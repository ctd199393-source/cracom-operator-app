const { ClientSecretCredential } = require("@azure/identity");
const { StorageSharedKeyCredential, generateBlobSASQueryParameters, BlobSASPermissions } = require("@azure/storage-blob");
const fetch = require("node-fetch");

// --- Helper: SASトークン生成関数 ---
function generateSasToken(connectionString, containerName, blobPath) {
    try {
        if (!connectionString || !containerName || !blobPath) return null;

        // 接続文字列の解析
        const parts = connectionString.split(';');
        const accountName = parts.find(p => p.startsWith('AccountName=')).split('=')[1];
        const accountKey = parts.find(p => p.startsWith('AccountKey=')).split('=')[1];
        
        // Blobパスの正規化 (先頭の / やコンテナ名を除去)
        // 例: "/mycontainer/folder/file.pdf" -> "folder/file.pdf"
        let blobName = blobPath;
        if (blobName.startsWith(`/${containerName}/`)) {
            blobName = blobName.substring(containerName.length + 2);
        } else if (blobName.startsWith("/")) {
             blobName = blobName.substring(1);
        }

        const sharedKeyCredential = new StorageSharedKeyCredential(accountName, accountKey);
        
        // 有効期限: 60分
        const expiresOn = new Date(new Date().valueOf() + 60 * 60 * 1000);

        // SAS生成
        const sasToken = generateBlobSASQueryParameters({
            containerName: containerName,
            blobName: blobName,
            permissions: BlobSASPermissions.parse("r"), // Read権限のみ
            expiresOn: expiresOn
        }, sharedKeyCredential).toString();

        // 署名付きURLを返す
        return `https://${accountName}.blob.core.windows.net/${containerName}/${blobName}?${sasToken}`;
    } catch (e) {
        console.error("SAS Gen Error:", e);
        return null; 
    }
}

module.exports = async function (context, req) {
    try {
        // --- 1. 環境変数チェック ---
        const tenantId = process.env.TENANT_ID;
        const clientId = process.env.CLIENT_ID;
        const clientSecret = process.env.CLIENT_SECRET;
        const dataverseUrl = process.env.DATAVERSE_URL;
        const storageConnString = process.env.AZURE_STORAGE_CONNECTION_STRING; 

        if (!tenantId || !clientId || !clientSecret || !dataverseUrl) {
            throw new Error("環境変数が不足しています (Dataverse)");
        }
        
        // --- 2. ユーザー特定 ---
        const header = req.headers["x-ms-client-principal"];
        let userEmail = "unknown";
        if (header) {
            const decoded = JSON.parse(Buffer.from(header, "base64").toString("ascii"));
            userEmail = decoded.userDetails;
        }
        if (userEmail.includes("#EXT#")) {
            let temp = userEmail.split("#EXT#")[0];
            const lastUnderscore = temp.lastIndexOf("_");
            if (lastUnderscore !== -1) {
                userEmail = temp.substring(0, lastUnderscore) + "@" + temp.substring(lastUnderscore + 1);
            }
        }

        // --- 3. Dataverse接続 ---
        const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        const tokenResp = await credential.getToken(`${dataverseUrl}/.default`);
        const token = tokenResp.token;

        // --- 4. 作業員マスタ検索 ---
        const userFilter = `new_mail eq '${userEmail}'`;
        const userQuery = `${dataverseUrl}/api/data/v9.2/new_sagyouin_mastas?$filter=${encodeURIComponent(userFilter)}&$select=new_sagyouin_id,_owningbusinessunit_value`;
        
        const userRes = await fetch(userQuery, { 
            headers: { "Authorization": `Bearer ${token}`, "Prefer": "odata.include-annotations=\"*\"" } 
        });
        if (!userRes.ok) throw new Error(await userRes.text());
        const userData = await userRes.json();
        
        if (!userData.value || userData.value.length === 0) {
            context.res = { status: 403, body: { error: "マスタ未登録ユーザー" } };
            return;
        }

        const user = userData.value[0];
        const buId = user._owningbusinessunit_value;
        const buName = user["_owningbusinessunit_value@OData.Community.Display.V1.FormattedValue"];

        // --- 5. 配車データ取得 ---
        // ★修正: 案件(new_id) と 現場(new_sagyouba) も取得リストに追加
        const selectCols = [
            "new_day", "new_start_time", "new_genbamei", "new_sagyou_naiyou", "new_shinkoujoukyou", 
            "new_table2id", "new_tokuisaki_mei", "new_kyakusaki", "new_sharyou", "new_kashikiri", 
            "new_renraku1", "new_renraku_jikou", "new_type", "new_haisha_zumi",
            "new_id",       // 案件
            "new_sagyouba"  // 現場
        ].join(",");

        const dt = new Date();
        dt.setDate(dt.getDate() - 1);
        const yesterdayStr = dt.toISOString().split('T')[0];

        const myDispatchFilter = `
            _new_operator_value eq '${user.new_sagyouin_mastaid}' and 
            statecode eq 0 and 
            new_type eq 100000000 and 
            new_day ge ${yesterdayStr} and
            new_haisha_zumi eq true
        `.replace(/\s+/g, ' ').trim();

        const myDispatchQuery = `${dataverseUrl}/api/data/v9.2/new_table2s?$filter=${encodeURIComponent(myDispatchFilter)}&$select=${selectCols}&$orderby=new_day asc`;

        const dispatchRes = await fetch(myDispatchQuery, { 
            headers: { "Authorization": `Bearer ${token}`, "Prefer": "odata.include-annotations=\"*\"" } 
        });
        if (!dispatchRes.ok) throw new Error(`配車取得エラー: ${await dispatchRes.text()}`);
        const dispatchData = await dispatchRes.json();
        let records = dispatchData.value;

        // --- 6. 資料データの取得 & SAS発行 ---
        if (records.length > 0 && storageConnString) {
            
            // 検索条件の構築 (配車ID OR 案件ID OR 現場ID)
            let filterParts = [];
            records.forEach(r => {
                filterParts.push(`_new_haisha_value eq '${r.new_table2id}'`);
                // 案件IDがあれば追加
                if (r._new_id_value) filterParts.push(`_new_anken_value eq '${r._new_id_value}'`);
                // 現場IDがあれば追加
                if (r._new_sagyouba_value) filterParts.push(`_new_genba_value eq '${r._new_sagyouba_value}'`);
            });

            // 重複除去してクエリ結合
            const uniqueFilters = [...new Set(filterParts)];

            if (uniqueFilters.length > 0) {
                // 件数が多いとURL長制限に引っかかる可能性がありますが、当面はこれでいきます
                const docFilter = uniqueFilters.join(" or ");
                
                // 資料テーブル (new_docment_tables) を検索
                const docQuery = `${dataverseUrl}/api/data/v9.2/new_docment_tables?$filter=${encodeURIComponent(docFilter)}&$select=new_name,new_kakuchoushi,new_blob_pass,new_container,new_blobthmbnailurl,_new_haisha_value,_new_anken_value,_new_genba_value`;

                const docRes = await fetch(docQuery, { headers: { "Authorization": `Bearer ${token}` } });
                
                if (docRes.ok) {
                    const docData = await docRes.json();
                    const allDocs = docData.value;

                    // 各配車レコードに、関連する資料をマッピング
                    records = records.map(rec => {
                        // 自分の 配車ID / 案件ID / 現場ID にマッチする資料を抽出
                        const myDocs = allDocs.filter(d => 
                            d._new_haisha_value === rec.new_table2id ||
                            (rec._new_id_value && d._new_anken_value === rec._new_id_value) ||
                            (rec._new_sagyouba_value && d._new_genba_value === rec._new_sagyouba_value)
                        );

                        // SAS URL生成してオブジェクト化
                        rec.documents = myDocs.map(d => {
                            // Blobパスがない場合はスキップ
                            if (!d.new_blob_pass || !d.new_container) return null;

                            // SAS付きURL発行
                            const sasUrl = generateSasToken(storageConnString, d.new_container, d.new_blob_pass);
                            
                            return {
                                new_name: d.new_name,
                                new_kakuchoushi: d.new_kakuchoushi,
                                new_url: sasUrl, // ★ここが有効期限付きURLになります
                                new_blobthmbnailurl: d.new_blobthmbnailurl // サムネイルはそのまま(SASなし)
                            };
                        }).filter(d => d !== null && d.new_url !== null);

                        return rec;
                    });
                }
            }
        }

        // --- 7. 部署稼働数 ---
        const today = new Date().toISOString().split('T')[0];
        const countFilter = `_owningbusinessunit_value eq '${buId}' and new_day eq ${today} and new_type eq 100000000 and new_haisha_zumi eq true and statecode eq 0`;
        const countQuery = `${dataverseUrl}/api/data/v9.2/new_table2s?$filter=${encodeURIComponent(countFilter)}&$count=true&$top=0`;
        const countRes = await fetch(countQuery, { headers: { "Authorization": `Bearer ${token}` } });
        const buCount = countRes.ok ? (await countRes.json())["@odata.count"] : 0;

        context.res = {
            status: 200,
            body: {
                user: { name: user.new_sagyouin_id, buId: buId, buName: buName },
                records: records,
                todayCount: buCount
            }
        };

    } catch (e) {
        context.log.error(e);
        context.res = { status: 500, body: { error: e.message } };
    }
};
