const { ClientSecretCredential } = require("@azure/identity");
const { StorageSharedKeyCredential, generateBlobSASQueryParameters, BlobSASPermissions } = require("@azure/storage-blob");
const fetch = require("node-fetch");

// SASトークン生成ヘルパー
function generateSasToken(connectionString, containerName, blobPath) {
    try {
        if (!connectionString || !containerName || !blobPath) return null;
        
        const parts = connectionString.split(';');
        const accountName = parts.find(p => p.startsWith('AccountName=')).split('=')[1];
        const accountKey = parts.find(p => p.startsWith('AccountKey=')).split('=')[1];

        // Blobパスの正規化
        let blobName = blobPath;
        // "/containerName/path" の形式から containerName を除去して純粋なBlob名にする
        if (blobName.startsWith(`/${containerName}/`)) {
            blobName = blobName.substring(containerName.length + 2);
        } else if (blobName.startsWith("/")) {
             blobName = blobName.substring(1);
        }

        const sharedKeyCredential = new StorageSharedKeyCredential(accountName, accountKey);
        const expiresOn = new Date(new Date().valueOf() + 60 * 60 * 1000); // 1時間有効

        const sasToken = generateBlobSASQueryParameters({
            containerName: containerName,
            blobName: blobName,
            permissions: BlobSASPermissions.parse("r"),
            expiresOn: expiresOn
        }, sharedKeyCredential).toString();

        return `https://${accountName}.blob.core.windows.net/${containerName}/${blobName}?${sasToken}`;
    } catch (e) {
        console.error("SAS Gen Error:", e);
        return null;
    }
}

module.exports = async function (context, req) {
    try {
        const tenantId = process.env.TENANT_ID;
        const clientId = process.env.CLIENT_ID;
        const clientSecret = process.env.CLIENT_SECRET;
        const dataverseUrl = process.env.DATAVERSE_URL;
        const storageConnString = process.env.AZURE_STORAGE_CONNECTION_STRING;

        if (!tenantId || !clientId || !clientSecret || !dataverseUrl) {
            throw new Error("環境設定不足: Dataverse接続情報がありません");
        }

        // 1. ユーザー特定
        const header = req.headers["x-ms-client-principal"];
        let userEmail = "unknown";
        if (header) {
            const decoded = JSON.parse(Buffer.from(header, "base64").toString("ascii"));
            userEmail = decoded.userDetails;
        }
        if (userEmail.includes("#EXT#")) {
            let temp = userEmail.split("#EXT#")[0];
            const last = temp.lastIndexOf("_");
            if (last !== -1) userEmail = temp.substring(0, last) + "@" + temp.substring(last + 1);
        }

        // 2. Dataverse接続
        const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        const tokenResp = await credential.getToken(`${dataverseUrl}/.default`);
        const token = tokenResp.token;

        // 3. マスタ検索
        const userFilter = `new_mail eq '${userEmail}'`;
        const userQuery = `${dataverseUrl}/api/data/v9.2/new_sagyouin_mastas?$filter=${encodeURIComponent(userFilter)}&$select=new_sagyouin_id,_owningbusinessunit_value`;
        
        const userRes = await fetch(userQuery, { headers: { "Authorization": `Bearer ${token}` } });
        if (!userRes.ok) throw new Error(await userRes.text());
        const userData = await userRes.json();
        
        if (!userData.value || userData.value.length === 0) {
            context.res = { status: 403, body: { error: "マスタ未登録" } };
            return;
        }
        const user = userData.value[0];
        const buId = user._owningbusinessunit_value;
        const buName = user["_owningbusinessunit_value@OData.Community.Display.V1.FormattedValue"];

        // 4. 配車データ取得
        // ★修正: new_end_time (終了時間), new_yousha_irai (傭車依頼) を追加
        const selectCols = [
            "new_day", "new_start_time", "new_end_time", // 時間関係
            "new_genbamei", "new_sagyou_naiyou", "new_shinkoujoukyou", 
            "new_table2id", "new_tokuisaki_mei", "new_kyakusaki", "new_sharyou", "new_kashikiri", 
            "new_renraku1", "new_renraku_jikou", "new_type", "new_haisha_zumi",
            "new_yousha_irai", // 傭車依頼フラグ
            "_new_id_value", "_new_sagyouba_value"
        ].join(",");

        const dt = new Date();
        dt.setDate(dt.getDate() - 1);
        const yesterdayStr = dt.toISOString().split('T')[0];

        const myDispatchFilter = `_new_operator_value eq '${user.new_sagyouin_mastaid}' and statecode eq 0 and new_type eq 100000000 and new_day ge ${yesterdayStr} and new_haisha_zumi eq true`;
        const myDispatchQuery = `${dataverseUrl}/api/data/v9.2/new_table2s?$filter=${encodeURIComponent(myDispatchFilter)}&$select=${selectCols}&$orderby=new_day asc`;

        const dispatchRes = await fetch(myDispatchQuery, { 
            headers: { "Authorization": `Bearer ${token}`, "Prefer": "odata.include-annotations=\"*\"" } 
        });

        if (!dispatchRes.ok) throw new Error(`配車取得エラー: ${await dispatchRes.text()}`);
        const dispatchData = await dispatchRes.json();
        let records = dispatchData.value;

        // 5. 資料データ取得 & SAS発行
        if (records.length > 0 && storageConnString) {
            let filterParts = [];
            records.forEach(r => {
                filterParts.push(`_new_haisha_value eq '${r.new_table2id}'`);
                if (r._new_id_value) filterParts.push(`_new_anken_value eq '${r._new_id_value}'`);
                if (r._new_sagyouba_value) filterParts.push(`_new_genba_value eq '${r._new_sagyouba_value}'`);
            });
            
            const uniqueFilters = [...new Set(filterParts)];
            
            if (uniqueFilters.length > 0) {
                const docFilter = uniqueFilters.join(" or ");
                const docQuery = `${dataverseUrl}/api/data/v9.2/new_docment_tables?$filter=${encodeURIComponent(docFilter)}&$select=new_name,new_kakuchoushi,new_blob_pass,new_container,new_blobthmbnailurl,_new_haisha_value,_new_anken_value,_new_genba_value`;

                const docRes = await fetch(docQuery, { headers: { "Authorization": `Bearer ${token}` } });
                if (docRes.ok) {
                    const docJson = await docRes.json();
                    const allDocs = docJson.value;

                    records = records.map(rec => {
                        const myDocs = allDocs.filter(d => 
                            d._new_haisha_value === rec.new_table2id ||
                            (rec._new_id_value && d._new_anken_value === rec._new_id_value) ||
                            (rec._new_sagyouba_value && d._new_genba_value === rec._new_sagyouba_value)
                        );

                        rec.documents = myDocs.map(d => {
                            if (!d.new_blob_pass || !d.new_container) return null;
                            
                            // 本体のSAS URL
                            const sasUrl = generateSasToken(storageConnString, d.new_container, d.new_blob_pass);
                            
                            // ★修正: サムネイル用SAS URL生成ロジック
                            let thumbSasUrl = null;
                            const ext = (d.new_kakuchoushi || "").toLowerCase();
                            const isImage = ['.jpg', '.jpeg', '.png', '.gif', '.bmp'].some(e => ext.includes(e));

                            if (d.new_blobthmbnailurl) {
                                // サムネイル列にパスがあればそれを使う
                                thumbSasUrl = generateSasToken(storageConnString, d.new_container, d.new_blobthmbnailurl);
                            } else if (isImage) {
                                // 画像ファイルなら、本体URLをサムネイルとして代用
                                thumbSasUrl = sasUrl;
                            }

                            return {
                                new_name: d.new_name,
                                new_kakuchoushi: d.new_kakuchoushi,
                                new_url: sasUrl, 
                                new_blobthmbnailurl: thumbSasUrl // SAS付きのサムネイルURL (なければnull)
                            };
                        }).filter(d => d !== null && d.new_url !== null);
                        return rec;
                    });
                }
            }
        }

        // 6. 部署稼働数
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
