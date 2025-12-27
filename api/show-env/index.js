const { ClientSecretCredential } = require("@azure/identity");
const { StorageSharedKeyCredential, generateBlobSASQueryParameters, BlobSASPermissions } = require("@azure/storage-blob");
const fetch = require("node-fetch");

// Helper: SASトークン生成
function generateSasToken(connectionString, containerName, blobPath) {
    try {
        if (!connectionString || !containerName || !blobPath) return null;
        
        const parts = connectionString.split(';');
        const accountNameMatch = parts.find(p => p.startsWith('AccountName='));
        const accountKeyMatch = parts.find(p => p.startsWith('AccountKey='));

        if (!accountNameMatch || !accountKeyMatch) return null;

        const accountName = accountNameMatch.split('=')[1];
        const accountKey = accountKeyMatch.split('=')[1];

        // Blobパス正規化
        let blobName = blobPath;
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
        // 1. 環境設定
        const tenantId = process.env.TENANT_ID;
        const clientId = process.env.CLIENT_ID;
        const clientSecret = process.env.CLIENT_SECRET;
        const dataverseUrl = process.env.DATAVERSE_URL; 
        const storageConnString = process.env.AZURE_STORAGE_CONNECTION_STRING;

        if (!tenantId || !clientId || !clientSecret || !dataverseUrl) {
            throw new Error("環境変数が不足しています (Dataverse)");
        }

        // 2. ユーザー特定 (SWA Authentication)
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

        // 3. Dataverse接続
        const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        const tokenResp = await credential.getToken(`${dataverseUrl}/.default`);
        const token = tokenResp.token;

        // 4. 作業員マスタ検索
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

        // 5. 配車データ取得
        const selectCols = [
            "new_day", "new_start_time", "new_genbamei", "new_sagyou_naiyou", "new_shinkoujoukyou", 
            "new_table2id", "new_tokuisaki_mei", "new_kyakusaki", "new_sharyou", "new_kashikiri", 
            "new_renraku1", "new_renraku_jikou", "new_type", "new_haisha_zumi",
            "new_end_time", "new_yousha_irai", "modifiedon",
            "_new_id_value", "_new_sagyouba_value" 
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

        const myDispatchQuery = `${dataverseUrl}/api/data/v9.2/new_table2s?$filter=${encodeURIComponent(myDispatchFilter)}&$select=${selectCols}&$expand=new_sharyou($select=new_shaban,new_tsuriage)&$orderby=new_day asc`;

        const dispatchRes = await fetch(myDispatchQuery, { 
            headers: { "Authorization": `Bearer ${token}`, "Prefer": "odata.include-annotations=\"*\"" } 
        });
        if (!dispatchRes.ok) throw new Error(`配車取得エラー: ${await dispatchRes.text()}`);
        const dispatchData = await dispatchRes.json();
        let records = dispatchData.value;

        if (records.length > 0) {
            // =========================================================
            // A. 中間テーブル（作業場）の取得 & 結合
            // =========================================================
            records.forEach(r => r.sagyouba_list = []);
            const haishaIds = records.map(r => r.new_table2id);
            
            // ★確定情報: 論理名 new_haisha_id なので _new_haisha_id_value でフィルタ
            const chuukanFilter = haishaIds.map(id => `_new_haisha_id_value eq '${id}'`).join(" or ");

            if (chuukanFilter) {
                // ★確定情報: テーブル名 new_haisha_sagyouba_chuukans (複数形)
                // ★確定情報: 展開列 new_sagyouba (論理名)
                const chuukanQuery = `${dataverseUrl}/api/data/v9.2/new_haisha_sagyouba_chuukans?$filter=${encodeURIComponent(chuukanFilter)}&$expand=new_sagyouba($select=new_name)`;
                
                try {
                    const chuukanRes = await fetch(chuukanQuery, { headers: { "Authorization": `Bearer ${token}` } });
                    if (chuukanRes.ok) {
                        const chuukanData = await chuukanRes.json();
                        chuukanData.value.forEach(c => {
                            // 親IDとの紐付け
                            const parentId = c._new_haisha_id_value; 
                            const targetRec = records.find(r => r.new_table2id === parentId);
                            
                            // 名称があればリストに追加
                            if (targetRec && c.new_sagyouba && c.new_sagyouba.new_name) {
                                targetRec.sagyouba_list.push({ name: c.new_sagyouba.new_name });
                            }
                        });
                    }
                } catch (e) { console.error("中間テーブル取得エラー:", e); }
            }

            // B. 案件・現場マスタ（GoogleMapリンク）の別途取得
            const ankenIds = [...new Set(records.map(r => r._new_id_value).filter(id => id))];
            if (ankenIds.length > 0) {
                const ankenFilter = ankenIds.map(id => `new_ankenid eq '${id}'`).join(" or ");
                const ankenQuery = `${dataverseUrl}/api/data/v9.2/new_ankens?$filter=${encodeURIComponent(ankenFilter)}&$select=new_ankenid&$expand=new_genba($select=new_googlemap_link)`;
                
                try {
                    const ankenRes = await fetch(ankenQuery, { headers: { "Authorization": `Bearer ${token}` } });
                    if (ankenRes.ok) {
                        const ankenData = await ankenRes.json();
                        ankenData.value.forEach(a => {
                            if (a.new_genba && a.new_genba.new_googlemap_link) {
                                records.filter(r => r._new_id_value === a.new_ankenid).forEach(r => {
                                    r.googlemap_link = a.new_genba.new_googlemap_link;
                                });
                            }
                        });
                    }
                } catch (e) { console.error("案件・現場取得エラー:", e); }
            }
        }

        // 6. 資料データ取得 & SAS発行
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
                    const docData = await docRes.json();
                    const allDocs = docData.value;

                    records = records.map(rec => {
                        const myDocs = allDocs.filter(d => 
                            d._new_haisha_value === rec.new_table2id ||
                            (rec._new_id_value && d._new_anken_value === rec._new_id_value) ||
                            (rec._new_sagyouba_value && d._new_genba_value === rec._new_sagyouba_value)
                        );

                        rec.documents = myDocs.map(d => {
                            if (!d.new_blob_pass || !d.new_container) return null;

                            const sasUrl = generateSasToken(storageConnString, d.new_container, d.new_blob_pass);
                            let thumbSasUrl = null;
                            const ext = (d.new_kakuchoushi || "").toLowerCase();
                            const isImage = ['.jpg', '.jpeg', '.png', '.gif', '.bmp'].some(e => ext.includes(e));

                            if (d.new_blobthmbnailurl) {
                                thumbSasUrl = generateSasToken(storageConnString, d.new_container, d.new_blobthmbnailurl);
                            } else if (isImage) {
                                thumbSasUrl = sasUrl;
                            }

                            return {
                                new_name: d.new_name,
                                new_kakuchoushi: d.new_kakuchoushi,
                                new_url: sasUrl,
                                new_blobthmbnailurl: thumbSasUrl 
                            };
                        }).filter(d => d !== null && d.new_url !== null);

                        return rec;
                    });
                }
            }
        }

        // 7. 部署稼働数
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
