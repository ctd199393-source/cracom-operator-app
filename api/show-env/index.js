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
        // 1. 環境設定の読み込みチェック
        const tenantId = process.env.TENANT_ID;
        const clientId = process.env.CLIENT_ID;
        const clientSecret = process.env.CLIENT_SECRET;
        const dataverseUrl = process.env.DATAVERSE_URL; 
        const storageConnString = process.env.AZURE_STORAGE_CONNECTION_STRING;

        if (!tenantId || !clientId || !clientSecret || !dataverseUrl) {
            throw new Error(`環境変数が不足しています。チェック状況: TENANT_ID:${!!tenantId}, CLIENT_ID:${!!clientId}, CLIENT_SECRET:${!!clientSecret}, DATAVERSE_URL:${!!dataverseUrl}`);
        }

        // 2. ユーザー特定 (SWA 認証ヘッダーの解析)
        const header = req.headers["x-ms-client-principal"];
        let userEmail = "unknown";
        if (header) {
            try {
                const decoded = JSON.parse(Buffer.from(header, "base64").toString("ascii"));
                userEmail = decoded.userDetails || "unknown";
            } catch (jsonErr) {
                throw new Error(`認証ヘッダーの解析に失敗しました: ${jsonErr.message}`);
            }
        } else {
            context.res = { status: 401, body: { error: "認証ヘッダーが見つかりません。ログインし直してください。" } };
            return;
        }

        // 外部ユーザー用メールアドレス正規化
        if (userEmail.includes("#EXT#")) {
            let temp = userEmail.split("#EXT#")[0];
            const last = temp.lastIndexOf("_");
            if (last !== -1) userEmail = temp.substring(0, last) + "@" + temp.substring(last + 1);
        }

        // 3. Dataverse接続 (認証トークンの取得)
        let token;
        try {
            const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
            const tokenResp = await credential.getToken(`${dataverseUrl}/.default`);
            token = tokenResp.token;
        } catch (authErr) {
            throw new Error(`Dataverseへの認証に失敗しました。シークレットが正しいか確認してください。詳細: ${authErr.message}`);
        }

        // 4. 作業員マスタ検索
        const userFilter = `new_mail eq '${userEmail}'`;
        const userQuery = `${dataverseUrl}/api/data/v9.2/new_sagyouin_mastas?$filter=${encodeURIComponent(userFilter)}&$select=new_sagyouin_id,new_sagyouin_mastaid,_owningbusinessunit_value`;
        
        const userRes = await fetch(userQuery, { 
            headers: { "Authorization": `Bearer ${token}`, "Prefer": "odata.include-annotations=\"*\"" } 
        });
        
        if (!userRes.ok) {
            const errText = await userRes.text();
            throw new Error(`作業員マスタの取得に失敗しました(Dataverse APIエラー): ${errText}`);
        }
        
        const userData = await userRes.json();
        
        if (!userData.value || userData.value.length === 0) {
            context.res = { status: 403, body: { error: `マスタ未登録ユーザーです。登録メールアドレス: ${userEmail}` } };
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

        // 6. 中間テーブルや資料データの紐付け処理 (ロジックは維持)
        if (records.length > 0) {
            // A. 中間テーブル（作業場）の取得 & 結合
            records.forEach(r => r.sagyouba_list = []);
            const haishaIds = records.map(r => r.new_table2id);
            const chuukanFilter = haishaIds.map(id => `_new_haisha_id_value eq '${id}'`).join(" or ");

            if (chuukanFilter) {
                const chuukanQuery = `${dataverseUrl}/api/data/v9.2/new_haisha_sagyouba_chuukans?$filter=${encodeURIComponent(chuukanFilter)}&$expand=new_sagyouba($select=new_title,new_googlemap)`;
                try {
                    const chuukanRes = await fetch(chuukanQuery, { headers: { "Authorization": `Bearer ${token}` } });
                    if (chuukanRes.ok) {
                        const chuukanData = await chuukanRes.json();
                        chuukanData.value.forEach(c => {
                            const parentId = c._new_haisha_id_value; 
                            const targetRec = records.find(r => r.new_table2id === parentId);
                            if (targetRec && c.new_sagyouba && c.new_sagyouba.new_title) {
                                const link = c.new_sagyouba.new_googlemap || c.new_sagyouba.new_Googlemap;
                                targetRec.sagyouba_list.push({ 
                                    name: c.new_sagyouba.new_title,
                                    mapLink: link || null
                                });
                            }
                        });
                    }
                } catch (e) { context.log.error("中間テーブル取得エラー:", e); }
            }
        }

        // 7. 部署稼働数の取得
        const today = new Date().toISOString().split('T')[0];
        const countFilter = `_owningbusinessunit_value eq '${buId}' and new_day eq ${today} and new_type eq 100000000 and new_haisha_zumi eq true and statecode eq 0`;
        const countQuery = `${dataverseUrl}/api/data/v9.2/new_table2s?$filter=${encodeURIComponent(countFilter)}&$count=true&$top=0`;
        const countRes = await fetch(countQuery, { headers: { "Authorization": `Bearer ${token}` } });
        const buCount = countRes.ok ? (await countRes.json())["@odata.count"] : 0;

        // 8. 成功レスポンス
        context.res = {
            status: 200,
            body: {
                user: { name: user.new_sagyouin_id, buId: buId, buName: buName },
                records: records,
                todayCount: buCount
            }
        };

    } catch (e) {
        context.log.error("API Error Summary:", e.message);
        context.res = {
            status: 500,
            body: {
                error: "API実行中にエラーが発生しました",
                details: e.message,
                hint: "環境変数(CLIENT_SECRETなど)や、Dataverse側のアプリケーションユーザー権限、APIのアクセス許可を再確認してください。"
            }
        };
    }
};
