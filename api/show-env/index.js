const { ClientSecretCredential } = require("@azure/identity");
const fetch = require("node-fetch");

module.exports = async function (context, req) {
    try {
        // --- 1. 環境変数の取得 ---
        const tenantId = process.env.TENANT_ID;
        const clientId = process.env.CLIENT_ID;
        const clientSecret = process.env.CLIENT_SECRET;
        const dataverseUrl = process.env.DATAVERSE_URL;

        if (!tenantId || !clientId || !clientSecret || !dataverseUrl) {
            throw new Error("サーバー設定エラー: 環境変数が不足しています。");
        }

        // --- 2. ユーザー特定 ---
        const header = req.headers["x-ms-client-principal"];
        let userEmail = "unknown";
        if (header) {
            const decoded = JSON.parse(Buffer.from(header, "base64").toString("ascii"));
            userEmail = decoded.userDetails;
        }
        // ゲスト対応
        if (userEmail.includes("#EXT#")) {
            let temp = userEmail.split("#EXT#")[0];
            const lastUnderscore = temp.lastIndexOf("_");
            if (lastUnderscore !== -1) {
                userEmail = temp.substring(0, lastUnderscore) + "@" + temp.substring(lastUnderscore + 1);
            }
        }

        // --- 3. Dataverse接続 (トークン取得) ---
        const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        const tokenResp = await credential.getToken(`${dataverseUrl}/.default`);
        const token = tokenResp.token;

        // --- 4. マスタ検索 (ユーザー情報) ---
        const userFilter = `new_mail eq '${userEmail}'`;
        const userQuery = `${dataverseUrl}/api/data/v9.2/new_sagyouin_mastas?$filter=${encodeURIComponent(userFilter)}&$select=new_sagyouin_id,_owningbusinessunit_value`;
        
        const userRes = await fetch(userQuery, { 
            headers: { 
                "Authorization": `Bearer ${token}`,
                "Prefer": "odata.include-annotations=\"*\""
            } 
        });

        if (!userRes.ok) {
            const errTxt = await userRes.text();
            throw new Error(`マスタ検索エラー: ${errTxt}`);
        }

        const userData = await userRes.json();
        if (!userData.value || userData.value.length === 0) {
            context.res = { status: 403, body: { error: `未登録のユーザーです (${userEmail})` } };
            return;
        }

        const user = userData.value[0];
        const buId = user._owningbusinessunit_value;
        const buName = user["_owningbusinessunit_value@OData.Community.Display.V1.FormattedValue"];

        // --- 5. 配車データ取得 ---
        // ★修正: 確実に存在する列のみを指定
        const selectCols = [
            "new_day", 
            "new_start_time", 
            "new_genbamei", 
            "new_sagyou_naiyou", 
            "new_shinkoujoukyou", 
            "new_table2id",
            "new_tokuisaki_meinvarchar", 
            "_new_kyakusaki_value", 
            "_new_sharyou_value", 
            "new_kashikiripicklist", 
            "_new_renraku1_value", 
            "new_renraku_jikountext"
        ].join(",");

        const myDispatchFilter = `_new_operator_value eq '${user.new_sagyouin_mastaid}' and statecode eq 0`; 
        const myDispatchQuery = `${dataverseUrl}/api/data/v9.2/new_table2s?$filter=${encodeURIComponent(myDispatchFilter)}&$select=${selectCols}&$orderby=new_day asc`;

        const dispatchRes = await fetch(myDispatchQuery, { 
            headers: { 
                "Authorization": `Bearer ${token}`,
                "Prefer": "odata.include-annotations=\"*\"" 
            } 
        });

        // ★エラー詳細をキャッチ
        if (!dispatchRes.ok) {
            const errTxt = await dispatchRes.text();
            throw new Error(`配車データ取得エラー: ${errTxt}`);
        }

        const dispatchData = await dispatchRes.json();
        let records = dispatchData.value;

        // --- 6. 資料データの取得 (Docs) ---
        if (records.length > 0) {
            const haishaIds = records.map(r => r.new_table2id);
            // 配車IDで紐づく資料を検索
            const docFilter = haishaIds.map(id => `_new_haisha_value eq '${id}'`).join(" or ");
            
            const docQuery = `${dataverseUrl}/api/data/v9.2/new_docment_tables?$filter=${encodeURIComponent(docFilter)}&$select=new_namenvarchar,new_kakuchoushin,new_url,new_blobthmbnailurl,_new_haisha_value`;

            const docRes = await fetch(docQuery, { headers: { "Authorization": `Bearer ${token}` } });
            
            // 資料取得は失敗してもメイン処理は止めない
            if (docRes.ok) {
                const docData = await docRes.json();
                const docs = docData.value;
                records = records.map(rec => {
                    rec.documents = docs.filter(d => d._new_haisha_value === rec.new_table2id);
                    return rec;
                });
            }
        }

        // --- 7. 部署稼働数カウント ---
        const today = new Date().toISOString().split('T')[0];
        const countFilter = `_owningbusinessunit_value eq '${buId}' and new_day eq ${today}`;
        const countQuery = `${dataverseUrl}/api/data/v9.2/new_table2s?$filter=${encodeURIComponent(countFilter)}&$count=true&$top=0`;
        
        const countRes = await fetch(countQuery, { 
            headers: { "Authorization": `Bearer ${token}` } 
        });
        const buCount = countRes.ok ? (await countRes.json())["@odata.count"] : 0;

        // 結果返却
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
        // エラー詳細を画面に返す
        context.res = { 
            status: 500, 
            body: { error: `System Error: ${e.message}` } 
        };
    }
};
