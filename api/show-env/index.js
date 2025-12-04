const { ClientSecretCredential } = require("@azure/identity");
const fetch = require("node-fetch");

module.exports = async function (context, req) {
    try {
        // --- 1. 環境設定 ---
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
        // ★修正: new_isdeleted を削除しました
        const selectCols = [
            "new_day", 
            "new_start_time", 
            "new_genbamei", 
            "new_sagyou_naiyou", 
            "new_shinkoujoukyou", 
            "new_table2id",
            "new_tokuisaki_mei",
            "new_kyakusaki",
            "new_sharyou",
            "new_kashikiri",
            "new_renraku1",
            "new_renraku_jikou",
            "new_type",
            "new_haisha_zumi"
        ].join(",");

        // 日付計算: 「昨日」以降
        const dt = new Date();
        dt.setDate(dt.getDate() - 1);
        const yesterdayStr = dt.toISOString().split('T')[0];

        // ★修正: フィルタ条件から new_isdeleted eq false を削除
        // statecode eq 0 (アクティブ) だけで有効データを判定します
        const myDispatchFilter = `
            _new_operator_value eq '${user.new_sagyouin_mastaid}' and 
            statecode eq 0 and 
            new_type eq 100000000 and 
            new_day ge ${yesterdayStr} and
            new_haisha_zumi eq true
        `.replace(/\s+/g, ' ').trim();

        const myDispatchQuery = `${dataverseUrl}/api/data/v9.2/new_table2s?$filter=${encodeURIComponent(myDispatchFilter)}&$select=${selectCols}&$orderby=new_day asc`;

        const dispatchRes = await fetch(myDispatchQuery, { 
            headers: { 
                "Authorization": `Bearer ${token}`,
                "Prefer": "odata.include-annotations=\"*\""
            } 
        });

        if (!dispatchRes.ok) {
            const errTxt = await dispatchRes.text();
            throw new Error(`配車データ取得エラー: ${errTxt}`);
        }

        const dispatchData = await dispatchRes.json();
        let records = dispatchData.value;

        // --- 6. 資料データの取得 ---
        if (records.length > 0) {
            const haishaIds = records.map(r => r.new_table2id);
            const docFilter = haishaIds.map(id => `_new_haisha_value eq '${id}'`).join(" or ");
            
            const docQuery = `${dataverseUrl}/api/data/v9.2/new_docment_tables?$filter=${encodeURIComponent(docFilter)}&$select=new_name,new_kakuchoushi,new_url,new_blobthmbnailurl,_new_haisha_value`;

            const docRes = await fetch(docQuery, { headers: { "Authorization": `Bearer ${token}` } });
            
            if (docRes.ok) {
                const docData = await docRes.json();
                const docs = docData.value;
                records = records.map(rec => {
                    rec.documents = docs.filter(d => d._new_haisha_value === rec.new_table2id);
                    return rec;
                });
            }
        }

        // --- 7. 部署稼働数 ---
        // ★修正: ここからも new_isdeleted を削除
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
        context.res = { status: 500, body: { error: `System Error: ${e.message}` } };
    }
};
