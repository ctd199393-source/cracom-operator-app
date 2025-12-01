const { ClientSecretCredential } = require("@azure/identity");
const fetch = require("node-fetch");

module.exports = async function (context, req) {
    try {
        const tenantId = process.env.TENANT_ID;
        const clientId = process.env.CLIENT_ID;
        const clientSecret = process.env.CLIENT_SECRET;
        const dataverseUrl = process.env.DATAVERSE_URL;

        if (!tenantId || !clientId || !clientSecret || !dataverseUrl) {
            throw new Error("環境設定が不足しています。");
        }

        // 1. ユーザー特定
        const header = req.headers["x-ms-client-principal"];
        let userEmail = "unknown";
        if (header) {
            const decoded = JSON.parse(Buffer.from(header, "base64").toString("ascii"));
            userEmail = decoded.userDetails;
        }
        // ゲストユーザー対応
        if (userEmail.includes("#EXT#")) {
            let temp = userEmail.split("#EXT#")[0];
            const lastUnderscore = temp.lastIndexOf("_");
            if (lastUnderscore !== -1) {
                userEmail = temp.substring(0, lastUnderscore) + "@" + temp.substring(lastUnderscore + 1);
            }
        }

        // 2. Dataverse接続
        const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        const tokenResp = await credential.getToken(`${dataverseUrl}/.default`);
        const token = tokenResp.token;

        // 3. 作業員マスタ検索
        const userFilter = `new_mail eq '${userEmail}'`;
        const userQuery = `${dataverseUrl}/api/data/v9.2/new_sagyouin_mastas?$filter=${encodeURIComponent(userFilter)}&$select=new_sagyouin_id,_owningbusinessunit_value`;
        
        const userRes = await fetch(userQuery, { headers: { "Authorization": `Bearer ${token}` } });
        const userData = await userRes.json();

        if (!userData.value || userData.value.length === 0) {
            context.res = { status: 403, body: { error: "マスタ未登録" } };
            return;
        }

        const user = userData.value[0];
        const buId = user._owningbusinessunit_value;

        // 4. 配車データ取得（自分の予定）
        // 過去1週間～未来のデータを取得（フィルタリングはJS側で行うため広めに取る）
        // 必要な列: 日付, 開始時間, 現場名, 作業内容, 進行状況, 案件ID
        const myDispatchFilter = `_new_operator_value eq '${user.new_sagyouin_mastaid}' and statecode eq 0`; // statecode 0 = Active
        const myDispatchQuery = `${dataverseUrl}/api/data/v9.2/new_table2s?$filter=${encodeURIComponent(myDispatchFilter)}&$select=new_day,new_start_time,new_genbamei,new_sagyou_naiyou,new_shinkoujoukyou,new_table2id&$orderby=new_day asc`;

        const dispatchRes = await fetch(myDispatchQuery, { headers: { "Authorization": `Bearer ${token}` } });
        const dispatchData = await dispatchRes.json();

        // 5. 今日の部署稼働数のカウント (集計)
        // 条件: 同じBusinessUnit かつ 日付が今日
        const today = new Date().toISOString().split('T')[0];
        const countFilter = `_owningbusinessunit_value eq '${buId}' and new_day eq ${today}`;
        const countQuery = `${dataverseUrl}/api/data/v9.2/new_table2s?$filter=${encodeURIComponent(countFilter)}&$count=true&$top=0`; // データは取らず件数だけ
        
        const countRes = await fetch(countQuery, { 
            headers: { 
                "Authorization": `Bearer ${token}`,
                "Prefer": "odata.include-annotations=\"*\"" // カウント取得に必要
            } 
        });
        const countJson = await countRes.json();
        const buCount = countJson["@odata.count"] || 0;

        context.res = {
            status: 200,
            body: {
                user: { name: user.new_sagyouin_id, buId: buId },
                records: dispatchData.value,
                todayCount: buCount
            }
        };

    } catch (e) {
        context.log.error(e);
        context.res = { status: 500, body: { error: e.message } };
    }
};
