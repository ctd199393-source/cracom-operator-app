const { ClientSecretCredential } = require("@azure/identity");
const fetch = require("node-fetch");

module.exports = async function (context, req) {
    context.log('Dataverse API request started.');

    try {
        // 1. 環境変数
        const tenantId = process.env.TENANT_ID; 
        const clientId = process.env.CLIENT_ID;
        const clientSecret = process.env.CLIENT_SECRET;
        const dataverseUrl = process.env.DATAVERSE_URL;

        if (!tenantId || !clientId || !clientSecret || !dataverseUrl) {
            throw new Error("環境変数が不足しています");
        }

        // 2. ユーザー情報
        const userEmail = req.headers["x-ms-client-principal-name"];
        if (!userEmail) {
            context.res = { status: 401, body: { error: "ログインが必要です" } };
            return;
        }

        // 3. 認証
        const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        const tokenResponse = await credential.getToken(`${dataverseUrl}/.default`);
        const accessToken = tokenResponse.token;
        
        const headers = {
            "Authorization": `Bearer ${accessToken}`,
            "Accept": "application/json",
            "Content-Type": "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            "Prefer": "odata.include-annotations=\"*\""
        };

        // 4. 作業員マスタ検索
        const workerTable = "new_sagyouin_mastas"; 
        const workerQuery = `?$select=_owningbusinessunit_value,new_sagyouin_mastaid,new_mei&$filter=new_mail eq '${userEmail}'`;
        const workerRes = await fetch(`${dataverseUrl}/api/data/v9.2/${workerTable}${workerQuery}`, { method: "GET", headers });
        
        if (!workerRes.ok) {
            throw new Error(`Worker Search Error: ${workerRes.status}`);
        }
        const workerData = await workerRes.json();

        if (!workerData.value || workerData.value.length === 0) {
            context.res = { status: 403, body: { error: "作業員マスタに登録がありません" } };
            return;
        }

        const worker = workerData.value[0];
        const myBusinessUnit = worker._owningbusinessunit_value;
        const myWorkerId = worker.new_sagyouin_mastaid;
        const myName = worker.new_mei || "担当者";
        const myBusinessUnitName = worker["_owningbusinessunit_value@OData.Community.Display.V1.FormattedValue"] || "";

        // 5. 配車データ取得 (配車テーブルから直接取得)
        // ★修正点: 正しい列名に変更しました
        const dispatchTable = "new_Table2s"; 

        const selectCols = [
            "new_table2id",
            "new_start_time",       // 開始時間
            "new_kashikiri",        // 貸切区分
            "statuscode",           // ステータス
            "new_sharyou",          // 車両
            "new_tokuisaki_mei",    // ★修正: 得意先名 (アンダーバーあり)
            "new_genbamei",         // 現場名
            "new_sagyou_naiyou",    // ★修正: 作業内容 (配車テーブルの列を使用)
            "new_renraku_jikou"     // ★修正: 連絡事項 (配車テーブルの列を使用)
        ].join(",");

        // フィルタリング
        // 過去のデータも含めてテストしたい場合は日付フィルタをコメントアウトしてください
        const todayStr = new Date().toISOString().split('T')[0];
        
        let filter = `_owningbusinessunit_value eq ${myBusinessUnit}`;
        filter += ` and _new_operator_value eq ${myWorkerId}`; 
        // filter += ` and new_start_time ge ${todayStr}`; // いったん全件出して確認する場合はここを無効化

        const query = `?$select=${selectCols}&$filter=${filter}&$orderby=new_start_time asc`;
        
        const apiUrl = `${dataverseUrl}/api/data/v9.2/${dispatchTable}${query}`;
        const response = await fetch(apiUrl, { method: "GET", headers });
        
        if (!response.ok) {
            // エラー詳細を取得して返す
            const errorText = await response.text();
            let errorDetail = errorText;
            try {
                const errJson = JSON.parse(errorText);
                errorDetail = errJson.error?.message || errorText;
            } catch(e) {}
            throw new Error(`Dataverse Error (${dispatchTable}): ${response.status} - ${errorDetail}`);
        }
        
        const data = await response.json();

        // 6. 整形
        const results = data.value.map(item => {
            let timeStr = "--:--";
            if (item.new_start_time) {
                const dateObj = new Date(item.new_start_time);
                timeStr = dateObj.toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit', timeZone: 'Asia/Tokyo' });
            }

            return {
                id: item.new_table2id,
                time: timeStr,
                type: item["new_kashikiri@OData.Community.Display.V1.FormattedValue"] || "-",
                car: "代車 4958", // 車両名は後で実装
                // 以下、配車テーブルから直接取った値
                client: item.new_tokuisaki_mei || "名称なし",
                location: item.new_genbamei || "",
                workContent: item.new_sagyou_naiyou || "",
                notes: item.new_renraku_jikou || "",
                contact: "連絡先未設定",
                status: item["statuscode@OData.Community.Display.V1.FormattedValue"] || "未確認",
                statusCode: item.statuscode
            };
        });

        context.res = {
            status: 200,
            body: { 
                message: "Success", 
                userName: myName,
                businessUnitName: myBusinessUnitName,
                count: results.length,
                results: results 
            }
        };

    } catch (error) {
        context.log.error(error);
        context.res = { 
            status: 500, 
            body: { 
                error: "API Error", 
                details: error.message 
            } 
        };
    }
};
