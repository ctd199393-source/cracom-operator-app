const { ClientSecretCredential } = require("@azure/identity");
const fetch = require("node-fetch");

module.exports = async function (context, req) {
    context.log('Dataverse API request started.');

    try {
        // 1. 環境変数チェック
        const tenantId = process.env.TENANT_ID; 
        const clientId = process.env.CLIENT_ID;
        const clientSecret = process.env.CLIENT_SECRET;
        const dataverseUrl = process.env.DATAVERSE_URL;

        if (!tenantId || !clientId || !clientSecret || !dataverseUrl) {
            throw new Error("環境変数が不足しています");
        }

        // 2. ログインユーザーのメールアドレス取得
        const userEmail = req.headers["x-ms-client-principal-name"];
        
        if (!userEmail) {
            context.res = { status: 401, body: { error: "ログインが必要です" } };
            return;
        }

        // 3. Dataverse認証
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

        // --------------------------------------------------
        // Step A: 作業員マスタ検索 (セキュリティと部署特定)
        // --------------------------------------------------
        const workerTable = "new_sagyouin_mastas"; 
        const workerQuery = `?$select=_owningbusinessunit_value,new_sagyouin_mastaid,new_mei&$filter=new_mail eq '${userEmail}'`;
        
        const workerRes = await fetch(`${dataverseUrl}/api/data/v9.2/${workerTable}${workerQuery}`, { method: "GET", headers });
        
        if (!workerRes.ok) {
            // 作業員マスタのエラーかどうかの切り分け
            const err = await workerRes.text();
            throw new Error(`Worker Master Error: ${workerRes.status} ${err}`);
        }

        const workerData = await workerRes.json();

        if (!workerData.value || workerData.value.length === 0) {
            context.res = { status: 403, body: { error: "作業員マスタにあなたのメールアドレスが登録されていません" } };
            return;
        }

        const worker = workerData.value[0];
        const myBusinessUnit = worker._owningbusinessunit_value;
        const myWorkerId = worker.new_sagyouin_mastaid;
        const myName = worker.new_mei || "担当者";
        const myBusinessUnitName = worker["_owningbusinessunit_value@OData.Community.Display.V1.FormattedValue"] || "所属なし";

        // --------------------------------------------------
        // Step B: 配車データ取得
        // --------------------------------------------------
        // ★ここを修正しました: new_table2s -> new_Table2s (Tを大文字に)
        const dispatchTable = "new_Table2s"; 

        const selectCols = [
            "new_table2id",
            "new_start_time",       // 開始時間
            "new_kashikiri",        // 貸切区分
            "statuscode",           // ステータス
            "new_sharyou",          // 車両
            "_new_id_value"         // 案件ID
        ].join(",");

        // 案件情報の展開 (new_id)
        const expandAnken = `new_id($select=new_tokuisakimei,new_genbamei,new_bikou,new_renraku_jikou)`;
        
        // フィルタリング
        const todayStr = new Date().toISOString().split('T')[0];
        
        let filter = `_owningbusinessunit_value eq ${myBusinessUnit}`;
        filter += ` and _new_operator_value eq ${myWorkerId}`; 
        filter += ` and new_start_time ge ${todayStr}`;

        const query = `?$select=${selectCols}&$expand=${expandAnken}&$filter=${filter}&$orderby=new_start_time asc`;
        const apiUrl = `${dataverseUrl}/api/data/v9.2/${dispatchTable}${query}`;

        const response = await fetch(apiUrl, { method: "GET", headers });
        
        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Dataverse Error (${dispatchTable}): ${response.status} ${errorText}`);
        }
        
        const data = await response.json();

        // 4. データ整形
        const results = data.value.map(item => {
            const anken = item.new_id || {};
            const dateObj = new Date(item.new_start_time);
            const timeStr = dateObj.toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit', timeZone: 'Asia/Tokyo' });

            return {
                id: item.new_table2id,
                time: timeStr,
                type: item["new_kashikiri@OData.Community.Display.V1.FormattedValue"] || "-",
                car: "代車 4958",
                client: anken.new_tokuisakimei || "名称未設定",
                location: anken.new_genbamei || "",
                workContent: anken.new_bikou || "",
                notes: anken.new_renraku_jikou || "",
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
