const { ClientSecretCredential } = require("@azure/identity");
const fetch = require("node-fetch");

module.exports = async function (context, req) {
    context.log('Dataverse API request started.');

    try {
        // --------------------------------------------------
        // 1. 設定と準備
        // --------------------------------------------------
        const tenantId = process.env.TENANT_ID; 
        const clientId = process.env.CLIENT_ID;
        const clientSecret = process.env.CLIENT_SECRET;
        const dataverseUrl = process.env.DATAVERSE_URL;

        if (!tenantId || !clientId || !clientSecret || !dataverseUrl) {
            throw new Error("環境変数が不足しています");
        }

        // ログインユーザーのメールアドレス
        const userEmail = req.headers["x-ms-client-principal-name"];
        
        if (!userEmail) {
            context.res = { status: 401, body: { error: "ログインが必要です" } };
            return;
        }

        // 認証トークン取得
        const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        const tokenResponse = await credential.getToken(`${dataverseUrl}/.default`);
        const accessToken = tokenResponse.token;
        
        const headers = {
            "Authorization": `Bearer ${accessToken}`,
            "Accept": "application/json",
            "Content-Type": "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            "Prefer": "odata.include-annotations=\"*\"" // これにより部署名(FormattedValue)が取得できます
        };

        // --------------------------------------------------
        // 2. 作業員マスタ検索 (セキュリティと部署特定)
        // --------------------------------------------------
        const workerTable = "new_sagyouin_mastas"; 
        const workerQuery = `?$select=_owningbusinessunit_value,new_sagyouin_mastaid,new_mei&$filter=new_mail eq '${userEmail}'`;
        
        const workerRes = await fetch(`${dataverseUrl}/api/data/v9.2/${workerTable}${workerQuery}`, { method: "GET", headers });
        const workerData = await workerRes.json();

        if (!workerData.value || workerData.value.length === 0) {
            context.res = { status: 403, body: { error: "作業員マスタに登録がありません" } };
            return;
        }

        const worker = workerData.value[0];
        const myBusinessUnit = worker._owningbusinessunit_value;
        const myWorkerId = worker.new_sagyouin_mastaid;
        const myName = worker.new_mei || "担当者";
        
        // ★追加: 部署名の取得 (FormattedValueを使用)
        const myBusinessUnitName = worker["_owningbusinessunit_value@OData.Community.Display.V1.FormattedValue"] || "所属なし";

        // --------------------------------------------------
        // 3. 配車データ取得 (配車テーブル + 案件情報)
        // --------------------------------------------------
        const dispatchTable = "new_table2s"; 

        // 取得したい列
        const selectCols = [
            "new_table2id",
            "new_start_time",
            "new_kashikiri",
            "statuscode",
            "new_sharyou",
            "_new_id_value"
        ].join(",");

        // 案件情報の展開
        const expandAnken = `new_id($select=new_tokuisakimei,new_genbamei,new_bikou,new_renraku_jikou)`;
        
        // フィルタリング (部署 && 担当者 && 日付)
        const todayStr = new Date().toISOString().split('T')[0];
        
        let filter = `_owningbusinessunit_value eq ${myBusinessUnit}`;
        filter += ` and _new_operator_value eq ${myWorkerId}`; 
        filter += ` and new_start_time ge ${todayStr}`;

        const query = `?$select=${selectCols}&$expand=${expandAnken}&$filter=${filter}&$orderby=new_start_time asc`;
        
        const apiUrl = `${dataverseUrl}/api/data/v9.2/${dispatchTable}${query}`;
        const response = await fetch(apiUrl, { method: "GET", headers });
        
        if (!response.ok) {
            throw new Error(`Dataverse Error: ${response.status} ${await response.text()}`);
        }
        
        const data = await response.json();

        // --------------------------------------------------
        // 4. データ整形
        // --------------------------------------------------
        const results = data.value.map(item => {
            const anken = item.new_id || {};
            const dateObj = new Date(item.new_start_time);
            const timeStr = dateObj.toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit', timeZone: 'Asia/Tokyo' });

            return {
                id: item.new_table2id,
                time: timeStr,
                type: item["new_kashikiri@OData.Community.Display.V1.FormattedValue"] || "-",
                car: "代車 4958", // ※車両テーブルとの連携が必要であれば後で実装
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
                businessUnitName: myBusinessUnitName, // ★画面に渡す部署名
                results: results 
            }
        };

    } catch (error) {
        context.log.error(error);
        context.res = { status: 500, body: { error: error.message } };
    }
};
