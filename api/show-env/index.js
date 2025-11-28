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

        // -------------------------------------------------------------
        // ★重要修正：ログイン情報の詳細(ClientPrincipal)からメールアドレスを抽出
        // -------------------------------------------------------------
        let userEmail = req.headers["x-ms-client-principal-name"];
        const header = req.headers["x-ms-client-principal"];
        
        if (header) {
            const encoded = Buffer.from(header, "base64");
            const decoded = JSON.parse(encoded.toString("ascii"));
            
            // クレームの中から 'emails' または 'email' を探す
            const emailClaim = decoded.claims.find(c => c.typ === "emails" || c.typ === "email" || c.typ === "preferred_username");
            if (emailClaim) {
                userEmail = emailClaim.val;
                context.log(`Identified User Email via Claims: ${userEmail}`);
            }
        }

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

        // 4. 作業員マスタ検索
        const workerTable = "new_sagyouin_mastas"; 
        const workerQuery = `?$select=_owningbusinessunit_value,new_sagyouin_mastaid,new_mei&$filter=new_mail eq '${userEmail}'`;
        const workerRes = await fetch(`${dataverseUrl}/api/data/v9.2/${workerTable}${workerQuery}`, { method: "GET", headers });
        
        if (!workerRes.ok) {
            const err = await workerRes.text();
            throw new Error(`Worker Search Error: ${workerRes.status} ${err}`);
        }

        const workerData = await workerRes.json();

        if (!workerData.value || workerData.value.length === 0) {
            // どのメアドで検索して失敗したかをログに出す
            context.log(`User not found in Master: ${userEmail}`);
            context.res = { status: 403, body: { error: "作業員マスタに登録がありません" } };
            return;
        }

        const worker = workerData.value[0];
        const myBusinessUnit = worker._owningbusinessunit_value;
        const myWorkerId = worker.new_sagyouin_mastaid;
        const myName = worker.new_mei || "担当者";
        const myBusinessUnitName = worker["_owningbusinessunit_value@OData.Community.Display.V1.FormattedValue"] || "";

        // 5. 配車データ取得
        const dispatchTable = "new_Table2s"; 

        const selectCols = [
            "new_table2id",
            "new_start_time",       
            "new_kashikiri",        
            "statuscode",           
            "new_sharyou",          
            "new_tokuisaki_mei",    
            "new_genbamei",         
            "new_sagyou_naiyou",    
            "new_renraku_jikou"     
        ].join(",");

        const todayStr = new Date().toISOString().split('T')[0];
        
        let filter = `_owningbusinessunit_value eq ${myBusinessUnit}`;
        filter += ` and _new_operator_value eq ${myWorkerId}`; 
        // filter += ` and new_start_time ge ${todayStr}`; 

        const query = `?$select=${selectCols}&$filter=${filter}&$orderby=new_start_time asc`;
        const apiUrl = `${dataverseUrl}/api/data/v9.2/${dispatchTable}${query}`;

        const response = await fetch(apiUrl, { method: "GET", headers });
        
        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Dataverse Error (${dispatchTable}): ${response.status} - ${errorText}`);
        }
        
        const data = await response.json();

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
                car: "車両情報",
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
        context.res = { status: 500, body: { error: "API Error", details: error.message } };
    }
};
