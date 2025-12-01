const { ClientSecretCredential } = require("@azure/identity");
const fetch = require("node-fetch");

module.exports = async function (context, req) {
    context.log('▼ API Request Started');

    try {
        // -----------------------------------------------------------
        // 0. 環境変数の取得 (Dataverse接続用)
        // -----------------------------------------------------------
        const tenantId = process.env.TENANT_ID; 
        const clientId = process.env.CLIENT_ID;
        const clientSecret = process.env.CLIENT_SECRET;
        const dataverseUrl = process.env.DATAVERSE_URL;

        if (!tenantId || !clientId || !clientSecret || !dataverseUrl) {
            throw new Error("API Config Error: 環境変数が不足しています。");
        }

        // -----------------------------------------------------------
        // 1. ユーザーのメールアドレスを「確実に」取得する処理
        // -----------------------------------------------------------
        let userEmail = null;
        const header = req.headers["x-ms-client-principal"];

        if (header) {
            try {
                const encoded = Buffer.from(header, "base64");
                const decoded = JSON.parse(encoded.toString("ascii"));
                
                // デバッグ用: どんな情報が来ているかログに残す（後で確認可能）
                context.log(`User Claims: ${JSON.stringify(decoded.claims)}`);
                context.log(`User Details: ${decoded.userDetails}`);

                // 【重要】以下の優先順位で「メールアドレス」を探す
                // 1. claim: "email" (一番確実)
                // 2. claim: "emails" (複数形の場合あり)
                // 3. claim: "preferred_username" (UPNだが、メール形式の場合がある)
                // 4. claim: "name" (メールの場合がある)
                // 5. userDetails (SWAが判定したID。ここがランダムIDの場合があるので優先度を下げる)
                
                const claims = decoded.claims || [];
                
                userEmail = claims.find(c => c.typ === "email" || c.typ === "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress")?.val
                         || claims.find(c => c.typ === "emails")?.val
                         || claims.find(c => c.typ === "preferred_username")?.val
                         || claims.find(c => c.typ === "name")?.val
                         || decoded.userDetails;

            } catch(e) {
                context.log.error("Header decode failed: " + e.message);
            }
        }

        // もしヘッダーから取れなければ、SWAのデフォルトIDを使う（最終手段）
        if (!userEmail) {
            userEmail = req.headers["x-ms-client-principal-name"];
        }

        // メールアドレスの正規化（ゲストユーザー等のゴミ除去）
        if (userEmail && userEmail.includes("#EXT#")) {
            // 例: user_gmail.com#EXT#@... -> user@gmail.com
            userEmail = userEmail.split("#")[0].replace(/_$/, "").replace(/_/, "@");
        }

        // ★ここで「ランダムID（UUID）」が残っていないか最終チェック
        // メールアドレス形式（@があるか）でなければエラーにする
        if (!userEmail || !userEmail.includes("@")) {
            context.log(`Invalid Email Detected: ${userEmail}`);
            context.res = { status: 403, body: { error: "InvalidID", details: `システムがメールアドレスを特定できませんでした。(ID: ${userEmail})` } };
            return;
        }

        context.log(`Target Email: ${userEmail}`);

        // -----------------------------------------------------------
        // 2. Dataverse 検索
        // -----------------------------------------------------------
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

        const workerTable = "new_sagyouin_mastas"; 
        const workerQuery = `?$select=_owningbusinessunit_value,new_sagyouin_mastaid,new_mei,new_mail&$filter=new_mail eq '${userEmail}'`;
        
        const workerRes = await fetch(`${dataverseUrl}/api/data/v9.2/${workerTable}${workerQuery}`, { method: "GET", headers });
        if (!workerRes.ok) throw new Error(`Dataverse Error: ${workerRes.status}`);
        
        const workerData = await workerRes.json();

        if (!workerData.value || workerData.value.length === 0) {
            context.res = { 
                status: 403, 
                body: { 
                    error: "NoRegistration", 
                    details: `メールアドレス [${userEmail}] はマスタにありません。`,
                    detectedEmail: userEmail
                } 
            };
            return;
        }

        const worker = workerData.value[0];
        const myBusinessUnit = worker._owningbusinessunit_value;
        const myWorkerId = worker.new_sagyouin_mastaid;
        const myName = worker.new_mei || "担当者";
        const myBusinessUnitName = worker["_owningbusinessunit_value@OData.Community.Display.V1.FormattedValue"] || "";

        // -----------------------------------------------------------
        // 3. 配車データ取得
        // -----------------------------------------------------------
        const dispatchTable = "new_table2s"; 
        const selectCols = "new_table2id,new_start_time,new_kashikiri,statuscode,new_sharyou,new_tokuisaki_mei,new_genbamei,new_sagyou_naiyou,new_renraku_jikou";
        
        let filter = `_owningbusinessunit_value eq ${myBusinessUnit} and _new_operator_value eq ${myWorkerId}`; 
        const query = `?$select=${selectCols}&$filter=${filter}&$orderby=new_start_time asc`;
        
        const jobsRes = await fetch(`${dataverseUrl}/api/data/v9.2/${dispatchTable}${query}`, { method: "GET", headers });
        if (!jobsRes.ok) throw new Error(`Jobs Error: ${jobsRes.status}`);
        
        const data = await jobsRes.json();

        const results = data.value.map(item => {
            let timeStr = "--:--";
            if (item.new_start_time) {
                timeStr = new Date(item.new_start_time).toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit', timeZone: 'Asia/Tokyo' });
            }
            return {
                id: item.new_table2id,
                time: timeStr,
                type: item["new_kashikiri@OData.Community.Display.V1.FormattedValue"] || "-",
                car: "代車",
                client: item.new_tokuisaki_mei || "名称なし",
                location: item.new_genbamei || "",
                workContent: item.new_sagyou_naiyou || "",
                notes: item.new_renraku_jikou || "",
                contact: "",
                status: item["statuscode@OData.Community.Display.V1.FormattedValue"] || "未確認",
                statusCode: item.statuscode
            };
        });

        context.res = {
            status: 200,
            body: { message: "Success", userName: myName, businessUnitName: myBusinessUnitName, count: results.length, results: results }
        };

    } catch (error) {
        context.log.error(error);
        context.res = { status: 500, body: { error: "SystemError", details: error.message } };
    }
};
