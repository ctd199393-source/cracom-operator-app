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

        // -----------------------------------------------------------
        // 2. ユーザー情報の「メールアドレス」を確実に取り出すロジック
        // -----------------------------------------------------------
        let userEmail = null;
        const header = req.headers["x-ms-client-principal"];

        if (header) {
            try {
                // Base64デコードして中身(JSON)を取り出す
                const encoded = Buffer.from(header, "base64");
                const decoded = JSON.parse(encoded.toString("ascii"));
                
                // クレームの中から「メールアドレス」を探す
                // 優先順位: email > emails > name > preferred_username
                const claims = decoded.claims;
                const emailClaim = claims.find(c => c.typ === "email" || c.typ === "emails" || c.typ === "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress");
                
                if (emailClaim) {
                    userEmail = emailClaim.val;
                } else {
                    // メールが見つからない場合、nameクレームを使う（ここがメアドの場合がある）
                    const nameClaim = claims.find(c => c.typ === "name");
                    if (nameClaim) userEmail = nameClaim.val;
                }
            } catch(e) {
                context.log("Header decode failed: " + e.message);
            }
        }

        // ヘッダー解読に失敗、またはローカルでテストしている場合のフォールバック
        if (!userEmail) {
            userEmail = req.headers["x-ms-client-principal-name"];
        }

        // ★重要: それでもまだ "live.com#..." や乱数IDの可能性がある場合の最終整形
        if (userEmail && userEmail.includes("#EXT#")) {
             // ゲストユーザーの場合の整形
             let extracted = userEmail.split("#EXT#")[0];
             const lastUnderscore = extracted.lastIndexOf("_");
             if (lastUnderscore !== -1) {
                 extracted = extracted.substring(0, lastUnderscore) + "@" + extracted.substring(lastUnderscore + 1);
            }
            userEmail = extracted;
        }

        if (!userEmail) {
            context.res = { status: 401, body: { error: "ログインが必要です" } };
            return;
        }
        
        // ログ出力：ここで正しいメアド（gmail.com）になっているか確認できます
        context.log(`Searching Dataverse for User: ${userEmail}`);


        // -----------------------------------------------------------
        // 3. Dataverse認証
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

        // 4. 作業員マスタ検索
        const workerTable = "new_sagyouin_mastas"; 
        const workerQuery = `?$select=_owningbusinessunit_value,new_sagyouin_mastaid,new_mei&$filter=new_mail eq '${userEmail}'`;
        const workerRes = await fetch(`${dataverseUrl}/api/data/v9.2/${workerTable}${workerQuery}`, { method: "GET", headers });
        
        if (!workerRes.ok) throw new Error(`Worker Search Error: ${workerRes.status}`);
        const workerData = await workerRes.json();

        // マスタにいない場合
        if (!workerData.value || workerData.value.length === 0) {
            context.log(`User not found in master. Input was: ${userEmail}`);
            context.res = { status: 403, body: { error: "作業員マスタに登録がありません" } };
            return;
        }

        const worker = workerData.value[0];
        const myBusinessUnit = worker._owningbusinessunit_value;
        const myWorkerId = worker.new_sagyouin_mastaid;
        const myName = worker.new_mei || "担当者";
        const myBusinessUnitName = worker["_owningbusinessunit_value@OData.Community.Display.V1.FormattedValue"] || "";

        // 5. 配車データ取得 (正しいテーブル名: new_table2s)
        const dispatchTable = "new_table2s"; 

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
        
        // フィルタ: 部署一致 && 自分担当
        let filter = `_owningbusinessunit_value eq ${myBusinessUnit}`;
        filter += ` and _new_operator_value eq ${myWorkerId}`; 

        const query = `?$select=${selectCols}&$filter=${filter}&$orderby=new_start_time asc`;
        const apiUrl = `${dataverseUrl}/api/data/v9.2/${dispatchTable}${query}`;

        const response = await fetch(apiUrl, { method: "GET", headers });
        
        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Dataverse Error (${dispatchTable}): ${response.status} - ${errorText}`);
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
                car: "代車 4958",
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
