const { ClientSecretCredential } = require("@azure/identity");
const fetch = require("node-fetch");

module.exports = async function (context, req) {
    context.log('Dataverse API request started.');

    try {
        const tenantId = process.env.TENANT_ID; 
        const clientId = process.env.CLIENT_ID;
        const clientSecret = process.env.CLIENT_SECRET;
        const dataverseUrl = process.env.DATAVERSE_URL;

        if (!tenantId || !clientId || !clientSecret || !dataverseUrl) {
            throw new Error("環境変数が不足しています");
        }

        // -----------------------------------------------------------
        // 2. ユーザー情報 (メールアドレス) の確実な抽出
        // -----------------------------------------------------------
        let userEmail = req.headers["x-ms-client-principal-name"];
        const header = req.headers["x-ms-client-principal"];

        // (A) クレームデータがあれば、そこから 'email' を探す (一番確実)
        if (header) {
            try {
                const encoded = Buffer.from(header, "base64");
                const decoded = JSON.parse(encoded.toString("ascii"));
                const emailClaim = decoded.claims.find(c => c.typ === "email" || c.typ === "emails");
                if (emailClaim) {
                    userEmail = emailClaim.val;
                }
            } catch(e) {
                context.log("Failed to parse client principal header");
            }
        }

        // (B) UPNからの抽出 (フォールバック)
        // ゲストユーザー (例: c.t.d..._gmail.com#EXT#@cracom...) の場合
        if (userEmail && userEmail.includes("#EXT#")) {
            // #EXT# の前を取り出す -> c.t.d.1993.93_gmail.com
            let extracted = userEmail.split("#EXT#")[0];
            // 最後の "_gmail.com" の "_" を "@" に戻す簡易ロジック
            // (※完全な復元は難しいが、多くのケースで "_" を "@" に置換すれば通る)
            // ここでは単純に "gmail.com" などの主要ドメインの前の "_" を "@" に変える
            extracted = extracted.replace("_gmail.com", "@gmail.com")
                                 .replace("_yahoo.co.jp", "@yahoo.co.jp")
                                 .replace("_icloud.com", "@icloud.com");
            // 汎用的な置換 (最後のアンダースコアを@に)
            if (!extracted.includes("@")) {
                 const lastUnderscore = extracted.lastIndexOf("_");
                 if (lastUnderscore !== -1) {
                     extracted = extracted.substring(0, lastUnderscore) + "@" + extracted.substring(lastUnderscore + 1);
                 }
            }
            userEmail = extracted;
        }

        if (!userEmail) {
            context.res = { status: 401, body: { error: "ログインが必要です" } };
            return;
        }
        
        context.log(`Target Email for Search: ${userEmail}`);

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

        // 5. 配車データ取得
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
