const { ClientSecretCredential } = require("@azure/identity");
const fetch = require("node-fetch");

module.exports = async function (context, req) {
    context.log('▼ Login Check Start');

    try {
        // -----------------------------------------------------------
        // 0. 環境変数のチェック
        // -----------------------------------------------------------
        const tenantId = process.env.TENANT_ID; 
        const clientId = process.env.CLIENT_ID;
        const clientSecret = process.env.CLIENT_SECRET;
        const dataverseUrl = process.env.DATAVERSE_URL;

        if (!tenantId || !clientId || !clientSecret || !dataverseUrl) {
            throw new Error("サーバー設定エラー: 環境変数(TENANT_ID等)が不足しています。");
        }

        // -----------------------------------------------------------
        // 1. ユーザー情報の抽出と正規化
        // -----------------------------------------------------------
        let userEmail = null;
        let rawPrincipalName = req.headers["x-ms-client-principal-name"]; // SWAがヘッダーに付与する生のID
        const header = req.headers["x-ms-client-principal"]; // 詳細情報（Base64エンコード）

        context.log(`Raw Principal Name: ${rawPrincipalName}`);

        // (A) 詳細情報(クレーム)からメールアドレスを探す
        if (header) {
            try {
                const encoded = Buffer.from(header, "base64");
                const decoded = JSON.parse(encoded.toString("ascii"));
                
                // デバッグ用にクレーム情報をログに出す（本番運用時は削除可）
                // context.log(`Decoded Claims: ${JSON.stringify(decoded)}`);

                // 優先順位: userDetails -> email -> emails -> name
                userEmail = decoded.userDetails 
                    || (decoded.claims && decoded.claims.find(c => c.typ === "email")?.val)
                    || (decoded.claims && decoded.claims.find(c => c.typ === "emails")?.val)
                    || (decoded.claims && decoded.claims.find(c => c.typ === "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress")?.val)
                    || (decoded.claims && decoded.claims.find(c => c.typ === "name")?.val);
            } catch(e) {
                context.log.error("Header decode warning: " + e.message);
            }
        }

        // (B) ヘッダーから直接取れる場合のバックアップ
        if (!userEmail && rawPrincipalName) {
            userEmail = rawPrincipalName;
        }

        // (C) ゲストユーザー特有のゴミ(#EXT#)除去処理
        // 例: user_gmail.com#EXT#@cracomsystem.onmicrosoft.com -> user@gmail.com
        if (userEmail) {
            // 大文字小文字区別なく #EXT# を含むかチェック
            if (userEmail.toUpperCase().includes("#EXT#")) {
                context.log(`Guest User Detected. Raw: ${userEmail}`);
                
                // #EXT# より前の部分を取得 (例: user_gmail.com)
                let extracted = userEmail.split("#")[0]; 
                
                // 最後のアンダースコア(_)をアットマーク(@)に置換
                const lastUnderscore = extracted.lastIndexOf("_");
                if (lastUnderscore !== -1) {
                     extracted = extracted.substring(0, lastUnderscore) + "@" + extracted.substring(lastUnderscore + 1);
                }
                userEmail = extracted;
                context.log(`Normalized Email: ${userEmail}`);
            } else {
                context.log(`Standard User Email: ${userEmail}`);
            }
        }

        // 抽出失敗時の処理
        if (!userEmail) {
            context.res = { 
                status: 401, 
                body: { error: "ログイン情報が取得できません。画面をリロードするか再ログインしてください。" } 
            };
            return;
        }

        // -----------------------------------------------------------
        // 2. Dataverse 認証 & 検索
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

        // 作業員マスタ検索
        // new_mail が userEmail と一致するものを探す
        const workerTable = "new_sagyouin_mastas"; 
        const workerQuery = `?$select=_owningbusinessunit_value,new_sagyouin_mastaid,new_mei,new_mail&$filter=new_mail eq '${userEmail}'`;
        
        context.log(`Querying Dataverse: ${workerQuery}`);
        const workerRes = await fetch(`${dataverseUrl}/api/data/v9.2/${workerTable}${workerQuery}`, { method: "GET", headers });
        
        if (!workerRes.ok) {
            throw new Error(`Dataverse Search Error: ${workerRes.status} ${workerRes.statusText}`);
        }
        
        const workerData = await workerRes.json();

        // マスタ登録なしの場合
        if (!workerData.value || workerData.value.length === 0) {
            context.log(`User not found in Dataverse: ${userEmail}`);
            context.res = { 
                status: 403, 
                body: { 
                    error: "NoRegistration", // フロントエンド側での判別コード
                    details: `メールアドレス [${userEmail}] は作業員マスタに登録されていません。管理者に連絡してください。`,
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

        context.log(`User Found: ${myName} (${myBusinessUnitName})`);

        // -----------------------------------------------------------
        // 3. 配車データ取得 (所属部署 & 担当者フィルタ)
        // -----------------------------------------------------------
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

        // 今日の日付以降などを入れたい場合はここに追記可能ですが、まずは全件取得でテスト
        const query = `?$select=${selectCols}&$filter=${filter}&$orderby=new_start_time asc`;
        const apiUrl = `${dataverseUrl}/api/data/v9.2/${dispatchTable}${query}`;

        const response = await fetch(apiUrl, { method: "GET", headers });
        
        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Dataverse GetJobs Error: ${response.status} - ${errorText}`);
        }
        
        const data = await response.json();

        // -----------------------------------------------------------
        // 4. データ整形して返却
        // -----------------------------------------------------------
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
                car: "代車", // 車番が必要ならここもDataverseから取得・結合が必要
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
                userEmail: userEmail, // デバッグ確認用にメアドも返す
                count: results.length, 
                results: results 
            }
        };

    } catch (error) {
        context.log.error(error);
        context.res = { 
            status: 500, 
            body: { error: "SystemError", details: error.message } 
        };
    }
};
