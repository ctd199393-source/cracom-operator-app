const { ClientSecretCredential } = require("@azure/identity");
const fetch = require("node-fetch");

module.exports = async function (context, req) {
    context.log("API Triggered: show-env (Debug Mode)");

    try {
        // 1. 環境変数のチェック
        const tenantId = process.env.TENANT_ID;
        const clientId = process.env.CLIENT_ID;
        const clientSecret = process.env.CLIENT_SECRET;
        const dataverseUrl = process.env.DATAVERSE_URL;

        // 変数が空ならエラーを出す
        if (!tenantId) throw new Error("環境変数 TENANT_ID が設定されていません");
        if (!clientId) throw new Error("環境変数 CLIENT_ID が設定されていません");
        if (!clientSecret) throw new Error("環境変数 CLIENT_SECRET が設定されていません");
        if (!dataverseUrl) throw new Error("環境変数 DATAVERSE_URL が設定されていません");

        // 2. ユーザー情報の取得
        const header = req.headers["x-ms-client-principal"];
        let rawEmail = "unknown";
        let searchEmail = "";

        if (header) {
            const decoded = JSON.parse(Buffer.from(header, "base64").toString("ascii"));
            rawEmail = decoded.userDetails || "unknown";
        } else {
            // ローカルテスト用などのフォールバック（今回は本番なのでエラーでもよいがログ出す）
            context.log("No x-ms-client-principal header found.");
        }

        // メールアドレスの正規化 (#EXT# 対策)
        if (rawEmail.includes("#EXT#")) {
            let temp = rawEmail.split("#EXT#")[0];
            const lastUnderscoreIndex = temp.lastIndexOf("_");
            if (lastUnderscoreIndex !== -1) {
                searchEmail = temp.substring(0, lastUnderscoreIndex) + "@" + temp.substring(lastUnderscoreIndex + 1);
            } else {
                searchEmail = temp;
            }
        } else {
            searchEmail = rawEmail;
        }

        context.log(`Searching Dataverse for: ${searchEmail}`);

        // 3. Dataverse 接続
        const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        // トークン取得を試みる
        const tokenResponse = await credential.getToken(`${dataverseUrl}/.default`);
        const accessToken = tokenResponse.token;

        // 4. Dataverse検索
        // エラー詳細確認のため、try-catchをここでも強化
        const filter = `emailaddress1 eq '${searchEmail}'`; 
        const queryUrl = `${dataverseUrl}/api/data/v9.2/new_sagyouin_mastas?$filter=${encodeURIComponent(filter)}&$select=new_name,_new_businessunit_value`;

        const dvRes = await fetch(queryUrl, {
            method: "GET",
            headers: {
                "Authorization": `Bearer ${accessToken}`,
                "Accept": "application/json",
                "OData-MaxVersion": "4.0",
                "OData-Version": "4.0"
            }
        });

        if (!dvRes.ok) {
            const errText = await dvRes.text();
            throw new Error(`Dataverse API Error (${dvRes.status}): ${errText}`);
        }

        const dvData = await dvRes.json();

        // 5. データ判定
        if (dvData.value.length === 0) {
            context.res = { 
                status: 403, 
                body: { error: `メールアドレス (${searchEmail}) がマスタに見つかりません。Dataverseを確認してください。` } 
            };
            return;
        }

        const userRecord = dvData.value[0];
        const businessUnitId = userRecord._new_businessunit_value;

        // 6. 配車データ取得
        const dispatchQuery = `${dataverseUrl}/api/data/v9.2/new_table2s?$filter=_craca_businessunit_value eq '${businessUnitId}'`;
        const dispatchRes = await fetch(dispatchQuery, {
            headers: { "Authorization": `Bearer ${accessToken}` }
        });
        
        if (!dispatchRes.ok) {
             const dispErr = await dispatchRes.text();
             throw new Error(`Dispatch Table Error: ${dispErr}`);
        }
        
        const dispatchData = await dispatchRes.json();

        context.res = {
            status: 200,
            body: {
                user: userRecord,
                records: dispatchData.value,
                debugEmail: searchEmail // デバッグ用に返却
            }
        };

    } catch (error) {
        context.log.error(error);
        // ★ここが重要：生のエラーメッセージを画面に返す
        context.res = { 
            status: 500, 
            body: { 
                error: `システムエラー詳細: ${error.message}`,
                stack: error.stack 
            } 
        };
    }
};
