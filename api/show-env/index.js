const { ClientSecretCredential } = require("@azure/identity");
const fetch = require("node-fetch");

module.exports = async function (context, req) {
    context.log("API Triggered: show-env");

    try {
        // 1. 環境変数の取得
        const tenantId = process.env.TENANT_ID;
        const clientId = process.env.CLIENT_ID;
        const clientSecret = process.env.CLIENT_SECRET;
        const dataverseUrl = process.env.DATAVERSE_URL;

        if (!tenantId || !clientId || !clientSecret || !dataverseUrl) {
            throw new Error("Server configurations are missing.");
        }

        // 2. ユーザー情報の取得と正規化
        const header = req.headers["x-ms-client-principal"];
        let rawEmail = "unknown";
        let searchEmail = "";

        if (header) {
            const decoded = JSON.parse(Buffer.from(header, "base64").toString("ascii"));
            // SWA認証では userDetails にメールアドレス(またはUPN)が入る
            rawEmail = decoded.userDetails; 
        }

        if (!rawEmail || rawEmail === "unknown") {
            context.res = { status: 401, body: { error: "ユーザー情報を取得できませんでした。" } };
            return;
        }

        // ★ ゲストユーザー特有の #EXT# 形式を元のメアドに戻す処理
        // 例: user_gmail.com#EXT#@cracomsystem.onmicrosoft.com -> user@gmail.com
        if (rawEmail.includes("#EXT#")) {
            // #EXT# より前を取得
            let temp = rawEmail.split("#EXT#")[0];
            // 最後の _ を @ に戻す (Azure ADの仕様に基づく変換)
            // ただし、単純な置換だと email_with_underscore@gmail.com で誤爆する可能性があるため
            // 一般的には「末尾のドメイン部分の直前のアンダースコア」を置換しますが、
            // 簡易的に「最後のアンダースコアを@にする」で実装します。
            const lastUnderscoreIndex = temp.lastIndexOf("_");
            if (lastUnderscoreIndex !== -1) {
                searchEmail = temp.substring(0, lastUnderscoreIndex) + "@" + temp.substring(lastUnderscoreIndex + 1);
            } else {
                searchEmail = temp; // 変換不能な場合はそのまま
            }
        } else {
            // 社内ユーザーなどはそのまま
            searchEmail = rawEmail;
        }

        context.log(`Searching Dataverse for: ${searchEmail} (Raw: ${rawEmail})`);

        // 3. Dataverse 接続設定
        const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        const tokenResponse = await credential.getToken(`${dataverseUrl}/.default`);
        const accessToken = tokenResponse.token;

        // 4. Dataverse検索 (作業員マスタ)
        // ※emailaddress1 は標準的なメール列名です。実際の列名に合わせてください (例: new_email)
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
            throw new Error(`Dataverse Error: ${dvRes.status} ${errText}`);
        }

        const dvData = await dvRes.json();

        // 5. 判定とデータ取得
        if (dvData.value.length === 0) {
            context.log("User not found in Dataverse.");
            context.res = { 
                status: 403, 
                body: { error: `メールアドレス (${searchEmail}) は作業員マスタに登録されていません。` } 
            };
            return;
        }

        const userRecord = dvData.value[0];
        const businessUnitId = userRecord._new_businessunit_value;

        // 6. 配車データの取得 (BusinessUnitでフィルタ)
        // ※new_table2 は実際の配車テーブル名に、craca_businessunit は実際の関連列名に修正してください
        const dispatchQuery = `${dataverseUrl}/api/data/v9.2/new_table2s?$filter=_craca_businessunit_value eq '${businessUnitId}'`;
        
        const dispatchRes = await fetch(dispatchQuery, {
            headers: { "Authorization": `Bearer ${accessToken}` }
        });
        const dispatchData = await dispatchRes.json();

        // 7. 結果返却
        context.res = {
            status: 200,
            body: {
                user: userRecord,
                records: dispatchData.value
            }
        };

    } catch (error) {
        context.log.error(error);
        context.res = { status: 500, body: { error: "サーバー内部エラーが発生しました。" } };
    }
};
