const { ClientSecretCredential } = require("@azure/identity");
const fetch = require("node-fetch");

module.exports = async function (context, req) {
    context.log("API Triggered: show-env (Corrected Schema)");

    try {
        // 1. 環境変数のチェック
        const tenantId = process.env.TENANT_ID;
        const clientId = process.env.CLIENT_ID;
        const clientSecret = process.env.CLIENT_SECRET;
        const dataverseUrl = process.env.DATAVERSE_URL;

        if (!tenantId || !clientId || !clientSecret || !dataverseUrl) {
            throw new Error("環境変数が不足しています。SWAの設定を確認してください。");
        }

        // 2. ユーザー情報の取得と正規化
        const header = req.headers["x-ms-client-principal"];
        let rawEmail = "unknown";
        let searchEmail = "";

        if (header) {
            const decoded = JSON.parse(Buffer.from(header, "base64").toString("ascii"));
            rawEmail = decoded.userDetails || "unknown";
        }

        // #EXT# 除去処理
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
        const tokenResponse = await credential.getToken(`${dataverseUrl}/.default`);
        const accessToken = tokenResponse.token;

        // 4. Dataverse検索 (作業員マスタ)
        // ★修正ポイント: 列名を実際の定義に合わせて変更しました
        // emailaddress1 -> new_mail
        const filter = `new_mail eq '${searchEmail}'`; 
        
        // select指定: new_name -> new_sagyouin_id (作業員名), new_businessunit -> owningbusinessunit (所属部署)
        const queryUrl = `${dataverseUrl}/api/data/v9.2/new_sagyouin_mastas?$filter=${encodeURIComponent(filter)}&$select=new_sagyouin_id,_owningbusinessunit_value`;

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
                body: { error: `メールアドレス (${searchEmail}) が作業員マスタに見つかりません。Dataverseの 'new_mail' 列を確認してください。` } 
            };
            return;
        }

        const userRecord = dvData.value[0];
        // ★修正ポイント: 取得した所属部署IDを取り出す
        const businessUnitId = userRecord._owningbusinessunit_value;

        // 6. 配車データ取得 (new_table2s)
        // ※注意: ここの 'new_table2s' と '_craca_businessunit_value' も、もしエラーが出たら正しい名前に直す必要があります。
        // 一旦、前のコードのまま進めます。
        const dispatchQuery = `${dataverseUrl}/api/data/v9.2/new_table2s?$filter=_craca_businessunit_value eq '${businessUnitId}'`;
        
        const dispatchRes = await fetch(dispatchQuery, {
            headers: { "Authorization": `Bearer ${accessToken}` }
        });

        // 配車テーブル側でエラーが出た場合のハンドリング
        if (!dispatchRes.ok) {
             const dispErr = await dispatchRes.text();
             // ここでエラーが出たら、new_table2s の列名も確認が必要です
             throw new Error(`配車データ取得エラー: ${dispErr}`);
        }
        
        const dispatchData = await dispatchRes.json();

        context.res = {
            status: 200,
            body: {
                user: {
                    name: userRecord.new_sagyouin_id, // 作業員名
                    bu: businessUnitId
                },
                records: dispatchData.value,
                debugEmail: searchEmail
            }
        };

    } catch (error) {
        context.log.error(error);
        context.res = { 
            status: 500, 
            body: { 
                error: `システムエラー: ${error.message}`,
                stack: error.stack
            } 
        };
    }
};
