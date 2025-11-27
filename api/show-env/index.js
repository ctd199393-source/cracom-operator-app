const { ClientSecretCredential } = require("@azure/identity");
const fetch = require("node-fetch");

module.exports = async function (context, req) {
    context.log('Dataverse API request started.');

    try {
        // 1. 環境変数の読み込み (Azureの設定から)
        const tenantId = process.env.TENANT_ID; 
        const clientId = process.env.CLIENT_ID;
        const clientSecret = process.env.CLIENT_SECRET;
        const dataverseUrl = process.env.DATAVERSE_URL;

        // 設定漏れチェック
        if (!tenantId || !clientId || !clientSecret || !dataverseUrl) {
            throw new Error("環境変数が不足しています (TENANT_ID, CLIENT_ID, CLIENT_SECRET, DATAVERSE_URL)");
        }

        // 2. 認証 (厨房の合鍵を使用)
        const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        const tokenResponse = await credential.getToken(`${dataverseUrl}/.default`);
        const accessToken = tokenResponse.token;

        // 3. Dataverseデータ取得
        // 例: 取引先企業 (accounts) テーブルから名前を3件取得
        const tableName = "accounts"; 
        const query = "?$select=name&$top=3";
        const apiUrl = `${dataverseUrl}/api/data/v9.2/${tableName}${query}`;

        const response = await fetch(apiUrl, {
            method: "GET",
            headers: {
                "Authorization": `Bearer ${accessToken}`,
                "Accept": "application/json",
                "Content-Type": "application/json",
                "OData-MaxVersion": "4.0",
                "OData-Version": "4.0"
            }
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Dataverse Error: ${response.status} ${errorText}`);
        }

        const data = await response.json();

        // 4. 結果をブラウザに返す
        context.res = {
            status: 200,
            body: {
                message: "Dataverse接続成功！",
                results: data.value
            }
        };

    } catch (error) {
        context.log.error(error);
        context.res = {
            status: 500,
            body: {
                error: "Dataverse connection failed",
                details: error.message
            }
        };
    }
};
