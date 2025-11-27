const { ClientSecretCredential } = require("@azure/identity");
const fetch = require("node-fetch");

module.exports = async function (context, req) {
    try {
        const tenantId = process.env.TENANT_ID; 
        const clientId = process.env.CLIENT_ID;
        const clientSecret = process.env.CLIENT_SECRET;
        const dataverseUrl = process.env.DATAVERSE_URL;

        // 認証
        const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        const tokenResponse = await credential.getToken(`${dataverseUrl}/.default`);
        
        // ★診断: テーブル定義情報(メタデータ)を検索して、正しい「セット名」を探す
        // 探すテーブルの論理名（小文字でOK）
        const targetTables = ["new_table2", "new_sagyouin_masta"];
        
        // フィルタ条件作成
        const filter = targetTables.map(t => `LogicalName eq '${t.toLowerCase()}'`).join(" or ");
        const apiUrl = `${dataverseUrl}/api/data/v9.2/EntityDefinitions?$select=LogicalName,EntitySetName&$filter=${filter}`;
        
        const response = await fetch(apiUrl, {
            method: "GET",
            headers: { "Authorization": `Bearer ${tokenResponse.token}`, "Accept": "application/json" }
        });

        if (!response.ok) {
            throw new Error(`Metadata Query Error: ${response.status} ${await response.text()}`);
        }

        const data = await response.json();
        
        // 結果を整形して表示
        const resultList = data.value.map(t => ({
            "論理名 (LogicalName)": t.LogicalName,
            "★正解のセット名 (EntitySetName)": t.EntitySetName
        }));

        context.res = {
            status: 200,
            body: {
                message: "テーブル名の診断結果",
                tables: resultList
            }
        };

    } catch (error) {
        context.res = { status: 500, body: { error: "Diagnosis Failed", details: error.message } };
    }
};
