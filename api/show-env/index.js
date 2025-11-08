module.exports = function (context, req) {
    context.log('JavaScript HTTP trigger function processed a request.');

    // Azureポータル（SWA）の「環境変数」から値を取得
    const clientId = process.env.ENTRA_CLIENT_ID || "（CLIENT_IDが設定されていません）";
    const tenantId = process.env.ENTRA_TENANT_ID || "（TENANT_IDが設定されていません）";

    // JSON形式で結果を返す
    context.res.json({
        ENTRA_CLIENT_ID_Check: clientId,
        ENTRA_TENANT_ID_Check: tenantId,
        ENTRA_CLIENT_SECRET_Check: "（シークレットはセキュリティ上、このAPIでは確認できません）"
    });
};
