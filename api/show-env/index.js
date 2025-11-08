module.exports = async function (context, req) {
  context.log('JavaScript HTTP trigger function processed a request.');

  const clientId = process.env.ENTRA_CLIENT_ID || "（CLIENT_IDが設定されていません）";
  const tenantId = process.env.ENTRA_TENANT_ID || "（TENANT_IDが設定されていません）";

  // 注意： CLIENT_SECRET はセキュリティ上、絶対に表示させてはいけません。
  // SWAの仕様により、そもそもこのAPIには渡されません（nullになります）。

  context.res = {
    // status: 200, /* Defaults to 200 */
    headers: {
        'Content-Type': 'application/json'
    },
    body: {
      ENTRA_CLIENT_ID_Check: clientId,
      ENTRA_TENANT_ID_Check: tenantId,
      ENTRA_CLIENT_SECRET_Check: "（シークレットはセキュリティ上表示できません）"
    }
  };
};
