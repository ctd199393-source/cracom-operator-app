export default async function (context, req) {
  context.log('Environment variable check invoked.');

  const clientId = process.env.ENTRA_CLIENT_ID || "(未設定)";
  const tenantId = process.env.ENTRA_TENANT_ID || "(未設定)";
  const hasSecret = !!process.env.ENTRA_CLIENT_SECRET;

  context.res = {
    status: 200,
    headers: { "Content-Type": "application/json" },
    body: {
      ENTRA_CLIENT_ID: clientId,
      ENTRA_TENANT_ID: tenantId,
      ENTRA_CLIENT_SECRET: hasSecret ? "(設定済み)" : "(未設定)"
    }
  };
}
