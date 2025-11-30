const { ClientSecretCredential } = require("@azure/identity");
const fetch = require("node-fetch");

module.exports = async function (context, req) {
    context.log('▼ Login Check Start');
    try {
        const tenantId = process.env.ENTRA_TENANT_ID || process.env.TENANT_ID; 
        const clientId = process.env.ENTRA_CLIENT_ID || process.env.CLIENT_ID;
        const clientSecret = process.env.ENTRA_CLIENT_SECRET || process.env.CLIENT_SECRET;
        const dataverseUrl = process.env.DATAVERSE_URL;

        if (!tenantId || !clientId || !clientSecret || !dataverseUrl) throw new Error("環境変数が不足しています。");

        let userEmail = null;
        let rawPrincipalName = req.headers["x-ms-client-principal-name"];
        const header = req.headers["x-ms-client-principal"];

        if (header) {
            try {
                const encoded = Buffer.from(header, "base64");
                const decoded = JSON.parse(encoded.toString("ascii"));
                userEmail = decoded.userDetails 
                    || (decoded.claims && decoded.claims.find(c => c.typ === "email")?.val)
                    || (decoded.claims && decoded.claims.find(c => c.typ === "emails")?.val)
                    || (decoded.claims && decoded.claims.find(c => c.typ === "name")?.val);
            } catch(e) { context.log.error(e); }
        }
        if (!userEmail && rawPrincipalName) userEmail = rawPrincipalName;

        if (userEmail && userEmail.toUpperCase().includes("#EXT#")) {
            let extracted = userEmail.split("#")[0]; 
            const lastUnderscore = extracted.lastIndexOf("_");
            if (lastUnderscore !== -1) {
                 extracted = extracted.substring(0, lastUnderscore) + "@" + extracted.substring(lastUnderscore + 1);
            }
            userEmail = extracted;
        }

        if (!userEmail) {
            context.res = { status: 401, body: { error: "ログイン情報がありません" } };
            return;
        }

        const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        const tokenResponse = await credential.getToken(`${dataverseUrl}/.default`);
        const accessToken = tokenResponse.token;
        
        const headers = {
            "Authorization": `Bearer ${accessToken}`, "Accept": "application/json", "Content-Type": "application/json",
            "OData-MaxVersion": "4.0", "OData-Version": "4.0", "Prefer": "odata.include-annotations=\"*\""
        };

        const workerTable = "new_sagyouin_mastas"; 
        const workerQuery = `?$select=_owningbusinessunit_value,new_sagyouin_mastaid,new_mei,new_mail&$filter=new_mail eq '${userEmail}'`;
        const workerRes = await fetch(`${dataverseUrl}/api/data/v9.2/${workerTable}${workerQuery}`, { method: "GET", headers });
        if (!workerRes.ok) throw new Error(`Dataverse Search Error: ${workerRes.status}`);
        const workerData = await workerRes.json();

        if (!workerData.value || workerData.value.length === 0) {
            context.res = { status: 403, body: { error: "NoRegistration", details: `メールアドレス [${userEmail}] は作業員マスタに登録されていません。`, detectedEmail: userEmail } };
            return;
        }

        const worker = workerData.value[0];
        const myBusinessUnit = worker._owningbusinessunit_value;
        const myWorkerId = worker.new_sagyouin_mastaid;
        const myName = worker.new_mei || "担当者";
        const myBusinessUnitName = worker["_owningbusinessunit_value@OData.Community.Display.V1.FormattedValue"] || "";

        const dispatchTable = "new_table2s"; 
        const selectCols = "new_table2id,new_start_time,new_kashikiri,statuscode,new_sharyou,new_tokuisaki_mei,new_genbamei,new_sagyou_naiyou,new_renraku_jikou";
        let filter = `_owningbusinessunit_value eq ${myBusinessUnit} and _new_operator_value eq ${myWorkerId}`; 
        const query = `?$select=${selectCols}&$filter=${filter}&$orderby=new_start_time asc`;
        const jobsRes = await fetch(`${dataverseUrl}/api/data/v9.2/${dispatchTable}${query}`, { method: "GET", headers });
        if (!jobsRes.ok) throw new Error(`Jobs Error: ${jobsRes.status}`);
        const data = await jobsRes.json();

        const results = data.value.map(item => {
            let timeStr = "--:--";
            if (item.new_start_time) timeStr = new Date(item.new_start_time).toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit', timeZone: 'Asia/Tokyo' });
            return {
                id: item.new_table2id, time: timeStr,
                type: item["new_kashikiri@OData.Community.Display.V1.FormattedValue"] || "-",
                car: "代車", client: item.new_tokuisaki_mei || "名称なし", location: item.new_genbamei || "",
                workContent: item.new_sagyou_naiyou || "", notes: item.new_renraku_jikou || "", contact: "",
                status: item["statuscode@OData.Community.Display.V1.FormattedValue"] || "未確認", statusCode: item.statuscode
            };
        });

        context.res = { status: 200, body: { message: "Success", userName: myName, businessUnitName: myBusinessUnitName, count: results.length, results: results } };
    } catch (error) {
        context.log.error(error); context.res = { status: 500, body: { error: "SystemError", details: error.message } };
    }
};
