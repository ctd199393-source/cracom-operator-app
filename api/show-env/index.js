const { ClientSecretCredential } = require("@azure/identity");
const fetch = require("node-fetch");

module.exports = async function (context, req) {
    context.log("API Triggered: Final Version");

    try {
        // --------------------------------------------------
        // 1. 環境変数のチェック
        // --------------------------------------------------
        const tenantId = process.env.TENANT_ID;
        const clientId = process.env.CLIENT_ID;
        const clientSecret = process.env.CLIENT_SECRET;
        const dataverseUrl = process.env.DATAVERSE_URL;

        if (!tenantId || !clientId || !clientSecret || !dataverseUrl) {
            throw new Error("環境変数が不足しています。SWAの設定を確認してください。");
        }

        // --------------------------------------------------
        // 2. ユーザー情報の取得と正規化
        // --------------------------------------------------
        const header = req.headers["x-ms-client-principal"];
        let rawEmail = "unknown";
        let searchEmail = "";

        if (header) {
            const decoded = JSON.parse(Buffer.from(header, "base64").toString("ascii"));
            rawEmail = decoded.userDetails || "unknown";
        }

        // #EXT# 除去処理（ゲストユーザー対策）
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

        context.log(`Searching Dataverse for User: ${searchEmail}`);

        // --------------------------------------------------
        // 3. Dataverse 接続 (認証トークン取得)
        // --------------------------------------------------
        const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        const tokenResponse = await credential.getToken(`${dataverseUrl}/.default`);
        const accessToken = tokenResponse.token;

        // --------------------------------------------------
        // 4. 作業員マスタ検索 (new_sagyouin_masta)
        // --------------------------------------------------
        // 検索キー: new_mail (メールアドレス)
        // 取得項目: new_sagyouin_id (作業員名), owningbusinessunit (所属部署)
        const filterUser = `new_mail eq '${searchEmail}'`; 
        const queryUser = `${dataverseUrl}/api/data/v9.2/new_sagyouin_mastas?$filter=${encodeURIComponent(filterUser)}&$select=new_sagyouin_id,_owningbusinessunit_value`;

        const userRes = await fetch(queryUser, {
            method: "GET",
            headers: {
                "Authorization": `Bearer ${accessToken}`,
                "Accept": "application/json",
                "OData-MaxVersion": "4.0",
                "OData-Version": "4.0"
            }
        });

        if (!userRes.ok) {
            const errText = await userRes.text();
            throw new Error(`Master Search Error (${userRes.status}): ${errText}`);
        }

        const userData = await userRes.json();

        // 判定: マスタにいない場合
        if (userData.value.length === 0) {
            context.res = { 
                status: 403, 
                body: { error: `あなたのメールアドレス (${searchEmail}) は作業員マスタに登録されていません。` } 
            };
            return;
        }

        const userRecord = userData.value[0];
        const businessUnitId = userRecord._owningbusinessunit_value; // 所属部署ID

        // --------------------------------------------------
        // 5. 配車データ検索 (new_Table2)
        // --------------------------------------------------
        // フィルタ: 所属部署 (owningbusinessunit) が一致するもの
        // 取得項目: 日付, 現場名, 作業内容, 開始時間
        const filterDispatch = `_owningbusinessunit_value eq '${businessUnitId}'`;
        
        // ※テーブル名の複数形は通常「s」がつきます。new_table2s と仮定します。
        // ※列名は提供いただいた定義に基づき new_day, new_genbamei, new_sagyou_naiyou, new_start_time を取得します。
        const queryDispatch = `${dataverseUrl}/api/data/v9.2/new_table2s?$filter=${encodeURIComponent(filterDispatch)}&$select=new_day,new_genbamei,new_sagyou_naiyou,new_start_time&$orderby=new_day desc`;
        
        const dispatchRes = await fetch(queryDispatch, {
            headers: { 
                "Authorization": `Bearer ${accessToken}`,
                "Accept": "application/json",
                "OData-MaxVersion": "4.0",
                "OData-Version": "4.0"
            }
        });

        if (!dispatchRes.ok) {
             const dispErr = await dispatchRes.text();
             throw new Error(`Dispatch Search Error (${dispatchRes.status}): ${dispErr}`);
        }
        
        const dispatchData = await dispatchRes.json();

        // --------------------------------------------------
        // 6. 結果返却
        // --------------------------------------------------
        context.res = {
            status: 200,
            body: {
                user: {
                    name: userRecord.new_sagyouin_id, // 作業員名
                    email: searchEmail
                },
                records: dispatchData.value
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
