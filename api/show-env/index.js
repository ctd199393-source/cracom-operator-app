const { ClientSecretCredential } = require("@azure/identity");
const fetch = require("node-fetch");

module.exports = async function (context, req) {
    context.log('Dataverse API request started.');

    try {
        // --------------------------------------------------
        // 1. 設定と準備
        // --------------------------------------------------
        const tenantId = process.env.TENANT_ID; 
        const clientId = process.env.CLIENT_ID;
        const clientSecret = process.env.CLIENT_SECRET;
        const dataverseUrl = process.env.DATAVERSE_URL;

        if (!tenantId || !clientId || !clientSecret || !dataverseUrl) {
            throw new Error("環境変数が不足しています");
        }

        // ログインユーザーのメールアドレス (テスト時は自分のメルアドに書き換えて確認可)
        const userEmail = req.headers["x-ms-client-principal-name"];
        
        if (!userEmail) {
            context.res = { status: 401, body: { error: "ログインが必要です" } };
            return;
        }

        // 認証トークン取得
        const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        const tokenResponse = await credential.getToken(`${dataverseUrl}/.default`);
        const accessToken = tokenResponse.token;
        
        const headers = {
            "Authorization": `Bearer ${accessToken}`,
            "Accept": "application/json",
            "Content-Type": "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            "Prefer": "odata.include-annotations=\"*\""
        };

        // --------------------------------------------------
        // 2. 作業員マスタ検索 (セキュリティと部署特定)
        // --------------------------------------------------
        // ※テーブル名や列名が変更になった場合はここを修正
        const workerTable = "new_sagyouin_mastas"; 
        const workerQuery = `?$select=_owningbusinessunit_value,new_sagyouin_mastaid,new_mei&$filter=new_mail eq '${userEmail}'`;
        
        const workerRes = await fetch(`${dataverseUrl}/api/data/v9.2/${workerTable}${workerQuery}`, { method: "GET", headers });
        const workerData = await workerRes.json();

        if (!workerData.value || workerData.value.length === 0) {
            context.res = { status: 403, body: { error: "作業員マスタに登録がありません" } };
            return;
        }

        const worker = workerData.value[0];
        const myBusinessUnit = worker._owningbusinessunit_value;
        const myWorkerId = worker.new_sagyouin_mastaid;
        const myName = worker.new_mei || "担当者";

        // --------------------------------------------------
        // 3. 配車データ取得 (配車テーブル + 案件情報)
        // --------------------------------------------------
        const dispatchTable = "new_table2s"; // 配車テーブル

        // ▼ 取得したい列 (API論理名)
        const selectCols = [
            "new_table2id",         // 配車ID
            "new_start_time",       // 開始日時
            "new_kashikiri",        // 貸切区分
            "statuscode",           // ステータス
            "new_sharyou",          // 車両 (Lookup)
            "_new_id_value"         // 案件ID (Lookup)
        ].join(",");

        // ▼ 案件情報の展開 ($expand)
        // ここで親テーブル(案件)の情報を引っ張ります
        const expandAnken = `new_id($select=new_tokuisakimei,new_genbamei,new_bikou,new_renraku_jikou)`;
        // もし車両名も展開して取得するなら追加 (例: new_sharyou($select=new_name))
        
        // ▼ フィルタリング (部署 && 担当者 && 日付)
        // 日付範囲：今日の0:00以降 (夜勤対応のため広めに取る場合は調整)
        const todayStr = new Date().toISOString().split('T')[0];
        
        let filter = `_owningbusinessunit_value eq ${myBusinessUnit}`;
        filter += ` and _new_operator_value eq ${myWorkerId}`; // 自分への配車のみ
        filter += ` and new_start_time ge ${todayStr}`;       // 今日以降

        const query = `?$select=${selectCols}&$expand=${expandAnken}&$filter=${filter}&$orderby=new_start_time asc`;
        
        const apiUrl = `${dataverseUrl}/api/data/v9.2/${dispatchTable}${query}`;
        const response = await fetch(apiUrl, { method: "GET", headers });
        
        if (!response.ok) {
            throw new Error(`Dataverse Error: ${response.status} ${await response.text()}`);
        }
        
        const data = await response.json();

        // --------------------------------------------------
        // 4. データ整形 (フロントエンド用のシンプルな名前に変換)
        // --------------------------------------------------
        const results = data.value.map(item => {
            const anken = item.new_id || {};
            
            // 日付フォーマット処理 (UTC -> JST)
            const dateObj = new Date(item.new_start_time);
            // 時間だけ抽出 (例: "8:00")
            const timeStr = dateObj.toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit', timeZone: 'Asia/Tokyo' });
            // AM/PM判定などが必要ならここでロジック追加

            return {
                id: item.new_table2id,
                
                // 表示用データマッピング
                time: timeStr,                                  // 時間
                type: item["new_kashikiri@OData.Community.Display.V1.FormattedValue"] || "-", // 貸切区分
                car: "代車 4958",                               // ※車両情報が取れたらここに入れる
                
                client: anken.new_tokuisakimei || "名称未設定",   // 得意先名
                location: anken.new_genbamei || "",             // 現場名
                workContent: anken.new_bikou || "",             // 作業内容
                notes: anken.new_renraku_jikou || "",           // 連絡事項
                contact: "連絡先未設定",                          // ※連絡先列があればここに入れる

                status: item["statuscode@OData.Community.Display.V1.FormattedValue"] || "未確認",
                statusCode: item.statuscode // 数値(ロジック判定用)
            };
        });

        context.res = {
            status: 200,
            body: { 
                message: "Success", 
                userName: myName, // 画面ヘッダー表示用
                results: results 
            }
        };

    } catch (error) {
        context.log.error(error);
        context.res = { status: 500, body: { error: error.message } };
    }
};
