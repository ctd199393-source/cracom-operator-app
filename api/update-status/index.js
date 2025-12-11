const fetch = require("node-fetch");

module.exports = async function (context, req) {
    context.log("Status Update Triggered");

    try {
        // 1. 環境変数の取得 (Power AutomateのURL)
        const flowUrl = process.env.FLOW_URL_COMPLETE;
        if (!flowUrl) throw new Error("環境変数 FLOW_URL_COMPLETE が設定されていません。");

        // 2. リクエストデータの受け取り
        // (ID, 緯度, 経度 を受け取る。位置情報がない場合は null が来る前提)
        const { haishaId, lat, long } = req.body;

        if (!haishaId) throw new Error("配車IDが指定されていません。");

        // 3. ユーザー情報の取得 (監査用ログ出力)
        const header = req.headers["x-ms-client-principal"];
        const userDetails = header ? JSON.parse(Buffer.from(header, "base64").toString("ascii")).userDetails : "DevUser";

        context.log(`User: ${userDetails} updated Haisha: ${haishaId} (Lat: ${lat}, Long: ${long})`);

        // 4. Power Automate へ転送 (Fire and Forget)
        // ※Dataverseの更新はPower Automate側で行うため、ここでは投げるだけ
        const flowRes = await fetch(flowUrl, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                haishaId: haishaId,
                lat: lat || 0, // 数値型エラー回避のため、nullなら0を送る等の調整
                long: long || 0,
                // 100000004 = 作業完了 (Power Appsの定義に合わせる)
                status: 100000004 
            })
        });

        if (!flowRes.ok) {
            const errText = await flowRes.text();
            throw new Error(`Flow Error: ${errText}`);
        }

        context.res = {
            status: 200,
            body: { message: "作業完了を通知しました" }
        };

    } catch (error) {
        context.log.error(error);
        // エラーでもフロントエンドを止めないよう、メッセージを返す（ステータスは500にするが）
        context.res = {
            status: 500,
            body: { error: error.message }
        };
    }
};
