// ★ 新規追加：セッション確認用のライブラリ
const jwt = require('jsonwebtoken');
const cookie = require('cookie');

module.exports = async function (context, req) {
    context.log("Status Update Triggered");

    // --- ▼ここから追加：48時間入館証のチェック（見張り番）▼ ---
    let decodedToken = null;
    try {
        const cookieHeader = req.headers.cookie;
        if (!cookieHeader) throw new Error("Cookieが存在しません");
        
        const cookies = cookie.parse(cookieHeader);
        const token = cookies.cracom_session;
        if (!token) throw new Error("トークンが存在しません");

        const secretKey = process.env.JWT_SECRET_KEY;
        if (!secretKey) throw new Error("サーバーの秘密鍵が設定されていません");

        // トークンが本物か、有効期限（48時間）内かをチェック
        decodedToken = jwt.verify(token, secretKey);
    } catch (error) {
        // チェックNG（偽造、または48時間経過）の場合はここで弾く
        context.log.warn("Session Check Failed:", error.message);
        context.res = { status: 401, body: { error: "セッションの有効期限が切れました。再ログインしてください。" } };
        return; // ← これ以上下の処理を実行させない
    }
    // --- ▲追加ここまで▲ ---

    try {
        const flowUrl = process.env.FLOW_URL_COMPLETE;
        if (!flowUrl) throw new Error("Server Error: FLOW_URL_COMPLETE is not set.");

        // mode: 'complete', 'undo', 'confirm'
        const { haishaId, lat, long, mode } = req.body;
        
        if (!haishaId) throw new Error("ID not provided.");

        let targetStatus;
        let targetLat = lat;
        let targetLong = long;
        let targetTime;

        if (mode === 'undo' || mode === 'confirm') {
            // --- 取消/確認モード ---
            targetStatus = 100000003; // 確認済み
            targetLat = null;
            targetLong = null;
            targetTime = null; 
        } else {
            // --- 完了モード ---
            targetStatus = 100000004; // 作業完了
            
            if (targetLat === undefined) targetLat = null;
            if (targetLong === undefined) targetLong = null;
            
            // ★修正: 日本時間 (JST) に変換してセット
            const now = new Date();
            now.setHours(now.getHours() + 9); // 9時間進める
            // 'Z' を削除して「現地時間」として扱うようにする
            targetTime = now.toISOString().replace('Z', '');
        }

        // Power Automate へ送信 (標準内蔵の fetch を使用)
        const flowRes = await fetch(flowUrl, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                haishaId: haishaId,
                lat: targetLat,
                long: targetLong,
                status: targetStatus,
                completionTime: targetTime
            })
        });

        if (!flowRes.ok) {
            throw new Error(`Flow Error: ${await flowRes.text()}`);
        }

        let msg = "作業完了を通知しました";
        if (mode === 'undo') msg = "作業完了を取り消しました";
        if (mode === 'confirm') msg = "確認済みに更新しました";

        context.res = { status: 200, body: { message: msg } };

    } catch (error) {
        context.log.error(error);
        context.res = { status: 500, body: { error: error.message } };
    }
};
