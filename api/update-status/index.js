const fetch = require("node-fetch");

module.exports = async function (context, req) {
    context.log("Status Update Triggered");

    try {
        const flowUrl = process.env.FLOW_URL_COMPLETE;
        if (!flowUrl) throw new Error("Server Error: FLOW_URL_COMPLETE is not set.");

        // mode: 'complete' (完了) or 'undo' (取消)
        const { haishaId, lat, long, mode } = req.body;
        
        if (!haishaId) throw new Error("ID not provided.");

        let targetStatus;
        let targetLat = lat;
        let targetLong = long;
        let targetTime; // ★追加: 完了時間用変数

        // モードによる値の切り替え
        if (mode === 'undo') {
            // --- 取消モード ---
            // ★変更: ステータスを「確認済み(100000003)」に戻す
            targetStatus = 100000003;
            
            // 位置情報と完了時間をクリアする (null)
            targetLat = null;
            targetLong = null;
            targetTime = null; 
        } else {
            // --- 完了モード ---
            // ステータスを「作業完了(100000004)」にする
            targetStatus = 100000004;
            
            // 位置情報 (なければnull)
            if (targetLat === undefined) targetLat = null;
            if (targetLong === undefined) targetLong = null;
            
            // ★追加: 現在時刻をセット (ISO形式)
            targetTime = new Date().toISOString();
        }

        // Power Automate へ送信
        const flowRes = await fetch(flowUrl, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                haishaId: haishaId,
                lat: targetLat,
                long: targetLong,
                status: targetStatus,
                completionTime: targetTime // ★追加
            })
        });

        if (!flowRes.ok) {
            throw new Error(`Flow Error: ${await flowRes.text()}`);
        }

        const msg = (mode === 'undo') ? "作業完了を取り消しました" : "作業完了を通知しました";
        context.res = { status: 200, body: { message: msg } };

    } catch (error) {
        context.log.error(error);
        context.res = { status: 500, body: { error: error.message } };
    }
};
