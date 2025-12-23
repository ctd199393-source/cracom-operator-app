const fetch = require("node-fetch");

module.exports = async function (context, req) {
    context.log("Status Update Triggered");

    try {
        const flowUrl = process.env.FLOW_URL_COMPLETE;
        if (!flowUrl) throw new Error("Server Error: FLOW_URL_COMPLETE is not set.");

        // mode: 'complete', 'undo', 'confirm' (★追加)
        const { haishaId, lat, long, mode } = req.body;
        
        if (!haishaId) throw new Error("ID not provided.");

        let targetStatus;
        let targetLat = lat;
        let targetLong = long;
        let targetTime;

        // ★修正: undo または confirm の場合は「確認済み」に戻す
        if (mode === 'undo' || mode === 'confirm') {
            // --- 取消/確認モード ---
            targetStatus = 100000003; // 確認済み
            
            // 位置情報と完了時間をクリアする
            targetLat = null;
            targetLong = null;
            targetTime = null; 
        } else {
            // --- 完了モード ---
            targetStatus = 100000004; // 作業完了
            
            if (targetLat === undefined) targetLat = null;
            if (targetLong === undefined) targetLong = null;
            
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
                completionTime: targetTime
            })
        });

        if (!flowRes.ok) {
            throw new Error(`Flow Error: ${await flowRes.text()}`);
        }

        // メッセージの出し分け
        let msg = "作業完了を通知しました";
        if (mode === 'undo') msg = "作業完了を取り消しました";
        if (mode === 'confirm') msg = "確認済みに更新しました";

        context.res = { status: 200, body: { message: msg } };

    } catch (error) {
        context.log.error(error);
        context.res = { status: 500, body: { error: error.message } };
    }
};
