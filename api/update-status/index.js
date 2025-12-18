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

        // モードによる値の切り替え
        if (mode === 'undo') {
            // --- 取消モード ---
            // ステータスを「現場確定(100000001)」に戻す
            targetStatus = 100000001;
            // 位置情報はクリアする (null)
            targetLat = null;
            targetLong = null;
        } else {
            // --- 完了モード (デフォルト) ---
            // ステータスを「作業完了(100000004)」にする
            targetStatus = 100000004;
            // 位置情報は送られてきた値 (なければ0またはnull)
            if (targetLat === undefined) targetLat = null;
            if (targetLong === undefined) targetLong = null;
        }

        // Power Automate へ送信
        const flowRes = await fetch(flowUrl, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                haishaId: haishaId,
                lat: targetLat,
                long: targetLong,
                status: targetStatus
            })
        });

        if (!flowRes.ok) {
            throw new Error(`Flow Error: ${await flowRes.text()}`);
        }

        // メッセージの出し分け
        const msg = (mode === 'undo') ? "作業完了を取り消しました" : "作業完了を通知しました";

        context.res = { status: 200, body: { message: msg } };

    } catch (error) {
        context.log.error(error);
        context.res = { status: 500, body: { error: error.message } };
    }
};
