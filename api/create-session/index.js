const jwt = require('jsonwebtoken');
const cookie = require('cookie');

module.exports = async function (context, req) {
    try {
        // 1. ASWAが付与する「ユーザーのログイン情報」を受け取る
        const header = req.headers['x-ms-client-principal'];
        
        if (!header) {
            context.res = { status: 401, body: "ログイン情報が見つかりません" };
            return;
        }

        // 2. Base64形式の情報をデコードしてJSONとして読み込む
        const encoded = Buffer.from(header, 'base64');
        const decoded = encoded.toString('ascii');
        const clientPrincipal = JSON.parse(decoded);

        // 3. Azureの環境変数から「秘密の鍵」を取り出す
        const secretKey = process.env.JWT_SECRET_KEY;
        if (!secretKey) {
            context.res = { status: 500, body: "サーバーの秘密鍵が設定されていません" };
            return;
        }

        // 4. 【変更点】7日間有効な「独自トークン（JWT）」を発行する
        const token = jwt.sign(
            { principal: clientPrincipal },
            secretKey,
            { expiresIn: '7d' } // ★ '48h' から '7d' に変更
        );

        // 5. 【変更点】発行したトークンを、ブラウザに保存させるためのCookieに包む
        const cookieString = cookie.serialize('cracom_session', token, {
            httpOnly: true,       // JavaScriptからのアクセスを禁止（安全）
            secure: true,         // HTTPS通信時のみ送信
            sameSite: 'strict',   // クロスサイトリクエストをブロック
            maxAge: 7 * 24 * 60 * 60, // ★ 7日間（秒数指定）に変更
            path: '/'             // サイト全体で有効
        });

        // 6. クッキーをブラウザに返却する
        context.res = {
            status: 200,
            headers: {
                'Set-Cookie': cookieString,
                'Content-Type': 'application/json'
            },
            body: { message: "7日間有効なセッションを発行しました" }
        };

    } catch (error) {
        context.log.error("Session creation error:", error);
        context.res = { status: 500, body: "サーバーエラーが発生しました" };
    }
};
