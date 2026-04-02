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

        // 2. Base64という形式で暗号化されている情報を解読する
        const encoded = Buffer.from(header, 'base64');
        const decoded = encoded.toString('ascii');
        const clientPrincipal = JSON.parse(decoded);

        // 3. 先ほどAzureに設定した「秘密の鍵」を取り出す
        const secretKey = process.env.JWT_SECRET_KEY;
        if (!secretKey) {
            context.res = { status: 500, body: "サーバーの秘密鍵が設定されていません" };
            return;
        }

        // 4. 48時間有効な「独自トークン（JWT）」を発行する
        const token = jwt.sign(
            { principal: clientPrincipal },
            secretKey,
            { expiresIn: '48h' }
        );

        // 5. 発行したトークンを、ブラウザに保存させるためのCookie（クッキー）に包む
        const cookieString = cookie.serialize('cracom_session', token, {
            httpOnly: true,       // JavaScriptからの盗み見を防止
            secure: true,         // HTTPS通信のときだけ送信
            sameSite: 'strict',   // 他のサイトからの干渉をブロック
            maxAge: 48 * 60 * 60, // 48時間（秒で指定）
            path: '/'             // アプリ全体で有効
        });

        // 6. 出来上がったCookieをブラウザにお持ち帰りさせる
        context.res = {
            status: 200,
            headers: {
                'Set-Cookie': cookieString,
                'Content-Type': 'application/json'
            },
            body: { message: "48時間有効なセッションを発行しました" }
        };

    } catch (error) {
        context.log.error("Session creation error:", error);
        context.res = { status: 500, body: "サーバーエラーが発生しました" };
    }
};
