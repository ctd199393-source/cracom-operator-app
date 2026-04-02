const cookie = require('cookie');

module.exports = async function (context, req) {
    // 過去の日付（有効期限切れ）を指定した空のCookieを発行して、ブラウザ上の入館証を強制的に上書き消去する
    const clearCookie = cookie.serialize('cracom_session', '', {
        httpOnly: true,
        secure: true,
        sameSite: 'strict',
        maxAge: 0, // 寿命ゼロ（即時削除）
        path: '/'
    });

    context.res = {
        status: 200,
        headers: {
            'Set-Cookie': clearCookie,
            'Content-Type': 'application/json'
        },
        body: { message: "セッションを破棄しました" }
    };
};
