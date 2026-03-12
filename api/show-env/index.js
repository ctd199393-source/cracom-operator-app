module.exports = async function (context, req) {
    context.res = {
        status: 200,
        headers: {
            "Content-Type": "application/json"
        },
        body: {
            "status": "success",
            "message": "API is ALIVE!"
        }
    };
};
