const APIError = require("../utils/errors")

const errorHandlerMiddleware = (err, req, res, next) => {
    if(err instanceof APIError){
        return res.status(err.statusCode || 400)
            .json({
                success: false,
                message: err.message,
                err:err,
            })
    }

    console.log(err.name);

    // if(err.name ==="CastError") console.log("CastError");

    return res.status(500).json({
        success: false,
        message: "You encountered with an error, please check your api!",
        err
    })
}

module.exports = errorHandlerMiddleware