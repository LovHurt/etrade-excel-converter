const router = require("express").Router()
const multer = require("multer")
const upload = require("../middlewares/lib/upload")
const APIError = require("../utils/errors")
const Response = require("../utils/response")

const auth = require("../app/auth/router")
const user = require("../app/users/router")
const excel = require("../app/process-excel/router")

router.use(auth)
router.use(user)
router.use(excel)


router.post("/upload", function(req,res){
    upload(req,res,function(err){
        if(err instanceof multer.MulterError)
            throw new APIError("resim yüklenirken multer kaynaklı hata oluştu : ", err)
        else if(err)
            throw new APIError("resim yüklenirken hata oluştu : ", err)
        else return new Response(req.savedImages, "Yükleme Başarılı").success(res)
    })
})

module.exports = router