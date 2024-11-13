const multer = require("multer")
const path = require("path")
const fs = require("fs")

const fileFilter = (req, file, cb) => {
    const allowedMimeTypes = ["image/jpg", "image/gif", "image/jpeg", "image/png"]

    if(!allowedMimeTypes.includes(file.mimetype)){
        cb(new Error("Bu resim tipi desteklenmemektedir, lütfen farklı bir resim seçinizi"), false)
    }
    cb(null, true)
}

const storage = multer.diskStorage({
    destination: function(req, file, cb){
        const rootDir = path.dirname(require.main.filename)
        const uploadPath = path.join(rootDir, "public", "uploads");

        console.log("require.main.filename : ", require.main.filename);
        
        if (!fs.existsSync(uploadPath)) {
            fs.mkdirSync(uploadPath, { recursive: true });
        }
        
        cb(null, uploadPath)
    },
    filename: function(req, file, cb){
        const extension = file.mimetype.split("/")[1]

        if(!req.savedImages) req.savedImages = []

        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9)

        let url = `image_${uniqueSuffix}.${extension}`

        req.savedImages.push(url);

        cb(null, url)
    }
})

const upload = multer({storage, fileFilter}).array("images")

module.exports = upload