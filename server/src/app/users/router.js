const router = require("express").Router()
const { tokenCheck } = require("../../middlewares/auth")
const {me} = require("./controller")

router.get("/me", tokenCheck, me)

module.exports = router