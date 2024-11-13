const joi = require("joi")

const APIError = require("../../utils/errors")

class authValidation{
    constructor() {}

    static register = async (req, res, next) => {
        try {
            await joi.object({
                name: joi.string().trim().min(3).max(100).required().messages({
                    "string.base": "İsim Alanı Normal Metin Olmalıdır",
                    "string.empty": "İsim alanı boş olamaz",
                    "string.min": "isim alanı 3 karakterden az olamaz",
                    "string.max": "isim alanı en fazla 100 karakter olabilir",
                    "string.required": "isim alanı zorunludur"
                }),
                lastname: joi.string().trim().min(3).max(100).required().messages({
                    "string.base": "Soyad Alanı Normal Metin Olmalıdır",
                    "string.empty": "Soyad alanı boş olamaz",
                    "string.min": "Soyad alanı 3 karakterden az olamaz",
                    "string.max": "Soyad alanı en fazla 100 karakter olabilir",
                    "string.required": "Soyad alanı zorunludur"
                }),
                email: joi.string().email().trim().min(3).max(100).required().messages({
                    "string.base": "Email Alanı Normal Metin Olmalıdır",
                    "string.empty": "Email alanı boş olamaz",
                    "string.email": "Lütfen geçerli bir email giriniz!",
                    "string.min": "Email alanı 3 karakterden az olamaz",
                    "string.max": "Email alanı en fazla 100 karakter olabilir",
                    "string.required": "Email alanı zorunludur"
                }),
                password: joi.string().trim().min(6).max(36).required().messages({
                    "string.base": "Şifre Alanı Normal Metin Olmalıdır",
                    "string.empty": "Şifre alanı boş olamaz",
                    "string.min": "Şifre alanı 6 karakterden az olamaz",
                    "string.max": "Şifre alanı en fazla 36 karakter olabilir",
                    "string.required": "Şifre alanı zorunludur"
                })
            }).validateAsync(req.body)
        } catch(error){
            if(error.details && error?.details[0].message)
                throw new APIError(error.details[0].message, 400)
            else throw new APIError("Lütfen validasyon kurallarına uyun", 400)
        }
        next()
    } 
    static login = async (req,res,next)=> {
        try {
            await joi.object({
                email: joi.string().email().trim().min(3).max(100).required().messages({
                    "string.base": "Email Alanı Normal Metin Olmalıdır",
                    "string.empty": "Email alanı boş olamaz",
                    "string.email": "Lütfen geçerli bir email giriniz!",
                    "string.min": "Email alanı 3 karakterden az olamaz",
                    "string.max": "Email alanı en fazla 100 karakter olabilir",
                    "string.required": "Email alanı zorunludur"
                }),
                password: joi.string().trim().min(6).max(36).required().messages({
                    "string.base": "Şifre Alanı Normal Metin Olmalıdır",
                    "string.empty": "Şifre alanı boş olamaz",
                    "string.min": "Şifre alanı 6 karakterden az olamaz",
                    "string.max": "Şifre alanı en fazla 36 karakter olabilir",
                    "string.required": "Şifre alanı zorunludur"
                })
            }).validateAsync(req.body)
        } catch(error){
            if(error.details && error?.details[0].message)
                throw new APIError(error.details[0].message, 400)
            else throw new APIError("Lütfen validasyon kurallarına uyun", 400)
        }
        next();
    }
}

module.exports = authValidation