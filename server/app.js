require("express-async-errors")
const express = require("express")
const app = express()
require("dotenv").config()
require("./src/db/dbconnection")
const port = process.env.PORT || 5001//reached to port inside of the env file, dotenv tanımlamasıyla oldu bu, dotenv yapmadığımızda 5001 e gidiyodu
const router = require("./src/routers")
const errorHandlerMiddleware = require("./src/middlewares/errorHandler")
const cors = require("cors")
const corsOptions = require("./src/helpers/corsOptions")
const mongoSanitize = require('express-mongo-sanitize');
const path = require("path")
const apiLimiter = require("./src/middlewares/rateLimit")
const moment = require("moment-timezone")
moment.tz.setDefault("Europe/Istanbul")

//Middlewares
app.use(express.json())
app.use(express.json({limit:"50mb"}))
app.use(express.urlencoded({limit:"50mb", extended: true, parameterLimit:50000}))

// app.use(express.static(path.join(__dirname, "public")))
// app.use("/uploads", express.static(__dirname))
app.use(express.static(path.join(__dirname, "public")));
app.use("/uploads", express.static(path.join(__dirname, "public", "uploads")));

app.use(cors(corsOptions))

app.use("/api", apiLimiter)

//rotalardan önceye koyuyoruz apimize erişmeden önceye
app.use(
    mongoSanitize({
      replaceWith: '_',
    }),
  );

app.use("/api",router)
 

app.get("/", (req,res)=>{
    res.json({
        message:"Hoş geldiniz!"
    })
})

//error catching
app.use(errorHandlerMiddleware)

app.listen(port, ()=>{
    console.log(`Server ${port} portundan çalışıyor..`);
})
