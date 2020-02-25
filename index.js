require('dotenv').config()
require("datejs")
const express = require('express')
const bodyParser = require("body-parser");

let app = express();

let apiRoutes = require("./api-routes");
app.use(bodyParser.urlencoded({ 
    extended: true 
}));
app.use(bodyParser.json());

app.use(function (req, res, next) {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Methods", "GET,HEAD,OPTIONS,POST,PUT");
  res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, contentType,Content-Type, Accept, Authorization");
  next();
});

app.use(process.env.APIROOT, apiRoutes);

var port = process.env.APIPORT
var server = app.listen(port, function () {
    console.log("App now running on port", port);
});