const router = require('express').Router();
const VerifyToken = require('./VerifyToken');

router.get('/', function(req, res) {
    res.send(`Hello! The API is at ${process.env.HOST}${process.env.APIROOT}`);
});

var controller = require('./controller');

router.get('/calendarevents/:emailaddress', VerifyToken, controller.getCalendarEvents)
    
module.exports = router;