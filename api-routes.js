const router = require('express').Router();
const VerifyToken = require('./VerifyToken');

router.get('/', function(req, res) {
    res.send(`Hello! The API is at ${process.env.HOST}${process.env.APIROOT}`);
});

var controller = require('./controller');

router.get('/calendaritems/emailaddress/:emailaddress', VerifyToken, controller.getCalendarItems)

router.get('/calendaritems/:itemid', VerifyToken, controller.getItems)

router.get('/calendaravailability/:emailaddress', VerifyToken, controller.getCalendarAvailability)

router.post('/calendaritems', VerifyToken, controller.createCalendarItems)

router.delete('/calendaritems', VerifyToken, controller.cancelCalendarItems)
    
module.exports = router;