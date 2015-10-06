var express = require('express');
var router = express.Router();


/* Piggyback other routes here */
router.use('/upload', require('./upload'));

/* GET home page. */
router.get('/', function(req, res, next) {
  res.render('index', { title: 'xcel2sqlite' });
});


module.exports = router;
