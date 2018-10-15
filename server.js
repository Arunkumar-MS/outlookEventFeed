// grab the packages we need
var express = require('express');
var app = express();
var port = process.env.PORT || 8080;
var bodyParser = require('body-parser');
app.use(bodyParser.json()); // support json encoded bodies
app.use(bodyParser.urlencoded({
	extended: true
})); // support encoded bodies
app.post('/V1/feed', function (req, res) {
	if (req.query && req.query.validationToken) {
		res.send(req.query.validationToken);
		// Send a status of 'Ok'
		status = 200;
	} else {
		console.log(JSON.stringify(req.body.value[0]));
		console.log('----\n', req.body.value[0].resourceData);
		//  websocket code
		res.send('post data');
	}
});

// start the server
app.listen(port);
console.log('Server started! At http://localhost:' + port);