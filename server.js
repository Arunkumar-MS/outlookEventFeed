// grab the packages we need
var express = require('express');
var app = express();
var port = process.env.PORT || 8080;
var bodyParser = require('body-parser');
var http = require('http').Server(app);
var io = require('socket.io')(http);
var cors = require('cors');

app.use(cors());
app.use(bodyParser.json()); // support json encoded bodies
app.use(bodyParser.urlencoded({
	extended: true
})); // support encoded bodies

io.on('connection', function(client) {  
    client.on('join', function(data) {
		console.log('----------------------------------------------------------------------------------------------------------------------------------');
        console.log(data);
    });
    client.on('messages', function(data) {
           client.emit('broad', data);
           client.broadcast.emit('broad',data);
    });
});


app.post('/V1/feed', function (req, res) {
	if (req.query && req.query.validationToken) {
		res.send(req.query.validationToken);
		// Send a status of 'Ok'
		status = 200;
	} else {
		//  websocket code
		console.log(req.body.value[0].resourceData);
		io.emit('outlookData', req.body.value[0].resourceData);
	}
});

app.get('/', (req, res) => {
	res.sendFile(__dirname+`/index.html`);
});



// start the server
http.listen(port);
console.log('Server started! At http://localhost:' + port);