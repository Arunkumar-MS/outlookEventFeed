// grab the packages we need
var express = require('express');
var app = express();
var port = process.env.PORT || 8080;
var bodyParser = require('body-parser');
var http = require('http').Server(app);
var io = require('socket.io')(http);
var cors = require('cors');
var mongoose = require('mongoose');
var axios = require('axios');
var microsoftGraph = require("@microsoft/microsoft-graph-client");


app.use(cors());
app.use(bodyParser.json()); // support json encoded bodies
app.use(bodyParser.urlencoded({
	extended: true
})); // support encoded bodies

const Schema = mongoose.Schema;
const USER = mongoose.model('user', new Schema({
	id: String,
	token: String,
}));

const COMPOSE_URI_DEFAULT = 'mongodb://arunkumar:arun123@ds233323.mlab.com:33323/user';
const connectionOptions = {
	server: {
		socketOptions: {
			socketTimeoutMS: 0,
			connectionTimeout: 0
		}
	}
};

mongoose.connect(COMPOSE_URI_DEFAULT, connectionOptions, function (error) {
	if (error) console.error(error)
	else console.log('mongo connected')
})
const conn = mongoose.connection;
conn.on('error', console.error.bind(console, 'connection error:'));


io.on('connection', function (client) {
	client.on('join', function (data) {
		console.log(data);
	});
	client.on('messages', function (data) {
		client.emit('broad', data);
		client.broadcast.emit('broad', data);
	});
});


app.post('/V1/feed', function (req, res) {
	if (req.query && req.query.validationToken) {
		res.send(req.query.validationToken);
		// Send a status of 'Ok'
		status = 200;
	} else {
		//  websocket code
		const userId = (req.body.value[0].resourceData['@odata.id'] || '').split('/')[1];
		USER.findOne({
			id: userId
		}, (err, userInfo) => {
			console.log('userInfo', userInfo);
			const URL = 'https://graph.microsoft.com/v1.0/' + req.body.value[0].resourceData['@odata.id'];
			console.log(URL);

			axios.get(URL, {
					headers: {
						Authorization: `Bearer ${userInfo.token}`
					}
				}).then(response => {
					console.log(response.data);
					io.emit('outlookData', response.data);
				})
				.catch((error) => {
					console.log('error ' + error);
				});
		});
	}
	res.send({});
});

app.get('/', (req, res) => {
	res.sendFile(__dirname + `/index.html`);
});

app.get('/V1/getOutlookfeed', (req, res) => {
	if (!req.query.token) {
		return res.send('please send valid token');
	}
	return getGraphClient(req.query.token)
		.api('/me')
		.get((err, data) => {
			USER.findOne({
				id: data.id
			}, (err, userData) => {
				if (userData) {
					userData.token = req.query.token;
					userData.save();
				} else {
					let user = new USER({
						id: data.id,
						token: req.query.token,
					});
					user.save();
				}
			});

			getGraphClient(req.query.token)
				.api('/me/events?$select=subject,body,bodyPreview,organizer,attendees,start,end,location,responseStatus')
				.get((err, events) => {
					events.userId = data.id;
					res.send(events);
				});
		});
});

app.post('/V1/accept', (req, res) => {
	if (!req.body.token) {
		return res.send('please send valid token');
	}
	const URL = `https://graph.microsoft.com/v1.0/me/events/${req.body.id}/accept`;
	return axios.post(URL, {}, {headers: {
		Authorization: `Bearer ${req.body.token}`
	}})
	.then((response) => {
			console.log(response);
			res.send(response);
		})
		.catch((err) => {
			console.log(err);
			res.send(err);
		});
})

app.post('/V1/decline', (req, res) => {
	if (!req.body.token) {
		return res.send('please send valid token');
	}
	return getGraphClient(req.body.token)
		.api(`/me/events/${req.body.id}/decline`)
		.version('beta')
		.post({}, (err, response) => {
			res.send(response);
		});
})

// start the server
http.listen(port);
console.log('Server started! At http://localhost:' + port);

function getGraphClient(accessToken) {
	return microsoftGraph.Client.init({
		defaultVersion: 'v1.0',
		debugLogging: true,
		authProvider: (done) => {
			done(null, accessToken);
		}
	});
}
