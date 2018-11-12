var path = require('path'); 
var fs = require('fs');
var express = require('express');
var app = express();

// Initialize variables. 
var port = process.env.PORT || 8080;  

// Set the front-end folder to serve public assets.
app.use(express.static(__dirname));

// Set the route to the index.html file.
app.get('/', function (req, res) {
    var homepage = path.join(__dirname, 'dist/index.html');
    res.sendFile(homepage);
});

// Start the app.  
app.listen(port);
console.log('Listening on port ' + port + '...');
