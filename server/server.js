const path = require('path'); 
const fs = require('fs');
const express = require('express');
const bodyParser = require('body-parser');
const localPath = path.join(__dirname, "../public");
const app = express();

// Initialize variables. 
const port = process.env.PORT || 8080;  

// Middlewares to Set the front-end folder to serve public assets and enable body Parser
app.use(express.static(localPath));
app.use(bodyParser.urlencoded({extended: true}));

// Set the route to the index.html file.
app.get('/', function (req, res) {
    var homepage = path.join(__dirname , 'index.html');
    res.sendFile(homepage);
});

// Route to New, Delete, Edit Profile
app.post('/', function(req, res){
    var jsonPath = path.join(localPath , 'json/profiles.json');
    var dataToWrite = JSON.stringify(req.body.profiles)
    fs.writeFileSync(jsonPath, dataToWrite)
    res.send('Success')
})

// Start the app.  
app.listen(port);
console.log('Listening on port ' + port + '...');
