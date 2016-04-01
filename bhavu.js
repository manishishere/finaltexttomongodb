var mongodb = require('mongodb');
var mongoose = require('mongoose');
var Schema = mongoose.Schema;
var fs = require("fs");
var tfs = require("fs");
var byline = require('byline');
var async = require('async');
var timestamp = require('log-timestamp');
var sys = require('util');
var exec = require('child_process').exec;
var cnt = 0;
var parseXlsx = require('excel');
mongoose.connect('mongodb://127.0.0.1:27017/manish', function(err, res) {
    if (err) console.log(err);
});
var userSchema = new Schema({
OrderDate:String,
Region:String,
Rep:String,
Item:String,
Units:String,
Unit_Cost:String,
Total:String
});
var b = [];
var app = require('express')();
app.listen(8080);
var bodyParser = require('body-parser');
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({
    extended: true
}));
var Sam = mongoose.model('samexl', userSchema);
module.exports = Sam;
var starttime = timestamp();
parseXlsx('Spreadsheet.xlsx', function(err, temp) {
    if (err) throw err;
    // data is an array of arrays
    for (var i = 0; i < temp.length; i++) {
       // for (var j = 0; j < temp[i].length; j++)
         {
            console.log(temp[i][0]);
            console.log("___________________");
            Sam.update({
                OrderDate: temp[i][0]
            }, {
                $set: {
                    OrderDate:temp[i][0],
                    Region:temp[i][1],
                    Rep:temp[i][2],
                    Item:temp[i][3],
                    Units:temp[i][4],
                    Unit_Cost:temp[i][5],
                    Total:temp[i][6],
                              
                }
            }, {
                upsert: true
            }, function(err, doc) {
                if (err) {
                    console.log(
                        "Something wrong when updating data!"
                    );
                }
                console.log(doc);
            });
        }
    }
    console.log("end");
});
app.get('/get', function(req, res) {
    Sam.find({}, function(err, result) {
        res.send(result);
    });
});
app.post('/post', function(req, res) {});
app.put('/put', function(req, res) {});
/*User.update({"name":name},{$set: {"name":name,"mobile":mobile} },function(err, data){   
          if (err) return res.send(err);
          res.send(data);
          }); */