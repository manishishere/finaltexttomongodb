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
var cnt=0;
// fs.readdir(__dirname+'/data', function(err, files) {
//     if (err) return err;
//     files.forEach(function(f) {
//             cnt++;

//     });
//    // console.log(cnt);
// }); 

//function puts(error, stdout, stderr) {sys.log(stdout)}
//exec("split -l 100000 -a 1 -d data.txt data/x", puts)

setTimeout(function(){  
mongoose.connect('mongodb://127.0.0.1:27017/manish',function(err,res)
  {
  if(err)console.log(err);
  });
var userSchema = new Schema({
  DOI:String,
  Date_Ingested:String,
  Portico_Article_ID:String,
  Portico_Article_bibCitation:String,
  Triggered:String   
 });
var b=[];
var app = require('express')();
app.listen(3000);
var bodyParser = require('body-parser');
app.use(bodyParser.json()); 
app.use(bodyParser.urlencoded({ extended: true }));

var Sam = mongoose.model('sam', userSchema);
module.exports = Sam;
 var starttime=timestamp();
;

//fs.createReadStream('c.csv',{start: start ,end:end})
// var stats=fs.statSync("data/data.txt");
//   var ts=stats["size"]
// var end=ts/1000;
// var start=0;
 var checkcnt=0;

//tfs.createReadStream('data.txt',{start: start ,end :end}).pipe(fs.createWriteStream('dt1.txt'));
  



        function process1(file)
        {

                var LineByLineReader = require('line-by-line'),
                lr = new LineByLineReader(file); 
                lr.on('error', function (err) {
                    // 'err' contains error object
                  });
                lr.on('line', function (line) {
                  // pause emitting of lines...
                  lr.pause();

                  // ...do your asynchronous line processing..
                  setTimeout(function () {
                            
                       
                            temp=line.split("\t");
                            var obj={DOI:temp[0],Date_Ingested:temp[1],Portico_Article_ID:temp[2],Portico_Article_bibCitation:temp[3],Triggered:temp[4]};
                           // console.log(obj);
                            var s = new Sam(obj);
                            Sam.update({DOI: temp[0]}, {$set:{DOI:temp[0],
                              Date_Ingested:temp[1],
                              Portico_Article_ID:temp[2],
                              Portico_Article_bibCitation:temp[3],
                              Triggered:temp[4] } },
                             {upsert: true}, function(err, doc){
                              if(err){
                                  console.log("Something wrong when updating data!");
                              }

                              console.log(doc);
                              });
                                                        
                                    //   s.save(function(err,result){
                                    //     if(err)console.log(err);
                                    //     if(result){
                                    //   // console.log(result);
                                    //     checkcnt++;
                                    //   //    console.log(checkcnt);
                                    //       }
                                    // });
                                       
                          
         
                        
                             // ...and continue emitting lines.
                    lr.resume();

                  }, 0);  
                });
          lr.on('end', function () {
              // All lines are read, file is closed now.
              console.log("**************done reading page***********");
            });
      }

// var asyncTasks = [];
//       console.log("-------------------")
//       console.log(cnt);
//       console.log("-------------------")
     

//      for(var j=0;j<cnt;j++)
//      {
     
//        // var st="data/x"+i;
//        i=0;       
//         asyncTasks.push(function(){ 
//          var st="data/x"+i;
         process1("data.txt");
//         console.log(st);
//         i=i+1; });

      
//     }

   
// console.log("content of array:");
// console.log(asyncTasks.toString());
//async.parallel(asyncTasks, function(){
// console.log("end");
//});



// {
//   async.parallel([
//   function(callback) {
//       var st="data/x"+i;
//      process1(st);
//      console.log(st);
    
//   },
//   // function(callback) {
//   //   setTimeout(function() {
//   //     console.log('--------------------------------------------Task 2');
//   //     callback(null, 2);
//   //   },0);
//   // },
//   // function(callback) {
//   //   setTimeout(function() {
//   //     console.log('----------------------------------------------Task 3');
//   //     callback(null, 3);
//   //   }, 0);
//   // }
//   ], function(error, results) {
//    // cb(error, results)
//    if(error)console.log(error);
//   });
// }

app.get('/get', function (req, res) {
  
    Sam.find({},function(err,result)
    {
       res.send(result);

    });

});

app.post('/post', function (req, res) {   });

app.put('/put', function (req, res) {  });

         /*User.update({"name":name},{$set: {"name":name,"mobile":mobile} },function(err, data){   
          if (err) return res.send(err);
          res.send(data);
          }); */

}, 10);