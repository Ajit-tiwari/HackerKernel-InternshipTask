const express=require("express");
const bodyParser=require("body-parser");
const mongoose=require("mongoose");
const excel=require('exceljs');
const fs=require("fs");
const nodemailer=require('nodemailer');

const App=express();


mongoose.connect("mongodb://localhost:27017/Ajit",{ useNewUrlParser: true }); //Database Connection
var db=mongoose.connection;
db.on('error',console.log.bind(console,"connection error"));
db.once('open',function(callback){                      //opening connection
  console.log("connection succeded");
});

const dataBaseInsertion=(data,collection)=>{                //Helper Function to insert data into database name collection
  db.collection(collection).insertOne(data,function(err, collection){
    if (err) throw err;
    console.log("Record inserted Successfully");
  });
};

const exportExcel=()=>{           //helper function to fetch data from database and write it into Excel File
  return new Promise(function(resolve,reject){         
      
      try {                             //check wether the file already exist or not
        var fileName="customer.xlsx";           
      if(fs.existsSync(fileName)) {
          console.log("The file exists.");
          fs.unlink('./customer.xlsx',function(err){
            if(err) return console.log(err);
            console.log('file deleted successfully');
          });  
          } else {
          console.log('The file does not exist.');
        }
        var workbook = new excel.Workbook();                  //creating Excel file
        var worksheetUser = workbook.addWorksheet("users");   //adding users worksheet to excel file
        var worksheetTask = workbook.addWorksheet("table");   //adding table worksheet to excel file
        worksheetUser.state='visible';
        worksheetTask.state='visible'; 
        db.collection("users").find({}).toArray(function(err, result) {    //fetching users data from database
          if (err) throw err;
          worksheetUser.columns = [                       //creating headers for worksheet
            { header: 'Name', key: 'name', width: 30 },
            { header: 'Email', key: 'email', width: 30},
            { header: 'Mobile', key: 'mobile', width: 10, outlineLevel: 1}
          ];
          worksheetUser.addRows(result);              //adding data to worksheet in row fashion
        });
        db.collection("table").find({}).toArray(function(err,result){    //fetching table data from database
          if(err) throw err;
          worksheetTask.columns = [                       //creating headers for worksheet
              { header: 'Name', key: 'name', width: 30 },
              { header: 'Task', key: 'task', width: 30},
              { header: 'Status', key: 'status', width: 10, outlineLevel: 1}
            ];
            worksheetTask.addRows(result);
            workbook.xlsx.writeFile("customer.xlsx")        //writing users and table data to customer.xlsx file
              .then(function() {
              resolve("success");
              });
          });
            } catch (err) {
              reject(err);
            }
        });
};

const transporter = nodemailer.createTransport({    //used nodemailer to send confirmation mails to registred users
  service: 'gmail',
  auth: {
    user: 'tajit40.at@gmail.com',
    pass: '#############'
  }
});

App.use(express.urlencoded({
    extended: true
  }))
App.use(bodyParser.json());
App.use(bodyParser.urlencoded({
    extended: true
}));

App.set('view engine','ejs')
App.set('views','views')

App.get('/',(req,res,next)=>{
    res.render('user');
});

App.post("/done",(req,res)=>{
  var name=req.body.Name;
  var email=req.body.Email;
  var mobile=req.body.Phone;
  
  var data={
    "name" :name,
    "email":email,
    "mobile":mobile
  }
  console.log(data);
  dataBaseInsertion(data,"users");      //calling Insertion function to insert data into database

    const mailOptions = {               // creating mail from tajit40.at@gmail.com to the current registered user
      from: 'tajit40.at@gmail.com',
      to: email,
      subject: 'Registration Done',
      text: 'Congratulation your record is added !!'
    };
  
transporter.sendMail(mailOptions, function(error, info){    //sending mails
  if (error) {
    console.log(error);
  } else {
    console.log('Email sent: ' + info.response);
  }
});
    res.send("Record Inserted");
});

App.get('/addTable',(req,res)=>{          //task 2: the Table File
    db.collection("users").find({}).toArray(function(err, result) {   //fetching data from users database
    if (err) throw err;
    res.render('addTable',{detail:result});         //passing fetched data to frontend using ejs functionality
    });
})

App.post('/userAddedIntoTable',(req,res)=>{         // handles the route for insertion of data from table form (task-2)
    var name=req.body.list;
    var task=req.body.task;
    var status=req.body.status;
    
    var data={
      "name":name,
      "task":task,
      "status":status
    }
    dataBaseInsertion(data,"table");            //inserting into databse 
    res.send("testing");

});

App.get('/exportExcel',(req,res)=>{           //handles the export req from front end
    exportExcel().then((data)=>{              //exporting data from database to excel file
      var file = "customer.xlsx";
      res.download('./customer.xlsx',file);      //sending the file from server to client directory
    }).catch((error) =>{
      res.status(500).json({"message":"Internal Server Error"}).end();
    });
});

App.listen(3000);   //port used for deploying server