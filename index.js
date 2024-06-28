const express = require('express');
// const bodyParser = require('body-parser');
const nodemailer = require('nodemailer');

const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs');
const cors= require("cors");
const app = express();
app.use(express.json())
app.use(express.urlencoded({extended:true}))
app.use(cors());
const dotenv=require("dotenv").config();

const upload = multer();
app.listen(process.env.PORT,()=>{
    console.log("server is running..............")
})



const mongoose = require('mongoose');
mongoose.connect("mongodb://localhost:27017/ZsorkData")
.then(()=>{
    console.log("connected from database")
})
.catch(()=>{
    console.log("Not connected from database")
})

const UserSchema=new mongoose.Schema({
    username:String,
    phone:Number,
    message:String
});
const User= mongoose.model("User",UserSchema)
app.post("/contact", async(req,res)=>{
    const{username, phone, message}=req.body;
    console.log(username, phone ,message);

    const store=new User({
        username:username,
        phone:phone,
        message:message
    })
    console.log(store);
    await store.save()
    .then(()=>{
        console.log("data is saved")
    })
    .catch(()=>{
        console.log("data is Not saved") 
    })

    //   // Save to Excel
    //   const filePath = 'form_data.xlsx';
    //   const workbook = xlsx.utils.book_new();
    //   const header = ['Username', 'Phone', 'Message'];
    //   const data = [ header,
    //     [username, phone, message]];
  
    //   if (fs.existsSync(filePath)) {
    //     const existingWorkbook = xlsx.readFile(filePath);
    //     const existingSheet = existingWorkbook.Sheets[existingWorkbook.SheetNames[0]];
    //     const existingData = xlsx.utils.sheet_to_json(existingSheet, { header: 1 });
    //     data.unshift(...existingData);
    //   }

    //  Save to Excel
     const filePath = 'form_data2.xlsx';
     let workbook;
     let data;
 
     if (fs.existsSync(filePath)) {
       workbook = xlsx.readFile(filePath);
       const existingSheet = workbook.Sheets[workbook.SheetNames[0]];
       const existingData = xlsx.utils.sheet_to_json(existingSheet, { header: 1 });
       data = existingData.length > 0 ? existingData : [['Username', 'Phone', 'Message']];
     } else {
       workbook = xlsx.utils.book_new();
       data = [['Username', 'Phone', 'Message']];
     }
 
     data.push([username, phone, message]);
     const worksheet = xlsx.utils.aoa_to_sheet(data);
     xlsx.utils.book_append_sheet(workbook, worksheet, 'Form Data 2');
     xlsx.writeFile(workbook, filePath);
  
    //   const worksheet = xlsx.utils.aoa_to_sheet(data);
    //   xlsx.utils.book_append_sheet(workbook, worksheet, 'Form Data');
    //   xlsx.writeFile(workbook, filePath);
  
    //   res.send('Form submitted and data saved!');


        const transporter = nodemailer.createTransport({
          service: "gmail",
          auth: {
            user: "mehakmalhotra200479@gmail.com",
            pass: "szdw abwt gszx brvs"
          }
        });
        const mailOptions = {
          from: "mehakmalhotra200479@gmail.com",
          to: "yachnamalhotra01@gmail.com",
          subject: "Zsork Contact From ",
          text: `Name:${username} \n Phone Number:${phone} \n Message:${message}`, 
        };
        transporter.sendMail(mailOptions, (err, info) => {
          if (err) {
            console.error("Error sending Email", err);
            return res.json({ message: "Error Sending Email" });
          } else {
            console.log("Email has been sent: " + info.response);
            return res.json({ status: true, message: "Email Sent" });
          }
          
         
        });

     
    
})


