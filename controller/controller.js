const { db } = require("../db");
const ages = db.collection("Age");
const excelJs = require("exceljs");
const fs = require('fs')


const upload = async(req,res)=>{
    // const session = await mongoose.startSession();
    // session.startTransaction();
      try {
        const data = [];
        const errMsg = [];
        // Read excel file
        const workbook = new excelJs.Workbook();
        const result = await workbook.xlsx.readFile(req.file.path);
        const sheetCount = workbook.worksheets.length;
        if (sheetCount === 0) {
          errMsg.push({ message: "Workbook empty." });
        } else {
          for (let i = 0; i < sheetCount; i++) {
            let sheet = workbook.worksheets[i];
            sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
              if (rowNumber === 1 && row.cellCount === 2 ) {
                // Checking if Header exists
                if (!row.hasValues) {
                  errMsg.push({ status: "Error", message: "Empty Headers" });
                } else if (row.values[1] !== "Name" || row.values[2] !== "Age") {
                  errMsg.push({
                    location: "Row " + rowNumber,
                    message: "Incorrect Headers",
                  });
                }
              }
              // Checking only those rows which have a value
              else if (row.hasValues) {
                const alphabetRegex = new RegExp(/^[a-zA-Z]+$/);
                const numberRegex = new RegExp(/^[0-9]+$/);
                if(row.cellCount === 2){
                  if(row.values[1] !== undefined &&
                    row.values[2] !== undefined){
                          if (alphabetRegex.test(row.values[1]) &&
                  numberRegex.test(row.values[2]) )
                   {
                  data.push({ name: row.values[1], age: row.values[2] });
                  console.log(data)
                } else {
                  errMsg.push({
                    location: `Row:${rowNumber}`,
                    message: `name should have only letter,
          age should contain only number`,
                  });
                }
              }else{
                errMsg.push({
                  location: "Row " + rowNumber,
                  message: "one column in excel can not be empty.",
                });
              }
              }else{
                errMsg.push({
                  location: "Row " + rowNumber,
                  message: "excel should not have more than two columns",
                });
              }
            }
           });
          }
        }
        if (errMsg.length > 0) {
          res.status(400).json({
            status: "failed",
            response: errMsg,
            message: "Invalid excel sheet",
          });
        } else {
          const record = await ages.insertMany(data);
          res.status(200).json({
            status: "success",
            response: {data},
            message: "excel file uploaded successfully",
          });
        }
        // console.log("errMsg: ", errMsg);
        // console.log("data: ", data);
      } catch (error) {
        console.log(error.message);
        res.status(400).json({
          status: "failed",
          message: error.message,
        });
      }
  }


  module.exports = { upload }
