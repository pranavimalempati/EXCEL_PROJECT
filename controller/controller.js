const { db } = require("../db");
const personal = db.collection("Personal_details");
const excelJs = require("exceljs");



const uploadfile = async (req, res) => {
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
                    if (rowNumber === 1 && row.cellCount === 5) {
                        // Checking if Header exists
                        if (!row.hasValues) {
                            errMsg.push({ status: "Error", message: "Empty Headers" });
                        } else if (row.values[1] !== "FirstName" || row.values[2] !== "LastName" || row.values[3] !== "Email" || row.values[4] !== "Gender" || row.values[5] !== "Age") {
                            errMsg.push({
                                location: "Row " + rowNumber,
                                message: "Incorrect Headers",
                            });
                        }
                    }
                    // Checking only those rows which have a value
                    else if (row.hasValues) {
                        const alphabetRegex = new RegExp(/^[a-zA-Z]+$/);
                        const numberRegex = new RegExp(/^[0-9]{2}$/);
                        const emailRegex = new RegExp(/^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*$/)
                        if (row.cellCount === 5) {
                            if (row.values[1] !== undefined && row.values[2] !== undefined && row.values[3] !== undefined
                                && row.values[4] !== undefined && row.values[5] !== undefined) {
                                if (alphabetRegex.test(row.values[1]) && alphabetRegex.test(row.values[2])
                                    && emailRegex.test(row.values[3]) && alphabetRegex.test(row.values[4]) && numberRegex.test(row.values[5])) {
                                    data.push({
                                        firstname: row.values[1], lastname: row.values[2],
                                        email: row.values[3], gender: row.values[4], age: row.values[5]
                                    });
                                } else {
                                    errMsg.push({
                                        location: `Row:${rowNumber}`,
                                        message: 'firstname,lastname and gender should have only letter,'+
                                                   'invalid email prototype,'+
                                                'age should contain only two digit number.'
                                     });
                                }
                            } else {
                                errMsg.push({
                                    location: "Row " + rowNumber,
                                    message: "one column in excel can not be empty.",
                                });
                            }
                        } else {
                            errMsg.push({
                                location: "Row " + rowNumber,
                                message: "excel should not have more than five columns",
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
            const record = await personal.insertMany(data);
            res.status(200).json({
                status: "success",
                response: { data },
                message: "excel file uploaded successfully",
            });
        }
    } catch (error) {
        console.log(error.message);
        res.status(400).json({
            status: "failed",
            message: error.message,
        });
    }
}


module.exports = { uploadfile }
