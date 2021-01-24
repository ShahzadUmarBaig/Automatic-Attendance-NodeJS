const ExcelForNode = require("excel4node");
const xlsxFile = require("read-excel-file/node");
const ExcelJS = require("exceljs");
const express = require("express");
var app = express();
var bodyParser = require("body-parser");

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";

var port = process.env.PORT || 3000; // set our port

app.get("/:student", async (req, res) => {
  var studentCode = req.query.student_code;
  var studentName = req.params.student;

  var columnIndex = await getRowLength();
  console.log(columnIndex);
  var rowIndex = await getName(studentCode);
  console.log(rowIndex);
  await editingExcelFile(rowIndex, columnIndex);
  res.send(`${studentName}` + `Your Attendance has been marked`);
});

app.listen(port);
console.log("Server is live at " + port);

async function main() {
  var columnIndex = await getRowLength();
  console.log(columnIndex);
  var rowIndex = await getName("a3b1c2");
  console.log(rowIndex);
  await editingExcelFile(rowIndex, columnIndex);
}

async function getName(code) {
  const workbook = new ExcelJS.Workbook();
  var rowNumberToWrite = 0;
  await workbook.xlsx.readFile("Excel.xlsx").then(async () => {
    var sheetOne = workbook.getWorksheet("Sheet 1");
    sheetOne.eachRow((row, rowNumber) => {
      console.log(row.values);
      if (row.values[4] == code) {
        rowNumberToWrite = rowNumber;
      }
    });
  });

  return rowNumberToWrite;
}

async function getRowLength() {
  var ts = Date.now();
  let date_ob = new Date(ts);
  let date = date_ob.getDate();
  let month = date_ob.getMonth() + 1;
  let year = date_ob.getFullYear();
  var dateToday = `${date}/${month}/${year}`;
  const workbook = new ExcelJS.Workbook();
  var rowValues = [];
  var allNames;
  var sheetOne;
  await workbook.xlsx.readFile("Excel.xlsx").then(async () => {
    sheetOne = await workbook.getWorksheet("Sheet 1");
    allNames = await sheetOne.getRow(1);
    allNames.eachCell((cell, colNum) => {
      rowValues.push(cell.value);
    });
  });
  var rowLastIndex = rowValues.length;

  if (rowValues[rowLastIndex - 1] == dateToday) {
    console.log("This Condition Is Working");
    return rowLastIndex;
  } else {
    var row = sheetOne.getRow(1);
    row.getCell(rowLastIndex + 1).value = dateToday;
    row.commit();
    workbook.xlsx.writeFile("Excel.xlsx");
    return rowLastIndex + 1;
  }
}

async function editingExcelFile(rowIndex, colIndex) {
  var workbook = new ExcelJS.Workbook();

  workbook.xlsx.readFile("Excel.xlsx").then(async function () {
    var worksheet = await workbook.getWorksheet(1);
    var row = await worksheet.getRow(rowIndex);
    row.getCell(colIndex).value = "P"; // A5's value set to 5
    row.commit();
    return workbook.xlsx.writeFile("Excel.xlsx");
  });
}

// main();
