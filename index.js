const axios = require('axios');
const { Client, Pool } = require("pg");
const Query = require('pg').Query;
const ExcelJS = require('exceljs');
const fs = require('fs');


const config ={
    user : 'postgres', 
    host : 'localhost',
    database : 'postgres', 
    password : 'postgres', 
    port : 5432,
    max: 20,
    idleTimeoutMillis:30000
  }

  
var pool = new Pool(config);

// add query functions
function getSqlData() {
    const text = "SELECT test1"
	+" FROM test limit 1";

      pool.query(text)
      .then(data => {
        console.log(data.rows[0].test1) // Array
      })
      .catch(e => {console.log(e)})

}

// ResultExcel = Teamplate + excelData
async function loadTemplate() {

    let sourceWorkbook = new ExcelJS.Workbook();
    let dataWorkbook = new ExcelJS.Workbook();
    
    dataWorkbook = await dataWorkbook.xlsx.readFile('data.xlsx');
    const dataWroksheet = dataWorkbook.getWorksheet(1);
    


    dataWroksheet.eachRow({ includeEmpty: false },async (row, rowNumber) => {
        if(rowNumber>2&&rowNumber<53){
            let targetWorkbook = new ExcelJS.Workbook();
            let targetWorksheet = targetWorkbook.addWorksheet();
            sourceWorkbook = await sourceWorkbook.xlsx.readFile('template.xlsx');
            const sourceWorksheet = sourceWorkbook.getWorksheet(1);
            targetWorksheet.model = Object.assign(sourceWorksheet.model, {
                mergeCells: sourceWorksheet.model.merges
            });
            sourceWorksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                var newRow = targetWorksheet.getRow(rowNumber);
                row.eachCell((cell, colNumber) => {
                    var newCell = newRow.getCell(colNumber)
                    for(var prop in cell)
                    {
                        newCell[prop] = cell[prop];
                    }
                })
             });

            
            await targetWorkbook.xlsx.writeFile('result/'+rowNumber+'.xlsx');
        }
     });


    
}

getSqlData()
loadTemplate();