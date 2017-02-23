if(typeof require !== 'undefined') XLSX = require('xlsx');
var workbook = XLSX.readFile('grafik.xlsx');
/* DO SOMETHING WITH workbook HERE */
var first_sheet_name = workbook.SheetNames[0];
var worksheet = workbook.Sheets[first_sheet_name];
