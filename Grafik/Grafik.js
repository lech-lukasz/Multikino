var fs = require('fs');
if(typeof require !== 'undefined') XLSX = require('xlsx');
var workbook = XLSX.readFile('Grafik.xlsx');
/* DO SOMETHING WITH workbook HERE */
var first_sheet_name = workbook.SheetNames[0];
var worksheet = workbook.Sheets[first_sheet_name];

var events = [];

var namesCol = "C";
var scheduleCol = "D";

var stringRowNumber = 4;
var dateRow = 2;

function nextChar(c, i) {
    return String.fromCharCode(c.charCodeAt(0) + i);
}

function searchRowNumberForEmployee(startingRowNumber, employeeName) {
    var rowNumber = stringRowNumber;

    while((worksheet[namesCol + String(rowNumber)].v) != employeeName) {
        rowNumber++;
    }
    return rowNumber;
}

function dateParser(dateToParse) {
    var date = dateToParse.toString().split("/");

    var day = date[1];
    var month = date[0];
    var year = "20" + date[2];

    var parsedDate = year + "-" + month + "-" + day;

    return parsedDate;
}

var rowNumberForEmployee = searchRowNumberForEmployee(stringRowNumber, "LEC")

while (scheduleCol < "Y") {

    if(worksheet[scheduleCol + rowNumberForEmployee].w != '0') {

        var event = new Object();

        event.summary = worksheet[scheduleCol + rowNumberForEmployee].w;

        scheduleCol = nextChar(scheduleCol, 1);

        var date = dateParser(worksheet[scheduleCol + String(dateRow)].w);

        var startTime = worksheet[scheduleCol + rowNumberForEmployee].w + ":00";
        event.start = new Object();
        event.start.dateTime = date + "T" + startTime;
        event.start.timeZone = "Europe/Warsaw";


        scheduleCol = nextChar(scheduleCol, 1);

        var endTime = worksheet[scheduleCol + rowNumberForEmployee].w + ":00";
        event.end = new Object();
        event.end.dateTime = date + "T" + endTime;
        event.end.timeZone = "Europe/Warsaw";

        scheduleCol = nextChar(scheduleCol, 1);

        event.location = "Arkady Wrocławskie, Powstańców Śląskich 2-4";

        events.push(event);

    } else {
        scheduleCol = nextChar(scheduleCol, 3);
    }
}

fs.writeFile('schedule.json', JSON.stringify(events), (err) => {
  if (err) throw err;
  console.log('It\'s saved!');
});
