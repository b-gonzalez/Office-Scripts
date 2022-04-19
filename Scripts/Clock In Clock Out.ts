function main(workbook:ExcelScript.Workbook){
  
  let sheetName = "Sheet1"
  let tableName = "Table1"

  let sh: ExcelScript.Worksheet = workbook.getWorksheet(sheetName);
  let tbl: ExcelScript.Table = sh.getTable(tableName);
  let clockIn: ExcelScript.TableColumn = tbl.getColumnByName("Clock in");
  let clockOut: ExcelScript.TableColumn = tbl.getColumnByName("Clock Out");
  let duration: ExcelScript.TableColumn = tbl.getColumnByName("Duration");
  let clockInLastRow: ExcelScript.Range = clockIn.getRangeBetweenHeaderAndTotal().getLastRow();
  let clockOutLastRow: ExcelScript.Range = clockOut.getRangeBetweenHeaderAndTotal().getLastRow();
  
  let date: Date = new Date();

  
  if (clockInLastRow.getValue() as string === ""){
    clockInLastRow.setValue(date.toLocaleString());
  } else if (clockOutLastRow.getValue() as string === "") {
    clockOutLastRow.setValue(date.toLocaleString());
    let clockInTime: Date = new Date(clockInLastRow.getValue() as string);
    let clockOutTime: Date = new Date(clockOutLastRow.getValue() as string);
    let clockDuration = Math.abs((clockOutTime.getTime() - clockInTime.getTime()))

    let durationString = getDurationMessage(clockDuration)

    duration.getRangeBetweenHeaderAndTotal().getLastRow().setValue(durationString);
  } else {
    tbl.addRow()
    clockInLastRow.getOffsetRange(1,0).setValue(date.toLocaleString());
  }
}

function getDurationMessage(delta : number){
//adapted from here: https://stackoverflow.com/questions/13903897/javascript-return-number-of-days-hours-minutes-seconds-between-two-dates
    // get total seconds between the times

  delta = delta / 1000

  let durationString = ""

  // calculate (and subtract) whole days
  let days = Math.floor(delta / 86400)

  if (days >= 1) {
    durationString += days
    days > 1 ? durationString += " days " : durationString += " day "
  }

  delta -= days * 86400;

  // calculate (and subtract) whole hours
  let hours = Math.floor(delta / 3600) % 24;
  if (hours >= 1) {
    durationString += hours
    hours > 1 ? durationString += " hours " : durationString += " hour "
  }
  delta -= hours * 3600;

  // calculate (and subtract) whole minutes
  let minutes = Math.floor(delta / 60) % 60;
  if (minutes >= 1) {
    durationString += minutes
    minutes > 1 ? durationString += " minutes " : durationString += " minute "
  }

/*delta -= minutes * 60;

// what's left is seconds
let seconds = delta % 60;  // in theory the modulus is not required*/

  return durationString
}