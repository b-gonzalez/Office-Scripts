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
  //adapted from here: 
  //https://stackoverflow.com/questions/13903897/javascript-return-number-of-days-hours-minutes-seconds-between-two-dates

  delta = delta / 1000

  let durationString = ""

  let days = Math.floor(delta / 86400)
  delta -= days * 86400;

  let hours = Math.floor(delta / 3600) % 24;
  delta -= hours * 3600;

  let minutes = Math.floor(delta / 60) % 60;

  if (days >= 1) {
    durationString += days
    durationString += (days > 1 ? " days" : " day")

    if (hours >= 1 && minutes >=1){
      durationString += ", "
    }
    else if (hours >= 1 || minutes > 1){
      durationString += " and "
    }
  }
  
  if (hours >= 1) {
    durationString += hours
    durationString += (hours > 1 ? " hours" : " hour")
    if (minutes>=1) durationString += " and "
  }

  if (minutes >= 1) {
    durationString += minutes
    durationString += (minutes > 1 ? " minutes" : " minute")
  }

  return durationString
}