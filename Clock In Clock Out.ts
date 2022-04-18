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
      let clockDuration = Math.round((clockOutTime.getTime() - clockInTime.getTime()) / 1000 / 60)
      duration.getRangeBetweenHeaderAndTotal().getLastRow().setValue(clockDuration);
    }
    else {
      tbl.addRow()
      clockInLastRow.getOffsetRange(1,0).setValue(date.toLocaleString());
    }
  }