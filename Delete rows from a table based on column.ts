//Office Script to delete rows from a table.
//Requires updating of sheetName, tableName,
//columnName, and columnValue variables with
//the values required in the workbook.
//
//Code is in reponse to this question on SO:
//https://stackoverflow.com/questions/70343448/how-to-delete-special-cells

function main(workbook: ExcelScript.Workbook) {

  let sheetName: string = "Sheet3";
  let tableName: string = "Table2";
  let columnName: string = "Col19";
  let columnValue: string = ""

  let sh: ExcelScript.Worksheet = workbook.getWorksheet(sheetName);
  let tbl: ExcelScript.Table = workbook.getTable(tableName);
  let tableColumn: ExcelScript.TableColumn = tbl.getColumn(columnName);
  let tableColumnRange: ExcelScript.Range = tableColumn.getRangeBetweenHeaderAndTotal();
  let tableColumnValues: (String|Number|Boolean)[][] = tableColumnRange.getValues();
  let rowCount: number = tableColumnRange.getRowCount();
  let colCount: number = tableColumnRange.getColumnCount();

  for (let i = rowCount-1; i>=0; i--){
      if (tableColumnValues[i][0] == columnValue){
        tbl.deleteRowsAt(i);
    }
  }
}
