//The code was inspired by the question here:
// https://stackoverflow.com/questions/69394723/incrementing-alphabet-counter-for-excel-online-in-office-script

//My updated code here utilizes various reusable functions and an
//interface to make the code in my initial answer more streamlined

function main(workbook: ExcelScript.Workbook) {
  let ws = workbook.getWorksheet("oldWorksheet");
  let ws2 = workbook.getWorksheet("newWorksheet");
  let rang = ws.getRange("A2");
  let data: (string|number|boolean)[][] = getDataFromSheetWithLoop(rang,true,2);
  let rang2 : ExcelScript.Range = ws2.getRange("A1");
  let newRange : ExcelScript.Range = getResizedRangeFromArray(rang2,data);
  newRange.setValues(data);
}

function getDataFromSheetWithLoop(dataRange: ExcelScript.Range, selectionCurrentRegion: boolean = false, step:number = 1): (string|number|boolean)[][]{
  if (selectionCurrentRegion === true){
    dataRange = dataRange.getSurroundingRegion();
  }
  let dataRowCount = dataRange.getRowCount();
  let dataColCount = dataRange.getColumnCount();
  let dataRangeVals: (string | number | boolean)[][] = dataRange.getValues();
  let newRangeVals: (string | number | boolean)[][] = [];
  let columnCounter = dataColCount / step;

  for (let i = 0; i < dataRowCount; i++) {
    let tempVals: (string | number | boolean)[] = [];
    for (let j = 0; j < dataColCount; j += step) {
      tempVals.push(dataRangeVals[i][j]);
    }
    newRangeVals.push(tempVals);
  }
  return newRangeVals;
}

function getRowsAndColumnsFromArray(rangeArray: (string | number | boolean)[][]): iRowsAndColumns {
  let rowsAndColumns: iRowsAndColumns = { rows: 0, columns: 0 };
  rowsAndColumns.rows = rangeArray.length;
  rowsAndColumns.columns = rangeArray[0].length;
  return rowsAndColumns;
}

function getResizedRangeFromArray(inputRange: ExcelScript.Range,rangeValues: (string|number|boolean)[][]): ExcelScript.Range{
  let rowsAndColumns : iRowsAndColumns = {rows:0,columns:0};
  rowsAndColumns =  getRowsAndColumnsFromArray(rangeValues);
  let newRange : ExcelScript.Range = inputRange.getResizedRange(rowsAndColumns.rows-1,rowsAndColumns.columns-1);
  return newRange;
}

interface iRowsAndColumns{
  rows: number,
  columns: number
}
