//Originally inspired by this question: https://stackoverflow.com/questions/69322235/combine-specific-columns-from-several-tables-using-excel-office-script-to-one-ta

//Here are a few additional examples

//Iterating throuth the tables collection, iterating through the values of certain columns in all of the tablse,
//and merging those table values into a final column in the final table

function main(workbook: ExcelScript.Workbook) {

  let columnNames : string[] = ["ColA","ColB","ColC"]

  const columnNamesAndCellValues = {
    "ColA": [],
    "ColB": [],
    "ColC": []
  }

  let tables : ExcelScript.Table[] = workbook.getTables()

  tables.forEach(tbl=>{
    columnNames.forEach(column=>{
      let tableColumn: ExcelScript.Range = tbl.getColumn(column).getRangeBetweenHeaderAndTotal()
      tableColumn.getValues().forEach(value=>{
        columnNamesAndCellValues[column].push(value)
        })
    })
  })

  let combinedSheet : ExcelScript.Worksheet = workbook.getWorksheet("Combined")

  combinedSheet.activate()

  let headerRange : ExcelScript.Range = combinedSheet.getRangeByIndexes(0,0,1,columnNames.length)

  headerRange.setValues([columnNames])

  columnNames.forEach((column,index)=>{
    combinedSheet.getRangeByIndexes(1, index, columnNamesAndCellValues[column].length, 1).setValues(columnNamesAndCellValues[column])
    })

  let combinedTableAddress : string = combinedSheet.getRange("A1").getSurroundingRegion().getAddress()

  combinedSheet.addTable(combinedTableAddress,true)
}


//Using sheets to pull specific columns data from the underlying tables that they contain.
//All column names in this example are expected to be unique


function main(workbook: ExcelScript.Workbook) {

  //JSON object called SheetAndColumnNames. On the left hand side is the sheet name. 
 //On the right hand side is an array with the column names for the table in the sheet.

  //NOTE: replace the sheet and column names with your own values
  const sheetAndColumnNames = {
    "Sheet1": ["ColA"],
    "Sheet2": ["ColD", "ColF"]
  }

  //Array to hold the column values we get from the tables. 
 //These values will include both the header name and column values
  let tableColumnData : (string|number|boolean)[][][] = []

  //Iterate through the JSON object
  for (let sheetName in sheetAndColumnNames) {

      //Use sheet name from JSON object to get sheet
      let sheet: ExcelScript.Worksheet = workbook.getWorksheet(sheetName)

      //get table from the previously assigned sheet
      let table: ExcelScript.Table = sheet.getTables()[0]

    //get array of column names to be iterated on the sheet
    let tableColumnNames: string[] = sheetAndColumnNames[sheetName]

      //Iterate the array of table column names
      tableColumnNames.forEach(columnName=> {

        //get range from the table for the current column name
        let tableColumn : ExcelScript.Range = table.getColumn(columnName).getRange()

        //Add values from the table column to the table column data array
         tableColumnData.push(tableColumn.getValues())
      })
  }
  //Delete previous worksheet named Combined
  workbook.getWorksheet("Combined")?.delete()

  //Add new worksheet named Combined and assign to combinedSheet variable
  let combinedSheet : ExcelScript.Worksheet = workbook.addWorksheet("Combined")

  //Activate the combined sheet
  combinedSheet.activate()

  //iterate through the tableColumnData array to write to the Combined sheet
  tableColumnData.forEach((column,index)=>{
    combinedSheet.getRangeByIndexes(0,index,column.length,1).setValues(column)
  })

  //Get the address for the current region of the data written
  //from the tableColumnData array to the sheet
  let combinedTableAddress : string = combinedSheet.getRange("A1").getSurroundingRegion().getAddress()

  //Add the table to the sheet using the address and setting the hasHeaders boolean value to true
  combinedSheet.addTable(combinedTableAddress,true)
}
