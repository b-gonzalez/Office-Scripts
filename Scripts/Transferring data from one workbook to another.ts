//Documents the code from transferring data from one workbook to
//another on Excel Online. This solutions requires PowerAutomate
//to run the script for both of the office scripts referenced.
//This snippet will just have the code samples. However, more
//details on the PowerAutomate required can be found on
//stackoverflow here: 
//
//https://stackoverflow.com/questions/70463127/excel-online-workbook-links-linking-full-row-range

//Scenario one: transfer a cell value from one cell in one
//workbook to another cell in a different workbook
    
    //first script
    function main(workbook: ExcelScript.Workbook): string {
        let sh: ExcelScript.Worksheet = workbook.getActiveWorksheet();
        
        //set the range of the cell to get the value from
        let rang: ExcelScript.Range = sh.getRange("A1");
        
        //assign the result variable to the value in the cell
        //and convert it to a string.
        let result: string = rang.getValue() as string;
        
        //return cell's value as a string
        return result;
    }

    //second script

    function main(workbook: ExcelScript.Workbook, cellValue: string)
    {
    let sh: ExcelScript.Worksheet = workbook.getWorksheet("Sheet1");
    
    //set the range of the cell where the value is to be written
    let rang: ExcelScript.Range = sh.getRange("A2");
    //sets the value of the resized range to the array
    rang.setValues(cellValue);
    }


//Scenario two: transfer a 2d array from an Excel table in
//one workbook to a range in another workbook
    
    //first script 
    function main(workbook: ExcelScript.Workbook): string {
        let sh: ExcelScript.Worksheet = workbook.getActiveWorksheet();
        
        //get table
        let tbl: ExcelScript.Table = sh.getTable("Table1");
        
        //get table's column count
        let tblColumnCount: number = tbl.getColumns().length;
        
        //set number of columns to keep
        let columnsToKeep: number = 3;
        
        //set the number of rows to remove
        let rowsToRemove: number = 0;
        
        //resize the table range
        let tblRange: ExcelScript.Range = tbl.getRangeBetweenHeaderAndTotal().getResizedRange(rowsToRemove,columnsToKeep - tblColumnCount);
        
        //get the table values
        let tblRangeValues: string[][] = tblRange.getValues() as string[][];
        
        //create a JSON string
        let result: string = JSON.stringify(tblRangeValues);
        
        //return JSON string
        return result;
    }

    //second script
    function main(workbook: ExcelScript.Workbook, tableValues: string) {
        let sh: ExcelScript.Worksheet = workbook.getWorksheet("Sheet1")
        
        //parses the JSON string to create array
        let tableValuesArray: string[][] = JSON.parse(tableValues);
        
        //gets row count from the array
        let valuesRowCount: number = tableValuesArray.length - 1
        
        //gets column count from the array
        let valuesColumnCount: number = tableValuesArray[0].length - 1
        
        //resizes the range
        let rang: ExcelScript.Range = sh.getRange("A1").getResizedRange(valuesRowCount,valuesColumnCount)
        
        //sets the value of the resized range to the array
        rang.setValues(tableValuesArray)
    }

//Scenario three: transfer a 2d array from an Excel table in
//one workbook to a range in another workbook
    
    //first script 
    function main(workbook: ExcelScript.Workbook): string {
        let sh: ExcelScript.Worksheet = workbook.getActiveWorksheet();

        //get table
        let tbl: ExcelScript.Table = sh.getTable("Table1");

        //get table's column count
        let tblColumnCount: number = tbl.getColumns().length;
        
        //set number of columns to keep
        let columnsToKeep: number = 3;
        
        //set the number of rows to remove
        let rowsToRemove: number = 0;
        
        //resize the table range
        let tblRange: ExcelScript.Range = tbl.getRangeBetweenHeaderAndTotal().getResizedRange(rowsToRemove,columnsToKeep - tblColumnCount);
        
        //get the table values
        let tblRangeValues: string[][] = tblRange.getValues() as string[][];
        
        //create a JSON string
        let result: string = JSON.stringify(tblRangeValues);
        
        //return JSON string
        return result;
    }

    //second script
    function main(workbook: ExcelScript.Workbook, tableValues: string) {
        let sh: ExcelScript.Worksheet = workbook.getWorksheet("Sheet1")

        //set the table that will hold the table values
        let resultTable: ExcelScript.Table = ws.getTable("Table1")
        
        //get the rowCount of the dataBodyRange for the table
        let rowCount: number = resultTable.getRangeBetweenHeaderAndTotal().getRowCount()
        
        //Delete the rowCount if it's one or more. It looks like this
        //code will throw an error if the table doesn't have any data in it.
        if (rowCount > 0){
        resultTable.getRangeBetweenHeaderAndTotal().delete(ExcelScript.DeleteShiftDirection.up)
        } 

        //parse the JSON string to a JSON array
        let tableValuesArray: string[][] = JSON.parse(tableValues);

        //update the table with the new values from the array
        resultTable.addRows(null,tableValuesArray)
    }