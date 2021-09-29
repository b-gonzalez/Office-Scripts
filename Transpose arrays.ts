//Utility functions to transpose arrays

//Function 1: Transforms a single dimensional array to a multidimensional array
//so that it can be used with the setValues method of the range object.
//This utilizes the setValueType enum so that the toColumns or toRows
//value can be specified.

function main(workbook: ExcelScript.Workbook)
{
  let arr = [1,2,3,4,5]
  console.log(arrayForSetValues(arr,setValueType.toColumns))
  console.log(arrayForSetValues(arr, setValueType.toRows))
}

function arrayForSetValues(arr: (string | number | boolean)[], setType: setValueType): (string | number | boolean)[][]{
  let tempArr: (string | number | boolean)[][] = []
  if (setType === setValueType.toRows){ 
    arr.forEach(value => {
      tempArr.push([value])
    })
  } else if (setType === setValueType.toColumns) {
    let toColumnsArr: (string | number | boolean)[] = []
    arr.forEach((value) => {
      toColumnsArr.push(value)
    })
    tempArr.push(toColumnsArr)
  }
  return tempArr
}

enum setValueType{
  toColumns,
  toRows
}

//Example 2: This function transposes an array. It's intended to be used with the getValues function 
//of the range object to transpose one type of array to the other.

function main(workbook:ExcelScript.Workbook) {
  let ws = workbook.getActiveWorksheet()
  //let rang = ws.getRange("E1:E3")
  let rang = ws.getRange("E1:G1")
  //let rang2 = ws.getRange("E1:G1")
  let rang2 = ws.getRange("E1:E3")
  console.log(rang.getValues())
  let vals : (string|number|boolean)[][] = transposeValues(rang.getValues())
  rang2.setValues(vals)
}

function transposeValues(inputValues: (string | number | boolean)[][]): (string | number | boolean)[][] {
  let tempArr: (string | number | boolean)[][] = []
  if (inputValues.length === 1){
    let arr: (string | number | boolean)[] = inputValues[0]
    arr.forEach(value=>{
      tempArr.push([value])
    })
  } else {
    let elseArr: (string|number|boolean)[] = []
    inputValues.forEach((value,index)=>{
      elseArr.push(value[0])
    })
    tempArr.push(elseArr)
  }
  return tempArr
}
