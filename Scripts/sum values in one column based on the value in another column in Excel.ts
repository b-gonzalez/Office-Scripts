//sums the values of a summed column based on the
//unique values from another column. Based on a
//question from StackOverflow that can be found
//here:

//https://stackoverflow.com/questions/71779149/typescript-code-to-sum-values-in-one-column-based-on-the-value-in-another-colum

function main(workbook: ExcelScript.Workbook) {
	let ws: ExcelScript.Worksheet = workbook.getWorksheet("Sheet1")
	let criteriaStart: string = "A1"
	let sumStart: string = "B1"
	let map: Map<string, number> = getColumnTotals(ws, criteriaStart, sumStart);
	console.log(map.get('A'));
	console.log(map.get('B'));
	console.log(map.get('C'));

	//map.forEach(e => console.log(e))
}

function getColumnTotals(worksheet: ExcelScript.Worksheet, criteriaStartAddress: string, sumStartAddress: string): Map<string, number> {
	let criteriaRange: ExcelScript.Range = worksheet.getRange(criteriaStartAddress).getExtendedRange(ExcelScript.KeyboardDirection.down);
	let criteriaVals: string[][] = criteriaRange.getValues() as string[][];
	let sumRange: ExcelScript.Range = worksheet.getRange(sumStartAddress).getExtendedRange(ExcelScript.KeyboardDirection.down);
	let sumVals: number[][] = sumRange.getValues() as number[][];
	let map: Map<string, number> = new Map<string, number>();
	let tempArr: string[] = criteriaVals.map(e => e[0]);
	let uniqueCalendarVals: string[] = Array.from(new Set(tempArr));

	uniqueCalendarVals.forEach(uniqueCalVal => {
		let tempTotal: number = 0;
		criteriaVals.forEach((criVal, index) => {
			if (criVal[0] === uniqueCalVal) {
				tempTotal += sumVals[index][0];
			}
		})
		map.set(uniqueCalVal, tempTotal);
	});
	return map;
}