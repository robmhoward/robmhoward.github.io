var ctx = new Excel.ExcelClientContext();
var selectedRange = ctx.workbook.getSelectedRange();
ctx.load(selectedRange);
ctx.executeAsync().then(function () {
	console.log(selectedRange.address);
	console.log("done");
});