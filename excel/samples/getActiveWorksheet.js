var ctx = new Excel.ExcelClientContext();
var activeWorksheet = ctx.workbook.getActiveWorksheet();
ctx.load(activeWorksheet);
ctx.executeAsync().then(function () {
	logComment(activeWorksheet.name);
});