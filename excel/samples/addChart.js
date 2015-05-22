var sheetName = "Charts";
var sourceData = sheetName + "!" + "E2:E5";

var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("ColumnClustered", sourceData, "rows");
ctx.executeAsync().then(function () {
		logComment"New Chart Added");
});