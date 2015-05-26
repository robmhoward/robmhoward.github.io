var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

ctx.load(chart);
ctx.executeAsync().then(function () {
		console.log("Chart1 Loaded");
});