var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getByName("Chart1");	

ctx.load(chart);
ctx.executeAsync().then(function () {
		logComment(chart.title.text);
});