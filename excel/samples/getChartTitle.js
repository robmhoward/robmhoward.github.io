var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
var title = chart.title;
ctx.load(title);
ctx.executeAsync().then(function () {
		logComment(title.text);
});