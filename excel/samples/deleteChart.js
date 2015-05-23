var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.deleteObject();
ctx.executeAsync().then(function () {
		logComment"Chart Deleted");
});