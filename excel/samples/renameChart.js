var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getByName("Chart1");
	
chart.name="NewChartName";
ctx.executeAsync().then(function () {
		logComment"Chart Renamed");
});