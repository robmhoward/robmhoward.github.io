var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");
	
chart.name="NewChartName";
ctx.executeAsync().then(function () {
		logComment("Chart Renamed");
});