var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getByName("Chart1");	

chart.height =200;
chart.width =200;
ctx.executeAsync().then(function () {
		logComment("Chart Resized");
});