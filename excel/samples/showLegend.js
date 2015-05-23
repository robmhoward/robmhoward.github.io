var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.legend.visible = true;

ctx.executeAsync().then(function () {
		logComment("Legend Shown ");
});