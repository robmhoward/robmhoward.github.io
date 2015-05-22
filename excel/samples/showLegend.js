var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getByName("Chart1");	

chart.legend.visible = true;

ctx.executeAsync().then(function () {
		logComment("Legend Shown ");
});