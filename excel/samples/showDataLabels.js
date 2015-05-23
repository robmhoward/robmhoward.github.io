var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.datalabels.visible = true;

ctx.executeAsync().then(function () {
		logComment("Datalabels Shown");
});