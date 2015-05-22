var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getByName("Chart1");	

chart.datalabels.visible = true;

ctx.executeAsync().then(function () {
		logComment("Datalabels Shown");
});