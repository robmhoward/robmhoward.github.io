var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.datalabels.position = "top";

ctx.executeAsync().then(function () {
		logComment("Datalabels Position Changed");
});