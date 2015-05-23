var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.datalabels.ShowSeriesName = true;

ctx.executeAsync().then(function () {
		logComment("Series Name Added to Datalabels");
});