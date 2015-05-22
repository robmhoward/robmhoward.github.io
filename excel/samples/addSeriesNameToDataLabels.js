var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getByName("Chart1");	

chart.datalabels.ShowSeriesName = true;

ctx.executeAsync().then(function () {
		logComment("Series Name Added to Datalabels");
});