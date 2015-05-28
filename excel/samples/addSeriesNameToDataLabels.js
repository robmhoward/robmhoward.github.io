var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.datalabels.ShowSeriesName = true;

ctx.executeAsync().then(function () {
		logComment("Series Name Added to Datalabels");
});