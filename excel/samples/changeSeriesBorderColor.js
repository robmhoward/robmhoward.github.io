var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getByName("Chart1");	

chart.series.GetItemAt(1).lineFormat.color = "#FF0000";

ctx.executeAsync().then(function () {
		logComment("Series Border Color Changed  ");
});