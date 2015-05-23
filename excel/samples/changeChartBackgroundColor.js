var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.fillFormat.SetSolidColor("#FF0000");

ctx.executeAsync().then(function () {
		logComment("Chart Color Changed ");
});