var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.series.GetItemAt(1).fillFormat.SetSolidColor("#FF0000");

ctx.executeAsync().then(function () {
		logComment("Series Fill Color Changed ");
});