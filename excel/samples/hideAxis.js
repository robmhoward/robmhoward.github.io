var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.axes.valueaxis.visible=false;

ctx.executeAsync().then(function () {
		logComment("Value Axis Hidden");
});