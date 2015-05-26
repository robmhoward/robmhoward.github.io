var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.axes.valueAxis.visible=false;

ctx.executeAsync().then(function () {
		console.log("Value Axis Hidden");
});