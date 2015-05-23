var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.axes.valueaxis.majorgridlines.visible = true;

ctx.executeAsync().then(function () {
		logComment("Axis Title Added ");
});