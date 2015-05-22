var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getByName("Chart1");	

chart.axes.valueaxis.title.text = "Catagory";

ctx.executeAsync().then(function () {
		logComment("Axis Title Added ");
});