var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.title.text="New Title";
ctx.executeAsync().then(function () {
		logComment("Title Added");
});