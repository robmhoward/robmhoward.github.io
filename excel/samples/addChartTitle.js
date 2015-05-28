var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.title.text="New Title";
ctx.executeAsync().then(function () {
		logComment("Title Added");
});