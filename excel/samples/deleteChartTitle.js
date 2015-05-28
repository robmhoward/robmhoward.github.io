var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.title.visible = false; 
ctx.executeAsync().then(function () {
		logComment("Title Hidden");
});