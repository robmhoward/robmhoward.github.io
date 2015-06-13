var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
ctx.load(chart);
ctx.executeAsync().then(function () {
		console.log(chart.name);
		console.log("done");
});