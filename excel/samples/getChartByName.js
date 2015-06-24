var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getByName("Chart1");	
ctx.load(chart);
ctx.executeAsync().then(function () {
		console.log(chart.name);
		console.log("done");
});