var ctx = new Excel.RequestContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getByName("Chart1").load();
ctx.executeAsync().then(function () {
		console.log(chart.name);
		console.log("done");
});