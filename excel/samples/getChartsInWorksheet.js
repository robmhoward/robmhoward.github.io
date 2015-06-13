var ctx = new Excel.ExcelClientContext();
var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
ctx.load(charts);
ctx.executeAsync().then(function () {
	for (var i = 0; i < charts.items.length; i++) {
		console.log(charts.items[i].name);
	}
	console.log("done");
});