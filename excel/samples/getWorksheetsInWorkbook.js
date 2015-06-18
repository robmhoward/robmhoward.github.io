var ctx = new Excel.ExcelClientContext();
var worksheets = ctx.workbook.worksheets;
ctx.load(worksheets);
ctx.executeAsync().then(function() {
	for (var i = 0; i < worksheets.items.length; i++) {
		console.log(worksheets.items[i].name);
	}
	console.log("done");
});