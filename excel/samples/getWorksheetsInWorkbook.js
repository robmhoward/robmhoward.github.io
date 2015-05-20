var ctx = new Excel.ExcelClientContext();
var worksheets = ctx.workbook.worksheets;
ctx.load(worksheets);
ctx.executeAsync().then(function() {
	for (var i = 0; i < worksheets.items.length; i++) {
		logComment(worksheets.items[i].name);
	}
});