var ctx = new Excel.ExcelClientContext();
var tableRows = ctx.workbook.tables.getItemAt(0).tableRows;
ctx.load(tableRows);
ctx.executeAsync().then(function () {
	for (var i = 0; i < tableRows.items.length; i++) {
		console.log(tableRows.items[i].values);
	}
	console.log("done");
});