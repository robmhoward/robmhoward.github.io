var ctx = new Excel.ExcelClientContext();
var tables = ctx.workbook.tables;
ctx.load(tables);
ctx.executeAsync().then(function () {
	for (var i = 0; i < tables.items.length; i++)
	{
		logComment(tables.items[i].name);
	}
});