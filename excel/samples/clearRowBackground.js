var ctx = new Excel.ExcelClientContext();
var rows = ctx.workbook.tables.getItem("Table1").tableRows;
ctx.load(rows);
ctx.executeAsync().then(function () {
	for (var i = 0; i < rows.items.length; i++){
		var row = rows.getItemAt(i).getRange();
		ctx.load(row);
		row.format.background.color = null;
		ctx.executeAsync().then();
	}
});	
