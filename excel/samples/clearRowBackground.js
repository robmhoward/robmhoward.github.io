var ctx = new Excel.ExcelClientContext();
RichApiTest.log.comment("Format the row with the largest value in column 2 in red.");
var rows = ctx.workbook.tables.getItem("Table1").tableRows;
ctx.load(rows);
ctx.executeAsync().then(function () {

	for (var i = 0; i < rows.items.length; i++){
		RichApiTest.log.comment(i.toString());
		var row = rows.getItemAt(i).getRange();
		ctx.load(row);
		row.format.background.color = null;
		ctx.executeAsync().then();
	}
});	
