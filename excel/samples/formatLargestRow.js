var ctx = new Excel.ExcelClientContext();
var rows = ctx.workbook.tables.getItem("Table1").tableRows;
ctx.load(rows);
ctx.executeAsync().then(function () {
	var largestRow = 0;
	var largestValue = 0;
	
	for (var i = 0; i < rows.items.length; i++){
		if (rows.items[i].values[0][1] > largestValue){
			largestRow = i;
			largestValue = rows.items[i].values[0][1];
		}
	}
	
	var largestRowRng = rows.getItemAt(largestRow).getRange();
	largestRowRng.format.background.color = "#ff0000";
	
	ctx.executeAsync().then();
});	
