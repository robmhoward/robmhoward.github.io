var ctx = new Excel.ExcelClientContext();
RichApiTest.log.comment("Format rows greater than 2 in green, other rows in red.");
var rows = ctx.workbook.tables.getItem("Table1").tableRows;
ctx.load(rows);
ctx.executeAsync().then(function () {
	
	RichApiTest.log.comment(rows.items.length.toString());
	for (var i = 0; i < rows.items.length; i++){
		
		RichApiTest.log.comment(rows.items[i].values[0][1]);
		var rng = rows.getItemAt(i).getRange();
		
		if (rows.items[i].values[0][1] > 2){
			rng.format.background.color = "#ff0000";
		}
		else{
			rng.format.background.color = "#00ff00";
		}
		ctx.executeAsync().then();
	}
});	
