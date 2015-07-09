var ctx = new Excel.RequestContext();
var rows = ctx.workbook.tables.getItem("Table1").tableRows.load();
ctx.executeAsync().then(function () {
	
	for (var i = 0; i < rows.items.length; i++){
		
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
