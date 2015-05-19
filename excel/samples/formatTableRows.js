var ctx = new Excel.ExcelClientContext();
var table = ctx.workbook.tables.getItem("Table1");
ctx.load(table);
ctx.executeAsync().then(function () {
	
	var rows = table.tableRows;
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
});
