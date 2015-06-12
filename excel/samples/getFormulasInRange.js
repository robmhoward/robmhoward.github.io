var ctx = new Excel.ExcelClientContext();
var range = ctx.workbook.worksheets.getItem("Sheet1").getRange("A1:C3");
ctx.load(range);
ctx.executeAsync().then(function() {
	for (var i = 0; i < range.formulas.length; i++) {
		for (var j = 0; j < range.formulas[i].length; j++) {
			console.log(range.formulas[i][j]);
		}
	}
	console.log("done");
});