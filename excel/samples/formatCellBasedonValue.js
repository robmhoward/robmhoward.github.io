var sheetName = "Sheet1";
var rangeAddress = "A1:A9";
var ctx = new Excel.ExcelClientContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

ctx.load(range);
ctx.references.add(range);
ctx.executeAsync().then(function () {
	for (var i=0;i<range.values.length;i++){
		for (var j=0;j<range.values[i].length;j++){
			if(range.values[i][j]%2==0){
				range.getCell(i,j).format.fill.color="FF0000";
			}
			else{
				range.getCell(i,j).format.fill.color="008000";
			}
		}
	}
	ctx.executeAsync().then(function(){
		ctx.references.remove(range);
		ctx.executeAsync().then(function()
		{
			
		});
	});
});