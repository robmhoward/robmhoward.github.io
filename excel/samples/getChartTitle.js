var ctx = new Excel.ExcelClientContext();
var title = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(0).title;	
ctx.load(title);
ctx.executeAsync().then(function () {
		console.log(title.text);
		console.log("done");
});