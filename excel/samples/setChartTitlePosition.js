var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.title.position = "top";
chart.title.overlay=true;

ctx.executeAsync().then(function () {
		console.log("Char Title Position Changed");
});