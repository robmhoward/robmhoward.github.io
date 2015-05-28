var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.datalabels.visible = true;

ctx.executeAsync().then(function () {
		console.log("Datalabels Shown");
});