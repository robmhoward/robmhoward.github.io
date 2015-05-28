var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
chart.height = 200;
chart.width = 200;
ctx.executeAsync().then();