var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(0);	
chart.top = 200;
chart.left = 200;
ctx.executeAsync().then();