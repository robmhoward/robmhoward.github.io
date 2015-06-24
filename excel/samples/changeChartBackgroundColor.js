var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(0);	
chart.fillFormat.setSolidColor("#FF0000");
ctx.executeAsync().then();