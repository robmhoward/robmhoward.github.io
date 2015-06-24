var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(0);	
chart.series.getItemAt(0).fillFormat.setSolidColor("#FF0000");
ctx.executeAsync().then();