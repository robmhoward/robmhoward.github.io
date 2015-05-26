var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	
chart.fillFormat.setSolidColor("#FF0000");
ctx.executeAsync().then();