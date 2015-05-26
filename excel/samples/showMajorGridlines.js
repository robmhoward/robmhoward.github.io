var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	
chart.axes.valueAxis.majorGridlines.visible = true;
ctx.executeAsync().then();