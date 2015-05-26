var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	
chart.dataLabels.position = Excel.ChartDataLabelPosition.top;
ctx.executeAsync().then();