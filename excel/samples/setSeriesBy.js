var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
chart.setData("A1:B4", Excel.ChartSeriesBy.rows);
ctx.executeAsync().then();