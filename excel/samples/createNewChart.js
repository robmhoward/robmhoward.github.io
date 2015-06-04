var ctx = new Excel.ExcelClientContext();
var range = ctx.workbook.names.getItem("MyChartData").getRange();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.add("pie", range, "auto");
ctx.executeAsync();