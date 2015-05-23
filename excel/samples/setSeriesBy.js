var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	
var sourceData = "A1:B4";

chart.SetData(sourceData, "Rows");
ctx.executeAsync().then();