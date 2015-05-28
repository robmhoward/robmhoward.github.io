var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
chart.axes.valueAxis.title.text = "Category";
ctx.executeAsync().then();