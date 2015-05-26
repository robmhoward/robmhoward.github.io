var ctx = new Excel.ExcelClientContext();
ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1").name = "NewChartName";
ctx.executeAsync().then();