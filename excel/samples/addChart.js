var ctx = new Excel.ExcelClientContext();
ctx.workbook.worksheets.getItem("Sheet1").charts.add("ColumnClustered", "Sheet1!A1:B4", Excel.ChartSeriesBy.auto);
ctx.executeAsync().then();
