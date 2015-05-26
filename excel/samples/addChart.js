var ctx = new Excel.ExcelClientContext();
ctx.workbook.worksheets.getItem("Charts").charts.add("ColumnClustered", "Charts!A1:B4", "auto");
ctx.executeAsync().then();