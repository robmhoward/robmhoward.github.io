var ctx = new Excel.ExcelClientContext();
ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1").title.text="New Title";
ctx.executeAsync().then();