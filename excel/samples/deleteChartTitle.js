var ctx = new Excel.ExcelClientContext();
ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1").title.visible = false; 
ctx.executeAsync().then();