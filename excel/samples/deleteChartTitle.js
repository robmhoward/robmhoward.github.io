var ctx = new Excel.ExcelClientContext();
ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").title.visible = false; 
ctx.executeAsync().then();