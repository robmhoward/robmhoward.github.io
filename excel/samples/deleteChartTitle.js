var ctx = new Excel.ExcelClientContext();
ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(0).title.visible = false; 
ctx.executeAsync().then();