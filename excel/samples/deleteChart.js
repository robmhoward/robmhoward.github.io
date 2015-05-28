var ctx = new Excel.ExcelClientContext();
ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").deleteObject();	
ctx.executeAsync().then();