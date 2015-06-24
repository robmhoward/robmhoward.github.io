var ctx = new Excel.ExcelClientContext();
ctx.workbook.worksheets.getActiveWorksheet().deleteObject();
ctx.executeAsync().then();