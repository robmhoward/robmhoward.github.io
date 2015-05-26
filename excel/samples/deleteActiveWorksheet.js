var ctx = new Excel.ExcelClientContext();
ctx.workbook.getActiveWorksheet().deleteObject();
ctx.executeAsync().then();