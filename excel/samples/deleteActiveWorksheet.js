var ctx = new Excel.RequestContext();
ctx.workbook.worksheets.getActiveWorksheet().deleteObject();
ctx.executeAsync().then();