var ctx = new Excel.ExcelClientContext();
ctx.workbook.application.calculate('full');
ctx.executeAsync().then();