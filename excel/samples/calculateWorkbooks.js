var ctx = new Excel.ExcelClientContext();
ctx.workbook.application.calculate(Excel.CalculationType.full);
ctx.executeAsync().then();