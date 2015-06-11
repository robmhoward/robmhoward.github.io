var ctx = new Excel.ExcelClientContext();
ctx.workbook.worksheets.getActiveWorksheet().getRange("A1").numberFormat = "d-mmm";
ctx.executeAsync().then();