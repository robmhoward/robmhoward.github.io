var ctx = new Excel.ExcelClientContext();
ctx.workbook.getActiveWorksheet().getRange("A1:C3").numberFormat = "d-mmm";
ctx.executeAsync();