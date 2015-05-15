var ctx = new Excel.ExcelClientContext();
ctx.workbook.tables.add('MyTable', 'Sheet1!A1:E7', true, false, null);
ctx.executeAsync().then();