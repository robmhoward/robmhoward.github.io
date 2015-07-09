var ctx = new Excel.RequestContext();
ctx.workbook.tables.add('Table1', 'Sheet1!A1:E7', true, false, null);
ctx.executeAsync().then();
