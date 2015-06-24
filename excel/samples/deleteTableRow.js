var ctx = new Excel.ExcelClientContext();
ctx.workbook.tables.getItem('Table1').tableRows.getItemAt(3).deleteObject();
ctx.executeAsync().then();