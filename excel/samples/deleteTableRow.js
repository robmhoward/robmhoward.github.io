var ctx = new Excel.ExcelClientContext();
ctx.workbook.tables.getItem('MyTable').tableRows.getItemAt(3).deleteObject();
ctx.executeAsync().then();