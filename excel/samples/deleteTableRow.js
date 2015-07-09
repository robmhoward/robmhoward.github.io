var ctx = new Excel.RequestContext();
ctx.workbook.tables.getItem('Table1').tableRows.getItemAt(3).deleteObject();
ctx.executeAsync().then();