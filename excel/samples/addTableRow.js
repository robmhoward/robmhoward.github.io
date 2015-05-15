var ctx = new Excel.ExcelClientContext();
var tableRows = ctx.workbook.tables.getItem('MyTable').tableRows;
tableRows.add(3, [[1,2,3,4,5]]);
ctx.executeAsync().then();