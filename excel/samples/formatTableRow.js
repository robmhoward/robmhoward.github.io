var ctx = new Excel.ExcelClientContext();
var range = ctx.workbook.tables.getItem('MyTable').tableRows.getItemAt(3).getRange();
range.format.background.color = "green";
ctx.executeAsync().then();