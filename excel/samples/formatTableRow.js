var ctx = new Excel.ExcelClientContext();
var range = ctx.workbook.tables.getItem('MyTable').tableRows.getItemAt(3).getRange();
range.format.background.color = "#00AA00";
ctx.executeAsync().then();