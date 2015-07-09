var ctx = new Excel.RequestContext();
var range = ctx.workbook.tables.getItem('Table1').tableRows.getItemAt(3).getRange();
range.format.background.color = "#00AA00";
ctx.executeAsync().then();