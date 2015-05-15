var ctx = new Excel.ExcelClientContext();
var range = ctx.workbook.worksheets.getItem("Sheet1").getRange("A1:B2");
range.formulas = [["12345", "=A1"], ["=B1", "=RAND()"]];
ctx.executeAsync().then();