var ctx = new Excel.ExcelClientContext();
ctx.workbook.tables.getItem("Table1").getDataBodyRange().clear(Excel.ClearApplyTo.formats);
ctx.executeAsync().then();