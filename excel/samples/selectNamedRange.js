var ctx = new Excel.ExcelClientContext();
ctx.workbook.names.getItem("myData").getRange().select();
ctx.executeAsync().then();