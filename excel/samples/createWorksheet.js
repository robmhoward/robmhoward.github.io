var ctx = new Excel.ExcelClientContext();
ctx.workbook.worksheets.add("Sheet" + Math.floor(Math.random()*100000).toString());
ctx.executeAsync().then();