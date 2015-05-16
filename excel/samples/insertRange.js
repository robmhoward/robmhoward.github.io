var ctx = new Excel.ExcelClientContext();
ctx.workbook.worksheets.getItem("Sheet1").getRange("A1:C3").insert("right");
ctx.executeAsync().then();