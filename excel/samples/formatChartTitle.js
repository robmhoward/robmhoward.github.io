var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(0);	

chart.title.font.bold = true; 
chart.title.font.color = "#FF0000";

ctx.executeAsync().then();
