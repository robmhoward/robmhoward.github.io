var ctx = new Excel.ExcelClientContext();
var sheet = ctx.workbook.worksheets.getItem("Sheet1");
 
var range = sheet.getRange("A1:B3");
range.values = [
       ["", "Gender"],
       ["Male", 12],
       ["Female", 14]
];
 
var chart = sheet.charts.add("pie", range, "auto");
 
chart.fillFormat.setSolidColor("F8F8FF");
 
chart.title.text = "Class Demographics";
chart.title.font.bold = true;
chart.title.font.size = 18;
chart.title.font.color = "568568";
 
chart.legend.position = "right";
chart.legend.font.name = "Algerian";
chart.legend.font.size = 13;
 
chart.dataLabels.showPercentage = true;
chart.dataLabels.font.size = 15;
chart.dataLabels.font.color = "444444";
 
var points = chart.series.getItemAt(0).points;
points.getItemAt(0).fillFormat.setSolidColor("8FBC8F");
points.getItemAt(1).fillFormat.setSolidColor("D87093");
 
ctx.executeAsync().then();

