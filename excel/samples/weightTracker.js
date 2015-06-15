//prepare data for weight tracker
var weightvalues=[["Average",""],["Maximum",""],["Week","Weight(lb)"],[1,125],[2,135],[3,145],[4,148],[5,151],[6,142],[7,145],[8,143],[9,145],[10,146]];

//define each area of the weight tracker. The first two rows will display average weight and maximum weight. From the 3rd row, it would be a table with weekly weight data.
var sheetName="Weight";
var rangeAddress=sheetName + "!" + "A1:B13";
var rangeValues=sheetName + "!" + "B3:B13";
var tableaddress=sheetName + "!" + "A3:B13";

//create table
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var table=ctx.workbook.tables.add("WeightByWeek",tableaddress, true);
var range = worksheet.getRange(rangeAddress);
var rangeMax = worksheet.getRange("B2:B2");
var rangeAvg = worksheet.getRange("B1:B1");

//fill data
range.values=weightvalues;

//use formula to calculate maximum and average of weight
rangeMax.formulas= "=MAX("+rangeValues+")";
rangeAvg.formulas= "=AVERAGE("+rangeValues+")";
rangeMax.format.font.bold=true;
rangeAvg.format.font.bold=true;
rangeMax.format.font.size=16;
rangeAvg.format.font.size=16;
range.getCell(0,0).format.font.size=16;
range.getCell(1,0).format.font.size=16;
range.getCell(0,0).format.font.bold=true;
range.getCell(1,0).format.font.bold=true;
ctx.load(range);
ctx.load(rangeAvg);
ctx.references.add(range);
ctx.executeAsync().then(function(){
//if the weight is more than average, highlight the cell in red.
	for(var i=4; i<range.values.length;i++){
		if(range.values[i][1]>rangeAvg.values[0][0]){
			range.getCell(i,1).format.fill.color="FF0000";
		}
	}

	ctx.executeAsync().then(function(){
		ctx.references.remove(range);
		ctx.executeAsync().then(function()
		{
			
		});
	});
});

//create a chart to present the data
var chart=worksheet.charts.add("ColumnClustered",rangeValues,"auto");

//update chart tile and change the color to green
chart.title.text="Weight Tracker";
chart.title.format.font.name="Impact";
chart.title.format.font.color="green";

//update the axis title and format
chart.axes.categoryAxis.title.text="Week";
chart.axes.valueAxis.title.text="Weight";
chart.axes.valueAxis.title.format.font.color="7F00FF";
chart.axes.valueAxis.title.format.font.color="7F00FF";
//show data from 120 as we don't have any weight below 120 and also set the chart to display datalabels
chart.axes.valueAxis.minimum=120;
chart.dataLabels.showValue=true;

ctx.executeAsync().then(function(){
});

