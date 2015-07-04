var ctx = new Excel.RequestContext();
var application = ctx.workbook.application;
ctx.load(application);
ctx.executeAsync().then(function() {
	console.log(application.calculationMode);
	console.log("done");
});
