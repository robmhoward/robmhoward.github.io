var ctx = new Excel.RequestContext();
var selectedRange = ctx.workbook.getSelectedRange();
ctx.load(selectedRange.worksheet);
ctx.executeAsync().then(function () {
    console.log(selectedRange.worksheet.name);
}, function (error) {
    console.log("An error occurred: " + error.errorCode + ":" + error.errorMessage);
});