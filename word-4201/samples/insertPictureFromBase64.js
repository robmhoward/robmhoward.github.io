var ctx = new Word.WordClientContext();
ctx.customData = OfficeExtension.Constants.iterativeExecutor;
var paras = ctx.document.body.paragraphs;

var pic = paras.getItemAt(0).insertInlinePictureBase64("iVBORw0KGgoAAAANSUhEUgAAAIAAAACABAMAAAAxEHz4AAAAJFBMVEX///9GRkZGRkZGRkZGRkZGRkZGRkZGRkYBpO9/ugDyUCL/uQGm4PjWAAAACHRSTlMBCQ0RFRknMx7uViEAAAB3SURBVGje7dcxCYBQGEXhi6izYBHB0RIiiAXkzW5iAMEKFnCwguVscJd/ecM5Ab79SNHK5FqlZXeNql/XIx23awMAAAAAAAAAAAAAAAAAyBwIvzNJxeyapLZ3Naou1ykNn6sDAAAAAAAAAAAAAAAAAMgcCL9ztB/UhshWs1l/WAAAAABJRU5ErkJggg==", Word.InsertLocation.end);

var pics = ctx.document.body.inlinePictures
ctx.load(pics);

ctx.executeAsync().then(
	function () {
		console.log("Picture Count=" + pics.count);
		console.log("Success");
	},
	function (result) {
		console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
		console.log(result.traceMessages);
	}
);