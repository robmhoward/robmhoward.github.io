var ctx = new Word.WordClientContext();
var paras = ctx.document.body.paragraphs;
ctx.load(paras);

var pic = paras.getItem(0).insertInlinePictureFromUrl("http://dev.office.com/Media/Default/App%20Awards/AppAwards.png", Word.InsertLocation.end, false, true);
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