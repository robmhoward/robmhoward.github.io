var ctx = new Word.WordClientContext();
var range = ctx.document.selection;

var myContentControl = range.insertContentControl();

ctx.load(myContentControl);

ctx.executeAsync().then(
     function () {
         console.log("Content control Id: " + myContentControl.id);
     },
     function (result) {
         console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
         console.log(result.traceMessages);
     }
);
