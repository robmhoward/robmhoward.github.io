var ctx = new Word.WordClientContext();
var par = ctx.document.body.paragraphs.getItemAt(0);
par.setLineSpacing(36);
var val = par.getLineSpacing();

ctx.executeAsync().then(
    function () {
        console.log("Success! Setting paragraph line spacing to " + val.value);
    },
    function (result) {
        console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
        console.log(result.traceMessages);
    }
);
