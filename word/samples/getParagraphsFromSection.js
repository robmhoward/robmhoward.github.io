var ctx = new Word.WordClientContext();
ctx.customData = OfficeExtension.Constants.iterativeExecutor;

var paras = ctx.document.sections.getItemAt(0).body.paragraphs;
ctx.load(paras);

ctx.executeAsync().then(
    function () {
        console.log("Number of paragraphs in section: " + paras.count);
    },
    function (result) {
        console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
        console.log(result.traceMessages);
    }
);