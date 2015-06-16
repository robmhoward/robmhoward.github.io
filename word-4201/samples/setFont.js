var ctx = new Word.WordClientContext();
var para = ctx.document.body.paragraphs.getItemAt(0);
var font = para.font;

font.size = 32;
font.bold = true;
font.color = "#0000ff";
font.highlightColor = "#ffff00";

ctx.executeAsync().then(
    function () {
        console.log("Success");
    },
    function (result) {
        console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
        console.log(result.traceMessages);
    }
);
