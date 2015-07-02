var ctx = new Word.WordClientContext();
var paras = ctx.document.body.paragraphs;
ctx.load(paras);
ctx.references.add(paras);

ctx.executeAsync().then(
    function () {
        var results = new Array();
        for (var i = 0; i < paras.items.length; i++) {
            results.push(paras.getItem(i).getText());
        }
        ctx.executeAsync().then(
            function () {
                for (var i = 0; i < results.length; i++) {
                    console.log("paras[" + i + "].content  = " + results[i].value);
                }
            }
        );
    },
    function (result) {
        console.log("Failed: ErrorCode=" + result.errorCode + ", ErrorMessage=" + result.errorMessage);
        console.log(result.traceMessages);
    }
);