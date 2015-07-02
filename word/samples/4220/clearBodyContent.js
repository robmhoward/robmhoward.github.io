var ctx = new Word.WordClientContext();
ctx.document.body.clear();

ctx.executeAsync().then(
    function () {
        console.log("Success");
    }
);
