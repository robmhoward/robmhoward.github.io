var MonacoEditorIntegration;
(function (MonacoEditorIntegration) {
    var localStorageKey = 'word-api-samples-4201';
    var jsMonacoEditor;

    function initializeJsEditor(textAreaId, intellisensePaths) {
        var defaultJsText = '';
        if (window.localStorage && (localStorageKey in window.localStorage)) {
            defaultJsText = window.localStorage[localStorageKey];
        }

        var editorMode = 'text/typescript';
        jsMonacoEditor = Monaco.Editor.create(document.getElementById(textAreaId), {
            value: defaultJsText,
            mode: editorMode,
            wrappingColumn: 0,
            tabSize: 4,
            lineNumbers: false,
            insertSpaces: false
        });
        document.getElementById(textAreaId).addEventListener('keyup', function () {
            storeCurrentJSBuffer();
        });

        if (window.parent.document.location.protocol == "file:") {
            intellisensePaths = [];
        } else {
            intellisensePaths = intellisensePaths.map(function (path) {
                if (path.indexOf("?") < 0) {
                    path += '?';
                } else {
                    path += '&';
                }
                return path += 'refresh=' + Math.floor(Math.random() * 1000000000);
            });
        }

        require([
            'vs/platform/platform',
            'vs/editor/modes/modesExtensions'
        ], function (Platform, ModesExt) {
            Platform.Registry.as(ModesExt.Extensions.EditorModes).configureMode(editorMode, {
                "validationSettings": {
                    "extraLibs": intellisensePaths
                }
            });
        });

        $(window).resize(function () {
            resizeEditor();
        });
    }
    MonacoEditorIntegration.initializeJsEditor = initializeJsEditor;

    function getJavaScriptToRun() {
        return jsMonacoEditor.getValue();
    }
    MonacoEditorIntegration.getJavaScriptToRun = getJavaScriptToRun;

    function setJavaScriptText(text) {
        require(["vs/editor/contrib/snippet/snippet"], function (snippet) {
            jsMonacoEditor.setSelection(jsMonacoEditor.getModel().getFullModelRange(), false);
            snippet.InsertSnippetHelper.run(jsMonacoEditor, jsMonacoEditor.getHandlerService(), new snippet.CodeSnippet(text));
            jsMonacoEditor.setSelection({ startColumn: 0, endColumn: 0, startLineNumber: 0, endLineNumber: 0 }, true);
            jsMonacoEditor.focus();
        });
    }
    MonacoEditorIntegration.setJavaScriptText = setJavaScriptText;

    function resizeEditor(scrollUp) {
        if (typeof scrollUp === "undefined") { scrollUp = false; }
        jsMonacoEditor.layout();
        if (scrollUp) {
            jsMonacoEditor.setScrollTop(0);
            jsMonacoEditor.setScrollLeft(0);
        }
        jsMonacoEditor.focus();
    }
    MonacoEditorIntegration.resizeEditor = resizeEditor;

    function storeCurrentJSBuffer() {
        console.log("storeCurrentJSBuffer");
        if (MonacoEditorIntegration.setDirty) {
            MonacoEditorIntegration.setDirty();
        }
        if (window.localStorage) {
            window.localStorage[localStorageKey] = jsMonacoEditor.getValue();
        }
    }
})(MonacoEditorIntegration || (MonacoEditorIntegration = {}));
