var wordSamplesApp = angular.module("wordSamplesApp", ['ngRoute']);
var insideOffice = false;
var debugOption = false;

var logComment = function (message) {
    var span = document.createElement('span');
    span.className = 'message-text';
    span.innerHTML = message + '<br/>';
    $('#console').append(span);
}

var logDebug = function (message) {
    if (debugOption) {
        var span = document.createElement('span');
        span.className = 'debug-text';
        span.innerHTML = message + '<br/>';
        $('#console').append(span);
    }
}

Office.initialize = function (reason) {
    insideOffice = true;

    // Override window.console to log framework debug info
    window.console.log = function (message) {
        logDebug(message);
    };

    // Log all unhandled exceptions
    window.onerror = function (em, url, ln) {
        logDebug("OnError: " + em + ", " + url + ", " + ln);
    };

    console.log('Initialized!');
};

wordSamplesApp.config(['$routeProvider', function ($routeProvider) {
    $routeProvider
        .when('/samples',
            {
                controller: 'SamplesController',
                templateUrl: 'partials/samples.html'
            })
        .otherwise({ redirectTo: '/samples' });
}]);

wordSamplesApp.factory("wordSamplesFactory", ['$http', function ($http) {
    var factory = {};

    factory.getSamples = function () {
        return $http.get('samples/samples.json');
    };

    factory.getSampleCode = function (filename) {
        return $http.get('samples/' + filename);
    };

    return factory;
}]);

wordSamplesApp.controller("SamplesController", function ($scope, wordSamplesFactory) {
    $scope.samples = [{ name: "Loading..." }];
    $scope.selectedSample = { description: "No sample loaded" };
    $scope.debugOption = { value: false };
    $scope.insideOffice = insideOffice;

    // Update to full path if word is not at the root folder
    MonacoEditorIntegration.initializeJsEditor('TxtRichApiScript', [
            "/word-4201/script/EditorIntelliSense/WordLatest.txt",
            "/word-4201/script/EditorIntelliSense/Office.Runtime.txt",
            "/word-4201/script/EditorIntelliSense/Helpers.txt",
            "/word-4201/script/EditorIntelliSense/jquery.txt",
    ]);

    MonacoEditorIntegration.setDirty = function () {
        if ($scope.selectedSample.code) {
            $scope.selectedSample = { description: $scope.selectedSample.description + " (modified)" };
            $scope.$apply();
        }
    }

    wordSamplesFactory.getSamples().then(function (response) {
        $scope.samples = response.data.values;
        $scope.groups = response.data.groups;
    });

    $scope.loadSampleCode = function () {
        console.log("loadSampleCode called");
        appInsights.trackEvent("SampleLoaded", { name: $scope.selectedSample.name });
        wordSamplesFactory.getSampleCode($scope.selectedSample.filename).then(function (response) {
            $scope.selectedSample.code = response.data;
            $scope.insideOffice = insideOffice;
            MonacoEditorIntegration.setJavaScriptText($scope.selectedSample.code);
        });
    };

    $scope.runSelectedSample = function () {
        var script = MonacoEditorIntegration.getJavaScriptToRun().replace("console.log", "logComment");
        script = "try {" + script + "} catch(e) { logComment(\"Exception: \" + e.message ? e.message : e);}";

        logComment("====="); // Add separators between executions
        eval(script);
    }

    $scope.toggleDebugOption = function () {
        debugOption = $scope.debugOption.value;
    }

    $scope.clearLog = function () {
        $('#console').empty();
    }
});
