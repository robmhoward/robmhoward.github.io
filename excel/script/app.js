var excelSamplesApp = angular.module("excelSamplesApp", ['ngRoute']);
var insideOffice = false;
var consoleErrorFunction;

var logComment = function(message) {
	var consoleElement;
	consoleElement = document.getElementById('console');
	consoleElement.innerHTML += message + '\n';
	consoleElement.scrollTop = consoleElement.scrollHeight;
}

Office.initialize = function (reason) {
	insideOffice = true;	
	console.log('Add-in initialized, redirecting console.log() to console textArea');
	consoleErrorFunction = console.error;
	console.
	console.error = logComment;
};

excelSamplesApp.config(['$routeProvider', function ($routeProvider) {
	$routeProvider
		.when('/samples',
			{
				controller: 'SamplesController',
				templateUrl: 'partials/samples.html'
			})
		.otherwise({redirectTo: '/samples' });
}]);

excelSamplesApp.factory("excelSamplesFactory", ['$http', function ($http) {
	var factory = {};
	
	factory.getSamples = function() {
		return $http.get('samples/samples.json');
	};

	factory.getSampleCode = function(filename) {
		return $http.get('samples/' + filename);
	};

	return factory;
}]);

excelSamplesApp.controller("SamplesController", function($scope, excelSamplesFactory) {
	$scope.samples = [{ name: "Loading..." }];
	$scope.selectedSample = { description: "No sample loaded" };
	$scope.insideOffice = insideOffice;
	
	MonacoEditorIntegration.initializeJsEditor('TxtRichApiScript', [
			"/excel/script/EditorIntelliSense/ExcelLatest.txt",
			"/excel/script/EditorIntelliSense/Office.Runtime.txt",
			"/excel/script/EditorIntelliSense/Helpers.txt",
			"/excel/script/EditorIntelliSense/jquery.txt",
		]);
	
	MonacoEditorIntegration.setDirty = function() {
		if ($scope.selectedSample.code) {
			$scope.selectedSample = { description: $scope.selectedSample.description + " (modified)" };
			$scope.$apply();
		}
	}
	
	excelSamplesFactory.getSamples().then(function (response) {
		$scope.samples = response.data.values;
		$scope.groups = response.data.groups;
	});

	$scope.loadSampleCode = function() {
		appInsights.trackEvent("SampleLoaded", {name:$scope.selectedSample.name});
		excelSamplesFactory.getSampleCode($scope.selectedSample.filename).then(function (response) {
			$scope.selectedSample.code = addErrorHandlingIfNeeded(response.data);
			$scope.insideOffice = insideOffice;
			MonacoEditorIntegration.setJavaScriptText($scope.selectedSample.code);
			MonacoEditorIntegration.resizeEditor();
		});
	};
	
	$scope.runSelectedSample = function() {
		var script = MonacoEditorIntegration.getJavaScriptToRun().replace("console.log", "logComment");
		try {
			eval(script);
		} catch (e) {
			logComment(e.name + ": " + e.message);
		}
	}

});

function addErrorHandlingIfNeeded(sampleCode) {
	if (!insideOffice) return sampleCode;
	return sampleCode.replace("ctx.executeAsync().then();", "ctx.executeAsync().then(function() {\r\n    console.log(\"done\");\r\n}, function(error) {\r\n    console.log(\"An error occurred: \" + error.name + \":\" + error.message);\r\n});");
}