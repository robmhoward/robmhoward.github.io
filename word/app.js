var wordSamplesApp = angular.module("wordSamplesApp", ['ngRoute']);
var insideOffice = false;

var logComment = function(message) {
	document.getElementById('console').innerHTML += message + '\n';
}

Office.initialize = function (reason) {
	insideOffice = true;	
	console.log('Initialized!');
};

wordSamplesApp.config(['$routeProvider', function ($routeProvider) {
	$routeProvider
		.when('/samples',
			{
				controller: 'SamplesController',
				templateUrl: 'partials/samples.html'
			})
		.otherwise({redirectTo: '/samples' });
}]);

wordSamplesApp.factory("wordSamplesFactory", ['$http', function ($http) {
	var factory = {};
	
	factory.getSamples = function() {
		return $http.get('samples/samples.json');
	};

	factory.getSampleCode = function(filename) {
		return $http.get('samples/' + filename);
	};

	return factory;
}]);

wordSamplesApp.controller("SamplesController", function($scope, wordSamplesFactory) {
	$scope.samples = [{ name: "Loading..." }];
	$scope.selectedSample = { description: "No sample loaded" };
	$scope.insideOffice = insideOffice;
	
	MonacoEditorIntegration.initializeJsEditor('TxtRichApiScript', [
			"/word/script/EditorIntelliSense/WordLatest.txt",
			"/word/script/EditorIntelliSense/Office.Runtime.txt",
			"/word/script/EditorIntelliSense/Helpers.txt",
			"/word/script/EditorIntelliSense/jquery.txt",
		]);
	
	MonacoEditorIntegration.setDirty = function() {
		if ($scope.selectedSample.code) {
			$scope.selectedSample = { description: $scope.selectedSample.description + " (modified)" };
			$scope.$apply();
		}
	}
	
	wordSamplesFactory.getSamples().then(function (response) {
		$scope.samples = response.data.values;
		$scope.groups = response.data.groups;
	});

	$scope.loadSampleCode = function() {
		console.log("loadSampleCode called");
		appInsights.trackEvent("SampleLoaded", {name:$scope.selectedSample.name});
		wordSamplesFactory.getSampleCode($scope.selectedSample.filename).then(function (response) {
			$scope.selectedSample.code = response.data;
			$scope.insideOffice = insideOffice;
			MonacoEditorIntegration.setJavaScriptText($scope.selectedSample.code);
		});
	};
	
	$scope.runSelectedSample = function() {
		var script = MonacoEditorIntegration.getJavaScriptToRun().replace("console.log", "logComment");
		eval(script);
	}
});