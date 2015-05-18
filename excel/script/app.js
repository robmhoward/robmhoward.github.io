var excelSamplesApp = angular.module("excelSamplesApp", ['ngRoute']);
var insideOffice = false;

console.log = function(message) {
	document.getElementById('console').innerHTML += message + '\n';
}

Office.initialize = function (reason) {
	insideOffice = true;	
	console.log('Initialized!');
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
	$scope.insideOffice = insideOffice;
	
	MonacoEditorIntegration.initializeJsEditor('TxtRichApiScript', [
			"/excel/script/EditorIntelliSense/Excel.txt",
			"/excel/script/EditorIntelliSense/Office.Runtime.txt",
			"/excel/script/EditorIntelliSense/Helpers.txt",
			"/excel/script/EditorIntelliSense/jquery.txt",
		]);
	
	excelSamplesFactory.getSamples().then(function (response) {
		$scope.samples = response.data.values;
		$scope.groups = response.data.groups;
	});

	$scope.loadSampleCode = function() {
		console.log("loadSampleCode called");
		appInsights.trackEvent("SampleLoaded", {name:$scope.selectedSample.name});
		excelSamplesFactory.getSampleCode($scope.selectedSample.filename).then(function (response) {
			$scope.selectedSample.code = response.data;
			$scope.insideOffice = insideOffice;
			MonacoEditorIntegration.setJavaScriptText($scope.selectedSample.code);
		});
	};
	
	$scope.runSelectedSample = function() {
		var script = MonacoEditorIntegration.getJavaScriptToRun();
		eval(script);
	}

});