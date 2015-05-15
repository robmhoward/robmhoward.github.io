var excelSamplesApp = angular.module("excelSamplesApp", ['ngRoute']);

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
	
	MonacoEditorIntegration.initializeJsEditor('TxtRichApiScript', [
			"/foo/script/EditorIntelliSense/Excel.txt",
			"/foo/script/EditorIntelliSense/Office.Runtime.txt",
			"/foo/script/EditorIntelliSense/Helpers.txt",
			"/foo/script/EditorIntelliSense/jquery.txt",
		]);
	
	excelSamplesFactory.getSamples().then(function (response) {
		$scope.samples = response.data.values;
		$scope.groups = response.data.groups;
	});

	$scope.loadSampleCode = function() {
		console.log("loadSampleCode called");
		excelSamplesFactory.getSampleCode($scope.selectedSample.filename).then(function (response) {
			$scope.selectedSample.code = response.data;
		});
	};

});