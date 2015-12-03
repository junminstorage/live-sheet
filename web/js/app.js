var app = angular.module('excelApp', []);

//I simply log the creation / linking of a DOM node to
//illustrate the way the DOM nodes are created with the
//various tracking approaches.
app.directive(
 "bnLogDomCreation",
 function() {
     // I bind the UI to the $scope.
     function link( $scope, element, attributes ) {
         console.log(
             attributes.bnLogDomCreation,
             $scope.$index
         );
     }
     // Return the directive configuration.
     return({
         link: link
     });
 }
);

app.config(function($locationProvider) { 
	$locationProvider.html5Mode({
	  enabled: true,
	  requireBase: false
	}); 
}); 