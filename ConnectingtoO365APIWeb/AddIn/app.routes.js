(function () {
    'use strict';

    var outlookApp = angular.module('graphapp');

    // load routes
    outlookApp.config(['$routeProvider', routeConfigurator]);

    function routeConfigurator($routeProvider) {
        $routeProvider
            .when('/home', {
                templateUrl: 'views/home.html',
                controller: 'homeController'
            });


        $routeProvider.otherwise({ redirectTo: '/home' });
    }
})();
