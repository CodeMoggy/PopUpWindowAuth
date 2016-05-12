(function () {
    'use strict';

    // create the angular app
    var outlookApp = angular.module('graphapp', [
      'ngRoute'
    ]);

    // configure the app
    outlookApp.config(['$logProvider', function ($logProvider) {
        // set debug logging to on
        if ($logProvider.debugEnabled) {
            $logProvider.debugEnabled(true);
        }
    }]);

    //gives access to a global variable
    outlookApp.run(function ($rootScope) {
        
    });


    // when office has initalized, manually bootstrap the app
    Office.initialize = function () {
        console.log(">>> Office.initialize()");
        try {
            angular.bootstrap(jQuery('#container'), ['graphapp']);
        } catch (Exception) { }

    };

})();
