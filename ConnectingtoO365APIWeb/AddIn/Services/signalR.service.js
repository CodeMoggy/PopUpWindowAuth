(function () {
    angular.module('graphapp')
        .service('signalRService', ['$rootScope', signalRService]);
            
    function signalRService ($rootScope) {

        return {
            startOAuthFlow: startOAuthFlow
        };

        var proxy = null;        
        
        function startOAuthFlow (callback) {         
            //Getting the connection object         
            connection = $.hubConnection();

            //Creating proxy         
            this.proxy = connection.createHubProxy('oauth');

            //publish an event when oauth is complete         
            this.proxy.on('completed', function () {
                callback();
            });

            //publish an event when connected to the server         
            this.proxy.on('connected', function (hubConnectionId) {
                $rootScope.$broadcast("signalRConnected", { hubConnectionId: hubConnectionId });
            });

            //Starting connection         
            connection.start();
        };        
    };

})();
