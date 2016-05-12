(function () {
    angular.module('graphapp')
        .service('graphService', ['$http', '$rootScope', graphService]);

    function graphService($http, $rootScope) {
        return {
            getPeople: getPeople,
            getMeProfile: getMeProfile,
            getAuthorizationUrl: getAuthorizationUrl
        };

        function getPeople(connectionId) {
            $http({
                method: "GET",
                url: '../../api/Graph/GetPeople?connectionId=' + connectionId,
                headers: {}
            }).success(function (data) {
                $rootScope.people = data;
                $rootScope.$broadcast("getPeopleComplete");
            }).error(function (er) {
                var msg = er;
            });
        };

        function getMeProfile(connectionId) {
            $http({
                method: "GET",
                url: '../../api/Graph/GetMeProfile?connectionId=' + connectionId,
                headers: {}
            }).success(function (data) {
                $rootScope.meProfile = data;
                $rootScope.$broadcast("getMeComplete");
            }).error(function (er) {
                var msg = er;
            });
        };

        function getAuthorizationUrl(callback, connectionId) {

            $http({
                method: "GET",
                url: '../../api/OAuth/GetAuthorizationUrl?connectionId=' + connectionId,
                headers: {}
            }).success(function (data) {
                callback(data);
            }).error(function (er) {
            });
        }
    };

})();