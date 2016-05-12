(function () {
    //'use strict';

    angular.module('graphapp')
        .controller('homeController',
        ['$scope', '$rootScope', '$window', 'graphService', 'signalRService', homeController]
    );

    function homeController($scope, $rootScope, $window, graphService, signalRService) {

        $scope.authenticated = false;

        $scope.sheetsAdded = false;

        $scope.getPeople = function () {
            if ($scope.authenticated == false) {

                try {
                    signalRService.startOAuthFlow(getPeople);
                }
                catch (e) {
                    var msg = e;
                }
            } else {
                getPeople();
            }
        };

        $scope.getMeProfile = function () {
            if ($scope.authenticated == false) {
                signalRService.startOAuthFlow(getMeProfile);
            } else {
                getMeProfile();
            }
        };

        function getMeProfile() {
            $scope.authenticated = true;
            graphService.getMeProfile($rootScope.connectionId);
        }

        function getPeople() {
            $scope.authenticated = true;
            graphService.getPeople($rootScope.connectionId);
        }

        function addSheets() {

            if ($scope.sheetsAdded == false) {

                Excel.run(function (ctx) {
                    ctx.workbook.worksheets.add('Me');
                    ctx.workbook.worksheets.add('People');
                    return ctx.sync();
                })
                .then(function () {
                    $scope.sheetsAdded = true;
                })
            }
        }

        $scope.$on("getPeopleComplete", function (event) {

            addSheets();

            $scope.people = $rootScope.people;

            Excel.run(function (ctx) {
                var sheet = ctx.workbook.worksheets.getItem('People');

                for (index = 0, len = $scope.people.value.length; index < len; ++index) {
                    var address = "A" + (index + 1).toString();
                    var range = sheet.getRange(address);
                    range.values = $scope.people.value[index].displayName;
                }

                sheet.activate();
                return ctx.sync();
            })
            .then(function () {
                // todo
            })
            .catch(function (er) {
                // todo
            });
        });

        $scope.$on("getMeComplete", function (event) {

            addSheets();

            $scope.meProfile = $rootScope.meProfile;

            Excel.run(function (ctx) {
                var sheet = ctx.workbook.worksheets.getItem('Me');

                var range = sheet.getRange("A1");
                range.values = "Name";
                range = sheet.getRange("B1")
                range.values = $scope.meProfile.givenName + ' ' + $scope.meProfile.surname;

                range = sheet.getRange("A2");
                range.values = "Email";
                range = sheet.getRange("B2")
                range.values = $scope.meProfile.mail;

                range = sheet.getRange("A3");
                range.values = "Mobile";
                range = sheet.getRange("B3")
                range.values = $scope.meProfile.mobilePhone;

                range = sheet.getRange("A4");
                range.values = "UPN";
                range = sheet.getRange("B4")
                range.values = $scope.meProfile.userPrincipalName;

                var colRange = sheet.getUsedRange().getColumn(0);
                colRange.format.fill.color = "yellow";

                sheet.activate();
                return ctx.sync();
            })
            .then(function () {
                // todo
            })
            .catch(function (er) {
                // todo
            });

        });

        $scope.$on("signalRConnected", function (event, args) {

            $rootScope.connectionId = args.hubConnectionId;

            graphService.getAuthorizationUrl(function (data) {
                $window.open(data);
            }, $rootScope.connectionId);
        });
    }
})();
