/*
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
*  See LICENSE in the source repository root for complete license information.
*/

'use strict';

var AuthenticationContext = (function () {
    function AuthenticationContext(token) {
        this.token = token || null;
    }

    AuthenticationContext.prototype.isAuthenticated = function () {
        if (this.token == null) {
            var token = window.localStorage.getItem('token');
            this.token = token != null ? JSON.parse(token) : null;
        }

        return this.token !== null;
    }

    AuthenticationContext.prototype.authenticate = function (force) {
        var self = this,
            deferred = Q.defer();

        if (!force && this.isAuthenticated()) {
            deferred.resolve(this.token);
            return deferred.promise;
        }

        try {
            // Construct a request to send to the service. It will look something like this:
            // <http endpoint>/authorize?client_id="<client id>"&response_type=token&redirect_uri="<encoded redirect uri>"
            // &scope="<encoded scope>"
            var authUrl = AUTH_ENDPOINT
                + "/authorize?"
                + "client_id=" + CLIENT_ID
                + "&response_type=token"
                + "&redirect_uri=" + encodeURIComponent(REDIRECT_URI)
                + "&scope=" + encodeURIComponent(SCOPES)
                + "&response_mode=fragment"
                + "&state=12345"
                + "&nonce=23232432465433";

            Office.context.ui.displayDialogAsync(authUrl, {
                height: 60,
                width: 40,
                enforceAppDomains: true,
                requireHTTPS: true
            }, function (result) {
                try {
                    var dialog = result.value;
                    dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, function (event) {
                        if (event.error === 12002) {
                            deferred.reject(null);
                        }
                        else {
                            deferred.notify(event);
                        }
                    });
                    dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogEventReceived, function (event) {
                        if (event.error === 12006) {
                            if (self.isAuthenticated()) {
                                deferred.resolve(self.token);
                            }
                            else {
                                deferred.reject(null);
                            }

                        };
                    });
                }
                catch (error) {
                    deferred.reject(error);
                }
            });
        }
        catch (error) {
            deferred.reject(error);
        }

        return deferred.promise;
    }

    AuthenticationContext.prototype.clearTokens = function () {
        this.token = null;
        window.localStorage.removeItem('token');
    }

    return AuthenticationContext;
})();
