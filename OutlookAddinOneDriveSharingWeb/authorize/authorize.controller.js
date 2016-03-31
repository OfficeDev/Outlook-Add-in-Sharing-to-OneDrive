/*
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
*  See LICENSE in the source repository root for complete license information.
*/

(function () {
    'use strict';

    $(document).ready(function () {

        //Checks the local storage for a token.
        var token = getToken(window.location.href);
        if (token !== null) {
            window.localStorage.setItem('token', JSON.stringify(token));
            window.localStorage.setItem('dump', 'stuff' + Date().toString());
            $('#success').show();
            $('#failed').hide();
        }
        else {
            $('#success').hide();
            $('#failed').show();
        }
    });

    function getToken(hash) {
        var parts = hash.split('#');
        if (parts == null || parts.length <= 0) return null;

        var rightPart = parts.length == 2 ? parts[1] : parts[0];
        var token = getTokenFromString(rightPart);
        return token;
    }

    function getTokenFromString(hash) {
        var params = {},
            regex = /([^&=]+)=([^&]*)/g,
            matches;

        while ((matches = regex.exec(hash)) != null) {
            params[decodeURIComponent(matches[1])] = decodeURIComponent(matches[2]);
        }

        if (params.access_token || params.code || params.error) {
            return params;
        }

        return null;
    }
})();