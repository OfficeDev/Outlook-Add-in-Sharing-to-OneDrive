/*
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
*  See LICENSE in the source repository root for complete license information.
*/

'use strict';

var TENANT_ID = "common",
    AUTH_ENDPOINT = "https://login.microsoftonline.com/" + TENANT_ID + "/oauth2/v2.0",
    CLIENT_ID = "",
    REDIRECT_URI = "https://localhost:44300/authorize/authorize.html",
    GRAPH_ID = "https://graph.microsoft.com",
    SCOPES = "wl.signin wl.offline_access onedrive.readwrite",
    TOKEN = "";
