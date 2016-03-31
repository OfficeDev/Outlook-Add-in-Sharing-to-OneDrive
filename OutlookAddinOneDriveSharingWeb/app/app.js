/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

/// <reference path="authenticate.js" />
/// <reference path="onedrive.object.js" />

'use strict';

(function () {
    // Global objects to use each time the app is run.
    var authenticationContext,
        shareContexts,
        renderingContext;

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            authenticationContext = new AuthenticationContext(TOKEN);
            renderingContext = new Renderer();
            authenticationContext.authenticate()
                .then(setOneDriveToken)
                .then(fetchLinksAndRecipients)
                .spread(checkLinksForPermissions)                
                .finally(renderUIElements)
                .catch(handleError);
        });
    }

    /*
    * Sets the token to make OneDrive API calls.
    */
    function setOneDriveToken(token) {
        if (!token) {
            OneDriveShare.token = null;
            throw 'Authentication failed';
        }

        OneDriveShare.token = token;
        return token;
    }

    /*
    * Get all links and recipients from the compose message.
    */
    function fetchLinksAndRecipients(token) {
        return Q.all([
            getOneDriveLinks(),
            getAllRecipients()
        ]);
    }

    /*
    * Checks all links in the message body for OneDrive links.
    */
    function getOneDriveLinks() {
        var deferred = Q.defer();

        Office.context.mailbox.item.body.getAsync("html", function (asyncResult) {
            try {
                var ONEDRIVE_URL_FORMAT = "^https:\/\/(onedrive\.live\.com|1drv\.ms)",
                    htmlParser = new DOMParser().parseFromString(asyncResult.value, "text/html"),
                    links = htmlParser.getElementsByTagName("a");

                var oneDriveLinks = $.map(links, function (link) {
                    var innerText = link.innerText.toLowerCase().trim();
                    var urlString = link.href.toLowerCase().trim();
                    if (urlString.search(ONEDRIVE_URL_FORMAT) !== -1) {
                        return urlString;
                    }
                });

                deferred.resolve(oneDriveLinks);
            }
            catch (e) {
                deferred.reject(e);
            }
        });

        return deferred.promise;
    }

    // Get an array of recipients from the compose message.
    function getAllRecipients() {
        // Local objects to point to recipients of the message that is being composed.
        var self = this;
        var toRecipients = Office.context.mailbox.item.to;
        var ccRecipients = Office.context.mailbox.item.cc;
        var bccRecipients = Office.context.mailbox.item.bcc;

       

        // Get all recipients' email addresses, including To, Cc, and Bcc recipients.
        return Q.all([
            getRecipientsFromSource(toRecipients),
            getRecipientsFromSource(ccRecipients)
        ]).then(function (sources) {
            var allRecipients = {};

            $.each(sources, function (index, source) { 
                $.each(source, function (name, properties) {
                    allRecipients[name] = properties;
                });
            })

            if (Object.keys(allRecipients).length == 0) throw "We couldn't find any recipients."

            return allRecipients;
        });
    }

    // Get the recipients from an Outlook email message.
    function getRecipientsFromSource(mailboxItem) {
        var deferred = Q.defer(),
            recipients = {};

        // Use asynchronous method getAsync to get each recipient
        // of the composed message. Each time, this example passes an anonymous 
        // callback function that doesn't take any parameters.
        mailboxItem.getAsync(function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                deferred.reject(asyncResult.error.message);
            }

            //Insert this recipient into recipients array.
            for (var i = 0; i < asyncResult.value.length; i++) {
                recipients[asyncResult.value[i].emailAddress] = {
                    "hasPermission": false,
                    "role": null
                };
            }

            deferred.resolve(recipients);
        });

        return deferred.promise;
    }

    
    // Check the link to see if its shared with any recipients.
    function checkLinksForPermissions(links, recipients) {
        if (!links) throw "We couldn't find any links";

        shareContexts = links.map(function (link) {
            if (link != null) {
                var shareContext = new OneDriveShare(link, recipients);
                return shareContext;
            }
        });
    }

    
    // Renders the display.
    function renderUIElements() {
        if (!shareContexts || shareContexts.length == 0) throw "We couldn't find any links or people."

        var currentView = shareContexts[0];

        // TODO: Add a dropdown selector to switch currentView.

        renderingContext.clearUI();
        return renderingContext.renderUI(currentView);
    }

    function handleError(error) {
        console.error(error);
    }
})();

// *********************************************************
//
// Outlook-Add-in-Sharing-to-OneDrive
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************