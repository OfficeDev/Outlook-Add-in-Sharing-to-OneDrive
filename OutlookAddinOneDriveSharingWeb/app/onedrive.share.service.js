/*
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
*  See LICENSE in the source repository root for complete license information.
*/

/*
* The OneDriveShare object definition.
* @param {Object} link - A list of html links. 
* @param {Object} recipients - Recipients of the email.
* @param {Object} metadata - The properties that are returned for a OneDrive item.
* @param {String} apiBase - The OneDrive API endpoint.
*/
'use strict';

var OneDriveShare = (function () {

    // Constructor for the OneDriveShare object.
    function OneDriveShare(link, recipients) {
        this.link = link;
        this.recipients = recipients;
        this.metadata = {};
        this.apiBase = 'https://api.onedrive.com/v1.0/shares/u!';
    }

    // Access token required to make calls to the OneDrive APIs.
    OneDriveShare.token = "";

    /*
    * Base64 encodes a URL. For more information about Base64, see https://en.wikipedia.org/wiki/Base64. 
    * @param {String} data - The string to be base64 encoded.
    * @return {encodedValue} - The value of the encoded string.
    */
    OneDriveShare.base64_encode = function (data) {
        
        if (!data) {
            return data;
        }

        var encodedValue = window.btoa(data);
        // Replace '+' with '-' and '/' with '_', and remove all trailing '='.
        encodedValue = encodedValue.replace(/\+/g, '-').replace(/\//g, '_').replace(/\=+$/, '');

        return encodedValue;
    }

    /* 
    * Helper function to create requests to the OneDrive API endpoint.
    * @param {String} url - The request URL.
    * @param {String} method - The type of request. The default is GET.
    * @param {Object} data - The data for the request body.
    * @return A promise that the request has succeeded.
    */
    OneDriveShare.request = function (url, method, data) {
        var request = {
            contentType: "application/json",
            type: method || 'GET',
            url: url,
            headers: {
                "Authorization": "Bearer " + OneDriveShare.token
            },
            data: JSON.stringify(data)
        };

        if (data) request["dataType"] = "json";
        return $.ajax(request);
    }

    // Initialize and get the OneDrive item metadata and permissions list.
    OneDriveShare.prototype.initialize = function () {
        var self = this;
        return Q.all([
            self.getFileMetadata(),
            self.getPermissionsForLink()
        ]).spread(function (metadata, permissions) {
            self.metadata = metadata;
            return permissions;
        });
    }

    /*
    * Get the OneDrive item's metadata from a URL by using the shares API in OneDrive.
    * For more information about shares, see https://dev.onedrive.com/items/shares.htm
    */
    OneDriveShare.prototype.getFileMetadata = function () {
        var rootUrl = "/root",
            self = this,
            encodedUrl = OneDriveShare.base64_encode(self.link);

        var requestUrl = self.apiBase + encodedUrl + rootUrl;

        //Send request to OneDrive for the item's metadata.
        return OneDriveShare.request(requestUrl);
    }

    /* 
    * Get an item's permissions from a URL by using the shares API in OneDrive.
    * For more information about shares, see https://dev.onedrive.com/items/shares.htm
    */
    OneDriveShare.prototype.getPermissionsForLink = function () {
        var permissionsUrl = '/root/permissions',
            self = this;

        var encodedUrl = OneDriveShare.base64_encode(self.link);
        var requestUrl = self.apiBase + encodedUrl + permissionsUrl;

        //Send request to OneDrive for the item's permissions.
        return OneDriveShare.request(requestUrl)
            .then(function (response) {
                return self.checkPermissionsForLink(response);
            });
    }

    /*
    * Sets permissions for an item by using a URL and the action.invite method in the 
    * OneDrive API. For more information about action.invite, see https://dev.onedrive.com.
    * @param {Object} recipients - The list of recipients' email addresses to share the OneDrive item to.
    * @param {String} sharedPermission - The type of permission to add to the OneDrive item.
    */
    OneDriveShare.prototype.grantPermissionsToRecipients = function (recipients, sharePermission) {
        var inviteUrl = "/root/action.invite",
            self = this;

        var encodedUrl = OneDriveShare.base64_encode(self.link);
        var requestUrl = self.apiBase + encodedUrl + inviteUrl;
        var role = sharePermission || "read";

        var emailList = recipients.map(function (recipient) {
            return {
                email: recipient
            };
        });

        if (!emailList || emailList.length <= 0) {
            return null;
        }

        var newRequest = {
            requireSignIn: false,
            sendInvitation: false,
            roles: [role],
            recipients: emailList
        };

        // Send the POST request to grant a list of emails view permissions to the OneDrive item.
        return OneDriveShare.request(requestUrl, "POST", newRequest)
            .then(function (response) {
                return self.checkPermissionsForLink(response);
            });
    }

    /*
    * Modify each recipient's permissions, based on the data received.
    * Permissions are determined by the property hasPermission. If a recipient's hasPermission property is set to true, the 
    * recipient has permission type view or edit. Whether the recipient has view or edit permissions, is defined by the 
    * role object. In this method, we determine the permission type by using only the first role, roles[0].
    */
    OneDriveShare.prototype.checkPermissionsForLink = function (response) {
        var self = this,
            deferred = Q.defer(),
            permissions = response.value;

        setTimeout(function () {
            // If the permission's object is empty then the OneDrive item is not shared.
            if (permissions == null || permissions.length === 0) {
                deferred.reject({
                    permission_status: 0,
                    message: 'This file hasn\'t been shared with any of your recipients. Click \'share with all\' to share it with them.',
                    link: self.link
                });
            }
            else {
                // Find the email invitation for each permission returned in the permissions array.
                permissions.forEach(function (permission) {

                    // If the permission has no invitation then the link is a public link.
                    // Otherwise, it's a shared link.
                    if (permission.invitation == null) {
                        self.recipients["public"] = {
                            "hasPermission": true,
                            "role": permission.roles[0]
                        };

                        deferred.reject({
                            permission_status: 1,
                            message: 'This file has already been shared publicly.',
                            link: self.link
                        });
                    }
                    // Store the email from the invitation property of the permission, and the role of the permission.
                    else if (self.recipients.hasOwnProperty(permission.invitation.email)) {
                        self.recipients[permission.invitation.email] = {
                            "hasPermission": true,
                            "role": permission.roles[0]
                        };
                    }

                });

                deferred.resolve(self.recipients);
            }
        }, 1);

        return deferred.promise;
    }

    return OneDriveShare;
})();
