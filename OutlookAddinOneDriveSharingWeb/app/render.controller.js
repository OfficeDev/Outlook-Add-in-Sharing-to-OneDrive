/*
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
*  See LICENSE in the source repository root for complete license information.
*/

'use strict';

/*
* The Renderer object definition. The Renderer controls what is displayed in the task pane.
* @param {String} peopleWithPermissionsContainer - The list of emails with permissions to the link.
* @param {String} peopleWithoutPermissionsContainer - The list of emails without permissions to the link.
* @param {Object} shareButton - A button that handles the event when the user wants to share the URL with all recipients.
* @param {Object} uiSpinner - Progress meter.
* @param {Object} permissionsToBeGranted - A list of emails to grant permissions to a OneDrive item.
* @param {Object} personaTemplate - Display of emails and permission level.
* @param {Object} messageTemplate - Display message.
*/
var Renderer = (function () {

    function Renderer() {
        this.peopleWithPermissionsContainer = document.getElementById("sharedWith");
        this.peopleWithoutPermissionsContainer = document.getElementById("shareWith");
        this.peopleWithPermissionsSection = document.getElementById("sharedWithSection");
        this.peopleWithoutPermissionsSection = document.getElementById("shareWithSection");
        this.shareButton = document.getElementById('setPermissionsToAll');
        this.uiSpinner = document.getElementById('progress');

        this.permissionsToBeGranted = [];

        this.personaTemplate =
            '<div class="ms-Persona">'
            + '<div class="ms-font-m-plus ms-fontWeight-semilight ms-fontColor-neutralPrimary">{{name}}</div>'
            + '<div class="ms-font-s ms-fontColor-neutralSecondaryAlt">{{permission}}</div>'
            + '</div>';

        this.messageTemplate = '<p class="ms-font-l ms-fontColor-neutralSecondaryAlt ms-fontWeight-light">{{message}}</p>';
    }

    /*
    * Renders a display based on the OneDriveShare object.
    */
    Renderer.prototype.renderUI = function (oneDriveContext) {
        var self = this;
        return oneDriveContext.initialize()
            .then(function () {
                return self.renderPeopleAccessLevels(oneDriveContext.recipients);
            }, function (error) {
                return self.renderSharingExceptions(error, oneDriveContext.recipients);
            })
            .then(function (recipientsWithoutPermissions) {
                if (recipientsWithoutPermissions && recipientsWithoutPermissions.length > 0) {
                    self.showShareButton();
                    $(self.shareButton).click(function () {
                        oneDriveContext
                            .grantPermissionsToRecipients(recipientsWithoutPermissions)
                            .then(function () {
                                return self.renderUI(oneDriveContext);
                            })
                    });
                }
                else {
                    self.hideShareButton();
                }
            })
    }

    Renderer.prototype.clearUI = function () {
        this.permissionsToBeGranted = [];
        this.clearSharedWithContainer();
        this.clearShareWithContainer("We're looking for recipients and files in your message.");
        this.hideShareButton();
        this.hideSharedWithSection();
    }

    Renderer.prototype.hideShareButton = function () {
        $(this.shareButton).hide();
        $(this.shareButton).unbind('click');
    }

    Renderer.prototype.showShareButton = function () {
        $(this.shareButton).show();
    }

    Renderer.prototype.hideSharedWithSection = function () {
        $(this.peopleWithPermissionsSection).hide();
    }

    Renderer.prototype.hideShareWithSection = function () {
        $(this.peopleWithoutPermissionsSection).hide();
    }

    Renderer.prototype.showSharedWithSection = function () {
        $(this.peopleWithPermissionsSection).show();
    }

    Renderer.prototype.showShareWithSection = function () {
        $(this.peopleWithoutPermissionsSection).show();
    }

    Renderer.prototype.showShareButton = function () {
        $(this.shareButton).show();
    }

    Renderer.prototype.showProgress = function (promise) {
        var self = this;
        if (promise && promise.then) {
            $(self.uiSpinner).show();
            promise.then(
                function (success) {
                    $(self.uiSpinner).hide();
                    return success;
                },
                function (failure) {
                    $(self.uiSpinner).hide();
                    return failure;
                }
            );
        }
    }

    Renderer.prototype.clearSharedWithContainer = function (message) {
        this.peopleWithPermissionsContainer.innerHTML = message ? this.messageTemplate.replace('{{message}}', message) : "";
    }

    Renderer.prototype.clearShareWithContainer = function (message) {
        this.peopleWithoutPermissionsContainer.innerHTML = message ? this.messageTemplate.replace('{{message}}', message) : "";
    }

    Renderer.prototype.renderPeopleAccessLevels = function (recipients) {
        var clearFlagForPeopleWithPermissions = true,
            clearFlagForPeopleWithoutPermissions = true;

        for (var recipient in recipients) {
            if (recipients[recipient].hasPermission) {
                if (clearFlagForPeopleWithPermissions) {
                    this.showSharedWithSection();
                    clearFlagForPeopleWithPermissions = false;
                    this.clearSharedWithContainer();
                    this.hideShareButton();
                }
                var template = this.personaTemplate.replace('{{name}}', recipient)
                template = template.replace('{{permission}}', recipients[recipient].role)
                this.peopleWithPermissionsContainer.innerHTML += template;
            }
            else {
                if (clearFlagForPeopleWithoutPermissions) {
                    this.showShareWithSection();
                    clearFlagForPeopleWithoutPermissions = false;
                    this.clearShareWithContainer();
                    this.showShareButton();
                }
                var template = this.personaTemplate.replace('{{name}}', recipient)
                template = template.replace('{{permission}}', '')
                this.peopleWithoutPermissionsContainer.innerHTML += template;
                this.permissionsToBeGranted.push(recipient);
            }
        }

        if (clearFlagForPeopleWithoutPermissions) {
            this.clearShareWithContainer("This file has already been shared with all recipients.")
        }

        if (clearFlagForPeopleWithPermissions) {
            this.clearSharedWithContainer("This file cannot be shared.")
        }

        return this.permissionsToBeGranted;
    }

    Renderer.prototype.renderSharingExceptions = function (error, recipients) {
        if (!error) return;

        if (error.hasOwnProperty('permission_status')) {
            this.clearSharedWithContainer(error.message);
            if (error.permission_status == 0) {
                this.showSharedWithSection();
                this.clearShareWithContainer();
                this.showShareButton();
                for (var recipient in recipients) {
                    var template = this.personaTemplate.replace('{{name}}', recipient)
                    template = template.replace('{{permission}}', '')
                    this.peopleWithoutPermissionsContainer.innerHTML += template;
                    this.permissionsToBeGranted.push(recipient);
                }
            }
            else {
                this.clearShareWithContainer("Good news! Everyone has permission to see this file.");
            }
        }

        if (error.hasOwnProperty('statusText')) {
            throw error;
        }

        return this.permissionsToBeGranted;
    }

    Renderer.prototype.notify = function (message) {
        Office.context.mailbox.item.notificationMessages
            .addAsync("subject", {
                type: "informationalMessage",
                icon: "blue-icon-16",
                message: message,
                persistent: false
            });
    }

    return Renderer;
})();
