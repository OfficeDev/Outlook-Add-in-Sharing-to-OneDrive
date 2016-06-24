# Microsoft Outlook Add-in Sharing to OneDrive

Users can now share a OneDrive item directly from within an Outlook add-in.
In this sample, we show you how to use the JavaScript API for Office, and the OneDrive API to create a Microsoft Outlook Add-in that displays which email recipients have permission to view the OneDrive link in the message body.
If there are recipients that don't have the proper permission to view the link(s), the user will have the option to grant permissions to selected recipients.

With the OneDrive `shares` API, you can programmatically get permissions for an item by using the item's link. You can then use the same `shares` API, with `action.invite`, to share the URL with email recipients.


## Table of Contents

* [Prerequisites](#prerequisites)
* [Configure the project](#configure-the-project)
* [Run the project](#run-the-project)
* [Understand the code](#understand-the-code)
* [Questions and comments](#questions-and-comments)
* [Additional resources](#additional-resources)

## Prerequisites

This sample requires the following:

* Visual Studio 2015. If you don't have Visual Studio 2015, you can install [Visual Studio Community 2015](http://aka.ms/vscommunity2015) for free. 
* [Microsoft Office Developer Tools for Visual Studio 2015](http://aka.ms/officedevtoolsforvs2015).
* [Microsoft Office Developer Tools Preview for Visual Studio 2015](http://www.microsoft.com/en-us/download/details.aspx?id=49972). Note that both base and preview of Microsoft Office Developer Tools for Visual Studio 2015 must be installed.
* Outlook 2016.
* A computer running Exchange with at least one email account, or an Office 365 account. You can sign up for an [Office 365 Developer subscription](http://aka.ms/ro9c62) and get an Office 365 account through it.
* A personal OneDrive account. This is different from an Exchange account.
* Internet Explorer 9 or later, which must be installed, but doesn't have to be the default browser. To support Office Add-ins, the Office client that acts as host uses browser components that are part of Internet Explorer 9 or later.

Note: This sample currently only works with consumer OneDrive service. 

## Configure the project

1. Get a token from the OneDrive developer site. To get a token, go to [OneDrive authentication and sign in](https://dev.onedrive.com/auth/msa_oauth.htm) and click **Get Token**. Copy the token, which is after the text _Authentication: bearer_ and save it to a text file. This token is valid for one hour, and gives you read/write access to the signed in user's OneDrive files. You'll be required to sign in to your personal OneDrive.
2. Open the solution file **OutlookAddinOneDriveSharing.sln** and in the `\app\authentication.config.js` file, paste the token, like this:
```
TOKEN = '<your_token>';
```
3. In **Solution Explorer**, click the **OutlookAddinOneDriveSharing** project and in the **Properties window**, change **Start Action** to **Office Desktop Client**.

4. Right-click the **OutlookAddinOneDriveSharing** project and choose **Set as StartUp Project**.
5. Close Outlook desktop client.

## Run the project

Press **F5** to run the project. You'll be prompted to enter an email and password to use for running Outlook. Enter your _Exchange_ email and password. **Note** This may be different from your personal OneDrive account email and password. 

Once the Outlook desktop client has started, click **New Email** to compose a new message.

**Important** If you weren't prompted to accept the installation for the IIS Express Development Certificate, navigate to **Control Panel** | **Add/Remove Programs** and choose **IIS Express**. Right-click and choose **Repair**. Restart Visual Studio and open the OutlookAddinOneDriveSharing.sln file.

This add-in uses [add-in commands](https://msdn.microsoft.com/EN-US/library/office/mt267547.aspx), so you launch the add-in by choosing this command button on the ribbon:

![Check access command button on the ribbon](/readme-images/commandbutton.PNG)

A task pane appears with the list of recipients. The list is divided by who has permission to view the link, and who doesn't. 
**Note** Any time you add or remove recipients, or change the link, click the command button again to refresh the list. 

To get a OneDrive link, sign in to your OneDrive account at www.onedrive.com and choose one of your files. Copy the link for that file and paste it into the body of the email message.

## Understand the code

* `app.js` - In app.js, a global object of recipients is created by using the `Office.context.mail.item.getAsync` to get recipients from the message. Links are obtained in the same way, with `Office.context.mail.item.body.getAsync`.
* `onedrive.share.service.js` - An object to handle requests to the OneDrive API. This object includes:
    - A link property to maintain links.
    - A request method to send requests to the OneDrive API endpoint, and to use the shares and permissions API.
    - A UI object to render the display to the task pane.
* `render.controller.js` - An object to control the display in the task pane. 

## Remarks

* The sample checks only the first link in the message body.
* You must use a personal OneDrive account to get the token.
* If you are using an Outlook account for your personal OneDrive account and it hasn't been migrated to Office 365, sharing may not work. To check if your email account was migrated, sign in to Outlook.com and if the upper left hand corner says Outlook.com, it's not migrated.

## Questions and comments

We'd love to get your feedback about the *Outlook Add-in Sharing to OneDrive* sample. You can send your feedback to us in the *Issues* section of this repository. 
Questions about Office 365 development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Make sure that your questions are tagged with [Office365] and [API].

## Additional resources

* [Office 365 APIs documentation](http://msdn.microsoft.com/office/office365/howto/platform-development-overview)
* [Microsoft Office 365 API Tools](https://visualstudiogallery.msdn.microsoft.com/a15b85e6-69a7-4fdf-adda-a38066bb5155)
* [Office Dev Center](http://dev.office.com/)
* [Office 365 APIs starter projects and code samples](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)
* [OneDrive developer center](http://dev.onedrive.com)
* [Outlook developer center](http://dev.outlook.com)

## Copyright
Copyright (c) 2016 Microsoft. All rights reserved.

