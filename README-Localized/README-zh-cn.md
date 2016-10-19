# <a name="microsoft-outlook-add-in-sharing-to-onedrive"></a>Microsoft Outlook 外接程序的 OneDrive 共享

现在，用户可以直接从 Outlook 外接程序内共享 OneDrive 项。本示例介绍如何使用适用于 Office 的 JavaScript API，以及如何使用 OneDrive API 创建 Microsoft Outlook 外接程序，用于显示哪些电子邮件收件人拥有查看邮件正文中 OneDrive 链接的权限。如果收件人不具有查看链接的正确权限，用户可以选择将权限授予所选收件人。

借助 OneDrive `shares` API，你可以通过使用项目的链接以编程方式获取相应项目的权限。然后，你可以结合使用相同的 `shares` API 和 `action.invite` 与电子邮件收件人共享 URL。


## <a name="table-of-contents"></a>目录

* [先决条件](#prerequisites)
* [配置项目](#configure-the-project)
* [运行项目](#run-the-project)
* [了解代码](#understand-the-code)
* [问题和意见](#questions-and-comments)
* [其他资源](#additional-resources)

## <a name="prerequisites"></a>先决条件

此示例需要以下各项：

* Visual Studio 2015。如果你未安装 Visual Studio 2015，则可以免费安装 [Visual Studio Community 2015](http://aka.ms/vscommunity2015)。 
* [Microsoft Visual Studio 的 Office 开发人员工具 2015](http://aka.ms/officedevtoolsforvs2015)。
* [Microsoft Visual Studio 的 Office 开发人员工具 2015 预览版](http://www.microsoft.com/en-us/download/details.aspx?id=49972)。请注意，必须同时安装 Microsoft Visual Studio 的 Office 开发人员工具 2015 基础版和预览版。
* Outlook 2016。
* 运行至少具有一个电子邮件帐户或 Office 365 帐户的 Exchange 的计算机。如果你没有任一帐户，可以 [参加 Office 365 开发人员计划并获取为期 1 年的免费 Office 365 订阅](https://aka.ms/devprogramsignup)。
* OneDrive 个人帐户。这不同于 Exchange 帐户。
* 必须安装 Internet Explorer 9 或更高版本，但不一定作为默认浏览器。为了支持 Office 外接程序，作为主机的 Office 客户端使用属于 Internet Explorer 9 或更高版本的一部分的浏览器组件。

注意：此示例目前仅适用于 Consumer OneDrive 服务。 

## <a name="configure-the-project"></a>配置项目

1. 从 OneDrive 开发人员网站获取令牌。若要获取令牌，请转到 [OneDrive authentication and sign in](https://dev.onedrive.com/auth/msa_oauth.htm)（OneDrive 身份验证和登录），然后单击“**Get Token**”（获取令牌）。复制 _Authentication: bearer_ 文本后面的令牌，并将其保存到文本文件中。此令牌的有效期为一小时，并为你提供对已登录用户的 OneDrive 文件的读取/写入访问权限。你将需要登录到你的个人 OneDrive。
2. 打开解决方案文件 **OutlookAddinOneDriveSharing.sln**，并在 `\app\authentication.config.js` 文件中粘贴此令牌，如下所示：
```
TOKEN = '<your_token>';
```
3. 在“**解决方案资源管理器**”中，单击“**OutlookAddinOneDriveSharing**”项目，并在“**属性窗口**”中，将“**启动操作**”更改为“**Office 桌面客户端**”。

4. 右键单击“**OutlookAddinOneDriveSharing**”项目，然后选择“**设为启动项目**”。
5. 关闭 Outlook 桌面客户端。

## <a name="run-the-project"></a>运行项目

按 **F5** 即可运行项目。系统将提示你输入用于运行 Outlook 的电子邮件和密码。输入你的 _Exchange_ 电子邮件和密码。**注意** 这可能不同于你的个人 OneDrive 帐户电子邮件和密码。 

在 Outlook 桌面客户端启动后，请单击“**新建电子邮件**”撰写一封新邮件。

**重要说明** 如果系统未提示你接受 IIS Express 开发证书的安装，请导航到“**控制面板** | **添加/删除程序**”并选择“**IIS Express**”。右键单击并选择“**修复**”。重启 Visual Studio 并打开 OutlookAddinOneDriveSharing.sln 文件。

此外接程序使用[外接程序命令](https://msdn.microsoft.com/EN-US/library/office/mt267547.aspx)，因此你可以通过在功能区上选择此命令按钮来启动该外接程序：

![查看功能区上的访问命令按钮](../readme-images/commandbutton.PNG)

任务窗格中显示收件人列表。该列表按拥有和没有查看链接权限的收件人进行划分。 **注意** 你可以随时添加或删除收件人，或者更改该链接，再次单击该命令按钮即可刷新该列表。 

若要获取 OneDrive 链接，请在 www.onedrive.com 登录到你的 OneDrive 帐户，并选择你的其中一个文件。复制该文件的链接并将其粘贴到电子邮件的正文中。

## <a name="understand-the-code"></a>了解代码

* `app.js` - 在 app.js 中，通过使用 `Office.context.mail.item.getAsync` 从邮件中获取收件人来创建收件人的全局对象。使用 `Office.context.mail.item.body.getAsync` 以相同的方式获取链接。
* `onedrive.share.service.js` - 用于处理 OneDrive API 请求的对象。此对象包括：
    - 维护链接的链接属性。
    - 发送请求到 OneDrive API 终结点以及使用共享和权限 API 的请求方法。
    - 呈现任务窗格中的显示内容的 UI 对象。
* `render.controller.js` - 控制任务窗格中的显示内容的对象。 

## <a name="remarks"></a>注解

* 该示例仅检查邮件正文中的第一个链接。
* 你必须使用个人的 OneDrive 帐户获得令牌。
* 如果你要将 Outlook 帐户用作你的个人 OneDrive 帐户，并且该帐户尚未迁移到 Office 365，则可能无法共享。若要检查你的电子邮件帐户是否已迁移，请登录到 Outlook.com 并且如果左上角出现 Outlook.com，则说明没有迁移。

## <a name="questions-and-comments"></a>问题和意见

我们希望得到你对 *Outlook 外接程序共享到 OneDrive* 示例的相关反馈。你可以在该存储库中的“*问题*”部分将反馈发送给我们。与 Office 365 开发相关的问题一般应发布到 [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API)。确保你的问题使用了 [Office365] 和 [API] 标记。

## <a name="additional-resources"></a>其他资源

* [Office 365 API 文档](http://msdn.microsoft.com/office/office365/howto/platform-development-overview)
* [Microsoft Office 365 API 工具](https://visualstudiogallery.msdn.microsoft.com/a15b85e6-69a7-4fdf-adda-a38066bb5155)
* [Office 开发人员中心](http://dev.office.com/)
* [Office 365 API 初学者项目和代码示例](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)
* [OneDrive 开发人员中心](http://dev.onedrive.com)
* [Outlook 开发人员中心](http://dev.outlook.com)

## <a name="copyright"></a>版权
版权所有 (c) 2016 Microsoft。保留所有权利。

