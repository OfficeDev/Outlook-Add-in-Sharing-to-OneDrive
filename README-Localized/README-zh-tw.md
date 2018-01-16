# <a name="microsoft-outlook-add-in-sharing-to-onedrive"></a>Microsoft Outlook 增益集共用至 OneDrive

使用者現在可以直接從 Outlook 增益集內共用 OneDrive 項目。在這個範例中，為您示範如何使用 JavaScript API for Office 和 OneDrive API 以建立 Microsoft Outlook 增益集，顯示哪一個電子郵件收件者有檢視郵件本文中 OneDrive 連結的權限。如果有收件者沒有適當的權限可以檢視連結，則使用者可以選擇授與權限給選取的收件者。

使用 OneDrive `shares` API，您可以以程式設計方式取得項目的權限，方法是使用項目的連結。然後您可以使用相同的 `shares` API 和 `action.invite`，與電子郵件收件者共用 URL。


## <a name="table-of-contents"></a>目錄

* [必要條件](#prerequisites)
* [設定專案](#configure-the-project)
* [執行專案](#run-the-project)
* [瞭解程式碼](#understand-the-code)
* [問題和建議](#questions-and-comments)
* [其他資源](#additional-resources)

## <a name="prerequisites"></a>必要條件

此範例需要下列項目：

* Visual Studio 2015。如果您沒有 Visual Studio 2015，您可以免費安裝 [Visual Studio Community 2015](http://aka.ms/vscommunity2015)。 
* [Microsoft Office Developer Tools for Visual Studio 2015](http://aka.ms/officedevtoolsforvs2015).
* [Microsoft Office Developer Tools for Visual Studio 2015 預覽](http://www.microsoft.com/en-us/download/details.aspx?id=49972).請注意，必須安裝 Microsoft Office Developer Tools for Visual Studio 2015 基本和預覽。
* Outlook 2016。
* 執行 Exchange 的電腦且具有至少一個電子郵件帳戶，或 Office 365 帳戶。如果也沒有的話，[參加 Office 365 開發人員計劃，並取得 Office 365 的免費 1 年訂用帳戶](https://aka.ms/devprogramsignup)。
* 個人的 OneDrive 帳戶。這與 Exchange 帳戶不同。
* Internet Explorer 9 或更新版本，必須先安裝，但不一定是預設瀏覽器。若要支援 Office 增益集，做為主機的 Office 用戶端會使用 Internet Explorer 9 或更新版本的瀏覽器元件。

附註：這個範例目前只適用於家庭用戶 OneDrive 服務。 

## <a name="configure-the-project"></a>設定專案

1. 從 OneDrive 開發人員網站取得權杖。若要取得權杖，請移至 [OneDrive 驗證和登入](https://dev.onedrive.com/auth/msa_oauth.htm)，然後按一下 [取得權杖]****複製權杖，它在 _Authentication: bearer_ 文字後面，並將它儲存到文字檔。權杖的有效期限為一小時，給予您登入的使用者的 OneDrive 檔案的讀取/寫入存取權。系統會要求您登入您個人的 OneDrive。
2. 開啟解決方案檔案 **OutlookAddinOneDriveSharing.sln**，並且在 `\app\authentication.config.js` 檔案中，貼上權杖，如下所示︰
```
TOKEN = '<your_token>';
```
3. 在 [方案總管]****中，按一下**OutlookAddinOneDriveSharing** 專案，並在[屬性]**** 視窗中將 [起始動作]**** 變更為 [Office 桌面用戶端]****。

4. 以滑鼠右鍵按一下 **OutlookAddinOneDriveSharing** 專案，然後選擇 [設定為啟始專案]****。
5. 關閉 Outlook 桌面用戶端。

## <a name="run-the-project"></a>執行專案

按下 **F5** 以執行專案。系統會提示您輸入用於執行 Outlook 的電子郵件和密碼。輸入您的 _Exchange_ 電子郵件和密碼。**附註** 這可能與您個人的 OneDrive 帳戶電子郵件和密碼不同。 

一旦啟動 Outlook 桌面用戶端之後，按一下 [新增電子郵件]**** 以撰寫新郵件。

**重要** 如果沒有提示您接受 IIS Express 開發憑證的安裝，請瀏覽至 [控制台] **** | [新增/移除程式]****，然後選擇 [IIS Express]****。以滑鼠右鍵按一下，並選擇 [修復]****。重新啟動 Visual Studio 並開啟 OutlookAddinOneDriveSharing.sln 檔案。

這個增益集會使用 [增益功能命令](https://msdn.microsoft.com/EN-US/library/office/mt267547.aspx)，所以您可以藉由選擇功能區上的命令按鈕，啟動增益集︰

![檢查功能區上的存取命令按鈕](/readme-images/commandbutton.PNG)

工作窗格會顯示，並且具有收件者清單。清單會分為具有和不具有檢視連結權限的使用者。 **附註** 新增或移除收件者或變更連結時，再次按一下命令按鈕以重新整理清單。 

若要取得 OneDrive 連結，請在 www.onedrive.com 登入您的 OneDrive 帳戶，並選擇其中一個檔案。複製該檔案的連結，並將它貼到電子郵件訊息的本文。

## <a name="understand-the-code"></a>瞭解程式碼

* `app.js` - 在 app.js 中，會建立收件者的全域物件，方法是使用 `Office.context.mail.item.getAsync` 以從訊息中取得收件者。連結是以相同的方式，使用 `Office.context.mail.item.body.getAsync` 取得。
* `onedrive.share.service.js` - 來處理 OneDrive API 要求的物件。此物件包含︰
    - 維護連結的連結屬性。
    - 傳送要求至 OneDrive API 端點的要求方法，以及使用共用和權限 API。
    - 將顯示轉譯到工作窗格的 UI 物件。
* `render.controller.js` - 控制工作窗格中的顯示的物件。 

## <a name="remarks"></a>備註

* 這個範例只會檢查郵件本文中的第一個連結。
* 您必須使用個人的 OneDrive 帳戶來取得權杖。
* 如果您使用 Outlook 帳戶做為您個人的 OneDrive 帳戶，而且尚未移轉至 Office 365，則共用可能無法運作。若要檢查您的電子郵件帳戶是否已移轉，請登入 Outlook.com，如果左上角顯示 Outlook.com，則表示尚未移轉。

## <a name="questions-and-comments"></a>問題和建議

我們很樂於收到您對於 *Outlook 增益集共用至 OneDrive* 範例的意見反應。您可以在此儲存機制的 [問題]** 區段中，將您的意見反應傳送給我們。請在 [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API) 提出有關 Office 365 開發的一般問題。務必以 [Office365] 和 [API] 標記您的問題。

## <a name="additional-resources"></a>其他資源

* [Office 365 API 文件](http://msdn.microsoft.com/office/office365/howto/platform-development-overview)
* [Microsoft Office 365 API 工具](https://visualstudiogallery.msdn.microsoft.com/a15b85e6-69a7-4fdf-adda-a38066bb5155)
* [Office 開發人員中心](http://dev.office.com/)
* [Office 365 API 入門專案和程式碼範例](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)
* [OneDrive 開發人員中心](http://dev.onedrive.com)
* [Outlook 開發人員中心](http://dev.outlook.com)

## <a name="copyright"></a>著作權
Copyright (c) 2016 Microsoft.著作權所有，並保留一切權利。



此專案已採用 [Microsoft 開放原始碼管理辦法](https://opensource.microsoft.com/codeofconduct/)。如需詳細資訊，請參閱[管理辦法常見問題集](https://opensource.microsoft.com/codeofconduct/faq/)，如果有其他問題或意見，請連絡 [opencode@microsoft.com](mailto:opencode@microsoft.com)。
