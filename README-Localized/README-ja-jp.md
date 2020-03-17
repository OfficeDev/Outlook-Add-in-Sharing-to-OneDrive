---
page_type: sample
products:
- office-outlook
- office-onedrive
- office-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 3/24/2016 9:32:55 AM
---
# OneDrive に対する Microsoft Outlook アドインの共有

ユーザーは OneDrive アイテムを Outlook アドイン内から直接共有できるようになりました。
このサンプルに、JavaScript API for Office と OneDrive API を使用して、どの電子メール受信者がメッセージ本文の OneDrive リンクを表示するアクセス許可を持つかを表示する Microsoft Outlook アドインを作成する方法を示します。
リンクを表示するための適切なアクセス許可を持たない受信者が存在する場合、ユーザーには選択した受信者にアクセス許可を付与するオプションがあります。

OneDrive `共有` API では、アイテムのリンクを使用してアイテムのアクセス許可をプログラムで取得できます。同じ `共有` API を `action.invite` と一緒に使用して、電子メールの受信者と URL を共有できます。


## 目次

* [前提条件](#prerequisites)
* [プロジェクトを構成する](#configure-the-project)
* [プロジェクトを実行する](#run-the-project)
* [コードを理解する](#understand-the-code)
* [質問とコメント](#questions-and-comments)
* [その他のリソース](#additional-resources)

## 前提条件

このサンプルを実行するには次のものが必要です。

* Visual Studio 2015。Visual Studio 2015 をお持ちでない場合は、無料版の [Visual Studio Community 2015](http://aka.ms/vscommunity2015) をインストールできます。 
* [Microsoft Office Developer Tools for Visual Studio 2015](http://aka.ms/officedevtoolsforvs2015)。
* [Microsoft Office Developer Tools Preview for Visual Studio 2015](http://www.microsoft.com/en-us/download/details.aspx?id=49972)。Microsoft Office Developer Tools for Visual Studio 2015 のベース版とプレビュー版の両方をインストールする必要があることに注意してください。
* Outlook 2016。
* 少なくとも 1 つの電子メール アカウントまたは Office 365 アカウントで Exchange を実行するコンピューター。どちらもお持ちでない場合は、[Office 365 Developer プログラムに参加して、Office 365 の 1 年間無料のサブスクリプションを取得](https://aka.ms/devprogramsignup)できます。
* 個人用の OneDrive アカウント。これは Exchange アカウントとは異なります。
* Internet Explorer 9 以降をインストールする必要がありますが、必ずしも既定のブラウザーにする必要はありません。Office アドインをサポートするために、ホストとして動作する Office のクライアントは、Internet Explorer 9 以降に組み込まれているブラウザー コンポーネントを使用します。

注:このサンプルは現在、コンシューマー OneDrive サービスでのみ機能します。 

## プロジェクトを構成する

1. OneDrive 開発者向けサイトからトークンを取得します。トークンを取得するには、「[OneDrive の認証とサインイン](https://dev.onedrive.com/auth/msa_oauth.htm)」に移動し、[**トークンを取得する**] をクリックします。「_Authentication: bearer_」より後にあるトークンをコピーし、テキスト ファイルに保存します。このトークンは 1 時間有効で、サインインしているユーザーの OneDrive ファイルへの読み取り/書き込みアクセスが付与されます。個人用の OneDrive にサインインする必要があります。
2. ソリューション ファイル **OutlookAddinOneDriveSharing.sln** を開き、`\app\authentication.config.js` ファイルに次のようにトークンを貼り付けます。
```
TOKEN = '<your_token>';
```
3. **ソリューション エクスプローラー**で、**OutlookAddinOneDriveSharing** プロジェクトをクリックし、[**プロパティ**] ウィンドウで [**開始動作**] を [**Office デスクトップ クライアント**] に変更します。

4. **OutlookAddinOneDriveSharing** プロジェクトを右クリックして、[**スタートアップ プロジェクトに設定**] を選択します。
5. Outlook デスクトップ クライアントを閉じます。

## プロジェクトを実行する

**F5** キーを押してプロジェクトを実行します。Outlook を実行するために使用する電子メールとパスワードを入力するよう求めるプロンプトが表示されます。_Exchange_ の電子メールとパスワードを入力します。**注** これは、個人用の OneDrive アカウントの電子メールとパスワードとは異なる場合もあります。 

Outlook デスクトップ クライアントが起動したら、[**電子メールの作成**] をクリックして、新しいメッセージを作成します。

**重要** IIS Express 開発証明書のインストールを受け入れるか確認するメッセージが表示されない場合は、[**コントロール パネル**]、[**プログラムの追加と削除**] の順に移動し、[**IIS Express**] を選択してください。右クリックして [**修復**] を選択します。Visual Studio を再起動し、OutlookAddinOneDriveSharing.sln ファイルを開きます。

このアドインでは[アドイン コマンド](https://msdn.microsoft.com/EN-US/library/office/mt267547.aspx)を使用します。リボンにあるこのコマンド ボタンを選択して、アドインを起動します。

![アクセスを確認するリボン上のコマンド ボタン](/readme-images/commandbutton.PNG)

受信者の一覧と共に作業ウィンドウが表示されます。この一覧は、リンクを表示するアクセス許可を持つユーザーと持たないユーザーで分けられています。
**注** 受信者の追加または削除を行ったり、リンクを変更したりした場合は、必ずコマンド ボタンをもう一度クリックしてリストを更新してください。 

OneDrive リンクを取得するには、www.onedrive.com で OneDrive アカウントにサインインし、ファイルを選択します。そのファイルのリンクをコピーし、電子メール メッセージの本文に貼り付けます。

## コードを理解する

* `app.js` - app.js では、`Office.context.mail.item.getAsync` を使用してメッセージから受信者を取得することによって、受信者のグローバル オブジェクトが作成されます。`Office.context.mail.item.body.getAsync` からでも、リンクは同じ方法で取得されます。
* `onedrive.share.service.js` - OneDrive API への要求を処理するオブジェクト。このオブジェクトには、以下が含まれます。
    - リンクをメンテナンスするためのリンク プロパティ。
    - 要求を OneDrive API エンドポイントに送信し、共有とアクセス許可の API を使用する要求メソッド。
    - 作業ウィンドウのディスプレイをレンダリングする UI オブジェクト。
* `render.controller.js` - 作業ウィンドウのディスプレイを制御するオブジェクト。 

## 注釈

* このサンプルは、メッセージ本文の最初のリンクのみをチェックします。
* 個人用の OneDrive アカウントを使用してトークンを取得する必要があります。
* 個人用の OneDrive アカウントに Outlook アカウントを使用していて、そのアカウントが Office 365 に移行されていない場合、共有は機能しない場合があります。電子メール アカウントが移行されているかを確認するには、Outlook.com にサインインします。左上隅に Outlook.com と表示される場合、移行されていません。

## 質問とコメント

*OneDrive サンプルへの Outlook アドインの共有*についてのフィードバックをお待ちしています。このリポジトリの「*問題*」セクションでフィードバックを送信できます。
Office 365 開発に関する全般的な質問は、「[Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API)」に投稿してください。質問には、必ず [Office365] および [API] のタグを付けてください。

## その他のリソース

* [Office 365 API ドキュメント](http://msdn.microsoft.com/office/office365/howto/platform-development-overview)
* [Microsoft Office 365 API ツール](https://visualstudiogallery.msdn.microsoft.com/a15b85e6-69a7-4fdf-adda-a38066bb5155)
* [Office デベロッパー センター](http://dev.office.com/)
* [Office 365 API スタート プロジェクトとコード サンプル](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)
* [OneDrive デベロッパー センター](http://dev.onedrive.com)
* [Outlook デベロッパー センター](http://dev.outlook.com)

## 著作権
Copyright (c) 2016 Microsoft.All rights reserved.



このプロジェクトでは、[Microsoft オープン ソース倫理規定](https://opensource.microsoft.com/codeofconduct/)が採用されています。詳細については、「[倫理規定の FAQ](https://opensource.microsoft.com/codeofconduct/faq/)」を参照してください。また、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までお問い合わせください。
