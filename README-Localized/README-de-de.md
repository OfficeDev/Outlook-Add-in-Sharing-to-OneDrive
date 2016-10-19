# <a name="microsoft-outlook-add-in-sharing-to-onedrive"></a>Microsoft Outlook-Add-In für die Freigabe in OneDrive

Benutzer können jetzt ein OneDrive-Element direkt von einem Outlook-Add-In aus freigeben. In diesem Beispiel zeigen wir Ihnen, wie die JavaScript-API für Office und die OneDrive-API verwendet werden, um ein Microsoft Outlook-Add-In zu erstellen, das angezeigt, welche E-Mail-Empfänger über die Berechtigung zum Anzeigen des OneDrive-Links im Nachrichtentext verfügen. Wenn Empfänger nicht über die erforderliche Berechtigung zum Anzeigen der Links verfügen, hat der Benutzer die Möglichkeit, ausgewählten Empfängern Berechtigungen zu gewähren.

Mit der OneDrive `shares`-API können Sie programmgesteuert Berechtigungen für ein Element erhalten, indem Sie den Link des Elements verwenden. Anschließend können Sie die gleiche `shares`-API mit `action.invite` verwenden, um die URL für E-Mail-Empfänger freizugeben.


## <a name="table-of-contents"></a>Inhaltsverzeichnis

* [Voraussetzungen](#prerequisites)
* [Konfigurieren des Projekts](#configure-the-project)
* [Ausführen des Projekts](#run-the-project)
* [Grundlegendes zum Code](#understand-the-code)
* [Fragen und Kommentare](#questions-and-comments)
* [Zusätzliche Ressourcen](#additional-resources)

## <a name="prerequisites"></a>Voraussetzungen

Für dieses Beispiel ist Folgendes erforderlich:

* Visual Studio 2015 Wenn Sie nicht über Visual Studio 2015 verfügen, können Sie [Visual Studio Community 2015](http://aka.ms/vscommunity2015) kostenlos installieren. 
* [Vorschau der Microsoft Office Developer Tools für Visual Studio 2015](http://aka.ms/officedevtoolsforvs2015). Beachten Sie, dass sowohl die Basis- als auch die Vorschauversion der Microsoft Office Developer Tools für Visual Studio 2015 installiert werden müssen.
* [Vorschau der Microsoft Office Developer Tools für Visual Studio 2015](http://www.microsoft.com/en-us/download/details.aspx?id=49972).
* Outlook 2016
* Ein Computer mit Exchange mit mindestens einem E-Mail-Konto oder ein Office 365-Konto. Wenn Sie keines dieser Konten besitzen, [nehmen Sie am Office 365 Entwicklerprogramm teil, und erhalten Sie ein kostenloses 1-Jahres-Abonnement für Office 365](https://aka.ms/devprogramsignup).
* Ein persönliches OneDrive-Konto. Dies unterscheidet sich von einem Exchange-Konto.
* Internet Explorer 9 oder höher muss installiert, aber nicht der Standardbrowser sein. Zur Unterstützung von Office-Add-Ins verwendet der Office-Client, der als Host agiert, Browserkomponenten, die Bestandteil von Internet Explorer 9 oder höher sind.

Konfigurieren des Projekts 

## <a name="configure-the-project"></a>Konfigurieren des Projekts

1. Rufen Sie ein Tokens von der OneDrive-Entwicklerwebsite ab. Wechseln Sie zum Abrufen eines Tokens zu [OneDrive-Authentifizierung und Anmeldung](https://dev.onedrive.com/auth/msa_oauth.htm), und klicken Sie auf **Token abrufen**. Kopieren Sie das Token hinter _Authentication: Bearer_, und speichern Sie es in einer Textdatei. Dieses Token ist eine Stunde lang gültig und gewährt Ihnen Lese-/Schreibzugriff auf die OneDrive-Dateien des angemeldeten Benutzers. Sie werden zur Anmeldung bei Ihrem persönlichen OneDrive-Konto aufgefordert.
2. Öffnen Sie die Projektmappendatei **OutlookAddinOneDriveSharing.sln**, und fügen Sie in der Datei `\app\authentication.config.js` das Token wie folgt ein:
```
TOKEN = '<your_token>';
```
3. Klicken Sie im **Projektmappen-Explorer** auf das Projekt **OutlookAddinOneDriveSharing**, und ändern Sie im **Eigenschaftenfenster** die Option **Aktion beginnen** in **Office-Desktopclient**.

4. Klicken Sie mit der rechten Maustaste auf das Projekt **OutlookAddinOneDriveSharing**, und wählen Sie **Als Startprojekt festlegen**.
5. Schließen Sie den Outlook-Desktopclient.

## <a name="run-the-project"></a>Ausführen des Projekts

Drücken Sie **F5**, um das Projekt auszuführen. Sie werden aufgefordert, eine E-Mail-Adresse und ein Kennwort zum Ausführen von Outlook einzugeben. Geben Sie Ihre _Exchange_-E-Mail-Adresse und das Kennwort ein. **Hinweis** Diese weichen möglicherweise von der E-Mail-Adresse und dem Kennwort für Ihr privates OneDrive-Konto ab. 

Sobald der Outlook-Desktopclient gestartet wurde, klicken Sie auf **Neue E-Mail-Nachricht**, um eine neue Nachricht zu erstellen.

**Wichtig** Wenn Sie nicht dazu aufgefordert wurden, die Installation für das IIS Express Development Certificate zu bestätigen, navigieren Sie zu **Systemsteuerung** | **Software**, und wählen Sie **IIS Express**. Klicken Sie mit der rechten Maustaste, und wählen Sie **Reparieren**. Starten Sie Visual Studio neu, und öffnen Sie die Datei „OutlookAddinOneDriveSharing.sln“.

Dieses Add-In verwendet [Add-In-Befehle](https://msdn.microsoft.com/EN-US/library/office/mt267547.aspx); daher starten Sie das Add-In, indem Sie diese Befehlsschaltfläche im Menüband auswählen:

![Befehlsschaltfläche zum Überprüfen des Zugriffs im Menüband](../readme-images/commandbutton.PNG)

Ein Aufgabenbereich mit der Liste der Empfänger wird angezeigt. In der Liste werden die Empfänger danach sortiert, ob Sie über Berechtigungen zum Anzeigen des Links verfügen. **Hinweis** Klicken Sie nach dem Hinzufügen oder Entfernen von Empfängern oder dem Ändern des Links erneut auf die Befehlsschaltfläche, um die Liste zu aktualisieren. 

Melden Sie sich zum Abrufen eines OneDrive-Links bei Ihrem OneDrive-Konto unter www.onedrive.com an, und wählen Sie eine Ihrer Dateien.

## <a name="understand-the-code"></a>Grundlegendes zum Code

* `app.js` - In der Datei „app.js“ wird ein globales Objekt der Empfänger mithilfe von `Office.context.mail.item.getAsync` erstellt, um die Empfänger aus der Nachricht abzurufen. Links werden auf die gleiche Weise mit `Office.context.mail.item.body.getAsync` abgerufen.
* `onedrive.share.service.js` Ein Objekt zum Verarbeiten von Anforderungen für die OneDrive-API.
    - Eine link-Eigenschaft zum Verwalten von Links.
    - Eine Anforderungsmethode zum Senden von Anforderungen an den OneDrive-API-Endpunkt und zum Verwenden der Freigaben- und Berechtigungs-API.
    - Ein UI-Objekt zum Anzeigen des Aufgabenbereichs.
* `render.controller.js` - Ein Objekt zum Steuern der Anzeige im Aufgabenbereich. 

## <a name="remarks"></a>Hinweise

* Das Beispiel überprüft nur den ersten Link im Nachrichtentext.
* Sie müssen ein persönliches OneDrive-Konto zum Abrufen des Tokens verwenden.
* Falls Sie ein Outlook-Konto für das persönliche OneDrive-Konto verwenden und dieses noch nicht zu Office 365 migriert wurde, funktioniert das Freigeben möglicherweise nicht.

## <a name="questions-and-comments"></a>Fragen und Kommentare

Wir schätzen Ihr Feedback hinsichtlich des *Outlook-Add-In-Beispiels für die Freigabe in OneDrive*. Sie können uns Ihr Feedback über den Abschnitt *Probleme* dieses Repositorys senden. Allgemeine Fragen zur Office 365-Entwicklung sollten in [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API) gestellt werden. Stellen Sie sicher, dass Ihre Fragen mit [Office365] und [API] markiert sind.

## <a name="additional-resources"></a>Zusätzliche Ressourcen

* [Dokumentation zu Office 365-APIs](http://msdn.microsoft.com/office/office365/howto/platform-development-overview)
* [Microsoft Office 365 API-Tools](https://visualstudiogallery.msdn.microsoft.com/a15b85e6-69a7-4fdf-adda-a38066bb5155)
* [Office Dev Center](http://dev.office.com/)
* [Office 365 APIs – Startprojekte und Codebeispiele](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)
* [OneDrive Developer Center](http://dev.onedrive.com)
* [Outlook Developer Center](http://dev.outlook.com)

## <a name="copyright"></a>Copyright
Copyright (c) 2016 Microsoft. Alle Rechte vorbehalten.

