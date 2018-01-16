# <a name="microsoft-outlook-add-in-sharing-to-onedrive"></a>Partage de complément Microsoft Outlook dans OneDrive

Les utilisateurs peuvent désormais partager un élément OneDrive directement depuis un complément Outlook. Dans cet exemple, nous allons vous montrer comment utiliser l’interface API JavaScript pour Office et l’API OneDrive afin de créer un complément Microsoft Outlook permettant d’indiquer les destinataires du message qui sont autorisés à visualiser le lien OneDrive dans le corps du message. S’il existe des destinataires ne disposant pas de l’autorisation appropriée pour visualiser les liens, l’utilisateur aura la possibilité d’octroyer des autorisations aux destinataires sélectionnés.

Avec l’API `shares` OneDrive, vous pouvez obtenir des autorisations par programmation pour un élément à l’aide du lien de l’élément. Vous pouvez ensuite utiliser la même API `shares`, avec `action.invite`, pour partager l’URL avec des destinataires du courrier électronique.


## <a name="table-of-contents"></a>Table des matières

* [Conditions préalables](#prerequisites)
* [Configuration du projet](#configure-the-project)
* [Exécution du projet](#run-the-project)
* [Présentation du code](#understand-the-code)
* [Questions et commentaires](#questions-and-comments)
* [Ressources supplémentaires](#additional-resources)

## <a name="prerequisites"></a>Conditions préalables

Cet exemple nécessite les éléments suivants :

* Visual Studio 2015. Si vous n’avez pas Visual Studio 2015, vous pouvez installer [Visual Studio Community 2015](http://aka.ms/vscommunity2015) gratuitement. 
* [Outils de développement Microsoft Office pour Visual Studio 2015](http://aka.ms/officedevtoolsforvs2015).
* [Aperçu des outils de développement Microsoft Office pour Visual Studio 2015](http://www.microsoft.com/en-us/download/details.aspx?id=49972). La base et l’aperçu des outils de développement Microsoft Office pour Visual Studio 2015 doivent être tous les deux installés.
* Outlook 2016.
* Un ordinateur exécutant Microsoft Exchange avec au moins un compte de messagerie ou un compte Office 365. Si vous n’avez aucun des deux, vous pouvez [participer au programme pour les développeurs Office 365 et obtenir un abonnement gratuit d’un an à Office 365](https://aka.ms/devprogramsignup).
* Un compte OneDrive personnel. Ce type de compte est différent d’un compte Exchange.
* Internet Explorer 9 ou une version ultérieure, qui doit être installé mais pas nécessairement défini comme navigateur par défaut. Pour prendre en charge des compléments Office, le client Office qui agit comme hôte utilise des composants de navigateur faisant partie d’Internet Explorer 9 ou d’une version ultérieure.

Remarque : Cet exemple ne fonctionne actuellement qu’avec le service OneDrive grand public. 

## <a name="configure-the-project"></a>Configurer le projet

1. Obtenez un jeton à partir du site de développeur OneDrive. Pour obtenir un jeton, accédez à [Connexion et authentification à OneDrive](https://dev.onedrive.com/auth/msa_oauth.htm) et cliquez sur **Obtenir un jeton**. Copiez le jeton, qui se trouve après le texte _Authentication: bearer_ et enregistrez-le dans un fichier texte. Ce jeton est valide pendant une heure et vous donne accès en lecture/écriture aux fichiers OneDrive de l’utilisateur connecté. Vous allez être amené à vous connecter à votre espace OneDrive personnel.
2. Ouvrez le fichier de solution **OutlookAddinOneDriveSharing.sln** puis, dans le fichier `\app\authentication.config.js`, collez le jeton comme suit :
```
TOKEN = '<your_token>';
```
3. Dans l’**Explorateur de solutions**, cliquez sur le projet **OutlookAddinOneDriveSharing**, puis dans la fenêtre **Propriétés**, modifiez **Action de démarrage** en **Client Office pour ordinateur de bureau**.

4. Cliquez avec le bouton droit sur le projet **OutlookAddinOneDriveSharing**, puis choisissez **Définir comme projet de démarrage**.
5. Fermez le client Outlook pour ordinateur de bureau.

## <a name="run-the-project"></a>Exécuter le projet

Appuyez sur **F5** pour exécuter le projet. Vous serez invité à saisir votre adresse électronique et votre mot de passe pour l’exécution d’Outlook. Saisissez votre adresse _Exchange_ et votre mot de passe. **Remarque** : votre adresse et votre mot de passe peuvent être différents de ceux que vous utilisez pour votre compte OneDrive personnel. 

Une fois que le client Outlook pour ordinateur de bureau a démarré, cliquez sur **Nouveau message électronique** pour écrire un nouveau message.

**Important** : si vous n’avez pas été invité à accepter l’installation du certificat de développement IIS Express, accédez à **Panneau de configuration** | **Ajout/Suppression de programmes** et choisissez **IIS Express**. Cliquez avec le bouton droit et choisissez **Réparer**. Redémarrez Visual Studio et ouvrez le fichier OutlookAddinOneDriveSharing.sln.

Ce complément utilise des [commandes de complément](https://msdn.microsoft.com/EN-US/library/office/mt267547.aspx). De ce fait, vous lancez le complément en cliquant sur ce bouton de commande dans le ruban :

![Bouton de commande de vérification d’accès sur le ruban](/readme-images/commandbutton.PNG)

Un volet Office s’affiche avec la liste des destinataires. La liste comporte deux groupes : les destinataires autorisés à visualiser le lien et ceux non autorisés à le faire. **Remarque** : lorsque vous ajoutez ou supprimez des destinataires, ou que vous modifiez le lien, cliquez à nouveau sur le bouton de commande pour actualiser la liste. 

Pour obtenir un lien OneDrive, connectez-vous à votre compte OneDrive sur www.onedrive.com et choisissez l’un de vos fichiers. Copiez le lien de ce fichier et collez-le dans le corps du message électronique.

## <a name="understand-the-code"></a>Comprendre le code

* `app.js` - Dans app.js, un objet global de destinataires est créé à l’aide de l’élément `Office.context.mail.item.getAsync` pour obtenir les destinataires du message. Les liens sont obtenus de la même façon, avec `Office.context.mail.item.body.getAsync`.
* `onedrive.share.service.js` - Objet pour gérer les demandes envoyées à l’API OneDrive. Cet objet inclut :
    - Une propriété de lien pour tenir les liens à jour.
    - Une méthode de demande pour envoyer des demandes au point de terminaison d’API OneDrive, ainsi que pour utiliser l’API de partages et d’autorisations.
    - Un objet de l’interface utilisateur pour le rendu dans un volet Office.
* `render.controller.js` - Un objet pour contrôler l’affichage dans le volet Office. 

## <a name="remarks"></a>Remarques

* L’exemple vérifie uniquement le premier lien dans le corps du message.
* Vous devez utiliser un compte OneDrive personnel pour obtenir le jeton.
* Si vous utilisez un compte Outlook pour votre compte OneDrive personnel et qu’il n’a pas été migré vers Office 365, il est possible que le partage ne fonctionne pas. Pour vérifier si votre compte de messagerie a été migré, connectez-vous à Outlook.com. Si dans l’angle supérieur gauche vous voyez qu’Outlook.com est indiqué, cela signifie que votre compte n’a pas été migré.

## <a name="questions-and-comments"></a>Questions et commentaires

Nous aimerions recevoir vos commentaires relatifs à l’exemple *Partage de complément Outlook dans OneDrive*. Vous pouvez nous envoyer vos commentaires via la section *Problèmes* de ce référentiel. Si vous avez des questions sur le développement d’Office 365, envoyez-les sur [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Veillez à poser vos questions avec les balises [API] et [Office365].

## <a name="additional-resources"></a>Ressources supplémentaires

* [Documentation sur les API Office 365](http://msdn.microsoft.com/office/office365/howto/platform-development-overview)
* [Outils d’API Microsoft Office 365](https://visualstudiogallery.msdn.microsoft.com/a15b85e6-69a7-4fdf-adda-a38066bb5155)
* [Centre de développement Office](http://dev.office.com/)
* [Projets de démarrage et exemples de code des API Office 365](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)
* [Centre pour développeurs OneDrive](http://dev.onedrive.com)
* [Centre pour développeurs Outlook](http://dev.outlook.com)

## <a name="copyright"></a>Copyright
Copyright (c) 2016 Microsoft. Tous droits réservés.



Ce projet a adopté le [code de conduite Open Source de Microsoft](https://opensource.microsoft.com/codeofconduct/). Pour plus d’informations, reportez-vous à la [FAQ relative au code de conduite](https://opensource.microsoft.com/codeofconduct/faq/) ou contactez [opencode@microsoft.com](mailto:opencode@microsoft.com) pour toute question ou tout commentaire.
