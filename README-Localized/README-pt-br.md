# <a name="microsoft-outlook-add-in-sharing-to-onedrive"></a>Compartilhamento do Suplemento do Microsoft Outlook no OneDrive

Os usuários já podem compartilhar itens do OneDrive diretamente de um suplemento do Outlook. Neste exemplo, mostraremos como usar a API JavaScript para Office e a API do OneDrive para criar um suplemento do Microsoft Outlook que exibe os destinatários de email que têm permissão para exibir links do OneDrive no corpo da mensagem. Se houver destinatários que não tenham a permissão adequada para exibir os links, o usuário terá a opção de conceder permissões para destinatários selecionados.

Com a API `shares` do OneDrive, você pode obter permissões programaticamente para um item usando o link do item. Você pode usar a mesma API `shares`, com `action.invite`, para compartilhar a URL com destinatários do email.


## <a name="table-of-contents"></a>Sumário

* [Pré-requisitos](#prerequisites)
* [Configurar o projeto](#configure-the-project)
* [Executar o projeto](#run-the-project)
* [Compreender o código](#understand-the-code)
* [Perguntas e comentários](#questions-and-comments)
* [Recursos adicionais](#additional-resources)

## <a name="prerequisites"></a>Pré-requisitos

Esse exemplo requer o seguinte:

* Visual Studio 2015. Se não tiver o Visual Studio 2015, você poderá instalar o [Visual Studio Community 2015](http://aka.ms/vscommunity2015) gratuitamente. 
* [Microsoft Office Developer Tools para Visual Studio 2015](http://aka.ms/officedevtoolsforvs2015).
* [Microsoft Office Developer Tools Preview para Visual Studio 2015](http://www.microsoft.com/en-us/download/details.aspx?id=49972). Observe que a base e a visualização do Microsoft Office Developer Tools para Visual Studio 2015 devem ser instaladas.
* Outlook 2016.
* Um computador executando o Exchange com pelo menos uma conta de email ou uma conta do Office 365. Caso não tenha nenhuma delas, [participe do Programa para Desenvolvedores do Office 365 e obtenha uma assinatura gratuita de 1 ano do Office 365](https://aka.ms/devprogramsignup).
* Uma conta pessoal do OneDrive. É diferente de uma conta do Exchange.
* Internet Explorer 9 ou posterior, que deve estar instalado, mas não precisa ser o navegador padrão. Para oferecer suporte aos Suplementos do Office, o cliente do Office que atua como host usa os componentes do navegador que fazem parte do Internet Explorer 9 ou posterior.

Observação: Este exemplo atualmente só funciona com o serviço de consumidor do OneDrive. 

## <a name="configure-the-project"></a>Configurar o projeto

1. Obtenha um token do site do desenvolvedor do OneDrive. Para obter um token, vá para [autenticação e entrada do OneDrive](https://dev.onedrive.com/auth/msa_oauth.htm) e clique em **Obter Token**. Copie o token, localizado após o texto _Autenticação: portador_ e salve-o em um arquivo de texto. Esse token é válido por uma hora e oferece acesso de leitura/gravação para os arquivos do OneDrive do usuário conectado. Você precisará entrar em seu OneDrive pessoal.
2. Abra o arquivo de solução **OutlookAddinOneDriveSharing.sln** e, no arquivo `\app\authentication.config.js`, cole o token da seguinte forma:
```
TOKEN = '<your_token>';
```
3. No **Gerenciador de Soluções**, clique no projeto **OutlookAddinOneDriveSharing** e, na **janela Propriedades**, altere **Iniciar Ação** para **Cliente do Office para Área de Trabalho**.

4. Clique com o botão direito do mouse no projeto **OutlookAddinOneDriveSharing** e escolha **Definir como Projeto de Inicialização**.
5. Feche o cliente do Outlook para área de trabalho.

## <a name="run-the-project"></a>Executar o projeto

Pressione **F5** para executar o projeto. Você será solicitado a inserir um email e senha que será usado para executar o Outlook. Insira seu email e senha do _Exchange_. **Observação** Isso pode ser diferente com seu email e senha da conta pessoal do OneDrive. 

Depois que o cliente do Outlook para área de trabalho iniciar, clique em **Novo Email** para redigir uma nova mensagem.

**Importante** Se você não foi solicitado a aceitar a instalação do Certificado de Desenvolvimento do IIS Express, navegue até **Painel de Controle** | **Adicionar/Remover Programas** e escolha **IIS Express**. Clique com o botão direito e escolha **Reparar**. Reinicie o Visual Studio e abra o arquivo OutlookAddinOneDriveSharing.sln.

Este suplemento usa [comandos de suplemento](https://msdn.microsoft.com/EN-US/library/office/mt267547.aspx) para que você inicie o suplemento escolhendo esse botão de comando na faixa de opções:

![Botão de comando de verificação de acesso na faixa de opções](/readme-images/commandbutton.PNG)

Um painel de tarefas é exibida com a lista de destinatários. A lista é dividida por quem têm ou não permissão para exibir o link. **Observação** Sempre que você adicionar ou remover destinatários, ou alterar o link, clique no botão de comando novamente para atualizar a lista. 

Para obter um link do OneDrive, entre em sua conta do OneDrive no endereço www.onedrive.com e escolha um dos seus arquivos. Copie o link desse arquivo e cole-o no corpo da mensagem de email.

## <a name="understand-the-code"></a>Compreender o código

* `app.js` - Em app.js, um objeto global de destinatários é criado usando o `Office.context.mail.item.getAsync` para obter os destinatários da mensagem. Os links são obtidos da mesma maneira com `Office.context.mail.item.body.getAsync`.
* `onedrive.share.service.js` – um objeto para lidar com solicitações à API do OneDrive. Este objeto inclui:
    - Uma propriedade de link para manter links.
    - Um método de solicitação para enviar solicitações para o ponto de extremidade da API do OneDrive e usar a API de compartilhamentos e permissões.
    - Um objeto da interface do usuário para renderizar a exibição para o painel de tarefas.
* `render.controller.js` - Um objeto da interface do usuário para renderizar a exibição para o painel de tarefas. 

## <a name="remarks"></a>Comentários

* O exemplo verifica apenas o primeiro link no corpo da mensagem.
* Você deve usar uma conta pessoal do OneDrive para obter o token.
* O compartilhamento pode não funcionar se você estiver usando uma conta do Outlook para sua conta pessoal do OneDrive e ainda não tiver migrado para o Office 365. Para verificar se sua conta de email foi migrada, entre no Outlook.com e se o canto superior esquerdo exibir Outlook.com sua conta não foi migrada.

## <a name="questions-and-comments"></a>Perguntas e comentários

Adoraríamos receber seus comentários sobre o exemplo do *Compartilhamento do Suplemento do Outlook para o OneDrive*. Você pode enviar comentários na seção *Problemas* deste repositório. As perguntas sobre o desenvolvimento do Office 365 em geral devem ser postadas no [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Não deixe de marcar as perguntas com [Office365] e [API].

## <a name="additional-resources"></a>Recursos adicionais

* [Documentação de APIs do Office 365](http://msdn.microsoft.com/office/office365/howto/platform-development-overview)
* [Ferramentas de API do Microsoft Office 365](https://visualstudiogallery.msdn.microsoft.com/a15b85e6-69a7-4fdf-adda-a38066bb5155)
* [Centro de Desenvolvimento do Office](http://dev.office.com/)
* [Exemplos de código e projetos iniciais de APIs do Office 365](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)
* [Central de desenvolvimento do OneDrive](http://dev.onedrive.com)
* [Central de desenvolvimento do Outlook](http://dev.outlook.com)

## <a name="copyright"></a>Copyright
Copyright © 2016 Microsoft. Todos os direitos reservados.



Este projeto adotou o [Código de Conduta de Software Livre da Microsoft](https://opensource.microsoft.com/codeofconduct/). Para saber mais, confira as [Perguntas frequentes sobre o Código de Conduta](https://opensource.microsoft.com/codeofconduct/faq/) ou contate [opencode@microsoft.com](mailto:opencode@microsoft.com) se tiver outras dúvidas ou comentários.
