# Uso compartido del complemento de Microsoft Outlook en OneDrive

Los usuarios pueden compartir ahora un elemento de OneDrive directamente desde un complemento de Outlook. En este ejemplo, le mostramos cómo usar la API de JavaScript para Office y la API de OneDrive para crear un complemento de Microsoft Outlook que muestre qué destinatarios de correo que tienen permiso para ver el vínculo de OneDrive en el cuerpo del mensaje. Si hay destinatarios que no tengan el permiso adecuado para ver los vínculos, el usuario tendrá la opción de conceder permisos a los destinatarios seleccionados.

Con la API `shares` de OneDrive, puede obtener mediante programación permisos de un elemento con el vínculo del elemento. Después, puede usar la misma API `shares`, con `action.invite`, para compartir la dirección URL con destinatarios de correo.


## Tabla de contenido

* [Requisitos previos](#prerequisites)
* [Configurar el proyecto](#configure-the-project)
* [Ejecutar el proyecto](#run-the-project)
* [Entender el código](#understand-the-code)
* [Preguntas y comentarios](#questions-and-comments)
* [Recursos adicionales](#additional-resources)

## Requisitos previos

Este ejemplo necesita lo siguiente:

* Visual Studio 2015. Si no tiene Visual Studio 2015, puede instalar [Visual Studio Community 2015](http://aka.ms/vscommunity2015) gratis. 
* [Microsoft Office Developer Tools para Visual Studio 2015](http://aka.ms/officedevtoolsforvs2015).
* [Microsoft Office Developer Tools Preview para Visual Studio 2015](http://www.microsoft.com/en-us/download/details.aspx?id=49972). Tenga en cuenta que deben estar instaladas tanto la versión base como la versión preliminar de Microsoft Office Developer Tools para Visual Studio 2015.
* Outlook 2016.
* Un equipo que ejecute Exchange con al menos una cuenta de correo o una cuenta de Office 365. Puede registrarse para una [suscripción a Office 365 Developer](http://aka.ms/ro9c62) y obtener una cuenta de Office 365.
* Una cuenta personal de OneDrive. Es diferente de una cuenta de Exchange.
* Internet Explorer 9 o posterior (se debe instalar, pero no es necesario que sea el explorador predeterminado). Para admitir Complementos de Office, el cliente de Office que actúa como host usa componentes del explorador que forman parte de Internet Explorer 9 o de una versión posterior.

Nota: Actualmente, este ejemplo solo funciona con el servicio OneDrive de consumidor. 

## Configurar el proyecto

1. Obtenga un token desde el sitio para desarrolladores de OneDrive. Para obtener un token, vaya a [Inicio de sesión y autenticación de OneDrive](https://dev.onedrive.com/auth/msa_oauth.htm) y haga clic en **Get Token** (Obtener token). Copie el token, que aparece después del texto _Authentication: bearer_ y guárdelo en un archivo de texto. Este token es válido durante una hora y le da acceso de lectura y escritura a los archivos de OneDrive del usuario que ha iniciado sesión. Deberá iniciar sesión en su OneDrive personal.
2. Abra el archivo de la solución **OutlookAddinOneDriveSharing.sln** y, en el archivo `\app\authentication.config.js`, pegue el token, de esta forma:
```
TOKEN = '<your_token>';
```
3. En el **Explorador de soluciones**, haga clic en el proyecto **OutlookAddinOneDriveSharing** y, en la **ventana Propiedades**, cambie la **Acción de inicio** a **Cliente de escritorio de Office**.

4. Haga clic con el botón derecho en el proyecto **OutlookAddinOneDriveSharing** y elija **Establecer como proyecto de inicio**.
5. Cierre el cliente para equipo de escritorio de Outlook.

## Ejecutar el proyecto

Pulse **F5** para ejecutar el proyecto. Se le pedirá que escriba un correo electrónico y contraseña para usar para ejecutar Outlook. Escriba el correo y contraseña de _Exchange_. **Nota** Puede ser diferente que el correo y contraseña de la cuenta personal de OneDrive. 

Una vez que se ha iniciado el cliente para equipo de escritorio de Outlook, haga clic en **Nuevo correo electrónico** para redactar un nuevo mensaje.

**Importante** Si se no se le ha pedido que acepte la instalación del certificado de desarrollo de IIS Express, vaya al **Panel de control** | **Agregar o quitar programas** y elija **IIS Express**. Haga clic con el botón derecho y seleccione **Reparar**. Reinicie Visual Studio y abra el archivo OutlookAddinOneDriveSharing.sln.

Este complemento usa [comandos de complemento](https://msdn.microsoft.com/es-es/library/office/mt267547.aspx), por lo que inicia el complemento al elegir este botón de comando en la cinta de opciones:

![Botón de comando Check access (Comprobar acceso) en la cinta de opciones](../readme-images/commandbutton.PNG)

Aparece un panel de tareas con la lista de destinatarios. La lista se divide en quién tiene permiso para ver el vínculo y quién no. 
**Nota** Siempre que agregue o quite destinatarios o cambie el vínculo, vuelva a hacer clic en el botón de comando para actualizar la lista. 

Para obtener un vínculo de OneDrive, inicie sesión en su cuenta de OneDrive en www.onedrive.com y elija uno de los archivos. Copie el vínculo de ese archivo y péguelo en el cuerpo del mensaje de correo.

## Entender el código

* `app.js`: en app.js, se crea un objeto global de destinatarios mediante `Office.context.mail.item.getAsync` para obtener los destinatarios del mensaje. Los vínculos se obtienen de la misma manera, con `Office.context.mail.item.body.getAsync`.
* `onedrive.share.service.js`: un objeto para controlar las solicitudes a la API de OneDrive. Este objeto incluye:
    - Una propiedad de vínculo para mantener los vínculos.
    - Un método de solicitud para enviar solicitudes al punto de conexión de API de OneDrive y usar la API de recursos compartidos y permisos.
    - Un objeto de interfaz de usuario para representar la visualización en el panel de tareas.
* `render.controller.js`: un objeto para controlar la visualización en el panel de tareas. 

## Observaciones

* El ejemplo comprueba solo el primer vínculo en el cuerpo del mensaje.
* Debe usar una cuenta personal de OneDrive para obtener el token.
* Si está usando una cuenta de Outlook para la cuenta personal de OneDrive y aún no se ha migrado a Office 365, puede que no funcione el uso compartido. Para comprobar si se ha migrado su cuenta de correo, inicie sesión en Outlook.com y, si la esquina superior izquierda aparece Outlook.com, no se ha migrado.

## Preguntas y comentarios

Nos encantaría recibir sus comentarios sobre el ejemplo *Uso compartido del complemento de Outlook en OneDrive*. Puede enviarnos comentarios a través de la sección *Problemas* de este repositorio. Las preguntas generales sobre desarrollo en Office 365 deben publicarse en [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Asegúrese de que sus preguntas se etiquetan con [Office365] y [API].

## Recursos adicionales

* [Documentación de las API de Office 365](http://msdn.microsoft.com/office/office365/howto/platform-development-overview)
* [Herramientas de API en Microsoft Office 365](https://visualstudiogallery.msdn.microsoft.com/a15b85e6-69a7-4fdf-adda-a38066bb5155)
* [Centro para desarrolladores de Office](http://dev.office.com/)
* [Proyectos de inicio de las API de Office 365 y ejemplos de código](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)
* [Centro para desarrolladores de OneDrive](http://dev.onedrive.com)
* [Centro para desarrolladores de Outlook](http://dev.outlook.com)

## Copyright
Copyright (c) 2016 Microsoft. Todos los derechos reservados.


