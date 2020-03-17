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
# Compartir el complemento de Microsoft Outlook con OneDrive

Ahora, los usuarios pueden compartir un elemento de OneDrive directamente desde un complemento de Outlook.
En este ejemplo se muestra cómo usar la API de JavaScript para Office y la API de OneDrive para crear un complemento de Microsoft Outlook que muestre los destinatarios de correo electrónico que tienen permiso para ver el vínculo de OneDrive en el cuerpo del mensaje.
Si hay destinatarios que no tienen el permiso adecuado para ver el enlace o enlaces, el usuario tendrá la opción de conceder permisos a los destinatarios seleccionados.

Con la API de `acción` de OneDrive, puedes obtener programáticamente los permisos para un artículo usando el enlace del mismo. Luego puede utilizar la misma `acción` de API, con `action.invite`, para compartir la URL con los destinatarios del correo electrónico.


## Tabla de contenido

* [Requisitos previos](#prerequisites)
* [Configurar el proyecto](#configure-the-project)
* [Ejecutar el proyecto](#run-the-project)
* [Entender el código](#understand-the-code)
* [Preguntas y comentarios](#questions-and-comments)
* [Recursos adicionales](#additional-resources)

## Requisitos previos

Este ejemplo necesita lo siguiente:

* Visual Studio 2015. Si no tienes Visual Studio 2015, puedes instalar [Visual Studio Community 2015](http://aka.ms/vscommunity2015) de forma gratuita. 
* [Microsoft Office Developer Tools para Visual Studio 2015](http://aka.ms/officedevtoolsforvs2015).
* [Microsoft Office Developer Tools Preview para Visual Studio 2015](http://www.microsoft.com/en-us/download/details.aspx?id=49972). Tenga en cuenta que tanto la base como la vista previa de Microsoft Office Developer Tools para Visual Studio 2015 deben estar instaladas.
* Outlook 2016.
* Una computadora con Exchange con al menos una cuenta de correo electrónico, o una cuenta de Office 365. Si no tienes ninguno de los dos, puedes[unirte al Programa de Desarrolladores de Office 365 y obtener una suscripción gratuita de 1 año a Office 365](https://aka.ms/devprogramsignup).
* Una cuenta personal de OneDrive. Esto es diferente de una cuenta de Exchange.
* Internet Explorer 9 o posterior, que debe ser instalado, pero no tiene que ser el navegador predeterminado. Para admitir los complementos de Office, el cliente de Office que actúa como host utiliza componentes del explorador que forman parte de Internet Explorer 9 o posterior.

Nota: Actualmente, este ejemplo solo funciona con el servicio OneDrive de consumidor. 

## Configurar el proyecto

1. Consiga un token del sitio de desarrollo de OneDrive. Para obtener un token, vaya a la [autenticación de OneDrive e inicie sesión](https://dev.onedrive.com/auth/msa_oauth.htm) y haga clic en **Get Token**. Copie el token, que está después del texto _Autenticación: portador_ y guárdela en un archivo de texto. Este token es válido por una hora, y le da acceso de lectura/escritura a los archivos de OneDrive del usuario registrado. Deberá iniciar sesión en su OneDrive personal.
2. Abra el archivo de la solución **OutlookAddinOneDriveSharing.sln** y en el archivo `\app\authentication.config.js`, pegue el token, así:
```
TOKEN = '<your_token>';
```
3. En el **Explorador de soluciones**, haga clic en el proyecto **OutlookAddinOneDriveSharing** y en la **ventana Propiedades**, cambie **Acción de inicio** a **Cliente de escritorio de Office**.

4. Haga clic con el botón derecho del ratón en el proyecto **OutlookAddinOneDriveSharing **y elija **Configurar como proyecto de inicio**.
5. Cierre el cliente de escritorio de Outlook.

## Ejecutar el proyecto

Pulse **F5** para ejecutar el proyecto. Se le pedirá que introduzcas un correo electrónico y una contraseña para usar en Outlook. Introduzca su correo electrónico y su contraseña de _Exchange_. **Nota** Esto puede ser diferente del correo electrónico y la contraseña de su cuenta personal de OneDrive. 

Una vez que el cliente de escritorio de Outlook se ha iniciado, haga clic en **Nuevo Correo Electrónico** para redactar un nuevo mensaje.

**Importante** Si no se le pidió que aceptara la instalación del Certificado de Desarrollo de IIS Express, vaya al **Panel de Control** | **Agregar/Quitar Programas** y seleccione **IIS Express**. Haga clic con el botón derecho y seleccione **Reparar**. Reinicie Visual Studio y abra el archivo OutlookAddinOneDriveSharing.sln.

Este complemento usa [comandos de complemento](https://msdn.microsoft.com/EN-US/library/office/mt267547.aspx), así que lanza el complemento eligiendo este botón de comando en la cinta:

![Compruebe el botón de comando de acceso en la cinta ](/readme-images/commandbutton.PNG)

Aparece un panel de tareas con la lista de destinatarios. La lista está dividida por quién tiene permiso para ver el enlace y quién no.
**Nota** Cada vez que añada o elimine destinatarios, o cambie el enlace, haga clic de nuevo en el botón de comando para actualizar la lista. 

Para obtener un enlace de OneDrive, entre en su cuenta de OneDrive en www.onedrive.com y elija uno de sus archivos. Copie el vínculo de ese archivo y péguelo en el cuerpo del mensaje de correo.

## Entender el código

* `app.js`: : En app.js, un objeto global de receptores se crea utilizando el `Office.context.mail.item.getAsync`para obtener los destinatarios del mensaje Los enlaces se obtienen de la misma manera, con `Office.context.mail.item.body.getAsync`.
* `onedrive.share.service.js`: Un objeto para manejar las solicitudes a la API de OneDrive. Este objeto incluye:
    - Una propiedad de vínculo para mantener los vínculos.
    - Un método de solicitud para enviar solicitudes al punto de conexión de API de OneDrive y usar la API de recursos compartidos y permisos.
    - Un objeto de interfaz de usuario para mostrar la pantalla en el panel de tareas.
* `render.controller.js`: un objeto para controlar la pantalla en el panel de tareas. 

## Comentarios

* El ejemplo comprueba solo el primer vínculo en el cuerpo del mensaje.
* Debe usar una cuenta personal de OneDrive para obtener el token.
* Si está usando una cuenta de Outlook para la cuenta personal de OneDrive y aún no se ha migrado a Office 365, puede que no funcione el uso compartido. Para comprobar si se ha migrado su cuenta de correo, inicie sesión en Outlook.com y, si la esquina superior izquierda aparece Outlook.com, no se ha migrado.

## Preguntas y comentarios

Nos encantaría recibir sus comentarios sobre el ejemplo *Uso compartido del complemento de Outlook en OneDrive*.
Puede enviarnos sus comentarios en la sección de *temas* de este repositorio Las preguntas sobre el desarrollo de Office 365 en general deben enviarse a [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Asegúrate de que sus preguntas estén etiquetadas con [Office365] y [API].

## Recursos adicionales

* [Documentación de las API de Office 365](http://msdn.microsoft.com/office/office365/howto/platform-development-overview)
* [Herramientas de API de Microsoft Office 365](https://visualstudiogallery.msdn.microsoft.com/a15b85e6-69a7-4fdf-adda-a38066bb5155)
* [Centro para desarrolladores de Office](http://dev.office.com/)
* [Office 365 APIs proyectos de inicio y muestras de código](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)
* [Centro para desarrolladores de OneDrive](http://dev.onedrive.com)
* [Centro para desarrolladores de Outlook](http://dev.outlook.com)

## Derechos de autor
Copyright (c) 2016 Microsoft. Todos los derechos reservados.



Este proyecto ha adoptado el [Código de conducta de código abierto de Microsoft](https://opensource.microsoft.com/codeofconduct/). Para obtener más información, vea [Preguntas frecuentes sobre el código de conducta](https://opensource.microsoft.com/codeofconduct/faq/) o póngase en contacto con [opencode@microsoft.com](mailto:opencode@microsoft.com) si tiene otras preguntas o comentarios.
