# correosOutlook
Ejemplo de lectura de correos desde python y volcado a mongo

Es un ejemplo de la utilización de python para leer Oulook y volcar cada uno de los correos de la bandeja de entrada en una base de datos mongo local

Va enlazado a una aplicación que proximamente liberaré, que consiste en la consulta en Nodejs para buscar por palabra dentro del texto del correo

Valores para GetDefaultFolder(6)

0: Carpeta de elementos eliminados (Deleted Items)

1: Bandeja de entrada (Inbox)

2: Elementos enviados (Sent Items)

3: Elementos eliminados (Deleted Items)

4: Elementos eliminados públicos (Public Folders\Deleted Items)

5: Borradores (Drafts)

6: Archivo de elementos eliminados (Outlook Data File\Deleted Items)

9: Elementos eliminados sincronizados (Sync Issues\Conflicts)

10: Elementos eliminados sincronizados (Sync Issues\Local Failures)

11: Elementos eliminados sincronizados (Sync Issues\Server Failures)

12: Calendario (Calendar)

13: Contactos (Contacts)

14: Tareas (Tasks)

15: Diario (Journal)

16: Notas (Notes)

18: Elementos eliminados de búsqueda (Outlook Data File\Search Folders\Deleted Items)

19: Elementos eliminados de búsqueda (Outlook Data File\Search Folders\Unread Mail)


Variables que encontré para poder extraer del correo 

Subject: Asunto del correo electrónico.

SenderName: Nombre del remitente.

ReceivedTime: Fecha y hora de recepción.

Body: Cuerpo del mensaje.

CC: Campo de copia a otros destinatarios (CC).

BCC: Campo de copia oculta a otros destinatarios (BCC).

Attachments: Lista de archivos adjuntos.

SentOn: Fecha y hora de envío.

Importance: Nivel de importancia (puede ser "alta", "normal" o "baja").

Categories: Categorías asociadas al mensaje.

Recipients: Lista de destinatarios.

HTMLBody: Cuerpo del mensaje en formato HTML.
