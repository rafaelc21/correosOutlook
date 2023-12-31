from pymongo import MongoClient
from bson.binary import Binary
import os
import tempfile
import base64
import win32com.client


# Configuración de conexión
client = MongoClient("mongodb://localhost:27017/")
# Seleccionar una base de datos
database = client["correos_king"]
# Seleccionar una colección
collection = database["correos"]

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox =  outlook.GetDefaultFolder(6) # 6 entrada 5 salida 
messages = inbox.Items

correo = {}

for mensaje in messages:
    correo = {
        'Tipo': 1,  # Tipo de correo (puedes personalizar según tus necesidades)
        'Tema': mensaje.Subject,  # Asunto del correo
        'EnviadoPor': mensaje.SenderName,  # Remitente del correo
        'Recibido': str(mensaje.ReceivedTime)[:19],  # Fecha y hora de recepción
        'Cuerpo': mensaje.Body,  # Cuerpo del correo en texto sin formato
        'CC': mensaje.CC,  # Copia de carbón (CC) del correo
        'BCC': mensaje.BCC,  # Copia oculta (BCC) del correo
        'Enviado': str(mensaje.SentOn)[:19],  # Fecha y hora de envío
        'Destinatarios': [],  # Lista para almacenar información de destinatarios
        'HTMLBody': mensaje.HTMLBody.encode('utf-8'),  # Cuerpo del correo en formato HTML codificado en UTF-8
        'Adjuntos': []  # Lista para almacenar información de adjuntos
    }
    
    correo_attachment = {
    }
    # Obtener información sobre los destinatarios
    recipients = mensaje.Recipients
    if recipients:
        for i, recipient in enumerate(recipients):
            try:
                # Acceder a las propiedades relevantes del destinatario
                recipient_info = {
                    'Nombre': recipient.Name,
                    'DireccionCorreo': recipient.Address
                }

                # Agregar información del destinatario al documento
                correo['Detinatarios'].append(recipient_info)
            except Exception as e:
                print(f'Error al procesar destinatario: {str(e)}')

    # Obtener información sobre los elementos adjuntos
    attachments = mensaje.Attachments
    if attachments:
        print("Adjuntos:")
        for i, attachment in enumerate(attachments):
            try:
                # Extraer solo algunas propiedades relevantes del adjunto
                temp_file = os.path.join(tempfile.gettempdir(), attachment.FileName)
                # Guardar el archivo adjunto en el sistema de archivos
                attachment.SaveAsFile(temp_file)

                # Leer el contenido del archivo
                with open(temp_file, "rb") as file:
                    file_content = file.read()

                attachment_info = {
                    'Nombre': attachment.FileName,
                    'Tipo': attachment.Type#,
                    #"content": Binary(file_content)
                }

                # Agregar información del adjunto al documento
                correo['Adjuntos'].append(attachment_info)
            except Exception as e:
                print(f'Error al procesar adjunto: {str(e)}')
    else:
        print("No hay adjuntos.")

    # Mostrar información del correo en la consola / Solo para revisar que esta enviando. Se puede omitir     
    print(correo)    

     # Insertar el documento del correo en la colección MongoDB
    collection.insert_one(correo)    

# Cerrar la conexión
client.close()



