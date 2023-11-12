from dotenv import load_dotenv
import os

from exchangelib import Account, Credentials

# Cargar variables de entorno desde el archivo .env
load_dotenv()

# Acceder a las variables de entorno
usuario = os.getenv("USUARIO")
password = os.getenv("PASSWORD")
cuenta=os.getenv("CUENTA")

# Configura las credenciales y la URL del servidor Exchange
credentials = Credentials(usuario, password)
account = Account(usuario, credentials=credentials, autodiscover=True)

# Accede a la bandeja de entrada
inbox = account.inbox

# Itera a través de los correos electrónicos en la bandeja de entrada
for item in inbox.all():
    print(f'Subject: {item.subject}')
    print(f'From: {item.sender}')
    print(f'To: {item.to}')
    print(f'Body: {item.body}')

# Cierra la sesión
account.logout()