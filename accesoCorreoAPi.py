from dotenv import load_dotenv
import os

from exchangelib import Credentials, Account, DELEGATE
from exchangelib.errors import UnauthorizedError


# Cargar variables de entorno desde el archivo .env
load_dotenv()

# Acceder a las variables de entorno
usuario = os.getenv("USUARIO")
password = os.getenv("PASSWORD")
cuenta=os.getenv("CUENTA")

try:
# Configurar las credenciales
    credentials = Credentials(usuario, password)

# Configurar la cuenta de correo
    account = Account(usuario, credentials=credentials, autodiscover=True, access_type=DELEGATE)

# Acceder a la bandeja de entrada
    inbox = account.inbox

# Imprimir información sobre los correos electrónicos en la bandeja de entrada
    for item in inbox.all().order_by('-datetime_received')[:5]:
        print(f'Subject: {item.subject}, Received: {item.datetime_received}, Sender: {item.sender}')
    
# Cerrar la sesión
    account.logout()
except UnauthorizedError as e:
    print(f"Error de autenticación: {e}")
except OSError as e:
    print(f"Error OS: {e}")