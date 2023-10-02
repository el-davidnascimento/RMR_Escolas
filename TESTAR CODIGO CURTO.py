import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

# Autenticar com o Google Drive
gauth = GoogleAuth()
gauth.LocalWebserverAuth()  # Abre uma janela do navegador para autenticação
drive = GoogleDrive(gauth)

# ID da pasta do Google Drive que você deseja listar
folder_id = '1-bPnm2O379Pcz8Ft-LZ_9Ly1OqoHixt4'

# Listar todos os arquivos na pasta do Google Drive
file_list = drive.ListFile({'q': f"'{folder_id}' in parents"}).GetList()





