import os
import sys
from dotenv import load_dotenv
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

#carregar .env
load_dotenv()
ROOT = os.getenv('ROOT')
USERNAME = os.getenv('sharepoint_email')
PASSWORD = os.getenv('sharepoint_password')
SHAREPOINT_URL_SITE = os.getenv('sharepoint_url_site')
FOLDER_RELATIVE_URL = os.getenv('folder_relative_url')

#Definição dos caminhos
SCRIPTS_PUBLIC_PATH = os.path.abspath(os.path.join(ROOT, 'scripts_public'))
CURRENT_DIR = os.path.abspath(os.path.join(ROOT, 'CC_copy'))
CC_UP = os.path.abspath(os.path.join(ROOT, 'CC_up'))
CC_BACKUP = os.path.abspath(os.path.join(ROOT, 'CC_backup'))
CC_DATA_RAW = os.path.abspath(os.path.join(ROOT, 'CC_data_raw'))
CC_STAGE_AREA = os.path.abspath(os.path.join(ROOT, 'CC_stage_area'))

# Adiciona o diretório correto ao sys.path
sys.path.append(SCRIPTS_PUBLIC_PATH)

from apagar_arquivos_pasta import apagar_arquivos_pasta

def get_files():
    # Conectar ao SharePoint com credenciais
    ctx = ClientContext(SHAREPOINT_URL_SITE).with_credentials(UserCredential(USERNAME, PASSWORD))

    try:
        # Obter a pasta usando o caminho relativo
        folder = ctx.web.get_folder_by_server_relative_url(FOLDER_RELATIVE_URL).expand(["Files"]).get().execute_query()
        files = folder.files

        # Baixar arquivos da pasta
        for file in files:
            file_name = file.name
            if file_name.endswith((".xls", ".xlsx")):  # Filtra arquivos Excel
                download_path = os.path.join(CURRENT_DIR, file_name)
                with open(download_path, "wb") as local_file:
                    file.download(local_file).execute_query()
                print(f"Arquivo '{file_name}' baixado com sucesso para {download_path}.")
    except Exception as e:
        print(f"Erro ao acessar os arquivos da pasta: {e}")



def buscar_arquivos_sharepoint():
    apagar_arquivos_pasta(CURRENT_DIR)
    apagar_arquivos_pasta(CC_UP)
    apagar_arquivos_pasta(CC_BACKUP)
    apagar_arquivos_pasta(CC_DATA_RAW)
    apagar_arquivos_pasta(CC_STAGE_AREA)

    get_files()