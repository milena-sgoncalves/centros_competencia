import os
import sys
from dotenv import load_dotenv
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.files.file import File

#carregar .env
load_dotenv()
ROOT = os.getenv('ROOT')

#Definição dos caminhos
SCRIPTS_PUBLIC_PATH = os.path.abspath(os.path.join(ROOT, 'scripts_public'))
CURRENT_DIR = os.path.abspath(os.path.join(ROOT, 'CC_copy'))
CC_UP = os.path.abspath(os.path.join(ROOT, 'CC_up'))
CC_BACKUP = os.path.abspath(os.path.join(ROOT, 'CC_backup'))
CC_DATA_RAW = os.path.abspath(os.path.join(ROOT, 'CC_data_raw'))
CC_STAGE_AREA = os.path.abspath(os.path.join(ROOT, 'CC_stage_area'))
PATH_OFFICE = os.path.abspath(os.path.join(ROOT, 'office365_api'))

# Adiciona o diretório correto ao sys.path
sys.path.append(SCRIPTS_PUBLIC_PATH)
sys.path.append(PATH_OFFICE)

from download_files import get_files
from apagar_arquivos_pasta import apagar_arquivos_pasta

def buscar_arquivos_sharepoint():
    apagar_arquivos_pasta(CURRENT_DIR)
    apagar_arquivos_pasta(CC_UP)
    apagar_arquivos_pasta(CC_BACKUP)
    apagar_arquivos_pasta(CC_DATA_RAW)
    apagar_arquivos_pasta(CC_STAGE_AREA)

    get_files("DWPII/centros_competencia", CURRENT_DIR)