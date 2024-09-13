import os
import sys
from dotenv import load_dotenv

#carregar .env
load_dotenv()
ROOT = os.getenv('ROOT')

#Definição dos caminhos
PASTA_ARQUIVOS = os.path.abspath(os.path.join(ROOT, 'CC_up'))
CC_COPY = os.path.abspath(os.path.join(ROOT, 'CC_copy'))
CC_BACKUP = os.path.abspath(os.path.join(ROOT, 'CC_backup'))
PATH_OFFICE = os.path.abspath(os.path.join(ROOT, 'office365_api'))

# Adiciona o diretório correto ao sys.path
sys.path.append(PATH_OFFICE)

from upload_files import upload_files
from zipar_arquivos import zipar_arquivos
from criar_db_sqlite import gerar_db_sqlite
from apagar_arquivos_pasta import apagar_arquivos_pasta

def levar_arquivos_sharepoint():

    gerar_db_sqlite()
    zipar_arquivos(CC_COPY, CC_BACKUP)
    upload_files(CC_BACKUP, "DWPII_backup")
    upload_files(PASTA_ARQUIVOS, "DWPII/centros_competencia")

#Executar função
if __name__ == "__main__":
    levar_arquivos_sharepoint()
