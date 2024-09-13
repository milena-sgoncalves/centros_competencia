import os
import sys
from dotenv import load_dotenv

#carregar .env
load_dotenv()
ROOT = os.getenv('ROOT')

#Definição dos caminhos
SCRIPTS_PUBLIC_PATH = os.path.abspath(os.path.join(ROOT, 'scripts_public'))

# Adiciona o diretório correto ao sys.path
sys.path.append(SCRIPTS_PUBLIC_PATH)

# from buscar_arquivos_sharepoint import buscar_arquivos_sharepoint
from buscar_arquivos_sharepoint2 import buscar_arquivos_sharepoint
from centros import processar_centros_competencia
from projetos import processar_projetos
from publicacoes import processar_publicacoes
from pedidos_pi import processar_pi
from formacao import processar_formacao
from atividades import processar_atividades
from acs_empresas import processar_acs_empresas
from at_prospeccao import processar_at_prospeccao
from equipe import processar_equipe
from licenciamento_pi import processar_licenciamento
from associadas import processar_associadas
from instituicao import processar_instituicao
from midia import processar_midia
from eventos import processar_eventos
from conselho import processar_conselho
from levar_arquivos_sharepoint import levar_arquivos_sharepoint


if __name__ == "__main__":
    buscar_arquivos_sharepoint()
    processar_centros_competencia()
    processar_projetos()
    processar_publicacoes()
    processar_pi()
    processar_formacao()
    processar_atividades()
    processar_acs_empresas()
    processar_at_prospeccao()
    processar_equipe()
    processar_licenciamento()
    processar_associadas()
    processar_instituicao()
    processar_midia()
    processar_eventos()
    processar_conselho()
    levar_arquivos_sharepoint()