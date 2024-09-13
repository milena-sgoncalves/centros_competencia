import os
import sys
import pandas as pd
from dotenv import load_dotenv

#carregar .env
load_dotenv()
ROOT = os.getenv('ROOT')

#sys.path
SCRIPTS_PUBLIC = os.path.abspath(os.path.join(ROOT, 'scripts_public'))
CC_COPY = os.path.abspath(os.path.join(ROOT, 'CC_copy'))
CC_DATA_RAW = os.path.abspath(os.path.join(ROOT, 'CC_data_raw'))
CC_STAGE_AREA = os.path.abspath(os.path.join(ROOT, 'CC_stage_area'))
CC_UP = os.path.abspath(os.path.join(ROOT, 'CC_up'))
sys.path.append(SCRIPTS_PUBLIC)

from processar_excel import processar_excel

def acs_pi():
    # Lendo o arquivo
    acs = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS TECNICO.xlsx')),
                    sheet_name = 'ACS - PI')

    # Salvando o arquivo
    acs.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'acs_pi.xlsx')), index = False)

    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_stage_area')
    nome_arquivo = 'acs_pi.xlsx'
    arquivo_origem = os.path.join(origem, nome_arquivo)
    arquivo_destino = os.path.join(destino, nome_arquivo)

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'Título da PI',
        'Em qual dos projetos, na ação ACS, decorreu o pedido de proteção desta PI?',
        'Tipo de pedido de proteção de PI no INPI',
        'Data do pedido',
        'Número do Pedido no INPI',
        'Link',
        'Observações',
    ]

    novos_nomes_e_ordem = {
        'Em qual dos projetos, na ação ACS, decorreu o pedido de proteção desta PI?': 'codigo_projeto',
        'Título da PI': 'titulo_pi',
        'Tipo de pedido de proteção de PI no INPI': 'tipo_pedido_pi',
        'Data do pedido': 'data_pedido',
        'Número do Pedido no INPI': 'num_pedido_inpi',
        'Link': 'link',
        'Observações': 'observacoes'
    }

    # Campos de data e valor
    campos_data = ['data_pedido']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_data)


def afcct_pi():
    # Lendo o arquivo
    afcct = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS TECNICO.xlsx')),
                        sheet_name = 'AFCCT - PI')
    # Salvando o arquivo
    afcct.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'afcct_pi.xlsx')), index = False)

    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_stage_area')
    nome_arquivo = 'afcct_pi.xlsx'
    arquivo_origem = os.path.join(origem, nome_arquivo)
    arquivo_destino = os.path.join(destino, nome_arquivo)

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'Título da PI',
        'Em qual dos projetos, na ação AFCCT, decorreu o pedido de proteção desta PI?',
        'Tipo de PI',
        'Data do pedido',
        'Número do Pedido no INPI',
        'Link',
        'Observações',
    ]

    novos_nomes_e_ordem = {
        'Em qual dos projetos, na ação AFCCT, decorreu o pedido de proteção desta PI?': 'codigo_projeto',
        'Título da PI': 'titulo_pi',
        'Tipo de PI': 'tipo_pedido_pi',
        'Data do pedido': 'data_pedido',
        'Número do Pedido no INPI': 'num_pedido_inpi',
        'Link': 'link',
        'Observações': 'observacoes'
    }

    # Campos de data e valor
    campos_data = []

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_data)


def juntar_pi():
    # Lendo os arquivos
    afcct = pd.read_excel(os.path.abspath(os.path.join(CC_STAGE_AREA, 'afcct_pi.xlsx')))
    acs = pd.read_excel(os.path.abspath(os.path.join(CC_STAGE_AREA, 'acs_pi.xlsx')))

    # Juntando os arquivos
    pi = pd.concat([afcct, acs], ignore_index=True)

    # Salvando os arquivos
    pi.to_excel(os.path.abspath(os.path.join(CC_UP, 'pi.xlsx')), index = False)



def processar_pi():
    afcct_pi()
    acs_pi()
    juntar_pi()