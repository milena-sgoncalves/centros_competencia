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

def afcct_pub():
    # Lendo o arquivo
    afcct = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS TECNICO.xlsx')),
                        sheet_name = 'AFCCT - Publicações')
    
    # Salvando o arquivo
    afcct.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'afcct_publicacoes.xlsx')), index = False)

    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_stage_area')
    nome_arquivo = 'afcct_publicacoes.xlsx'
    arquivo_origem = os.path.join(origem, nome_arquivo)
    arquivo_destino = os.path.join(destino, nome_arquivo)

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'Título da publicação em periódico de excelência',
        'Em qual dos projetos, na ação AFCCT, decorreu a geração desta publicação?',
        'Data da publicação ou do aceite da publicação',
        'Título do periódico',
        'Classificação do periódico',
        'Nome(s) do(s) autor(es)',
        'Outras informações do periódico (DOI, volume, página etc.)',
        'Link',
        'Observações',
    ]

    novos_nomes_e_ordem = {
        'Em qual dos projetos, na ação AFCCT, decorreu a geração desta publicação?': 'codigo_projeto',
        'Título da publicação em periódico de excelência': 'titulo_publicacao',
        'Data da publicação ou do aceite da publicação': 'data_publicacao',
        'Título do periódico': 'titulo_periodico',
        'Classificação do periódico': 'classificacao_periodico',
        'Nome(s) do(s) autor(es)': 'nome_autor',
        'Outras informações do periódico (DOI, volume, página etc.)': 'info_periodico',
        'Link': 'link',
        'Observações': 'observacoes',
    }

    # Campos de data e valor
    campos_data = ['data_publicacao']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_data)


def acs_pub():
    # Lendo o arquivo
    acs = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS TECNICO.xlsx')),
                    sheet_name = 'ACS - Publicações')

    # Salvando o arquivo
    acs.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'acs_publicacoes.xlsx')), index = False)
    
    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_stage_area')
    nome_arquivo = 'acs_publicacoes.xlsx'
    arquivo_origem = os.path.join(origem, nome_arquivo)
    arquivo_destino = os.path.join(destino, nome_arquivo)

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'Título da publicação em periódico de excelência',
        'Em qual dos projetos, na ação ACS, decorreu a geração desta publicação?',
        'Data da publicação ou do aceite da publicação',
        'Título do periódico',
        'Classificação do periódico',
        'Nome(s) do(s) autor(es)',
        'Outras informações do periódico (DOI, volume, página etc.)',
        'Link',
        'Observações',
    ]

    novos_nomes_e_ordem = {
        'Em qual dos projetos, na ação ACS, decorreu a geração desta publicação?': 'codigo_projeto',
        'Título da publicação em periódico de excelência': 'titulo_publicacao',
        'Data da publicação ou do aceite da publicação': 'data_publicacao',
        'Título do periódico': 'titulo_periodico',
        'Classificação do periódico': 'classificacao_periodico',
        'Nome(s) do(s) autor(es)': 'nome_autor',
        'Outras informações do periódico (DOI, volume, página etc.)': 'info_periodico',
        'Link': 'link',
        'Observações': 'observacoes',
    }

    # Campos de data e valor
    campos_data = ['data_publicacao']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_data)


def juntar_publi():
    # Lendo os arquivos
    afcct = pd.read_excel(os.path.abspath(os.path.join(CC_STAGE_AREA, 'afcct_publicacoes.xlsx')))
    acs = pd.read_excel(os.path.abspath(os.path.join(CC_STAGE_AREA, 'acs_publicacoes.xlsx')))

    # Juntando os arquivos
    publicacoes = pd.concat([afcct, acs], ignore_index=True)

    # Salvando os arquivos
    publicacoes.to_excel(os.path.abspath(os.path.join(CC_UP, 'publicacoes.xlsx')), index = False)

def processar_publicacoes():
    afcct_pub()
    acs_pub()
    juntar_publi()