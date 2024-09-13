import os
import sys
import pandas as pd
from dotenv import load_dotenv
import re

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

def centros_associadas():
    # Lendo o arquivo
    associadas = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS FINANCEIROS.xlsx')),
                        sheet_name = 'Acompanhamento de Associadas')

    # Salvando o arquivo
    associadas.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'associadas.xlsx')), index = False)
    
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_up')
    arquivo_origem = os.path.join(origem, 'associadas.xlsx')
    arquivo_destino = os.path.join(destino, 'centros_associadas.xlsx')

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'Centro de Competência',
        'CNPJ',
        'Data de início da associação',
        'Data de encerramento da associação',
        'Valor da contribuição no Ano 1',
        'Valor da contribuição no Ano 2',
        'Valor da contribuição no Ano 3',
        'Valor da contribuição no Ano 4',
        'Valor da contribuição financeira para o período de associação',
        'Houve doação de infraestrutura pela entidade associada?',
        'Observações',
    ]

    novos_nomes_e_ordem = {
        'Centro de Competência': 'centro_competencia',
        'CNPJ': 'cnpj',
        'Data de início da associação': 'data_inicio',
        'Data de encerramento da associação': 'data_encerramento',
        'Valor da contribuição no Ano 1': 'valor_ano1',
        'Valor da contribuição no Ano 2': 'valor_ano2',
        'Valor da contribuição no Ano 3': 'valor_ano3',
        'Valor da contribuição no Ano 4': 'valor_ano4',
        'Valor da contribuição financeira para o período de associação': 'valor_periodo',
        'Houve doação de infraestrutura pela entidade associada?': 'doacao_infraestrutura',
        'Observações': 'observacoes',
    }

    # Campos especiais
    campos_string = ['cnpj']
    campos_data = ['data_inicio', 'data_encerramento']
    campos_valor = ['valor_ano1', 'valor_ano2', 'valor_ano2', 'valor_ano4', 'valor_periodo']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino,
                    campos_data=campos_data, campos_valor=campos_valor, campos_string=campos_string,
                    cnpj = 'cnpj')
    


def emp_associadas():
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_up')
    nome_arquivo = 'associadas.xlsx'
    arquivo_origem = os.path.join(origem, nome_arquivo)
    arquivo_destino = os.path.join(destino, nome_arquivo)

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'Razão social da entidade',
        'Nome fantasia da entidade',
        'CNPJ',
        'A empresa é uma startup?',
        'Nome do contato da entidade associada',
        'Cargo',
        'Email do contato da entidade associada',
    ]

    novos_nomes_e_ordem = {
        'CNPJ': 'cnpj',
        'Razão social da entidade': 'razao_social',
        'Nome fantasia da entidade': 'nome_fantasia',
        'A empresa é uma startup?': 'startup',
        'Nome do contato da entidade associada': 'contato',
        'Cargo': 'cargo_contato',
        'Email do contato da entidade associada': 'email_contato',
    }

    # Campos especiais
    campos_string = ['cnpj']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_string=campos_string,
                    cnpj = 'cnpj', coluna = 'cnpj')



def processar_associadas():
    centros_associadas()
    emp_associadas()