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

def acs_projetos_empresas():
    # Lendo o arquivo
    acs = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS TECNICO.xlsx')),
                    sheet_name = 'ACS - Empresas Envolvidas')

    # Salvando o arquivo
    acs.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'acs_empresas.xlsx')), index = False)
    
    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_up')
    arquivo_origem = os.path.join(origem, 'acs_empresas.xlsx')
    arquivo_destino = os.path.join(destino, 'acs_projetos_empresas.xlsx')

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'Código do projeto onde a empresa está/esteve envolvida',
        'CNPJ',
        'Caso a empresa seja uma Startup, ela foi atraída ou criada pelas ações do CC?',
        'Houve alavancagem?',
        'Qual valor da alavancagem?',
        'Quem realizou a alavancagem?',
        'Modelo de alavancagem',
        'Observação',
    ]

    novos_nomes_e_ordem = {
        'Código do projeto onde a empresa está/esteve envolvida': 'codigo_projeto',
        'CNPJ': 'cnpj',
        'Caso a empresa seja uma Startup, ela foi atraída ou criada pelas ações do CC?': 'tipo_acao_startup',
        'Houve alavancagem?': 'alavancagem',
        'Qual valor da alavancagem?': 'valor_alavancagem',
        'Quem realizou a alavancagem?': 'responsavel_alavancagem',
        'Modelo de alavancagem': 'modelo_alavancagem',
        'Observação': 'observacoes',
    }

    # Campos especiais
    campos_string = ['cnpj']
    valores_a_remover = ['N/A']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_string=campos_string,
                    cnpj = 'cnpj', rm_valor_especifico=True, coluna_valor='valor_alavancagem', valores_a_remover=valores_a_remover,
                    dropna=True, subset='codigo_projeto')


def acs_empresas():
    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_up')
    nome_arquivo = 'acs_empresas.xlsx'
    arquivo_origem = os.path.join(origem, nome_arquivo)
    arquivo_destino = os.path.join(destino, nome_arquivo)

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'CNPJ',
        'Razão social da empresa',
        'Nome fantasia da empresa',
        'A empresa envolvida é uma Empresa de Base Tecnológica (incluindo Startup)?',
        'Nome do contato da empresa',
        'Cargo',
        'Telefone do contato da empresa',
        'Email do contato da empresa',
    ]

    novos_nomes_e_ordem = {
        'CNPJ': 'cnpj',
        'Razão social da empresa': 'razao_social',
        'Nome fantasia da empresa': 'nome_fantasia',
        'A empresa envolvida é uma Empresa de Base Tecnológica (incluindo Startup)?': 'empresa_tecnologica',
        'Nome do contato da empresa': 'contato',
        'Cargo': 'cargo_contato',
        'Telefone do contato da empresa': 'telefone_contato',
        'Email do contato da empresa': 'email_contato',
    }

    campos_string = ['cnpj']
    valores_a_remover = ['']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_string=campos_string,
                    cnpj = 'cnpj', coluna = 'cnpj', rm_valor_especifico=True, coluna_valor='cnpj', valores_a_remover=valores_a_remover)


def processar_acs_empresas():
    acs_projetos_empresas()
    acs_empresas()