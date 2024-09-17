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

def at_prospeccao():
    # Lendo o arquivo
    at = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS TECNICO.xlsx')),
                    sheet_name = 'AT - Prospecção')

    # Salvando o arquivo
    at.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'at_prospeccao.xlsx')), index = False)
    
    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_up')
    nome_arquivo = 'at_prospeccao.xlsx'
    arquivo_origem = os.path.join(origem, nome_arquivo)
    arquivo_destino = os.path.join(destino, nome_arquivo)

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'Centro de Competência',
        'Iniciativa da prospecção',
        'Tipo de interação com a Entidade',
        'CNPJ',
        'Data da prospecção',
        'Breve descrição do resultado da prospecção',
        'Foi emitida proposta para se associar? ',
        'Resultado da proposta',
        'Observações',
    ]

    novos_nomes_e_ordem = {
        'Centro de Competência': 'centro_competencia',
        'CNPJ': 'cnpj',
        'Data da prospecção': 'data_prospeccao',
        'Iniciativa da prospecção': 'iniciativa',
        'Tipo de interação com a Entidade': 'tipo_interacao',
        'Breve descrição do resultado da prospecção': 'resultado_prospeccao',
        'Foi emitida proposta para se associar? ': 'proposta',
        'Resultado da proposta': 'resultado_proposta',
        'Observações': 'observacoes',
    }

    # Campos especiais
    campos_string = ['cnpj']
    campos_data = ['data_prospeccao']
    valores_a_remover = ['']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_data=campos_data, campos_string=campos_string,
                    cnpj = 'cnpj', rm_valor_especifico=True, coluna_valor='cnpj', valores_a_remover=valores_a_remover)


def at_empresas():
    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_up')
    arquivo_origem = os.path.join(origem, 'at_prospeccao.xlsx')
    arquivo_destino = os.path.join(destino, 'at_empresas.xlsx')

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'Razão social da entidade',
        'Nome fantasia da entidade',
        'CNPJ',
        'Nome(s) do(s) contato(s) da entidade',
        'Cargo',
        'Ponto Focal',
    ]

    novos_nomes_e_ordem = {
        'CNPJ': 'cnpj',
        'Razão social da entidade': 'razao_social',
        'Nome fantasia da entidade': 'nome_fantasia',
        'Nome(s) do(s) contato(s) da entidade': 'contato',
        'Cargo': 'cargo_contato',
        'Ponto Focal': 'ponto_focal',
    }

    campos_string = ['cnpj']
    valores_a_remover = ['']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_string=campos_string,
                    cnpj = 'cnpj', coluna = 'cnpj', rm_valor_especifico=True, coluna_valor='cnpj', valores_a_remover=valores_a_remover)


def processar_at_prospeccao():
    at_prospeccao()
    at_empresas()