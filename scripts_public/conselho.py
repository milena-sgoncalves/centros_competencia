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

def processar_conselho():
    # Lendo o arquivo
    conselho = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS FINANCEIROS.xlsx')),
                        sheet_name = 'Conselho Consultivo')
    
    # Salvando o arquivo
    conselho.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'conselho.xlsx')), index = False)

    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_up')
    nome_arquivo = 'conselho.xlsx'
    arquivo_origem = os.path.join(origem, nome_arquivo)
    arquivo_destino = os.path.join(destino, nome_arquivo)

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'Centro de Competência',
        'Nome do membro',
        'CPF',
        'Nome da instituição a qual é vinculado',
        'Categoria (representante)',
        'Nível de representação',
        'Duração do mandato (em anos)',
        'Data de entrada do Conselho',
        'Membro ativo?',
        'Data de saída do Conselho',
        'Observações',
    ]

    novos_nomes_e_ordem = {
        'Centro de Competência': 'centro_competencia',
        'Nome do membro': 'nome',
        'CPF': 'cpf',
        'Nome da instituição a qual é vinculado': 'instituicao',
        'Categoria (representante)': 'categoria',
        'Nível de representação': 'nivel',
        'Duração do mandato (em anos)': 'duracao_mandato',
        'Data de entrada do Conselho': 'data_entrada',
        'Membro ativo?': 'ativo',
        'Data de saída do Conselho': 'data_saida',
        'Observações': 'observacoes',
    }

    # Campos especiais
    campos_data = ['data_entrada', 'data_saida']
    campos_string = ['cpf']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_data, campos_string=campos_string)


#Executar função
if __name__ == "__main__":
    processar_conselho()