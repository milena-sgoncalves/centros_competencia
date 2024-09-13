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

def processar_midia():
    # Lendo o arquivo
    midia = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS FINANCEIROS.xlsx')),
                        sheet_name = 'Mídia')
    
    # Salvando o arquivo
    midia.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'midia.xlsx')), index = False)

    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_up')
    nome_arquivo = 'midia.xlsx'
    arquivo_origem = os.path.join(origem, nome_arquivo)
    arquivo_destino = os.path.join(destino, nome_arquivo)

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'Centro de Competência',
        'Título da publicação',
        'Nome do veículo de divulgação da publicação',
        'Porta-voz/Fonte',
        'Link da publicação',
        'Data da publicação',
    ]

    novos_nomes_e_ordem = {
        'Centro de Competência': 'centro_competencia',
        'Título da publicação': 'titulo',
        'Nome do veículo de divulgação da publicação': 'veiculo',
        'Porta-voz/Fonte': 'fonte',
        'Link da publicação': 'link',
        'Data da publicação': 'data',
    }

    # Campos especiais
    campos_data = ['data']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_data)


#Executar função
if __name__ == "__main__":
    processar_midia()