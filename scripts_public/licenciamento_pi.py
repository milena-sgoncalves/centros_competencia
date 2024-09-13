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

def processar_licenciamento():
    # Lendo o arquivo
    licenciamento = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS FINANCEIROS.xlsx')),
                        sheet_name = 'Licenciamento de PI\'s')
    
    # Salvando o arquivo
    licenciamento.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'licenciamento_pi.xlsx')), index = False)

    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_up')
    nome_arquivo = 'licenciamento_pi.xlsx'
    arquivo_origem = os.path.join(origem, nome_arquivo)
    arquivo_destino = os.path.join(destino, nome_arquivo)

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'Número do Pedido no INPI',
        'Empresa que licenciou é estrangeira?',
        'Razão social da empresa',
        'CNPJ',
        'País sede da empresa que licenciou',
        'Participação nos resultados da exploração do licenciamento',
        'Observações',
    ]

    novos_nomes_e_ordem = {
        'Número do Pedido no INPI': 'num_pedido_inpi',
        'CNPJ': 'cnpj',
        'Razão social da empresa': 'empresa',
        'Empresa que licenciou é estrangeira?': 'empresa_estrangeira',
        'País sede da empresa que licenciou': 'pais',
        'Participação nos resultados da exploração do licenciamento': 'participacao',
        'Observações': 'observacoes',
    }

    # Campos especiais
    campos_string = ['cnpj']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_string=campos_string)


#Executar função
if __name__ == "__main__":
    processar_licenciamento()