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

def processar_eventos():
    # Lendo o arquivo
    eventos = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS FINANCEIROS.xlsx')),
                        sheet_name = 'Eventos')
    
    # Salvando o arquivo
    eventos.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'eventos.xlsx')), index = False)

    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_up')
    nome_arquivo = 'eventos.xlsx'
    arquivo_origem = os.path.join(origem, nome_arquivo)
    arquivo_destino = os.path.join(destino, nome_arquivo)

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'Centro de Competência',
        'Título do evento',
        'Tipo de evento',
        'Formato do evento',
        'Tipo de participação do CC',
        'Local de realização (Cidade/Estado/País)',
        'Data de realização',
        'Nome dos Participantes do CC',
        'Resultado alcançado com a participação no evento',
        'Informar entidades as quais foram feitos contatos',
        'Observações',
    ]

    novos_nomes_e_ordem = {
        'Centro de Competência': 'centro_competencia',
        'Título do evento': 'titulo',
        'Tipo de evento': 'tipo_evento',
        'Formato do evento': 'formato',
        'Tipo de participação do CC': 'tipo_participacao',
        'Local de realização (Cidade/Estado/País)': 'local',
        'Data de realização': 'data',
        'Nome dos Participantes do CC': 'participantes',
        'Resultado alcançado com a participação no evento': 'resultado',
        'Informar entidades as quais foram feitos contatos': 'entidades_contatadas',
        'Observações': 'observacoes',
    }

    # Campos especiais
    campos_data = ['data']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_data)


#Executar função
if __name__ == "__main__":
    processar_eventos()