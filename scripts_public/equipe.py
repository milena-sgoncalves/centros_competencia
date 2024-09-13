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

def processar_equipe():
    # Lendo o arquivo
    equipe = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS FINANCEIROS.xlsx')),
                        sheet_name = 'Equipe')
    
    # Salvando o arquivo
    equipe.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'equipe.xlsx')), index = False)

    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_up')
    nome_arquivo = 'equipe.xlsx'
    arquivo_origem = os.path.join(origem, nome_arquivo)
    arquivo_destino = os.path.join(destino, nome_arquivo)

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'Centro de Competência',
        'Nome Completo',
        'CPF / Passaporte',
        'Titulação',
        'Formação Acadêmica (da Graduação)',
        'Link do Currículo Lattes',
        'Atividade / Função',
        'Data de Admissão no CC',
        'Data de Saída do CC',
        'Disponibilidade (horas/mês)',
    ]

    novos_nomes_e_ordem = {
        'Centro de Competência': 'centro_competencia',
        'CPF / Passaporte': 'cpf',
        'Nome Completo': 'nome',
        'Titulação': 'titulacao',
        'Formação Acadêmica (da Graduação)': 'formacao',
        'Link do Currículo Lattes': 'link_lattes',
        'Atividade / Função': 'atividade',
        'Data de Admissão no CC': 'data_admissao',
        'Data de Saída do CC': 'data_saida',
        'Disponibilidade (horas/mês)': 'disponibilidade',
    }

    # Campos especiais
    campos_string = ['cpf']
    campos_data = ['data_admissao', 'data_saida']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_data=campos_data, campos_string=campos_string)


#Executar função
if __name__ == "__main__":
    processar_equipe()