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

def ler_salvar_arquivos():
    # Lendo os arquivos
    acs = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS TECNICO.xlsx')),
                        sheet_name = 'ACS - Info de Atividades')
    at = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS TECNICO.xlsx')),
                        sheet_name = 'AT - Atividade')

    # Salvando os arquivos
    acs.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'acs_atividade.xlsx')), index = False)
    at.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'at_atividade.xlsx')), index = False)

def acs_atividade():
    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_stage_area')
    arquivo_origem = os.path.join(origem, 'acs_atividade.xlsx')
    arquivo_destino = os.path.join(destino, 'acs_atividade.xlsx')

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'Centro de Competência',
        'Título da Atividade de Atração e Criação de Startup',
        'Código da atividade',
        'Objetivo',
        'Descrição Pública',
        'Status atual da atividade',
        'Valor total planejado do projeto (todas as fontes)',
        'Valor aportado pela EMBRAPII (até o momento)',
        'Valor aportado pela AT (até o momento)',
        'Valor aportado por Outras Fontes (até o momento)',
        'Valor aportado por empresas (fora dos aportes da AT) (até o momento)',
        ' Data real de início',
        ' Data real de término',
        'Quantidade startups envolvidas',
        'Observações',
    ]

    novos_nomes_e_ordem = {
        'Centro de Competência': 'centro_competencia',
        'Código da atividade': 'codigo_atividade',
        'Título da Atividade de Atração e Criação de Startup': 'titulo_atividade',
        'Objetivo': 'objetivo_atividade',
        'Descrição Pública': 'descricao_publica_atividade',
        'Status atual da atividade': 'status_atividade',
        'Valor total planejado do projeto (todas as fontes)': 'valor_total',
        'Valor aportado pela EMBRAPII (até o momento)': 'valor_embrapii',
        'Valor aportado pela AT (até o momento)': 'valor_at',
        'Valor aportado por Outras Fontes (até o momento)': 'valor_outras_fontes',
        'Valor aportado por empresas (fora dos aportes da AT) (até o momento)': 'valor_empresas',
        ' Data real de início': 'data_inicio_real',
        ' Data real de término': 'data_termino_real',
        'Quantidade startups envolvidas': 'num_startups',
        'Observações': 'observacoes',
    }

    # Campos de data e valor
    campos_data = ['data_inicio_real', 'data_termino_real']
    campos_valor = ['valor total', 'valor_embrapii', 'valor_at', 'valor_outras_fontes', 'valor_empresas']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_data, campos_valor)


def at_atividade():
    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_stage_area')
    arquivo_origem = os.path.join(origem, 'at_atividade.xlsx')
    arquivo_destino = os.path.join(destino, 'at_atividade.xlsx')

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'Centro de Competência',
        'Título da atividade executada na Associação Tecnológica',
        'Código da atividade',
        'Quais associadas participaram?',
        'Objetivo',
        'Descrição Pública',
        'Status atual da atividade',
        'Valor executado para esta atividade (se houver)',
        'Valor da AT ou outra fonte?',
        'Valor executado (AT)',
        'Valor executado (Outra fonte)',
        ' Data real de início',
        ' Data real de término',
        'Observações',
    ]

    novos_nomes_e_ordem = {
        'Centro de Competência': 'centro_competencia',
        'Código da atividade': 'codigo_atividade',
        'Título da atividade executada na Associação Tecnológica': 'titulo_atividade',
        'Objetivo': 'objetivo_atividade',
        'Descrição Pública': 'descricao_publica',
        'Status atual da atividade': 'status_atividade',
        'Valor executado para esta atividade (se houver)': 'valor_total',
        'Valor executado (AT)': 'valor_at',
        'Valor executado (Outra fonte)': 'valor_outras_fontes',
        ' Data real de início': 'data_inicio_real',
        ' Data real de término': 'data_termino_real',
        'Observações': 'observacoes',
        'Quais associadas participaram?': 'associadas',
        'Valor da AT ou outra fonte?': 'fonte_valor',
    }

    # Campos de data e valor
    campos_data = ['data_inicio_real', 'data_termino_real']
    campos_valor = ['valor total', 'valor_at', 'valor_outras_fontes', 'valor_empresas']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_data, campos_valor)


def juntar_salvar_arq():
    # Lendo os arquivos
    acs = pd.read_excel(os.path.abspath(os.path.join(CC_STAGE_AREA, 'acs_atividade.xlsx')))
    at = pd.read_excel(os.path.abspath(os.path.join(CC_STAGE_AREA, 'at_atividade.xlsx')))

    # Juntando os arquivos
    atividades = pd.concat([acs, at], ignore_index=True)
    atividades = atividades.dropna(subset='codigo_atividade')

    # Salvando os arquivos
    atividades.to_excel(os.path.abspath(os.path.join(CC_UP, 'atividades.xlsx')), index = False)


def processar_atividades():
    ler_salvar_arquivos()
    acs_atividade()
    at_atividade()
    juntar_salvar_arq()