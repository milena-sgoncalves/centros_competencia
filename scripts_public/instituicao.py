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
    nac = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS FINANCEIROS.xlsx')),
                        sheet_name = 'Cooperação Nacional')
    inter = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS FINANCEIROS.xlsx')),
                        sheet_name = 'Cooperação Internacional')

    # Salvando os arquivos
    nac.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'coop_nac.xlsx')), index = False)
    inter.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'coop_inter.xlsx')), index = False)

def nacional():
    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_stage_area')
    nome_arquivo = 'coop_nac.xlsx'
    arquivo_origem = os.path.join(origem, nome_arquivo)
    arquivo_destino = os.path.join(destino, nome_arquivo)

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'Nome da Instituição',
        'Data de início da cooperação',
        'Data de encerramento da cooperação',
        'Código(s) do(s) projeto(s) realizado(s) de forma cooperada',
        'Objetivo da cooperação',
        'Resultados e atividades da cooperação até o momento',
        'Observações',
    ]

    novos_nomes_e_ordem = {
        'Código(s) do(s) projeto(s) realizado(s) de forma cooperada': 'codigo_projeto',
        'Nome da Instituição': 'instituicao',
        'Data de início da cooperação': 'data_inicio',
        'Data de encerramento da cooperação': 'data_encerramento',
        'Objetivo da cooperação': 'objetivo',
        'Resultados e atividades da cooperação até o momento': 'resultados',
        'Observações': 'observacoes',
    }

    # Campos de data e valor
    campos_data = ['data_inicio', 'data_encerramento']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_data)

def emp_nac():
    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_stage_area')
    arquivo_origem = os.path.join(origem, 'coop_nac.xlsx')
    arquivo_destino = os.path.join(destino, 'emp_nac.xlsx')

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'Nome da Instituição',
        'Nome do contato na Instituição',
        'Cargo do contato na Instituição',
        'Email do contato na Instituição',
    ]

    novos_nomes_e_ordem = {
        'Nome da Instituição': 'instituicao',
        'Nome do contato na Instituição': 'contato',
        'Cargo do contato na Instituição': 'cargo_contato',
        'Email do contato na Instituição': 'email_contato',
    }

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino)


def internacional():
    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_stage_area')
    nome_arquivo = 'coop_inter.xlsx'
    arquivo_origem = os.path.join(origem, nome_arquivo)
    arquivo_destino = os.path.join(destino, nome_arquivo)

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'Nome da Instituição',
        'Data de início da cooperação',
        'Data de encerramento da cooperação',
        'Código(s) do(s) projeto(s) realizado(s) de forma cooperada',
        'Objetivo da cooperação',
        'Resultados e atividades da cooperação até o momento',
        'Observações',
    ]

    novos_nomes_e_ordem = {
        'Código(s) do(s) projeto(s) realizado(s) de forma cooperada': 'codigo_projeto',
        'Nome da Instituição': 'instituicao',
        'Data de início da cooperação': 'data_inicio',
        'Data de encerramento da cooperação': 'data_encerramento',
        'Objetivo da cooperação': 'objetivo',
        'Resultados e atividades da cooperação até o momento': 'resultados',
        'Observações': 'observacoes',
    }

    campos_data = ['data_inicio', 'data_encerramento']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_data)

def emp_inter():
    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_stage_area')
    arquivo_origem = os.path.join(origem, 'coop_inter.xlsx')
    arquivo_destino = os.path.join(destino, 'emp_inter.xlsx')

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'Nome da Instituição',
        'País',
        'Nome do contato na Instituição',
        'Cargo do contato na Instituição',
        'Email do contato na Instituição',
    ]

    novos_nomes_e_ordem = {
        'Nome da Instituição': 'instituicao',
        'País': 'pais',
        'Nome do contato na Instituição': 'contato',
        'Cargo do contato na Instituição': 'cargo_contato',
        'Email do contato na Instituição': 'email_contato',
    }

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino)


def juntar_salvar_arq(lista_arquivos, nome_arquivo):
    # Lista para armazenar os DataFrames carregados
    dataframes = []

    # Carregando cada arquivo Excel e adicionando à lista de DataFrames
    for arquivo in lista_arquivos:
        caminho_arquivo = os.path.abspath(os.path.join(CC_STAGE_AREA, f'{arquivo}.xlsx'))
        df = pd.read_excel(caminho_arquivo)
        dataframes.append(df)

    # Juntando todos os DataFrames
    arquivos_juntos = pd.concat(dataframes, ignore_index=True)

    # Salvando o DataFrame concatenado em um arquivo Excel
    caminho_destino = os.path.abspath(os.path.join(CC_UP, nome_arquivo))
    arquivos_juntos.to_excel(caminho_destino, index=False)


def remover_duplicados(arquivo, coluna, rm_na=False):
    caminho_arquivo = os.path.abspath(os.path.join(CC_UP, arquivo))
    df = pd.read_excel(caminho_arquivo)
    if rm_na:
        df_rm = df.dropna(subset=[coluna])
        df_unicos = df_rm.drop_duplicates(subset=[coluna])
        df_unicos.to_excel(caminho_arquivo, index=False)
    else:
        df_unicos = df.drop_duplicates(subset=[coluna])
        df_unicos.to_excel(caminho_arquivo, index=False)



def processar_instituicao():
    ler_salvar_arquivos()
    nacional()
    internacional()
    emp_inter()
    emp_nac()
    juntar_salvar_arq(['coop_inter', 'coop_nac'], 'projeto_instituicao.xlsx')
    juntar_salvar_arq(['emp_inter', 'emp_nac'], 'instituicao.xlsx')
    remover_duplicados('instituicao.xlsx', 'instituicao', rm_na=True)

processar_instituicao()