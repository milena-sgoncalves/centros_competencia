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
    afcct = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS TECNICO.xlsx')),
                            sheet_name = 'AFCCT - Prof Capacitados')
    acs = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS TECNICO.xlsx')),
                        sheet_name = 'ACS - Prof Capacitados')
    fcrh = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS TECNICO.xlsx')),
                        sheet_name = 'FCRH - Prof Capacitados')

    # Salvando os arquivos
    afcct.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'afcct_formacao.xlsx')), index = False)
    acs.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'acs_formacao.xlsx')), index = False)
    fcrh.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'fcrh_formacao.xlsx')), index = False)

def formacao(nome_arquivo):
    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_stage_area')
    arquivo_origem = os.path.join(origem, nome_arquivo)
    arquivo_destino = os.path.join(destino, nome_arquivo)

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'CPF',
        'Código do projeto de formação e capacitação onde o profissional foi formado / capacitado',
        'Data de conclusão da formação',
        'Certificado emitido?',
        'Link para acesso aos certificados',
        'Observação',
    ]

    novos_nomes_e_ordem = {
        'Código do projeto de formação e capacitação onde o profissional foi formado / capacitado': 'codigo_projeto',
        'CPF': 'cpf',
        'Data de conclusão da formação': 'data_conclusao',
        'Certificado emitido?': 'certificado',
        'Link para acesso aos certificados': 'link_certificado',
        'Observação': 'observacoes',
    }

    # Campos de data e valor
    campos_data = ['data_conclusao']
    campos_string = ['cpf']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_data, campos_string=campos_string,
                    cpf = 'cpf')

def formacao_fcrh():
    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_stage_area')
    arquivo_origem = os.path.join(origem, 'fcrh_formacao.xlsx')
    arquivo_destino = os.path.join(destino, 'fcrh_formacao.xlsx')

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'CPF',
        'Código do projeto onde o profissional foi formado / capacitado',
        'Data de conclusão da formação',
        'Certificado emitido?',
        'Link para acesso aos certificados',
        'Observação',
    ]

    novos_nomes_e_ordem = {
        'Código do projeto onde o profissional foi formado / capacitado': 'codigo_projeto',
        'CPF': 'cpf',
        'Data de conclusão da formação': 'data_conclusao',
        'Certificado emitido?': 'certificado',
        'Link para acesso aos certificados': 'link_certificado',
        'Observação': 'observacoes',
    }

    # Campos de data e valor
    campos_data = ['data_conclusao']
    campos_string = ['cpf']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_data, campos_string=campos_string,
                    cpf = 'cpf')


def prof(nome_arquivo_origem, nome_arquivo_destino):
    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_stage_area')
    arquivo_origem = os.path.join(origem, nome_arquivo_origem)
    arquivo_destino = os.path.join(destino, nome_arquivo_destino)

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'Nome Completo',
        'CPF',
        'Email',
        'Vínculo deste profissional',
    ]

    novos_nomes_e_ordem = {
        'CPF': 'cpf',
        'Nome Completo': 'nome',
        'Email': 'email',
        'Vínculo deste profissional': 'vinculo',
    }

    campos_string = ['cpf']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_string=campos_string,
                    cpf = 'cpf', coluna = 'cpf', rm_na = True)


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
    arquivos_juntos = arquivos_juntos.dropna(how='all')

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



def processar_formacao():
    ler_salvar_arquivos()
    formacao('afcct_formacao.xlsx')
    formacao('acs_formacao.xlsx')
    formacao_fcrh()
    prof('afcct_formacao.xlsx', 'afcct_prof.xlsx')
    prof('acs_formacao.xlsx', 'acs_prof.xlsx')
    prof('fcrh_formacao.xlsx', 'fcrh_prof.xlsx')
    juntar_salvar_arq(['afcct_formacao', 'acs_formacao', 'fcrh_formacao'], 'formacao_prof.xlsx')
    juntar_salvar_arq(['afcct_prof', 'acs_prof', 'fcrh_prof'], 'prof_capacitados.xlsx')
    remover_duplicados('prof_capacitados.xlsx', 'cpf', rm_na=True)