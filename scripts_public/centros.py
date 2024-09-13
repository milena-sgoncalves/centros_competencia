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

def processar_centros_competencia():
    # Lendo o arquivo
    centros_competencia = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS TECNICO.xlsx')),
                                        sheet_name = 'CENTROS - Dados Gerais')
    centros_competencia.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'centros_competencia.xlsx')), index = False)
    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_up')
    nome_arquivo = 'centros_competencia.xlsx'
    arquivo_origem = os.path.join(origem, nome_arquivo)
    arquivo_destino = os.path.join(destino, 'centros_competencia.xlsx')

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
        'Centro de Competência',
        'Nome fantasia',
        'Tipo de Instituição',
        'Chamada de Credenciamento',
        'Cidade',
        'Estado',
        'Coordenador do CC',
        'Email do Coordenador do CC',
        'Telefone do Coordenador do CC',
        'Gerente Executivo do CC',
        'Email do Gerente Executivo do CC',
        'Telefone do Gerente Executivo do CC',
        'Responsável Comunicação',
        'Email do Responsável de Comunicação',
        'Telefone do Responsável de Comunicação',
        'Responsável EMBRAPII',
        'Responsável Institucional',
        'Telefone do Responsável Institucional',
        'Nome do responsável pelo sistema tickets no CC',
        'Email cadastrado no tickets',
        'Número do Termo de Cooperação',
        'Data assinatura do Termo de Cooperação',
        'Linhas de Pesquisa',
        'Status do Credenciamento',
    ]

    novos_nomes_e_ordem = {
        'Centro de Competência': 'centro_competencia',
        'Nome fantasia': 'nome_fantasia',
        'Tipo de Instituição': 'tipo_instituicao',
        'Chamada de Credenciamento': 'chamada_credenciamento',
        'Cidade': 'cidade',
        'Estado': 'estado',
        'Coordenador do CC': 'coordenador',
        'Email do Coordenador do CC': 'email_coordenador',
        'Telefone do Coordenador do CC': 'telefone_coordenador',
        'Gerente Executivo do CC': 'gerente_executivo',
        'Email do Gerente Executivo do CC': 'email_gerente',
        'Telefone do Gerente Executivo do CC': 'telefone_gerente',
        'Responsável Comunicação': 'responsavel_comunicacao',
        'Email do Responsável de Comunicação': 'email_responsavel_comunicacao',
        'Telefone do Responsável de Comunicação': 'telefone_responsavel_comunicacao',
        'Responsável EMBRAPII': 'responsavel_embrapii',
        'Responsável Institucional': 'responsavel_institucional',
        'Telefone do Responsável Institucional': 'telefone_responsavel_institucional',
        'Nome do responsável pelo sistema tickets no CC': 'responsavel_tickets',
        'Email cadastrado no tickets': 'email_tickets',
        'Número do Termo de Cooperação': 'numero_termo_cooperacao',
        'Data assinatura do Termo de Cooperação': 'data_assinatura',
        'Linhas de Pesquisa': 'linhas_pesquisa',
        'Status do Credenciamento': 'status_credenciamento'
    }

    # Campos de data e valor
    campos_data = ['data_assinatura']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_data)

#Executar função
if __name__ == "__main__":
    processar_centros_competencia()