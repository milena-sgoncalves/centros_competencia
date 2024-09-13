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
CC_UP = os.path.abspath(os.path.join(ROOT, 'CC_up'))
CC_DATA_RAW = os.path.abspath(os.path.join(ROOT, 'CC_data_raw'))
CC_STAGE_AREA = os.path.abspath(os.path.join(ROOT, 'CC_stage_area'))
sys.path.append(SCRIPTS_PUBLIC)

from processar_excel import processar_excel

def afcct_pdi():
    # Lendo o arquivo
    projetos_afcct_pdi = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS TECNICO.xlsx')),
                                        sheet_name = 'AFCCT - Projetos PDI')
    projetos_afcct_pdi['tipo_projeto'] = 'pdi'
    projetos_afcct_pdi['modelo'] = 'afcct'

    # Gerando a planilha
    projetos_afcct_pdi.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'projetos_afcct_pdi.xlsx')), index = False)
        
    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_stage_area')
    nome_arquivo = 'projetos_afcct_pdi.xlsx'
    arquivo_origem = os.path.join(origem, nome_arquivo)
    arquivo_destino = os.path.join(destino, nome_arquivo)

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
            'modelo',
            'tipo_projeto',
            'Centro de Competência',
            'Título do Projeto de PD&I de Ampliação e Fortalecimento de Competências Científicas e Tecnológicas',
            'Código do projeto',
            'Resumo do projeto',
            'Descrição Pública',
            'Há envolvimento de profissionais/pesquisadores das empresas da AT participando do projeto?',
            'Nome do(s) profissional(is) da(s) empresa(s) da AT envolvido(s)',
            'Nome da empresa da AT a qual o(s) profissional(is) é(são) vinculado(s)',
            'TRL Inicial',
            'TRL Final',
            'Status atual do projeto',
            'Resultados executados (até o momento)',
            'Valor total planejado do projeto (todas as fontes)',
            'Valor aportado pela EMBRAPII (até o momento)',
            'Valor aportado pela AT (até o momento)',
            'Valor aportado por Outras Fontes (até o momento)',
            ' Data real de início',
            ' Data prevista de término',
            ' Data real de término',
            'Observações',
        ]

    novos_nomes_e_ordem = {
            'Código do projeto': 'codigo_projeto',
            'Centro de Competência': 'centro_competencia',
            'tipo_projeto': 'tipo_projeto',
            'modelo': 'modelo',
            'Título do Projeto de PD&I de Ampliação e Fortalecimento de Competências Científicas e Tecnológicas': 'titulo_projeto',
            'Descrição Pública': 'descriacao_publica',
            'Status atual do projeto': 'status_projeto',
            'Valor total planejado do projeto (todas as fontes)': 'valor_total',
            'Valor aportado pela EMBRAPII (até o momento)': 'valor_embrapii',
            'Valor aportado pela AT (até o momento)': 'valor_at',
            'Valor aportado por Outras Fontes (até o momento)': 'valor_outras_fontes',
            ' Data real de início': 'data_inicio_real',
            ' Data real de término': 'data_termino_real',
            'Observações': 'observacoes',
            'Resumo do projeto': 'resumo_projeto',
            'Há envolvimento de profissionais/pesquisadores das empresas da AT participando do projeto?': 'prof_empresas_envolvidos',
            'Nome do(s) profissional(is) da(s) empresa(s) da AT envolvido(s)': 'nome_prof_empresas_envolvidos',
            'Nome da empresa da AT a qual o(s) profissional(is) é(são) vinculado(s)': 'empresa_envolvida',
            'TRL Inicial': 'trl_inicial',
            'TRL Final': 'trl_final',
            'Resultados executados (até o momento)': 'resultados',
            ' Data prevista de término': 'data_termino_prevista'
        }

    # Campos de data e valor
    campos_data = ['data_inicio_real', 'data_termino_real', 'data_termino_prevista']
    campos_valor = ['valor_total', 'valor_embrapii', 'valor_at', 'valor_outras_fontes']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_data, campos_valor)


def acs_pdi():
    # Lendo o arquivo
    projetos_acs_pdi = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS TECNICO.xlsx')),
                                        sheet_name = 'ACS - Projetos PDI')
    projetos_acs_pdi['tipo_projeto'] = 'pdi'
    projetos_acs_pdi['modelo'] = 'acs'

    # Gerando a planilha
    projetos_acs_pdi.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'projetos_acs_pdi.xlsx')), index = False)
        
    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_stage_area')
    nome_arquivo = 'projetos_acs_pdi.xlsx'
    arquivo_origem = os.path.join(origem, nome_arquivo)
    arquivo_destino = os.path.join(destino, nome_arquivo)

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
            'modelo',
            'tipo_projeto',
            'Centro de Competência',
            'Título do Projeto PD&I de Atração e Criação de Startup',
            'Código do projeto',
            'Projeto executado no ambiente de inovação aberta?',
            'Projeto desenvolvido de forma cooperada (com participação de mais de 1 empresa)?',
            'Existiu envolvimento de empresas de base tecnológica?',
            'Objetivo',
            'Descrição Pública',
            'Status atual do projeto',
            'Valor total planejado do projeto (todas as fontes)',
            'Valor executado pela EMBRAPII (até o momento)',
            'Valor executado pela AT (até o momento)',
            'Valor executado por Outras Fontes (até o momento)',
            'Valor executado por empresas (fora dos aportes da AT) (até o momento)',
            ' Data real de início',
            ' Data real de término',
            'Quantidade de empresas de base tecnológica envolvidas',
            'Observações',
        ]

    novos_nomes_e_ordem = {
            'Código do projeto': 'codigo_projeto',
            'Centro de Competência': 'centro_competencia',
            'tipo_projeto': 'tipo_projeto',
            'modelo': 'modelo',
            'Título do Projeto PD&I de Atração e Criação de Startup': 'titulo_projeto',
            'Objetivo': 'objetivo_projeto',
            'Descrição Pública': 'descricao_publica',
            'Status atual do projeto': 'status_projeto',
            'Valor total planejado do projeto (todas as fontes)': 'valor_total',
            'Valor executado pela EMBRAPII (até o momento)': 'valor_embrapii',
            'Valor executado pela AT (até o momento)': 'valor_at',
            'Valor executado por Outras Fontes (até o momento)': 'valor_outras_fontes',
            'Valor executado por empresas (fora dos aportes da AT) (até o momento)': 'valor_empresas',
            ' Data real de início': 'data_inicio_real',
            ' Data real de término': 'data_termino_real',
            'Observações': 'observacoes',
            'Projeto executado no ambiente de inovação aberta?': 'inovacao_aberta',
            'Projeto desenvolvido de forma cooperada (com participação de mais de 1 empresa)?': 'cooperativo',
            'Existiu envolvimento de empresas de base tecnológica?': 'empresa_tecnologica',
            'Quantidade de empresas de base tecnológica envolvidas': 'quant_empresa_tecnologica'
        }

    # Campos de data e valor
    campos_data = ['data_inicio_real', 'data_termino_real']
    campos_valor = ['valor_total', 'valor_embrapii', 'valor_at', 'valor_outras_fontes', 'valor_empresas']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_data, campos_valor)



def afcct_formacao():
    # Lendo o arquivo
    projetos_afcct_formacao = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS TECNICO.xlsx')),
                                        sheet_name = 'AFCCT - Projetos Formação')
    projetos_afcct_formacao['tipo_projeto'] = 'formacao'
    projetos_afcct_formacao['modelo'] = 'afcct'

    # Gerando a planilha
    projetos_afcct_formacao.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'projetos_afcct_formacao.xlsx')), index = False)
        
    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_stage_area')
    nome_arquivo = 'projetos_afcct_formacao.xlsx'
    arquivo_origem = os.path.join(origem, nome_arquivo)
    arquivo_destino = os.path.join(destino, nome_arquivo)

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
            'modelo',
            'tipo_projeto',
            'Centro de Competência',
            'Título do Projeto de Formação e Capacitação (exclusiva aos colaboradores do CC)',
            'Código do projeto',
            'Objetivo',
            'Descrição Pública',
            'Status atual do projeto',
            'Valor total planejado do projeto (todas as fontes)',
            'Valor aportado pela EMBRAPII (até o momento)',
            'Valor aportado pela AT (até o momento)',
            'Valor aportado por Outras Fontes (até o momento)',
            ' Data real de início',
            ' Data real de término',
            'Número de profissionais que iniciaram a formação',
            'Número de profissionais que concluiram a formação',
            'Qual tipo de formação será realizada?',
            'Onde a formação será realizada?',
            'Vínculo da formação à(s) linha(s) temática(s) de pesquisa do CC',
            'Observações',
        ]

    novos_nomes_e_ordem = {
            'Código do projeto': 'codigo_projeto',
            'Centro de Competência': 'centro_competencia',
            'tipo_projeto': 'tipo_projeto',
            'modelo': 'modelo',
            'Título do Projeto de Formação e Capacitação (exclusiva aos colaboradores do CC)': 'titulo_projeto',
            'Objetivo': 'objetivo_projeto',
            'Descrição Pública': 'descricao_publica',
            'Status atual do projeto': 'status_projeto',
            'Valor total planejado do projeto (todas as fontes)': 'valor_total',
            'Valor aportado pela EMBRAPII (até o momento)': 'valor_embrapii',
            'Valor aportado pela AT (até o momento)': 'valor_at',
            'Valor aportado por Outras Fontes (até o momento)': 'valor_outras_fontes',
            ' Data real de início': 'data_inicio_real',
            ' Data real de término': 'data_termino_real',
            'Observações': 'observacoes',
            'Número de profissionais que iniciaram a formação': 'num_prof_ingressantes',
            'Número de profissionais que concluiram a formação': 'num_prof_concluintes',
            'Qual tipo de formação será realizada?': 'tipo_formacao',
            'Onde a formação será realizada?': 'local_formacao',
            'Vínculo da formação à(s) linha(s) temática(s) de pesquisa do CC': 'vinculo_linha_tematica',
        }

    # Campos de data e valor
    campos_data = ['data_inicio_real', 'data_termino_real']
    campos_valor = ['valor_total', 'valor_embrapii', 'valor_at', 'valor_outras_fontes']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_data, campos_valor)


def acs_formacao():
    # Lendo o arquivo
    projetos_acs_formacao = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS TECNICO.xlsx')),
                                        sheet_name = 'ACS - Projetos de Formação')
    projetos_acs_formacao['tipo_projeto'] = 'formacao'
    projetos_acs_formacao['modelo'] = 'acs'

    # Gerando a planilha
    projetos_acs_formacao.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'projetos_acs_formacao.xlsx')), index = False)
        
    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_stage_area')
    nome_arquivo = 'projetos_acs_formacao.xlsx'
    arquivo_origem = os.path.join(origem, nome_arquivo)
    arquivo_destino = os.path.join(destino, nome_arquivo)

    # Campos de interesse e novos nomes das colunas
    campos_interesse = [
            'modelo',
            'tipo_projeto',
            'Centro de Competência',
            'Título do Projeto de Formação e Capacitação (no contexto da ação ACS)',
            'Código do projeto',
            'Objetivo',
            'Descrição Pública',
            'Status atual do projeto',
            'Valor total planejado do projeto (todas as fontes)',
            'Valor executado pela EMBRAPII (até o momento)',
            'Valor executado pela AT (até o momento)',
            'Valor executado por Outras Fontes (até o momento)',
            ' Data real de início',
            ' Data real de término',
            'Número de profissionais que iniciaram a formação',
            'Número de profissionais que concluiram a formação',
            'Qual tipo de formação será realizada?',
            'Onde a formação será realizada?',
            'Vínculo da formação à(s) linha(s) temática(s) de pesquisa do CC',
            'Observações',
        ]

    novos_nomes_e_ordem = {
            'Código do projeto': 'codigo_projeto',
            'Centro de Competência': 'centro_competencia',
            'tipo_projeto': 'tipo_projeto',
            'modelo': 'modelo',
            'Título do Projeto de Formação e Capacitação (no contexto da ação ACS)': 'titulo_projeto',
            'Objetivo': 'objetivo_projeto',
            'Descrição Pública': 'descricao_publica',
            'Status atual do projeto': 'status_projeto',
            'Valor total planejado do projeto (todas as fontes)': 'valor_total',
            'Valor executado pela EMBRAPII (até o momento)': 'valor_embrapii',
            'Valor executado pela AT (até o momento)': 'valor_at',
            'Valor executado por Outras Fontes (até o momento)': 'valor_outras_fontes',
            ' Data real de início': 'data_inicio_real',
            ' Data real de término': 'data_termino_real',
            'Observações': 'observacoes',
            'Número de profissionais que iniciaram a formação': 'num_prof_ingressantes',
            'Número de profissionais que concluiram a formação': 'num_prof_concluintes',
            'Qual tipo de formação será realizada?': 'tipo_formacao',
            'Onde a formação será realizada?': 'local_formacao',
            'Vínculo da formação à(s) linha(s) temática(s) de pesquisa do CC': 'vinculo_linha_tematica',
        }

    # Campos de data e valor
    campos_data = ['data_inicio_real', 'data_termino_real']
    campos_valor = ['valor_total', 'valor_embrapii', 'valor_at', 'valor_outras_fontes']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_data, campos_valor)


def fcrh():
    # Lendo o arquivo
    projetos_fcrh = pd.read_excel(os.path.abspath(os.path.join(CC_COPY, 'BI - DADOS TECNICO.xlsx')),
                                        sheet_name = 'FCRH - Projetos')
    projetos_fcrh['tipo_projeto'] = 'formacao'
    projetos_fcrh['modelo'] = 'fcrh'

    # Gerando a planilha
    projetos_fcrh.to_excel(os.path.abspath(os.path.join(CC_DATA_RAW, 'projetos_fcrh.xlsx')), index = False)
        
    # Definições dos caminhos e nomes de arquivos
    origem = os.path.join(ROOT, 'CC_data_raw')
    destino = os.path.join(ROOT, 'CC_stage_area')
    nome_arquivo = 'projetos_fcrh.xlsx'
    arquivo_origem = os.path.join(origem, nome_arquivo)
    arquivo_destino = os.path.join(destino, nome_arquivo)

        # Campos de interesse e novos nomes das colunas
    campos_interesse = [
            'modelo',
            'tipo_projeto',
            'Centro de Competência',
            'Título do Projeto de Formação e Capacitação de RH para PD&I',
            'Código do projeto',
            'Objetivo',
            'Descrição Pública',
            'Ementa detalhada',
            'Carga horária total',
            'Público alvo do curso',
            'Status atual do projeto',
            'Valor total planejado do projeto (todas as fontes)',
            'Valor executado pela EMBRAPII (até o momento)',
            'Valor executado pela AT (até o momento)',
            'Valor executado por Outras Fontes (até o momento)',
            ' Data real de início',
            ' Data real de término',
            'Número de profissionais que iniciaram a formação',
            'Número de profissionais que concluiram a formação',
            'Qual tipo de formação será realizada?',
            'Onde a formação será realizada?',
            'Vínculo da formação à(s) linha(s) temática(s) de pesquisa do CC',
            'Quantidade de empresas associadas a AT que participaram da formação',
            'Observações',
        ]

    novos_nomes_e_ordem = {
            'Código do projeto': 'codigo_projeto',
            'Centro de Competência': 'centro_competencia',
            'tipo_projeto': 'tipo_projeto',
            'modelo': 'modelo',
            'Título do Projeto de Formação e Capacitação de RH para PD&I': 'titulo_projeto',
            'Objetivo': 'objetivo_projeto',
            'Descrição Pública': 'descricao_publica',
            'Status atual do projeto': 'status_projeto',
            'Valor total planejado do projeto (todas as fontes)': 'valor_total',
            'Valor executado pela EMBRAPII (até o momento)': 'valor_embrapii',
            'Valor executado pela AT (até o momento)': 'valor_at',
            'Valor executado por Outras Fontes (até o momento)': 'valor_outras_fontes',
            ' Data real de início': 'data_inicio_real',
            ' Data real de término': 'data_termino_real',
            'Observações': 'observacoes',
            'Número de profissionais que iniciaram a formação': 'num_prof_ingressantes',
            'Número de profissionais que concluiram a formação': 'num_prof_concluintes',
            'Qual tipo de formação será realizada?': 'tipo_formacao',
            'Onde a formação será realizada?': 'local_formacao',
            'Vínculo da formação à(s) linha(s) temática(s) de pesquisa do CC': 'vinculo_linha_tematica',
            'Ementa detalhada': 'ementa',
            'Carga horária total': 'carga_horaria',
            'Público alvo do curso': 'publico_alvo',
            'Quantidade de empresas associadas a AT que participaram da formação': 'quant_empresas_at',
        }

    # Campos de data e valor
    campos_data = ['data_inicio_real', 'data_termino_real']
    campos_valor = ['valor_total', 'valor_embrapii', 'valor_at', 'valor_outras_fontes']

    processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_data, campos_valor)


def projetos():
    # Lendo os arquivos
    afcct_pdi = pd.read_excel(os.path.abspath(os.path.join(CC_STAGE_AREA, 'projetos_afcct_pdi.xlsx')))
    acs_pdi = pd.read_excel(os.path.abspath(os.path.join(CC_STAGE_AREA, 'projetos_acs_pdi.xlsx')))
    afcct_formacao = pd.read_excel(os.path.abspath(os.path.join(CC_STAGE_AREA, 'projetos_afcct_formacao.xlsx')))
    acs_formacao = pd.read_excel(os.path.abspath(os.path.join(CC_STAGE_AREA, 'projetos_acs_formacao.xlsx')))
    fcrh = pd.read_excel(os.path.abspath(os.path.join(CC_STAGE_AREA, 'projetos_fcrh.xlsx')))

    # Juntando os arquivos
    projetos = pd.concat([afcct_pdi, acs_pdi, afcct_formacao, acs_formacao, fcrh], ignore_index=True)

    campos_interesse = [
            'codigo_projeto',
            'centro_competencia',
            'tipo_projeto',
            'modelo',
            'titulo_projeto',
            'objetivo_projeto',
            'descricao_publica',
            'status_projeto',
            'valor_total',
            'valor_embrapii',
            'valor_at',
            'valor_outras_fontes',
            'valor_empresas',
            'data_inicio_real',
            'data_termino_real',
            'observacoes',
        ]

    projetos = projetos[campos_interesse]
    projetos = projetos.dropna(subset=['codigo_projeto'])
    projetos.to_excel(os.path.abspath(os.path.join(CC_UP, 'projetos.xlsx')), index = False)



def projetos_pdi():
    # Lendo os arquivos
    afcct = pd.read_excel(os.path.abspath(os.path.join(CC_STAGE_AREA, 'projetos_afcct_pdi.xlsx')))
    acs = pd.read_excel(os.path.abspath(os.path.join(CC_STAGE_AREA, 'projetos_acs_pdi.xlsx')))

    # Juntando os arquivos
    projetos_pdi = pd.concat([afcct, acs], ignore_index=True)

    campos_interesse = [
            'codigo_projeto',
            'modelo',
            'resumo_projeto',
            'prof_empresas_envolvidos',
            'nome_prof_empresas_envolvidos',
            'empresa_envolvida',
            'trl_inicial',
            'trl_final',
            'resultados',
            'data_termino_prevista',
            'inovacao_aberta',
            'cooperativo',
            'empresa_tecnologica',
            'quant_empresa_tecnologica',
        ]

    projetos_pdi = projetos_pdi[campos_interesse]
    projetos_pdi = projetos_pdi.dropna(subset=['codigo_projeto'])
    projetos_pdi.to_excel(os.path.abspath(os.path.join(CC_UP, 'projetos_pdi.xlsx')), index = False)


def projetos_formacao():
    # Lendo os arquivos
    afcct = pd.read_excel(os.path.abspath(os.path.join(CC_STAGE_AREA, 'projetos_afcct_formacao.xlsx')))
    acs = pd.read_excel(os.path.abspath(os.path.join(CC_STAGE_AREA, 'projetos_acs_formacao.xlsx')))
    fcrh = pd.read_excel(os.path.abspath(os.path.join(CC_STAGE_AREA, 'projetos_fcrh.xlsx')))

    # Juntando os arquivos
    projetos_formacao = pd.concat([afcct, acs, fcrh], ignore_index=True)

    campos_interesse = [
            'codigo_projeto',
            'modelo',
            'num_prof_ingressantes',
            'num_prof_concluintes',
            'tipo_formacao',
            'local_formacao',
            'vinculo_linha_tematica',
            'ementa',
            'carga_horaria',
            'publico_alvo',
            'quant_empresas_at',
        ]

    projetos_formacao = projetos_formacao[campos_interesse]
    projetos_formacao = projetos_formacao.dropna(subset=['codigo_projeto'])
    projetos_formacao.to_excel(os.path.abspath(os.path.join(CC_UP, 'projetos_formacao.xlsx')), index = False)

def processar_projetos():
    afcct_pdi()
    acs_pdi()
    afcct_formacao()
    acs_formacao()
    fcrh()
    projetos()
    projetos_pdi()
    projetos_formacao()