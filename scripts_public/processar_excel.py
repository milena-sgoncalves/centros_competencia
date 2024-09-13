import os
import pandas as pd
import numpy as np
import re

# Função para formatar o CPF
def formatar_cpf(cpf):
    # Verifica se o valor é um número ou está vazio, e se for o caso, retorna um valor vazio
    if pd.isna(cpf) or not isinstance(cpf, str):
        return ''
    # Remove todos os caracteres que não são números
    cpf = re.sub(r'\D', '', cpf)
    # Verifica se o CPF tem 11 dígitos antes de formatar
    if len(cpf) == 11:
        # Formata o CPF para ###.###.###-##
        return f'{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}'
    if len(cpf) == 10:
    # Formata o CPF, adicionando o dígito 0 antes
        return f'0{cpf[:2]}.{cpf[2:5]}.{cpf[5:8]}-{cpf[8:]}'
    else:
        # Retorna o CPF sem alterações se não tiver 11, nem 10 dígitos
        return cpf
    
# Função para formatar o CNPJ
def formatar_cnpj(cnpj):
    # Verifica se o valor é um número ou está vazio, e se for o caso, retorna um valor vazio
    if pd.isna(cnpj) or not isinstance(cnpj, str):
        return ''
    # Remove todos os caracteres que não são números
    cnpj = re.sub(r'\D', '', cnpj)
    
    # Verifica se o CNPJ tem 14 dígitos antes de formatar
    if len(cnpj) == 14:
        # Formata o CNPJ para ##.###.###/####-##
        return f'{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}'
    if len(cnpj) == 13:
        # Formata o CNPJ, adicionando o dígito 0 antes
        return f'0{cnpj[:1]}.{cnpj[1:4]}.{cnpj[4:7]}/{cnpj[7:11]}-{cnpj[11:]}'
    else:
        # Retorna o CNPJ sem alterações se não tiver 14 dígitos
        return cnpj

# Função para remover valores duplicados e remover na's
def remover_duplicados(arquivo, coluna, rm_na=False):
    if rm_na:
        df_rm = arquivo.dropna(subset=[coluna])
        df_unicos = df_rm.drop_duplicates(subset=[coluna])
        return df_unicos
    else:
        df_unicos = arquivo.drop_duplicates(subset=[coluna])
        return df_unicos
    
# Função para remover valores específicos
def remover_valor_especifico(arquivo, coluna, valores_a_remover):
    arquivo[coluna] = arquivo[coluna].str.strip()
    arquivo = arquivo[~arquivo[coluna].isin(valores_a_remover)]
    return arquivo




def processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_data=None, campos_valor=None, campos_string=None,
                    cpf=None, cnpj=None, coluna=None, rm_na=False, rm_valor_especifico=False, coluna_valor = None, valores_a_remover=None):
    # Ler o arquivo Excel
    df = pd.read_excel(arquivo_origem)

    # Selecionar apenas as colunas de interesse
    df_selecionado = df[campos_interesse]

    # Renomear as colunas e definir a nova ordem
    df_renomeado = df_selecionado.rename(columns=novos_nomes_e_ordem)

    # Ajustar campos de data, se fornecidos
    if campos_data:
        for campo in campos_data:
            if campo in df_renomeado.columns:
                df_renomeado[campo] = pd.to_datetime(df_renomeado[campo], format='%d/%m/%Y', errors='coerce')

    # Ajustar campos de string, se fornecidos
    if campos_string:
        for campo in campos_string:
            if campo in df_renomeado.columns:
                df_renomeado[campo] = df_renomeado[campo].astype(str)

    # Ajustar cpf, se fornecidos
    if cpf:
        if cpf in df_renomeado.columns:
            df_renomeado[cpf] = df_renomeado[cpf].apply(formatar_cpf)

    # Ajustar cnpj, se fornecidos
    if cnpj:
        if cnpj in df_renomeado.columns:
            df_renomeado[cnpj] = df_renomeado[cnpj].apply(formatar_cnpj)

    # Remover duplicadas, se fornecido
    if coluna:
        if coluna in df_renomeado.columns:
            df_renomeado = remover_duplicados(df_renomeado, coluna, rm_na)

    # Remover valor específico
    if rm_valor_especifico:
        if coluna_valor in df_renomeado.columns:
            df_renomeado = remover_valor_especifico(df_renomeado, coluna_valor, valores_a_remover)


        
    

    # Reordenar as colunas conforme especificado
    df_final = df_renomeado[list(novos_nomes_e_ordem.values())]

    # Garantir que o diretório de destino existe
    os.makedirs(os.path.dirname(arquivo_destino), exist_ok=True)

    # Verificar se o arquivo de destino está sendo usado e remover se necessário
    if os.path.exists(arquivo_destino):
        os.remove(arquivo_destino)

    # Salvar o arquivo resultante
    with pd.ExcelWriter(arquivo_destino, engine='xlsxwriter') as writer:
        df_final.to_excel(writer, index=False, sheet_name='Sheet1')

        # Acessar o workbook e worksheet
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # Aplicar formatação numérica aos campos de valor
        if campos_valor:
            for coluna in campos_valor:
                if coluna in df_final.columns:
                    # Transformando em string
                    df_final[coluna] = df_final[coluna].astype(str)

                    # Removendo "R$ "
                    df_final[coluna] = df_final[coluna].str.replace(r' R\$ -   ', '', regex=True)

                    # Convertendo para número, se necessário (removendo os pontos de milhar e trocando vírgula por ponto)
                    df_final[coluna] = df_final[coluna].str.replace('.', '', regex=True).str.replace(',', '.').replace('', '0').astype(float)

        # Aplicar formatação de data aos campos de data
        if campos_data:
            format_date = workbook.add_format({'num_format': 'dd/mm/yyyy'})
            for coluna in campos_data:
                if coluna in df_final.columns:
                    col_idx = df_final.columns.get_loc(coluna)
                    worksheet.set_column(col_idx, col_idx, 20, format_date)



        # Definir a largura das colunas
        for i, coluna in enumerate(df_final.columns):
            col_idx = i
            worksheet.set_column(col_idx, col_idx, 20)

# Exemplo de chamada da função
# processar_excel(arquivo_origem, campos_interesse, novos_nomes_e_ordem, arquivo_destino, campos_data, campos_valor)
