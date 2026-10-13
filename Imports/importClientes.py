import pandas as pd
import os
import sys
import openpyxl
import re
from datetime import datetime


def obter_diretorio_base():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

def limpar_valor(valor):
    if pd.isna(valor) or valor == 'NULL' or valor == '':
        return 'NULL'
    if isinstance(valor, (int, float)):
        return str(int(valor)) if isinstance(valor, float) and valor.is_integer() else str(valor)
    valor_limpo = re.sub(r'[^\d]', '', str(valor))
    return valor_limpo if valor_limpo else 'NULL'

def formatar_valor(valor):
    if pd.isna(valor) or valor == 'NULL' or valor == '':
        return 'NULL'
    
    valor = str(valor).upper()
    
    
    valor = valor.replace("'", "''").replace('"', '') 
    
    return f"'{valor}'"


import pandas as pd

def gerar_insert_cliente(row):
    colunas_numericas = [
        'cliCpfCgc', 'cliRgInsc', 'cliTipoCad', 'cliTipo', 'cliId',
        'cliFatCep', 'cliFone', 'cliCelular', 'cliFatEndNumero',
        'cliDesativa', 'cliLimitCred', 'cliCobCep','cliRotaId','cliUsuLoginId'
    ]
    
    if 'cliCpfCgc' in colunas_numericas:
        colunas_numericas.remove('cliCpfCgc')
    
    row['cliCobCep'] = limpar_valor(row.get('cliFatCep', ''))
    row['cliCpfCgc'] = limpar_valor(row.get('cliCpfCgc', ''))

    campos_cobranca = ['cliCobEnd', 'cliCobBairro', 'cliCobCidade', 'cliCobUf', 'cliCobCep', 'cliCobEndNumero']
    campos_faturamento = ['cliFatEnd', 'cliFatBairro', 'cliFatCidade', 'cliFatUf', 'cliFatCep', 'cliFatEndNumero']

    for cob, fat in zip(campos_cobranca, campos_faturamento):
        row[cob] = row.get(fat, '')

    values = {}
    for col in row:
        if col == 'empId':
            continue
        valor = row[col] if pd.notna(row[col]) else ''
        if col in colunas_numericas:
            values[col] = limpar_valor(valor)
        elif col == 'cliCpfCgc':  
            values[col] = f"'{valor}'" if valor else 'NULL'
        else:
            values[col] = formatar_valor(valor)


    cliNome = str(row.get('cliNome', ''))[:50]
    cliFatEndRef = str(row.get('cliFatEndRef', ''))[:50]
    cliFantasia = str(row.get('cliFantasia', ''))[:40] if pd.notna(row.get('cliFantasia')) else ''
    cliObsGeral = str(row.get('cliObsGeral', '')) if pd.notna(row.get('cliObsGeral')) else ''
    values['cliObsGeral'] = formatar_valor(cliObsGeral)
    values['cliNome'] = formatar_valor(cliNome)
    values['cliFatEndRef'] = formatar_valor(cliFatEndRef)
    values['cliFantasia'] = formatar_valor(cliFantasia)

    columns = ', '.join(values.keys())
    values_str = ', '.join(map(str, values.values()))
    
    insert = f"INSERT INTO cliente({columns}) VALUES({values_str}) GO"
    return insert

from datetime import datetime

def gerar_insert_empresafiltro(row):
    empId = str(row.get('empId', '')).strip()
    cliId = row.get('cliId')

    if not empId or empId.upper() == 'NULL':
        return None
    
    empId_formatado = limpar_valor(empId)
    cliId_formatado = limpar_valor(cliId)

    data_hoje = datetime.now().strftime('%Y-%d-%m')
    data_formatada = f"'{data_hoje}'"

    return f"INSERT INTO empresafiltro (empId, emfTable, emfPkField, emfUsuId, emfPkValue, emfDataOcorrencia) VALUES ({empId_formatado}, 'cliente', 'cliId', '7',{cliId_formatado}, {data_formatada}) GO"


def gerar_sql_inserts(df, mapeamento):
    inserts = []
    inserts_empresafiltro = []

    inserts.append('EXEC SP_MSFOREACHTABLE "ALTER TABLE ? NOCHECK CONSTRAINT ALL" GO')
    inserts.append('SET IDENTITY_INSERT cliente ON GO')
    inserts.append('DBCC TRACEON(460, -1) GO')

    for _, linha_excel in df.iterrows():
        # Cria uma nova linha com os nomes do banco (chaves do mapeamento)
        row = {}
        for campo_banco, coluna_excel in mapeamento.items():
            if campo_banco == 'empId' and not coluna_excel:
                row[campo_banco] = 'NULL'
            else:
                row[campo_banco] = linha_excel.get(coluna_excel, '')


        inserts.append(gerar_insert_cliente(row))

        insert_empfiltro = gerar_insert_empresafiltro(row)
        if insert_empfiltro:
            inserts_empresafiltro.append(insert_empfiltro)

    inserts.append('SET IDENTITY_INSERT cliente OFF GO')
    inserts.append('EXEC SP_MSFOREACHTABLE "ALTER TABLE ? CHECK CONSTRAINT ALL" GO')
    inserts.append('DBCC TRACEOFF(460, -1) GO')

    inserts.append('update cliente set cliente.cliFatCidCodIBGE = municipio.munCodigoIBGE from municipio inner join cliente on cliente.cliFatCidade = municipio.munDescricao inner join uf on municipio.munUfId = uf.ufCodigo and cliente.cliFatUf = uf.ufSigla go')
    inserts.append('update cliente set cliente.cliCobCidCodIBGE = municipio.munCodigoIBGE from municipio inner join cliente on cliente.cliCobCidade = municipio.munDescricao inner join uf on municipio.munUfId = uf.ufCodigo and cliente.cliFatUf = uf.ufSigla go')
    inserts.append("update cliente SET cliContribuinte = 'False' WHERE cliRgInsc IS NULL OR cliRgInsc = '' GO")
    inserts.append("update cliente SET cliRgInsc = 'ISENTO' WHERE cliRgInsc IS NULL OR cliRgInsc = '' GO")
    inserts.append("UPDATE cliente SET clinome = LTRIM(clinome) WHERE clinome LIKE ' %' GO")

    inserts.extend(inserts_empresafiltro)
    
    return inserts

def corrigir_null_no_arquivo(caminho_arquivo):
    with open(caminho_arquivo, 'r', encoding='utf-8') as f:
        conteudo = f.read()

    # Substitui 'NULL' isolado, ou seja, aspas simples + NULL + aspas simples
    # Usando replace simples, pois queremos trocar exatamente isso:
    conteudo_corrigido = conteudo.replace("'NULL'", "NULL")

    with open(caminho_arquivo, 'w', encoding='utf-8') as f:
        f.write(conteudo_corrigido)

def salvar_sql_inserts(arquivo, inserts):
    with open(arquivo, 'w', encoding='utf-8') as f:
        for insert in inserts:
            f.write(insert + '\n')
    corrigir_null_no_arquivo(arquivo)
            
def carregar_excel(caminho, sheet_name=None, fill_value='NULL'):
    if not os.path.exists(caminho):
        raise FileNotFoundError(f"Arquivo Excel '{caminho}' não encontrado.")
    
    try:
        df = pd.read_excel(caminho, sheet_name=sheet_name)
        df = df.fillna(fill_value)
        return df
    except Exception as e:
        raise Exception(f"Erro ao carregar o arquivo Excel: {e}")




