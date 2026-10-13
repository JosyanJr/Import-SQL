import pandas as pd
import os
import sys
from datetime import datetime

def obter_diretorio_base():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

diretorio_base = obter_diretorio_base()

# ================= Funções de formatação =================

def formatar_valor(valor):
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    if valor is None:
        return 'NULL'
    return str(valor)

def nome_valido(valor):
    if not valor or str(valor).strip().upper() in ('NULL','', 'NAN'):
        return False
    return True

def formatar_valor_com_aspas(valor):
    if valor is None:
        return 'NULL'
    if isinstance(valor, str) and valor.strip().lower() in ('null', 'nan', ''):
        return 'NULL'
    # Tenta converter string numerica com .0 para inteiro
    if isinstance(valor, str):
        try:
            # Se for algo como "789.0" ou "123.0"
            if valor.strip().endswith('.0'):
                return str(int(float(valor)))
        except (ValueError, TypeError):
            pass
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return f"'{valor}'"

def to_int(valor):
    try:
        return int(valor)
    except (ValueError, TypeError):
        return 'NULL'

def nome_valido(valor):
    if not valor or str(valor).strip().lower() in ('null', 'nan', ''):
        return False
    return True

def converter_str_para_int_se_possivel(valor):
    """Converte string tipo '123.0' ou '123' para inteiro, 
    ou retorna a string original se nao for numerico."""
    if valor is None:
        return 'NULL'
    s = str(valor).strip()
    if s.upper() in ('NULL', 'NAN', ''):
        return 'NULL'
    try:
        # Tenta converter para float, depois int se for .0
        f = float(s)
        if f == int(f):
            return str(int(f))
        return s
    except ValueError:
        return s

# ================= Funções para gerar inserts =================

def gerar_insert_grupo_prod(row):
    gdp_nome = formatar_valor_com_aspas(row['gdpNome'])
    desativa = 0 
    columns_values = {'gdpNome': gdp_nome, 'gdpdesativa': desativa}
    columns = ', '.join(columns_values.keys())
    values = ', '.join(map(str, columns_values.values()))
    return f"INSERT INTO grupoProd ({columns}) VALUES ({values}) GO"

def gerar_insert_fabricante(row):
    fabNome = formatar_valor_com_aspas(row['fabNome'])
    fabAtivo = 0 
    columns_values = {'fabNome': fabNome, 'fabAtivo': fabAtivo}
    columns = ', '.join(columns_values.keys())
    values = ', '.join(map(str, columns_values.values()))
    return f"INSERT INTO fabricante ({columns}) VALUES ({values}) GO"

def gerar_insert_classe(row):
    classeNome = formatar_valor_com_aspas(row['pclDescricao'])
    classeAtivo = 0 
    columns_values = {'pclDescricao': classeNome, 'pclDesativa': classeAtivo}
    columns = ', '.join(columns_values.keys())
    values = ', '.join(map(str, columns_values.values()))
    return f"INSERT INTO produtoClasse ({columns}) VALUES ({values}) GO"

def gerar_insert_subgrupo_prod(row):
    sgp_nome = formatar_valor_com_aspas(row.get('proSubGrupo'))
    gdp_nome = formatar_valor_com_aspas(row.get('gdpNome'))
    desativa = 0 

    columns_values = {
        'sgpNome': sgp_nome,
        'sgpIdGdp': 'NULL',
        'sgpDesativa': desativa,
        'gdpNome_Temp': gdp_nome
    }

    columns = ', '.join(columns_values.keys())
    values = ', '.join(map(str, columns_values.values()))

    return f"INSERT INTO subGrupoProd ({columns}) VALUES ({values}) GO"

def gerar_insert_produto(row):
    desc = str(row['proDescricao']).upper()
    desc = desc.replace("'", "").replace('"', '')
    ncm = ''.join(filter(str.isdigit, str(row.get('zzz_proCodigoNCM', ''))))[:8]
    ncm = 'NULL' if not ncm else f"'{ncm}'"

    columns_values = {
        'proid': to_int(row.get('proid')),
        'proDescricao': formatar_valor_com_aspas(desc[:96]),
        'proDescPdv': formatar_valor_com_aspas(desc[:29]),
        'zzz_proCodigoNCM': ncm,
        'proUsaM2': 0,
        'proUsaPpauta': 0,
        'proControlado': 0,
        'proAplicacao': formatar_valor_com_aspas(row.get('proAplicacao')),
        'proPeso': formatar_valor(row.get('proPeso')),
        'protipo': str(row.get('protipo', 'P'))
    }

    if nome_valido(row.get('proSubGrupo')):
        columns_values['proSubGrupo_Temp'] = formatar_valor_com_aspas(row.get('proSubGrupo'))
        
    # Colunas opcionais: adiciona somente se o valor for válido
    for col in ['pclDescricao', 'gdpNome', 'fabNome']:
        if col in row and nome_valido(row[col]):
            columns_values[col] = formatar_valor_com_aspas(row[col])

    # Adiciona todas as colunas extras/dinâmicas somente se tiver valor válido
    for col in row.index:
        if col not in columns_values and col != 'proSubGrupo':
            valor = row[col]
            if valor is None or (isinstance(valor, str) and valor.strip().upper() in ('NULL', '', 'NAN')):
                continue  # ignora colunas vazias
            if isinstance(valor, str):
                valor = f"'{valor}'"
            columns_values[col] = valor

    columns = ', '.join(columns_values.keys())
    values = ', '.join(map(str, columns_values.values()))

    return f"INSERT INTO produto({columns}) VALUES({values}) GO"


def formatar_valor_monetario(valor):
    if valor is None or str(valor).strip().upper() in ('NULL', '', 'NAN'):
        return 'NULL'
    # Transforma em string
    s = str(valor).strip()
    if s.count('.') > 1:
        partes = s.split('.')
        s = ''.join(partes[:-1]) + '.' + partes[-1]
    return s

def gerar_insert_produto_empresa(row):

    proCodCst2_raw = row.get('proCodCst2', 0)
    if proCodCst2_raw is None or str(proCodCst2_raw).strip().upper() in ('NULL', '', 'NAN'):
        proCodCst2_num = 0
    else:
        try:
            proCodCst2_num = float(proCodCst2_raw)
        except:
            proCodCst2_num = 0
    Cst2 = str(int(proCodCst2_num)).zfill(2)

    columns_values = {
        'empid': to_int(row.get('empid')),
        'preid': to_int(row.get('preid')),
        'proid': to_int(row.get('proid')),
        'proCodigo': formatar_valor_com_aspas(row.get('proCodigo')),
        'ProCusto': formatar_valor_monetario(row.get('ProCusto')),
        'ProVenda': formatar_valor_monetario(row.get('ProVenda')),
        'ProUn': formatar_valor_com_aspas(str(row.get('ProUn', 'UN'))[:2]),
        'ProUnTrib': formatar_valor_com_aspas(str(row.get('ProUnTrib', 'UN'))[:2]),
        'proCodCst2': formatar_valor_com_aspas(Cst2),
        'proCodCSOSN': formatar_valor_com_aspas(row.get('proCodCSOSN')),
        'proLocalizador': formatar_valor_com_aspas(row.get('proLocalizador')),
        'proPrateleira': formatar_valor_com_aspas(str(row.get('proPrateleira') or '')[:20]),
        'proListaInvent': -1
    }

    # Adiciona todas as colunas extras (dinâmicas) somente se houver valor válido
    for col in row.index:
        if col not in columns_values:
            valor = row[col]
            # Verifica se o valor é válido
            if valor is None or (isinstance(valor, str) and valor.strip().upper() in ('NULL', '', 'NAN')):
                continue  
            if isinstance(valor, str):
                valor = f"'{valor}'"
            columns_values[col] = valor

    columns = ', '.join(columns_values.keys())
    values = ', '.join(map(str, columns_values.values()))



    return f"INSERT INTO produto_empresa({columns}) VALUES({values}) GO"

def gerar_insert_cod_barras(row):
    cdb_codigo_raw = row.get('cdbCodigo', '')
    # Converte corretamente removendo .0 de numeros
    cdb_codigo = converter_str_para_int_se_possivel(cdb_codigo_raw)
    if not cdb_codigo or cdb_codigo.upper() in ('NULL', 'NAN'):
        return None
    cdb_codigo = str(cdb_codigo)[:20]
    cdb_id_prod = to_int(row.get('proid'))
    cdb_codigo = formatar_valor_com_aspas(cdb_codigo)
    return f"INSERT INTO codBarras (cdbIdProd, cdbCodigo) VALUES ({cdb_id_prod}, {cdb_codigo}) GO"

def gerar_insert_empresafiltro_produto(row):
    empId = row.get('empid')
    proid = row.get('proid')

    if not empId or str(empId).upper() == 'NULL' or int(empId) == 1:
        return None  # Não gera insert se empid for 1 ou inválido

    data_hoje = datetime.now().strftime('%Y-%m-%d')
    data_formatada = f"'{data_hoje}'"

    return f"INSERT INTO empresafiltro (empId, emfTable, emfPkField, emfUsuId, emfPkValue, emfDataOcorrencia) " \
           f"VALUES ({empId}, 'produto', 'proid', '7', {proid}, {data_formatada}) GO"

# ================= Função principal =================

def gerar_sql_inserts(df, colunas_mapeadas=None):
    inserts = []

    # cria cópias para produto e produto_empresa
    df_produto = pd.DataFrame()
    df_empresa = pd.DataFrame()

    if colunas_mapeadas:
        if "produto" in colunas_mapeadas:
            for chave_padrao, coluna_excel in colunas_mapeadas["produto"].items():
                if coluna_excel in df.columns:
                    df_produto[chave_padrao] = df[coluna_excel]
                else:
                    df_produto[chave_padrao] = 'NULL'

        if "produto_empresa" in colunas_mapeadas:
            for chave_padrao, coluna_excel in colunas_mapeadas["produto_empresa"].items():
                if coluna_excel in df.columns:
                    df_empresa[chave_padrao] = df[coluna_excel]
                else:
                    df_empresa[chave_padrao] = 'NULL'

    # garante proid/preid
    if 'proid' not in df_produto.columns or df_produto['proid'].isnull().all():
        df_produto['proid'] = range(1, len(df) + 1)
        df_empresa['proid'] = df_produto['proid']

    if 'preid' not in df_empresa.columns:
        df_empresa['preid'] = range(1, len(df) + 1)

    # ==== Inserts auxiliares ====
    inserts.append('EXEC SP_MSFOREACHTABLE "ALTER TABLE ? NOCHECK CONSTRAINT ALL" GO')
    inserts.append('DBCC TRACEON(460, -1) GO')

    grupos_adicionados = set()
    fabricantes_adicionados = set()
    classes_adicionados = set()
    subgrupos_adicionados = set()

    # Flags para saber se houve inserts nas tabelas auxiliares
    teve_grupo = False
    teve_fabricante = False
    teve_classe = False
    teve_subgrupo = False

    # Adicionar colunas temporárias só se existirem no mapeamento
    if "gdpNome" in df_produto.columns:
        inserts.append("ALTER TABLE produto ADD gdpNome NVARCHAR(100) GO")
    if "fabNome" in df_produto.columns:
        inserts.append("ALTER TABLE produto ADD fabNome NVARCHAR(100) GO")
    if "pclDescricao" in df_produto.columns:
        inserts.append("ALTER TABLE produto ADD pclDescricao NVARCHAR(100) GO")
    if "proSubGrupo" in df_produto.columns:
        inserts.append("ALTER TABLE produto ADD proSubGrupo_Temp NVARCHAR(100) GO")   
        inserts.append("ALTER TABLE subGrupoProd ADD gdpNome_Temp NVARCHAR(100) GO")     

    for _, row in df_produto.iterrows():
        grupo = row.get('gdpNome')
        if "gdpNome" in df_produto.columns and nome_valido(grupo) and grupo not in grupos_adicionados:
            inserts.append(gerar_insert_grupo_prod(row))
            grupos_adicionados.add(grupo)
            teve_grupo = True

        fabricante = row.get('fabNome')
        if "fabNome" in df_produto.columns and nome_valido(fabricante) and fabricante not in fabricantes_adicionados:
            inserts.append(gerar_insert_fabricante(row))
            fabricantes_adicionados.add(fabricante)
            teve_fabricante = True

        classe = row.get('pclDescricao')
        if "pclDescricao" in df_produto.columns and nome_valido(classe) and classe not in classes_adicionados:
            inserts.append(gerar_insert_classe(row))
            classes_adicionados.add(classe)
            teve_classe = True

        subgrupo = row.get('proSubGrupo')
        if "proSubGrupo" in df_produto.columns and nome_valido(subgrupo):
            chave_sub = f"{row.get('gdpNome')}_{subgrupo}"
            if chave_sub not in subgrupos_adicionados:
                inserts.append(gerar_insert_subgrupo_prod(row))
                subgrupos_adicionados.add(chave_sub)
                teve_subgrupo = True
                
    # ==== Inserir produto ====
    inserts.append('SET IDENTITY_INSERT produto ON GO')
    for _, row in df_produto.iterrows():
        inserts.append(gerar_insert_produto(row))
    inserts.append('SET IDENTITY_INSERT produto OFF GO')

    # ==== Inserir produto_empresa ====
    inserts.append('SET IDENTITY_INSERT produto_empresa ON GO')
    for _, row in df_empresa.iterrows():
        inserts.append(gerar_insert_produto_empresa(row))
    inserts.append('SET IDENTITY_INSERT produto_empresa OFF GO')

    # ==== Inserir codBarras (se tiver no DF original) ====
    df["proid"] = df_produto["proid"]

    for _, row in df.iterrows():
        insert_cod = gerar_insert_cod_barras(row)
        if insert_cod:
            inserts.append(insert_cod)

    # ==== Inserir empresafiltro ====
    for _, row in df_empresa.iterrows():
        insert_empfiltro = gerar_insert_empresafiltro_produto(row)
        if insert_empfiltro:
            inserts.append(insert_empfiltro)

    # ==== Atualizações e limpeza (somente se houve insert) ====
    if "gdpNome" in df_produto.columns and teve_grupo:
        inserts.append('UPDATE produto SET proGrupo = grupoProd.gdpId FROM produto JOIN grupoProd ON produto.gdpNome = grupoProd.gdpNome GO')
        inserts.append('ALTER TABLE produto DROP COLUMN gdpNome GO')

    if "fabNome" in df_produto.columns and teve_fabricante:
        inserts.append('UPDATE produto SET proFab = fabricante.fabId FROM produto JOIN fabricante ON produto.fabNome = fabricante.fabNome GO')
        inserts.append('ALTER TABLE produto DROP COLUMN fabNome GO')

    if "pclDescricao" in df_produto.columns and teve_classe:
        inserts.append('UPDATE produto SET proClasseId = produtoClasse.pclId FROM produto JOIN produtoClasse ON produto.pclDescricao = produtoClasse.pclDescricao GO')
        inserts.append('ALTER TABLE produto DROP COLUMN pclDescricao GO')

    if "proSubGrupo" in df_produto.columns and teve_subgrupo:
        inserts.append('''
            UPDATE produto 
            SET proSubGrupo = subGrupoProd.sgpId 
            FROM produto 
            JOIN grupoProd ON produto.proGrupo = grupoProd.gdpId
            JOIN subGrupoProd 
                ON produto.proSubGrupo_Temp = subGrupoProd.sgpNome 
                AND subGrupoProd.sgpIdGdp = grupoProd.gdpId
            GO''')

        inserts.append('ALTER TABLE produto DROP COLUMN proSubGrupo_Temp GO')
    
    # UPDATE subGrupoProd somente se houve subgrupo
    if "proSubGrupo" in df_produto.columns and teve_subgrupo:
        inserts.append('''UPDATE subGrupoProd SET sgpIdGdp = grupoProd.gdpId FROM subGrupoProd JOIN grupoProd ON subGrupoProd.gdpNome_Temp = grupoProd.gdpNome GO''')
        inserts.append('ALTER TABLE subGrupoProd DROP COLUMN gdpNome_Temp GO')

    inserts.append('UPDATE produto_empresa SET proUnComercialId = produtoUN.unpId FROM produto_empresa JOIN produtoUN ON produto_empresa.proUn = produtoUN.unpUn GO')
    inserts.append('UPDATE produto_empresa SET proUnTribId = produtoUN.unpId FROM produto_empresa JOIN produtoUN ON produto_empresa.proUn = produtoUN.unpUn GO')
    inserts.append('UPDATE produto_empresa SET ProUnTrib = UPPER(ProUnTrib), ProUn = UPPER(ProUn) GO')
    inserts.append('UPDATE produto_empresa SET proClasseId = produto.proClasseId FROM produto_empresa JOIN produto ON produto_empresa.proid = produto.proid GO')
    inserts.append('EXEC SP_MSFOREACHTABLE "ALTER TABLE ? CHECK CONSTRAINT ALL" GO')
    inserts.append('update produto set produto.proNcmId = proNCM.ncmId from proncm inner join produto on produto.zzz_proCodigoNCM = proNCM.ncmCodigoCript GO')
    inserts.append('update produto set produto.proCestId = proNcmCest.pncCesId from proNcmCest inner join produto on proNcmCest.pncNcmId = produto.proNcmId go')
    inserts.append('DBCC TRACEOFF(460, -1) GO')

    return inserts




# ================= Funções auxiliares =================

def corrigir_null_no_arquivo(caminho_arquivo):
    with open(caminho_arquivo, 'r', encoding='utf-8') as f:
        conteudo = f.read()
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

