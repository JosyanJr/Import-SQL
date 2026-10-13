import pandas as pd
import os
from datetime import datetime


def obter_diretorio_base():
    import sys
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


def limitar_qtde_estoque(valor):
    try:
        valor_float = float(valor)
        valor_int = int(abs(valor_float))
        if valor_int > 99999:
            valor_int = 99999
        if valor_float < 0:
            valor_int = -valor_int
        return valor_int
    except ValueError:
        return 0


def gerar_sql_nf_entrada(df, fornecedor, id_empresa, campo_id_produto, campo_qtde_estoque, nome_tabela=""):
    linhas_sql = []
    agora = datetime.now()

    limite_itens = 160
    num_itens = len(df)
    total_nfs = 0
    contador_itens = 0
    itens_ignorados = 0

    for i in range(0, num_itens, limite_itens):
        total_nfs += 1
        nf_num = total_nfs

        linhas_sql.append(f"-- ===== NF {nf_num} - Lote {nf_num} =====")

        # INSERT na tabela NF
        linhas_sql.append( 
            f"INSERT INTO NF (nfTipo, nfOperacao, nfCliente, nfBaseCalc, nfVlrIcms, "
            f"nfBaseCalcIcmsSub, nfVlrIcmsSub, nfVlrFrete, nfVlrOutraSDesp, nfVlrTotalNota, "
            f"nfStatus, nfDataEmissao, nfDataContabil, nfCliNome, nfAtendente, nfTipoNf, "
            f"nfDataLancamento, nfIndPgto, nfTipoNfSped, empid) "
            f"VALUES (4, '1102', {fornecedor}, 0, 0, 0, 0, 0, 0, 0, 'P', "
            f"GETDATE(), GETDATE(), '{fornecedor}', 7, 'E', GETDATE(), 9, '01', {id_empresa});"
        )
        linhas_sql.append(f"DECLARE @nfId{nf_num} INT = SCOPE_IDENTITY();")
        linhas_sql.append("")

        itens_atual = df.iloc[i:i + limite_itens]
        for index, row in itens_atual.iterrows():
            id_produto = row[campo_id_produto]
            qtde_estoque = row[campo_qtde_estoque]

            if pd.isna(id_produto) or pd.isna(qtde_estoque):
                continue

            id_produto_int = int(id_produto)
            qtde_estoque_int = limitar_qtde_estoque(qtde_estoque)

            # Pular itens com estoque zerado ou negativo
            if qtde_estoque_int <= 0:
                itens_ignorados += 1
                continue            

            contador_itens += 1

            # Verificar se o item já existe antes de inserir ou atualizar
            linhas_sql.append(
                f"IF NOT EXISTS (SELECT 1 FROM nfitem WHERE nfiNf = @nfId{nf_num} AND nfiProd = {id_produto_int})"
            )
            linhas_sql.append(
                f"    INSERT INTO nfitem (nfiNf, nfiProd, nfiSitTrib, nfiQtde, nfiQtdeCom, "
                f"nfiVlr, nfiIcms, nfiDesc, nfiBaseCalc, nfiVlrIcms, nfiVlrD, nfiLote, "
                f"nfiLoteVenc, nfiBaseSubst, nfiVlrSubst, nfiQtdFrac, nfiCfop) "
                f"VALUES (@nfId{nf_num}, {id_produto_int}, '020', {float(qtde_estoque_int)}, "
                f"{float(qtde_estoque_int)}, 0, 0, 0, 0, 0, 0, 'X', "
                f"CONVERT(DATETIME, '2026-01-01', 120), 0, 0, 1, '1102');"
            )
            linhas_sql.append("ELSE")
            linhas_sql.append(
                f"    UPDATE nfitem "
                f"SET nfiQtde = {float(qtde_estoque_int)}, nfiQtdeCom = {float(qtde_estoque_int)} "
                f"WHERE nfiNf = @nfId{nf_num} AND nfiProd = {id_produto_int};"
            )
            linhas_sql.append("")

        # Update para preencher valores dos itens desta NF
        linhas_sql.append(f"-- Atualizar valores dos itens da NF {nf_num}")
        linhas_sql.append(
            f"UPDATE nfitem "
            f"SET nfiVlr = produto_empresa.proCusto, "
            f"    nfiProdNome = produto.proDescricao, "
            f"    nfiProUn = produto.proUn "
            f"FROM nfitem "
            f"INNER JOIN produto ON produto.proid = nfitem.nfiProd "
            f"INNER JOIN produto_empresa ON produto.proid = produto_empresa.proId "
            f"WHERE nfitem.nfiNf = @nfId{nf_num} "
            f"  AND produto_empresa.empId = {id_empresa};"        
        )
        linhas_sql.append(
            f"UPDATE nf "
            f"SET nfCliNome = cliente.cliNome "
            f"FROM nf "
            f"INNER JOIN cliente ON nf.nfCliente = cliente.cliId "
            f"WHERE nf.nfId = @nfId{nf_num};"
        )
        linhas_sql.append("")
        linhas_sql.append("GO")
        linhas_sql.append("")

    return linhas_sql


def salvar_sql_script(linhas_sql, fornecedor, id_empresa, diretorio=None):
    """
    Salva as linhas SQL em um arquivo .txt
    
    Parâmetros:
        linhas_sql (list): Lista de strings SQL
        fornecedor (int): ID do fornecedor (usado no nome do arquivo)
        id_empresa (int): ID da empresa (usado no nome do arquivo)
        diretorio (str, optional): Diretório de saída. Se None, usa o diretório base.
    
    Retorna:
        str: Caminho completo do arquivo salvo
    """
    if diretorio is None:
        diretorio = obter_diretorio_base()
    
    agora = datetime.now()
    timestamp = agora.strftime('%Y%d%m_%H%M%S')
    nome_arquivo = f"SQL_NF_Fornecedor{fornecedor}_Empresa{id_empresa}_{timestamp}.txt"
    caminho_arquivo = os.path.join(diretorio, nome_arquivo)

    with open(caminho_arquivo, 'w', encoding='utf-8') as f:
        f.write('\n'.join(linhas_sql))

    return caminho_arquivo

