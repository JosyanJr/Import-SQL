import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import datetime
import os
import sys
from ttkthemes import ThemedTk


def obter_diretorio_base():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


def carregar_arquivo():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if filepath:
        try:
            global excel_file
            excel_file = pd.ExcelFile(filepath)
            messagebox.showinfo("Sucesso", "Arquivo carregado com sucesso!")
            
            tabelas = excel_file.sheet_names
            tabela_combobox['values'] = tabelas
            if tabelas:
                tabela_combobox.set(tabelas[0])
                selecionar_tabela(None)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar arquivo: {e}")


def selecionar_tabela(event):
    tabela_selecionada = tabela_combobox.get()
    if tabela_selecionada:
        try:
            global df
            df = pd.read_excel(excel_file, sheet_name=tabela_selecionada, dtype=str)
            
            colunas = df.columns.tolist()
            campo_id_produto_combobox['values'] = colunas
            campo_qtde_estoque_combobox['values'] = colunas
            if colunas:
                campo_id_produto_combobox.set(colunas[0])
                campo_qtde_estoque_combobox.set(colunas[0])
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao selecionar tabela: {e}")


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


def atualizar_log(mensagem):
    log_text.configure(state=tk.NORMAL)
    log_text.insert(tk.END, mensagem + '\n')
    log_text.yview(tk.END)
    log_text.configure(state=tk.DISABLED)


def gerar_nf():
    fornecedor = fornecedor_entry.get()
    id_empresa = id_empresa_entry.get()
    tabela = tabela_combobox.get()
    campo_id_produto = campo_id_produto_combobox.get()
    campo_qtde_estoque = campo_qtde_estoque_combobox.get()

    if not all([fornecedor, id_empresa, tabela, campo_id_produto, campo_qtde_estoque]):
        messagebox.showwarning("Aviso", "Todos os campos devem ser preenchidos")
        return

    if 'df' not in globals():
        messagebox.showwarning("Aviso", "Nenhum arquivo carregado")
        return

    try:
        fornecedor = int(fornecedor)
        id_empresa = int(id_empresa)
    except ValueError:
        messagebox.showerror("Erro", "ID Fornecedor e ID Empresa devem ser números inteiros")
        return

    try:
        linhas_sql = []
        agora = datetime.now()

        # Cabeçalho do script
        linhas_sql.append("-- ============================================")
        linhas_sql.append(f"-- SQL gerado em: {agora.strftime('%d/%m/%Y %H:%M:%S')}")
        linhas_sql.append(f"-- Fornecedor: {fornecedor} | Empresa: {id_empresa}")
        linhas_sql.append(f"-- Tabela Excel: {tabela}")
        linhas_sql.append(f"-- Campo ID Produto: {campo_id_produto}")
        linhas_sql.append(f"-- Campo Qtde Estoque: {campo_qtde_estoque}")
        linhas_sql.append("-- ============================================")
        linhas_sql.append("")
        linhas_sql.append("USE [max4];")
        linhas_sql.append("GO")
        linhas_sql.append("")

        limite_itens = 160
        num_itens = len(df)
        total_nfs = 0
        contador_itens = 0

        for i in range(0, num_itens, limite_itens):
            total_nfs += 1
            nf_num = total_nfs

            linhas_sql.append(f"-- ===== NF {nf_num} - Lote {nf_num} =====")

            # INSERT na tabela NF
            data_str = agora.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
            linhas_sql.append(
                f"INSERT INTO NF (nfTipo, nfOperacao, nfCliente, nfBaseCalc, nfVlrIcms, "
                f"nfBaseCalcIcmsSub, nfVlrIcmsSub, nfVlrFrete, nfVlrOutraSDesp, nfVlrTotalNota, "
                f"nfStatus, nfDataEmissao, nfDataContabil, nfCliNome, nfAtendente, nfTipoNf, "
                f"nfDataLancamento, nfTipoNfSped, empid) "
                f"VALUES (4, '1102', {fornecedor}, 0, 0, 0, 0, 0, 0, 0, 'P', "
                f"'{data_str}', '{data_str}', '{fornecedor}', 7, 'E', '{data_str}', '01', {id_empresa});"
            )
            linhas_sql.append(f"DECLARE @nfId{nf_num} INT = SCOPE_IDENTITY();")
            linhas_sql.append("")

            itens_atual = df.iloc[i:i+limite_itens]
            for index, row in itens_atual.iterrows():
                id_produto = row[campo_id_produto]
                qtde_estoque = row[campo_qtde_estoque]

                if pd.isna(id_produto) or pd.isna(qtde_estoque):
                    continue

                id_produto_int = int(id_produto)
                qtde_estoque_int = limitar_qtde_estoque(qtde_estoque)

                contador_itens += 1
                mensagem_log = f"Item {contador_itens}: ID Produto: {id_produto_int}, Qtde Estoque: {qtde_estoque_int}"
                atualizar_log(mensagem_log)

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
                    f"'2000-01-01', 0, 0, 1, '1102');"
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
            linhas_sql.append("")
            linhas_sql.append("GO")
            linhas_sql.append("")

        # Rodapé
        linhas_sql.append("-- ============================================")
        linhas_sql.append(f"-- Total de itens processados: {contador_itens}")
        linhas_sql.append(f"-- Total de NFs geradas: {total_nfs}")
        linhas_sql.append("-- ============================================")

        # Gerar nome do arquivo
        timestamp = agora.strftime('%Y%m%d_%H%M%S')
        nome_arquivo = f"SQL_NF_Fornecedor{fornecedor}_Empresa{id_empresa}_{timestamp}.txt"
        caminho_arquivo = os.path.join(obter_diretorio_base(), nome_arquivo)

        # Salvar arquivo
        with open(caminho_arquivo, 'w', encoding='utf-8') as f:
            f.write('\n'.join(linhas_sql))

        ProgressBar1['value'] = 100
        app.update_idletasks()

        messagebox.showinfo("Sucesso", f"Arquivo SQL gerado com sucesso!\n\n{caminho_arquivo}")

        # Abrir a pasta contendo o arquivo
        try:
            os.startfile(obter_diretorio_base())
        except Exception:
            pass

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao gerar SQL: {e}")
        import traceback
        traceback.print_exc()


def limpar_dados():
    fornecedor_entry.delete(0, tk.END)
    id_empresa_entry.delete(0, tk.END)
    tabela_combobox.set('')
    campo_id_produto_combobox.set('')
    campo_qtde_estoque_combobox.set('')
    log_text.configure(state=tk.NORMAL)
    log_text.delete(1.0, tk.END)
    log_text.configure(state=tk.DISABLED)
    ProgressBar1['value'] = 0


app = ThemedTk(theme="arc")
app.title("NF Entrada - Gerar SQL")
app.geometry("450x500")

frame1 = ttk.Frame(app)
frame1.pack(pady=10)

fornecedor_label = ttk.Label(frame1, text="ID Fornecedor:")
fornecedor_label.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
fornecedor_entry = ttk.Entry(frame1)
fornecedor_entry.grid(row=0, column=1, padx=5, pady=5)

id_empresa_label = ttk.Label(frame1, text="ID Empresa:")
id_empresa_label.grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
id_empresa_entry = ttk.Entry(frame1)
id_empresa_entry.grid(row=1, column=1, padx=5, pady=5)

frame2 = ttk.Frame(app)
frame2.pack(pady=10)

tabela_label = ttk.Label(frame2, text="Tabela:")
tabela_label.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
tabela_combobox = ttk.Combobox(frame2, state='readonly')
tabela_combobox.grid(row=0, column=1, padx=5, pady=5)
tabela_combobox.bind("<<ComboboxSelected>>", selecionar_tabela)

campo_id_produto_label = ttk.Label(frame2, text="Campo ID Produto:")
campo_id_produto_label.grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
campo_id_produto_combobox = ttk.Combobox(frame2, state='readonly')
campo_id_produto_combobox.grid(row=1, column=1, padx=5, pady=5)

campo_qtde_estoque_label = ttk.Label(frame2, text="Campo Qtde Estoque:")
campo_qtde_estoque_label.grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
campo_qtde_estoque_combobox = ttk.Combobox(frame2, state='readonly')
campo_qtde_estoque_combobox.grid(row=2, column=1, padx=5, pady=5)

frame3 = ttk.Frame(app)
frame3.pack(pady=10)

carregar_button = ttk.Button(frame3, text="Carregar Arquivo", command=carregar_arquivo)
carregar_button.grid(row=0, column=0, padx=5, pady=5)

gerar_nf_button = ttk.Button(frame3, text="Gerar SQL", command=gerar_nf)
gerar_nf_button.grid(row=0, column=1, padx=5, pady=5)

limpar_dados_button = ttk.Button(frame3, text="Limpar Dados", command=limpar_dados)
limpar_dados_button.grid(row=0, column=2, padx=5, pady=5)

frame4 = ttk.Frame(app)
frame4.pack(pady=10)

log_text = tk.Text(frame4, height=10, width=50, state=tk.DISABLED)
log_text.pack()

ProgressBar1 = ttk.Progressbar(app, orient='horizontal', mode='determinate')
ProgressBar1.pack(fill=tk.X, padx=10, pady=10)

app.mainloop()

