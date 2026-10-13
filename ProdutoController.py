from PySide6.QtWidgets import (
    QMessageBox, QFileDialog, QComboBox, QDialog, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QSizePolicy, QFrame, QGridLayout, QWidget
)
from PySide6.QtGui import QIntValidator
from PySide6.QtCore import Qt
import os
import sys
import pandas as pd

from Imports.importProdutos import (
    obter_diretorio_base, gerar_sql_inserts, salvar_sql_inserts,
    carregar_excel, nome_valido
)
from Imports.importEstoque import gerar_sql_nf_entrada



# ==========================================================
# Dialog para adicionar campo dinâmico
# ==========================================================
class AddCampoDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Adicionar Novo Campo")
        self.resize(400, 180)

        layout = QVBoxLayout(self)

        h1 = QHBoxLayout()
        h1.addWidget(QLabel("Nome do Label:"))
        self.label_name_input = QLineEdit()
        h1.addWidget(self.label_name_input)
        layout.addLayout(h1)

        h2 = QHBoxLayout()
        h2.addWidget(QLabel("Tabela:"))
        self.table_combo = QComboBox()
        self.table_combo.addItems(["produto", "produto_empresa"])
        h2.addWidget(self.table_combo)
        layout.addLayout(h2)

        h3 = QHBoxLayout()
        h3.addWidget(QLabel("Nome da Coluna:"))
        self.column_input = QLineEdit()
        h3.addWidget(self.column_input)
        layout.addLayout(h3)

        h4 = QHBoxLayout()
        h4.addWidget(QLabel("Tipo:"))
        self.type_combo = QComboBox()
        self.type_combo.addItems(["Texto", "Número"])
        h4.addWidget(self.type_combo)
        layout.addLayout(h4)

        self.save_btn = QPushButton("Gravar")
        self.save_btn.clicked.connect(self.save_field)
        layout.addWidget(self.save_btn)

        self.saved = False
        self.data = None

    def save_field(self):
        if not self.label_name_input.text().strip() or not self.column_input.text().strip():
            QMessageBox.warning(self, "Erro", "Preencha todos os campos.")
            return

        self.data = {
            "label_text": self.label_name_input.text().strip(),
            "tabela": self.table_combo.currentText(),
            "coluna": self.column_input.text().strip(),
            "tipo": self.type_combo.currentText()
        }
        self.saved = True
        self.accept()


# ==========================================================
# Controller principal
# ==========================================================
class ProdutoController:
    def __init__(self, ui):
        self.ui = ui
        self.arquivo_excel = None
        self.colunas_disponiveis = []

        # ---------------- Estado inicial ----------------
        self.ui.verde.hide()
        self.ui.vermelho.show()
        self.ui.preidInput.setEnabled(False)

        # =========================================================
        # LAYOUT MESTRE DO pg_prod (CORRECAO DEFINITIVA)
        # =========================================================
        layout_master = QVBoxLayout(self.ui.pg_prod)
        layout_master.setContentsMargins(9, 10, 9, 10)
        layout_master.setSpacing(6)

        # ---------- TOPO ----------
        layout_master.addWidget(self.ui.horizontalLayoutWidget)

        # ---------- CONTEUDO CENTRAL ----------
        conteudo = QWidget(self.ui.pg_prod)
        conteudo_layout = QVBoxLayout(conteudo)
        conteudo_layout.setContentsMargins(0, 0, 0, 0)
        conteudo_layout.setSpacing(6)

        # ----- linha superior (3 grids fixos) -----
        linha_grids = QWidget(conteudo)
        lg = QHBoxLayout(linha_grids)
        lg.setContentsMargins(0, 0, 0, 0)
        lg.setSpacing(6)

        lg.addWidget(self.ui.layoutWidget2)  # col 1
        lg.addWidget(self.ui.layoutWidget3)  # col 2
        lg.addWidget(self.ui.layoutWidget1)  # col 3

        conteudo_layout.addWidget(linha_grids)

        # ----- colunas dinamicas -----
        linha_colunas = QWidget(conteudo)
        lc = QHBoxLayout(linha_colunas)
        lc.setContentsMargins(0, 0, 0, 0)
        lc.setSpacing(6)

        lc.addWidget(self.ui.horizontalLayoutWidget_2)
        lc.addWidget(self.ui.horizontalLayoutWidget_3)
        lc.addWidget(self.ui.horizontalLayoutWidget_4)

        conteudo_layout.addWidget(linha_colunas)

        layout_master.addWidget(conteudo)

        # espaco elastico (empurra botoes para o fundo)
        layout_master.addStretch(1)

        # BOTOES NO FUNDO (gridLayout_4)
        layout_master.addWidget(self.ui.layoutWidget)

        # =========================================================
        # ---------------- Colunas dinamicas ----------------
        # =========================================================
        self.col1 = QGridLayout()
        self.col2 = QGridLayout()
        self.col3 = QGridLayout()

        for col in (self.col1, self.col2, self.col3):
            col.setColumnStretch(1, 1)
            col.setRowStretch(999, 1)
            col.setAlignment(Qt.AlignTop)
            col.setHorizontalSpacing(6)
            col.setVerticalSpacing(4)

        self.ui.horizontalLayout_6.addLayout(self.col1)
        self.ui.horizontalLayout_7.addLayout(self.col2)
        self.ui.horizontalLayout_8.addLayout(self.col3)

        # =========================================================
        # ---------------- Configuracoes iniciais ----------------
        # =========================================================
        self.ui.preidInput.setValidator(QIntValidator(1, 999999))

        self.ui.protipo.clear()
        self.ui.protipo.addItem("PRODUTO", "'P'")
        self.ui.protipo.addItem("SERVICO", "'S'")

        self.ui.empId_pro.setValidator(QIntValidator(1, 99))
        self.ui.empId_pro.setMaxLength(2)
        self.ui.empId_pro.setText("1")

        self.comboboxes_produtos = [
            self.ui.proid, self.ui.proCodigo, self.ui.proDescricao,
            self.ui.proClasse, self.ui.proSubGrupo, self.ui.gdpNome,
            self.ui.fabNome, self.ui.cdbCodigo,
            self.ui.zzz_proCodigoNCM, self.ui.ProCusto,
            self.ui.ProVenda, self.ui.ProUn,
            self.ui.proCodCst2, self.ui.proCodCSOSN,
            self.ui.proLocalizador, self.ui.proPrateleira,
            self.ui.proAplicacao, self.ui.proPeso,
            self.ui.ProUnTrib, self.ui.proestoque
        ]

        for combo in self.comboboxes_produtos:
            combo.currentIndexChanged.connect(
                self.atualizar_opcoes_comboboxes
            )

        self.ui.selecionarpro.clicked.connect(self.selecionar_arquivo)
        self.ui.gerarpro.clicked.connect(self.gerar_script_produtos)
        self.ui.addnovo_prod.clicked.connect(self.add_novo_campo)

        self.campos_dinamicos = []
        self.arquivo_campos = os.path.join(
            os.path.dirname(sys.argv[0]),
            "campos_produtos.xlsx"
        )

        self.carregar_campos_existentes()

    # -------------------------------------------------
    def adicionar_campo_layout(self, campo_info):
        label = QLabel(campo_info["label_text"].upper())
        label.setAlignment(Qt.AlignCenter)
        label.setFrameShape(QFrame.Box)
        label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        label.setMinimumWidth(110)

        combo = QComboBox()
        combo.addItem("")
        combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        combo.setMinimumWidth(110)

        index = len(self.campos_dinamicos)
        coluna = index % 3
        linha = index // 3

        layout = self.col1 if coluna == 0 else self.col2 if coluna == 1 else self.col3
        layout.addWidget(label, linha, 0)
        layout.addWidget(combo, linha, 1)

        self.campos_dinamicos.append({
            "label": label,
            "combo": combo,
            **campo_info
        })

    # -------------------------------------------------
    def carregar_campos_existentes(self):
        if os.path.exists(self.arquivo_campos):
            df = pd.read_excel(self.arquivo_campos)
            for _, row in df.iterrows():
                self.adicionar_campo_layout(row)

    # -------------------------------------------------
    def resetar_estado(self):
        self.arquivo_excel = None
        self.colunas_disponiveis = []

        self.ui.verde.hide()
        self.ui.vermelho.show()
        self.ui.preidInput.clear()
        self.ui.preidInput.setEnabled(False)

        for combo in self.comboboxes_produtos:
            combo.blockSignals(True)
            combo.clear()
            combo.addItem("")
            combo.blockSignals(False)

    # -------------------------------------------------
    def add_novo_campo(self):
        dialog = AddCampoDialog(self.ui)
        if dialog.exec() and dialog.saved:
            df = pd.DataFrame([dialog.data])
            if os.path.exists(self.arquivo_campos):
                df = pd.concat([pd.read_excel(self.arquivo_campos), df], ignore_index=True)
            df.to_excel(self.arquivo_campos, index=False)
            self.adicionar_campo_layout(dialog.data)

    # -------------------------------------------------
    def selecionar_arquivo(self):
        self.resetar_estado()
        arquivo, _ = QFileDialog.getOpenFileName(
            self.ui, "Selecionar Arquivo Excel", "", "Arquivos Excel (*.xlsx *.xls)"
        )
        if arquivo:
            self.arquivo_excel = arquivo
            self.ui.verde.show()
            self.ui.vermelho.hide()
            self.ui.preidInput.setEnabled(True)
            planilhas = pd.ExcelFile(arquivo).sheet_names
            self.ui.plani_2.clear()
            self.ui.plani_2.addItems(planilhas)
            self.ui.plani_2.currentIndexChanged.connect(self.atualizar_colunas_planilha)
            self.atualizar_colunas_planilha()

    # -------------------------------------------------
    def atualizar_colunas_planilha(self):
        if not self.arquivo_excel:
            return

        df = pd.read_excel(self.arquivo_excel, sheet_name=self.ui.plani_2.currentText())
        self.colunas_disponiveis = df.columns.tolist()

        for combo in self.comboboxes_produtos + [c["combo"] for c in self.campos_dinamicos]:
            combo.blockSignals(True)
            combo.clear()
            combo.addItem("")
            combo.addItems(self.colunas_disponiveis)
            combo.blockSignals(False)

    # -------------------------------------------------
    def atualizar_opcoes_comboboxes(self):
        usadas = {
            combo.currentText()
            for combo in self.comboboxes_produtos + [c["combo"] for c in self.campos_dinamicos]
            if combo.currentText()
        }

        for combo in self.comboboxes_produtos + [c["combo"] for c in self.campos_dinamicos]:
            atual = combo.currentText()
            combo.blockSignals(True)
            combo.clear()
            combo.addItem("")
            for col in self.colunas_disponiveis:
                if col == atual or col not in usadas:
                    combo.addItem(col)
            combo.setCurrentText(atual)
            combo.blockSignals(False)

    # -------------------------------------------------
    def gerar_script_produtos(self):
        if not self.arquivo_excel:
            QMessageBox.warning(self.ui, "Erro", "Selecione um arquivo Excel.")
            return
        try:
            # Carrega Excel e normaliza nomes das colunas
            df = carregar_excel(self.arquivo_excel, sheet_name=self.ui.plani_2.currentText())
            df.columns = df.columns.str.strip()
            df = df.fillna('NULL')

            # ---- Verificar se deve incluir estoque ----
            proforst_texto = self.ui.proforst.text().strip()
            tem_estoque = bool(proforst_texto)

            # ---- Verificar se deve incluir produtos ----
            col_desc = self.ui.proDescricao.currentText().strip()
            tem_produtos = bool(col_desc)

            # Se nao tem nem produtos nem estoque, avisa e sai
            if not tem_produtos and not tem_estoque:
                QMessageBox.warning(
                    self.ui, "Aviso",
                    "Selecione ao menos a coluna 'DESC. PROD.' para gerar produtos,"
                    "ou preencha 'FOR. EST.' e selecione 'ESTOQUE' para gerar estoque."
                )
                return

            protipo = self.ui.protipo.currentData() or "P"

            # --- Preparacao das colunas para os inserts de PRODUTOS ---
            sql_columns_produto = {}
            sql_columns_empresa = {}

            if tem_produtos:
                df['__protipo_fixo__'] = protipo
                sql_columns_produto["protipo"] = '__protipo_fixo__'

                # preid
                preid_inicial_str = self.ui.preidInput.text().strip()
                preid_inicial = int(preid_inicial_str) if preid_inicial_str.isdigit() else 1
                df['__preid_fixo__'] = range(preid_inicial, preid_inicial + len(df))
                sql_columns_empresa["preid"] = '__preid_fixo__'

                # empid
                empid_str = self.ui.empId_pro.text().strip()
                empid = int(empid_str) if empid_str.isdigit() else 1
                df['__empid_fixo__'] = empid
                sql_columns_empresa["empid"] = '__empid_fixo__'

                # ==== Colunas fixas para 'produto' ====
                if nome_valido(self.ui.proid.currentText()):
                    sql_columns_produto["proid"] = self.ui.proid.currentText()
                    sql_columns_empresa["proid"] = self.ui.proid.currentText()

                sql_columns_produto["proDescricao"] = col_desc
                sql_columns_produto["proDescPdv"] = col_desc

                if nome_valido(self.ui.proSubGrupo.currentText()):
                    sql_columns_produto["proSubGrupo"] = self.ui.proSubGrupo.currentText()
                if nome_valido(self.ui.zzz_proCodigoNCM.currentText()):
                    sql_columns_produto["zzz_proCodigoNCM"] = self.ui.zzz_proCodigoNCM.currentText()
                if nome_valido(self.ui.proAplicacao.currentText()):
                    sql_columns_produto["proAplicacao"] = self.ui.proAplicacao.currentText()
                if nome_valido(self.ui.proPeso.currentText()):
                    sql_columns_produto["proPeso"] = self.ui.proPeso.currentText()
                if nome_valido(self.ui.proClasse.currentText()):
                    sql_columns_produto["pclDescricao"] = self.ui.proClasse.currentText()
                if nome_valido(self.ui.gdpNome.currentText()):
                    sql_columns_produto["gdpNome"] = self.ui.gdpNome.currentText()
                if nome_valido(self.ui.fabNome.currentText()):
                    sql_columns_produto["fabNome"] = self.ui.fabNome.currentText()

                sql_columns_produto["proUsaM2"] = 0
                sql_columns_produto["proUsaPpauta"] = 0
                sql_columns_produto["proControlado"] = 0

                # ==== Colunas fixas para 'produto_empresa' ====
                if nome_valido(self.ui.proCodigo.currentText()):
                    sql_columns_empresa["proCodigo"] = self.ui.proCodigo.currentText()
                if nome_valido(self.ui.ProCusto.currentText()):
                    sql_columns_empresa["ProCusto"] = self.ui.ProCusto.currentText()
                if nome_valido(self.ui.ProVenda.currentText()):
                    sql_columns_empresa["ProVenda"] = self.ui.ProVenda.currentText()
                if nome_valido(self.ui.ProUn.currentText()):
                    sql_columns_empresa["ProUn"] = self.ui.ProUn.currentText()
                if nome_valido(self.ui.ProUnTrib.currentText()):
                    sql_columns_empresa["ProUnTrib"] = self.ui.ProUnTrib.currentText()
                if nome_valido(self.ui.proCodCst2.currentText()):
                    sql_columns_empresa["proCodCst2"] = self.ui.proCodCst2.currentText()
                if nome_valido(self.ui.proCodCSOSN.currentText()):
                    sql_columns_empresa["proCodCSOSN"] = self.ui.proCodCSOSN.currentText()
                if nome_valido(self.ui.proLocalizador.currentText()):
                    sql_columns_empresa["proLocalizador"] = self.ui.proLocalizador.currentText()
                if nome_valido(self.ui.proPrateleira.currentText()):
                    sql_columns_empresa["proPrateleira"] = self.ui.proPrateleira.currentText()

                # ==== CODIGO DE BARRAS ====
                col_excel_cdb = self.ui.cdbCodigo.currentText().strip()
                if col_excel_cdb and col_excel_cdb in df.columns:
                    df["cdbCodigo"] = df[col_excel_cdb]
                else:
                    df["cdbCodigo"] = 'NULL'

                # ==== Campos dinamicos ====
                for campo in self.campos_dinamicos:
                    tabela = campo['tabela']
                    coluna_sql = campo['coluna']
                    coluna_excel = campo['combo'].currentText().strip()

                    if coluna_excel and coluna_excel in df.columns:
                        if tabela == "produto":
                            sql_columns_produto[coluna_sql] = coluna_excel
                        else:
                            sql_columns_empresa[coluna_sql] = coluna_excel
                    else:
                        temp_null_col_name = f"__{coluna_sql}_NULL__"
                        df[temp_null_col_name] = 'NULL'
                        if tabela == "produto":
                            sql_columns_produto[coluna_sql] = temp_null_col_name
                        else:
                            sql_columns_empresa[coluna_sql] = temp_null_col_name

            # ==== Geracao do SQL ====
            diretorio = obter_diretorio_base()
            nome_arquivo = "insertsProdutos.txt" if protipo.strip("'").upper() == 'P' else "insertsServico.txt"
            arquivo_saida = os.path.join(diretorio, nome_arquivo)

            # Lista que vai acumular TODAS as linhas SQL
            linhas_totais = []

            # 1. Gera inserts de PRODUTOS (se houver)
            if tem_produtos:
                inserts = gerar_sql_inserts(
                    df,
                    {"produto": sql_columns_produto, "produto_empresa": sql_columns_empresa}
                )
                linhas_totais.extend(inserts)

            # 2. Gera inserts de ESTOQUE (se houver fornecedor preenchido)
            if tem_estoque:
                fornecedor = int(proforst_texto)
                empid_str = self.ui.empId_pro.text().strip()
                id_empresa = int(empid_str) if empid_str.isdigit() else 1
                campo_id_produto = self.ui.proid.currentText().strip()
                campo_qtde_estoque = self.ui.proestoque.currentText().strip()
                nome_tabela = self.ui.plani_2.currentText()

                if not campo_id_produto:
                    QMessageBox.warning(self.ui, "Erro", "Selecione a coluna 'ID PRODUTO' para o estoque.")
                    return
                if not campo_qtde_estoque:
                    QMessageBox.warning(self.ui, "Erro", "Selecione a coluna 'ESTOQUE' para a quantidade.")
                    return

                # Recarrega df sem fillna('NULL') para o estoque precisar de valores reais
                df_estoque = pd.read_excel(self.arquivo_excel, sheet_name=nome_tabela)

                for numero_linha, valor in enumerate(
                    df_estoque[campo_id_produto], start=2
                ):
                    if pd.isna(valor):
                        continue
                    try:
                        int(valor)
                    except (ValueError, TypeError):
                        QMessageBox.critical(
                            self.ui,
                            "Erro no arquivo de estoque",
                            f"Foi encontrado um valor inválido no arquivo.\n\n"
                            f"Linha: {numero_linha}\n"
                            f"Coluna: {campo_id_produto}\n"
                            f"Valor encontrado: '{valor}'\n\n"
                            f"Essa coluna deve conter apenas números.\n\n"
                            f"Corrija o arquivo de estoque e tente novamente."
                        )
                        return

                linhas_estoque = gerar_sql_nf_entrada(
                    df_estoque, fornecedor, id_empresa,
                    campo_id_produto, campo_qtde_estoque, nome_tabela
                )

                # Separador se tiver produtos antes
                if tem_produtos:
                    linhas_totais.append("")
                    linhas_totais.append("-- ============================================")
                    linhas_totais.append("-- INSERTS DE ESTOQUE (NF ENTRADA)")
                    linhas_totais.append("-- ============================================")
                    linhas_totais.append("")

                linhas_totais.extend(linhas_estoque)

            # Salva TUDO em um unico arquivo
            conteudo_final = '\n'.join(linhas_totais)
            conteudo_final = conteudo_final.replace("'NULL'", "NULL")
            with open(arquivo_saida, 'w', encoding='utf-8') as f:
                f.write(conteudo_final)

            QMessageBox.information(self.ui, "Sucesso", f"Script gerado:\n{arquivo_saida}")

        except Exception as e:
            QMessageBox.critical(self.ui, "Erro", f"Erro ao gerar script:\n{e}")

