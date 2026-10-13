from PySide6.QtWidgets import QMessageBox, QFileDialog
from PySide6.QtGui import QIntValidator
import os
import pandas as pd

from Imports.importFinanceiro import (
    obter_diretorio_base,
    gerar_sql_inserts,
    salvar_sql_inserts,
    carregar_excel,
)


class FinanceiroController:
    def __init__(self, ui):
        self.ui = ui
        self.arquivo_excel = None
        self.colunas_disponiveis = []

        self.ui.vermelho_3.show()
        self.setup_ui()

    # ------------------------------------------------------------------
    # Setup
    # ------------------------------------------------------------------
    def setup_ui(self):
        # Preenche o combo fixo de tipo de conta com userData
        self.ui.pgtTipoConta.clear()
        self.ui.pgtTipoConta.addItem("A PAGAR",   "P")
        self.ui.pgtTipoConta.addItem("A RECEBER", "R")

        # empId_fi: valor fixo digitado pelo usuário (igual ao empId_pro de produtos)
        self.ui.empId_fi.setValidator(QIntValidator(1, 999))
        self.ui.empId_fi.setMaxLength(3)
        self.ui.empId_fi.setText("1")

        # Todos os combos que recebem colunas do Excel
        self.comboboxes_financeiro = [
            self.ui.pgtCliNome,
            self.ui.pgtValor,
            self.ui.pgtData,
            self.ui.pgtVecmto,
            self.ui.pgtDataQuitou,
            self.ui.pgtObs,
            self.ui.temp_pgtPlc,
            self.ui.temp_pgtSubPc,
            self.ui.pgtBancoRef,
            self.ui.pgtCecId,
            self.ui.pgtValorJuros,
            self.ui.pgtValorDescFin,
            self.ui.pgtValorMulta,
            self.ui.formapgt,
        ]

        for combo in self.comboboxes_financeiro:
            combo.currentIndexChanged.connect(self.atualizar_opcoes_comboboxes)

        self.ui.selecionarfi.clicked.connect(self.selecionar_arquivo)
        self.ui.gerar_fi.clicked.connect(self.gerar_script_financeiro)

    # ------------------------------------------------------------------
    # Reset
    # ------------------------------------------------------------------
    def resetar_estado(self):
        self.arquivo_excel = None
        self.colunas_disponiveis = []

        self.ui.vermelho_3.show()

        try:
            self.ui.plani_3.blockSignals(True)
            try:
                self.ui.plani_3.currentIndexChanged.disconnect(
                    self.atualizar_colunas_planilha
                )
            except TypeError:
                pass
            self.ui.plani_3.clear()
        finally:
            self.ui.plani_3.blockSignals(False)

        for combo in self.comboboxes_financeiro:
            combo.blockSignals(True)
            combo.clear()
            combo.addItem("")
            combo.blockSignals(False)

    # ------------------------------------------------------------------
    # Seleção de arquivo
    # ------------------------------------------------------------------
    def selecionar_arquivo(self):
        self.resetar_estado()
        arquivo, _ = QFileDialog.getOpenFileName(
            self.ui,
            "Selecionar Arquivo Excel",
            "",
            "Arquivos Excel (*.xlsx *.xls)",
        )
        if not arquivo:
            return

        self.arquivo_excel = arquivo
        self.ui.vermelho_3.hide()

        try:
            planilhas = pd.ExcelFile(self.arquivo_excel).sheet_names
            self.ui.plani_3.blockSignals(True)
            self.ui.plani_3.clear()
            self.ui.plani_3.addItems(planilhas)
            self.ui.plani_3.setCurrentIndex(0)
            self.ui.plani_3.blockSignals(False)

            self.ui.plani_3.currentIndexChanged.connect(
                self.atualizar_colunas_planilha
            )
            self.atualizar_colunas_planilha()

        except Exception as e:
            QMessageBox.warning(self.ui, "Erro", f"Erro ao carregar Excel:\n{e}")

    # ------------------------------------------------------------------
    # Atualiza colunas ao trocar planilha
    # ------------------------------------------------------------------
    def atualizar_colunas_planilha(self):
        if not self.arquivo_excel:
            return
        planilha = self.ui.plani_3.currentText()
        try:
            df = pd.read_excel(self.arquivo_excel, sheet_name=planilha)
            self.colunas_disponiveis = df.columns.tolist()
            self.atualizar_opcoes_comboboxes()
        except Exception as e:
            QMessageBox.warning(
                self.ui, "Erro", f"Erro ao carregar planilha {planilha}:\n{e}"
            )

    # ------------------------------------------------------------------
    # Impede seleção duplicada
    # ------------------------------------------------------------------
    def atualizar_opcoes_comboboxes(self):
        if not self.colunas_disponiveis:
            return

        usadas = {
            c.currentText()
            for c in self.comboboxes_financeiro
            if c.currentText()
        }

        for combo in self.comboboxes_financeiro:
            atual = combo.currentText()
            combo.blockSignals(True)
            combo.clear()
            combo.addItem("")
            for col in self.colunas_disponiveis:
                if col == atual or col not in usadas:
                    combo.addItem(col)
            combo.setCurrentText(atual)
            combo.blockSignals(False)

    # ------------------------------------------------------------------
    # Geração do script SQL
    # ------------------------------------------------------------------
    def gerar_script_financeiro(self):
        if not self.arquivo_excel:
            QMessageBox.warning(self.ui, "Erro", "Selecione um arquivo Excel.")
            return

        try:
            planilha   = self.ui.plani_3.currentText()
            df         = carregar_excel(self.arquivo_excel, sheet_name=planilha)
            tipo_conta = self.ui.pgtTipoConta.currentData()   # 'P' ou 'R'

            mapeamento = {
                "pgtCliNome":      self.ui.pgtCliNome.currentText(),
                "pgtValor":        self.ui.pgtValor.currentText(),
                "pgtData":         self.ui.pgtData.currentText(),
                "pgtVecmto":       self.ui.pgtVecmto.currentText(),
                "pgtDataQuitou":   self.ui.pgtDataQuitou.currentText(),
                "pgtObs":          self.ui.pgtObs.currentText(),
                "temp_pgtPlc":     self.ui.temp_pgtPlc.currentText(),
                "temp_pgtSubPc":   self.ui.temp_pgtSubPc.currentText(),
                "pgtBancoRef":     self.ui.pgtBancoRef.currentText(),
                "pgtCecId":        self.ui.pgtCecId.currentText(),
                "pgtValorJuros":   self.ui.pgtValorJuros.currentText(),
                "pgtValorDescFin": self.ui.pgtValorDescFin.currentText(),
                "pgtValorMulta":   self.ui.pgtValorMulta.currentText(),
                "formapgt":        self.ui.formapgt.currentText(),
            }

            empid_str = self.ui.empId_fi.text().strip()
            empid = int(empid_str) if empid_str.isdigit() else 1

            inserts = gerar_sql_inserts(df, mapeamento, tipo_conta, empid)

            diretorio    = obter_diretorio_base()
            nome_arquivo = (
                "insertsAPagar.txt" if tipo_conta == "P" else "insertsAReceber.txt"
            )
            arquivo_saida = os.path.join(diretorio, nome_arquivo)
            salvar_sql_inserts(arquivo_saida, inserts)

            QMessageBox.information(
                self.ui, "Sucesso", f"Script gerado:\n{arquivo_saida}"
            )

        except Exception as e:
            QMessageBox.critical(self.ui, "Erro", f"Erro ao gerar script:\n{e}")