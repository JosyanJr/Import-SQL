from PySide6.QtWidgets import QMessageBox, QFileDialog, QComboBox
from Imports.importClientes import (
    obter_diretorio_base,
    gerar_sql_inserts, salvar_sql_inserts, carregar_excel
)
import os
import pandas as pd


class ClienteController:
    def __init__(self, ui):
        self.ui = ui
        self.arquivo_excel = None
        self.colunas_disponiveis = []
        self.ui.verde_2.hide()
        self.ui.vermelho_2.show()        

        self.setup_ui()

    def setup_ui(self):
        self.ui.cliTipoCad.clear()
        self.ui.cliTipoCad.addItem("CLIENTE", 0)
        self.ui.cliTipoCad.addItem("FORNECEDOR", 1)
        self.ui.cliTipoCad.addItem("CLIENTE E FORNECEDOR", 2)
        self.ui.cliTipoCad.addItem("TRANSPORTADORA", 3)
        self.ui.cliTipoCad.addItem("USUARIOS", 5)

        self.comboboxes_clientes = [
            self.ui.cliId, self.ui.cliTipo, self.ui.cliNome, self.ui.cliFantasia, self.ui.cliCpfCgc,
            self.ui.cliRgInsc, self.ui.cliFatEnd, self.ui.cliFatBairro,
            self.ui.cliFatCidade, self.ui.cliFatUf, self.ui.cliFatCep,
            self.ui.cliFone, self.ui.cliCelular, self.ui.cliEmail,
            self.ui.cliFatEndNumero, self.ui.cliFatEndRef, self.ui.cliObsGeral,
            self.ui.empId, self.ui.cliUsuLoginId
        ]

        for combo in self.comboboxes_clientes:
            combo.currentIndexChanged.connect(self.atualizar_opcoes_comboboxes)

        self.ui.selecionarcli.clicked.connect(self.selecionar_arquivo)
        self.ui.gerarcli.clicked.connect(self.gerar_script_clientes)
        self.ui.cliTipoCad.currentIndexChanged.connect(self.ajustar_visibilidade_campos)
        self.ajustar_visibilidade_campos()

    def resetar_estado(self):
        # estado interno
        self.arquivo_excel = None
        self.colunas_disponiveis = []

        # ícones e campos
        self.ui.verde_2.hide()
        self.ui.vermelho_2.show()
        self.ui.cliTipoCad.setCurrentIndex(0)

        # planilhas
        try:
            self.ui.plani.blockSignals(True)
            try:
                self.ui.plani.currentIndexChanged.disconnect(self.atualizar_colunas_planilha)
            except TypeError:
                pass
            self.ui.plani.clear()
        finally:
            self.ui.plani.blockSignals(False)

        # comboboxes de mapeamento
        for combo in self.comboboxes_clientes:
            combo.blockSignals(True)
            combo.clear()
            combo.addItem("")
            combo.blockSignals(False)

    def selecionar_arquivo(self):
        self.resetar_estado()        
        arquivo, _ = QFileDialog.getOpenFileName(
            self.ui, "Selecionar Arquivo Excel", "", "Arquivos Excel (*.xlsx *.xls)"
        )
        if arquivo:
            self.arquivo_excel = arquivo
            self.ui.verde_2.show()
            self.ui.vermelho_2.hide()            
            try:
                # lista planilhas
                planilhas = pd.ExcelFile(self.arquivo_excel).sheet_names
                self.ui.plani.clear()
                self.ui.plani.addItems(planilhas)
                self.ui.plani.setCurrentIndex(0)

                # conecta troca de planilha
                self.ui.plani.currentIndexChanged.connect(self.atualizar_colunas_planilha)

                # carrega a primeira planilha
                self.atualizar_colunas_planilha()

            except Exception as e:
                QMessageBox.warning(self.ui, "Erro", f"Erro ao carregar Excel:\n{e}")

    def atualizar_colunas_planilha(self):
        """Carrega as colunas da planilha selecionada e atualiza os comboboxes"""
        if not self.arquivo_excel:
            return

        planilha_selecionada = self.ui.plani.currentText()
        try:
            df = pd.read_excel(self.arquivo_excel, sheet_name=planilha_selecionada)
            self.colunas_disponiveis = df.columns.tolist()
            self.atualizar_opcoes_comboboxes()
        except Exception as e:
            QMessageBox.warning(
                self.ui, "Erro", f"Erro ao carregar planilha {planilha_selecionada}:\n{e}"
            )

    def ajustar_visibilidade_campos(self):
        tipo = self.ui.cliTipoCad.currentData()

        for widget in self.ui.findChildren(QComboBox):
            if widget != self.ui.cliTipoCad:
                widget.setCurrentIndex(0)

        if tipo == 5:
            self.ui.empId.hide()
            self.ui.label_21.hide()
            self.ui.cliUsuLoginId.show()
            self.ui.label_23.show()
        else:
            self.ui.empId.show()
            self.ui.label_21.show()
            self.ui.cliUsuLoginId.hide()
            self.ui.label_23.hide()

        self.atualizar_opcoes_comboboxes()

    def atualizar_opcoes_comboboxes(self):
        if not self.colunas_disponiveis:
            return

        usadas = {combo.currentText() for combo in self.comboboxes_clientes if combo.currentText()}
        for combo in self.comboboxes_clientes:
            texto_atual = combo.currentText()
            combo.blockSignals(True)
            combo.clear()
            combo.addItem("")
            for coluna in self.colunas_disponiveis:
                if coluna == texto_atual or coluna not in usadas:
                    combo.addItem(coluna)
            combo.setCurrentText(texto_atual)
            combo.blockSignals(False)

    def gerar_script_clientes(self):
        if not self.arquivo_excel:
            QMessageBox.warning(self.ui, "Erro", "Selecione um arquivo Excel.")
            return
        try:
            planilha_selecionada = self.ui.plani.currentText()
            df = carregar_excel(self.arquivo_excel, sheet_name=planilha_selecionada).fillna('NULL')

            colunas_mapeadas = {
                "cliId": self.ui.cliId.currentText(),
                "cliTipo": self.ui.cliTipo.currentText(),
                "cliNome": self.ui.cliNome.currentText(),
                "cliFantasia": self.ui.cliFantasia.currentText(),
                "cliCpfCgc": self.ui.cliCpfCgc.currentText(),
                "cliRgInsc": self.ui.cliRgInsc.currentText(),
                "cliFatEnd": self.ui.cliFatEnd.currentText(),
                "cliFatBairro": self.ui.cliFatBairro.currentText(),
                "cliFatCidade": self.ui.cliFatCidade.currentText(),
                "cliFatUf": self.ui.cliFatUf.currentText(),
                "cliFatCep": self.ui.cliFatCep.currentText(),
                "cliFone": self.ui.cliFone.currentText(),
                "cliCelular": self.ui.cliCelular.currentText(),
                "cliEmail": self.ui.cliEmail.currentText(),
                "cliFatEndNumero": self.ui.cliFatEndNumero.currentText(),
                "cliFatEndRef": self.ui.cliFatEndRef.currentText(),
                "cliObsGeral": self.ui.cliObsGeral.currentText(),
                "cliTipoCad": "cliTipoCad"
            }

            if self.ui.cliTipoCad.currentData() == 5:
                colunas_mapeadas["cliUsuLoginId"] = self.ui.cliUsuLoginId.currentText()
            else:
                colunas_mapeadas["empId"] = self.ui.empId.currentText()

            df['cliTipoCad'] = self.ui.cliTipoCad.currentData()

            inserts = gerar_sql_inserts(df, colunas_mapeadas)

            tipo = self.ui.cliTipoCad.currentData()
            diretorio = obter_diretorio_base()

            if tipo == 0:
                nome_arquivo = "insertsClientes.txt"
            elif tipo == 1:
                nome_arquivo = "insertsFornecedores.txt"
            elif tipo == 2:
                nome_arquivo = "insertsClientesFornecedores.txt"
            elif tipo == 3:
                nome_arquivo = "insertsTransportadoras.txt"
            elif tipo == 5:
                nome_arquivo = "insertsUsuarios.txt"
            else:
                nome_arquivo = "insertsClientes.txt"  # padrão

            arquivo_saida = os.path.join(diretorio, nome_arquivo)
            salvar_sql_inserts(arquivo_saida, inserts)

            QMessageBox.information(self.ui, "Sucesso", f"Script gerado:\n{arquivo_saida}")

        except Exception as e:
            QMessageBox.critical(self.ui, "Erro", f"Erro ao gerar script:\n{e}")
