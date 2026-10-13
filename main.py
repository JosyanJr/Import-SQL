import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QSizePolicy
from ui_main import Ui_MainWindow
from ClienteController import ClienteController
from ProdutoController import ProdutoController
from FinanceiroController import FinanceiroController


class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.setWindowTitle("Importador")

        self.setMinimumSize(0, 0)
        self.setMaximumSize(16777215, 16777215)

        self.setSizePolicy(
            QSizePolicy.Expanding,
            QSizePolicy.Expanding,
        )

        self.cliente_controller    = ClienteController(self)
        self.produto_controller    = ProdutoController(self)
        self.financeiro_controller = FinanceiroController(self)

        self.Pages.currentChanged.connect(self.ajustar_tamanho_janela)

        self.select.currentIndexChanged.connect(self.mudar_pagina)
        self.mudar_pagina(0)

    # ------------------------------------------------------------------

    def mudar_pagina(self, index):
        self.cliente_controller.resetar_estado()
        self.produto_controller.resetar_estado()
        self.financeiro_controller.resetar_estado()

        if index == 0:
            self.Pages.setCurrentWidget(self.pg_clifor)
        elif index == 1:
            self.Pages.setCurrentWidget(self.pg_prod)
        elif index == 2:
            self.Pages.setCurrentWidget(self.pg_fin)

        self.ajustar_tamanho_janela()

    # ------------------------------------------------------------------

    def ajustar_tamanho_janela(self):
        pagina = self.Pages.currentWidget()
        if not pagina:
            return

        largura_fixa = 800

        if pagina == self.pg_clifor:
            self.resize(largura_fixa, 408)
            return

        if pagina == self.pg_fin:
            self.resize(largura_fixa, 380)
            return

        pagina.adjustSize()
        self.adjustSize()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    janela = MainWindow()
    janela.show()
    sys.exit(app.exec())