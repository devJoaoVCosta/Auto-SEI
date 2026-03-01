import os
import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLineEdit, QPushButton, QMessageBox, QFileDialog, QCheckBox,
    QLabel, QComboBox, QScrollArea, QFrame, QSizePolicy, QCompleter,
)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt
from dotenv import load_dotenv, set_key
from selenium_handler import SEIAutomation


TIPOS_DOCUMENTO = [
    "Documentos",
    "Comprovante",
    "Demonstrativo",
    "Autorizacao",
    "Requerimento",
    "Oficio",
    "Relatorio",
    "Outros",
]

STYLE_INPUT = """
    background-color: #ffffff;
    color: black;
    border-radius: 5px;
    padding: 6px 6px 6px 28px;
    font-size: 13px;
"""


class DocumentoRow(QFrame):
    def __init__(self, numero, nome_arquivo: str, on_remove):
        super().__init__()
        self.setStyleSheet("""
            background-color: #e9eef4;
            border-radius: 5px;
        """)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        row = QHBoxLayout(self)
        row.setContentsMargins(6, 4, 6, 4)
        row.setSpacing(6)

        lbl_num = QLabel(str(numero))
        lbl_num.setFixedWidth(22)
        lbl_num.setStyleSheet("font-size: 12px; color: #666;")
        lbl_num.setAlignment(Qt.AlignCenter)
        row.addWidget(lbl_num)

        self.combo_tipo = QComboBox()
        self.combo_tipo.setEditable(True)
        self.combo_tipo.addItems(TIPOS_DOCUMENTO)
        self.combo_tipo.setFixedWidth(160)
        self.combo_tipo.setFixedHeight(30)
        self.combo_tipo.setStyleSheet(
            "QComboBox { background-color: #fff; color: black; font-size: 12px;"
            "  border-radius: 4px; padding: 2px 6px; }"
            "QComboBox QAbstractItemView { background: #fff; color: black; }"
        )
        completer_tipo = QCompleter(TIPOS_DOCUMENTO, self.combo_tipo)
        completer_tipo.setCaseSensitivity(Qt.CaseInsensitive)
        completer_tipo.setFilterMode(Qt.MatchContains)
        self.combo_tipo.setCompleter(completer_tipo)
        row.addWidget(self.combo_tipo)

        self.entry_nome = QLineEdit()
        self.entry_nome.setText(nome_arquivo)
        self.entry_nome.setFixedHeight(30)
        self.entry_nome.setStyleSheet("background-color: #fff; border-radius: 5px; padding: 6px; font-size: 12px;")
        row.addWidget(self.entry_nome, 1)

        self.lbl_status = QLabel("")
        self.lbl_status.setFixedWidth(32)
        self.lbl_status.setAlignment(Qt.AlignCenter)
        self.lbl_status.setStyleSheet("font-size: 11px; font-weight: bold;")
        row.addWidget(self.lbl_status)

        btn_rem = QPushButton("X")
        btn_rem.setFixedSize(28, 28)
        btn_rem.setStyleSheet(
            "background-color: #b22222; color: white; border-radius: 4px;"
            "font-size: 12px; font-weight: bold;"
        )
        btn_rem.clicked.connect(lambda: on_remove(self))
        row.addWidget(btn_rem)

    def dados(self):
        return self.combo_tipo.currentText().strip(), self.entry_nome.text().strip()

    def set_status(self, status: str, cor: str = "#333"):
        self.lbl_status.setText(status)
        self.lbl_status.setStyleSheet(f"font-size: 11px; font-weight: bold; color: {cor};")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Autobot")
        self.setGeometry(100, 100, 680, 560)
        self.setFixedSize(680, 560)
        self.setWindowIcon(QIcon("Icon.ico"))
        self.setStyleSheet("background-color: #dfdfdf;")

        load_dotenv()

        self._linhas = []
        self._contador = 0

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(8)

        titulo = QLabel("Automacao SEI - Funpresp-jud")
        titulo.setFixedHeight(40)
        titulo.setStyleSheet(
            "font-size: 20px; font-weight: bold; color: black;"
        )
        layout.addWidget(titulo)

        cred_layout = QHBoxLayout()

        self.usuario_input = self._input("Usuario")
        self.usuario_input.addAction(QIcon("user.png"), QLineEdit.LeadingPosition)

        self.senha_input = self._input("Senha", password=True)

        cred_layout.addWidget(self.usuario_input)
        cred_layout.addSpacing(8)
        cred_layout.addWidget(self.senha_input)
        layout.addLayout(cred_layout)

        proc_pasta_layout = QHBoxLayout()
        proc_pasta_layout.setSpacing(8)

        self.processo_input = self._input("Nº do processo (Ex: 00001/2026)")
        self.processo_input.setFixedWidth(230)
        proc_pasta_layout.addWidget(self.processo_input)

        self.pasta_input = QLineEdit()
        self.pasta_input.setPlaceholderText("Pasta com os arquivos PDF...")
        self.pasta_input.setFixedHeight(34)
        self.pasta_input.setStyleSheet(STYLE_INPUT)

        btn_pasta = QPushButton("Selecionar")
        btn_pasta.setFixedHeight(34)
        btn_pasta.setFixedWidth(90)
        btn_pasta.setStyleSheet("background-color: #f2b705; color: black; border-radius: 5px;")
        btn_pasta.clicked.connect(self._selecionar_pasta)

        proc_pasta_layout.addWidget(self.pasta_input, 1)
        proc_pasta_layout.addWidget(btn_pasta)
        layout.addLayout(proc_pasta_layout)

        header = QFrame()
        header.setStyleSheet("background-color: #0e509a; border-radius: 4px;")
        header.setFixedHeight(26)
        h_row = QHBoxLayout(header)
        h_row.setContentsMargins(6, 0, 6, 0)

        for txt, fixed_w in [
            ("  Tipo de Documento", 160),
            ("Nome do Arquivo", 0),
            ("", 32),
            ("", 28),
        ]:
            lbl = QLabel(txt)
            lbl.setStyleSheet("color: white; font-size: 11px; font-weight: bold;")
            if fixed_w:
                lbl.setFixedWidth(fixed_w)
            h_row.addWidget(lbl)
        layout.addWidget(header)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFixedHeight(230)
        self.scroll_area.setStyleSheet(
            "QScrollArea { border: 1px solid #cdd6e0; border-radius: 4px; background: #f2f4f7; }"
        )

        self.linhas_widget = QWidget()
        self.linhas_widget.setStyleSheet("background: #f2f4f7;")
        self.linhas_layout = QVBoxLayout(self.linhas_widget)
        self.linhas_layout.setContentsMargins(4, 4, 4, 4)
        self.linhas_layout.setSpacing(6)
        self.linhas_layout.addStretch()

        self.scroll_area.setWidget(self.linhas_widget)
        layout.addWidget(self.scroll_area)

        btns_layout = QHBoxLayout()

        btn_buscar = QPushButton("🔍 Buscar Arquivos")
        btn_buscar.setFixedHeight(32)
        btn_buscar.setStyleSheet("background-color: #0e509a; color: white; border-radius: 5px;")
        btn_buscar.clicked.connect(self._buscar_arquivos)

        btn_add = QPushButton("+ Adicionar Documento")
        btn_add.setFixedHeight(32)
        btn_add.setStyleSheet("background-color: #f2b705; color: black; border-radius: 5px;")
        btn_add.clicked.connect(self._adicionar_linha_vazia)

        btn_reset = QPushButton("Limpar Tudo")
        btn_reset.setFixedHeight(32)
        btn_reset.setStyleSheet("background-color: #9ea4aa; color: white; border-radius: 5px;")
        btn_reset.clicked.connect(self.limpar_linhas)

        btns_layout.addWidget(btn_buscar)
        btns_layout.addSpacing(8)
        btns_layout.addWidget(btn_add)
        btns_layout.addSpacing(8)
        btns_layout.addWidget(btn_reset)
        btns_layout.addStretch()
        layout.addLayout(btns_layout)
        btn_buscar.setFixedSize(160, 38)
        btn_add.setFixedSize(210, 38)
        btn_reset.setFixedSize(120, 38)

        btn_exec = QPushButton("EXECUTAR AUTOMACAO")
        btn_exec.setFixedHeight(38)
        btn_exec.setStyleSheet(
            "background-color: #0e509a; color: white; border-radius: 6px; font-size: 14px; font-weight: bold;"
        )
        btn_exec.clicked.connect(self.executar_automacao)
        layout.addWidget(btn_exec)

        self.checkbox_salvar = QCheckBox("Lembrar usuario e senha")
        self.checkbox_salvar.setVisible(False)

        self._carregar_login_salvo()

    def _input(self, placeholder, password=False):
        f = QLineEdit()
        f.setPlaceholderText(placeholder)
        f.setFixedHeight(34)
        f.setStyleSheet(STYLE_INPUT)
        if password:
            f.setEchoMode(QLineEdit.Password)
        return f

    def _carregar_login_salvo(self):
        usuario = os.getenv("SEI_USUARIO")
        senha = os.getenv("SEI_SENHA")
        if usuario and senha:
            self.usuario_input.setText(usuario)
            self.senha_input.setText(senha)
            self.checkbox_salvar.setChecked(True)

    def _selecionar_pasta(self):
        pasta = QFileDialog.getExistingDirectory(self, "Selecionar Pasta")
        if pasta:
            self.pasta_input.setText(pasta)

    def _buscar_arquivos(self):
        pasta = self.pasta_input.text().strip()
        if not pasta or not os.path.isdir(pasta):
            QMessageBox.warning(self, "Erro", "Selecione uma pasta valida.")
            return

        pdfs = sorted([f for f in os.listdir(pasta) if f.lower().endswith(".pdf")])
        if not pdfs:
            QMessageBox.information(self, "Aviso", "Nenhum PDF encontrado.")
            return

        self.limpar_linhas()
        for nome in pdfs:
            self._contador += 1
            linha = DocumentoRow(self._contador, nome, self._remover_linha)
            self.linhas_layout.insertWidget(self.linhas_layout.count() - 1, linha)
            self._linhas.append(linha)

    def _adicionar_linha_vazia(self):
        self._contador += 1
        linha = DocumentoRow(self._contador, "", self._remover_linha)
        self.linhas_layout.insertWidget(self.linhas_layout.count() - 1, linha)
        self._linhas.append(linha)

    def _remover_linha(self, linha):
        self._linhas.remove(linha)
        linha.deleteLater()

    def limpar_linhas(self):
        for linha in self._linhas:
            linha.deleteLater()
        self._linhas.clear()
        self._contador = 0

    def executar_automacao(self):
        pass


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())