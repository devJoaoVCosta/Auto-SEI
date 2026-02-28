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
    padding: 6px;
    font-size: 13px;
"""


class DocumentoRow(QFrame):
    """Linha: # | Tipo (editavel) | Nome do arquivo (editavel) | status | remover"""

    def __init__(self, numero, nome_arquivo: str, on_remove):
        super().__init__()
        self.setStyleSheet("background-color: #ececec; border-radius: 4px;")
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        row = QHBoxLayout(self)
        row.setContentsMargins(6, 4, 6, 4)
        row.setSpacing(6)

        # Numero
        lbl_num = QLabel(str(numero))
        lbl_num.setFixedWidth(22)
        lbl_num.setStyleSheet("font-size: 12px; color: #888;")
        lbl_num.setAlignment(Qt.AlignCenter)
        row.addWidget(lbl_num)

        # Tipo de Documento (combo editavel com autocomplete)
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

        # Nome do arquivo — preenchido automaticamente, mas editavel se necessario
        self.entry_nome = QLineEdit()
        self.entry_nome.setText(nome_arquivo)
        self.entry_nome.setFixedHeight(30)
        self.entry_nome.setStyleSheet(STYLE_INPUT + "font-size: 12px;")
        self.entry_nome.setToolTip(nome_arquivo)
        row.addWidget(self.entry_nome, 1)

        # Status
        self.lbl_status = QLabel("")
        self.lbl_status.setFixedWidth(32)
        self.lbl_status.setAlignment(Qt.AlignCenter)
        self.lbl_status.setStyleSheet("font-size: 11px; font-weight: bold;")
        row.addWidget(self.lbl_status)

        # Remover
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

        self._linhas: list[DocumentoRow] = []
        self._contador = 0

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(8)

        # ── Titulo ───────────────────────────────────────────────
        titulo = QLabel("Automacao SEI - Funpresp-jud")
        titulo.setFixedHeight(40)
        titulo.setStyleSheet(
            "font-size: 20px; font-family: Arial; font-weight: bold;"
            "color: #333333; padding-bottom: 4px;"
        )
        layout.addWidget(titulo)

        # ── Credenciais ──────────────────────────────────────────
        cred_layout = QHBoxLayout()
        self.usuario_input = self._input("Usuario")
        self.senha_input = self._input("Senha", password=True)
        cred_layout.addWidget(self.usuario_input)
        cred_layout.addSpacing(8)
        cred_layout.addWidget(self.senha_input)
        layout.addLayout(cred_layout)

        # ── Processo + Pasta (lado a lado) ───────────────────────
        proc_pasta_layout = QHBoxLayout()
        proc_pasta_layout.setSpacing(8)

        self.processo_input = self._input("No processo  (Ex: 00001/000100/2024)")
        self.processo_input.setFixedWidth(230)
        proc_pasta_layout.addWidget(self.processo_input)

        self.pasta_input = QLineEdit()
        self.pasta_input.setPlaceholderText("Pasta com os arquivos PDF...")
        self.pasta_input.setFixedHeight(34)
        self.pasta_input.setStyleSheet(STYLE_INPUT)

        btn_pasta = QPushButton("Selecionar")
        btn_pasta.setFixedHeight(34)
        btn_pasta.setFixedWidth(90)
        btn_pasta.setStyleSheet(
            "background-color: #f7a833; color: black; border-radius: 5px; font-size: 13px;"
        )
        btn_pasta.clicked.connect(self._selecionar_pasta)

        proc_pasta_layout.addWidget(self.pasta_input, 1)
        proc_pasta_layout.addWidget(btn_pasta)
        layout.addLayout(proc_pasta_layout)

        # ── Cabecalho da tabela ──────────────────────────────────
        header = QFrame()
        header.setStyleSheet("background-color: #0e509a; border-radius: 4px;")
        header.setFixedHeight(26)
        h_row = QHBoxLayout(header)
        h_row.setContentsMargins(6, 0, 6, 0)
        h_row.setSpacing(6)

        for txt, fixed_w in [
            ("#",                 22),
            ("Tipo de Documento", 160),
            ("Nome do Arquivo",    0),
            ("",                  32),
            ("",                  28),
        ]:
            lbl = QLabel(txt)
            lbl.setStyleSheet("color: white; font-size: 11px; font-weight: bold;")
            if fixed_w:
                lbl.setFixedWidth(fixed_w)
            else:
                lbl.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
            h_row.addWidget(lbl)
        layout.addWidget(header)

        # ── Area rolavel ─────────────────────────────────────────
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFixedHeight(230)
        self.scroll_area.setStyleSheet(
            "QScrollArea { border: 1px solid #cccccc; border-radius: 4px; background: #f5f5f5; }"
        )

        self.linhas_widget = QWidget()
        self.linhas_widget.setStyleSheet("background: #f5f5f5;")
        self.linhas_layout = QVBoxLayout(self.linhas_widget)
        self.linhas_layout.setContentsMargins(4, 4, 4, 4)
        self.linhas_layout.setSpacing(4)
        self.linhas_layout.addStretch()

        self.scroll_area.setWidget(self.linhas_widget)
        layout.addWidget(self.scroll_area)

        # ── Botoes de acao ───────────────────────────────────────
        btns_layout = QHBoxLayout()

        btn_buscar = QPushButton("🔍  Buscar Arquivos")
        btn_buscar.setFixedHeight(32)
        btn_buscar.setStyleSheet(
            "background-color: #5a7a0a; color: white; border-radius: 5px;"
            "font-size: 13px; padding: 0 12px;"
        )
        btn_buscar.clicked.connect(self._buscar_arquivos)

        btn_add = QPushButton("  +  Adicionar Documento")
        btn_add.setFixedHeight(32)
        btn_add.setStyleSheet(
            "background-color: #0e509a; color: white; border-radius: 5px;"
            "font-size: 13px; padding: 0 12px;"
        )
        btn_add.clicked.connect(self._adicionar_linha_vazia)

        btn_reset = QPushButton("  Limpar Tudo")
        btn_reset.setFixedHeight(32)
        btn_reset.setStyleSheet(
            "background-color: #646464; color: white; border-radius: 5px;"
            "font-size: 13px; padding: 0 12px;"
        )
        btn_reset.clicked.connect(self.limpar_linhas)

        btns_layout.addWidget(btn_buscar)
        btns_layout.addSpacing(8)
        btns_layout.addWidget(btn_add)
        btns_layout.addSpacing(8)
        btns_layout.addWidget(btn_reset)
        btns_layout.addStretch()
        layout.addLayout(btns_layout)

        # ── Botao Executar ───────────────────────────────────────
        layout.addSpacing(4)
        btn_exec = QPushButton("  EXECUTAR AUTOMACAO")
        btn_exec.setFixedHeight(38)
        btn_exec.setStyleSheet(
            "background-color: #1a7a1a; color: white; border-radius: 6px;"
            "font-size: 14px; font-weight: bold;"
        )
        btn_exec.clicked.connect(self.executar_automacao)
        layout.addWidget(btn_exec)

        # Checkbox oculto — mantido no backend para reativacao futura.
        self.checkbox_salvar = QCheckBox("Lembrar usuario e senha")
        self.checkbox_salvar.setVisible(False)

        self._carregar_login_salvo()

    # ── Helpers ──────────────────────────────────────────────────

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

    # ── Pasta ─────────────────────────────────────────────────────

    def _selecionar_pasta(self):
        pasta = QFileDialog.getExistingDirectory(self, "Selecionar Pasta")
        if pasta:
            self.pasta_input.setText(pasta)

    def _buscar_arquivos(self):
        """Le a pasta e cria uma linha por PDF encontrado, preenchendo o nome automaticamente."""
        pasta = self.pasta_input.text().strip()

        if not pasta:
            QMessageBox.warning(self, "Aviso", "Selecione uma pasta antes de buscar arquivos.")
            return

        if not os.path.isdir(pasta):
            QMessageBox.warning(self, "Erro", f"Pasta nao encontrada:\n{pasta}")
            return

        pdfs = sorted([f for f in os.listdir(pasta) if f.lower().endswith(".pdf")])

        if not pdfs:
            QMessageBox.information(self, "Aviso", "Nenhum arquivo PDF encontrado na pasta selecionada.")
            return

        # Limpa as linhas existentes antes de preencher
        self.limpar_linhas()

        for nome in pdfs:
            self._contador += 1
            linha = DocumentoRow(self._contador, nome, self._remover_linha)
            self.linhas_layout.insertWidget(self.linhas_layout.count() - 1, linha)
            self._linhas.append(linha)

        # Rola para o topo para mostrar as linhas adicionadas
        self.scroll_area.verticalScrollBar().setValue(0)

    # ── Linhas ───────────────────────────────────────────────────

    def _adicionar_linha_vazia(self):
        """Adiciona uma linha em branco para o usuario preencher manualmente."""
        self._contador += 1
        linha = DocumentoRow(self._contador, "", self._remover_linha)
        self.linhas_layout.insertWidget(self.linhas_layout.count() - 1, linha)
        self._linhas.append(linha)
        self.scroll_area.verticalScrollBar().setValue(
            self.scroll_area.verticalScrollBar().maximum()
        )

    def _remover_linha(self, linha: DocumentoRow):
        self._linhas.remove(linha)
        linha.deleteLater()

    def limpar_linhas(self):
        for linha in self._linhas:
            linha.deleteLater()
        self._linhas.clear()
        self._contador = 0

    # ── Execucao ─────────────────────────────────────────────────

    def executar_automacao(self):
        usuario  = self.usuario_input.text().strip()
        senha    = self.senha_input.text().strip()
        processo = self.processo_input.text().strip()
        pasta    = self.pasta_input.text().strip()

        if not all([usuario, senha, processo, pasta]):
            QMessageBox.warning(
                self, "Erro",
                "Preencha todos os campos: usuario, senha, numero do processo e pasta."
            )
            return

        if not os.path.isdir(pasta):
            QMessageBox.warning(self, "Erro", f"Pasta nao encontrada:\n{pasta}")
            return

        if not self._linhas:
            QMessageBox.warning(self, "Erro", "Nenhum documento na lista. Clique em Buscar Arquivos.")
            return

        documentos = []
        for i, linha in enumerate(self._linhas, 1):
            tipo, nome = linha.dados()
            if not nome:
                QMessageBox.warning(self, "Erro", f"Linha {i}: nome do arquivo esta vazio.")
                return
            caminho = os.path.join(pasta, nome).replace("/", "\\")
            if not os.path.isfile(caminho):
                QMessageBox.warning(
                    self, "Erro",
                    f"Linha {i}: arquivo nao encontrado na pasta:\n{nome}"
                )
                return
            documentos.append({"tipo": tipo, "caminho": caminho})

        if self.checkbox_salvar.isChecked():
            set_key(".env", "SEI_USUARIO", usuario)
            set_key(".env", "SEI_SENHA", senha)

        try:
            automacao = SEIAutomation()
            resultados = automacao.executar(usuario, senha, processo, documentos)
            for linha, ok in zip(self._linhas, resultados):
                linha.set_status("OK" if ok else "ERR", cor="#1a7a1a" if ok else "#b22222")
            sucesso = sum(resultados)
            QMessageBox.information(
                self, "Concluido",
                f"Automacao finalizada!\nSucesso: {sucesso}   |   Erro: {len(resultados) - sucesso}"
            )
        except Exception as e:
            QMessageBox.critical(self, "Erro", str(e))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())