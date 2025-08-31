"""
App Profissional de OPs de Calçados (PyQt5 + SQLite)
----------------------------------------------------
Recursos principais:
- Banco SQLite com criação automática e índices.
- Tela inicial com busca e lista de OPs (abrir por duplo-clique, excluir, exportar CSV).
- Tela de criação de OP com validações.
- Tela de visualização/edição da OP com abas por GIRO (1–5).
  * Edição direta das quantidades por tamanho com QSpinBox.
  * Linha TOTAL LOTE automática por GIRO e linha TOTAL GERAL no rodapé.
  * Validação: cada TALÃO deve somar 12 pares (personalizável por constante).
  * Botões: Salvar alterações, Exportar CSV da OP.
- Aparência moderna (tema claro com acento roxo + cantos arredondados) e responsiva (auto stretch de colunas).

Observações:
- Dependências: PyQt5 (pip install PyQt5). Somente stdlib + PyQt5.
- Rode assim:  python app_op_calcados.py
- O arquivo de banco padrão é: producao_calcados.db (na mesma pasta).

Autor: você 💜
"""

import sys
import os
import csv
import sqlite3
from datetime import datetime, time
from typing import Dict, List, Tuple

from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex, QVariant
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QStackedWidget, QMessageBox, QTableWidget, QTableWidgetItem,
    QGroupBox, QFormLayout, QSizePolicy, QSpacerItem, QTabWidget, QTableView,
    QHeaderView, QAbstractItemView, QToolButton, QStyle, QFileDialog, QComboBox,
    QStyledItemDelegate, QSpinBox
)
import pyautogui

# ==========================
# Configurações da Aplicação
# ==========================
DATABASE_PATH = "producao_calcados.db"
TAMANHOS = ["4", "4x", "5x", "6", "7", "7x", "8x", "9x", "10", "11", "12"]
GIROS = [1, 2, 3, 4, 5]
PARES_POR_TALAO = 20  # Agora cada talão terá 20 pares

# ==============
# Banco de Dados
# ==============

def criar_banco():
    novo_banco = not os.path.exists(DATABASE_PATH)
    conn = sqlite3.connect(DATABASE_PATH)
    c = conn.cursor()

    # Tabelas
    c.execute(
        """
        CREATE TABLE IF NOT EXISTS ops (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cliente TEXT NOT NULL,
            num_op INTEGER UNIQUE NOT NULL,
            data_criacao TEXT NOT NULL,
            total_pares INTEGER NOT NULL
        )
        """
    )
    c.execute(
        """
        CREATE TABLE IF NOT EXISTS taloes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            op_id INTEGER NOT NULL,
            giro INTEGER NOT NULL,
            talao_num INTEGER NOT NULL,
            numeracao TEXT NOT NULL,
            quantidade INTEGER NOT NULL DEFAULT 0,
            status TEXT NOT NULL DEFAULT 'pendente',
            FOREIGN KEY (op_id) REFERENCES ops (id)
        )
        """
    )

    # Índices para performance em listas/pesquisas
    c.execute("CREATE INDEX IF NOT EXISTS idx_ops_numop ON ops(num_op)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_taloes_op_giro ON taloes(op_id, giro)")

    conn.commit()
    conn.close()
    return novo_banco

# ==============
# Estilos (QSS)
# ==============

APP_QSS = """
* { font-family: 'Segoe UI'; }
QMainWindow { background-color: #000000; }
QLabel[cls="title"] { font-size: 24px; font-weight: 800; color: #cdd9e5; }
QLabel[cls="subtitle"] { font-size: 14px; color: #768390; }
QGroupBox { border: 1px solid #222; border-radius: 12px; margin-top: 16px; background: #111; }
QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 4px 8px; color: #cdd9e5; }
QLineEdit, QComboBox {
    background: #000; border: 1px solid #222; border-radius: 8px; padding: 10px; font-size: 14px; color: #cdd9e5;
}
QLineEdit:focus, QComboBox:focus { border: 1px solid #7c4dff; }
QPushButton {
    background-color: #7c4dff; color: #fff; padding: 8px 18px; border-radius: 10px; font-weight: 700; border: 0;
    min-width: 120px; max-width: 220px;
}
QPushButton[destructive="true"] { background: #ff4d6d; }
QPushButton:hover { background-color: #6a3dff; }
QToolButton { background: #000; border: 1px solid #222; border-radius: 8px; padding: 8px; color: #cdd9e5; }
QToolButton:hover { border-color: #7c4dff; background: #222; }
QTabBar::tab {
    padding: 10px 16px;
    border: 1px solid #222;
    border-bottom: none;
    border-top-left-radius: 10px;
    border-top-right-radius: 10px;
    background: #111;
    color: #cdd9e5;
}
QTabBar::tab:selected { background: #000; color: #fff; }
QTabWidget::pane { border: 1px solid #222; border-radius: 10px; top: -1px; }
QTableWidget, QTableView, QWidget, QMainWindow, QTabWidget, QGroupBox {
    background: #000;
    color: #cdd9e5;
}
QTableWidget::item {
    background: #000;
    color: #cdd9e5;
}
QHeaderView::section {
    background: #111;
    border: none;
    padding: 8px;
    font-weight: 700;
    color: #cdd9e5;
}
QTableWidget QTableCornerButton::section {
    background: #111;
    border: 1px solid #222;
}
QTableWidget::item {
    border: 1px solid #222;
    padding: 2px;
    color: #cdd9e5;
    background: #000;
}
QTableWidget::item:selected {
    background: #7c4dff;
    color: #fff;
}
"""

# =====================
# Utilidades de Banco
# =====================

def inserir_op(cliente: str, num_op: int, total_pares: int, tipo: str) -> int:
    conn = sqlite3.connect(DATABASE_PATH)
    c = conn.cursor()
    data_criacao = datetime.now().strftime("%Y-%m-%d %H:%M")
    c.execute(
        "INSERT INTO ops (cliente, num_op, data_criacao, total_pares, tipo) VALUES (?, ?, ?, ?, ?)",
        (cliente, num_op, data_criacao, total_pares, tipo),
    )
    op_id = c.lastrowid
    conn.commit()
    conn.close()
    return op_id


def gerar_taloes_iniciais(op_id: int, total_pares: int, tipo: str):
    if tipo == "Masculino":
        tamanhos = ["7", "7x", "8x", "9x", "10", "11", "12"]
        taloes_por_tamanho = [1, 1, 2, 3, 2, 2, 1]
    else:
        tamanhos = ["4", "4x", "5", "6", "7", "7x", "8x"]
        taloes_por_tamanho = [1, 1, 3, 3, 2, 1, 1]

    conn = sqlite3.connect(DATABASE_PATH)
    c = conn.cursor()
    pares_restantes = total_pares

    for giro in GIROS:
        talao_num = 1
        for idx, tam in enumerate(tamanhos):
            for _ in range(taloes_por_tamanho[idx]):
                if pares_restantes <= 0:
                    break
                qtd = min(PARES_POR_TALAO, pares_restantes)
                c.execute(
                    "INSERT INTO taloes (op_id, giro, talao_num, numeracao, quantidade) VALUES (?, ?, ?, ?, ?)",
                    (op_id, giro, talao_num, tam, qtd),
                )
                pares_restantes -= qtd
                talao_num += 1
        if pares_restantes <= 0:
            break
    conn.commit()
    conn.close()


def listar_ops(filtro: str = "") -> List[Tuple]:
    conn = sqlite3.connect(DATABASE_PATH)
    c = conn.cursor()
    if filtro:
        like = f"%{filtro}%"
        c.execute(
            "SELECT id, cliente, num_op, data_criacao, total_pares FROM ops WHERE cliente LIKE ? OR CAST(num_op AS TEXT) LIKE ? ORDER BY data_criacao DESC",
            (like, like),
        )
    else:
        c.execute(
            "SELECT id, cliente, num_op, data_criacao, total_pares FROM ops ORDER BY data_criacao DESC"
        )
    rows = c.fetchall()
    conn.close()
    return rows


def carregar_taloes(op_id: int) -> Dict[int, Dict[int, Dict[str, int]]]:
    """Retorna estrutura: {giro: {talao_num: {tamanho: qtd}}} ordenada por giro, talão."""
    conn = sqlite3.connect(DATABASE_PATH)
    c = conn.cursor()
    c.execute(
        "SELECT giro, talao_num, numeracao, quantidade FROM taloes WHERE op_id = ? ORDER BY giro, talao_num",
        (op_id,),
    )
    dados = c.fetchall()
    conn.close()

    giros: Dict[int, Dict[int, Dict[str, int]]] = {g: {} for g in GIROS}
    for giro, talao_num, numeracao, quantidade in dados:
        if talao_num not in giros[giro]:
            giros[giro][talao_num] = {num: 0 for num in TAMANHOS}
        giros[giro][talao_num][numeracao] = quantidade
    return giros


def salvar_taloes(op_id: int, giros: Dict[int, Dict[int, Dict[str, int]]]) -> None:
    conn = sqlite3.connect(DATABASE_PATH)
    c = conn.cursor()
    for giro, taloes in giros.items():
        for talao_num, tamanhos in taloes.items():
            for numeracao, quantidade in tamanhos.items():
                c.execute(
                    """
                    UPDATE taloes SET quantidade = ?
                    WHERE op_id = ? AND giro = ? AND talao_num = ? AND numeracao = ?
                    """,
                    (quantidade, op_id, giro, talao_num, numeracao),
                )
    conn.commit()
    conn.close()

# ========================
# Delegates (Editor célula)
# ========================

class SpinBoxDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        editor = QSpinBox(parent)
        editor.setRange(0, 99)
        return editor

    def setEditorData(self, editor, index):
        try:
            value = int(index.model().data(index, Qt.EditRole) or 0)
        except ValueError:
            value = 0
        editor.setValue(value)

    def setModelData(self, editor, model, index):
        model.setData(index, editor.value(), Qt.EditRole)

# ===================
# Telas da Aplicação
# ===================

class ListaOPsPage(QWidget):
    def __init__(self, abrir_op_callback, criar_op_callback):
        super().__init__()
        self.abrir_op_callback = abrir_op_callback
        self.criar_op_callback = criar_op_callback
        self._setup_ui()
        self.atualizar()

    def _setup_ui(self):
        self.setStyleSheet(APP_QSS)
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)

        header = QHBoxLayout()
        title = QLabel("Ordens de Produção")
        title.setProperty("cls", "title")
        header.addWidget(title)
        header.addStretch()

        self.busca = QLineEdit()
        self.busca.setPlaceholderText("Buscar por cliente ou Nº OP…")
        self.busca.textChanged.connect(self.atualizar)
        header.addWidget(self.busca)

        bt_novo = QPushButton("+ Nova OP")
        bt_novo.setMinimumWidth(120)
        bt_novo.clicked.connect(self.criar_op_callback)
        header.addWidget(bt_novo)
        root.addLayout(header)

        # Painel de botões de ação acima da tabela
        self.botoes_acoes = QHBoxLayout()
        self.bt_abrir = QPushButton("Abrir")
        self.bt_export = QPushButton("Exportar CSV")
        self.bt_excluir = QPushButton("Excluir")
        self.bt_excluir.setProperty("destructive", True)
        for bt in [self.bt_abrir, self.bt_export, self.bt_excluir]:
            bt.setMinimumWidth(120)
            bt.setMaximumWidth(220)
            bt.setMinimumHeight(32)
            bt.setMaximumHeight(36)
            self.botoes_acoes.addWidget(bt)
        self.botoes_acoes.addStretch()
        root.addLayout(self.botoes_acoes)

        self.tabela = QTableWidget()
        self.tabela.setColumnCount(6)
        self.tabela.setHorizontalHeaderLabels(["ID", "Cliente", "Nº OP", "Criada em", "Total Pares", "Ações"])
        self.tabela.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tabela.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.tabela.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tabela.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tabela.cellDoubleClicked.connect(self._duplo_clique)
        self.tabela.verticalHeader().setVisible(False)
        root.addWidget(self.tabela)

    def atualizar(self):
        filtro = self.busca.text().strip()
        dados = listar_ops(filtro)
        self.tabela.setRowCount(len(dados))
        for r, (op_id, cliente, num_op, data_criacao, total_pares) in enumerate(dados):
            self.tabela.setItem(r, 0, QTableWidgetItem(str(op_id)))
            self.tabela.setItem(r, 1, QTableWidgetItem(cliente))
            self.tabela.setItem(r, 2, QTableWidgetItem(str(num_op)))
            self.tabela.setItem(r, 3, QTableWidgetItem(data_criacao))
            self.tabela.setItem(r, 4, QTableWidgetItem(str(total_pares)))

            acao_widget = QWidget()
            lay = QHBoxLayout(acao_widget)
            lay.setContentsMargins(0, 0, 0, 0)
            lay.setSpacing(16)  # Espaço maior entre botões

            bt_abrir = QPushButton("Abrir")
            bt_abrir.setMinimumWidth(90)
            bt_abrir.setMaximumWidth(90)
            bt_abrir.setMinimumHeight(36)
            bt_abrir.setMaximumHeight(36)
            bt_abrir.setStyleSheet("""
                QPushButton {
                    background-color: #7c4dff;
                    color: #fff;
                    border-radius: 0px;
                    font-weight: bold;
                    font-size: 14px;
                    padding: 0px;
                    text-align: center;
                }
                QPushButton:hover { background-color: #6a3dff; }
            """)

            bt_export = QPushButton("Exportar CSV")
            bt_export.setMinimumWidth(90)
            bt_export.setMaximumWidth(90)
            bt_export.setMinimumHeight(36)
            bt_export.setMaximumHeight(36)
            bt_export.setStyleSheet("""
                QPushButton {
                    background-color: #43d96b;
                    color: #fff;
                    border-radius: 0px;
                    font-weight: bold;
                    font-size: 14px;
                    padding: 0px;
                    text-align: center;
                }
                QPushButton:hover { background-color: #2fa74c; }
            """)

            bt_excluir = QPushButton("Excluir")
            bt_excluir.setMinimumWidth(90)
            bt_excluir.setMaximumWidth(90)
            bt_excluir.setMinimumHeight(36)
            bt_excluir.setMaximumHeight(36)
            bt_excluir.setStyleSheet("""
                QPushButton {
                    background-color: #ff4d6d;
                    color: #fff;
                    border-radius: 0px;
                    font-weight: bold;
                    font-size: 14px;
                    padding: 0px;
                    text-align: center;
                }
                QPushButton:hover { background-color: #d93c5c; }
            """)
# ...restante do método...
            bt_abrir.clicked.connect(lambda _, x=op_id: self.abrir_op_callback(x))
            bt_excluir.clicked.connect(lambda _, x=op_id: self._excluir(x))
            bt_export.clicked.connect(lambda _, x=op_id: self._exportar_csv(x))

            lay.addWidget(bt_abrir)
            lay.addWidget(bt_export)
            lay.addWidget(bt_excluir)
            acao_widget.setMinimumWidth(480)
            self.tabela.setCellWidget(r, 5, acao_widget)

        # Após preencher a tabela, ajuste o tamanho da coluna de ações:
        self.tabela.setColumnWidth(5, 540)  # Aumente para 540 ou mais

    def _duplo_clique(self, row, _col):
        op_id = int(self.tabela.item(row, 0).text())
        self.abrir_op_callback(op_id)

    def _excluir(self, op_id: int):
        r = QMessageBox.question(self, "Confirmar", f"Excluir OP {op_id}? Esta ação não pode ser desfeita.")
        if r != QMessageBox.Yes:
            return
        conn = sqlite3.connect(DATABASE_PATH)
        c = conn.cursor()
        c.execute("DELETE FROM taloes WHERE op_id = ?", (op_id,))
        c.execute("DELETE FROM ops WHERE id = ?", (op_id,))
        conn.commit()
        conn.close()
        self.atualizar()

    def _exportar_csv(self, op_id: int):
        caminho, _ = QFileDialog.getSaveFileName(self, "Salvar CSV", f"op_{op_id}.csv", "CSV (*.csv)")
        if not caminho:
            return
        # Exporta cabeçalho + todas as linhas de talões
        conn = sqlite3.connect(DATABASE_PATH)
        c = conn.cursor()
        c.execute("SELECT cliente, num_op, data_criacao, total_pares FROM ops WHERE id = ?", (op_id,))
        op = c.fetchone()
        c.execute(
            "SELECT giro, talao_num, numeracao, quantidade FROM taloes WHERE op_id = ? ORDER BY giro, talao_num, numeracao",
            (op_id,),
        )
        linhas = c.fetchall()
        conn.close()
        with open(caminho, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["OP ID", op_id])
            if op:
                w.writerow(["Cliente", op[0]])
                w.writerow(["Nº OP", op[1]])
                w.writerow(["Criada em", op[2]])
                w.writerow(["Total Pares", op[3]])
            w.writerow([])
            w.writerow(["Giro", "Talão", "Tamanho", "Quantidade"]) 
            for giro, talao, tam, qtd in linhas:
                w.writerow([giro, talao, tam, qtd])
        QMessageBox.information(self, "Exportar", "CSV gerado com sucesso!")


class CriarOPPage(QWidget):
    def __init__(self, on_created_callback):
        super().__init__()
        self.on_created_callback = on_created_callback
        self._setup_ui()

    def _setup_ui(self):
        self.setStyleSheet(APP_QSS)
        root = QVBoxLayout(self)

        title = QLabel("Criar Nova OP")
        title.setProperty("cls", "title")
        root.addWidget(title)

        form_group = QGroupBox("Dados da OP")
        form_layout = QFormLayout(form_group)
        self.cliente_input = QLineEdit()
        self.num_op_input = QLineEdit()
        self.total_pares_input = QLineEdit()
        self.total_pares_input.setPlaceholderText(f"Múltiplos de {PARES_POR_TALAO} serão distribuídos em talões")
        form_layout.addRow("Cliente:", self.cliente_input)
        form_layout.addRow("Número da OP:", self.num_op_input)
        form_layout.addRow("Total de Pares:", self.total_pares_input)

        self.tipo_input = QComboBox()
        self.tipo_input.addItems(["Masculino", "Feminino"])
        form_layout.addRow("Tipo:", self.tipo_input)

        root.addWidget(form_group)

        botoes = QHBoxLayout()
        salvar_btn = QPushButton("Salvar OP")
        salvar_btn.clicked.connect(self.salvar_op)
        limpar_btn = QPushButton("Limpar")
        limpar_btn.setProperty("secondary", True)
        limpar_btn.clicked.connect(self._limpar)
        botoes.addWidget(salvar_btn)
        botoes.addWidget(limpar_btn)
        botoes.addStretch()
        root.addLayout(botoes)
        root.addStretch()

        salvar_btn.setMinimumHeight(32)
        salvar_btn.setMaximumHeight(36)
        limpar_btn.setMinimumHeight(32)
        limpar_btn.setMaximumHeight(36)

    def _limpar(self):
        self.cliente_input.clear()
        self.num_op_input.clear()
        self.total_pares_input.clear()

    def salvar_op(self):
        cliente = self.cliente_input.text().strip()
        num_op_txt = self.num_op_input.text().strip()
        total_txt = self.total_pares_input.text().strip()

        if not cliente or not num_op_txt.isdigit() or not total_txt.isdigit():
            QMessageBox.warning(self, "Dados inválidos", "Preencha todos os campos corretamente.")
            return

        num_op = int(num_op_txt)
        total_pares = int(total_txt)
        if total_pares <= 0:
            QMessageBox.warning(self, "Dados inválidos", "Total de pares deve ser maior que zero.")
            return

        tipo = self.tipo_input.currentText()
        try:
            op_id = inserir_op(cliente, num_op, total_pares, tipo)
            gerar_taloes_iniciais(op_id, total_pares, tipo)
        except sqlite3.IntegrityError as e:
            QMessageBox.critical(self, "Erro", f"Não foi possível salvar. Nº OP já existente?\n\n{e}")
            return
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao salvar OP.\n\n{e}")
            return

        QMessageBox.information(self, "Sucesso", f"OP criada (ID {op_id}). Agora você pode distribuir os pares por tamanho.")
        self.on_created_callback(op_id)


class VisualizarOPPage(QWidget):
    def __init__(self, op_id: int, voltar_callback=None):
        super().__init__()
        self.op_id = op_id
        self.voltar_callback = voltar_callback
        self._setup_ui()
        self._carregar()

    def _setup_ui(self):
        self.setStyleSheet(APP_QSS)
        self.layout = QVBoxLayout(self)

        header = QHBoxLayout()
        self.lb_title = QLabel(f"OP {self.op_id}")
        self.lb_title.setProperty("cls", "title")
        header.addWidget(self.lb_title)
        header.addStretch()

        # Botão Voltar
        self.bt_voltar = QPushButton("Voltar")
        self.bt_voltar.setMinimumWidth(100)
        self.bt_voltar.setMaximumWidth(140)
        self.bt_voltar.setMinimumHeight(32)
        self.bt_voltar.setMaximumHeight(36)
        self.bt_voltar.setStyleSheet("""
            QPushButton {
                background-color: #222;
                color: #fff;
                border-radius: 8px;
                font-weight: bold;
                font-size: 14px;
                padding: 0px;
            }
            QPushButton:hover { background-color: #444; }
        """)
        self.bt_voltar.clicked.connect(self._voltar)
        header.addWidget(self.bt_voltar)

        self.bt_salvar = QPushButton("Salvar alterações")
        self.bt_salvar.clicked.connect(self._salvar)
        self.bt_exportar = QPushButton("Exportar CSV")
        self.bt_exportar.clicked.connect(self._exportar_csv)

        header.addWidget(self.bt_exportar)
        header.addWidget(self.bt_salvar)
        self.layout.addLayout(header)

        self.abas = QTabWidget()
        self.layout.addWidget(self.abas)

        # Rodapé totalizador
        rodape = QHBoxLayout()
        self.lb_resumo = QLabel()
        self.lb_resumo.setProperty("cls", "subtitle")
        rodape.addWidget(self.lb_resumo)
        rodape.addStretch()
        self.layout.addLayout(rodape)

    def _carregar(self):
        # Cabeçalho da OP
        conn = sqlite3.connect(DATABASE_PATH)
        c = conn.cursor()
        c.execute("SELECT cliente, num_op, data_criacao, total_pares FROM ops WHERE id = ?", (self.op_id,))
        row = c.fetchone()
        conn.close()
        if row:
            cliente, num_op, data_criacao, total_pares = row
            self.lb_title.setText(f"OP {self.op_id} · Cliente: {cliente} · Nº OP: {num_op} · Criada em: {data_criacao} · Total informado: {total_pares}")

        # Montar tabelas por GIRO
        self.abas.clear()
        self.giros_data = carregar_taloes(self.op_id)
        self.tabelas_por_giro: Dict[int, QTableWidget] = {}

        for giro in GIROS:
            taloes = self.giros_data.get(giro, {})
            pagina = QWidget()
            vbox = QVBoxLayout(pagina)
            vbox.setContentsMargins(0, 0, 0, 0)  # Remover margens

            headers = ["Talão", "Talão id", "Numeração"] + TAMANHOS + ["Qtd. Produto", "TOTAL TALÃO", "Status", "Ações"]
            tabela = QTableWidget()
            tabela.setColumnCount(len(headers))
            tabela.setHorizontalHeaderLabels(headers)
            tabela.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            tabela.horizontalHeader().setSectionResizeMode(len(headers) - 1, QHeaderView.Fixed)
            tabela.setColumnWidth(len(headers) - 1, 260)  # Largura maior para os botões
            tabela.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            tabela.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            tabela.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            tabela.setStyleSheet("QTableWidget { background: #000; selection-background-color: #7c4dff; }")
            vbox.addWidget(tabela)
            self.abas.addTab(pagina, f"Giro {giro}")
            self.tabelas_por_giro[giro] = tabela

            row_count = len(taloes) + 1
            tabela.setRowCount(row_count)

            total_por_tamanho = {num: 0 for num in TAMANHOS}

            for r, (talao_num, tamanhos) in enumerate(taloes.items()):
                tabela.setItem(r, 0, QTableWidgetItem(str(talao_num)))
                tabela.setItem(r, 1, QTableWidgetItem(f"ID_{self.op_id}_{giro}_{talao_num}"))
                tabela.setItem(r, 2, QTableWidgetItem("/".join([k for k, v in tamanhos.items() if v > 0])))

                soma_linha = 0
                for c, tam in enumerate(TAMANHOS, start=3):
                    valor = int(tamanhos.get(tam, 0))
                    item = QTableWidgetItem(str(valor))
                    item.setTextAlignment(Qt.AlignCenter)
                    item.setFlags(item.flags() | Qt.ItemIsEditable)
                    tabela.setItem(r, c, item)
                    soma_linha += valor
                    total_por_tamanho[tam] += valor

                tabela.setItem(r, len(TAMANHOS) + 3, QTableWidgetItem(str(soma_linha)))
                total_item = QTableWidgetItem(str(soma_linha))
                total_item.setTextAlignment(Qt.AlignCenter)
                total_item.setFlags(Qt.ItemIsEnabled)
                tabela.setItem(r, len(headers) - 1, total_item)

                # Status
                status_item = QTableWidgetItem("Pendente")
                status_item.setBackground(Qt.red)
                tabela.setItem(r, len(headers) - 2, status_item)
                # Botão de ação (somente nas linhas de talão, não na linha TOTAL LOTE)
                if r < len(taloes):  # Adiciona apenas nas linhas de talão
                    acao_widget = QWidget()
                    lay = QHBoxLayout(acao_widget)
                    lay.setContentsMargins(0, 0, 0, 0)
                    lay.setSpacing(8)
                    lay.setAlignment(Qt.AlignVCenter)  # Centraliza verticalmente

                    # Botão Pendente
                    bt_pendente = QToolButton()
                    bt_pendente.setText("Pendente")
                    bt_pendente.setMinimumWidth(100)
                    bt_pendente.setMaximumWidth(100)
                    bt_pendente.setMinimumHeight(40)
                    bt_pendente.setMaximumHeight(40)
                    bt_pendente.setStyleSheet("""
                        QToolButton {
                            background-color: #ff4d6d;
                            color: #fff;
                            border-radius: 0px;
                            font-weight: bold;
                            font-size: 15px;
                            padding: 0px;
                            text-align: center;
                        }
                        QToolButton:pressed {
                            background-color: #d93c5c;
                        }
                    """)

                    # Botão OK
                    bt_status = QToolButton()
                    bt_status.setText("OK")
                    bt_status.setMinimumWidth(100)
                    bt_status.setMaximumWidth(100)
                    bt_status.setMinimumHeight(40)
                    bt_status.setMaximumHeight(40)
                    bt_status.setStyleSheet("""
                        QToolButton {
                            background-color: #43d96b;
                            color: #fff;
                            border-radius: 0px;
                            font-weight: bold;
                            font-size: 15px;
                            padding: 0px;
                            text-align: center;
                        }
                        QToolButton:pressed {
                            background-color: #2fa74c;
                        }
                    """)

                    lay.addWidget(bt_pendente)
                    lay.addWidget(bt_status)
                    acao_widget.setMinimumHeight(40)
                    acao_widget.setMaximumHeight(40)
                    acao_widget.setMinimumWidth(200)
                    acao_widget.setMaximumWidth(200)
                    tabela.setCellWidget(r, len(headers) - 1, acao_widget)
            tabela.setColumnWidth(len(headers) - 1, 200)  # Ajusta largura da coluna de ações
        # ...restante do método...

        self._atualizar_resumo()

    def _on_cell_changed(self, row: int, col: int):
        # Ignora se é a linha totalizadora ou a coluna não editável
        tabela: QTableWidget = self.abas.currentWidget().layout().itemAt(0).widget()  # tabela da aba atual
        last_row = tabela.rowCount() - 1
        if row == last_row or col == 0 or col == tabela.columnCount() - 1:
            return

        # Atualiza TOTAL da linha
        soma = 0
        for c in range(1, tabela.columnCount() - 1):
            try:
                soma += int(tabela.item(row, c).text())
            except Exception:
                pass
        total_item = QTableWidgetItem(str(soma))
        total_item.setTextAlignment(Qt.AlignCenter)
        total_item.setFlags(Qt.ItemIsEnabled)
        tabela.setItem(row, tabela.columnCount() - 1, total_item)

        # Recalcula total por coluna
        self._recalcular_totais_da_tabela(tabela)
        self._atualizar_resumo()

    def _recalcular_totais_da_tabela(self, tabela: QTableWidget):
        last_row = tabela.rowCount() - 1
        # zera
        for c in range(1, tabela.columnCount() - 1):
            tabela.setItem(last_row, c, QTableWidgetItem("0"))
            tabela.item(last_row, c).setFlags(Qt.ItemIsEnabled)
            tabela.item(last_row, c).setTextAlignment(Qt.AlignCenter)
        # soma
        for r in range(0, last_row):
            for c in range(1, tabela.columnCount() - 1):
                try:
                    val = int(tabela.item(r, c).text())
                except Exception:
                    val = 0
                cur = int(tabela.item(last_row, c).text())
                tabela.item(last_row, c).setText(str(cur + val))
        # total geral
        soma_geral = 0
        for c in range(1, tabela.columnCount() - 1):
            soma_geral += int(tabela.item(last_row, c).text())
        total_item = QTableWidgetItem(str(soma_geral))
        total_item.setFlags(Qt.ItemIsEnabled)
        total_item.setTextAlignment(Qt.AlignCenter)
        tabela.setItem(last_row, tabela.columnCount() - 1, total_item)

    def _validar(self) -> bool:
        # Cada talão deve somar exatamente PARES_POR_TALAO
        for giro, tabela in self.tabelas_por_giro.items():
            last_row = tabela.rowCount() - 1
            for r in range(0, last_row):
                soma = 0
                for c in range(1, tabela.columnCount() - 1):
                    try:
                        soma += int(tabela.item(r, c).text())
                    except Exception:
                        pass
                if soma != PARES_POR_TALAO:
                    QMessageBox.warning(
                        self,
                        "Validação",
                        f"No Giro {giro}, Talão {tabela.item(r,0).text()} soma {soma}.\nCada talão deve somar {PARES_POR_TALAO} pares.",
                    )
                    return False
        return True

    def _salvar(self):
        if not self._validar():
            return
        # Coleta dados editados
        for giro, tabela in self.tabelas_por_giro.items():
            last_row = tabela.rowCount() - 1
            for r in range(0, last_row):
                talao_num = int(tabela.item(r, 0).text())
                for c, tam in enumerate(TAMANHOS, start=1):
                    try:
                        val = int(tabela.item(r, c).text())
                    except Exception:
                        val = 0
                    self.giros_data[giro][talao_num][tam] = val
        try:
            salvar_taloes(self.op_id, self.giros_data)
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao salvar alterações.\n\n{e}")
            return
        QMessageBox.information(self, "Salvo", "Alterações gravadas com sucesso.")
        self._atualizar_resumo()

    def _exportar_csv(self):
        caminho, _ = QFileDialog.getSaveFileName(self, "Salvar CSV da OP", f"op_{self.op_id}.csv", "CSV (*.csv)")
        if not caminho:
            return
        # Cabeçalho
        conn = sqlite3.connect(DATABASE_PATH)
        c = conn.cursor()
        c.execute("SELECT cliente, num_op, data_criacao, total_pares FROM ops WHERE id = ?", (self.op_id,))
        op = c.fetchone()
        c.execute(
            "SELECT giro, talao_num, numeracao, quantidade FROM taloes WHERE op_id = ? ORDER BY giro, talao_num, numeracao",
            (self.op_id,),
        )
        linhas = c.fetchall()
        conn.close()
        with open(caminho, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["OP ID", self.op_id])
            if op:
                w.writerow(["Cliente", op[0]])
                w.writerow(["Nº OP", op[1]])
                w.writerow(["Criada em", op[2]])
                w.writerow(["Total Pares (pedido)", op[3]])
            w.writerow([])
            w.writerow(["Giro", "Talão"] + TAMANHOS + ["TOTAL TALÃO"]) 

            # Gera por giro/talão conforme tela (já valida somas)
            for giro in GIROS:
                taloes = self.giros_data.get(giro, {})
                for talao_num, tamanhos in taloes.items():
                    row = [giro, talao_num]
                    soma = 0
                    for tam in TAMANHOS:
                        qtd = int(tamanhos.get(tam, 0))
                        row.append(qtd)
                        soma += qtd
                    row.append(soma)
                    w.writerow(row)
        QMessageBox.information(self, "Exportar", "CSV da OP gerado com sucesso!")

    def _atualizar_resumo(self):
        total_geral = 0
        for tabela in self.tabelas_por_giro.values():
            last_row = tabela.rowCount() - 1
            try:
                total_geral += int(tabela.item(last_row, tabela.columnCount() - 1).text())
            except Exception:
                pass
        self.lb_resumo.setText(f"Total geral (somando todos os GIROS): {total_geral} pares")

    def _excluir_talao(self, talao_num, giro):
        # Implemente a lógica de exclusão do talão específico do giro
        pass

    def _alternar_status(self, row, giro):
        # Implemente a lógica de alternar status do talão específico do giro
        pass

    def _voltar(self):
        if self.voltar_callback:
            self.voltar_callback()


# ================
# Janela Principal
# ================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("OP de Calçados · Profissional")
        self.setMinimumSize(1100, 700)
        self.stack = QStackedWidget()
        self.stack.setContentsMargins(0, 0, 0, 0)  # Remove margens do stack
        self.setCentralWidget(self.stack)

        self.page_lista = ListaOPsPage(self.abrir_op, self.nova_op)
        self.stack.addWidget(self.page_lista)

        # Status bar
        self.statusBar().showMessage("Pronto")

    def nova_op(self):
        self.page_criar = CriarOPPage(self._on_op_criada)
        self.stack.addWidget(self.page_criar)
        self.stack.setCurrentWidget(self.page_criar)

    def _on_op_criada(self, op_id: int):
        # Após criar, abre diretamente a tela de edição da OP
        self.abrir_op(op_id)
        # E atualiza a lista ao voltar para ela
        self.page_lista.atualizar()

    def abrir_op(self, op_id: int):
        self.page_visualizar = VisualizarOPPage(op_id, voltar_callback=self._voltar_lista)
        self.stack.addWidget(self.page_visualizar)
        self.stack.setCurrentWidget(self.page_visualizar)

    def _voltar_lista(self):
        self.stack.setCurrentWidget(self.page_lista)
        self.page_lista.atualizar()


# =====
# Main
# =====

def main():
    criar_banco()
    app = QApplication(sys.argv)
    app.setStyleSheet(APP_QSS)
    win = MainWindow()
    win.showMaximized()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()

def anexar_arquivo(caminho_arquivo):
    # Aguarde o WhatsApp Web carregar a conversa
    time.sleep(2)
    # Clique no ícone de clipe (ajuste as coordenadas conforme seu monitor)
    pyautogui.click(x=1200, y=950)  # Troque x e y pelas coordenadas do ícone de clipe
    time.sleep(1)
    # Clique no ícone de documento (ajuste as coordenadas conforme seu monitor)
    pyautogui.click(x=1150, y=800)  # Troque x e y pelas coordenadas do ícone de documento
    time.sleep(2)
    # Digite o caminho do arquivo no seletor de arquivos do Windows
    pyautogui.write(caminho_arquivo)
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(5)  # Aguarde o upload do arquivo
    # Clique no botão de enviar (ícone de avião de papel)
    pyautogui.click(x=1250, y=950)  # Troque x e y pelas coordenadas do botão de enviar
    time.sleep(2)
