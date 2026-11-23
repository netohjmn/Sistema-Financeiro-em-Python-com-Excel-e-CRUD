"""Sistema Financeiro em Python e Excel.

Controle de itens e movimentações com MySQL, interface Tkinter, gráficos
Matplotlib e exportação de relatórios formatados para planilha (Google Sheets,
que abre e exporta para Excel).
"""

import os
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

import gspread
import matplotlib
import mysql.connector
import pandas as pd
from google.oauth2.service_account import Credentials

matplotlib.use("TkAgg")

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure


# --------------------------------------------------------------------------
# Configuração
# --------------------------------------------------------------------------

APP_TITLE = "Sistema Financeiro"
APP_SLOGAN = "Controle claro, números confiáveis."

DB_CONFIG = {
    "host": os.getenv("DB_HOST", "localhost"),
    "user": os.getenv("DB_USER", "root"),
    "password": os.getenv("DB_PASSWORD", ""),
    "database": os.getenv("DB_NAME", "sistema_financeiro"),
}
CREDENTIALS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE", "")
SPREADSHEET_ID = os.getenv("GOOGLE_SHEETS_ID", "")
SPREADSHEET_NAME = "relatorio_financeiro"
GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
SHEET_DETAIL = "Dados Consolidados"
SHEET_SUMMARY = "Resumo Gerencial"

PALETTE = {
    "bg": "#F4F5F9",
    "card": "#FFFFFF",
    "text": "#1B2033",
    "muted": "#7A8099",
    "primary": "#2D3A8C",
    "primary_dark": "#1A235E",
    "primary_soft": "#E8EBFA",
    "line": "#E1E4EF",
    "danger": "#C0392B",
    "accent": "#E08E2B",
    "stripe": "#F7F8FC",
}
FONT = "Segoe UI"
MONEY_WORDS = ("preco", "valor", "total")
STOCK_DELTA = "CASE WHEN {t}tipo = 'Entrada' THEN {t}quantidade ELSE -{t}quantidade END"

PRODUCTS_SQL = f"""
    SELECT p.id, p.nome, p.categoria, p.preco_unitario, p.estoque_inicial,
           p.estoque_inicial + COALESCE(SUM({STOCK_DELTA.format(t='m.')}), 0) AS estoque_atual
    FROM produtos AS p
    LEFT JOIN movimentacoes AS m ON m.produto_id = p.id
    GROUP BY p.id, p.nome, p.categoria, p.preco_unitario, p.estoque_inicial
    ORDER BY p.id
"""
MOVEMENTS_SQL = """
    SELECT m.id, m.data_movimento, p.nome AS produto, p.categoria, m.tipo,
           m.quantidade, m.valor_unitario, m.observacao, m.produto_id
    FROM movimentacoes AS m
    JOIN produtos AS p ON p.id = m.produto_id
    ORDER BY m.data_movimento DESC, m.id DESC
"""
DETAIL_SQL = """
    SELECT m.id AS movimentacao_id, m.data_movimento, p.id AS produto_id,
           p.nome AS produto, p.categoria, m.tipo, m.quantidade,
           m.valor_unitario, m.quantidade * m.valor_unitario AS valor_total,
           m.observacao
    FROM movimentacoes AS m
    JOIN produtos AS p ON p.id = m.produto_id
    ORDER BY m.data_movimento, m.id
"""
CATEGORY_CHART_SQL = """
    SELECT p.categoria, SUM(m.quantidade) AS quantidade
    FROM movimentacoes AS m
    JOIN produtos AS p ON p.id = m.produto_id
    GROUP BY p.categoria ORDER BY p.categoria
"""
MONTHLY_CHART_SQL = f"""
    SELECT DATE_FORMAT(data_movimento, '%Y-%m') AS periodo,
           SUM({STOCK_DELTA.format(t='')}) AS saldo
    FROM movimentacoes
    GROUP BY DATE_FORMAT(data_movimento, '%Y-%m')
    ORDER BY periodo
"""


# --------------------------------------------------------------------------
# Especificação das telas (a interface é gerada a partir destes dados)
# --------------------------------------------------------------------------

def brl(value):
    return f"R$ {float(value):.2f}".replace(".", ",")


ENTITIES = {
    "product": {
        "noun": "produto",
        "heading": "Catálogo de produtos",
        "new_label": "+ Novo produto",
        "tab": "Produtos",
        "columns": [
            ("id", "ID", 55, lambda r: r.id),
            ("nome", "Produto", 220, lambda r: r.nome),
            ("categoria", "Categoria", 150, lambda r: r.categoria),
            ("preco", "Preço unitário", 130, lambda r: brl(r.preco_unitario)),
            ("inicial", "Estoque inicial", 120, lambda r: r.estoque_inicial),
            ("atual", "Estoque atual", 120, lambda r: r.estoque_atual),
        ],
    },
    "movement": {
        "noun": "movimentação",
        "heading": "Histórico de movimentações",
        "new_label": "+ Nova movimentação",
        "tab": "Movimentações",
        "columns": [
            ("id", "ID", 45, lambda r: r.id),
            ("data", "Data", 100, lambda r: r.data_movimento),
            ("produto", "Produto", 170, lambda r: r.produto),
            ("categoria", "Categoria", 110, lambda r: r.categoria),
            ("tipo", "Tipo", 75, lambda r: r.tipo),
            ("quantidade", "Quantidade", 85, lambda r: r.quantidade),
            ("valor", "Valor unitário", 110, lambda r: brl(r.valor_unitario)),
            ("observacao", "Observação", 200, lambda r: r.observacao),
        ],
    },
}

METRICS = (
    ("products", "PRODUTOS CADASTRADOS"),
    ("stock", "UNIDADES EM ESTOQUE"),
    ("movements", "MOVIMENTAÇÕES"),
)


# --------------------------------------------------------------------------
# Acesso ao banco
# --------------------------------------------------------------------------

def open_connection():
    return mysql.connector.connect(**DB_CONFIG)


def run_query(cursor, sql, params=None):
    cursor.execute(sql, params or ())
    names = [col[0] for col in cursor.description]
    return pd.DataFrame(cursor.fetchall(), columns=names)


def fetch_frames(*queries):
    connection = open_connection()
    try:
        cursor = connection.cursor()
        frames = [run_query(cursor, sql) for sql in queries]
        cursor.close()
    finally:
        connection.close()
    return frames


def stock_balance(cursor, product_id, skip_movement=None):
    extra, params = "", [product_id]
    if skip_movement is not None:
        extra, params = " AND id <> %s", [skip_movement, product_id]
    cursor.execute(
        f"""
        SELECT estoque_inicial + COALESCE((
            SELECT SUM({STOCK_DELTA.format(t='')})
            FROM movimentacoes
            WHERE produto_id = produtos.id{extra}
        ), 0) AS saldo
        FROM produtos
        WHERE id = %s
        """,
        params,
    )
    row = cursor.fetchone()
    if row is None:
        raise ValueError("O produto selecionado não existe.")
    return int(row[0])


class Transaction:
    """Abre uma conexão, confirma ao final e desfaz em caso de erro."""

    def __enter__(self):
        self.connection = open_connection()
        self.cursor = self.connection.cursor()
        return self.cursor

    def __exit__(self, exc_type, exc, tb):
        try:
            if exc_type is None:
                self.connection.commit()
            else:
                self.connection.rollback()
        finally:
            self.connection.close()
        return False


# --------------------------------------------------------------------------
# Relatórios e exportação
# --------------------------------------------------------------------------

def build_reports():
    detail, products, movements = fetch_frames(DETAIL_SQL, PRODUCTS_SQL, MOVEMENTS_SQL)

    if products.empty:
        by_category = pd.DataFrame(
            columns=["indicador", "grupo", "produtos", "unidades_em_estoque", "valor_estoque"]
        )
    else:
        by_category = products.groupby("categoria", as_index=False).agg(
            produtos=("id", "count"),
            unidades_em_estoque=("estoque_atual", "sum"),
        )
        stock_value = (
            products.assign(
                valor_estoque=products["preco_unitario"].astype(float)
                * products["estoque_atual"].astype(int)
            )
            .groupby("categoria")["valor_estoque"]
            .sum()
        )
        by_category["valor_estoque"] = by_category["categoria"].map(stock_value)
        by_category.insert(0, "indicador", "Estoque por categoria")
        by_category = by_category.rename(columns={"categoria": "grupo"})

    if movements.empty:
        by_type = pd.DataFrame(
            columns=["tipo", "quantidade_total", "media_por_movimentacao", "movimentacoes"]
        )
    else:
        by_type = movements.groupby("tipo", as_index=False).agg(
            quantidade_total=("quantidade", "sum"),
            media_por_movimentacao=("quantidade", "mean"),
            movimentacoes=("id", "count"),
        )
        by_type.insert(0, "indicador", "Movimentações por tipo")
        by_type = by_type.rename(columns={"tipo": "grupo"})

    summary = pd.concat([by_category, by_type], ignore_index=True, sort=False)
    return detail, summary


def to_rgb(hex_color):
    hex_color = hex_color.lstrip("#")
    red, green, blue = (int(hex_color[i:i + 2], 16) / 255 for i in (0, 2, 4))
    return {"red": red, "green": green, "blue": blue}


def sheet_cell(value):
    if pd.isna(value):
        return ""
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, Decimal):
        return float(value)
    if hasattr(value, "item"):
        return value.item()
    return value


def frame_to_rows(frame):
    return [[sheet_cell(v) for v in row] for row in frame.itertuples(index=False, name=None)]


def authorize_client():
    if not CREDENTIALS_FILE:
        raise ValueError(
            "Configure a variável GOOGLE_CREDENTIALS_FILE com o caminho "
            "do JSON da conta de serviço do Google."
        )
    path = Path(CREDENTIALS_FILE)
    if not path.is_file():
        raise FileNotFoundError(f"Arquivo de credenciais não encontrado: {path}")
    credentials = Credentials.from_service_account_file(str(path), scopes=GOOGLE_SCOPES)
    return gspread.authorize(credentials)


def open_spreadsheet(client):
    if SPREADSHEET_ID:
        return client.open_by_key(SPREADSHEET_ID)
    try:
        return client.open(SPREADSHEET_NAME)
    except gspread.SpreadsheetNotFound:
        return client.create(SPREADSHEET_NAME)


def get_worksheet(spreadsheet, title, frame):
    try:
        return spreadsheet.worksheet(title)
    except gspread.WorksheetNotFound:
        if title == SHEET_DETAIL and len(spreadsheet.worksheets()) == 1:
            sheet = spreadsheet.sheet1
            sheet.update_title(title)
            return sheet
        return spreadsheet.add_worksheet(
            title=title, rows=max(len(frame) + 1, 100), cols=max(len(frame.columns), 10)
        )


def export_reports(detail, summary):
    spreadsheet = open_spreadsheet(authorize_client())
    for title, frame in ((SHEET_DETAIL, detail), (SHEET_SUMMARY, summary)):
        sheet = get_worksheet(spreadsheet, title, frame)
        sheet.clear()
        sheet.update(
            range_name="A1",
            values=[frame.columns.tolist()] + frame_to_rows(frame),
            value_input_option="USER_ENTERED",
        )
        style_worksheet(sheet, frame)
    return spreadsheet.url


def style_worksheet(sheet, frame):
    rows = max(len(frame) + 1, 2)
    cols = max(len(frame.columns), 1)
    sheet.resize(rows=rows, cols=cols)
    last_cell = gspread.utils.rowcol_to_a1(rows, cols)
    last_head = gspread.utils.rowcol_to_a1(1, cols)
    head_color = to_rgb(PALETTE["primary_dark"])
    white = {"red": 1, "green": 1, "blue": 1}

    sheet.format(f"A1:{last_cell}", {
        "textFormat": {"fontFamily": "Calibri", "fontSize": 10,
                       "foregroundColor": to_rgb(PALETTE["text"])},
        "verticalAlignment": "MIDDLE",
    })
    sheet.format(f"A1:{last_head}", {
        "backgroundColor": head_color,
        "textFormat": {"fontFamily": "Calibri", "fontSize": 10, "bold": True,
                       "foregroundColor": white},
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE",
        "wrapStrategy": "WRAP",
    })
    sheet.format(f"A2:{last_cell}", {
        "verticalAlignment": "MIDDLE",
        "borders": {"bottom": {"style": "SOLID", "color": to_rgb(PALETTE["line"])}},
    })
    sheet.freeze(rows=1)
    sheet.columns_auto_resize(0, cols)

    metadata = sheet.spreadsheet.fetch_sheet_metadata(
        params={"fields": "sheets(properties(sheetId),bandedRanges(bandedRangeId))"}
    )
    current = next(s for s in metadata["sheets"] if s["properties"]["sheetId"] == sheet.id)
    requests = [
        {"deleteBanding": {"bandedRangeId": band["bandedRangeId"]}}
        for band in current.get("bandedRanges", [])
    ]
    requests += [
        {"addBanding": {"bandedRange": {
            "range": {"sheetId": sheet.id, "startRowIndex": 0, "endRowIndex": rows,
                      "startColumnIndex": 0, "endColumnIndex": cols},
            "rowProperties": {
                "headerColor": head_color,
                "firstBandColor": white,
                "secondBandColor": to_rgb(PALETTE["primary_soft"]),
            },
        }}},
        {"updateSheetProperties": {
            "properties": {"sheetId": sheet.id, "gridProperties": {"hideGridlines": True}},
            "fields": "gridProperties.hideGridlines",
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sheet.id, "dimension": "ROWS", "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 42},
            "fields": "pixelSize",
        }},
    ]
    sheet.batch_update({"requests": requests})

    for position, name in enumerate(frame.columns, start=1):
        letter = gspread.utils.rowcol_to_a1(1, position)[:-1]
        target = f"{letter}2:{letter}{rows}"
        lowered = str(name).lower()
        if any(word in lowered for word in MONEY_WORDS):
            number_format = {"type": "CURRENCY", "pattern": '"R$" #,##0.00'}
        elif "data" in lowered:
            number_format = {"type": "DATE", "pattern": "dd/mm/yyyy"}
        elif pd.api.types.is_numeric_dtype(frame[name]):
            number_format = {"type": "NUMBER", "pattern": "#,##0.##"}
        else:
            continue
        sheet.format(target, {"numberFormat": number_format})


# --------------------------------------------------------------------------
# Validação
# --------------------------------------------------------------------------

def parse_int(value, label, minimum):
    try:
        number = int(value)
    except (TypeError, ValueError) as error:
        raise ValueError(f"{label} deve ser um número inteiro.") from error
    if number < minimum:
        raise ValueError(f"{label} deve ser no mínimo {minimum}.")
    return number


def parse_money(value):
    try:
        number = Decimal(value.replace(",", "."))
    except (InvalidOperation, AttributeError) as error:
        raise ValueError("Informe um valor monetário válido.") from error
    if not number.is_finite() or number < 0:
        raise ValueError("O valor monetário não pode ser negativo.")
    return number


def thousands(number):
    return f"{number:,}".replace(",", ".")


# --------------------------------------------------------------------------
# Interface
# --------------------------------------------------------------------------

class FinanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_TITLE} | Python e Excel")
        self.root.geometry("1240x800")
        self.root.minsize(1000, 680)
        self.products = pd.DataFrame()
        self.movements = pd.DataFrame()
        self.trees = {}
        self.metric_labels = {}
        self.chart_canvas = None
        self._apply_theme()
        self._build_layout()
        self.reload()

    # ---- tema --------------------------------------------------------

    def _apply_theme(self):
        c = PALETTE
        self.root.configure(bg=c["bg"])
        style = ttk.Style(self.root)
        style.theme_use("clam")

        def button(name, bg, fg, border, hover, pressed, pad=(12, 8)):
            style.configure(name, padding=pad, background=bg, foreground=fg,
                            bordercolor=border, font=(FONT, 9, "bold"))
            style.map(name, background=[("active", hover), ("pressed", pressed)],
                      foreground=[("disabled", c["muted"])])

        style.configure("TFrame", background=c["bg"])
        style.configure("TLabel", background=c["bg"], foreground=c["text"], font=(FONT, 10))
        button("TButton", c["card"], c["text"], c["line"], c["primary_soft"], "#D5DAF2")
        button("Accent.TButton", c["primary"], "#FFFFFF", c["primary"], "#232F78",
               c["primary_dark"], pad=(14, 9))
        button("Danger.TButton", "#FBECEA", c["danger"], "#F0D2CE", "#F6DCD8", "#EDCBC5")

        style.configure("TNotebook", background=c["bg"], borderwidth=0, tabmargins=(0, 0, 0, 0))
        style.configure("TNotebook.Tab", padding=(18, 11), background="#E6E9F4",
                        foreground=c["muted"], font=(FONT, 10, "bold"), borderwidth=0)
        style.map("TNotebook.Tab",
                  background=[("selected", c["card"]), ("active", "#EEF0F9")],
                  foreground=[("selected", c["primary_dark"]), ("active", c["text"])])
        style.configure("Treeview", background=c["card"], fieldbackground=c["card"],
                        foreground=c["text"], rowheight=36, borderwidth=0, font=(FONT, 9))
        style.configure("Treeview.Heading", background="#EBEEF8", foreground=c["muted"],
                        padding=(10, 11), relief="flat", font=(FONT, 9, "bold"))
        style.map("Treeview", background=[("selected", c["primary"])],
                  foreground=[("selected", "#FFFFFF")])
        style.map("Treeview.Heading", background=[("active", "#DEE3F4")])
        style.configure("Vertical.TScrollbar", background="#D9DEF0", troughcolor=c["card"],
                        bordercolor=c["card"], arrowsize=12)

    # ---- layout ------------------------------------------------------

    def _build_layout(self):
        c = PALETTE
        header = tk.Frame(self.root, bg=c["primary_dark"], padx=28, pady=20)
        header.pack(fill="x")
        brand = tk.Frame(header, bg=c["primary_dark"])
        brand.pack(side="left", fill="y")
        tk.Label(brand, text=APP_TITLE.upper(), bg=c["primary_dark"], fg="#FFFFFF",
                 font=(FONT, 19, "bold")).pack(anchor="w")
        tk.Label(brand, text=APP_SLOGAN, bg=c["primary_dark"], fg="#B9C1EA",
                 font=(FONT, 9)).pack(anchor="w", pady=(2, 0))

        tools = tk.Frame(header, bg=c["primary_dark"])
        tools.pack(side="right", pady=5)
        ttk.Button(tools, text="Atualizar dados", command=self.reload).pack(side="left", padx=(0, 8))
        ttk.Button(tools, text="Exportar para planilha", style="Accent.TButton",
                   command=self.export).pack(side="left")

        cards = tk.Frame(self.root, bg=c["bg"], padx=24, pady=18)
        cards.pack(fill="x")
        for key, title in METRICS:
            self._metric_card(cards, key, title)

        self.tabs = ttk.Notebook(self.root, padding=0)
        self.tabs.pack(fill="both", expand=True, padx=24, pady=(0, 22))
        for kind, spec in ENTITIES.items():
            self._build_table_tab(kind, spec)
        self._build_charts_tab()

    def _metric_card(self, parent, key, title):
        c = PALETTE
        card = tk.Frame(parent, bg=c["card"], highlightbackground=c["line"],
                        highlightthickness=1, padx=18, pady=13)
        card.pack(side="left", fill="both", expand=True, padx=(0, 12))
        tk.Label(card, text=title, bg=c["card"], fg=c["muted"],
                 font=(FONT, 8, "bold")).pack(anchor="w")
        label = tk.Label(card, text="0", bg=c["card"], fg=c["primary_dark"],
                         font=(FONT, 21, "bold"))
        label.pack(anchor="w", pady=(4, 0))
        self.metric_labels[key] = label

    def _toolbar(self, parent, heading):
        bar = ttk.Frame(parent, padding=(0, 0, 0, 12))
        bar.pack(fill="x")
        ttk.Label(bar, text=heading, font=(FONT, 14, "bold")).pack(side="left", padx=(0, 16))
        return bar

    def _build_table_tab(self, kind, spec):
        tab = ttk.Frame(self.tabs, padding=(16, 16))
        self.tabs.add(tab, text=spec["tab"])
        bar = self._toolbar(tab, spec["heading"])
        ttk.Button(bar, text=spec["new_label"], style="Accent.TButton",
                   command=lambda: self.open_form(kind)).pack(side="right", padx=(8, 0))
        ttk.Button(bar, text="Editar",
                   command=lambda: self.edit_selected(kind)).pack(side="right")
        ttk.Button(bar, text="Excluir", style="Danger.TButton",
                   command=lambda: self.delete_selected(kind)).pack(side="right", padx=8)
        self.trees[kind] = self._make_tree(tab, spec["columns"])

    def _build_charts_tab(self):
        tab = ttk.Frame(self.tabs, padding=(16, 16))
        self.tabs.add(tab, text="Gráficos")
        bar = self._toolbar(tab, "Visão analítica")
        ttk.Button(bar, text="Atualizar gráficos", style="Accent.TButton",
                   command=self.draw_charts).pack(side="right")
        self.chart_frame = ttk.Frame(tab)
        self.chart_frame.pack(fill="both", expand=True)

    @staticmethod
    def _make_tree(parent, columns):
        box = ttk.Frame(parent, padding=1)
        box.pack(fill="both", expand=True)
        keys = [col[0] for col in columns]
        tree = ttk.Treeview(box, columns=keys, show="headings", selectmode="browse")
        vbar = ttk.Scrollbar(box, orient="vertical", command=tree.yview)
        hbar = ttk.Scrollbar(box, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vbar.grid(row=0, column=1, sticky="ns")
        hbar.grid(row=1, column=0, sticky="ew")
        box.rowconfigure(0, weight=1)
        box.columnconfigure(0, weight=1)
        for key, label, width, _ in columns:
            tree.heading(key, text=label)
            tree.column(key, width=width, minwidth=45, anchor="w")
        tree.tag_configure("even", background=PALETTE["card"])
        tree.tag_configure("odd", background=PALETTE["stripe"])
        return tree

    # ---- dados -------------------------------------------------------

    def reload(self):
        try:
            self.products, self.movements = fetch_frames(PRODUCTS_SQL, MOVEMENTS_SQL)
        except (mysql.connector.Error, ValueError) as error:
            messagebox.showerror("Erro ao consultar o banco", str(error), parent=self.root)
            return
        self._fill("product", self.products)
        self._fill("movement", self.movements)
        self._update_metrics()

    def _fill(self, kind, frame):
        tree = self.trees[kind]
        tree.delete(*tree.get_children())
        columns = ENTITIES[kind]["columns"]
        for index, row in enumerate(frame.itertuples(index=False)):
            tree.insert("", "end", iid=str(row.id),
                        tags=("even" if index % 2 == 0 else "odd",),
                        values=tuple(col[3](row) for col in columns))

    def _update_metrics(self):
        units = int(self.products["estoque_atual"].sum()) if len(self.products) else 0
        numbers = {
            "products": len(self.products),
            "stock": units,
            "movements": len(self.movements),
        }
        for key, value in numbers.items():
            self.metric_labels[key].configure(text=thousands(value))

    # ---- ações sobre registros ---------------------------------------

    def _frame_for(self, kind):
        return self.products if kind == "product" else self.movements

    def _selected_id(self, kind, action):
        selection = self.trees[kind].selection()
        if not selection:
            messagebox.showinfo("Seleção necessária",
                                f"Selecione um registro para {action}.", parent=self.root)
            return None
        return int(selection[0])

    def edit_selected(self, kind):
        record_id = self._selected_id(kind, "editar")
        if record_id is None:
            return
        frame = self._frame_for(kind)
        self.open_form(kind, frame.loc[frame["id"] == record_id].iloc[0].to_dict())

    def delete_selected(self, kind):
        record_id = self._selected_id(kind, "excluir")
        if record_id is None:
            return
        noun = ENTITIES[kind]["noun"]
        if not messagebox.askyesno("Confirmar exclusão", f"Deseja excluir este {noun}?",
                                   parent=self.root):
            return
        try:
            with Transaction() as cursor:
                if kind == "product":
                    cursor.execute("DELETE FROM produtos WHERE id=%s", (record_id,))
                else:
                    cursor.execute("SELECT produto_id FROM movimentacoes WHERE id=%s FOR UPDATE",
                                   (record_id,))
                    found = cursor.fetchone()
                    if found is None:
                        raise ValueError("A movimentação selecionada não existe mais.")
                    if stock_balance(cursor, int(found[0]), record_id) < 0:
                        raise ValueError("A exclusão deixaria o estoque negativo.")
                    cursor.execute("DELETE FROM movimentacoes WHERE id=%s", (record_id,))
        except (mysql.connector.Error, ValueError) as error:
            messagebox.showerror(
                "Não foi possível excluir",
                "O produto pode ter movimentações vinculadas. "
                "Exclua ou ajuste os vínculos primeiro.\n\n" + str(error),
                parent=self.root,
            )
            return
        self.reload()

    # ---- formulário --------------------------------------------------

    def _form_fields(self, kind):
        if kind == "product":
            return [
                ("nome", "Nome", "entry"),
                ("categoria", "Categoria", "entry"),
                ("preco_unitario", "Preço unitário", "entry"),
                ("estoque_inicial", "Estoque inicial", "entry"),
            ]
        return [
            ("produto_id", "Produto", "product"),
            ("tipo", "Tipo", "type"),
            ("quantidade", "Quantidade", "entry"),
            ("data_movimento", "Data (AAAA-MM-DD)", "entry"),
            ("valor_unitario", "Valor unitário", "entry"),
            ("observacao", "Observação", "entry"),
        ]

    def open_form(self, kind, record=None):
        is_product = kind == "product"
        if not is_product and self.products.empty:
            messagebox.showinfo("Sem produtos",
                                "Cadastre um produto antes de registrar movimentações.",
                                parent=self.root)
            return
        window = tk.Toplevel(self.root)
        verb = "Editar" if record else "Cadastrar"
        window.title(f"{verb} {ENTITIES[kind]['noun']}")
        window.transient(self.root)
        window.grab_set()
        body = ttk.Frame(window, padding=16)
        body.pack(fill="both", expand=True)

        options = [(int(r.id), f"{int(r.id)} - {r.nome}")
                   for r in self.products.itertuples(index=False)]
        fields = self._form_fields(kind)
        inputs = {}
        for line, (key, label, widget_type) in enumerate(fields):
            ttk.Label(body, text=label).grid(row=line, column=0, sticky="w", pady=5)
            initial = record.get(key, "") if record else ""
            if widget_type == "product":
                widget = ttk.Combobox(body, state="readonly", width=34,
                                      values=[name for _, name in options])
                widget.set(next((name for pid, name in options if pid == initial), ""))
            elif widget_type == "type":
                widget = ttk.Combobox(body, state="readonly", width=34,
                                      values=("Entrada", "Saida"))
                widget.set(initial or "Entrada")
            else:
                if isinstance(initial, (date, datetime)):
                    initial = initial.isoformat()
                widget = ttk.Entry(body, width=37)
                widget.insert(0, str(initial))
            widget.grid(row=line, column=1, sticky="ew", pady=5, padx=(10, 0))
            inputs[key] = widget
        body.columnconfigure(1, weight=1)

        def submit():
            try:
                values = self._read_form(inputs)
                record_id = record.get("id") if record else None
                if is_product:
                    if not self._confirm_product(values, record, window):
                        return
                    self._save_product(values, record_id)
                else:
                    if not self._confirm_movement(values, record, window):
                        return
                    self._save_movement(values, record_id)
                window.destroy()
                self.reload()
            except (ValueError, mysql.connector.Error) as error:
                messagebox.showerror("Não foi possível salvar", str(error), parent=window)

        footer = ttk.Frame(body)
        footer.grid(row=len(fields), column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(footer, text="Cancelar", command=window.destroy).pack(side="right", padx=(6, 0))
        ttk.Button(footer, text="Salvar", command=submit).pack(side="right")
        window.bind("<Return>", lambda _event: submit())
        window.wait_window()

    def _read_form(self, inputs):
        values = {}
        for key, widget in inputs.items():
            if key == "produto_id":
                chosen = widget.get().split(" - ", 1)
                if not chosen or not chosen[0]:
                    raise ValueError("Selecione um produto.")
                values[key] = int(chosen[0])
            else:
                values[key] = widget.get().strip()
        return values

    def _confirm_product(self, values, record, window):
        values["preco_unitario"] = parse_money(values["preco_unitario"])
        values["estoque_inicial"] = parse_int(values["estoque_inicial"], "Estoque inicial", 0)
        if not values["nome"] or len(values["nome"]) > 100:
            raise ValueError("Informe o nome do produto (até 100 caracteres).")
        if not values["categoria"] or len(values["categoria"]) > 60:
            raise ValueError("Informe a categoria (até 60 caracteres).")
        action = "atualizar" if record else "cadastrar"
        return messagebox.askyesno(
            "Confirmar operação", f"Confirma {action} o produto '{values['nome']}'?",
            parent=window)

    def _confirm_movement(self, values, record, window):
        values["quantidade"] = parse_int(values["quantidade"], "Quantidade", 1)
        values["valor_unitario"] = parse_money(values["valor_unitario"])
        try:
            datetime.strptime(values["data_movimento"], "%Y-%m-%d")
        except ValueError as error:
            raise ValueError("A data deve estar no formato AAAA-MM-DD.") from error
        if len(values["observacao"]) > 255:
            raise ValueError("A observação deve ter até 255 caracteres.")
        name = self.products.loc[self.products["id"] == values["produto_id"], "nome"].iloc[0]
        action = "atualizar" if record else "registrar"
        return messagebox.askyesno(
            "Confirmar operação",
            f"Confirma {action} {values['tipo'].lower()} de "
            f"{values['quantidade']} unidade(s) de '{name}'?",
            parent=window)

    # ---- gravação ----------------------------------------------------

    @staticmethod
    def _save_product(values, product_id):
        with Transaction() as cursor:
            if product_id is None:
                cursor.execute(
                    """INSERT INTO produtos (nome, categoria, preco_unitario, estoque_inicial)
                       VALUES (%s, %s, %s, %s)""",
                    (values["nome"], values["categoria"],
                     values["preco_unitario"], values["estoque_inicial"]),
                )
                return
            cursor.execute(
                f"""SELECT COALESCE(SUM({STOCK_DELTA.format(t='')}), 0)
                    FROM movimentacoes WHERE produto_id = %s""",
                (product_id,),
            )
            balance = int(cursor.fetchone()[0])
            if values["estoque_inicial"] + balance < 0:
                raise ValueError("Esse estoque inicial deixaria o saldo atual negativo.")
            cursor.execute(
                """UPDATE produtos SET nome=%s, categoria=%s, preco_unitario=%s,
                   estoque_inicial=%s WHERE id=%s""",
                (values["nome"], values["categoria"], values["preco_unitario"],
                 values["estoque_inicial"], product_id),
            )

    @staticmethod
    def _save_movement(values, movement_id):
        row = (values["produto_id"], values["tipo"], values["quantidade"],
               values["data_movimento"], values["valor_unitario"], values["observacao"])
        with Transaction() as cursor:
            cursor.execute("SELECT id FROM produtos WHERE id=%s FOR UPDATE",
                           (values["produto_id"],))
            if cursor.fetchone() is None:
                raise ValueError("O produto selecionado não existe.")
            if movement_id is not None:
                cursor.execute("SELECT produto_id FROM movimentacoes WHERE id=%s FOR UPDATE",
                               (movement_id,))
                previous = cursor.fetchone()
                if previous is None:
                    raise ValueError("A movimentação selecionada não existe mais.")
                if stock_balance(cursor, int(previous[0]), movement_id) < 0:
                    raise ValueError("A alteração deixaria o estoque do produto anterior negativo.")
            available = stock_balance(cursor, values["produto_id"], movement_id)
            if values["tipo"] == "Saida" and values["quantidade"] > available:
                raise ValueError(f"Estoque insuficiente. Disponível: {available}.")
            if movement_id is None:
                cursor.execute(
                    """INSERT INTO movimentacoes
                       (produto_id, tipo, quantidade, data_movimento, valor_unitario, observacao)
                       VALUES (%s, %s, %s, %s, %s, %s)""",
                    row,
                )
            else:
                cursor.execute(
                    """UPDATE movimentacoes SET produto_id=%s, tipo=%s, quantidade=%s,
                       data_movimento=%s, valor_unitario=%s, observacao=%s WHERE id=%s""",
                    row + (movement_id,),
                )

    # ---- exportação e gráficos ---------------------------------------

    def export(self):
        if not messagebox.askyesno(
            "Confirmar exportação",
            "Deseja enviar os dados atuais para a planilha?\n\n"
            f"As abas {SHEET_DETAIL} e {SHEET_SUMMARY} serão atualizadas e formatadas.",
            parent=self.root,
        ):
            return
        try:
            url = export_reports(*build_reports())
        except (mysql.connector.Error, gspread.exceptions.GSpreadException,
                OSError, ValueError) as error:
            messagebox.showerror("Falha na exportação", str(error), parent=self.root)
            return
        messagebox.showinfo("Exportação concluída",
                            f"Relatório enviado para a planilha:\n{url}", parent=self.root)

    def draw_charts(self):
        c = PALETTE
        try:
            category, monthly = fetch_frames(CATEGORY_CHART_SQL, MONTHLY_CHART_SQL)
        except (mysql.connector.Error, ValueError) as error:
            messagebox.showerror("Falha ao gerar gráficos", str(error), parent=self.root)
            return

        figure = Figure(figsize=(10, 4.8), dpi=100)
        figure.patch.set_facecolor(c["card"])
        left, right = figure.subplots(1, 2)
        for axis in (left, right):
            axis.set_facecolor(c["card"])
            for side in ("top", "right"):
                axis.spines[side].set_visible(False)
            for side in ("left", "bottom"):
                axis.spines[side].set_color(c["line"])
            axis.tick_params(colors=c["muted"], labelsize=9)
            axis.grid(axis="y", color=c["line"], linewidth=0.8, alpha=0.8)
            axis.set_axisbelow(True)

        title_style = dict(loc="left", pad=14, fontweight="bold", color=c["text"])
        left.bar(category["categoria"], category["quantidade"],
                 color=c["primary"], width=0.58, zorder=3)
        left.set_title("Movimentações por categoria", **title_style)
        left.set_ylabel("Unidades movimentadas", color=c["muted"])
        left.tick_params(axis="x", rotation=25)

        positions = list(range(len(monthly)))
        balance = monthly["saldo"].tolist()
        right.plot(positions, balance, marker="o", markersize=6, linewidth=2.5,
                   color=c["accent"])
        right.fill_between(positions, balance, color=c["accent"], alpha=0.12)
        right.set_title("Saldo mensal de movimentações", **title_style)
        right.set_ylabel("Entradas - saídas", color=c["muted"])
        right.set_xticks(positions, monthly["periodo"].tolist())
        right.tick_params(axis="x", rotation=35)
        figure.tight_layout()

        if self.chart_canvas is not None:
            self.chart_canvas.get_tk_widget().destroy()
        self.chart_canvas = FigureCanvasTkAgg(figure, master=self.chart_frame)
        self.chart_canvas.draw()
        self.chart_canvas.get_tk_widget().pack(fill="both", expand=True)
        try:
            figure.savefig(Path(__file__).with_name("graficos_financeiro.png"))
        except OSError as error:
            messagebox.showerror("Falha ao gerar gráficos", str(error), parent=self.root)


def main():
    root = tk.Tk()
    FinanceApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
