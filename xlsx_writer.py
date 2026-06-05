"""
Cria o arquivo Excel (.xlsx) com as transações extraídas do PDF.

Formato:
  - Uma aba por titular (Fábio, Luiz, Fátima, José).
  - Dentro de cada aba:
      Linha 1: "Lançamentos  |  Fatura BB – <título>"
      Linha 2: cabeçalhos das colunas (coluna A vazia)
      Linhas 3+: dados
"""
import io
from datetime import datetime
from typing import Optional

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

COLUMNS = ["Titular", "Data", "Descrição", "País", "Valor (R$)", "Categoria", "Tipo"]

BRL_FORMAT = '"R$ "#,##0.00'  # formato R$ para células numéricas


def _to_float(value) -> float:
    """Converte string no formato brasileiro (1.234,56) para float."""
    if isinstance(value, (int, float)):
        return value
    s = str(value).strip().replace("R$", "").replace(" ", "").replace("-", "")
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return value

# Cores
HEADER_BG = "1F497D"   # azul escuro BB
HEADER_FG = "FFFFFF"   # branco
TITLE_BG  = "DCE6F1"   # azul claro
ALT_ROW   = "EBF1DE"   # verde claro linhas alternadas


def _thin_border():
    thin = Side(style="thin", color="BFBFBF")
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def _write_sheet(ws, transactions: list[dict], fatura_title: str) -> None:
    """Preenche uma aba com cabeçalho + dados."""
    title_fill = PatternFill("solid", fgColor=TITLE_BG)
    title_font = Font(bold=True, size=12)
    header_fill = PatternFill("solid", fgColor=HEADER_BG)
    header_font = Font(bold=True, color=HEADER_FG, size=10)
    alt_fill    = PatternFill("solid", fgColor=ALT_ROW)
    data_font   = Font(size=10)
    border      = _thin_border()

    # Linha 1: título
    ws.cell(row=1, column=1, value="Lançamentos")
    ws.cell(row=1, column=2, value=fatura_title)
    ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=8)
    for col in range(1, 9):
        cell = ws.cell(row=1, column=col)
        cell.fill = title_fill
        cell.font = title_font
        cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 22

    # Linha 2: cabeçalhos
    ws.cell(row=2, column=1, value="")
    for i, col_name in enumerate(COLUMNS, start=2):
        cell = ws.cell(row=2, column=i, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[2].height = 18

    # Dados
    for row_idx, tx in enumerate(transactions, start=3):
        ws.cell(row=row_idx, column=1, value="")
        for col_idx, col_name in enumerate(COLUMNS, start=2):
            value = tx.get(col_name, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = data_font
            cell.border = border
            cell.alignment = Alignment(vertical="center")
            if row_idx % 2 == 0:
                cell.fill = alt_fill
            if col_name == "Valor (R$)":
                cell.value = _to_float(value)
                cell.number_format = BRL_FORMAT
                cell.alignment = Alignment(horizontal="right", vertical="center")

    # Largura das colunas
    col_widths = {"A": 3, "B": 18, "C": 8, "D": 44, "E": 5, "F": 16, "G": 22, "H": 20}
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    ws.freeze_panes = "B3"


def build_xlsx(
    transactions_by_owner: dict[str, list[dict]],
    fatura_title: str = "Fatura BB",
    ref_date: Optional[datetime] = None,
) -> bytes:
    """
    Recebe transações agrupadas por titular e retorna xlsx com uma aba por titular.

    transactions_by_owner: {tab_name: [transações]}
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove aba padrão vazia

    for tab_name, txs in transactions_by_owner.items():
        ws = wb.create_sheet(title=tab_name)
        _write_sheet(ws, txs, fatura_title)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Compatibilidade retroativa ──────────────────────────────
# Se algum código ainda chamar build_xlsx com lista flat, agrupa tudo
# em uma aba "Lançamentos".
def build_xlsx_flat(
    transactions: list[dict],
    fatura_title: str = "Fatura BB",
    ref_date: Optional[datetime] = None,
) -> bytes:
    """Versão legada: lista plana → aba única 'Lançamentos'."""
    return build_xlsx({"Lançamentos": transactions}, fatura_title, ref_date)
