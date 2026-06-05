"""
Cria o xls dos pais do mês seguinte a partir do atual.

Lógica da aba "Gastos cartões Fábio":
  - (recorrente) na Descrição  → copia, coluna G = "pending" (fundo amarelo)
  - X/Y na Descrição, X < Y   → copia, incrementa parcela para (X+1)/Y
  - X/Y na Descrição, X = Y   → última parcela, não copia
  - Outros (one-off)           → não copia

Demais ajustes:
  - Nome do arquivo atualizado (ex: faturas_Cartões_Luiz_Jul_ref_Jun_26.xlsx)
  - Aba "Total": célula B2 atualizada (ex: FATURAS Jul ref Jun 26)
  - Abas BB Altus Visa Luiz, BB Altus Visa Fátima, BTG: limpas (só cabeçalho)
"""
import io
import re
import logging
from datetime import datetime
from typing import Optional

import openpyxl
from openpyxl.styles import PatternFill

log = logging.getLogger(__name__)

YELLOW_FILL = PatternFill("solid", fgColor="FFFF00")
PARCELA_RE  = re.compile(r"(\d+)/(\d+)")
RECORRENTE  = "(recorrente)"

# Cor padrão de input (SmartMoney Ventures)
INPUT_FONT_COLOR = "342EFF"

from openpyxl.styles import Font

# Abas de cartões a limpar (mantém só o cabeçalho)
CARD_TABS = ["BB Altus Visa Luiz", "BB Altus Visa Fátima", "BTG"]

# Nome da aba de gastos do Fábio
FABIO_TAB = "Gastos cartões Fábio"

# Nome da aba de totais
TOTAL_TAB = "Total"


def _next_month_date(ref_date: datetime) -> datetime:
    if ref_date.month == 12:
        return datetime(ref_date.year + 1, 1, 1)
    return datetime(ref_date.year, ref_date.month + 1, 1)


def _month_abbr_pt(month: int) -> str:
    months = {1:"Jan",2:"Fev",3:"Mar",4:"Abr",5:"Mai",6:"Jun",
              7:"Jul",8:"Ago",9:"Set",10:"Out",11:"Nov",12:"Dez"}
    return months[month]


def _build_next_filename(ref_date: datetime) -> str:
    """
    Ex: ref_date = Jun/2026 → faturas_Cartões_Luiz_Jul_ref_Jun_26.xlsx
    """
    next_d   = _next_month_date(ref_date)
    next_m   = _month_abbr_pt(next_d.month)   # Jul
    ref_m    = _month_abbr_pt(ref_date.month)  # Jun
    year_s   = str(next_d.year)[-2:]           # 26
    return f"faturas_Cartões_Luiz_{next_m}_ref_{ref_m}_{year_s}.xlsx"


def _process_fabio_tab(ws, next_date: datetime):
    """
    Processa a aba 'Gastos cartões Fábio':
    - Detecta a linha de cabeçalho
    - Filtra e transforma cada linha de dado
    - Recria as linhas de dado abaixo do cabeçalho
    """
    # 1. Encontra linha de cabeçalho (tem "Data" e "Descrição" ou similar)
    header_row = None
    for row in ws.iter_rows():
        vals = [str(c.value or "").lower() for c in row]
        if "data" in vals and any("descri" in v for v in vals):
            header_row = row[0].row
            break

    if header_row is None:
        log.warning("Cabeçalho não encontrado em '%s'", FABIO_TAB)
        return

    # 2. Descobre qual coluna é Descrição e qual é Link (G)
    header_cells = list(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=False))[0]
    col_desc = col_link = None
    for cell in header_cells:
        v = str(cell.value or "").lower()
        if "descri" in v:
            col_desc = cell.column
        if v == "link" and col_link is None:
            col_link = cell.column

    if not col_desc:
        log.warning("Coluna Descrição não encontrada")
        return

    data_start = header_row + 1

    # 3. Coleta todas as linhas de dado existentes
    existing_rows = []
    for row in ws.iter_rows(min_row=data_start, values_only=False):
        if not any(c.value for c in row):
            break  # linha vazia = fim dos dados
        existing_rows.append(row)

    if not existing_rows:
        return

    # 4. Determina total de colunas a preservar
    max_col = max(c.column for row in existing_rows for c in row if c.value is not None)

    # 5. Apaga linhas de dado existentes
    ws.delete_rows(data_start, len(existing_rows) + 5)

    blue_font = Font(color=INPUT_FONT_COLOR)

    def write_row(row_cells, override_desc=None, pending_link=False):
        """
        Copia uma linha com fonte azul.
        pending_link=True  → coluna Link = "pending" com fundo amarelo (recorrentes)
        pending_link=False → coluna Link mantém valor original, sem fundo amarelo (parcelados)
        Ignora células completamente vazias sem valor (evita recriar colunas vazias).
        """
        for cell in row_cells:
            if cell.value is None:
                continue  # não recria células vazias
            value = override_desc if (override_desc and cell.column == col_desc) else cell.value
            new_cell = ws.cell(row=insert_row, column=cell.column, value=value)
            new_cell.font = blue_font
            if cell.column == col_link and pending_link:
                new_cell.value = "pending"
                new_cell.fill = YELLOW_FILL

    # 6. Processa cada linha e reinsere as qualificadas
    insert_row = data_start
    for row_cells in existing_rows:
        desc_cell = next((c for c in row_cells if c.column == col_desc), None)
        desc = str(desc_cell.value or "") if desc_cell else ""

        is_recorrente = RECORRENTE.lower() in desc.lower()
        parcela_m = PARCELA_RE.search(desc)

        if is_recorrente:
            # pending + fundo amarelo na coluna Link
            write_row(row_cells, pending_link=True)
            insert_row += 1

        elif parcela_m:
            current = int(parcela_m.group(1))
            total   = int(parcela_m.group(2))
            if current < total:
                new_desc = PARCELA_RE.sub(f"{current+1}/{total}", desc, count=1)
                # mantém link original, sem amarelo
                write_row(row_cells, override_desc=new_desc, pending_link=False)
                insert_row += 1
            # else: última parcela → não copia

        # else: one-off → não copia

    log.info("'%s': %d linhas → %d linhas no próximo mês.",
             FABIO_TAB, len(existing_rows), insert_row - data_start)


def _clear_card_tab(wb, tab_name: str):
    """Limpa dados de uma aba de cartão mantendo o cabeçalho."""
    if tab_name not in wb.sheetnames:
        return
    ws = wb[tab_name]
    # Encontra linha de cabeçalho (Data + Descrição)
    header_row = None
    for row in ws.iter_rows():
        vals = [str(c.value or "").lower() for c in row]
        if "data" in vals and any("descri" in v for v in vals):
            header_row = row[0].row
            break
    if header_row and ws.max_row > header_row:
        ws.delete_rows(header_row + 1, ws.max_row - header_row)
    log.info("Aba '%s' limpa (só cabeçalho).", tab_name)


def create_next_month_xls(xlsx_bytes: bytes, ref_date: datetime) -> tuple[bytes, str]:
    """
    Recebe o xls atual dos pais e retorna (novo_xls_bytes, novo_filename).

    ref_date: data de referência do mês atual (ex: datetime(2026, 6, 1))
    """
    next_date = _next_month_date(ref_date)
    filename  = _build_next_filename(ref_date)

    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))

    # 1. Atualiza célula B2 da aba Total
    if TOTAL_TAB in wb.sheetnames:
        ws_total = wb[TOTAL_TAB]
        next_m  = _month_abbr_pt(next_date.month)
        ref_m   = _month_abbr_pt(ref_date.month)
        year_s  = str(next_date.year)[-2:]
        ws_total["B2"] = f"FATURAS {next_m} ref {ref_m} {year_s}"
        log.info("Aba '%s': B2 = '%s'", TOTAL_TAB, ws_total["B2"].value)

    # 2. Processa aba Gastos cartões Fábio
    if FABIO_TAB in wb.sheetnames:
        _process_fabio_tab(wb[FABIO_TAB], next_date)

    # 3. Limpa abas de cartões
    for tab in CARD_TABS:
        _clear_card_tab(wb, tab)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), filename
