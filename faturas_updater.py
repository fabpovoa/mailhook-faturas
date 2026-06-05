"""
Atualiza as abas "BB Altus Visa Luiz" e "BB Altus Visa Fátima" na planilha
mestre de faturas dos pais, localizada em:

  Drive → Cartões / <ano> / <mês> / Pais / faturas_Cartões_Luiz_<Mês>_ref_<MêsAnterior>_<AA>.xlsx

Usa a Google Sheets API para:
  1. Localizar o cabeçalho (linha com "Data", "Descrição", etc.) em cada aba.
  2. Limpar tudo abaixo do cabeçalho.
  3. Escrever as transações correspondentes.
"""
import logging
from datetime import datetime
from typing import Optional

from googleapiclient.discovery import build
from openpyxl.styles import PatternFill

from gmail import get_credentials

BRL_FORMAT = '"R$ "#,##0.00'

def _to_float(value) -> float:
    if isinstance(value, (int, float)):
        return value
    s = str(value).strip().replace("R$", "").replace(" ", "").replace("-", "")
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return value

log = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────────
# Mapeamento de titulares → abas na planilha mestre
# ─────────────────────────────────────────────────────────────
OWNER_MAP: dict[str, list[str]] = {
    "Fábio":  ["FABIO R POVOA",   "6688"],
    "Luiz":   ["LUIZ C M POVOA",  "0631"],
    "Fátima": ["MARIA F R POVOA", "2176"],
    "José":   ["JOSE RODRIGUES",  "3304"],
}

# Abas da planilha mestre que recebem os dados (só Luiz e Fátima por enquanto)
TAB_MAP: dict[str, str] = {
    "Luiz":   "BB Altus Visa Luiz",
    "Fátima": "BB Altus Visa Fátima",
}

# Marcador de parada no extrato
STOP_MARKERS = [
    "PARCELAMENTOS PRÓXIMA FATURA",
    "PARCELAMENTOS PROXIMA FATURA",
    "PARCELAMENTOS - PRÓXIMA FATURA",
    "PARCELAMENTOS - PROXIMA FATURA",
]

# Colunas a gravar na aba de destino (mesma ordem do cabeçalho)
OUTPUT_COLUMNS = ["Data", "Descrição", "País", "Valor (R$)", "Categoria", "Tipo"]


# ─────────────────────────────────────────────────────────────
# Serviços Google
# ─────────────────────────────────────────────────────────────

def _creds():
    return get_credentials()

def get_sheets_service():
    return build("sheets", "v4", credentials=_creds())

def get_drive_service():
    from drive import get_drive_service as _ds
    return _ds()


# ─────────────────────────────────────────────────────────────
# Agrupamento por titular
# ─────────────────────────────────────────────────────────────

def _match_owner_tab(titular: str) -> Optional[str]:
    t = titular.upper()
    for tab_name, tokens in OWNER_MAP.items():
        if any(token in t for token in tokens):
            return tab_name
    return None


def _is_stop_marker(desc: str) -> bool:
    d = desc.upper().strip()
    return any(m in d for m in STOP_MARKERS)


def group_by_owner(transactions: list[dict]) -> dict[str, list[dict]]:
    """
    Agrupa transações pelo campo Titular (preenchido pelo extrator).
    Para em "Parcelamentos Próxima Fatura".
    """
    grouped: dict[str, list[dict]] = {tab: [] for tab in OWNER_MAP}

    for tx in transactions:
        desc = str(tx.get("Descrição", "")).strip()
        desc_upper = desc.upper()

        if _is_stop_marker(desc):
            log.info("🛑 Marcador de parada: '%s'. Ignorando restante.", desc)
            break

        if any(k in desc_upper for k in ["SUBTOTAL", "SUB-TOTAL", "SALDO ANTERIOR", "TOTAL PARCELADO"]):
            continue

        titular = str(tx.get("Titular", ""))
        tab = _match_owner_tab(titular)
        if not tab:
            continue

        tx = dict(tx)
        tx["Titular"] = tab
        grouped[tab].append(tx)

    totals = {k: len(v) for k, v in grouped.items()}
    log.info("Separação concluída → %s", totals)
    return grouped


# ─────────────────────────────────────────────────────────────
# Localização da planilha mestre no Drive
# ─────────────────────────────────────────────────────────────

def _build_filename(ref_date: datetime) -> str:
    """
    Gera o nome esperado da planilha mestre.
    Exemplo: faturas_Cartões_Luiz_Jun_ref_May_26.xlsx
    Usa abreviações em inglês (strftime %b) pois é o padrão do arquivo no Drive.
    """
    from datetime import date
    month_en = ref_date.strftime("%b")       # Jun
    prev_num = ref_date.month - 1 if ref_date.month > 1 else 12
    prev_year = ref_date.year if ref_date.month > 1 else ref_date.year - 1
    prev_en  = date(prev_year, prev_num, 1).strftime("%b")  # May
    year_s   = ref_date.strftime("%y")       # 26
    return f"faturas_Cartões_Luiz_{month_en}_ref_{prev_en}_{year_s}.xlsx"


def _get_month_folder_id(drive_service, ref_date: datetime) -> Optional[str]:
    """Navega Cartões → ano → mês e devolve o ID da pasta do mês."""
    from drive import (
        CARTOES_ROOT_ID,
        find_year_folder,
        find_month_folder,
        list_child_folders,
    )

    log.info("🔍 Navegando para pasta do mês (ano=%d mês=%d)...", ref_date.year, ref_date.month)

    year_id = find_year_folder(drive_service, ref_date.year, CARTOES_ROOT_ID)
    if not year_id:
        log.error("❌ Pasta do ano %d não encontrada em Cartões.", ref_date.year)
        return None
    log.info("   ✅ Ano encontrado: %s", year_id)

    month_id = find_month_folder(drive_service, ref_date.month, year_id)
    if not month_id:
        log.error("❌ Pasta do mês %d não encontrada. Subpastas:", ref_date.month)
        for f in list_child_folders(drive_service, year_id):
            log.error("     • %s (id=%s)", f["name"], f["id"])
        return None
    log.info("   ✅ Mês encontrado: %s", month_id)
    return month_id


def _find_spreadsheet_id(drive_service, folder_id: str, filename: str) -> Optional[str]:
    """
    Procura o arquivo na pasta pelo nome (com ou sem .xlsx, Google Sheets ou xlsx).
    Devolve o ID do arquivo/spreadsheet.
    """
    base = filename.replace(".xlsx", "")
    candidates = [filename, base]

    for name in candidates:
        from drive import _escape_drive_query, _drive_list_kwargs
        q = (
            f"'{folder_id}' in parents and name = '{_escape_drive_query(name)}'"
            f" and trashed = false"
        )
        results = drive_service.files().list(
            q=q,
            fields="files(id, name, mimeType)",
            **_drive_list_kwargs(),
        ).execute()
        files = results.get("files", [])
        if files:
            f = files[0]
            log.info("📄 Planilha mestre encontrada: %s (ID: %s)", f["name"], f["id"])
            return f["id"]

    log.warning("Planilha mestre não encontrada na pasta Pais com nome: %s", filename)
    return None


# ─────────────────────────────────────────────────────────────
# Atualização via openpyxl (arquivo é .xlsx, não Google Sheet nativo)
# ─────────────────────────────────────────────────────────────

def _find_header_info(ws) -> tuple[int, dict[str, int]]:
    """
    Localiza a linha de cabeçalho (contém "Data" e "Descrição") e devolve:
      (row_number_1based, {col_name: col_number_1based})
    Retorna (-1, {}) se não encontrada.
    """
    for row in ws.iter_rows():
        cells = {str(c.value or "").strip(): c.column for c in row if c.value}
        keys_lower = {k.lower(): k for k in cells}
        if any("data" in k for k in keys_lower) and any("descri" in k for k in keys_lower):
            col_map = {orig: cells[orig] for orig in cells}
            return row[0].row, col_map
    return -1, {}


def _update_xlsx_tab(wb, tab_name: str, transactions: list[dict]) -> bool:
    """
    Localiza o cabeçalho na aba, detecta o mapeamento de colunas,
    limpa os dados abaixo e grava as transações respeitando:
      - Coluna A sempre vazia
      - Colunas determinadas pelo cabeçalho real do arquivo
    """
    if tab_name not in wb.sheetnames:
        log.warning("Aba '%s' não encontrada. Disponíveis: %s", tab_name, wb.sheetnames)
        return False

    ws = wb[tab_name]
    header_row, col_map = _find_header_info(ws)
    if header_row < 0:
        log.warning("Cabeçalho não encontrado na aba '%s'.", tab_name)
        return False

    log.info("Aba '%s': cabeçalho na linha %d, colunas: %s", tab_name, header_row, col_map)

    data_start = header_row + 1

    # Limpa linhas de dados existentes
    if ws.max_row >= data_start:
        ws.delete_rows(data_start, ws.max_row - data_start + 1)

    # Grava cada transação nas colunas corretas (col A sempre vazia)
    for tx in transactions:
        row_num = ws.max_row + 1
        ws.cell(row=row_num, column=1, value=None)  # coluna A vazia
        for col_name, col_idx in col_map.items():
            value = tx.get(col_name, "")
            cell = ws.cell(row=row_num, column=col_idx)
            if col_name == "Valor (R$)" and value:
                cell.value = _to_float(value)
                cell.number_format = BRL_FORMAT
            else:
                cell.value = value

    log.info("Aba '%s': %d transações gravadas (linha %d em diante).", tab_name, len(transactions), data_start)
    return True


# ─────────────────────────────────────────────────────────────
# Ponto de entrada principal
# ─────────────────────────────────────────────────────────────

def update_faturas_luiz(
    transactions: list[dict],
    ref_date: Optional[datetime] = None,
) -> Optional[str]:
    """
    Agrupa as transações por titular e atualiza as abas "BB Altus Visa Luiz"
    e "BB Altus Visa Fátima" na planilha mestre localizada em:
    Cartões / <ano> / <mês> / Pais /

    Retorna o link web da planilha ou None em caso de falha.
    """
    if ref_date is None:
        ref_date = datetime.now()

    import io
    import openpyxl
    from googleapiclient.http import MediaIoBaseUpload

    drive_service = get_drive_service()

    # 1. Navega até a pasta do mês
    month_folder_id = _get_month_folder_id(drive_service, ref_date)
    if not month_folder_id:
        log.error("Pasta do mês não encontrada. Atualização cancelada.")
        return None

    # 2. Localiza a subpasta Pais
    from drive import find_folder
    pais_folder_id = None
    for name in ("Pais", "País"):
        pais_folder_id = find_folder(drive_service, name, month_folder_id)
        if pais_folder_id:
            log.info("📁 Pasta '%s' encontrada: %s", name, pais_folder_id)
            break
    if not pais_folder_id:
        log.error("Subpasta 'Pais' não encontrada. Atualização cancelada.")
        return None

    # 3. Localiza o arquivo xlsx
    filename = _build_filename(ref_date)
    file_id  = _find_spreadsheet_id(drive_service, pais_folder_id, filename)
    if not file_id:
        log.error("Planilha '%s' não encontrada. Atualização cancelada.", filename)
        return None

    # 4. Baixa o arquivo como xlsx
    log.info("⬇️  Baixando planilha mestre (ID: %s)...", file_id)
    content = drive_service.files().get_media(
        fileId=file_id, supportsAllDrives=True
    ).execute()
    wb = openpyxl.load_workbook(io.BytesIO(content))

    # 5. Agrupa transações e atualiza as abas
    grouped = group_by_owner(transactions)
    for owner_tab, target_tab in TAB_MAP.items():
        txs = grouped.get(owner_tab, [])
        log.info("Atualizando aba '%s' (%d transações)...", target_tab, len(txs))
        _update_xlsx_tab(wb, target_tab, txs)

    # 6. Salva e reenvia para o Drive (sobrescreve o arquivo existente)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    drive_service.files().update(
        fileId=file_id,
        media_body=MediaIoBaseUpload(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            resumable=True,
        ),
        fields="id, webViewLink",
        supportsAllDrives=True,
    ).execute()

    link = f"https://docs.google.com/spreadsheets/d/{file_id}/edit"
    log.info("✨ Planilha mestre atualizada: %s", link)
    return link


# ─────────────────────────────────────────────────────────────
# Atualização da aba BTG na planilha dos pais
# ─────────────────────────────────────────────────────────────

# Mapeamento de aba BTG → número do cartão final
BTG_CARD_MAP = {
    "Luiz":       "0627",
    "Fátima 7387": "7387",
    "Fátima 5310": "5310",
}

# Colunas da aba BTG (cabeçalho detectado dinamicamente, mas esta é a ordem esperada)
BTG_COLUMNS = ["Data", "Descrição", "Parcela", "Valor (R$)", "Tipo"]

BTG_DATA_START_ROW = 11  # linha onde os dados começam (conforme planilha)


def _find_btg_tab(wb) -> Optional[str]:
    """Encontra a aba BTG no workbook (busca case-insensitive)."""
    for name in wb.sheetnames:
        if "btg" in name.lower():
            return name
    return None


def _update_btg_xlsx_tab(wb, tab_name: str, grouped: dict[str, list[dict]], pdf_link: str = "") -> bool:
    """
    Atualiza a aba BTG da planilha dos pais a partir da linha BTG_DATA_START_ROW.
    Escreve Luiz e Fátima (7387 + 5310) — sem Hiero.
    Colunas: A=vazia, B=Data, C=Descrição, D=Parcela, E=Valor(float), F=Tipo, G=vazia, H=Final Cartão
    """
    if tab_name not in wb.sheetnames:
        log.warning("Aba BTG '%s' não encontrada. Disponíveis: %s", tab_name, wb.sheetnames)
        return False

    ws = wb[tab_name]

    # G2 — link do PDF com fundo branco
    if pdf_link:
        cell_g2 = ws.cell(row=2, column=7, value=pdf_link)
        cell_g2.fill = PatternFill("solid", fgColor="FFFFFF")
        log.info("Aba BTG: link PDF gravado em G2.")

    # Limpa a partir da linha de dados
    if ws.max_row >= BTG_DATA_START_ROW:
        ws.delete_rows(BTG_DATA_START_ROW, ws.max_row - BTG_DATA_START_ROW + 1)

    row = BTG_DATA_START_ROW
    for owner_tab, card_final in BTG_CARD_MAP.items():
        txs = grouped.get(owner_tab, [])
        for tx in txs:
            ws.cell(row=row, column=1, value=None)                        # A vazia
            ws.cell(row=row, column=2, value=tx.get("Data", ""))          # B
            ws.cell(row=row, column=3, value=tx.get("Descrição", ""))     # C
            ws.cell(row=row, column=4, value=tx.get("Parcela", ""))       # D
            valor_cell = ws.cell(row=row, column=5, value=_to_float(tx.get("Valor (R$)", "")))  # E
            valor_cell.number_format = BRL_FORMAT
            ws.cell(row=row, column=6, value=tx.get("Tipo", "Compra"))    # F
            ws.cell(row=row, column=7, value=None)                         # G vazia
            ws.cell(row=row, column=8, value=card_final)                   # H
            row += 1

    total = row - BTG_DATA_START_ROW
    log.info("Aba BTG '%s': %d transações gravadas (linha %d em diante).", tab_name, total, BTG_DATA_START_ROW)
    return True


def update_btg_pais(
    grouped: dict[str, list[dict]],
    ref_date: Optional[datetime] = None,
    pdf_link: str = "",
) -> Optional[str]:
    """
    Atualiza a aba BTG na planilha dos pais com transações de Luiz e Fátima.
    Navega Cartões / <ano> / <mês> / Pais / <filename>.
    """
    import io
    import openpyxl
    from googleapiclient.http import MediaIoBaseUpload

    if ref_date is None:
        ref_date = datetime.now()

    drive_service = get_drive_service()

    # Navega até a pasta Pais
    month_folder_id = _get_month_folder_id(drive_service, ref_date)
    if not month_folder_id:
        log.error("Pasta do mês não encontrada.")
        return None

    from drive import find_folder
    pais_folder_id = None
    for name in ("Pais", "País"):
        pais_folder_id = find_folder(drive_service, name, month_folder_id)
        if pais_folder_id:
            break
    if not pais_folder_id:
        log.error("Subpasta 'Pais' não encontrada.")
        return None

    filename = _build_filename(ref_date)
    file_id  = _find_spreadsheet_id(drive_service, pais_folder_id, filename)
    if not file_id:
        log.error("Planilha '%s' não encontrada.", filename)
        return None

    # Baixa e atualiza
    content = drive_service.files().get_media(fileId=file_id, supportsAllDrives=True).execute()
    wb = openpyxl.load_workbook(io.BytesIO(content))

    btg_tab = _find_btg_tab(wb)
    if not btg_tab:
        log.error("Aba BTG não encontrada. Abas: %s", wb.sheetnames)
        return None

    _update_btg_xlsx_tab(wb, btg_tab, grouped, pdf_link=pdf_link)

    # Re-upload
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    drive_service.files().update(
        fileId=file_id,
        media_body=MediaIoBaseUpload(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            resumable=True,
        ),
        fields="id",
        supportsAllDrives=True,
    ).execute()

    link = f"https://docs.google.com/spreadsheets/d/{file_id}/edit"
    log.info("✨ Aba BTG atualizada: %s", link)
    return link
