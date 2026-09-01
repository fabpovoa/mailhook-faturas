"""
Atualiza a planilha "Boletos" (aba "Acomp mensal May/26 em diante", tabela
"Gastos e boletos a pagar") com o lançamento de uma conta processada por um
handler (ex.: claro_handler.py).

A tabela não tem um fim fixo — cresce a cada mês. Este módulo localiza a
última linha preenchida da coluna "Payee" a partir do início dos dados
(HEADER_ROW + 1) e decide entre atualizar uma linha já existente (mesmo Mês +
Payee — idempotência) ou anexar uma linha nova logo depois da última.
"""
import logging
from datetime import datetime
from typing import Optional

from googleapiclient.discovery import build

from gmail import get_credentials

log = logging.getLogger(__name__)

BOLETOS_SHEET_ID = "1G3S57VJmz-3EQIwK5Lo_iuJhVaI7X2h1EtPvPBpLxZo"
BOLETOS_TAB = "Acomp mensal May/26 em diante"

# Cabeçalho da tabela "Gastos e boletos a pagar" nessa aba:
# B=Mês C=Dia D=Payee E=Valor F=Link pdf G=Conta débito H=Forma de pagto
# I=Obs J=Link pagamento K=Válido até
HEADER_ROW = 4
DATA_START_ROW = HEADER_ROW + 1

MONTHS_PT_ABBR = {
    1: "jan", 2: "fev", 3: "mar", 4: "abr", 5: "mai", 6: "jun",
    7: "jul", 8: "ago", 9: "set", 10: "out", 11: "nov", 12: "dez",
}


def _mes_str(ref_date: datetime) -> str:
    return f"{MONTHS_PT_ABBR[ref_date.month]}./{str(ref_date.year)[-2:]}"


def _sheets_service():
    return build("sheets", "v4", credentials=get_credentials())


def upsert_boleto_row(
    *,
    ref_date: datetime,
    dia: int,
    payee: str,
    valor: float,
    link_pdf: str,
    conta_debito: str,
    forma_pagamento: str,
    link_pagamento: str = "pending",
) -> str:
    """
    Cria ou atualiza (idempotente por Mês+Payee) a linha do boleto na tabela
    "Gastos e boletos a pagar". Retorna a referência A1 da linha afetada.
    """
    service = _sheets_service()
    mes = _mes_str(ref_date)

    # Lê as colunas B (Mês) e D (Payee) de toda a tabela existente.
    result = service.spreadsheets().values().get(
        spreadsheetId=BOLETOS_SHEET_ID,
        range=f"'{BOLETOS_TAB}'!B{DATA_START_ROW}:D",
    ).execute()
    rows = result.get("values", [])

    target_row: Optional[int] = None
    last_filled_row = DATA_START_ROW - 1
    for i, row in enumerate(rows):
        row_num = DATA_START_ROW + i
        row_mes = row[0] if len(row) > 0 else ""
        row_payee = row[2] if len(row) > 2 else ""
        if row_mes or row_payee:
            last_filled_row = row_num
        if row_mes == mes and row_payee == payee:
            target_row = row_num
            break

    if target_row is None:
        target_row = last_filled_row + 1
        log.info("Boletos: nenhuma linha existente pra %s/%s — anexando na linha %d.", mes, payee, target_row)
    else:
        log.info("Boletos: linha existente %d pra %s/%s — atualizando.", target_row, mes, payee)

    values = [[
        mes, dia, payee, valor, link_pdf, conta_debito, forma_pagamento, "", link_pagamento, "",
    ]]
    service.spreadsheets().values().update(
        spreadsheetId=BOLETOS_SHEET_ID,
        range=f"'{BOLETOS_TAB}'!B{target_row}:K{target_row}",
        valueInputOption="USER_ENTERED",
        body={"values": values},
    ).execute()

    return f"B{target_row}:K{target_row}"
