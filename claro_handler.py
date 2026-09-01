"""
Handler — Fatura Claro Net Fibra, conta "Pais" (faturada no CNPJ do SMV1 —
Smartmoney Ventures Serviços Administrativos, 30.243.467/0001-23).

Fluxo (chamado a partir de main.py quando chega e-mail de faturadigital@minhaclaro.com.br,
direto ou encaminhado):
  1. Desbloqueia o PDF (senha = 5 primeiros dígitos do CNPJ, CLARO_CNPJ_PREFIX).
  2. Extrai vencimento / valor / forma de pagamento do texto do PDF.
  3. Salva o PDF em Contas/<ano>/<mês>/Pais/Claro Net Fibra/ (pasta assumida já
     criada pelo automation "Folders mensais"; get_or_create_folder é
     idempotente, então funciona também se a pasta ainda não existir).
  4. Cria o lançamento em Contas a Pagar no app money via webhook HTTP
     (o template recorrente "Claro Net Fibra" já precisa existir lá, vinculado
     ao Titular SMV1 — ver prisma/schema.prisma no repo money).
  5. Atualiza a linha do mês na planilha Boletos (boletos_updater.py).

Non-goals: o rollover mensal da planilha de cartões do pai (aba "Gastos
cartões Fábio", incluindo a linha "(recorrente) Claro Fibra") já é coberto por
next_month_xls.py — não duplicado aqui.
"""
import io
import os
import re
import logging
from datetime import datetime
from typing import Optional

import httpx
import pdfplumber
from googleapiclient.http import MediaIoBaseUpload

from drive import (
    get_drive_service,
    get_or_resolve_year_folder,
    get_or_resolve_month_folder,
    get_or_create_folder,
)
from pdf_handler import unlock_pdf
from boletos_updater import upsert_boleto_row

log = logging.getLogger(__name__)

CONTAS_ROOT_ID = "1CsrkyGFQGwtAyhykp6TrJhMvXZGokt5L"  # Personal stuff/Finanças/Contas

CLARO_SENDER = "faturadigital@minhaclaro.com.br"
# 5 primeiros dígitos do CNPJ do SMV1 (30.243.467/0001-23) — senha do PDF da Claro.
CLARO_CNPJ_PREFIX = os.getenv("CLARO_CNPJ_PREFIX", "30243")

MONEY_WEBHOOK_URL = os.getenv(
    "MONEY_WEBHOOK_URL", "https://money.smartmoney.ventures/api/webhooks/contas-a-pagar"
)
MONEY_WEBHOOK_SECRET = os.getenv("MONEY_WEBHOOK_SECRET", "")

CLARO_CONTA_DEBITO = os.getenv("CLARO_CONTA_DEBITO", "99 Fábio")
CLARO_PAYEE_APP = "Claro Net Fibra"       # payee no template do app money
CLARO_PAYEE_SHEET = "Claro NET Virtua"    # payee histórico na planilha Boletos

VENCIMENTO_RE = re.compile(r"Vencimento\s*[\r\n]+\s*(\d{2}/\d{2}/\d{4})")
VALOR_RE = re.compile(r"Valor\s*[\r\n]+\s*([\d.]+,\d{2})")
FORMA_PGTO_RE = re.compile(r"Forma de Pagamento\s*[\r\n]+\s*([A-ZÀ-Ú ]+)")


def is_fatura_claro(message: dict) -> bool:
    """E-mail direto da Claro, ou encaminhado citando o remetente/assunto original."""
    headers = {h["name"].lower(): h["value"] for h in message.get("payload", {}).get("headers", [])}
    from_email = headers.get("from", "").lower()
    subject = headers.get("subject", "").lower()
    if CLARO_SENDER in from_email:
        return True
    return "claro" in subject and ("fatura" in subject or "conta" in subject)


def extract_header_text(pdf_bytes: bytes) -> str:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages[:1])


def extract_claro_fields(header_text: str) -> dict:
    venc_m = VENCIMENTO_RE.search(header_text)
    valor_m = VALOR_RE.search(header_text)
    forma_m = FORMA_PGTO_RE.search(header_text)
    vencimento_br = venc_m.group(1) if venc_m else None
    vencimento_iso = None
    if vencimento_br:
        vencimento_iso = datetime.strptime(vencimento_br, "%d/%m/%Y").strftime("%Y-%m-%d")
    return {
        "vencimento_br": vencimento_br,
        "vencimento_iso": vencimento_iso,
        "valor": float(valor_m.group(1).replace(".", "").replace(",", ".")) if valor_m else None,
        "forma_pagamento": (forma_m.group(1).strip().title() if forma_m else None),
    }


def get_claro_folder_id(service, ref_date: datetime) -> str:
    """Contas/<ano>/<mês>/Pais/Claro Net Fibra — cria só as subpastas que faltarem."""
    year_id = get_or_resolve_year_folder(service, ref_date.year, CONTAS_ROOT_ID)
    month_id = get_or_resolve_month_folder(service, ref_date.month, year_id)
    pais_id = get_or_create_folder(service, "Pais", month_id)
    return get_or_create_folder(service, "Claro Net Fibra", pais_id)


def _next_available_name(service, folder_id: str, base_name: str, ext: str) -> str:
    """Segue a convenção já usada manualmente: 'Fatura Claro.pdf', '(3)', '(4)'..."""
    existing = service.files().list(
        q=f"'{folder_id}' in parents and trashed = false",
        fields="files(name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute().get("files", [])
    names = {f["name"] for f in existing}
    if f"{base_name}.{ext}" not in names:
        return f"{base_name}.{ext}"
    n = 2
    while f"{base_name} ({n}).{ext}" in names:
        n += 1
    return f"{base_name} ({n}).{ext}"


def upload_claro_pdf(pdf_bytes: bytes, ref_date: datetime) -> dict:
    service = get_drive_service()
    folder_id = get_claro_folder_id(service, ref_date)
    filename = _next_available_name(service, folder_id, "Fatura Claro", "pdf")

    media = MediaIoBaseUpload(io.BytesIO(pdf_bytes), mimetype="application/pdf")
    file_info = service.files().create(
        body={"name": filename, "parents": [folder_id]},
        media_body=media,
        fields="id, name, webViewLink",
        supportsAllDrives=True,
    ).execute()
    return file_info


async def notify_money_contas_a_pagar(valor: float, vencimento_iso: str, link_pdf: str, forma_pagamento: str) -> bool:
    """Cria o lançamento em Contas a Pagar no app money (idempotente lá, por template+vencimento)."""
    if not MONEY_WEBHOOK_SECRET:
        log.warning("MONEY_WEBHOOK_SECRET não configurado — lançamento em Contas a Pagar pulado.")
        return False
    async with httpx.AsyncClient(timeout=15) as client:
        resp = await client.post(
            MONEY_WEBHOOK_URL,
            headers={"Authorization": f"Bearer {MONEY_WEBHOOK_SECRET}"},
            json={
                "payee": CLARO_PAYEE_APP,
                "valor": valor,
                "vencimento": vencimento_iso,
                "linkPdf": link_pdf,
                "contaDebito": CLARO_CONTA_DEBITO,
                "formaPagamento": "Boleto" if "boleto" in (forma_pagamento or "").lower() else forma_pagamento,
            },
        )
        resp.raise_for_status()
        data = resp.json()
        log.info("Contas a Pagar (money): %s", data)
        return bool(data.get("created", data.get("ok")))


async def process_claro_attachment(att: dict, email_date_str: str) -> bool:
    """Ponto de entrada chamado por main.py (mesma assinatura de process_attachment,
    pra poder ser passado como `handler` em _search_and_process) para um anexo PDF
    de e-mail da Claro."""
    filename = att["filename"]
    pdf_bytes = att["data"]
    log.info("Processando fatura Claro: %s (%d bytes)", filename, len(pdf_bytes))

    try:
        pdf_bytes = unlock_pdf(pdf_bytes, CLARO_CNPJ_PREFIX)
        log.info("PDF Claro desbloqueado.")
    except ValueError as e:
        log.error("Senha incorreta pro PDF da Claro: %s", e)
        return False

    header_text = extract_header_text(pdf_bytes)
    fields = extract_claro_fields(header_text)
    if not fields["vencimento_iso"] or fields["valor"] is None:
        log.error("Não consegui extrair vencimento/valor do PDF da Claro. Texto: %r", header_text[:500])
        return False

    ref_date = datetime.strptime(fields["vencimento_iso"], "%Y-%m-%d")

    file_info = upload_claro_pdf(pdf_bytes, ref_date)
    link_pdf = file_info.get("webViewLink", "")
    log.info("PDF Claro no Drive: %s", link_pdf)

    is_boleto = "boleto" in (fields["forma_pagamento"] or "").lower()
    forma_pagto_str = "Boleto" if is_boleto else (fields["forma_pagamento"] or "")

    try:
        await notify_money_contas_a_pagar(fields["valor"], fields["vencimento_iso"], link_pdf, forma_pagto_str)
    except Exception:
        log.exception("Falha ao notificar o app money — segue o resto do fluxo mesmo assim.")

    try:
        upsert_boleto_row(
            ref_date=ref_date,
            dia=ref_date.day,
            payee=CLARO_PAYEE_SHEET,
            valor=fields["valor"],
            link_pdf=link_pdf,
            conta_debito=CLARO_CONTA_DEBITO,
            forma_pagamento=forma_pagto_str,
            link_pagamento="pending" if is_boleto else "Débito automático",
        )
    except Exception:
        log.exception("Falha ao atualizar a planilha Boletos — segue o resto do fluxo mesmo assim.")

    return True
