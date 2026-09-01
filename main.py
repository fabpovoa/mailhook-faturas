"""
FastAPI webhook — recebe notificações do Gmail via Google Cloud Pub/Sub,
extrai PDFs de fatura BB Altus, desbloqueia, salva no Drive e gera Excel.

Fluxo:
  Gmail (digital@faturaourocard.com.br) → Pub/Sub → /webhook
       → desbloqueia PDF (pikepdf)
       → salva PDF no Drive (Cartões/<ano>/<mês>/BB Altus/)
       → extrai transações (pdfplumber) → gera xlsx (openpyxl)
       → salva xlsx no Drive
       → atualiza faturas_Cartões_Luiz (abas Luiz + Fátima)
"""
import base64
import json
import os
import logging
from contextlib import asynccontextmanager
from datetime import datetime

from dotenv import load_dotenv
load_dotenv()

from fastapi import FastAPI, Request, HTTPException, BackgroundTasks

from gmail import get_gmail_service, get_message, extract_pdf_attachments, mark_message_read
from drive import upload_file, get_drive_service, get_btg_folder_id
from pdf_handler import unlock_pdf
from pdf_extractor import extract_transactions, extract_header_text, get_fatura_title
from xlsx_writer import build_xlsx
from faturas_updater import group_by_owner, update_faturas_luiz, update_btg_pais
from btg_excel_extractor import extract_btg_excel
from claro_handler import process_claro_attachment

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

GMAIL_USER         = os.getenv("GMAIL_USER", "me")
PDF_PASSWORD       = os.getenv("PDF_PASSWORD", "")
BTG_EXCEL_PASSWORD = os.getenv("BTG_EXCEL_PASSWORD", "")
WEBHOOK_SECRET     = os.getenv("WEBHOOK_SECRET", "")

# Filtro preciso: só e-mails da faturaourocard com PDF anexo
GMAIL_QUERY = "from:digital@faturaourocard.com.br has:attachment filename:pdf is:unread"

# Fatura Claro (conta Pais/SMV1) — pode chegar direto ou encaminhada (fabio@povoa.com),
# por isso o filtro de assunto além do remetente direto (ver claro_handler.is_fatura_claro).
CLARO_GMAIL_QUERY = '(from:faturadigital@minhaclaro.com.br OR subject:"Fatura Digital Claro" OR subject:"Fatura Claro") has:attachment filename:pdf is:unread'


@asynccontextmanager
async def lifespan(app: FastAPI):
    log.info("🚀 Mailhook iniciado. Aguardando notificações do Pub/Sub...")
    yield
    log.info("Mailhook encerrado.")


app = FastAPI(title="Mailhook — Fatura BB Altus", lifespan=lifespan)


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/webhook")
async def webhook(request: Request, background_tasks: BackgroundTasks):
    """Endpoint chamado pelo Google Cloud Pub/Sub quando chega e-mail novo."""
    token = request.query_params.get("token", "")
    if WEBHOOK_SECRET and token != WEBHOOK_SECRET:
        raise HTTPException(status_code=401, detail="Token inválido.")

    body = await request.json()
    message = body.get("message", {})
    data_b64 = message.get("data", "")

    if not data_b64:
        log.warning("Webhook sem campo 'data'. Ignorando.")
        return {"status": "ignored"}

    payload = json.loads(base64.b64decode(data_b64).decode("utf-8"))
    email_address = payload.get("emailAddress", GMAIL_USER)
    log.info(f"Notificação recebida — historyId={payload.get('historyId')}")

    background_tasks.add_task(process_new_emails, email_address)
    return {"status": "accepted"}


async def process_new_emails(email_address: str):
    """Busca e-mails não lidos com PDF: fatura BB Altus e fatura Claro (Pais)."""
    try:
        service = get_gmail_service()
        await _search_and_process(service, email_address, GMAIL_QUERY, process_attachment, "fatura BB")
        await _search_and_process(service, email_address, CLARO_GMAIL_QUERY, process_claro_attachment, "fatura Claro")
    except Exception as e:
        log.exception(f"Erro ao processar e-mails: {e}")


async def _search_and_process(service, email_address: str, query: str, handler, label: str):
    """Busca mensagens pela query, roda `handler(att, email_date_str)` em cada anexo PDF
    e marca como lida só se todos os anexos da mensagem forem processados com sucesso."""
    results = service.users().messages().list(
        userId=email_address,
        q=query,
        maxResults=10,
    ).execute()

    messages = results.get("messages", [])
    if not messages:
        log.info(f"Nenhuma {label} nova.")
        return

    for msg_ref in messages:
        msg = get_message(service, email_address, msg_ref["id"])
        headers = {h["name"].lower(): h["value"]
                   for h in msg.get("payload", {}).get("headers", [])}
        email_date_str = headers.get("date", "")

        attachments = extract_pdf_attachments(service, email_address, msg)
        if not attachments:
            log.info(f"Mensagem {msg_ref['id']} sem PDF.")
            continue

        all_ok = True
        for att in attachments:
            if not await handler(att, email_date_str):
                all_ok = False

        if all_ok:
            mark_message_read(service, email_address, msg_ref["id"])
            log.info(f"Mensagem {msg_ref['id']} ({label}) marcada como lida.")


async def process_attachment(att: dict, email_date_str: str) -> bool:
    """Desbloqueia PDF, salva no Drive, gera xlsx e atualiza faturas_Cartões_Luiz."""
    filename = att["filename"]
    pdf_bytes = att["data"]
    log.info(f"Processando: {filename} ({len(pdf_bytes)} bytes)")

    # 1. Desbloqueia PDF
    if PDF_PASSWORD:
        try:
            pdf_bytes = unlock_pdf(pdf_bytes, PDF_PASSWORD)
            log.info("PDF desbloqueado.")
        except ValueError as e:
            log.error(f"Senha incorreta: {e}")
            return False

    # 2. Extrai data e título do PDF
    header_text = extract_header_text(pdf_bytes)
    fatura_title = get_fatura_title(header_text)
    transactions, ref_date = extract_transactions(pdf_bytes)
    if ref_date is None:
        ref_date = datetime.now()

    month_str = ref_date.strftime("%b%Y").lower()  # ex: jun2026
    pdf_name  = f"fatura_bb_{month_str}.pdf"
    xlsx_name = f"fatura_bb_{month_str}.xlsx"

    # 3. Salva PDF no Drive
    pdf_info = upload_file(pdf_bytes, pdf_name, "application/pdf", ref_date)
    log.info(f"PDF no Drive: {pdf_info.get('webViewLink')}")

    if not transactions:
        log.warning("Nenhuma transação extraída — Excel e atualização pulados.")
        return True

    log.info(f"{len(transactions)} transações extraídas.")

    # 4. Agrupa por titular e gera xlsx com uma aba por cartão
    grouped = group_by_owner(transactions)
    xlsx_bytes = build_xlsx(grouped, fatura_title, ref_date)
    xlsx_info = upload_file(
        xlsx_bytes, xlsx_name,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ref_date,
    )
    log.info(f"Excel no Drive: {xlsx_info.get('webViewLink')}")

    # 5. Atualiza planilha mestre (já usa grouped internamente)
    link = update_faturas_luiz(transactions, ref_date)
    if link:
        log.info(f"faturas_Cartões_Luiz atualizado: {link}")
    else:
        log.warning("faturas_Cartões_Luiz não encontrado ou não atualizado.")

    return True


# ─────────────────────────────────────────────────────────────
# Endpoint Make — recebe folder_id com anexo já salvo no Drive
# ─────────────────────────────────────────────────────────────

@app.post("/webhook/make")
async def webhook_make(request: Request, background_tasks: BackgroundTasks):
    """
    Chamado pelo Make após salvar o anexo em uma pasta temp no Drive.
    Payload: {"folder_id": "xxx"}
    """
    token = request.query_params.get("token", "")
    if WEBHOOK_SECRET and token != WEBHOOK_SECRET:
        raise HTTPException(status_code=401, detail="Token inválido.")

    body = await request.json()
    folder_id = body.get("folder_id", "")
    card_type = body.get("card_type", "").upper()  # "BB", "BTG" ou vazio (auto-detect)

    if not folder_id:
        raise HTTPException(status_code=400, detail="folder_id obrigatório.")

    log.info("📬 Make webhook — folder_id: %s | card_type: %s", folder_id, card_type or "auto")
    background_tasks.add_task(process_make_folder, folder_id, card_type)
    return {"status": "accepted"}


async def process_make_folder(folder_id: str, card_type: str = ""):
    """Lista arquivos na pasta temp, roteia para BB ou BTG e apaga a pasta ao final."""
    from googleapiclient.http import MediaIoBaseDownload
    import io as _io

    def download(file_id: str) -> bytes:
        buf = _io.BytesIO()
        dl = MediaIoBaseDownload(buf, drive.files().get_media(fileId=file_id, supportsAllDrives=True))
        done = False
        while not done:
            _, done = dl.next_chunk()
        return buf.getvalue()

    try:
        drive = get_drive_service()

        results = drive.files().list(
            q=f"'{folder_id}' in parents and trashed = false",
            fields="files(id, name, mimeType)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        files = results.get("files", [])

        if not files:
            log.warning("Pasta temp vazia: %s", folder_id)
            return

        log.info("%d arquivo(s) encontrado(s) na pasta temp.", len(files))
        for f in files:
            log.info("  • %s", f["name"])

        if card_type == "BB" or (not card_type and all(f["name"].lower().endswith(".pdf") for f in files)):
            # Pipeline BB — único PDF
            pdf_file = next((f for f in files if f["name"].lower().endswith(".pdf")), None)
            if pdf_file:
                await _process_bb(download(pdf_file["id"]), pdf_file["name"])
        elif card_type == "BTG" or any(f["name"].lower().endswith(".xlsx") for f in files):
            # Pipeline BTG — tem PDF + XLSX; desbloqueia ambos, copia para BTG Ultra, processa só o XLSX
            await _process_btg_full(files, drive, download)
        else:
            log.warning("Não foi possível determinar o pipeline para os arquivos recebidos.")

        # Apaga pasta temp
        drive.files().delete(fileId=folder_id, supportsAllDrives=True).execute()
        log.info("🗑️  Pasta temp apagada: %s", folder_id)

    except Exception as e:
        log.exception("Erro ao processar pasta Make: %s", e)


async def _process_btg_full(files: list, drive, download_fn):
    """
    Pipeline BTG completo:
    1. Baixa PDF e XLSX da pasta temp
    2. Desbloqueia ambos com a senha BTG (CPF da mãe)
    3. Cria pasta BTG Ultra no Drive e copia ambos (compartilhados publicamente)
    4. Processa apenas o XLSX
    """
    from googleapiclient.http import MediaIoBaseUpload
    import io as _io
    import pikepdf

    pdf_file  = next((f for f in files if f["name"].lower().endswith(".pdf")),  None)
    xlsx_file = next((f for f in files if f["name"].lower().endswith(".xlsx") or f["name"].lower().endswith(".xls")), None)

    if not xlsx_file:
        log.error("BTG: XLSX não encontrado na pasta temp.")
        return

    # 1. Baixa os arquivos
    xlsx_bytes = download_fn(xlsx_file["id"])
    pdf_bytes  = download_fn(pdf_file["id"]) if pdf_file else None

    # 2. Desbloqueia XLSX
    try:
        import msoffcrypto
        enc = msoffcrypto.OfficeFile(_io.BytesIO(xlsx_bytes))
        enc.load_key(password=BTG_EXCEL_PASSWORD)
        buf = _io.BytesIO()
        enc.decrypt(buf)
        xlsx_unlocked = buf.getvalue()
        log.info("BTG: XLSX desbloqueado.")
    except Exception as e:
        log.error("BTG: erro ao desbloquear XLSX: %s", e)
        return

    # 3. Desbloqueia PDF (se existir)
    pdf_unlocked = None
    if pdf_bytes:
        try:
            tmp_in  = _io.BytesIO(pdf_bytes)
            tmp_out = _io.BytesIO()
            with pikepdf.open(tmp_in, password=BTG_EXCEL_PASSWORD) as p:
                p.save(tmp_out)
            pdf_unlocked = tmp_out.getvalue()
            log.info("BTG: PDF desbloqueado.")
        except Exception as e:
            log.warning("BTG: PDF não pôde ser desbloqueado (%s) — copiando original.", e)
            pdf_unlocked = pdf_bytes

    # 4. Extrai ref_date do XLSX para criar pasta no mês correto
    grouped, ref_date = extract_btg_excel(xlsx_unlocked, BTG_EXCEL_PASSWORD)
    if ref_date is None:
        ref_date = datetime.now()

    month_str     = ref_date.strftime("%b%Y").lower()
    btg_folder_id = get_btg_folder_id(drive, ref_date)
    log.info("BTG: pasta BTG Ultra pronta: %s", btg_folder_id)

    def upload_shared(data: bytes, name: str, mime: str) -> str:
        info = drive.files().create(
            body={"name": name, "parents": [btg_folder_id]},
            media_body=MediaIoBaseUpload(_io.BytesIO(data), mimetype=mime),
            fields="id, webViewLink",
            supportsAllDrives=True,
        ).execute()
        drive.permissions().create(
            fileId=info["id"],
            body={"type": "anyone", "role": "reader"},
            supportsAllDrives=True,
        ).execute()
        log.info("BTG: %s → %s", name, info.get("webViewLink"))
        return info.get("webViewLink", "")

    # 5. Copia ambos para a pasta BTG Ultra
    xlsx_link = upload_shared(xlsx_unlocked, f"fatura_btg_{month_str}.xlsx",
                              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if pdf_unlocked:
        upload_shared(pdf_unlocked, f"fatura_btg_{month_str}.pdf", "application/pdf")

    # 6. Processa o XLSX (atualiza planilha dos pais)
    update_btg_pais(grouped, ref_date, pdf_link=xlsx_link)
    log.info("BTG: pipeline completo.")


async def _process_bb(pdf_bytes: bytes, filename: str):
    """Pipeline BB completo."""
    if PDF_PASSWORD:
        try:
            pdf_bytes = unlock_pdf(pdf_bytes, PDF_PASSWORD)
        except ValueError as e:
            log.error("Senha BB incorreta: %s", e)
            return

    header_text = extract_header_text(pdf_bytes)
    fatura_title = get_fatura_title(header_text)
    transactions, ref_date = extract_transactions(pdf_bytes)
    if ref_date is None:
        ref_date = datetime.now()

    month_str = ref_date.strftime("%b%Y").lower()
    upload_file(pdf_bytes, f"fatura_bb_{month_str}.pdf", "application/pdf", ref_date)

    if not transactions:
        log.warning("BB: nenhuma transação extraída.")
        return

    grouped = group_by_owner(transactions)
    xlsx_bytes = build_xlsx(grouped, fatura_title, ref_date)
    upload_file(xlsx_bytes, f"fatura_bb_{month_str}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ref_date)
    update_faturas_luiz(transactions, ref_date)
    log.info("BB: pipeline concluído.")


async def _process_btg_pdf(pdf_bytes: bytes, filename: str):
    """Pipeline BTG via PDF (fallback quando BTG envia PDF em vez de Excel)."""
    from btg_extractor import extract_btg_transactions
    from googleapiclient.http import MediaIoBaseUpload
    import io as _io

    grouped, ref_date = extract_btg_transactions(pdf_bytes)
    if ref_date is None:
        ref_date = datetime.now()

    drive = get_drive_service()
    month_str = ref_date.strftime("%b%Y").lower()
    btg_folder_id = get_btg_folder_id(drive, ref_date)

    file_info = drive.files().create(
        body={"name": f"fatura_btg_{month_str}.pdf", "parents": [btg_folder_id]},
        media_body=MediaIoBaseUpload(_io.BytesIO(pdf_bytes), mimetype="application/pdf"),
        fields="id, webViewLink",
        supportsAllDrives=True,
    ).execute()
    drive.permissions().create(
        fileId=file_info["id"],
        body={"type": "anyone", "role": "reader"},
        supportsAllDrives=True,
    ).execute()

    update_btg_pais(grouped, ref_date, pdf_link=file_info.get("webViewLink", ""))
    log.info("BTG PDF: pipeline concluído.")


async def _process_btg(xlsx_bytes: bytes, filename: str):
    """Pipeline BTG completo."""
    from googleapiclient.http import MediaIoBaseUpload
    import io as _io

    grouped, ref_date = extract_btg_excel(xlsx_bytes, BTG_EXCEL_PASSWORD)
    if ref_date is None:
        ref_date = datetime.now()

    drive = get_drive_service()
    month_str = ref_date.strftime("%b%Y").lower()
    btg_folder_id = get_btg_folder_id(drive, ref_date)

    file_info = drive.files().create(
        body={"name": f"fatura_btg_{month_str}.xlsx", "parents": [btg_folder_id]},
        media_body=MediaIoBaseUpload(
            _io.BytesIO(xlsx_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ),
        fields="id, webViewLink",
        supportsAllDrives=True,
    ).execute()
    drive.permissions().create(
        fileId=file_info["id"],
        body={"type": "anyone", "role": "reader"},
        supportsAllDrives=True,
    ).execute()

    update_btg_pais(grouped, ref_date, pdf_link=file_info.get("webViewLink", ""))
    log.info("BTG: pipeline concluído.")
