"""
Autenticação Gmail e leitura de anexos PDF.
"""
import base64
import os
from typing import Optional

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = [
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]


def get_credentials() -> Credentials:
    """Carrega ou renova credenciais OAuth2."""
    creds = None
    token_file = os.getenv("GOOGLE_TOKEN_FILE", "token.json")
    creds_file = os.getenv("GOOGLE_CREDENTIALS_FILE", "credentials.json")

    if os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(creds_file, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_file, "w") as f:
            f.write(creds.to_json())

    return creds


def get_gmail_service():
    return build("gmail", "v1", credentials=get_credentials())


def get_message(service, user_id: str, message_id: str) -> dict:
    return service.users().messages().get(
        userId=user_id, id=message_id, format="full"
    ).execute()


def mark_message_read(service, user_id: str, message_id: str) -> None:
    """Remove o marcador UNREAD para evitar reprocessamento via Pub/Sub."""
    service.users().messages().modify(
        userId=user_id,
        id=message_id,
        body={"removeLabelIds": ["UNREAD"]},
    ).execute()


def extract_pdf_attachments(service, user_id: str, message: dict) -> list[dict]:
    """
    Extrai todos os anexos PDF de uma mensagem.
    Retorna lista de {"filename": str, "data": bytes}.
    """
    attachments = []
    parts = message.get("payload", {}).get("parts", [])

    for part in parts:
        mime_type = part.get("mimeType", "")
        filename = part.get("filename", "")

        if "pdf" not in mime_type.lower() and not filename.lower().endswith(".pdf"):
            continue

        body = part.get("body", {})
        attachment_id = body.get("attachmentId")
        data = body.get("data")

        if attachment_id:
            # Anexo separado — busca pelo ID
            att = service.users().messages().attachments().get(
                userId=user_id, messageId=message["id"], id=attachment_id
            ).execute()
            data = att.get("data", "")

        if data:
            pdf_bytes = base64.urlsafe_b64decode(data + "==")
            attachments.append({"filename": filename or "fatura.pdf", "data": pdf_bytes})

    return attachments


def is_fatura_bb(message: dict) -> bool:
    """
    Verifica se o e-mail é de fatura do BB Cartões.
    Ajuste os filtros conforme o remetente real.
    """
    headers = {h["name"].lower(): h["value"] for h in message.get("payload", {}).get("headers", [])}
    from_email = headers.get("from", "").lower()
    subject = headers.get("subject", "").lower()

    return (
        "bb" in from_email or "bancodobrasil" in from_email or "bradescard" in from_email
        or "fatura" in subject
        or "cartão" in subject
        or "cartao" in subject
    )
