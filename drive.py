"""
Google Drive: navega/cria estrutura de pastas e faz upload de arquivos.

Estrutura esperada:
  Cartões (1p2sOYnyL4AI1oDoz8TvbKDFr3N2PNhkw)
    └── 2026
        └── 06_jun
            └── BB Altus
                ├── fatura_bb_jun2026.pdf
                └── fatura_bb_jun2026.xlsx
"""
import io
from datetime import datetime
from typing import Optional

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

from gmail import get_credentials

# ID da pasta raiz "Cartões" no Drive
CARTOES_ROOT_ID = "1p2sOYnyL4AI1oDoz8TvbKDFr3N2PNhkw"

MONTHS_PT = {
    1: "01_jan", 2: "02_fev", 3: "03_mar", 4: "04_abr",
    5: "05_mai", 6: "06_jun", 7: "07_jul", 8: "08_ago",
    9: "09_set", 10: "10_out", 11: "11_nov", 12: "12_dez",
}


def month_folder_names(month: int) -> list[str]:
    """Nomes a procurar: canónico (06_jun) e legado (6_jun), para não duplicar pastas."""
    canonical = MONTHS_PT[month]
    names = [canonical]
    if month < 10:
        suffix = canonical.split("_", 1)[1]
        legacy = f"{month}_{suffix}"
        if legacy not in names:
            names.append(legacy)
    return names


def _escape_drive_query(value: str) -> str:
    return value.replace("\\", "\\\\").replace("'", "\\'")


def _drive_list_kwargs() -> dict:
    """Pastas partilhadas (ex.: Cartões) exigem estes parâmetros na API v3."""
    return {
        "supportsAllDrives": True,
        "includeItemsFromAllDrives": True,
    }


def get_drive_service():
    return build("drive", "v3", credentials=get_credentials())


def list_child_folders(service, parent_id: str) -> list[dict]:
    """Lista subpastas diretas de parent_id (com paginação)."""
    q = (
        f"'{parent_id}' in parents and mimeType = 'application/vnd.google-apps.folder'"
        f" and trashed = false"
    )
    folders: list[dict] = []
    page_token = None
    while True:
        kwargs = {
            "q": q,
            "fields": "nextPageToken, files(id, name)",
            "pageSize": 200,
            **_drive_list_kwargs(),
        }
        if page_token:
            kwargs["pageToken"] = page_token
        results = service.files().list(**kwargs).execute()
        folders.extend(results.get("files", []))
        page_token = results.get("nextPageToken")
        if not page_token:
            break
    return folders


def find_folder(service, name: str, parent_id: str) -> Optional[str]:
    """Retorna o ID da pasta se existir sob parent_id (incluindo drives compartilhados)."""
    safe_name = _escape_drive_query(name)
    q = (
        f"name = '{safe_name}' and mimeType = 'application/vnd.google-apps.folder'"
        f" and '{parent_id}' in parents and trashed = false"
    )
    # Correção aqui: Incluindo os kwargs de drives compartilhados na busca por nome
    results = service.files().list(
        q=q, fields="files(id, name)", **_drive_list_kwargs()
    ).execute()
    files = results.get("files", [])
    if files:
        return files[0]["id"]

    target = name.strip()
    for folder in list_child_folders(service, parent_id):
        if folder["name"].strip() == target:
            return folder["id"]
    return None


def find_year_folder(service, year: int, parent_id: str) -> Optional[str]:
    """Procura pasta do ano (ex.: 2026) em Cartões sem criar duplicata."""
    return find_folder(service, str(year), parent_id)


def find_month_folder(service, month: int, year_id: str) -> Optional[str]:
    """Procura pasta mensal existente (formato novo ou legado)."""
    for name in month_folder_names(month):
        folder_id = find_folder(service, name, year_id)
        if folder_id:
            return folder_id

    targets = {n.strip() for n in month_folder_names(month)}
    for folder in list_child_folders(service, year_id):
        if folder["name"].strip() in targets:
            return folder["id"]
    return None


def get_or_create_folder(service, name: str, parent_id: str) -> str:
    """Retorna o ID de uma pasta (cria só se não existir)."""
    existing = find_folder(service, name, parent_id)
    if existing:
        return existing

    folder = service.files().create(
        body={
            "name": name,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [parent_id],
        },
        fields="id",
        supportsAllDrives=True,
    ).execute()
    return folder["id"]


def get_or_resolve_year_folder(service, year: int, parent_id: str) -> str:
    """Reutiliza pasta do ano existente; só cria se não houver nenhuma."""
    existing = find_year_folder(service, year, parent_id)
    if existing:
        return existing
    return get_or_create_folder(service, str(year), parent_id)


def get_or_resolve_month_folder(service, month: int, year_id: str) -> str:
    """Reutiliza pasta mensal existente; só cria com nome canónico se nenhuma existir."""
    existing = find_month_folder(service, month, year_id)
    if existing:
        return existing
    return get_or_create_folder(service, MONTHS_PT[month], year_id)


def get_target_folder_id(service, ref_date: Optional[datetime] = None) -> str:
    """
    Retorna o ID da pasta destino: Cartões / <ano> / <mês> / BB Altus
    Cria as subpastas se não existirem.
    """
    if ref_date is None:
        ref_date = datetime.now()

    year_id = get_or_resolve_year_folder(service, ref_date.year, CARTOES_ROOT_ID)
    month_id = get_or_resolve_month_folder(service, ref_date.month, year_id)
    bb_id = get_or_create_folder(service, "BB Altus", month_id)
    return bb_id


def get_btg_folder_id(service, ref_date: Optional[datetime] = None) -> str:
    """
    Retorna o ID da pasta destino: Cartões / <ano> / <mês> / BTG Ultra
    Cria as subpastas se não existirem.
    """
    if ref_date is None:
        ref_date = datetime.now()

    year_id = get_or_resolve_year_folder(service, ref_date.year, CARTOES_ROOT_ID)
    month_id = get_or_resolve_month_folder(service, ref_date.month, year_id)
    btg_id = get_or_create_folder(service, "BTG Ultra", month_id)
    return btg_id


def upload_file(
    file_bytes: bytes,
    filename: str,
    mime_type: str,
    ref_date: Optional[datetime] = None,
) -> dict:
    """
    Faz upload de um arquivo para Cartões/<ano>/<mês>/BB Altus/.
    Retorna: {"id": str, "name": str, "webViewLink": str}
    """
    service = get_drive_service()
    folder_id = get_target_folder_id(service, ref_date)

    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mime_type)
    file_info = service.files().create(
        body={"name": filename, "parents": [folder_id]},
        media_body=media,
        fields="id, name, webViewLink",
        supportsAllDrives=True,
    ).execute()
    return file_info
