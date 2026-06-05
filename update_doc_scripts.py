"""
Documentação de Scripts — utilitários.

Funções:
  move_doc_to_versions()  — move doc antigo de Documentation/ para Documentation/Versions/<versão>/
  append_row()            — adiciona linha nova na planilha Documentação de Scripts

Uso: cd /Users/fabpovoa/mailhook && python3 update_doc_scripts.py
"""
from dotenv import load_dotenv
load_dotenv()
from gmail import get_credentials
from googleapiclient.discovery import build


SPREADSHEET_ID    = "1Y8MdFr6K0izpeejk4daW1JCGD5yu_dIXcMzX5BJBTbA"
DOC_FOLDER_ID     = "10wykUVtZMJsEfr2eBRgYB17GoVJFNI_y"   # Documentation/
VERSIONS_10_ID    = "1fqA-4Pz2OyRel3_XVwIn1OCfV3ZW5FGb"   # Documentation/Versions/1.0/
BAD_COPY_ID       = "1fY3OSEQxo7IVmKI5molbHhioswt3lmTJ_6h7Cw1sQbg"  # cópia vazia criada por engano
ORIGINAL_DOC_ID   = "1L6tjWhIMlkUVh8X52FXLODIjtx1y875W1kKUGGlW18Q"  # v1.0 ainda em Documentation/


def move_doc_to_versions():
    """Move o doc v1.0 de Documentation/ para Documentation/Versions/1.0/."""
    drive = build("drive", "v3", credentials=get_credentials())

    # 1. Apaga a cópia vazia criada anteriormente
    try:
        drive.files().delete(fileId=BAD_COPY_ID).execute()
        print(f"🗑️  Cópia vazia apagada: {BAD_COPY_ID}")
    except Exception as e:
        print(f"ℹ️  Cópia vazia não encontrada (ok): {e}")

    # 2. Move o original para Documentation/Versions/1.0/
    drive.files().update(
        fileId=ORIGINAL_DOC_ID,
        addParents=VERSIONS_10_ID,
        removeParents=DOC_FOLDER_ID,
        fields="id, parents",
    ).execute()
    print(f"✅ Doc v1.0 movido para Documentation/Versions/1.0/")


def append_row():
    """Remove todas as linhas 'Fatura BB' existentes e insere uma linha v1.1 limpa."""
    svc = build("sheets", "v4", credentials=get_credentials())
    sheet = svc.spreadsheets()

    # 1. Pega ID da primeira aba
    meta = sheet.get(spreadsheetId=SPREADSHEET_ID).execute()
    sheet_id = meta["sheets"][0]["properties"]["sheetId"]

    # 2. Lê colunas A:C para achar "Fatura BB" em qualquer uma delas
    all_rows = sheet.values().get(
        spreadsheetId=SPREADSHEET_ID, range="A:C"
    ).execute().get("values", [])

    to_delete = [
        i for i, r in enumerate(all_rows)
        if any("Fatura BB" in str(cell) for cell in r)
    ]
    for i in sorted(to_delete, reverse=True):
        sheet.batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": [{"deleteDimension": {"range": {
                "sheetId": sheet_id, "dimension": "ROWS",
                "startIndex": i, "endIndex": i + 1,
            }}}]}
        ).execute()
    print(f"🗑️  {len(to_delete)} linha(s) antigas removidas.")

    # 3. Insere linha nova — SEM coluna A vazia (o template já deixa A em branco)
    nova_linha = [
        "Fatura BB Cartões (Mailhook)",
        "✅ Ativo",
        "Processa faturas BB (PDF) e BTG (Excel+PDF) via Make mailhook. Extrai transações, separa por titular, gera xlsx multi-abas, salva no Drive e atualiza planilha dos pais.",
        "Python Local (Mac mini) — FastAPI + ngrok + Make",
        "/Users/fabpovoa/mailhook/",
        "fabio@smartmoney.ventures",
        "Gmail API; Drive API v3; Sheets API v4; Make (mailhook)",
        "BB: PDF + xlsx (4 abas) em BB Altus/; BTG: Excel+PDF em BTG Ultra/; Planilha dos pais atualizada",
        "Event-driven (Make mailhook — BB: digital@faturaourocard.com.br; BTG: email Fátima)",
        "Make mailhook + LaunchAgents macOS",
        "https://drive.google.com/drive/folders/1BuPhXc_6wZOqptlRpBosW_opV6u7TZj8",
        "https://docs.google.com/document/d/18ddlJgA34TL8auzJjJPndNKgfI5ZB1BK_xtWNUkIbgQ/edit",
        "03/06/2026",
        "1.1",
        "04/06/2026",
        "Fábio Póvoa",
        "fastapi, pdfplumber, pikepdf, openpyxl, msoffcrypto-tool, google-api-python-client, ngrok, make.com",
    ]
    svc = build("sheets", "v4", credentials=get_credentials())
    svc.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range="B:R",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": [nova_linha]},
    ).execute()
    print("✅ Linha v1.1 adicionada em Documentação de Scripts.")


if __name__ == "__main__":
    move_doc_to_versions()
    append_row()
