"""
Pipeline BTG — baseado no Excel fornecido pelo banco (sem processamento de PDF).

Fluxo:
  1. Desbloqueia e extrai transações do Excel BTG
  2. Cria pasta BTG Ultra no Drive (Cartões/<ano>/<mês>/BTG Ultra/)
  3. Salva o Excel original na pasta BTG Ultra (compartilhado publicamente)
  4. Atualiza aba BTG da planilha dos pais (Luiz + Fátima, sem Hiero)

Uso:
  cd /Users/fabpovoa/mailhook
  python3 test_btg_drive.py <caminho_do_excel.xlsx>
"""
import sys
import os
import logging
import io
from datetime import datetime

from dotenv import load_dotenv
load_dotenv()

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

BTG_EXCEL_PASSWORD = os.getenv("BTG_EXCEL_PASSWORD", "57189560149")


def main(excel_path: str):
    from btg_excel_extractor import extract_btg_excel
    from drive import get_drive_service, get_btg_folder_id
    from faturas_updater import update_btg_pais
    from googleapiclient.http import MediaIoBaseUpload

    log.info("=" * 60)
    log.info("BTG Excel: %s", excel_path)

    with open(excel_path, "rb") as f:
        xlsx_bytes = f.read()

    # 1. Extrai transações do Excel
    log.info("1. Extraindo transações do Excel BTG...")
    grouped, ref_date = extract_btg_excel(xlsx_bytes, BTG_EXCEL_PASSWORD)
    if ref_date is None:
        ref_date = datetime.now()
    log.info("   Data ref: %s", ref_date.strftime("%B/%Y"))
    for tab, txs in grouped.items():
        log.info("   %s: %d transações", tab, len(txs))

    excel_name = f"fatura_btg_{ref_date.strftime('%b%Y').lower()}.xlsx"

    # 2. Cria/localiza pasta BTG Ultra no Drive
    log.info("2. Criando pasta BTG Ultra no Drive...")
    drive_service = get_drive_service()
    btg_folder_id = get_btg_folder_id(drive_service, ref_date)
    log.info("   ✅ Pasta BTG Ultra: %s", btg_folder_id)

    # 3. Salva o Excel original na pasta BTG Ultra e compartilha publicamente
    log.info("3. Salvando Excel no Drive (BTG Ultra)...")
    file_info = drive_service.files().create(
        body={"name": excel_name, "parents": [btg_folder_id]},
        media_body=MediaIoBaseUpload(
            io.BytesIO(xlsx_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ),
        fields="id, webViewLink",
        supportsAllDrives=True,
    ).execute()
    file_id = file_info["id"]

    # Compartilha para qualquer pessoa com o link
    drive_service.permissions().create(
        fileId=file_id,
        body={"type": "anyone", "role": "reader"},
        supportsAllDrives=True,
    ).execute()
    file_link = file_info.get("webViewLink", "")
    log.info("   ✅ Excel → %s", file_link)

    # 4. Atualiza aba BTG na planilha dos pais
    # NOTA: para teste usa a pasta Pais de Junho (onde está a planilha atual)
    pais_ref_date = datetime(2026, 6, 1)  # TODO: remover após junho ter planilha própria
    log.info("5. Atualizando aba BTG na planilha dos pais (Pais de %s)...", pais_ref_date.strftime("%b/%Y"))
    link = update_btg_pais(grouped, pais_ref_date, pdf_link=file_link)
    if link:
        log.info("   ✅ Planilha dos pais → %s", link)
    else:
        log.error("   ❌ Planilha dos pais não atualizada.")

    log.info("=" * 60)
    log.info("DONE")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python3 test_btg_drive.py <caminho_excel.xlsx>")
        sys.exit(1)
    main(sys.argv[1])
