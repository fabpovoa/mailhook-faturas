"""
Testa a criação do xls do próximo mês a partir do atual.

Uso:
  cd /Users/fabpovoa/mailhook
  python3 test_next_month_xls.py

Baixa o xls de Junho do Drive, aplica a transformação e salva
preview_proximo_mes.xlsx localmente para conferir.
"""
import io
import logging
from datetime import datetime
from dotenv import load_dotenv
load_dotenv()

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

from gmail import get_credentials
from googleapiclient.discovery import build
from next_month_xls import create_next_month_xls

# ID do xls atual dos pais (Jun ref May 26)
FILE_ID  = "1VB4L2d2-kfJRm4DzVODMPcDYk6bteUBo"
REF_DATE = datetime(2026, 6, 1)


def main():
    drive = build("drive", "v3", credentials=get_credentials())

    log.info("Baixando xls atual (Jun ref May)...")
    content = drive.files().get_media(fileId=FILE_ID, supportsAllDrives=True).execute()
    log.info("Baixado: %d bytes", len(content))

    log.info("Aplicando transformação para %s...", REF_DATE.strftime("%b/%Y"))
    new_bytes, filename = create_next_month_xls(content, REF_DATE)

    out_path = f"/Users/fabpovoa/mailhook/preview_{filename}"
    with open(out_path, "wb") as f:
        f.write(new_bytes)

    log.info("✅ Preview salvo: %s", out_path)
    log.info("   Arquivo gerado: %s (%d bytes)", filename, len(new_bytes))


if __name__ == "__main__":
    main()
