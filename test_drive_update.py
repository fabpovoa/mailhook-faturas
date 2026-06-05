"""
Teste completo do pipeline: desbloqueio → extração → split → Drive upload + Sheets update.

Como usar:
  cd /Users/fabpovoa/mailhook
  python3 test_drive_update.py <caminho_do_pdf>

Exemplo:
  python3 test_drive_update.py "CARTAO ALTUS VISA (1).pdf"
"""
import sys
import os
import logging
from datetime import datetime

from dotenv import load_dotenv
load_dotenv()

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

PDF_PASSWORD = os.getenv("PDF_PASSWORD", "")


def main(pdf_path: str):
    from pdf_handler import unlock_pdf
    from pdf_extractor import extract_transactions, extract_header_text, get_fatura_title
    from faturas_updater import group_by_owner, update_faturas_luiz
    from xlsx_writer import build_xlsx
    from drive import upload_file

    log.info("=" * 60)
    log.info("PDF: %s", pdf_path)

    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()

    # 1. Desbloqueia
    if PDF_PASSWORD:
        log.info("1. Desbloqueando PDF (senha: %s)...", PDF_PASSWORD)
        pdf_bytes = unlock_pdf(pdf_bytes, PDF_PASSWORD)
        log.info("   ✅ PDF desbloqueado (%d bytes)", len(pdf_bytes))
    else:
        log.info("1. PDF_PASSWORD não definida — assumindo PDF sem senha.")

    # 2. Extrai
    log.info("2. Extraindo transações...")
    header_text = extract_header_text(pdf_bytes)
    fatura_title = get_fatura_title(header_text)
    transactions, ref_date = extract_transactions(pdf_bytes)
    if ref_date is None:
        ref_date = datetime.now()
    log.info("   Título: %s", fatura_title)
    log.info("   Data ref: %s", ref_date)
    log.info("   Total bruto: %d transações", len(transactions))

    month_str = ref_date.strftime("%b%Y").lower()
    pdf_name  = f"fatura_bb_{month_str}.pdf"
    xlsx_name = f"fatura_bb_{month_str}.xlsx"

    # 3. Salva PDF desbloqueado no Drive
    log.info("3. Salvando PDF desbloqueado no Drive...")
    pdf_info = upload_file(pdf_bytes, pdf_name, "application/pdf", ref_date)
    log.info("   ✅ PDF → %s", pdf_info.get("webViewLink"))

    if not transactions:
        log.warning("Nenhuma transação extraída — encerrando.")
        return

    # 4. Agrupa por titular
    log.info("4. Agrupando por titular...")
    grouped = group_by_owner(transactions)
    for owner, txs in grouped.items():
        log.info("   %s: %d transações", owner, len(txs))

    # 5. Gera e salva xlsx splitado no Drive
    log.info("5. Gerando xlsx splitado e salvando no Drive...")
    xlsx_bytes = build_xlsx(grouped, fatura_title, ref_date)
    xlsx_info = upload_file(
        xlsx_bytes, xlsx_name,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ref_date,
    )
    log.info("   ✅ xlsx → %s", xlsx_info.get("webViewLink"))

    # 6. Atualiza planilha dos pais (BB Altus Visa Luiz + Fátima)
    log.info("6. Atualizando planilha dos pais no Drive...")
    link = update_faturas_luiz(transactions, ref_date)
    if link:
        log.info("   ✅ Planilha dos pais → %s", link)
    else:
        log.error("   ❌ Planilha dos pais não atualizada.")

    log.info("=" * 60)
    log.info("DONE")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python3 test_drive_update.py <caminho_do_pdf>")
        sys.exit(1)
    main(sys.argv[1])
