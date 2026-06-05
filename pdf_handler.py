"""
Desbloqueia PDFs protegidos por senha usando pikepdf.
"""
import pikepdf
import os
import tempfile
from pathlib import Path


def unlock_pdf(pdf_bytes: bytes, password: str) -> bytes:
    """
    Recebe o PDF em bytes com senha, retorna PDF desbloqueado em bytes.
    Lança ValueError se a senha estiver errada.
    """
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp_in:
        tmp_in.write(pdf_bytes)
        tmp_in_path = tmp_in.name

    tmp_out_path = tmp_in_path.replace(".pdf", "_unlocked.pdf")

    try:
        with pikepdf.open(tmp_in_path, password=password) as pdf:
            pdf.save(tmp_out_path)

        with open(tmp_out_path, "rb") as f:
            return f.read()

    except pikepdf.PasswordError:
        raise ValueError("Senha do PDF incorreta.")
    finally:
        Path(tmp_in_path).unlink(missing_ok=True)
        Path(tmp_out_path).unlink(missing_ok=True)
