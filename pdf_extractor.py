"""
Extrai transações de uma fatura BB Altus (PDF desbloqueado).

Estrutura esperada do PDF BB Cartões:
  Linha de cabeçalho: "Lançamentos  Fatura BB – <Mês> <data>"
  Colunas: Titular | Data | Descrição | País | Valor (R$) | Categoria | Tipo

Ative diagnóstico: PDF_EXTRACT_DEBUG=1 (imprime linhas descartadas no stdout).
"""
import io
import os
import re
import logging
from datetime import datetime
from typing import Optional

import pdfplumber

log = logging.getLogger(__name__)

DEBUG = os.getenv("PDF_EXTRACT_DEBUG", "").lower() in ("1", "true", "yes")

COLUMNS = ["Titular", "Data", "Descrição", "País", "Valor (R$)", "Categoria", "Tipo"]

DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{2,4}$")
VALOR_RE = re.compile(
    r"^[\s()\-]*(?:R\$\s*)?[\d]{1,3}(?:\.[\d]{3})*,[\d]{2}[\s)]*$|"
    r"^[\s()\-]*[\d]+,[\d]{2}[\s)]*$"
)

HEADER_MARKERS = ("titular", "data", "descri", "descrição", "país", "pais", "valor", "categoria", "tipo")
SKIP_ROW_KEYWORDS = ("lançamentos", "lancamentos", "fatura bb", "saldo anterior", "total da fatura", "pagamento")


def _debug(msg: str) -> None:
    if DEBUG:
        print(f"[pdf_extractor] {msg}")


def _parse_date_from_header(text: str) -> Optional[datetime]:
    match = re.search(r"(\d{2}/\d{2}/\d{4})", text)
    if match:
        try:
            return datetime.strptime(match.group(1), "%d/%m/%Y")
        except ValueError:
            pass
    return None


def _clean_value(val) -> str:
    if val is None:
        return ""
    return str(val).strip().replace("\n", " ")


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", s.lower().strip())


def _looks_like_money(val: str) -> bool:
    v = _clean_value(val)
    if not v or v in ("-", "—"):
        return False
    return bool(VALOR_RE.match(v)) or bool(re.search(r",\d{2}\)?$", v))


def _looks_like_date(val: str) -> bool:
    return bool(DATE_RE.match(_clean_value(val)))


def _is_table_header_row(row: list) -> bool:
    """Cabeçalho de colunas (não linha de transação)."""
    cells = [_norm(_clean_value(c)) for c in row if _clean_value(c)]
    if not cells:
        return False
    hits = sum(1 for c in cells if any(m in c for m in HEADER_MARKERS))
    return hits >= 3


def _is_summary_row(row_text: str) -> bool:
    low = _norm(row_text)
    return any(kw in low for kw in SKIP_ROW_KEYWORDS)


def _detect_column_map(header_row: list) -> Optional[dict[str, int]]:
    mapping: dict[str, int] = {}
    for i, cell in enumerate(header_row):
        key = _norm(_clean_value(cell))
        if not key:
            continue
        if "titular" in key:
            mapping["Titular"] = i
        elif key == "data" or key.startswith("data "):
            mapping["Data"] = i
        elif "descri" in key:
            mapping["Descrição"] = i
        elif "país" in key or key == "pais":
            mapping["País"] = i
        elif "valor" in key:
            mapping["Valor (R$)"] = i
        elif "categor" in key:
            mapping["Categoria"] = i
        elif key == "tipo":
            mapping["Tipo"] = i
    if len(mapping) >= 4:
        return mapping
    return None


def _row_from_column_map(row: list, col_map: dict[str, int]) -> dict:
    tx = {col: "" for col in COLUMNS}
    for col, idx in col_map.items():
        if idx < len(row):
            tx[col] = _clean_value(row[idx])
    return tx


def _guess_transaction_from_row(row: list, last_titular: str) -> Optional[dict]:
    """Mapeia colunas por heurística quando não há cabeçalho detectado."""
    cells = [_clean_value(c) for c in row]
    while cells and not cells[0]:
        cells.pop(0)

    if len(cells) < 4:
        return None

    date_idx = next((i for i, c in enumerate(cells) if _looks_like_date(c)), None)
    if date_idx is None:
        return None

    valor_idx = next(
        (i for i in range(len(cells) - 1, -1, -1) if _looks_like_money(cells[i])),
        None,
    )
    if valor_idx is None:
        return None

    titular = last_titular
    if date_idx > 0:
        titular = " ".join(cells[:date_idx]).strip() or last_titular

    data = cells[date_idx]
    middle = cells[date_idx + 1 : valor_idx]
    if not middle:
        return None

    país = ""
    categoria = ""
    tipo = ""
    descrição_parts = list(middle)

    if len(middle) >= 3:
        if len(middle[-1]) <= 20 and not _looks_like_money(middle[-1]):
            tipo = middle[-1]
            descrição_parts = middle[:-1]
        if len(descrição_parts) >= 2 and len(descrição_parts[-1]) <= 4:
            país = descrição_parts[-1]
            descrição_parts = descrição_parts[:-1]
        if len(descrição_parts) >= 2 and len(descrição_parts[-1]) <= 30:
            categoria = descrição_parts[-1]
            descrição_parts = descrição_parts[:-1]

    descrição = " ".join(descrição_parts).strip()
    valor = cells[valor_idx]

    if not descrição and not titular:
        return None

    return {
        "Titular": titular,
        "Data": data,
        "Descrição": descrição,
        "País": país,
        "Valor (R$)": valor,
        "Categoria": categoria,
        "Tipo": tipo,
    }


def _is_valid_transaction(tx: dict) -> bool:
    """Aceita linha com data + valor; titular ou descrição."""
    if not _looks_like_date(tx.get("Data", "")):
        return False
    if not _looks_like_money(tx.get("Valor (R$)", "")):
        return False
    return bool(tx.get("Descrição") or tx.get("Titular"))


def _legacy_row_to_tx(row: list) -> Optional[dict]:
    if len(row) >= 8:
        return {
            "Titular": _clean_value(row[1]),
            "Data": _clean_value(row[2]),
            "Descrição": _clean_value(row[3]),
            "País": _clean_value(row[4]),
            "Valor (R$)": _clean_value(row[5]),
            "Categoria": _clean_value(row[6]),
            "Tipo": _clean_value(row[7]),
        }
    if len(row) >= 7:
        return {
            "Titular": _clean_value(row[0]),
            "Data": _clean_value(row[1]),
            "Descrição": _clean_value(row[2]),
            "País": _clean_value(row[3]),
            "Valor (R$)": _clean_value(row[4]),
            "Categoria": _clean_value(row[5]),
            "Tipo": _clean_value(row[6]),
        }
    return None


def _parse_text_lines(text: str, last_titular: str) -> tuple[list[dict], str]:
    """Fallback cirúrgico: extrai transações guardando o titular em memória."""
    found: list[dict] = []
    titular = last_titular

    date_re = re.compile(r"(\d{2}/\d{2}(?:/\d{2,4})?)")
    valor_re = re.compile(r"([\d.]*,\d{2})")

    # Padrão de cabeçalho de bloco de cartão:
    #   "FABIO R POVOA (Cartão 6688)"  → grupo 1 = nome, grupo 2 = número
    #   "01- FABIO R POVOA Cartao N. 6688"
    _card_header_re = re.compile(
        r"(?:\d+[-–]\s*)?([A-ZÁÉÍÓÚÃÕÂÊÎÔÛÀÜ][A-ZÁÉÍÓÚÃÕÂÊÎÔÛÀÜ\s]+?)"
        r"\s*(?:\(Cartão|Cartao N\.)\s*(\d{4})\)?",
        re.IGNORECASE,
    )

    for raw_line in text.splitlines():
        line = _clean_value(raw_line)

        if not line or _is_summary_row(line) or _is_table_header_row([line]):
            continue

        # 1. Busca Data e Valor
        date_match = date_re.search(line)
        valor_match = valor_re.search(line)

        # SE NÃO TEM DATA NEM VALOR: tenta detectar cabeçalho de bloco de cartão
        if not date_match and not valor_match:
            # Cabeçalho explícito de cartão: "NOME (Cartão NNNN)" ou "NN- NOME Cartao N. NNNN"
            m = _card_header_re.search(line)
            if m:
                titular = m.group(1).strip().upper()
                _debug(f"[titular via cartão] {line!r} → {titular!r}")
            elif len(line) > 4 and line.isupper() and not any(k in line.lower() for k in ["fatura", "ourocard", "banco"]):
                # Linha toda maiúscula (formato legado)
                titular = line.strip()
            continue

        # Se chegou aqui, é uma linha de transação (precisa ter data e valor)
        if not date_match or not valor_match:
            continue

        data_str = date_match.group(1)
        valor_str = valor_match.group(1)

        # 2. Identifica o Titular da linha atual
        # Se a linha NÃO começa com a data, limpa o texto de antes para ver se há um novo titular ali
        if not line.startswith(data_str):
            partes_antes = line.split(data_str)[0].strip()
            if len(partes_antes) > 4 and partes_antes.isupper() and not any(k in partes_antes.lower() for k in ["fatura", "ourocard"]):
                titular = partes_antes
        
        linha_titular = titular

        # 3. Isola a Descrição comercial
        if data_str in line:
            resto_linha = line.split(data_str)[1].strip()
        else:
            resto_linha = line.replace(data_str, "").strip()
            
        desc_clean = resto_linha.replace(valor_str, "").replace(" BR ", " ").replace(" BR", "").strip()
        desc_clean = re.sub(r'\s+', ' ', desc_clean)

        # 4. Define o Tipo
        tipo_tx = "Compra"
        if "parcela" in line.lower():
            tipo_tx = "Parcela"
        elif "anuidade" in line.lower():
            tipo_tx = "Anuidade"
        elif "estorno" in line.lower() or "credito" in line.lower() or "crédito" in line.lower():
            tipo_tx = "Estorno/Crédito"

        tx = {
            "Titular": linha_titular if linha_titular else "TITULAR NÃO IDENTIFICADO",
            "Data": data_str,
            "Descrição": desc_clean if desc_clean else "Lançamento Fatura",
            "País": "BR",
            "Valor (R$)": valor_str,
            "Categoria": "Outros",
            "Tipo": tipo_tx
        }

        found.append(tx)

    return found, titular


def _extract_all_tables(page) -> list[list[list]]:
    """Tenta várias estratégias de tabela do pdfplumber."""
    seen: set[str] = set()
    tables: list[list[list]] = []

    strategies = [
        {},
        {"vertical_strategy": "lines", "horizontal_strategy": "lines"},
        {"vertical_strategy": "text", "horizontal_strategy": "text"},
    ]

    for settings in strategies:
        try:
            chunk = page.extract_tables(table_settings=settings) if settings else page.extract_tables()
        except Exception:
            chunk = None
        if not chunk:
            continue
        for table in chunk:
            if not table:
                continue
            key = repr(table[:3])
            if key in seen:
                continue
            seen.add(key)
            tables.append(table)

    return tables


def _process_table_rows(
    table: list[list],
    *,
    discard_log: list[str],
) -> list[dict]:
    transactions: list[dict] = []
    col_map: Optional[dict[str, int]] = None
    last_titular = ""

    for row in table:
        if not row:
            discard_log.append(f"(vazia) {row!r}")
            continue

        row_text = " ".join(_clean_value(c) for c in row if c)
        if not row_text.strip():
            discard_log.append("(só vazios)")
            continue

        if _is_summary_row(row_text):
            discard_log.append(f"[resumo] {row!r}")
            continue

        if _is_table_header_row(row):
            col_map = _detect_column_map(row) or col_map
            discard_log.append(f"[cabeçalho] {row!r}")
            continue

        non_empty = [c for c in row if c and str(c).strip()]
        if len(non_empty) < 3:
            discard_log.append(f"[poucas colunas] {row!r}")
            continue

        tx = None
        if col_map:
            tx = _row_from_column_map(row, col_map)
        if not tx or not tx.get("Data"):
            tx = _legacy_row_to_tx(row)
        if not tx or not _looks_like_date(tx.get("Data", "")):
            tx = _guess_transaction_from_row(row, last_titular)

        if not tx:
            discard_log.append(f"[não mapeada] {row!r}")
            continue

        if tx.get("Titular"):
            last_titular = tx["Titular"]
        elif last_titular:
            tx["Titular"] = last_titular

        if _is_valid_transaction(tx):
            transactions.append(tx)
        else:
            discard_log.append(
                f"[filtro] titular={tx.get('Titular')!r} data={tx.get('Data')!r} "
                f"desc={tx.get('Descrição')!r} valor={tx.get('Valor (R$)')!r} | raw={row!r}"
            )

    return transactions


def extract_transactions(pdf_bytes: bytes) -> tuple[list[dict], Optional[datetime]]:
    """
    Extrai transações e data de referência do PDF.

    Returns:
        (transactions, ref_date)
    """
    transactions: list[dict] = []
    ref_date = None
    header_text = ""
    discard_log: list[str] = []
    seen_keys: set[tuple] = set()

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        last_titular = ""

        for page_num, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            header_text += text

            page_tx: list[dict] = []
            for table in _extract_all_tables(page):
                page_tx.extend(_process_table_rows(table, discard_log=discard_log))

            if not page_tx and text:
                text_tx, last_titular = _parse_text_lines(text, last_titular)
                page_tx.extend(text_tx)
                if DEBUG and text_tx:
                    _debug(f"página {page_num}: {len(text_tx)} transações via texto")

            for tx in page_tx:
                key = (tx.get("Data"), tx.get("Descrição"), tx.get("Valor (R$)"))
                if key in seen_keys:
                    continue
                seen_keys.add(key)
                transactions.append(tx)

    ref_date = _parse_date_from_header(header_text)

    if not transactions:
        log.warning(
            "Nenhuma transação extraída do PDF (%d linhas descartadas). "
            "Defina PDF_EXTRACT_DEBUG=1 para ver detalhes no stdout.",
            len(discard_log),
        )
        _dump_discard_sample(discard_log, header_text)
    elif DEBUG:
        _debug(f"extraídas {len(transactions)} transações; descartadas {len(discard_log)}")

    return transactions, ref_date


def _dump_discard_sample(discard_log: list[str], header_text: str, limit: int = 40) -> None:
    """Diagnóstico: log sempre; stdout com amostra (ou dump completo com PDF_EXTRACT_DEBUG=1)."""
    sample = discard_log[: (limit if DEBUG else 25)]
    for line in sample[:15]:
        log.info("PDF descartado: %s", line)

    print("\n=== PDF extract: linhas descartadas (amostra) ===")
    for line in sample:
        print(line)
    if len(discard_log) > len(sample):
        print(f"... e mais {len(discard_log) - len(sample)} linhas")

    if DEBUG:
        print("\n=== PDF extract: trecho do texto bruto ===")
        for i, line in enumerate(header_text.splitlines()):
            low = line.lower()
            if "lanç" in low or "fatura bb" in low or re.search(r"\d{2}/\d{2}/\d{4}", line):
                print(line[:200])
            if i > 120:
                break
    print("=== fim amostra (use PDF_EXTRACT_DEBUG=1 para mais detalhe) ===\n")


def get_fatura_title(header_text: str) -> str:
    match = re.search(r"Fatura BB[^\n]+", header_text)
    if match:
        return match.group(0).strip()
    return "Fatura BB"


def extract_header_text(pdf_bytes: bytes) -> str:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages[:2])
