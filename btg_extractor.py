"""
Extrai transações de uma fatura BTG Pactual (PDF sem senha).

Estrutura do PDF BTG:
  - Cada página tem cabeçalho: "Nome | fatura de Mês Ano | cartão final XXXX"
  - Blocos de titular delimitados por:
      "Lançamentos do cartão físico | Nome | Final XXXX"
      "Lançamentos do cartão adicional | Nome | Final XXXX"
  - Transações em duas colunas por linha:
      "DD Mmm Descrição (X/Y) R$ valor   DD Mmm Descrição R$ valor"
  - Ignorar: pagamentos, benefícios/cashback, totais, créditos

Retorna: dict {tab_name: [{"Data", "Descrição", "Parcela", "Valor (R$)", "Tipo"}]}
"""
import io
import re
import logging
from datetime import datetime
from typing import Optional

import pdfplumber

log = logging.getLogger(__name__)

# Meses PT abreviados
MESES = r"(?:Jan|Fev|Mar|Abr|Mai|Jun|Jul|Ago|Set|Out|Nov|Dez)"

# Padrão de início de transação: "DD Mmm "
TX_START = re.compile(rf"\b(\d{{2}})\s+({MESES})\s+", re.IGNORECASE)

# Padrão de valor: "R$ 1.234,56" com sinal opcional
VALOR_RE = re.compile(r"R\$\s*([\d.,]+)")

# Padrão de parcela: "(3/10)"
PARCELA_RE = re.compile(r"\((\d+/\d+)\)")

# Cabeçalho de bloco de titular
BLOCK_RE = re.compile(
    r"Lançamentos do cartão (?:físico|adicional)\s*\|\s*(.+?)\s*\|\s*Final\s+(\d{4})",
    re.IGNORECASE,
)

# Linhas a ignorar
SKIP_PATTERNS = [
    re.compile(p, re.IGNORECASE) for p in [
        r"^Total de",
        r"^Pagamento de fatura",
        r"^Benefício do cartão",
        r"^Cashback",
        r"^Cancelamento",
        r"^BTG Pactual",
        r"^CNPJ",
        r"^Av\.",
        r"^São Paulo",
        r"^Resumo",
        r"^Lançamentos",
        r"^Período",
        r"^Vencimento",
        r"^Pagamento mínimo",
        r"^Fatura",
        r"^Olá",
        r"^Mensalidade",
        r"^Desconto",
        r"^Valor a pagar",
        r"^Saldo",
        r"^Total da fatura",
        r"^Total a pagar",
        r"^Seguros",
        r"^Outros",
        r"^Sua fatura",
        r"^Se preferir",
        r"^Para ",
        r"^Cartão >",
        r"^Sabia",
        r"^Boa notícia",
        r"^\+$",
        r"^=$",
        r"FATURA ATUAL",
        r"Pagamentos feitos pelo cliente",
    ]
]


def _should_skip(line: str) -> bool:
    s = line.strip()
    if not s:
        return True
    for pat in SKIP_PATTERNS:
        if pat.search(s):
            return True
    return False


def _parse_tx_from_chunk(chunk: str, date_str: str) -> Optional[dict]:
    """
    Dado um fragmento de texto após a data (ex: 'Autoglass (4/5) R$ 574,20'),
    extrai descrição, parcela e valor.
    """
    chunk = chunk.strip()
    valor_m = VALOR_RE.search(chunk)
    if not valor_m:
        return None

    valor = valor_m.group(1).strip()
    # Remove o "R$ valor" do chunk para isolar a descrição
    desc_part = chunk[:valor_m.start()].strip()

    # Extrai parcela se existir
    parcela_m = PARCELA_RE.search(desc_part)
    parcela = ""
    if parcela_m:
        parcela = parcela_m.group(1)
        desc_part = desc_part[:parcela_m.start()].strip()

    if not desc_part:
        return None

    tipo = "Parcela" if parcela else "Compra"

    return {
        "Data": date_str,
        "Descrição": desc_part,
        "Parcela": parcela,
        "Valor (R$)": valor,
        "Tipo": tipo,
    }


def _extract_transactions_from_line(line: str) -> list[dict]:
    """
    Uma linha BTG pode conter 1 ou 2 transações lado a lado.
    Divide pelos starts de data e extrai cada uma.
    """
    results = []
    matches = list(TX_START.finditer(line))
    if not matches:
        return results

    for i, m in enumerate(matches):
        date_str = f"{m.group(1)}/{m.group(2)}"
        # Texto até o próximo match (ou fim da linha)
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(line)
        chunk = line[start:end]

        tx = _parse_tx_from_chunk(chunk, date_str)
        if tx:
            # Filtra créditos/benefícios que aparecem como transações com data
            desc = tx["Descrição"].lower()
            if any(k in desc for k in ["benefício do cartão", "beneficio do cartao", "cashback"]):
                continue
            results.append(tx)

    return results


def _detect_block_header(line: str) -> Optional[tuple[str, str]]:
    """Retorna (nome, final_cartao) se a linha for um cabeçalho de bloco."""
    m = BLOCK_RE.search(line)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return None


def _name_to_tab(name: str, card_final: str, tab_names: dict) -> str:
    """
    Mapeia nome + final do cartão para nome da aba.
    tab_names é um dict acumulado para detectar duplicatas (ex: Fátima com 2 cartões).
    """
    name_upper = name.upper()

    if "LUIZ" in name_upper:
        return "Luiz"
    if "HIERO" in name_upper:
        return "Hiero"
    if "FATIMA" in name_upper or "FÁTIMA" in name_upper or "MARIA DE FATIMA" in name_upper:
        return f"Fátima {card_final}"
    # Fallback: usa o próprio nome
    return name.split()[0].capitalize()


def extract_btg_transactions(pdf_bytes: bytes) -> tuple[dict[str, list[dict]], Optional[datetime]]:
    """
    Extrai transações do PDF BTG agrupadas por titular.

    Returns:
        (grouped, ref_date)
        grouped: {tab_name: [transações]}
        ref_date: data de referência extraída do cabeçalho
    """
    grouped: dict[str, list[dict]] = {}
    ref_date: Optional[datetime] = None
    current_tab: Optional[str] = None
    tab_names: dict[str, str] = {}  # tab_name -> card_final

    # Regex para extrair data de referência do cabeçalho
    header_date_re = re.compile(
        r"fatura de (\w+) de (\d{4})", re.IGNORECASE
    )
    meses_map = {
        "janeiro":1,"fevereiro":2,"março":3,"abril":4,"maio":5,"junho":6,
        "julho":7,"agosto":8,"setembro":9,"outubro":10,"novembro":11,"dezembro":12,
    }

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""

            for line in text.splitlines():
                line = line.strip()
                if not line:
                    continue

                # Extrai data de referência do cabeçalho da página
                if ref_date is None:
                    hm = header_date_re.search(line)
                    if hm:
                        mes_str = hm.group(1).lower()
                        ano = int(hm.group(2))
                        mes = meses_map.get(mes_str)
                        if mes:
                            ref_date = datetime(ano, mes, 1)

                # Detecta cabeçalho de bloco de titular
                block = _detect_block_header(line)
                if block:
                    nome, final = block
                    tab = _name_to_tab(nome, final, tab_names)
                    if tab not in grouped:
                        grouped[tab] = []
                        tab_names[tab.split()[0]] = final  # registra para detectar duplicatas
                    current_tab = tab
                    log.info("📌 Bloco BTG: %s (Final %s) → aba '%s'", nome, final, tab)
                    continue

                # Ignora linhas de controle
                if _should_skip(line):
                    continue

                # Tenta extrair transações da linha
                if current_tab is not None:
                    txs = _extract_transactions_from_line(line)
                    if txs:
                        grouped[current_tab].extend(txs)

    totals = {k: len(v) for k, v in grouped.items()}
    log.info("BTG — extração concluída: %s", totals)
    return grouped, ref_date


def get_btg_ref_date(pdf_bytes: bytes) -> Optional[datetime]:
    """Extrai só a data de referência do cabeçalho."""
    _, ref_date = extract_btg_transactions(pdf_bytes)
    return ref_date
