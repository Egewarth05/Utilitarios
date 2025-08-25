import io
import re
import fitz  # PyMuPDF
import pdfplumber
import unicodedata
from PIL import Image
import pytesseract
from decimal import Decimal, InvalidOperation
from typing import List, Dict, Optional, Tuple

# palavras-gatilho que marcam início de descrição no meio da linha
KNOWN_STARTS = [
    "TAXA DE INTERMEDIACAO",
    "CREDITO REF",
    "DEBITO CBLC IRRF",
    "OPERACOES EM BOLSA LIQ",
    "DIVIDENDOS DE CLIENTES",
    "CREDITO DE REEMBOLSO",
]

# detecta um SEGUNDO par de datas na mesma linha (quebra em novo lançamento)
DATEPAIR_RE = re.compile(r'\d{2}/\d{2}/\d{4}\s+\d{2}/\d{2}/\d{4}')

# fallback para quando as 2 datas vêm quebradas em LINHAS diferentes
ONE_DATE_ONLY_RE = re.compile(r'^\s*(\d{2}/\d{2}/\d{4})\s*$')
NEXT_DATE_AND_TEXT_RE = re.compile(r'^\s*(\d{2}/\d{2}/\d{4})(?:\s+(.*))?$')

DECIMAL_RE = re.compile(r'([-\u2212]?)R\$\s*([\d\.\,\u00A0]+)')  # aceita NBSP e sinal U+2212
HEADER_RE = re.compile(r'^\s*Liq\s+Mov\s+Histórico\s+Valor\s+Saldo', re.IGNORECASE)
FUTUROS_RE = re.compile(r'^\s*Lançamentos futuros', re.IGNORECASE)
# Documento no formato "Nº 116011339" (ou "No 116011339")
DOC_NO_RE = re.compile(r'^(?:N[ºo]\s*)?(\d{6,})$', re.IGNORECASE)
DOC_NO_PURE_RE = re.compile(r'^(?:N[ºo]\s*)(\d{6,})$', re.IGNORECASE)
DOC_NO_WITH_TEXT_RE = re.compile(r'^(?:N[ºo]\s*)(\d{6,})\s+(.+)$', re.IGNORECASE)
DATE_LINE_RE = re.compile(
    r'^\s*(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})(?:\s+(.*))?$'
)

# ---------------------------
# OCR helpers (usado só quando a página não tem texto)
# ---------------------------
def _ocr_page_to_lines(fitz_page: fitz.Page, dpi: int = 300, lang: Optional[str] = None) -> List[str]:
    """Renderiza a página como imagem e roda OCR, retornando as linhas de texto."""
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    pix = fitz_page.get_pixmap(matrix=mat, alpha=False)
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    txt = pytesseract.image_to_string(img, lang=lang) if lang else pytesseract.image_to_string(img)
    # normaliza quebras de linha
    lines = [ln.rstrip() for ln in txt.splitlines()]
    # remove linhas vazias excessivas
    return [l for l in lines if l.strip()]

# ---------------------------

def _parse_doc_line(line: str) -> tuple[Optional[str], Optional[str]]:
    """
    Retorna (doc_id, texto_pos_doc) se a linha for:
      - "Nº 123456 ..."  -> (123456, "...")
      - "Nº 123456"      -> (123456, None)
      - caso contrário   -> (None, None)
    """
    s = line.strip()
    m2 = DOC_NO_WITH_TEXT_RE.match(s)
    if m2:
        return m2.group(1), _clean_spaces(m2.group(2))
    m1 = DOC_NO_PURE_RE.match(s)
    if m1:
        return m1.group(1), None
    return None, None

def _is_doc_no_line(line: str) -> Optional[str]:
    m = DOC_NO_RE.match(line.strip())
    return m.group(1) if m else None

def _to_decimal_br(valor_str: str) -> Decimal:
    s = (valor_str or "").replace('\u00A0','').replace('.', '').replace(',', '.')
    try:
        return Decimal(s)
    except InvalidOperation:
        raise ValueError(f"Valor inválido: {valor_str!r}")

def _clean_spaces(s: str) -> str:
    return re.sub(r'\s+', ' ', s).strip()

def _strip_trailing_amounts(s: str) -> str:
    return re.sub(r'(\s*-?\s*R\$\s*[\d\.\,]+\s*){1,2}$', '', s).rstrip(' .,-')

def _is_big_doc_code(line: str) -> bool:
    return bool(re.fullmatch(r'\d{10,}(?:-\d+)?', line.strip()))

def _norm(s: str) -> str:
    s = s or ""
    s = unicodedata.normalize('NFKD', s)
    s = ''.join(ch for ch in s if not unicodedata.combining(ch))
    return s.upper()

def _is_header_line(s: str) -> bool:
    z = _norm(s)
    return all(k in z for k in ["LIQ","MOV","HISTORICO","VALOR","SALDO"])

def _classificar_conta_historico(descricao: str):
    d = _norm(descricao)
    if "TAXA DE INTERMEDIACAO" in d:
        return "4698","32","25016"
    if "DEBITO CBLC IRRF S/ RENDIMENTO" in d:
        return "4698","32","25016"
    if "CREDITO REF" in d and "REMUNERACAO" in d:
        return "25016","21","25017"
    if "OPERACOES EM BOLSA LIQ" in d:
        return "25016","21","25017"
    if "DIVIDENDOS DE CLIENTES" in d:
        return "25016","21","25017"
    if "CREDITO DE REEMBOLSO DE EVENTO" in d or "CREDITO DE REEMBOLSO" in d:
        return "25016","21","25017"
    return None, None, None

def _page_to_lines(page) -> list[str]:
    words = page.extract_words(
        x_tolerance=2,
        y_tolerance=4,          # << antes era 2
        keep_blank_chars=False,
        use_text_flow=True
    )
    lines, row, last_top = [], [], None
    for w in words:
        t = round(w["top"], 1)
        if last_top is None or abs(t - last_top) <= 3.5:   # << antes era <= 2
            row.append((w["x0"], w["text"]))
        else:
            row.sort(key=lambda z: z[0])
            lines.append(" ".join(z[1] for z in row))
            row = [(w["x0"], w["text"])]
        last_top = t
    if row:
        row.sort(key=lambda z: z[0])
        lines.append(" ".join(z[1] for z in row))
    return lines

def _normalize_with_map(s: str) -> tuple[str, list[int]]:
    """Normaliza (sem acentos, maiúsc.) e devolve o mapa de posições norm->orig."""
    norm_chars = []
    pos_map = []
    for i, ch in enumerate(s):
        base = unicodedata.normalize('NFKD', ch)
        base = ''.join(c for c in base if not c.isascii() and unicodedata.combining(c)) or base
        base = ''.join(c for c in base if not unicodedata.combining(c))
        if not base:
            continue
        for c in base:
            norm_chars.append(c.upper())
            pos_map.append(i)
    return ''.join(norm_chars), pos_map

def _split_on_known_starts(line: str) -> list[str]:
    """Divide a linha em várias quando encontra uma sentença conhecida ou um novo par de datas."""
    if not line.strip():
        return []
    norm, pos_map = _normalize_with_map(line)
    cut_norm_pos = set()

    # 1) sentenças conhecidas
    for needle in KNOWN_STARTS:
        start = 0
        while True:
            idx = norm.find(needle, start)
            if idx == -1:
                break
            if idx > 0:                 # não corta se for no início da linha
                cut_norm_pos.add(idx)
            start = idx + 1

    # 2) pares de datas extras no meio
    for m in DATEPAIR_RE.finditer(norm):
        if m.start() > 0:
            cut_norm_pos.add(m.start())

    if not cut_norm_pos:
        return [line.strip()]

    # converte posição normalizada -> posição no original
    cut_pos = sorted({pos_map[p] for p in cut_norm_pos if p < len(pos_map)})

    parts, start = [], 0
    for p in cut_pos:
        seg = line[start:p].strip()
        if seg:
            parts.append(seg)
        start = p
    last = line[start:].strip()
    if last:
        parts.append(last)
    return parts

def _explode_lines(lines: list[str]) -> list[str]:
    out = []
    for l in lines:
        out.extend(_split_on_known_starts(l))
    return out

def parse_xp_extrato_pdf(pdf_path: str, *, ocr_dpi: int = 300, ocr_lang: Optional[str] = "por") -> Tuple[List[Dict], Dict]:
    """
    Lê o PDF de extrato.
    - Usa pdfplumber normalmente (sem alterar seu comportamento atual).
    - Se a página não tiver texto (imagem/digitalizada), usa OCR (pytesseract) naquela página.
    Retorna (rows, meta) onde meta indica se houve OCR e em quais páginas.
    """
    rows: List[Dict] = []
    in_table = False
    current = None

    pending_desc = ""
    candidate_desc = ""
    first_date_only: Optional[str] = None

    fitz_doc = fitz.open(pdf_path)
    paginas_ocr: List[int] = []

    with pdfplumber.open(pdf_path) as pdf:
        for idx, page in enumerate(pdf.pages):
            # garante flush do lançamento anterior ao mudar de página
            if current:
                if not current.get('descricao') and candidate_desc:
                    current['descricao'] = candidate_desc
                rows.append(current)
                current = None

            # --- zera o contexto por página ---
            in_table = False
            pending_desc = ""
            candidate_desc = ""
            first_date_only = None

            # 1) tentativa padrão (mantém o que você já fazia)
            raw_lines = _page_to_lines(page)

            # 2) fallback: extract_text()
            if not raw_lines:
                text = page.extract_text() or ""
                raw_lines = [l.rstrip() for l in text.splitlines() if l.strip()]

            # 3) fallback final: OCR somente se ainda não houver nada
            if not raw_lines:
                fpage = fitz_doc.load_page(idx)
                raw_lines = _ocr_page_to_lines(fpage, dpi=ocr_dpi, lang=ocr_lang)
                paginas_ocr.append(idx + 1)  # páginas 1-based

            lines = _explode_lines(raw_lines)

            for raw_line in lines:
                line = raw_line.strip()
                if not line:
                    continue

                if first_date_only is None:
                    m_one = ONE_DATE_ONLY_RE.match(line)
                    if m_one:
                        first_date_only = m_one.group(1)
                        continue
                else:
                    m_next = NEXT_DATE_AND_TEXT_RE.match(line)
                    if m_next:
                        segunda = m_next.group(1)
                        resto = m_next.group(2) or ""
                        line = f"{first_date_only} {segunda} {resto}".strip()
                        first_date_only = None
                    else:
                        # a primeira "data solta" não formou par — trate como texto pendente
                        pending_desc = _clean_spaces(((pending_desc + " " + first_date_only).strip()) if pending_desc else first_date_only)
                        first_date_only = None

                m_dates = DATE_LINE_RE.match(line)  # <- sempre calcule aqui

                # 0) Entrar na tabela
                if not in_table:
                    # entra na tabela se ver o header OU já ver a primeira linha com datas
                    if HEADER_RE.search(line) or _is_header_line(line) or m_dates:
                        in_table = True
                        # se NÃO for uma linha de datas (ex.: cabeçalho), vai pra próxima
                        if not m_dates:
                            continue
                    else:
                        continue

                # 1) Fim da tabela
                if FUTUROS_RE.search(line):
                    # garante que "data solta" não se perca ao fechar a seção
                    if first_date_only:
                        pending_desc = _clean_spaces(
                            ((pending_desc + " " + first_date_only).strip()) if pending_desc else first_date_only
                        )
                        first_date_only = None

                    if current:
                        if not current.get('descricao') and candidate_desc:
                            current['descricao'] = candidate_desc
                        rows.append(current)
                    in_table = False
                    current = None
                    pending_desc = ""
                    candidate_desc = ""
                    continue

                # 2) Linha com datas (abre/fecha lançamento)
                if m_dates:
                    if first_date_only:
                        pending_desc = _clean_spaces(
                            ((pending_desc + " " + first_date_only).strip()) if pending_desc else first_date_only
                        )
                        first_date_only = None

                    if current:
                        if not current.get('descricao') and candidate_desc:
                            current['descricao'] = candidate_desc
                        rows.append(current)

                    # abre novo lançamento
                    desc_inline = _strip_trailing_amounts(_clean_spaces(m_dates.group(3) or ""))
                    current = {
                        'data_liq': m_dates.group(1),
                        'data_mov': m_dates.group(2),
                        'descricao': desc_inline if desc_inline else None,
                        'documento': None,
                        'valor': None,   # só valor (sem saldo)
                        'tipo': None,
                        'conta': None,
                        'historico_code': None,
                        'contrapartida': None,
                    }

                    # valor pode (ou não) estar na mesma linha
                    m_val_here = DECIMAL_RE.search(line)
                    if m_val_here:
                        sign_v, v = m_val_here.groups()
                        valor = _to_decimal_br(v)
                        if sign_v in ('-', '\u2212'):
                            valor = -valor
                        current['valor'] = valor
                        current['tipo'] = 'D' if valor < 0 else 'C'

                    # descrição/categorização
                    candidate_desc = current['descricao'] or ""
                    if not candidate_desc and pending_desc:
                        candidate_desc = pending_desc
                        current['descricao'] = pending_desc

                    if current.get('descricao'):
                        conta, hist, cp = _classificar_conta_historico(current['descricao'])
                        current['conta'], current['historico_code'], current['contrapartida'] = conta, hist, cp

                    pending_desc = ""
                    continue

                # 3) Ainda não começou nenhum lançamento: acumula descrição pré-datas
                if current is None:
                    doc_id, after = _parse_doc_line(line)
                    if doc_id:
                        if after:
                            pending_desc = _clean_spaces((pending_desc + " " + after).strip()) if pending_desc else after
                        continue
                    if DECIMAL_RE.search(line) or _is_big_doc_code(line):
                        continue
                    pending_desc = _clean_spaces((pending_desc + " " + line).strip()) if pending_desc else line
                    continue

                # 4.1) Documento “Nº 123 …”
                doc_id, after = _parse_doc_line(line)
                if doc_id:
                    current['documento'] = doc_id
                    if after and not current.get('descricao'):
                        current['descricao'] = after
                        candidate_desc = after
                        conta, hist, cp = _classificar_conta_historico(current['descricao'])
                        current['conta'], current['historico_code'], current['contrapartida'] = conta, hist, cp
                    continue

                # 4.2) Documento longo 2025…-1
                if _is_big_doc_code(line):
                    current['documento'] = line.strip()
                    continue

                m_val = DECIMAL_RE.search(line)
                if m_val and current.get('valor') is None:
                    sign_v, v = m_val.groups()
                    valor = _to_decimal_br(v)
                    if sign_v in ('-', '\u2212'):
                        valor = -valor
                    current['valor'] = valor
                    current['tipo'] = 'D' if valor < 0 else 'C'

                    # tentar extrair a descrição desta linha de valor
                    if not current.get('descricao'):
                        maybe_desc = _clean_spaces(_strip_trailing_amounts(line))
                        # ignora se sobrou apenas datas (ou nada)
                        if maybe_desc and not re.fullmatch(r'\d{2}/\d{2}/\d{4}(?:\s+\d{2}/\d{2}/\d{4})?', maybe_desc):
                            current['descricao'] = maybe_desc
                            candidate_desc = maybe_desc
                            conta, hist, cp = _classificar_conta_historico(maybe_desc)
                            current['conta'], current['historico_code'], current['contrapartida'] = conta, hist, cp
                    continue

                # 4.4) Texto comum -> descrição (concatena)
                if not current.get('descricao'):
                    current['descricao'] = _clean_spaces(line)
                    candidate_desc = current['descricao']
                    conta, hist, cp = _classificar_conta_historico(current['descricao'])
                    current['conta'], current['historico_code'], current['contrapartida'] = conta, hist, cp
                else:
                    current['descricao'] = _clean_spaces(current['descricao'] + ' ' + line)

    fitz_doc.close()

    if current:
        if not current.get('descricao') and candidate_desc:
            current['descricao'] = candidate_desc
        rows.append(current)

    rows = [r for r in rows if (r.get('descricao') or r.get('documento') or r.get('valor') is not None)]
    meta = {
        "usou_ocr": len(paginas_ocr) > 0,
        "paginas_ocr": paginas_ocr,
        "total_paginas": len(paginas_ocr) + (fitz.open(pdf_path).page_count - len(paginas_ocr)) if paginas_ocr else fitz.open(pdf_path).page_count
    }
    return rows, meta

def export_to_xlsx(rows: List[Dict], out_xlsx_path: str) -> None:
    import pandas as pd
    df = pd.DataFrame(rows)
    df = df.rename(columns={
        'descricao': 'Descrição',
        'historico_code': 'Histórico',
        'conta': 'Conta',
        'contrapartida': 'Contrapartida',
    })
    col_ordem = ['data_liq','data_mov','Descrição','documento','tipo','valor','Conta','Histórico','Contrapartida']
    for c in col_ordem:
        if c not in df.columns:
            df[c] = None
    df = df[col_ordem]
    df['valor'] = pd.to_numeric(df['valor'], errors='coerce')
    with pd.ExcelWriter(out_xlsx_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Extrato')

def export_to_txt_contabil(
    rows: List[Dict],
    out_txt_path: str,
    *,
    codigo_prefixo="1",   # continua 1
    codigo_meio="5",      # fallback se não houver contrapartida no lançamento
    codigo_cc="337"       # fallback se não houver histórico no lançamento
) -> None:
    from decimal import Decimal

    def _fmt_data_ddmmaa(dmy: str) -> str:
        return dmy.replace('/', '')

    def _fmt_valor_2p(v) -> str:
        try:
            v = Decimal(str(v))
        except Exception:
            pass
        return f"{abs(v):.2f}"

    with open(out_txt_path, 'w', encoding='utf-8') as f:
        for r in rows:
            if r.get('valor') is None:
                continue

            data = _fmt_data_ddmmaa(r['data_liq'])
            conta = (r.get('conta') or "").strip()
            contrapartida = (r.get('contrapartida') or codigo_meio)
            valor = _fmt_valor_2p(r['valor'])
            historico = (r.get('historico_code') or codigo_cc)
            hist_txt = (r.get('descricao') or '').replace('"', "'")
            linha = f'{codigo_prefixo},{data},{conta},{contrapartida},{valor},{historico},"{hist_txt}"'
            f.write(linha + '\n')

def processar_extrato_pdf(in_pdf_path: str, out_xlsx_path: str, out_txt_path: Optional[str]=None, config: Optional[Dict]=None) -> Dict:
    config = config or {}
    ocr_dpi = int(config.get("ocr_dpi", 300))
    ocr_lang = config.get("ocr_lang", "por")

    rows, meta = parse_xp_extrato_pdf(in_pdf_path, ocr_dpi=ocr_dpi, ocr_lang=ocr_lang)
    export_to_xlsx(rows, out_xlsx_path)
    if out_txt_path:
        export_to_txt_contabil(
            rows,
            out_txt_path,
            codigo_prefixo=config.get('codigo_prefixo', '1'),
            codigo_meio=config.get('codigo_meio', '5'),
            codigo_cc=config.get('codigo_cc', '337'),
        )

    resultado = {
        'quantidade_lancamentos': len(rows),
        'paginas_ocr': meta.get("paginas_ocr", []),
        'usou_ocr': meta.get("usou_ocr", False),
    }
    if meta.get("usou_ocr"):
        aviso = "⚠️ Atenção: OCR foi utilizado nas páginas {}. Confira os valores, pois o OCR pode confundir números (ex.: 0 ↔ O, , ↔ .).".format(
            ", ".join(map(str, meta["paginas_ocr"]))
        )
        resultado['aviso'] = aviso
        print(aviso)
    return resultado

def debug_dump_pdf(pdf_path: str, max_lines_per_page: int = 9999, *, ocr_lang: str = "por", ocr_dpi: int = 300) -> None:
    print("==== DEBUG XP EXTRATO ====")
    try:
        import pdfplumber
    except Exception as e:
        print("pdfplumber não disponível:", e)
        return

    # roda o parser (com OCR condicional) para obter rows + meta (inclusive páginas com OCR)
    try:
        rows_preview, meta = parse_xp_extrato_pdf(pdf_path, ocr_dpi=ocr_dpi, ocr_lang=ocr_lang)
    except Exception as e:
        print("ERRO rodando parse_xp_extrato_pdf:", e)
        rows_preview, meta = [], {"usou_ocr": False, "paginas_ocr": []}

    if meta.get("usou_ocr"):
        print(f"⚠️ AVISO GERAL: OCR foi utilizado neste arquivo nas páginas {', '.join(map(str, meta['paginas_ocr']))}.")
        print("   Confira os valores, pois podem haver erros de reconhecimento (0↔O, vírgula/ponto, etc.).")

    fitz_doc = fitz.open(pdf_path)
    total_rows = rows_preview

    with pdfplumber.open(pdf_path) as pdf:
        for pidx, page in enumerate(pdf.pages, start=1):
            # Se esta página consta como OCR pelo parser, já renderizamos direto via OCR
            if pidx in meta.get("paginas_ocr", []):
                used_ocr = True
                fpage = fitz_doc.load_page(pidx - 1)
                raw_lines = _ocr_page_to_lines(fpage, dpi=ocr_dpi, lang=ocr_lang)
            else:
                # tenta texto; se falhar, gera OCR (não deve acontecer se meta já tinha)
                raw_lines = _page_to_lines(page)
                used_ocr = False
                if not raw_lines:
                    text = page.extract_text() or ""
                    raw_lines = [l.rstrip() for l in text.splitlines()]
                if not raw_lines:
                    used_ocr = True
                    fpage = fitz_doc.load_page(pidx - 1)
                    raw_lines = _ocr_page_to_lines(fpage, dpi=ocr_dpi, lang=ocr_lang)

            exploded = _explode_lines(raw_lines)

            tag_ocr = " | OCR" if used_ocr else ""
            print(f"\n--- PÁGINA {pidx} | cruas={len(raw_lines)} | explodidas={len(exploded)}{tag_ocr} ---")
            if used_ocr:
                print("⚠️ AVISO: Esta página foi processada com OCR. Confira os valores extraídos.")
            first_date_only = None
            for i, raw in enumerate(exploded[:max_lines_per_page], start=1):
                line = raw.strip()
                if not line:
                    continue

                tags = []
                if HEADER_RE.search(line) or _is_header_line(line):
                    tags.append("HEADER")
                if FUTUROS_RE.search(line):
                    tags.append("FUTUROS")
                if DATE_LINE_RE.match(line):
                    tags.append("DATE-LINE")
                else:
                    if ONE_DATE_ONLY_RE.match(line):
                        tags.append("DATE1")
                    elif NEXT_DATE_AND_TEXT_RE.match(line):
                        tags.append("DATE2+TXT")

                if DOC_NO_WITH_TEXT_RE.match(line) or DOC_NO_RE.match(line):
                    tags.append("DOC-NO")
                if _is_big_doc_code(line):
                    tags.append("DOC-LONGO")
                if DECIMAL_RE.search(line):
                    tags.append("VALOR")

                parts = _split_on_known_starts(line)
                if len(parts) > 1:
                    tags.append(f"SPLITx{len(parts)}")

                tagtxt = " | ".join(tags) if tags else "..."
                print(f"[P{pidx}:{i:03d}] {tagtxt}  {line}")

    fitz_doc.close()

    print("\n==== RESUMO PARSER ====")
    print(f"Lançamentos extraídos: {len(total_rows)}")
    for j, r in enumerate(total_rows[:20], start=1):
        dv = r.get('valor')
        try:
            val = f"{Decimal(dv):.2f}"
        except Exception:
            val = str(dv)
        print(f"{j:02d}. {r.get('data_liq')} {r.get('data_mov')} | {r.get('documento') or ''} | {val} | {r.get('descricao')[:80] if r.get('descricao') else ''}")

if __name__ == "__main__":
    import argparse, os
    parser = argparse.ArgumentParser(description="Debug do parser de extrato XP (com OCR condicional + aviso)")
    parser.add_argument("pdf", help="Caminho do extrato PDF")
    parser.add_argument("--max", type=int, default=200, help="Máx. linhas por página no debug")
    parser.add_argument("--ocr-lang", default="por", help="Idioma do Tesseract (ex.: 'por' ou 'por+eng')")
    parser.add_argument("--ocr-dpi", type=int, default=300, help="DPI para renderização antes do OCR")
    args = parser.parse_args()

    path = os.path.expanduser(args.pdf)
    if not os.path.isfile(path):
        raise FileNotFoundError(f"Arquivo não encontrado: {path}")
    debug_dump_pdf(path, max_lines_per_page=args.max, ocr_lang=args.ocr_lang, ocr_dpi=args.ocr_dpi)
