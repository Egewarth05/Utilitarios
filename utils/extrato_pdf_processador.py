
import re
import pdfplumber
import unicodedata
from decimal import Decimal, InvalidOperation
from typing import List, Dict, Optional, Tuple

DECIMAL_RE = re.compile(r'([-\u2212]?)R\$\s*([\d\.\,\u00A0]+)')  # aceita NBSP e sinal U+2212
HEADER_RE = re.compile(r'^\s*Liq\s+Mov\s+Histórico\s+Valor\s+Saldo', re.IGNORECASE)
FUTUROS_RE = re.compile(r'^\s*Lançamentos futuros', re.IGNORECASE)
# Documento no formato "Nº 116011339" (ou "No 116011339")
DOC_NO_RE = re.compile(r'^(?:N[ºo]\s*)?(\d{6,})$', re.IGNORECASE)
DOC_NO_PURE_RE = re.compile(r'^(?:N[ºo]\s*)(\d{6,})$', re.IGNORECASE)
DOC_NO_WITH_TEXT_RE = re.compile(r'^(?:N[ºo]\s*)(\d{6,})\s+(.+)$', re.IGNORECASE)
# linha que começa um lançamento: duas datas; o texto depois é OPCIONAL
DATE_LINE_RE = re.compile(
    r'^(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})(?:\s+(.*))?$'
)

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
    """
    Constrói linhas estáveis a partir das palavras do PDF,
    agrupando por coordenada vertical (top) com tolerância.
    Evita que duas linhas visuais virem uma só.
    """
    words = page.extract_words(
        x_tolerance=2,     # não juntar palavras de colunas diferentes
        y_tolerance=2,     # sensibilidade a quebras de linha
        keep_blank_chars=False,
        use_text_flow=True
    )
    lines = []
    row = []
    last_top = None
    for w in words:
        t = round(w["top"], 1)
        if last_top is None or abs(t - last_top) <= 2:
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

KNOWN_STARTS = [
    "TAXA DE INTERMEDIACAO",
    "CREDITO REF",               
    "DEBITO CBLC IRRF",          
    "OPERACOES EM BOLSA LIQ",    
    "DIVIDENDOS DE CLIENTES",
    "CREDITO DE REEMBOLSO",      
]

# um segundo par de datas no meio da linha também inicia novo lançamento
DATEPAIR_RE = re.compile(r'\d{2}/\d{2}/\d{4}\s+\d{2}/\d{2}/\d{4}')

def _normalize_with_map(s: str) -> tuple[str, list[int]]:
    """Normaliza (sem acentos, maiúsc.) e devolve o mapa de posições norm->orig."""
    norm_chars = []
    pos_map = []
    for i, ch in enumerate(s):
        base = unicodedata.normalize('NFKD', ch)
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
    """Aplica o corte acima a todas as linhas da página."""
    out = []
    for l in lines:
        out.extend(_split_on_known_starts(l))
    return out

def parse_xp_extrato_pdf(pdf_path: str) -> List[Dict]:
    rows: List[Dict] = []
    in_table = False
    current = None

    pending_desc = ""              
    candidate_desc = ""            
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            raw_lines = _page_to_lines(page)
            if not raw_lines:  # fallback extra
                text = page.extract_text() or ""
                raw_lines = [l.rstrip() for l in text.splitlines()]
            lines = _explode_lines(raw_lines)
            for raw_line in lines:
                line = raw_line.strip()
                if not line:
                    continue

                m_dates = DATE_LINE_RE.match(line)

                if not in_table:
                    # entra na tabela se ver o header OU já ver a primeira linha com datas
                    if HEADER_RE.search(line) or _is_header_line(line) or m_dates:
                        in_table = True
                        # se não for uma linha de datas, passa pra próxima
                        if not m_dates:
                            continue
                    else:
                        continue

                # 1) Fim da tabela
                if FUTUROS_RE.search(line):
                        if m_dates:
                            # fecha o anterior (sem checar valor/saldo)
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
                                'valor': None,        # só valor
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
                    # ignora valores/doc longo antes da data
                    if DECIMAL_RE.search(line) or _is_big_doc_code(line):
                        continue
                    # resto é texto: acumula
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

                # 4.3) Linha com valor (se ainda não setado)
                m_val = DECIMAL_RE.search(line)
                if m_val and current.get('valor') is None:
                    sign_v, v = m_val.groups()
                    valor = _to_decimal_br(v)
                    if sign_v in ('-', '\u2212'):
                        valor = -valor
                    current['valor'] = valor
                    current['tipo'] = 'D' if valor < 0 else 'C'
                    continue

                # 4.4) Texto comum -> descrição
                if not current.get('descricao'):
                    current['descricao'] = _clean_spaces(line)
                    candidate_desc = current['descricao']
                    conta, hist, cp = _classificar_conta_historico(current['descricao'])
                    current['conta'], current['historico_code'], current['contrapartida'] = conta, hist, cp
                else:
                    # só concatena se não for claramente valor/documento
                    current['descricao'] = _clean_spaces(current['descricao'] + ' ' + line)

    if current:
        if not current.get('descricao') and candidate_desc:
            current['descricao'] = candidate_desc
        rows.append(current)  # << inclui mesmo sem valor

    # remove lançamentos sem valor
    rows = [r for r in rows if (r.get('descricao') or r.get('documento'))]
    return rows

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
    df['valor'] = df['valor'].astype(float)
    with pd.ExcelWriter(out_xlsx_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Extrato')

def export_to_txt_contabil(rows: List[Dict], out_txt_path: str, *, codigo_prefixo="1", codigo_meio="5", codigo_cc="337") -> None:
    import hashlib
    def _doc_or_hash(r):
        d = (r.get('documento') or '').strip()
        if d:
            return d
        key = (r.get('data_liq','') + '|' + r.get('descricao','')).encode('utf-8')
        return hashlib.md5(key).hexdigest()[:10].upper()
    def _fmt_data_ddmmaa(dmy: str) -> str:
        return dmy.replace('/','')
    def _fmt_valor_2p(v: Decimal) -> str:
        return f"{abs(v):.2f}"
    with open(out_txt_path, 'w', encoding='utf-8') as f:
        for r in rows:
            if r.get('valor') is None:
                continue
            data = _fmt_data_ddmmaa(r['data_liq'])
            doc = _doc_or_hash(r)
            valor = _fmt_valor_2p(r['valor'])
            hist_txt = (r.get('descricao') or '').replace('"', "'")
            linha = f'{codigo_prefixo},{data},{doc},{codigo_meio},{valor},{codigo_cc},"{hist_txt}"'
            f.write(linha + '\n')

def processar_extrato_pdf(in_pdf_path: str, out_xlsx_path: str, out_txt_path: Optional[str]=None, config: Optional[Dict]=None) -> Dict:
    config = config or {}
    rows = parse_xp_extrato_pdf(in_pdf_path)
    export_to_xlsx(rows, out_xlsx_path)
    if out_txt_path:
        export_to_txt_contabil(
            rows,
            out_txt_path,
            codigo_prefixo=config.get('codigo_prefixo', '1'),
            codigo_meio=config.get('codigo_meio', '5'),
            codigo_cc=config.get('codigo_cc', '337'),
        )
    return {'quantidade_lancamentos': len(rows)}
