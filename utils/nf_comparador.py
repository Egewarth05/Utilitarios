import os
import io
import re
import stat
import shutil
import tempfile
import rarfile
from collections import defaultdict
from decimal import Decimal, InvalidOperation
import fitz  # PyMuPDF
import pytesseract
import pdfplumber
from PIL import Image
from datetime import date
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

DEBUG_NF = True

def _dbg(arq, msg):
    if DEBUG_NF:
        print(f"[NFDBG:{os.path.basename(arq)}] {msg}")

def _ctx(s, pos, w=90):
    a = max(0, pos - w)
    b = min(len(s), pos + w)
    return s[a:b].replace("\n", " ")

# aceita qualquer NFS-e (com ou sem hífen) ou o texto "Nota Fiscal de Serviços"
PATTERN_NFSE = re.compile(r'(NFS[–—-]?E|Nota\s+Fiscal\b)', re.IGNORECASE)

def _money_to_decimal(txt: str):
    if not txt:
        return None
    # remove separadores de milhar: ponto, espaço normal e NBSP
    clean = re.sub(r"[.\s\u00A0\u202F]", "", txt)
    try:
        return Decimal(clean.replace(",", ".")).quantize(Decimal("0.01"))
    except InvalidOperation:
        return None

# configurações externas
rarfile.UNRAR_TOOL = r"C:\Program Files\UnRAR.exe"
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
os.environ["TESSDATA_PREFIX"] = r"C:\Program Files\Tesseract-OCR\tessdata"

UNICODE_SPACES = "\u00A0\u202F\u2007\u2009\u200A\u2008\u2006\u205F"
SPACES = rf"[.\s{UNICODE_SPACES}]"
MONEY  = rf"(?:\d{{1,3}}(?:{SPACES}\d{{3}})*|\d{{4,6}})\s*,\s*\d\s*\d"
SPAN   = r"[\s\S]{0,300}?"                            

def extrair_info_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    texto_puro, ocr_text = "", []
    for page in doc:
        texto_puro += page.get_text() or ""
        pix = page.get_pixmap(dpi=300, alpha=False)
        img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L")
        bw  = img.point(lambda x: 0 if x < 180 else 255, "1")
        cfg = r"--oem 3 --psm 6"
        ocr_text.append(pytesseract.image_to_string(bw, lang="por", config=cfg))
    doc.close()

    texto = (texto_puro or "") + "\n" + "\n".join(ocr_text)
    _dbg(pdf_path, f"text_len={len(texto)}  puro={len(texto_puro)}  ocr_len_total={sum(len(t) for t in ocr_text)}")

    # Ignora NFe (não NFS-e)
    if re.search(r'\bNFE\b', texto, re.IGNORECASE) and not PATTERN_NFSE.search(texto):
        return None

    def _norm(s: str) -> str:
        s = (s or "").replace('\u00A0', ' ')
        return re.sub(r'[ \t]+', ' ', s)

    nome = os.path.basename(pdf_path)
    m = re.search(r"(\d+)", nome)
    raw = m.group(1) if m else ""
    parts = re.split(r"0{4,}", raw)
    numero = parts[-1].lstrip("0") if len(parts) > 1 and parts[-1] else raw.lstrip("0")

    # Normalização agressiva de espaços e separadores
    header_section = _norm(texto)
    header_section = (header_section
                      .replace('\u00A0', ' ')
                      .replace('\u202F', ' ')
                      .replace('\u2007', ' ')
                      .replace('\u2009', ' '))
    def _compact_money_digits(s: str) -> str:
        def repl(m):
            left = re.sub(r"\s+", "", m.group(1))
            return f"{left},{m.group(2)}{m.group(3)}"
        return re.sub(r'(?<!\d)(\d(?:\s?\d){0,10})\s*,\s*(\d)\s*(\d)(?!\d)', repl, s)
    header_section = re.sub(r'R\$\s*(?=\d)', 'R$ ', header_section)
    header_section = re.sub(r'(\d)\s*,\s*(\d{2})', r'\1,\2', header_section)
    header_section = _compact_money_digits(header_section)

    # ===== DATA =====
    RX_DATA_ANY = re.compile(r"(\d{2}[\/\.-]\d{2}[\/\.-](?:\d{2}|\d{4}))", re.I)
    RX_EMISSAO_NEAR = re.compile(
        r"(?:data(?:\s+\w+){0,5}\s*(?:da\s*nota\s*|de\s*)?emiss[aã]o)\D{0,120}" + RX_DATA_ANY.pattern,
        re.I
    )
    
    RX_CNPJ = re.compile(r'\b\d{2}[.\s]?\d{3}[.\s]?\d{3}\s*/\s*\d{4}\s*-\s*\d{2}\b')
    RX_CPF  = re.compile(r'\b\d{3}[.\s]?\d{3}[.\s]?\d{3}\s*-\s*\d{2}\b')
    header_section = RX_CNPJ.sub(' [CNPJ] ', header_section)
    header_section = RX_CPF.sub(' [CPF] ', header_section)
    linhas = [x.strip() for x in header_section.splitlines()]
    flat   = " ".join(linhas)
    
    def _canon_date(s: str) -> str:
        d, m, y = re.split(r"[\/\.-]", s)
        if len(y) == 2:
            y = ("20" if int(y) <= 49 else "19") + y
        return f"{d}/{m}/{y}"

    def _first_date(s: str):
        m = RX_DATA_ANY.search(s)
        return _canon_date(m.group(1)) if m else None

    def _pick_date(lines):
        HAS_TIME   = re.compile(r"\b\d{2}:\d{2}:\d{2}\b")
        BAN_DATE   = ("impress", "compet", "venc", "parcela", "boleto")
        IS_PERIOD  = re.compile(r"\bper[íi]odo\b", re.IGNORECASE)
        TWO_DATES  = re.compile(r"\d{2}/\d{2}/\d{4}\s*(?:a|-|–)\s*\d{2}/\d{2}/\d{4}", re.IGNORECASE)

        for i, ln in enumerate(lines):
            l = ln.lower()
            if "data" in l and "emiss" in l and "impress" not in l:
                d = _first_date(ln)
                if d: return d
                for j in range(max(0, i-2), min(len(lines), i+10)):
                    lj = lines[j].lower()
                    if any(b in lj for b in BAN_DATE) or IS_PERIOD.search(lines[j]) or TWO_DATES.search(lines[j]):
                        continue
                    d = _first_date(lines[j])
                    if d: return d

        for i, ln in enumerate(lines):
            l = ln.lower()
            if "data" in l and ("servi" in l or "execu" in l) and not IS_PERIOD.search(ln) and not TWO_DATES.search(ln):
                d = _first_date(ln)
                if d: return d
                for j in range(max(0, i-2), min(len(lines), i+3)):
                    lj = lines[j].lower()
                    if any(b in lj for b in BAN_DATE) or HAS_TIME.search(lines[j]) or IS_PERIOD.search(lines[j]) or TWO_DATES.search(lines[j]):
                        continue
                    d = _first_date(lines[j])
                    if d: return d

        BAN_DATE_EXTRA = ("alvar","licen","vigênc","vigenc","simples","cnae","fund","constit","abertura")
        def _ymd_tuple(dstr: str):
            return (int(dstr[6:10]), int(dstr[3:5]), int(dstr[0:2]))
        cands = []
        for ln in lines:
            l = ln.lower()
            if any(b in l for b in BAN_DATE) or any(b in l for b in BAN_DATE_EXTRA):
                continue
            if HAS_TIME.search(ln) or IS_PERIOD.search(ln) or TWO_DATES.search(ln):
                continue
            for m in RX_DATA_ANY.finditer(ln):
                d = _canon_date(m.group(1))
                cands.append((_ymd_tuple(d), d))
        if cands:
            cands.sort(reverse=True)
            return cands[0][1]
        return None

    m_emissao = RX_EMISSAO_NEAR.search(flat)
    if m_emissao:
        data = _canon_date(m_emissao.group(1))
    else:
        data = _pick_date(linhas)
    if not data:
        todas = [_canon_date(m.group(1)) for m in RX_DATA_ANY.finditer(flat)]
        if todas:
            data = max(todas, key=lambda d: (int(d[6:10]), int(d[3:5]), int(d[0:2])))
    if data:
        md = re.match(r"(\d{2})/(\d{2})/\d{2,4}$", data)
        if md:
            data = f"{md.group(1)}/{md.group(2)}/2025"

    # --- rótulos bons e termos a evitar perto do número ---
    GOOD_ANCHOR = re.compile(
        r"(?:(?:valor|vlr\.?)\s+bruto\s+da\s+nota|"
        r"valor\s+total\s+(?:da\s+nfs[–—-]?e|da\s+nota|do\s+documento)|"
        r"valor\s+dos?\s+servi[cç]os?|"
        r"total\s+do\s+servi[cç]o|"
        r"valor\s+l[ií]quido|"
        r"valor\s+da\s+nota|"
        r"vlr\.?\s*total|"
        r"valor\s+do\s+documento|"
        r"valor\s+do\s+servi[cç]o|"
        r"valor\s+(?:bruto|l[ií]quido|unit[áa]rio)\s+do\s+servi[cç]o|"
        r"valor\s+l[ií]quido\s+do\s+servi[cç]o|"
        r"valor\s+a\s+pagar|total\s+a\s+pagar|"
        r"valor\s+total\s+da\s+nota(?:\s+fiscal)?(?:\s+de\s+servi[cç]os)?|"
        r"valor\s+bruto\s+da\s+nota(?:\s+fiscal)?(?:\s+de\s+servi[cç]os)?)",
        re.IGNORECASE
    )

    BAD_TAXES = r"(?:iss|issqn|pis|cofins|csll|inss|irrf)"
    BAD_ISS_FUZZ = r"(?:iss|i[s5]{2}|[1iIl][s5]{2})"

    BAD_CTX = re.compile(
        rf"\b({BAD_ISS_FUZZ}|issqn|pis|cofins|csll|inss|irrf|"
        r"al[ií]q|al[ií]quota|ret(?:id|en)|dedu[cç][aã]o|descon|"
        r"base\s+de\s+c[aá]lculo|base\s+calc|ibpt|aprox(?:imad[oa])?|"
        r"tribut|imposto|parcela|parcelas|venc(?:imento)?|juros|multa|"
        r"pagamento|boleto|duplicata|carn[eé]|"
        r"periodo|per[íi]odo|compet[eê]ncia|compet|"
        r"quantidade|qtd|descri[cç][aã]o|cod\.?\s*serv|abatiment|percent|%|cep|"
        r"cnpj|cpf|inscri[cç][aã]o|rps|c[oó]d|verifica[cç][aã]o|autenticidade)\b",
        re.IGNORECASE
    )

    # usados no PASSO 3
    rx_val_plain = re.compile(rf'(?<!\d)({MONEY})(?!\d)')
    rx_val_rs    = re.compile(rf'(?<!\d)R\$\s*({MONEY})(?!\d)', re.IGNORECASE)

    NEAR_LABELS = re.compile(
        r"(valor\s+total\s+da\s+(?:nfs[–—-]?e|nota(?:\s+fiscal)?(?:\s+de\s+servi[cç]os)?|documento)|"
        r"valor\s+dos?\s+servi[cç]os?|"
        r"valor\s+do\s+servi[cç]o|"
        r"valor\s+unit[áa]rio\s+do\s+servi[cç]o|"
        r"valor\s+bruto\s+do\s+servi[cç]o|"
        r"valor\s+l[ií]quido(?:\s+do\s+servi[cç]o)?|"
        r"valor\s+a\s+pagar|total\s+a\s+pagar|"
        r"valor\s+da\s+nota)",
        re.IGNORECASE
    )
    RX_MONEY_ANY = re.compile(rf"(R\$\s*)?({MONEY})", re.IGNORECASE)

    best = None  # tuple(score, valor_decimal)

    def _score_candidate(has_rs: bool, same_line: bool, dist_chars: int, val: Decimal) -> int:
        # quanto menor melhor
        score  = 0 if same_line else 120          # penaliza forte se não estiver na mesma linha
        score += 0 if has_rs else 60              # ter "R$" ajuda bem
        score += min(dist_chars, 60)              # proximidade do rótulo
        return score

    for i, ln in enumerate(linhas):
        if not NEAR_LABELS.search(ln) or BAD_CTX.search(ln):
            continue

        bloco_linhas = linhas[i:i+3]            # linha do rótulo + duas seguintes
        bloco = " ".join(bloco_linhas)
        lab_end = max((m.end() for m in NEAR_LABELS.finditer(ln.lower())), default=len(ln))

        # 1) juntar todos candidatos do bloco
        cands = []
        for m in RX_MONEY_ANY.finditer(bloco):
            has_rs = bool(m.group(1))
            raw_val = m.group(2)
            dec = _money_to_decimal(raw_val)
            if dec is None or dec <= 0:
                continue

            ctx = bloco[max(0, m.start()-60): m.end()+60].lower()
            if BAD_CTX.search(ctx):
                continue

            # está na mesma linha do rótulo?
            off = 0
            same_line = False
            for k, ltxt in enumerate(bloco_linhas):
                limite = off + len(ltxt) + (1 if k > 0 else 0)
                if m.start() < limite:
                    same_line = (k == 0)
                    dist = abs(lab_end - (m.start()-off)) if same_line else 120
                    break
                off = limite

            has_iss_like = re.search(BAD_ISS_FUZZ, ctx) is not None
            cands.append((dec, has_rs, same_line, dist, has_iss_like))

        if not cands:
            continue

        # 2) penalizar valores muito pequenos quando existe um muito maior no bloco
        bloco_max = max(c[0] for c in cands)
        for dec, has_rs, same_line, dist, has_iss_like in cands:
            score = _score_candidate(has_rs, same_line, dist, dec)

            if bloco_max and dec <= bloco_max * Decimal("0.15"):
                score += 400

            if has_iss_like and bloco_max and dec <= bloco_max * Decimal("0.25"):
                score += 1000

            pref = (dec >= Decimal("200.00"), dec)

            if best is None or score < best[0] or (score == best[0] and pref > (best[2] if len(best) > 2 else (False, Decimal("0")))):
                best = (score, dec, pref)

    if best is not None:
        valor = str(best[1])
        _dbg(pdf_path, f"[NEAR-LINE] score={best[0]} valor={valor}")
        return {"numero": numero, "data": data, "valor": valor}

    def _canon_date(s: str) -> str:
        d, m, y = re.split(r"[\/\.-]", s)
        if len(y) == 2:
            y = ("20" if int(y) <= 49 else "19") + y
        return f"{d}/{m}/{y}"

    def _first_date(s: str):
        m = RX_DATA_ANY.search(s)
        return _canon_date(m.group(1)) if m else None

    def _pick_date(lines):
        # preferir "Data de emissão" (evitar "impressão")
        HAS_TIME   = re.compile(r"\b\d{2}:\d{2}:\d{2}\b")
        BAN_DATE   = ("impress", "compet", "venc", "parcela", "boleto")
        IS_PERIOD  = re.compile(r"\bper[íi]odo\b", re.IGNORECASE)
        TWO_DATES  = re.compile(r"\d{2}/\d{2}/\d{4}\s*(?:a|-|–)\s*\d{2}/\d{2}/\d{4}", re.IGNORECASE)

        for i, ln in enumerate(lines):
            l = ln.lower()
            if "data" in l and "emiss" in l and "impress" not in l:
                d = _first_date(ln)
                if d:
                    return d
                for j in range(max(0, i-2), min(len(lines), i+10)):
                    lj = lines[j].lower()
                    if any(b in lj for b in BAN_DATE) or IS_PERIOD.search(lines[j]) or TWO_DATES.search(lines[j]):
                        continue
                    d = _first_date(lines[j])
                    if d:
                        return d

        # Data do serviço/execução (sem período)
        for i, ln in enumerate(lines):
            l = ln.lower()
            if "data" in l and ("servi" in l or "execu" in l) and not IS_PERIOD.search(ln) and not TWO_DATES.search(ln):
                d = _first_date(ln)
                if d:
                    return d
                for j in range(max(0, i-2), min(len(lines), i+3)):
                    lj = lines[j].lower()
                    if any(b in lj for b in BAN_DATE) or HAS_TIME.search(lines[j]) or IS_PERIOD.search(lines[j]) or TWO_DATES.search(lines[j]):
                        continue
                    d = _first_date(lines[j])
                    if d:
                        return d

        # fallback: maior data “limpa”
        BAN_DATE_EXTRA = ("alvar","licen","vigênc","vigenc","simples","cnae","fund","constit","abertura")
        def _ymd_tuple(dstr: str):
            return (int(dstr[6:10]), int(dstr[3:5]), int(dstr[0:2]))
        cands = []
        for ln in lines:
            l = ln.lower()
            if any(b in l for b in BAN_DATE) or any(b in l for b in BAN_DATE_EXTRA):
                continue
            if HAS_TIME.search(ln) or IS_PERIOD.search(ln) or TWO_DATES.search(ln):
                continue
            for m in RX_DATA_ANY.finditer(ln):
                d = _canon_date(m.group(1))
                cands.append((_ymd_tuple(d), d))
        if cands:
            cands.sort(reverse=True)
            return cands[0][1]
        return None

    linhas = [x.strip() for x in header_section.splitlines()]
    flat = " ".join(linhas)

    # ------ DATA ------
    m_emissao = RX_EMISSAO_NEAR.search(flat)
    if m_emissao:
        data = _canon_date(m_emissao.group(1))
    else:
        data = _pick_date(linhas)
    if not data:
        todas = [_canon_date(m.group(1)) for m in RX_DATA_ANY.finditer(flat)]
        if todas:
            data = max(todas, key=lambda d: (int(d[6:10]), int(d[3:5]), int(d[0:2])) )
    if data:
        md = re.match(r"(\d{2})/(\d{2})/\d{2,4}$", data)
        if md:
            data = f"{md.group(1)}/{md.group(2)}/2025"

    # ===== 1) label -> número (apenas para FRENTE, janela curta) =====
    PREF_FLOOR = Decimal("200.00")  
    WIN = 180                        

    best_val = None
    for ml in NEAR_LABELS.finditer(flat):
        s = ml.end()
        e = min(len(flat), ml.end() + WIN)
        trecho = flat[s:e]

        local_cands = []

        # 1) Tente primeiro com "R$"
        for mg in re.finditer(rf"R\$\s*({MONEY})", trecho, re.IGNORECASE):
            ctx = trecho[max(0, mg.start()-40): mg.end()+40].lower()
            if BAD_CTX.search(ctx):
                continue
            dec = _money_to_decimal(mg.group(1))
            if dec is not None and dec > 0:
                local_cands.append(dec)

        if not local_cands:
            for mg in re.finditer(rf"(?<!\d)({MONEY})(?!\d)", trecho):
                ctx = trecho[max(0, mg.start()-40): mg.end()+40].lower()
                if BAD_CTX.search(ctx):
                    continue
                dec = _money_to_decimal(mg.group(1))
                if dec is not None and dec > 0:
                    local_cands.append(dec)

        if local_cands:
            prefer = [v for v in local_cands if v >= PREF_FLOOR]
            cand = max(prefer) if prefer else max(local_cands)
            if best_val is None or cand > best_val:
                best_val = cand

    if best_val is not None:
        valor = str(best_val)
        _dbg(pdf_path, f"[LABEL-FWD] valor={valor}")
        return {"numero": numero, "data": data, "valor": valor}

    # ===== SEGUNDO: global FWD (rótulo -> número), janela curta =====
    SPAN_NEAR = r"[\s\S]{0,80}?"
    RX_FWD = [
        re.compile(rf"valor\s+total\s+da\s+(?:nfs[–—-]?e|nota(?:\s+fiscal)?(?:\s+de\s+servi[cç]os)?)"
                   rf"{SPAN_NEAR}(?:R\$\s*)?({MONEY})", re.I|re.S),
        re.compile(rf"valor\s+dos?\s+servi[cç]os?{SPAN_NEAR}(?:R\$\s*)?({MONEY})", re.I|re.S),
        re.compile(rf"(?:valor|vlr\.?)\s+bruto\s+da\s+(?:nota(?:\s+fiscal)?(?:\s+de\s+servi[cç]os)?)"
                   rf"{SPAN_NEAR}(?:R\$\s*)?({MONEY})", re.I|re.S),
        re.compile(rf"valor\s+l[ií]quido(?:\s+da\s+nota\s+fiscal)?{SPAN_NEAR}(?:R\$\s*)?({MONEY})", re.I|re.S),
        re.compile(rf"(?:fatura|duplicata){SPAN_NEAR}valor{SPAN_NEAR}(?:R\$\s*)?({MONEY})",
               re.I | re.S),
    ]

    best_val = None
    best_dist = 10**9
    for rx in RX_FWD:
        for m in rx.finditer(flat):
            dec = _money_to_decimal(m.group(1))
            ctx = flat[max(0, m.start(1)-80): m.end(1)+80].lower()
            label_txt = flat[max(0, m.start()-80): m.start(1)].lower()
            is_fatura = ('fatura' in label_txt) or ('duplicata' in label_txt)
            if dec is None or dec <= 0:
                continue
            # só aplica BAD_CTX se não for o caso FATURA/DUPLICATA
            if (not is_fatura) and BAD_CTX.search(ctx):
                continue
            dist = m.start(1) - m.start()
            if dist < best_dist or (dist == best_dist and (best_val is None or dec > best_val)):
                best_dist, best_val = dist, dec

    if best_val is not None:
        valor = str(best_val)
        _dbg(pdf_path, f"[GLOBAL-FWD] dist={best_dist} valor={valor}")
        return {"numero": numero, "data": data, "valor": valor}

    # ------ PASSO 3: janela ancorada (±3 linhas) ------
    header_section2 = re.sub(r'(,\d{2})(?=\d)', r'\1 ', header_section)
    linhas = [_norm(x) for x in header_section2.splitlines()]
    flat   = " ".join(linhas)

    for i, ln in enumerate(linhas):
        if not GOOD_ANCHOR.search(ln) or BAD_CTX.search(ln):
            continue
        janela = linhas[max(0, i-3):min(len(linhas), i+4)]
        cands_txt = []
        for w in janela:
            cands_txt += [m.group(1) for m in rx_val_rs.finditer(w)]
            cands_txt += [m.group(1) for m in rx_val_plain.finditer(w)]

        decs = [(d, t) for t in cands_txt if (d := _money_to_decimal(t)) is not None]
        if decs:
            _, melhor = max(decs, key=lambda t: t[0])
            valor = str(_money_to_decimal(melhor))
            _dbg(pdf_path, f"[ANC-JANELA] -> {melhor}")
            return {"numero": numero, "data": data, "valor": valor}

    # ------ Fallback conservador ------
    def _plausivel(v: Decimal) -> bool:
        return Decimal('0.01') <= v <= Decimal('100000.00')

    GOOD_NEAR = re.compile(r'(valor|total|nfs|nota|servi[cç]o|l[ií]quido|bruto)', re.IGNORECASE)
    def _ok_context(pos: int) -> bool:
        janela = flat[max(0, pos-150):pos+150].lower()
        perto  = flat[max(0, pos-10):pos+10]
        return (GOOD_NEAR.search(janela)
                and '/' not in perto
                and not BAD_CTX.search(janela))

    candidatos = []
    for m in rx_val_rs.finditer(flat):
        if not _ok_context(m.start()):
            continue
        v = _money_to_decimal(m.group(1))
        if v is not None and _plausivel(v):
            candidatos.append((v, m.group(1), m.start()))
    if not candidatos:
        for m in rx_val_plain.finditer(flat):
            if not _ok_context(m.start()):
                continue
            v = _money_to_decimal(m.group(1))
            if v is not None and _plausivel(v):
                candidatos.append((v, m.group(1), m.start()))

    valor = None
    if candidatos:
        v, valor_str, pos = max(candidatos, key=lambda r: r[0])
        _dbg(pdf_path, f"[FALLBACK] pos={pos} val={valor_str} ctx='{_ctx(flat, pos)}'")
        valor = str(_money_to_decimal(valor_str)) if valor_str else None

    _dbg(pdf_path, f"[RETORNO] numero={numero} data={data} valor={valor}")
    return {"numero": numero, "data": data, "valor": valor}

# === 2) Extrai notas (info) do RAR ===
def extrair_notas_zip(zip_path, temp_dir):
    """
    Extrai e processa todos os arquivos PDF dentro do RAR, inclusive em subpastas.
    Retorna lista de dicionários com infos extraídas e lista de arquivos sem dados (sem número).
    """
    os.makedirs(temp_dir, exist_ok=True)
    with rarfile.RarFile(zip_path) as rar:
        rar.extractall(temp_dir)

    # varre recursivamente todos os PDFs extraídos
    pdfs = []
    for root, dirs, files in os.walk(temp_dir):
        for fn in files:
            if not fn.lower().endswith('.pdf'):
                continue
            if 'fatura' in fn.lower():
                continue
            # Ignorar NFe pelo NOME: tem "nfe" e NÃO tem "nfs"/"nfs-e"
            fn_lower = fn.lower()
            if re.search(r'\bnf[\s\-_.]*e\b', fn_lower) and not re.search(r'\bnfs[\s\-_.]*e?\b', fn_lower):
                continue
            pdfs.append(os.path.join(root, fn))

    notas, sem_dados = [], []
    for pdf in pdfs:
        nome_arquivo = os.path.basename(pdf)
        m_num = re.search(r"(\d+)", nome_arquivo)
        if not m_num:
            sem_dados.append(nome_arquivo)
            continue
        numero = m_num.group(1).lstrip("0")

        info = extrair_info_pdf(pdf)
        data = info.get("data") if info else None
        valor = info.get("valor") if info else None

        notas.append({
            "numero": numero,
            "data": data,
            "valor": valor,
            "arquivo": nome_arquivo
        })

    return notas, sem_dados

# === 3) Extrai relatório via pdfplumber ===
def extrair_relatorio(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        rows = []
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                for row in table:
                    if any(cell for cell in row):
                        rows.append([cell or "" for cell in row])
    if not rows:
        raise ValueError("Nenhuma linha extraída do relatório.")
    header = [c.replace("\n", " ").strip().lower() for c in rows[0]]
    idx_doc = next(i for i,h in enumerate(header) if "docum" in h)
    idx_esp = next(i for i,h in enumerate(header) if "espécie" in h)
    idx_date = next(i for i,h in enumerate(header) if "entrada" in h)
    idx_val = next(i for i,h in enumerate(header) if "valor" in h)
    rel = []
    for row in rows[1:]:
        esp = row[idx_esp].strip().upper()
        if esp == "NFE" or not esp.startswith("NFS"):
            continue
        numero = row[idx_doc].strip()
        data = row[idx_date].strip()
        raw = row[idx_val].strip()
        try:
            valor = str(Decimal(raw.replace(".", "").replace(",", ".")).quantize(Decimal("0.01")))
        except InvalidOperation:
            continue
        rel.append({"numero": numero, "data": data, "valor": valor})
    return rel

# === 4) Compara e gera PDF de validação ===
def gerar_pdf_relatorio(encontradas, divergentes, nao_encontradas, sem_dados, output_path):
    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4
    y = height - 50

    def fmt_val(v):
        return f"R$ {v}" if v not in (None, "", "—") else "—"

    def fmt_date(d):
        return d if d not in (None, "", "—") else "—"

    def add_secao(titulo, itens):
        nonlocal y
        c.setFont("Helvetica-Bold", 12)
        c.drawString(40, y, f"{titulo} ({len(itens)})")
        y -= 20
        c.setFont("Helvetica", 10)
        if not itens:
            c.drawString(50, y, "Nenhum item.")
            y -= 30
            return
        for it in itens:
            if y < 50:
                c.showPage()
                y = height - 50

            linha = (
                f"Nº: {it['numero']} | "
                f"Data: {fmt_date(it.get('data'))} | "
                f"Valor: {fmt_val(it.get('valor'))}"
            )
            if it.get("arquivo"):
                linha += f" | Arquivo: {it['arquivo']}"

            if it.get("esperado"):
                exp = it["esperado"]
                linha += (
                    f" | Esperado: Nº {exp.get('numero')}, "
                    f"{fmt_date(exp.get('data'))}, {fmt_val(exp.get('valor'))}"
                )

            c.drawString(50, y, linha)
            y -= 15
        y -= 20

    c.setFont("Helvetica-Bold", 14)
    c.drawString(40, y, "Relatório de Validação de NFS-e")
    y -= 30
    add_secao("✔ Encontradas e Correspondentes", encontradas)
    add_secao("⚠ Divergentes (data ou valor)", divergentes)
    add_secao("❌ Não encontradas no relatório", nao_encontradas)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(40, y, f"🛈 Arquivos sem extração ({len(sem_dados)})")
    y -= 20
    c.setFont("Helvetica", 10)
    for fname in sem_dados:
        if y < 50:
            c.showPage()
            y = height - 50
        c.drawString(50, y, fname)
        y -= 15
    c.save()

def to_decimal_br(s):
    if s is None:
        return None
    try:
        s = str(s).strip()
        s = s.replace('.', '').replace(',', '.')
        return Decimal(s).quantize(Decimal('0.01'))
    except (InvalidOperation, AttributeError, ValueError):
        return None

# === 5) Função principal ===
def processar_comparacao_nf(zip_path, relatorio_pdf_path, output_dir):

    temp = os.path.join(os.path.dirname(output_dir), "temp_notas")
    if os.path.isdir(temp):
        shutil.rmtree(temp, onerror=lambda f, p, e: os.chmod(p, stat.S_IWRITE) or f(p))
    if os.path.isdir(output_dir):
        shutil.rmtree(output_dir, onerror=lambda f, p, e: os.chmod(p, stat.S_IWRITE) or f(p))
    os.makedirs(temp, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    # Extrai NFS-e do RAR (NF-e já são ignoradas) e lê o relatório
    notas, sem_dados = extrair_notas_zip(zip_path, temp)
    rel = extrair_relatorio(relatorio_pdf_path)

    encontradas, divergentes, nao_encontradas = [], [], []

    # Indexa relatório por número (pode haver mais de uma linha por número)
    idx = defaultdict(list)
    for r in rel:
        num_str = r.get("numero", "")
        if num_str.isdigit():
            idx[int(num_str)].append(r)

    for nf in notas:
        num_str = nf.get("numero", "")
        if not num_str.isdigit():
            nao_encontradas.append(nf)
            continue

        num = int(num_str)
        matches = idx.get(num, [])
        if not matches:
            nao_encontradas.append(nf)
            continue

        # Dados vindos EXCLUSIVAMENTE do OCR/texto da nota
        val_nf  = to_decimal_br(nf.get('valor'))
        data_nf = nf.get('data')

        # Match exato: mesma data e mesmo valor
        exact = None
        for r in matches:
            val_r = to_decimal_br(r.get('valor'))
            mesma_data  = (data_nf and r.get('data') == data_nf)
            mesmo_valor = (val_nf is not None and val_r is not None and val_r == val_nf)
            if mesma_data and mesmo_valor:
                exact = r
                break

        if exact:
            encontradas.append(nf)
        else:
            nf['esperado'] = matches[0]  
            divergentes.append(nf)

    MISSING = "—"  

    def _sanitize(items):
        for it in items:
            if not it.get("valor"): it["valor"] = MISSING
            if not it.get("data"):  it["data"]  = MISSING
            if it.get("esperado"):
                exp = it["esperado"]
                if not exp.get("valor"): exp["valor"] = MISSING
                if not exp.get("data"):  exp["data"]  = MISSING

    _sanitize(encontradas)
    _sanitize(divergentes)
    _sanitize(nao_encontradas)

    pdf_out = os.path.join(output_dir, "relatorio_validacao.pdf")
    gerar_pdf_relatorio(encontradas, divergentes, nao_encontradas, sem_dados, pdf_out)

    resultado = {
        "encontradas": encontradas,
        "divergentes": divergentes,
        "nao_encontradas": nao_encontradas,
        "sem_dados": sem_dados,
        "pdf": pdf_out
    }
    return resultado, pdf_out
