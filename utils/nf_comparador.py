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

# configurações externas
rarfile.UNRAR_TOOL = r"C:\Program Files\UnRAR.exe"
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
os.environ["TESSDATA_PREFIX"] = r"C:\Program Files\Tesseract-OCR\tessdata"

def extrair_info_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    texto_puro, ocr_text = "", []
    for page in doc:
        texto_puro += page.get_text() or ""
        pix = page.get_pixmap(dpi=300, alpha=False)
        img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L")
        bw  = img.point(lambda x: 0 if x < 180 else 255, "1")
        cfg = r"--oem 3 --psm 6"   # sem whitelist
        ocr_text.append(pytesseract.image_to_string(bw, lang="por", config=cfg))
    doc.close()

    texto = (texto_puro or "") + "\n" + "\n".join(ocr_text)
    _dbg(pdf_path, f"text_len={len(texto)}  puro={len(texto_puro)}  ocr_len_total={sum(len(t) for t in ocr_text)}")

    # Se for NFe (e não NFS-e), ignore
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

    header_section = _norm(texto)

    # ------- helpers / regex -------
    HAS_TIME   = re.compile(r"\b\d{2}:\d{2}:\d{2}\b")
    BAN_DATE   = ("impress", "compet", "venc", "parcela", "boleto")
    IS_PERIOD  = re.compile(r"\bper[íi]odo\b", re.IGNORECASE)
    TWO_DATES  = re.compile(r"\d{2}/\d{2}/\d{4}\s*(?:a|-|–)\s*\d{2}/\d{2}/\d{4}", re.IGNORECASE)

    RX_DATA_ANY = re.compile(r"(\d{2}[\/\.-]\d{2}[\/\.-](?:\d{2}|\d{4}))", re.I)
    RX_EMISSAO_NEAR = re.compile(
        r"(?:data(?:\s+\w+){0,5}\s*(?:da\s*nota\s*|de\s*)?emiss[aã]o)\D{0,120}" + RX_DATA_ANY.pattern,
        re.I
    )

    # Âncoras e contexto
    GOOD_ANCHOR = re.compile(
        r'((?:valor|vlr\.?)\s+bruto\s+da\s+nota|'
        r'valor\s+total\s+(?:da\s+nfs[–—-]?e|da\s+nota)|'
        r'valor\s+dos?\s+servi[cç]os?|'
        r'total\s+do\s+servi[cç]o|'
        r'valor\s+l[ií]quido|'
        r'valor\s+da\s+nota|'
        r'vlr\.?\s*total)',
        re.IGNORECASE
    )
    RX_TOTAL_GLOBAL = re.compile(
        r'valor\s+total\s+(?:da\s+nfs[–—-]?e|da\s+nota)\D{0,60}(?:R\$\s*)?(\d{1,3}(?:\.\d{3})*,\d{2})',
        re.IGNORECASE
    )
    RX_SERV_GLOBAL = re.compile(
        r'valor\s+dos?\s+servi[cç]os?\D{0,60}(?:R\$\s*)?(\d{1,3}(?:\.\d{3})*,\d{2})',
        re.IGNORECASE
    )
    RX_BRUTO_GLOBAL_FWD = re.compile(
        r'(?:valor|vlr\.?)\s+bruto\s+da\s+nota\D{0,60}(?:R\$\s*)?(\d{1,3}(?:\.\d{3})*,\d{2})',
        re.IGNORECASE
    )
    RX_BRUTO_GLOBAL_REV = re.compile(
        r'(\d{1,3}(?:\.\d{3})*,\d{2})\D{0,12}(?:valor|vlr\.?)\s+bruto\s+da\s+nota',
        re.IGNORECASE
    )
    RX_TOTAL_GLOBAL_REV = re.compile(
        r'(?:R\$\s*)?(\d{1,3}(?:\.\d{3})*,\d{2})\D{0,12}'
        r'valor\s+total\s+(?:da\s+nfs[–—-]?e|da\s+nota)',
        re.IGNORECASE
    )

    RX_SERV_GLOBAL_REV = re.compile(
        r'(?:R\$\s*)?(\d{1,3}(?:\.\d{3})*,\d{2})\D{0,12}'
        r'valor\s+dos?\s+servi[cç]os?',
        re.IGNORECASE
    )
    BAD_CTX = re.compile(
        r'\b(iss|issqn|al[ií]q|al[ií]quota|ret(?:id|en)|dedu[cç][aã]o|descon|'
        r'base\s+de\s+c[aá]lculo|base\s+calc|ibpt|aprox(?:imad[oa])?|'
        r'tribut|imposto|parcela|parcelas|venc(?:imento)?|juros|multa|'
        r'pagamento|boleto|duplicata|carn[eé]|'
        r'periodo|per[íi]odo|compet[eê]ncia|compet|'
        r'quantidade|descri[cç][aã]o|cod\.?\s*serv)'
        r'\b',
        re.IGNORECASE
    )

    rx_val_plain = re.compile(r'(?<!\d)(\d{1,3}(?:\.\d{3})*,\d{2})(?!\d)')
    rx_val_rs    = re.compile(r'(?<!\d)R\$\s*(\d{1,3}(?:\.\d{3})*,\d{2})(?!\d)', re.IGNORECASE)

    def _canon_date(s: str) -> str:
        d, m, y = re.split(r"[\/\.-]", s)
        if len(y) == 2:
            y = ("20" if int(y) <= 49 else "19") + y
        return f"{d}/{m}/{y}"

    def _first_date(s: str):
        m = RX_DATA_ANY.search(s)
        return _canon_date(m.group(1)) if m else None

    def _pick_date(lines):
        # 1) Emissão
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
        # 2) Serviço/Execução
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
        # 3) mais recente em contexto ok
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

    # preparar linhas/flat (1ª vez)
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
            data = max(todas, key=lambda d: (int(d[6:10]), int(d[3:5]), int(d[0:2])))

    # força 2025
    if data:
        md = re.match(r"(\d{2})/(\d{2})/\d{2,4}$", data)
        if md:
            data = f"{md.group(1)}/{md.group(2)}/2025"

    # ------ PASSO 1: global cross-line ------
    m = (
        RX_BRUTO_GLOBAL_FWD.search(flat) or
        RX_BRUTO_GLOBAL_REV.search(flat) or
        RX_TOTAL_GLOBAL.search(flat)     or
        RX_TOTAL_GLOBAL_REV.search(flat) or
        RX_SERV_GLOBAL.search(flat)      or
        RX_SERV_GLOBAL_REV.search(flat)
    )
    if m:
        valor = str(Decimal(m.group(1).replace('.','').replace(',','.')).quantize(Decimal('0.01')))
        _dbg(pdf_path, f"[GLOBAL-OK] valor={valor}")
        return {"numero": numero, "data": data, "valor": valor}

    # ------ PASSO 2: janela ancorada ------
    # reprocessa header para separar melhor valores colados
    header_section2 = re.sub(r'(,\d{2})(?=\d)', r'\1 ', header_section)
    linhas = [_norm(x) for x in header_section2.splitlines()]   # redefine linhas
    flat   = " ".join(linhas)

    for i, ln in enumerate(linhas):
        if not GOOD_ANCHOR.search(ln) or BAD_CTX.search(ln):
            continue
        janela = linhas[max(0, i-1):min(len(linhas), i+2)]
        cands = []
        for w in janela:
            for mrs in rx_val_rs.finditer(w):
                cands.append(mrs.group(1))
            for mpl in rx_val_plain.finditer(w):
                cands.append(mpl.group(1))

        if cands:
            # escolhe o MAIOR valor na janela
            decs = []
            for txt in cands:
                try:
                    decs.append((Decimal(txt.replace('.','').replace(',','.')).quantize(Decimal('0.01')), txt))
                except InvalidOperation:
                    pass
            if decs:
                _, melhor = max(decs, key=lambda t: t[0])
                valor = str(Decimal(melhor.replace('.','').replace(',','.')).quantize(Decimal('0.01')))
                _dbg(pdf_path, f"[ANC-JANELA] -> {melhor}")
                return {"numero": numero, "data": data, "valor": valor}

    # ------ Fallback conservador ------
    def _plausivel(v: Decimal) -> bool:
        return Decimal('0.01') <= v <= Decimal('100000.00')

    def _ok_context(pos: int) -> bool:
        janela = flat[max(0, pos-80):pos+80].lower()
        return not BAD_CTX.search(janela)

    candidatos = []
    for m in rx_val_rs.finditer(flat):
        if not _ok_context(m.start()):
            continue
        try:
            v = Decimal(m.group(1).replace('.','').replace(',','.')).quantize(Decimal('0.01'))
            if _plausivel(v):
                candidatos.append((v, m.group(1), m.start()))
        except InvalidOperation:
            pass
    if not candidatos:
        for m in rx_val_plain.finditer(flat):
            if not _ok_context(m.start()):
                continue
            try:
                v = Decimal(m.group(1).replace('.','').replace(',','.')).quantize(Decimal('0.01'))
                if _plausivel(v):
                    candidatos.append((v, m.group(1), m.start()))
            except InvalidOperation:
                pass

    valor = None
    if candidatos:
        v, valor_str, pos = max(candidatos, key=lambda r: r[0])
        _dbg(pdf_path, f"[FALLBACK] pos={pos} val={valor_str} ctx='{_ctx(flat, pos)}'")
        try:
            valor = str(Decimal(valor_str.replace('.', '').replace(',', '.')).quantize(Decimal('0.01')))
        except InvalidOperation:
            valor = None

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
            # >>> IGNORAR NFe PELO NOME <<<
            # se contiver "nfe" e NÃO contiver "nfs" (nem "nfs-e"), ignore
            fn_lower = fn.lower()
            if re.search(r'\bnf[\s\-_.]*e\b', fn_lower) and not re.search(r'\bnfs[\s\-_.]*e?\b', fn_lower):
                continue

            pdfs.append(os.path.join(root, fn))

    notas, sem_dados = [], []
    for pdf in pdfs:
        nome_arquivo = os.path.basename(pdf)
        # extrai número do nome do arquivo
        m_num = re.search(r"(\d+)", nome_arquivo)
        if not m_num:
            sem_dados.append(nome_arquivo)
            continue
        numero = m_num.group(1).lstrip("0")

        # extrai data/valor via OCR/texto (se falhar, insere None)
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
            linha = f"Nº: {it['numero']} | Data: {it['data']} | Valor: R$ {it['valor']}"
            if it.get("arquivo"): linha += f" | Arquivo: {it['arquivo']}"
            if it.get("esperado"): exp = it['esperado']; linha += f" | Esperado: Nº {exp['numero']}, {exp['data']}, R$ {exp['valor']}"
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
        # aceita "1.234,56" e "1234.56"
        s = s.replace('.', '').replace(',', '.')
        return Decimal(s).quantize(Decimal('0.01'))
    except (InvalidOperation, AttributeError, ValueError):
        return None

# === 5) Função principal chamada pelo Flask ===
def processar_comparacao_nf(zip_path, relatorio_pdf_path, output_dir):
    """
    Compara as NFS-e do RAR com o relatório PDF.
    - NÃO herda valor/data do relatório para a NF extraída (OCR).
    - Marca como 'encontrada' SOMENTE se (número, data e valor) do OCR coincidirem com o relatório.
    - Se não houver match exato, vai para 'divergentes' com 'esperado' (do relatório) só para exibição.
    - NF-e são ignoradas antes (filtradas em extrair_notas_zip / extrair_info_pdf).
    """
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

        # Match exato: precisa MESMA DATA e MESMO VALOR (sem herdar do relatório)
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
            # Não mexe em nf['data'] nem nf['valor']; usa relatório só para exibir o "esperado"
            nf['esperado'] = matches[0]
            divergentes.append(nf)

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

