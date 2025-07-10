import os
import io
import re
import stat
import shutil
import datetime
import statistics
import concurrent.futures
import subprocess
from subprocess import run, DEVNULL
import tempfile
from concurrent.futures import ProcessPoolExecutor
import rarfile
from collections import defaultdict
from decimal import Decimal, InvalidOperation
import fitz  # PyMuPDF
import pytesseract
import pdfplumber
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

# aceita qualquer NFS-e (com ou sem hífen) ou o texto "Nota Fiscal de Serviços"
PATTERN_NFSE = re.compile(r'(NFS[–—-]?E|Nota\s+Fiscal\b)', re.IGNORECASE)
# Configurações externas
rarfile.UNRAR_TOOL = r"C:\Program Files\UnRAR.exe"
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
os.environ["TESSDATA_PREFIX"] = r"C:\Program Files\Tesseract-OCR\tessdata"

def pdf_page_to_png_with_gs(pdf_path, page_number=1, dpi=300):
    """
    Usa Ghostscript para renderizar a página `page_number` do PDF em PNG
    e retorna o caminho do .png gerado.
    """
    tmp = tempfile.gettempdir()
    out = os.path.join(tmp, f"gs_page{page_number}.png")
    args = [
        "gswin64c", "-dSAFER", "-dBATCH", "-dNOPAUSE",
        f"-r{dpi}", "-sDEVICE=pngalpha",
        f"-dFirstPage={page_number}", f"-dLastPage={page_number}",
        f"-sOutputFile={out}", pdf_path
    ]
    run(args, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return out

def ocr_from_png(png_path):
    try:
        img = Image.open(png_path).convert("L")
    except Exception as e:
        print(f"❌ Falha ao abrir {png_path}: {e}")
        raise
    bw = img.point(lambda x: 0 if x < 180 else 255, "1")
    """
    Carrega o PNG, aplica binarização simples e devolve o texto OCR.
    """
    img = Image.open(png_path).convert("L")
    bw  = img.point(lambda x: 0 if x < 180 else 255, "1")
    cfg = r"--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789,\.\/"
    return pytesseract.image_to_string(bw, lang="por", config=cfg)

def extrair_info_pdf(pdf_path):
    # 1) abre PDF e extrai texto "nativo"
    doc = fitz.open(pdf_path)
    nat = ""
    ocr_text = []
    for page in doc:
        nat += page.get_text() or ""

        # renderiza com PyMuPDF em vez de Ghostscript
        pix = page.get_pixmap(dpi=300, alpha=False)
        img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L")
        bw  = img.point(lambda x: 0 if x < 180 else 255, "1")
        cfg = r"--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789,\/"
        ocr_text.append(pytesseract.image_to_string(bw, lang="por", config=cfg))

    doc.close()
    # monta o texto completo
    texto = nat + "\n\n" + "\n\n".join(ocr_text)

    # rejeita NFe puro que não seja NFSe
    if re.search(r'\bNFE\b', texto, re.IGNORECASE) and not re.search(r'NFS[–—-]?E', texto, re.IGNORECASE):
        return None
    # garante que é alguma Nota Fiscal
    if not PATTERN_NFSE.search(texto):
        return None

    # 1) limite o escopo até antes de “Detalhamento dos Tributos”
    texto_sem_tributos = re.split(r"(?i)Detalhamento\s+dos\s+Tributos", texto)[0]
    # 2) continue usando discriminação para dividir header x discriminação
    header_section = re.split(r"(?i)discriminação", texto_sem_tributos)[0]
    print("---- RAW HEADER_SECTION ----")
    print(repr(header_section))
    print("---- LINHAS ----")
    for i, l in enumerate(header_section.splitlines()):
        print(f"{i:02d}: {l!r}")
    header_clean   = re.sub(r"\s+", " ", header_section).strip()

    # extrai número da NF pelo nome do arquivo
    nome = os.path.basename(pdf_path)
    m = re.search(r"(\d+)", nome)
    raw = m.group(1) if m else ""
    # se raw tiver 4 ou mais zeros seguidos, pega só o que vem depois
    parts = re.split(r"0{4,}", raw)
    if len(parts) > 1 and parts[-1]:
        numero = parts[-1].lstrip("0")
    else:
        numero = raw.lstrip("0")

     # ===== EXTRAÇÃO DA DATA =====
    data = None

    # tenta data+hora em qualquer lugar do texto_sem_tributos
    m = re.search(r"(\d{2}/\d{2}/\d{4})\s+\d{2}:\d{2}:\d{2}",
                texto_sem_tributos)
    if m:
        data = m.group(1)

    # se ainda não achou, tenta outras variações...
    if not data:
        m = re.search(r"Data\s+Emissão\D*?(\d{2}/\d{2}/\d{4})",
                    texto, re.IGNORECASE)
        if m:
            data = m.group(1)

    # C) “DD/MM/AAAA HH:MM:SS” sozinho em alguma linha do header
    if not data:
        for line in header_section.splitlines():
            m = re.match(r"\s*(\d{2}/\d{2}/\d{4})\s+\d{2}:\d{2}:\d{2}", line)
            if m:
                data = m.group(1)
                break

    # D) Data da emissão (variação textual)
    if not data:
        m = re.search(r"Data\s+da\s+emiss(?:ão|ao)\D*?(\d{2}/\d{2}/\d{4})",
                    texto, re.IGNORECASE)
        if m:
            data = m.group(1)
        # 5) Fallback 1: primeira data no header (ignorando “venc”)
    if not data:
        for mm in re.finditer(r"(\d{2}/\d{2}/\d{4})", header_section):
            ctx = header_section[max(0, mm.start()-30): mm.end()+30]
            if re.search(r"venc", ctx, re.IGNORECASE):
                continue
            data = mm.group(1)
            break

    # 6) Fallback 2: última tentativa em todo o texto
    if not data:
        for mm in re.finditer(r"(\d{2}/\d{2}/\d{4})", texto):
            ctx = texto[max(0, mm.start()-30): mm.start()].lower()
            if re.search(r"venc", ctx):
                continue
            data = mm.group(1)
            break

    # valida ano mínimo 2025
    if data:
        try:
            dt = datetime.datetime.strptime(data, "%d/%m/%Y")
            if dt.year < 2025:
                data = None
        except ValueError:
            data = None
            
    # ========== EXTRAÇÃO DO VALOR TOTAL ==========
    valor_str = None

    # 1) Valor Total da NFS-e (com R$ e dois-pontos)
    m = re.search(
        r"Valor\s+Total\s+da\s+NFS[–—-]?e[:\s]*R\$\s*(\d{1,3}(?:\.\d{3})*,\d{2})",
        texto, re.IGNORECASE
    )
    if m:
        valor_str = m.group(1)
    else:
        # 2) Valor Bruto da Nota (caso exista)
        m = re.search(
            r"(\d{1,3}(?:\.\d{3})*,\d{2})\s*Valor\s+Bruto\s+da\s+Nota[:\s]",
            texto, re.IGNORECASE
        )
        if m:
            valor_str = m.group(1)
        else:
            # 3) Valor Total do RPS
            m = re.search(
                r"Valor\s+Total\s+do\s+RPS\D*?R\$\s*(\d{1,3}(?:\.\d{3})*,\d{2})",
                texto, re.IGNORECASE
            )
            if m:
                valor_str = m.group(1)        
                
            elif not valor_str:
                m = re.search(
                    r"FATURA\s*/\s*DUPLICATA.*?^\s*\d+\s+\d{2}/\d{2}/\d{4}\s+([\d\.]+,\d{2})",
                    texto, re.IGNORECASE | re.MULTILINE | re.DOTALL
                )
                if m:
                    valor_str = m.group(1)        
                
        # 4) Valor da Nota (demonstrativo)
        if not valor_str:
            m = re.search(
                r"Valor\s+da\s+Nota[:\s]*R\$\s*([\d\.]+,\d{2})",
                texto, re.IGNORECASE
            )
            if m:
                valor_str = m.group(1)

        # 5) fallback genérico: escolhe o maior R$ encontrado (ignora “imposto”)
        if not valor_str:
            candidates = []
            for line in texto.splitlines():
                if re.search(r"R\$\s*[\d\.]+,\d{2}", line, re.IGNORECASE) \
                and not re.search(r"imposto|iss", line, re.IGNORECASE):
                    mm = re.search(r"R\$\s*([\d\.]+,\d{2})", line)
                    if mm:
                        candidates.append(mm.group(1))
            if candidates:
                valor_str = max(
                    candidates,
                    key=lambda v: Decimal(v.replace(".", "").replace(",", "."))
                )

       # converte e retorna (agora no nível da função, não dentro do if)
    valor = None
    if valor_str:
        valor = str(Decimal(valor_str.replace(".", "").replace(",", "."))
                    .quantize(Decimal("0.01")))

    info = {"numero": numero, "data": data, "valor": valor}
    print(f"[DEBUG][NF] {os.path.basename(pdf_path)} → "
          f"número: {numero!r}, data: {data!r}, valor: {valor!r}")
    return info

def extrair_notas_zip(zip_path, temp_dir):
    os.makedirs(temp_dir, exist_ok=True)
    with rarfile.RarFile(zip_path) as rar:
        rar.extractall(temp_dir)
    
    # usa ProcessPoolExecutor para rodar extrair_info_pdf em paralelo
    pdfs = [
        os.path.join(temp_dir, fn)
        for fn in os.listdir(temp_dir)
        if fn.lower().endswith('.pdf') and 'fatura' not in fn.lower()
    ]
    notas, sem_dados = [], []
    with ProcessPoolExecutor(max_workers=os.cpu_count()) as pool:
        futures = { pool.submit(extrair_info_pdf, pdf): pdf for pdf in pdfs }
        for fut in concurrent.futures.as_completed(futures):
            pdf = futures[fut]
            fn = os.path.basename(pdf)
            info = fut.result()
            if info and info.get("numero") and info.get("valor"):
                info["arquivo"] = fn
                notas.append(info)
            else:
                sem_dados.append(fn)
    if sem_dados:
          print(f"[DEBUG][SEM_DADOS] Arquivos sem extração ({len(sem_dados)}): {sem_dados}")
    return notas

def extrair_relatorio_com_pdfplumber(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        rows = []
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                for row in table:
                    if any(cell not in (None, "") for cell in row):
                        rows.append([cell or "" for cell in row])

    print(f"[DEBUG][RELATÓRIO] {os.path.basename(pdf_path)} → {len(rows)} linhas extraídas")
    # DEBUG: mostre cabeçalho e primeiras 5 linhas
    header = [c.replace("\n", " ").strip() for c in rows[0]]
    sample = rows[1:6]
    print(f"[DEBUG][RELATÓRIO] Cabeçalho: {header}")
    print(f"[DEBUG][RELATÓRIO] Primeiras linhas: {sample}")
    if not rows:
        raise ValueError("Nenhuma linha extraída do relatório.")

    header      = [c.replace("\n", " ").strip().lower() for c in rows[0]]
    idx_doc     = next(i for i,h in enumerate(header) if "docum"   in h)
    idx_especie = next(i for i,h in enumerate(header) if "espécie" in h)
    idx_date    = next(i for i,h in enumerate(header) if "entrada" in h)
    idx_valor   = next(i for i,h in enumerate(header) if "valor"   in h)

    rel = []
    for row in rows[1:]:
        esp = row[idx_especie].strip().upper()
        if esp == "NFE":
            continue
        if not esp.startswith("NFS"):
            continue

        numero  = row[idx_doc].strip()
        data    = row[idx_date].strip()
        raw_val = row[idx_valor].strip()
        try:
            valor = str(Decimal(raw_val.replace(".", "").replace(",", "."))
                        .quantize(Decimal("0.01")))
        except InvalidOperation:
            continue

        rel.append({"numero": numero, "data": data, "valor": valor})

    return rel

def _fallback_extrair(pdf_path):
    import fitz
    rel = []
    for p in fitz.open(pdf_path):
        for b in p.get_text("blocks"):
            txt = b[4]
            if "\n" not in txt:
                continue
            lines = [l.strip() for l in txt.splitlines() if l.strip()]
            if "NFSE" not in [l.upper() for l in lines]:
                continue

            nums  = [l for l in lines if l.isdigit()]
            dates = [l for l in lines if re.match(r"\d{2}/\d{2}/\d{4}", l)]
            vals  = [l for l in lines if re.fullmatch(r"\d{1,3}(?:\.\d{3})*,\d{2}", l)]
            if not (nums and dates and vals):
                continue
            try:
                v = str(Decimal(vals[0].replace(".", "").replace(",", "."))
                        .quantize(Decimal("0.01")))
            except InvalidOperation:
                continue

            rel.append({"numero": nums[0].lstrip("0"), "data": dates[0], "valor": v})

    # dedupe
    seen = set()
    uniq = []
    for r in rel:
        key = (r["numero"], r["data"], r["valor"])
        if key not in seen:
            seen.add(key)
            uniq.append(r)
    return uniq


def gerar_pdf_relatorio(encontradas, divergentes, nao_encontradas, caminho_saida):
    c = canvas.Canvas(caminho_saida, pagesize=A4)
    width, height = A4
    y = height - 50

    def add_secao(titulo, itens):
        nonlocal y
        c.setFont("Helvetica-Bold", 12)
        c.drawString(40, y, f"{titulo} ({len(itens)}):")
        y -= 20
        c.setFont("Helvetica", 10)
        if not itens:
            c.drawString(50, y, "Nenhum item.")
            y -= 30
            return
        for it in itens:
            if y < 50:
                c.showPage(); y = height - 50
            linha = f"Nº: {it['numero']} | Data: {it['data']} | Valor: R$ {it['valor']}"
            if "esperado" in it:
                exp = it["esperado"]
                linha += f" | Esperado: Nº {exp['numero']}, {exp['data']}, R$ {exp['valor']}"
            c.drawString(50, y, linha)
            y -= 15
        y -= 20

    c.setFont("Helvetica-Bold", 14)
    c.drawString(40, y, "Relatório de Validação de NFS-e")
    y -= 30
    add_secao("✔ Encontradas e Correspondentes", encontradas)
    add_secao("⚠ Divergentes (data ou valor)", divergentes)
    add_secao("❌ Não encontradas no relatório", nao_encontradas)
    c.save()


def comparar_nfs(notas_zip, relatorio, output_dir):
    encontradas, divergentes, nao_encontradas = [], [], []

    idx = defaultdict(list)
    for r in relatorio:
        if r.get("numero","").isdigit():
            idx[int(r["numero"])].append(r)

    for nf in notas_zip:
        num_str = nf.get("numero", "")
        if not num_str.isdigit():
            nao_encontradas.append(nf)
            continue
        num      = int(num_str)
        nf_data  = nf.get("data")
        nf_valor = Decimal(nf["valor"].replace(",", "."))

        matches = [
            r for r in relatorio
            if r.get("numero","").isdigit() and int(r["numero"]) == num
        ]
        if not matches:
            nao_encontradas.append(nf)
            continue

        if len(matches) > 1:
            exact = next(
                (r for r in matches if r["data"] == nf_data and Decimal(r["valor"]) == nf_valor),
                None
            )
            if exact:
                encontradas.append(nf)
                continue
            same_val = next((r for r in matches if Decimal(r["valor"]) == nf_valor), None)
            if same_val:
                nf["esperado"] = same_val
                divergentes.append(nf)
            else:
                nao_encontradas.append(nf)
            continue

        r = matches[0]
        valor_rel = Decimal(r["valor"])
        data_rel  = r["data"]

        if valor_rel == nf_valor:
            if data_rel == nf_data or nf_data is None:
                encontradas.append(nf)
            else:
                nf["esperado"] = r
                divergentes.append(nf)
        else:
            nao_encontradas.append(nf)

    os.makedirs(output_dir, exist_ok=True)
    pdf_out = os.path.join(output_dir, "relatorio_validacao.pdf")
    gerar_pdf_relatorio(encontradas, divergentes, nao_encontradas, pdf_out)

    return {
        "encontradas": encontradas,
        "divergentes": divergentes,
        "nao_encontradas": nao_encontradas,
        "pdf": pdf_out
    }

# para compatibilidade com quem importava extrair_relatorio
extrair_relatorio = extrair_relatorio_com_pdfplumber

def processar_comparacao_nf(zip_path, relatorio_pdf_path, output_dir):
    temp = os.path.join(os.path.dirname(output_dir), "temp_notas")
    if os.path.isdir(temp):
        shutil.rmtree(temp, onerror=lambda f,p,e: os.chmod(p, stat.S_IWRITE) or f(p))
    if os.path.isdir(output_dir):
        shutil.rmtree(output_dir, onerror=lambda f,p,e: os.chmod(p, stat.S_IWRITE) or f(p))

    os.makedirs(temp, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    notas = extrair_notas_zip(zip_path, temp)
    rel   = extrair_relatorio_com_pdfplumber(relatorio_pdf_path)

    resultado = comparar_nfs(notas, rel, output_dir)
    notas.sort(key=lambda n: int(n["numero"]))
    rel.sort(key=lambda r: int(r["numero"]))
    resultado["encontradas"].sort(key=lambda n: int(n["numero"]))
    resultado["divergentes"].sort(key=lambda n: int(n["numero"]))
    resultado["nao_encontradas"].sort(key=lambda n: int(n["numero"]))

    pdf_path  = resultado.pop("pdf")
    return resultado, pdf_path
