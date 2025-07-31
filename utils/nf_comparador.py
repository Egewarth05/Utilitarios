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
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from subprocess import run, DEVNULL

# aceita qualquer NFS-e (com ou sem hífen) ou o texto "Nota Fiscal de Serviços"
PATTERN_NFSE = re.compile(r'(NFS[–—-]?E|Nota\s+Fiscal\b)', re.IGNORECASE)

# configurações externas
rarfile.UNRAR_TOOL = r"C:\Program Files\UnRAR.exe"
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
os.environ["TESSDATA_PREFIX"] = r"C:\Program Files\Tesseract-OCR\tessdata"


def extrair_info_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    nat = ""
    ocr_text = []
    for page in doc:
        nat += page.get_text() or ""
        pix = page.get_pixmap(dpi=300, alpha=False)
        img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L")
        bw = img.point(lambda x: 0 if x < 180 else 255, "1")
        cfg = r"--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789,\/"
        ocr_text.append(pytesseract.image_to_string(bw, lang="por", config=cfg))
    doc.close()
    texto = nat + "\n\n" + "\n\n".join(ocr_text)

    if re.search(r'\bNFE\b', texto, re.IGNORECASE) and not PATTERN_NFSE.search(texto):
        return None

    texto_sem_tributos = re.split(r"(?i)Detalhamento\s+dos\s+Tributos", texto)[0]
    header_section = re.split(r"(?i)discriminação", texto_sem_tributos)[0]

    debug_path = os.path.join(tempfile.gettempdir(), "header_section.txt")
    with open(debug_path, "w", encoding="utf-8") as f:
        f.write(header_section)
    print(f"[DEBUG] header_section salvo em: {debug_path}")

    nome = os.path.basename(pdf_path)
    m = re.search(r"(\d+)", nome)
    raw = m.group(1) if m else ""
    parts = re.split(r"0{4,}", raw)
    numero = parts[-1].lstrip("0") if len(parts) > 1 and parts[-1] else raw.lstrip("0")

    data = None
    m = re.search(r"(\d{2}/\d{2}/\d{4})\s+\d{2}:\d{2}:\d{2}", texto_sem_tributos)
    if m:
        data = m.group(1)
    else:
        m = re.search(r"Data\s+Emissão\D*?(\d{2}/\d{2}/\d{4})", texto, re.IGNORECASE)
        if m:
            data = m.group(1)

    valor_str = None
    for pat in [
        r"Valor\s+Total\s+da\s+NFS[–—-]?E[:\s]*R\$\s*([\d\.]+,\d{2})",
        r"([\d\.]+,\d{2})\s*Valor\s+Bruto\s+da\s+Nota"
    ]:
        m = re.search(pat, texto, re.IGNORECASE)
        if m:
            valor_str = m.group(1)
            break
    if not valor_str:
        for line in header_section.splitlines():
            if re.search(r"\bValor\b", line, re.IGNORECASE):
                m = re.search(r"([\d\.]+,\d{2})", line)
                if m:
                    valor_str = m.group(1)
                    break

    valor = None
    if valor_str:
        valor = str(
            Decimal(valor_str.replace(".", "").replace(",", "."))
                   .quantize(Decimal("0.01"))
        )

    print(f"[DEBUG][NF] {os.path.basename(pdf_path)} → número: {numero!r}, data: {data!r}, valor: {valor!r}")
    return {"numero": numero, "data": data, "valor": valor}


def extrair_notas_zip(zip_path, temp_dir):
    os.makedirs(temp_dir, exist_ok=True)
    with rarfile.RarFile(zip_path) as rar:
        rar.extractall(temp_dir)

    pdfs = [
        os.path.join(temp_dir, fn)
        for fn in os.listdir(temp_dir)
        if fn.lower().endswith('.pdf') and 'fatura' not in fn.lower()
    ]
    notas, sem_dados = [], []
    for pdf in pdfs:
        print(f"\n--- PROCESSANDO {os.path.basename(pdf)} ---")
        info = extrair_info_pdf(pdf)
        if info and info.get("numero") and info.get("valor"):
            info["arquivo"] = os.path.basename(pdf)
            notas.append(info)
        else:
            sem_dados.append(os.path.basename(pdf))

    if sem_dados:
        print(f"[DEBUG][SEM_DADOS] Arquivos sem extração ({len(sem_dados)}): {sem_dados}")
    return notas, sem_dados


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
    if not rows:
        raise ValueError("Nenhuma linha extraída do relatório.")

    header = [c.replace("\n", " ").strip().lower() for c in rows[0]]
    idx_doc     = next(i for i,h in enumerate(header) if "docum"   in h)
    idx_especie = next(i for i,h in enumerate(header) if "espécie" in h)
    idx_date    = next(i for i,h in enumerate(header) if "entrada" in h)
    idx_valor   = next(i for i,h in enumerate(header) if "valor"   in h)

    rel = []
    for row in rows[1:]:
        esp = row[idx_especie].strip().upper()
        if esp == "NFE" or not esp.startswith("NFS"):
            continue
        numero  = row[idx_doc].strip()
        data    = row[idx_date].strip()
        raw_val = row[idx_valor].strip()
        try:
            valor = str(Decimal(raw_val.replace(".", "").replace(",", "."))\
                        .quantize(Decimal("0.01")))
        except InvalidOperation:
            continue
        rel.append({"numero": numero, "data": data, "valor": valor})

    return rel

def gerar_pdf_relatorio(encontradas, divergentes, nao_encontradas, sem_dados, caminho_saida):
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
                c.showPage()
                y = height - 50
            linha = f"Nº: {it['numero']} | Data: {it['data']} | Valor: R$ {it['valor']}"
            # **agora inclui nome do arquivo**
            if it.get("arquivo"):
                linha += f" | Arquivo: {it['arquivo']}"
            if "esperado" in it:
                exp = it["esperado"]
                linha += f" | Esperado: Nº {exp['numero']}, {exp['data']}, R$ {exp['valor']}"
            c.drawString(50, y, linha)
            y -= 15
        y -= 20

    # cabeçalho
    c.setFont("Helvetica-Bold", 14)
    c.drawString(40, y, "Relatório de Validação de NFS-e")
    y -= 30

    # seções
    add_secao("✔ Encontradas e Correspondentes", encontradas)
    add_secao("⚠ Divergentes (data ou valor)", divergentes)
    add_secao("❌ Não encontradas no relatório", nao_encontradas)

    # arquivos sem extração
    c.setFont("Helvetica-Bold", 12)
    c.drawString(40, y, f"🛈 Arquivos sem extração ({len(sem_dados)}):")
    y -= 20
    c.setFont("Helvetica", 10)
    if not sem_dados:
        c.drawString(50, y, "Nenhum arquivo sem extração.")
        y -= 30
    else:
        for fname in sem_dados:
            if y < 50:
                c.showPage()
                y = height - 50
            c.drawString(50, y, fname)
            y -= 15
        y -= 20

    c.save()

def comparar_nfs(notas_zip, relatorio, sem_dados, output_dir):
    encontradas, divergentes, nao_encontradas = [], [], []

    idx = defaultdict(list)
    for r in relatorio:
        if r.get("numero", "").isdigit():
            idx[int(r["numero"])].append(r)

    for nf in notas_zip:
        num_str = nf.get("numero", "")
        if not num_str.isdigit():
            nao_encontradas.append(nf)
            continue
        num = int(num_str)
        nf_data = nf.get("data")
        nf_valor = Decimal(nf["valor"].replace(",", "."))

        matches = [r for r in relatorio if r.get("numero", "").isdigit() and int(r["numero"]) == num]
        if not matches:
            nao_encontradas.append(nf)
            continue

        if len(matches) > 1:
            exact = next((r for r in matches if r["data"] == nf_data and Decimal(r["valor"]) == nf_valor), None)
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
        data_rel = r["data"]

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
    gerar_pdf_relatorio(encontradas, divergentes, nao_encontradas, sem_dados, pdf_out)

    return {
        "encontradas": encontradas,
        "divergentes": divergentes,
        "nao_encontradas": nao_encontradas,
        "sem_dados": sem_dados,
        "pdf": pdf_out
    }

# compatibilidade
extrair_relatorio = extrair_relatorio_com_pdfplumber

def processar_comparacao_nf(zip_path, relatorio_pdf_path, output_dir):
    temp = os.path.join(os.path.dirname(output_dir), "temp_notas")
    if os.path.isdir(temp):
        shutil.rmtree(temp, onerror=lambda f,p,e: os.chmod(p, stat.S_IWRITE) or f(p))
    if os.path.isdir(output_dir):
        shutil.rmtree(output_dir, onerror=lambda f,p,e: os.chmod(p, stat.S_IWRITE) or f(p))

    os.makedirs(temp, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    notas, sem_dados = extrair_notas_zip(zip_path, temp)
    rel            = extrair_relatorio_com_pdfplumber(relatorio_pdf_path)

    resultado = comparar_nfs(notas, rel, sem_dados, output_dir)
    resultado["sem_dados"] = sem_dados
    notas.sort(key=lambda n: int(n["numero"]))
    rel.sort(key=lambda r: int(r["numero"]))
    resultado["encontradas"].sort(key=lambda n: int(n["numero"]))
    resultado["divergentes"].sort(key=lambda n: int(n["numero"]))
    resultado["nao_encontradas"].sort(key=lambda n: int(n["numero"]))

    pdf_path  = resultado.pop("pdf")
    return resultado, pdf_path
