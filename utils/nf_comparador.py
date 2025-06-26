import os
import rarfile
import datetime
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
from decimal import Decimal, InvalidOperation
import io
import re
import statistics
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

# caminho para o unrar e para o Tesseract
rarfile.UNRAR_TOOL = r"C:\Program Files\UnRAR.exe"
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
os.environ["TESSDATA_PREFIX"] = r"C:\Program Files\Tesseract-OCR\tessdata"

def extrair_info_pdf(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        if doc.is_encrypted:
            print(f"[DEBUG] Ignorando encriptado: {pdf_path}")
            return None

        # —— REINCLUA ISTO —— 
        texto = ""
        for page in doc:
            t = page.get_text()
            if t.strip():
                texto += t
            else:
                pix = page.get_pixmap(dpi=300)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                texto += pytesseract.image_to_string(img, lang='por')
        doc.close()
        
            # permite NFSe **ou** Prestação/Prestador de Serviço
        if not re.search(r"\bNFS[- ]?E\b", texto, re.IGNORECASE) \
        and not re.search(r"\bPrestação de Serviços\b", texto, re.IGNORECASE) \
        and not re.search(r"\bPrestador de Serviços\b", texto, re.IGNORECASE):
            print(f"[DEBUG] Ignorando sem padrão de NFSe ou Serviço: {pdf_path}")
            return None

        # número pelo nome de arquivo
        nome = os.path.basename(pdf_path)
        m = re.search(r"(\d+)", nome)
        numero = str(int(m.group(1))).lstrip("0") if m else None

        # data e valor pelo maior match
        dm = re.search(r"(\d{2}/\d{2}/\d{4})", texto)
        raw_vals = re.findall(r"\d{1,3}(?:\.\d{3})*,\d{2}", texto)
        valor = None
        if raw_vals:
            vf = max(Decimal(v.replace(".", "").replace(",", ".")) for v in raw_vals)
            valor = str(vf.quantize(Decimal("0.01")))

        return {"numero": numero, "data": dm.group(1) if dm else None, "valor": valor}
    except Exception as e:
        print(f"Erro extraindo {pdf_path}: {e}")
        return None

def extrair_notas_zip(zip_path, temp_dir):
    """Descompacta o RAR e aplica extrair_info_pdf em cada PDF."""
    os.makedirs(temp_dir, exist_ok=True)
    with rarfile.RarFile(zip_path) as rar:
        rar.extractall(temp_dir)
    notas = []
    for fn in os.listdir(temp_dir):
        if fn.lower().endswith(".pdf"):
            info = extrair_info_pdf(os.path.join(temp_dir, fn))
            if info and info.get("numero") and info.get("valor"):
                info["arquivo"] = fn
                notas.append(info)
    return notas

def extrair_relatorio(pdf_path):
    import fitz, re, statistics
    from decimal import Decimal, InvalidOperation

    # --- abre todo o PDF e coleta blocos para detectar cabeçalhos ---
    doc = fitz.open(pdf_path)
    header_blocks = []
    for p in doc:
        header_blocks.extend(p.get_text("blocks"))
    doc.close()

    # --- calcula médias X dos cabeçalhos ---
    xs_doc     = [b[0] for b in header_blocks if "docum"         in b[4].lower()]
    xs_date    = [b[0] for b in header_blocks if "entrada/saíd"  in b[4].lower()]
    xs_valor   = [b[0] for b in header_blocks if "valor contábil" in b[4].lower()]
    xs_especie = [b[0] for b in header_blocks if "espécie"       in b[4].lower()]

    if not (xs_doc and xs_date and xs_valor and xs_especie):
        print("[DEBUG #1] cabeçalhos faltando, usando fallback")
        rel = _fallback_extrair(pdf_path)
    else:
        doc_x     = statistics.mean(xs_doc)
        date_x    = statistics.mean(xs_date)
        valor_x   = statistics.mean(xs_valor)
        especie_x = statistics.mean(xs_especie)
        print(f"[DEBUG #1] posições X → doc:{doc_x:.1f}, date:{date_x:.1f}, valor:{valor_x:.1f}, espécie:{especie_x:.1f}")

        # --- monta linhas agrupando por y0 arredondado ---
        all_blocks = []
        for p in fitz.open(pdf_path):
            all_blocks.extend(p.get_text("blocks"))
        rows = {}
        for b in all_blocks:
            rows.setdefault(round(b[1],1), []).append(b)

        def pick_closest(row, x0, tol=25):
            cands = [b for b in row if abs(b[0]-x0) <= tol]
            return min(cands, key=lambda b: abs(b[0]-x0))[4].strip() if cands else ""

        rel = []
        # --- loop principal extração via colunas X ---
        for y0 in sorted(rows):
            row       = rows[y0]
            txt_num   = pick_closest(row, doc_x)
            txt_date  = pick_closest(row, date_x)
            txt_valor = pick_closest(row, valor_x)
            txt_esp   = pick_closest(row, especie_x)

            # --- DEBUG #2: veja o que capturou em cada linha ---
            print(f"[DEBUG #2] Linha y={y0}: num={txt_num!r}, date={txt_date!r}, valor={txt_valor!r}, esp={txt_esp!r}")

            if txt_esp.strip().upper() != "NFSE":
                continue

            m_num  = re.search(r"\b(\d+)\b",           txt_num)
            m_date = re.search(r"(\d{2}/\d{2}/\d{4})", txt_date)
            m_val  = re.search(r"(\d{1,3}(?:\.\d{3})*,\d{2})", txt_valor)
            if not (m_num and m_date and m_val):
                continue

            valor = str(Decimal(m_val.group(1).replace(".","").replace(",",".")) \
                        .quantize(Decimal("0.01")))
            rel.append({
                "numero": m_num.group(1).lstrip("0"),
                "data":   m_date.group(1),
                "valor":  valor
            })

        if not rel:
            # se nada saiu, fallback refinado
            rel = _fallback_extrair(pdf_path)

    # --- DEBUG FINAL: imprime todas as entradas extraídas ---
    print("\n[DEBUG] Entradas extraídas do relatório:")
    for entrada in rel:
        print(f"  - Nº: {entrada['numero']}, Data: {entrada['data']}, Valor: R$ {entrada['valor']}")

    return rel

def _fallback_extrair(pdf_path):
    import fitz, re
    from decimal import Decimal, InvalidOperation

    doc = fitz.open(pdf_path)
    rel = []
    for p in doc:
        for b in p.get_text("blocks"):
            txt = b[4]
            if "\n" not in txt:
                continue
            lines = [l.strip() for l in txt.splitlines() if l.strip()]
            if "NFSE" not in [l.upper() for l in lines]:
                continue

            # extrai primeiro número, data e valor válidos
            nums  = [l for l in lines if l.isdigit()]
            dates = [l for l in lines if re.match(r"\d{2}/\d{2}/\d{4}", l)]
            vals  = [l for l in lines if re.fullmatch(r"\d{1,3}(?:\.\d{3})*,\d{2}", l)]
            if not (nums and dates and vals):
                continue

            try:
                v = str(Decimal(vals[0].replace(".","").replace(",","."))
                        .quantize(Decimal("0.01")))
            except InvalidOperation:
                continue

            rel.append({
                "numero": nums[0].lstrip("0"),
                "data":   dates[0],
                "valor":  v
            })

    if rel:
        print(f"[DEBUG #3] Fallback refinado extraiu {len(rel)} entradas via split de blocos")
        # <<< dedupe aqui >>>
        seen = set()
        dedup = []
        for r in rel:
            key = (r['numero'], r['data'], r['valor'])
            if key not in seen:
                seen.add(key)
                dedup.append(r)
        print(f"[DEBUG] Após deduplicação: {len(dedup)} entradas únicas")
        return dedup

    print("[DEBUG #3] Nada encontrado nem pelo fallback")
    return []

def gerar_pdf_relatorio(encontradas, divergentes, nao_encontradas, caminho_saida):
    """Mantém sua lógica de geração de PDF igual de antes."""
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

    for nf in notas_zip:
        # **PASSO 2: validação da string antes de converter em int**
        num_str = nf.get("numero", "")
        if not num_str.isdigit():
            # se veio vazio ou com letras, já marca como não encontrada
            nao_encontradas.append(nf)
            continue

        num = int(num_str)

        # resto da sua lógica de comparação
        match = next(
            (r for r in relatorio
                if r.get("numero") and r["numero"].isdigit()
                and int(r["numero"]) == num),
            None
        )
        if not match:
            nao_encontradas.append(nf)
            continue

        if nf["data"] == match["data"] and Decimal(nf["valor"]) == Decimal(match["valor"]):
            encontradas.append(nf)
        else:
            nf["esperado"] = match
            divergentes.append(nf)

    # geração do PDF, etc...
    os.makedirs(output_dir, exist_ok=True)
    pdf = os.path.join(output_dir, "relatorio_validacao.pdf")
    gerar_pdf_relatorio(encontradas, divergentes, nao_encontradas, pdf)

    return {
        "encontradas": encontradas,
        "divergentes": divergentes,
        "nao_encontradas": nao_encontradas,
        "pdf": pdf
    }

def processar_comparacao_nf(zip_path, relatorio_pdf_path, output_dir):
    """Wrapper para o Flask."""
    temp = os.path.join(os.path.dirname(output_dir), "temp_notas")
    notas = extrair_notas_zip(zip_path, temp)
    rel = extrair_relatorio(relatorio_pdf_path)
    res = comparar_nfs(notas, rel, output_dir)
    pdf = res.pop("pdf")
    return res, pdf
