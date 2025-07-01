import os
import rarfile
import shutil
import datetime
import stat
from collections import defaultdict
import fitz  # PyMuPDF
import pytesseract
import pdfplumber
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
    doc   = fitz.open(pdf_path)
    texto = ""
    for page in doc:
        t = page.get_text()
        if t.strip():
            texto += t
        else:
            pix   = page.get_pixmap(dpi=300)
            img   = Image.open(io.BytesIO(pix.tobytes("png")))
            texto += pytesseract.image_to_string(img, lang='por')
    doc.close()

    # --- 0) filtra só NFS-e (não NFe) ---
    if not re.search(r"\bNFS[- ]?E\b", texto, re.IGNORECASE):
        print(f"[DEBUG][PDF] Pulando {os.path.basename(pdf_path)} (não é NFS-e)")
        return None

    # --- 1) número da NF pelo nome do arquivo ---
    m_num  = re.search(r"(\d+)", os.path.basename(pdf_path))
    numero = str(int(m_num.group(1))).lstrip("0") if m_num else None

    # --- 2) isola até “DISCRIMINAÇÃO” e normaliza espaços ---
    header_section = re.split(r"(?i)discriminação", texto)[0]
    header_clean   = re.sub(r"\s+", " ", header_section).strip()
    print(f"[DEBUG][PDF] {os.path.basename(pdf_path)} → {header_clean[:200]!r}")

    # --- 3) extrai data: primeiro label específico ---
    m_label = re.search(
        r"Data e Hora de Emissão\D*(\d{2}/\d{2}/\d{4})",
        header_clean, re.IGNORECASE
    )
    if m_label:
        data = m_label.group(1)
    else:
        # fallback: primeira data no texto que não esteja perto de “Vencimento”
        data = None
        for mm in re.finditer(r"(\d{2}/\d{2}/\d{4})", texto):
            ctx = texto[max(0, mm.start()-30): mm.start()]
            if not re.search(r"venci", ctx, re.IGNORECASE):
                data = mm.group(1)
                break

    # --- 4) extrai valor (maior ocorrência) ---
    raw_vals = re.findall(r"\d{1,3}(?:\.\d{3})*,\d{2}", texto)
    valor    = None
    if raw_vals:
        vf    = max(Decimal(v.replace(".", "").replace(",", ".")) for v in raw_vals)
        valor = str(vf.quantize(Decimal("0.01")))

    return {"numero": numero, "data": data, "valor": valor}

def extrair_relatorio_com_pdfplumber(pdf_path):
    # extrai todas as linhas não-brancas
    with pdfplumber.open(pdf_path) as pdf:
        rows = []
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                for row in table:
                    if any(cell not in (None, "") for cell in row):
                        rows.append([cell or "" for cell in row])

    print(f"[DEBUG][RELATÓRIO] {os.path.basename(pdf_path)} → {len(rows)} linhas")
    if not rows:
        raise ValueError("Nenhuma linha extraída do relatório.")

    # identifica cabeçalho e índices
    header = [c.replace("\n", " ").strip().lower() for c in rows[0]]
    idx_doc     = next(i for i,h in enumerate(header) if "docum" in h)
    idx_especie = next(i for i,h in enumerate(header) if "espécie" in h)
    idx_date    = next(i for i,h in enumerate(header) if "entrada" in h)
    idx_valor   = next(i for i,h in enumerate(header) if "valor" in h)

    # coleta apenas linhas de NFSe
    rel = []
    for row in rows[1:]:
        if row[idx_especie].strip().upper() != "NFSE":
            continue
        numero  = row[idx_doc].strip()
        data    = row[idx_date].strip()
        raw_val = row[idx_valor].strip()
        try:
            valor = str(Decimal(raw_val.replace(".", "").replace(",", ".")).quantize(Decimal("0.01")))
        except InvalidOperation:
            continue
        rel.append({"numero": numero, "data": data, "valor": valor})

    return rel

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

def extrair_relatorio_com_pdfplumber(pdf_path):
    rel = []
    print(f"[DEBUG] Abrindo relatório: {pdf_path}") # Nova linha de depuração
    with pdfplumber.open(pdf_path) as pdf:
        rows = []
        for i, page in enumerate(pdf.pages): # Adicionado 'i' para número da página
            print(f"[DEBUG] Processando página {i+1} de {pdf_path}") # Nova linha de depuração
            table = page.extract_table()
            if not table:
                print(f"[DEBUG] Nenhuma tabela encontrada na página {i+1}.") # Nova linha de depuração
                continue
            
            # Nova linha de depuração: Mostra a tabela bruta extraída por página
            print(f"[DEBUG] Tabela bruta da página {i+1}:\n{table}") 

            # remove linhas totalmente em branco
            for row in table:
                if any(cell not in (None, '') for cell in row):
                    rows.append(row)
    
    # Nova linha de depuração: Mostra todas as linhas acumuladas após limpeza
    print(f"[DEBUG] Todas as linhas extraídas (após remoção de brancos):\n{rows}")

    if not rows:
        raise ValueError("Não foi possível extrair nenhuma linha de nenhum página.")

    # 2) Identifica cabeçalho real na primeira linha não-vazia
    header_row = rows[0]
    headers = [
        (cell or '').replace('\n', ' ').strip().lower()
        for cell in header_row
    ]
    print("[DEBUG cabeçalhos normalizados]", headers)

    # 3) Mapeia índices (seu código atual permanece)
    try:
        idx_doc     = next(i for i,h in enumerate(headers) if 'docum'   in h or 'documento' in h)
        idx_especie = next(i for i,h in enumerate(headers) if 'espécie' in h)
        idx_date    = next(i for i,h in enumerate(headers) if 'entrada' in h)
        idx_valor   = next(i for i,h in enumerate(headers) if 'valor'   in h)
        print(f"[DEBUG] Índices de colunas: Doc={idx_doc}, Espécie={idx_especie}, Data={idx_date}, Valor={idx_valor}") # Nova linha de depuração
    except StopIteration as e:
        raise ValueError(f"Cabeçalho não encontrado: {e}")

    # 4) Percorre todas as linhas de dados (a partir da segunda)
    for row_idx, row in enumerate(rows[1:]): # Adicionado 'row_idx' para depuração
        especie = (row[idx_especie] or '').strip().upper()
        numero  = (row[idx_doc]  or '').strip()
        data    = (row[idx_date] or '').strip()
        raw_val = (row[idx_valor] or '').strip()

        print(f"[DEBUG] Linha {row_idx+2} (após cabeçalho): Espécie='{especie}', Número='{numero}', Data='{data}', Valor Bruto='{raw_val}'") # Nova linha de depuração

        if especie != 'NFSE':
            print(f"[DEBUG] Ignorando linha {row_idx+2}: Não é NFSE.") # Nova linha de depuração
            continue

        try:
            valor = Decimal(raw_val.replace('.', '').replace(',', '.'))
        except InvalidOperation: # Usar InvalidOperation para ser mais específico
            print(f"[DEBUG] Erro de conversão de valor na linha {row_idx+2}: '{raw_val}' não é um valor válido.") # Nova linha de depuração
            continue
        except Exception as e: # Captura outros erros de conversão
            print(f"[DEBUG] Erro inesperado na conversão de valor da linha {row_idx+2}: {e} para '{raw_val}'.")
            continue

        if numero and data and valor:
            print(f"[DEBUG] Adicionando ao relatório: {{'numero': '{numero}', 'data': '{data}', 'valor': '{valor}'}}") # Nova linha de depuração
            rel.append({'numero': numero, 'data': data, 'valor': valor})
        else:
            print(f"[DEBUG] Linha {row_idx+2} incompleta: numero={numero}, data={data}, valor={valor}. Ignorando.") # Nova linha de depuração

    return rel

def extrair_relatorio(pdf_path):
    doc  = fitz.open(pdf_path)
    page = doc[0]
    words = page.get_text("words")   # (x0,y0,x1,y1, word, bno, lno, wno)

    # 1) agrupa por linha
    linhas = defaultdict(list)
    for x0, y0, x1, y1, w, bno, lno, wno in words:
        linhas[lno].append((w.lower(), x0))

    # 2) para cada coluna, ache o lno e o x0 médio
    def achar_x(header_kw):
        # devolve média de todos x0 cujo word contenha header_kw
        matches = [(lno, x) 
                   for lno, ws in linhas.items() 
                   for w,x in ws if header_kw in w]
        if not matches:
            raise ValueError(f"Não achei header contendo '{header_kw}'")
        # pega só os x0
        xs = [x for _,x in matches]
        return statistics.mean(xs)

    doc_x     = achar_x("docum")        # vai achar “Docum” e “ento”
    valor_x   = achar_x("valor")        # acha “valor” em “Valor Contábil”
    date_x    = achar_x("entrada")      # acha “Entrada/Saíd”
    especie_x = achar_x("espécie")      # acha “Espécie”

    print(f"[DEBUG #1] X cols → doc:{doc_x:.1f}, valor:{valor_x:.1f}, "
          f"date:{date_x:.1f}, esp:{especie_x:.1f}")

    # 3) monta rows a partir de blocks e usa pick_closest como antes
    all_blocks = page.get_text("blocks")
    rows = {}
    for x0,y0,x1,y1,txt, *_ in all_blocks:
        key = round(y0,1)
        rows.setdefault(key, []).append((x0, txt.strip()))
        
    col_positions = sorted([doc_x, valor_x, date_x, especie_x])
    tol = min(b - a for a, b in zip(col_positions, col_positions[1:])) / 2
    print(f"[DEBUG] tolerância dinâmica: {tol:.1f}px")    

    def pick_closest(row, x_ref, tol=tol):
        cands = [(x,txt) for x,txt in row if abs(x - x_ref) <= tol]
        return min(cands, key=lambda t: abs(t[0] - x_ref))[1] if cands else ""

    rel = []
    for y in sorted(rows):
        row   = rows[y]
        num   = pick_closest(row, doc_x)
        val   = pick_closest(row, valor_x)
        date  = pick_closest(row, date_x)
        esp   = pick_closest(row, especie_x)
        print(f"[DEBUG #2] y={y}: num={num!r}, valor={val!r}, date={date!r}, esp={esp!r}")

        if esp.upper() != "NFSE":
            continue
        m_num  = re.search(r"\b(\d+)\b",           num)
        m_date = re.search(r"\d{2}/\d{2}/\d{4}",  date)
        m_val  = re.search(r"\d{1,3}(?:\.\d{3})*,\d{2}", val)
        if m_num and m_date and m_val:
            rel.append({
                "numero": m_num.group(1).lstrip("0"),
                "data":   m_date.group(0),
                "valor":  str(Decimal(m_val.group(0).replace(".","").replace(",",".")))
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
        num_str = nf.get("numero", "")
        if not num_str.isdigit():
            nao_encontradas.append(nf)
            continue
        num = int(num_str)
        nf_data  = nf.get("data")
        nf_valor = Decimal(nf["valor"].replace(",", "."))

        # 1) todas as ocorrências no relatório com o mesmo número
        matches = [r for r in relatorio
                   if r.get("numero","").isdigit() and int(r["numero"]) == num]

        # 2) se não encontrar nenhuma, vai para 'não encontradas'
        if not matches:
            nao_encontradas.append(nf)
            continue

        # 3) se houver múltiplas, tente achar uma com os 3 fatores iguais
        if len(matches) > 1:
            # procura por correspondência exata
            exact = next(
                (r for r in matches
                 if r.get("data") == nf_data and r.get("valor") == nf_valor),
                None
            )
            if exact:
                encontradas.append(nf)
            else:
                nao_encontradas.append(nf)
            continue

        # 4) se há exatamente 1 match, compara normalmente
        r = matches[0]
        if r.get("data") == nf_data and r.get("valor") == nf_valor:
            encontradas.append(nf)
        else:
            # aqui há divergência de data ou valor
            nf["esperado"] = r
            divergentes.append(nf)

    # gera o PDF de saída
    os.makedirs(output_dir, exist_ok=True)
    pdf = os.path.join(output_dir, "relatorio_validacao.pdf")
    gerar_pdf_relatorio(encontradas, divergentes, nao_encontradas, pdf)

    return {
        "encontradas": encontradas,
        "divergentes": divergentes,
        "nao_encontradas": nao_encontradas,
        "pdf": pdf
    }
    
def _rm_error_handler(func, path, exc_info):
    # torna o arquivo gravável e tenta remover de novo
    os.chmod(path, stat.S_IWRITE)
    func(path)

def processar_comparacao_nf(zip_path, relatorio_pdf_path, output_dir):
    temp  = os.path.join(os.path.dirname(output_dir), "temp_notas")

    # limpa pastas antigas
    if os.path.isdir(temp):
        shutil.rmtree(temp, onerror=_rm_error_handler)
    if os.path.isdir(output_dir):
        shutil.rmtree(output_dir, onerror=_rm_error_handler)

    os.makedirs(temp, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    notas = extrair_notas_zip(zip_path, temp)
    rel   = extrair_relatorio_com_pdfplumber(relatorio_pdf_path)

    # ao invés de “return comparar_nfs(notas, rel, output_dir)”, faça:
    resultado = comparar_nfs(notas, rel, output_dir)
    # retira o caminho do PDF do dict e devolve como segundo valor
    pdf_path = resultado.pop("pdf")
    return resultado, pdf_path
