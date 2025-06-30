import os
import rarfile
import datetime
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
    try:
        # abre o PDF
        doc = fitz.open(pdf_path)
        if doc.is_encrypted:
            print(f"[DEBUG] Ignorando encriptado: {pdf_path}")
            return None

        # lê todo o texto (com OCR se necessário)
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

        # filtra apenas NFSe / prestação de serviço
        if not re.search(r"\bNFS[- ]?E\b", texto, re.IGNORECASE) \
        and not re.search(r"\bPrestação de Serviços\b", texto, re.IGNORECASE) \
        and not re.search(r"\bPrestador de Serviços\b", texto, re.IGNORECASE):
            print(f"[DEBUG] Ignorando sem padrão de NFSe ou Serviço: {pdf_path}")
            return None

        # número pelo nome do arquivo
        nome = os.path.basename(pdf_path)
        m_num = re.search(r"(\d+)", nome)
        numero = str(int(m_num.group(1))).lstrip("0") if m_num else None

        # monta trecho de cabeçalho (antes de "DISCRIMINAÇÃO DOS SERVIÇOS")
        header_text = texto.split("DISCRIMINAÇÃO")[0]

        # padrões de data de emissão
        date_patterns = [
            r"Data de Geração da NFS[- ]?e[^\d]*(\d{2}/\d{2}/\d{4})",    # DF-DF usa esse label :contentReference[oaicite:0]{index=0}
            r"Data de Competência[^\d]*(\d{2}/\d{2}/\d{4})",
            r"Emitido em[:\s]*[\r\n\s]*(\d{2}/\d{2}/\d{4})",
            r"Data e Hora de Emissão[^\d]*([\d]{2}/[\d]{2}/[\d]{4})",
            r"Emissão\s*[:]\s*[\r\n\s]*(\d{2}/\d{2}/\d{4})",
            r"Emissão em[:\s]*[\r\n\s]*(\d{2}/\d{2}/\d{4})"
        ]

        data = None
        for pat in date_patterns:
            m = re.search(pat, header_text, re.IGNORECASE)
            if m:
                data = m.group(1)
                break

        if not data:
            non_vencimento_date_found = None
            
            # Encontra todas as ocorrências de datas no formato DD/MM/AAAA com suas posições
            for match in re.finditer(r"(\d{2}/\d{2}/\d{4})", header_text):
                current_date_str = match.group(1)
                start_pos = match.start()
                
                # Define uma janela antes da data para verificar a presença de "Vencimentos"
                # Procura até 50 caracteres para trás
                check_start = max(0, start_pos - 50) 
                pre_date_context = header_text[check_start:start_pos]
                
                # Se "Vencimentos" não for encontrado no contexto anterior, considera essa data
                if not re.search(r"Vencimento(s)?", pre_date_context, re.IGNORECASE):
                    non_vencimento_date_found = current_date_str
                    break # Pega a primeira data adequada e sai do loop
            
            if non_vencimento_date_found:
                data = non_vencimento_date_found

        # extrai valor (maior valor encontrado em todo o texto)
        raw_vals = re.findall(r"\d{1,3}(?:\.\d{3})*,\d{2}", texto)
        valor = None
        if raw_vals:
            vf = max(Decimal(v.replace(".", "").replace(",", ".")) for v in raw_vals)
            valor = str(vf.quantize(Decimal("0.01")))

        return {"numero": numero, "data": data, "valor": valor}

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

        # Converte o valor da NF-e individual para Decimal para comparação consistente.
        # O valor é uma string em 'notas', então o 'replace' é necessário aqui.
        nf_valor_decimal = Decimal(nf.get("valor", "0").replace(",", "."))

        # 1. Tentar encontrar uma correspondência EXATA (número, data e valor)
        exact_match = None
        for r in relatorio:
            if r.get("numero") and r["numero"].isdigit():
                # O valor de 'r["valor"]' já é um Decimal vindo de 'extrair_relatorio_com_pdfplumber',
                # então NÃO precisamos do '.replace()' aqui.
                r_valor_decimal = r.get("valor", Decimal("0"))

                if int(r["numero"]) == num and \
                   r.get("data") == nf.get("data") and \
                   r_valor_decimal == nf_valor_decimal:
                    exact_match = r
                    break

        if exact_match:
            encontradas.append(nf)
            continue

        # 2. Se não houver correspondência exata, tentar encontrar por NÚMERO e DATA
        date_match = None
        for r in relatorio:
            if r.get("numero") and r["numero"].isdigit():
                if int(r["numero"]) == num and r.get("data") == nf.get("data"):
                    date_match = r
                    break

        if date_match:
            # Encontrou por número e data, mas o valor pode ser diferente. É uma divergência.
            # Compara nf_valor_decimal (já Decimal) com date_match['valor'] (já Decimal)
            if nf_valor_decimal == date_match["valor"]:
                encontradas.append(nf)
            else:
                nf["esperado"] = date_match
                divergentes.append(nf)
            continue

        # 3. Se não houver correspondência por número e data, tentar apenas por NÚMERO
        num_only_match = None
        for r in relatorio:
            if r.get("numero") and r["numero"].isdigit() and int(r["numero"]) == num:
                num_only_match = r
                break

        if num_only_match:
            # Encontrou apenas por número. É uma divergência.
            nf["esperado"] = num_only_match
            divergentes.append(nf)
            continue

        # Se não encontrou nenhuma correspondência
        nao_encontradas.append(nf)

    # O restante da função 'comparar_nfs' permanece inalterado
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
    temp  = os.path.join(os.path.dirname(output_dir), "temp_notas")
    notas = extrair_notas_zip(zip_path, temp)

    # usa pdfplumber para extrair o relatório
    rel = extrair_relatorio_com_pdfplumber(relatorio_pdf_path)

    res = comparar_nfs(notas, rel, output_dir)
    pdf = res.pop("pdf")
    return res, pdf
