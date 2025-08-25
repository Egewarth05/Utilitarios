import pandas as pd
import os
import difflib
import sys
import re

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# Nome do arquivo Excel principal que contém todas as abas.
NOME_ARQUIVO_EXCEL_PRINCIPAL = None

# Nomes exatos das abas (sheets) dentro do arquivo Excel que você quer processar.
NOME_ABA_ESQUADRIAS = 'Pagamentos com NF Esquadrias'
POSSIVEIS_FERRO_NOBRE = ['Pagamentos com NF Ferro Nobre', 'Paamentos com NF Ferro Nobre']

HEADER_ROW_ESQUADRIAS = 3
HEADER_ROW_FERRO_NOBRE = 1

COLUNAS_ORIGINAIS = {
    'Nome': 'Fornecedor',
    'Documento': 'Documento',
    'Dt.Baixa': 'Data',
    'Vlr.Recebido': 'Valor',
    'Centro Custo': 'Centro de Custo',
}

ARQUIVO_SAIDA = None

CODIGOS_FORNECEDORES = {
    'ACEVILLE TRANSPORTES LTDA': 5072093,
    'ACO METAIS': 5022258,
    'ASX AGRODRONE': 5121940,
    'ATACADO VELHO OESTE LTDA': 5040323,
    'AUTO POSTO CANARINHO LTDA': 5006241,
    'BOLD': 5021250,
    'BONDMANN QUIMICA': 5120465,
    'BRASTELHA INDUSTRIAL': 5023466,
    'CASA DAS CHAVES E FECHADURAS': 5196809,
    'CETRIC': 5020836,
    'COMFORT DOOR': 5196636,
    'CONSTRUTECH SOLUCOES LTDA': 5038628,
    'COOPERATIVA A1': 5000007,
    'Copema Distribuidora Automotiva': 5002567,
    'CPEL INDUSTRIA DE PAPEL LTDA.': 5196648,
    'DISK VIDROS': 5020397,
    'ENERVOLT SOLUCOES EM ENERGIA LTDA': 5015186,
    'EXPRESSO SAO MIGUEL S/A' : 5000475,
    'CERTISIGN CERTIFICADORA DIGITAL S.A' : 5196739,
    'G.V. COMERCIO DE MATERIAIS DE FERRAGENS LTDA': 5015199,
    'G2 COMPONENTES' : 5050002,
    'GMAD CHAPECOMP SUPRIMENTOS PARA MOVEIS LTDA': 5020804,
    'GRD E-COMMERCE LTDA' : 5122699,
    'GURGELMIX MAQUINAS E FERRAMENTAS S.A.' : 5026594,
    'HDU INDUSTRIA METALURGICA LTDA EPP' : 5196659,
    'INATEC CONTABILIDADE' : 5000107,
    'ARTEFERRO BRASIL' : 5122443,
    'JAIR MOREIRA' : 5015851,
    'JGI COMERCIO DE OXIGENIOS PATO BRANCO' : 5121354,
    'JOSE ANTONIO SANTIN  FILHOS LTDA' : 5014909,
    'KARLINSKI COMERCIO DE TINTAS LTDA - ME': 5017934,
    'LEILIANE PIETRO BIAZI ME' : 5042424,
    'LUTHIER GESTAO ESTRATEGICA EM PROPRIEDADE INTELECTUAL' : 5199807,
    'MULLER DISTRIBUIDORA' : 5021276,
    'MAQFER TELHAS DE ALUZINCO' : 5043163,
    'Megamaq Ferramentas e Acessorios para Industria Lt' : 5122798,
    'MEGAMAQ FERRAMENTAS E ACESSORIOS PARA INDUSTRIA' : 5122798,
    'METALURGICA CORTESA LTDA' : 5043494,
    'METALURGICA DE TONI' : 5044274,
    'MINNER SC' : 5002494,
    'MULTI-ACO INDUSTRIA E COMERCIO LTDA' : 5023464,
    'MUNDIAL EXPRESS' : 5199808,
    'NAPORT PORTAS E PORTOES' : 5196788,
    'NESSI PROMOCAO DE EVENTOS E COMERCIAL LTDA' : 5199900,
    'PERFIL INDUSTRIA DE PERFILADOS' : 5030796,
    'PERFILADOS VANZIN LTDA' : 5056904,
    'PERFYACO MATRIZ' : 5020848,
    'POSTO FREY' : 5122057,
    'PRIMA FERRAGENS COMERCIAL LTDA' : 5196808,
    'PRIME LOCK BLINDAGENS E FECHADURAS DE SEGURANCA LTDA' : 5122553,
    'RIBEIRO NEGOCIOS DIGITAIS E CONSULTORIA LTDA' : 5122551,
    'RODONAVES TRANSPORTES E ENCOMENDAS LTDA' : 5000946,
    'ROYER ENCOMENDAS LTDA' : 5000718,
    'S.O.S SEGURANCA E SAUDE NO TRABALHO' : 5030084,
    'SIDERACO DISTRIBUIDORA DE ACOS LTDA' : 5200103,
    'SMO CRONOTACOGRAFOS LTDA' : 5121079,
    'SOLDAS PLANALTO' : 5121604,
    'T. F. PEREIRA FERRAGENS' : 5196791,
    'TECNO COMERCIO, INDUSTRIA, IMPORTACAO E EXPORTACAO' : 5200105,
    'TEVERE' : 5000738,
    'TR COMERCIO DE INOX E VARIEDADES' : 5200104,
    'UNIAO METAIS' : 5200102,
    'VITRAL-SUL' : 5020407,
    'WEG TINTAS LTDA' : 5122666,
    'WERKEUG ULTRA LTDA' : 5196790,
    'WURTH BRASIL FIXAÇAO' : 5001429,
    'MAVAMPAR' : 5015199,
    'CAMARA DE DIRIGENTES LOJISTAS DE SAO MIGUEL DO OESTE/SC' : 4553,
    'W VETRO' : 5122552,
    'DAL CORTIVO ADVOCACIA EMPRESARIAL' : 1496,
    'SOMAQ ASSISTENCIA E EQUIPAMENTOS' : 5001697,
    'CERUMAR SERVICOS EM PROPRIEDADE INTELECTUAL LTDA' : 5030518,
    'INVIOLAVEL SAO MIGUEL' : 5001525,
    'CELESC DISTRIBUICAO S.A' : 5001704,
    'ENERTEC ENERGIA E TECNOLOGIA' : 5001704,
    'DLOCAL' : 1496,
    'BOLD CHAPECO' : 5021250,
    'MINNER COMERCIAL LTDA' : 5002494,
    'MICHELIN VIDROS' : 5023090,
    'ALVANE COMERCIO DE ALUMINIOS E COMPONENTES LTDA.' : 5196624,
    'REAL VIDROS' : 5020413,
    'SOMAPRINT IMPRESSAO DIGITAL' : 4376,
    'ACISMO' : 4553,
    'SANTO ANTONIO INDUSTRIA DE VIDROS LTDA' : 5196649,
    'CORSO COMERCIO DE ALUMINIOS E ACESSORIOS' : 5122554,
    'HMB GESTAO DE ATIVOS' : 5196643,
    'NOPLA METAIS' : 5196673,
    'DALLSUL COMERCIAL LTDA' : 1496,
    'FERMAX' : 5122555,
    'SANTO ANTONIO INDUSTRIA DE VIDROS LTDA' : 5196649,
    'VIDRACARIA LINDE' : 5196623,
    'LF COMERCIAL DE BENS LTDA' : 1496,
    'ALUMASA INDUSTRIA DE PLASTICO E ALUMINIO LTDA' : 5026491,
    'DXMAX BRASIL LTDA - PE' : 1496,
    'PADO S/A' : 5051482,
    'M5 SOLUTION GLASS' : 3525,
    'RADIO PEPERI' : 5008111,
    'EMTECO SOLUCOES' : 5196695,
    'IPASA' : 1496,
    'AEST' : 4553,
    'ICILEGEL INDUSTRIA E COMERCIO IBAITI LTDA' : 5196633,
    'LOGICA INFORMATICA' : 5122797,
    'ALUFORTE' : 5020405,
    'CHIRU FM' : 5002539,
    'OPTIDATA' : 5122052,
    'TEMPERMED' : 5196621,
    'CREA SC' : 4532,
    'MALLMANN PROJETOS E EMPREENDIMENTOS' : 1496,
    'SCHERER SA COMERCIO DE AUTOPECAS' : 5020506,
    'VIDROMAX FERRAGENS' : 5196838,
    'MEPAR' : 5002389,
    'DESPACHANTE TUNAPOLIS' : 4537,
    "WOOD'S POWR-GRIP DO BRASIL" : 5196837,
    'INGA VEICULOS LTDA' : 5020757,
    'ELITINOX' : 5196647,
    'PERFIL ALUMINIO DO BRASIL S/A' : 5200077,
    'MC MATRIZES CARDEAL' : 5028761,
    'ALUESTE ALUMINIOS' : 5020411,
    'BAUER CARGAS' : 5000728,
    'CARPBRASIL FERRAGENS ESPECIAIS' : 5200076,
    'REUNIDAS' : 5002377,
    'AVANTPAR' : 5022435,
    'CONSTANTI METAIS LTDA' : 5177353,
    'DONNA FERRAGENS' : 5200079,
    'PINTA RISCO' : 5199704,
    'JORNAL REGIONAL DO OESTE' : 4376,
    'ANCORA GROUP' : 5199805,
    'ALUMICONTE COMPONENTES DE ALUMINIO LTDA' : 5200080,
    'SOMA PUXADORES' : 5200083,
    'C & A EMBALAGENS' : 5196644,
    'SUKIRA IMPORTACAO E EXPORTACAO' : 5200082,
    'ALINE SCHROEDER' : 25043,
    'EXPRESSO SAO MIGUEL S/A CURITIBA' : 5000475,
    'IEM CARDEALMAR' : 5196617,
    'KRG INDUSTRIA' : 5200081,
    'COOPER A1 TUNAS' : 5000007,
    'CONSTRUTECH' : 5038628,
    'FREIBERGER SERVIÇOS ELETROMECANICOS' : 5001125,
    'INVIOLAVEL CEDRO SISTEMAS DE ALARMES LTDA ME' : 5001525,
    'ESQUADRISOFT SISTEMAS LTDA' : 5200108,
    'PORTAL INFORMATICA' : 5007621,
    'RADIO ITAPIRANGA LTDA' : 25056,
    'POSTO BELA VISTA' : 5030809,
    'SUPERMERCADO VENEZA LTDA' : 5010201,
    'BOLFE EMPREENDIMENTOS E PARTICIPACOES LTDA' : 25060,
}

# --- Overrides específicos para Schroeder Esquadrias ---
ESQUADRIAS_OVERRIDES = {
    'NAPORT PORTAS E PORTOES': 25063,
    'ROBERT BOSCH' : 5021397,
    'MULTI-ACO INDUSTRIA E COMERCIO LTDA': 25064,
    'ARTEFERRO BRASIL': 25065,
    'PRIMA FERRAGENS' : 5200158,
    'Copema Distribuidora Automotiva': 25066,
    'PERFILADOS VANZIN LTDA': 25067,
    'JGI COMERCIO DE OXIGENIOS PATO BRANCO': 25068,
    'TEVERE': 25069,
    'SUPER CATARINA COMERCIO DE VIDROS': 5196693,
    'INVIOLAVEL SAO MIGUEL' : 1496,
    'W VETRO' : 1496,
    'PORTAL INFORMATICA' : 1496,
    'CELESC DISTRIBUICAO S.A' : 1496,
    'PERFYACO MATRIZ': 25070,
    'BONDMANN QUIMICA': 25071,
    'NESSI PROMOCAO DE EVENTOS E COMERCIAL LTDA': 25072,
}

def _rebuild_normalized_map():
    global CODIGOS_FORNECEDORES_NORMALIZADO
    CODIGOS_FORNECEDORES_NORMALIZADO = {
        normalize_name(k): v for k, v in CODIGOS_FORNECEDORES.items()
    }

def set_empresa(empresa: str | None):
    if not empresa:
        return
    empresa = empresa.strip().lower()
    if empresa in ('esquadrias', 'schroeder esquadrias'):
        CODIGOS_FORNECEDORES.update(ESQUADRIAS_OVERRIDES)
        _rebuild_normalized_map()  # <-- RECONSTRÓI o mapa normalizado após atualizar

def normalize_name(s):
    return re.sub(r'\s+', ' ', str(s).strip().upper())

# versão normalizada para lookup mais robusto
CODIGOS_FORNECEDORES_NORMALIZADO = {
    normalize_name(k): v for k, v in CODIGOS_FORNECEDORES.items()
}

def fuzzy_match_column(target, candidates, cutoff=0.6):
    target_norm = re.sub(r'[\W_]+', '', target).lower()
    best = None
    best_score = 0
    for c in candidates:
        c_norm = re.sub(r'[\W_]+', '', c).lower()
        score = difflib.SequenceMatcher(None, target_norm, c_norm).ratio()
        if score > best_score:
            best_score = score
            best = c
    if best_score >= cutoff:
        return best
    return None

def detectar_header(df_path, sheet_name, dtypes, esperados, max_scan=8):
    for header_row in range(0, max_scan):
        try:
            df_try = pd.read_excel(df_path, sheet_name=sheet_name, header=header_row, dtype=dtypes)
        except Exception:
            continue
        cols = [str(c).strip() for c in df_try.columns]
        matches = sum(1 for e in esperados if fuzzy_match_column(e, cols))
        if matches >= 2:  # heurística mínima
            print(f"Usando header_row={header_row} para '{sheet_name}' com {matches} matches.")
            return df_try
    print(f"Não encontrou header consistente em '{sheet_name}', usando row 0 como fallback.")
    return pd.read_excel(df_path, sheet_name=sheet_name, header=0, dtype=dtypes)

def processar_dados_sheet(df, original_cols_map, sheet_name):
    print(f"Processando dados da aba: '{sheet_name}'...")

    # limpa nomes brutos
    df.columns = df.columns.astype(str).str.strip()

    # se faltam colunas exatas, tenta fazer fuzzy match
    mapped_columns = {}
    for orig in original_cols_map.keys():
        if orig in df.columns:
            mapped_columns[orig] = orig
        else:
            # tenta match aproximado entre os cabeçalhos disponíveis
            match = fuzzy_match_column(orig, df.columns.tolist())
            if match:
                print(f"⚠️ Cabeçalho '{orig}' não encontrado exatamente em '{sheet_name}', usando aproximação '{match}'.")
                mapped_columns[orig] = match
            else:
                print(f"ERRO: Na aba '{sheet_name}', a coluna original '{orig}' não foi encontrada nem aproximada.")
                # não retorna de imediato: permite montar com colunas faltantes
                mapped_columns[orig] = None

    # seleciona e renomeia: cria df_final com colunas disponíveis
    df_final = pd.DataFrame()
    for orig, new_name in original_cols_map.items():
        source_col = mapped_columns.get(orig)
        if source_col and source_col in df.columns:
            df_final[new_name] = df[source_col]
        else:
            # coluna faltante: preenche com vazio (ou 0 se for valor)
            if new_name in ['Valor']:
                df_final[new_name] = 0
            else:
                df_final[new_name] = ''
            print(f"[WARN] Coluna '{new_name}' não encontrada: preenchida com padrão.")

    # --- Tratamento e Formatação dos Dados ---

    # 1. Coluna 'Data' (Dt.Baixa -> Data)
    df_final['Data'] = pd.to_datetime(df_final['Data'], errors='coerce', dayfirst=True)
    df_final['Data'] = df_final['Data'].dt.strftime('%d%m%Y').fillna('')

    # Tratamento robusto de valores com vírgula como separador decimal (formato BR)
    df_valor = df_final['Valor'].astype(str).str.strip()
    df_valor = df_valor.apply(lambda x: x.replace('.', '').replace(',', '.') if ',' in x else x)
    df_final['Valor'] = pd.to_numeric(df_valor, errors='coerce').fillna(0).apply(lambda x: f"{x:.2f}")

    # 3. Coluna 'Documento' (CNPJ/CPF)
    df_final['Documento'] = df_final['Documento'].astype(str).str.strip()
    df_final['Documento'] = df_final['Documento'].str.replace(r'[./-]', '', regex=True)

    # garante string e normaliza fornecedor para lookup
    df_final['Fornecedor'] = df_final['Fornecedor'].fillna('').astype(str)
    df_final['Fornecedor_NORM'] = df_final['Fornecedor'].apply(normalize_name)
    df_final['Código'] = df_final['Fornecedor_NORM'].map(CODIGOS_FORNECEDORES_NORMALIZADO).fillna('NÃO DEFINIDO')
    # opcional: remover coluna auxiliar se não quiser mantê-la
    df_final.drop(columns=['Fornecedor_NORM'], inplace=True)

    # exclusão robusta de fornecedores indesejados (trata NaN e normaliza)
    mask_excluir = df_final['Fornecedor'].fillna('').astype(str).str.upper().str.startswith((
        'SICOOB', 'MUNICIPIO', 'MINISTERIO', 'VIVO', 'NEDEL', 'SECRETARIA', 'TUNAPOLIS', 'INMETRO'
    ))
    df_final = df_final[~mask_excluir]

    return df_final

def gerar_planilhas_convertidas(processed_dfs, arquivo_saida):
    print("Gerando planilhas convertidas...")

    with pd.ExcelWriter(arquivo_saida, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        for nome_original, nome_convertido in [
            ('Esquadrias Processada', 'Converter Esquadrias'),
            ('Ferro Nobre Processada', 'Converter Ferros Nobres')
        ]:
            if nome_original not in processed_dfs:
                print(f"AVISO: Aba '{nome_original}' não encontrada para conversão.")
                continue

            df = processed_dfs[nome_original].copy()

            # Remove linhas sem código
            if 'Código' in df.columns:
                df = df[df['Código'].astype(str).str.strip().str.upper() != 'NÃO DEFINIDO']

            if 'Código' not in df.columns:
                print(f"⚠️ Coluna 'Código' não encontrada na aba '{nome_original}'. Usando vazio.")
                df['CodigoLimpo'] = ""
            else:
                def _coerce_codigo(x):
                    if pd.isna(x):
                        return ""
                    if isinstance(x, (int, float)):
                        return str(int(x))          # sem .0 e sem zeros a mais
                    s = str(x).strip()
                    if s.lower() == 'não definido':
                        return ""
                    m = re.search(r'\d+', s)       # pega só os dígitos (evita “.0”)
                    return m.group(0) if m else ""

                # ✅ Atribuição que faltava:
                df['CodigoLimpo'] = df['Código'].apply(_coerce_codigo)

            # (opcional) se quiser descartar linhas ainda sem código após o coerce:
            # df = df[df['CodigoLimpo'] != ""]

            df['Linha Convertida'] = df.apply(
                lambda row: ",".join([
                    "1",
                    row["Data"],
                    row["CodigoLimpo"],
                    "5",
                    row["Valor"].replace(",", "."),
                    "337",
                    f'"{row["Fornecedor"]} - {row["Documento"]} - {row["Centro de Custo"]}"'
                ]),
                axis=1
            )

            df[['Linha Convertida']].to_excel(writer, sheet_name=nome_convertido, index=False, header=False)
            print(f"Aba '{nome_convertido}' criada com sucesso.")

def carregar_sheet_com_header_detectado(caminho, sheet_name, dtypes, esperados, max_scan=8):
    """
    Tenta ler a sheet com header em várias linhas, escolhendo aquela
    que corresponde melhor aos nomes esperados.
    """
    for header_row in range(max_scan):
        try:
            df_try = pd.read_excel(
                caminho,
                sheet_name=sheet_name,
                header=header_row,
                dtype=dtypes,
            )
        except Exception:
            continue
        cols = [str(c).strip() for c in df_try.columns]
        # conta quantos esperados “batem” (por fuzzy match ou exato)
        matches = sum(
            1
            for exp in esperados
            if exp in cols or fuzzy_match_column(exp, cols)
        )
        if matches >= 2:  # achou pelo menos 2 colunas-chave
            print(f"→ Usando header_row={header_row} em '{sheet_name}' ({matches} matches)")
            return df_try
    # fallback
    print(f"⚠️ Não detectou header automático em '{sheet_name}', usando header_row=0")
    return pd.read_excel(caminho, sheet_name=sheet_name, header=0, dtype=dtypes)

def processar_planilha_pagamentos_separado():
    if not NOME_ARQUIVO_EXCEL_PRINCIPAL or not ARQUIVO_SAIDA:
        raise ValueError("Arquivo de entrada/saída não definido. Use a função *_custom(...) passando os caminhos.")

    print("DEBUG lendo:", NOME_ARQUIVO_EXCEL_PRINCIPAL)
    print("DEBUG escrevendo:", ARQUIVO_SAIDA)

    processed_dfs = {}

    POSSIVEIS_ESQUADRIAS  = ['Pagamentos com NF Esquadrias', 'Paagamentos com NF Esquadrias']
    POSSIVEIS_FERRO_NOBRE = ['Pagamentos com NF Ferro Nobre', 'Paamentos com NF Ferro Nobre']

    xls = pd.ExcelFile(NOME_ARQUIVO_EXCEL_PRINCIPAL)
    available_sheets = [s.strip() for s in xls.sheet_names]

    def pick_existing_sheet(possiveis):
        for candidate in possiveis:
            for existing in available_sheets:
                if existing.strip().lower() == candidate.strip().lower():
                    return existing
        return None

    esperados = list(COLUNAS_ORIGINAIS.keys())
    dtypes = {'Vlr.Recebido': str, 'Documento': str}

    # 1) Esquadrias (se houver)
    sheet_esquad = pick_existing_sheet(POSSIVEIS_ESQUADRIAS)
    if sheet_esquad:
        print(f"Iniciando processamento da aba Esquadrias: '{sheet_esquad}'")
        df_raw = carregar_sheet_com_header_detectado(
            NOME_ARQUIVO_EXCEL_PRINCIPAL,
            sheet_esquad,
            dtypes,
            esperados,
            max_scan=10
        )
        processed_dfs['Esquadrias Processada'] = processar_dados_sheet(
            df_raw, COLUNAS_ORIGINAIS, 'Esquadrias Processada'
        )

    # 2) Ferro Nobre (independente de ter Esquadrias)
    sheet_ferro = pick_existing_sheet(POSSIVEIS_FERRO_NOBRE)
    if sheet_ferro:
        print(f"Iniciando processamento da aba Ferro Nobre: '{sheet_ferro}'")
        df_raw = carregar_sheet_com_header_detectado(
            NOME_ARQUIVO_EXCEL_PRINCIPAL,
            sheet_ferro,
            dtypes,
            esperados,
            max_scan=10
        )
        processed_dfs['Ferro Nobre Processada'] = processar_dados_sheet(
            df_raw, COLUNAS_ORIGINAIS, 'Ferro Nobre Processada'
        )

    if not processed_dfs:
        print("Nenhuma aba de Esquadrias nem de Ferro Nobre encontrada. Abortando.")
        return

    # 3) Grava saída e planilhas convertidas
    with pd.ExcelWriter(ARQUIVO_SAIDA, engine='openpyxl') as writer:
        for nome_aba, df in processed_dfs.items():
            df.to_excel(writer, sheet_name=nome_aba, index=False)

    gerar_planilhas_convertidas(processed_dfs, ARQUIVO_SAIDA)

def processar_planilha_pagamentos_separado_custom(arquivo_entrada, arquivo_saida, empresa=None):
    global NOME_ARQUIVO_EXCEL_PRINCIPAL, ARQUIVO_SAIDA
    NOME_ARQUIVO_EXCEL_PRINCIPAL = arquivo_entrada
    ARQUIVO_SAIDA = arquivo_saida
    set_empresa(empresa)  # <-- aplica regras por empresa
    processar_planilha_pagamentos_separado()

def main():
    if len(sys.argv) >= 3:
        entrada = sys.argv[1]
        saida = sys.argv[2]
        empresa = sys.argv[3] if len(sys.argv) >= 4 else None  # <-- NOVO
        processar_planilha_pagamentos_separado_custom(entrada, saida, empresa)
    else:
        processar_planilha_pagamentos_separado()

if __name__ == "__main__":
    main()
