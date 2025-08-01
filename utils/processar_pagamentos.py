import pandas as pd
import os

# --- Configurações ---
# Nome do arquivo Excel principal que contém todas as abas.
NOME_ARQUIVO_EXCEL_PRINCIPAL = 'Relat cont_bil.xlsx'

# Nomes exatos das abas (sheets) dentro do arquivo Excel que você quer processar.
NOME_ABA_ESQUADRIAS = 'Pagamentos com NF Esquadrias'
NOME_ABA_FERRO_NOBRE = 'Paamentos com NF Ferro Nobre' # Corrigido: "Paamentos"

HEADER_ROW_ESQUADRIAS = 3
HEADER_ROW_FERRO_NOBRE = 1 


COLUNAS_ORIGINAIS = {
    'Nome': 'Fornecedor',
    'Documento': 'Documento',
    'Dt.Baixa': 'Data',
    'Vlr.Recebido': 'Valor',
    'Centro Custo': 'Centro de Custo',
}

ARQUIVO_SAIDA = 'pagamentos_processados_final.xlsx' 

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
    'IBAMA/CGFIN - COORDENACAO GERAL DE FINANCAS' : 5001525,
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
    
}

def processar_dados_sheet(df, original_cols_map, sheet_name):
    """
    Processa um DataFrame de uma única aba:
    - Limpa espaços em branco dos nomes das colunas.
    - Verifica a existência das colunas necessárias.
    - Seleciona, renomeia e formata as colunas.
    """
    print(f"Processando dados da aba: '{sheet_name}'...")

    # Limpa espaços em branco dos nomes das colunas imediatamente após a leitura
    df.columns = df.columns.str.strip()

    # Verifica se as colunas necessárias estão no DataFrame
    for col_original in original_cols_map.keys():
        if col_original not in df.columns:
            print(f"ERRO: Na aba '{sheet_name}', a coluna original '{col_original}' não foi encontrada.")
            print(f"Colunas encontradas após limpeza: {df.columns.tolist()}")
            return None

    # Seleciona as colunas desejadas e as renomeia
    colunas_para_selecionar = list(original_cols_map.keys())
    df_final = df[colunas_para_selecionar].copy()
    df_final.rename(columns=original_cols_map, inplace=True)

    # --- Tratamento e Formatação dos Dados ---

    # 1. Coluna 'Data' (Dt.Baixa -> Data)
    df_final['Data'] = pd.to_datetime(df_final['Data'], errors='coerce', dayfirst=True)
    df_final['Data'] = df_final['Data'].dt.strftime('%d%m%Y').fillna('')

    # Tratamento robusto de valores com vírgula como separador decimal (formato BR)
    df_valor = df_final['Valor'].astype(str).str.strip()

    # Se contém vírgula, então é BR (ex: 2.020,83)
    df_valor = df_valor.apply(lambda x: x.replace('.', '').replace(',', '.') if ',' in x else x)

    df_final['Valor'] = pd.to_numeric(df_valor, errors='coerce').fillna(0).apply(lambda x: f"{x:.2f}")

    # 3. Coluna 'Documento' (CNPJ/CPF)
    df_final['Documento'] = df_final['Documento'].astype(str).str.strip()
    df_final['Documento'] = df_final['Documento'].str.replace(r'[./-]', '', regex=True)
    
    # 4. Coluna 'Código' com base no dicionário
    df_final['Código'] = df_final['Fornecedor'].map(CODIGOS_FORNECEDORES).fillna('NÃO DEFINIDO')
    # 5. Remover fornecedores que começam com SICOOB, MUNICÍPIO ou MINISTÉRIO
    df_final = df_final[~df_final['Fornecedor'].str.upper().str.startswith(('SICOOB', 'MUNICIPIO', 'MINISTERIO', 'VIVO', 'NEDEL', 'SECRETARIA', 'TUNAPOLIS', 'INMETRO'))]

    return df_final

def gerar_planilhas_convertidas(processed_dfs, arquivo_saida):
    """
    Gera duas novas abas convertidas no formato solicitado:
    1,data,codigo,5,valor,337,"fornecedor - documento - centro de custo"
    """
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

            if 'Código' not in df.columns:
                print(f"⚠️ Coluna 'Código' não encontrada na aba '{nome_original}'. Usando '0000000' como padrão.")
                df['CodigoLimpo'] = "0000000"
            else:
                df['CodigoLimpo'] = df['Código'].apply(lambda x: str(int(float(x))).zfill(7) if pd.notna(x) and str(x).strip().lower() != 'não definido' else "0000000")

            # Constrói a linha final convertida
            df['Linha Convertida'] = df.apply(lambda row: ",".join([
                "1",
                row["Data"],
                row["CodigoLimpo"],
                "5",
                row["Valor"].replace(",", "."),
                "337",
                f'"{row["Fornecedor"]} - {row["Documento"]} - {row["Centro de Custo"]}"'
            ]), axis=1)

            # Salva apenas a coluna final convertida
            df[['Linha Convertida']].to_excel(writer, sheet_name=nome_convertido, index=False, header=False)
            print(f"Aba '{nome_convertido}' criada com sucesso.")

def processar_planilha_pagamentos_separado():
    """
    Função principal:
    Lê o arquivo Excel, processa cada aba de interesse separadamente,
    e salva os resultados em um novo arquivo Excel com abas distintas.
    Suporta o caso em que apenas uma das abas existe (processa só a que estiver presente).
    """
    if not os.path.exists(NOME_ARQUIVO_EXCEL_PRINCIPAL):
        print(f"ERRO: O arquivo Excel principal '{NOME_ARQUIVO_EXCEL_PRINCIPAL}' não foi encontrado na pasta do script.")
        print("Por favor, verifique se o nome está correto e se o arquivo está na mesma pasta.")
        return

    processed_dfs = {}  # Dicionário para armazenar os DataFrames processados por nome de aba

    # Tenta listar as abas disponíveis para evitar exceções desnecessárias
    try:
        xls = pd.ExcelFile(NOME_ARQUIVO_EXCEL_PRINCIPAL)
        available_sheets = [s.strip() for s in xls.sheet_names]
    except Exception as e:
        print(f"ERRO: Não foi possível abrir o arquivo '{NOME_ARQUIVO_EXCEL_PRINCIPAL}' para listar abas. Erro: {e}")
        return

    def sheet_exists(name):
        return any(s.lower() == name.lower() for s in available_sheets)

    # Isso é crucial para que leitura com strings funcione conforme esperado
    dtypes_leitura = {'Vlr.Recebido': str, 'Documento': str}

    # --- Processa a aba da Schroeder Esquadrias ---
    if sheet_exists(NOME_ABA_ESQUADRIAS):
        try:
            print(f"Iniciando leitura da aba '{NOME_ABA_ESQUADRIAS}'...")
            df_esquadrias_raw = pd.read_excel(
                NOME_ARQUIVO_EXCEL_PRINCIPAL,
                sheet_name=NOME_ABA_ESQUADRIAS,
                header=HEADER_ROW_ESQUADRIAS,
                dtype=dtypes_leitura
            )
            df_processed_esquadrias = processar_dados_sheet(
                df_esquadrias_raw.copy(), COLUNAS_ORIGINAIS, "Esquadrias Processada"
            )
            if df_processed_esquadrias is not None:
                processed_dfs['Esquadrias Processada'] = df_processed_esquadrias
        except Exception as e:
            print(f"AVISO: Não foi possível ler ou processar a aba '{NOME_ABA_ESQUADRIAS}'. Erro: {e}")
    else:
        print(f"AVISO: Aba esperada '{NOME_ABA_ESQUADRIAS}' não encontrada. Pulando.")

    # --- Processa a aba da Schroeder Ferros Nobres ---
    if sheet_exists(NOME_ABA_FERRO_NOBRE):
        try:
            print(f"Iniciando leitura da aba '{NOME_ABA_FERRO_NOBRE}'...")
            df_ferro_nobre_raw = pd.read_excel(
                NOME_ARQUIVO_EXCEL_PRINCIPAL,
                sheet_name=NOME_ABA_FERRO_NOBRE,
                header=HEADER_ROW_FERRO_NOBRE,
                dtype=dtypes_leitura
            )
            df_processed_ferro_nobre = processar_dados_sheet(
                df_ferro_nobre_raw.copy(), COLUNAS_ORIGINAIS, "Ferro Nobre Processada"
            )
            if df_processed_ferro_nobre is not None:
                processed_dfs['Ferro Nobre Processada'] = df_processed_ferro_nobre
        except Exception as e:
            print(f"AVISO: Não foi possível ler ou processar a aba '{NOME_ABA_FERRO_NOBRE}'. Erro: {e}")
    else:
        print(f"AVISO: Aba esperada '{NOME_ABA_FERRO_NOBRE}' não encontrada. Pulando.")

    if not processed_dfs:
        print("Nenhuma aba foi processada com sucesso. Nenhuma saída será gerada.")
        return

    # --- Salva em um único arquivo Excel com abas separadas ---
    try:
        print(f"\nSalvando dados processados em: {ARQUIVO_SAIDA}...")
        with pd.ExcelWriter(ARQUIVO_SAIDA, engine='openpyxl') as writer:
            for sheet_name_output, df_to_write in processed_dfs.items():
                df_to_write.to_excel(writer, sheet_name=sheet_name_output, index=False)
        print(f"Dados salvos com sucesso em '{ARQUIVO_SAIDA}' com abas separadas.")
    except Exception as e:
        print(f"Erro ao salvar o arquivo Excel: {e}")
        return

    # Gera as abas convertidas (usa o arquivo de saída já escrito)
    gerar_planilhas_convertidas(processed_dfs, ARQUIVO_SAIDA)

def processar_planilha_pagamentos_separado_custom(arquivo_entrada, arquivo_saida):
    global NOME_ARQUIVO_EXCEL_PRINCIPAL, ARQUIVO_SAIDA
    NOME_ARQUIVO_EXCEL_PRINCIPAL = arquivo_entrada
    ARQUIVO_SAIDA = arquivo_saida
    processar_planilha_pagamentos_separado()

# Só executa diretamente se for rodado como script
if __name__ == "__main__":
    processar_planilha_pagamentos_separado()
