#!/usr/bin/env python3
import pandas as pd
from decimal import Decimal

def format_val(v):
    s = str(v).replace('.', '').replace(',', '.')
    try:
        num = float(s)
        return f"{num:.2f}"
    except:
        return ''

def fmt_competencia(c) -> str:
    # tenta padronizar para MM/AAAA
    if pd.isna(c):
        return ""
    if isinstance(c, pd.Timestamp):
        return c.strftime("%m/%Y")
    s = str(c).strip()
    # casos comuns: 2025-07-01, 01/07/2025, 07/2025
    try:
        dt = pd.to_datetime(s, dayfirst=True, errors="raise")
        return dt.strftime("%m/%Y")
    except Exception:
        # se já vier como MM/AAAA, mantém
        return s

def map_conta(tipo):
    tipo = str(tipo)
    if 'Cálculo Normal' in tipo or 'Folha Complementar' in tipo:
        return 1634
    elif 'Pensão Judicial Normal' in tipo:
        return 1637
    elif 'Rescisão Normal' in tipo:
        return 1634
    elif 'Férias' in tipo:
        return 313
    else:
        return 0

def map_historico(tipo):
    tipo = str(tipo)
    if 'Cálculo Normal' in tipo or 'Folha Complementar' in tipo:
        return 700
    elif 'Pensão Judicial Normal' in tipo:
        return 147
    elif 'Rescisão Normal' in tipo:
        return 132
    elif 'Férias' in tipo:
        return 139
    else:
        return 0

def normalize_name(n):
    if pd.isna(n):
        return ''
    s = str(n).strip()
    s = s.replace('"', '').replace("'", "")
    s = ' '.join(s.split())
    return s.upper()

def format_date_ddmmyyyy(d):
    if pd.isna(d):
        return ''
    s = str(d).strip()
    return ''.join(filter(str.isdigit, s))

def process_sheet(input_csv, output_excel, output_txt=None):
    rows = []
    with open(input_csv, encoding='latin1', errors='ignore') as f:
        for idx, line in enumerate(f):
            if idx <= 6:  # pula cabeçalho não estrutural
                continue
            cols = line.strip().split(';')
            if len(cols) < 8:
                continue
            for i in range(0, len(cols), 8):
                block = cols[i:i+8]
                if len(block) == 8:
                    rows.append(block)

    if not rows:
        raise RuntimeError("Nenhum dado válido encontrado no CSV.")

    df = pd.DataFrame(rows, columns=[
        'Chave Débito',
        'Contrato Empregado',
        'Nome',
        'Vencimento',
        'Tipo Movimento',           
        'Tipo_Para_Mapeamento',     
        'Situação Débito_Ignorar',  
        'Valor Débito'              
    ])

    print("\n--- Valores Originais de 'Valor Débito' ---")
    print(df['Valor Débito'].head(10))

    # Normaliza nome para override
    df['Nome_norm'] = df['Nome'].apply(normalize_name)

    df_out = pd.DataFrame({
        'Data': df['Tipo Movimento'].astype(str).apply(format_date_ddmmyyyy), # Aplica a formatação aqui
        'Nome': df['Nome_norm'],
        'Tipo Movimento': df['Vencimento'].astype(str),
        'Valor': df['Valor Débito'].apply(format_val),
        'Conta': df['Tipo_Para_Mapeamento'].apply(map_conta),
        'Histórico': df['Tipo_Para_Mapeamento'].apply(map_historico),
    })

    df_out['Competência'] = df['Vencimento'].astype(str).apply(fmt_competencia)

    print("\n--- Valores Formatados de 'Valor' (df_out) ---")
    print(df_out['Valor'].head(10))

    override = {'ALINE SCHROEDER', 'LUIZE SCHROEDER'}
    mask = df['Nome_norm'].isin(override)
    df_out.loc[mask, 'Conta'] = 1635
    df_out.loc[mask, 'Histórico'] = 48

    # Salva Excel
    df_out.to_excel(output_excel, index=False)
    print(f'► Excel gerado: {output_excel}')

    if output_txt:
        with open(output_txt, 'w', encoding='utf-8') as tx:
            for _, row in df_out.iterrows():
                # data sem prefixo "Competência:"
                data_campo = row['Data']

                if row.get('Competência') and str(row['Competência']).strip():
                    descricao = f"{row['Nome']} - Competencia: {row['Competência']}"
                else:
                    descricao = f"{row['Nome']}"

                line = (
                    f"1,"
                    f"{data_campo},"
                    f"{row['Conta']},"
                    f"5,"
                    f"{row['Valor']},"
                    f"{row['Histórico']},"
                    f"\"{descricao}\""
                )
                tx.write(line + '\n')
        print(f'► TXT gerado: {output_txt}')
