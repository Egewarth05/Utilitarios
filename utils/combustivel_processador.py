import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
import os
import csv

def to_float(val):
    try:
        return float(val.replace(",", "."))
    except:
        return None

def processar_combustivel(caminho_csv, valor_gasolina, valor_diesel, caminho_saida):
    fill_gasolina = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    fill_diesel = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
    alinhamento_centro = Alignment(horizontal="center", vertical="center")

    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "Relatório Combustível"

    # Legenda
    ws_out["B2"].fill = fill_gasolina
    ws_out["C2"] = "→ Gasolina"
    ws_out["E2"].fill = fill_diesel
    ws_out["F2"] = "→ Diesel"

    # Cabeçalho
    headers = ["Data", "Produto", "Nº da Nota", "Valor Total", "Quantidade", "TOTAL:"]
    for col_num, header in enumerate(headers, start=1):
        cell = ws_out.cell(row=6, column=col_num, value=header)
        cell.font = Font(bold=True)
        cell.border = thin_border
        cell.alignment = alinhamento_centro

    with open(caminho_csv, encoding="latin1") as f:
        reader = list(csv.reader(f, delimiter=";"))

    linha_saida = 7
    ultima_data = ""
    ultima_nota = ""

    for i in range(7, len(reader) - 3):
        try:
            linha_superior = reader[i]
            linha_produto = reader[i + 3]

            data_raw = linha_superior[1].strip() if len(linha_superior) > 1 else ""
            nota_raw = linha_superior[2].strip() if len(linha_superior) > 2 else ""
            if data_raw: ultima_data = data_raw
            if nota_raw: ultima_nota = nota_raw

            produto = linha_produto[1].strip()
            quantidade = linha_produto[3].strip()
            valor_total = linha_produto[5].strip()

            qtde_float = to_float(quantidade)
            if not qtde_float:
                continue

            if "gasolina" in produto.lower():
                valor_litro = valor_gasolina
                cor = fill_gasolina
            elif "diesel" in produto.lower():
                valor_litro = valor_diesel
                cor = fill_diesel
            else:
                continue

            total = round(qtde_float * valor_litro, 2)

            dados = [ultima_data, produto, ultima_nota, valor_total, qtde_float, total]
            for col, val in enumerate(dados, start=1):
                cell = ws_out.cell(row=linha_saida, column=col, value=val)
                cell.border = thin_border
                cell.alignment = alinhamento_centro
                if col == 6:
                    cell.fill = cor
                    cell.number_format = '#,##0.00'
                elif isinstance(val, float):
                    cell.number_format = '#,##0.00'

            linha_saida += 1
        except:
            continue

    for col in ['A', 'B', 'C', 'D', 'E', 'F']:
        ws_out.column_dimensions[col].width = 15

    wb_out.save(caminho_saida)
