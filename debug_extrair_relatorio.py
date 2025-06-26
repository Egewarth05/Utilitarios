#!/usr/bin/env python3
import sys
import os

# Ajuste este path se o seu módulo estiver em outro lugar
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), ".")))

from utils.nf_comparador import extrair_relatorio

def main():
    # Pega o PDF via argumento de linha de comando ou usa o padrão
    pdf_path = sys.argv[1] if len(sys.argv) > 1 else "uploads/Questor_Exportacao.pdf"
    print(f"Extraindo relatório de: {pdf_path}\n")

    rel = extrair_relatorio(pdf_path)

    print("===== EXTRAÇÃO VIA BLOCOS (primeiras 10 linhas) =====")
    for item in rel[:10]:
        print(item)
    print()
    print(f"Total de linhas extraídas: {len(rel)}")

if __name__ == "__main__":
    main()
