#!/usr/bin/env python3

import os
import pprint
from utils.nf_comparador import extrair_relatorio

# Ajuste este caminho para onde está o seu PDF do Questor
PDF_PATH = os.path.join("uploads", "Questor_Exportacao.pdf")

if not os.path.isfile(PDF_PATH):
    print(f"[ERRO] PDF não encontrado em {PDF_PATH}")
    exit(1)

print(f"Extraindo relatório de: {PDF_PATH}\n")

rel = extrair_relatorio(PDF_PATH)

print("\n===== RESULTADO da extração (primeiras 10 linhas) =====")
pprint.pprint(rel[:10])
print(f"\nTotal de linhas extraídas: {len(rel)}")
