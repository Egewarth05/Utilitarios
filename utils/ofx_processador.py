import os
import re
import sys
import tkinter as tk
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox


def sgml_para_xml(raw: str) -> str:
    raw = raw[raw.find('<OFX>') : ]
    def fechar(m):
        tag, texto = m.group(1), m.group(2)
        return f"<{tag}>{texto}</{tag}>"
    pattern = re.compile(r"<([A-Z0-9]+)>([^<\r\n]+)(?![^<]*</\1>)")
    body = pattern.sub(fechar, raw)
    return '<?xml version="1.0" encoding="ISO-8859-1"?>\n' + body


def processar_ofx_caixa(ofx_path, save_path):
    with open(ofx_path, 'r', encoding='ISO-8859-1') as f:
        texto = f.read()

    parts = texto.split('<OFX>', 1)
    cabecalho = parts[0]
    corpo = '<OFX>' + parts[1]
    corpo = re.sub(r'><', '>\n<', corpo)
    linhas = corpo.splitlines()

    temp = []
    memo_valor = None
    for linha in linhas:
        s = linha.strip()
        if s.startswith('<MEMO>') and s.endswith('</MEMO>'):
            memo_valor = re.search(r'<MEMO>(.*?)</MEMO>', s).group(1)
            temp.append(s)
        elif memo_valor and s.startswith('<NAME>') and s.endswith('</NAME>'):
            name_valor = re.search(r'<NAME>(.*?)</NAME>', s).group(1)
            temp[-1] = f'<MEMO>{memo_valor} - {name_valor}</MEMO>'
            memo_valor = None
        else:
            temp.append(s)

    resultado = []
    resultado.extend([linha + '\n' for linha in cabecalho.splitlines()])
    indent = 0
    for linha in temp:
        if linha.startswith('</'):
            indent -= 1
        resultado.append(' ' * indent + linha + '\n')
        if linha.startswith('<') and not linha.startswith('</') and not re.match(r'<[^>]+>.*</[^>]+>$', linha):
            indent += 1

    with open(save_path, 'w', encoding='ISO-8859-1') as f:
        f.writelines(resultado)


def processar_ofx_sicoob(ofx_path, save_path):
    with open(ofx_path, 'r', encoding='ISO-8859-1') as f:
        linhas = f.readlines()
    out, buffer, inside = [], [], False
    name = checknum = ''
    for linha in linhas:
        if '<STMTTRN>' in linha:
            inside = True
            buffer = [linha]
            name = checknum = ''
            continue
        if inside:
            buffer.append(linha)
            m_name = re.search(r'<NAME>(.*?)</NAME>', linha)
            if m_name:
                name = m_name.group(1).strip()
            m_chk = re.search(r'<CHECKNUM>(.*?)</CHECKNUM>', linha)
            if m_chk:
                checknum = m_chk.group(1).strip()
            if '</STMTTRN>' in linha:
                for bl in buffer:
                    if '<MEMO>' in bl:
                        indent = bl[:bl.find('<')]
                        orig = re.search(r'<MEMO>(.*?)</MEMO>', bl).group(1).strip()
                        parts = [orig]
                        if name: parts.append(name)
                        if checknum: parts.append(checknum)
                        out.append(f"{indent}<MEMO>{' - '.join(parts)}</MEMO>\n")
                    else:
                        out.append(bl)
                inside = False
            continue
        out.append(linha)
    with open(save_path, 'w', encoding='ISO-8859-1') as f:
        f.writelines(out)


def processar_ofx(ofx_path, save_path, banco):
    if banco == "caixa":
        processar_ofx_caixa(ofx_path, save_path)
    elif banco == "sicoob":
        processar_ofx_sicoob(ofx_path, save_path)
    else:
        raise ValueError("Banco não suportado: " + banco)


if __name__ == "__main__":
    print("Este módulo não deve ser executado diretamente.")
