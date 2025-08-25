import os
import re
import sys
from typing import Optional

def sgml_para_xml(raw: str) -> str:
    raw = raw[raw.find('<OFX>') : ]
    def fechar(m):
        tag, texto = m.group(1), m.group(2)
        return f"<{tag}>{texto}</{tag}>"
    pattern = re.compile(r"<([A-Z0-9]+)>([^<\r\n]+)(?![^<]*</\1>)")
    body = pattern.sub(fechar, raw)
    return '<?xml version="1.0" encoding="ISO-8859-1"?>\n' + body

_STMTTRN_RX = re.compile(r'(<STMTTRN>.*?</STMTTRN>)', re.IGNORECASE | re.DOTALL)

_HAS_MEMO_RX            = re.compile(r'<MEMO\b', re.IGNORECASE)
_NAME_CLOSED_RX         = re.compile(r'<NAME>(.*?)</NAME>', re.IGNORECASE | re.DOTALL)
_NAME_SINGLELINE_RX     = re.compile(r'^\s*<NAME>([^\r\n<]*)\s*$', re.IGNORECASE | re.MULTILINE)
_NAME_OPEN_UNCLOSED_RX  = re.compile(r'<NAME>([^<\r\n]+)', re.IGNORECASE)

def _remove_all_name(block: str) -> str:
    # Remove NAME fechado
    block = _NAME_CLOSED_RX.sub('', block)
    # Remove NAME em linha sem fechamento
    block = _NAME_SINGLELINE_RX.sub('', block)
    # Remove NAME aberto sem fechamento na mesma linha
    block = _NAME_OPEN_UNCLOSED_RX.sub('', block)
    return block

def _extract_first_name(block: str) -> Optional[str]:
    m = _NAME_CLOSED_RX.search(block)
    if m:
        return m.group(1).strip()
    m = _NAME_SINGLELINE_RX.search(block)
    if m:
        return m.group(1).strip()
    m = _NAME_OPEN_UNCLOSED_RX.search(block)
    if m:
        return m.group(1).strip()
    return None

def _to_memo_only(block: str) -> str:
    # Se já tem MEMO, só limpe quaisquer NAME restantes
    if _HAS_MEMO_RX.search(block):
        return _remove_all_name(block)

    # Senão, tenta extrair um NAME e convertê-lo para MEMO
    desc = _extract_first_name(block)
    if desc:
        # Remove todos os NAME
        block = _remove_all_name(block)
        # Insere MEMO antes de </STMTTRN>
        block = re.sub(r'</STMTTRN>', f'\n<MEMO>{desc}</MEMO>\n</STMTTRN>',
                       block, flags=re.IGNORECASE, count=1)
    return block

def processar_ofx_caixa(ofx_path, save_path):
    with open(ofx_path, 'r', encoding='ISO-8859-1') as f:
        txt = f.read()

    parts = txt.split('<OFX>', 1)
    if len(parts) == 2:
        cabecalho, corpo = parts[0], '<OFX>' + parts[1]
    else:
        cabecalho, corpo = '', txt

    def repl(m):
        try:
            return _to_memo_only(m.group(1))
        except Exception:
            # Se algo der errado, devolve o bloco original para não quebrar
            return m.group(1)

    corpo = _STMTTRN_RX.sub(repl, corpo)
    corpo = re.sub(r'><', '>\n<', corpo)

    out = (cabecalho + '\n' if cabecalho else '') + corpo
    with open(save_path, 'w', encoding='ISO-8859-1') as f:
        f.write(out)

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
