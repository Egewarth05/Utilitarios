# utils/caixa_financeiro.py
import re
import unicodedata
import calendar
from datetime import date
from decimal import Decimal, InvalidOperation
from typing import List, Dict, Optional

LINE_RX = re.compile(
    r"^\s*"
    r"(?P<grupo>\d{2})\s+"
    r"(?P<grupo_desc>.+?)\s+"
    r"(?P<conta_codigo>\d{3})"
    r"(?P<conta_desc>.+?)\s+"
    r"(?P<valor1>\d[\d\.]*\.\d{2})"
    r"(?:\s+(?P<valor2>\d[\d\.]*\.\d{2}))?"
    r"\s*$",
    re.MULTILINE
)

START_RX = re.compile(r"Resumo\s+das\s+Contas", re.IGNORECASE)
END_RX = re.compile(
    r"^\s*-{5,}|^\s*\*{3,}\s*FIM|LAN[ÇC]AMENTOS\s+CANCELADOS",
    re.IGNORECASE | re.MULTILINE,
)

def _parse_amount(s: str) -> Optional[Decimal]:
    if not s:
        return None
    # Mantém só dígitos e ponto; remove milhares e deixa o último ponto como decimal
    s = re.sub(r"[^\d.]", "", s)
    if s.count(".") > 1:
        head, tail = s.rsplit(".", 1)   # última ocorrência é o decimal
        s = head.replace(".", "") + "." + tail
    try:
        return Decimal(s).quantize(Decimal("0.01"))
    except InvalidOperation:
        return None

def _decide_entrada_saida(grupo_desc: str, conta_desc: str,
                          v1: Optional[Decimal], v2: Optional[Decimal]) -> tuple[Decimal, Decimal]:
    """2 valores -> (entrada=v1, saída=v2).
       1 valor -> RECEIT*/RECEB*/SOBRA* ou ENTRADA/ABERTURA => entrada; senão => saída."""
    if v1 is None and v2 is None:
        return Decimal("0.00"), Decimal("0.00")

    if v1 is not None and v2 is not None:
        return v1, v2

    g = (grupo_desc or "").upper()
    c = (conta_desc or "").upper()

    entrada_like = (
        g.startswith("RECEIT")
        or g.startswith("RECEB")
        or "SOBRA" in g
        or "ENTRADA" in c
        or "ABERTURA" in c
    )

    v = v1 if v1 is not None else v2
    if entrada_like:
        return v, Decimal("0.00")
    else:
        return Decimal("0.00"), v

def parse_resumo_contas(txt_path: str) -> List[Dict]:
    text = open(txt_path, "r", encoding="latin-1", errors="ignore").read()

    mstart = START_RX.search(text)
    if not mstart:
        raise ValueError("Seção 'Resumo das Contas' não encontrada.")

    mend = END_RX.search(text, mstart.end())
    end_pos = mend.start() if mend else len(text)
    block = text[mstart.end():end_pos]

    rows: List[Dict] = []
    for m in LINE_RX.finditer(block):
        g = m.groupdict()
        v1 = _parse_amount(g.get("valor1"))
        v2 = _parse_amount(g.get("valor2"))
        entrada, saida = _decide_entrada_saida(g["grupo_desc"], g["conta_desc"], v1, v2)

        rows.append({
            "grupo": g["grupo"].strip(),
            "grupo_desc": g["grupo_desc"].strip(),
            "conta_codigo": g["conta_codigo"].strip(),
            "conta_desc": g["conta_desc"].strip(),
            "entrada": entrada,
            "saida": saida,
        })

    if not rows:
        raise ValueError("Nenhuma linha válida encontrada no 'Resumo das Contas'.")
    return rows

def _usa_primeiro_dia(grupo_desc: str, conta_desc: str) -> bool:
    texto = _norm(f"{grupo_desc} {conta_desc}")
    return ("VALOR ABERTURA DE CAIXA" in texto) or ("VALOR ABERTURA CAIXA" in texto)

def _primeiro_dia_mes(data_ddmmaaaa: str) -> str:
    mm = data_ddmmaaaa[2:4]
    aaaa = data_ddmmaaaa[4:8]
    return f"01{mm}{aaaa}"

def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.upper().strip()

MAPEAMENTO_FIXO: dict[str, tuple[str, str]] = {
    # RECEITAS
    _norm("CARTAO DE DÉBITO"): ("142", "21"),
    _norm("CARTAO DE DEBITO"): ("142", "21"),
    _norm("CREDIÁRIO"): ("142", "21"),
    _norm("CREDIARIO"): ("142", "21"),
    _norm("ENTRADAS DE CAIXA"): ("142", "21"),
    _norm("DINHEIRO/PIX SICOOB"): ("142", "21"),
    _norm("VALOR ABERTURA DE CAIXA"): ("25091", "66"),
    _norm("BAIXA TELE ENTREGA"): ("142", "21"),
    _norm("SOBRA CAIXA DIÁRIO"): ("142", "21"),
    _norm("CARTAO DE CREDITO"): ("142", "21"),
    _norm("CONVENIOS EMPRESARIAIS"): ("142", "21"),
    # EXEMPLOS DE DESPESAS/CUSTOS
    _norm("SISTEMA FIDELIDADE"): ("4556", "449"),
    _norm("BRINDE/PATROCI/RIFA/DOAÇÃ"): ("4556", "449"),
    _norm("MAT. DE USO E CONSUMO INT"): ("4546", "449"),
    _norm("ALUGUEL SALA COMERCIAL 2"): ("4430", "449"),
    _norm("ALUGUEL SALA COMERCIAL"): ("4430", "449"),
    _norm("CARTÓRIOS"): ("4555", "449"),
    _norm("SANGRIA"): ("25091", "66"),
    _norm("COMPRAS CONCORRÊNCIA"): ("3035", "337"),
    _norm("DEVOLUCAO CONVENIOS"): ("142", "21"),
    _norm("TRANSF. CX FINANCEIRO P/"): ("25091", "66"),
    _norm("RETORNO CLIENTE (FAT105)"): ("142", "21"),
    _norm("DEVOLUÇÃO VALOR P/ CLIENT"): ("142", "21"),
    _norm("FALTA CAIXA DIÁRIO"): ("142", "21"),
    _norm("SERV.TERCEIRIZADOS/ASSESS"): ("1496", "337"),
    _norm("ALUGUEL ESTACIONAMENTO"): ("4430", "449"),
    _norm("DESP. PESSOAL - HORA EXTRA"): ("4014", "449"),
    _norm("DESP. PESSOAL - VALE TRAN"): ("4014", "449"),
    _norm("PROLABORE SUELE FRANZEN"): ("680", "10003"),
    _norm("PROLABORE HELMUT FUHR"): ("680", "10003"),
    _norm("MATERIAIS EXPEDIENTE/PAPE"): ("4534", "449"),
    _norm("LIMPEZA/FAXINA (MUTIRÃO/P"): ("4546", "449"),
}

def _deve_excluir(grupo_desc: str, conta_desc: str) -> bool:

    texto = _norm(f"{grupo_desc} {conta_desc}")
    termos_banidos = (
    #Inserir os termos que quiser ignorar. Ex: "SANGRIA",
    )
    return any(t in texto for t in termos_banidos)

def _map_conta_historico(grupo_desc: str, conta_desc: str) -> tuple[str, str]:
    key = _norm(conta_desc)
    if key in MAPEAMENTO_FIXO:
        return MAPEAMENTO_FIXO[key]
    g = _norm(grupo_desc)
    # default de RECEITA
    if g.startswith("RECEIT") or g.startswith("RECEB") or "SOBRA" in g:
        return "142", "21"
    # default de DESPESA/CUSTO
    return "4698", "32"

DATE_RX = re.compile(r"\b(\d{2})[\/\-.](\d{2})[\/\-.](\d{2,4})\b")

def _canon_year(y: int) -> int:
    return 2000 + y if y < 100 and y <= 49 else (1900 + y if y < 100 else y)

def _inferir_data_ddmmaaaa_pelo_conteudo(full_text: str) -> str:
    contagem: dict[tuple[int, int], int] = {}
    latest: Optional[tuple[int, int, int]] = None  # (y, m, d)

    for d, m, y in DATE_RX.findall(full_text):
        try:
            dd = int(d); mm = int(m); yy = _canon_year(int(y))
        except ValueError:
            continue
        if not (1 <= mm <= 12 and 1 <= dd <= 31):
            continue
        contagem[(yy, mm)] = contagem.get((yy, mm), 0) + 1
        if latest is None or (yy, mm, dd) > latest:
            latest = (yy, mm, dd)

    if contagem:
        yy, mm = max(contagem.items(), key=lambda kv: kv[1])[0]
    elif latest:
        yy, mm = latest[0], latest[1]
    else:
        today = date.today()
        yy, mm = today.year, today.month

    last_day = calendar.monthrange(yy, mm)[1]
    return f"{last_day:02d}{mm:02d}{yy:04d}"

def _dec_to_str(x: Optional[Decimal], decimal_comma: bool) -> str:
    if x is None:
        x = Decimal("0.00")
    s = f"{x:.2f}"
    return s.replace(".", ",") if decimal_comma else s

def export_import_txt(rows: List[Dict], out_path: str, data_ddmmaaaa: str,
                      *, decimal: str = "dot",
                      prefixo: str = "1", contrapartida: str = "5") -> int:
    use_comma = (decimal == "comma")
    linhas_out: List[str] = []

    primeiro_dia = _primeiro_dia_mes(data_ddmmaaaa)

    for r in rows:
        if _deve_excluir(r["grupo_desc"], r["conta_desc"]):
            continue

        conta, hist = _map_conta_historico(r["grupo_desc"], r["conta_desc"])
        desc = r["conta_desc"]

        data_linha = primeiro_dia if _usa_primeiro_dia(r["grupo_desc"], r["conta_desc"]) else data_ddmmaaaa

        # ENTRADA
        if r.get("entrada") and r["entrada"] > 0:
            valor = _dec_to_str(r["entrada"], use_comma)
            linhas_out.append(
                f'{prefixo},{data_linha},{contrapartida},{conta},{valor},{hist},"{desc}"'
            )

        # SAÍDA (custo)
        if r.get("saida") and r["saida"] > 0:
            valor = _dec_to_str(r["saida"], use_comma)
            linhas_out.append(
                f'{prefixo},{data_linha},{conta},{contrapartida},{valor},{hist},"{desc}"'
            )

    with open(out_path, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(linhas_out))

    return len(linhas_out)

def processar_resumo_contas(in_txt_path: str, out_txt_path: str, *, decimal: str = "dot") -> Dict:
    """
    Lê o TXT, identifica a data (último dia do mês), extrai o Resumo das Contas,
    aplica filtros/mapeamentos e gera o TXT final.
    Retorna dict para a view: {"linhas": <int>, "saida": <path>, "decimal": <str>, "data": <ddmmaaaa>}
    """
    full_text = open(in_txt_path, "r", encoding="latin-1", errors="ignore").read()
    data_ddmmaaaa = _inferir_data_ddmmaaaa_pelo_conteudo(full_text)

    rows = parse_resumo_contas(in_txt_path)
    linhas_out = export_import_txt(rows, out_txt_path, data_ddmmaaaa, decimal=decimal)

    return {
        "linhas": linhas_out,
        "saida": out_txt_path,
        "decimal": decimal,
        "data": data_ddmmaaaa
    }
