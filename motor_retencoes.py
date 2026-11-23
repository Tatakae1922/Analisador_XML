"""
Motor de Extracao de Retencoes em NFS-e - HEC Assessoria Contabil S/S Ltda.
(arquivo independente, sem interface grafica - para integrar em outro programa)

COMO USAR
---------
    from motor_retencoes_standalone import analisar_arquivo
    notas = analisar_arquivo("C:/caminho/nota.pdf")   # ou .xml
    for n in notas:                      # 1 dict por nota fiscal encontrada
        print(n["prestador"], n["cnpj_prestador"], n["tem_retencao"],
              n["total_retencoes_calculado"], n["valor_liquido"])

FORMATOS ACEITOS
  - PDF NFS-e legado municipal (ex.: Sao Paulo) - PIS/COFINS pela data de
    emissao (corte 01/07/2026, reforma tributaria)
  - PDF DANFSe v2.0 (padrao nacional) e PDFs com varias notas (relatorios/taloes)
  - PDF escaneado via OCR (se o Tesseract estiver instalado)
  - XML NFS-e padrao nacional v1.00 (leitura exata por campos)

CAMPOS DE CADA DICT (texto ja formatado em "R$ 1.234,56"; versao numerica
em chaves com "_" e sufixo _float, ex.: _iss_float):
  arquivo, paginas, layout, numero_nota, data_emissao, prestador,
  cnpj_prestador, tomador, cnpj_tomador, valor_total, pis, cofins, csll,
  inss, irrf, iss, total_retencoes_calculado, valor_liquido,
  retencoes_encontradas, tem_retencao (bool), obs_iss, obs_geral,
  discriminacao, erro

OUTROS ITENS UTEIS: formatar_situacao_iss(r, modo), COLUNAS_EXCEL,
COLUNAS_EXCEL_PRESTADOS, CAMPOS_NUMERICOS_EXCEL, OCR_DISPONIVEL.

DEPENDENCIAS: pip install pdfplumber   (opcional: pytesseract + Tesseract OCR)
"""

import os
import re
import shutil
import xml.etree.ElementTree as ET
from datetime import date
from pathlib import Path

import pdfplumber

# ══════════════════════════════════════════════════════════════════════════
#  OCR (para PDFs escaneados, ex.: Talao Fiscal Eletronico - TF-e)
# ══════════════════════════════════════════════════════════════════════════

try:
    import pytesseract
    _HAS_PYTESSERACT = True
except ImportError:
    _HAS_PYTESSERACT = False

_CAMINHOS_TESSERACT_CANDIDATOS = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    os.path.join(os.path.expanduser("~"), "Documents", "tesseract.exe"),
]


def _localizar_tesseract() -> str | None:
    env = os.environ.get("TESSERACT_CMD")
    if env and os.path.isfile(env):
        return env
    achou = shutil.which("tesseract")
    if achou:
        return achou
    for c in _CAMINHOS_TESSERACT_CANDIDATOS:
        if os.path.isfile(c):
            return c
    return None


_TESSERACT_CMD = _localizar_tesseract() if _HAS_PYTESSERACT else None
if _TESSERACT_CMD:
    pytesseract.pytesseract.tesseract_cmd = _TESSERACT_CMD
OCR_DISPONIVEL = bool(_TESSERACT_CMD)


def _ocr_pagina(pagina) -> str:
    """OCR de uma pagina escaneada (sem camada de texto). Devolve string
    vazia se o OCR nao estiver disponivel nesta maquina."""
    if not OCR_DISPONIVEL:
        return ""
    try:
        img = pagina.to_image(resolution=300).original
        return pytesseract.image_to_string(img, lang="por") or ""
    except Exception:
        return ""


# ══════════════════════════════════════════════════════════════════════════
#  HELPERS NUMERICOS
# ══════════════════════════════════════════════════════════════════════════

NUM = r"([0-9]{1,3}(?:\.[0-9]{3})*(?:,[0-9]+)|[0-9]+(?:,[0-9]+)?)"


def _norm(v):
    if v is None:
        return None
    v = str(v).strip()
    if v in ("", "-", "--"):
        return None
    if re.match(r"^\d+\.\d{2}$", v):
        return float(v)
    if re.search(r"\d\.\d{3},", v):
        v = v.replace(".", "").replace(",", ".")
    else:
        v = v.replace(",", ".")
    try:
        return float(v)
    except ValueError:
        return None


def _fmt(v):
    if v is None:
        return ""
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _buscar(texto, pats):
    for p in pats:
        m = re.search(p, texto, re.IGNORECASE | re.MULTILINE)
        if m:
            n = _norm(m.group(1))
            if n is not None:
                return _fmt(n), n
    return "", None


def _meta(texto, pats):
    for p in pats:
        m = re.search(p, texto, re.IGNORECASE | re.MULTILINE)
        if m:
            return m.group(1).strip()
    return ""


CAMPOS_RETENCAO = ["pis", "cofins", "csll", "inss", "irrf", "iss"]

COLUNAS_EXCEL = [
    ("arquivo", "Arquivo"), ("paginas", "Paginas"), ("layout", "Layout da NFS-e"),
    ("situacao", "Situacao"),
    ("numero_nota", "No Nota"), ("data_emissao", "Data Emissao"),
    ("prestador", "Prestador (quem reteve)"), ("cnpj_prestador", "CNPJ Prestador"),
    ("tomador", "Tomador"), ("valor_total", "Valor Total da Nota"),
    ("pis", "PIS Retido"), ("cofins", "COFINS Retido"), ("csll", "CSLL Retido"),
    ("inss", "INSS Retido"), ("irrf", "IRRF Retido"), ("iss", "ISS Retido"),
    ("total_retencoes_calculado", "Total Retido"),
    ("valor_liquido", "Valor Liquido (pos retencao)"),
    ("retencoes_encontradas", "Tributos Retidos"), ("obs_iss", "Situacao do ISS"),
    ("discriminacao", "Discriminacao do Servico"),
    ("obs_geral", "Observacoes"), ("erro", "Erro / Aviso"),
]

# Servicos PRESTADOS: aqui o prestador e sempre o proprio escritorio/cliente
# contabil (fixo), entao o foco passa a ser o TOMADOR (cliente que comprou o
# servico) e TODAS as notas entram no relatorio, com ou sem retencao.
COLUNAS_EXCEL_PRESTADOS = [
    ("arquivo", "Arquivo"), ("paginas", "Paginas"), ("layout", "Layout da NFS-e"),
    ("situacao", "Situacao"),
    ("numero_nota", "No Nota"), ("data_emissao", "Data Emissao"),
    ("tomador", "Tomador do Servico"), ("cnpj_tomador", "CNPJ Tomador"),
    ("valor_total", "Valor Total da Nota"),
    ("pis", "PIS Retido"), ("cofins", "COFINS Retido"), ("csll", "CSLL Retido"),
    ("inss", "INSS Retido"), ("irrf", "IRRF Retido"), ("iss", "ISS Retido"),
    ("total_retencoes_calculado", "Total Retido"),
    ("valor_liquido", "Valor Liquido (pos retencao)"),
    ("retencoes_encontradas", "Tributos Retidos"), ("obs_iss", "Situacao do ISS"),
    ("discriminacao", "Discriminacao do Servico"),
    ("obs_geral", "Observacoes"), ("erro", "Erro / Aviso"),
]

# Para as colunas monetarias, o campo interno guarda o texto ja formatado
# (ex.: "R$ 1.234,56"); este mapa aponta para o campo float correspondente,
# usado na exportacao para Excel para gravar numero de verdade (nao texto),
# permitindo formatacao contabil nativa e formulas de soma no relatorio.
CAMPOS_NUMERICOS_EXCEL = {
    "valor_total": "_valor_total_float",
    "pis": "_pis_float",
    "cofins": "_cofins_float",
    "csll": "_csll_float",
    "inss": "_inss_float",
    "irrf": "_irrf_float",
    "iss": "_iss_float",
    "total_retencoes_calculado": "_total_retencoes_float",
    "valor_liquido": "_valor_liquido_float",
}


# ══════════════════════════════════════════════════════════════════════════
#  SEPARACAO DO PDF EM BLOCOS (1 bloco = 1 nota fiscal)
# ══════════════════════════════════════════════════════════════════════════

_MARCA_DANFSE = re.compile(r"DANFSe", re.IGNORECASE)
_MARCA_LEGADO = re.compile(
    r"N[uú\W]?mero\s+da\s+Nota", re.IGNORECASE
)
_MARCA_LEGADO_TITULO = re.compile(
    r"NOTA\s+FISCAL\s+ELETR[ÔOo\W]NICA\s+DE\s+SERVI[ÇCç\W]OS", re.IGNORECASE
)


def _pagina_texto_e_tabelas(pagina) -> tuple[str, bool]:
    """Devolve (texto_da_pagina, usou_ocr). Se a pagina nao tiver camada
    de texto (PDF escaneado, ex.: Talao Fiscal Eletronico) e houver
    imagem, tenta OCR como ultimo recurso."""
    texto_puro = pagina.extract_text(layout=False) or ""
    usou_ocr = False
    if not texto_puro.strip() and pagina.images:
        texto_puro = _ocr_pagina(pagina)
        usou_ocr = bool(texto_puro.strip())
    partes = [texto_puro]
    for tbl in (pagina.extract_tables() or []):
        for row in tbl:
            partes.append("  ".join(str(c or "") for c in row))
    return "\n".join(partes), usou_ocr


def _e_inicio_de_nota(texto_pagina: str) -> bool:
    cabecalho = texto_pagina[:700]
    if _MARCA_DANFSE.search(cabecalho):
        return True
    # So o titulo basta (nao exige "Numero da Nota" junto): em paginas
    # vindas de OCR, rotulos curtos como "Nota" as vezes saem
    # corrompidos (ex.: "Mota"), mas o titulo da nota - mais longo e
    # distintivo - costuma sobreviver ao OCR.
    if _MARCA_LEGADO_TITULO.search(cabecalho):
        return True
    return False


def _dividir_nome_municipio(ws: list) -> str:
    """Dado um conjunto de palavras de uma mesma linha (com posicao x),
    separa o nome (coluna esquerda) do municipio/UF colado ao lado
    (coluna direita) achando o maior espaco horizontal entre palavras."""
    ws = sorted(ws, key=lambda w: w["x0"])
    if not ws:
        return ""
    maior_gap, idx_corte = 0, len(ws)
    for i in range(1, len(ws)):
        gap = ws[i]["x0"] - ws[i - 1]["x1"]
        if gap > maior_gap:
            maior_gap, idx_corte = gap, i
    if maior_gap < 10:
        return " ".join(w["text"] for w in ws)
    return " ".join(w["text"] for w in ws[:idx_corte])


def _extrair_nomes_por_posicao(pg) -> dict:
    """No layout DANFSe o nome do prestador/tomador fica colado ao
    municipio na mesma linha de texto (ex.: 'VR BENEFICIOS ... Sao Paulo
    / SP'), o que quebra uma extracao so por regex de texto. Aqui usamos
    as coordenadas (x,y) das palavras da propria pagina para separar as
    duas colunas corretamente."""
    try:
        words = pg.extract_words()
    except Exception:
        return {}
    linhas = {}
    for w in words:
        linhas.setdefault(round(w["top"]), []).append(w)
    tops = sorted(linhas.keys())

    def _linha_texto(top):
        return " ".join(w["text"] for w in sorted(linhas[top], key=lambda w: w["x0"]))

    def _nome_apos_rotulo(a_partir_de_top):
        for t in tops:
            if t < a_partir_de_top:
                continue
            if re.search(r"Nome\s*/\s*Nome\s+Empresarial", _linha_texto(t), re.IGNORECASE):
                for t2 in tops:
                    if t2 > t:
                        return _dividir_nome_municipio(linhas[t2]), t2
                return "", None
        return "", None

    nome_prestador, top_apos = _nome_apos_rotulo(0)
    nome_tomador, _ = _nome_apos_rotulo((top_apos or 0) + 1)
    return {"_nome_prestador_pos": nome_prestador, "_nome_tomador_pos": nome_tomador}


def extrair_blocos(caminho: Path) -> list[dict]:
    """Abre o PDF e devolve uma lista de blocos {texto, paginas, ocr}, um
    por nota."""
    blocos = []
    atual = None
    with pdfplumber.open(caminho) as pdf:
        for i, pg in enumerate(pdf.pages, start=1):
            texto_pag_completo, usou_ocr = _pagina_texto_e_tabelas(pg)
            if atual is None or _e_inicio_de_nota(texto_pag_completo):
                if atual is not None:
                    blocos.append(atual)
                atual = {"texto": texto_pag_completo, "paginas": [i], "ocr": usou_ocr}
                atual.update(_extrair_nomes_por_posicao(pg))
            else:
                atual["texto"] += "\n" + texto_pag_completo
                atual["paginas"].append(i)
                atual["ocr"] = atual["ocr"] or usou_ocr
        if atual is not None:
            blocos.append(atual)
    return blocos


def _detectar_layout(texto: str) -> str:
    cabecalho = texto[:700]
    if _MARCA_DANFSE.search(cabecalho):
        return "danfse"
    if _MARCA_LEGADO_TITULO.search(cabecalho):
        return "legado"
    return "desconhecido"


# ══════════════════════════════════════════════════════════════════════════
#  LAYOUT LEGADO (modelo proprio de prefeitura, ex. Sao Paulo pre-DANFSe)
# ══════════════════════════════════════════════════════════════════════════

# ══════════════════════════════════════════════════════════════════════════
#  PIS/COFINS no layout legado — o significado MUDOU com a reforma
#  tributaria a partir de 07/2026 (competencia de emissao):
#
#    ATE 06/2026: a coluna "COFINS (R$)" / "PIS/PASEP (R$)" da tabela
#    de tributos representava RETENCAO de verdade.
#
#    A PARTIR DE 07/2026: essa mesma coluna passou a ser so o tributo
#    que o PRESTADOR recolhe sobre a propria operacao (nao mais uma
#    retencao do tomador) — daí a nota vir com o aviso "Informacoes
#    preenchidas nos campos de PIS e COFINS sao referentes aos valores
#    totais sobre a operacao." em OUTRAS INFORMACOES.
#
#  A data que decide qual regra aplicar e a DATA DE EMISSAO da propria
#  nota, nao a data em que o programa roda.
# ══════════════════════════════════════════════════════════════════════════
_CORTE_REFORMA_PIS_COFINS = date(2026, 7, 1)


def _data_emissao_para_date(data_str: str):
    if not data_str:
        return None
    try:
        d, m, a = data_str.strip().split("/")
        return date(int(a), int(m), int(d))
    except Exception:
        return None


# Aviso que a propria NFS-e imprime quando os campos de PIS/COFINS sao
# o tributo da OPERACAO (regra nova), e nao retencao do tomador. Vale
# mais do que a data: ha notas emitidas ANTES de 07/2026 que ja saem
# com esse aviso (ex.: Verisure, emitida em 13/06/2026) porque o
# emissor antecipou a mudanca.
_AVISO_PIS_COFINS_SOBRE_OPERACAO = re.compile(
    r"PIS\s+e\s+COFINS.{0,120}?valores\s+totais\s+sobre\s+a\s+opera",
    re.IGNORECASE | re.DOTALL,
)


def _eh_pre_reforma_pis_cofins(data_emissao_str: str, texto: str = "") -> bool:
    """True = tratar PIS/COFINS da tabela como RETENCAO de verdade
    (comportamento ate 06/2026).

    Decide em duas etapas:
      1. Se a nota TRAZ o aviso de que PIS/COFINS sao valores totais
         sobre a operacao, vale o aviso (= pos-reforma), nao importa a
         data — ha emissores que anteciparam a mudanca.
      2. Senao, decide pela data de emissao (corte 01/07/2026).

    Sem data legivel e sem aviso, assume-se o comportamento NOVO
    (pos-reforma) por seguranca — contar retencao a mais e pior do que
    contar a menos."""
    if texto and _AVISO_PIS_COFINS_SOBRE_OPERACAO.search(texto):
        return False
    d = _data_emissao_para_date(data_emissao_str)
    if d is None:
        return False
    return d < _CORTE_REFORMA_PIS_COFINS


_PADROES_PIS_COFINS_RETENCAO_PRE_REFORMA = {
    "pis": [
        rf"PIS\s+Ret\.?:?\s*{NUM}",
        rf"PIS\W?PASEP\s*\(R\$\)\s*\n\s*{NUM}",
        rf"PIS\W?PASEP\s*\(R\$\)\s+{NUM}",
    ],
    "cofins": [
        rf"COFINS\s+Ret\.?:?\s*{NUM}",
        rf"COFINS\s*\(R\$\)\s*\n\s*{NUM}",
        rf"\bCOFINS\b[\s:R$]*{NUM}",
    ],
}
_PADROES_PIS_COFINS_APENAS_EXPLICITO_POS_REFORMA = {
    # Pos-reforma, so conta como retencao se o texto disser "Ret."
    # explicitamente (ex.: "PIS Ret.:2,06") — a coluna generica da
    # tabela (sem a palavra "Ret.") deixou de ser retencao.
    "pis": [rf"PIS\s+Ret\.?:?\s*{NUM}"],
    "cofins": [rf"COFINS\s+Ret\.?:?\s*{NUM}"],
}
_PADROES_TRIBUTO_OPERACAO_LEGADO = {
    "pis": [rf"PIS\W?PASEP\s*\(R\$\)\s*\n\s*{NUM}", rf"PIS\W?PASEP\s*\(R\$\)\s+{NUM}"],
    "cofins": [rf"COFINS\s*\(R\$\)\s*\n\s*{NUM}", rf"\bCOFINS\b[\s:R$]*{NUM}"],
}


def _nota_tributo_operacao_legado(texto: str) -> str:
    """So chamada quando POS-reforma (>= 07/2026): PIS/COFINS
    informados na nota (coluna sem 'Retido' no nome) sao so o tributo
    que o PRESTADOR recolhe sobre a propria operacao — nao retencao do
    tomador. Devolve uma frase explicando isso quando esses valores
    existem e sao maiores que zero, para o contador nao estranhar ver
    o valor na NFS-e mas nao no relatorio."""
    _, pis_v = _buscar(texto, _PADROES_TRIBUTO_OPERACAO_LEGADO["pis"])
    _, cofins_v = _buscar(texto, _PADROES_TRIBUTO_OPERACAO_LEGADO["cofins"])
    partes = []
    if pis_v:
        partes.append(f"PIS {_fmt(pis_v)}")
    if cofins_v:
        partes.append(f"COFINS {_fmt(cofins_v)}")
    if not partes:
        return ""
    return ("Tributo sobre a operacao informado na nota (NAO e retencao do tomador, "
            "regra valida a partir de 07/2026): " + " / ".join(partes) + ".")


_PADROES_LEGADO = {
    "csll": [
        rf"CSL(?:L)?\s+Ret\.?:?\s*{NUM}",
        rf"Contribui[çc\W][õo\W]es\s+Sociais\s*-\s*Retidas?\s*\(R\$\)\s*\n?\s*{NUM}",
        rf"CSLL\s*\(R\$\)\s*\n\s*{NUM}",
        rf"CSLL[\s:R$]*{NUM}",
        rf"contribui[çc\W][aã\W]o\s+social[\s:R$]*{NUM}",
    ],
    "inss": [
        rf"Contribui[çc\W][aã\W]o\s+Previdenci[aá\W]ria\s*-\s*Retida\s*\(R\$\)\s*\n?\s*{NUM}",
        rf"INSS\s+Ret\.?:?\s*{NUM}",
        rf"INSS\s*\(R\$\)\s*\n\s*{NUM}",
        rf"\bINSS\b[\s:R$]*{NUM}",
        rf"previd[êe\W]ncia\s+social[\s:R$]*{NUM}",
    ],
    "irrf": [
        rf"IRRF\s*\(R\$\)\s*\n\s*{NUM}",
        rf"IRRF\s+[\d,]+\s*%[:\s]*R\$\s*{NUM}",
        rf"IRRF\s+Ret\.?:?\s*{NUM}",
        rf"IR\s+Ret\.?:?\s*{NUM}",
        rf"imposto\s+de\s+renda[\s:R$]*{NUM}",
        rf"\bIRRF\b[\s:R$]*{NUM}",
    ],
    "valor_total": [
        rf"VALOR\s+TOTAL\s+DO\s+SERVI[ÇC\W]O\s*=\s*R\$\s*{NUM}",
        rf"VALOR\s+TOTAL\s+COBRADO\s*=\s*R\$\s*{NUM}",
        rf"valor\s+total[\s:R$]*{NUM}",
    ],
}

_META_PADROES_LEGADO = {
    "numero_nota": [r"N[úu\W]mero\s+da\s+[NM]ota\s*\n?\s*([0-9]+)", r"NFS?-?e?\s*[Nn][°º#]?\s*([0-9]+)", r"RPS\s+N[°º]\s*([0-9]+)", r"\b(\d{5,8})\b\s*\n?\s*SECRETARIA\s+MUNICIPAL"],
    "data_emissao": [r"Data\s+e\s+Hora\s+de\s+Emiss[aã\W]o[\s\S]{0,4}?([0-9]{2}/[0-9]{2}/[0-9]{4})", r"emiss[aã\W]o[\s\S]{0,4}?([0-9]{2}[/.-][0-9]{2}[/.-][0-9]{2,4})"],
    "cnpj_prestador": [r"PRESTADOR[\s\S]*?CPF/CNPJ:\s*([\d./-]+)", r"CPF/CNPJ:\s*([\d./-]+)"],
}

_PAT_NOME_LEGADO = [
    r"Nome/Raz[aã\W]o\s+Social:\s*([^\n]+)",
    r"Social:\s*([^\n]+)",
]


def _extrair_prestador_tomador_legado(texto: str) -> dict:
    """Nome e CNPJ de prestador/tomador, com a busca ESCOPADA a cada
    secao (PRESTADOR .. TOMADOR .. INTERMEDIARIO), para nao pegar por
    engano o nome/CNPJ da secao errada - risco real quando o texto vem
    de OCR (rotulos como 'Nome' podem sair corrompidos)."""
    bloco_prest = _bloco_secao(texto, r"PRESTADOR\s+DE\s+SERVI[ÇC\W]OS", r"TOMADOR\s+DE\s+SERVI[ÇC\W]OS")
    bloco_tom = _bloco_secao(texto, r"TOMADOR\s+DE\s+SERVI[ÇC\W]OS", r"INTERMEDI[ÁA\W]RIO\s+DE\s+SERVI[ÇC\W]OS")
    if not bloco_tom:
        bloco_tom = _bloco_secao(texto, r"TOMADOR\s+DE\s+SERVI[ÇC\W]OS", r"DISCRIMINA[ÇC\W][ÃA\W]O")

    def _nome(bloco):
        return _meta(bloco, _PAT_NOME_LEGADO)

    def _cnpj(bloco):
        # tolera espaco espurio de OCR entre o bloco final e o traco (ex.: "0001 -46")
        m = re.search(r"CPF/CNPJ:\s*([\d./]+\s*-?\s*[\d]*)", bloco, re.IGNORECASE)
        if not m:
            return ""
        return re.sub(r"\s+", "", m.group(1))

    return {
        "prestador": _nome(bloco_prest),
        "cnpj_prestador": _cnpj(bloco_prest),
        "tomador": _nome(bloco_tom),
        "cnpj_tomador": _cnpj(bloco_tom),
    }

_ISS_RETIDO_LEGADO = [r"ISS.*ser[aá\W]\s+RETIDO\s+pelo\s+Tomador", r"ISS.*retido\s+pelo\s+tomador", r"ISS\s+ser[aá\W]\s+retido", r"recolhimento.*ISS.*tomador", r"tomador.*reter.*ISS", r"ISS\s+Ret\.?:"]
_ISS_NAO_RETIDO_LEGADO = [r"Data\s+de\s+vencimento\s+do\s+ISS\s+desta\s+NFS-e", r"recolher\s+o\s+ISS\s+desta\s+NFS-e"]

_PAT_ISS_TABELA_LEGADO = (
    r"Valor\s+Total\s+das\s+Dedu[çc\W][õo\W]es[^\n]+Valor\s+do\s+ISS[^\n]*\n"
    r"\s*" + NUM + r"\s+" + NUM + r"\s+[\d,]+%\s+" + NUM
)

_PAT_TABELA_INLINE_LEGADO = (
    r"(INSS)\s*\(R\$\)\s+"
    r"(IRRF)\s*\(R\$\)\s+"
    r"(CSLL)\s*\(R\$\)\s+"
    r"(COFINS)\s*\(R\$\)\s+"
    r"(PIS\W?PASEP)\s*\(R\$\)[^\n]*\n"
    r"\s*([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)\s+([0-9,]+)\s+([0-9,.-]+)"
)

_PAT_DISCRIMINACAO_LEGADO = re.compile(
    r"DISCRIMINA[ÇC\W][ÃA\W]O\s+D[EO]S?\s+SERVI[ÇC\W]OS\s*\n(.*?)"
    r"(?:VALOR\s+TOTAL\s+DO\s+SERVI[ÇC\W]O|C[óo\W]digo\s+do\s+Servi[çc\W]o)",
    re.IGNORECASE | re.DOTALL,
)


def _tabela_inline_legado(texto: str, incluir_pis_cofins: bool) -> dict:
    """INSS/IRRF/CSLL sempre entram como retencao nessa tabela. COFINS
    e PIS/PASEP so entram quando `incluir_pis_cofins` for True (nota
    PRE-reforma, emitida ate 06/2026) — pos-reforma eles ficam de fora
    de proposito (ver comentario acima de _PADROES_LEGADO); quem
    precisar do valor deles pos-reforma usa
    _nota_tributo_operacao_legado, nao este dicionario."""
    resultado = {}
    m = re.search(_PAT_TABELA_INLINE_LEGADO, texto, re.IGNORECASE)
    if m:
        mapa = {"inss": m.group(6), "irrf": m.group(7), "csll": m.group(8)}
        if incluir_pis_cofins:
            mapa["cofins"] = m.group(9)
            mapa["pis"] = m.group(10)
        for campo, val in mapa.items():
            n = _norm(val)
            resultado[campo] = (_fmt(n), n)
    return resultado


_PAT_CAMPO_AGREGADO_CONTRIB_SOCIAIS = re.compile(
    r"Contribui[çc\W][õo\W]es\s+Sociais\s*-\s*Retidas?\s*\(R\$\)", re.IGNORECASE)
_PAT_CSLL_EXPLICITO = re.compile(rf"CSL(?:L)?\s+Ret\.?:?\s*{NUM}", re.IGNORECASE)


def _nota_contribuicoes_sociais_agregadas(texto: str, csll_valor) -> str:
    """No layout legado, o campo da tabela chamado "Contribuicoes
    Sociais - Retidas (R$)" NAO e so a CSLL: ele soma PIS + COFINS +
    CSLL federais (4,65% quando cheio). Como o motor guarda esse valor
    na coluna "csll" (nao ha coluna propria), devolve um aviso para o
    contador nao concluir que PIS e COFINS ficaram sem retencao.

    So avisa quando o valor veio MESMO desse campo agregado, ou seja,
    quando a nota nao traz "CSL Ret.:" explicito na discriminacao (nesse
    caso o valor e CSLL de verdade e nao ha o que avisar)."""
    if not csll_valor:
        return ""
    if _PAT_CSLL_EXPLICITO.search(texto):
        return ""
    if not _PAT_CAMPO_AGREGADO_CONTRIB_SOCIAIS.search(texto):
        return ""
    return (f"O valor {_fmt(csll_valor)} na coluna CSLL veio do campo "
            "\"Contribuicoes Sociais - Retidas\" da nota, que soma PIS + COFINS + "
            "CSLL federais (nao e so CSLL) - conferir a composicao na nota.")


_PAT_ISS_NA_DISCRIMINACAO = re.compile(rf"ISS\s+Ret\.?:?\s*{NUM}", re.IGNORECASE)


def _divergencia_iss_legado(texto: str, iss_da_tabela) -> str:
    """A nota traz o ISS em DOIS lugares: o campo "Valor do ISS (R$)"
    da tabela e, as vezes, escrito a mao na discriminacao (ex.: "ISS
    Ret.:6,34"). Quando os dois existem e nao batem, avisa - e
    diferenca de centavos que reaparece depois na conciliacao.

    O valor mantido e sempre o da TABELA (campo oficial da nota); a
    discriminacao e texto livre digitado pelo emissor."""
    if iss_da_tabela is None:
        return ""
    m = _PAT_ISS_NA_DISCRIMINACAO.search(texto)
    if not m:
        return ""
    iss_no_texto = _norm(m.group(1))
    if iss_no_texto is None or abs(iss_no_texto - iss_da_tabela) < 0.005:
        return ""
    return (f"ISS divergente dentro da propria nota: campo da tabela {_fmt(iss_da_tabela)} "
            f"x discriminacao {_fmt(iss_no_texto)} - foi usado o da tabela.")


def _resolver_iss_legado(texto, fmt_tab, num_tab):
    retido = any(re.search(p, texto, re.IGNORECASE) for p in _ISS_RETIDO_LEGADO)
    nao_retido = any(re.search(p, texto, re.IGNORECASE) for p in _ISS_NAO_RETIDO_LEGADO)
    if retido and not nao_retido:
        return fmt_tab, num_tab, "ISS retido pelo tomador"
    if nao_retido and not retido:
        return "", None, "ISS recolhido pelo prestador"
    if retido and nao_retido:
        return "", None, "ISS: sinais contraditorios - conferir manualmente"
    return "", None, ""


def _analisar_legado(texto: str) -> dict:
    r = {"layout": "SP / Municipal (legado)"}

    for campo, pats in _META_PADROES_LEGADO.items():
        r[campo] = _meta(texto, pats)
    r.update(_extrair_prestador_tomador_legado(texto))

    # PIS/COFINS dependem da DATA DE EMISSAO da nota (regra mudou com a
    # reforma tributaria a partir de 07/2026 - ver comentario acima de
    # _CORTE_REFORMA_PIS_COFINS).
    pre_reforma = _eh_pre_reforma_pis_cofins(r.get("data_emissao"), texto)
    padroes_pis_cofins = (_PADROES_PIS_COFINS_RETENCAO_PRE_REFORMA if pre_reforma
                          else _PADROES_PIS_COFINS_APENAS_EXPLICITO_POS_REFORMA)

    for campo, pats in _PADROES_LEGADO.items():
        fmt, num = _buscar(texto, pats)
        r[campo] = fmt
        r[f"_{campo}_float"] = num
    for campo, pats in padroes_pis_cofins.items():
        fmt, num = _buscar(texto, pats)
        r[campo] = fmt
        r[f"_{campo}_float"] = num

    tabela_inline = _tabela_inline_legado(texto, incluir_pis_cofins=pre_reforma)
    for campo, (fmt, num) in tabela_inline.items():
        if not r.get(f"_{campo}_float") and num and num > 0:
            r[campo] = fmt
            r[f"_{campo}_float"] = num

    m_iss = re.search(_PAT_ISS_TABELA_LEGADO, texto, re.IGNORECASE | re.MULTILINE)
    if m_iss:
        nt = _norm(m_iss.group(3))
        ft = _fmt(nt)
    else:
        nt = ft = None
    r["iss"], r["_iss_float"], r["obs_iss"] = _resolver_iss_legado(texto, ft, nt)

    m_disc = _PAT_DISCRIMINACAO_LEGADO.search(texto)
    r["discriminacao"] = re.sub(r"[ \t]+", " ", m_disc.group(1)).strip() if m_disc else ""

    r["_valor_liquido_declarado_float"] = None
    # O aviso de "tributo sobre a operacao" existe para explicar um
    # valor que aparece NA NOTA mas NAO no relatorio. Se o PIS/COFINS
    # foi reconhecido como retencao de verdade (a nota traz "PIS Ret.:"
    # explicito), ele ESTA no relatorio - avisar o contrario seria
    # contradizer as proprias colunas.
    pis_cofins_entraram = bool(r.get("_pis_float") or r.get("_cofins_float"))
    avisos = [
        "" if (pre_reforma or pis_cofins_entraram) else _nota_tributo_operacao_legado(texto),
        _nota_contribuicoes_sociais_agregadas(texto, r.get("_csll_float")),
        _divergencia_iss_legado(texto, r.get("_iss_float")),
    ]
    r["obs_geral"] = " | ".join(a for a in avisos if a)
    return r


# ══════════════════════════════════════════════════════════════════════════
#  LAYOUT DANFSe v2.0 (padrao nacional)
# ══════════════════════════════════════════════════════════════════════════

_PAT_NUMERO_DANFSE = re.compile(
    r"N[ÚU]MERO\s+DA\s+NFS-e[^\n]*\n\s*([0-9]+)\s+([0-9]{2}/[0-9]{2}/[0-9]{4})\s+"
    r"([0-9]{2}/[0-9]{2}/[0-9]{4})\s+([0-9:]+)",
    re.IGNORECASE,
)
_PAT_MUNICIPIO_DANFSE = re.compile(r"Munic[íi]pio:\s*([^\n]+?)(?:\s{2,}|\n)", re.IGNORECASE)

_PAT_CNPJ = re.compile(r"([\d]{2}\.[\d]{3}\.[\d]{3}/[\d]{4}-[\d]{2})")
_PAT_NOME_EMPRESARIAL = re.compile(
    r"Nome\s*/\s*Nome\s+Empresarial[^\n]*\n\s*(.+?)\s+"
    r"[A-ZÀ-Ú][\wÀ-ÿ.'-]*(?:\s[A-ZÀ-Ú][\wÀ-ÿ.'-]*)*\s*/\s*[A-Z]{2}\s+[\d]",
    re.IGNORECASE,
)

_PAT_VALORES_DANFSE = re.compile(
    r"VALOR\s+TOTAL\s+DA\s+NFS-e\s+VALOR\s+DA\s+OPERA[ÇC][ÃA]O\s*/\s*SERVI[ÇC]O[^\n]*\n"
    r"\s*(?:R\$\s*)?([0-9.,]+|-)\s+(?:R\$\s*)?([0-9.,]+|-)",
    re.IGNORECASE,
)
_PAT_LIQUIDO_DANFSE = re.compile(
    r"Total\s+das\s+Reten[çc][õo]es\s*\(ISSQN\s*/\s*Federais\)\s+VALOR\s+L[ÍI]QUIDO\s+DA\s+NFS-e[^\n]*\n"
    r"\s*(?:R\$\s*)?([0-9.,]+|-)\s+(?:R\$\s*)?([0-9.,]+|-)",
    re.IGNORECASE,
)
_PAT_ISSQN_DANFSE = re.compile(
    r"BC\s+ISSQN\s+Al[íi]quota\s+Aplicada\s+Reten[çc][ãa]o\s+do\s+ISSQN\s+ISSQN\s+Apurado[^\n]*\n"
    r"\s*(?:R\$\s*)?([0-9.,]+|-)\s+([0-9,]+%|-)\s+"
    r"(Retido(?:\s+pelo\s+Tomador)?|N[ãa]o\s+Retido)\s+(?:R\$\s*)?([0-9.,]+|-)",
    re.IGNORECASE,
)
_PAT_FEDERAL_DANFSE = re.compile(
    r"TRIBUTA[ÇC][ÃA]O\s+FEDERAL\s*\(EXCETO\s+CBS\)\s+IRRF\s+Contribui[çc][ãa]o\s+Previdenci[áa]ria\s*-\s*Retida\s+"
    r"Contribui[çc][õo]es\s+Sociais\s*-\s*Retidas[^\n]*\n"
    r"\s*([0-9.,]+|-)\s+([0-9.,]+|-)\s+([0-9.,]+|-)",
    re.IGNORECASE,
)
_PAT_PIS_COFINS_DEBITO_DANFSE = re.compile(
    r"PIS\s*-\s*D[ée]bito\s+Apura[çc][ãa]o\s+Pr[óo]pria\s+COFINS\s*-\s*D[ée]bito\s+Apura[çc][ãa]o\s+Pr[óo]pria[^\n]*\n"
    r"\s*(?:R\$\s*)?([0-9.,]+|-)\s+(?:R\$\s*)?([0-9.,]+|-)",
    re.IGNORECASE,
)
_PAT_DISCRIMINACAO_DANFSE = re.compile(
    r"Descri[çc][ãa]o\s+do\s+Servi[çc]o\s*\n(.*?)"
    r"(?:TRIBUTA[ÇC][ÃA]O\s+MUNICIPAL|TRIBUTA[ÇC][ÃA]O\s+FEDERAL)",
    re.IGNORECASE | re.DOTALL,
)


def _bloco_secao(texto: str, inicio_pat: str, fim_pat: str) -> str:
    m = re.search(inicio_pat + r"(.*?)" + fim_pat, texto, re.IGNORECASE | re.DOTALL)
    return m.group(1) if m else ""


def _analisar_danfse(texto: str, bloco: dict | None = None) -> dict:
    r = {"layout": "Padrao Nacional (DANFSe)"}
    bloco = bloco or {}

    m = _PAT_NUMERO_DANFSE.search(texto)
    if m:
        r["numero_nota"] = m.group(1)
        r["data_emissao"] = m.group(3)
    else:
        r["numero_nota"] = ""
        r["data_emissao"] = ""

    bloco_prestador = _bloco_secao(texto, r"PRESTADOR\s*/\s*FORNECEDOR", r"TOMADOR\s*/\s*ADQUIRENTE")
    cnpj_m = _PAT_CNPJ.search(bloco_prestador)
    r["cnpj_prestador"] = cnpj_m.group(1) if cnpj_m else ""

    nome_pos = bloco.get("_nome_prestador_pos") or ""
    if nome_pos:
        r["prestador"] = nome_pos
    else:
        nome_m = _PAT_NOME_EMPRESARIAL.search(bloco_prestador)
        r["prestador"] = re.sub(r"\s+", " ", nome_m.group(1)).strip() if nome_m else ""

    bloco_tomador = _bloco_secao(
        texto, r"TOMADOR\s*/\s*ADQUIRENTE",
        r"DESTINAT[ÁA]RIO\s+DA\s+OPERA[ÇC][ÃA]O|INTERMEDI[ÁA]RIO\s+DA\s+OPERA[ÇC][ÃA]O|SERVI[ÇC]O\s+PRESTADO",
    )
    nome_tom_pos = bloco.get("_nome_tomador_pos") or ""
    if nome_tom_pos:
        r["tomador"] = nome_tom_pos
    else:
        nome_tom = _PAT_NOME_EMPRESARIAL.search(bloco_tomador)
        r["tomador"] = re.sub(r"\s+", " ", nome_tom.group(1)).strip() if nome_tom else ""
    cnpj_tom_m = _PAT_CNPJ.search(bloco_tomador)
    r["cnpj_tomador"] = cnpj_tom_m.group(1) if cnpj_tom_m else ""

    m_disc = _PAT_DISCRIMINACAO_DANFSE.search(texto)
    r["discriminacao"] = re.sub(r"[ \t]+", " ", m_disc.group(1)).strip() if m_disc else ""

    mv = _PAT_VALORES_DANFSE.search(texto)
    valor_operacao = _norm(mv.group(2)) if mv else None
    valor_total_nfse = _norm(mv.group(1)) if mv else None
    valor_total = valor_operacao if valor_operacao is not None else valor_total_nfse
    r["valor_total"] = _fmt(valor_total)
    r["_valor_total_float"] = valor_total

    mliq = _PAT_LIQUIDO_DANFSE.search(texto)
    r["_valor_liquido_declarado_float"] = _norm(mliq.group(2)) if mliq else None

    # PIS e COFINS no DANFSe aparecem so como "Debito Apuracao Propria"
    # (o que o PRESTADOR recolhe sobre a propria operacao) - nunca como
    # retencao do tomador; por isso ficam sempre de fora do total
    # retido, so viram uma nota informativa quando houver valor.
    r["pis"] = ""
    r["_pis_float"] = None
    r["cofins"] = ""
    r["_cofins_float"] = None
    obs = []

    mdeb = _PAT_PIS_COFINS_DEBITO_DANFSE.search(texto)
    if mdeb:
        pis_deb, cofins_deb = _norm(mdeb.group(1)), _norm(mdeb.group(2))
        partes_deb = []
        if pis_deb:
            partes_deb.append(f"PIS {_fmt(pis_deb)}")
        if cofins_deb:
            partes_deb.append(f"COFINS {_fmt(cofins_deb)}")
        if partes_deb:
            obs.append("Debito de apuracao propria do prestador (NAO e retencao do tomador): "
                        + " / ".join(partes_deb) + ".")

    mfed = _PAT_FEDERAL_DANFSE.search(texto)
    if mfed:
        irrf_v = _norm(mfed.group(1))
        inss_v = _norm(mfed.group(2))
        cs_v = _norm(mfed.group(3))
    else:
        irrf_v = inss_v = cs_v = None
    r["irrf"], r["_irrf_float"] = _fmt(irrf_v), irrf_v
    r["inss"], r["_inss_float"] = _fmt(inss_v), inss_v
    r["csll"], r["_csll_float"] = _fmt(cs_v), cs_v
    if cs_v:
        obs.append("\"CSLL\" no padrao DANFSe e o bloco unico \"Contribuicoes "
                    "Sociais - Retidas\", que soma PIS+COFINS+CSLL federais.")

    miss = _PAT_ISSQN_DANFSE.search(texto)
    if miss:
        status = miss.group(3)
        issqn_apurado = _norm(miss.group(4))
        if re.match(r"Retido", status, re.IGNORECASE):
            r["iss"], r["_iss_float"] = _fmt(issqn_apurado), issqn_apurado
            r["obs_iss"] = "ISS retido pelo tomador (campo 'Retencao do ISSQN')"
        else:
            r["iss"], r["_iss_float"] = "", None
            r["obs_iss"] = "ISS recolhido pelo prestador (campo 'Retencao do ISSQN' = Nao Retido)"
    else:
        r["iss"], r["_iss_float"] = "", None
        r["obs_iss"] = ""

    r["obs_geral"] = " | ".join(obs)
    return r


# ══════════════════════════════════════════════════════════════════════════
#  FALLBACK GENERICO (layout nao mapeado — outras prefeituras)
# ══════════════════════════════════════════════════════════════════════════

_PAT_VALOR_LIQUIDO_GENERICO = [
    rf"Valor\s+L[íi]quido\s+d[aeo]s?\s+(?:NFS-?e|Nota|Documento)[^\n]*?{NUM}",
    rf"Valor\s+L[íi]quido[^\n]*?R\$\s*{NUM}",
]
_PAT_VALOR_OPERACAO_GENERICO = [
    rf"Valor\s+(?:da\s+Opera[çc][ãa]o|Total\s+d[oe]s?\s+Servi[çc]os?|dos?\s+Servi[çc]os?)[^\n]*?{NUM}",
]


def _analisar_generico(texto: str) -> dict:
    r = {"layout": "Nao identificado (fallback generico)"}
    for campo in ("numero_nota", "data_emissao", "prestador", "cnpj_prestador", "tomador",
                  "cnpj_tomador", "discriminacao"):
        r[campo] = ""
    for campo in ("pis", "cofins", "csll", "inss", "irrf"):
        r[campo] = ""
        r[f"_{campo}_float"] = None
    r["_valor_liquido_declarado_float"] = None

    _, liquido = _buscar(texto, _PAT_VALOR_LIQUIDO_GENERICO)
    _, operacao = _buscar(texto, _PAT_VALOR_OPERACAO_GENERICO)
    r["valor_total"] = _fmt(operacao if operacao is not None else liquido)
    r["_valor_total_float"] = operacao if operacao is not None else liquido

    if liquido is not None and operacao is not None and abs(liquido - operacao) > 0.009:
        diferenca = round(operacao - liquido, 2)
        r["iss"], r["_iss_float"] = _fmt(diferenca), diferenca
        r["obs_iss"] = ("Layout nao mapeado: retencao ESTIMADA pela diferenca entre "
                         "Valor da Operacao e Valor Liquido - CONFERIR MANUALMENTE.")
    else:
        r["iss"], r["_iss_float"] = "", None
        r["obs_iss"] = "" if liquido is None or operacao is None else ""
    r["obs_geral"] = ("Layout de NFS-e nao reconhecido pelo motor (nem SP legado, nem "
                       "DANFSe nacional). Valores acima sao apenas um melhor esforco; "
                       "confira a nota manualmente.")
    return r


# ══════════════════════════════════════════════════════════════════════════
#  ORQUESTRACAO
# ══════════════════════════════════════════════════════════════════════════

def _calcular_totais(r: dict) -> None:
    """Fecha o resultado: total retido, lista de tributos retidos,
    flag tem_retencao e valor liquido (usa o declarado na nota quando
    existe, senao calcula valor_total - total_retido). Compartilhado
    por todos os parsers (PDF legado/DANFSe/generico e XML nacional)."""
    total_ret = sum(r.get(f"_{c}_float") or 0 for c in CAMPOS_RETENCAO)
    achados = [c.upper() for c in CAMPOS_RETENCAO if (r.get(f"_{c}_float") or 0) > 0]
    r["tem_retencao"] = bool(achados)
    r["retencoes_encontradas"] = ", ".join(achados)
    r["total_retencoes_calculado"] = _fmt(total_ret) if total_ret else ""
    r["_total_retencoes_float"] = total_ret if total_ret else None

    liquido_declarado = r.get("_valor_liquido_declarado_float")
    valor_total = r.get("_valor_total_float")
    if liquido_declarado is not None:
        r["valor_liquido"] = _fmt(liquido_declarado)
        r["_valor_liquido_float"] = liquido_declarado
    elif valor_total is not None:
        vl = round(valor_total - total_ret, 2)
        r["valor_liquido"] = _fmt(vl)
        r["_valor_liquido_float"] = vl
    else:
        r["valor_liquido"] = ""
        r["_valor_liquido_float"] = None

    r.setdefault("obs_geral", "")
    r.setdefault("cnpj_tomador", "")
    r.setdefault("discriminacao", "")
    r.setdefault("erro", "")
    # Padrao: nota valida. Vira "CANCELADA" em aplicar_cancelamentos,
    # quando o evento de cancelamento aparecer no mesmo lote.
    r.setdefault("chave_acesso", "")
    r.setdefault("cancelada", False)
    r.setdefault("situacao", "Normal")
    r.setdefault("obs_cancelamento", "")
    r.setdefault("tipo_registro", "nota")


def _analisar_bloco(bloco: dict) -> dict:
    texto = bloco["texto"]
    layout = _detectar_layout(texto)
    if layout == "legado":
        r = _analisar_legado(texto)
    elif layout == "danfse":
        r = _analisar_danfse(texto, bloco)
    else:
        r = _analisar_generico(texto)

    _calcular_totais(r)

    if bloco.get("ocr"):
        aviso_ocr = ("Nota lida por OCR (PDF escaneado, sem camada de texto) - "
                      "valores podem conter erro de leitura, CONFERIR MANUALMENTE.")
        r["obs_geral"] = (aviso_ocr if not r["obs_geral"] else f"{aviso_ocr} | {r['obs_geral']}")
    return r


def _resultado_vazio(arquivo: str, erro: str) -> dict:
    r = {"arquivo": arquivo, "paginas": 0, "layout": "", "erro": erro, "tem_retencao": False}
    for campo in ("numero_nota", "data_emissao", "prestador", "cnpj_prestador", "tomador",
                  "cnpj_tomador", "valor_total", "pis", "cofins", "csll", "inss", "irrf", "iss",
                  "total_retencoes_calculado", "valor_liquido", "retencoes_encontradas",
                  "obs_iss", "obs_geral", "discriminacao"):
        r[campo] = ""
    return r


# ══════════════════════════════════════════════════════════════════════════
#  NFS-e PADRAO NACIONAL — arquivo XML (leiaute nacional versao 1.00)
# ══════════════════════════════════════════════════════════════════════════
#  O XML nacional e o formato mais limpo de todos: cada retencao tem um
#  campo proprio e explicito, sem precisar de regex em texto livre nem
#  de regra por data. Os campos usados:
#    infNFSe/valores/vISSQN, vTotalRet, vLiq          (totais oficiais)
#    infDPS/valores/vServPrest/vServ                  (valor bruto)
#    infDPS/valores/trib/tribMun/tpRetISSQN           1=nao retido,
#                                                     2=retido pelo tomador,
#                                                     3=retido pelo intermediario
#    infDPS/valores/trib/tribFed/piscofins/tpRetPisCofins   0=nao, 1=sim
#    infDPS/valores/trib/tribFed/piscofins/vPis, vCofins
#    infDPS/valores/trib/tribFed/vRetCP     (INSS retido)
#    infDPS/valores/trib/tribFed/vRetIRRF   (IRRF retido)
#    infDPS/valores/trib/tribFed/vRetCSLL   (CSLL retido)
# ══════════════════════════════════════════════════════════════════════════

_NS_NFSE = "{http://www.sped.fazenda.gov.br/nfse}"


def _xt(el, caminho: str) -> str:
    """Texto de um sub-elemento (caminho ja com o namespace via _p), ''
    se nao existir."""
    if el is None:
        return ""
    achado = el.find(caminho)
    return (achado.text or "").strip() if achado is not None and achado.text else ""


def _p(*tags) -> str:
    """Monta um caminho ElementTree com o namespace do NFS-e nacional em
    cada segmento: _p('DPS', 'infDPS') -> '{ns}DPS/{ns}infDPS'."""
    return "/".join(_NS_NFSE + t for t in tags)


def _num_xml(s: str):
    if not s:
        return None
    try:
        return float(str(s).strip().replace(",", "."))
    except ValueError:
        return None


def _fmt_cnpj(c: str) -> str:
    c = re.sub(r"\D", "", c or "")
    if len(c) == 14:
        return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}"
    if len(c) == 11:
        return f"{c[:3]}.{c[3:6]}.{c[6:9]}-{c[9:]}"
    return c


def _iso_para_br(s: str) -> str:
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", s or "")
    return f"{m.group(3)}/{m.group(2)}/{m.group(1)}" if m else ""


def _extrair_nfse_nacional(inf) -> dict:
    r = {"layout": "Padrao Nacional (XML)"}
    dps = inf.find(_p("DPS", "infDPS"))

    # A chave de acesso e o que liga a nota ao evento de cancelamento
    # (que vem em arquivo separado). O atributo Id e "NFS" + a chave.
    id_attr = inf.get("Id", "") or ""
    r["chave_acesso"] = id_attr[3:] if id_attr.startswith("NFS") else id_attr

    r["numero_nota"] = _xt(inf, _p("nNFSe"))
    r["data_emissao"] = _iso_para_br(_xt(dps, _p("dhEmi")) or _xt(inf, _p("dhProc")))

    emit = inf.find(_p("emit"))
    r["prestador"] = _xt(emit, _p("xNome"))
    r["cnpj_prestador"] = _fmt_cnpj(_xt(emit, _p("CNPJ")) or _xt(emit, _p("CPF")))

    toma = dps.find(_p("toma")) if dps is not None else None
    r["tomador"] = _xt(toma, _p("xNome"))
    r["cnpj_tomador"] = _fmt_cnpj(_xt(toma, _p("CNPJ")) or _xt(toma, _p("CPF")))

    r["discriminacao"] = re.sub(
        r"\s+", " ", _xt(dps, _p("serv", "cServ", "xDescServ"))).strip()

    val = inf.find(_p("valores"))            # totais oficiais (infNFSe)
    v_issqn = _num_xml(_xt(val, _p("vISSQN")))
    v_total_ret = _num_xml(_xt(val, _p("vTotalRet")))
    v_liq = _num_xml(_xt(val, _p("vLiq")))

    v_serv = _num_xml(_xt(dps, _p("valores", "vServPrest", "vServ")))
    if v_serv is None:
        v_serv = _num_xml(_xt(val, _p("vBC")))
    r["valor_total"] = _fmt(v_serv)
    r["_valor_total_float"] = v_serv
    r["_valor_liquido_declarado_float"] = v_liq

    trib = dps.find(_p("valores", "trib")) if dps is not None else None

    # ISS
    tp_ret_iss = _xt(trib, _p("tribMun", "tpRetISSQN"))
    if tp_ret_iss in ("2", "3"):
        r["iss"], r["_iss_float"] = _fmt(v_issqn), v_issqn
        quem = "tomador" if tp_ret_iss == "2" else "intermediario"
        r["obs_iss"] = f"ISS retido pelo {quem} (tpRetISSQN={tp_ret_iss})"
    else:
        r["iss"], r["_iss_float"] = "", None
        r["obs_iss"] = ("ISS recolhido pelo prestador (tpRetISSQN=1, nao retido)"
                        if tp_ret_iss == "1" else "")

    # PIS / COFINS
    pc = trib.find(_p("tribFed", "piscofins")) if trib is not None else None
    if _xt(pc, _p("tpRetPisCofins")) == "1":
        pis_v = _num_xml(_xt(pc, _p("vPis")))
        cofins_v = _num_xml(_xt(pc, _p("vCofins")))
        r["pis"], r["_pis_float"] = _fmt(pis_v), pis_v
        r["cofins"], r["_cofins_float"] = _fmt(cofins_v), cofins_v
    else:
        r["pis"], r["_pis_float"] = "", None
        r["cofins"], r["_cofins_float"] = "", None

    # INSS / IRRF / CSLL
    tf = trib.find(_p("tribFed")) if trib is not None else None
    inss_v = _num_xml(_xt(tf, _p("vRetCP")))
    irrf_v = _num_xml(_xt(tf, _p("vRetIRRF")))
    csll_v = _num_xml(_xt(tf, _p("vRetCSLL")))
    r["inss"], r["_inss_float"] = _fmt(inss_v), inss_v
    r["irrf"], r["_irrf_float"] = _fmt(irrf_v), irrf_v
    r["csll"], r["_csll_float"] = _fmt(csll_v), csll_v

    # cruzamento: total retido declarado x soma dos campos lidos
    obs = []
    soma = sum(x or 0 for x in (r["_pis_float"], r["_cofins_float"], r["_csll_float"],
                                r["_inss_float"], r["_irrf_float"], r["_iss_float"]))
    if v_total_ret is not None and abs(v_total_ret - soma) > 0.01:
        obs.append(f"Total retido declarado na NFS-e ({_fmt(v_total_ret)}) difere da soma "
                    f"dos campos de retencao lidos ({_fmt(round(soma, 2))}) - conferir.")
    r["obs_geral"] = " | ".join(obs)
    return r


def _extrair_evento_cancelamento(raiz, nome_arquivo: str) -> dict | None:
    """Os cancelamentos de NFS-e nao vem dentro da nota: sao arquivos
    <evento> separados, que apontam a chave da nota cancelada
    (chNFSe) e o motivo. Devolve um registro de EVENTO (nao e uma
    nota), ou None se o XML nao for um evento de cancelamento."""
    ped = raiz.find(f".//{_NS_NFSE}chNFSe")
    if ped is None or not (ped.text or "").strip():
        return None
    descricao = _xt(raiz, f".//{_NS_NFSE}xDesc")
    if descricao and "cancel" not in descricao.lower():
        return None  # evento de outro tipo (ex.: carta de correcao)
    return {
        "tipo_registro": "evento_cancelamento",
        "arquivo": nome_arquivo,
        "paginas": "XML",
        "layout": "Evento de cancelamento (XML)",
        "chave_acesso": ped.text.strip(),
        "descricao_evento": descricao or "Cancelamento de NFS-e",
        "motivo_cancelamento": _xt(raiz, f".//{_NS_NFSE}xMotivo"),
        "data_evento": _iso_para_br(_xt(raiz, f".//{_NS_NFSE}dhEvento")
                                    or _xt(raiz, f".//{_NS_NFSE}dhProc")),
        "tem_retencao": False,
        "erro": "",
    }


def _analisar_xml_nacional(caminho: Path) -> list[dict]:
    try:
        raiz = ET.parse(caminho).getroot()
    except ET.ParseError as e:
        return [_resultado_vazio(caminho.name, f"XML invalido: {e}")]

    infs = raiz.findall(f".//{_NS_NFSE}infNFSe")
    if not infs and raiz.tag == f"{_NS_NFSE}infNFSe":
        infs = [raiz]
    if not infs:
        evento = _extrair_evento_cancelamento(raiz, caminho.name)
        if evento is not None:
            return [evento]
        return [_resultado_vazio(
            caminho.name,
            "XML sem elemento <infNFSe> - nao parece uma NFS-e padrao nacional.")]

    resultados = []
    for i, inf in enumerate(infs, 1):
        try:
            r = _extrair_nfse_nacional(inf)
            _calcular_totais(r)
        except Exception as e:  # nota malformada nao derruba o lote
            r = _resultado_vazio(caminho.name, f"Falha ao ler a NFS-e do XML: {e}")
        r["arquivo"] = caminho.name
        r["paginas"] = "XML" if len(infs) == 1 else f"XML {i}/{len(infs)}"
        resultados.append(r)
    return resultados


def formatar_situacao_iss(r: dict, modo: str = "tomados") -> str:
    """Texto da coluna 'Situacao do ISS' ajustado ao ponto de vista do
    relatorio: em Servicos Tomados fala do ISS que o TOMADOR (nosso
    cliente) retem do prestador; em Servicos Prestados fala do ISS que o
    TOMADOR do servico (cliente externo) retem do nosso cliente
    (prestador)."""
    retido = bool(r.get("_iss_float"))
    if modo == "prestados":
        if retido:
            return f"ISS devido pelo tomador (retido nesta nota) - {r.get('iss', '')}"
        return "ISS devido pelo prestador (nao retido pelo tomador)"
    return r.get("obs_iss", "") or ("ISS retido pelo tomador" if retido else "")


def aplicar_cancelamentos(registros: list[dict]) -> list[dict]:
    """Cruza os EVENTOS de cancelamento com as NOTAS do mesmo lote e
    marca as notas canceladas. Devolve so as notas (os eventos saem da
    lista, ja que nao sao notas fiscais).

    E preciso analisar o lote INTEIRO antes de chamar esta funcao: o
    cancelamento vem num arquivo separado da nota, entao uma nota so
    pode ser marcada como cancelada depois que o arquivo de evento
    correspondente tambem tiver sido lido.

    A nota cancelada CONTINUA na lista, com todos os valores e
    retencoes que ela tinha - ela so passa a vir marcada
    (cancelada=True e situacao="CANCELADA"), para o contador ver que
    existe e saber que nao deve entrar na apuracao."""
    cancelamentos = {}
    for reg in registros:
        if reg.get("tipo_registro") == "evento_cancelamento":
            chave = reg.get("chave_acesso")
            if chave:
                cancelamentos[chave] = reg

    notas = []
    for reg in registros:
        if reg.get("tipo_registro") == "evento_cancelamento":
            continue
        evento = cancelamentos.get(reg.get("chave_acesso") or "\x00")
        if evento is not None:
            reg["cancelada"] = True
            motivo = evento.get("motivo_cancelamento") or evento.get("descricao_evento") or ""
            data = evento.get("data_evento") or ""
            reg["situacao"] = "CANCELADA"
            reg["obs_cancelamento"] = (
                f"NOTA CANCELADA em {data}. Motivo: {motivo}." if data
                else f"NOTA CANCELADA. Motivo: {motivo}.")
            reg["obs_geral"] = (reg["obs_cancelamento"] if not reg.get("obs_geral")
                                else f"{reg['obs_cancelamento']} | {reg['obs_geral']}")
        notas.append(reg)

    # Eventos cujo XML da nota nao veio junto: avisa, senao o
    # cancelamento passaria despercebido.
    chaves_lidas = {n.get("chave_acesso") for n in notas}
    for chave, evento in cancelamentos.items():
        if chave in chaves_lidas:
            continue
        orfao = _resultado_vazio(
            evento["arquivo"],
            "Evento de cancelamento sem o XML da nota correspondente na pasta - "
            "a nota cancelada pode estar sendo contada em outro relatorio.")
        orfao.update({"chave_acesso": chave, "situacao": "CANCELADA (nota ausente)",
                      "cancelada": True, "layout": evento["layout"],
                      "obs_cancelamento": f"Motivo: {evento.get('motivo_cancelamento', '')}"})
        notas.append(orfao)
    return notas


def analisar_varios(caminhos) -> list[dict]:
    """Analisa VARIOS arquivos de uma vez e ja devolve as notas com os
    cancelamentos aplicados (ver aplicar_cancelamentos). E a forma
    recomendada de processar uma pasta inteira - analisar_arquivo, um
    a um, nao tem como saber que uma nota foi cancelada, porque o
    cancelamento esta em outro arquivo."""
    registros = []
    for caminho in caminhos:
        registros.extend(analisar_arquivo(caminho))
    return aplicar_cancelamentos(registros)


def analisar_arquivo(caminho) -> list[dict]:
    """Ponto de entrada unico: aceita PDF (NFS-e legado, DANFSe ou
    relatorio/talao com varias notas) e XML (NFS-e padrao nacional).
    Devolve sempre uma lista de dicts, um por nota fiscal encontrada."""
    caminho = Path(caminho)
    if caminho.suffix.lower() == ".xml":
        return _analisar_xml_nacional(caminho)
    return _analisar_pdf(caminho)


def _analisar_pdf(caminho: Path) -> list[dict]:
    """Analisa um PDF (que pode conter 1 ou varias notas) e devolve uma
    lista de dicts, um por nota fiscal encontrada."""
    caminho = Path(caminho)
    try:
        blocos = extrair_blocos(caminho)
    except Exception as e:
        return [_resultado_vazio(caminho.name, str(e))]

    houve_pagina_sem_texto_e_sem_ocr = False
    resultados = []
    for bloco in blocos:
        texto = bloco["texto"]
        if not texto.strip():
            if not OCR_DISPONIVEL:
                houve_pagina_sem_texto_e_sem_ocr = True
            continue
        r = _analisar_bloco(bloco)
        if r["layout"].startswith("Nao identificado") and not r["numero_nota"] and not r["_valor_total_float"]:
            continue  # pagina de capa/rosto sem dado nenhum de nota (ex.: capa do Talao Fiscal)
        r["arquivo"] = caminho.name
        pags = bloco["paginas"]
        r["paginas"] = str(pags[0]) if len(pags) == 1 else f"{pags[0]}-{pags[-1]}"
        resultados.append(r)

    if not resultados:
        if houve_pagina_sem_texto_e_sem_ocr:
            erro = ("PDF escaneado (sem camada de texto) e o OCR (Tesseract) nao foi "
                     "encontrado nesta maquina. Instale o Tesseract OCR para processar "
                     "este arquivo, ou peca a NFS-e em formato com texto.")
        else:
            erro = "Sem texto extraivel (PDF vazio ou corrompido?)"
        resultados.append(_resultado_vazio(caminho.name, erro))
    return resultados


# Nome antigo mantido por compatibilidade (o dispatcher e analisar_arquivo)
analisar_pdf = analisar_arquivo
