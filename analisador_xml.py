# -*- coding: utf-8 -*-
"""
====================================================================
 ANALISADOR DE NF-e  --  Programa Unico
 HEC ASSESSORIA CONTABIL S/S LTDA.
====================================================================
Le arquivos XML de NF-e, NFC-e (Cupom Fiscal), CT-e e NFS-e Nacional
dentro de uma ou mais pastas (varre subpastas tambem, inclusive XMLs
dentro de arquivos .zip, sem precisar extrair manualmente), identifica
automaticamente qual desses tipos cada um e (a NFC-e e diferenciada da
NF-e pela tag <mod>65</mod>), e extrai os campos principais para uma
planilha Excel (.xlsx) ou CSV -- cada tipo de documento numa aba
separada -- ja formatada com as cores da HEC.

Campos extraidos da NF-e (a NFC-e/Cupom Fiscal usa os mesmos campos,
so vai para a aba "NFC-e (Cupom Fiscal)" em vez de "NFe"):
    - chNFe   -> Chave de Acesso (44 digitos)
    - nNF     -> Numero da Nota
    - dhEmi   -> Data de Emissao (formatada dd/mm/aaaa + ISO original)
    - CNPJ    -> CNPJ do Emitente
    - xNome   -> Nome/Razao Social do Emitente
    - vNF     -> Valor Total da Nota Fiscal
    - CFOP    -> CFOP(s) dos itens da nota (um ou mais, separados por " / ")
    - dest/xNome -> Nome do Comprador/Destinatario
    - dest/CNPJ ou dest/CPF -> CNPJ/CPF do Comprador/Destinatario
    - infCpl  -> Informacoes Complementares (observacoes da nota)
    - cobr/fat/nFat -> Numero da Fatura/Duplicata (quando a nota tem)
    - cobr/dup      -> Parcelas (numero, vencimento e valor) -- vao para
                       uma aba separada "Fatura e Duplicatas", uma linha
                       por parcela, só das notas que tem esse campo
    - infCpl (texto) -> orcamento, sinal, saldo, a vista -- extracao
                       heuristica, aba separada "Condicoes de Pagamento"

Campos extraidos do CT-e (aba "CT-e"): chave de acesso, numero, data
de emissao, CFOP, transportadora/remetente/destinatario, municipios e
UF de inicio/fim, valor total da prestacao, valor a receber,
vencimento(s) e observacoes.

Campos extraidos da NFS-e Nacional (aba "NFS-e"): numero, data,
competencia, local de emissao, prestador, tomador, descricao do
servico, base de calculo/aliquota/valor do ISS, valor liquido e
informacoes complementares.

INSTALACAO (se for rodar o .py direto, sem o executavel)
----------------------------------------------------------
    pip install pandas openpyxl customtkinter pillow
====================================================================
"""

import os
import re
import sys
import threading
import xml.etree.ElementTree as ET
import zipfile

import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinter as tk

try:
    from PIL import Image
    PIL_OK = True
except ImportError:
    PIL_OK = False


def _caminho_recurso(nome_arquivo: str) -> str:
    """Resolve o caminho de um arquivo (ex.: logo.png) tanto rodando o
    .py direto quanto empacotado pelo PyInstaller (onedir/onefile)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, nome_arquivo)


# ====================================================================
# IDENTIDADE VISUAL -- mesma paleta usada nos demais programas HEC
# (Integrador de Extratos, Analisador de Tributos, Analisador de
# Retencoes, Conciliador), extraida da logo: globo verde-azulado. As
# cores de MARCA (verde/cinza) sao fixas nos dois temas -- so
# fundo/texto/borda mudam entre claro/escuro. A troca de tema
# acontece AO VIVO, sem reiniciar (ver App._aplicar_tema_ao_vivo).
# ====================================================================
VERDE_HEC = "#00926E"
VERDE_ESCURO = "#00674E"
VERDE_CLARO = "#E3F3EE"
CINZA_HEC = "#9A9A9C"
CINZA_ESCURO = "#5A5A5C"
COR_ERRO = "#C00000"
COR_ALERTA = "#E08A00"

_PALETA_ESCURA = {
    "COR_FUNDO": "#101214",
    "COR_FUNDO_MENU": "#181B1E",
    "COR_CARD": "#1C1F22",
    "COR_CARD_ATIVO": "#153029",
    "COR_BORDA": "#2B2F33",
    "COR_TEXTO": "#FFFFFF",
    "COR_MUTED": "#9AA0A6",
}
_PALETA_CLARA = {
    "COR_FUNDO": "#F4F5F6",
    "COR_FUNDO_MENU": "#E9EBED",
    "COR_CARD": "#FFFFFF",
    "COR_CARD_ATIVO": VERDE_CLARO,
    "COR_BORDA": "#D3D7DA",
    "COR_TEXTO": "#181A1B",
    "COR_MUTED": "#54585C",
}

FONTE = "Calibri"
NOME_ESCRITORIO = "HEC ASSESSORIA CONTABIL S/S LTDA."
VERSAO_PROGRAMA = "v01.8"  # atualize a cada nova versao gerada


def _caminho_preferencia_tema():
    """Fica ao lado do .exe (ou do .py) -- sobrevive a atualizacoes de
    versao porque nao fica dentro da pasta que o PyInstaller recria."""
    base = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "tema_preferido.txt")


def carregar_tema_preferido():
    try:
        with open(_caminho_preferencia_tema(), "r", encoding="utf-8-sig") as arquivo:
            return "claro" if arquivo.read().strip().lower() == "claro" else "escuro"
    except Exception:
        return "escuro"


def salvar_tema_preferido(tema):
    try:
        with open(_caminho_preferencia_tema(), "w", encoding="utf-8") as arquivo:
            arquivo.write(tema)
    except Exception:
        pass


TEMA_ATUAL = carregar_tema_preferido()
_paleta = _PALETA_CLARA if TEMA_ATUAL == "claro" else _PALETA_ESCURA

COR_FUNDO = _paleta["COR_FUNDO"]
COR_FUNDO_MENU = _paleta["COR_FUNDO_MENU"]
COR_CARD = _paleta["COR_CARD"]
COR_CARD_ATIVO = _paleta["COR_CARD_ATIVO"]
COR_BORDA = _paleta["COR_BORDA"]
COR_TEXTO = _paleta["COR_TEXTO"]
COR_MUTED = _paleta["COR_MUTED"]


# ------------------------------------------------------------------
# NAMESPACE DA NF-e
# ------------------------------------------------------------------
# Todo XML de NF-e usa esse "xmlns" (namespace). Sem informar isso
# para o ElementTree, as buscas por tag (ex: "nNF") nao encontram
# nada, porque internamente a tag "real" e algo como
# "{http://www.portalfiscal.inf.br/nfe}nNF".
NS = {"nfe": "http://www.portalfiscal.inf.br/nfe"}

# Os outros dois tipos de documento que o programa tambem reconhece --
# CT-e (Conhecimento de Transporte Eletronico) e NFS-e Nacional (Nota
# Fiscal de Servico padrao nacional) -- cada um com o seu proprio
# namespace XML, igual a NF-e.
NS_CTE = {"cte": "http://www.portalfiscal.inf.br/cte"}
NS_NFSE = {"nfse": "http://www.sped.fazenda.gov.br/nfse"}

# NFS-e MUNICIPAL no padrao ABRASF -- layout usado pelas prefeituras que
# ainda tem sistema proprio de nota de servico (inclusive a Nota Fiscal
# Paulistana, de Sao Paulo), diferente da NFS-e Nacional acima. O mesmo
# namespace serve para varias cidades; o municipio vem dentro do XML.
NS_ABRASF = {"a": "http://www.abrasf.org.br/nfse.xsd"}

# Codigo IBGE de Sao Paulo -- usado para identificar a Nota Fiscal
# Paulistana dentro das NFS-e ABRASF.
CODIGO_MUNICIPIO_SAO_PAULO = "3550308"


def extrair_texto(elemento, caminho, namespaces=NS):
    """
    Funcao auxiliar para buscar um campo dentro do XML e devolver o
    texto dele, ja tratando os casos em que o campo nao existe.

    Se o campo nao for encontrado, devolve uma string vazia "" em
    vez de quebrar o script - isso e importante porque nem toda
    nota tem, por exemplo, informacoes complementares (infCpl).
    """
    encontrado = elemento.find(caminho, namespaces)
    if encontrado is not None and encontrado.text is not None:
        return encontrado.text.strip()
    return ""


def extrair_chave_acesso(root):
    """
    A chave de acesso (44 digitos) pode aparecer em dois lugares,
    dependendo de como o XML foi baixado/salvo:

    1) Dentro de <protNFe><infProt><chNFe> - quando o arquivo e o
       "XML completo" (nfeProc), que inclui o protocolo de
       autorizacao da SEFAZ.

    2) Caso o arquivo seja "so a NFe" (sem o protocolo), a chave fica
       embutida no atributo "Id" da tag <infNFe>, no formato "NFe" +
       44 digitos. Nesse caso, extraimos a chave removendo o prefixo
       "NFe".
    """
    chave = extrair_texto(root, ".//nfe:protNFe/nfe:infProt/nfe:chNFe")
    if chave:
        return chave

    inf_nfe = root.find(".//nfe:infNFe", NS)
    if inf_nfe is not None:
        id_attr = inf_nfe.get("Id", "")  # ex: "NFe3526060161..."
        if id_attr.startswith("NFe"):
            return id_attr[3:]
        return id_attr

    return ""


def extrair_data_emissao_formatada(data_emissao_iso):
    """
    Recebe a data/hora de emissao no formato ISO da NF-e
    (ex: "2026-08-14T10:30:00-03:00") e devolve apenas a data no
    formato brasileiro "dd/mm/aaaa". Se o formato vier diferente do
    esperado (ou vazio), devolve o valor original sem alteracoes.
    """
    if not data_emissao_iso:
        return ""
    try:
        ano, mes, dia = data_emissao_iso[:10].split("-")
        return f"{dia}/{mes}/{ano}"
    except (ValueError, IndexError):
        return data_emissao_iso


def extrair_cfops(inf_nfe):
    """
    Uma nota fiscal pode ter varios itens (tag <det>), e cada item
    tem o seu proprio CFOP dentro de <det><prod><CFOP>. Essa funcao
    coleta todos os CFOPs encontrados, remove repetidos (mantendo a
    ordem) e devolve como uma unica string, separada por " / " quando
    houver mais de um.
    """
    cfops_encontrados = []
    for cfop_elemento in inf_nfe.findall(".//nfe:det/nfe:prod/nfe:CFOP", NS):
        if cfop_elemento.text:
            cfop = cfop_elemento.text.strip()
            if cfop and cfop not in cfops_encontrados:
                cfops_encontrados.append(cfop)
    return " / ".join(cfops_encontrados)


def extrair_descricao_produtos(inf_nfe):
    """
    Cada item da nota (tag <det>) tem a descricao do produto/servico
    em <det><prod><xProd>. Essa funcao junta a descricao de TODOS os
    itens numa unica string, separadas por " / " -- assim a planilha
    mostra, numa coluna so, tudo o que foi vendido naquela nota.

    Diferente do CFOP (ver extrair_cfops), aqui as repeticoes NAO sao
    removidas: se a nota tem o mesmo produto em dois itens diferentes
    (ex.: mesma bebida em dois lancamentos), os dois aparecem, porque
    representam linhas reais da nota.
    """
    descricoes = []
    for produto in inf_nfe.findall(".//nfe:det/nfe:prod/nfe:xProd", NS):
        if produto.text and produto.text.strip():
            descricoes.append(produto.text.strip())
    return " / ".join(descricoes)


def extrair_itens_nfe(inf_nfe, nome_arquivo, chave_acesso, numero_nota):
    """
    Devolve uma lista de dicts, UMA LINHA POR ITEM da nota (cada tag
    <det>), com codigo, descricao, NCM, CFOP, unidade, quantidade e
    valores do item. Alimenta a aba "Itens da Nota" da planilha.

    Ter uma linha por item (em vez de tudo espremido numa celula so)
    e o que permite, no Excel, filtrar por produto, somar por NCM ou
    por CFOP e conferir item a item -- mesma ideia da aba "Fatura e
    Duplicatas", que tem uma linha por parcela.
    """
    linhas_itens = []
    for indice, det in enumerate(inf_nfe.findall("nfe:det", NS), start=1):
        # O numero do item vem no atributo nItem de <det>; se vier
        # vazio (nota fora do padrao), cai na contagem sequencial.
        numero_item = det.get("nItem") or str(indice)
        linhas_itens.append({
            "Arquivo de Origem": nome_arquivo,
            "Chave de Acesso (chNFe)": chave_acesso,
            "Numero da Nota (nNF)": numero_nota,
            "Item": numero_item,
            "Codigo do Produto (cProd)": extrair_texto(det, "nfe:prod/nfe:cProd"),
            "Descricao (xProd)": extrair_texto(det, "nfe:prod/nfe:xProd"),
            "NCM": extrair_texto(det, "nfe:prod/nfe:NCM"),
            "CFOP": extrair_texto(det, "nfe:prod/nfe:CFOP"),
            "Unidade (uCom)": extrair_texto(det, "nfe:prod/nfe:uCom"),
            "Quantidade (qCom)": extrair_texto(det, "nfe:prod/nfe:qCom"),
            "Valor Unitario (vUnCom)": extrair_texto(det, "nfe:prod/nfe:vUnCom"),
            "Valor Total do Item (vProd)": extrair_texto(det, "nfe:prod/nfe:vProd"),
        })
    return linhas_itens


def extrair_cnpj_cpf_destinatario(inf_nfe):
    """
    O destinatario (comprador) pode ser identificado por CNPJ (pessoa
    juridica) ou CPF (pessoa fisica) -- a NF-e so preenche um dos dois
    campos, nunca os dois. Essa funcao tenta o CNPJ primeiro e, se nao
    encontrar, cai para o CPF.
    """
    cnpj = extrair_texto(inf_nfe, "nfe:dest/nfe:CNPJ")
    if cnpj:
        return cnpj
    return extrair_texto(inf_nfe, "nfe:dest/nfe:CPF")


def extrair_fatura_e_duplicatas(inf_nfe):
    """
    O bloco <cobr> (cobranca) e OPCIONAL na NF-e -- só aparece quando a
    venda foi feita a prazo/faturada. Quando existe, traz:

        <fat>  -> os dados da FATURA/DUPLICATA como um todo (numero)
        <dup>  -> uma ou mais parcelas, cada uma com numero (nDup),
                  data de vencimento (dVenc) e valor (vDup)

    Devolve uma tupla (numero_fatura, lista_parcelas):
        numero_fatura  -> string (nFat), ou "" se a nota nao tem cobr/fat
        lista_parcelas -> lista de dicts {"nDup", "dVenc", "vDup"},
                           vazia se a nota nao tem nenhuma parcela
    """
    numero_fatura = extrair_texto(inf_nfe, "nfe:cobr/nfe:fat/nfe:nFat")

    lista_parcelas = []
    for dup_elemento in inf_nfe.findall("nfe:cobr/nfe:dup", NS):
        lista_parcelas.append({
            "nDup": extrair_texto(dup_elemento, "nfe:nDup"),
            "dVenc": extrair_texto(dup_elemento, "nfe:dVenc"),
            "vDup": extrair_texto(dup_elemento, "nfe:vDup"),
        })

    return numero_fatura, lista_parcelas


def _texto_valor_br_para_float(texto_valor):
    """Converte um valor no formato BR ("25.000,00") para float. Devolve
    None se o texto nao for um numero valido."""
    texto_limpo = texto_valor.replace(".", "").replace(",", ".")
    try:
        return float(texto_limpo)
    except ValueError:
        return None


# Formato de valor monetario BR dentro de texto livre: grupos de 3
# digitos separados por ponto (milhar, opcional) + virgula + 2 casas
# decimais -- ex.: "25.000,00" ou "199,90". Usar so "[\d.,]+" (sem essa
# exigencia de 2 casas decimais no final) pega de brinde a pontuacao
# da FRASE (ex.: a virgula de "R$25.000,00, pago via..."), quebrando a
# conversao para numero.
PADRAO_VALOR_BR = r"(\d{1,3}(?:\.\d{3})*,\d{2})"


def extrair_condicoes_pagamento_do_texto(texto_infcpl, valor_total_nota=None):
    """
    A NF-e NAO tem um campo estruturado para "sinal/entrada",
    "orcamento", "a vista" ou "condicao de pagamento por extenso" --
    quando o emissor descreve isso, e sempre como TEXTO LIVRE dentro
    de infCpl (Informacoes Complementares / Dados Adicionais). Por
    ser texto livre, cada empresa/ERP escreve de um jeito diferente --
    esta funcao e uma extracao HEURISTICA (por padroes de texto,
    regex) que funciona para redacoes parecidas com o exemplo:

        "VENDA EFETUADA CONFORME ORCAMENTO 573/2025G. FORMA DE
        PAGAMENTO: SINAL NO VALOR DE R$25.000,00, PAGO VIA
        PIX/TRANSF E SALDO NO VALOR DE R$10.225,56 EM 02 BOLETOS
        BANCARIOS COM VENCIMENTO 30/60 DIAS APOS DATA DE EMISSAO..."

    ou, para venda a vista:

        "PAGAMENTO A VISTA" / "VENDA A VISTA NO VALOR DE R$1.500,00"

    Quando o texto tiver uma redacao muito diferente, os campos que
    nao forem reconhecidos ficam de fora do dicionario devolvido (nao
    "adivinha" nem inventa valor) -- EXCECAO: quando a palavra "A
    VISTA" aparece mas nenhum valor e informado junto, assume-se o
    Valor Total da Nota (vNF) como o valor a vista, ja que e o unico
    valor que faz sentido nesse caso. Por isso o texto original tambem
    e mantido na planilha, para conferencia manual de qualquer nota.

    'valor_total_nota' -> float com o vNF da nota (ou None), usado
    apenas como fallback do "Valor a Vista" quando o texto nao traz
    um valor explicito.

    Devolve um dicionario (vazio se nada foi reconhecido) com as
    chaves que forem encontradas dentre:
        "Numero do Orcamento", "Valor do Sinal/Entrada",
        "Valor Residual/Saldo", "Quantidade de Parcelas (texto)",
        "Vencimento/Prazo (texto)", "Valor a Vista"
    """
    if not texto_infcpl:
        return {}

    texto = texto_infcpl.upper()
    resultado = {}

    combinacao = re.search(r"OR[CÇ]AMENTO\s*N?[ºO°:]*\s*([A-Z0-9./-]+)", texto)
    if combinacao:
        resultado["Numero do Orcamento"] = combinacao.group(1).rstrip(".,")

    combinacao = re.search(r"SINAL(?:\s+NO\s+VALOR\s+DE)?\s*R\$\s*" + PADRAO_VALOR_BR, texto)
    if combinacao:
        resultado["Valor do Sinal/Entrada"] = _texto_valor_br_para_float(combinacao.group(1))

    combinacao = re.search(r"(?:SALDO|VALOR\s+RESIDUAL)(?:\s+NO\s+VALOR\s+DE)?\s*R\$\s*" + PADRAO_VALOR_BR, texto)
    if combinacao:
        resultado["Valor Residual/Saldo"] = _texto_valor_br_para_float(combinacao.group(1))

    combinacao = re.search(r"(\d+)\s*BOLET", texto) or re.search(r"EM\s+(\d+)\s*(?:PARCELAS|VEZES)", texto)
    if combinacao:
        resultado["Quantidade de Parcelas (texto)"] = int(combinacao.group(1))

    combinacao = re.search(r"VENCIMENTO\S*\s+(.+?)(?:\.|$)", texto)
    if combinacao:
        resultado["Vencimento/Prazo (texto)"] = combinacao.group(1).strip().rstrip(".")

    # "A VISTA" ("A" ou "À") -- procura um valor logo depois da
    # expressao (ex.: "A VISTA NO VALOR DE R$1.500,00"); se a palavra
    # aparecer mas sem valor nenhum junto, cai no valor total da nota.
    if re.search(r"[AÀ]\s*VISTA", texto):
        combinacao = re.search(r"[AÀ]\s*VISTA(?:\s+NO\s+VALOR\s+DE)?\s*R\$\s*" + PADRAO_VALOR_BR, texto)
        if combinacao:
            resultado["Valor a Vista"] = _texto_valor_br_para_float(combinacao.group(1))
        else:
            resultado["Valor a Vista"] = valor_total_nota

    return resultado


def _caminho_longo_windows(caminho):
    """
    O Windows, por padrao, nao abre arquivos cujo caminho completo passe
    de 260 caracteres -- e as pastas de XML dos clientes chegam nisso
    facil (ex.: ...\\RelatorioMensal-2026-05_CNPJ_...\\CTe\\CTE\\NF 85884 -
    BRINKS SEGURANCA E TRANSPORTES ... .xml tem 261). O prefixo "\\\\?\\"
    (ou "\\\\?\\UNC\\" para pastas de rede \\\\servidor\\...) desliga esse
    limite. Fora do Windows, devolve o caminho sem mudanca.
    """
    if os.name != "nt":
        return caminho
    caminho_absoluto = os.path.abspath(caminho)
    if caminho_absoluto.startswith("\\\\?\\"):
        return caminho_absoluto
    if caminho_absoluto.startswith("\\\\"):
        return "\\\\?\\UNC\\" + caminho_absoluto[2:]
    return "\\\\?\\" + caminho_absoluto


def _ler_xml_da_fonte(fonte, callback_log):
    """
    Le o XML de uma 'fonte' (ver listar_fontes_xml) e devolve uma
    tupla (root, nome_arquivo):

        root         -> a raiz (Element) do XML ja parseado, ou None
                         se nao foi possivel ler (erro ja avisado via
                         callback_log)
        nome_arquivo -> nome para exibicao/planilha (ex.: "3016.xml",
                         ou "3016.xml (dentro de notas.zip)")

    'fonte' aceita dois formatos:
        - uma string: caminho de um arquivo .xml solto no disco
        - uma tupla ("zip", caminho_do_zip, nome_interno): um .xml que
          esta DENTRO de um arquivo .zip, sem precisar extrair primeiro

    Erro de leitura (XML corrompido, arquivo sumiu, sem permissao etc.)
    vira so um aviso no log -- nunca derruba o processamento do lote.
    """
    if isinstance(fonte, tuple):
        _tipo, caminho_zip, nome_interno = fonte
        nome_exibicao = f"{nome_interno} (dentro de {os.path.basename(caminho_zip)})"
        try:
            with zipfile.ZipFile(_caminho_longo_windows(caminho_zip)) as arquivo_zip:
                conteudo = arquivo_zip.read(nome_interno)
            root = ET.fromstring(conteudo)
        except (ET.ParseError, zipfile.BadZipFile, KeyError, OSError) as erro:
            callback_log(f"[AVISO] Nao foi possivel ler o XML '{nome_exibicao}': {erro}")
            return None, nome_exibicao
        return root, nome_exibicao

    try:
        arvore = ET.parse(_caminho_longo_windows(fonte))
        root = arvore.getroot()
    except (ET.ParseError, OSError) as erro:
        callback_log(f"[AVISO] Nao foi possivel ler o XML '{os.path.basename(fonte)}': {erro}")
        return None, os.path.basename(fonte)
    return root, os.path.basename(fonte)


def identificar_tipo_documento(root):
    """
    Identifica qual dos tipos de documento reconhecidos pelo programa
    o XML representa, procurando a tag "principal" de cada um em
    qualquer lugar da arvore (funciona tanto para o XML "puro" quanto
    para a versao com o protocolo de autorizacao anexado, ex.:
    nfeProc/cteProc). Devolve "nfe", "nfce", "cte", "nfse", ou None se
    nao for nenhum desses.

    NF-e e NFC-e (Nota Fiscal de Consumidor Eletronica -- o "cupom
    fiscal" eletronico que substituiu o antigo cupom do ECF) usam
    exatamente o mesmo layout/namespace de XML -- a UNICA diferenca
    entre os dois e a tag <ide><mod>: "55" e NF-e (venda para outra
    empresa), "65" e NFC-e (venda direta ao consumidor final, o
    "cupom"). Por isso, sem checar o <mod>, os dois ficariam
    misturados na mesma aba.
    """
    inf_nfe = root.find(".//nfe:infNFe", NS)
    if inf_nfe is not None:
        modelo = extrair_texto(inf_nfe, "nfe:ide/nfe:mod")
        return "nfce" if modelo == "65" else "nfe"
    if root.find(".//cte:infCte", NS_CTE) is not None:
        return "cte"
    if root.find(".//nfse:infNFSe", NS_NFSE) is not None:
        return "nfse"
    if root.find(".//a:InfNfse", NS_ABRASF) is not None:
        return "nfse_abrasf"
    return None


def _processar_nfe(root, nome_arquivo):
    """
    Extrai os campos de uma NF-e ja identificada. Devolve uma tupla
    (dados, linhas_parcelas, linha_condicao_pagamento, linhas_itens) --
    ver processar_arquivo_xml. 'dados' vem None se a tag infNFe nao for
    encontrada (nao deveria acontecer, ja que identificar_tipo_documento
    ja confirmou -- e so uma segunda checagem defensiva).
    """
    inf_nfe = root.find(".//nfe:infNFe", NS)
    if inf_nfe is None:
        return None, [], None, []

    dhEmi = extrair_texto(inf_nfe, "nfe:ide/nfe:dhEmi")
    chave_acesso = extrair_chave_acesso(root)
    numero_nota = extrair_texto(inf_nfe, "nfe:ide/nfe:nNF")
    texto_infcpl = extrair_texto(inf_nfe, "nfe:infAdic/nfe:infCpl")
    texto_valor_total_nota = extrair_texto(inf_nfe, "nfe:total/nfe:ICMSTot/nfe:vNF")
    try:
        valor_total_nota = float(texto_valor_total_nota) if texto_valor_total_nota else None
    except ValueError:
        valor_total_nota = None

    numero_fatura, lista_parcelas = extrair_fatura_e_duplicatas(inf_nfe)

    dados = {
        "Chave de Acesso (chNFe)": chave_acesso,
        "Numero da Nota (nNF)": numero_nota,
        "Data de Emissao": extrair_data_emissao_formatada(dhEmi),
        "Data/Hora de Emissao (dhEmi)": dhEmi,
        "CNPJ Emitente": extrair_texto(inf_nfe, "nfe:emit/nfe:CNPJ"),
        "Nome Emitente (xNome)": extrair_texto(inf_nfe, "nfe:emit/nfe:xNome"),
        "Nome do Comprador (dest/xNome)": extrair_texto(inf_nfe, "nfe:dest/nfe:xNome"),
        "CNPJ/CPF do Comprador": extrair_cnpj_cpf_destinatario(inf_nfe),
        "Descricao dos Produtos": extrair_descricao_produtos(inf_nfe),
        "Valor Total da Nota (vNF)": texto_valor_total_nota,
        "CFOP": extrair_cfops(inf_nfe),
        # infCpl fica dentro de <infAdic>, por isso o caminho tem os dois niveis:
        "Informacoes Complementares (infCpl)": texto_infcpl,
        "Numero da Fatura (nFat)": numero_fatura,
        "Quantidade de Parcelas": len(lista_parcelas),
        "Arquivo de Origem": nome_arquivo,
    }

    linhas_parcelas = []
    for indice, parcela in enumerate(lista_parcelas, start=1):
        linhas_parcelas.append({
            "Arquivo de Origem": nome_arquivo,
            "Chave de Acesso (chNFe)": chave_acesso,
            "Numero da Nota (nNF)": numero_nota,
            "Numero da Fatura (nFat)": numero_fatura,
            "Parcela": indice,
            "Numero da Parcela (nDup)": parcela["nDup"],
            "Vencimento": extrair_data_emissao_formatada(parcela["dVenc"]),
            "Valor da Parcela (vDup)": parcela["vDup"],
        })

    condicoes_encontradas = extrair_condicoes_pagamento_do_texto(texto_infcpl, valor_total_nota)
    linha_condicao_pagamento = None
    if condicoes_encontradas:
        linha_condicao_pagamento = {
            "Arquivo de Origem": nome_arquivo,
            "Chave de Acesso (chNFe)": chave_acesso,
            "Numero da Nota (nNF)": numero_nota,
            "Numero do Orcamento": condicoes_encontradas.get("Numero do Orcamento", ""),
            "Valor do Sinal/Entrada": condicoes_encontradas.get("Valor do Sinal/Entrada"),
            "Valor Residual/Saldo": condicoes_encontradas.get("Valor Residual/Saldo"),
            "Valor a Vista": condicoes_encontradas.get("Valor a Vista"),
            "Quantidade de Parcelas (texto)": condicoes_encontradas.get("Quantidade de Parcelas (texto)"),
            "Vencimento/Prazo (texto)": condicoes_encontradas.get("Vencimento/Prazo (texto)", ""),
            "Texto Original (infCpl)": texto_infcpl,
        }

    linhas_itens = extrair_itens_nfe(inf_nfe, nome_arquivo, chave_acesso, numero_nota)

    return dados, linhas_parcelas, linha_condicao_pagamento, linhas_itens


def _processar_cte(root, nome_arquivo):
    """
    Extrai os campos de um CT-e (Conhecimento de Transporte
    Eletronico) ja identificado. Devolve o dicionario de dados, ou
    None se a tag infCte nao for encontrada.
    """
    inf_cte = root.find(".//cte:infCte", NS_CTE)
    if inf_cte is None:
        return None

    def txt(caminho):
        return extrair_texto(inf_cte, caminho, NS_CTE)

    dhEmi = txt("cte:ide/cte:dhEmi")

    # A chave de acesso normalmente vem no protocolo de autorizacao
    # (protCTe/infProt/chCTe); se o arquivo nao tiver o protocolo,
    # cai para o atributo "Id" de <infCte> (formato "CTe" + 44 digitos),
    # igual se faz com a NF-e.
    chave_acesso = extrair_texto(root, ".//cte:protCTe/cte:infProt/cte:chCTe", NS_CTE)
    if not chave_acesso:
        id_attr = inf_cte.get("Id", "")
        chave_acesso = id_attr[3:] if id_attr.startswith("CTe") else id_attr

    # Vencimento(s) -- o bloco cobr/dup do CT-e tem a mesma estrutura
    # do da NF-e, mas normalmente e uma parcela unica (o frete e pago
    # de uma vez). Junta todas as datas encontradas, caso existam mais.
    vencimentos = [
        extrair_data_emissao_formatada(dup.findtext("cte:dVenc", default="", namespaces=NS_CTE))
        for dup in inf_cte.findall(".//cte:cobr/cte:dup", NS_CTE)
    ]
    vencimentos = [v for v in vencimentos if v]

    return {
        "Chave de Acesso (chCTe)": chave_acesso,
        "Numero do CT-e (nCT)": txt("cte:ide/cte:nCT"),
        "Data de Emissao": extrair_data_emissao_formatada(dhEmi),
        "Data/Hora de Emissao (dhEmi)": dhEmi,
        "CFOP": txt("cte:ide/cte:CFOP"),
        "CNPJ Transportadora (emit)": txt("cte:emit/cte:CNPJ"),
        "Nome Transportadora (emit/xNome)": txt("cte:emit/cte:xNome"),
        "CNPJ Remetente (rem)": txt("cte:rem/cte:CNPJ"),
        "Nome Remetente (rem/xNome)": txt("cte:rem/cte:xNome"),
        "CNPJ Destinatario (dest)": txt("cte:dest/cte:CNPJ"),
        "Nome Destinatario (dest/xNome)": txt("cte:dest/cte:xNome"),
        "Municipio de Inicio": txt("cte:ide/cte:xMunIni"),
        "UF de Inicio": txt("cte:ide/cte:UFIni"),
        "Municipio de Fim": txt("cte:ide/cte:xMunFim"),
        "UF de Fim": txt("cte:ide/cte:UFFim"),
        "Valor Total da Prestacao (vTPrest)": txt("cte:vPrest/cte:vTPrest"),
        "Valor a Receber (vRec)": txt("cte:vPrest/cte:vRec"),
        "Vencimento(s)": " / ".join(vencimentos),
        "Observacoes (xObs)": txt("cte:compl/cte:xObs"),
        "Arquivo de Origem": nome_arquivo,
    }


def _processar_nfse(root, nome_arquivo):
    """
    Extrai os campos de uma NFS-e Nacional (Nota Fiscal de Servico no
    padrao nacional, layout do Ambiente de Dados Nacional/ADN) ja
    identificada. Devolve o dicionario de dados, ou None se a tag
    infNFSe nao for encontrada.
    """
    inf_nfse = root.find(".//nfse:infNFSe", NS_NFSE)
    if inf_nfse is None:
        return None

    def txt(caminho):
        return extrair_texto(inf_nfse, caminho, NS_NFSE)

    dhProc = txt("nfse:dhProc")
    id_attr = inf_nfse.get("Id", "")

    return {
        "Chave/Id da NFS-e": id_attr,
        "Numero da NFS-e (nNFSe)": txt("nfse:nNFSe"),
        "Data de Emissao": extrair_data_emissao_formatada(dhProc),
        "Data/Hora de Processamento (dhProc)": dhProc,
        "Competencia (dCompet)": txt("nfse:DPS/nfse:infDPS/nfse:dCompet"),
        "Local de Emissao": txt("nfse:xLocEmi"),
        "CNPJ Prestador (emit)": txt("nfse:emit/nfse:CNPJ"),
        "Nome Prestador (emit/xNome)": txt("nfse:emit/nfse:xNome"),
        "CNPJ/CPF Tomador (toma)": (
            txt("nfse:DPS/nfse:infDPS/nfse:toma/nfse:CNPJ")
            or txt("nfse:DPS/nfse:infDPS/nfse:toma/nfse:CPF")
        ),
        "Nome Tomador (toma/xNome)": txt("nfse:DPS/nfse:infDPS/nfse:toma/nfse:xNome"),
        "Descricao do Servico": txt("nfse:DPS/nfse:infDPS/nfse:serv/nfse:cServ/nfse:xDescServ"),
        "Valor do Servico (vServ)": txt("nfse:DPS/nfse:infDPS/nfse:valores/nfse:vServPrest/nfse:vServ"),
        "Base de Calculo ISS (vBC)": txt("nfse:valores/nfse:vBC"),
        "Aliquota ISS % (pAliqAplic)": txt("nfse:valores/nfse:pAliqAplic"),
        "Valor do ISS (vISSQN)": txt("nfse:valores/nfse:vISSQN"),
        "Valor Liquido (vLiq)": txt("nfse:valores/nfse:vLiq"),
        "Informacoes Complementares": txt("nfse:DPS/nfse:infDPS/nfse:serv/nfse:infoCompl/nfse:xInfComp"),
        "Arquivo de Origem": nome_arquivo,
    }


def _processar_nfse_abrasf(root, nome_arquivo):
    """
    Extrai os campos de uma NFS-e MUNICIPAL no padrao ABRASF (ex.:
    Nota Fiscal Paulistana, e notas de servico de Campinas, Guarulhos,
    Barueri, Rio de Janeiro etc.). Devolve o dicionario de dados, ou
    None se a tag InfNfse nao for encontrada.

    Cada campo e buscado em mais de um caminho possivel, porque o
    ABRASF tem duas geracoes de layout: na 2.x os dados do servico e do
    tomador ficam dentro de DeclaracaoPrestacaoServico, na 1.x ficam
    direto em InfNfse (Servico / TomadorServico).
    """
    inf = root.find(".//a:InfNfse", NS_ABRASF)
    if inf is None:
        return None

    def primeiro(*caminhos):
        for caminho in caminhos:
            valor = extrair_texto(inf, caminho, NS_ABRASF)
            if valor:
                return valor
        return ""

    data_emissao = primeiro("a:DataEmissao")

    codigo_municipio = primeiro(
        "a:OrgaoGerador/a:CodigoMunicipio",
        ".//a:Servico/a:MunicipioIncidencia",
        ".//a:Servico/a:CodigoMunicipio",
        "a:PrestadorServico/a:Endereco/a:CodigoMunicipio",
    )
    if codigo_municipio == CODIGO_MUNICIPIO_SAO_PAULO:
        municipio = "SAO PAULO (Nota Fiscal Paulistana)"
    else:
        municipio = codigo_municipio

    # A discriminacao costuma vir com varias quebras de linha (formatada
    # para impressao); junta tudo numa linha so pra caber na celula.
    discriminacao = " ".join(primeiro(".//a:Servico/a:Discriminacao").split())

    iss_retido = primeiro(".//a:Servico/a:IssRetido")
    iss_retido = {"1": "Sim", "2": "Nao"}.get(iss_retido, iss_retido)

    return {
        "Numero da NFS-e": primeiro("a:Numero"),
        "Codigo de Verificacao": primeiro("a:CodigoVerificacao"),
        "Data de Emissao": extrair_data_emissao_formatada(data_emissao),
        "Data/Hora de Emissao": data_emissao,
        "Competencia": primeiro(".//a:InfDeclaracaoPrestacaoServico/a:Competencia", "a:Competencia"),
        "Municipio": municipio,
        "CNPJ/CPF Prestador": primeiro(
            "a:PrestadorServico/a:IdentificacaoPrestador/a:CpfCnpj/a:Cnpj",
            "a:PrestadorServico/a:IdentificacaoPrestador/a:CpfCnpj/a:Cpf",
            "a:PrestadorServico/a:IdentificacaoPrestador/a:Cnpj",
            ".//a:Prestador/a:CpfCnpj/a:Cnpj",
            ".//a:Prestador/a:CpfCnpj/a:Cpf",
        ),
        "Nome Prestador": primeiro("a:PrestadorServico/a:RazaoSocial"),
        "CNPJ/CPF Tomador": primeiro(
            ".//a:Tomador/a:IdentificacaoTomador/a:CpfCnpj/a:Cnpj",
            ".//a:Tomador/a:IdentificacaoTomador/a:CpfCnpj/a:Cpf",
            ".//a:TomadorServico/a:IdentificacaoTomador/a:CpfCnpj/a:Cnpj",
            ".//a:TomadorServico/a:IdentificacaoTomador/a:CpfCnpj/a:Cpf",
        ),
        "Nome Tomador": primeiro(".//a:Tomador/a:RazaoSocial", ".//a:TomadorServico/a:RazaoSocial"),
        "Item da Lista de Servico": primeiro(".//a:Servico/a:ItemListaServico"),
        "Descricao do Servico (Discriminacao)": discriminacao,
        "Valor dos Servicos": primeiro(".//a:Servico/a:Valores/a:ValorServicos"),
        "Base de Calculo ISS": primeiro("a:ValoresNfse/a:BaseCalculo", ".//a:Servico/a:Valores/a:BaseCalculo"),
        "Aliquota ISS": primeiro("a:ValoresNfse/a:Aliquota", ".//a:Servico/a:Valores/a:Aliquota"),
        "Valor do ISS": primeiro("a:ValoresNfse/a:ValorIss", ".//a:Servico/a:Valores/a:ValorIss"),
        "ISS Retido": iss_retido,
        # Retencoes federais -- so vem preenchidas quando o tomador retem
        # o tributo; em branco = nao houve retencao informada na nota.
        "PIS Retido": primeiro(".//a:Servico/a:Valores/a:ValorPis"),
        "COFINS Retido": primeiro(".//a:Servico/a:Valores/a:ValorCofins"),
        "CSLL Retido": primeiro(".//a:Servico/a:Valores/a:ValorCsll"),
        "IR Retido": primeiro(".//a:Servico/a:Valores/a:ValorIr"),
        "INSS Retido": primeiro(".//a:Servico/a:Valores/a:ValorInss"),
        # Muitas prefeituras (inclusive SP, em boa parte das notas) NAO
        # informam o valor liquido -- nesse caso a celula fica em branco
        # (nao e calculado aqui, para nao mostrar um valor que nao esta
        # na nota).
        "Valor Liquido": primeiro("a:ValoresNfse/a:ValorLiquidoNfse", ".//a:Servico/a:Valores/a:ValorLiquidoNfse"),
        "Arquivo de Origem": nome_arquivo,
    }


def _chave_unica_documento(tipo, dados):
    """
    Identificador unico de um documento fiscal, para detectar o MESMO
    documento lido mais de uma vez (ex.: o XML solto na pasta e a
    copia dele dentro de um .zip, ou pastas copiadas). Devolve None se
    o documento nao tiver chave (nesse caso nao da pra afirmar que e
    repetido, entao ele entra normalmente).
    """
    if tipo in ("nfe", "nfce"):
        chave = dados.get("Chave de Acesso (chNFe)")
    elif tipo == "cte":
        chave = dados.get("Chave de Acesso (chCTe)")
    elif tipo == "nfse":
        chave = dados.get("Chave/Id da NFS-e")
    elif tipo == "nfse_abrasf":
        partes = [dados.get("CNPJ/CPF Prestador"), dados.get("Numero da NFS-e"), dados.get("Codigo de Verificacao")]
        chave = "|".join(partes) if all(partes) else None
    else:
        chave = None
    return f"{tipo}:{chave}" if chave else None


def processar_arquivo_xml(fonte, callback_log=print):
    """
    Le um XML e devolve uma tupla
    (tipo_documento, dados, linhas_parcelas, linha_condicao_pagamento):

        tipo_documento            -> "nfe", "nfce" (Cupom Fiscal), "cte",
                                      "nfse", ou None se o arquivo nao
                                      foi reconhecido/nao pode ser lido
                                      (motivo ja avisado via callback_log)
        dados                     -> dicionario com os campos do
                                      documento (o formato depende do
                                      tipo -- ver _processar_nfe/_cte/_nfse)
        linhas_parcelas           -> lista de dicts com as parcelas da
                                      fatura/duplicata -- so preenchida
                                      para NF-e/NFC-e, vazia para os demais
        linha_condicao_pagamento  -> dict com a condicao de pagamento
                                      reconhecida no texto -- so
                                      preenchida para NF-e/NFC-e, None
                                      para os demais
        linhas_itens              -> lista de dicts, uma por item/produto
                                      da nota -- so preenchida para
                                      NF-e/NFC-e, vazia para os demais

    Quando 'tipo_documento' vem None, os demais valores vem None/[]/None/[].

    'fonte' aceita dois formatos (ver listar_fontes_xml):
        - uma string: caminho de um arquivo .xml solto no disco
        - uma tupla ("zip", caminho_do_zip, nome_interno): um .xml que
          esta DENTRO de um arquivo .zip, sem precisar extrair primeiro
    """
    root, nome_arquivo = _ler_xml_da_fonte(fonte, callback_log)
    if root is None:
        return None, None, [], None, []

    tipo = identificar_tipo_documento(root)

    if tipo in ("nfe", "nfce"):
        # NFC-e (Cupom Fiscal) usa exatamente a mesma extracao da NF-e
        # -- so muda a aba pra onde vai (ver identificar_tipo_documento).
        dados, linhas_parcelas, linha_condicao, linhas_itens = _processar_nfe(root, nome_arquivo)
        if dados is None:
            nome_tipo = "NFC-e" if tipo == "nfce" else "NF-e"
            callback_log(f"[AVISO] Arquivo '{nome_arquivo}' nao parece ser uma {nome_tipo} valida (tag infNFe nao encontrada).")
            return None, None, [], None, []
        return tipo, dados, linhas_parcelas, linha_condicao, linhas_itens

    if tipo == "cte":
        dados = _processar_cte(root, nome_arquivo)
        if dados is None:
            callback_log(f"[AVISO] Arquivo '{nome_arquivo}' nao parece ser um CT-e valido (tag infCte nao encontrada).")
            return None, None, [], None, []
        return "cte", dados, [], None, []

    if tipo == "nfse":
        dados = _processar_nfse(root, nome_arquivo)
        if dados is None:
            callback_log(f"[AVISO] Arquivo '{nome_arquivo}' nao parece ser uma NFS-e valida (tag infNFSe nao encontrada).")
            return None, None, [], None, []
        return "nfse", dados, [], None, []

    if tipo == "nfse_abrasf":
        dados = _processar_nfse_abrasf(root, nome_arquivo)
        if dados is None:
            callback_log(f"[AVISO] Arquivo '{nome_arquivo}' nao parece ser uma NFS-e municipal valida (tag InfNfse nao encontrada).")
            return None, None, [], None, []
        return "nfse_abrasf", dados, [], None, []

    callback_log(f"[AVISO] Arquivo '{nome_arquivo}' nao e um XML de NF-e, NFC-e, CT-e ou NFS-e reconhecido.")
    return None, None, [], None, []


def listar_fontes_xml(pasta):
    """
    Varre a pasta informada (incluindo subpastas) e devolve a lista de
    "fontes" de XML encontradas -- tanto arquivos .xml soltos quanto
    XMLs que estao DENTRO de arquivos .zip (sem precisar extrair
    manualmente antes).

    Cada item da lista e:
        - uma string, para um .xml solto (o caminho completo dele)
        - uma tupla ("zip", caminho_do_zip, nome_interno), para um .xml
          que esta dentro de um .zip (nome_interno e o caminho dele
          DENTRO do zip, ex.: "notas/3016.xml")
    """
    fontes_encontradas = []
    # A varredura parte do caminho ja com o prefixo de caminho longo (ver
    # _caminho_longo_windows) -- sem isso, o os.walk PULA EM SILENCIO as
    # subpastas que passam de 260 caracteres, e os XMLs delas nem
    # chegariam a ser lidos.
    for pasta_atual, _subpastas, arquivos in os.walk(_caminho_longo_windows(pasta)):
        for nome_arquivo in arquivos:
            caminho_completo = os.path.join(pasta_atual, nome_arquivo)
            if nome_arquivo.lower().endswith(".xml"):
                fontes_encontradas.append(caminho_completo)
            elif nome_arquivo.lower().endswith(".zip"):
                try:
                    with zipfile.ZipFile(caminho_completo) as arquivo_zip:
                        for nome_interno in arquivo_zip.namelist():
                            if nome_interno.lower().endswith(".xml"):
                                fontes_encontradas.append(("zip", caminho_completo, nome_interno))
                except (zipfile.BadZipFile, OSError):
                    pass  # zip corrompido/invalido/ilegivel -- ignora, sem travar o resto do processamento
    return fontes_encontradas


# Verde da logo da HEC Assessoria Contabil -- mesma cor de marca usada
# na interface do programa (ver VERDE_HEC mais abaixo), aplicada aqui
# no cabecalho das planilhas geradas para manter a identidade visual.
VERDE_HEC_EXCEL = "00926E"


def formatar_planilha_hec(planilha, df):
    """
    Aplica a formatacao padrao HEC numa aba ja escrita pelo
    pandas/openpyxl: cabecalho verde (cor da logo) com texto branco em
    negrito, linha de cabecalho congelada (fica visivel ao rolar) e
    largura de coluna ajustada ao conteudo.
    """
    preenchimento_cabecalho = PatternFill(start_color=VERDE_HEC_EXCEL, end_color=VERDE_HEC_EXCEL, fill_type="solid")
    fonte_cabecalho = Font(color="000000", bold=True)
    alinhamento_cabecalho = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for indice_coluna, nome_coluna in enumerate(df.columns, start=1):
        celula = planilha.cell(row=1, column=indice_coluna)
        celula.fill = preenchimento_cabecalho
        celula.font = fonte_cabecalho
        celula.alignment = alinhamento_cabecalho

        maior_texto = max([len(str(nome_coluna))] + [len(str(valor)) for valor in df.iloc[:, indice_coluna - 1]])
        planilha.column_dimensions[get_column_letter(indice_coluna)].width = min(max(maior_texto + 2, 12), 45)

    planilha.freeze_panes = "A2"
    planilha.row_dimensions[1].height = 28


def _forcar_colunas_como_texto(df, colunas_texto):
    """Converte as colunas informadas (se existirem no df) para texto --
    evita que o Excel "coma" zeros a esquerda de codigos/chaves."""
    for coluna in colunas_texto:
        if coluna in df.columns:
            df[coluna] = df[coluna].astype(str)


def _escrever_aba_excel(writer, df, nome_aba, colunas_texto):
    """
    Escreve um DataFrame como uma aba do Excel ja com a formatacao HEC
    (ver formatar_planilha_hec) e com as 'colunas_texto' formatadas em
    celula-texto ("@"), para o Excel nao reinterpretar codigos/chaves
    como numero.
    """
    df.to_excel(writer, index=False, sheet_name=nome_aba)
    planilha = writer.sheets[nome_aba]
    for coluna in colunas_texto:
        if coluna not in df.columns:
            continue
        indice_coluna = df.columns.get_loc(coluna) + 1
        letra_coluna = planilha.cell(row=1, column=indice_coluna).column_letter
        for linha_num in range(2, len(df) + 2):
            celula = planilha[f"{letra_coluna}{linha_num}"]
            celula.number_format = "@"
    formatar_planilha_hec(planilha, df)


def gerar_planilha(pastas_origem, arquivo_saida, callback_log=print):
    """
    Funcao principal: varre a(s) pasta(s), processa cada XML -- reconhece
    NF-e, CT-e e NFS-e Nacional, cada um indo para a sua propria aba --
    e salva o resultado consolidado em Excel (.xlsx) ou CSV, dependendo
    da extensao informada em 'arquivo_saida'. Cada etapa e reportada
    via 'callback_log' (por padrao, print no console).

    'pastas_origem' aceita uma unica pasta (string) ou uma lista de
    pastas -- quando e uma lista, os XMLs de TODAS elas (e suas
    subpastas) entram no mesmo resultado consolidado.

    Devolve uma tupla (sucesso, mensagem):
        sucesso  -> True/False
        mensagem -> texto explicando o resultado final
    """
    if isinstance(pastas_origem, str):
        pastas_origem = [pastas_origem]

    fontes_xml = []
    for pasta in pastas_origem:
        callback_log(f"Procurando arquivos XML em: {pasta}")
        fontes_xml.extend(listar_fontes_xml(pasta))

    if not fontes_xml:
        mensagem = "Nenhum arquivo .xml (solto ou dentro de .zip) foi encontrado na(s) pasta(s) de origem selecionada(s)."
        callback_log(mensagem)
        return False, mensagem

    callback_log(f"{len(fontes_xml)} arquivo(s) XML encontrado(s). Processando...")

    linhas_nfe = []
    linhas_nfce = []
    linhas_cte = []
    linhas_nfse = []
    linhas_nfse_abrasf = []
    linhas_parcelas = []
    linhas_condicoes_pagamento = []
    linhas_itens = []
    chaves_ja_lidas = set()
    quantidade_repetidos = 0
    for fonte in fontes_xml:
        nome_para_log = f"{fonte[2]} (dentro de {os.path.basename(fonte[1])})" if isinstance(fonte, tuple) else os.path.basename(fonte)
        callback_log(f"Lendo: {nome_para_log}")
        tipo, dados, parcelas_da_nota, condicao_pagamento, itens_da_nota = processar_arquivo_xml(fonte, callback_log=callback_log)

        # Mesmo documento lido de novo (outra copia do XML) -- ignora
        # inteiro (inclusive itens/parcelas), senao os valores somam em
        # dobro na planilha.
        chave_unica = _chave_unica_documento(tipo, dados) if tipo else None
        if chave_unica is not None:
            if chave_unica in chaves_ja_lidas:
                quantidade_repetidos += 1
                continue
            chaves_ja_lidas.add(chave_unica)

        if tipo == "nfe":
            linhas_nfe.append(dados)
            linhas_parcelas.extend(parcelas_da_nota)
            linhas_itens.extend(itens_da_nota)
            if condicao_pagamento is not None:
                linhas_condicoes_pagamento.append(condicao_pagamento)
        elif tipo == "nfce":
            linhas_nfce.append(dados)
            linhas_parcelas.extend(parcelas_da_nota)
            linhas_itens.extend(itens_da_nota)
            if condicao_pagamento is not None:
                linhas_condicoes_pagamento.append(condicao_pagamento)
        elif tipo == "cte":
            linhas_cte.append(dados)
        elif tipo == "nfse":
            linhas_nfse.append(dados)
        elif tipo == "nfse_abrasf":
            linhas_nfse_abrasf.append(dados)

    if not (linhas_nfe or linhas_nfce or linhas_cte or linhas_nfse or linhas_nfse_abrasf):
        mensagem = "Nenhum documento valido (NF-e, NFC-e, CT-e ou NFS-e) foi extraido dos XMLs encontrados. Nada foi salvo."
        callback_log(mensagem)
        return False, mensagem

    # ---- NF-e e NFC-e (Cupom Fiscal) ------------------------------------
    # Usam exatamente o mesmo layout de campos (ver identificar_tipo_documento
    # / _processar_nfe) -- so vao para abas diferentes.
    df_nfe = None
    df_nfce = None
    df_parcelas = None
    df_condicoes = None
    df_itens = None
    colunas_texto_nfe = ["Chave de Acesso (chNFe)", "CNPJ Emitente", "Numero da Nota (nNF)", "CFOP",
                          "CNPJ/CPF do Comprador", "Numero da Fatura (nFat)"]
    colunas_texto_parcelas = ["Chave de Acesso (chNFe)", "Numero da Nota (nNF)",
                               "Numero da Fatura (nFat)", "Numero da Parcela (nDup)"]
    colunas_texto_condicoes = ["Chave de Acesso (chNFe)", "Numero da Nota (nNF)", "Numero do Orcamento"]
    colunas_texto_itens = ["Chave de Acesso (chNFe)", "Numero da Nota (nNF)", "Item",
                            "Codigo do Produto (cProd)", "NCM", "CFOP"]

    if linhas_nfe:
        df_nfe = pd.DataFrame(linhas_nfe)
        # IMPORTANTE: CNPJ, Chave de Acesso, Numero da Nota e CFOP sao
        # codigos, nao numeros -- ver _forcar_colunas_como_texto.
        _forcar_colunas_como_texto(df_nfe, colunas_texto_nfe)
        if "Valor Total da Nota (vNF)" in df_nfe.columns:
            df_nfe["Valor Total da Nota (vNF)"] = pd.to_numeric(df_nfe["Valor Total da Nota (vNF)"], errors="coerce")

    if linhas_nfce:
        df_nfce = pd.DataFrame(linhas_nfce)
        _forcar_colunas_como_texto(df_nfce, colunas_texto_nfe)
        if "Valor Total da Nota (vNF)" in df_nfce.columns:
            df_nfce["Valor Total da Nota (vNF)"] = pd.to_numeric(df_nfce["Valor Total da Nota (vNF)"], errors="coerce")

    # Aba separada "Fatura e Duplicatas" -- so existe se pelo menos uma
    # nota (NF-e ou NFC-e) do lote tiver o bloco cobr/dup (venda a
    # prazo). Uma linha por PARCELA (nao por nota), para facilitar
    # conferencia de vencimentos/valores no Excel (filtro, soma, etc.).
    if linhas_parcelas:
        df_parcelas = pd.DataFrame(linhas_parcelas)
        _forcar_colunas_como_texto(df_parcelas, colunas_texto_parcelas)
        df_parcelas["Valor da Parcela (vDup)"] = pd.to_numeric(df_parcelas["Valor da Parcela (vDup)"], errors="coerce")

    # Aba separada "Condicoes de Pagamento (Texto)" -- extracao
    # HEURISTICA do texto livre de infCpl (ver
    # extrair_condicoes_pagamento_do_texto). Mantem o texto original
    # ao lado, para conferencia manual.
    if linhas_condicoes_pagamento:
        df_condicoes = pd.DataFrame(linhas_condicoes_pagamento)
        _forcar_colunas_como_texto(df_condicoes, colunas_texto_condicoes)

    # Aba separada "Itens da Nota" -- uma linha por ITEM/produto das
    # NF-e/NFC-e do lote (ver extrair_itens_nfe). Quantidade e valores
    # ficam numericos para permitir soma/filtro direto no Excel.
    if linhas_itens:
        df_itens = pd.DataFrame(linhas_itens)
        _forcar_colunas_como_texto(df_itens, colunas_texto_itens)
        for coluna in ["Quantidade (qCom)", "Valor Unitario (vUnCom)", "Valor Total do Item (vProd)"]:
            df_itens[coluna] = pd.to_numeric(df_itens[coluna], errors="coerce")

    # ---- CT-e ----------------------------------------------------------
    df_cte = None
    colunas_texto_cte = ["Chave de Acesso (chCTe)", "Numero do CT-e (nCT)", "CFOP",
                          "CNPJ Transportadora (emit)", "CNPJ Remetente (rem)", "CNPJ Destinatario (dest)"]
    if linhas_cte:
        df_cte = pd.DataFrame(linhas_cte)
        _forcar_colunas_como_texto(df_cte, colunas_texto_cte)
        for coluna in ["Valor Total da Prestacao (vTPrest)", "Valor a Receber (vRec)"]:
            df_cte[coluna] = pd.to_numeric(df_cte[coluna], errors="coerce")

    # ---- NFS-e Nacional --------------------------------------------------
    df_nfse = None
    colunas_texto_nfse = ["Chave/Id da NFS-e", "Numero da NFS-e (nNFSe)",
                           "CNPJ Prestador (emit)", "CNPJ/CPF Tomador (toma)"]
    if linhas_nfse:
        df_nfse = pd.DataFrame(linhas_nfse)
        _forcar_colunas_como_texto(df_nfse, colunas_texto_nfse)
        for coluna in ["Valor do Servico (vServ)", "Base de Calculo ISS (vBC)", "Aliquota ISS % (pAliqAplic)",
                       "Valor do ISS (vISSQN)", "Valor Liquido (vLiq)"]:
            df_nfse[coluna] = pd.to_numeric(df_nfse[coluna], errors="coerce")

    # ---- NFS-e Municipal (ABRASF / Nota Fiscal Paulistana) ----------------
    df_nfse_abrasf = None
    colunas_texto_nfse_abrasf = ["Numero da NFS-e", "Codigo de Verificacao", "CNPJ/CPF Prestador",
                                  "CNPJ/CPF Tomador", "Item da Lista de Servico", "Municipio"]
    if linhas_nfse_abrasf:
        df_nfse_abrasf = pd.DataFrame(linhas_nfse_abrasf)
        _forcar_colunas_como_texto(df_nfse_abrasf, colunas_texto_nfse_abrasf)
        for coluna in ["Valor dos Servicos", "Base de Calculo ISS", "Aliquota ISS", "Valor do ISS",
                       "PIS Retido", "COFINS Retido", "CSLL Retido", "IR Retido", "INSS Retido", "Valor Liquido"]:
            df_nfse_abrasf[coluna] = pd.to_numeric(df_nfse_abrasf[coluna], errors="coerce")

    extensao = os.path.splitext(arquivo_saida)[1].lower()

    # Cada aba que existir vira um arquivo -- a primeira da lista usa o
    # nome exato escolhido pelo usuario; as demais ganham um sufixo.
    abas_presentes = [
        ("NFe", df_nfe, colunas_texto_nfe, ""),
        ("NFC-e (Cupom Fiscal)", df_nfce, colunas_texto_nfe, "_nfce"),
        ("CT-e", df_cte, colunas_texto_cte, "_cte"),
        ("NFS-e Nacional", df_nfse, colunas_texto_nfse, "_nfse"),
        ("NFS-e Municipal (ABRASF)", df_nfse_abrasf, colunas_texto_nfse_abrasf, "_nfse_municipal"),
        ("Itens da Nota", df_itens, colunas_texto_itens, "_itens"),
        ("Fatura e Duplicatas", df_parcelas, colunas_texto_parcelas, "_fatura_duplicatas"),
        ("Condicoes de Pagamento (Texto)", df_condicoes, colunas_texto_condicoes, "_condicoes_pagamento"),
    ]
    abas_presentes = [aba for aba in abas_presentes if aba[1] is not None]

    if extensao == ".csv":
        # CSV so suporta uma tabela por arquivo -- a primeira aba usa o
        # nome exato escolhido; as demais viram arquivos .csv ao lado.
        for indice, (_nome_aba, df_atual, _colunas, sufixo) in enumerate(abas_presentes):
            caminho = arquivo_saida if indice == 0 else os.path.splitext(arquivo_saida)[0] + sufixo + ".csv"
            df_atual.to_csv(caminho, index=False, sep=";", encoding="utf-8-sig")
    else:
        if extensao != ".xlsx":
            arquivo_saida = os.path.splitext(arquivo_saida)[0] + ".xlsx"

        # Gravamos usando o ExcelWriter para poder aplicar um FORMATO
        # DE CELULA "texto" (@) nas colunas de codigo, garantindo que
        # o Excel nao vai reinterpretar esses valores como numero.
        with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as writer:
            for nome_aba, df_atual, colunas_texto_aba, _sufixo in abas_presentes:
                _escrever_aba_excel(writer, df_atual, nome_aba, colunas_texto_aba)

    partes_resumo = []
    if linhas_nfe:
        partes_resumo.append(f"{len(linhas_nfe)} NF-e")
    if linhas_nfce:
        partes_resumo.append(f"{len(linhas_nfce)} NFC-e (Cupom Fiscal)")
    if linhas_cte:
        partes_resumo.append(f"{len(linhas_cte)} CT-e")
    if linhas_nfse:
        partes_resumo.append(f"{len(linhas_nfse)} NFS-e Nacional")
    if linhas_nfse_abrasf:
        quantidade_paulistana = sum(1 for l in linhas_nfse_abrasf if l["Municipio"].startswith("SAO PAULO"))
        partes_resumo.append(f"{len(linhas_nfse_abrasf)} NFS-e Municipal ({quantidade_paulistana} Paulistana)")

    quantidade_com_parcelas = sum(1 for l in linhas_nfe + linhas_nfce if l["Quantidade de Parcelas"] > 0)
    mensagem = (
        f"Concluido! {', '.join(partes_resumo)} exportado(s) com sucesso"
        + (f", com {len(linhas_itens)} item(ns)/produto(s) detalhado(s) na aba 'Itens da Nota'" if linhas_itens else "")
        + (f", sendo {len(linhas_parcelas)} parcela(s) de fatura/duplicata em "
           f"{quantidade_com_parcelas} nota(s)." if linhas_parcelas else ".")
        + (f" {len(linhas_condicoes_pagamento)} nota(s) com condicao de pagamento reconhecida no texto de observacoes." if linhas_condicoes_pagamento else "")
        + (f" {quantidade_repetidos} arquivo(s) repetido(s) (mesmo documento em mais de um XML/.zip) foram ignorados "
           f"para nao contar em dobro." if quantidade_repetidos else "")
        + f"\nArquivo salvo em: {os.path.abspath(arquivo_saida)}"
    )
    callback_log(mensagem)
    return True, mensagem


NOME_ARQUIVO_SAIDA = "resultado_final_XML.xlsx"


# Cinza classico do Windows 98 -- pedido do usuario para os botoes do
# programa terem o visual "retro" (cinza, em relevo/3D, fonte preta).
CINZA_WIN98 = "#C0C0C0"


def criar_botao_win98(parent, texto, comando):
    """
    Cria um botao no estilo classico do Windows 98: cinza, com relevo
    3D (borda clara/escura simulando luz vindo de cima-esquerda) e
    texto preto. O customtkinter (CTkButton) so desenha botoes "chapados"
    com cantos arredondados -- nao da pra fazer esse efeito de relevo
    com ele, entao aqui usamos o tk.Button classico (nao-temático), que
    e o unico que desenha esse bevel automaticamente via relief="raised".
    """
    return tk.Button(
        parent, text=texto, command=comando,
        bg=CINZA_WIN98, fg="#000000",
        activebackground=CINZA_WIN98, activeforeground="#000000",
        disabledforeground="#808080",
        font=(FONTE, 12, "bold"), relief="raised", bd=3,
        highlightthickness=0, cursor="hand2", padx=14, pady=6,
    )


# ====================================================================
# JANELA PRINCIPAL -- menu lateral fixo (225px) com logo/titulo/tema,
# + area de conteudo roladora a direita, organizada em "cards", igual
# ao padrao usado nos demais programas HEC.
# ====================================================================
class App(ctk.CTk):

    NOME_PROGRAMA_MENU = ("ANALISADOR DE", "NF-e")
    SUBTITULO_MENU = "Extracao de dados de\ndocumentos fiscais"

    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("Light" if TEMA_ATUAL == "claro" else "Dark")
        ctk.set_default_color_theme("green")

        self.title(f"{' '.join(self.NOME_PROGRAMA_MENU)} {VERSAO_PROGRAMA} — HEC Assessoria Contabil")
        self.geometry("1200x800")
        self.minsize(980, 620)
        self.configure(fg_color=COR_FUNDO)

        try:
            self.iconbitmap(_caminho_recurso("logo_hec.ico"))
        except Exception:
            pass

        self.pastas_origem = []
        self.pasta_destino = ""
        self._processando = False

        self._build_menu()
        self._build_area_principal()

    # ---------------------------------------------------------------
    def _build_menu(self):
        m = ctk.CTkFrame(self, width=225, corner_radius=0, fg_color=COR_FUNDO_MENU)
        m.pack(side="left", fill="y")
        m.pack_propagate(False)

        if PIL_OK:
            try:
                im = Image.open(_caminho_recurso("logo.png"))
                ctk.CTkLabel(m, image=ctk.CTkImage(im, im, size=(120, 120)),
                             text="").pack(pady=(28, 10))
            except Exception:
                ctk.CTkLabel(m, text="HEC", font=ctk.CTkFont(FONTE, 42, "bold"),
                             text_color=VERDE_HEC).pack(pady=(40, 10))
        else:
            ctk.CTkLabel(m, text="HEC", font=ctk.CTkFont(FONTE, 42, "bold"),
                         text_color=VERDE_HEC).pack(pady=(40, 10))

        linha1, linha2 = self.NOME_PROGRAMA_MENU
        ctk.CTkLabel(m, text=linha1, font=ctk.CTkFont(FONTE, 18, "bold"),
                     text_color=COR_TEXTO).pack()
        ctk.CTkLabel(m, text=linha2, font=ctk.CTkFont(FONTE, 17, "bold"),
                     text_color=VERDE_HEC, wraplength=190, justify="center").pack(pady=(0, 4))
        ctk.CTkLabel(m, text=self.SUBTITULO_MENU,
                     font=ctk.CTkFont(FONTE, 13), text_color=COR_MUTED,
                     justify="center").pack(pady=(0, 20))

        ctk.CTkLabel(m, text=VERSAO_PROGRAMA, font=ctk.CTkFont(FONTE, 11),
                     text_color=CINZA_HEC).pack(side="bottom", pady=(0, 4))
        ctk.CTkLabel(m, text=NOME_ESCRITORIO, font=ctk.CTkFont(FONTE, 11),
                     text_color=CINZA_HEC, wraplength=190,
                     justify="center").pack(side="bottom", pady=(16, 4))

        self._var_tema_claro = ctk.BooleanVar(value=(TEMA_ATUAL == "claro"))
        ctk.CTkSwitch(
            m, text="Modo Claro", variable=self._var_tema_claro,
            onvalue=True, offvalue=False, command=self._alternar_tema,
            progress_color=VERDE_HEC, button_color="#FFFFFF", button_hover_color="#E8E8E8",
            text_color=COR_MUTED, font=ctk.CTkFont(FONTE, 11),
        ).pack(side="bottom", pady=(0, 10))

    def _alternar_tema(self):
        novo_tema = "claro" if self._var_tema_claro.get() else "escuro"
        salvar_tema_preferido(novo_tema)
        self._aplicar_tema_ao_vivo(novo_tema)

    def _aplicar_tema_ao_vivo(self, novo_tema):
        """Troca o tema SEM reiniciar o processo: reconstroi a tela do
        zero com as cores novas, preservando pastas ja selecionadas e
        o log atual."""
        estado_salvo = {
            "pastas_origem": list(self.pastas_origem),
            "pasta_destino": self.pasta_destino,
            "log": self._texto_log.get("1.0", "end-1c") if hasattr(self, "_texto_log") else "",
        }

        globals()["TEMA_ATUAL"] = novo_tema
        nova_paleta = _PALETA_CLARA if novo_tema == "claro" else _PALETA_ESCURA
        for nome_constante, valor in nova_paleta.items():
            globals()[nome_constante] = valor
        ctk.set_appearance_mode("Light" if novo_tema == "claro" else "Dark")

        for widget in self.winfo_children():
            widget.destroy()
        self.configure(fg_color=COR_FUNDO)
        self._build_menu()
        self._build_area_principal()

        self.pastas_origem = estado_salvo["pastas_origem"]
        self.pasta_destino = estado_salvo["pasta_destino"]
        self._atualizar_labels_pastas()
        if estado_salvo["log"]:
            self._texto_log.configure(state="normal")
            self._texto_log.insert("1.0", estado_salvo["log"])
            self._texto_log.configure(state="disabled")

    # ---------------------------------------------------------------
    def _build_area_principal(self):
        area = ctk.CTkScrollableFrame(self, fg_color=COR_FUNDO)
        area.pack(side="left", fill="both", expand=True, padx=22, pady=20)

        self._card_pasta_origem(area)
        self._card_pasta_destino(area)
        self._card_processar(area)

    def _card_pasta_origem(self, area):
        card = ctk.CTkFrame(area, fg_color=COR_CARD, border_color=COR_BORDA,
                             border_width=1, corner_radius=10)
        card.pack(fill="x", pady=(0, 14))

        ctk.CTkLabel(card, text="1.  Pasta(s) de origem (onde estao os arquivos XML)",
                     font=ctk.CTkFont(FONTE, 16, "bold"),
                     text_color=COR_TEXTO).pack(anchor="w", padx=16, pady=(14, 2))
        ctk.CTkLabel(card, text="Pode adicionar mais de uma pasta -- os documentos de todas elas entram no "
                                 "mesmo resultado. Cada pasta pode conter subpastas -- o programa varre tudo "
                                 "recursivamente, inclusive XMLs que estiverem dentro de arquivos .zip.",
                     font=ctk.CTkFont(FONTE, 13), text_color=COR_MUTED,
                     justify="left", wraplength=760).pack(anchor="w", padx=16, pady=(0, 10))

        linha_botoes = ctk.CTkFrame(card, fg_color="transparent")
        linha_botoes.pack(fill="x", padx=16, pady=(0, 8))

        criar_botao_win98(linha_botoes, "Adicionar Pasta de Origem...",
                          self._selecionar_pasta_origem).pack(side="left", padx=(0, 8))
        criar_botao_win98(linha_botoes, "Limpar Pastas",
                          self._limpar_pastas_origem).pack(side="left")

        self._label_pasta_origem = ctk.CTkLabel(
            card, text="Nenhuma pasta selecionada.",
            font=ctk.CTkFont(FONTE, 13), text_color=COR_MUTED,
            justify="left", wraplength=760, anchor="w",
        )
        self._label_pasta_origem.pack(anchor="w", padx=16, pady=(0, 16))

    def _card_pasta_destino(self, area):
        card = ctk.CTkFrame(area, fg_color=COR_CARD, border_color=COR_BORDA,
                             border_width=1, corner_radius=10)
        card.pack(fill="x", pady=(0, 14))

        ctk.CTkLabel(card, text="2.  Pasta de destino (onde salvar a planilha)",
                     font=ctk.CTkFont(FONTE, 16, "bold"),
                     text_color=COR_TEXTO).pack(anchor="w", padx=16, pady=(14, 2))
        ctk.CTkLabel(card, text=f"O arquivo sera salvo como '{NOME_ARQUIVO_SAIDA}' na pasta escolhida.",
                     font=ctk.CTkFont(FONTE, 13), text_color=COR_MUTED,
                     justify="left", wraplength=760).pack(anchor="w", padx=16, pady=(0, 10))

        linha = ctk.CTkFrame(card, fg_color="transparent")
        linha.pack(fill="x", padx=16, pady=(0, 16))

        criar_botao_win98(linha, "Selecionar Pasta de Destino...",
                          self._selecionar_pasta_destino).pack(side="left", padx=(0, 12))

        self._label_pasta_destino = ctk.CTkLabel(
            linha, text="Nenhuma pasta selecionada.",
            font=ctk.CTkFont(FONTE, 13), text_color=COR_MUTED,
            justify="left", wraplength=560, anchor="w",
        )
        self._label_pasta_destino.pack(side="left", fill="x", expand=True)

    def _card_processar(self, area):
        card = ctk.CTkFrame(area, fg_color=COR_CARD, border_color=COR_BORDA,
                             border_width=1, corner_radius=10)
        card.pack(fill="x", pady=(0, 14))

        ctk.CTkLabel(card, text="3.  Processar e gerar a planilha",
                     font=ctk.CTkFont(FONTE, 16, "bold"),
                     text_color=COR_TEXTO).pack(anchor="w", padx=16, pady=(14, 2))
        ctk.CTkLabel(card, text="Reconhece NF-e, NFC-e (Cupom Fiscal), CT-e, NFS-e Nacional e NFS-e Municipal "
                                 "(padrao ABRASF, incluindo a Nota Fiscal Paulistana), cada um numa aba propria da "
                                 "planilha. Documentos repetidos (o mesmo XML solto e dentro de .zip, por exemplo) "
                                 "sao contados uma vez so. Da NF-e/NFC-e extrai Chave de Acesso, Numero, Data de "
                                 "Emissao, CNPJ/Nome do Emitente, Nome e CNPJ/CPF do Comprador, Descricao dos "
                                 "Produtos, Valor Total, CFOP, Observacoes, Fatura/Duplicatas (quantidade de "
                                 "parcelas, vencimento e valor) e condicoes de pagamento descritas no texto das "
                                 "observacoes (orcamento, sinal, saldo, a vista) -- alem de uma aba 'Itens da Nota' "
                                 "com uma linha por produto (codigo, descricao, NCM, CFOP, quantidade e valores). "
                                 "Do CT-e extrai dados do transporte (transportadora, remetente, "
                                 "destinatario, valores, vencimento). Das NFS-e extrai prestador, tomador, municipio, "
                                 "descricao do servico, valores/ISS e retencoes federais -- inclusive documentos que "
                                 "estiverem dentro de arquivos .zip.",
                     font=ctk.CTkFont(FONTE, 13), text_color=COR_MUTED,
                     justify="left", wraplength=760).pack(anchor="w", padx=16, pady=(0, 10))

        linha_botoes = ctk.CTkFrame(card, fg_color="transparent")
        linha_botoes.pack(fill="x", padx=16, pady=(0, 12))

        self._botao_processar = criar_botao_win98(
            linha_botoes, "Gerar Planilha", self._iniciar_processamento)
        self._botao_processar.pack(side="left", padx=(0, 8))

        criar_botao_win98(linha_botoes, "Limpar Log", self._limpar_log).pack(side="left")

        self._texto_log = ctk.CTkTextbox(
            card, height=260, fg_color=COR_FUNDO, text_color=COR_TEXTO,
            font=ctk.CTkFont("Consolas", 12), border_width=1, border_color=COR_BORDA,
            wrap="word", state="disabled",
        )
        self._texto_log.pack(fill="x", padx=16, pady=(0, 16))

    # ---------------------------------------------------------------
    def _atualizar_labels_pastas(self):
        if self.pastas_origem:
            texto = f"{len(self.pastas_origem)} pasta(s) selecionada(s):\n" + "\n".join(
                f" - {pasta}" for pasta in self.pastas_origem)
        else:
            texto = "Nenhuma pasta selecionada."
        self._label_pasta_origem.configure(text=texto)
        self._label_pasta_destino.configure(
            text=self.pasta_destino if self.pasta_destino else "Nenhuma pasta selecionada.")

    def _selecionar_pasta_origem(self):
        pasta = filedialog.askdirectory(
            parent=self, title="Selecione uma PASTA DE ORIGEM (onde estao os arquivos XML)")
        if pasta and pasta not in self.pastas_origem:
            self.pastas_origem.append(pasta)
            self._atualizar_labels_pastas()

    def _limpar_pastas_origem(self):
        self.pastas_origem = []
        self._atualizar_labels_pastas()

    def _selecionar_pasta_destino(self):
        pasta = filedialog.askdirectory(
            parent=self, title="Selecione a PASTA DE DESTINO (onde a planilha sera salva)")
        if pasta:
            self.pasta_destino = pasta
            self._label_pasta_destino.configure(text=pasta)

    def _limpar_log(self):
        self._texto_log.configure(state="normal")
        self._texto_log.delete("1.0", "end")
        self._texto_log.configure(state="disabled")

    def _logar(self, mensagem):
        self._texto_log.configure(state="normal")
        self._texto_log.insert("end", mensagem + "\n")
        self._texto_log.see("end")
        self._texto_log.configure(state="disabled")

    # ---------------------------------------------------------------
    def _iniciar_processamento(self):
        if self._processando:
            return

        if not self.pastas_origem:
            messagebox.showwarning("Pasta de origem nao selecionada",
                                    "Adicione ao menos uma pasta de origem (onde estao os arquivos XML) antes de continuar.",
                                    parent=self)
            return
        if not self.pasta_destino:
            messagebox.showwarning("Pasta de destino nao selecionada",
                                    "Selecione a pasta de destino (onde a planilha sera salva) antes de continuar.",
                                    parent=self)
            return

        self._limpar_log()
        self._processando = True
        self._botao_processar.configure(state="disabled", text="Processando...")

        caminho_saida = os.path.join(self.pasta_destino, NOME_ARQUIVO_SAIDA)
        # Roda em uma thread separada para nao travar a janela enquanto
        # processa muitos arquivos XML.
        thread = threading.Thread(
            target=self._processar_em_segundo_plano,
            args=(list(self.pastas_origem), caminho_saida),
            daemon=True,
        )
        thread.start()

    def _processar_em_segundo_plano(self, pastas_origem, caminho_saida):
        def log_thread_safe(mensagem):
            self.after(0, self._logar, mensagem)

        try:
            sucesso, mensagem = gerar_planilha(pastas_origem, caminho_saida, callback_log=log_thread_safe)
        except Exception as erro:
            sucesso, mensagem = False, f"Erro inesperado ao processar: {erro}"
            log_thread_safe(mensagem)

        self.after(0, self._finalizar_processamento, sucesso, mensagem)

    def _finalizar_processamento(self, sucesso, mensagem):
        self._processando = False
        self._botao_processar.configure(state="normal", text="Gerar Planilha")
        if sucesso:
            messagebox.showinfo("Concluido", mensagem, parent=self)
        else:
            messagebox.showerror("Nao foi possivel concluir", mensagem, parent=self)


if __name__ == "__main__":
    app = App()
    app.mainloop()
