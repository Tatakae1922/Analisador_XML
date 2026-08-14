# -*- coding: utf-8 -*-
"""
====================================================================
 ANALISADOR DE NF-e  --  Programa Unico
 HEC ASSESSORIA CONTABIL S/S LTDA.
====================================================================
Le todos os arquivos XML de NF-e (Nota Fiscal Eletronica) dentro de
uma pasta (varre subpastas tambem) e extrai os campos principais
para uma planilha Excel (.xlsx) ou CSV.

Campos extraidos:
    - chNFe   -> Chave de Acesso (44 digitos)
    - nNF     -> Numero da Nota
    - dhEmi   -> Data de Emissao (formatada dd/mm/aaaa + ISO original)
    - CNPJ    -> CNPJ do Emitente
    - xNome   -> Nome/Razao Social do Emitente
    - vNF     -> Valor Total da Nota Fiscal
    - CFOP    -> CFOP(s) dos itens da nota (um ou mais, separados por " / ")
    - dest/xNome -> Nome do Comprador/Destinatario
    - infCpl  -> Informacoes Complementares (observacoes da nota)

INSTALACAO (se for rodar o .py direto, sem o executavel)
----------------------------------------------------------
    pip install pandas openpyxl customtkinter pillow
====================================================================
"""

import os
import sys
import threading
import xml.etree.ElementTree as ET

import pandas as pd

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
VERSAO_PROGRAMA = "v01.0"  # atualize a cada nova versao gerada


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


def processar_arquivo_xml(caminho_arquivo, callback_log=print):
    """
    Le um unico arquivo XML de NF-e e devolve um dicionario com os
    campos que nos interessam. Se o arquivo nao for uma NF-e valida
    (por exemplo, um XML corrompido ou de outro tipo), devolve None
    e avisa via 'callback_log', para o processamento nao parar no
    meio do lote.
    """
    try:
        arvore = ET.parse(caminho_arquivo)
        root = arvore.getroot()
    except ET.ParseError as erro:
        callback_log(f"[AVISO] Nao foi possivel ler o XML '{os.path.basename(caminho_arquivo)}': {erro}")
        return None

    inf_nfe = root.find(".//nfe:infNFe", NS)
    if inf_nfe is None:
        callback_log(f"[AVISO] Arquivo '{os.path.basename(caminho_arquivo)}' nao parece ser uma NF-e (tag infNFe nao encontrada).")
        return None

    dhEmi = extrair_texto(inf_nfe, "nfe:ide/nfe:dhEmi")

    dados = {
        "Chave de Acesso (chNFe)": extrair_chave_acesso(root),
        "Numero da Nota (nNF)": extrair_texto(inf_nfe, "nfe:ide/nfe:nNF"),
        "Data de Emissao": extrair_data_emissao_formatada(dhEmi),
        "Data/Hora de Emissao (dhEmi)": dhEmi,
        "CNPJ Emitente": extrair_texto(inf_nfe, "nfe:emit/nfe:CNPJ"),
        "Nome Emitente (xNome)": extrair_texto(inf_nfe, "nfe:emit/nfe:xNome"),
        "Nome do Comprador (dest/xNome)": extrair_texto(inf_nfe, "nfe:dest/nfe:xNome"),
        "Valor Total da Nota (vNF)": extrair_texto(inf_nfe, "nfe:total/nfe:ICMSTot/nfe:vNF"),
        "CFOP": extrair_cfops(inf_nfe),
        # infCpl fica dentro de <infAdic>, por isso o caminho tem os dois niveis:
        "Informacoes Complementares (infCpl)": extrair_texto(inf_nfe, "nfe:infAdic/nfe:infCpl"),
        "Arquivo de Origem": os.path.basename(caminho_arquivo),
    }
    return dados


def listar_arquivos_xml(pasta):
    """
    Varre a pasta informada (incluindo subpastas) e devolve a lista
    de caminhos completos de todos os arquivos que terminam com
    ".xml" (sem diferenciar maiusculas/minusculas).
    """
    arquivos_encontrados = []
    for pasta_atual, _subpastas, arquivos in os.walk(pasta):
        for nome_arquivo in arquivos:
            if nome_arquivo.lower().endswith(".xml"):
                caminho_completo = os.path.join(pasta_atual, nome_arquivo)
                arquivos_encontrados.append(caminho_completo)
    return arquivos_encontrados


def gerar_planilha(pasta_xmls, arquivo_saida, callback_log=print):
    """
    Funcao principal: varre a pasta, processa cada XML e salva o
    resultado consolidado em Excel (.xlsx) ou CSV, dependendo da
    extensao informada em 'arquivo_saida'. Cada etapa e reportada via
    'callback_log' (por padrao, print no console).

    Devolve uma tupla (sucesso, mensagem):
        sucesso  -> True/False
        mensagem -> texto explicando o resultado final
    """
    callback_log(f"Procurando arquivos XML em: {pasta_xmls}")
    arquivos_xml = listar_arquivos_xml(pasta_xmls)

    if not arquivos_xml:
        mensagem = "Nenhum arquivo .xml foi encontrado na pasta de origem selecionada."
        callback_log(mensagem)
        return False, mensagem

    callback_log(f"{len(arquivos_xml)} arquivo(s) XML encontrado(s). Processando...")

    linhas = []
    for caminho in arquivos_xml:
        callback_log(f"Lendo: {os.path.basename(caminho)}")
        dados = processar_arquivo_xml(caminho, callback_log=callback_log)
        if dados is not None:
            linhas.append(dados)

    if not linhas:
        mensagem = "Nenhuma nota fiscal valida foi extraida dos XMLs encontrados. Nada foi salvo."
        callback_log(mensagem)
        return False, mensagem

    df = pd.DataFrame(linhas)

    # IMPORTANTE: CNPJ, Chave de Acesso, Numero da Nota e CFOP sao
    # codigos, nao numeros. Se deixarmos o pandas/Excel decidir o
    # tipo sozinho, ele trata esses campos como numero e "come" os
    # zeros a esquerda. Por isso forcamos essas colunas a
    # permanecerem como texto.
    colunas_texto = ["Chave de Acesso (chNFe)", "CNPJ Emitente", "Numero da Nota (nNF)", "CFOP"]
    for coluna in colunas_texto:
        if coluna in df.columns:
            df[coluna] = df[coluna].astype(str)

    # O Valor Total da Nota (vNF) e numerico, entao convertemos para
    # float para que a planilha permita somas/formulas diretamente.
    if "Valor Total da Nota (vNF)" in df.columns:
        df["Valor Total da Nota (vNF)"] = pd.to_numeric(
            df["Valor Total da Nota (vNF)"], errors="coerce"
        )

    extensao = os.path.splitext(arquivo_saida)[1].lower()

    if extensao == ".csv":
        df.to_csv(arquivo_saida, index=False, sep=";", encoding="utf-8-sig")
    else:
        if extensao != ".xlsx":
            arquivo_saida = os.path.splitext(arquivo_saida)[0] + ".xlsx"

        # Gravamos usando o ExcelWriter para poder aplicar um FORMATO
        # DE CELULA "texto" (@) nas colunas de codigo, garantindo que
        # o Excel nao vai reinterpretar esses valores como numero.
        with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="NFe")
            planilha = writer.sheets["NFe"]
            for coluna in colunas_texto:
                if coluna not in df.columns:
                    continue
                indice_coluna = df.columns.get_loc(coluna) + 1
                letra_coluna = planilha.cell(row=1, column=indice_coluna).column_letter
                for linha_num in range(2, len(df) + 2):
                    celula = planilha[f"{letra_coluna}{linha_num}"]
                    celula.number_format = "@"

    mensagem = (
        f"Concluido! {len(linhas)} nota(s) fiscal(is) exportada(s) com sucesso.\n"
        f"Arquivo salvo em: {os.path.abspath(arquivo_saida)}"
    )
    callback_log(mensagem)
    return True, mensagem


NOME_ARQUIVO_SAIDA = "resultado_final_XML.xlsx"


# ====================================================================
# JANELA PRINCIPAL -- menu lateral fixo (225px) com logo/titulo/tema,
# + area de conteudo roladora a direita, organizada em "cards", igual
# ao padrao usado nos demais programas HEC.
# ====================================================================
class App(ctk.CTk):

    NOME_PROGRAMA_MENU = ("ANALISADOR DE", "NF-e")
    SUBTITULO_MENU = "Extracao de dados de\nNotas Fiscais Eletronicas"

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

        self.pasta_origem = ""
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
            "pasta_origem": self.pasta_origem,
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

        self.pasta_origem = estado_salvo["pasta_origem"]
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

        ctk.CTkLabel(card, text="1.  Pasta de origem (onde estao os arquivos XML)",
                     font=ctk.CTkFont(FONTE, 16, "bold"),
                     text_color=COR_TEXTO).pack(anchor="w", padx=16, pady=(14, 2))
        ctk.CTkLabel(card, text="Pode conter subpastas -- o programa varre tudo recursivamente.",
                     font=ctk.CTkFont(FONTE, 13), text_color=COR_MUTED,
                     justify="left", wraplength=760).pack(anchor="w", padx=16, pady=(0, 10))

        linha = ctk.CTkFrame(card, fg_color="transparent")
        linha.pack(fill="x", padx=16, pady=(0, 16))

        ctk.CTkButton(linha, text="Selecionar Pasta de Origem...", height=34,
                      fg_color=VERDE_HEC, hover_color=VERDE_ESCURO,
                      font=ctk.CTkFont(FONTE, 13, "bold"),
                      command=self._selecionar_pasta_origem).pack(side="left", padx=(0, 12))

        self._label_pasta_origem = ctk.CTkLabel(
            linha, text="Nenhuma pasta selecionada.",
            font=ctk.CTkFont(FONTE, 13), text_color=COR_MUTED,
            justify="left", wraplength=560, anchor="w",
        )
        self._label_pasta_origem.pack(side="left", fill="x", expand=True)

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

        ctk.CTkButton(linha, text="Selecionar Pasta de Destino...", height=34,
                      fg_color=VERDE_HEC, hover_color=VERDE_ESCURO,
                      font=ctk.CTkFont(FONTE, 13, "bold"),
                      command=self._selecionar_pasta_destino).pack(side="left", padx=(0, 12))

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
        ctk.CTkLabel(card, text="Extrai Chave de Acesso, Numero, Data de Emissao, CNPJ/Nome do Emitente, "
                                 "Nome do Comprador, Valor Total, CFOP e Observacoes de cada nota.",
                     font=ctk.CTkFont(FONTE, 13), text_color=COR_MUTED,
                     justify="left", wraplength=760).pack(anchor="w", padx=16, pady=(0, 10))

        linha_botoes = ctk.CTkFrame(card, fg_color="transparent")
        linha_botoes.pack(fill="x", padx=16, pady=(0, 12))

        self._botao_processar = ctk.CTkButton(
            linha_botoes, text="Gerar Planilha", height=34,
            fg_color=VERDE_HEC, hover_color=VERDE_ESCURO,
            font=ctk.CTkFont(FONTE, 13, "bold"),
            command=self._iniciar_processamento,
        )
        self._botao_processar.pack(side="left", padx=(0, 8))

        ctk.CTkButton(linha_botoes, text="Limpar Log", height=34,
                      fg_color="transparent", border_width=1, border_color=CINZA_HEC,
                      text_color=CINZA_HEC, hover_color=COR_CARD_ATIVO,
                      font=ctk.CTkFont(FONTE, 13),
                      command=self._limpar_log).pack(side="left")

        self._texto_log = ctk.CTkTextbox(
            card, height=260, fg_color=COR_FUNDO, text_color=COR_TEXTO,
            font=ctk.CTkFont("Consolas", 12), border_width=1, border_color=COR_BORDA,
            wrap="word", state="disabled",
        )
        self._texto_log.pack(fill="x", padx=16, pady=(0, 16))

    # ---------------------------------------------------------------
    def _atualizar_labels_pastas(self):
        self._label_pasta_origem.configure(
            text=self.pasta_origem if self.pasta_origem else "Nenhuma pasta selecionada.")
        self._label_pasta_destino.configure(
            text=self.pasta_destino if self.pasta_destino else "Nenhuma pasta selecionada.")

    def _selecionar_pasta_origem(self):
        pasta = filedialog.askdirectory(
            parent=self, title="Selecione a PASTA DE ORIGEM (onde estao os arquivos XML)")
        if pasta:
            self.pasta_origem = pasta
            self._label_pasta_origem.configure(text=pasta)

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

        if not self.pasta_origem:
            messagebox.showwarning("Pasta de origem nao selecionada",
                                    "Selecione a pasta de origem (onde estao os arquivos XML) antes de continuar.",
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
            args=(self.pasta_origem, caminho_saida),
            daemon=True,
        )
        thread.start()

    def _processar_em_segundo_plano(self, pasta_origem, caminho_saida):
        def log_thread_safe(mensagem):
            self.after(0, self._logar, mensagem)

        try:
            sucesso, mensagem = gerar_planilha(pasta_origem, caminho_saida, callback_log=log_thread_safe)
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
