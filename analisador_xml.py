# -*- coding: utf-8 -*-
"""
extrai_nfe.py
=============

Script para ler todos os arquivos XML de NF-e (Nota Fiscal Eletronica)
dentro de uma pasta e extrair os campos principais para uma planilha
Excel (.xlsx) ou CSV.

Campos extraidos:
    - chNFe   -> Chave de Acesso (44 digitos)
    - nNF     -> Numero da Nota
    - dhEmi   -> Data/Hora de Emissao (tambem exportada so a data, ja formatada)
    - CNPJ    -> CNPJ do Emitente
    - xNome   -> Nome/Razao Social do Emitente
    - vNF     -> Valor Total da Nota Fiscal
    - CFOP    -> CFOP(s) dos itens da nota (um ou mais, separados por " / ")
    - dest/xNome -> Nome do Comprador/Destinatario
    - infCpl  -> Informacoes Complementares (observacoes da nota)

------------------------------------------------------------------
COMO INSTALAR AS DEPENDENCIAS
------------------------------------------------------------------
Este script usa a biblioteca padrao do Python para ler o XML
(xml.etree.ElementTree), entao a unica dependencia externa e o
pandas (para montar a tabela e exportar para Excel/CSV).

No terminal (cmd/PowerShell), com o ambiente virtual ativado (se
voce usar um), rode:

    pip install pandas openpyxl

O "openpyxl" e necessario para o pandas conseguir gravar arquivos
.xlsx.

------------------------------------------------------------------
COMO RODAR
------------------------------------------------------------------
1) Coloque todos os arquivos .xml das notas fiscais dentro de uma
   pasta (pode ter subpastas, o script varre tudo recursivamente).

2) Rode o script normalmente:

    python extrai_nfe.py

   O script vai abrir duas janelas do Windows, nessa ordem:

   a) Uma janela pedindo para voce selecionar a PASTA DE ORIGEM
      (onde estao os arquivos XML).

   b) Uma janela pedindo para voce selecionar a PASTA DE DESTINO
      (onde a planilha final sera salva).

   Nao existe mais a opcao de passar os caminhos por linha de
   comando - tudo e feito pelas janelas, para facilitar o uso no
   dia a dia.

3) O resultado sera um arquivo "notas_fiscais_consolidado.xlsx"
   salvo na pasta de destino escolhida, com uma linha por nota
   fiscal. Ao final, aparece uma mensagem na tela confirmando
   quantas notas foram processadas e onde o arquivo foi salvo.
"""

import os
import sys
import xml.etree.ElementTree as ET

import pandas as pd

# tkinter e filedialog/messagebox fazem parte da biblioteca padrao
# do Python (nao precisa instalar nada extra) e sao responsaveis
# pelas janelas graficas de selecao de pasta e pelas mensagens de
# aviso/erro que aparecem na tela.
import tkinter as tk
from tkinter import filedialog, messagebox


# ------------------------------------------------------------------
# NAMESPACE DA NF-e
# ------------------------------------------------------------------
# Todo XML de NF-e usa esse "xmlns" (namespace). Sem informar isso
# para o ElementTree, as buscas por tag (ex: "nNF") nao encontram
# nada, porque internamente a tag "real" e algo como
# "{http://www.portalfiscal.inf.br/nfe}nNF".
#
# Por isso criamos esse dicionario "ns" e usamos o prefixo "nfe:"
# em todas as buscas (find/findall) mais abaixo.
NS = {"nfe": "http://www.portalfiscal.inf.br/nfe"}


def extrair_texto(elemento, caminho, namespaces=NS):
    """
    Funcao auxiliar para buscar um campo dentro do XML e devolver o
    texto dele, ja tratando os casos em que o campo nao existe.

    elemento   -> o "no" do XML onde vamos comecar a busca (ex: a
                  tag <infNFe>)
    caminho    -> o caminho da tag que queremos, no formato do
                  ElementTree, usando o prefixo "nfe:" para respeitar
                  o namespace (ex: "ide/nNF")

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

    1) Dentro de <protNFe><infProt><chNFe> - isso acontece quando o
       arquivo e o "XML completo" (nfeProc), que inclui o protocolo
       de autorizacao da SEFAZ. E o caso mais comum quando voce
       baixa o XML pelo portal ou pelo sistema do seu ERP.

    2) Caso o arquivo seja "so a NFe" (sem o protocolo), a chave nao
       existe como um campo separado - ela fica embutida no
       atributo "Id" da tag <infNFe>, no formato "NFe" + 44 digitos
       (ex: Id="NFe35260601616867000139550030000158021923256329").
       Nesse caso, extraimos a chave removendo o prefixo "NFe".

    Essa funcao tenta o caminho 1 primeiro e, se nao achar nada,
    cai para o caminho 2.
    """
    # Tentativa 1: dentro do protocolo de autorizacao
    chave = extrair_texto(root, ".//nfe:protNFe/nfe:infProt/nfe:chNFe")
    if chave:
        return chave

    # Tentativa 2: a partir do atributo "Id" de <infNFe>
    inf_nfe = root.find(".//nfe:infNFe", NS)
    if inf_nfe is not None:
        id_attr = inf_nfe.get("Id", "")  # ex: "NFe3526060161..."
        if id_attr.startswith("NFe"):
            return id_attr[3:]  # remove o prefixo "NFe"
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
        # Os primeiros 10 caracteres de uma data ISO sao sempre "aaaa-mm-dd"
        ano, mes, dia = data_emissao_iso[:10].split("-")
        return f"{dia}/{mes}/{ano}"
    except (ValueError, IndexError):
        return data_emissao_iso


def extrair_cfops(inf_nfe):
    """
    Uma nota fiscal pode ter varios itens (tag <det>), e cada item
    tem o seu proprio CFOP dentro de <det><prod><CFOP>. Normalmente
    todos os itens de uma mesma nota usam o mesmo CFOP, mas notas
    mistas podem ter mais de um.

    Essa funcao coleta todos os CFOPs encontrados, remove repetidos
    (mantendo a ordem) e devolve como uma unica string, separada por
    " / " quando houver mais de um.
    """
    cfops_encontrados = []
    for cfop_elemento in inf_nfe.findall(".//nfe:det/nfe:prod/nfe:CFOP", NS):
        if cfop_elemento.text:
            cfop = cfop_elemento.text.strip()
            if cfop and cfop not in cfops_encontrados:
                cfops_encontrados.append(cfop)
    return " / ".join(cfops_encontrados)


def processar_arquivo_xml(caminho_arquivo):
    """
    Le um unico arquivo XML de NF-e e devolve um dicionario com os
    campos que nos interessam. Se o arquivo nao for uma NF-e valida
    (por exemplo, um XML corrompido ou de outro tipo), devolve None
    e avisa no console, para o script nao parar no meio do processo.
    """
    try:
        arvore = ET.parse(caminho_arquivo)
        root = arvore.getroot()
    except ET.ParseError as erro:
        print(f"  [AVISO] Nao foi possivel ler o XML '{caminho_arquivo}': {erro}")
        return None

    # <infNFe> e o "no principal" que contem quase todos os dados da nota.
    inf_nfe = root.find(".//nfe:infNFe", NS)
    if inf_nfe is None:
        print(f"  [AVISO] Arquivo '{caminho_arquivo}' nao parece ser uma NF-e (tag infNFe nao encontrada).")
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


def gerar_planilha(pasta_xmls, arquivo_saida):
    """
    Funcao principal: varre a pasta, processa cada XML e salva o
    resultado consolidado em Excel (.xlsx) ou CSV, dependendo da
    extensao informada em 'arquivo_saida'.

    Devolve uma tupla (sucesso, mensagem):
        sucesso  -> True/False
        mensagem -> texto explicando o resultado (usado tanto no
                    console quanto nas janelas de aviso do tkinter)
    """
    print(f"Procurando arquivos XML em: {pasta_xmls}")
    arquivos_xml = listar_arquivos_xml(pasta_xmls)

    if not arquivos_xml:
        mensagem = "Nenhum arquivo .xml foi encontrado na pasta de origem selecionada."
        print(mensagem)
        return False, mensagem

    print(f"{len(arquivos_xml)} arquivo(s) XML encontrado(s). Processando...\n")

    linhas = []
    for caminho in arquivos_xml:
        print(f"Lendo: {caminho}")
        dados = processar_arquivo_xml(caminho)
        if dados is not None:
            linhas.append(dados)

    if not linhas:
        mensagem = "Nenhuma nota fiscal valida foi extraida dos XMLs encontrados. Nada foi salvo."
        print(mensagem)
        return False, mensagem

    # Monta a tabela com pandas. Cada item da lista "linhas" vira uma
    # linha da planilha, e as chaves do dicionario viram as colunas.
    df = pd.DataFrame(linhas)

    # IMPORTANTE: CNPJ e Chave de Acesso sao codigos, nao numeros.
    # Se deixarmos o pandas/Excel decidir o tipo sozinho, ele trata
    # esses campos como numero e "come" os zeros a esquerda (ex: o
    # CNPJ "01616867000139" viraria "1616867000139", o que e errado).
    # Por isso forcamos essas colunas a permanecerem como texto.
    colunas_texto = ["Chave de Acesso (chNFe)", "CNPJ Emitente", "Numero da Nota (nNF)", "CFOP"]
    for coluna in colunas_texto:
        if coluna in df.columns:
            df[coluna] = df[coluna].astype(str)

    # O Valor Total da Nota (vNF) e numerico, entao convertemos para
    # float para que a planilha permita somas/formulas diretamente.
    # Se algum valor vier vazio ou invalido, ele vira NaN (celula em
    # branco) em vez de quebrar o script.
    if "Valor Total da Nota (vNF)" in df.columns:
        df["Valor Total da Nota (vNF)"] = pd.to_numeric(
            df["Valor Total da Nota (vNF)"], errors="coerce"
        )

    # Decide se salva como Excel ou CSV com base na extensao do
    # arquivo de saida escolhido pelo usuario.
    extensao = os.path.splitext(arquivo_saida)[1].lower()

    if extensao == ".csv":
        # utf-8-sig evita problemas de acentuacao ao abrir no Excel
        df.to_csv(arquivo_saida, index=False, sep=";", encoding="utf-8-sig")
    else:
        # Qualquer outra extensao (ou nenhuma) vira .xlsx
        if extensao != ".xlsx":
            arquivo_saida = os.path.splitext(arquivo_saida)[0] + ".xlsx"

        # Gravamos usando o ExcelWriter para poder aplicar um FORMATO
        # DE CELULA "texto" (@) nas colunas de codigo, garantindo que
        # o Excel nao vai reinterpretar esses valores como numero
        # mesmo depois do arquivo ja estar salvo.
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
        f"Concluido!\n\n"
        f"{len(linhas)} nota(s) fiscal(is) exportada(s) com sucesso.\n\n"
        f"Arquivo salvo em:\n{os.path.abspath(arquivo_saida)}"
    )
    print("\n" + mensagem)
    return True, mensagem


def escolher_pasta_origem(janela_raiz):
    """
    Abre a janela nativa do Windows para o usuario escolher a PASTA
    DE ORIGEM, ou seja, a pasta onde estao os arquivos XML das notas
    fiscais.

    'janela_raiz' e a janela "invisivel" do tkinter que serve apenas
    de base para as caixas de dialogo (ela nunca aparece na tela).
    """
    pasta = filedialog.askdirectory(
        parent=janela_raiz,
        title="Selecione a PASTA DE ORIGEM (onde estao os arquivos XML)",
    )
    return pasta  # vem como string vazia "" se o usuario clicar em Cancelar


def escolher_pasta_destino(janela_raiz):
    """
    Abre a janela nativa do Windows para o usuario escolher a PASTA
    DE DESTINO, ou seja, onde a planilha final (.xlsx) sera salva.
    """
    pasta = filedialog.askdirectory(
        parent=janela_raiz,
        title="Selecione a PASTA DE DESTINO (onde a planilha sera salva)",
    )
    return pasta


if __name__ == "__main__":
    NOME_ARQUIVO_SAIDA = "resultado_final_XML.xlsx"

    # Cria a janela "raiz" do tkinter, mas mantem ela escondida -
    # ela so existe para servir de base para as janelas de dialogo
    # (selecionar pasta) e as mensagens de aviso/erro abaixo.
    janela_raiz = tk.Tk()
    janela_raiz.withdraw()

    # ----------------------------------------------------------
    # PASSO 1: pedir a pasta de origem (onde estao os XMLs)
    # ----------------------------------------------------------
    pasta_xmls = escolher_pasta_origem(janela_raiz)
    if not pasta_xmls:
        # Usuario clicou em "Cancelar" -> encerra o script sem erro.
        messagebox.showwarning(
            "Operacao cancelada",
            "Nenhuma pasta de origem foi selecionada. O script sera encerrado.",
        )
        sys.exit(0)

    # ----------------------------------------------------------
    # PASSO 2: pedir a pasta de destino (onde salvar a planilha)
    # ----------------------------------------------------------
    pasta_destino = escolher_pasta_destino(janela_raiz)
    if not pasta_destino:
        messagebox.showwarning(
            "Operacao cancelada",
            "Nenhuma pasta de destino foi selecionada. O script sera encerrado.",
        )
        sys.exit(0)

    # Monta o caminho completo do arquivo de saida, juntando a pasta
    # de destino escolhida com o nome padrao da planilha.
    caminho_saida = os.path.join(pasta_destino, NOME_ARQUIVO_SAIDA)

    # ----------------------------------------------------------
    # PASSO 3: processar os XMLs e gerar a planilha
    # ----------------------------------------------------------
    sucesso, mensagem = gerar_planilha(pasta_xmls, caminho_saida)

    # Mostra o resultado final numa janela, para o usuario nao
    # precisar olhar o console (que normalmente fica escondido
    # quando o script e rodado com duplo-clique no Windows).
    if sucesso:
        messagebox.showinfo("Concluido", mensagem)
    else:
        messagebox.showerror("Nao foi possivel concluir", mensagem)

    janela_raiz.destroy()