"""
QUALIFICADOR DE CEDULAS - aplicativo unico
============================================
Decodifica um .PRN, qualifica emitentes/avalistas e gera o Word --
tudo em um só programa, com uma janela simples para escolher o arquivo.

Para rodar direto:
    python qualificador_app.py

Para transformar em .exe (rodar no Windows, com Python instalado):
    pip install python-docx pyinstaller
    pyinstaller --onefile --windowed --name Qualificador qualificador_app.py

O executavel fica em dist\\Qualificador.exe -- pode copiar esse arquivo
sozinho para qualquer pasta/computador (nao precisa mais do Python nem
do Node instalado para RODAR, só para GERAR o .exe da primeira vez).
"""
import os
import re
import sys
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

from docx import Document
from docx.shared import Cm, Pt, Emu, RGBColor
from docx.enum.text import WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ======================================================================
# ETAPA 1: DECODIFICACAO DO PRN
# ======================================================================

MAPA_OVERSTRIKE = {
    (b'a', b'~'): 'ã', (b'A', b'~'): 'Ã',
    (b'o', b'~'): 'õ', (b'O', b'~'): 'Õ',
    (b'c', b','): 'ç', (b'C', b','): 'Ç',
    (b'a', b"'"): 'á', (b'A', b"'"): 'Á',
    (b'e', b"'"): 'é', (b'E', b"'"): 'É',
    (b'i', b"'"): 'í', (b'I', b"'"): 'Í',
    (b'o', b"'"): 'ó', (b'O', b"'"): 'Ó',
    (b'u', b"'"): 'ú', (b'U', b"'"): 'Ú',
    (b'e', b'^'): 'ê', (b'E', b'^'): 'Ê',
    (b'o', b'^'): 'ô', (b'O', b'^'): 'Ô',
    (b'a', b'^'): 'â',
    (b'a', b'`'): 'à',
}


def decodificar_prn(caminho, codepage="cp850"):
    with open(caminho, "rb") as f:
        raw = f.read()

    sem_controle = re.sub(rb'\x1bC.', b'', raw)          # ESC C <n> (3 bytes, com parametro)
    sem_controle = re.sub(rb'\x1b[A-Za-z]', b'', sem_controle)  # demais ESC + 1 letra
    sem_controle = re.sub(rb'[\x0f\x12]', b'', sem_controle)    # SI / DC2

    texto = sem_controle.decode(codepage, errors="replace")

    def resolver(m):
        base, marca = m.group(1).encode(), m.group(2).encode()
        return MAPA_OVERSTRIKE.get((base, marca), m.group(1))

    texto = re.sub(r'(.)\x08(.)', resolver, texto)
    return texto


# ======================================================================
# ETAPA 2: REESTRUTURACAO (Interveniente Garantidor -> dentro de "Por aval...")
# ======================================================================

def reestruturar_documento(texto):
    padrao_span = re.compile(
        r"INTERVENIENTE\(S\)\s+GARANTIDOR\(ES\):.*?(?=CARTILHA\s+DO\s+CR[EÉ]DITO\s+RURAL:)",
        re.IGNORECASE | re.DOTALL,
    )
    m = padrao_span.search(texto)
    if not m:
        return texto

    span_original = m.group(0)
    m_intro = re.search(r"(Assino\(amos\).*?pelo\s+emitente\.)", span_original, re.IGNORECASE | re.DOTALL)
    paragrafo_intro = m_intro.group(1) if m_intro else ""

    padrao_continua = re.compile(r"Continua\s+Proxima\s+Pagina", re.IGNORECASE)
    padrao_continuacao_header = re.compile(
        r"Continua[çc][aã]o\s+do\s+instrumento\s+de\s+cr[eé]dito\s+do\s+t[ií]tulo", re.IGNORECASE
    )

    linhas_mantidas = []
    for linha in span_original.split("\n"):
        if padrao_continua.search(linha) or padrao_continuacao_header.search(linha) or linha.strip() == "":
            linhas_mantidas.append(linha)
    span_limpo = "\n".join(linhas_mantidas)

    texto = texto.replace(span_original, span_limpo, 1)
    texto = re.sub(
        r"Por\s+aval\s+ao\(s\)\s+emitente\(s\):",
        "POR AVAL AO(S) EMITENTE(S) / INTERVENIENTE(S) GARANTIDOR(ES):\r\n      \r\n      " + paragrafo_intro,
        texto, count=1, flags=re.IGNORECASE,
    )
    return texto


# ======================================================================
# ETAPA 3: BLOCOS DE DADOS COMPLETOS (cabecalho do documento)
# ======================================================================

ROTULOS_BLOCO_DADOS = [
    ("EMITENTE", r"EMITENTE\(S\):"),
    ("AVALISTA", r"Avalista\(s\):"),
    ("CONJUGE_AVALISTA", r"C[oô]njuge\s+do\s+Avalista:"),
    ("FIDUCIANTE", r"Propriet[aá]rio\(s\):"),
]

CPF_RE = re.compile(r"CPF\.?:?\s*(?:sob\s*n[°.]?\s*)?([\d.\-]{11,})", re.IGNORECASE)
RG_RE = re.compile(r"\bRG\s*([\w./-]+)\s*-\s*([A-Z/]+)", re.IGNORECASE)
NACIONALIDADE_RE = re.compile(r"Nacionalidade\s+([A-ZÀ-Ú]+)", re.IGNORECASE)
ESTADO_CIVIL_RE = re.compile(r"Nacionalidade\s+\S+,\s*([A-ZÀ-Ú]+)", re.IGNORECASE)
FILIACAO_RE = re.compile(r"filho\(a\)\s*de\s*(.+?)\s+e\s+(.+?),", re.IGNORECASE | re.DOTALL)
PROFISSAO_RE = re.compile(
    r"(?:BENS|UNIVERSAL(?:\s+DE\s+BENS)?),\s*filho\(a\).+?,\s*(.+?),\s*(?:residente|Nacionalidade)",
    re.IGNORECASE | re.DOTALL,
)
ENDERECO_RE = re.compile(r"residente\s+e\s+domiciliado\(a\)\s+no\(a\)\s*(.+?),\s*bairro\s*(.+?),", re.IGNORECASE | re.DOTALL)
TELEFONE_RE = re.compile(r"telefone\s*(.+?),", re.IGNORECASE | re.DOTALL)
EMAIL_RE = re.compile(r"endere[çc]o\s+eletr[ôo]nico\s*(.+?)\.\s*$", re.IGNORECASE | re.DOTALL)


def _limpar_espacos(valor):
    return re.sub(r"\s+", " ", valor).strip() if valor else None


def _fim_do_paragrafo(texto, inicio, limite):
    m = re.search(r"\r?\n[ \t]*\r?\n", texto[inicio:limite])
    return inicio + m.start() if m else limite


def _extrair_campos(bloco_completo):
    cpf_m = CPF_RE.search(bloco_completo)
    rg_m = RG_RE.search(bloco_completo)
    nac_m = NACIONALIDADE_RE.search(bloco_completo)
    ec_m = ESTADO_CIVIL_RE.search(bloco_completo)
    fil_m = FILIACAO_RE.search(bloco_completo)
    prof_m = PROFISSAO_RE.search(bloco_completo)
    end_m = ENDERECO_RE.search(bloco_completo)
    tel_m = TELEFONE_RE.search(bloco_completo)
    email_m = EMAIL_RE.search(bloco_completo)
    return {
        "cpf": cpf_m.group(1) if cpf_m else None,
        "rg": rg_m.group(1) if rg_m else None,
        "orgao_emissor": rg_m.group(2) if rg_m else None,
        "nacionalidade": nac_m.group(1) if nac_m else None,
        "estado_civil": ec_m.group(1) if ec_m else None,
        "pai": _limpar_espacos(fil_m.group(1)) if fil_m else None,
        "mae": _limpar_espacos(fil_m.group(2)) if fil_m else None,
        "profissao": _limpar_espacos(prof_m.group(1)) if prof_m else None,
        "endereco": _limpar_espacos(end_m.group(1)) if end_m else None,
        "bairro": _limpar_espacos(end_m.group(2)) if end_m else None,
        "telefone": _limpar_espacos(tel_m.group(1)) if tel_m else None,
        "email": _limpar_espacos(email_m.group(1)) if email_m else None,
    }


def extrair_blocos_dados(texto):
    marcadores = []
    for papel, padrao in ROTULOS_BLOCO_DADOS:
        for m in re.finditer(padrao + r"(?=[ \t]+\S)", texto):
            marcadores.append((m.start(), papel, m.group(0)))
    marcadores.sort()

    blocos = []
    for idx, (pos, papel, rotulo_txt) in enumerate(marcadores):
        limite = marcadores[idx + 1][0] if idx + 1 < len(marcadores) else pos + 1000
        fim = _fim_do_paragrafo(texto, pos, limite)
        bloco_completo = texto[pos:fim].strip()

        campos = _extrair_campos(bloco_completo)
        nome_m = re.search(re.escape(rotulo_txt.rstrip()) + r"\s*(.+?),", bloco_completo, re.IGNORECASE)
        if not campos["cpf"] and not nome_m:
            continue

        bloco_sem_rotulo = re.sub(r"^" + re.escape(rotulo_txt.rstrip()) + r"\s*", "", bloco_completo)

        blocos.append({
            "papel": papel,
            "nome": nome_m.group(1).strip() if nome_m else None,
            **campos,
            "bloco_completo": bloco_completo,
            "bloco_sem_rotulo": bloco_sem_rotulo,
        })
    return blocos


# ======================================================================
# ETAPA 4: CAMPO DE ASSINATURAS
# ======================================================================

CABECALHOS_ASSINATURA = [
    ("EMITENTE", r"^[ \t]*EMITENTE\(S\):[ \t]*\r?$"),
    ("AVALISTA", r"^[ \t]*Por\s+aval\s+ao\(s\)\s+emitente\(s\)(?:\s*/\s*INTERVENIENTE\(S\)\s+GARANTIDOR\(ES\))?:"),
    ("INTERVENIENTE_GARANTIDOR", r"^[ \t]*INTERVENIENTE\(S\)\s+GARANTIDOR\(ES\):"),
]

CPF_ASSINATURA_RE = re.compile(r"\bCPF\.?:?\s*([\d.\-]{11,})", re.IGNORECASE)
AUTORIZACAO_CONJUGE_RE = re.compile(
    r"Autoriza[çc][aã]o\s+para\s+os\s+fins\s+do\s*\r?\n?\s*Art\.\s*1\.647\s+do\s+C[oó]digo\s+Civil",
    re.IGNORECASE,
)


def extrair_assinaturas(texto):
    marcadores_papel = []
    for papel, padrao in CABECALHOS_ASSINATURA:
        for m in re.finditer(padrao, texto, re.IGNORECASE | re.MULTILINE):
            marcadores_papel.append((m.start(), papel))
    marcadores_papel.sort()

    def papel_na_posicao(pos):
        papel = None
        for p, nome_papel in marcadores_papel:
            if p <= pos:
                papel = nome_papel
            else:
                break
        return papel

    assinaturas = []
    for m in re.finditer(r"\bNOME:\s*(.+)", texto, re.IGNORECASE):
        nome = m.group(1).strip()
        pos = m.start()
        janela = texto[pos:pos + 200]
        cpf_m = CPF_ASSINATURA_RE.search(janela)
        cpf = cpf_m.group(1) if cpf_m else None
        contexto_anterior = texto[max(0, pos - 150):pos]
        eh_autorizacao_conjuge = bool(AUTORIZACAO_CONJUGE_RE.search(contexto_anterior))
        assinaturas.append({
            "papel_assinatura": papel_na_posicao(pos),
            "autorizacao_conjuge_art_1647": eh_autorizacao_conjuge,
            "nome": nome,
            "cpf": cpf,
        })
    return assinaturas


def papel_header_esperado(assinatura):
    if assinatura["autorizacao_conjuge_art_1647"]:
        return "CONJUGE_AVALISTA"
    return assinatura["papel_assinatura"]  # INTERVENIENTE_GARANTIDOR cai no fallback


def qualificar_com_blocos(assinaturas, blocos):
    blocos_por_cpf_e_papel = {}
    blocos_por_cpf_qualquer = {}
    for b in blocos:
        if not b["cpf"]:
            continue
        blocos_por_cpf_e_papel[(b["cpf"], b["papel"])] = b
        blocos_por_cpf_qualquer.setdefault(b["cpf"], b)

    qualificadas = []
    for assinatura in assinaturas:
        papel_esperado = papel_header_esperado(assinatura)
        dados_exatos = blocos_por_cpf_e_papel.get((assinatura["cpf"], papel_esperado))
        dados_completos = dados_exatos or blocos_por_cpf_qualquer.get(assinatura["cpf"])
        qualificadas.append({
            **assinatura,
            "encontrado_no_cabecalho": dados_completos is not None,
            "match_exato_de_papel": dados_exatos is not None,
            "dados_completos": dados_completos,
        })
    return qualificadas


# ======================================================================
# ETAPA 4.5: SECAO "2 - IMOVEIS" -- deteccao de hipoteca e avaliacao
# ======================================================================

PADRAO_SECAO_IMOVEIS = re.compile(r"2\s*-\s*IM[OÓ]VEIS:", re.IGNORECASE)
PADRAO_FIM_SECAO_IMOVEIS = re.compile(
    r"CONSTITUI[ÇC][ÃA]O\s+DA\s+ALIENA[ÇC][ÃA]O\s+FIDUCI[ÁA]RIA|II\s*-\s*CL[ÁA]USULAS",
    re.IGNORECASE,
)
PADRAO_HIPOTECA = re.compile(r"\bHIPOTEC\w*\b", re.IGNORECASE)
PADRAO_AVALIADO_EM = re.compile(r"Avaliado\s+em\s+([\d.,]+)\s*\(\s*(.*?)\s*\)", re.IGNORECASE | re.DOTALL)
PADRAO_VALOR_GARANTIA = re.compile(
    r"Valor\s+de\s+avalia[çc][ãa]o\s+do\s+im[óo]vel\s+para\s+fins\s+de\s+garantia\s+e\s+venda\s+em\s+p[úu]blico\s+leil[ãa]o:\s*R\$\s*([\d.,]+)\s*\(\s*(.*?)\s*\)",
    re.IGNORECASE | re.DOTALL,
)


def _normalizar_espacos(s):
    return re.sub(r"\s+", " ", s).strip()


def detectar_secao_imoveis(texto):
    """
    Localiza a secao "2 - IMOVEIS:" e verifica: (a) se ha mencao a
    hipoteca dentro dela, (b) os valores de avaliacao informados, e
    (c) se os dois valores de avaliacao sao identicos (possivel erro
    de preenchimento duplicado).　Retorna None se a secao nao existir
    no documento (ex: cedulas sem esse tipo de garantia).
    """
    m_inicio = PADRAO_SECAO_IMOVEIS.search(texto)
    if not m_inicio:
        return None

    m_fim = PADRAO_FIM_SECAO_IMOVEIS.search(texto, m_inicio.end())
    fim = m_fim.start() if m_fim else min(m_inicio.end() + 4000, len(texto))
    bloco = texto[m_inicio.start():fim]

    hipoteca_m = PADRAO_HIPOTECA.search(bloco)
    avaliado_m = PADRAO_AVALIADO_EM.search(bloco)
    garantia_m = PADRAO_VALOR_GARANTIA.search(bloco)

    valores = []
    if avaliado_m:
        valores.append({
            "rotulo": "Avaliado em",
            "valor": avaliado_m.group(1).strip(),
            "extenso": _normalizar_espacos(avaliado_m.group(2)),
        })
    if garantia_m:
        valores.append({
            "rotulo": "Valor de avaliação para fins de garantia e venda em público leilão",
            "valor": garantia_m.group(1).strip(),
            "extenso": _normalizar_espacos(garantia_m.group(2)),
        })

    duplicado = (
        len(valores) == 2
        and valores[0]["valor"].replace(" ", "") == valores[1]["valor"].replace(" ", "")
    )

    return {
        "span_inicio": m_inicio.start(),
        "span_fim": fim,
        "bloco_original": bloco,
        "tem_hipoteca": hipoteca_m is not None,
        "trecho_hipoteca": _normalizar_espacos(bloco[max(0, hipoteca_m.start() - 60):hipoteca_m.end() + 60]) if hipoteca_m else None,
        "valores_avaliacao": valores,
        "avaliacoes_duplicadas": duplicado,
    }


def montar_texto_revisao(deteccao):
    """Monta o texto EDITAVEL mostrado na tela de revisao: tudo que vem
    depois de "2 - IMOVEIS:" ate onde a secao termina (onde a avaliacao
    e encontrada) -- reflui as quebras de linha de largura fixa do PRN
    em paragrafos corridos, mais faceis de ler/editar na caixa de texto."""
    conteudo = PADRAO_SECAO_IMOVEIS.sub("", deteccao["bloco_original"], count=1)

    paragrafos_brutos = re.split(r"\r?\n[ \t]*\r?\n", conteudo)
    paragrafos = []
    for p in paragrafos_brutos:
        linha_unica = _normalizar_espacos(p.replace("\r\n", " ").replace("\n", " "))
        if linha_unica:
            paragrafos.append(linha_unica)

    aviso_linhas = []
    if deteccao["tem_hipoteca"]:
        aviso_linhas.append(f"[AVISO: HIPOTECA ENCONTRADA -- trecho: {deteccao['trecho_hipoteca']}]")
        aviso_linhas.append("")

    return "\n".join(aviso_linhas) + "\n\n".join(paragrafos)


def aplicar_correcao_imoveis(texto, deteccao, texto_revisado):
    """Substitui o bloco original da secao '2 - IMOVEIS:' pelo texto
    revisado (editado na tela), formatado com o mesmo recuo padrao do
    resto do documento."""
    linhas_novas = ["      2 - IMÓVEIS:", "      "]
    for linha in texto_revisado.split("\n"):
        linhas_novas.append("      " + linha if linha.strip() else "      ")
    bloco_novo = "\r\n".join(linhas_novas) + "\r\n      \r\n"

    return texto[:deteccao["span_inicio"]] + bloco_novo + texto[deteccao["span_fim"]:]


# ======================================================================
# ETAPA 5: GERAR O WORD (python-docx)
# ======================================================================

FATOR_LARGURA_CHAR = 0.6  # Courier New: ~0.6em de largura por caractere
FATOR_ALTURA_LINHA = 1.05
FONTE_PT = 12
FONTE_MENOR_PT = 8
MARCADOR_FONTE_MENOR = "Extrato de Operações SICOR"
RECUO_PADRAO = "      "
PADRAO_UNDERLINE = re.compile(r"^\s*_{5,}\s*$")
PADRAO_AUTORIZACAO_UMA_LINHA = re.compile(r"Autoriza[çc][aã]o\s+para\s+os\s+fins\s+do", re.IGNORECASE)
PADRAO_AUTORIZACAO_L1 = re.compile(r"Autoriza[çc][aã]o\s+para\s+os\s+fins\s+do\s*$", re.IGNORECASE)
PADRAO_AUTORIZACAO_L2 = re.compile(r"^\s*Art\.\s*1\.647\s+do\s+C[oó]digo\s+Civil", re.IGNORECASE)
PADRAO_CONTINUA_PROXIMA = re.compile(r"Continua\s+Proxima\s+Pagina", re.IGNORECASE)
PADRAO_CONTINUACAO_HEADER = re.compile(
    r"Continua[çc][aã]o\s+do\s+instrumento\s+de\s+cr[eé]dito\s+do\s+t[ií]tulo", re.IGNORECASE
)
PADRAO_PAGINA_SOZINHA = re.compile(r"^\s*Pagina:\s*\d+\s*$", re.IGNORECASE)


def _pagina_so_tem_boilerplate(linhas):
    """True se a pagina nao tem nenhum conteudo real -- só linhas em
    branco e/ou os avisos de paginacao (Continua Proxima Pagina /
    Continuacao do instrumento...)."""
    for linha in linhas:
        if linha.strip() == "":
            continue
        if PADRAO_CONTINUA_PROXIMA.search(linha) or PADRAO_CONTINUACAO_HEADER.search(linha):
            continue
        return False
    return True


def _linhas_de_qualificacao(dados):
    if not dados or not dados.get("bloco_sem_rotulo"):
        return []
    bloco = dados["bloco_sem_rotulo"]
    linhas = bloco.split("\r\n") if "\r\n" in bloco else bloco.split("\n")
    return [RECUO_PADRAO + l if i == 0 else l for i, l in enumerate(linhas)]


PADRAO_TABULAR_ESPACO = re.compile(r"\S {4,}\S")
PADRAO_SEPARADOR = re.compile(r"^[\-=_]{5,}$")
PADRAO_CAMPO_FORMULARIO = re.compile(r"\.{3,}\s*:")
PADRAO_CABECALHO_TABELA = re.compile(r"Nro\s+Data|Ref\.BACEN|Quadro\s+Resumo", re.IGNORECASE)


def _reflow_prosa(linhas_texto):
    """Junta linhas quebradas por largura fixa em texto corrido, normalizando
    espacamento duplo (usado no PRN para justificar) para espaco simples."""
    texto = " ".join(l.strip() for l in linhas_texto if l.strip())
    return re.sub(r" {2,}", " ", texto)


def _classificar_bloco(linhas_texto):
    if any(PADRAO_UNDERLINE.match(l) for l in linhas_texto):
        return "assinatura"

    linhas_stripped = [l.strip() for l in linhas_texto]
    if any(PADRAO_SEPARADOR.match(l) or PADRAO_CABECALHO_TABELA.search(l) for l in linhas_stripped):
        return "tabular"
    # so classifica como tabular se a MAIORIA das linhas do bloco tiverem
    # esse padrao de colunas -- uma unica linha de prosa muito justificada
    # tambem pode ter 1 gap grande, mas nao o bloco inteiro
    if len(linhas_stripped) >= 2:
        com_gap = sum(1 for l in linhas_stripped if PADRAO_TABULAR_ESPACO.search(l))
        if com_gap / len(linhas_stripped) >= 0.6:
            return "tabular"

    if any(PADRAO_CAMPO_FORMULARIO.search(l) for l in linhas_texto):
        return "formulario"
    return "prosa"


def _adicionar_campo_pagina(paragraph):
    """Insere um campo de numero de pagina automatico (se atualiza sozinho
    no Word, nao e texto fixo)."""
    run = paragraph.add_run()
    fld_inicio = OxmlElement("w:fldChar")
    fld_inicio.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld_fim = OxmlElement("w:fldChar")
    fld_fim.set(qn("w:fldCharType"), "end")
    run._r.append(fld_inicio)
    run._r.append(instr)
    run._r.append(fld_fim)
    run.font.name = "Courier New"
    run.font.size = Pt(FONTE_PT)
    run.font.color.rgb = RGBColor(0, 0, 0)


PADRAO_CARTILHA = re.compile(r"CARTILHA\s+DO\s+CR[EÉ]DITO\s+RURAL:", re.IGNORECASE)


def gerar_docx(texto, blocos, assinaturas_qualificadas, caminho_saida):
    # a partir da "CARTILHA DO CREDITO RURAL:" (se houver), o documento
    # nao usa mais a paginacao "Continuacao do instrumento..." no PRN
    # original -- e um anexo informativo separado. Por pedido do usuario,
    # essa parte fica fiel ao original (sem reflow/tabela/qualificacao,
    # sem cabecalho/rodape fixo), como uma segunda secao do Word.
    m_cartilha = PADRAO_CARTILHA.search(texto)
    if m_cartilha:
        texto_estruturado = texto[: m_cartilha.start()]
        texto_verbatim = texto[m_cartilha.start():]
    else:
        texto_estruturado = texto
        texto_verbatim = ""

    paginas_textos = texto_estruturado.split("\x0c")

    # posicao (no texto todo) do ultimo campo de assinatura -- so
    # removemos paginas em branco QUE VENHAM DEPOIS desse ponto (efeito
    # colateral da reestruturacao do Interveniente Garantidor, que pode
    # deixar uma pagina so com "Continua Proxima Pagina").
    pos_ultima_assinatura = 0
    for m in re.finditer(r"\bNOME:\s*.+", texto_estruturado, re.IGNORECASE):
        pos_ultima_assinatura = m.end()

    paginas_linhas = []
    offsets_paginas = []
    offset = 0
    for pagina in paginas_textos:
        offsets_paginas.append(offset)
        offset += len(pagina) + 1  # +1 pelo \x0c removido no split
        paginas_linhas.append([l.replace("\r", "") for l in pagina.replace("\r\n", "\n").split("\n")])

    while paginas_linhas and all(l.strip() == "" for l in paginas_linhas[-1]):
        paginas_linhas.pop()
        offsets_paginas.pop()

    i = 1  # nunca remove a pagina 0 (cabecalho do titulo)
    while i < len(paginas_linhas):
        depois_da_ultima_assinatura = offsets_paginas[i] >= pos_ultima_assinatura
        if depois_da_ultima_assinatura and _pagina_so_tem_boilerplate(paginas_linhas[i]):
            paginas_linhas[i - 1] = [l for l in paginas_linhas[i - 1] if not PADRAO_CONTINUA_PROXIMA.search(l)]
            paginas_linhas.pop(i)
            offsets_paginas.pop(i)
            # nao incrementa i -- a lista encolheu
        else:
            i += 1

    for linhas in paginas_linhas:
        idx = 0
        while idx < len(linhas) - 1:
            if PADRAO_AUTORIZACAO_L1.search(linhas[idx]) and PADRAO_AUTORIZACAO_L2.search(linhas[idx + 1]):
                linhas[idx] = linhas[idx].rstrip() + " " + linhas[idx + 1].strip()
                linhas.pop(idx + 1)
            idx += 1

    apartir_sicor = False
    linhas_flat = []
    for idx_pagina, linhas in enumerate(paginas_linhas):
        for idx_linha, linha in enumerate(linhas):
            if MARCADOR_FONTE_MENOR in linha:
                apartir_sicor = True
            linhas_flat.append({
                "texto": linha,
                "quebra_forcada": idx_pagina > 0 and idx_linha == 0,
                "fonte_menor": apartir_sicor,
            })

    indice_assinatura = 0
    unidades = []
    i = 0
    while i < len(linhas_flat):
        atual = linhas_flat[i]

        eh_marcador_paginacao = (
            PADRAO_CONTINUA_PROXIMA.search(atual["texto"])
            or PADRAO_CONTINUACAO_HEADER.search(atual["texto"])
            or PADRAO_PAGINA_SOZINHA.match(atual["texto"])
        )
        if eh_marcador_paginacao:
            i += 1
            continue  # nao vira paragrafo -- agora vem do cabecalho/rodape fixo
        if atual["texto"].strip() == "":
            unidades.append({"tipo": "linha", "linha": atual})
            i += 1
            continue

        # junta linhas nao-vazias consecutivas em um bloco, pra classificar
        # (para tambem nos marcadores de paginacao, que ficam sempre isolados)
        bloco = [atual]
        j = i + 1
        while j < len(linhas_flat) and linhas_flat[j]["texto"].strip() != "" and not (
            PADRAO_CONTINUA_PROXIMA.search(linhas_flat[j]["texto"])
            or PADRAO_CONTINUACAO_HEADER.search(linhas_flat[j]["texto"])
            or PADRAO_PAGINA_SOZINHA.match(linhas_flat[j]["texto"])
        ):
            bloco.append(linhas_flat[j])
            j += 1

        tipo = _classificar_bloco([l["texto"] for l in bloco])

        if tipo == "assinatura":
            linhas_bloco = [{**bloco[0], "texto": bloco[0]["texto"].strip()}]
            resto = bloco[1:]
            if resto and PADRAO_AUTORIZACAO_UMA_LINHA.search(resto[0]["texto"]):
                linhas_bloco.append({**resto[0], "texto": resto[0]["texto"].strip()})
                resto = resto[1:]
            if resto and re.search(r"\bNOME:\s*.+", resto[0]["texto"], re.IGNORECASE):
                linha_nome = resto[0]
                linha_cpf = resto[1] if len(resto) > 1 else None
                tem_cpf = linha_cpf is not None and re.match(r"^\s*CPF", linha_cpf["texto"], re.IGNORECASE)
                assinatura = assinaturas_qualificadas[indice_assinatura] if indice_assinatura < len(assinaturas_qualificadas) else None

                if assinatura and tem_cpf:
                    indice_assinatura += 1
                    dados = assinatura["dados_completos"]
                    bloco_texto = dados.get("bloco_sem_rotulo") if dados else None
                    if bloco_texto:
                        linhas_originais = bloco_texto.split("\r\n") if "\r\n" in bloco_texto else bloco_texto.split("\n")
                        texto_reflow = _reflow_prosa(linhas_originais)
                        linhas_bloco.append({"texto": texto_reflow, "quebra_forcada": False, "fonte_menor": atual["fonte_menor"]})
                    else:
                        linhas_bloco.append(linha_nome)
                        linhas_bloco.append(linha_cpf)
                else:
                    linhas_bloco.append(linha_nome)
                    if tem_cpf:
                        linhas_bloco.append(linha_cpf)
            unidades.append({"tipo": "bloco", "linhas": linhas_bloco, "quebra_forcada": atual["quebra_forcada"]})

        elif tipo == "tabular":
            for l in bloco:
                unidades.append({"tipo": "linha", "linha": l})

        elif tipo == "formulario":
            # mantem cada linha separada (campo a campo), so normaliza
            # espacamento duplo e tira o recuo fixo
            for l in bloco:
                l2 = {**l, "texto": re.sub(r" {2,}", " ", l["texto"].strip())}
                unidades.append({"tipo": "linha", "linha": l2})

        else:  # prosa
            texto_reflow = _reflow_prosa([l["texto"] for l in bloco])
            unidades.append({
                "tipo": "linha",
                "linha": {"texto": texto_reflow, "quebra_forcada": atual["quebra_forcada"], "fonte_menor": atual["fonte_menor"]},
            })

        i = j

    doc = Document()
    secao = doc.sections[0]
    secao.orientation = WD_ORIENT.PORTRAIT
    secao.page_width = Cm(21.59)   # 8.5in em retrato = largura (Oficio/Legal)
    secao.page_height = Cm(35.56)  # 14in em retrato = altura
    # topo com folga suficiente pra caber o cabecalho (distancia + altura da
    # linha) sem invadir o espaco do corpo do texto -- se nao, o corpo
    # "perde" espaco que o calculo de linhas por pagina nao contava, e
    # sobra um pouco de texto vazando pra pagina seguinte.
    secao.top_margin = Cm(0.7)
    secao.bottom_margin = Cm(1.25)
    secao.left_margin = Cm(1.5)
    secao.right_margin = Cm(1.5)

    # -------- cabecalho e rodape fixos (Continuacao/Pagina/Continua Proxima) --------
    m_titulo = re.search(r"T[IÍ]TULO\.+:\s*(\S+)", texto, re.IGNORECASE)
    numero_titulo = m_titulo.group(1) if m_titulo else ""
    largura_util_cm = 21.59 - 1.5 - 1.5  # largura da pagina menos as margens laterais

    secao.header_distance = Cm(0.15)
    secao.footer_distance = Cm(0.3)
    secao.different_first_page_header_footer = True

    PRETO = RGBColor(0, 0, 0)

    def _run_padrao(paragraph, texto_run):
        r = paragraph.add_run(texto_run)
        r.font.name = "Courier New"
        r.font.size = Pt(FONTE_PT)
        r.font.color.rgb = PRETO
        return r

    # cabecalho da 1a pagina: so "Pagina: 1", alinhado a direita
    p_header_1 = secao.first_page_header.paragraphs[0]
    p_header_1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _run_padrao(p_header_1, "Pagina: ")
    _adicionar_campo_pagina(p_header_1)

    # cabecalho das demais paginas: "Continuacao do instrumento..." a
    # esquerda + "Pagina: N" (campo automatico) a direita, usando tab
    # com parada alinhada a direita na borda util da pagina
    p_header = secao.header.paragraphs[0]
    p_header.paragraph_format.tab_stops.add_tab_stop(Cm(largura_util_cm), WD_TAB_ALIGNMENT.RIGHT)
    _run_padrao(p_header, f"Continuação do instrumento de crédito do título {numero_titulo}.\tPagina: ")
    _adicionar_campo_pagina(p_header)

    # rodape (todas as paginas): "Continua Proxima Pagina" alinhado a direita
    for rodape in (secao.footer, secao.first_page_footer):
        p_footer = rodape.paragraphs[0]
        p_footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        _run_padrao(p_footer, "Continua Proxima Pagina")
    # -------------------------------------------------------------------------

    altura_pagina_pt = 14 * 72  # Legal em retrato: altura = 14"
    margem_topo_pt = 0.7 / 2.54 * 72
    margem_rodape_pt = 1.25 / 2.54 * 72
    altura_util_pt = altura_pagina_pt - margem_topo_pt - margem_rodape_pt

    largura_pagina_pt = 8.5 * 72  # Legal em retrato: largura = 8.5"
    margem_lateral_pt = 1.5 / 2.54 * 72
    largura_util_pt = largura_pagina_pt - (2 * margem_lateral_pt)
    max_chars_por_linha = max(1, int(largura_util_pt // (FONTE_PT * FATOR_LARGURA_CHAR)))

    def linhas_visuais(texto_linha):
        """Estima em quantas linhas visuais uma linha logica vai quebrar
        no Word, ja que em retrato varias linhas do PRN (ate 127
        caracteres) nao cabem mais numa unica linha da pagina."""
        n = len(texto_linha)
        return max(1, -(-n // max_chars_por_linha))  # ceil sem precisar de math.ceil

    altura_linha_pt = FONTE_PT * FATOR_ALTURA_LINHA
    # folga de seguranca de 15% -- em paragrafos corridos longos (texto em
    # prosa), pequenos erros na estimativa de largura de caractere se
    # acumulam ao longo do paragrafo inteiro, e podem fazer sobrar um
    # pouco de texto vazando pra proxima pagina. Preferimos terminar a
    # pagina um pouco mais cedo a arriscar esse vazamento.
    max_linhas_por_pagina = int((altura_util_pt // altura_linha_pt) * 0.85)

    def adicionar_paragrafo(texto_linha, quebra_antes, fonte_menor):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.space_after = Pt(0)
        pf.space_before = Pt(0)
        # AT_LEAST (nao EXACTLY): se uma linha longa quebrar em 2 linhas
        # visuais (retrato e mais estreito que paisagem), o paragrafo
        # cresce em vez de sobrepor o texto da linha seguinte.
        pf.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
        pf.line_spacing = Pt(FONTE_MENOR_PT * FATOR_ALTURA_LINHA if fonte_menor else altura_linha_pt)
        pf.page_break_before = quebra_antes
        run = p.add_run(texto_linha if len(texto_linha) else " ")
        run.font.name = "Courier New"
        run.font.size = Pt(FONTE_MENOR_PT if fonte_menor else FONTE_PT)
        return p

    linhas_restantes = max_linhas_por_pagina
    primeira_unidade = True
    for u in unidades:
        if u["tipo"] == "bloco":
            num_linhas = sum(linhas_visuais(l["texto"]) for l in u["linhas"])
        else:
            num_linhas = linhas_visuais(u["linha"]["texto"])
        quebra_forcada = u["quebra_forcada"] if u["tipo"] == "bloco" else u["linha"]["quebra_forcada"]
        quebra_antes = False

        if not primeira_unidade and quebra_forcada:
            quebra_antes = True
            linhas_restantes = max_linhas_por_pagina
        elif not primeira_unidade and num_linhas <= max_linhas_por_pagina and num_linhas > linhas_restantes:
            quebra_antes = True
            linhas_restantes = max_linhas_por_pagina

        linhas_restantes -= num_linhas
        primeira_unidade = False

        if u["tipo"] == "linha":
            adicionar_paragrafo(u["linha"]["texto"], quebra_antes, u["linha"]["fonte_menor"])
        else:
            for idx, l in enumerate(u["linhas"]):
                adicionar_paragrafo(l["texto"], quebra_antes if idx == 0 else False, l["fonte_menor"])

    # -------- 2a secao: Cartilha em diante, fiel ao original -------------
    # nenhuma das regras (reflow/tabela/qualificacao/cabecalho fixo) se
    # aplica aqui -- cada linha do PRN vira um paragrafo tal como veio,
    # e essa secao nao tem cabecalho/rodape (a paginacao "oficial" do
    # instrumento acaba na ultima assinatura).
    if texto_verbatim:
        secao2 = doc.add_section(WD_SECTION.NEW_PAGE)
        secao2.page_width = secao.page_width
        secao2.page_height = secao.page_height
        secao2.top_margin = secao.top_margin
        secao2.bottom_margin = secao.bottom_margin
        secao2.left_margin = secao.left_margin
        secao2.right_margin = secao.right_margin

        secao2.header.is_linked_to_previous = False
        secao2.footer.is_linked_to_previous = False
        secao2.first_page_header.is_linked_to_previous = False
        secao2.first_page_footer.is_linked_to_previous = False
        for p in secao2.header.paragraphs:
            p.text = ""
        for p in secao2.footer.paragraphs:
            p.text = ""
        for p in secao2.first_page_header.paragraphs:
            p.text = ""
        for p in secao2.first_page_footer.paragraphs:
            p.text = ""

        paginas_verbatim = texto_verbatim.split("\x0c")
        for idx_pag, pagina in enumerate(paginas_verbatim):
            linhas_pagina = [l.replace("\r", "") for l in pagina.replace("\r\n", "\n").split("\n")]
            for idx_lin, linha in enumerate(linhas_pagina):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                pf = p.paragraph_format
                pf.space_after = Pt(0)
                pf.space_before = Pt(0)
                pf.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
                pf.line_spacing = Pt(altura_linha_pt)
                pf.page_break_before = idx_pag > 0 and idx_lin == 0
                run = p.add_run(linha if linha else " ")
                run.font.name = "Courier New"
                run.font.size = Pt(FONTE_PT)
                run.font.color.rgb = RGBColor(0, 0, 0)

    doc.save(caminho_saida)


# ======================================================================
# ORQUESTRACAO
# ======================================================================

def sanitizar_nome_arquivo(s):
    return re.sub(r'[\\/:*?"<>|]', "-", s).strip()


def preparar_arquivo(caminho_prn, log=print):
    """Decodifica, reestrutura, qualifica e detecta a secao de imoveis.
    NAO gera o Word ainda -- devolve tudo que a tela de revisao precisa."""
    log(f"Decodificando {os.path.basename(caminho_prn)} ...")
    texto = decodificar_prn(caminho_prn)
    texto = reestruturar_documento(texto)

    log("Extraindo blocos de dados e qualificando assinaturas ...")
    blocos = extrair_blocos_dados(texto)
    assinaturas = extrair_assinaturas(texto)
    assinaturas_qualificadas = qualificar_com_blocos(assinaturas, blocos)
    log(f"  {len(assinaturas_qualificadas)} assinatura(s) qualificada(s).")

    deteccao_imoveis = detectar_secao_imoveis(texto)
    if deteccao_imoveis:
        avisos = []
        if deteccao_imoveis["tem_hipoteca"]:
            avisos.append("hipoteca encontrada")
        if deteccao_imoveis["avaliacoes_duplicadas"]:
            avisos.append("valores de avaliação idênticos (possível duplicidade)")
        if avisos:
            log("  Seção '2 - IMÓVEIS': " + "; ".join(avisos) + ".")
        else:
            log("  Seção '2 - IMÓVEIS' encontrada (sem hipoteca, valores de avaliação distintos).")

    emitente = next((b for b in blocos if b["papel"] == "EMITENTE"), None)
    nome_emitente = emitente["nome"].strip() if emitente and emitente.get("nome") else None
    m_titulo = re.search(r"T[IÍ]TULO\.+:\s*(\S+)", texto, re.IGNORECASE)
    numero_titulo = m_titulo.group(1) if m_titulo else None

    return {
        "caminho_prn": caminho_prn,
        "texto": texto,
        "blocos": blocos,
        "assinaturas_qualificadas": assinaturas_qualificadas,
        "deteccao_imoveis": deteccao_imoveis,
        "nome_emitente": nome_emitente,
        "numero_titulo": numero_titulo,
    }


def finalizar_geracao(dados, pasta_saida, texto_revisao_imoveis=None, log=print):
    """Aplica a correcao da secao de imoveis (se houver texto revisado) e
    gera o Word."""
    texto = dados["texto"]
    if dados["deteccao_imoveis"] and texto_revisao_imoveis is not None:
        texto = aplicar_correcao_imoveis(texto, dados["deteccao_imoveis"], texto_revisao_imoveis)
        log("Aplicando correções da seção '2 - IMÓVEIS' ...")
        # a secao de imoveis mudou de tamanho -- refaz blocos/assinaturas
        # em cima do texto corrigido, pra qualificacao continuar correta.
        blocos = extrair_blocos_dados(texto)
        assinaturas_qualificadas = qualificar_com_blocos(extrair_assinaturas(texto), blocos)
    else:
        blocos = dados["blocos"]
        assinaturas_qualificadas = dados["assinaturas_qualificadas"]

    if dados["nome_emitente"] and dados["numero_titulo"]:
        nome_arquivo = sanitizar_nome_arquivo(f"{dados['nome_emitente']} - {dados['numero_titulo']}") + ".docx"
    else:
        nome_arquivo = os.path.splitext(os.path.basename(dados["caminho_prn"]))[0] + "_decodificado.docx"

    caminho_saida = os.path.join(pasta_saida, nome_arquivo)

    log("Gerando o Word ...")
    try:
        gerar_docx(texto, blocos, assinaturas_qualificadas, caminho_saida)
    except PermissionError:
        alternativo = os.path.join(pasta_saida, f"{os.path.splitext(nome_arquivo)[0]}_novo.docx")
        log(f'AVISO: "{nome_arquivo}" esta aberto/bloqueado. Salvando como {os.path.basename(alternativo)} ...')
        gerar_docx(texto, blocos, assinaturas_qualificadas, alternativo)
        caminho_saida = alternativo

    log(f"Pronto: {caminho_saida}")
    return caminho_saida


# ======================================================================
# INTERFACE GRAFICA
# ======================================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Qualificador de Cédulas")
        self.geometry("700x560")
        self.resizable(True, True)

        self.caminho_prn = tk.StringVar()
        self.dados_preparados = None  # resultado de preparar_arquivo()

        frame_top = tk.Frame(self, padx=12, pady=12)
        frame_top.pack(fill="x")
        tk.Label(frame_top, text="Arquivo .PRN:").pack(side="left")
        tk.Entry(frame_top, textvariable=self.caminho_prn, width=55).pack(side="left", padx=6, fill="x", expand=True)
        tk.Button(frame_top, text="Selecionar...", command=self.selecionar_arquivo).pack(side="left")

        frame_botoes = tk.Frame(self, padx=12)
        frame_botoes.pack(fill="x")
        self.botao_processar = tk.Button(
            frame_botoes, text="1. Processar", command=self.processar, bg="#1565c0", fg="white"
        )
        self.botao_processar.pack(side="left", pady=6)
        self.botao_gerar = tk.Button(
            frame_botoes, text="2. Gerar Word", command=self.gerar, bg="#2e7d32", fg="white", state="disabled"
        )
        self.botao_gerar.pack(side="left", padx=(8, 0), pady=6)

        # -------- painel de revisao da secao "2 - IMOVEIS" (aparece só quando aplicável) --------
        self.frame_revisao = tk.LabelFrame(
            self, text="Revisão: seção \"2 - IMÓVEIS\" (edite abaixo se precisar corrigir)", padx=8, pady=8
        )
        self.aviso_label = tk.Label(self.frame_revisao, text="", fg="#b71c1c", justify="left")
        self.aviso_label.pack(fill="x", anchor="w")
        self.texto_revisao = scrolledtext.ScrolledText(self.frame_revisao, height=8, wrap="word")
        self.texto_revisao.pack(fill="both", expand=True, pady=(4, 0))
        # o frame_revisao só é exibido (pack) quando ha algo pra revisar

        self.log_text = scrolledtext.ScrolledText(self, padx=8, pady=8, state="disabled", wrap="word", height=10)
        self.log_text.pack(fill="both", expand=True, padx=12, pady=(6, 12))

    def log(self, mensagem):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", mensagem + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        self.update_idletasks()

    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
            title="Selecione o arquivo .PRN",
            filetypes=[("Arquivos PRN", "*.PRN;*.prn"), ("Todos os arquivos", "*.*")],
        )
        if caminho:
            self.caminho_prn.set(caminho)

    def processar(self):
        caminho = self.caminho_prn.get().strip()
        if not caminho or not os.path.isfile(caminho):
            messagebox.showerror("Erro", "Selecione um arquivo .PRN valido primeiro.")
            return

        self.botao_processar.config(state="disabled")
        self.botao_gerar.config(state="disabled")
        self.frame_revisao.pack_forget()
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        try:
            self.dados_preparados = preparar_arquivo(caminho, log=self.log)
            deteccao = self.dados_preparados["deteccao_imoveis"]

            if deteccao:
                avisos = []
                if deteccao["tem_hipoteca"]:
                    avisos.append("⚠ Hipoteca encontrada na seção de imóveis.")
                if deteccao["avaliacoes_duplicadas"]:
                    avisos.append("⚠ Os dois valores de avaliação do imóvel são IDÊNTICOS -- confira se está certo.")
                if not avisos:
                    avisos.append("Nenhum problema encontrado -- revise/ajuste os valores abaixo se quiser.")
                self.aviso_label.config(text="\n".join(avisos))

                self.texto_revisao.delete("1.0", "end")
                self.texto_revisao.insert("1.0", montar_texto_revisao(deteccao))
                self.frame_revisao.pack(fill="both", expand=False, padx=12, pady=(0, 6))
            else:
                self.log("  (Documento não tem seção '2 - IMÓVEIS' -- nada para revisar aqui.)")

            self.log("\nProcessado. Confira acima (e o painel de revisão, se apareceu) e clique em \"2. Gerar Word\".")
            self.botao_gerar.config(state="normal")
        except Exception as e:
            self.log("\nERRO: " + str(e))
            self.log(traceback.format_exc())
            messagebox.showerror("Erro ao processar", str(e))
        finally:
            self.botao_processar.config(state="normal")

    def gerar(self):
        if not self.dados_preparados:
            messagebox.showerror("Erro", "Clique em \"1. Processar\" primeiro.")
            return

        self.botao_gerar.config(state="disabled")
        try:
            texto_revisao = None
            if self.dados_preparados["deteccao_imoveis"]:
                texto_revisao = self.texto_revisao.get("1.0", "end-1c")

            pasta_saida = os.path.dirname(self.caminho_prn.get().strip())
            caminho_saida = finalizar_geracao(
                self.dados_preparados, pasta_saida, texto_revisao_imoveis=texto_revisao, log=self.log
            )
            self.log("\nConcluído com sucesso!")
            resposta = messagebox.askyesno(
                "Concluído", f"Arquivo gerado:\n{caminho_saida}\n\nDeseja abrir a pasta agora?"
            )
            if resposta:
                os.startfile(pasta_saida)  # Windows
        except Exception as e:
            self.log("\nERRO: " + str(e))
            self.log(traceback.format_exc())
            messagebox.showerror("Erro ao gerar", str(e))
        finally:
            self.botao_gerar.config(state="normal")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        # modo linha de comando: python qualificador_app.py caminho.PRN
        # (usa a secao de imoveis tal como veio, sem tela de revisao)
        caminho = sys.argv[1]
        pasta_saida = os.path.dirname(os.path.abspath(caminho))
        dados = preparar_arquivo(caminho)
        finalizar_geracao(dados, pasta_saida)
    else:
        app = App()
        app.mainloop()
