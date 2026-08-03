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
from tkinter import filedialog, messagebox, scrolledtext, ttk

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


def _remover_boilerplate_paginacao(bloco):
    """Remove as linhas de paginacao (Continua Proxima Pagina /
    Continuacao do instrumento... / Pagina: N sozinha) de dentro de um
    trecho de texto -- usado pra descricao do imovel nao "puxar" esses
    artefatos quando a secao atravessa mais de uma pagina do PRN."""
    linhas = bloco.split("\n")
    linhas_limpas = [
        l for l in linhas
        if not (
            PADRAO_CONTINUA_PROXIMA.search(l)
            or PADRAO_CONTINUACAO_HEADER.search(l)
            or PADRAO_PAGINA_SOZINHA.match(l)
        )
    ]
    return "\n".join(linhas_limpas)


def detectar_secao_imoveis(texto):
    """
    Localiza a secao "2 - IMOVEIS:" e verifica: (a) se ha mencao a
    hipoteca dentro dela, (b) os valores de avaliacao informados, e
    (c) se os dois valores de avaliacao sao identicos (possivel erro
    de preenchimento duplicado). Retorna None se a secao nao existir
    no documento (ex: cedulas sem esse tipo de garantia).

    A secao pode atravessar varias paginas do PRN (descricao de imovel
    grande) -- o limite de seguranca e bem folgado (50 mil caracteres)
    pra nao cortar nada nesse caso, e as linhas de paginacao que
    aparecem no meio do caminho sao removidas antes de analisar.
    """
    m_inicio = PADRAO_SECAO_IMOVEIS.search(texto)
    if not m_inicio:
        return None

    m_fim = PADRAO_FIM_SECAO_IMOVEIS.search(texto, m_inicio.end())
    fim = m_fim.start() if m_fim else min(m_inicio.end() + 50000, len(texto))
    bloco = _remover_boilerplate_paginacao(texto[m_inicio.start():fim])

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
    resto do documento. Devolve (novo_texto, posicao_onde_a_descricao_termina)
    -- essa posicao e usada por aplicar_clausula_superveniencia para
    inserir a clausula logo apos a descricao, mesmo que ela tenha mudado
    de tamanho."""
    linhas_novas = ["      2 - IMÓVEIS:", "      "]
    for linha in texto_revisado.split("\n"):
        linhas_novas.append("      " + linha if linha.strip() else "      ")
    bloco_novo = "\r\n".join(linhas_novas) + "\r\n      \r\n"

    novo_texto = texto[:deteccao["span_inicio"]] + bloco_novo + texto[deteccao["span_fim"]:]
    posicao_fim = deteccao["span_inicio"] + len(bloco_novo)
    return novo_texto, posicao_fim


# --------------------------------------------------------------------
# Clausula de Superveniencia (aba "Opcoes")
# --------------------------------------------------------------------

TEXTO_CLAUSULA_SUPERVENIENCIA = (
    "CONSTITUIÇÃO DA ALIENAÇÃO FIDUCIÁRIA DA PROPRIEDADE SUPERVENIENTE: "
    "O imóvel objeto da garantia é constituído em ALIENAÇÃO FIDUCIÁRIA SUPERVENIENTE, "
    "nos termos do §3º do art. 22 da Lei n° 9.514/1997, com a redação dada pela Lei "
    "14.711/2023. Parágrafo Único: Na hipótese de constituição de alienações "
    "fiduciárias sucessivas da propriedade superveniente e/ou quando houver a extensão "
    "da garantia fiduciária, havendo o inadimplemento de quaisquer das obrigações "
    "garantidas pelo mesmo imóvel, fica desde já autorizado ao CREDOR determinar o "
    "vencimento antecipado das demais obrigações de que for titular, nos termos do §6º "
    "do artigo 22 da Lei 9.514/2023, com a redação dada pela Lei 14.711/2023."
)

PADRAO_CLAUSULA_SUPERVENIENCIA = re.compile(
    r"CONSTITUI[ÇC][ÃA]O\s+DA\s+ALIENA[ÇC][ÃA]O\s+FIDUCI[ÁA]RIA\s+DA\s+PROPRIEDADE\s+SUPERVENIENTE",
    re.IGNORECASE,
)

# marcadores invisiveis (area de uso privado do Unicode -- nao aparecem
# como texto de verdade) usados so pra sinalizar pro gerador do Word
# qual trecho foi ACRESCENTADO pelo script, pra colorir diferente do
# resto e o colaborador saber o que foi adicionado x o que veio no PRN.
MARCADOR_ADICIONADO_INICIO = "\ue000"
MARCADOR_ADICIONADO_FIM = "\ue001"


def detectar_clausula_superveniencia(texto):
    """True se a clausula de superveniencia ja existe no documento."""
    return bool(PADRAO_CLAUSULA_SUPERVENIENCIA.search(texto))


def aplicar_clausula_superveniencia(texto, posicao_insercao, adicionar):
    """Insere a clausula de superveniencia logo na posicao indicada (logo
    apos a descricao do imovel), marcada como "acrescentada" (cor
    diferente no Word). So insere se adicionar=True -- quem decide ANTES
    se deve chamar isso ou nao e a logica de "so acrescenta se nao
    existia e o colaborador marcou o checkbox" (ver finalizar_geracao)."""
    if not adicionar:
        return texto

    bloco = (
        "\r\n      \r\n      "
        + MARCADOR_ADICIONADO_INICIO
        + TEXTO_CLAUSULA_SUPERVENIENCIA
        + MARCADOR_ADICIONADO_FIM
        + "\r\n      \r\n"
    )
    return texto[:posicao_insercao] + bloco + texto[posicao_insercao:]


# --------------------------------------------------------------------
# Cartorio (aba "Opcoes") -- deteccao simples por enquanto; no futuro vai
# cruzar com uma planilha de criterios por cartorio.
# --------------------------------------------------------------------

PADRAO_CARTORIO = re.compile(
    r"Registro\s+de\s+Im[oó]veis[^\r\n,\.]*(?:Comarca\s+de\s+[^\r\n,\.]+)?",
    re.IGNORECASE,
)


def detectar_cartorio(texto):
    """Tenta achar o nome do cartorio/comarca citado na descricao do
    imovel (ex: 'Registro de Imoveis Comarca de Mandaguacu - PR').
    Retorna None se nao encontrar nada parecido."""
    m = PADRAO_CARTORIO.search(texto)
    return _normalizar_espacos(m.group(0)) if m else None



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
PADRAO_CABECALHO_TABELA = re.compile(r"^(Nro\s+Data|Ref\.BACEN|Quadro\s+Resumo\s+da)", re.IGNORECASE)
# numeracao de clausula/item no comeco da linha (1., 3., (iii), a), i.) --
# sozinha ja cria um "gap" de varios espacos que nao e uma tabela de verdade
PADRAO_NUMERACAO_CLAUSULA = re.compile(r"^\s*(?:\(?[ivxlcdm]{1,6}\)|\(?[a-zA-Z]\)|\d+[.\)])\s+", re.IGNORECASE)


def _reflow_prosa(linhas_texto):
    """Junta linhas quebradas por largura fixa em texto corrido, normalizando
    espacamento duplo (usado no PRN para justificar) para espaco simples."""
    texto = " ".join(l.strip() for l in linhas_texto if l.strip())
    return re.sub(r" {2,}", " ", texto)


def _classificar_bloco(linhas_texto):
    if any(PADRAO_UNDERLINE.match(l) for l in linhas_texto):
        return "assinatura"

    linhas_stripped = [l.strip() for l in linhas_texto]
    # so classifica como tabular com marcadores confiaveis: linha
    # separadora (----/====) ou cabecalho de tabela conhecido. A
    # heuristica antiga (maioria das linhas com espaco largo) foi
    # removida -- texto justificado normal tambem pode ter gaps largos
    # quando a linha tem poucas palavras pra preencher a largura, e
    # isso gerava falso positivo em clausulas normais.
    if any(PADRAO_SEPARADOR.match(l) or PADRAO_CABECALHO_TABELA.search(l) for l in linhas_stripped):
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


def _adicionar_texto_exceto_ultima_pagina(paragraph, texto_condicional):
    """Insere um campo condicional do Word: so mostra 'texto_condicional'
    quando a pagina atual NAO for a ultima pagina da secao (campo IF com
    PAGE e SECTIONPAGES aninhados). Assim o rodape "Continua Proxima
    Pagina" desaparece sozinho na ultima pagina, sem precisar criar
    nenhuma secao/pagina extra -- o Word calcula isso dinamicamente,
    entao continua correto mesmo se o texto for editado depois."""

    def _fld_char(tipo):
        el = OxmlElement("w:fldChar")
        el.set(qn("w:fldCharType"), tipo)
        return el

    def _instr(texto):
        el = OxmlElement("w:instrText")
        el.set(qn("xml:space"), "preserve")
        el.text = texto
        return el

    run = paragraph.add_run()
    run.font.name = "Courier New"
    run.font.size = Pt(FONTE_PT)
    run.font.color.rgb = RGBColor(0, 0, 0)
    r = run._r

    r.append(_fld_char("begin"))  # campo IF (externo)
    r.append(_instr(' IF '))
    r.append(_fld_char("begin"))  # campo PAGE (aninhado)
    r.append(_instr(' PAGE '))
    r.append(_fld_char("end"))
    r.append(_instr(' <> '))
    r.append(_fld_char("begin"))  # campo SECTIONPAGES (aninhado)
    r.append(_instr(' SECTIONPAGES '))
    r.append(_fld_char("end"))
    r.append(_instr(f' "{texto_condicional}" ""'))
    r.append(_fld_char("separate"))
    # texto de exibicao inicial (o Word recalcula ao abrir o arquivo)
    texto_inicial = OxmlElement("w:t")
    texto_inicial.set(qn("xml:space"), "preserve")
    texto_inicial.text = texto_condicional
    r.append(texto_inicial)
    r.append(_fld_char("end"))


PADRAO_CARTILHA = re.compile(r"CARTILHA\s+DO\s+CR[EÉ]DITO\s+RURAL:", re.IGNORECASE)


PADRAO_ITEM_LISTA_CARTILHA = re.compile(r"^\s*-\s")
PADRAO_TERMINA_FRASE = re.compile(r"[.:;]\s*$")
PADRAO_SICREDI_FONE = re.compile(r"SICREDI\s+FONE", re.IGNORECASE)


def _mesclar_topicos_cartilha(linhas):
    """Junta linhas quebradas por largura fixa em um paragrafo por
    'topico' (paragrafo, cabecalho ou item de lista com '-'), evitando
    que um heading fique despedacado por uma quebra de pagina do PRN
    original caindo no meio da frase (2+ linhas em branco seguidas e a
    frase anterior nao termina em pontuacao = artefato de pagina, nao
    paragrafo novo). Itens de lista com '-' sempre comecam um topico
    novo, mesmo sem linha em branco antes (e como vem no PRN)."""
    resultado = []
    topico_atual = []

    def ultima_linha_nao_vazia():
        for l in reversed(topico_atual):
            if l.strip():
                return l.strip()
        return ""

    def fechar_topico():
        if topico_atual:
            resultado.append(_reflow_prosa(topico_atual))
            topico_atual.clear()

    i = 0
    while i < len(linhas):
        linha = linhas[i]
        if linha.strip() == "":
            j = i
            while j < len(linhas) and linhas[j].strip() == "":
                j += 1
            n_blanks = j - i
            ultima = ultima_linha_nao_vazia()
            termina_frase = bool(PADRAO_TERMINA_FRASE.search(ultima)) if ultima else True

            if n_blanks == 1 or termina_frase:
                fechar_topico()
                resultado.append("")
            # senao (2+ linhas em branco e a frase anterior nao termina em
            # pontuacao): artefato de quebra de pagina no meio da frase --
            # ignora as linhas em branco e continua o topico atual
            i = j
            continue

        if PADRAO_ITEM_LISTA_CARTILHA.match(linha) and topico_atual:
            fechar_topico()

        topico_atual.append(linha)
        i += 1

    fechar_topico()
    return resultado


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
                # NAO forca quebra de pagina nos limites originais do PRN --
                # agora que o texto reflui em prosa compacta, isso so
                # desperdicava espaco (paginas terminando com muito branco
                # sobrando). O conteudo flui naturalmente; so o bloco de
                # assinatura continua protegido contra ficar dividido.
                "quebra_forcada": False,
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

    # rodape (todas as paginas, exceto a ultima da secao): "Continua
    # Proxima Pagina" alinhado a direita -- campo condicional, some
    # sozinho na ultima pagina (onde termina o instrumento e comeca a
    # Cartilha, um documento diferente).
    for rodape in (secao.footer, secao.first_page_footer):
        p_footer = rodape.paragraphs[0]
        p_footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        _adicionar_texto_exceto_ultima_pagina(p_footer, "Continua Proxima Pagina")
    # -------------------------------------------------------------------------

    altura_linha_pt = FONTE_PT * FATOR_ALTURA_LINHA

    def adicionar_paragrafo(texto_linha, fonte_menor, manter_com_proximo=False, manter_junto=False):
        eh_adicionado = MARCADOR_ADICIONADO_INICIO in texto_linha
        texto_limpo = texto_linha.replace(MARCADOR_ADICIONADO_INICIO, "").replace(MARCADOR_ADICIONADO_FIM, "")

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
        # deixa o WORD decidir onde quebrar a pagina (dinamico de
        # verdade) -- so pedimos pra ele manter certos paragrafos juntos,
        # em vez de calcular por fora quantas linhas cabem.
        pf.keep_with_next = manter_com_proximo
        pf.keep_together = manter_junto
        run = p.add_run(texto_limpo if len(texto_limpo) else " ")
        run.font.name = "Courier New"
        run.font.size = Pt(FONTE_MENOR_PT if fonte_menor else FONTE_PT)
        if eh_adicionado:
            run.font.color.rgb = RGBColor(0, 51, 153)  # azul -- sinaliza texto acrescentado pelo script
        return p

    for u in unidades:
        if u["tipo"] == "linha":
            adicionar_paragrafo(u["linha"]["texto"], u["linha"]["fonte_menor"])
        else:
            n_linhas_bloco = len(u["linhas"])
            for idx, l in enumerate(u["linhas"]):
                eh_ultima_linha_do_bloco = idx == n_linhas_bloco - 1
                adicionar_paragrafo(
                    l["texto"], l["fonte_menor"],
                    manter_com_proximo=not eh_ultima_linha_do_bloco,
                    manter_junto=True,
                )

    # -------- 2a e 3a secoes: Cartilha e Extrato SICOR, fieis ao original -----
    # nenhuma das regras (reflow/tabela/qualificacao/cabecalho fixo) se
    # aplica aqui -- cada linha do PRN vira um paragrafo tal como veio,
    # e essas secoes nao tem cabecalho/rodape (a paginacao "oficial" do
    # instrumento acaba na ultima assinatura). Margens quase zero nas
    # duas ("sem espaçamento"); a do Extrato SICOR fica em fonte 7pt.
    m_sicor = re.search(re.escape(MARCADOR_FONTE_MENOR), texto_verbatim)
    if m_sicor:
        texto_cartilha_only = texto_verbatim[: m_sicor.start()]
        texto_sicor_only = texto_verbatim[m_sicor.start():]
    else:
        texto_cartilha_only = texto_verbatim
        texto_sicor_only = ""

    def _nova_secao_sem_espacamento(margem_lateral):
        s = doc.add_section(WD_SECTION.NEW_PAGE)
        s.page_width = secao.page_width
        s.page_height = secao.page_height
        s.top_margin = Cm(0.2)
        s.bottom_margin = Cm(0.2)
        s.left_margin = margem_lateral
        s.right_margin = margem_lateral
        s.header.is_linked_to_previous = False
        s.footer.is_linked_to_previous = False
        s.first_page_header.is_linked_to_previous = False
        s.first_page_footer.is_linked_to_previous = False
        for p in s.header.paragraphs:
            p.text = ""
        for p in s.footer.paragraphs:
            p.text = ""
        for p in s.first_page_header.paragraphs:
            p.text = ""
        for p in s.first_page_footer.paragraphs:
            p.text = ""
        return s

    def _adicionar_paginas_verbatim(texto_bloco, tamanho_fonte, mesclar_topicos=False):
        linhas_todas = []
        for pagina in texto_bloco.split("\x0c"):
            linhas_todas.extend(l.replace("\r", "") for l in pagina.replace("\r\n", "\n").split("\n"))

        linhas_centralizadas = set()

        if mesclar_topicos:
            # separa o bloco final "SICREDI FONE" (4 linhas fixas de
            # contato) do resto -- ele NAO deve ser mesclado com o
            # topico anterior, e cada uma das 4 linhas fica centralizada
            # e sem o preenchimento de espacos antigo (calculado pra
            # largura de pagina antiga).
            idx_sicredi = next(
                (i for i, l in enumerate(linhas_todas) if PADRAO_SICREDI_FONE.search(l)), None
            )
            if idx_sicredi is not None:
                antes = linhas_todas[:idx_sicredi]
                bloco_contato = [l.strip() for l in linhas_todas[idx_sicredi:idx_sicredi + 4] if l.strip()]
                depois = linhas_todas[idx_sicredi + 4:]

                topicos = _mesclar_topicos_cartilha(antes)
                # tira blanks do final e poe exatamente 3 antes do bloco de contato
                while topicos and topicos[-1] == "":
                    topicos.pop()
                topicos.extend([""] * 3)

                inicio_centralizadas = len(topicos)
                topicos.extend(bloco_contato)
                linhas_centralizadas = set(range(inicio_centralizadas, len(topicos)))

                topicos.extend(_mesclar_topicos_cartilha(depois))
                linhas_todas = topicos
            else:
                linhas_todas = _mesclar_topicos_cartilha(linhas_todas)

        altura_linha_local_pt = tamanho_fonte * FATOR_ALTURA_LINHA

        # deixa o Word decidir onde quebrar a pagina -- so pede pra
        # manter cada topico/linha inteiro junto (nao dividir no meio),
        # sem calcular quantas linhas cabem por pagina.
        for idx_linha, linha in enumerate(linhas_todas):
            p = doc.add_paragraph()
            centralizar = idx_linha in linhas_centralizadas or MARCADOR_FONTE_MENOR in linha
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if centralizar else WD_ALIGN_PARAGRAPH.LEFT
            pf = p.paragraph_format
            pf.space_after = Pt(0)
            pf.space_before = Pt(0)
            pf.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
            pf.line_spacing = Pt(altura_linha_local_pt)
            pf.keep_together = True
            run = p.add_run(linha if linha else " ")
            run.font.name = "Courier New"
            run.font.size = Pt(tamanho_fonte)
            run.font.color.rgb = RGBColor(0, 0, 0)

    if texto_cartilha_only:
        _nova_secao_sem_espacamento(Cm(1.5))
        _adicionar_paginas_verbatim(texto_cartilha_only, FONTE_PT, mesclar_topicos=True)

    if texto_sicor_only:
        _nova_secao_sem_espacamento(Cm(0))
        _adicionar_paginas_verbatim(texto_sicor_only, 7)

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

    tem_clausula_superveniencia = detectar_clausula_superveniencia(texto)
    cartorio_detectado = detectar_cartorio(texto)
    log(f"  Cláusula de Superveniência: {'encontrada' if tem_clausula_superveniencia else 'não encontrada'}.")
    log(f"  Cartório detectado: {cartorio_detectado or '(nenhum)'}")

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
        "tem_clausula_superveniencia": tem_clausula_superveniencia,
        "cartorio_detectado": cartorio_detectado,
        "nome_emitente": nome_emitente,
        "numero_titulo": numero_titulo,
    }


def finalizar_geracao(dados, pasta_saida, texto_revisao_imoveis=None, checkbox_superveniencia=None, log=print):
    """Aplica a correcao da secao de imoveis (se houver texto revisado),
    acrescenta a clausula de Superveniencia se o checkbox estiver marcado
    E ela nao existia no documento original (evita duplicidade), e gera
    o Word."""
    texto = dados["texto"]
    posicao_fim_descricao = dados["deteccao_imoveis"]["span_fim"] if dados["deteccao_imoveis"] else None

    if dados["deteccao_imoveis"] and texto_revisao_imoveis is not None:
        log("Aplicando correções da seção '2 - IMÓVEIS' ...")
        texto, posicao_fim_descricao = aplicar_correcao_imoveis(texto, dados["deteccao_imoveis"], texto_revisao_imoveis)

    # so acrescenta a clausula se: (a) o checkbox esta marcado agora, e
    # (b) ela NAO existia no documento original -- se ja existia, nao
    # precisa (evita duplicidade), mesmo que o checkbox continue marcado.
    deve_acrescentar_clausula = (
        checkbox_superveniencia
        and not dados["tem_clausula_superveniencia"]
        and posicao_fim_descricao is not None
    )
    if deve_acrescentar_clausula:
        log("Acrescentando cláusula de Superveniência (marcada em azul no Word) ...")
        texto = aplicar_clausula_superveniencia(texto, posicao_fim_descricao, adicionar=True)

    if texto_revisao_imoveis is not None or deve_acrescentar_clausula:
        # o texto mudou -- refaz blocos/assinaturas em cima do texto
        # corrigido, pra qualificacao continuar correta.
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
        self.geometry("760x620")
        self.resizable(True, True)

        self.caminho_prn = tk.StringVar()
        self.var_superveniencia = tk.BooleanVar(value=False)
        self.var_cartorio = tk.StringVar()
        self.var_procurador = tk.StringVar()
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

        # -------- abas --------
        self.abas = ttk.Notebook(self)
        self.abas.pack(fill="both", expand=True, padx=12, pady=(6, 6))

        self._montar_aba_descricao()
        self._montar_aba_opcoes()

        self.log_text = scrolledtext.ScrolledText(self, padx=8, pady=8, state="disabled", wrap="word", height=8)
        self.log_text.pack(fill="both", expand=False, padx=12, pady=(0, 12))

    # ------------------------------------------------------------------
    # montagem das abas
    # ------------------------------------------------------------------
    def _montar_aba_descricao(self):
        aba = tk.Frame(self.abas, padx=8, pady=8)
        self.abas.add(aba, text="Descrição da Matrícula")

        self.aviso_label = tk.Label(aba, text="", fg="#b71c1c", justify="left", anchor="w")
        self.aviso_label.pack(fill="x", anchor="w")

        tk.Label(
            aba, text="Descrição do imóvel (seção \"2 - IMÓVEIS\") -- edite aqui se precisar corrigir:",
            anchor="w", justify="left",
        ).pack(fill="x", anchor="w", pady=(6, 2))

        self.texto_revisao = scrolledtext.ScrolledText(aba, height=16, wrap="word")
        self.texto_revisao.pack(fill="both", expand=True)

    def _montar_aba_opcoes(self):
        aba = tk.Frame(self.abas, padx=12, pady=12)
        self.abas.add(aba, text="Opções")

        frame_superveniencia = tk.Frame(aba)
        frame_superveniencia.pack(fill="x", anchor="w", pady=(0, 4))
        self.check_superveniencia = tk.Checkbutton(
            frame_superveniencia, text="Cláusula de Superveniência", variable=self.var_superveniencia,
        )
        self.check_superveniencia.pack(side="left")
        self.label_status_superveniencia = tk.Label(frame_superveniencia, text="", fg="#555555")
        self.label_status_superveniencia.pack(side="left", padx=(8, 0))

        tk.Label(
            aba,
            text="Se desmarcado ao processar, a cláusula não foi encontrada na cédula -- marque\n"
                 "para acrescentá-la (aparece em azul no Word gerado). Se já existir, marcar de\n"
                 "novo não duplica.",
            fg="#555555", justify="left", anchor="w",
        ).pack(fill="x", anchor="w", pady=(0, 12))

        tk.Label(aba, text="Cartório:").pack(anchor="w")
        self.combo_cartorio = ttk.Combobox(aba, textvariable=self.var_cartorio, values=[], width=60)
        self.combo_cartorio.pack(fill="x", anchor="w", pady=(0, 12))

        tk.Label(aba, text="Procuradores:").pack(anchor="w")
        self.combo_procurador = ttk.Combobox(aba, textvariable=self.var_procurador, values=[], width=60)
        self.combo_procurador.pack(fill="x", anchor="w")

    # ------------------------------------------------------------------
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
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        try:
            self.dados_preparados = preparar_arquivo(caminho, log=self.log)
            deteccao = self.dados_preparados["deteccao_imoveis"]

            self.texto_revisao.delete("1.0", "end")
            if deteccao:
                avisos = []
                if deteccao["tem_hipoteca"]:
                    avisos.append("⚠ Hipoteca encontrada na seção de imóveis.")
                if deteccao["avaliacoes_duplicadas"]:
                    avisos.append("⚠ Os dois valores de avaliação do imóvel são IDÊNTICOS -- confira se está certo.")
                if not avisos:
                    avisos.append("Nenhum problema encontrado -- revise/ajuste os valores abaixo se quiser.")
                self.aviso_label.config(text="\n".join(avisos))
                self.texto_revisao.insert("1.0", montar_texto_revisao(deteccao))
            else:
                self.aviso_label.config(text="Documento não tem seção \"2 - IMÓVEIS\" -- nada para revisar aqui.")

            # aba Opcoes: clausula de superveniencia
            tem_clausula = self.dados_preparados["tem_clausula_superveniencia"]
            self.var_superveniencia.set(tem_clausula)
            self.label_status_superveniencia.config(
                text="(encontrada na cédula)" if tem_clausula else "(não encontrada -- marque para acrescentar)"
            )

            # aba Opcoes: cartorio detectado
            cartorio = self.dados_preparados["cartorio_detectado"]
            if cartorio:
                self.combo_cartorio.config(values=[cartorio])
                self.var_cartorio.set(cartorio)
            else:
                self.combo_cartorio.config(values=[])
                self.var_cartorio.set("")

            self.log("\nProcessado. Confira as abas acima e clique em \"2. Gerar Word\".")
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
                self.dados_preparados,
                pasta_saida,
                texto_revisao_imoveis=texto_revisao,
                checkbox_superveniencia=self.var_superveniencia.get(),
                log=self.log,
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
