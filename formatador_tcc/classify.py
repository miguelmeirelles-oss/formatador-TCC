"""Classifica cada parágrafo do .docx do aluno em uma categoria de formatação.

A classificação nunca lê nem altera o conteúdo textual -- apenas decide qual
estilo ABNT (ver config.py) deve ser aplicado a cada parágrafo. Dois sinais
são usados, nessa ordem de prioridade:

1. O nome do estilo do Word já aplicado ao parágrafo (se o aluno partiu do
   modelo oficial e usou o painel de Estilos corretamente).
2. Heurísticas sobre o texto (títulos numerados digitados manualmente,
   palavras-chave de seções sem número, "Figura"/"Quadro"/"Tabela"/"Fonte:").

Quando nenhum sinal é conclusivo, o parágrafo é classificado como "outro" e
recebe apenas normalização leve (fonte/tamanho), preservando o layout que o
aluno já tinha (ex.: capa, folha de rosto, ficha catalográfica).
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from .texto import normalizar
from .config import SECOES_SEM_NUMERO, SECOES_POS_TEXTUAIS_TITULO1

RE_TITULO_NUMERADO = re.compile(
    r"^(?P<prefixo>\d{1,2}(?:\.\d{1,2}){0,4})\.?\s+[A-ZÀ-Ú0-9]"
)
RE_LEGENDA = re.compile(
    r"^(figura|quadro|tabela|gr[aá]fico|equa[cç][aã]o)\s*\d+", re.IGNORECASE
)
RE_FONTE = re.compile(r"^fonte\s*:", re.IGNORECASE)
# Linha de sumário: termina em número de página (com ou sem tab/pontos de
# preenchimento antes), ex.: "1\tINTRODUÇÃO\t12" ou "REFERÊNCIAS  22". As
# páginas pré-textuais (antes da Introdução) costumam ser numeradas em
# algarismos romanos minúsculos (ex.: "Sumário\tix", "Lista de Figuras\txii"),
# por isso também são aceitos aqui.
RE_LINHA_SUMARIO = re.compile(r"(\d{1,4}|[ivxlcdmIVXLCDM]{1,7})\s*$")

_MAPA_ESTILO_WORD = {
    "heading 1": "titulo1",
    "título 1": "titulo1",
    "titulo 1": "titulo1",
    "heading 2": "titulo2",
    "título 2": "titulo2",
    "titulo 2": "titulo2",
    "heading 3": "titulo3",
    "título 3": "titulo3",
    "titulo 3": "titulo3",
    "heading 4": "titulo4",
    "título 4": "titulo4",
    "titulo 4": "titulo4",
    "heading 5": "titulo5",
    "título 5": "titulo5",
    "titulo 5": "titulo5",
    "título de seção": "titulo_sem_numero",
    "titulo de secao": "titulo_sem_numero",
    "title": "titulo_sem_numero",
    "quadro": "legenda",
    "caption": "fonte_ilustracao",
}

# Zonas de documento usadas para saber como classificar parágrafos "neutros".
ZONA_PRE_TEXTUAL = "PRE_TEXTUAL"
ZONA_RESUMO = "RESUMO"
ZONA_ABSTRACT = "ABSTRACT"
ZONA_SUMARIO = "SUMARIO"
ZONA_CORPO = "BODY"
ZONA_REFERENCIAS = "REFERENCIAS"
ZONA_POS_TEXTUAL = "POS_TEXTUAL"


@dataclass
class EstadoClassificacao:
    zona: str = ZONA_PRE_TEXTUAL
    logo_apos_titulo: bool = False


_ESTILOS_TITULO_NUMERADO = {
    nome for nome, categoria in _MAPA_ESTILO_WORD.items()
    if categoria in ("titulo1", "titulo2", "titulo3", "titulo4", "titulo5")
}


def eh_nome_de_estilo_titulo_numerado(nome_estilo: str | None) -> bool:
    """True se `nome_estilo` for um dos estilos Heading 1-5 (ou seus
    equivalentes em português). Usado para "rebaixar" o estilo de um
    parágrafo que carrega um Heading indevido mas que a classificação
    concluiu não ser, de fato, um título -- senão o Sumário nativo do Word
    (que lê o nível de tópico do estilo, não a formatação visual aplicada)
    continuaria listando esse parágrafo como se fosse um capítulo."""
    if not nome_estilo:
        return False
    return nome_estilo.strip().lower() in _ESTILOS_TITULO_NUMERADO


def _estilo_word(paragraph) -> str | None:
    try:
        nome = paragraph.style.name if paragraph.style else None
    except Exception:
        nome = None
    if not nome:
        return None
    chave = nome.strip().lower()
    return _MAPA_ESTILO_WORD.get(chave)


def _eh_citacao_longa_provavel(texto: str) -> bool:
    t = texto.strip()
    if len(t) < 250:
        return False
    if t[0] in "\"“'" and t[-1] in "\"”'":
        return True
    return False


_LIMITE_PALAVRAS_TITULO = 15


def _parece_titulo(texto: str) -> bool:
    """Proteção contra estilo de título mal aplicado: em TCCs reais é comum
    um parágrafo de corpo (uma frase inteira, longa) acabar com um estilo
    Heading 1-5 aplicado por engano (colagem de outro documento, clique
    errado etc.). Isso não é uma regra ABNT -- é só uma heurística de
    segurança: título de seção é curto, então um parágrafo comprido demais
    não deve ser tratado como título mesmo que o Word diga que é um."""
    return len(texto.split()) <= _LIMITE_PALAVRAS_TITULO


def classificar_paragrafo(paragraph, estado: EstadoClassificacao) -> str:
    """Retorna a categoria do parágrafo e atualiza `estado` in place."""
    texto = paragraph.text.strip()

    if not texto:
        estado.logo_apos_titulo = False
        return "vazio"

    norm = normalizar(texto)

    # 1) Título "sem indicativo numérico" centralizado (pré-textual/listas).
    if norm in SECOES_SEM_NUMERO:
        if norm == "RESUMO":
            estado.zona = ZONA_RESUMO
        elif norm == "ABSTRACT":
            estado.zona = ZONA_ABSTRACT
        elif norm == "SUMARIO":
            estado.zona = ZONA_SUMARIO
        else:
            estado.zona = ZONA_PRE_TEXTUAL
        estado.logo_apos_titulo = True
        return "titulo_sem_numero"

    # 2) REFERÊNCIAS/APÊNDICE/ANEXO/GLOSSÁRIO: no modelo oficial usam o mesmo
    #    estilo "Heading 1" dos capítulos (entram no Sumário como titulo1).
    #    Fora da zona de SUMÁRIO -- senão uma entrada antiga em cache do
    #    sumário (ex.: "REFERÊNCIAS\t22") seria confundida com o título real.
    if estado.zona != ZONA_SUMARIO and (
        norm in SECOES_POS_TEXTUAIS_TITULO1
        or any(norm.startswith(prefixo) for prefixo in SECOES_POS_TEXTUAIS_TITULO1)
    ):
        estado.zona = ZONA_REFERENCIAS if norm.startswith("REFERENCIAS") else ZONA_POS_TEXTUAL
        estado.logo_apos_titulo = True
        return "titulo1"

    # 2b) "Figura N", "Quadro N", "Tabela N", "Fonte:" -- um padrão de texto
    #     inequívoco (nunca é o título de uma seção real) tem prioridade
    #     sobre qualquer estilo Heading que porventura esteja aplicado ao
    #     parágrafo (acontece em documentos reais: a legenda herda um estilo
    #     de título por engano ao colar de outro lugar).
    if estado.zona == ZONA_CORPO:
        if RE_LEGENDA.match(texto):
            estado.logo_apos_titulo = False
            return "legenda"
        if RE_FONTE.match(texto):
            estado.logo_apos_titulo = False
            return "fonte_ilustracao"

    # 3) Estilo do Word explícito tem prioridade sobre heurística textual,
    #    exceto para reclassificar a zona corrente -- e exceto quando o
    #    texto é longo demais para ser plausivelmente um título (proteção
    #    contra estilo de título aplicado por engano a um parágrafo comum).
    estilo_mapeado = _estilo_word(paragraph)
    if estilo_mapeado in ("titulo1", "titulo2", "titulo3", "titulo4", "titulo5") and _parece_titulo(texto):
        estado.zona = ZONA_CORPO
        estado.logo_apos_titulo = True
        return estilo_mapeado

    # 4) Título numerado digitado manualmente (sem estilo Heading aplicado).
    #    Apêndice I limita a numeração progressiva até a seção quinária
    #    (1.1.1.1.1), inclusive.
    m = RE_TITULO_NUMERADO.match(texto)
    if m and estado.zona in (ZONA_PRE_TEXTUAL, ZONA_CORPO):
        prefixo = m.group("prefixo")
        nivel = min(prefixo.count("."), 4) + 1  # 1 a 5
        estado.zona = ZONA_CORPO
        estado.logo_apos_titulo = True
        return f"titulo{nivel}"

    if estilo_mapeado == "legenda":
        estado.logo_apos_titulo = False
        return "legenda"
    if estilo_mapeado == "fonte_ilustracao":
        estado.logo_apos_titulo = False
        return "fonte_ilustracao"

    # 5) Conteúdo dentro de zonas conhecidas.
    if estado.zona == ZONA_REFERENCIAS:
        estado.logo_apos_titulo = False
        return "referencia"

    if estado.zona in (ZONA_RESUMO, ZONA_ABSTRACT):
        estado.logo_apos_titulo = False
        return "resumo_corpo"

    if estado.zona == ZONA_SUMARIO:
        if RE_LINHA_SUMARIO.search(texto):
            # ainda parece uma entrada de sumário (termina em nº de página)
            # -- será removida e substituída pelo campo TOC nativo.
            estado.logo_apos_titulo = False
            return "sumario_entrada"
        # não parece mais uma entrada de sumário: o bloco acabou (ex.: já
        # chegamos ao primeiro capítulo, digitado sem estilo de título).
        estado.zona = ZONA_PRE_TEXTUAL
        return classificar_paragrafo(paragraph, estado)

    if estado.zona == ZONA_CORPO:
        # legenda/fonte já foram tratadas no passo 2b, antes do estilo Word
        if _eh_citacao_longa_provavel(texto):
            estado.logo_apos_titulo = False
            return "citacao_longa"
        categoria = "corpo_sem_recuo" if estado.logo_apos_titulo else "corpo"
        estado.logo_apos_titulo = False
        return categoria

    # 6) Fora de qualquer zona reconhecida (capa, folha de rosto, ficha
    #    catalográfica, dedicatória, agradecimentos, epígrafe, listas...).
    estado.logo_apos_titulo = False
    return "outro"
