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
    r"^(?P<prefixo>\d{1,2}(?:\.\d{1,2}){0,3})\.?\s+[A-ZÀ-Ú0-9]"
)
RE_LEGENDA = re.compile(
    r"^(figura|quadro|tabela|gr[aá]fico|equa[cç][aã]o)\s*\d+", re.IGNORECASE
)
RE_FONTE = re.compile(r"^fonte\s*:", re.IGNORECASE)
# Linha de sumário: termina em número de página (com ou sem tab/pontos de
# preenchimento antes), ex.: "1\tINTRODUÇÃO\t12" ou "REFERÊNCIAS  22".
RE_LINHA_SUMARIO = re.compile(r"\d{1,4}\s*$")

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
    if norm in SECOES_POS_TEXTUAIS_TITULO1 or any(
        norm.startswith(prefixo) for prefixo in SECOES_POS_TEXTUAIS_TITULO1
    ):
        estado.zona = ZONA_REFERENCIAS if norm.startswith("REFERENCIAS") else ZONA_POS_TEXTUAL
        estado.logo_apos_titulo = True
        return "titulo1"

    # 2) Estilo do Word explícito tem prioridade sobre heurística textual,
    #    exceto para reclassificar a zona corrente.
    estilo_mapeado = _estilo_word(paragraph)
    if estilo_mapeado in ("titulo1", "titulo2", "titulo3"):
        estado.zona = ZONA_CORPO
        estado.logo_apos_titulo = True
        return estilo_mapeado

    # 3) Título numerado digitado manualmente (sem estilo Heading aplicado).
    m = RE_TITULO_NUMERADO.match(texto)
    if m and estado.zona in (ZONA_PRE_TEXTUAL, ZONA_CORPO):
        prefixo = m.group("prefixo")
        nivel = min(prefixo.count("."), 2) + 1  # 1, 2 ou 3
        estado.zona = ZONA_CORPO
        estado.logo_apos_titulo = True
        return f"titulo{nivel}"

    if estilo_mapeado == "legenda":
        estado.logo_apos_titulo = False
        return "legenda"
    if estilo_mapeado == "fonte_ilustracao":
        estado.logo_apos_titulo = False
        return "fonte_ilustracao"

    # 4) Conteúdo dentro de zonas conhecidas.
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
        if RE_LEGENDA.match(texto):
            estado.logo_apos_titulo = False
            return "legenda"
        if RE_FONTE.match(texto):
            estado.logo_apos_titulo = False
            return "fonte_ilustracao"
        if _eh_citacao_longa_provavel(texto):
            estado.logo_apos_titulo = False
            return "citacao_longa"
        categoria = "corpo_sem_recuo" if estado.logo_apos_titulo else "corpo"
        estado.logo_apos_titulo = False
        return categoria

    # 5) Fora de qualquer zona reconhecida (capa, folha de rosto, ficha
    #    catalográfica, dedicatória, agradecimentos, epígrafe, listas...).
    estado.logo_apos_titulo = False
    return "outro"
