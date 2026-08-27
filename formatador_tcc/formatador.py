"""Motor de normalização de formatação ABNT.

`formatar_documento` recebe um documento python-docx já aberto (o TCC do
aluno) e aplica a formatação ABNT do modelo oficial em cada parágrafo do
corpo do documento -- fonte, tamanho, alinhamento, espaçamento, recuo -- sem
jamais reescrever o texto dos parágrafos. Retorna uma lista de eventos
(auditoria do que foi tocado) para entrar no relatório final.
"""
from __future__ import annotations

from dataclasses import dataclass, replace
from typing import Optional

from docx.shared import Pt

from . import config
from .classify import EstadoClassificacao, classificar_paragrafo
from .estilos import aplicar_configuracao_pagina, aplicar_estilo

_CORPO_SEM_RECUO = replace(config.CORPO_TEXTO, recuo_primeira_linha_cm=0.0)

# categoria -> (estilo, forcar_negrito_italico, forcar_maiusculas, estilo_word)
_REGRAS = {
    "titulo1": (config.TITULO_SECAO_PRIMARIA, True, True, "Heading 1"),
    "titulo2": (config.TITULO_SECAO_SECUNDARIA, True, False, "Heading 2"),
    "titulo3": (config.TITULO_SECAO_TERCIARIA, True, False, "Heading 3"),
    "titulo_sem_numero": (config.TITULO_SEM_NUMERO, True, True, "Título de Seção"),
    "legenda": (config.LEGENDA, True, False, None),
    "fonte_ilustracao": (config.FONTE_ILUSTRACAO, True, False, None),
    "citacao_longa": (config.CITACAO_LONGA, False, False, None),
    "referencia": (config.REFERENCIA, False, False, None),
    "resumo_corpo": (config.CORPO_TEXTO, False, False, None),
    "corpo": (config.CORPO_TEXTO, False, False, None),
    "corpo_sem_recuo": (_CORPO_SEM_RECUO, False, False, None),
}


@dataclass
class EventoFormatacao:
    indice: int
    categoria: str
    trecho: str


def _normalizar_apenas_fonte(paragraph) -> None:
    for run in paragraph.runs:
        run.font.name = config.CORPO_TEXTO.fonte
        run.font.size = Pt(config.CORPO_TEXTO.tamanho_pt)


def _tentar_aplicar_estilo_word(paragraph, nome_estilo: Optional[str]) -> None:
    if not nome_estilo:
        return
    try:
        paragraph.style = paragraph.part.document.styles[nome_estilo]
    except KeyError:
        # o documento do aluno não tem esse estilo definido -- a formatação
        # direta aplicada em seguida garante a aparência correta mesmo assim.
        pass


def formatar_documento(document) -> list[EventoFormatacao]:
    aplicar_configuracao_pagina(document, config.PAGE_SETUP)

    estado = EstadoClassificacao()
    eventos: list[EventoFormatacao] = []

    for indice, paragraph in enumerate(document.paragraphs):
        categoria = classificar_paragrafo(paragraph, estado)

        if categoria in ("vazio", "sumario_entrada"):
            continue

        if categoria == "outro":
            _normalizar_apenas_fonte(paragraph)
            continue

        regra = _REGRAS.get(categoria)
        if regra is None:
            _normalizar_apenas_fonte(paragraph)
            continue

        estilo, forcar_negrito_italico, forcar_maiusculas, nome_estilo_word = regra
        _tentar_aplicar_estilo_word(paragraph, nome_estilo_word)
        aplicar_estilo(
            paragraph,
            estilo,
            forcar_fonte=True,
            forcar_negrito_italico=forcar_negrito_italico,
            forcar_maiusculas=forcar_maiusculas,
        )

        trecho = paragraph.text.strip()[:60]
        eventos.append(EventoFormatacao(indice=indice, categoria=categoria, trecho=trecho))

    return eventos
