"""Aplica um EstiloParagrafo (config.py) a um parágrafo python-docx.

Importante: estas funções só tocam em propriedades de formatação (fonte,
tamanho, negrito/itálico, alinhamento, espaçamento, recuo). O texto dos
`run`s nunca é lido para decisão nem reescrito -- inclusive "maiúsculas" é
aplicado via o efeito de caractere w:caps (font.all_caps), que renderiza o
texto em caixa alta sem alterar a string armazenada no documento.
"""
from __future__ import annotations

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

from .config import EstiloParagrafo

_ALINHAMENTO = {
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
}


def aplicar_estilo(
    paragraph,
    estilo: EstiloParagrafo,
    *,
    forcar_fonte: bool = True,
    forcar_negrito_italico: bool = False,
    forcar_maiusculas: bool = False,
) -> None:
    pf = paragraph.paragraph_format
    pf.alignment = _ALINHAMENTO[estilo.alinhamento]
    pf.line_spacing = estilo.espacamento_linha
    pf.space_before = Pt(estilo.espaco_antes_pt)
    pf.space_after = Pt(estilo.espaco_depois_pt)
    pf.first_line_indent = Cm(estilo.recuo_primeira_linha_cm) if estilo.recuo_primeira_linha_cm else Cm(0)
    pf.left_indent = Cm(estilo.recuo_esquerdo_cm) if estilo.recuo_esquerdo_cm else None

    for run in paragraph.runs:
        if forcar_fonte:
            run.font.name = estilo.fonte
            run.font.size = Pt(estilo.tamanho_pt)
        if forcar_negrito_italico:
            if estilo.negrito is not None:
                run.font.bold = estilo.negrito
            if estilo.italico is not None:
                run.font.italic = estilo.italico
        if forcar_maiusculas and estilo.maiusculas:
            run.font.all_caps = True


def aplicar_configuracao_pagina(document, page_setup) -> None:
    for section in document.sections:
        section.page_width = Cm(page_setup.largura_cm)
        section.page_height = Cm(page_setup.altura_cm)
        section.top_margin = Cm(page_setup.margem_superior_cm)
        section.bottom_margin = Cm(page_setup.margem_inferior_cm)
        section.left_margin = Cm(page_setup.margem_esquerda_cm)
        section.right_margin = Cm(page_setup.margem_direita_cm)
