"""Aplica um EstiloParagrafo (config.py) a um parágrafo python-docx.

Importante: estas funções só tocam em propriedades de formatação (fonte,
tamanho, negrito/itálico, alinhamento, espaçamento, recuo). O texto dos
`run`s nunca é lido para decisão nem reescrito -- inclusive "maiúsculas" é
aplicado via o efeito de caractere w:caps (font.all_caps), que renderiza o
texto em caixa alta sem alterar a string armazenada no documento.
"""
from __future__ import annotations

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.text.run import Run

from .config import EstiloParagrafo


def runs_completos(paragraph) -> list[Run]:
    """Todos os `w:r` do parágrafo, incluindo os que estão dentro de um
    `w:hyperlink` -- comum em referências cruzadas automáticas (ex.: o
    número de página numa Lista de Figuras/Tabelas gerada via Inserir >
    Referência Cruzada). `paragraph.runs` do python-docx só enxerga os
    `w:r` soltos direto dentro do parágrafo, então um hyperlink inteiro
    ficava sem receber nenhuma normalização de fonte -- mesmo depois de
    "formatado", esses trechos continuavam na fonte original do aluno."""
    runs: list[Run] = []
    for filho in paragraph._p:
        if filho.tag == qn("w:r"):
            runs.append(Run(filho, paragraph))
        elif filho.tag == qn("w:hyperlink"):
            for r_el in filho.findall(qn("w:r")):
                runs.append(Run(r_el, paragraph))
    return runs


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

    for run in runs_completos(paragraph):
        if forcar_fonte:
            run.font.name = estilo.fonte
            run.font.size = Pt(estilo.tamanho_pt)
        if forcar_negrito_italico:
            if estilo.negrito is not None:
                run.font.bold = estilo.negrito
            if estilo.italico is not None:
                run.font.italic = estilo.italico
        if forcar_maiusculas:
            run.font.all_caps = estilo.maiusculas


def aplicar_negrito_prefixo(paragraph, prefixo: str) -> None:
    """Deixa em negrito só o prefixo literal do parágrafo (ex.: "Fonte:"),
    sem alterar o texto nem o negrito do restante da linha -- é o padrão
    observado nos exemplos reais do modelo oficial ("**Fonte:** Autor,
    ano.", só a palavra "Fonte:" em negrito).
    """
    texto = paragraph.text
    if not texto.lower().startswith(prefixo.lower()):
        return

    restante = len(prefixo)
    for run in list(paragraph.runs):
        if restante <= 0:
            break
        n = len(run.text)
        if n == 0:
            continue
        if n <= restante:
            run.font.bold = True
            restante -= n
            continue
        # o prefixo termina no meio deste run: separa em dois runs.
        texto_run = run.text
        parte_prefixo = texto_run[:restante]
        parte_resto = texto_run[restante:]
        novo = OxmlElement("w:r")
        rpr_original = run._r.find(qn("w:rPr"))
        if rpr_original is not None:
            novo.append(_clonar_elemento(rpr_original))
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = parte_prefixo
        novo.append(t)
        run._r.addprevious(novo)
        novo_run = type(run)(novo, run._parent)
        novo_run.font.bold = True

        run.text = parte_resto
        restante = 0


def _clonar_elemento(el):
    import copy
    return copy.deepcopy(el)


def aplicar_configuracao_pagina(document, page_setup) -> None:
    for section in document.sections:
        section.page_width = Cm(page_setup.largura_cm)
        section.page_height = Cm(page_setup.altura_cm)
        section.top_margin = Cm(page_setup.margem_superior_cm)
        section.bottom_margin = Cm(page_setup.margem_inferior_cm)
        section.left_margin = Cm(page_setup.margem_esquerda_cm)
        section.right_margin = Cm(page_setup.margem_direita_cm)
