"""Numeração de página -- Apêndice I, seção PAGINAÇÃO:

"Todas as folhas do trabalho, a partir da folha de rosto, devem ser
contadas sequencialmente, mas não numeradas. A numeração deve ser colocada
a partir da primeira folha da parte textual (Introdução), em algarismos
arábicos, no canto superior direito da folha. Havendo anexo(s) e
apêndice(s), as suas folhas devem ser numeradas e paginadas de maneira
contínua."

O comentário da própria autora do modelo oficial detalha o cálculo: contam-se
todas as páginas físicas anteriores exceto a CAPA e a FICHA CATALOGRÁFICA --
por isso o número mostrado na Introdução é sempre "página física atual menos
2" (config.PAGINAS_EXCLUIDAS_DA_CONTAGEM).

Assim como o Sumário, não há como calcular a paginação real sem um motor de
layout de verdade (python-docx não faz isso). A solução é a mesma: um campo
nativo do Word -- aqui um campo de fórmula `{ =PAGE-2 }` --, que o próprio
Word recalcula automaticamente ao abrir o documento (settings.xml já pede
atualização de campos ao abrir, configurado por sumario.py).
"""
from __future__ import annotations

from dataclasses import dataclass

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .classify import EstadoClassificacao, classificar_paragrafo
from .config import PAGINAS_EXCLUIDAS_DA_CONTAGEM


@dataclass
class ResultadoPaginacao:
    aplicada: bool
    secao_introducao: int | None = None


def _mapear_paragrafo_para_secao(document) -> list[int]:
    """Para cada parágrafo de nível superior (mesma ordem de
    `document.paragraphs`), a que seção (`document.sections`) ele pertence."""
    body = document.element.body
    mapa: list[int] = []
    secao_atual = 0
    for child in body:
        if child.tag == qn("w:p"):
            mapa.append(secao_atual)
            ppr = child.find(qn("w:pPr"))
            if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
                secao_atual += 1
    return mapa


def _indice_primeiro_titulo1(document) -> int | None:
    estado = EstadoClassificacao()
    for i, paragraph in enumerate(document.paragraphs):
        categoria = classificar_paragrafo(paragraph, estado)
        if categoria == "titulo1":
            return i
    return None


def _construir_campo_numero_pagina(paragraph) -> None:
    """Substitui o conteúdo do parágrafo por um campo `{ =PAGE-N }`."""
    p = paragraph._p
    for child in list(p.findall(qn("w:r"))):
        p.remove(child)

    def novo_run():
        r = OxmlElement("w:r")
        p.append(r)
        return r

    r1 = novo_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")
    r1.append(begin)

    r2 = novo_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" =PAGE-{PAGINAS_EXCLUIDAS_DA_CONTAGEM} "
    r2.append(instr)

    r3 = novo_run()
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    r3.append(sep)

    r4 = novo_run()
    t = OxmlElement("w:t")
    t.text = "1"
    r4.append(t)

    r5 = novo_run()
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    r5.append(end)


def aplicar_numeracao_paginas(document) -> ResultadoPaginacao:
    indice_intro = _indice_primeiro_titulo1(document)
    if indice_intro is None:
        return ResultadoPaginacao(aplicada=False)

    mapa_secoes = _mapear_paragrafo_para_secao(document)
    if indice_intro >= len(mapa_secoes):
        return ResultadoPaginacao(aplicada=False)
    secao_idx = mapa_secoes[indice_intro]

    secoes = document.sections
    if secao_idx >= len(secoes):
        return ResultadoPaginacao(aplicada=False)

    secao = secoes[secao_idx]
    secao.header.is_linked_to_previous = False
    header = secao.header

    # remove qualquer conteúdo pré-existente no cabeçalho -- inclusive
    # blocos de "Número de página" prontos do Word (`w:sdt`), que não
    # aparecem em header.paragraphs mas ficariam duplicando o número.
    for child in list(header._element):
        header._element.remove(child)
    paragrafo = header.add_paragraph()

    paragrafo.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _construir_campo_numero_pagina(paragrafo)

    # Seções seguintes (Referências, Apêndices, Anexos) herdam este mesmo
    # cabeçalho -- numeração contínua, conforme o Apêndice I.
    for s in secoes[secao_idx + 1:]:
        s.header.is_linked_to_previous = True

    return ResultadoPaginacao(aplicada=True, secao_introducao=secao_idx)
