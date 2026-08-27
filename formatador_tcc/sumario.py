"""Reconstrói o SUMÁRIO como um campo TOC nativo do Word.

Em vez de tentar calcular números de página manualmente (o que exigiria um
motor de paginação real, algo que o python-docx não tem), inserimos o mesmo
campo `{ TOC }` que o Word usa nativamente: ele é recalculado pelo próprio
Word ao abrir o documento (via `updateFields`) ou com F9, sempre com os
números de página corretos, a partir dos títulos que o formatador marcou
como "Heading 1/2/3" (ver formatador.py). As entradas antigas do sumário
(texto colado ou desatualizado) são removidas.
"""
from __future__ import annotations

from dataclasses import dataclass

from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from .classify import EstadoClassificacao, classificar_paragrafo
from .texto import normalizar


@dataclass
class ResultadoSumario:
    encontrado: bool
    entradas_removidas: int = 0
    campo_inserido: bool = False


def _localizar_bloco_sumario(document):
    estado = EstadoClassificacao()
    indice_titulo = None
    indices_entradas: list[int] = []
    for i, paragraph in enumerate(document.paragraphs):
        categoria = classificar_paragrafo(paragraph, estado)
        if categoria == "titulo_sem_numero" and normalizar(paragraph.text) == "SUMARIO":
            indice_titulo = i
            continue
        if indice_titulo is not None and categoria == "sumario_entrada":
            indices_entradas.append(i)
        elif indice_titulo is not None and categoria != "sumario_entrada" and indices_entradas:
            break
    return indice_titulo, indices_entradas


def _forcar_atualizacao_de_campos_ao_abrir(document) -> None:
    settings = document.settings.element
    if settings.find(qn("w:updateFields")) is not None:
        return
    el = OxmlElement("w:updateFields")
    el.set(qn("w:val"), "true")
    settings.append(el)


def _construir_campo_toc(paragraph) -> None:
    """Substitui o conteúdo do parágrafo por um campo `{ TOC \\o "1-3" \\h \\z \\u }`."""
    p = paragraph._p
    for child in list(p.findall(qn("w:r"))):
        p.remove(child)
    for child in list(p.findall(qn("w:hyperlink"))):
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
    instr.text = ' TOC \\o "1-3" \\h \\z \\u '
    r2.append(instr)

    r3 = novo_run()
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    r3.append(sep)

    r4 = novo_run()
    t = OxmlElement("w:t")
    t.text = "Sumário gerado automaticamente -- clique com o botão direito e escolha “Atualizar campo” (ou F9) após abrir no Word."
    r4.append(t)

    r5 = novo_run()
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    r5.append(end)


def reconstruir_sumario(document) -> ResultadoSumario:
    indice_titulo, indices_entradas = _localizar_bloco_sumario(document)
    if indice_titulo is None:
        return ResultadoSumario(encontrado=False)

    paragrafos = document.paragraphs
    titulo = paragrafos[indice_titulo]

    novo_p = OxmlElement("w:p")
    titulo._p.addnext(novo_p)
    from docx.text.paragraph import Paragraph
    campo_paragraph = Paragraph(novo_p, titulo._parent)
    _construir_campo_toc(campo_paragraph)

    removidos = 0
    for i in indices_entradas:
        alvo = paragrafos[i]._p
        alvo.getparent().remove(alvo)
        removidos += 1

    _forcar_atualizacao_de_campos_ao_abrir(document)

    return ResultadoSumario(encontrado=True, entradas_removidas=removidos, campo_inserido=True)
