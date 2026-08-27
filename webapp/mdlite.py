"""Conversor mínimo de Markdown -> HTML, só para o formato que
`formatador_tcc.relatorio.gerar_relatorio_markdown` produz (títulos #/##/###,
listas com "- ", negrito **assim**, itálico _assim_). Não é um parser de
Markdown genérico -- existe só para não precisar de mais uma dependência
externa para renderizar o relatório na página de resultado.

Todo o texto é escapado com html.escape ANTES de interpretar a marcação,
porque o relatório contém trechos copiados do próprio documento do aluno
(títulos de referência, trechos de citação) -- texto não confiável que não
pode ir para o HTML sem escapar.
"""
from __future__ import annotations

import html
import re

_RE_NEGRITO = re.compile(r"\*\*(.+?)\*\*")
_RE_ITALICO = re.compile(r"(?<!\w)_(.+?)_(?!\w)")


def _inline(texto: str) -> str:
    texto = _RE_NEGRITO.sub(r"<strong>\1</strong>", texto)
    texto = _RE_ITALICO.sub(r"<em>\1</em>", texto)
    return texto


def markdown_para_html(markdown: str) -> str:
    linhas = markdown.splitlines()
    html_partes: list[str] = []
    em_lista = False

    def fechar_lista():
        nonlocal em_lista
        if em_lista:
            html_partes.append("</ul>")
            em_lista = False

    for linha_bruta in linhas:
        linha = html.escape(linha_bruta.rstrip())

        if not linha.strip():
            fechar_lista()
            continue

        if linha.startswith("### "):
            fechar_lista()
            html_partes.append(f"<h3>{_inline(linha[4:])}</h3>")
        elif linha.startswith("## "):
            fechar_lista()
            html_partes.append(f"<h2>{_inline(linha[3:])}</h2>")
        elif linha.startswith("# "):
            fechar_lista()
            html_partes.append(f"<h1>{_inline(linha[2:])}</h1>")
        elif linha.startswith("- "):
            if not em_lista:
                html_partes.append("<ul>")
                em_lista = True
            html_partes.append(f"<li>{_inline(linha[2:])}</li>")
        else:
            fechar_lista()
            html_partes.append(f"<p>{_inline(linha)}</p>")

    fechar_lista()
    return "\n".join(html_partes)
