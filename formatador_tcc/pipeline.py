"""Orquestra o processamento completo de um documento, do jeito que tanto o
CLI (`cli.py`) quanto a interface web (`webapp/app.py`) precisam: roda as
checagens sobre o documento original, formata, reconstrói o sumário e monta
o relatório -- tudo em memória, sem exigir nada em disco além do que o
chamador decidir salvar.
"""
from __future__ import annotations

import io
from dataclasses import dataclass

import docx

from .citacoes import cruzar_citacoes_e_referencias
from .contagem import verificar_contagem_resumo
from .formatador import formatar_documento
from .paginacao import aplicar_numeracao_paginas
from .referencias_check import verificar_referencias
from .relatorio import gerar_relatorio_markdown
from .sumario import reconstruir_sumario


@dataclass
class ResultadoProcessamento:
    docx_bytes: bytes
    relatorio_markdown: str


def processar_docx_bytes(conteudo: bytes, *, nome_entrada: str, nome_saida: str) -> ResultadoProcessamento:
    """Recebe os bytes de um .docx do aluno e devolve o .docx formatado +
    o relatório em Markdown, sem tocar no disco."""
    document = docx.Document(io.BytesIO(conteudo))

    relatorio_referencias = verificar_referencias(document)
    relatorio_contagem = verificar_contagem_resumo(document)
    resultado_citacoes = cruzar_citacoes_e_referencias(document)

    eventos = formatar_documento(document)
    resultado_sumario = reconstruir_sumario(document)
    resultado_paginacao = aplicar_numeracao_paginas(document)

    buffer = io.BytesIO()
    document.save(buffer)

    relatorio_md = gerar_relatorio_markdown(
        nome_entrada=nome_entrada,
        nome_saida=nome_saida,
        eventos_formatacao=eventos,
        resultado_sumario=resultado_sumario,
        resultado_paginacao=resultado_paginacao,
        resultado_citacoes=resultado_citacoes,
        relatorio_referencias=relatorio_referencias,
        relatorio_contagem=relatorio_contagem,
    )

    return ResultadoProcessamento(docx_bytes=buffer.getvalue(), relatorio_markdown=relatorio_md)
