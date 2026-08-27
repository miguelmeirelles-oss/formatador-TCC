"""Confere a contagem de palavras do Resumo e do Abstract.

Regra extraída literalmente do "Modelo de TCC oficial": "O resumo deve
conter no mínimo 150 palavras e no máximo 500 palavras" -- e o mesmo vale
para o abstract, por ser a tradução do resumo.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

RE_PALAVRAS_CHAVE = re.compile(r"^(palavras[\s-]?chave|keywords)\s*:", re.IGNORECASE)

from .classify import EstadoClassificacao, ZONA_RESUMO, ZONA_ABSTRACT, classificar_paragrafo
from .config import RESUMO_MIN_PALAVRAS, RESUMO_MAX_PALAVRAS
from .texto import contar_palavras


@dataclass
class ContagemSecao:
    nome: str
    encontrada: bool
    palavras: int = 0
    dentro_do_limite: bool = True
    mensagem: str = ""


@dataclass
class RelatorioContagem:
    resumo: ContagemSecao
    abstract: ContagemSecao


def _avaliar(nome: str, palavras: int, encontrada: bool) -> ContagemSecao:
    if not encontrada:
        return ContagemSecao(nome, False, 0, False, f"{nome} não foi encontrado no documento.")
    if palavras < RESUMO_MIN_PALAVRAS:
        return ContagemSecao(
            nome, True, palavras, False,
            f"{nome} tem {palavras} palavras -- abaixo do mínimo de {RESUMO_MIN_PALAVRAS}.",
        )
    if palavras > RESUMO_MAX_PALAVRAS:
        return ContagemSecao(
            nome, True, palavras, False,
            f"{nome} tem {palavras} palavras -- acima do máximo de {RESUMO_MAX_PALAVRAS}.",
        )
    return ContagemSecao(nome, True, palavras, True, f"{nome} tem {palavras} palavras -- dentro do limite.")


def verificar_contagem_resumo(document) -> RelatorioContagem:
    estado = EstadoClassificacao()
    palavras_resumo = 0
    palavras_abstract = 0
    encontrou_resumo = False
    encontrou_abstract = False

    for paragraph in document.paragraphs:
        categoria = classificar_paragrafo(paragraph, estado)
        if categoria == "titulo_sem_numero":
            if estado.zona == ZONA_RESUMO:
                encontrou_resumo = True
            elif estado.zona == ZONA_ABSTRACT:
                encontrou_abstract = True
            continue
        if categoria != "resumo_corpo":
            continue
        texto = paragraph.text.strip()
        if RE_PALAVRAS_CHAVE.match(texto):
            continue
        if estado.zona == ZONA_RESUMO:
            palavras_resumo += contar_palavras(texto)
        elif estado.zona == ZONA_ABSTRACT:
            palavras_abstract += contar_palavras(texto)

    return RelatorioContagem(
        resumo=_avaliar("Resumo", palavras_resumo, encontrou_resumo),
        abstract=_avaliar("Abstract", palavras_abstract, encontrou_abstract),
    )
