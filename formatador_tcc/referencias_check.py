"""Checagem heurística de formatação das Referências -- NBR 6023.

Roda sobre o documento ORIGINAL do aluno (antes do formatador.py normalizar
alinhamento/espaçamento/recuo), para que o relatório mostre o que estava
errado e o que foi corrigido automaticamente. Problemas de conteúdo (autor
sem maiúsculas, ano ausente, sem ponto final, destaque de título
inconsistente) não são corrigidos automaticamente -- exigem revisão do
aluno, porque corrigi-los exigiria reescrever o texto da referência.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from .classify import EstadoClassificacao, classificar_paragrafo
from .texto import RE_ANO

RE_AUTOR_MAIUSCULO = re.compile(r"^[A-ZÀ-Ú][A-ZÀ-Ú\.\-' ]*[,\.]")


@dataclass
class ProblemaReferencia:
    paragrafo_indice: int
    texto: str
    corrigivel_automaticamente: bool
    mensagem: str


@dataclass
class RelatorioReferencias:
    problemas: list[ProblemaReferencia] = field(default_factory=list)
    total_entradas: int = 0


def _resolver_propriedade(paragraph, nome: str):
    """Resolve uma propriedade de paragraph_format subindo a cadeia de
    estilos (Word só aplica o valor direto do parágrafo quando ele existe;
    caso contrário, herda do estilo, que pode herdar de outro estilo)."""
    valor = getattr(paragraph.paragraph_format, nome)
    if valor is not None:
        return valor
    estilo = paragraph.style
    visitados = set()
    while estilo is not None and id(estilo) not in visitados:
        visitados.add(id(estilo))
        valor = getattr(estilo.paragraph_format, nome, None)
        if valor is not None:
            return valor
        estilo = getattr(estilo, "base_style", None)
    return None


def _checar_entrada(indice: int, paragraph) -> list[ProblemaReferencia]:
    texto = paragraph.text.strip()
    problemas: list[ProblemaReferencia] = []

    if not RE_AUTOR_MAIUSCULO.match(texto):
        problemas.append(ProblemaReferencia(
            indice, texto, False,
            "o elemento de entrada (autor ou entidade) não parece estar em "
            "MAIÚSCULAS seguido de vírgula, como pede a NBR 6023 "
            "(ex.: 'SILVA, João.')."
        ))

    if not RE_ANO.search(texto):
        problemas.append(ProblemaReferencia(
            indice, texto, False,
            "não foi encontrado um ano de publicação (4 dígitos) nesta referência."
        ))

    if not texto.rstrip().endswith("."):
        problemas.append(ProblemaReferencia(
            indice, texto, False,
            "a referência não termina com ponto final."
        ))

    alinhamento = _resolver_propriedade(paragraph, "alignment")
    if alinhamento is not None and str(alinhamento) != "LEFT (0)":
        problemas.append(ProblemaReferencia(
            indice, texto, True,
            "alinhamento não está à esquerda (referências não devem ser "
            "justificadas, para preservar os espaços entre os elementos)."
        ))

    recuo = _resolver_propriedade(paragraph, "first_line_indent")
    if recuo is not None and recuo.cm > 0.05:
        problemas.append(ProblemaReferencia(
            indice, texto, True,
            "possui recuo de primeira linha (referências não levam recuo)."
        ))

    espacamento = _resolver_propriedade(paragraph, "line_spacing")
    if espacamento not in (None, 1, 1.0):
        problemas.append(ProblemaReferencia(
            indice, texto, True,
            "espaçamento entre linhas dentro da referência não é simples."
        ))

    return problemas


def _destaque_do_titulo(paragraph) -> str:
    tem_negrito = any(r.font.bold for r in paragraph.runs if r.text.strip())
    tem_italico = any(r.font.italic for r in paragraph.runs if r.text.strip())
    if tem_negrito and tem_italico:
        return "misto"
    if tem_negrito:
        return "negrito"
    if tem_italico:
        return "italico"
    return "nenhum"


def verificar_referencias(document) -> RelatorioReferencias:
    estado = EstadoClassificacao()
    relatorio = RelatorioReferencias()
    destaques: list[str] = []

    for i, paragraph in enumerate(document.paragraphs):
        categoria = classificar_paragrafo(paragraph, estado)
        if categoria != "referencia" or not paragraph.text.strip():
            continue
        relatorio.total_entradas += 1
        relatorio.problemas.extend(_checar_entrada(i, paragraph))
        destaques.append(_destaque_do_titulo(paragraph))

    rotulos_destaque = {"negrito": "negrito", "italico": "itálico", "misto": "negrito+itálico numa mesma entrada"}
    usados = {d for d in destaques if d != "nenhum"}
    if len(usados) > 1:
        descricao = ", ".join(rotulos_destaque[d] for d in sorted(usados))
        relatorio.problemas.append(ProblemaReferencia(
            -1, "(lista de referências)", False,
            "o destaque tipográfico do título das obras não é consistente: "
            f"a lista mistura {descricao} entre as entradas. "
            "A NBR 6023 exige um único padrão (negrito OU itálico) para "
            "todas as referências do trabalho."
        ))
    elif relatorio.total_entradas > 0 and not usados:
        relatorio.problemas.append(ProblemaReferencia(
            -1, "(lista de referências)", False,
            "nenhuma referência da lista usa negrito ou itálico para destacar "
            "o título das obras -- confira se isso é exigido para os tipos de "
            "fonte utilizados (livros, monografias, artigos de periódico)."
        ))

    return relatorio
