"""Formata a Capa e a Folha de Rosto (Apêndice I).

Diferente dos títulos de seção, o texto da Capa/Folha de Rosto (nome do
autor, título do trabalho, instituição...) é digitado livremente pelo aluno
-- não tem nenhum estilo Word nem padrão de texto que diga "isto é o nome
do autor". Adivinhar isso às cegas seria inventar uma regra.

Mas o Apêndice I define a ORDEM exata dos elementos em cada folha, e dois
desses elementos são textos fixos, iguais para qualquer aluno (não são
digitados livremente):

- "TRABALHO DE CONCLUSÃO DE CURSO" / "MONOGRAFIA" / "DISSERTAÇÃO" sozinho
  numa linha -- o "tipo do documento" da Capa.
- Uma frase que começa com "Trabalho de Conclusão de Curso apresentado..."
  (ou "Monografia apresentada...", "Dissertação apresentada...") -- a
  "natureza do trabalho" da Folha de Rosto.

A partir desses dois pontos fixos, a posição (não o conteúdo) dos blocos de
texto vizinhos identifica os outros elementos, na ordem que o Apêndice I
descreve:

  CAPA:         [cabeçalho?] AUTOR TÍTULO **TIPO** (local) (ano)
  FOLHA DE ROSTO:            AUTOR TÍTULO **NATUREZA** (orientador...)

(** = âncora de texto fixo; os demais são só "os dois blocos de texto
anteriores a essa âncora").
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from . import config
from .estilos import aplicar_estilo
from .texto import normalizar

RE_INSTITUICAO = re.compile(
    r"CENTRO FEDERAL DE EDUCA[CÇ][AÃ]O TECNOL[OÓ]GICA", re.IGNORECASE
)
RE_NATUREZA_TRABALHO = re.compile(
    r"^(trabalho de conclus[aã]o de curso|monografia|disserta[cç][aã]o|projeto final)\s+apresentad[oa]",
    re.IGNORECASE,
)
# Procurado em qualquer posição da linha (não só no início) -- documentos
# reais variam entre "Orientador: Nome" e "Prof. Orientador: Nome".
RE_ORIENTADOR = re.compile(r"orientador", re.IGNORECASE)


@dataclass
class ResultadoCapa:
    titulos_formatados: int = 0
    autores_formatados: int = 0
    tipos_documento_formatados: int = 0
    naturezas_formatadas: int = 0
    orientadores_formatados: int = 0
    institucoes_formatadas: int = 0


def _blocos(paragraphs, inicio: int, fim: int) -> list[list[int]]:
    """Agrupa paragraphs[inicio:fim] em blocos de índices consecutivos com
    texto (um "bloco" é o que fica entre duas linhas em branco)."""
    blocos: list[list[int]] = []
    atual: list[int] = []
    for i in range(inicio, fim):
        if paragraphs[i].text.strip():
            atual.append(i)
        elif atual:
            blocos.append(atual)
            atual = []
    if atual:
        blocos.append(atual)
    return blocos


def _aplicar(paragraphs, indices: list[int], estilo) -> None:
    for i in indices:
        aplicar_estilo(
            paragraphs[i],
            estilo,
            forcar_fonte=True,
            forcar_negrito_italico=True,
            forcar_maiusculas=True,
        )


def _indice_fim_zona_pre_capa(document) -> int:
    """Até onde vai a região de Capa/Folha de Rosto/Ficha Catalográfica --
    o primeiro título reconhecido (Dedicatória, Resumo, Sumário, um
    capítulo...) marca o fim dessa região."""
    from .classify import EstadoClassificacao, classificar_paragrafo

    estado = EstadoClassificacao()
    for i, paragraph in enumerate(document.paragraphs):
        categoria = classificar_paragrafo(paragraph, estado)
        if categoria not in ("outro", "vazio"):
            return i
    return len(document.paragraphs)


def formatar_capa_e_folha_de_rosto(document) -> ResultadoCapa:
    paragraphs = document.paragraphs
    fim = _indice_fim_zona_pre_capa(document)
    blocos = _blocos(paragraphs, 0, fim)
    resultado = ResultadoCapa()

    for idx, bloco in enumerate(blocos):
        texto_normalizado = normalizar(paragraphs[bloco[0]].text)

        if len(bloco) == 1 and texto_normalizado in config.TIPOS_DE_DOCUMENTO:
            _aplicar(paragraphs, bloco, config.ELEMENTO_CAPA)
            resultado.tipos_documento_formatados += 1
            if idx >= 1:
                _aplicar(paragraphs, blocos[idx - 1], config.TITULO_TRABALHO_CAPA)
                resultado.titulos_formatados += 1
            if idx >= 2:
                _aplicar(paragraphs, blocos[idx - 2], config.ELEMENTO_CAPA)
                resultado.autores_formatados += 1
            continue

        if RE_NATUREZA_TRABALHO.match(paragraphs[bloco[0]].text.strip()):
            _aplicar(paragraphs, bloco, config.NATUREZA_TRABALHO)
            resultado.naturezas_formatadas += 1
            if idx >= 1:
                _aplicar(paragraphs, blocos[idx - 1], config.TITULO_TRABALHO_CAPA)
                resultado.titulos_formatados += 1
            if idx >= 2:
                _aplicar(paragraphs, blocos[idx - 2], config.ELEMENTO_CAPA)
                resultado.autores_formatados += 1
            if idx + 1 < len(blocos) and RE_ORIENTADOR.search(paragraphs[blocos[idx + 1][0]].text.strip()):
                _aplicar(paragraphs, blocos[idx + 1], config.NATUREZA_TRABALHO)
                resultado.orientadores_formatados += 1
            continue

        if RE_INSTITUICAO.search(paragraphs[bloco[0]].text):
            _aplicar(paragraphs, bloco, config.ELEMENTO_CAPA)
            resultado.institucoes_formatadas += 1

    return resultado
