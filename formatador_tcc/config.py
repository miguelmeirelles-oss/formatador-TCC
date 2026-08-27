"""Perfil de formatação ABNT extraído do documento 'Modelo de TCC oficial'.

Os valores abaixo foram lidos diretamente de templates/Modelo_TCC_oficial.docx
(margens de página e propriedades de cada estilo). Eles são a fonte da
verdade para o formatador -- se a instituição atualizar o modelo oficial,
estes valores devem ser reconferidos.
"""
from __future__ import annotations

from dataclasses import dataclass, field


CM = 360000  # EMU por centímetro


@dataclass(frozen=True)
class PageSetup:
    largura_cm: float = 21.0   # A4
    altura_cm: float = 29.7    # A4
    margem_superior_cm: float = 3.0
    margem_inferior_cm: float = 2.0
    margem_esquerda_cm: float = 3.0
    margem_direita_cm: float = 2.0


@dataclass(frozen=True)
class EstiloParagrafo:
    """Formatação a aplicar a um parágrafo, sem alterar o texto/runs."""
    fonte: str = "Times New Roman"
    tamanho_pt: float = 12.0
    negrito: bool | None = None
    italico: bool | None = None
    maiusculas: bool = False
    alinhamento: str = "justify"       # justify | left | center | right
    espacamento_linha: float = 1.5      # 1.0 = simples, 1.5, 2.0
    espaco_antes_pt: float = 0.0
    espaco_depois_pt: float = 0.0
    recuo_primeira_linha_cm: float = 0.0
    recuo_esquerdo_cm: float = 0.0


PAGE_SETUP = PageSetup()

# Estilo do texto corrido (corpo do trabalho).
CORPO_TEXTO = EstiloParagrafo(
    alinhamento="justify",
    espacamento_linha=1.5,
    recuo_primeira_linha_cm=1.25,
)

# Títulos numerados de seção primária (1 INTRODUÇÃO, 2 OBJETIVOS, ...).
TITULO_SECAO_PRIMARIA = EstiloParagrafo(
    tamanho_pt=12.0,
    negrito=True,
    maiusculas=True,
    alinhamento="left",
    espacamento_linha=1.5,
    espaco_antes_pt=10.0,
    espaco_depois_pt=5.0,
)

# Subtítulos de seção (1.1, 1.2, ...).
TITULO_SECAO_SECUNDARIA = EstiloParagrafo(
    tamanho_pt=12.0,
    negrito=True,
    maiusculas=False,
    alinhamento="left",
    espacamento_linha=1.5,
    espaco_antes_pt=10.0,
    espaco_depois_pt=5.0,
)

TITULO_SECAO_TERCIARIA = EstiloParagrafo(
    tamanho_pt=12.0,
    negrito=True,
    maiusculas=False,
    alinhamento="left",
    espacamento_linha=1.5,
    espaco_antes_pt=12.0,
    espaco_depois_pt=3.0,
)

# Títulos sem indicativo numérico (RESUMO, ABSTRACT, SUMÁRIO, REFERÊNCIAS,
# AGRADECIMENTOS, DEDICATÓRIA, EPÍGRAFE, listas...) -- centralizados, ABNT NBR 14724.
TITULO_SEM_NUMERO = EstiloParagrafo(
    tamanho_pt=16.0,
    negrito=True,
    maiusculas=True,
    alinhamento="center",
    espacamento_linha=1.5,
    espaco_antes_pt=12.0,
    espaco_depois_pt=3.0,
)

# Legenda de figura/quadro/tabela (linha "Figura 1 - ...").
LEGENDA = EstiloParagrafo(
    tamanho_pt=12.0,
    negrito=True,
    alinhamento="center",
    espacamento_linha=1.5,
)

# Fonte da ilustração (linha "Fonte: ...", abaixo da imagem/tabela).
FONTE_ILUSTRACAO = EstiloParagrafo(
    tamanho_pt=10.0,
    italico=True,
    alinhamento="center",
    espacamento_linha=1.0,
    espaco_depois_pt=10.0,
)

# Citação direta longa (mais de 3 linhas) -- NBR 10520.
CITACAO_LONGA = EstiloParagrafo(
    tamanho_pt=10.0,
    alinhamento="justify",
    espacamento_linha=1.0,
    recuo_esquerdo_cm=4.0,
    espaco_antes_pt=6.0,
    espaco_depois_pt=6.0,
)

# Referência bibliográfica -- NBR 6023: alinhada à esquerda, espaço simples
# dentro da referência, espaço entre referências, sem recuo de primeira linha.
REFERENCIA = EstiloParagrafo(
    tamanho_pt=12.0,
    alinhamento="left",
    espacamento_linha=1.0,
    espaco_depois_pt=6.0,
)

# Nomes de seções "sem indicativo numérico" que, no modelo oficial, usam o
# estilo centralizado (Título de Seção / 16pt) e NÃO entram no Sumário:
# elementos pré-textuais + listas. Comparação feita normalizada (maiúsculas,
# sem acento).
SECOES_SEM_NUMERO = {
    "DEDICATORIA",
    "AGRADECIMENTOS",
    "EPIGRAFE",
    "RESUMO",
    "ABSTRACT",
    "LISTA DE ILUSTRACOES",
    "LISTA DE FIGURAS",
    "LISTA DE QUADROS",
    "LISTA DE TABELAS",
    "LISTA DE ABREVIATURAS E SIGLAS",
    "SUMARIO",
    "FOLHA DE APROVACAO",
}

# No modelo oficial, REFERÊNCIAS/APÊNDICE/ANEXO usam o mesmo estilo "Heading 1"
# dos capítulos numerados (negrito, maiúsculas, alinhado à esquerda) -- e por
# isso entram no Sumário junto com os capítulos, diferente das seções acima.
SECOES_POS_TEXTUAIS_TITULO1 = {"REFERENCIAS", "APENDICE", "ANEXO", "GLOSSARIO"}

# Marca o início da lista de referências bibliográficas.
INICIO_REFERENCIAS = "REFERENCIAS"

# Regra explícita do modelo oficial: "O resumo deve conter no mínimo 150
# palavras e no máximo 500 palavras." Vale também para o abstract.
RESUMO_MIN_PALAVRAS = 150
RESUMO_MAX_PALAVRAS = 500
