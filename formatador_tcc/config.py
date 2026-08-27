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

# Subtítulos de seção (1.1, 1.2, ...). Apêndice I -- "Regras gerais de
# apresentação": "Para as seções secundárias: somente caixa alta e sem
# negrito" (confirmado também pelo comentário da autora do próprio modelo
# oficial: "Seção secundária: 1.1. CAIXA ALTA").
TITULO_SECAO_SECUNDARIA = EstiloParagrafo(
    tamanho_pt=12.0,
    negrito=False,
    maiusculas=True,
    alinhamento="left",
    espacamento_linha=1.5,
    espaco_antes_pt=10.0,
    espaco_depois_pt=5.0,
)

# Apêndice I: "Para as seções terciárias: a primeira letra de cada palavra em
# maiúscula" (Title Case -- não é algo que se force por formatação de
# caractere sem reescrever o texto; fica só o negrito, que o comentário do
# modelo oficial confirma: "Seção terciária: ... Caixa baixa e destaque
# (negrito)").
TITULO_SECAO_TERCIARIA = EstiloParagrafo(
    tamanho_pt=12.0,
    negrito=True,
    maiusculas=False,
    alinhamento="left",
    espacamento_linha=1.5,
    espaco_antes_pt=12.0,
    espaco_depois_pt=3.0,
)

# Seção quaternária (1.1.1.1). Apêndice I: "somente a primeira letra do
# título da seção em maiúscula" (sem negrito) -- o modelo oficial não define
# um estilo próprio para este nível (o "Heading 4" da galeria do Word está
# com a formatação padrão de fábrica, não customizada para este trabalho),
# por isso o espaçamento usado é o mesmo da terciária, que o Apêndice I
# descreve como equivalente ("também devem ser separados... por 1 espaço").
TITULO_SECAO_QUATERNARIA = EstiloParagrafo(
    tamanho_pt=12.0,
    negrito=False,
    maiusculas=False,
    alinhamento="left",
    espacamento_linha=1.5,
    espaco_antes_pt=12.0,
    espaco_depois_pt=3.0,
)

# Seção quinária (1.1.1.1.1) -- último nível permitido pelo Apêndice I
# ("Deve-se limitar a numeração progressiva até a seção quinária, inclusive").
# Apêndice I: "Primeira letra do título maiúscula e em itálico".
TITULO_SECAO_QUINARIA = EstiloParagrafo(
    tamanho_pt=12.0,
    negrito=False,
    italico=True,
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
# Medido diretamente nos 3 exemplos reais do modelo oficial (não é itálico,
# não é centralizado, 11pt -- só a palavra "Fonte:" vem em negrito, o resto
# da linha não). O negrito parcial é aplicado à parte em formatador.py, não
# aqui (este EstiloParagrafo não força negrito/itálico na linha toda).
FONTE_ILUSTRACAO = EstiloParagrafo(
    tamanho_pt=11.0,
    alinhamento="justify",
    espacamento_linha=1.0,
    espaco_depois_pt=10.0,
)
PREFIXO_FONTE_NEGRITO = "Fonte:"

# Apêndice I: "O tamanho da fonte na parte interna da tabela é 10" (e também
# "na parte interna das ilustrações"). Só o tamanho é normatizado -- o
# Apêndice I não define alinhamento/negrito para o conteúdo de dentro da
# tabela, então isso fica como o aluno escreveu.
TABELA_CONTEUDO_TAMANHO_PT = 10.0

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

# PAGINAÇÃO -- Apêndice I + comentário do próprio modelo oficial:
# "Todas as folhas do trabalho, a partir da folha de rosto, devem ser
# contadas sequencialmente, mas não numeradas. A numeração deve ser colocada
# a partir da primeira folha da parte textual (Introdução), em algarismos
# arábicos, no canto superior direito da folha." O comentário do modelo
# oficial detalha o cálculo: contam-se todas as páginas anteriores exceto a
# CAPA e a FICHA CATALOGRÁFICA -- por isso o número mostrado na Introdução é
# sempre "página física atual menos 2".
PAGINAS_EXCLUIDAS_DA_CONTAGEM = 2  # Capa + Ficha Catalográfica
