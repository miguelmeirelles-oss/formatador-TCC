"""Cruzamento entre citações no corpo do texto e a lista de Referências.

Cobre os dois formatos de citação-autor-data da NBR 10520:
  - Parentética:  "...blá blá (SILVA, 2020)."      / "(SILVA; SOUZA, 2020)"
                  "(SILVA, 2019; SOUZA, 2020)"       / "(SILVA et al., 2020)"
  - Narrativa:    "Segundo Silva (2020), blá blá."   / "Silva e Souza (2020)"

E o caso "apud" (citação de citação): só a fonte efetivamente consultada
(o que vem depois de "apud") precisa constar nas Referências -- o autor
original citado por meio dela não precisa ter entrada própria.

O módulo não altera o texto do aluno em nenhum momento: apenas lê.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from .classify import EstadoClassificacao, classificar_paragrafo
from .texto import normalizar, RE_ANO

CATEGORIAS_CITAVEIS = {
    "corpo",
    "corpo_sem_recuo",
    "resumo_corpo",
    "citacao_longa",
    "legenda",
    "fonte_ilustracao",
}

RE_PARENS = re.compile(r"\(([^()]*)\)")
RE_SOMENTE_ANO = re.compile(r"^\d{4}[a-z]?(\s*[,;]\s*p\.?\s*\d+[-–—]?\d*)?$", re.IGNORECASE)
RE_NOME_PRECEDENTE = re.compile(
    r"([A-ZÀ-Ú][\wÀ-ÿ.'\-]*(?:\s+(?:e|&|;)\s+[A-ZÀ-Ú][\wÀ-ÿ.'\-]*)*)(?:\s+et\s*al\.?)?\s*$"
)
RE_APUD = re.compile(r"\bapud\b", re.IGNORECASE)
RE_ET_AL = re.compile(r"\bet\s*al\.?", re.IGNORECASE)


def _norm_autor(s: str) -> str:
    s = re.sub(r"\(org[s]?\.?\)|\(coord\.?\)", "", s, flags=re.IGNORECASE)
    s = s.strip(" .;,")
    return normalizar(s)


@dataclass
class Citacao:
    autores: tuple[str, ...]
    ano: str
    et_al: bool
    forma: str          # "parentetica" | "narrativa"
    apud: bool
    trecho: str          # trecho de contexto, para o relatório
    paragrafo_indice: int


@dataclass
class EntradaReferencia:
    paragrafo_indice: int
    texto: str
    autores: tuple[str, ...] = field(default_factory=tuple)
    ano: str | None = None
    valida: bool = False


def _extrair_autores_ano_de_bloco(bloco: str) -> list[tuple[tuple[str, ...], str, bool]]:
    """Recebe o conteúdo já sem 'apud' à esquerda (se houver) e devolve uma
    lista de (autores, ano, et_al) -- pode haver mais de uma citação quando
    há ';' separando obras diferentes dentro do mesmo parêntese."""
    resultado: list[tuple[tuple[str, ...], str, bool]] = []
    pendentes: list[str] = []
    for segmento in bloco.split(";"):
        segmento = segmento.strip()
        if not segmento:
            continue
        m = RE_ANO.search(segmento)
        if not m:
            nome = segmento.strip(" .")
            if nome:
                pendentes.append(nome)
            continue
        antes_bruto = segmento[: m.start()]
        # Exige vírgula (ou "et al.") imediatamente antes do ano -- é o que
        # diferencia uma citação autor-data ("SILVA, 2020") de outro número
        # entre parênteses que por coincidência contém 4 dígitos plausíveis
        # (ex.: "NBR 14720/2002", uma norma técnica, não uma citação).
        if not pendentes and not re.search(r"(,|\bet\s*al\.?)\s*$", antes_bruto, re.IGNORECASE):
            continue
        antes_do_ano = antes_bruto.strip(" ,")
        et_al = bool(RE_ET_AL.search(antes_do_ano))
        antes_do_ano = RE_ET_AL.sub("", antes_do_ano).strip(" ,")
        autores = [a for a in re.split(r"\s+e\s+|&|,", antes_do_ano) if a.strip()]
        autores = [_norm_autor(a) for a in autores]
        autores = [a for a in autores if a]
        todos = tuple(dict.fromkeys([_norm_autor(p) for p in pendentes] + autores))
        pendentes = []
        if todos:
            resultado.append((todos, m.group(0), et_al))
    return resultado


def _capturar_nome_precedente(texto: str, pos_abre_parens: int) -> str | None:
    trecho_antes = texto[max(0, pos_abre_parens - 80): pos_abre_parens]
    m = RE_NOME_PRECEDENTE.search(trecho_antes)
    if not m:
        return None
    nome = m.group(0).strip()
    return nome or None


def extrair_citacoes_do_paragrafo(texto: str, indice: int) -> list[Citacao]:
    citacoes: list[Citacao] = []
    for m in RE_PARENS.finditer(texto):
        conteudo = m.group(1).strip()
        if not RE_ANO.search(conteudo):
            continue

        apud = bool(RE_APUD.search(conteudo))
        if apud:
            conteudo_efetivo = RE_APUD.split(conteudo, maxsplit=1)[-1]
        else:
            conteudo_efetivo = conteudo

        if RE_SOMENTE_ANO.match(conteudo.strip()):
            # forma narrativa: "Silva (2020)" -- autor está antes do "("
            nome_precedente = _capturar_nome_precedente(texto, m.start())
            if not nome_precedente:
                continue
            if RE_APUD.search(nome_precedente):
                apud = True
                nome_precedente = RE_APUD.split(nome_precedente, maxsplit=1)[-1]
            et_al = bool(RE_ET_AL.search(nome_precedente))
            nome_precedente = RE_ET_AL.sub("", nome_precedente).strip()
            autores = tuple(
                _norm_autor(a)
                for a in re.split(r"\s+e\s+|&|;", nome_precedente)
                if _norm_autor(a)
            )
            ano_m = RE_ANO.search(conteudo)
            if not autores or not ano_m:
                continue
            citacoes.append(
                Citacao(
                    autores=autores,
                    ano=ano_m.group(0),
                    et_al=et_al,
                    forma="narrativa",
                    apud=apud,
                    trecho=texto[max(0, m.start() - 40): m.end() + 5].strip(),
                    paragrafo_indice=indice,
                )
            )
            continue

        for autores, ano, et_al in _extrair_autores_ano_de_bloco(conteudo_efetivo):
            citacoes.append(
                Citacao(
                    autores=autores,
                    ano=ano,
                    et_al=et_al,
                    forma="parentetica",
                    apud=apud,
                    trecho=texto[max(0, m.start() - 20): m.end() + 5].strip(),
                    paragrafo_indice=indice,
                )
            )
    return citacoes


def _extrair_referencia(paragrafo_indice: int, texto: str) -> EntradaReferencia:
    entrada = EntradaReferencia(paragrafo_indice=paragrafo_indice, texto=texto)

    primeiro_ponto = texto.find(". ")
    autores_bruto = texto[:primeiro_ponto] if primeiro_ponto > 0 else texto
    if "," not in autores_bruto:
        # não começa com "SOBRENOME, Nome" -- provavelmente obra sem autor
        # pessoal (ex.: entidade em maiúsculas sem vírgula, ou item mal
        # formatado). Ainda tentamos extrair o ano para não perder o cruzamento.
        m = list(RE_ANO.finditer(texto))
        entrada.ano = m[-1].group(0) if m else None
        return entrada

    autores = [a.strip() for a in autores_bruto.split(";") if a.strip()]
    sobrenomes = []
    for a in autores:
        antes_virgula = a.split(",")[0]
        s = _norm_autor(antes_virgula)
        if s:
            sobrenomes.append(s)
    entrada.autores = tuple(dict.fromkeys(sobrenomes))

    anos = list(RE_ANO.finditer(texto))
    entrada.ano = anos[-1].group(0) if anos else None
    entrada.valida = bool(entrada.autores and entrada.ano)
    return entrada


@dataclass
class ResultadoCruzamento:
    citacoes: list[Citacao]
    referencias: list[EntradaReferencia]
    citacoes_sem_referencia: list[Citacao]
    referencias_sem_citacao: list[EntradaReferencia]
    citacoes_nao_identificadas: int = 0


def cruzar_citacoes_e_referencias(document) -> ResultadoCruzamento:
    estado = EstadoClassificacao()
    citacoes: list[Citacao] = []
    referencias: list[EntradaReferencia] = []

    for i, paragraph in enumerate(document.paragraphs):
        categoria = classificar_paragrafo(paragraph, estado)
        texto = paragraph.text.strip()
        if not texto:
            continue
        if categoria == "referencia":
            referencias.append(_extrair_referencia(i, texto))
        elif categoria in CATEGORIAS_CITAVEIS:
            citacoes.extend(extrair_citacoes_do_paragrafo(texto, i))

    chaves_referenciadas: set[tuple[str, str]] = set()
    primeiro_autor_ano: set[tuple[str, str]] = set()
    for ref in referencias:
        if not ref.ano:
            continue
        for autor in ref.autores:
            chaves_referenciadas.add((autor, ref.ano))
            chaves_referenciadas.add((autor, ref.ano[:4]))
        if ref.autores:
            primeiro_autor_ano.add((ref.autores[0], ref.ano))
            primeiro_autor_ano.add((ref.autores[0], ref.ano[:4]))

    def citacao_tem_referencia(c: Citacao) -> bool:
        if c.apud:
            # a fonte original (antes do apud) não precisa estar nas referências
            pass
        if not c.autores:
            return True  # não foi possível extrair autor -- não reportar como órfã
        if c.et_al or len(c.autores) == 1:
            return (c.autores[0], c.ano) in primeiro_autor_ano
        return all((a, c.ano) in chaves_referenciadas for a in c.autores)

    citacoes_orfas = [c for c in citacoes if c.autores and not citacao_tem_referencia(c)]

    autores_citados: set[str] = set()
    for c in citacoes:
        autores_citados.update(c.autores)

    referencias_sem_citacao = [
        ref
        for ref in referencias
        if ref.valida and not any(a in autores_citados for a in ref.autores)
    ]

    return ResultadoCruzamento(
        citacoes=citacoes,
        referencias=referencias,
        citacoes_sem_referencia=citacoes_orfas,
        referencias_sem_citacao=referencias_sem_citacao,
    )
