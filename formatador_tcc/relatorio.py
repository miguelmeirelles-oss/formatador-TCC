"""Monta o relatório final (Markdown) a partir dos resultados de cada checagem."""
from __future__ import annotations

from collections import Counter

from .capa import ResultadoCapa
from .citacoes import ResultadoCruzamento
from .contagem import RelatorioContagem
from .formatador import EventoFormatacao
from .paginacao import ResultadoPaginacao
from .referencias_check import RelatorioReferencias
from .sumario import ResultadoSumario


def _bloco_formatacao(eventos: list[EventoFormatacao]) -> list[str]:
    contagem = Counter(e.categoria for e in eventos)
    linhas = ["## 1. Formatação aplicada", ""]
    if not eventos:
        linhas.append("Nenhum parágrafo reconhecido foi reformatado -- confira se o "
                       "documento segue a estrutura do Modelo de TCC oficial (títulos, "
                       "RESUMO, ABSTRACT, SUMÁRIO, REFERÊNCIAS).")
    else:
        linhas.append(f"{len(eventos)} parágrafos tiveram a formatação normalizada "
                       "(fonte, tamanho, alinhamento, espaçamento e recuo -- o texto em si "
                       "não foi alterado). Por categoria:")
        linhas.append("")
        rotulos = {
            "titulo1": "Títulos de capítulo (seção primária)",
            "titulo2": "Subtítulos (seção secundária)",
            "titulo3": "Subtítulos (seção terciária)",
            "titulo4": "Subtítulos (seção quaternária)",
            "titulo5": "Subtítulos (seção quinária)",
            "titulo_sem_numero": "Títulos sem indicativo numérico",
            "legenda": "Legendas de figura/quadro/tabela",
            "fonte_ilustracao": "Linhas de fonte de ilustração",
            "citacao_longa": "Citações diretas longas (recuo de 4cm aplicado)",
            "referencia": "Entradas da lista de referências",
            "resumo_corpo": "Parágrafos do Resumo/Abstract",
            "corpo": "Parágrafos de texto corrido",
            "corpo_sem_recuo": "Parágrafos de texto corrido (1º após título)",
            "tabela_conteudo": "Parágrafos dentro de tabelas (fonte normalizada)",
            "quebra_pagina_removida": "Quebras de página manuais removidas (redundantes)",
        }
        for categoria, total in contagem.most_common():
            linhas.append(f"- {rotulos.get(categoria, categoria)}: {total}")
    return linhas


def _bloco_capa(resultado: ResultadoCapa) -> list[str]:
    linhas = ["", "## 2. Capa e Folha de Rosto", ""]
    total = (
        resultado.titulos_formatados + resultado.autores_formatados
        + resultado.tipos_documento_formatados + resultado.naturezas_formatadas
        + resultado.orientadores_formatados + resultado.institucoes_formatadas
    )
    if total == 0:
        linhas.append(
            "Não foi possível reconhecer com segurança os elementos da Capa/Folha de "
            "Rosto (título, nome do autor...) -- confira manualmente se seguem o "
            "Apêndice I (negrito, caixa alta, centralizado, tamanho 14 para o título "
            "e 12 para o restante)."
        )
        return linhas
    linhas.append(
        "Reconhecidos a partir de âncoras de texto fixo (\"TRABALHO DE CONCLUSÃO DE "
        "CURSO\" na Capa, \"Trabalho de Conclusão de Curso apresentado...\" na Folha "
        "de Rosto) -- a posição dos blocos de texto ao redor dessas âncoras segue a "
        "ordem descrita no Apêndice I:"
    )
    linhas.append("")
    if resultado.titulos_formatados:
        linhas.append(f"- Título do trabalho: {resultado.titulos_formatados} (negrito, caixa alta, tamanho 14, centralizado)")
    if resultado.autores_formatados:
        linhas.append(f"- Nome do(s) autor(es): {resultado.autores_formatados} (negrito, caixa alta, tamanho 12, centralizado)")
    if resultado.tipos_documento_formatados:
        linhas.append(f"- Tipo do documento (Capa): {resultado.tipos_documento_formatados}")
    if resultado.naturezas_formatadas:
        linhas.append(f"- Natureza do trabalho (Folha de Rosto): {resultado.naturezas_formatadas}")
    if resultado.orientadores_formatados:
        linhas.append(f"- Orientador(es): {resultado.orientadores_formatados}")
    if resultado.institucoes_formatadas:
        linhas.append(f"- Nome da instituição: {resultado.institucoes_formatadas}")
    linhas.append("")
    linhas.append(
        "**Importante:** Local e Ano (últimas linhas da Capa/Folha de Rosto) não são "
        "reconhecidos automaticamente -- confira manualmente se estão em negrito, "
        "caixa alta, centralizados, tamanho 12."
    )
    return linhas


def _bloco_sumario(resultado: ResultadoSumario) -> list[str]:
    linhas = ["", "## 3. Sumário", ""]
    if not resultado.encontrado:
        linhas.append("Não foi encontrado um título \"SUMÁRIO\" no documento -- o sumário "
                       "não pôde ser reconstruído automaticamente. Verifique se essa página "
                       "existe e usa exatamente a palavra \"SUMÁRIO\".")
        return linhas
    linhas.append(
        f"{resultado.entradas_removidas} entrada(s) antiga(s) do sumário foram removidas e "
        "substituídas por um campo de Sumário automático do Word."
    )
    linhas.append(
        "**Importante:** ao abrir o arquivo gerado no Word, os números de página do sumário "
        "são recalculados automaticamente (o documento está configurado para atualizar campos "
        "ao abrir). Se não atualizar sozinho, clique com o botão direito sobre o sumário e "
        "escolha \"Atualizar campo\" -> \"Atualizar o índice inteiro\" (ou selecione o texto e "
        "pressione F9)."
    )
    return linhas


def _bloco_paginacao(resultado: ResultadoPaginacao) -> list[str]:
    linhas = ["", "## 4. Numeração de página", ""]
    if not resultado.aplicada:
        linhas.append(
            "Não foi possível localizar o início da Introdução -- a numeração de página não "
            "foi aplicada. Verifique se o documento tem um título de capítulo (\"1 "
            "INTRODUÇÃO\" ou equivalente com estilo Heading 1)."
        )
        return linhas
    linhas.append(
        "Número de página inserido no canto superior direito, a partir da Introdução (as "
        "páginas anteriores não são numeradas, mas contam para a numeração, exceto Capa e "
        "Ficha Catalográfica -- por isso o número já sai descontado em 2, conforme o "
        "Apêndice I). Referências, apêndices e anexos continuam a mesma numeração."
    )
    linhas.append(
        "Assim como o sumário, esse número é recalculado automaticamente pelo Word ao abrir "
        "o arquivo."
    )
    return linhas


def _bloco_contagem(rel: RelatorioContagem) -> list[str]:
    linhas = ["", "## 5. Contagem de palavras (Resumo/Abstract)", ""]
    for secao in (rel.resumo, rel.abstract):
        marca = "✅" if secao.dentro_do_limite else "⚠️"
        linhas.append(f"- {marca} {secao.mensagem}")
    return linhas


def _bloco_citacoes(res: ResultadoCruzamento) -> list[str]:
    linhas = ["", "## 6. Cruzamento entre citações e referências", ""]
    linhas.append(f"- {len(res.citacoes)} citação(ões) identificada(s) no corpo do texto.")
    linhas.append(f"- {len(res.referencias)} entrada(s) na lista de Referências.")
    linhas.append("")

    if res.citacoes_sem_referencia:
        linhas.append(f"### ⚠️ {len(res.citacoes_sem_referencia)} citação(ões) sem referência correspondente")
        linhas.append("")
        for c in res.citacoes_sem_referencia:
            autores = "; ".join(c.autores)
            linhas.append(f"- **{autores}, {c.ano}** -- \"...{c.trecho}...\"")
        linhas.append("")
    else:
        linhas.append("✅ Todas as citações identificadas têm referência correspondente.")
        linhas.append("")

    if res.referencias_sem_citacao:
        linhas.append(f"### ⚠️ {len(res.referencias_sem_citacao)} referência(s) nunca citada(s) no texto")
        linhas.append("")
        for r in res.referencias_sem_citacao:
            linhas.append(f"- {r.texto[:120]}")
        linhas.append("")
    else:
        linhas.append("✅ Todas as referências da lista foram citadas em algum ponto do texto.")

    linhas.append("")
    linhas.append(
        "_Observação: citações no formato \"apud\" só exigem referência da fonte "
        "efetivamente consultada (após o \"apud\"), conforme a NBR 10520. A detecção de "
        "citações é automática e pode não reconhecer formatos incomuns -- revise "
        "manualmente antes de considerar a lista definitiva._"
    )
    return linhas


def _bloco_referencias(rel: RelatorioReferencias) -> list[str]:
    linhas = ["", "## 7. Formatação das Referências (NBR 6023)", ""]
    linhas.append(f"{rel.total_entradas} entrada(s) verificada(s).")
    linhas.append("")

    corrigidos = [p for p in rel.problemas if p.corrigivel_automaticamente]
    manuais = [p for p in rel.problemas if not p.corrigivel_automaticamente]

    if corrigidos:
        linhas.append(f"### Corrigidos automaticamente ({len(corrigidos)})")
        linhas.append("")
        for p in corrigidos:
            linhas.append(f"- {p.mensagem} -- \"{p.texto[:80]}\"")
        linhas.append("")

    if manuais:
        linhas.append(f"### ⚠️ Precisam de revisão manual ({len(manuais)})")
        linhas.append("")
        for p in manuais:
            linhas.append(f"- {p.mensagem} -- \"{p.texto[:80]}\"")
    else:
        linhas.append("✅ Nenhum problema de conteúdo encontrado nas referências.")

    return linhas


def gerar_relatorio_markdown(
    *,
    nome_entrada: str,
    nome_saida: str,
    eventos_formatacao: list[EventoFormatacao],
    resultado_capa: ResultadoCapa,
    resultado_sumario: ResultadoSumario,
    resultado_paginacao: ResultadoPaginacao,
    resultado_citacoes: ResultadoCruzamento,
    relatorio_referencias: RelatorioReferencias,
    relatorio_contagem: RelatorioContagem,
) -> str:
    linhas = [
        "# Relatório de Formatação ABNT",
        "",
        f"- Arquivo de entrada: `{nome_entrada}`",
        f"- Arquivo formatado gerado: `{nome_saida}`",
        "",
        "Este relatório é gerado automaticamente. Ele aponta o que foi corrigido e o "
        "que ainda precisa da sua revisão -- ele não substitui a leitura atenta das "
        "normas nem a orientação do seu professor orientador.",
    ]
    linhas += _bloco_formatacao(eventos_formatacao)
    linhas += _bloco_capa(resultado_capa)
    linhas += _bloco_sumario(resultado_sumario)
    linhas += _bloco_paginacao(resultado_paginacao)
    linhas += _bloco_contagem(relatorio_contagem)
    linhas += _bloco_citacoes(resultado_citacoes)
    linhas += _bloco_referencias(relatorio_referencias)
    return "\n".join(linhas) + "\n"
