"""Interface de linha de comando.

Uso:
    python -m formatador_tcc entrada.docx [--saida saida.docx] [--relatorio relatorio.md]
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import docx

from .citacoes import cruzar_citacoes_e_referencias
from .contagem import verificar_contagem_resumo
from .formatador import formatar_documento
from .referencias_check import verificar_referencias
from .relatorio import gerar_relatorio_markdown
from .sumario import reconstruir_sumario


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="formatador_tcc",
        description="Formata um TCC (.docx) segundo o Modelo de TCC oficial (ABNT) "
                    "e gera um relatório de conferência.",
    )
    parser.add_argument("entrada", type=Path, help="Arquivo .docx do aluno")
    parser.add_argument("--saida", type=Path, default=None, help="Caminho do .docx formatado (padrão: <entrada>_formatado.docx)")
    parser.add_argument("--relatorio", type=Path, default=None, help="Caminho do relatório .md (padrão: <entrada>_relatorio.md)")
    args = parser.parse_args(argv)

    if not args.entrada.exists():
        print(f"Arquivo não encontrado: {args.entrada}", file=sys.stderr)
        return 1

    saida = args.saida or args.entrada.with_name(args.entrada.stem + "_formatado.docx")
    relatorio_path = args.relatorio or args.entrada.with_name(args.entrada.stem + "_relatorio.md")

    document = docx.Document(str(args.entrada))

    # Checagens de conteúdo/formatação original -- rodam antes de o
    # formatador normalizar alinhamento/espaçamento/recuo, para o relatório
    # mostrar o que estava errado e o que foi corrigido automaticamente.
    relatorio_referencias = verificar_referencias(document)
    relatorio_contagem = verificar_contagem_resumo(document)
    resultado_citacoes = cruzar_citacoes_e_referencias(document)

    eventos = formatar_documento(document)
    resultado_sumario = reconstruir_sumario(document)

    document.save(str(saida))

    texto_relatorio = gerar_relatorio_markdown(
        nome_entrada=args.entrada.name,
        nome_saida=saida.name,
        eventos_formatacao=eventos,
        resultado_sumario=resultado_sumario,
        resultado_citacoes=resultado_citacoes,
        relatorio_referencias=relatorio_referencias,
        relatorio_contagem=relatorio_contagem,
    )
    relatorio_path.write_text(texto_relatorio, encoding="utf-8")

    print(f"Documento formatado salvo em: {saida}")
    print(f"Relatório salvo em: {relatorio_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
