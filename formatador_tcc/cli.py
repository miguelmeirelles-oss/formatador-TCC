"""Interface de linha de comando.

Uso:
    python -m formatador_tcc entrada.docx [--saida saida.docx] [--relatorio relatorio.md]
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .pipeline import processar_docx_bytes


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

    resultado = processar_docx_bytes(
        args.entrada.read_bytes(),
        nome_entrada=args.entrada.name,
        nome_saida=saida.name,
    )

    saida.write_bytes(resultado.docx_bytes)
    relatorio_path.write_text(resultado.relatorio_markdown, encoding="utf-8")

    print(f"Documento formatado salvo em: {saida}")
    print(f"Relatório salvo em: {relatorio_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
