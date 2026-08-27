"""App web mínimo do Formatador de TCC: o aluno sobe o .docx, a página
mostra o relatório de conferência e oferece o download do .docx formatado.

Roda tudo em memória (nada é gravado em disco) -- o mesmo motor usado pelo
CLI, via `formatador_tcc.pipeline.processar_docx_bytes`.

Uso local:
    pip install -r requirements-dev.txt
    python -m webapp.app
"""
from __future__ import annotations

import io
import sys
import uuid
from collections import OrderedDict
from pathlib import Path

RAIZ_PROJETO = Path(__file__).resolve().parent.parent
if str(RAIZ_PROJETO) not in sys.path:
    sys.path.insert(0, str(RAIZ_PROJETO))

from flask import Flask, abort, redirect, render_template, request, send_file, url_for

from formatador_tcc.pipeline import processar_docx_bytes
from webapp.mdlite import markdown_para_html

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024  # 30 MB

# Guarda em memória os últimos resultados gerados, para o link de download
# funcionar numa segunda requisição sem precisar de disco/banco. Isto é
# suficiente para um servidor único de demonstração/uso departamental; para
# produção com múltiplos processos, trocar por um storage compartilhado
# (ex.: arquivo temporário, S3, etc.) seria o próximo passo.
#
# IMPORTANTE: por isso o Procfile sobe o gunicorn com --workers 1. Com mais
# de um worker, o upload e o download poderiam cair em processos diferentes
# e o link de download "não seria encontrado" -- não aumente os workers sem
# antes trocar esse storage por algo compartilhado entre processos.
_MAX_RESULTADOS_EM_MEMORIA = 50
_resultados: "OrderedDict[str, tuple[bytes, str]]" = OrderedDict()


def _guardar_resultado(conteudo: bytes, nome_arquivo: str) -> str:
    token = uuid.uuid4().hex
    _resultados[token] = (conteudo, nome_arquivo)
    _resultados.move_to_end(token)
    while len(_resultados) > _MAX_RESULTADOS_EM_MEMORIA:
        _resultados.popitem(last=False)
    return token


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/processar")
def processar():
    arquivo = request.files.get("documento")
    if arquivo is None or not arquivo.filename:
        return render_template("index.html", erro="Selecione um arquivo .docx antes de enviar."), 400

    nome_original = arquivo.filename
    if not nome_original.lower().endswith(".docx"):
        return render_template(
            "index.html",
            erro="O arquivo precisa ser um .docx (Word). Se o seu TCC está em .doc, "
                 "abra no Word e use \"Salvar como\" -> .docx antes de enviar.",
        ), 400

    conteudo = arquivo.read()
    try:
        nome_saida = nome_original[:-len(".docx")] + "_formatado.docx"
        resultado = processar_docx_bytes(conteudo, nome_entrada=nome_original, nome_saida=nome_saida)
    except Exception:
        app.logger.exception("Falha ao processar %s", nome_original)
        return render_template(
            "index.html",
            erro="Não foi possível processar esse arquivo. Confira se ele não está "
                 "corrompido e se realmente é um .docx do Word, e tente novamente.",
        ), 400

    token = _guardar_resultado(resultado.docx_bytes, nome_saida)
    relatorio_html = markdown_para_html(resultado.relatorio_markdown)

    return render_template(
        "resultado.html",
        relatorio_html=relatorio_html,
        token=token,
        nome_saida=nome_saida,
    )


@app.get("/baixar/<token>")
def baixar(token: str):
    item = _resultados.get(token)
    if item is None:
        abort(404, "Esse link de download expirou. Envie o documento novamente.")
    conteudo, nome_arquivo = item
    return send_file(
        io.BytesIO(conteudo),
        as_attachment=True,
        download_name=nome_arquivo,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.get("/nova")
def nova():
    return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)
