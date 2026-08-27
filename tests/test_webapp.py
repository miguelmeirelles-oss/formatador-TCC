import sys
from io import BytesIO
from pathlib import Path

import pytest

pytest.importorskip("flask")

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from webapp.app import app  # noqa: E402

TEMPLATE = Path(__file__).resolve().parent.parent / "templates" / "Modelo_TCC_oficial.docx"


@pytest.fixture
def client():
    app.config.update(TESTING=True)
    return app.test_client()


def test_pagina_inicial(client):
    r = client.get("/")
    assert r.status_code == 200
    assert b"Formatar e conferir" in r.data


def test_extensao_invalida_retorna_erro(client):
    r = client.post(
        "/processar",
        data={"documento": (BytesIO(b"texto qualquer"), "arquivo.txt")},
        content_type="multipart/form-data",
    )
    assert r.status_code == 400
    assert "precisa ser um .docx" in r.get_data(as_text=True)


def test_docx_invalido_retorna_erro_amigavel(client):
    r = client.post(
        "/processar",
        data={"documento": (BytesIO(b"nao e um zip valido"), "quebrado.docx")},
        content_type="multipart/form-data",
    )
    assert r.status_code == 400
    assert "Não foi possível processar" in r.get_data(as_text=True)


def test_download_com_token_invalido_404(client):
    r = client.get("/baixar/token-que-nao-existe")
    assert r.status_code == 404


@pytest.mark.skipif(not TEMPLATE.exists(), reason="template oficial não versionado neste checkout")
def test_fluxo_completo_upload_e_download(client):
    conteudo = TEMPLATE.read_bytes()
    r = client.post(
        "/processar",
        data={"documento": (BytesIO(conteudo), "Modelo_TCC_oficial.docx")},
        content_type="multipart/form-data",
    )
    assert r.status_code == 200
    pagina = r.get_data(as_text=True)
    assert "Baixar" in pagina
    assert "Formatação aplicada" in pagina

    import re
    m = re.search(r"/baixar/([0-9a-f]+)", pagina)
    assert m is not None
    token = m.group(1)

    r2 = client.get(f"/baixar/{token}")
    assert r2.status_code == 200
    assert r2.data[:2] == b"PK"  # .docx é um arquivo zip
