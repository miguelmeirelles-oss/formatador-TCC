from pathlib import Path

import docx
from docx.shared import Cm

from formatador_tcc.classify import EstadoClassificacao, classificar_paragrafo
from formatador_tcc.formatador import formatar_documento
from formatador_tcc.sumario import reconstruir_sumario

TEMPLATE = Path(__file__).resolve().parent.parent / "templates" / "Modelo_TCC_oficial.docx"


def test_formatar_nao_altera_texto(construir_docx):
    d = construir_docx([
        ("titulo_sem_numero", "SUMÁRIO"),
        ("heading1", "INTRODUÇÃO"),
        ("texto", "Um parágrafo qualquer com texto do aluno que não deve mudar."),
        ("heading1", "REFERÊNCIAS"),
        ("texto", "SILVA, João. Um livro. São Paulo: Editora, 2020."),
    ])
    textos_antes = [p.text for p in d.paragraphs]
    formatar_documento(d)
    textos_depois = [p.text for p in d.paragraphs]
    assert textos_antes == textos_depois


def test_titulo_numerado_digitado_manualmente_e_reconhecido(construir_docx):
    d = construir_docx([
        ("titulo_sem_numero", "SUMÁRIO"),
        ("texto", "1 INTRODUÇÃO"),
        ("texto", "Texto do primeiro parágrafo."),
        ("texto", "3.1 Um subtítulo qualquer"),
    ])
    estado = EstadoClassificacao()
    categorias = [classificar_paragrafo(p, estado) for p in d.paragraphs]
    assert categorias[1] == "titulo1"
    assert categorias[3] == "titulo2"


def test_paragrafo_corpo_fica_justificado_com_recuo(construir_docx):
    d = construir_docx([
        ("titulo_sem_numero", "SUMÁRIO"),
        ("heading1", "INTRODUÇÃO"),
        ("texto", "Primeiro parágrafo logo após o título."),
        ("texto", "Segundo parágrafo, este sim com recuo de primeira linha."),
    ])
    formatar_documento(d)
    paragrafos = d.paragraphs
    # segundo parágrafo de corpo (índice 3) deve ter recuo ~1.25cm
    recuo = paragrafos[3].paragraph_format.first_line_indent
    assert recuo is not None
    assert abs(recuo.cm - 1.25) < 0.05


def test_sumario_reconstroi_campo_toc(construir_docx):
    d = construir_docx([
        ("titulo_sem_numero", "SUMÁRIO"),
        ("texto", "1\tINTRODUÇÃO\t12"),
        ("texto", "2\tOBJETIVOS\t13"),
        ("heading1", "INTRODUÇÃO"),
        ("texto", "Corpo do texto."),
    ])
    formatar_documento(d)
    resultado = reconstruir_sumario(d)
    assert resultado.encontrado is True
    assert resultado.entradas_removidas == 2
    textos = [p.text for p in d.paragraphs]
    assert "1\tINTRODUÇÃO\t12" not in textos
    assert any("Sumário gerado automaticamente" in t for t in textos)


def test_documento_oficial_end_to_end():
    if not TEMPLATE.exists():
        return  # template não versionado neste checkout -- pula silenciosamente
    d = docx.Document(str(TEMPLATE))

    estado = EstadoClassificacao()
    textos_esperados = [p.text for p in d.paragraphs if classificar_paragrafo(p, estado) != "sumario_entrada"]

    eventos = formatar_documento(d)
    assert len(eventos) > 50  # documento extenso, deve reconhecer bastante coisa

    resultado_sumario = reconstruir_sumario(d)
    assert resultado_sumario.encontrado is True

    textos_finais = [p.text for p in d.paragraphs if "Sumário gerado automaticamente" not in p.text]
    assert textos_finais == textos_esperados
