import docx
from docx.enum.section import WD_SECTION

from formatador_tcc.paginacao import aplicar_numeracao_paginas


def _construir_docx_com_secoes():
    d = docx.Document()
    d.add_paragraph("CAPA")
    d.add_section(WD_SECTION.NEW_PAGE)
    d.add_paragraph("FICHA CATALOGRÁFICA")
    d.add_section(WD_SECTION.NEW_PAGE)
    d.add_paragraph("SUMÁRIO")
    p = d.add_paragraph("INTRODUÇÃO")
    p.style = d.styles["Heading 1"]
    d.add_paragraph("Texto do primeiro parágrafo.")
    p2 = d.add_paragraph("REFERÊNCIAS")
    p2.style = d.styles["Heading 1"]
    d.add_paragraph("SILVA, João. Um livro. São Paulo: Editora, 2020.")
    return d


def test_numeracao_aplicada_a_partir_da_introducao():
    d = _construir_docx_com_secoes()
    resultado = aplicar_numeracao_paginas(d)

    assert resultado.aplicada is True
    # a Introdução está na última seção (2 quebras de seção antes dela)
    assert resultado.secao_introducao == 2

    secoes = d.sections
    # seções anteriores continuam sem cabeçalho próprio
    assert secoes[0].header.is_linked_to_previous is True
    assert secoes[1].header.is_linked_to_previous is True
    # a seção da Introdução tem cabeçalho próprio com o campo de página
    assert secoes[2].header.is_linked_to_previous is False


def test_campo_de_pagina_tem_formula_com_offset_correto():
    d = _construir_docx_com_secoes()
    aplicar_numeracao_paginas(d)

    header = d.sections[2].header
    xml = header._element.xml
    assert "=PAGE-2" in xml
    assert 'w:jc w:val="right"' in xml


def test_nao_duplica_campo_se_ja_houver_conteudo_no_cabecalho(construir_docx):
    d = _construir_docx_com_secoes()
    # simula um bloco de número de página já existente (como no modelo oficial)
    secao = d.sections[2]
    secao.header.is_linked_to_previous = False
    secao.header.add_paragraph("2")

    aplicar_numeracao_paginas(d)

    header = d.sections[2].header
    assert len(header.paragraphs) == 1
    assert "=PAGE-2" in header._element.xml


def test_documento_sem_titulo1_nao_aplica_numeracao():
    d = docx.Document()
    d.add_paragraph("Texto qualquer sem nenhum capítulo numerado.")
    resultado = aplicar_numeracao_paginas(d)
    assert resultado.aplicada is False


def test_documento_final_continua_valido_apos_numeracao(tmp_path):
    import zipfile
    from lxml import etree

    d = _construir_docx_com_secoes()
    aplicar_numeracao_paginas(d)
    caminho = tmp_path / "saida.docx"
    d.save(caminho)

    with zipfile.ZipFile(caminho) as z:
        assert z.testzip() is None
        for name in z.namelist():
            if name.endswith(".xml") or name.endswith(".rels"):
                etree.fromstring(z.read(name))
