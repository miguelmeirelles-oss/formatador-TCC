"""Testes de capa.py -- formatação da Capa e Folha de Rosto.

O texto (nome do autor, título) é digitado livremente pelo aluno, sem
nenhum padrão confiável -- por isso a detecção se apoia em duas âncoras de
texto FIXO (não digitadas livremente): "TRABALHO DE CONCLUSÃO DE CURSO" /
"MONOGRAFIA" / "DISSERTAÇÃO" / "PROJETO FINAL" sozinho numa linha (Capa) e
"Trabalho de Conclusão de Curso apresentado..." (Folha de Rosto). A partir
delas, a POSIÇÃO dos blocos de texto vizinhos (não o conteúdo) identifica
autor/título, na ordem que o Apêndice I descreve.
"""
from formatador_tcc.capa import formatar_capa_e_folha_de_rosto


def _bold_caps_size(paragraph):
    runs = paragraph.runs
    return (
        all(r.font.bold for r in runs),
        all(r.font.all_caps for r in runs),
        {r.font.size.pt for r in runs if r.font.size},
    )


def test_capa_completa_reconhecida(construir_docx):
    d = construir_docx([
        ("texto", "CENTRO FEDERAL DE EDUCAÇÃO TECNOLÓGICA CELSO SUCKOW DA FONSECA"),
        ("texto", ""),
        ("texto", "Fulano de Tal"),
        ("texto", ""),
        ("texto", "TÍTULO DO TRABALHO DE EXEMPLO"),
        ("texto", ""),
        ("texto", "TRABALHO DE CONCLUSÃO DE CURSO"),
        ("texto", ""),
        ("texto", "VALENÇA"),
        ("texto", ""),
        ("texto", "2025"),
    ])
    resultado = formatar_capa_e_folha_de_rosto(d)

    assert resultado.institucoes_formatadas == 1
    assert resultado.autores_formatados == 1
    assert resultado.titulos_formatados == 1
    assert resultado.tipos_documento_formatados == 1

    bold, caps, sizes = _bold_caps_size(d.paragraphs[0])
    assert bold and caps and sizes == {12.0}
    bold, caps, sizes = _bold_caps_size(d.paragraphs[2])
    assert bold and caps and sizes == {12.0}
    bold, caps, sizes = _bold_caps_size(d.paragraphs[4])
    assert bold and caps and sizes == {14.0}
    bold, caps, sizes = _bold_caps_size(d.paragraphs[6])
    assert bold and caps and sizes == {12.0}


def test_folha_de_rosto_reconhecida_sem_negrito_na_natureza(construir_docx):
    d = construir_docx([
        ("texto", "Fulano de Tal"),
        ("texto", ""),
        ("texto", "TÍTULO DO TRABALHO DE EXEMPLO"),
        ("texto", ""),
        ("texto", "Trabalho de Conclusão de Curso apresentado como requisito parcial."),
        ("texto", ""),
        ("texto", "Orientador: Prof. Dr. Fulano"),
    ])
    resultado = formatar_capa_e_folha_de_rosto(d)

    assert resultado.titulos_formatados == 1
    assert resultado.autores_formatados == 1
    assert resultado.naturezas_formatadas == 1
    assert resultado.orientadores_formatados == 1

    bold, caps, _ = _bold_caps_size(d.paragraphs[0])
    assert bold and caps
    bold, caps, _ = _bold_caps_size(d.paragraphs[2])
    assert bold and caps
    bold_natureza, caps_natureza, _ = _bold_caps_size(d.paragraphs[4])
    assert bold_natureza is False and caps_natureza is False
    bold_orientador, caps_orientador, _ = _bold_caps_size(d.paragraphs[6])
    assert bold_orientador is False and caps_orientador is False


def test_projeto_final_tambem_reconhecido(construir_docx):
    """Variante real desta instituição: alguns documentos usam "Projeto
    Final" em vez de "Trabalho de Conclusão de Curso"."""
    d = construir_docx([
        ("texto", "Fulano de Tal"),
        ("texto", ""),
        ("texto", "TÍTULO DO TRABALHO DE EXEMPLO"),
        ("texto", ""),
        ("texto", "Projeto final apresentado em cumprimento às normas do curso."),
    ])
    resultado = formatar_capa_e_folha_de_rosto(d)
    assert resultado.naturezas_formatadas == 1
    assert resultado.titulos_formatados == 1
    assert resultado.autores_formatados == 1


def test_sem_ancora_nao_formata_nada(construir_docx):
    """Sem nenhuma âncora de texto fixo reconhecível, a ferramenta não deve
    arriscar adivinhar qual linha é o quê -- melhor não formatar do que
    formatar errado."""
    d = construir_docx([
        ("texto", "Fulano de Tal"),
        ("texto", ""),
        ("texto", "Um título qualquer sem nenhum marcador reconhecível por perto."),
    ])
    resultado = formatar_capa_e_folha_de_rosto(d)
    assert resultado.titulos_formatados == 0
    assert resultado.autores_formatados == 0
    assert resultado.tipos_documento_formatados == 0
    assert resultado.naturezas_formatadas == 0


def test_para_de_procurar_apos_o_primeiro_titulo_reconhecido(construir_docx):
    """A busca por âncoras de Capa/Folha de Rosto não deve avançar para
    dentro do corpo do trabalho -- se um parágrafo qualquer do corpo
    mencionar "orientador" ou citar um trabalho "apresentado", isso não
    pode ser confundido com a Folha de Rosto."""
    d = construir_docx([
        ("titulo_sem_numero", "SUMÁRIO"),
        ("heading1", "INTRODUÇÃO"),
        ("texto", "Projeto final apresentado por outro autor foi citado aqui no corpo."),
    ])
    resultado = formatar_capa_e_folha_de_rosto(d)
    assert resultado.naturezas_formatadas == 0
