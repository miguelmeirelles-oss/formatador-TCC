from formatador_tcc.contagem import verificar_contagem_resumo


def test_resumo_abaixo_do_minimo(construir_docx):
    d = construir_docx([
        ("titulo_sem_numero", "RESUMO"),
        ("texto", "Um resumo muito curto."),
        ("texto", "Palavras-chave: Um; Dois."),
        ("titulo_sem_numero", "ABSTRACT"),
        ("texto", " ".join(["word"] * 200)),
        ("titulo_sem_numero", "SUMÁRIO"),
        ("heading1", "INTRODUÇÃO"),
    ])
    rel = verificar_contagem_resumo(d)
    assert rel.resumo.encontrada is True
    assert rel.resumo.dentro_do_limite is False
    assert rel.abstract.dentro_do_limite is True


def test_resumo_dentro_do_limite_nao_conta_palavras_chave(construir_docx):
    d = construir_docx([
        ("titulo_sem_numero", "RESUMO"),
        ("texto", " ".join(["palavra"] * 150)),
        ("texto", "Palavras-chave: " + " ".join(["chave"] * 100)),
        ("titulo_sem_numero", "ABSTRACT"),
        ("texto", " ".join(["word"] * 150)),
        ("titulo_sem_numero", "SUMÁRIO"),
    ])
    rel = verificar_contagem_resumo(d)
    assert rel.resumo.palavras == 150
    assert rel.resumo.dentro_do_limite is True


def test_resumo_nao_encontrado(construir_docx):
    d = construir_docx([
        ("heading1", "INTRODUÇÃO"),
        ("texto", "Texto qualquer."),
    ])
    rel = verificar_contagem_resumo(d)
    assert rel.resumo.encontrada is False
    assert rel.resumo.dentro_do_limite is False
