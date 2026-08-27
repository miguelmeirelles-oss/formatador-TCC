from formatador_tcc.citacoes import extrair_citacoes_do_paragrafo, cruzar_citacoes_e_referencias


def test_citacao_parentetica_simples():
    cs = extrair_citacoes_do_paragrafo("Um trecho qualquer (SILVA, 2020).", 0)
    assert len(cs) == 1
    assert cs[0].autores == ("SILVA",)
    assert cs[0].ano == "2020"
    assert cs[0].forma == "parentetica"


def test_citacao_parentetica_dois_autores():
    cs = extrair_citacoes_do_paragrafo("(SILVA; SOUZA, 2020)", 0)
    assert len(cs) == 1
    assert cs[0].autores == ("SILVA", "SOUZA")


def test_citacao_parentetica_duas_obras():
    cs = extrair_citacoes_do_paragrafo("(SILVA, 2019; SOUZA, 2020)", 0)
    assert len(cs) == 2
    assert (cs[0].autores, cs[0].ano) == (("SILVA",), "2019")
    assert (cs[1].autores, cs[1].ano) == (("SOUZA",), "2020")


def test_citacao_et_al():
    cs = extrair_citacoes_do_paragrafo("(SILVA et al., 2020)", 0)
    assert len(cs) == 1
    assert cs[0].autores == ("SILVA",)
    assert cs[0].et_al is True


def test_citacao_narrativa():
    cs = extrair_citacoes_do_paragrafo("Segundo Silva (2020), blá blá.", 0)
    assert len(cs) == 1
    assert cs[0].autores == ("SILVA",)
    assert cs[0].forma == "narrativa"


def test_citacao_narrativa_et_al():
    cs = extrair_citacoes_do_paragrafo("Conforme Silva et al. (2020)...", 0)
    assert len(cs) == 1
    assert cs[0].et_al is True


def test_citacao_apud_parentetica():
    cs = extrair_citacoes_do_paragrafo("(Fulano apud SILVA, 2020)", 0)
    assert len(cs) == 1
    assert cs[0].autores == ("SILVA",)
    assert cs[0].apud is True


def test_nao_confunde_norma_tecnica_com_citacao():
    cs = extrair_citacoes_do_paragrafo("Ver especificação (NBR 14720/2002).", 0)
    assert cs == []


def test_nao_confunde_referencia_de_item_com_citacao():
    cs = extrair_citacoes_do_paragrafo("Os resultados (ver item 3.2) confirmam a hipótese.", 0)
    assert cs == []


def test_cruzamento_completo(construir_docx):
    d = construir_docx([
        ("titulo_sem_numero", "SUMÁRIO"),
        ("heading1", "INTRODUÇÃO"),
        ("texto", "Conforme Silva (2020), o tema é relevante. Outro ponto (Souza; Lima, 2019) "
                  "reforça isso. Um terceiro autor (Gomes, 2022) não tem referência."),
        ("heading1", "REFERÊNCIAS"),
        ("texto", "SILVA, João. Um livro qualquer. São Paulo: Editora, 2020."),
        ("texto", "SOUZA, Maria; LIMA, Pedro. Outro livro. Rio de Janeiro: Editora, 2019."),
        ("texto", "OLIVEIRA, Ana. Nunca citada no texto. Curitiba: Editora, 2018."),
    ])
    res = cruzar_citacoes_e_referencias(d)

    assert len(res.citacoes) == 3
    assert len(res.referencias) == 3

    orfas = {c.autores for c in res.citacoes_sem_referencia}
    assert orfas == {("GOMES",)}

    nao_citadas = {r.autores for r in res.referencias_sem_citacao}
    assert nao_citadas == {("OLIVEIRA",)}
