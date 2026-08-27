# Formatador de TCC (ABNT)

Ferramenta para formatar automaticamente o `.docx` do TCC de um aluno
segundo as normas ABNT usadas pela instituição, com base no
`templates/Modelo_TCC_oficial.docx` (documento oficial fornecido). A
ferramenta **nunca reescreve o texto do aluno** -- ela só toca em
propriedades de formatação (fonte, tamanho, negrito/itálico, alinhamento,
espaçamento, recuo, margens) e gera um relatório apontando o que foi
corrigido automaticamente e o que ainda precisa de revisão manual.

## O que ela faz

1. **Normaliza a formatação** de cada parágrafo (fonte Times New Roman 12,
   espaçamento 1,5, recuo de primeira linha 1,25cm, margens 3/2/3/2cm,
   títulos em negrito/maiúsculas conforme o nível, legendas de
   figura/quadro/tabela, citações diretas longas com recuo de 4cm etc.),
   classificando cada parágrafo pelo estilo do Word que o aluno usou (se
   partiu do modelo oficial) ou por heurísticas de texto (título numerado
   digitado manualmente, palavras-chave como RESUMO/ABSTRACT/REFERÊNCIAS).
   Ver `formatador_tcc/classify.py` e `formatador_tcc/config.py`.

2. **Reconstrói o SUMÁRIO** como um campo TOC nativo do Word (`{ TOC }`),
   em vez de tentar calcular números de página manualmente -- o próprio
   Word recalcula a paginação correta ao abrir o arquivo (o documento é
   configurado para atualizar campos automaticamente). Ver
   `formatador_tcc/sumario.py`.

3. **Confere a contagem de palavras** do Resumo e do Abstract contra a
   regra explícita do modelo oficial (mínimo 150, máximo 500 palavras).
   Ver `formatador_tcc/contagem.py`.

4. **Cruza citações do texto com a lista de Referências** (NBR 10520):
   aponta citações no corpo do texto sem entrada correspondente nas
   Referências, e referências da lista que nunca foram citadas. Reconhece
   citação parentética (`(SILVA, 2020)`), narrativa (`Silva (2020)`),
   múltiplos autores, "et al." e "apud" (nesse caso só exige referência da
   fonte efetivamente consultada). Ver `formatador_tcc/citacoes.py`.

5. **Confere a formatação de cada referência** (NBR 6023): autor em
   maiúsculas seguido de vírgula ou ponto, presença de ano, pontuação
   final, alinhamento à esquerda, recuo zero, espaçamento simples, e
   consistência do destaque tipográfico do título (negrito OU itálico, não
   os dois ao longo da lista). Ver `formatador_tcc/referencias_check.py`.

## Uso

```bash
pip install -r requirements.txt
python -m formatador_tcc caminho/do/tcc_do_aluno.docx
```

Isso gera, ao lado do arquivo de entrada:

- `tcc_do_aluno_formatado.docx` -- o documento com a formatação normalizada;
- `tcc_do_aluno_relatorio.md` -- o relatório de conferência.

Parâmetros opcionais:

```bash
python -m formatador_tcc entrada.docx --saida saida.docx --relatorio relatorio.md
```

## O que é corrigido automaticamente vs. o que exige revisão do aluno

| Corrigido automaticamente | Reportado para revisão manual |
|---|---|
| Fonte, tamanho, negrito/maiúsculas dos títulos | Autor de referência fora do padrão (ex.: minúsculas) |
| Alinhamento, espaçamento, recuo de parágrafo | Ano ausente ou pontuação final ausente numa referência |
| Margens e tamanho de página | Destaque de título inconsistente entre referências |
| Entradas antigas do sumário → campo TOC nativo | Citação sem referência correspondente |
| Recuo de citação direta longa (heurística) | Referência nunca citada no texto |
| | Resumo/Abstract fora do intervalo de 150–500 palavras |

O motivo de não corrigir a segunda coluna automaticamente é que fazer isso
exigiria **decidir por conta própria** o que o aluno quis dizer (ex.:
inventar um ano de publicação, ou decidir sozinho qual autor está certo) --
o que violaria a regra de não alterar o conteúdo escrito pelo aluno.

## Limitações conhecidas

- A detecção de citações é heurística (baseada em expressões regulares).
  Ela cobre os formatos mais comuns da NBR 10520, mas pode não reconhecer
  formatos incomuns ou citações mal formatadas pelo próprio aluno -- por
  isso o relatório sempre pede revisão humana antes de considerar a lista
  definitiva.
- A detecção automática de "citação direta longa" (recuo de 4cm) é
  conservadora (só age quando o parágrafo inteiro está entre aspas e é
  longo); citações longas sem aspas não são identificadas.
- A checagem de contagem de palavras hoje cobre apenas Resumo/Abstract,
  porque é a única regra numérica explícita no modelo oficial. Se a
  instituição tiver uma faixa de páginas/palavras exigida para o trabalho
  inteiro, ela pode ser adicionada em `formatador_tcc/config.py` e
  `formatador_tcc/contagem.py`.
- Testado neste ambiente apenas por validação estrutural do XML/zip e por
  comparação de texto antes/depois (não há LibreOffice funcional neste
  sandbox para gerar um PDF de conferência visual) -- recomenda-se abrir o
  `_formatado.docx` gerado no Word real antes de considerar o processo
  concluído.

## Rodando os testes

```bash
pip install -r requirements-dev.txt
python -m pytest tests/ -v
```

## Próximos passos sugeridos

- Empacotar como serviço web (ex.: reaproveitando o padrão do projeto de
  Ata de Defesa já existente: formulário → upload do `.docx` → download do
  `.docx` formatado + relatório), já que os alunos escrevem no Word e não
  devem precisar rodar um script manualmente.
- Ampliar `referencias_check.py` com validações específicas por tipo de
  referência (livro, artigo, site, norma, legislação), hoje tratadas de
  forma genérica.
- Se a instituição definir uma faixa de páginas/palavras para o trabalho
  completo, adicionar essa checagem em `contagem.py`.
