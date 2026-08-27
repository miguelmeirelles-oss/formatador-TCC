# Formatador de TCC (ABNT)

Ferramenta para formatar automaticamente o `.docx` do TCC de um aluno de
Engenharia de Alimentos do CEFET/RJ segundo as normas da instituição, com
base em duas fontes oficiais: o `templates/Modelo_TCC_oficial.docx` (fonte
principal -- usado sempre que ele define a regra) e o **Apêndice I**
("Normas para Elaboração e Defesa de Projeto Final") como complemento,
usado só quando o modelo oficial não cobre algo. Nenhuma regra usada aqui
foi inventada -- todas vêm de um desses dois documentos, ou de comentários
deixados pela própria autora do modelo oficial dentro do arquivo. A
ferramenta **nunca reescreve o texto do aluno** -- ela só toca em
propriedades de formatação (fonte, tamanho, negrito/itálico, alinhamento,
espaçamento, recuo, margens, quebras de página, numeração) e gera um
relatório apontando o que foi corrigido automaticamente e o que ainda
precisa de revisão manual.

## O que ela faz

1. **Normaliza a formatação** de cada parágrafo (fonte Times New Roman 12,
   espaçamento 1,5, recuo de primeira linha 1,25cm, margens 3/2/3/2cm,
   legendas de figura/quadro/tabela, citações diretas longas com recuo de
   4cm, conteúdo interno de tabelas em 10pt etc.), classificando cada
   parágrafo pelo estilo do Word que o aluno usou (se partiu do modelo
   oficial) ou por heurísticas de texto (título numerado digitado
   manualmente, palavras-chave como RESUMO/ABSTRACT/REFERÊNCIAS). Suporta os
   5 níveis de título do Apêndice I (seção primária a quinária), cada um com
   seu padrão de negrito/maiúsculas/itálico, e insere quebra de página antes
   de cada seção primária ("As seções primárias devem iniciar SEMPRE em
   páginas distintas" -- Apêndice I) -- inclusive removendo quebras de
   página manuais que o aluno tenha colocado em outro lugar, já que ficam
   redundantes ou erradas com a quebra automática. Também tem uma proteção
   contra estilo de título aplicado por engano (comum ao colar conteúdo de
   outro documento): um parágrafo comprido demais ou uma legenda de
   figura/tabela com estilo "Heading N" não é tratado como título, e perde
   esse estilo (senão o Sumário nativo do Word continuaria listando esse
   parágrafo como se fosse um capítulo). Ver `formatador_tcc/classify.py` e
   `formatador_tcc/config.py`.

2. **Reconstrói o SUMÁRIO** como um campo TOC nativo do Word (`{ TOC }`),
   em vez de tentar calcular números de página manualmente -- o próprio
   Word recalcula a paginação correta ao abrir o arquivo (o documento é
   configurado para atualizar campos automaticamente). Ver
   `formatador_tcc/sumario.py`.

3. **Numera as páginas** no canto superior direito a partir da Introdução,
   como manda o Apêndice I ("Todas as folhas ... devem ser contadas
   sequencialmente, mas não numeradas [até a Introdução] ... no canto
   superior direito da folha"). Usa um campo de fórmula nativo do Word
   (`{ =PAGE-2 }`, onde o "-2" desconta Capa e Ficha Catalográfica da
   contagem, conforme o próprio modelo oficial detalha em um comentário) --
   pelo mesmo motivo do Sumário, é o Word quem recalcula o número certo ao
   abrir o arquivo, não há como calcular isso sem um motor de paginação de
   verdade. Cabeçalho é uma propriedade de *seção* do Word, não de página --
   por isso, se o aluno não tiver preservado as quebras de seção do modelo
   oficial (comum: o trabalho inteiro numa seção só), a ferramenta insere a
   quebra de seção que falta antes da Introdução automaticamente. Ver
   `formatador_tcc/paginacao.py`.

4. **Confere a contagem de palavras** do Resumo e do Abstract contra a
   regra explícita do modelo oficial (mínimo 150, máximo 500 palavras).
   Ver `formatador_tcc/contagem.py`.

5. **Cruza citações do texto com a lista de Referências** (NBR 10520):
   aponta citações no corpo do texto sem entrada correspondente nas
   Referências, e referências da lista que nunca foram citadas. Reconhece
   citação parentética (`(SILVA, 2020)`), narrativa (`Silva (2020)`),
   múltiplos autores, "et al." e "apud" (nesse caso só exige referência da
   fonte efetivamente consultada). Ver `formatador_tcc/citacoes.py`.

6. **Confere a formatação de cada referência** (NBR 6023): autor em
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
| Fonte, tamanho, negrito/maiúsculas/itálico dos 5 níveis de título | Autor de referência fora do padrão (ex.: minúsculas) |
| Alinhamento, espaçamento, recuo de parágrafo | Ano ausente ou pontuação final ausente numa referência |
| Margens e tamanho de página | Destaque de título inconsistente entre referências |
| Quebra de página antes de cada seção primária | Citação sem referência correspondente |
| Entradas antigas do sumário → campo TOC nativo | Referência nunca citada no texto |
| Numeração de página (canto superior direito, a partir da Introdução) | Resumo/Abstract fora do intervalo de 150–500 palavras |
| Recuo de citação direta longa (heurística) | Casos de figura/tabela que a heurística de legenda não reconheceu |
| Fonte do conteúdo interno de tabelas (10pt) | |
| Quebras de página manuais redundantes/erradas removidas | |
| Estilo de título (Heading N) tirado de parágrafo/legenda que herdou por engano | |

O motivo de não corrigir a segunda coluna automaticamente é que fazer isso
exigiria **decidir por conta própria** o que o aluno quis dizer (ex.:
inventar um ano de publicação, ou decidir sozinho qual autor está certo) --
o que violaria a regra de não alterar o conteúdo escrito pelo aluno.

## Limitações conhecidas

- O documento final ainda pode conter comentários do Word herdados do
  modelo oficial (aquelas anotações da autora com orientações, tipo
  "Esta é a primeira página numerada..."), se o aluno não os removeu ao
  preencher o próprio modelo. A ferramenta não tenta apagá-los
  automaticamente -- mexer no XML de comentários tem risco real de corromper
  o arquivo, e não há como testar isso de forma confiável neste ambiente
  (sem Word/LibreOffice funcional para validar visualmente). Se algum
  comentário sobrar no arquivo final, é rápido apagar manualmente no Word
  (Revisão → Excluir → Excluir Todos os Comentários no Documento).
- A detecção de citações é heurística (baseada em expressões regulares).
  Ela cobre os formatos mais comuns da NBR 10520, mas pode não reconhecer
  formatos incomuns ou citações mal formatadas pelo próprio aluno -- por
  isso o relatório sempre pede revisão humana antes de considerar a lista
  definitiva.
- A detecção automática de "citação direta longa" (recuo de 4cm) é
  conservadora (só age quando o parágrafo inteiro está entre aspas e é
  longo); citações longas sem aspas não são identificadas.
- A proteção contra estilo de título aplicado por engano (ver seção 1 acima)
  usa uma heurística de tamanho: um parágrafo com estilo Heading 1-5 e mais
  de 15 palavras é tratado como corpo de texto, não título. Isso não é uma
  regra ABNT, é só uma heurística de segurança (`_LIMITE_PALAVRAS_TITULO` em
  `classify.py`) -- na prática rara de o trabalho ter um título de seção
  genuinamente longo, ele seria rebaixado incorretamente para corpo de
  texto; o valor pode ser ajustado se isso acontecer.
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

## Interface web

Para os alunos não precisarem rodar nada em linha de comando, há um app web
mínimo em `webapp/` (Flask): o aluno sobe o `.docx`, a página mostra o
relatório de conferência e oferece o botão de download do `.docx`
formatado. Tudo roda em memória (o motor é o mesmo do CLI, via
`formatador_tcc.pipeline`) -- nada é gravado em disco.

```bash
pip install -r webapp/requirements.txt
python -m webapp.app
```

Abre em `http://127.0.0.1:5000`.

### Colocando no ar para os alunos (deploy no Render, grátis)

O jeito mais simples de dar aos alunos um link público, sem precisar
instalar nada em servidor nenhum manualmente, é o [Render](https://render.com)
(tem plano gratuito). O repositório já vem pronto para isso (`Procfile` +
`webapp/requirements.txt`). Passo a passo:

1. Crie uma conta em [render.com](https://render.com) (dá para entrar direto
   com a conta do GitHub que já tem acesso a este repositório).
2. No painel do Render, clique em **New +** → **Web Service**.
3. Escolha **Build and deploy from a Git repository** e conecte o repositório
   `miguelmeirelles-oss/formatador-TCC` (o Render vai pedir autorização para
   acessar seus repositórios do GitHub -- autorize).
4. Preencha:
   - **Name**: qualquer nome (ex.: `formatador-tcc`).
   - **Branch**: `claude/tcc-formatter-program-1j4zbt` (ou `main`, depois que
     o PR for aceito).
   - **Runtime**: `Python 3`.
   - **Build Command**: `pip install -r requirements.txt -r webapp/requirements.txt`
   - **Start Command**: `gunicorn --workers 1 --bind 0.0.0.0:$PORT webapp.app:app`
   - **Instance Type**: `Free`.
5. Clique em **Create Web Service** e aguarde o build (leva alguns minutos
   na primeira vez).
6. Quando terminar, o Render mostra uma URL pública tipo
   `https://formatador-tcc.onrender.com` -- esse é o link que os alunos vão
   acessar. Não precisa mexer em mais nada.

**Sobre o plano gratuito:** o serviço "dorme" depois de um tempo sem uso e
demora ~30-60 segundos para acordar no primeiro acesso do dia -- aceitável
para uso departamental, mas se isso incomodar dá para migrar para um plano
pago do Render (ou outro provedor) sem mudar nada no código.

**Por que `--workers 1`:** o link de download fica guardado em memória do
processo que gerou o arquivo; com mais de um worker, um pedido de download
poderia cair num processo diferente do que processou o upload e falhar. Não
aumente os workers sem antes trocar esse armazenamento por algo
compartilhado (arquivo temporário em disco compartilhado, S3 etc.) em
`webapp/app.py` -- é uma mudança pequena e isolada, não afeta o motor de
formatação.

### Embutindo na página do Wix

O app não envia cabeçalho `X-Frame-Options` nem `Content-Security-Policy`,
então pode ser embutido normalmente num `<iframe>` dentro do Wix (mesmo
padrão usado no formulário de Ata de Defesa). No editor do Wix: adicione um
elemento **Embed → Custom Embed → Embed a Widget/iframe** na página
`projeto-final-tcc`, e aponte para a URL pública do Render (ex.:
`https://formatador-tcc.onrender.com`).

Se preferir manter tudo dentro do Google Workspace/Wix como o projeto de
Ata de Defesa que vocês já têm, dá para reaproveitar o mesmo formulário
HTML como frontend e usar `UrlFetchApp` do Apps Script para chamar esse
backend Python hospedado à parte (o Apps Script sozinho, via
`DocumentApp`, não consegue reproduzir o nível de controle de formatação
que este motor faz).

## Rodando os testes

```bash
pip install -r requirements-dev.txt
python -m pytest tests/ -v
```

## Fontes das regras usadas

- `templates/Modelo_TCC_oficial.docx` -- fonte principal, tem prioridade
  sempre que define a regra (estilos, margens, contagem de palavras do
  resumo etc.), incluindo comentários do Word deixados pela própria autora
  do modelo (ex.: a explicação de como numerar a página da Introdução).
- `templates/Apendice_I_normas_formatacao.txt` -- texto extraído do
  Apêndice I ("Normas para Elaboração e Defesa de Projeto Final", CEFET/RJ),
  usado só para preencher lacunas que o modelo oficial não cobre (ex.:
  numeração de página, seções quaternária/quinária). O arquivo original é
  um `.doc` binário antigo (Word 97-2003); o texto foi extraído com
  `antiword` porque não há LibreOffice funcional neste ambiente para
  converter para `.docx`.

## Próximos passos sugeridos

- Ampliar `referencias_check.py` com validações específicas por tipo de
  referência (livro, artigo, site, norma, legislação), hoje tratadas de
  forma genérica -- assim que houver mais detalhes das regras da
  instituição para cada tipo.
- Se a instituição definir uma faixa de páginas/palavras para o trabalho
  completo, adicionar essa checagem em `contagem.py`.
- Remoção segura de comentários do Word herdados do modelo (ver
  "Limitações conhecidas").
