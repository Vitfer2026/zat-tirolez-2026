# Agente de atualização semanal — Apresentação RSN/SSMA

Gera a apresentação semanal de Segurança, Saúde e Meio Ambiente (RSN) a
partir da planilha de ocorrências e da apresentação da semana anterior,
**sem alterar o layout do template** — só números, textos de resumo e a
quantidade de linhas/cartões que dependem do volume de dados da semana.

## Uso

```bash
pip install -r requirements.txt
python3 gerar_apresentacao.py \
    --planilha SSMA_2026_Dados_e_Ocorrencias_Unificados_v01.xlsx \
    --pptx-anterior Pre_RSN_Atualizada_v9_1.pptx \
    --saida Pre_RSN_Atualizada_v10.pptx
```

O intervalo da semana é inferido automaticamente a partir do `--pptx-anterior`
(veja a regra de fronteira abaixo). Para forçar um intervalo diferente:

```bash
python3 gerar_apresentacao.py ... --inicio 18/07/2026 --fim 24/07/2026
```

### Consolidado semanal opcional (`--semana`)

Se a planilha mestre ainda não tiver todos os lançamentos da semana atual
(ex.: um consolidado separado, feito à parte pelo time), passe-o com
`--semana`:

```bash
python3 gerar_apresentacao.py \
    --planilha SSMA_2026_Dados_e_Ocorrencias_Unificados_v01.xlsx \
    --semana Ocorrencias_Semana_Consolidado.xlsx \
    --pptx-anterior Pre_RSN_Atualizada_v10.pptx \
    --saida Pre_RSN_Atualizada_v11.pptx
```

Formato esperado (aba única, cabeçalho na linha 1): `Data, Unidade, Tipo,
Classificação, Descrição` — sem coluna Farol, sem separar Tirolez de
Levitare\|Regina (o bloco é inferido pela unidade: Levitare/Regina vão para
o bloco "Levitare\|Regina", o resto para "Tirolez"). As datas podem estar
como texto `dd/mm/aaaa` ou como data real do Excel.

Os dados desse arquivo **substituem** os de `--planilha` só para a semana
atual (datas ≥ início); `--planilha` continua sendo a única fonte para o
histórico anterior, necessário para o YTD do slide 4.

## Escopo (o que o agente atualiza)

| Slide | O que é atualizado |
|---|---|
| 1 — Capa | Intervalo de datas |
| 2 — Régua de Maturidade (metodologia) | **Nada** — é conteúdo de referência fixo |
| 3 — Termômetro de Cultura | Contagens por classificação, nível da régua (sugestão), quadro "Evolução por unidade" (redimensionado conforme unidades ativas), totais da semana, headline |
| 4 — Pirâmide de Segurança YTD | Tabela YTD completa (recomputada do zero a cada semana), KPIs, cartões de ACA/irreversível da semana (redimensionados), headline |
| 5 — Log de ocorrências | Tabela detalhada (linhas adicionadas/removidas conforme o volume real da semana, cor do "Farol" por classificação) |
| 6 — ETE (DQO/O₂/SD), se existir | **Nunca tocado.** Esta planilha não tem dados de ETE, então o agente não atualiza esse slide — mas também não o remove: se o `--pptx-anterior` tiver um slide 6 (mantido por outro processo/planilha), ele é copiado para a saída sem nenhuma alteração. Sem slide 6 no `--pptx-anterior`, a saída simplesmente continua com 5 slides. |

## Fonte de dados

Abas **"Tirolez"** e **"Levitare\|Regina"** da planilha, colunas `Data,
Unidade, Tipo, Classif., Descrição` (a partir da linha 5). As duas abas
juntas também definem os dois blocos da tabela do slide 5.

## Decisões de projeto (confirmadas com o time)

- **Semanas normais vão de sábado a sexta (7 dias) e nunca se sobrepõem.**
  A semana atual é sempre inferida como o sábado seguinte à sexta de
  fechamento da semana anterior. Toda ocorrência conta para a semana em
  que sua própria data cai — **sem excecões por classificação**: um ACA
  nos primeiros dias da semana atual é contado na semana atual, não é
  "adiantado" para a anterior. (Uma versão anterior deste agente tinha
  essa exceção, criada para um caso histórico único de puxar um ACA de
  sábado/domingo/segunda para o relatório que fechava. Ficou obsoleta
  quando "semana passada" passou a vir do que a apresentação anterior já
  publicou — ver seção abaixo — e continuar puxando escondia ACAs reais
  da semana atual do log, dos destaques e das contagens semanais, mesmo
  eles aparecendo corretamente no YTD/mês. Foi assim que o bug foi
  percebido e removido.)
- **Quadro "Evolução por unidade" (slide 3) e cartões de ACA (slide 4)
  encolhem** para mostrar só as unidades/eventos realmente ativos —
  não ficam vagas fixas com "–". Se o número de unidades ativas (soma das
  duas semanas) passar de 8, ou o número de ACA/irreversíveis da semana
  passar de 3, o script **para com erro** em vez de arriscar layout
  quebrado — nesse caso, adicione manualmente uma linha/cartão ao template
  antes de gerar.
- **Nível da Régua de Maturidade é uma sugestão automática**, calculada a
  partir das regras do slide 2 aplicadas às ocorrências da semana. Só o
  gatilho mais severo presente na semana vale (não soma vários), igual ao
  comportamento observado no template original (3 ACA + 1 ASA na mesma
  semana produziram apenas −0,5, não −0,7):
  - Acidente irreversível → reset para 1,0.
  - Senão, ACA na semana → −0,5 (teto).
  - Senão, ASA na semana → −0,2.
  - Senão (semana sem ACA e sem ASA) → **+0,3**, usando a ausência de
    acidentes como *proxy* de "IPS/IPA ≥ 90%" (proatividade sustentada),
    já que a planilha não traz o indicador real de IPS/IPA nem os
    quase-acidentes relatados. **Sempre revisar com o time de SSMA antes
    de publicar** — é uma aproximação, não o cálculo oficial completo.

## Textos "de rascunho" — revisar antes de publicar

As seguintes frases são geradas por template a partir dos dados (factuais,
mas sem o polimento editorial de quem escreve o resumo manualmente):
headline do slide 3, headline e subtítulo "OCORRÊNCIAS MAIS GRAVES" do
slide 4, e a nota "Irreversível: ... · +N ASA" do slide 3. Todas aparecem
no `stdout` do script para conferência rápida.

## "Semana passada" vem da apresentação anterior, não da planilha

No slide 3 (contagens por tipo, quadro "Evolução por unidade", GERAL e
"Totais da semana"), os números de **"semana passada" são lidos direto da
apresentação anterior** — especificamente do que ela publicou como "esta
semana" (`extract_prev_published`) — em vez de recalculados da planilha.

Isso é proposital: a planilha é um documento vivo e pode ser corrigida
depois que uma semana já foi publicada (reclassificar um Desvio como
Incidente, adicionar um lançamento retroativo). Se "semana passada" fosse
recalculado toda vez, a comparação semana-a-semana ficaria inconsistente
com o que o time já viu — o mesmo evento que na v10 subiu "6→5" apareceria
na v11 como "6→6", por exemplo, mesmo sem nada ter mudado de verdade. Só a
**"esta semana"** (a coluna nova) é calculada da planilha; o YTD do slide 4
também é recalculado do zero a cada rodada (isso é o esperado para um
total acumulado).

Consequência prática: **o `--pptx-anterior` passado ao script precisa ser
sempre a última apresentação realmente publicada**, não uma versão
descartada ou um rascunho — é dali que a próxima rodada lê a referência.

## Limitações conhecidas

- O slide 6 (ETE) nunca é atualizado por este agente (sem dados na
  planilha) — mas também nunca é removido. Se um slide 6 mantido por
  outro processo estiver no `--pptx-anterior`, ele passa para a saída
  intocado; ele já ficou desatualizado (ex.: título/datas de outra
  semana) até que o processo dele próprio o atualize separadamente.
- A validação visual (renderização das imagens do slide) depende do
  LibreOffice; em ambientes sem ele, valide com:
  ```bash
  python3 /root/.claude/skills/pptx/scripts/office/validate.py saida.pptx --original anterior.pptx
  ```
  e revisando o texto com `markitdown saida.pptx`.
