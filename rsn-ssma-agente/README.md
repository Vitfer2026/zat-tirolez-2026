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

## Escopo (o que o agente atualiza)

| Slide | O que é atualizado |
|---|---|
| 1 — Capa | Intervalo de datas |
| 2 — Régua de Maturidade (metodologia) | **Nada** — é conteúdo de referência fixo |
| 3 — Termômetro de Cultura | Contagens por classificação, nível da régua (sugestão), quadro "Evolução por unidade" (redimensionado conforme unidades ativas), totais da semana, headline |
| 4 — Pirâmide de Segurança YTD | Tabela YTD completa (recomputada do zero a cada semana), KPIs, cartões de ACA/irreversível da semana (redimensionados), headline |
| 5 — Log de ocorrências | Tabela detalhada (linhas adicionadas/removidas conforme o volume real da semana, cor do "Farol" por classificação) |
| 6 — ETE (DQO/O₂/SD) | **Removido da apresentação gerada** — esta planilha não tem dados de ETE. Se ele precisar entrar na apresentação final, é um passo manual separado (feito em outra cópia/processo, fora deste agente). |

## Fonte de dados

Abas **"Tirolez"** e **"Levitare\|Regina"** da planilha, colunas `Data,
Unidade, Tipo, Classif., Descrição` (a partir da linha 5). As duas abas
juntas também definem os dois blocos da tabela do slide 5.

## Decisões de projeto (confirmadas com o time)

- **Semanas normais vão de sábado a sexta (7 dias) e nunca se sobrepõem.**
  A semana atual é sempre inferida como o sábado seguinte à sexta de
  fechamento da semana anterior — mesmo que o `.pptx` anterior mostre uma
  data de fim que não é sexta (isso acontece por causa da exceção do ACA
  abaixo; o script volta para a sexta mais recente antes de calcular).
- **Exceção do ACA (sexta a segunda)**: um ACA que acontece entre sábado e
  segunda da semana nova é reportado de imediato no relatório que está
  sendo fechado (conta como "semana anterior", não espera o ciclo
  seguinte) — por isso um relatório pode aparecer com a data de fim
  "estendida" até esse ACA. O agente detecta esse ACA automaticamente
  pela data e pelo classificador "ACA", soma-o na contagem da semana
  anterior e **não o duplica** na semana atual — mesmo que a data dele já
  esteja dentro do intervalo novo. Todo ACA puxado aparece no `stdout` do
  script para conferência.
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

## Limitações conhecidas

- O slide 6 (ETE) é excluído da apresentação gerada — a lógica de
  atualizá-lo (DQO/O₂/SD/kgDQO por unidade) não faz parte deste agente,
  pois essa planilha não tem esses dados. A saída sempre tem 5 slides.
- A validação visual (renderização das imagens do slide) depende do
  LibreOffice; em ambientes sem ele, valide com:
  ```bash
  python3 /root/.claude/skills/pptx/scripts/office/validate.py saida.pptx --original anterior.pptx
  ```
  e revisando o texto com `markitdown saida.pptx`.
