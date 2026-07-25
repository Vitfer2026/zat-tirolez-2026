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

O intervalo da semana é inferido automaticamente (semana anterior lida do
próprio `--pptx-anterior`, semana atual = dia seguinte até +7 dias, sem
sobreposição). Para forçar um intervalo diferente:

```bash
python3 gerar_apresentacao.py ... --inicio 19/07/2026 --fim 25/07/2026
```

## Escopo (o que o agente atualiza)

| Slide | O que é atualizado |
|---|---|
| 1 — Capa | Intervalo de datas |
| 2 — Régua de Maturidade (metodologia) | **Nada** — é conteúdo de referência fixo |
| 3 — Termômetro de Cultura | Contagens por classificação, nível da régua (sugestão), quadro "Evolução por unidade" (redimensionado conforme unidades ativas), totais da semana, headline |
| 4 — Pirâmide de Segurança YTD | Tabela YTD completa (recomputada do zero a cada semana), KPIs, cartões de ACA/irreversível da semana (redimensionados), headline |
| 5 — Log de ocorrências | Tabela detalhada (linhas adicionadas/removidas conforme o volume real da semana, cor do "Farol" por classificação) |
| 6 — ETE (DQO/O₂/SD) | **Não tocado** — esta planilha não tem dados de ETE. Continua sendo atualizado manualmente. |

## Fonte de dados

Abas **"Tirolez"** e **"Levitare\|Regina"** da planilha, colunas `Data,
Unidade, Tipo, Classif., Descrição` (a partir da linha 5). As duas abas
juntas também definem os dois blocos da tabela do slide 5.

## Decisões de projeto (confirmadas com o time)

- **Semanas não se sobrepõem**: a semana atual começa no dia seguinte ao
  fim da semana anterior (ex.: anterior 11/07–18/07 → atual 19/07–25/07).
- **Quadro "Evolução por unidade" (slide 3) e cartões de ACA (slide 4)
  encolhem** para mostrar só as unidades/eventos realmente ativos —
  não ficam vagas fixas com "–". Se o número de unidades ativas (soma das
  duas semanas) passar de 8, ou o número de ACA/irreversíveis da semana
  passar de 3, o script **para com erro** em vez de arriscar layout
  quebrado — nesse caso, adicione manualmente uma linha/cartão ao template
  antes de gerar.
- **Nível da Régua de Maturidade é uma sugestão automática**, calculada a
  partir das regras do slide 2 aplicadas às ocorrências da semana (ACA →
  −0,5 teto; ASA isolado → −0,2; reincidência do mesmo desvio/unidade →
  −0,3; autuação/notificação de terceiro citada na descrição → −0,2;
  acidente irreversível → reset para 1,0). **Os gatilhos de alta (IPS/IPA
  ≥ 90%, quase-acidente relatado) não estão na planilha e por isso nunca
  são aplicados** — o nível só cai ou fica estável nesta versão do agente.
  **Sempre revisar com o time de SSMA antes de publicar.**

## Textos "de rascunho" — revisar antes de publicar

As seguintes frases são geradas por template a partir dos dados (factuais,
mas sem o polimento editorial de quem escreve o resumo manualmente):
headline do slide 3, headline e subtítulo "OCORRÊNCIAS MAIS GRAVES" do
slide 4, e a nota "Irreversível: ... · +N ASA" do slide 3. Todas aparecem
no `stdout` do script para conferência rápida.

## Limitações conhecidas

- Slide 6 (ETE) precisa de outra fonte de dados (DQO/O₂/SD/kgDQO por
  unidade) — não incluída nesta planilha. Continua manual.
- A validação visual (renderização das imagens do slide) depende do
  LibreOffice; em ambientes sem ele, valide com:
  ```bash
  python3 /root/.claude/skills/pptx/scripts/office/validate.py saida.pptx --original anterior.pptx
  ```
  e revisando o texto com `markitdown saida.pptx`.
