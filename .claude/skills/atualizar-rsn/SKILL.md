---
name: atualizar-rsn
description: Gera a apresentação semanal RSN/SSMA (Segurança, Saúde e Meio Ambiente) a partir da planilha de ocorrências ("Tirolez" + "Levitare|Regina") e da apresentação da semana anterior, sem alterar o layout. Use quando o usuário pedir para atualizar/gerar a apresentação RSN, o relatório semanal de SSMA, ou mencionar as abas Tirolez/Levitare|Regina junto com um .pptx anterior.
---

# Atualizar apresentação RSN/SSMA

Este skill executa `rsn-ssma-agente/gerar_apresentacao.py` para produzir a
apresentação semanal a partir de dois insumos que o usuário deve fornecer:

1. A planilha de ocorrências (abas "Tirolez" e "Levitare|Regina").
2. A apresentação `.pptx` da semana anterior.

Leia `rsn-ssma-agente/README.md` primeiro — ele documenta o escopo exato
(quais slides são tocados), as decisões de projeto já validadas com o
time (semanas sem sobreposição, quadros que encolhem em vez de manter
vagas vazias, régua de maturidade como sugestão automática) e os limites
em que o script para com erro em vez de arriscar quebrar o layout
(mais de 8 unidades ativas, mais de 3 ACA/irreversíveis na semana).

## Passos

1. Instalar dependências se necessário: `pip install -r rsn-ssma-agente/requirements.txt`.
2. Rodar:
   ```bash
   python3 rsn-ssma-agente/gerar_apresentacao.py \
       --planilha <planilha.xlsx> \
       --pptx-anterior <semana_anterior.pptx> \
       --saida <nome_da_saida.pptx>
   ```
3. Ler a saída do script no terminal — ela lista o intervalo de datas
   inferido, a sugestão de nível da régua de maturidade e os motivos.
   **Sinalizar ao usuário que a régua e as frases de resumo (headlines)
   são sugestões automáticas e precisam de revisão do time de SSMA antes
   de publicar** — isso está documentado no README.
4. Validar a saída:
   ```bash
   python3 /root/.claude/skills/pptx/scripts/office/validate.py <saida.pptx> --original <semana_anterior.pptx>
   markitdown <saida.pptx>
   ```
   Revisar o texto extraído por `markitdown` procurando por conteúdo
   estranho ou datas erradas antes de entregar.
5. Se o script parar com erro de capacidade excedida (mais unidades
   ativas ou mais ACA/irreversíveis do que o template suporta), **não
   tente contornar automaticamente** — avisar o usuário que o template
   precisa de um ajuste manual (adicionar linha/cartão) antes de rodar
   de novo.
6. O slide 6 (ETE/DQO) nunca é tocado por este agente — se a planilha de
   ETE estiver disponível, isso é um passo manual separado (ou uma
   extensão futura do script, fora do escopo atual).
