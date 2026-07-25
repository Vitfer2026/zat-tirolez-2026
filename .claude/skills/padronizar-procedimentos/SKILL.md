---
name: padronizar-procedimentos
description: "Padroniza qualquer procedimento, POP, IT, manual, programa, tabela de higienização ou procedimento analítico (PAP) que tenha sido criado fora do padrão corporativo, ajustando layout, formatação e redação para o modelo oficial CORP-GQ-ANX-001 (TEMPLATE PARA DOCUMENTOS). Use quando o usuário enviar um documento .docx de procedimento fora do padrão pedindo para 'padronizar', 'ajustar ao modelo', 'corrigir formatação/layout', 'adequar ao template' ou similar — mesmo sem anexar o template, desde que o cache dele já exista em assets/. Cobre: fonte/margens/alinhamento, cor de destaque das tabelas, cabeçalho/rodapé/capa, codificação do documento, e a estrutura de seções obrigatória por tipo de documento (Manual/Procedimento/Programa, POP/IT, Tabela de Higienização, PAP, APPCC)."
license: Proprietary
---

# Padronização de procedimentos

## O que esta skill faz

Recebe qualquer documento de procedimento (POP, IT, Manual, Programa,
Tabela de Higienização, PAP) que não segue o padrão corporativo e o
reconstrói no formato oficial definido pelo anexo `CORP-GQ-ANX-001 —
TEMPLATE PARA DOCUMENTOS`, preservando **todo** o conteúdo técnico
original (passos, responsáveis, EPIs, imagens, tabelas embutidas,
documentos relacionados) — só a forma muda, nunca o conteúdo técnico.

O template oficial (r04) está salvo em
`assets/CORP-GQ-ANX-001_r04_template.docx`. As regras que ele documenta já
foram extraídas para os arquivos de referência abaixo — leia-os em vez de
reabrir o `.docx` toda vez:

- `references/formatacao.md` — fonte, margens, alinhamento, espaçamento,
  cor de destaque das tabelas, e o padrão de redação (verbos no
  infinitivo, maiúsculas em títulos de item, etc.).
- `references/codificacao.md` — tabela de localidade/área/tipo de
  documento para validar ou identificar o código do documento.
- `references/cabecalho-rodape-capa.md` — estrutura exata do cabeçalho,
  rodapé e capa.
- `references/escopos-por-tipo.md` — a estrutura de seções obrigatória
  para cada tipo de documento. **Este é o arquivo mais importante**: é
  ele que diz que seções o documento final precisa ter e em que ordem.

Se o usuário enviar uma versão mais nova do template (`CORP-GQ-ANX-001`
com revisão maior que r04), atualize o arquivo em `assets/` e revise os
arquivos de referência contra a nova versão antes de padronizar qualquer
documento com ela.

## Princípio central

**Padronizar é reorganizar e reformatar, não recriar do zero.** O
documento de entrada já contém o conhecimento técnico correto (como fazer
o queijo, como higienizar o equipamento, quais EPIs usar) — o trabalho é
mapear esse conteúdo para as seções e a formatação certas, não reescrever
o processo com base em suposições. Nunca invente, remova ou "melhore"
conteúdo técnico (quantidades, tempos, temperaturas, responsáveis) só
porque ele parece estranho — se algo parecer errado tecnicamente (não é
mais formatação, é o conteúdo), sinalize ao usuário em vez de corrigir
silenciosamente.

## Fluxo de trabalho

### Fase 1 — Diagnóstico

1. Ler o documento de entrada por completo (texto de parágrafos, todas as
   tabelas — inclusive tabelas aninhadas dentro de células — e contar
   imagens embutidas). Use `python-docx` para isso; não confie apenas no
   texto extraído por conversores que "achatam" tabelas, porque este
   padrão de documento depende muito de estrutura de tabela.
2. Identificar o tipo de documento (POP, IT, Manual, Procedimento,
   Programa, Tabela de Higienização, PAP) pelo código existente, pelo
   título do cabeçalho, ou pelo conteúdo — ver a seção "Como decidir qual
   escopo usar" em `references/escopos-por-tipo.md`.
3. Montar uma lista do que já está certo vs. o que está fora do padrão:
   fonte/tamanho errado, cor de tabela errada, cabeçalho/rodapé ausente
   ou incompleto, seções faltando ou fora de ordem, redação em estilo
   inconsistente, código de documento ausente/suspeito. Apresentar esse
   diagnóstico ao usuário antes de gerar o arquivo final, especialmente se
   houver ambiguidade sobre o tipo de documento ou o código correto.

### Fase 2 — Reconstrução

Trabalhar por edição direta do XML do `.docx` (abordagem "editar
documento existente" da skill `docx`: `unzip` → editar `word/document.xml`
→ `zip`), não recriando o arquivo do zero em `docx-js` — isso preserva
imagens, tabelas aninhadas e relações internas do documento original sem
precisar re-inserir cada imagem manualmente.

1. **Cabeçalho/rodapé/capa**: reconstruir conforme
   `references/cabecalho-rodape-capa.md`, preenchendo com os metadados
   reais do documento (código, revisão, datas, título, unidade). Se o
   documento de entrada já tiver essas informações só que malformatadas,
   extrair os valores de lá — não inventar datas ou números de revisão.
2. **Formatação**: aplicar as regras de `references/formatacao.md` (fonte
   Arial, tamanhos 12/9, margens 1,5 cm, alinhamento, cor `#DBE5F1` no
   cabeçalho de toda tabela, espaçamento).
3. **Estrutura de seções**: mapear o conteúdo existente para as seções do
   escopo correto (`references/escopos-por-tipo.md`). Ao mover o conteúdo
   de uma etapa que contém imagem ou tabela aninhada, mover o nó XML
   inteiro (parágrafo/tabela com seus `w:drawing`/relações), não
   retranscrever a imagem como texto.
4. **Redação**: ajustar verbos para infinitivo/imperativo, títulos de
   item para maiúsculo, ordenar Definições e Documentos de Referência
   conforme pedido no escopo, sem alterar o significado técnico.
5. Se o documento original não tiver alguma seção obrigatória do escopo
   (ex.: falta "Registros" ou "Verificação"), criar a seção com o rótulo
   padrão e sinalizar ao usuário que o conteúdo precisa ser preenchido —
   não inventar o conteúdo da seção.

### Fase 3 — Validação

1. `python scripts/office/validate.py <saida.docx> --original <entrada.docx>`
   (script da skill `docx`) para checar XML contra o schema OOXML. **Isso
   não é suficiente sozinho** — passar no XSD não garante que o Word vai
   conseguir abrir o arquivo (ver próximo item).
2. Reabrir o `.docx` gerado com `python-docx` e conferir: contagem de
   imagens preservada, todas as seções do escopo presentes na ordem
   certa, cabeçalho/rodapé com os campos certos, cor de tabela `#DBE5F1`
   aplicada nos cabeçalhos de coluna. Se o cabeçalho/rodapé usar campos
   dinâmicos (`PAGE`/`NUMPAGES`), verificar que cada `<w:fldChar>` e
   `<w:instrText>` está em um `<w:r>` próprio, nunca todos empacotados no
   mesmo run — isso passa no XSD mas já causou o cabeçalho inteiro
   sumir ao abrir no Word de verdade (ver `references/cabecalho-rodape-capa.md`).
   Buscar por esse padrão especificamente após qualquer alteração no
   código que gera campos.
3. Tentar renderizar com `soffice`/`pdftoppm` (ver skill `docx`) e olhar
   as páginas — isso é o ideal para pegar problemas visuais que a
   inspeção de XML não mostra. **Se o ambiente não conseguir converter
   para PDF** (algumas sandboxes têm o LibreOffice quebrado — teste com um
   `.docx` trivial antes de assumir que o problema é do arquivo), avisar
   explicitamente o usuário que a verificação visual não pôde ser feita
   neste ambiente e recomendar abrir o arquivo no Word antes de publicar.

## Postura

- **Diagnostique e mostre o plano antes de gerar o arquivo final**,
  principalmente quando o tipo de documento ou o código não forem óbvios.
- **Nunca altere conteúdo técnico** (quantidades, tempos, temperaturas,
  responsáveis, nomes de produtos/documentos relacionados) — só a forma.
- **Nunca invente código de documento, data de aprovação ou nome de
  aprovador** — se faltar, sinalizar a lacuna.
- **Preserve imagens e tabelas aninhadas** movendo o XML original, nunca
  descrevendo a imagem em texto ou omitindo-a.
