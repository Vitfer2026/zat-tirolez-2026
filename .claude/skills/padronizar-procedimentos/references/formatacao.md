# Regras de formatação (fonte: CORP-GQ-ANX-001, item 1)

Extraído da tabela oficial do template. Aplicar por formatação direta nos
parágrafos/tabelas — os estilos "Normal" do .docx corporativo **não** têm
essas propriedades embutidas, então cada run precisa da formatação direta.

| Item | Texto corrido | Dentro de tabela |
|---|---|---|
| Margens (todas) | 1,5 cm | 1,5 cm |
| Layout cabeçalho/rodapé (distância) | 1,25 cm | 1,25 cm |
| Tipo de letra | Arial | Arial |
| Tamanho | 12 pt | 9 pt |
| Cor da fonte | Automático (preto) | Automático (preto) |
| Títulos e subtítulos | 12 pt, negrito, MAIÚSCULO | — |
| Alinhamento do texto | Justificado | Centralizado |
| Alinhamento de figuras/tabelas | Centralizado | Centralizado |
| Espaçamento entre parágrafos | 0 pt | 0 pt |
| Espaçamento entre linhas | Simples 1,5 | Simples 1,15 |
| Recuo | Primeira linha 1,25 cm | Nenhum |

Tamanho de página: A4 (21 x 29,7 cm), retrato. Usar paisagem apenas quando a
tabela exigir mais colunas do que cabe em retrato (ex.: tabela de
codificação) — nesse caso é uma seção `.docx` própria, não a página inteira.

**Atenção ao recuo de primeira linha (1,25 cm):** confirmado no XML do
`CORP-GQ-ANX-001.r04` que essa regra é aplicada por **formatação direta em
cada parágrafo de texto corrido** (`w:ind w:firstLine="709"` = 1,25 cm),
não é algo que vem "de graça" do estilo Normal (que no template está vazio
de propriedades). Ao gerar `.docx` programaticamente, é fácil esquecer
esse recuo porque nenhum estilo garante — aplicar
`paragraph_format.first_line_indent = Cm(1.25)` em cada parágrafo de texto
corrido explicitamente. **Não aplicar em:** títulos/subtítulos, itens de
lista com marcador (já têm recuo próprio do nível de lista) e texto dentro
de tabela (regra da tabela é "Nenhum").

## Cor de destaque das tabelas ("azul claro")

Toda tabela usada em qualquer documento deve ter a **primeira linha** (linha
de cabeçalho das colunas) preenchida com azul claro:

- Cor canônica: `#DBE5F1` (usar esta por padrão).
- Variante aceitável observada em um exemplo do próprio template: `#D9E2F3`
  — não é a regra, apenas não é motivo para rejeitar um documento antigo
  que já a usa; ao padronizar, migrar para `#DBE5F1`.
- Nunca usar preenchimento cinza (`#D9D9D9`, `#FAFAFA`), amarelo
  (`#FFFF00`) ou qualquer outra cor de destaque no cabeçalho de tabela —
  esses são sinais típicos de documento fora do padrão.

## Estilos de título (Heading)

O `.docx` corporativo usa os estilos internos `Heading 1/2/3` (nomes PT-BR
"Título 1/2/3") com Arial 12 pt. `Heading 1` e `Heading 3` são negrito por
definição de estilo; ao usá-los para títulos numerados (`1. OBJETIVOS`,
`6.1.1 TIPOS DE DOCUMENTOS`) sempre digitar o texto em MAIÚSCULO — o estilo
não força isso automaticamente.

## Estilo de escrita (o "e escrita" do pedido do usuário)

Além do layout, padronizar também a redação:

- Passos de procedimento em **infinitivo/imperativo** ("Verificar...",
  "Adicionar...", "Manter..."), nunca em primeira pessoa ou gerúndio solto.
  É o padrão observado em todos os documentos corporativos analisados.
- Nomes de item/etapa em MAIÚSCULO, numerados sequencialmente
  (`1. PREPARO DOS EQUIPAMENTOS`, `2. ADIÇÃO DE INGREDIENTES`...).
- Definições e documentos de referência em ordem alfabética / cronológica,
  conforme pedido explícito no escopo de cada seção (ver
  `escopos-por-tipo.md`).
- Sem abreviações inconsistentes ou apelidos internos sem explicação na
  primeira ocorrência (ex.: siglas de equipamento devem aparecer por
  extenso ao menos uma vez).
- Unidades e faixas numéricas com o mesmo formato em todo o documento
  (ex.: "25 a 30 minutos", não misturar com "25-30 min" no mesmo texto).
