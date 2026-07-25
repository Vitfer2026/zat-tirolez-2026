# Capa, cabeçalho e rodapé (fonte: CORP-GQ-ANX-001, itens 3 e 4)

## Capa

Obrigatória para **Manuais, PAC, Programas e Procedimentos** (documentos
de "gestão", tipicamente maiores). Confirmada em 3 exemplos reais
(`CORP-GQ-PAC-003`, `TIR-GQ-MAN-002`, `CORP-GQ-PRO-003`) — a receita é bem
mais enxuta do que se poderia supor:

1. Vários parágrafos em branco (espaçamento vertical, sem conteúdo).
2. **Título do documento, sozinho, 28pt, negrito, centralizado,
   maiúsculo** (ex.: "CONTROLE INTEGRADO DE PRAGAS",
   "SISTEMA DE GESTÃO DA QUALIDADE E SEGURANÇA DE ALIMENTOS",
   "GESTÃO DE OCORRÊNCIAS"). É só isso — **não** repetir o tipo de
   documento ("PROGRAMA", "MANUAL") nem o nome da unidade na capa: esses
   dados já estão no cabeçalho, que se repete em toda página inclusive a
   primeira.
3. Uma imagem centralizada (mascote/ícone do departamento — nos 3
   exemplos reais é o mesmo mascote em formato de queijo com uma prancheta,
   variando só o texto/selo impresso na prancheta conforme o assunto do
   documento; é a "imagem/slogan do departamento ou projeto" citada no
   template). Se não houver arte específica do departamento disponível,
   usar a logo da empresa como alternativa e sinalizar que é um
   placeholder — não inventar um mascote/ícone novo.
4. Mais parágrafos em branco.
5. Quebra de página / início direto da primeira seção do escopo (`1.
   OBJETIVO(S)` etc.) — **sem** repetir código/revisão na capa.

POP e IT normalmente **não têm capa separada** — a primeira página já
começa pelo cabeçalho seguido direto da tabela de informações do
procedimento (ver `escopos-por-tipo.md`). Isso foi confirmado nos
exemplos reais `TIR-1-FAB-POP-004` e `MON-EBP-POP-002`, que não têm capa.

## Cabeçalho (repete em todas as páginas)

**Estrutura real confirmada no XML do `CORP-GQ-ANX-001.r04`** (não é uma
tabela de 6 linhas com células mescladas — é uma **tabela de 1 linha e 3
colunas**, e cada coluna guarda várias informações como **parágrafos
empilhados dentro da mesma célula**, sem nenhuma borda horizontal entre
eles):

| Coluna 1 — logo (~20% da largura) | Coluna 2 — tipo + título (~57%) | Coluna 3 — metadados (~23%) |
|---|---|---|
| Logo da empresa, ~2,5–3 cm de largura, centralizado, alinhamento vertical centralizado na célula | 5 parágrafos empilhados, todos centralizados: (vazio) / `TIPO DE DOCUMENTO` (Arial 9pt negrito) / (vazio) / **título do documento em maiúsculo** (Arial 14pt negrito) / (vazio) | 6 parágrafos empilhados, todos **centralizados** (`jc=center`, confirmado no XML — não é alinhado à esquerda apesar de parecer em capturas de tela pequenas), Arial 8pt: rótulo em negrito + valor sem negrito na mesma linha |

Conteúdo exato dos 6 parágrafos da coluna 3, na ordem:
1. `Código: ` (negrito) + valor (não negrito)
2. `Revisão: ` (negrito) + valor (não negrito)
3. `Data Revisão: ` (negrito) + valor (não negrito)
4. `Data Aprovação: ` (negrito) + valor (não negrito)
5. `Página: ` (negrito) + campo dinâmico `PAGE` + ` de ` + campo dinâmico `NUMPAGES` (usar campos de verdade, não número fixo — é assim que o software da empresa preenche e é o único jeito de ficar correto em qualquer paginação)

**Bug real já cometido e corrigido ao gerar campo `PAGE`/`NUMPAGES` via
`python-docx`:** cada `<w:fldChar>` (begin/separate/end) e o
`<w:instrText>` precisam estar em **runs (`<w:r>`) separados**, um
elemento por run — nunca todos dentro do mesmo `<w:r>`. Empacotar tudo
num único run passa despercebido pela validação de schema (XSD não
reclama), mas quebra o parser de campo complexo do Word na prática: o
sintoma foi o **cabeçalho inteiro sumir ao abrir o arquivo**, não só o
número da página. Conferido contra `CORP-GQ-PAC-003` (que também usa
`PAGE`/`NUMPAGES`): lá cada `fldChar`/`instrText` tem seu próprio `<w:r>`,
com a mesma `rPr` repetida em cada um. Incluir também um valor
"cacheado" (texto literal) entre `separate` e `end` (ex.: `1`), como
fallback visual antes do Word recalcular o campo.
6. Nome da unidade/localidade, sozinho, **em negrito** (ex.: "TIROLEZ" no anexo corporativo, "TIROS - MG" num documento de planta)

Exemplo real confirmado (`TIR-1-FAB-POP-004`, mesma estrutura):
```
PROCEDIMENTO OPERACIONAL PADRÃO — FABRICAÇÃO DO QUEIJO ESTEPE
Código: TIR-1-FAB-POP-004
Revisão: 5
Data Revisão: 28/04/2025
Data Aprovação: 02/07/2025
Página: 1 de 2
TIROS - MG
```

**Atualização após checar 4 documentos reais adicionais
(`CORP-GQ-PAC-003`, `TIR-GQ-MAN-002`, `MON-EBP-POP-002`,
`CORP-GQ-PRO-003`):** as duas construções abaixo são **igualmente
válidas** e visualmente indistinguíveis — a empresa usa as duas
(3 dos 4 documentos novos usam a primeira; o anexo template e
`CORP-GQ-PRO-003` usam a segunda):

- **Tabela de 6 linhas x 3 colunas**, com `merge` vertical nas colunas 1
  (logo) e 2 (tipo+título), e a coluna 3 com uma linha de metadado por
  linha da tabela. **O detalhe que faz isso funcionar**: a borda
  horizontal entre as linhas da coluna 3 é pintada de **branco**
  (`w:color="FFFFFF"` no `tcBorders`), não removida — por isso não
  aparece nenhuma linha visível entre "Código:", "Revisão:" etc., mesmo a
  tabela usando o estilo "Tabela com Grade" (que desenha toda borda por
  padrão). Esquecer de branquear essa borda é o erro mais fácil de
  cometer nessa construção.
- **Tabela de 1 linha x 3 colunas**, com a coluna 2 e a coluna 3 tendo
  vários parágrafos empilhados dentro da mesma célula (usada no anexo
  `CORP-GQ-ANX-001` e em `CORP-GQ-PRO-003`). Não tem bordas internas para
  esconder porque não há células extras.

Qualquer uma das duas é aceitável para gerar `.docx` programaticamente —
a de 1 linha é mais simples de implementar (não precisa branquear
bordas) e foi a adotada nesta skill.

## Rodapé (só na primeira página, não replicar nas demais)

Tabela de 3 colunas, **2 linhas**, estilo "Rodapé" do Word, Arial 8pt,
tudo centralizado:

| Elaboração: Nome | Verificação: Nome | Aprovação: Nome |
|---|---|---|
| Área/cargo de quem elaborou | Área/cargo de quem verificou | Área/cargo de quem aprovou |

Exemplo real confirmado no template (`CORP-GQ-ANX-001`):
```
Elaboração: Nome1        | Verificação: Nome2       | Aprovação: Nome3
Garantia da Qualidade    | Garantia da Qualidade    | Gerência da Qualidade
```

Preenchido automaticamente pelo software a partir do fluxo cadastrado para
cada pasta. Se mais de um responsável precisar dar parecer na mesma etapa,
preencher manualmente (o software não suporta múltiplos aprovadores por
etapa automaticamente). Ao reconstruir manualmente sem a 2ª linha
disponível (área/cargo desconhecidos), preencher com o que for possível
inferir das Responsabilidades do próprio documento e sinalizar a
suposição — não deixar a linha simplesmente ausente, pois isso muda a
estrutura da tabela (2 linhas é o padrão).

**Nota Excel:** planilhas (`.xlsx`) têm restrições de cabeçalho/rodapé
automático no software de gestão — nesses casos, preencher previamente
seguindo o exemplo `CORP-GQ-ANX-003`. Isso não se aplica a documentos
`.docx`.

## Ao padronizar um documento existente

- Se o documento já tem cabeçalho/rodapé com os campos certos mas em
  formatação diferente (fonte errada, cores erradas, ordem trocada),
  reconstruir a tabela no formato acima preservando os valores
  (código, revisão, datas, título) — não resetar revisão nem datas sem
  necessidade.
- Se faltar algum campo (ex.: "Página: N de M" ausente), adicionar mas
  sinalizar ao usuário que o valor precisa ser conferido, especialmente
  paginação, já que ela depende da diagramação final do corpo do
  documento.
