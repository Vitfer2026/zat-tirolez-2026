# Capa, cabeçalho e rodapé (fonte: CORP-GQ-ANX-001, itens 3 e 4)

## Capa

Obrigatória para **Manuais, Programas e Procedimentos** (documentos de
"gestão", tipicamente maiores). Composta por: cabeçalho padrão, título do
documento, imagem/slogan do departamento ou projeto (se aplicável) e
rodapé.

POP e IT normalmente **não têm capa separada** — a primeira página já
começa pelo cabeçalho seguido direto da tabela de informações do
procedimento (ver `escopos-por-tipo.md`). Isso foi confirmado no exemplo
real `TIR-1-FAB-POP-004`, que não tem capa.

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

**Erro comum a evitar:** construir isso como uma tabela de 6 linhas com
`merge` vertical nas colunas 1 e 2. Visualmente pode parecer parecido, mas
gera bordas horizontais indesejadas entre cada linha de metadado (porque o
estilo de tabela do padrão é "Tabela com Grade", que desenha toda borda de
linha) e não é fiel à estrutura real do documento oficial.

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
