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

Tabela de 3 colunas presente no cabeçalho de seção do Word (preenchimento
automático pelo software de gestão de documentos da empresa — ao produzir
o `.docx` manualmente, reproduzir os mesmos campos):

| Coluna 1 (logo) | Coluna 2 (tipo + título) | Coluna 3 (metadados, uma info por linha) |
|---|---|---|
| (logo/marca, quando aplicável) | `TIPO DE DOCUMENTO` (ex.: "INSTRUÇÃO DE TRABALHO", "PROCEDIMENTO OPERACIONAL PADRÃO") seguido do **título do documento em maiúsculo** | `Código: XXX-YYY-ZZZ-000` |
| | | `Revisão: N` |
| | | `Data Revisão: DD/MM/AAAA` |
| | | `Data Aprovação: DD/MM/AAAA` |
| | | `Página: N de M` |
| | | Nome da unidade/localidade (ex.: "TIROLEZ" para o anexo corporativo, "TIROS - MG" para um documento de planta) |

Exemplo real confirmado (`TIR-1-FAB-POP-004`):
```
PROCEDIMENTO OPERACIONAL PADRÃO — FABRICAÇÃO DO QUEIJO ESTEPE
Código: TIR-1-FAB-POP-004
Revisão: 5
Data Revisão: 28/04/2025
Data Aprovação: 02/07/2025
Página: 1 de 2
TIROS - MG
```

## Rodapé (só na primeira página, não replicar nas demais)

Tabela de 3 colunas, uma linha:

| Elaboração: Nome | Verificação: Nome | Aprovação: Nome |
|---|---|---|
| Cargo/área de quem elaborou | Cargo/área de quem verificou | Cargo/área de quem aprovou |

Preenchido automaticamente pelo software a partir do fluxo cadastrado para
cada pasta. Se mais de um responsável precisar dar parecer na mesma etapa,
preencher manualmente (o software não suporta múltiplos aprovadores por
etapa automaticamente).

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
