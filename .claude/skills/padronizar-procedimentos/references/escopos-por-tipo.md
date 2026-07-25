# Escopos padrão por tipo de documento (fonte: CORP-GQ-ANX-001)

Cada tipo de documento tem uma estrutura de seções obrigatória. **Antes de
padronizar, identifique o tipo do documento recebido** (pelo código, pelo
título do cabeçalho, ou pelo conteúdo) e aplique o escopo correspondente.
Não misture escopos de tipos diferentes.

## 1. Manuais, PAC, Programas e Procedimentos (documentos de gestão, com capa)

Estrutura de seções numeradas (`Heading 1` para o nível principal):

1. **OBJETIVOS** — o que o documento define (ex.: "Descrever como
   codificar e elaborar procedimentos ou documentos na empresa").
2. **DOCUMENTOS DE REFERÊNCIA** — tabela em ordem cronológica:
   `Nº | Norma/Documento | Origem | Data da Publicação | Descrição`.
3. **CAMPOS DE APLICAÇÃO** — empresa/unidade a que o documento se aplica
   (ex.: "Este documento aplica-se a todas as unidades da empresa
   Laticínios Tirolez Ltda").
4. **DEFINIÇÕES** — termos e definições em **ordem alfabética**.
5. **RESPONSABILIDADES** — matriz de responsabilidades: linhas = requisito,
   colunas = cargos/áreas, marcação `X` em quem é responsável. Colunas
   típicas: Colaboradores, Diretoria, Gerência Industrial, Gerência da
   Qualidade, Produção, Qualidade Corporativa, Qualidade Fábrica (ajustar
   colunas ao público real do documento).
6. **DESCRIÇÃO** — corpo principal do processo, pode ter subtítulos
   `Heading 2` (6.1) e `Heading 3` (6.1.1) e assim por diante.
7. **MONITORAMENTO** (se aplicável) — tabela
   `O QUE? | COMO? | QUANDO? | QUEM? | DOCUMENTOS`.
8. **CORREÇÃO / AÇÃO CORRETIVA** (se aplicável) — tabela
   `NÃO CONFORMIDADE | CORREÇÃO | AÇÃO CORRETIVA | QUANDO? | QUEM? | DOCUMENTOS`.
9. **VERIFICAÇÃO** (se aplicável) — tabela
   `O QUE? | COMO? | QUANDO? | QUEM?`.
10. **REGISTROS** — tabela
    `DOCUMENTO | RETENÇÃO | DISPOSIÇÃO | RECUPERAÇÃO | ARMAZENAMENTO`.
11. **ANEXOS** (se aplicável) — lista de documentos (código + descrição) e
    fluxogramas vinculados, ex.: "CORP-GQ-ANX-028 – Modelo de Fluxograma".

Seções marcadas "(Se aplicável)" podem ser omitidas se genuinamente não se
aplicarem ao documento — mas omitir não é o padrão, é a exceção; se o
documento original tinha conteúdo equivalente a uma dessas seções, ele
**deve** virar a seção correspondente, não ficar solto em outro lugar.

## 2. Procedimento Operacional Padrão (POP) / Instrução de Trabalho (IT)

Estrutura mais enxuta, **sem capa**, sem headings numerados — direto do
cabeçalho para a tabela de informações do procedimento:

**Tabela de informações** (2 colunas, uma linha por campo; campos sem
conteúdo aplicável podem ser omitidos, mas os 3 primeiros são obrigatórios):

- Responsável pela verificação: cargo (ex.: "Supervisor / Coordenador do
  setor").
- EPI's e acessórios: todos os EPIs obrigatórios para a execução.
- Equipamentos / Utensílios: (se aplicável).
- Matéria-prima e Ingredientes: (se aplicável).
- Linha de Produtos: incluir imagens da linha, se aplicável.
- Local: onde o documento é utilizado (se aplicável).
- Equipamentos / Produtos / Utensílios utilizados: para IT de
  higienização — produtos químicos e concentração, utensílios de limpeza.
- Considerações: qualquer observação relevante (se aplicável).
- Documentos relacionados: outros POP/IT/Formulários referenciados
  (`Código – Nome do documento`), repetido ao lado dos campos acima.

**Tabela de procedimento** (3 colunas):
`ITEM | DESCRIÇÃO | RESPONSÁVEL`

- Coluna ITEM: `N. NOME DA ETAPA` em maiúsculo (ex.: "1. SEPARAÇÃO",
  "2. ADIÇÃO"). Repetir o mesmo número/nome em todas as linhas que
  pertencem à mesma etapa, se a etapa precisar de mais de uma linha.
- Coluna DESCRIÇÃO: texto em infinitivo/imperativo. Imagens e tabelas
  auxiliares citadas no texto ("conforme foto abaixo", "seguir tabela
  abaixo") ficam **centralizadas dentro da mesma célula**, logo após o
  texto que as referencia — nunca extraídas para fora da tabela.
- Coluna RESPONSÁVEL: cargo/função de quem executa.
- Linha `VERIFICAÇÃO (Se aplicável)` mesclada nas 3 colunas, seguida de
  linha(s) com a descrição da verificação e o responsável por ela.
- Linha `REGISTRO (Se aplicável)` mesclada, seguida do(s) formulário(s)
  de registro (`Código – Nome do documento`).

**Fluxograma (se aplicável)**: tabela de 1 coluna com título
"FLUXOGRAMA* (se aplicável)" e o fluxograma (feito no Bizagi) inserido
como imagem na linha abaixo.

## 3. Tabela de Higienização

Tabela única com colunas:
`ONDE | O QUE | PRODUTOS | MATERIAIS | FREQUÊNCIA | COMO | QUEM`

Linha final mesclada `VERIFICAÇÃO` com a descrição de quem verifica
(diferente de quem executou a higienização) e a referência ao plano de
amostragem/análises laboratoriais.

## 4. Procedimento Analítico Padrão (PAP)

**Tabela de informações** (2 colunas):
- Aplicação: o que o método mede/verifica.
- Método de Referência: nome do método (ex.: "Método Gerber -
  butirométrico").
- EPI's e Acessórios.
- Documentos relacionados (Programas e Formulários).
- Equipamentos / Utensílios.
- Reagentes.
- Princípio: explicação resumida do fundamento técnico do método (célula
  mesclada nas duas colunas).

**Tabela de procedimento** (2 colunas):
`ITEM | DESCRIÇÃO`
- `1. PROCEDIMENTO` em maiúsculo — passo a passo em infinitivo.
- `2. INTERPRETAÇÃO / LEITURA DO RESULTADO` — como ler/interpretar o
  resultado, com ilustração se necessário.

## 5. APPCC (Estudos de Análise de Perigos e Pontos Críticos de Controle)

Usar o modelo específico `CORP-GQ-FOR-043` — não recriar a estrutura do
zero; se o usuário enviar um estudo de APPCC fora do padrão, sinalizar que
o formulário correto a seguir é aquele código e pedir para localizá-lo
antes de prosseguir.

## Como decidir qual escopo usar

1. Olhar o tipo no código do documento (`POP`, `IT`, `PAP`, `TAB`, `MAN`,
   `PR`, `PRO`, `POL`) — ver `codificacao.md`.
2. Se o código estiver ausente/errado, inferir pelo conteúdo: uma
   sequência linear de passos de execução por um operador é POP/IT; uma
   tabela de limpeza por área é Tabela de Higienização; um método de
   análise laboratorial é PAP; um documento de política/diretriz mais
   amplo com responsabilidades e monitoramento é Manual/Procedimento/
   Programa.
3. Em caso de dúvida real entre dois escopos, perguntar ao usuário antes
   de escolher — aplicar o escopo errado obriga a reescrever o documento
   inteiro depois.
