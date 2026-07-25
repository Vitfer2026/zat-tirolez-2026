# Codificação dos documentos (fonte: CORP-GQ-ANX-001, item 2)

Padrão de código: `[LOCALIDADE]-[ÁREA]-[TIPO]-[SEQUENCIAL]`

Exemplos reais confirmados nos documentos analisados:
`CORP-GQ-ANX-001`, `TIR-1-FAB-POP-004`, `TIR-1-FAB-FOR-019`,
`TIR-1-FAB-TAB-002`, `TIR-1-FAB-IT-002/006/007/008`, `TIR-1-SAL-POP-002`.

(No exemplo `TIR-1-...`, o "1" depois de TIR é a planta/fábrica dentro da
localidade Tiros — manter esse segmento se já existir no código original;
não é parte da tabela de Localidade abaixo, é um detalhe de numeração
interna da unidade e deve ser preservado como está.)

## Localidade

| Sigla | Nome |
|---|---|
| CORP | Corporativo |
| ARA | Arapuá |
| CXB | Caxambu |
| LIN | Lins |
| MON | Monte |
| TIR | Tiros |
| CD | Centro de Armazenamento e Distribuição |
| MTZ | Matriz |

## Área

| Sigla | Nome |
|---|---|
| REG | Assuntos Regulatórios |
| CQ | Controle de Qualidade |
| GQ | Garantia da Qualidade |
| SUP | Suprimentos |
| MNT | Manutenção |
| UTL | Utilidades |
| REC | Recepção de Matéria-prima |
| BEN | Beneficiamento |
| PRE | Preparo de Leite |
| FAB | Fabricação |
| SAL | Salga |
| SEC | Secagem |
| MAT | Maturação |
| ACA | Acabamento |
| EXT | Área Externa |
| P&D | Pesquisa & Desenvolvimento |
| PLT | Política Leiteira |
| P&C | Pessoas & Cultura |

## Tipo de documento

| Sigla | Nome | Código sequencial (se listado) |
|---|---|---|
| POL | Política | 001 |
| MAN | Manual | 002 |
| PR | Procedimento | 003 |
| PRO | Programa | 004 |
| ANX | Anexo | 005 |
| POP | Procedimento Operacional Padrão | — |
| IT | Instrução de Trabalho | — |
| PAP | Procedimento Analítico Padrão | — |
| REL | Relatório | — |
| DES | Descritivo de Processo | — |
| FOR | Formulário | — |
| TAB | Tabela | — |

As Especificações Técnicas seguem codificação própria (documento
específico, não coberto por esta tabela).

## Ao padronizar um documento existente

- **Nunca invente ou troque o código de outro documento.** Se o documento
  já tem código válido (bate com a tabela acima), preservar exatamente,
  só ajustando revisão/datas conforme o histórico de alterações.
- Se o código estiver ausente, obviamente errado (ex.: sigla de área que
  não existe na tabela, ou copiado de outra unidade/planta), sinalizar a
  divergência ao usuário em vez de decidir sozinho qual código correto
  atribuir — pedir confirmação antes de gravar um código novo.
- Sequencial (o número final, ex. `004` em `POP-004`) nunca é reaproveitado
  de outro documento da mesma área/tipo — se precisar de um código novo,
  perguntar qual é o próximo sequencial livre em vez de supor.
