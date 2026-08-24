# Consolidações de Embarques e DUIMP

Automação operacional da C·TRÊS para tratamento de relações de embarques de seguros de transportes.

O projeto usa PowerShell e Excel Desktop e oferece dois processos:

- **Consolidação de Embarques:** gera um workbook por apólice e competência, com uma aba para cada relação do mesmo lote.
- **Consolidação de DUIMP:** preserva as DIs e consolida adições e itens pelo número-base da DUIMP.

## Requisitos

- Windows;
- Microsoft Excel Desktop;
- Windows PowerShell;
- arquivos de entrada no formato `.xlsx`.

Não é necessário instalar Python, Node.js ou alterar permanentemente a política de execução do PowerShell.

## Estrutura operacional

```text
.
├── PROCESSAR_CONSOLIDACOES.cmd
├── 01_CONSOLIDAÇÃO_DE_EMBARQUES
│   ├── INPUT
│   ├── OUTPUT
│   └── PROCESSADOS
├── 02_CONSOLIDAÇÃO_DE_DUIMP
│   ├── INPUT
│   ├── OUTPUT
│   └── PROCESSADOS
└── AUTOMACAO
    ├── Consolidar.ps1
    ├── Consolidacao.Core.psm1
    ├── GRUPOS_EMBARQUES.csv
    └── ASSETS
```

As pastas operacionais são criadas automaticamente pelo motor quando não existem.

## Como usar

1. Feche no Excel as relações que serão processadas.
2. Coloque cada `.xlsx` no `INPUT` do processo desejado.
3. Execute `PROCESSAR_CONSOLIDACOES.cmd`.
4. Confira o resumo apresentado no terminal.
5. Confira o resultado em `OUTPUT`.

Após o sucesso, o original é movido para `PROCESSADOS`. Resultados existentes nunca são sobrescritos: novas versões recebem `_V2`, `_V3` e assim por diante.

Consulte [COMO_USAR.txt](COMO_USAR.txt) para as regras operacionais completas.

## Segurança dos dados

Este repositório não contém relações de segurados, resultados, arquivos processados, logs, relatórios de erro ou histórico operacional. O `.gitignore` impede o versionamento acidental desses dados.

Antes de publicar qualquer alteração, confirme que nenhum documento de cliente foi incluído no commit.

## Configuração de nomes amigáveis

O arquivo `AUTOMACAO/GRUPOS_EMBARQUES.csv` associa uma apólice a um nome de lote:

```csv
Apolice;NomeLote
0000000000000;NOME_DO_LOTE
```

A apólice deve ser mantida como texto, inclusive com zeros à esquerda.
