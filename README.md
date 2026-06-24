# Backend FloraGold v8 — filtro obrigatório por período

Esta versão corrige o comportamento do relatório para **não baixar tudo do sistema**.

## Regras v8

- O frontend envia apenas `dataInicial` e `dataFinal`.
- O backend aplica o filtro no FloraGold antes de capturar a tabela.
- O período enviado ao campo do sistema é: `DD/MM/AAAA - DD/MM/AAAA`.
- A captura só começa depois de clicar em Buscar/Pesquisar/Filtrar.
- O contador de páginas é dinâmico: ele tenta ler o total real da paginação/tabela.
- Não existe mais contador fixo `1 de 21`.
- Em `DEMO_MODE=true`, o total também é proporcional ao período escolhido, não 61.358 fixo.

## Render

Build Command:

```bash
npm install && npx playwright install chromium
```

Start Command:

```bash
npm start
```

Variáveis principais:

```env
DEMO_MODE=true
HEADLESS=true
AUTO_PAUSE=true
SISTEMA_URL=https://floragold.acessocloud.com/index.php
REPORT_DIR=./relatorios
LINHAS_POR_PAGINA=3000
```

Para sistema real, altere:

```env
DEMO_MODE=false
```
