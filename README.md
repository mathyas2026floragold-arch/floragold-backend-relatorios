# Backend FloraGold v9 — Relatório fiel ao sistema

Esta versão corrige a captura para priorizar o **exportador nativo do FloraGold**.

## O que mudou na v9

- O robô **nunca captura antes de aplicar o filtro de período**.
- O robô clica em **Filtros**, preenche **data inicial** e **data final**, clica em **Buscar/Pesquisar/Filtrar** e espera a tabela atualizar.
- Depois disso, em vez de ler só a tabela visível, ele usa o botão nativo **Excel/CSV** do sistema.
- Isso deixa o arquivo final mais fiel ao relatório oficial baixado manualmente no FloraGold.
- O contador de páginas é dinâmico: página `X de Y`, conforme o sistema mostrar.

## Variável importante

```env
EXPORT_MODE=native
```

Use `native` em produção. Use `table` só para teste.

## Render

Build command:

```bash
npm install && npx playwright install chromium
```

Start command:

```bash
npm start
```

## Importante

Para dados reais, use:

```env
DEMO_MODE=false
```

Com `DEMO_MODE=true`, os dados são simulados e não representam o relatório real.


## v10 — Modelo oficial FloraGold

Ajustado usando o arquivo oficial enviado:
- Aba: `Relatório`
- Colunas obrigatórias:
  `Protocolo`, `Status`, `Origem`, `Data`, `Hora`, `Fila`, `Destino`, `Operador`, `Conversa`, `Espera`, `Duracao`

A versão v10 usa o exportador nativo do FloraGold e valida o cabeçalho. Se o arquivo baixado não trouxer as colunas oficiais, o job falha para evitar relatório incompleto.

Para relatório verídico:
- `DEMO_MODE=false`
- `EXPORT_MODE=native`


## v12 - Relatório verídico obrigatório

Esta versão não gera relatório simulado. Para baixar relatório real, configure no Render:

```env
DEMO_MODE=false
EXPORT_MODE=native
```

Regras da v12:
- O robô deve aplicar o filtro de data antes de exportar.
- O robô prioriza o botão Excel/CSV nativo do FloraGold.
- Telefone, protocolo, origem, operador, data e hora nunca são inventados.
- Se `DEMO_MODE=true`, a geração de relatório é bloqueada com erro para evitar dados falsos.
- O arquivo final só é aceito se tiver as colunas oficiais do relatório Entrante.


## Ajuste v12
- Front-end sem campos de hora.
- Datas em texto `dd/mm/aaaa`, sem limite 2025/2027 e sem calendário nativo que falhava.
- Backend continua aplicando o dia inteiro internamente: 00:00 até 23:59.
