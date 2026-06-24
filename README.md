# Backend Render — Central Automática de Relatórios

Backend Node.js + Express + Playwright para rodar no Render e operar o sistema original sem API oficial.

## O que já vem pronto

- API para o front-end do Netlify.
- Modo demonstração (`DEMO_MODE=true`) para testar sem entrar no sistema real.
- Modo real (`DEMO_MODE=false`) com Playwright/Chromium.
- Validação obrigatória de segurança:
  - usuário deve estar em pausa;
  - status 0800 deve estar sem ligação ativa;
  - cancela se houver chamada ativa.
- Geração de CSV e Excel.
- Histórico resumido em JSON.
- Download de arquivos.
- CORS configurável para Netlify.
- Prints automáticos em caso de erro no modo real.
- Seletores configuráveis por variáveis de ambiente, sem precisar alterar código.

## Rotas da API

### `GET /health`
Verifica se o backend está online.

### `GET /api/status`
Retorna status do backend e configurações principais.

### `POST /api/validar-acesso`
Valida login e segurança.

Body:

```json
{
  "usuario": "seu_usuario",
  "senha": "sua_senha"
}
```

### `POST /api/relatorios/gerar`
Inicia a automação.

Body:

```json
{
  "usuario": "seu_usuario",
  "senha": "sua_senha",
  "tipoRelatorio": "Entrante",
  "dataInicial": "23/06/2026",
  "horaInicial": "00:00",
  "dataFinal": "23/06/2026",
  "horaFinal": "23:59",
  "operador": "Todos os operadores",
  "fila": "Todas as filas",
  "statusFiltro": "Todos os status"
}
```

Resposta:

```json
{
  "ok": true,
  "jobId": "..."
}
```

### `GET /api/jobs/:id`
Consulta progresso da execução.

### `GET /api/historico`
Lista últimas execuções.

### `GET /api/download/:filename`
Baixa CSV ou Excel gerado.

### `POST /api/opcoes`
Tenta carregar operadores, filas e status do sistema real. Em modo demo retorna opções simuladas.

## Como subir no Render

### Build Command

```bash
npm install && npx playwright install --with-deps chromium
```

### Start Command

```bash
npm start
```

## Variáveis de ambiente principais

```env
DEMO_MODE=true
HEADLESS=true
REPORT_DIR=./relatorios
LINHAS_POR_PAGINA=3000
SISTEMA_URL=https://floragold.acessocloud.com/index.php
CORS_ORIGIN=*
AUTO_PAUSE=false
WAIT_FOR_CALL_FINISH=false
```

Para produção, troque `CORS_ORIGIN=*` pelo domínio do Netlify:

```env
CORS_ORIGIN=https://seu-site.netlify.app
```

## Como testar primeiro

1. Suba no Render com:

```env
DEMO_MODE=true
```

2. Abra o front-end no Netlify.
3. Configure `API_BASE_URL` com a URL do Render.
4. Clique em Validar acesso e Baixar relatório.
5. Ele vai gerar CSV/Excel simulado.

## Como conectar no sistema real

Depois do teste:

```env
DEMO_MODE=false
```

A automação real já está preparada, mas talvez você precise ajustar os seletores do sistema original.

Principais seletores configuráveis:

```env
LOGIN_USER_SELECTOR=
LOGIN_PASS_SELECTOR=
LOGIN_BUTTON_SELECTOR=
PAUSE_BUTTON_SELECTOR=
REPORT_MENU_SELECTOR=
CALLS_MENU_SELECTOR=
ENTRANTE_MENU_SELECTOR=
FILTERS_BUTTON_SELECTOR=
FILTER_DATE_SELECTOR=
FILTER_START_DATE_SELECTOR=
FILTER_END_DATE_SELECTOR=
FILTER_START_TIME_SELECTOR=
FILTER_END_TIME_SELECTOR=
FILTER_OPERATOR_SELECTOR=
FILTER_QUEUE_SELECTOR=
FILTER_STATUS_SELECTOR=
FILTER_SEARCH_BUTTON_SELECTOR=
ROWS_SELECT_SELECTOR=
NEXT_BUTTON_SELECTOR=
```

Se deixar vazio, o backend usa seletores automáticos genéricos.

## Sobre a pausa

Por segurança, o backend só continua se detectar pausa. Os textos padrão são:

```env
PAUSE_TEXTS=pausa,pausado,administrativo,almoco,almoço
ACTIVE_CALL_TEXTS=em chamada,chamada ativa,atendendo,discando,ligacao ativa,ligação ativa
AVAILABLE_TEXTS=disponivel,disponível,online,livre
```

Se o sistema usa outro texto para pausa, adicione aqui.

## Importante

Este backend fica 100% pronto para estrutura, Render, API, geração de arquivos e modo real configurável. Para garantir funcionamento no sistema FloraGold real, ainda pode ser necessário confirmar os seletores HTML exatos da tela de login, pausa, filtros e paginação. Se o sistema tiver captcha, 2FA ou bloquear IP de servidor, será necessário ajustar a forma de autenticação.

## Progresso detalhado do relatório

Esta versão envia para o front-end campos extras durante a execução:

- `paginaAtual`
- `totalPaginas`
- `registrosLidos`
- `linhasPorPagina`
- `velocidadeMedia`
- `tempoEstimadoRestante`

Esses dados alimentam o contador visual de relatório com estimativa de tempo.


## Regra de pausa FloraGold v0.8.5

Este backend foi atualizado para o fluxo real informado:

- `1002 - Disponível` = ainda não está seguro para gerar relatório;
- selecionar `Ligação 0800` no campo de pausa;
- clicar no botão amarelo `Iniciar pausa`;
- aguardar o status mudar para `1002 - Ligação 0800`;
- quando aparece `Retirar pausa`, o robô entende que a pausa 0800 está ativa.

Variáveis principais no Render:

```env
AUTO_PAUSE=true
PAUSE_REASON_LABEL=Ligação 0800
PAUSE_TEXTS=ligacao 0800,ligação 0800,retirar pausa
AVAILABLE_TEXTS=disponivel,disponível,online,livre
```

Se algum botão não for encontrado, use o DevTools para descobrir o seletor e configure:

```env
PAUSE_BUTTON_SELECTOR=
PAUSE_REASON_SELECTOR=
REMOVE_PAUSE_BUTTON_SELECTOR=
```
