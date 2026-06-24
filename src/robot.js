const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { chromium } = require('playwright');

const REPORT_DIR = path.resolve(process.env.REPORT_DIR || './relatorios');
const DEMO_MODE = String(process.env.DEMO_MODE || 'true').toLowerCase() === 'true';
const HEADLESS = String(process.env.HEADLESS || 'true').toLowerCase() !== 'false';
const LINHAS_POR_PAGINA = Number(process.env.LINHAS_POR_PAGINA || 3000);

function envList(name, fallback) {
  const raw = process.env[name];
  if (!raw) return fallback;
  return raw.split(',').map(v => v.trim()).filter(Boolean);
}

function envSelector(name, fallback = []) {
  const value = process.env[name];
  return value ? [value, ...fallback] : fallback;
}

const TEXTS = {
  // No FloraGold, a pausa segura aparece como "1002 - Ligação 0800".
  // Não use apenas o texto do select, porque "Ligação 0800" também aparece no campo de seleção.
  pause: envList('PAUSE_TEXTS', ['ligacao 0800', 'ligação 0800', 'retirar pausa']),
  activeCall: envList('ACTIVE_CALL_TEXTS', ['em chamada', 'chamada ativa', 'atendendo', 'discando', 'ligacao ativa', 'ligação ativa', 'atendimento em andamento']),
  available: envList('AVAILABLE_TEXTS', ['disponivel', 'disponível', 'online', 'livre'])
};

const SELECTORS = {
  loginUser: envSelector('LOGIN_USER_SELECTOR', [
    'input[placeholder="Usuário"]',
    'input[placeholder="Usuario"]',
    'input[name="usuario"]',
    'input[name="user"]',
    'input[name="login"]',
    'input[type="text"]'
  ]),
  loginPass: envSelector('LOGIN_PASS_SELECTOR', [
    'input[placeholder="Senha"]',
    'input[name="senha"]',
    'input[name="password"]',
    'input[type="password"]'
  ]),
  loginButton: envSelector('LOGIN_BUTTON_SELECTOR', [
    'button:has-text("Logar")',
    'button:has-text("Entrar")',
    'input[type="submit"]',
    '.btn:has-text("Logar")',
    '.btn:has-text("Entrar")'
  ]),
  pauseButton: envSelector('PAUSE_BUTTON_SELECTOR', [
    '[title*="Iniciar pausa"]',
    '[title*="iniciar pausa"]',
    'button:has-text("Iniciar pausa")',
    'a:has-text("Iniciar pausa")',
    'button:has-text("Pausa")',
    'a:has-text("Pausa")',
    'button:has-text("Pausar")',
    'a:has-text("Pausar")',
    '[title*="Pausa"]',
    '[title*="pausa"]',
    '.btn-warning',
    '.btn.btn-warning'
  ]),
  removePauseButton: envSelector('REMOVE_PAUSE_BUTTON_SELECTOR', [
    '[title*="Retirar pausa"]',
    '[title*="retirar pausa"]',
    'button:has-text("Retirar pausa")',
    'a:has-text("Retirar pausa")'
  ]),
  pauseReasonSelect: envSelector('PAUSE_REASON_SELECTOR', [
    'select:has(option:text("Ligação 0800"))',
    'select:has(option:text("Ligacao 0800"))',
    'select'
  ]),
  reportMenu: envSelector('REPORT_MENU_SELECTOR', [
    'text=Relatórios',
    'text=Relatorios',
    'a:has-text("Relatórios")',
    'a:has-text("Relatorios")',
    'button:has-text("Relatórios")',
    'button:has-text("Relatorios")'
  ]),
  callsMenu: envSelector('CALLS_MENU_SELECTOR', [
    'text=Chamadas',
    'a:has-text("Chamadas")',
    'button:has-text("Chamadas")'
  ]),
  entranteMenu: envSelector('ENTRANTE_MENU_SELECTOR', [
    'text=Entrante',
    'a:has-text("Entrante")',
    'button:has-text("Entrante")'
  ]),
  filtersButton: envSelector('FILTERS_BUTTON_SELECTOR', [
    'button:has-text("Filtros")',
    'a:has-text("Filtros")',
    '.btn:has-text("Filtros")',
    '[title*="Filtros"]'
  ]),
  dateRange: envSelector('FILTER_DATE_SELECTOR', [
    'input[name="data"]',
    'input[name*="data"]',
    'input[placeholder*="Data"]',
    'input[placeholder*="Período"]',
    'input[placeholder*="Periodo"]',
    '.modal input[type="text"]'
  ]),
  startDate: envSelector('FILTER_START_DATE_SELECTOR', []),
  endDate: envSelector('FILTER_END_DATE_SELECTOR', []),
  startTime: envSelector('FILTER_START_TIME_SELECTOR', [
    'input[name*="hora_inicio"]',
    'input[name*="inicio"]',
    'input[placeholder*="Hora inicial"]',
    'input[type="time"]'
  ]),
  endTime: envSelector('FILTER_END_TIME_SELECTOR', [
    'input[name*="hora_fim"]',
    'input[name*="fim"]',
    'input[placeholder*="Hora final"]',
    'input[type="time"]'
  ]),
  operator: envSelector('FILTER_OPERATOR_SELECTOR', [
    'select[name*="operador"]',
    'select[name*="usuario"]',
    'select:has(option:text("Todos os operadores"))'
  ]),
  queue: envSelector('FILTER_QUEUE_SELECTOR', [
    'select[name*="fila"]',
    'select[name*="queue"]',
    'select:has(option:text("Todas as filas"))'
  ]),
  status: envSelector('FILTER_STATUS_SELECTOR', [
    'select[name*="status"]'
  ]),
  searchButton: envSelector('FILTER_SEARCH_BUTTON_SELECTOR', [
    'button:has-text("Buscar")',
    'button:has-text("Pesquisar")',
    'button:has-text("Filtrar")',
    '.modal .btn-primary',
    '.btn:has-text("Buscar")'
  ]),
  rowsSelect: envSelector('ROWS_SELECT_SELECTOR', [
    'select[name*="length"]',
    'select'
  ]),
  nextButton: envSelector('NEXT_BUTTON_SELECTOR', [
    'a:has-text("Próximo")',
    'a:has-text("Proximo")',
    'button:has-text("Próximo")',
    'button:has-text("Proximo")',
    '.paginate_button.next:not(.disabled)',
    '.pagination .next:not(.disabled) a',
    '.pagination .next:not(.disabled) button'
  ])
};

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function slug(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .toLowerCase();
}

function parseDateBR(date) {
  const [d, m, y] = String(date || '').split('/');
  return `${y || '0000'}-${m || '00'}-${d || '00'}`;
}

function elapsed(start) {
  const seconds = Math.floor((Date.now() - start) / 1000);
  return formatDuration(seconds);
}

function formatDuration(secondsInput) {
  const seconds = Math.max(0, Math.round(Number(secondsInput || 0)));
  const min = Math.floor(seconds / 60);
  const sec = seconds % 60;
  return `${String(min).padStart(2, '0')}m ${String(sec).padStart(2, '0')}s`;
}

function estimateRemaining(started, currentPage, totalPages) {
  if (!currentPage || !totalPages) return '--';
  const elapsedSeconds = Math.max((Date.now() - started) / 1000, 1);
  const avgPerPage = elapsedSeconds / currentPage;
  const remaining = Math.max(0, (totalPages - currentPage) * avgPerPage);
  return formatDuration(remaining);
}

function averagePageSpeed(started, currentPage) {
  if (!currentPage) return '—';
  const elapsedMinutes = Math.max((Date.now() - started) / 60000, 0.1);
  return `${Math.max(1, Math.round(currentPage / elapsedMinutes))} páginas/min`;
}

function csvEscape(v) {
  return `"${String(v ?? '').replace(/"/g, '""')}"`;
}

function normalizeText(value) {
  return String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
}

function containsAny(text, terms) {
  const clean = normalizeText(text);
  return terms.some(term => clean.includes(normalizeText(term)));
}

function getRuntimeConfig() {
  return {
    demoMode: DEMO_MODE,
    sistemaUrl: process.env.SISTEMA_URL || '',
    headless: HEADLESS,
    linhasPorPagina: LINHAS_POR_PAGINA,
    autoPause: String(process.env.AUTO_PAUSE || 'true').toLowerCase() === 'true',
    waitForCallFinish: String(process.env.WAIT_FOR_CALL_FINISH || 'false').toLowerCase() === 'true'
  };
}

function saveFiles({ rows, headers, filenameBase }) {
  fs.mkdirSync(REPORT_DIR, { recursive: true });

  if (!headers?.length && rows[0]) headers = Object.keys(rows[0]);
  if (!headers?.length) headers = ['Resultado'];

  const csvFile = `${filenameBase}.csv`;
  const xlsxFile = `${filenameBase}.xlsx`;
  const csvPath = path.join(REPORT_DIR, csvFile);
  const xlsxPath = path.join(REPORT_DIR, xlsxFile);

  const csv = [
    headers.map(csvEscape).join(';'),
    ...rows.map(row => headers.map(h => csvEscape(row[h])).join(';'))
  ].join('\n');

  fs.writeFileSync(csvPath, '\uFEFF' + csv, 'utf8');

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows, { header: headers });
  XLSX.utils.book_append_sheet(wb, ws, 'Relatório');
  XLSX.writeFile(wb, xlsxPath);

  return { csvFile, xlsxFile, csvPath, xlsxPath };
}

function buildDemoRows(params, total = 61358) {
  const operadores = ['Ismaias Mathyas', 'Jailson Maceno', 'Midian Silva', 'Vinicius Araújo', 'Kayki Lima', 'Amanda Franciele', 'Rennan Victor', 'Genilson Luz'];
  const statusList = ['Atendida', 'Atendida', 'Atendida', 'Abandono'];
  const filas = ['rcpt0800', '0800', 'Atendimento', 'Equipe 1'];
  const ddds = ['82', '81', '91', '83', '68', '97', '65', '28'];
  const rows = [];

  for (let i = 1; i <= total; i++) {
    const operador = params.operador && !String(params.operador).includes('Todos') ? params.operador : operadores[i % operadores.length];
    const fila = params.fila && !String(params.fila).includes('Todas') ? params.fila : filas[i % filas.length];
    const status = params.statusFiltro && !String(params.statusFiltro).includes('Todos') ? params.statusFiltro : statusList[i % statusList.length];
    const ddd = ddds[i % ddds.length];

    rows.push({
      Protocolo: `ENT-${String(i).padStart(7, '0')}`,
      Status: status,
      Origem: `(${ddd}) 9${String(80000000 + i).slice(-8)}`,
      Data: params.dataInicial,
      Hora: `${String(8 + (i % 12)).padStart(2, '0')}:${String(i % 60).padStart(2, '0')}:${String((i * 7) % 60).padStart(2, '0')}`,
      Fila: fila,
      Destino: String(1000 + (i % 30)),
      Operador: operador,
      Conversa: `00:0${i % 6}:${String((i * 3) % 60).padStart(2, '0')}`,
      Espera: `00:00:${String(i % 20).padStart(2, '0')}`,
      Duracao: `00:0${i % 7}:${String((i * 5) % 60).padStart(2, '0')}`
    });
  }
  return rows;
}

async function launchBrowser() {
  const browser = await chromium.launch({
    headless: HEADLESS,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage']
  });
  const context = await browser.newContext({
    viewport: { width: 1440, height: 900 },
    acceptDownloads: true,
    locale: 'pt-BR'
  });
  const page = await context.newPage();
  page.setDefaultTimeout(20000);
  return { browser, context, page };
}

async function validateAccessJob(params) {
  if (DEMO_MODE) {
    await sleep(700);
    return {
      login: true,
      pausa: true,
      status0800: 0,
      ligacaoAtiva: false,
      message: 'Validação simulada concluída. DEMO_MODE=true.'
    };
  }

  const { browser, page } = await launchBrowser();
  try {
    await login(page, params);
    const safety = await ensureSafety(page, null, { validateOnly: true });
    return { ...safety, message: 'Acesso validado no sistema real.' };
  } finally {
    await browser.close();
  }
}

async function readFilterOptionsJob(params) {
  if (DEMO_MODE || !params?.usuario || !params?.senha) {
    return demoOptions();
  }

  const { browser, page } = await launchBrowser();
  try {
    await login(page, params);
    await ensureSafety(page);
    await openReport(page, params.tipoRelatorio || 'Entrante');
    await openFilters(page);
    const options = await extractSelectOptions(page);
    return {
      operadores: options.operadores.length ? options.operadores : demoOptions().operadores,
      filas: options.filas.length ? options.filas : demoOptions().filas,
      status: options.status.length ? options.status : demoOptions().status
    };
  } finally {
    await browser.close();
  }
}

function demoOptions() {
  return {
    operadores: ['Todos os operadores', 'Ismaias Mathyas', 'Jailson Maceno', 'Midian Silva', 'Vinicius Araújo', 'Kayki Lima', 'Amanda Franciele', 'Rennan Victor'],
    filas: ['Todas as filas', 'rcpt0800', '0800', 'Atendimento', 'Cobrança', 'Equipe 1', 'Equipe 2'],
    status: ['Todos os status', 'Atendida', 'Abandono', 'Ocupado', 'Não atendida']
  };
}

async function runReportJob(params, update = () => {}) {
  if (DEMO_MODE) return runDemoJob(params, update);
  return runRealJob(params, update);
}

async function runDemoJob(params, update) {
  const started = Date.now();
  const steps = [
    ['Login concluído', 12],
    ['Pausa 0800 validada', 24],
    ['Relatório aberto', 35],
    ['Filtros aplicados', 45]
  ];

  for (const [step, progress] of steps) {
    await sleep(550);
    update({ step, progress, log: step });
  }

  const totalRegistros = 61358;
  const paginas = Math.ceil(totalRegistros / LINHAS_POR_PAGINA);

  for (let p = 1; p <= paginas; p++) {
    await sleep(120);
    const registrosLidos = Math.min(totalRegistros, p * LINHAS_POR_PAGINA);
    update({
      step: `Coletando página ${p} de ${paginas}`,
      progress: Math.min(92, 45 + Math.round((p / paginas) * 47)),
      paginaAtual: p,
      totalPaginas: paginas,
      registrosLidos,
      linhasPorPagina: LINHAS_POR_PAGINA,
      velocidadeMedia: averagePageSpeed(started, p),
      tempoEstimadoRestante: estimateRemaining(started, p, paginas),
      log: `Página ${p} de ${paginas} capturada.`
    });
  }

  update({ step: 'Consolidando arquivo', progress: 96, tempoEstimadoRestante: '00m 00s', log: 'Gerando CSV e Excel.' });
  await sleep(500);

  const rows = buildDemoRows(params, totalRegistros);
  const headers = Object.keys(rows[0]);
  const base = `relatorio_${slug(params.tipoRelatorio || 'entrante')}_${parseDateBR(params.dataInicial)}_${parseDateBR(params.dataFinal)}_${Date.now()}`;
  const files = saveFiles({ rows, headers, filenameBase: base });

  return {
    totalRegistros,
    paginasCapturadas: paginas,
    tempoTotal: elapsed(started),
    ...files,
    demoMode: true
  };
}

async function runRealJob(params, update) {
  const started = Date.now();
  const { browser, page } = await launchBrowser();

  try {
    update({ step: 'Acessando sistema', progress: 5, log: 'Abrindo sistema de origem.' });
    await login(page, params);
    update({ step: 'Login concluído', progress: 15, log: 'Login realizado.' });

    await ensureSafety(page, update);
    update({ step: 'Pausa 0800 validada', progress: 25, log: 'Usuário em pausa Ligação 0800, sem ligação ativa.' });

    await openReport(page, params.tipoRelatorio || 'Entrante');
    update({ step: 'Relatório aberto', progress: 35, log: `Relatório ${params.tipoRelatorio || 'Entrante'} aberto.` });

    await applyFilters(page, params);
    update({ step: 'Filtros aplicados', progress: 45, log: 'Filtros aplicados.' });

    await setRowsPerPage(page, LINHAS_POR_PAGINA);
    const { rows, headers, paginasCapturadas } = await captureAllPages(page, update, started);

    if (!rows.length) {
      throw new Error('Nenhuma linha capturada. Verifique período, filtro, permissões ou seletores do sistema.');
    }

    const base = `relatorio_${slug(params.tipoRelatorio || 'entrante')}_${parseDateBR(params.dataInicial)}_${parseDateBR(params.dataFinal)}_${Date.now()}`;
    const files = saveFiles({ rows, headers, filenameBase: base });

    return {
      totalRegistros: rows.length,
      paginasCapturadas,
      tempoTotal: elapsed(started),
      ...files,
      demoMode: false
    };
  } catch (err) {
    await saveErrorScreenshot(page, params.jobId || `erro_${Date.now()}`);
    throw err;
  } finally {
    await browser.close();
  }
}

async function login(page, { usuario, senha }) {
  const url = process.env.SISTEMA_URL;
  if (!url) throw new Error('SISTEMA_URL não configurada.');

  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 70000 });
  await sleep(800);

  await fillFirst(page, SELECTORS.loginUser, usuario, 'campo usuário');
  await fillFirst(page, SELECTORS.loginPass, senha, 'campo senha');
  await handleOptionalLoginSelects(page);
  await clickFirst(page, SELECTORS.loginButton, 'botão logar');

  await page.waitForLoadState('networkidle', { timeout: 70000 }).catch(() => {});
  await sleep(2500);

  const body = await page.locator('body').innerText({ timeout: 10000 }).catch(() => '');
  if (/senha inválida|usuario inválido|usuário inválido|login inválido|erro de autenticação/i.test(body)) {
    throw new Error('Login recusado pelo sistema. Confira usuário e senha.');
  }
}

async function handleOptionalLoginSelects(page) {
  const values = [process.env.LOGIN_SELECT_1_VALUE, process.env.LOGIN_SELECT_2_VALUE].filter(Boolean);
  if (!values.length) return;

  const selects = page.locator('select');
  const count = await selects.count().catch(() => 0);
  for (let i = 0; i < Math.min(values.length, count); i++) {
    const value = values[i];
    const loc = selects.nth(i);
    await loc.selectOption(value).catch(async () => {
      await loc.selectOption({ label: value }).catch(() => {});
    });
  }
}

async function readBody(page) {
  return page.locator('body').innerText({ timeout: 15000 }).catch(() => '');
}

function extractOperatorStatus(body) {
  const normalizedLines = String(body || '')
    .split(/\n+/)
    .map(line => line.replace(/\s+/g, ' ').trim())
    .filter(Boolean);

  // Exemplo real FloraGold: "1002 - Disponível" ou "1002 - Ligação 0800".
  const statusLine = normalizedLines.find(line => /\b\d{3,5}\s*-\s*/.test(line));
  if (statusLine) return statusLine;

  const compact = String(body || '').replace(/\s+/g, ' ').trim();
  const match = compact.match(/\b\d{3,5}\s*-\s*(Dispon[ií]vel|Liga[cç][aã]o 0800|Campanha|Entrante|Manual|Atendimento|Em chamada)[^\s]*/i);
  return match ? match[0] : '';
}

async function hasAnyVisible(page, selectors) {
  for (const selector of selectors.filter(Boolean)) {
    const loc = page.locator(selector).first();
    const count = await loc.count().catch(() => 0);
    if (!count) continue;
    const visible = await loc.isVisible().catch(() => false);
    if (visible) return true;
  }
  return false;
}

async function readSafety(page) {
  const body = await readBody(page);
  const operatorStatus = extractOperatorStatus(body);
  const removePauseVisible = await hasAnyVisible(page, SELECTORS.removePauseButton);
  const normalizedStatus = normalizeText(operatorStatus);

  const paused0800 =
    normalizedStatus.includes('ligacao 0800') ||
    removePauseVisible ||
    (/\b\d{3,5}\s*-\s*liga[cç][aã]o 0800/i.test(operatorStatus));

  const available =
    normalizedStatus.includes('disponivel') ||
    /\b\d{3,5}\s*-\s*Dispon[ií]vel/i.test(operatorStatus);

  const activeCall = containsAny(body, TEXTS.activeCall);
  const unsafeStatus = containsAny(operatorStatus, ['campanha', 'entrante', 'manual', 'atendimento', 'em chamada']);
  const ligacaoAtiva = Boolean(activeCall || unsafeStatus);

  return {
    login: true,
    pausa: paused0800,
    pausa0800: paused0800,
    disponivel: available && !paused0800,
    ligacaoAtiva,
    status0800: paused0800 && !ligacaoAtiva ? 0 : 1,
    statusOperador: operatorStatus || 'não detectado',
    textoDetectado: body.slice(0, 900)
  };
}

async function selectPause0800(page) {
  const reason = process.env.PAUSE_REASON_LABEL || 'Ligação 0800';

  for (const selector of SELECTORS.pauseReasonSelect.filter(Boolean)) {
    const loc = page.locator(selector).first();
    const count = await loc.count().catch(() => 0);
    if (!count) continue;

    try {
      await loc.selectOption({ label: reason });
      await sleep(400);
      return true;
    } catch {}

    try {
      await loc.selectOption(reason);
      await sleep(400);
      return true;
    } catch {}
  }

  // Fallback para selects customizados: tenta clicar no texto se ele aparecer.
  await page.getByText(reason, { exact: false }).first().click({ timeout: 3000 }).catch(() => null);
  await sleep(400);
  return true;
}

async function activatePause0800(page, update = null) {
  update?.({ step: 'Ativando pausa 0800', progress: 20, log: 'Selecionando Ligação 0800 e clicando em Iniciar pausa.' });

  await selectPause0800(page);
  await clickFirst(page, SELECTORS.pauseButton, 'botão Iniciar pausa');

  const started = Date.now();
  while (Date.now() - started < 30000) {
    await sleep(1200);
    const safety = await readSafety(page);
    if (safety.pausa0800 && !safety.ligacaoAtiva) return safety;
  }

  const safety = await readSafety(page);
  throw new Error(`Não foi possível confirmar a pausa 0800. Status detectado: ${safety.statusOperador}`);
}

async function ensureSafety(page, update = null, options = {}) {
  const waitForCallFinish = String(process.env.WAIT_FOR_CALL_FINISH || 'false').toLowerCase() === 'true';
  const autoPause = String(process.env.AUTO_PAUSE || 'true').toLowerCase() === 'true';

  let safety = await readSafety(page);

  if (safety.ligacaoAtiva && waitForCallFinish && !options.validateOnly) {
    update?.({ step: 'Aguardando ligação finalizar', progress: 18, log: 'Ligação ativa detectada. Aguardando finalizar.' });
    const started = Date.now();
    while (Date.now() - started < 10 * 60 * 1000) {
      await sleep(5000);
      safety = await readSafety(page);
      if (!safety.ligacaoAtiva) break;
    }
  }

  if (safety.ligacaoAtiva) {
    throw new Error(`Execução bloqueada: existe ligação ativa ou status inseguro. Status detectado: ${safety.statusOperador}`);
  }

  if (!safety.pausa0800 && safety.disponivel && autoPause && !options.validateOnly) {
    safety = await activatePause0800(page, update);
  }

  if (!safety.pausa0800) {
    const msg = options.validateOnly
      ? 'Validação bloqueada: o usuário ainda não está em pausa 0800.'
      : 'Execução bloqueada: o usuário precisa estar em pausa 0800 antes de gerar relatório.';
    throw new Error(`${msg} Status detectado: ${safety.statusOperador}`);
  }

  return safety;
}

async function openReport(page, tipoRelatorio) {
  await clickFirst(page, SELECTORS.reportMenu, 'menu Relatórios');
  await sleep(700);

  await clickFirst(page, SELECTORS.callsMenu, 'submenu Chamadas').catch(() => null);
  await sleep(500);

  const tipo = String(tipoRelatorio || 'Entrante').trim();
  if (normalizeText(tipo).includes('entrant')) {
    await clickFirst(page, SELECTORS.entranteMenu, 'relatório Entrante');
  } else {
    await clickFirst(page, [`text=${tipo}`, `a:has-text("${tipo}")`, `button:has-text("${tipo}")`], `relatório ${tipo}`);
  }

  await page.waitForLoadState('networkidle', { timeout: 70000 }).catch(() => {});
  await sleep(1500);
}

async function openFilters(page) {
  await clickFirst(page, SELECTORS.filtersButton, 'botão Filtros');
  await sleep(1000);
}

async function applyFilters(page, params) {
  await openFilters(page);

  const periodo = `${params.dataInicial} ${params.horaInicial} - ${params.dataFinal} ${params.horaFinal}`;

  if (SELECTORS.startDate.length && SELECTORS.endDate.length) {
    await fillFirst(page, SELECTORS.startDate, params.dataInicial, 'data inicial').catch(() => null);
    await fillFirst(page, SELECTORS.endDate, params.dataFinal, 'data final').catch(() => null);
  }

  await fillDateRange(page, SELECTORS.dateRange, periodo, 'campo de período/data').catch(async () => {
    // Se não encontrar campo único de período, tenta preencher os 2 primeiros inputs de data visíveis.
    const dateInputs = page.locator('.modal input, input').filter({ hasNotText: '' });
    const count = await dateInputs.count().catch(() => 0);
    if (count >= 2) {
      await dateInputs.nth(0).fill(params.dataInicial).catch(() => {});
      await dateInputs.nth(1).fill(params.dataFinal).catch(() => {});
    } else {
      throw new Error('Não encontrei o campo de período/data. Ajuste FILTER_DATE_SELECTOR.');
    }
  });

  await fillTimeFields(page, params);
  await selectIfPresent(page, SELECTORS.operator, params.operador, ['Todos os operadores', 'Todos']);
  await selectIfPresent(page, SELECTORS.queue, params.fila, ['Todas as filas', 'Todas']);
  await selectIfPresent(page, SELECTORS.status, params.statusFiltro || params.status, ['Todos os status', 'Todos']);

  await clickFirst(page, SELECTORS.searchButton, 'botão Buscar/Pesquisar');
  await page.waitForLoadState('networkidle', { timeout: 70000 }).catch(() => {});
  await sleep(2500);
}

async function fillDateRange(page, selectors, value, label) {
  const loc = await firstExisting(page, selectors);
  if (!loc) throw new Error(`Não encontrei ${label}.`);

  await loc.click({ timeout: 10000 }).catch(() => {});
  await loc.fill('').catch(() => {});
  await loc.fill(String(value)).catch(async () => {
    await loc.evaluate((el, val) => {
      el.value = val;
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
    }, String(value));
  });
  await loc.press('Enter').catch(() => {});
  await sleep(300);
}

async function fillTimeFields(page, params) {
  const startSet = await fillFirst(page, SELECTORS.startTime, params.horaInicial, 'hora inicial', { optional: true });
  const endSet = await fillFirst(page, SELECTORS.endTime, params.horaFinal, 'hora final', { optional: true, skipFirstMatchIfSame: startSet?.selector });

  if (!startSet || !endSet) {
    const inputs = page.locator('.modal input[type="time"], input[type="time"], input[placeholder="HH:MM"], input[name*="hora"]');
    const count = await inputs.count().catch(() => 0);
    if (count >= 2) {
      await inputs.nth(0).fill(params.horaInicial).catch(() => {});
      await inputs.nth(1).fill(params.horaFinal).catch(() => {});
    }
  }
}

async function selectIfPresent(page, selectors, value, ignored = []) {
  if (!value || ignored.some(v => normalizeText(value).includes(normalizeText(v)))) return false;
  const loc = await firstExisting(page, selectors);
  if (!loc) return false;

  await loc.selectOption({ label: value }).catch(async () => {
    await loc.selectOption(value).catch(() => null);
  });
  await sleep(300);
  return true;
}

async function setRowsPerPage(page, rowsPerPage) {
  const frames = page.frames();
  for (const frame of frames) {
    const selects = frame.locator(SELECTORS.rowsSelect.join(','));
    const count = await selects.count().catch(() => 0);
    for (let i = 0; i < count; i++) {
      const s = selects.nth(i);
      const text = await s.innerText().catch(() => '');
      if (text.includes(String(rowsPerPage))) {
        await s.selectOption(String(rowsPerPage)).catch(async () => {
          await s.selectOption({ label: String(rowsPerPage) }).catch(() => {});
        });
        await sleep(1800);
        return true;
      }
    }
  }
  return false;
}

async function captureAllPages(page, update, started = Date.now()) {
  const rows = [];
  let headers = [];
  let paginasCapturadas = 0;
  const seen = new Set();
  let lastSignature = '';

  for (let pageNumber = 1; pageNumber <= 250; pageNumber++) {
    const source = await bestTableSource(page);
    if (!source) break;

    const extracted = await source.frame.evaluate(() => {
      const text = el => (el?.innerText || el?.textContent || '').replace(/\s+/g, ' ').trim();
      const visible = el => !!(el && el.getClientRects().length);
      const tables = Array.from(document.querySelectorAll('table'))
        .filter(visible)
        .map(t => ({ table: t, cells: t.querySelectorAll('tbody td').length, len: text(t).length }))
        .sort((a, b) => (b.cells + b.len) - (a.cells + a.len));
      const table = tables[0]?.table;
      if (!table) return { headers: [], rows: [], signature: '' };

      let headers = Array.from(table.querySelectorAll('thead tr:last-child th')).map(text).filter(Boolean);
      const rows = Array.from(table.querySelectorAll('tbody tr')).map(tr => Array.from(tr.querySelectorAll('td,th')).map(text))
        .filter(r => r.length && r.some(c => c) && !/nenhum registro|no matching records|carregando|processando/i.test(r.join(' ')));
      if (!headers.length && rows[0]) headers = rows[0].map((_, i) => `Coluna ${i + 1}`);
      const signature = rows.slice(0, 3).map(r => r.join('|')).join('§') + `#${rows.length}`;
      return { headers, rows, signature };
    });

    if (!extracted.rows.length && pageNumber === 1) break;
    if (!headers.length && extracted.headers.length) headers = extracted.headers;

    let added = 0;
    for (const arr of extracted.rows) {
      const key = arr.join('||');
      if (seen.has(key)) continue;
      seen.add(key);
      const obj = {};
      headers.forEach((h, i) => { obj[h || `Coluna ${i + 1}`] = arr[i] || ''; });
      rows.push(obj);
      added++;
    }

    paginasCapturadas++;
    update({
      step: `Coletando página ${pageNumber}`,
      progress: Math.min(92, 45 + Math.round(pageNumber * 2)),
      paginaAtual: pageNumber,
      totalPaginas: 0,
      registrosLidos: rows.length,
      linhasPorPagina: LINHAS_POR_PAGINA,
      velocidadeMedia: averagePageSpeed(started, pageNumber),
      tempoEstimadoRestante: '--',
      log: `Página ${pageNumber}: ${added} linhas capturadas.`
    });

    const moved = await clickNext(source.frame);
    if (!moved) break;

    await waitTableChange(source.frame, extracted.signature || lastSignature);
    lastSignature = extracted.signature;
  }

  return { rows, headers, paginasCapturadas };
}

async function bestTableSource(page) {
  let best = null;
  for (const frame of page.frames()) {
    const score = await frame.evaluate(() => {
      const visible = el => !!(el && el.getClientRects().length);
      return Array.from(document.querySelectorAll('table')).filter(visible).reduce((acc, t) => {
        const cells = t.querySelectorAll('tbody td').length;
        const len = (t.innerText || '').length;
        return Math.max(acc, cells + len);
      }, 0);
    }).catch(() => 0);
    if (score > 0 && (!best || score > best.score)) best = { frame, score };
  }
  return best;
}

async function clickNext(frame) {
  for (const selector of SELECTORS.nextButton) {
    const loc = frame.locator(selector).first();
    const count = await loc.count().catch(() => 0);
    if (!count) continue;
    const disabled = await loc.evaluate(el => el.disabled || el.classList.contains('disabled') || el.parentElement?.classList.contains('disabled') || el.getAttribute('aria-disabled') === 'true').catch(() => true);
    if (disabled) continue;
    await loc.click({ timeout: 10000 }).catch(() => null);
    await sleep(1000);
    return true;
  }
  return false;
}

async function waitTableChange(frame, previousSignature) {
  const started = Date.now();
  while (Date.now() - started < 30000) {
    await sleep(700);
    const signature = await frame.evaluate(() => {
      const text = el => (el?.innerText || el?.textContent || '').replace(/\s+/g, ' ').trim();
      const table = Array.from(document.querySelectorAll('table')).sort((a, b) => b.querySelectorAll('tbody td').length - a.querySelectorAll('tbody td').length)[0];
      if (!table) return '';
      const rows = Array.from(table.querySelectorAll('tbody tr')).map(tr => Array.from(tr.querySelectorAll('td,th')).map(text));
      return rows.slice(0, 3).map(r => r.join('|')).join('§') + `#${rows.length}`;
    }).catch(() => '');
    if (signature && signature !== previousSignature) return true;
  }
  return false;
}

async function extractSelectOptions(page) {
  return page.evaluate(() => {
    const text = el => (el?.innerText || el?.textContent || '').replace(/\s+/g, ' ').trim();
    const result = { operadores: [], filas: [], status: [] };

    Array.from(document.querySelectorAll('select')).forEach(sel => {
      const label = (sel.name || sel.id || sel.closest('label')?.innerText || '').toLowerCase();
      const options = Array.from(sel.options).map(o => o.textContent.trim()).filter(Boolean);
      if (!options.length) return;
      if (/operador|usuario|usuário/.test(label)) result.operadores.push(...options);
      else if (/fila|queue/.test(label)) result.filas.push(...options);
      else if (/status|situacao|situação/.test(label)) result.status.push(...options);
    });

    for (const key of Object.keys(result)) result[key] = Array.from(new Set(result[key]));
    return result;
  }).catch(() => ({ operadores: [], filas: [], status: [] }));
}

async function firstExisting(pageOrFrame, selectors) {
  for (const selector of selectors.filter(Boolean)) {
    const loc = pageOrFrame.locator(selector).first();
    const count = await loc.count().catch(() => 0);
    if (count) return loc;
  }
  return null;
}

async function fillFirst(page, selectors, value, label, options = {}) {
  for (const selector of selectors.filter(Boolean)) {
    if (options.skipFirstMatchIfSame && selector === options.skipFirstMatchIfSame) continue;
    const loc = page.locator(selector).first();
    if (await loc.count().catch(() => 0)) {
      try {
        await loc.fill(String(value));
        return { selector };
      } catch {}
    }
  }
  if (options.optional) return null;
  throw new Error(`Não encontrei ${label}. Ajuste o seletor no Render.`);
}

async function clickFirst(page, selectors, label) {
  for (const selector of selectors.filter(Boolean)) {
    const loc = page.locator(selector).first();
    if (await loc.count().catch(() => 0)) {
      try {
        await loc.click({ timeout: 12000 });
        return true;
      } catch {}
    }
  }
  throw new Error(`Não encontrei ${label}. Ajuste o seletor no Render.`);
}

async function saveErrorScreenshot(page, jobId) {
  try {
    const dir = path.join(REPORT_DIR, 'erros');
    fs.mkdirSync(dir, { recursive: true });
    await page.screenshot({ path: path.join(dir, `${slug(jobId)}.png`), fullPage: true });
  } catch {}
}

module.exports = {
  runReportJob,
  validateAccessJob,
  readFilterOptionsJob,
  getRuntimeConfig
};
