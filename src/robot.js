const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { chromium } = require('playwright');

const REPORT_DIR = path.resolve(process.env.REPORT_DIR || './relatorios');
const DEMO_MODE = String(process.env.DEMO_MODE || 'false').toLowerCase() === 'true';
const HEADLESS = String(process.env.HEADLESS || 'true').toLowerCase() !== 'false';
const LINHAS_POR_PAGINA = Number(process.env.LINHAS_POR_PAGINA || 3000);

// Modelo oficial confirmado pelo relatório exportado pelo FloraGold.
// O backend preserva essas colunas e falha se o exportador não trouxer esse conjunto básico.
const OFFICIAL_HEADERS = ['Protocolo', 'Status', 'Origem', 'Data', 'Hora', 'Fila', 'Destino', 'Operador', 'Conversa', 'Espera', 'Duracao'];

function normalizeHeaderKey(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, '')
    .toLowerCase();
}

function findHeader(headers, expected) {
  const expectedKey = normalizeHeaderKey(expected);
  return (headers || []).find(h => normalizeHeaderKey(h) === expectedKey);
}

function alignRowsToOfficialSchema(rows, headers) {
  const headerMap = {};
  for (const expected of OFFICIAL_HEADERS) {
    const found = findHeader(headers, expected);
    if (!found) return { ok: false, missing: expected, rows, headers };
    headerMap[expected] = found;
  }

  const aligned = (rows || []).map(row => {
    const out = {};
    for (const expected of OFFICIAL_HEADERS) out[expected] = row[headerMap[expected]] ?? '';
    for (const h of headers || []) {
      if (!Object.values(headerMap).includes(h) && !(h in out)) out[h] = row[h] ?? '';
    }
    return out;
  });

  const extraHeaders = (headers || []).filter(h => !Object.values(headerMap).includes(h));
  return { ok: true, rows: aligned, headers: [...OFFICIAL_HEADERS, ...extraHeaders] };
}

function validateOfficialReportSchema(rows, headers, context = 'relatório') {
  const aligned = alignRowsToOfficialSchema(rows, headers);
  if (!aligned.ok) {
    throw new Error(`O ${context} exportado não bate com o modelo oficial do FloraGold. Coluna ausente: ${aligned.missing}. O robô não vai gerar relatório incompleto.`);
  }
  return aligned;
}


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
  confirmPauseButton: envSelector('CONFIRM_PAUSE_SELECTOR', [
    'button:has-text("Confirmar")',
    'button:has-text("Sim")',
    'button:has-text("OK")',
    '.modal button.btn-primary',
    '.swal2-confirm'
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
  ]),
  exportExcelButton: envSelector('EXPORT_EXCEL_SELECTOR', [
    '.buttons-excel',
    'button:has-text("Excel")',
    'a:has-text("Excel")',
    '[title*="Excel"]',
    'button:has-text("XLS")',
    'a:has-text("XLS")'
  ]),
  exportCsvButton: envSelector('EXPORT_CSV_SELECTOR', [
    '.buttons-csv',
    'button:has-text("CSV")',
    'a:has-text("CSV")',
    '[title*="CSV"]'
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

function dateBRToDate(date) {
  const [d, m, y] = String(date || '').split('/').map(Number);
  if (!d || !m || !y) return null;
  return new Date(y, m - 1, d);
}

function daysBetweenBR(start, end) {
  const a = dateBRToDate(start);
  const b = dateBRToDate(end);
  if (!a || !b) return 1;
  return Math.max(1, Math.floor((b - a) / 86400000) + 1);
}

function addDays(date, days) {
  const d = new Date(date.getTime());
  d.setDate(d.getDate() + days);
  return d;
}

function formatDateBR(date) {
  return `${String(date.getDate()).padStart(2, '0')}/${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}`;
}

function demoTotalForRange(params) {
  const days = daysBetweenBR(params.dataInicial, params.dataFinal);
  const basePerDay = Number(process.env.DEMO_REGISTROS_POR_DIA || 135);
  return Math.max(1, Math.min(61358, days * basePerDay));
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

  // Se for relatório com modelo oficial, grava na ordem exata do FloraGold.
  const official = alignRowsToOfficialSchema(rows, headers);
  if (official.ok) {
    rows = official.rows;
    headers = official.headers;
  }


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

function buildDemoRows(params, total = demoTotalForRange(params)) {
  const operadores = ['Ismaias Mathyas', 'Jailson Maceno', 'Midian Silva', 'Vinicius Araújo', 'Kayki Lima', 'Amanda Franciele', 'Rennan Victor', 'Genilson Luz'];
  const statusList = ['Atendida', 'Atendida', 'Atendida', 'Abandono'];
  const filas = ['rcpt0800', '0800', 'Atendimento', 'Equipe 1'];
  const ddds = ['82', '81', '91', '83', '68', '97', '65', '28'];
  const rows = [];

  const startDate = dateBRToDate(params.dataInicial) || new Date();
  const days = daysBetweenBR(params.dataInicial, params.dataFinal);

  for (let i = 1; i <= total; i++) {
    const operador = params.operador && !String(params.operador).includes('Todos') ? params.operador : operadores[i % operadores.length];
    const fila = params.fila && !String(params.fila).includes('Todas') ? params.fila : filas[i % filas.length];
    const status = params.statusFiltro && !String(params.statusFiltro).includes('Todos') ? params.statusFiltro : statusList[i % statusList.length];
    const ddd = ddds[i % ddds.length];

    rows.push({
      Protocolo: `ENT-${String(i).padStart(7, '0')}`,
      Status: status,
      Origem: `(${ddd}) 9${String(80000000 + i).slice(-8)}`,
      Data: formatDateBR(addDays(startDate, (i - 1) % days)),
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
    const safety = await ensureSafety(page, null, { validateOnly: false });
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
  // v11: relatório verídico obrigatório.
  // DEMO_MODE=true serve apenas para testar conexão/validação visual, não para gerar arquivo,
  // porque dados simulados podem confundir telefone/protocolo/origem.
  if (DEMO_MODE) {
    throw new Error('DEMO_MODE=true está ativo. Para gerar relatório verídico, altere DEMO_MODE=false no Render. Nenhum dado simulado será gerado.');
  }
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

  const totalRegistros = demoTotalForRange(params);
  const paginas = Math.max(1, Math.ceil(totalRegistros / LINHAS_POR_PAGINA));

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

    // v9: por padrão usamos o exportador NATIVO do FloraGold após aplicar os filtros.
    // Isso evita relatório "meia boca": o arquivo final passa a usar as mesmas colunas
    // que o sistema exporta no botão Excel/CSV, em vez de depender só das colunas visíveis da tabela.
    const exportMode = String(process.env.EXPORT_MODE || 'native').toLowerCase();
    const captured = exportMode === 'table'
      ? await captureAllPages(page, update, started)
      : await exportNativeFilteredReport(page, update, started);

    let { rows, headers, paginasCapturadas, totalExpected, sourceMode } = captured;

    if (!rows.length) {
      throw new Error('Nenhuma linha capturada. Verifique se o filtro de período foi aplicado e se o botão Excel/CSV do sistema exportou dados.');
    }

    // v10: proteção contra relatório incompleto.
    // O arquivo final só é gerado se o exportador nativo trouxer as colunas do relatório oficial FloraGold.
    if (normalizeText(params.tipoRelatorio || 'Entrante').includes('entrant')) {
      const official = validateOfficialReportSchema(rows, headers, 'relatório Entrante');
      rows = official.rows;
      headers = official.headers;
    }

    const base = `relatorio_${slug(params.tipoRelatorio || 'entrante')}_${parseDateBR(params.dataInicial)}_${parseDateBR(params.dataFinal)}_${Date.now()}`;
    const files = saveFiles({ rows, headers, filenameBase: base });

    return {
      totalRegistros: totalExpected || rows.length,
      registrosCapturados: rows.length,
      paginasCapturadas,
      tempoTotal: elapsed(started),
      modoCaptura: sourceMode || exportMode,
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

  // Primeiro tenta por select nativo. Pelo print, "Ligação 0800" fica em um select antes do botão amarelo.
  for (const selector of SELECTORS.pauseReasonSelect.filter(Boolean)) {
    const loc = page.locator(selector).first();
    const count = await loc.count().catch(() => 0);
    if (!count) continue;

    try {
      await loc.selectOption({ label: reason });
      await sleep(500);
      return { ok: true, mode: 'select-label', selector };
    } catch {}

    try {
      await loc.selectOption({ label: 'Ligacao 0800' });
      await sleep(500);
      return { ok: true, mode: 'select-label-sem-acento', selector };
    } catch {}

    try {
      await loc.evaluate((el, wanted) => {
        const clean = v => String(v || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
        const opt = Array.from(el.options || []).find(o => clean(o.textContent).includes(clean(wanted)));
        if (opt) {
          el.value = opt.value;
          el.dispatchEvent(new Event('input', { bubbles: true }));
          el.dispatchEvent(new Event('change', { bubbles: true }));
        }
      }, reason);
      await sleep(500);
      return { ok: true, mode: 'select-evaluate', selector };
    } catch {}
  }

  // Fallback para dropdown customizado.
  await page.getByText(reason, { exact: false }).first().click({ timeout: 3000 }).catch(() => null);
  await sleep(500);
  return { ok: true, mode: 'text-click' };
}

async function clickPauseButtonNearReason(page) {
  // Preferência: clicar no botão amarelo do mesmo bloco/linha do select onde existe "Ligação 0800".
  const clicked = await page.evaluate(() => {
    const norm = v => String(v || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
    const visible = el => !!(el && el.getClientRects().length);
    const selects = Array.from(document.querySelectorAll('select')).filter(visible);

    for (const sel of selects) {
      const has0800 = Array.from(sel.options || []).some(o => norm(o.textContent).includes('ligacao 0800'));
      if (!has0800) continue;

      let root = sel.parentElement;
      for (let depth = 0; root && depth < 5; depth++, root = root.parentElement) {
        const buttons = Array.from(root.querySelectorAll('button,a,input[type="button"],input[type="submit"]')).filter(visible);
        const target = buttons.find(btn => {
          const txt = norm(btn.innerText || btn.textContent || btn.title || btn.getAttribute('aria-label') || btn.value || '');
          const cls = norm(btn.className || '');
          return txt.includes('iniciar pausa') || txt.includes('pausa') || cls.includes('btn-warning') || cls.includes('warning');
        }) || buttons.find(btn => norm(btn.className || '').includes('warning'));
        if (target) {
          target.click();
          return true;
        }
      }
    }
    return false;
  }).catch(() => false);

  if (clicked) {
    await sleep(800);
    return true;
  }

  return clickFirst(page, SELECTORS.pauseButton, 'botão Iniciar pausa');
}

async function confirmPauseIfPrompt(page) {
  // Alguns sistemas abrem modal de confirmação depois de clicar em iniciar pausa.
  await sleep(700);
  for (const selector of SELECTORS.confirmPauseButton.filter(Boolean)) {
    const loc = page.locator(selector).first();
    const count = await loc.count().catch(() => 0);
    if (!count) continue;
    const visible = await loc.isVisible().catch(() => false);
    if (!visible) continue;
    await loc.click({ timeout: 5000 }).catch(() => null);
    await sleep(1000);
    return true;
  }
  return false;
}

async function activatePause0800(page, update = null) {
  update?.({ step: 'Ativando pausa 0800', progress: 20, log: 'Selecionando Ligação 0800 e clicando em Iniciar pausa.' });

  const before = await readSafety(page);
  if (before.pausa0800 && !before.ligacaoAtiva) return before;

  await selectPause0800(page);
  await clickPauseButtonNearReason(page);
  await confirmPauseIfPrompt(page);

  const started = Date.now();
  while (Date.now() - started < 45000) {
    await sleep(1500);
    const safety = await readSafety(page);
    update?.({
      step: 'Aguardando confirmação da pausa 0800',
      progress: 22,
      statusOperador: safety.statusOperador,
      log: `Status detectado: ${safety.statusOperador}`
    });
    if (safety.pausa0800 && !safety.ligacaoAtiva) return safety;
  }

  const safety = await readSafety(page);
  throw new Error(`Não foi possível confirmar a pausa 0800. Status detectado: ${safety.statusOperador}. Verifique se o botão Iniciar pausa abriu confirmação ou se o seletor PAUSE_BUTTON_SELECTOR precisa ser ajustado.`);
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
  const beforeSignature = await readTableSignature(page).catch(() => '');
  await openFilters(page);

  // Regra v8: o relatório SEMPRE precisa ser filtrado pelo período escolhido.
  // A interface só envia data inicial e data final; não há filtro de horário no HTML.
  const periodo = `${params.dataInicial} - ${params.dataFinal}`;

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
  await waitForProcessingEnd(page);
  await waitTableChangeAnyFrame(page, beforeSignature);
  await sleep(1200);

  const afterSignature = await readTableSignature(page).catch(() => '');
  if (!afterSignature) {
    throw new Error('Filtro aplicado, mas não encontrei a tabela de resultado. Confira o seletor do relatório.');
  }
}

async function waitForProcessingEnd(page) {
  const started = Date.now();
  while (Date.now() - started < 30000) {
    const loading = await page.evaluate(() => {
      const visible = el => !!(el && el.getClientRects().length);
      const txt = el => (el?.innerText || el?.textContent || '').toLowerCase();
      return Array.from(document.querySelectorAll('.dataTables_processing,.loading,.loader,.overlay,.blockUI'))
        .some(el => visible(el) && /processando|carregando|aguarde|loading/.test(txt(el)));
    }).catch(() => false);
    if (!loading) return true;
    await sleep(500);
  }
  return false;
}

async function readTableSignature(page) {
  for (const frame of page.frames()) {
    const sig = await frame.evaluate(() => {
      const text = el => (el?.innerText || el?.textContent || '').replace(/\s+/g, ' ').trim();
      const table = Array.from(document.querySelectorAll('table'))
        .filter(t => t.getClientRects().length)
        .sort((a, b) => b.querySelectorAll('tbody td').length - a.querySelectorAll('tbody td').length)[0];
      if (!table) return '';
      const rows = Array.from(table.querySelectorAll('tbody tr')).map(tr => Array.from(tr.querySelectorAll('td,th')).map(text));
      return rows.slice(0, 3).map(r => r.join('|')).join('§') + `#${rows.length}`;
    }).catch(() => '');
    if (sig) return sig;
  }
  return '';
}

async function waitTableChangeAnyFrame(page, previousSignature) {
  if (!previousSignature) {
    await sleep(1000);
    return true;
  }
  const started = Date.now();
  while (Date.now() - started < 30000) {
    await sleep(700);
    const sig = await readTableSignature(page).catch(() => '');
    if (sig && sig !== previousSignature) return true;
  }
  // Não falha aqui porque alguns filtros podem manter a primeira linha igual; a captura ainda continua.
  return false;
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
  // Não preenchemos filtro de hora separado. O período já é enviado como dia completo (00:00 até 23:59).
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


function normalizeHeaderName(value, index) {
  const clean = String(value || '').replace(/\s+/g, ' ').trim();
  return clean || `Coluna ${index + 1}`;
}

function normalizeRowsFromSheet(matrix) {
  const cleanRows = (matrix || [])
    .map(r => (r || []).map(c => String(c ?? '').replace(/\s+/g, ' ').trim()))
    .filter(r => r.length && r.some(c => c));

  if (!cleanRows.length) return { headers: [], rows: [] };

  let headerIndex = 0;
  // Procura uma linha provável de cabeçalho. Relatórios exportados às vezes têm título nas primeiras linhas.
  for (let i = 0; i < Math.min(8, cleanRows.length); i++) {
    const joined = cleanRows[i].join(' ').toLowerCase();
    const filled = cleanRows[i].filter(Boolean).length;
    if (filled >= 2 && /(data|hora|operador|origem|destino|status|protocolo|fila|duração|duracao|espera)/i.test(joined)) {
      headerIndex = i;
      break;
    }
  }

  const headers = cleanRows[headerIndex].map(normalizeHeaderName);
  const rows = cleanRows.slice(headerIndex + 1)
    .filter(r => r.some(c => c) && !/total geral|nenhum registro|no matching records/i.test(r.join(' ')))
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = r[i] || ''; });
      return obj;
    });

  return { headers, rows };
}

function parseDownloadedFile(filePath) {
  const ext = path.extname(filePath).toLowerCase();

  if (['.xlsx', '.xls', '.xlsm', '.ods'].includes(ext)) {
    const wb = XLSX.readFile(filePath, { cellDates: false, raw: false });
    const sheetName = wb.SheetNames[0];
    const matrix = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: '' });
    return normalizeRowsFromSheet(matrix);
  }

  const raw = fs.readFileSync(filePath, 'utf8').replace(/^\uFEFF/, '');
  const lines = raw.split(/\r?\n/).filter(Boolean);
  const sep = raw.includes(';') ? ';' : ',';
  const matrix = lines.map(line => {
    // Parser simples suficiente para CSV do sistema; se houver aspas, respeita o básico.
    const out = [];
    let cur = '';
    let quoted = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; continue; }
      if (ch === '"') { quoted = !quoted; continue; }
      if (ch === sep && !quoted) { out.push(cur); cur = ''; continue; }
      cur += ch;
    }
    out.push(cur);
    return out;
  });
  return normalizeRowsFromSheet(matrix);
}

async function clickExportButtonAndWaitDownload(pageOrFrame, label = 'Excel/CSV') {
  const candidates = [...SELECTORS.exportExcelButton, ...SELECTORS.exportCsvButton];
  for (const selector of candidates.filter(Boolean)) {
    const loc = pageOrFrame.locator(selector).first();
    const count = await loc.count().catch(() => 0);
    if (!count) continue;
    const visible = await loc.isVisible().catch(() => true);
    if (!visible) continue;

    const downloadPromise = pageOrFrame.page ? pageOrFrame.page().waitForEvent('download', { timeout: 45000 }) : null;
    // Quando pageOrFrame é Page, não existe .page().
    const realPromise = downloadPromise || pageOrFrame.waitForEvent('download', { timeout: 45000 });
    await loc.click({ timeout: 12000 });
    const download = await realPromise;
    return download;
  }
  throw new Error(`Não encontrei o botão de exportação ${label}. Ajuste EXPORT_EXCEL_SELECTOR ou EXPORT_CSV_SELECTOR no Render.`);
}

async function getPageTotals(page) {
  for (const frame of page.frames()) {
    const totals = await frame.evaluate((linhasPorPagina) => {
      const text = el => (el?.innerText || el?.textContent || '').replace(/\s+/g, ' ').trim();
      const body = text(document.body);
      const parseNumber = v => Number(String(v || '').replace(/\./g, '').replace(',', '.').replace(/[^0-9.]/g, '')) || 0;
      let totalRecords = 0;
      const patterns = [
        /de\s+([\d\.]+)\s+registros/i,
        /total\s+de\s+([\d\.]+)\s+registros/i,
        /([\d\.]+)\s+registros\s+no\s+total/i,
        /filtered\s+from\s+([\d\.]+)\s+total/i
      ];
      for (const p of patterns) {
        const m = body.match(p);
        if (m) { totalRecords = parseNumber(m[1]); break; }
      }

      let totalPages = totalRecords ? Math.max(1, Math.ceil(totalRecords / linhasPorPagina)) : 0;
      const pageNumbers = Array.from(document.querySelectorAll('.paginate_button, .pagination a, .pagination button, a, button'))
        .map(el => text(el))
        .filter(v => /^\d+$/.test(v))
        .map(Number);
      if (pageNumbers.length) totalPages = Math.max(totalPages, Math.max(...pageNumbers));
      return { totalRecords, totalPages };
    }, LINHAS_POR_PAGINA).catch(() => ({ totalRecords: 0, totalPages: 0 }));
    if (totals.totalRecords || totals.totalPages) return totals;
  }
  return { totalRecords: 0, totalPages: 0 };
}

async function exportNativeFilteredReport(page, update, started = Date.now()) {
  fs.mkdirSync(REPORT_DIR, { recursive: true });

  const allRows = [];
  let headers = [];
  const seen = new Set();
  let paginasCapturadas = 0;
  let { totalRecords, totalPages } = await getPageTotals(page);

  update({
    step: totalPages ? `Exportando arquivo do sistema 1 de ${totalPages}` : 'Exportando arquivo do sistema',
    progress: 48,
    paginaAtual: 1,
    totalPaginas: totalPages || 0,
    registrosLidos: 0,
    log: 'Usando botão nativo Excel/CSV do FloraGold para manter o relatório completo.'
  });

  for (let p = 1; p <= 250; p++) {
    const source = await bestTableSource(page);
    const frame = source?.frame || page.mainFrame();

    const download = await clickExportButtonAndWaitDownload(frame, 'Excel/CSV');
    const suggested = download.suggestedFilename() || `floragold_pagina_${p}.xlsx`;
    const tmpFile = path.join(REPORT_DIR, `tmp_${Date.now()}_${p}_${slug(suggested) || 'export'}`);
    await download.saveAs(tmpFile);

    let parsed = parseDownloadedFile(tmpFile);
    fs.unlink(tmpFile, () => {});

    // Valida e ordena pelo modelo oficial do FloraGold.
    // Assim não aceitamos planilha resumida ou tabela visual com colunas faltando.
    if (parsed.headers?.length) {
      const official = validateOfficialReportSchema(parsed.rows, parsed.headers, `arquivo nativo página ${p}`);
      parsed = { headers: official.headers, rows: official.rows };
    }

    if (!headers.length && parsed.headers.length) headers = parsed.headers;

    let added = 0;
    for (const row of parsed.rows) {
      const key = headers.map(h => row[h] || '').join('||') || JSON.stringify(row);
      if (seen.has(key)) continue;
      seen.add(key);
      allRows.push(row);
      added++;
    }

    paginasCapturadas = p;
    const totals = await getPageTotals(page);
    if (totals.totalPages) totalPages = totals.totalPages;
    if (totals.totalRecords) totalRecords = totals.totalRecords;

    update({
      step: totalPages ? `Exportando arquivo do sistema ${p} de ${totalPages}` : `Exportando arquivo do sistema ${p}`,
      progress: totalPages ? Math.min(94, 48 + Math.round((p / totalPages) * 44)) : Math.min(92, 48 + p * 2),
      paginaAtual: p,
      totalPaginas: totalPages || 0,
      registrosLidos: allRows.length,
      linhasPorPagina: LINHAS_POR_PAGINA,
      velocidadeMedia: averagePageSpeed(started, p),
      tempoEstimadoRestante: totalPages ? estimateRemaining(started, p, totalPages) : '--',
      log: `Exportação nativa ${p}${totalPages ? ` de ${totalPages}` : ''}: ${added} linhas importadas do arquivo do sistema.`
    });

    // Se o botão nativo exportou tudo filtrado de uma vez, não precisa passar páginas.
    if (totalRecords && allRows.length >= totalRecords) break;
    if (totalPages && p >= totalPages) break;

    const moved = await clickNext(frame);
    if (!moved) break;
    await waitProcessingAfterNext(page);
  }

  if (!headers.length && allRows[0]) headers = Object.keys(allRows[0]);
  return { rows: allRows, headers, paginasCapturadas, totalExpected: totalRecords || allRows.length, sourceMode: 'native-export' };
}

async function waitProcessingAfterNext(page) {
  await page.waitForLoadState('networkidle', { timeout: 45000 }).catch(() => {});
  await waitForProcessingEnd(page).catch(() => {});
  await sleep(1000);
}

async function captureAllPages(page, update, started = Date.now()) {
  const rows = [];
  let headers = [];
  let paginasCapturadas = 0;
  const seen = new Set();
  let lastSignature = '';
  let totalPages = 0;
  let totalExpected = 0;

  for (let pageNumber = 1; pageNumber <= 250; pageNumber++) {
    const source = await bestTableSource(page);
    if (!source) break;

    const extracted = await source.frame.evaluate((linhasPorPagina) => {
      const text = el => (el?.innerText || el?.textContent || '').replace(/\s+/g, ' ').trim();
      const visible = el => !!(el && el.getClientRects().length);
      const tables = Array.from(document.querySelectorAll('table'))
        .filter(visible)
        .map(t => ({ table: t, cells: t.querySelectorAll('tbody td').length, len: text(t).length }))
        .sort((a, b) => (b.cells + b.len) - (a.cells + a.len));
      const table = tables[0]?.table;
      if (!table) return { headers: [], rows: [], signature: '', totalPages: 0, totalRecords: 0 };

      let headers = Array.from(table.querySelectorAll('thead tr:last-child th')).map(text).filter(Boolean);
      const rows = Array.from(table.querySelectorAll('tbody tr')).map(tr => Array.from(tr.querySelectorAll('td,th')).map(text))
        .filter(r => r.length && r.some(c => c) && !/nenhum registro|no matching records|carregando|processando/i.test(r.join(' ')));
      if (!headers.length && rows[0]) headers = rows[0].map((_, i) => `Coluna ${i + 1}`);
      const signature = rows.slice(0, 3).map(r => r.join('|')).join('§') + `#${rows.length}`;

      const pageText = text(document.body);
      const parseNumber = v => Number(String(v || '').replace(/\./g, '').replace(',', '.').replace(/[^0-9.]/g, '')) || 0;
      let totalRecords = 0;

      const patterns = [
        /de\s+([\d\.]+)\s+registros/i,
        /total\s+de\s+([\d\.]+)\s+registros/i,
        /([\d\.]+)\s+registros\s+no\s+total/i,
        /filtered\s+from\s+([\d\.]+)\s+total/i
      ];
      for (const p of patterns) {
        const m = pageText.match(p);
        if (m) { totalRecords = parseNumber(m[1]); break; }
      }

      let totalPages = totalRecords ? Math.max(1, Math.ceil(totalRecords / linhasPorPagina)) : 0;
      const pageNumbers = Array.from(document.querySelectorAll('.paginate_button, .pagination a, .pagination button, a, button'))
        .map(el => text(el))
        .filter(v => /^\d+$/.test(v))
        .map(Number);
      if (pageNumbers.length) totalPages = Math.max(totalPages, Math.max(...pageNumbers));

      return { headers, rows, signature, totalPages, totalRecords };
    }, LINHAS_POR_PAGINA);

    if (!extracted.rows.length && pageNumber === 1) break;
    if (!headers.length && extracted.headers.length) headers = extracted.headers;
    if (extracted.totalPages) totalPages = extracted.totalPages;
    if (extracted.totalRecords) totalExpected = extracted.totalRecords;

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
    const currentTotalPages = totalPages || 0;
    update({
      step: `Coletando página ${pageNumber}${currentTotalPages ? ` de ${currentTotalPages}` : ''}`,
      progress: currentTotalPages ? Math.min(92, 45 + Math.round((pageNumber / currentTotalPages) * 47)) : Math.min(90, 45 + Math.round(pageNumber * 2)),
      paginaAtual: pageNumber,
      totalPaginas: currentTotalPages,
      registrosLidos: rows.length,
      linhasPorPagina: LINHAS_POR_PAGINA,
      velocidadeMedia: averagePageSpeed(started, pageNumber),
      tempoEstimadoRestante: currentTotalPages ? estimateRemaining(started, pageNumber, currentTotalPages) : '--',
      log: `Página ${pageNumber}${currentTotalPages ? ` de ${currentTotalPages}` : ''}: ${added} linhas capturadas.`
    });

    if (currentTotalPages && pageNumber >= currentTotalPages) break;

    const moved = await clickNext(source.frame);
    if (!moved) break;

    await waitTableChange(source.frame, extracted.signature || lastSignature);
    lastSignature = extracted.signature;
  }

  return { rows, headers, paginasCapturadas, totalExpected };
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
