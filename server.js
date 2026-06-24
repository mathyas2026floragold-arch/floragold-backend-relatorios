require('dotenv').config();

const fs = require('fs');
const path = require('path');
const express = require('express');
const helmet = require('helmet');
const morgan = require('morgan');
const cors = require('cors');
const { randomUUID } = require('crypto');
const {
  runReportJob,
  validateAccessJob,
  readFilterOptionsJob,
  getRuntimeConfig
} = require('./src/robot');

const app = express();
const PORT = Number(process.env.PORT || 3000);
const REPORT_DIR = path.resolve(process.env.REPORT_DIR || './relatorios');
const HISTORY_PATH = path.join(REPORT_DIR, 'historico.json');
const JOB_TTL_MINUTES = Number(process.env.JOB_TTL_MINUTES || 180);
const LINHAS_POR_PAGINA = Number(process.env.LINHAS_POR_PAGINA || 3000);

function normalizeDateOnly(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  // Frontend novo envia YYYY-MM-DD. O FloraGold costuma usar DD/MM/AAAA.
  const iso = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return `${iso[3]}/${iso[2]}/${iso[1]}`;
  return raw;
}

function normalizeParams(params = {}) {
  return {
    ...params,
    dataInicial: normalizeDateOnly(params.dataInicial),
    dataFinal: normalizeDateOnly(params.dataFinal),
    horaInicial: '00:00',
    horaFinal: '23:59'
  };
}

fs.mkdirSync(REPORT_DIR, { recursive: true });
fs.mkdirSync(path.join(REPORT_DIR, 'erros'), { recursive: true });

const allowedOrigin = process.env.CORS_ORIGIN || '*';
app.use(cors({
  origin(origin, callback) {
    if (allowedOrigin === '*' || !origin) return callback(null, true);
    const list = allowedOrigin.split(',').map(v => v.trim()).filter(Boolean);
    return callback(null, list.includes(origin));
  },
  credentials: false
}));
app.use(helmet({ contentSecurityPolicy: false, crossOriginEmbedderPolicy: false }));
app.use(morgan('dev'));
app.use(express.json({ limit: '2mb' }));

const jobs = new Map();

function nowIso() {
  return new Date().toISOString();
}

function readHistory() {
  try {
    const data = JSON.parse(fs.readFileSync(HISTORY_PATH, 'utf8'));
    return Array.isArray(data) ? data : [];
  } catch {
    return [];
  }
}

function writeHistory(items) {
  fs.writeFileSync(HISTORY_PATH, JSON.stringify(items.slice(0, 300), null, 2), 'utf8');
}

function safeJob(job) {
  const copy = JSON.parse(JSON.stringify(job));
  if (copy.params) {
    delete copy.params.senha;
    copy.params.senha = '[PROTEGIDA]';
  }
  return copy;
}

function addJobLog(job, message) {
  job.logs.push({ at: nowIso(), message: String(message || '') });
}

function createUpdater(jobId) {
  return (patch = {}) => {
    const job = jobs.get(jobId);
    if (!job) return;

    const { log, ...rest } = patch;
    Object.assign(job, rest, { updatedAt: nowIso() });
    if (log) addJobLog(job, log);
    jobs.set(jobId, job);
  };
}

function cleanOldJobs() {
  const cutoff = Date.now() - JOB_TTL_MINUTES * 60 * 1000;
  for (const [id, job] of jobs.entries()) {
    const time = new Date(job.updatedAt || job.createdAt).getTime();
    if (Number.isFinite(time) && time < cutoff) jobs.delete(id);
  }
}
setInterval(cleanOldJobs, 15 * 60 * 1000).unref();

function validateParams(params) {
  const required = ['usuario', 'senha', 'tipoRelatorio', 'dataInicial', 'dataFinal'];
  const missing = required.filter(key => !String(params?.[key] || '').trim());
  if (missing.length) return `Campos obrigatórios ausentes: ${missing.join(', ')}`;
  return null;
}

function historySuccess(job, result) {
  const p = job.params || {};
  return {
    id: job.id,
    tipoRelatorio: p.tipoRelatorio,
    periodo: `${p.dataInicial} até ${p.dataFinal}`,
    operador: p.operador || 'Todos os operadores',
    fila: p.fila || 'Todas as filas',
    statusFiltro: p.statusFiltro || p.status || 'Todos os status',
    registros: result.totalRegistros || 0,
    paginas: result.paginasCapturadas || 0,
    tempoTotal: result.tempoTotal || '-',
    arquivoCsv: result.csvFile,
    arquivoXlsx: result.xlsxFile,
    status: 'Sucesso',
    createdAt: nowIso()
  };
}

function historyError(job, err) {
  const p = job.params || {};
  return {
    id: job.id,
    tipoRelatorio: p.tipoRelatorio || 'Relatório',
    periodo: `${p.dataInicial || ''} até ${p.dataFinal || ''}`.trim(),
    operador: p.operador || 'Todos os operadores',
    fila: p.fila || 'Todas as filas',
    registros: 0,
    paginas: 0,
    tempoTotal: '-',
    status: 'Falha',
    erro: err?.message || String(err || 'Erro desconhecido'),
    createdAt: nowIso()
  };
}

app.get('/', (req, res) => {
  res.type('html').send(`<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><title>Backend Relatórios+</title></head><body style="font-family:Arial;padding:32px"><h1>Backend Relatórios+ ativo</h1><p>Use o front-end no Netlify para operar a Central Automática de Relatórios.</p><p><a href="/health">/health</a> • <a href="/api/status">/api/status</a></p></body></html>`);
});

app.get('/health', (req, res) => {
  res.json({ ok: true, service: 'central-relatorios-backend', time: nowIso() });
});

app.get('/api/status', (req, res) => {
  res.json({ ok: true, ...getRuntimeConfig(), activeJobs: jobs.size, serverTime: nowIso() });
});

app.get('/api/historico', (req, res) => {
  res.json({ ok: true, items: readHistory().slice(0, Number(req.query.limit || 20)) });
});

app.delete('/api/historico', (req, res) => {
  writeHistory([]);
  res.json({ ok: true, message: 'Histórico limpo.' });
});

app.post('/api/validar-acesso', async (req, res) => {
  const params = normalizeParams(req.body || {});
  if (!params.usuario || !params.senha) {
    return res.status(400).json({ ok: false, error: 'Informe usuário e senha.' });
  }

  try {
    const result = await validateAccessJob(params);
    res.json({ ok: true, result });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message || 'Falha ao validar acesso.' });
  }
});

app.post('/api/opcoes', async (req, res) => {
  const params = normalizeParams(req.body || {});
  try {
    const result = await readFilterOptionsJob(params);
    res.json({ ok: true, result });
  } catch (err) {
    res.status(500).json({ ok: false, error: err.message || 'Falha ao carregar opções.' });
  }
});

app.post('/api/relatorios/gerar', (req, res) => {
  const params = normalizeParams(req.body || {});
  const error = validateParams(params);
  if (error) return res.status(400).json({ ok: false, error });

  const id = randomUUID();
  const job = {
    id,
    status: 'queued',
    progress: 0,
    step: 'Na fila',
    paginaAtual: 0,
    totalPaginas: 0,
    registrosLidos: 0,
    linhasPorPagina: LINHAS_POR_PAGINA,
    velocidadeMedia: '—',
    tempoEstimadoRestante: '--',
    logs: [],
    createdAt: nowIso(),
    updatedAt: nowIso(),
    result: null,
    error: null,
    params: { ...params, senha: '[PROTEGIDA]' }
  };

  jobs.set(id, job);
  addJobLog(job, 'Job criado.');
  const update = createUpdater(id);

  setImmediate(async () => {
    try {
      update({ status: 'running', step: 'Iniciando automação', progress: 2, log: 'Processo iniciado.' });
      const result = await runReportJob({ ...params, jobId: id }, update);
      update({ status: 'done', step: 'Relatório concluído', progress: 100, result, log: 'Relatório gerado com sucesso.' });

      const savedJob = jobs.get(id) || job;
      const history = readHistory();
      history.unshift(historySuccess(savedJob, result));
      writeHistory(history);
    } catch (err) {
      update({ status: 'error', step: 'Erro na execução', progress: 100, error: err.message || String(err), log: `Erro: ${err.message || err}` });

      const savedJob = jobs.get(id) || job;
      const history = readHistory();
      history.unshift(historyError(savedJob, err));
      writeHistory(history);
    }
  });

  res.json({ ok: true, jobId: id });
});

app.get('/api/jobs/:id', (req, res) => {
  const job = jobs.get(req.params.id);
  if (!job) return res.status(404).json({ ok: false, error: 'Job não encontrado.' });
  res.json({ ok: true, job: safeJob(job) });
});

app.get('/api/download/:filename', (req, res) => {
  const filename = path.basename(req.params.filename || '');
  if (!filename) return res.status(400).send('Arquivo inválido.');

  const filePath = path.join(REPORT_DIR, filename);
  if (!filePath.startsWith(REPORT_DIR) || !fs.existsSync(filePath)) {
    return res.status(404).send('Arquivo não encontrado.');
  }
  res.download(filePath);
});

app.use((req, res) => res.status(404).json({ ok: false, error: 'Rota não encontrada.' }));

app.listen(PORT, () => {
  console.log(`Backend Relatórios+ rodando na porta ${PORT}`);
  console.log(`DEMO_MODE=${process.env.DEMO_MODE || 'true'}`);
});
