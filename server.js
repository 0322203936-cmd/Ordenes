require('dotenv').config();

const http = require('http');
const fs = require('fs');
const path = require('path');
const {
  closeDatabase,
  getWorkbook,
  isDatabaseConfigured,
  resetWorkbook,
  saveWorkbook
} = require('./database');
const { createPoscoWorkbook } = require('./posco-export');
const { createOrderPdf } = require('./order-pdf');
const { createPassingSheetPdf } = require('./passing-pdf');
const { sendPassingSheetEmail, sendWholesaleEmail } = require('./email-sender');
const {
  canSendPending,
  cancelPendingPassingSheet,
  finishSendingPending,
  getPendingPassingSheet,
  listPendingPassingSheets,
  markPendingError,
  markPendingSent,
  pendingWorkbook,
  queueManualPassingSheet,
  queuePassingSheet,
  resumePendingPassingSheet,
  startPassingSheetScheduler,
  updatePendingPassingSheet
} = require('./passing-sheets');

const PORT = process.env.PORT || 3050;
const ROOT_DIR = __dirname;
const DATA_DIR = path.join(ROOT_DIR, 'data');
const INITIAL_FILE = path.join(DATA_DIR, 'initial_workbook.json');
const STATE_FILE = path.join(DATA_DIR, 'workbook_state.json');

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon'
};

function sendJson(res, statusCode, payload) {
  res.writeHead(statusCode, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify(payload));
}

function readLocalWorkbook() {
  const targetFile = fs.existsSync(STATE_FILE) ? STATE_FILE : INITIAL_FILE;
  return JSON.parse(fs.readFileSync(targetFile, 'utf8'));
}

function readRequestBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', chunk => {
      body += chunk;
      if (body.length > 50 * 1024 * 1024) {
        reject(new Error('Request body too large'));
        req.destroy();
      }
    });
    req.on('end', () => resolve(body));
    req.on('error', reject);
  });
}

function localEmailPreviewEnabled() {
  return String(process.env.LOCAL_EMAIL_PREVIEW).toLowerCase() === 'true';
}

async function sendPendingPassingSheet(id) {
  const pending = await getPendingPassingSheet(id);
  if (!pending) throw new Error('Passing Sheet pendiente no encontrado.');
  if (pending.status === 'ENVIADO') return { record: pending, sentTo: [] };
  if (!canSendPending(id)) throw new Error('Este Passing Sheet ya se está enviando.');

  try {
    const workbook = pendingWorkbook(pending);
    const pdf = await createPassingSheetPdf(workbook);
    const email = localEmailPreviewEnabled()
      ? { sentTo: ['LOCAL PREVIEW'], simulated: true }
      : await sendPassingSheetEmail({
        workbook,
        pdfBuffer: pdf.buffer,
        pdfFileName: pdf.fileName
      });
    const record = await markPendingSent(id, email);
    return { record, ...email };
  } catch (error) {
    await markPendingError(id, error);
    throw error;
  } finally {
    finishSendingPending(id);
  }
}

const server = http.createServer(async (req, res) => {
  const parsedUrl = new URL(req.url, `http://${req.headers.host}`);
  const pathname = decodeURIComponent(parsedUrl.pathname);

  if (pathname === '/api/posco/export' && req.method === 'POST') {
    try {
      const parsed = JSON.parse(await readRequestBody(req));
      const result = await createPoscoWorkbook(parsed);
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${result.fileName.replace(/[\r\n"]/g, '_')}"`,
        'Content-Length': result.buffer.length,
        'X-BOM-Pending': encodeURIComponent(JSON.stringify(result.boomPending)),
        'X-POSCO-Rows': String(result.rows)
      });
      res.end(result.buffer);
    } catch (err) {
      sendJson(res, err.status || 400, { error: err.message });
    }
    return;
  }

  if (pathname === '/api/order/send' && req.method === 'POST') {
    try {
      const parsed = JSON.parse(await readRequestBody(req));
      const posco = await createPoscoWorkbook(parsed);
      const pdf = await createOrderPdf(parsed);
      const email = localEmailPreviewEnabled()
        ? { sentTo: ['LOCAL PREVIEW'], pendingBom: posco.boomPending.length, simulated: true }
        : await sendWholesaleEmail({
          workbook: parsed,
          pdfBuffer: pdf.buffer,
          pdfFileName: pdf.fileName,
          poscoBuffer: posco.buffer,
          poscoFileName: posco.fileName,
          boomPending: posco.boomPending
        });
      const pending = await queuePassingSheet(parsed);
      sendJson(res, 200, {
        ok: true,
        ...email,
        rows: posco.rows,
        passingQueued: true,
        passingPendingId: pending.id,
        passingScheduledAt: pending.scheduledAt,
        passingSent: false
      });
    } catch (err) {
      const status = err.message?.includes('Falta configurar') || err.message?.includes('Microsoft Graph') ? 503 : 400;
      sendJson(res, err.status || status, { error: err.message });
    }
    return;
  }

  if (pathname === '/api/passing/pending' && req.method === 'GET') {
    sendJson(res, 200, { ok: true, records: await listPendingPassingSheets() });
    return;
  }

  if (pathname === '/api/passing/pending' && req.method === 'POST') {
    try {
      const input = JSON.parse(await readRequestBody(req));
      const record = await queueManualPassingSheet(input);
      sendJson(res, 200, { ok: true, record });
    } catch (err) {
      sendJson(res, 400, { error: err.message });
    }
    return;
  }

  const pendingMatch = pathname.match(/^\/api\/passing\/pending\/([^/]+)$/);
  if (pendingMatch && req.method === 'PATCH') {
    try {
      const input = JSON.parse(await readRequestBody(req));
      const updated = await updatePendingPassingSheet(pendingMatch[1], input);
      if (!updated) {
        sendJson(res, 404, { error: 'Passing Sheet pendiente no encontrado.' });
      } else {
        sendJson(res, 200, { ok: true, record: updated });
      }
    } catch (err) {
      sendJson(res, 400, { error: err.message });
    }
    return;
  }

  const pendingSendMatch = pathname.match(/^\/api\/passing\/pending\/([^/]+)\/send$/);
  if (pendingSendMatch && req.method === 'POST') {
    try {
      const result = await sendPendingPassingSheet(pendingSendMatch[1]);
      sendJson(res, 200, { ok: true, ...result, record: result.record });
    } catch (err) {
      const status = err.message?.includes('Falta configurar') || err.message?.includes('Microsoft Graph') ? 503 : 400;
      sendJson(res, err.status || status, { error: err.message });
    }
    return;
  }

  const pendingCancelMatch = pathname.match(/^\/api\/passing\/pending\/([^/]+)\/cancel$/);
  if (pendingCancelMatch && req.method === 'POST') {
    try {
      const record = await cancelPendingPassingSheet(pendingCancelMatch[1]);
      if (!record) {
        sendJson(res, 404, { error: 'Passing Sheet pendiente no encontrado.' });
      } else {
        sendJson(res, 200, { ok: true, record });
      }
    } catch (err) {
      sendJson(res, 400, { error: err.message });
    }
    return;
  }

  const pendingResumeMatch = pathname.match(/^\/api\/passing\/pending\/([^/]+)\/resume$/);
  if (pendingResumeMatch && req.method === 'POST') {
    try {
      const record = await resumePendingPassingSheet(pendingResumeMatch[1]);
      if (!record) {
        sendJson(res, 404, { error: 'Passing Sheet pendiente no encontrado.' });
      } else {
        sendJson(res, 200, { ok: true, record });
      }
    } catch (err) {
      sendJson(res, 400, { error: err.message });
    }
    return;
  }

  if (pathname === '/api/passing/send' && req.method === 'POST') {
    try {
      const parsed = JSON.parse(await readRequestBody(req));
      const pdf = await createPassingSheetPdf(parsed);
      const email = await sendPassingSheetEmail({
        workbook: parsed,
        pdfBuffer: pdf.buffer,
        pdfFileName: pdf.fileName
      });
      sendJson(res, 200, { ok: true, ...email });
    } catch (err) {
      const status = err.message?.includes('Falta configurar') || err.message?.includes('Microsoft Graph') ? 503 : 400;
      sendJson(res, err.status || status, { error: err.message });
    }
    return;
  }

  // API: Get Workbook State
  if (pathname === '/api/workbook' && req.method === 'GET') {
    try {
      const workbook = isDatabaseConfigured() ? await getWorkbook() : readLocalWorkbook();
      sendJson(res, 200, workbook);
    } catch (err) {
      sendJson(res, 500, { error: err.message });
    }
    return;
  }

  // API: Save Workbook State
  if (pathname === '/api/workbook' && req.method === 'POST') {
    try {
      const parsed = JSON.parse(await readRequestBody(req));
      if (isDatabaseConfigured()) {
        sendJson(res, 200, await saveWorkbook(parsed));
      } else {
        fs.writeFileSync(STATE_FILE, JSON.stringify(parsed, null, 2), 'utf8');
        sendJson(res, 200, { ok: true, savedAt: new Date().toISOString(), storage: 'json-local' });
      }
    } catch (err) {
      sendJson(res, 400, { error: err.message });
    }
    return;
  }

  // API: Reset Workbook to Initial .xlsm State
  if (pathname === '/api/reset' && req.method === 'POST') {
    try {
      if (isDatabaseConfigured()) {
        sendJson(res, 200, await resetWorkbook());
      } else {
        const initial = JSON.parse(fs.readFileSync(INITIAL_FILE, 'utf8'));
        const current = fs.existsSync(STATE_FILE) ? readLocalWorkbook() : initial;
        for (const sheetName of ['Productos', 'Transporte']) {
          if (current?.sheets?.[sheetName]) initial.sheets[sheetName] = current.sheets[sheetName];
        }
        fs.writeFileSync(STATE_FILE, JSON.stringify(initial, null, 2), 'utf8');
        sendJson(res, 200, initial);
      }
    } catch (err) {
      sendJson(res, 500, { error: err.message });
    }
    return;
  }

  // Serve Static Files
  const filePath = pathname === '/' ? path.join(ROOT_DIR, 'index.html') : path.join(ROOT_DIR, pathname);
  if (!filePath.startsWith(ROOT_DIR)) {
    res.writeHead(403);
    res.end('Forbidden');
    return;
  }

  if (fs.existsSync(filePath) && fs.statSync(filePath).isFile()) {
    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || 'application/octet-stream';
    res.writeHead(200, { 'Content-Type': contentType });
    fs.createReadStream(filePath).pipe(res);
  } else {
    res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('No encontrado');
  }
});

server.listen(PORT, '0.0.0.0', () => {
  console.log('====================================================');
  console.log('  SISTEMA EXCEL: ORDENES WHOLESALE (Pacifica Farms)');
  console.log(`  URL: http://localhost:${PORT}`);
  console.log(`  Almacenamiento: ${isDatabaseConfigured() ? 'Neon PostgreSQL' : 'JSON local (DATABASE_URL no configurada)'}`);
  console.log(`  Passing Sheet automático: ${String(process.env.PASSING_AUTO_SEND).toLowerCase() === 'true' ? 'activo' : 'manual/local'}`);
  console.log('====================================================');
});

const passingScheduler = startPassingSheetScheduler(async id => {
  await sendPendingPassingSheet(id);
});

async function shutdown() {
  if (passingScheduler) clearInterval(passingScheduler);
  await closeDatabase();
  server.close(() => process.exit(0));
}

process.on('SIGTERM', shutdown);
process.on('SIGINT', shutdown);
