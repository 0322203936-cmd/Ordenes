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
const { sendWholesaleEmail } = require('./email-sender');

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
      const email = await sendWholesaleEmail({
        workbook: parsed,
        pdfBuffer: pdf.buffer,
        pdfFileName: pdf.fileName,
        poscoBuffer: posco.buffer,
        poscoFileName: posco.fileName,
        boomPending: posco.boomPending
      });
      sendJson(res, 200, { ok: true, ...email, rows: posco.rows });
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
        if (fs.existsSync(STATE_FILE)) fs.unlinkSync(STATE_FILE);
        sendJson(res, 200, readLocalWorkbook());
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
  console.log('====================================================');
});

async function shutdown() {
  await closeDatabase();
  server.close(() => process.exit(0));
}

process.on('SIGTERM', shutdown);
process.on('SIGINT', shutdown);
