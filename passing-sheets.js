const fs = require('fs');
const path = require('path');
const { ensureDatabase, getDatabasePool, isDatabaseConfigured } = require('./database');

const DATA_FILE = path.join(__dirname, 'data', 'passing_sheets_pending.json');
const DEFAULT_SEND_HOUR = 9;
const MAX_PENDING_RECORDS = 20;
const sendingIds = new Set();
let pendingTablePromise = null;

function clean(value) {
  return String(value ?? '').replace(/\s+/g, ' ').trim();
}

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function cellValue(sheet, coord) {
  const cell = sheet?.cells?.[coord];
  if (cell && typeof cell === 'object' && Object.prototype.hasOwnProperty.call(cell, 'v')) {
    return cell.v;
  }
  return cell ?? '';
}

function setCellValue(sheet, coord, value) {
  if (!sheet) return;
  const current = sheet.cells?.[coord];
  sheet.cells = sheet.cells || {};
  sheet.cells[coord] = { ...(current && typeof current === 'object' ? current : {}), v: value };
  delete sheet.cells[coord].f;
}

function dateOnly(value) {
  const text = clean(value);
  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) return `${match[1]}-${match[2]}-${match[3]}`;
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return '';
  return [date.getFullYear(), String(date.getMonth() + 1).padStart(2, '0'), String(date.getDate()).padStart(2, '0')].join('-');
}

function addDays(value, amount) {
  const normalized = dateOnly(value);
  if (!normalized) return '';
  const date = new Date(`${normalized}T00:00:00`);
  date.setDate(date.getDate() + amount);
  return [date.getFullYear(), String(date.getMonth() + 1).padStart(2, '0'), String(date.getDate()).padStart(2, '0')].join('-');
}

function sendHour() {
  const configured = Number(process.env.PASSING_SEND_HOUR ?? DEFAULT_SEND_HOUR);
  if (!Number.isFinite(configured)) return DEFAULT_SEND_HOUR;
  return Math.min(23, Math.max(0, Math.floor(configured)));
}

function scheduledAtFor(loadDate) {
  const sendDate = addDays(loadDate, -1);
  if (!sendDate) return null;
  const hour = String(sendHour()).padStart(2, '0');
  return new Date(`${sendDate}T${hour}:00:00`).toISOString();
}

function readLocalStore() {
  if (!fs.existsSync(DATA_FILE)) return { version: 1, records: [] };
  try {
    const parsed = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
    return { version: 1, records: Array.isArray(parsed.records) ? parsed.records : [] };
  } catch {
    return { version: 1, records: [] };
  }
}

function writeLocalStore(store) {
  fs.mkdirSync(path.dirname(DATA_FILE), { recursive: true });
  store.records = trimRecords(store.records);
  const serialized = JSON.stringify(store, null, 2);
  const tempFile = `${DATA_FILE}.${process.pid}.${Date.now()}.${Math.random().toString(36).slice(2, 8)}.tmp`;
  try {
    fs.writeFileSync(tempFile, serialized, 'utf8');
    fs.renameSync(tempFile, DATA_FILE);
  } catch (error) {
    try {
      if (fs.existsSync(tempFile)) fs.unlinkSync(tempFile);
    } catch {
      // El archivo temporal se limpia en el siguiente guardado si Windows lo libera después.
    }
    if (error.code !== 'EPERM') throw error;
    // Fallback para Windows cuando el reemplazo atómico está bloqueado momentáneamente.
    fs.writeFileSync(DATA_FILE, serialized, 'utf8');
  }
}

function recordTime(record) {
  const value = Date.parse(record?.updatedAt || record?.createdAt || '');
  return Number.isFinite(value) ? value : 0;
}

function trimRecords(records = []) {
  if (records.length <= MAX_PENDING_RECORDS) return records;
  const removable = records.slice().sort((a, b) => {
    const aTerminal = ['ENVIADO', 'CANCELADO'].includes(a.status) ? 0 : 1;
    const bTerminal = ['ENVIADO', 'CANCELADO'].includes(b.status) ? 0 : 1;
    if (aTerminal !== bTerminal) return aTerminal - bTerminal;
    return recordTime(a) - recordTime(b);
  });
  const removeIds = new Set(removable.slice(0, records.length - MAX_PENDING_RECORDS).map(record => record.id));
  return records.filter(record => !removeIds.has(record.id));
}

async function trimDatabaseRecords(db) {
  const result = await db.query('SELECT id, record FROM passing_sheets_pending');
  const records = result.rows.map(row => row.record);
  const trimmed = trimRecords(records);
  if (trimmed.length === records.length) return;
  const keepIds = new Set(trimmed.map(record => record.id));
  const removeIds = records.filter(record => !keepIds.has(record.id)).map(record => record.id);
  await db.query('DELETE FROM passing_sheets_pending WHERE id = ANY($1::text[])', [removeIds]);
}

async function getPendingDatabase() {
  if (!isDatabaseConfigured()) return null;
  const db = getDatabasePool();
  await ensureDatabase();
  if (!pendingTablePromise) {
    pendingTablePromise = db.query(`
      CREATE TABLE IF NOT EXISTS passing_sheets_pending (
        id TEXT PRIMARY KEY,
        status TEXT NOT NULL,
        scheduled_at TIMESTAMPTZ NULL,
        record JSONB NOT NULL,
        updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
      )
    `).then(() => db.query(`
      CREATE INDEX IF NOT EXISTS passing_sheets_pending_schedule_idx
      ON passing_sheets_pending (scheduled_at)
    `)).catch(error => {
      pendingTablePromise = null;
      throw error;
    });
  }
  await pendingTablePromise;
  return db;
}

async function savePendingRecord(record) {
  const db = await getPendingDatabase();
  if (!db) return;
  await db.query(`
    INSERT INTO passing_sheets_pending (id, status, scheduled_at, record, updated_at)
    VALUES ($1, $2, $3, $4::jsonb, NOW())
    ON CONFLICT (id) DO UPDATE SET
      status = EXCLUDED.status,
      scheduled_at = EXCLUDED.scheduled_at,
      record = EXCLUDED.record,
      updated_at = NOW()
  `, [record.id, record.status, record.scheduledAt || null, JSON.stringify(record)]);
  await trimDatabaseRecords(db);
}

async function readPendingRecords() {
  const db = await getPendingDatabase();
  if (!db) return readLocalStore().records;
  const result = await db.query('SELECT record FROM passing_sheets_pending ORDER BY scheduled_at NULLS LAST, id');
  const records = result.rows.map(row => row.record);
  const trimmed = trimRecords(records);
  if (trimmed.length !== records.length) {
    const keepIds = new Set(trimmed.map(record => record.id));
    const removeIds = records.filter(record => !keepIds.has(record.id)).map(record => record.id);
    await db.query('DELETE FROM passing_sheets_pending WHERE id = ANY($1::text[])', [removeIds]);
  }
  return trimmed;
}

function passingFields(workbook) {
  const order = workbook?.sheets?.['PO Nuevo'];
  const passing = workbook?.sheets?.['PASSING SHEET'];
  return {
    poNumber: clean(cellValue(order, 'A9')) || 'Sin PO',
    customer: clean(cellValue(passing, 'B5')) || clean(cellValue(order, 'F11')) || 'Cliente',
    loadDate: dateOnly(cellValue(passing, 'B6')) || dateOnly(cellValue(order, 'B19')),
    releaseDate: dateOnly(cellValue(passing, 'B7')) || dateOnly(cellValue(order, 'F19')),
    truckLine: clean(cellValue(passing, 'B8')),
    account: clean(cellValue(passing, 'B9')),
    boxes: cellValue(passing, 'B10'),
    farmShipTo: clean(cellValue(order, 'B18')) || 'FLORAL TRADE',
    instructions: [13, 14, 15].map(row => clean(cellValue(passing, `B${row}`)))
  };
}

function buildManualPassingWorkbook(input = {}) {
  const customer = clean(input.customer);
  const loadDate = dateOnly(input.loadDate);
  const releaseDate = dateOnly(input.releaseDate);
  const rows = Array.isArray(input.lines) ? input.lines : [];
  if (!customer) throw new Error('Selecciona un cliente para el Passing Sheet.');
  if (!loadDate) throw new Error('Captura una fecha de Load Date válida.');
  if (!releaseDate) throw new Error('Captura una fecha de Release Date válida.');
  if (!rows.length) throw new Error('Agrega al menos una flor al Passing Sheet.');

  const orderSheet = { maxRow: 45, maxCol: 12, cells: {} };
  const passingSheet = { maxRow: 22, maxCol: 4, cells: {} };
  const set = (sheet, coord, value) => setCellValue(sheet, coord, value);
  const normalizedRows = [];
  let totalBoxes = 0;
  const dimensionGroups = new Map();

  rows.forEach((line, index) => {
    const flower = clean(line.flower);
    const cases = Number(line.cases);
    if (!flower || !Number.isFinite(cases) || cases <= 0) return;
    const row = 22 + normalizedRows.length;
    const boxType = clean(line.boxType);
    const dimension = clean(line.boxDimension);
    set(orderSheet, `A${row}`, clean(input.reference) || 'MANUAL');
    set(orderSheet, `B${row}`, flower);
    set(orderSheet, `C${row}`, clean(line.color));
    set(orderSheet, `D${row}`, line.stems ?? '');
    set(orderSheet, `E${row}`, line.pack ?? '');
    set(orderSheet, `F${row}`, boxType);
    set(orderSheet, `G${row}`, dimension);
    set(orderSheet, `H${row}`, cases);
    normalizedRows.push({ flower, cases, boxType, dimension });
    totalBoxes += cases;
    const label = [boxType, dimension].filter(Boolean).join('-');
    if (label) dimensionGroups.set(label, (dimensionGroups.get(label) || 0) + cases);
  });

  if (!normalizedRows.length) throw new Error('Las líneas deben tener flor y un número de cajas mayor que cero.');
  const dimensionText = Array.from(dimensionGroups, ([label, cases]) => `(${cases}) ${label}`).join(', ');
  const reference = clean(input.reference) || 'MANUAL';

  set(orderSheet, 'A9', reference);
  set(orderSheet, 'F11', customer);
  set(orderSheet, 'E12', clean(input.customerShipTo));
  set(orderSheet, 'F15', clean(input.truckLine));
  set(orderSheet, 'F16', input.account ?? '');
  set(orderSheet, 'B18', clean(input.farmShipTo) || 'FLORAL TRADE');
  set(orderSheet, 'B19', loadDate);
  set(orderSheet, 'F19', releaseDate);
  const totalRow = 22 + normalizedRows.length;
  set(orderSheet, `G${totalRow}`, 'TOTAL');
  set(orderSheet, `H${totalRow}`, totalBoxes);

  set(passingSheet, 'A2', 'Please receive the following shipment and release it to the designated trucking line as instructed below.');
  set(passingSheet, 'B4', 'Pacifica Produce Farms LLC');
  set(passingSheet, 'B5', customer);
  set(passingSheet, 'B6', loadDate);
  set(passingSheet, 'B7', releaseDate);
  set(passingSheet, 'B8', clean(input.truckLine));
  set(passingSheet, 'B9', input.account ?? '');
  set(passingSheet, 'B10', totalBoxes);
  set(passingSheet, 'B11', dimensionText);
  const instructions = Array.isArray(input.instructions) ? input.instructions : [];
  instructions.slice(0, 3).forEach((value, index) => set(passingSheet, `B${13 + index}`, clean(value)));

  return {
    filename: `Passing Sheet Manual - ${reference}`,
    sheets: { 'PO Nuevo': orderSheet, 'PASSING SHEET': passingSheet }
  };
}

async function queueManualPassingSheet(input = {}) {
  const workbook = buildManualPassingWorkbook(input);
  return queuePassingSheet(workbook, 'manual');
}

function createPendingRecord(workbook, source = 'order-send') {
  const orderSheet = clone(workbook?.sheets?.['PO Nuevo'] || { cells: {} });
  const passingSheet = clone(workbook?.sheets?.['PASSING SHEET'] || { cells: {} });
  const fields = passingFields(workbook);
  const now = new Date().toISOString();
  const id = `ps_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
  return {
    id,
    source,
    filename: workbook?.filename || 'workbook',
    poNumber: fields.poNumber,
    customer: fields.customer,
    loadDate: fields.loadDate,
    releaseDate: fields.releaseDate,
    truckLine: fields.truckLine,
    account: fields.account,
    boxes: fields.boxes,
    farmShipTo: fields.farmShipTo,
    instructions: fields.instructions,
    orderSheet,
    passingSheet,
    status: 'PENDIENTE',
    revision: 1,
    scheduledAt: scheduledAtFor(fields.loadDate),
    createdAt: now,
    updatedAt: now,
    sentAt: null,
    lastError: null,
    history: [{ at: now, event: 'CREATED', revision: 1, source }]
  };
}

async function queuePassingSheet(workbook, source = 'order-send') {
  const record = createPendingRecord(workbook, source);
  const db = await getPendingDatabase();
  if (db) {
    await savePendingRecord(record);
    return clone(record);
  }
  const store = readLocalStore();
  store.records.push(record);
  writeLocalStore(store);
  return clone(record);
}

async function listPendingPassingSheets() {
  const records = await readPendingRecords();
  return records
    .slice()
    .sort((a, b) => String(a.scheduledAt || '').localeCompare(String(b.scheduledAt || '')))
    .map(clone);
}

async function getPendingPassingSheet(id) {
  const records = await readPendingRecords();
  const record = records.find(item => item.id === id);
  return record ? clone(record) : null;
}

function pendingWorkbook(record) {
  return {
    filename: record.filename,
    sheets: {
      'PO Nuevo': clone(record.orderSheet),
      'PASSING SHEET': clone(record.passingSheet)
    }
  };
}

async function updatePendingPassingSheet(id, input = {}) {
  const db = await getPendingDatabase();
  const store = db ? null : readLocalStore();
  const record = db
    ? (await readPendingRecords()).find(item => item.id === id)
    : store.records.find(item => item.id === id);
  if (!record) return null;
  if (record.status === 'ENVIADO') throw new Error('Este Passing Sheet ya fue enviado. Crea una nueva revisión desde la orden.');

  const editable = ['loadDate', 'releaseDate', 'truckLine', 'account', 'boxes', 'farmShipTo'];
  const changes = {};
  for (const field of editable) {
    if (!Object.prototype.hasOwnProperty.call(input, field)) continue;
    const value = field.includes('Date') ? dateOnly(input[field]) : input[field];
    if (field.includes('Date') && input[field] && !value) throw new Error(`Fecha inválida en ${field}.`);
    if (String(record[field] ?? '') !== String(value ?? '')) changes[field] = { from: record[field] ?? '', to: value ?? '' };
    record[field] = value ?? '';
  }

  if (Array.isArray(input.instructions)) {
    const nextInstructions = input.instructions.slice(0, 3).map(clean);
    if (JSON.stringify(record.instructions) !== JSON.stringify(nextInstructions)) {
      changes.instructions = { from: record.instructions, to: nextInstructions };
      record.instructions = nextInstructions;
    }
  }

  if (Object.keys(changes).length === 0) return clone(record);
  setCellValue(record.passingSheet, 'B6', record.loadDate);
  setCellValue(record.passingSheet, 'B7', record.releaseDate);
  setCellValue(record.passingSheet, 'B8', record.truckLine);
  setCellValue(record.passingSheet, 'B9', record.account);
  setCellValue(record.passingSheet, 'B10', record.boxes);
  setCellValue(record.orderSheet, 'B18', record.farmShipTo);
  record.instructions.forEach((value, index) => setCellValue(record.passingSheet, `B${13 + index}`, value));
  record.revision += 1;
  record.status = 'PENDIENTE';
  record.scheduledAt = scheduledAtFor(record.loadDate);
  record.updatedAt = new Date().toISOString();
  record.lastError = null;
  record.history.push({
    at: record.updatedAt,
    event: 'EDITED',
    revision: record.revision,
    changes
  });
  if (db) await savePendingRecord(record);
  else writeLocalStore(store);
  return clone(record);
}

async function markPendingSent(id, result = {}) {
  const db = await getPendingDatabase();
  const store = db ? null : readLocalStore();
  const record = db
    ? (await readPendingRecords()).find(item => item.id === id)
    : store.records.find(item => item.id === id);
  if (!record) return null;
  record.status = 'ENVIADO';
  record.sentAt = new Date().toISOString();
  record.updatedAt = record.sentAt;
  record.lastError = null;
  record.history.push({
    at: record.sentAt,
    event: 'SENT',
    revision: record.revision,
    sentTo: result.sentTo || [],
    simulated: Boolean(result.simulated)
  });
  if (db) await savePendingRecord(record);
  else writeLocalStore(store);
  return clone(record);
}

async function cancelPendingPassingSheet(id) {
  const db = await getPendingDatabase();
  const store = db ? null : readLocalStore();
  const record = db
    ? (await readPendingRecords()).find(item => item.id === id)
    : store.records.find(item => item.id === id);
  if (!record) return null;
  if (record.status === 'ENVIADO') throw new Error('Este Passing Sheet ya fue enviado y no se puede cancelar.');
  if (sendingIds.has(id)) throw new Error('Este Passing Sheet se está enviando en este momento.');
  if (record.status === 'CANCELADO') return clone(record);

  const now = new Date().toISOString();
  record.status = 'CANCELADO';
  record.canceledAt = now;
  record.updatedAt = now;
  record.lastError = null;
  record.history.push({
    at: now,
    event: 'CANCELLED',
    revision: record.revision
  });
  if (db) await savePendingRecord(record);
  else writeLocalStore(store);
  return clone(record);
}

async function resumePendingPassingSheet(id) {
  const db = await getPendingDatabase();
  const store = db ? null : readLocalStore();
  const record = db
    ? (await readPendingRecords()).find(item => item.id === id)
    : store.records.find(item => item.id === id);
  if (!record) return null;
  if (record.status !== 'CANCELADO') throw new Error('Solo se pueden reanudar Passing Sheets cancelados.');

  const now = new Date().toISOString();
  record.status = 'PENDIENTE';
  record.scheduledAt = scheduledAtFor(record.loadDate);
  record.updatedAt = now;
  record.lastError = null;
  record.history.push({
    at: now,
    event: 'RESUMED',
    revision: record.revision,
    scheduledAt: record.scheduledAt
  });
  if (db) await savePendingRecord(record);
  else writeLocalStore(store);
  return clone(record);
}

async function markPendingError(id, error) {
  const db = await getPendingDatabase();
  const store = db ? null : readLocalStore();
  const record = db
    ? (await readPendingRecords()).find(item => item.id === id)
    : store.records.find(item => item.id === id);
  if (!record) return null;
  record.status = 'ERROR';
  record.lastError = String(error?.message || error);
  record.updatedAt = new Date().toISOString();
  record.history.push({ at: record.updatedAt, event: 'SEND_ERROR', revision: record.revision, error: record.lastError });
  if (db) await savePendingRecord(record);
  else writeLocalStore(store);
  return clone(record);
}

function canSendPending(id) {
  if (sendingIds.has(id)) return false;
  sendingIds.add(id);
  return true;
}

function finishSendingPending(id) {
  sendingIds.delete(id);
}

function isSendingPending(id) {
  return sendingIds.has(id);
}

function startPassingSheetScheduler(sendDue) {
  if (String(process.env.PASSING_AUTO_SEND).toLowerCase() !== 'true') return null;
  const run = async () => {
    const now = Date.now();
    for (const record of await listPendingPassingSheets()) {
      if (record.status !== 'PENDIENTE' && record.status !== 'ERROR') continue;
      if (!record.scheduledAt || new Date(record.scheduledAt).getTime() > now) continue;
      if (isSendingPending(record.id)) continue;
      await sendDue(record.id).catch(() => {});
    }
  };
  void run();
  return setInterval(run, 30_000);
}

module.exports = {
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
};
