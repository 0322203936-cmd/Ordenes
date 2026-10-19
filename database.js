const fs = require('fs');
const path = require('path');
const { Pool } = require('pg');

const DATA_DIR = path.join(__dirname, 'data');
const INITIAL_FILE = path.join(DATA_DIR, 'initial_workbook.json');

let pool = null;
let initializationPromise = null;

function isDatabaseConfigured() {
  return Boolean(process.env.DATABASE_URL);
}

function getPool() {
  if (!isDatabaseConfigured()) return null;
  if (!pool) {
    pool = new Pool({
      connectionString: process.env.DATABASE_URL,
      ssl: { rejectUnauthorized: false },
      max: 5,
      idleTimeoutMillis: 30000,
      connectionTimeoutMillis: 10000
    });
  }
  return pool;
}

function readInitialWorkbook() {
  return JSON.parse(fs.readFileSync(INITIAL_FILE, 'utf8'));
}

async function ensureDatabase() {
  const db = getPool();
  if (!db) return;
  if (!initializationPromise) {
    initializationPromise = (async () => {
      await db.query(`
        CREATE TABLE IF NOT EXISTS workbook_state (
          id INTEGER PRIMARY KEY CHECK (id = 1),
          filename TEXT NOT NULL,
          workbook JSONB NOT NULL,
          updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
        )
      `);

      const result = await db.query('SELECT id FROM workbook_state WHERE id = 1');
      if (result.rowCount === 0) {
        const initial = readInitialWorkbook();
        await db.query(
          `INSERT INTO workbook_state (id, filename, workbook)
           VALUES (1, $1, $2::jsonb)`,
          [initial.filename || 'workbook', JSON.stringify(initial)]
        );
        console.log('Neon inicializado con initial_workbook.json');
      }
    })().catch(error => {
      initializationPromise = null;
      throw error;
    });
  }
  await initializationPromise;
}

async function getWorkbook() {
  const db = getPool();
  if (!db) return null;
  await ensureDatabase();
  const result = await db.query('SELECT workbook FROM workbook_state WHERE id = 1');
  return result.rows[0]?.workbook || readInitialWorkbook();
}

async function saveWorkbook(workbook) {
  const db = getPool();
  if (!db) return null;
  await ensureDatabase();
  await db.query(
    `INSERT INTO workbook_state (id, filename, workbook, updated_at)
     VALUES (1, $1, $2::jsonb, NOW())
     ON CONFLICT (id) DO UPDATE SET
       filename = EXCLUDED.filename,
       workbook = EXCLUDED.workbook,
       updated_at = NOW()`,
    [workbook.filename || 'workbook', JSON.stringify(workbook)]
  );
  return { ok: true, savedAt: new Date().toISOString(), storage: 'neon' };
}

async function resetWorkbook() {
  const initial = readInitialWorkbook();
  const db = getPool();
  if (!db) return initial;
  await ensureDatabase();
  await db.query(
    `INSERT INTO workbook_state (id, filename, workbook, updated_at)
     VALUES (1, $1, $2::jsonb, NOW())
     ON CONFLICT (id) DO UPDATE SET
       filename = EXCLUDED.filename,
       workbook = EXCLUDED.workbook,
       updated_at = NOW()`,
    [initial.filename || 'workbook', JSON.stringify(initial)]
  );
  return initial;
}

async function closeDatabase() {
  if (pool) await pool.end();
}

module.exports = {
  closeDatabase,
  getWorkbook,
  isDatabaseConfigured,
  resetWorkbook,
  saveWorkbook
};
