import type Database from 'better-sqlite3';
import initSqlJs, { type Database as SqlJsDatabase, type SqlValue } from 'sql.js';
import sqlWasmUrl from 'sql.js/dist/sql-wasm-browser.wasm?url';

const IDB_NAME = 'fmcg_bi_client_store';
const IDB_STORE = 'sqlite_snapshots';
const IDB_KEY = 'main_db_v1';

const SCHEMA_SQL = `
  PRAGMA foreign_keys = ON;

  CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE NOT NULL,
    full_name TEXT NOT NULL,
    email TEXT,
    role TEXT NOT NULL CHECK(role IN ('admin', 'sector_manager', 'standard_user')),
    allowed_sectors TEXT NOT NULL DEFAULT '[]',
    password_hash TEXT NOT NULL,
    must_change_password INTEGER NOT NULL DEFAULT 0,
    is_active INTEGER NOT NULL DEFAULT 1,
    created_at TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS import_batches (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_uuid TEXT UNIQUE NOT NULL,
    filename TEXT NOT NULL,
    imported_by_id INTEGER REFERENCES users(id),
    imported_by_name TEXT NOT NULL,
    status TEXT NOT NULL CHECK(status IN ('active', 'archived')),
    vf05n_count INTEGER NOT NULL DEFAULT 0,
    vf05n_sales_total REAL NOT NULL DEFAULT 0,
    vf05n_tax_total REAL NOT NULL DEFAULT 0,
    zsales_count INTEGER NOT NULL DEFAULT 0,
    zsales_sales_total REAL NOT NULL DEFAULT 0,
    zsales_tax_total REAL NOT NULL DEFAULT 0,
    unlinked_branch_count INTEGER NOT NULL DEFAULT 0,
    unlinked_sector_count INTEGER NOT NULL DEFAULT 0,
    has_category_targets_sheet INTEGER NOT NULL DEFAULT 0,
    validation_report TEXT NOT NULL DEFAULT '{}',
    created_at TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS sectors (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_id INTEGER NOT NULL REFERENCES import_batches(id) ON DELETE CASCADE,
    sector_code TEXT NOT NULL,
    sector_name TEXT NOT NULL,
    monthly_target REAL NOT NULL DEFAULT 0,
    is_original_sector INTEGER NOT NULL DEFAULT 1
  );

  CREATE TABLE IF NOT EXISTS departments_chains (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_id INTEGER NOT NULL REFERENCES import_batches(id) ON DELETE CASCADE,
    chain_code TEXT NOT NULL,
    chain_name TEXT NOT NULL,
    department_name TEXT NOT NULL,
    sector_name TEXT NOT NULL,
    supervisor_name TEXT NOT NULL DEFAULT ''
  );

  CREATE TABLE IF NOT EXISTS branches (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_id INTEGER NOT NULL REFERENCES import_batches(id) ON DELETE CASCADE,
    branch_code TEXT NOT NULL,
    branch_name TEXT NOT NULL,
    sector_name TEXT NOT NULL,
    branch_target REAL NOT NULL DEFAULT 0,
    chain_code TEXT,
    chain_name TEXT NOT NULL DEFAULT '',
    rep_code TEXT,
    rep_name TEXT NOT NULL DEFAULT '',
    supervisor_name TEXT NOT NULL DEFAULT ''
  );

  CREATE TABLE IF NOT EXISTS sales_reps (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_id INTEGER NOT NULL REFERENCES import_batches(id) ON DELETE CASCADE,
    rep_code TEXT NOT NULL,
    rep_name TEXT NOT NULL,
    sector_name TEXT,
    supervisor_name TEXT,
    monthly_target REAL NOT NULL DEFAULT 0
  );

  CREATE TABLE IF NOT EXISTS items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_id INTEGER NOT NULL REFERENCES import_batches(id) ON DELETE CASCADE,
    item_code TEXT NOT NULL,
    item_name TEXT NOT NULL,
    category_name TEXT,
    brand_name TEXT,
    unit_name TEXT,
    chains_category TEXT NOT NULL DEFAULT '',
    army_category TEXT NOT NULL DEFAULT '',
    top_retail_category TEXT NOT NULL DEFAULT ''
  );

  CREATE TABLE IF NOT EXISTS category_targets (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_id INTEGER NOT NULL REFERENCES import_batches(id) ON DELETE CASCADE,
    sector_name TEXT NOT NULL DEFAULT '',
    job_title TEXT NOT NULL,
    employee_code TEXT NOT NULL DEFAULT '',
    employee_name TEXT NOT NULL,
    category_name TEXT NOT NULL,
    target_value REAL NOT NULL DEFAULT 0
  );

  CREATE TABLE IF NOT EXISTS competitions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    is_active INTEGER NOT NULL DEFAULT 1,
    created_by_id INTEGER REFERENCES users(id),
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS competition_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    competition_id INTEGER NOT NULL REFERENCES competitions(id) ON DELETE CASCADE,
    item_code TEXT NOT NULL,
    item_name TEXT NOT NULL DEFAULT '',
    equivalent_factor REAL NOT NULL DEFAULT 1,
    UNIQUE(competition_id, item_code)
  );

  CREATE TABLE IF NOT EXISTS competition_targets (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    competition_id INTEGER NOT NULL REFERENCES competitions(id) ON DELETE CASCADE,
    participant_key TEXT NOT NULL,
    target_qty REAL NOT NULL DEFAULT 0,
    industrial_target_qty REAL NOT NULL DEFAULT 0,
    commercial_target_qty REAL NOT NULL DEFAULT 0,
    UNIQUE(competition_id, participant_key)
  );

  CREATE TABLE IF NOT EXISTS item_category_targets (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    target_key TEXT UNIQUE NOT NULL,
    sector_name TEXT NOT NULL,
    job_title TEXT NOT NULL,
    employee_name TEXT NOT NULL,
    category_name TEXT NOT NULL,
    target_value REAL NOT NULL DEFAULT 0,
    updated_at TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS sales_records (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_id INTEGER NOT NULL REFERENCES import_batches(id) ON DELETE CASCADE,
    source_type TEXT NOT NULL CHECK(source_type IN ('VF05N', 'ZSALES')),
    doc_number TEXT,
    billing_date TEXT,
    payer_code TEXT,
    payer_name TEXT,
    sold_to_code TEXT,
    sold_to_name TEXT,
    item_code TEXT,
    item_name TEXT,
    rep_code TEXT,
    quantity REAL NOT NULL DEFAULT 0,
    sales_value REAL NOT NULL DEFAULT 0,
    tax_value REAL NOT NULL DEFAULT 0,
    resolved_sector TEXT,
    resolved_department TEXT,
    resolved_chain_name TEXT,
    resolved_branch_code TEXT,
    resolved_branch_name TEXT,
    resolved_rep_code TEXT,
    resolved_rep_name TEXT,
    resolved_item_name TEXT,
    is_return INTEGER NOT NULL DEFAULT 0
  );

  CREATE TABLE IF NOT EXISTS runtime_error_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    kind TEXT NOT NULL DEFAULT 'runtime_error',
    tab_name TEXT NOT NULL DEFAULT '',
    message TEXT NOT NULL,
    stack TEXT NOT NULL DEFAULT '',
    source_url TEXT NOT NULL DEFAULT '',
    line_no INTEGER NOT NULL DEFAULT 0,
    col_no INTEGER NOT NULL DEFAULT 0,
    user_id INTEGER,
    username TEXT NOT NULL DEFAULT '',
    resolved INTEGER NOT NULL DEFAULT 0,
    ai_analysis TEXT NOT NULL DEFAULT '',
    occurred_at TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS audit_runs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    triggered_by_id INTEGER REFERENCES users(id),
    triggered_by_name TEXT NOT NULL DEFAULT '',
    health_score INTEGER NOT NULL DEFAULT 100,
    total_findings INTEGER NOT NULL DEFAULT 0,
    critical_count INTEGER NOT NULL DEFAULT 0,
    high_count INTEGER NOT NULL DEFAULT 0,
    medium_count INTEGER NOT NULL DEFAULT 0,
    low_count INTEGER NOT NULL DEFAULT 0,
    findings_json TEXT NOT NULL DEFAULT '[]',
    ai_summary TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS audit_fix_backups (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    finding_code TEXT NOT NULL,
    action_type TEXT NOT NULL,
    action_label TEXT NOT NULL,
    summary TEXT NOT NULL,
    snapshot_json TEXT NOT NULL,
    created_by_id INTEGER REFERENCES users(id),
    created_by_name TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    restored_at TEXT
  );

  CREATE INDEX IF NOT EXISTS idx_sales_batch_source ON sales_records(batch_id, source_type);
  CREATE INDEX IF NOT EXISTS idx_sales_sector ON sales_records(batch_id, source_type, resolved_sector);
  CREATE INDEX IF NOT EXISTS idx_sales_branch ON sales_records(batch_id, source_type, resolved_branch_code);
  CREATE INDEX IF NOT EXISTS idx_sales_date ON sales_records(batch_id, source_type, billing_date);
  CREATE INDEX IF NOT EXISTS idx_branches_batch ON branches(batch_id, branch_code);
  CREATE INDEX IF NOT EXISTS idx_chains_batch ON departments_chains(batch_id, chain_code);
  CREATE INDEX IF NOT EXISTS idx_cat_targets_batch ON category_targets(batch_id, sector_name, job_title);
  CREATE INDEX IF NOT EXISTS idx_comp_items_comp ON competition_items(competition_id);
  CREATE INDEX IF NOT EXISTS idx_comp_targets_comp ON competition_targets(competition_id);
  CREATE INDEX IF NOT EXISTS idx_item_cat_targets_key ON item_category_targets(target_key);
  CREATE INDEX IF NOT EXISTS idx_runtime_errors_resolved ON runtime_error_logs(resolved, occurred_at);
  CREATE INDEX IF NOT EXISTS idx_audit_runs_created ON audit_runs(created_at);
  CREATE INDEX IF NOT EXISTS idx_audit_backups_created ON audit_fix_backups(created_at);
`;

function openIdb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = window.indexedDB.open(IDB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(IDB_STORE)) {
        db.createObjectStore(IDB_STORE);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error('IndexedDB open error'));
  });
}

async function readIdbSnapshot(): Promise<Uint8Array | null> {
  if (typeof window === 'undefined' || !window.indexedDB) return null;
  try {
    const idb = await openIdb();
    return await new Promise((resolve, reject) => {
      const tx = idb.transaction(IDB_STORE, 'readonly');
      const store = tx.objectStore(IDB_STORE);
      const req = store.get(IDB_KEY);
      req.onsuccess = () => {
        const val = req.result;
        if (val instanceof Uint8Array) resolve(val);
        else if (val instanceof ArrayBuffer) resolve(new Uint8Array(val));
        else resolve(null);
      };
      req.onerror = () => reject(req.error);
    });
  } catch {
    return null;
  }
}

async function writeIdbSnapshot(bytes: Uint8Array): Promise<void> {
  if (typeof window === 'undefined' || !window.indexedDB) return;
  try {
    const idb = await openIdb();
    await new Promise<void>((resolve, reject) => {
      const tx = idb.transaction(IDB_STORE, 'readwrite');
      const store = tx.objectStore(IDB_STORE);
      const req = store.put(bytes, IDB_KEY);
      req.onsuccess = () => resolve();
      req.onerror = () => reject(req.error);
    });
  } catch {
    // Best-effort persistence
  }
}

function normalizeSqlParams(args: unknown[]): SqlValue[] {
  const flat: unknown[] =
    args.length === 1 && Array.isArray(args[0]) ? args[0] : args;
  return flat.map((v) => {
    if (v === undefined || v === null) return null;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (typeof v === 'number' || typeof v === 'string' || v instanceof Uint8Array) {
      return v;
    }
    return String(v);
  });
}

function createBetterSqliteShim(rawDb: SqlJsDatabase): Database.Database {
  let txDepth = 0;

  const adapter = {
    exec(sql: string) {
      rawDb.exec(sql);
      return adapter;
    },
    pragma(pragmaStr: string) {
      const clean = pragmaStr.trim();
      if (clean.toLowerCase().startsWith('journal_mode')) {
        return 'memory';
      }
      const res = rawDb.exec(`PRAGMA ${clean}`);
      if (res.length > 0 && res[0].values.length > 0) {
        return res[0].values[0][0];
      }
      return undefined;
    },
    prepare(sql: string) {
      return {
        all(...args: unknown[]) {
          const stmt = rawDb.prepare(sql);
          try {
            const params = normalizeSqlParams(args);
            if (params.length > 0) stmt.bind(params);
            const rows: Record<string, unknown>[] = [];
            while (stmt.step()) {
              rows.push(stmt.getAsObject());
            }
            return rows;
          } finally {
            stmt.free();
          }
        },
        get(...args: unknown[]) {
          const stmt = rawDb.prepare(sql);
          try {
            const params = normalizeSqlParams(args);
            if (params.length > 0) stmt.bind(params);
            if (stmt.step()) {
              return stmt.getAsObject();
            }
            return undefined;
          } finally {
            stmt.free();
          }
        },
        run(...args: unknown[]) {
          const stmt = rawDb.prepare(sql);
          try {
            const params = normalizeSqlParams(args);
            if (params.length > 0) stmt.bind(params);
            stmt.step();
          } finally {
            stmt.free();
          }
          const changes = rawDb.getRowsModified();
          const lastRowRes = rawDb.exec('SELECT last_insert_rowid() AS id');
          const lastInsertRowid =
            lastRowRes.length > 0 && lastRowRes[0].values.length > 0
              ? Number(lastRowRes[0].values[0][0])
              : 0;
          return { changes, lastInsertRowid };
        },
      };
    },
    transaction<F extends (...args: never[]) => unknown>(fn: F) {
      return (...args: Parameters<F>): ReturnType<F> => {
        const spName = `sp_client_${txDepth}`;
        if (txDepth === 0) {
          rawDb.run('BEGIN TRANSACTION;');
        } else {
          rawDb.run(`SAVEPOINT ${spName};`);
        }
        txDepth++;
        try {
          const result = fn(...args);
          txDepth--;
          if (txDepth === 0) {
            rawDb.run('COMMIT;');
          } else {
            rawDb.run(`RELEASE SAVEPOINT ${spName};`);
          }
          return result as ReturnType<F>;
        } catch (err) {
          txDepth--;
          try {
            if (txDepth === 0) {
              rawDb.run('ROLLBACK;');
            } else {
              rawDb.run(`ROLLBACK TO SAVEPOINT ${spName};`);
            }
          } catch {
            // ignore rollback secondary error
          }
          throw err;
        }
      };
    },
  };

  return adapter as unknown as Database.Database;
}

let clientRawDb: SqlJsDatabase | null = null;
let clientDbAdapter: Database.Database | null = null;
let initPromise: Promise<Database.Database> | null = null;

export async function getClientDb(): Promise<Database.Database> {
  if (clientDbAdapter) return clientDbAdapter;
  if (initPromise) return initPromise;

  initPromise = (async () => {
    const SQL = await initSqlJs({
      locateFile: () => sqlWasmUrl,
    });
    const existingBytes = await readIdbSnapshot();
    const raw = existingBytes ? new SQL.Database(existingBytes) : new SQL.Database();
    raw.exec(SCHEMA_SQL);
    clientRawDb = raw;
    clientDbAdapter = createBetterSqliteShim(raw);
    return clientDbAdapter;
  })();

  return initPromise;
}

export async function persistClientDb(): Promise<void> {
  if (!clientRawDb) return;
  const bytes = clientRawDb.export();
  await writeIdbSnapshot(bytes);
  try {
    window.localStorage.setItem('fmcg_bi_client_mode_active', '1');
  } catch {
    // ignore storage quota
  }
}

export function exportClientDbBytes(): Uint8Array | null {
  return clientRawDb ? clientRawDb.export() : null;
}
