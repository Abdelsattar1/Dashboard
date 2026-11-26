import Database from 'better-sqlite3';
import fs from 'fs';
import path from 'path';

let dbInstance: Database.Database | null = null;

export function getDbPath(): string {
  const configured = process.env.SQLITE_DB_PATH?.trim();
  if (configured) {
    return path.resolve(process.cwd(), configured);
  }
  return path.resolve(process.cwd(), 'data', 'sales_bi.db');
}

export function initDatabase(customPath?: string): Database.Database {
  if (dbInstance && !customPath) {
    return dbInstance;
  }

  const dbPath = customPath ?? getDbPath();
  if (dbPath !== ':memory:') {
    const dir = path.dirname(dbPath);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
  }

  const db = new Database(dbPath);
  db.pragma('journal_mode = WAL');
  db.pragma('foreign_keys = ON');

  db.exec(`
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
  `);

  // Idempotent migration for existing SQLite databases missing supervisor_name on departments_chains
  const chainCols = db.prepare(`PRAGMA table_info(departments_chains)`).all() as Array<{ name: string }>;
  if (!chainCols.some((c) => c.name === 'supervisor_name')) {
    db.exec(`ALTER TABLE departments_chains ADD COLUMN supervisor_name TEXT NOT NULL DEFAULT ''`);
  }

  // Idempotent migration for existing SQLite databases missing chain_name / rep_name / supervisor_name on branches
  const branchCols = db.prepare(`PRAGMA table_info(branches)`).all() as Array<{ name: string }>;
  if (!branchCols.some((c) => c.name === 'chain_name')) {
    db.exec(`ALTER TABLE branches ADD COLUMN chain_name TEXT NOT NULL DEFAULT ''`);
  }
  if (!branchCols.some((c) => c.name === 'rep_name')) {
    db.exec(`ALTER TABLE branches ADD COLUMN rep_name TEXT NOT NULL DEFAULT ''`);
  }
  if (!branchCols.some((c) => c.name === 'supervisor_name')) {
    db.exec(`ALTER TABLE branches ADD COLUMN supervisor_name TEXT NOT NULL DEFAULT ''`);
  }

  // Idempotent migration for existing SQLite databases missing monthly_target on sales_reps
  const repCols = db.prepare(`PRAGMA table_info(sales_reps)`).all() as Array<{ name: string }>;
  if (!repCols.some((c) => c.name === 'monthly_target')) {
    db.exec(`ALTER TABLE sales_reps ADD COLUMN monthly_target REAL NOT NULL DEFAULT 0`);
  }

  // Idempotent migration for existing SQLite databases missing sector classification columns on items
  const itemCols = db.prepare(`PRAGMA table_info(items)`).all() as Array<{ name: string }>;
  if (!itemCols.some((c) => c.name === 'chains_category')) {
    db.exec(`ALTER TABLE items ADD COLUMN chains_category TEXT NOT NULL DEFAULT ''`);
  }
  if (!itemCols.some((c) => c.name === 'army_category')) {
    db.exec(`ALTER TABLE items ADD COLUMN army_category TEXT NOT NULL DEFAULT ''`);
  }
  if (!itemCols.some((c) => c.name === 'top_retail_category')) {
    db.exec(`ALTER TABLE items ADD COLUMN top_retail_category TEXT NOT NULL DEFAULT ''`);
  }

  // Idempotent migration for has_category_targets_sheet on import_batches
  const batchCols = db.prepare(`PRAGMA table_info(import_batches)`).all() as Array<{ name: string }>;
  if (!batchCols.some((c) => c.name === 'has_category_targets_sheet')) {
    db.exec(`ALTER TABLE import_batches ADD COLUMN has_category_targets_sheet INTEGER NOT NULL DEFAULT 0`);
  }

  // Idempotent migration: upgrade users table CHECK constraint to ('admin', 'sector_manager', 'standard_user')
  const usersSchemaRow = db
    .prepare(`SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'users'`)
    .get() as { sql?: string } | undefined;
  const usersSchemaSql = usersSchemaRow?.sql ?? '';
  if (
    usersSchemaSql &&
    (!usersSchemaSql.includes("'sector_manager'") || !usersSchemaSql.includes("'standard_user'"))
  ) {
    db.pragma('foreign_keys = OFF');
    try {
      const migrateUsers = db.transaction(() => {
        db.exec(`
          CREATE TABLE users_migrated (
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
          INSERT INTO users_migrated (
            id, username, full_name, email, role, allowed_sectors,
            password_hash, must_change_password, is_active, created_at
          )
          SELECT
            id,
            username,
            full_name,
            email,
            CASE
              WHEN role = 'admin' THEN 'admin'
              WHEN role = 'sector_manager' THEN 'sector_manager'
              ELSE 'standard_user'
            END,
            allowed_sectors,
            password_hash,
            must_change_password,
            is_active,
            created_at
          FROM users;
          DROP TABLE users;
          ALTER TABLE users_migrated RENAME TO users;
        `);
      });
      migrateUsers();
    } finally {
      db.pragma('foreign_keys = ON');
    }
  }

  if (!customPath) {
    dbInstance = db;
  }
  return db;
}

export function getDb(): Database.Database {
  if (!dbInstance) {
    return initDatabase();
  }
  return dbInstance;
}
