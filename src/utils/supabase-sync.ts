import { createClient, type SupabaseClient } from '@supabase/supabase-js';
import type Database from 'better-sqlite3';
import { ORIGINAL_SECTORS } from '../../server/excel-engine.ts';
import type { ImportBatchHistoryItem, UserProfile, UserRole } from '../types/bi.ts';
import {
  CLIENT_DEFAULT_ADMIN_SEED,
  CLIENT_DEFAULT_ADMIN_USERNAME,
  hashPasswordClient,
  normalizeRoleClient,
} from './client-auth.ts';

let cachedSupabase: SupabaseClient | null = null;

export function getSupabaseConfig(): { url: string; anonKey: string } {
  const url = String(import.meta.env.VITE_SUPABASE_URL || '').trim();
  const anonKey = String(import.meta.env.VITE_SUPABASE_ANON_KEY || '').trim();
  return { url, anonKey };
}

export function isSupabaseConfigured(): boolean {
  const { url, anonKey } = getSupabaseConfig();
  return Boolean(url && anonKey && /^https?:\/\//i.test(url));
}

export function getSupabaseClient(): SupabaseClient {
  if (cachedSupabase) return cachedSupabase;
  const { url, anonKey } = getSupabaseConfig();
  if (!url || !anonKey) {
    throw new Error('متغيرات البيئة VITE_SUPABASE_URL و VITE_SUPABASE_ANON_KEY غير مضبوطة.');
  }
  cachedSupabase = createClient(url, anonKey, {
    auth: {
      persistSession: false,
      autoRefreshToken: false,
    },
  });
  return cachedSupabase;
}

function formatSupabaseError(err: { code?: string; message?: string } | null | undefined, fallback: string): Error {
  if (!err) return new Error(fallback);
  if (err.code === '42P01' || err.code === 'PGRST205') {
    return new Error('جداول Supabase غير منشأة بعد. يرجى نسخ كود SQL من تبويب «ربط Supabase و SQL» وتشغيله في Supabase SQL Editor.');
  }
  return new Error(err.message || fallback);
}

export interface SupabaseAppUserRow {
  id: number;
  username: string;
  full_name: string;
  email: string | null;
  role: string;
  allowed_sectors: string[] | string;
  password_hash: string;
  must_change_password: boolean;
  is_active: boolean;
  created_at: string;
}

export function mapSupabaseUser(row: SupabaseAppUserRow): UserProfile {
  let allowedSectors: string[] = [];
  if (Array.isArray(row.allowed_sectors)) {
    allowedSectors = row.allowed_sectors.map((s) => String(s));
  } else if (typeof row.allowed_sectors === 'string') {
    try {
      const parsed = JSON.parse(row.allowed_sectors);
      if (Array.isArray(parsed)) allowedSectors = parsed.map((s) => String(s));
    } catch {
      allowedSectors = [];
    }
  }
  return {
    id: Number(row.id),
    username: row.username,
    fullName: row.full_name,
    email: row.email ?? null,
    role: normalizeRoleClient(row.role),
    allowedSectors,
    isActive: Boolean(row.is_active),
    mustChangePassword: Boolean(row.must_change_password),
    createdAt: row.created_at,
  };
}

export async function supabaseTestConnection(): Promise<{ ok: boolean; userCount: number; message: string }> {
  if (!isSupabaseConfigured()) {
    return {
      ok: false,
      userCount: 0,
      message: 'متغيرات VITE_SUPABASE_URL و VITE_SUPABASE_ANON_KEY غير مضبوطة في بيئة النشر.',
    };
  }
  const sb = getSupabaseClient();
  const { count, error } = await sb.from('app_users').select('id', { count: 'exact', head: true });
  if (error) {
    return {
      ok: false,
      userCount: 0,
      message: formatSupabaseError(error, 'تعذر الاتصال بـ Supabase').message,
    };
  }
  return {
    ok: true,
    userCount: count ?? 0,
    message: `تم الاتصال بنجاح بـ Supabase السحابية (${count ?? 0} مستخدم مسجل).`,
  };
}

export async function supabaseEnsureDefaultAdmin(): Promise<{ seededNow: boolean; defaultAdminPendingChange: boolean }> {
  const sb = getSupabaseClient();
  const { data: admins, error } = await sb
    .from('app_users')
    .select('id, username, must_change_password, is_active')
    .eq('role', 'admin');

  if (error) throw formatSupabaseError(error, 'تعذر التحقق من حساب المدير في Supabase');

  if (!admins || admins.length === 0) {
    const hash = await hashPasswordClient(CLIENT_DEFAULT_ADMIN_SEED);
    const { error: insErr } = await sb.from('app_users').insert({
      username: CLIENT_DEFAULT_ADMIN_USERNAME,
      full_name: 'مدير النظام',
      email: null,
      role: 'admin',
      allowed_sectors: [...ORIGINAL_SECTORS],
      password_hash: hash,
      must_change_password: true,
      is_active: true,
      created_at: new Date().toISOString(),
    });
    if (insErr) throw formatSupabaseError(insErr, 'تعذر إنشاء حساب الأدمن الافتراضي في Supabase');
    return { seededNow: true, defaultAdminPendingChange: true };
  }

  const defaultAdmin = admins.find((a) => String(a.username).toLowerCase() === 'admin' && a.is_active);
  return {
    seededNow: false,
    defaultAdminPendingChange: Boolean(defaultAdmin && defaultAdmin.must_change_password),
  };
}

export async function supabaseFindUserByUsername(username: string): Promise<SupabaseAppUserRow | null> {
  const sb = getSupabaseClient();
  const clean = username.trim().toLowerCase();
  const { data, error } = await sb
    .from('app_users')
    .select('*')
    .ilike('username', clean)
    .limit(1)
    .maybeSingle();
  if (error) throw formatSupabaseError(error, 'تعذر البحث عن المستخدم');
  return (data as SupabaseAppUserRow | null) ?? null;
}

export async function supabaseFindUserById(id: number): Promise<SupabaseAppUserRow | null> {
  const sb = getSupabaseClient();
  const { data, error } = await sb
    .from('app_users')
    .select('*')
    .eq('id', id)
    .limit(1)
    .maybeSingle();
  if (error) throw formatSupabaseError(error, 'تعذر جلب بيانات المستخدم');
  return (data as SupabaseAppUserRow | null) ?? null;
}

export async function supabaseListUsers(): Promise<UserProfile[]> {
  const sb = getSupabaseClient();
  const { data, error } = await sb
    .from('app_users')
    .select('*')
    .order('id', { ascending: true });
  if (error) throw formatSupabaseError(error, 'تعذر تحميل قائمة المستخدمين');
  return ((data || []) as SupabaseAppUserRow[]).map(mapSupabaseUser);
}

export async function supabaseCreateUser(input: {
  username: string;
  fullName: string;
  email: string | null;
  role: UserRole;
  allowedSectors: string[];
  passwordHash: string;
  mustChangePassword: boolean;
}): Promise<UserProfile> {
  const sb = getSupabaseClient();
  const { data, error } = await sb
    .from('app_users')
    .insert({
      username: input.username,
      full_name: input.fullName,
      email: input.email,
      role: input.role,
      allowed_sectors: input.allowedSectors,
      password_hash: input.passwordHash,
      must_change_password: input.mustChangePassword,
      is_active: true,
      created_at: new Date().toISOString(),
    })
    .select('*')
    .single();
  if (error) {
    if (error.code === '23505') {
      throw new Error('اسم المستخدم مسجل بالفعل');
    }
    throw formatSupabaseError(error, 'تعذر إنشاء المستخدم في Supabase');
  }
  return mapSupabaseUser(data as SupabaseAppUserRow);
}

export async function supabaseUpdateUser(
  id: number,
  patch: {
    full_name?: string;
    email?: string | null;
    role?: UserRole;
    allowed_sectors?: string[];
    is_active?: boolean;
    password_hash?: string;
    must_change_password?: boolean;
  }
): Promise<UserProfile> {
  const sb = getSupabaseClient();
  const { data, error } = await sb
    .from('app_users')
    .update(patch)
    .eq('id', id)
    .select('*')
    .single();
  if (error) throw formatSupabaseError(error, 'تعذر تحديث المستخدم في Supabase');
  return mapSupabaseUser(data as SupabaseAppUserRow);
}

export async function supabaseDeleteUser(id: number): Promise<void> {
  const sb = getSupabaseClient();
  const { error } = await sb.from('app_users').delete().eq('id', id);
  if (error) throw formatSupabaseError(error, 'تعذر حذف المستخدم من Supabase');
}

async function fetchAllPaged<T>(table: string, batchId?: number): Promise<T[]> {
  const sb = getSupabaseClient();
  const pageSize = 1000;
  let from = 0;
  const out: T[] = [];
  while (true) {
    let query = sb.from(table).select('*').order('id', { ascending: true }).range(from, from + pageSize - 1);
    if (typeof batchId === 'number') {
      query = query.eq('batch_id', batchId);
    }
    const { data, error } = await query;
    if (error) throw formatSupabaseError(error, `تعذر جلب بيانات ${table}`);
    const chunk = (data || []) as T[];
    out.push(...chunk);
    if (chunk.length < pageSize) break;
    from += pageSize;
  }
  return out;
}

async function insertInChunks(table: string, rows: Record<string, unknown>[], chunkSize = 400): Promise<void> {
  if (rows.length === 0) return;
  const sb = getSupabaseClient();
  for (let i = 0; i < rows.length; i += chunkSize) {
    const slice = rows.slice(i, i + chunkSize);
    const { error } = await sb.from(table).insert(slice);
    if (error) throw formatSupabaseError(error, `تعذر حفظ بيانات ${table} في Supabase`);
  }
}

async function upsertInChunks(
  table: string,
  rows: Record<string, unknown>[],
  onConflict: string,
  chunkSize = 400
): Promise<void> {
  if (rows.length === 0) return;
  const sb = getSupabaseClient();
  for (let i = 0; i < rows.length; i += chunkSize) {
    const slice = rows.slice(i, i + chunkSize);
    const { error } = await sb.from(table).upsert(slice, { onConflict });
    if (error) throw formatSupabaseError(error, `تعذر تحديث بيانات ${table} في Supabase`);
  }
}

function dedupeByKey<T>(rows: T[], keyFn: (row: T) => string): T[] {
  const map = new Map<string, T>();
  for (const row of rows) {
    const k = keyFn(row);
    if (!k) continue;
    map.set(k, row);
  }
  return Array.from(map.values());
}

function extractMonthKey(billingDate: string | null | undefined): string {
  const d = String(billingDate || '').trim();
  const m = d.match(/^(\d{4})-(\d{2})(?:-\d{2})?/);
  if (m) {
    const y = Number(m[1]);
    const mo = Number(m[2]);
    if (y >= 1990 && y <= 2099 && mo >= 1 && mo <= 12) {
      return `${m[1]}-${m[2]}`;
    }
  }
  return 'UNSPECIFIED';
}

export async function syncSupabaseToLocalDb(db: Database.Database, opts?: { forceBatch?: boolean }): Promise<void> {
  const sb = getSupabaseClient();

  const [usersRes, catTargetsRes, compsRes, compItemsRes, compTargetsRes, errorsRes, backupsRes, activeBatchRes] =
    await Promise.all([
      sb.from('app_users').select('*').order('id', { ascending: true }),
      fetchAllPaged<{
        id: number;
        target_key: string;
        sector_name: string;
        job_title: string;
        employee_name: string;
        category_name: string;
        target_value: number;
        updated_at: string;
      }>('item_category_targets'),
      sb.from('competitions').select('*').order('id', { ascending: true }),
      fetchAllPaged<{ id: number; competition_id: number; item_code: string; item_name: string; equivalent_factor: number }>(
        'competition_items'
      ),
      fetchAllPaged<{
        id: number;
        competition_id: number;
        participant_key: string;
        target_qty: number;
        industrial_target_qty: number;
        commercial_target_qty: number;
      }>('competition_targets'),
      sb.from('runtime_error_logs').select('*').eq('resolved', false).order('id', { ascending: false }).limit(100),
      sb.from('audit_fix_backups').select('*').order('id', { ascending: false }).limit(30),
      sb.from('import_batches').select('*').eq('status', 'active').order('id', { ascending: false }).limit(1).maybeSingle(),
    ]);

  if (usersRes.error) throw formatSupabaseError(usersRes.error, 'تعذر مزامنة المستخدمين');
  if (compsRes.error) throw formatSupabaseError(compsRes.error, 'تعذر مزامنة المسابقات');
  if (errorsRes.error) throw formatSupabaseError(errorsRes.error, 'تعذر مزامنة سجل الأخطاء');
  if (backupsRes.error) throw formatSupabaseError(backupsRes.error, 'تعذر مزامنة النسخ الاحتياطية');
  if (activeBatchRes.error) throw formatSupabaseError(activeBatchRes.error, 'تعذر قراءة الدفعة النشطة من Supabase');

  const remoteUsers = (usersRes.data || []) as SupabaseAppUserRow[];
  const remoteBatch = activeBatchRes.data as
    | {
        id: number;
        batch_uuid: string;
        filename: string;
        imported_by_id: number | null;
        imported_by_name: string;
        status: 'active' | 'archived';
        vf05n_count: number;
        vf05n_sales_total: number;
        vf05n_tax_total: number;
        zsales_count: number;
        zsales_sales_total: number;
        zsales_tax_total: number;
        unlinked_branch_count: number;
        unlinked_sector_count: number;
        has_category_targets_sheet: boolean;
        validation_report: unknown;
        created_at: string;
      }
    | null;

  const localActiveBatch = db
    .prepare(`SELECT id, batch_uuid FROM import_batches WHERE status = 'active' ORDER BY id DESC LIMIT 1`)
    .get() as { id: number; batch_uuid: string } | undefined;

  const shouldHydrateBatch = Boolean(
    remoteBatch && (opts?.forceBatch || !localActiveBatch || localActiveBatch.batch_uuid !== remoteBatch.batch_uuid)
  );
  const shouldClearBatch = Boolean(!remoteBatch && localActiveBatch);

  let batchSectors: Array<{ id: number; batch_id: number; sector_code: string; sector_name: string; monthly_target: number; is_original_sector: boolean }> = [];
  let batchChains: Array<{ id: number; batch_id: number; chain_code: string; chain_name: string; department_name: string; sector_name: string; supervisor_name: string }> = [];
  let batchBranches: Array<{
    id: number;
    batch_id: number;
    branch_code: string;
    branch_name: string;
    sector_name: string;
    branch_target: number;
    chain_code: string | null;
    chain_name: string;
    rep_code: string | null;
    rep_name: string;
    supervisor_name: string;
  }> = [];
  let batchReps: Array<{ id: number; batch_id: number; rep_code: string; rep_name: string; sector_name: string | null; supervisor_name: string | null; monthly_target: number }> = [];
  let batchItems: Array<{
    id: number;
    batch_id: number;
    item_code: string;
    item_name: string;
    category_name: string | null;
    brand_name: string | null;
    unit_name: string | null;
    chains_category: string;
    army_category: string;
    top_retail_category: string;
  }> = [];
  let batchCatTargets: Array<{
    id: number;
    batch_id: number;
    sector_name: string;
    job_title: string;
    employee_code: string;
    employee_name: string;
    category_name: string;
    target_value: number;
  }> = [];
  let batchSales: Array<{
    id: number;
    batch_id: number;
    source_type: string;
    doc_number: string | null;
    billing_date: string | null;
    payer_code: string | null;
    payer_name: string | null;
    sold_to_code: string | null;
    sold_to_name: string | null;
    item_code: string | null;
    item_name: string | null;
    rep_code: string | null;
    quantity: number;
    sales_value: number;
    tax_value: number;
    resolved_sector: string | null;
    resolved_department: string | null;
    resolved_chain_name: string | null;
    resolved_branch_code: string | null;
    resolved_branch_name: string | null;
    resolved_rep_code: string | null;
    resolved_rep_name: string | null;
    resolved_item_name: string | null;
    is_return: boolean;
  }> = [];

  if (shouldHydrateBatch && remoteBatch) {
    const [rawSectors, rawChains, rawBranches, rawReps, rawItems, rawCatTargets, rawSales] = await Promise.all([
      fetchAllPaged<(typeof batchSectors)[number]>('sectors'),
      fetchAllPaged<(typeof batchChains)[number]>('departments_chains'),
      fetchAllPaged<(typeof batchBranches)[number]>('branches'),
      fetchAllPaged<(typeof batchReps)[number]>('sales_reps'),
      fetchAllPaged<(typeof batchItems)[number]>('items'),
      fetchAllPaged<(typeof batchCatTargets)[number]>('category_targets'),
      fetchAllPaged<(typeof batchSales)[number]>('sales_records'),
    ]);

    const byBatchAsc = <T extends { batch_id: number; id: number }>(arr: T[]) =>
      [...arr].sort((a, b) => a.batch_id - b.batch_id || a.id - b.id);

    batchSectors = dedupeByKey(byBatchAsc(rawSectors), (s) => s.sector_name);
    batchChains = dedupeByKey(byBatchAsc(rawChains), (c) => c.chain_code);
    batchBranches = dedupeByKey(byBatchAsc(rawBranches), (b) => b.branch_code);
    batchReps = dedupeByKey(byBatchAsc(rawReps), (r) => r.rep_code);
    batchItems = dedupeByKey(byBatchAsc(rawItems), (it) => it.item_code);
    batchCatTargets = dedupeByKey(
      byBatchAsc(rawCatTargets),
      (ct) => `${(ct.sector_name || '').trim()}\u0000${(ct.job_title || '').trim()}\u0000${(ct.employee_name || '').trim()}\u0000${(ct.category_name || '').trim()}`
    );

    // Accumulate distinct months across batches, preferring the latest batch_id per month (or active remoteBatch.id)
    const monthOwnerBatch = new Map<string, number>();
    for (const sr of rawSales) {
      const m = extractMonthKey(sr.billing_date);
      const prevOwner = monthOwnerBatch.get(m);
      if (prevOwner === remoteBatch.id) continue;
      if (sr.batch_id === remoteBatch.id || prevOwner === undefined || sr.batch_id > prevOwner) {
        monthOwnerBatch.set(m, sr.batch_id);
      }
    }
    batchSales = rawSales.filter((sr) => sr.batch_id === monthOwnerBatch.get(extractMonthKey(sr.billing_date)));
  }

  db.pragma('foreign_keys = OFF');
  try {
    const tx = db.transaction(() => {
      // 1. Sync users
      db.prepare(`DELETE FROM users`).run();
      const insUser = db.prepare(
        `INSERT INTO users (id, username, full_name, email, role, allowed_sectors, password_hash, must_change_password, is_active, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
      );
      for (const u of remoteUsers) {
        const sectorsJson = Array.isArray(u.allowed_sectors) ? JSON.stringify(u.allowed_sectors) : String(u.allowed_sectors || '[]');
        insUser.run(
          u.id,
          u.username,
          u.full_name,
          u.email ?? null,
          u.role,
          sectorsJson,
          u.password_hash,
          u.must_change_password ? 1 : 0,
          u.is_active ? 1 : 0,
          u.created_at
        );
      }

      // 2. Sync item_category_targets
      db.prepare(`DELETE FROM item_category_targets`).run();
      const insCatT = db.prepare(
        `INSERT INTO item_category_targets (id, target_key, sector_name, job_title, employee_name, category_name, target_value, updated_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
      );
      for (const ct of catTargetsRes) {
        insCatT.run(ct.id, ct.target_key, ct.sector_name, ct.job_title, ct.employee_name, ct.category_name, ct.target_value, ct.updated_at);
      }

      // 3. Sync competitions + items + targets
      db.prepare(`DELETE FROM competition_targets`).run();
      db.prepare(`DELETE FROM competition_items`).run();
      db.prepare(`DELETE FROM competitions`).run();
      const insComp = db.prepare(
        `INSERT INTO competitions (id, name, is_active, created_by_id, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?)`
      );
      for (const c of (compsRes.data || []) as Array<{
        id: number;
        name: string;
        is_active: boolean;
        created_by_id: number | null;
        created_at: string;
        updated_at: string;
      }>) {
        insComp.run(c.id, c.name, c.is_active ? 1 : 0, c.created_by_id ?? null, c.created_at, c.updated_at);
      }
      const insCompItem = db.prepare(
        `INSERT INTO competition_items (id, competition_id, item_code, item_name, equivalent_factor) VALUES (?, ?, ?, ?, ?)`
      );
      for (const ci of compItemsRes) {
        insCompItem.run(ci.id, ci.competition_id, ci.item_code, ci.item_name, ci.equivalent_factor);
      }
      const insCompTarget = db.prepare(
        `INSERT INTO competition_targets (id, competition_id, participant_key, target_qty, industrial_target_qty, commercial_target_qty)
         VALUES (?, ?, ?, ?, ?, ?)`
      );
      for (const ct of compTargetsRes) {
        insCompTarget.run(ct.id, ct.competition_id, ct.participant_key, ct.target_qty, ct.industrial_target_qty, ct.commercial_target_qty);
      }

      // 4. Sync runtime errors & backups
      db.prepare(`DELETE FROM runtime_error_logs`).run();
      const insErr = db.prepare(
        `INSERT INTO runtime_error_logs (id, kind, tab_name, message, stack, source_url, line_no, col_no, user_id, username, resolved, ai_analysis, occurred_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
      );
      for (const e of (errorsRes.data || []) as Array<{
        id: number;
        kind: string;
        tab_name: string;
        message: string;
        stack: string;
        source_url: string;
        line_no: number;
        col_no: number;
        user_id: number | null;
        username: string;
        resolved: boolean;
        ai_analysis: string;
        occurred_at: string;
      }>) {
        insErr.run(
          e.id,
          e.kind,
          e.tab_name,
          e.message,
          e.stack,
          e.source_url,
          e.line_no,
          e.col_no,
          e.user_id ?? null,
          e.username,
          e.resolved ? 1 : 0,
          e.ai_analysis,
          e.occurred_at
        );
      }

      db.prepare(`DELETE FROM audit_fix_backups`).run();
      const insBackup = db.prepare(
        `INSERT INTO audit_fix_backups (id, finding_code, action_type, action_label, summary, snapshot_json, created_by_id, created_by_name, created_at, restored_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
      );
      for (const b of (backupsRes.data || []) as Array<{
        id: number;
        finding_code: string;
        action_type: string;
        action_label: string;
        summary: string;
        snapshot_json: unknown;
        created_by_id: number | null;
        created_by_name: string;
        created_at: string;
        restored_at: string | null;
      }>) {
        insBackup.run(
          b.id,
          b.finding_code,
          b.action_type,
          b.action_label,
          b.summary,
          typeof b.snapshot_json === 'string' ? b.snapshot_json : JSON.stringify(b.snapshot_json || {}),
          b.created_by_id ?? null,
          b.created_by_name,
          b.created_at,
          b.restored_at ?? null
        );
      }

      // 5. Hydrate active import batch if needed
      if (shouldClearBatch || shouldHydrateBatch) {
        db.prepare(`DELETE FROM sales_records`).run();
        db.prepare(`DELETE FROM category_targets`).run();
        db.prepare(`DELETE FROM items`).run();
        db.prepare(`DELETE FROM sales_reps`).run();
        db.prepare(`DELETE FROM branches`).run();
        db.prepare(`DELETE FROM departments_chains`).run();
        db.prepare(`DELETE FROM sectors`).run();
        db.prepare(`DELETE FROM import_batches`).run();
      }

      if (shouldHydrateBatch && remoteBatch) {
        db.prepare(
          `INSERT INTO import_batches (
            id, batch_uuid, filename, imported_by_id, imported_by_name, status,
            vf05n_count, vf05n_sales_total, vf05n_tax_total,
            zsales_count, zsales_sales_total, zsales_tax_total,
            unlinked_branch_count, unlinked_sector_count, has_category_targets_sheet, validation_report, created_at
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).run(
          remoteBatch.id,
          remoteBatch.batch_uuid,
          remoteBatch.filename,
          remoteBatch.imported_by_id ?? null,
          remoteBatch.imported_by_name,
          remoteBatch.status,
          remoteBatch.vf05n_count,
          remoteBatch.vf05n_sales_total,
          remoteBatch.vf05n_tax_total,
          remoteBatch.zsales_count,
          remoteBatch.zsales_sales_total,
          remoteBatch.zsales_tax_total,
          remoteBatch.unlinked_branch_count,
          remoteBatch.unlinked_sector_count,
          remoteBatch.has_category_targets_sheet ? 1 : 0,
          typeof remoteBatch.validation_report === 'string'
            ? remoteBatch.validation_report
            : JSON.stringify(remoteBatch.validation_report || {}),
          remoteBatch.created_at
        );

        const insSec = db.prepare(
          `INSERT INTO sectors (id, batch_id, sector_code, sector_name, monthly_target, is_original_sector) VALUES (?, ?, ?, ?, ?, ?)`
        );
        for (const s of batchSectors) {
          insSec.run(s.id, remoteBatch.id, s.sector_code, s.sector_name, s.monthly_target, s.is_original_sector ? 1 : 0);
        }

        const insCh = db.prepare(
          `INSERT INTO departments_chains (id, batch_id, chain_code, chain_name, department_name, sector_name, supervisor_name) VALUES (?, ?, ?, ?, ?, ?, ?)`
        );
        for (const c of batchChains) {
          insCh.run(c.id, remoteBatch.id, c.chain_code, c.chain_name, c.department_name, c.sector_name, c.supervisor_name);
        }

        const insBr = db.prepare(
          `INSERT INTO branches (id, batch_id, branch_code, branch_name, sector_name, branch_target, chain_code, chain_name, rep_code, rep_name, supervisor_name) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        );
        for (const b of batchBranches) {
          insBr.run(
            b.id,
            remoteBatch.id,
            b.branch_code,
            b.branch_name,
            b.sector_name,
            b.branch_target,
            b.chain_code,
            b.chain_name,
            b.rep_code,
            b.rep_name,
            b.supervisor_name
          );
        }

        const insRep = db.prepare(
          `INSERT INTO sales_reps (id, batch_id, rep_code, rep_name, sector_name, supervisor_name, monthly_target) VALUES (?, ?, ?, ?, ?, ?, ?)`
        );
        for (const r of batchReps) {
          insRep.run(r.id, remoteBatch.id, r.rep_code, r.rep_name, r.sector_name, r.supervisor_name, r.monthly_target);
        }

        const insIt = db.prepare(
          `INSERT INTO items (id, batch_id, item_code, item_name, category_name, brand_name, unit_name, chains_category, army_category, top_retail_category) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        );
        for (const it of batchItems) {
          insIt.run(
            it.id,
            remoteBatch.id,
            it.item_code,
            it.item_name,
            it.category_name,
            it.brand_name,
            it.unit_name,
            it.chains_category,
            it.army_category,
            it.top_retail_category
          );
        }

        const insCt = db.prepare(
          `INSERT INTO category_targets (id, batch_id, sector_name, job_title, employee_code, employee_name, category_name, target_value) VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
        );
        for (const ct of batchCatTargets) {
          insCt.run(ct.id, remoteBatch.id, ct.sector_name, ct.job_title, ct.employee_code, ct.employee_name, ct.category_name, ct.target_value);
        }

        const insSale = db.prepare(
          `INSERT INTO sales_records (
            id, batch_id, source_type, doc_number, billing_date,
            payer_code, payer_name, sold_to_code, sold_to_name,
            item_code, item_name, rep_code, quantity, sales_value, tax_value,
            resolved_sector, resolved_department, resolved_chain_name,
            resolved_branch_code, resolved_branch_name, resolved_rep_code, resolved_rep_name, resolved_item_name,
            is_return
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        );
        for (const sr of batchSales) {
          insSale.run(
            sr.id,
            remoteBatch.id,
            sr.source_type,
            sr.doc_number,
            sr.billing_date,
            sr.payer_code,
            sr.payer_name,
            sr.sold_to_code,
            sr.sold_to_name,
            sr.item_code,
            sr.item_name,
            sr.rep_code,
            sr.quantity,
            sr.sales_value,
            sr.tax_value,
            sr.resolved_sector,
            sr.resolved_department,
            sr.resolved_chain_name,
            sr.resolved_branch_code,
            sr.resolved_branch_name,
            sr.resolved_rep_code,
            sr.resolved_rep_name,
            sr.resolved_item_name,
            sr.is_return ? 1 : 0
          );
        }
      }
    });
    tx();
  } finally {
    db.pragma('foreign_keys = ON');
  }
}

export async function supabasePushCommittedBatch(db: Database.Database, localBatchId: number): Promise<ImportBatchHistoryItem> {
  const sb = getSupabaseClient();
  const batchRow = db.prepare(`SELECT * FROM import_batches WHERE id = ?`).get(localBatchId) as {
    id: number;
    batch_uuid: string;
    filename: string;
    imported_by_id: number | null;
    imported_by_name: string;
    vf05n_count: number;
    vf05n_sales_total: number;
    vf05n_tax_total: number;
    zsales_count: number;
    zsales_sales_total: number;
    zsales_tax_total: number;
    unlinked_branch_count: number;
    unlinked_sector_count: number;
    has_category_targets_sheet: number;
    validation_report: string;
    created_at: string;
  };

  const sectors = db.prepare(`SELECT sector_code, sector_name, monthly_target, is_original_sector FROM sectors WHERE batch_id = ?`).all(localBatchId) as Array<{
    sector_code: string;
    sector_name: string;
    monthly_target: number;
    is_original_sector: number;
  }>;
  const chains = db
    .prepare(`SELECT chain_code, chain_name, department_name, sector_name, supervisor_name FROM departments_chains WHERE batch_id = ?`)
    .all(localBatchId) as Array<Record<string, unknown>>;
  const branches = db
    .prepare(`SELECT branch_code, branch_name, sector_name, branch_target, chain_code, chain_name, rep_code, rep_name, supervisor_name FROM branches WHERE batch_id = ?`)
    .all(localBatchId) as Array<Record<string, unknown>>;
  const reps = db
    .prepare(`SELECT rep_code, rep_name, sector_name, supervisor_name, monthly_target FROM sales_reps WHERE batch_id = ?`)
    .all(localBatchId) as Array<Record<string, unknown>>;
  const items = db
    .prepare(`SELECT item_code, item_name, category_name, brand_name, unit_name, chains_category, army_category, top_retail_category FROM items WHERE batch_id = ?`)
    .all(localBatchId) as Array<Record<string, unknown>>;
  const catTargets = db
    .prepare(`SELECT sector_name, job_title, employee_code, employee_name, category_name, target_value FROM category_targets WHERE batch_id = ?`)
    .all(localBatchId) as Array<Record<string, unknown>>;
  const sales = db
    .prepare(
      `SELECT source_type, doc_number, billing_date, payer_code, payer_name, sold_to_code, sold_to_name, item_code, item_name, rep_code, quantity, sales_value, tax_value, resolved_sector, resolved_department, resolved_chain_name, resolved_branch_code, resolved_branch_name, resolved_rep_code, resolved_rep_name, resolved_item_name, is_return
       FROM sales_records WHERE batch_id = ?`
    )
    .all(localBatchId) as Array<Record<string, unknown>>;

  // Archive existing active batches in Supabase
  const { error: archErr } = await sb.from('import_batches').update({ status: 'archived' }).eq('status', 'active');
  if (archErr) throw formatSupabaseError(archErr, 'تعذر أرشفة النسخ السابقة في Supabase');

  let parsedValidation: Record<string, unknown> = {};
  try {
    parsedValidation = JSON.parse(batchRow.validation_report || '{}') as Record<string, unknown>;
  } catch {
    parsedValidation = {};
  }

  const { data: insertedBatch, error: batchErr } = await sb
    .from('import_batches')
    .insert({
      batch_uuid: batchRow.batch_uuid,
      filename: batchRow.filename,
      imported_by_id: batchRow.imported_by_id,
      imported_by_name: batchRow.imported_by_name,
      status: 'active',
      vf05n_count: batchRow.vf05n_count,
      vf05n_sales_total: batchRow.vf05n_sales_total,
      vf05n_tax_total: batchRow.vf05n_tax_total,
      zsales_count: batchRow.zsales_count,
      zsales_sales_total: batchRow.zsales_sales_total,
      zsales_tax_total: batchRow.zsales_tax_total,
      unlinked_branch_count: batchRow.unlinked_branch_count,
      unlinked_sector_count: batchRow.unlinked_sector_count,
      has_category_targets_sheet: batchRow.has_category_targets_sheet === 1,
      validation_report: parsedValidation,
      created_at: batchRow.created_at,
    })
    .select('*')
    .single();

  if (batchErr) throw formatSupabaseError(batchErr, 'تعذر حفظ سجل الاستيراد في Supabase');
  const remoteBatchId = Number((insertedBatch as { id: number }).id);

  // 1. Upsert Master Reference Data without duplicating existing entities
  const upsertSectors = dedupeByKey(
    sectors.map((s) => ({
      batch_id: remoteBatchId,
      sector_code: s.sector_code,
      sector_name: s.sector_name,
      monthly_target: s.monthly_target,
      is_original_sector: s.is_original_sector === 1,
    })),
    (s) => String(s.sector_name)
  );
  await upsertInChunks('sectors', upsertSectors, 'sector_name');

  const upsertChains = dedupeByKey(
    chains.map((c): Record<string, unknown> => ({ ...c, batch_id: remoteBatchId })),
    (c) => String(c.chain_code || '')
  );
  await upsertInChunks('departments_chains', upsertChains, 'chain_code');

  const upsertBranches = dedupeByKey(
    branches.map((b): Record<string, unknown> => ({ ...b, batch_id: remoteBatchId })),
    (b) => String(b.branch_code || '')
  );
  await upsertInChunks('branches', upsertBranches, 'branch_code');

  const upsertReps = dedupeByKey(
    reps.map((r): Record<string, unknown> => ({ ...r, batch_id: remoteBatchId })),
    (r) => String(r.rep_code || '')
  );
  await upsertInChunks('sales_reps', upsertReps, 'rep_code');

  const upsertItems = dedupeByKey(
    items.map((it): Record<string, unknown> => ({ ...it, batch_id: remoteBatchId })),
    (it) => String(it.item_code || '')
  );
  await upsertInChunks('items', upsertItems, 'item_code');

  const upsertCatTargets = dedupeByKey(
    catTargets.map((ct): Record<string, unknown> => ({
      ...ct,
      sector_name: String(ct.sector_name || '').trim(),
      job_title: String(ct.job_title || '').trim(),
      employee_name: String(ct.employee_name || '').trim(),
      category_name: String(ct.category_name || '').trim(),
      batch_id: remoteBatchId,
    })),
    (ct) => `${String(ct.sector_name)}\u0000${String(ct.job_title)}\u0000${String(ct.employee_name)}\u0000${String(ct.category_name)}`
  );
  await upsertInChunks('category_targets', upsertCatTargets, 'sector_name,job_title,employee_name,category_name');

  // 2. Partitioned Monthly Sales Sync using billing_date (YYYY-MM):
  //    Delete ONLY overlapping months from Supabase, leaving all other months untouched,
  //    then insert only the newly uploaded month records.
  const incSummary = parsedValidation.incrementalSummary as
    | { replacedPeriods?: unknown; addedPeriods?: unknown }
    | undefined;
  let incomingPeriods: string[] = [];
  if (incSummary) {
    const repP = Array.isArray(incSummary.replacedPeriods) ? incSummary.replacedPeriods : [];
    const addP = Array.isArray(incSummary.addedPeriods) ? incSummary.addedPeriods : [];
    incomingPeriods = Array.from(new Set([...repP, ...addP].map((p) => String(p || '').trim()).filter(Boolean)));
  }
  if (incomingPeriods.length === 0) {
    incomingPeriods = Array.from(
      new Set(sales.map((sr) => extractMonthKey(sr.billing_date as string | null | undefined)))
    ).filter(Boolean);
  }

  const validMonthPrefixes = incomingPeriods.filter((p) => /^\d{4}-\d{2}$/.test(p));
  for (const period of validMonthPrefixes) {
    const { error: delErr } = await sb.from('sales_records').delete().like('billing_date', `${period}%`);
    if (delErr) throw formatSupabaseError(delErr, `تعذر تحديث مبيعات شهر ${period} في Supabase`);
  }
  if (incomingPeriods.includes('UNSPECIFIED')) {
    const { error: delErr } = await sb.from('sales_records').delete().or('billing_date.is.null,billing_date.eq.');
    if (delErr) throw formatSupabaseError(delErr, 'تعذر تحديث السجلات غير المؤرخة في Supabase');
  }

  const incomingSet = new Set(incomingPeriods);
  const salesToPush = sales.filter((sr) => incomingSet.has(extractMonthKey(sr.billing_date as string | null | undefined)));

  await insertInChunks(
    'sales_records',
    salesToPush.map((sr) => ({
      ...sr,
      batch_id: remoteBatchId,
      is_return: Boolean(sr.is_return),
    }))
  );

  await syncSupabaseToLocalDb(db, { forceBatch: true });

  return {
    id: remoteBatchId,
    batchUuid: batchRow.batch_uuid,
    filename: batchRow.filename,
    importedByName: batchRow.imported_by_name,
    status: 'active',
    vf05nCount: batchRow.vf05n_count,
    vf05nSalesTotal: batchRow.vf05n_sales_total,
    vf05nTaxTotal: batchRow.vf05n_tax_total,
    zsalesCount: batchRow.zsales_count,
    zsalesSalesTotal: batchRow.zsales_sales_total,
    zsalesTaxTotal: batchRow.zsales_tax_total,
    unlinkedBranchCount: batchRow.unlinked_branch_count,
    unlinkedSectorCount: batchRow.unlinked_sector_count,
    createdAt: batchRow.created_at,
  };
}

export async function supabaseListImportBatches(): Promise<ImportBatchHistoryItem[]> {
  const sb = getSupabaseClient();
  const { data, error } = await sb.from('import_batches').select('*').order('id', { ascending: false });
  if (error) throw formatSupabaseError(error, 'تعذر تحميل سجل الاستيراد');
  return ((data || []) as Array<{
    id: number;
    batch_uuid: string;
    filename: string;
    imported_by_name: string;
    status: 'active' | 'archived';
    vf05n_count: number;
    vf05n_sales_total: number;
    vf05n_tax_total: number;
    zsales_count: number;
    zsales_sales_total: number;
    zsales_tax_total: number;
    unlinked_branch_count: number;
    unlinked_sector_count: number;
    created_at: string;
  }>).map((r) => ({
    id: Number(r.id),
    batchUuid: r.batch_uuid,
    filename: r.filename,
    importedByName: r.imported_by_name,
    status: r.status,
    vf05nCount: r.vf05n_count,
    vf05nSalesTotal: r.vf05n_sales_total,
    vf05nTaxTotal: r.vf05n_tax_total,
    zsalesCount: r.zsales_count,
    zsalesSalesTotal: r.zsales_sales_total,
    zsalesTaxTotal: r.zsales_tax_total,
    unlinkedBranchCount: r.unlinked_branch_count,
    unlinkedSectorCount: r.unlinked_sector_count,
    createdAt: r.created_at,
  }));
}

export async function supabaseRollbackToBatch(db: Database.Database, batchId: number): Promise<ImportBatchHistoryItem> {
  const sb = getSupabaseClient();
  const { data: target, error: findErr } = await sb.from('import_batches').select('*').eq('id', batchId).maybeSingle();
  if (findErr) throw formatSupabaseError(findErr, 'تعذر العثور على النسخة');
  if (!target) throw new Error('النسخة المحددة غير موجودة');

  const { error: archErr } = await sb.from('import_batches').update({ status: 'archived' }).eq('status', 'active');
  if (archErr) throw formatSupabaseError(archErr, 'تعذر تحديث حالة النسخ');

  const { error: actErr } = await sb.from('import_batches').update({ status: 'active' }).eq('id', batchId);
  if (actErr) throw formatSupabaseError(actErr, 'تعذر تفعيل النسخة المحددة');

  await syncSupabaseToLocalDb(db, { forceBatch: true });

  const r = target as {
    id: number;
    batch_uuid: string;
    filename: string;
    imported_by_name: string;
    vf05n_count: number;
    vf05n_sales_total: number;
    vf05n_tax_total: number;
    zsales_count: number;
    zsales_sales_total: number;
    zsales_tax_total: number;
    unlinked_branch_count: number;
    unlinked_sector_count: number;
    created_at: string;
  };
  return {
    id: Number(r.id),
    batchUuid: r.batch_uuid,
    filename: r.filename,
    importedByName: r.imported_by_name,
    status: 'active',
    vf05nCount: r.vf05n_count,
    vf05nSalesTotal: r.vf05n_sales_total,
    vf05nTaxTotal: r.vf05n_tax_total,
    zsalesCount: r.zsales_count,
    zsalesSalesTotal: r.zsales_sales_total,
    zsalesTaxTotal: r.zsales_tax_total,
    unlinkedBranchCount: r.unlinked_branch_count,
    unlinkedSectorCount: r.unlinked_sector_count,
    createdAt: r.created_at,
  };
}

export async function supabaseSyncItemCategoryTargetsFromDb(db: Database.Database): Promise<void> {
  const sb = getSupabaseClient();
  const rows = db
    .prepare(`SELECT target_key, sector_name, job_title, employee_name, category_name, target_value, updated_at FROM item_category_targets`)
    .all() as Array<{
    target_key: string;
    sector_name: string;
    job_title: string;
    employee_name: string;
    category_name: string;
    target_value: number;
    updated_at: string;
  }>;
  if (rows.length === 0) return;
  for (let i = 0; i < rows.length; i += 300) {
    const slice = rows.slice(i, i + 300);
    const { error } = await sb.from('item_category_targets').upsert(slice, { onConflict: 'target_key' });
    if (error) throw formatSupabaseError(error, 'تعذر مزامنة أهداف الفئات مع Supabase');
  }
}

export async function supabaseCreateCompetition(
  db: Database.Database,
  createdById: number,
  name: string,
  items: Array<{ itemCode: string; itemName: string; equivalentFactor: number }>
): Promise<number> {
  const sb = getSupabaseClient();
  const now = new Date().toISOString();
  const { data: comp, error } = await sb
    .from('competitions')
    .insert({ name, is_active: true, created_by_id: createdById, created_at: now, updated_at: now })
    .select('id')
    .single();
  if (error) throw formatSupabaseError(error, 'تعذر إنشاء المسابقة في Supabase');
  const compId = Number((comp as { id: number }).id);
  if (items.length > 0) {
    const { error: itemErr } = await sb.from('competition_items').insert(
      items.map((it) => ({
        competition_id: compId,
        item_code: it.itemCode,
        item_name: it.itemName,
        equivalent_factor: it.equivalentFactor,
      }))
    );
    if (itemErr) throw formatSupabaseError(itemErr, 'تعذر حفظ أصناف المسابقة في Supabase');
  }
  await syncSupabaseToLocalDb(db);
  return compId;
}

export async function supabaseUpdateCompetition(
  db: Database.Database,
  compId: number,
  patch: { name?: string; isActive?: boolean; items?: Array<{ itemCode: string; itemName: string; equivalentFactor: number }> }
): Promise<void> {
  const sb = getSupabaseClient();
  const updateObj: Record<string, unknown> = { updated_at: new Date().toISOString() };
  if (patch.name !== undefined) updateObj.name = patch.name;
  if (patch.isActive !== undefined) updateObj.is_active = patch.isActive;
  const { error } = await sb.from('competitions').update(updateObj).eq('id', compId);
  if (error) throw formatSupabaseError(error, 'تعذر تحديث المسابقة في Supabase');

  if (patch.items !== undefined) {
    const { error: delErr } = await sb.from('competition_items').delete().eq('competition_id', compId);
    if (delErr) throw formatSupabaseError(delErr, 'تعذر تحديث أصناف المسابقة');
    if (patch.items.length > 0) {
      const { error: insErr } = await sb.from('competition_items').insert(
        patch.items.map((it) => ({
          competition_id: compId,
          item_code: it.itemCode,
          item_name: it.itemName,
          equivalent_factor: it.equivalentFactor,
        }))
      );
      if (insErr) throw formatSupabaseError(insErr, 'تعذر حفظ أصناف المسابقة الجديدة');
    }
  }
  await syncSupabaseToLocalDb(db);
}

export async function supabaseDeleteCompetition(db: Database.Database, compId: number): Promise<void> {
  const sb = getSupabaseClient();
  const { error } = await sb.from('competitions').delete().eq('id', compId);
  if (error) throw formatSupabaseError(error, 'تعذر حذف المسابقة من Supabase');
  await syncSupabaseToLocalDb(db);
}

export async function supabaseSyncCompetitionTargetsFromDb(db: Database.Database, compId: number): Promise<void> {
  const sb = getSupabaseClient();
  const rows = db
    .prepare(`SELECT competition_id, participant_key, target_qty, industrial_target_qty, commercial_target_qty FROM competition_targets WHERE competition_id = ?`)
    .all(compId) as Array<{
    competition_id: number;
    participant_key: string;
    target_qty: number;
    industrial_target_qty: number;
    commercial_target_qty: number;
  }>;
  if (rows.length === 0) return;
  for (let i = 0; i < rows.length; i += 300) {
    const slice = rows.slice(i, i + 300);
    const { error } = await sb.from('competition_targets').upsert(slice, { onConflict: 'competition_id,participant_key' });
    if (error) throw formatSupabaseError(error, 'تعذر مزامنة أهداف المسابقة مع Supabase');
  }
}

export async function supabaseInsertRuntimeError(err: {
  kind: string;
  tabName: string;
  message: string;
  stack: string;
  sourceUrl: string;
  lineNo: number;
  colNo: number;
  userId: number | null;
  username: string;
  occurredAt: string;
}): Promise<void> {
  const sb = getSupabaseClient();
  await sb.from('runtime_error_logs').insert({
    kind: err.kind,
    tab_name: err.tabName,
    message: err.message,
    stack: err.stack,
    source_url: err.sourceUrl,
    line_no: err.lineNo,
    col_no: err.colNo,
    user_id: err.userId,
    username: err.username,
    resolved: false,
    ai_analysis: '',
    occurred_at: err.occurredAt,
  });
}
