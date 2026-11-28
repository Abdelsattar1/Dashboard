import {
  getBranchesAnalytics,
  getDashboardSummary,
  getDepartmentsChainsBreakdown,
  getItemCategoriesAnalytics,
  getItemsAnalytics,
  getRepsAnalytics,
  getSalesTransactions,
  getSavedRepTargetMode,
  getTimelineSeries,
  saveItemCategoryTargets,
  saveRepManualTargets,
  saveRepTargetMode,
  type ItemCategoryTargetInput,
  type RepManualTargetInput,
} from '../../server/analytics-engine.ts';
import {
  createCompetition,
  deleteCompetition,
  getAvailableCatalogItems,
  getCompetitionAnalytics,
  listCompetitions,
  saveCompetitionTargets,
  updateCompetition,
  type CompetitionItemInput,
  type CompetitionTargetInput,
} from '../../server/competitions-engine.ts';
import {
  canonicalizeSectorName,
  commitValidatedImport,
  generateStandardTemplateWorkbook,
  getActiveSalesPeriods,
  ORIGINAL_SECTORS,
  rollbackToBatch,
  validateExcelBuffer,
} from '../../server/excel-engine.ts';
import {
  executeAutoFix,
  recordRuntimeError,
  rollbackAutoFix,
  runSmartAudit,
} from '../../server/smart-audit-core.ts';
import type {
  AutoFixActionType,
  ImportBatchHistoryItem,
  SalesSource,
  UserProfile,
  UserRole,
} from '../types/bi.ts';
import {
  ensureDefaultAdminSeeded,
  hashPasswordClient,
  issueClientSessionToken,
  mapClientUserRow,
  resolveClientSessionUser,
  verifyPasswordClient,
  type ClientDbUserRow,
} from './client-auth.ts';
import { exportClientDbBytes, getClientDb, persistClientDb } from './client-db.ts';
import { AuthorizationError } from './errors.ts';
import {
  isSupabaseConfigured,
  supabaseCreateCompetition,
  supabaseCreateUser,
  supabaseDeleteCompetition,
  supabaseDeleteUser,
  supabaseEnsureDefaultAdmin,
  supabaseFindUserByUsername,
  supabaseInsertRuntimeError,
  supabaseListImportBatches,
  supabaseListUsers,
  supabasePushCommittedBatch,
  supabaseRollbackToBatch,
  supabaseSaveRepTargetMode,
  supabaseSyncCompetitionTargetsFromDb,
  supabaseSyncItemCategoryTargetsFromDb,
  supabaseTestConnection,
  supabaseUpdateCompetition,
  supabaseUpdateUser,
  supabaseUpsertRepManualTargetsFromDb,
  syncSupabaseToLocalDb,
} from './supabase-sync.ts';
import type Database from 'better-sqlite3';

export class ClientHttpError extends Error {
  status: number;
  constructor(status: number, message: string) {
    super(message);
    this.status = status;
  }
}

const parseSource = (raw: string | null): SalesSource => (raw === 'ZSALES' ? 'ZSALES' : 'VF05N');

const parseStrictUserRole = (raw: unknown): UserRole | null => {
  if (raw === 'admin' || raw === 'sector_manager' || raw === 'standard_user') {
    return raw;
  }
  return null;
};

const sanitizeSectorsInput = (raw: unknown): string[] => {
  if (!Array.isArray(raw)) return [];
  const list = raw.map((s) => String(s || '').trim()).filter(Boolean);
  if (list.some((s) => s === 'الكل' || s.toLowerCase() === 'all')) {
    return [...ORIGINAL_SECTORS];
  }
  const valid = list
    .map((s) => canonicalizeSectorName(s))
    .filter((s) => ORIGINAL_SECTORS.includes(s as (typeof ORIGINAL_SECTORS)[number]));
  return Array.from(new Set(valid));
};

const ALLOWED_AUTOFIX_ACTIONS = new Set<AutoFixActionType>([
  'FIX_NON_FINITE_NUMBERS',
  'FIX_DUPLICATE_BRANCH_RECORDS',
  'FIX_DUPLICATE_ITEM_CATALOG',
  'FIX_SYNC_COMPETITION_TARGET_TOTALS',
  'FIX_CANONICALIZE_TARGETS_AND_SCOPES',
  'FIX_RESOLVE_RUNTIME_ERRORS',
]);

function parseJsonBody<T>(body: BodyInit | null | undefined): T {
  if (!body) return {} as T;
  if (typeof body === 'string') {
    return JSON.parse(body) as T;
  }
  return {} as T;
}

function requireUser(user: UserProfile | null): UserProfile {
  if (!user) {
    throw new ClientHttpError(401, 'مطلوب تسجيل الدخول للوصول إلى هذه البيانات');
  }
  return user;
}

function requireAdminUser(user: UserProfile | null): UserProfile {
  const u = requireUser(user);
  if (u.role !== 'admin') {
    throw new ClientHttpError(403, 'هذه العملية متاحة لمدير النظام (Admin) فقط');
  }
  return u;
}

function requireTargetEditorUser(user: UserProfile | null): UserProfile {
  const u = requireUser(user);
  if (u.role !== 'admin' && u.role !== 'sector_manager') {
    throw new ClientHttpError(403, 'تعديل الأهداف متاح للأدمن ومدير القطاع أو الإدارة فقط');
  }
  return u;
}

let lastCloudSyncAt = 0;
let cloudSyncInFlight: Promise<void> | null = null;

async function ensureCloudSynced(db: Database.Database, opts?: { force?: boolean; forceBatch?: boolean }): Promise<void> {
  if (!isSupabaseConfigured()) return;
  const now = Date.now();
  if (!opts?.force && !opts?.forceBatch && now - lastCloudSyncAt < 4000) {
    return;
  }
  if (cloudSyncInFlight) {
    await cloudSyncInFlight;
    return;
  }
  cloudSyncInFlight = (async () => {
    try {
      await syncSupabaseToLocalDb(db, { forceBatch: opts?.forceBatch });
      lastCloudSyncAt = Date.now();
    } finally {
      cloudSyncInFlight = null;
    }
  })();
  await cloudSyncInFlight;
}

export async function handleClientApiRequest<T>(
  rawUrl: string,
  options: RequestInit = {},
  authToken: string | null
): Promise<T> {
  const db = await getClientDb();
  const useSupabase = isSupabaseConfigured();

  const parsedUrl = new URL(rawUrl, 'http://localhost');
  const pathname = parsedUrl.pathname;
  const searchParams = parsedUrl.searchParams;
  const method = (options.method || 'GET').toUpperCase();

  // ==================== HEALTH & AUTH ====================
  if (pathname === '/api/health' && method === 'GET') {
    return { status: 'ok', mode: useSupabase ? 'supabase' : 'client' } as T;
  }

  if (pathname === '/api/auth/status' && method === 'GET') {
    if (useSupabase) {
      try {
        const sbSeed = await supabaseEnsureDefaultAdmin();
        return {
          needsSetup: false,
          hasEnvPasscode: false,
          clientMode: true,
          defaultAdminAvailable: sbSeed.defaultAdminPendingChange,
          supabaseConfigured: true,
        } as T;
      } catch (err) {
        return {
          needsSetup: false,
          hasEnvPasscode: false,
          clientMode: true,
          defaultAdminAvailable: false,
          supabaseConfigured: true,
          supabaseSetupError: err instanceof Error ? err.message : 'تعذر الاتصال بـ Supabase',
        } as T;
      }
    }

    const seedInfo = await ensureDefaultAdminSeeded(db);
    if (seedInfo.seededNow) {
      await persistClientDb();
    }
    return {
      needsSetup: false,
      hasEnvPasscode: false,
      clientMode: true,
      defaultAdminAvailable: seedInfo.defaultAdminPendingChange,
      supabaseConfigured: false,
    } as T;
  }

  if (pathname === '/api/auth/setup-admin' && method === 'POST') {
    throw new ClientHttpError(403, 'حساب مدير النظام الافتراضي جاهز مسبقاً. يرجى تسجيل الدخول.');
  }

  if (pathname === '/api/auth/login' && method === 'POST') {
    const { username, password } = parseJsonBody<{ username?: string; password?: string }>(options.body);
    const cleanUsername = String(username || '').trim().toLowerCase();
    const cleanPass = String(password || '');

    if (!cleanUsername || !cleanPass) {
      throw new ClientHttpError(400, 'يرجى إدخال اسم المستخدم وكلمة المرور');
    }

    if (useSupabase) {
      try {
        await supabaseEnsureDefaultAdmin();
        const remoteRow = await supabaseFindUserByUsername(cleanUsername);
        if (!remoteRow || !remoteRow.is_active) {
          throw new ClientHttpError(401, 'بيانات الدخول غير صحيحة أو الحساب غير مفعل');
        }
        const ok = await verifyPasswordClient(cleanPass, remoteRow.password_hash);
        if (!ok) {
          throw new ClientHttpError(401, 'بيانات الدخول غير صحيحة');
        }
        await ensureCloudSynced(db, { force: true });
        await persistClientDb();
        const localUser = db.prepare('SELECT * FROM users WHERE id = ?').get(Number(remoteRow.id)) as ClientDbUserRow | undefined;
        const profile = localUser ? mapClientUserRow(localUser) : mapClientUserRow({
          id: Number(remoteRow.id),
          username: remoteRow.username,
          full_name: remoteRow.full_name,
          email: remoteRow.email ?? null,
          role: remoteRow.role,
          allowed_sectors: Array.isArray(remoteRow.allowed_sectors)
            ? JSON.stringify(remoteRow.allowed_sectors)
            : String(remoteRow.allowed_sectors || '[]'),
          password_hash: remoteRow.password_hash,
          must_change_password: remoteRow.must_change_password ? 1 : 0,
          is_active: remoteRow.is_active ? 1 : 0,
          created_at: remoteRow.created_at,
        });
        const token = issueClientSessionToken(profile.id);
        return { token, user: profile } as T;
      } catch (err) {
        if (err instanceof ClientHttpError) throw err;
        throw new ClientHttpError(400, err instanceof Error ? err.message : 'تعذر تسجيل الدخول عبر Supabase');
      }
    }

    const seedInfo = await ensureDefaultAdminSeeded(db);
    if (seedInfo.seededNow) {
      await persistClientDb();
    }

    const row = db
      .prepare('SELECT * FROM users WHERE lower(username) = ?')
      .get(cleanUsername) as ClientDbUserRow | undefined;

    if (!row || row.is_active !== 1) {
      throw new ClientHttpError(401, 'بيانات الدخول غير صحيحة أو الحساب غير مفعل');
    }

    const ok = await verifyPasswordClient(cleanPass, row.password_hash);
    if (!ok) {
      throw new ClientHttpError(401, 'بيانات الدخول غير صحيحة');
    }

    const profile = mapClientUserRow(row);
    const token = issueClientSessionToken(profile.id);
    return { token, user: profile } as T;
  }

  if (!useSupabase) {
    const seedInfo = await ensureDefaultAdminSeeded(db);
    if (seedInfo.seededNow) {
      await persistClientDb();
    }
  } else if (method === 'GET' && !pathname.startsWith('/api/admin/supabase-status')) {
    await ensureCloudSynced(db);
  }

  const currentUser = resolveClientSessionUser(db, authToken);

  if (pathname === '/api/auth/me' && method === 'GET') {
    const u = requireUser(currentUser);
    return { user: u } as T;
  }

  if (pathname === '/api/auth/change-password' && method === 'POST') {
    const u = requireUser(currentUser);
    const { currentPassword, newPassword } = parseJsonBody<{ currentPassword?: string; newPassword?: string }>(
      options.body
    );
    if (!newPassword || newPassword.length < 8) {
      throw new ClientHttpError(400, 'كلمة المرور الجديدة يجب أن لا تقل عن 8 أحرف');
    }
    const row = db.prepare('SELECT * FROM users WHERE id = ?').get(u.id) as ClientDbUserRow | undefined;
    if (!row || !(await verifyPasswordClient(String(currentPassword || ''), row.password_hash))) {
      throw new ClientHttpError(400, 'كلمة المرور الحالية غير صحيحة');
    }
    if (row.must_change_password === 1 && String(currentPassword || '') === newPassword) {
      throw new ClientHttpError(400, 'كلمة المرور الجديدة يجب أن تختلف عن كلمة المرور المبدئية');
    }
    const nextHash = await hashPasswordClient(newPassword);
    if (useSupabase) {
      const updatedUser = await supabaseUpdateUser(u.id, {
        password_hash: nextHash,
        must_change_password: false,
      });
      db.prepare('UPDATE users SET password_hash = ?, must_change_password = 0 WHERE id = ?').run(nextHash, u.id);
      lastCloudSyncAt = Date.now();
      await persistClientDb();
      return { success: true, user: updatedUser } as T;
    }
    db.prepare('UPDATE users SET password_hash = ?, must_change_password = 0 WHERE id = ?').run(nextHash, u.id);
    await persistClientDb();
    const updated = db.prepare('SELECT * FROM users WHERE id = ?').get(u.id) as ClientDbUserRow;
    return { success: true, user: mapClientUserRow(updated) } as T;
  }

  // ==================== ANALYTICS ====================
  if (pathname === '/api/analytics/dashboard' && method === 'GET') {
    const u = requireUser(currentUser);
    const source = parseSource(searchParams.get('source'));
    const startDate = searchParams.get('startDate')?.trim() || undefined;
    const endDate = searchParams.get('endDate')?.trim() || undefined;
    const sector = searchParams.get('sector')?.trim() || undefined;
    return getDashboardSummary(db, u, { source, startDate, endDate, sector }) as T;
  }

  if (pathname === '/api/analytics/sectors' && method === 'GET') {
    const u = requireUser(currentUser);
    const source = parseSource(searchParams.get('source'));
    const startDate = searchParams.get('startDate')?.trim() || undefined;
    const endDate = searchParams.get('endDate')?.trim() || undefined;
    const sector = searchParams.get('sector')?.trim() || undefined;
    const summary = getDashboardSummary(db, u, { source, startDate, endDate, sector });
    const chains = getDepartmentsChainsBreakdown(db, u, { source, startDate, endDate, sector });
    return { sectors: summary.sectors, aggregations: summary.aggregations, chains } as T;
  }

  if (pathname === '/api/analytics/branches' && method === 'GET') {
    const u = requireUser(currentUser);
    const source = parseSource(searchParams.get('source'));
    const startDate = searchParams.get('startDate')?.trim() || undefined;
    const endDate = searchParams.get('endDate')?.trim() || undefined;
    const sector = searchParams.get('sector')?.trim() || undefined;
    return getBranchesAnalytics(db, u, { source, startDate, endDate, sector }) as T;
  }

  if (pathname === '/api/analytics/reps' && method === 'GET') {
    const u = requireUser(currentUser);
    const source = parseSource(searchParams.get('source'));
    const startDate = searchParams.get('startDate')?.trim() || undefined;
    const endDate = searchParams.get('endDate')?.trim() || undefined;
    const sector = searchParams.get('sector')?.trim() || undefined;
    const rawMode = searchParams.get('targetMode')?.trim() || '';
    const targetMode =
      rawMode === 'manual' || rawMode === 'auto' ? rawMode : getSavedRepTargetMode(db);
    const reps = getRepsAnalytics(db, u, { source, startDate, endDate, sector, targetMode });
    const { diagnostics } = getBranchesAnalytics(db, u, { source, startDate, endDate, sector });
    return { reps, targetMode, diagnostics } as T;
  }

  if (pathname === '/api/analytics/reps/target-mode' && method === 'PUT') {
    const u = requireTargetEditorUser(currentUser);
    const { targetMode } = parseJsonBody<{ targetMode?: string }>(options.body);
    if (targetMode !== 'auto' && targetMode !== 'manual') {
      throw new ClientHttpError(400, 'وضع الأهداف غير صالح (يجب أن يكون auto أو manual)');
    }
    try {
      const saved = saveRepTargetMode(db, targetMode, u);
      if (useSupabase) {
        await supabaseSaveRepTargetMode(saved);
        lastCloudSyncAt = Date.now();
      }
      await persistClientDb();
      return { success: true, targetMode: saved } as T;
    } catch (err) {
      if (err instanceof AuthorizationError) {
        throw new ClientHttpError(403, err.message);
      }
      throw new ClientHttpError(400, err instanceof Error ? err.message : 'تعذر تحديث وضع أهداف المناديب');
    }
  }

  if (pathname === '/api/analytics/reps/manual-targets' && method === 'PUT') {
    const u = requireTargetEditorUser(currentUser);
    const { targets } = parseJsonBody<{ targets?: RepManualTargetInput[] }>(options.body);
    if (!Array.isArray(targets)) {
      throw new ClientHttpError(400, 'صيغة أهداف المناديب غير صالحة');
    }
    try {
      const savedCodes = saveRepManualTargets(db, targets, u);
      if (useSupabase) {
        await supabaseUpsertRepManualTargetsFromDb(db, savedCodes);
        lastCloudSyncAt = Date.now();
      }
      await persistClientDb();
      return { success: true } as T;
    } catch (err) {
      if (err instanceof AuthorizationError) {
        throw new ClientHttpError(403, err.message);
      }
      throw new ClientHttpError(400, err instanceof Error ? err.message : 'تعذر حفظ الأهداف اليدوية للمناديب');
    }
  }

  if (pathname === '/api/analytics/items' && method === 'GET') {
    const u = requireUser(currentUser);
    const source = parseSource(searchParams.get('source'));
    const startDate = searchParams.get('startDate')?.trim() || undefined;
    const endDate = searchParams.get('endDate')?.trim() || undefined;
    const sector = searchParams.get('sector')?.trim() || undefined;
    const jobTitle = searchParams.get('jobTitle')?.trim() || undefined;
    const categoryAnalysis = getItemCategoriesAnalytics(db, u, { source, startDate, endDate, sector, jobTitle });
    const items = getItemsAnalytics(db, u, { source, startDate, endDate, sector });
    return { ...categoryAnalysis, items } as T;
  }

  if (pathname === '/api/analytics/item-category-targets' && method === 'PUT') {
    const u = requireTargetEditorUser(currentUser);
    const { targets } = parseJsonBody<{ targets?: ItemCategoryTargetInput[] }>(options.body);
    if (!Array.isArray(targets)) {
      throw new ClientHttpError(400, 'صيغة أهداف الفئات غير صالحة');
    }
    try {
      saveItemCategoryTargets(db, targets, u);
      if (useSupabase) {
        await supabaseSyncItemCategoryTargetsFromDb(db);
        lastCloudSyncAt = Date.now();
      }
      await persistClientDb();
      return { success: true } as T;
    } catch (err) {
      if (err instanceof AuthorizationError) {
        throw new ClientHttpError(403, err.message);
      }
      throw new ClientHttpError(400, err instanceof Error ? err.message : 'تعذر حفظ أهداف الفئات');
    }
  }

  if (pathname === '/api/analytics/timeline' && method === 'GET') {
    const u = requireUser(currentUser);
    const source = parseSource(searchParams.get('source'));
    const startDate = searchParams.get('startDate')?.trim() || undefined;
    const endDate = searchParams.get('endDate')?.trim() || undefined;
    const sector = searchParams.get('sector')?.trim() || undefined;
    const series = getTimelineSeries(db, u, { source, startDate, endDate, sector });
    return { series } as T;
  }

  if (pathname === '/api/analytics/transactions' && method === 'GET') {
    const u = requireUser(currentUser);
    const source = parseSource(searchParams.get('source'));
    const startDate = searchParams.get('startDate')?.trim() || undefined;
    const endDate = searchParams.get('endDate')?.trim() || undefined;
    const sector = searchParams.get('sector')?.trim() || undefined;
    const search = searchParams.get('search') || undefined;
    const linkStatusRaw = searchParams.get('linkStatus') || 'all';
    const linkStatus = ['all', 'linked', 'unlinked_branch', 'unlinked_sector', 'returns'].includes(linkStatusRaw)
      ? (linkStatusRaw as 'all' | 'linked' | 'unlinked_branch' | 'unlinked_sector' | 'returns')
      : 'all';
    const page = Number(searchParams.get('page') || 1);
    const pageSize = Number(searchParams.get('pageSize') || 50);
    return getSalesTransactions(db, u, { source, startDate, endDate, sector, search, linkStatus, page, pageSize }) as T;
  }

  // ==================== COMPETITIONS ====================
  if (pathname === '/api/competitions' && method === 'GET') {
    requireUser(currentUser);
    return {
      competitions: listCompetitions(db),
      catalogItems: getAvailableCatalogItems(db),
    } as T;
  }

  if (pathname === '/api/competitions/analytics' && method === 'GET') {
    const u = requireUser(currentUser);
    const compIdRaw = searchParams.get('competitionId') ? Number(searchParams.get('competitionId')) : null;
    const compId = compIdRaw && Number.isInteger(compIdRaw) && compIdRaw > 0 ? compIdRaw : null;
    const sector = searchParams.get('sector')?.trim() || undefined;
    const startDate = searchParams.get('startDate')?.trim() || undefined;
    const endDate = searchParams.get('endDate')?.trim() || undefined;
    return getCompetitionAnalytics(db, u, compId, { sector, startDate, endDate }) as T;
  }

  const compAnalyticsMatch = pathname.match(/^\/api\/competitions\/(\d+)\/analytics$/);
  if (compAnalyticsMatch && method === 'GET') {
    const u = requireUser(currentUser);
    const compId = Number(compAnalyticsMatch[1]);
    const sector = searchParams.get('sector')?.trim() || undefined;
    const startDate = searchParams.get('startDate')?.trim() || undefined;
    const endDate = searchParams.get('endDate')?.trim() || undefined;
    return getCompetitionAnalytics(db, u, compId > 0 ? compId : null, { sector, startDate, endDate }) as T;
  }

  if (pathname === '/api/competitions' && method === 'POST') {
    const u = requireAdminUser(currentUser);
    const { name, items } = parseJsonBody<{ name?: string; items?: CompetitionItemInput[] }>(options.body);
    try {
      const created = createCompetition(db, { id: u.id }, { name: String(name || ''), items });
      if (useSupabase) {
        const remoteCompId = await supabaseCreateCompetition(db, u.id, created.name, created.items);
        lastCloudSyncAt = Date.now();
        await persistClientDb();
        const syncedList = listCompetitions(db);
        const syncedComp = syncedList.find((c) => c.id === remoteCompId) || created;
        return { competition: syncedComp } as T;
      }
      await persistClientDb();
      return { competition: created } as T;
    } catch (err) {
      throw new ClientHttpError(400, err instanceof Error ? err.message : 'تعذر إنشاء المسابقة');
    }
  }

  const compItemMatch = pathname.match(/^\/api\/competitions\/(\d+)$/);
  if (compItemMatch && method === 'PUT') {
    requireAdminUser(currentUser);
    const compId = Number(compItemMatch[1]);
    const { name, isActive, items } = parseJsonBody<{ name?: string; isActive?: boolean; items?: CompetitionItemInput[] }>(
      options.body
    );
    try {
      const updated = updateCompetition(db, compId, { name, isActive, items });
      if (useSupabase) {
        await supabaseUpdateCompetition(db, compId, {
          name: updated.name,
          isActive: updated.isActive,
          items: updated.items,
        });
        lastCloudSyncAt = Date.now();
      }
      await persistClientDb();
      return { competition: updated } as T;
    } catch (err) {
      throw new ClientHttpError(400, err instanceof Error ? err.message : 'تعذر تحديث المسابقة');
    }
  }

  if (compItemMatch && method === 'DELETE') {
    requireAdminUser(currentUser);
    const compId = Number(compItemMatch[1]);
    try {
      deleteCompetition(db, compId);
      if (useSupabase) {
        await supabaseDeleteCompetition(db, compId);
        lastCloudSyncAt = Date.now();
      }
      await persistClientDb();
      return { success: true } as T;
    } catch (err) {
      throw new ClientHttpError(404, err instanceof Error ? err.message : 'تعذر حذف المسابقة');
    }
  }

  const compTargetsMatch = pathname.match(/^\/api\/competitions\/(\d+)\/targets$/);
  if (compTargetsMatch && method === 'PUT') {
    const u = requireTargetEditorUser(currentUser);
    const compId = Number(compTargetsMatch[1]);
    const { targets } = parseJsonBody<{ targets?: CompetitionTargetInput[] }>(options.body);
    if (!Array.isArray(targets)) {
      throw new ClientHttpError(400, 'صيغة الأهداف غير صالحة');
    }
    try {
      saveCompetitionTargets(db, compId, targets, u);
      if (useSupabase) {
        await supabaseSyncCompetitionTargetsFromDb(db, compId);
        lastCloudSyncAt = Date.now();
      }
      await persistClientDb();
      return { success: true } as T;
    } catch (err) {
      if (err instanceof AuthorizationError) {
        throw new ClientHttpError(403, err.message);
      }
      throw new ClientHttpError(400, err instanceof Error ? err.message : 'تعذر حفظ الأهداف الكمية');
    }
  }

  // ==================== EXCEL IMPORT ====================
  if (pathname === '/api/import/validate' && method === 'POST') {
    requireAdminUser(currentUser);
    if (!(options.body instanceof FormData)) {
      throw new ClientHttpError(400, 'لم يتم إرفاق ملف Excel صالح');
    }
    const file = options.body.get('file');
    if (!(file instanceof File) || file.size === 0) {
      throw new ClientHttpError(400, 'لم يتم إرفاق ملف Excel صالح');
    }
    try {
      if (useSupabase) {
        await ensureCloudSynced(db);
      }
      const existingPeriods = getActiveSalesPeriods(db);
      const arrBuf = await file.arrayBuffer();
      return validateExcelBuffer(new Uint8Array(arrBuf), file.name, existingPeriods) as T;
    } catch (err) {
      throw new ClientHttpError(400, err instanceof Error ? err.message : 'تعذر تحليل ملف Excel');
    }
  }

  if (pathname === '/api/import/commit' && method === 'POST') {
    const u = requireAdminUser(currentUser);
    const { validationToken } = parseJsonBody<{ validationToken?: string }>(options.body);
    if (!validationToken) {
      throw new ClientHttpError(400, 'رمز التحقق مطلوب لاعتماد الاستيراد');
    }
    try {
      if (useSupabase) {
        await ensureCloudSynced(db);
      }
      const batch = commitValidatedImport(db, validationToken, { id: u.id, fullName: u.fullName });
      if (useSupabase) {
        const remoteBatch = await supabasePushCommittedBatch(db, batch.id);
        lastCloudSyncAt = Date.now();
        await persistClientDb();
        return { success: true, batch: remoteBatch } as T;
      }
      await persistClientDb();
      return { success: true, batch } as T;
    } catch (err) {
      throw new ClientHttpError(400, err instanceof Error ? err.message : 'فشل اعتماد البيانات');
    }
  }

  if (pathname === '/api/import/history' && method === 'GET') {
    requireAdminUser(currentUser);
    if (useSupabase) {
      const batches = await supabaseListImportBatches();
      return { batches } as T;
    }
    const rows = db.prepare(`SELECT * FROM import_batches ORDER BY id DESC`).all() as Array<{
      id: number;
      batch_uuid: string;
      filename: string;
      imported_by_name: string;
      status: 'active' | 'archived';
      vf05n_count: number;
      vf05n_sales_total: number;
      vf05n_net_sales_total?: number;
      vf05n_tax_total: number;
      zsales_count: number;
      zsales_sales_total: number;
      zsales_net_sales_total?: number;
      zsales_tax_total: number;
      unlinked_branch_count: number;
      unlinked_sector_count: number;
      created_at: string;
    }>;
    const batches: ImportBatchHistoryItem[] = rows.map((r) => ({
      id: r.id,
      batchUuid: r.batch_uuid,
      filename: r.filename,
      importedByName: r.imported_by_name,
      status: r.status,
      vf05nCount: r.vf05n_count,
      vf05nSalesTotal: r.vf05n_sales_total,
      vf05nNetSalesTotal: r.vf05n_net_sales_total ?? 0,
      vf05nTaxTotal: r.vf05n_tax_total,
      zsalesCount: r.zsales_count,
      zsalesSalesTotal: r.zsales_sales_total,
      zsalesNetSalesTotal: r.zsales_net_sales_total ?? 0,
      zsalesTaxTotal: r.zsales_tax_total,
      unlinkedBranchCount: r.unlinked_branch_count,
      unlinkedSectorCount: r.unlinked_sector_count,
      createdAt: r.created_at,
    }));
    return { batches } as T;
  }

  const rollbackBatchMatch = pathname.match(/^\/api\/import\/rollback\/(\d+)$/);
  if (rollbackBatchMatch && method === 'POST') {
    requireAdminUser(currentUser);
    const batchId = Number(rollbackBatchMatch[1]);
    try {
      if (useSupabase) {
        const restored = await supabaseRollbackToBatch(db, batchId);
        lastCloudSyncAt = Date.now();
        await persistClientDb();
        return { success: true, batch: restored } as T;
      }
      const restored = rollbackToBatch(db, batchId);
      await persistClientDb();
      return { success: true, batch: restored } as T;
    } catch (err) {
      throw new ClientHttpError(404, err instanceof Error ? err.message : 'تعذر استرجاع النسخة');
    }
  }

  // ==================== USER MANAGEMENT ====================
  if (pathname === '/api/users' && method === 'GET') {
    requireAdminUser(currentUser);
    if (useSupabase) {
      const users = await supabaseListUsers();
      return { users } as T;
    }
    const rows = db.prepare('SELECT * FROM users ORDER BY id ASC').all() as ClientDbUserRow[];
    return { users: rows.map(mapClientUserRow) } as T;
  }

  if (pathname === '/api/users' && method === 'POST') {
    requireAdminUser(currentUser);
    const { username, fullName, email, role, allowedSectors, password, mustChangePassword } = parseJsonBody<{
      username?: string;
      fullName?: string;
      email?: string;
      role?: string;
      allowedSectors?: string[];
      password?: string;
      mustChangePassword?: boolean;
    }>(options.body);

    const cleanUsername = String(username || '').trim().toLowerCase();
    const cleanName = String(fullName || '').trim();
    const cleanPass = String(password || '');
    const cleanRole = parseStrictUserRole(role ?? 'standard_user');
    const forceFirstChange = typeof mustChangePassword === 'boolean' ? mustChangePassword : true;

    if (!cleanRole) throw new ClientHttpError(400, 'الدور الوظيفي المحدد غير صالح');
    if (!cleanUsername || cleanUsername.length < 3) {
      throw new ClientHttpError(400, 'اسم المستخدم يجب أن يكون 3 أحرف على الأقل');
    }
    if (!cleanName) throw new ClientHttpError(400, 'يرجى إدخال الاسم الكامل');
    if (cleanPass.length < 8) {
      throw new ClientHttpError(400, 'كلمة المرور المبدئية يجب أن لا تقل عن 8 أحرف');
    }

    const sectors = cleanRole === 'admin' ? [...ORIGINAL_SECTORS] : sanitizeSectorsInput(allowedSectors);
    if (cleanRole !== 'admin' && sectors.length === 0) {
      throw new ClientHttpError(400, 'يرجى تحديد قطاع واحد على الأقل لهذا المستخدم');
    }

    const pwHash = await hashPasswordClient(cleanPass);

    if (useSupabase) {
      const created = await supabaseCreateUser({
        username: cleanUsername,
        fullName: cleanName,
        email: email ? String(email).trim() : null,
        role: cleanRole,
        allowedSectors: sectors,
        passwordHash: pwHash,
        mustChangePassword: forceFirstChange,
      });
      await ensureCloudSynced(db, { force: true });
      await persistClientDb();
      return { user: created } as T;
    }

    const exists = db.prepare('SELECT id FROM users WHERE lower(username) = ?').get(cleanUsername);
    if (exists) throw new ClientHttpError(409, 'اسم المستخدم مسجل بالفعل');

    const now = new Date().toISOString();
    const resInsert = db
      .prepare(
        `INSERT INTO users (username, full_name, email, role, allowed_sectors, password_hash, must_change_password, is_active, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, 1, ?)`
      )
      .run(
        cleanUsername,
        cleanName,
        email ? String(email).trim() : null,
        cleanRole,
        JSON.stringify(sectors),
        pwHash,
        forceFirstChange ? 1 : 0,
        now
      );

    await persistClientDb();
    const created = db
      .prepare('SELECT * FROM users WHERE id = ?')
      .get(Number(resInsert.lastInsertRowid)) as ClientDbUserRow;
    return { user: mapClientUserRow(created) } as T;
  }

  const resetPwMatch = pathname.match(/^\/api\/users\/(\d+)\/reset-password$/);
  if (resetPwMatch && method === 'POST') {
    requireAdminUser(currentUser);
    const targetId = Number(resetPwMatch[1]);
    const { temporaryPassword } = parseJsonBody<{ temporaryPassword?: string }>(options.body);
    const cleanPass = String(temporaryPassword || '');
    if (cleanPass.length < 8) {
      throw new ClientHttpError(400, 'كلمة المرور المبدئية يجب أن لا تقل عن 8 أحرف');
    }
    const existing = db.prepare('SELECT * FROM users WHERE id = ?').get(targetId) as ClientDbUserRow | undefined;
    if (!existing) throw new ClientHttpError(404, 'المستخدم غير موجود');

    const nextHash = await hashPasswordClient(cleanPass);
    if (useSupabase) {
      const updated = await supabaseUpdateUser(targetId, {
        password_hash: nextHash,
        must_change_password: true,
      });
      db.prepare('UPDATE users SET password_hash = ?, must_change_password = 1 WHERE id = ?').run(nextHash, targetId);
      lastCloudSyncAt = Date.now();
      await persistClientDb();
      return { user: updated } as T;
    }

    db.prepare('UPDATE users SET password_hash = ?, must_change_password = 1 WHERE id = ?').run(nextHash, targetId);
    await persistClientDb();
    const updated = db.prepare('SELECT * FROM users WHERE id = ?').get(targetId) as ClientDbUserRow;
    return { user: mapClientUserRow(updated) } as T;
  }

  const userPatchMatch = pathname.match(/^\/api\/users\/(\d+)$/);
  if (userPatchMatch && method === 'PATCH') {
    const admin = requireAdminUser(currentUser);
    const targetId = Number(userPatchMatch[1]);
    const existing = db.prepare('SELECT * FROM users WHERE id = ?').get(targetId) as ClientDbUserRow | undefined;
    if (!existing) throw new ClientHttpError(404, 'المستخدم غير موجود');

    const { fullName, email, role, allowedSectors, isActive, newPassword, mustChangePassword } = parseJsonBody<{
      fullName?: string;
      email?: string | null;
      role?: string;
      allowedSectors?: string[];
      isActive?: boolean;
      newPassword?: string;
      mustChangePassword?: boolean;
    }>(options.body);

    let nextRole: UserRole =
      existing.role === 'admin'
        ? 'admin'
        : existing.role === 'sector_manager'
          ? 'sector_manager'
          : 'standard_user';
    if (role !== undefined) {
      const parsed = parseStrictUserRole(role);
      if (!parsed) throw new ClientHttpError(400, 'الدور الوظيفي المحدد غير صالح');
      nextRole = parsed;
    }

    const nextActive = typeof isActive === 'boolean' ? (isActive ? 1 : 0) : existing.is_active;
    if (targetId === admin.id && (nextRole !== 'admin' || nextActive === 0)) {
      throw new ClientHttpError(400, 'لا يمكنك تعطيل حسابك أو خفض صلاحيتك أثناء تسجيل الدخول');
    }

    let nextSectorsList: string[] = [];
    if (nextRole === 'admin') {
      nextSectorsList = [...ORIGINAL_SECTORS];
    } else if (allowedSectors !== undefined) {
      const cleaned = sanitizeSectorsInput(allowedSectors);
      if (cleaned.length === 0) throw new ClientHttpError(400, 'يرجى تحديد قطاع واحد على الأقل لهذا المستخدم');
      nextSectorsList = cleaned;
    } else {
      let existingList: string[] = [];
      try {
        const parsed = JSON.parse(existing.allowed_sectors);
        if (Array.isArray(parsed)) existingList = sanitizeSectorsInput(parsed);
      } catch {
        existingList = [];
      }
      if (existingList.length === 0) throw new ClientHttpError(400, 'يرجى تحديد قطاع واحد على الأقل لهذا المستخدم');
      nextSectorsList = existingList;
    }
    const nextSectorsJson = JSON.stringify(nextSectorsList);

    if (newPassword && newPassword.length > 0 && newPassword.length < 8) {
      throw new ClientHttpError(400, 'كلمة المرور الجديدة يجب أن تكون 8 أحرف على الأقل');
    }
    const hasNewPass = Boolean(newPassword && newPassword.length >= 8);
    const nextHash = hasNewPass ? await hashPasswordClient(newPassword!) : existing.password_hash;
    const nextMustChange =
      typeof mustChangePassword === 'boolean'
        ? mustChangePassword
          ? 1
          : 0
        : hasNewPass
          ? 0
          : existing.must_change_password;

    const nextFullName = fullName ? String(fullName).trim() : existing.full_name;
    const nextEmail = email !== undefined ? (email ? String(email).trim() : null) : existing.email;

    if (useSupabase) {
      const updatedUser = await supabaseUpdateUser(targetId, {
        full_name: nextFullName,
        email: nextEmail,
        role: nextRole,
        allowed_sectors: nextSectorsList,
        is_active: nextActive === 1,
        password_hash: nextHash,
        must_change_password: nextMustChange === 1,
      });
      db.prepare(
        `UPDATE users
         SET full_name = ?, email = ?, role = ?, allowed_sectors = ?, is_active = ?, password_hash = ?, must_change_password = ?
         WHERE id = ?`
      ).run(nextFullName, nextEmail, nextRole, nextSectorsJson, nextActive, nextHash, nextMustChange, targetId);
      lastCloudSyncAt = Date.now();
      await persistClientDb();
      return { user: updatedUser } as T;
    }

    db.prepare(
      `UPDATE users
       SET full_name = ?, email = ?, role = ?, allowed_sectors = ?, is_active = ?, password_hash = ?, must_change_password = ?
       WHERE id = ?`
    ).run(nextFullName, nextEmail, nextRole, nextSectorsJson, nextActive, nextHash, nextMustChange, targetId);

    await persistClientDb();
    const updated = db.prepare('SELECT * FROM users WHERE id = ?').get(targetId) as ClientDbUserRow;
    return { user: mapClientUserRow(updated) } as T;
  }

  if (userPatchMatch && method === 'DELETE') {
    const admin = requireAdminUser(currentUser);
    const targetId = Number(userPatchMatch[1]);
    if (targetId === admin.id) {
      throw new ClientHttpError(400, 'لا يمكنك حذف حسابك النشط');
    }
    if (useSupabase) {
      await supabaseDeleteUser(targetId);
      lastCloudSyncAt = 0;
    }
    db.prepare('DELETE FROM users WHERE id = ?').run(targetId);
    await persistClientDb();
    return { success: true } as T;
  }

  // ==================== SUPABASE STATUS ====================
  if (pathname === '/api/admin/supabase-status' && method === 'GET') {
    requireAdminUser(currentUser);
    const result = await supabaseTestConnection();
    return {
      configured: useSupabase,
      ...result,
    } as T;
  }

  // ==================== SMART AUDITOR & RUNTIME ERRORS ====================
  if (pathname === '/api/audit/runtime-errors' && method === 'POST') {
    const u = requireUser(currentUser);
    const body = parseJsonBody<{
      kind?: string;
      tabName?: string;
      message?: string;
      stack?: string;
      sourceUrl?: string;
      lineNo?: number;
      colNo?: number;
    }>(options.body);
    if (!String(body.message || '').trim()) {
      throw new ClientHttpError(400, 'رسالة الخطأ مطلوبة');
    }
    const log = recordRuntimeError(db, { id: u.id, username: u.username }, body);
    if (useSupabase && log.id > 0) {
      supabaseInsertRuntimeError({
        kind: log.kind,
        tabName: log.tabName,
        message: log.message,
        stack: log.stack,
        sourceUrl: log.sourceUrl,
        lineNo: log.lineNo,
        colNo: log.colNo,
        userId: u.id,
        username: u.username,
        occurredAt: log.occurredAt,
      }).catch(() => {});
    }
    await persistClientDb();
    return { recorded: true, id: log.id } as T;
  }

  if ((pathname === '/api/admin/audit/latest' && method === 'GET') || (pathname === '/api/admin/audit/run' && method === 'POST')) {
    const admin = requireAdminUser(currentUser);
    const report = runSmartAudit(db, admin, { aiConfigured: false });
    await persistClientDb();
    return { report } as T;
  }

  if (pathname === '/api/admin/audit/ai-analyze' && method === 'POST') {
    const admin = requireAdminUser(currentUser);
    const report = runSmartAudit(db, admin, { aiConfigured: false });
    return {
      usedAi: false,
      aiSummary:
        'يعمل المراجع الذكي في وضع المتصفح المحلي بمحرك القواعد الكامل بدون الحاجة إلى خادم خارجي.',
      report,
    } as T;
  }

  if (pathname === '/api/admin/audit/autofix' && method === 'POST') {
    const admin = requireAdminUser(currentUser);
    const { actionType, findingCode } = parseJsonBody<{ actionType?: string; findingCode?: string }>(options.body);
    if (!actionType || !ALLOWED_AUTOFIX_ACTIONS.has(actionType as AutoFixActionType)) {
      throw new ClientHttpError(400, 'نوع الإصلاح التلقائي غير مدعوم');
    }
    const result = executeAutoFix(db, admin, actionType as AutoFixActionType, findingCode, {
      aiConfigured: false,
    });
    await persistClientDb();
    return result as T;
  }

  const auditRollbackMatch = pathname.match(/^\/api\/admin\/audit\/rollback\/(\d+)$/);
  if (auditRollbackMatch && method === 'POST') {
    const admin = requireAdminUser(currentUser);
    const backupId = Number(auditRollbackMatch[1]);
    const result = rollbackAutoFix(db, admin, backupId, { aiConfigured: false });
    await persistClientDb();
    return result as T;
  }

  throw new ClientHttpError(404, `المسار غير معروف: ${pathname}`);
}

export async function handleClientDownload(
  rawUrl: string,
  fallbackFilename: string,
  authToken: string | null
): Promise<{ blob: Blob; filename: string }> {
  const db = await getClientDb();
  await ensureDefaultAdminSeeded(db);
  const currentUser = resolveClientSessionUser(db, authToken);
  requireAdminUser(currentUser);

  const parsedUrl = new URL(rawUrl, 'http://localhost');
  if (parsedUrl.pathname === '/api/import/template') {
    const bytes = generateStandardTemplateWorkbook();
    return {
      blob: new Blob([bytes.buffer as ArrayBuffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      }),
      filename: fallbackFilename || 'FMCG_Sales_Template.xlsx',
    };
  }

  if (parsedUrl.pathname === '/api/admin/export-project-zip') {
    const bytes = exportClientDbBytes();
    if (!bytes) {
      throw new ClientHttpError(400, 'لا توجد بيانات محفوظة للتصدير بعد');
    }
    return {
      blob: new Blob([bytes.buffer as ArrayBuffer], { type: 'application/octet-stream' }),
      filename: 'fmcg-sales-database-backup.sqlite',
    };
  }

  throw new ClientHttpError(404, 'تعذر العثور على الملف المطلوب');
}
