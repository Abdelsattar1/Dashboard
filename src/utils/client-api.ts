import {
  getBranchesAnalytics,
  getDashboardSummary,
  getDepartmentsChainsBreakdown,
  getItemCategoriesAnalytics,
  getItemsAnalytics,
  getRepsAnalytics,
  getSalesTransactions,
  getTimelineSeries,
  saveItemCategoryTargets,
  type ItemCategoryTargetInput,
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

export async function handleClientApiRequest<T>(
  rawUrl: string,
  options: RequestInit = {},
  authToken: string | null
): Promise<T> {
  const db = await getClientDb();
  const seedInfo = await ensureDefaultAdminSeeded(db);
  if (seedInfo.seededNow) {
    await persistClientDb();
  }

  const parsedUrl = new URL(rawUrl, 'http://localhost');
  const pathname = parsedUrl.pathname;
  const searchParams = parsedUrl.searchParams;
  const method = (options.method || 'GET').toUpperCase();
  const currentUser = resolveClientSessionUser(db, authToken);

  // ==================== HEALTH & AUTH ====================
  if (pathname === '/api/health' && method === 'GET') {
    return { status: 'ok', mode: 'client' } as T;
  }

  if (pathname === '/api/auth/status' && method === 'GET') {
    return {
      needsSetup: false,
      hasEnvPasscode: false,
      clientMode: true,
      defaultAdminAvailable: seedInfo.defaultAdminPendingChange,
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
    const nextHash = await hashPasswordClient(newPassword);
    db.prepare('UPDATE users SET password_hash = ?, must_change_password = 0 WHERE id = ?').run(nextHash, u.id);
    await persistClientDb();
    return { success: true } as T;
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
    const reps = getRepsAnalytics(db, u, { source, startDate, endDate, sector });
    const { diagnostics } = getBranchesAnalytics(db, u, { source, startDate, endDate, sector });
    return { reps, diagnostics } as T;
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
      const arrBuf = await file.arrayBuffer();
      return validateExcelBuffer(new Uint8Array(arrBuf), file.name) as T;
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
      const batch = commitValidatedImport(db, validationToken, { id: u.id, fullName: u.fullName });
      await persistClientDb();
      return { success: true, batch } as T;
    } catch (err) {
      throw new ClientHttpError(400, err instanceof Error ? err.message : 'فشل اعتماد البيانات');
    }
  }

  if (pathname === '/api/import/history' && method === 'GET') {
    requireAdminUser(currentUser);
    const rows = db.prepare(`SELECT * FROM import_batches ORDER BY id DESC`).all() as Array<{
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
    }>;
    const batches: ImportBatchHistoryItem[] = rows.map((r) => ({
      id: r.id,
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
    return { batches } as T;
  }

  const rollbackBatchMatch = pathname.match(/^\/api\/import\/rollback\/(\d+)$/);
  if (rollbackBatchMatch && method === 'POST') {
    requireAdminUser(currentUser);
    const batchId = Number(rollbackBatchMatch[1]);
    try {
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
    const rows = db.prepare('SELECT * FROM users ORDER BY id ASC').all() as ClientDbUserRow[];
    return { users: rows.map(mapClientUserRow) } as T;
  }

  if (pathname === '/api/users' && method === 'POST') {
    requireAdminUser(currentUser);
    const { username, fullName, email, role, allowedSectors, password } = parseJsonBody<{
      username?: string;
      fullName?: string;
      email?: string;
      role?: string;
      allowedSectors?: string[];
      password?: string;
    }>(options.body);

    const cleanUsername = String(username || '').trim().toLowerCase();
    const cleanName = String(fullName || '').trim();
    const cleanPass = String(password || '');
    const cleanRole = parseStrictUserRole(role ?? 'standard_user');

    if (!cleanRole) throw new ClientHttpError(400, 'الدور الوظيفي المحدد غير صالح');
    if (!cleanUsername || cleanUsername.length < 3) {
      throw new ClientHttpError(400, 'اسم المستخدم يجب أن يكون 3 أحرف على الأقل');
    }
    if (!cleanName) throw new ClientHttpError(400, 'يرجى إدخال الاسم الكامل');
    if (cleanPass.length < 8) {
      throw new ClientHttpError(400, 'كلمة المرور يجب أن لا تقل عن 8 أحرف');
    }

    const exists = db.prepare('SELECT id FROM users WHERE lower(username) = ?').get(cleanUsername);
    if (exists) throw new ClientHttpError(409, 'اسم المستخدم مسجل بالفعل');

    const sectors = cleanRole === 'admin' ? [...ORIGINAL_SECTORS] : sanitizeSectorsInput(allowedSectors);
    if (cleanRole !== 'admin' && sectors.length === 0) {
      throw new ClientHttpError(400, 'يرجى تحديد قطاع واحد على الأقل لهذا المستخدم');
    }

    const now = new Date().toISOString();
    const pwHash = await hashPasswordClient(cleanPass);
    const resInsert = db
      .prepare(
        `INSERT INTO users (username, full_name, email, role, allowed_sectors, password_hash, must_change_password, is_active, created_at)
         VALUES (?, ?, ?, ?, ?, ?, 0, 1, ?)`
      )
      .run(
        cleanUsername,
        cleanName,
        email ? String(email).trim() : null,
        cleanRole,
        JSON.stringify(sectors),
        pwHash,
        now
      );

    await persistClientDb();
    const created = db
      .prepare('SELECT * FROM users WHERE id = ?')
      .get(Number(resInsert.lastInsertRowid)) as ClientDbUserRow;
    return { user: mapClientUserRow(created) } as T;
  }

  const userPatchMatch = pathname.match(/^\/api\/users\/(\d+)$/);
  if (userPatchMatch && method === 'PATCH') {
    const admin = requireAdminUser(currentUser);
    const targetId = Number(userPatchMatch[1]);
    const existing = db.prepare('SELECT * FROM users WHERE id = ?').get(targetId) as ClientDbUserRow | undefined;
    if (!existing) throw new ClientHttpError(404, 'المستخدم غير موجود');

    const { fullName, email, role, allowedSectors, isActive, newPassword } = parseJsonBody<{
      fullName?: string;
      email?: string | null;
      role?: string;
      allowedSectors?: string[];
      isActive?: boolean;
      newPassword?: string;
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

    let nextSectorsJson = existing.allowed_sectors;
    if (nextRole === 'admin') {
      nextSectorsJson = JSON.stringify([...ORIGINAL_SECTORS]);
    } else if (allowedSectors !== undefined) {
      const cleaned = sanitizeSectorsInput(allowedSectors);
      if (cleaned.length === 0) throw new ClientHttpError(400, 'يرجى تحديد قطاع واحد على الأقل لهذا المستخدم');
      nextSectorsJson = JSON.stringify(cleaned);
    } else {
      let existingList: string[] = [];
      try {
        const parsed = JSON.parse(existing.allowed_sectors);
        if (Array.isArray(parsed)) existingList = sanitizeSectorsInput(parsed);
      } catch {
        existingList = [];
      }
      if (existingList.length === 0) throw new ClientHttpError(400, 'يرجى تحديد قطاع واحد على الأقل لهذا المستخدم');
      nextSectorsJson = JSON.stringify(existingList);
    }

    if (newPassword && newPassword.length > 0 && newPassword.length < 8) {
      throw new ClientHttpError(400, 'كلمة المرور الجديدة يجب أن تكون 8 أحرف على الأقل');
    }
    const nextHash =
      newPassword && newPassword.length >= 8 ? await hashPasswordClient(newPassword) : existing.password_hash;
    const nextMustChange =
      newPassword && newPassword.length >= 8 ? 0 : existing.must_change_password;

    db.prepare(
      `UPDATE users
       SET full_name = ?, email = ?, role = ?, allowed_sectors = ?, is_active = ?, password_hash = ?, must_change_password = ?
       WHERE id = ?`
    ).run(
      fullName ? String(fullName).trim() : existing.full_name,
      email !== undefined ? (email ? String(email).trim() : null) : existing.email,
      nextRole,
      nextSectorsJson,
      nextActive,
      nextHash,
      nextMustChange,
      targetId
    );

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
    db.prepare('DELETE FROM users WHERE id = ?').run(targetId);
    await persistClientDb();
    return { success: true } as T;
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
