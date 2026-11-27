import 'dotenv/config';
import express from 'express';
import fs from 'fs';
import http from 'http';
import JSZip from 'jszip';
import multer from 'multer';
import path from 'path';
import { fileURLToPath } from 'url';
import { createServer as createViteServer } from 'vite';
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
} from './server/analytics-engine.ts';
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
} from './server/competitions-engine.ts';
import {
  AuthorizationError,
  createAuthMiddleware,
  hashPassword,
  mapDbUserToProfile,
  requireAdmin,
  requireTargetEditor,
  signToken,
  verifyPassword,
  type AuthenticatedRequest,
  type DbUserRow,
} from './server/auth.ts';
import { initDatabase } from './server/db.ts';
import {
  canonicalizeSectorName,
  commitValidatedImport,
  generateStandardTemplateWorkbook,
  ORIGINAL_SECTORS,
  rollbackToBatch,
  validateExcelBuffer,
} from './server/excel-engine.ts';
import {
  enrichAuditWithGemini,
  executeAutoFix,
  recordRuntimeError,
  rollbackAutoFix,
  runSmartAudit,
} from './server/smart-auditor.ts';
import type { AutoFixActionType, ImportBatchHistoryItem, SalesSource, UserRole } from './src/types/bi.ts';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const isProd = process.argv.includes('--prod') || process.env.NODE_ENV === 'production';

function resolveServerPort(): number {
  const portArgIdx = process.argv.indexOf('--port');
  if (portArgIdx !== -1 && process.argv[portArgIdx + 1]) {
    const parsed = Number(process.argv[portArgIdx + 1]);
    if (Number.isFinite(parsed) && parsed > 0) return parsed;
  }
  const envPort = Number(process.env.PORT);
  const nginxPort = Number(process.env.NGINX_PORT || 8080);
  if (!isProd && (!Number.isFinite(envPort) || envPort <= 0 || envPort === nginxPort)) {
    return Number(process.env.DEFAULT_APP_PORT || 3000);
  }
  return Number.isFinite(envPort) && envPort > 0 ? envPort : 3000;
}

const PORT = resolveServerPort();

async function startServer() {
  const db = initDatabase();
  const app = express();
  const httpServer = http.createServer(app);
  app.use(express.json({ limit: '5mb' }));

  app.get('/api/health', (_req, res) => {
    res.json({ status: 'ok' });
  });

  const upload = multer({
    storage: multer.memoryStorage(),
    limits: { fileSize: 50 * 1024 * 1024 },
    fileFilter: (_req, file, cb) => {
      const ext = path.extname(file.originalname).toLowerCase();
      if (ext !== '.xlsx' && ext !== '.xls') {
        cb(new Error('يرجى رفع ملف Excel بصيغة .xlsx فقط'));
        return;
      }
      cb(null, true);
    },
  });

  const requireAuth = createAuthMiddleware(db);

  // ==================== AUTH ROUTES ====================
  app.get('/api/auth/status', (_req, res) => {
    const adminCountRow = db
      .prepare(`SELECT COUNT(*) AS cnt FROM users WHERE role = 'admin' AND is_active = 1`)
      .get() as { cnt: number };
    const hasEnvPasscode = Boolean(process.env.ADMIN_PASSCODE && process.env.ADMIN_PASSCODE.trim().length > 0);
    res.json({
      needsSetup: adminCountRow.cnt === 0,
      hasEnvPasscode,
    });
  });

  app.post('/api/auth/setup-admin', (req, res) => {
    const existingAdmin = db
      .prepare(`SELECT COUNT(*) AS cnt FROM users WHERE role = 'admin'`)
      .get() as { cnt: number };

    if (existingAdmin.cnt > 0) {
      res.status(403).json({ error: 'تم إعداد حساب مدير النظام مسبقًا. يرجى تسجيل الدخول.' });
      return;
    }

    const { username, fullName, email, password } = req.body as {
      username?: string;
      fullName?: string;
      email?: string;
      password?: string;
    };

    const cleanUsername = String(username || '').trim().toLowerCase();
    const cleanName = String(fullName || '').trim();
    const cleanPass = String(password || '');

    if (!cleanUsername || cleanUsername.length < 3) {
      res.status(400).json({ error: 'يجب أن يكون اسم المستخدم 3 أحرف على الأقل' });
      return;
    }
    if (!cleanName) {
      res.status(400).json({ error: 'يرجى إدخال الاسم الكامل' });
      return;
    }
    if (cleanPass.length < 8) {
      res.status(400).json({ error: 'يجب أن تكون كلمة المرور 8 أحرف على الأقل' });
      return;
    }

    const now = new Date().toISOString();
    const pwHash = hashPassword(cleanPass);

    const result = db
      .prepare(
        `INSERT INTO users (username, full_name, email, role, allowed_sectors, password_hash, must_change_password, is_active, created_at)
         VALUES (?, ?, ?, 'admin', ?, ?, 0, 1, ?)`
      )
      .run(cleanUsername, cleanName, email ? String(email).trim() : null, JSON.stringify([...ORIGINAL_SECTORS]), pwHash, now);

    const createdRow = db
      .prepare('SELECT * FROM users WHERE id = ?')
      .get(Number(result.lastInsertRowid)) as DbUserRow;
    const profile = mapDbUserToProfile(createdRow);
    const token = signToken({ id: profile.id, username: profile.username, role: profile.role });

    res.status(201).json({ token, user: profile });
  });

  app.post('/api/auth/login', (req, res) => {
    const { username, password } = req.body as { username?: string; password?: string };
    const cleanUsername = String(username || '').trim().toLowerCase();
    const cleanPass = String(password || '');

    if (!cleanUsername || !cleanPass) {
      res.status(400).json({ error: 'يرجى إدخال اسم المستخدم وكلمة المرور' });
      return;
    }

    const row = db
      .prepare('SELECT * FROM users WHERE lower(username) = ?')
      .get(cleanUsername) as DbUserRow | undefined;

    if (!row || row.is_active !== 1) {
      res.status(401).json({ error: 'بيانات الدخول غير صحيحة أو الحساب غير مفعل' });
      return;
    }

    const envPasscode = process.env.ADMIN_PASSCODE?.trim();
    const envMatches = Boolean(envPasscode && row.role === 'admin' && cleanPass === envPasscode);
    const hashMatches = verifyPassword(cleanPass, row.password_hash);

    if (!envMatches && !hashMatches) {
      res.status(401).json({ error: 'بيانات الدخول غير صحيحة' });
      return;
    }

    const profile = mapDbUserToProfile(row);
    const token = signToken({ id: profile.id, username: profile.username, role: profile.role });
    res.json({ token, user: profile });
  });

  app.get('/api/auth/me', requireAuth, (req: AuthenticatedRequest, res) => {
    res.json({ user: req.user });
  });

  app.post('/api/auth/change-password', requireAuth, (req: AuthenticatedRequest, res) => {
    const { currentPassword, newPassword } = req.body as { currentPassword?: string; newPassword?: string };
    if (!newPassword || newPassword.length < 8) {
      res.status(400).json({ error: 'كلمة المرور الجديدة يجب أن لا تقل عن 8 أحرف' });
      return;
    }
    const row = db.prepare('SELECT * FROM users WHERE id = ?').get(req.user!.id) as DbUserRow | undefined;
    if (!row || !verifyPassword(String(currentPassword || ''), row.password_hash)) {
      res.status(400).json({ error: 'كلمة المرور الحالية غير صحيحة' });
      return;
    }
    if (row.must_change_password === 1 && String(currentPassword || '') === newPassword) {
      res.status(400).json({ error: 'كلمة المرور الجديدة يجب أن تختلف عن كلمة المرور المبدئية' });
      return;
    }
    db.prepare('UPDATE users SET password_hash = ?, must_change_password = 0 WHERE id = ?').run(
      hashPassword(newPassword),
      req.user!.id
    );
    const updatedRow = db.prepare('SELECT * FROM users WHERE id = ?').get(req.user!.id) as DbUserRow;
    res.json({ success: true, user: mapDbUserToProfile(updatedRow) });
  });

  // ==================== ANALYTICS ROUTES ====================
  const parseSource = (raw: unknown): SalesSource => (raw === 'ZSALES' ? 'ZSALES' : 'VF05N');

  app.get('/api/analytics/dashboard', requireAuth, (req: AuthenticatedRequest, res) => {
    const source = parseSource(req.query.source);
    const startDate = typeof req.query.startDate === 'string' ? req.query.startDate.trim() : undefined;
    const endDate = typeof req.query.endDate === 'string' ? req.query.endDate.trim() : undefined;
    const sector = typeof req.query.sector === 'string' ? req.query.sector.trim() : undefined;

    const data = getDashboardSummary(db, req.user!, { source, startDate, endDate, sector });
    res.json(data);
  });

  app.get('/api/analytics/sectors', requireAuth, (req: AuthenticatedRequest, res) => {
    const source = parseSource(req.query.source);
    const startDate = typeof req.query.startDate === 'string' ? req.query.startDate.trim() : undefined;
    const endDate = typeof req.query.endDate === 'string' ? req.query.endDate.trim() : undefined;
    const sector = typeof req.query.sector === 'string' ? req.query.sector.trim() : undefined;

    const summary = getDashboardSummary(db, req.user!, { source, startDate, endDate, sector });
    const chains = getDepartmentsChainsBreakdown(db, req.user!, { source, startDate, endDate, sector });
    res.json({ sectors: summary.sectors, aggregations: summary.aggregations, chains });
  });

  app.get('/api/analytics/branches', requireAuth, (req: AuthenticatedRequest, res) => {
    const source = parseSource(req.query.source);
    const startDate = typeof req.query.startDate === 'string' ? req.query.startDate.trim() : undefined;
    const endDate = typeof req.query.endDate === 'string' ? req.query.endDate.trim() : undefined;
    const sector = typeof req.query.sector === 'string' ? req.query.sector.trim() : undefined;

    const result = getBranchesAnalytics(db, req.user!, { source, startDate, endDate, sector });
    if (result.diagnostics.differenceCount !== 0 || result.diagnostics.consolidatedDuplicateNameCount > 0) {
      console.log('[Branches/Reps Diagnostics]', JSON.stringify(result.diagnostics));
    }
    res.json(result);
  });

  app.get('/api/analytics/reps', requireAuth, (req: AuthenticatedRequest, res) => {
    const source = parseSource(req.query.source);
    const startDate = typeof req.query.startDate === 'string' ? req.query.startDate.trim() : undefined;
    const endDate = typeof req.query.endDate === 'string' ? req.query.endDate.trim() : undefined;
    const sector = typeof req.query.sector === 'string' ? req.query.sector.trim() : undefined;

    const reps = getRepsAnalytics(db, req.user!, { source, startDate, endDate, sector });
    const { diagnostics } = getBranchesAnalytics(db, req.user!, { source, startDate, endDate, sector });
    res.json({ reps, diagnostics });
  });

  app.get('/api/analytics/items', requireAuth, (req: AuthenticatedRequest, res) => {
    const source = parseSource(req.query.source);
    const startDate = typeof req.query.startDate === 'string' ? req.query.startDate.trim() : undefined;
    const endDate = typeof req.query.endDate === 'string' ? req.query.endDate.trim() : undefined;
    const sector = typeof req.query.sector === 'string' ? req.query.sector.trim() : undefined;
    const jobTitle = typeof req.query.jobTitle === 'string' ? req.query.jobTitle.trim() : undefined;

    const categoryAnalysis = getItemCategoriesAnalytics(db, req.user!, {
      source,
      startDate,
      endDate,
      sector,
      jobTitle,
    });
    const items = getItemsAnalytics(db, req.user!, { source, startDate, endDate, sector });
    res.json({ ...categoryAnalysis, items });
  });

  app.put('/api/analytics/item-category-targets', requireAuth, requireTargetEditor, (req: AuthenticatedRequest, res) => {
    const { targets } = req.body as { targets?: ItemCategoryTargetInput[] };
    if (!Array.isArray(targets)) {
      res.status(400).json({ error: 'صيغة أهداف الفئات غير صالحة' });
      return;
    }
    try {
      saveItemCategoryTargets(db, targets, req.user!);
      res.json({ success: true });
    } catch (err) {
      if (err instanceof AuthorizationError) {
        res.status(403).json({ error: err.message });
        return;
      }
      res.status(400).json({ error: err instanceof Error ? err.message : 'تعذر حفظ أهداف الفئات' });
    }
  });

  app.get('/api/analytics/timeline', requireAuth, (req: AuthenticatedRequest, res) => {
    const source = parseSource(req.query.source);
    const startDate = typeof req.query.startDate === 'string' ? req.query.startDate.trim() : undefined;
    const endDate = typeof req.query.endDate === 'string' ? req.query.endDate.trim() : undefined;
    const sector = typeof req.query.sector === 'string' ? req.query.sector.trim() : undefined;

    const series = getTimelineSeries(db, req.user!, { source, startDate, endDate, sector });
    res.json({ series });
  });

  app.get('/api/analytics/transactions', requireAuth, (req: AuthenticatedRequest, res) => {
    const source = parseSource(req.query.source);
    const startDate = typeof req.query.startDate === 'string' ? req.query.startDate.trim() : undefined;
    const endDate = typeof req.query.endDate === 'string' ? req.query.endDate.trim() : undefined;
    const sector = typeof req.query.sector === 'string' ? req.query.sector.trim() : undefined;
    const search = typeof req.query.search === 'string' ? req.query.search : undefined;
    const linkStatusRaw = typeof req.query.linkStatus === 'string' ? req.query.linkStatus : 'all';
    const linkStatus = ['all', 'linked', 'unlinked_branch', 'unlinked_sector', 'returns'].includes(linkStatusRaw)
      ? (linkStatusRaw as 'all' | 'linked' | 'unlinked_branch' | 'unlinked_sector' | 'returns')
      : 'all';
    const page = Number(req.query.page || 1);
    const pageSize = Number(req.query.pageSize || 50);

    const result = getSalesTransactions(db, req.user!, {
      source,
      startDate,
      endDate,
      sector,
      search,
      linkStatus,
      page,
      pageSize,
    });
    res.json(result);
  });

  // ==================== COMPETITIONS ROUTES ====================
  app.get('/api/competitions', requireAuth, (_req, res) => {
    const competitions = listCompetitions(db);
    const catalogItems = getAvailableCatalogItems(db);
    res.json({ competitions, catalogItems });
  });

  app.get('/api/competitions/analytics', requireAuth, (req: AuthenticatedRequest, res) => {
    const compIdRaw = req.query.competitionId ? Number(req.query.competitionId) : null;
    const compId = compIdRaw && Number.isInteger(compIdRaw) && compIdRaw > 0 ? compIdRaw : null;
    const sector = typeof req.query.sector === 'string' ? req.query.sector.trim() : undefined;
    const startDate = typeof req.query.startDate === 'string' ? req.query.startDate.trim() : undefined;
    const endDate = typeof req.query.endDate === 'string' ? req.query.endDate.trim() : undefined;

    const data = getCompetitionAnalytics(db, req.user!, compId, { sector, startDate, endDate });
    res.json(data);
  });

  app.get('/api/competitions/:id/analytics', requireAuth, (req: AuthenticatedRequest, res) => {
    const compIdRaw = Number(req.params.id);
    const compId = Number.isInteger(compIdRaw) && compIdRaw > 0 ? compIdRaw : null;
    const sector = typeof req.query.sector === 'string' ? req.query.sector.trim() : undefined;
    const startDate = typeof req.query.startDate === 'string' ? req.query.startDate.trim() : undefined;
    const endDate = typeof req.query.endDate === 'string' ? req.query.endDate.trim() : undefined;

    const data = getCompetitionAnalytics(db, req.user!, compId, { sector, startDate, endDate });
    res.json(data);
  });

  app.post('/api/competitions', requireAuth, requireAdmin, (req: AuthenticatedRequest, res) => {
    const { name, items } = req.body as {
      name?: string;
      items?: CompetitionItemInput[];
    };
    try {
      const created = createCompetition(db, { id: req.user!.id }, { name: String(name || ''), items });
      res.status(201).json({ competition: created });
    } catch (err) {
      res.status(400).json({ error: err instanceof Error ? err.message : 'تعذر إنشاء المسابقة' });
    }
  });

  app.put('/api/competitions/:id', requireAuth, requireAdmin, (req, res) => {
    const compId = Number(req.params.id);
    if (!Number.isInteger(compId) || compId <= 0) {
      res.status(400).json({ error: 'معرف المسابقة غير صالح' });
      return;
    }
    const { name, isActive, items } = req.body as {
      name?: string;
      isActive?: boolean;
      items?: CompetitionItemInput[];
    };
    try {
      const updated = updateCompetition(db, compId, { name, isActive, items });
      res.json({ competition: updated });
    } catch (err) {
      res.status(400).json({ error: err instanceof Error ? err.message : 'تعذر تحديث المسابقة' });
    }
  });

  app.delete('/api/competitions/:id', requireAuth, requireAdmin, (req, res) => {
    const compId = Number(req.params.id);
    if (!Number.isInteger(compId) || compId <= 0) {
      res.status(400).json({ error: 'معرف المسابقة غير صالح' });
      return;
    }
    try {
      deleteCompetition(db, compId);
      res.json({ success: true });
    } catch (err) {
      res.status(404).json({ error: err instanceof Error ? err.message : 'تعذر حذف المسابقة' });
    }
  });

  app.put('/api/competitions/:id/targets', requireAuth, requireTargetEditor, (req: AuthenticatedRequest, res) => {
    const compId = Number(req.params.id);
    if (!Number.isInteger(compId) || compId <= 0) {
      res.status(400).json({ error: 'معرف المسابقة غير صالح' });
      return;
    }
    const { targets } = req.body as {
      targets?: CompetitionTargetInput[];
    };
    if (!Array.isArray(targets)) {
      res.status(400).json({ error: 'صيغة الأهداف غير صالحة' });
      return;
    }
    try {
      saveCompetitionTargets(db, compId, targets, req.user!);
      res.json({ success: true });
    } catch (err) {
      if (err instanceof AuthorizationError) {
        res.status(403).json({ error: err.message });
        return;
      }
      res.status(400).json({ error: err instanceof Error ? err.message : 'تعذر حفظ الأهداف الكمية' });
    }
  });

  // ==================== EXCEL IMPORT & VERSIONING ROUTES ====================
  app.post('/api/import/validate', requireAuth, requireAdmin, upload.single('file'), (req: AuthenticatedRequest, res) => {
    if (!req.file || !req.file.buffer || req.file.buffer.length === 0) {
      res.status(400).json({ error: 'لم يتم إرفاق ملف Excel صالح' });
      return;
    }
    try {
      const decodedFilename = Buffer.from(req.file.originalname, 'latin1').toString('utf8');
      const filename = /[\u0600-\u06FF]/.test(decodedFilename) ? decodedFilename : req.file.originalname;
      const report = validateExcelBuffer(req.file.buffer, filename);
      res.json(report);
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'تعذر تحليل ملف Excel';
      res.status(400).json({ error: msg });
    }
  });

  app.post('/api/import/commit', requireAuth, requireAdmin, (req: AuthenticatedRequest, res) => {
    const { validationToken } = req.body as { validationToken?: string };
    if (!validationToken) {
      res.status(400).json({ error: 'رمز التحقق مطلوب لاعتماد الاستيراد' });
      return;
    }
    try {
      const batch = commitValidatedImport(db, validationToken, {
        id: req.user!.id,
        fullName: req.user!.fullName,
      });
      res.json({ success: true, batch });
    } catch (err) {
      const message = err instanceof Error ? err.message : 'فشل اعتماد البيانات';
      res.status(400).json({ error: message });
    }
  });

  app.get('/api/import/history', requireAuth, requireAdmin, (_req, res) => {
    const rows = db
      .prepare(`SELECT * FROM import_batches ORDER BY id DESC`)
      .all() as Array<{
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

    res.json({ batches });
  });

  app.post('/api/import/rollback/:batchId', requireAuth, requireAdmin, (req, res) => {
    const batchId = Number(req.params.batchId);
    if (!Number.isInteger(batchId) || batchId <= 0) {
      res.status(400).json({ error: 'معرف النسخة غير صالح' });
      return;
    }
    try {
      const restored = rollbackToBatch(db, batchId);
      res.json({ success: true, batch: restored });
    } catch (err) {
      res.status(404).json({ error: err instanceof Error ? err.message : 'تعذر استرجاع النسخة' });
    }
  });

  app.get('/api/import/template', requireAuth, requireAdmin, (_req, res) => {
    const buffer = Buffer.from(generateStandardTemplateWorkbook());
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="FMCG_Sales_Template.xlsx"');
    res.send(buffer);
  });

  // ==================== USER MANAGEMENT ROUTES ====================
  app.get('/api/users', requireAuth, requireAdmin, (_req, res) => {
    const rows = db.prepare('SELECT * FROM users ORDER BY id ASC').all() as DbUserRow[];
    res.json({ users: rows.map(mapDbUserToProfile) });
  });

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

  app.post('/api/users', requireAuth, requireAdmin, (req, res) => {
    const { username, fullName, email, role, allowedSectors, password, mustChangePassword } = req.body as {
      username?: string;
      fullName?: string;
      email?: string;
      role?: string;
      allowedSectors?: string[];
      password?: string;
      mustChangePassword?: boolean;
    };

    const cleanUsername = String(username || '').trim().toLowerCase();
    const cleanName = String(fullName || '').trim();
    const cleanPass = String(password || '');
    const cleanRole = parseStrictUserRole(role ?? 'standard_user');
    const forceFirstChange = typeof mustChangePassword === 'boolean' ? mustChangePassword : true;

    if (!cleanRole) {
      res.status(400).json({ error: 'الدور الوظيفي المحدد غير صالح' });
      return;
    }
    if (!cleanUsername || cleanUsername.length < 3) {
      res.status(400).json({ error: 'اسم المستخدم يجب أن يكون 3 أحرف على الأقل' });
      return;
    }
    if (!cleanName) {
      res.status(400).json({ error: 'يرجى إدخال الاسم الكامل' });
      return;
    }
    if (cleanPass.length < 8) {
      res.status(400).json({ error: 'كلمة المرور المبدئية يجب أن لا تقل عن 8 أحرف' });
      return;
    }

    const exists = db.prepare('SELECT id FROM users WHERE lower(username) = ?').get(cleanUsername);
    if (exists) {
      res.status(409).json({ error: 'اسم المستخدم مسجل بالفعل' });
      return;
    }

    const sectors = cleanRole === 'admin' ? [...ORIGINAL_SECTORS] : sanitizeSectorsInput(allowedSectors);
    if (cleanRole !== 'admin' && sectors.length === 0) {
      res.status(400).json({ error: 'يرجى تحديد قطاع واحد على الأقل لهذا المستخدم' });
      return;
    }

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
        hashPassword(cleanPass),
        forceFirstChange ? 1 : 0,
        now
      );

    const created = db.prepare('SELECT * FROM users WHERE id = ?').get(Number(resInsert.lastInsertRowid)) as DbUserRow;
    res.status(201).json({ user: mapDbUserToProfile(created) });
  });

  app.post('/api/users/:id/reset-password', requireAuth, requireAdmin, (req, res) => {
    const targetId = Number(req.params.id);
    const { temporaryPassword } = req.body as { temporaryPassword?: string };
    const cleanPass = String(temporaryPassword || '');
    if (cleanPass.length < 8) {
      res.status(400).json({ error: 'كلمة المرور المبدئية يجب أن لا تقل عن 8 أحرف' });
      return;
    }
    const existing = db.prepare('SELECT * FROM users WHERE id = ?').get(targetId) as DbUserRow | undefined;
    if (!existing) {
      res.status(404).json({ error: 'المستخدم غير موجود' });
      return;
    }
    db.prepare('UPDATE users SET password_hash = ?, must_change_password = 1 WHERE id = ?').run(
      hashPassword(cleanPass),
      targetId
    );
    const updated = db.prepare('SELECT * FROM users WHERE id = ?').get(targetId) as DbUserRow;
    res.json({ user: mapDbUserToProfile(updated) });
  });

  app.patch('/api/users/:id', requireAuth, requireAdmin, (req: AuthenticatedRequest, res) => {
    const targetId = Number(req.params.id);
    const existing = db.prepare('SELECT * FROM users WHERE id = ?').get(targetId) as DbUserRow | undefined;
    if (!existing) {
      res.status(404).json({ error: 'المستخدم غير موجود' });
      return;
    }

    const { fullName, email, role, allowedSectors, isActive, newPassword, mustChangePassword } = req.body as {
      fullName?: string;
      email?: string | null;
      role?: string;
      allowedSectors?: string[];
      isActive?: boolean;
      newPassword?: string;
      mustChangePassword?: boolean;
    };

    let nextRole: UserRole =
      existing.role === 'admin'
        ? 'admin'
        : existing.role === 'sector_manager'
          ? 'sector_manager'
          : 'standard_user';

    if (role !== undefined) {
      const parsed = parseStrictUserRole(role);
      if (!parsed) {
        res.status(400).json({ error: 'الدور الوظيفي المحدد غير صالح' });
        return;
      }
      nextRole = parsed;
    }

    const nextActive = typeof isActive === 'boolean' ? (isActive ? 1 : 0) : existing.is_active;

    if (targetId === req.user!.id && (nextRole !== 'admin' || nextActive === 0)) {
      res.status(400).json({ error: 'لا يمكنك تعطيل حسابك أو خفض صلاحيتك أثناء تسجيل الدخول' });
      return;
    }

    let nextSectorsJson = existing.allowed_sectors;
    if (nextRole === 'admin') {
      nextSectorsJson = JSON.stringify([...ORIGINAL_SECTORS]);
    } else if (allowedSectors !== undefined) {
      const cleaned = sanitizeSectorsInput(allowedSectors);
      if (cleaned.length === 0) {
        res.status(400).json({ error: 'يرجى تحديد قطاع واحد على الأقل لهذا المستخدم' });
        return;
      }
      nextSectorsJson = JSON.stringify(cleaned);
    } else {
      // Ensure existing sectors are non-empty when demoting from admin without explicit sector list
      let existingList: string[] = [];
      try {
        const parsed = JSON.parse(existing.allowed_sectors);
        if (Array.isArray(parsed)) existingList = sanitizeSectorsInput(parsed);
      } catch {
        existingList = [];
      }
      if (existingList.length === 0) {
        res.status(400).json({ error: 'يرجى تحديد قطاع واحد على الأقل لهذا المستخدم' });
        return;
      }
      nextSectorsJson = JSON.stringify(existingList);
    }

    if (newPassword && newPassword.length > 0 && newPassword.length < 8) {
      res.status(400).json({ error: 'كلمة المرور الجديدة يجب أن تكون 8 أحرف على الأقل' });
      return;
    }
    const hasNewPass = Boolean(newPassword && newPassword.length >= 8);
    const nextHash = hasNewPass ? hashPassword(newPassword!) : existing.password_hash;
    const nextMustChange =
      typeof mustChangePassword === 'boolean'
        ? mustChangePassword
          ? 1
          : 0
        : hasNewPass
          ? 0
          : existing.must_change_password;

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

    const updated = db.prepare('SELECT * FROM users WHERE id = ?').get(targetId) as DbUserRow;
    res.json({ user: mapDbUserToProfile(updated) });
  });

  app.delete('/api/users/:id', requireAuth, requireAdmin, (req: AuthenticatedRequest, res) => {
    const targetId = Number(req.params.id);
    if (targetId === req.user!.id) {
      res.status(400).json({ error: 'لا يمكنك حذف حسابك النشط' });
      return;
    }
    db.prepare('DELETE FROM users WHERE id = ?').run(targetId);
    res.json({ success: true });
  });

  app.get('/api/admin/supabase-status', requireAuth, requireAdmin, (_req, res) => {
    const url = String(process.env.VITE_SUPABASE_URL || '').trim();
    const anonKey = String(process.env.VITE_SUPABASE_ANON_KEY || '').trim();
    const configured = Boolean(url && anonKey && /^https?:\/\//i.test(url));
    res.json({
      configured,
      ok: configured,
      userCount: 0,
      message: configured
        ? 'تم العثور على متغيرات VITE_SUPABASE_URL و VITE_SUPABASE_ANON_KEY.'
        : 'يعمل التطبيق حالياً على قاعدة SQLite المحلية. للربط السحابي المشترك على Vercel، أضف متغيرات VITE_SUPABASE_URL و VITE_SUPABASE_ANON_KEY ونفذ كود SQL المرفق.',
    });
  });

  // ==================== SMART AUDITOR & RUNTIME ERROR MONITORING ROUTES ====================
  const ALLOWED_AUTOFIX_ACTIONS = new Set<AutoFixActionType>([
    'FIX_NON_FINITE_NUMBERS',
    'FIX_DUPLICATE_BRANCH_RECORDS',
    'FIX_DUPLICATE_ITEM_CATALOG',
    'FIX_SYNC_COMPETITION_TARGET_TOTALS',
    'FIX_CANONICALIZE_TARGETS_AND_SCOPES',
    'FIX_RESOLVE_RUNTIME_ERRORS',
  ]);

  // Any authenticated user's client runtime errors are recorded silently for Admin inspection
  app.post('/api/audit/runtime-errors', requireAuth, (req: AuthenticatedRequest, res) => {
    const body = (req.body || {}) as {
      kind?: string;
      tabName?: string;
      message?: string;
      stack?: string;
      sourceUrl?: string;
      lineNo?: number;
      colNo?: number;
    };
    const msg = String(body.message || '').trim();
    if (!msg) {
      res.status(400).json({ error: 'رسالة الخطأ مطلوبة' });
      return;
    }
    const log = recordRuntimeError(
      db,
      req.user ? { id: req.user.id, username: req.user.username } : undefined,
      body
    );
    res.status(201).json({ recorded: true, id: log.id });
  });

  app.get('/api/admin/audit/latest', requireAuth, requireAdmin, (req: AuthenticatedRequest, res) => {
    const report = runSmartAudit(db, req.user!);
    res.json({ report });
  });

  app.post('/api/admin/audit/run', requireAuth, requireAdmin, (req: AuthenticatedRequest, res) => {
    const report = runSmartAudit(db, req.user!);
    res.json({ report });
  });

  app.post('/api/admin/audit/ai-analyze', requireAuth, requireAdmin, async (req: AuthenticatedRequest, res) => {
    try {
      const result = await enrichAuditWithGemini(db, req.user!);
      res.json(result);
    } catch (err) {
      res.status(502).json({
        error: err instanceof Error ? err.message : 'تعذر إجراء التحليل الإضافي عبر الذكاء الاصطناعي',
      });
    }
  });

  app.post('/api/admin/audit/autofix', requireAuth, requireAdmin, (req: AuthenticatedRequest, res) => {
    const { actionType, findingCode } = req.body as { actionType?: string; findingCode?: string };
    if (!actionType || !ALLOWED_AUTOFIX_ACTIONS.has(actionType as AutoFixActionType)) {
      res.status(400).json({ error: 'نوع الإصلاح التلقائي غير مدعوم' });
      return;
    }
    try {
      const result = executeAutoFix(db, req.user!, actionType as AutoFixActionType, findingCode);
      res.json(result);
    } catch (err) {
      res.status(400).json({ error: err instanceof Error ? err.message : 'تعذر تنفيذ الإصلاح التلقائي' });
    }
  });

  app.post('/api/admin/audit/rollback/:backupId', requireAuth, requireAdmin, (req: AuthenticatedRequest, res) => {
    const backupId = Number(req.params.backupId);
    if (!Number.isInteger(backupId) || backupId <= 0) {
      res.status(400).json({ error: 'معرف النسخة الاحتياطية غير صالح' });
      return;
    }
    try {
      const result = rollbackAutoFix(db, req.user!, backupId);
      res.json(result);
    } catch (err) {
      res.status(400).json({ error: err instanceof Error ? err.message : 'تعذر استعادة النسخة الاحتياطية' });
    }
  });

  // ==================== EXPORT PROJECT ZIP FOR WINDOWS 10 ====================
  app.get('/api/admin/export-project-zip', requireAuth, requireAdmin, async (_req, res) => {
    try {
      const zip = new JSZip();
      const rootDir = __dirname;
      const includeRoots = [
        'package.json',
        'tsconfig.json',
        'vite.config.ts',
        'index.html',
        'metadata.json',
        'vercel.json',
        '.env.example',
        'server.ts',
        'src',
        'server',
        'supabase',
        'tests',
      ];

      const addPathToZip = (relPath: string) => {
        const abs = path.resolve(rootDir, relPath);
        if (!abs.startsWith(rootDir + path.sep) || !fs.existsSync(abs)) return;
        const stat = fs.statSync(abs);
        if (stat.isDirectory()) {
          for (const entry of fs.readdirSync(abs)) {
            if (entry === 'node_modules' || entry === 'dist' || entry === '.git') continue;
            addPathToZip(path.join(relPath, entry));
          }
        } else {
          zip.file(relPath.replace(/\\/g, '/'), fs.readFileSync(abs));
        }
      };

      for (const item of includeRoots) {
        addPathToZip(item);
      }

      zip.file(
        'README_WINDOWS.txt',
        [
          'Sales Performance Dashboard - Windows 10 Setup Guide',
          '====================================================',
          '1. Install Node.js v20+ or v22+ LTS from https://nodejs.org',
          '2. Open Command Prompt (CMD) or PowerShell inside this folder',
          '3. Run: npm install',
          '4. Run: npm run build',
          '5. Run: npm start',
          '6. Open browser at: http://localhost:3000',
          '',
          'Run calculation verification tests anytime via: npm test',
        ].join('\r\n')
      );

      const zipBuf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
      res.setHeader('Content-Type', 'application/zip');
      res.setHeader('Content-Disposition', 'attachment; filename="sales-performance-dashboard.zip"');
      res.send(zipBuf);
    } catch (err) {
      res.status(500).json({ error: err instanceof Error ? err.message : 'تعذر إنشاء ملف ZIP' });
    }
  });

  // Error handler for multer & API routes
  app.use((err: unknown, _req: express.Request, res: express.Response, _next: express.NextFunction) => {
    const msg = err instanceof Error ? err.message : 'حدث خطأ غير متوقع في الخادم';
    res.status(400).json({ error: msg });
  });

  // ==================== VITE / STATIC SERVING ====================
  if (!isProd) {
    let viteMiddleware: express.RequestHandler | null = null;
    const hmrDisabled = process.env.DISABLE_HMR === 'true';
    const viteReady = createViteServer({
      server: {
        middlewareMode: true,
        hmr: hmrDisabled ? false : true,
        ...(hmrDisabled ? {} : { ws: { server: httpServer } }),
      },
      appType: 'spa',
    })
      .then((vite) => {
        viteMiddleware = vite.middlewares;
        return vite;
      })
      .catch((err) => {
        console.error('Failed to initialize Vite dev middleware:', err);
        throw err;
      });

    app.use(async (req, res, next) => {
      try {
        if (!viteMiddleware) {
          await viteReady;
        }
        viteMiddleware!(req, res, next);
      } catch (err) {
        next(err);
      }
    });
  } else {
    const distDir = path.resolve(__dirname, 'dist');
    app.use(express.static(distDir));
    app.get('*', (_req, res) => {
      res.sendFile(path.join(distDir, 'index.html'));
    });
  }

  httpServer.listen(PORT, '0.0.0.0', () => {
    console.log(`Sales Performance BI Server running at http://localhost:${PORT}`);
  });
}

startServer();
