import type Database from 'better-sqlite3';
import type {
  AuditFinding,
  AuditFixBackup,
  AutoFixActionType,
  RuntimeErrorLog,
  SmartAuditReport,
  UserProfile,
} from '../src/types/bi.ts';
import { AuthorizationError } from '../src/utils/errors.ts';
import {
  buildItemCategoryTargetKey,
  getBranchesAnalytics,
  getDashboardSummary,
  getItemCategoriesAnalytics,
  getRepsAnalytics,
  getTimelineSeries,
  saveItemCategoryTargets,
} from './analytics-engine.ts';
import { getCompetitionAnalytics, listCompetitions, saveCompetitionTargets } from './competitions-engine.ts';
import {
  canonicalizeSectorName,
  normalizeArabic,
  normalizeCode,
  ORIGINAL_SECTORS,
} from './excel-engine.ts';

function isBadNumber(val: unknown): boolean {
  if (val === null || val === undefined) return false;
  if (typeof val === 'number') return !Number.isFinite(val);
  if (typeof val === 'string') {
    const s = val.trim().toLowerCase();
    return s === 'nan' || s === 'infinity' || s === '-infinity';
  }
  return false;
}

function scanObjectForNonFinite(obj: unknown, pathPrefix: string, out: string[], limit = 15): void {
  if (out.length >= limit || obj === null || obj === undefined) return;
  if (typeof obj === 'number') {
    if (!Number.isFinite(obj)) {
      out.push(`${pathPrefix} = ${String(obj)}`);
    }
    return;
  }
  if (typeof obj === 'string') {
    const low = obj.trim().toLowerCase();
    if (low === 'nan' || low === 'infinity' || low === '-infinity') {
      out.push(`${pathPrefix} = "${obj}"`);
    }
    return;
  }
  if (Array.isArray(obj)) {
    for (let i = 0; i < obj.length && out.length < limit; i++) {
      scanObjectForNonFinite(obj[i], `${pathPrefix}[${i}]`, out, limit);
    }
    return;
  }
  if (typeof obj === 'object') {
    for (const [k, v] of Object.entries(obj as Record<string, unknown>)) {
      if (out.length >= limit) break;
      scanObjectForNonFinite(v, `${pathPrefix}.${k}`, out, limit);
    }
  }
}

export interface RecordRuntimeErrorInput {
  kind?: string;
  tabName?: string;
  message?: string;
  stack?: string;
  sourceUrl?: string;
  lineNo?: number;
  colNo?: number;
}

export function recordRuntimeError(
  db: Database.Database,
  user: { id: number; username: string } | undefined,
  input: RecordRuntimeErrorInput
): RuntimeErrorLog {
  const kindRaw = String(input.kind || 'runtime_error').trim();
  const kind: RuntimeErrorLog['kind'] =
    kindRaw === 'unhandled_rejection' || kindRaw === 'react_error' ? kindRaw : 'runtime_error';
  const tabName = String(input.tabName || 'غير محدد').trim().slice(0, 120);
  const message = String(input.message || 'خطأ برمجي غير معروف').trim().slice(0, 1000);
  const stack = String(input.stack || '').trim().slice(0, 4000);
  const sourceUrl = String(input.sourceUrl || '').trim().slice(0, 500);
  const lineNo = Number.isFinite(Number(input.lineNo)) ? Math.max(0, Math.floor(Number(input.lineNo))) : 0;
  const colNo = Number.isFinite(Number(input.colNo)) ? Math.max(0, Math.floor(Number(input.colNo))) : 0;
  const now = new Date().toISOString();

  const combinedErr = `${message}\n${stack}\n${sourceUrl}`.toLowerCase();
  if (
    combinedErr.includes('websocket closed without opened') ||
    (combinedErr.includes('@vite/client') && combinedErr.includes('websocket'))
  ) {
    return {
      id: 0,
      kind,
      tabName,
      message,
      stack,
      sourceUrl,
      lineNo,
      colNo,
      username: user?.username ?? '',
      resolved: true,
      aiAnalysis: '',
      occurredAt: now,
    };
  }

  // Deduplicate identical unresolved errors occurring within 60 seconds
  const recent = db
    .prepare(
      `SELECT id, occurred_at FROM runtime_error_logs
       WHERE resolved = 0 AND kind = ? AND tab_name = ? AND message = ?
       ORDER BY id DESC LIMIT 1`
    )
    .get(kind, tabName, message) as { id: number; occurred_at: string } | undefined;

  if (recent) {
    const diffMs = Date.now() - Date.parse(recent.occurred_at);
    if (Number.isFinite(diffMs) && diffMs >= 0 && diffMs < 60_000) {
      db.prepare(`UPDATE runtime_error_logs SET occurred_at = ?, stack = COALESCE(NULLIF(?, ''), stack) WHERE id = ?`).run(
        now,
        stack,
        recent.id
      );
      const updated = db.prepare(`SELECT * FROM runtime_error_logs WHERE id = ?`).get(recent.id) as {
        id: number;
        kind: RuntimeErrorLog['kind'];
        tab_name: string;
        message: string;
        stack: string;
        source_url: string;
        line_no: number;
        col_no: number;
        username: string;
        resolved: number;
        ai_analysis: string;
        occurred_at: string;
      };
      return mapRuntimeErrorRow(updated);
    }
  }

  const res = db
    .prepare(
      `INSERT INTO runtime_error_logs (
        kind, tab_name, message, stack, source_url, line_no, col_no, user_id, username, resolved, ai_analysis, occurred_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0, '', ?)`
    )
    .run(kind, tabName, message, stack, sourceUrl, lineNo, colNo, user?.id ?? null, user?.username ?? '', now);

  const row = db.prepare(`SELECT * FROM runtime_error_logs WHERE id = ?`).get(Number(res.lastInsertRowid)) as {
    id: number;
    kind: RuntimeErrorLog['kind'];
    tab_name: string;
    message: string;
    stack: string;
    source_url: string;
    line_no: number;
    col_no: number;
    username: string;
    resolved: number;
    ai_analysis: string;
    occurred_at: string;
  };
  return mapRuntimeErrorRow(row);
}

function mapRuntimeErrorRow(r: {
  id: number;
  kind: string;
  tab_name: string;
  message: string;
  stack: string;
  source_url: string;
  line_no: number;
  col_no: number;
  username: string;
  resolved: number;
  ai_analysis: string;
  occurred_at: string;
}): RuntimeErrorLog {
  const kind: RuntimeErrorLog['kind'] =
    r.kind === 'unhandled_rejection' || r.kind === 'react_error' ? r.kind : 'runtime_error';
  return {
    id: r.id,
    kind,
    tabName: r.tab_name,
    message: r.message,
    stack: r.stack,
    sourceUrl: r.source_url,
    lineNo: r.line_no,
    colNo: r.col_no,
    username: r.username,
    resolved: r.resolved === 1,
    aiAnalysis: r.ai_analysis || '',
    occurredAt: r.occurred_at,
  };
}

export function listRuntimeErrorLogs(db: Database.Database, includeResolved = false): RuntimeErrorLog[] {
  db.prepare(
    `UPDATE runtime_error_logs
     SET resolved = 1
     WHERE resolved = 0
       AND (lower(message) LIKE '%websocket closed without opened%'
         OR (lower(stack) LIKE '%@vite/client%' AND lower(message) LIKE '%websocket%'))`
  ).run();
  const sql = includeResolved
    ? `SELECT * FROM runtime_error_logs ORDER BY id DESC LIMIT 100`
    : `SELECT * FROM runtime_error_logs WHERE resolved = 0 ORDER BY id DESC LIMIT 100`;
  const rows = db.prepare(sql).all() as Array<Parameters<typeof mapRuntimeErrorRow>[0]>;
  return rows.map(mapRuntimeErrorRow);
}

export function listAuditBackups(db: Database.Database): AuditFixBackup[] {
  const rows = db
    .prepare(
      `SELECT id, finding_code, action_type, action_label, summary, created_by_name, created_at, restored_at
       FROM audit_fix_backups
       ORDER BY id DESC
       LIMIT 30`
    )
    .all() as Array<{
    id: number;
    finding_code: string;
    action_type: string;
    action_label: string;
    summary: string;
    created_by_name: string;
    created_at: string;
    restored_at: string | null;
  }>;
  return rows.map((r) => ({
    id: r.id,
    findingCode: r.finding_code,
    actionType: r.action_type,
    actionLabel: r.action_label,
    summary: r.summary,
    createdByName: r.created_by_name,
    createdAt: r.created_at,
    restoredAt: r.restored_at,
  }));
}

function buildLocalRuntimeErrorPrompt(err: RuntimeErrorLog): string {
  const hint = (() => {
    const m = err.message.toLowerCase();
    if (m.includes('maximum update depth')) {
      return 'السبب المرجح: استدعاء setState داخل useEffect مع وجود مصفوفة أو كائن يتغير مرجعه في كل render (مثل `?? []` أو props غير مستقرة). استخدم ثوابت فارغة على مستوى الموديول أو useMemo وتحقّق من تساوي القيم قبل تحديث الحالة.';
    }
    if (m.includes('cannot read properties of undefined') || m.includes('cannot read properties of null')) {
      return 'السبب المرجح: الوصول إلى خاصية داخل كائن أو مصفوفة قد تكون undefined/null أثناء التحميل أو عند غياب بيانات الدفعة. أضف Optional Chaining (`?.`) وقيمة افتراضية آمنة.';
    }
    if (m.includes('nan') || m.includes('infinity')) {
      return 'السبب المرجح: قسمة على صفر أو تحويل نص غير رقمي عبر Number(). استخدم دالة حماية القسمة (target > 0) وتحقّق من Number.isFinite().';
    }
    if (m.includes('401') || m.includes('403') || m.includes('unauthorized') || m.includes('forbidden')) {
      return 'السبب المرجح: انتهاء الجلسة أو عدم تطابق صلاحية الدور / القطاع المسموح في مسار API.';
    }
    return 'راجع مسار التنفيذ الموضح في Stack Trace وتأكد من معالجة الحالات الحدية والقيم الفارغة بأمان.';
  })();

  const stackExcerpt = err.stack
    ? err.stack.split('\n').slice(0, 6).join('\n')
    : `${err.sourceUrl || 'غير متوفر'}:${err.lineNo}:${err.colNo}`;

  return [
    `أصلح الخطأ البرمجي التالي في التطبيق بدقة ودون المساس بباقي قواعد الحساب أو الصلاحيات:`,
    `- التبويب: ${err.tabName || 'غير محدد'}`,
    `- النوع: ${err.kind}`,
    `- رسالة الخطأ: ${err.message}`,
    `- الموقع / المكدس:\n${stackExcerpt}`,
    `- التشخيص المبدئي: ${hint}`,
    `- شرط الأمان: حافظ على الفصل التام بين مصدري VF05N وZsales_Price، وعلى صلاحيات الأدوار الثلاثة (Admin / Sector Manager / Standard User)، وعلى عدم ظهور قيم NaN أو Infinity.`,
  ].join('\n');
}

export function runSmartAudit(
  db: Database.Database,
  adminUser: UserProfile,
  options?: { auditAdminBypassNote?: boolean; aiConfigured?: boolean }
): SmartAuditReport {
  void options;
  const now = new Date().toISOString();
  const findings: AuditFinding[] = [];

  const activeBatch = db
    .prepare(`SELECT id, filename, unlinked_branch_count, unlinked_sector_count FROM import_batches WHERE status = 'active' ORDER BY id DESC LIMIT 1`)
    .get() as
    | { id: number; filename: string; unlinked_branch_count: number; unlinked_sector_count: number }
    | undefined;

  // ==================== CHECK 1: NUMERIC INTEGRITY (NaN / Infinity / Negative Targets) ====================
  const badDbNumericDetails: string[] = [];
  if (activeBatch) {
    const badSalesRows = db
      .prepare(
        `SELECT id, source_type, quantity, sales_value, tax_value
         FROM sales_records
         WHERE batch_id = ?`
      )
      .all(activeBatch.id) as Array<{
      id: number;
      source_type: string;
      quantity: unknown;
      sales_value: unknown;
      tax_value: unknown;
    }>;
    for (const r of badSalesRows) {
      if (isBadNumber(r.quantity) || isBadNumber(r.sales_value) || isBadNumber(r.tax_value)) {
        badDbNumericDetails.push(`sales_records#${r.id} (${r.source_type}): qty=${r.quantity}, sales=${r.sales_value}, tax=${r.tax_value}`);
        if (badDbNumericDetails.length >= 20) break;
      }
    }

    const badBranches = db
      .prepare(`SELECT id, branch_code, branch_target FROM branches WHERE batch_id = ?`)
      .all(activeBatch.id) as Array<{ id: number; branch_code: string; branch_target: unknown }>;
    for (const b of badBranches) {
      if (isBadNumber(b.branch_target) || Number(b.branch_target) < 0) {
        badDbNumericDetails.push(`branches#${b.id} (${b.branch_code}): target=${b.branch_target}`);
      }
    }
  }

  const badCatTargets = db
    .prepare(`SELECT id, target_key, target_value FROM item_category_targets`)
    .all() as Array<{ id: number; target_key: string; target_value: unknown }>;
  for (const ct of badCatTargets) {
    if (isBadNumber(ct.target_value) || Number(ct.target_value) < 0) {
      badDbNumericDetails.push(`item_category_targets#${ct.id} (${ct.target_key}): target=${ct.target_value}`);
    }
  }

  const badCompTargets = db
    .prepare(`SELECT id, participant_key, target_qty, industrial_target_qty, commercial_target_qty FROM competition_targets`)
    .all() as Array<{
    id: number;
    participant_key: string;
    target_qty: unknown;
    industrial_target_qty: unknown;
    commercial_target_qty: unknown;
  }>;
  for (const cp of badCompTargets) {
    if (
      isBadNumber(cp.target_qty) ||
      isBadNumber(cp.industrial_target_qty) ||
      isBadNumber(cp.commercial_target_qty) ||
      Number(cp.target_qty) < 0 ||
      Number(cp.industrial_target_qty) < 0 ||
      Number(cp.commercial_target_qty) < 0
    ) {
      badDbNumericDetails.push(
        `competition_targets#${cp.id} (${cp.participant_key}): total=${cp.target_qty}, ind=${cp.industrial_target_qty}, com=${cp.commercial_target_qty}`
      );
    }
  }

  const badCompItems = db
    .prepare(`SELECT id, item_code, equivalent_factor FROM competition_items`)
    .all() as Array<{ id: number; item_code: string; equivalent_factor: unknown }>;
  for (const ci of badCompItems) {
    if (isBadNumber(ci.equivalent_factor) || Number(ci.equivalent_factor) < 0) {
      badDbNumericDetails.push(`competition_items#${ci.id} (${ci.item_code}): factor=${ci.equivalent_factor}`);
    }
  }

  if (badDbNumericDetails.length > 0) {
    findings.push({
      id: 'NUM_DB_NON_FINITE',
      code: 'NUM_DB_NON_FINITE',
      title: 'قيم رقمية غير صالحة أو سالبة (NaN / Infinity) في قاعدة البيانات',
      description: `تم رصد ${badDbNumericDetails.length} سجل يحتوي على قيم NaN أو Infinity أو أهداف/معاملات سالبة قد تشوه الحسابات.`,
      severity: 'critical',
      category: 'numeric_integrity',
      fixType: 'auto_fixable',
      affectedArea: 'جداول المبيعات والأهداف',
      details: badDbNumericDetails.slice(0, 10),
      autoFixAction: 'FIX_NON_FINITE_NUMBERS',
      autoFixLabel: 'تنظيف وتصحيح القيم الرقمية تلقائياً (مع نسخة احتياطية)',
    });
  }

  // Scan live computed API outputs for NaN / Infinity across VF05N and ZSALES
  const computedBadValues: string[] = [];
  const allComps = listCompetitions(db);
  if (activeBatch) {
    for (const src of ['VF05N', 'ZSALES'] as const) {
      scanObjectForNonFinite(getDashboardSummary(db, adminUser, { source: src, sector: 'ALL' }), `dashboard(${src})`, computedBadValues);
      scanObjectForNonFinite(getBranchesAnalytics(db, adminUser, { source: src, sector: 'ALL' }), `branches(${src})`, computedBadValues);
      scanObjectForNonFinite(getRepsAnalytics(db, adminUser, { source: src, sector: 'ALL' }), `reps(${src})`, computedBadValues);
      scanObjectForNonFinite(getItemCategoriesAnalytics(db, adminUser, { source: src, sector: 'ALL' }), `items(${src})`, computedBadValues);
      scanObjectForNonFinite(getTimelineSeries(db, adminUser, { source: src, sector: 'ALL' }), `timeline(${src})`, computedBadValues);
    }
    for (const comp of allComps) {
      scanObjectForNonFinite(
        getCompetitionAnalytics(db, adminUser, comp.id, { sector: 'ALL' }),
        `competition#${comp.id}`,
        computedBadValues
      );
    }
  }

  if (computedBadValues.length > 0) {
    findings.push({
      id: 'NUM_COMPUTED_NON_FINITE',
      code: 'NUM_COMPUTED_NON_FINITE',
      title: 'ظهور قيم NaN أو Infinity في مخرجات الحسابات التحليلية',
      description: 'رصد الفاحص قيمًا غير محدودة في المخرجات المحسوبة للتقارير (مثل نسب التحقيق أو الأهداف).',
      severity: 'critical',
      category: 'numeric_integrity',
      fixType: 'code_fix_required',
      affectedArea: 'محرك التحليلات / المسابقات',
      details: computedBadValues.slice(0, 10),
      readyPrompt: [
        'أصلح ظهور قيم NaN أو Infinity في مخرجات الحسابات بالخادم (server/analytics-engine.ts أو server/competitions-engine.ts):',
        `- القيم المكتشفة: ${computedBadValues.slice(0, 5).join(' | ')}`,
        '- تأكد من استخدام Number.isFinite() وحماية القسمة على صفر (target > 0) قبل إرجاع نسب التحقيق أو الإجماليات.',
      ].join('\n'),
    });
  }

  // ==================== CHECK 2: SECTOR CONSISTENCY & BRANCH/ITEM INTEGRITY ====================
  let consistencyIssueCount = 0;

  if (activeBatch) {
    // 2a. Verify dashboard sector sales/tax parity against raw SQLite sales_records + cross-tier item categories parity
    const sectorMismatchDetails: string[] = [];
    for (const src of ['VF05N', 'ZSALES'] as const) {
      const dash = getDashboardSummary(db, adminUser, { source: src, sector: 'ALL' });
      const sqlSectorSums = db
        .prepare(
          `SELECT resolved_sector AS sec, COALESCE(SUM(sales_value), 0) AS sales, COALESCE(SUM(tax_value), 0) AS tax
           FROM sales_records
           WHERE batch_id = ? AND source_type = ? AND resolved_sector IS NOT NULL
           GROUP BY resolved_sector`
        )
        .all(activeBatch.id, src) as Array<{ sec: string; sales: number; tax: number }>;

      const sqlMap = new Map(sqlSectorSums.map((r) => [r.sec, r]));
      for (const sRow of dash.sectors) {
        const raw = sqlMap.get(sRow.sectorName);
        const rawSales = raw ? Number(raw.sales.toFixed(2)) : 0;
        const rawTax = raw ? Number(raw.tax.toFixed(2)) : 0;
        if (Math.abs(sRow.sales - rawSales) > 0.05 || Math.abs(sRow.tax - rawTax) > 0.05) {
          sectorMismatchDetails.push(`${src} / ${sRow.sectorName}: لوحة=${sRow.sales} مقابل القاعدة=${rawSales}`);
        }
      }

      // Cross-job-title parity check on Item Categories tab per sector
      const itemCatsAll = getItemCategoriesAnalytics(db, adminUser, { source: src, sector: 'ALL' });
      const actualSrc = itemCatsAll.effectiveSource;
      const dashForActual = actualSrc === src ? dash : getDashboardSummary(db, adminUser, { source: actualSrc, sector: 'ALL' });

      for (const sec of ORIGINAL_SECTORS) {
        const dashSec = dashForActual.sectors.find((x) => x.sectorName === sec);
        if (!dashSec) continue;
        for (const jt of itemCatsAll.jobTitleOptions) {
          if (normalizeArabic(jt).includes('مدير اداره') || normalizeArabic(jt).includes('مدير الاداره')) continue;
          const jtRes = getItemCategoriesAnalytics(db, adminUser, { source: src, sector: sec, jobTitle: jt });
          if (jtRes.rows.length === 0) continue;
          const sumSales = jtRes.rows.reduce((acc, r) => acc + r.sales, 0);
          const sumTax = jtRes.rows.reduce((acc, r) => acc + r.tax, 0);
          if (Math.abs(sumSales - dashSec.sales) > 1 || Math.abs(sumTax - dashSec.tax) > 1) {
            sectorMismatchDetails.push(
              `تحليل الفئات (${actualSrc}) / ${sec} / ${jt}: إجمالي الفئات=${sumSales.toFixed(2)} مقابل القطاع=${dashSec.sales.toFixed(2)}`
            );
          }
        }
      }
    }

    if (sectorMismatchDetails.length > 0) {
      consistencyIssueCount++;
      findings.push({
        id: 'SECTOR_TOTALS_MISMATCH',
        code: 'SECTOR_TOTALS_MISMATCH',
        title: 'تفاوت في إجماليات القطاع بين التبويبات أو المسميات الوظيفية',
        description: 'رصد الفاحص اختلافًا في إجمالي المبيعات أو الضرائب لنفس القطاع عند المقارنة مع لوحة المعلومات أو بين المسميات الوظيفية.',
        severity: 'high',
        category: 'data_consistency',
        fixType: 'code_fix_required',
        affectedArea: 'لوحة المعلومات / تحليل فئات الأصناف',
        details: sectorMismatchDetails.slice(0, 10),
        readyPrompt: [
          'راجع منطق تجميع القطاعات في server/analytics-engine.ts لإصلاح تفاوت الإجماليات التالي:',
          ...sectorMismatchDetails.slice(0, 6).map((d) => `- ${d}`),
          '- تأكد من توحيد إجمالي المبيعات والضرائب لنفس القطاع عبر المسميات الوظيفية دون إسقاط أي فروع أو فواتير ودون خلط مصدري VF05N وZsales_Price.',
        ].join('\n'),
      });
    }

    // 2b. Check duplicate branch records in `branches` table & branch/rep parity
    const batchBranches = db
      .prepare(`SELECT id, branch_code, branch_name, sector_name, rep_code FROM branches WHERE batch_id = ?`)
      .all(activeBatch.id) as Array<{
      id: number;
      branch_code: string;
      branch_name: string;
      sector_name: string;
      rep_code: string | null;
    }>;
    const seenBranchCodes = new Map<string, number>();
    const dupBranchDetails: string[] = [];
    for (const br of batchBranches) {
      const code = normalizeCode(br.branch_code);
      const sec = canonicalizeSectorName(br.sector_name);
      if (!code) continue;
      const key = `${sec}::${code}`;
      if (seenBranchCodes.has(key)) {
        dupBranchDetails.push(`كود مكرر ${code} في قطاع ${sec} (${br.branch_name})`);
      } else {
        seenBranchCodes.set(key, br.id);
      }
    }
    if (dupBranchDetails.length > 0) {
      consistencyIssueCount++;
      findings.push({
        id: 'DUP_BRANCH_ROWS',
        code: 'DUP_BRANCH_ROWS',
        title: 'تكرار أكواد فروع في جدول الفروع للدفعة النشطة',
        description: `تم العثور على ${dupBranchDetails.length} صف فرع مكرر بنفس الكود والقطاع مما قد يؤثر على دقة عداد الفروع أو الأهداف.`,
        severity: 'high',
        category: 'data_consistency',
        fixType: 'auto_fixable',
        affectedArea: 'تاب الفروع / جهاز البيع',
        details: dupBranchDetails.slice(0, 10),
        autoFixAction: 'FIX_DUPLICATE_BRANCH_RECORDS',
        autoFixLabel: 'دمج وإزالة صفوف الفروع المكررة تلقائياً',
      });
    }

    const branchDiag = getBranchesAnalytics(db, adminUser, { source: 'VF05N', sector: 'ALL' }).diagnostics;
    const repsAuditList = getRepsAnalytics(db, adminUser, { source: 'VF05N', sector: 'ALL' });
    const actualRepsBranchCount = repsAuditList.reduce((acc, r) => acc + Math.trunc(r.branchCount || 0), 0);
    const actualBranchRepDiff = actualRepsBranchCount - branchDiag.totalTargetedAssignedBranches;
    if (actualBranchRepDiff !== 0) {
      consistencyIssueCount++;
      findings.push({
        id: 'BRANCH_REP_COUNT_DIFF',
        code: 'BRANCH_REP_COUNT_DIFF',
        title: 'فرق في إجمالي عدد الفروع بين حساب الفروع وحساب المناديب',
        description: `يوجد فرق بمقدار ${actualBranchRepDiff} فرع بين إجمالي الفروع المنسوبة والمناديب.`,
        severity: 'medium',
        category: 'data_consistency',
        fixType: 'code_fix_required',
        affectedArea: 'الفروع وجهاز البيع',
        details: [
          `المتوقع على المناديب: ${branchDiag.totalTargetedAssignedBranches}`,
          `المحسوب في جهاز البيع: ${actualRepsBranchCount}`,
          `الفرق: ${actualBranchRepDiff}`,
        ],
        readyPrompt: [
          'راجع منطق تجميع الفروع الفريدة لكل مندوب في server/analytics-engine.ts (getBranchesAnalytics / getRepsAnalytics):',
          `- الفرق المرصود: ${actualBranchRepDiff} فرع`,
          '- تحقّق من تفرّد كود الفرع واسم الفرع بعد التطبيع والحفاظ على الأصفار البادئة ومنع احتساب فروع مندوب آخر.',
        ].join('\n'),
      });
    }

    // 2c. Check duplicate item codes in `items` catalog
    const batchItems = db
      .prepare(`SELECT id, item_code, item_name FROM items WHERE batch_id = ?`)
      .all(activeBatch.id) as Array<{ id: number; item_code: string; item_name: string }>;
    const seenItems = new Set<string>();
    const dupItemDetails: string[] = [];
    for (const it of batchItems) {
      const code = normalizeCode(it.item_code);
      if (!code) continue;
      if (seenItems.has(code)) {
        dupItemDetails.push(`كود الصنف المكرر ${code} (${it.item_name})`);
      } else {
        seenItems.add(code);
      }
    }
    if (dupItemDetails.length > 0) {
      consistencyIssueCount++;
      findings.push({
        id: 'DUP_ITEM_CATALOG',
        code: 'DUP_ITEM_CATALOG',
        title: 'تكرار أكواد أصناف في جدول تعريف الأصناف',
        description: `تم رصد ${dupItemDetails.length} صنف مكرر بنفس الكود في الدفعة النشطة.`,
        severity: 'medium',
        category: 'data_consistency',
        fixType: 'auto_fixable',
        affectedArea: 'تعريف الأصناف / مسابقات',
        details: dupItemDetails.slice(0, 10),
        autoFixAction: 'FIX_DUPLICATE_ITEM_CATALOG',
        autoFixLabel: 'تنظيف وتوحيد الأصناف المكررة تلقائياً',
      });
    }

    // 2d. Unlinked records advisory
    if (activeBatch.unlinked_branch_count > 0 || activeBatch.unlinked_sector_count > 0) {
      findings.push({
        id: 'UNLINKED_EXCEL_RECORDS',
        code: 'UNLINKED_EXCEL_RECORDS',
        title: 'سجلات غير مرتبطة بفرع أو قطاع في الدفعة النشطة',
        description: 'توجد فواتير لم تطابق تعريف الفروع أو السلاسل أثناء الاستيراد؛ يُنصح بمراجعتها في ملف Excel الأصلي.',
        severity: 'low',
        category: 'data_consistency',
        fixType: 'info',
        affectedArea: 'الفروع / تفاصيل المبيعات',
        details: [
          `سجلات غير مرتبطة بفرع: ${activeBatch.unlinked_branch_count}`,
          `سجلات غير مرتبطة بقطاع: ${activeBatch.unlinked_sector_count}`,
        ],
      });
    }
  }

  // 2e. Competition target_qty vs industrial + commercial sync check
  const desyncedCompTargets = db
    .prepare(
      `SELECT id, competition_id, participant_key, target_qty, industrial_target_qty, commercial_target_qty
       FROM competition_targets
       WHERE (industrial_target_qty > 0 OR commercial_target_qty > 0)
         AND ABS(target_qty - (industrial_target_qty + commercial_target_qty)) > 0.005`
    )
    .all() as Array<{
    id: number;
    competition_id: number;
    participant_key: string;
    target_qty: number;
    industrial_target_qty: number;
    commercial_target_qty: number;
  }>;
  if (desyncedCompTargets.length > 0) {
    consistencyIssueCount++;
    findings.push({
      id: 'COMP_TARGET_TOTAL_DESYNC',
      code: 'COMP_TARGET_TOTAL_DESYNC',
      title: 'عدم تطابق إجمالي هدف المسابقة مع مجموع أهداف الصناعى والتجارى',
      description: `يوجد ${desyncedCompTargets.length} سجل هدف مسابقة لا يساوي فيه الإجمالي مجموع (الصناعى + التجارى).`,
      severity: 'medium',
      category: 'data_consistency',
      fixType: 'auto_fixable',
      affectedArea: 'تبويب مسابقات',
      details: desyncedCompTargets
        .slice(0, 10)
        .map((r) => `${r.participant_key}: مسجل=${r.target_qty} بينما (صناعى+تجارى)=${(r.industrial_target_qty + r.commercial_target_qty).toFixed(3)}`),
      autoFixAction: 'FIX_SYNC_COMPETITION_TARGET_TOTALS',
      autoFixLabel: 'مزامنة إجماليات أهداف المسابقات تلقائياً',
    });
  }

  // 2f. Item category targets key canonicalization check
  const itemCatTargets = db
    .prepare(`SELECT id, target_key, sector_name, job_title, employee_name, category_name FROM item_category_targets`)
    .all() as Array<{
    id: number;
    target_key: string;
    sector_name: string;
    job_title: string;
    employee_name: string;
    category_name: string;
  }>;
  const nonCanonKeys: string[] = [];
  for (const row of itemCatTargets) {
    const expectedKey = buildItemCategoryTargetKey(row.sector_name, row.job_title, row.employee_name, row.category_name);
    if (expectedKey !== row.target_key) {
      nonCanonKeys.push(`${row.sector_name} / ${row.employee_name} / ${row.category_name}`);
    }
  }

  // ==================== CHECK 3: RBAC & PERMISSIONS COMPLIANCE ====================
  let rbacIssueCount = 0;
  const allUsers = db
    .prepare(`SELECT id, username, full_name, role, allowed_sectors, is_active FROM users`)
    .all() as Array<{
    id: number;
    username: string;
    full_name: string;
    role: string;
    allowed_sectors: string;
    is_active: number;
  }>;

  const activeAdmins = allUsers.filter((u) => u.role === 'admin' && u.is_active === 1);
  if (activeAdmins.length === 0) {
    rbacIssueCount++;
    findings.push({
      id: 'RBAC_NO_ACTIVE_ADMIN',
      code: 'RBAC_NO_ACTIVE_ADMIN',
      title: 'لا يوجد حساب Admin مفعل في النظام',
      description: 'يجب وجود حساب مدير نظام واحد على الأقل لضمان إمكانية إدارة النظام ورفع ملفات Excel.',
      severity: 'critical',
      category: 'rbac_permissions',
      fixType: 'info',
      affectedArea: 'إدارة المستخدمين والصلاحيات',
      details: ['تأكد من تفعيل حساب أدمن من تبويب إدارة المستخدمين.'],
    });
  }

  const invalidRoleUsers: string[] = [];
  const emptySectorUsers: string[] = [];
  const nonCanonSectorUsers: string[] = [];

  for (const u of allUsers) {
    if (!['admin', 'sector_manager', 'standard_user'].includes(u.role)) {
      invalidRoleUsers.push(`${u.username} (${u.role})`);
      continue;
    }
    if (u.role !== 'admin') {
      let sectors: string[] = [];
      try {
        const parsed = JSON.parse(u.allowed_sectors);
        if (Array.isArray(parsed)) sectors = parsed.map((s) => String(s || '').trim()).filter(Boolean);
      } catch {
        sectors = [];
      }
      const canonicalized = Array.from(
        new Set(
          sectors
            .map((s) => canonicalizeSectorName(s))
            .filter((s) => ORIGINAL_SECTORS.includes(s as (typeof ORIGINAL_SECTORS)[number]))
        )
      );
      if (u.is_active === 1 && canonicalized.length === 0) {
        emptySectorUsers.push(`${u.full_name} (@${u.username})`);
      } else if (JSON.stringify(sectors) !== JSON.stringify(canonicalized)) {
        nonCanonSectorUsers.push(`${u.full_name} (@${u.username}): ${sectors.join('، ')}`);
      }
    }
  }

  if (invalidRoleUsers.length > 0 || emptySectorUsers.length > 0) {
    rbacIssueCount++;
    findings.push({
      id: 'RBAC_USER_SCOPE_INVALID',
      code: 'RBAC_USER_SCOPE_INVALID',
      title: 'حسابات مستخدمين بأدوار غير معرفة أو بدون قطاعات مصرح بها',
      description: 'توجد حسابات غير أدمن تفتقر إلى قطاعات صالحة، مما يمنعها من عرض البيانات بشكل سليم.',
      severity: 'high',
      category: 'rbac_permissions',
      fixType: 'info',
      affectedArea: 'المستخدمون والصلاحيات',
      details: [...invalidRoleUsers.map((x) => `دور غير صالح: ${x}`), ...emptySectorUsers.map((x) => `بدون قطاعات: ${x}`)],
    });
  }

  if (nonCanonKeys.length > 0 || nonCanonSectorUsers.length > 0) {
    findings.push({
      id: 'RBAC_NON_CANONICAL_SCOPES',
      code: 'RBAC_NON_CANONICAL_SCOPES',
      title: 'أسماء قطاعات أو مفاتيح أهداف غير موحدة التهجئة',
      description: 'توجد أسماء قطاعات أو مفاتيح أهداف مكتوبة بتهجئة/همزات غير معيارية وقد تسبب عدم تطابق عند التصفية.',
      severity: 'low',
      category: 'rbac_permissions',
      fixType: 'auto_fixable',
      affectedArea: 'الأهداف وصلاحيات القطاعات',
      details: [...nonCanonKeys.slice(0, 5), ...nonCanonSectorUsers.slice(0, 5)],
      autoFixAction: 'FIX_CANONICALIZE_TARGETS_AND_SCOPES',
      autoFixLabel: 'توحيد أسماء القطاعات ومفاتيح الأهداف تلقائياً',
    });
  }

  // Non-mutating RBAC guard verification probes
  const rbacGuardFailures: string[] = [];
  const probeStdUser: UserProfile = {
    id: 999991,
    username: 'audit_std_probe',
    fullName: 'Probe Standard',
    email: null,
    role: 'standard_user',
    allowedSectors: ['التوريد المباشر'],
    mustChangePassword: false,
    isActive: true,
    createdAt: now,
  };
  const probeSectorMgr: UserProfile = {
    id: 999992,
    username: 'audit_mgr_probe',
    fullName: 'Probe SectorMgr',
    email: null,
    role: 'sector_manager',
    allowedSectors: ['التوريد المباشر'],
    mustChangePassword: false,
    isActive: true,
    createdAt: now,
  };

  try {
    saveItemCategoryTargets(
      db,
      [{ sectorName: 'التوريد المباشر', jobTitle: 'مندوب مبيعات', employeeName: 'probe', categoryName: 'probe', targetValue: 1 }],
      probeStdUser
    );
    rbacGuardFailures.push('saveItemCategoryTargets لم يمنع standard_user!');
  } catch (err) {
    if (!(err instanceof AuthorizationError)) {
      rbacGuardFailures.push(`saveItemCategoryTargets أعاد خطأ غير متوقع لـ standard_user: ${String(err)}`);
    }
  }

  try {
    saveItemCategoryTargets(
      db,
      [{ sectorName: 'التوريد المركزي', jobTitle: 'مندوب مبيعات', employeeName: 'probe', categoryName: 'probe', targetValue: 1 }],
      probeSectorMgr
    );
    rbacGuardFailures.push('saveItemCategoryTargets لم يمنع مدير قطاع خارج نطاقه!');
  } catch (err) {
    if (!(err instanceof AuthorizationError)) {
      rbacGuardFailures.push(`saveItemCategoryTargets أعاد خطأ غير متوقع خارج النطاق: ${String(err)}`);
    }
  }

  if (allComps.length > 0) {
    try {
      saveCompetitionTargets(
        db,
        allComps[0].id,
        [{ participantKey: 'REP::التوريد المباشر::PROBE', industrialTargetQty: 1, commercialTargetQty: 1 }],
        probeStdUser
      );
      rbacGuardFailures.push('saveCompetitionTargets لم يمنع standard_user!');
    } catch (err) {
      if (!(err instanceof AuthorizationError)) {
        rbacGuardFailures.push(`saveCompetitionTargets أعاد خطأ غير متوقع لـ standard_user: ${String(err)}`);
      }
    }

    try {
      saveCompetitionTargets(
        db,
        allComps[0].id,
        [{ participantKey: 'REP::التوريد المركزي::PROBE', industrialTargetQty: 1, commercialTargetQty: 1 }],
        probeSectorMgr
      );
      rbacGuardFailures.push('saveCompetitionTargets لم يمنع مدير قطاع خارج نطاقه!');
    } catch (err) {
      if (!(err instanceof AuthorizationError)) {
        rbacGuardFailures.push(`saveCompetitionTargets أعاد خطأ غير متوقع خارج النطاق: ${String(err)}`);
      }
    }
  }

  if (rbacGuardFailures.length > 0) {
    rbacIssueCount++;
    findings.push({
      id: 'RBAC_GUARD_FAILURE',
      code: 'RBAC_GUARD_FAILURE',
      title: 'ثغرة في التحقق من صلاحيات تعديل الأهداف',
      description: 'فشل فحص الصلاحيات التلقائي في منع مستخدم غير مصرح له أو خارج قطاعه من محاولة تعديل الأهداف.',
      severity: 'critical',
      category: 'rbac_permissions',
      fixType: 'code_fix_required',
      affectedArea: 'حماية الصلاحيات (RBAC)',
      details: rbacGuardFailures,
      readyPrompt: [
        'أصلح فحص صلاحيات الأدوار والنطاق في الخادم (server/auth.ts، server/analytics-engine.ts، server/competitions-engine.ts):',
        ...rbacGuardFailures.map((f) => `- ${f}`),
        '- تأكد من قصر رفع ملفات Excel على Admin فقط، ومنع Standard User من تعديل أي أهداف، وتقييد Sector Manager بالقطاعات المصرح بها فقط.',
      ].join('\n'),
    });
  }

  // ==================== CHECK 4: CAPTURED RUNTIME ERRORS ====================
  const runtimeErrors = listRuntimeErrorLogs(db, false);
  for (const err of runtimeErrors.slice(0, 15)) {
    const lowMsg = err.message.toLowerCase();
    const isCritical =
      err.kind === 'react_error' ||
      lowMsg.includes('maximum update depth') ||
      lowMsg.includes('referenceerror') ||
      lowMsg.includes('typeerror');
    const severity: AuditFinding['severity'] = isCritical
      ? 'critical'
      : err.kind === 'unhandled_rejection'
        ? 'high'
        : 'medium';

    findings.push({
      id: `RUNTIME_ERR_${err.id}`,
      code: `RUNTIME_ERR_${err.id}`,
      title: `خطأ تشغيل مسجل في تبويب «${err.tabName || 'غير محدد'}»`,
      description: err.message,
      severity,
      category: 'runtime_error',
      fixType: 'code_fix_required',
      affectedArea: `التبويب: ${err.tabName || 'التطبيق'}`,
      details: [
        `النوع: ${err.kind}`,
        `الوقت: ${err.occurredAt}`,
        ...(err.sourceUrl ? [`المصدر: ${err.sourceUrl}:${err.lineNo}:${err.colNo}`] : []),
        ...(err.stack ? err.stack.split('\n').slice(0, 3) : []),
      ],
      readyPrompt: buildLocalRuntimeErrorPrompt(err),
      aiInsight: err.aiAnalysis || undefined,
    });
  }

  if (runtimeErrors.length > 0) {
    findings.push({
      id: 'RUNTIME_ERR_ARCHIVE',
      code: 'RUNTIME_ERR_ARCHIVE',
      title: `توجد ${runtimeErrors.length} أخطاء تشغيل غير مؤرشفة في السجل`,
      description: 'بعد مراجعة الأخطاء وتطبيق الإصلاحات، يمكنك أرشفة السجلات المعلقة للحفاظ على نظافة لوحة المراقبة.',
      severity: 'low',
      category: 'runtime_error',
      fixType: 'auto_fixable',
      affectedArea: 'سجل أخطاء التشغيل',
      details: runtimeErrors.slice(0, 5).map((e) => `#${e.id} (${e.tabName}): ${e.message}`),
      autoFixAction: 'FIX_RESOLVE_RUNTIME_ERRORS',
      autoFixLabel: 'أرشفة الأخطاء المسجلة بعد المراجعة',
    });
  }

  // Sort findings by severity
  const sevRank: Record<AuditFinding['severity'], number> = { critical: 0, high: 1, medium: 2, low: 3 };
  findings.sort((a, b) => sevRank[a.severity] - sevRank[b.severity]);

  const counts = {
    total: findings.length,
    critical: findings.filter((f) => f.severity === 'critical').length,
    high: findings.filter((f) => f.severity === 'high').length,
    medium: findings.filter((f) => f.severity === 'medium').length,
    low: findings.filter((f) => f.severity === 'low').length,
    autoFixable: findings.filter((f) => f.fixType === 'auto_fixable').length,
    codeFixRequired: findings.filter((f) => f.fixType === 'code_fix_required').length,
    unresolvedRuntimeErrors: runtimeErrors.length,
  };

  const penalty = counts.critical * 25 + counts.high * 12 + counts.medium * 5 + counts.low * 2;
  const healthScore = Math.max(0, Math.min(100, 100 - penalty));

  const checksExecuted = [
    {
      key: 'numeric_integrity',
      label: 'فحص القيم غير المعرفة (NaN / Infinity) والأرقام السالبة',
      passed: badDbNumericDetails.length === 0 && computedBadValues.length === 0,
      issueCount: (badDbNumericDetails.length > 0 ? 1 : 0) + (computedBadValues.length > 0 ? 1 : 0),
    },
    {
      key: 'data_consistency',
      label: 'فحص تناسق إجماليات القطاعات وتكرار الفروع والأصناف والأهداف',
      passed: consistencyIssueCount === 0,
      issueCount: consistencyIssueCount,
    },
    {
      key: 'rbac_permissions',
      label: 'فحص التزام الصلاحيات (Admin / مدير قطاع / مستخدم قياسي)',
      passed: rbacIssueCount === 0 && nonCanonKeys.length === 0 && nonCanonSectorUsers.length === 0,
      issueCount: rbacIssueCount + (nonCanonKeys.length > 0 || nonCanonSectorUsers.length > 0 ? 1 : 0),
    },
    {
      key: 'runtime_errors',
      label: 'مراقبة أخطاء التشغيل (window.onerror / unhandledrejection)',
      passed: runtimeErrors.length === 0,
      issueCount: runtimeErrors.length,
    },
  ];

  const aiConfigured = Boolean(options?.aiConfigured);
  const localSummary =
    counts.total === 0
      ? 'النظام سليم 100%: لا توجد قيم NaN/Infinity، إجماليات القطاعات والفروع متناسقة، قواعد الصلاحيات محكمة، ولا توجد أخطاء تشغيل مسجلة.'
      : `تم رصد ${counts.total} ملاحظة (${counts.critical} حرجة، ${counts.high} عالية، ${counts.medium} متوسطة، ${counts.low} منخفضة). يتوفر ${counts.autoFixable} إصلاح تلقائي آمن و${counts.codeFixRequired} برومبت إصلاح برمجي جاهز للنسخ.`;

  const insertRun = db
    .prepare(
      `INSERT INTO audit_runs (
        triggered_by_id, triggered_by_name, health_score, total_findings, critical_count, high_count, medium_count, low_count, findings_json, ai_summary, created_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
    )
    .run(
      adminUser.id,
      adminUser.fullName,
      healthScore,
      counts.total,
      counts.critical,
      counts.high,
      counts.medium,
      counts.low,
      JSON.stringify(findings),
      localSummary,
      now
    );

  return {
    runId: Number(insertRun.lastInsertRowid),
    ranAt: now,
    triggeredByName: adminUser.fullName,
    healthScore,
    aiConfigured,
    aiSummary: localSummary,
    activeBatchId: activeBatch?.id ?? null,
    activeBatchFilename: activeBatch?.filename ?? null,
    checksExecuted,
    summaryCounts: counts,
    findings,
    runtimeErrors,
    backups: listAuditBackups(db),
  };
}

export function executeAutoFix(
  db: Database.Database,
  adminUser: UserProfile,
  actionType: AutoFixActionType,
  findingCode?: string,
  options?: { aiConfigured?: boolean }
): { backupId: number; summary: string; report: SmartAuditReport } {
  const now = new Date().toISOString();
  const activeBatch = db
    .prepare(`SELECT id FROM import_batches WHERE status = 'active' ORDER BY id DESC LIMIT 1`)
    .get() as { id: number } | undefined;

  let backupId = 0;
  let summary = '';

  const tx = db.transaction(() => {
    if (actionType === 'FIX_NON_FINITE_NUMBERS') {
      const snapshot = {
        salesRows: activeBatch
          ? db.prepare(`SELECT id, quantity, sales_value, tax_value FROM sales_records WHERE batch_id = ?`).all(activeBatch.id)
          : [],
        branches: activeBatch
          ? db.prepare(`SELECT id, branch_target FROM branches WHERE batch_id = ?`).all(activeBatch.id)
          : [],
        itemCategoryTargets: db.prepare(`SELECT id, target_value FROM item_category_targets`).all(),
        competitionTargets: db
          .prepare(`SELECT id, target_qty, industrial_target_qty, commercial_target_qty FROM competition_targets`)
          .all(),
        competitionItems: db.prepare(`SELECT id, equivalent_factor FROM competition_items`).all(),
      };

      let fixedCount = 0;
      if (activeBatch) {
        for (const r of snapshot.salesRows as Array<{ id: number; quantity: unknown; sales_value: unknown; tax_value: unknown }>) {
          if (isBadNumber(r.quantity) || isBadNumber(r.sales_value) || isBadNumber(r.tax_value)) {
            db.prepare(`UPDATE sales_records SET quantity = ?, sales_value = ?, tax_value = ? WHERE id = ?`).run(
              isBadNumber(r.quantity) ? 0 : Number(r.quantity),
              isBadNumber(r.sales_value) ? 0 : Number(r.sales_value),
              isBadNumber(r.tax_value) ? 0 : Number(r.tax_value),
              r.id
            );
            fixedCount++;
          }
        }
        for (const b of snapshot.branches as Array<{ id: number; branch_target: unknown }>) {
          if (isBadNumber(b.branch_target) || Number(b.branch_target) < 0) {
            db.prepare(`UPDATE branches SET branch_target = 0 WHERE id = ?`).run(b.id);
            fixedCount++;
          }
        }
      }
      for (const ct of snapshot.itemCategoryTargets as Array<{ id: number; target_value: unknown }>) {
        if (isBadNumber(ct.target_value) || Number(ct.target_value) < 0) {
          db.prepare(`UPDATE item_category_targets SET target_value = 0 WHERE id = ?`).run(ct.id);
          fixedCount++;
        }
      }
      for (const cp of snapshot.competitionTargets as Array<{
        id: number;
        target_qty: unknown;
        industrial_target_qty: unknown;
        commercial_target_qty: unknown;
      }>) {
        if (
          isBadNumber(cp.target_qty) ||
          isBadNumber(cp.industrial_target_qty) ||
          isBadNumber(cp.commercial_target_qty) ||
          Number(cp.target_qty) < 0 ||
          Number(cp.industrial_target_qty) < 0 ||
          Number(cp.commercial_target_qty) < 0
        ) {
          const ind = isBadNumber(cp.industrial_target_qty) || Number(cp.industrial_target_qty) < 0 ? 0 : Number(cp.industrial_target_qty);
          const com = isBadNumber(cp.commercial_target_qty) || Number(cp.commercial_target_qty) < 0 ? 0 : Number(cp.commercial_target_qty);
          const tot = isBadNumber(cp.target_qty) || Number(cp.target_qty) < 0 ? Number((ind + com).toFixed(3)) : Number(cp.target_qty);
          db.prepare(`UPDATE competition_targets SET target_qty = ?, industrial_target_qty = ?, commercial_target_qty = ? WHERE id = ?`).run(
            tot,
            ind,
            com,
            cp.id
          );
          fixedCount++;
        }
      }
      for (const ci of snapshot.competitionItems as Array<{ id: number; equivalent_factor: unknown }>) {
        if (isBadNumber(ci.equivalent_factor) || Number(ci.equivalent_factor) < 0) {
          db.prepare(`UPDATE competition_items SET equivalent_factor = 1 WHERE id = ?`).run(ci.id);
          fixedCount++;
        }
      }

      summary = `تم تنظيف وتصحيح ${fixedCount} سجل يحتوي على قيم رقمية غير صالحة أو سالبة.`;
      const res = db
        .prepare(
          `INSERT INTO audit_fix_backups (finding_code, action_type, action_label, summary, snapshot_json, created_by_id, created_by_name, created_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
        )
        .run(
          findingCode || 'NUM_DB_NON_FINITE',
          actionType,
          'تنظيف القيم الرقمية غير الصالحة',
          summary,
          JSON.stringify(snapshot),
          adminUser.id,
          adminUser.fullName,
          now
        );
      backupId = Number(res.lastInsertRowid);
    } else if (actionType === 'FIX_DUPLICATE_BRANCH_RECORDS') {
      if (!activeBatch) throw new Error('لا توجد دفعة بيانات نشطة');
      const allBranches = db.prepare(`SELECT * FROM branches WHERE batch_id = ? ORDER BY id ASC`).all(activeBatch.id) as Array<{
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
      }>;
      const keepMap = new Map<string, (typeof allBranches)[number]>();
      const deleteIds: number[] = [];
      for (const b of allBranches) {
        const key = `${canonicalizeSectorName(b.sector_name)}::${normalizeCode(b.branch_code)}`;
        const existing = keepMap.get(key);
        if (!existing) {
          keepMap.set(key, { ...b });
        } else {
          existing.branch_target = Number((Number(existing.branch_target || 0) + Number(b.branch_target || 0)).toFixed(2));
          deleteIds.push(b.id);
        }
      }
      for (const kept of keepMap.values()) {
        db.prepare(`UPDATE branches SET branch_target = ? WHERE id = ?`).run(kept.branch_target, kept.id);
      }
      for (const delId of deleteIds) {
        db.prepare(`DELETE FROM branches WHERE id = ?`).run(delId);
      }

      summary = `تم دمج وإزالة ${deleteIds.length} سجل فرع مكرر في الدفعة النشطة.`;
      const res = db
        .prepare(
          `INSERT INTO audit_fix_backups (finding_code, action_type, action_label, summary, snapshot_json, created_by_id, created_by_name, created_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
        )
        .run(
          findingCode || 'DUP_BRANCH_ROWS',
          actionType,
          'دمج صفوف الفروع المكررة',
          summary,
          JSON.stringify({ batchId: activeBatch.id, branches: allBranches }),
          adminUser.id,
          adminUser.fullName,
          now
        );
      backupId = Number(res.lastInsertRowid);
    } else if (actionType === 'FIX_DUPLICATE_ITEM_CATALOG') {
      if (!activeBatch) throw new Error('لا توجد دفعة بيانات نشطة');
      const allItems = db.prepare(`SELECT * FROM items WHERE batch_id = ? ORDER BY id ASC`).all(activeBatch.id) as Array<{
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
      }>;
      const seen = new Set<string>();
      const deleteIds: number[] = [];
      for (const it of allItems) {
        const code = normalizeCode(it.item_code);
        if (!code) continue;
        if (seen.has(code)) {
          deleteIds.push(it.id);
        } else {
          seen.add(code);
        }
      }
      for (const delId of deleteIds) {
        db.prepare(`DELETE FROM items WHERE id = ?`).run(delId);
      }

      summary = `تمت إزالة ${deleteIds.length} سجل صنف مكرر من جدول تعريف الأصناف.`;
      const res = db
        .prepare(
          `INSERT INTO audit_fix_backups (finding_code, action_type, action_label, summary, snapshot_json, created_by_id, created_by_name, created_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
        )
        .run(
          findingCode || 'DUP_ITEM_CATALOG',
          actionType,
          'توحيد جدول تعريف الأصناف',
          summary,
          JSON.stringify({ batchId: activeBatch.id, items: allItems }),
          adminUser.id,
          adminUser.fullName,
          now
        );
      backupId = Number(res.lastInsertRowid);
    } else if (actionType === 'FIX_SYNC_COMPETITION_TARGET_TOTALS') {
      const allCompTargets = db.prepare(`SELECT * FROM competition_targets`).all() as Array<{
        id: number;
        target_qty: number;
        industrial_target_qty: number;
        commercial_target_qty: number;
      }>;
      let synced = 0;
      for (const row of allCompTargets) {
        if (row.industrial_target_qty > 0 || row.commercial_target_qty > 0) {
          const expected = Number((Number(row.industrial_target_qty) + Number(row.commercial_target_qty)).toFixed(3));
          if (Math.abs(Number(row.target_qty) - expected) > 0.001) {
            db.prepare(`UPDATE competition_targets SET target_qty = ? WHERE id = ?`).run(expected, row.id);
            synced++;
          }
        }
      }
      summary = `تمت مزامنة ${synced} سجل هدف مسابقة ليطابق مجموع (الصناعى + التجارى).`;
      const res = db
        .prepare(
          `INSERT INTO audit_fix_backups (finding_code, action_type, action_label, summary, snapshot_json, created_by_id, created_by_name, created_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
        )
        .run(
          findingCode || 'COMP_TARGET_TOTAL_DESYNC',
          actionType,
          'مزامنة إجماليات أهداف المسابقات',
          summary,
          JSON.stringify({ competitionTargets: allCompTargets }),
          adminUser.id,
          adminUser.fullName,
          now
        );
      backupId = Number(res.lastInsertRowid);
    } else if (actionType === 'FIX_CANONICALIZE_TARGETS_AND_SCOPES') {
      const usersSnap = db.prepare(`SELECT id, role, allowed_sectors FROM users`).all() as Array<{
        id: number;
        role: string;
        allowed_sectors: string;
      }>;
      const catTargetsSnap = db.prepare(`SELECT * FROM item_category_targets`).all() as Array<{
        id: number;
        target_key: string;
        sector_name: string;
        job_title: string;
        employee_name: string;
        category_name: string;
        target_value: number;
        updated_at: string;
      }>;

      let fixedCount = 0;
      for (const u of usersSnap) {
        if (u.role === 'admin') continue;
        let sectors: string[] = [];
        try {
          const p = JSON.parse(u.allowed_sectors);
          if (Array.isArray(p)) sectors = p.map((s) => String(s || '').trim()).filter(Boolean);
        } catch {
          sectors = [];
        }
        const canon = Array.from(
          new Set(
            sectors
              .map((s) => canonicalizeSectorName(s))
              .filter((s) => ORIGINAL_SECTORS.includes(s as (typeof ORIGINAL_SECTORS)[number]))
          )
        );
        if (canon.length > 0 && JSON.stringify(canon) !== u.allowed_sectors) {
          db.prepare(`UPDATE users SET allowed_sectors = ? WHERE id = ?`).run(JSON.stringify(canon), u.id);
          fixedCount++;
        }
      }

      for (const ct of catTargetsSnap) {
        const expected = buildItemCategoryTargetKey(ct.sector_name, ct.job_title, ct.employee_name, ct.category_name);
        if (expected !== ct.target_key) {
          const conflict = db.prepare(`SELECT id FROM item_category_targets WHERE target_key = ?`).get(expected) as
            | { id: number }
            | undefined;
          if (conflict) {
            db.prepare(`DELETE FROM item_category_targets WHERE id = ?`).run(ct.id);
          } else {
            db.prepare(`UPDATE item_category_targets SET target_key = ? WHERE id = ?`).run(expected, ct.id);
          }
          fixedCount++;
        }
      }

      summary = `تم توحيد ${fixedCount} سجل من أسماء القطاعات ومفاتيح الأهداف.`;
      const res = db
        .prepare(
          `INSERT INTO audit_fix_backups (finding_code, action_type, action_label, summary, snapshot_json, created_by_id, created_by_name, created_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
        )
        .run(
          findingCode || 'RBAC_NON_CANONICAL_SCOPES',
          actionType,
          'توحيد أسماء القطاعات ومفاتيح الأهداف',
          summary,
          JSON.stringify({ users: usersSnap, itemCategoryTargets: catTargetsSnap }),
          adminUser.id,
          adminUser.fullName,
          now
        );
      backupId = Number(res.lastInsertRowid);
    } else if (actionType === 'FIX_RESOLVE_RUNTIME_ERRORS') {
      const unresolved = db.prepare(`SELECT id FROM runtime_error_logs WHERE resolved = 0`).all() as Array<{ id: number }>;
      db.prepare(`UPDATE runtime_error_logs SET resolved = 1 WHERE resolved = 0`).run();
      summary = `تمت أرشفة ${unresolved.length} خطأ تشغيل من السجل النشط.`;
      const res = db
        .prepare(
          `INSERT INTO audit_fix_backups (finding_code, action_type, action_label, summary, snapshot_json, created_by_id, created_by_name, created_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
        )
        .run(
          findingCode || 'RUNTIME_ERR_ARCHIVE',
          actionType,
          'أرشفة أخطاء التشغيل المسجلة',
          summary,
          JSON.stringify({ unresolvedIds: unresolved.map((x) => x.id) }),
          adminUser.id,
          adminUser.fullName,
          now
        );
      backupId = Number(res.lastInsertRowid);
    } else {
      throw new Error(`إجراء الإصلاح التلقائي غير مدعوم: ${actionType}`);
    }
  });

  tx();
  const report = runSmartAudit(db, adminUser, { aiConfigured: options?.aiConfigured });
  return { backupId, summary, report };
}

export function rollbackAutoFix(
  db: Database.Database,
  adminUser: UserProfile,
  backupId: number,
  options?: { aiConfigured?: boolean }
): { summary: string; report: SmartAuditReport } {
  const backup = db
    .prepare(`SELECT * FROM audit_fix_backups WHERE id = ?`)
    .get(backupId) as
    | {
        id: number;
        action_type: string;
        action_label: string;
        snapshot_json: string;
        restored_at: string | null;
      }
    | undefined;

  if (!backup) {
    throw new Error('النسخة الاحتياطية المطلوبة غير موجودة');
  }
  if (backup.restored_at) {
    throw new Error('تمت استعادة هذه النسخة الاحتياطية مسبقاً');
  }

  const snap = JSON.parse(backup.snapshot_json) as Record<string, unknown>;
  const now = new Date().toISOString();

  const tx = db.transaction(() => {
    if (backup.action_type === 'FIX_NON_FINITE_NUMBERS') {
      for (const r of (snap.salesRows as Array<{ id: number; quantity: number; sales_value: number; tax_value: number }>) || []) {
        db.prepare(`UPDATE sales_records SET quantity = ?, sales_value = ?, tax_value = ? WHERE id = ?`).run(
          r.quantity,
          r.sales_value,
          r.tax_value,
          r.id
        );
      }
      for (const b of (snap.branches as Array<{ id: number; branch_target: number }>) || []) {
        db.prepare(`UPDATE branches SET branch_target = ? WHERE id = ?`).run(b.branch_target, b.id);
      }
      for (const ct of (snap.itemCategoryTargets as Array<{ id: number; target_value: number }>) || []) {
        db.prepare(`UPDATE item_category_targets SET target_value = ? WHERE id = ?`).run(ct.target_value, ct.id);
      }
      for (const cp of (snap.competitionTargets as Array<{
        id: number;
        target_qty: number;
        industrial_target_qty: number;
        commercial_target_qty: number;
      }>) || []) {
        db.prepare(`UPDATE competition_targets SET target_qty = ?, industrial_target_qty = ?, commercial_target_qty = ? WHERE id = ?`).run(
          cp.target_qty,
          cp.industrial_target_qty,
          cp.commercial_target_qty,
          cp.id
        );
      }
      for (const ci of (snap.competitionItems as Array<{ id: number; equivalent_factor: number }>) || []) {
        db.prepare(`UPDATE competition_items SET equivalent_factor = ? WHERE id = ?`).run(ci.equivalent_factor, ci.id);
      }
    } else if (backup.action_type === 'FIX_DUPLICATE_BRANCH_RECORDS') {
      const batchId = Number(snap.batchId);
      const branches = (snap.branches as Array<{
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
      }>) || [];
      db.prepare(`DELETE FROM branches WHERE batch_id = ?`).run(batchId);
      const ins = db.prepare(
        `INSERT INTO branches (id, batch_id, branch_code, branch_name, sector_name, branch_target, chain_code, chain_name, rep_code, rep_name, supervisor_name)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
      );
      for (const b of branches) {
        ins.run(b.id, b.batch_id, b.branch_code, b.branch_name, b.sector_name, b.branch_target, b.chain_code, b.chain_name, b.rep_code, b.rep_name, b.supervisor_name);
      }
    } else if (backup.action_type === 'FIX_DUPLICATE_ITEM_CATALOG') {
      const batchId = Number(snap.batchId);
      const items = (snap.items as Array<{
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
      }>) || [];
      db.prepare(`DELETE FROM items WHERE batch_id = ?`).run(batchId);
      const ins = db.prepare(
        `INSERT INTO items (id, batch_id, item_code, item_name, category_name, brand_name, unit_name, chains_category, army_category, top_retail_category)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
      );
      for (const it of items) {
        ins.run(it.id, it.batch_id, it.item_code, it.item_name, it.category_name, it.brand_name, it.unit_name, it.chains_category, it.army_category, it.top_retail_category);
      }
    } else if (backup.action_type === 'FIX_SYNC_COMPETITION_TARGET_TOTALS') {
      const rows = (snap.competitionTargets as Array<{
        id: number;
        target_qty: number;
        industrial_target_qty: number;
        commercial_target_qty: number;
      }>) || [];
      for (const r of rows) {
        db.prepare(`UPDATE competition_targets SET target_qty = ?, industrial_target_qty = ?, commercial_target_qty = ? WHERE id = ?`).run(
          r.target_qty,
          r.industrial_target_qty,
          r.commercial_target_qty,
          r.id
        );
      }
    } else if (backup.action_type === 'FIX_CANONICALIZE_TARGETS_AND_SCOPES') {
      const usersSnap = (snap.users as Array<{ id: number; allowed_sectors: string }>) || [];
      for (const u of usersSnap) {
        db.prepare(`UPDATE users SET allowed_sectors = ? WHERE id = ?`).run(u.allowed_sectors, u.id);
      }
      const catTargets = (snap.itemCategoryTargets as Array<{
        id: number;
        target_key: string;
        sector_name: string;
        job_title: string;
        employee_name: string;
        category_name: string;
        target_value: number;
        updated_at: string;
      }>) || [];
      db.prepare(`DELETE FROM item_category_targets`).run();
      const ins = db.prepare(
        `INSERT INTO item_category_targets (id, target_key, sector_name, job_title, employee_name, category_name, target_value, updated_at)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
      );
      for (const ct of catTargets) {
        ins.run(ct.id, ct.target_key, ct.sector_name, ct.job_title, ct.employee_name, ct.category_name, ct.target_value, ct.updated_at);
      }
    } else if (backup.action_type === 'FIX_RESOLVE_RUNTIME_ERRORS') {
      const ids = (snap.unresolvedIds as number[]) || [];
      for (const id of ids) {
        db.prepare(`UPDATE runtime_error_logs SET resolved = 0 WHERE id = ?`).run(id);
      }
    }

    db.prepare(`UPDATE audit_fix_backups SET restored_at = ? WHERE id = ?`).run(now, backupId);
  });

  tx();
  const report = runSmartAudit(db, adminUser, { aiConfigured: options?.aiConfigured });
  return {
    summary: `تمت استعادة الحالة قبل «${backup.action_label}» بنجاح.`,
    report,
  };
}
