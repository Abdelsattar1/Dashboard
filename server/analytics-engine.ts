import type Database from 'better-sqlite3';
import type {
  ActiveBatchInfo,
  AggregationMetricRow,
  BranchMetricRow,
  BranchRepBreakdown,
  BranchRepDiagnostics,
  DashboardSummaryResponse,
  DepartmentChainRow,
  ItemCategoriesAnalyticsResponse,
  ItemCategoryAnalysisRow,
  ItemMetricRow,
  RepMetricRow,
  SalesSource,
  SalesTransactionRow,
  SectorMetricRow,
  TimelinePoint,
  UnlinkedRecordRow,
  UserProfile,
} from '../src/types/bi.ts';
import {
  canonicalizeSectorName,
  normalizeArabic,
  normalizeCode,
  ORIGINAL_SECTORS,
  stripLeadingZeros,
} from './excel-engine.ts';

export function calculateAchievementPct(sales: number, target: number, tax: number): number | null {
  const denominator = target - tax;
  if (denominator <= 0) {
    return null;
  }
  return Number(((sales / denominator) * 100).toFixed(2));
}

export interface AnalyticsFilterParams {
  source: SalesSource;
  startDate?: string;
  endDate?: string;
  sector?: string;
}

interface BatchDbRow {
  id: number;
  batch_uuid: string;
  filename: string;
  imported_by_name: string;
  created_at: string;
  vf05n_count: number;
  vf05n_sales_total: number;
  vf05n_tax_total: number;
  zsales_count: number;
  zsales_sales_total: number;
  zsales_tax_total: number;
  unlinked_branch_count: number;
  unlinked_sector_count: number;
}

function getActiveBatch(db: Database.Database): ActiveBatchInfo | null {
  const row = db
    .prepare(`SELECT * FROM import_batches WHERE status = 'active' ORDER BY id DESC LIMIT 1`)
    .get() as BatchDbRow | undefined;

  if (!row) return null;

  return {
    id: row.id,
    batchUuid: row.batch_uuid,
    filename: row.filename,
    importedByName: row.imported_by_name,
    createdAt: row.created_at,
    vf05nCount: row.vf05n_count,
    vf05nSalesTotal: row.vf05n_sales_total,
    vf05nTaxTotal: row.vf05n_tax_total,
    zsalesCount: row.zsales_count,
    zsalesSalesTotal: row.zsales_sales_total,
    zsalesTaxTotal: row.zsales_tax_total,
    unlinkedBranchCount: row.unlinked_branch_count,
    unlinkedSectorCount: row.unlinked_sector_count,
  };
}

function resolveAuthorizedSectors(user: UserProfile): string[] {
  if (user.role === 'admin') {
    return [...ORIGINAL_SECTORS];
  }
  const normalized = (Array.isArray(user.allowedSectors) ? user.allowedSectors : [])
    .map((s) => canonicalizeSectorName(s))
    .filter(Boolean);
  if (normalized.some((s) => normalizeArabic(s) === 'الكل' || normalizeArabic(s) === 'all')) {
    return [...ORIGINAL_SECTORS];
  }
  return Array.from(new Set(normalized));
}

function buildScopeFilter(
  user: UserProfile,
  params: AnalyticsFilterParams,
  batchId: number,
  alias = 'sr'
): { sql: string; args: unknown[]; authorizedSectors: string[] } {
  const conditions: string[] = [`${alias}.batch_id = ?`, `${alias}.source_type = ?`];
  const args: unknown[] = [batchId, params.source];

  const isAdmin = user.role === 'admin';
  const authorizedSectors = resolveAuthorizedSectors(user);

  if (!isAdmin) {
    if (authorizedSectors.length === 0) {
      conditions.push('1 = 0');
    } else {
      const placeholders = authorizedSectors.map(() => '?').join(', ');
      conditions.push(`${alias}.resolved_sector IN (${placeholders})`);
      args.push(...authorizedSectors);
    }
  }

  const sectorParam =
    params.sector && params.sector !== 'ALL' && params.sector !== 'undefined' && params.sector !== 'null'
      ? params.sector === 'UNLINKED'
        ? 'UNLINKED'
        : canonicalizeSectorName(params.sector)
      : 'ALL';

  if (sectorParam !== 'ALL') {
    if (!isAdmin && !authorizedSectors.includes(sectorParam)) {
      conditions.push('1 = 0');
    } else if (sectorParam === 'UNLINKED') {
      if (isAdmin) {
        conditions.push(`${alias}.resolved_sector IS NULL`);
      } else {
        conditions.push('1 = 0');
      }
    } else {
      conditions.push(`${alias}.resolved_sector = ?`);
      args.push(sectorParam);
    }
  }

  if (params.startDate) {
    conditions.push(`${alias}.billing_date >= ?`);
    args.push(params.startDate);
  }
  if (params.endDate) {
    conditions.push(`${alias}.billing_date <= ?`);
    args.push(params.endDate);
  }

  return {
    sql: conditions.join(' AND '),
    args,
    authorizedSectors,
  };
}

export function getDashboardSummary(
  db: Database.Database,
  user: UserProfile,
  params: AnalyticsFilterParams
): DashboardSummaryResponse {
  const activeBatch = getActiveBatch(db);
  if (!activeBatch) {
    return {
      hasData: false,
      activeBatch: null,
      source: params.source,
      dateRange: { minDate: null, maxDate: null },
      kpis: {
        privateMarketsSales: 0,
        privateMarketsTarget: 0,
        privateMarketsTax: 0,
        privateMarketsVariance: 0,
        privateMarketsAchievementPct: null,
        directAndArmySales: 0,
        directAndArmyTarget: 0,
        directAndArmyAchievementPct: null,
        chainsSectorSales: 0,
        chainsSectorTarget: 0,
        chainsSectorAchievementPct: null,
        unlinkedSectorSales: 0,
        unlinkedSectorCount: 0,
        unlinkedBranchSales: 0,
        unlinkedBranchCount: 0,
        totalSourceRecords: 0,
        totalSourceSales: 0,
        totalSourceTax: 0,
      },
      sectors: [],
      aggregations: [],
      timeline: [],
      sourceComparison: null,
    };
  }

  const scope = buildScopeFilter(user, params, activeBatch.id);

  const dateBounds = db
    .prepare(
      `SELECT MIN(billing_date) AS min_date, MAX(billing_date) AS max_date
       FROM sales_records sr
       WHERE ${scope.sql} AND billing_date != ''`
    )
    .get(...scope.args) as { min_date: string | null; max_date: string | null };

  // Fetch sector targets from DB for this batch
  const dbSectors = db
    .prepare(`SELECT sector_name, monthly_target, is_original_sector FROM sectors WHERE batch_id = ?`)
    .all(activeBatch.id) as Array<{ sector_name: string; monthly_target: number; is_original_sector: number }>;

  const targetBySector = new Map<string, { target: number; isOriginal: boolean }>();
  for (const s of dbSectors) {
    targetBySector.set(s.sector_name, {
      target: s.monthly_target,
      isOriginal: s.is_original_sector === 1,
    });
  }

  // Group sales by resolved_sector
  const groupedSales = db
    .prepare(
      `SELECT
        resolved_sector AS sector_name,
        COALESCE(SUM(sales_value), 0) AS sales,
        COALESCE(SUM(tax_value), 0) AS tax,
        COUNT(*) AS record_count,
        SUM(CASE WHEN is_return = 1 THEN 1 ELSE 0 END) AS return_count,
        COALESCE(SUM(CASE WHEN is_return = 1 THEN sales_value ELSE 0 END), 0) AS return_value
       FROM sales_records sr
       WHERE ${scope.sql}
       GROUP BY resolved_sector`
    )
    .all(...scope.args) as Array<{
    sector_name: string | null;
    sales: number;
    tax: number;
    record_count: number;
    return_count: number;
    return_value: number;
  }>;

  const salesBySector = new Map<
    string,
    { sales: number; tax: number; count: number; returnCount: number; returnValue: number }
  >();
  let unlinkedSectorSales = 0;
  let unlinkedSectorCount = 0;
  let unlinkedSectorTax = 0;

  for (const row of groupedSales) {
    if (!row.sector_name) {
      unlinkedSectorSales += row.sales;
      unlinkedSectorTax += row.tax;
      unlinkedSectorCount += row.record_count;
    } else {
      salesBySector.set(row.sector_name, {
        sales: row.sales,
        tax: row.tax,
        count: row.record_count,
        returnCount: row.return_count,
        returnValue: row.return_value,
      });
    }
  }

  const selectedSector =
    params.sector && params.sector !== 'ALL' && params.sector !== 'undefined' && params.sector !== 'null'
      ? canonicalizeSectorName(params.sector)
      : 'ALL';

  // Determine which original sectors are visible to this user & filter
  const visibleOriginalSectors = ORIGINAL_SECTORS.filter((sName) => {
    if (user.role !== 'admin' && !scope.authorizedSectors.includes(sName)) return false;
    if (selectedSector !== 'ALL' && selectedSector !== sName) return false;
    return true;
  });

  // Also check any non-original sector present in data (if admin)
  const extraSectorNames = new Set<string>();
  for (const name of salesBySector.keys()) {
    if (!ORIGINAL_SECTORS.includes(name as (typeof ORIGINAL_SECTORS)[number])) {
      extraSectorNames.add(name);
    }
  }

  const sectorRows: SectorMetricRow[] = [];
  for (const sName of [...visibleOriginalSectors, ...Array.from(extraSectorNames)]) {
    const agg = salesBySector.get(sName) ?? { sales: 0, tax: 0, count: 0, returnCount: 0, returnValue: 0 };
    const targetInfo = targetBySector.get(sName) ?? {
      target: 0,
      isOriginal: ORIGINAL_SECTORS.includes(sName as (typeof ORIGINAL_SECTORS)[number]),
    };
    const sales = Number(agg.sales.toFixed(2));
    const tax = Number(agg.tax.toFixed(2));
    const target = Number(targetInfo.target.toFixed(2));
    const variance = Number((sales - target).toFixed(2));

    sectorRows.push({
      sectorName: sName,
      sales,
      tax,
      target,
      variance,
      achievementPct: calculateAchievementPct(sales, target, tax),
      recordCount: agg.count,
      returnCount: agg.returnCount,
      returnValue: Number(agg.returnValue.toFixed(2)),
      isOriginalSector: targetInfo.isOriginal,
    });
  }

  // Helper to compute executive aggregations ONLY from authorized original sectors (no double counting)
  const sumSectors = (names: readonly string[]) => {
    let sales = 0;
    let tax = 0;
    let target = 0;
    let countIncluded = 0;
    for (const name of names) {
      if (user.role !== 'admin' && !scope.authorizedSectors.includes(name)) continue;
      if (selectedSector !== 'ALL' && selectedSector !== name) continue;
      const row = sectorRows.find((r) => r.sectorName === name);
      if (row) {
        sales += row.sales;
        tax += row.tax;
        target += row.target;
        countIncluded++;
      }
    }
    sales = Number(sales.toFixed(2));
    tax = Number(tax.toFixed(2));
    target = Number(target.toFixed(2));
    return {
      sales,
      tax,
      target,
      variance: Number((sales - target).toFixed(2)),
      achievementPct: calculateAchievementPct(sales, target, tax),
      isPartial: countIncluded < names.length,
    };
  };

  const privateMarkets = sumSectors(['التوريد المباشر', 'التوريد المركزي', 'التوب ريتيل', 'الجيش']);
  const directAndArmy = sumSectors(['التوريد المباشر', 'التوب ريتيل', 'الجيش']);
  const chainsSector = sumSectors(['التوريد المباشر', 'التوريد المركزي']);

  const aggregations: AggregationMetricRow[] = [
    {
      key: 'private_markets',
      title: 'إجمالي الأسواق الخاصة',
      subtitle: 'مجموع القطاعات الأربعة الأصلية',
      includedSectors: ['التوريد المباشر', 'التوريد المركزي', 'التوب ريتيل', 'الجيش'],
      sales: privateMarkets.sales,
      tax: privateMarkets.tax,
      target: privateMarkets.target,
      variance: privateMarkets.variance,
      achievementPct: privateMarkets.achievementPct,
      isPartialForUser: privateMarkets.isPartial,
    },
    {
      key: 'direct_and_army',
      title: 'إجمالي المباشر والجيش',
      subtitle: 'التوريد المباشر + التوب ريتيل + الجيش',
      includedSectors: ['التوريد المباشر', 'التوب ريتيل', 'الجيش'],
      sales: directAndArmy.sales,
      tax: directAndArmy.tax,
      target: directAndArmy.target,
      variance: directAndArmy.variance,
      achievementPct: directAndArmy.achievementPct,
      isPartialForUser: directAndArmy.isPartial,
    },
    {
      key: 'chains_sector',
      title: 'إجمالي قطاع السلاسل',
      subtitle: 'التوريد المباشر + التوريد المركزي',
      includedSectors: ['التوريد المباشر', 'التوريد المركزي'],
      sales: chainsSector.sales,
      tax: chainsSector.tax,
      target: chainsSector.target,
      variance: chainsSector.variance,
      achievementPct: chainsSector.achievementPct,
      isPartialForUser: chainsSector.isPartial,
    },
  ];

  // Unlinked branches within scope
  const unlinkedBranchRow = db
    .prepare(
      `SELECT COUNT(*) AS cnt, COALESCE(SUM(sales_value), 0) AS sales
       FROM sales_records sr
       WHERE ${scope.sql} AND sr.resolved_branch_code IS NULL`
    )
    .get(...scope.args) as { cnt: number; sales: number };

  const totalScopeRow = db
    .prepare(
      `SELECT COUNT(*) AS cnt, COALESCE(SUM(sales_value), 0) AS sales, COALESCE(SUM(tax_value), 0) AS tax
       FROM sales_records sr
       WHERE ${scope.sql}`
    )
    .get(...scope.args) as { cnt: number; sales: number; tax: number };

  // Timeline
  const timeline = getTimelineSeries(db, user, params);

  // Source comparison (admin sees full batch comparison; sector user sees their authorized scope comparison across VF05N vs ZSALES)
  const buildSourceSum = (src: SalesSource) => {
    const srcScope = buildScopeFilter(user, { ...params, source: src }, activeBatch.id);
    const row = db
      .prepare(
        `SELECT
          COUNT(*) AS cnt,
          COALESCE(SUM(sales_value), 0) AS sales,
          COALESCE(SUM(tax_value), 0) AS tax,
          SUM(CASE WHEN is_return = 1 THEN 1 ELSE 0 END) AS return_cnt
         FROM sales_records sr
         WHERE ${srcScope.sql}`
      )
      .get(...srcScope.args) as { cnt: number; sales: number; tax: number; return_cnt: number };

    return {
      count: row.cnt,
      salesTotal: Number(row.sales.toFixed(2)),
      taxTotal: Number(row.tax.toFixed(2)),
      returnCount: row.return_cnt || 0,
    };
  };

  const vfComp = buildSourceSum('VF05N');
  const zsComp = buildSourceSum('ZSALES');
  const salesDiff = Number((vfComp.salesTotal - zsComp.salesTotal).toFixed(2));
  const taxDiff = Number((vfComp.taxTotal - zsComp.taxTotal).toFixed(2));

  return {
    hasData: true,
    activeBatch,
    source: params.source,
    dateRange: {
      minDate: dateBounds.min_date,
      maxDate: dateBounds.max_date,
    },
    kpis: {
      privateMarketsSales: privateMarkets.sales,
      privateMarketsTarget: privateMarkets.target,
      privateMarketsTax: privateMarkets.tax,
      privateMarketsVariance: privateMarkets.variance,
      privateMarketsAchievementPct: privateMarkets.achievementPct,
      directAndArmySales: directAndArmy.sales,
      directAndArmyTarget: directAndArmy.target,
      directAndArmyAchievementPct: directAndArmy.achievementPct,
      chainsSectorSales: chainsSector.sales,
      chainsSectorTarget: chainsSector.target,
      chainsSectorAchievementPct: chainsSector.achievementPct,
      unlinkedSectorSales: Number(unlinkedSectorSales.toFixed(2)),
      unlinkedSectorCount,
      unlinkedBranchSales: Number(unlinkedBranchRow.sales.toFixed(2)),
      unlinkedBranchCount: unlinkedBranchRow.cnt,
      totalSourceRecords: totalScopeRow.cnt,
      totalSourceSales: Number(totalScopeRow.sales.toFixed(2)),
      totalSourceTax: Number(totalScopeRow.tax.toFixed(2)),
    },
    sectors: sectorRows,
    aggregations,
    timeline,
    sourceComparison: {
      vf05n: vfComp,
      zsales: zsComp,
      salesDifference: salesDiff,
      taxDifference: taxDiff,
      isMatched: Math.abs(salesDiff) <= 0.05 && vfComp.count > 0 && zsComp.count > 0,
    },
  };
}

function classifyCommercialIndustrial(rawCategory: string | null | undefined): 'commercial' | 'industrial' | 'other' {
  const norm = normalizeArabic(rawCategory || '');
  if (!norm) return 'other';
  if (norm.includes('تجاري')) return 'commercial';
  if (norm.includes('صناعي')) return 'industrial';
  return 'other';
}

export function getDepartmentsChainsBreakdown(
  db: Database.Database,
  user: UserProfile,
  params: AnalyticsFilterParams
): DepartmentChainRow[] {
  const activeBatch = getActiveBatch(db);
  if (!activeBatch) return [];
  const scope = buildScopeFilter(user, params, activeBatch.id);

  const hasItemRow = db
    .prepare(`SELECT 1 AS ok FROM sales_records sr WHERE ${scope.sql} AND sr.item_code != '' LIMIT 1`)
    .get(...scope.args) as { ok: number } | undefined;
  const sourceHasItems = Boolean(hasItemRow);

  interface DeptChainDef {
    chain_code: string;
    chain_name: string;
    department_name: string;
    sector_name: string;
    supervisor_name: string;
  }

  const chainDefs = db
    .prepare(
      `SELECT chain_code, chain_name, department_name, sector_name, supervisor_name
       FROM departments_chains
       WHERE batch_id = ?`
    )
    .all(activeBatch.id) as DeptChainDef[];

  const chainByExactCode = new Map<string, DeptChainDef>();
  const chainByTrimmedCode = new Map<string, DeptChainDef>();
  const chainByNormName = new Map<string, DeptChainDef>();

  for (const dc of chainDefs) {
    const exactCode = normalizeCode(dc.chain_code);
    if (exactCode && !chainByExactCode.has(exactCode)) {
      chainByExactCode.set(exactCode, dc);
    }
    const trimmed = stripLeadingZeros(exactCode);
    if (trimmed && !chainByTrimmedCode.has(trimmed)) {
      chainByTrimmedCode.set(trimmed, dc);
    }
    const normName = normalizeArabic(dc.chain_name);
    if (normName && !chainByNormName.has(normName)) {
      chainByNormName.set(normName, dc);
    }
  }

  const resolveChainDef = (rawCode: string, rawName?: string | null): DeptChainDef | undefined => {
    const exact = normalizeCode(rawCode);
    if (exact) {
      const m1 = chainByExactCode.get(exact);
      if (m1) return m1;
      const m2 = chainByTrimmedCode.get(stripLeadingZeros(exact));
      if (m2) return m2;
    }
    const normN = normalizeArabic(rawName || '');
    if (normN) {
      const m3 = chainByNormName.get(normN);
      if (m3) return m3;
    }
    return undefined;
  };

  const rows = db
    .prepare(
      `SELECT
        COALESCE(sr.payer_code, '') AS payer_code,
        COALESCE(NULLIF(sr.resolved_chain_name, ''), NULLIF(sr.payer_name, ''), '') AS raw_chain_name,
        COALESCE(MAX(NULLIF(sr.resolved_department, '')), '') AS raw_department,
        COALESCE(sr.resolved_sector, 'غير مرتبط بقطاع') AS sector_name,
        COALESCE(NULLIF(TRIM(rep.supervisor_name), ''), '') AS rep_supervisor,
        COALESCE(it.category_name, '') AS item_category,
        COALESCE(SUM(sr.sales_value), 0) AS sales,
        COALESCE(SUM(sr.tax_value), 0) AS tax,
        COUNT(*) AS record_count
       FROM sales_records sr
       LEFT JOIN sales_reps rep
         ON rep.batch_id = sr.batch_id AND rep.rep_code = sr.resolved_rep_code
       LEFT JOIN items it
         ON it.batch_id = sr.batch_id AND it.item_code = sr.item_code
       WHERE ${scope.sql}
       GROUP BY sr.payer_code, COALESCE(NULLIF(sr.resolved_chain_name, ''), NULLIF(sr.payer_name, ''), ''), sr.resolved_sector, COALESCE(NULLIF(TRIM(rep.supervisor_name), ''), ''), COALESCE(it.category_name, '')`
    )
    .all(...scope.args) as Array<{
    payer_code: string;
    raw_chain_name: string;
    raw_department: string;
    sector_name: string;
    rep_supervisor: string;
    item_category: string;
    sales: number;
    tax: number;
    record_count: number;
  }>;

  // Consolidate rows per (supervisor_name, chain_code, sector_name)
  interface AggChain {
    chainCode: string;
    chainName: string;
    departmentName: string;
    sectorName: string;
    supervisorName: string;
    commercialSales: number;
    industrialSales: number;
    sales: number;
    tax: number;
    recordCount: number;
  }

  interface AggChainWithKey extends AggChain {
    codeSet: Set<string>;
  }

  const chainMap = new Map<string, AggChainWithKey>();
  const masterKeyByDef = new WeakMap<DeptChainDef, string>();

  const buildUnifiedChainKey = (supervisorName: string, chainName: string, fallbackCode: string): string => {
    const normSup = normalizeArabic(supervisorName) || 'غير محدد';
    const normChain = normalizeArabic(chainName);
    return `unified::${normSup}::${normChain || normalizeCode(fallbackCode) || fallbackCode}`;
  };

  const selectedSectorParam =
    params.sector && params.sector !== 'ALL' && params.sector !== 'undefined' && params.sector !== 'null'
      ? params.sector === 'UNLINKED'
        ? 'UNLINKED'
        : canonicalizeSectorName(params.sector)
      : 'ALL';

  // Master List: Pre-seed ALL chains/supervisors defined in `تعريف الإدارات والقطاعات` (filtered only by RBAC & sector filter)
  // consolidated by normalized chain name within each supervisor so multi-code chains (e.g. ال اس جى) appear as one unified row.
  for (const dc of chainDefs) {
    const dcSector = canonicalizeSectorName(dc.sector_name) || 'غير مرتبط بقطاع';
    if (user.role !== 'admin' && !scope.authorizedSectors.includes(dcSector)) continue;
    if (selectedSectorParam !== 'ALL') {
      if (selectedSectorParam === 'UNLINKED') {
        if (dcSector !== 'غير مرتبط بقطاع') continue;
      } else if (dcSector !== selectedSectorParam) {
        continue;
      }
    }

    const exactCode = normalizeCode(dc.chain_code) || dc.chain_code;
    const dcName = (dc.chain_name || exactCode).trim();
    const dcDept = dc.department_name || 'غير مصنف';
    const dcSupervisor = (dc.supervisor_name && dc.supervisor_name.trim()) || 'غير محدد';
    const masterKey = buildUnifiedChainKey(dcSupervisor, dcName, exactCode);

    const existing = chainMap.get(masterKey);
    if (!existing) {
      const codeSet = new Set<string>();
      if (exactCode) codeSet.add(exactCode);
      chainMap.set(masterKey, {
        chainCode: exactCode,
        chainName: dcName,
        departmentName: dcDept,
        sectorName: dcSector,
        supervisorName: dcSupervisor,
        commercialSales: 0,
        industrialSales: 0,
        sales: 0,
        tax: 0,
        recordCount: 0,
        codeSet,
      });
    } else if (exactCode && !existing.codeSet.has(exactCode)) {
      existing.codeSet.add(exactCode);
      existing.chainCode = Array.from(existing.codeSet).join(', ');
    }
    masterKeyByDef.set(dc, masterKey);
  }

  // Left-join fold sales_records into Master List (or add unmatched sales rows unified by supervisor + normalized chain name)
  for (const r of rows) {
    const dcMatch = resolveChainDef(r.payer_code, r.raw_chain_name);
    const resolvedCode = dcMatch?.chain_code || normalizeCode(r.payer_code) || r.payer_code;
    const resolvedName = (dcMatch?.chain_name || r.raw_chain_name || resolvedCode).trim();
    const resolvedDept = dcMatch?.department_name || r.raw_department || 'غير مصنف';
    const resolvedSector = dcMatch
      ? canonicalizeSectorName(dcMatch.sector_name) || r.sector_name || 'غير مرتبط بقطاع'
      : r.sector_name || 'غير مرتبط بقطاع';
    const resolvedSupervisor =
      (dcMatch?.supervisor_name && dcMatch.supervisor_name.trim()) ||
      (r.rep_supervisor && r.rep_supervisor.trim()) ||
      'غير محدد';

    const masterKey = dcMatch ? masterKeyByDef.get(dcMatch) : undefined;
    const key = masterKey ?? buildUnifiedChainKey(resolvedSupervisor, resolvedName, resolvedCode);

    let existing = chainMap.get(key);
    if (!existing) {
      const codeSet = new Set<string>();
      if (resolvedCode) codeSet.add(resolvedCode);
      existing = {
        chainCode: resolvedCode,
        chainName: resolvedName,
        departmentName: resolvedDept,
        sectorName: resolvedSector,
        supervisorName: resolvedSupervisor,
        commercialSales: 0,
        industrialSales: 0,
        sales: 0,
        tax: 0,
        recordCount: 0,
        codeSet,
      };
      chainMap.set(key, existing);
    } else {
      if (existing.supervisorName === 'غير محدد' && resolvedSupervisor !== 'غير محدد') {
        existing.supervisorName = resolvedSupervisor;
      }
      if (resolvedCode && !existing.codeSet.has(resolvedCode)) {
        existing.codeSet.add(resolvedCode);
        existing.chainCode = Array.from(existing.codeSet).join(', ');
      }
    }

    const cls = classifyCommercialIndustrial(r.item_category);
    if (cls === 'commercial') existing.commercialSales += r.sales;
    else if (cls === 'industrial') existing.industrialSales += r.sales;

    existing.sales += r.sales;
    existing.tax += r.tax;
    existing.recordCount += r.record_count;
  }

  // If the active source (e.g. VF05N) does not contain Material item codes, derive Commercial/Industrial split per unified chain from ZSALES
  if (!sourceHasItems && params.source === 'VF05N') {
    const zsScope = buildScopeFilter(user, { ...params, source: 'ZSALES' }, activeBatch.id);
    const zsRows = db
      .prepare(
        `SELECT
          COALESCE(sr.payer_code, '') AS payer_code,
          COALESCE(NULLIF(sr.resolved_chain_name, ''), NULLIF(sr.payer_name, ''), '') AS raw_chain_name,
          COALESCE(sr.resolved_sector, 'غير مرتبط بقطاع') AS sector_name,
          COALESCE(NULLIF(TRIM(rep.supervisor_name), ''), '') AS rep_supervisor,
          COALESCE(it.category_name, '') AS item_category,
          COALESCE(SUM(sr.sales_value), 0) AS sales
         FROM sales_records sr
         LEFT JOIN sales_reps rep
           ON rep.batch_id = sr.batch_id AND rep.rep_code = sr.resolved_rep_code
         LEFT JOIN items it
           ON it.batch_id = sr.batch_id AND it.item_code = sr.item_code
         WHERE ${zsScope.sql}
         GROUP BY sr.payer_code, COALESCE(NULLIF(sr.resolved_chain_name, ''), NULLIF(sr.payer_name, ''), ''), sr.resolved_sector, COALESCE(NULLIF(TRIM(rep.supervisor_name), ''), ''), COALESCE(it.category_name, '')`
      )
      .all(...zsScope.args) as Array<{
      payer_code: string;
      raw_chain_name: string;
      sector_name: string;
      rep_supervisor: string;
      item_category: string;
      sales: number;
    }>;

    const zsSplitMap = new Map<string, { commercial: number; industrial: number; total: number }>();
    for (const zr of zsRows) {
      const dcMatch = resolveChainDef(zr.payer_code, zr.raw_chain_name);
      const resolvedCode = dcMatch?.chain_code || normalizeCode(zr.payer_code) || zr.payer_code;
      const resolvedName = (dcMatch?.chain_name || zr.raw_chain_name || resolvedCode).trim();
      const resolvedSupervisor =
        (dcMatch?.supervisor_name && dcMatch.supervisor_name.trim()) ||
        (zr.rep_supervisor && zr.rep_supervisor.trim()) ||
        'غير محدد';
      const masterKey = dcMatch ? masterKeyByDef.get(dcMatch) : undefined;
      const k = masterKey ?? buildUnifiedChainKey(resolvedSupervisor, resolvedName, resolvedCode);

      const entry = zsSplitMap.get(k) ?? { commercial: 0, industrial: 0, total: 0 };
      const cls = classifyCommercialIndustrial(zr.item_category);
      if (cls === 'commercial') entry.commercial += zr.sales;
      else if (cls === 'industrial') entry.industrial += zr.sales;
      entry.total += zr.sales;
      zsSplitMap.set(k, entry);
    }

    for (const [key, chain] of chainMap.entries()) {
      if (chain.sales === 0) continue;
      const zs = zsSplitMap.get(key);
      if (zs && zs.total !== 0) {
        if (Math.abs(zs.total - chain.sales) <= 0.05) {
          chain.commercialSales = zs.commercial;
          chain.industrialSales = zs.industrial;
        } else {
          const ratio = chain.sales / zs.total;
          chain.commercialSales = zs.commercial * ratio;
          chain.industrialSales = zs.industrial * ratio;
        }
      }
    }
  }

  const sectorTotals = new Map<string, number>();
  const supervisorTotals = new Map<string, number>();
  for (const c of chainMap.values()) {
    sectorTotals.set(c.sectorName, (sectorTotals.get(c.sectorName) ?? 0) + c.sales);
    const normSup = normalizeArabic(c.supervisorName) || c.supervisorName;
    supervisorTotals.set(normSup, (supervisorTotals.get(normSup) ?? 0) + c.sales);
  }

  const result: DepartmentChainRow[] = Array.from(chainMap.values()).map((c) => {
    const secTotal = sectorTotals.get(c.sectorName) ?? 0;
    const normSup = normalizeArabic(c.supervisorName) || c.supervisorName;
    const supTotal = supervisorTotals.get(normSup) ?? 0;
    return {
      chainCode: c.chainCode,
      chainName: c.chainName,
      departmentName: c.departmentName,
      sectorName: c.sectorName,
      supervisorName: c.supervisorName,
      commercialSales: Number(c.commercialSales.toFixed(2)),
      industrialSales: Number(c.industrialSales.toFixed(2)),
      sales: Number(c.sales.toFixed(2)),
      tax: Number(c.tax.toFixed(2)),
      recordCount: c.recordCount,
      shareOfSectorPct: secTotal !== 0 ? Number(((c.sales / secTotal) * 100).toFixed(2)) : 0,
      shareOfSupervisorPct: supTotal !== 0 ? Number(((c.sales / supTotal) * 100).toFixed(1)) : 0,
    };
  });

  // Sort supervisors by total sales descending ('غير محدد' last), and chains within each supervisor by sales descending
  result.sort((a, b) => {
    if (a.supervisorName !== b.supervisorName) {
      if (a.supervisorName === 'غير محدد') return 1;
      if (b.supervisorName === 'غير محدد') return -1;
      const supDiff = (supervisorTotals.get(b.supervisorName) ?? 0) - (supervisorTotals.get(a.supervisorName) ?? 0);
      if (supDiff !== 0) return supDiff;
      return a.supervisorName.localeCompare(b.supervisorName, 'ar');
    }
    const salesDiff = b.sales - a.sales;
    if (salesDiff !== 0) return salesDiff;
    return a.chainName.localeCompare(b.chainName, 'ar');
  });

  return result;
}

export function getBranchesAnalytics(
  db: Database.Database,
  user: UserProfile,
  params: AnalyticsFilterParams
): {
  branches: BranchMetricRow[];
  unlinkedRecords: UnlinkedRecordRow[];
  totalUnlinkedSales: number;
  totalUnlinkedCount: number;
  diagnostics: BranchRepDiagnostics;
} {
  const activeBatch = getActiveBatch(db);
  if (!activeBatch) {
    return {
      branches: [],
      unlinkedRecords: [],
      totalUnlinkedSales: 0,
      totalUnlinkedCount: 0,
      diagnostics: {
        totalTargetedAssignedBranches: 0,
        previousInflatedCountInBranchesView: 0,
        differenceCount: 0,
        consolidatedDuplicateNameCount: 0,
        consolidatedDuplicateNames: [],
        repDiscrepancies: [],
      },
    };
  }

  const scope = buildScopeFilter(user, params, activeBatch.id);

  // Get branch master & targets defined in `تعريف الفروع و الاهداف`
  const definedBranches = db
    .prepare(
      `SELECT
        branch_code,
        branch_name,
        sector_name,
        branch_target,
        COALESCE(chain_code, '') AS chain_code,
        COALESCE(chain_name, '') AS chain_name,
        COALESCE(rep_code, '') AS rep_code,
        COALESCE(rep_name, '') AS rep_name,
        COALESCE(supervisor_name, '') AS supervisor_name
       FROM branches
       WHERE batch_id = ?`
    )
    .all(activeBatch.id) as Array<{
    branch_code: string;
    branch_name: string;
    sector_name: string;
    branch_target: number;
    chain_code: string;
    chain_name: string;
    rep_code: string;
    rep_name: string;
    supervisor_name: string;
  }>;

  // Load `departments_chains` lookup maps
  const chainDefs = db
    .prepare(
      `SELECT chain_code, chain_name, sector_name, COALESCE(supervisor_name, '') AS supervisor_name
       FROM departments_chains
       WHERE batch_id = ?`
    )
    .all(activeBatch.id) as Array<{
    chain_code: string;
    chain_name: string;
    sector_name: string;
    supervisor_name: string;
  }>;

  type BranchChainDef = (typeof chainDefs)[number] & { normName: string; canonicalSector: string };
  const chainByExactCode = new Map<string, BranchChainDef>();
  const chainByTrimmedCode = new Map<string, BranchChainDef>();
  const chainByNormName = new Map<string, BranchChainDef>();
  const chainsBySector = new Map<string, BranchChainDef[]>();
  const allNormalizedChains: BranchChainDef[] = [];

  for (const dc of chainDefs) {
    const normName = normalizeArabic(dc.chain_name);
    const canonicalSector = canonicalizeSectorName(dc.sector_name);
    const enriched: BranchChainDef = { ...dc, normName, canonicalSector };
    allNormalizedChains.push(enriched);

    const exact = normalizeCode(dc.chain_code);
    if (exact) {
      if (!chainByExactCode.has(exact)) chainByExactCode.set(exact, enriched);
      const trimmed = stripLeadingZeros(exact);
      if (trimmed && !chainByTrimmedCode.has(trimmed)) chainByTrimmedCode.set(trimmed, enriched);
    }
    if (normName && !chainByNormName.has(normName)) {
      chainByNormName.set(normName, enriched);
    }
    const secList = chainsBySector.get(canonicalSector) ?? [];
    secList.push(enriched);
    chainsBySector.set(canonicalSector, secList);
  }
  for (const list of chainsBySector.values()) {
    list.sort((a, b) => b.normName.length - a.normName.length);
  }
  allNormalizedChains.sort((a, b) => b.normName.length - a.normName.length);

  // Load `sales_reps` lookup maps
  const repRows = db
    .prepare(`SELECT rep_code, rep_name, COALESCE(supervisor_name, '') AS supervisor_name FROM sales_reps WHERE batch_id = ?`)
    .all(activeBatch.id) as Array<{ rep_code: string; rep_name: string; supervisor_name: string }>;
  const repByExactCode = new Map<string, (typeof repRows)[number]>();
  const repByTrimmedCode = new Map<string, (typeof repRows)[number]>();
  for (const r of repRows) {
    const exact = normalizeCode(r.rep_code);
    if (exact) {
      if (!repByExactCode.has(exact)) repByExactCode.set(exact, r);
      const trimmed = stripLeadingZeros(exact);
      if (trimmed && !repByTrimmedCode.has(trimmed)) repByTrimmedCode.set(trimmed, r);
    }
  }
  const resolveRepDef = (rawRepCode: string) => {
    const exact = normalizeCode(rawRepCode);
    if (!exact) return undefined;
    return repByExactCode.get(exact) ?? repByTrimmedCode.get(stripLeadingZeros(exact));
  };

  // Query sales grouped by (branch, rep) joining strictly on resolved_branch_code within date/source scope
  // Branch sector attribution comes from `تعريف الفروع و الاهداف` (branches.sector_name), not invoice Payer sector
  const branchScopeConditions: string[] = [`sr.batch_id = ?`, `sr.source_type = ?`, `sr.resolved_branch_code IS NOT NULL`];
  const branchScopeArgs: unknown[] = [activeBatch.id, params.source];
  if (params.startDate) {
    branchScopeConditions.push(`sr.billing_date >= ?`);
    branchScopeArgs.push(params.startDate);
  }
  if (params.endDate) {
    branchScopeConditions.push(`sr.billing_date <= ?`);
    branchScopeArgs.push(params.endDate);
  }

  const salesRows = db
    .prepare(
      `SELECT
        sr.resolved_branch_code AS branch_code,
        COALESCE(NULLIF(sr.resolved_rep_code, ''), '') AS rep_code,
        COALESCE(MAX(NULLIF(sr.resolved_rep_name, '')), '') AS rep_name,
        COALESCE(MAX(NULLIF(sr.payer_code, '')), '') AS payer_code,
        COALESCE(MAX(NULLIF(sr.resolved_chain_name, '')), '') AS resolved_chain_name,
        COALESCE(SUM(sr.sales_value), 0) AS sales,
        COALESCE(SUM(sr.tax_value), 0) AS tax,
        COUNT(*) AS record_count,
        SUM(CASE WHEN sr.is_return = 1 THEN 1 ELSE 0 END) AS return_count
       FROM sales_records sr
       WHERE ${branchScopeConditions.join(' AND ')}
       GROUP BY sr.resolved_branch_code, COALESCE(NULLIF(sr.resolved_rep_code, ''), '')`
    )
    .all(...branchScopeArgs) as Array<{
    branch_code: string;
    rep_code: string;
    rep_name: string;
    payer_code: string;
    resolved_chain_name: string;
    sales: number;
    tax: number;
    record_count: number;
    return_count: number;
  }>;

  interface BranchAggBucket {
    sales: number;
    tax: number;
    recordCount: number;
    returnCount: number;
    payerCode: string;
    chainName: string;
    reps: Map<string, BranchRepBreakdown>;
  }
  const branchSalesMap = new Map<string, BranchAggBucket>();
  const branchSalesTrimmedMap = new Map<string, BranchAggBucket>();

  for (const row of salesRows) {
    const normBranchCode = normalizeCode(row.branch_code) || row.branch_code;
    const trimmedBranchCode = stripLeadingZeros(normBranchCode);
    let bucket = branchSalesMap.get(normBranchCode);
    if (!bucket) {
      bucket = {
        sales: 0,
        tax: 0,
        recordCount: 0,
        returnCount: 0,
        payerCode: row.payer_code,
        chainName: row.resolved_chain_name,
        reps: new Map(),
      };
      branchSalesMap.set(normBranchCode, bucket);
      if (trimmedBranchCode && !branchSalesTrimmedMap.has(trimmedBranchCode)) {
        branchSalesTrimmedMap.set(trimmedBranchCode, bucket);
      }
    }
    bucket.sales += row.sales;
    bucket.tax += row.tax;
    bucket.recordCount += row.record_count;
    bucket.returnCount += row.return_count;
    if (!bucket.payerCode && row.payer_code) bucket.payerCode = row.payer_code;
    if (!bucket.chainName && row.resolved_chain_name) bucket.chainName = row.resolved_chain_name;

    const cleanRepCode = normalizeCode(row.rep_code);
    if (cleanRepCode) {
      const repDef = resolveRepDef(cleanRepCode);
      const repDisplay = (repDef?.rep_name || row.rep_name || cleanRepCode).trim();
      const existingRep = bucket.reps.get(cleanRepCode) ?? {
        repCode: cleanRepCode,
        repName: repDisplay,
        sales: 0,
        tax: 0,
        recordCount: 0,
      };
      existingRep.sales += row.sales;
      existingRep.tax += row.tax;
      existingRep.recordCount += row.record_count;
      bucket.reps.set(cleanRepCode, existingRep);
    }
  }

  // Batch-wide fallback payer lookup per branch (for legacy batches where branches.chain_code is empty or date filter/ZSALES hides VF05N rows)
  const batchBranchChainFallback = new Map<string, { payerCode: string; chainName: string }>();
  if (definedBranches.some((b) => !b.chain_code)) {
    const fbRows = db
      .prepare(
        `SELECT
          resolved_branch_code AS branch_code,
          MAX(NULLIF(payer_code, '')) AS payer_code,
          MAX(NULLIF(resolved_chain_name, '')) AS chain_name
         FROM sales_records
         WHERE batch_id = ? AND source_type = 'VF05N' AND resolved_branch_code IS NOT NULL
         GROUP BY resolved_branch_code`
      )
      .all(activeBatch.id) as Array<{ branch_code: string; payer_code: string | null; chain_name: string | null }>;
    for (const r of fbRows) {
      if (r.payer_code || r.chain_name) {
        batchBranchChainFallback.set(r.branch_code, {
          payerCode: r.payer_code ?? '',
          chainName: r.chain_name ?? '',
        });
      }
    }
  }

  const resolveChainForBranch = (
    branchCode: string,
    branchName: string,
    sector: string,
    storedChainCode: string,
    storedChainName: string,
    salesPayerCode?: string,
    salesChainName?: string
  ): BranchChainDef | undefined => {
    const lookupByCode = (code: string) => {
      const exact = normalizeCode(code);
      if (!exact) return undefined;
      return chainByExactCode.get(exact) ?? chainByTrimmedCode.get(stripLeadingZeros(exact));
    };

    if (storedChainCode) {
      const m = lookupByCode(storedChainCode);
      if (m) return m;
    }
    if (storedChainName) {
      const m = chainByNormName.get(normalizeArabic(storedChainName));
      if (m) return m;
    }
    if (salesPayerCode) {
      const m = lookupByCode(salesPayerCode);
      if (m) return m;
    }
    if (salesChainName) {
      const m = chainByNormName.get(normalizeArabic(salesChainName));
      if (m) return m;
    }
    const selfMatch = lookupByCode(branchCode);
    if (selfMatch) return selfMatch;

    // Prefix/alias fallback on normalized branch name for legacy batches without chain_code
    const normBranch = normalizeArabic(branchName);
    if (!normBranch) return undefined;

    const candidates = chainsBySector.get(sector) ?? allNormalizedChains;
    for (const c of candidates) {
      if (c.normName && (normBranch === c.normName || normBranch.startsWith(`${c.normName} `))) {
        return c;
      }
    }
    // Common Arabic spelling variations between SAP branch names & chain sheet (e.g. سبنيس vs سبينس, سعودى vs سعودى ماركت)
    const firstWord = normBranch.split(' ')[0];
    if (firstWord && firstWord.length >= 4) {
      for (const c of candidates) {
        if (c.normName && (c.normName === firstWord || c.normName.startsWith(`${firstWord} `))) {
          return c;
        }
      }
      if (firstWord === 'سبنيس' || firstWord === 'سبينيس' || firstWord === 'سبينس') {
        const sp = candidates.find((c) => c.normName === 'سبينس' || c.normName === 'سبينيس' || c.normName === 'سبنيس');
        if (sp) return sp;
      }
    }
    return undefined;
  };

  interface ConsolidatedBranchBucket {
    codeSet: Set<string>;
    branchName: string;
    sectorName: string;
    supervisorName: string;
    chainCode: string;
    chainName: string;
    assignedRepCode: string;
    assignedRepName: string;
    repMap: Map<string, BranchRepBreakdown>;
    target: number;
    sales: number;
    tax: number;
    recordCount: number;
    returnCount: number;
  }

  const seenExactCodes = new Set<string>();
  const consolidatedMap = new Map<string, ConsolidatedBranchBucket>();
  const selectedSector =
    params.sector && params.sector !== 'ALL' && params.sector !== 'undefined' && params.sector !== 'null'
      ? canonicalizeSectorName(params.sector)
      : 'ALL';

  for (const b of definedBranches) {
    const sector = canonicalizeSectorName(b.sector_name || 'التوريد المباشر');
    if (user.role !== 'admin' && !scope.authorizedSectors.includes(sector)) continue;
    if (selectedSector !== 'ALL' && selectedSector !== sector) continue;

    // Preserve leading zeros in branch_code and deduplicate exact code rows
    const normBranchCode = normalizeCode(b.branch_code) || String(b.branch_code ?? '').trim();
    if (!normBranchCode || seenExactCodes.has(normBranchCode)) continue;
    seenExactCodes.add(normBranchCode);

    const s =
      branchSalesMap.get(normBranchCode) ?? branchSalesTrimmedMap.get(stripLeadingZeros(normBranchCode));
    const batchFb = batchBranchChainFallback.get(b.branch_code);
    const sales = s?.sales ?? 0;
    const tax = s?.tax ?? 0;
    const target = Number(b.branch_target) || 0;

    const hasStoredChain = Boolean(b.chain_code.trim() || b.chain_name.trim());
    const chainDef = resolveChainForBranch(
      b.branch_code,
      b.branch_name,
      sector,
      b.chain_code,
      b.chain_name,
      hasStoredChain ? undefined : s?.payerCode || batchFb?.payerCode,
      hasStoredChain ? undefined : s?.chainName || batchFb?.chainName
    );
    const assignedRepCode = normalizeCode(b.rep_code);
    const assignedRepDef = assignedRepCode ? resolveRepDef(assignedRepCode) : undefined;
    const assignedRepName = (b.rep_name || assignedRepDef?.rep_name || assignedRepCode || '').trim();

    // Always prefer `كود الادارة` / `اسم الادارة` recorded per branch in `تعريف الفروع و الاهداف`
    const chainCode = normalizeCode(b.chain_code) || chainDef?.chain_code || s?.payerCode || batchFb?.payerCode || '';
    const chainName =
      (b.chain_name.trim() || chainDef?.chain_name || s?.chainName || batchFb?.chainName || '').trim() || 'غير محدد';
    // Supervisor MUST come strictly from `اسم المشرف` on that branch row in `تعريف الفروع و الاهداف`
    const supervisorName = b.supervisor_name.trim() || 'غير محدد';

    const cleanBranchName = (b.branch_name || '').trim() || normBranchCode;
    const normBranchName = normalizeArabic(cleanBranchName);
    const normSup = normalizeArabic(supervisorName);
    // Consolidate distinct codes sharing the same normalized branch name under the same (sector, supervisor, assignedRep)
    const unifiedKey = normBranchName
      ? `${sector}::${normSup}::${assignedRepCode}::${normBranchName}`
      : `code::${normBranchCode}`;

    let bucket = consolidatedMap.get(unifiedKey);
    if (!bucket) {
      bucket = {
        codeSet: new Set([normBranchCode]),
        branchName: cleanBranchName,
        sectorName: sector,
        supervisorName,
        chainCode,
        chainName,
        assignedRepCode,
        assignedRepName,
        repMap: new Map(),
        target: 0,
        sales: 0,
        tax: 0,
        recordCount: 0,
        returnCount: 0,
      };
      consolidatedMap.set(unifiedKey, bucket);
    } else {
      bucket.codeSet.add(normBranchCode);
      if (bucket.chainName === 'غير محدد' && chainName !== 'غير محدد') {
        bucket.chainName = chainName;
        bucket.chainCode = chainCode;
      }
      if (!bucket.assignedRepName && assignedRepName) {
        bucket.assignedRepName = assignedRepName;
      }
    }

    bucket.target += target;
    bucket.sales += sales;
    bucket.tax += tax;
    bucket.recordCount += s?.recordCount ?? 0;
    bucket.returnCount += s?.returnCount ?? 0;

    if (s) {
      for (const rb of s.reps.values()) {
        const existingRb = bucket.repMap.get(rb.repCode);
        if (!existingRb) {
          bucket.repMap.set(rb.repCode, { ...rb });
        } else {
          existingRb.sales += rb.sales;
          existingRb.tax += rb.tax;
          existingRb.recordCount += rb.recordCount;
        }
      }
    }
  }

  const branches: BranchMetricRow[] = [];
  const consolidatedDuplicateNames: BranchRepDiagnostics['consolidatedDuplicateNames'] = [];

  for (const bucket of consolidatedMap.values()) {
    const sales = Number(bucket.sales.toFixed(2));
    const tax = Number(bucket.tax.toFixed(2));
    const target = Number(bucket.target.toFixed(2));
    const codes = Array.from(bucket.codeSet);

    if (codes.length > 1) {
      consolidatedDuplicateNames.push({
        branchName: bucket.branchName,
        sectorName: bucket.sectorName,
        supervisorName: bucket.supervisorName,
        repCode: bucket.assignedRepCode,
        branchCodes: codes,
        combinedSales: sales,
        combinedTarget: target,
      });
    }

    const repBreakdowns: BranchRepBreakdown[] = Array.from(bucket.repMap.values())
      .map((rb) => ({
        repCode: rb.repCode,
        repName: rb.repName,
        sales: Number(rb.sales.toFixed(2)),
        tax: Number(rb.tax.toFixed(2)),
        recordCount: rb.recordCount,
      }))
      .sort((a, bRep) => bRep.sales - a.sales);

    branches.push({
      branchCode: codes.join(', '),
      branchName: bucket.branchName,
      sectorName: bucket.sectorName,
      supervisorName: bucket.supervisorName,
      chainCode: bucket.chainCode,
      chainName: bucket.chainName,
      assignedRepCode: bucket.assignedRepCode,
      assignedRepName: bucket.assignedRepName,
      repBreakdowns,
      target,
      sales,
      tax,
      variance: Number((sales - target).toFixed(2)),
      achievementPct: target > 0 ? calculateAchievementPct(sales, target, tax) : null,
      recordCount: bucket.recordCount,
      returnCount: bucket.returnCount,
    });
  }

  branches.sort((a, b) => {
    const chainCmp = a.chainName.localeCompare(b.chainName, 'ar');
    if (chainCmp !== 0) return chainCmp;
    return a.branchName.localeCompare(b.branchName, 'ar');
  });

  // Diagnostics comparing strict assigned branches (target > 0) vs cross-rep invoice matches
  const repDiagMap = new Map<
    string,
    {
      repCode: string;
      repName: string;
      expectedAssignedTargetedCount: number;
      extraCrossRepBranches: BranchRepDiagnostics['repDiscrepancies'][number]['extraCrossRepBranches'];
    }
  >();

  for (const br of branches) {
    if (br.target <= 0) continue;
    const assignedKey = normalizeCode(br.assignedRepCode);
    if (assignedKey) {
      const entry = repDiagMap.get(assignedKey) ?? {
        repCode: assignedKey,
        repName: br.assignedRepName || assignedKey,
        expectedAssignedTargetedCount: 0,
        extraCrossRepBranches: [],
      };
      entry.expectedAssignedTargetedCount += 1;
      repDiagMap.set(assignedKey, entry);
    }

    for (const rb of br.repBreakdowns) {
      const rbKey = normalizeCode(rb.repCode);
      if (!rbKey || rbKey === assignedKey) continue;
      if (!assignedKey) continue;
      const entry = repDiagMap.get(rbKey) ?? {
        repCode: rbKey,
        repName: rb.repName || rbKey,
        expectedAssignedTargetedCount: 0,
        extraCrossRepBranches: [],
      };
      entry.extraCrossRepBranches.push({
        branchCode: br.branchCode,
        branchName: br.branchName,
        target: br.target,
        assignedRepCode: assignedKey,
        assignedRepName: br.assignedRepName || assignedKey,
        supervisorName: br.supervisorName,
      });
      repDiagMap.set(rbKey, entry);
    }
  }

  let totalTargetedAssignedBranches = 0;
  let previousInflatedCountInBranchesView = 0;
  const repDiscrepancies: BranchRepDiagnostics['repDiscrepancies'] = [];

  for (const entry of repDiagMap.values()) {
    if (entry.expectedAssignedTargetedCount === 0 && entry.extraCrossRepBranches.length === 0) continue;
    const prevCount = entry.expectedAssignedTargetedCount + entry.extraCrossRepBranches.length;
    if (entry.expectedAssignedTargetedCount > 0) {
      totalTargetedAssignedBranches += entry.expectedAssignedTargetedCount;
      previousInflatedCountInBranchesView += prevCount;
    }
    if (entry.extraCrossRepBranches.length > 0) {
      repDiscrepancies.push({
        repCode: entry.repCode,
        repName: entry.repName,
        expectedAssignedTargetedCount: entry.expectedAssignedTargetedCount,
        previousBranchesViewCount: prevCount,
        extraCount: entry.extraCrossRepBranches.length,
        extraCrossRepBranches: entry.extraCrossRepBranches,
      });
    }
  }

  const diagnostics: BranchRepDiagnostics = {
    totalTargetedAssignedBranches,
    previousInflatedCountInBranchesView,
    differenceCount: previousInflatedCountInBranchesView - totalTargetedAssignedBranches,
    consolidatedDuplicateNameCount: consolidatedDuplicateNames.length,
    consolidatedDuplicateNames,
    repDiscrepancies,
  };

  const unlinkedRows = db
    .prepare(
      `SELECT
        sr.id, sr.doc_number, sr.billing_date, sr.payer_code, sr.payer_name,
        sr.sold_to_code, sr.sold_to_name, sr.resolved_sector,
        sr.sales_value, sr.tax_value, sr.item_code, sr.item_name
       FROM sales_records sr
       WHERE ${scope.sql} AND sr.resolved_branch_code IS NULL
       ORDER BY sr.sales_value DESC`
    )
    .all(...scope.args) as Array<{
    id: number;
    doc_number: string;
    billing_date: string;
    payer_code: string;
    payer_name: string;
    sold_to_code: string;
    sold_to_name: string;
    resolved_sector: string | null;
    sales_value: number;
    tax_value: number;
    item_code: string;
    item_name: string;
  }>;

  let totalUnlinkedSales = 0;
  const unlinkedRecords: UnlinkedRecordRow[] = unlinkedRows.map((r) => {
    totalUnlinkedSales += r.sales_value;
    return {
      id: r.id,
      docNumber: r.doc_number,
      billingDate: r.billing_date,
      payerCode: r.payer_code,
      payerName: r.payer_name,
      soldToCode: r.sold_to_code,
      soldToName: r.sold_to_name,
      resolvedSector: r.resolved_sector,
      salesValue: Number(r.sales_value.toFixed(2)),
      taxValue: Number(r.tax_value.toFixed(2)),
      itemCode: r.item_code,
      itemName: r.item_name,
    };
  });

  return {
    branches,
    unlinkedRecords,
    totalUnlinkedSales: Number(totalUnlinkedSales.toFixed(2)),
    totalUnlinkedCount: unlinkedRecords.length,
    diagnostics,
  };
}

export function getRepsAnalytics(
  db: Database.Database,
  user: UserProfile,
  params: AnalyticsFilterParams
): RepMetricRow[] {
  const activeBatch = getActiveBatch(db);
  if (!activeBatch) return [];
  const scope = buildScopeFilter(user, params, activeBatch.id);

  const selectedSector =
    params.sector && params.sector !== 'ALL' && params.sector !== 'undefined' && params.sector !== 'null'
      ? params.sector === 'UNLINKED'
        ? 'UNLINKED'
        : canonicalizeSectorName(params.sector)
      : 'ALL';

  const repDefs = db
    .prepare(
      `SELECT rep_code, rep_name, sector_name, supervisor_name, COALESCE(monthly_target, 0) AS monthly_target
       FROM sales_reps
       WHERE batch_id = ?`
    )
    .all(activeBatch.id) as Array<{
    rep_code: string;
    rep_name: string;
    sector_name: string;
    supervisor_name: string;
    monthly_target: number;
  }>;

  const repDefByExact = new Map<string, (typeof repDefs)[number]>();
  const repDefByTrimmed = new Map<string, (typeof repDefs)[number]>();
  for (const d of repDefs) {
    const exact = normalizeCode(d.rep_code);
    if (exact) {
      if (!repDefByExact.has(exact)) repDefByExact.set(exact, d);
      const trimmed = stripLeadingZeros(exact);
      if (trimmed && !repDefByTrimmed.has(trimmed)) repDefByTrimmed.set(trimmed, d);
    }
  }
  const resolveRepDef = (code: string) => {
    const exact = normalizeCode(code);
    if (!exact) return undefined;
    return repDefByExact.get(exact) ?? repDefByTrimmed.get(stripLeadingZeros(exact));
  };

  // Aggregate branch targets & unique targeted branches (branch_target > 0) per rep within authorized/selected sectors
  // Deduplicates by BOTH branch_code (preserving leading zeros) AND normalized branch_name per rep/supervisor/sector
  const branchRows = db
    .prepare(
      `SELECT
        branch_code,
        COALESCE(branch_name, '') AS branch_name,
        COALESCE(rep_code, '') AS rep_code,
        sector_name,
        branch_target,
        COALESCE(supervisor_name, '') AS supervisor_name
       FROM branches
       WHERE batch_id = ? AND COALESCE(rep_code, '') != ''`
    )
    .all(activeBatch.id) as Array<{
    branch_code: string;
    branch_name: string;
    rep_code: string;
    sector_name: string;
    branch_target: number;
    supervisor_name: string;
  }>;

  interface RepBranchMeta {
    target: number;
    seenBranchCodes: Set<string>;
    branchGroups: Map<string, { target: number }>;
    targetBranchKeys: Set<string>;
    supervisors: Set<string>;
    sectors: Set<string>;
  }
  const repBranchMetaMap = new Map<string, RepBranchMeta>();
  const repBranchMetaTrimmedMap = new Map<string, RepBranchMeta>();

  if (selectedSector !== 'UNLINKED') {
    for (const b of branchRows) {
      const sec = canonicalizeSectorName(b.sector_name || '');
      if (user.role !== 'admin' && !scope.authorizedSectors.includes(sec)) continue;
      if (selectedSector !== 'ALL' && sec !== selectedSector) continue;

      const normRep = normalizeCode(b.rep_code);
      if (!normRep) continue;
      const defMatch = resolveRepDef(normRep);
      const canonicalRepKey = defMatch ? normalizeCode(defMatch.rep_code) : normRep;
      const trimmedKey = stripLeadingZeros(canonicalRepKey);

      let meta = repBranchMetaMap.get(canonicalRepKey);
      if (!meta) {
        meta = {
          target: 0,
          seenBranchCodes: new Set(),
          branchGroups: new Map(),
          targetBranchKeys: new Set(),
          supervisors: new Set(),
          sectors: new Set(),
        };
        repBranchMetaMap.set(canonicalRepKey, meta);
        if (trimmedKey && !repBranchMetaTrimmedMap.has(trimmedKey)) {
          repBranchMetaTrimmedMap.set(trimmedKey, meta);
        }
      }

      // Preserve leading zeros in branch_code and ignore exact duplicate code rows for same rep
      const normBranchCode = normalizeCode(b.branch_code) || String(b.branch_code ?? '').trim();
      if (!normBranchCode || meta.seenBranchCodes.has(normBranchCode)) continue;
      meta.seenBranchCodes.add(normBranchCode);

      const normBranchName = normalizeArabic(b.branch_name);
      const normSup = normalizeArabic(b.supervisor_name);
      const groupKey = normBranchName
        ? `${sec}::${normSup}::${normBranchName}`
        : `code::${normBranchCode}`;

      const targetVal = Number(b.branch_target);
      const safeTarget = Number.isFinite(targetVal) && targetVal > 0 ? targetVal : 0;

      const existingGroup = meta.branchGroups.get(groupKey) ?? { target: 0 };
      existingGroup.target += safeTarget;
      meta.branchGroups.set(groupKey, existingGroup);

      if (sec) meta.sectors.add(sec);
      if (b.supervisor_name && b.supervisor_name.trim()) {
        meta.supervisors.add(b.supervisor_name.trim());
      }
    }

    for (const meta of repBranchMetaMap.values()) {
      let sumTarget = 0;
      for (const [groupKey, group] of meta.branchGroups.entries()) {
        if (group.target > 0) {
          sumTarget += group.target;
          meta.targetBranchKeys.add(groupKey);
        }
      }
      meta.target = sumTarget;
    }
  }

  const rows = db
    .prepare(
      `SELECT
        COALESCE(NULLIF(sr.resolved_rep_code, ''), 'UNLINKED') AS rep_code,
        MAX(sr.resolved_rep_name) AS resolved_rep_name,
        MAX(sr.resolved_sector) AS sector_name,
        COALESCE(SUM(sr.sales_value), 0) AS sales,
        COALESCE(SUM(sr.tax_value), 0) AS tax,
        COUNT(*) AS record_count,
        SUM(CASE WHEN sr.is_return = 1 THEN 1 ELSE 0 END) AS return_count
       FROM sales_records sr
       WHERE ${scope.sql}
       GROUP BY COALESCE(NULLIF(sr.resolved_rep_code, ''), 'UNLINKED')
       ORDER BY sales DESC`
    )
    .all(...scope.args) as Array<{
    rep_code: string;
    resolved_rep_name: string | null;
    sector_name: string | null;
    sales: number;
    tax: number;
    record_count: number;
    return_count: number;
  }>;

  interface AggRepBucket {
    repCode: string;
    repName: string;
    sectorName: string;
    supervisorName: string;
    sales: number;
    tax: number;
    target: number;
    recordCount: number;
    returnCount: number;
    branchCount: number;
  }

  const repBuckets = new Map<string, AggRepBucket>();

  for (const r of rows) {
    const isUnlinked = r.rep_code === 'UNLINKED';
    const def = isUnlinked ? undefined : resolveRepDef(r.rep_code);
    const key = isUnlinked ? 'UNLINKED' : normalizeCode(def?.rep_code || r.rep_code) || r.rep_code;

    const existing = repBuckets.get(key);
    if (!existing) {
      const branchMeta = isUnlinked
        ? undefined
        : repBranchMetaMap.get(key) ?? repBranchMetaTrimmedMap.get(stripLeadingZeros(key));

      const branchSup =
        branchMeta && branchMeta.supervisors.size > 0
          ? Array.from(branchMeta.supervisors).join('، ')
          : '';
      const defSector = def?.sector_name ? canonicalizeSectorName(def.sector_name) : '';
      const defTargetMatchesScope =
        def &&
        def.monthly_target > 0 &&
        selectedSector !== 'UNLINKED' &&
        (user.role === 'admin' || !defSector || scope.authorizedSectors.includes(defSector)) &&
        (selectedSector === 'ALL' || !defSector || defSector === selectedSector);
      const targetVal =
        branchMeta && branchMeta.target > 0
          ? branchMeta.target
          : defTargetMatchesScope
            ? def!.monthly_target
            : 0;

      repBuckets.set(key, {
        repCode: isUnlinked ? 'غير محدد' : key,
        repName: isUnlinked
          ? 'سجلات بدون كود مندوب'
          : def?.rep_name || r.resolved_rep_name || r.rep_code,
        sectorName: defSector || r.sector_name || 'غير محدد',
        supervisorName: def?.supervisor_name?.trim() || branchSup || '-',
        sales: r.sales,
        tax: r.tax,
        target: targetVal,
        recordCount: r.record_count,
        returnCount: r.return_count,
        branchCount: branchMeta ? branchMeta.targetBranchKeys.size : 0,
      });
    } else {
      existing.sales += r.sales;
      existing.tax += r.tax;
      existing.recordCount += r.record_count;
      existing.returnCount += r.return_count;
    }
  }

  // Ensure defined reps with targets in the current sector scope stay visible even if 0 sales in a narrowed date range
  if (selectedSector !== 'UNLINKED') {
    for (const d of repDefs) {
      const key = normalizeCode(d.rep_code);
      if (!key || repBuckets.has(key)) continue;
      const branchMeta = repBranchMetaMap.get(key) ?? repBranchMetaTrimmedMap.get(stripLeadingZeros(key));
      const defSector = d.sector_name ? canonicalizeSectorName(d.sector_name) : '';
      const inAuth = user.role === 'admin' || !defSector || scope.authorizedSectors.includes(defSector);
      const inSelected = selectedSector === 'ALL' || (defSector && defSector === selectedSector);
      const hasBranchTarget = (branchMeta?.target ?? 0) > 0;
      const hasDefTarget = d.monthly_target > 0 && inAuth && inSelected;
      if (!hasBranchTarget && !hasDefTarget) continue;

      const branchSup =
        branchMeta && branchMeta.supervisors.size > 0
          ? Array.from(branchMeta.supervisors).join('، ')
          : '';
      const targetVal = hasBranchTarget ? branchMeta!.target : d.monthly_target;
      repBuckets.set(key, {
        repCode: key,
        repName: d.rep_name || key,
        sectorName: defSector || (branchMeta ? Array.from(branchMeta.sectors)[0] : '') || 'غير محدد',
        supervisorName: d.supervisor_name?.trim() || branchSup || '-',
        sales: 0,
        tax: 0,
        target: targetVal,
        recordCount: 0,
        returnCount: 0,
        branchCount: branchMeta ? branchMeta.targetBranchKeys.size : 0,
      });
    }
  }

  const result: RepMetricRow[] = Array.from(repBuckets.values()).map((b) => {
    const sales = Number(b.sales.toFixed(2));
    const tax = Number(b.tax.toFixed(2));
    const target = Number(b.target.toFixed(2));
    return {
      repCode: b.repCode,
      repName: b.repName,
      sectorName: b.sectorName,
      supervisorName: b.supervisorName,
      sales,
      tax,
      target,
      achievementPct: calculateAchievementPct(sales, target, tax),
      recordCount: b.recordCount,
      returnCount: b.returnCount,
      branchCount: Math.trunc(b.branchCount),
    };
  });

  result.sort((a, b) => b.sales - a.sales);
  return result;
}

export function getItemsAnalytics(
  db: Database.Database,
  user: UserProfile,
  params: AnalyticsFilterParams
): ItemMetricRow[] {
  const activeBatch = getActiveBatch(db);
  if (!activeBatch) return [];
  const scope = buildScopeFilter(user, params, activeBatch.id);

  const itemDefs = db
    .prepare(`SELECT item_code, item_name, category_name, brand_name FROM items WHERE batch_id = ?`)
    .all(activeBatch.id) as Array<{
    item_code: string;
    item_name: string;
    category_name: string;
    brand_name: string;
  }>;
  const itemDefMap = new Map(itemDefs.map((it) => [it.item_code, it]));

  const rows = db
    .prepare(
      `SELECT
        COALESCE(NULLIF(sr.item_code, ''), 'UNLINKED') AS item_code,
        MAX(sr.resolved_item_name) AS resolved_item_name,
        MAX(sr.item_name) AS raw_item_name,
        COALESCE(SUM(sr.quantity), 0) AS quantity,
        COALESCE(SUM(sr.sales_value), 0) AS sales,
        COALESCE(SUM(sr.tax_value), 0) AS tax,
        COUNT(*) AS record_count,
        SUM(CASE WHEN sr.is_return = 1 THEN 1 ELSE 0 END) AS return_count,
        COALESCE(SUM(CASE WHEN sr.is_return = 1 THEN sr.sales_value ELSE 0 END), 0) AS return_value
       FROM sales_records sr
       WHERE ${scope.sql}
       GROUP BY COALESCE(NULLIF(sr.item_code, ''), 'UNLINKED')
       ORDER BY sales DESC`
    )
    .all(...scope.args) as Array<{
    item_code: string;
    resolved_item_name: string | null;
    raw_item_name: string | null;
    quantity: number;
    sales: number;
    tax: number;
    record_count: number;
    return_count: number;
    return_value: number;
  }>;

  return rows.map((r) => {
    const def = itemDefMap.get(r.item_code);
    return {
      itemCode: r.item_code === 'UNLINKED' ? 'غير محدد' : r.item_code,
      itemName:
        r.item_code === 'UNLINKED'
          ? 'سجلات بدون كود صنف'
          : def?.item_name || r.resolved_item_name || r.raw_item_name || r.item_code,
      categoryName: def?.category_name || '-',
      brandName: def?.brand_name || '-',
      quantity: Number(r.quantity.toFixed(2)),
      sales: Number(r.sales.toFixed(2)),
      tax: Number(r.tax.toFixed(2)),
      recordCount: r.record_count,
      returnCount: r.return_count,
      returnValue: Number(r.return_value.toFixed(2)),
    };
  });
}

function isSalesRepJobTitle(jobTitle: string): boolean {
  const norm = normalizeArabic(jobTitle);
  return norm.includes('مندوب');
}

function isSupervisorJobTitle(jobTitle: string): boolean {
  const norm = normalizeArabic(jobTitle);
  return norm.includes('مشرف');
}

function isDepartmentManagerJobTitle(jobTitle: string): boolean {
  const noSpace = normalizeArabic(jobTitle).replace(/\s+/g, '');
  return (
    noSpace === 'مديراداره' ||
    noSpace === 'مديرالاداره' ||
    noSpace.includes('مديراداره') ||
    noSpace.includes('مديرالاداره')
  );
}

function isRafatMohamedName(employeeName: string): boolean {
  return normalizeArabic(employeeName).replace(/\s+/g, '').includes('رافتمحمد');
}

function isAbdelSattarName(employeeName: string): boolean {
  return normalizeArabic(employeeName).replace(/\s+/g, '').includes('عبدالستارمصطفي');
}

function resolveDepartmentManagerSectors(employeeName: string, rawSectorName: string): string[] {
  if (isRafatMohamedName(employeeName)) {
    return ['التوريد المركزي'];
  }
  if (isAbdelSattarName(employeeName)) {
    return ['التوريد المباشر', 'الجيش', 'التوب ريتيل'];
  }
  const normSec = normalizeArabic(rawSectorName);
  const matched: string[] = [];
  if (normSec.includes('مباشر')) matched.push('التوريد المباشر');
  if (normSec.includes('مركزي')) matched.push('التوريد المركزي');
  if ((normSec.includes('توب') && normSec.includes('ريتيل')) || (normSec.includes('تجزئ') && normSec.includes('مميز'))) {
    matched.push('التوب ريتيل');
  }
  if (normSec.includes('جيش')) matched.push('الجيش');
  if (matched.length > 0) return matched;
  const canon = rawSectorName ? canonicalizeSectorName(rawSectorName) : '';
  return canon ? [canon] : [...ORIGINAL_SECTORS];
}

function normalizeCategoryText(raw: string): string {
  return normalizeArabic(raw).replace(/\s+/g, '');
}

const UNCLASSIFIED_CAT_NORM = normalizeCategoryText('غير مصنف');
const ABDEL_SATTAR_SECTOR_LABEL = 'التوريد المباشر + الجيش + التوب ريتيل';

export function buildItemCategoryTargetKey(
  sectorName: string,
  jobTitle: string,
  employeeName: string,
  categoryName: string
): string {
  const rawSector = String(sectorName || '').trim();
  const canonSec = rawSector.includes('+')
    ? rawSector
        .split('+')
        .map((part) => canonicalizeSectorName(part.trim()) || part.trim())
        .filter(Boolean)
        .join('+')
    : canonicalizeSectorName(rawSector) || rawSector;

  return [
    normalizeArabic(canonSec).replace(/\s+/g, ''),
    normalizeArabic(jobTitle).replace(/\s+/g, ''),
    normalizeArabic(employeeName).replace(/\s+/g, ''),
    normalizeCategoryText(categoryName),
  ].join('::');
}

export interface ItemCategoryTargetInput {
  sectorName: string;
  jobTitle: string;
  employeeName: string;
  categoryName: string;
  targetValue: number;
}

export function saveItemCategoryTargets(
  db: Database.Database,
  targets: ItemCategoryTargetInput[]
): void {
  const now = new Date().toISOString();
  const tx = db.transaction(() => {
    const stmt = db.prepare(
      `INSERT INTO item_category_targets (
        target_key, sector_name, job_title, employee_name, category_name, target_value, updated_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(target_key) DO UPDATE SET
        sector_name = excluded.sector_name,
        job_title = excluded.job_title,
        employee_name = excluded.employee_name,
        category_name = excluded.category_name,
        target_value = excluded.target_value,
        updated_at = excluded.updated_at`
    );

    for (const item of targets) {
      const sectorName = String(item.sectorName || '').trim();
      const jobTitle = String(item.jobTitle || '').trim();
      const employeeName = String(item.employeeName || '').trim();
      const categoryName = String(item.categoryName || '').trim();
      if (!sectorName || !jobTitle || !employeeName || !categoryName) continue;

      // Rafat Mohamed is strictly auto-computed from Central Supply targets; never store manual overrides for him
      if (isDepartmentManagerJobTitle(jobTitle) && isRafatMohamedName(employeeName)) {
        continue;
      }

      const rawVal = item.targetValue === undefined || item.targetValue === null ? 0 : Number(item.targetValue);
      if (!Number.isFinite(rawVal) || rawVal < 0) {
        throw new Error(`قيمة الهدف غير صالحة لـ ${employeeName} (${categoryName})`);
      }

      const cleanVal = Number(rawVal.toFixed(2));
      const key = buildItemCategoryTargetKey(sectorName, jobTitle, employeeName, categoryName);
      stmt.run(key, sectorName, jobTitle, employeeName, categoryName, cleanVal, now);
    }
  });
  tx();
}

export function getItemCategoriesAnalytics(
  db: Database.Database,
  user: UserProfile,
  params: AnalyticsFilterParams & { jobTitle?: string }
): ItemCategoriesAnalyticsResponse {
  const activeBatch = getActiveBatch(db);
  if (!activeBatch) {
    return {
      rows: [],
      jobTitleOptions: [],
      hasCategoryTargetsSheet: false,
      effectiveSource: params.source,
    };
  }

  // If VF05N lacks Material item_code rows, use ZSALES for item-level analysis (never summing both)
  const hasItemsInRequested = db
    .prepare(
      `SELECT 1 AS ok FROM sales_records WHERE batch_id = ? AND source_type = ? AND COALESCE(item_code, '') != '' LIMIT 1`
    )
    .get(activeBatch.id, params.source) as { ok: number } | undefined;

  const effectiveSource: SalesSource =
    !hasItemsInRequested && params.source === 'VF05N' ? 'ZSALES' : params.source;

  // Query all authorized sectors in date/source scope so multi-sector `مدير ادارة` & auto-rollups work even under single-sector filter
  const scope = buildScopeFilter(user, { ...params, source: effectiveSource, sector: 'ALL' }, activeBatch.id);
  const authorizedSectors = user.role === 'admin' ? [...ORIGINAL_SECTORS] : scope.authorizedSectors;
  const authSectorSet = new Set(authorizedSectors);

  const selectedSector =
    params.sector && params.sector !== 'ALL' && params.sector !== 'undefined' && params.sector !== 'null'
      ? params.sector === 'UNLINKED'
        ? 'UNLINKED'
        : canonicalizeSectorName(params.sector)
      : 'ALL';

  const selectedJobTitle =
    params.jobTitle && params.jobTitle !== 'ALL' && params.jobTitle !== 'undefined' && params.jobTitle !== 'null'
      ? params.jobTitle.trim()
      : 'ALL';
  const selectedJobNorm = selectedJobTitle === 'ALL' ? '' : normalizeArabic(selectedJobTitle);

  // 1. Load Items and build sector-specific category mapping
  const itemRows = db
    .prepare(
      `SELECT
        item_code,
        item_name,
        COALESCE(category_name, '') AS category_name,
        COALESCE(chains_category, '') AS chains_category,
        COALESCE(army_category, '') AS army_category,
        COALESCE(top_retail_category, '') AS top_retail_category
       FROM items
       WHERE batch_id = ?`
    )
    .all(activeBatch.id) as Array<{
    item_code: string;
    item_name: string;
    category_name: string;
    chains_category: string;
    army_category: string;
    top_retail_category: string;
  }>;

  const itemByExact = new Map<string, (typeof itemRows)[number]>();
  const itemByTrimmed = new Map<string, (typeof itemRows)[number]>();
  for (const it of itemRows) {
    const exact = normalizeCode(it.item_code);
    if (exact) {
      if (!itemByExact.has(exact)) itemByExact.set(exact, it);
      const trimmed = stripLeadingZeros(exact);
      if (trimmed && !itemByTrimmed.has(trimmed)) itemByTrimmed.set(trimmed, it);
    }
  }

  // Strictly use sector-specific classification columns (`تصنيف السلاسل`, `تصنيف الجيش`, `التوب ريتيل`)
  // Never fall back to `التصنيف A` (`it.category_name`)
  const getItemCategoryForSector = (it: (typeof itemRows)[number], sectorName: string): string => {
    const canonSec = canonicalizeSectorName(sectorName);
    const normSec = normalizeArabic(sectorName);
    if (
      canonSec === 'التوريد المباشر' ||
      canonSec === 'التوريد المركزي' ||
      normSec.includes('سلاسل') ||
      normSec.includes('مباشر') ||
      normSec.includes('مركزي')
    ) {
      return it.chains_category.trim();
    }
    if (canonSec === 'الجيش' || normSec.includes('جيش')) {
      return it.army_category.trim();
    }
    if (
      canonSec === 'التوب ريتيل' ||
      (normSec.includes('توب') && normSec.includes('ريتيل')) ||
      (normSec.includes('تجزئ') && normSec.includes('مميز'))
    ) {
      return it.top_retail_category.trim();
    }
    return (it.chains_category || it.army_category || it.top_retail_category).trim();
  };

  // Precompute defined categories per sector & unified `تصنيف السلاسل` from `تعريف الاصناف`
  const sectorCategoriesMap = new Map<string, Map<string, string>>();
  const allChainsCategoriesMap = new Map<string, string>();
  for (const sec of ORIGINAL_SECTORS) {
    const catMap = new Map<string, string>();
    for (const it of itemRows) {
      const rawCat = getItemCategoryForSector(it, sec);
      if (rawCat) {
        const norm = normalizeCategoryText(rawCat);
        if (norm && !catMap.has(norm)) {
          catMap.set(norm, rawCat.replace(/\s+/g, ' ').trim());
        }
      }
      const rawChainCat = it.chains_category.trim();
      if (rawChainCat) {
        const normChain = normalizeCategoryText(rawChainCat);
        if (normChain && !allChainsCategoriesMap.has(normChain)) {
          allChainsCategoriesMap.set(normChain, rawChainCat.replace(/\s+/g, ' ').trim());
        }
      }
    }
    sectorCategoriesMap.set(sec, catMap);
  }

  // 2. Load Sales Reps, Chains/Supervisors, and Branches metadata (strictly sector-aware)
  const repRows = db
    .prepare(
      `SELECT rep_code, rep_name, COALESCE(sector_name, '') AS sector_name, COALESCE(supervisor_name, '') AS supervisor_name
       FROM sales_reps
       WHERE batch_id = ?`
    )
    .all(activeBatch.id) as Array<{
    rep_code: string;
    rep_name: string;
    sector_name: string;
    supervisor_name: string;
  }>;

  const repBySecExact = new Map<string, (typeof repRows)[number]>();
  const repBySecTrimmed = new Map<string, (typeof repRows)[number]>();
  const repByExact = new Map<string, (typeof repRows)[number]>();
  const repByTrimmed = new Map<string, (typeof repRows)[number]>();
  const repByNormName = new Map<string, (typeof repRows)[number]>();

  for (const r of repRows) {
    const sec = canonicalizeSectorName(r.sector_name);
    const exact = normalizeCode(r.rep_code);
    const trimmed = stripLeadingZeros(exact);
    if (exact) {
      if (sec) {
        if (!repBySecExact.has(`${sec}::${exact}`)) repBySecExact.set(`${sec}::${exact}`, r);
        if (trimmed && !repBySecTrimmed.has(`${sec}::${trimmed}`)) repBySecTrimmed.set(`${sec}::${trimmed}`, r);
      }
      if (!repByExact.has(exact)) repByExact.set(exact, r);
      if (trimmed && !repByTrimmed.has(trimmed)) repByTrimmed.set(trimmed, r);
    }
    const normName = normalizeArabic(r.rep_name);
    if (normName && !repByNormName.has(normName)) {
      repByNormName.set(normName, r);
    }
  }

  const resolveRep = (code: string, sectorHint?: string) => {
    const exact = normalizeCode(code);
    if (!exact) return undefined;
    const trimmed = stripLeadingZeros(exact);
    if (sectorHint) {
      const secMatch =
        repBySecExact.get(`${sectorHint}::${exact}`) ??
        (trimmed ? repBySecTrimmed.get(`${sectorHint}::${trimmed}`) : undefined);
      if (secMatch) return secMatch;
    }
    return repByExact.get(exact) ?? (trimmed ? repByTrimmed.get(trimmed) : undefined);
  };

  const chainRows = db
    .prepare(
      `SELECT chain_code, chain_name, department_name, sector_name, COALESCE(supervisor_name, '') AS supervisor_name
       FROM departments_chains
       WHERE batch_id = ?`
    )
    .all(activeBatch.id) as Array<{
    chain_code: string;
    chain_name: string;
    department_name: string;
    sector_name: string;
    supervisor_name: string;
  }>;

  const chainByExact = new Map<string, (typeof chainRows)[number]>();
  const chainByTrimmed = new Map<string, (typeof chainRows)[number]>();
  const chainByNormName = new Map<string, (typeof chainRows)[number]>();
  const chainSupervisorsBySector = new Map<string, Map<string, string>>(); // sec -> Map<normSup, displaySup>
  const departmentsBySector = new Map<string, Set<string>>();
  const chainsBySector = new Map<string, Set<string>>();

  for (const c of chainRows) {
    const exact = normalizeCode(c.chain_code);
    if (exact) {
      if (!chainByExact.has(exact)) chainByExact.set(exact, c);
      const trimmed = stripLeadingZeros(exact);
      if (trimmed && !chainByTrimmed.has(trimmed)) chainByTrimmed.set(trimmed, c);
    }
    const normChain = normalizeArabic(c.chain_name);
    if (normChain && !chainByNormName.has(normChain)) chainByNormName.set(normChain, c);

    const sec = canonicalizeSectorName(c.sector_name);
    const supClean = c.supervisor_name.trim();
    const supNorm = normalizeArabic(supClean);
    if (sec && supNorm && supClean !== 'غير محدد') {
      const sMap = chainSupervisorsBySector.get(sec) ?? new Map<string, string>();
      if (!sMap.has(supNorm)) sMap.set(supNorm, supClean);
      chainSupervisorsBySector.set(sec, sMap);
    }
    const deptNorm = normalizeArabic(c.department_name);
    if (sec && deptNorm) {
      const dSet = departmentsBySector.get(sec) ?? new Set<string>();
      dSet.add(deptNorm);
      departmentsBySector.set(sec, dSet);
    }
    if (sec && normChain) {
      const cSet = chainsBySector.get(sec) ?? new Set<string>();
      cSet.add(normChain);
      chainsBySector.set(sec, cSet);
    }
  }

  const repSupRepsMap = new Map<string, Set<string>>();
  const repSupDisplayMap = new Map<string, Map<string, string>>();
  for (const r of repRows) {
    const sec = canonicalizeSectorName(r.sector_name);
    const supClean = r.supervisor_name.trim();
    const supNorm = normalizeArabic(supClean);
    const repKey = normalizeCode(r.rep_code);
    if (sec && supNorm && supClean !== 'غير محدد') {
      const dispMap = repSupDisplayMap.get(sec) ?? new Map<string, string>();
      if (!dispMap.has(supNorm)) dispMap.set(supNorm, supClean);
      repSupDisplayMap.set(sec, dispMap);
      if (repKey) {
        const k = `${sec}::${supNorm}`;
        const set = repSupRepsMap.get(k) ?? new Set<string>();
        set.add(repKey);
        repSupRepsMap.set(k, set);
      }
    }
  }

  const branchRows = db
    .prepare(
      `SELECT sector_name, COALESCE(rep_code, '') AS rep_code, COALESCE(supervisor_name, '') AS supervisor_name
       FROM branches
       WHERE batch_id = ?`
    )
    .all(activeBatch.id) as Array<{ sector_name: string; rep_code: string; supervisor_name: string }>;
  const branchSupRepsMap = new Map<string, Set<string>>();
  const branchSupDisplayMap = new Map<string, Map<string, string>>();
  for (const b of branchRows) {
    const sec = canonicalizeSectorName(b.sector_name);
    const supClean = b.supervisor_name.trim();
    const supNorm = normalizeArabic(supClean);
    const repDef = resolveRep(b.rep_code, sec);
    const repKey = normalizeCode(repDef?.rep_code || b.rep_code);
    if (sec && supNorm && supClean !== 'غير محدد') {
      const dispMap = branchSupDisplayMap.get(sec) ?? new Map<string, string>();
      if (!dispMap.has(supNorm)) dispMap.set(supNorm, supClean);
      branchSupDisplayMap.set(sec, dispMap);
      if (repKey) {
        const k = `${sec}::${supNorm}`;
        const set = branchSupRepsMap.get(k) ?? new Set<string>();
        set.add(repKey);
        branchSupRepsMap.set(k, set);
      }
    }
  }

  // 3. Load Saved In-App Targets (`item_category_targets`) and Roster-Only Role Metadata
  // NOTE: Excel `category_targets.target_value` is intentionally IGNORED! Targets come strictly from `item_category_targets`.
  const savedTargetRows = db
    .prepare(`SELECT target_key, target_value FROM item_category_targets`)
    .all() as Array<{ target_key: string; target_value: number }>;
  const savedTargetMap = new Map<string, number>();
  for (const sr of savedTargetRows) {
    savedTargetMap.set(sr.target_key, Number(sr.target_value || 0));
  }

  const rawCatRoster = db
    .prepare(
      `SELECT DISTINCT sector_name, job_title, employee_code, employee_name, category_name
       FROM category_targets
       WHERE batch_id = ?`
    )
    .all(activeBatch.id) as Array<{
    sector_name: string;
    job_title: string;
    employee_code: string;
    employee_name: string;
    category_name: string;
  }>;

  const supervisorSectorLookup = new Map<string, string>();
  for (const [sec, sups] of chainSupervisorsBySector.entries()) {
    for (const s of sups.keys()) {
      if (!supervisorSectorLookup.has(s)) supervisorSectorLookup.set(s, sec);
    }
  }

  const resolvedRoster = rawCatRoster.map((ct) => {
    const jobTitle = ct.job_title.trim();
    const employeeName = ct.employee_name.trim() || normalizeCode(ct.employee_code);
    const isDeptMgr = isDepartmentManagerJobTitle(jobTitle);
    const managedSectors = isDeptMgr ? resolveDepartmentManagerSectors(employeeName, ct.sector_name) : [];

    let sec = ct.sector_name ? canonicalizeSectorName(ct.sector_name) : '';
    const normCode = normalizeCode(ct.employee_code);
    const normName = normalizeArabic(ct.employee_name);
    if (!sec && !isDeptMgr) {
      const repMatch =
        (normCode ? resolveRep(normCode) : undefined) ?? (normName ? repByNormName.get(normName) : undefined);
      if (repMatch && repMatch.sector_name) {
        sec = canonicalizeSectorName(repMatch.sector_name);
      } else if (normName && supervisorSectorLookup.has(normName)) {
        sec = supervisorSectorLookup.get(normName)!;
      }
    }
    return {
      sectorName: sec,
      rawSectorName: ct.sector_name,
      managedSectors,
      isDepartmentManager: isDeptMgr,
      employeeCode: normCode,
      employeeName,
      jobTitle,
      categoryName: ct.category_name.replace(/\s+/g, ' ').trim(),
    };
  });

  // 4. Query sales aggregated strictly by (sector, payer, rep, item) within active scope
  const salesRows = db
    .prepare(
      `SELECT
        COALESCE(sr.resolved_sector, 'غير مرتبط بقطاع') AS sector_name,
        COALESCE(sr.payer_code, '') AS payer_code,
        COALESCE(NULLIF(sr.resolved_chain_name, ''), NULLIF(sr.payer_name, ''), '') AS chain_name,
        COALESCE(NULLIF(sr.resolved_rep_code, ''), sr.rep_code, '') AS rep_code,
        COALESCE(NULLIF(sr.resolved_rep_name, ''), '') AS rep_name,
        COALESCE(sr.item_code, '') AS item_code,
        COALESCE(SUM(sr.sales_value), 0) AS sales,
        COALESCE(SUM(sr.tax_value), 0) AS tax
       FROM sales_records sr
       WHERE ${scope.sql}
       GROUP BY sr.resolved_sector, sr.payer_code, COALESCE(NULLIF(sr.resolved_chain_name, ''), NULLIF(sr.payer_name, ''), ''), COALESCE(NULLIF(sr.resolved_rep_code, ''), sr.rep_code, ''), COALESCE(NULLIF(sr.resolved_rep_name, ''), ''), sr.item_code`
    )
    .all(...scope.args) as Array<{
    sector_name: string;
    payer_code: string;
    chain_name: string;
    rep_code: string;
    rep_name: string;
    item_code: string;
    sales: number;
    tax: number;
  }>;

  interface SalesTaxBucket {
    sales: number;
    tax: number;
  }
  const addBucket = (map: Map<string, SalesTaxBucket>, key: string, sales: number, tax: number) => {
    const b = map.get(key) ?? { sales: 0, tax: 0 };
    b.sales += sales;
    b.tax += tax;
    map.set(key, b);
  };

  const sectorCatAgg = new Map<string, SalesTaxBucket>();
  const sectorChainsCatAgg = new Map<string, SalesTaxBucket>();
  const repCatAgg = new Map<string, SalesTaxBucket>();
  const supCatAgg = new Map<string, SalesTaxBucket>();
  const deptCatAgg = new Map<string, SalesTaxBucket>();
  const chainCatAgg = new Map<string, SalesTaxBucket>();

  const activeRepsBySector = new Map<string, Map<string, { repCode: string; repName: string }>>();
  const activeRepUnclassifiedCatsBySector = new Map<string, Map<string, string>>();

  for (const r of salesRows) {
    const sec = r.sector_name === 'غير مرتبط بقطاع' ? r.sector_name : canonicalizeSectorName(r.sector_name);
    const normItem = normalizeCode(r.item_code);
    const itemDef = normItem ? itemByExact.get(normItem) ?? itemByTrimmed.get(stripLeadingZeros(normItem)) : undefined;
    const rawCat = itemDef ? getItemCategoryForSector(itemDef, sec) : '';
    const cleanCat = rawCat ? rawCat.replace(/\s+/g, ' ').trim() : 'غير مصنف';
    const normCat = normalizeCategoryText(cleanCat) || UNCLASSIFIED_CAT_NORM;

    addBucket(sectorCatAgg, `${sec}::${normCat}`, r.sales, r.tax);

    // Unified `تصنيف السلاسل` aggregation per sector (used by `مدير ادارة`)
    const rawChainsCat = itemDef ? itemDef.chains_category.trim() : '';
    const cleanChainsCat = rawChainsCat ? rawChainsCat.replace(/\s+/g, ' ').trim() : 'غير مصنف';
    const normChainsCat = normalizeCategoryText(cleanChainsCat) || UNCLASSIFIED_CAT_NORM;
    addBucket(sectorChainsCatAgg, `${sec}::${normChainsCat}`, r.sales, r.tax);

    // Resolve rep strictly scoped to `sec`
    const rawRep = normalizeCode(r.rep_code);
    const repDef = rawRep ? resolveRep(rawRep, sec) : undefined;
    const repKey = repDef ? normalizeCode(repDef.rep_code) : rawRep;
    if (repKey) {
      addBucket(repCatAgg, `${sec}::${repKey}::${normCat}`, r.sales, r.tax);
      const secReps = activeRepsBySector.get(sec) ?? new Map<string, { repCode: string; repName: string }>();
      if (!secReps.has(repKey)) {
        secReps.set(repKey, {
          repCode: repKey,
          repName: (repDef?.rep_name || r.rep_name || repKey).trim(),
        });
      }
      activeRepsBySector.set(sec, secReps);
    }

    if (normCat === UNCLASSIFIED_CAT_NORM) {
      const unMap = activeRepUnclassifiedCatsBySector.get(sec) ?? new Map<string, string>();
      unMap.set(normCat, cleanCat);
      activeRepUnclassifiedCatsBySector.set(sec, unMap);
    }

    // Resolve chain -> supervisor & department strictly inside `sec`
    const normPayer = normalizeCode(r.payer_code);
    const chainDef =
      (normPayer ? chainByExact.get(normPayer) ?? chainByTrimmed.get(stripLeadingZeros(normPayer)) : undefined) ??
      (r.chain_name ? chainByNormName.get(normalizeArabic(r.chain_name)) : undefined);
    const supNorm = normalizeArabic(chainDef?.supervisor_name || repDef?.supervisor_name || '');
    if (supNorm) {
      addBucket(supCatAgg, `${sec}::${supNorm}::${normCat}`, r.sales, r.tax);
    }
    const deptNorm = normalizeArabic(chainDef?.department_name || '');
    if (deptNorm) {
      addBucket(deptCatAgg, `${sec}::${deptNorm}::${normCat}`, r.sales, r.tax);
    }
    const chainNorm = normalizeArabic(chainDef?.chain_name || r.chain_name || '');
    if (chainNorm) {
      addBucket(chainCatAgg, `${sec}::${chainNorm}::${normCat}`, r.sales, r.tax);
    }
  }

  // 5. Sector visibility helpers
  const isSectorVisible = (sec: string): boolean => {
    if (!sec) return selectedSector === 'ALL';
    if (user.role !== 'admin' && !authSectorSet.has(sec)) return false;
    if (selectedSector !== 'ALL' && sec !== selectedSector) return false;
    return true;
  };
  const isDeptManagerVisible = (managedSectors: string[]): boolean => {
    if (selectedSector === 'UNLINKED') return false;
    const authManaged =
      user.role === 'admin' ? managedSectors : managedSectors.filter((s) => authSectorSet.has(s));
    if (authManaged.length === 0) return false;
    if (selectedSector === 'ALL') return true;
    return authManaged.includes(selectedSector);
  };

  // 6. Build full authorized matrix across all 4 tiers BEFORE applying jobTitle/sector display filters
  //    so `رافت محمد` always resolves Central Supply's per-category auto-target cleanly without double-counting!
  interface InternalMatrixRow {
    targetKey: string;
    sectorName: string;
    managedSectors: string[];
    isDepartmentManager: boolean;
    roleTier: 'dept_mgr' | 'sec_mgr' | 'sup' | 'rep';
    jobTitle: string;
    employeeCode: string;
    employeeName: string;
    categoryName: string;
    normCat: string;
    sales: number;
    tax: number;
    target: number;
    isAutoTarget: boolean;
  }

  const allMatrixRows: InternalMatrixRow[] = [];

  // 6A. Sector-scoped Rep rows (every rep in authorized sector x every category of that sector)
  for (const sec of authorizedSectors) {
    const secCatsMap = sectorCategoriesMap.get(sec) ?? new Map<string, string>();
    const secRepRoster = resolvedRoster.filter(
      (ct) => isSalesRepJobTitle(ct.jobTitle) && (!ct.sectorName || ct.sectorName === sec)
    );
    const repJobLabel = secRepRoster.find((ct) => ct.sectorName === sec)?.jobTitle || secRepRoster[0]?.jobTitle || 'مندوب مبيعات';

    const secRepsMap = new Map<string, { repKey: string; repCode: string; repName: string }>();
    for (const r of repRows) {
      if (canonicalizeSectorName(r.sector_name) !== sec) continue;
      const rk = normalizeCode(r.rep_code);
      if (!rk || secRepsMap.has(rk)) continue;
      secRepsMap.set(rk, {
        repKey: rk,
        repCode: rk,
        repName: r.rep_name.trim() || rk,
      });
    }
    for (const ct of secRepRoster) {
      if (ct.sectorName && ct.sectorName !== sec) continue;
      const resolved =
        (ct.employeeCode ? resolveRep(ct.employeeCode, sec) : undefined) ??
        repByNormName.get(normalizeArabic(ct.employeeName));
      if (resolved && resolved.sector_name && canonicalizeSectorName(resolved.sector_name) !== sec) continue;
      const rk = resolved
        ? normalizeCode(resolved.rep_code)
        : ct.employeeCode || normalizeArabic(ct.employeeName);
      if (!rk || secRepsMap.has(rk)) continue;
      secRepsMap.set(rk, {
        repKey: rk,
        repCode: resolved?.rep_code || ct.employeeCode,
        repName: resolved?.rep_name || ct.employeeName,
      });
    }
    const activeInSec = activeRepsBySector.get(sec);
    if (activeInSec) {
      for (const [rk, info] of activeInSec.entries()) {
        if (!secRepsMap.has(rk)) {
          secRepsMap.set(rk, {
            repKey: rk,
            repCode: info.repCode,
            repName: info.repName,
          });
        }
      }
    }

    const hasUnclassified = (activeRepUnclassifiedCatsBySector.get(sec)?.size ?? 0) > 0;

    for (const rep of secRepsMap.values()) {
      for (const [normCat, displayCat] of secCatsMap.entries()) {
        const st = repCatAgg.get(`${sec}::${rep.repKey}::${normCat}`);
        const sales = Number((st?.sales ?? 0).toFixed(2));
        const tax = Number((st?.tax ?? 0).toFixed(2));
        const targetKey = buildItemCategoryTargetKey(sec, repJobLabel, rep.repName, displayCat);
        const fallbackKey = buildItemCategoryTargetKey(sec, 'مندوب مبيعات', rep.repName, displayCat);
        const target = Number((savedTargetMap.get(targetKey) ?? savedTargetMap.get(fallbackKey) ?? 0).toFixed(2));

        allMatrixRows.push({
          targetKey,
          sectorName: sec,
          managedSectors: [sec],
          isDepartmentManager: false,
          roleTier: 'rep',
          jobTitle: repJobLabel,
          employeeCode: rep.repCode,
          employeeName: rep.repName,
          categoryName: displayCat,
          normCat,
          sales,
          tax,
          target,
          isAutoTarget: false,
        });
      }

      if (hasUnclassified && !secCatsMap.has(UNCLASSIFIED_CAT_NORM)) {
        const unBucket = repCatAgg.get(`${sec}::${rep.repKey}::${UNCLASSIFIED_CAT_NORM}`);
        if (unBucket && (unBucket.sales !== 0 || unBucket.tax !== 0)) {
          const sales = Number(unBucket.sales.toFixed(2));
          const tax = Number(unBucket.tax.toFixed(2));
          const targetKey = buildItemCategoryTargetKey(sec, repJobLabel, rep.repName, 'غير مصنف');
          const fallbackKey = buildItemCategoryTargetKey(sec, 'مندوب مبيعات', rep.repName, 'غير مصنف');
          const target = Number((savedTargetMap.get(targetKey) ?? savedTargetMap.get(fallbackKey) ?? 0).toFixed(2));
          allMatrixRows.push({
            targetKey,
            sectorName: sec,
            managedSectors: [sec],
            isDepartmentManager: false,
            roleTier: 'rep',
            jobTitle: repJobLabel,
            employeeCode: rep.repCode,
            employeeName: rep.repName,
            categoryName: 'غير مصنف',
            normCat: UNCLASSIFIED_CAT_NORM,
            sales,
            tax,
            target,
            isAutoTarget: false,
          });
        }
      }
    }
  }

  // 6B. Sector-scoped Supervisors (`مشرف قطاع` / `مشرف مبيعات`) & Sector Managers (`مدير قطاع` / `قائد قطاع الجيش`)
  //     Strictly use ONLY the categories assigned to each manager/supervisor in `category_targets` (when present)
  //     so managers in the same sector/job title do NOT inherit each other's categories or duplicate sector totals.
  interface SectorStaffMember {
    sectorName: string;
    roleTier: 'sec_mgr' | 'sup';
    jobTitle: string;
    employeeCode: string;
    employeeName: string;
    assignedCategories: Map<string, string>;
    hasExplicitCategories: boolean;
  }
  const sectorStaffMap = new Map<string, SectorStaffMember>();

  const registerSectorStaff = (
    sec: string,
    roleTier: 'sec_mgr' | 'sup',
    jobTitle: string,
    employeeCode: string,
    employeeName: string,
    categoryName?: string
  ) => {
    if (!sec || !authSectorSet.has(sec)) return;
    const cleanName = employeeName.trim() || employeeCode;
    const normEmp = normalizeArabic(cleanName);
    if (!normEmp || cleanName === 'غير محدد') return;
    const dedupKey =
      roleTier === 'sup'
        ? `${sec}::SUP::${normEmp}`
        : `${sec}::${normalizeArabic(jobTitle)}::${normEmp}`;
    let staff = sectorStaffMap.get(dedupKey);
    if (!staff) {
      staff = {
        sectorName: sec,
        roleTier,
        jobTitle: jobTitle.trim(),
        employeeCode,
        employeeName: cleanName,
        assignedCategories: new Map<string, string>(),
        hasExplicitCategories: false,
      };
      sectorStaffMap.set(dedupKey, staff);
    }
    if (categoryName) {
      const normCat = normalizeCategoryText(categoryName);
      if (normCat) {
        staff.hasExplicitCategories = true;
        if (!staff.assignedCategories.has(normCat)) {
          const canonicalDisplay = sectorCategoriesMap.get(sec)?.get(normCat) ?? categoryName;
          staff.assignedCategories.set(normCat, canonicalDisplay);
        }
      }
    }
  };

  // Register non-rep, non-dept-manager roster entries from `category_targets` WITH their assigned categories
  for (const ct of resolvedRoster) {
    if (ct.isDepartmentManager || isSalesRepJobTitle(ct.jobTitle)) continue;
    if (!ct.sectorName || !authSectorSet.has(ct.sectorName)) continue;
    const tier: 'sec_mgr' | 'sup' = isSupervisorJobTitle(ct.jobTitle) ? 'sup' : 'sec_mgr';
    registerSectorStaff(ct.sectorName, tier, ct.jobTitle, ct.employeeCode, ct.employeeName, ct.categoryName);
  }

  // Supplement supervisors from `departments_chains`, and if a sector still lacks supervisors, fall back to `branches`/`sales_reps`
  for (const sec of authorizedSectors) {
    const chainSups = chainSupervisorsBySector.get(sec);
    if (chainSups) {
      for (const displaySup of chainSups.values()) {
        registerSectorStaff(sec, 'sup', 'مشرف قطاع', '', displaySup);
      }
    }
    const hasSupForSec = Array.from(sectorStaffMap.values()).some(
      (s) => s.sectorName === sec && s.roleTier === 'sup'
    );
    if (!hasSupForSec) {
      const brSups = branchSupDisplayMap.get(sec);
      if (brSups) {
        for (const displaySup of brSups.values()) {
          registerSectorStaff(sec, 'sup', 'مشرف قطاع', '', displaySup);
        }
      }
      const rSups = repSupDisplayMap.get(sec);
      if (rSups) {
        for (const displaySup of rSups.values()) {
          registerSectorStaff(sec, 'sup', 'مشرف قطاع', '', displaySup);
        }
      }
    }
  }

  const staffRenderedCats = new Map<SectorStaffMember, Map<string, string>>();
  const tierCategoryAssigneeCount = new Map<string, number>();

  for (const staff of sectorStaffMap.values()) {
    const catsToRender = staff.hasExplicitCategories
      ? staff.assignedCategories
      : sectorCategoriesMap.get(staff.sectorName) ?? new Map<string, string>();
    staffRenderedCats.set(staff, catsToRender);

    const tierGroupKey = staff.roleTier === 'sup' ? 'SUP' : normalizeArabic(staff.jobTitle);
    for (const normCat of catsToRender.keys()) {
      const k = `${staff.sectorName}::${tierGroupKey}::${normCat}`;
      tierCategoryAssigneeCount.set(k, (tierCategoryAssigneeCount.get(k) ?? 0) + 1);
    }
  }

  const lookupSectorStaffSalesTax = (staff: SectorStaffMember, normCat: string): SalesTaxBucket => {
    const sec = staff.sectorName;
    const empNorm = normalizeArabic(staff.employeeName);
    const tierGroupKey = staff.roleTier === 'sup' ? 'SUP' : normalizeArabic(staff.jobTitle);
    const assigneeCount = tierCategoryAssigneeCount.get(`${sec}::${tierGroupKey}::${normCat}`) ?? 0;

    // If ONLY ONE employee in this `(sector, jobTitle/tier)` owns `normCat`, they account for the full
    // sector sales & tax of that category so sector totals stay identical across Rep / Supervisor / Manager.
    if (assigneeCount <= 1) {
      return sectorCatAgg.get(`${sec}::${normCat}`) ?? { sales: 0, tax: 0 };
    }

    if (staff.roleTier === 'sup') {
      const isChainSup = chainSupervisorsBySector.get(sec)?.has(empNorm) ?? false;
      const branchReps = branchSupRepsMap.get(`${sec}::${empNorm}`);
      const repSupReps = repSupRepsMap.get(`${sec}::${empNorm}`);
      if (isChainSup) {
        return supCatAgg.get(`${sec}::${empNorm}::${normCat}`) ?? { sales: 0, tax: 0 };
      }
      let totalSales = 0;
      let totalTax = 0;
      if (branchReps) {
        for (const rk of branchReps) {
          const b = repCatAgg.get(`${sec}::${rk}::${normCat}`);
          if (b) {
            totalSales += b.sales;
            totalTax += b.tax;
          }
        }
        return { sales: totalSales, tax: totalTax };
      }
      if (repSupReps) {
        for (const rk of repSupReps) {
          const b = repCatAgg.get(`${sec}::${rk}::${normCat}`);
          if (b) {
            totalSales += b.sales;
            totalTax += b.tax;
          }
        }
        return { sales: totalSales, tax: totalTax };
      }
      return supCatAgg.get(`${sec}::${empNorm}::${normCat}`) ?? { sales: 0, tax: 0 };
    }

    if (departmentsBySector.get(sec)?.has(empNorm)) {
      return deptCatAgg.get(`${sec}::${empNorm}::${normCat}`) ?? { sales: 0, tax: 0 };
    }
    if (chainsBySector.get(sec)?.has(empNorm)) {
      return chainCatAgg.get(`${sec}::${empNorm}::${normCat}`) ?? { sales: 0, tax: 0 };
    }
    return sectorCatAgg.get(`${sec}::${normCat}`) ?? { sales: 0, tax: 0 };
  };

  for (const staff of sectorStaffMap.values()) {
    const catsToRender = staffRenderedCats.get(staff) ?? new Map<string, string>();
    for (const [normCat, displayCat] of catsToRender.entries()) {
      const st = lookupSectorStaffSalesTax(staff, normCat);
      const sales = Number(st.sales.toFixed(2));
      const tax = Number(st.tax.toFixed(2));
      const targetKey = buildItemCategoryTargetKey(staff.sectorName, staff.jobTitle, staff.employeeName, displayCat);
      const target = Number((savedTargetMap.get(targetKey) ?? 0).toFixed(2));

      allMatrixRows.push({
        targetKey,
        sectorName: staff.sectorName,
        managedSectors: [staff.sectorName],
        isDepartmentManager: false,
        roleTier: staff.roleTier,
        jobTitle: staff.jobTitle,
        employeeCode: staff.employeeCode,
        employeeName: staff.employeeName,
        categoryName: displayCat,
        normCat,
        sales,
        tax,
        target,
        isAutoTarget: false,
      });
    }
  }

  // 6C. Compute Central Supply (`التوريد المركزي`) per-category waterfall auto-target for `رافت محمد`
  // Priority per category: `مدير قطاع` > `مشرف قطاع` > `مندوب مبيعات` (prevents multi-level double counting!)
  const centralAutoTargetByCat = new Map<string, number>();
  const centralCatSums = new Map<string, { secMgr: number; sup: number; rep: number }>();
  for (const row of allMatrixRows) {
    if (row.sectorName !== 'التوريد المركزي') continue;
    const entry = centralCatSums.get(row.normCat) ?? { secMgr: 0, sup: 0, rep: 0 };
    if (row.roleTier === 'sec_mgr') entry.secMgr += row.target;
    else if (row.roleTier === 'sup') entry.sup += row.target;
    else if (row.roleTier === 'rep') entry.rep += row.target;
    centralCatSums.set(row.normCat, entry);
  }
  for (const [normCat, sums] of centralCatSums.entries()) {
    const resolved = sums.secMgr > 0 ? sums.secMgr : sums.sup > 0 ? sums.sup : sums.rep;
    centralAutoTargetByCat.set(normCat, Number(resolved.toFixed(2)));
  }

  // 6D. Department Managers (`مدير إدارة` / `مدير ادارة`)
  interface DeptMgrDef {
    jobTitle: string;
    employeeCode: string;
    employeeName: string;
    managedSectors: string[];
    sectorLabel: string;
    isRafat: boolean;
    isAbdelSattar: boolean;
    assignedCategories: Map<string, string>;
    hasExplicitCategories: boolean;
  }
  const deptMgrMap = new Map<string, DeptMgrDef>();
  const registerDeptMgr = (
    rawName: string,
    rawCode: string,
    rawJobTitle: string,
    rawSectorName: string,
    categoryName?: string
  ) => {
    const cleanName = rawName.trim() || rawCode;
    const noSpace = normalizeArabic(cleanName).replace(/\s+/g, '');
    if (!noSpace) return;
    let dm = deptMgrMap.get(noSpace);
    if (!dm) {
      const managedSectors = resolveDepartmentManagerSectors(cleanName, rawSectorName);
      const isRafat = isRafatMohamedName(cleanName);
      const isAbdelSattar = isAbdelSattarName(cleanName);
      const sectorLabel = isRafat
        ? 'التوريد المركزي'
        : isAbdelSattar
          ? ABDEL_SATTAR_SECTOR_LABEL
          : managedSectors.join(' + ');

      dm = {
        jobTitle: rawJobTitle.trim() || 'مدير إدارة',
        employeeCode: rawCode,
        employeeName: cleanName,
        managedSectors,
        sectorLabel,
        isRafat,
        isAbdelSattar,
        assignedCategories: new Map<string, string>(),
        hasExplicitCategories: false,
      };
      deptMgrMap.set(noSpace, dm);
    }
    if (categoryName) {
      const normCat = normalizeCategoryText(categoryName);
      if (normCat) {
        dm.hasExplicitCategories = true;
        if (!dm.assignedCategories.has(normCat)) {
          const canonicalDisplay =
            (dm.isRafat
              ? sectorCategoriesMap.get('التوريد المركزي')?.get(normCat)
              : undefined) ??
            allChainsCategoriesMap.get(normCat) ??
            categoryName;
          dm.assignedCategories.set(normCat, canonicalDisplay);
        }
      }
    }
  };

  for (const ct of resolvedRoster) {
    if (ct.isDepartmentManager) {
      registerDeptMgr(ct.employeeName, ct.employeeCode, ct.jobTitle, ct.rawSectorName, ct.categoryName);
    }
  }

  // Fallback: ensure `رافت محمد` & `عبدالستار مصطفى` exist whenever their managed sectors exist in active batch
  const batchSectorsRows = db
    .prepare(`SELECT sector_name FROM sectors WHERE batch_id = ?`)
    .all(activeBatch.id) as Array<{ sector_name: string }>;
  const activeSectorsPresent = new Set<string>([
    ...batchSectorsRows.map((s) => canonicalizeSectorName(s.sector_name)),
    ...repRows.map((r) => canonicalizeSectorName(r.sector_name)),
    ...Array.from(activeRepsBySector.keys()),
  ]);

  if (
    !Array.from(deptMgrMap.values()).some((d) => d.isRafat) &&
    activeSectorsPresent.has('التوريد المركزي') &&
    authSectorSet.has('التوريد المركزي')
  ) {
    registerDeptMgr('رافت محمد', '', 'مدير إدارة', 'التوريد المركزي');
  }
  if (
    !Array.from(deptMgrMap.values()).some((d) => d.isAbdelSattar) &&
    ['التوريد المباشر', 'الجيش', 'التوب ريتيل'].some(
      (s) => activeSectorsPresent.has(s) && authSectorSet.has(s)
    )
  ) {
    registerDeptMgr('عبدالستار مصطفى', '', 'مدير إدارة', ABDEL_SATTAR_SECTOR_LABEL);
  }

  for (const dm of deptMgrMap.values()) {
    const authManaged =
      user.role === 'admin'
        ? dm.managedSectors
        : dm.managedSectors.filter((s) => authSectorSet.has(s));
    if (authManaged.length === 0) continue;

    const fallbackCats = dm.isRafat
      ? sectorCategoriesMap.get('التوريد المركزي') ?? allChainsCategoriesMap
      : allChainsCategoriesMap;
    const catsMap = dm.hasExplicitCategories ? dm.assignedCategories : fallbackCats;

    for (const [normCat, displayCat] of catsMap.entries()) {
      let totalSales = 0;
      let totalTax = 0;
      for (const sec of authManaged) {
        const b = sectorChainsCatAgg.get(`${sec}::${normCat}`);
        if (b) {
          totalSales += b.sales;
          totalTax += b.tax;
        }
      }
      const sales = Number(totalSales.toFixed(2));
      const tax = Number(totalTax.toFixed(2));
      const targetKey = buildItemCategoryTargetKey(dm.sectorLabel, dm.jobTitle, dm.employeeName, displayCat);

      // `رافت محمد` auto-computes from Central Supply; `عبدالستار مصطفى` (and any other dept mgr) is strictly manual!
      const target = dm.isRafat
        ? centralAutoTargetByCat.get(normCat) ?? 0
        : Number((savedTargetMap.get(targetKey) ?? 0).toFixed(2));

      allMatrixRows.push({
        targetKey,
        sectorName: dm.sectorLabel,
        managedSectors: dm.managedSectors,
        isDepartmentManager: true,
        roleTier: 'dept_mgr',
        jobTitle: dm.jobTitle,
        employeeCode: dm.employeeCode,
        employeeName: dm.employeeName,
        categoryName: displayCat,
        normCat,
        sales,
        tax,
        target,
        isAutoTarget: dm.isRafat,
      });
    }
  }

  // 7. Filter by selectedSector & collect cascading jobTitleOptions, then filter by selectedJobTitle
  const sectorScopedRows = allMatrixRows.filter((row) =>
    row.isDepartmentManager
      ? isDeptManagerVisible(row.managedSectors)
      : isSectorVisible(row.sectorName)
  );

  const jobTitleMap = new Map<string, string>();
  for (const row of sectorScopedRows) {
    const normJob = normalizeArabic(row.jobTitle);
    if (normJob && !jobTitleMap.has(normJob)) {
      jobTitleMap.set(normJob, row.jobTitle);
    }
  }
  const jobTitleOptions = Array.from(jobTitleMap.values());

  const outputRows: ItemCategoryAnalysisRow[] = [];
  for (const row of sectorScopedRows) {
    if (selectedJobNorm && normalizeArabic(row.jobTitle) !== selectedJobNorm) continue;
    outputRows.push({
      targetKey: row.targetKey,
      sectorName: row.sectorName,
      jobTitle: row.jobTitle,
      employeeCode: row.employeeCode,
      employeeName: row.employeeName,
      categoryName: row.categoryName,
      sales: row.sales,
      tax: row.tax,
      target: row.target,
      achievementPct: calculateAchievementPct(row.sales, row.target, row.tax),
      isAutoTarget: row.isAutoTarget,
    });
  }

  outputRows.sort((a, b) => {
    const secCmp = a.sectorName.localeCompare(b.sectorName, 'ar');
    if (secCmp !== 0) return secCmp;
    const jobCmp = a.jobTitle.localeCompare(b.jobTitle, 'ar');
    if (jobCmp !== 0) return jobCmp;
    const empCmp = a.employeeName.localeCompare(b.employeeName, 'ar');
    if (empCmp !== 0) return empCmp;
    const salesDiff = b.sales - a.sales;
    if (salesDiff !== 0) return salesDiff;
    return a.categoryName.localeCompare(b.categoryName, 'ar');
  });

  return {
    rows: outputRows,
    jobTitleOptions,
    hasCategoryTargetsSheet: true,
    effectiveSource,
  };
}

export function getTimelineSeries(
  db: Database.Database,
  user: UserProfile,
  params: AnalyticsFilterParams
): TimelinePoint[] {
  const activeBatch = getActiveBatch(db);
  if (!activeBatch) return [];
  const scope = buildScopeFilter(user, params, activeBatch.id);

  const rows = db
    .prepare(
      `SELECT
        CASE WHEN sr.billing_date = '' OR sr.billing_date IS NULL THEN 'غير مؤرخ' ELSE sr.billing_date END AS b_date,
        COALESCE(SUM(sr.sales_value), 0) AS sales,
        COALESCE(SUM(sr.tax_value), 0) AS tax,
        COALESCE(SUM(CASE WHEN sr.sales_value >= 0 THEN sr.sales_value ELSE 0 END), 0) AS gross_sales,
        COALESCE(SUM(CASE WHEN sr.sales_value < 0 THEN sr.sales_value ELSE 0 END), 0) AS returns_value,
        COUNT(*) AS record_count
       FROM sales_records sr
       WHERE ${scope.sql}
       GROUP BY b_date
       ORDER BY b_date ASC`
    )
    .all(...scope.args) as Array<{
    b_date: string;
    sales: number;
    tax: number;
    gross_sales: number;
    returns_value: number;
    record_count: number;
  }>;

  let cumulative = 0;
  return rows.map((r) => {
    cumulative += r.sales;
    return {
      date: r.b_date,
      sales: Number(r.sales.toFixed(2)),
      tax: Number(r.tax.toFixed(2)),
      grossSales: Number(r.gross_sales.toFixed(2)),
      returnsValue: Number(r.returns_value.toFixed(2)),
      recordCount: r.record_count,
      cumulativeSales: Number(cumulative.toFixed(2)),
    };
  });
}

export interface TransactionsQueryParams extends AnalyticsFilterParams {
  search?: string;
  linkStatus?: 'all' | 'linked' | 'unlinked_branch' | 'unlinked_sector' | 'returns';
  page?: number;
  pageSize?: number;
}

export function getSalesTransactions(
  db: Database.Database,
  user: UserProfile,
  params: TransactionsQueryParams
): { rows: SalesTransactionRow[]; totalCount: number; page: number; pageSize: number } {
  const activeBatch = getActiveBatch(db);
  if (!activeBatch) {
    return { rows: [], totalCount: 0, page: 1, pageSize: 50 };
  }

  const scope = buildScopeFilter(user, params, activeBatch.id);
  const extraConds: string[] = [scope.sql];
  const extraArgs: unknown[] = [...scope.args];

  if (params.linkStatus === 'linked') {
    extraConds.push('sr.resolved_sector IS NOT NULL AND sr.resolved_branch_code IS NOT NULL');
  } else if (params.linkStatus === 'unlinked_branch') {
    extraConds.push('sr.resolved_branch_code IS NULL');
  } else if (params.linkStatus === 'unlinked_sector') {
    extraConds.push('sr.resolved_sector IS NULL');
  } else if (params.linkStatus === 'returns') {
    extraConds.push('sr.is_return = 1');
  }

  if (params.search && params.search.trim()) {
    const q = `%${params.search.trim()}%`;
    extraConds.push(
      `(sr.doc_number LIKE ? OR sr.payer_code LIKE ? OR sr.payer_name LIKE ? OR sr.sold_to_code LIKE ? OR sr.sold_to_name LIKE ? OR sr.item_code LIKE ? OR sr.item_name LIKE ? OR sr.resolved_branch_name LIKE ? OR sr.resolved_chain_name LIKE ?)`
    );
    for (let i = 0; i < 9; i++) extraArgs.push(q);
  }

  const whereClause = extraConds.join(' AND ');
  const countRow = db
    .prepare(`SELECT COUNT(*) AS total FROM sales_records sr WHERE ${whereClause}`)
    .get(...extraArgs) as { total: number };

  const pageSize = Math.min(Math.max(params.pageSize ?? 50, 10), 200);
  const page = Math.max(params.page ?? 1, 1);
  const offset = (page - 1) * pageSize;

  const dbRows = db
    .prepare(
      `SELECT
        sr.id, sr.source_type, sr.doc_number, sr.billing_date,
        sr.payer_code, sr.payer_name, sr.sold_to_code, sr.sold_to_name,
        sr.item_code, COALESCE(sr.resolved_item_name, sr.item_name) AS item_name,
        COALESCE(sr.resolved_rep_code, sr.rep_code) AS rep_code,
        COALESCE(sr.resolved_rep_name, '') AS rep_name,
        sr.quantity, sr.sales_value, sr.tax_value,
        sr.resolved_sector, sr.resolved_department, sr.resolved_chain_name,
        sr.resolved_branch_code, sr.resolved_branch_name, sr.is_return
       FROM sales_records sr
       WHERE ${whereClause}
       ORDER BY sr.billing_date DESC, sr.id DESC
       LIMIT ? OFFSET ?`
    )
    .all(...extraArgs, pageSize, offset) as Array<{
    id: number;
    source_type: SalesSource;
    doc_number: string;
    billing_date: string;
    payer_code: string;
    payer_name: string;
    sold_to_code: string;
    sold_to_name: string;
    item_code: string;
    item_name: string;
    rep_code: string;
    rep_name: string;
    quantity: number;
    sales_value: number;
    tax_value: number;
    resolved_sector: string | null;
    resolved_department: string | null;
    resolved_chain_name: string | null;
    resolved_branch_code: string | null;
    resolved_branch_name: string | null;
    is_return: number;
  }>;

  return {
    rows: dbRows.map((r) => ({
      id: r.id,
      sourceType: r.source_type,
      docNumber: r.doc_number,
      billingDate: r.billing_date,
      payerCode: r.payer_code,
      payerName: r.payer_name,
      soldToCode: r.sold_to_code,
      soldToName: r.sold_to_name,
      itemCode: r.item_code,
      itemName: r.item_name,
      repCode: r.rep_code,
      repName: r.rep_name,
      quantity: r.quantity,
      salesValue: Number(r.sales_value.toFixed(2)),
      taxValue: Number(r.tax_value.toFixed(2)),
      resolvedSector: r.resolved_sector,
      resolvedDepartment: r.resolved_department,
      resolvedChainName: r.resolved_chain_name,
      resolvedBranchCode: r.resolved_branch_code,
      resolvedBranchName: r.resolved_branch_name,
      isReturn: r.is_return === 1,
    })),
    totalCount: countRow.total,
    page,
    pageSize,
  };
}
