import type Database from 'better-sqlite3';
import type {
  ClassAType,
  CompetitionAnalyticsResponse,
  CompetitionCatalogItem,
  CompetitionItemConfig,
  CompetitionParticipantRow,
  CompetitionSummary,
  UserProfile,
} from '../src/types/bi.ts';
import {
  canonicalizeSectorName,
  normalizeArabic,
  normalizeCode,
  ORIGINAL_SECTORS,
  stripLeadingZeros,
} from './excel-engine.ts';
import { AuthorizationError } from './auth.ts';

export function classifyClassA(rawClassA: string): ClassAType {
  const norm = normalizeArabic(rawClassA).replace(/\s+/g, '');
  if (!norm) return 'other';
  if (norm.includes('صناعي') || norm.includes('صناع')) return 'industrial';
  if (norm.includes('تجاري') || norm.includes('تجار')) return 'commercial';
  return 'other';
}

function normalizeCategoryText(raw: string): string {
  return normalizeArabic(raw).replace(/\s+/g, '');
}

const UNCLASSIFIED_CAT_NORM = normalizeCategoryText('غير مصنف');

function isCategoryScopedSectorManagerSector(sec: string): boolean {
  return sec === 'التوريد المباشر' || sec === 'التوريد المركزي';
}

function isSalesRepJobTitle(jobTitle: string): boolean {
  return normalizeArabic(jobTitle).includes('مندوب');
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

function resolveDepartmentManagerSectors(employeeName: string, rawSectorName: string): string[] {
  const noSpaceName = normalizeArabic(employeeName).replace(/\s+/g, '');
  if (noSpaceName.includes('رافتمحمد')) {
    return ['التوريد المركزي'];
  }
  if (noSpaceName.includes('عبدالستارمصطفي')) {
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

function roundQty(val: number): number {
  if (!Number.isFinite(val)) return 0;
  return Number(val.toFixed(3));
}

export function calculateQtyAchievementPct(soldQty: number, targetQty: number): number | null {
  if (!Number.isFinite(soldQty) || !Number.isFinite(targetQty) || targetQty <= 0) return null;
  const pct = (soldQty / targetQty) * 100;
  if (!Number.isFinite(pct)) return null;
  return Number(pct.toFixed(2));
}

function getActiveBatchId(db: Database.Database): number | null {
  const row = db
    .prepare(`SELECT id FROM import_batches WHERE status = 'active' ORDER BY id DESC LIMIT 1`)
    .get() as { id: number } | undefined;
  return row ? row.id : null;
}

export function getAvailableCatalogItems(db: Database.Database): CompetitionCatalogItem[] {
  const batchId = getActiveBatchId(db);
  if (!batchId) return [];

  const rows = db
    .prepare(
      `SELECT
        item_code,
        item_name,
        COALESCE(category_name, '') AS category_name,
        COALESCE(chains_category, '') AS chains_category,
        COALESCE(army_category, '') AS army_category,
        COALESCE(top_retail_category, '') AS top_retail_category,
        COALESCE(unit_name, '') AS unit_name
       FROM items
       WHERE batch_id = ?
       ORDER BY item_name ASC, item_code ASC`
    )
    .all(batchId) as Array<{
    item_code: string;
    item_name: string;
    category_name: string;
    chains_category: string;
    army_category: string;
    top_retail_category: string;
    unit_name: string;
  }>;

  const seen = new Set<string>();
  const catalog: CompetitionCatalogItem[] = [];
  for (const r of rows) {
    const code = normalizeCode(r.item_code);
    if (!code || seen.has(code)) continue;
    seen.add(code);
    const classA = r.category_name.trim();
    catalog.push({
      itemCode: code,
      itemName: r.item_name.trim() || code,
      categoryName: classA,
      classAType: classifyClassA(classA),
      chainsCategory: r.chains_category.trim(),
      armyCategory: r.army_category.trim(),
      topRetailCategory: r.top_retail_category.trim(),
      unitName: r.unit_name.trim(),
    });
  }
  return catalog;
}

export function listCompetitions(db: Database.Database): CompetitionSummary[] {
  const compRows = db
    .prepare(`SELECT id, name, is_active, created_at, updated_at FROM competitions ORDER BY id DESC`)
    .all() as Array<{
    id: number;
    name: string;
    is_active: number;
    created_at: string;
    updated_at: string;
  }>;

  if (compRows.length === 0) return [];

  const batchId = getActiveBatchId(db);
  const activeItemMetaMap = new Map<string, { itemName: string; classificationA: string; classAType: ClassAType }>();
  if (batchId) {
    const itemRows = db
      .prepare(`SELECT item_code, item_name, COALESCE(category_name, '') AS category_name FROM items WHERE batch_id = ?`)
      .all(batchId) as Array<{ item_code: string; item_name: string; category_name: string }>;
    for (const ir of itemRows) {
      const code = normalizeCode(ir.item_code);
      if (code && !activeItemMetaMap.has(code)) {
        const classA = ir.category_name.trim();
        activeItemMetaMap.set(code, {
          itemName: (ir.item_name || code).trim(),
          classificationA: classA,
          classAType: classifyClassA(classA),
        });
      }
    }
  }

  const allItems = db
    .prepare(`SELECT competition_id, item_code, item_name, equivalent_factor FROM competition_items ORDER BY id ASC`)
    .all() as Array<{
    competition_id: number;
    item_code: string;
    item_name: string;
    equivalent_factor: number;
  }>;

  const targetCounts = db
    .prepare(`SELECT competition_id, COUNT(*) AS cnt FROM competition_targets GROUP BY competition_id`)
    .all() as Array<{ competition_id: number; cnt: number }>;

  const itemsByComp = new Map<number, CompetitionItemConfig[]>();
  for (const it of allItems) {
    const list = itemsByComp.get(it.competition_id) ?? [];
    const code = normalizeCode(it.item_code);
    const meta = activeItemMetaMap.get(code);
    list.push({
      itemCode: code,
      itemName: meta?.itemName || it.item_name || code,
      equivalentFactor: Number(it.equivalent_factor),
      classificationA: meta?.classificationA ?? '',
      classAType: meta?.classAType ?? 'other',
    });
    itemsByComp.set(it.competition_id, list);
  }

  const targetCountMap = new Map<number, number>();
  for (const tc of targetCounts) {
    targetCountMap.set(tc.competition_id, tc.cnt);
  }

  return compRows.map((c) => ({
    id: c.id,
    name: c.name,
    isActive: c.is_active === 1,
    items: itemsByComp.get(c.id) ?? [],
    targetCount: targetCountMap.get(c.id) ?? 0,
    createdAt: c.created_at,
    updatedAt: c.updated_at,
  }));
}

export interface CompetitionItemInput {
  itemCode: string;
  itemName?: string;
  equivalentFactor?: number;
}

export function createCompetition(
  db: Database.Database,
  user: { id: number },
  input: { name: string; items?: CompetitionItemInput[] }
): CompetitionSummary {
  const cleanName = String(input.name || '').trim();
  if (!cleanName) {
    throw new Error('اسم المسابقة مطلوب');
  }

  const now = new Date().toISOString();
  const tx = db.transaction(() => {
    const res = db
      .prepare(
        `INSERT INTO competitions (name, is_active, created_by_id, created_at, updated_at) VALUES (?, 1, ?, ?, ?)`
      )
      .run(cleanName, user.id, now, now);

    const compId = Number(res.lastInsertRowid);
    if (Array.isArray(input.items) && input.items.length > 0) {
      upsertCompetitionItemsTx(db, compId, input.items);
    }
    return compId;
  });

  const compId = tx();
  return listCompetitions(db).find((c) => c.id === compId)!;
}

function upsertCompetitionItemsTx(
  db: Database.Database,
  competitionId: number,
  items: CompetitionItemInput[]
): void {
  db.prepare(`DELETE FROM competition_items WHERE competition_id = ?`).run(competitionId);
  const insertItem = db.prepare(
    `INSERT INTO competition_items (competition_id, item_code, item_name, equivalent_factor) VALUES (?, ?, ?, ?)`
  );
  const seen = new Set<string>();
  for (const raw of items) {
    const code = normalizeCode(raw.itemCode);
    if (!code || seen.has(code)) continue;
    seen.add(code);
    const factorNum =
      raw.equivalentFactor === undefined || raw.equivalentFactor === null
        ? 1
        : Number(raw.equivalentFactor);
    if (!Number.isFinite(factorNum) || factorNum < 0) {
      throw new Error(`معامل الشدة غير صالح للصنف ${code}`);
    }
    insertItem.run(competitionId, code, String(raw.itemName || code).trim(), factorNum);
  }
}

export function updateCompetition(
  db: Database.Database,
  competitionId: number,
  input: { name?: string; isActive?: boolean; items?: CompetitionItemInput[] }
): CompetitionSummary {
  const existing = db
    .prepare(`SELECT id, name, is_active FROM competitions WHERE id = ?`)
    .get(competitionId) as { id: number; name: string; is_active: number } | undefined;
  if (!existing) {
    throw new Error('المسابقة غير موجودة');
  }

  const nextName = input.name !== undefined ? String(input.name).trim() : existing.name;
  if (!nextName) {
    throw new Error('اسم المسابقة مطلوب');
  }
  const nextActive = input.isActive !== undefined ? (input.isActive ? 1 : 0) : existing.is_active;
  const now = new Date().toISOString();

  const tx = db.transaction(() => {
    db.prepare(`UPDATE competitions SET name = ?, is_active = ?, updated_at = ? WHERE id = ?`).run(
      nextName,
      nextActive,
      now,
      competitionId
    );
    if (Array.isArray(input.items)) {
      upsertCompetitionItemsTx(db, competitionId, input.items);
    }
  });
  tx();

  return listCompetitions(db).find((c) => c.id === competitionId)!;
}

export function deleteCompetition(db: Database.Database, competitionId: number): void {
  const res = db.prepare(`DELETE FROM competitions WHERE id = ?`).run(competitionId);
  if (res.changes === 0) {
    throw new Error('المسابقة غير موجودة');
  }
}

export interface CompetitionTargetInput {
  participantKey: string;
  targetQty?: number;
  industrialTargetQty?: number;
  commercialTargetQty?: number;
}

export function saveCompetitionTargets(
  db: Database.Database,
  competitionId: number,
  targets: CompetitionTargetInput[],
  user?: UserProfile
): void {
  const exists = db.prepare(`SELECT id FROM competitions WHERE id = ?`).get(competitionId);
  if (!exists) {
    throw new Error('المسابقة غير موجودة');
  }

  if (user && user.role !== 'admin') {
    if (user.role !== 'sector_manager') {
      throw new AuthorizationError('صلاحية تعديل أهداف المسابقات متاحة للأدمن ومدير القطاع أو الإدارة فقط');
    }
    const authSet = new Set(resolveAuthorizedSectors(user));
    for (const row of targets) {
      const key = String(row.participantKey || '').trim();
      if (!key) continue;
      const parts = key.split('::');
      const prefix = parts[0];
      const rawSec = parts[1];
      if (!['REP', 'SUP', 'SEC_MGR'].includes(prefix) || !rawSec) {
        throw new AuthorizationError('غير مصرح لك بتعديل هذا الهدف خارج نطاق صلاحياتك');
      }
      const canonSec = canonicalizeSectorName(rawSec);
      if (!canonSec || !authSet.has(canonSec)) {
        throw new AuthorizationError('غير مصرح لك بتعديل أهداف قطاع خارج نطاق صلاحياتك');
      }
    }
  }

  const now = new Date().toISOString();
  const tx = db.transaction(() => {
    const upsertStmt = db.prepare(
      `INSERT INTO competition_targets (
         competition_id, participant_key, target_qty, industrial_target_qty, commercial_target_qty
       )
       VALUES (?, ?, ?, ?, ?)
       ON CONFLICT(competition_id, participant_key)
       DO UPDATE SET
         target_qty = excluded.target_qty,
         industrial_target_qty = excluded.industrial_target_qty,
         commercial_target_qty = excluded.commercial_target_qty`
    );
    for (const row of targets) {
      const key = String(row.participantKey || '').trim();
      if (!key) continue;

      const hasSplit = row.industrialTargetQty !== undefined || row.commercialTargetQty !== undefined;
      if (hasSplit) {
        const indQty = Number(row.industrialTargetQty ?? 0);
        const comQty = Number(row.commercialTargetQty ?? 0);
        if (!Number.isFinite(indQty) || indQty < 0) {
          throw new Error(`قيمة هدف أصناف الصناعى غير صالحة للمشارك ${key}`);
        }
        if (!Number.isFinite(comQty) || comQty < 0) {
          throw new Error(`قيمة هدف الأصناف التجارى غير صالحة للمشارك ${key}`);
        }
        const cleanInd = roundQty(indQty);
        const cleanCom = roundQty(comQty);
        const total = roundQty(cleanInd + cleanCom);
        upsertStmt.run(competitionId, key, total, cleanInd, cleanCom);
      } else {
        const qty = Number(row.targetQty ?? 0);
        if (!Number.isFinite(qty) || qty < 0) {
          throw new Error(`قيمة الهدف الكمي غير صالحة للمشارك ${key}`);
        }
        upsertStmt.run(competitionId, key, roundQty(qty), 0, 0);
      }
    }
    db.prepare(`UPDATE competitions SET updated_at = ? WHERE id = ?`).run(now, competitionId);
  });
  tx();
}

export interface CompetitionFilterParams {
  sector?: string;
  startDate?: string;
  endDate?: string;
}

interface SplitBucket {
  ind: number;
  com: number;
  total: number;
}

const ZERO_SPLIT: SplitBucket = { ind: 0, com: 0, total: 0 };

function addSplitBucket(map: Map<string, SplitBucket>, key: string, ind: number, com: number, total: number): void {
  const b = map.get(key) ?? { ind: 0, com: 0, total: 0 };
  b.ind += ind;
  b.com += com;
  b.total += total;
  map.set(key, b);
}

interface TargetSplit {
  ind: number;
  com: number;
  total: number;
}

export function getCompetitionAnalytics(
  db: Database.Database,
  user: UserProfile,
  competitionId: number | null,
  params: CompetitionFilterParams = {}
): CompetitionAnalyticsResponse {
  const competitions = listCompetitions(db);
  const catalogItems = getAvailableCatalogItems(db);

  const selectedComp =
    (competitionId ? competitions.find((c) => c.id === competitionId) : undefined) ??
    competitions[0] ??
    null;

  if (!selectedComp) {
    return {
      competition: null,
      competitions: [],
      catalogItems,
      participants: [],
      summary: {
        totalSelectedItems: 0,
        totalTeamMembers: 0,
        totalStandardSoldQty: 0,
        totalTargetQty: 0,
        overallAchievementPct: null,
        totalIndustrialSoldQty: 0,
        totalCommercialSoldQty: 0,
        totalIndustrialTargetQty: 0,
        totalCommercialTargetQty: 0,
        industrialAchievementPct: null,
        commercialAchievementPct: null,
      },
    };
  }

  const batchId = getActiveBatchId(db);
  if (!batchId) {
    return {
      competition: selectedComp,
      competitions,
      catalogItems,
      participants: [],
      summary: {
        totalSelectedItems: selectedComp.items.length,
        totalTeamMembers: 0,
        totalStandardSoldQty: 0,
        totalTargetQty: 0,
        overallAchievementPct: null,
        totalIndustrialSoldQty: 0,
        totalCommercialSoldQty: 0,
        totalIndustrialTargetQty: 0,
        totalCommercialTargetQty: 0,
        industrialAchievementPct: null,
        commercialAchievementPct: null,
      },
    };
  }

  const authorizedSectors = resolveAuthorizedSectors(user);
  const selectedSector =
    params.sector && params.sector !== 'ALL' && params.sector !== 'undefined' && params.sector !== 'null'
      ? params.sector === 'UNLINKED'
        ? 'UNLINKED'
        : canonicalizeSectorName(params.sector)
      : 'ALL';

  const visibleSectors =
    selectedSector === 'UNLINKED'
      ? []
      : ORIGINAL_SECTORS.filter((sec) => {
          if (user.role !== 'admin' && !authorizedSectors.includes(sec)) return false;
          if (selectedSector !== 'ALL' && selectedSector !== sec) return false;
          return true;
        });

  const isSectorVisible = (sec: string): boolean => {
    if (!sec) return selectedSector === 'ALL';
    if (user.role !== 'admin' && !authorizedSectors.includes(sec)) return false;
    if (selectedSector !== 'ALL' && sec !== selectedSector) return false;
    return true;
  };

  // Load saved targets (split industrial + commercial + total)
  const targetRows = db
    .prepare(
      `SELECT participant_key, target_qty, COALESCE(industrial_target_qty, 0) AS industrial_target_qty, COALESCE(commercial_target_qty, 0) AS commercial_target_qty
       FROM competition_targets
       WHERE competition_id = ?`
    )
    .all(selectedComp.id) as Array<{
    participant_key: string;
    target_qty: number;
    industrial_target_qty: number;
    commercial_target_qty: number;
  }>;

  const savedTargetMap = new Map<string, TargetSplit>();
  for (const tr of targetRows) {
    const ind = roundQty(Number(tr.industrial_target_qty || 0));
    const com = roundQty(Number(tr.commercial_target_qty || 0));
    const total = ind > 0 || com > 0 ? roundQty(ind + com) : roundQty(Number(tr.target_qty || 0));
    savedTargetMap.set(tr.participant_key, { ind, com, total });
  }
  const getSavedTarget = (key: string, legacyKey?: string): TargetSplit => {
    return (
      savedTargetMap.get(key) ??
      (legacyKey ? savedTargetMap.get(legacyKey) : undefined) ?? { ind: 0, com: 0, total: 0 }
    );
  };

  // Load items catalog (`التصنيف A` -> category_name, `تصنيف السلاسل` -> chains_category)
  const itemRows = db
    .prepare(
      `SELECT
        item_code,
        COALESCE(category_name, '') AS category_name,
        COALESCE(chains_category, '') AS chains_category
       FROM items
       WHERE batch_id = ?`
    )
    .all(batchId) as Array<{ item_code: string; category_name: string; chains_category: string }>;

  const itemByExact = new Map<string, { classAType: ClassAType; chainsCategory: string; normChainsCat: string }>();
  const itemByTrimmed = new Map<string, { classAType: ClassAType; chainsCategory: string; normChainsCat: string }>();
  for (const it of itemRows) {
    const exact = normalizeCode(it.item_code);
    if (!exact) continue;
    const cleanChains = it.chains_category.replace(/\s+/g, ' ').trim();
    const meta = {
      classAType: classifyClassA(it.category_name),
      chainsCategory: cleanChains,
      normChainsCat: cleanChains ? normalizeCategoryText(cleanChains) : UNCLASSIFIED_CAT_NORM,
    };
    if (!itemByExact.has(exact)) itemByExact.set(exact, meta);
    const trimmed = stripLeadingZeros(exact);
    if (trimmed && !itemByTrimmed.has(trimmed)) itemByTrimmed.set(trimmed, meta);
  }

  // Load hierarchy metadata from active batch
  const repRows = db
    .prepare(
      `SELECT rep_code, rep_name, COALESCE(sector_name, '') AS sector_name, COALESCE(supervisor_name, '') AS supervisor_name
       FROM sales_reps
       WHERE batch_id = ?`
    )
    .all(batchId) as Array<{ rep_code: string; rep_name: string; sector_name: string; supervisor_name: string }>;

  const repByExact = new Map<string, (typeof repRows)[number]>();
  const repByTrimmed = new Map<string, (typeof repRows)[number]>();
  for (const r of repRows) {
    const exact = normalizeCode(r.rep_code);
    if (!exact) continue;
    if (!repByExact.has(exact)) repByExact.set(exact, r);
    const trimmed = stripLeadingZeros(exact);
    if (trimmed && !repByTrimmed.has(trimmed)) repByTrimmed.set(trimmed, r);
  }
  const resolveRep = (code: string) => {
    const exact = normalizeCode(code);
    if (!exact) return undefined;
    return repByExact.get(exact) ?? repByTrimmed.get(stripLeadingZeros(exact));
  };

  const chainRows = db
    .prepare(
      `SELECT chain_code, chain_name, department_name, sector_name, COALESCE(supervisor_name, '') AS supervisor_name
       FROM departments_chains
       WHERE batch_id = ?`
    )
    .all(batchId) as Array<{
    chain_code: string;
    chain_name: string;
    department_name: string;
    sector_name: string;
    supervisor_name: string;
  }>;

  const chainByExact = new Map<string, (typeof chainRows)[number]>();
  const chainByTrimmed = new Map<string, (typeof chainRows)[number]>();
  const chainByNormName = new Map<string, (typeof chainRows)[number]>();
  const chainSupExists = new Set<string>(); // `${sec}::${supNorm}`
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
    const supNorm = normalizeArabic(c.supervisor_name);
    if (sec && supNorm) {
      chainSupExists.add(`${sec}::${supNorm}`);
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

  const branchRows = db
    .prepare(
      `SELECT sector_name, COALESCE(rep_code, '') AS rep_code, COALESCE(supervisor_name, '') AS supervisor_name
       FROM branches
       WHERE batch_id = ?`
    )
    .all(batchId) as Array<{ sector_name: string; rep_code: string; supervisor_name: string }>;

  const branchSupRepsMap = new Map<string, Set<string>>();
  for (const b of branchRows) {
    const sec = canonicalizeSectorName(b.sector_name);
    const supNorm = normalizeArabic(b.supervisor_name);
    const repDef = resolveRep(b.rep_code);
    const repKey = normalizeCode(repDef?.rep_code || b.rep_code);
    if (sec && supNorm) {
      const k = `${sec}::${supNorm}`;
      const set = branchSupRepsMap.get(k) ?? new Set<string>();
      if (repKey) set.add(repKey);
      branchSupRepsMap.set(k, set);
    }
  }

  const repSupRepsMap = new Map<string, Set<string>>();
  for (const r of repRows) {
    const sec = canonicalizeSectorName(r.sector_name);
    const supNorm = normalizeArabic(r.supervisor_name);
    const repKey = normalizeCode(r.rep_code);
    if (sec && supNorm && repKey) {
      const k = `${sec}::${supNorm}`;
      const set = repSupRepsMap.get(k) ?? new Set<string>();
      set.add(repKey);
      repSupRepsMap.set(k, set);
    }
  }

  // Load roster + assigned categories from `category_targets`
  const catTargetRoles = db
    .prepare(
      `SELECT DISTINCT sector_name, job_title, employee_code, employee_name, COALESCE(category_name, '') AS category_name
       FROM category_targets
       WHERE batch_id = ?`
    )
    .all(batchId) as Array<{
    sector_name: string;
    job_title: string;
    employee_code: string;
    employee_name: string;
    category_name: string;
  }>;

  // Build item equivalent factor lookup
  const factorByExact = new Map<string, number>();
  const factorByTrimmed = new Map<string, number>();
  for (const ci of selectedComp.items) {
    const exact = normalizeCode(ci.itemCode);
    if (!exact) continue;
    factorByExact.set(exact, Number(ci.equivalentFactor));
    const trimmed = stripLeadingZeros(exact);
    if (trimmed && !factorByTrimmed.has(trimmed)) {
      factorByTrimmed.set(trimmed, Number(ci.equivalentFactor));
    }
  }

  // Query Zsales_Price (2) (`source_type = 'ZSALES'`) within authorized sectors and date range
  const salesConds: string[] = [`sr.batch_id = ?`, `sr.source_type = 'ZSALES'`];
  const salesArgs: unknown[] = [batchId];
  if (user.role !== 'admin') {
    if (authorizedSectors.length === 0) {
      salesConds.push(`1 = 0`);
    } else {
      salesConds.push(`sr.resolved_sector IN (${authorizedSectors.map(() => '?').join(', ')})`);
      salesArgs.push(...authorizedSectors);
    }
  }
  if (params.startDate) {
    salesConds.push(`sr.billing_date >= ?`);
    salesArgs.push(params.startDate);
  }
  if (params.endDate) {
    salesConds.push(`sr.billing_date <= ?`);
    salesArgs.push(params.endDate);
  }

  const zsalesRows = db
    .prepare(
      `SELECT
        COALESCE(sr.resolved_sector, '') AS sector_name,
        COALESCE(sr.payer_code, '') AS payer_code,
        COALESCE(NULLIF(sr.resolved_chain_name, ''), NULLIF(sr.payer_name, ''), '') AS chain_name,
        COALESCE(NULLIF(sr.resolved_rep_code, ''), sr.rep_code, '') AS rep_code,
        COALESCE(NULLIF(sr.resolved_rep_name, ''), '') AS rep_name,
        COALESCE(sr.item_code, '') AS item_code,
        COALESCE(SUM(sr.quantity), 0) AS qty
       FROM sales_records sr
       WHERE ${salesConds.join(' AND ')}
       GROUP BY sr.resolved_sector, sr.payer_code, COALESCE(NULLIF(sr.resolved_chain_name, ''), NULLIF(sr.payer_name, ''), ''), COALESCE(NULLIF(sr.resolved_rep_code, ''), sr.rep_code, ''), COALESCE(NULLIF(sr.resolved_rep_name, ''), ''), sr.item_code`
    )
    .all(...salesArgs) as Array<{
    sector_name: string;
    payer_code: string;
    chain_name: string;
    rep_code: string;
    rep_name: string;
    item_code: string;
    qty: number;
  }>;

  const sectorSplitMap = new Map<string, SplitBucket>(); // sec
  const sectorChainsCatSplitMap = new Map<string, SplitBucket>(); // `${sec}::${normChainsCat}`
  const deptChainsCatSplitMap = new Map<string, SplitBucket>(); // `${sec}::${deptNorm}::${normChainsCat}`
  const chainChainsCatSplitMap = new Map<string, SplitBucket>(); // `${sec}::${chainNorm}::${normChainsCat}`
  const repSplitMap = new Map<string, SplitBucket>(); // `${sec}::${repKey}`
  const supSplitMap = new Map<string, SplitBucket>(); // `${sec}::${supNorm}`
  const activeRepsBySector = new Map<string, Map<string, { repCode: string; repName: string }>>();

  for (const r of zsalesRows) {
    const sec = canonicalizeSectorName(r.sector_name);
    if (!sec) continue;
    const rawItem = normalizeCode(r.item_code);
    const factor = rawItem
      ? factorByExact.get(rawItem) ?? factorByTrimmed.get(stripLeadingZeros(rawItem))
      : undefined;

    const rawRep = normalizeCode(r.rep_code);
    const repDef = rawRep ? resolveRep(rawRep) : undefined;
    const repKey = repDef ? normalizeCode(repDef.rep_code) : rawRep;
    if (repKey) {
      const secReps = activeRepsBySector.get(sec) ?? new Map<string, { repCode: string; repName: string }>();
      if (!secReps.has(repKey)) {
        secReps.set(repKey, {
          repCode: repKey,
          repName: (repDef?.rep_name || r.rep_name || repKey).trim(),
        });
      }
      activeRepsBySector.set(sec, secReps);
    }

    if (factor === undefined) continue; // Item not in competition
    const stdQty = Number(r.qty || 0) * factor;
    const itemMeta = rawItem
      ? itemByExact.get(rawItem) ?? itemByTrimmed.get(stripLeadingZeros(rawItem))
      : undefined;
    const classAType = itemMeta?.classAType ?? 'other';
    const indAdd = classAType === 'industrial' ? stdQty : 0;
    const comAdd = classAType === 'commercial' ? stdQty : 0;
    const normChainsCat = itemMeta?.normChainsCat ?? UNCLASSIFIED_CAT_NORM;

    addSplitBucket(sectorSplitMap, sec, indAdd, comAdd, stdQty);
    addSplitBucket(sectorChainsCatSplitMap, `${sec}::${normChainsCat}`, indAdd, comAdd, stdQty);

    if (repKey) {
      addSplitBucket(repSplitMap, `${sec}::${repKey}`, indAdd, comAdd, stdQty);
    }

    const normPayer = normalizeCode(r.payer_code);
    const chainDef =
      (normPayer ? chainByExact.get(normPayer) ?? chainByTrimmed.get(stripLeadingZeros(normPayer)) : undefined) ??
      (r.chain_name ? chainByNormName.get(normalizeArabic(r.chain_name)) : undefined);
    const supNorm = normalizeArabic(chainDef?.supervisor_name || repDef?.supervisor_name || '');
    if (supNorm) {
      addSplitBucket(supSplitMap, `${sec}::${supNorm}`, indAdd, comAdd, stdQty);
    }
    const deptNorm = normalizeArabic(chainDef?.department_name || '');
    if (deptNorm) {
      addSplitBucket(deptChainsCatSplitMap, `${sec}::${deptNorm}::${normChainsCat}`, indAdd, comAdd, stdQty);
    }
    const chainNorm = normalizeArabic(chainDef?.chain_name || r.chain_name || '');
    if (chainNorm) {
      addSplitBucket(chainChainsCatSplitMap, `${sec}::${chainNorm}::${normChainsCat}`, indAdd, comAdd, stdQty);
    }
  }

  // ==================== BUILD TEAM HIERARCHY & PARTICIPANT LIST ====================
  // Note: Compute all authorized-sector subordinates first so `مدير إدارة` auto-rollup works even when filtered to one managed sector.
  const authSet = new Set(user.role === 'admin' ? ORIGINAL_SECTORS : authorizedSectors);

  // A. Sales Reps across all authorized sectors
  interface InternalRep {
    participantKey: string;
    employeeCode: string;
    employeeName: string;
    sectorName: string;
    target: TargetSplit;
    sold: SplitBucket;
  }
  const allAuthReps: InternalRep[] = [];
  const repSeenBySec = new Set<string>();

  for (const r of repRows) {
    const sec = canonicalizeSectorName(r.sector_name);
    if (!sec || !authSet.has(sec)) continue;
    const rk = normalizeCode(r.rep_code);
    if (!rk) continue;
    const dedupKey = `${sec}::${rk}`;
    if (repSeenBySec.has(dedupKey)) continue;
    repSeenBySec.add(dedupKey);
    const participantKey = `REP::${sec}::${rk}`;
    const legacyKey = `REP::${rk}`;
    allAuthReps.push({
      participantKey,
      employeeCode: rk,
      employeeName: r.rep_name.trim() || rk,
      sectorName: sec,
      target: getSavedTarget(participantKey, legacyKey),
      sold: repSplitMap.get(dedupKey) ?? ZERO_SPLIT,
    });
  }
  for (const [sec, map] of activeRepsBySector.entries()) {
    if (!authSet.has(sec)) continue;
    for (const [rk, info] of map.entries()) {
      const dedupKey = `${sec}::${rk}`;
      if (repSeenBySec.has(dedupKey)) continue;
      repSeenBySec.add(dedupKey);
      const participantKey = `REP::${sec}::${rk}`;
      const legacyKey = `REP::${rk}`;
      allAuthReps.push({
        participantKey,
        employeeCode: rk,
        employeeName: info.repName,
        sectorName: sec,
        target: getSavedTarget(participantKey, legacyKey),
        sold: repSplitMap.get(dedupKey) ?? ZERO_SPLIT,
      });
    }
  }

  // B. Supervisors across all authorized sectors
  interface InternalSup {
    participantKey: string;
    employeeName: string;
    roleLabel: string;
    sectorName: string;
    normName: string;
    target: TargetSplit;
    sold: SplitBucket;
  }
  const supMap = new Map<string, InternalSup>();
  const registerSupervisor = (sec: string, rawName: string, roleLabel = 'مشرف قطاع') => {
    const cleanName = rawName.trim();
    const norm = normalizeArabic(cleanName);
    if (!sec || !norm || !authSet.has(sec)) return;
    const key = `${sec}::${norm}`;
    if (supMap.has(key)) return;
    const participantKey = `SUP::${sec}::${norm}`;
    const sold: SplitBucket = { ind: 0, com: 0, total: 0 };
    if (chainSupExists.has(key)) {
      const s = supSplitMap.get(key);
      if (s) {
        sold.ind = s.ind;
        sold.com = s.com;
        sold.total = s.total;
      }
    } else if (branchSupRepsMap.has(key)) {
      for (const rk of branchSupRepsMap.get(key)!) {
        const rs = repSplitMap.get(`${sec}::${rk}`);
        if (rs) {
          sold.ind += rs.ind;
          sold.com += rs.com;
          sold.total += rs.total;
        }
      }
    } else if (repSupRepsMap.has(key)) {
      for (const rk of repSupRepsMap.get(key)!) {
        const rs = repSplitMap.get(`${sec}::${rk}`);
        if (rs) {
          sold.ind += rs.ind;
          sold.com += rs.com;
          sold.total += rs.total;
        }
      }
    } else {
      const s = supSplitMap.get(key);
      if (s) {
        sold.ind = s.ind;
        sold.com = s.com;
        sold.total = s.total;
      }
    }
    supMap.set(key, {
      participantKey,
      employeeName: cleanName,
      roleLabel,
      sectorName: sec,
      normName: norm,
      target: getSavedTarget(participantKey),
      sold,
    });
  };

  for (const ct of catTargetRoles) {
    if (!isSalesRepJobTitle(ct.job_title) && !isDepartmentManagerJobTitle(ct.job_title) && normalizeArabic(ct.job_title).includes('مشرف')) {
      registerSupervisor(canonicalizeSectorName(ct.sector_name), ct.employee_name, ct.job_title.trim() || 'مشرف قطاع');
    }
  }
  for (const c of chainRows) {
    registerSupervisor(canonicalizeSectorName(c.sector_name), c.supervisor_name, 'مشرف قطاع');
  }
  for (const b of branchRows) {
    registerSupervisor(canonicalizeSectorName(b.sector_name), b.supervisor_name, 'مشرف قطاع');
  }
  for (const r of repRows) {
    registerSupervisor(canonicalizeSectorName(r.sector_name), r.supervisor_name, 'مشرف قطاع');
  }

  // C. Sector Managers (`مدير قطاع`) across all authorized sectors
  //    In `التوريد المباشر` and `التوريد المركزي`: scoped STRICTLY to their assigned categories from `category_targets` (matched via `chains_category`),
  //    whereas `الجيش` and `التوب ريتيل` reflect their entire sector.
  interface InternalSecMgr {
    participantKey: string;
    employeeCode: string;
    employeeName: string;
    normName: string;
    roleLabel: string;
    sectorName: string;
    assignedCats: Map<string, string>; // normCat -> displayCat
    target: TargetSplit;
    sold: SplitBucket;
  }
  const secMgrMap = new Map<string, InternalSecMgr>();
  for (const ct of catTargetRoles) {
    if (isSalesRepJobTitle(ct.job_title) || isDepartmentManagerJobTitle(ct.job_title)) continue;
    if (normalizeArabic(ct.job_title).includes('مشرف')) continue;
    const sec = canonicalizeSectorName(ct.sector_name);
    const empName = ct.employee_name.trim() || normalizeCode(ct.employee_code);
    const normEmp = normalizeArabic(empName);
    if (!sec || !normEmp || !authSet.has(sec)) continue;
    const dedup = `${sec}::${normEmp}`;
    let mgr = secMgrMap.get(dedup);
    if (!mgr) {
      const participantKey = `SEC_MGR::${sec}::${normEmp}`;
      mgr = {
        participantKey,
        employeeCode: normalizeCode(ct.employee_code),
        employeeName: empName,
        normName: normEmp,
        roleLabel: ct.job_title.trim() || 'مدير قطاع',
        sectorName: sec,
        assignedCats: new Map<string, string>(),
        target: getSavedTarget(participantKey),
        sold: { ind: 0, com: 0, total: 0 },
      };
      secMgrMap.set(dedup, mgr);
    }
    const rawCat = ct.category_name.replace(/\s+/g, ' ').trim();
    const normCat = normalizeCategoryText(rawCat);
    if (rawCat && normCat && !mgr.assignedCats.has(normCat)) {
      mgr.assignedCats.set(normCat, rawCat);
    }
  }

  // Track category assignee multiplicity within each (sector, jobTitle) for Direct & Central Supply
  const secMgrCatAssigneeCount = new Map<string, number>();
  for (const mgr of secMgrMap.values()) {
    if (!isCategoryScopedSectorManagerSector(mgr.sectorName)) continue;
    const roleNorm = normalizeArabic(mgr.roleLabel);
    for (const normCat of mgr.assignedCats.keys()) {
      const k = `${mgr.sectorName}::${roleNorm}::${normCat}`;
      secMgrCatAssigneeCount.set(k, (secMgrCatAssigneeCount.get(k) ?? 0) + 1);
    }
  }

  for (const mgr of secMgrMap.values()) {
    if (isCategoryScopedSectorManagerSector(mgr.sectorName) && mgr.assignedCats.size > 0) {
      const roleNorm = normalizeArabic(mgr.roleLabel);
      const sum: SplitBucket = { ind: 0, com: 0, total: 0 };
      for (const normCat of mgr.assignedCats.keys()) {
        const assigneeCount = secMgrCatAssigneeCount.get(`${mgr.sectorName}::${roleNorm}::${normCat}`) ?? 0;
        let bucket: SplitBucket | undefined;
        if (assigneeCount > 1 && departmentsBySector.get(mgr.sectorName)?.has(mgr.normName)) {
          bucket = deptChainsCatSplitMap.get(`${mgr.sectorName}::${mgr.normName}::${normCat}`);
        } else if (assigneeCount > 1 && chainsBySector.get(mgr.sectorName)?.has(mgr.normName)) {
          bucket = chainChainsCatSplitMap.get(`${mgr.sectorName}::${mgr.normName}::${normCat}`);
        } else {
          bucket = sectorChainsCatSplitMap.get(`${mgr.sectorName}::${normCat}`);
        }
        if (bucket) {
          sum.ind += bucket.ind;
          sum.com += bucket.com;
          sum.total += bucket.total;
        }
      }
      mgr.sold = sum;
    } else {
      mgr.sold = sectorSplitMap.get(mgr.sectorName) ?? ZERO_SPLIT;
    }
  }

  // Per-sector target totals by hierarchy level (used by `مدير إدارة` auto-rollup)
  const sectorRollupTarget = (sec: string): TargetSplit => {
    let secMgrInd = 0;
    let secMgrCom = 0;
    let secMgrTotal = 0;
    for (const sm of secMgrMap.values()) {
      if (sm.sectorName === sec) {
        secMgrInd += sm.target.ind;
        secMgrCom += sm.target.com;
        secMgrTotal += sm.target.total;
      }
    }
    if (secMgrTotal > 0 || secMgrInd > 0 || secMgrCom > 0) {
      return { ind: secMgrInd, com: secMgrCom, total: secMgrTotal };
    }

    let supInd = 0;
    let supCom = 0;
    let supTotal = 0;
    for (const sp of supMap.values()) {
      if (sp.sectorName === sec) {
        supInd += sp.target.ind;
        supCom += sp.target.com;
        supTotal += sp.target.total;
      }
    }
    if (supTotal > 0 || supInd > 0 || supCom > 0) {
      return { ind: supInd, com: supCom, total: supTotal };
    }

    let repInd = 0;
    let repCom = 0;
    let repTotal = 0;
    for (const rp of allAuthReps) {
      if (rp.sectorName === sec) {
        repInd += rp.target.ind;
        repCom += rp.target.com;
        repTotal += rp.target.total;
      }
    }
    return { ind: repInd, com: repCom, total: repTotal };
  };

  // D. Department Managers (`مدير إدارة`)
  interface InternalDeptMgr {
    participantKey: string;
    employeeCode: string;
    employeeName: string;
    roleLabel: string;
    managedSectors: string[];
    target: TargetSplit;
    sold: SplitBucket;
  }
  const deptMgrMap = new Map<string, InternalDeptMgr>();

  const registerDeptMgr = (rawName: string, rawCode: string, roleLabel: string, rawSector: string) => {
    const cleanName = rawName.trim() || normalizeCode(rawCode);
    const noSpace = normalizeArabic(cleanName).replace(/\s+/g, '');
    if (!noSpace) return;
    if (deptMgrMap.has(noSpace)) return;
    const allManagedSectors = resolveDepartmentManagerSectors(cleanName, rawSector);
    if (user.role !== 'admin' && !allManagedSectors.every((s) => authSet.has(s))) return;
    const managedSectors = allManagedSectors.filter((s) => authSet.has(s));
    if (managedSectors.length === 0) return;

    const target: TargetSplit = { ind: 0, com: 0, total: 0 };
    const sold: SplitBucket = { ind: 0, com: 0, total: 0 };
    for (const sec of managedSectors) {
      const rt = sectorRollupTarget(sec);
      target.ind += rt.ind;
      target.com += rt.com;
      target.total += rt.total;

      const ss = sectorSplitMap.get(sec);
      if (ss) {
        sold.ind += ss.ind;
        sold.com += ss.com;
        sold.total += ss.total;
      }
    }

    deptMgrMap.set(noSpace, {
      participantKey: `DEPT_MGR::${normalizeArabic(cleanName)}`,
      employeeCode: normalizeCode(rawCode),
      employeeName: cleanName,
      roleLabel: roleLabel || 'مدير إدارة',
      managedSectors,
      target,
      sold,
    });
  };

  for (const ct of catTargetRoles) {
    if (isDepartmentManagerJobTitle(ct.job_title)) {
      registerDeptMgr(ct.employee_name, ct.employee_code, ct.job_title.trim(), ct.sector_name);
    }
  }
  // Fallback: if workbook has no `اهداف الفئات` sheet, include standard department managers if their sectors exist
  if (deptMgrMap.size === 0) {
    const sectorsPresent = new Set<string>([
      ...allAuthReps.map((r) => r.sectorName),
      ...Array.from(sectorSplitMap.keys()),
    ]);
    if (sectorsPresent.has('التوريد المركزي')) {
      registerDeptMgr('رافت محمد', '', 'مدير إدارة', 'التوريد المركزي');
    }
    if (['التوريد المباشر', 'الجيش', 'التوب ريتيل'].some((s) => sectorsPresent.has(s))) {
      registerDeptMgr('عبدالستار مصطفى', '', 'مدير إدارة', 'التوريد المباشر + الجيش + التوب ريتيل');
    }
  }

  // ==================== ASSEMBLE FILTERED PARTICIPANTS ====================
  const participants: CompetitionParticipantRow[] = [];

  const toRow = (base: {
    participantKey: string;
    employeeCode: string;
    employeeName: string;
    roleLabel: string;
    roleKind: CompetitionParticipantRow['roleKind'];
    sectorName: string;
    managedSectors: string[];
    assignedCategories: string[];
    isAutoTarget: boolean;
    target: TargetSplit;
    sold: SplitBucket;
  }): CompetitionParticipantRow => {
    const indSold = roundQty(base.sold.ind);
    const comSold = roundQty(base.sold.com);
    const totalSold = roundQty(base.sold.total);
    const indTarget = roundQty(base.target.ind);
    const comTarget = roundQty(base.target.com);
    const totalTarget = roundQty(base.target.total);
    const remaining = totalTarget > 0 ? Math.max(0, roundQty(totalTarget - totalSold)) : 0;

    return {
      participantKey: base.participantKey,
      employeeCode: base.employeeCode,
      employeeName: base.employeeName,
      roleLabel: base.roleLabel,
      roleKind: base.roleKind,
      sectorName: base.sectorName,
      managedSectors: base.managedSectors,
      assignedCategories: base.assignedCategories,
      isAutoTarget: base.isAutoTarget,
      targetQty: totalTarget,
      soldStandardQty: totalSold,
      remainingQty: remaining,
      achievementPct: calculateQtyAchievementPct(totalSold, totalTarget),
      industrialSoldQty: indSold,
      commercialSoldQty: comSold,
      industrialTargetQty: indTarget,
      commercialTargetQty: comTarget,
      industrialAchievementPct: calculateQtyAchievementPct(indSold, indTarget),
      commercialAchievementPct: calculateQtyAchievementPct(comSold, comTarget),
    };
  };

  for (const dm of deptMgrMap.values()) {
    const isVisible =
      selectedSector === 'UNLINKED'
        ? false
        : selectedSector === 'ALL' || dm.managedSectors.includes(selectedSector);
    if (!isVisible) continue;
    participants.push(
      toRow({
        participantKey: dm.participantKey,
        employeeCode: dm.employeeCode,
        employeeName: dm.employeeName,
        roleLabel: dm.roleLabel,
        roleKind: 'department_manager',
        sectorName: dm.managedSectors.join(' + '),
        managedSectors: dm.managedSectors,
        assignedCategories: [],
        isAutoTarget: true,
        target: dm.target,
        sold: dm.sold,
      })
    );
  }

  for (const sm of secMgrMap.values()) {
    if (!isSectorVisible(sm.sectorName)) continue;
    participants.push(
      toRow({
        participantKey: sm.participantKey,
        employeeCode: sm.employeeCode,
        employeeName: sm.employeeName,
        roleLabel: sm.roleLabel,
        roleKind: 'sector_manager',
        sectorName: sm.sectorName,
        managedSectors: [sm.sectorName],
        assignedCategories: isCategoryScopedSectorManagerSector(sm.sectorName)
          ? Array.from(sm.assignedCats.values())
          : [],
        isAutoTarget: false,
        target: sm.target,
        sold: sm.sold,
      })
    );
  }

  for (const sp of supMap.values()) {
    if (!isSectorVisible(sp.sectorName)) continue;
    participants.push(
      toRow({
        participantKey: sp.participantKey,
        employeeCode: '',
        employeeName: sp.employeeName,
        roleLabel: sp.roleLabel,
        roleKind: 'supervisor',
        sectorName: sp.sectorName,
        managedSectors: [sp.sectorName],
        assignedCategories: [],
        isAutoTarget: false,
        target: sp.target,
        sold: sp.sold,
      })
    );
  }

  for (const rp of allAuthReps) {
    if (!isSectorVisible(rp.sectorName)) continue;
    participants.push(
      toRow({
        participantKey: rp.participantKey,
        employeeCode: rp.employeeCode,
        employeeName: rp.employeeName,
        roleLabel: 'مندوب مبيعات',
        roleKind: 'sales_rep',
        sectorName: rp.sectorName,
        managedSectors: [rp.sectorName],
        assignedCategories: [],
        isAutoTarget: false,
        target: rp.target,
        sold: rp.sold,
      })
    );
  }

  const roleOrder: Record<CompetitionParticipantRow['roleKind'], number> = {
    department_manager: 0,
    sector_manager: 1,
    supervisor: 2,
    sales_rep: 3,
  };

  participants.sort((a, b) => {
    const rDiff = roleOrder[a.roleKind] - roleOrder[b.roleKind];
    if (rDiff !== 0) return rDiff;
    const sDiff = b.soldStandardQty - a.soldStandardQty;
    if (sDiff !== 0) return sDiff;
    return a.employeeName.localeCompare(b.employeeName, 'ar');
  });

  // Non-overlapping summary across visible sectors
  let totalStandardSoldQty = 0;
  let totalIndustrialSoldQty = 0;
  let totalCommercialSoldQty = 0;
  let totalTargetQty = 0;
  let totalIndustrialTargetQty = 0;
  let totalCommercialTargetQty = 0;

  for (const sec of visibleSectors) {
    const ss = sectorSplitMap.get(sec);
    if (ss) {
      totalStandardSoldQty += ss.total;
      totalIndustrialSoldQty += ss.ind;
      totalCommercialSoldQty += ss.com;
    }
    const rt = sectorRollupTarget(sec);
    totalTargetQty += rt.total;
    totalIndustrialTargetQty += rt.ind;
    totalCommercialTargetQty += rt.com;
  }

  const roundedSoldTotal = roundQty(totalStandardSoldQty);
  const roundedIndSoldTotal = roundQty(totalIndustrialSoldQty);
  const roundedComSoldTotal = roundQty(totalCommercialSoldQty);
  const roundedTargetTotal = roundQty(totalTargetQty);
  const roundedIndTargetTotal = roundQty(totalIndustrialTargetQty);
  const roundedComTargetTotal = roundQty(totalCommercialTargetQty);

  return {
    competition: selectedComp,
    competitions,
    catalogItems,
    participants,
    summary: {
      totalSelectedItems: selectedComp.items.length,
      totalTeamMembers: participants.length,
      totalStandardSoldQty: roundedSoldTotal,
      totalTargetQty: roundedTargetTotal,
      overallAchievementPct: calculateQtyAchievementPct(roundedSoldTotal, roundedTargetTotal),
      totalIndustrialSoldQty: roundedIndSoldTotal,
      totalCommercialSoldQty: roundedComSoldTotal,
      totalIndustrialTargetQty: roundedIndTargetTotal,
      totalCommercialTargetQty: roundedComTargetTotal,
      industrialAchievementPct: calculateQtyAchievementPct(roundedIndSoldTotal, roundedIndTargetTotal),
      commercialAchievementPct: calculateQtyAchievementPct(roundedComSoldTotal, roundedComTargetTotal),
    },
  };
}
