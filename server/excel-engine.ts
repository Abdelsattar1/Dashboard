import type Database from 'better-sqlite3';
import * as XLSXModule from 'xlsx';
import type {
  ImportBatchHistoryItem,
  ImportPeriodSummary,
  ImportValidationReport,
  SheetCheckSummary,
  UserRole,
  ValidationIssue,
} from '../src/types/bi.ts';

const XLSX = (
  (XLSXModule as unknown as { default?: typeof XLSXModule }).default &&
  typeof (XLSXModule as unknown as { default?: typeof XLSXModule }).default?.read === 'function'
    ? (XLSXModule as unknown as { default: typeof XLSXModule }).default
    : XLSXModule
) as typeof XLSXModule;

export const REQUIRED_SHEETS = [
  'VF05N (2)',
  'Zsales_Price (2)',
  'القطاعات',
  'اهداف القطاعات',
  'تعريف المستخدمين',
  'تعريف الفروع و الاهداف',
  'تعريف الادارات والقطاعات',
  'تعريف مندوبى البيع',
  'تعريف الاصناف',
] as const;

export const ORIGINAL_SECTORS = [
  'التوريد المباشر',
  'التوريد المركزي',
  'التوب ريتيل',
  'الجيش',
] as const;

export function normalizeArabic(input: string): string {
  return String(input || '')
    .replace(/[\u00A0\u200B-\u200F\uFEFF]/g, ' ')
    .trim()
    .replace(/[\u064B-\u065F\u0670\u0640]/g, '')
    .replace(/[أإآ]/g, 'ا')
    .replace(/ى/g, 'ي')
    .replace(/ة/g, 'ه')
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

export function normalizeCode(raw: string): string {
  const cleaned = String(raw || '')
    .replace(/[\u00A0\u200B-\u200F\uFEFF]/g, '')
    .replace(/[٠-٩]/g, (d) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d)))
    .replace(/[۰-۹]/g, (d) => String('۰۱۲۳۴۵۶۷۸۹'.indexOf(d)))
    .trim();
  if (/^\d+\.0+$/.test(cleaned)) {
    return cleaned.replace(/\.0+$/, '');
  }
  return cleaned;
}

export function stripLeadingZeros(code: string): string {
  const norm = normalizeCode(code);
  if (!norm) return '';
  if (/^0+$/.test(norm)) return '0';
  return norm.replace(/^0+/, '');
}

export function canonicalizeSectorName(raw: string): string {
  const trimmed = String(raw || '').replace(/[\u00A0\u200B-\u200F\uFEFF]/g, ' ').trim();
  if (!trimmed) return '';
  const norm = normalizeArabic(trimmed);
  for (const s of ORIGINAL_SECTORS) {
    if (normalizeArabic(s) === norm) {
      return s;
    }
  }
  if (norm.includes('مباشر') && norm.includes('توريد')) return 'التوريد المباشر';
  if (norm.includes('مركزي') && norm.includes('توريد')) return 'التوريد المركزي';
  if ((norm.includes('توب') && norm.includes('ريتيل')) || (norm.includes('تجزئ') && norm.includes('مميز'))) return 'التوب ريتيل';
  if (norm === 'الجيش' || norm.includes('جيش')) return 'الجيش';
  return trimmed;
}

function normalizeHeader(h: string): string {
  return normalizeArabic(String(h || '').replace(/[\u00A0\u200B-\u200F\uFEFF]/g, ' ').replace(/[._:-]+/g, ' '));
}

function readCellAsText(cell: XLSXModule.CellObject | undefined): string {
  if (!cell || cell.v === undefined || cell.v === null) return '';
  const clean = (val: string) => val.replace(/[\u00A0\u200B-\u200F\uFEFF]/g, ' ').replace(/\s+/g, ' ').trim();
  if (cell.t === 's') return clean(String(cell.v));
  if (cell.t === 'n') {
    if (cell.w !== undefined && cell.w !== null) {
      const w = clean(String(cell.w));
      if (/^0\d+$/.test(w) || /^\d+$/.test(w)) {
        return w;
      }
    }
    return clean(String(cell.v));
  }
  if (cell.t === 'd' && cell.v instanceof Date) {
    return cell.v.toISOString().slice(0, 10);
  }
  return clean(String(cell.w ?? cell.v));
}

function readCellAsCode(cell: XLSXModule.CellObject | undefined): string {
  if (!cell || cell.v === undefined || cell.v === null) return '';
  if (cell.t === 's') return normalizeCode(String(cell.v));
  if (cell.t === 'n') {
    if (cell.w !== undefined && cell.w !== null) {
      const w = normalizeCode(String(cell.w));
      if (/^\d+$/.test(w)) {
        return w;
      }
    }
    return normalizeCode(String(Math.trunc(Number(cell.v)) === Number(cell.v) ? Number(cell.v) : cell.v));
  }
  return normalizeCode(String(cell.w ?? cell.v));
}

export function parseNumericValue(raw: unknown): number | null {
  if (raw === undefined || raw === null || raw === '') return null;
  if (typeof raw === 'number') {
    return Number.isFinite(raw) ? raw : null;
  }
  let s = String(raw).trim();
  if (!s) return null;

  // Convert Arabic-Indic digits to Western 0-9 if present and clean invisible/whitespace marks
  s = s
    .replace(/[٠-٩]/g, (d) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d)))
    .replace(/[\u00A0\u200E\u200F]/g, '')
    .trim();

  let negative = false;
  if (s.endsWith('-')) {
    negative = true;
    s = s.slice(0, -1).trim();
  } else if (s.startsWith('(') && s.endsWith(')')) {
    negative = true;
    s = s.slice(1, -1).trim();
  }

  s = s.replace(/[,\s]/g, '').replace(/(?:EGP|L\.?E\.?|SAR|USD|ج\.?م\.?|جنيه|\$|£|€)/gi, '');
  if (!/^[-+]?\d*(\.\d+)?$/.test(s) || s === '-' || s === '+' || s === '.') {
    return null;
  }
  const num = Number(s);
  if (!Number.isFinite(num)) return null;
  return negative ? -Math.abs(num) : num;
}

function readCellAsNumber(cell: XLSXModule.CellObject | undefined): number | null {
  if (!cell) return null;
  if (cell.t === 'n' && cell.v !== undefined && cell.v !== null) {
    return Number.isFinite(Number(cell.v)) ? Number(cell.v) : null;
  }
  const raw = cell.v !== undefined && cell.v !== null && cell.v !== '' ? cell.v : cell.w;
  return parseNumericValue(raw);
}

export function excelSerialToIsoDate(serial: number): string {
  if (!Number.isFinite(serial) || serial < 1 || serial > 2958465) {
    return '';
  }
  const wholeDays = Math.floor(serial);
  // Excel's 1900 date system treats 1900 as a leap year (serial 60 = 1900-02-29)
  const utcDays = wholeDays >= 60 ? wholeDays - 25569 : wholeDays - 25568;
  const dt = new Date(utcDays * 86400000);
  if (Number.isNaN(dt.getTime())) return '';
  return dt.toISOString().slice(0, 10);
}

function readCellAsDate(cell: XLSXModule.CellObject | undefined): string {
  if (!cell || cell.v === undefined || cell.v === null) return '';
  if (cell.t === 'n' && typeof cell.v === 'number') {
    const parseDateCode = XLSX?.SSF?.parse_date_code;
    if (typeof parseDateCode === 'function') {
      const parsed = parseDateCode(cell.v);
      if (parsed && parsed.y && parsed.m && parsed.d) {
        const mm = String(parsed.m).padStart(2, '0');
        const dd = String(parsed.d).padStart(2, '0');
        return `${parsed.y}-${mm}-${dd}`;
      }
    }
    const iso = excelSerialToIsoDate(cell.v);
    if (iso) return iso;
  }
  if (cell.t === 'd' && cell.v instanceof Date && !Number.isNaN(cell.v.getTime())) {
    return cell.v.toISOString().slice(0, 10);
  }
  const str = String(cell.w ?? cell.v)
    .replace(/[\u00A0\u200B-\u200F\uFEFF]/g, ' ')
    .replace(/[٠-٩]/g, (d) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d)))
    .replace(/[۰-۹]/g, (d) => String('۰۱۲۳۴۵۶۷۸۹'.indexOf(d)))
    .trim();
  if (!str) return '';

  // Match DD.MM.YYYY or DD/MM/YYYY (or MM/DD/YYYY if second part > 12)
  const dmy = str.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/);
  if (dmy) {
    const p1 = Number(dmy[1]);
    const p2 = Number(dmy[2]);
    const yearRaw = dmy[3];
    const yyyy = yearRaw.length === 2 ? `20${yearRaw}` : yearRaw;
    if (p2 > 12 && p1 <= 12) {
      return `${yyyy}-${String(p1).padStart(2, '0')}-${String(p2).padStart(2, '0')}`;
    }
    return `${yyyy}-${String(p2).padStart(2, '0')}-${String(p1).padStart(2, '0')}`;
  }

  // Match YYYY-MM-DD or YYYY/MM/DD
  const ymd = str.match(/^(\d{4})[./-](\d{1,2})[./-](\d{1,2})/);
  if (ymd) {
    const yyyy = ymd[1];
    const mm = ymd[2].padStart(2, '0');
    const dd = ymd[3].padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  }

  return str;
}

const ARABIC_MONTH_MAP: Array<[string, number]> = [
  ['يناير', 1],
  ['فبراير', 2],
  ['مارس', 3],
  ['ابريل', 4],
  ['مايو', 5],
  ['يونيو', 6],
  ['يونيه', 6],
  ['يوليو', 7],
  ['يوليه', 7],
  ['اغسطس', 8],
  ['سبتمبر', 9],
  ['اكتوبر', 10],
  ['نوفمبر', 11],
  ['ديسمبر', 12],
  ['january', 1],
  ['jan', 1],
  ['february', 2],
  ['feb', 2],
  ['march', 3],
  ['mar', 3],
  ['april', 4],
  ['apr', 4],
  ['may', 5],
  ['june', 6],
  ['jun', 6],
  ['july', 7],
  ['jul', 7],
  ['august', 8],
  ['aug', 8],
  ['september', 9],
  ['sept', 9],
  ['sep', 9],
  ['october', 10],
  ['oct', 10],
  ['november', 11],
  ['nov', 11],
  ['december', 12],
  ['dec', 12],
];

function isValidYearMonth(year: number, month: number): boolean {
  return Number.isInteger(year) && year >= 1990 && year <= 2099 && Number.isInteger(month) && month >= 1 && month <= 12;
}

export function parseSalesPeriodFromCell(cell: XLSXModule.CellObject | undefined): string {
  if (!cell || cell.v === undefined || cell.v === null) return '';

  if (cell.t === 'n' && typeof cell.v === 'number' && Number.isFinite(cell.v)) {
    const intVal = Math.trunc(cell.v);
    if (intVal === cell.v && intVal >= 199001 && intVal <= 209912) {
      const y = Math.floor(intVal / 100);
      const m = intVal % 100;
      if (isValidYearMonth(y, m)) {
        return `${y}-${String(m).padStart(2, '0')}`;
      }
    }
    const iso = readCellAsDate(cell);
    if (/^\d{4}-\d{2}/.test(iso)) {
      return iso.slice(0, 7);
    }
  }

  if (cell.t === 'd' && cell.v instanceof Date && !Number.isNaN(cell.v.getTime())) {
    return cell.v.toISOString().slice(0, 7);
  }

  const rawText = String(cell.w ?? cell.v)
    .replace(/[\u00A0\u200B-\u200F\uFEFF]/g, ' ')
    .replace(/[٠-٩]/g, (d) => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d)))
    .replace(/[۰-۹]/g, (d) => String('۰۱۲۳۴۵۶۷۸۹'.indexOf(d)))
    .trim();
  if (!rawText) return '';

  // YYYY-MM or YYYY/MM or YYYY-MM-DD
  const ym = rawText.match(/^(\d{4})[./-](\d{1,2})(?:[./-]\d{1,2})?/);
  if (ym) {
    const y = Number(ym[1]);
    const m = Number(ym[2]);
    if (isValidYearMonth(y, m)) {
      return `${y}-${String(m).padStart(2, '0')}`;
    }
  }

  // MM/YYYY or MM-YYYY
  const my = rawText.match(/^(\d{1,2})[./-](\d{4})$/);
  if (my) {
    const m = Number(my[1]);
    const y = Number(my[2]);
    if (isValidYearMonth(y, m)) {
      return `${y}-${String(m).padStart(2, '0')}`;
    }
  }

  // DD/MM/YYYY or DD.MM.YYYY
  const dmy = rawText.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/);
  if (dmy) {
    const p1 = Number(dmy[1]);
    const p2 = Number(dmy[2]);
    const yearRaw = dmy[3];
    const y = Number(yearRaw.length === 2 ? `20${yearRaw}` : yearRaw);
    const m = p2 > 12 && p1 <= 12 ? p1 : p2;
    if (isValidYearMonth(y, m)) {
      return `${y}-${String(m).padStart(2, '0')}`;
    }
  }

  // YYYYMM compact text
  const compact = rawText.match(/^(20\d{2}|19\d{2})(\d{2})$/);
  if (compact) {
    const y = Number(compact[1]);
    const m = Number(compact[2]);
    if (isValidYearMonth(y, m)) {
      return `${y}-${String(m).padStart(2, '0')}`;
    }
  }

  // Named Arabic/English month + year (e.g. "نوفمبر 2026", "Nov-2026", "2026 نوفمبر")
  const norm = normalizeArabic(rawText);
  const yearMatch = norm.match(/\b(19\d{2}|20\d{2})\b/);
  if (yearMatch) {
    const y = Number(yearMatch[1]);
    for (const [label, m] of ARABIC_MONTH_MAP) {
      if (norm.includes(label) && isValidYearMonth(y, m)) {
        return `${y}-${String(m).padStart(2, '0')}`;
      }
    }
  }

  return '';
}

export function extractMonthKey(billingDate: string | null | undefined): string {
  const trimmed = String(billingDate || '').trim();
  const m = trimmed.match(/^(\d{4})-(\d{2})(?:-\d{2})?/);
  if (m && isValidYearMonth(Number(m[1]), Number(m[2]))) {
    return `${m[1]}-${m[2]}`;
  }
  return 'UNSPECIFIED';
}

export function extractBillingMonthAndDate(
  billingDate: string,
  rawDateCell?: XLSXModule.CellObject,
  rawPeriodCell?: XLSXModule.CellObject
): { monthKey: string; billingDate: string } {
  let fromDate = '';
  const trimmedDate = String(billingDate || '').trim();
  const isoMatch = trimmedDate.match(/^(\d{4})-(\d{2})(?:-\d{2})?$/);
  if (isoMatch && isValidYearMonth(Number(isoMatch[1]), Number(isoMatch[2]))) {
    fromDate = `${isoMatch[1]}-${isoMatch[2]}`;
  } else if (rawDateCell) {
    fromDate = parseSalesPeriodFromCell(rawDateCell);
  }

  const fromPeriod = rawPeriodCell ? parseSalesPeriodFromCell(rawPeriodCell) : '';
  const monthKey = fromDate || fromPeriod || 'UNSPECIFIED';

  let normalizedDate = trimmedDate;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(normalizedDate) && monthKey !== 'UNSPECIFIED') {
    normalizedDate = `${monthKey}-01`;
  }

  return { monthKey, billingDate: normalizedDate };
}

export function getActiveSalesPeriods(db: Database.Database): string[] {
  const active = db
    .prepare(`SELECT id FROM import_batches WHERE status = 'active' ORDER BY id DESC LIMIT 1`)
    .get() as { id: number } | undefined;
  if (!active) return [];
  const rows = db
    .prepare(`SELECT DISTINCT billing_date FROM sales_records WHERE batch_id = ?`)
    .all(active.id) as Array<{ billing_date: string | null }>;
  const periods = new Set<string>();
  for (const r of rows) {
    periods.add(extractMonthKey(r.billing_date));
  }
  return Array.from(periods).sort();
}

interface ParsedSheetTable {
  headers: string[];
  headerRowIndex: number;
  rows: Array<Record<string, XLSXModule.CellObject | undefined>>;
}

function extractSheetTable(ws: XLSXModule.WorkSheet, hintKeywords: string[]): ParsedSheetTable {
  const ref = ws['!ref'];
  if (!ref) {
    return { headers: [], headerRowIndex: 0, rows: [] };
  }
  const range = XLSX.utils.decode_range(ref);
  const maxScanRow = Math.min(range.e.r, range.s.r + 12);

  let bestHeaderRow = range.s.r;
  let bestScore = -1;

  for (let r = range.s.r; r <= maxScanRow; r++) {
    let score = 0;
    let nonEmpty = 0;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr] as XLSXModule.CellObject | undefined;
      const val = readCellAsText(cell);
      if (val) {
        nonEmpty++;
        const norm = normalizeHeader(val);
        if (hintKeywords.some((kw) => norm.includes(normalizeHeader(kw)))) {
          score += 3;
        }
      }
    }
    const totalScore = (nonEmpty >= 2 ? 100 : 0) + score * 10 + nonEmpty;
    if (nonEmpty > 0 && score > 0 && totalScore > bestScore) {
      bestScore = totalScore;
      bestHeaderRow = r;
    } else if (bestScore === -1 && nonEmpty > 0) {
      bestHeaderRow = r;
      bestScore = nonEmpty;
    }
  }

  const headers: string[] = [];
  const colIndices: number[] = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r: bestHeaderRow, c });
    const headerText = readCellAsText(ws[addr] as XLSXModule.CellObject | undefined);
    if (headerText) {
      headers.push(headerText);
      colIndices.push(c);
    }
  }

  const rows: Array<Record<string, XLSXModule.CellObject | undefined>> = [];
  for (let r = bestHeaderRow + 1; r <= range.e.r; r++) {
    const rowObj: Record<string, XLSXModule.CellObject | undefined> = {};
    let hasAnyValue = false;
    for (let i = 0; i < headers.length; i++) {
      const c = colIndices[i];
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr] as XLSXModule.CellObject | undefined;
      rowObj[headers[i]] = cell;
      if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '') {
        hasAnyValue = true;
      }
    }
    if (hasAnyValue) {
      rows.push(rowObj);
    }
  }

  return { headers, headerRowIndex: bestHeaderRow, rows };
}

function findColumnHeader(headers: string[], aliases: string[]): string | null {
  const normMap = new Map<string, string>();
  const noSpaceMap = new Map<string, string>();
  for (const h of headers) {
    const nh = normalizeHeader(h);
    normMap.set(nh, h);
    noSpaceMap.set(nh.replace(/\s+/g, ''), h);
  }
  for (const alias of aliases) {
    const na = normalizeHeader(alias);
    const match = normMap.get(na) ?? noSpaceMap.get(na.replace(/\s+/g, ''));
    if (match) return match;
  }
  // Partial fallback only for multi-word unambiguous aliases
  for (const alias of aliases) {
    const normAlias = normalizeHeader(alias);
    if (normAlias.length < 4) continue;
    for (const [normH, origH] of normMap.entries()) {
      if (normH.includes(normAlias)) return origH;
    }
  }
  return null;
}

function findSheetByName(wb: XLSXModule.WorkBook, targetName: string): { actualName: string; ws: XLSXModule.WorkSheet } | null {
  if (wb.Sheets[targetName]) {
    return { actualName: targetName, ws: wb.Sheets[targetName] };
  }
  const normTarget = normalizeArabic(targetName);
  for (const name of wb.SheetNames) {
    if (normalizeArabic(name) === normTarget) {
      return { actualName: name, ws: wb.Sheets[name] };
    }
  }
  return null;
}

interface ParsedSalesRow {
  sourceType: 'VF05N' | 'ZSALES';
  docNumber: string;
  billingDate: string;
  monthKey: string;
  payerCode: string;
  payerName: string;
  soldToCode: string;
  soldToName: string;
  itemCode: string;
  itemName: string;
  repCode: string;
  quantity: number;
  salesValue: number;
  netSalesValue: number;
  taxValue: number;
  resolvedSector: string | null;
  resolvedDepartment: string | null;
  resolvedChainName: string | null;
  resolvedBranchCode: string | null;
  resolvedBranchName: string | null;
  resolvedRepCode: string | null;
  resolvedRepName: string | null;
  resolvedItemName: string | null;
  isReturn: number;
}

interface StagedImportPayload {
  validationToken: string;
  filename: string;
  createdAtMs: number;
  report: ImportValidationReport;
  sectors: Array<{ sectorCode: string; sectorName: string; monthlyTarget: number; isOriginal: number }>;
  chains: Array<{ chainCode: string; chainName: string; departmentName: string; sectorName: string; supervisorName: string }>;
  branches: Array<{
    branchCode: string;
    branchName: string;
    sectorName: string;
    branchTarget: number;
    chainCode: string;
    chainName: string;
    repCode: string;
    repName: string;
    supervisorName: string;
  }>;
  reps: Array<{ repCode: string; repName: string; sectorName: string; supervisorName: string; monthlyTarget: number }>;
  items: Array<{
    itemCode: string;
    itemName: string;
    categoryName: string;
    brandName: string;
    unitName: string;
    chainsCategory: string;
    armyCategory: string;
    topRetailCategory: string;
  }>;
  categoryTargets: Array<{
    sectorName: string;
    jobTitle: string;
    employeeCode: string;
    employeeName: string;
    categoryName: string;
    targetValue: number;
  }>;
  hasCategoryTargetsSheet: boolean;
  salesRecords: ParsedSalesRow[];
  importedUsers: Array<{ username: string; fullName: string; email: string | null; role: UserRole; allowedSectors: string[] }>;
}

const stagedImports = new Map<string, StagedImportPayload>();

function pruneExpiredStages() {
  const now = Date.now();
  for (const [k, v] of stagedImports.entries()) {
    if (now - v.createdAtMs > 1000 * 60 * 30) {
      stagedImports.delete(k);
    }
  }
}

export function validateExcelBuffer(
  buffer: Uint8Array | ArrayBuffer,
  filename: string,
  existingPeriods: string[] = []
): ImportValidationReport {
  pruneExpiredStages();

  const errors: ValidationIssue[] = [];
  const warnings: ValidationIssue[] = [];
  const sheetsSummary: SheetCheckSummary[] = [];

  let wb: XLSXModule.WorkBook;
  try {
    const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    wb = XLSX.read(bytes, { type: 'array', cellDates: false });
  } catch (err) {
    const token = globalThis.crypto.randomUUID();
    const msg = err instanceof Error ? err.message : 'ملف Excel غير صالح';
    errors.push({
      severity: 'error',
      sheet: 'المصنف',
      code: 'INVALID_WORKBOOK',
      message: `تعذر قراءة ملف Excel: ${msg}`,
    });
    return {
      validationToken: token,
      filename,
      valid: false,
      canCommit: false,
      errors,
      warnings,
      sheets: [],
      comparison: {
        vf05n: { count: 0, salesTotal: 0, netSalesTotal: 0, taxTotal: 0, returnCount: 0 },
        zsales: { count: 0, salesTotal: 0, netSalesTotal: 0, taxTotal: 0, returnCount: 0 },
        salesDifference: 0,
        netSalesDifference: 0,
        taxDifference: 0,
        isMatched: false,
      },
      periodsDetected: [],
      linkStats: {
        vf05nTotalRows: 0,
        vf05nSectorLinkedRows: 0,
        vf05nSectorUnlinkedRows: 0,
        vf05nBranchLinkedRows: 0,
        vf05nBranchUnlinkedRows: 0,
        unlinkedBranchCodes: [],
        unlinkedSectorCodes: [],
      },
      targetsSummary: {
        sectors: [],
        totalSectorTargets: 0,
        directBranchesTargetTotal: 0,
        directBranchCount: 0,
      },
      usersFound: [],
    };
  }

  // 1. Check presence of all 9 required sheets
  const sheetTables = new Map<string, ParsedSheetTable>();
  const sheetHints: Record<string, string[]> = {
    'VF05N (2)': ['New Growth Value BD&T', 'New Net Value BT', 'New Total Tax', 'Payer', 'Sold-to Party', 'Billing Date', 'Material', 'الفترة', 'الشهر', 'Period', 'Month'],
    'Zsales_Price (2)': ['New Growth Value BD&T', 'Net Value Without TA', 'Total Tax', 'Customer', 'Sales Emp.', 'Material', 'Billing Dt.', 'Billing Date', 'الفترة', 'الشهر', 'Period', 'Month'],
    'القطاعات': ['القطاع', 'اسم القطاع', 'Sector'],
    'اهداف القطاعات': ['القطاع', 'الهدف', 'المستهدف', 'Target'],
    'تعريف المستخدمين': ['اسم المستخدم', 'المستخدم', 'الصلاحية', 'القطاع', 'القطاعات'],
    'تعريف الفروع و الاهداف': ['كود الفرع', 'اسم الفرع', 'القطاع', 'كود الادارة', 'اسم الادارة', 'كود المندوب', 'اسم المندوب', 'اسم المشرف', 'إجمالى الهدف', 'الهدف', 'الفرع'],
    'تعريف الادارات والقطاعات': ['كود السلسلة', 'السلسلة', 'اسم السلسلة', 'الادارة', 'القطاع', 'اسم المشرف'],
    'تعريف مندوبى البيع': ['كود المندوب', 'اسم المندوب', 'المندوب'],
    'تعريف الاصناف': ['كود الصنف', 'اسم الصنف', 'Material', 'الصنف'],
  };

  for (const reqSheet of REQUIRED_SHEETS) {
    const found = findSheetByName(wb, reqSheet);
    if (!found) {
      errors.push({
        severity: 'error',
        sheet: reqSheet,
        code: 'MISSING_SHEET',
        message: `ورقة العمل المطلوبة "${reqSheet}" غير موجودة في الملف.`,
      });
      sheetsSummary.push({
        sheetName: reqSheet,
        found: false,
        rowCount: 0,
        detectedColumns: [],
        missingColumns: [],
      });
    } else {
      if (found.actualName !== reqSheet) {
        warnings.push({
          severity: 'warning',
          sheet: reqSheet,
          code: 'SHEET_NAME_WHITESPACE',
          message: `تم العثور على الورقة باسم "${found.actualName}" ومطابقتها مع "${reqSheet}".`,
        });
      }
      const table = extractSheetTable(found.ws, sheetHints[reqSheet] ?? []);
      sheetTables.set(reqSheet, table);
      sheetsSummary.push({
        sheetName: reqSheet,
        found: true,
        rowCount: table.rows.length,
        detectedColumns: table.headers,
        missingColumns: [],
      });
    }
  }

  // Helper to update missingColumns on sheetsSummary
  const markMissingCols = (sheetName: string, missing: string[]) => {
    const item = sheetsSummary.find((s) => s.sheetName === sheetName);
    if (item) {
      item.missingColumns.push(...missing);
    }
  };

  // 2. Parse reference tables first (Sectors, Sector Targets, Chains/Departments, Branches, Reps, Items, Users)
  const sectorTargetMap = new Map<string, number>();
  const sectorsList: Array<{ sectorCode: string; sectorName: string; monthlyTarget: number; isOriginal: number }> = [];

  const sectorsTable = sheetTables.get('القطاعات');
  if (sectorsTable) {
    const colSector = findColumnHeader(sectorsTable.headers, ['القطاع', 'اسم القطاع', 'Sector', 'Sector Name']);
    const colCode = findColumnHeader(sectorsTable.headers, ['كود القطاع', 'Sector Code', 'الكود']);
    if (!colSector) {
      markMissingCols('القطاعات', ['القطاع']);
      errors.push({
        severity: 'error',
        sheet: 'القطاعات',
        code: 'MISSING_COLUMN',
        message: 'عمود "القطاع" مفقود في ورقة القطاعات.',
      });
    } else {
      const seen = new Set<string>();
      for (const row of sectorsTable.rows) {
        const name = canonicalizeSectorName(readCellAsText(row[colSector]));
        if (!name) continue;
        if (seen.has(name)) {
          warnings.push({
            severity: 'warning',
            sheet: 'القطاعات',
            code: 'DUPLICATE_SECTOR',
            message: `تكرار اسم القطاع "${name}" في ورقة القطاعات.`,
          });
          continue;
        }
        seen.add(name);
        const code = colCode ? readCellAsCode(row[colCode]) || name : name;
        const isOrig = ORIGINAL_SECTORS.includes(name as (typeof ORIGINAL_SECTORS)[number]) ? 1 : 0;
        sectorsList.push({ sectorCode: code, sectorName: name, monthlyTarget: 0, isOriginal: isOrig });
      }
    }
  }

  const sectorTargetsTable = sheetTables.get('اهداف القطاعات');
  if (sectorTargetsTable) {
    const colSector = findColumnHeader(sectorTargetsTable.headers, ['القطاع', 'اسم القطاع', 'Sector']);
    const colTarget = findColumnHeader(sectorTargetsTable.headers, ['الهدف', 'المستهدف', 'الهدف الشهري', 'Target', 'Monthly Target']);
    const missing: string[] = [];
    if (!colSector) missing.push('القطاع');
    if (!colTarget) missing.push('الهدف');
    if (missing.length > 0) {
      markMissingCols('اهداف القطاعات', missing);
      errors.push({
        severity: 'error',
        sheet: 'اهداف القطاعات',
        code: 'MISSING_COLUMN',
        message: `أعمدة مطلوبة مفقودة في ورقة اهداف القطاعات: ${missing.join('، ')}`,
      });
    } else if (colSector && colTarget) {
      for (const row of sectorTargetsTable.rows) {
        const sName = canonicalizeSectorName(readCellAsText(row[colSector]));
        if (!sName) continue;
        const tVal = readCellAsNumber(row[colTarget]);
        if (tVal === null) {
          errors.push({
            severity: 'error',
            sheet: 'اهداف القطاعات',
            code: 'INVALID_NUMBER',
            message: `قيمة هدف غير رقمية للقطاع "${sName}".`,
          });
        } else {
          sectorTargetMap.set(sName, tVal);
        }
      }
    }
  }

  // Merge targets into sectorsList (ensure all 4 original sectors exist if present in targets)
  for (const [sName, target] of sectorTargetMap.entries()) {
    const existing = sectorsList.find((s) => s.sectorName === sName);
    if (existing) {
      existing.monthlyTarget = target;
    } else {
      const isOrig = ORIGINAL_SECTORS.includes(sName as (typeof ORIGINAL_SECTORS)[number]) ? 1 : 0;
      sectorsList.push({ sectorCode: sName, sectorName: sName, monthlyTarget: target, isOriginal: isOrig });
    }
  }

  // Parse Chains & Departments (`تعريف الادارات والقطاعات`)
  type ChainEntry = {
    chainCode: string;
    chainName: string;
    departmentName: string;
    sectorName: string;
    supervisorName: string;
  };
  const chainMap = new Map<string, ChainEntry>();
  const chainTrimmedMap = new Map<string, ChainEntry>();
  const chainByNameMap = new Map<string, ChainEntry>();
  const chainsList: ChainEntry[] = [];
  const chainsTable = sheetTables.get('تعريف الادارات والقطاعات');
  if (chainsTable) {
    const colCode = findColumnHeader(chainsTable.headers, ['كود السلسلة', 'كود العميل', 'Payer', 'Chain Code']);
    const rawColChain = findColumnHeader(chainsTable.headers, ['السلسلة', 'اسم السلسلة', 'Chain Name', 'Name of Payer']);
    const colChain = rawColChain && rawColChain !== colCode ? rawColChain : null;
    const colDept = findColumnHeader(chainsTable.headers, ['الادارة', 'الإدارة', 'Department', 'القناة']);
    const colSector = findColumnHeader(chainsTable.headers, ['القطاع', 'اسم القطاع', 'Sector']);
    const colSup = findColumnHeader(chainsTable.headers, ['اسم المشرف', 'المشرف', 'مدير المنطقة', 'Supervisor', 'مشرف']);

    const missing: string[] = [];
    if (!colCode) missing.push('كود السلسلة');
    if (!colSector) missing.push('القطاع');
    if (missing.length > 0) {
      markMissingCols('تعريف الادارات والقطاعات', missing);
      errors.push({
        severity: 'error',
        sheet: 'تعريف الادارات والقطاعات',
        code: 'MISSING_COLUMN',
        message: `أعمدة مطلوبة مفقودة في ورقة تعريف الادارات والقطاعات: ${missing.join('، ')}`,
      });
    } else if (colCode && colSector) {
      const dupes: string[] = [];
      for (const row of chainsTable.rows) {
        const code = readCellAsCode(row[colCode]);
        const sector = canonicalizeSectorName(readCellAsText(row[colSector]));
        const chainName = colChain ? readCellAsText(row[colChain]) : code;
        const deptName = colDept ? readCellAsText(row[colDept]) : '';
        const supervisorName = colSup ? readCellAsText(row[colSup]) : '';
        if (!code) {
          warnings.push({
            severity: 'warning',
            sheet: 'تعريف الادارات والقطاعات',
            code: 'MISSING_CHAIN_CODE',
            message: 'يوجد صف بدون كود سلسلة في ورقة تعريف الادارات والقطاعات.',
          });
          continue;
        }
        if (chainMap.has(code)) {
          dupes.push(code);
          continue;
        }
        const entry: ChainEntry = {
          chainCode: code,
          chainName: chainName || code,
          departmentName: deptName || supervisorName || sector,
          sectorName: sector,
          supervisorName,
        };
        chainMap.set(code, entry);
        const trimmedCode = stripLeadingZeros(code);
        if (trimmedCode && !chainTrimmedMap.has(trimmedCode)) {
          chainTrimmedMap.set(trimmedCode, entry);
        }
        const normName = normalizeArabic(entry.chainName);
        if (normName && !chainByNameMap.has(normName)) {
          chainByNameMap.set(normName, entry);
        }
        chainsList.push(entry);
      }
      if (dupes.length > 0) {
        warnings.push({
          severity: 'warning',
          sheet: 'تعريف الادارات والقطاعات',
          code: 'DUPLICATE_CHAIN_CODE',
          message: `تم اكتشاف ${dupes.length} أكواد سلاسل مكررة في تعريف الادارات والقطاعات (تم اعتماد أول سجل لكل كود).`,
          details: dupes.slice(0, 15),
        });
      }
    }
  }

  // Parse Branches & Targets (`تعريف الفروع و الاهداف`)
  const branchMap = new Map<
    string,
    {
      branchCode: string;
      branchName: string;
      sectorName: string;
      branchTarget: number;
      chainCode: string;
      chainName: string;
      repCode: string;
      repName: string;
      supervisorName: string;
    }
  >();
  const branchTrimmedMap = new Map<
    string,
    {
      branchCode: string;
      branchName: string;
      sectorName: string;
      branchTarget: number;
      chainCode: string;
      chainName: string;
      repCode: string;
      repName: string;
      supervisorName: string;
    }
  >();
  const branchesList: Array<{
    branchCode: string;
    branchName: string;
    sectorName: string;
    branchTarget: number;
    chainCode: string;
    chainName: string;
    repCode: string;
    repName: string;
    supervisorName: string;
  }> = [];
  const branchesTable = sheetTables.get('تعريف الفروع و الاهداف');
  if (branchesTable) {
    const colCode = findColumnHeader(branchesTable.headers, ['كود الفرع', 'Branch Code', 'Sold-to Party', 'Sold-To']);
    const colName = findColumnHeader(branchesTable.headers, ['اسم الفرع', 'الفرع', 'Branch Name']);
    const colTarget = findColumnHeader(branchesTable.headers, ['إجمالى الهدف', 'اجمالي الهدف', 'الهدف', 'المستهدف', 'هدف الفرع', 'Target']);
    const colSector = findColumnHeader(branchesTable.headers, ['القطاع', 'Sector']);
    const colChain = findColumnHeader(branchesTable.headers, ['كود السلسلة', 'كود الادارة', 'كود الإدارة', 'Payer', 'Chain Code']);
    const colChainName = findColumnHeader(branchesTable.headers, ['اسم السلسلة', 'اسم الادارة', 'اسم الإدارة', 'السلسلة', 'Chain Name']);
    const colRep = findColumnHeader(branchesTable.headers, ['كود المندوب', 'Rep Code']);
    const colRepName = findColumnHeader(branchesTable.headers, ['اسم المندوب', 'Rep Name']);
    const colSup = findColumnHeader(branchesTable.headers, ['اسم المشرف', 'المشرف', 'مدير المنطقة', 'Supervisor']);

    const missing: string[] = [];
    if (!colCode) missing.push('كود الفرع');
    if (!colName) missing.push('اسم الفرع');
    if (missing.length > 0) {
      markMissingCols('تعريف الفروع و الاهداف', missing);
      errors.push({
        severity: 'error',
        sheet: 'تعريف الفروع و الاهداف',
        code: 'MISSING_COLUMN',
        message: `أعمدة مطلوبة مفقودة في ورقة تعريف الفروع و الاهداف: ${missing.join('، ')}`,
      });
    } else if (colCode && colName) {
      const dupes: string[] = [];
      for (const row of branchesTable.rows) {
        const code = readCellAsCode(row[colCode]);
        const name = readCellAsText(row[colName]);
        if (!code) continue;
        if (branchMap.has(code)) {
          dupes.push(code);
          continue;
        }
        const target = colTarget ? (readCellAsNumber(row[colTarget]) ?? 0) : 0;
        const sector = colSector ? canonicalizeSectorName(readCellAsText(row[colSector])) || 'التوريد المباشر' : 'التوريد المباشر';
        const rawChainCode = colChain ? readCellAsCode(row[colChain]) : '';
        const rawChainName = colChainName ? readCellAsText(row[colChainName]) : '';
        const chainDef =
          (rawChainCode ? chainMap.get(rawChainCode) : undefined) ??
          (rawChainCode ? chainTrimmedMap.get(stripLeadingZeros(rawChainCode)) : undefined) ??
          (rawChainName ? chainByNameMap.get(normalizeArabic(rawChainName)) : undefined);
        const resolvedChainCode = rawChainCode || chainDef?.chainCode || '';
        const resolvedChainName = rawChainName || (!colChainName ? chainDef?.chainName || '' : '');
        // Supervisor MUST come strictly from `تعريف الفروع و الاهداف` (`اسم المشرف`) — never inferred from chain or rep
        const resolvedSupervisor = colSup ? readCellAsText(row[colSup]) : '';

        const entry = {
          branchCode: code,
          branchName: name || code,
          sectorName: sector,
          branchTarget: target,
          chainCode: resolvedChainCode,
          chainName: resolvedChainName,
          repCode: colRep ? readCellAsCode(row[colRep]) : '',
          repName: colRepName ? readCellAsText(row[colRepName]) : '',
          supervisorName: resolvedSupervisor,
        };
        branchMap.set(code, entry);
        const trimmedBranchCode = stripLeadingZeros(code);
        if (trimmedBranchCode && !branchTrimmedMap.has(trimmedBranchCode)) {
          branchTrimmedMap.set(trimmedBranchCode, entry);
        }
        branchesList.push(entry);
      }
      if (dupes.length > 0) {
        warnings.push({
          severity: 'warning',
          sheet: 'تعريف الفروع و الاهداف',
          code: 'DUPLICATE_BRANCH_CODE',
          message: `تم اكتشاف ${dupes.length} أكواد فروع مكررة في تعريف الفروع و الاهداف.`,
          details: dupes.slice(0, 15),
        });
      }
    }
  }

  // Parse Sales Reps (`تعريف مندوبى البيع`)
  const repMap = new Map<string, { repCode: string; repName: string; sectorName: string; supervisorName: string; monthlyTarget: number }>();
  const repsList: Array<{ repCode: string; repName: string; sectorName: string; supervisorName: string; monthlyTarget: number }> = [];
  const repsTable = sheetTables.get('تعريف مندوبى البيع');
  if (repsTable) {
    const colCode = findColumnHeader(repsTable.headers, ['كود المندوب', 'Rep Code', 'Sales Employee', 'Sales Rep']);
    const colName = findColumnHeader(repsTable.headers, ['اسم المندوب', 'المندوب', 'Rep Name']);
    const colSector = findColumnHeader(repsTable.headers, ['القطاع', 'Sector']);
    const colSup = findColumnHeader(repsTable.headers, ['اسم المشرف', 'المشرف', 'مدير المنطقة', 'Supervisor', 'مشرف']);
    const colTarget = findColumnHeader(repsTable.headers, ['إجمالى الهدف', 'اجمالي الهدف', 'هدف المندوب', 'الهدف', 'المستهدف', 'Target']);

    const missing: string[] = [];
    if (!colCode) missing.push('كود المندوب');
    if (!colName) missing.push('اسم المندوب');
    if (missing.length > 0) {
      markMissingCols('تعريف مندوبى البيع', missing);
      errors.push({
        severity: 'error',
        sheet: 'تعريف مندوبى البيع',
        code: 'MISSING_COLUMN',
        message: `أعمدة مطلوبة مفقودة في ورقة تعريف مندوبى البيع: ${missing.join('، ')}`,
      });
    } else if (colCode && colName) {
      const dupes: string[] = [];
      for (const row of repsTable.rows) {
        const code = readCellAsCode(row[colCode]);
        const name = readCellAsText(row[colName]);
        if (!code) continue;
        if (repMap.has(code)) {
          dupes.push(code);
          continue;
        }
        const entry = {
          repCode: code,
          repName: name || code,
          sectorName: colSector ? canonicalizeSectorName(readCellAsText(row[colSector])) : '',
          supervisorName: colSup ? readCellAsText(row[colSup]) : '',
          monthlyTarget: colTarget ? (readCellAsNumber(row[colTarget]) ?? 0) : 0,
        };
        repMap.set(code, entry);
        repsList.push(entry);
      }
      if (dupes.length > 0) {
        warnings.push({
          severity: 'warning',
          sheet: 'تعريف مندوبى البيع',
          code: 'DUPLICATE_REP_CODE',
          message: `تكرار ${dupes.length} أكواد مناديب في تعريف مندوبى البيع.`,
          details: dupes.slice(0, 15),
        });
      }
    }
  }

  // Enrich branch repName from parsed repMap ONLY if `اسم المندوب` was blank in `تعريف الفروع و الاهداف` (never override supervisor!)
  for (const b of branchesList) {
    if (b.repCode && !b.repName) {
      const repDef = repMap.get(b.repCode);
      if (repDef) {
        b.repName = repDef.repName;
      }
    }
  }

  // Parse Items (`تعريف الاصناف`)
  interface ParsedItemDef {
    itemCode: string;
    itemName: string;
    categoryName: string;
    brandName: string;
    unitName: string;
    chainsCategory: string;
    armyCategory: string;
    topRetailCategory: string;
  }
  const itemMap = new Map<string, ParsedItemDef>();
  const itemsList: ParsedItemDef[] = [];
  const itemsTable = sheetTables.get('تعريف الاصناف');
  if (itemsTable) {
    const colCode = findColumnHeader(itemsTable.headers, ['كود الصنف', 'Material', 'Item Code', 'SKU']);
    const colName = findColumnHeader(itemsTable.headers, ['اسم الصنف', 'الصنف', 'Material Description', 'Item Name']);
    const colClassA = findColumnHeader(itemsTable.headers, [
      'التصنيف A',
      'التصنيف أ',
      'تصنيف A',
      'التصنيفA',
      'Classification A',
      'Class A',
    ]);
    const colCat =
      colClassA ?? findColumnHeader(itemsTable.headers, ['التصنيف', 'المجموعة', 'الفئة', 'Category']);
    const colChainsCat = findColumnHeader(itemsTable.headers, [
      'تصنيف السلاسل',
      'فئة السلاسل',
      'تصنيف سلاسل',
      'السلاسل',
      'Chains Category',
    ]);
    const colArmyCat = findColumnHeader(itemsTable.headers, [
      'تصنيف الجيش',
      'فئة الجيش',
      'تصنيف جيش',
      'الجيش',
      'Army Category',
    ]);
    const colTopRetailCat = findColumnHeader(itemsTable.headers, [
      'التوب ريتيل',
      'تصنيف التوب ريتيل',
      'فئة التوب ريتيل',
      'توب ريتيل',
      'التجزئة المميزة',
      'التجزئه المميزه',
      'تصنيف التجزئة المميزة',
      'تجزئة مميزة',
      'Top Retail',
      'Top Retail Category',
    ]);
    const colBrand = findColumnHeader(itemsTable.headers, ['العلامة التجارية', 'الماركة', 'Brand']);
    const colUnit = findColumnHeader(itemsTable.headers, ['الوحدة', 'Unit']);

    const missing: string[] = [];
    if (!colCode) missing.push('كود الصنف');
    if (!colName) missing.push('اسم الصنف');
    if (missing.length > 0) {
      markMissingCols('تعريف الاصناف', missing);
      errors.push({
        severity: 'error',
        sheet: 'تعريف الاصناف',
        code: 'MISSING_COLUMN',
        message: `أعمدة مطلوبة مفقودة في ورقة تعريف الاصناف: ${missing.join('، ')}`,
      });
    } else if (colCode && colName) {
      const dupes: string[] = [];
      for (const row of itemsTable.rows) {
        const code = readCellAsCode(row[colCode]);
        const name = readCellAsText(row[colName]);
        if (!code) continue;
        if (itemMap.has(code)) {
          dupes.push(code);
          continue;
        }
        const entry: ParsedItemDef = {
          itemCode: code,
          itemName: name || code,
          categoryName: colCat ? readCellAsText(row[colCat]) : '',
          brandName: colBrand ? readCellAsText(row[colBrand]) : '',
          unitName: colUnit ? readCellAsText(row[colUnit]) : '',
          chainsCategory: colChainsCat ? readCellAsText(row[colChainsCat]) : '',
          armyCategory: colArmyCat ? readCellAsText(row[colArmyCat]) : '',
          topRetailCategory: colTopRetailCat ? readCellAsText(row[colTopRetailCat]) : '',
        };
        itemMap.set(code, entry);
        itemsList.push(entry);
      }
      if (dupes.length > 0) {
        warnings.push({
          severity: 'warning',
          sheet: 'تعريف الاصناف',
          code: 'DUPLICATE_ITEM_CODE',
          message: `تكرار ${dupes.length} أكواد أصناف في تعريف الاصناف.`,
          details: dupes.slice(0, 15),
        });
      }
    }
  }

  // Parse Optional Category Targets Sheets (`أهداف الفئات` / `اهداف الفئات ...`)
  const categoryTargetsList: Array<{
    sectorName: string;
    jobTitle: string;
    employeeCode: string;
    employeeName: string;
    categoryName: string;
    targetValue: number;
  }> = [];

  const requiredNormSet = new Set(REQUIRED_SHEETS.map((s) => normalizeArabic(s)));
  const categoryTargetSheets = wb.SheetNames.filter((sName) => {
    const norm = normalizeArabic(sName);
    if (requiredNormSet.has(norm)) return false;
    const noSpace = norm.replace(/\s+/g, '');
    return (
      norm.includes('اهداف الفئ') ||
      norm.includes('اهداف فئ') ||
      norm.includes('هدف الفئ') ||
      norm.includes('هدف فئ') ||
      noSpace.includes('اهدافالفئ') ||
      noSpace.includes('اهداففئ') ||
      noSpace.includes('هدفالفئ') ||
      noSpace.includes('هدففئ')
    );
  });

  const inferSectorFromName = (name: string): string => {
    const norm = normalizeArabic(name);
    if (norm.includes('مباشر') && norm.includes('توريد')) return 'التوريد المباشر';
    if (norm.includes('مركزي') && norm.includes('توريد')) return 'التوريد المركزي';
    if ((norm.includes('توب') && norm.includes('ريتيل')) || (norm.includes('تجزئ') && norm.includes('مميز'))) {
      return 'التوب ريتيل';
    }
    if (norm.includes('جيش')) return 'الجيش';
    return '';
  };

  const inferJobTitleFromSheetName = (sheetName: string): string => {
    const cleaned = sheetName
      .replace(/أهداف\s*الفئات|اهداف\s*الفئات|أهداف\s*فئات|اهداف\s*فئات|هدف\s*الفئات|أهداف\s*الفئة|اهداف\s*الفئة/gi, '')
      .replace(/[-–—_:|/\\]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    const norm = normalizeArabic(cleaned);
    if (!norm) return '';
    if (norm === 'التوريد المباشر' || norm === 'التوريد المركزي' || norm === 'التوب ريتيل' || norm === 'الجيش') {
      return '';
    }
    return cleaned;
  };

  for (const catSheetName of categoryTargetSheets) {
    const ws = wb.Sheets[catSheetName];
    if (!ws) continue;
    const table = extractSheetTable(ws, [
      'القطاع',
      'المسمى الوظيفى',
      'المسمى الوظيفي',
      'المسمى',
      'الوظيفة',
      'المستوى الإدارى',
      'اسم الموظف',
      'الموظف',
      'الاسم',
      'كود الموظف',
      'كود المندوب',
      'اسم المندوب',
      'اسم المشرف',
      'الفئة',
      'الفئه',
      'الهدف',
      'المستهدف',
    ]);
    sheetsSummary.push({
      sheetName: catSheetName,
      found: true,
      rowCount: table.rows.length,
      detectedColumns: table.headers,
      missingColumns: [],
    });
    if (table.rows.length === 0) continue;

    const colSector = findColumnHeader(table.headers, ['القطاع', 'اسم القطاع', 'Sector']);
    const colJobTitle = findColumnHeader(table.headers, [
      'المسمى الوظيفى',
      'المسمى الوظيفي',
      'المسمي الوظيفي',
      'المسمى',
      'المسمي',
      'الوظيفة',
      'الوظيفه',
      'المستوى الادارى',
      'المستوى الاداري',
      'المستوى الإداري',
      'المستوى الوظيفى',
      'المستوى',
      'الصفة',
      'الدور',
      'Job Title',
      'Title',
      'Position',
      'Role',
      'Level',
    ]);
    const colEmpCode = findColumnHeader(table.headers, [
      'كود الموظف',
      'كود المندوب',
      'كود المشرف',
      'كود المدير',
      'الكود',
      'Employee Code',
      'Rep Code',
    ]);
    const colEmpName = findColumnHeader(table.headers, [
      'اسم الموظف',
      'الموظف',
      'اسم المسئول',
      'اسم المسؤول',
      'المسئول',
      'المسؤول',
      'اسم المندوب',
      'المندوب',
      'اسم المشرف',
      'المشرف',
      'اسم المدير',
      'المدير',
      'مدير القطاع',
      'رئيس القطاع',
      'الاسم',
      'اسم',
      'Employee Name',
      'Employee',
      'Rep Name',
      'Supervisor',
      'Manager',
      'Name',
    ]);
    const colCategory = findColumnHeader(table.headers, [
      'الفئة',
      'الفئه',
      'فئة',
      'فئه',
      'الفئات',
      'فئة الصنف',
      'فئه الصنف',
      'اسم الفئة',
      'اسم الفئه',
      'تصنيف الفئة',
      'تصنيف الصنف',
      'التصنيف',
      'Category',
      'Item Category',
    ]);
    const colTarget = findColumnHeader(table.headers, [
      'الهدف',
      'هدف',
      'هدف الفئة',
      'هدف الفئه',
      'المستهدف',
      'مستهدف',
      'قيمة الهدف',
      'اجمالى الهدف',
      'اجمالي الهدف',
      'الهدف الشهرى',
      'الهدف الشهري',
      'Target',
      'Category Target',
    ]);

    const fallbackSheetSector = inferSectorFromName(catSheetName);
    const fallbackSheetJob = inferJobTitleFromSheetName(catSheetName);

    if (colCategory && colTarget) {
      // Long format: one row per (employee/role, category, target)
      for (const row of table.rows) {
        const rawSector = colSector ? readCellAsText(row[colSector]) : '';
        const sectorName = canonicalizeSectorName(rawSector) || fallbackSheetSector;
        const empCode = colEmpCode ? readCellAsCode(row[colEmpCode]) : '';
        const rawJobTitle =
          (colJobTitle ? readCellAsText(row[colJobTitle]) : '') ||
          fallbackSheetJob ||
          (empCode ? 'مندوب مبيعات' : 'غير محدد');
        const empName =
          (colEmpName ? readCellAsText(row[colEmpName]) : '') ||
          empCode ||
          (rawJobTitle !== 'غير محدد' ? rawJobTitle : '');
        const categoryName = readCellAsText(row[colCategory]);
        if (!categoryName || (!empName && !empCode)) continue;
        const targetVal = readCellAsNumber(row[colTarget]) ?? 0;
        categoryTargetsList.push({
          sectorName,
          jobTitle: rawJobTitle,
          employeeCode: empCode,
          employeeName: empName,
          categoryName,
          targetValue: targetVal,
        });
      }
    } else if (colEmpName || colEmpCode || colJobTitle) {
      // Wide matrix format: category names are column headers with numeric target cells
      const metaHeaders = new Set<string>(
        [colSector, colJobTitle, colEmpCode, colEmpName, colTarget].filter(Boolean) as string[]
      );
      const ignoredNorms = new Set([
        'الاجمالي',
        'اجمالي',
        'اجمالي الهدف',
        'المجموع',
        'total',
        'sum',
        'ملاحظات',
        'notes',
      ]);
      const catHeaders = table.headers.filter((h) => {
        if (metaHeaders.has(h)) return false;
        const n = normalizeHeader(h);
        if (!n || ignoredNorms.has(n)) return false;
        return table.rows.some((r) => readCellAsNumber(r[h]) !== null);
      });

      for (const row of table.rows) {
        const rawSector = colSector ? readCellAsText(row[colSector]) : '';
        const sectorName = canonicalizeSectorName(rawSector) || fallbackSheetSector;
        const empCode = colEmpCode ? readCellAsCode(row[colEmpCode]) : '';
        const rawJobTitle =
          (colJobTitle ? readCellAsText(row[colJobTitle]) : '') ||
          fallbackSheetJob ||
          (empCode ? 'مندوب مبيعات' : 'غير محدد');
        const empName =
          (colEmpName ? readCellAsText(row[colEmpName]) : '') ||
          empCode ||
          (rawJobTitle !== 'غير محدد' ? rawJobTitle : '');
        if (!empName && !empCode) continue;

        for (const catHeader of catHeaders) {
          const val = readCellAsNumber(row[catHeader]);
          if (val === null) continue;
          categoryTargetsList.push({
            sectorName,
            jobTitle: rawJobTitle,
            employeeCode: empCode,
            employeeName: empName,
            categoryName: catHeader.trim(),
            targetValue: val,
          });
        }
      }
    }
  }

  // Deduplicate parsed category targets by natural key so DB upserts remain unique
  const categoryTargetDedupMap = new Map<string, (typeof categoryTargetsList)[number]>();
  for (const ct of categoryTargetsList) {
    const cleanCt = {
      sectorName: ct.sectorName.trim(),
      jobTitle: ct.jobTitle.trim(),
      employeeCode: ct.employeeCode.trim(),
      employeeName: ct.employeeName.trim(),
      categoryName: ct.categoryName.trim(),
      targetValue: ct.targetValue,
    };
    const key = `${cleanCt.sectorName}\u0000${cleanCt.jobTitle}\u0000${cleanCt.employeeName}\u0000${cleanCt.categoryName}`;
    categoryTargetDedupMap.set(key, cleanCt);
  }
  const dedupedCategoryTargetsList = Array.from(categoryTargetDedupMap.values());

  // Parse Users (`تعريف المستخدمين`)
  const importedUsers: Array<{ username: string; fullName: string; email: string | null; role: UserRole; allowedSectors: string[] }> = [];
  const usersTable = sheetTables.get('تعريف المستخدمين');
  if (usersTable) {
    const colUser = findColumnHeader(usersTable.headers, ['اسم المستخدم', 'المستخدم', 'Username', 'كود المستخدم', 'Email', 'البريد الإلكتروني']);
    const colName = findColumnHeader(usersTable.headers, ['الاسم', 'الاسم الكامل', 'Full Name', 'اسم الموظف']);
    const colEmail = findColumnHeader(usersTable.headers, ['البريد الإلكتروني', 'Email', 'البريد']);
    const colRole = findColumnHeader(usersTable.headers, ['الدور', 'Role', 'نوع الصلاحية', 'الدور الوظيفي', 'الصلاحية', 'الصلاحيات']);
    const colSectors = findColumnHeader(usersTable.headers, ['القطاعات المسموحة', 'القطاعات', 'القطاع', 'Allowed Sectors', 'Sector', 'الصلاحيات']);

    if (!colUser && !colName) {
      markMissingCols('تعريف المستخدمين', ['اسم المستخدم']);
      errors.push({
        severity: 'error',
        sheet: 'تعريف المستخدمين',
        code: 'MISSING_COLUMN',
        message: 'عمود اسم المستخدم مفقود في ورقة تعريف المستخدمين.',
      });
    } else {
      const seenUsers = new Set<string>();
      for (const row of usersTable.rows) {
        const rawUser = colUser ? readCellAsText(row[colUser]) : '';
        const rawName = colName ? readCellAsText(row[colName]) : rawUser;
        const username = (rawUser || rawName).trim().toLowerCase().replace(/\s+/g, '.');
        if (!username) continue;
        if (seenUsers.has(username)) continue;
        seenUsers.add(username);

        const rawRole = colRole ? normalizeArabic(readCellAsText(row[colRole])) : '';
        const isAdmin =
          rawRole.includes('ادمن') ||
          rawRole.includes('admin') ||
          rawRole.includes('مدير نظام') ||
          rawRole.includes('مدير النظام') ||
          rawRole.includes('مدير عام') ||
          rawRole === 'الكل' ||
          rawRole === 'all';

        const isSectorManager =
          !isAdmin &&
          (rawRole.includes('مدير قطاع') ||
            rawRole.includes('مدير ادارة') ||
            rawRole.includes('مدير الادارة') ||
            rawRole.includes('مدير ادار') ||
            rawRole.includes('sector_manager') ||
            rawRole.includes('manager') ||
            rawRole.includes('مدير'));

        const role: UserRole = isAdmin
          ? 'admin'
          : isSectorManager
            ? 'sector_manager'
            : 'standard_user';

        const rawSectors = colSectors ? readCellAsText(row[colSectors]) : '';
        const allowedSectors = isAdmin
          ? [...ORIGINAL_SECTORS]
          : rawSectors
              .split(/[,،;|/]/)
              .map((s) => canonicalizeSectorName(s))
              .filter((s) => ORIGINAL_SECTORS.includes(s as (typeof ORIGINAL_SECTORS)[number]));

        importedUsers.push({
          username,
          fullName: rawName || username,
          email: colEmail ? readCellAsText(row[colEmail]) || null : null,
          role,
          allowedSectors,
        });
      }
    }
  }

  // 3. Parse Sales Sheets (`VF05N (2)` and `Zsales_Price (2)`)
  const parseSalesSheet = (sheetName: 'VF05N (2)' | 'Zsales_Price (2)', sourceType: 'VF05N' | 'ZSALES') => {
    const table = sheetTables.get(sheetName);
    const records: ParsedSalesRow[] = [];
    if (!table) return records;

    const isZsales = sourceType === 'ZSALES';

    const colSales = findColumnHeader(table.headers, [
      'New Growth Value BD&T',
      'New Growth Value BD & T',
      'New Growth Value',
    ]);
    const colNetSales = isZsales
      ? findColumnHeader(table.headers, [
          'Net Value Without TA',
          'Net Value Without Tax',
          'Net Value W/O TA',
        ])
      : findColumnHeader(table.headers, [
          'New Net Value BT',
          'New Net Value B&T',
          'Net Value BT',
        ]);
    const colTax = isZsales
      ? findColumnHeader(table.headers, [
          'Total Tax',
          'New Total Tax',
          'Output Tax',
          'Tax Amount',
          'Tax Value',
          'VAT Amount',
          'VAT Value',
          'VAT',
          'Tax',
          'قيمة الضريبة',
          'إجمالي الضريبة',
          'اجمالي الضريبة',
          'ضريبة القيمة المضافة',
          'الضريبة',
          'ضريبة',
          'MWSBP',
        ])
      : findColumnHeader(table.headers, [
          'New Total Tax',
          'Total Tax',
          'Output Tax',
          'Tax Amount',
          'Tax Value',
          'VAT Amount',
          'VAT Value',
          'VAT',
          'Tax',
          'قيمة الضريبة',
          'إجمالي الضريبة',
          'اجمالي الضريبة',
          'ضريبة القيمة المضافة',
          'الضريبة',
          'ضريبة',
          'MWSBP',
        ]);
    const colPayer = isZsales
      ? findColumnHeader(table.headers, ['Customer', 'Customer Code', 'كود العميل', 'كود السلسلة'])
      : findColumnHeader(table.headers, ['Payer', 'Payer Code', 'كود السلسلة']);
    const colPayerName = isZsales
      ? findColumnHeader(table.headers, ['Customer Name', 'Name of Customer', 'Name 1', 'اسم العميل', 'اسم السلسلة'])
      : findColumnHeader(table.headers, ['Name of Payer', 'Payer Name', 'اسم السلسلة']);
    const colSoldTo = isZsales
      ? null
      : findColumnHeader(table.headers, ['Sold-to Party', 'Sold-To Party', 'Sold to party', 'Sold-To', 'كود الفرع']);
    const colSoldToName = isZsales
      ? null
      : findColumnHeader(table.headers, ['Name of Sold-to Party', 'Sold-to Party Name', 'Sold-To Name', 'اسم الفرع']);
    const colDate = isZsales
      ? findColumnHeader(table.headers, ['Billing Dt.', 'Billing Dt', 'Billing Date', 'Document Date', 'Date', 'التاريخ', 'تاريخ الفاتورة'])
      : findColumnHeader(table.headers, ['Billing Date', 'Billing Dt.', 'Billing Dt', 'Document Date', 'Date', 'التاريخ', 'تاريخ الفاتورة']);
    const colPeriod = findColumnHeader(table.headers, [
      'الفترة',
      'الشهر',
      'الشهر/السنة',
      'الشهر والسنة',
      'الشهر - السنة',
      'تاريخ الفترة',
      'Period',
      'Month',
      'Fiscal Period',
      'Year-Month',
      'Year Month',
      'CalMonth',
      'Billing Month',
    ]);
    const colDoc = findColumnHeader(table.headers, ['Billing Document', 'Document', 'Invoice', 'رقم الفاتورة']);
    const colItem = findColumnHeader(table.headers, ['Material', 'Material Code', 'Item Code', 'كود الصنف']);
    const colItemName = findColumnHeader(table.headers, ['Material Description', 'Item Name', 'Description', 'اسم الصنف']);
    const colRep = findColumnHeader(table.headers, ['Sales Emp.', 'Sales Emp', 'Sales Employee Code', 'Sales Employee', 'Sales Rep', 'Rep Code', 'كود المندوب']);
    const colQty = findColumnHeader(table.headers, ['Billed Quantity', 'Qty', 'Quantity', 'الكمية']);

    const missing: string[] = [];
    if (!colSales) missing.push('New Growth Value BD&T');
    if (isZsales) {
      if (!colPayer) missing.push('Customer');
      if (!colRep) missing.push('Sales Emp.');
      if (!colItem) missing.push('Material');
    } else {
      if (!colPayer) missing.push('Payer');
      if (!colSoldTo) missing.push('Sold-to Party');
    }

    if (missing.length > 0) {
      markMissingCols(sheetName, missing);
      errors.push({
        severity: 'error',
        sheet: sheetName,
        code: 'MISSING_COLUMN',
        message: `أعمدة أساسية مفقودة في ورقة ${sheetName}: ${missing.join('، ')}`,
      });
      return records;
    }

    if (!colNetSales) {
      const expectedNetCol = isZsales ? 'Net Value Without TA' : 'New Net Value BT';
      warnings.push({
        severity: 'warning',
        sheet: sheetName,
        code: 'NET_SALES_COLUMN_NOT_FOUND',
        message: `لم يتم العثور على عمود صافي قيمة المبيعات (${expectedNetCol}) في ورقة ${sheetName}، سيتم احتساب صافي قيمة المبيعات كـ 0 لهذه الورقة.`,
      });
    }

    if (!colTax) {
      warnings.push({
        severity: 'warning',
        sheet: sheetName,
        code: 'TAX_COLUMN_NOT_FOUND',
        message: `لم يتم العثور على عمود للضريبة في ورقة ${sheetName}، سيتم احتساب الضريبة كـ 0 لهذه الورقة.`,
      });
    }

    if (table.rows.length === 0) {
      errors.push({
        severity: 'error',
        sheet: sheetName,
        code: 'EMPTY_SALES_SHEET',
        message: `ورقة المبيعات "${sheetName}" فارغة لا تحتوي على أي سجلات.`,
      });
      return records;
    }

    let invalidSalesRows = 0;
    let invalidNetSalesRows = 0;
    for (const row of table.rows) {
      const salesNum = readCellAsNumber(row[colSales!]);
      if (salesNum === null) {
        invalidSalesRows++;
        continue;
      }
      let netSalesNum = 0;
      if (colNetSales) {
        const netCell = row[colNetSales];
        const hasNetContent = netCell !== undefined && netCell.v !== undefined && netCell.v !== null && String(netCell.v).trim() !== '';
        if (hasNetContent) {
          const parsedNet = readCellAsNumber(netCell);
          if (parsedNet === null) {
            invalidNetSalesRows++;
            continue;
          }
          netSalesNum = parsedNet;
        }
      }
      const taxNum = colTax ? (readCellAsNumber(row[colTax]) ?? 0) : 0;
      const qtyNum = colQty ? (readCellAsNumber(row[colQty]) ?? 0) : 0;

      const payerCode = readCellAsCode(row[colPayer!]);
      const payerName = colPayerName ? readCellAsText(row[colPayerName]) : '';
      const soldToCode = colSoldTo ? readCellAsCode(row[colSoldTo]) : '';
      const soldToName = colSoldToName ? readCellAsText(row[colSoldToName]) : '';
      const docNumber = colDoc ? readCellAsCode(row[colDoc]) : '';
      const rawBillingDate = colDate ? readCellAsDate(row[colDate]) : '';
      const { monthKey, billingDate } = extractBillingMonthAndDate(
        rawBillingDate,
        colDate ? row[colDate] : undefined,
        colPeriod ? row[colPeriod] : undefined
      );
      const itemCode = colItem ? readCellAsCode(row[colItem]) : '';
      const itemName = colItemName ? readCellAsText(row[colItemName]) : '';
      const rawRepCode = colRep ? readCellAsCode(row[colRep]) : '';

      // Join 1: Payer (or Customer in Zsales_Price) -> Chain/Department -> Sector
      // Priority: exact preserved-leading-zero code -> zero-trimmed normalized code -> normalized chain name (`السلسلة`)
      const chainMatch =
        (payerCode ? chainMap.get(payerCode) : undefined) ??
        (payerCode ? chainTrimmedMap.get(stripLeadingZeros(payerCode)) : undefined) ??
        (payerName ? chainByNameMap.get(normalizeArabic(payerName)) : undefined);
      // Join 2: Sold-to Party -> Branch (VF05N only; Zsales_Price has no Sold-to Party)
      // Priority: exact preserved-leading-zero code -> zero-trimmed normalized code
      const branchMatch =
        (soldToCode ? branchMap.get(soldToCode) : undefined) ??
        (soldToCode ? branchTrimmedMap.get(stripLeadingZeros(soldToCode)) : undefined);
      // Join 3: Rep code (Sales Emp. / Sales Employee or branch metadata if defined)
      const effectiveRepCode = rawRepCode || branchMatch?.repCode || '';
      const repMatch = effectiveRepCode ? repMap.get(effectiveRepCode) : undefined;
      // Join 4: Item code (Material)
      const itemMatch = itemCode ? itemMap.get(itemCode) : undefined;

      records.push({
        sourceType,
        docNumber,
        billingDate,
        monthKey,
        payerCode: payerCode || chainMatch?.chainCode || '',
        payerName,
        soldToCode,
        soldToName,
        itemCode,
        itemName: itemName || itemMatch?.itemName || '',
        repCode: effectiveRepCode,
        quantity: qtyNum,
        salesValue: salesNum,
        netSalesValue: netSalesNum,
        taxValue: taxNum,
        resolvedSector: chainMatch ? chainMatch.sectorName : null,
        resolvedDepartment: chainMatch ? chainMatch.departmentName : null,
        resolvedChainName: chainMatch ? chainMatch.chainName : null,
        resolvedBranchCode: branchMatch ? branchMatch.branchCode : null,
        resolvedBranchName: branchMatch ? branchMatch.branchName : null,
        resolvedRepCode: repMatch ? repMatch.repCode : effectiveRepCode || null,
        resolvedRepName: repMatch ? repMatch.repName : null,
        resolvedItemName: itemMatch ? itemMatch.itemName : itemName || null,
        isReturn: salesNum < 0 ? 1 : 0,
      });
    }

    if (invalidSalesRows > 0) {
      errors.push({
        severity: 'error',
        sheet: sheetName,
        code: 'INVALID_SALES_VALUES',
        message: `تم العثور على ${invalidSalesRows} صفوف تحتوي على قيم غير رقمية في عمود ${colSales} داخل ورقة ${sheetName}.`,
      });
    }
    if (invalidNetSalesRows > 0) {
      errors.push({
        severity: 'error',
        sheet: sheetName,
        code: 'INVALID_NET_SALES_VALUES',
        message: `تم العثور على ${invalidNetSalesRows} صفوف تحتوي على قيم غير رقمية في عمود ${colNetSales} داخل ورقة ${sheetName}.`,
      });
    }

    return records;
  };

  const vf05nRows = parseSalesSheet('VF05N (2)', 'VF05N');
  const zsalesRows = parseSalesSheet('Zsales_Price (2)', 'ZSALES');

  const summarizeRows = (rows: ParsedSalesRow[]) => {
    let salesTotal = 0;
    let netSalesTotal = 0;
    let taxTotal = 0;
    let returnCount = 0;
    for (const r of rows) {
      salesTotal += r.salesValue;
      netSalesTotal += r.netSalesValue;
      taxTotal += r.taxValue;
      if (r.salesValue < 0) returnCount++;
    }
    return {
      count: rows.length,
      salesTotal: Number(salesTotal.toFixed(2)),
      netSalesTotal: Number(netSalesTotal.toFixed(2)),
      taxTotal: Number(taxTotal.toFixed(2)),
      returnCount,
    };
  };

  const vf05nSum = summarizeRows(vf05nRows);
  const zsalesSum = summarizeRows(zsalesRows);
  const salesDiff = Number((vf05nSum.salesTotal - zsalesSum.salesTotal).toFixed(2));
  const netSalesDiff = Number((vf05nSum.netSalesTotal - zsalesSum.netSalesTotal).toFixed(2));
  const taxDiff = Number((vf05nSum.taxTotal - zsalesSum.taxTotal).toFixed(2));
  const isMatched = Math.abs(salesDiff) <= 0.05 && vf05nSum.count > 0 && zsalesSum.count > 0;

  if (vf05nSum.count > 0 && zsalesSum.count > 0 && Math.abs(salesDiff) > 0.05) {
    warnings.push({
      severity: 'warning',
      sheet: 'VF05N (2) / Zsales_Price (2)',
      code: 'SOURCE_TOTAL_MISMATCH',
      message: `يوجد فرق بقيمة ${salesDiff.toLocaleString('en-US', { minimumFractionDigits: 2 })} EGP بين إجمالي مبيعات VF05N (${vf05nSum.salesTotal.toLocaleString('en-US')}) وZsales_Price (${zsalesSum.salesTotal.toLocaleString('en-US')}).`,
    });
  }

  // Link diagnostics on VF05N
  let vf05nSectorLinked = 0;
  let vf05nBranchLinked = 0;
  const unlinkedBranchAgg = new Map<string, { soldToCode: string; soldToName: string; rowCount: number; salesTotal: number }>();
  const unlinkedSectorAgg = new Map<string, { payerCode: string; payerName: string; rowCount: number; salesTotal: number }>();

  for (const r of vf05nRows) {
    if (r.resolvedSector) {
      vf05nSectorLinked++;
    } else {
      const key = r.payerCode || '(بدون كود)';
      const cur = unlinkedSectorAgg.get(key) ?? {
        payerCode: key,
        payerName: r.payerName || 'غير محدد',
        rowCount: 0,
        salesTotal: 0,
      };
      cur.rowCount++;
      cur.salesTotal += r.salesValue;
      unlinkedSectorAgg.set(key, cur);
    }

    if (r.resolvedBranchCode) {
      vf05nBranchLinked++;
    } else {
      const key = r.soldToCode || '(بدون كود)';
      const cur = unlinkedBranchAgg.get(key) ?? {
        soldToCode: key,
        soldToName: r.soldToName || 'غير محدد',
        rowCount: 0,
        salesTotal: 0,
      };
      cur.rowCount++;
      cur.salesTotal += r.salesValue;
      unlinkedBranchAgg.set(key, cur);
    }
  }

  const unlinkedBranchCodes = Array.from(unlinkedBranchAgg.values()).sort((a, b) => b.rowCount - a.rowCount);
  const unlinkedSectorCodes = Array.from(unlinkedSectorAgg.values()).sort((a, b) => b.rowCount - a.rowCount);

  if (unlinkedBranchCodes.length > 0) {
    const totalUnlinked = vf05nRows.length - vf05nBranchLinked;
    warnings.push({
      severity: 'warning',
      sheet: 'VF05N (2)',
      code: 'UNLINKED_BRANCHES',
      message: `تم ربط ${vf05nBranchLinked} سجلًا بالفروع من أصل ${vf05nRows.length}، وهناك ${totalUnlinked} سجلًا غير مرتبط بفرع (سيتم عرضها ضمن السجلات غير المرتبطة دون إسناد تخميني).`,
      details: unlinkedBranchCodes
        .slice(0, 15)
        .map((u) => `كود ${u.soldToCode} (${u.soldToName}): ${u.rowCount} سجل - ${u.salesTotal.toFixed(2)} EGP`),
    });
  }

  if (unlinkedSectorCodes.length > 0) {
    const totalUnlinked = vf05nRows.length - vf05nSectorLinked;
    warnings.push({
      severity: 'warning',
      sheet: 'VF05N (2)',
      code: 'UNLINKED_SECTORS',
      message: `يوجد ${totalUnlinked} سجلًا في VF05N غير مرتبط بقطاع عبر عمود Payer (كود السلسلة).`,
      details: unlinkedSectorCodes
        .slice(0, 15)
        .map((u) => `Payer ${u.payerCode} (${u.payerName}): ${u.rowCount} سجل - ${u.salesTotal.toFixed(2)} EGP`),
    });
  }

  const existingPeriodSet = new Set(existingPeriods.map((p) => String(p || '').trim()).filter(Boolean));
  const periodMap = new Map<string, ImportPeriodSummary>();
  for (const r of [...vf05nRows, ...zsalesRows]) {
    const key = r.monthKey || extractMonthKey(r.billingDate);
    const cur = periodMap.get(key) ?? {
      period: key,
      vf05nCount: 0,
      vf05nSalesTotal: 0,
      vf05nNetSalesTotal: 0,
      zsalesCount: 0,
      zsalesSalesTotal: 0,
      zsalesNetSalesTotal: 0,
      action: existingPeriodSet.has(key) ? ('replace' as const) : ('add' as const),
    };
    if (r.sourceType === 'VF05N') {
      cur.vf05nCount += 1;
      cur.vf05nSalesTotal = Number((cur.vf05nSalesTotal + r.salesValue).toFixed(2));
      cur.vf05nNetSalesTotal = Number((cur.vf05nNetSalesTotal + r.netSalesValue).toFixed(2));
    } else {
      cur.zsalesCount += 1;
      cur.zsalesSalesTotal = Number((cur.zsalesSalesTotal + r.salesValue).toFixed(2));
      cur.zsalesNetSalesTotal = Number((cur.zsalesNetSalesTotal + r.netSalesValue).toFixed(2));
    }
    periodMap.set(key, cur);
  }
  const periodsDetected = Array.from(periodMap.values()).sort((a, b) => {
    if (a.period === 'UNSPECIFIED') return 1;
    if (b.period === 'UNSPECIFIED') return -1;
    return a.period.localeCompare(b.period);
  });

  const totalSectorTargets = sectorsList.reduce((acc, s) => acc + s.monthlyTarget, 0);
  const directBranchesTargetTotal = branchesList.reduce((acc, b) => acc + b.branchTarget, 0);

  const validationToken = globalThis.crypto.randomUUID();
  const canCommit = errors.length === 0;

  const report: ImportValidationReport = {
    validationToken,
    filename,
    valid: canCommit,
    canCommit,
    errors,
    warnings,
    sheets: sheetsSummary,
    comparison: {
      vf05n: vf05nSum,
      zsales: zsalesSum,
      salesDifference: salesDiff,
      netSalesDifference: netSalesDiff,
      taxDifference: taxDiff,
      isMatched,
    },
    periodsDetected,
    linkStats: {
      vf05nTotalRows: vf05nRows.length,
      vf05nSectorLinkedRows: vf05nSectorLinked,
      vf05nSectorUnlinkedRows: vf05nRows.length - vf05nSectorLinked,
      vf05nBranchLinkedRows: vf05nBranchLinked,
      vf05nBranchUnlinkedRows: vf05nRows.length - vf05nBranchLinked,
      unlinkedBranchCodes,
      unlinkedSectorCodes,
    },
    targetsSummary: {
      sectors: sectorsList.map((s) => ({ sectorName: s.sectorName, target: s.monthlyTarget })),
      totalSectorTargets: Number(totalSectorTargets.toFixed(2)),
      directBranchesTargetTotal: Number(directBranchesTargetTotal.toFixed(2)),
      directBranchCount: branchesList.length,
    },
    usersFound: importedUsers.map((u) => ({
      username: u.username,
      fullName: u.fullName,
      role: u.role,
      allowedSectors: u.allowedSectors,
    })),
  };

  if (canCommit) {
    stagedImports.set(validationToken, {
      validationToken,
      filename,
      createdAtMs: Date.now(),
      report,
      sectors: sectorsList,
      chains: chainsList,
      branches: branchesList,
      reps: repsList,
      items: itemsList,
      categoryTargets: dedupedCategoryTargetsList,
      hasCategoryTargetsSheet: categoryTargetSheets.length > 0,
      salesRecords: [...vf05nRows, ...zsalesRows],
      importedUsers,
    });
  }

  return report;
}

export function commitValidatedImport(
  db: Database.Database,
  validationToken: string,
  user: { id: number; fullName: string }
): ImportBatchHistoryItem {
  const staged = stagedImports.get(validationToken);
  if (!staged) {
    throw new Error('انتهت صلاحية جلسة الفحص أو الرمز غير صالح. يرجى إعادة فحص الملف.');
  }

  const now = new Date().toISOString();
  const batchUuid = globalThis.crypto.randomUUID();
  const { report } = staged;

  const tx = db.transaction(() => {
    const prevBatch = db
      .prepare(`SELECT id, has_category_targets_sheet FROM import_batches WHERE status = 'active' ORDER BY id DESC LIMIT 1`)
      .get() as { id: number; has_category_targets_sheet: number } | undefined;

    // 1. Merge / Upsert Master Reference Data by natural unique keys (never duplicate, preserve unmentioned entries)
    const mergedSectorsMap = new Map<string, StagedImportPayload['sectors'][number]>();
    const mergedChainsMap = new Map<string, StagedImportPayload['chains'][number]>();
    const mergedBranchesMap = new Map<string, StagedImportPayload['branches'][number]>();
    const mergedRepsMap = new Map<string, StagedImportPayload['reps'][number]>();
    const mergedItemsMap = new Map<string, StagedImportPayload['items'][number]>();
    const mergedCatTargetsMap = new Map<string, StagedImportPayload['categoryTargets'][number]>();

    if (prevBatch) {
      const prevSectors = db
        .prepare(`SELECT sector_code, sector_name, monthly_target, is_original_sector FROM sectors WHERE batch_id = ?`)
        .all(prevBatch.id) as Array<{
        sector_code: string;
        sector_name: string;
        monthly_target: number;
        is_original_sector: number;
      }>;
      for (const s of prevSectors) {
        const key = canonicalizeSectorName(s.sector_name);
        if (!key) continue;
        mergedSectorsMap.set(key, {
          sectorCode: s.sector_code,
          sectorName: key,
          monthlyTarget: s.monthly_target,
          isOriginal: s.is_original_sector,
        });
      }

      const prevChains = db
        .prepare(`SELECT chain_code, chain_name, department_name, sector_name, supervisor_name FROM departments_chains WHERE batch_id = ?`)
        .all(prevBatch.id) as Array<{
        chain_code: string;
        chain_name: string;
        department_name: string;
        sector_name: string;
        supervisor_name: string;
      }>;
      for (const c of prevChains) {
        const key = normalizeCode(c.chain_code);
        if (!key) continue;
        mergedChainsMap.set(key, {
          chainCode: key,
          chainName: c.chain_name,
          departmentName: c.department_name,
          sectorName: canonicalizeSectorName(c.sector_name),
          supervisorName: c.supervisor_name || '',
        });
      }

      const prevBranches = db
        .prepare(
          `SELECT branch_code, branch_name, sector_name, branch_target, chain_code, chain_name, rep_code, rep_name, supervisor_name
           FROM branches WHERE batch_id = ?`
        )
        .all(prevBatch.id) as Array<{
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
      for (const b of prevBranches) {
        const key = normalizeCode(b.branch_code);
        if (!key) continue;
        mergedBranchesMap.set(key, {
          branchCode: key,
          branchName: b.branch_name,
          sectorName: canonicalizeSectorName(b.sector_name),
          branchTarget: b.branch_target,
          chainCode: normalizeCode(b.chain_code || ''),
          chainName: b.chain_name || '',
          repCode: normalizeCode(b.rep_code || ''),
          repName: b.rep_name || '',
          supervisorName: b.supervisor_name || '',
        });
      }

      const prevReps = db
        .prepare(`SELECT rep_code, rep_name, sector_name, supervisor_name, monthly_target FROM sales_reps WHERE batch_id = ?`)
        .all(prevBatch.id) as Array<{
        rep_code: string;
        rep_name: string;
        sector_name: string | null;
        supervisor_name: string | null;
        monthly_target: number;
      }>;
      for (const r of prevReps) {
        const key = normalizeCode(r.rep_code);
        if (!key) continue;
        mergedRepsMap.set(key, {
          repCode: key,
          repName: r.rep_name,
          sectorName: r.sector_name ? canonicalizeSectorName(r.sector_name) : '',
          supervisorName: r.supervisor_name || '',
          monthlyTarget: r.monthly_target,
        });
      }

      const prevItems = db
        .prepare(
          `SELECT item_code, item_name, category_name, brand_name, unit_name, chains_category, army_category, top_retail_category
           FROM items WHERE batch_id = ?`
        )
        .all(prevBatch.id) as Array<{
        item_code: string;
        item_name: string;
        category_name: string | null;
        brand_name: string | null;
        unit_name: string | null;
        chains_category: string;
        army_category: string;
        top_retail_category: string;
      }>;
      for (const it of prevItems) {
        const key = normalizeCode(it.item_code);
        if (!key) continue;
        mergedItemsMap.set(key, {
          itemCode: key,
          itemName: it.item_name,
          categoryName: it.category_name || '',
          brandName: it.brand_name || '',
          unitName: it.unit_name || '',
          chainsCategory: it.chains_category || '',
          armyCategory: it.army_category || '',
          topRetailCategory: it.top_retail_category || '',
        });
      }

      const prevCatTargets = db
        .prepare(`SELECT sector_name, job_title, employee_code, employee_name, category_name, target_value FROM category_targets WHERE batch_id = ?`)
        .all(prevBatch.id) as Array<{
        sector_name: string;
        job_title: string;
        employee_code: string;
        employee_name: string;
        category_name: string;
        target_value: number;
      }>;
      for (const ct of prevCatTargets) {
        const cleanCt = {
          sectorName: (ct.sector_name || '').trim(),
          jobTitle: (ct.job_title || '').trim(),
          employeeCode: (ct.employee_code || '').trim(),
          employeeName: (ct.employee_name || '').trim(),
          categoryName: (ct.category_name || '').trim(),
          targetValue: ct.target_value,
        };
        const key = `${cleanCt.sectorName}\u0000${cleanCt.jobTitle}\u0000${cleanCt.employeeName}\u0000${cleanCt.categoryName}`;
        mergedCatTargetsMap.set(key, cleanCt);
      }
    }

    // Apply Upserts from newly staged Excel payload (updates matching natural keys, inserts new ones once)
    for (const s of staged.sectors) {
      const key = canonicalizeSectorName(s.sectorName);
      if (!key) continue;
      mergedSectorsMap.set(key, { ...s, sectorName: key });
    }
    for (const c of staged.chains) {
      const key = normalizeCode(c.chainCode);
      if (!key) continue;
      mergedChainsMap.set(key, { ...c, chainCode: key });
    }
    for (const b of staged.branches) {
      const key = normalizeCode(b.branchCode);
      if (!key) continue;
      mergedBranchesMap.set(key, { ...b, branchCode: key });
    }
    for (const r of staged.reps) {
      const key = normalizeCode(r.repCode);
      if (!key) continue;
      mergedRepsMap.set(key, { ...r, repCode: key });
    }
    for (const it of staged.items) {
      const key = normalizeCode(it.itemCode);
      if (!key) continue;
      mergedItemsMap.set(key, { ...it, itemCode: key });
    }
    for (const ct of staged.categoryTargets) {
      const cleanCt = {
        sectorName: ct.sectorName.trim(),
        jobTitle: ct.jobTitle.trim(),
        employeeCode: ct.employeeCode.trim(),
        employeeName: ct.employeeName.trim(),
        categoryName: ct.categoryName.trim(),
        targetValue: ct.targetValue,
      };
      const key = `${cleanCt.sectorName}\u0000${cleanCt.jobTitle}\u0000${cleanCt.employeeName}\u0000${cleanCt.categoryName}`;
      mergedCatTargetsMap.set(key, cleanCt);
    }

    const mergedSectors = Array.from(mergedSectorsMap.values());
    const mergedChains = Array.from(mergedChainsMap.values());
    const mergedBranches = Array.from(mergedBranchesMap.values());
    const mergedReps = Array.from(mergedRepsMap.values());
    const mergedItems = Array.from(mergedItemsMap.values());
    const mergedCatTargets = Array.from(mergedCatTargetsMap.values());

    // Helper lookup maps on merged catalog for re-linking carried-over historical sales rows
    const chainTrimmedLookup = new Map<string, StagedImportPayload['chains'][number]>();
    const chainByNameLookup = new Map<string, StagedImportPayload['chains'][number]>();
    for (const c of mergedChains) {
      const trimmed = stripLeadingZeros(c.chainCode);
      if (trimmed && !chainTrimmedLookup.has(trimmed)) chainTrimmedLookup.set(trimmed, c);
      const normName = normalizeArabic(c.chainName);
      if (normName && !chainByNameLookup.has(normName)) chainByNameLookup.set(normName, c);
    }
    const branchTrimmedLookup = new Map<string, StagedImportPayload['branches'][number]>();
    for (const b of mergedBranches) {
      const trimmed = stripLeadingZeros(b.branchCode);
      if (trimmed && !branchTrimmedLookup.has(trimmed)) branchTrimmedLookup.set(trimmed, b);
    }

    // 2. Partition Sales Records by extracted Year-Month (from billing_date):
    //    Replace ONLY overlapping periods present in the uploaded file; preserve all other periods!
    const incomingPeriods = Array.from(new Set(staged.salesRecords.map((r) => r.monthKey || extractMonthKey(r.billingDate))));
    const incomingPeriodSet = new Set(incomingPeriods);
    const prevPeriodSet = new Set<string>();
    const retainedSales: ParsedSalesRow[] = [];

    if (prevBatch) {
      const prevSalesRows = db
        .prepare(
          `SELECT source_type, doc_number, billing_date, payer_code, payer_name,
                  sold_to_code, sold_to_name, item_code, item_name, rep_code, quantity, sales_value, net_sales_value, tax_value,
                  resolved_sector, resolved_department, resolved_chain_name,
                  resolved_branch_code, resolved_branch_name, resolved_rep_code, resolved_rep_name, resolved_item_name,
                  is_return
           FROM sales_records WHERE batch_id = ?`
        )
        .all(prevBatch.id) as Array<{
        source_type: 'VF05N' | 'ZSALES';
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
        net_sales_value: number;
        tax_value: number;
        resolved_sector: string | null;
        resolved_department: string | null;
        resolved_chain_name: string | null;
        resolved_branch_code: string | null;
        resolved_branch_name: string | null;
        resolved_rep_code: string | null;
        resolved_rep_name: string | null;
        resolved_item_name: string | null;
        is_return: number;
      }>;

      for (const r of prevSalesRows) {
        const bDate = (r.billing_date || '').trim();
        const monthKey = extractMonthKey(bDate);
        prevPeriodSet.add(monthKey);
        if (incomingPeriodSet.has(monthKey)) {
          // Replaced by the newly uploaded data for this period
          continue;
        }

        const pCode = normalizeCode(r.payer_code || '');
        const pName = r.payer_name || '';
        const sCode = normalizeCode(r.sold_to_code || '');
        const rCode = normalizeCode(r.rep_code || '');
        const iCode = normalizeCode(r.item_code || '');

        const chainMatch =
          (pCode ? mergedChainsMap.get(pCode) : undefined) ??
          (pCode ? chainTrimmedLookup.get(stripLeadingZeros(pCode)) : undefined) ??
          (pName ? chainByNameLookup.get(normalizeArabic(pName)) : undefined);
        const branchMatch =
          (sCode ? mergedBranchesMap.get(sCode) : undefined) ??
          (sCode ? branchTrimmedLookup.get(stripLeadingZeros(sCode)) : undefined);
        const effectiveRepCode = rCode || branchMatch?.repCode || '';
        const repMatch = effectiveRepCode ? mergedRepsMap.get(effectiveRepCode) : undefined;
        const itemMatch = iCode ? mergedItemsMap.get(iCode) : undefined;

        retainedSales.push({
          sourceType: r.source_type,
          docNumber: r.doc_number || '',
          billingDate: bDate,
          monthKey,
          payerCode: pCode || chainMatch?.chainCode || '',
          payerName: pName,
          soldToCode: sCode,
          soldToName: r.sold_to_name || '',
          itemCode: iCode,
          itemName: r.item_name || itemMatch?.itemName || '',
          repCode: effectiveRepCode,
          quantity: r.quantity,
          salesValue: r.sales_value,
          netSalesValue: r.net_sales_value ?? 0,
          taxValue: r.tax_value,
          resolvedSector: chainMatch ? chainMatch.sectorName : r.resolved_sector,
          resolvedDepartment: chainMatch ? chainMatch.departmentName : r.resolved_department,
          resolvedChainName: chainMatch ? chainMatch.chainName : r.resolved_chain_name,
          resolvedBranchCode: branchMatch ? branchMatch.branchCode : r.resolved_branch_code,
          resolvedBranchName: branchMatch ? branchMatch.branchName : r.resolved_branch_name,
          resolvedRepCode: repMatch ? repMatch.repCode : effectiveRepCode || r.resolved_rep_code,
          resolvedRepName: repMatch ? repMatch.repName : r.resolved_rep_name,
          resolvedItemName: itemMatch ? itemMatch.itemName : r.item_name || r.resolved_item_name,
          isReturn: r.is_return,
        });
      }
    }

    const replacedPeriods = incomingPeriods.filter((p) => prevPeriodSet.has(p)).sort();
    const addedPeriods = incomingPeriods.filter((p) => !prevPeriodSet.has(p)).sort();
    const retainedPeriods = Array.from(prevPeriodSet)
      .filter((p) => !incomingPeriodSet.has(p))
      .sort();

    const allSalesRecords = [...retainedSales, ...staged.salesRecords];

    let vf05nCount = 0;
    let vf05nSalesTotal = 0;
    let vf05nNetSalesTotal = 0;
    let vf05nTaxTotal = 0;
    let zsalesCount = 0;
    let zsalesSalesTotal = 0;
    let zsalesNetSalesTotal = 0;
    let zsalesTaxTotal = 0;
    let unlinkedBranchCount = 0;
    let unlinkedSectorCount = 0;

    for (const r of allSalesRecords) {
      if (r.sourceType === 'VF05N') {
        vf05nCount++;
        vf05nSalesTotal += r.salesValue;
        vf05nNetSalesTotal += r.netSalesValue;
        vf05nTaxTotal += r.taxValue;
        if (!r.resolvedBranchCode) unlinkedBranchCount++;
        if (!r.resolvedSector) unlinkedSectorCount++;
      } else {
        zsalesCount++;
        zsalesSalesTotal += r.salesValue;
        zsalesNetSalesTotal += r.netSalesValue;
        zsalesTaxTotal += r.taxValue;
      }
    }

    vf05nSalesTotal = Number(vf05nSalesTotal.toFixed(2));
    vf05nNetSalesTotal = Number(vf05nNetSalesTotal.toFixed(2));
    vf05nTaxTotal = Number(vf05nTaxTotal.toFixed(2));
    zsalesSalesTotal = Number(zsalesSalesTotal.toFixed(2));
    zsalesNetSalesTotal = Number(zsalesNetSalesTotal.toFixed(2));
    zsalesTaxTotal = Number(zsalesTaxTotal.toFixed(2));

    const finalReport: ImportValidationReport = {
      ...report,
      periodsDetected: (report.periodsDetected ?? []).map((p) => ({
        ...p,
        action: prevPeriodSet.has(p.period) ? 'replace' : 'add',
      })),
      incrementalSummary: {
        replacedPeriods,
        addedPeriods,
        retainedPeriods,
      },
    };

    const hasCatTargetsFlag = (prevBatch?.has_category_targets_sheet === 1) || staged.hasCategoryTargetsSheet ? 1 : 0;

    // Archive existing active batches without deleting them!
    db.prepare(`UPDATE import_batches SET status = 'archived' WHERE status = 'active'`).run();

    const batchRes = db
      .prepare(
        `INSERT INTO import_batches (
          batch_uuid, filename, imported_by_id, imported_by_name, status,
          vf05n_count, vf05n_sales_total, vf05n_net_sales_total, vf05n_tax_total,
          zsales_count, zsales_sales_total, zsales_net_sales_total, zsales_tax_total,
          unlinked_branch_count, unlinked_sector_count, has_category_targets_sheet, validation_report, created_at
        ) VALUES (?, ?, ?, ?, 'active', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
      )
      .run(
        batchUuid,
        staged.filename,
        user.id,
        user.fullName,
        vf05nCount,
        vf05nSalesTotal,
        vf05nNetSalesTotal,
        vf05nTaxTotal,
        zsalesCount,
        zsalesSalesTotal,
        zsalesNetSalesTotal,
        zsalesTaxTotal,
        unlinkedBranchCount,
        unlinkedSectorCount,
        hasCatTargetsFlag,
        JSON.stringify(finalReport),
        now
      );

    const batchId = Number(batchRes.lastInsertRowid);

    const insertSector = db.prepare(
      `INSERT INTO sectors (batch_id, sector_code, sector_name, monthly_target, is_original_sector) VALUES (?, ?, ?, ?, ?)`
    );
    for (const s of mergedSectors) {
      insertSector.run(batchId, s.sectorCode, s.sectorName, s.monthlyTarget, s.isOriginal);
    }

    const insertChain = db.prepare(
      `INSERT INTO departments_chains (batch_id, chain_code, chain_name, department_name, sector_name, supervisor_name) VALUES (?, ?, ?, ?, ?, ?)`
    );
    for (const c of mergedChains) {
      insertChain.run(batchId, c.chainCode, c.chainName, c.departmentName, c.sectorName, c.supervisorName);
    }

    const insertBranch = db.prepare(
      `INSERT INTO branches (batch_id, branch_code, branch_name, sector_name, branch_target, chain_code, chain_name, rep_code, rep_name, supervisor_name) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
    );
    for (const b of mergedBranches) {
      insertBranch.run(
        batchId,
        b.branchCode,
        b.branchName,
        b.sectorName,
        b.branchTarget,
        b.chainCode,
        b.chainName,
        b.repCode,
        b.repName,
        b.supervisorName
      );
    }

    const insertRep = db.prepare(
      `INSERT INTO sales_reps (batch_id, rep_code, rep_name, sector_name, supervisor_name, monthly_target) VALUES (?, ?, ?, ?, ?, ?)`
    );
    for (const r of mergedReps) {
      insertRep.run(batchId, r.repCode, r.repName, r.sectorName, r.supervisorName, r.monthlyTarget);
    }

    const insertItem = db.prepare(
      `INSERT INTO items (batch_id, item_code, item_name, category_name, brand_name, unit_name, chains_category, army_category, top_retail_category) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
    );
    for (const it of mergedItems) {
      insertItem.run(
        batchId,
        it.itemCode,
        it.itemName,
        it.categoryName,
        it.brandName,
        it.unitName,
        it.chainsCategory,
        it.armyCategory,
        it.topRetailCategory
      );
    }

    const insertCatTarget = db.prepare(
      `INSERT INTO category_targets (batch_id, sector_name, job_title, employee_code, employee_name, category_name, target_value) VALUES (?, ?, ?, ?, ?, ?, ?)`
    );
    for (const ct of mergedCatTargets) {
      insertCatTarget.run(batchId, ct.sectorName, ct.jobTitle, ct.employeeCode, ct.employeeName, ct.categoryName, ct.targetValue);
    }

    const insertSale = db.prepare(
      `INSERT INTO sales_records (
        batch_id, source_type, doc_number, billing_date,
        payer_code, payer_name, sold_to_code, sold_to_name,
        item_code, item_name, rep_code, quantity, sales_value, net_sales_value, tax_value,
        resolved_sector, resolved_department, resolved_chain_name,
        resolved_branch_code, resolved_branch_name, resolved_rep_code, resolved_rep_name, resolved_item_name,
        is_return
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
    );

    for (const rec of allSalesRecords) {
      insertSale.run(
        batchId,
        rec.sourceType,
        rec.docNumber,
        rec.billingDate,
        rec.payerCode,
        rec.payerName,
        rec.soldToCode,
        rec.soldToName,
        rec.itemCode,
        rec.itemName,
        rec.repCode,
        rec.quantity,
        rec.salesValue,
        rec.netSalesValue,
        rec.taxValue,
        rec.resolvedSector,
        rec.resolvedDepartment,
        rec.resolvedChainName,
        rec.resolvedBranchCode,
        rec.resolvedBranchName,
        rec.resolvedRepCode,
        rec.resolvedRepName,
        rec.resolvedItemName,
        rec.isReturn
      );
    }

    return {
      id: batchId,
      batchUuid,
      filename: staged.filename,
      importedByName: user.fullName,
      status: 'active' as const,
      vf05nCount,
      vf05nSalesTotal,
      vf05nNetSalesTotal,
      vf05nTaxTotal,
      zsalesCount,
      zsalesSalesTotal,
      zsalesNetSalesTotal,
      zsalesTaxTotal,
      unlinkedBranchCount,
      unlinkedSectorCount,
      createdAt: now,
    };
  });

  const batch = tx();
  stagedImports.delete(validationToken);
  return batch;
}

export function rollbackToBatch(db: Database.Database, batchId: number): ImportBatchHistoryItem {
  const target = db.prepare('SELECT * FROM import_batches WHERE id = ?').get(batchId) as
    | {
        id: number;
        batch_uuid: string;
        filename: string;
        imported_by_name: string;
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
      }
    | undefined;

  if (!target) {
    throw new Error('النسخة المطلوبة غير موجودة');
  }

  const tx = db.transaction(() => {
    db.prepare(`UPDATE import_batches SET status = 'archived' WHERE status = 'active'`).run();
    db.prepare(`UPDATE import_batches SET status = 'active' WHERE id = ?`).run(batchId);
  });
  tx();

  return {
    id: target.id,
    batchUuid: target.batch_uuid,
    filename: target.filename,
    importedByName: target.imported_by_name,
    status: 'active',
    vf05nCount: target.vf05n_count,
    vf05nSalesTotal: target.vf05n_sales_total,
    vf05nNetSalesTotal: target.vf05n_net_sales_total ?? 0,
    vf05nTaxTotal: target.vf05n_tax_total,
    zsalesCount: target.zsales_count,
    zsalesSalesTotal: target.zsales_sales_total,
    zsalesNetSalesTotal: target.zsales_net_sales_total ?? 0,
    zsalesTaxTotal: target.zsales_tax_total,
    unlinkedBranchCount: target.unlinked_branch_count,
    unlinkedSectorCount: target.unlinked_sector_count,
    createdAt: target.created_at,
  };
}

export function generateStandardTemplateWorkbook(): Uint8Array {
  const wb = XLSX.utils.book_new();

  const sheetsData: Record<string, string[][]> = {
    'VF05N (2)': [
      ['Billing Document', 'Billing Date', 'Payer', 'Name of Payer', 'Sold-to Party', 'Name of Sold-to Party', 'Material', 'Material Description', 'Sales Employee', 'Billed Quantity', 'New Growth Value BD&T', 'New Net Value BT', 'New Total Tax'],
    ],
    'Zsales_Price (2)': [
      ['Billing Document', 'Billing Dt.', 'Customer', 'Customer Name', 'Material', 'Material Description', 'Sales Emp.', 'Billed Quantity', 'New Growth Value BD&T', 'Net Value Without TA', 'Total Tax'],
    ],
    'القطاعات': [['كود القطاع', 'القطاع']],
    'اهداف القطاعات': [['القطاع', 'الهدف']],
    'تعريف المستخدمين': [['اسم المستخدم', 'الاسم الكامل', 'البريد الإلكتروني', 'الصلاحية', 'القطاعات']],
    'تعريف الفروع و الاهداف': [['كود الفرع', 'اسم الفرع', 'القطاع', 'كود الادارة', 'اسم الادارة', 'كود المندوب', 'اسم المندوب', 'اسم المشرف', 'إجمالى الهدف']],
    'تعريف الادارات والقطاعات': [['كود السلسلة', 'السلسلة', 'القطاع', 'اسم المشرف']],
    'تعريف مندوبى البيع': [['كود المندوب', 'اسم المندوب', 'القطاع', 'المشرف']],
    'تعريف الاصناف': [['كود الصنف', 'اسم الصنف', 'التصنيف A', 'تصنيف السلاسل', 'تصنيف الجيش', 'التوب ريتيل', 'العلامة التجارية', 'الوحدة']],
    'أهداف الفئات': [['القطاع', 'المسمى الوظيفى', 'كود الموظف', 'اسم الموظف', 'الفئة', 'الهدف']],
  };

  for (const [name, rows] of Object.entries(sheetsData)) {
    const ws = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, name);
  }

  return new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer);
}
