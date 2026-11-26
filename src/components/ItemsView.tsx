import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  ArrowUpDown,
  Check,
  CheckCircle2,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  Info,
  RotateCcw,
  Save,
  Search,
  User,
  X,
} from 'lucide-react';
import type { ItemCategoryAnalysisRow, SalesSource } from '../types/bi';
import { apiRequest } from '../utils/api';
import {
  buildActiveFilterMeta,
  buildExportFilename,
  exportToExcel,
  exportToPdf,
  type ExportColumn,
} from '../utils/exporters';
import { formatEGP, formatNum } from '../utils/format';
import { ExportButtons } from './ExportButtons';

interface ItemsViewProps {
  rows: ItemCategoryAnalysisRow[];
  effectiveSource: SalesSource;
  requestedSource: SalesSource;
  selectedSector: string;
  selectedJobTitle: string;
  startDate: string;
  endDate: string;
  filterResetSignal?: number;
  isAdmin: boolean;
  loading: boolean;
  onRefresh: () => void;
}

type SortKey = 'default' | 'sales' | 'tax' | 'target' | 'achievementPct';

function formatCategoryAchievement(pct: number | null | undefined): string {
  if (pct === null || pct === undefined || !Number.isFinite(pct)) {
    return '—';
  }
  return `${pct.toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}%`;
}

function normalizeArabicCompact(input: string): string {
  return String(input || '')
    .replace(/[\u00A0\u200B-\u200F\uFEFF]/g, ' ')
    .trim()
    .replace(/[\u064B-\u065F\u0670\u0640]/g, '')
    .replace(/[أإآ]/g, 'ا')
    .replace(/ى/g, 'ي')
    .replace(/ة/g, 'ه')
    .replace(/\s+/g, '')
    .toLowerCase();
}

function computeAchievementPct(sales: number, target: number, tax: number): number | null {
  const denom = target - tax;
  if (!Number.isFinite(denom) || denom <= 0) return null;
  const pct = (sales / denom) * 100;
  if (!Number.isFinite(pct)) return null;
  return Number(pct.toFixed(2));
}

export const ItemsView: React.FC<ItemsViewProps> = ({
  rows,
  effectiveSource,
  requestedSource,
  selectedSector,
  selectedJobTitle,
  startDate,
  endDate,
  filterResetSignal = 0,
  isAdmin,
  loading,
  onRefresh,
}) => {
  const [search, setSearch] = useState('');
  const [employeeFilter, setEmployeeFilter] = useState('ALL');
  const [empDropdownOpen, setEmpDropdownOpen] = useState(false);
  const [empSearch, setEmpSearch] = useState('');
  const empComboboxRef = useRef<HTMLDivElement>(null);
  const empSearchInputRef = useRef<HTMLInputElement>(null);

  const [categoryFilter, setCategoryFilter] = useState('ALL');
  const [sortKey, setSortKey] = useState<SortKey>('default');
  const [sortAsc, setSortAsc] = useState(false);

  const [drafts, setDrafts] = useState<Record<string, string>>({});
  const [saving, setSaving] = useState(false);
  const [saveMessage, setSaveMessage] = useState<string | null>(null);
  const [saveError, setSaveError] = useState<string | null>(null);

  useEffect(() => {
    setSaveMessage(null);
    setSaveError(null);
  }, [selectedSector, selectedJobTitle, effectiveSource]);

  // Reset cascading employee & category filters when parent Sector, Job Title, Source, or Reset button triggers
  useEffect(() => {
    setEmployeeFilter('ALL');
    setEmpSearch('');
    setEmpDropdownOpen(false);
    setCategoryFilter('ALL');
  }, [selectedSector, selectedJobTitle, effectiveSource, filterResetSignal]);

  // Close employee dropdown when clicking outside or pressing Escape
  useEffect(() => {
    if (!empDropdownOpen) return;
    const onMouseDown = (e: MouseEvent) => {
      if (empComboboxRef.current && !empComboboxRef.current.contains(e.target as Node)) {
        setEmpDropdownOpen(false);
      }
    };
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        setEmpDropdownOpen(false);
      }
    };
    document.addEventListener('mousedown', onMouseDown);
    document.addEventListener('keydown', onKeyDown);
    return () => {
      document.removeEventListener('mousedown', onMouseDown);
      document.removeEventListener('keydown', onKeyDown);
    };
  }, [empDropdownOpen]);

  useEffect(() => {
    if (empDropdownOpen) {
      setTimeout(() => empSearchInputRef.current?.focus(), 0);
    }
  }, [empDropdownOpen]);

  // Compute live target values (including live Central Supply waterfall rollup for رافت محمد)
  const liveRows = useMemo(() => {
    const parseDraftTarget = (row: ItemCategoryAnalysisRow): number => {
      if (row.isAutoTarget) return row.target;
      const raw = drafts[row.targetKey];
      if (raw === undefined) return row.target;
      const trimmed = raw.trim();
      if (!trimmed) return 0;
      const parsed = Number(trimmed);
      return Number.isFinite(parsed) && parsed >= 0 ? parsed : row.target;
    };

    const centralSubordinateRows = rows.filter(
      (r) => !r.isAutoTarget && normalizeArabicCompact(r.sectorName) === 'التوريدالمركزي'
    );
    const hasCentralSubordinates = centralSubordinateRows.length > 0;

    const centralWaterfallByCat = new Map<string, number>();
    if (hasCentralSubordinates) {
      const catBuckets = new Map<string, { secMgr: number; sup: number; rep: number }>();
      for (const cr of centralSubordinateRows) {
        const normCat = normalizeArabicCompact(cr.categoryName);
        const b = catBuckets.get(normCat) ?? { secMgr: 0, sup: 0, rep: 0 };
        const val = parseDraftTarget(cr);
        const normJob = normalizeArabicCompact(cr.jobTitle);
        if (normJob.includes('مندوب')) {
          b.rep += val;
        } else if (normJob.includes('مشرف')) {
          b.sup += val;
        } else {
          b.secMgr += val;
        }
        catBuckets.set(normCat, b);
      }
      for (const [normCat, b] of catBuckets.entries()) {
        const resolved = b.secMgr > 0 ? b.secMgr : b.sup > 0 ? b.sup : b.rep;
        centralWaterfallByCat.set(normCat, Number(resolved.toFixed(2)));
      }
    }

    return rows.map((r) => {
      const nextTarget =
        r.isAutoTarget && hasCentralSubordinates
          ? centralWaterfallByCat.get(normalizeArabicCompact(r.categoryName)) ?? 0
          : parseDraftTarget(r);
      return {
        ...r,
        target: Number(nextTarget.toFixed(2)),
        achievementPct: computeAchievementPct(r.sales, nextTarget, r.tax),
      };
    });
  }, [rows, drafts]);

  const dirtyRows = useMemo(() => {
    return rows.filter((r) => {
      if (r.isAutoTarget) return false;
      const raw = drafts[r.targetKey];
      if (raw === undefined) return false;
      const parsed = raw.trim() === '' ? 0 : Number(raw.trim());
      if (!Number.isFinite(parsed)) return true;
      return Number(parsed.toFixed(2)) !== Number(r.target.toFixed(2));
    });
  }, [rows, drafts]);

  // Cascading Employee / Rep options strictly derived from rows matching active Sector + Job Title
  const employeeOptions = useMemo(() => {
    const map = new Map<
      string,
      { key: string; name: string; code: string; jobTitle: string; sectorName: string }
    >();
    for (const r of liveRows) {
      const key = normalizeArabicCompact(r.employeeName);
      if (!key) continue;
      if (!map.has(key)) {
        map.set(key, {
          key,
          name: r.employeeName,
          code: r.employeeCode || '',
          jobTitle: r.jobTitle,
          sectorName: r.sectorName,
        });
      }
    }
    return Array.from(map.values()).sort((a, b) => a.name.localeCompare(b.name, 'ar'));
  }, [liveRows]);

  // If data refresh drops the selected employee from current scope, reset cleanly to ALL
  useEffect(() => {
    if (employeeFilter !== 'ALL' && !employeeOptions.some((o) => o.key === employeeFilter)) {
      setEmployeeFilter('ALL');
    }
  }, [employeeOptions, employeeFilter]);

  const filteredEmployeeOptions = useMemo(() => {
    const qNorm = normalizeArabicCompact(empSearch);
    const qRaw = empSearch.trim().toLowerCase();
    if (!qNorm && !qRaw) return employeeOptions;
    return employeeOptions.filter(
      (opt) =>
        normalizeArabicCompact(opt.name).includes(qNorm) ||
        (opt.code && opt.code.toLowerCase().includes(qRaw))
    );
  }, [employeeOptions, empSearch]);

  const selectedEmployeeOption = useMemo(
    () => (employeeFilter === 'ALL' ? null : employeeOptions.find((o) => o.key === employeeFilter) ?? null),
    [employeeOptions, employeeFilter]
  );

  const selectedEmployeeIndex = useMemo(
    () => (employeeFilter === 'ALL' ? -1 : employeeOptions.findIndex((o) => o.key === employeeFilter)),
    [employeeOptions, employeeFilter]
  );

  const handleStepEmployee = (delta: -1 | 1) => {
    if (employeeOptions.length === 0) return;
    if (selectedEmployeeIndex === -1) {
      const target = delta > 0 ? employeeOptions[0] : employeeOptions[employeeOptions.length - 1];
      setEmployeeFilter(target.key);
      return;
    }
    const nextIdx = (selectedEmployeeIndex + delta + employeeOptions.length) % employeeOptions.length;
    setEmployeeFilter(employeeOptions[nextIdx].key);
  };

  const categoryOptions = useMemo(() => {
    const set = new Set<string>();
    for (const r of liveRows) {
      if (employeeFilter !== 'ALL' && normalizeArabicCompact(r.employeeName) !== employeeFilter) {
        continue;
      }
      if (r.categoryName) set.add(r.categoryName);
    }
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'ar'));
  }, [liveRows, employeeFilter]);

  useEffect(() => {
    if (categoryFilter !== 'ALL' && !categoryOptions.includes(categoryFilter)) {
      setCategoryFilter('ALL');
    }
  }, [categoryOptions, categoryFilter]);

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    const list = liveRows.filter((r) => {
      if (employeeFilter !== 'ALL' && normalizeArabicCompact(r.employeeName) !== employeeFilter) {
        return false;
      }
      if (categoryFilter !== 'ALL' && r.categoryName !== categoryFilter) return false;
      if (!q) return true;
      return (
        r.sectorName.toLowerCase().includes(q) ||
        r.jobTitle.toLowerCase().includes(q) ||
        r.employeeName.toLowerCase().includes(q) ||
        r.employeeCode.toLowerCase().includes(q) ||
        r.categoryName.toLowerCase().includes(q)
      );
    });

    if (sortKey === 'default') return list;

    return [...list].sort((a, b) => {
      if (sortKey === 'achievementPct') {
        const av = a.achievementPct ?? -Infinity;
        const bv = b.achievementPct ?? -Infinity;
        return sortAsc ? av - bv : bv - av;
      }
      const diff = a[sortKey] - b[sortKey];
      return sortAsc ? diff : -diff;
    });
  }, [liveRows, employeeFilter, search, categoryFilter, sortKey, sortAsc]);

  const totals = useMemo(() => {
    const jobSet = new Set<string>();
    for (const r of filtered) {
      jobSet.add(normalizeArabicCompact(r.jobTitle));
    }

    if (jobSet.size <= 1) {
      let sales = 0;
      let tax = 0;
      let target = 0;
      for (const r of filtered) {
        sales += r.sales;
        tax += r.tax;
        target += r.target;
      }
      return {
        sales,
        tax,
        target,
        achievementPct: computeAchievementPct(sales, target, tax),
        isDeduplicated: false,
      };
    }

    // Multiple hierarchical job titles are displayed together (e.g., `جميع المسميات`):
    // Avoid double/triple-counting the same `(sector, category)` across `مندوب` / `مشرف` / `مدير قطاع` / `مدير إدارة`.
    const getTierRank = (jobTitle: string): number => {
      const j = normalizeArabicCompact(jobTitle);
      if (j.includes('مديرقطاع') || j.includes('قائد')) return 4;
      if (j.includes('مشرف')) return 3;
      if (j.includes('مندوب')) return 2;
      if (j.includes('مديرادارة') || j.includes('مديرالادارة') || j.includes('مديراداره') || j.includes('مديرالاداره')) return 1;
      return 2;
    };

    const singleSectorBuckets = new Map<
      string,
      Map<string, { rank: number; sales: number; tax: number; target: number }>
    >();
    const coveredSingleSectors = new Set<string>();
    const multiSectorRows: ItemCategoryAnalysisRow[] = [];

    for (const r of filtered) {
      if (r.sectorName.includes('+')) {
        multiSectorRows.push(r);
        continue;
      }
      const secNorm = normalizeArabicCompact(r.sectorName);
      const catNorm = normalizeArabicCompact(r.categoryName);
      const jobNorm = normalizeArabicCompact(r.jobTitle);
      const secCatKey = `${secNorm}::${catNorm}`;
      coveredSingleSectors.add(secNorm);

      let jobMap = singleSectorBuckets.get(secCatKey);
      if (!jobMap) {
        jobMap = new Map();
        singleSectorBuckets.set(secCatKey, jobMap);
      }
      const existing = jobMap.get(jobNorm) ?? {
        rank: getTierRank(r.jobTitle),
        sales: 0,
        tax: 0,
        target: 0,
      };
      existing.sales += r.sales;
      existing.tax += r.tax;
      existing.target += r.target;
      jobMap.set(jobNorm, existing);
    }

    let sales = 0;
    let tax = 0;
    let target = 0;

    for (const jobMap of singleSectorBuckets.values()) {
      let chosen: { rank: number; sales: number; tax: number; target: number } | null = null;
      for (const b of jobMap.values()) {
        if (!chosen || b.rank > chosen.rank) {
          chosen = b;
        }
      }
      if (chosen) {
        sales += chosen.sales;
        tax += chosen.tax;
        target += chosen.target;
      }
    }

    // Only add multi-sector rows (`التوريد المباشر + الجيش + التوب ريتيل`) if their sectors aren't already covered
    for (const r of multiSectorRows) {
      const constituentNorms = r.sectorName
        .split('+')
        .map((part) => normalizeArabicCompact(part))
        .filter(Boolean);
      const overlapsSingleSector = constituentNorms.some((s) => coveredSingleSectors.has(s));
      if (!overlapsSingleSector) {
        sales += r.sales;
        tax += r.tax;
        target += r.target;
      }
    }

    return {
      sales,
      tax,
      target,
      achievementPct: computeAchievementPct(sales, target, tax),
      isDeduplicated: true,
    };
  }, [filtered]);

  const toggleSort = (k: SortKey) => {
    if (sortKey === k) {
      setSortAsc(!sortAsc);
    } else {
      setSortKey(k);
      setSortAsc(false);
    }
  };

  const handleTargetChange = (targetKey: string, val: string) => {
    setSaveMessage(null);
    setSaveError(null);
    setDrafts((prev) => ({
      ...prev,
      [targetKey]: val,
    }));
  };

  const handleResetDrafts = () => {
    setDrafts({});
    setSaveError(null);
    setSaveMessage(null);
  };

  const handleSaveTargets = async () => {
    if (dirtyRows.length === 0) return;
    setSaveError(null);
    setSaveMessage(null);

    const payload: Array<{
      sectorName: string;
      jobTitle: string;
      employeeName: string;
      categoryName: string;
      targetValue: number;
    }> = [];

    for (const r of dirtyRows) {
      const raw = drafts[r.targetKey] ?? '';
      const num = raw.trim() === '' ? 0 : Number(raw.trim());
      if (!Number.isFinite(num) || num < 0) {
        setSaveError(`يرجى إدخال قيمة هدف صالحة (>= 0) لـ ${r.employeeName} — ${r.categoryName}`);
        return;
      }
      payload.push({
        sectorName: r.sectorName,
        jobTitle: r.jobTitle,
        employeeName: r.employeeName,
        categoryName: r.categoryName,
        targetValue: Number(num.toFixed(2)),
      });
    }

    setSaving(true);
    try {
      await apiRequest('/api/analytics/item-category-targets', {
        method: 'PUT',
        body: JSON.stringify({ targets: payload }),
      });
      setDrafts((prev) => {
        const next = { ...prev };
        for (const r of dirtyRows) {
          delete next[r.targetKey];
        }
        return next;
      });
      setSaveMessage(`تم حفظ ${formatNum(payload.length)} هدف بنجاح وتحديث النتائج.`);
      onRefresh();
    } catch (err) {
      setSaveError(err instanceof Error ? err.message : 'تعذر حفظ أهداف الفئات');
    } finally {
      setSaving(false);
    }
  };

  const getExportMeta = () => {
    const meta = buildActiveFilterMeta({
      source: effectiveSource,
      sector: selectedSector,
      startDate,
      endDate,
      extra: [
        selectedJobTitle === 'ALL'
          ? { label: 'المسمى الوظيفي', value: 'جميع المسميات الوظيفية' }
          : { label: 'المسمى الوظيفي', value: selectedJobTitle, fileTag: selectedJobTitle },
        selectedEmployeeOption
          ? { label: 'اسم الموظف', value: selectedEmployeeOption.name, fileTag: selectedEmployeeOption.name }
          : { label: 'اسم الموظف', value: `الكل (${formatNum(employeeOptions.length)})` },
        ...(categoryFilter !== 'ALL' ? [{ label: 'الفئة', value: categoryFilter, fileTag: categoryFilter }] : []),
        ...(search.trim() ? [{ label: 'بحث', value: search.trim(), fileTag: search.trim() }] : []),
      ],
    });

    const columns: ExportColumn[] = [
      { header: 'القطاع', key: 'sectorName', type: 'string', minWidth: 16 },
      { header: 'المسمى الوظيفى', key: 'jobTitle', type: 'string', minWidth: 15 },
      { header: 'اسم الموظف', key: 'employeeName', type: 'string', minWidth: 22 },
      { header: 'الفئة', key: 'categoryName', type: 'string', minWidth: 18 },
      { header: 'المبيعات', key: 'sales', type: 'currency', minWidth: 16 },
      { header: 'الضريبة', key: 'tax', type: 'currency', minWidth: 14 },
      { header: 'الهدف', key: 'target', type: 'currency', minWidth: 16 },
      { header: 'نسبة التحقيق', key: 'achievementPct', type: 'percent', minWidth: 14 },
    ];

    const exportRows = filtered.map((r) => ({
      sectorName: r.sectorName,
      jobTitle: r.jobTitle,
      employeeName: r.employeeName,
      categoryName: r.categoryName,
      sales: r.sales,
      tax: r.tax,
      target: r.target,
      achievementPct: r.achievementPct,
    }));

    const totalRow =
      filtered.length > 0
        ? {
            sectorName: totals.isDeduplicated ? 'الإجمالي غير المكرر للقطاع' : 'الإجمالي',
            jobTitle: '',
            employeeName: `${formatNum(filtered.length)} صف`,
            categoryName: '',
            sales: totals.sales,
            tax: totals.tax,
            target: totals.target,
            achievementPct: totals.achievementPct,
          }
        : null;

    return { meta, columns, exportRows, totalRow };
  };

  const handleExportExcel = () => {
    const { meta, columns, exportRows, totalRow } = getExportMeta();
    exportToExcel({
      filename: buildExportFilename('تحليل_فئات_الأصناف', meta.fileTags, 'xlsx'),
      sheets: [
        {
          name: 'تحليل فئات الأصناف',
          title: 'تحليل فئات الأصناف',
          subtitle: meta.summary.join(' | '),
          columns,
          rows: exportRows,
          totalRow,
        },
      ],
    });
  };

  const handleExportPdf = async () => {
    const { meta, columns, exportRows, totalRow } = getExportMeta();
    await exportToPdf({
      filename: buildExportFilename('تحليل_فئات_الأصناف', meta.fileTags, 'pdf'),
      tabTitle: 'تحليل فئات الأصناف',
      filtersSummary: meta.summary,
      orientation: 'landscape',
      tables: [
        {
          title: 'جدول تحليل فئات الأصناف والأهداف',
          subtitle: `عدد الصفوف الظاهرة: ${formatNum(filtered.length)}${totals.isDeduplicated ? ' · إجمالي غير مكرر للقطاع' : ''}`,
          columns,
          rows: exportRows,
          totalRow,
        },
      ],
    });
  };

  if (loading && rows.length === 0) {
    return <div className="h-64 bg-white border border-slate-200 rounded-xl animate-pulse" />;
  }

  return (
    <div className="space-y-4">
      {requestedSource === 'VF05N' && effectiveSource === 'ZSALES' && (
        <div className="bg-slate-50 border border-slate-200 rounded-xl px-4 py-3 text-xs text-slate-700 flex items-center gap-2.5">
          <Info className="w-4 h-4 text-teal-700 shrink-0" />
          <span>
            ورقة <span className="font-semibold">VF05N (2)</span> لا تحتوي على عمود أكواد الأصناف؛ يتم عرض تحليل فئات الأصناف تلقائيًا من ورقة{' '}
            <span className="font-semibold">Zsales_Price (2)</span> المطابقة في الإجماليات.
          </span>
        </div>
      )}

      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="p-5 border-b border-slate-200 flex flex-wrap items-center justify-between gap-4">
          <div>
            <h3 className="text-sm font-bold text-slate-900">تحليل فئات الأصناف</h3>
            <div className="flex flex-wrap items-center gap-2 text-xs text-slate-500 mt-1">
              <span>
                إجمالي السجلات: <span className="font-num font-semibold text-slate-700">{formatNum(filtered.length)}</span>
              </span>
              <span aria-hidden="true">·</span>
              <span>
                القطاع: <span className="font-semibold text-slate-700">{selectedSector === 'ALL' ? 'جميع القطاعات' : selectedSector}</span>
              </span>
              <span aria-hidden="true">·</span>
              <span>
                المسمى الوظيفي:{' '}
                <span className="font-semibold text-slate-700">
                  {selectedJobTitle === 'ALL' ? 'جميع المسميات' : selectedJobTitle}
                </span>
              </span>
              <span aria-hidden="true">·</span>
              <span>
                الموظف / المندوب:{' '}
                <span className="font-semibold text-slate-700">
                  {selectedEmployeeOption ? selectedEmployeeOption.name : `الكل (${formatNum(employeeOptions.length)})`}
                </span>
              </span>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2.5">
            {/* Cascading Searchable Employee / Rep Selector + One-by-One Stepper */}
            <div className="flex items-center gap-1.5">
              <div ref={empComboboxRef} className="relative">
                <button
                  type="button"
                  aria-label="اسم الموظف / المندوب"
                  aria-expanded={empDropdownOpen}
                  onClick={() => setEmpDropdownOpen((o) => !o)}
                  className={`inline-flex items-center justify-between gap-2 px-3 py-2 text-xs border rounded-lg bg-white transition-colors min-w-[210px] max-w-[270px] cursor-pointer focus:outline-none focus:ring-2 focus:ring-teal-600 ${
                    selectedEmployeeOption
                      ? 'border-teal-600 bg-teal-50/40 text-teal-900 font-semibold'
                      : 'border-slate-300 text-slate-700 hover:border-slate-400'
                  }`}
                >
                  <span className="inline-flex items-center gap-1.5 truncate">
                    <User className="w-3.5 h-3.5 text-slate-500 shrink-0" />
                    <span className="truncate">
                      {selectedEmployeeOption
                        ? selectedEmployeeOption.name
                        : `اسم الموظف / المندوب (${formatNum(employeeOptions.length)})`}
                    </span>
                  </span>
                  <ChevronDown className="w-3.5 h-3.5 text-slate-400 shrink-0" />
                </button>

                {empDropdownOpen && (
                  <div className="absolute right-0 mt-1.5 w-80 bg-white border border-slate-200 rounded-xl shadow-lg z-30 overflow-hidden">
                    <div className="p-2 border-b border-slate-100 bg-slate-50">
                      <div className="relative">
                        <Search className="w-3.5 h-3.5 text-slate-400 absolute right-2.5 top-1/2 -translate-y-1/2" />
                        <input
                          ref={empSearchInputRef}
                          type="text"
                          value={empSearch}
                          onChange={(e) => setEmpSearch(e.target.value)}
                          placeholder="بحث سريع بالاسم أو الكود..."
                          className="w-full pr-8 pl-7 py-1.5 text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-teal-600"
                        />
                        {empSearch && (
                          <button
                            type="button"
                            onClick={() => setEmpSearch('')}
                            className="absolute left-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 cursor-pointer"
                            aria-label="مسح البحث"
                          >
                            <X className="w-3.5 h-3.5" />
                          </button>
                        )}
                      </div>
                    </div>

                    <div className="max-h-64 overflow-y-auto divide-y divide-slate-100">
                      <button
                        type="button"
                        onClick={() => {
                          setEmployeeFilter('ALL');
                          setEmpDropdownOpen(false);
                        }}
                        className={`w-full px-3 py-2 text-right text-xs flex items-center justify-between gap-2 hover:bg-slate-50 cursor-pointer ${
                          employeeFilter === 'ALL' ? 'bg-teal-50/60 text-teal-800 font-semibold' : 'text-slate-700'
                        }`}
                      >
                        <span>الكل — جميع الموظفين ({formatNum(employeeOptions.length)})</span>
                        {employeeFilter === 'ALL' && <Check className="w-3.5 h-3.5 text-teal-700 shrink-0" />}
                      </button>

                      {filteredEmployeeOptions.length === 0 ? (
                        <div className="px-3 py-4 text-center text-xs text-slate-500">
                          لا يوجد موظف مطابق للبحث
                        </div>
                      ) : (
                        filteredEmployeeOptions.map((opt) => {
                          const isSelected = employeeFilter === opt.key;
                          return (
                            <button
                              key={opt.key}
                              type="button"
                              onClick={() => {
                                setEmployeeFilter(opt.key);
                                setEmpDropdownOpen(false);
                              }}
                              className={`w-full px-3 py-2 text-right text-xs flex items-center justify-between gap-2 hover:bg-slate-50 cursor-pointer ${
                                isSelected ? 'bg-teal-50/70 text-teal-900 font-semibold' : 'text-slate-800'
                              }`}
                            >
                              <div className="min-w-0">
                                <div className="truncate">{opt.name}</div>
                                {(opt.code || selectedJobTitle === 'ALL' || selectedSector === 'ALL') && (
                                  <div className="text-[11px] text-slate-500 truncate">
                                    {[
                                      opt.code ? `كود ${opt.code}` : '',
                                      selectedJobTitle === 'ALL' ? opt.jobTitle : '',
                                      selectedSector === 'ALL' ? opt.sectorName : '',
                                    ]
                                      .filter(Boolean)
                                      .join(' · ')}
                                  </div>
                                )}
                              </div>
                              {isSelected && <Check className="w-3.5 h-3.5 text-teal-700 shrink-0" />}
                            </button>
                          );
                        })
                      )}
                    </div>
                  </div>
                )}
              </div>

              {employeeOptions.length > 1 && (
                <div className="inline-flex items-center border border-slate-300 rounded-lg bg-white overflow-hidden">
                  <button
                    type="button"
                    onClick={() => handleStepEmployee(-1)}
                    title="الموظف السابق"
                    aria-label="الموظف السابق"
                    className="p-2 text-slate-600 hover:bg-slate-50 hover:text-slate-900 cursor-pointer border-l border-slate-200"
                  >
                    <ChevronRight className="w-3.5 h-3.5" />
                  </button>
                  {selectedEmployeeIndex >= 0 && (
                    <span
                      className="px-2 text-[11px] font-num tabular-nums text-slate-600 select-none"
                      dir="ltr"
                    >
                      {selectedEmployeeIndex + 1}/{employeeOptions.length}
                    </span>
                  )}
                  <button
                    type="button"
                    onClick={() => handleStepEmployee(1)}
                    title="الموظف التالي"
                    aria-label="الموظف التالي"
                    className="p-2 text-slate-600 hover:bg-slate-50 hover:text-slate-900 cursor-pointer border-r border-slate-200"
                  >
                    <ChevronLeft className="w-3.5 h-3.5" />
                  </button>
                </div>
              )}

              {selectedEmployeeOption && (
                <button
                  type="button"
                  onClick={() => setEmployeeFilter('ALL')}
                  className="px-2 py-1.5 text-xs font-medium text-slate-600 hover:text-slate-900 underline cursor-pointer whitespace-nowrap"
                >
                  عرض الكل
                </button>
              )}
            </div>

            {categoryOptions.length > 0 && (
              <select
                aria-label="الفئة"
                value={categoryFilter}
                onChange={(e) => setCategoryFilter(e.target.value)}
                className="px-3 py-2 text-xs border border-slate-300 rounded-lg bg-white text-slate-700 focus:outline-none focus:ring-2 focus:ring-teal-600"
              >
                <option value="ALL">جميع الفئات</option>
                {categoryOptions.map((cat) => (
                  <option key={cat} value={cat}>
                    {cat}
                  </option>
                ))}
              </select>
            )}
            <div className="relative">
              <Search className="w-4 h-4 text-slate-400 absolute right-3 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="بحث بالقطاع أو الموظف أو الفئة..."
                className="pr-9 pl-3 py-2 text-xs border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-teal-600 w-56"
              />
            </div>

            {isAdmin && (
              <div className="flex items-center gap-2">
                {dirtyRows.length > 0 && (
                  <button
                    type="button"
                    onClick={handleResetDrafts}
                    disabled={saving}
                    className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold border border-slate-300 rounded-lg text-slate-700 hover:bg-slate-50 cursor-pointer disabled:opacity-50"
                  >
                    <RotateCcw className="w-3.5 h-3.5" />
                    <span>تراجع</span>
                  </button>
                )}
                <button
                  type="button"
                  onClick={handleSaveTargets}
                  disabled={saving || dirtyRows.length === 0}
                  className={`inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold rounded-lg transition-colors whitespace-nowrap ${
                    dirtyRows.length > 0
                      ? 'bg-teal-700 hover:bg-teal-800 text-white cursor-pointer'
                      : 'bg-slate-100 text-slate-400 cursor-not-allowed'
                  }`}
                >
                  <Save className="w-3.5 h-3.5" />
                  <span>
                    {saving
                      ? 'جاري الحفظ...'
                      : dirtyRows.length > 0
                        ? `حفظ الأهداف (${formatNum(dirtyRows.length)})`
                        : 'حفظ الأهداف'}
                  </span>
                </button>
              </div>
            )}

            <ExportButtons
              onExportExcel={handleExportExcel}
              onExportPdf={handleExportPdf}
              disabled={loading || filtered.length === 0}
            />
          </div>
        </div>

        {(saveError || saveMessage) && (
          <div
            className={`px-5 py-2.5 text-xs border-b flex items-center gap-2 ${
              saveError
                ? 'bg-red-50 border-red-200 text-red-800'
                : 'bg-emerald-50 border-emerald-200 text-emerald-800'
            }`}
          >
            {saveMessage && <CheckCircle2 className="w-4 h-4 shrink-0 text-emerald-700" />}
            <span>{saveError ?? saveMessage}</span>
          </div>
        )}

        <div className="overflow-x-auto">
          <table className="w-full text-sm text-right">
            <thead className="bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-600">
              <tr>
                <th className="py-3 px-4 whitespace-nowrap">القطاع</th>
                <th className="py-3 px-4 whitespace-nowrap">المسمى الوظيفى</th>
                <th className="py-3 px-4 whitespace-nowrap">اسم الموظف</th>
                <th className="py-3 px-4 whitespace-nowrap">الفئة</th>
                <th className="py-3 px-4 whitespace-nowrap">
                  <button
                    onClick={() => toggleSort('sales')}
                    className="inline-flex items-center gap-1 cursor-pointer hover:text-slate-900"
                  >
                    <span>المبيعات</span>
                    <ArrowUpDown className="w-3.5 h-3.5" />
                  </button>
                </th>
                <th className="py-3 px-4 whitespace-nowrap">
                  <button
                    onClick={() => toggleSort('tax')}
                    className="inline-flex items-center gap-1 cursor-pointer hover:text-slate-900"
                  >
                    <span>الضريبة</span>
                    <ArrowUpDown className="w-3.5 h-3.5" />
                  </button>
                </th>
                <th className="py-3 px-4 whitespace-nowrap">
                  <button
                    onClick={() => toggleSort('target')}
                    className="inline-flex items-center gap-1 cursor-pointer hover:text-slate-900"
                  >
                    <span>الهدف</span>
                    <ArrowUpDown className="w-3.5 h-3.5" />
                  </button>
                </th>
                <th className="py-3 px-4 whitespace-nowrap">
                  <button
                    onClick={() => toggleSort('achievementPct')}
                    className="inline-flex items-center gap-1 cursor-pointer hover:text-slate-900"
                  >
                    <span>نسبة التحقيق</span>
                    <ArrowUpDown className="w-3.5 h-3.5" />
                  </button>
                </th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {filtered.length === 0 ? (
                <tr>
                  <td colSpan={8} className="py-10 text-center text-xs text-slate-500">
                    لا توجد بيانات فئات مطابقة للفلاتر المحددة
                  </td>
                </tr>
              ) : (
                filtered.map((r) => {
                  const draftVal = drafts[r.targetKey];
                  const isDirty = draftVal !== undefined && Number(draftVal || 0) !== r.target;
                  return (
                    <tr key={r.targetKey} className="hover:bg-slate-50/70">
                      <td className="py-2.5 px-4 text-xs font-medium text-slate-700 whitespace-nowrap">
                        {r.sectorName}
                      </td>
                      <td className="py-2.5 px-4 text-xs font-medium text-slate-700 whitespace-nowrap">
                        {r.jobTitle}
                      </td>
                      <td className="py-2.5 px-4 font-medium text-slate-900 whitespace-nowrap">
                        {r.employeeName}
                      </td>
                      <td className="py-2.5 px-4 text-slate-700 whitespace-nowrap">{r.categoryName}</td>
                      <td
                        className={`py-2.5 px-4 font-num tabular-nums font-semibold whitespace-nowrap ${
                          r.sales === 0 ? 'text-slate-400' : 'text-slate-900'
                        }`}
                        dir="ltr"
                      >
                        {formatEGP(r.sales)}
                      </td>
                      <td className="py-2.5 px-4 font-num tabular-nums text-slate-600 whitespace-nowrap" dir="ltr">
                        {formatEGP(r.tax)}
                      </td>
                      <td className="py-2 px-4 whitespace-nowrap">
                        {r.isAutoTarget ? (
                          <div className="flex flex-col items-end">
                            <span className="font-num tabular-nums font-semibold text-teal-800" dir="ltr">
                              {formatEGP(r.target)}
                            </span>
                            <span className="text-[11px] text-slate-500">تجميع تلقائي (التوريد المركزي)</span>
                          </div>
                        ) : isAdmin ? (
                          <input
                            type="number"
                            min="0"
                            step="0.01"
                            aria-label={`هدف ${r.employeeName} - ${r.categoryName}`}
                            value={draftVal !== undefined ? draftVal : r.target === 0 ? '' : String(r.target)}
                            placeholder="0"
                            onChange={(e) => handleTargetChange(r.targetKey, e.target.value)}
                            dir="ltr"
                            className={`w-32 px-2.5 py-1.5 text-xs font-num tabular-nums text-left border rounded-lg focus:outline-none focus:ring-2 focus:ring-teal-600 ${
                              isDirty
                                ? 'border-teal-600 bg-teal-50/40 text-slate-900 font-semibold'
                                : 'border-slate-300 bg-white text-slate-800'
                            }`}
                          />
                        ) : (
                          <span className="font-num tabular-nums text-slate-700" dir="ltr">
                            {formatEGP(r.target)}
                          </span>
                        )}
                      </td>
                      <td
                        className={`py-2.5 px-4 font-num tabular-nums font-semibold whitespace-nowrap ${
                          r.achievementPct === null
                            ? 'text-slate-400'
                            : r.achievementPct >= 100
                              ? 'text-emerald-700'
                              : 'text-teal-700'
                        }`}
                        dir="ltr"
                      >
                        {formatCategoryAchievement(r.achievementPct)}
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
            {filtered.length > 0 && (
              <tfoot className="bg-sky-50 border-t-2 border-slate-300 font-bold text-slate-900">
                <tr>
                  <td colSpan={4} className="py-2.5 px-4 whitespace-nowrap">
                    {totals.isDeduplicated
                      ? `الإجمالي غير المكرر للقطاع (${formatNum(filtered.length)} صف)`
                      : `الإجمالي (${formatNum(filtered.length)} صف)`}
                  </td>
                  <td className="py-2.5 px-4 font-num tabular-nums whitespace-nowrap" dir="ltr">
                    {formatEGP(totals.sales)}
                  </td>
                  <td className="py-2.5 px-4 font-num tabular-nums whitespace-nowrap" dir="ltr">
                    {formatEGP(totals.tax)}
                  </td>
                  <td className="py-2.5 px-4 font-num tabular-nums whitespace-nowrap" dir="ltr">
                    {formatEGP(totals.target)}
                  </td>
                  <td className="py-2.5 px-4 font-num tabular-nums text-teal-900 whitespace-nowrap" dir="ltr">
                    {formatCategoryAchievement(totals.achievementPct)}
                  </td>
                </tr>
              </tfoot>
            )}
          </table>
        </div>
      </div>
    </div>
  );
};
