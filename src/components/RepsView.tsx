import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { ArrowUpDown, CheckCircle2, Cloud, Lock, Save, Search } from 'lucide-react';
import type { RepMetricRow, RepTargetMode, SalesSource } from '../types/bi';
import { apiRequest } from '../utils/api';
import {
  buildActiveFilterMeta,
  buildExportFilename,
  exportToExcel,
  exportToPdf,
} from '../utils/exporters';
import { formatEGP, formatNum } from '../utils/format';
import { ExportButtons } from './ExportButtons';

interface RepsViewProps {
  reps: RepMetricRow[];
  targetMode: RepTargetMode;
  source: SalesSource;
  sectorFilter: string;
  startDate: string;
  endDate: string;
  isAdmin: boolean;
  canEditTargets: boolean;
  authorizedSectors: string[];
  loading: boolean;
  onTargetModeChange: (nextMode: RepTargetMode) => void;
  onRefresh: () => void;
}

type SortKey = 'sales' | 'tax' | 'target' | 'achievementPct' | 'recordCount' | 'branchCount';

function formatRepAchievement(pct: number | null | undefined): string {
  if (pct === null || pct === undefined || !Number.isFinite(pct)) {
    return '—';
  }
  return `${pct.toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}%`;
}

function computeRepAchievement(sales: number, target: number, tax: number): number | null {
  const denom = target - tax;
  if (!Number.isFinite(denom) || denom <= 0) return null;
  const pct = (sales / denom) * 100;
  if (!Number.isFinite(pct)) return null;
  return Number(pct.toFixed(2));
}

export const RepsView: React.FC<RepsViewProps> = ({
  reps,
  targetMode,
  source,
  sectorFilter,
  startDate,
  endDate,
  isAdmin,
  canEditTargets,
  authorizedSectors,
  loading,
  onTargetModeChange,
  onRefresh,
}) => {
  const [search, setSearch] = useState('');
  const [sortKey, setSortKey] = useState<SortKey>('sales');
  const [sortAsc, setSortAsc] = useState(false);

  const [drafts, setDrafts] = useState<Record<string, string>>({});
  const [savedManualOverrides, setSavedManualOverrides] = useState<Record<string, number>>({});
  const [savingTargets, setSavingTargets] = useState(false);
  const [savingMode, setSavingMode] = useState(false);
  const [saveMessage, setSaveMessage] = useState<string | null>(null);
  const [saveError, setSaveError] = useState<string | null>(null);

  const debounceTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const draftsRef = useRef<Record<string, string>>({});
  draftsRef.current = drafts;

  useEffect(() => {
    // Sync server manual values when reps reload & clear stale drafts that match persisted value
    setSavedManualOverrides((prev) => {
      const next = { ...prev };
      for (const r of reps) {
        if (r.manualTarget !== null && r.manualTarget !== undefined) {
          next[r.repCode] = r.manualTarget;
        }
      }
      return next;
    });
    setDrafts((prev) => {
      const next: Record<string, string> = {};
      for (const [code, raw] of Object.entries(prev)) {
        const rep = reps.find((r) => r.repCode === code);
        if (!rep) continue;
        const parsed = raw.trim() === '' ? 0 : Number(raw.trim());
        const savedVal = rep.manualTarget ?? 0;
        if (Number.isFinite(parsed) && Number(parsed.toFixed(2)) === Number(savedVal.toFixed(2))) {
          continue;
        }
        next[code] = raw;
      }
      return next;
    });
  }, [reps]);

  useEffect(() => {
    return () => {
      if (debounceTimerRef.current) {
        clearTimeout(debounceTimerRef.current);
      }
    };
  }, []);

  const canEditRepRow = useCallback(
    (row: RepMetricRow): boolean => {
      if (!canEditTargets) return false;
      if (!row.repCode || row.repCode === 'غير محدد') return false;
      if (isAdmin) return true;
      return Boolean(row.sectorName && authorizedSectors.includes(row.sectorName));
    },
    [canEditTargets, isAdmin, authorizedSectors]
  );

  const getPersistedManualValue = useCallback(
    (row: RepMetricRow): number | null => {
      if (Object.prototype.hasOwnProperty.call(savedManualOverrides, row.repCode)) {
        return savedManualOverrides[row.repCode];
      }
      return row.manualTarget ?? null;
    },
    [savedManualOverrides]
  );

  const liveReps = useMemo(() => {
    return reps.map((r) => {
      const isUnlinked = r.repCode === 'غير محدد';
      if (isUnlinked) {
        return {
          ...r,
          target: 0,
          achievementPct: null,
        };
      }

      const persistedManual = getPersistedManualValue(r);
      let nextTarget = r.autoTarget;

      if (targetMode === 'manual') {
        const draftRaw = drafts[r.repCode];
        if (draftRaw !== undefined && canEditRepRow(r)) {
          const trimmed = draftRaw.trim();
          if (trimmed === '') {
            nextTarget = 0;
          } else {
            const parsed = Number(trimmed);
            nextTarget =
              Number.isFinite(parsed) && parsed >= 0
                ? Number(parsed.toFixed(2))
                : persistedManual ?? 0;
          }
        } else {
          nextTarget = persistedManual ?? 0;
        }
      }

      const cleanTarget = Number(nextTarget.toFixed(2));
      return {
        ...r,
        manualTarget: persistedManual,
        target: cleanTarget,
        achievementPct: computeRepAchievement(r.sales, cleanTarget, r.tax),
      };
    });
  }, [reps, targetMode, drafts, canEditRepRow, getPersistedManualValue]);

  const dirtyEntries = useMemo(() => {
    if (!canEditTargets) return [];
    const list: Array<{ rep: RepMetricRow; targetValue: number }> = [];
    for (const r of reps) {
      if (!canEditRepRow(r)) continue;
      const raw = drafts[r.repCode];
      if (raw === undefined) continue;
      const trimmed = raw.trim();
      const parsed = trimmed === '' ? 0 : Number(trimmed);
      if (!Number.isFinite(parsed) || parsed < 0) continue;
      const cleanVal = Number(parsed.toFixed(2));
      const existingManual = getPersistedManualValue(r);
      if (existingManual === null || Number(existingManual.toFixed(2)) !== cleanVal) {
        list.push({ rep: r, targetValue: cleanVal });
      }
    }
    return list;
  }, [reps, drafts, canEditTargets, canEditRepRow, getPersistedManualValue]);

  const flushSaveDrafts = useCallback(
    async (targetCodes?: string[]) => {
      if (!canEditTargets) return true;
      if (debounceTimerRef.current) {
        clearTimeout(debounceTimerRef.current);
        debounceTimerRef.current = null;
      }

      const currentDrafts = draftsRef.current;
      const payload: Array<{ repCode: string; repName: string; sectorName: string; targetValue: number }> = [];

      for (const r of reps) {
        if (!canEditRepRow(r)) continue;
        if (targetCodes && !targetCodes.includes(r.repCode)) continue;
        const raw = currentDrafts[r.repCode];
        if (raw === undefined) continue;
        const trimmed = raw.trim();
        const parsed = trimmed === '' ? 0 : Number(trimmed);
        if (!Number.isFinite(parsed) || parsed < 0) {
          setSaveError(`يرجى إدخال رقم صحيح غير سالب لهدف المندوب ${r.repName}`);
          return false;
        }
        const cleanVal = Number(parsed.toFixed(2));
        const existingManual = getPersistedManualValue(r);
        if (existingManual !== null && Number(existingManual.toFixed(2)) === cleanVal) {
          continue;
        }
        payload.push({
          repCode: r.repCode,
          repName: r.repName,
          sectorName: r.sectorName,
          targetValue: cleanVal,
        });
      }

      if (payload.length === 0) return true;

      setSavingTargets(true);
      setSaveError(null);
      try {
        await apiRequest('/api/analytics/reps/manual-targets', {
          method: 'PUT',
          body: JSON.stringify({ targets: payload }),
        });

        setSavedManualOverrides((prev) => {
          const next = { ...prev };
          for (const item of payload) {
            next[item.repCode] = item.targetValue;
          }
          return next;
        });

        setDrafts((prev) => {
          const next = { ...prev };
          for (const item of payload) {
            const curRaw = next[item.repCode];
            if (curRaw !== undefined) {
              const curNum = curRaw.trim() === '' ? 0 : Number(curRaw.trim());
              if (Number.isFinite(curNum) && Number(curNum.toFixed(2)) === item.targetValue) {
                delete next[item.repCode];
              }
            }
          }
          return next;
        });

        setSaveMessage(
          payload.length === 1
            ? `تم حفظ وتحديث هدف المندوب (${payload[0].repName}) في قاعدة البيانات.`
            : `تم حفظ وتحديث أهداف ${formatNum(payload.length)} مندوب في قاعدة البيانات.`
        );
        onRefresh();
        return true;
      } catch (err) {
        setSaveError(err instanceof Error ? err.message : 'تعذر حفظ الأهداف اليدوية في قاعدة البيانات');
        return false;
      } finally {
        setSavingTargets(false);
      }
    },
    [canEditTargets, reps, canEditRepRow, getPersistedManualValue, onRefresh]
  );

  const handleManualDraftChange = (repCode: string, nextValue: string) => {
    setSaveError(null);
    setSaveMessage(null);
    setDrafts((prev) => ({
      ...prev,
      [repCode]: nextValue,
    }));

    if (debounceTimerRef.current) {
      clearTimeout(debounceTimerRef.current);
    }
    debounceTimerRef.current = setTimeout(() => {
      flushSaveDrafts();
    }, 650);
  };

  const handleSwitchMode = async (nextMode: RepTargetMode) => {
    if (nextMode === targetMode) return;
    setSaveError(null);

    if (targetMode === 'manual' && nextMode === 'auto' && canEditTargets) {
      await flushSaveDrafts();
    }

    onTargetModeChange(nextMode);

    if (canEditTargets) {
      setSavingMode(true);
      try {
        await apiRequest('/api/analytics/reps/target-mode', {
          method: 'PUT',
          body: JSON.stringify({ targetMode: nextMode }),
        });
        setSaveMessage(
          nextMode === 'auto'
            ? 'تم التبديل إلى الوضع الأوتوماتيك (تجميع أهداف الفروع تلقائياً مع الاحتفاظ بالقيم اليدوية المحفوظة).'
            : 'تم التبديل إلى الوضع اليدوي وعرض آخر أهداف يدوية محفوظة في قاعدة البيانات.'
        );
        onRefresh();
      } catch (err) {
        setSaveError(err instanceof Error ? err.message : 'تعذر حفظ وضع الأهداف في قاعدة البيانات');
      } finally {
        setSavingMode(false);
      }
    }
  };

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    return liveReps
      .filter((r) => {
        if (!q) return true;
        return (
          r.repCode.toLowerCase().includes(q) ||
          r.repName.toLowerCase().includes(q) ||
          r.supervisorName.toLowerCase().includes(q) ||
          r.sectorName.toLowerCase().includes(q)
        );
      })
      .sort((a, b) => {
        if (sortKey === 'achievementPct') {
          const av = a.achievementPct ?? -Infinity;
          const bv = b.achievementPct ?? -Infinity;
          return sortAsc ? av - bv : bv - av;
        }
        const diff = a[sortKey] - b[sortKey];
        return sortAsc ? diff : -diff;
      });
  }, [liveReps, search, sortKey, sortAsc]);

  const totals = useMemo(() => {
    let sales = 0;
    let tax = 0;
    let target = 0;
    let recordCount = 0;
    let returnCount = 0;
    let branchCount = 0;
    for (const r of filtered) {
      sales += r.sales;
      tax += r.tax;
      target += r.target;
      recordCount += r.recordCount;
      returnCount += r.returnCount;
      branchCount += Math.trunc(r.branchCount || 0);
    }
    const denom = target - tax;
    const achievementPct = denom > 0 ? (sales / denom) * 100 : null;
    return {
      sales,
      tax,
      target,
      achievementPct:
        achievementPct !== null && Number.isFinite(achievementPct)
          ? Number(achievementPct.toFixed(2))
          : null,
      recordCount,
      returnCount,
      branchCount,
    };
  }, [filtered]);

  const toggleSort = (k: SortKey) => {
    if (sortKey === k) setSortAsc(!sortAsc);
    else {
      setSortKey(k);
      setSortAsc(false);
    }
  };

  const modeLabel = targetMode === 'manual' ? 'الوضع اليدوي (أهداف محفوظة)' : 'الوضع الأوتوماتيك (تجميع الفروع)';

  const filterMeta = buildActiveFilterMeta({
    source,
    sector: sectorFilter,
    startDate,
    endDate,
    extra: [
      { label: 'نظام الهدف', value: modeLabel },
      ...(search.trim() ? [{ label: 'بحث', value: search.trim(), fileTag: 'مفلتر' }] : []),
    ],
  });

  const handleExportExcel = () => {
    const filename = buildExportFilename('جهاز_البيع', filterMeta.fileTags, 'xlsx');
    exportToExcel({
      filename,
      sheets: [
        {
          name: 'جهاز البيع',
          title: 'تحليلات قطاع السلاسل — جدول أداء مندوبي المبيعات',
          subtitle: filterMeta.summary.join(' | '),
          columns: [
            { header: 'كود المندوب', key: 'repCode', type: 'string', minWidth: 14 },
            { header: 'اسم المندوب', key: 'repName', type: 'string', minWidth: 24 },
            { header: 'القطاع', key: 'sectorName', type: 'string', minWidth: 18 },
            { header: 'المشرف', key: 'supervisorName', type: 'string', minWidth: 18 },
            { header: 'المبيعات', key: 'sales', type: 'currency', minWidth: 16 },
            { header: 'الضريبة', key: 'tax', type: 'currency', minWidth: 15 },
            { header: 'الهدف', key: 'target', type: 'currency', minWidth: 16 },
            { header: 'نسبة التحقيق', key: 'achievementPct', type: 'percent', minWidth: 14 },
            { header: 'الفروع', key: 'branchCount', type: 'integer', minWidth: 10 },
            { header: 'عدد السجلات', key: 'recordCount', type: 'integer', minWidth: 12 },
            { header: 'عدد المرتجعات', key: 'returnCount', type: 'integer', minWidth: 12 },
          ],
          rows: filtered.map((r) => ({
            repCode: r.repCode,
            repName: r.repName,
            sectorName: r.sectorName,
            supervisorName: r.supervisorName,
            sales: r.sales,
            tax: r.tax,
            target: r.target,
            achievementPct: r.achievementPct,
            branchCount: Math.trunc(r.branchCount || 0),
            recordCount: r.recordCount,
            returnCount: r.returnCount,
          })),
          totalRow:
            filtered.length > 0
              ? {
                  repCode: `الإجمالي (${formatNum(filtered.length)} مندوب)`,
                  repName: '',
                  sectorName: '',
                  supervisorName: '',
                  sales: totals.sales,
                  tax: totals.tax,
                  target: totals.target,
                  achievementPct: totals.achievementPct,
                  branchCount: Math.trunc(totals.branchCount),
                  recordCount: totals.recordCount,
                  returnCount: totals.returnCount,
                }
              : null,
        },
      ],
    });
  };

  const handleExportPdf = async () => {
    const filename = buildExportFilename('جهاز_البيع', filterMeta.fileTags, 'pdf');
    await exportToPdf({
      filename,
      tabTitle: 'جهاز البيع — أداء مندوبي المبيعات',
      filtersSummary: filterMeta.summary,
      orientation: 'landscape',
      tables: [
        {
          title: 'جدول أداء مندوبي المبيعات',
          subtitle: `إجمالي المناديب الظاهرين: ${formatNum(filtered.length)} · ${modeLabel}`,
          columns: [
            { header: 'الكود', key: 'repCode', type: 'string', align: 'center' },
            { header: 'اسم المندوب', key: 'repName', type: 'string', align: 'right' },
            { header: 'القطاع', key: 'sectorName', type: 'string', align: 'right' },
            { header: 'المشرف', key: 'supervisorName', type: 'string', align: 'right' },
            { header: 'المبيعات', key: 'sales', type: 'currency' },
            { header: 'الضريبة', key: 'tax', type: 'currency' },
            { header: 'الهدف', key: 'target', type: 'currency' },
            { header: 'نسبة التحقيق', key: 'achievementPct', type: 'percent' },
            { header: 'الفروع', key: 'branchCount', type: 'integer' },
            { header: 'السجلات / المرتجعات', key: 'recordsSummary', type: 'string', align: 'center' },
          ],
          rows: filtered.map((r) => ({
            repCode: r.repCode,
            repName: r.repName,
            sectorName: r.sectorName,
            supervisorName: r.supervisorName,
            sales: r.sales,
            tax: r.tax,
            target: r.target,
            achievementPct: r.achievementPct,
            branchCount: Math.trunc(r.branchCount || 0),
            recordsSummary: `${formatNum(r.recordCount)} · ${formatNum(r.returnCount)} مرتجع`,
          })),
          totalRow:
            filtered.length > 0
              ? {
                  repCode: 'الإجمالي',
                  repName: `${formatNum(filtered.length)} مندوب`,
                  sectorName: '',
                  supervisorName: '',
                  sales: totals.sales,
                  tax: totals.tax,
                  target: totals.target,
                  achievementPct: totals.achievementPct,
                  branchCount: Math.trunc(totals.branchCount),
                  recordsSummary: `${formatNum(totals.recordCount)} · ${formatNum(totals.returnCount)} مرتجع`,
                }
              : null,
        },
      ],
    });
  };

  if (loading && reps.length === 0) {
    return <div className="h-64 bg-white border border-slate-200 rounded-xl animate-pulse" />;
  }

  return (
    <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
      <div className="p-5 border-b border-slate-200 space-y-4">
        <div className="flex flex-wrap items-center justify-between gap-4">
          <div>
            <h3 className="text-sm font-bold text-slate-900">جدول أداء مندوبي المبيعات</h3>
            <p className="text-xs text-slate-500 mt-0.5">
              إجمالي المناديب: <span className="font-num">{formatNum(filtered.length)}</span>
              <span className="mx-2 text-slate-300">|</span>
              <span>
                {targetMode === 'auto'
                  ? 'الوضع الأوتوماتيك: هدف كل مندوب محسوب تلقائياً من تجميع أهداف الفروع التابعة له'
                  : 'الوضع اليدوي: عرض وتعديل الأهداف اليدوية المحفوظة على آخر إدخال في قاعدة البيانات'}
              </span>
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-3">
            {/* Target Mode Switch: Auto vs Manual */}
            <div className="inline-flex items-center gap-2.5 px-3 py-1.5 bg-slate-50 border border-slate-200 rounded-lg">
              <span className="text-xs font-semibold text-slate-700 whitespace-nowrap">نظام الأهداف:</span>
              <button
                type="button"
                onClick={() => handleSwitchMode('auto')}
                disabled={savingMode}
                className={`text-xs font-bold transition-colors cursor-pointer whitespace-nowrap ${
                  targetMode === 'auto' ? 'text-teal-800' : 'text-slate-400 hover:text-slate-700'
                }`}
              >
                أوتوماتيك
              </button>

              <button
                type="button"
                role="switch"
                aria-checked={targetMode === 'manual'}
                aria-label="التبديل بين الوضع الأوتوماتيك والوضع اليدوي لأهداف المناديب"
                disabled={savingMode}
                onClick={() => handleSwitchMode(targetMode === 'auto' ? 'manual' : 'auto')}
                dir="ltr"
                className={`relative inline-flex h-6 w-11 shrink-0 items-center rounded-full transition-colors cursor-pointer focus:outline-none focus:ring-2 focus:ring-teal-600 focus:ring-offset-1 ${
                  targetMode === 'manual' ? 'bg-amber-600' : 'bg-teal-700'
                } ${savingMode ? 'opacity-60' : ''}`}
              >
                <span
                  className={`inline-block h-4 w-4 transform rounded-full bg-white shadow transition-transform ${
                    targetMode === 'manual' ? 'translate-x-6' : 'translate-x-1'
                  }`}
                />
              </button>

              <button
                type="button"
                onClick={() => handleSwitchMode('manual')}
                disabled={savingMode}
                className={`text-xs font-bold transition-colors cursor-pointer whitespace-nowrap ${
                  targetMode === 'manual' ? 'text-amber-800' : 'text-slate-400 hover:text-slate-700'
                }`}
              >
                يدوي
              </button>
            </div>

            {targetMode === 'manual' && canEditTargets && dirtyEntries.length > 0 && (
              <button
                type="button"
                onClick={() => flushSaveDrafts()}
                disabled={savingTargets}
                className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-white bg-teal-700 hover:bg-teal-800 rounded-lg transition-colors cursor-pointer disabled:opacity-50"
              >
                <Save className="w-3.5 h-3.5" />
                <span>{savingTargets ? 'جاري الحفظ...' : `حفظ التعديلات (${formatNum(dirtyEntries.length)})`}</span>
              </button>
            )}

            <div className="relative">
              <Search className="w-4 h-4 text-slate-400 absolute right-3 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="بحث بكود المندوب أو الاسم..."
                className="pr-9 pl-3 py-2 text-xs border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-teal-600 w-64"
              />
            </div>

            <ExportButtons
              onExportExcel={handleExportExcel}
              onExportPdf={handleExportPdf}
              disabled={filtered.length === 0}
            />
          </div>
        </div>

        {/* Status / Info Strip */}
        {(saveMessage || saveError || savingTargets || targetMode === 'manual') && (
          <div className="flex flex-wrap items-center justify-between gap-2 text-xs">
            <div className="flex flex-wrap items-center gap-3">
              {targetMode === 'manual' && (
                <div className="inline-flex items-center gap-1.5 text-slate-600">
                  {canEditTargets ? (
                    <>
                      <Cloud className="w-3.5 h-3.5 text-teal-700 shrink-0" />
                      <span>
                        أدخل أو عدّل الهدف لأي مندوب وسيتم الحفظ والتحديث تلقائياً (Upsert) على آخر إدخال، مع الاحتفاظ بالقيم عند التبديل إلى الأوتوماتيك.
                      </span>
                    </>
                  ) : (
                    <>
                      <Lock className="w-3.5 h-3.5 text-slate-500 shrink-0" />
                      <span>عرض فقط للأهداف اليدوية المحفوظة (التعديل متاح للأدمن ومدير القطاع فقط).</span>
                    </>
                  )}
                </div>
              )}
            </div>

            <div className="flex items-center gap-2">
              {savingTargets && (
                <span className="text-teal-700 font-semibold">جاري حفظ التحديث في قاعدة البيانات...</span>
              )}
              {!savingTargets && saveMessage && !saveError && (
                <span className="inline-flex items-center gap-1 text-emerald-700 font-semibold">
                  <CheckCircle2 className="w-3.5 h-3.5 shrink-0" />
                  <span>{saveMessage}</span>
                </span>
              )}
              {saveError && (
                <span className="text-rose-700 font-semibold bg-rose-50 border border-rose-200 px-2.5 py-1 rounded-md">
                  {saveError}
                </span>
              )}
            </div>
          </div>
        )}
      </div>

      <div className="overflow-x-auto">
        <table className="w-full text-sm text-right">
          <thead className="bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-600">
            <tr>
              <th className="py-3 px-4 whitespace-nowrap">كود المندوب</th>
              <th className="py-3 px-4 whitespace-nowrap">اسم المندوب</th>
              <th className="py-3 px-4 whitespace-nowrap">القطاع</th>
              <th className="py-3 px-4 whitespace-nowrap">المشرف</th>
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
                  <span>{targetMode === 'manual' ? 'الهدف (يدوي)' : 'الهدف (أوتوماتيك)'}</span>
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
              <th className="py-3 px-4 whitespace-nowrap">
                <button
                  onClick={() => toggleSort('branchCount')}
                  className="inline-flex items-center gap-1 cursor-pointer hover:text-slate-900"
                >
                  <span>الفروع</span>
                  <ArrowUpDown className="w-3.5 h-3.5" />
                </button>
              </th>
              <th className="py-3 px-4 whitespace-nowrap">
                <button
                  onClick={() => toggleSort('recordCount')}
                  className="inline-flex items-center gap-1 cursor-pointer hover:text-slate-900"
                >
                  <span>السجلات / المرتجعات</span>
                  <ArrowUpDown className="w-3.5 h-3.5" />
                </button>
              </th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {filtered.length === 0 ? (
              <tr>
                <td colSpan={10} className="py-8 text-center text-xs text-slate-500">
                  لا توجد بيانات مناديب مطابقة
                </td>
              </tr>
            ) : (
              filtered.map((r) => {
                const editable = targetMode === 'manual' && canEditRepRow(r);
                const persistedManual = getPersistedManualValue(r);
                const inputVal =
                  drafts[r.repCode] !== undefined
                    ? drafts[r.repCode]
                    : persistedManual !== null
                      ? String(persistedManual)
                      : '';
                const isDirty = dirtyEntries.some((d) => d.rep.repCode === r.repCode);

                return (
                  <tr key={r.repCode} className="hover:bg-slate-50/70">
                    <td className="py-2.5 px-4 font-num text-xs font-medium text-slate-700 whitespace-nowrap">
                      {r.repCode}
                    </td>
                    <td className="py-2.5 px-4 font-medium text-slate-900 whitespace-nowrap">{r.repName}</td>
                    <td className="py-2.5 px-4 text-xs text-slate-600 whitespace-nowrap">{r.sectorName}</td>
                    <td className="py-2.5 px-4 text-xs text-slate-600 whitespace-nowrap">{r.supervisorName}</td>
                    <td
                      className="py-2.5 px-4 font-num tabular-nums font-semibold text-slate-900 whitespace-nowrap"
                      dir="ltr"
                    >
                      {formatEGP(r.sales)}
                    </td>
                    <td className="py-2.5 px-4 font-num tabular-nums text-slate-600 whitespace-nowrap" dir="ltr">
                      {formatEGP(r.tax)}
                    </td>
                    <td className="py-2 px-4 font-num tabular-nums text-slate-700 whitespace-nowrap">
                      {editable ? (
                        <div className="space-y-0.5 min-w-[140px]">
                          <input
                            type="number"
                            min="0"
                            step="0.01"
                            dir="ltr"
                            aria-label={`الهدف اليدوي للمندوب ${r.repName}`}
                            value={inputVal}
                            placeholder="0"
                            onChange={(e) => handleManualDraftChange(r.repCode, e.target.value)}
                            onBlur={() => {
                              if (draftsRef.current[r.repCode] !== undefined) {
                                flushSaveDrafts([r.repCode]);
                              }
                            }}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') {
                                (e.currentTarget as HTMLInputElement).blur();
                              }
                            }}
                            className={`w-36 px-2.5 py-1 text-xs font-num tabular-nums text-left border rounded-md focus:outline-none focus:ring-2 focus:ring-teal-600 ${
                              isDirty
                                ? 'border-amber-400 bg-amber-50/50 text-slate-900'
                                : 'border-slate-300 bg-white text-slate-900'
                            }`}
                          />
                          <div className="text-[10px] text-slate-400 whitespace-nowrap">
                            الأوتوماتيك: <span dir="ltr">{formatEGP(r.autoTarget)}</span>
                          </div>
                        </div>
                      ) : (
                        <div dir="ltr">
                          <span>{formatEGP(r.target)}</span>
                          {targetMode === 'manual' && r.repCode !== 'غير محدد' && (
                            <div className="text-[10px] text-slate-400" dir="rtl">
                              الأوتوماتيك: <span dir="ltr">{formatEGP(r.autoTarget)}</span>
                            </div>
                          )}
                        </div>
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
                      {formatRepAchievement(r.achievementPct)}
                    </td>
                    <td className="py-2.5 px-4 font-num tabular-nums text-slate-700 whitespace-nowrap">
                      {formatNum(Math.trunc(r.branchCount || 0))}
                    </td>
                    <td className="py-2.5 px-4 font-num tabular-nums text-xs text-slate-600 whitespace-nowrap">
                      {formatNum(r.recordCount)} · {formatNum(r.returnCount)} مرتجع
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
                  الإجمالي ({formatNum(filtered.length)} مندوب)
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
                  {formatRepAchievement(totals.achievementPct)}
                </td>
                <td className="py-2.5 px-4 font-num tabular-nums whitespace-nowrap">
                  {formatNum(Math.trunc(totals.branchCount))}
                </td>
                <td className="py-2.5 px-4 font-num tabular-nums text-xs whitespace-nowrap">
                  {formatNum(totals.recordCount)} · {formatNum(totals.returnCount)} مرتجع
                </td>
              </tr>
            </tfoot>
          )}
        </table>
      </div>
    </div>
  );
};
