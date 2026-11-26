import React, { useMemo, useState } from 'react';
import { ArrowUpDown, Search } from 'lucide-react';
import type { RepMetricRow, SalesSource } from '../types/bi';
import { formatEGP, formatNum } from '../utils/format';
import {
  buildActiveFilterMeta,
  buildExportFilename,
  exportToExcel,
  exportToPdf,
} from '../utils/exporters';
import { ExportButtons } from './ExportButtons';

interface RepsViewProps {
  reps: RepMetricRow[];
  source: SalesSource;
  sectorFilter: string;
  startDate: string;
  endDate: string;
  loading: boolean;
}

type SortKey = 'sales' | 'tax' | 'target' | 'achievementPct' | 'recordCount' | 'branchCount';

function formatRepAchievement(pct: number | null | undefined): string {
  if (pct === null || pct === undefined || !Number.isFinite(pct)) {
    return '—';
  }
  return `${pct.toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}%`;
}

export const RepsView: React.FC<RepsViewProps> = ({
  reps,
  source,
  sectorFilter,
  startDate,
  endDate,
  loading,
}) => {
  const [search, setSearch] = useState('');
  const [sortKey, setSortKey] = useState<SortKey>('sales');
  const [sortAsc, setSortAsc] = useState(false);

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    return reps
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
  }, [reps, search, sortKey, sortAsc]);

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
      achievementPct: achievementPct !== null && Number.isFinite(achievementPct) ? Number(achievementPct.toFixed(2)) : null,
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

  const filterMeta = buildActiveFilterMeta({
    source,
    sector: sectorFilter,
    startDate,
    endDate,
    extra: search.trim() ? [{ label: 'بحث', value: search.trim(), fileTag: 'مفلتر' }] : [],
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
          totalRow: filtered.length > 0
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
          subtitle: `إجمالي المناديب الظاهرين: ${formatNum(filtered.length)}`,
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
          totalRow: filtered.length > 0
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
        <div className="p-5 border-b border-slate-200 flex flex-wrap items-center justify-between gap-4">
          <div>
            <h3 className="text-sm font-bold text-slate-900">جدول أداء مندوبي المبيعات</h3>
            <p className="text-xs text-slate-500 mt-0.5">
              إجمالي المناديب: <span className="font-num">{formatNum(filtered.length)}</span>
            </p>
          </div>
          <div className="flex flex-wrap items-center gap-3">
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

        <div className="overflow-x-auto">
          <table className="w-full text-sm text-right">
            <thead className="bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-600">
              <tr>
                <th className="py-3 px-4 whitespace-nowrap">كود المندوب</th>
                <th className="py-3 px-4 whitespace-nowrap">اسم المندوب</th>
                <th className="py-3 px-4 whitespace-nowrap">القطاع</th>
                <th className="py-3 px-4 whitespace-nowrap">المشرف</th>
                <th className="py-3 px-4 whitespace-nowrap">
                  <button onClick={() => toggleSort('sales')} className="inline-flex items-center gap-1 cursor-pointer hover:text-slate-900">
                    <span>المبيعات</span>
                    <ArrowUpDown className="w-3.5 h-3.5" />
                  </button>
                </th>
                <th className="py-3 px-4 whitespace-nowrap">
                  <button onClick={() => toggleSort('tax')} className="inline-flex items-center gap-1 cursor-pointer hover:text-slate-900">
                    <span>الضريبة</span>
                    <ArrowUpDown className="w-3.5 h-3.5" />
                  </button>
                </th>
                <th className="py-3 px-4 whitespace-nowrap">
                  <button onClick={() => toggleSort('target')} className="inline-flex items-center gap-1 cursor-pointer hover:text-slate-900">
                    <span>الهدف</span>
                    <ArrowUpDown className="w-3.5 h-3.5" />
                  </button>
                </th>
                <th className="py-3 px-4 whitespace-nowrap">
                  <button onClick={() => toggleSort('achievementPct')} className="inline-flex items-center gap-1 cursor-pointer hover:text-slate-900">
                    <span>نسبة التحقيق</span>
                    <ArrowUpDown className="w-3.5 h-3.5" />
                  </button>
                </th>
                <th className="py-3 px-4 whitespace-nowrap">
                  <button onClick={() => toggleSort('branchCount')} className="inline-flex items-center gap-1 cursor-pointer hover:text-slate-900">
                    <span>الفروع</span>
                    <ArrowUpDown className="w-3.5 h-3.5" />
                  </button>
                </th>
                <th className="py-3 px-4 whitespace-nowrap">
                  <button onClick={() => toggleSort('recordCount')} className="inline-flex items-center gap-1 cursor-pointer hover:text-slate-900">
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
                filtered.map((r) => (
                  <tr key={r.repCode} className="hover:bg-slate-50/70">
                    <td className="py-2.5 px-4 font-num text-xs font-medium text-slate-700 whitespace-nowrap">{r.repCode}</td>
                    <td className="py-2.5 px-4 font-medium text-slate-900 whitespace-nowrap">{r.repName}</td>
                    <td className="py-2.5 px-4 text-xs text-slate-600 whitespace-nowrap">{r.sectorName}</td>
                    <td className="py-2.5 px-4 text-xs text-slate-600 whitespace-nowrap">{r.supervisorName}</td>
                    <td className="py-2.5 px-4 font-num tabular-nums font-semibold text-slate-900 whitespace-nowrap" dir="ltr">
                      {formatEGP(r.sales)}
                    </td>
                    <td className="py-2.5 px-4 font-num tabular-nums text-slate-600 whitespace-nowrap" dir="ltr">
                      {formatEGP(r.tax)}
                    </td>
                    <td className="py-2.5 px-4 font-num tabular-nums text-slate-700 whitespace-nowrap" dir="ltr">
                      {formatEGP(r.target)}
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
                    <td className="py-2.5 px-4 font-num tabular-nums text-slate-700 whitespace-nowrap">{formatNum(Math.trunc(r.branchCount || 0))}</td>
                    <td className="py-2.5 px-4 font-num tabular-nums text-xs text-slate-600 whitespace-nowrap">
                      {formatNum(r.recordCount)} · {formatNum(r.returnCount)} مرتجع
                    </td>
                  </tr>
                ))
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
                  <td className="py-2.5 px-4 font-num tabular-nums whitespace-nowrap">{formatNum(Math.trunc(totals.branchCount))}</td>
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
