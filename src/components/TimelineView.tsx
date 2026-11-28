import React, { useMemo } from 'react';
import {
  Area,
  AreaChart,
  Bar,
  BarChart,
  CartesianGrid,
  Legend,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';
import type { SalesSource, TimelinePoint } from '../types/bi';
import { formatEGP, formatNum } from '../utils/format';
import {
  buildActiveFilterMeta,
  buildExportFilename,
  exportToExcel,
  exportToPdf,
} from '../utils/exporters';
import { ExportButtons } from './ExportButtons';

interface TimelineViewProps {
  series: TimelinePoint[];
  source: SalesSource;
  sectorFilter: string;
  startDate: string;
  endDate: string;
  loading: boolean;
}

export const TimelineView: React.FC<TimelineViewProps> = ({
  series,
  source,
  sectorFilter,
  startDate,
  endDate,
  loading,
}) => {
  const totals = useMemo(() => {
    let grossSales = 0;
    let returnsValue = 0;
    let sales = 0;
    let tax = 0;
    let recordCount = 0;
    for (const p of series) {
      grossSales += p.grossSales;
      returnsValue += p.returnsValue;
      sales += p.sales;
      tax += p.tax;
      recordCount += p.recordCount;
    }
    const cumulativeSales = series.length > 0 ? series[series.length - 1].cumulativeSales : 0;
    return { grossSales, returnsValue, sales, tax, cumulativeSales, recordCount };
  }, [series]);

  const filterMeta = buildActiveFilterMeta({
    source,
    sector: sectorFilter,
    startDate,
    endDate,
  });

  const handleExportExcel = () => {
    const filename = buildExportFilename('التحليل_الزمني', filterMeta.fileTags, 'xlsx');
    exportToExcel({
      filename,
      sheets: [
        {
          name: 'التحليل الزمني',
          title: 'تحليلات قطاع السلاسل — التفصيل الزمني اليومي',
          subtitle: filterMeta.summary.join(' | '),
          columns: [
            { header: 'التاريخ', key: 'date', type: 'string', minWidth: 14 },
            { header: 'المبيعات الموجبة', key: 'grossSales', type: 'currency', minWidth: 18 },
            { header: 'المرتجعات', key: 'returnsValue', type: 'currency', minWidth: 16 },
            { header: 'إجمالي قيمة المبيعات', key: 'sales', type: 'currency', minWidth: 18 },
            { header: 'الضريبة', key: 'tax', type: 'currency', minWidth: 15 },
            { header: 'الإجمالي التراكمي', key: 'cumulativeSales', type: 'currency', minWidth: 18 },
            { header: 'عدد السجلات', key: 'recordCount', type: 'integer', minWidth: 14 },
          ],
          rows: series.map((p) => ({
            date: p.date,
            grossSales: p.grossSales,
            returnsValue: p.returnsValue,
            sales: p.sales,
            tax: p.tax,
            cumulativeSales: p.cumulativeSales,
            recordCount: p.recordCount,
          })),
          totalRow: series.length > 0
            ? {
                date: `الإجمالي (${formatNum(series.length)} يوم)`,
                grossSales: totals.grossSales,
                returnsValue: totals.returnsValue,
                sales: totals.sales,
                tax: totals.tax,
                cumulativeSales: totals.cumulativeSales,
                recordCount: totals.recordCount,
              }
            : null,
        },
      ],
    });
  };

  const handleExportPdf = async () => {
    const filename = buildExportFilename('التحليل_الزمني', filterMeta.fileTags, 'pdf');
    await exportToPdf({
      filename,
      tabTitle: 'التحليل الزمني',
      filtersSummary: filterMeta.summary,
      orientation: 'landscape',
      tables: [
        {
          title: 'التفصيل الزمني اليومي',
          columns: [
            { header: 'التاريخ', key: 'date', type: 'string', align: 'center' },
            { header: 'المبيعات الموجبة', key: 'grossSales', type: 'currency' },
            { header: 'المرتجعات', key: 'returnsValue', type: 'currency' },
            { header: 'إجمالي قيمة المبيعات', key: 'sales', type: 'currency' },
            { header: 'الضريبة', key: 'tax', type: 'currency' },
            { header: 'الإجمالي التراكمي', key: 'cumulativeSales', type: 'currency' },
            { header: 'عدد السجلات', key: 'recordCount', type: 'integer' },
          ],
          rows: series.map((p) => ({
            date: p.date,
            grossSales: p.grossSales,
            returnsValue: p.returnsValue,
            sales: p.sales,
            tax: p.tax,
            cumulativeSales: p.cumulativeSales,
            recordCount: p.recordCount,
          })),
          totalRow: series.length > 0
            ? {
                date: `الإجمالي (${formatNum(series.length)} يوم)`,
                grossSales: totals.grossSales,
                returnsValue: totals.returnsValue,
                sales: totals.sales,
                tax: totals.tax,
                cumulativeSales: totals.cumulativeSales,
                recordCount: totals.recordCount,
              }
            : null,
        },
      ],
    });
  };

  if (loading && series.length === 0) {
    return <div className="h-80 bg-white border border-slate-200 rounded-xl animate-pulse" />;
  }

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Cumulative Growth */}
        <div className="bg-white border border-slate-200 rounded-xl p-5">
          <h3 className="text-sm font-bold text-slate-900 mb-1">المبيعات التراكمية عبر الفترة</h3>
          <p className="text-xs text-slate-500 mb-4">تصاعد إجمالي قيمة المبيعات (New Growth Value BD&amp;T) يومًا بيوم</p>
          <div className="h-72 w-full" dir="ltr">
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={series} margin={{ top: 10, right: 20, left: 10, bottom: 20 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" />
                <XAxis dataKey="date" tick={{ fontSize: 11, fill: '#64748B' }} />
                <YAxis tickFormatter={(v: number) => `${(v / 1_000_000).toFixed(1)}M`} tick={{ fontSize: 11, fill: '#64748B' }} />
                <Tooltip
                  formatter={(value) => [formatEGP(Number(value ?? 0)), 'المبيعات التراكمية']}
                  contentStyle={{ borderRadius: '8px', borderColor: '#E2E8F0', fontSize: '12px', direction: 'rtl' }}
                />
                <Area type="monotone" dataKey="cumulativeSales" stroke="#1D4ED8" fill="#DBEAFE" strokeWidth={2} />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* Daily Gross vs Returns */}
        <div className="bg-white border border-slate-200 rounded-xl p-5">
          <h3 className="text-sm font-bold text-slate-900 mb-1">المبيعات الإجمالية مقابل المرتجعات اليومية</h3>
          <p className="text-xs text-slate-500 mb-4">تُعرض المرتجعات بإشارتها السالبة الأصلية دون تجاهل</p>
          <div className="h-72 w-full" dir="ltr">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={series} margin={{ top: 10, right: 20, left: 10, bottom: 20 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" />
                <XAxis dataKey="date" tick={{ fontSize: 11, fill: '#64748B' }} />
                <YAxis tickFormatter={(v: number) => `${(v / 1_000_000).toFixed(1)}M`} tick={{ fontSize: 11, fill: '#64748B' }} />
                <Tooltip
                  formatter={(value, name) => [formatEGP(Number(value ?? 0)), String(name ?? '')]}
                  contentStyle={{ borderRadius: '8px', borderColor: '#E2E8F0', fontSize: '12px', direction: 'rtl' }}
                />
                <Legend wrapperStyle={{ fontSize: '12px' }} />
                <Bar dataKey="grossSales" name="المبيعات الموجبة" fill="#0F766E" radius={[4, 4, 0, 0]} />
                <Bar dataKey="returnsValue" name="المرتجعات (سالبة)" fill="#DC2626" radius={[0, 0, 4, 4]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="px-5 py-4 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3">
          <h3 className="text-sm font-bold text-slate-900">التفصيل الزمني اليومي</h3>
          <ExportButtons
            onExportExcel={handleExportExcel}
            onExportPdf={handleExportPdf}
            disabled={series.length === 0}
          />
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-sm text-right">
            <thead className="bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-600">
              <tr>
                <th className="py-3 px-4">التاريخ</th>
                <th className="py-3 px-4">المبيعات الموجبة</th>
                <th className="py-3 px-4">المرتجعات</th>
                <th className="py-3 px-4">إجمالي قيمة المبيعات</th>
                <th className="py-3 px-4">الضريبة</th>
                <th className="py-3 px-4">الإجمالي التراكمي</th>
                <th className="py-3 px-4">عدد السجلات</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {series.length === 0 ? (
                <tr>
                  <td colSpan={7} className="py-8 text-center text-xs text-slate-500">
                    لا توجد بيانات زمنية متاحة
                  </td>
                </tr>
              ) : (
                series.map((p) => (
                  <tr key={p.date} className="hover:bg-slate-50/70">
                    <td className="py-2.5 px-4 font-num text-xs font-semibold text-slate-800">{p.date}</td>
                    <td className="py-2.5 px-4 font-num text-slate-700">{formatEGP(p.grossSales)}</td>
                    <td className={`py-2.5 px-4 font-num ${p.returnsValue < 0 ? 'text-rose-700' : 'text-slate-500'}`}>
                      {formatEGP(p.returnsValue)}
                    </td>
                    <td className="py-2.5 px-4 font-num font-semibold text-slate-900">{formatEGP(p.sales)}</td>
                    <td className="py-2.5 px-4 font-num text-slate-600">{formatEGP(p.tax)}</td>
                    <td className="py-2.5 px-4 font-num font-medium text-teal-700">{formatEGP(p.cumulativeSales)}</td>
                    <td className="py-2.5 px-4 font-num text-slate-600">{formatNum(p.recordCount)}</td>
                  </tr>
                ))
              )}
            </tbody>
            {series.length > 0 && (
              <tfoot className="bg-sky-50 border-t-2 border-slate-300 font-bold text-slate-900">
                <tr>
                  <td className="py-2.5 px-4 whitespace-nowrap">الإجمالي ({formatNum(series.length)} يوم)</td>
                  <td className="py-2.5 px-4 font-num tabular-nums whitespace-nowrap">{formatEGP(totals.grossSales)}</td>
                  <td className={`py-2.5 px-4 font-num tabular-nums whitespace-nowrap ${totals.returnsValue < 0 ? 'text-rose-700' : ''}`}>
                    {formatEGP(totals.returnsValue)}
                  </td>
                  <td className="py-2.5 px-4 font-num tabular-nums whitespace-nowrap">{formatEGP(totals.sales)}</td>
                  <td className="py-2.5 px-4 font-num tabular-nums whitespace-nowrap">{formatEGP(totals.tax)}</td>
                  <td className="py-2.5 px-4 font-num tabular-nums whitespace-nowrap text-teal-900">{formatEGP(totals.cumulativeSales)}</td>
                  <td className="py-2.5 px-4 font-num tabular-nums whitespace-nowrap">{formatNum(totals.recordCount)}</td>
                </tr>
              </tfoot>
            )}
          </table>
        </div>
      </div>
    </div>
  );
};
