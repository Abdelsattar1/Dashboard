import React, { useMemo, useState } from 'react';
import { Search } from 'lucide-react';
import type { AggregationMetricRow, DepartmentChainRow, SalesSource, SectorMetricRow } from '../types/bi';
import { formatEGP, formatNum, formatPct, formatPctOneDecimal } from '../utils/format';
import {
  buildActiveFilterMeta,
  buildExportFilename,
  exportToExcel,
  exportToPdf,
  type PdfTableRow,
} from '../utils/exporters';
import { ExportButtons } from './ExportButtons';

interface SectorsViewProps {
  sectors: SectorMetricRow[];
  aggregations: AggregationMetricRow[];
  chains: DepartmentChainRow[];
  source: SalesSource;
  globalSectorFilter: string;
  startDate: string;
  endDate: string;
  loading: boolean;
}

interface SupervisorGroup {
  supervisorName: string;
  chains: DepartmentChainRow[];
  commercialTotal: number;
  industrialTotal: number;
  salesTotal: number;
  taxTotal: number;
}

function normalizeArabicText(input: string): string {
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

export const SectorsView: React.FC<SectorsViewProps> = ({
  sectors,
  aggregations,
  chains,
  source,
  globalSectorFilter,
  startDate,
  endDate,
  loading,
}) => {
  const [search, setSearch] = useState('');
  const [sectorFilter, setSectorFilter] = useState('ALL');
  const [supervisorFilter, setSupervisorFilter] = useState('ALL');

  const availableSectors = useMemo(() => {
    const set = new Set<string>();
    for (const s of sectors) {
      if (s.sectorName) set.add(s.sectorName);
    }
    for (const c of chains) {
      if (c.sectorName) set.add(c.sectorName);
    }
    return Array.from(set);
  }, [sectors, chains]);

  const availableSupervisors = useMemo(() => {
    const map = new Map<string, string>();
    for (const c of chains) {
      if (c.supervisorName) {
        const norm = normalizeArabicText(c.supervisorName);
        if (norm && !map.has(norm)) map.set(norm, c.supervisorName.trim());
      }
    }
    return Array.from(map.values());
  }, [chains]);

  const filteredChains = useMemo(() => {
    const q = normalizeArabicText(search);
    const normSupFilter = supervisorFilter === 'ALL' ? 'ALL' : normalizeArabicText(supervisorFilter);
    return chains.filter((c) => {
      if (sectorFilter !== 'ALL' && c.sectorName !== sectorFilter) return false;
      if (normSupFilter !== 'ALL' && normalizeArabicText(c.supervisorName) !== normSupFilter) return false;
      if (!q) return true;
      return (
        normalizeArabicText(c.supervisorName).includes(q) ||
        c.chainCode.toLowerCase().includes(q) ||
        normalizeArabicText(c.chainName).includes(q) ||
        normalizeArabicText(c.sectorName).includes(q)
      );
    });
  }, [chains, search, sectorFilter, supervisorFilter]);

  const supervisorGroups = useMemo<SupervisorGroup[]>(() => {
    const map = new Map<string, { group: SupervisorGroup; chainByNormName: Map<string, DepartmentChainRow> }>();

    for (const c of filteredChains) {
      const supDisplay = (c.supervisorName || 'غير محدد').trim() || 'غير محدد';
      const supKey = normalizeArabicText(supDisplay) || 'غير محدد';
      let bucket = map.get(supKey);
      if (!bucket) {
        bucket = {
          group: {
            supervisorName: supDisplay,
            chains: [],
            commercialTotal: 0,
            industrialTotal: 0,
            salesTotal: 0,
            taxTotal: 0,
          },
          chainByNormName: new Map(),
        };
        map.set(supKey, bucket);
      }

      const normChain = normalizeArabicText(c.chainName) || c.chainCode;
      const existingChain = bucket.chainByNormName.get(normChain);
      if (!existingChain) {
        const cloned: DepartmentChainRow = { ...c, supervisorName: supDisplay };
        bucket.chainByNormName.set(normChain, cloned);
        bucket.group.chains.push(cloned);
      } else {
        existingChain.commercialSales = Number((existingChain.commercialSales + c.commercialSales).toFixed(2));
        existingChain.industrialSales = Number((existingChain.industrialSales + c.industrialSales).toFixed(2));
        existingChain.sales = Number((existingChain.sales + c.sales).toFixed(2));
        existingChain.tax = Number((existingChain.tax + c.tax).toFixed(2));
        existingChain.recordCount += c.recordCount;
        if (c.chainCode && !existingChain.chainCode.includes(c.chainCode)) {
          existingChain.chainCode = existingChain.chainCode ? `${existingChain.chainCode}, ${c.chainCode}` : c.chainCode;
        }
      }

      bucket.group.commercialTotal += c.commercialSales;
      bucket.group.industrialTotal += c.industrialSales;
      bucket.group.salesTotal += c.sales;
      bucket.group.taxTotal += c.tax;
    }

    const groups = Array.from(map.values()).map((b) => b.group);
    for (const g of groups) {
      g.commercialTotal = Number(g.commercialTotal.toFixed(2));
      g.industrialTotal = Number(g.industrialTotal.toFixed(2));
      g.salesTotal = Number(g.salesTotal.toFixed(2));
      g.taxTotal = Number(g.taxTotal.toFixed(2));
      for (const ch of g.chains) {
        ch.shareOfSupervisorPct = g.salesTotal !== 0 ? Number(((ch.sales / g.salesTotal) * 100).toFixed(1)) : 0;
      }
      g.chains.sort((a, b) => {
        const diff = b.sales - a.sales;
        if (diff !== 0) return diff;
        return a.chainName.localeCompare(b.chainName, 'ar');
      });
    }

    groups.sort((a, b) => {
      if (a.supervisorName === 'غير محدد' && b.supervisorName !== 'غير محدد') return 1;
      if (b.supervisorName === 'غير محدد' && a.supervisorName !== 'غير محدد') return -1;
      const diff = b.salesTotal - a.salesTotal;
      if (diff !== 0) return diff;
      return a.supervisorName.localeCompare(b.supervisorName, 'ar');
    });

    return groups;
  }, [filteredChains]);

  const unifiedChainCount = useMemo(
    () => supervisorGroups.reduce((acc, g) => acc + g.chains.length, 0),
    [supervisorGroups]
  );

  const grandTotals = useMemo(() => {
    let commercial = 0;
    let industrial = 0;
    let sales = 0;
    let tax = 0;
    for (const g of supervisorGroups) {
      commercial += g.commercialTotal;
      industrial += g.industrialTotal;
      sales += g.salesTotal;
      tax += g.taxTotal;
    }
    return {
      commercial: Number(commercial.toFixed(2)),
      industrial: Number(industrial.toFixed(2)),
      sales: Number(sales.toFixed(2)),
      tax: Number(tax.toFixed(2)),
    };
  }, [supervisorGroups]);

  const hasLegacyMissingSupervisors = useMemo(
    () => chains.length > 0 && chains.every((c) => !c.supervisorName || c.supervisorName === 'غير محدد'),
    [chains]
  );

  const effectiveSectorDisplay = sectorFilter !== 'ALL' ? sectorFilter : globalSectorFilter;
  const filterMeta = buildActiveFilterMeta({
    source,
    sector: effectiveSectorDisplay,
    startDate,
    endDate,
    extra: [
      supervisorFilter !== 'ALL' ? { label: 'المشرف', value: supervisorFilter, fileTag: supervisorFilter } : { label: 'المشرف' },
      search.trim() ? { label: 'بحث', value: search.trim(), fileTag: 'مفلتر' } : { label: 'بحث' },
    ],
  });

  const buildChainExportRows = (): PdfTableRow[] => {
    const out: PdfTableRow[] = [];
    for (const group of supervisorGroups) {
      group.chains.forEach((c, idx) => {
        out.push({
          supervisorName: idx === 0 ? group.supervisorName : '',
          chainName: c.chainName,
          chainCode: c.chainCode,
          sectorName: c.sectorName,
          commercialSales: c.commercialSales,
          industrialSales: c.industrialSales,
          sales: c.sales,
          tax: c.tax,
          sharePct: c.shareOfSupervisorPct,
        });
      });
      out.push({
        supervisorName: `إجمالي ${group.supervisorName}`,
        chainName: `(${formatNum(group.chains.length)} سلسلة)`,
        chainCode: '',
        sectorName: '',
        commercialSales: group.commercialTotal,
        industrialSales: group.industrialTotal,
        sales: group.salesTotal,
        tax: group.taxTotal,
        sharePct: group.salesTotal !== 0 ? 100 : 0,
        _tone: 'subtotal',
      });
    }
    return out;
  };

  const chainColumns = [
    { header: 'اسم المشرف', key: 'supervisorName', type: 'string' as const, align: 'right' as const, minWidth: 20 },
    { header: 'اسم السلسلة', key: 'chainName', type: 'string' as const, align: 'right' as const, minWidth: 24 },
    { header: 'مبيعات التجارى', key: 'commercialSales', type: 'currency' as const, minWidth: 16 },
    { header: 'مبيعات الصناعى', key: 'industrialSales', type: 'currency' as const, minWidth: 16 },
    { header: 'الإجمالى', key: 'sales', type: 'currency' as const, minWidth: 16 },
    { header: 'الضريبة', key: 'tax', type: 'currency' as const, minWidth: 15 },
    { header: 'نسبة المشاركة', key: 'sharePct', type: 'percent' as const, minWidth: 14 },
  ];

  const handleExportExcel = () => {
    const filename = buildExportFilename('القطاعات_والإدارات', filterMeta.fileTags, 'xlsx');
    const chainRows = buildChainExportRows();
    exportToExcel({
      filename,
      sheets: [
        {
          name: 'الإدارات والسلاسل',
          title: 'تحليلات قطاع السلاسل — تفاصيل الإدارات والسلاسل',
          subtitle: filterMeta.summary.join(' | '),
          columns: [
            { header: 'اسم المشرف', key: 'supervisorName', type: 'string', minWidth: 20 },
            { header: 'اسم السلسلة', key: 'chainName', type: 'string', minWidth: 24 },
            { header: 'كود السلسلة', key: 'chainCode', type: 'string', minWidth: 14 },
            { header: 'القطاع', key: 'sectorName', type: 'string', minWidth: 18 },
            { header: 'مبيعات التجارى', key: 'commercialSales', type: 'currency', minWidth: 16 },
            { header: 'مبيعات الصناعى', key: 'industrialSales', type: 'currency', minWidth: 16 },
            { header: 'الإجمالى', key: 'sales', type: 'currency', minWidth: 16 },
            { header: 'الضريبة', key: 'tax', type: 'currency', minWidth: 15 },
            { header: 'نسبة المشاركة', key: 'sharePct', type: 'percent', minWidth: 14 },
          ],
          rows: chainRows,
          totalRow: supervisorGroups.length > 0
            ? {
                supervisorName: 'الإجمالي العام',
                chainName: `${formatNum(unifiedChainCount)} سلسلة`,
                chainCode: '',
                sectorName: '',
                commercialSales: grandTotals.commercial,
                industrialSales: grandTotals.industrial,
                sales: grandTotals.sales,
                tax: grandTotals.tax,
                sharePct: grandTotals.sales !== 0 ? 100 : 0,
              }
            : null,
        },
        {
          name: 'ملخص القطاعات',
          title: 'تحليلات قطاع السلاسل — ملخص القطاعات',
          subtitle: filterMeta.summary.join(' | '),
          columns: [
            { header: 'القطاع', key: 'sectorName', type: 'string', minWidth: 20 },
            { header: 'المبيعات', key: 'sales', type: 'currency', minWidth: 18 },
            { header: 'الضريبة', key: 'tax', type: 'currency', minWidth: 16 },
            { header: 'الهدف', key: 'target', type: 'currency', minWidth: 18 },
            { header: 'الفرق', key: 'variance', type: 'currency', minWidth: 18 },
            { header: 'نسبة التحقيق', key: 'achievementPct', type: 'percent', minWidth: 14 },
          ],
          rows: sectors.map((s) => ({
            sectorName: s.sectorName,
            sales: s.sales,
            tax: s.tax,
            target: s.target,
            variance: s.variance,
            achievementPct: s.achievementPct,
          })),
        },
      ],
    });
  };

  const handleExportPdf = async () => {
    const filename = buildExportFilename('القطاعات_والإدارات', filterMeta.fileTags, 'pdf');
    await exportToPdf({
      filename,
      tabTitle: 'القطاعات والإدارات',
      filtersSummary: filterMeta.summary,
      orientation: 'landscape',
      tables: [
        {
          title: 'تفاصيل الإدارات والسلاسل',
          subtitle: `المشرفون: ${formatNum(supervisorGroups.length)} · السلاسل: ${formatNum(unifiedChainCount)}`,
          columns: chainColumns,
          rows: buildChainExportRows(),
          totalRow: supervisorGroups.length > 0
            ? {
                supervisorName: 'الإجمالي العام',
                chainName: `${formatNum(unifiedChainCount)} سلسلة`,
                commercialSales: grandTotals.commercial,
                industrialSales: grandTotals.industrial,
                sales: grandTotals.sales,
                tax: grandTotals.tax,
                sharePct: grandTotals.sales !== 0 ? 100 : 0,
              }
            : null,
        },
      ],
    });
  };

  if (loading && sectors.length === 0) {
    return <div className="h-64 bg-white border border-slate-200 rounded-xl animate-pulse" />;
  }

  return (
    <div className="space-y-6">
      {/* Sector Cards */}
      <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-4">
        {sectors.map((s) => (
          <div key={s.sectorName} className="bg-white border border-slate-200 rounded-xl p-5">
            <div className="flex items-center justify-between mb-2">
              <h3 className="text-sm font-bold text-slate-900">{s.sectorName}</h3>
              <span className="font-num text-xs font-semibold text-teal-700">{formatPct(s.achievementPct)}</span>
            </div>
            <div className="text-xl font-bold font-num text-slate-900 mb-3">{formatEGP(s.sales)}</div>
            <div className="space-y-1 text-xs text-slate-600 border-t border-slate-100 pt-3">
              <div className="flex justify-between">
                <span>الهدف:</span>
                <span className="font-num text-slate-800">{formatEGP(s.target)}</span>
              </div>
              <div className="flex justify-between">
                <span>الضريبة:</span>
                <span className="font-num text-slate-800">{formatEGP(s.tax)}</span>
              </div>
              <div className="flex justify-between">
                <span>الفرق:</span>
                <span className={`font-num font-medium ${s.variance >= 0 ? 'text-emerald-700' : 'text-rose-700'}`}>
                  {formatEGP(s.variance)}
                </span>
              </div>
            </div>
          </div>
        ))}
      </div>

      {/* Aggregations strip */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        {aggregations.map((agg) => (
          <div key={agg.key} className="bg-slate-100/70 border border-slate-200 rounded-xl p-4">
            <div className="text-xs font-bold text-slate-800">{agg.title}</div>
            <div className="text-[11px] text-slate-500 mb-2">{agg.subtitle}</div>
            <div className="text-lg font-bold font-num text-slate-900">{formatEGP(agg.sales)}</div>
            <div className="text-xs text-slate-600 mt-1">
              نسبة التحقيق: <span className="font-num font-semibold text-teal-700">{formatPct(agg.achievementPct)}</span>
            </div>
          </div>
        ))}
      </div>

      {/* Hierarchical Supervisor-Grouped Chains Table with Internal Scroll */}
      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="p-4 lg:p-5 border-b border-slate-200 flex flex-wrap items-center justify-between gap-4">
          <div>
            <h3 className="text-sm font-bold text-slate-900">
              تفاصيل الإدارات والسلاسل (مرتبطة عبر Payer / كود السلسلة)
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              المشرفون: <span className="font-num font-semibold text-slate-700">{formatNum(supervisorGroups.length)}</span>
              {' · '}
              السلاسل: <span className="font-num font-semibold text-slate-700">{formatNum(unifiedChainCount)}</span>
              {' · '}
              التصنيف التجاري/الصناعي مرتبط بشيت تعريف الأصناف (التصنيف A)
            </p>
            {hasLegacyMissingSupervisors && (
              <p className="text-[11px] text-sky-800 mt-1">
                ملاحظة: البيانات الحالية مستوردة قبل تفعيل ربط «اسم المشرف» من ورقة «تعريف الإدارات والقطاعات» — أعد رفع ملف Excel لتحديث المشرفين.
              </p>
            )}
          </div>

          <div className="flex flex-wrap items-center gap-2.5">
            <select
              value={sectorFilter}
              onChange={(e) => setSectorFilter(e.target.value)}
              className="px-3 py-1.5 text-xs border border-slate-300 rounded-lg bg-white text-slate-700 focus:outline-none focus:ring-2 focus:ring-teal-600"
            >
              <option value="ALL">جميع القطاعات المصرح بها</option>
              {availableSectors.map((sectorName) => (
                <option key={sectorName} value={sectorName}>
                  {sectorName}
                </option>
              ))}
            </select>

            <select
              value={supervisorFilter}
              onChange={(e) => setSupervisorFilter(e.target.value)}
              className="px-3 py-1.5 text-xs border border-slate-300 rounded-lg bg-white text-slate-700 focus:outline-none focus:ring-2 focus:ring-teal-600"
            >
              <option value="ALL">جميع المشرفين</option>
              {availableSupervisors.map((sup) => (
                <option key={sup} value={sup}>
                  {sup}
                </option>
              ))}
            </select>

            <div className="relative">
              <Search className="w-3.5 h-3.5 text-slate-400 absolute right-3 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="بحث بالمشرف أو السلسلة..."
                className="pr-8 pl-3 py-1.5 text-xs border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-teal-600 w-52"
              />
            </div>

            <ExportButtons
              onExportExcel={handleExportExcel}
              onExportPdf={handleExportPdf}
              disabled={supervisorGroups.length === 0 && sectors.length === 0}
            />
          </div>
        </div>

        {/* Fixed-Height Internal Scroll Container with Sticky Header */}
        <div className="max-h-[550px] overflow-auto">
          <table className="w-full text-xs text-right border-collapse">
            <thead className="sticky top-0 z-10 bg-slate-100 border-b-2 border-slate-300 text-slate-800 font-bold shadow-[0_1px_0_0_rgba(203,213,225,1)]">
              <tr>
                <th className="py-2.5 px-3 whitespace-nowrap">اسم المشرف</th>
                <th className="py-2.5 px-3 whitespace-nowrap">اسم السلسلة</th>
                <th className="py-2.5 px-3 whitespace-nowrap">مبيعات التجارى</th>
                <th className="py-2.5 px-3 whitespace-nowrap">مبيعات الصناعى</th>
                <th className="py-2.5 px-3 whitespace-nowrap">الإجمالى</th>
                <th className="py-2.5 px-3 whitespace-nowrap">الضريبة</th>
                <th className="py-2.5 px-3 whitespace-nowrap">نسبة المشاركة</th>
              </tr>
            </thead>
            <tbody>
              {supervisorGroups.length === 0 ? (
                <tr>
                  <td colSpan={7} className="py-10 text-center text-xs text-slate-500">
                    لا توجد بيانات سلاسل مطابقة للبحث
                  </td>
                </tr>
              ) : (
                supervisorGroups.map((group) => (
                  <React.Fragment key={group.supervisorName}>
                    {group.chains.map((c, idx) => {
                      const sharePct = c.shareOfSupervisorPct;

                      return (
                        <tr
                          key={`${group.supervisorName}-${c.chainName}-${idx}`}
                          className="border-b border-slate-100 hover:bg-slate-50/80 transition-colors"
                        >
                          <td className="py-1.5 px-3 whitespace-nowrap font-semibold text-slate-900">
                            {idx === 0 ? group.supervisorName : ''}
                          </td>
                          <td className="py-1.5 px-3 whitespace-nowrap text-slate-800" title={`الكود: ${c.chainCode} | ${c.sectorName}`}>
                            {c.chainName}
                          </td>
                          <td
                            className={`py-1.5 px-3 font-num tabular-nums whitespace-nowrap ${
                              c.commercialSales < 0 ? 'text-rose-700' : 'text-slate-700'
                            }`}
                            dir="ltr"
                            title={formatEGP(c.commercialSales)}
                          >
                            {formatNum(c.commercialSales, 0)}
                          </td>
                          <td
                            className={`py-1.5 px-3 font-num tabular-nums whitespace-nowrap ${
                              c.industrialSales < 0 ? 'text-rose-700' : 'text-slate-700'
                            }`}
                            dir="ltr"
                            title={formatEGP(c.industrialSales)}
                          >
                            {formatNum(c.industrialSales, 0)}
                          </td>
                          <td
                            className={`py-1.5 px-3 font-num tabular-nums whitespace-nowrap font-semibold ${
                              c.sales < 0 ? 'text-rose-700' : 'text-slate-900'
                            }`}
                            dir="ltr"
                            title={formatEGP(c.sales)}
                          >
                            {formatNum(c.sales, 0)}
                          </td>
                          <td
                            className={`py-1.5 px-3 font-num tabular-nums whitespace-nowrap ${
                              c.tax < 0 ? 'text-rose-700' : 'text-slate-600'
                            }`}
                            dir="ltr"
                            title={formatEGP(c.tax)}
                          >
                            {formatNum(c.tax, 0)}
                          </td>
                          <td
                            className={`py-1.5 px-3 font-num tabular-nums whitespace-nowrap font-medium ${
                              sharePct < 0 ? 'text-rose-700' : 'text-teal-800'
                            }`}
                            dir="ltr"
                          >
                            {formatPctOneDecimal(sharePct)}
                          </td>
                        </tr>
                      );
                    })}

                    {/* Supervisor Subtotal Row (Bold + Bottom Underline) */}
                    <tr className="bg-slate-50/95 font-bold text-slate-900 border-t border-slate-300 border-b-2 border-b-slate-400">
                      <td className="py-2 px-3 whitespace-nowrap">إجمالي {group.supervisorName}</td>
                      <td className="py-2 px-3 whitespace-nowrap text-slate-500 text-[11px] font-normal">
                        ({formatNum(group.chains.length)} سلسلة)
                      </td>
                      <td
                        className={`py-2 px-3 font-num tabular-nums whitespace-nowrap ${
                          group.commercialTotal < 0 ? 'text-rose-700' : 'text-slate-900'
                        }`}
                        dir="ltr"
                        title={formatEGP(group.commercialTotal)}
                      >
                        {formatNum(group.commercialTotal, 0)}
                      </td>
                      <td
                        className={`py-2 px-3 font-num tabular-nums whitespace-nowrap ${
                          group.industrialTotal < 0 ? 'text-rose-700' : 'text-slate-900'
                        }`}
                        dir="ltr"
                        title={formatEGP(group.industrialTotal)}
                      >
                        {formatNum(group.industrialTotal, 0)}
                      </td>
                      <td
                        className={`py-2 px-3 font-num tabular-nums whitespace-nowrap ${
                          group.salesTotal < 0 ? 'text-rose-700' : 'text-slate-900'
                        }`}
                        dir="ltr"
                        title={formatEGP(group.salesTotal)}
                      >
                        {formatNum(group.salesTotal, 0)}
                      </td>
                      <td
                        className={`py-2 px-3 font-num tabular-nums whitespace-nowrap ${
                          group.taxTotal < 0 ? 'text-rose-700' : 'text-slate-800'
                        }`}
                        dir="ltr"
                        title={formatEGP(group.taxTotal)}
                      >
                        {formatNum(group.taxTotal, 0)}
                      </td>
                      <td className="py-2 px-3 font-num tabular-nums whitespace-nowrap text-teal-900" dir="ltr">
                        {group.salesTotal !== 0 ? '100.0%' : '0.0%'}
                      </td>
                    </tr>
                  </React.Fragment>
                ))
              )}
            </tbody>

            {supervisorGroups.length > 0 && (
              <tfoot className="bg-sky-50 border-t-2 border-slate-400 font-bold text-slate-900">
                <tr>
                  <td className="py-2.5 px-3 whitespace-nowrap">الإجمالي العام</td>
                  <td className="py-2.5 px-3 whitespace-nowrap text-slate-600 font-normal text-[11px]">
                    {formatNum(unifiedChainCount)} سلسلة
                  </td>
                  <td className="py-2.5 px-3 font-num tabular-nums whitespace-nowrap" dir="ltr" title={formatEGP(grandTotals.commercial)}>
                    {formatNum(grandTotals.commercial, 0)}
                  </td>
                  <td className="py-2.5 px-3 font-num tabular-nums whitespace-nowrap" dir="ltr" title={formatEGP(grandTotals.industrial)}>
                    {formatNum(grandTotals.industrial, 0)}
                  </td>
                  <td className="py-2.5 px-3 font-num tabular-nums whitespace-nowrap" dir="ltr" title={formatEGP(grandTotals.sales)}>
                    {formatNum(grandTotals.sales, 0)}
                  </td>
                  <td className="py-2.5 px-3 font-num tabular-nums whitespace-nowrap" dir="ltr" title={formatEGP(grandTotals.tax)}>
                    {formatNum(grandTotals.tax, 0)}
                  </td>
                  <td className="py-2.5 px-3 font-num tabular-nums whitespace-nowrap text-teal-900" dir="ltr">
                    {grandTotals.sales !== 0 ? '100.0%' : '0.0%'}
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
