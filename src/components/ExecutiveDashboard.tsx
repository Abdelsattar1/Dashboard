import React, { useMemo, useRef } from 'react';
import {
  AlertTriangle,
  ArrowDown,
  ArrowUp,
  CheckCircle2,
  FileSpreadsheet,
  Upload,
} from 'lucide-react';
import {
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Legend,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';
import type { DashboardSummaryResponse, SalesSource, UserProfile } from '../types/bi';
import {
  formatEGP,
  formatMillionOrFull,
  formatNum,
  formatPct,
  formatPctOneDecimal,
} from '../utils/format';
import {
  buildActiveFilterMeta,
  buildExportFilename,
  captureChartSvgToPng,
  exportToExcel,
  exportToPdf,
  type PdfChartItem,
} from '../utils/exporters';
import { ExportButtons } from './ExportButtons';

interface ExecutiveDashboardProps {
  data: DashboardSummaryResponse | null;
  loading: boolean;
  source: SalesSource;
  sectorFilter: string;
  startDate: string;
  endDate: string;
  user: UserProfile;
  onNavigateImport: () => void;
  onNavigateBranchesUnlinked: () => void;
}

const SECTOR_COLORS: Record<string, string> = {
  'التوب ريتيل': '#0F5A78',
  'التوريد المباشر': '#E06D2B',
  'التوريد المركزي': '#1A651A',
  'الجيش': '#0097DB',
};

const FALLBACK_COLORS = ['#0F5A78', '#E06D2B', '#1A651A', '#0097DB', '#475569'];

const TABLE_SECTOR_LABELS: Record<string, string> = {
  'التوريد المباشر': 'قطاع السلاسل مباشر',
  'التوريد المركزي': 'قطاع السلاسل مركزي',
  'التوب ريتيل': 'قطاع التجزئة المميزة',
  'الجيش': 'قطاع الجيش',
};

function getTableSectorLabel(sectorName: string): string {
  return TABLE_SECTOR_LABELS[sectorName] ?? sectorName;
}

function getSectorColor(sectorName: string, index: number): string {
  return SECTOR_COLORS[sectorName] ?? FALLBACK_COLORS[index % FALLBACK_COLORS.length];
}

interface PieLabelProps {
  cx?: number;
  cy?: number;
  midAngle?: number;
  innerRadius?: number;
  outerRadius?: number;
  percent?: number;
}

function renderPieSliceLabel({
  cx = 0,
  cy = 0,
  midAngle = 0,
  innerRadius = 0,
  outerRadius = 0,
  percent = 0,
}: PieLabelProps) {
  const pct = Math.round(percent * 100);
  if (pct < 4) return null;
  const radius = innerRadius + (outerRadius - innerRadius) * 0.56;
  const rad = (-midAngle * Math.PI) / 180;
  const x = cx + radius * Math.cos(rad);
  const y = cy + radius * Math.sin(rad);

  return (
    <text
      x={x}
      y={y}
      fill="#FFFFFF"
      textAnchor="middle"
      dominantBaseline="central"
      fontSize={12}
      fontWeight={700}
    >
      {`${pct}%`}
    </text>
  );
}

interface TableDisplayRow {
  key: string;
  label: string;
  sales: number;
  tax: number;
  target: number;
  variance: number;
  achievementPct: number | null;
  rowTone: 'private' | 'direct_army' | 'chains' | 'sector';
}

export const ExecutiveDashboard: React.FC<ExecutiveDashboardProps> = ({
  data,
  loading,
  source,
  sectorFilter,
  startDate,
  endDate,
  user,
  onNavigateImport,
  onNavigateBranchesUnlinked,
}) => {
  const sectors = data?.sectors ?? [];
  const aggregations = data?.aggregations ?? [];
  const pieChartContainerRef = useRef<HTMLDivElement>(null);
  const barChartContainerRef = useRef<HTMLDivElement>(null);

  const pieData = useMemo(() => {
    const positive = sectors.filter((s) => s.sales > 0);
    const total = positive.reduce((acc, s) => acc + s.sales, 0);
    return positive.map((s, idx) => ({
      name: s.sectorName,
      value: s.sales,
      sharePct: total > 0 ? Number(((s.sales / total) * 100).toFixed(1)) : 0,
      color: getSectorColor(s.sectorName, idx),
    }));
  }, [sectors]);

  const barData = useMemo(() => {
    return [...sectors]
      .sort((a, b) => b.target - a.target || b.sales - a.sales)
      .map((s) => ({
        sectorName: s.sectorName,
        salesMillions: Number((s.sales / 1_000_000).toFixed(2)),
        targetMillions: Number((s.target / 1_000_000).toFixed(2)),
        rawSales: s.sales,
        rawTarget: s.target,
      }));
  }, [sectors]);

  const tableRows = useMemo<TableDisplayRow[]>(() => {
    const rows: TableDisplayRow[] = [];
    const aggPrivate = aggregations.find((a) => a.key === 'private_markets');
    const aggDirectArmy = aggregations.find((a) => a.key === 'direct_and_army');
    const aggChains = aggregations.find((a) => a.key === 'chains_sector');

    const sectorNamesPresent = new Set(sectors.map((s) => s.sectorName));

    if (aggPrivate && sectors.length > 1) {
      rows.push({
        key: 'agg-private',
        label: 'إجمالي مبيعات الأسواق الخاصة',
        sales: aggPrivate.sales,
        tax: aggPrivate.tax,
        target: aggPrivate.target,
        variance: aggPrivate.variance,
        achievementPct: aggPrivate.achievementPct,
        rowTone: 'private',
      });
    }

    if (
      aggDirectArmy &&
      sectors.length > 1 &&
      aggDirectArmy.includedSectors.some((n) => sectorNamesPresent.has(n))
    ) {
      rows.push({
        key: 'agg-direct-army',
        label: 'إجمالي مبيعات المباشر والجيش',
        sales: aggDirectArmy.sales,
        tax: aggDirectArmy.tax,
        target: aggDirectArmy.target,
        variance: aggDirectArmy.variance,
        achievementPct: aggDirectArmy.achievementPct,
        rowTone: 'direct_army',
      });
    }

    if (
      aggChains &&
      sectors.length > 1 &&
      aggChains.includedSectors.some((n) => sectorNamesPresent.has(n))
    ) {
      rows.push({
        key: 'agg-chains',
        label: 'إجمالي مبيعات قطاع السلاسل',
        sales: aggChains.sales,
        tax: aggChains.tax,
        target: aggChains.target,
        variance: aggChains.variance,
        achievementPct: aggChains.achievementPct,
        rowTone: 'chains',
      });
    }

    for (const s of sectors) {
      rows.push({
        key: `sector-${s.sectorName}`,
        label: getTableSectorLabel(s.sectorName),
        sales: s.sales,
        tax: s.tax,
        target: s.target,
        variance: s.variance,
        achievementPct: s.achievementPct,
        rowTone: 'sector',
      });
    }

    return rows;
  }, [sectors, aggregations]);

  if (loading && !data) {
    return (
      <div className="space-y-5">
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3">
          {Array.from({ length: 5 }).map((_, i) => (
            <div key={i} className="h-28 bg-white border border-slate-200 rounded-lg p-4 animate-pulse">
              <div className="h-3 w-28 bg-slate-200 rounded mb-4 mx-auto" />
              <div className="h-6 w-24 bg-slate-200 rounded mx-auto" />
            </div>
          ))}
        </div>
        <div className="h-96 bg-white border border-slate-200 rounded-lg animate-pulse" />
      </div>
    );
  }

  if (!data || !data.hasData) {
    return (
      <div className="bg-white border border-slate-200 rounded-xl p-10 text-center max-w-2xl mx-auto mt-8">
        <div className="w-12 h-12 rounded-xl bg-teal-50 border border-teal-200 text-teal-700 flex items-center justify-center mx-auto mb-4">
          <FileSpreadsheet className="w-6 h-6" />
        </div>
        <h2 className="text-lg font-bold text-slate-900 mb-2">لا توجد بيانات مبيعات مستوردة حتى الآن</h2>
        <p className="text-sm text-slate-600 leading-relaxed mb-6">
          تعتمد لوحة المعلومات على قراءة أوراق العمل التسع من ملف Excel الفعلي (VF05N (2)، Zsales_Price (2)، القطاعات، الأهداف، الفروع، الإدارات، المندوبين، الأصناف، والمستخدمين).
        </p>
        {user.role === 'admin' ? (
          <button
            onClick={onNavigateImport}
            className="inline-flex items-center gap-2 px-5 py-2.5 bg-teal-700 hover:bg-teal-800 text-white text-sm font-semibold rounded-lg transition-colors cursor-pointer"
          >
            <Upload className="w-4 h-4" />
            <span>الانتقال إلى استيراد بيانات Excel</span>
          </button>
        ) : (
          <p className="text-xs text-slate-500">يرجى التواصل مع مدير النظام لرفع ملف البيانات الشهري.</p>
        )}
      </div>
    );
  }

  const { kpis, sourceComparison } = data;
  const salesCard = formatMillionOrFull(kpis.privateMarketsSales);
  const targetCard = formatMillionOrFull(kpis.privateMarketsTarget);
  const varianceCard = formatMillionOrFull(kpis.privateMarketsVariance);
  const directArmyCard = formatMillionOrFull(kpis.directAndArmySales);

  const mainAch = kpis.privateMarketsAchievementPct;
  const mainAchPositive = mainAch !== null && mainAch >= 100;
  const variancePositive = kpis.privateMarketsVariance >= 0;
  const directArmyAch = kpis.directAndArmyAchievementPct;
  const directArmyAchPositive = directArmyAch !== null && directArmyAch >= 100;

  const filterMeta = buildActiveFilterMeta({
    source,
    sector: sectorFilter,
    startDate,
    endDate,
  });

  const handleExportExcel = () => {
    const filename = buildExportFilename('لوحة_المعلومات_الرئيسية', filterMeta.fileTags, 'xlsx');
    const sourceRows = sourceComparison
      ? [
          {
            sourceName: 'VF05N (2)',
            count: sourceComparison.vf05n.count,
            salesTotal: sourceComparison.vf05n.salesTotal,
            taxTotal: sourceComparison.vf05n.taxTotal,
          },
          {
            sourceName: 'Zsales_Price (2)',
            count: sourceComparison.zsales.count,
            salesTotal: sourceComparison.zsales.salesTotal,
            taxTotal: sourceComparison.zsales.taxTotal,
          },
        ]
      : [];

    exportToExcel({
      filename,
      sheets: [
        {
          name: 'ملخص المؤشرات',
          title: 'تحليلات قطاع السلاسل — لوحة المعلومات الرئيسية',
          subtitle: filterMeta.summary.join(' | '),
          summaryPairs: [
            { label: 'إجمالي مبيعات الأسواق الخاصة', value: kpis.privateMarketsSales, type: 'currency' },
            { label: 'إجمالي الهدف (الأسواق الخاصة)', value: kpis.privateMarketsTarget, type: 'currency' },
            { label: 'نسبة التحقيق (الأسواق الخاصة)', value: kpis.privateMarketsAchievementPct, type: 'percent' },
            { label: 'الفرق عن الهدف', value: kpis.privateMarketsVariance, type: 'currency' },
            { label: 'البيع المباشر و الجيش', value: kpis.directAndArmySales, type: 'currency' },
            { label: 'نسبة تحقيق البيع المباشر و الجيش', value: kpis.directAndArmyAchievementPct, type: 'percent' },
            { label: 'ضريبة المصدر النشط', value: kpis.totalSourceTax, type: 'currency' },
            { label: 'سجلات غير مرتبطة بفرع', value: kpis.unlinkedBranchCount, type: 'integer' },
            { label: 'سجلات غير مرتبطة بقطاع', value: kpis.unlinkedSectorCount, type: 'integer' },
          ],
          columns: sourceRows.length > 0
            ? [
                { header: 'مصدر البيانات', key: 'sourceName', type: 'string', minWidth: 20 },
                { header: 'عدد السجلات', key: 'count', type: 'integer', minWidth: 14 },
                { header: 'إجمالي المبيعات', key: 'salesTotal', type: 'currency', minWidth: 18 },
                { header: 'إجمالي الضريبة', key: 'taxTotal', type: 'currency', minWidth: 18 },
              ]
            : [],
          rows: sourceRows,
        },
        {
          name: 'التفاصيل حسب القطاع',
          title: 'تحليلات قطاع السلاسل — التفاصيل حسب القطاع',
          subtitle: filterMeta.summary.join(' | '),
          columns: [
            { header: 'القطاع', key: 'label', type: 'string', minWidth: 28 },
            { header: 'المبيعات', key: 'sales', type: 'currency', minWidth: 18 },
            { header: 'الضريبة', key: 'tax', type: 'currency', minWidth: 16 },
            { header: 'الهدف', key: 'target', type: 'currency', minWidth: 18 },
            { header: 'الفرق', key: 'variance', type: 'currency', minWidth: 18 },
            { header: 'نسبة التحقيق', key: 'achievementPct', type: 'percent', minWidth: 15 },
          ],
          rows: tableRows.map((r) => ({
            label: r.label,
            sales: r.sales,
            tax: r.tax,
            target: r.target,
            variance: r.variance,
            achievementPct: r.achievementPct,
          })),
        },
        {
          name: 'حصة القطاعات',
          title: 'تحليلات قطاع السلاسل — حصة القطاعات من المبيعات',
          subtitle: filterMeta.summary.join(' | '),
          columns: [
            { header: 'القطاع', key: 'name', type: 'string', minWidth: 24 },
            { header: 'المبيعات', key: 'value', type: 'currency', minWidth: 18 },
            { header: 'نسبة الحصة', key: 'sharePct', type: 'percent', minWidth: 14 },
          ],
          rows: pieData.map((p) => ({
            name: p.name,
            value: p.value,
            sharePct: p.sharePct,
          })),
        },
      ],
    });
  };

  const handleExportPdf = async () => {
    const filename = buildExportFilename('لوحة_المعلومات_الرئيسية', filterMeta.fileTags, 'pdf');
    const infoBlocks = sourceComparison
      ? [
          {
            title: 'التحقق من تطابق المصادر (دون جمع)',
            items: [
              `VF05N (2): ${formatEGP(sourceComparison.vf05n.salesTotal)} (ضريبة ${formatEGP(sourceComparison.vf05n.taxTotal)} · ${formatNum(sourceComparison.vf05n.count)} سجل)`,
              `Zsales_Price (2): ${formatEGP(sourceComparison.zsales.salesTotal)} (ضريبة ${formatEGP(sourceComparison.zsales.taxTotal)} · ${formatNum(sourceComparison.zsales.count)} سجل)`,
              sourceComparison.isMatched
                ? 'الحالة: المصدران متطابقان بالكامل'
                : `الحالة: يوجد فرق ${formatEGP(sourceComparison.salesDifference)}`,
            ],
          },
        ]
      : [];

    const [pieCapture, barCapture] = await Promise.all([
      pieData.length > 0 ? captureChartSvgToPng(pieChartContainerRef.current, { scale: 3 }) : Promise.resolve(null),
      barData.length > 0 ? captureChartSvgToPng(barChartContainerRef.current, { scale: 3 }) : Promise.resolve(null),
    ]);

    const charts: PdfChartItem[] = [];
    if (pieCapture) {
      charts.push({
        title: 'حصة القطاعات من المبيعات',
        subtitle: `إجمالي القطاعات الظاهرة (${formatNum(pieData.length)})`,
        imageDataUrl: pieCapture.dataUrl,
        width: pieCapture.width,
        height: pieCapture.height,
        legend: pieData.map((entry) => ({
          label: entry.name,
          color: entry.color,
          valueText: `${entry.sharePct.toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}%`,
        })),
      });
    }
    if (barCapture) {
      charts.push({
        title: 'المبيعات مقابل الهدف',
        subtitle: 'بالمليون EGP',
        imageDataUrl: barCapture.dataUrl,
        width: barCapture.width,
        height: barCapture.height,
        legend: [
          { label: 'المبيعات', color: '#0F5A78' },
          { label: 'الهدف', color: '#E06D2B' },
        ],
      });
    }

    await exportToPdf({
      filename,
      tabTitle: 'لوحة المعلومات الرئيسية',
      filtersSummary: filterMeta.summary,
      orientation: 'landscape',
      kpis: [
        {
          label: 'إجمالي مبيعات الأسواق الخاصة',
          value: formatNum(kpis.privateMarketsSales, 2),
          unit: 'EGP',
        },
        {
          label: 'إجمالي الهدف',
          value: formatNum(kpis.privateMarketsTarget, 2),
          unit: 'EGP',
        },
        {
          label: 'نسبة التحقيق',
          value: formatPct(mainAch),
          tone: mainAch === null ? 'default' : mainAchPositive ? 'positive' : 'negative',
        },
        {
          label: 'الفرق عن الهدف',
          value: formatNum(kpis.privateMarketsVariance, 2),
          unit: 'EGP',
          tone: variancePositive ? 'positive' : 'negative',
        },
        {
          label: 'البيع المباشر و الجيش',
          value: formatNum(kpis.directAndArmySales, 2),
          unit: 'EGP',
          sublabel: `من الهدف ${formatPct(directArmyAch)}`,
        },
      ],
      infoBlocks,
      tables: [
        {
          title: 'التفاصيل حسب القطاع',
          columns: [
            { header: 'القطاع', key: 'label', type: 'string', align: 'right' },
            { header: 'المبيعات', key: 'sales', type: 'currency' },
            { header: 'الضريبة', key: 'tax', type: 'currency' },
            { header: 'الهدف', key: 'target', type: 'currency' },
            { header: 'الفرق', key: 'variance', type: 'currency' },
            { header: 'نسبة التحقيق', key: 'achievementPct', type: 'percent' },
          ],
          rows: tableRows.map((r) => ({
            label: r.label,
            sales: r.sales,
            tax: r.tax,
            target: r.target,
            variance: r.variance,
            achievementPct: r.achievementPct,
            _tone: r.rowTone === 'sector' ? undefined : r.rowTone,
          })),
        },
      ],
      chartsSectionTitle: 'الرسوم البيانية التحليلية',
      charts,
    });
  };

  return (
    <div className="space-y-5">
      {/* Unlinked Records Warning Strip */}
      {(kpis.unlinkedBranchCount > 0 || kpis.unlinkedSectorCount > 0) && (
        <div className="bg-amber-50 border border-amber-200 rounded-lg px-4 py-3 flex flex-wrap items-center justify-between gap-4">
          <div className="flex items-start gap-3 text-xs sm:text-sm text-amber-950">
            <AlertTriangle className="w-4 h-4 text-amber-700 shrink-0 mt-0.5" />
            <span>
              <strong className="font-semibold">تنبيه سلامة الربط: </strong>
              يوجد <strong className="font-num">{formatNum(kpis.unlinkedBranchCount)}</strong> سجل غير مرتبط بفرع بقيمة{' '}
              <strong className="font-num">{formatEGP(kpis.unlinkedBranchSales)}</strong>
              {kpis.unlinkedSectorCount > 0 && (
                <>
                  {' '}
                  · و <strong className="font-num">{formatNum(kpis.unlinkedSectorCount)}</strong> سجل غير مرتبط بقطاع بقيمة{' '}
                  <strong className="font-num">{formatEGP(kpis.unlinkedSectorSales)}</strong>
                </>
              )}
              .
            </span>
          </div>
          <button
            onClick={onNavigateBranchesUnlinked}
            className="px-3 py-1.5 text-xs font-semibold text-amber-900 bg-white border border-amber-300 rounded-md hover:bg-amber-100 transition-colors whitespace-nowrap cursor-pointer"
          >
            فحص السجلات غير المرتبطة
          </button>
        </div>
      )}

      {/* Prompt admin if stored legacy batch has 0 tax so they know to re-import */}
      {kpis.totalSourceTax === 0 && kpis.totalSourceSales !== 0 && user.role === 'admin' && (
        <div className="bg-sky-50 border border-sky-200 rounded-lg px-4 py-2.5 flex flex-wrap items-center justify-between gap-3 text-xs text-sky-950">
          <span>
            النسخة الحالية مخزنة بضريبة إجمالية <strong className="font-num">0</strong> (مستوردة قبل ربط عمودي{' '}
            <span className="font-mono">New Total Tax</span> و<span className="font-mono">Total Tax</span>). أعد رفع ملف Excel لتحديث أرقام الضريبة.
          </span>
          <button
            onClick={onNavigateImport}
            className="px-3 py-1.5 text-xs font-semibold text-sky-900 bg-white border border-sky-300 rounded-md hover:bg-sky-100 transition-colors whitespace-nowrap cursor-pointer"
          >
            تحديث البيانات من Excel
          </button>
        </div>
      )}

      {/* Top Row: 5 Equal-Height KPI Cards (Right to Left in RTL) */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3.5 items-stretch">
        {/* Card 1: Total Private Markets Sales */}
        <div
          className="bg-white border border-slate-200 rounded-lg px-4 py-3.5 flex flex-col justify-between text-center min-h-[108px]"
          title={formatEGP(kpis.privateMarketsSales)}
        >
          <span className="text-xs font-semibold text-slate-500">إجمالي مبيعات الأسواق الخاصة</span>
          <div className="text-2xl font-bold text-slate-900 flex items-baseline justify-center gap-1.5 my-auto pt-1">
            <span dir="ltr" className="font-num tabular-nums">
              {salesCard.numText}
            </span>
            <span className="text-base font-bold text-slate-800">{salesCard.unit}</span>
          </div>
        </div>

        {/* Card 2: Total Target */}
        <div
          className="bg-white border border-slate-200 rounded-lg px-4 py-3.5 flex flex-col justify-between text-center min-h-[108px]"
          title={formatEGP(kpis.privateMarketsTarget)}
        >
          <span className="text-xs font-semibold text-slate-500">إجمالي الهدف</span>
          <div className="text-2xl font-bold text-slate-900 flex items-baseline justify-center gap-1.5 my-auto pt-1">
            <span dir="ltr" className="font-num tabular-nums">
              {targetCard.numText}
            </span>
            <span className="text-base font-bold text-slate-800">{targetCard.unit}</span>
          </div>
        </div>

        {/* Card 3: Achievement Percentage */}
        <div className="bg-white border border-slate-200 rounded-lg px-4 py-3.5 flex flex-col justify-between text-center min-h-[108px]">
          <span className="text-xs font-semibold text-slate-500">نسبة التحقيق</span>
          <div className="flex items-center justify-center gap-2.5 my-auto pt-1">
            <span dir="ltr" className="text-2xl font-bold font-num tabular-nums text-slate-900">
              {formatPct(mainAch)}
            </span>
            {mainAch !== null &&
              (mainAchPositive ? (
                <ArrowUp className="w-5 h-5 text-emerald-700 shrink-0" aria-label="فوق المستهدف" />
              ) : (
                <ArrowDown className="w-5 h-5 text-rose-700 shrink-0" aria-label="دون المستهدف" />
              ))}
          </div>
        </div>

        {/* Card 4: Variance from Target */}
        <div
          className="bg-white border border-slate-200 rounded-lg px-4 py-3.5 flex flex-col justify-between text-center min-h-[108px]"
          title={formatEGP(kpis.privateMarketsVariance)}
        >
          <span className="text-xs font-semibold text-slate-500">الفرق عن الهدف</span>
          <div
            className={`text-2xl font-bold flex items-baseline justify-center gap-1.5 my-auto pt-1 ${
              variancePositive ? 'text-emerald-700' : 'text-rose-700'
            }`}
          >
            <span dir="ltr" className="font-num tabular-nums">
              {varianceCard.numText}
            </span>
            <span className="text-base font-bold">{varianceCard.unit}</span>
          </div>
        </div>

        {/* Card 5: Direct Sales & Army */}
        <div
          className="bg-white border border-slate-200 rounded-lg px-4 py-3.5 flex flex-col justify-between text-center min-h-[108px]"
          title={formatEGP(kpis.directAndArmySales)}
        >
          <span className="text-xs font-semibold text-slate-500">البيع المباشر و الجيش</span>
          <div className="my-auto pt-1">
            <div className="text-2xl font-bold text-slate-900 flex items-baseline justify-center gap-1.5">
              <span dir="ltr" className="font-num tabular-nums">
                {directArmyCard.numText}
              </span>
              <span className="text-base font-bold text-slate-800">{directArmyCard.unit}</span>
            </div>
            <div
              className={`mt-1 inline-flex items-center justify-center gap-1 text-xs font-semibold ${
                directArmyAch === null
                  ? 'text-slate-500'
                  : directArmyAchPositive
                    ? 'text-emerald-700'
                    : 'text-rose-700'
              }`}
            >
              <span>من الهدف</span>
              <span dir="ltr" className="font-num tabular-nums">
                {formatPct(directArmyAch)}
              </span>
              {directArmyAch !== null &&
                (directArmyAchPositive ? (
                  <ArrowUp className="w-3.5 h-3.5" />
                ) : (
                  <ArrowDown className="w-3.5 h-3.5" />
                ))}
            </div>
          </div>
        </div>
      </div>

      {/* Main Executive Section: Details by Sector Table (Right) + Pie Chart (Center) + Grouped Bar Chart (Left) */}
      <div className="space-y-2.5">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div className="flex flex-wrap items-center gap-3">
            <h2 className="text-base font-bold text-slate-900">التفاصيل حسب القطاع</h2>
            <span className="text-xs text-slate-500">
              المصدر النشط: <strong className="text-slate-700">{source === 'VF05N' ? 'VF05N (2)' : 'Zsales_Price (2)'}</strong>
            </span>
          </div>
          <ExportButtons
            onExportExcel={handleExportExcel}
            onExportPdf={handleExportPdf}
            disabled={tableRows.length === 0}
          />
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 xl:grid-cols-[minmax(470px,1.45fr)_minmax(240px,0.85fr)_minmax(280px,1fr)] gap-4 items-stretch">
          {/* Right: Sector Details Table */}
          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden flex flex-col lg:col-span-2 xl:col-span-1">
            <div className="overflow-x-auto flex-1">
              <table className="w-full text-xs text-right border-collapse">
                <thead className="bg-slate-100 border-b border-slate-200 text-slate-700 font-bold">
                  <tr>
                    <th className="py-2.5 px-3 whitespace-nowrap">القطاع</th>
                    <th className="py-2.5 px-2.5 whitespace-nowrap">المبيعات</th>
                    <th className="py-2.5 px-2.5 whitespace-nowrap">الضريبة</th>
                    <th className="py-2.5 px-2.5 whitespace-nowrap">الهدف</th>
                    <th className="py-2.5 px-2.5 whitespace-nowrap">الفرق</th>
                    <th className="py-2.5 px-3 whitespace-nowrap">نسبة التحقيق</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {tableRows.map((row) => {
                    const rowBg =
                      row.rowTone === 'private'
                        ? 'bg-sky-50 font-semibold text-slate-900'
                        : row.rowTone === 'direct_army'
                          ? 'bg-emerald-50/80 font-semibold text-slate-900'
                          : row.rowTone === 'chains'
                            ? 'bg-slate-100/90 font-semibold text-slate-900'
                            : 'bg-white hover:bg-slate-50/70 text-slate-800';

                    const achPositive = row.achievementPct !== null && row.achievementPct >= 100;
                    const varPositive = row.variance >= 0;

                    return (
                      <tr key={row.key} className={`${rowBg} transition-colors`}>
                        <td className="py-2.5 px-3 whitespace-nowrap">{row.label}</td>
                        <td className="py-2.5 px-2.5 font-num tabular-nums whitespace-nowrap" dir="ltr" title={formatEGP(row.sales)}>
                          {formatNum(row.sales, 0)}
                        </td>
                        <td
                          className="py-2.5 px-2.5 font-num tabular-nums whitespace-nowrap text-slate-700"
                          dir="ltr"
                          title={formatEGP(row.tax)}
                        >
                          {formatNum(row.tax, 0)}
                        </td>
                        <td className="py-2.5 px-2.5 font-num tabular-nums whitespace-nowrap" dir="ltr" title={formatEGP(row.target)}>
                          {formatNum(row.target, 0)}
                        </td>
                        <td
                          className={`py-2.5 px-2.5 font-num tabular-nums whitespace-nowrap font-semibold ${
                            varPositive ? 'text-emerald-700' : 'text-rose-700'
                          }`}
                          dir="ltr"
                          title={formatEGP(row.variance)}
                        >
                          {formatNum(row.variance, 0)}
                        </td>
                        <td className="py-2.5 px-3 whitespace-nowrap">
                          <div className="inline-flex items-center gap-1 font-num tabular-nums">
                            {row.achievementPct !== null &&
                              (achPositive ? (
                                <ArrowUp className="w-3.5 h-3.5 text-emerald-700 shrink-0" />
                              ) : (
                                <ArrowDown className="w-3.5 h-3.5 text-rose-700 shrink-0" />
                              ))}
                            <span className={achPositive ? 'font-bold text-emerald-700' : 'font-semibold text-slate-800'}>
                              {formatPctOneDecimal(row.achievementPct)}
                            </span>
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>

          {/* Center: Pie Chart — Sector Share of Sales */}
          <div className="bg-white border border-slate-200 rounded-lg p-4 flex flex-col justify-between">
            <div>
              <h3 className="text-sm font-bold text-slate-900 text-center">حصة القطاعات من المبيعات</h3>
            </div>

            {pieData.length === 0 ? (
              <div className="h-56 flex items-center justify-center text-xs text-slate-500">
                لا توجد مبيعات موجبة للعرض
              </div>
            ) : (
              <>
                <div ref={pieChartContainerRef} className="h-56 w-full my-1" dir="ltr">
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie
                        data={pieData}
                        dataKey="value"
                        nameKey="name"
                        cx="50%"
                        cy="50%"
                        outerRadius={82}
                        stroke="#FFFFFF"
                        strokeWidth={2}
                        labelLine={false}
                        label={renderPieSliceLabel}
                      >
                        {pieData.map((entry) => (
                          <Cell key={entry.name} fill={entry.color} />
                        ))}
                      </Pie>
                      <Tooltip
                        formatter={(value, name) => [formatEGP(Number(value ?? 0)), String(name ?? '')]}
                        contentStyle={{
                          borderRadius: '8px',
                          borderColor: '#E2E8F0',
                          fontSize: '12px',
                          direction: 'rtl',
                        }}
                      />
                    </PieChart>
                  </ResponsiveContainer>
                </div>

                {/* Clear RTL Legend */}
                <div className="border-t border-slate-100 pt-2.5 grid grid-cols-2 gap-x-3 gap-y-1.5 text-xs">
                  {pieData.map((entry) => (
                    <div key={entry.name} className="flex items-center gap-1.5 min-w-0">
                      <span
                        className="w-2.5 h-2.5 rounded-xs shrink-0"
                        style={{ backgroundColor: entry.color }}
                        aria-hidden="true"
                      />
                      <span className="text-slate-700 truncate" title={entry.name}>
                        {entry.name}
                      </span>
                    </div>
                  ))}
                </div>
              </>
            )}
          </div>

          {/* Left: Grouped Bar Chart — Sales vs Target (Millions) */}
          <div className="bg-white border border-slate-200 rounded-lg p-4 flex flex-col justify-between">
            <div className="flex items-center justify-between gap-2 mb-1">
              <h3 className="text-sm font-bold text-slate-900">المبيعات مقابل الهدف</h3>
              <span className="text-[11px] text-slate-500">بالمليون EGP</span>
            </div>

            <div ref={barChartContainerRef} className="h-64 w-full" dir="ltr">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={barData} margin={{ top: 10, right: 8, left: -12, bottom: 26 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#E2E8F0" vertical={false} />
                  <XAxis
                    dataKey="sectorName"
                    interval={0}
                    tick={{ fontSize: 10, fill: '#334155' }}
                    angle={-15}
                    textAnchor="end"
                    height={46}
                  />
                  <YAxis tick={{ fontSize: 11, fill: '#64748B' }} />
                  <Tooltip
                    formatter={(value, name) => [`${Number(value ?? 0).toFixed(2)} مليون EGP`, String(name ?? '')]}
                    contentStyle={{
                      borderRadius: '8px',
                      borderColor: '#E2E8F0',
                      fontSize: '12px',
                      direction: 'rtl',
                    }}
                  />
                  <Legend wrapperStyle={{ fontSize: '11px', paddingTop: '2px' }} />
                  <Bar dataKey="salesMillions" name="المبيعات" fill="#0F5A78" radius={[3, 3, 0, 0]} maxBarSize={22} />
                  <Bar dataKey="targetMillions" name="الهدف" fill="#E06D2B" radius={[3, 3, 0, 0]} maxBarSize={22} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>
      </div>

      {/* Compact Source Comparison Verification Strip (VF05N vs Zsales_Price - NEVER Summed) */}
      {sourceComparison && (
        <div className="bg-white border border-slate-200 rounded-lg px-4 py-3.5 flex flex-wrap items-center justify-between gap-4 text-xs">
          <div className="flex flex-wrap items-center gap-4">
            <span className="font-bold text-slate-800">التحقق من تطابق المصادر (دون جمع):</span>
            <span className="text-slate-600">
              VF05N (2): <strong className="font-num text-slate-900">{formatEGP(sourceComparison.vf05n.salesTotal)}</strong>{' '}
              · الضريبة: <strong className="font-num text-slate-800">{formatEGP(sourceComparison.vf05n.taxTotal)}</strong>{' '}
              ({formatNum(sourceComparison.vf05n.count)} سجل)
            </span>
            <span className="text-slate-300" aria-hidden="true">|</span>
            <span className="text-slate-600">
              Zsales_Price (2): <strong className="font-num text-slate-900">{formatEGP(sourceComparison.zsales.salesTotal)}</strong>{' '}
              · الضريبة: <strong className="font-num text-slate-800">{formatEGP(sourceComparison.zsales.taxTotal)}</strong>{' '}
              ({formatNum(sourceComparison.zsales.count)} سجل)
            </span>
          </div>

          <div className="flex items-center gap-1.5 font-semibold">
            {sourceComparison.isMatched ? (
              <span className="inline-flex items-center gap-1.5 text-emerald-700">
                <CheckCircle2 className="w-4 h-4" />
                <span>المصدران متطابقان بالكامل</span>
              </span>
            ) : (
              <span className="inline-flex items-center gap-1.5 text-amber-700">
                <AlertTriangle className="w-4 h-4" />
                <span>
                  الفرق: <span className="font-num">{formatEGP(sourceComparison.salesDifference)}</span>
                </span>
              </span>
            )}
          </div>
        </div>
      )}
    </div>
  );
};
