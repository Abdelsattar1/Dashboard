import React, { useEffect, useMemo, useState } from 'react';
import { AlertTriangle, ArrowUpDown, RotateCcw, Search } from 'lucide-react';
import type { BranchMetricRow, SalesSource, UnlinkedRecordRow } from '../types/bi';
import {
  buildActiveFilterMeta,
  buildExportFilename,
  exportToExcel,
  exportToPdf,
  type ExportColumn,
} from '../utils/exporters';
import { formatEGP, formatNum, formatPctOneDecimal } from '../utils/format';
import { ExportButtons } from './ExportButtons';

interface BranchesViewProps {
  branches: BranchMetricRow[];
  unlinkedRecords: UnlinkedRecordRow[];
  totalUnlinkedSales: number;
  totalUnlinkedNetSales?: number;
  totalUnlinkedCount: number;
  loading: boolean;
  initialTab?: 'linked' | 'unlinked';
  authorizedSectors: string[];
  source: SalesSource;
  selectedSector: string;
  startDate: string;
  endDate: string;
  onSectorChange: (sector: string) => void;
}

type SortKey = 'default' | 'chainName' | 'branchName' | 'sales' | 'netSales' | 'tax' | 'target' | 'achievementPct';

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

function normalizeRepKey(code: string): string {
  return String(code || '')
    .replace(/[\u00A0\u200B-\u200F\uFEFF]/g, ' ')
    .trim();
}

export const BranchesView: React.FC<BranchesViewProps> = ({
  branches,
  unlinkedRecords,
  loading,
  initialTab = 'linked',
  authorizedSectors,
  source,
  selectedSector,
  startDate,
  endDate,
  onSectorChange,
}) => {
  const [tab, setTab] = useState<'linked' | 'unlinked'>(initialTab);
  const [supervisorFilter, setSupervisorFilter] = useState<string>('ALL');
  const [chainFilter, setChainFilter] = useState<string>('ALL');
  const [repFilter, setRepFilter] = useState<string>('ALL');
  const [search, setSearch] = useState('');
  const [sortKey, setSortKey] = useState<SortKey>('default');
  const [sortAsc, setSortAsc] = useState(true);

  useEffect(() => {
    setTab(initialTab);
  }, [initialTab]);

  // Normalize global sector value ('UNLINKED' applies only to unlinked records, not linked branches)
  const effectiveSector = selectedSector === 'UNLINKED' ? 'ALL' : selectedSector;

  // Reset dependent child filters whenever parent sector changes
  useEffect(() => {
    setSupervisorFilter('ALL');
    setChainFilter('ALL');
    setRepFilter('ALL');
  }, [effectiveSector]);

  // 1. Branches scoped by Sector
  const branchesBySector = useMemo(() => {
    if (effectiveSector === 'ALL') return branches;
    return branches.filter((b) => b.sectorName === effectiveSector);
  }, [branches, effectiveSector]);

  // Sector filter dropdown options
  const sectorOptions = useMemo(() => {
    const set = new Set<string>(authorizedSectors);
    for (const b of branches) {
      if (b.sectorName) set.add(b.sectorName);
    }
    return Array.from(set);
  }, [authorizedSectors, branches]);

  // 2. Supervisor dropdown options (cascaded from Sector)
  const supervisorOptions = useMemo(() => {
    const map = new Map<string, string>();
    for (const b of branchesBySector) {
      const display = (b.supervisorName || '').trim();
      if (!display) continue;
      const norm = normalizeArabicText(display);
      if (norm && !map.has(norm)) {
        map.set(norm, display);
      }
    }
    return Array.from(map.values()).sort((a, b) => {
      if (a === 'غير محدد') return 1;
      if (b === 'غير محدد') return -1;
      return a.localeCompare(b, 'ar');
    });
  }, [branchesBySector]);

  // Branches scoped by Sector + Supervisor
  const branchesBySupervisor = useMemo(() => {
    if (supervisorFilter === 'ALL') return branchesBySector;
    const targetNorm = normalizeArabicText(supervisorFilter);
    return branchesBySector.filter((b) => normalizeArabicText(b.supervisorName) === targetNorm);
  }, [branchesBySector, supervisorFilter]);

  // 3. Chain Management dropdown options (cascaded from Sector + Supervisor)
  const chainOptions = useMemo(() => {
    const map = new Map<string, string>();
    for (const b of branchesBySupervisor) {
      const display = (b.chainName || '').trim();
      if (!display) continue;
      const norm = normalizeArabicText(display);
      if (norm && !map.has(norm)) {
        map.set(norm, display);
      }
    }
    return Array.from(map.values()).sort((a, b) => {
      if (a === 'غير محدد') return 1;
      if (b === 'غير محدد') return -1;
      return a.localeCompare(b, 'ar');
    });
  }, [branchesBySupervisor]);

  // Branches scoped by Sector + Supervisor + Chain
  const branchesByChain = useMemo(() => {
    if (chainFilter === 'ALL') return branchesBySupervisor;
    const targetNorm = normalizeArabicText(chainFilter);
    return branchesBySupervisor.filter((b) => normalizeArabicText(b.chainName) === targetNorm);
  }, [branchesBySupervisor, chainFilter]);

  // 4. Sales Rep dropdown options (strictly assigned reps from `تعريف الفروع والأهداف` to prevent cross-rep branch inflation)
  const repOptions = useMemo(() => {
    const map = new Map<string, { repCode: string; repName: string }>();
    for (const b of branchesByChain) {
      const assignedKey = normalizeRepKey(b.assignedRepCode);
      if (!assignedKey) continue;
      const existing = map.get(assignedKey);
      if (!existing || (!existing.repName && b.assignedRepName)) {
        map.set(assignedKey, {
          repCode: assignedKey,
          repName: b.assignedRepName || b.assignedRepCode,
        });
      }
    }
    return Array.from(map.values()).sort((a, b) => a.repName.localeCompare(b.repName, 'ar'));
  }, [branchesByChain]);

  // Branches scoped by all 4 interdependent filters + search
  const filteredBranches = useMemo(() => {
    const q = normalizeArabicText(search);
    const list = branchesByChain.filter((b) => {
      if (repFilter !== 'ALL') {
        if (normalizeRepKey(b.assignedRepCode) !== repFilter) return false;
      }
      if (!q) return true;
      return (
        b.branchCode.toLowerCase().includes(q) ||
        normalizeArabicText(b.branchName).includes(q) ||
        normalizeArabicText(b.chainName).includes(q) ||
        normalizeArabicText(b.supervisorName).includes(q) ||
        normalizeArabicText(b.assignedRepName).includes(q)
      );
    });

    const sorted = [...list];
    sorted.sort((a, b) => {
      if (sortKey === 'default') {
        const chainCmp = a.chainName.localeCompare(b.chainName, 'ar');
        if (chainCmp !== 0) return chainCmp;
        return a.branchName.localeCompare(b.branchName, 'ar');
      }
      if (sortKey === 'chainName') {
        const cmp = a.chainName.localeCompare(b.chainName, 'ar');
        return sortAsc ? cmp : -cmp;
      }
      if (sortKey === 'branchName') {
        const cmp = a.branchName.localeCompare(b.branchName, 'ar');
        return sortAsc ? cmp : -cmp;
      }
      if (sortKey === 'achievementPct') {
        const av = a.achievementPct ?? -Infinity;
        const bv = b.achievementPct ?? -Infinity;
        return sortAsc ? av - bv : bv - av;
      }
      const diff = a[sortKey] - b[sortKey];
      return sortAsc ? diff : -diff;
    });

    return sorted;
  }, [branchesByChain, repFilter, search, sortKey, sortAsc]);

  // Dynamic KPIs strictly calculated on filteredBranches with base rule: target > 0
  const kpis = useMemo(() => {
    const branchesWithTarget = filteredBranches.filter((b) => b.target > 0);
    const totalWithTarget = branchesWithTarget.length;
    const activeCount = branchesWithTarget.filter((b) => b.sales > 0).length;
    const inactiveCount = branchesWithTarget.filter((b) => b.sales <= 0).length;
    const efficientCount = branchesWithTarget.filter(
      (b) => b.achievementPct !== null && b.achievementPct >= 70.5
    ).length;

    const activePct = totalWithTarget > 0 ? (activeCount / totalWithTarget) * 100 : 0;
    const inactivePct = totalWithTarget > 0 ? (inactiveCount / totalWithTarget) * 100 : 0;
    const efficiencyPct = totalWithTarget > 0 ? (efficientCount / totalWithTarget) * 100 : 0;

    return {
      totalWithTarget,
      activeCount,
      activePct,
      inactiveCount,
      inactivePct,
      efficientCount,
      efficiencyPct,
    };
  }, [filteredBranches]);

  const tableTotals = useMemo(() => {
    let sales = 0;
    let netSales = 0;
    let tax = 0;
    let target = 0;
    for (const b of filteredBranches) {
      sales += b.sales;
      netSales += b.netSales ?? 0;
      tax += b.tax;
      target += b.target;
    }
    const denom = target - tax;
    const achievementPct = target > 0 && denom > 0 ? (sales / denom) * 100 : null;
    return { sales, netSales, tax, target, achievementPct };
  }, [filteredBranches]);

  const filteredUnlinked = useMemo(() => {
    const q = search.trim().toLowerCase();
    return unlinkedRecords.filter((r) => {
      if (!q) return true;
      return (
        r.soldToCode.toLowerCase().includes(q) ||
        r.soldToName.toLowerCase().includes(q) ||
        r.payerCode.toLowerCase().includes(q) ||
        r.docNumber.toLowerCase().includes(q)
      );
    });
  }, [unlinkedRecords, search]);

  const handleSectorSelect = (val: string) => {
    onSectorChange(val);
    setSupervisorFilter('ALL');
    setChainFilter('ALL');
    setRepFilter('ALL');
  };

  const handleSupervisorSelect = (val: string) => {
    setSupervisorFilter(val);
    setChainFilter('ALL');
    setRepFilter('ALL');
  };

  const handleChainSelect = (val: string) => {
    setChainFilter(val);
    setRepFilter('ALL');
  };

  const handleResetFilters = () => {
    onSectorChange('ALL');
    setSupervisorFilter('ALL');
    setChainFilter('ALL');
    setRepFilter('ALL');
    setSearch('');
  };

  const toggleSort = (k: SortKey) => {
    if (sortKey === k) {
      setSortAsc(!sortAsc);
    } else {
      setSortKey(k);
      setSortAsc(k === 'chainName' || k === 'branchName');
    }
  };

  const hasActiveFilters =
    effectiveSector !== 'ALL' ||
    supervisorFilter !== 'ALL' ||
    chainFilter !== 'ALL' ||
    repFilter !== 'ALL' ||
    search.trim() !== '';

  const hasLegacyMissingBranchSupervisors = useMemo(
    () => branches.length > 0 && branches.every((b) => !b.supervisorName || b.supervisorName === 'غير محدد'),
    [branches]
  );

  const unlinkedTotals = useMemo(() => {
    let sales = 0;
    let netSales = 0;
    let tax = 0;
    for (const r of filteredUnlinked) {
      sales += r.salesValue;
      netSales += r.netSalesValue ?? 0;
      tax += r.taxValue;
    }
    return { sales, netSales, tax };
  }, [filteredUnlinked]);

  const getExportMeta = () => {
    const selectedRepLabel =
      repFilter === 'ALL'
        ? null
        : repOptions.find((r) => r.repCode === repFilter)?.repName || repFilter;

    const meta = buildActiveFilterMeta({
      source,
      sector: effectiveSector,
      startDate,
      endDate,
      extra: [
        {
          label: 'التبويب',
          value: tab === 'linked' ? 'تفصيل تحليل الفروع' : 'السجلات غير المرتبطة بفروع',
          fileTag: tab === 'linked' ? 'الفروع_المرتبطة' : 'غير_مرتبط_بفرع',
        },
        ...(supervisorFilter !== 'ALL' ? [{ label: 'المشرف', value: supervisorFilter, fileTag: supervisorFilter }] : []),
        ...(chainFilter !== 'ALL' ? [{ label: 'إدارة السلسلة', value: chainFilter, fileTag: chainFilter }] : []),
        ...(selectedRepLabel ? [{ label: 'مندوب البيع', value: selectedRepLabel, fileTag: selectedRepLabel }] : []),
        ...(search.trim() ? [{ label: 'بحث', value: search.trim(), fileTag: search.trim() }] : []),
      ],
    });
    return meta;
  };

  const handleExportExcel = () => {
    const meta = getExportMeta();

    if (tab === 'linked') {
      const columns: ExportColumn[] = [
        { header: 'اسم السلسلة', key: 'chainName', type: 'string', minWidth: 22 },
        { header: 'اسم الفرع', key: 'branchName', type: 'string', minWidth: 26 },
        { header: 'إجمالي قيمة المبيعات', key: 'sales', type: 'currency', minWidth: 16 },
        { header: 'صافي قيمة المبيعات', key: 'netSales', type: 'currency', minWidth: 16 },
        { header: 'الضريبة', key: 'tax', type: 'currency', minWidth: 14 },
        { header: 'الهدف', key: 'target', type: 'currency', minWidth: 16 },
        { header: 'ن.التحقيق', key: 'achievementPct', type: 'percent', minWidth: 13 },
      ];
      const exportRows = filteredBranches.map((b) => ({
        chainName: b.chainName,
        branchName: b.branchName,
        sales: b.sales,
        netSales: b.netSales ?? 0,
        tax: b.tax,
        target: b.target,
        achievementPct: b.target > 0 && b.achievementPct !== null ? b.achievementPct : null,
      }));
      const totalRow =
        filteredBranches.length > 0
          ? {
              chainName: 'الإجمالي',
              branchName: `${formatNum(filteredBranches.length)} فرع (${formatNum(kpis.totalWithTarget)} بهدف)`,
              sales: tableTotals.sales,
              netSales: tableTotals.netSales,
              tax: tableTotals.tax,
              target: tableTotals.target,
              achievementPct: tableTotals.achievementPct,
            }
          : null;

      exportToExcel({
        filename: buildExportFilename('الفروع', meta.fileTags, 'xlsx'),
        sheets: [
          {
            name: 'تحليل الفروع',
            title: 'الفروع — تفصيل تحليل الفروع',
            subtitle: meta.summary.join(' | '),
            summaryPairs: [
              { label: 'إجمالي قيمة مبيعات الفروع', value: tableTotals.sales, type: 'currency' },
              { label: 'صافي قيمة مبيعات الفروع', value: tableTotals.netSales, type: 'currency' },
              { label: 'نسبة تحقيق الفروع', value: tableTotals.achievementPct, type: 'percent' },
              { label: 'عدد الفروع النشطة (هدف > 0)', value: kpis.activeCount, type: 'integer' },
              { label: 'نسبة الفروع النشطة', value: kpis.activePct, type: 'percent' },
              { label: 'عدد الفروع غير النشطة', value: kpis.inactiveCount, type: 'integer' },
              { label: 'نسبة الفروع غير النشطة', value: kpis.inactivePct, type: 'percent' },
              { label: 'كفاءة الفروع (>= 70.5%)', value: kpis.efficiencyPct, type: 'percent' },
            ],
            columns,
            rows: exportRows,
            totalRow,
          },
        ],
      });
      return;
    }

    const columns: ExportColumn[] = [
      { header: 'التاريخ', key: 'billingDate', type: 'string', minWidth: 12 },
      { header: 'رقم الفاتورة', key: 'docNumber', type: 'string', minWidth: 14 },
      { header: 'Sold-to Party', key: 'soldToCode', type: 'string', minWidth: 14 },
      { header: 'اسم العميل (Sold-to)', key: 'soldToName', type: 'string', minWidth: 26 },
      { header: 'Payer', key: 'payerCode', type: 'string', minWidth: 14 },
      { header: 'القطاع المرتبط', key: 'resolvedSector', type: 'string', minWidth: 18 },
      { header: 'إجمالي قيمة المبيعات', key: 'salesValue', type: 'currency', minWidth: 16 },
      { header: 'صافي قيمة المبيعات', key: 'netSalesValue', type: 'currency', minWidth: 16 },
      { header: 'الضريبة', key: 'taxValue', type: 'currency', minWidth: 14 },
    ];
    const exportRows = filteredUnlinked.map((r) => ({
      billingDate: r.billingDate || '—',
      docNumber: r.docNumber || '—',
      soldToCode: r.soldToCode || '—',
      soldToName: r.soldToName || '—',
      payerCode: r.payerCode || '—',
      resolvedSector: r.resolvedSector || 'غير مرتبط بقطاع',
      salesValue: r.salesValue,
      netSalesValue: r.netSalesValue ?? 0,
      taxValue: r.taxValue,
    }));
    const totalRow =
      filteredUnlinked.length > 0
        ? {
            billingDate: 'الإجمالي',
            docNumber: `${formatNum(filteredUnlinked.length)} سجل`,
            soldToCode: '',
            soldToName: '',
            payerCode: '',
            resolvedSector: '',
            salesValue: unlinkedTotals.sales,
            netSalesValue: unlinkedTotals.netSales,
            taxValue: unlinkedTotals.tax,
          }
        : null;

    exportToExcel({
      filename: buildExportFilename('الفروع', meta.fileTags, 'xlsx'),
      sheets: [
        {
          name: 'السجلات غير المرتبطة',
          title: 'الفروع — السجلات غير المرتبطة بفروع',
          subtitle: meta.summary.join(' | '),
          columns,
          rows: exportRows,
          totalRow,
        },
      ],
    });
  };

  const handleExportPdf = async () => {
    const meta = getExportMeta();

    if (tab === 'linked') {
      const columns: ExportColumn[] = [
        { header: 'اسم السلسلة', key: 'chainName', type: 'string' },
        { header: 'اسم الفرع', key: 'branchName', type: 'string' },
        { header: 'إجمالي قيمة المبيعات', key: 'sales', type: 'currency' },
        { header: 'صافي قيمة المبيعات', key: 'netSales', type: 'currency' },
        { header: 'الضريبة', key: 'tax', type: 'currency' },
        { header: 'الهدف', key: 'target', type: 'currency' },
        { header: 'ن.التحقيق', key: 'achievementPct', type: 'percent' },
      ];
      const exportRows = filteredBranches.map((b) => ({
        chainName: b.chainName,
        branchName: b.branchName,
        sales: b.sales,
        netSales: b.netSales ?? 0,
        tax: b.tax,
        target: b.target,
        achievementPct: b.target > 0 && b.achievementPct !== null ? b.achievementPct : null,
      }));
      const totalRow =
        filteredBranches.length > 0
          ? {
              chainName: 'الإجمالي',
              branchName: `${formatNum(filteredBranches.length)} فرع (${formatNum(kpis.totalWithTarget)} بهدف)`,
              sales: tableTotals.sales,
              netSales: tableTotals.netSales,
              tax: tableTotals.tax,
              target: tableTotals.target,
              achievementPct: tableTotals.achievementPct,
            }
          : null;

      await exportToPdf({
        filename: buildExportFilename('الفروع', meta.fileTags, 'pdf'),
        tabTitle: 'الفروع — تفصيل تحليل الفروع',
        filtersSummary: meta.summary,
        kpis: [
          {
            label:
              effectiveSector !== 'ALL'
                ? `مبيعات فروع ${effectiveSector}`
                : 'مبيعات الفروع',
            value: formatNum(tableTotals.sales, 2),
            unit: 'EGP',
            sublabel: `صافي قيمة المبيعات: ${formatEGP(tableTotals.netSales)} · الهدف ${formatNum(tableTotals.target, 0)} · التحقيق ${
              tableTotals.achievementPct !== null ? formatPctOneDecimal(tableTotals.achievementPct) : '—'
            }`,
            tone: 'accent',
          },
          {
            label: 'الفروع النشطة',
            value: `${formatNum(kpis.activeCount)} (${formatPctOneDecimal(kpis.activePct)})`,
            sublabel: `من أصل ${formatNum(kpis.totalWithTarget)} فرع له هدف`,
            tone: 'positive',
          },
          {
            label: 'الفروع غير النشطة',
            value: `${formatNum(kpis.inactiveCount)} (${formatPctOneDecimal(kpis.inactivePct)})`,
            sublabel: 'مبيعات صفرية',
            tone: kpis.inactiveCount > 0 ? 'negative' : 'default',
          },
          {
            label: 'كفاءة الفروع (>= 70.5%)',
            value: formatPctOneDecimal(kpis.efficiencyPct),
            sublabel: `${formatNum(kpis.efficientCount)} / ${formatNum(kpis.totalWithTarget)} فرع`,
            tone: 'accent',
          },
        ],
        tables: [
          {
            title: 'تفصيل تحليل الفروع',
            columns,
            rows: exportRows,
            totalRow,
          },
        ],
      });
      return;
    }

    const columns: ExportColumn[] = [
      { header: 'التاريخ', key: 'billingDate', type: 'string' },
      { header: 'رقم الفاتورة', key: 'docNumber', type: 'string' },
      { header: 'Sold-to Party', key: 'soldToCode', type: 'string' },
      { header: 'اسم العميل (Sold-to)', key: 'soldToName', type: 'string' },
      { header: 'Payer', key: 'payerCode', type: 'string' },
      { header: 'القطاع المرتبط', key: 'resolvedSector', type: 'string' },
      { header: 'إجمالي قيمة المبيعات', key: 'salesValue', type: 'currency' },
      { header: 'صافي قيمة المبيعات', key: 'netSalesValue', type: 'currency' },
      { header: 'الضريبة', key: 'taxValue', type: 'currency' },
    ];
    const exportRows = filteredUnlinked.map((r) => ({
      billingDate: r.billingDate || '—',
      docNumber: r.docNumber || '—',
      soldToCode: r.soldToCode || '—',
      soldToName: r.soldToName || '—',
      payerCode: r.payerCode || '—',
      resolvedSector: r.resolvedSector || 'غير مرتبط بقطاع',
      salesValue: r.salesValue,
      netSalesValue: r.netSalesValue ?? 0,
      taxValue: r.taxValue,
    }));
    const totalRow =
      filteredUnlinked.length > 0
        ? {
            billingDate: 'الإجمالي',
            docNumber: `${formatNum(filteredUnlinked.length)} سجل`,
            soldToCode: '',
            soldToName: '',
            payerCode: '',
            resolvedSector: '',
            salesValue: unlinkedTotals.sales,
            netSalesValue: unlinkedTotals.netSales,
            taxValue: unlinkedTotals.tax,
          }
        : null;

    await exportToPdf({
      filename: buildExportFilename('الفروع', meta.fileTags, 'pdf'),
      tabTitle: 'الفروع — السجلات غير المرتبطة',
      filtersSummary: meta.summary,
      infoBlocks: [
        {
          title: 'تنبيه عدم التطابق',
          items: [
            'هذه السجلات لم يطابق كود Sold-to Party الخاص بها أي كود في ورقة تعريف الفروع والأهداف.',
          ],
        },
      ],
      tables: [
        {
          title: 'السجلات غير المرتبطة بفروع',
          columns,
          rows: exportRows,
          totalRow,
        },
      ],
    });
  };

  if (loading && branches.length === 0 && unlinkedRecords.length === 0) {
    return <div className="h-64 bg-white border border-slate-200 rounded-xl animate-pulse" />;
  }

  return (
    <div className="space-y-5">
      {/* Top KPI Cards (Filtered & Computed Dynamically) */}
      <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-4 gap-4">
        {/* Card 1: Total Sales & Net Sales (Branches / Unlinked) */}
        <div
          className="bg-white border border-slate-200 rounded-xl p-5 flex flex-col justify-between"
          title={`${formatEGP(tab === 'linked' ? tableTotals.sales : unlinkedTotals.sales)} | صافي: ${formatEGP(tab === 'linked' ? tableTotals.netSales : unlinkedTotals.netSales)}`}
        >
          <div className="text-xs font-semibold text-slate-500 mb-2">
            {tab === 'unlinked'
              ? 'مبيعات السجلات غير المرتبطة بفروع'
              : effectiveSector !== 'ALL'
                ? `مبيعات فروع ${effectiveSector}`
                : 'مبيعات الفروع'}
          </div>
          <div className="space-y-1.5">
            <div>
              <div className="text-[11px] text-slate-500">إجمالي قيمة المبيعات</div>
              <span className="text-xl font-bold font-num tabular-nums text-slate-900" dir="ltr">
                {formatEGP(tab === 'linked' ? tableTotals.sales : unlinkedTotals.sales)}
              </span>
            </div>
            <div className="pt-1 border-t border-slate-100">
              <div className="text-[11px] text-teal-700">صافي قيمة المبيعات</div>
              <span className="text-base font-bold font-num tabular-nums text-teal-800" dir="ltr">
                {formatEGP(tab === 'linked' ? tableTotals.netSales : unlinkedTotals.netSales)}
              </span>
            </div>
          </div>
          <div className="text-xs text-slate-500 mt-2">
            {tab === 'linked' ? (
              <>
                من <span className="font-num font-semibold text-slate-700">{formatNum(filteredBranches.length)}</span> فرع · التحقيق:{' '}
                <span className="font-num font-semibold text-teal-700" dir="ltr">
                  {tableTotals.achievementPct !== null ? formatPctOneDecimal(tableTotals.achievementPct) : '—'}
                </span>
              </>
            ) : (
              <>
                عبر <span className="font-num font-semibold text-slate-700">{formatNum(filteredUnlinked.length)}</span> سجل · الضريبة:{' '}
                <span className="font-num font-semibold text-slate-700" dir="ltr">
                  {formatEGP(unlinkedTotals.tax)}
                </span>
              </>
            )}
          </div>
        </div>

        {/* Card 2: Active Branches */}
        <div className="bg-white border border-slate-200 rounded-xl p-5 flex flex-col justify-between">
          <div className="text-xs font-semibold text-slate-500 mb-2">
            عدد الفروع النشطة ونسبتها إلى إجمالي عدد الفروع
          </div>
          <div className="flex items-baseline justify-between gap-3">
            <span className="text-2xl font-bold font-num tabular-nums text-emerald-700" dir="ltr">
              {formatNum(kpis.activeCount)}
            </span>
            <span className="text-lg font-bold font-num tabular-nums text-teal-800" dir="ltr">
              {formatPctOneDecimal(kpis.activePct)}
            </span>
          </div>
          <div className="text-xs text-slate-500 mt-2">
            من إجمالي <span className="font-num font-semibold text-slate-700">{formatNum(kpis.totalWithTarget)}</span> فرع له هدف (&gt; 0)
          </div>
        </div>

        {/* Card 3: Inactive Branches */}
        <div className="bg-white border border-slate-200 rounded-xl p-5 flex flex-col justify-between">
          <div className="text-xs font-semibold text-slate-500 mb-2">
            عدد الفروع الغير نشطة ونسبتها إلى إجمالي عدد الفروع
          </div>
          <div className="flex items-baseline justify-between gap-3">
            <span className="text-2xl font-bold font-num tabular-nums text-rose-700" dir="ltr">
              {formatNum(kpis.inactiveCount)}
            </span>
            <span className="text-lg font-bold font-num tabular-nums text-rose-700" dir="ltr">
              {formatPctOneDecimal(kpis.inactivePct)}
            </span>
          </div>
          <div className="text-xs text-slate-500 mt-2">
            مبيعات صفرية من أصل <span className="font-num font-semibold text-slate-700">{formatNum(kpis.totalWithTarget)}</span> فرع له هدف
          </div>
        </div>

        {/* Card 4: Branch Efficiency */}
        <div className="bg-white border border-slate-200 rounded-xl p-5 flex flex-col justify-between">
          <div className="text-xs font-semibold text-slate-500 mb-2">كفاءة الفروع</div>
          <div className="flex items-baseline justify-between gap-3">
            <span className="text-2xl font-bold font-num tabular-nums text-teal-800" dir="ltr">
              {formatPctOneDecimal(kpis.efficiencyPct)}
            </span>
            <span className="text-sm font-semibold font-num tabular-nums text-slate-700" dir="ltr">
              {formatNum(kpis.efficientCount)} / {formatNum(kpis.totalWithTarget)}
            </span>
          </div>
          <div className="text-xs text-slate-500 mt-2">
            الفروع المحققة <span className="font-num font-semibold text-slate-700" dir="ltr">&ge; 70.5%</span> من إجمالي الفروع ذات الهدف
          </div>
        </div>
      </div>

      {/* 4 Interdependent Cascading Filters */}
      <div className="bg-white border border-slate-200 rounded-xl p-4">
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
          <div>
            <label className="block text-[11px] font-semibold text-slate-600 mb-1">القطاع</label>
            <select
              value={effectiveSector}
              onChange={(e) => handleSectorSelect(e.target.value)}
              className="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg bg-white text-slate-800 focus:outline-none focus:ring-2 focus:ring-teal-600"
            >
              <option value="ALL">جميع القطاعات (الكل)</option>
              {sectorOptions.map((sec) => (
                <option key={sec} value={sec}>
                  {sec}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className="block text-[11px] font-semibold text-slate-600 mb-1">المشرف</label>
            <select
              value={supervisorFilter}
              onChange={(e) => handleSupervisorSelect(e.target.value)}
              className="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg bg-white text-slate-800 focus:outline-none focus:ring-2 focus:ring-teal-600"
            >
              <option value="ALL">جميع المشرفين ({formatNum(supervisorOptions.length)})</option>
              {supervisorOptions.map((sup) => (
                <option key={sup} value={sup}>
                  {sup}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className="block text-[11px] font-semibold text-slate-600 mb-1">إدارة السلسلة</label>
            <select
              value={chainFilter}
              onChange={(e) => handleChainSelect(e.target.value)}
              className="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg bg-white text-slate-800 focus:outline-none focus:ring-2 focus:ring-teal-600"
            >
              <option value="ALL">جميع السلاسل ({formatNum(chainOptions.length)})</option>
              {chainOptions.map((ch) => (
                <option key={ch} value={ch}>
                  {ch}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className="block text-[11px] font-semibold text-slate-600 mb-1">مندوب البيع</label>
            <select
              value={repFilter}
              onChange={(e) => setRepFilter(e.target.value)}
              className="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg bg-white text-slate-800 focus:outline-none focus:ring-2 focus:ring-teal-600"
            >
              <option value="ALL">جميع المناديب ({formatNum(repOptions.length)})</option>
              {repOptions.map((r) => (
                <option key={r.repCode} value={r.repCode}>
                  {r.repName !== r.repCode ? `${r.repName} (${r.repCode})` : r.repCode}
                </option>
              ))}
            </select>
          </div>
        </div>

        {hasLegacyMissingBranchSupervisors && (
          <p className="text-[11px] text-sky-800 mt-2.5">
            ملاحظة: النسخة الحالية مستوردة قبل تفعيل قراءة عمود «اسم المشرف» وكود/اسم الإدارة من ورقة «تعريف الفروع والأهداف» — يرجى إعادة رفع ملف Excel لاعتماد تبعية الفروع والمشرفين بدقة.
          </p>
        )}

        {hasActiveFilters && (
          <div className="mt-3 pt-2.5 border-t border-slate-100 flex items-center justify-between text-xs">
            <span className="text-slate-500">
              عرض <span className="font-num font-semibold text-slate-800">{formatNum(filteredBranches.length)}</span> فرع مطابق للفلترة الحالية
            </span>
            <button
              onClick={handleResetFilters}
              className="inline-flex items-center gap-1.5 text-teal-700 hover:text-teal-900 font-semibold cursor-pointer"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              <span>إعادة ضبط الفلاتر</span>
            </button>
          </div>
        )}
      </div>

      {/* Branches Analysis Table & Unlinked Records Tab */}
      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="p-4 lg:p-5 border-b border-slate-200 flex flex-wrap items-center justify-between gap-4">
          <div className="flex items-center gap-1 p-1 bg-slate-100 rounded-lg">
            <button
              onClick={() => setTab('linked')}
              className={`px-3.5 py-1.5 text-xs font-semibold rounded-md transition-colors whitespace-nowrap cursor-pointer ${
                tab === 'linked' ? 'bg-white text-slate-900 shadow-sm' : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              تفصيل تحليل الفروع ({formatNum(filteredBranches.length)})
            </button>
            <button
              onClick={() => setTab('unlinked')}
              className={`px-3.5 py-1.5 text-xs font-semibold rounded-md transition-colors whitespace-nowrap cursor-pointer ${
                tab === 'unlinked' ? 'bg-white text-amber-800 shadow-sm' : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              السجلات غير المرتبطة ({formatNum(unlinkedRecords.length)})
            </button>
          </div>

          <div className="flex flex-wrap items-center gap-2.5">
            <div className="relative">
              <Search className="w-4 h-4 text-slate-400 absolute right-3 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder={tab === 'linked' ? 'بحث بالسلسلة أو الفرع أو المندوب...' : 'بحث بكود العميل أو رقم الفاتورة...'}
                className="pr-9 pl-3 py-1.5 text-xs border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-teal-600 w-64"
              />
            </div>

            <ExportButtons
              onExportExcel={handleExportExcel}
              onExportPdf={handleExportPdf}
              disabled={loading || (tab === 'linked' ? filteredBranches.length === 0 : filteredUnlinked.length === 0)}
            />
          </div>
        </div>

        {tab === 'linked' ? (
          <div className="max-h-[560px] overflow-auto">
            <table className="w-full text-sm text-right border-collapse">
              <thead className="sticky top-0 z-10 bg-slate-100 border-b-2 border-slate-300 text-xs font-bold text-slate-800 shadow-[0_1px_0_0_rgba(203,213,225,1)]">
                <tr>
                  <th className="py-2.5 px-4 whitespace-nowrap border-l border-slate-200">
                    <button
                      onClick={() => toggleSort('chainName')}
                      className="inline-flex items-center gap-1 cursor-pointer hover:text-teal-800 underline decoration-slate-400 underline-offset-4"
                    >
                      <span>اسم السلسلة</span>
                      <ArrowUpDown className="w-3 h-3 text-slate-400" />
                    </button>
                  </th>
                  <th className="py-2.5 px-4 whitespace-nowrap">
                    <button
                      onClick={() => toggleSort('branchName')}
                      className="inline-flex items-center gap-1 cursor-pointer hover:text-teal-800 underline decoration-slate-400 underline-offset-4"
                    >
                      <span>اسم الفرع</span>
                      <ArrowUpDown className="w-3 h-3 text-slate-400" />
                    </button>
                  </th>
                  <th className="py-2.5 px-4 whitespace-nowrap text-center">
                    <button
                      onClick={() => toggleSort('sales')}
                      className="inline-flex items-center gap-1 cursor-pointer hover:text-teal-800 underline decoration-slate-400 underline-offset-4"
                    >
                      <span>إجمالي قيمة المبيعات</span>
                      <ArrowUpDown className="w-3 h-3 text-slate-400" />
                    </button>
                  </th>
                  <th className="py-2.5 px-4 whitespace-nowrap text-center">
                    <button
                      onClick={() => toggleSort('netSales')}
                      className="inline-flex items-center gap-1 cursor-pointer hover:text-teal-800 underline decoration-slate-400 underline-offset-4"
                    >
                      <span>صافي قيمة المبيعات</span>
                      <ArrowUpDown className="w-3 h-3 text-slate-400" />
                    </button>
                  </th>
                  <th className="py-2.5 px-4 whitespace-nowrap text-center">
                    <button
                      onClick={() => toggleSort('tax')}
                      className="inline-flex items-center gap-1 cursor-pointer hover:text-teal-800 underline decoration-slate-400 underline-offset-4"
                    >
                      <span>الضريبة</span>
                      <ArrowUpDown className="w-3 h-3 text-slate-400" />
                    </button>
                  </th>
                  <th className="py-2.5 px-4 whitespace-nowrap text-center">
                    <button
                      onClick={() => toggleSort('target')}
                      className="inline-flex items-center gap-1 cursor-pointer hover:text-teal-800 underline decoration-slate-400 underline-offset-4"
                    >
                      <span>الهدف</span>
                      <ArrowUpDown className="w-3 h-3 text-slate-400" />
                    </button>
                  </th>
                  <th className="py-2.5 px-4 whitespace-nowrap text-center">
                    <button
                      onClick={() => toggleSort('achievementPct')}
                      className="inline-flex items-center gap-1 cursor-pointer hover:text-teal-800 underline decoration-slate-400 underline-offset-4"
                    >
                      <span>ن.التحقيق</span>
                      <ArrowUpDown className="w-3 h-3 text-slate-400" />
                    </button>
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {filteredBranches.length === 0 ? (
                  <tr>
                    <td colSpan={7} className="py-10 text-center text-xs text-slate-500">
                      لا توجد فروع مطابقة للفلاتر المحددة
                    </td>
                  </tr>
                ) : (
                  filteredBranches.map((b) => {
                    const showPct = b.target > 0 && b.achievementPct !== null;
                    const roundedPct = showPct ? Math.round(b.achievementPct!) : null;
                    return (
                      <tr key={`${b.sectorName}-${b.supervisorName}-${b.assignedRepCode}-${b.branchCode}`} className="hover:bg-slate-50/80 transition-colors">
                        <td
                          className="py-2 px-4 whitespace-nowrap font-semibold text-slate-900 border-l border-slate-200"
                          title={`المشرف: ${b.supervisorName} | القطاع: ${b.sectorName}`}
                        >
                          {b.chainName}
                        </td>
                        <td
                          className="py-2 px-4 whitespace-nowrap text-slate-800"
                          title={`كود الفرع: ${b.branchCode}${b.assignedRepName ? ` | المندوب: ${b.assignedRepName}` : ''}`}
                        >
                          {b.branchName}
                        </td>
                        <td
                          className={`py-2 px-4 text-center font-num tabular-nums whitespace-nowrap ${
                            b.sales < 0 ? 'text-rose-700 font-semibold' : 'text-slate-900'
                          }`}
                          dir="ltr"
                          title={formatEGP(b.sales)}
                        >
                          {formatNum(b.sales, 0)}
                        </td>
                        <td
                          className={`py-2 px-4 text-center font-num tabular-nums whitespace-nowrap font-medium ${
                            (b.netSales ?? 0) < 0 ? 'text-rose-700 font-semibold' : 'text-teal-800'
                          }`}
                          dir="ltr"
                          title={formatEGP(b.netSales ?? 0)}
                        >
                          {formatNum(b.netSales ?? 0, 0)}
                        </td>
                        <td
                          className={`py-2 px-4 text-center font-num tabular-nums whitespace-nowrap ${
                            b.tax < 0 ? 'text-rose-700' : 'text-slate-700'
                          }`}
                          dir="ltr"
                          title={formatEGP(b.tax)}
                        >
                          {formatNum(b.tax, 0)}
                        </td>
                        <td
                          className="py-2 px-4 text-center font-num tabular-nums whitespace-nowrap text-slate-800"
                          dir="ltr"
                          title={formatEGP(b.target)}
                        >
                          {formatNum(b.target, 0)}
                        </td>
                        <td
                          className={`py-2 px-4 text-center font-num tabular-nums whitespace-nowrap font-semibold ${
                            roundedPct === null
                              ? ''
                              : b.achievementPct! >= 70.5
                                ? 'text-emerald-700'
                                : 'text-slate-800'
                          }`}
                          dir="ltr"
                        >
                          {roundedPct !== null ? `${roundedPct}%` : ''}
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
              {filteredBranches.length > 0 && (
                <tfoot className="bg-sky-50 border-t-2 border-slate-400 font-bold text-slate-900">
                  <tr>
                    <td className="py-2.5 px-4 whitespace-nowrap border-l border-slate-200">الإجمالي</td>
                    <td className="py-2.5 px-4 whitespace-nowrap text-xs font-normal text-slate-600">
                      {formatNum(filteredBranches.length)} فرع ({formatNum(kpis.totalWithTarget)} بهدف)
                    </td>
                    <td className="py-2.5 px-4 text-center font-num tabular-nums whitespace-nowrap" dir="ltr" title={formatEGP(tableTotals.sales)}>
                      {formatNum(tableTotals.sales, 0)}
                    </td>
                    <td className="py-2.5 px-4 text-center font-num tabular-nums whitespace-nowrap text-teal-800" dir="ltr" title={formatEGP(tableTotals.netSales)}>
                      {formatNum(tableTotals.netSales, 0)}
                    </td>
                    <td className="py-2.5 px-4 text-center font-num tabular-nums whitespace-nowrap" dir="ltr" title={formatEGP(tableTotals.tax)}>
                      {formatNum(tableTotals.tax, 0)}
                    </td>
                    <td className="py-2.5 px-4 text-center font-num tabular-nums whitespace-nowrap" dir="ltr" title={formatEGP(tableTotals.target)}>
                      {formatNum(tableTotals.target, 0)}
                    </td>
                    <td className="py-2.5 px-4 text-center font-num tabular-nums whitespace-nowrap text-teal-900" dir="ltr">
                      {tableTotals.achievementPct !== null ? `${Math.round(tableTotals.achievementPct)}%` : ''}
                    </td>
                  </tr>
                </tfoot>
              )}
            </table>
          </div>
        ) : (
          <div>
            <div className="px-5 py-3 bg-amber-50 border-b border-amber-200 text-xs text-amber-900 flex items-center gap-2">
              <AlertTriangle className="w-4 h-4 text-amber-700 shrink-0" />
              <span>
                هذه السجلات لم يطابق كود <span className="font-num font-semibold">Sold-to Party</span> الخاص بها أي كود في ورقة تعريف الفروع والأهداف، ولذلك لم يتم إسنادها لأي فرع بالتخمين.
              </span>
            </div>
            <div className="overflow-x-auto max-h-[520px]">
              <table className="w-full text-sm text-right">
                <thead className="sticky top-0 bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-600">
                  <tr>
                    <th className="py-3 px-4">التاريخ</th>
                    <th className="py-3 px-4">رقم الفاتورة</th>
                    <th className="py-3 px-4">Sold-to Party</th>
                    <th className="py-3 px-4">اسم العميل (Sold-to)</th>
                    <th className="py-3 px-4">Payer</th>
                    <th className="py-3 px-4">القطاع المرتبط</th>
                    <th className="py-3 px-4">إجمالي قيمة المبيعات</th>
                    <th className="py-3 px-4">صافي قيمة المبيعات</th>
                    <th className="py-3 px-4">الضريبة</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {filteredUnlinked.length === 0 ? (
                    <tr>
                      <td colSpan={9} className="py-8 text-center text-xs text-slate-500">
                        جميع السجلات مرتبطة بفروع بنجاح
                      </td>
                    </tr>
                  ) : (
                    filteredUnlinked.map((r) => (
                      <tr key={r.id} className="hover:bg-slate-50/70">
                        <td className="py-2.5 px-4 font-num text-xs text-slate-600">{r.billingDate || '—'}</td>
                        <td className="py-2.5 px-4 font-num text-xs text-slate-700">{r.docNumber || '—'}</td>
                        <td className="py-2.5 px-4 font-num text-xs font-semibold text-amber-800">{r.soldToCode || '—'}</td>
                        <td className="py-2.5 px-4 text-slate-900">{r.soldToName || '—'}</td>
                        <td className="py-2.5 px-4 font-num text-xs text-slate-600">{r.payerCode || '—'}</td>
                        <td className="py-2.5 px-4 text-xs text-slate-700">{r.resolvedSector || 'غير مرتبط بقطاع'}</td>
                        <td className="py-2.5 px-4 font-num font-semibold text-slate-900">{formatEGP(r.salesValue)}</td>
                        <td className="py-2.5 px-4 font-num font-semibold text-teal-800">{formatEGP(r.netSalesValue ?? 0)}</td>
                        <td className="py-2.5 px-4 font-num text-slate-600">{formatEGP(r.taxValue)}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};
