import React, { useEffect, useMemo, useState } from 'react';
import { ChevronLeft, ChevronRight, Download, Search } from 'lucide-react';
import type { SalesSource, SalesTransactionRow } from '../types/bi';
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

interface TransactionsViewProps {
  source: SalesSource;
  sectorFilter: string;
  startDate: string;
  endDate: string;
}

const LINK_STATUS_LABELS: Record<
  'all' | 'linked' | 'unlinked_branch' | 'unlinked_sector' | 'returns',
  string
> = {
  all: 'جميع الحالات',
  linked: 'مرتبط بالكامل',
  unlinked_branch: 'غير مرتبط بفرع',
  unlinked_sector: 'غير مرتبط بقطاع',
  returns: 'المرتجعات (قيم سالبة)',
};

export const TransactionsView: React.FC<TransactionsViewProps> = ({
  source,
  sectorFilter,
  startDate,
  endDate,
}) => {
  const [search, setSearch] = useState('');
  const [debouncedSearch, setDebouncedSearch] = useState('');
  const [linkStatus, setLinkStatus] = useState<'all' | 'linked' | 'unlinked_branch' | 'unlinked_sector' | 'returns'>('all');
  const [page, setPage] = useState(1);
  const [rows, setRows] = useState<SalesTransactionRow[]>([]);
  const [totalCount, setTotalCount] = useState(0);
  const [loading, setLoading] = useState(false);
  const pageSize = 50;

  useEffect(() => {
    const t = setTimeout(() => {
      setDebouncedSearch(search);
      setPage(1);
    }, 300);
    return () => clearTimeout(t);
  }, [search]);

  useEffect(() => {
    setPage(1);
  }, [source, sectorFilter, startDate, endDate, linkStatus]);

  useEffect(() => {
    let active = true;
    const load = async () => {
      setLoading(true);
      try {
        const qs = new URLSearchParams({
          source,
          sector: sectorFilter,
          linkStatus,
          page: String(page),
          pageSize: String(pageSize),
        });
        if (startDate) qs.set('startDate', startDate);
        if (endDate) qs.set('endDate', endDate);
        if (debouncedSearch) qs.set('search', debouncedSearch);

        const res = await apiRequest<{
          rows: SalesTransactionRow[];
          totalCount: number;
          page: number;
          pageSize: number;
        }>(`/api/analytics/transactions?${qs.toString()}`);

        if (active) {
          setRows(res.rows);
          setTotalCount(res.totalCount);
        }
      } catch {
        if (active) setRows([]);
      } finally {
        if (active) setLoading(false);
      }
    };
    load();
    return () => {
      active = false;
    };
  }, [source, sectorFilter, startDate, endDate, debouncedSearch, linkStatus, page]);

  const totalPages = Math.max(1, Math.ceil(totalCount / pageSize));

  const pageTotals = useMemo(() => {
    let quantity = 0;
    let salesValue = 0;
    let taxValue = 0;
    for (const r of rows) {
      quantity += r.quantity;
      salesValue += r.salesValue;
      taxValue += r.taxValue;
    }
    return { quantity, salesValue, taxValue };
  }, [rows]);

  const getExportMeta = () => {
    const meta = buildActiveFilterMeta({
      source,
      sector: sectorFilter,
      startDate,
      endDate,
      extra: [
        linkStatus !== 'all'
          ? { label: 'حالة الربط', value: LINK_STATUS_LABELS[linkStatus], fileTag: LINK_STATUS_LABELS[linkStatus] }
          : { label: 'حالة الربط', value: 'جميع الحالات' },
        ...(debouncedSearch.trim() ? [{ label: 'بحث', value: debouncedSearch.trim(), fileTag: debouncedSearch.trim() }] : []),
        { label: 'الصفحة', value: `${page} من ${totalPages}`, fileTag: totalPages > 1 ? `ص${page}` : null },
      ],
    });

    const columns: ExportColumn[] = [
      { header: 'الفاتورة', key: 'docNumber', type: 'string', minWidth: 14 },
      { header: 'التاريخ', key: 'billingDate', type: 'string', minWidth: 12 },
      { header: 'Payer', key: 'payerCode', type: 'string', minWidth: 14 },
      { header: 'القطاع', key: 'resolvedSector', type: 'string', minWidth: 16 },
      { header: 'السلسلة', key: 'chainName', type: 'string', minWidth: 18 },
      { header: 'Sold-to Party', key: 'soldToCode', type: 'string', minWidth: 14 },
      { header: 'الفرع', key: 'branchName', type: 'string', minWidth: 22 },
      { header: 'كود الصنف', key: 'itemCode', type: 'string', minWidth: 14 },
      { header: 'الصنف', key: 'itemName', type: 'string', minWidth: 24 },
      { header: 'الكمية', key: 'quantity', type: 'qty', minWidth: 12 },
      { header: 'المبيعات (BD&T)', key: 'salesValue', type: 'currency', minWidth: 16 },
      { header: 'الضريبة', key: 'taxValue', type: 'currency', minWidth: 14 },
    ];

    const exportRows = rows.map((r) => ({
      docNumber: r.docNumber || '—',
      billingDate: r.billingDate || '—',
      payerCode: r.payerCode || '—',
      resolvedSector: r.resolvedSector || 'غير مرتبط بقطاع',
      chainName: r.resolvedChainName || r.payerName || '—',
      soldToCode: r.soldToCode || '—',
      branchName: r.resolvedBranchName || (r.soldToName ? `غير مرتبط (${r.soldToName})` : 'غير مرتبط'),
      itemCode: r.itemCode || '—',
      itemName: r.itemName || '—',
      quantity: r.quantity,
      salesValue: r.salesValue,
      taxValue: r.taxValue,
    }));

    const totalRow =
      rows.length > 0
        ? {
            docNumber: `إجمالي الصفحة (${formatNum(rows.length)} سجل)`,
            billingDate: '',
            payerCode: '',
            resolvedSector: '',
            chainName: '',
            soldToCode: '',
            branchName: '',
            itemCode: '',
            itemName: '',
            quantity: pageTotals.quantity,
            salesValue: pageTotals.salesValue,
            taxValue: pageTotals.taxValue,
          }
        : null;

    return { meta, columns, exportRows, totalRow };
  };

  const handleExportExcel = () => {
    const { meta, columns, exportRows, totalRow } = getExportMeta();
    exportToExcel({
      filename: buildExportFilename('تفاصيل_المبيعات', meta.fileTags, 'xlsx'),
      sheets: [
        {
          name: 'تفاصيل المبيعات',
          title: `تفاصيل المبيعات (${source === 'VF05N' ? 'VF05N (2)' : 'Zsales_Price (2)'})`,
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
      filename: buildExportFilename('تفاصيل_المبيعات', meta.fileTags, 'pdf'),
      tabTitle: 'تفاصيل المبيعات',
      filtersSummary: meta.summary,
      orientation: 'landscape',
      tables: [
        {
          title: `سجل تفاصيل المبيعات (${source === 'VF05N' ? 'VF05N (2)' : 'Zsales_Price (2)'})`,
          subtitle: `إجمالي السجلات المطابقة: ${formatNum(totalCount)} · معروض في هذه الصفحة: ${formatNum(rows.length)}`,
          columns,
          rows: exportRows,
          totalRow,
        },
      ],
    });
  };

  const exportCsv = () => {
    const headers = [
      'المصدر',
      'رقم الفاتورة',
      'التاريخ',
      'Payer',
      'السلسلة',
      'القطاع',
      'Sold-to Party',
      'الفرع',
      'كود الصنف',
      'الصنف',
      'الكمية',
      'المبيعات',
      'الضريبة',
    ];
    const csvRows = [
      headers.join(','),
      ...rows.map((r) =>
        [
          r.sourceType,
          `"${r.docNumber}"`,
          `"${r.billingDate}"`,
          `"${r.payerCode}"`,
          `"${(r.resolvedChainName || r.payerName).replace(/"/g, '""')}"`,
          `"${r.resolvedSector || 'غير مرتبط'}"`,
          `"${r.soldToCode}"`,
          `"${(r.resolvedBranchName || 'غير مرتبط').replace(/"/g, '""')}"`,
          `"${r.itemCode}"`,
          `"${r.itemName.replace(/"/g, '""')}"`,
          r.quantity,
          r.salesValue,
          r.taxValue,
        ].join(',')
      ),
    ];
    const bom = '\uFEFF';
    const blob = new Blob([bom + csvRows.join('\n')], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `sales-transactions-${source}-page-${page}.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
      <div className="p-5 border-b border-slate-200 flex flex-wrap items-center justify-between gap-4">
        <div>
          <h3 className="text-sm font-bold text-slate-900">سجل تفاصيل المبيعات ({source === 'VF05N' ? 'VF05N (2)' : 'Zsales_Price (2)'})</h3>
          <p className="text-xs text-slate-500 mt-0.5">
            إجمالي السجلات المطابقة: <span className="font-num">{formatNum(totalCount)}</span>
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2.5">
          <select
            value={linkStatus}
            onChange={(e) => setLinkStatus(e.target.value as typeof linkStatus)}
            className="px-3 py-2 text-xs border border-slate-300 rounded-lg bg-white text-slate-700 focus:outline-none focus:ring-2 focus:ring-teal-600"
          >
            <option value="all">جميع الحالات</option>
            <option value="linked">مرتبط بالكامل</option>
            <option value="unlinked_branch">غير مرتبط بفرع</option>
            <option value="unlinked_sector">غير مرتبط بقطاع</option>
            <option value="returns">المرتجعات (قيم سالبة)</option>
          </select>

          <div className="relative">
            <Search className="w-4 h-4 text-slate-400 absolute right-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="بحث برقم الفاتورة / العميل / الصنف..."
              className="pr-9 pl-3 py-2 text-xs border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-teal-600 w-64"
            />
          </div>

          <button
            onClick={exportCsv}
            disabled={rows.length === 0}
            className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold border border-slate-300 rounded-lg text-slate-700 hover:bg-slate-50 disabled:opacity-50 cursor-pointer"
          >
            <Download className="w-3.5 h-3.5" />
            <span>CSV</span>
          </button>

          <ExportButtons
            onExportExcel={handleExportExcel}
            onExportPdf={handleExportPdf}
            disabled={loading || rows.length === 0}
          />
        </div>
      </div>

      <div className="overflow-x-auto">
        <table className="w-full text-xs text-right">
          <thead className="bg-slate-50 border-b border-slate-200 font-semibold text-slate-600">
            <tr>
              <th className="py-3 px-3">الفاتورة</th>
              <th className="py-3 px-3">التاريخ</th>
              <th className="py-3 px-3">Payer</th>
              <th className="py-3 px-3">القطاع / السلسلة</th>
              <th className="py-3 px-3">Sold-to Party</th>
              <th className="py-3 px-3">الفرع</th>
              <th className="py-3 px-3">الصنف</th>
              <th className="py-3 px-3">الكمية</th>
              <th className="py-3 px-3">المبيعات (BD&T)</th>
              <th className="py-3 px-3">الضريبة</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {loading ? (
              <tr>
                <td colSpan={10} className="py-10 text-center text-slate-500">
                  جاري تحميل السجلات...
                </td>
              </tr>
            ) : rows.length === 0 ? (
              <tr>
                <td colSpan={10} className="py-10 text-center text-slate-500">
                  لا توجد سجلات مطابقة لمعايير البحث الحالية
                </td>
              </tr>
            ) : (
              rows.map((r) => (
                <tr key={r.id} className="hover:bg-slate-50/70">
                  <td className="py-2.5 px-3 font-num text-slate-700">{r.docNumber || '—'}</td>
                  <td className="py-2.5 px-3 font-num text-slate-600 whitespace-nowrap">{r.billingDate || '—'}</td>
                  <td className="py-2.5 px-3 font-num text-slate-700">{r.payerCode || '—'}</td>
                  <td className="py-2.5 px-3">
                    <div className="font-medium text-slate-900">{r.resolvedSector || 'غير مرتبط بقطاع'}</div>
                    <div className="text-[11px] text-slate-500">{r.resolvedChainName || r.payerName || '—'}</div>
                  </td>
                  <td className="py-2.5 px-3 font-num text-slate-700">{r.soldToCode || '—'}</td>
                  <td className="py-2.5 px-3">
                    {r.resolvedBranchName ? (
                      <span className="text-slate-900 font-medium">{r.resolvedBranchName}</span>
                    ) : (
                      <span className="text-amber-700 font-medium">غير مرتبط ({r.soldToName || r.soldToCode})</span>
                    )}
                  </td>
                  <td className="py-2.5 px-3">
                    <div className="text-slate-900">{r.itemName || '—'}</div>
                    <div className="font-num text-[11px] text-slate-500">{r.itemCode}</div>
                  </td>
                  <td className="py-2.5 px-3 font-num text-slate-700">{formatNum(r.quantity, 2)}</td>
                  <td className={`py-2.5 px-3 font-num font-semibold whitespace-nowrap ${r.isReturn ? 'text-rose-700' : 'text-slate-900'}`}>
                    {formatEGP(r.salesValue)}
                  </td>
                  <td className="py-2.5 px-3 font-num text-slate-600 whitespace-nowrap">{formatEGP(r.taxValue)}</td>
                </tr>
              ))
            )}
          </tbody>
          {!loading && rows.length > 0 && (
            <tfoot className="bg-sky-50 border-t-2 border-slate-300 font-bold text-slate-900">
              <tr>
                <td colSpan={7} className="py-2.5 px-3 whitespace-nowrap">
                  إجمالي الصفحة ({formatNum(rows.length)} سجل)
                </td>
                <td className="py-2.5 px-3 font-num tabular-nums whitespace-nowrap" dir="ltr">
                  {formatNum(pageTotals.quantity, 2)}
                </td>
                <td className="py-2.5 px-3 font-num tabular-nums whitespace-nowrap" dir="ltr">
                  {formatEGP(pageTotals.salesValue)}
                </td>
                <td className="py-2.5 px-3 font-num tabular-nums whitespace-nowrap" dir="ltr">
                  {formatEGP(pageTotals.taxValue)}
                </td>
              </tr>
            </tfoot>
          )}
        </table>
      </div>

      {/* Pagination */}
      <div className="px-5 py-3.5 border-t border-slate-200 flex items-center justify-between text-xs text-slate-600">
        <div>
          صفحة <span className="font-num font-semibold text-slate-900">{formatNum(page)}</span> من{' '}
          <span className="font-num font-semibold text-slate-900">{formatNum(totalPages)}</span>
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={() => setPage((p) => Math.max(1, p - 1))}
            disabled={page <= 1}
            className="inline-flex items-center gap-1 px-3 py-1.5 border border-slate-300 rounded-lg hover:bg-slate-50 disabled:opacity-40 cursor-pointer"
          >
            <ChevronRight className="w-3.5 h-3.5" />
            <span>السابق</span>
          </button>
          <button
            onClick={() => setPage((p) => Math.min(totalPages, p + 1))}
            disabled={page >= totalPages}
            className="inline-flex items-center gap-1 px-3 py-1.5 border border-slate-300 rounded-lg hover:bg-slate-50 disabled:opacity-40 cursor-pointer"
          >
            <span>التالي</span>
            <ChevronLeft className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>
    </div>
  );
};
