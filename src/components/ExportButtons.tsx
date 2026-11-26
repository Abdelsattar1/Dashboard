import React, { useState } from 'react';
import { FileSpreadsheet, FileText, Loader2 } from 'lucide-react';

interface ExportButtonsProps {
  onExportExcel: () => void | Promise<void>;
  onExportPdf: () => void | Promise<void>;
  disabled?: boolean;
}

export const ExportButtons: React.FC<ExportButtonsProps> = ({
  onExportExcel,
  onExportPdf,
  disabled = false,
}) => {
  const [busyKind, setBusyKind] = useState<'excel' | 'pdf' | null>(null);
  const [errorText, setErrorText] = useState<string | null>(null);

  const handleExcel = async () => {
    if (disabled || busyKind) return;
    setErrorText(null);
    setBusyKind('excel');
    try {
      await onExportExcel();
    } catch (err) {
      setErrorText(err instanceof Error ? err.message : 'تعذر تصدير ملف Excel');
    } finally {
      setBusyKind(null);
    }
  };

  const handlePdf = async () => {
    if (disabled || busyKind) return;
    setErrorText(null);
    setBusyKind('pdf');
    try {
      await onExportPdf();
    } catch (err) {
      setErrorText(err instanceof Error ? err.message : 'تعذر تصدير ملف PDF');
    } finally {
      setBusyKind(null);
    }
  };

  return (
    <div className="inline-flex flex-col items-end gap-1">
      <div className="inline-flex items-center gap-2">
        <button
          type="button"
          onClick={handleExcel}
          disabled={disabled || busyKind !== null}
          className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold border border-emerald-700/30 bg-emerald-50/70 hover:bg-emerald-100/80 text-emerald-900 rounded-lg transition-colors whitespace-nowrap cursor-pointer disabled:opacity-45 disabled:cursor-not-allowed"
          title="تصدير البيانات الظاهرة إلى ملف Excel (.xlsx)"
        >
          {busyKind === 'excel' ? (
            <Loader2 className="w-3.5 h-3.5 animate-spin shrink-0" />
          ) : (
            <FileSpreadsheet className="w-3.5 h-3.5 shrink-0" />
          )}
          <span>تصدير Excel</span>
        </button>

        <button
          type="button"
          onClick={handlePdf}
          disabled={disabled || busyKind !== null}
          className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold border border-rose-700/30 bg-rose-50/70 hover:bg-rose-100/80 text-rose-900 rounded-lg transition-colors whitespace-nowrap cursor-pointer disabled:opacity-45 disabled:cursor-not-allowed"
          title="تصدير التقرير المنسق إلى ملف PDF"
        >
          {busyKind === 'pdf' ? (
            <Loader2 className="w-3.5 h-3.5 animate-spin shrink-0" />
          ) : (
            <FileText className="w-3.5 h-3.5 shrink-0" />
          )}
          <span>تصدير PDF</span>
        </button>
      </div>
      {errorText && <span className="text-[11px] text-rose-700">{errorText}</span>}
    </div>
  );
};
