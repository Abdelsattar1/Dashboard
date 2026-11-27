import React, { useCallback, useEffect, useMemo, useState } from 'react';
import {
  AlertCircle,
  AlertTriangle,
  Check,
  CheckCircle2,
  ClipboardCopy,
  Clock,
  History,
  Info,
  RefreshCw,
  RotateCcw,
  ShieldCheck,
  Sparkles,
  Wrench,
} from 'lucide-react';
import type { AuditFinding, AuditSeverity, AutoFixActionType, SmartAuditReport } from '../types/bi';
import { apiRequest } from '../utils/api';
import { formatDateTime } from '../utils/format';

interface SmartAuditorViewProps {
  onDataMutated: () => void;
}

const SEVERITY_META: Record<
  AuditSeverity,
  { label: string; badgeClass: string; borderClass: string }
> = {
  critical: {
    label: 'حرج',
    badgeClass: 'bg-rose-100 text-rose-800 border-rose-200',
    borderClass: 'border-rose-200',
  },
  high: {
    label: 'عالي',
    badgeClass: 'bg-amber-100 text-amber-800 border-amber-200',
    borderClass: 'border-amber-200',
  },
  medium: {
    label: 'متوسط',
    badgeClass: 'bg-sky-100 text-sky-800 border-sky-200',
    borderClass: 'border-sky-200',
  },
  low: {
    label: 'منخفض',
    badgeClass: 'bg-slate-100 text-slate-700 border-slate-200',
    borderClass: 'border-slate-200',
  },
};

export const SmartAuditorView: React.FC<SmartAuditorViewProps> = ({ onDataMutated }) => {
  const [report, setReport] = useState<SmartAuditReport | null>(null);
  const [loading, setLoading] = useState(true);
  const [running, setRunning] = useState(false);
  const [aiAnalyzing, setAiAnalyzing] = useState(false);
  const [busyKey, setBusyKey] = useState<string | null>(null);
  const [autoScanEnabled, setAutoScanEnabled] = useState(true);
  const [severityFilter, setSeverityFilter] = useState<'ALL' | AuditSeverity>('ALL');
  const [fixTypeFilter, setFixTypeFilter] = useState<'ALL' | 'auto_fixable' | 'code_fix_required'>('ALL');
  const [copiedId, setCopiedId] = useState<string | null>(null);
  const [feedback, setFeedback] = useState<{ type: 'success' | 'error'; text: string } | null>(null);

  const runAudit = useCallback(async (isInitial = false) => {
    if (isInitial) {
      setLoading(true);
    } else {
      setRunning(true);
    }
    try {
      const res = await apiRequest<{ report: SmartAuditReport }>('/api/admin/audit/run', {
        method: 'POST',
      });
      setReport(res.report);
    } catch (err) {
      setFeedback({
        type: 'error',
        text: err instanceof Error ? err.message : 'تعذر تشغيل الفحص الذكي للنظام',
      });
    } finally {
      setLoading(false);
      setRunning(false);
    }
  }, []);

  useEffect(() => {
    runAudit(true);
  }, [runAudit]);

  useEffect(() => {
    if (!autoScanEnabled) return;
    const timer = window.setInterval(() => {
      runAudit(false);
    }, 120_000);
    return () => window.clearInterval(timer);
  }, [autoScanEnabled, runAudit]);

  const handleAiDeepAnalysis = async () => {
    setAiAnalyzing(true);
    setFeedback(null);
    try {
      const res = await apiRequest<{ usedAi: boolean; aiSummary: string; report: SmartAuditReport }>(
        '/api/admin/audit/ai-analyze',
        { method: 'POST' }
      );
      setReport(res.report);
      setFeedback({
        type: 'success',
        text: res.usedAi
          ? 'تم إثراء التقرير بتحليل الذكاء الاصطناعي وتحديث توصيات الإصلاح.'
          : 'تم التحقق باستخدام محرك القواعد المحلي.',
      });
    } catch (err) {
      setFeedback({
        type: 'error',
        text: err instanceof Error ? err.message : 'تعذر تنفيذ التحليل الإضافي للنظام',
      });
    } finally {
      setAiAnalyzing(false);
    }
  };

  const handleAutoFix = async (finding: AuditFinding) => {
    if (!finding.autoFixAction) return;
    setBusyKey(`fix_${finding.id}`);
    setFeedback(null);
    try {
      const res = await apiRequest<{ backupId: number; summary: string; report: SmartAuditReport }>(
        '/api/admin/audit/autofix',
        {
          method: 'POST',
          body: JSON.stringify({
            actionType: finding.autoFixAction as AutoFixActionType,
            findingCode: finding.code,
          }),
        }
      );
      setReport(res.report);
      setFeedback({
        type: 'success',
        text: `${res.summary} (تم إنشاء نسخة احتياطية #${res.backupId} قبل التعديل)`,
      });
      onDataMutated();
    } catch (err) {
      setFeedback({
        type: 'error',
        text: err instanceof Error ? err.message : 'تعذر تنفيذ الإصلاح التلقائي',
      });
    } finally {
      setBusyKey(null);
    }
  };

  const handleRollback = async (backupId: number) => {
    setBusyKey(`rollback_${backupId}`);
    setFeedback(null);
    try {
      const res = await apiRequest<{ summary: string; report: SmartAuditReport }>(
        `/api/admin/audit/rollback/${backupId}`,
        { method: 'POST' }
      );
      setReport(res.report);
      setFeedback({
        type: 'success',
        text: res.summary,
      });
      onDataMutated();
    } catch (err) {
      setFeedback({
        type: 'error',
        text: err instanceof Error ? err.message : 'تعذر استعادة النسخة الاحتياطية',
      });
    } finally {
      setBusyKey(null);
    }
  };

  const handleCopyPrompt = async (key: string, text: string) => {
    try {
      await navigator.clipboard.writeText(text);
      setCopiedId(key);
      window.setTimeout(() => {
        setCopiedId((prev) => (prev === key ? null : prev));
      }, 2500);
    } catch {
      setFeedback({
        type: 'error',
        text: 'تعذر النسخ التلقائي إلى الحافظة؛ يمكنك تحديد النص ونسخه يدوياً.',
      });
    }
  };

  const filteredFindings = useMemo(() => {
    if (!report) return [];
    return report.findings.filter((f) => {
      if (severityFilter !== 'ALL' && f.severity !== severityFilter) return false;
      if (fixTypeFilter !== 'ALL' && f.fixType !== fixTypeFilter) return false;
      return true;
    });
  }, [report, severityFilter, fixTypeFilter]);

  if (loading && !report) {
    return (
      <div className="bg-white border border-slate-200 rounded-xl p-10 text-center text-sm text-slate-600">
        جاري تشغيل الفحص الذكي الشامل لقواعد البيانات والحسابات والصلاحيات...
      </div>
    );
  }

  return (
    <div className="space-y-6">
      {/* Header Toolbar */}
      <div className="bg-white border border-slate-200 rounded-xl p-5 flex flex-wrap items-center justify-between gap-4">
        <div className="space-y-1">
          <div className="flex items-center gap-2.5">
            <ShieldCheck className="w-5 h-5 text-teal-700 shrink-0" />
            <h2 className="text-base font-bold text-slate-900">المراجع الذكي وفحص سلامة النظام</h2>
          </div>
          <p className="text-xs text-slate-500">
            أداة حصرية لمدير النظام لمراقبة أخطاء التشغيل، القيم غير الصالحة (NaN / Infinity)، تناسق القطاعات، وصلاحيات الأدوار مع إصلاح تلقائي آمن ونسخ احتياطي فوري.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2.5">
          <label className="inline-flex items-center gap-2 px-3 py-2 border border-slate-200 rounded-lg text-xs font-semibold text-slate-700 bg-slate-50 cursor-pointer">
            <input
              type="checkbox"
              checked={autoScanEnabled}
              onChange={(e) => setAutoScanEnabled(e.target.checked)}
              className="rounded border-slate-300 text-teal-700 focus:ring-teal-600"
            />
            <span>فحص دوري كل دقيقتين</span>
          </label>

          {report?.aiConfigured && (
            <button
              type="button"
              onClick={handleAiDeepAnalysis}
              disabled={aiAnalyzing || running}
              className="inline-flex items-center gap-2 px-3.5 py-2 text-xs font-semibold border border-teal-200 bg-teal-50 text-teal-800 hover:bg-teal-100 rounded-lg transition-colors cursor-pointer disabled:opacity-50"
            >
              <Sparkles className="w-3.5 h-3.5" />
              <span>{aiAnalyzing ? 'جاري التحليل الذكي...' : 'تحليل ذكي إضافي (AI)'}</span>
            </button>
          )}

          <button
            type="button"
            onClick={() => runAudit(false)}
            disabled={running}
            className="inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold text-white bg-teal-700 hover:bg-teal-800 rounded-lg transition-colors cursor-pointer disabled:opacity-50"
          >
            <RefreshCw className={`w-3.5 h-3.5 ${running ? 'animate-spin' : ''}`} />
            <span>{running ? 'جاري الفحص...' : 'تشغيل الفحص الذكي'}</span>
          </button>
        </div>
      </div>

      {feedback && (
        <div
          className={`p-3.5 rounded-xl border text-xs font-semibold flex items-center justify-between gap-3 ${
            feedback.type === 'success'
              ? 'bg-emerald-50 border-emerald-200 text-emerald-800'
              : 'bg-rose-50 border-rose-200 text-rose-800'
          }`}
        >
          <span>{feedback.text}</span>
          <button
            type="button"
            onClick={() => setFeedback(null)}
            className="text-[11px] underline cursor-pointer shrink-0"
          >
            إخفاء
          </button>
        </div>
      )}

      {report && (
        <>
          {/* Summary KPI Cards */}
          <div className="grid grid-cols-2 lg:grid-cols-5 gap-4">
            <div className="bg-white border border-slate-200 rounded-xl p-4 space-y-1.5">
              <div className="text-xs font-semibold text-slate-500">مؤشر سلامة النظام</div>
              <div
                className={`text-2xl font-bold font-num ${
                  report.healthScore >= 90
                    ? 'text-emerald-700'
                    : report.healthScore >= 65
                      ? 'text-amber-600'
                      : 'text-rose-700'
                }`}
              >
                {report.healthScore}%
              </div>
              <div className="text-[11px] text-slate-500 font-num">
                آخر فحص: {formatDateTime(report.ranAt)}
              </div>
            </div>

            <div className="bg-white border border-slate-200 rounded-xl p-4 space-y-1.5">
              <div className="text-xs font-semibold text-slate-500">مشاكل حرجة</div>
              <div className="text-2xl font-bold font-num text-rose-700">
                {report.summaryCounts.critical}
              </div>
              <div className="text-[11px] text-slate-500">أولوية فورية</div>
            </div>

            <div className="bg-white border border-slate-200 rounded-xl p-4 space-y-1.5">
              <div className="text-xs font-semibold text-slate-500">عالية / متوسطة</div>
              <div className="text-2xl font-bold font-num text-amber-600">
                {report.summaryCounts.high + report.summaryCounts.medium}
              </div>
              <div className="text-[11px] text-slate-500 font-num">
                {report.summaryCounts.high} عالية · {report.summaryCounts.medium} متوسطة
              </div>
            </div>

            <div className="bg-white border border-slate-200 rounded-xl p-4 space-y-1.5">
              <div className="text-xs font-semibold text-slate-500">قابلة للإصلاح التلقائي</div>
              <div className="text-2xl font-bold font-num text-teal-700">
                {report.summaryCounts.autoFixable}
              </div>
              <div className="text-[11px] text-slate-500">مع نسخة احتياطية مسبقة</div>
            </div>

            <div className="bg-white border border-slate-200 rounded-xl p-4 space-y-1.5">
              <div className="text-xs font-semibold text-slate-500">أخطاء تشغيل مسجلة</div>
              <div className="text-2xl font-bold font-num text-slate-900">
                {report.summaryCounts.unresolvedRuntimeErrors}
              </div>
              <div className="text-[11px] text-slate-500">window.onerror / React</div>
            </div>
          </div>

          {/* Executive Summary & Core Checks */}
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-5">
            <div className="lg:col-span-1 bg-white border border-slate-200 rounded-xl p-5 space-y-3">
              <div className="flex items-center justify-between gap-2">
                <h3 className="text-sm font-bold text-slate-900">ملخص المراجع الذكي</h3>
                <span className="text-[11px] font-semibold px-2 py-0.5 rounded bg-slate-100 text-slate-700">
                  {report.aiConfigured ? 'محرك محلي + AI' : 'محرك الفحص المحلي'}
                </span>
              </div>
              <p className="text-xs text-slate-700 leading-relaxed">{report.aiSummary}</p>
              {report.activeBatchFilename && (
                <div className="pt-2 border-t border-slate-100 text-[11px] text-slate-500">
                  الدفعة المفحوصة: <span className="font-semibold text-slate-700">{report.activeBatchFilename}</span>
                </div>
              )}
            </div>

            <div className="lg:col-span-2 bg-white border border-slate-200 rounded-xl p-5 space-y-3">
              <h3 className="text-sm font-bold text-slate-900">نتائج الفحوصات الأساسية</h3>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                {report.checksExecuted.map((chk) => (
                  <div
                    key={chk.key}
                    className={`p-3 rounded-lg border flex items-start justify-between gap-3 ${
                      chk.passed ? 'bg-emerald-50/50 border-emerald-200' : 'bg-amber-50/60 border-amber-200'
                    }`}
                  >
                    <div className="flex items-start gap-2.5 min-w-0">
                      {chk.passed ? (
                        <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                      ) : (
                        <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
                      )}
                      <span className="text-xs font-semibold text-slate-800 leading-snug">{chk.label}</span>
                    </div>
                    <span
                      className={`text-[11px] font-bold font-num px-2 py-0.5 rounded shrink-0 ${
                        chk.passed ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-900'
                      }`}
                    >
                      {chk.passed ? 'سليم' : `${chk.issueCount} ملاحظة`}
                    </span>
                  </div>
                ))}
              </div>
            </div>
          </div>

          {/* Findings List */}
          <div className="bg-white border border-slate-200 rounded-xl p-5 space-y-4">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <h3 className="text-sm font-bold text-slate-900">
                المشاكل والملاحظات المكتشفة ({filteredFindings.length})
              </h3>
              <div className="flex flex-wrap items-center gap-2">
                <select
                  aria-label="تصفية حسب الخطورة"
                  value={severityFilter}
                  onChange={(e) => setSeverityFilter(e.target.value as 'ALL' | AuditSeverity)}
                  className="px-2.5 py-1.5 text-xs border border-slate-300 rounded-lg bg-white text-slate-700"
                >
                  <option value="ALL">جميع المستويات</option>
                  <option value="critical">حرج فقط</option>
                  <option value="high">عالي</option>
                  <option value="medium">متوسط</option>
                  <option value="low">منخفض</option>
                </select>
                <select
                  aria-label="تصفية حسب نوع الإجراء"
                  value={fixTypeFilter}
                  onChange={(e) =>
                    setFixTypeFilter(e.target.value as 'ALL' | 'auto_fixable' | 'code_fix_required')
                  }
                  className="px-2.5 py-1.5 text-xs border border-slate-300 rounded-lg bg-white text-slate-700"
                >
                  <option value="ALL">كل أنواع الإصلاح</option>
                  <option value="auto_fixable">إصلاح تلقائي متاح</option>
                  <option value="code_fix_required">يتطلب تعديل كود (برومبت جاهز)</option>
                </select>
              </div>
            </div>

            {filteredFindings.length === 0 ? (
              <div className="p-6 text-center bg-emerald-50/50 border border-emerald-200 rounded-lg text-xs font-semibold text-emerald-800">
                لا توجد مشاكل مطابقة للتصفية الحالية — جميع الفحوصات في هذا النطاق سليمة.
              </div>
            ) : (
              <div className="space-y-3">
                {filteredFindings.map((finding) => {
                  const meta = SEVERITY_META[finding.severity];
                  const isFixing = busyKey === `fix_${finding.id}`;
                  return (
                    <div
                      key={finding.id}
                      className={`border ${meta.borderClass} rounded-xl p-4 space-y-3 bg-slate-50/40`}
                    >
                      <div className="flex flex-wrap items-start justify-between gap-3">
                        <div className="space-y-1 min-w-0">
                          <div className="flex flex-wrap items-center gap-2">
                            <span
                              className={`text-[11px] font-bold px-2 py-0.5 rounded border ${meta.badgeClass}`}
                            >
                              {meta.label}
                            </span>
                            <span className="text-[11px] font-semibold text-slate-500">
                              {finding.affectedArea}
                            </span>
                            <span className="text-[11px] font-mono text-slate-400" dir="ltr">
                              {finding.code}
                            </span>
                          </div>
                          <h4 className="text-sm font-bold text-slate-900">{finding.title}</h4>
                          <p className="text-xs text-slate-600 leading-relaxed">{finding.description}</p>
                        </div>

                        {finding.fixType === 'auto_fixable' && finding.autoFixAction && (
                          <button
                            type="button"
                            onClick={() => handleAutoFix(finding)}
                            disabled={Boolean(busyKey)}
                            className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-white bg-teal-700 hover:bg-teal-800 rounded-lg transition-colors cursor-pointer shrink-0 disabled:opacity-50"
                          >
                            <Wrench className="w-3.5 h-3.5" />
                            <span>
                              {isFixing ? 'جاري الإصلاح...' : finding.autoFixLabel || 'إصلاح تلقائي'}
                            </span>
                          </button>
                        )}
                      </div>

                      {finding.details.length > 0 && (
                        <ul className="space-y-1 bg-white border border-slate-200 rounded-lg p-3 text-[11px] text-slate-700 font-num">
                          {finding.details.map((d, idx) => (
                            <li key={idx} className="break-words">
                              • {d}
                            </li>
                          ))}
                        </ul>
                      )}

                      {finding.aiInsight && (
                        <div className="p-3 bg-teal-50/70 border border-teal-200 rounded-lg text-xs text-teal-900 space-y-1">
                          <div className="font-bold flex items-center gap-1.5">
                            <Sparkles className="w-3.5 h-3.5 text-teal-700" />
                            <span>تشخيص المراجع الذكي:</span>
                          </div>
                          <p className="leading-relaxed">{finding.aiInsight}</p>
                        </div>
                      )}

                      {finding.readyPrompt && (
                        <div className="bg-slate-900 text-slate-100 rounded-lg p-3.5 space-y-2">
                          <div className="flex items-center justify-between gap-2">
                            <span className="text-[11px] font-bold text-teal-300 flex items-center gap-1.5">
                              <Info className="w-3.5 h-3.5" />
                              برومبت إصلاح جاهز للنسخ
                            </span>
                            <button
                              type="button"
                              onClick={() => handleCopyPrompt(`finding_${finding.id}`, finding.readyPrompt!)}
                              className="inline-flex items-center gap-1.5 px-2.5 py-1 text-[11px] font-semibold bg-slate-800 hover:bg-slate-700 text-white border border-slate-700 rounded cursor-pointer"
                            >
                              {copiedId === `finding_${finding.id}` ? (
                                <>
                                  <Check className="w-3 h-3 text-emerald-400" />
                                  <span>تم النسخ</span>
                                </>
                              ) : (
                                <>
                                  <ClipboardCopy className="w-3 h-3" />
                                  <span>نسخ البرومبت</span>
                                </>
                              )}
                            </button>
                          </div>
                          <pre className="text-[11px] leading-relaxed whitespace-pre-wrap font-sans text-slate-200">
                            {finding.readyPrompt}
                          </pre>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </div>

          {/* Runtime Error Monitor Log */}
          <div className="bg-white border border-slate-200 rounded-xl p-5 space-y-3">
            <div className="flex items-center justify-between gap-3">
              <div className="flex items-center gap-2">
                <AlertCircle className="w-4 h-4 text-rose-600" />
                <h3 className="text-sm font-bold text-slate-900">
                  سجل أخطاء التشغيل الملتقطة ({report.runtimeErrors.length})
                </h3>
              </div>
              <span className="text-[11px] text-slate-500">
                يلتقط تلقائياً window.onerror و unhandledrejection وأخطاء React
              </span>
            </div>

            {report.runtimeErrors.length === 0 ? (
              <div className="p-4 text-center text-xs text-slate-500 bg-slate-50 border border-slate-200 rounded-lg">
                لا توجد أخطاء تشغيل غير معالجة مسجلة.
              </div>
            ) : (
              <div className="overflow-x-auto border border-slate-200 rounded-lg">
                <table className="w-full text-xs text-right">
                  <thead className="bg-slate-50 text-slate-600 border-b border-slate-200">
                    <tr>
                      <th className="py-2.5 px-3 font-semibold">التبويب</th>
                      <th className="py-2.5 px-3 font-semibold">النوع</th>
                      <th className="py-2.5 px-3 font-semibold">الوقت</th>
                      <th className="py-2.5 px-3 font-semibold">المستخدم</th>
                      <th className="py-2.5 px-3 font-semibold">تفاصيل الخطأ</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {report.runtimeErrors.map((err) => (
                      <tr key={err.id} className="hover:bg-slate-50/70">
                        <td className="py-2.5 px-3 font-semibold text-slate-800 whitespace-nowrap">
                          {err.tabName}
                        </td>
                        <td className="py-2.5 px-3 font-mono text-[11px] text-slate-600 whitespace-nowrap" dir="ltr">
                          {err.kind}
                        </td>
                        <td className="py-2.5 px-3 font-num text-slate-600 whitespace-nowrap">
                          {formatDateTime(err.occurredAt)}
                        </td>
                        <td className="py-2.5 px-3 font-num text-slate-600 whitespace-nowrap" dir="ltr">
                          {err.username ? `@${err.username}` : '—'}
                        </td>
                        <td className="py-2.5 px-3 text-rose-700 font-mono text-[11px] break-all" dir="ltr">
                          {err.message}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>

          {/* Pre-Fix Backup Snapshots & Rollback */}
          <div className="bg-white border border-slate-200 rounded-xl p-5 space-y-3">
            <div className="flex items-center justify-between gap-3">
              <div className="flex items-center gap-2">
                <History className="w-4 h-4 text-teal-700" />
                <h3 className="text-sm font-bold text-slate-900">
                  سجل النسخ الاحتياطية للإصلاح التلقائي ({report.backups.length})
                </h3>
              </div>
              <span className="text-[11px] text-slate-500">
                يتم حفظ نسخة احتياطية كاملة قبل أي إصلاح تلقائي مع إمكانية التراجع الفوري
              </span>
            </div>

            {report.backups.length === 0 ? (
              <div className="p-4 text-center text-xs text-slate-500 bg-slate-50 border border-slate-200 rounded-lg">
                لم يتم تنفيذ إصلاحات تلقائية بعد.
              </div>
            ) : (
              <div className="space-y-2">
                {report.backups.map((b) => {
                  const isRestoring = busyKey === `rollback_${b.id}`;
                  return (
                    <div
                      key={b.id}
                      className="flex flex-wrap items-center justify-between gap-3 p-3.5 border border-slate-200 rounded-lg bg-slate-50/50"
                    >
                      <div className="space-y-1 min-w-0">
                        <div className="flex flex-wrap items-center gap-2">
                          <span className="text-xs font-bold text-slate-900">
                            #{b.id} — {b.actionLabel}
                          </span>
                          <span className="text-[11px] text-slate-500 flex items-center gap-1 font-num">
                            <Clock className="w-3 h-3" />
                            {formatDateTime(b.createdAt)}
                          </span>
                        </div>
                        <p className="text-xs text-slate-600">{b.summary}</p>
                      </div>

                      {b.restoredAt ? (
                        <span className="text-[11px] font-semibold px-2.5 py-1 rounded bg-slate-200 text-slate-700 shrink-0">
                          تمت الاستعادة ({formatDateTime(b.restoredAt)})
                        </span>
                      ) : (
                        <button
                          type="button"
                          onClick={() => handleRollback(b.id)}
                          disabled={Boolean(busyKey)}
                          className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-amber-800 bg-amber-50 hover:bg-amber-100 border border-amber-200 rounded-lg transition-colors cursor-pointer shrink-0 disabled:opacity-50"
                        >
                          <RotateCcw className="w-3.5 h-3.5" />
                          <span>{isRestoring ? 'جاري الاستعادة...' : 'تراجع / استعادة'}</span>
                        </button>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </>
      )}
    </div>
  );
};
