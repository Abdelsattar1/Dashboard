import React, { useEffect, useState } from 'react';
import {
  AlertTriangle,
  ArchiveRestore,
  CheckCircle2,
  Download,
  FileSpreadsheet,
  FolderArchive,
  History,
  KeyRound,
  Plus,
  Trash2,
  Upload,
  Users,
  XCircle,
} from 'lucide-react';
import type {
  ImportBatchHistoryItem,
  ImportValidationReport,
  UserProfile,
  UserRole,
} from '../types/bi';
import { apiRequest, downloadProtectedFile } from '../utils/api';
import { formatDateTime, formatEGP, formatNum } from '../utils/format';

const ALL_SECTORS = ['التوريد المباشر', 'التوريد المركزي', 'التوب ريتيل', 'الجيش'] as const;

interface ImportAdminViewProps {
  currentUser: UserProfile;
  onDataUpdated: () => void;
}

export const ImportAdminView: React.FC<ImportAdminViewProps> = ({ currentUser, onDataUpdated }) => {
  const [activeTab, setActiveTab] = useState<'import' | 'history' | 'users' | 'export'>('import');

  // Import state
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [validating, setValidating] = useState(false);
  const [committing, setCommitting] = useState(false);
  const [report, setReport] = useState<ImportValidationReport | null>(null);
  const [importMessage, setImportMessage] = useState<{ type: 'success' | 'error'; text: string } | null>(null);

  // History state
  const [batches, setBatches] = useState<ImportBatchHistoryItem[]>([]);
  const [loadingBatches, setLoadingBatches] = useState(false);

  // Users state
  const [users, setUsers] = useState<UserProfile[]>([]);
  const [loadingUsers, setLoadingUsers] = useState(false);
  const [userError, setUserError] = useState<string | null>(null);
  const [newUsername, setNewUsername] = useState('');
  const [newFullName, setNewFullName] = useState('');
  const [newEmail, setNewEmail] = useState('');
  const [newRole, setNewRole] = useState<UserRole>('sector_user');
  const [newSectors, setNewSectors] = useState<string[]>(['التوريد المباشر']);
  const [newPassword, setNewPassword] = useState('');
  const [editingUserId, setEditingUserId] = useState<number | null>(null);
  const [editPassword, setEditPassword] = useState('');
  const [editSectors, setEditSectors] = useState<string[]>([]);
  const [editRole, setEditRole] = useState<UserRole>('sector_user');

  const loadHistory = async () => {
    setLoadingBatches(true);
    try {
      const res = await apiRequest<{ batches?: ImportBatchHistoryItem[] }>('/api/import/history');
      setBatches(Array.isArray(res?.batches) ? res.batches : []);
    } catch {
      setBatches([]);
    } finally {
      setLoadingBatches(false);
    }
  };

  const loadUsers = async () => {
    setLoadingUsers(true);
    try {
      const res = await apiRequest<{ users?: UserProfile[] }>('/api/users');
      setUsers(Array.isArray(res?.users) ? res.users : []);
    } catch {
      setUsers([]);
    } finally {
      setLoadingUsers(false);
    }
  };

  useEffect(() => {
    if (activeTab === 'history') loadHistory();
    if (activeTab === 'users') loadUsers();
  }, [activeTab]);

  const handleValidateFile = async (file: File) => {
    setSelectedFile(file);
    setReport(null);
    setImportMessage(null);
    setValidating(true);

    try {
      const fd = new FormData();
      fd.append('file', file);
      const res = await apiRequest<ImportValidationReport>('/api/import/validate', {
        method: 'POST',
        body: fd,
      });
      if (!res || !Array.isArray(res.sheets)) {
        throw new Error('استجابة الفحص من الخادم غير مكتملة');
      }
      setReport({
        ...res,
        errors: Array.isArray(res.errors) ? res.errors : [],
        warnings: Array.isArray(res.warnings) ? res.warnings : [],
        sheets: res.sheets.map((s) => ({
          ...s,
          detectedColumns: Array.isArray(s.detectedColumns) ? s.detectedColumns : [],
          missingColumns: Array.isArray(s.missingColumns) ? s.missingColumns : [],
        })),
        usersFound: Array.isArray(res.usersFound) ? res.usersFound : [],
      });
    } catch (err) {
      setImportMessage({
        type: 'error',
        text: err instanceof Error ? err.message : 'فشل فحص ملف Excel',
      });
    } finally {
      setValidating(false);
    }
  };

  const handleCommitImport = async () => {
    if (!report || !report.canCommit) return;
    setCommitting(true);
    setImportMessage(null);

    try {
      await apiRequest('/api/import/commit', {
        method: 'POST',
        body: JSON.stringify({ validationToken: report.validationToken }),
      });
      setImportMessage({
        type: 'success',
        text: `تم اعتماد وحفظ بيانات "${report.filename}" بنجاح مع الاحتفاظ بالنسخ السابقة في سجل الإصدارات.`,
      });
      setReport(null);
      setSelectedFile(null);
      onDataUpdated();
      loadHistory();
    } catch (err) {
      setImportMessage({
        type: 'error',
        text: err instanceof Error ? err.message : 'تعذر اعتماد البيانات',
      });
    } finally {
      setCommitting(false);
    }
  };

  const handleRollback = async (batchId: number) => {
    setImportMessage(null);
    try {
      await apiRequest(`/api/import/rollback/${batchId}`, { method: 'POST' });
      setImportMessage({ type: 'success', text: 'تم استرجاع النسخة المحددة وتفعيلها بنجاح.' });
      await loadHistory();
      onDataUpdated();
    } catch (err) {
      setImportMessage({
        type: 'error',
        text: err instanceof Error ? err.message : 'فشل استرجاع النسخة',
      });
    }
  };

  const handleCreateUser = async (e: React.FormEvent) => {
    e.preventDefault();
    setUserError(null);
    try {
      await apiRequest('/api/users', {
        method: 'POST',
        body: JSON.stringify({
          username: newUsername,
          fullName: newFullName,
          email: newEmail || undefined,
          role: newRole,
          allowedSectors: newRole === 'admin' ? [...ALL_SECTORS] : newSectors,
          password: newPassword,
        }),
      });
      setNewUsername('');
      setNewFullName('');
      setNewEmail('');
      setNewPassword('');
      setNewRole('sector_user');
      setNewSectors(['التوريد المباشر']);
      await loadUsers();
    } catch (err) {
      setUserError(err instanceof Error ? err.message : 'تعذر إنشاء المستخدم');
    }
  };

  const handleSaveUserEdit = async (userId: number) => {
    setUserError(null);
    try {
      await apiRequest(`/api/users/${userId}`, {
        method: 'PATCH',
        body: JSON.stringify({
          role: editRole,
          allowedSectors: editRole === 'admin' ? [...ALL_SECTORS] : editSectors,
          newPassword: editPassword || undefined,
        }),
      });
      setEditingUserId(null);
      setEditPassword('');
      await loadUsers();
    } catch (err) {
      setUserError(err instanceof Error ? err.message : 'تعذر تحديث المستخدم');
    }
  };

  const handleToggleUserActive = async (u: UserProfile) => {
    setUserError(null);
    try {
      await apiRequest(`/api/users/${u.id}`, {
        method: 'PATCH',
        body: JSON.stringify({ isActive: !u.isActive }),
      });
      await loadUsers();
    } catch (err) {
      setUserError(err instanceof Error ? err.message : 'تعذر تغيير حالة الحساب');
    }
  };

  const handleDeleteUser = async (u: UserProfile) => {
    setUserError(null);
    try {
      await apiRequest(`/api/users/${u.id}`, { method: 'DELETE' });
      await loadUsers();
    } catch (err) {
      setUserError(err instanceof Error ? err.message : 'تعذر حذف المستخدم');
    }
  };

  const toggleSectorInList = (list: string[], setter: (v: string[]) => void, sector: string) => {
    if (list.includes(sector)) setter(list.filter((s) => s !== sector));
    else setter([...list, sector]);
  };

  return (
    <div className="space-y-6">
      {/* Navigation Sub-tabs */}
      <div className="flex flex-wrap items-center gap-1 p-1 bg-slate-100 border border-slate-200 rounded-xl w-fit">
        <button
          onClick={() => setActiveTab('import')}
          className={`inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold rounded-lg transition-colors whitespace-nowrap cursor-pointer ${
            activeTab === 'import' ? 'bg-white text-slate-900 shadow-sm' : 'text-slate-600 hover:text-slate-900'
          }`}
        >
          <Upload className="w-3.5 h-3.5" />
          <span>استيراد البيانات (Excel)</span>
        </button>
        <button
          onClick={() => setActiveTab('history')}
          className={`inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold rounded-lg transition-colors whitespace-nowrap cursor-pointer ${
            activeTab === 'history' ? 'bg-white text-slate-900 shadow-sm' : 'text-slate-600 hover:text-slate-900'
          }`}
        >
          <History className="w-3.5 h-3.5" />
          <span>سجل النسخ والاسترجاع</span>
        </button>
        <button
          onClick={() => setActiveTab('users')}
          className={`inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold rounded-lg transition-colors whitespace-nowrap cursor-pointer ${
            activeTab === 'users' ? 'bg-white text-slate-900 shadow-sm' : 'text-slate-600 hover:text-slate-900'
          }`}
        >
          <Users className="w-3.5 h-3.5" />
          <span>المستخدمون والصلاحيات</span>
        </button>
        <button
          onClick={() => setActiveTab('export')}
          className={`inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold rounded-lg transition-colors whitespace-nowrap cursor-pointer ${
            activeTab === 'export' ? 'bg-white text-slate-900 shadow-sm' : 'text-slate-600 hover:text-slate-900'
          }`}
        >
          <FolderArchive className="w-3.5 h-3.5" />
          <span>تصدير المشروع ZIP</span>
        </button>
      </div>

      {importMessage && (
        <div
          className={`p-4 rounded-xl border text-sm font-medium flex items-center gap-2.5 ${
            importMessage.type === 'success'
              ? 'bg-emerald-50 border-emerald-200 text-emerald-900'
              : 'bg-rose-50 border-rose-200 text-rose-900'
          }`}
        >
          {importMessage.type === 'success' ? (
            <CheckCircle2 className="w-5 h-5 text-emerald-700 shrink-0" />
          ) : (
            <XCircle className="w-5 h-5 text-rose-700 shrink-0" />
          )}
          <span>{importMessage.text}</span>
        </div>
      )}

      {/* TAB 1: EXCEL IMPORT & VALIDATION */}
      {activeTab === 'import' && (
        <div className="space-y-6">
          <div className="bg-white border border-slate-200 rounded-xl p-6">
            <div className="flex flex-wrap items-start justify-between gap-4 mb-5">
              <div>
                <h3 className="text-base font-bold text-slate-900">استيراد بيانات المبيعات من ملف Excel (.xlsx)</h3>
                <p className="text-xs text-slate-500 mt-1">
                  يقرأ النظام الأوراق التسع عبر SheetJS ويفحص الأعمدة وعلاقات الربط وتطابق المصادر قبل استبدال البيانات الحالية.
                </p>
              </div>
              <button
                onClick={() => downloadProtectedFile('/api/import/template', 'FMCG_Sales_Template.xlsx')}
                className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold border border-slate-300 rounded-lg text-slate-700 hover:bg-slate-50 cursor-pointer"
              >
                <Download className="w-3.5 h-3.5" />
                <span>تنزيل قالب Excel القياسي</span>
              </button>
            </div>

            <label
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => {
                e.preventDefault();
                const file = e.dataTransfer.files?.[0];
                if (file) handleValidateFile(file);
              }}
              className="flex flex-col items-center justify-center border-2 border-dashed border-slate-300 hover:border-teal-600 rounded-xl p-8 bg-slate-50/70 transition-colors cursor-pointer"
            >
              <FileSpreadsheet className="w-9 h-9 text-teal-700 mb-3" />
              <span className="text-sm font-semibold text-slate-900 mb-1">
                {validating
                  ? 'جاري قراءة وفحص أوراق العمل التسع...'
                  : selectedFile
                    ? `الملف المحدد: ${selectedFile.name}`
                    : 'اختر ملف Excel (.xlsx) من جهازك أو اسحبه وأفلته هنا'}
              </span>
              <span className="text-xs text-slate-500">
                يحافظ النظام على أكواد الفروع والسلاسل كما هي مع الأصفار البادئة (Leading Zeros)
              </span>
              <input
                type="file"
                accept=".xlsx,.xls"
                className="hidden"
                onChange={(e) => {
                  const file = e.target.files?.[0];
                  e.target.value = '';
                  if (file) handleValidateFile(file);
                }}
              />
            </label>
          </div>

          {/* Validation Report Preview */}
          {report && (
            <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
              <div className="p-5 border-b border-slate-200 flex flex-wrap items-center justify-between gap-4">
                <div>
                  <div className="flex items-center gap-2">
                    {report.canCommit ? (
                      <CheckCircle2 className="w-5 h-5 text-emerald-700" />
                    ) : (
                      <XCircle className="w-5 h-5 text-rose-700" />
                    )}
                    <h4 className="text-sm font-bold text-slate-900">
                      تقرير فحص الملف: {report.filename}
                    </h4>
                  </div>
                  <p className="text-xs text-slate-500 mt-1">
                    {report.canCommit
                      ? 'اجتاز الملف الفحص الأساسي بنجاح. راجع الإجماليات والتحذيرات أدناه ثم اضغط اعتماد.'
                      : 'يحتوي الملف على أخطاء تمنع الاعتماد. لن يتم المساس بالبيانات الحالية.'}
                  </p>
                </div>

                <div className="flex items-center gap-3">
                  <button
                    onClick={() => {
                      setReport(null);
                      setSelectedFile(null);
                    }}
                    className="px-4 py-2 text-xs font-semibold text-slate-700 border border-slate-300 rounded-lg hover:bg-slate-50 cursor-pointer"
                  >
                    إلغاء
                  </button>
                  <button
                    onClick={handleCommitImport}
                    disabled={!report.canCommit || committing}
                    className="px-5 py-2 text-xs font-semibold text-white bg-teal-700 hover:bg-teal-800 disabled:opacity-40 rounded-lg transition-colors cursor-pointer"
                  >
                    {committing ? 'جاري الحفظ...' : 'اعتماد وحفظ البيانات'}
                  </button>
                </div>
              </div>

              {/* Errors & Warnings */}
              {((report.errors?.length ?? 0) > 0 || (report.warnings?.length ?? 0) > 0) && (
                <div className="p-5 border-b border-slate-200 space-y-3 bg-slate-50/50">
                  {(report.errors ?? []).map((err, i) => (
                    <div key={`err-${i}`} className="p-3 bg-rose-50 border border-rose-200 rounded-lg text-xs text-rose-900">
                      <div className="font-bold mb-0.5">
                        خطأ [{err.sheet}]: {err.message}
                      </div>
                      {Array.isArray(err.details) && err.details.length > 0 && (
                        <ul className="list-disc list-inside text-[11px] mt-1 space-y-0.5">
                          {err.details.map((d, idx) => (
                            <li key={idx}>{d}</li>
                          ))}
                        </ul>
                      )}
                    </div>
                  ))}
                  {(report.warnings ?? []).map((warn, i) => (
                    <div key={`warn-${i}`} className="p-3 bg-amber-50 border border-amber-200 rounded-lg text-xs text-amber-950">
                      <div className="font-semibold flex items-center gap-1.5">
                        <AlertTriangle className="w-3.5 h-3.5 text-amber-700 shrink-0" />
                        <span>
                          تحذير [{warn.sheet}]: {warn.message}
                        </span>
                      </div>
                      {Array.isArray(warn.details) && warn.details.length > 0 && (
                        <ul className="list-disc list-inside text-[11px] mt-1.5 space-y-0.5 text-amber-900">
                          {warn.details.map((d, idx) => (
                            <li key={idx}>{d}</li>
                          ))}
                        </ul>
                      )}
                    </div>
                  ))}
                </div>
              )}

              {/* Metrics Comparison Preview */}
              <div className="p-5 grid grid-cols-1 md:grid-cols-4 gap-4 border-b border-slate-200 text-xs">
                <div>
                  <div className="text-slate-500 mb-1">مبيعات VF05N (2)</div>
                  <div className="text-sm font-bold font-num text-slate-900">{formatEGP(report.comparison?.vf05n?.salesTotal)}</div>
                  <div className="text-slate-500 mt-0.5 font-num">
                    الضريبة: {formatEGP(report.comparison?.vf05n?.taxTotal)} · {formatNum(report.comparison?.vf05n?.count)} سجل
                  </div>
                </div>
                <div>
                  <div className="text-slate-500 mb-1">مبيعات Zsales_Price (2)</div>
                  <div className="text-sm font-bold font-num text-slate-900">{formatEGP(report.comparison?.zsales?.salesTotal)}</div>
                  <div className="text-slate-500 mt-0.5 font-num">
                    الضريبة: {formatEGP(report.comparison?.zsales?.taxTotal)} · {formatNum(report.comparison?.zsales?.count)} سجل
                  </div>
                </div>
                <div>
                  <div className="text-slate-500 mb-1">الربط بالقطاعات والفروع</div>
                  <div className="text-sm font-bold font-num text-slate-900">
                    الفروع: {formatNum(report.linkStats?.vf05nBranchLinkedRows)} / {formatNum(report.linkStats?.vf05nTotalRows)}
                  </div>
                  <div className="text-slate-500 mt-0.5 font-num">
                    القطاعات: {formatNum(report.linkStats?.vf05nSectorLinkedRows)} / {formatNum(report.linkStats?.vf05nTotalRows)}
                  </div>
                </div>
                <div>
                  <div className="text-slate-500 mb-1">أهداف القطاعات</div>
                  <div className="text-sm font-bold font-num text-slate-900">{formatEGP(report.targetsSummary?.totalSectorTargets)}</div>
                  <div className="text-slate-500 mt-0.5 font-num">
                    فروع التوريد المباشر: {formatEGP(report.targetsSummary?.directBranchesTargetTotal)}
                  </div>
                </div>
              </div>

              {/* 9 Sheets Table */}
              <div className="overflow-x-auto">
                <table className="w-full text-xs text-right">
                  <thead className="bg-slate-50 border-b border-slate-200 font-semibold text-slate-600">
                    <tr>
                      <th className="py-2.5 px-4">ورقة العمل المطلوبة</th>
                      <th className="py-2.5 px-4">الحالة</th>
                      <th className="py-2.5 px-4">عدد الصفوف</th>
                      <th className="py-2.5 px-4">الأعمدة المكتشفة / المفقودة</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {(report.sheets ?? []).map((s) => {
                      const missingCols = Array.isArray(s.missingColumns) ? s.missingColumns : [];
                      const detectedCols = Array.isArray(s.detectedColumns) ? s.detectedColumns : [];
                      return (
                        <tr key={s.sheetName}>
                          <td className="py-2.5 px-4 font-semibold text-slate-900">{s.sheetName}</td>
                          <td className="py-2.5 px-4">
                            {s.found && missingCols.length === 0 ? (
                              <span className="text-emerald-700 font-semibold">سليمة</span>
                            ) : (
                              <span className="text-rose-700 font-semibold">ناقصة</span>
                            )}
                          </td>
                          <td className="py-2.5 px-4 font-num text-slate-700">{formatNum(s.rowCount)}</td>
                          <td className="py-2.5 px-4 text-slate-600">
                            {missingCols.length > 0 ? (
                              <span className="text-rose-700 font-medium">مفقود: {missingCols.join('، ')}</span>
                            ) : (
                              <span>{detectedCols.slice(0, 6).join(' · ')}{detectedCols.length > 6 ? ' ...' : ''}</span>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
      )}

      {/* TAB 2: SNAPSHOT HISTORY & ROLLBACK */}
      {activeTab === 'history' && (
        <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
          <div className="p-5 border-b border-slate-200">
            <h3 className="text-sm font-bold text-slate-900">سجل عمليات الاستيراد والنسخ المحفوظة</h3>
            <p className="text-xs text-slate-500 mt-0.5">
              يتم الاحتفاظ بجميع النسخ السابقة في قاعدة البيانات مع إمكانية التراجع الفوري لأي نسخة سابقة.
            </p>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-xs text-right">
              <thead className="bg-slate-50 border-b border-slate-200 font-semibold text-slate-600">
                <tr>
                  <th className="py-3 px-4">#</th>
                  <th className="py-3 px-4">اسم الملف</th>
                  <th className="py-3 px-4">تاريخ الاستيراد</th>
                  <th className="py-3 px-4">بواسطة</th>
                  <th className="py-3 px-4">إجمالي VF05N</th>
                  <th className="py-3 px-4">الضريبة</th>
                  <th className="py-3 px-4">السجلات / غير المرتبط</th>
                  <th className="py-3 px-4">الحالة</th>
                  <th className="py-3 px-4">إجراء</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {loadingBatches ? (
                  <tr>
                    <td colSpan={9} className="py-8 text-center text-slate-500">جاري التحميل...</td>
                  </tr>
                ) : batches.length === 0 ? (
                  <tr>
                    <td colSpan={9} className="py-8 text-center text-slate-500">لا توجد عمليات استيراد مسجلة بعد</td>
                  </tr>
                ) : (
                  batches.map((b) => (
                    <tr key={b.id} className="hover:bg-slate-50/70">
                      <td className="py-3 px-4 font-num text-slate-500">{b.id}</td>
                      <td className="py-3 px-4 font-semibold text-slate-900">{b.filename}</td>
                      <td className="py-3 px-4 font-num text-slate-600">{formatDateTime(b.createdAt)}</td>
                      <td className="py-3 px-4 text-slate-700">{b.importedByName}</td>
                      <td className="py-3 px-4 font-num font-semibold text-slate-900">{formatEGP(b.vf05nSalesTotal)}</td>
                      <td className="py-3 px-4 font-num text-slate-600">{formatEGP(b.vf05nTaxTotal)}</td>
                      <td className="py-3 px-4 font-num text-slate-600">
                        {formatNum(b.vf05nCount)} سجل · {formatNum(b.unlinkedBranchCount)} غير مرتبط
                      </td>
                      <td className="py-3 px-4">
                        {b.status === 'active' ? (
                          <span className="text-emerald-700 font-bold">النسخة النشطة</span>
                        ) : (
                          <span className="text-slate-500">مؤرشفة</span>
                        )}
                      </td>
                      <td className="py-3 px-4">
                        {b.status !== 'active' && (
                          <button
                            onClick={() => handleRollback(b.id)}
                            className="inline-flex items-center gap-1 px-3 py-1.5 text-xs font-semibold text-teal-800 border border-teal-300 bg-teal-50 hover:bg-teal-100 rounded-lg cursor-pointer"
                          >
                            <ArchiveRestore className="w-3.5 h-3.5" />
                            <span>استرجاع هذه النسخة</span>
                          </button>
                        )}
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* TAB 3: USER MANAGEMENT & SECTOR RBAC */}
      {activeTab === 'users' && (
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          {/* Create User Form */}
          <div className="bg-white border border-slate-200 rounded-xl p-5">
            <h3 className="text-sm font-bold text-slate-900 mb-1">إضافة مستخدم جديد</h3>
            <p className="text-xs text-slate-500 mb-4">تُحفظ كلمات المرور مشفرة عبر scrypt على الخادم</p>

            {userError && (
              <div className="mb-3 p-2.5 bg-rose-50 border border-rose-200 rounded-lg text-xs text-rose-800">
                {userError}
              </div>
            )}

            <form onSubmit={handleCreateUser} className="space-y-3">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">اسم المستخدم</label>
                <input
                  type="text"
                  required
                  dir="ltr"
                  value={newUsername}
                  onChange={(e) => setNewUsername(e.target.value)}
                  placeholder="user.direct"
                  className="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg text-left"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">الاسم الكامل</label>
                <input
                  type="text"
                  required
                  value={newFullName}
                  onChange={(e) => setNewFullName(e.target.value)}
                  placeholder="الاسم الكامل"
                  className="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">البريد الإلكتروني</label>
                <input
                  type="email"
                  dir="ltr"
                  value={newEmail}
                  onChange={(e) => setNewEmail(e.target.value)}
                  placeholder="user@company.com"
                  className="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg text-left"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">الدور الوظيفي</label>
                <select
                  value={newRole}
                  onChange={(e) => setNewRole(e.target.value as UserRole)}
                  className="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg bg-white"
                >
                  <option value="sector_user">مستخدم بصلاحيات قطاعية</option>
                  <option value="admin">مدير نظام (جميع الصلاحيات)</option>
                </select>
              </div>

              {newRole === 'sector_user' && (
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1.5">القطاعات المصرح بها</label>
                  <div className="space-y-1.5">
                    {ALL_SECTORS.map((sec) => (
                      <label key={sec} className="flex items-center gap-2 text-xs text-slate-700 cursor-pointer">
                        <input
                          type="checkbox"
                          checked={newSectors.includes(sec)}
                          onChange={() => toggleSectorInList(newSectors, setNewSectors, sec)}
                          className="rounded border-slate-300 text-teal-700 focus:ring-teal-600"
                        />
                        <span>{sec}</span>
                      </label>
                    ))}
                  </div>
                </div>
              )}

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">كلمة المرور الأولية (8 أحرف+)</label>
                <input
                  type="password"
                  required
                  minLength={8}
                  dir="ltr"
                  value={newPassword}
                  onChange={(e) => setNewPassword(e.target.value)}
                  placeholder="••••••••"
                  className="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg text-left"
                />
              </div>

              <button
                type="submit"
                className="w-full inline-flex items-center justify-center gap-1.5 px-4 py-2 bg-teal-700 hover:bg-teal-800 text-white text-xs font-semibold rounded-lg cursor-pointer"
              >
                <Plus className="w-4 h-4" />
                <span>إنشاء المستخدم</span>
              </button>
            </form>
          </div>

          {/* Users Table */}
          <div className="lg:col-span-2 bg-white border border-slate-200 rounded-xl overflow-hidden">
            <div className="p-5 border-b border-slate-200">
              <h3 className="text-sm font-bold text-slate-900">المستخدمون المسجلون وصلاحيات الوصول</h3>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-xs text-right">
                <thead className="bg-slate-50 border-b border-slate-200 font-semibold text-slate-600">
                  <tr>
                    <th className="py-3 px-4">المستخدم</th>
                    <th className="py-3 px-4">الدور</th>
                    <th className="py-3 px-4">القطاعات المسموحة</th>
                    <th className="py-3 px-4">الحالة</th>
                    <th className="py-3 px-4">إجراءات</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {loadingUsers ? (
                    <tr>
                      <td colSpan={5} className="py-8 text-center text-slate-500">جاري التحميل...</td>
                    </tr>
                  ) : (
                    users.map((u) => (
                      <React.Fragment key={u.id}>
                        <tr className="hover:bg-slate-50/70">
                          <td className="py-3 px-4">
                            <div className="font-semibold text-slate-900">{u.fullName}</div>
                            <div className="font-num text-slate-500">{u.username}</div>
                          </td>
                          <td className="py-3 px-4 text-slate-700 font-medium">
                            {u.role === 'admin' ? 'مدير نظام' : 'مستخدم قطاعي'}
                          </td>
                          <td className="py-3 px-4 text-slate-600">
                            {u.role === 'admin'
                              ? 'جميع القطاعات'
                              : (Array.isArray(u.allowedSectors) ? u.allowedSectors.join(' · ') : '') || 'لا يوجد'}
                          </td>
                          <td className="py-3 px-4">
                            <button
                              onClick={() => handleToggleUserActive(u)}
                              disabled={u.id === currentUser.id}
                              className={`font-semibold cursor-pointer disabled:opacity-50 ${
                                u.isActive ? 'text-emerald-700' : 'text-rose-700'
                              }`}
                            >
                              {u.isActive ? 'مفعل' : 'موقوف'}
                            </button>
                          </td>
                          <td className="py-3 px-4">
                            <div className="flex items-center gap-2">
                              <button
                                onClick={() => {
                                  if (editingUserId === u.id) {
                                    setEditingUserId(null);
                                  } else {
                                    setEditingUserId(u.id);
                                    setEditRole(u.role);
                                    setEditSectors(u.allowedSectors);
                                    setEditPassword('');
                                  }
                                }}
                                className="inline-flex items-center gap-1 px-2.5 py-1 border border-slate-300 rounded text-slate-700 hover:bg-slate-50 cursor-pointer"
                              >
                                <KeyRound className="w-3 h-3" />
                                <span>تعديل / كلمة مرور</span>
                              </button>
                              {u.id !== currentUser.id && (
                                <button
                                  onClick={() => handleDeleteUser(u)}
                                  className="p-1.5 text-rose-700 hover:bg-rose-50 rounded cursor-pointer"
                                  title="حذف المستخدم"
                                >
                                  <Trash2 className="w-3.5 h-3.5" />
                                </button>
                              )}
                            </div>
                          </td>
                        </tr>
                        {editingUserId === u.id && (
                          <tr className="bg-slate-50">
                            <td colSpan={5} className="p-4">
                              <div className="flex flex-wrap items-end gap-4">
                                <div>
                                  <label className="block text-[11px] font-semibold text-slate-700 mb-1">الدور</label>
                                  <select
                                    value={editRole}
                                    onChange={(e) => setEditRole(e.target.value as UserRole)}
                                    disabled={u.id === currentUser.id}
                                    className="px-3 py-1.5 text-xs border border-slate-300 rounded bg-white"
                                  >
                                    <option value="sector_user">مستخدم قطاعي</option>
                                    <option value="admin">مدير نظام</option>
                                  </select>
                                </div>
                                {editRole === 'sector_user' && (
                                  <div className="flex flex-wrap items-center gap-3">
                                    {ALL_SECTORS.map((sec) => (
                                      <label key={sec} className="flex items-center gap-1.5 text-xs text-slate-700">
                                        <input
                                          type="checkbox"
                                          checked={editSectors.includes(sec)}
                                          onChange={() => toggleSectorInList(editSectors, setEditSectors, sec)}
                                        />
                                        <span>{sec}</span>
                                      </label>
                                    ))}
                                  </div>
                                )}
                                <div>
                                  <label className="block text-[11px] font-semibold text-slate-700 mb-1">كلمة مرور جديدة (اختياري)</label>
                                  <input
                                    type="password"
                                    dir="ltr"
                                    value={editPassword}
                                    onChange={(e) => setEditPassword(e.target.value)}
                                    placeholder="اتركه فارغًا دون تغيير"
                                    className="px-3 py-1.5 text-xs border border-slate-300 rounded bg-white text-left"
                                  />
                                </div>
                                <button
                                  onClick={() => handleSaveUserEdit(u.id)}
                                  className="px-4 py-1.5 bg-teal-700 text-white text-xs font-semibold rounded hover:bg-teal-800 cursor-pointer"
                                >
                                  حفظ التعديلات
                                </button>
                              </div>
                            </td>
                          </tr>
                        )}
                      </React.Fragment>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* TAB 4: PROJECT ZIP EXPORT FOR WINDOWS 10 */}
      {activeTab === 'export' && (
        <div className="bg-white border border-slate-200 rounded-xl p-6 max-w-3xl space-y-5">
          <div>
            <h3 className="text-base font-bold text-slate-900">تصدير المشروع كاملًا (ZIP) للتشغيل على Windows 10</h3>
            <p className="text-xs text-slate-600 mt-1">
              يجمع هذا الخيار جميع ملفات الكود المصدري للواجهة والخادم وقاعدة البيانات SQLite والاختبارات في ملف مضغوط واحد جاهز للتشغيل المحلي.
            </p>
          </div>

          <button
            onClick={() => downloadProtectedFile('/api/admin/export-project-zip', 'sales-performance-dashboard.zip')}
            className="inline-flex items-center gap-2 px-5 py-2.5 bg-teal-700 hover:bg-teal-800 text-white text-sm font-semibold rounded-lg transition-colors cursor-pointer"
          >
            <Download className="w-4 h-4" />
            <span>تحميل حزمة المشروع (ZIP)</span>
          </button>

          <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 text-xs text-slate-700 space-y-2">
            <div className="font-bold text-slate-900">خطوات التشغيل على Windows 10 بعد فك الضغط:</div>
            <ol className="list-decimal list-inside space-y-1.5">
              <li>تثبيت Node.js (الإصدار 20 أو 22 LTS).</li>
              <li>فتح موجه الأوامر (CMD أو PowerShell) داخل مجلد المشروع.</li>
              <li>
                تثبيت المكتبات: <code className="font-num bg-white px-1.5 py-0.5 border border-slate-200 rounded">npm install</code>
              </li>
              <li>
                بناء الواجهة: <code className="font-num bg-white px-1.5 py-0.5 border border-slate-200 rounded">npm run build</code>
              </li>
              <li>
                تشغيل النظام: <code className="font-num bg-white px-1.5 py-0.5 border border-slate-200 rounded">npm start</code> ثم فتح{' '}
                <code className="font-num bg-white px-1.5 py-0.5 border border-slate-200 rounded">http://localhost:3000</code>
              </li>
              <li>
                تشغيل اختبارات التحقق الحسابي: <code className="font-num bg-white px-1.5 py-0.5 border border-slate-200 rounded">npm test</code>
              </li>
            </ol>
          </div>
        </div>
      )}
    </div>
  );
};
