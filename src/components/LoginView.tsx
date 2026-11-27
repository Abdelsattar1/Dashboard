import React, { useState } from 'react';
import { BarChart3, Check, Cloud, Copy, KeyRound, Lock, ShieldCheck, User } from 'lucide-react';
import type { LoginResponse, UserProfile } from '../types/bi';
import { apiRequest, setAuthToken } from '../utils/api';
import { CLIENT_DEFAULT_ADMIN_SEED, CLIENT_DEFAULT_ADMIN_USERNAME } from '../utils/client-auth';
import { SUPABASE_SCHEMA_SQL } from '../utils/supabase-sql';

interface LoginViewProps {
  needsSetup: boolean;
  clientMode?: boolean;
  defaultAdminAvailable?: boolean;
  supabaseConfigured?: boolean;
  supabaseSetupError?: string | null;
  onAuthenticated: (user: UserProfile, usedPassword: string) => void;
}

export const LoginView: React.FC<LoginViewProps> = ({
  needsSetup,
  clientMode = false,
  defaultAdminAvailable = false,
  supabaseConfigured = false,
  supabaseSetupError = null,
  onAuthenticated,
}) => {
  const [username, setUsername] = useState(needsSetup ? 'admin' : '');
  const [fullName, setFullName] = useState(needsSetup ? 'مدير النظام' : '');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [copiedSql, setCopiedSql] = useState(false);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);

    if (needsSetup && password !== confirmPassword) {
      setError('كلمة المرور وتأكيدها غير متطابقين');
      return;
    }

    setLoading(true);
    try {
      const endpoint = needsSetup ? '/api/auth/setup-admin' : '/api/auth/login';
      const body = needsSetup
        ? { username, fullName, email: email || undefined, password }
        : { username, password };

      const res = await apiRequest<LoginResponse>(endpoint, {
        method: 'POST',
        body: JSON.stringify(body),
      });

      setAuthToken(res.token);
      onAuthenticated(res.user, password);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'تعذر إتمام العملية');
    } finally {
      setLoading(false);
    }
  };

  const handleFillDefaultAdmin = () => {
    setUsername(CLIENT_DEFAULT_ADMIN_USERNAME);
    setPassword(CLIENT_DEFAULT_ADMIN_SEED);
    setError(null);
  };

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col justify-center items-center p-6">
      <div className="w-full max-w-md bg-white border border-slate-200 rounded-xl p-8">
        <div className="flex items-center gap-3 mb-6 pb-5 border-b border-slate-100">
          <div className="w-10 h-10 rounded-lg bg-teal-700 text-white flex items-center justify-center shrink-0">
            <BarChart3 className="w-5 h-5" />
          </div>
          <div>
            <h1 className="text-lg font-bold text-slate-900">نظام متابعة أداء المبيعات</h1>
            <p className="text-xs text-slate-500">تحليل مبيعات السلع الاستهلاكية FMCG</p>
          </div>
        </div>

        {needsSetup ? (
          <div className="mb-5 p-3.5 bg-teal-50 border border-teal-200 rounded-lg text-xs text-teal-900 leading-relaxed">
            <div className="flex items-center gap-2 font-semibold mb-1">
              <ShieldCheck className="w-4 h-4 text-teal-700 shrink-0" />
              <span>إعداد حساب مدير النظام الأول</span>
            </div>
            <p>
              لا توجد حسابات مسجلة بعد في قاعدة البيانات. أنشئ حساب مدير النظام الآن للبدء في استيراد ملفات Excel وإدارة الصلاحيات.
            </p>
          </div>
        ) : (
          <div className="mb-5 space-y-3">
            <div>
              <h2 className="text-base font-semibold text-slate-900">تسجيل الدخول</h2>
              <p className="text-xs text-slate-500 mt-1 leading-relaxed">
                أدخل بيانات حسابك للوصول إلى لوحة المعلومات والقطاعات المصرح لك بها (أدمن، مدير قطاع/إدارة، أو مستخدم قياسي).
              </p>
            </div>

            {supabaseSetupError && (
              <div className="p-3.5 bg-amber-50 border border-amber-200 rounded-lg text-xs text-amber-950 space-y-2">
                <div className="font-bold flex items-center gap-1.5 text-amber-900">
                  <Cloud className="w-4 h-4 text-amber-700 shrink-0" />
                  <span>تهيئة جداول Supabase مطلوبة</span>
                </div>
                <p className="leading-relaxed">{supabaseSetupError}</p>
                <button
                  type="button"
                  onClick={() => {
                    navigator.clipboard.writeText(SUPABASE_SCHEMA_SQL).then(() => {
                      setCopiedSql(true);
                      setTimeout(() => setCopiedSql(false), 2500);
                    }).catch(() => {});
                  }}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-amber-700 hover:bg-amber-800 text-white rounded text-[11px] font-semibold cursor-pointer"
                >
                  {copiedSql ? <Check className="w-3.5 h-3.5" /> : <Copy className="w-3.5 h-3.5" />}
                  <span>{copiedSql ? 'تم نسخ كود SQL!' : 'نسخ كود إنشاء الجداول (SQL)'}</span>
                </button>
              </div>
            )}

            {clientMode && defaultAdminAvailable && (
              <div className="p-3.5 bg-teal-50/80 border border-teal-200 rounded-lg text-xs text-teal-950 space-y-2">
                <div className="flex items-center justify-between gap-2">
                  <span className="font-bold flex items-center gap-1.5 text-teal-900">
                    <KeyRound className="w-3.5 h-3.5 text-teal-700 shrink-0" />
                    بيانات الدخول الافتراضية لمدير النظام
                  </span>
                  <button
                    type="button"
                    onClick={handleFillDefaultAdmin}
                    className="px-2.5 py-1 bg-teal-700 hover:bg-teal-800 text-white rounded text-[11px] font-semibold transition-colors cursor-pointer shrink-0"
                  >
                    تعبئة تلقائية
                  </button>
                </div>
                <div className="grid grid-cols-2 gap-2 bg-white border border-teal-100 rounded p-2 font-num">
                  <div>
                    <span className="text-slate-500">اسم المستخدم: </span>
                    <span className="font-bold text-slate-900" dir="ltr">
                      {CLIENT_DEFAULT_ADMIN_USERNAME}
                    </span>
                  </div>
                  <div>
                    <span className="text-slate-500">كلمة المرور: </span>
                    <span className="font-bold text-slate-900" dir="ltr">
                      {CLIENT_DEFAULT_ADMIN_SEED}
                    </span>
                  </div>
                </div>
                <p className="text-[11px] text-teal-800 leading-relaxed">
                  {supabaseConfigured
                    ? 'النظام متصل بقاعدة بيانات Supabase السحابية المشتركة. سيُطلب تعيين كلمة مرور جديدة عند أول دخول.'
                    : 'يعمل التطبيق داخل المتصفح مع حفظ المستخدمين والأهداف والبيانات تلقائياً. سيُطلب تعيين كلمة مرور جديدة عند أول دخول.'}
                </p>
              </div>
            )}
          </div>
        )}

        {error && (
          <div className="mb-4 p-3 rounded-lg bg-rose-50 border border-rose-200 text-xs text-rose-800 font-medium">
            {error}
          </div>
        )}

        <form onSubmit={handleSubmit} className="space-y-4">
          <div>
            <label className="block text-xs font-semibold text-slate-700 mb-1.5">اسم المستخدم</label>
            <div className="relative">
              <User className="w-4 h-4 text-slate-400 absolute right-3 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                required
                dir="ltr"
                value={username}
                onChange={(e) => setUsername(e.target.value)}
                placeholder="admin"
                className="w-full pr-9 pl-3 py-2.5 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-teal-600 focus:border-teal-600 text-left"
              />
            </div>
          </div>

          {needsSetup && (
            <>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1.5">الاسم الكامل</label>
                <input
                  type="text"
                  required
                  value={fullName}
                  onChange={(e) => setFullName(e.target.value)}
                  placeholder="الاسم الكامل للمدير"
                  className="w-full px-3 py-2.5 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-teal-600 focus:border-teal-600"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1.5">البريد الإلكتروني (اختياري)</label>
                <input
                  type="email"
                  dir="ltr"
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  placeholder="admin@company.com"
                  className="w-full px-3 py-2.5 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-teal-600 focus:border-teal-600 text-left"
                />
              </div>
            </>
          )}

          <div>
            <label className="block text-xs font-semibold text-slate-700 mb-1.5">كلمة المرور</label>
            <div className="relative">
              <Lock className="w-4 h-4 text-slate-400 absolute right-3 top-1/2 -translate-y-1/2" />
              <input
                type="password"
                required
                dir="ltr"
                minLength={needsSetup ? 8 : 1}
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                placeholder="••••••••"
                className="w-full pr-9 pl-3 py-2.5 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-teal-600 focus:border-teal-600 text-left"
              />
            </div>
            {needsSetup && <p className="text-[11px] text-slate-500 mt-1">8 أحرف على الأقل</p>}
          </div>

          {needsSetup && (
            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1.5">تأكيد كلمة المرور</label>
              <input
                type="password"
                required
                dir="ltr"
                minLength={8}
                value={confirmPassword}
                onChange={(e) => setConfirmPassword(e.target.value)}
                placeholder="••••••••"
                className="w-full px-3 py-2.5 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-teal-600 focus:border-teal-600 text-left"
              />
            </div>
          )}

          <button
            type="submit"
            disabled={loading}
            className="w-full py-2.5 px-4 bg-teal-700 hover:bg-teal-800 disabled:opacity-60 text-white text-sm font-semibold rounded-lg transition-colors cursor-pointer"
          >
            {loading
              ? 'جاري المعالجة...'
              : needsSetup
                ? 'إنشاء حساب المدير والدخول'
                : 'تسجيل الدخول'}
          </button>
        </form>
      </div>
    </div>
  );
};
