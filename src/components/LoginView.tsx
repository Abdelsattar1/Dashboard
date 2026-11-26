import React, { useState } from 'react';
import { BarChart3, Lock, ShieldCheck, User } from 'lucide-react';
import type { LoginResponse, UserProfile } from '../types/bi';
import { apiRequest, setAuthToken } from '../utils/api';

interface LoginViewProps {
  needsSetup: boolean;
  onAuthenticated: (user: UserProfile) => void;
}

export const LoginView: React.FC<LoginViewProps> = ({ needsSetup, onAuthenticated }) => {
  const [username, setUsername] = useState(needsSetup ? 'admin' : '');
  const [fullName, setFullName] = useState(needsSetup ? 'مدير النظام' : '');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);

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
      onAuthenticated(res.user);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'تعذر إتمام العملية');
    } finally {
      setLoading(false);
    }
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
          <div className="mb-5">
            <h2 className="text-base font-semibold text-slate-900">تسجيل الدخول</h2>
            <p className="text-xs text-slate-500 mt-1">أدخل بيانات حسابك للوصول إلى لوحة المعلومات المصرح بها</p>
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
