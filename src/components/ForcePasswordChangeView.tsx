import React, { useState } from 'react';
import { CheckCircle2, Eye, EyeOff, KeyRound, Lock, LogOut, ShieldCheck } from 'lucide-react';
import type { ChangePasswordResponse, UserProfile } from '../types/bi';
import { apiRequest } from '../utils/api';
import { getRoleShortLabel } from '../utils/roles';

interface ForcePasswordChangeViewProps {
  user: UserProfile;
  tempPassword: string;
  onCompleted: (updatedUser: UserProfile) => void;
  onLogout: () => void;
}

export const ForcePasswordChangeView: React.FC<ForcePasswordChangeViewProps> = ({
  user,
  tempPassword,
  onCompleted,
  onLogout,
}) => {
  const [manualCurrentPw, setManualCurrentPw] = useState('');
  const [newPassword, setNewPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const [showNew, setShowNew] = useState(false);
  const [showConfirm, setShowConfirm] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const effectiveCurrentPw = tempPassword || manualCurrentPw;

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);

    if (!effectiveCurrentPw) {
      setError('يرجى إدخال كلمة المرور المبدئية الحالية');
      return;
    }
    if (newPassword.length < 8) {
      setError('كلمة المرور الجديدة يجب أن لا تقل عن 8 أحرف');
      return;
    }
    if (newPassword === effectiveCurrentPw) {
      setError('كلمة المرور الجديدة يجب أن تختلف عن كلمة المرور المبدئية');
      return;
    }
    if (newPassword !== confirmPassword) {
      setError('كلمة المرور الجديدة وتأكيدها غير متطابقين');
      return;
    }

    setSaving(true);
    try {
      const res = await apiRequest<ChangePasswordResponse>('/api/auth/change-password', {
        method: 'POST',
        body: JSON.stringify({
          currentPassword: effectiveCurrentPw,
          newPassword,
        }),
      });
      const updatedUser: UserProfile = res.user
        ? { ...res.user, mustChangePassword: false }
        : { ...user, mustChangePassword: false };
      onCompleted(updatedUser);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'تعذر حفظ كلمة المرور الجديدة');
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col justify-center items-center p-6">
      <div className="w-full max-w-md bg-white border border-slate-200 rounded-xl p-8 shadow-sm">
        <div className="flex items-center justify-between gap-3 mb-5 pb-4 border-b border-slate-100">
          <div className="flex items-center gap-3 min-w-0">
            <div className="w-10 h-10 rounded-lg bg-amber-600 text-white flex items-center justify-center shrink-0">
              <KeyRound className="w-5 h-5" />
            </div>
            <div className="min-w-0">
              <h1 className="text-base font-bold text-slate-900 truncate">تعيين كلمة المرور الدائمة</h1>
              <p className="text-xs text-slate-500 truncate">
                {user.fullName} · <span dir="ltr">@{user.username}</span> ({getRoleShortLabel(user.role)})
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onLogout}
            className="inline-flex items-center gap-1 px-2.5 py-1.5 text-xs font-semibold text-slate-600 hover:text-slate-900 border border-slate-200 rounded-lg cursor-pointer shrink-0"
          >
            <LogOut className="w-3.5 h-3.5" />
            <span>خروج</span>
          </button>
        </div>

        <div className="mb-5 p-3.5 bg-amber-50 border border-amber-200 rounded-lg text-xs text-amber-950 leading-relaxed space-y-1">
          <div className="flex items-center gap-1.5 font-bold text-amber-900">
            <ShieldCheck className="w-4 h-4 text-amber-700 shrink-0" />
            <span>إجراء أمني إلزامي عند أول دخول</span>
          </div>
          <p>
            تم تسجيل دخولك باستخدام كلمة مرور مبدئية. لحماية حسابك، يجب تعيين كلمة مرور جديدة خاصة بك وتأكيدها قبل الدخول إلى لوحة المعلومات.
          </p>
        </div>

        {error && (
          <div className="mb-4 p-3 rounded-lg bg-rose-50 border border-rose-200 text-xs text-rose-800 font-medium">
            {error}
          </div>
        )}

        <form onSubmit={handleSubmit} className="space-y-4">
          {!tempPassword && (
            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1.5">كلمة المرور المبدئية الحالية</label>
              <input
                type="password"
                required
                dir="ltr"
                value={manualCurrentPw}
                onChange={(e) => setManualCurrentPw(e.target.value)}
                placeholder="••••••••"
                className="w-full px-3 py-2.5 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-teal-600 text-left"
              />
            </div>
          )}

          <div>
            <label className="block text-xs font-semibold text-slate-700 mb-1.5">كلمة المرور الجديدة (8 أحرف على الأقل)</label>
            <div className="relative">
              <Lock className="w-4 h-4 text-slate-400 absolute right-3 top-1/2 -translate-y-1/2" />
              <input
                type={showNew ? 'text' : 'password'}
                required
                minLength={8}
                dir="ltr"
                value={newPassword}
                onChange={(e) => setNewPassword(e.target.value)}
                placeholder="أدخل كلمة المرور الجديدة"
                className="w-full pr-9 pl-10 py-2.5 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-teal-600 text-left"
              />
              <button
                type="button"
                onClick={() => setShowNew((v) => !v)}
                className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-700 cursor-pointer"
                aria-label="إظهار أو إخفاء كلمة المرور"
              >
                {showNew ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
              </button>
            </div>
          </div>

          <div>
            <label className="block text-xs font-semibold text-slate-700 mb-1.5">تأكيد كلمة المرور الجديدة</label>
            <div className="relative">
              <Lock className="w-4 h-4 text-slate-400 absolute right-3 top-1/2 -translate-y-1/2" />
              <input
                type={showConfirm ? 'text' : 'password'}
                required
                minLength={8}
                dir="ltr"
                value={confirmPassword}
                onChange={(e) => setConfirmPassword(e.target.value)}
                placeholder="أعد كتابة كلمة المرور الجديدة"
                className="w-full pr-9 pl-10 py-2.5 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-teal-600 text-left"
              />
              <button
                type="button"
                onClick={() => setShowConfirm((v) => !v)}
                className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-700 cursor-pointer"
                aria-label="إظهار أو إخفاء التأكيد"
              >
                {showConfirm ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
              </button>
            </div>
          </div>

          <button
            type="submit"
            disabled={saving}
            className="w-full inline-flex items-center justify-center gap-2 py-2.5 px-4 bg-teal-700 hover:bg-teal-800 disabled:opacity-60 text-white text-sm font-semibold rounded-lg transition-colors cursor-pointer"
          >
            <CheckCircle2 className="w-4 h-4" />
            <span>{saving ? 'جاري حفظ كلمة المرور...' : 'حفظ كلمة المرور والدخول للنظام'}</span>
          </button>
        </form>
      </div>
    </div>
  );
};
