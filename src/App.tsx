import React, { useCallback, useEffect, useState } from 'react';
import {
  BarChart3,
  Building2,
  CalendarRange,
  Database,
  FileSpreadsheet,
  GitBranch,
  KeyRound,
  LayoutDashboard,
  Lock,
  LogOut,
  Menu,
  Package,
  ReceiptText,
  Shield,
  Trophy,
  UserCog,
  Users,
  X,
} from 'lucide-react';
import { BranchesView } from './components/BranchesView';
import { CompetitionsView } from './components/CompetitionsView';
import { ExecutiveDashboard } from './components/ExecutiveDashboard';
import { ImportAdminView, type ImportAdminSubTab } from './components/ImportAdminView';
import { ItemsView } from './components/ItemsView';
import { LoginView } from './components/LoginView';
import { RepsView } from './components/RepsView';
import { SectorsView } from './components/SectorsView';
import { TimelineView } from './components/TimelineView';
import { TransactionsView } from './components/TransactionsView';
import type {
  AggregationMetricRow,
  AuthStatusResponse,
  BranchMetricRow,
  CompetitionAnalyticsResponse,
  DashboardSummaryResponse,
  DepartmentChainRow,
  ItemCategoriesAnalyticsResponse,
  RepMetricRow,
  SalesSource,
  SectorMetricRow,
  TimelinePoint,
  UnlinkedRecordRow,
  UserProfile,
} from './types/bi';
import { apiRequest, setAuthToken, setOnUnauthorized } from './utils/api';
import { formatDateTime } from './utils/format';
import { canEditTargetsRole, getRoleShortLabel } from './utils/roles';

type NavSection =
  | 'dashboard'
  | 'sectors'
  | 'branches'
  | 'reps'
  | 'items'
  | 'competitions'
  | 'timeline'
  | 'transactions'
  | 'import-admin';

const ALL_SECTORS = ['التوريد المباشر', 'التوريد المركزي', 'التوب ريتيل', 'الجيش'];

function resolveUserSectors(u: UserProfile | null): string[] {
  if (!u) return [];
  if (u.role === 'admin') return ALL_SECTORS;
  const raw = Array.isArray(u.allowedSectors) ? u.allowedSectors : [];
  if (raw.some((s) => s.trim() === 'الكل' || s.trim().toLowerCase() === 'all')) {
    return ALL_SECTORS;
  }
  return Array.from(new Set(raw.filter((s) => ALL_SECTORS.includes(s))));
}

const NAV_ITEMS: Array<{ id: NavSection; label: string; icon: React.FC<{ className?: string }>; adminOnly?: boolean }> = [
  { id: 'dashboard', label: 'لوحة المعلومات الرئيسية', icon: LayoutDashboard },
  { id: 'sectors', label: 'القطاعات والإدارات', icon: GitBranch },
  { id: 'branches', label: 'الفروع', icon: Building2 },
  { id: 'reps', label: 'جهاز البيع', icon: Users },
  { id: 'items', label: 'تحليل فئات الأصناف', icon: Package },
  { id: 'competitions', label: 'مسابقات', icon: Trophy },
  { id: 'timeline', label: 'التحليل الزمني', icon: CalendarRange },
  { id: 'transactions', label: 'تفاصيل المبيعات', icon: ReceiptText },
  { id: 'import-admin', label: 'استيراد البيانات والنظام', icon: Database, adminOnly: true },
];

export default function App() {
  const [authChecked, setAuthChecked] = useState(false);
  const [needsSetup, setNeedsSetup] = useState(false);
  const [user, setUser] = useState<UserProfile | null>(null);

  const [section, setSection] = useState<NavSection>('dashboard');
  const [importAdminTab, setImportAdminTab] = useState<ImportAdminSubTab>('import');
  const [branchesInitialTab, setBranchesInitialTab] = useState<'linked' | 'unlinked'>('linked');
  const [sidebarOpen, setSidebarOpen] = useState(false);

  // Self-service password change state
  const [pwModalOpen, setPwModalOpen] = useState(false);
  const [currentPw, setCurrentPw] = useState('');
  const [newPw, setNewPw] = useState('');
  const [confirmPw, setConfirmPw] = useState('');
  const [pwError, setPwError] = useState<string | null>(null);
  const [pwSuccess, setPwSuccess] = useState<string | null>(null);
  const [pwSaving, setPwSaving] = useState(false);

  // Global interactive filters
  const [source, setSource] = useState<SalesSource>('VF05N');
  const [sectorFilter, setSectorFilter] = useState<string>('ALL');
  const [jobTitleFilter, setJobTitleFilter] = useState<string>('ALL');
  const [startDate, setStartDate] = useState<string>('');
  const [endDate, setEndDate] = useState<string>('');
  const [refreshTick, setRefreshTick] = useState(0);
  const [filterResetSignal, setFilterResetSignal] = useState(0);

  // Data states
  const [dashData, setDashData] = useState<DashboardSummaryResponse | null>(null);
  const [sectorsData, setSectorsData] = useState<{
    sectors: SectorMetricRow[];
    aggregations: AggregationMetricRow[];
    chains: DepartmentChainRow[];
  } | null>(null);
  const [branchesData, setBranchesData] = useState<{
    branches: BranchMetricRow[];
    unlinkedRecords: UnlinkedRecordRow[];
    totalUnlinkedSales: number;
    totalUnlinkedCount: number;
  } | null>(null);
  const [repsData, setRepsData] = useState<RepMetricRow[]>([]);
  const [itemCategoriesData, setItemCategoriesData] = useState<ItemCategoriesAnalyticsResponse | null>(null);
  const [competitionsData, setCompetitionsData] = useState<CompetitionAnalyticsResponse | null>(null);
  const [selectedCompetitionId, setSelectedCompetitionId] = useState<number | null>(null);
  const [timelineData, setTimelineData] = useState<TimelinePoint[]>([]);
  const [loading, setLoading] = useState(false);

  const checkAuthStatus = useCallback(async () => {
    try {
      const status = await apiRequest<AuthStatusResponse>('/api/auth/status');
      setNeedsSetup(status.needsSetup);
    } catch {
      setNeedsSetup(false);
    } finally {
      setAuthChecked(true);
    }
  }, []);

  useEffect(() => {
    setOnUnauthorized(() => {
      setUser(null);
      checkAuthStatus();
    });
    checkAuthStatus();
  }, [checkAuthStatus]);

  const buildQueryString = useCallback(() => {
    const qs = new URLSearchParams({ source, sector: sectorFilter });
    if (section === 'items' && jobTitleFilter && jobTitleFilter !== 'ALL') {
      qs.set('jobTitle', jobTitleFilter);
    }
    if (section === 'competitions' && selectedCompetitionId) {
      qs.set('competitionId', String(selectedCompetitionId));
    }
    if (startDate) qs.set('startDate', startDate);
    if (endDate) qs.set('endDate', endDate);
    return qs.toString();
  }, [source, sectorFilter, section, jobTitleFilter, selectedCompetitionId, startDate, endDate]);

  useEffect(() => {
    if (!user) return;
    let active = true;
    const qs = buildQueryString();

    const loadSectionData = async () => {
      setLoading(true);
      try {
        if (section === 'dashboard') {
          const res = await apiRequest<DashboardSummaryResponse>(`/api/analytics/dashboard?${qs}`);
          if (active) setDashData(res);
        } else if (section === 'sectors') {
          const res = await apiRequest<{
            sectors: SectorMetricRow[];
            aggregations: AggregationMetricRow[];
            chains: DepartmentChainRow[];
          }>(`/api/analytics/sectors?${qs}`);
          if (active) setSectorsData(res);
        } else if (section === 'branches') {
          const res = await apiRequest<{
            branches: BranchMetricRow[];
            unlinkedRecords: UnlinkedRecordRow[];
            totalUnlinkedSales: number;
            totalUnlinkedCount: number;
          }>(`/api/analytics/branches?${qs}`);
          if (active) setBranchesData(res);
        } else if (section === 'reps') {
          const res = await apiRequest<{ reps: RepMetricRow[] }>(`/api/analytics/reps?${qs}`);
          if (active) setRepsData(res.reps);
        } else if (section === 'items') {
          const res = await apiRequest<ItemCategoriesAnalyticsResponse>(`/api/analytics/items?${qs}`);
          if (active) {
            setItemCategoriesData(res);
            if (jobTitleFilter !== 'ALL' && !res.jobTitleOptions.includes(jobTitleFilter)) {
              setJobTitleFilter('ALL');
            }
          }
        } else if (section === 'competitions') {
          const res = await apiRequest<CompetitionAnalyticsResponse>(`/api/competitions/analytics?${qs}`);
          if (active) {
            setCompetitionsData(res);
            if (res.competition && res.competition.id !== selectedCompetitionId) {
              setSelectedCompetitionId(res.competition.id);
            } else if (!res.competition && selectedCompetitionId !== null) {
              setSelectedCompetitionId(null);
            }
          }
        } else if (section === 'timeline') {
          const res = await apiRequest<{ series: TimelinePoint[] }>(`/api/analytics/timeline?${qs}`);
          if (active) setTimelineData(res.series);
        }
      } catch {
        // handled cleanly via empty states
      } finally {
        if (active) setLoading(false);
      }
    };

    loadSectionData();
    return () => {
      active = false;
    };
  }, [user, section, buildQueryString, refreshTick]);

  const handleLogout = () => {
    setAuthToken(null);
    setUser(null);
    setDashData(null);
    setSectorsData(null);
    setBranchesData(null);
    setRepsData([]);
    setItemCategoriesData(null);
    setCompetitionsData(null);
    setTimelineData([]);
    setSectorFilter('ALL');
    setJobTitleFilter('ALL');
    setStartDate('');
    setEndDate('');
    setSource('VF05N');
    setSection('dashboard');
    setPwModalOpen(false);
    checkAuthStatus();
  };

  const handleChangeOwnPassword = async (e: React.FormEvent) => {
    e.preventDefault();
    setPwError(null);
    setPwSuccess(null);
    if (newPw.length < 8) {
      setPwError('كلمة المرور الجديدة يجب أن تكون 8 أحرف على الأقل');
      return;
    }
    if (newPw !== confirmPw) {
      setPwError('كلمة المرور الجديدة وتأكيدها غير متطابقين');
      return;
    }
    setPwSaving(true);
    try {
      await apiRequest('/api/auth/change-password', {
        method: 'POST',
        body: JSON.stringify({ currentPassword: currentPw, newPassword: newPw }),
      });
      setPwSuccess('تم تحديث كلمة المرور بنجاح.');
      setCurrentPw('');
      setNewPw('');
      setConfirmPw('');
    } catch (err) {
      setPwError(err instanceof Error ? err.message : 'تعذر تغيير كلمة المرور');
    } finally {
      setPwSaving(false);
    }
  };

  if (!authChecked) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center text-sm text-slate-600">
        جاري التحقق من الجلسة...
      </div>
    );
  }

  if (!user) {
    return (
      <LoginView
        needsSetup={needsSetup}
        onAuthenticated={(loggedInUser) => {
          const userSectors = resolveUserSectors(loggedInUser);
          setUser(loggedInUser);
          setNeedsSetup(false);
          setSectorFilter(loggedInUser.role !== 'admin' && userSectors.length === 1 ? userSectors[0] : 'ALL');
          setJobTitleFilter('ALL');
          setStartDate('');
          setEndDate('');
          setSource('VF05N');
          setSection('dashboard');
          setImportAdminTab('import');
          setRefreshTick((t) => t + 1);
        }}
      />
    );
  }

  const isAdmin = user.role === 'admin';
  const canEditTargets = canEditTargetsRole(user.role);
  const authorizedSectors = resolveUserSectors(user);
  const isSingleSectorScoped = !isAdmin && authorizedSectors.length === 1;
  const effectiveSectorFilter = isSingleSectorScoped ? authorizedSectors[0] : sectorFilter;
  const visibleNavItems = NAV_ITEMS.filter((item) => !item.adminOnly || isAdmin);
  const currentNavLabel = visibleNavItems.find((i) => i.id === section)?.label ?? 'لوحة المعلومات الرئيسية';

  return (
    <div className="min-h-screen bg-slate-50 flex">
      {/* Mobile Sidebar Backdrop */}
      {sidebarOpen && (
        <div
          onClick={() => setSidebarOpen(false)}
          className="fixed inset-0 bg-slate-900/40 z-30 lg:hidden"
        />
      )}

      {/* Right Sidebar Navigation (RTL) */}
      <aside
        className={`fixed lg:static inset-y-0 right-0 z-40 w-64 bg-white border-l border-slate-200 flex flex-col transition-transform duration-200 lg:translate-x-0 ${
          sidebarOpen ? 'translate-x-0' : 'translate-x-full lg:translate-x-0'
        }`}
      >
        <div className="h-16 px-5 border-b border-slate-200 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-teal-700 text-white flex items-center justify-center shrink-0">
              <BarChart3 className="w-4 h-4" />
            </div>
            <span className="text-sm font-bold text-slate-900 whitespace-nowrap">أداء المبيعات FMCG</span>
          </div>
          <button
            onClick={() => setSidebarOpen(false)}
            className="lg:hidden p-1.5 text-slate-500 hover:text-slate-900"
            aria-label="إغلاق القائمة"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        <nav className="flex-1 p-3 space-y-1 overflow-y-auto">
          {visibleNavItems.map((item) => {
            const Icon = item.icon;
            const active = section === item.id;
            return (
              <button
                key={item.id}
                onClick={() => {
                  if (item.id === 'branches') setBranchesInitialTab('linked');
                  setSection(item.id);
                  setSidebarOpen(false);
                }}
                className={`w-full flex items-center gap-3 px-3.5 py-2.5 text-xs font-semibold rounded-lg transition-colors whitespace-nowrap cursor-pointer ${
                  active
                    ? 'bg-teal-50 text-teal-800 border border-teal-200'
                    : 'text-slate-600 hover:bg-slate-100 hover:text-slate-900 border border-transparent'
                }`}
              >
                <Icon className={`w-4 h-4 shrink-0 ${active ? 'text-teal-700' : 'text-slate-400'}`} />
                <span>{item.label}</span>
              </button>
            );
          })}
        </nav>

        <div className="p-4 border-t border-slate-200 text-xs space-y-2.5 bg-slate-50/70">
          <div className="space-y-1">
            <div className="flex items-center justify-between gap-2">
              <span className="font-bold text-slate-900 truncate">{user.fullName}</span>
              <span className="text-[11px] font-num text-slate-500 shrink-0" dir="ltr">
                @{user.username}
              </span>
            </div>
            <div className="inline-flex items-center gap-1.5 text-[11px] font-semibold text-teal-800">
              <Shield className="w-3 h-3 text-teal-700 shrink-0" />
              <span>{getRoleShortLabel(user.role)}</span>
            </div>
            <div className="text-[11px] text-slate-600 leading-relaxed">
              {isAdmin ? 'النطاق: جميع القطاعات' : `القطاعات: ${authorizedSectors.join(' · ') || 'غير محدد'}`}
            </div>
          </div>

          {isAdmin && (
            <button
              onClick={() => {
                setImportAdminTab('users');
                setSection('import-admin');
                setSidebarOpen(false);
              }}
              className="w-full flex items-center justify-center gap-2 px-3 py-1.5 border border-teal-200 bg-teal-50/60 hover:bg-teal-100/70 rounded-lg text-teal-900 font-semibold transition-colors cursor-pointer"
            >
              <UserCog className="w-3.5 h-3.5" />
              <span>إدارة المستخدمين والصلاحيات</span>
            </button>
          )}

          <div className="grid grid-cols-2 gap-1.5">
            <button
              onClick={() => {
                setPwError(null);
                setPwSuccess(null);
                setPwModalOpen(true);
              }}
              className="flex items-center justify-center gap-1 px-2 py-2 border border-slate-300 bg-white rounded-lg text-slate-700 hover:bg-slate-100 font-semibold transition-colors cursor-pointer whitespace-nowrap"
              title="تغيير كلمة المرور"
            >
              <KeyRound className="w-3.5 h-3.5 shrink-0" />
              <span>كلمة المرور</span>
            </button>
            <button
              onClick={handleLogout}
              className="flex items-center justify-center gap-1 px-2 py-2 border border-slate-300 bg-white rounded-lg text-slate-700 hover:bg-slate-100 font-semibold transition-colors cursor-pointer whitespace-nowrap"
              title="تبديل الحساب أو تسجيل الخروج"
            >
              <LogOut className="w-3.5 h-3.5 shrink-0" />
              <span>تبديل الحساب</span>
            </button>
          </div>
        </div>
      </aside>

      {/* Main Content Area */}
      <div className="flex-1 min-w-0 flex flex-col">
        {/* Top Bar Contract: Breadcrumb/Title — Clean Workspace Info — Primary Action */}
        <header className="h-16 bg-white border-b border-slate-200 px-6 flex items-center justify-between gap-8 shrink-0">
          <div className="flex items-center gap-3 min-w-0">
            <button
              onClick={() => setSidebarOpen(true)}
              className="lg:hidden p-2 text-slate-600 hover:text-slate-900 border border-slate-200 rounded-lg"
              aria-label="فتح القائمة"
            >
              <Menu className="w-4 h-4" />
            </button>
            <h1 className="text-base font-bold text-slate-900 truncate whitespace-nowrap">{currentNavLabel}</h1>
          </div>

          <div className="hidden md:flex items-center gap-2 text-xs text-slate-500 whitespace-nowrap">
            {dashData?.activeBatch ? (
              <>
                <FileSpreadsheet className="w-3.5 h-3.5 text-teal-700 shrink-0" />
                <span>النسخة النشطة: {dashData.activeBatch.filename}</span>
                <span aria-hidden="true">·</span>
                <span className="font-num">{formatDateTime(dashData.activeBatch.createdAt)}</span>
              </>
            ) : (
              <span>لم يتم استيراد ملف بيانات بعد</span>
            )}
          </div>

          <div className="flex items-center gap-3 shrink-0">
            {isAdmin && section !== 'import-admin' && (
              <button
                onClick={() => {
                  setImportAdminTab('import');
                  setSection('import-admin');
                }}
                className="px-4 py-2 text-xs font-semibold text-white bg-teal-700 hover:bg-teal-800 rounded-lg transition-colors whitespace-nowrap shrink-0 cursor-pointer"
              >
                استيراد ملف Excel
              </button>
            )}
          </div>
        </header>

        {/* Workspace Filter Toolbar (below header, outside Top Bar Contract) */}
        {section !== 'import-admin' && (
          <div className="bg-white border-b border-slate-200 px-6 py-3 flex flex-wrap items-center justify-between gap-4">
            <div className="flex flex-wrap items-center gap-3">
              {/* Single-source selector (NEVER summed together) */}
              <div className="flex items-center gap-2">
                <span className="text-xs font-semibold text-slate-600">مصدر المبيعات:</span>
                <div className="flex items-center gap-1 p-1 bg-slate-100 rounded-lg">
                  <button
                    onClick={() => setSource('VF05N')}
                    className={`px-3 py-1 text-xs font-semibold rounded-md transition-colors whitespace-nowrap cursor-pointer ${
                      source === 'VF05N' ? 'bg-white text-teal-800 shadow-sm' : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    VF05N (2)
                  </button>
                  <button
                    onClick={() => setSource('ZSALES')}
                    className={`px-3 py-1 text-xs font-semibold rounded-md transition-colors whitespace-nowrap cursor-pointer ${
                      source === 'ZSALES' ? 'bg-white text-teal-800 shadow-sm' : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Zsales_Price (2)
                  </button>
                </div>
              </div>

              {/* Sector Filter (locked cleanly for single-sector non-admin users) */}
              {isSingleSectorScoped ? (
                <div className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-slate-200 rounded-lg bg-slate-50 text-slate-700">
                  <Lock className="w-3 h-3 text-teal-700 shrink-0" />
                  <span>القطاع: {authorizedSectors[0]}</span>
                </div>
              ) : (
                <select
                  aria-label="القطاع"
                  value={effectiveSectorFilter}
                  onChange={(e) => {
                    setSectorFilter(e.target.value);
                    setJobTitleFilter('ALL');
                  }}
                  className="px-3 py-1.5 text-xs border border-slate-300 rounded-lg bg-white text-slate-700 focus:outline-none focus:ring-2 focus:ring-teal-600"
                >
                  <option value="ALL">
                    {isAdmin ? 'جميع القطاعات (الكل)' : `جميع القطاعات المصرح بها (${authorizedSectors.length})`}
                  </option>
                  {authorizedSectors.map((sec) => (
                    <option key={sec} value={sec}>
                      {sec}
                    </option>
                  ))}
                  {isAdmin && <option value="UNLINKED">سجلات غير مرتبطة بقطاع</option>}
                </select>
              )}

              {/* Job Title Filter beside Sector Filter (cascades from selected Sector on Item Categories tab) */}
              {section === 'items' && (
                <select
                  aria-label="المسمى الوظيفي"
                  value={jobTitleFilter}
                  onChange={(e) => setJobTitleFilter(e.target.value)}
                  className="px-3 py-1.5 text-xs border border-slate-300 rounded-lg bg-white text-slate-700 focus:outline-none focus:ring-2 focus:ring-teal-600"
                >
                  <option value="ALL">جميع المسميات الوظيفية</option>
                  {(itemCategoriesData?.jobTitleOptions ?? []).map((jt) => (
                    <option key={jt} value={jt}>
                      {jt}
                    </option>
                  ))}
                </select>
              )}
            </div>

            {/* Date Range Filters */}
            <div className="flex flex-wrap items-center gap-2 text-xs">
              <label className="text-slate-500">من:</label>
              <input
                type="date"
                value={startDate}
                onChange={(e) => setStartDate(e.target.value)}
                className="px-2.5 py-1.5 border border-slate-300 rounded-lg font-num text-slate-700"
              />
              <label className="text-slate-500">إلى:</label>
              <input
                type="date"
                value={endDate}
                onChange={(e) => setEndDate(e.target.value)}
                className="px-2.5 py-1.5 border border-slate-300 rounded-lg font-num text-slate-700"
              />
              {(startDate || endDate || (!isSingleSectorScoped && sectorFilter !== 'ALL') || jobTitleFilter !== 'ALL') && (
                <button
                  onClick={() => {
                    setStartDate('');
                    setEndDate('');
                    setSectorFilter(isSingleSectorScoped ? authorizedSectors[0] : 'ALL');
                    setJobTitleFilter('ALL');
                    setFilterResetSignal((n) => n + 1);
                  }}
                  className="px-2.5 py-1.5 text-xs font-semibold text-slate-600 hover:text-slate-900 underline cursor-pointer"
                >
                  إعادة ضبط
                </button>
              )}
            </div>
          </div>
        )}

        {/* Main Viewport */}
        <main className="flex-1 p-5 lg:p-6 max-w-[1600px] w-full mx-auto">
          {section === 'dashboard' && (
            <ExecutiveDashboard
              data={dashData}
              loading={loading}
              source={source}
              sectorFilter={effectiveSectorFilter}
              startDate={startDate}
              endDate={endDate}
              user={user}
              onNavigateImport={() => {
                setImportAdminTab('import');
                setSection('import-admin');
              }}
              onNavigateBranchesUnlinked={() => {
                setBranchesInitialTab('unlinked');
                setSection('branches');
              }}
            />
          )}

          {section === 'sectors' && (
            <SectorsView
              sectors={sectorsData?.sectors ?? []}
              aggregations={sectorsData?.aggregations ?? []}
              chains={sectorsData?.chains ?? []}
              source={source}
              globalSectorFilter={effectiveSectorFilter}
              startDate={startDate}
              endDate={endDate}
              loading={loading}
            />
          )}

          {section === 'branches' && (
            <BranchesView
              branches={branchesData?.branches ?? []}
              unlinkedRecords={branchesData?.unlinkedRecords ?? []}
              totalUnlinkedSales={branchesData?.totalUnlinkedSales ?? 0}
              totalUnlinkedCount={branchesData?.totalUnlinkedCount ?? 0}
              loading={loading}
              initialTab={branchesInitialTab}
              authorizedSectors={authorizedSectors}
              source={source}
              selectedSector={effectiveSectorFilter}
              startDate={startDate}
              endDate={endDate}
              onSectorChange={setSectorFilter}
            />
          )}

          {section === 'reps' && (
            <RepsView
              reps={repsData}
              source={source}
              sectorFilter={effectiveSectorFilter}
              startDate={startDate}
              endDate={endDate}
              loading={loading}
            />
          )}

          {section === 'items' && (
            <ItemsView
              rows={itemCategoriesData?.rows ?? []}
              effectiveSource={itemCategoriesData?.effectiveSource ?? source}
              requestedSource={source}
              selectedSector={effectiveSectorFilter}
              selectedJobTitle={jobTitleFilter}
              startDate={startDate}
              endDate={endDate}
              filterResetSignal={filterResetSignal}
              isAdmin={isAdmin}
              canEditTargets={canEditTargets}
              authorizedSectors={authorizedSectors}
              loading={loading}
              onRefresh={() => setRefreshTick((t) => t + 1)}
            />
          )}

          {section === 'competitions' && (
            <CompetitionsView
              data={competitionsData}
              loading={loading}
              isAdmin={isAdmin}
              canEditTargets={canEditTargets}
              requestedSource={source}
              selectedSector={effectiveSectorFilter}
              startDate={startDate}
              endDate={endDate}
              selectedCompetitionId={selectedCompetitionId}
              onSelectCompetition={(id) => {
                setSelectedCompetitionId(id);
              }}
              onRefresh={() => setRefreshTick((t) => t + 1)}
            />
          )}

          {section === 'timeline' && (
            <TimelineView
              series={timelineData}
              source={source}
              sectorFilter={effectiveSectorFilter}
              startDate={startDate}
              endDate={endDate}
              loading={loading}
            />
          )}

          {section === 'transactions' && (
            <TransactionsView
              source={source}
              sectorFilter={effectiveSectorFilter}
              startDate={startDate}
              endDate={endDate}
            />
          )}

          {section === 'import-admin' && isAdmin && (
            <ImportAdminView
              currentUser={user}
              initialTab={importAdminTab}
              onDataUpdated={() => {
                setRefreshTick((t) => t + 1);
              }}
            />
          )}
        </main>
      </div>

      {/* Change Password Modal */}
      {pwModalOpen && (
        <div className="fixed inset-0 bg-slate-900/40 z-50 flex items-center justify-center p-4">
          <div className="w-full max-w-sm bg-white border border-slate-200 rounded-xl p-6 space-y-4 shadow-xl">
            <div className="flex items-center justify-between">
              <h3 className="text-sm font-bold text-slate-900">تغيير كلمة المرور</h3>
              <button
                onClick={() => setPwModalOpen(false)}
                className="p-1 text-slate-400 hover:text-slate-700 cursor-pointer"
                aria-label="إغلاق"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            {pwError && (
              <div className="p-2.5 bg-rose-50 border border-rose-200 rounded-lg text-xs text-rose-800">{pwError}</div>
            )}
            {pwSuccess && (
              <div className="p-2.5 bg-emerald-50 border border-emerald-200 rounded-lg text-xs text-emerald-800">
                {pwSuccess}
              </div>
            )}
            <form onSubmit={handleChangeOwnPassword} className="space-y-3">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">كلمة المرور الحالية</label>
                <input
                  type="password"
                  required
                  dir="ltr"
                  value={currentPw}
                  onChange={(e) => setCurrentPw(e.target.value)}
                  className="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg text-left"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">كلمة المرور الجديدة (8 أحرف+)</label>
                <input
                  type="password"
                  required
                  minLength={8}
                  dir="ltr"
                  value={newPw}
                  onChange={(e) => setNewPw(e.target.value)}
                  className="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg text-left"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">تأكيد كلمة المرور الجديدة</label>
                <input
                  type="password"
                  required
                  minLength={8}
                  dir="ltr"
                  value={confirmPw}
                  onChange={(e) => setConfirmPw(e.target.value)}
                  className="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg text-left"
                />
              </div>
              <div className="flex items-center justify-end gap-2 pt-2">
                <button
                  type="button"
                  onClick={() => setPwModalOpen(false)}
                  className="px-3 py-1.5 text-xs font-semibold text-slate-600 hover:text-slate-900 cursor-pointer"
                >
                  إغلاق
                </button>
                <button
                  type="submit"
                  disabled={pwSaving}
                  className="px-4 py-1.5 text-xs font-semibold bg-teal-700 hover:bg-teal-800 text-white rounded-lg cursor-pointer disabled:opacity-50"
                >
                  {pwSaving ? 'جاري الحفظ...' : 'حفظ كلمة المرور'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}
