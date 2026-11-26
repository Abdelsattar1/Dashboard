import React, { useEffect, useMemo, useState } from 'react';
import {
  ArrowUpDown,
  Check,
  ChevronDown,
  ChevronUp,
  Info,
  Plus,
  Save,
  Search,
  Trash2,
  Trophy,
  X,
} from 'lucide-react';
import type {
  CompetitionAnalyticsResponse,
  CompetitionCatalogItem,
  CompetitionItemConfig,
  CompetitionParticipantRow,
  SalesSource,
} from '../types/bi';
import { apiRequest } from '../utils/api';
import {
  buildActiveFilterMeta,
  buildExportFilename,
  exportToExcel,
  exportToPdf,
  type ExportColumn,
} from '../utils/exporters';
import { formatNum } from '../utils/format';
import { ExportButtons } from './ExportButtons';

interface CompetitionsViewProps {
  data: CompetitionAnalyticsResponse | null;
  loading: boolean;
  isAdmin: boolean;
  canEditTargets: boolean;
  requestedSource: SalesSource;
  selectedSector: string;
  startDate: string;
  endDate: string;
  selectedCompetitionId: number | null;
  onSelectCompetition: (id: number | null) => void;
  onRefresh: () => void;
}

type RoleFilter = 'ALL' | CompetitionParticipantRow['roleKind'];
type SortKey = 'default' | 'targetQty' | 'soldStandardQty' | 'remainingQty' | 'achievementPct';

function formatQty(value: number | null | undefined): string {
  if (value === null || value === undefined || !Number.isFinite(value)) return '—';
  return value.toLocaleString('en-US', {
    minimumFractionDigits: 0,
    maximumFractionDigits: 3,
  });
}

function formatAchievement(pct: number | null | undefined): string {
  if (pct === null || pct === undefined || !Number.isFinite(pct)) return '—';
  return `${pct.toLocaleString('en-US', {
    minimumFractionDigits: 1,
    maximumFractionDigits: 1,
  })}%`;
}

function roundQty(val: number): number {
  if (!Number.isFinite(val)) return 0;
  return Number(val.toFixed(3));
}

function computeSectorRollupFromList(
  rows: Array<{ roleKind: CompetitionParticipantRow['roleKind']; sectorName: string; targetQty: number }>,
  sector: string
): number {
  let secMgr = 0;
  let sup = 0;
  let rep = 0;
  for (const r of rows) {
    if (r.sectorName !== sector) continue;
    if (r.roleKind === 'sector_manager') secMgr += r.targetQty;
    else if (r.roleKind === 'supervisor') sup += r.targetQty;
    else if (r.roleKind === 'sales_rep') rep += r.targetQty;
  }
  if (secMgr > 0) return roundQty(secMgr);
  if (sup > 0) return roundQty(sup);
  return roundQty(rep);
}

export const CompetitionsView: React.FC<CompetitionsViewProps> = ({
  data,
  loading,
  isAdmin,
  canEditTargets,
  requestedSource,
  selectedSector,
  startDate,
  endDate,
  selectedCompetitionId,
  onSelectCompetition,
  onRefresh,
}) => {
  const competition = data?.competition ?? null;
  const competitions = data?.competitions ?? [];
  const catalogItems = data?.catalogItems ?? [];
  const participants = data?.participants ?? [];

  // Competition create/rename/delete UI state
  const [creatingNew, setCreatingNew] = useState(false);
  const [newCompName, setNewCompName] = useState('');
  const [renamingComp, setRenamingComp] = useState(false);
  const [renameDraft, setRenameDraft] = useState('');
  const [confirmDeleteId, setConfirmDeleteId] = useState<number | null>(null);
  const [busyAction, setBusyAction] = useState(false);
  const [feedback, setFeedback] = useState<{ type: 'success' | 'error'; text: string } | null>(null);

  // Item selection + standard strength factor drafts
  const [showItemsPanel, setShowItemsPanel] = useState(false);
  const [itemSearch, setItemSearch] = useState('');
  const [selectedItemDrafts, setSelectedItemDrafts] = useState<
    Array<{ itemCode: string; itemName: string; factorStr: string }>
  >([]);
  const [itemsDirty, setItemsDirty] = useState(false);

  // Target qty drafts per participantKey
  const [targetDrafts, setTargetDrafts] = useState<Record<string, string>>({});
  const [targetsDirty, setTargetsDirty] = useState(false);

  // Table filter & sort
  const [roleFilter, setRoleFilter] = useState<RoleFilter>('ALL');
  const [searchText, setSearchText] = useState('');
  const [sortKey, setSortKey] = useState<SortKey>('default');
  const [sortAsc, setSortAsc] = useState(false);

  // Sync local item drafts whenever competition changes
  useEffect(() => {
    if (!competition) {
      setSelectedItemDrafts([]);
      setItemsDirty(false);
      setRenamingComp(false);
      setConfirmDeleteId(null);
      return;
    }
    setSelectedItemDrafts(
      competition.items.map((it) => ({
        itemCode: it.itemCode,
        itemName: it.itemName,
        factorStr: String(it.equivalentFactor),
      }))
    );
    setItemsDirty(false);
    setRenameDraft(competition.name);
    setRenamingComp(false);
    setConfirmDeleteId(null);
    if (competition.items.length === 0 && isAdmin) {
      setShowItemsPanel(true);
    }
  }, [competition?.id, competition?.updatedAt, isAdmin]);

  // Sync local target drafts when participants load/refresh
  useEffect(() => {
    const next: Record<string, string> = {};
    for (const p of participants) {
      if (!p.isAutoTarget) {
        next[p.participantKey] = String(p.targetQty);
      }
    }
    setTargetDrafts(next);
    setTargetsDirty(false);
  }, [competition?.id, participants]);

  // Catalog lookup map
  const catalogByCode = useMemo(() => {
    const map = new Map<string, CompetitionCatalogItem>();
    for (const item of catalogItems) {
      map.set(item.itemCode, item);
    }
    return map;
  }, [catalogItems]);

  const selectedCodesSet = useMemo(
    () => new Set(selectedItemDrafts.map((i) => i.itemCode)),
    [selectedItemDrafts]
  );

  const filteredCatalog = useMemo(() => {
    const q = itemSearch.trim().toLowerCase();
    if (!q) return catalogItems;
    return catalogItems.filter(
      (it) =>
        it.itemName.toLowerCase().includes(q) ||
        it.itemCode.toLowerCase().includes(q) ||
        it.categoryName.toLowerCase().includes(q) ||
        it.chainsCategory.toLowerCase().includes(q)
    );
  }, [catalogItems, itemSearch]);

  const toggleCatalogItem = (item: CompetitionCatalogItem) => {
    if (!isAdmin) return;
    setFeedback(null);
    setItemsDirty(true);
    setSelectedItemDrafts((prev) => {
      if (prev.some((x) => x.itemCode === item.itemCode)) {
        return prev.filter((x) => x.itemCode !== item.itemCode);
      }
      return [
        ...prev,
        {
          itemCode: item.itemCode,
          itemName: item.itemName,
          factorStr: '1',
        },
      ];
    });
  };

  const selectAllFilteredCatalog = () => {
    if (!isAdmin) return;
    setFeedback(null);
    setItemsDirty(true);
    setSelectedItemDrafts((prev) => {
      const existing = new Set(prev.map((x) => x.itemCode));
      const additions = filteredCatalog
        .filter((it) => !existing.has(it.itemCode))
        .map((it) => ({ itemCode: it.itemCode, itemName: it.itemName, factorStr: '1' }));
      return [...prev, ...additions];
    });
  };

  const clearSelectedItems = () => {
    if (!isAdmin) return;
    setFeedback(null);
    setItemsDirty(true);
    setSelectedItemDrafts([]);
  };

  const updateFactorStr = (itemCode: string, val: string) => {
    setFeedback(null);
    setItemsDirty(true);
    setSelectedItemDrafts((prev) =>
      prev.map((it) => (it.itemCode === itemCode ? { ...it, factorStr: val } : it))
    );
  };

  const handleCreateCompetition = async (e: React.FormEvent) => {
    e.preventDefault();
    const clean = newCompName.trim();
    if (!clean) {
      setFeedback({ type: 'error', text: 'يرجى إدخال اسم المسابقة أولاً' });
      return;
    }
    setBusyAction(true);
    setFeedback(null);
    try {
      const res = await apiRequest<{ competition: { id: number; name: string } }>('/api/competitions', {
        method: 'POST',
        body: JSON.stringify({ name: clean, items: [] }),
      });
      setNewCompName('');
      setCreatingNew(false);
      setShowItemsPanel(true);
      onSelectCompetition(res.competition.id);
      onRefresh();
      setFeedback({ type: 'success', text: `تم إنشاء المسابقة «${res.competition.name}» بنجاح` });
    } catch (err) {
      setFeedback({ type: 'error', text: err instanceof Error ? err.message : 'تعذر إنشاء المسابقة' });
    } finally {
      setBusyAction(false);
    }
  };

  const handleSaveRename = async () => {
    if (!competition) return;
    const clean = renameDraft.trim();
    if (!clean) {
      setFeedback({ type: 'error', text: 'اسم المسابقة لا يمكن أن يكون فارغاً' });
      return;
    }
    setBusyAction(true);
    setFeedback(null);
    try {
      await apiRequest(`/api/competitions/${competition.id}`, {
        method: 'PUT',
        body: JSON.stringify({ name: clean }),
      });
      setRenamingComp(false);
      onRefresh();
      setFeedback({ type: 'success', text: 'تم تحديث اسم المسابقة' });
    } catch (err) {
      setFeedback({ type: 'error', text: err instanceof Error ? err.message : 'تعذر تحديث اسم المسابقة' });
    } finally {
      setBusyAction(false);
    }
  };

  const handleDeleteCompetition = async (compId: number) => {
    setBusyAction(true);
    setFeedback(null);
    try {
      await apiRequest(`/api/competitions/${compId}`, { method: 'DELETE' });
      setConfirmDeleteId(null);
      onSelectCompetition(null);
      onRefresh();
      setFeedback({ type: 'success', text: 'تم حذف المسابقة' });
    } catch (err) {
      setFeedback({ type: 'error', text: err instanceof Error ? err.message : 'تعذر حذف المسابقة' });
    } finally {
      setBusyAction(false);
    }
  };

  const handleSaveItems = async () => {
    if (!competition) return;
    const payloadItems: CompetitionItemConfig[] = [];
    for (const draft of selectedItemDrafts) {
      const parsed = draft.factorStr.trim() === '' ? 1 : Number(draft.factorStr);
      if (!Number.isFinite(parsed) || parsed < 0) {
        setFeedback({
          type: 'error',
          text: `معامل الشدة المعيارية غير صالح للصنف ${draft.itemName} (${draft.itemCode})`,
        });
        return;
      }
      payloadItems.push({
        itemCode: draft.itemCode,
        itemName: draft.itemName,
        equivalentFactor: parsed,
      });
    }

    setBusyAction(true);
    setFeedback(null);
    try {
      await apiRequest(`/api/competitions/${competition.id}`, {
        method: 'PUT',
        body: JSON.stringify({ items: payloadItems }),
      });
      setItemsDirty(false);
      onRefresh();
      setFeedback({ type: 'success', text: 'تم حفظ مجموعة الأصناف ومعاملات الشدة المعيارية' });
    } catch (err) {
      setFeedback({ type: 'error', text: err instanceof Error ? err.message : 'تعذر حفظ الأصناف' });
    } finally {
      setBusyAction(false);
    }
  };

  const handleSaveTargets = async () => {
    if (!competition) return;
    const targetsPayload: Array<{ participantKey: string; targetQty: number }> = [];
    for (const p of participants) {
      if (p.isAutoTarget) continue;
      const raw = (targetDrafts[p.participantKey] ?? '0').trim();
      const parsed = raw === '' ? 0 : Number(raw);
      if (!Number.isFinite(parsed) || parsed < 0) {
        setFeedback({
          type: 'error',
          text: `قيمة الهدف الكمي غير صالحة للمشارك ${p.employeeName}`,
        });
        return;
      }
      targetsPayload.push({
        participantKey: p.participantKey,
        targetQty: parsed,
      });
    }

    setBusyAction(true);
    setFeedback(null);
    try {
      await apiRequest(`/api/competitions/${competition.id}/targets`, {
        method: 'PUT',
        body: JSON.stringify({ targets: targetsPayload }),
      });
      setTargetsDirty(false);
      onRefresh();
      setFeedback({ type: 'success', text: 'تم حفظ الأهداف الكمية وتحديث إجماليات مديري الإدارات' });
    } catch (err) {
      setFeedback({ type: 'error', text: err instanceof Error ? err.message : 'تعذر حفظ الأهداف الكمية' });
    } finally {
      setBusyAction(false);
    }
  };

  // Compute live participants with instant client-side recalculation for targets, remaining qty, % and مدير إدارة rollup
  const liveParticipants = useMemo(() => {
    const subRowsOld = participants.filter((p) => !p.isAutoTarget);
    const subRowsNew = subRowsOld.map((p) => {
      const raw = targetDrafts[p.participantKey];
      const parsed = raw === undefined ? p.targetQty : raw.trim() === '' ? 0 : Number(raw);
      const validTarget = Number.isFinite(parsed) && parsed >= 0 ? roundQty(parsed) : p.targetQty;
      return {
        ...p,
        liveTargetQty: validTarget,
      };
    });

    const newTargetByKey = new Map<string, number>();
    for (const sr of subRowsNew) {
      newTargetByKey.set(sr.participantKey, sr.liveTargetQty);
    }

    // Determine sectors present in current subordinate view so even single-sector filter deltas stay exact
    const sectorsInView = Array.from(new Set(subRowsOld.map((r) => r.sectorName).filter(Boolean)));
    const deltaBySector = new Map<string, number>();
    for (const sec of sectorsInView) {
      const oldRoll = computeSectorRollupFromList(subRowsOld, sec);
      const newRoll = computeSectorRollupFromList(
        subRowsNew.map((r) => ({ roleKind: r.roleKind, sectorName: r.sectorName, targetQty: r.liveTargetQty })),
        sec
      );
      deltaBySector.set(sec, newRoll - oldRoll);
    }

    return participants.map((p) => {
      let targetQty = p.targetQty;
      if (p.isAutoTarget) {
        let delta = 0;
        for (const sec of p.managedSectors) {
          delta += deltaBySector.get(sec) ?? 0;
        }
        targetQty = Math.max(0, roundQty(p.targetQty + delta));
      } else {
        targetQty = newTargetByKey.get(p.participantKey) ?? p.targetQty;
      }
      const remainingQty = targetQty > 0 ? Math.max(0, roundQty(targetQty - p.soldStandardQty)) : 0;
      const pct =
        targetQty > 0 && Number.isFinite(p.soldStandardQty / targetQty)
          ? Number(((p.soldStandardQty / targetQty) * 100).toFixed(2))
          : null;

      return {
        ...p,
        targetQty,
        remainingQty,
        achievementPct: pct,
      };
    });
  }, [participants, targetDrafts]);

  // Live summary across visible sectors (non-overlapping)
  const liveSummary = useMemo(() => {
    const subRows = liveParticipants.filter((p) => !p.isAutoTarget);
    const sectors = Array.from(new Set(subRows.map((p) => p.sectorName).filter(Boolean)));
    let totalTargetQty = 0;
    for (const sec of sectors) {
      totalTargetQty += computeSectorRollupFromList(subRows, sec);
    }
    const roundedTarget = roundQty(totalTargetQty);
    const soldTotal = data?.summary.totalStandardSoldQty ?? 0;
    const pct =
      roundedTarget > 0 && Number.isFinite(soldTotal / roundedTarget)
        ? Number(((soldTotal / roundedTarget) * 100).toFixed(2))
        : null;
    return {
      totalSelectedItems: selectedItemDrafts.length,
      totalStandardSoldQty: soldTotal,
      totalTargetQty: roundedTarget,
      overallAchievementPct: pct,
    };
  }, [liveParticipants, data?.summary.totalStandardSoldQty, selectedItemDrafts.length]);

  const filteredRows = useMemo(() => {
    const q = searchText.trim().toLowerCase();
    const list = liveParticipants.filter((r) => {
      if (roleFilter !== 'ALL' && r.roleKind !== roleFilter) return false;
      if (!q) return true;
      return (
        r.employeeName.toLowerCase().includes(q) ||
        r.roleLabel.toLowerCase().includes(q) ||
        r.sectorName.toLowerCase().includes(q) ||
        r.employeeCode.toLowerCase().includes(q)
      );
    });

    if (sortKey === 'default') return list;

    return [...list].sort((a, b) => {
      if (sortKey === 'achievementPct') {
        const av = a.achievementPct ?? -Infinity;
        const bv = b.achievementPct ?? -Infinity;
        return sortAsc ? av - bv : bv - av;
      }
      const diff = a[sortKey] - b[sortKey];
      return sortAsc ? diff : -diff;
    });
  }, [liveParticipants, roleFilter, searchText, sortKey, sortAsc]);

  const toggleSort = (k: SortKey) => {
    if (sortKey === k) {
      setSortAsc(!sortAsc);
    } else {
      setSortKey(k);
      setSortAsc(false);
    }
  };

  const exportTotals = useMemo(() => {
    const distinctRoles = new Set(filteredRows.map((r) => r.roleKind));
    if (roleFilter !== 'ALL' || distinctRoles.size <= 1) {
      let targetQty = 0;
      let soldStandardQty = 0;
      let remainingQty = 0;
      for (const r of filteredRows) {
        targetQty += r.targetQty;
        soldStandardQty += r.soldStandardQty;
        remainingQty += r.remainingQty;
      }
      targetQty = roundQty(targetQty);
      soldStandardQty = roundQty(soldStandardQty);
      remainingQty = roundQty(remainingQty);
      const achievementPct =
        targetQty > 0 && Number.isFinite(soldStandardQty / targetQty)
          ? Number(((soldStandardQty / targetQty) * 100).toFixed(2))
          : null;
      return { targetQty, soldStandardQty, remainingQty, achievementPct, isDeduplicated: false };
    }

    if (!searchText.trim()) {
      const targetQty = liveSummary.totalTargetQty;
      const soldStandardQty = liveSummary.totalStandardSoldQty;
      const remainingQty = targetQty > 0 ? Math.max(0, roundQty(targetQty - soldStandardQty)) : 0;
      return {
        targetQty,
        soldStandardQty,
        remainingQty,
        achievementPct: liveSummary.overallAchievementPct,
        isDeduplicated: true,
      };
    }

    // Search narrowed rows across multiple hierarchical tiers: compute non-overlapping sector rollup + sold by lowest visible tier per sector
    const singleSectors = new Set(
      filteredRows.filter((r) => r.roleKind !== 'department_manager').map((r) => r.sectorName).filter(Boolean)
    );
    let targetQty = 0;
    let soldStandardQty = 0;
    if (singleSectors.size > 0) {
      for (const sec of singleSectors) {
        targetQty += computeSectorRollupFromList(filteredRows, sec);
        const secRows = filteredRows.filter((r) => r.sectorName === sec && r.roleKind !== 'department_manager');
        const repRows = secRows.filter((r) => r.roleKind === 'sales_rep');
        const supRows = secRows.filter((r) => r.roleKind === 'supervisor');
        const secMgrRows = secRows.filter((r) => r.roleKind === 'sector_manager');
        const tierRows = secMgrRows.length > 0 ? secMgrRows : supRows.length > 0 ? supRows : repRows;
        for (const tr of tierRows) soldStandardQty += tr.soldStandardQty;
      }
    } else {
      for (const r of filteredRows) {
        targetQty += r.targetQty;
        soldStandardQty += r.soldStandardQty;
      }
    }
    targetQty = roundQty(targetQty);
    soldStandardQty = roundQty(soldStandardQty);
    const remainingQty = targetQty > 0 ? Math.max(0, roundQty(targetQty - soldStandardQty)) : 0;
    const achievementPct =
      targetQty > 0 && Number.isFinite(soldStandardQty / targetQty)
        ? Number(((soldStandardQty / targetQty) * 100).toFixed(2))
        : null;
    return { targetQty, soldStandardQty, remainingQty, achievementPct, isDeduplicated: true };
  }, [filteredRows, roleFilter, searchText, liveSummary]);

  const getExportMeta = () => {
    const roleFilterLabels: Record<RoleFilter, string> = {
      ALL: 'جميع الأدوار الوظيفية',
      department_manager: 'مدير إدارة',
      sector_manager: 'مدير قطاع',
      supervisor: 'مشرف قطاع',
      sales_rep: 'مندوب مبيعات',
    };

    const meta = buildActiveFilterMeta({
      source: 'ZSALES',
      sector: selectedSector,
      startDate,
      endDate,
      extra: [
        ...(competition ? [{ label: 'المسابقة', value: competition.name, fileTag: competition.name }] : []),
        roleFilter !== 'ALL'
          ? { label: 'الدور الوظيفي', value: roleFilterLabels[roleFilter], fileTag: roleFilterLabels[roleFilter] }
          : { label: 'الدور الوظيفي', value: 'جميع الأدوار الوظيفية' },
        ...(searchText.trim() ? [{ label: 'بحث', value: searchText.trim(), fileTag: searchText.trim() }] : []),
      ],
    });

    const columns: ExportColumn[] = [
      { header: 'الاسم', key: 'employeeName', type: 'string', minWidth: 22 },
      { header: 'المسمى الوظيفي / الدور', key: 'roleLabel', type: 'string', minWidth: 16 },
      { header: 'القطاع', key: 'sectorName', type: 'string', minWidth: 18 },
      { header: 'الكود', key: 'employeeCode', type: 'string', minWidth: 12 },
      { header: 'الهدف الكمي', key: 'targetQty', type: 'qty', minWidth: 14 },
      { header: 'الكمية المباعة', key: 'soldStandardQty', type: 'qty', minWidth: 14 },
      { header: 'الكمية المتبقية', key: 'remainingQty', type: 'qty', minWidth: 14 },
      { header: 'نسبة التحقيق', key: 'achievementPct', type: 'percent', minWidth: 13 },
    ];

    const exportRows = filteredRows.map((r) => ({
      employeeName: r.employeeName,
      roleLabel: r.roleLabel,
      sectorName: r.sectorName,
      employeeCode: r.employeeCode || '—',
      targetQty: r.targetQty,
      soldStandardQty: r.soldStandardQty,
      remainingQty: r.remainingQty,
      achievementPct: r.achievementPct,
    }));

    const totalRow =
      filteredRows.length > 0
        ? {
            employeeName: exportTotals.isDeduplicated
              ? `الإجمالي غير المكرر (${formatNum(filteredRows.length)} مشارك)`
              : `الإجمالي (${formatNum(filteredRows.length)} مشارك)`,
            roleLabel: '',
            sectorName: '',
            employeeCode: '',
            targetQty: exportTotals.targetQty,
            soldStandardQty: exportTotals.soldStandardQty,
            remainingQty: exportTotals.remainingQty,
            achievementPct: exportTotals.achievementPct,
          }
        : null;

    const itemRows = selectedItemDrafts.map((it) => {
      const cat = catalogByCode.get(it.itemCode);
      const factorNum = Number(it.factorStr);
      return {
        itemCode: it.itemCode,
        itemName: it.itemName,
        category: cat?.chainsCategory || cat?.categoryName || '—',
        factor: Number.isFinite(factorNum) ? factorNum : 1,
      };
    });

    return { meta, columns, exportRows, totalRow, itemRows };
  };

  const handleExportExcel = () => {
    if (!competition) return;
    const { meta, columns, exportRows, totalRow, itemRows } = getExportMeta();

    const itemColumns: ExportColumn[] = [
      { header: 'كود الصنف', key: 'itemCode', type: 'string', minWidth: 16 },
      { header: 'اسم الصنف', key: 'itemName', type: 'string', minWidth: 32 },
      { header: 'التصنيف', key: 'category', type: 'string', minWidth: 20 },
      { header: 'معامل الشدة المعيارية', key: 'factor', type: 'qty', minWidth: 20 },
    ];

    const itemsSummaryText =
      selectedItemDrafts.length > 0
        ? selectedItemDrafts.map((i) => `${i.itemName} (×${i.factorStr || '1'})`).join('، ')
        : 'لا توجد أصناف مختارة';

    exportToExcel({
      filename: buildExportFilename('مسابقات', meta.fileTags, 'xlsx'),
      sheets: [
        {
          name: 'أداء المسابقة',
          title: `مسابقة: ${competition.name}`,
          subtitle: `${meta.summary.join(' | ')} — الأصناف: ${itemsSummaryText}`,
          summaryPairs: [
            { label: 'اسم المسابقة', value: competition.name, type: 'string' },
            { label: 'الأصناف المختارة بالمسابقة', value: liveSummary.totalSelectedItems, type: 'integer' },
            { label: 'إجمالي الكمية المباعة المعيارية', value: liveSummary.totalStandardSoldQty, type: 'qty' },
            { label: 'الهدف الكمي الإجمالي', value: liveSummary.totalTargetQty, type: 'qty' },
            { label: 'نسبة التحقيق الإجمالية', value: liveSummary.overallAchievementPct, type: 'percent' },
          ],
          columns,
          rows: exportRows,
          totalRow,
        },
        {
          name: 'الأصناف والشدة المعيارية',
          title: `الأصناف المختارة ومعاملات الشدة المعيارية — ${competition.name}`,
          subtitle: `عدد الأصناف: ${formatNum(itemRows.length)}`,
          columns: itemColumns,
          rows: itemRows,
        },
      ],
    });
  };

  const handleExportPdf = async () => {
    if (!competition) return;
    const { meta, columns, exportRows, totalRow } = getExportMeta();
    const itemStrings =
      selectedItemDrafts.length > 0
        ? selectedItemDrafts.map((it) => `${it.itemName} (${it.itemCode}) × معامل ${it.factorStr || '1'}`)
        : ['لم يتم اختيار أصناف بعد'];

    await exportToPdf({
      filename: buildExportFilename('مسابقات', meta.fileTags, 'pdf'),
      tabTitle: `مسابقات — ${competition.name}`,
      filtersSummary: meta.summary,
      orientation: 'landscape',
      kpis: [
        {
          label: 'الأصناف المختارة بالمسابقة',
          value: formatNum(liveSummary.totalSelectedItems),
        },
        {
          label: 'إجمالي الكمية المباعة المعيارية',
          value: formatQty(liveSummary.totalStandardSoldQty),
          sublabel: 'كرتونة معيارية (Zsales_Price)',
        },
        {
          label: 'الهدف الكمي الإجمالي',
          value: formatQty(liveSummary.totalTargetQty),
        },
        {
          label: 'نسبة التحقيق الإجمالية',
          value: formatAchievement(liveSummary.overallAchievementPct),
          tone:
            liveSummary.overallAchievementPct !== null && liveSummary.overallAchievementPct >= 100
              ? 'positive'
              : 'accent',
        },
      ],
      infoBlocks: [
        {
          title: `أصناف المسابقة ومعاملات الشدة المعيارية (${formatNum(selectedItemDrafts.length)} صنف)`,
          items: itemStrings,
        },
      ],
      tables: [
        {
          title: 'جدول أداء المسابقة بالكميات',
          columns,
          rows: exportRows,
          totalRow,
        },
      ],
    });
  };

  if (loading && !data) {
    return <div className="h-64 bg-white border border-slate-200 rounded-xl animate-pulse" />;
  }

  return (
    <div className="space-y-5">
      {/* Zsales quantity source notice */}
      {requestedSource === 'VF05N' && (
        <div className="bg-slate-50 border border-slate-200 rounded-xl px-4 py-3 text-xs text-slate-700 flex items-center gap-2.5">
          <Info className="w-4 h-4 text-teal-700 shrink-0" />
          <span>
            تُحتسب كميات المسابقات (<span className="font-semibold font-num">QTY</span>) حصريًا من شيت{' '}
            <span className="font-semibold">Zsales_Price (2)</span> وتُضرب تلقائيًا في معامل الشدة المعيارية المحدد لكل صنف.
          </span>
        </div>
      )}

      {/* Feedback Banner */}
      {feedback && (
        <div
          className={`px-4 py-3 rounded-xl border text-xs font-medium flex items-center justify-between gap-3 ${
            feedback.type === 'success'
              ? 'bg-emerald-50 border-emerald-200 text-emerald-900'
              : 'bg-rose-50 border-rose-200 text-rose-900'
          }`}
        >
          <span>{feedback.text}</span>
          <button
            onClick={() => setFeedback(null)}
            className="p-1 hover:opacity-75 cursor-pointer"
            aria-label="إغلاق التنبيه"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        </div>
      )}

      {/* Competition Management Bar */}
      <div className="bg-white border border-slate-200 rounded-xl p-5 space-y-4">
        <div className="flex flex-wrap items-center justify-between gap-4">
          <div className="flex flex-wrap items-center gap-3">
            <div className="flex items-center gap-2">
              <Trophy className="w-4 h-4 text-teal-700 shrink-0" />
              <span className="text-sm font-bold text-slate-900">المسابقة النشطة:</span>
            </div>

            {competitions.length > 0 ? (
              <select
                aria-label="اختيار المسابقة"
                value={competition?.id ?? ''}
                onChange={(e) => {
                  const id = Number(e.target.value);
                  onSelectCompetition(Number.isFinite(id) && id > 0 ? id : null);
                }}
                className="px-3 py-2 text-xs font-semibold border border-slate-300 rounded-lg bg-white text-slate-900 focus:outline-none focus:ring-2 focus:ring-teal-600 min-w-[220px]"
              >
                {competitions.map((c) => (
                  <option key={c.id} value={c.id}>
                    {c.name} ({c.items.length} صنف)
                  </option>
                ))}
              </select>
            ) : (
              <span className="text-xs text-slate-500">لا توجد مسابقات منشأة بعد</span>
            )}

            {isAdmin && competition && !renamingComp && (
              <button
                onClick={() => {
                  setRenameDraft(competition.name);
                  setRenamingComp(true);
                }}
                className="px-2.5 py-1.5 text-xs font-semibold text-slate-600 hover:text-slate-900 border border-slate-300 rounded-lg hover:bg-slate-50 cursor-pointer"
              >
                تعديل الاسم
              </button>
            )}

            {isAdmin && competition && renamingComp && (
              <div className="flex items-center gap-2">
                <input
                  type="text"
                  value={renameDraft}
                  onChange={(e) => setRenameDraft(e.target.value)}
                  className="px-3 py-1.5 text-xs border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-teal-600 w-52"
                />
                <button
                  onClick={handleSaveRename}
                  disabled={busyAction}
                  className="px-3 py-1.5 text-xs font-semibold text-white bg-teal-700 hover:bg-teal-800 rounded-lg cursor-pointer disabled:opacity-50"
                >
                  حفظ
                </button>
                <button
                  onClick={() => setRenamingComp(false)}
                  className="px-2.5 py-1.5 text-xs text-slate-600 hover:text-slate-900 cursor-pointer"
                >
                  إلغاء
                </button>
              </div>
            )}
          </div>

          <div className="flex flex-wrap items-center gap-2.5">
            {competition && (
              <button
                onClick={() => setShowItemsPanel((v) => !v)}
                className="inline-flex items-center gap-2 px-3.5 py-2 text-xs font-semibold border border-slate-300 rounded-lg bg-white text-slate-700 hover:bg-slate-50 cursor-pointer"
              >
                <span>
                  {isAdmin
                    ? `إدارة الأصناف والشدة المعيارية (${selectedItemDrafts.length})`
                    : `عرض الأصناف والشدة المعيارية (${selectedItemDrafts.length})`}
                </span>
                {showItemsPanel ? <ChevronUp className="w-3.5 h-3.5" /> : <ChevronDown className="w-3.5 h-3.5" />}
              </button>
            )}

            {isAdmin && !creatingNew && (
              <button
                onClick={() => {
                  setCreatingNew(true);
                  setFeedback(null);
                }}
                className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-semibold text-white bg-teal-700 hover:bg-teal-800 rounded-lg transition-colors cursor-pointer"
              >
                <Plus className="w-3.5 h-3.5" />
                <span>إنشاء مسابقة جديدة</span>
              </button>
            )}

            {isAdmin && competition && confirmDeleteId !== competition.id && (
              <button
                onClick={() => setConfirmDeleteId(competition.id)}
                className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold text-rose-700 border border-rose-200 rounded-lg hover:bg-rose-50 cursor-pointer"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>حذف</span>
              </button>
            )}

            {isAdmin && competition && confirmDeleteId === competition.id && (
              <div className="flex items-center gap-2 bg-rose-50 border border-rose-200 px-3 py-1.5 rounded-lg text-xs">
                <span className="text-rose-900 font-semibold">تأكيد الحذف؟</span>
                <button
                  onClick={() => handleDeleteCompetition(competition.id)}
                  disabled={busyAction}
                  className="px-2.5 py-1 bg-rose-700 text-white rounded font-semibold hover:bg-rose-800 cursor-pointer"
                >
                  نعم، احذف
                </button>
                <button
                  onClick={() => setConfirmDeleteId(null)}
                  className="px-2 py-1 text-slate-700 hover:text-slate-900 cursor-pointer"
                >
                  تراجع
                </button>
              </div>
            )}
          </div>
        </div>

        {/* Inline Create Competition Form */}
        {isAdmin && creatingNew && (
          <form onSubmit={handleCreateCompetition} className="pt-3 border-t border-slate-200 flex flex-wrap items-center gap-3">
            <input
              type="text"
              value={newCompName}
              onChange={(e) => setNewCompName(e.target.value)}
              placeholder="اسم المسابقة (مثال: مسابقة الشدة المعيارية - البسكويت والويفر)"
              className="flex-1 min-w-[260px] px-3.5 py-2 text-xs border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-teal-600"
              autoFocus
            />
            <button
              type="submit"
              disabled={busyAction}
              className="px-4 py-2 text-xs font-semibold text-white bg-teal-700 hover:bg-teal-800 rounded-lg cursor-pointer disabled:opacity-50"
            >
              إنشاء واختيار الأصناف
            </button>
            <button
              type="button"
              onClick={() => {
                setCreatingNew(false);
                setNewCompName('');
              }}
              className="px-3 py-2 text-xs font-semibold text-slate-600 hover:text-slate-900 cursor-pointer"
            >
              إلغاء
            </button>
          </form>
        )}
      </div>

      {/* Empty state when no competition exists */}
      {!competition ? (
        <div className="bg-white border border-slate-200 rounded-xl p-10 text-center space-y-3">
          <Trophy className="w-8 h-8 text-slate-400 mx-auto" />
          <div className="text-sm font-bold text-slate-900">لم يتم إنشاء أي مسابقة بعد</div>
          <p className="text-xs text-slate-500 max-w-md mx-auto">
            أنشئ مسابقة جديدة وحدد مجموعة الأصناف ومعامل الشدة المعيارية لكل صنف ثم وزّع الأهداف الكمية مباشرة على فريق المبيعات.
          </p>
          {isAdmin && !creatingNew && (
            <button
              onClick={() => setCreatingNew(true)}
              className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold text-white bg-teal-700 hover:bg-teal-800 rounded-lg cursor-pointer"
            >
              <Plus className="w-4 h-4" />
              <span>إنشاء أول مسابقة</span>
            </button>
          )}
        </div>
      ) : (
        <>
          {/* 4-Column Summary Strip */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <div className="bg-white border border-slate-200 rounded-xl p-4">
              <div className="text-xs font-medium text-slate-500">الأصناف المختارة بالمسابقة</div>
              <div className="text-2xl font-bold text-slate-900 font-num tabular-nums mt-1.5" dir="ltr">
                {formatNum(liveSummary.totalSelectedItems)}
              </div>
              <div className="text-[11px] text-slate-500 mt-1">من تعريف الأصناف</div>
            </div>

            <div className="bg-white border border-slate-200 rounded-xl p-4">
              <div className="text-xs font-medium text-slate-500">إجمالي الكمية المباعة المعيارية</div>
              <div className="text-2xl font-bold text-slate-900 font-num tabular-nums mt-1.5" dir="ltr">
                {formatQty(liveSummary.totalStandardSoldQty)}
              </div>
              <div className="text-[11px] text-slate-500 mt-1">كرتونة معيارية · من Zsales_Price (2)</div>
            </div>

            <div className="bg-white border border-slate-200 rounded-xl p-4">
              <div className="text-xs font-medium text-slate-500">الهدف الكمي الإجمالي</div>
              <div className="text-2xl font-bold text-slate-900 font-num tabular-nums mt-1.5" dir="ltr">
                {formatQty(liveSummary.totalTargetQty)}
              </div>
              <div className="text-[11px] text-slate-500 mt-1">
                {selectedSector === 'ALL' ? 'إجمالي القطاعات الظاهرة' : `قطاع ${selectedSector}`}
              </div>
            </div>

            <div className="bg-white border border-slate-200 rounded-xl p-4">
              <div className="text-xs font-medium text-slate-500">نسبة التحقيق الإجمالية</div>
              <div
                className={`text-2xl font-bold font-num tabular-nums mt-1.5 ${
                  liveSummary.overallAchievementPct === null
                    ? 'text-slate-400'
                    : liveSummary.overallAchievementPct >= 100
                      ? 'text-emerald-700'
                      : 'text-teal-700'
                }`}
                dir="ltr"
              >
                {formatAchievement(liveSummary.overallAchievementPct)}
              </div>
              <div className="text-[11px] text-slate-500 mt-1">الكمية المباعة المعيارية ÷ الهدف الكمي</div>
            </div>
          </div>

          {/* Item Multi-Select + Standard Strength Factors Panel */}
          {showItemsPanel && (
            <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
              <div className="p-5 border-b border-slate-200 flex flex-wrap items-center justify-between gap-4">
                <div>
                  <h3 className="text-sm font-bold text-slate-900">أصناف المسابقة ومعامل الشدة المعيارية</h3>
                  <p className="text-xs text-slate-500 mt-0.5">
                    اختر الأصناف من تعريف الأصناف ثم حدد لكل صنف: الكرتونة من هذا الصنف تساوي كام كرتونة من الشدة المعيارية؟
                  </p>
                </div>
                {isAdmin && (
                  <button
                    onClick={handleSaveItems}
                    disabled={busyAction}
                    className={`inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold rounded-lg transition-colors cursor-pointer disabled:opacity-50 ${
                      itemsDirty
                        ? 'bg-teal-700 hover:bg-teal-800 text-white'
                        : 'bg-slate-800 hover:bg-slate-900 text-white'
                    }`}
                  >
                    <Save className="w-3.5 h-3.5" />
                    <span>{itemsDirty ? 'حفظ الأصناف والمعاملات *' : 'حفظ الأصناف والمعاملات'}</span>
                  </button>
                )}
              </div>

              <div className="grid grid-cols-1 lg:grid-cols-12 divide-y lg:divide-y-0 lg:divide-x lg:divide-x-reverse divide-slate-200">
                {/* Left/Right 1: Multi-select Item Picker */}
                {isAdmin && (
                  <div className="lg:col-span-5 p-4 space-y-3">
                    <div className="flex items-center justify-between gap-2">
                      <span className="text-xs font-bold text-slate-800">
                        قائمة تعريف الأصناف ({formatNum(catalogItems.length)})
                      </span>
                      <div className="flex items-center gap-2 text-[11px]">
                        <button
                          type="button"
                          onClick={selectAllFilteredCatalog}
                          className="text-teal-700 hover:underline font-semibold cursor-pointer"
                        >
                          تحديد الظاهر
                        </button>
                        <span className="text-slate-300">|</span>
                        <button
                          type="button"
                          onClick={clearSelectedItems}
                          className="text-slate-600 hover:text-rose-700 hover:underline cursor-pointer"
                        >
                          مسح الكل
                        </button>
                      </div>
                    </div>

                    <div className="relative">
                      <Search className="w-4 h-4 text-slate-400 absolute right-3 top-1/2 -translate-y-1/2" />
                      <input
                        type="text"
                        value={itemSearch}
                        onChange={(e) => setItemSearch(e.target.value)}
                        placeholder="بحث سريع باسم الصنف أو الكود..."
                        className="w-full pr-9 pl-3 py-2 text-xs border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-teal-600"
                      />
                    </div>

                    <div className="max-h-80 overflow-y-auto border border-slate-200 rounded-lg divide-y divide-slate-100">
                      {filteredCatalog.length === 0 ? (
                        <div className="p-6 text-center text-xs text-slate-500">
                          لا توجد أصناف مطابقة للبحث في جدول تعريف الأصناف
                        </div>
                      ) : (
                        filteredCatalog.map((item) => {
                          const checked = selectedCodesSet.has(item.itemCode);
                          const subParts = [item.itemCode, item.chainsCategory || item.categoryName, item.unitName]
                            .filter(Boolean)
                            .join(' · ');
                          return (
                            <label
                              key={item.itemCode}
                              className={`flex items-start gap-3 px-3 py-2.5 text-xs cursor-pointer transition-colors ${
                                checked ? 'bg-teal-50/60' : 'hover:bg-slate-50'
                              }`}
                            >
                              <input
                                type="checkbox"
                                checked={checked}
                                onChange={() => toggleCatalogItem(item)}
                                className="mt-0.5 h-4 w-4 rounded border-slate-300 text-teal-700 focus:ring-teal-600"
                              />
                              <div className="min-w-0 flex-1">
                                <div className="font-semibold text-slate-900 truncate">{item.itemName}</div>
                                <div className="text-[11px] text-slate-500 font-num mt-0.5 truncate">{subParts}</div>
                              </div>
                              {checked && <Check className="w-3.5 h-3.5 text-teal-700 shrink-0 mt-0.5" />}
                            </label>
                          );
                        })
                      )}
                    </div>
                  </div>
                )}

                {/* Selected Items & Standard Strength Factor Inputs */}
                <div className={`${isAdmin ? 'lg:col-span-7' : 'lg:col-span-12'} p-4 space-y-3`}>
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-800">
                      الأصناف المختارة ({formatNum(selectedItemDrafts.length)})
                    </span>
                    <span className="text-[11px] text-slate-500">
                      الكرتونة من هذا الصنف تساوي كام كرتونة من الشدة المعيارية؟ (يدعم حتى 3 أرقام عشرية مثل 0.333، 0.667، 1.250)
                    </span>
                  </div>

                  {selectedItemDrafts.length === 0 ? (
                    <div className="border border-dashed border-slate-300 rounded-lg p-8 text-center text-xs text-slate-500">
                      اختر صنفاً واحداً على الأقل من القائمة لتفعيل احتساب الكمية المعيارية للمسابقة.
                    </div>
                  ) : (
                    <div className="max-h-96 overflow-y-auto border border-slate-200 rounded-lg">
                      <table className="w-full text-xs text-right">
                        <thead className="bg-slate-50 border-b border-slate-200 text-slate-600 font-semibold sticky top-0">
                          <tr>
                            <th className="py-2.5 px-3 whitespace-nowrap">كود الصنف</th>
                            <th className="py-2.5 px-3">اسم الصنف</th>
                            <th className="py-2.5 px-3 whitespace-nowrap">معامل الشدة المعيارية</th>
                            {isAdmin && <th className="py-2.5 px-3 text-center whitespace-nowrap">إزالة</th>}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {selectedItemDrafts.map((item) => {
                            const cat = catalogByCode.get(item.itemCode);
                            return (
                              <tr key={item.itemCode} className="hover:bg-slate-50/70">
                                <td className="py-2 px-3 font-num text-slate-600 whitespace-nowrap" dir="ltr">
                                  {item.itemCode}
                                </td>
                                <td className="py-2 px-3 font-medium text-slate-900">
                                  <div>{item.itemName}</div>
                                  {cat?.chainsCategory && (
                                    <div className="text-[11px] text-slate-500">{cat.chainsCategory}</div>
                                  )}
                                </td>
                                <td className="py-2 px-3 whitespace-nowrap">
                                  {isAdmin ? (
                                    <input
                                      type="number"
                                      step="0.001"
                                      min="0"
                                      aria-label={`معامل الشدة المعيارية للصنف ${item.itemName}`}
                                      value={item.factorStr}
                                      onChange={(e) => updateFactorStr(item.itemCode, e.target.value)}
                                      className="w-28 px-2.5 py-1.5 text-xs font-num tabular-nums border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-teal-600 bg-white"
                                      dir="ltr"
                                    />
                                  ) : (
                                    <span className="font-num tabular-nums font-semibold text-slate-800" dir="ltr">
                                      {item.factorStr}
                                    </span>
                                  )}
                                </td>
                                {isAdmin && (
                                  <td className="py-2 px-3 text-center">
                                    <button
                                      type="button"
                                      onClick={() => {
                                        setItemsDirty(true);
                                        setSelectedItemDrafts((prev) =>
                                          prev.filter((x) => x.itemCode !== item.itemCode)
                                        );
                                      }}
                                      className="p-1 text-slate-400 hover:text-rose-700 cursor-pointer"
                                      aria-label={`إزالة ${item.itemName}`}
                                    >
                                      <X className="w-4 h-4" />
                                    </button>
                                  </td>
                                )}
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              </div>
            </div>
          )}

          {/* Leaderboard Performance Table */}
          <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
            <div className="p-5 border-b border-slate-200 flex flex-wrap items-center justify-between gap-4">
              <div>
                <h3 className="text-sm font-bold text-slate-900">جدول أداء المسابقة بالكميات</h3>
                <div className="flex flex-wrap items-center gap-2 text-xs text-slate-500 mt-1">
                  <span>
                    المشاركون الظاهرون:{' '}
                    <span className="font-num font-semibold text-slate-700">{formatNum(filteredRows.length)}</span>
                  </span>
                  <span aria-hidden="true">·</span>
                  <span>
                    القطاع:{' '}
                    <span className="font-semibold text-slate-700">
                      {selectedSector === 'ALL' ? 'جميع القطاعات' : selectedSector}
                    </span>
                  </span>
                  <span aria-hidden="true">·</span>
                  <span>المصدر: Zsales_Price (2)</span>
                </div>
              </div>

              <div className="flex flex-wrap items-center gap-3">
                <select
                  aria-label="تصفية الدور الوظيفي"
                  value={roleFilter}
                  onChange={(e) => setRoleFilter(e.target.value as RoleFilter)}
                  className="px-3 py-2 text-xs border border-slate-300 rounded-lg bg-white text-slate-700 focus:outline-none focus:ring-2 focus:ring-teal-600"
                >
                  <option value="ALL">جميع الأدوار الوظيفية</option>
                  <option value="department_manager">مدير إدارة</option>
                  <option value="sector_manager">مدير قطاع</option>
                  <option value="supervisor">مشرف قطاع</option>
                  <option value="sales_rep">مندوب مبيعات</option>
                </select>

                <div className="relative">
                  <Search className="w-4 h-4 text-slate-400 absolute right-3 top-1/2 -translate-y-1/2" />
                  <input
                    type="text"
                    value={searchText}
                    onChange={(e) => setSearchText(e.target.value)}
                    placeholder="بحث بالاسم أو القطاع..."
                    className="pr-9 pl-3 py-2 text-xs border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-teal-600 w-56"
                  />
                </div>

                {canEditTargets && (
                  <button
                    onClick={handleSaveTargets}
                    disabled={busyAction}
                    className={`inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold rounded-lg transition-colors cursor-pointer disabled:opacity-50 ${
                      targetsDirty
                        ? 'bg-teal-700 hover:bg-teal-800 text-white'
                        : 'bg-slate-800 hover:bg-slate-900 text-white'
                    }`}
                  >
                    <Save className="w-3.5 h-3.5" />
                    <span>{targetsDirty ? 'حفظ الأهداف الكمية *' : 'حفظ الأهداف الكمية'}</span>
                  </button>
                )}

                <ExportButtons
                  onExportExcel={handleExportExcel}
                  onExportPdf={handleExportPdf}
                  disabled={loading || !competition || filteredRows.length === 0}
                />
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-sm text-right">
                <thead className="bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-600">
                  <tr>
                    <th className="py-3 px-4 whitespace-nowrap">الاسم</th>
                    <th className="py-3 px-4 whitespace-nowrap">المسمى الوظيفي / الدور</th>
                    <th className="py-3 px-4 whitespace-nowrap">
                      <button
                        onClick={() => toggleSort('targetQty')}
                        className="inline-flex items-center gap-1 cursor-pointer hover:text-slate-900"
                      >
                        <span>الهدف الكمي</span>
                        <ArrowUpDown className="w-3.5 h-3.5" />
                      </button>
                    </th>
                    <th className="py-3 px-4 whitespace-nowrap">
                      <button
                        onClick={() => toggleSort('soldStandardQty')}
                        className="inline-flex items-center gap-1 cursor-pointer hover:text-slate-900"
                      >
                        <span>الكمية المباعة</span>
                        <ArrowUpDown className="w-3.5 h-3.5" />
                      </button>
                    </th>
                    <th className="py-3 px-4 whitespace-nowrap">
                      <button
                        onClick={() => toggleSort('remainingQty')}
                        className="inline-flex items-center gap-1 cursor-pointer hover:text-slate-900"
                      >
                        <span>الكمية المتبقية</span>
                        <ArrowUpDown className="w-3.5 h-3.5" />
                      </button>
                    </th>
                    <th className="py-3 px-4 whitespace-nowrap">
                      <button
                        onClick={() => toggleSort('achievementPct')}
                        className="inline-flex items-center gap-1 cursor-pointer hover:text-slate-900"
                      >
                        <span>نسبة التحقيق</span>
                        <ArrowUpDown className="w-3.5 h-3.5" />
                      </button>
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {filteredRows.length === 0 ? (
                    <tr>
                      <td colSpan={6} className="py-10 text-center text-xs text-slate-500">
                        لا توجد نتائج مطابقة للفلاتر المحددة
                      </td>
                    </tr>
                  ) : (
                    filteredRows.map((r) => (
                      <tr
                        key={r.participantKey}
                        className={r.roleKind === 'department_manager' ? 'bg-slate-50/70 hover:bg-slate-100/70' : 'hover:bg-slate-50/70'}
                      >
                        <td className="py-2.5 px-4 whitespace-nowrap">
                          <div className="font-semibold text-slate-900">{r.employeeName}</div>
                          <div className="text-[11px] text-slate-500 mt-0.5">
                            {r.sectorName}
                            {r.employeeCode ? ` · ${r.employeeCode}` : ''}
                          </div>
                        </td>

                        <td className="py-2.5 px-4 whitespace-nowrap">
                          <div className="text-xs font-semibold text-slate-800">{r.roleLabel}</div>
                          {r.isAutoTarget && (
                            <div className="text-[11px] text-teal-700 mt-0.5">تجميع تلقائي لأهداف النطاق</div>
                          )}
                        </td>

                        <td className="py-2.5 px-4 whitespace-nowrap">
                          {r.isAutoTarget || !canEditTargets ? (
                            <span
                              className="font-num tabular-nums font-semibold text-slate-900 inline-block"
                              dir="ltr"
                            >
                              {formatQty(r.targetQty)}
                            </span>
                          ) : (
                            <input
                              type="number"
                              step="0.001"
                              min="0"
                              aria-label={`الهدف الكمي لـ ${r.employeeName}`}
                              value={targetDrafts[r.participantKey] ?? String(r.targetQty)}
                              onChange={(e) => {
                                setFeedback(null);
                                setTargetsDirty(true);
                                setTargetDrafts((prev) => ({
                                  ...prev,
                                  [r.participantKey]: e.target.value,
                                }));
                              }}
                              className="w-28 px-2.5 py-1.5 text-xs font-num tabular-nums border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-teal-600 bg-white"
                              dir="ltr"
                            />
                          )}
                        </td>

                        <td
                          className={`py-2.5 px-4 font-num tabular-nums font-semibold whitespace-nowrap ${
                            r.soldStandardQty === 0 ? 'text-slate-500' : 'text-slate-900'
                          }`}
                          dir="ltr"
                        >
                          {formatQty(r.soldStandardQty)}
                        </td>

                        <td className="py-2.5 px-4 font-num tabular-nums text-slate-700 whitespace-nowrap" dir="ltr">
                          {formatQty(r.remainingQty)}
                        </td>

                        <td
                          className={`py-2.5 px-4 font-num tabular-nums font-semibold whitespace-nowrap ${
                            r.achievementPct === null
                              ? 'text-slate-400'
                              : r.achievementPct >= 100
                                ? 'text-emerald-700'
                                : 'text-teal-700'
                          }`}
                          dir="ltr"
                        >
                          {formatAchievement(r.achievementPct)}
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
                {filteredRows.length > 0 && (
                  <tfoot className="bg-sky-50 border-t-2 border-slate-300 font-bold text-slate-900">
                    <tr>
                      <td colSpan={2} className="py-2.5 px-4 whitespace-nowrap">
                        {exportTotals.isDeduplicated
                          ? `الإجمالي غير المكرر (${formatNum(filteredRows.length)} مشارك)`
                          : `الإجمالي (${formatNum(filteredRows.length)} مشارك)`}
                      </td>
                      <td className="py-2.5 px-4 font-num tabular-nums whitespace-nowrap" dir="ltr">
                        {formatQty(exportTotals.targetQty)}
                      </td>
                      <td className="py-2.5 px-4 font-num tabular-nums whitespace-nowrap" dir="ltr">
                        {formatQty(exportTotals.soldStandardQty)}
                      </td>
                      <td className="py-2.5 px-4 font-num tabular-nums whitespace-nowrap" dir="ltr">
                        {formatQty(exportTotals.remainingQty)}
                      </td>
                      <td className="py-2.5 px-4 font-num tabular-nums text-teal-900 whitespace-nowrap" dir="ltr">
                        {formatAchievement(exportTotals.achievementPct)}
                      </td>
                    </tr>
                  </tfoot>
                )}
              </table>
            </div>
          </div>
        </>
      )}
    </div>
  );
};
