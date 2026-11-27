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
  ClassAType,
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
type SortKey =
  | 'default'
  | 'industrialSoldQty'
  | 'commercialSoldQty'
  | 'industrialTargetQty'
  | 'commercialTargetQty'
  | 'industrialAchievementPct'
  | 'commercialAchievementPct';

interface TargetSplitDraft {
  ind: string;
  com: string;
}

const EMPTY_COMPETITIONS: CompetitionAnalyticsResponse['competitions'] = [];
const EMPTY_CATALOG_ITEMS: CompetitionCatalogItem[] = [];
const EMPTY_PARTICIPANTS: CompetitionParticipantRow[] = [];

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

function safeAchievementPct(sold: number, target: number): number | null {
  if (!Number.isFinite(sold) || !Number.isFinite(target) || target <= 0) return null;
  const pct = (sold / target) * 100;
  if (!Number.isFinite(pct)) return null;
  return Number(pct.toFixed(2));
}

function formatClassALabel(classAType?: ClassAType, raw?: string): string {
  if (classAType === 'industrial') return 'صناعى';
  if (classAType === 'commercial') return 'تجارى';
  return raw?.trim() || '—';
}

function computeSectorRollupSplitFromList(
  rows: Array<{
    roleKind: CompetitionParticipantRow['roleKind'];
    sectorName: string;
    industrialTargetQty: number;
    commercialTargetQty: number;
  }>,
  sector: string
): { ind: number; com: number; total: number } {
  let secMgrInd = 0;
  let secMgrCom = 0;
  let supInd = 0;
  let supCom = 0;
  let repInd = 0;
  let repCom = 0;

  for (const r of rows) {
    if (r.sectorName !== sector) continue;
    if (r.roleKind === 'sector_manager') {
      secMgrInd += r.industrialTargetQty;
      secMgrCom += r.commercialTargetQty;
    } else if (r.roleKind === 'supervisor') {
      supInd += r.industrialTargetQty;
      supCom += r.commercialTargetQty;
    } else if (r.roleKind === 'sales_rep') {
      repInd += r.industrialTargetQty;
      repCom += r.commercialTargetQty;
    }
  }

  if (secMgrInd > 0 || secMgrCom > 0) {
    const ind = roundQty(secMgrInd);
    const com = roundQty(secMgrCom);
    return { ind, com, total: roundQty(ind + com) };
  }
  if (supInd > 0 || supCom > 0) {
    const ind = roundQty(supInd);
    const com = roundQty(supCom);
    return { ind, com, total: roundQty(ind + com) };
  }
  const ind = roundQty(repInd);
  const com = roundQty(repCom);
  return { ind, com, total: roundQty(ind + com) };
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
  onSelectCompetition,
  onRefresh,
}) => {
  const competition = data?.competition ?? null;
  const competitions = data?.competitions ?? EMPTY_COMPETITIONS;
  const catalogItems = data?.catalogItems ?? EMPTY_CATALOG_ITEMS;
  const participants = data?.participants ?? EMPTY_PARTICIPANTS;

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

  // Dual target qty drafts per participantKey (`صناعى` & `تجارى`)
  const [targetDrafts, setTargetDrafts] = useState<Record<string, TargetSplitDraft>>({});
  const [targetsDirty, setTargetsDirty] = useState(false);

  // Table filter & sort
  const [roleFilter, setRoleFilter] = useState<RoleFilter>('ALL');
  const [searchText, setSearchText] = useState('');
  const [sortKey, setSortKey] = useState<SortKey>('default');
  const [sortAsc, setSortAsc] = useState(false);

  // Sync local item drafts whenever competition changes
  useEffect(() => {
    if (!competition) {
      setSelectedItemDrafts((prev) => (prev.length === 0 ? prev : []));
      setItemsDirty((prev) => (prev ? false : prev));
      setRenamingComp((prev) => (prev ? false : prev));
      setConfirmDeleteId((prev) => (prev !== null ? null : prev));
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

  // Sync local dual target drafts when participants load/refresh
  useEffect(() => {
    const next: Record<string, TargetSplitDraft> = {};
    for (const p of participants) {
      if (!p.isAutoTarget) {
        next[p.participantKey] = {
          ind: String(p.industrialTargetQty ?? 0),
          com: String(p.commercialTargetQty ?? 0),
        };
      }
    }
    setTargetDrafts((prev) => {
      const prevKeys = Object.keys(prev);
      const nextKeys = Object.keys(next);
      if (
        prevKeys.length === nextKeys.length &&
        nextKeys.every((k) => prev[k]?.ind === next[k].ind && prev[k]?.com === next[k].com)
      ) {
        return prev;
      }
      return next;
    });
    setTargetsDirty((prev) => (prev ? false : prev));
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

  const selectedClassACounts = useMemo(() => {
    let industrial = 0;
    let commercial = 0;
    for (const it of selectedItemDrafts) {
      const cat = catalogByCode.get(it.itemCode);
      if (cat?.classAType === 'industrial') industrial++;
      else if (cat?.classAType === 'commercial') commercial++;
    }
    return { industrial, commercial };
  }, [selectedItemDrafts, catalogByCode]);

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
    const targetsPayload: Array<{
      participantKey: string;
      industrialTargetQty: number;
      commercialTargetQty: number;
    }> = [];

    for (const p of participants) {
      if (p.isAutoTarget) continue;
      const draft = targetDrafts[p.participantKey];
      const rawInd = (draft?.ind ?? String(p.industrialTargetQty ?? 0)).trim();
      const rawCom = (draft?.com ?? String(p.commercialTargetQty ?? 0)).trim();
      const parsedInd = rawInd === '' ? 0 : Number(rawInd);
      const parsedCom = rawCom === '' ? 0 : Number(rawCom);
      if (!Number.isFinite(parsedInd) || parsedInd < 0) {
        setFeedback({
          type: 'error',
          text: `قيمة هدف أصناف الصناعى غير صالحة للمشارك ${p.employeeName}`,
        });
        return;
      }
      if (!Number.isFinite(parsedCom) || parsedCom < 0) {
        setFeedback({
          type: 'error',
          text: `قيمة هدف الأصناف التجارى غير صالحة للمشارك ${p.employeeName}`,
        });
        return;
      }
      targetsPayload.push({
        participantKey: p.participantKey,
        industrialTargetQty: roundQty(parsedInd),
        commercialTargetQty: roundQty(parsedCom),
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
      setFeedback({ type: 'success', text: 'تم حفظ أهداف الصناعى والتجارى وتحديث إجماليات مديري الإدارات' });
    } catch (err) {
      setFeedback({ type: 'error', text: err instanceof Error ? err.message : 'تعذر حفظ الأهداف الكمية' });
    } finally {
      setBusyAction(false);
    }
  };

  // Compute live participants with instant client-side recalculation for dual targets, %s, and مدير إدارة rollup
  const liveParticipants = useMemo(() => {
    const subRowsOld = participants.filter((p) => !p.isAutoTarget);
    const subRowsNew = subRowsOld.map((p) => {
      const draft = targetDrafts[p.participantKey];
      const rawInd = draft?.ind;
      const rawCom = draft?.com;
      const parsedInd = rawInd === undefined ? p.industrialTargetQty : rawInd.trim() === '' ? 0 : Number(rawInd);
      const parsedCom = rawCom === undefined ? p.commercialTargetQty : rawCom.trim() === '' ? 0 : Number(rawCom);
      const validInd = Number.isFinite(parsedInd) && parsedInd >= 0 ? roundQty(parsedInd) : p.industrialTargetQty;
      const validCom = Number.isFinite(parsedCom) && parsedCom >= 0 ? roundQty(parsedCom) : p.commercialTargetQty;
      return {
        ...p,
        liveIndTarget: validInd,
        liveComTarget: validCom,
      };
    });

    const newSplitByKey = new Map<string, { ind: number; com: number }>();
    for (const sr of subRowsNew) {
      newSplitByKey.set(sr.participantKey, { ind: sr.liveIndTarget, com: sr.liveComTarget });
    }

    const sectorsInView = Array.from(new Set(subRowsOld.map((r) => r.sectorName).filter(Boolean)));
    const deltaBySector = new Map<string, { ind: number; com: number }>();
    for (const sec of sectorsInView) {
      const oldRoll = computeSectorRollupSplitFromList(subRowsOld, sec);
      const newRoll = computeSectorRollupSplitFromList(
        subRowsNew.map((r) => ({
          roleKind: r.roleKind,
          sectorName: r.sectorName,
          industrialTargetQty: r.liveIndTarget,
          commercialTargetQty: r.liveComTarget,
        })),
        sec
      );
      deltaBySector.set(sec, {
        ind: newRoll.ind - oldRoll.ind,
        com: newRoll.com - oldRoll.com,
      });
    }

    return participants.map((p) => {
      let industrialTargetQty = p.industrialTargetQty;
      let commercialTargetQty = p.commercialTargetQty;
      if (p.isAutoTarget) {
        let deltaInd = 0;
        let deltaCom = 0;
        for (const sec of p.managedSectors) {
          const d = deltaBySector.get(sec);
          if (d) {
            deltaInd += d.ind;
            deltaCom += d.com;
          }
        }
        industrialTargetQty = Math.max(0, roundQty(p.industrialTargetQty + deltaInd));
        commercialTargetQty = Math.max(0, roundQty(p.commercialTargetQty + deltaCom));
      } else {
        const split = newSplitByKey.get(p.participantKey);
        if (split) {
          industrialTargetQty = split.ind;
          commercialTargetQty = split.com;
        }
      }

      const targetQty =
        industrialTargetQty > 0 || commercialTargetQty > 0
          ? roundQty(industrialTargetQty + commercialTargetQty)
          : p.targetQty;
      const remainingQty = targetQty > 0 ? Math.max(0, roundQty(targetQty - p.soldStandardQty)) : 0;

      return {
        ...p,
        industrialTargetQty,
        commercialTargetQty,
        targetQty,
        remainingQty,
        achievementPct: safeAchievementPct(p.soldStandardQty, targetQty),
        industrialAchievementPct: safeAchievementPct(p.industrialSoldQty, industrialTargetQty),
        commercialAchievementPct: safeAchievementPct(p.commercialSoldQty, commercialTargetQty),
      };
    });
  }, [participants, targetDrafts]);

  // Live summary across visible sectors (non-overlapping)
  const liveSummary = useMemo(() => {
    const subRows = liveParticipants.filter((p) => !p.isAutoTarget);
    const sectors = Array.from(new Set(subRows.map((p) => p.sectorName).filter(Boolean)));
    let totalIndTarget = 0;
    let totalComTarget = 0;
    for (const sec of sectors) {
      const r = computeSectorRollupSplitFromList(subRows, sec);
      totalIndTarget += r.ind;
      totalComTarget += r.com;
    }
    const roundedIndTarget = roundQty(totalIndTarget);
    const roundedComTarget = roundQty(totalComTarget);
    const roundedTarget = roundQty(roundedIndTarget + roundedComTarget);

    const indSoldTotal = data?.summary.totalIndustrialSoldQty ?? 0;
    const comSoldTotal = data?.summary.totalCommercialSoldQty ?? 0;
    const soldTotal = data?.summary.totalStandardSoldQty ?? 0;

    return {
      totalSelectedItems: selectedItemDrafts.length,
      totalIndustrialSoldQty: indSoldTotal,
      totalCommercialSoldQty: comSoldTotal,
      totalStandardSoldQty: soldTotal,
      totalIndustrialTargetQty: roundedIndTarget,
      totalCommercialTargetQty: roundedComTarget,
      totalTargetQty: roundedTarget,
      industrialAchievementPct: safeAchievementPct(indSoldTotal, roundedIndTarget),
      commercialAchievementPct: safeAchievementPct(comSoldTotal, roundedComTarget),
      overallAchievementPct: safeAchievementPct(soldTotal, roundedTarget),
    };
  }, [
    liveParticipants,
    data?.summary.totalIndustrialSoldQty,
    data?.summary.totalCommercialSoldQty,
    data?.summary.totalStandardSoldQty,
    selectedItemDrafts.length,
  ]);

  const filteredRows = useMemo(() => {
    const q = searchText.trim().toLowerCase();
    const list = liveParticipants.filter((r) => {
      if (roleFilter !== 'ALL' && r.roleKind !== roleFilter) return false;
      if (!q) return true;
      return (
        r.employeeName.toLowerCase().includes(q) ||
        r.roleLabel.toLowerCase().includes(q) ||
        r.sectorName.toLowerCase().includes(q) ||
        r.employeeCode.toLowerCase().includes(q) ||
        r.assignedCategories.some((c) => c.toLowerCase().includes(q))
      );
    });

    if (sortKey === 'default') return list;

    return [...list].sort((a, b) => {
      if (sortKey === 'industrialAchievementPct' || sortKey === 'commercialAchievementPct') {
        const av = a[sortKey] ?? -Infinity;
        const bv = b[sortKey] ?? -Infinity;
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
      let industrialSoldQty = 0;
      let commercialSoldQty = 0;
      let industrialTargetQty = 0;
      let commercialTargetQty = 0;
      for (const r of filteredRows) {
        industrialSoldQty += r.industrialSoldQty;
        commercialSoldQty += r.commercialSoldQty;
        industrialTargetQty += r.industrialTargetQty;
        commercialTargetQty += r.commercialTargetQty;
      }
      industrialSoldQty = roundQty(industrialSoldQty);
      commercialSoldQty = roundQty(commercialSoldQty);
      industrialTargetQty = roundQty(industrialTargetQty);
      commercialTargetQty = roundQty(commercialTargetQty);
      return {
        industrialSoldQty,
        commercialSoldQty,
        industrialTargetQty,
        commercialTargetQty,
        industrialAchievementPct: safeAchievementPct(industrialSoldQty, industrialTargetQty),
        commercialAchievementPct: safeAchievementPct(commercialSoldQty, commercialTargetQty),
        isDeduplicated: false,
      };
    }

    if (!searchText.trim()) {
      return {
        industrialSoldQty: liveSummary.totalIndustrialSoldQty,
        commercialSoldQty: liveSummary.totalCommercialSoldQty,
        industrialTargetQty: liveSummary.totalIndustrialTargetQty,
        commercialTargetQty: liveSummary.totalCommercialTargetQty,
        industrialAchievementPct: liveSummary.industrialAchievementPct,
        commercialAchievementPct: liveSummary.commercialAchievementPct,
        isDeduplicated: true,
      };
    }

    const singleSectors = new Set(
      filteredRows.filter((r) => r.roleKind !== 'department_manager').map((r) => r.sectorName).filter(Boolean)
    );
    let industrialTargetQty = 0;
    let commercialTargetQty = 0;
    let industrialSoldQty = 0;
    let commercialSoldQty = 0;

    if (singleSectors.size > 0) {
      for (const sec of singleSectors) {
        const roll = computeSectorRollupSplitFromList(filteredRows, sec);
        industrialTargetQty += roll.ind;
        commercialTargetQty += roll.com;
        const secRows = filteredRows.filter((r) => r.sectorName === sec && r.roleKind !== 'department_manager');
        const repRows = secRows.filter((r) => r.roleKind === 'sales_rep');
        const supRows = secRows.filter((r) => r.roleKind === 'supervisor');
        const secMgrRows = secRows.filter((r) => r.roleKind === 'sector_manager');
        const tierRows = repRows.length > 0 ? repRows : supRows.length > 0 ? supRows : secMgrRows;
        for (const tr of tierRows) {
          industrialSoldQty += tr.industrialSoldQty;
          commercialSoldQty += tr.commercialSoldQty;
        }
      }
    } else {
      for (const r of filteredRows) {
        industrialTargetQty += r.industrialTargetQty;
        commercialTargetQty += r.commercialTargetQty;
        industrialSoldQty += r.industrialSoldQty;
        commercialSoldQty += r.commercialSoldQty;
      }
    }
    industrialSoldQty = roundQty(industrialSoldQty);
    commercialSoldQty = roundQty(commercialSoldQty);
    industrialTargetQty = roundQty(industrialTargetQty);
    commercialTargetQty = roundQty(commercialTargetQty);
    return {
      industrialSoldQty,
      commercialSoldQty,
      industrialTargetQty,
      commercialTargetQty,
      industrialAchievementPct: safeAchievementPct(industrialSoldQty, industrialTargetQty),
      commercialAchievementPct: safeAchievementPct(commercialSoldQty, commercialTargetQty),
      isDeduplicated: true,
    };
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
      { header: 'مبيعات أصناف الصناعى', key: 'industrialSoldQty', type: 'qty', minWidth: 16 },
      { header: 'مبيعات الأصناف التجارى', key: 'commercialSoldQty', type: 'qty', minWidth: 16 },
      { header: 'هدف أصناف الصناعى', key: 'industrialTargetQty', type: 'qty', minWidth: 15 },
      { header: 'هدف أصناف التجارى', key: 'commercialTargetQty', type: 'qty', minWidth: 15 },
      { header: 'نسبة تحقيق الصناعى', key: 'industrialAchievementPct', type: 'percent', minWidth: 15 },
      { header: 'نسبة تحقيق التجارى', key: 'commercialAchievementPct', type: 'percent', minWidth: 15 },
    ];

    const exportRows = filteredRows.map((r) => ({
      employeeName:
        r.assignedCategories.length > 0
          ? `${r.employeeName} (${r.assignedCategories.join('، ')})`
          : r.employeeName,
      roleLabel: r.roleLabel,
      sectorName: r.sectorName,
      employeeCode: r.employeeCode || '—',
      industrialSoldQty: r.industrialSoldQty,
      commercialSoldQty: r.commercialSoldQty,
      industrialTargetQty: r.industrialTargetQty,
      commercialTargetQty: r.commercialTargetQty,
      industrialAchievementPct: r.industrialAchievementPct,
      commercialAchievementPct: r.commercialAchievementPct,
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
            industrialSoldQty: exportTotals.industrialSoldQty,
            commercialSoldQty: exportTotals.commercialSoldQty,
            industrialTargetQty: exportTotals.industrialTargetQty,
            commercialTargetQty: exportTotals.commercialTargetQty,
            industrialAchievementPct: exportTotals.industrialAchievementPct,
            commercialAchievementPct: exportTotals.commercialAchievementPct,
          }
        : null;

    const itemRows = selectedItemDrafts.map((it) => {
      const cat = catalogByCode.get(it.itemCode);
      const factorNum = Number(it.factorStr);
      return {
        itemCode: it.itemCode,
        itemName: it.itemName,
        classA: formatClassALabel(cat?.classAType, cat?.categoryName),
        category: cat?.chainsCategory || '—',
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
      { header: 'التصنيف A', key: 'classA', type: 'string', minWidth: 14 },
      { header: 'تصنيف السلاسل', key: 'category', type: 'string', minWidth: 20 },
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
            { label: 'مبيعات أصناف الصناعى', value: liveSummary.totalIndustrialSoldQty, type: 'qty' },
            { label: 'هدف أصناف الصناعى', value: liveSummary.totalIndustrialTargetQty, type: 'qty' },
            { label: 'نسبة تحقيق الصناعى', value: liveSummary.industrialAchievementPct, type: 'percent' },
            { label: 'مبيعات الأصناف التجارى', value: liveSummary.totalCommercialSoldQty, type: 'qty' },
            { label: 'هدف أصناف التجارى', value: liveSummary.totalCommercialTargetQty, type: 'qty' },
            { label: 'نسبة تحقيق التجارى', value: liveSummary.commercialAchievementPct, type: 'percent' },
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
        ? selectedItemDrafts.map((it) => {
            const cat = catalogByCode.get(it.itemCode);
            const classALabel = formatClassALabel(cat?.classAType, cat?.categoryName);
            return `${it.itemName} (${it.itemCode} · ${classALabel}) × معامل ${it.factorStr || '1'}`;
          })
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
          sublabel: `صناعى ${formatNum(selectedClassACounts.industrial)} · تجارى ${formatNum(selectedClassACounts.commercial)}`,
        },
        {
          label: 'أصناف الصناعى (مبيعات / هدف)',
          value: formatQty(liveSummary.totalIndustrialSoldQty),
          sublabel: `الهدف: ${formatQty(liveSummary.totalIndustrialTargetQty)} · التحقيق: ${formatAchievement(liveSummary.industrialAchievementPct)}`,
          tone:
            liveSummary.industrialAchievementPct !== null && liveSummary.industrialAchievementPct >= 100
              ? 'positive'
              : 'accent',
        },
        {
          label: 'الأصناف التجارى (مبيعات / هدف)',
          value: formatQty(liveSummary.totalCommercialSoldQty),
          sublabel: `الهدف: ${formatQty(liveSummary.totalCommercialTargetQty)} · التحقيق: ${formatAchievement(liveSummary.commercialAchievementPct)}`,
          tone:
            liveSummary.commercialAchievementPct !== null && liveSummary.commercialAchievementPct >= 100
              ? 'positive'
              : 'accent',
        },
        {
          label: 'نسبة التحقيق الإجمالية',
          value: formatAchievement(liveSummary.overallAchievementPct),
          sublabel: `${formatQty(liveSummary.totalStandardSoldQty)} ÷ ${formatQty(liveSummary.totalTargetQty)}`,
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
          title: 'جدول أداء المسابقة بالكميات (تصنيف A: صناعى / تجارى)',
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
            <span className="font-semibold">Zsales_Price (2)</span> وتُضرب تلقائيًا في معامل الشدة المعيارية المحدد لكل صنف،
            مع تقسيمها وفق <span className="font-semibold">التصنيف A (صناعى / تجارى)</span>.
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
            أنشئ مسابقة جديدة وحدد مجموعة الأصناف ومعامل الشدة المعيارية لكل صنف ثم وزّع أهداف الصناعى والتجارى مباشرة على فريق المبيعات.
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
              <div className="text-[11px] text-slate-500 mt-1">
                صناعى: <span className="font-num font-semibold text-slate-700">{formatNum(selectedClassACounts.industrial)}</span>
                {' · '}
                تجارى: <span className="font-num font-semibold text-slate-700">{formatNum(selectedClassACounts.commercial)}</span>
              </div>
            </div>

            <div className="bg-white border border-slate-200 rounded-xl p-4">
              <div className="text-xs font-medium text-slate-500">مبيعات أصناف الصناعى</div>
              <div className="text-2xl font-bold text-slate-900 font-num tabular-nums mt-1.5" dir="ltr">
                {formatQty(liveSummary.totalIndustrialSoldQty)}
              </div>
              <div className="text-[11px] text-slate-500 mt-1">
                الهدف: <span className="font-num font-semibold text-slate-700">{formatQty(liveSummary.totalIndustrialTargetQty)}</span>
                {' · '}
                التحقيق:{' '}
                <span className="font-num font-semibold text-teal-700">
                  {formatAchievement(liveSummary.industrialAchievementPct)}
                </span>
              </div>
            </div>

            <div className="bg-white border border-slate-200 rounded-xl p-4">
              <div className="text-xs font-medium text-slate-500">مبيعات الأصناف التجارى</div>
              <div className="text-2xl font-bold text-slate-900 font-num tabular-nums mt-1.5" dir="ltr">
                {formatQty(liveSummary.totalCommercialSoldQty)}
              </div>
              <div className="text-[11px] text-slate-500 mt-1">
                الهدف: <span className="font-num font-semibold text-slate-700">{formatQty(liveSummary.totalCommercialTargetQty)}</span>
                {' · '}
                التحقيق:{' '}
                <span className="font-num font-semibold text-teal-700">
                  {formatAchievement(liveSummary.commercialAchievementPct)}
                </span>
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
              <div className="text-[11px] text-slate-500 mt-1">
                المبيعات <span className="font-num font-semibold">{formatQty(liveSummary.totalStandardSoldQty)}</span>
                {' ÷ '}
                الهدف <span className="font-num font-semibold">{formatQty(liveSummary.totalTargetQty)}</span>
              </div>
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
                {/* Multi-select Item Picker */}
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
                        placeholder="بحث سريع باسم الصنف أو الكود أو التصنيف A..."
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
                          const classALabel = formatClassALabel(item.classAType, item.categoryName);
                          const subParts = [item.itemCode, classALabel, item.chainsCategory, item.unitName]
                            .filter((p) => p && p !== '—')
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
                  <div className="flex items-center justify-between flex-wrap gap-2">
                    <span className="text-xs font-bold text-slate-800">
                      الأصناف المختارة ({formatNum(selectedItemDrafts.length)})
                    </span>
                    <span className="text-[11px] text-slate-500">
                      الكرتونة من هذا الصنف تساوي كام كرتونة من الشدة المعيارية؟ (حتى 3 أرقام عشرية)
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
                            <th className="py-2.5 px-3 whitespace-nowrap">التصنيف A</th>
                            <th className="py-2.5 px-3 whitespace-nowrap">معامل الشدة المعيارية</th>
                            {isAdmin && <th className="py-2.5 px-3 text-center whitespace-nowrap">إزالة</th>}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {selectedItemDrafts.map((item) => {
                            const cat = catalogByCode.get(item.itemCode);
                            const classALabel = formatClassALabel(cat?.classAType, cat?.categoryName);
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
                                <td className="py-2 px-3 whitespace-nowrap font-semibold text-slate-700">
                                  {classALabel}
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
                        onClick={() => toggleSort('industrialSoldQty')}
                        className="inline-flex items-center gap-1 cursor-pointer hover:text-slate-900"
                      >
                        <span>مبيعات أصناف الصناعى</span>
                        <ArrowUpDown className="w-3.5 h-3.5" />
                      </button>
                    </th>
                    <th className="py-3 px-4 whitespace-nowrap">
                      <button
                        onClick={() => toggleSort('commercialSoldQty')}
                        className="inline-flex items-center gap-1 cursor-pointer hover:text-slate-900"
                      >
                        <span>مبيعات الأصناف التجارى</span>
                        <ArrowUpDown className="w-3.5 h-3.5" />
                      </button>
                    </th>
                    <th className="py-3 px-4 whitespace-nowrap">
                      <button
                        onClick={() => toggleSort('industrialTargetQty')}
                        className="inline-flex items-center gap-1 cursor-pointer hover:text-slate-900"
                      >
                        <span>هدف أصناف الصناعى</span>
                        <ArrowUpDown className="w-3.5 h-3.5" />
                      </button>
                    </th>
                    <th className="py-3 px-4 whitespace-nowrap">
                      <button
                        onClick={() => toggleSort('commercialTargetQty')}
                        className="inline-flex items-center gap-1 cursor-pointer hover:text-slate-900"
                      >
                        <span>هدف أصناف التجارى</span>
                        <ArrowUpDown className="w-3.5 h-3.5" />
                      </button>
                    </th>
                    <th className="py-3 px-4 whitespace-nowrap">
                      <button
                        onClick={() => toggleSort('industrialAchievementPct')}
                        className="inline-flex items-center gap-1 cursor-pointer hover:text-slate-900"
                      >
                        <span>نسبة تحقيق الصناعى</span>
                        <ArrowUpDown className="w-3.5 h-3.5" />
                      </button>
                    </th>
                    <th className="py-3 px-4 whitespace-nowrap">
                      <button
                        onClick={() => toggleSort('commercialAchievementPct')}
                        className="inline-flex items-center gap-1 cursor-pointer hover:text-slate-900"
                      >
                        <span>نسبة تحقيق التجارى</span>
                        <ArrowUpDown className="w-3.5 h-3.5" />
                      </button>
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {filteredRows.length === 0 ? (
                    <tr>
                      <td colSpan={8} className="py-10 text-center text-xs text-slate-500">
                        لا توجد نتائج مطابقة للفلاتر المحددة
                      </td>
                    </tr>
                  ) : (
                    filteredRows.map((r) => {
                      const draft = targetDrafts[r.participantKey];
                      return (
                        <tr
                          key={r.participantKey}
                          className={
                            r.roleKind === 'department_manager'
                              ? 'bg-slate-50/70 hover:bg-slate-100/70'
                              : 'hover:bg-slate-50/70'
                          }
                        >
                          <td className="py-2.5 px-4 whitespace-nowrap">
                            <div className="font-semibold text-slate-900">{r.employeeName}</div>
                            <div className="text-[11px] text-slate-500 mt-0.5">
                              {r.sectorName}
                              {r.employeeCode ? ` · ${r.employeeCode}` : ''}
                            </div>
                            {r.assignedCategories.length > 0 && (
                              <div className="text-[11px] text-teal-800 mt-0.5 whitespace-normal max-w-[240px]">
                                الفئات: {r.assignedCategories.join('، ')}
                              </div>
                            )}
                          </td>

                          <td className="py-2.5 px-4 whitespace-nowrap">
                            <div className="text-xs font-semibold text-slate-800">{r.roleLabel}</div>
                            {r.isAutoTarget && (
                              <div className="text-[11px] text-teal-700 mt-0.5">تجميع تلقائي لأهداف النطاق</div>
                            )}
                          </td>

                          <td
                            className={`py-2.5 px-4 font-num tabular-nums font-semibold whitespace-nowrap ${
                              r.industrialSoldQty === 0 ? 'text-slate-500' : 'text-slate-900'
                            }`}
                            dir="ltr"
                          >
                            {formatQty(r.industrialSoldQty)}
                          </td>

                          <td
                            className={`py-2.5 px-4 font-num tabular-nums font-semibold whitespace-nowrap ${
                              r.commercialSoldQty === 0 ? 'text-slate-500' : 'text-slate-900'
                            }`}
                            dir="ltr"
                          >
                            {formatQty(r.commercialSoldQty)}
                          </td>

                          <td className="py-2.5 px-4 whitespace-nowrap">
                            {r.isAutoTarget || !canEditTargets ? (
                              <span
                                className="font-num tabular-nums font-semibold text-slate-900 inline-block"
                                dir="ltr"
                              >
                                {formatQty(r.industrialTargetQty)}
                              </span>
                            ) : (
                              <input
                                type="number"
                                step="0.001"
                                min="0"
                                aria-label={`هدف أصناف الصناعى لـ ${r.employeeName}`}
                                value={draft?.ind ?? String(r.industrialTargetQty)}
                                onChange={(e) => {
                                  const val = e.target.value;
                                  setFeedback(null);
                                  setTargetsDirty(true);
                                  setTargetDrafts((prev) => ({
                                    ...prev,
                                    [r.participantKey]: {
                                      ind: val,
                                      com: prev[r.participantKey]?.com ?? String(r.commercialTargetQty),
                                    },
                                  }));
                                }}
                                className="w-28 px-2.5 py-1.5 text-xs font-num tabular-nums border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-teal-600 bg-white"
                                dir="ltr"
                              />
                            )}
                          </td>

                          <td className="py-2.5 px-4 whitespace-nowrap">
                            {r.isAutoTarget || !canEditTargets ? (
                              <span
                                className="font-num tabular-nums font-semibold text-slate-900 inline-block"
                                dir="ltr"
                              >
                                {formatQty(r.commercialTargetQty)}
                              </span>
                            ) : (
                              <input
                                type="number"
                                step="0.001"
                                min="0"
                                aria-label={`هدف أصناف التجارى لـ ${r.employeeName}`}
                                value={draft?.com ?? String(r.commercialTargetQty)}
                                onChange={(e) => {
                                  const val = e.target.value;
                                  setFeedback(null);
                                  setTargetsDirty(true);
                                  setTargetDrafts((prev) => ({
                                    ...prev,
                                    [r.participantKey]: {
                                      ind: prev[r.participantKey]?.ind ?? String(r.industrialTargetQty),
                                      com: val,
                                    },
                                  }));
                                }}
                                className="w-28 px-2.5 py-1.5 text-xs font-num tabular-nums border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-teal-600 bg-white"
                                dir="ltr"
                              />
                            )}
                          </td>

                          <td
                            className={`py-2.5 px-4 font-num tabular-nums font-semibold whitespace-nowrap ${
                              r.industrialAchievementPct === null
                                ? 'text-slate-400'
                                : r.industrialAchievementPct >= 100
                                  ? 'text-emerald-700'
                                  : 'text-teal-700'
                            }`}
                            dir="ltr"
                          >
                            {formatAchievement(r.industrialAchievementPct)}
                          </td>

                          <td
                            className={`py-2.5 px-4 font-num tabular-nums font-semibold whitespace-nowrap ${
                              r.commercialAchievementPct === null
                                ? 'text-slate-400'
                                : r.commercialAchievementPct >= 100
                                  ? 'text-emerald-700'
                                  : 'text-teal-700'
                            }`}
                            dir="ltr"
                          >
                            {formatAchievement(r.commercialAchievementPct)}
                          </td>
                        </tr>
                      );
                    })
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
                        {formatQty(exportTotals.industrialSoldQty)}
                      </td>
                      <td className="py-2.5 px-4 font-num tabular-nums whitespace-nowrap" dir="ltr">
                        {formatQty(exportTotals.commercialSoldQty)}
                      </td>
                      <td className="py-2.5 px-4 font-num tabular-nums whitespace-nowrap" dir="ltr">
                        {formatQty(exportTotals.industrialTargetQty)}
                      </td>
                      <td className="py-2.5 px-4 font-num tabular-nums whitespace-nowrap" dir="ltr">
                        {formatQty(exportTotals.commercialTargetQty)}
                      </td>
                      <td className="py-2.5 px-4 font-num tabular-nums text-teal-900 whitespace-nowrap" dir="ltr">
                        {formatAchievement(exportTotals.industrialAchievementPct)}
                      </td>
                      <td className="py-2.5 px-4 font-num tabular-nums text-teal-900 whitespace-nowrap" dir="ltr">
                        {formatAchievement(exportTotals.commercialAchievementPct)}
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
