export type SalesSource = 'VF05N' | 'ZSALES';
export type UserRole = 'admin' | 'sector_manager' | 'standard_user';

export interface UserProfile {
  id: number;
  username: string;
  fullName: string;
  email: string | null;
  role: UserRole;
  allowedSectors: string[];
  isActive: boolean;
  mustChangePassword?: boolean;
  createdAt: string;
}

export interface AuthStatusResponse {
  needsSetup: boolean;
  hasEnvPasscode: boolean;
}

export interface LoginResponse {
  token: string;
  user: UserProfile;
}

export interface SectorMetricRow {
  sectorName: string;
  sales: number;
  tax: number;
  target: number;
  variance: number;
  achievementPct: number | null;
  recordCount: number;
  returnCount: number;
  returnValue: number;
  isOriginalSector: boolean;
}

export interface AggregationMetricRow {
  key: 'private_markets' | 'direct_and_army' | 'chains_sector';
  title: string;
  subtitle: string;
  includedSectors: string[];
  sales: number;
  tax: number;
  target: number;
  variance: number;
  achievementPct: number | null;
  isPartialForUser: boolean;
}

export interface SourceComparisonSummary {
  vf05n: {
    count: number;
    salesTotal: number;
    taxTotal: number;
    returnCount: number;
  };
  zsales: {
    count: number;
    salesTotal: number;
    taxTotal: number;
    returnCount: number;
  };
  salesDifference: number;
  taxDifference: number;
  isMatched: boolean;
}

export interface ActiveBatchInfo {
  id: number;
  batchUuid: string;
  filename: string;
  importedByName: string;
  createdAt: string;
  vf05nCount: number;
  vf05nSalesTotal: number;
  vf05nTaxTotal: number;
  zsalesCount: number;
  zsalesSalesTotal: number;
  zsalesTaxTotal: number;
  unlinkedBranchCount: number;
  unlinkedSectorCount: number;
}

export interface DashboardSummaryResponse {
  hasData: boolean;
  activeBatch: ActiveBatchInfo | null;
  source: SalesSource;
  dateRange: {
    minDate: string | null;
    maxDate: string | null;
  };
  kpis: {
    privateMarketsSales: number;
    privateMarketsTarget: number;
    privateMarketsTax: number;
    privateMarketsVariance: number;
    privateMarketsAchievementPct: number | null;
    directAndArmySales: number;
    directAndArmyTarget: number;
    directAndArmyAchievementPct: number | null;
    chainsSectorSales: number;
    chainsSectorTarget: number;
    chainsSectorAchievementPct: number | null;
    unlinkedSectorSales: number;
    unlinkedSectorCount: number;
    unlinkedBranchSales: number;
    unlinkedBranchCount: number;
    totalSourceRecords: number;
    totalSourceSales: number;
    totalSourceTax: number;
  };
  sectors: SectorMetricRow[];
  aggregations: AggregationMetricRow[];
  timeline: TimelinePoint[];
  sourceComparison: SourceComparisonSummary | null;
}

export interface DepartmentChainRow {
  chainCode: string;
  chainName: string;
  departmentName: string;
  sectorName: string;
  supervisorName: string;
  commercialSales: number;
  industrialSales: number;
  sales: number;
  tax: number;
  recordCount: number;
  shareOfSectorPct: number;
  shareOfSupervisorPct: number;
}

export interface BranchRepBreakdown {
  repCode: string;
  repName: string;
  sales: number;
  tax: number;
  recordCount: number;
}

export interface BranchMetricRow {
  branchCode: string;
  branchName: string;
  sectorName: string;
  supervisorName: string;
  chainCode: string;
  chainName: string;
  assignedRepCode: string;
  assignedRepName: string;
  repBreakdowns: BranchRepBreakdown[];
  target: number;
  sales: number;
  tax: number;
  variance: number;
  achievementPct: number | null;
  recordCount: number;
  returnCount: number;
}

export interface UnlinkedRecordRow {
  id: number;
  docNumber: string;
  billingDate: string;
  payerCode: string;
  payerName: string;
  soldToCode: string;
  soldToName: string;
  resolvedSector: string | null;
  salesValue: number;
  taxValue: number;
  itemCode: string;
  itemName: string;
}

export interface RepMetricRow {
  repCode: string;
  repName: string;
  sectorName: string;
  supervisorName: string;
  sales: number;
  tax: number;
  target: number;
  achievementPct: number | null;
  recordCount: number;
  returnCount: number;
  branchCount: number;
}

export interface BranchRepDiagnostics {
  totalTargetedAssignedBranches: number;
  previousInflatedCountInBranchesView: number;
  differenceCount: number;
  consolidatedDuplicateNameCount: number;
  consolidatedDuplicateNames: Array<{
    branchName: string;
    sectorName: string;
    supervisorName: string;
    repCode: string;
    branchCodes: string[];
    combinedSales: number;
    combinedTarget: number;
  }>;
  repDiscrepancies: Array<{
    repCode: string;
    repName: string;
    expectedAssignedTargetedCount: number;
    previousBranchesViewCount: number;
    extraCount: number;
    extraCrossRepBranches: Array<{
      branchCode: string;
      branchName: string;
      target: number;
      assignedRepCode: string;
      assignedRepName: string;
      supervisorName: string;
    }>;
  }>;
}

export interface ItemMetricRow {
  itemCode: string;
  itemName: string;
  categoryName: string;
  brandName: string;
  quantity: number;
  sales: number;
  tax: number;
  recordCount: number;
  returnCount: number;
  returnValue: number;
}

export interface ItemCategoryAnalysisRow {
  targetKey: string;
  sectorName: string;
  jobTitle: string;
  employeeCode: string;
  employeeName: string;
  categoryName: string;
  sales: number;
  tax: number;
  target: number;
  achievementPct: number | null;
  isAutoTarget: boolean;
}

export interface ItemCategoriesAnalyticsResponse {
  rows: ItemCategoryAnalysisRow[];
  jobTitleOptions: string[];
  hasCategoryTargetsSheet: boolean;
  effectiveSource: SalesSource;
}

export interface TimelinePoint {
  date: string;
  sales: number;
  tax: number;
  grossSales: number;
  returnsValue: number;
  recordCount: number;
  cumulativeSales: number;
}

export interface SalesTransactionRow {
  id: number;
  sourceType: SalesSource;
  docNumber: string;
  billingDate: string;
  payerCode: string;
  payerName: string;
  soldToCode: string;
  soldToName: string;
  itemCode: string;
  itemName: string;
  repCode: string;
  repName: string;
  quantity: number;
  salesValue: number;
  taxValue: number;
  resolvedSector: string | null;
  resolvedDepartment: string | null;
  resolvedChainName: string | null;
  resolvedBranchCode: string | null;
  resolvedBranchName: string | null;
  isReturn: boolean;
}

export interface ValidationIssue {
  severity: 'error' | 'warning';
  sheet: string;
  code: string;
  message: string;
  details?: string[];
}

export interface SheetCheckSummary {
  sheetName: string;
  found: boolean;
  rowCount: number;
  detectedColumns: string[];
  missingColumns: string[];
}

export interface ImportValidationReport {
  validationToken: string;
  filename: string;
  valid: boolean;
  canCommit: boolean;
  errors: ValidationIssue[];
  warnings: ValidationIssue[];
  sheets: SheetCheckSummary[];
  comparison: SourceComparisonSummary;
  linkStats: {
    vf05nTotalRows: number;
    vf05nSectorLinkedRows: number;
    vf05nSectorUnlinkedRows: number;
    vf05nBranchLinkedRows: number;
    vf05nBranchUnlinkedRows: number;
    unlinkedBranchCodes: Array<{ soldToCode: string; soldToName: string; rowCount: number; salesTotal: number }>;
    unlinkedSectorCodes: Array<{ payerCode: string; payerName: string; rowCount: number; salesTotal: number }>;
  };
  targetsSummary: {
    sectors: Array<{ sectorName: string; target: number }>;
    totalSectorTargets: number;
    directBranchesTargetTotal: number;
    directBranchCount: number;
  };
  usersFound: Array<{ username: string; fullName: string; role: UserRole; allowedSectors: string[] }>;
}

export interface ImportBatchHistoryItem {
  id: number;
  batchUuid: string;
  filename: string;
  importedByName: string;
  status: 'active' | 'archived';
  vf05nCount: number;
  vf05nSalesTotal: number;
  vf05nTaxTotal: number;
  zsalesCount: number;
  zsalesSalesTotal: number;
  zsalesTaxTotal: number;
  unlinkedBranchCount: number;
  unlinkedSectorCount: number;
  createdAt: string;
}

export type ClassAType = 'industrial' | 'commercial' | 'other';

export interface CompetitionCatalogItem {
  itemCode: string;
  itemName: string;
  categoryName: string;
  classAType: ClassAType;
  chainsCategory: string;
  armyCategory: string;
  topRetailCategory: string;
  unitName: string;
}

export interface CompetitionItemConfig {
  itemCode: string;
  itemName: string;
  equivalentFactor: number;
  classificationA?: string;
  classAType?: ClassAType;
}

export interface CompetitionSummary {
  id: number;
  name: string;
  isActive: boolean;
  items: CompetitionItemConfig[];
  targetCount: number;
  createdAt: string;
  updatedAt: string;
}

export interface CompetitionParticipantRow {
  participantKey: string;
  employeeCode: string;
  employeeName: string;
  roleLabel: string;
  roleKind: 'department_manager' | 'sector_manager' | 'supervisor' | 'sales_rep';
  sectorName: string;
  managedSectors: string[];
  assignedCategories: string[];
  isAutoTarget: boolean;
  targetQty: number;
  soldStandardQty: number;
  remainingQty: number;
  achievementPct: number | null;
  industrialSoldQty: number;
  commercialSoldQty: number;
  industrialTargetQty: number;
  commercialTargetQty: number;
  industrialAchievementPct: number | null;
  commercialAchievementPct: number | null;
}

export interface CompetitionAnalyticsResponse {
  competition: CompetitionSummary | null;
  competitions: CompetitionSummary[];
  catalogItems: CompetitionCatalogItem[];
  participants: CompetitionParticipantRow[];
  summary: {
    totalSelectedItems: number;
    totalTeamMembers: number;
    totalStandardSoldQty: number;
    totalTargetQty: number;
    overallAchievementPct: number | null;
    totalIndustrialSoldQty: number;
    totalCommercialSoldQty: number;
    totalIndustrialTargetQty: number;
    totalCommercialTargetQty: number;
    industrialAchievementPct: number | null;
    commercialAchievementPct: number | null;
  };
}

export type AuditSeverity = 'critical' | 'high' | 'medium' | 'low';
export type AuditCategory = 'numeric_integrity' | 'data_consistency' | 'rbac_permissions' | 'runtime_error';
export type AuditFixType = 'auto_fixable' | 'code_fix_required' | 'info';

export type AutoFixActionType =
  | 'FIX_NON_FINITE_NUMBERS'
  | 'FIX_DUPLICATE_BRANCH_RECORDS'
  | 'FIX_DUPLICATE_ITEM_CATALOG'
  | 'FIX_SYNC_COMPETITION_TARGET_TOTALS'
  | 'FIX_CANONICALIZE_TARGETS_AND_SCOPES'
  | 'FIX_RESOLVE_RUNTIME_ERRORS';

export interface AuditFinding {
  id: string;
  code: string;
  title: string;
  description: string;
  severity: AuditSeverity;
  category: AuditCategory;
  fixType: AuditFixType;
  affectedArea: string;
  details: string[];
  autoFixAction?: AutoFixActionType;
  autoFixLabel?: string;
  readyPrompt?: string;
  aiInsight?: string;
}

export interface RuntimeErrorLog {
  id: number;
  kind: 'runtime_error' | 'unhandled_rejection' | 'react_error';
  tabName: string;
  message: string;
  stack: string;
  sourceUrl: string;
  lineNo: number;
  colNo: number;
  username: string;
  resolved: boolean;
  aiAnalysis: string;
  occurredAt: string;
}

export interface AuditFixBackup {
  id: number;
  findingCode: string;
  actionType: string;
  actionLabel: string;
  summary: string;
  createdByName: string;
  createdAt: string;
  restoredAt: string | null;
}

export interface SmartAuditReport {
  runId: number;
  ranAt: string;
  triggeredByName: string;
  healthScore: number;
  aiConfigured: boolean;
  aiSummary: string;
  activeBatchId: number | null;
  activeBatchFilename: string | null;
  checksExecuted: Array<{
    key: string;
    label: string;
    passed: boolean;
    issueCount: number;
  }>;
  summaryCounts: {
    total: number;
    critical: number;
    high: number;
    medium: number;
    low: number;
    autoFixable: number;
    codeFixRequired: number;
    unresolvedRuntimeErrors: number;
  };
  findings: AuditFinding[];
  runtimeErrors: RuntimeErrorLog[];
  backups: AuditFixBackup[];
}


