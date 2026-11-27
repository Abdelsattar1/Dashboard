import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import Database from 'better-sqlite3';
import * as XLSX from 'xlsx';
import {
  calculateAchievementPct,
  getBranchesAnalytics,
  getDashboardSummary,
  getDepartmentsChainsBreakdown,
  getItemCategoriesAnalytics,
  getRepsAnalytics,
  saveItemCategoryTargets,
} from '../server/analytics-engine.ts';
import { AuthorizationError } from '../server/auth.ts';
import {
  calculateQtyAchievementPct,
  classifyClassA,
  createCompetition,
  deleteCompetition,
  getAvailableCatalogItems,
  getCompetitionAnalytics,
  listCompetitions,
  saveCompetitionTargets,
  updateCompetition,
} from '../server/competitions-engine.ts';
import { initDatabase } from '../server/db.ts';
import { commitValidatedImport, excelSerialToIsoDate, rollbackToBatch, validateExcelBuffer } from '../server/excel-engine.ts';
import { executeAutoFix, recordRuntimeError, rollbackAutoFix, runSmartAudit } from '../server/smart-auditor.ts';
import type { UserProfile } from '../src/types/bi.ts';

function buildReferenceTestWorkbook(): Buffer {
  const wb = XLSX.utils.book_new();

  // 1. Sectors
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      ['كود القطاع', 'القطاع'],
      ['SEC-1', 'التوريد المباشر'],
      ['SEC-2', 'التوريد المركزي'],
      ['SEC-3', 'التوب ريتيل'],
      ['SEC-4', 'الجيش'],
    ]),
    'القطاعات'
  );

  // 2. Sector Targets (Total = 65,100,000)
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      ['القطاع', 'الهدف'],
      ['التوريد المباشر', 15000000],
      ['التوريد المركزي', 25300000],
      ['التوب ريتيل', 13800000],
      ['الجيش', 11000000],
    ]),
    'اهداف القطاعات'
  );

  // 3. Chains & Departments (Payer -> Sector) - preserving leading zeros!
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      ['كود السلسلة', 'اسم السلسلة', 'الادارة', 'القطاع'],
      ['001001', 'عملاء التوريد المباشر', 'إدارة المباشر', 'التوريد المباشر'],
      ['002001', 'سلسلة هايبر المركزي', 'إدارة السلاسل', 'التوريد المركزي'],
      ['003001', 'سلسلة التوب ريتيل', 'إدارة التوب ريتيل', 'التوب ريتيل'],
      ['004001', 'منافذ الجيش', 'إدارة الجيش', 'الجيش'],
    ]),
    'تعريف الادارات والقطاعات'
  );

  // 4. Branches & Targets (Direct Supply ONLY, sum = 15,000,000)
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      ['كود الفرع', 'اسم الفرع', 'القطاع', 'الهدف'],
      ['000501', 'فرع القاهرة المباشر', 'التوريد المباشر', 9000000],
      ['000502', 'فرع الإسكندرية المباشر', 'التوريد المباشر', 6000000],
    ]),
    'تعريف الفروع و الاهداف'
  );

  // 5. Reps
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      ['كود المندوب', 'اسم المندوب', 'القطاع', 'المشرف'],
      ['REP01', 'أحمد محمود', 'التوريد المباشر', 'خالد علي'],
      ['REP02', 'محمد حسن', 'التوريد المركزي', 'خالد علي'],
    ]),
    'تعريف مندوبى البيع'
  );

  // 6. Items
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      ['كود الصنف', 'اسم الصنف', 'المجموعة', 'العلامة التجارية', 'الوحدة'],
      ['SKU001', 'منتج غذائي أ', 'أغذية', 'العلامة الأولى', 'كرتونة'],
      ['SKU002', 'منتج استهلاكي ب', 'مشروبات', 'العلامة الثانية', 'كرتونة'],
    ]),
    'تعريف الاصناف'
  );

  // 7. Users
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      ['اسم المستخدم', 'الاسم الكامل', 'البريد الإلكتروني', 'الصلاحية', 'القطاعات'],
      ['admin', 'مدير النظام', 'admin@company.com', 'مدير', 'الكل'],
      ['direct.user', 'مدير المباشر', 'direct@company.com', 'مستخدم قطاعي', 'التوريد المباشر'],
    ]),
    'تعريف المستخدمين'
  );

  // 8 & 9. Build 1,193 rows: 1,178 branch-linked + 15 unlinked branch rows
  // Total sales = 52,171,054.36, Total tax = 211,626.87, including negative return rows!
  const salesHeader = [
    'Billing Document',
    'Billing Date',
    'Payer',
    'Name of Payer',
    'Sold-to Party',
    'Name of Sold-to Party',
    'Material',
    'Material Description',
    'Sales Employee',
    'Billed Quantity',
    'New Growth Value BD&T',
    'Tax Amount',
  ];

  const salesRows: unknown[][] = [salesHeader];
  // Distribute across 1193 rows:
  // First 1177 linked rows of 40,000 each = 47,080,000
  // 1 return linked row = -50,000
  // 15 unlinked branch rows (Sold-to '999999' not in branches) totaling 5,141,054.36
  // Total = 47,080,000 - 50,000 + 5,141,054.36 = 52,171,054.36!
  const payers = ['001001', '002001', '003001', '004001'];
  for (let i = 0; i < 1177; i++) {
    const payer = payers[i % 4];
    const soldTo = i % 2 === 0 ? '000501' : '000502';
    const tax = i === 0 ? 111626.87 : 0;
    salesRows.push([
      `INV${String(i + 1).padStart(5, '0')}`,
      '2026-11-10',
      payer,
      `عميل ${payer}`,
      soldTo,
      `فرع ${soldTo}`,
      i % 2 === 0 ? 'SKU001' : 'SKU002',
      'صنف اختباري',
      i % 2 === 0 ? 'REP01' : 'REP02',
      10,
      40000,
      tax,
    ]);
  }
  // 1 negative return row (linked to branch 000501, payer 001001)
  salesRows.push([
    'RET0001',
    '2026-11-11',
    '001001',
    'عملاء التوريد المباشر',
    '000501',
    'فرع القاهرة المباشر',
    'SKU001',
    'صنف اختباري',
    'REP01',
    -2,
    -50000,
    0,
  ]);
  // 15 unlinked branch rows (Sold-to Party = '009999' not in branch master!)
  for (let j = 0; j < 15; j++) {
    const val = j === 0 ? 141054.36 : 357142.85714285716;
    const tax = j === 0 ? 100000 : 0;
    salesRows.push([
      `UNL${String(j + 1).padStart(4, '0')}`,
      '2026-11-12',
      '002001',
      'سلسلة هايبر المركزي',
      '009999',
      'نقطة توزيع غير مسجلة',
      'SKU002',
      'صنف اختباري',
      'REP02',
      5,
      Number(val.toFixed(2)),
      tax,
    ]);
  }

  // Adjust exact penny on row 1 of unlinked if rounding shifted by 0.01
  let currentTotal = 0;
  for (let r = 1; r < salesRows.length; r++) {
    currentTotal += Number(salesRows[r][10]);
  }
  const diff = Number((52171054.36 - currentTotal).toFixed(2));
  salesRows[salesRows.length - 1][10] = Number((Number(salesRows[salesRows.length - 1][10]) + diff).toFixed(2));

  // Build Zsales_Price (2) using Customer (instead of Payer), Sales Emp., Material, and NO Sold-to Party
  const zsalesHeader = [
    'Billing Document',
    'Billing Date',
    'Customer',
    'Customer Name',
    'Material',
    'Material Description',
    'Sales Emp.',
    'Billed Quantity',
    'New Growth Value BD&T',
    'Tax Amount',
  ];
  const zsalesRows: unknown[][] = [
    zsalesHeader,
    ...salesRows.slice(1).map((r) => [r[0], r[1], r[2], r[3], r[6], r[7], r[8], r[9], r[10], r[11]]),
  ];

  const wsVf = XLSX.utils.aoa_to_sheet(salesRows);
  const wsZs = XLSX.utils.aoa_to_sheet(zsalesRows);

  // Force string cell type on leading-zero codes and test numeric Excel serial date + SAP DD.MM.YYYY date
  for (let r = 1; r < salesRows.length; r++) {
    for (const colIdx of [2, 4]) {
      const ref = XLSX.utils.encode_cell({ r, c: colIdx });
      if (wsVf[ref]) {
        wsVf[ref].t = 's';
        wsVf[ref].v = String(salesRows[r][colIdx]);
      }
    }
    const zsCustRef = XLSX.utils.encode_cell({ r, c: 2 });
    if (wsZs[zsCustRef]) {
      wsZs[zsCustRef].t = 's';
      wsZs[zsCustRef].v = String(zsalesRows[r][2]);
    }
    if (r === 1) {
      wsVf[XLSX.utils.encode_cell({ r, c: 1 })] = { t: 'n', v: 46336 };
      wsZs[XLSX.utils.encode_cell({ r, c: 1 })] = { t: 'n', v: 46336 };
    } else if (r === 2) {
      wsVf[XLSX.utils.encode_cell({ r, c: 1 })] = { t: 's', v: '10.11.2026' };
      wsZs[XLSX.utils.encode_cell({ r, c: 1 })] = { t: 's', v: '10.11.2026' };
    }
  }

  XLSX.utils.book_append_sheet(wb, wsVf, 'VF05N (2)');
  XLSX.utils.book_append_sheet(wb, wsZs, 'Zsales_Price (2)');

  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Buffer;
}

function runTests() {
  console.log('Running Sales BI Engine Verification Tests...');
  assert.equal(excelSerialToIsoDate(46336), '2026-11-10', 'Excel serial date 46336 must convert to 2026-11-10');

  const db = initDatabase(':memory:');
  db.prepare(
    `INSERT INTO users (username, full_name, email, role, allowed_sectors, password_hash, created_at)
     VALUES ('admin', 'مدير النظام', 'admin@test.com', 'admin', '[]', 'salt:hash', '2026-11-23T00:00:00Z')`
  ).run();

  const adminUser: UserProfile = {
    id: 1,
    username: 'admin',
    fullName: 'مدير النظام',
    email: 'admin@test.com',
    role: 'admin',
    allowedSectors: ['التوريد المباشر', 'التوريد المركزي', 'التوب ريتيل', 'الجيش'],
    isActive: true,
    createdAt: '2026-11-23T00:00:00Z',
  };

  const directUser: UserProfile = {
    id: 2,
    username: 'direct.user',
    fullName: 'مستخدم المباشر',
    email: 'direct@test.com',
    role: 'standard_user',
    allowedSectors: ['التوريد المباشر'],
    isActive: true,
    createdAt: '2026-11-23T00:00:00Z',
  };

  // 1. Validate Excel Buffer
  const xlsxBuf = buildReferenceTestWorkbook();
  const report = validateExcelBuffer(xlsxBuf, 'Reference_Test.xlsx');

  assert.equal(report.canCommit, true, 'Expected workbook to pass validation');
  assert.equal(report.comparison.vf05n.count, 1193, 'Expected 1,193 rows in VF05N');
  assert.equal(report.comparison.zsales.count, 1193, 'Expected 1,193 rows in Zsales_Price');
  assert.equal(report.comparison.vf05n.salesTotal, 52171054.36, 'VF05N total must match 52,171,054.36');
  assert.equal(report.comparison.zsales.salesTotal, 52171054.36, 'Zsales_Price total must match 52,171,054.36');
  assert.equal(report.comparison.vf05n.taxTotal, 211626.87, 'Tax total must match 211,626.87');
  assert.equal(report.comparison.isMatched, true, 'VF05N and Zsales_Price totals must match');
  assert.equal(report.linkStats.vf05nBranchLinkedRows, 1178, 'Expected 1,178 branch-linked rows');
  assert.equal(report.linkStats.vf05nBranchUnlinkedRows, 15, 'Expected 15 unlinked branch rows');
  assert.equal(report.targetsSummary.totalSectorTargets, 65100000, 'Expected total sector targets = 65,100,000');
  assert.equal(report.targetsSummary.directBranchesTargetTotal, 15000000, 'Expected direct branch targets = 15,000,000');

  // 2. Commit Batch 1
  const batch1 = commitValidatedImport(db, report.validationToken, { id: 1, fullName: 'مدير النظام' });
  assert.equal(batch1.status, 'active');

  // 3. Verify Dashboard Calculation (Single Source, No Double Counting, Exact Formula)
  const dashVf = getDashboardSummary(db, adminUser, { source: 'VF05N' });
  const dashZs = getDashboardSummary(db, adminUser, { source: 'ZSALES' });

  assert.equal(dashVf.kpis.privateMarketsSales, 52171054.36, 'Must NOT sum VF05N and Zsales_Price together');
  assert.equal(dashZs.kpis.privateMarketsSales, 52171054.36, 'ZSALES must also equal 52,171,054.36');
  assert.equal(dashVf.kpis.privateMarketsTarget, 65100000, 'Target must remain 65,100,000 (branch targets not double-added)');
  assert.equal(dashVf.kpis.privateMarketsTax, 211626.87);

  // Verify Achievement Formula: Sales / (Target - Tax) * 100
  const expectedAch = calculateAchievementPct(52171054.36, 65100000, 211626.87);
  assert.equal(dashVf.kpis.privateMarketsAchievementPct, expectedAch);
  assert.equal(calculateAchievementPct(1000, 500, 600), null, 'Zero or negative denominator must return null');

  // Verify overlapping aggregations do NOT inflate Private Markets total and carry accurate tax sums
  const aggPrivate = dashVf.aggregations.find((a) => a.key === 'private_markets')!;
  const aggDirectArmy = dashVf.aggregations.find((a) => a.key === 'direct_and_army')!;
  const aggChains = dashVf.aggregations.find((a) => a.key === 'chains_sector')!;
  assert.equal(aggPrivate.sales, 52171054.36);
  assert.equal(aggPrivate.tax, 211626.87, 'Private markets subtotal must sum all 4 sectors tax');
  const sumSectorTax = Number(dashVf.sectors.reduce((acc, s) => acc + s.tax, 0).toFixed(2));
  assert.equal(sumSectorTax, 211626.87, 'Sum of per-sector tax must equal total tax');
  const expectedDirectArmyTax = Number(
    dashVf.sectors
      .filter((s) => ['التوريد المباشر', 'التوب ريتيل', 'الجيش'].includes(s.sectorName))
      .reduce((acc, s) => acc + s.tax, 0)
      .toFixed(2)
  );
  assert.equal(aggDirectArmy.tax, expectedDirectArmyTax, 'Direct & Army subtotal tax must match its 3 sectors');
  const expectedChainsTax = Number(
    dashVf.sectors
      .filter((s) => ['التوريد المباشر', 'التوريد المركزي'].includes(s.sectorName))
      .reduce((acc, s) => acc + s.tax, 0)
      .toFixed(2)
  );
  assert.equal(aggChains.tax, expectedChainsTax, 'Chains subtotal tax must match its 2 sectors');
  assert.ok(aggDirectArmy.sales < aggPrivate.sales);
  assert.ok(aggChains.sales < aggPrivate.sales);

  // 4. Verify Leading-Zero Code Preservation & Unlinked Branch Handling (15 records never guessed)
  const branchesRes = getBranchesAnalytics(db, adminUser, { source: 'VF05N' });
  assert.equal(branchesRes.totalUnlinkedCount, 15, '15 unlinked branch records preserved accurately');
  assert.ok(branchesRes.branches.some((b) => b.branchCode === '000501'), 'Leading zero on branch code 000501 preserved');

  // 5. Verify Negative Return Handling preserved
  const directSectorRow = dashVf.sectors.find((s) => s.sectorName === 'التوريد المباشر')!;
  assert.equal(directSectorRow.returnCount, 1, 'Return record counted');
  assert.equal(directSectorRow.returnValue, -50000, 'Negative return value preserved with original sign');

  // 6. Verify Backend Sector RBAC & Multi-Sector Aggregation (never restricted to allowedSectors[0])
  assert.equal(dashVf.sectors.length, 4, 'Admin with ALL filter must see all 4 original sectors');

  const directDash = getDashboardSummary(db, directUser, { source: 'VF05N' });
  assert.equal(directDash.sectors.length, 1, 'Single-sector user only sees authorized sector');
  assert.equal(directDash.sectors[0].sectorName, 'التوريد المباشر');
  assert.equal(directDash.kpis.privateMarketsTarget, 15000000, 'Sector user only sees target of authorized sector');

  const multiSectorUser: UserProfile = {
    id: 3,
    username: 'chains.mgr',
    fullName: 'مدير السلاسل',
    email: 'chains@test.com',
    role: 'sector_manager',
    allowedSectors: ['التوريد المباشر', 'التوريد المركزي'],
    isActive: true,
    createdAt: '2026-11-23T00:00:00Z',
  };
  const multiDash = getDashboardSummary(db, multiSectorUser, { source: 'VF05N', sector: 'ALL' });
  assert.equal(multiDash.sectors.length, 2, 'Multi-sector user must see ALL authorized sectors, not allowedSectors[0]');
  const expectedMultiSales = Number(
    (
      dashVf.sectors.find((s) => s.sectorName === 'التوريد المباشر')!.sales +
      dashVf.sectors.find((s) => s.sectorName === 'التوريد المركزي')!.sales
    ).toFixed(2)
  );
  assert.equal(multiDash.kpis.privateMarketsSales, expectedMultiSales, 'Must aggregate sales of ALL authorized sectors');
  assert.equal(multiDash.kpis.privateMarketsTarget, 40300000, 'Must aggregate targets of ALL authorized sectors');

  const multiFilteredDash = getDashboardSummary(db, multiSectorUser, { source: 'VF05N', sector: 'التوريد المركزي' });
  assert.equal(multiFilteredDash.sectors.length, 1, 'Manual sector filter narrows to selected authorized sector');
  assert.equal(multiFilteredDash.sectors[0].sectorName, 'التوريد المركزي');

  const unauthorizedFilteredDash = getDashboardSummary(db, multiSectorUser, { source: 'VF05N', sector: 'الجيش' });
  assert.equal(unauthorizedFilteredDash.sectors.length, 0, 'Unauthorized manual sector filter must return empty');

  const allKeywordUser: UserProfile = {
    ...multiSectorUser,
    id: 4,
    allowedSectors: ['الكل'],
  };
  const allKeywordDash = getDashboardSummary(db, allKeywordUser, { source: 'VF05N' });
  assert.equal(allKeywordDash.sectors.length, 4, 'User with الكل permission must see all 4 sectors');

  // 7. Verify SAP header aliases (New Total Tax, Total Tax, Billing Dt., Qty, Sales Employee Code, الصلاحيات)
  const sapWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(sapWb, XLSX.utils.aoa_to_sheet([['كود القطاع', 'القطاع'], ['S1', 'التوريد المباشر'], ['S2', 'التوريد المركزي'], ['S3', 'التوب ريتيل'], ['S4', 'الجيش']]), 'القطاعات');
  XLSX.utils.book_append_sheet(sapWb, XLSX.utils.aoa_to_sheet([['القطاع', 'الهدف'], ['التوريد المباشر', 1000], ['التوريد المركزي', 2000], ['التوب ريتيل', 3000], ['الجيش', 4000]]), 'اهداف القطاعات');
  XLSX.utils.book_append_sheet(sapWb, XLSX.utils.aoa_to_sheet([['كود السلسلة', 'اسم السلسلة', 'الادارة', 'القطاع'], ['101', 'عميل', 'إدارة', 'التوريد المباشر']]), 'تعريف الادارات والقطاعات');
  XLSX.utils.book_append_sheet(sapWb, XLSX.utils.aoa_to_sheet([['كود الفرع', 'اسم الفرع', 'القطاع', 'الهدف'], ['501', 'فرع', 'التوريد المباشر', 1000]]), 'تعريف الفروع و الاهداف');
  XLSX.utils.book_append_sheet(sapWb, XLSX.utils.aoa_to_sheet([['كود المندوب', 'اسم المندوب', 'القطاع'], ['R1', 'مندوب', 'التوريد المباشر']]), 'تعريف مندوبى البيع');
  XLSX.utils.book_append_sheet(sapWb, XLSX.utils.aoa_to_sheet([['كود الصنف', 'اسم الصنف'], ['M1', 'صنف']]), 'تعريف الاصناف');
  XLSX.utils.book_append_sheet(sapWb, XLSX.utils.aoa_to_sheet([['اسم المستخدم', 'الصلاحيات'], ['u1', 'التوريد المباشر']]), 'تعريف المستخدمين');
  XLSX.utils.book_append_sheet(sapWb, XLSX.utils.aoa_to_sheet([['Sold-to Party', 'Payer', 'Billing Date', 'Sales Employee Code', 'Material', 'New Total Tax', 'New Growth Value BD&T'], ['501', '101', '2026-11-15', 'R1', 'M1', 140, 1000]]), 'VF05N (2)');
  XLSX.utils.book_append_sheet(sapWb, XLSX.utils.aoa_to_sheet([['Billing Dt.', 'Customer', 'Material', 'Sales Emp.', 'Qty', 'Total Tax', 'New Growth Value BD&T'], [46341, '101', 'M1', 'R1', 5, 140, 1000]]), 'Zsales_Price (2)');
  const sapReport = validateExcelBuffer(XLSX.write(sapWb, { type: 'buffer', bookType: 'xlsx' }) as Buffer, 'SAP_Headers.xlsx');
  assert.equal(sapReport.canCommit, true, 'SAP headers workbook must validate cleanly');
  assert.equal(sapReport.comparison.vf05n.taxTotal, 140, 'New Total Tax must be parsed');
  assert.equal(sapReport.comparison.zsales.taxTotal, 140, 'Total Tax must be parsed');

  // Verify flexible Arabic/VAT tax column aliases with spaces, commas, and currency suffixes
  const taxAliasWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(taxAliasWb, XLSX.utils.aoa_to_sheet([['كود القطاع', 'القطاع'], ['S1', 'التوريد المباشر']]), 'القطاعات');
  XLSX.utils.book_append_sheet(taxAliasWb, XLSX.utils.aoa_to_sheet([['القطاع', 'الهدف'], ['التوريد المباشر', 10000]]), 'اهداف القطاعات');
  XLSX.utils.book_append_sheet(taxAliasWb, XLSX.utils.aoa_to_sheet([['كود السلسلة', 'اسم السلسلة', 'الادارة', 'القطاع'], ['101', 'عميل', 'إدارة', 'التوريد المباشر']]), 'تعريف الادارات والقطاعات');
  XLSX.utils.book_append_sheet(taxAliasWb, XLSX.utils.aoa_to_sheet([['كود الفرع', 'اسم الفرع', 'القطاع', 'الهدف'], ['501', 'فرع', 'التوريد المباشر', 10000]]), 'تعريف الفروع و الاهداف');
  XLSX.utils.book_append_sheet(taxAliasWb, XLSX.utils.aoa_to_sheet([['كود المندوب', 'اسم المندوب', 'القطاع'], ['R1', 'مندوب', 'التوريد المباشر']]), 'تعريف مندوبى البيع');
  XLSX.utils.book_append_sheet(taxAliasWb, XLSX.utils.aoa_to_sheet([['كود الصنف', 'اسم الصنف'], ['M1', 'صنف']]), 'تعريف الاصناف');
  XLSX.utils.book_append_sheet(taxAliasWb, XLSX.utils.aoa_to_sheet([['اسم المستخدم', 'الصلاحيات'], ['u1', 'التوريد المباشر']]), 'تعريف المستخدمين');
  XLSX.utils.book_append_sheet(taxAliasWb, XLSX.utils.aoa_to_sheet([['Sold-to Party', 'Payer', 'Billing Date', 'Material', ' قيمة الضريبة ', 'New Growth Value BD&T'], ['501', '101', '2026-11-15', 'M1', '1,250.50 EGP', '10,000.00']]), 'VF05N (2)');
  XLSX.utils.book_append_sheet(taxAliasWb, XLSX.utils.aoa_to_sheet([['Billing Dt.', 'Customer', 'Material', 'Sales Emp.', ' VAT ', 'New Growth Value BD&T'], [46341, '101', 'M1', 'R1', ' 1,250.50 ج.م ', 10000]]), 'Zsales_Price (2)');
  const taxAliasReport = validateExcelBuffer(XLSX.write(taxAliasWb, { type: 'buffer', bookType: 'xlsx' }) as Buffer, 'Tax_Aliases.xlsx');
  assert.equal(taxAliasReport.canCommit, true, 'Flexible Arabic/VAT tax headers must validate');
  assert.equal(taxAliasReport.comparison.vf05n.taxTotal, 1250.5, 'Arabic قيمة الضريبة with currency/commas must parse');
  assert.equal(taxAliasReport.comparison.zsales.taxTotal, 1250.5, 'VAT header with Arabic currency/spaces must parse');

  // 8. Verify Snapshot Versioning & Rollback
  const report2 = validateExcelBuffer(xlsxBuf, 'Batch_2.xlsx');
  const batch2 = commitValidatedImport(db, report2.validationToken, { id: 1, fullName: 'مدير النظام' });
  assert.equal(batch2.id, 2);
  const restored = rollbackToBatch(db, batch1.id);
  assert.equal(restored.id, 1);
  assert.equal(restored.status, 'active');

  // 9. Verify Hierarchical Supervisor-Grouped Chains Breakdown, التصنيف A (تجاري/صناعي), Negative Share %, & VF05N Fallback Split
  const supWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    supWb,
    XLSX.utils.aoa_to_sheet([
      ['كود القطاع', 'القطاع'],
      ['S1', 'التوريد المركزي'],
    ]),
    'القطاعات'
  );
  XLSX.utils.book_append_sheet(
    supWb,
    XLSX.utils.aoa_to_sheet([
      ['القطاع', 'الهدف'],
      ['التوريد المركزي', 2000000],
    ]),
    'اهداف القطاعات'
  );
  XLSX.utils.book_append_sheet(
    supWb,
    XLSX.utils.aoa_to_sheet([
      ['كود السلسلة', 'السلسلة', 'القطاع', 'اسم المشرف'],
      ['C101', 'سلسلة ألفا', 'التوريد المركزي', 'محمد عبدالمحسن'],
      ['C102', 'سلسلة بيتا (مرتجعات)', 'التوريد المركزي', 'محمد عبدالمحسن'],
      ['C201', 'سلسلة جاما', 'التوريد المركزي', 'محمد مجدى'],
    ]),
    'تعريف الادارات والقطاعات'
  );
  XLSX.utils.book_append_sheet(
    supWb,
    XLSX.utils.aoa_to_sheet([
      ['كود الفرع', 'اسم الفرع', 'القطاع', 'الهدف'],
      ['B1', 'فرع رئيسي', 'التوريد المركزي', 2000000],
    ]),
    'تعريف الفروع و الاهداف'
  );
  XLSX.utils.book_append_sheet(
    supWb,
    XLSX.utils.aoa_to_sheet([
      ['كود المندوب', 'اسم المندوب', 'اسم القطاع'],
      ['R1', 'مندوب 1', 'التوريد المركزي'],
    ]),
    'تعريف مندوبى البيع'
  );
  XLSX.utils.book_append_sheet(
    supWb,
    XLSX.utils.aoa_to_sheet([
      ['كود الصنف', 'الصنف', 'التصنيف A', 'التصنيف B'],
      ['IT_COM', 'صنف تجاري 1', 'تجارى', 'بقالة'],
      ['IT_IND', 'صنف صناعي 1', 'صناعى', 'شركات'],
    ]),
    'تعريف الاصناف'
  );
  XLSX.utils.book_append_sheet(
    supWb,
    XLSX.utils.aoa_to_sheet([
      ['اسم المستخدم', 'الصلاحيات'],
      ['u1', 'الكل'],
    ]),
    'تعريف المستخدمين'
  );

  // VF05N (2) intentionally WITHOUT Material column (matching real SAP VF05N export)
  XLSX.utils.book_append_sheet(
    supWb,
    XLSX.utils.aoa_to_sheet([
      ['Sold-to Party', 'Payer', 'Billing Date', 'Sales Employee Code', 'New Total Tax', 'New Growth Value BD&T'],
      ['B1', 'C101', '2026-11-16', 'R1', 14000, 1000000],
      ['B1', 'C102', '2026-11-16', 'R1', -180, -12830],
      ['B1', 'C201', '2026-11-16', 'R1', 7000, 500000],
    ]),
    'VF05N (2)'
  );

  // Zsales_Price (2) WITH Material column splitting C101 (600k تجارى / 400k صناعى), C102 (-12,830 تجارى), C201 (500k صناعى)
  XLSX.utils.book_append_sheet(
    supWb,
    XLSX.utils.aoa_to_sheet([
      ['Billing Dt.', 'Customer', 'Material', 'Sales Emp.', 'Qty', 'Total Tax', 'New Growth Value BD&T'],
      [46342, 'C101', 'IT_COM', 'R1', 10, 8400, 600000],
      [46342, 'C101', 'IT_IND', 'R1', 5, 5600, 400000],
      [46342, 'C102', 'IT_COM', 'R1', -1, -180, -12830],
      [46342, 'C201', 'IT_IND', 'R1', 8, 7000, 500000],
    ]),
    'Zsales_Price (2)'
  );

  const supReport = validateExcelBuffer(XLSX.write(supWb, { type: 'buffer', bookType: 'xlsx' }) as Buffer, 'Supervisor_Split.xlsx');
  assert.equal(supReport.canCommit, true, 'Supervisor commercial/industrial workbook must validate');
  commitValidatedImport(db, supReport.validationToken, { id: 1, fullName: 'مدير النظام' });

  for (const src of ['ZSALES', 'VF05N'] as const) {
    const chainsOut = getDepartmentsChainsBreakdown(db, adminUser, { source: src });
    assert.equal(chainsOut.length, 3, `[${src}] Expected 3 chain rows`);

    const c101 = chainsOut.find((c) => c.chainCode === 'C101')!;
    const c102 = chainsOut.find((c) => c.chainCode === 'C102')!;
    const c201 = chainsOut.find((c) => c.chainCode === 'C201')!;

    assert.equal(c101.supervisorName, 'محمد عبدالمحسن');
    assert.equal(c101.commercialSales, 600000, `[${src}] C101 commercialSales mismatch`);
    assert.equal(c101.industrialSales, 400000, `[${src}] C101 industrialSales mismatch`);
    assert.equal(c101.sales, 1000000);
    assert.equal(c101.tax, 14000);

    // Supervisor total = 1,000,000 - 12,830 = 987,170 -> C102 share = -12,830 / 987,170 * 100 = -1.3%
    assert.equal(c102.supervisorName, 'محمد عبدالمحسن');
    assert.equal(c102.commercialSales, -12830);
    assert.equal(c102.industrialSales, 0);
    assert.equal(c102.shareOfSupervisorPct, -1.3, `[${src}] Negative return shareOfSupervisorPct must be -1.3`);
    assert.equal(c101.shareOfSupervisorPct, 101.3, `[${src}] Positive shareOfSupervisorPct must be 101.3`);

    assert.equal(c201.supervisorName, 'محمد مجدى');
    assert.equal(c201.commercialSales, 0);
    assert.equal(c201.industrialSales, 500000);
    assert.equal(c201.shareOfSupervisorPct, 100);
  }

  // 10. Verify "تعريف الإدارات والقطاعات" (with Hamza), Leading-Zero Payer matching, Chain Name fallback, and "غير محدد" only for unmatched rows
  const hamzaWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(hamzaWb, XLSX.utils.aoa_to_sheet([['القطاع'], ['التوب ريتيل']]), 'القطاعات');
  XLSX.utils.book_append_sheet(hamzaWb, XLSX.utils.aoa_to_sheet([['القطاع', 'الهدف'], ['التوب ريتيل', 1000000]]), 'اهداف القطاعات');
  XLSX.utils.book_append_sheet(
    hamzaWb,
    XLSX.utils.aoa_to_sheet([
      ['كود السلسلة', 'السلسلة', 'القطاع', 'اسم المشرف'],
      ['0064000030', 'سلسلة أصفار بادئة', 'التوب ريتيل', 'حازم السيد'],
      ['0064000099', 'أولاد رجب المعادي', 'التوب ريتيل', 'مصطفى مناع'],
      ['0064000100', 'سلسلة بدون مبيعات', 'التوب ريتيل', 'مصطفى مناع'],
      ['0064000200', 'سلسلة مشرف صفري', 'التوب ريتيل', 'مشرف بدون مبيعات'],
    ]),
    'تعريف الإدارات والقطاعات'
  );
  XLSX.utils.book_append_sheet(hamzaWb, XLSX.utils.aoa_to_sheet([['كود الفرع', 'اسم الفرع', 'القطاع', 'الهدف'], ['B1', 'فرع', 'التوب ريتيل', 1000000]]), 'تعريف الفروع و الاهداف');
  XLSX.utils.book_append_sheet(hamzaWb, XLSX.utils.aoa_to_sheet([['كود المندوب', 'اسم المندوب', 'القطاع'], ['R1', 'مندوب', 'التوب ريتيل']]), 'تعريف مندوبى البيع');
  XLSX.utils.book_append_sheet(hamzaWb, XLSX.utils.aoa_to_sheet([['كود الصنف', 'اسم الصنف', 'التصنيف A'], ['M1', 'صنف', 'تجارى']]), 'تعريف الاصناف');
  XLSX.utils.book_append_sheet(hamzaWb, XLSX.utils.aoa_to_sheet([['اسم المستخدم', 'الصلاحيات'], ['u1', 'الكل']]), 'تعريف المستخدمين');
  XLSX.utils.book_append_sheet(
    hamzaWb,
    XLSX.utils.aoa_to_sheet([
      ['Sold-to Party', 'Payer', 'Name of Payer', 'Billing Date', 'New Total Tax', 'New Growth Value BD&T'],
      // Numeric Payer missing leading zeros vs '0064000030' in master sheet
      ['B1', 64000030, 'عميل مختلف الاسم', '2026-11-17', 1400, 100000],
      // Empty Payer code but matching chain name ('اولاد رجب المعادى' normalized hamza/yeh)
      ['B1', '', 'اولاد رجب المعادى', '2026-11-17', 700, 50000],
      // Unmatched chain code not present in تعريف الإدارات والقطاعات
      ['B1', '99999999', 'عميل غير معرف', '2026-11-17', 0, 25000],
    ]),
    'VF05N (2)'
  );
  XLSX.utils.book_append_sheet(
    hamzaWb,
    XLSX.utils.aoa_to_sheet([
      ['Billing Dt.', 'Customer', 'Customer Name', 'Material', 'Sales Emp.', 'Qty', 'Total Tax', 'New Growth Value BD&T'],
      [46343, 64000030, 'عميل مختلف الاسم', 'M1', 'R1', 10, 1400, 100000],
      [46343, '', 'اولاد رجب المعادى', 'M1', 'R1', 5, 700, 50000],
      [46343, '99999999', 'عميل غير معرف', 'M1', 'R1', 2, 0, 25000],
    ]),
    'Zsales_Price (2)'
  );

  const hamzaReport = validateExcelBuffer(XLSX.write(hamzaWb, { type: 'buffer', bookType: 'xlsx' }) as Buffer, 'Hamza_LeadingZeros.xlsx');
  assert.equal(hamzaReport.canCommit, true, 'Workbook with تعريف الإدارات والقطاعات and leading zero tolerance must validate');
  commitValidatedImport(db, hamzaReport.validationToken, { id: 1, fullName: 'مدير النظام' });

  const hamzaChains = getDepartmentsChainsBreakdown(db, adminUser, { source: 'VF05N' });
  const lzChain = hamzaChains.find((c) => c.chainCode === '0064000030')!;
  assert.ok(lzChain, 'Chain 0064000030 must resolve via leading-zero tolerance');
  assert.equal(lzChain.supervisorName, 'حازم السيد');
  assert.equal(lzChain.commercialSales, 100000);

  const nameChain = hamzaChains.find((c) => c.chainCode === '0064000099')!;
  assert.ok(nameChain, 'Chain 0064000099 must resolve via normalized chain name السلسلة');
  assert.equal(nameChain.supervisorName, 'مصطفى مناع');
  assert.equal(nameChain.commercialSales, 50000);

  const zeroChainSameSup = hamzaChains.find((c) => c.chainCode === '0064000100')!;
  assert.ok(zeroChainSameSup, 'Zero-sales master chain 0064000100 must remain visible');
  assert.equal(zeroChainSameSup.supervisorName, 'مصطفى مناع');
  assert.equal(zeroChainSameSup.sales, 0);
  assert.equal(zeroChainSameSup.tax, 0);
  assert.equal(zeroChainSameSup.shareOfSupervisorPct, 0);

  const zeroSupChain = hamzaChains.find((c) => c.chainCode === '0064000200')!;
  assert.ok(zeroSupChain, 'Zero-sales supervisor master chain 0064000200 must remain visible');
  assert.equal(zeroSupChain.supervisorName, 'مشرف بدون مبيعات');
  assert.equal(zeroSupChain.sales, 0);
  assert.equal(zeroSupChain.shareOfSupervisorPct, 0);

  // Even in an empty date window with zero sales records, all 4 master chains must still appear with 0s
  const emptyWindowChains = getDepartmentsChainsBreakdown(db, adminUser, {
    source: 'VF05N',
    startDate: '2099-01-01',
    endDate: '2099-01-31',
  });
  assert.equal(emptyWindowChains.length, 4, 'All 4 master chains must still appear when filtered date range has 0 sales');
  assert.ok(emptyWindowChains.every((c) => c.sales === 0 && c.tax === 0 && c.shareOfSupervisorPct === 0));

  const unmatchedChain = hamzaChains.find((c) => c.chainCode === '99999999')!;
  assert.ok(unmatchedChain, 'Unmatched chain 99999999 present');
  assert.equal(unmatchedChain.supervisorName, 'غير محدد', 'Unmatched chain must show غير محدد');

  // 11. Verify Multi-Code Chain Consolidation by Normalized Chain Name within Supervisor (e.g. ال اس جى across 64001238 & 64001380)
  const multiCodeWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(multiCodeWb, XLSX.utils.aoa_to_sheet([['القطاع'], ['التوريد المباشر']]), 'القطاعات');
  XLSX.utils.book_append_sheet(multiCodeWb, XLSX.utils.aoa_to_sheet([['القطاع', 'الهدف'], ['التوريد المباشر', 2000000]]), 'اهداف القطاعات');
  XLSX.utils.book_append_sheet(
    multiCodeWb,
    XLSX.utils.aoa_to_sheet([
      ['كود السلسلة', 'السلسلة', 'القطاع', 'اسم المشرف'],
      ['64001238', 'ال اس جى', 'التوريد المباشر', 'محمد مجدى'],
      ['64001380', 'ال اس جي ', 'التوريد المباشر', 'محمد مجدي'],
      ['64009999', 'سلسلة اخرى', 'التوريد المباشر', 'محمد مجدى'],
      ['64008888', 'سلسلة صفرية موحدة', 'التوريد المباشر', 'محمد مجدى'],
      ['64008889', 'سلسلة  صفرية  موحدة', 'التوريد المباشر', 'محمد مجدى'],
    ]),
    'تعريف الإدارات والقطاعات'
  );
  XLSX.utils.book_append_sheet(multiCodeWb, XLSX.utils.aoa_to_sheet([['كود الفرع', 'اسم الفرع', 'القطاع', 'الهدف'], ['B1', 'فرع', 'التوريد المباشر', 2000000]]), 'تعريف الفروع و الاهداف');
  XLSX.utils.book_append_sheet(multiCodeWb, XLSX.utils.aoa_to_sheet([['كود المندوب', 'اسم المندوب', 'القطاع'], ['R1', 'مندوب', 'التوريد المباشر']]), 'تعريف مندوبى البيع');
  XLSX.utils.book_append_sheet(
    multiCodeWb,
    XLSX.utils.aoa_to_sheet([
      ['كود الصنف', 'اسم الصنف', 'التصنيف A'],
      ['M_COM', 'صنف تجاري', 'تجارى'],
      ['M_IND', 'صنف صناعي', 'صناعى'],
    ]),
    'تعريف الاصناف'
  );
  XLSX.utils.book_append_sheet(multiCodeWb, XLSX.utils.aoa_to_sheet([['اسم المستخدم', 'الصلاحيات'], ['u1', 'الكل']]), 'تعريف المستخدمين');
  XLSX.utils.book_append_sheet(
    multiCodeWb,
    XLSX.utils.aoa_to_sheet([
      ['Sold-to Party', 'Payer', 'Billing Date', 'New Total Tax', 'New Growth Value BD&T'],
      ['B1', '64001238', '2026-11-18', 1000, 300000],
      ['B1', '64001380', '2026-11-18', 500, 200000],
      ['B1', '64009999', '2026-11-18', 2500, 500000],
    ]),
    'VF05N (2)'
  );
  XLSX.utils.book_append_sheet(
    multiCodeWb,
    XLSX.utils.aoa_to_sheet([
      ['Billing Dt.', 'Customer', 'Material', 'Sales Emp.', 'Qty', 'Total Tax', 'New Growth Value BD&T'],
      [46344, '64001238', 'M_COM', 'R1', 10, 1000, 300000],
      [46344, '64001380', 'M_IND', 'R1', 5, 500, 200000],
      [46344, '64009999', 'M_COM', 'R1', 20, 2500, 500000],
    ]),
    'Zsales_Price (2)'
  );

  const multiCodeReport = validateExcelBuffer(XLSX.write(multiCodeWb, { type: 'buffer', bookType: 'xlsx' }) as Buffer, 'MultiCode_Unified.xlsx');
  assert.equal(multiCodeReport.canCommit, true);
  commitValidatedImport(db, multiCodeReport.validationToken, { id: 1, fullName: 'مدير النظام' });

  for (const src of ['ZSALES', 'VF05N'] as const) {
    const unified = getDepartmentsChainsBreakdown(db, adminUser, { source: src });
    // 5 master rows consolidate into 3 unified rows: ال اس جى (1), سلسلة اخرى (1), سلسلة صفرية موحدة (1)
    assert.equal(unified.length, 3, `[${src}] Expected 3 unified chain rows under محمد مجدى`);
    const lsg = unified.find((c) => c.chainName.includes('اس ج'))!;
    assert.ok(lsg, `[${src}] Unified ال اس جى row must exist`);
    assert.equal(lsg.commercialSales, 300000, `[${src}] ال اس جى commercialSales must sum code 64001238`);
    assert.equal(lsg.industrialSales, 200000, `[${src}] ال اس جى industrialSales must sum code 64001380`);
    assert.equal(lsg.sales, 500000, `[${src}] ال اس جى total sales must merge both codes`);
    assert.equal(lsg.tax, 1500, `[${src}] ال اس جى tax must sum both codes`);
    assert.equal(lsg.shareOfSupervisorPct, 50, `[${src}] 500k / 1M supervisor total = 50.0%`);
    assert.ok(lsg.chainCode.includes('64001238') && lsg.chainCode.includes('64001380'));

    const zeroUnified = unified.find((c) => c.chainCode.includes('64008888'))!;
    assert.ok(zeroUnified, `[${src}] Consolidated zero-sales chain must remain visible`);
    assert.equal(zeroUnified.sales, 0);
    assert.equal(zeroUnified.shareOfSupervisorPct, 0);
  }

  // 12. Verify Branches Analytics with كود الادارة / اسم الادارة / اسم المشرف / إجمالى الهدف, repBreakdowns, target === 0 -> null, and >= 70.5% boundary
  const brWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(brWb, XLSX.utils.aoa_to_sheet([['القطاع'], ['التوريد المباشر']]), 'القطاعات');
  XLSX.utils.book_append_sheet(brWb, XLSX.utils.aoa_to_sheet([['القطاع', 'الهدف'], ['التوريد المباشر', 500000]]), 'اهداف القطاعات');
  XLSX.utils.book_append_sheet(
    brWb,
    XLSX.utils.aoa_to_sheet([
      ['كود السلسلة', 'السلسلة', 'القطاع', 'اسم المشرف'],
      ['64000059', 'البنيان', 'التوريد المباشر', 'محمد عبدالمحسن'],
    ]),
    'تعريف الإدارات والقطاعات'
  );
  XLSX.utils.book_append_sheet(
    brWb,
    XLSX.utils.aoa_to_sheet([
      ['القطاع', 'كود الفرع', 'اسم الفرع', 'كود الادارة', 'اسم الادارة', 'كود المندوب', 'اسم المندوب', 'اسم المشرف', 'هدف التجارى', 'هدف الصناعى', 'إجمالى الهدف'],
      ['التوريد المباشر', '64000209', 'البنيان الثورة', '64000059', 'البنيان', '102342', 'عبدالله الملا - عقود', 'محمد عبدالمحسن', 100000, 30000, 130000],
      ['التوريد المباشر', '64000264', 'البنيان زهراء عين شمس', '64000059', 'البنيان', '102342', 'عبدالله الملا - عقود', 'محمد عبدالمحسن', 40000, 15000, 55000],
      ['التوريد المباشر', '64000164', 'البنيان المطرية', '64000059', 'البنيان', '102342', 'عبدالله الملا - عقود', 'محمد عبدالمحسن', 23000, 10000, 33000],
      ['التوريد المباشر', '64000170', 'البنيان الزيتون مركزى', '64000059', 'البنيان', '102342', 'عبدالله الملا - عقود', 'محمد عبدالمحسن', 0, 0, 0],
      ['التوريد المباشر', '64000999', 'البنيان حد 70.4', '64000059', 'البنيان', '101179', 'مندوب اخر', 'محمد عبدالمحسن', 100000, 0, 100000],
    ]),
    'تعريف الفروع و الاهداف'
  );
  XLSX.utils.book_append_sheet(
    brWb,
    XLSX.utils.aoa_to_sheet([
      ['كود المندوب', 'اسم المندوب', 'القطاع'],
      ['102342', 'عبدالله الملا - عقود', 'التوريد المباشر'],
      ['101222', 'مندوب مساعد', 'التوريد المباشر'],
      ['101179', 'مندوب اخر', 'التوريد المباشر'],
    ]),
    'تعريف مندوبى البيع'
  );
  XLSX.utils.book_append_sheet(brWb, XLSX.utils.aoa_to_sheet([['كود الصنف', 'اسم الصنف', 'التصنيف A'], ['M1', 'صنف', 'تجارى']]), 'تعريف الاصناف');
  XLSX.utils.book_append_sheet(brWb, XLSX.utils.aoa_to_sheet([['اسم المستخدم', 'الصلاحيات'], ['u1', 'الكل']]), 'تعريف المستخدمين');
  XLSX.utils.book_append_sheet(
    brWb,
    XLSX.utils.aoa_to_sheet([
      ['Sold-to Party', 'Payer', 'Billing Date', 'Sales Employee Code', 'New Total Tax', 'New Growth Value BD&T'],
      ['64000209', '64000059', '2026-11-18', '102342', 337.99, 91549.89],
      ['64000264', '64000059', '2026-11-18', '102342', 0, 9156],
      ['64000264', '64000059', '2026-11-18', '101222', 0, 594],
      ['64000999', '64000059', '2026-11-18', '101179', 0, 70400],
    ]),
    'VF05N (2)'
  );
  XLSX.utils.book_append_sheet(
    brWb,
    XLSX.utils.aoa_to_sheet([
      ['Billing Dt.', 'Customer', 'Material', 'Sales Emp.', 'Qty', 'Total Tax', 'New Growth Value BD&T'],
      [46344, '64000059', 'M1', '102342', 10, 337.99, 171699.89],
    ]),
    'Zsales_Price (2)'
  );

  const brReport = validateExcelBuffer(XLSX.write(brWb, { type: 'buffer', bookType: 'xlsx' }) as Buffer, 'Branches_Analysis.xlsx');
  assert.equal(brReport.canCommit, true, 'Branches analysis workbook must validate cleanly');
  commitValidatedImport(db, brReport.validationToken, { id: 1, fullName: 'مدير النظام' });

  const brOut = getBranchesAnalytics(db, adminUser, { source: 'VF05N' }).branches;
  assert.equal(brOut.length, 5, 'All 5 defined branches including zero-sales and zero-target must be present');

  const thawra = brOut.find((b) => b.branchCode === '64000209')!;
  assert.equal(thawra.chainCode, '64000059');
  assert.equal(thawra.chainName, 'البنيان');
  assert.equal(thawra.supervisorName, 'محمد عبدالمحسن');
  assert.equal(thawra.assignedRepCode, '102342');
  assert.equal(thawra.assignedRepName, 'عبدالله الملا - عقود');
  assert.equal(thawra.target, 130000);
  assert.equal(thawra.achievementPct, 70.61);
  assert.ok(thawra.achievementPct! >= 70.5, '70.61% must pass >= 70.5% efficiency threshold');

  const zahraa = brOut.find((b) => b.branchCode === '64000264')!;
  assert.equal(zahraa.sales, 9750, 'Multi-rep supplies must sum to branch total 9,750');
  assert.equal(zahraa.repBreakdowns.length, 2, 'Both rep 102342 and 101222 recorded in repBreakdowns');

  const matariya = brOut.find((b) => b.branchCode === '64000164')!;
  assert.equal(matariya.sales, 0);
  assert.equal(matariya.target, 33000);
  assert.equal(matariya.achievementPct, 0, 'Zero sales with positive target -> 0%');

  const zaytounCentral = brOut.find((b) => b.branchCode === '64000170')!;
  assert.equal(zaytounCentral.target, 0);
  assert.equal(zaytounCentral.achievementPct, null, 'Target === 0 must produce null achievementPct for blank cell rendering');

  const edge704 = brOut.find((b) => b.branchCode === '64000999')!;
  assert.equal(edge704.achievementPct, 70.4);
  assert.ok(edge704.achievementPct! < 70.5, '70.40% must NOT pass >= 70.5% efficiency threshold');

  // 13. Verify strict branch-row Supervisor truth (no override from chain or rep), leading-zero Sold-to Party matching, & sector join by branch_code
  const supIsoWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    supIsoWb,
    XLSX.utils.aoa_to_sheet([
      ['القطاع'],
      ['التوريد المباشر'],
      ['الجيش'],
      ['التوب ريتيل'],
    ]),
    'القطاعات'
  );
  XLSX.utils.book_append_sheet(
    supIsoWb,
    XLSX.utils.aoa_to_sheet([
      ['القطاع', 'الهدف'],
      ['التوريد المباشر', 500000],
      ['الجيش', 100000],
      ['التوب ريتيل', 100000],
    ]),
    'اهداف القطاعات'
  );
  XLSX.utils.book_append_sheet(
    supIsoWb,
    XLSX.utils.aoa_to_sheet([
      ['كود السلسلة', 'السلسلة', 'القطاع', 'اسم المشرف'],
      ['0064000021', 'سعودى ماركت', 'التوريد المباشر', 'محمد عبدالمحسن'],
      ['0064000888', 'إدارة توب ريتيل', 'التوب ريتيل', 'مشرف التوب ريتيل'],
    ]),
    'تعريف الإدارات والقطاعات'
  );
  XLSX.utils.book_append_sheet(
    supIsoWb,
    XLSX.utils.aoa_to_sheet([
      ['القطاع', 'كود الفرع', 'اسم الفرع', 'كود الادارة', 'اسم الادارة', 'كود المندوب', 'اسم المندوب', 'اسم المشرف', 'هدف التجارى', 'هدف الصناعى', 'إجمالى الهدف'],
      // Same chain + same rep, different supervisors on each branch row
      ['التوريد المباشر', '0064000081', 'سعودى دستريكت', '0064000021', 'سعودى ماركت', '102342', 'عبدالله الملا', 'محمد عبدالمحسن', 50000, 0, 50000],
      ['التوريد المباشر', '0064000082', 'سعودى سموحة', '0064000021', 'سعودى ماركت', '102342', 'عبدالله الملا', 'أحمد سمير', 60000, 0, 60000],
      // Empty supervisor cell on branch row -> MUST remain 'غير محدد' (never inferred from chain/rep)
      ['التوريد المباشر', '0064000083', 'سعودى غير مسند لمشرف', '0064000021', 'سعودى ماركت', '102342', 'عبدالله الملا', '', 40000, 0, 40000],
      // Branch sector = الجيش while VF05N Payer = التوب ريتيل -> branch sales must still appear under الجيش
      ['الجيش', '0064000666', 'نادي الصيد', '0064000888', 'نادي الصيد إدارة', '101071', 'مندوب الجيش', 'كامل حسن', 100000, 0, 100000],
    ]),
    'تعريف الفروع و الاهداف'
  );
  XLSX.utils.book_append_sheet(
    supIsoWb,
    XLSX.utils.aoa_to_sheet([
      ['كود المندوب', 'اسم المندوب', 'القطاع', 'اسم المشرف'],
      ['102342', 'عبدالله الملا', 'التوريد المباشر', 'خالد علي'],
      ['101071', 'مندوب الجيش', 'الجيش', 'كامل حسن'],
    ]),
    'تعريف مندوبى البيع'
  );
  XLSX.utils.book_append_sheet(supIsoWb, XLSX.utils.aoa_to_sheet([['كود الصنف', 'اسم الصنف', 'التصنيف A'], ['M1', 'صنف', 'تجارى']]), 'تعريف الاصناف');
  XLSX.utils.book_append_sheet(supIsoWb, XLSX.utils.aoa_to_sheet([['اسم المستخدم', 'الصلاحيات'], ['u1', 'الكل']]), 'تعريف المستخدمين');
  XLSX.utils.book_append_sheet(
    supIsoWb,
    XLSX.utils.aoa_to_sheet([
      ['Sold-to Party', 'Payer', 'Billing Date', 'Sales Employee Code', 'New Total Tax', 'New Growth Value BD&T'],
      // Numeric Sold-to missing leading zeros vs '0064000082'
      [64000082, '0064000021', '2026-11-19', '102342', 140, 45000],
      ['0064000666', '0064000888', '2026-11-19', '101071', 0, 8958.08],
    ]),
    'VF05N (2)'
  );
  XLSX.utils.book_append_sheet(
    supIsoWb,
    XLSX.utils.aoa_to_sheet([
      ['Billing Dt.', 'Customer', 'Material', 'Sales Emp.', 'Qty', 'Total Tax', 'New Growth Value BD&T'],
      [46345, '0064000021', 'M1', '102342', 10, 140, 53958.08],
    ]),
    'Zsales_Price (2)'
  );

  const supIsoReport = validateExcelBuffer(XLSX.write(supIsoWb, { type: 'buffer', bookType: 'xlsx' }) as Buffer, 'Branch_Supervisor_Truth.xlsx');
  assert.equal(supIsoReport.canCommit, true);
  commitValidatedImport(db, supIsoReport.validationToken, { id: 1, fullName: 'مدير النظام' });

  const supIsoBranches = getBranchesAnalytics(db, adminUser, { source: 'VF05N' }).branches;
  const brA = supIsoBranches.find((b) => b.branchCode === '0064000081')!;
  const brB = supIsoBranches.find((b) => b.branchCode === '0064000082')!;
  const brC = supIsoBranches.find((b) => b.branchCode === '0064000083')!;
  assert.equal(brA.supervisorName, 'محمد عبدالمحسن');
  assert.equal(brB.supervisorName, 'أحمد سمير', 'Branch supervisor must come from تعريف الفروع والأهداف, never chain or rep');
  assert.equal(brB.sales, 45000, 'Numeric Sold-to 64000082 must link to leading-zero branch 0064000082');
  assert.equal(brB.tax, 140);
  assert.equal(brC.supervisorName, 'غير محدد', 'Empty supervisor cell must NOT be inferred from chain or rep');

  const armyOnlyBranches = getBranchesAnalytics(db, adminUser, { source: 'VF05N', sector: 'الجيش' }).branches;
  assert.equal(armyOnlyBranches.length, 1);
  assert.equal(armyOnlyBranches[0].branchCode, '0064000666');
  assert.equal(armyOnlyBranches[0].sales, 8958.08, 'Branch sales must join on branch_code even when Payer sector differs');

  // 14. Verify Sales Reps Target aggregation & Achievement % = Sales / (Target - Tax)
  const repsWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    repsWb,
    XLSX.utils.aoa_to_sheet([['القطاع'], ['التوريد المباشر'], ['الجيش']]),
    'القطاعات'
  );
  XLSX.utils.book_append_sheet(
    repsWb,
    XLSX.utils.aoa_to_sheet([
      ['القطاع', 'الهدف'],
      ['التوريد المباشر', 200000],
      ['الجيش', 100000],
    ]),
    'اهداف القطاعات'
  );
  XLSX.utils.book_append_sheet(
    repsWb,
    XLSX.utils.aoa_to_sheet([
      ['كود السلسلة', 'السلسلة', 'القطاع', 'اسم المشرف'],
      ['0064000021', 'سعودى ماركت', 'التوريد المباشر', 'محمد عبدالمحسن'],
      ['0064000777', 'منافذ الجيش', 'الجيش', 'كامل حسن'],
    ]),
    'تعريف الإدارات والقطاعات'
  );
  XLSX.utils.book_append_sheet(
    repsWb,
    XLSX.utils.aoa_to_sheet([
      ['القطاع', 'كود الفرع', 'اسم الفرع', 'كود الادارة', 'اسم الادارة', 'كود المندوب', 'اسم المندوب', 'اسم المشرف', 'هدف التجارى', 'هدف الصناعى', 'إجمالى الهدف'],
      ['التوريد المباشر', '0064000081', 'سعودى 1', '0064000021', 'سعودى ماركت', '102342', 'عبدالله الملا', 'محمد عبدالمحسن', 80000, 20000, 100000],
      ['التوريد المباشر', '0064000082', 'سعودى 2', '0064000021', 'سعودى ماركت', '102342', 'عبدالله الملا', 'محمد عبدالمحسن', 50000, 0, 50000],
      // Zero-target and empty-target branches assigned to 102342 MUST NOT be counted in الفروع
      ['التوريد المباشر', '0064000084', 'سعودى بلا هدف 1', '0064000021', 'سعودى ماركت', '102342', 'عبدالله الملا', 'محمد عبدالمحسن', 0, 0, 0],
      ['التوريد المباشر', '0064000085', 'سعودى بلا هدف 2', '0064000021', 'سعودى ماركت', '102342', 'عبدالله الملا', 'محمد عبدالمحسن', '', '', ''],
      ['الجيش', '0064000771', 'منفذ الجيش 1', '0064000777', 'منافذ الجيش', '101071', 'مندوب الجيش', 'كامل حسن', 100000, 0, 100000],
      // Branch whose target is less than tax (tests negative/zero denominator safety)
      ['التوريد المباشر', '0064000099', 'فرع ضريبة أعلى', '0064000021', 'سعودى ماركت', '109999', 'مندوب عجز', 'محمد عبدالمحسن', 500, 0, 500],
    ]),
    'تعريف الفروع و الاهداف'
  );
  XLSX.utils.book_append_sheet(
    repsWb,
    XLSX.utils.aoa_to_sheet([
      ['كود المندوب', 'اسم المندوب', 'القطاع', 'اسم المشرف', 'الهدف'],
      ['102342', 'عبدالله الملا', 'التوريد المباشر', 'محمد عبدالمحسن', 0],
      ['101071', 'مندوب الجيش', 'الجيش', 'كامل حسن', 0],
      ['108888', 'مندوب بدون فروع', 'التوريد المباشر', 'محمد عبدالمحسن', 75000],
      ['109999', 'مندوب عجز', 'التوريد المباشر', 'محمد عبدالمحسن', 0],
      ['107777', 'مندوب بلا هدف', 'التوريد المباشر', 'محمد عبدالمحسن', 0],
    ]),
    'تعريف مندوبى البيع'
  );
  XLSX.utils.book_append_sheet(repsWb, XLSX.utils.aoa_to_sheet([['كود الصنف', 'اسم الصنف', 'التصنيف A'], ['M1', 'صنف', 'تجارى']]), 'تعريف الاصناف');
  XLSX.utils.book_append_sheet(repsWb, XLSX.utils.aoa_to_sheet([['اسم المستخدم', 'الصلاحيات'], ['u1', 'الكل']]), 'تعريف المستخدمين');
  XLSX.utils.book_append_sheet(
    repsWb,
    XLSX.utils.aoa_to_sheet([
      ['Sold-to Party', 'Payer', 'Billing Date', 'Sales Employee Code', 'New Total Tax', 'New Growth Value BD&T'],
      ['0064000081', '0064000021', '2026-11-19', '102342', 140, 45000],
      ['0064000771', '0064000777', '2026-11-19', '101071', 250, 75000],
      ['0064000099', '0064000021', '2026-11-19', '109999', 600, 1200],
      // Rep 107777 invoiced a targeted branch & zero-target branch, but owns 0 branches with target > 0
      ['0064000082', '0064000021', '2026-11-19', '107777', 100, 5000],
      ['0064000084', '0064000021', '2026-11-19', '107777', 0, 1000],
    ]),
    'VF05N (2)'
  );
  XLSX.utils.book_append_sheet(
    repsWb,
    XLSX.utils.aoa_to_sheet([
      ['Billing Dt.', 'Customer', 'Material', 'Sales Emp.', 'Qty', 'Total Tax', 'New Growth Value BD&T'],
      [46345, '0064000021', 'M1', '102342', 10, 140, 45000],
      [46345, '0064000777', 'M1', '101071', 15, 250, 75000],
    ]),
    'Zsales_Price (2)'
  );

  const repsReport = validateExcelBuffer(XLSX.write(repsWb, { type: 'buffer', bookType: 'xlsx' }) as Buffer, 'Reps_Achievement.xlsx');
  assert.equal(repsReport.canCommit, true);
  commitValidatedImport(db, repsReport.validationToken, { id: 1, fullName: 'مدير النظام' });

  const repsAllVf = getRepsAnalytics(db, adminUser, { source: 'VF05N' });
  const rep102342 = repsAllVf.find((r) => r.repCode === '102342')!;
  assert.ok(rep102342, 'Rep 102342 must be present');
  assert.equal(rep102342.target, 150000, 'Rep 102342 target must sum 100k + 50k from assigned branches');
  assert.equal(rep102342.sales, 45000);
  assert.equal(rep102342.tax, 140);
  // 45000 / (150000 - 140) * 100 = 30.03%
  assert.equal(rep102342.achievementPct, 30.03);
  assert.equal(rep102342.branchCount, 2, 'Only unique assigned branches with target > 0 must count (exclude 0 and empty target branches)');
  assert.ok(Number.isInteger(rep102342.branchCount));

  // Verify fallback monthly_target from sales_reps when rep has no branch targets
  const repFallback = repsAllVf.find((r) => r.repCode === '108888')!;
  assert.ok(repFallback);
  assert.equal(repFallback.target, 75000);
  assert.equal(repFallback.sales, 0);
  assert.equal(repFallback.achievementPct, 0);
  assert.equal(repFallback.branchCount, 0);

  // Verify safe null (never NaN / Infinity) when target === 0 or (target - tax) <= 0, and branchCount === 0
  const repNoTarget = repsAllVf.find((r) => r.repCode === '107777')!;
  assert.equal(repNoTarget.target, 0);
  assert.equal(repNoTarget.achievementPct, null, 'Zero target must safely return null achievementPct');
  assert.equal(repNoTarget.branchCount, 0, 'Rep with 0 assigned targeted branches must show 0 even if selling to other branches');

  const repNegDenom = repsAllVf.find((r) => r.repCode === '109999')!;
  assert.equal(repNegDenom.target, 500);
  assert.equal(repNegDenom.tax, 600);
  assert.equal(repNegDenom.achievementPct, null, 'Negative (Target - Tax) denominator must safely return null');
  assert.equal(repNegDenom.branchCount, 1);
  assert.equal(calculateAchievementPct(1000, 500, 500), null, 'Zero (Target - Tax) denominator must safely return null');

  // Verify ZSALES consistency with VF05N for target, achievementPct, and branchCount
  const repsAllZs = getRepsAnalytics(db, adminUser, { source: 'ZSALES' });
  const rep102342Zs = repsAllZs.find((r) => r.repCode === '102342')!;
  assert.equal(rep102342Zs.target, 150000);
  assert.equal(rep102342Zs.achievementPct, 30.03);
  assert.equal(rep102342Zs.branchCount, 2);

  // Verify rep target and targeted branch count preserved during zero-sales date filter & safe 0.0% achievement
  const repsEmptyDate = getRepsAnalytics(db, adminUser, { source: 'VF05N', startDate: '2099-01-01', endDate: '2099-01-31' });
  const repEmpty102342 = repsEmptyDate.find((r) => r.repCode === '102342')!;
  assert.ok(repEmpty102342, 'Rep with target must remain visible in narrowed date range');
  assert.equal(repEmpty102342.sales, 0);
  assert.equal(repEmpty102342.target, 150000);
  assert.equal(repEmpty102342.achievementPct, 0);
  assert.equal(repEmpty102342.branchCount, 2);

  // Verify rep filtered by sector (الجيش -> 101071 target = 100,000, sales = 75,000, tax = 250 -> 75.19%, branchCount = 1)
  const repsArmy = getRepsAnalytics(db, adminUser, { source: 'VF05N', sector: 'الجيش' });
  assert.equal(repsArmy.length, 1);
  assert.equal(repsArmy[0].repCode, '101071');
  assert.equal(repsArmy[0].target, 100000);
  assert.equal(repsArmy[0].achievementPct, 75.19);
  assert.equal(repsArmy[0].branchCount, 1);

  // =========================================================================
  // 15. Item Categories Analysis (`تحليل فئات الأصناف`) + Sector Classifications + Category Targets
  // =========================================================================
  const catWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    catWb,
    XLSX.utils.aoa_to_sheet([
      ['كود القطاع', 'القطاع'],
      ['S1', 'التوريد المباشر'],
      ['S2', 'التوريد المركزي'],
      ['S3', 'التجزئة المميزة'], // Variant name -> canonical 'التوب ريتيل'
      ['S4', 'الجيش'],
    ]),
    'القطاعات'
  );
  XLSX.utils.book_append_sheet(
    catWb,
    XLSX.utils.aoa_to_sheet([
      ['القطاع', 'الهدف'],
      ['التوريد المباشر', 500000],
      ['التوريد المركزي', 400000],
      ['التوب ريتيل', 300000],
      ['الجيش', 200000],
    ]),
    'اهداف القطاعات'
  );
  XLSX.utils.book_append_sheet(
    catWb,
    XLSX.utils.aoa_to_sheet([
      ['كود السلسلة', 'اسم السلسلة', 'الادارة', 'اسم المشرف', 'القطاع'],
      ['0064000021', 'سلسلة المباشر', 'ادارة المباشر', 'أحمد زينهم', 'التوريد المباشر'],
      ['0064000022', 'سلسلة المركزي', 'ادارة المركزي', 'كامل حسن', 'التوريد المركزي'],
      ['0064000023', 'سلسلة التوب ريتيل', 'ادارة التجزئة', 'محمد مجدى', 'التجزئة المميزة'],
      ['0064000024', 'منافذ الجيش', 'ادارة الجيش', 'كمال رضا', 'الجيش'],
    ]),
    'تعريف الادارات والقطاعات'
  );
  XLSX.utils.book_append_sheet(
    catWb,
    XLSX.utils.aoa_to_sheet([
      ['كود الفرع', 'اسم الفرع', 'الهدف', 'القطاع'],
      ['0064000081', 'فرع المباشر', 100000, 'التوريد المباشر'],
      ['0064000082', 'فرع المركزي', 100000, 'التوريد المركزي'],
      ['0064000083', 'فرع التوب', 100000, 'التوب ريتيل'],
      ['0064000084', 'فرع الجيش', 100000, 'الجيش'],
    ]),
    'تعريف الفروع و الاهداف'
  );
  XLSX.utils.book_append_sheet(
    catWb,
    XLSX.utils.aoa_to_sheet([
      ['كود المندوب', 'اسم المندوب', 'القطاع', 'اسم المشرف', 'الهدف'],
      ['201', 'أحمد مباشر', 'التوريد المباشر', 'أحمد زينهم', 0],
      ['202', 'خالد مباشر', 'التوريد المباشر', 'أحمد زينهم', 0],
      ['203', 'سالم مركزي', 'التوريد المركزي', 'كامل حسن', 0],
      ['204', 'طارق جيش', 'الجيش', 'كمال رضا', 0],
      ['205', 'ياسر توب', 'التجزئة المميزة', 'محمد مجدى', 0],
    ]),
    'تعريف مندوبى البيع'
  );
  XLSX.utils.book_append_sheet(
    catWb,
    XLSX.utils.aoa_to_sheet([
      ['كود الصنف', 'الصنف', 'التصنيف A', 'تصنيف السلاسل', 'تصنيف الجيش', 'التوب ريتيل'],
      ['ITM-1', 'أرز بسمتي', 'تجارى', '  بقوليات  و  أرز ', 'حصص تموينية', 'بقالة فاخرة'],
      ['ITM-2', 'مكرونة فاخرة', 'تجارى', 'مكرونات', 'مكرونات الجيش', 'معكرونة مستوردة'],
      ['ITM-3', 'زيت عباد', 'تجارى', 'زيوت', 'زيوت الجيش', 'زيوت فاخرة'],
    ]),
    'تعريف الاصناف'
  );
  XLSX.utils.book_append_sheet(
    catWb,
    XLSX.utils.aoa_to_sheet([['اسم المستخدم', 'الصلاحيات'], ['u1', 'الكل']]),
    'تعريف المستخدمين'
  );
  // VF05N intentionally has NO Material codes to assert automatic fallback to ZSALES
  XLSX.utils.book_append_sheet(
    catWb,
    XLSX.utils.aoa_to_sheet([
      ['Sold-to Party', 'Payer', 'Billing Date', 'Sales Employee Code', 'New Total Tax', 'New Growth Value BD&T'],
      ['0064000081', '0064000021', '2026-11-20', '201', 500, 30000],
      ['0064000084', '0064000024', '2026-11-20', '204', 200, 10000],
    ]),
    'VF05N (2)'
  );
  XLSX.utils.book_append_sheet(
    catWb,
    XLSX.utils.aoa_to_sheet([
      ['Billing Dt.', 'Customer', 'Material', 'Sales Emp.', 'Qty', 'Total Tax', 'New Growth Value BD&T'],
      [46346, '0064000021', 'ITM-1', '201', 100, 500, 30000], // احمد مباشر -> بقوليات و ارز
      [46346, '0064000022', 'ITM-2', '203', 50, 100, 15000],  // سالم مركزي -> مكرونات (السلاسل)
      [46346, '0064000023', 'ITM-3', '205', 40, 80, 12000],   // ياسر توب -> زيوت فاخرة (التوب ريتيل)
      [46346, '0064000024', 'ITM-1', '204', 60, 200, 10000],  // طارق جيش -> حصص تموينية (الجيش)
    ]),
    'Zsales_Price (2)'
  );
  // Long-format `اهداف الفئات` sheet (WITHOUT hamza!) + Arabic variant spelling (`بقوليات وارز` without hamza/spaces)
  XLSX.utils.book_append_sheet(
    catWb,
    XLSX.utils.aoa_to_sheet([
      ['القطاع', 'المسمى الوظيفي', 'كود الموظف', 'اسم الموظف', 'الفئة', 'الهدف'],
      ['التوريد المباشر', 'مدير قطاع', 'MGR-1', 'طارق مدير المباشر', 'بقوليات وارز', 60500],
      ['التوريد المباشر', 'مدير قطاع', 'MGR-1', 'طارق مدير المباشر', 'مكرونات', 20000],
      ['التوريد المباشر', 'مدير قطاع', 'MGR-2', 'محمد فؤاد', 'زيوت', 10000],
      ['التوريد المباشر', 'مشرف مبيعات', '', 'احمد زينهم', 'بقوليات و أرز', 50500],
      ['التوريد المباشر', 'مندوب مبيعات', '201', 'أحمد مباشر', 'بقوليات و ارز', 40500],
      ['التوريد المباشر', 'مندوب مبيعات', '201', 'أحمد مباشر', 'زيوت', 0],
      ['الجيش', 'قائد قطاع الجيش', 'ARMY-M1', 'سعيد جيش', 'حصص تموينية', 20200],
    ]),
    'اهداف الفئات'
  );
  // Wide-matrix `أهداف الفئات` sheet (categories as numeric columns + sector/job inferred from sheet name)
  XLSX.utils.book_append_sheet(
    catWb,
    XLSX.utils.aoa_to_sheet([
      ['اسم الموظف', 'القطاع', 'مكرونات الجيش', 'زيوت الجيش'],
      ['كمال رضا', 'الجيش', 15000, 5000],
    ]),
    'أهداف الفئات - مشرف مبيعات'
  );

  const catReport = validateExcelBuffer(
    XLSX.write(catWb, { type: 'buffer', bookType: 'xlsx' }) as Buffer,
    'Category_Targets.xlsx'
  );
  assert.equal(catReport.canCommit, true, `Category workbook errors: ${JSON.stringify(catReport.errors)}`);
  commitValidatedImport(db, catReport.validationToken, { id: 1, fullName: 'مدير النظام' });

  // 15a. Cascading Job Titles depend on selected Sector; Excel `category_targets.target_value` MUST be ignored!
  const directAll = getItemCategoriesAnalytics(db, adminUser, { source: 'VF05N', sector: 'التوريد المباشر' });
  assert.equal(directAll.effectiveSource, 'ZSALES', 'Must auto-fallback from VF05N to ZSALES when VF05N has no Material codes');
  assert.ok(directAll.jobTitleOptions.includes('مدير قطاع'));
  assert.ok(directAll.jobTitleOptions.includes('مشرف مبيعات'));
  assert.ok(directAll.jobTitleOptions.includes('مندوب مبيعات'));
  assert.ok(!directAll.jobTitleOptions.includes('قائد قطاع الجيش'), 'Army job title must NOT leak into Direct Supply dropdown');
  assert.ok(
    directAll.rows.every((r) => r.target === 0),
    'Imported Excel `اهداف الفئات` numeric targets MUST be ignored (all initial targets must be 0)'
  );

  const armyAll = getItemCategoriesAnalytics(db, adminUser, { source: 'ZSALES', sector: 'الجيش' });
  assert.ok(armyAll.jobTitleOptions.includes('قائد قطاع الجيش'));
  assert.ok(armyAll.jobTitleOptions.includes('مشرف قطاع') || armyAll.jobTitleOptions.includes('مشرف مبيعات'));
  assert.ok(armyAll.jobTitleOptions.includes('مندوب مبيعات'));
  assert.ok(!armyAll.jobTitleOptions.includes('مدير قطاع'));

  // 15b. Save direct in-app targets via (القطاع + المسمى الوظيفي + اسم الموظف + الفئة) & verify all 3 categories exist (including 0-sales!)
  saveItemCategoryTargets(db, [
    { sectorName: 'التوريد المباشر', jobTitle: 'مدير قطاع', employeeName: 'طارق مدير المباشر', categoryName: 'بقوليات و أرز', targetValue: 60500 },
    { sectorName: 'التوريد المباشر', jobTitle: 'مدير قطاع', employeeName: 'طارق مدير المباشر', categoryName: 'مكرونات', targetValue: 20000 },
    { sectorName: 'التوريد المباشر', jobTitle: 'مشرف مبيعات', employeeName: 'احمد زينهم', categoryName: 'بقوليات و أرز', targetValue: 50500 },
    { sectorName: 'التوريد المباشر', jobTitle: 'مندوب مبيعات', employeeName: 'أحمد مباشر', categoryName: 'بقوليات و أرز', targetValue: 40500 },
    { sectorName: 'الجيش', jobTitle: 'قائد قطاع الجيش', employeeName: 'سعيد جيش', categoryName: 'حصص تموينية', targetValue: 20200 },
    { sectorName: 'الجيش', jobTitle: 'مشرف مبيعات', employeeName: 'كمال رضا', categoryName: 'مكرونات الجيش', targetValue: 15000 },
    { sectorName: 'الجيش', jobTitle: 'مشرف مبيعات', employeeName: 'كمال رضا', categoryName: 'زيوت الجيش', targetValue: 5000 },
  ]);

  const directMgr = getItemCategoriesAnalytics(db, adminUser, {
    source: 'ZSALES',
    sector: 'التوريد المباشر',
    jobTitle: 'مدير قطاع',
  });
  const tariqRows = directMgr.rows.filter((r) => r.employeeName === 'طارق مدير المباشر');
  const fouadRows = directMgr.rows.filter((r) => r.employeeName === 'محمد فؤاد');
  assert.equal(tariqRows.length, 2, 'طارق مدير المباشر must see ONLY his 2 assigned categories');
  assert.equal(fouadRows.length, 1, 'محمد فؤاد must see ONLY his 1 assigned category (زيوت)');
  assert.ok(!tariqRows.some((r) => r.categoryName === 'زيوت'), 'طارق must NOT inherit محمد فؤاد category');
  assert.ok(!fouadRows.some((r) => r.categoryName.includes('بقوليات')), 'محمد فؤاد must NOT duplicate طارق sales/categories');
  assert.equal(fouadRows[0].categoryName, 'زيوت');
  assert.equal(fouadRows[0].sales, 0);
  assert.equal(fouadRows[0].tax, 0);

  const mgrGrains = tariqRows.find((r) => r.categoryName.includes('بقوليات'))!;
  assert.equal(mgrGrains.sectorName, 'التوريد المباشر');
  assert.equal(mgrGrains.sales, 30000);
  assert.equal(mgrGrains.tax, 500);
  assert.equal(mgrGrains.target, 60500);
  assert.equal(mgrGrains.isAutoTarget, false);
  // 30000 / (60500 - 500) * 100 = 50.0%
  assert.equal(mgrGrains.achievementPct, 50);

  const mgrPasta = tariqRows.find((r) => r.categoryName === 'مكرونات')!;
  assert.equal(mgrPasta.sales, 0, 'Zero-sales assigned category for manager must NOT be filtered out');
  assert.equal(mgrPasta.target, 20000);
  assert.equal(mgrPasta.achievementPct, 0);

  // 15c. Supervisor (`مشرف مبيعات` with Arabic hamza normalization): restricted to his 1 assigned category
  const directSup = getItemCategoriesAnalytics(db, adminUser, {
    source: 'ZSALES',
    sector: 'التوريد المباشر',
    jobTitle: 'مشرف مبيعات',
  });
  assert.equal(directSup.rows.length, 1, 'Supervisor must see ONLY his assigned categories');
  const supGrains = directSup.rows.find((r) => r.categoryName.includes('بقوليات'))!;
  assert.equal(supGrains.employeeName, 'احمد زينهم');
  assert.equal(supGrains.sales, 30000);
  assert.equal(supGrains.tax, 500);
  // 30000 / (50500 - 500) * 100 = 60.0%
  assert.equal(supGrains.achievementPct, 60);

  // 15d. `مندوب مبيعات` in `التوريد المباشر`: 2 reps x 3 Chains categories = 6 rows (preserving ALL 0-sales rows!)
  const directReps = getItemCategoriesAnalytics(db, adminUser, {
    source: 'ZSALES',
    sector: 'التوريد المباشر',
    jobTitle: 'مندوب مبيعات',
  });
  assert.equal(directReps.rows.length, 6, 'Each rep in Direct Supply must get all 3 `تصنيف السلاسل` categories');
  const rep201Grains = directReps.rows.find((r) => r.employeeCode === '201' && r.categoryName.includes('بقوليات'))!;
  assert.equal(rep201Grains.sectorName, 'التوريد المباشر');
  assert.equal(rep201Grains.sales, 30000);
  assert.equal(rep201Grains.tax, 500);
  assert.equal(rep201Grains.target, 40500);
  assert.equal(rep201Grains.achievementPct, 75);

  const rep201Oils = directReps.rows.find((r) => r.employeeCode === '201' && r.categoryName === 'زيوت')!;
  assert.equal(rep201Oils.sales, 0);
  assert.equal(rep201Oils.target, 0);
  assert.equal(rep201Oils.achievementPct, null, 'Target 0 must return null achievementPct (never NaN or Infinity)');

  const rep202Rows = directReps.rows.filter((r) => r.employeeCode === '202');
  assert.equal(rep202Rows.length, 3, 'Zero-sales rep 202 must still appear for all 3 sector categories');
  assert.ok(rep202Rows.every((r) => r.sales === 0));

  // 15e. Central Supply uses `تصنيف السلاسل`, Top Retail (`التجزئة المميزة` -> `التوب ريتيل`) uses `التوب ريتيل`, Army uses `تصنيف الجيش`
  const centralReps = getItemCategoriesAnalytics(db, adminUser, {
    source: 'ZSALES',
    sector: 'التوريد المركزي',
    jobTitle: 'مندوب مبيعات',
  });
  assert.ok(centralReps.rows.some((r) => r.categoryName === 'مكرونات' && r.sales === 15000));

  const topRetailReps = getItemCategoriesAnalytics(db, adminUser, {
    source: 'ZSALES',
    sector: 'التوب ريتيل',
    jobTitle: 'مندوب مبيعات',
  });
  assert.equal(topRetailReps.rows.length, 3);
  assert.ok(topRetailReps.rows.some((r) => r.categoryName === 'زيوت فاخرة' && r.sales === 12000));

  const armyMgr = getItemCategoriesAnalytics(db, adminUser, {
    source: 'ZSALES',
    sector: 'الجيش',
    jobTitle: 'قائد قطاع الجيش',
  });
  assert.equal(armyMgr.rows.length, 1, 'Army sector manager must see ONLY his 1 assigned category');
  const armyRations = armyMgr.rows.find((r) => r.categoryName === 'حصص تموينية')!;
  assert.equal(armyRations.sales, 10000);
  assert.equal(armyRations.tax, 200);
  assert.equal(armyRations.target, 20200);
  assert.equal(armyRations.achievementPct, 50);

  const armySup = getItemCategoriesAnalytics(db, adminUser, {
    source: 'ZSALES',
    sector: 'الجيش',
    jobTitle: 'مشرف مبيعات',
  });
  assert.equal(armySup.rows.length, 2, 'Army supervisor must see ONLY his 2 assigned categories');
  assert.ok(armySup.rows.some((r) => r.categoryName === 'مكرونات الجيش' && r.target === 15000 && r.sales === 0));
  assert.ok(armySup.rows.some((r) => r.categoryName === 'زيوت الجيش' && r.target === 5000 && r.sales === 0));

  // Verify Class A (`التصنيف A`: 'تجارى' / 'صناعى') NEVER appears in Item Categories Analysis
  const allCatAnalysis = getItemCategoriesAnalytics(db, adminUser, { source: 'ZSALES', sector: 'ALL' });
  assert.ok(!allCatAnalysis.rows.some((r) => r.categoryName === 'تجارى' || r.categoryName === 'صناعى'), 'Class A must never appear in Item Categories Analysis');

  // Verify DB persisted has_category_targets_sheet = 1 for hamza-free `اهداف الفئات`
  const dbBatchFlag = db.prepare(`SELECT has_category_targets_sheet FROM import_batches WHERE status = 'active'`).get() as { has_category_targets_sheet: number };
  assert.equal(dbBatchFlag.has_category_targets_sheet, 1);

  // =========================================================================
  // 16. Department Managers (`مدير ادارة` / `مدير إدارة`) + Matrix Isolation:
  //     - Excel targets are ignored
  //     - رافت محمد: auto-computed per category from Central Supply (avoiding double counting!)
  //     - عبدالستار مصطفى: strictly manual in-app editable (NEVER auto-summed from Direct+Army+TopRetail!)
  // =========================================================================
  const dmWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    dmWb,
    XLSX.utils.aoa_to_sheet([
      ['كود القطاع', 'القطاع'],
      ['S1', 'التوريد المباشر'],
      ['S2', 'التوريد المركزي'],
      ['S3', 'التوب ريتيل'],
      ['S4', 'الجيش'],
    ]),
    'القطاعات'
  );
  XLSX.utils.book_append_sheet(
    dmWb,
    XLSX.utils.aoa_to_sheet([
      ['القطاع', 'الهدف'],
      ['التوريد المباشر', 500000],
      ['التوريد المركزي', 400000],
      ['التوب ريتيل', 300000],
      ['الجيش', 200000],
    ]),
    'اهداف القطاعات'
  );
  XLSX.utils.book_append_sheet(
    dmWb,
    XLSX.utils.aoa_to_sheet([
      ['كود السلسلة', 'اسم السلسلة', 'الادارة', 'اسم المشرف', 'القطاع'],
      ['0064000021', 'سلسلة المباشر', 'ادارة المباشر', 'محمد مجدى', 'التوريد المباشر'],
      ['0064000022', 'سلسلة المركزي', 'ادارة المركزي', 'كمال رضا', 'التوريد المركزي'],
      ['0064000023', 'سلسلة التوب ريتيل', 'ادارة التجزئة', 'أحمد زينهم', 'التوب ريتيل'],
      ['0064000024', 'منافذ الجيش', 'ادارة الجيش', 'كامل حسن', 'الجيش'],
    ]),
    'تعريف الادارات والقطاعات'
  );
  XLSX.utils.book_append_sheet(
    dmWb,
    XLSX.utils.aoa_to_sheet([
      ['كود الفرع', 'اسم الفرع', 'الهدف', 'القطاع'],
      ['0064000081', 'فرع المباشر', 100000, 'التوريد المباشر'],
      ['0064000082', 'فرع المركزي', 100000, 'التوريد المركزي'],
      ['0064000083', 'فرع التوب', 100000, 'التوب ريتيل'],
      ['0064000084', 'فرع الجيش', 100000, 'الجيش'],
    ]),
    'تعريف الفروع و الاهداف'
  );
  XLSX.utils.book_append_sheet(
    dmWb,
    XLSX.utils.aoa_to_sheet([
      ['كود المندوب', 'اسم المندوب', 'القطاع', 'اسم المشرف', 'الهدف'],
      ['301', 'مندوب مباشر', 'التوريد المباشر', 'محمد مجدى', 0],
      ['302', 'مندوب مركزي', 'التوريد المركزي', 'كمال رضا', 0],
      ['303', 'مندوب توب', 'التوب ريتيل', 'أحمد زينهم', 0],
      ['304', 'مندوب جيش', 'الجيش', 'كامل حسن', 0],
    ]),
    'تعريف مندوبى البيع'
  );
  // Note: ITM-A & ITM-B have DIFFERENT category names in `تصنيف الجيش` and `التوب ريتيل` vs `تصنيف السلاسل`!
  XLSX.utils.book_append_sheet(
    dmWb,
    XLSX.utils.aoa_to_sheet([
      ['كود الصنف', 'الصنف', 'التصنيف A', 'تصنيف السلاسل', 'تصنيف الجيش', 'التوب ريتيل'],
      ['ITM-A', 'بسكويت سادة', 'تجارى', 'بسكويت + حلويات', 'بسكويت أنواع', 'كيك + حلويات'],
      ['ITM-B', 'كيك شيكولاتة', 'تجارى', 'كيك + كرواسون', 'كيك أنواع', 'كيك + حلويات'],
      ['ITM-C', 'قهوة', 'تجارى', 'مشروبات ساخنة', 'متنوع', 'مشروبات ساخنة'],
    ]),
    'تعريف الاصناف'
  );
  XLSX.utils.book_append_sheet(
    dmWb,
    XLSX.utils.aoa_to_sheet([['اسم المستخدم', 'الصلاحيات'], ['u1', 'الكل']]),
    'تعريف المستخدمين'
  );
  XLSX.utils.book_append_sheet(
    dmWb,
    XLSX.utils.aoa_to_sheet([
      ['Sold-to Party', 'Payer', 'Billing Date', 'Sales Employee Code', 'New Total Tax', 'New Growth Value BD&T'],
      ['0064000081', '0064000021', '2026-11-21', '301', 100, 10000],
      ['0064000082', '0064000022', '2026-11-21', '302', 200, 20000],
      ['0064000083', '0064000023', '2026-11-21', '303', 300, 30000],
      ['0064000084', '0064000024', '2026-11-21', '304', 400, 40000],
    ]),
    'VF05N (2)'
  );
  XLSX.utils.book_append_sheet(
    dmWb,
    XLSX.utils.aoa_to_sheet([
      ['Billing Dt.', 'Customer', 'Material', 'Sales Emp.', 'Qty', 'Total Tax', 'New Growth Value BD&T'],
      [46347, '0064000021', 'ITM-A', '301', 10, 100, 10000],
      [46347, '0064000022', 'ITM-A', '302', 20, 200, 20000],
      [46347, '0064000023', 'ITM-A', '303', 30, 300, 30000],
      [46347, '0064000024', 'ITM-A', '304', 40, 400, 40000],
      [46347, '0064000024', 'ITM-B', '304', 15, 0, 15000],
    ]),
    'Zsales_Price (2)'
  );
  // Excel sheet includes bogus targets to assert they are completely ignored by the app!
  XLSX.utils.book_append_sheet(
    dmWb,
    XLSX.utils.aoa_to_sheet([
      ['اسم القطاع', 'المسمى الوظيفى', 'اسم الموظف', 'الفئة', 'الهدف'],
      ['التوريد المركزي', 'مدير إدارة', 'رأفت محمد', 'بسكويت + حلويات', 999999],
      ['التوريد المركزي', 'مدير إدارة', 'رأفت محمد', 'مشروبات ساخنة', 111111],
      ['التوريد المركزي', 'مدير قطاع', 'مصطفى مناع', 'بسكويت + حلويات', 888888],
      ['التوريد المركزي', 'مدير قطاع', 'محمد الملا', 'مشروبات ساخنة', 666666],
      ['التوريد المباشر + الجيش + التوب ريتيل', 'مدير ادارة', 'عبد الستار مصطفي', 'بسكويت + حلويات', 777777],
      ['التوريد المباشر + الجيش + التوب ريتيل', 'مدير ادارة', 'عبد الستار مصطفي', 'كيك + كرواسون', 555555],
      ['التوريد المباشر + الجيش + التوب ريتيل', 'مدير ادارة', 'عبد الستار مصطفي', 'مشروبات ساخنة', 444444],
    ]),
    'اهداف الفئات'
  );

  const dmReport = validateExcelBuffer(
    XLSX.write(dmWb, { type: 'buffer', bookType: 'xlsx' }) as Buffer,
    'Dept_Managers.xlsx'
  );
  assert.equal(dmReport.canCommit, true, `Dept managers workbook errors: ${JSON.stringify(dmReport.errors)}`);
  commitValidatedImport(db, dmReport.validationToken, { id: 1, fullName: 'مدير النظام' });

  // 16a. Initially after import, NO Excel targets are used -> all targets in dmWb are 0 until entered in-app
  const dmInitial = getItemCategoriesAnalytics(db, adminUser, {
    source: 'ZSALES',
    sector: 'ALL',
    jobTitle: 'مدير ادارة',
  });
  assert.ok(dmInitial.rows.every((r) => r.target === 0), 'Excel targets 999999 / 777777 MUST be ignored');

  // 16b. Save Central Supply rep + supervisor targets -> رافت محمد auto-computes without double counting!
  // Also save Direct Supply rep target -> عبدالستار مصطفى MUST remain 0 (NO auto-rollup!)
  saveItemCategoryTargets(db, [
    { sectorName: 'التوريد المركزي', jobTitle: 'مندوب مبيعات', employeeName: 'مندوب مركزي', categoryName: 'بسكويت + حلويات', targetValue: 30000 },
    { sectorName: 'التوريد المركزي', jobTitle: 'مشرف قطاع', employeeName: 'كمال رضا', categoryName: 'بسكويت + حلويات', targetValue: 40200 },
    { sectorName: 'التوريد المركزي', jobTitle: 'مشرف قطاع', employeeName: 'كمال رضا', categoryName: 'مشروبات ساخنة', targetValue: 10000 },
    { sectorName: 'التوريد المباشر', jobTitle: 'مندوب مبيعات', employeeName: 'مندوب مباشر', categoryName: 'بسكويت + حلويات', targetValue: 55000 },
  ]);

  const centralDm = getItemCategoriesAnalytics(db, adminUser, {
    source: 'ZSALES',
    sector: 'التوريد المركزي',
    jobTitle: 'مدير ادارة',
  });
  assert.ok(centralDm.jobTitleOptions.some((j) => j.includes('مدير')));
  assert.equal(centralDm.rows.length, 2, 'رافت محمد must see ONLY his 2 assigned Central Supply categories');
  assert.ok(centralDm.rows.every((r) => (r.employeeName.includes('رافت') || r.employeeName.includes('رأفت')) && r.isAutoTarget === true));

  const rafatBiscuit = centralDm.rows.find((r) => r.categoryName === 'بسكويت + حلويات')!;
  assert.equal(rafatBiscuit.sectorName, 'التوريد المركزي');
  assert.equal(rafatBiscuit.sales, 20000);
  assert.equal(rafatBiscuit.tax, 200);
  // Waterfall prefers supervisor sum (40,200) over rep sum (30,000) -> 40,200 (NO 70,200 double counting!)
  assert.equal(rafatBiscuit.target, 40200);
  assert.equal(rafatBiscuit.achievementPct, 50);

  const rafatZero = centralDm.rows.find((r) => r.categoryName === 'مشروبات ساخنة')!;
  assert.equal(rafatZero.sales, 0);
  assert.equal(rafatZero.target, 10000);
  assert.equal(rafatZero.achievementPct, 0);

  // When a Central Supply Sector Manager (`مدير قطاع`) target is entered (e.g. 50,200), رافت محمد automatically updates to 50,200!
  saveItemCategoryTargets(db, [
    { sectorName: 'التوريد المركزي', jobTitle: 'مدير قطاع', employeeName: 'مصطفى مناع', categoryName: 'بسكويت + حلويات', targetValue: 50200 },
  ]);
  const centralDmUpdated = getItemCategoriesAnalytics(db, adminUser, {
    source: 'ZSALES',
    sector: 'التوريد المركزي',
    jobTitle: 'مدير ادارة',
  });
  assert.equal(
    centralDmUpdated.rows.find((r) => r.categoryName === 'بسكويت + حلويات')!.target,
    50200,
    'رافت محمد must auto-update when Central Supply مدير قطاع target is saved without double counting'
  );

  // Reset مصطفى مناع to 0 -> reverts cleanly to supervisor sum (40,200)
  saveItemCategoryTargets(db, [
    { sectorName: 'التوريد المركزي', jobTitle: 'مدير قطاع', employeeName: 'مصطفى مناع', categoryName: 'بسكويت + حلويات', targetValue: 0 },
  ]);

  // 16c. عبدالستار مصطفى MUST remain 0 until edited directly (never auto-summed from Direct+Army+TopRetail!),
  //      and then uses his directly saved targets!
  const abdelBeforeEdit = getItemCategoriesAnalytics(db, adminUser, {
    source: 'ZSALES',
    sector: 'التوريد المباشر',
    jobTitle: 'مدير ادارة',
  });
  assert.ok(abdelBeforeEdit.rows.every((r) => r.isAutoTarget === false && r.target === 0), 'عبدالستار مصطفى must NOT auto-sum subordinate sectors');

  saveItemCategoryTargets(db, [
    { sectorName: 'التوريد المباشر + الجيش + التوب ريتيل', jobTitle: 'مدير ادارة', employeeName: 'عبد الستار مصطفي', categoryName: 'بسكويت + حلويات', targetValue: 160800 },
    { sectorName: 'التوريد المباشر + الجيش + التوب ريتيل', jobTitle: 'مدير ادارة', employeeName: 'عبد الستار مصطفي', categoryName: 'كيك + كرواسون', targetValue: 30000 },
    { sectorName: 'التوريد المباشر + الجيش + التوب ريتيل', jobTitle: 'مدير ادارة', employeeName: 'عبد الستار مصطفي', categoryName: 'مشروبات ساخنة', targetValue: 25000 },
  ]);

  for (const sec of ['التوريد المباشر', 'الجيش', 'التوب ريتيل'] as const) {
    const abdelRes = getItemCategoriesAnalytics(db, adminUser, {
      source: 'ZSALES',
      sector: sec,
      jobTitle: 'مدير ادارة',
    });
    assert.ok(abdelRes.jobTitleOptions.some((j) => j.includes('مدير')), `مدير ادارة must appear under ${sec}`);
    assert.equal(abdelRes.rows.length, 3, `عبدالستار مصطفى must show all 3 categories under ${sec}`);
    assert.ok(abdelRes.rows.every((r) => r.employeeName.includes('عبد') && r.isAutoTarget === false));

    const biscuitRow = abdelRes.rows.find((r) => r.categoryName === 'بسكويت + حلويات')!;
    assert.equal(biscuitRow.sales, 80000, `Under ${sec}, عبدالستار must sum Direct+Army+TopRetail using تصنيف السلاسل`);
    assert.equal(biscuitRow.tax, 800);
    assert.equal(biscuitRow.target, 160800);
    assert.equal(biscuitRow.achievementPct, 50);

    const cakeRow = abdelRes.rows.find((r) => r.categoryName === 'كيك + كرواسون')!;
    assert.equal(cakeRow.sales, 15000);
    assert.equal(cakeRow.target, 30000);
    assert.equal(cakeRow.achievementPct, 50);

    const zeroHotRow = abdelRes.rows.find((r) => r.categoryName === 'مشروبات ساخنة')!;
    assert.equal(zeroHotRow.sales, 0);
    assert.equal(zeroHotRow.target, 25000);
    assert.equal(zeroHotRow.achievementPct, 0);
  }

  // Verify Central Supply `مدير قطاع` separation: مصطفى مناع sees only `بسكويت + حلويات`, محمد الملا sees only `مشروبات ساخنة`
  const centralSecMgrs = getItemCategoriesAnalytics(db, adminUser, {
    source: 'ZSALES',
    sector: 'التوريد المركزي',
    jobTitle: 'مدير قطاع',
  });
  assert.equal(centralSecMgrs.rows.length, 2);
  const mostafaRow = centralSecMgrs.rows.find((r) => r.employeeName === 'مصطفى مناع')!;
  const mullaRow = centralSecMgrs.rows.find((r) => r.employeeName === 'محمد الملا')!;
  assert.equal(mostafaRow.categoryName, 'بسكويت + حلويات');
  assert.equal(mostafaRow.sales, 20000);
  assert.equal(mullaRow.categoryName, 'مشروبات ساخنة');
  assert.equal(mullaRow.sales, 0);

  // 16d. Under `ALL` sectors + `مدير ادارة`, BOTH رافت محمد (2 assigned rows) and عبدالستار مصطفى (3 assigned rows) appear = 5 rows
  const allDm = getItemCategoriesAnalytics(db, adminUser, {
    source: 'ZSALES',
    sector: 'ALL',
    jobTitle: 'مدير ادارة',
  });
  assert.equal(allDm.rows.length, 5);

  // 16e. Verify in-app saved targets persist even after committing a fresh Excel import!
  const dmReportPersistence = validateExcelBuffer(
    XLSX.write(dmWb, { type: 'buffer', bookType: 'xlsx' }) as Buffer,
    'Dept_Managers_Reimport.xlsx'
  );
  commitValidatedImport(db, dmReportPersistence.validationToken, { id: 1, fullName: 'مدير النظام' });
  const afterReimportAbdel = getItemCategoriesAnalytics(db, adminUser, {
    source: 'ZSALES',
    sector: 'التوريد المباشر',
    jobTitle: 'مدير ادارة',
  });
  assert.equal(
    afterReimportAbdel.rows.find((r) => r.categoryName === 'بسكويت + حلويات')!.target,
    160800,
    'In-app saved category targets must survive new Excel imports'
  );

  // =========================================================================
  // 17. Quantity-Based Competitions (`مسابقات`):
  //     - Multi-select catalog items from `تعريف الاصناف`
  //     - Standard strength factor (`معامل الشدة المعيارية`) up to 3 decimals (0.333, 0.667, 1.25)
  //     - `QTY * factor` pulled strictly from Zsales_Price (2)
  //     - Editable targets persisted across reloads & new imports
  //     - Auto-rollup target for مديري الإدارات (رافت محمد & عبدالستار مصطفى)
  //     - Zero-quantity participants retained with 0 (never dropped), safe remaining/pct
  // =========================================================================
  const catalog = getAvailableCatalogItems(db);
  assert.equal(catalog.length, 3);
  assert.ok(catalog.some((c) => c.itemCode === 'ITM-A'));

  const comp1 = createCompetition(
    db,
    { id: 1 },
    {
      name: 'مسابقة الشدة المعيارية',
      items: [
        { itemCode: 'ITM-A', itemName: 'بسكويت سادة', equivalentFactor: 0.333 },
        { itemCode: 'ITM-B', itemName: 'كيك شيكولاتة', equivalentFactor: 1.25 },
      ],
    }
  );
  assert.equal(comp1.items.length, 2);
  assert.equal(comp1.items.find((i) => i.itemCode === 'ITM-A')!.equivalentFactor, 0.333);
  assert.equal(comp1.items.find((i) => i.itemCode === 'ITM-B')!.equivalentFactor, 1.25);

  // 17a. Verify QTY * equivalentFactor from Zsales_Price (2)
  const compAnalytics1 = getCompetitionAnalytics(db, adminUser, comp1.id, { sector: 'ALL' });
  const rep301 = compAnalytics1.participants.find((p) => p.roleKind === 'sales_rep' && p.employeeCode === '301')!;
  const rep302 = compAnalytics1.participants.find((p) => p.roleKind === 'sales_rep' && p.employeeCode === '302')!;
  const rep303 = compAnalytics1.participants.find((p) => p.roleKind === 'sales_rep' && p.employeeCode === '303')!;
  const rep304 = compAnalytics1.participants.find((p) => p.roleKind === 'sales_rep' && p.employeeCode === '304')!;

  assert.equal(rep301.soldStandardQty, 3.33); // 10 * 0.333
  assert.equal(rep302.soldStandardQty, 6.66); // 20 * 0.333
  assert.equal(rep303.soldStandardQty, 9.99); // 30 * 0.333
  assert.equal(rep304.soldStandardQty, 32.07); // 40 * 0.333 + 15 * 1.25 = 13.32 + 18.75

  // With target = 0 -> safe null achievementPct and remainingQty = 0
  assert.equal(rep301.targetQty, 0);
  assert.equal(rep301.remainingQty, 0);
  assert.equal(rep301.achievementPct, null);
  assert.equal(calculateQtyAchievementPct(10, 0), null);

  // 17b. Save targets for reps and verify live rollup for مديري الإدارات + remainingQty + %
  saveCompetitionTargets(db, comp1.id, [
    { participantKey: rep301.participantKey, targetQty: 10 },
    { participantKey: rep302.participantKey, targetQty: 10 },
    { participantKey: rep303.participantKey, targetQty: 20 },
    { participantKey: rep304.participantKey, targetQty: 50 },
  ]);

  const compAnalyticsWithTargets = getCompetitionAnalytics(db, adminUser, comp1.id, { sector: 'ALL' });
  const rep301After = compAnalyticsWithTargets.participants.find((p) => p.participantKey === rep301.participantKey)!;
  assert.equal(rep301After.targetQty, 10);
  assert.equal(rep301After.remainingQty, 6.67);
  assert.equal(rep301After.achievementPct, 33.3);

  const rep304After = compAnalyticsWithTargets.participants.find((p) => p.participantKey === rep304.participantKey)!;
  assert.equal(rep304After.targetQty, 50);
  assert.equal(rep304After.remainingQty, 17.93);
  assert.equal(rep304After.achievementPct, 64.14);

  const rafatComp = compAnalyticsWithTargets.participants.find(
    (p) => p.roleKind === 'department_manager' && p.employeeName.includes('رافت')
  ) ?? compAnalyticsWithTargets.participants.find(
    (p) => p.roleKind === 'department_manager' && p.employeeName.includes('رأفت')
  )!;
  assert.equal(rafatComp.isAutoTarget, true);
  assert.equal(rafatComp.targetQty, 10); // rep302 target
  assert.equal(rafatComp.soldStandardQty, 6.66);
  assert.equal(rafatComp.remainingQty, 3.34);
  assert.equal(rafatComp.achievementPct, 66.6);

  const abdelComp = compAnalyticsWithTargets.participants.find(
    (p) => p.roleKind === 'department_manager' && (p.employeeName.includes('عبدالستار') || p.employeeName.includes('عبد الستار'))
  )!;
  assert.equal(abdelComp.isAutoTarget, true);
  assert.equal(abdelComp.targetQty, 80); // 10 + 20 + 50
  assert.equal(abdelComp.soldStandardQty, 45.39); // 3.33 + 9.99 + 32.07
  assert.equal(abdelComp.remainingQty, 34.61);
  assert.equal(abdelComp.achievementPct, 56.74);

  // When target is already exceeded -> remainingQty stays safely 0 (no negative remaining)
  saveCompetitionTargets(db, comp1.id, [{ participantKey: rep301.participantKey, targetQty: 2 }]);
  const exceededRes = getCompetitionAnalytics(db, adminUser, comp1.id, { sector: 'التوريد المباشر' });
  const rep301Exceeded = exceededRes.participants.find((p) => p.participantKey === rep301.participantKey)!;
  assert.equal(rep301Exceeded.remainingQty, 0);
  assert.equal(rep301Exceeded.achievementPct, 166.5);
  // Restore rep301 target to 10 and test sector filter scoping
  saveCompetitionTargets(db, comp1.id, [{ participantKey: rep301.participantKey, targetQty: 10 }]);
  const directScoped = getCompetitionAnalytics(db, adminUser, comp1.id, { sector: 'التوريد المباشر' });
  assert.ok(directScoped.participants.some((p) => p.roleKind === 'department_manager' && p.targetQty === 80));
  assert.ok(!directScoped.participants.some((p) => p.roleKind === 'department_manager' && (p.employeeName.includes('رافت') || p.employeeName.includes('رأفت'))));

  // 17c. Zero-quantity preservation: competition with ITM-C (0 sales) keeps every participant with 0
  const zeroComp = createCompetition(
    db,
    { id: 1 },
    {
      name: 'مسابقة بدون مبيعات',
      items: [{ itemCode: 'ITM-C', itemName: 'قهوة', equivalentFactor: 0.667 }],
    }
  );
  const zeroRes = getCompetitionAnalytics(db, adminUser, zeroComp.id, { sector: 'ALL' });
  assert.ok(zeroRes.participants.length >= 10, 'Zero-qty participants must never be excluded');
  assert.ok(zeroRes.participants.every((p) => p.soldStandardQty === 0));

  // 17d. Update competition factor to 0.667 & persist across a new Excel batch commit
  const updatedComp = updateCompetition(db, comp1.id, {
    items: [{ itemCode: 'ITM-A', itemName: 'بسكويت سادة', equivalentFactor: 0.667 }],
  });
  assert.equal(updatedComp.items.length, 1);
  assert.equal(updatedComp.items[0].equivalentFactor, 0.667);

  const dmReport2 = validateExcelBuffer(
    XLSX.write(dmWb, { type: 'buffer', bookType: 'xlsx' }) as Buffer,
    'Dept_Managers_Reloaded.xlsx'
  );
  commitValidatedImport(db, dmReport2.validationToken, { id: 1, fullName: 'مدير النظام' });
  const afterNewImport = getCompetitionAnalytics(db, adminUser, comp1.id, { sector: 'ALL' });
  const rep301Reloaded = afterNewImport.participants.find((p) => p.participantKey === rep301.participantKey)!;
  assert.equal(rep301Reloaded.targetQty, 10, 'Saved competition targets must survive new Excel imports');
  assert.equal(rep301Reloaded.soldStandardQty, 6.67, '10 * 0.667 must recalculate cleanly after update & new import');

  deleteCompetition(db, zeroComp.id);
  assert.equal(listCompetitions(db).length, 1);

  // =========================================================================
  // 18. Branch & Rep Deduplication, Consolidated Same-Name Multi-Code Branches,
  //     Supervisor Isolation, Leading Zeros, & Cross-Rep Diagnostics
  // =========================================================================
  const dedupWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    dedupWb,
    XLSX.utils.aoa_to_sheet([
      ['كود القطاع', 'القطاع'],
      ['S1', 'التوريد المباشر'],
      ['S2', 'التوريد المركزي'],
    ]),
    'القطاعات'
  );
  XLSX.utils.book_append_sheet(
    dedupWb,
    XLSX.utils.aoa_to_sheet([
      ['القطاع', 'الهدف'],
      ['التوريد المباشر', 500000],
      ['التوريد المركزي', 300000],
    ]),
    'اهداف القطاعات'
  );
  XLSX.utils.book_append_sheet(
    dedupWb,
    XLSX.utils.aoa_to_sheet([
      ['كود السلسلة', 'اسم السلسلة', 'الادارة', 'اسم المشرف', 'القطاع'],
      ['0064000050', 'ال اس جى', 'ادارة المباشر', 'محمد مجدى', 'التوريد المباشر'],
      ['0064000051', 'سبينس', 'ادارة المباشر', 'محمد عبدالمحسن', 'التوريد المباشر'],
      ['0064000052', 'الشريف للجملة', 'ادارة المركزي', 'كمال رضا', 'التوريد المركزي'],
    ]),
    'تعريف الادارات والقطاعات'
  );
  XLSX.utils.book_append_sheet(
    dedupWb,
    XLSX.utils.aoa_to_sheet([
      ['القطاع', 'كود الفرع', 'اسم الفرع', 'كود الادارة', 'اسم الادارة', 'كود المندوب', 'اسم المندوب', 'اسم المشرف', 'إجمالى الهدف'],
      // Same branch name (`ال اس جى`) with TWO codes & leading zeros for rep 101385: one has sales, other has target
      ['التوريد المباشر', '0064001238', 'ال اس جى', '0064000050', 'ال اس جى', '101385', 'محمد منصور', 'محمد مجدى', 0],
      ['التوريد المباشر', '0064001383', 'ال اس جي', '0064000050', 'ال اس جى', '101385', 'محمد منصور', 'محمد مجدى', 50000],
      // Same branch name (`سبينس المنصورة`) with TWO codes where BOTH have target > 0 for rep 101385 -> must count as 1 unique branch!
      ['التوريد المباشر', '0064001266', 'سبينس المنصورة', '0064000051', 'سبينس', '101385', 'محمد منصور', 'محمد عبدالمحسن', 10000],
      ['التوريد المباشر', '0064001269', 'سبينس المنصوره', '0064000051', 'سبينس', '101385', 'محمد منصور', 'محمد عبدالمحسن', 20000],
      // Branch assigned to rep 103309 (`معاز محمود`), which another rep (103209) bills to in VF05N -> must NOT inflate 103209!
      ['التوريد المباشر', '0064000349', 'اسواق فتح الله', '0064000050', 'فتح الله', '103309', 'معاز محمود', 'محمد مجدى', 60000],
      // Same branch name (`الشريف للجملة`) under TWO DIFFERENT supervisors -> must NOT merge across supervisors!
      ['التوريد المركزي', '0064000044', 'الشريف للجملة', '0064000052', 'الشريف', '', '', 'احمد زينهم', 0],
      ['التوريد المركزي', '0064000463', 'الشريف للجمله', '0064000052', 'الشريف', '', '', 'كمال رضا', 0],
    ]),
    'تعريف الفروع و الاهداف'
  );
  XLSX.utils.book_append_sheet(
    dedupWb,
    XLSX.utils.aoa_to_sheet([
      ['كود المندوب', 'اسم المندوب', 'القطاع', 'اسم المشرف', 'الهدف'],
      ['101385', 'محمد منصور', 'التوريد المباشر', 'محمد مجدى', 0],
      ['103309', 'معاز محمود', 'التوريد المباشر', 'محمد مجدى', 0],
      ['103209', 'عمر خالد', 'التوريد المباشر', 'محمد عبدالمحسن', 0],
    ]),
    'تعريف مندوبى البيع'
  );
  XLSX.utils.book_append_sheet(
    dedupWb,
    XLSX.utils.aoa_to_sheet([['كود الصنف', 'اسم الصنف', 'التصنيف A'], ['M1', 'صنف', 'تجارى']]),
    'تعريف الاصناف'
  );
  XLSX.utils.book_append_sheet(dedupWb, XLSX.utils.aoa_to_sheet([['اسم المستخدم', 'الصلاحيات'], ['u1', 'الكل']]), 'تعريف المستخدمين');
  XLSX.utils.book_append_sheet(
    dedupWb,
    XLSX.utils.aoa_to_sheet([
      ['Sold-to Party', 'Payer', 'Billing Date', 'Sales Employee Code', 'New Total Tax', 'New Growth Value BD&T'],
      // Sales split across both `ال اس جى` codes:
      ['0064001238', '0064000050', '2026-11-22', '101385', 500, 40000],
      ['0064001383', '0064000050', '2026-11-22', '101385', 100, 15000],
      // Cross-rep supply: rep 103209 bills to rep 103309's branch `0064000349`
      ['0064000349', '0064000050', '2026-11-22', '103209', 0, 12000],
    ]),
    'VF05N (2)'
  );
  XLSX.utils.book_append_sheet(
    dedupWb,
    XLSX.utils.aoa_to_sheet([
      ['Billing Dt.', 'Customer', 'Material', 'Sales Emp.', 'Qty', 'Total Tax', 'New Growth Value BD&T'],
      [46348, '0064000050', 'M1', '101385', 10, 600, 67000],
    ]),
    'Zsales_Price (2)'
  );

  const dedupReport = validateExcelBuffer(
    XLSX.write(dedupWb, { type: 'buffer', bookType: 'xlsx' }) as Buffer,
    'Dedup_Branches.xlsx'
  );
  assert.equal(dedupReport.canCommit, true, `Dedup workbook errors: ${JSON.stringify(dedupReport.errors)}`);
  commitValidatedImport(db, dedupReport.validationToken, { id: 1, fullName: 'مدير النظام' });

  const dedupBranchesRes = getBranchesAnalytics(db, adminUser, { source: 'VF05N' });
  // 7 raw rows -> `ال اس جى` (2->1) and `سبينس المنصورة` (2->1) consolidate -> 5 rows total!
  // `الشريف للجملة` (2 different supervisors) MUST stay 2 separate rows!
  assert.equal(dedupBranchesRes.branches.length, 5, 'Duplicate names under same rep/supervisor must consolidate; different supervisors must NOT merge');

  const lsg = dedupBranchesRes.branches.find((b) => b.branchName.includes('ال اس'))!;
  assert.ok(lsg.branchCode.includes('0064001238') && lsg.branchCode.includes('0064001383'), 'Leading zeros preserved on consolidated branch codes');
  assert.equal(lsg.sales, 55000, 'Sales across both `ال اس جى` codes must sum without duplication');
  assert.equal(lsg.tax, 600);
  assert.equal(lsg.target, 50000);

  const spinneysMans = dedupBranchesRes.branches.find((b) => b.branchName.includes('سبينس'))!;
  assert.equal(spinneysMans.target, 30000, 'Targets across merged codes sum once without duplication');

  const sharifRows = dedupBranchesRes.branches.filter((b) => b.branchName.includes('الشريف'));
  assert.equal(sharifRows.length, 2, 'Branches under distinct supervisors must never be merged');

  const dedupReps = getRepsAnalytics(db, adminUser, { source: 'VF05N' });
  const repMansour = dedupReps.find((r) => r.repCode === '101385')!;
  // 4 raw rows in Excel across 2 unique normalized branch names (`ال اس جى` + `سبينس المنصورة`), both with combined target > 0 -> branchCount === 2!
  assert.equal(repMansour.branchCount, 2, 'Rep 101385 must count exactly 2 unique branches despite 4 branch codes');
  assert.equal(repMansour.target, 80000); // 50,000 + 30,000

  const repOmar = dedupReps.find((r) => r.repCode === '103209')!;
  assert.equal(repOmar.branchCount, 0, 'Cross-rep invoice on 0064000349 must NOT inflate rep 103209 branchCount');

  assert.equal(dedupBranchesRes.diagnostics.totalTargetedAssignedBranches, 3); // 2 (101385) + 1 (103309)
  assert.equal(dedupBranchesRes.diagnostics.consolidatedDuplicateNameCount, 2);
  assert.ok(dedupBranchesRes.diagnostics.repDiscrepancies.some((d) => d.repCode === '103209' && d.extraCount === 1));

  // =========================================================================
  // 19. Unified Sector Sales & Tax Totals Across ALL Job Titles (`مندوب` / `مشرف` / `مدير قطاع`)
  //     Even when geographical chain supervisors in `departments_chains` sell ALL categories or omit a category-only supervisor!
  // =========================================================================
  const unifiedWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    unifiedWb,
    XLSX.utils.aoa_to_sheet([
      ['كود القطاع', 'القطاع'],
      ['S1', 'التوريد المباشر'],
      ['S2', 'التوريد المركزي'],
    ]),
    'القطاعات'
  );
  XLSX.utils.book_append_sheet(
    unifiedWb,
    XLSX.utils.aoa_to_sheet([
      ['القطاع', 'الهدف'],
      ['التوريد المباشر', 500000],
      ['التوريد المركزي', 500000],
    ]),
    'اهداف القطاعات'
  );
  XLSX.utils.book_append_sheet(
    unifiedWb,
    XLSX.utils.aoa_to_sheet([
      ['كود السلسلة', 'اسم السلسلة', 'الادارة', 'اسم المشرف', 'القطاع'],
      // In Direct Supply, two chain supervisors divide CHAINS geographically while each chain sells ALL categories!
      ['0064000001', 'سلسلة الراية', 'محمد مجدى', 'محمد مجدى', 'التوريد المباشر'],
      ['0064000002', 'سلسلة سبينس', 'محمد عبدالمحسن', 'محمد عبدالمحسن', 'التوريد المباشر'],
      // In Central Supply, ONLY `كمال رضا` is listed on the chain (`أحمد سيد` is NOT listed on any chain!)
      ['0064000003', 'سلسلة كازيون', 'كمال رضا', 'كمال رضا', 'التوريد المركزي'],
    ]),
    'تعريف الادارات والقطاعات'
  );
  XLSX.utils.book_append_sheet(
    unifiedWb,
    XLSX.utils.aoa_to_sheet([
      ['كود الفرع', 'اسم الفرع', 'الهدف', 'القطاع'],
      ['0064000101', 'فرع 1', 10000, 'التوريد المباشر'],
      ['0064000102', 'فرع 2', 10000, 'التوريد المركزي'],
    ]),
    'تعريف الفروع و الاهداف'
  );
  XLSX.utils.book_append_sheet(
    unifiedWb,
    XLSX.utils.aoa_to_sheet([
      ['كود المندوب', 'اسم المندوب', 'القطاع', 'اسم المشرف', 'الهدف'],
      ['401', 'مندوب مباشر 1', 'التوريد المباشر', '', 0],
      ['402', 'مندوب مباشر 2', 'التوريد المباشر', '', 0],
      ['403', 'مندوب مركزي 1', 'التوريد المركزي', '', 0],
    ]),
    'تعريف مندوبى البيع'
  );
  XLSX.utils.book_append_sheet(
    unifiedWb,
    XLSX.utils.aoa_to_sheet([
      ['كود الصنف', 'الصنف', 'التصنيف A', 'تصنيف السلاسل', 'تصنيف الجيش', 'التوب ريتيل'],
      ['U-1', 'بسكويت', 'تجارى', 'بسكويت + حلويات', '', ''],
      ['U-2', 'كيك', 'تجارى', 'كيك + كرواسون', '', ''],
      ['U-3', 'قهوة', 'تجارى', 'مشروبات ساخنة', '', ''],
      ['U-4', 'متنوع', 'تجارى', 'متنوع', '', ''],
    ]),
    'تعريف الاصناف'
  );
  XLSX.utils.book_append_sheet(
    unifiedWb,
    XLSX.utils.aoa_to_sheet([['اسم المستخدم', 'الصلاحيات'], ['u1', 'الكل']]),
    'تعريف المستخدمين'
  );
  XLSX.utils.book_append_sheet(
    unifiedWb,
    XLSX.utils.aoa_to_sheet([
      ['Sold-to Party', 'Payer', 'Billing Date', 'Sales Employee Code', 'New Total Tax', 'New Growth Value BD&T'],
      ['0064000101', '0064000001', '2026-11-23', '401', 1100, 110000],
      ['0064000102', '0064000003', '2026-11-23', '403', 500, 50000],
    ]),
    'VF05N (2)'
  );
  XLSX.utils.book_append_sheet(
    unifiedWb,
    XLSX.utils.aoa_to_sheet([
      ['Billing Dt.', 'Customer', 'Material', 'Sales Emp.', 'Qty', 'Total Tax', 'New Growth Value BD&T'],
      // Direct Supply cross-chain sales: both `محمد مجدى` chain (0064000001) & `محمد عبدالمحسن` chain (0064000002) sell both categories!
      [46349, '0064000001', 'U-1', '401', 10, 400, 40000],
      [46349, '0064000002', 'U-1', '402', 5, 100, 10000], // U-1 total = 50,000 sales / 500 tax
      [46349, '0064000001', 'U-3', '401', 10, 200, 20000],
      [46349, '0064000002', 'U-3', '402', 15, 400, 40000], // U-3 total = 60,000 sales / 600 tax
      // Central Supply: all billed on `كمال رضا` chain (`0064000003`), split between `بسكويت + حلويات` (35k) and `مشروبات ساخنة` (15k)
      [46349, '0064000003', 'U-1', '403', 35, 350, 35000],
      [46349, '0064000003', 'U-3', '403', 15, 150, 15000],
    ]),
    'Zsales_Price (2)'
  );
  XLSX.utils.book_append_sheet(
    unifiedWb,
    XLSX.utils.aoa_to_sheet([
      ['اسم القطاع', 'المسمى الوظيفى', 'اسم الموظف', 'الفئة', 'الهدف'],
      ['التوريد المباشر', 'مدير قطاع', 'حازم السيد', 'بسكويت + حلويات', 0],
      ['التوريد المباشر', 'مدير قطاع', 'حازم السيد', 'كيك + كرواسون', 0],
      ['التوريد المباشر', 'مدير قطاع', 'محمد فؤاد', 'مشروبات ساخنة', 0],
      ['التوريد المباشر', 'مدير قطاع', 'محمد فؤاد', 'متنوع', 0],
      ['التوريد المباشر', 'مشرف قطاع', 'محمد مجدى', 'بسكويت + حلويات', 0],
      ['التوريد المباشر', 'مشرف قطاع', 'محمد مجدى', 'كيك + كرواسون', 0],
      ['التوريد المباشر', 'مشرف قطاع', 'محمد عبدالمحسن', 'مشروبات ساخنة', 0],
      ['التوريد المباشر', 'مشرف قطاع', 'محمد عبدالمحسن', 'متنوع', 0],
      ['التوريد المركزي', 'مدير قطاع', 'مصطفى مناع', 'بسكويت + حلويات', 0],
      ['التوريد المركزي', 'مدير قطاع', 'محمد الملا', 'مشروبات ساخنة', 0],
      ['التوريد المركزي', 'مشرف قطاع', 'كمال رضا', 'بسكويت + حلويات', 0],
      ['التوريد المركزي', 'مشرف قطاع', 'أحمد سيد', 'مشروبات ساخنة', 0],
      ['التوريد المركزي', 'مدير إدارة', 'رافت محمد', 'بسكويت + حلويات', 0],
      ['التوريد المركزي', 'مدير إدارة', 'رافت محمد', 'مشروبات ساخنة', 0],
    ]),
    'اهداف الفئات'
  );

  const unifiedReport = validateExcelBuffer(
    XLSX.write(unifiedWb, { type: 'buffer', bookType: 'xlsx' }) as Buffer,
    'Unified_Totals.xlsx'
  );
  assert.equal(unifiedReport.canCommit, true, `Unified workbook errors: ${JSON.stringify(unifiedReport.errors)}`);
  commitValidatedImport(db, unifiedReport.validationToken, { id: 1, fullName: 'مدير النظام' });

  const sumRows = (rows: Array<{ sales: number; tax: number }>) => ({
    sales: Number(rows.reduce((acc, r) => acc + r.sales, 0).toFixed(2)),
    tax: Number(rows.reduce((acc, r) => acc + r.tax, 0).toFixed(2)),
  });

  // Direct Supply: Rep === Supervisor === Sector Manager === (110,000 sales, 1,100 tax)
  for (const jt of ['مندوب مبيعات', 'مشرف قطاع', 'مدير قطاع'] as const) {
    const res = getItemCategoriesAnalytics(db, adminUser, {
      source: 'ZSALES',
      sector: 'التوريد المباشر',
      jobTitle: jt,
    });
    const tot = sumRows(res.rows);
    assert.equal(tot.sales, 110000, `Direct Supply sales under ${jt} must equal exact sector total 110,000`);
    assert.equal(tot.tax, 1100, `Direct Supply tax under ${jt} must equal exact sector total 1,100`);
  }

  // Verify Direct Supply Supervisors keep ONLY their assigned categories & capture 100% of cross-chain sales
  const directSupUnified = getItemCategoriesAnalytics(db, adminUser, {
    source: 'ZSALES',
    sector: 'التوريد المباشر',
    jobTitle: 'مشرف قطاع',
  });
  assert.equal(directSupUnified.rows.find((r) => r.employeeName === 'محمد مجدى' && r.categoryName === 'بسكويت + حلويات')!.sales, 50000);
  assert.equal(directSupUnified.rows.find((r) => r.employeeName === 'محمد مجدى' && r.categoryName === 'كيك + كرواسون')!.sales, 0);
  assert.equal(directSupUnified.rows.find((r) => r.employeeName === 'محمد عبدالمحسن' && r.categoryName === 'مشروبات ساخنة')!.sales, 60000);
  assert.equal(directSupUnified.rows.find((r) => r.employeeName === 'محمد عبدالمحسن' && r.categoryName === 'متنوع')!.sales, 0);

  // Central Supply: Rep === Supervisor === Sector Manager === Department Manager === (50,000 sales, 500 tax)
  for (const jt of ['مندوب مبيعات', 'مشرف قطاع', 'مدير قطاع', 'مدير إدارة'] as const) {
    const res = getItemCategoriesAnalytics(db, adminUser, {
      source: 'ZSALES',
      sector: 'التوريد المركزي',
      jobTitle: jt,
    });
    const tot = sumRows(res.rows);
    assert.equal(tot.sales, 50000, `Central Supply sales under ${jt} must equal exact sector total 50,000`);
    assert.equal(tot.tax, 500, `Central Supply tax under ${jt} must equal exact sector total 500`);
  }
  const centralSupUnified = getItemCategoriesAnalytics(db, adminUser, {
    source: 'ZSALES',
    sector: 'التوريد المركزي',
    jobTitle: 'مشرف قطاع',
  });
  assert.equal(centralSupUnified.rows.find((r) => r.employeeName === 'أحمد سيد' && r.categoryName === 'مشروبات ساخنة')!.sales, 15000);
  assert.equal(centralSupUnified.rows.find((r) => r.employeeName === 'كمال رضا' && r.categoryName === 'بسكويت + حلويات')!.sales, 35000);

  // 20. Verify 3-Role RBAC (admin, sector_manager, standard_user), sector-scoped target writes, Excel role parsing & SQLite migration
  const directMgrUser: UserProfile = {
    id: 10,
    username: 'direct.mgr',
    fullName: 'مدير التوريد المباشر',
    email: null,
    role: 'sector_manager',
    allowedSectors: ['التوريد المباشر'],
    isActive: true,
    createdAt: '2026-11-26T00:00:00Z',
  };
  const directStandardUser: UserProfile = {
    id: 11,
    username: 'direct.viewer',
    fullName: 'مستخدم التوريد المباشر',
    email: null,
    role: 'standard_user',
    allowedSectors: ['التوريد المباشر'],
    isActive: true,
    createdAt: '2026-11-26T00:00:00Z',
  };

  // Both sector_manager and standard_user only see their authorized sector in items analytics
  const mgrItemsView = getItemCategoriesAnalytics(db, directMgrUser, { source: 'ZSALES' });
  assert.ok(mgrItemsView.rows.length > 0);
  assert.ok(mgrItemsView.rows.every((r) => r.sectorName === 'التوريد المباشر'));

  const viewerItemsView = getItemCategoriesAnalytics(db, directStandardUser, { source: 'ZSALES' });
  assert.equal(viewerItemsView.rows.length, mgrItemsView.rows.length);

  // sector_manager CAN save targets inside their authorized sector
  saveItemCategoryTargets(
    db,
    [
      {
        sectorName: 'التوريد المباشر',
        jobTitle: 'مدير قطاع',
        employeeName: 'حازم السيد',
        categoryName: 'بسكويت + حلويات',
        targetValue: 75000,
      },
    ],
    directMgrUser
  );
  const afterMgrSave = getItemCategoriesAnalytics(db, directMgrUser, { source: 'ZSALES', jobTitle: 'مدير قطاع' });
  assert.equal(
    afterMgrSave.rows.find((r) => r.employeeName === 'حازم السيد' && r.categoryName === 'بسكويت + حلويات')?.target,
    75000
  );

  // sector_manager MUST be rejected (403 AuthorizationError) for sectors outside their scope
  assert.throws(
    () =>
      saveItemCategoryTargets(
        db,
        [
          {
            sectorName: 'التوريد المركزي',
            jobTitle: 'مدير قطاع',
            employeeName: 'مصطفى مناع',
            categoryName: 'بسكويت + حلويات',
            targetValue: 99999,
          },
        ],
        directMgrUser
      ),
    AuthorizationError
  );
  assert.throws(
    () =>
      saveItemCategoryTargets(
        db,
        [
          {
            sectorName: 'التوريد المباشر + الجيش + التوب ريتيل',
            jobTitle: 'مدير إدارة',
            employeeName: 'عبدالستار مصطفى',
            categoryName: 'بسكويت + حلويات',
            targetValue: 99999,
          },
        ],
        directMgrUser
      ),
    AuthorizationError
  );

  // standard_user MUST be rejected (403 AuthorizationError) from modifying targets even in their own sector
  assert.throws(
    () =>
      saveItemCategoryTargets(
        db,
        [
          {
            sectorName: 'التوريد المباشر',
            jobTitle: 'مدير قطاع',
            employeeName: 'حازم السيد',
            categoryName: 'بسكويت + حلويات',
            targetValue: 12345,
          },
        ],
        directStandardUser
      ),
    AuthorizationError
  );

  // Competition targets RBAC check
  const rbacComp = createCompetition(db, { id: 1 }, { name: 'مسابقة الصلاحيات', items: [{ itemCode: 'U-1', equivalentFactor: 1 }] });
  saveCompetitionTargets(db, rbacComp.id, [{ participantKey: 'REP::التوريد المباشر::401', targetQty: 150 }], directMgrUser);
  assert.throws(
    () =>
      saveCompetitionTargets(
        db,
        rbacComp.id,
        [{ participantKey: 'REP::التوريد المركزي::403', targetQty: 200 }],
        directMgrUser
      ),
    AuthorizationError
  );
  assert.throws(
    () =>
      saveCompetitionTargets(
        db,
        rbacComp.id,
        [{ participantKey: 'REP::التوريد المباشر::401', targetQty: 300 }],
        directStandardUser
      ),
    AuthorizationError
  );

  // Excel users sheet parses 3 roles accurately
  const rolesWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(rolesWb, XLSX.utils.aoa_to_sheet([['كود القطاع', 'القطاع'], ['S1', 'التوريد المباشر']]), 'القطاعات');
  XLSX.utils.book_append_sheet(rolesWb, XLSX.utils.aoa_to_sheet([['القطاع', 'الهدف'], ['التوريد المباشر', 1000]]), 'اهداف القطاعات');
  XLSX.utils.book_append_sheet(rolesWb, XLSX.utils.aoa_to_sheet([['كود السلسلة', 'اسم السلسلة', 'الادارة', 'القطاع'], ['101', 'عميل', 'إدارة', 'التوريد المباشر']]), 'تعريف الادارات والقطاعات');
  XLSX.utils.book_append_sheet(rolesWb, XLSX.utils.aoa_to_sheet([['كود الفرع', 'اسم الفرع', 'القطاع', 'الهدف'], ['501', 'فرع', 'التوريد المباشر', 1000]]), 'تعريف الفروع و الاهداف');
  XLSX.utils.book_append_sheet(rolesWb, XLSX.utils.aoa_to_sheet([['كود المندوب', 'اسم المندوب', 'القطاع'], ['R1', 'مندوب', 'التوريد المباشر']]), 'تعريف مندوبى البيع');
  XLSX.utils.book_append_sheet(rolesWb, XLSX.utils.aoa_to_sheet([['كود الصنف', 'اسم الصنف'], ['M1', 'صنف']]), 'تعريف الاصناف');
  XLSX.utils.book_append_sheet(
    rolesWb,
    XLSX.utils.aoa_to_sheet([
      ['اسم المستخدم', 'الاسم الكامل', 'الدور', 'الصلاحيات'],
      ['sys.admin', 'مدير النظام', 'أدمن', 'الكل'],
      ['sec.mgr', 'مدير المباشر', 'مدير قطاع', 'التوريد المباشر'],
      ['std.view', 'مستخدم المباشر', 'مستخدم', 'التوريد المباشر'],
    ]),
    'تعريف المستخدمين'
  );
  XLSX.utils.book_append_sheet(rolesWb, XLSX.utils.aoa_to_sheet([['Sold-to Party', 'Payer', 'Billing Date', 'Sales Employee Code', 'Material', 'New Total Tax', 'New Growth Value BD&T'], ['501', '101', '2026-11-15', 'R1', 'M1', 14, 100]]), 'VF05N (2)');
  XLSX.utils.book_append_sheet(rolesWb, XLSX.utils.aoa_to_sheet([['Billing Dt.', 'Customer', 'Material', 'Sales Emp.', 'Qty', 'Total Tax', 'New Growth Value BD&T'], [46341, '101', 'M1', 'R1', 1, 14, 100]]), 'Zsales_Price (2)');
  const rolesReport = validateExcelBuffer(XLSX.write(rolesWb, { type: 'buffer', bookType: 'xlsx' }) as Buffer, 'Roles_Test.xlsx');
  assert.equal(rolesReport.usersFound.find((u) => u.username === 'sys.admin')?.role, 'admin');
  assert.equal(rolesReport.usersFound.find((u) => u.username === 'sec.mgr')?.role, 'sector_manager');
  assert.equal(rolesReport.usersFound.find((u) => u.username === 'std.view')?.role, 'standard_user');

  // Legacy SQLite database with CHECK(role IN ('admin', 'sector_user')) migrates cleanly via initDatabase
  const tmpDbPath = path.join(os.tmpdir(), `rbac_migrate_${Date.now()}.db`);
  try {
    const legacyDb = new Database(tmpDbPath);
    legacyDb.exec(`
      CREATE TABLE users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        full_name TEXT NOT NULL,
        email TEXT,
        role TEXT NOT NULL CHECK(role IN ('admin', 'sector_user')),
        allowed_sectors TEXT NOT NULL DEFAULT '[]',
        password_hash TEXT NOT NULL,
        must_change_password INTEGER NOT NULL DEFAULT 0,
        is_active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL
      );
      INSERT INTO users (username, full_name, role, allowed_sectors, password_hash, created_at)
      VALUES
        ('old.admin', 'Old Admin', 'admin', '[]', 'salt:hash', '2026-11-01T00:00:00Z'),
        ('old.user', 'Old Sector User', 'sector_user', '["التوريد المباشر"]', 'salt:hash', '2026-11-01T00:00:00Z');
    `);
    legacyDb.close();

    const migratedDb = initDatabase(tmpDbPath);
    const oldAdminRow = migratedDb.prepare(`SELECT role FROM users WHERE username = 'old.admin'`).get() as { role: string };
    const oldUserRow = migratedDb.prepare(`SELECT role FROM users WHERE username = 'old.user'`).get() as { role: string };
    assert.equal(oldAdminRow.role, 'admin');
    assert.equal(oldUserRow.role, 'standard_user');
    migratedDb
      .prepare(
        `INSERT INTO users (username, full_name, role, allowed_sectors, password_hash, created_at)
         VALUES ('new.mgr', 'Sector Manager', 'sector_manager', '["التوريد المباشر"]', 'salt:hash', '2026-11-26T00:00:00Z')`
      )
      .run();
    migratedDb.close();
  } finally {
    if (fs.existsSync(tmpDbPath)) fs.unlinkSync(tmpDbPath);
  }

  // =========================================================================
  // 21. Competitions Class A Split (`صناعى` / `تجارى`), Sector-Manager Category
  //     Isolation in المباشر/المركزي vs Full-Sector in الجيش/التوب ريتيل,
  //     Dual Targets Persistence, & Department Manager Rollup
  // =========================================================================
  assert.equal(classifyClassA('صناعى'), 'industrial');
  assert.equal(classifyClassA('صناعي'), 'industrial');
  assert.equal(classifyClassA('أصناف صناعية'), 'industrial');
  assert.equal(classifyClassA('تجارى'), 'commercial');
  assert.equal(classifyClassA('تجاري'), 'commercial');
  assert.equal(classifyClassA('أصناف تجارية'), 'commercial');
  assert.equal(classifyClassA(''), 'other');

  const classAWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    classAWb,
    XLSX.utils.aoa_to_sheet([
      ['كود القطاع', 'القطاع'],
      ['S1', 'التوريد المباشر'],
      ['S2', 'التوريد المركزي'],
      ['S3', 'الجيش'],
      ['S4', 'التوب ريتيل'],
    ]),
    'القطاعات'
  );
  XLSX.utils.book_append_sheet(
    classAWb,
    XLSX.utils.aoa_to_sheet([
      ['القطاع', 'الهدف'],
      ['التوريد المباشر', 1000],
      ['التوريد المركزي', 1000],
      ['الجيش', 1000],
      ['التوب ريتيل', 1000],
    ]),
    'اهداف القطاعات'
  );
  XLSX.utils.book_append_sheet(
    classAWb,
    XLSX.utils.aoa_to_sheet([
      ['كود السلسلة', 'اسم السلسلة', 'الادارة', 'اسم المشرف', 'القطاع'],
      ['C1', 'سلسلة المباشر', 'ادارة المباشر', 'مشرف مباشر', 'التوريد المباشر'],
      ['C2', 'سلسلة المركزي', 'ادارة المركزي', 'مشرف مركزي', 'التوريد المركزي'],
      ['C3', 'سلسلة الجيش', 'ادارة الجيش', 'مشرف جيش', 'الجيش'],
      ['C4', 'سلسلة التوب', 'ادارة التوب', 'مشرف توب', 'التوب ريتيل'],
    ]),
    'تعريف الادارات والقطاعات'
  );
  XLSX.utils.book_append_sheet(
    classAWb,
    XLSX.utils.aoa_to_sheet([
      ['كود الفرع', 'اسم الفرع', 'الهدف', 'القطاع'],
      ['B1', 'فرع مباشر', 1000, 'التوريد المباشر'],
      ['B2', 'فرع مركزي', 1000, 'التوريد المركزي'],
      ['B3', 'فرع جيش', 1000, 'الجيش'],
      ['B4', 'فرع توب', 1000, 'التوب ريتيل'],
    ]),
    'تعريف الفروع و الاهداف'
  );
  XLSX.utils.book_append_sheet(
    classAWb,
    XLSX.utils.aoa_to_sheet([
      ['كود المندوب', 'اسم المندوب', 'القطاع', 'اسم المشرف', 'الهدف'],
      ['R-D', 'مندوب مباشر', 'التوريد المباشر', 'مشرف مباشر', 0],
      ['R-C', 'مندوب مركزي', 'التوريد المركزي', 'مشرف مركزي', 0],
      ['R-A', 'مندوب جيش', 'الجيش', 'مشرف جيش', 0],
      ['R-T', 'مندوب توب', 'التوب ريتيل', 'مشرف توب', 0],
    ]),
    'تعريف مندوبى البيع'
  );
  XLSX.utils.book_append_sheet(
    classAWb,
    XLSX.utils.aoa_to_sheet([
      ['كود الصنف', 'الصنف', 'التصنيف A', 'تصنيف السلاسل', 'تصنيف الجيش', 'التوب ريتيل'],
      ['IND-1', 'دقيق صناعى 50ك', 'صناعى', 'بسكويت + حلويات', 'جيش صناعى', 'توب صناعى'],
      ['COM-1', 'بسكويت تجارى', 'تجاري', 'بسكويت + حلويات', 'جيش تجارى', 'توب تجارى'],
      ['IND-2', 'خامات كيك صناعى', 'صناعي', 'كيك + كرواسون', 'جيش صناعى', 'توب صناعى'],
      ['COM-2', 'قهوة تجارى', 'تجارى', 'مشروبات ساخنة', 'جيش تجارى', 'توب تجارى'],
    ]),
    'تعريف الاصناف'
  );
  XLSX.utils.book_append_sheet(
    classAWb,
    XLSX.utils.aoa_to_sheet([['اسم المستخدم', 'الصلاحيات'], ['u1', 'الكل']]),
    'تعريف المستخدمين'
  );
  XLSX.utils.book_append_sheet(
    classAWb,
    XLSX.utils.aoa_to_sheet([
      ['Sold-to Party', 'Payer', 'Billing Date', 'Sales Employee Code', 'New Total Tax', 'New Growth Value BD&T'],
      ['B1', 'C1', '2026-11-22', 'R-D', 100, 1000],
    ]),
    'VF05N (2)'
  );
  XLSX.utils.book_append_sheet(
    classAWb,
    XLSX.utils.aoa_to_sheet([
      ['Billing Dt.', 'Customer', 'Material', 'Sales Emp.', 'Qty', 'Total Tax', 'New Growth Value BD&T'],
      // Direct supply: IND-1 (بسكويت + حلويات, 10 * 0.5 = 5), COM-1 (بسكويت + حلويات, 12 * 1 = 12), IND-2 (كيك + كرواسون, 6 * 1 = 6)
      [46348, 'C1', 'IND-1', 'R-D', 10, 0, 100],
      [46348, 'C1', 'COM-1', 'R-D', 12, 0, 120],
      [46348, 'C1', 'IND-2', 'R-D', 6, 0, 60],
      // Central supply: COM-2 (مشروبات ساخنة, 20 * 0.25 = 5), IND-1 (بسكويت + حلويات, 4 * 0.5 = 2)
      [46348, 'C2', 'COM-2', 'R-C', 20, 0, 200],
      [46348, 'C2', 'IND-1', 'R-C', 4, 0, 40],
      // Army: IND-1 (8 * 0.5 = 4), COM-2 (10 * 0.25 = 2.5)
      [46348, 'C3', 'IND-1', 'R-A', 8, 0, 80],
      [46348, 'C3', 'COM-2', 'R-A', 10, 0, 100],
      // Top Retail: IND-2 (3 * 1 = 3), COM-1 (9 * 1 = 9)
      [46348, 'C4', 'IND-2', 'R-T', 3, 0, 30],
      [46348, 'C4', 'COM-1', 'R-T', 9, 0, 90],
    ]),
    'Zsales_Price (2)'
  );
  XLSX.utils.book_append_sheet(
    classAWb,
    XLSX.utils.aoa_to_sheet([
      ['اسم القطاع', 'المسمى الوظيفى', 'اسم الموظف', 'الفئة', 'الهدف'],
      ['التوريد المباشر', 'مدير قطاع', 'حازم السيد', 'بسكويت + حلويات', 0],
      ['التوريد المباشر', 'مدير قطاع', 'محمد فؤاد', 'كيك + كرواسون', 0],
      ['التوريد المركزي', 'مدير قطاع', 'مصطفى مناع', 'بسكويت + حلويات', 0],
      ['التوريد المركزي', 'مدير قطاع', 'محمد الملا', 'مشروبات ساخنة', 0],
      // Army & Top Retail sector managers MUST see full sector even if only one category row appears in category_targets!
      ['الجيش', 'مدير قطاع', 'قائد الجيش', 'بسكويت + حلويات', 0],
      ['التوب ريتيل', 'مدير قطاع', 'مدير التوب', 'بسكويت + حلويات', 0],
      ['التوريد المركزي', 'مدير إدارة', 'رأفت محمد', 'بسكويت + حلويات', 0],
      ['التوريد المباشر + الجيش + التوب ريتيل', 'مدير ادارة', 'عبد الستار مصطفي', 'بسكويت + حلويات', 0],
    ]),
    'اهداف الفئات'
  );

  const classAReport = validateExcelBuffer(
    XLSX.write(classAWb, { type: 'buffer', bookType: 'xlsx' }) as Buffer,
    'ClassA_Competitions.xlsx'
  );
  assert.equal(classAReport.canCommit, true);
  commitValidatedImport(db, classAReport.validationToken, { id: 1, fullName: 'مدير النظام' });

  const classAComp = createCompetition(
    db,
    { id: 1 },
    {
      name: 'مسابقة التصنيف A',
      items: [
        { itemCode: 'IND-1', itemName: 'دقيق صناعى 50ك', equivalentFactor: 0.5 },
        { itemCode: 'COM-1', itemName: 'بسكويت تجارى', equivalentFactor: 1 },
        { itemCode: 'IND-2', itemName: 'خامات كيك صناعى', equivalentFactor: 1 },
        { itemCode: 'COM-2', itemName: 'قهوة تجارى', equivalentFactor: 0.25 },
      ],
    }
  );

  const caRes1 = getCompetitionAnalytics(db, adminUser, classAComp.id, { sector: 'ALL' });

  // Sales Reps: full industrial & commercial split
  const repD = caRes1.participants.find((p) => p.roleKind === 'sales_rep' && p.employeeCode === 'R-D')!;
  assert.equal(repD.industrialSoldQty, 11); // IND-1 (5) + IND-2 (6)
  assert.equal(repD.commercialSoldQty, 12); // COM-1 (12)
  assert.equal(repD.industrialAchievementPct, null); // 0 target -> null, never NaN
  assert.equal(repD.commercialAchievementPct, null);

  const repC = caRes1.participants.find((p) => p.roleKind === 'sales_rep' && p.employeeCode === 'R-C')!;
  assert.equal(repC.industrialSoldQty, 2); // IND-1 (2)
  assert.equal(repC.commercialSoldQty, 5); // COM-2 (5)

  // Sector Managers in التوريد المباشر: scoped strictly to their assigned categories!
  const hazem = caRes1.participants.find((p) => p.roleKind === 'sector_manager' && p.employeeName === 'حازم السيد')!;
  const fouad = caRes1.participants.find((p) => p.roleKind === 'sector_manager' && p.employeeName === 'محمد فؤاد')!;
  assert.deepEqual(hazem.assignedCategories, ['بسكويت + حلويات']);
  assert.equal(hazem.industrialSoldQty, 5); // ONLY IND-1 from بسكويت + حلويات
  assert.equal(hazem.commercialSoldQty, 12); // ONLY COM-1 from بسكويت + حلويات
  assert.deepEqual(fouad.assignedCategories, ['كيك + كرواسون']);
  assert.equal(fouad.industrialSoldQty, 6); // ONLY IND-2 from كيك + كرواسون
  assert.equal(fouad.commercialSoldQty, 0);

  // Sector Managers in التوريد المركزي: scoped strictly to their assigned categories!
  const mostafaManna = caRes1.participants.find((p) => p.roleKind === 'sector_manager' && p.employeeName === 'مصطفى مناع')!;
  const mulla = caRes1.participants.find((p) => p.roleKind === 'sector_manager' && p.employeeName === 'محمد الملا')!;
  assert.equal(mostafaManna.industrialSoldQty, 2); // IND-1
  assert.equal(mostafaManna.commercialSoldQty, 0);
  assert.equal(mulla.industrialSoldQty, 0);
  assert.equal(mulla.commercialSoldQty, 5); // COM-2

  // Sector Managers in الجيش & التوب ريتيل: full sector (NOT filtered down to a single category)!
  const armySecMgrCA = caRes1.participants.find((p) => p.roleKind === 'sector_manager' && p.employeeName === 'قائد الجيش')!;
  const topSecMgrCA = caRes1.participants.find((p) => p.roleKind === 'sector_manager' && p.employeeName === 'مدير التوب')!;
  assert.equal(armySecMgrCA.industrialSoldQty, 4); // Full Army industrial
  assert.equal(armySecMgrCA.commercialSoldQty, 2.5); // Full Army commercial (includes COM-2!)
  assert.equal(topSecMgrCA.industrialSoldQty, 3); // Full Top Retail industrial (includes IND-2!)
  assert.equal(topSecMgrCA.commercialSoldQty, 9); // Full Top Retail commercial

  // Department Managers: full scope split into industrial & commercial
  const abdelCA = caRes1.participants.find((p) => p.roleKind === 'department_manager' && p.employeeName.includes('عبد'))!;
  const rafatCA = caRes1.participants.find(
    (p) => p.roleKind === 'department_manager' && (p.employeeName.includes('رافت') || p.employeeName.includes('رأفت'))
  )!;
  assert.equal(abdelCA.industrialSoldQty, 18); // 11 Direct + 4 Army + 3 TopRetail
  assert.equal(abdelCA.commercialSoldQty, 23.5); // 12 Direct + 2.5 Army + 9 TopRetail
  assert.equal(rafatCA.industrialSoldQty, 2);
  assert.equal(rafatCA.commercialSoldQty, 5);

  // Save dual Industrial + Commercial targets directly -> check DB persistence, % calc, & مدير إدارة rollup!
  saveCompetitionTargets(db, classAComp.id, [
    { participantKey: hazem.participantKey, industrialTargetQty: 10, commercialTargetQty: 24 },
    { participantKey: fouad.participantKey, industrialTargetQty: 12, commercialTargetQty: 0 },
    { participantKey: armySecMgrCA.participantKey, industrialTargetQty: 8, commercialTargetQty: 5 },
    { participantKey: topSecMgrCA.participantKey, industrialTargetQty: 6, commercialTargetQty: 18 },
    { participantKey: repC.participantKey, industrialTargetQty: 4, commercialTargetQty: 10 },
  ]);

  const caResAfter = getCompetitionAnalytics(db, adminUser, classAComp.id, { sector: 'ALL' });
  const hazemAfter = caResAfter.participants.find((p) => p.participantKey === hazem.participantKey)!;
  assert.equal(hazemAfter.industrialTargetQty, 10);
  assert.equal(hazemAfter.commercialTargetQty, 24);
  assert.equal(hazemAfter.industrialAchievementPct, 50);
  assert.equal(hazemAfter.commercialAchievementPct, 50);

  const fouadAfter = caResAfter.participants.find((p) => p.participantKey === fouad.participantKey)!;
  assert.equal(fouadAfter.industrialAchievementPct, 50); // 6 / 12
  assert.equal(fouadAfter.commercialAchievementPct, null); // 0 target -> null (division-by-zero protected)

  const abdelAfter = caResAfter.participants.find((p) => p.participantKey === abdelCA.participantKey)!;
  assert.equal(abdelAfter.industrialTargetQty, 36); // Direct (10+12) + Army (8) + Top (6)
  assert.equal(abdelAfter.commercialTargetQty, 47); // Direct (24+0) + Army (5) + Top (18)
  assert.equal(abdelAfter.industrialAchievementPct, 50); // 18 / 36
  assert.equal(abdelAfter.commercialAchievementPct, 50); // 23.5 / 47

  const rafatAfter = caResAfter.participants.find((p) => p.participantKey === rafatCA.participantKey)!;
  assert.equal(rafatAfter.industrialTargetQty, 4); // repC industrial target
  assert.equal(rafatAfter.commercialTargetQty, 10); // repC commercial target
  assert.equal(rafatAfter.industrialAchievementPct, 50); // 2 / 4
  assert.equal(rafatAfter.commercialAchievementPct, 50); // 5 / 10

  // Verify summary totals across visible sectors
  assert.equal(caResAfter.summary.totalIndustrialSoldQty, 20); // 18 + 2
  assert.equal(caResAfter.summary.totalCommercialSoldQty, 28.5); // 23.5 + 5
  assert.equal(caResAfter.summary.totalIndustrialTargetQty, 40); // 36 + 4
  assert.equal(caResAfter.summary.totalCommercialTargetQty, 57); // 47 + 10
  assert.equal(caResAfter.summary.industrialAchievementPct, 50);
  assert.equal(caResAfter.summary.commercialAchievementPct, 50);

  // Verify dual competition targets survive a new Excel import
  const classAReimport = validateExcelBuffer(
    XLSX.write(classAWb, { type: 'buffer', bookType: 'xlsx' }) as Buffer,
    'ClassA_Competitions_Reimport.xlsx'
  );
  commitValidatedImport(db, classAReimport.validationToken, { id: 1, fullName: 'مدير النظام' });
  const caResReloaded = getCompetitionAnalytics(db, adminUser, classAComp.id, { sector: 'ALL' });
  assert.equal(
    caResReloaded.participants.find((p) => p.participantKey === hazem.participantKey)!.industrialTargetQty,
    10
  );
  assert.equal(
    caResReloaded.participants.find((p) => p.participantKey === hazem.participantKey)!.commercialTargetQty,
    24
  );

  // =========================================================================
  // 22. Smart Auditor (`المراجع الذكي`): Runtime Error Monitor, System Audit,
  //     Safe Auto-Fix with Backup Snapshot, and Exact Rollback Restoration
  // =========================================================================
  // Ensure at least one active admin exists in `users` for RBAC check
  db.prepare(
    `INSERT OR IGNORE INTO users (username, full_name, role, allowed_sectors, password_hash, must_change_password, is_active, created_at)
     VALUES ('smart_admin', 'مدير المراجعة', 'admin', ?, 'hash', 0, 1, ?)`
  ).run(JSON.stringify(['التوريد المباشر', 'التوريد المركزي', 'التوب ريتيل', 'الجيش']), new Date().toISOString());

  // 22a. Clean audit pass check (no runtime errors, no numeric corruption, RBAC probes pass)
  const cleanAudit = runSmartAudit(db, adminUser);
  assert.ok(cleanAudit.checksExecuted.find((c) => c.key === 'numeric_integrity')!.passed);
  assert.ok(cleanAudit.checksExecuted.find((c) => c.key === 'rbac_permissions')!.passed);
  assert.ok(cleanAudit.checksExecuted.find((c) => c.key === 'runtime_errors')!.passed);

  // 22b. Runtime error capture & 60s deduplication + readyPrompt generation
  const errLog1 = recordRuntimeError(
    db,
    { id: 1, username: 'smart_admin' },
    {
      kind: 'runtime_error',
      tabName: 'مسابقات',
      message: 'Maximum update depth exceeded inside CompetitionsView',
      stack: 'Error: Maximum update depth exceeded\n  at CompetitionsView (CompetitionsView.tsx:120)',
    }
  );
  const errLogDup = recordRuntimeError(
    db,
    { id: 1, username: 'smart_admin' },
    {
      kind: 'runtime_error',
      tabName: 'مسابقات',
      message: 'Maximum update depth exceeded inside CompetitionsView',
      stack: 'Error: Maximum update depth exceeded\n  at CompetitionsView (CompetitionsView.tsx:120)',
    }
  );
  assert.equal(errLogDup.id, errLog1.id, 'Identical errors within 60s must deduplicate');

  const auditWithErr = runSmartAudit(db, adminUser);
  const runtimeFinding = auditWithErr.findings.find((f) => f.id === `RUNTIME_ERR_${errLog1.id}`)!;
  assert.ok(runtimeFinding, 'Runtime error finding must surface in Admin Smart Audit');
  assert.equal(runtimeFinding.severity, 'critical');
  assert.ok(runtimeFinding.readyPrompt && runtimeFinding.readyPrompt.includes('مسابقات'));

  // 22c. Inject negative/non-finite target, duplicate branch row, and competition total desync -> detect -> autofix -> rollback
  db.prepare(`INSERT INTO item_category_targets (target_key, sector_name, job_title, employee_name, category_name, target_value, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?)`)
    .run('bad::key::1', 'التوريد المباشر', 'مندوب مبيعات', 'اختبار', 'بسكويت + حلويات', -500, new Date().toISOString());

  const activeBatchRow = db.prepare(`SELECT id FROM import_batches WHERE status = 'active' ORDER BY id DESC LIMIT 1`).get() as { id: number };
  const sampleBranch = db.prepare(`SELECT * FROM branches WHERE batch_id = ? LIMIT 1`).get(activeBatchRow.id) as {
    batch_id: number;
    branch_code: string;
    branch_name: string;
    sector_name: string;
    branch_target: number;
    chain_code: string | null;
    chain_name: string;
    rep_code: string | null;
    rep_name: string;
    supervisor_name: string;
  };
  db.prepare(
    `INSERT INTO branches (batch_id, branch_code, branch_name, sector_name, branch_target, chain_code, chain_name, rep_code, rep_name, supervisor_name)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
  ).run(
    sampleBranch.batch_id,
    sampleBranch.branch_code,
    sampleBranch.branch_name,
    sampleBranch.sector_name,
    250,
    sampleBranch.chain_code,
    sampleBranch.chain_name,
    sampleBranch.rep_code,
    sampleBranch.rep_name,
    sampleBranch.supervisor_name
  );
  const countBranchesBeforeFix = (db.prepare(`SELECT COUNT(*) AS cnt FROM branches WHERE batch_id = ?`).get(activeBatchRow.id) as { cnt: number }).cnt;

  db.prepare(`UPDATE competition_targets SET target_qty = 999 WHERE competition_id = ? AND industrial_target_qty > 0`).run(classAComp.id);

  const dirtyAudit = runSmartAudit(db, adminUser);
  assert.ok(dirtyAudit.findings.some((f) => f.code === 'NUM_DB_NON_FINITE'), 'Must detect negative/non-finite DB numeric values');
  assert.ok(dirtyAudit.findings.some((f) => f.code === 'DUP_BRANCH_ROWS'), 'Must detect duplicate branch codes');
  assert.ok(dirtyAudit.findings.some((f) => f.code === 'COMP_TARGET_TOTAL_DESYNC'), 'Must detect desynced competition target totals');

  const fixNumRes = executeAutoFix(db, adminUser, 'FIX_NON_FINITE_NUMBERS', 'NUM_DB_NON_FINITE');
  assert.ok(fixNumRes.backupId > 0);
  const fixedTargetRow = db.prepare(`SELECT target_value FROM item_category_targets WHERE target_key = 'bad::key::1'`).get() as { target_value: number };
  assert.equal(fixedTargetRow.target_value, 0, 'Negative target must be sanitized to 0');

  const fixDupBrRes = executeAutoFix(db, adminUser, 'FIX_DUPLICATE_BRANCH_RECORDS', 'DUP_BRANCH_ROWS');
  const countBranchesAfterFix = (db.prepare(`SELECT COUNT(*) AS cnt FROM branches WHERE batch_id = ?`).get(activeBatchRow.id) as { cnt: number }).cnt;
  assert.equal(countBranchesAfterFix, countBranchesBeforeFix - 1, 'Duplicate branch row must be merged & removed');
  assert.ok(!fixDupBrRes.report.findings.some((f) => f.code === 'DUP_BRANCH_ROWS'));

  const fixCompSync = executeAutoFix(db, adminUser, 'FIX_SYNC_COMPETITION_TARGET_TOTALS', 'COMP_TARGET_TOTAL_DESYNC');
  assert.ok(!fixCompSync.report.findings.some((f) => f.code === 'COMP_TARGET_TOTAL_DESYNC'));

  const fixErrArchive = executeAutoFix(db, adminUser, 'FIX_RESOLVE_RUNTIME_ERRORS', 'RUNTIME_ERR_ARCHIVE');
  assert.equal(fixErrArchive.report.runtimeErrors.length, 0);

  // Verify Rollback restores exact pre-fix state
  const rollbackNum = rollbackAutoFix(db, adminUser, fixNumRes.backupId);
  const restoredTargetRow = db.prepare(`SELECT target_value FROM item_category_targets WHERE target_key = 'bad::key::1'`).get() as { target_value: number };
  assert.equal(restoredTargetRow.target_value, -500, 'Rollback must restore exact pre-fix numeric snapshot');
  assert.ok(rollbackNum.report.backups.find((b) => b.id === fixNumRes.backupId)!.restoredAt !== null);

  const rollbackDupBr = rollbackAutoFix(db, adminUser, fixDupBrRes.backupId);
  const countBranchesAfterRollback = (db.prepare(`SELECT COUNT(*) AS cnt FROM branches WHERE batch_id = ?`).get(activeBatchRow.id) as { cnt: number }).cnt;
  assert.equal(countBranchesAfterRollback, countBranchesBeforeFix, 'Rollback must restore exact pre-fix branch rows');
  assert.ok(rollbackDupBr.report.findings.some((f) => f.code === 'DUP_BRANCH_ROWS'));

  console.log('All Sales BI Engine tests PASSED (22/22 assertion groups verified).');
}

runTests();
