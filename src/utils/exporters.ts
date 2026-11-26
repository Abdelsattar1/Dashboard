import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import html2canvas from 'html2canvas';
import cairoArabic400Url from '@fontsource/cairo/files/cairo-arabic-400-normal.woff2?url';
import cairoArabic700Url from '@fontsource/cairo/files/cairo-arabic-700-normal.woff2?url';
import cairoLatin400Url from '@fontsource/cairo/files/cairo-latin-400-normal.woff2?url';
import cairoLatin700Url from '@fontsource/cairo/files/cairo-latin-700-normal.woff2?url';

export type ExportColumnType = 'string' | 'currency' | 'number' | 'percent' | 'qty' | 'integer';

export interface ExportColumn {
  header: string;
  key: string;
  type: ExportColumnType;
  align?: 'right' | 'center' | 'left';
  minWidth?: number;
}

export type ExportRowData = Record<string, string | number | null | undefined>;

export interface ExcelSummaryPair {
  label: string;
  value: string | number | null | undefined;
  type?: ExportColumnType;
}

export interface ExcelSheetSpec {
  name: string;
  title?: string;
  subtitle?: string;
  summaryPairs?: ExcelSummaryPair[];
  columns: ExportColumn[];
  rows: ExportRowData[];
  totalRow?: ExportRowData | null;
}

export interface ExportToExcelParams {
  filename: string;
  sheets: ExcelSheetSpec[];
}

export interface PdfKpiItem {
  label: string;
  value: string;
  unit?: string;
  sublabel?: string;
  tone?: 'default' | 'positive' | 'negative' | 'accent';
}

export interface PdfChartLegendItem {
  label: string;
  color: string;
  valueText?: string;
}

export interface PdfChartItem {
  title: string;
  subtitle?: string;
  imageDataUrl: string;
  width?: number;
  height?: number;
  legend?: PdfChartLegendItem[];
}

export interface PdfTableRow extends ExportRowData {
  _tone?: 'subtotal' | 'private' | 'direct_army' | 'chains' | 'highlight';
}

export interface PdfTableSection {
  title?: string;
  subtitle?: string;
  columns: ExportColumn[];
  rows: PdfTableRow[];
  totalRow?: ExportRowData | null;
}

export interface ExportToPdfParams {
  filename: string;
  tabTitle: string;
  filtersSummary: string[];
  orientation?: 'auto' | 'portrait' | 'landscape';
  kpis?: PdfKpiItem[];
  infoBlocks?: Array<{ title: string; items: string[] }>;
  tables: PdfTableSection[];
  charts?: PdfChartItem[];
  chartsSectionTitle?: string;
}

function sanitizeFilenameSegment(raw: string): string {
  return String(raw || '')
    .replace(/[\\/:*?"<>|]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/\s+/g, '_');
}

export function getExportDateString(date = new Date()): string {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

export function getExportDateTimeDisplay(date = new Date()): string {
  const datePart = getExportDateString(date);
  const hh = String(date.getHours()).padStart(2, '0');
  const min = String(date.getMinutes()).padStart(2, '0');
  return `${datePart} — ${hh}:${min}`;
}

export function buildExportFilename(
  tabName: string,
  activeFilterLabels: Array<string | null | undefined>,
  ext: 'xlsx' | 'pdf'
): string {
  const parts: string[] = ['تحليلات_قطاع_السلاسل', sanitizeFilenameSegment(tabName)];
  for (const label of activeFilterLabels) {
    if (!label) continue;
    const clean = sanitizeFilenameSegment(label);
    if (clean) parts.push(clean);
  }
  parts.push(getExportDateString());

  let base = parts.filter(Boolean).join('_');
  if (base.length > 150) {
    base = base.slice(0, 140) + '_' + getExportDateString();
  }
  return `${base}.${ext}`;
}

export function buildActiveFilterMeta(opts: {
  source?: string;
  sector?: string;
  startDate?: string;
  endDate?: string;
  extra?: Array<{ label: string; value?: string | null; fileTag?: string | null; includeWhenAll?: boolean }>;
}): { summary: string[]; fileTags: string[] } {
  const summary: string[] = [];
  const fileTags: string[] = [];

  if (opts.source) {
    const srcName = opts.source === 'VF05N' ? 'VF05N (2)' : opts.source === 'ZSALES' ? 'Zsales_Price (2)' : opts.source;
    summary.push(`المصدر: ${srcName}`);
  }

  if (opts.sector) {
    if (opts.sector === 'ALL') {
      summary.push('القطاع: جميع القطاعات');
    } else if (opts.sector === 'UNLINKED') {
      summary.push('القطاع: سجلات غير مرتبطة بقطاع');
      fileTags.push('غير_مرتبط_بقطاع');
    } else {
      summary.push(`القطاع: ${opts.sector}`);
      fileTags.push(opts.sector);
    }
  }

  if (opts.startDate || opts.endDate) {
    const rangeStr = `${opts.startDate || '...'} إلى ${opts.endDate || '...'}`;
    summary.push(`الفترة: ${rangeStr}`);
    fileTags.push(`${opts.startDate || 'بداية'}_${opts.endDate || 'نهاية'}`);
  }

  if (opts.extra) {
    for (const item of opts.extra) {
      if (item.value) {
        summary.push(`${item.label}: ${item.value}`);
      }
      if (item.fileTag) {
        fileTags.push(item.fileTag);
      }
    }
  }

  return { summary, fileTags };
}

function sanitizeSheetName(raw: string, used: Set<string>): string {
  let clean = String(raw || 'تقرير')
    .replace(/[\\/?*[\]:]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 31);
  if (!clean) clean = 'تقرير';
  let candidate = clean;
  let idx = 2;
  while (used.has(candidate)) {
    const suffix = ` (${idx++})`;
    candidate = `${clean.slice(0, Math.max(1, 31 - suffix.length))}${suffix}`;
  }
  used.add(candidate);
  return candidate;
}

function toExcelCell(value: string | number | null | undefined, type: ExportColumnType): { cell: XLSX.CellObject; displayLen: number } {
  if (type === 'string') {
    const str = value === null || value === undefined ? '' : String(value);
    return {
      cell: { t: 's', v: str },
      displayLen: str.length,
    };
  }

  if (value === null || value === undefined || value === '') {
    return { cell: { t: 's', v: '—' }, displayLen: 2 };
  }

  const num = Number(value);
  if (!Number.isFinite(num)) {
    const str = String(value);
    return { cell: { t: 's', v: str }, displayLen: str.length };
  }

  if (type === 'percent') {
    const pctDecimal = Number((num / 100).toFixed(4));
    const formatted = `${num.toFixed(1)}%`;
    return {
      cell: { t: 'n', v: pctDecimal, z: '0.0%' },
      displayLen: formatted.length + 1,
    };
  }

  if (type === 'currency') {
    const rounded = Number(num.toFixed(2));
    const formatted = rounded.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    return {
      cell: { t: 'n', v: rounded, z: '#,##0.00;-#,##0.00;0.00' },
      displayLen: formatted.length + 1,
    };
  }

  if (type === 'qty') {
    const rounded = Number(num.toFixed(3));
    const formatted = rounded.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 3 });
    return {
      cell: { t: 'n', v: rounded, z: '#,##0.###;-#,##0.###;0' },
      displayLen: formatted.length + 1,
    };
  }

  if (type === 'integer') {
    const rounded = Math.trunc(num);
    const formatted = rounded.toLocaleString('en-US');
    return {
      cell: { t: 'n', v: rounded, z: '#,##0;-#,##0;0' },
      displayLen: formatted.length + 1,
    };
  }

  const rounded = Number(num.toFixed(2));
  const formatted = rounded.toLocaleString('en-US', { maximumFractionDigits: 2 });
  return {
    cell: { t: 'n', v: rounded, z: '#,##0.##;-#,##0.##;0' },
    displayLen: formatted.length + 1,
  };
}

export function exportToExcel(params: ExportToExcelParams): void {
  const wb = XLSX.utils.book_new();
  wb.Workbook = { Views: [{ RTL: true }] };
  const usedNames = new Set<string>();

  for (const sheet of params.sheets) {
    const ws: XLSX.WorkSheet = {};
    const colCount = Math.max(1, sheet.columns.length, sheet.summaryPairs ? 2 : 1);
    const colWidths = new Array<number>(colCount).fill(12);
    for (let c = 0; c < sheet.columns.length; c++) {
      colWidths[c] = Math.max(
        colWidths[c],
        sheet.columns[c].minWidth ?? 12,
        sheet.columns[c].header.length + 5
      );
    }

    let r = 0;
    const setCell = (rowIdx: number, colIdx: number, cell: XLSX.CellObject) => {
      const addr = XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });
      ws[addr] = cell;
    };

    if (sheet.title) {
      setCell(r, 0, { t: 's', v: sheet.title });
      r++;
    }
    if (sheet.subtitle) {
      setCell(r, 0, { t: 's', v: sheet.subtitle });
      r++;
    }
    if (sheet.title || sheet.subtitle) {
      r++; // clean spacer row
    }

    if (sheet.summaryPairs && sheet.summaryPairs.length > 0) {
      setCell(r, 0, { t: 's', v: 'المؤشر / البيان' });
      setCell(r, 1, { t: 's', v: 'القيمة' });
      colWidths[0] = Math.max(colWidths[0], 30);
      colWidths[1] = Math.max(colWidths[1], 22);
      r++;

      for (const pair of sheet.summaryPairs) {
        setCell(r, 0, { t: 's', v: pair.label });
        colWidths[0] = Math.max(colWidths[0], pair.label.length + 5);
        const { cell, displayLen } = toExcelCell(pair.value, pair.type ?? 'string');
        setCell(r, 1, cell);
        colWidths[1] = Math.max(colWidths[1], displayLen + 5);
        r++;
      }
      if (sheet.columns.length > 0) {
        r++;
      }
    }

    if (sheet.columns.length > 0) {
      // Table header row
      for (let c = 0; c < sheet.columns.length; c++) {
        const header = sheet.columns[c].header;
        setCell(r, c, { t: 's', v: header });
        colWidths[c] = Math.max(colWidths[c], header.length + 5);
      }
      r++;

      // Data rows (including zero rows!)
      for (const row of sheet.rows) {
        for (let c = 0; c < sheet.columns.length; c++) {
          const col = sheet.columns[c];
          const { cell, displayLen } = toExcelCell(row[col.key], col.type);
          setCell(r, c, cell);
          colWidths[c] = Math.max(colWidths[c], displayLen + 4);
        }
        r++;
      }

      // Total row
      if (sheet.totalRow) {
        for (let c = 0; c < sheet.columns.length; c++) {
          const col = sheet.columns[c];
          const rawVal = sheet.totalRow[col.key];
          // Total label cells in string columns might contain text like 'الإجمالي'
          const cellType: ExportColumnType =
            col.type !== 'string' && typeof rawVal === 'string' && rawVal.trim() !== '' && !Number.isFinite(Number(rawVal))
              ? 'string'
              : col.type;
          const { cell, displayLen } = toExcelCell(rawVal, cellType);
          setCell(r, c, cell);
          colWidths[c] = Math.max(colWidths[c], displayLen + 4);
        }
        r++;
      }
    }

    const finalRows = Math.max(1, r);
    ws['!ref'] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: finalRows - 1, c: colCount - 1 },
    });
    ws['!cols'] = colWidths.map((wch) => ({ wch: Math.min(60, Math.max(11, Math.ceil(wch))) }));
    ws['!views'] = [{ rightToLeft: true }];

    const sheetName = sanitizeSheetName(sheet.name, usedNames);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  }

  XLSX.writeFile(wb, params.filename, { compression: true });
}

function escapeHtml(raw: string): string {
  return String(raw || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function formatPdfCellHtml(value: string | number | null | undefined, type: ExportColumnType): { html: string; isLtr: boolean; extraClass?: string } {
  if (type === 'string') {
    const str = value === null || value === undefined || value === '' ? '—' : String(value);
    return { html: escapeHtml(str), isLtr: false };
  }

  if (value === null || value === undefined || value === '') {
    return { html: '—', isLtr: true, extraClass: 'muted' };
  }

  const num = Number(value);
  if (!Number.isFinite(num)) {
    return { html: escapeHtml(String(value)), isLtr: false };
  }

  if (type === 'percent') {
    const cls = num >= 100 ? 'pos' : '';
    return {
      html: `${num.toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}%`,
      isLtr: true,
      extraClass: cls,
    };
  }

  if (type === 'currency') {
    const cls = num < 0 ? 'neg' : num === 0 ? 'muted' : '';
    return {
      html: num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
      isLtr: true,
      extraClass: cls,
    };
  }

  if (type === 'qty') {
    const cls = num < 0 ? 'neg' : num === 0 ? 'muted' : '';
    return {
      html: num.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 3 }),
      isLtr: true,
      extraClass: cls,
    };
  }

  if (type === 'integer') {
    return {
      html: Math.trunc(num).toLocaleString('en-US'),
      isLtr: true,
      extraClass: num === 0 ? 'muted' : '',
    };
  }

  return {
    html: num.toLocaleString('en-US', { maximumFractionDigits: 2 }),
    isLtr: true,
    extraClass: num < 0 ? 'neg' : '',
  };
}

function buildTableHeadHtml(columns: ExportColumn[]): string {
  return `<thead><tr>${columns
    .map((c) => {
      const align = c.align ?? (c.type === 'string' ? 'right' : 'center');
      return `<th style="text-align:${align};">${escapeHtml(c.header)}</th>`;
    })
    .join('')}</tr></thead>`;
}

function buildTableRowHtml(columns: ExportColumn[], row: PdfTableRow, isTotal = false): string {
  const toneClass = isTotal
    ? 'row-total'
    : row._tone === 'subtotal'
      ? 'row-subtotal'
      : row._tone === 'private'
        ? 'row-private'
        : row._tone === 'direct_army'
          ? 'row-direct-army'
          : row._tone === 'chains'
            ? 'row-chains'
            : row._tone === 'highlight'
              ? 'row-highlight'
              : '';

  const cells = columns
    .map((col) => {
      const rawVal = row[col.key];
      const effType: ExportColumnType =
        isTotal && col.type !== 'string' && typeof rawVal === 'string' && rawVal.trim() !== '' && !Number.isFinite(Number(rawVal))
          ? 'string'
          : col.type;
      const { html, isLtr, extraClass } = formatPdfCellHtml(rawVal, effType);
      const align = col.align ?? (effType === 'string' ? 'right' : 'center');
      const dirAttr = isLtr ? ' dir="ltr"' : '';
      return `<td style="text-align:${align};" class="${extraClass ?? ''}"${dirAttr}>${html}</td>`;
    })
    .join('');

  return `<tr class="${toneClass}">${cells}</tr>`;
}

let cachedSvgFontCssPromise: Promise<string> | null = null;

async function fetchAsBase64DataUri(url: string, mime: string): Promise<string> {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Failed to load asset: ${url}`);
  const buf = await res.arrayBuffer();
  const bytes = new Uint8Array(buf);
  let binary = '';
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return `data:${mime};base64,${btoa(binary)}`;
}

async function getEmbeddedSvgFontCss(): Promise<string> {
  if (!cachedSvgFontCssPromise) {
    cachedSvgFontCssPromise = (async () => {
      try {
        const [ar400, ar700, lat400, lat700] = await Promise.all([
          fetchAsBase64DataUri(cairoArabic400Url, 'font/woff2'),
          fetchAsBase64DataUri(cairoArabic700Url, 'font/woff2'),
          fetchAsBase64DataUri(cairoLatin400Url, 'font/woff2'),
          fetchAsBase64DataUri(cairoLatin700Url, 'font/woff2'),
        ]);
        return `
          @font-face { font-family: 'CairoExport'; src: url('${ar400}') format('woff2'); font-weight: 400; font-style: normal; }
          @font-face { font-family: 'CairoExport'; src: url('${lat400}') format('woff2'); font-weight: 400; font-style: normal; unicode-range: U+0000-00FF, U+0131, U+0152-0153, U+02BB-02BC, U+02C6, U+02DA, U+02DC, U+2000-206F, U+2074, U+20AC, U+2122, U+2191, U+2193, U+2212, U+2215, U+FEFF, U+FFFD; }
          @font-face { font-family: 'CairoExport'; src: url('${ar700}') format('woff2'); font-weight: 700; font-style: normal; }
          @font-face { font-family: 'CairoExport'; src: url('${lat700}') format('woff2'); font-weight: 700; font-style: normal; unicode-range: U+0000-00FF, U+0131, U+0152-0153, U+02BB-02BC, U+02C6, U+02DA, U+02DC, U+2000-206F, U+2074, U+20AC, U+2122, U+2191, U+2193, U+2212, U+2215, U+FEFF, U+FFFD; }
          text, tspan { font-family: 'CairoExport', 'Tahoma', 'Arial', sans-serif !important; }
        `;
      } catch {
        return `text, tspan { font-family: 'Tahoma', 'Arial', sans-serif !important; }`;
      }
    })();
  }
  return cachedSvgFontCssPromise;
}

function waitMs(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function nextFrames(count = 2): Promise<void> {
  return new Promise((resolve) => {
    const step = (rem: number) => {
      if (rem <= 0) resolve();
      else requestAnimationFrame(() => step(rem - 1));
    };
    step(count);
  });
}

export async function captureChartSvgToPng(
  containerEl: HTMLElement | null,
  options: { scale?: number; timeoutMs?: number } = {}
): Promise<{ dataUrl: string; width: number; height: number } | null> {
  if (!containerEl) return null;
  const scale = options.scale ?? 3;
  const timeoutMs = options.timeoutMs ?? 1800;

  if (document.fonts && document.fonts.ready) {
    await document.fonts.ready;
  }

  // Wait until Recharts mounts the SVG with non-zero size and stops animating
  const startedAt = Date.now();
  let svg: SVGSVGElement | null = null;
  let prevMarkup = '';
  let stableCount = 0;

  while (Date.now() - startedAt < timeoutMs) {
    await nextFrames(1);
    const candidate = containerEl.querySelector('svg.recharts-surface') as SVGSVGElement | null;
    if (candidate) {
      const rect = candidate.getBoundingClientRect();
      if (rect.width >= 20 && rect.height >= 20) {
        const curMarkup = candidate.innerHTML;
        if (curMarkup && curMarkup === prevMarkup) {
          stableCount++;
          if (stableCount >= 1) {
            svg = candidate;
            break;
          }
        } else {
          prevMarkup = curMarkup;
          stableCount = 0;
        }
      }
    }
    await waitMs(70);
  }

  if (!svg) {
    svg = containerEl.querySelector('svg.recharts-surface') as SVGSVGElement | null;
  }
  if (!svg) return null;

  const rect = svg.getBoundingClientRect();
  const width = Math.max(1, Math.round(rect.width || Number(svg.getAttribute('width')) || 320));
  const height = Math.max(1, Math.round(rect.height || Number(svg.getAttribute('height')) || 220));

  const clone = svg.cloneNode(true) as SVGSVGElement;
  clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
  clone.setAttribute('xmlns:xlink', 'http://www.w3.org/1999/xlink');
  clone.setAttribute('width', String(width));
  clone.setAttribute('height', String(height));
  if (!clone.getAttribute('viewBox')) {
    clone.setAttribute('viewBox', `0 0 ${width} ${height}`);
  }

  // Inline computed SVG presentation attributes from the live SVG elements so styles survive serialization
  const origNodes = Array.from(svg.querySelectorAll('*'));
  const cloneNodes = Array.from(clone.querySelectorAll('*'));
  const len = Math.min(origNodes.length, cloneNodes.length);
  const svgAttrs = [
    'fill',
    'fill-opacity',
    'stroke',
    'stroke-width',
    'stroke-opacity',
    'stroke-dasharray',
    'opacity',
    'font-size',
    'font-weight',
    'text-anchor',
    'dominant-baseline',
  ];

  for (let i = 0; i < len; i++) {
    const origEl = origNodes[i];
    const cloneEl = cloneNodes[i];
    const cs = window.getComputedStyle(origEl);
    for (const prop of svgAttrs) {
      if (!cloneEl.getAttribute(prop)) {
        const val = cs.getPropertyValue(prop);
        if (val && val !== 'none' && val !== 'normal' && !val.includes('oklch')) {
          cloneEl.setAttribute(prop, val);
        }
      }
    }
  }

  const fontCss = await getEmbeddedSvgFontCss();
  const defs = document.createElementNS('http://www.w3.org/2000/svg', 'defs');
  const styleEl = document.createElementNS('http://www.w3.org/2000/svg', 'style');
  styleEl.textContent = fontCss;
  defs.appendChild(styleEl);
  clone.insertBefore(defs, clone.firstChild);

  const serialized = new XMLSerializer().serializeToString(clone);
  const svgDataUrl = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(serialized)}`;

  const img = new Image();
  img.decoding = 'async';
  const loaded = await new Promise<boolean>((resolve) => {
    img.onload = () => resolve(true);
    img.onerror = () => resolve(false);
    img.src = svgDataUrl;
  });

  if (!loaded) return null;

  const canvas = document.createElement('canvas');
  canvas.width = Math.round(width * scale);
  canvas.height = Math.round(height * scale);
  const ctx = canvas.getContext('2d');
  if (!ctx) return null;

  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.scale(scale, scale);
  ctx.drawImage(img, 0, 0, width, height);

  return {
    dataUrl: canvas.toDataURL('image/png'),
    width,
    height,
  };
}

export async function exportToPdf(params: ExportToPdfParams): Promise<void> {
  const maxCols = params.tables.reduce((m, t) => Math.max(m, t.columns.length), 0);
  const orientation: 'portrait' | 'landscape' =
    params.orientation === 'portrait' || params.orientation === 'landscape'
      ? params.orientation
      : maxCols >= 6
        ? 'landscape'
        : 'portrait';

  const pageWidthMm = orientation === 'landscape' ? 297 : 210;
  const pageHeightMm = orientation === 'landscape' ? 210 : 297;
  // Content canvas size at 96dpi with ~10mm page margins
  const pagePxW = orientation === 'landscape' ? 1050 : 730;
  const pagePxH = orientation === 'landscape' ? 720 : 1040;
  const headerFooterReserve = 42;
  const maxBodyHeight = pagePxH - headerFooterReserve;

  const toAbsUrl = (u: string) => new URL(u, window.location.href).href;
  const exportTime = getExportDateTimeDisplay();

  const iframe = document.createElement('iframe');
  iframe.setAttribute('aria-hidden', 'true');
  iframe.style.position = 'fixed';
  iframe.style.left = '-10000px';
  iframe.style.top = '0';
  iframe.style.width = `${pagePxW + 60}px`;
  iframe.style.height = `${pagePxH + 100}px`;
  iframe.style.border = '0';
  iframe.style.visibility = 'hidden';
  document.body.appendChild(iframe);

  try {
    const doc = iframe.contentDocument;
    if (!doc) {
      throw new Error('تعذر تهيئة مستند PDF');
    }

    const htmlDoc = `<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="UTF-8" />
  <style>
    @font-face {
      font-family: 'CairoExport';
      src: url('${toAbsUrl(cairoArabic400Url)}') format('woff2');
      font-weight: 400;
      font-style: normal;
    }
    @font-face {
      font-family: 'CairoExport';
      src: url('${toAbsUrl(cairoLatin400Url)}') format('woff2');
      font-weight: 400;
      font-style: normal;
      unicode-range: U+0000-00FF, U+0131, U+0152-0153, U+02BB-02BC, U+02C6, U+02DA, U+02DC, U+2000-206F, U+2074, U+20AC, U+2122, U+2191, U+2193, U+2212, U+2215, U+FEFF, U+FFFD;
    }
    @font-face {
      font-family: 'CairoExport';
      src: url('${toAbsUrl(cairoArabic700Url)}') format('woff2');
      font-weight: 700;
      font-style: normal;
    }
    @font-face {
      font-family: 'CairoExport';
      src: url('${toAbsUrl(cairoLatin700Url)}') format('woff2');
      font-weight: 700;
      font-style: normal;
      unicode-range: U+0000-00FF, U+0131, U+0152-0153, U+02BB-02BC, U+02C6, U+02DA, U+02DC, U+2000-206F, U+2074, U+20AC, U+2122, U+2191, U+2193, U+2212, U+2215, U+FEFF, U+FFFD;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      padding: 0;
      background: #ffffff;
      color: #0f172a;
      font-family: 'CairoExport', 'Tahoma', 'Arial', sans-serif;
      direction: rtl;
      -webkit-font-smoothing: antialiased;
    }
    .pdf-page {
      width: ${pagePxW}px;
      height: ${pagePxH}px;
      background: #ffffff;
      padding: 18px 20px 14px 20px;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      overflow: hidden;
    }
    .pdf-body {
      flex: 1 1 auto;
      overflow: hidden;
    }
    .report-header {
      border-bottom: 2px solid #0f766e;
      padding-bottom: 10px;
      margin-bottom: 12px;
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 16px;
    }
    .report-header.compact {
      padding-bottom: 6px;
      margin-bottom: 10px;
      border-bottom-width: 1px;
    }
    .brand-title {
      font-size: 18px;
      font-weight: 700;
      color: #0f172a;
      line-height: 1.3;
    }
    .report-header.compact .brand-title {
      font-size: 13px;
    }
    .tab-subtitle {
      font-size: 13px;
      font-weight: 700;
      color: #0f766e;
      margin-top: 2px;
    }
    .header-meta {
      text-align: left;
      font-size: 11px;
      color: #475569;
      white-space: nowrap;
    }
    .filters-strip {
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
      margin-bottom: 10px;
      font-size: 10.5px;
      color: #334155;
      background: #f8fafc;
      border: 1px solid #e2e8f0;
      border-radius: 6px;
      padding: 6px 10px;
    }
    .filter-item {
      padding-left: 8px;
      border-left: 1px solid #cbd5e1;
    }
    .filter-item:last-child {
      border-left: 0;
      padding-left: 0;
    }
    .kpi-grid {
      display: grid;
      grid-template-columns: repeat(5, minmax(0, 1fr));
      gap: 8px;
      margin-bottom: 10px;
    }
    .kpi-card {
      border: 1px solid #e2e8f0;
      background: #f8fafc;
      border-radius: 6px;
      padding: 8px 8px;
      text-align: center;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      min-height: 68px;
    }
    .kpi-label {
      font-size: 9.5px;
      color: #475569;
      font-weight: 700;
      margin-bottom: 4px;
      line-height: 1.3;
    }
    .kpi-value-wrap {
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      gap: 2px;
    }
    .kpi-value {
      font-size: 13px;
      font-weight: 700;
      color: #0f172a;
      font-variant-numeric: tabular-nums;
      white-space: nowrap;
      line-height: 1.25;
    }
    .kpi-unit {
      font-size: 9px;
      font-weight: 700;
      color: #64748b;
      letter-spacing: 0.05em;
      line-height: 1.1;
    }
    .kpi-value.pos { color: #047857; }
    .kpi-value.neg { color: #be123c; }
    .kpi-value.accent { color: #0f766e; }
    .kpi-sub {
      font-size: 9.5px;
      color: #64748b;
      margin-top: 3px;
      line-height: 1.25;
    }
    .charts-section {
      margin-top: 8px;
    }
    .charts-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 12px;
      align-items: stretch;
    }
    .chart-card {
      border: 1px solid #cbd5e1;
      background: #ffffff;
      border-radius: 8px;
      padding: 10px 12px;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
    }
    .chart-card-header {
      display: flex;
      align-items: baseline;
      justify-content: space-between;
      gap: 8px;
      margin-bottom: 6px;
      border-bottom: 1px solid #f1f5f9;
      padding-bottom: 4px;
    }
    .chart-card-title {
      font-size: 11.5px;
      font-weight: 700;
      color: #0f172a;
    }
    .chart-card-sub {
      font-size: 9.5px;
      color: #64748b;
      font-weight: 600;
    }
    .chart-img-wrap {
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 2px 0;
    }
    .chart-img {
      display: block;
      max-width: 100%;
      max-height: 230px;
      width: auto;
      height: auto;
      object-fit: contain;
    }
    .chart-legend {
      border-top: 1px solid #f1f5f9;
      padding-top: 6px;
      margin-top: 6px;
      display: flex;
      flex-wrap: wrap;
      justify-content: center;
      gap: 6px 14px;
      font-size: 10px;
      color: #334155;
    }
    .chart-legend-item {
      display: inline-flex;
      align-items: center;
      gap: 5px;
    }
    .chart-legend-swatch {
      width: 9px;
      height: 9px;
      border-radius: 2px;
      display: inline-block;
      flex-shrink: 0;
    }
    .chart-legend-val {
      font-weight: 700;
      color: #0f172a;
      font-variant-numeric: tabular-nums;
    }
    .info-block {
      border: 1px solid #e2e8f0;
      background: #f8fafc;
      border-radius: 6px;
      padding: 7px 10px;
      margin-bottom: 10px;
      font-size: 10.5px;
      color: #334155;
    }
    .info-block-title {
      font-weight: 700;
      color: #0f172a;
      margin-bottom: 4px;
      font-size: 11px;
    }
    .info-list {
      display: flex;
      flex-wrap: wrap;
      gap: 6px 14px;
    }
    .section-heading {
      font-size: 12.5px;
      font-weight: 700;
      color: #0f172a;
      margin: 8px 0 4px 0;
    }
    .section-sub {
      font-size: 10.5px;
      color: #64748b;
      margin: 0 0 6px 0;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin-bottom: 10px;
      font-size: 10.5px;
      table-layout: auto;
    }
    thead th {
      background: #f1f5f9;
      color: #1e293b;
      font-weight: 700;
      padding: 6px 7px;
      border: 1px solid #cbd5e1;
      white-space: nowrap;
    }
    tbody td, tfoot td {
      padding: 5px 7px;
      border: 1px solid #e2e8f0;
      color: #0f172a;
      font-variant-numeric: tabular-nums;
      line-height: 1.35;
    }
    tbody tr:nth-child(even) {
      background: #fafcfd;
    }
    tbody tr.row-subtotal td {
      background: #f1f5f9;
      font-weight: 700;
      border-top: 1.5px solid #94a3b8;
      border-bottom: 1.5px solid #64748b;
    }
    tbody tr.row-private td {
      background: #e0f2fe;
      font-weight: 700;
    }
    tbody tr.row-direct-army td {
      background: #ecfdf5;
      font-weight: 700;
    }
    tbody tr.row-chains td {
      background: #f1f5f9;
      font-weight: 700;
    }
    tbody tr.row-highlight td {
      background: #f8fafc;
      font-weight: 700;
    }
    tfoot tr.row-total td, tbody tr.row-total td {
      background: #e0f2fe;
      color: #0f172a;
      font-weight: 700;
      border-top: 2px solid #0284c7;
      border-bottom: 1.5px solid #94a3b8;
    }
    td.neg { color: #be123c; font-weight: 700; }
    td.pos { color: #047857; font-weight: 700; }
    td.muted { color: #94a3b8; }
    .pdf-footer {
      border-top: 1px solid #cbd5e1;
      padding-top: 6px;
      margin-top: 6px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      font-size: 10px;
      color: #64748b;
    }
  </style>
</head>
<body><div id="pages-root"></div></body>
</html>`;

    doc.open();
    doc.write(htmlDoc);
    doc.close();

    // Wait for Cairo Arabic & Latin fonts to finish loading inside the iframe
    if (doc.fonts && doc.fonts.ready) {
      await doc.fonts.ready;
    }

    const pagesRoot = doc.getElementById('pages-root')!;

    const createPage = (isFirst: boolean) => {
      const pageEl = doc.createElement('div');
      pageEl.className = 'pdf-page';

      const bodyEl = doc.createElement('div');
      bodyEl.className = 'pdf-body';

      const headerEl = doc.createElement('div');
      headerEl.className = isFirst ? 'report-header' : 'report-header compact';
      headerEl.innerHTML = isFirst
        ? `<div>
             <div class="brand-title">تحليلات قطاع السلاسل</div>
             <div class="tab-subtitle">${escapeHtml(params.tabTitle)}</div>
           </div>
           <div class="header-meta">
             <div>تاريخ التصدير: <span dir="ltr">${escapeHtml(exportTime)}</span></div>
           </div>`
        : `<div class="brand-title">تحليلات قطاع السلاسل — ${escapeHtml(params.tabTitle)}</div>
           <div class="header-meta"><span dir="ltr">${escapeHtml(exportTime)}</span></div>`;

      bodyEl.appendChild(headerEl);

      if (isFirst && params.filtersSummary.length > 0) {
        const filtersEl = doc.createElement('div');
        filtersEl.className = 'filters-strip';
        filtersEl.innerHTML = params.filtersSummary
          .map((f) => `<span class="filter-item">${escapeHtml(f)}</span>`)
          .join('');
        bodyEl.appendChild(filtersEl);
      }

      const footerEl = doc.createElement('div');
      footerEl.className = 'pdf-footer';
      footerEl.innerHTML = `<span>تحليلات قطاع السلاسل</span><span class="page-num"></span>`;

      pageEl.appendChild(bodyEl);
      pageEl.appendChild(footerEl);
      pagesRoot.appendChild(pageEl);

      return { pageEl, bodyEl, footerEl };
    };

    let current = createPage(true);

    const overflows = () => current.bodyEl.scrollHeight > maxBodyHeight;

    const startNextPage = () => {
      current = createPage(false);
    };

    // Add KPI summary cards on Page 1 if supplied
    if (params.kpis && params.kpis.length > 0) {
      const grid = doc.createElement('div');
      grid.className = 'kpi-grid';
      grid.style.gridTemplateColumns = `repeat(${Math.min(5, params.kpis.length)}, minmax(0, 1fr))`;
      grid.innerHTML = params.kpis
        .map((k) => {
          const toneCls =
            k.tone === 'positive' ? 'pos' : k.tone === 'negative' ? 'neg' : k.tone === 'accent' ? 'accent' : '';
          let displayValue = k.value;
          let displayUnit = k.unit ?? '';
          if (!displayUnit && /\s+EGP$/i.test(displayValue)) {
            displayValue = displayValue.replace(/\s+EGP$/i, '').trim();
            displayUnit = 'EGP';
          }
          return `<div class="kpi-card">
            <div class="kpi-label">${escapeHtml(k.label)}</div>
            <div class="kpi-value-wrap">
              <div class="kpi-value ${toneCls}" dir="ltr">${escapeHtml(displayValue)}</div>
              ${displayUnit ? `<div class="kpi-unit" dir="ltr">${escapeHtml(displayUnit)}</div>` : ''}
            </div>
            ${k.sublabel ? `<div class="kpi-sub">${escapeHtml(k.sublabel)}</div>` : ''}
          </div>`;
        })
        .join('');
      current.bodyEl.appendChild(grid);
    }

    // Add informational blocks (e.g. competition selected items & factors, source verification)
    if (params.infoBlocks && params.infoBlocks.length > 0) {
      for (const block of params.infoBlocks) {
        const div = doc.createElement('div');
        div.className = 'info-block';
        div.innerHTML = `<div class="info-block-title">${escapeHtml(block.title)}</div>
          <div class="info-list">${block.items.map((item) => `<span>• ${escapeHtml(item)}</span>`).join('')}</div>`;
        current.bodyEl.appendChild(div);
        if (overflows()) {
          current.bodyEl.removeChild(div);
          startNextPage();
          current.bodyEl.appendChild(div);
        }
      }
    }

    // Add each table with clean row-boundary pagination & repeating <thead>
    for (const tableSpec of params.tables) {
      const mountTableHeader = (continued: boolean) => {
        if (tableSpec.title) {
          const h = doc.createElement('div');
          h.className = 'section-heading';
          h.textContent = continued ? `${tableSpec.title} (تابع)` : tableSpec.title;
          current.bodyEl.appendChild(h);
        }
        if (!continued && tableSpec.subtitle) {
          const sub = doc.createElement('div');
          sub.className = 'section-sub';
          sub.textContent = tableSpec.subtitle;
          current.bodyEl.appendChild(sub);
        }

        const tableEl = doc.createElement('table');
        tableEl.innerHTML = `${buildTableHeadHtml(tableSpec.columns)}<tbody></tbody>`;
        current.bodyEl.appendChild(tableEl);
        return {
          tableEl,
          tbodyEl: tableEl.querySelector('tbody')!,
        };
      };

      let { tableEl, tbodyEl } = mountTableHeader(false);

      // If even the table header alone doesn't fit on current page, push to a new page
      if (overflows()) {
        // Remove newly mounted header elements and open next page
        while (current.bodyEl.lastChild && current.bodyEl.lastChild !== tableEl) {
          current.bodyEl.removeChild(current.bodyEl.lastChild);
        }
        // Actually track nodes added cleanly:
        current.bodyEl.removeChild(tableEl);
        const lastH = current.bodyEl.lastElementChild;
        if (lastH?.classList.contains('section-sub')) {
          current.bodyEl.removeChild(lastH);
          const maybeH = current.bodyEl.lastElementChild;
          if (maybeH?.classList.contains('section-heading')) current.bodyEl.removeChild(maybeH);
        } else if (lastH?.classList.contains('section-heading')) {
          current.bodyEl.removeChild(lastH);
        }
        startNextPage();
        ({ tableEl, tbodyEl } = mountTableHeader(false));
      }

      if (tableSpec.rows.length === 0) {
        const tr = doc.createElement('tr');
        tr.innerHTML = `<td colspan="${tableSpec.columns.length}" style="text-align:center;color:#64748b;padding:12px;">لا توجد بيانات مطابقة للفلاتر الحالية</td>`;
        tbodyEl.appendChild(tr);
      } else {
        for (let idx = 0; idx < tableSpec.rows.length; idx++) {
          const row = tableSpec.rows[idx];
          const tmp = doc.createElement('tbody');
          tmp.innerHTML = buildTableRowHtml(tableSpec.columns, row, false);
          const tr = tmp.firstElementChild!;
          tbodyEl.appendChild(tr);

          if (overflows()) {
            tbodyEl.removeChild(tr);
            // Only start a new page if at least one row was rendered on this page (prevents infinite loop on giant single row)
            if (tbodyEl.children.length > 0) {
              startNextPage();
              ({ tableEl, tbodyEl } = mountTableHeader(true));
              tbodyEl.appendChild(tr);
            } else {
              tbodyEl.appendChild(tr);
            }
          }
        }
      }

      if (tableSpec.totalRow) {
        const tfoot = doc.createElement('tfoot');
        tfoot.innerHTML = buildTableRowHtml(tableSpec.columns, tableSpec.totalRow, true);
        tableEl.appendChild(tfoot);

        if (overflows()) {
          tableEl.removeChild(tfoot);
          startNextPage();
          ({ tableEl, tbodyEl } = mountTableHeader(true));
          tableEl.appendChild(tfoot);
        }
      }
    }

    // Add High-DPI Charts section cleanly (pushes cleanly to next page if insufficient height remains)
    if (params.charts && params.charts.length > 0) {
      const buildChartCardHtml = (c: PdfChartItem) => {
        const legendHtml =
          c.legend && c.legend.length > 0
            ? `<div class="chart-legend">${c.legend
                .map(
                  (item) => `<span class="chart-legend-item">
                    <span class="chart-legend-swatch" style="background:${escapeHtml(item.color)};"></span>
                    <span>${escapeHtml(item.label)}</span>
                    ${item.valueText ? `<span class="chart-legend-val" dir="ltr">${escapeHtml(item.valueText)}</span>` : ''}
                  </span>`
                )
                .join('')}</div>`
            : '';
        return `<div class="chart-card">
          <div class="chart-card-header">
            <span class="chart-card-title">${escapeHtml(c.title)}</span>
            ${c.subtitle ? `<span class="chart-card-sub">${escapeHtml(c.subtitle)}</span>` : ''}
          </div>
          <div class="chart-img-wrap">
            <img class="chart-img" src="${c.imageDataUrl}" alt="${escapeHtml(c.title)}" />
          </div>
          ${legendHtml}
        </div>`;
      };

      // Render charts in pairs of 2 per grid block for clean horizontal side-by-side display
      for (let i = 0; i < params.charts.length; i += 2) {
        const pair = params.charts.slice(i, i + 2);
        const sectionEl = doc.createElement('div');
        sectionEl.className = 'charts-section';
        const cols = pair.length === 1 ? 1 : 2;
        sectionEl.innerHTML = `${
          i === 0 && params.chartsSectionTitle
            ? `<div class="section-heading">${escapeHtml(params.chartsSectionTitle)}</div>`
            : ''
        }<div class="charts-grid" style="grid-template-columns: repeat(${cols}, minmax(0, 1fr));">${pair
          .map(buildChartCardHtml)
          .join('')}</div>`;

        current.bodyEl.appendChild(sectionEl);

        // Wait for embedded chart <img> elements to decode so scrollHeight is exact
        const imgs = Array.from(sectionEl.querySelectorAll('img'));
        await Promise.all(
          imgs.map((img) =>
            img.complete
              ? Promise.resolve()
              : new Promise<void>((res) => {
                  img.onload = () => res();
                  img.onerror = () => res();
                })
          )
        );

        if (overflows()) {
          current.bodyEl.removeChild(sectionEl);
          startNextPage();
          current.bodyEl.appendChild(sectionEl);
        }
      }
    }

    // Stamp page numbers
    const pageEls = Array.from(pagesRoot.querySelectorAll('.pdf-page')) as HTMLElement[];
    const totalPages = pageEls.length;
    pageEls.forEach((p, idx) => {
      const span = p.querySelector('.page-num');
      if (span) {
        span.textContent = `صفحة ${idx + 1} من ${totalPages}`;
      }
    });

    // Wait two frames for layout paint
    await new Promise<void>((resolve) =>
      requestAnimationFrame(() => requestAnimationFrame(() => resolve()))
    );

    const pdf = new jsPDF({
      orientation,
      unit: 'mm',
      format: 'a4',
      compress: true,
    });

    for (let i = 0; i < pageEls.length; i++) {
      const pageEl = pageEls[i];
      const canvas = await html2canvas(pageEl, {
        scale: 2,
        backgroundColor: '#ffffff',
        useCORS: true,
        logging: false,
        windowWidth: pagePxW,
        windowHeight: pagePxH,
      });
      const imgData = canvas.toDataURL('image/jpeg', 0.95);
      if (i > 0) {
        pdf.addPage('a4', orientation);
      }
      pdf.addImage(imgData, 'JPEG', 0, 0, pageWidthMm, pageHeightMm, undefined, 'FAST');
    }

    pdf.save(params.filename);
  } finally {
    iframe.remove();
  }
}
