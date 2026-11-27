import path from 'node:path';
import ExcelJS from 'exceljs';
import XLSX from 'xlsx';
import { evaluateWorkbookFormulas, getCellFormulaError } from './excelFormulaEval.js';
import { expandScoreFiles, type ScoreUploadedFile } from './expertScoreAudit.js';
import { sanitizeLongSheetNamesForExcelJs } from './performanceMatcher.js';
import { Dec } from './toolboxWorkbook.js';

export interface PrecheckScoringColumnMeta {
  key: string;
  colLetter: string;
  level1: string;
  level2?: string;
  detail?: string;
  maxScore?: string;
  minScore?: string;
  stepHint?: string;
}

export interface PrecheckSheetMeta {
  sheetName: string;
  headerRow: number; // 0-based
  scoringColumns: string[];
  scoringColumnDetails?: PrecheckScoringColumnMeta[];
  totalColumn: string;
  templateMarkedColumns: string[];
}

export interface PrecheckObjectiveSelection {
  sheetName: string;
  criterion: string;
  step?: number | null;
}

export interface PrecheckTotalIssue {
  sheet_name: string;
  section: string;
  package_name: string;
  supplier: string;
  majority_same: string;
  diff_exceeded: string;
  expert_scores: string;
  source_location: string;
}

export interface PrecheckObjectiveIssue {
  sheet_name: string;
  section: string;
  package_name: string;
  supplier: string;
  criterion: string;
  issue_type: '分数不一致' | '分值步长不符' | '专家记录缺失' | '专家数量不足' | '细则未找到';
  detail: string;
  source_location: string;
}

export interface PrecheckRejectionRecord {
  lot_code?: string;
  lot_name?: string;
  package_no?: string;
  supplier_name: string;
  review_type?: string;
  reason_category?: string;
}

export interface PrecheckExcludedItem {
  sheet_name: string;
  section: string;
  package_name: string;
  supplier: string;
  review_type: string;
  reason_category: string;
}

export interface ExpertPrecheckResult {
  fileCount: number;
  sheetCount: number;
  expertNames: string[];
  maxScoreDiff: number;
  excludeRejected: boolean;
  excludedRejectedCount: number;
  excludedRejectedItems: PrecheckExcludedItem[];
  totalCheckedCount: number;
  objectiveCheckedCount: number;
  totalItems: PrecheckTotalIssue[];
  objectiveItems: PrecheckObjectiveIssue[];
  warnings: string[];
  counts: {
    totalIssues: number;
    majoritySame: number;
    diffExceeded: number;
    objectiveIssues: number;
    scoreMismatch: number;
    stepMismatch: number;
    expertMissing: number;
    expertInsufficient: number;
    criterionMissing: number;
  };
}

interface ParsedExpertRow {
  rowIndex: number; // 1-based Excel row number
  dataSeq: number; // 1-based data row index
  section: string;
  packageName: string;
  supplier: string;
  key: string;
  itemScores: Map<string, { value: Dec; col: number }>;
  computedTotal: Dec;
  sheetTotal: Dec | null;
  totalCol: number;
}

interface ParsedExpertSheet {
  sheetName: string;
  headerRow: number; // 0-based
  headers: string[];
  scoringColumns: string[];
  totalColumn: string;
  totalColIndex: number;
  rows: ParsedExpertRow[];
}

function readWorkbookSafe(buffer: Buffer, fileName: string): XLSX.WorkBook {
  let wb: XLSX.WorkBook;
  try {
    wb = XLSX.read(buffer, {
      type: 'buffer',
      cellDates: false,
      sheetStubs: true,
      cellNF: true,
      cellStyles: true,
    });
  } catch (err) {
    throw new Error(`读取工作簿「${fileName}」失败：${(err as Error).message}`);
  }
  evaluateWorkbookFormulas(wb);
  return wb;
}

function getCellText(ws: XLSX.WorkSheet, r: number, c: number, fileName: string, sheetName: string): string {
  const cell = ws[XLSX.utils.encode_cell({ r, c })] as XLSX.CellObject | undefined;
  if (!cell || cell.t === 'z') return '';
  if (cell.t === 'e') {
    const fErr = getCellFormulaError(cell);
    throw new Error(fErr || `${fileName} / ${sheetName} 第${r + 1}行第${c + 1}列存在错误单元格（${cell.w || cell.v || '#ERROR'}）`);
  }
  if (cell.w !== undefined && cell.w !== null) return String(cell.w);
  if (cell.v === undefined || cell.v === null) return '';
  return String(cell.v);
}

function getCellNumericText(ws: XLSX.WorkSheet, r: number, c: number, fileName: string, sheetName: string): string {
  const cell = ws[XLSX.utils.encode_cell({ r, c })] as XLSX.CellObject | undefined;
  if (!cell || cell.t === 'z') return '';
  if (cell.t === 'e') {
    const fErr = getCellFormulaError(cell);
    throw new Error(fErr || `${fileName} / ${sheetName} 第${r + 1}行第${c + 1}列存在错误单元格（${cell.w || cell.v || '#ERROR'}）`);
  }
  if (cell.t === 'n' && typeof cell.v === 'number') {
    return Dec.of(String(cell.v)).fmt();
  }
  if (cell.v !== undefined && cell.v !== null) {
    return String(cell.v).trim();
  }
  return cell.w ? String(cell.w).trim() : '';
}

function getSheetRange(ws: XLSX.WorkSheet): { maxRow: number; maxCol: number } {
  const ref = ws['!ref'];
  if (!ref) return { maxRow: -1, maxCol: -1 };
  const range = XLSX.utils.decode_range(ref);
  return { maxRow: range.e.r, maxCol: range.e.c };
}

/**
 * Locate header row using the exact offline template rule:
 * Find the first row where column A (index 0) is numeric 1; the header row is immediately above it (r - 1).
 */
function findHeaderRow(ws: XLSX.WorkSheet, fileName: string, sheetName: string): number {
  const { maxRow } = getSheetRange(ws);
  for (let r = 0; r <= maxRow; r++) {
    const cell = ws[XLSX.utils.encode_cell({ r, c: 0 })] as XLSX.CellObject | undefined;
    if (!cell || cell.t === 'z') continue;
    if (cell.t === 'e') {
      const fErr = getCellFormulaError(cell);
      throw new Error(fErr || `${fileName} / ${sheetName} 第${r + 1}行第1列存在错误单元格`);
    }
    if (cell.t === 'n' && cell.v === 1) {
      if (r < 1) {
        throw new Error(`${fileName} / ${sheetName} 第1行即为序号 1，缺少表头行`);
      }
      return r - 1;
    }
  }
  throw new Error(`${fileName} / ${sheetName} 未找到首列数值序号 1，无法定位表头行`);
}

function isPureNumericLabel(s: string): boolean {
  return /^-?\d+(?:\.\d+)?$/.test(s.trim());
}

function colIndexToLetter(colIdx0: number): string {
  let n = colIdx0 + 1;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function truncateLabel(s: string, maxLen: number): string {
  return s.length > maxLen ? `${s.slice(0, maxLen)}…` : s;
}

function extractScoringAndTotalColumns(
  ws: XLSX.WorkSheet,
  headerRow: number,
  fileName: string,
  sheetName: string
): {
  headers: string[];
  scoringColumns: string[];
  scoringColumnDetails: PrecheckScoringColumnMeta[];
  scoringColIndices: number[];
  totalColumn: string;
  totalColIndex: number;
} {
  const { maxCol } = getSheetRange(ws);
  const headers: string[] = [];
  for (let c = 0; c <= maxCol; c++) {
    headers.push(getCellText(ws, headerRow, c, fileName, sheetName).trim());
  }

  let totalColumn = '';
  let totalColIndex = -1;
  for (let c = 6; c <= maxCol; c++) {
    const h = headers[c];
    if (h && h.includes('合计')) {
      totalColumn = h;
      totalColIndex = c;
      break;
    }
  }

  if (totalColIndex < 0) {
    throw new Error(`${fileName} / ${sheetName} 第${headerRow + 1}行表头第7列起未找到名称包含「合计」的总分列`);
  }

  // Row 0 is sheet title; Row 1 is fixed meta header ("序号 | ... | 评审要素"); scoring element headers start at Row 2 (R3)
  const startHeaderRow = headerRow >= 2 ? 2 : 1;

  const scoringColIndices: number[] = [];
  for (let c = 6; c < totalColIndex; c++) {
    let hasContent = false;
    for (let r = startHeaderRow; r <= headerRow; r++) {
      if (getCellText(ws, r, c, fileName, sheetName).trim() !== '') {
        hasContent = true;
        break;
      }
    }
    if (hasContent) scoringColIndices.push(c);
  }

  if (scoringColIndices.length === 0) {
    throw new Error(`${fileName} / ${sheetName} 第${headerRow + 1}行表头「合计」列之前没有评分项列`);
  }

  // Strip trailing pure-numeric metadata rows (最高分 / 最低分 / 极差步长) from bottom upward
  const numericHeaderRows: number[] = [];
  for (let r = headerRow; r >= startHeaderRow; r--) {
    let nonEmpty = 0;
    let allNumeric = true;
    for (const c of scoringColIndices) {
      const txt = getCellText(ws, r, c, fileName, sheetName).trim();
      if (!txt) continue;
      nonEmpty++;
      if (!isPureNumericLabel(txt)) {
        allNumeric = false;
        break;
      }
    }
    if (nonEmpty > 0 && allNumeric) {
      numericHeaderRows.unshift(r);
    } else {
      break;
    }
  }

  const lastTextRow = numericHeaderRows.length > 0 ? numericHeaderRows[0] - 1 : headerRow;
  const textHeaderRows: number[] = [];
  for (let r = startHeaderRow; r <= lastTextRow; r++) {
    const nonEmptyTexts: string[] = [];
    for (const c of scoringColIndices) {
      const txt = getCellText(ws, r, c, fileName, sheetName).replace(/\s+/g, ' ').trim();
      if (txt) nonEmptyTexts.push(txt);
    }
    if (nonEmptyTexts.length === 0) continue;
    if (nonEmptyTexts.every((t) => t === '评审要素')) continue;
    textHeaderRows.push(r);
  }
  if (textHeaderRows.length === 0) {
    textHeaderRows.push(headerRow);
  }

  // Forward-fill horizontally across scoring columns for each text header row (handles unmerged ECP parent headers)
  const textMatrix: string[][] = textHeaderRows.map((r) => {
    const rowVals: string[] = [];
    let last = '';
    for (const c of scoringColIndices) {
      const raw = getCellText(ws, r, c, fileName, sheetName).replace(/\s+/g, ' ').trim();
      if (raw) {
        last = raw;
        rowVals.push(raw);
      } else {
        rowVals.push(last);
      }
    }
    return rowVals;
  });

  const maxRowIdx = numericHeaderRows.length >= 2 ? numericHeaderRows[0] : -1;
  const minRowIdx = numericHeaderRows.length >= 2 ? numericHeaderRows[1] : -1;
  const stepRowIdx =
    numericHeaderRows.length >= 3
      ? numericHeaderRows[2]
      : numericHeaderRows.length === 1
        ? numericHeaderRows[0]
        : -1;

  const rawScoringLabels: string[] = [];
  const rawMeta: Array<Omit<PrecheckScoringColumnMeta, 'key'>> = [];

  for (let idx = 0; idx < scoringColIndices.length; idx++) {
    const c = scoringColIndices[idx];
    const colLetter = colIndexToLetter(c);
    const level1 = textMatrix[0][idx] || `${colLetter}列`;
    let level2 = textMatrix.length >= 2 ? textMatrix[1][idx] || '' : '';
    if (level2 && level2 === level1) level2 = '';
    const detail =
      textMatrix.length >= 3
        ? textMatrix
            .slice(2)
            .map((row) => row[idx])
            .filter((t) => Boolean(t) && t !== level2 && t !== level1)
            .join('；')
        : '';

    const maxScore = maxRowIdx >= 0 ? getCellNumericText(ws, maxRowIdx, c, fileName, sheetName).trim() || undefined : undefined;
    const minScore = minRowIdx >= 0 ? getCellNumericText(ws, minRowIdx, c, fileName, sheetName).trim() || undefined : undefined;
    const stepHint = stepRowIdx >= 0 ? getCellNumericText(ws, stepRowIdx, c, fileName, sheetName).trim() || undefined : undefined;

    const canonical = level2 ? `${level1} / ${truncateLabel(level2, 36)}` : level1;
    rawScoringLabels.push(canonical);
    rawMeta.push({
      colLetter,
      level1,
      level2: level2 || undefined,
      detail: detail || undefined,
      maxScore,
      minScore,
      stepHint,
    });
  }

  // Deduplicate any identical scoring column labels by appending Excel column letter
  const labelCounts = new Map<string, number>();
  for (const lbl of rawScoringLabels) labelCounts.set(lbl, (labelCounts.get(lbl) || 0) + 1);
  const scoringColumns = rawScoringLabels.map((lbl, idx) =>
    (labelCounts.get(lbl) || 0) > 1 ? `${lbl} (${rawMeta[idx].colLetter}列)` : lbl
  );
  const scoringColumnDetails: PrecheckScoringColumnMeta[] = rawMeta.map((m, idx) => ({
    key: scoringColumns[idx],
    ...m,
  }));

  return { headers, scoringColumns, scoringColumnDetails, scoringColIndices, totalColumn, totalColIndex };
}

function isExcelJsCellFilled(cell: ExcelJS.Cell): boolean {
  const fill = cell.fill;
  if (!fill || fill.type !== 'pattern') return false;
  if (!fill.pattern || fill.pattern === 'none') return false;
  const fg = fill.fgColor as { argb?: string; theme?: number; indexed?: number } | undefined;
  if (!fg) return false;
  if (fg.theme !== undefined) return true;
  if (fg.indexed !== undefined) return true;
  if (fg.argb) {
    const upper = fg.argb.toUpperCase();
    if (upper !== '00000000' && upper !== 'FFFFFFFF') return true;
  }
  return false;
}

/**
 * Filter out macOS/Office junk files (.DS_Store, ._*, ~$*) and expand .zip archives.
 */
export function expandPrecheckFiles(files: ScoreUploadedFile[]): ScoreUploadedFile[] {
  const expanded = expandScoreFiles(files);
  return expanded.filter((f) => {
    const base = path.basename(f.originalName);
    if (!base) return false;
    if (base === '.DS_Store' || base.startsWith('._') || base.startsWith('~$')) return false;
    return true;
  });
}

/**
 * Derive clean expert display names from filenames.
 * Follows original Mac checker convention: substring before the first dot (e.g. "李.技术.xlsx" -> "李").
 * If duplicate short names occur, falls back to full stem without extension; if still duplicated, throws.
 */
function resolveExpertNames(files: ScoreUploadedFile[]): string[] {
  const shortNames = files.map((f) => {
    const base = path.basename(f.originalName);
    const dotIdx = base.indexOf('.');
    const short = dotIdx > 0 ? base.slice(0, dotIdx).trim() : base.replace(/\.xlsx?$/i, '').trim();
    return short || base;
  });

  const shortCounts = new Map<string, number>();
  for (const s of shortNames) shortCounts.set(s, (shortCounts.get(s) || 0) + 1);

  const finalNames = files.map((f, idx) => {
    const s = shortNames[idx];
    if ((shortCounts.get(s) || 0) > 1) {
      return path.basename(f.originalName).replace(/\.xlsx?$/i, '').trim() || f.originalName;
    }
    return s;
  });

  const seen = new Set<string>();
  for (const name of finalNames) {
    if (seen.has(name)) {
      throw new Error(`专家打分表重复导入：「${name}」`);
    }
    seen.add(name);
  }
  return finalNames;
}

/**
 * Discover sheets, scoring columns, and template-marked objective columns (via header cell fill color).
 */
export async function discoverPrecheckMetadata(params: {
  expertFiles: ScoreUploadedFile[];
  templateFile?: ScoreUploadedFile | null;
}): Promise<{
  expandedExperts: ScoreUploadedFile[];
  sheets: PrecheckSheetMeta[];
  defaultSelections: PrecheckObjectiveSelection[];
}> {
  const expandedExperts = expandPrecheckFiles(params.expertFiles);
  if (expandedExperts.length === 0) {
    throw new Error('请至少上传一份专家线下打分表（支持多选 .xlsx / .xls 或 .zip 压缩包）');
  }
  resolveExpertNames(expandedExperts);

  // Extract fill-color marked objective columns from templateFile (or fallback to first .xlsx expert file)
  const templateSheetOrder: string[] = [];
  const templateMarkedBySheet = new Map<string, string[]>();

  const colorSourceFile:
    | { file: ScoreUploadedFile; isTemplate: boolean }
    | null = params.templateFile
    ? { file: params.templateFile, isTemplate: true }
    : (() => {
        const firstXlsx = expandedExperts.find((f) => path.extname(f.originalName).toLowerCase() === '.xlsx');
        return firstXlsx ? { file: firstXlsx, isTemplate: false } : null;
      })();

  if (colorSourceFile) {
    const srcName = colorSourceFile.file.originalName;
    const ext = path.extname(srcName).toLowerCase();
    if (colorSourceFile.isTemplate && ext !== '.xlsx') {
      throw new Error('基础模板请使用 .xlsx 格式（.xls 无法读取表头填充色）');
    }
    const srcWbXlsx = readWorkbookSafe(colorSourceFile.file.buffer, srcName);
    const srcHeaderRows = new Map<string, number>();
    for (const sName of srcWbXlsx.SheetNames) {
      if (colorSourceFile.isTemplate) templateSheetOrder.push(sName);
      srcHeaderRows.set(sName, findHeaderRow(srcWbXlsx.Sheets[sName], srcName, sName));
    }

    if (ext === '.xlsx') {
      const ejWb = new ExcelJS.Workbook();
      const safeBuf = sanitizeLongSheetNamesForExcelJs(colorSourceFile.file.buffer);
      await ejWb.xlsx.load(safeBuf as unknown as Parameters<typeof ejWb.xlsx.load>[0]);
      srcWbXlsx.SheetNames.forEach((sName, sIdx) => {
        const wsEj = ejWb.getWorksheet(sName) || ejWb.worksheets[sIdx];
        const hRow0 = srcHeaderRows.get(sName);
        if (!wsEj || hRow0 === undefined) return;
        const marked: string[] = [];
        const wsX = srcWbXlsx.Sheets[sName];
        const { scoringColumns, scoringColIndices } = extractScoringAndTotalColumns(wsX, hRow0, srcName, sName);
        const minColorRow = hRow0 >= 2 ? 2 : 1;
        for (let idx = 0; idx < scoringColIndices.length; idx++) {
          const c = scoringColIndices[idx];
          let filled = false;
          for (let r = hRow0; r >= minColorRow; r--) {
            if (isExcelJsCellFilled(wsEj.getRow(r + 1).getCell(c + 1))) {
              filled = true;
              break;
            }
          }
          if (filled) {
            marked.push(scoringColumns[idx]);
          }
        }
        templateMarkedBySheet.set(sName, marked);
      });
    }
  }

  // Use first expert file (or template sheet order when provided) to build sheet metadata
  const firstExpert = expandedExperts[0];
  const firstWb = readWorkbookSafe(firstExpert.buffer, firstExpert.originalName);
  const sheetNames = templateSheetOrder.length > 0 ? templateSheetOrder : firstWb.SheetNames;

  const sheets: PrecheckSheetMeta[] = [];
  const defaultSelections: PrecheckObjectiveSelection[] = [];

  for (const sheetName of sheetNames) {
    const ws = firstWb.Sheets[sheetName];
    if (!ws) {
      throw new Error(`专家打分表「${firstExpert.originalName}」缺少基础模板中的工作表「${sheetName}」`);
    }
    const headerRow = findHeaderRow(ws, firstExpert.originalName, sheetName);
    const { scoringColumns, scoringColumnDetails, totalColumn } = extractScoringAndTotalColumns(
      ws,
      headerRow,
      firstExpert.originalName,
      sheetName
    );

    const rawMarked = templateMarkedBySheet.get(sheetName) || [];
    // Keep only marked columns that exist in scoringColumns (or all marked columns for strict validation)
    const markedSet = new Set(rawMarked);
    const markedScoring = scoringColumns.filter((col) => markedSet.has(col));
    // Also preserve any marked column outside scoringColumns (in case template marked a renamed/missing column)
    for (const m of rawMarked) {
      if (!markedScoring.includes(m) && !m.includes('合计')) {
        markedScoring.push(m);
      }
    }

    sheets.push({
      sheetName,
      headerRow,
      scoringColumns,
      scoringColumnDetails,
      totalColumn,
      templateMarkedColumns: markedScoring,
    });

    for (const col of markedScoring) {
      defaultSelections.push({ sheetName, criterion: col, step: null });
    }
  }

  return { expandedExperts, sheets, defaultSelections };
}

function normKeyText(raw: string): string {
  return String(raw ?? '')
    .replace(/[\s\u3000\u00A0]+/g, '')
    .replace(/（/g, '(')
    .replace(/）/g, ')')
    .trim();
}

function parseChineseSmallInt(cn: string): number | null {
  const map: Record<string, number> = {
    零: 0, 一: 1, 二: 2, 两: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9,
  };
  if (cn === '十') return 10;
  if (cn.includes('十')) {
    const [tens, ones] = cn.split('十');
    const t = tens ? map[tens] : 1;
    const o = ones ? map[ones] : 0;
    if (t === undefined || o === undefined) return null;
    return t * 10 + o;
  }
  if (cn.length === 1 && map[cn] !== undefined) return map[cn];
  return null;
}

/**
 * Normalize package number so "包1", "包01", "包一", "第1包", "1" all normalize to "1".
 */
function normPackageKey(raw: string): string {
  const s = normKeyText(raw).toUpperCase();
  if (!s) return '';
  const digitMatch = s.match(/(\d+)/);
  if (digitMatch) {
    return String(parseInt(digitMatch[1], 10));
  }
  const cnMatch = s.match(/([一二两三四五六七八九十]+)/);
  if (cnMatch) {
    const n = parseChineseSmallInt(cnMatch[1]);
    if (n !== null) return String(n);
  }
  return s;
}

type RejectionMatcher = (
  section: string,
  packageName: string,
  supplier: string
) => { review_type: string; reason_category: string } | null;

function buildRejectionMatcher(rejections: PrecheckRejectionRecord[]): RejectionMatcher {
  interface AggRej {
    lotCodeNorm: string;
    lotNameNorm: string;
    pkgNorm: string;
    reviewTypes: Set<string>;
    reasons: Set<string>;
  }
  const bySupplier = new Map<string, AggRej[]>();

  for (const r of rejections) {
    const supNorm = normKeyText(r.supplier_name);
    if (!supNorm) continue;
    const lotCodeNorm = normKeyText(r.lot_code || '').toUpperCase();
    const lotNameNorm = normKeyText(r.lot_name || '');
    const pkgNorm = normPackageKey(r.package_no || '');
    let list = bySupplier.get(supNorm);
    if (!list) {
      list = [];
      bySupplier.set(supNorm, list);
    }
    let existing = list.find(
      (x) => x.lotCodeNorm === lotCodeNorm && x.lotNameNorm === lotNameNorm && x.pkgNorm === pkgNorm
    );
    if (!existing) {
      existing = { lotCodeNorm, lotNameNorm, pkgNorm, reviewTypes: new Set(), reasons: new Set() };
      list.push(existing);
    }
    if (r.review_type && r.review_type.trim()) existing.reviewTypes.add(r.review_type.trim());
    if (r.reason_category && r.reason_category.trim()) existing.reasons.add(r.reason_category.trim());
  }

  return (section: string, packageName: string, supplier: string) => {
    const supNorm = normKeyText(supplier);
    if (!supNorm) return null;
    const candidates = bySupplier.get(supNorm);
    if (!candidates || candidates.length === 0) return null;

    const secNorm = normKeyText(section);
    const secUpper = secNorm.toUpperCase();
    const pkgNorm = normPackageKey(packageName);

    const matchedReviewTypes = new Set<string>();
    const matchedReasons = new Set<string>();
    let hit = false;

    for (const c of candidates) {
      // Package number must match whenever either side specifies a package
      if (c.pkgNorm && pkgNorm && c.pkgNorm !== pkgNorm) continue;
      if (c.pkgNorm && !pkgNorm) continue;

      // Section must match lot_name or lot_code when rejection record carries section info
      const hasRejSection = Boolean(c.lotNameNorm || c.lotCodeNorm);
      if (hasRejSection) {
        if (!secNorm) continue;
        const sectionHit =
          (c.lotNameNorm && c.lotNameNorm === secNorm) ||
          (c.lotCodeNorm && c.lotCodeNorm === secUpper);
        if (!sectionHit) continue;
      }

      hit = true;
      for (const rt of c.reviewTypes) matchedReviewTypes.add(rt);
      for (const rs of c.reasons) matchedReasons.add(rs);
    }

    if (!hit) return null;
    return {
      review_type: Array.from(matchedReviewTypes).join(' / ') || '初评否决',
      reason_category: Array.from(matchedReasons).join('；') || '未进入详评',
    };
  };
}

function parseExpertSheet(
  ws: XLSX.WorkSheet,
  fileName: string,
  sheetName: string,
  rejectionMatcher?: RejectionMatcher | null
): ParsedExpertSheet & { excludedRows: PrecheckExcludedItem[]; blankCellCount: number } {
  const headerRow = findHeaderRow(ws, fileName, sheetName);
  const { headers, scoringColumns, scoringColIndices, totalColumn, totalColIndex } = extractScoringAndTotalColumns(
    ws,
    headerRow,
    fileName,
    sheetName
  );
  const { maxRow } = getSheetRange(ws);

  // Locate metadata columns (prefer header match, fallback to standard offline template positions: B=1, C=2, E=4)
  let sectionCol = headers.findIndex((h) => h === '标段名称' || h === '分标名称');
  if (sectionCol < 0) sectionCol = 1;
  let pkgCol = headers.findIndex((h) => h === '包号' || h === '包名称');
  if (pkgCol < 0) pkgCol = 2;
  let supplierCol = headers.findIndex((h) => h === '应答人' || h === '投标人' || h === '供应商' || h === '供应商名称');
  if (supplierCol < 0) supplierCol = 4;

  const rows: ParsedExpertRow[] = [];
  const excludedRows: PrecheckExcludedItem[] = [];
  const seenKeys = new Set<string>();
  let dataSeq = 0;
  let blankCellCount = 0;

  for (let r = headerRow + 1; r <= maxRow; r++) {
    const seqText = getCellText(ws, r, 0, fileName, sheetName).trim();
    const section = getCellText(ws, r, sectionCol, fileName, sheetName).trim();
    const packageName = getCellText(ws, r, pkgCol, fileName, sheetName).trim();
    const supplier = getCellText(ws, r, supplierCol, fileName, sheetName).trim();

    // Skip rejected suppliers BEFORE reading score cells so blank/default/invalid cells on rejected rows never raise false errors
    if (supplier && rejectionMatcher) {
      const rejHit = rejectionMatcher(section, packageName, supplier);
      if (rejHit) {
        excludedRows.push({
          sheet_name: sheetName,
          section,
          package_name: packageName,
          supplier,
          review_type: rejHit.review_type,
          reason_category: rejHit.reason_category,
        });
        continue;
      }
    }

    // Check if the entire row is blank (trailing empty rows)
    let hasAnyScore = false;
    for (const scIdx of scoringColIndices) {
      if (getCellNumericText(ws, r, scIdx, fileName, sheetName).trim() !== '') {
        hasAnyScore = true;
        break;
      }
    }
    const totalStr = getCellNumericText(ws, r, totalColIndex, fileName, sheetName).trim();

    if (!seqText && !section && !packageName && !supplier && !hasAnyScore && !totalStr) {
      continue;
    }

    if (!supplier) {
      if (hasAnyScore || totalStr) {
        throw new Error(`${fileName} / ${sheetName} 第${r + 1}行有分数但缺少应答人/投标人名称`);
      }
      continue;
    }

    dataSeq++;
    const key = `${section}::${packageName}::${supplier}`;
    if (seenKeys.has(key)) {
      throw new Error(`${fileName} / ${sheetName} 第${r + 1}行存在重复应答人记录：「${supplier}」（${section} ${packageName}）`);
    }
    seenKeys.add(key);

    const itemScores = new Map<string, { value: Dec; col: number }>();
    let computedTotal = Dec.ZERO;

    for (let i = 0; i < scoringColumns.length; i++) {
      const colName = scoringColumns[i];
      const colIdx = scoringColIndices[i];
      const numStr = getCellNumericText(ws, r, colIdx, fileName, sheetName).trim();
      let valDec: Dec;
      if (!numStr) {
        // Match Excel SUM() / pandas sum(): treat blank scoring cells (e.g. un-deducted items) as 0
        valDec = Dec.ZERO;
        blankCellCount++;
      } else {
        try {
          // Allow negative deduction scores (e.g. 诚信评价扣分项 -15 / -30)
          valDec = Dec.of(numStr).stripTrailingZeros();
        } catch {
          throw new Error(`${fileName} / ${sheetName} 第${r + 1}行第${colIdx + 1}列（${supplier} / ${colName}）分数无效：${numStr}`);
        }
      }
      itemScores.set(colName, { value: valDec, col: colIdx + 1 });
      computedTotal = computedTotal.add(valDec);
    }

    computedTotal = computedTotal.stripTrailingZeros();

    let sheetTotal: Dec | null = null;
    if (totalStr !== '') {
      try {
        sheetTotal = Dec.of(totalStr).stripTrailingZeros();
      } catch {
        throw new Error(`${fileName} / ${sheetName} 第${r + 1}行第${totalColIndex + 1}列（${supplier} / ${totalColumn}）合计分无效：${totalStr}`);
      }
    }

    rows.push({
      rowIndex: r + 1,
      dataSeq,
      section,
      packageName,
      supplier,
      key,
      itemScores,
      computedTotal,
      sheetTotal,
      totalCol: totalColIndex + 1,
    });
  }

  return {
    sheetName,
    headerRow,
    headers,
    scoringColumns,
    totalColumn,
    totalColIndex,
    rows,
    excludedRows,
    blankCellCount,
  };
}

function isMultipleOfStep(value: Dec, step: Dec): boolean {
  const [, rem] = value.divAndRem(step);
  return rem.signum() === 0;
}

/**
 * Run offline expert scoring template pre-check (both total scores and objective scores).
 * Strictly aligned 1:1 with online `runExpertTotalAudit` and `runExpertObjectiveAudit`:
 * - Exact `Dec` arithmetic (rejects commas, blanks, negatives, formula errors)
 * - Total score: majority same (`cnt > Math.floor(experts / 2)`) + pairwise spread strictly > threshold (default 8)
 * - Objective score: `分数不一致`, `分值步长不符`, `专家记录缺失`, `专家数量不足`, `细则未找到`
 * - Primary-key alignment by (section, packageName, supplier) so row reordering never causes false matches or missed errors
 * - Automatically skips suppliers rejected in initial review (eval_rejections) when `excludeRejected` is enabled
 */
export function runExpertPrecheckAudit(params: {
  expertFiles: ScoreUploadedFile[];
  sheetNames?: string[];
  maxScoreDiff?: number;
  objectiveSelections: PrecheckObjectiveSelection[];
  excludeRejected?: boolean;
  rejections?: PrecheckRejectionRecord[];
}): ExpertPrecheckResult {
  const expandedExperts = expandPrecheckFiles(params.expertFiles);
  if (expandedExperts.length === 0) {
    throw new Error('请至少上传一份专家线下打分表（支持 .xlsx / .xls / .zip）');
  }

  const expertNames = resolveExpertNames(expandedExperts);
  const threshold =
    params.maxScoreDiff !== undefined && Number.isFinite(params.maxScoreDiff) ? params.maxScoreDiff : 8;
  const thresholdDec = Dec.of(String(threshold));
  if (thresholdDec.signum() < 0) {
    throw new Error('分差阈值不能为负数');
  }

  // Validate objective selections & parse step values with Dec
  const objectiveRulesBySheet = new Map<string, Map<string, Dec | null>>();
  for (const sel of params.objectiveSelections || []) {
    const sName = String(sel.sheetName || '').trim();
    const criterion = String(sel.criterion || '').trim();
    if (!sName || !criterion) continue;
    let rules = objectiveRulesBySheet.get(sName);
    if (!rules) {
      rules = new Map();
      objectiveRulesBySheet.set(sName, rules);
    }
    if (rules.has(criterion)) {
      throw new Error(`同一工作表重复配置客观项：${sName} / ${criterion}`);
    }
    const rawStep = sel.step !== undefined && sel.step !== null ? String(sel.step).trim() : '';
    let stepDec: Dec | null = null;
    if (rawStep !== '') {
      try {
        stepDec = Dec.of(rawStep);
      } catch {
        throw new Error(`分值步长必须大于 0：${criterion}`);
      }
      if (stepDec.signum() <= 0) {
        throw new Error(`分值步长必须大于 0：${criterion}`);
      }
    }
    rules.set(criterion, stepDec);
  }

  // Read all workbooks once
  const workbooks = expandedExperts.map((f, idx) => ({
    expertName: expertNames[idx],
    fileName: f.originalName,
    wb: readWorkbookSafe(f.buffer, f.originalName),
  }));

  const targetSheets =
    params.sheetNames && params.sheetNames.length > 0 ? params.sheetNames : workbooks[0].wb.SheetNames;

  const excludeRejected = params.excludeRejected !== false;
  const rejectionMatcher =
    excludeRejected && params.rejections && params.rejections.length > 0
      ? buildRejectionMatcher(params.rejections)
      : null;

  const totalItems: PrecheckTotalIssue[] = [];
  const objectiveItems: PrecheckObjectiveIssue[] = [];
  const warnings: string[] = [];
  const excludedMap = new Map<string, PrecheckExcludedItem>();
  let totalCheckedCount = 0;
  let objectiveCheckedCount = 0;

  if (workbooks.length < 2) {
    warnings.push('当前仅上传了 1 位专家的打分表，无法核验多位专家之间的分差、超半数同分与客观分一致性。');
  }

  for (const sheetName of targetSheets) {
    const parsedByExpert: Array<{
      expertName: string;
      fileName: string;
      sheet: ParsedExpertSheet;
    }> = [];

    for (const entry of workbooks) {
      const ws = entry.wb.Sheets[sheetName];
      if (!ws) {
        warnings.push(`专家「${entry.expertName}」（${entry.fileName}）缺少工作表「${sheetName}」`);
        continue;
      }
      const parsed = parseExpertSheet(ws, entry.fileName, sheetName, rejectionMatcher);
      if (parsed.blankCellCount > 0) {
        warnings.push(`${entry.fileName} / ${sheetName} 有 ${parsed.blankCellCount} 处空白评分格，已按 0 分参与合计`);
      }
      for (const ex of parsed.excludedRows) {
        const exKey = `${ex.sheet_name}::${ex.section}::${ex.package_name}::${ex.supplier}`;
        if (!excludedMap.has(exKey)) excludedMap.set(exKey, ex);
      }
      parsedByExpert.push({
        expertName: entry.expertName,
        fileName: entry.fileName,
        sheet: parsed,
      });
    }

    if (parsedByExpert.length === 0) {
      warnings.push(`工作表「${sheetName}」在所有专家表中均不存在，已跳过`);
      continue;
    }

    // Check row order consistency across experts (informational warning; alignment still uses primary key)
    const baseOrder = parsedByExpert[0].sheet.rows.map((r) => r.key);
    for (let i = 1; i < parsedByExpert.length; i++) {
      const curOrder = parsedByExpert[i].sheet.rows.map((r) => r.key);
      if (curOrder.length !== baseOrder.length) {
        warnings.push(
          `工作表「${sheetName}」中专家「${parsedByExpert[i].expertName}」有效行数（${curOrder.length}）与「${parsedByExpert[0].expertName}」（${baseOrder.length}）不一致，已按应答人主键对齐核验`
        );
      } else if (curOrder.some((k, idx) => k !== baseOrder[idx])) {
        warnings.push(
          `工作表「${sheetName}」中专家「${parsedByExpert[i].expertName}」的应答人行序与其他专家不同，已自动按应答人主键对齐核验`
        );
      }
    }

    // Collect union of all suppliers in deterministic order (first expert order, then any extras)
    const supplierMetaMap = new Map<
      string,
      { section: string; packageName: string; supplier: string; key: string }
    >();
    for (const p of parsedByExpert) {
      for (const row of p.sheet.rows) {
        if (!supplierMetaMap.has(row.key)) {
          supplierMetaMap.set(row.key, {
            section: row.section,
            packageName: row.packageName,
            supplier: row.supplier,
            key: row.key,
          });
        }
        // Check sheet's own "合计" column vs exact sum of scoring items
        if (row.sheetTotal !== null && !row.sheetTotal.eq(row.computedTotal)) {
          warnings.push(
            `合计不一致：${p.fileName} / ${sheetName} 第${row.rowIndex}行「${row.supplier}」表内「${p.sheet.totalColumn}」为 ${row.sheetTotal.fmt()}，但各项评分精确累加为 ${row.computedTotal.fmt()}（已按各项精确累加值参与总分核验）`
          );
        }
      }
    }

    const rowMapByExpert = parsedByExpert.map((p) => {
      const m = new Map<string, ParsedExpertRow>();
      for (const r of p.sheet.rows) m.set(r.key, r);
      return { ...p, rowMap: m };
    });

    // ==================== 1. Total Score Check (1:1 parity with runExpertTotalAudit) ====================
    for (const supMeta of supplierMetaMap.values()) {
      const presentExperts: Array<{
        expertName: string;
        fileName: string;
        score: Dec;
        rowIndex: number;
      }> = [];
      const missingExpertNames: string[] = [];

      for (const p of rowMapByExpert) {
        const row = p.rowMap.get(supMeta.key);
        if (!row) {
          missingExpertNames.push(p.expertName);
        } else {
          presentExperts.push({
            expertName: p.expertName,
            fileName: p.fileName,
            score: row.computedTotal,
            rowIndex: row.rowIndex,
          });
        }
      }

      if (missingExpertNames.length > 0) {
        warnings.push(
          `总分核验缺漏：工作表「${sheetName}」应答人「${supMeta.supplier}」缺少专家 ${missingExpertNames.join('、')} 的评分记录`
        );
      }

      if (presentExperts.length === 0) continue;
      totalCheckedCount++;

      const scores = presentExperts.map((x) => x.score);
      const labels = presentExperts.map((x) => x.expertName);
      const expertsCount = presentExperts.length;

      // 1.1 超过半数同分 (cnt > Math.floor(experts / 2)) — only meaningful when experts >= 2
      const majorityParts: string[] = [];
      if (expertsCount >= 2) {
        const countsMap = new Map<string, { score: Dec; count: number }>();
        for (const sc of scores) {
          const k = sc.fmt();
          const cur = countsMap.get(k);
          if (cur) cur.count++;
          else countsMap.set(k, { score: sc, count: 1 });
        }
        for (const { score: scoreVal, count: cnt } of countsMap.values()) {
          if (cnt > Math.floor(expertsCount / 2)) {
            const namesAtScore: string[] = [];
            for (let k = 0; k < scores.length; k++) {
              if (scores[k].eq(scoreVal)) namesAtScore.push(labels[k]);
            }
            majorityParts.push(`${cnt}个相同分数：${scoreVal.fmt()}；专家：${namesAtScore.join('、')}`);
          }
        }
      }

      // 1.2 任意两名专家分差严格大于阈值 (diff > threshold)
      const diffSet = new Set<string>();
      for (let i = 0; i < scores.length; i++) {
        for (let j = i + 1; j < scores.length; j++) {
          const high = scores[i].max(scores[j]);
          const low = scores[i].min(scores[j]);
          const diff = high.sub(low);
          if (diff.cmp(thresholdDec) > 0) {
            const highIdx = scores[i].cmp(scores[j]) > 0 ? i : j;
            const lowIdx = highIdx === i ? j : i;
            diffSet.add(
              `分差${diff.fmt()}分：高分${high.fmt()}（${labels[highIdx]}），低分${low.fmt()}（${labels[lowIdx]}）`
            );
          }
        }
      }

      if (majorityParts.length > 0 || diffSet.size > 0) {
        totalItems.push({
          sheet_name: sheetName,
          section: supMeta.section,
          package_name: supMeta.packageName,
          supplier: supMeta.supplier,
          majority_same: majorityParts.join('；'),
          diff_exceeded: Array.from(diffSet).join('；\n'),
          expert_scores: presentExperts.map((x) => `${x.expertName}：${x.score.fmt()}`).join('；'),
          source_location: presentExperts.map((x) => `${x.fileName} / 第${x.rowIndex}行`).join('\n'),
        });
      }
    }

    // ==================== 2. Objective Score Check (1:1 parity with runExpertObjectiveAudit) ====================
    const sheetObjRules = objectiveRulesBySheet.get(sheetName);
    if (!sheetObjRules || sheetObjRules.size === 0) continue;

    // 2.1 Check missing columns ("细则未找到") per expert
    for (const criterion of sheetObjRules.keys()) {
      for (const p of rowMapByExpert) {
        if (!p.sheet.scoringColumns.includes(criterion)) {
          objectiveItems.push({
            sheet_name: sheetName,
            section: '',
            package_name: '',
            supplier: '',
            criterion,
            issue_type: '细则未找到',
            detail: `${p.expertName}：未找到完全一致的评分项列名`,
            source_location: `${p.fileName} / ${sheetName}`,
          });
          warnings.push(`${p.fileName} / ${sheetName} 未找到客观项列：${criterion}`);
        }
      }
    }

    // 2.2 Check per (supplier, criterion) across all experts
    const allExpertNamesInSheet = rowMapByExpert.map((p) => p.expertName);
    for (const supMeta of supplierMetaMap.values()) {
      for (const [criterion, stepDec] of sheetObjRules.entries()) {
        const valuesArr: Array<{ expertName: string; value: Dec; source: string }> = [];
        const missingExperts: string[] = [];

        for (const p of rowMapByExpert) {
          if (!p.sheet.scoringColumns.includes(criterion)) continue;
          const row = p.rowMap.get(supMeta.key);
          if (!row) {
            missingExperts.push(p.expertName);
            continue;
          }
          const cellVal = row.itemScores.get(criterion);
          if (!cellVal) {
            missingExperts.push(p.expertName);
            continue;
          }
          valuesArr.push({
            expertName: p.expertName,
            value: cellVal.value,
            source: `${p.fileName} / ${sheetName} / ${row.rowIndex}行${cellVal.col}列`,
          });
        }

        if (valuesArr.length === 0) continue;
        objectiveCheckedCount++;

        const detailText = valuesArr.map((v) => `${v.expertName}：${v.value.fmt()}分`).join('\n');
        const sourcesText = valuesArr.map((v) => v.source).join('\n');

        // Also include experts who lacked this entire sheet in workbooks
        const sheetMissingExperts = expertNames.filter((n) => !allExpertNamesInSheet.includes(n));
        const combinedMissing = [...missingExperts, ...sheetMissingExperts];
        if (combinedMissing.length > 0) {
          objectiveItems.push({
            sheet_name: sheetName,
            section: supMeta.section,
            package_name: supMeta.packageName,
            supplier: supMeta.supplier,
            criterion,
            issue_type: '专家记录缺失',
            detail: `未找到：${combinedMissing.join('、')}`,
            source_location: sourcesText,
          });
        }

        if (valuesArr.length < 2) {
          objectiveItems.push({
            sheet_name: sheetName,
            section: supMeta.section,
            package_name: supMeta.packageName,
            supplier: supMeta.supplier,
            criterion,
            issue_type: '专家数量不足',
            detail: `仅1位专家有此项评分，无法核验一致性；${detailText}`,
            source_location: sourcesText,
          });
        }

        const distinctScores = new Set(valuesArr.map((v) => v.value.fmt()));
        if (distinctScores.size > 1) {
          objectiveItems.push({
            sheet_name: sheetName,
            section: supMeta.section,
            package_name: supMeta.packageName,
            supplier: supMeta.supplier,
            criterion,
            issue_type: '分数不一致',
            detail: detailText,
            source_location: sourcesText,
          });
        }

        if (stepDec !== null) {
          for (const v of valuesArr) {
            if (!isMultipleOfStep(v.value, stepDec)) {
              objectiveItems.push({
                sheet_name: sheetName,
                section: supMeta.section,
                package_name: supMeta.packageName,
                supplier: supMeta.supplier,
                criterion,
                issue_type: '分值步长不符',
                detail: `${v.expertName}：${v.value.fmt()}分，不是${stepDec.fmt()}的倍数`,
                source_location: v.source,
              });
            }
          }
        }
      }
    }
  }

  const excludedRejectedItems = Array.from(excludedMap.values());

  if (totalCheckedCount === 0 && excludedRejectedItems.length === 0) {
    throw new Error('未找到可核验的专家评分记录，请检查表格格式是否包含首列序号 1 及应答人信息');
  }
  if (totalCheckedCount === 0 && excludedRejectedItems.length > 0) {
    warnings.push(`当前表格中全部 ${excludedRejectedItems.length} 条应答人均已命中初评否决并自动排除，无剩余进入详评的应答人需要核验。`);
  }

  return {
    fileCount: expandedExperts.length,
    sheetCount: targetSheets.length,
    expertNames,
    maxScoreDiff: threshold,
    excludeRejected,
    excludedRejectedCount: excludedRejectedItems.length,
    excludedRejectedItems,
    totalCheckedCount,
    objectiveCheckedCount,
    totalItems,
    objectiveItems,
    warnings,
    counts: {
      totalIssues: totalItems.length,
      majoritySame: totalItems.filter((x) => Boolean(x.majority_same)).length,
      diffExceeded: totalItems.filter((x) => Boolean(x.diff_exceeded)).length,
      objectiveIssues: objectiveItems.length,
      scoreMismatch: objectiveItems.filter((x) => x.issue_type === '分数不一致').length,
      stepMismatch: objectiveItems.filter((x) => x.issue_type === '分值步长不符').length,
      expertMissing: objectiveItems.filter((x) => x.issue_type === '专家记录缺失').length,
      expertInsufficient: objectiveItems.filter((x) => x.issue_type === '专家数量不足').length,
      criterionMissing: objectiveItems.filter((x) => x.issue_type === '细则未找到').length,
    },
  };
}
