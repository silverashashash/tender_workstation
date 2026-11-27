import path from 'node:path';
import ExcelJS from 'exceljs';
import { expandPrecheckFiles } from './expertPrecheckAudit.js';
import type { ScoreUploadedFile } from './expertScoreAudit.js';
import { sanitizeLongSheetNamesForExcelJs } from './performanceMatcher.js';

export interface ScoreMergeSkippedFile {
  fileName: string;
  reason: string;
}

export interface ScoreMergeSheetPreview {
  name: string;
  totalRows: number;
  colCount: number;
  headerRows: number;
  rows: Array<{
    rowNum: number;
    cells: Array<{ text: string; highlighted?: boolean }>;
  }>;
}

export interface ScoreMergeRunResult {
  buffer: Buffer;
  outputFileName: string;
  totalInputFiles: number;
  validFileCount: number;
  skippedFiles: ScoreMergeSkippedFile[];
  usedTemplate: boolean;
  activeRows: number;
  ignoredRows: number;
  unmatchedRows: number;
  conflictCount: number;
  highlightCount: number;
  headerSummary: Array<{ sheetName: string; headerRows: number }>;
  sheetPreviews: ScoreMergeSheetPreview[];
  logs: string[];
}

const YELLOW_FILL: ExcelJS.Fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFFFF00' },
};

/** Clone cell.style before setting fill so ExcelJS does not mutate shared xf styles across other cells */
function applyYellowFill(cell: ExcelJS.Cell) {
  cell.style = {
    ...cell.style,
    fill: YELLOW_FILL,
  };
}

/**
 * Convert ExcelJS shared formula clones/masters into independent standard formulas
 * so overwriting or copying cells never breaks ExcelJS writeBuffer with
 * "Shared Formula master must exist above and or left of clone".
 */
function detachSharedFormulas(wb: ExcelJS.Workbook) {
  for (const ws of wb.worksheets) {
    const resolvedClones: Array<{ cell: ExcelJS.Cell; value: ExcelJS.CellValue }> = [];
    const masterCells: Array<{ cell: ExcelJS.Cell; value: ExcelJS.CellValue }> = [];

    ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const v = cell.value;
        if (!v || typeof v !== 'object') return;
        if ('sharedFormula' in v) {
          let formulaStr = '';
          try {
            formulaStr = cell.formula || '';
          } catch {
            formulaStr = '';
          }
          const res = 'result' in v ? (v as ExcelJS.CellFormulaValue).result : undefined;
          if (formulaStr) {
            resolvedClones.push({
              cell,
              value: res !== undefined ? { formula: formulaStr, result: res } : { formula: formulaStr },
            });
          } else {
            resolvedClones.push({ cell, value: res !== undefined ? res : null });
          }
        } else if ('shareType' in v && 'formula' in v) {
          const formulaStr = String((v as { formula?: string }).formula || '');
          const res = 'result' in v ? (v as ExcelJS.CellFormulaValue).result : undefined;
          masterCells.push({
            cell,
            value: formulaStr
              ? res !== undefined
                ? { formula: formulaStr, result: res }
                : { formula: formulaStr }
              : res !== undefined
                ? res
                : null,
          });
        }
      });
    });

    // Materialize clones first while masters are still intact, then strip shareType from masters
    for (const item of resolvedClones) item.cell.value = item.value;
    for (const item of masterCells) item.cell.value = item.value;
  }
}

/** Unhide all rows and clear active autoFilter inherited from a filtered base file */
function unhideAllRowsAndClearFilter(wb: ExcelJS.Workbook) {
  for (const ws of wb.worksheets) {
    if (ws.autoFilter) {
      ws.autoFilter = undefined;
    }
    for (let r = 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      if (row.hidden) row.hidden = false;
    }
  }
}

// ==================== Cell & Merge Helpers ====================

function unwrapCellValue(val: ExcelJS.CellValue): string | number | boolean | null {
  if (val === null || val === undefined) return null;
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return val;
  if (val instanceof Date) return val.toISOString();
  if (typeof val === 'object') {
    if ('result' in val && val.result !== undefined) {
      return unwrapCellValue(val.result as ExcelJS.CellValue);
    }
    if ('formula' in val) {
      return `=${String((val as { formula?: string }).formula || '')}`;
    }
    if ('richText' in val && Array.isArray((val as { richText?: Array<{ text?: string }> }).richText)) {
      return (val as { richText: Array<{ text?: string }> }).richText.map((r) => r.text || '').join('');
    }
    if ('text' in val && typeof (val as { text?: unknown }).text === 'string') {
      return (val as { text: string }).text;
    }
    if ('error' in val) {
      return String((val as { error?: string }).error || '#ERROR');
    }
  }
  return String(val);
}

function unwrapFormulaOrScalar(val: ExcelJS.CellValue): ExcelJS.CellValue {
  if (val === null || val === undefined) return null;
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return val;
  if (typeof val === 'object' && 'richText' in val && Array.isArray((val as { richText?: Array<{ text?: string }> }).richText)) {
    return (val as { richText: Array<{ text?: string }> }).richText.map((r) => r.text || '').join('');
  }
  return val;
}

function hasValue(val: unknown): boolean {
  if (val === null || val === undefined) return false;
  if (typeof val === 'object') {
    const u = unwrapCellValue(val as ExcelJS.CellValue);
    return u !== null && String(u).trim() !== '';
  }
  return String(val).trim() !== '';
}

function canonicalValue(val: unknown): string {
  const u = typeof val === 'object' && val !== null ? unwrapCellValue(val as ExcelJS.CellValue) : val;
  if (!hasValue(u)) return 'blank:';
  if (typeof u === 'number' && Number.isFinite(u)) return `num:${u}`;
  const s = String(u).trim();
  if (/^-?\d+(?:\.\d+)?$/.test(s)) {
    const n = Number(s);
    if (Number.isFinite(n)) return `num:${n}`;
  }
  return `txt:${s}`;
}

function colLettersToNum(letters: string): number {
  let num = 0;
  for (let i = 0; i < letters.length; i++) {
    num = num * 26 + (letters.charCodeAt(i) - 64);
  }
  return num;
}

function evalSimpleSumOnSheet(
  ws: ExcelJS.Worksheet,
  formulaStr: string
): { sum: number; hasConflictText: boolean } | null {
  const sumMatch = /^=?SUM\(\s*([A-Z]+)(\d+):([A-Z]+)(\d+)\s*\)$/i.exec(formulaStr.trim());
  if (!sumMatch) return null;
  const c1 = colLettersToNum(sumMatch[1].toUpperCase());
  const r1 = parseInt(sumMatch[2], 10);
  const c2 = colLettersToNum(sumMatch[3].toUpperCase());
  const r2 = parseInt(sumMatch[4], 10);
  let sum = 0;
  let hasConflictText = false;
  for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) {
    const rowObj = ws.getRow(r);
    for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) {
      const v = unwrapCellValue(rowObj.getCell(c).value);
      if (typeof v === 'number' && Number.isFinite(v)) {
        sum += v;
      } else if (typeof v === 'string') {
        const s = v.trim();
        if (/^-?\d+(?:\.\d+)?$/.test(s)) {
          sum += Number(s);
        } else if (s.includes('|')) {
          hasConflictText = true;
        }
      }
    }
  }
  return { sum: Math.round(sum * 10000) / 10000, hasConflictText };
}

/** Resolve a cell's evaluated display value (computing SUM(...) when cached formula result is missing) */
function resolveEvaluatedCellValue(
  ws: ExcelJS.Worksheet,
  rawVal: ExcelJS.CellValue
): string | number | boolean | null {
  if (rawVal && typeof rawVal === 'object' && 'formula' in rawVal) {
    const cached = 'result' in rawVal ? (rawVal as ExcelJS.CellFormulaValue).result : undefined;
    if (cached !== undefined && cached !== null) {
      return unwrapCellValue(cached as ExcelJS.CellValue);
    }
    const formulaStr = String((rawVal as { formula?: string }).formula || '');
    const evaluated = evalSimpleSumOnSheet(ws, formulaStr);
    if (evaluated !== null) return evaluated.sum;
  }
  return unwrapCellValue(rawVal);
}

function parseA1Range(ref: string): { minRow: number; minCol: number; maxRow: number; maxCol: number } | null {
  const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i.exec(ref.trim());
  if (!m) return null;
  const c1 = colLettersToNum(m[1].toUpperCase());
  const r1 = parseInt(m[2], 10);
  const c2 = m[3] ? colLettersToNum(m[3].toUpperCase()) : c1;
  const r2 = m[4] ? parseInt(m[4], 10) : r1;
  return {
    minRow: Math.min(r1, r2),
    minCol: Math.min(c1, c2),
    maxRow: Math.max(r1, r2),
    maxCol: Math.max(c1, c2),
  };
}

function getMergedRangeStrings(ws: ExcelJS.Worksheet): string[] {
  const modelMerges = (ws.model as { merges?: string[] } | undefined)?.merges;
  if (!Array.isArray(modelMerges)) return [];
  return [...modelMerges].map((m) => String(m)).sort();
}

function buildMergedMap(ws: ExcelJS.Worksheet): Map<string, [number, number]> {
  const map = new Map<string, [number, number]>();
  for (const ref of getMergedRangeStrings(ws)) {
    const rng = parseA1Range(ref);
    if (!rng) continue;
    for (let r = rng.minRow; r <= rng.maxRow; r++) {
      for (let c = rng.minCol; c <= rng.maxCol; c++) {
        map.set(`${r},${c}`, [rng.minRow, rng.minCol]);
      }
    }
  }
  return map;
}

function getMergedCellValue(
  ws: ExcelJS.Worksheet,
  row: number,
  col: number,
  mergedMap: Map<string, [number, number]>
): string | number | boolean | null {
  const base = mergedMap.get(`${row},${col}`);
  const cell = base ? ws.getRow(base[0]).getCell(base[1]) : ws.getRow(row).getCell(col);
  return unwrapCellValue(cell.value);
}

// ==================== Layout & Business Key ====================

interface ScoreLayout {
  section: number;
  package: number;
  respondent: number;
  compliance?: number;
  total: number;
  expert?: number;
  dataStart: number;
  scoreColumns: number[];
}

function findScoreLayout(ws: ExcelJS.Worksheet, maxScan = 20): ScoreLayout | null {
  let identityHeaderRow: number | null = null;
  let totalHeaderRow: number | null = null;
  const cols: Partial<Record<'section' | 'package' | 'respondent' | 'compliance' | 'total' | 'expert', number>> = {};

  const rowLimit = Math.min(maxScan, ws.rowCount);
  const colLimit = ws.columnCount;

  for (let r = 1; r <= rowLimit; r++) {
    const rowObj = ws.getRow(r);
    for (let c = 1; c <= colLimit; c++) {
      const raw = unwrapCellValue(rowObj.getCell(c).value);
      const text = raw === null ? '' : String(raw).trim();
      if (!text) continue;
      if (text === '标段' || (text === '标段名称' && cols.section === undefined)) {
        cols.section = c;
        identityHeaderRow = r;
      } else if (text === '包号' || (text === '包名称' && cols.package === undefined)) {
        cols.package = c;
        identityHeaderRow = r;
      } else if (text === '应答人') {
        cols.respondent = c;
        identityHeaderRow = r;
      } else if (text.includes('符合性评审')) {
        cols.compliance = c;
      } else if (text === '合计') {
        cols.total = c;
        totalHeaderRow = r;
      } else if (text === '专家名称') {
        cols.expert = c;
      }
    }
  }

  if (
    cols.section === undefined ||
    cols.package === undefined ||
    cols.respondent === undefined ||
    cols.total === undefined ||
    identityHeaderRow === null ||
    totalHeaderRow === null
  ) {
    return null;
  }

  const scoreStart = (cols.compliance ?? cols.respondent + 1) + 1;
  const scoreEnd = cols.total - 1;
  if (scoreStart > scoreEnd) return null;

  const scoreColumns: number[] = [];
  for (let c = scoreStart; c <= scoreEnd; c++) scoreColumns.push(c);

  return {
    section: cols.section,
    package: cols.package,
    respondent: cols.respondent,
    compliance: cols.compliance,
    total: cols.total,
    expert: cols.expert,
    dataStart: Math.max(identityHeaderRow, totalHeaderRow) + 1,
    scoreColumns,
  };
}

function getBusinessKey(
  ws: ExcelJS.Worksheet,
  row: number,
  layout: ScoreLayout,
  mergedMap: Map<string, [number, number]>
): string | null {
  const sec = getMergedCellValue(ws, row, layout.section, mergedMap);
  const pkg = getMergedCellValue(ws, row, layout.package, mergedMap);
  const res = getMergedCellValue(ws, row, layout.respondent, mergedMap);
  if (!hasValue(sec) || !hasValue(pkg) || !hasValue(res)) return null;
  return `${String(sec).trim()}::${String(pkg).trim()}::${String(res).trim()}`;
}

function appendUnique(
  map: Map<string, Array<string | number | boolean>>,
  cellKey: string,
  value: string | number | boolean | null
) {
  if (!hasValue(value) || value === null) return;
  let list = map.get(cellKey);
  if (!list) {
    list = [];
    map.set(cellKey, list);
  }
  const canon = canonicalValue(value);
  if (list.every((existing) => canonicalValue(existing) !== canon)) {
    list.push(value);
  }
}

// ==================== Skeleton & Header Fingerprint ====================

interface SheetFingerprint {
  skeletonKey: string;
  grid: string[][];
}

interface WorkbookFingerprint {
  sheetNames: string[];
  sheets: Map<string, SheetFingerprint>;
}

async function computeFingerprint(buffer: Buffer): Promise<WorkbookFingerprint> {
  const wb = new ExcelJS.Workbook();
  const safeBuf = sanitizeLongSheetNamesForExcelJs(buffer);
  await wb.xlsx.load(safeBuf as unknown as Parameters<typeof wb.xlsx.load>[0]);
  const sheetNames = wb.worksheets.map((ws) => ws.name);
  const sheets = new Map<string, SheetFingerprint>();

  for (const ws of wb.worksheets) {
    const merged = getMergedRangeStrings(ws);
    const skeletonKey = JSON.stringify(merged);
    const grid: string[][] = [];
    const maxR = Math.min(ws.rowCount, 25); // header scan only needs top rows
    const maxC = ws.columnCount;
    for (let r = 1; r <= maxR; r++) {
      const rowObj = ws.getRow(r);
      const vals: string[] = [];
      for (let c = 1; c <= maxC; c++) {
        const v = unwrapCellValue(rowObj.getCell(c).value);
        vals.push(v === null ? '' : String(v).trim());
      }
      // Trim trailing empty cells so formatting-only columns don't cause false header mismatches
      while (vals.length > 0 && vals[vals.length - 1] === '') vals.pop();
      grid.push(vals);
    }
    sheets.set(ws.name, { skeletonKey, grid });
  }
  return { sheetNames, sheets };
}

function skeletonsMatch(a: WorkbookFingerprint, b: WorkbookFingerprint): boolean {
  if (a.sheetNames.length !== b.sheetNames.length) return false;
  for (let i = 0; i < a.sheetNames.length; i++) {
    const name = a.sheetNames[i];
    if (name !== b.sheetNames[i]) return false;
    if (a.sheets.get(name)?.skeletonKey !== b.sheets.get(name)?.skeletonKey) return false;
  }
  return true;
}

function inferHeaderRows(grids: string[][][], maxScan = 20): number {
  if (grids.length === 0) return 0;
  const limit = Math.min(maxScan, ...grids.map((g) => g.length));
  let n = 0;
  for (let r = 0; r < limit; r++) {
    const first = JSON.stringify(grids[0][r]);
    if (grids.slice(1).every((g) => JSON.stringify(g[r]) === first)) {
      n++;
    } else {
      break;
    }
  }
  return n;
}

function headersMatch(
  a: WorkbookFingerprint,
  b: WorkbookFingerprint,
  headerRowsBySheet: Map<string, number>
): boolean {
  for (const [name, n] of headerRowsBySheet.entries()) {
    if (n === 0) continue;
    const ga = a.sheets.get(name)?.grid || [];
    const gb = b.sheets.get(name)?.grid || [];
    if (ga.length < n || gb.length < n) return false;
    for (let r = 0; r < n; r++) {
      if (JSON.stringify(ga[r]) !== JSON.stringify(gb[r])) return false;
    }
  }
  return true;
}

// ==================== Success List (优秀名单) ====================

async function loadSuccessList(
  file: ScoreUploadedFile,
  logs: string[]
): Promise<Set<string> | null> {
  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(file.buffer as unknown as Parameters<typeof wb.xlsx.load>[0]);
    const ws = wb.worksheets[0];
    if (!ws) {
      logs.push('  ⚠️ 优秀名单文件没有工作表，跳过高亮。');
      return null;
    }
    const mergedMap = buildMergedMap(ws);
    let headerRow: number | null = null;
    const cols: Partial<Record<'section' | 'package' | 'respondent', number>> = {};

    for (let r = 1; r <= Math.min(5, ws.rowCount); r++) {
      const rowObj = ws.getRow(r);
      let hasRes = false;
      for (let c = 1; c <= ws.columnCount; c++) {
        const v = unwrapCellValue(rowObj.getCell(c).value);
        if (v !== null && String(v).includes('应答人')) {
          hasRes = true;
          break;
        }
      }
      if (hasRes) {
        headerRow = r;
        for (let c = 1; c <= ws.columnCount; c++) {
          const v = unwrapCellValue(rowObj.getCell(c).value);
          const text = v === null ? '' : String(v).trim();
          if (text.includes('分标名称') || text.includes('标段')) cols.section = c;
          else if (text.includes('分包名称') || text.includes('包号')) cols.package = c;
          else if (text.includes('应答人')) cols.respondent = c;
        }
        break;
      }
    }

    if (!headerRow || !cols.section || !cols.package || !cols.respondent) {
      logs.push('  ⚠️ 优秀名单格式无法识别，需包含：分标名称（或标段）、分包名称（或包号）、应答人。跳过高亮。');
      return null;
    }

    const set = new Set<string>();
    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const sec = getMergedCellValue(ws, r, cols.section, mergedMap);
      const pkg = getMergedCellValue(ws, r, cols.package, mergedMap);
      const res = getMergedCellValue(ws, r, cols.respondent, mergedMap);
      if (hasValue(sec) && hasValue(pkg) && hasValue(res)) {
        set.add(`${String(sec).trim()}::${String(pkg).trim()}::${String(res).trim()}`);
      }
    }
    logs.push(`  ✅ 优秀名单加载成功，共 ${set.size} 条。`);
    return set;
  } catch (err) {
    logs.push(`  ⚠️ 读取优秀名单出错：${(err as Error).message}`);
    return null;
  }
}

// ==================== Main Pipeline ====================

export async function runScoreTemplateMerge(params: {
  scoreFiles: ScoreUploadedFile[];
  templateFile?: ScoreUploadedFile | null;
  successListFile?: ScoreUploadedFile | null;
}): Promise<ScoreMergeRunResult> {
  const expanded = expandPrecheckFiles(params.scoreFiles).filter((f) =>
    /\.xlsx$/i.test(f.originalName)
  );
  if (expanded.length === 0) {
    throw new Error('请至少上传一份专家打分表（.xlsx 或包含 .xlsx 的 .zip 压缩包）。');
  }

  const logs: string[] = [`共选中 ${expanded.length} 个专家打分表文件。`];
  const skippedFiles: ScoreMergeSkippedFile[] = [];
  let validFiles: ScoreUploadedFile[] = [];
  const headerRowsBySheet = new Map<string, number>();

  // 1. Fingerprint validation & header inference
  if (params.templateFile) {
    if (!/\.xlsx$/i.test(params.templateFile.originalName)) {
      throw new Error('原始模板文件请使用 .xlsx 格式。');
    }
    logs.push(`使用用户提供的基准模板：${params.templateFile.originalName}`);
    let baseFp: WorkbookFingerprint;
    try {
      baseFp = await computeFingerprint(params.templateFile.buffer);
    } catch (err) {
      throw new Error(`基准模板读取失败：${(err as Error).message}`);
    }

    const candidates: Array<{ file: ScoreUploadedFile; fp: WorkbookFingerprint }> = [];
    for (const f of expanded) {
      try {
        const fp = await computeFingerprint(f.buffer);
        if (skeletonsMatch(fp, baseFp)) {
          candidates.push({ file: f, fp });
        } else {
          skippedFiles.push({ fileName: f.originalName, reason: '与基准模板骨架不一致' });
        }
      } catch (err) {
        skippedFiles.push({ fileName: f.originalName, reason: `读取失败: ${(err as Error).message}` });
      }
    }

    for (const sheetName of baseFp.sheetNames) {
      const grids = [
        baseFp.sheets.get(sheetName)!.grid,
        ...candidates.map((c) => c.fp.sheets.get(sheetName)!.grid),
      ];
      headerRowsBySheet.set(sheetName, inferHeaderRows(grids));
    }

    for (const c of candidates) {
      if (headersMatch(c.fp, baseFp, headerRowsBySheet)) {
        validFiles.push(c.file);
      } else {
        skippedFiles.push({ fileName: c.file.originalName, reason: '表头与基准模板不一致' });
      }
    }
  } else {
    logs.push('未提供原始模板，使用众数法自动识别基准骨架（兼容旧版手工清空合并规则）…');
    const fps = new Map<ScoreUploadedFile, WorkbookFingerprint>();
    for (const f of expanded) {
      try {
        fps.set(f, await computeFingerprint(f.buffer));
      } catch (err) {
        skippedFiles.push({ fileName: f.originalName, reason: `读取指纹失败: ${(err as Error).message}` });
      }
    }
    if (fps.size === 0) {
      throw new Error('所有上传的打分表均无法读取，请检查文件是否为未加密的 .xlsx。');
    }

    const groups = new Map<string, ScoreUploadedFile[]>();
    for (const [f, fp] of fps.entries()) {
      const key = JSON.stringify([
        fp.sheetNames,
        ...fp.sheetNames.map((n) => fp.sheets.get(n)?.skeletonKey || ''),
      ]);
      const list = groups.get(key) || [];
      list.push(f);
      groups.set(key, list);
    }

    let bestKey = '';
    let bestList: ScoreUploadedFile[] = [];
    for (const [k, list] of groups.entries()) {
      if (list.length > bestList.length) {
        bestKey = k;
        bestList = list;
      }
    }

    bestList.sort((a, b) => a.originalName.localeCompare(b.originalName, 'zh-Hans-CN'));
    const baseFp = fps.get(bestList[0])!;
    logs.push(`  识别到基准模板：${bestList[0].originalName}`);

    for (const sheetName of baseFp.sheetNames) {
      const grids = bestList.map((f) => fps.get(f)!.sheets.get(sheetName)!.grid);
      headerRowsBySheet.set(sheetName, inferHeaderRows(grids));
    }

    for (const [k, list] of groups.entries()) {
      if (k === bestKey) continue;
      for (const f of list) {
        skippedFiles.push({ fileName: f.originalName, reason: '模板格式与多数文件不一致' });
      }
    }
    validFiles = bestList;
  }

  const headerSummary = Array.from(headerRowsBySheet.entries()).map(([sheetName, headerRows]) => ({
    sheetName,
    headerRows,
  }));
  logs.push(`  表头推断：${headerSummary.map((h) => `${h.sheetName}（表头${h.headerRows}行）`).join('、')}`);

  if (skippedFiles.length > 0) {
    logs.push(`⚠️ 跳过 ${skippedFiles.length} 个文件（模板不匹配或读取失败）：`);
    for (const s of skippedFiles) {
      logs.push(`  · ${s.fileName} — ${s.reason}`);
    }
  }

  if (validFiles.length === 0) {
    throw new Error('没有与基准模板匹配的可用打分表文件。');
  }

  // 2. Optional 优秀名单
  let successSet: Set<string> | null = null;
  if (params.successListFile) {
    successSet = await loadSuccessList(params.successListFile, logs);
  } else {
    logs.push('未提供优秀名单，跳过高亮标记。');
  }

  // 3. Load base workbook (preserving all styles, merged cells, formulas)
  validFiles.sort((a, b) => a.originalName.localeCompare(b.originalName, 'zh-Hans-CN'));
  const baseWb = new ExcelJS.Workbook();
  let sourceFiles: ScoreUploadedFile[];
  if (params.templateFile) {
    const safeBase = sanitizeLongSheetNamesForExcelJs(params.templateFile.buffer);
    await baseWb.xlsx.load(safeBase as unknown as Parameters<typeof baseWb.xlsx.load>[0]);
    detachSharedFormulas(baseWb);
    sourceFiles = validFiles;
    logs.push(`结构底表：${params.templateFile.originalName}（初始分仅作基线，不直接参与合并）`);
  } else {
    const safeBase = sanitizeLongSheetNamesForExcelJs(validFiles[0].buffer);
    await baseWb.xlsx.load(safeBase as unknown as Parameters<typeof baseWb.xlsx.load>[0]);
    detachSharedFormulas(baseWb);
    sourceFiles = validFiles.slice(1);
    logs.push(`底表：${validFiles[0].originalName}`);
  }

  let activeRows = 0;
  let ignoredRows = 0;
  let unmatchedRows = 0;
  let conflictCount = 0;
  let highlightCount = 0;

  // 4. Merge
  if (sourceFiles.length > 0) {
    logs.push(`开始合并 ${sourceFiles.length} 个专家打分表…`);
    if (params.templateFile) {
      // Template baseline merge (1:1 parity with merge_with_template_baseline in merge4.py,
      // plus automatic cell-diff baseline fallback for non-score sheets like 初评表/响应情况一览表)
      const templateSheets = new Map<
        string,
        {
          layout: ScoreLayout;
          rowsByKey: Map<string, { row: number; baseline: Array<string | number | boolean | null> }>;
        }
      >();
      const fallbackDiffSheets = new Map<
        string,
        {
          startRow: number;
          baselineMap: Map<string, string>;
        }
      >();
      const scoreValues = new Map<string, Array<string | number | boolean>>();
      const expertValues = new Map<string, Array<string | number | boolean>>();
      const diffCellValues = new Map<string, Array<string | number | boolean>>();

      for (const ws of baseWb.worksheets) {
        const layout = findScoreLayout(ws);
        if (!layout) {
          const startRow = (headerRowsBySheet.get(ws.name) || 0) + 1;
          const baselineMap = new Map<string, string>();
          for (let r = startRow; r <= ws.rowCount; r++) {
            const rowObj = ws.getRow(r);
            for (let c = 1; c <= ws.columnCount; c++) {
              const v = unwrapCellValue(rowObj.getCell(c).value);
              baselineMap.set(`${r}::${c}`, canonicalValue(v));
            }
          }
          fallbackDiffSheets.set(ws.name, { startRow, baselineMap });
          logs.push(
            `  ℹ️ ${ws.name}：未检测到带「合计」的评分区域，自动切换为「按模板基线逐格差异合并」（起始行：第 ${startRow} 行）。`
          );
          continue;
        }
        const mergedMap = buildMergedMap(ws);
        const rowsByKey = new Map<
          string,
          { row: number; baseline: Array<string | number | boolean | null> }
        >();
        const duplicateKeys = new Set<string>();

        for (let r = layout.dataStart; r <= ws.rowCount; r++) {
          const key = getBusinessKey(ws, r, layout, mergedMap);
          if (!key) continue;
          if (rowsByKey.has(key)) {
            duplicateKeys.add(key);
            continue;
          }
          const rowObj = ws.getRow(r);
          const baseline = layout.scoreColumns.map((c) => unwrapCellValue(rowObj.getCell(c).value));
          rowsByKey.set(key, { row: r, baseline });

          // Clear template's score & expert cells (preserve total formula!)
          for (const c of layout.scoreColumns) {
            rowObj.getCell(c).value = null;
          }
          if (layout.expert) {
            rowObj.getCell(layout.expert).value = null;
          }
        }

        if (duplicateKeys.size > 0) {
          logs.push(`  ⚠️ ${ws.name}：发现 ${duplicateKeys.size} 个重复业务主键，仅使用第一行。`);
        }
        templateSheets.set(ws.name, { layout, rowsByKey });
      }

      for (const sf of sourceFiles) {
        logs.push(`  合并中：${sf.originalName}`);
        try {
          const srcWb = new ExcelJS.Workbook();
          const safeSrc = sanitizeLongSheetNamesForExcelJs(sf.buffer);
          await srcWb.xlsx.load(safeSrc as unknown as Parameters<typeof srcWb.xlsx.load>[0]);
          detachSharedFormulas(srcWb);

          for (const [sheetName, { layout: targetLayout, rowsByKey: targetRows }] of templateSheets.entries()) {
            const srcWs = srcWb.getWorksheet(sheetName);
            if (!srcWs) continue;
            const srcLayout = findScoreLayout(srcWs);
            if (!srcLayout) {
              logs.push(`  ⚠️ ${sf.originalName} / ${sheetName}：无法识别评分区域。`);
              continue;
            }
            if (srcLayout.scoreColumns.length !== targetLayout.scoreColumns.length) {
              logs.push(`  ⚠️ ${sf.originalName} / ${sheetName}：评分列数量不一致。`);
              continue;
            }

            const srcMergedMap = buildMergedMap(srcWs);
            for (let sr = srcLayout.dataStart; sr <= srcWs.rowCount; sr++) {
              const key = getBusinessKey(srcWs, sr, srcLayout, srcMergedMap);
              if (!key) continue;
              const target = targetRows.get(key);
              if (!target) {
                unmatchedRows++;
                continue;
              }
              const srcRowObj = srcWs.getRow(sr);
              const srcScores = srcLayout.scoreColumns.map((c) =>
                unwrapCellValue(srcRowObj.getCell(c).value)
              );

              if (!srcScores.some((v) => hasValue(v))) {
                ignoredRows++;
                continue;
              }

              const changed = srcScores.some(
                (sv, idx) => canonicalValue(sv) !== canonicalValue(target.baseline[idx])
              );
              if (!changed) {
                ignoredRows++;
                continue;
              }

              activeRows++;
              targetLayout.scoreColumns.forEach((tc, idx) => {
                appendUnique(scoreValues, `${sheetName}::${target.row}::${tc}`, srcScores[idx]);
              });

              if (srcLayout.expert && targetLayout.expert) {
                const expVal = unwrapCellValue(srcRowObj.getCell(srcLayout.expert).value);
                appendUnique(
                  expertValues,
                  `${sheetName}::${target.row}::${targetLayout.expert}`,
                  expVal
                );
              }
            }
          }

          // Fallback cell-diff merge for non-score sheets (e.g. 初评表 / 供应商投标响应情况一览表)
          for (const [sheetName, { startRow, baselineMap }] of fallbackDiffSheets.entries()) {
            const srcWs = srcWb.getWorksheet(sheetName);
            const baseWs = baseWb.getWorksheet(sheetName);
            if (!srcWs || !baseWs) continue;
            const maxR = Math.max(baseWs.rowCount, srcWs.rowCount);
            const maxC = Math.max(baseWs.columnCount, srcWs.columnCount);

            for (let r = startRow; r <= maxR; r++) {
              const srcRowObj = srcWs.getRow(r);
              let rowChanged = false;
              let rowHasAnyVal = false;
              for (let c = 1; c <= maxC; c++) {
                const srcVal = unwrapCellValue(srcRowObj.getCell(c).value);
                if (!hasValue(srcVal)) continue;
                rowHasAnyVal = true;
                const baseCanon = baselineMap.get(`${r}::${c}`) ?? 'blank:';
                if (canonicalValue(srcVal) !== baseCanon) {
                  rowChanged = true;
                  appendUnique(diffCellValues, `${sheetName}::${r}::${c}`, srcVal);
                }
              }
              if (rowChanged) {
                activeRows++;
              } else if (rowHasAnyVal) {
                ignoredRows++;
              }
            }
          }
        } catch (err) {
          logs.push(`  ❌ 合并出错 ${sf.originalName}：${(err as Error).message}`);
        }
      }

      for (const [cellKey, vals] of scoreValues.entries()) {
        const [sheetName, rStr, cStr] = cellKey.split('::');
        const ws = baseWb.getWorksheet(sheetName);
        if (!ws) continue;
        const cell = ws.getRow(Number(rStr)).getCell(Number(cStr));
        if (vals.length === 1) {
          cell.value = vals[0];
        } else {
          cell.value = vals.map((v) => String(v).trim()).join(' | ');
          applyYellowFill(cell);
          conflictCount++;
        }
      }

      for (const [cellKey, vals] of expertValues.entries()) {
        const [sheetName, rStr, cStr] = cellKey.split('::');
        const ws = baseWb.getWorksheet(sheetName);
        if (!ws) continue;
        const cell = ws.getRow(Number(rStr)).getCell(Number(cStr));
        cell.value = vals.length === 1 ? vals[0] : vals.map((v) => String(v).trim()).join(' | ');
      }

      for (const [cellKey, vals] of diffCellValues.entries()) {
        const [sheetName, rStr, cStr] = cellKey.split('::');
        const ws = baseWb.getWorksheet(sheetName);
        if (!ws) continue;
        const cell = ws.getRow(Number(rStr)).getCell(Number(cStr));
        if (vals.length === 1) {
          cell.value = vals[0];
        } else {
          cell.value = vals.map((v) => String(v).trim()).join(' | ');
          applyYellowFill(cell);
          conflictCount++;
        }
      }

      logs.push(`  有效专家填写行：${activeRows}；忽略未调整/未填写行：${ignoredRows}。`);
      if (unmatchedRows > 0) {
        logs.push(`  ⚠️ 有 ${unmatchedRows} 行未在模板中找到对应的标段、包号和应答人。`);
      }
    } else {
      // Legacy cell-by-cell merge when no template is supplied
      // First materialize evaluated values for any formula cells missing cached result on baseWb
      for (const baseWs of baseWb.worksheets) {
        for (let r = 1; r <= baseWs.rowCount; r++) {
          const bRow = baseWs.getRow(r);
          for (let c = 1; c <= baseWs.columnCount; c++) {
            const bCell = bRow.getCell(c);
            const v = bCell.value;
            if (v && typeof v === 'object' && 'formula' in v) {
              const ev = resolveEvaluatedCellValue(baseWs, v);
              if (ev !== null) {
                bCell.value = { formula: String((v as { formula?: string }).formula || ''), result: ev };
              }
            }
          }
        }
      }

      for (const sf of sourceFiles) {
        logs.push(`  合并中：${sf.originalName}`);
        try {
          const srcWb = new ExcelJS.Workbook();
          const safeSrc = sanitizeLongSheetNamesForExcelJs(sf.buffer);
          await srcWb.xlsx.load(safeSrc as unknown as Parameters<typeof srcWb.xlsx.load>[0]);
          detachSharedFormulas(srcWb);
          for (const baseWs of baseWb.worksheets) {
            const srcWs = srcWb.getWorksheet(baseWs.name);
            if (!srcWs) continue;
            const maxR = Math.max(baseWs.rowCount, srcWs.rowCount);
            const maxC = Math.max(baseWs.columnCount, srcWs.columnCount);
            for (let r = 1; r <= maxR; r++) {
              const bRow = baseWs.getRow(r);
              const sRow = srcWs.getRow(r);
              for (let c = 1; c <= maxC; c++) {
                const bCell = bRow.getCell(c);
                const sCell = sRow.getCell(c);
                const bRaw = unwrapFormulaOrScalar(bCell.value);
                const sRaw = unwrapFormulaOrScalar(sCell.value);
                const sVal = resolveEvaluatedCellValue(srcWs, sRaw);
                if (hasValue(sVal)) {
                  const bVal = resolveEvaluatedCellValue(baseWs, bRaw);
                  if (!hasValue(bVal)) {
                    if (sRaw && typeof sRaw === 'object' && 'formula' in sRaw && sVal !== null) {
                      bCell.value = {
                        formula: String((sRaw as { formula?: string }).formula || ''),
                        result: sVal,
                      };
                    } else {
                      bCell.value = sVal;
                    }
                  } else {
                    const bText = String(bVal ?? '').trim();
                    const sText = String(sVal ?? '').trim();
                    const existingParts = bText.split(' | ').map((p) => p.trim());
                    if (!existingParts.includes(sText)) {
                      bCell.value = `${bText} | ${sText}`;
                      applyYellowFill(bCell);
                      conflictCount++;
                    }
                  }
                }
              }
            }
          }
        } catch (err) {
          logs.push(`  ❌ 合并出错 ${sf.originalName}：${(err as Error).message}`);
        }
      }
    }
  }

  // 5. Highlight matched suppliers from 优秀名单
  if (successSet && successSet.size > 0) {
    logs.push('开始根据优秀名单进行高亮标记…');
    for (const ws of baseWb.worksheets) {
      const mergedMap = buildMergedMap(ws);
      let headerRow: number | null = null;
      const cols: Partial<Record<'section' | 'package' | 'respondent', number>> = {};

      for (let r = 1; r <= Math.min(5, ws.rowCount); r++) {
        const rowObj = ws.getRow(r);
        const rowVals: string[] = [];
        for (let c = 1; c <= ws.columnCount; c++) {
          const v = unwrapCellValue(rowObj.getCell(c).value);
          rowVals.push(v === null ? '' : String(v).trim());
        }
        if (rowVals.includes('应答人') || rowVals.includes('标段')) {
          headerRow = r;
          rowVals.forEach((val, idx) => {
            if (val.includes('标段')) cols.section = idx + 1;
            else if (val.includes('包号')) cols.package = idx + 1;
            else if (val.includes('应答人')) cols.respondent = idx + 1;
          });
          break;
        }
      }

      if (!headerRow || !cols.section || !cols.package || !cols.respondent) continue;

      for (let r = headerRow + 1; r <= ws.rowCount; r++) {
        const sec = getMergedCellValue(ws, r, cols.section, mergedMap);
        const pkg = getMergedCellValue(ws, r, cols.package, mergedMap);
        const res = getMergedCellValue(ws, r, cols.respondent, mergedMap);
        if (hasValue(sec) && hasValue(pkg) && hasValue(res)) {
          const key = `${String(sec).trim()}::${String(pkg).trim()}::${String(res).trim()}`;
          if (successSet.has(key)) {
            applyYellowFill(ws.getRow(r).getCell(cols.respondent));
            highlightCount++;
            logs.push(`  高亮：${String(res).trim()} (${String(sec).trim()}-${String(pkg).trim()})`);
          }
        }
      }
    }
  }

  unhideAllRowsAndClearFilter(baseWb);

  const evalPreviewCellText = (ws: ExcelJS.Worksheet, rawVal: ExcelJS.CellValue): string => {
    if (rawVal && typeof rawVal === 'object' && 'formula' in rawVal) {
      const formulaStr = String((rawVal as { formula?: string }).formula || '').trim();
      const evaluated = evalSimpleSumOnSheet(ws, formulaStr);
      if (evaluated !== null) {
        return evaluated.hasConflictText ? `${evaluated.sum} (含冲突)` : String(evaluated.sum);
      }
    }
    const v = unwrapCellValue(rawVal);
    return v === null || v === undefined ? '' : String(v).trim();
  };

  const sheetPreviews: ScoreMergeSheetPreview[] = baseWb.worksheets.map((ws) => {
    const previewLimit = 40;
    const maxRow = Math.min(ws.rowCount, previewLimit);
    let maxUsedCol = 0;
    for (let r = 1; r <= maxRow; r++) {
      const rowObj = ws.getRow(r);
      for (let c = 1; c <= ws.columnCount; c++) {
        const v = unwrapCellValue(rowObj.getCell(c).value);
        if (hasValue(v) && c > maxUsedCol) maxUsedCol = c;
      }
    }
    const colCount = Math.min(Math.max(maxUsedCol, 1), 40);
    const rows: ScoreMergeSheetPreview['rows'] = [];
    for (let r = 1; r <= maxRow; r++) {
      const rowObj = ws.getRow(r);
      const cells: Array<{ text: string; highlighted?: boolean }> = [];
      for (let c = 1; c <= colCount; c++) {
        const cell = rowObj.getCell(c);
        const text = evalPreviewCellText(ws, cell.value);
        const fill = cell.fill;
        const argb =
          fill && fill.type === 'pattern'
            ? String((fill as ExcelJS.FillPattern).fgColor?.argb || '').toUpperCase()
            : '';
        const highlighted = argb.endsWith('FFFF00');
        cells.push(highlighted ? { text, highlighted: true } : { text });
      }
      rows.push({ rowNum: r, cells });
    }
    return {
      name: ws.name,
      totalRows: ws.rowCount,
      colCount,
      headerRows: headerRowsBySheet.get(ws.name) || 0,
      rows,
    };
  });

  const outBuf = Buffer.from(await baseWb.xlsx.writeBuffer());
  const baseTitle = params.templateFile
    ? path.basename(params.templateFile.originalName, path.extname(params.templateFile.originalName))
    : '已合并_汇总表';
  const outputFileName = `${baseTitle}_已合并(含标记).xlsx`;

  logs.push('✅ 处理完成！');
  if (conflictCount > 0) logs.push(`⚠️ 发现数据冲突：${conflictCount} 处（已用黄色标注）`);
  if (highlightCount > 0) logs.push(`✅ 已高亮优秀供应商：${highlightCount} 处`);

  return {
    buffer: outBuf,
    outputFileName,
    totalInputFiles: expanded.length,
    validFileCount: validFiles.length,
    skippedFiles,
    usedTemplate: Boolean(params.templateFile),
    activeRows,
    ignoredRows,
    unmatchedRows,
    conflictCount,
    highlightCount,
    headerSummary,
    sheetPreviews,
    logs,
  };
}
