import path from 'node:path';
import ExcelJS from 'exceljs';
import type { ScoreUploadedFile } from './expertScoreAudit.js';
import { sanitizeLongSheetNamesForExcelJs } from './performanceMatcher.js';

const MAX_RETRY = 500;

export interface RandomScoreSkippedRow {
  sheetName: string;
  row: number;
  section: string;
  packageNo: string;
  supplier: string;
  reason: string;
}

export interface RandomScoreSheetSummary {
  sheetName: string;
  scoreColRange: string;
  objCount: number;
  subjCount: number;
  dataRowCount: number;
  skipCount: number;
}

export interface RandomScoreOutputFile {
  expertName: string;
  filename: string;
  buffer: Buffer;
}

export interface RandomScoreRunResult {
  nExperts: number;
  maxTotalDiff: number;
  maxSameCount: number;
  expertNames: string[];
  sheetSummaries: RandomScoreSheetSummary[];
  skippedRows: RandomScoreSkippedRow[];
  totalCheckedGroups: number;
  violations: number;
  outputs: RandomScoreOutputFile[];
  logs: string[];
}

interface ScoreColParam {
  max: number;
  min: number;
  step: number;
  values: number[];
}

interface SheetMeta {
  sumCol: number;
  expertCol: number | null;
  scoreCols: number[];
  objCols: Set<number>;
  subjCols: number[];
  params: Map<number, ScoreColParam>;
  dataRows: number[];
}

function unwrapCellValue(val: ExcelJS.CellValue): string | number | boolean | null {
  if (val === null || val === undefined) return null;
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return val;
  if (typeof val === 'object') {
    if ('result' in val && val.result !== undefined) {
      return unwrapCellValue(val.result as ExcelJS.CellValue);
    }
    if ('richText' in val && Array.isArray((val as { richText?: Array<{ text?: string }> }).richText)) {
      return (val as { richText: Array<{ text?: string }> }).richText.map((r) => r.text || '').join('');
    }
    if ('text' in val && typeof (val as { text?: unknown }).text === 'string') {
      return (val as { text: string }).text;
    }
  }
  return String(val);
}

function colNumToLetters(col: number): string {
  let n = col;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function isYellowCell(cell: ExcelJS.Cell): boolean {
  const fill = cell.fill;
  if (!fill || fill.type !== 'pattern') return false;
  const fg = fill.fgColor as { argb?: string } | undefined;
  const bg = fill.bgColor as { argb?: string } | undefined;
  const fgArgb = fg?.argb ? String(fg.argb).toUpperCase() : '';
  const bgArgb = bg?.argb ? String(bg.argb).toUpperCase() : '';
  return fgArgb.includes('FFFF00') || bgArgb.includes('FFFF00');
}

function round4(n: number): number {
  return Math.round(n * 10000) / 10000;
}

function buildPossibleValues(min: number, max: number, step: number): number[] {
  const vals: number[] = [];
  let v = min;
  let guard = 0;
  while (v <= max + 1e-9 && guard < 10000) {
    vals.push(round4(v));
    v += step;
    guard++;
  }
  return vals;
}

function isNumericText(s: string): boolean {
  const cleaned = s.replace(/[.\-\s]/g, '');
  return cleaned.length > 0 && /^\d+$/.test(cleaned);
}

function autoDetectStructure(ws: ExcelJS.Worksheet): SheetMeta | null {
  let sumRow: number | null = null;
  let sumCol: number | null = null;
  let expertCol: number | null = null;

  const scanRows = Math.min(20, ws.rowCount);
  for (let r = 1; r <= scanRows; r++) {
    const rowObj = ws.getRow(r);
    for (let c = 1; c <= ws.columnCount; c++) {
      const v = unwrapCellValue(rowObj.getCell(c).value);
      const text = v === null ? '' : String(v).trim();
      if (text === '合计') {
        sumRow = r;
        sumCol = c;
      }
      if (text === '专家名称') {
        expertCol = c;
      }
    }
  }

  if (sumRow === null || sumCol === null) return null;
  const stepRow = sumRow;
  const minRow = stepRow - 1;
  const maxRow = stepRow - 2;
  if (maxRow < 1 || minRow < 1) return null;

  const scoreCols: number[] = [];
  for (let c = 7; c < sumCol; c++) scoreCols.push(c);
  if (scoreCols.length === 0) return null;

  const dataStart = stepRow + 1;
  const dataRows: number[] = [];
  for (let r = dataStart; r <= ws.rowCount; r++) {
    // Skip completely empty trailing rows
    const rowObj = ws.getRow(r);
    const hasMeta =
      unwrapCellValue(rowObj.getCell(2).value) !== null ||
      unwrapCellValue(rowObj.getCell(3).value) !== null ||
      unwrapCellValue(rowObj.getCell(5).value) !== null;
    if (hasMeta) dataRows.push(r);
  }
  if (dataRows.length === 0) return null;

  // Detect yellow objective columns (prefer bottom-most non-numeric text header row in [2..sumRow])
  const objCols = new Set<number>();
  for (const c of scoreCols) {
    let foundText = false;
    for (let r = sumRow; r >= 2; r--) {
      const cell = ws.getRow(r).getCell(c);
      const v = unwrapCellValue(cell.value);
      if (v === null || String(v).trim() === '') continue;
      if (!isNumericText(String(v).trim())) {
        if (isYellowCell(cell)) objCols.add(c);
        foundText = true;
        break;
      }
    }
    if (!foundText) {
      for (let r = sumRow; r >= 2; r--) {
        const cell = ws.getRow(r).getCell(c);
        const v = unwrapCellValue(cell.value);
        if (v !== null && String(v).trim() !== '') {
          if (isYellowCell(cell)) objCols.add(c);
          break;
        }
      }
    }
  }

  const subjCols = scoreCols.filter((c) => !objCols.has(c));

  const params = new Map<number, ScoreColParam>();
  for (const c of scoreCols) {
    const mxRaw = unwrapCellValue(ws.getRow(maxRow).getCell(c).value);
    const mnRaw = unwrapCellValue(ws.getRow(minRow).getCell(c).value);
    const stRaw = unwrapCellValue(ws.getRow(stepRow).getCell(c).value);
    if (mxRaw === null || mnRaw === null || stRaw === null) continue;
    let mx = Number(mxRaw);
    let mn = Number(mnRaw);
    const st = Number(stRaw);
    if (!Number.isFinite(mx) || !Number.isFinite(mn) || !Number.isFinite(st) || st <= 0) continue;
    if (mx < mn) [mx, mn] = [mn, mx];
    params.set(c, { max: mx, min: mn, step: st, values: buildPossibleValues(mn, mx, st) });
  }

  // Fallback for templates without numeric 最高分/最低分/级差 rows (e.g. 4-row text-criteria header templates)
  if (params.size === 0) {
    for (const c of scoreCols) {
      const criteriaRaw = unwrapCellValue(ws.getRow(sumRow).getCell(c).value);
      const criteriaText = criteriaRaw === null ? '' : String(criteriaRaw);
      const titleRaw = sumRow > 1 ? unwrapCellValue(ws.getRow(sumRow - 1).getCell(c).value) : null;
      const titleText = titleRaw === null ? '' : String(titleRaw);

      const extracted: number[] = [];
      // 1) Range expressions in criteria text, e.g. "13-15分", "9~10分", "5—6分"
      const rangeRe = /(\d+(?:\.\d+)?)\s*[-~～至—–]\s*(\d+(?:\.\d+)?)\s*分/g;
      let m: RegExpExecArray | null;
      while ((m = rangeRe.exec(criteriaText)) !== null) {
        const a = Number(m[1]);
        const b = Number(m[2]);
        if (Number.isFinite(a)) extracted.push(a);
        if (Number.isFinite(b)) extracted.push(b);
      }
      // 2) Discrete score mentions in criteria text, e.g. "得5分", "得【2分】", "。5分"
      const singleRe = /(?:得|为|计|[：:。；;])\s*【?\s*(\d+(?:\.\d+)?)\s*分\s*】?/g;
      while ((m = singleRe.exec(criteriaText)) !== null) {
        const v = Number(m[1]);
        if (Number.isFinite(v)) extracted.push(v);
      }

      let mn: number | null = null;
      let mx: number | null = null;
      if (extracted.length >= 2) {
        mn = Math.min(...extracted);
        mx = Math.max(...extracted);
      } else {
        // Fallback: parse "(X分)" or "（X分）" from title row
        const titleMatch = /[（(]\s*(\d+(?:\.\d+)?)\s*分\s*[）)]/.exec(titleText);
        if (titleMatch) {
          const cap = Number(titleMatch[1]);
          if (Number.isFinite(cap) && cap > 0) {
            mn = 0;
            mx = cap;
          }
        }
      }

      if (mn === null || mx === null || mx <= mn) continue;

      // Determine step: 0.5 if any extracted bound or observed score has a fractional part, otherwise 1
      let hasHalf = extracted.some((n) => Math.abs(n - Math.round(n)) > 1e-6);
      if (!hasHalf) {
        for (const r of dataRows) {
          const v = unwrapCellValue(ws.getRow(r).getCell(c).value);
          const num = typeof v === 'number' ? v : Number(String(v ?? '').trim());
          if (Number.isFinite(num) && Math.abs(num - Math.round(num)) > 1e-6) {
            hasHalf = true;
            break;
          }
        }
      }
      const st = hasHalf ? 0.5 : 1;
      params.set(c, { max: mx, min: mn, step: st, values: buildPossibleValues(mn, mx, st) });
    }
  }

  return { sumCol, expertCol, scoreCols, objCols, subjCols, params, dataRows };
}

export async function parseExpertNamesFromExcel(buffer: Buffer): Promise<string[]> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(sanitizeLongSheetNamesForExcelJs(buffer) as unknown as Parameters<typeof wb.xlsx.load>[0]);
  const ws = wb.worksheets[0];
  if (!ws) return [];
  const headerSkip = new Set(['姓名', '专家姓名', '专家名称', '名称', 'name', 'Name']);
  const names: string[] = [];
  for (let r = 1; r <= ws.rowCount; r++) {
    const v = unwrapCellValue(ws.getRow(r).getCell(1).value);
    if (v === null) continue;
    const s = String(v).trim();
    if (s && !headerSkip.has(s)) names.push(s);
  }
  return names;
}

export async function runRandomScoreGenerator(params: {
  templateFile: ScoreUploadedFile;
  expertNames?: string[] | null;
  nExperts?: number;
  maxTotalDiff?: number;
}): Promise<RandomScoreRunResult> {
  if (!/\.xlsx$/i.test(params.templateFile.originalName)) {
    throw new Error('打分表模板请使用 .xlsx 格式。');
  }

  const maxTotalDiff =
    params.maxTotalDiff !== undefined && Number.isFinite(params.maxTotalDiff) && params.maxTotalDiff >= 0
      ? params.maxTotalDiff
      : 7;

  let expertNames: string[];
  let nExperts: number;
  if (params.expertNames && params.expertNames.length > 0) {
    expertNames = params.expertNames.map((s) => s.trim()).filter(Boolean);
    nExperts = expertNames.length + 1;
  } else {
    nExperts = Math.floor(Number(params.nExperts || 5));
    if (!Number.isFinite(nExperts) || nExperts < 2) {
      throw new Error('专家总数至少为 2（含已有的 1 位专家）。');
    }
    if (nExperts > 30) {
      throw new Error('单次专家总数不能超过 30。');
    }
    expertNames = Array.from({ length: nExperts - 1 }, (_, i) => `专家${i + 2}`);
  }

  const nNew = nExperts - 1;
  const maxSame = Math.floor(nExperts / 2);
  const logs: string[] = [
    `源文件：${params.templateFile.originalName}`,
    `生成专家：${expertNames.join('、')}`,
    `共 ${nExperts} 位专家（含原表 1 位），生成 ${nNew} 份，总分最大分差 ≤ ${maxTotalDiff}，同分上限 ≤ ${maxSame} 人`,
  ];

  const safeTemplateBuf = sanitizeLongSheetNamesForExcelJs(params.templateFile.buffer);
  const wbOrig = new ExcelJS.Workbook();
  await wbOrig.xlsx.load(safeTemplateBuf as unknown as Parameters<typeof wbOrig.xlsx.load>[0]);

  const sheetMetas = new Map<string, SheetMeta>();
  const origScores = new Map<string, number>();
  const origTotals = new Map<string, number>();
  const skipRowsMap = new Map<string, RandomScoreSkippedRow>();
  const sheetSummaries: RandomScoreSheetSummary[] = [];

  for (const ws of wbOrig.worksheets) {
    const meta = autoDetectStructure(ws);
    if (!meta) {
      logs.push(`  [跳过] 工作表「${ws.name}」未找到合计列或数据行`);
      continue;
    }
    sheetMetas.set(ws.name, meta);
    let skipCount = 0;

    for (const r of meta.dataRows) {
      const rowObj = ws.getRow(r);
      const section = String(unwrapCellValue(rowObj.getCell(2).value) ?? '').trim();
      const packageNo = String(unwrapCellValue(rowObj.getCell(3).value) ?? '').trim();
      const supplier = String(unwrapCellValue(rowObj.getCell(5).value) ?? '').trim();
      const reviewVal = unwrapCellValue(rowObj.getCell(6).value);
      const rowKey = `${ws.name}::${r}`;

      if (reviewVal !== null && String(reviewVal).includes('否')) {
        skipRowsMap.set(rowKey, {
          sheetName: ws.name,
          row: r,
          section,
          packageNo,
          supplier,
          reason: '符合性评审未通过',
        });
        skipCount++;
      }

      let rowTotal = 0;
      for (const c of meta.scoreCols) {
        const raw = unwrapCellValue(rowObj.getCell(c).value);
        let num = 0;
        if (typeof raw === 'number' && Number.isFinite(raw)) {
          num = raw;
        } else if (raw !== null && raw !== '') {
          const parsed = Number(String(raw).trim());
          num = Number.isFinite(parsed) ? parsed : 0;
        }
        origScores.set(`${ws.name}::${r}::${c}`, num);
        rowTotal += num;
      }
      origTotals.set(rowKey, round4(rowTotal));

      if (!skipRowsMap.has(rowKey)) {
        for (const c of meta.subjCols) {
          const p = meta.params.get(c);
          if (!p) continue;
          const ov = origScores.get(`${ws.name}::${r}::${c}`) ?? 0;
          if (ov < p.min - 1e-9 || ov > p.max + 1e-9) {
            let colName = '';
            for (let hr = meta.dataRows[0] - 1; hr >= 2; hr--) {
              const cv = unwrapCellValue(ws.getRow(hr).getCell(c).value);
              if (cv !== null && String(cv).trim() !== '' && !isNumericText(String(cv).trim())) {
                colName = String(cv).trim();
                break;
              }
            }
            skipRowsMap.set(rowKey, {
              sheetName: ws.name,
              row: r,
              section,
              packageNo,
              supplier,
              reason: `异常分数: ${colNumToLetters(c)}列「${colName}」=${ov}，有效范围 [${p.min}, ${p.max}]`,
            });
            skipCount++;
            break;
          }
        }
      }
    }

    const scoreRange = `${colNumToLetters(meta.scoreCols[0])}-${colNumToLetters(
      meta.scoreCols[meta.scoreCols.length - 1]
    )}`;
    sheetSummaries.push({
      sheetName: ws.name,
      scoreColRange: scoreRange,
      objCount: meta.objCols.size,
      subjCount: meta.subjCols.length,
      dataRowCount: meta.dataRows.length,
      skipCount,
    });
    logs.push(
      `  [识别]「${ws.name}」打分列 ${scoreRange}，客观 ${meta.objCols.size} 列，主观 ${meta.subjCols.length} 列，数据 ${meta.dataRows.length} 行` +
        (skipCount > 0 ? `，跳过 ${skipCount} 行` : '')
    );
  }

  if (sheetMetas.size === 0) {
    throw new Error('未能识别任何工作表的打分结构（需包含「合计」列及上方的最高分/最低分/级差行）。');
  }

  // Generate random scores for nNew experts
  const allExpertScores: Array<Map<string, number>> = Array.from({ length: nNew }, () => new Map());
  const allExpertTotals: Array<Map<string, number>> = Array.from({ length: nNew }, () => new Map());

  for (const [sn, meta] of sheetMetas.entries()) {
    for (const r of meta.dataRows) {
      const rowKey = `${sn}::${r}`;
      const origT = origTotals.get(rowKey) ?? 0;

      if (skipRowsMap.has(rowKey)) {
        for (let ei = 0; ei < nNew; ei++) {
          for (const c of meta.scoreCols) {
            allExpertScores[ei].set(`${sn}::${r}::${c}`, origScores.get(`${sn}::${r}::${c}`) ?? 0);
          }
          allExpertTotals[ei].set(rowKey, origT);
        }
        continue;
      }

      for (let ei = 0; ei < nNew; ei++) {
        const prevTotals: number[] = [origT];
        for (let k = 0; k < ei; k++) {
          prevTotals.push(allExpertTotals[k].get(rowKey) ?? origT);
        }

        let found = false;
        let lastScores = new Map<number, number>();

        for (let attempt = 0; attempt < MAX_RETRY; attempt++) {
          const radius =
            attempt < Math.floor(MAX_RETRY / 3)
              ? 3
              : attempt < Math.floor((MAX_RETRY * 2) / 3)
                ? 5
                : null;

          const curScores = new Map<number, number>();
          let total = 0;

          for (const c of meta.scoreCols) {
            const origV = origScores.get(`${sn}::${r}::${c}`) ?? 0;
            let chosen = origV;
            if (!meta.objCols.has(c)) {
              const p = meta.params.get(c);
              if (p && origV >= p.min - 1e-9 && origV <= p.max + 1e-9 && p.values.length > 0) {
                let candidates = p.values;
                if (radius !== null) {
                  const nearby = p.values.filter((v) => Math.abs(v - origV) <= p.step * radius + 1e-9);
                  if (nearby.length > 0) candidates = nearby;
                }
                chosen = candidates[Math.floor(Math.random() * candidates.length)];
              }
            }
            curScores.set(c, chosen);
            total += chosen;
          }

          total = round4(total);
          lastScores = curScores;

          if (prevTotals.some((pt) => Math.abs(total - pt) > maxTotalDiff + 1e-9)) {
            continue;
          }
          const sameCount = prevTotals.filter((pt) => Math.abs(pt - total) < 1e-9).length + 1;
          if (sameCount > maxSame) {
            continue;
          }

          for (const [c, v] of curScores.entries()) {
            allExpertScores[ei].set(`${sn}::${r}::${c}`, v);
          }
          allExpertTotals[ei].set(rowKey, total);
          found = true;
          break;
        }

        if (!found) {
          let fallbackTotal = 0;
          for (const [c, v] of lastScores.entries()) {
            allExpertScores[ei].set(`${sn}::${r}::${c}`, v);
            fallbackTotal += v;
          }
          allExpertTotals[ei].set(rowKey, round4(fallbackTotal));
          logs.push(`  [提示]「${sn}」第 ${r} 行主观分调整空间不足，可能无法完全满足所有约束`);
        }
      }
    }
  }

  // Verify rules
  let violations = 0;
  let totalCheckedGroups = 0;
  for (const [sn, meta] of sheetMetas.entries()) {
    for (const r of meta.dataRows) {
      const rowKey = `${sn}::${r}`;
      if (skipRowsMap.has(rowKey)) continue;
      totalCheckedGroups++;
      const totals = [origTotals.get(rowKey) ?? 0];
      for (let ei = 0; ei < nNew; ei++) {
        totals.push(allExpertTotals[ei].get(rowKey) ?? 0);
      }
      const diff = round4(Math.max(...totals) - Math.min(...totals));
      if (diff > maxTotalDiff + 1e-9) {
        violations++;
        logs.push(`  ✗ 规则1违规 [${sn}] 第${r}行: 总分差=${diff} > ${maxTotalDiff}`);
      }
      const freq = new Map<number, number>();
      for (const t of totals) freq.set(t, (freq.get(t) || 0) + 1);
      for (const [t, cnt] of freq.entries()) {
        if (cnt > maxSame) {
          violations++;
          logs.push(`  ✗ 规则2违规 [${sn}] 第${r}行: 总分 ${t} 出现 ${cnt} 次 > ${maxSame}`);
        }
      }
    }
  }

  // Write nNew output workbooks preserving original template formatting
  const baseStem = path.basename(params.templateFile.originalName, path.extname(params.templateFile.originalName));
  const outputs: RandomScoreOutputFile[] = [];

  for (let ei = 0; ei < nNew; ei++) {
    const label = expertNames[ei];
    const safeLabel = label.replace(/[\\/:*?"<>|]/g, '_');
    const filename = `${baseStem}_${safeLabel}.xlsx`;

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(safeTemplateBuf as unknown as Parameters<typeof outWb.xlsx.load>[0]);

    for (const ws of outWb.worksheets) {
      const meta = sheetMetas.get(ws.name);
      if (!meta) continue;
      const firstCol = colNumToLetters(meta.scoreCols[0]);
      const lastCol = colNumToLetters(meta.scoreCols[meta.scoreCols.length - 1]);

      for (const r of meta.dataRows) {
        const rowObj = ws.getRow(r);
        if (meta.expertCol) {
          rowObj.getCell(meta.expertCol).value = label;
        }
        for (const c of meta.scoreCols) {
          const val = allExpertScores[ei].get(`${ws.name}::${r}::${c}`);
          if (val !== undefined) {
            rowObj.getCell(c).value = val;
          }
        }
        rowObj.getCell(meta.sumCol).value = {
          formula: `SUM(${firstCol}${r}:${lastCol}${r})`,
        };
      }
    }

    const buf = Buffer.from(await outWb.xlsx.writeBuffer());
    outputs.push({ expertName: label, filename, buffer: buf });
  }

  logs.push(
    `规则校验：共核验 ${totalCheckedGroups} 组应答文件 × ${nExperts} 位专家，` +
      (violations === 0 ? '全部通过（0 违规）' : `发现 ${violations} 处违规`)
  );

  return {
    nExperts,
    maxTotalDiff,
    maxSameCount: maxSame,
    expertNames,
    sheetSummaries,
    skippedRows: Array.from(skipRowsMap.values()),
    totalCheckedGroups,
    violations,
    outputs,
    logs,
  };
}
