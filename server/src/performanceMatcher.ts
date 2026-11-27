import path from 'node:path';
import ExcelJS from 'exceljs';
import PizZip from 'pizzip';
import { expandPrecheckFiles } from './expertPrecheckAudit.js';
import type { ScoreUploadedFile } from './expertScoreAudit.js';

const PERF_KEYWORD = '绩效评价';
const HEADER_BIAODUAN = new Set(['标段', '分标名称', '标段名称']);
const HEADER_BAO = new Set(['包号', '分包名称', '包']);
const HEADER_SUPPLIER = new Set(['应答人', '投标人', '投标人名称', '供应商', '供应商名称']);
const HEADER_SCAN_ROWS = 20;
const PERF_SCAN_ROWS = 10;
const PERF_LABEL_MAX_LEN = 50;

export interface PerfMatchUnmatchedItem {
  templateName: string;
  sheetName: string;
  section: string;
  packageNo: string;
  supplier: string;
}

export interface PerfMatchTemplateResult {
  originalName: string;
  outputFileName: string;
  buffer: Buffer;
  sheetsProcessed: number;
  sheetsSkipped: string[];
  totalRows: number;
  matched: number;
  defaulted: number;
  unmatched: PerfMatchUnmatchedItem[];
}

export interface PerfMatchRunResult {
  lyfFileName: string;
  lyfRecordCount: number;
  lyfDuplicates: number;
  defaultScore: number;
  templates: PerfMatchTemplateResult[];
  totalRows: number;
  totalMatched: number;
  totalDefaulted: number;
  totalSkippedSheets: number;
  reportText: string;
  logs: string[];
}

export function sanitizeLongSheetNamesForExcelJs(buffer: Buffer): Buffer {
  try {
    const zip = new PizZip(buffer);
    const wbEntry = zip.file('xl/workbook.xml');
    if (!wbEntry) return buffer;
    const xml = wbEntry.asText();
    const sheetTagRe = /<sheet\b[^>]*\bname="([^"]+)"[^>]*>/g;
    const matches: Array<{ full: string; rawName: string; index: number }> = [];
    let m: RegExpExecArray | null;
    while ((m = sheetTagRe.exec(xml)) !== null) {
      matches.push({ full: m[0], rawName: m[1], index: m.index });
    }
    if (matches.length === 0) return buffer;

    const hasLongOrDup =
      matches.some((item) => item.rawName.length > 31) ||
      new Set(matches.map((item) => item.rawName)).size !== matches.length;
    if (!hasLongOrDup) return buffer;

    const used = new Set<string>();
    const renamePairs: Array<[string, string]> = [];
    let rebuilt = '';
    let cursor = 0;
    for (const item of matches) {
      rebuilt += xml.slice(cursor, item.index);
      let nextName = item.rawName.slice(0, 31);
      if (used.has(nextName)) {
        let seq = 2;
        while (true) {
          const suffix = `_${seq}`;
          const cand = `${item.rawName.slice(0, 31 - suffix.length)}${suffix}`;
          if (!used.has(cand)) {
            nextName = cand;
            break;
          }
          seq++;
        }
      }
      used.add(nextName);
      if (nextName !== item.rawName) {
        renamePairs.push([item.rawName, nextName]);
      }
      const replacedTag = item.full.replace(/\bname="[^"]+"/, `name="${nextName}"`);
      rebuilt += replacedTag;
      cursor = item.index + item.full.length;
    }
    rebuilt += xml.slice(cursor);
    for (const [oldName, newName] of renamePairs) {
      rebuilt = rebuilt.split(`'${oldName}'!`).join(`'${newName}'!`);
    }
    zip.file('xl/workbook.xml', rebuilt);
    return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }) as Buffer;
  } catch {
    return buffer;
  }
}

function unwrapCellValue(val: ExcelJS.CellValue): string | number | boolean | null {
  if (val === null || val === undefined) return null;
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return val;
  if (val instanceof Date) return val.toISOString();
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

function normStr(val: unknown): string {
  if (val === null || val === undefined) return '';
  return String(val).trim();
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

interface SheetLayout {
  headerRow: number;
  colBiaoduan: number;
  colBao: number;
  colSupplier: number;
  perfCol: number;
  dataStart: number;
}

function detectSheetLayout(ws: ExcelJS.Worksheet): SheetLayout | null {
  const maxRows = ws.rowCount || 0;
  if (maxRows === 0) return null;

  let headerRow: number | null = null;
  let colBiaoduan: number | null = null;
  let colBao: number | null = null;
  let colSupplier: number | null = null;

  const scanLimit = Math.min(HEADER_SCAN_ROWS, maxRows);
  for (let r = 1; r <= scanLimit; r++) {
    const rowObj = ws.getRow(r);
    let bCol: number | null = null;
    let pCol: number | null = null;
    let sCol: number | null = null;

    for (let c = 1; c <= ws.columnCount; c++) {
      const v = normStr(unwrapCellValue(rowObj.getCell(c).value));
      if (!v) continue;
      if (bCol === null && HEADER_BIAODUAN.has(v)) bCol = c;
      if (pCol === null && HEADER_BAO.has(v)) pCol = c;
      if (sCol === null && HEADER_SUPPLIER.has(v)) sCol = c;
    }
    if (bCol !== null && sCol !== null) {
      headerRow = r;
      colBiaoduan = bCol;
      colSupplier = sCol;
      colBao = pCol;
      break;
    }
  }

  if (headerRow === null || colBiaoduan === null || colSupplier === null) return null;
  if (colBao === null) colBao = colBiaoduan + 1;

  // Step 2: find "绩效评价" column within next PERF_SCAN_ROWS rows (label length <= 50)
  let perfCol: number | null = null;
  const perfEnd = Math.min(headerRow + PERF_SCAN_ROWS, maxRows);
  for (let r = headerRow + 1; r <= perfEnd; r++) {
    const rowObj = ws.getRow(r);
    for (let c = 1; c <= ws.columnCount; c++) {
      const raw = unwrapCellValue(rowObj.getCell(c).value);
      if (raw === null || raw === '') continue;
      const sv = String(raw);
      if (sv.includes(PERF_KEYWORD) && sv.length <= PERF_LABEL_MAX_LEN) {
        perfCol = c;
        break;
      }
    }
    if (perfCol !== null) break;
  }
  if (perfCol === null) return null;

  // Step 3: first row after headerRow where colBiaoduan is non-empty (excluding vertical-merge slave cells of the header)
  let dataStart: number | null = null;
  for (let r = headerRow + 1; r <= maxRows; r++) {
    const cell = ws.getRow(r).getCell(colBiaoduan);
    if (cell.isMerged && cell.master && cell.master.address !== cell.address) {
      continue;
    }
    const val = normStr(unwrapCellValue(cell.value));
    if (val && !HEADER_BIAODUAN.has(val)) {
      dataStart = r;
      break;
    }
  }
  if (dataStart === null) return null;

  return { headerRow, colBiaoduan, colBao, colSupplier, perfCol, dataStart };
}

async function loadLyfMap(
  lyfFile: ScoreUploadedFile,
  logs: string[]
): Promise<{ map: Map<string, number | string>; duplicates: number }> {
  logs.push(`读取履约分文件：${lyfFile.originalName}`);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(lyfFile.buffer as unknown as Parameters<typeof wb.xlsx.load>[0]);
  const ws = wb.worksheets[0];
  if (!ws || ws.rowCount < 2) {
    throw new Error(`履约分文件「${lyfFile.originalName}」为空或缺少数据行。`);
  }

  // Try header-based column detection on row 1, fallback to standard columns B(2), C(3), D(4), E(5)
  const hRow = ws.getRow(1);
  let colB = 0;
  let colP = 0;
  let colS = 0;
  let colScore = 0;
  for (let c = 1; c <= ws.columnCount; c++) {
    const h = normStr(unwrapCellValue(hRow.getCell(c).value));
    if (!h) continue;
    if (!colB && (h === '分标名称' || h === '标段' || h === '标段名称')) colB = c;
    if (!colP && (h === '分包名称' || h === '包号')) colP = c;
    if (!colS && (h === '投标人' || h === '投标人名称' || h === '应答人' || h === '供应商' || h === '供应商名称')) {
      colS = c;
    }
    if (!colScore && (h === '履约分' || h === '履约分结果' || h.includes('履约分'))) colScore = c;
  }
  if (!colB || !colP || !colS || !colScore) {
    // Fallback to matcher.py default 0-indexed (1, 2, 3, 4) => 1-based (2, 3, 4, 5)
    colB = 2;
    colP = 3;
    colS = 4;
    colScore = 5;
  }

  const map = new Map<string, number | string>();
  let duplicates = 0;

  for (let r = 2; r <= ws.rowCount; r++) {
    const rowObj = ws.getRow(r);
    const bVal = unwrapCellValue(rowObj.getCell(colB).value);
    if (bVal === null || bVal === '') continue;
    const key = `${normStr(bVal)}::${normStr(unwrapCellValue(rowObj.getCell(colP).value))}::${normStr(
      unwrapCellValue(rowObj.getCell(colS).value)
    )}`;
    const rawScore = unwrapCellValue(rowObj.getCell(colScore).value);
    let scoreVal: number | string = '';
    if (typeof rawScore === 'number' && Number.isFinite(rawScore)) {
      scoreVal = rawScore;
    } else if (rawScore !== null && rawScore !== '') {
      const s = String(rawScore).trim();
      const n = Number(s);
      scoreVal = s !== '' && Number.isFinite(n) ? n : s;
    }

    if (map.has(key) && map.get(key) !== scoreVal) {
      duplicates++;
    }
    map.set(key, scoreVal);
  }

  logs.push(
    `  共载入 ${map.size} 条履约分记录` +
      (duplicates > 0 ? `（其中 ${duplicates} 条重复键值不一致，已用最后一条覆盖）` : '')
  );
  return { map, duplicates };
}

function buildCombinedReportText(
  templates: PerfMatchTemplateResult[],
  defaultScore: number
): string {
  const gTotal = templates.reduce((s, t) => s + t.totalRows, 0);
  const gMatched = templates.reduce((s, t) => s + t.matched, 0);
  const gDefault = templates.reduce((s, t) => s + t.defaulted, 0);
  const gSkipped = templates.reduce((s, t) => s + t.sheetsSkipped.length, 0);

  const lines: string[] = [
    '履约分匹配 - 汇总报告',
    '='.repeat(60),
    `共处理 ${templates.length} 个详评打分模板。`,
    '',
    `总待打分行数：${gTotal}`,
    `已匹配履约分：${gMatched}`,
    `未匹配（已默认填 ${defaultScore} 分）：${gDefault}`,
    `跳过 sheet 总数：${gSkipped}`,
    '',
  ];

  templates.forEach((t, idx) => {
    lines.push('-'.repeat(60));
    lines.push(`[${idx + 1}] 模板：${t.originalName}`);
    lines.push(`    输出：${t.outputFileName}`);
    lines.push(`    处理 sheet 数：${t.sheetsProcessed}`);
    lines.push(`    待打分行总数：${t.totalRows}`);
    lines.push(`    已匹配履约分：${t.matched}`);
    lines.push(`    未匹配（已默认填 ${defaultScore} 分）：${t.defaulted}`);
    if (t.sheetsSkipped.length > 0) {
      lines.push(`    跳过的 sheet（未找到「${PERF_KEYWORD}」列）：`);
      for (const s of t.sheetsSkipped) lines.push(`      - ${s}`);
    }
    lines.push('');
  });

  const allUnmatched = templates.flatMap((t) => t.unmatched);
  if (allUnmatched.length > 0) {
    lines.push('='.repeat(60));
    lines.push(`未匹配明细（共 ${allUnmatched.length} 条，已按默认 ${defaultScore} 分填入）`);
    lines.push('模板\tsheet\t标段\t包号\t应答人');
    for (const u of allUnmatched) {
      lines.push(`${u.templateName}\t${u.sheetName}\t${u.section}\t${u.packageNo}\t${u.supplier}`);
    }
  }

  return lines.join('\n');
}

export async function runPerformanceMatch(params: {
  lyfFile: ScoreUploadedFile;
  templateFiles: ScoreUploadedFile[];
  defaultScore?: number;
}): Promise<PerfMatchRunResult> {
  const defaultScore =
    params.defaultScore !== undefined && Number.isFinite(params.defaultScore)
      ? params.defaultScore
      : 4;

  const expandedTemplates = expandPrecheckFiles(params.templateFiles).filter((f) =>
    /\.xlsx$/i.test(f.originalName)
  );
  if (expandedTemplates.length === 0) {
    throw new Error('请至少上传一份详评打分模板（.xlsx 或包含 .xlsx 的 .zip 压缩包）。');
  }
  if (!/\.xlsx$/i.test(params.lyfFile.originalName)) {
    throw new Error('履约分文件请使用 .xlsx 格式。');
  }

  const logs: string[] = [];
  const { map: lyfMap, duplicates } = await loadLyfMap(params.lyfFile, logs);
  const templates: PerfMatchTemplateResult[] = [];

  for (const tpl of expandedTemplates) {
    logs.push(`读取详评打分模板：${tpl.originalName}`);
    const safeBuf = sanitizeLongSheetNamesForExcelJs(tpl.buffer);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(safeBuf as unknown as Parameters<typeof wb.xlsx.load>[0]);

    let sheetsProcessed = 0;
    const sheetsSkipped: string[] = [];
    let totalRows = 0;
    let matched = 0;
    let defaulted = 0;
    const unmatched: PerfMatchUnmatchedItem[] = [];

    for (const ws of wb.worksheets) {
      const layout = detectSheetLayout(ws);
      if (!layout) {
        logs.push(`  [跳过] sheet「${ws.name}」未识别到「标段/应答人」表头或「${PERF_KEYWORD}」列`);
        sheetsSkipped.push(ws.name);
        continue;
      }

      const shortName = ws.name.length <= 30 ? ws.name : `${ws.name.slice(0, 30)}...`;
      logs.push(
        `  处理 sheet「${shortName}」表头行=${layout.headerRow} 数据起始行=${layout.dataStart} ` +
          `标段列=${colNumToLetters(layout.colBiaoduan)} 包号列=${colNumToLetters(layout.colBao)} ` +
          `应答人列=${colNumToLetters(layout.colSupplier)} 绩效评价列=${colNumToLetters(layout.perfCol)}`
      );

      let sheetTotal = 0;
      let sheetMatched = 0;
      let sheetDefault = 0;

      for (let r = layout.dataStart; r <= ws.rowCount; r++) {
        const rowObj = ws.getRow(r);
        const biaoduan = normStr(unwrapCellValue(rowObj.getCell(layout.colBiaoduan).value));
        const bao = normStr(unwrapCellValue(rowObj.getCell(layout.colBao).value));
        const supplier = normStr(unwrapCellValue(rowObj.getCell(layout.colSupplier).value));
        if (!biaoduan && !bao && !supplier) continue;

        sheetTotal++;
        const key = `${biaoduan}::${bao}::${supplier}`;
        if (lyfMap.has(key)) {
          rowObj.getCell(layout.perfCol).value = lyfMap.get(key)!;
          sheetMatched++;
        } else {
          rowObj.getCell(layout.perfCol).value = defaultScore;
          sheetDefault++;
          unmatched.push({
            templateName: tpl.originalName,
            sheetName: ws.name,
            section: biaoduan,
            packageNo: bao,
            supplier,
          });
        }
      }

      logs.push(
        `    -> 总 ${sheetTotal} 行，匹配 ${sheetMatched}，默认填 ${defaultScore} 分 ${sheetDefault} 条`
      );
      sheetsProcessed++;
      totalRows += sheetTotal;
      matched += sheetMatched;
      defaulted += sheetDefault;
    }

    const stem = path.basename(tpl.originalName, path.extname(tpl.originalName));
    const outputFileName = `${stem}_已填履约分.xlsx`;
    const outBuf = Buffer.from(await wb.xlsx.writeBuffer());
    templates.push({
      originalName: tpl.originalName,
      outputFileName,
      buffer: outBuf,
      sheetsProcessed,
      sheetsSkipped,
      totalRows,
      matched,
      defaulted,
      unmatched,
    });
  }

  const reportText = buildCombinedReportText(templates, defaultScore);
  const totalRows = templates.reduce((s, t) => s + t.totalRows, 0);
  const totalMatched = templates.reduce((s, t) => s + t.matched, 0);
  const totalDefaulted = templates.reduce((s, t) => s + t.defaulted, 0);
  const totalSkippedSheets = templates.reduce((s, t) => s + t.sheetsSkipped.length, 0);

  logs.push(
    `✅ 全部完成：共 ${templates.length} 个模板，待打分 ${totalRows} 行，匹配 ${totalMatched} 行，默认填 ${defaultScore} 分 ${totalDefaulted} 条。`
  );

  return {
    lyfFileName: params.lyfFile.originalName,
    lyfRecordCount: lyfMap.size,
    lyfDuplicates: duplicates,
    defaultScore,
    templates,
    totalRows,
    totalMatched,
    totalDefaulted,
    totalSkippedSheets,
    reportText,
    logs,
  };
}
