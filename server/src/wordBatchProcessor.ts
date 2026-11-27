import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import PizZip from 'pizzip';
import { Dec } from './toolboxWorkbook.js';

export interface WordBatchReplacePair {
  find: string;
  replace: string;
}

export interface WordBatchConfig {
  modifyEnabled: boolean;
  modifyTemplate: string;
  multiplier: string;
  deleteEnabled: boolean;
  deleteTemplate: string;
  replaceEnabled: boolean;
  replaceUseRegex: boolean;
  replacePairs: WordBatchReplacePair[];
  extractEnabled: boolean;
}

export interface WordBatchOperation {
  name: string;
  pattern: RegExp;
  replacer: (match: string, ...args: string[]) => string;
}

export interface WordBatchFileOutput {
  filename: string;
  kind: 'docx' | 'xlsx';
  buffer: Buffer;
}

export interface WordBatchSingleFileResult {
  originalName: string;
  ok: boolean;
  convertedFromDoc: boolean;
  counts: Record<string, number>;
  outputs: WordBatchFileOutput[];
  logs: string[];
  error?: string;
}

// Numeric placeholder in template: xx+, XX+, *+, ＊+, or "数字"
const PLACEHOLDER_RE = /[xX]{2,}|＊+|\*+|数字/;
const TRAILING_PUNCT = '[，。、；：,.;:]?';

function escapeRegex(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export function buildNumberPattern(template: string): RegExp {
  const m = PLACEHOLDER_RE.exec(template);
  if (!m) {
    throw new Error('模板里需要包含数字占位符（例如 * 或 xxx 或“数字”）来代表要匹配的数字。');
  }
  const prefix = template.slice(0, m.index);
  const suffix = template.slice(m.index + m[0].length);
  if (!prefix && !suffix) {
    throw new Error('模板不能只有数字占位符，请保留前后文字以便精确定位。');
  }
  return new RegExp(
    `(${escapeRegex(prefix)})(\\d+(?:\\.\\d+)?)(${escapeRegex(suffix)})`,
    'g'
  );
}

export function buildDeletePattern(template: string): RegExp {
  const m = PLACEHOLDER_RE.exec(template);
  if (!m) {
    throw new Error('模板里需要包含数字占位符（例如 * 或 xxx 或“数字”）来代表要匹配的数字。');
  }
  const prefix = template.slice(0, m.index);
  const suffix = template.slice(m.index + m[0].length);
  if (!prefix && !suffix) {
    throw new Error('模板不能只有数字占位符，请保留前后文字以便精确定位。');
  }
  return new RegExp(
    `${escapeRegex(prefix)}\\d+(?:\\.\\d+)?${escapeRegex(suffix)}${TRAILING_PUNCT}`,
    'g'
  );
}

export function buildWordBatchOperations(cfg: WordBatchConfig): WordBatchOperation[] {
  const ops: WordBatchOperation[] = [];

  if (cfg.modifyEnabled) {
    const mulStr = String(cfg.multiplier ?? '').trim();
    let mulDec: Dec;
    try {
      if (!mulStr) throw new Error('empty');
      mulDec = Dec.of(mulStr);
    } catch {
      throw new Error('倍数必须是有效数字。');
    }
    const tpl = String(cfg.modifyTemplate ?? '').trim();
    const pattern = buildNumberPattern(tpl);
    ops.push({
      name: '修改',
      pattern,
      replacer: (_match, g1, g2, g3) => {
        const newVal = Dec.of(g2).mul(mulDec).stripTrailingZeros().fmt();
        return `${g1}${newVal}${g3}`;
      },
    });
  }

  if (cfg.deleteEnabled) {
    const tpl = String(cfg.deleteTemplate ?? '').trim();
    const pattern = buildDeletePattern(tpl);
    ops.push({
      name: '删除',
      pattern,
      replacer: () => '',
    });
  }

  if (cfg.replaceEnabled) {
    const validPairs = (cfg.replacePairs || []).filter((p) => p && String(p.find || '').length > 0);
    if (validPairs.length === 0) {
      throw new Error('已启用「查找替换」，但尚未添加任何替换规则。');
    }
    for (const pair of validPairs) {
      const find = String(pair.find);
      const rep = String(pair.replace ?? '');
      if (cfg.replaceUseRegex) {
        let pattern: RegExp;
        try {
          pattern = new RegExp(find, 'g');
        } catch (err) {
          throw new Error(`正则表达式「${find}」不合法：${(err as Error).message}`);
        }
        // Support both Python/Word \1..\9 and JS $1..$9 backreferences
        const jsRep = rep.replace(/\\(\d+)/g, '$$$1');
        ops.push({
          name: `替换「${find}→${rep}」`,
          pattern,
          replacer: (match, ...args) => {
            // Re-run single replacement with JS backreference semantics
            const singleRe = new RegExp(find);
            void args;
            return match.replace(singleRe, jsRep);
          },
        });
      } else {
        const pattern = new RegExp(escapeRegex(find), 'g');
        ops.push({
          name: `替换「${find}→${rep}」`,
          pattern,
          replacer: () => rep,
        });
      }
    }
  }

  if (ops.length === 0 && !cfg.extractEnabled) {
    throw new Error('请至少启用「批量修改」「批量删除」「查找替换」「提取表格」中的一项。');
  }

  return ops;
}

// ==================== XML Helpers (lossless OOXML processing) ====================

function decodeXmlEntities(s: string): string {
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_m, n) => String.fromCodePoint(parseInt(n, 10)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_m, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&amp;/g, '&');
}

function encodeXmlEntities(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function formatWtTag(openTag: string, text: string): string {
  const escaped = encodeXmlEntities(text);
  const needsPreserve = /^\s|\s$|\s{2,}/.test(text);
  let tag = openTag;
  if (needsPreserve && !/xml:space\s*=/.test(tag)) {
    tag = tag.replace(/>$/, ' xml:space="preserve">');
  }
  return `${tag}${escaped}</w:t>`;
}

function applyPatternCount(
  text: string,
  op: WordBatchOperation
): { text: string; count: number } {
  let count = 0;
  op.pattern.lastIndex = 0;
  const next = text.replace(op.pattern, (...args: unknown[]) => {
    count++;
    return op.replacer(
      String(args[0]),
      ...args.slice(1).map((x) => (typeof x === 'string' ? x : ''))
    );
  });
  return { text: next, count };
}

interface TagSpan {
  start: number;
  end: number;
}

/**
 * Find all non-overlapping or nested spans of a specific XML element (e.g. "w:p", "w:tbl", "w:tr", "w:tc").
 * When `topLevelOnly` is true, returns only outermost spans (relative to the input string).
 */
function findElementSpans(xml: string, tagName: string, topLevelOnly = false): TagSpan[] {
  const openRe = new RegExp(`<${tagName}(?:\\s[^>]*)?>|<${tagName}\\s*/>|</${tagName}>`, 'g');
  const stack: number[] = [];
  const spans: TagSpan[] = [];
  let m: RegExpExecArray | null;

  while ((m = openRe.exec(xml)) !== null) {
    const tok = m[0];
    if (tok.startsWith('</')) {
      if (stack.length > 0) {
        const start = stack.pop()!;
        if (!topLevelOnly || stack.length === 0) {
          spans.push({ start, end: m.index + tok.length });
        }
      }
    } else if (tok.endsWith('/>')) {
      if (!topLevelOnly || stack.length === 0) {
        spans.push({ start: m.index, end: m.index + tok.length });
      }
    } else {
      stack.push(m.index);
    }
  }
  return spans;
}

/**
 * Process a single `<w:p>...</w:p>` fragment with an operation.
 * Preserves run-level formatting when a match falls inside a single `<w:r>`,
 * and falls back to paragraph-level replacement (writing back to the first `<w:t>`)
 * when Word split a phrase across multiple `<w:r>` runs.
 * Leaves any nested `<w:txbxContent>` sub-paragraphs untouched (they are processed separately).
 */
function processSingleParagraphXml(
  pXml: string,
  op: WordBatchOperation
): { xml: string; count: number } {
  // Mask out any nested <w:txbxContent>...</w:txbxContent> blocks so we don't touch their inner <w:p> here
  const maskedBlocks: string[] = [];
  let work = pXml.replace(/<w:txbxContent\b[\s\S]*?<\/w:txbxContent>/g, (blk) => {
    const idx = maskedBlocks.length;
    maskedBlocks.push(blk);
    return `__TXBX_MASK_${idx}__`;
  });

  const restoreMasks = (s: string) =>
    s.replace(/__TXBX_MASK_(\d+)__/g, (_m, i) => maskedBlocks[Number(i)]);

  // Collect all <w:r> spans inside this paragraph
  const wtRe = /(<w:t(?:\s[^>]*)?>)([\s\S]*?)(<\/w:t>)/g;
  const runSpans = findElementSpans(work, 'w:r', true);
  if (runSpans.length === 0) {
    return { xml: pXml, count: 0 };
  }

  // Pass 1: run-level replacement
  let runCount = 0;
  let rebuilt = '';
  let cursor = 0;

  for (const rs of runSpans) {
    rebuilt += work.slice(cursor, rs.start);
    const runXml = work.slice(rs.start, rs.end);
    // Concatenate all <w:t> inside this single run
    const wtMatches: Array<{ start: number; end: number; openTag: string; text: string }> = [];
    wtRe.lastIndex = 0;
    let wm: RegExpExecArray | null;
    while ((wm = wtRe.exec(runXml)) !== null) {
      wtMatches.push({
        start: wm.index,
        end: wm.index + wm[0].length,
        openTag: wm[1],
        text: decodeXmlEntities(wm[2]),
      });
    }
    if (wtMatches.length > 0) {
      const runText = wtMatches.map((x) => x.text).join('');
      const res = applyPatternCount(runText, op);
      if (res.count > 0) {
        runCount += res.count;
        let newRun = '';
        let rCur = 0;
        wtMatches.forEach((wt, idx) => {
          newRun += runXml.slice(rCur, wt.start);
          newRun += formatWtTag(wt.openTag, idx === 0 ? res.text : '');
          rCur = wt.end;
        });
        newRun += runXml.slice(rCur);
        rebuilt += newRun;
      } else {
        rebuilt += runXml;
      }
    } else {
      rebuilt += runXml;
    }
    cursor = rs.end;
  }
  rebuilt += work.slice(cursor);

  if (runCount > 0) {
    return { xml: restoreMasks(rebuilt), count: runCount };
  }

  // Pass 2: paragraph-level fallback (phrase split across multiple runs)
  const allWt: Array<{ start: number; end: number; openTag: string; text: string }> = [];
  wtRe.lastIndex = 0;
  let wm: RegExpExecArray | null;
  while ((wm = wtRe.exec(work)) !== null) {
    allWt.push({
      start: wm.index,
      end: wm.index + wm[0].length,
      openTag: wm[1],
      text: decodeXmlEntities(wm[2]),
    });
  }
  if (allWt.length === 0) {
    return { xml: pXml, count: 0 };
  }

  const fullText = allWt.map((x) => x.text).join('');
  const res = applyPatternCount(fullText, op);
  if (res.count === 0) {
    return { xml: pXml, count: 0 };
  }

  let out = '';
  let pCur = 0;
  allWt.forEach((wt, idx) => {
    out += work.slice(pCur, wt.start);
    out += formatWtTag(wt.openTag, idx === 0 ? res.text : '');
    pCur = wt.end;
  });
  out += work.slice(pCur);
  return { xml: restoreMasks(out), count: res.count };
}

/**
 * Apply a single operation across every `<w:p>` element inside `<w:body>` (including tables & textboxes),
 * visiting each `<w:p>` exactly once.
 */
function applyOperationToDocumentXml(
  docXml: string,
  op: WordBatchOperation
): { xml: string; count: number } {
  const bodyMatch = /<w:body\b[^>]*>[\s\S]*<\/w:body>/.exec(docXml);
  if (!bodyMatch) return { xml: docXml, count: 0 };

  const bodyStart = bodyMatch.index;
  const bodyEnd = bodyStart + bodyMatch[0].length;
  let bodyXml = bodyMatch[0];
  let totalCount = 0;

  // Process from innermost to outermost or top-level paragraphs + textbox paragraphs
  // First process any paragraphs inside <w:txbxContent> blocks
  bodyXml = bodyXml.replace(/<w:txbxContent\b[\s\S]*?<\/w:txbxContent>/g, (txbx) => {
    const innerSpans = findElementSpans(txbx, 'w:p', true);
    if (innerSpans.length === 0) return txbx;
    let out = '';
    let cur = 0;
    for (const sp of innerSpans) {
      out += txbx.slice(cur, sp.start);
      const r = processSingleParagraphXml(txbx.slice(sp.start, sp.end), op);
      totalCount += r.count;
      out += r.xml;
      cur = sp.end;
    }
    out += txbx.slice(cur);
    return out;
  });

  // Then process all regular (non-textbox) <w:p> spans (masked inside processSingleParagraphXml)
  const outerSpans = findElementSpans(bodyXml, 'w:p', true);
  let out = '';
  let cur = 0;
  for (const sp of outerSpans) {
    out += bodyXml.slice(cur, sp.start);
    const r = processSingleParagraphXml(bodyXml.slice(sp.start, sp.end), op);
    totalCount += r.count;
    out += r.xml;
    cur = sp.end;
  }
  out += bodyXml.slice(cur);

  return {
    xml: docXml.slice(0, bodyStart) + out + docXml.slice(bodyEnd),
    count: totalCount,
  };
}

// ==================== Evaluation Table Extraction (Word -> Excel) ====================

const EVAL_COLS = [
  '序号',
  '应答人',
  '应答报价（万元）',
  '技术加权分',
  '价格加权分',
  '商务加权分',
  '总分',
  '排序',
  '推荐的成交候选人',
];
export const EVAL_EXCEL_HEADER = ['分标名称', '分包', ...EVAL_COLS];
const INT_COLS = new Set(['序号', '排序']);
const FLOAT_COLS = new Set(['应答报价（万元）', '技术加权分', '价格加权分', '商务加权分', '总分']);
const BID_RE = /第\s*\d+\s*分标[:：]\s*(.+)/;
const PKG_RE = /^(包\d+)/;

function extractPlainTextFromXml(fragment: string): string {
  const parts: string[] = [];
  const wtRe = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
  let m: RegExpExecArray | null;
  while ((m = wtRe.exec(fragment)) !== null) {
    parts.push(decodeXmlEntities(m[1]));
  }
  return parts.join('');
}

function collapseConsecutive(vals: string[]): string[] {
  const out: string[] = [];
  for (const v of vals) {
    if (out.length === 0 || out[out.length - 1] !== v) {
      out.push(v);
    }
  }
  return out;
}

function parseTableRows(tblXml: string): string[][] {
  // Strip outer <w:tbl ...> and </w:tbl>
  const firstClose = tblXml.indexOf('>');
  const lastOpen = tblXml.lastIndexOf('</w:tbl>');
  if (firstClose < 0 || lastOpen <= firstClose) return [];
  const inner = tblXml.slice(firstClose + 1, lastOpen);

  const trSpans = findElementSpans(inner, 'w:tr', true);
  const rows: string[][] = [];
  let prevRowVals: string[] = [];

  for (const trSpan of trSpans) {
    const trXml = inner.slice(trSpan.start, trSpan.end);
    const trFirstClose = trXml.indexOf('>');
    const trLastOpen = trXml.lastIndexOf('</w:tr>');
    if (trFirstClose < 0 || trLastOpen <= trFirstClose) continue;
    const trInner = trXml.slice(trFirstClose + 1, trLastOpen);

    const tcSpans = findElementSpans(trInner, 'w:tc', true);
    const rowVals: string[] = [];

    for (const tcSpan of tcSpans) {
      const tcXml = trInner.slice(tcSpan.start, tcSpan.end);
      // Extract cell text from all direct paragraphs (joined without newlines, matching python-docx strip().replace('\n',''))
      const pSpans = findElementSpans(tcXml, 'w:p', true);
      const rawCellText = pSpans
        .map((ps) => extractPlainTextFromXml(tcXml.slice(ps.start, ps.end)))
        .join('')
        .replace(/[\r\n]+/g, '')
        .trim();

      // Check vertical merge (<w:vMerge/> or <w:vMerge w:val="continue"/> inherits previous row's cell value)
      const vMergeMatch = /<w:vMerge\b([^>]*)\/?>/.exec(tcXml);
      const isVMergeContinue = Boolean(vMergeMatch && !/w:val="restart"/.test(vMergeMatch[1]));

      // Respect <w:gridSpan w:val="N"/> so merged cells behave identically to python-docx row.cells
      const gridSpanMatch = /<w:gridSpan\b[^>]*w:val="(\d+)"/.exec(tcXml);
      const spanCount = gridSpanMatch ? Math.max(1, parseInt(gridSpanMatch[1], 10)) : 1;
      for (let k = 0; k < spanCount; k++) {
        const colIdx = rowVals.length;
        const val = isVMergeContinue && !rawCellText ? (prevRowVals[colIdx] ?? '') : rawCellText;
        rowVals.push(val);
      }
    }
    rows.push(rowVals);
    prevRowVals = rowVals;
  }

  return rows;
}

/**
 * Walk direct children of `<w:body>` in document order and extract all candidate scoring tables
 * (`序号` / `应答人`), matching `extract_eval_tables` in `fix_limit_price_gui.py` 1:1.
 */
export function extractEvalTablesFromDocumentXml(docXml: string): string[][] {
  const bodyOpen = /<w:body\b[^>]*>/.exec(docXml);
  const bodyCloseIdx = docXml.lastIndexOf('</w:body>');
  if (!bodyOpen || bodyCloseIdx < 0) return [];

  const bodyInner = docXml.slice(bodyOpen.index + bodyOpen[0].length, bodyCloseIdx);

  // Find direct top-level <w:p> and <w:tbl> children in document order
  const pSpans = findElementSpans(bodyInner, 'w:p', true).map((s) => ({ ...s, type: 'p' as const }));
  const tblSpans = findElementSpans(bodyInner, 'w:tbl', true).map((s) => ({ ...s, type: 'tbl' as const }));

  // Filter out <w:p> that are inside a <w:tbl>
  const directP = pSpans.filter(
    (p) => !tblSpans.some((t) => p.start >= t.start && p.end <= t.end)
  );
  const children = [...directP, ...tblSpans].sort((a, b) => a.start - b.start);

  let curBid = '';
  const rowsOut: string[][] = [];

  for (const child of children) {
    const frag = bodyInner.slice(child.start, child.end);
    if (child.type === 'p') {
      const text = extractPlainTextFromXml(frag).trim();
      const m = BID_RE.exec(text);
      if (m) {
        curBid = m[1].split(/[（(]/)[0].trim();
      }
    } else {
      const tblRows = parseTableRows(frag);
      const isTargetTable = tblRows.some((r) => {
        const col = collapseConsecutive(r);
        return col.length >= 2 && col[0] === '序号' && col[1] === '应答人';
      });
      if (!isTargetTable) continue;

      let curPkg = '';
      for (const vals of tblRows) {
        const col = collapseConsecutive(vals);
        if (col.length === 1) {
          const txt = col[0];
          const mp = PKG_RE.exec(txt);
          if (mp) {
            curPkg = mp[1];
          } else if ((txt.includes('流标') || txt.includes('不足')) && curPkg) {
            rowsOut.push([curBid, curPkg, '', txt, '', '', '', '', '', '', '']);
          }
          continue;
        }
        if (vals.length > 0 && vals[0] === '序号') continue;
        if (vals.length > 0 && /^\d+$/.test(vals[0])) {
          const padded = vals.slice(0, 9);
          while (padded.length < 9) padded.push('');
          rowsOut.push([curBid, curPkg, ...padded]);
        }
      }
    }
  }

  return rowsOut;
}

function coerceEvalCell(colName: string, value: string): string | number {
  if (!value) return '';
  if (INT_COLS.has(colName) && /^-?\d+$/.test(value)) {
    return parseInt(value, 10);
  }
  if (FLOAT_COLS.has(colName) && /^-?\d+(?:\.\d+)?$/.test(value)) {
    const n = Number(value);
    if (Number.isFinite(n)) return n;
  }
  return value;
}

export async function buildEvalTablesXlsxBuffer(rows: string[][]): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = '国网招标代理 SOP 工作台';
  const ws = wb.addWorksheet('评标结果', {
    views: [{ state: 'frozen', ySplit: 1 }],
  });

  const headerRow = ws.addRow(EVAL_EXCEL_HEADER);
  headerRow.height = 24;
  headerRow.eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3B5B7E' } };
    cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFD1D5DB' } },
      left: { style: 'thin', color: { argb: 'FFD1D5DB' } },
      bottom: { style: 'thin', color: { argb: 'FFD1D5DB' } },
      right: { style: 'thin', color: { argb: 'FFD1D5DB' } },
    };
  });

  for (const r of rows) {
    const coerced = EVAL_EXCEL_HEADER.map((col, idx) => coerceEvalCell(col, r[idx] || ''));
    const row = ws.addRow(coerced);
    row.eachCell((cell) => {
      cell.font = { name: 'Calibri', size: 11 };
      cell.alignment = { vertical: 'middle', wrapText: true };
    });
  }

  // Estimate column widths (CJK chars count as 2), clamped to [8, 60]
  EVAL_EXCEL_HEADER.forEach((colName, idx) => {
    let maxW = colName.replace(/[^\x00-\xff]/g, 'xx').length;
    for (const r of rows) {
      const s = String(r[idx] || '');
      const w = s.replace(/[^\x00-\xff]/g, 'xx').length;
      if (w > maxW) maxW = w;
    }
    ws.getColumn(idx + 1).width = Math.min(Math.max(maxW + 3, 8), 60);
  });

  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
}

// ==================== Legacy .doc Detection & Conversion ====================

const OLE2_MAGIC = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);

export function isLegacyDocBuffer(buf: Buffer): boolean {
  return buf.length >= 8 && buf.subarray(0, 8).equals(OLE2_MAGIC);
}

function findSofficeBinary(): string | null {
  const candidates =
    process.platform === 'win32'
      ? [
          'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
          'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
          'soffice.exe',
        ]
      : [
          '/Applications/LibreOffice.app/Contents/MacOS/soffice',
          '/usr/bin/soffice',
          '/usr/local/bin/soffice',
          'soffice',
          'libreoffice',
        ];

  for (const c of candidates) {
    if (path.isAbsolute(c)) {
      if (fs.existsSync(c)) return c;
    } else {
      try {
        execFileSync(process.platform === 'win32' ? 'where' : 'which', [c], {
          stdio: 'ignore',
        });
        return c;
      } catch {
        // continue
      }
    }
  }
  return null;
}

function convertLegacyDocBufferToDocx(buf: Buffer, originalName: string): Buffer {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'word_batch_doc_'));
  const srcPath = path.join(tmpDir, 'src.doc');
  const outPath = path.join(tmpDir, 'src.docx');
  fs.writeFileSync(srcPath, buf);

  try {
    const soffice = findSofficeBinary();
    if (soffice) {
      execFileSync(soffice, ['--headless', '--convert-to', 'docx', '--outdir', tmpDir, srcPath], {
        stdio: 'pipe',
        timeout: 120_000,
      });
      if (fs.existsSync(outPath)) {
        return fs.readFileSync(outPath);
      }
      throw new Error(`LibreOffice 转换失败：${originalName}`);
    }

    if (process.platform === 'win32') {
      const psScript = `
$ErrorActionPreference = 'Stop'
$src = '${srcPath.replace(/'/g, "''")}'
$out = '${outPath.replace(/'/g, "''")}'
$app = $null
foreach ($progId in @('Word.Application','KWPS.Application','kwps.Application','wps.Application','WPS.Application')) {
  try { $app = New-Object -ComObject $progId; break } catch {}
}
if ($null -eq $app) { throw '未检测到已安装的 Microsoft Word 或 WPS' }
try { $app.Visible = $false } catch {}
try { $app.DisplayAlerts = 0 } catch {}
try {
  $doc = $app.Documents.Open($src, $false, $true)
  try { $doc.SaveAs2($out, 16) } catch { $doc.SaveAs($out, 12) }
  $doc.Close($false)
} finally {
  try { $app.Quit() } catch {}
}
`;
      execFileSync('powershell', ['-NoProfile', '-NonInteractive', '-Command', psScript], {
        stdio: 'pipe',
        timeout: 120_000,
      });
      if (fs.existsSync(outPath)) {
        return fs.readFileSync(outPath);
      }
      throw new Error(`Word/WPS 转换失败：${originalName}`);
    }

    throw new Error(
      `「${originalName}」为旧版二进制 .doc 格式，当前服务器未安装 LibreOffice（soffice），无法自动转换。请先在 WPS / Word 中将其另存为 .docx 后再上传。`
    );
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
}

// ==================== Single File Pipeline ====================

export async function processSingleWordFile(
  originalName: string,
  rawBuffer: Buffer,
  operations: WordBatchOperation[],
  extractExcel: boolean
): Promise<WordBatchSingleFileResult> {
  const logs: string[] = [`[${originalName}]`];
  const counts: Record<string, number> = {};
  const outputs: WordBatchFileOutput[] = [];
  let convertedFromDoc = false;

  try {
    let workBuffer = rawBuffer;
    if (isLegacyDocBuffer(rawBuffer)) {
      logs.push('  · 识别为旧版二进制 .doc 格式，正在自动转换为 .docx…');
      workBuffer = convertLegacyDocBufferToDocx(rawBuffer, originalName);
      convertedFromDoc = true;
    }

    let zip: PizZip;
    try {
      zip = new PizZip(workBuffer);
    } catch {
      throw new Error(`无法作为 .docx 打开「${originalName}」，请确认文件未损坏且未加密。`);
    }

    const docEntry = zip.file('word/document.xml');
    if (!docEntry) {
      throw new Error(`「${originalName}」缺少 word/document.xml，不是有效的 Word 文档。`);
    }

    let docXml = docEntry.asText();
    const baseStem = originalName.replace(/\.docx?$/i, '');
    const tag = convertedFromDoc ? '（已自动转换格式）' : '';

    // 1. Text operations: modify / delete / find-replace
    if (operations.length > 0) {
      for (const op of operations) {
        const res = applyOperationToDocumentXml(docXml, op);
        docXml = res.xml;
        counts[op.name] = res.count;
      }
      zip.file('word/document.xml', docXml);
      const outDocxBuf = zip.generate({
        type: 'nodebuffer',
        compression: 'DEFLATE',
      }) as Buffer;
      const outDocxName = `${baseStem}_modified.docx`;
      outputs.push({ filename: outDocxName, kind: 'docx', buffer: outDocxBuf });
      const detail = Object.entries(counts)
        .map(([k, v]) => `${k} ${v} 处`)
        .join('，');
      logs.push(`  · ${detail} -> ${outDocxName} ${tag}`.trimEnd());
    }

    // 2. Extract evaluation tables -> *_表格.xlsx (run on updated docXml so modified numbers carry into Excel!)
    if (extractExcel) {
      const tableRows = extractEvalTablesFromDocumentXml(docXml);
      const xlsxBuf = await buildEvalTablesXlsxBuffer(tableRows);
      const outXlsxName = `${baseStem}_表格.xlsx`;
      outputs.push({ filename: outXlsxName, kind: 'xlsx', buffer: xlsxBuf });
      counts['提取表格'] = tableRows.length;
      logs.push(
        `  · 提取评标表格 ${tableRows.length} 行 -> ${outXlsxName} ${operations.length === 0 ? tag : ''}`.trimEnd()
      );
    }

    return {
      originalName,
      ok: true,
      convertedFromDoc,
      counts,
      outputs,
      logs,
    };
  } catch (err) {
    const msg = (err as Error).message;
    logs.push(`  · 跳过（处理失败）：${msg}`);
    return {
      originalName,
      ok: false,
      convertedFromDoc,
      counts,
      outputs: [],
      logs,
      error: msg,
    };
  }
}
