import fs from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';
import * as XLSXNamespace from 'xlsx';

const XLSX = (XLSXNamespace as unknown as { default?: typeof XLSXNamespace }).default ?? XLSXNamespace;

export interface ParsedSheetTable {
  name: string;
  headers: string[];
  rows: Record<string, string>[];
  totalRows: number;
}

export interface ParsedWorkbookTables {
  sheets: ParsedSheetTable[];
  parsedAt: string;
}

export interface ParsedArchiveRow {
  sheetName: string;
  seq: string;
  docName: string;
  isRequired: boolean;
  requiredSource: 'color' | 'keyword' | 'manual';
  highlightColor?: string;
  mediaType: string;
  paperReq: string;
  electronicReq: string;
  policyBasis: string;
  templateSource: string;
  status: 'present' | 'missing' | 'na';
  notes: string;
}

export interface ParsedArchiveWorkbook {
  sheetNames: string[];
  items: ParsedArchiveRow[];
  colorDetectedCount: number;
  format: 'xlsx' | 'xls' | 'other';
}

/**
 * Parse general spreadsheets (Bidding / Evaluation stage tables) into structured JSON
 * for preview and downstream Award stage field extraction.
 */
export function parseSpreadsheetTables(filePath: string, maxRows = 200): ParsedWorkbookTables | null {
  const ext = path.extname(filePath).toLowerCase();
  if (!['.xlsx', '.xls', '.csv'].includes(ext)) return null;

  const buf = fs.readFileSync(filePath);
  const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });
  const sheets: ParsedSheetTable[] = [];

  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    const aoa = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: '' });
    if (!aoa.length) continue;

    // Find the first non-empty row as header
    let headerIdx = 0;
    for (let i = 0; i < Math.min(aoa.length, 5); i++) {
      const nonEmpty = (aoa[i] || []).filter((c) => String(c ?? '').trim() !== '').length;
      if (nonEmpty >= 2) {
        headerIdx = i;
        break;
      }
    }

    const rawHeaders = (aoa[headerIdx] || []).map((h, idx) => {
      const text = String(h ?? '').trim();
      return text || `列${idx + 1}`;
    });

    const dataRows = aoa.slice(headerIdx + 1).filter((r) =>
      (r || []).some((cell) => String(cell ?? '').trim() !== '')
    );

    const rows = dataRows.slice(0, maxRows).map((r) => {
      const obj: Record<string, string> = {};
      rawHeaders.forEach((h, i) => {
        const val = r[i];
        if (val instanceof Date) {
          obj[h] = val.toISOString().slice(0, 10);
        } else {
          obj[h] = String(val ?? '').trim();
        }
      });
      return obj;
    });

    sheets.push({
      name,
      headers: rawHeaders,
      rows,
      totalRows: dataRows.length,
    });
  }

  return {
    sheets,
    parsedAt: new Date().toISOString(),
  };
}

/**
 * Check if an ExcelJS cell has a meaningful non-white background fill or non-black highlight font color.
 */
function hasExcelJsHighlight(cell: ExcelJS.Cell): { highlighted: boolean; color?: string } {
  const fill = cell.fill as ExcelJS.FillPattern | undefined;
  if (fill && fill.type === 'pattern' && fill.pattern && fill.pattern !== 'none') {
    const fg = fill.fgColor as { argb?: string; theme?: number; indexed?: number } | undefined;
    if (fg) {
      const argb = fg.argb?.toUpperCase();
      if (argb && !['FFFFFFFF', '00FFFFFF', '00000000', 'FF000000'].includes(argb)) {
        return { highlighted: true, color: `#${argb.slice(-6)}` };
      }
      if (typeof fg.theme === 'number' && fg.theme !== 0 && fg.theme !== 1) {
        return { highlighted: true, color: `theme:${fg.theme}` };
      }
      if (typeof fg.indexed === 'number' && fg.indexed !== 64 && fg.indexed !== 65 && fg.indexed !== 9) {
        return { highlighted: true, color: `indexed:${fg.indexed}` };
      }
    }
  }
  const fontColor = cell.font?.color as { argb?: string; theme?: number } | undefined;
  if (fontColor?.argb) {
    const argb = fontColor.argb.toUpperCase();
    if (!['FF000000', '00000000', 'FFFFFFFF'].includes(argb)) {
      return { highlighted: true, color: `font:#${argb.slice(-6)}` };
    }
  }
  return { highlighted: false };
}

/**
 * Pure-Node BIFF8 (.xls) parser to extract cell background color fill flags for Column B
 * across each sheet in legacy .xls files!
 */
function extractBiff8ColumnBHighlights(filePath: string): Map<string, Set<number>> {
  const result = new Map<string, Set<number>>();
  try {
    const buf = fs.readFileSync(filePath);
    const cfb = XLSX.CFB.read(buf, { type: 'buffer' });
    const entry =
      XLSX.CFB.find(cfb, '/Workbook') ||
      XLSX.CFB.find(cfb, '/Book') ||
      XLSX.CFB.find(cfb, 'Workbook');
    if (!entry || !entry.content) return result;

    const data = Buffer.from(entry.content as Uint8Array);
    let offset = 0;

    // Track XF records: xfIndex -> hasNonWhiteFill
    const xfHasColor: boolean[] = [];
    const sheetOffsets: Array<{ name: string; pos: number }> = [];

    // First pass: collect XF records & BOUNDSHEET records from Workbook Globals
    while (offset + 4 <= data.length) {
      const recId = data.readUInt16LE(offset);
      const recLen = data.readUInt16LE(offset + 2);
      const recStart = offset + 4;
      if (recStart + recLen > data.length) break;

      if (recId === 0x00e0 && recLen >= 20) {
        // BIFF8 XF record
        // Bytes 16..17 (offset 16): pattern & flags; bits 26..31 of DWORD at offset 14 -> actually offset 14..15 = border2, offset 16..17 = fill pattern bits 10..15, offset 18..19 = icvFore (bits 0..6) & icvBack (bits 7..13)
        const fillPatternWord = data.readUInt16LE(recStart + 16);
        const fillPattern = (fillPatternWord >> 10) & 0x3f;
        const colorWord = data.readUInt16LE(recStart + 18);
        const icvFore = colorWord & 0x7f;
        // Standard BIFF8: fillPattern > 0 and icvFore not automatic (64/65) and not white (9)
        const isColored = fillPattern > 0 && icvFore !== 64 && icvFore !== 65 && icvFore !== 9;
        xfHasColor.push(isColored);
      } else if (recId === 0x0085 && recLen >= 8) {
        // BOUNDSHEET
        const sheetPos = data.readUInt32LE(recStart);
        const nameLen = data.readUInt8(recStart + 6);
        const flags = data.readUInt8(recStart + 7);
        const isUtf16 = (flags & 0x01) !== 0;
        const nameBuf = data.subarray(recStart + 8, recStart + 8 + (isUtf16 ? nameLen * 2 : nameLen));
        const name = isUtf16 ? nameBuf.toString('utf16le') : nameBuf.toString('latin1');
        sheetOffsets.push({ name, pos: sheetPos });
      } else if (recId === 0x000a) {
        // EOF of globals
        break;
      }

      offset = recStart + recLen;
    }

    // Second pass: for each sheet, scan cell records in Column B (col === 1)
    for (const sh of sheetOffsets) {
      const coloredRows = new Set<number>();
      let pos = sh.pos;
      while (pos + 4 <= data.length) {
        const recId = data.readUInt16LE(pos);
        const recLen = data.readUInt16LE(pos + 2);
        const recStart = pos + 4;
        if (recStart + recLen > data.length) break;

        if (recId === 0x000a) break; // Sheet EOF

        // Cell records with [row: u16, col: u16, ixfe: u16]:
        // LABELSST (0x00FD), LABEL (0x0204), BLANK (0x0201), RK (0x027E), NUMBER (0x0203), FORMULA (0x0006)
        if ([0x00fd, 0x0204, 0x0201, 0x027e, 0x0203, 0x0006].includes(recId) && recLen >= 6) {
          const row = data.readUInt16LE(recStart);
          const col = data.readUInt16LE(recStart + 2);
          const ixfe = data.readUInt16LE(recStart + 4);
          if (col === 1 && xfHasColor[ixfe]) {
            coloredRows.add(row);
          }
        }
        pos = recStart + recLen;
      }
      result.set(sh.name, coloredRows);
    }
  } catch {
    // Graceful fallback if binary structure differs
  }
  return result;
}

/**
 * Normalize a single archive row from column mapping.
 */
function buildArchiveRow(
  sheetName: string,
  rowValues: string[],
  colMap: {
    seq: number;
    docName: number;
    mediaType: number;
    paperReq: number;
    electronicReq: number;
    policyBasis: number;
    templateSource: number;
    notes: number;
  },
  isColorHighlighted: boolean,
  highlightColor?: string
): ParsedArchiveRow | null {
  const docName = (rowValues[colMap.docName] || '').trim();
  if (!docName || docName === '归档文件' || docName === '保管文件' || docName === '资料名称') {
    return null;
  }

  let seq = (rowValues[colMap.seq] || '').trim();
  if (/^\d+\.0$/.test(seq)) seq = seq.replace(/\.0$/, '');

  const notes = colMap.notes >= 0 ? (rowValues[colMap.notes] || '').trim() : '';
  const isNA = notes === '无' || notes === '不涉及';
  const isReturned = /已归还|已齐|已有|完成/.test(notes);

  let isRequired = isColorHighlighted;
  let requiredSource: 'color' | 'keyword' | 'manual' = isColorHighlighted ? 'color' : 'manual';

  return {
    sheetName,
    seq,
    docName,
    isRequired,
    requiredSource,
    highlightColor,
    mediaType: (rowValues[colMap.mediaType] || '纸质+电子').trim() || '纸质+电子',
    paperReq: (rowValues[colMap.paperReq] || '').trim(),
    electronicReq: (rowValues[colMap.electronicReq] || '').trim(),
    policyBasis: (rowValues[colMap.policyBasis] || '').trim(),
    templateSource: (rowValues[colMap.templateSource] || '').trim(),
    status: isNA ? 'na' : isReturned ? 'present' : 'missing',
    notes,
  };
}

function detectColumnMap(headers: string[]) {
  const findIdx = (patterns: RegExp[], fallback: number) => {
    const idx = headers.findIndex((h) => patterns.some((p) => p.test(h)));
    return idx >= 0 ? idx : fallback;
  };

  const seq = findIdx([/序号|编号|no/i], 0);
  const docName = findIdx([/归档文件|保管文件|资料名称|文件名称|名称/i], 1);
  const mediaType = findIdx([/保存介质|介质/i], -1);
  const paperReq = findIdx([/纸质要求|纸质/i], -1);
  const electronicReq = findIdx([/电子要求|电子/i], -1);
  const policyBasis = findIdx([/制度依据|依据/i], -1);
  const templateSource = findIdx([/模板来源|模板/i], -1);
  let notes = findIdx([/备注|说明|进度|状态/i], -1);
  // In the user's xls Sheet 1, the last column (after 模板来源) has no header text but holds notes like "叶勤已归还"
  if (notes < 0 && templateSource >= 0 && headers.length > templateSource + 1) {
    notes = templateSource + 1;
  }

  return { seq, docName, mediaType, paperReq, electronicReq, policyBasis, templateSource, notes };
}

/**
 * Parse Archive Catalog Excel (.xlsx or .xls) with multi-sheet support and B-column color detection!
 */
export async function parseArchiveCatalogExcel(filePath: string): Promise<ParsedArchiveWorkbook> {
  const ext = path.extname(filePath).toLowerCase();
  const items: ParsedArchiveRow[] = [];
  const sheetNames: string[] = [];
  let colorDetectedCount = 0;

  if (ext === '.xlsx') {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);

    wb.eachSheet((ws) => {
      const rowsData: Array<{ values: string[]; colBHighlighted: boolean; color?: string }> = [];
      ws.eachRow({ includeEmpty: false }, (row) => {
        const vals: string[] = [];
        const maxCol = Math.max(row.cellCount, 9);
        for (let c = 1; c <= maxCol; c++) {
          const cell = row.getCell(c);
          vals.push(String(cell.text ?? cell.value ?? '').trim());
        }
        const cellB = row.getCell(2);
        const hl = hasExcelJsHighlight(cellB);
        rowsData.push({ values: vals, colBHighlighted: hl.highlighted, color: hl.color });
      });

      if (rowsData.length < 2) return;
      const headerRowIdx = rowsData.findIndex((r) =>
        r.values.some((v) => /序号|归档文件|保管文件|资料名称/.test(v))
      );
      const hIdx = headerRowIdx >= 0 ? headerRowIdx : 0;
      const headers = rowsData[hIdx].values;
      const colMap = detectColumnMap(headers);

      // Normalize sheet display name
      let displaySheet = ws.name.trim();
      if (/^sheet/i.test(displaySheet)) {
        displaySheet = headers.some((h) => /保管文件/.test(h)) ? '重要保管文件' : '归档文件';
      }
      sheetNames.push(displaySheet);

      for (let i = hIdx + 1; i < rowsData.length; i++) {
        const r = rowsData[i];
        const parsed = buildArchiveRow(displaySheet, r.values, colMap, r.colBHighlighted, r.color);
        if (parsed) {
          if (parsed.isRequired) colorDetectedCount++;
          items.push(parsed);
        }
      }
    });

    // If no color was detected at all in the file, fallback heuristic: items without "（如有）" are marked required
    if (colorDetectedCount === 0 && items.length > 0) {
      items.forEach((it) => {
        if (!/（如有）|\(如有\)/.test(it.docName)) {
          it.isRequired = true;
          it.requiredSource = 'keyword';
        }
      });
    }

    return { sheetNames, items, colorDetectedCount, format: 'xlsx' };
  }

  // Legacy .xls (BIFF8) or .csv
  const wb = XLSX.read(fs.readFileSync(filePath), { type: 'buffer', cellStyles: true });
  const biffHighlights = ext === '.xls' ? extractBiff8ColumnBHighlights(filePath) : new Map<string, Set<number>>();

  for (const rawSheetName of wb.SheetNames) {
    const ws = wb.Sheets[rawSheetName];
    const aoa = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: '' });
    if (aoa.length < 2) continue;

    const headerRowIdx = aoa.findIndex((r) =>
      (r || []).some((v) => /序号|归档文件|保管文件|资料名称/.test(String(v ?? '')))
    );
    const hIdx = headerRowIdx >= 0 ? headerRowIdx : 0;
    const headers = (aoa[hIdx] || []).map((v) => String(v ?? '').trim());
    const colMap = detectColumnMap(headers);

    let displaySheet = rawSheetName.trim();
    if (/^sheet/i.test(displaySheet)) {
      displaySheet = headers.some((h) => /保管文件/.test(h)) ? '重要保管文件' : '归档文件';
    }
    sheetNames.push(displaySheet);

    const coloredSet = biffHighlights.get(rawSheetName) || new Set<number>();

    for (let i = hIdx + 1; i < aoa.length; i++) {
      const vals = (aoa[i] || []).map((v) => String(v ?? '').trim());
      const isHighlighted = coloredSet.has(i);
      const parsed = buildArchiveRow(
        displaySheet,
        vals,
        colMap,
        isHighlighted,
        isHighlighted ? 'xls-fill' : undefined
      );
      if (parsed) {
        if (parsed.isRequired) colorDetectedCount++;
        items.push(parsed);
      }
    }
  }

  if (colorDetectedCount === 0 && items.length > 0) {
    items.forEach((it) => {
      if (!/（如有）|\(如有\)/.test(it.docName)) {
        it.isRequired = true;
        it.requiredSource = 'keyword';
      }
    });
  }

  return {
    sheetNames,
    items,
    colorDetectedCount,
    format: ext === '.xls' ? 'xls' : 'other',
  };
}
