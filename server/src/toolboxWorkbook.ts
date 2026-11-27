import ExcelJS from 'exceljs';
import XLSXModule from 'xlsx';
import { evaluateWorkbookFormulas, getCellFormulaError } from './excelFormulaEval.js';

const XLSX = (XLSXModule as unknown as { default?: typeof XLSXModule }).default ?? XLSXModule;

export type RoundingMode = 'HALF_UP' | 'HALF_DOWN' | 'DOWN';

function pow10(n: number): bigint {
  if (n <= 0) return 1n;
  return 10n ** BigInt(n);
}

/**
 * Exact arbitrary-precision decimal matching Java BigDecimal semantics
 * (HALF_UP, HALF_DOWN, DOWN, divideAndRemainder, stripTrailingZeros).
 */
export class Dec {
  readonly unscaled: bigint;
  readonly scale: number;

  constructor(unscaled: bigint, scale: number) {
    this.unscaled = unscaled;
    this.scale = scale;
  }

  static ZERO = new Dec(0n, 0);
  static ONE = new Dec(1n, 0);
  static HUNDRED = new Dec(100n, 0);
  static THOUSAND = new Dec(1000n, 0);

  static of(val: string | number | bigint | Dec): Dec {
    if (val instanceof Dec) return val;
    if (typeof val === 'bigint') return new Dec(val, 0);
    const s = typeof val === 'number' ? String(val) : String(val ?? '').trim();
    if (!s) throw new Error('空数字');
    const m = s.match(/^([+-]?)(\d+)(?:\.(\d*))?(?:[eE]([+-]?\d+))?$/);
    if (!m) {
      const m2 = s.match(/^([+-]?)\.(\d+)(?:[eE]([+-]?\d+))?$/);
      if (!m2) throw new Error(`无效数字：${val}`);
      const sign = m2[1] === '-' ? -1n : 1n;
      const frac = m2[2];
      const exp = m2[3] ? parseInt(m2[3], 10) : 0;
      return Dec.fromParts(sign, frac, frac.length - exp);
    }
    const sign = m[1] === '-' ? -1n : 1n;
    const intPart = m[2];
    const fracPart = m[3] || '';
    const exp = m[4] ? parseInt(m[4], 10) : 0;
    return Dec.fromParts(sign, intPart + fracPart, fracPart.length - exp);
  }

  private static fromParts(sign: bigint, digits: string, scale: number): Dec {
    const raw = BigInt(digits || '0') * sign;
    if (scale < 0) {
      return new Dec(raw * pow10(-scale), 0);
    }
    return new Dec(raw, scale);
  }

  signum(): -1 | 0 | 1 {
    if (this.unscaled > 0n) return 1;
    if (this.unscaled < 0n) return -1;
    return 0;
  }

  isZero(): boolean {
    return this.unscaled === 0n;
  }

  abs(): Dec {
    return this.unscaled < 0n ? new Dec(-this.unscaled, this.scale) : this;
  }

  neg(): Dec {
    return new Dec(-this.unscaled, this.scale);
  }

  cmp(other: Dec): -1 | 0 | 1 {
    const s = Math.max(this.scale, other.scale);
    const a = this.unscaled * pow10(s - this.scale);
    const b = other.unscaled * pow10(s - other.scale);
    return a < b ? -1 : a > b ? 1 : 0;
  }

  eq(other: Dec): boolean {
    return this.cmp(other) === 0;
  }

  min(other: Dec): Dec {
    return this.cmp(other) <= 0 ? this : other;
  }

  max(other: Dec): Dec {
    return this.cmp(other) >= 0 ? this : other;
  }

  add(other: Dec): Dec {
    const s = Math.max(this.scale, other.scale);
    const a = this.unscaled * pow10(s - this.scale);
    const b = other.unscaled * pow10(s - other.scale);
    return new Dec(a + b, s);
  }

  sub(other: Dec): Dec {
    return this.add(other.neg());
  }

  mul(other: Dec): Dec {
    return new Dec(this.unscaled * other.unscaled, this.scale + other.scale);
  }

  private static roundDiv(num: bigint, den: bigint, mode: RoundingMode): bigint {
    if (den === 0n) throw new Error('除数不能为 0');
    const neg = (num < 0n) !== (den < 0n);
    const a = num < 0n ? -num : num;
    const b = den < 0n ? -den : den;
    let q = a / b;
    const r = a % b;
    if (r !== 0n) {
      const twiceR = r * 2n;
      if (mode === 'HALF_UP' && twiceR >= b) {
        q += 1n;
      } else if (mode === 'HALF_DOWN' && twiceR > b) {
        q += 1n;
      }
      // DOWN truncates toward zero
    }
    return neg ? -q : q;
  }

  setScale(newScale: number, mode: RoundingMode = 'HALF_UP'): Dec {
    if (newScale === this.scale) return this;
    if (newScale > this.scale) {
      return new Dec(this.unscaled * pow10(newScale - this.scale), newScale);
    }
    const factor = pow10(this.scale - newScale);
    return new Dec(Dec.roundDiv(this.unscaled, factor, mode), newScale);
  }

  div(other: Dec, scale: number, mode: RoundingMode = 'HALF_UP'): Dec {
    if (other.isZero()) throw new Error('除数不能为 0');
    // (a / 10^sa) / (b / 10^sb) = (a * 10^(sb - sa + scale)) / b / 10^scale
    const shift = other.scale - this.scale + scale;
    let num = this.unscaled;
    let den = other.unscaled;
    if (shift >= 0) {
      num *= pow10(shift);
    } else {
      den *= pow10(-shift);
    }
    return new Dec(Dec.roundDiv(num, den, mode), scale);
  }

  divExact(other: Dec): Dec {
    if (other.isZero()) throw new Error('除数不能为 0');
    // For power-of-10 divisors (like 100 or 1000) or exact divisions
    const res = this.div(other, Math.max(this.scale + 6, 12), 'HALF_UP').stripTrailingZeros();
    if (!res.mul(other).eq(this)) {
      throw new Error(`无法精确整除：${this.toPlainString()} / ${other.toPlainString()}`);
    }
    return res;
  }

  /**
   * Matches BigDecimal.divideAndRemainder(step): returns [integerQuotient, remainder].
   */
  divAndRem(step: Dec): [Dec, Dec] {
    if (step.isZero()) throw new Error('步长不能为 0');
    const s = Math.max(this.scale, step.scale);
    const a = this.unscaled * pow10(s - this.scale);
    const b = step.unscaled * pow10(s - step.scale);
    const q = a / b;
    const rem = a % b;
    return [new Dec(q, 0), new Dec(rem, s).stripTrailingZeros()];
  }

  toBigIntExact(): bigint {
    const stripped = this.stripTrailingZeros();
    if (stripped.scale > 0) throw new Error(`不是整数：${this.toPlainString()}`);
    return stripped.unscaled;
  }

  intValueExact(): number {
    const bi = this.toBigIntExact();
    if (bi < -2147483648n || bi > 2147483647n) {
      throw new Error(`整数超出范围：${bi.toString()}`);
    }
    return Number(bi);
  }

  toNumber(): number {
    return Number(this.toPlainString());
  }

  stripTrailingZeros(): Dec {
    if (this.unscaled === 0n) return Dec.ZERO;
    let u = this.unscaled;
    let s = this.scale;
    while (s > 0 && u % 10n === 0n) {
      u /= 10n;
      s -= 1;
    }
    return new Dec(u, s);
  }

  toPlainString(): string {
    if (this.unscaled === 0n) {
      return this.scale > 0 ? `0.${'0'.repeat(this.scale)}` : '0';
    }
    const neg = this.unscaled < 0n;
    const digits = (neg ? -this.unscaled : this.unscaled).toString();
    if (this.scale <= 0) {
      return (neg ? '-' : '') + digits + (this.scale < 0 ? '0'.repeat(-this.scale) : '');
    }
    if (digits.length <= this.scale) {
      return `${neg ? '-' : ''}0.${'0'.repeat(this.scale - digits.length)}${digits}`;
    }
    const cut = digits.length - this.scale;
    return `${neg ? '-' : ''}${digits.slice(0, cut)}.${digits.slice(cut)}`;
  }

  fmt(): string {
    return this.stripTrailingZeros().toPlainString();
  }
}

export interface CellVal {
  display: string;
  number: Dec | null;
  numericText: string;
}

export function readCellVal(ws: XLSXModule.WorkSheet, r: number, c: number, sheetName: string): CellVal {
  const addr = XLSX.utils.encode_cell({ r, c });
  const cell = ws[addr] as XLSXModule.CellObject | undefined;
  if (!cell || cell.t === 'z') return { display: '', number: null, numericText: '' };
  if (cell.t === 'e') {
    const fErr = getCellFormulaError(cell);
    throw new Error(fErr || `${sheetName}!${addr} 为错误单元格`);
  }
  if (cell.t === 'n' && typeof cell.v === 'number') {
    if (!Number.isFinite(cell.v)) throw new Error(`${sheetName}!${addr} 为无效数值`);
    // Match Java BigDecimal.valueOf(double) -> Double.toString(double) exact round-trip representation
    const dec = Dec.of(String(cell.v));
    const display = cell.w !== undefined && cell.w !== null ? String(cell.w) : dec.fmt();
    return { display, number: dec, numericText: dec.fmt() };
  }
  const display =
    cell.w !== undefined && cell.w !== null
      ? String(cell.w)
      : cell.v !== undefined && cell.v !== null
        ? String(cell.v)
        : '';
  return { display, number: null, numericText: display };
}

export interface ParsedWorkbookTable {
  rows: Record<string, string>[];
  headerRow: number;
}

function getHeaderCellsForRow(ws: XLSXModule.WorkSheet, rowIdx: number, sheetName: string): Map<number, string> {
  const map = new Map<number, string>();
  const ref = ws['!ref'];
  if (!ref) return map;
  const range = XLSX.utils.decode_range(ref);
  for (let c = range.s.c; c <= range.e.c; c++) {
    const v = readCellVal(ws, rowIdx, c, sheetName).display.trim();
    if (v) map.set(c, v);
  }
  // Vertical header merge inheritance (matches WorkbookTable.headerCells)
  const merges = ws['!merges'] || [];
  for (const m of merges) {
    if (m.s.c === m.e.c && m.s.r <= rowIdx && m.e.r >= rowIdx) {
      const anchorVal = readCellVal(ws, m.s.r, m.s.c, sheetName).display.trim();
      if (anchorVal) map.set(m.s.c, anchorVal);
    }
  }
  return map;
}

function readSheetTable(
  ws: XLSXModule.WorkSheet,
  sheetName: string,
  requiredCols: Record<string, string[]>,
  optionalCols: Record<string, string[]> = {}
): ParsedWorkbookTable {
  const ref = ws['!ref'];
  if (!ref) {
    throw new Error(`${sheetName} 缺少必要表头：${Object.keys(requiredCols).join('、')}`);
  }
  const range = XLSX.utils.decode_range(ref);
  const reqEntries = Object.entries(requiredCols);

  let indexes: Record<string, number> | null = null;
  let header = -1;
  let bestKeys: Set<string> = new Set();
  let bestRow = -1;

  for (let r = range.s.r; r <= range.e.r; r++) {
    const hCells = getHeaderCellsForRow(ws, r, sheetName);
    const found: Record<string, number> = {};
    for (const [key, aliases] of reqEntries) {
      for (const [colIdx, text] of hCells.entries()) {
        if (aliases.includes(text)) {
          if (found[key] !== undefined) {
            throw new Error(`表头重复：${key}，行${r + 1}`);
          }
          found[key] = colIdx;
        }
      }
    }
    const foundKeys = Object.keys(found);
    if (foundKeys.length > bestKeys.size) {
      bestKeys = new Set(foundKeys);
      bestRow = r + 1;
    }
    if (foundKeys.length === reqEntries.length) {
      indexes = found;
      header = r;
      break;
    }
  }

  if (!indexes) {
    const missing = Object.keys(requiredCols).filter((k) => !bestKeys.has(k));
    throw new Error(
      `${sheetName} 缺少必要表头：${missing.join('、')}${bestRow > 0 ? `（最接近的表头在第${bestRow}行）` : ''}`
    );
  }

  const hCells = getHeaderCellsForRow(ws, header, sheetName);
  for (const [key, aliases] of Object.entries(optionalCols)) {
    for (const [colIdx, text] of hCells.entries()) {
      if (aliases.includes(text)) {
        if (indexes[key] !== undefined) {
          throw new Error(`表头重复：${key}，行${header + 1}`);
        }
        indexes[key] = colIdx;
      }
    }
  }

  let dataStart = header + 1;
  const colValues = new Set(Object.values(indexes));
  for (const m of ws['!merges'] || []) {
    if (m.s.c === m.e.c && colValues.has(m.s.c) && m.s.r <= header && m.e.r >= header) {
      dataStart = Math.max(dataStart, m.e.r + 1);
    }
  }

  const rows: Record<string, string>[] = [];
  for (let i = dataStart; i <= range.e.r; i++) {
    const values: Record<string, string> = {};
    let empty = true;
    for (const [key, colIdx] of Object.entries(indexes)) {
      const parsed = readCellVal(ws, i, colIdx, sheetName);
      values[key] = parsed.display;
      if (parsed.number !== null) {
        values[`__numeric:${key}`] = parsed.numericText;
      }
      if (parsed.display.trim() !== '') empty = false;
    }
    if (!empty) {
      values.__row = String(i + 1);
      rows.push(values);
    }
  }

  return { rows, headerRow: header + 1 };
}

export function readWorkbookFirstSheet(
  buffer: Buffer,
  fileName: string,
  requiredCols: Record<string, string[]>,
  optionalCols: Record<string, string[]> = {}
): ParsedWorkbookTable {
  let wb: XLSXModule.WorkBook;
  try {
    wb = XLSX.read(buffer, { type: 'buffer', cellDates: false, sheetStubs: true, cellNF: true, cellStyles: true });
  } catch (err) {
    throw new Error(`${fileName} 无法读取：${(err as Error).message}`);
  }
  evaluateWorkbookFormulas(wb);
  if (wb.SheetNames.length === 0) throw new Error(`工作簿没有工作表：${fileName}`);
  const first = wb.SheetNames[0];
  return readSheetTable(wb.Sheets[first], first, requiredCols, optionalCols);
}

export function readWorkbookAllSheets(
  buffer: Buffer,
  fileName: string,
  requiredCols: Record<string, string[]>,
  optionalCols: Record<string, string[]> = {}
): Map<string, ParsedWorkbookTable> {
  let wb: XLSXModule.WorkBook;
  try {
    wb = XLSX.read(buffer, { type: 'buffer', cellDates: false, sheetStubs: true, cellNF: true, cellStyles: true });
  } catch (err) {
    throw new Error(`${fileName} 无法读取：${(err as Error).message}`);
  }
  evaluateWorkbookFormulas(wb);
  const result = new Map<string, ParsedWorkbookTable>();
  for (const name of wb.SheetNames) {
    result.set(name, readSheetTable(wb.Sheets[name], name, requiredCols, optionalCols));
  }
  return result;
}

export function numericText(row: Record<string, string>, key: string): string {
  return row[`__numeric:${key}`] ?? row[key] ?? '';
}

export function colsFromList(names: string[]): Record<string, string[]> {
  const out: Record<string, string[]> = {};
  for (const n of names) out[n] = [n];
  return out;
}

export interface ExcelExportSheet {
  name: string;
  headers: string[];
  rows: (string | number)[][];
}

/**
 * Unified workstation Excel exporter: pure black Calibri 11pt, bold header without fill,
 * all thin black borders, AutoFilter, freeze top row, printTitlesRow.
 */
export async function buildStyledExcelBuffer(sheets: ExcelExportSheet[]): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FF000000' } },
  };
  const headerFont: Partial<ExcelJS.Font> = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF000000' } };
  const bodyFont: Partial<ExcelJS.Font> = { name: 'Calibri', size: 11, color: { argb: 'FF000000' } };

  for (const spec of sheets) {
    const safeName = spec.name.replace(/[\\/*?:[\]]/g, '_').slice(0, 31) || 'Sheet1';
    const ws = wb.addWorksheet(safeName, {
      views: [{ state: 'frozen', ySplit: 1 }],
      pageSetup: { printTitlesRow: '1:1' },
    });

    const headerRow = ws.addRow(spec.headers);
    headerRow.height = 22;
    headerRow.eachCell({ includeEmpty: true }, (cell) => {
      cell.font = headerFont;
      cell.border = thinBorder;
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    });

    // Compute column widths first (matches WorkbookTable.write: up to 100 chars for multiline columns)
    const colWidths: number[] = [];
    for (let c = 0; c < spec.headers.length; c++) {
      let maxLen = Math.max(14, displayWidth(spec.headers[c]) + 4);
      let hasMultiline = false;
      for (const r of spec.rows) {
        const val = String(r[c] ?? '');
        if (val.includes('\n')) hasMultiline = true;
        for (const line of val.split('\n')) {
          maxLen = Math.max(maxLen, displayWidth(line) + 3);
        }
      }
      const width = Math.min(hasMultiline ? 100 : 54, Math.max(12, maxLen));
      colWidths.push(width);
      ws.getColumn(c + 1).width = width;
    }

    for (const r of spec.rows) {
      const row = ws.addRow(r);
      let heightLines = 1;
      for (let c = 0; c < spec.headers.length; c++) {
        const val = String(r[c] ?? '');
        if (val.includes('\n')) {
          const effWidth = Math.max(1, Math.floor((colWidths[c] || 14) - 2));
          let lines = 0;
          for (const line of val.split('\n')) {
            lines += Math.max(1, Math.ceil(displayWidth(line) / effWidth));
          }
          heightLines = Math.max(heightLines, lines);
        }
      }
      row.height = heightLines > 1 ? Math.min(409, heightLines * 17 + 4) : 20;
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const val = String(r[colNumber - 1] ?? '');
        cell.font = bodyFont;
        cell.border = thinBorder;
        cell.alignment = { vertical: val.includes('\n') ? 'top' : 'middle', wrapText: true };
      });
    }

    if (spec.headers.length > 0) {
      ws.autoFilter = {
        from: { row: 1, column: 1 },
        to: { row: Math.max(1, spec.rows.length + 1), column: spec.headers.length },
      };
    }
  }

  const arr = await wb.xlsx.writeBuffer();
  return Buffer.from(arr);
}

function displayWidth(text: string): number {
  let w = 0;
  for (const ch of text) {
    w += ch.charCodeAt(0) > 255 ? 2 : 1;
  }
  return w;
}
