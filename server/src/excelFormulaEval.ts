import XLSXModule from 'xlsx';

const XLSX = (XLSXModule as unknown as { default?: typeof XLSXModule }).default ?? XLSXModule;

const UNSUPPORTED_ERR = '#UNSUPPORTED!';

type EvalVal =
  | { kind: 'num'; v: number }
  | { kind: 'str'; v: string }
  | { kind: 'bool'; v: boolean }
  | { kind: 'empty' }
  | { kind: 'err'; code: string };

class UnsupportedFormulaError extends Error {
  constructor(msg: string) {
    super(msg);
    this.name = 'UnsupportedFormulaError';
  }
}

export interface FormulaErrorCell extends XLSXModule.CellObject {
  __formulaError?: string;
}

export function getCellFormulaError(cell: XLSXModule.CellObject | undefined): string | null {
  if (!cell) return null;
  const msg = (cell as FormulaErrorCell).__formulaError;
  return msg || null;
}

function isUnsupportedErr(v: EvalVal): boolean {
  return v.kind === 'err' && v.code === UNSUPPORTED_ERR;
}

type TokType = 'num' | 'str' | 'ident' | 'op' | 'lparen' | 'rparen' | 'comma' | 'colon' | 'eof';

interface Token {
  type: TokType;
  value: string;
}

function tokenize(formula: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  const n = formula.length;

  while (i < n) {
    const ch = formula[i];
    if (/\s/.test(ch)) {
      i++;
      continue;
    }
    if (ch === '"') {
      let s = '';
      i++;
      while (i < n) {
        if (formula[i] === '"') {
          if (i + 1 < n && formula[i + 1] === '"') {
            s += '"';
            i += 2;
          } else {
            i++;
            break;
          }
        } else {
          s += formula[i++];
        }
      }
      tokens.push({ type: 'str', value: s });
      continue;
    }
    // Quoted sheet reference: 'Sheet Name'!A1 or 'Sheet Name'!$D:$D
    if (ch === "'") {
      let ref = "'";
      i++;
      while (i < n) {
        if (formula[i] === "'") {
          ref += "'";
          i++;
          if (i < n && formula[i] === "'") {
            ref += "'";
            i++;
          } else {
            break;
          }
        } else {
          ref += formula[i++];
        }
      }
      if (i < n && formula[i] === '!') {
        ref += '!';
        i++;
        while (i < n && /[A-Za-z0-9_$]/.test(formula[i])) {
          ref += formula[i++];
        }
      }
      tokens.push({ type: 'ident', value: ref });
      continue;
    }
    // Numbers (including scientific notation like 1e-5)
    if (/[0-9]/.test(ch) || (ch === '.' && i + 1 < n && /[0-9]/.test(formula[i + 1]))) {
      let num = '';
      while (i < n && /[0-9.]/.test(formula[i])) num += formula[i++];
      if (i < n && (formula[i] === 'e' || formula[i] === 'E')) {
        const next = formula[i + 1];
        if (next && (/[0-9]/.test(next) || ((next === '+' || next === '-') && /[0-9]/.test(formula[i + 2] || '')))) {
          num += formula[i++];
          if (formula[i] === '+' || formula[i] === '-') num += formula[i++];
          while (i < n && /[0-9]/.test(formula[i])) num += formula[i++];
        }
      }
      tokens.push({ type: 'num', value: num });
      continue;
    }
    // Two-character comparison operators
    if ((ch === '<' && (formula[i + 1] === '>' || formula[i + 1] === '=')) || (ch === '>' && formula[i + 1] === '=')) {
      tokens.push({ type: 'op', value: ch + formula[i + 1] });
      i += 2;
      continue;
    }
    if ('+-*/^&=<%>'.includes(ch)) {
      tokens.push({ type: 'op', value: ch });
      i++;
      continue;
    }
    if (ch === '(') {
      tokens.push({ type: 'lparen', value: '(' });
      i++;
      continue;
    }
    if (ch === ')') {
      tokens.push({ type: 'rparen', value: ')' });
      i++;
      continue;
    }
    if (ch === ',') {
      tokens.push({ type: 'comma', value: ',' });
      i++;
      continue;
    }
    if (ch === ':') {
      tokens.push({ type: 'colon', value: ':' });
      i++;
      continue;
    }
    // Excel error literal like #N/A, #DIV/0!, #VALUE!, #REF!, #NAME?, #NUM!
    if (ch === '#') {
      let err = '#';
      i++;
      while (i < n && /[A-Za-z0-9/!?]/.test(formula[i])) err += formula[i++];
      tokens.push({ type: 'ident', value: err });
      continue;
    }
    // Identifiers: functions, booleans, cell references, Sheet1!A1, $D, $1
    if (/[A-Za-z_$\u4e00-\u9fa5]/.test(ch)) {
      let id = '';
      while (i < n && /[A-Za-z0-9_.$!\u4e00-\u9fa5]/.test(formula[i])) {
        id += formula[i++];
      }
      tokens.push({ type: 'ident', value: id });
      continue;
    }
    throw new UnsupportedFormulaError(`无法解析公式字符：${ch}`);
  }
  tokens.push({ type: 'eof', value: '' });
  return tokens;
}

const CELL_REF_RE = /^(?:(?:'((?:[^']|'')+)'|([^'!]+))!)?\$?([A-Za-z]{1,3})\$?(\d+)$/;
const COL_ONLY_REF_RE = /^(?:(?:'((?:[^']|'')+)'|([^'!]+))!)?\$?([A-Za-z]{1,3})$/;
const ROW_ONLY_REF_RE = /^(?:(?:'((?:[^']|'')+)'|([^'!]+))!)?\$?(\d+)$/;

function parseCellRef(token: string, currentSheet: string): { sheet: string; r: number; c: number } | null {
  const m = token.match(CELL_REF_RE);
  if (!m) return null;
  const sheet = m[1] ? m[1].replace(/''/g, "'") : m[2] ? m[2] : currentSheet;
  const colStr = m[3].toUpperCase();
  const rowNum = parseInt(m[4], 10);
  if (rowNum < 1 || rowNum > 1048576) return null;
  const decoded = XLSX.utils.decode_cell(`${colStr}${rowNum}`);
  return { sheet, r: decoded.r, c: decoded.c };
}

function parseColOnlyRef(token: string, currentSheet: string): { sheet: string; c: number } | null {
  const m = token.match(COL_ONLY_REF_RE);
  if (!m) return null;
  const sheet = m[1] ? m[1].replace(/''/g, "'") : m[2] ? m[2] : currentSheet;
  const colStr = m[3].toUpperCase();
  const c = XLSX.utils.decode_col(colStr);
  if (c < 0 || c > 16383) return null;
  return { sheet, c };
}

function parseRowOnlyRef(token: string, currentSheet: string): { sheet: string; r: number } | null {
  const m = token.match(ROW_ONLY_REF_RE);
  if (!m) return null;
  const sheet = m[1] ? m[1].replace(/''/g, "'") : m[2] ? m[2] : currentSheet;
  const rowNum = parseInt(m[3], 10);
  if (rowNum < 1 || rowNum > 1048576) return null;
  return { sheet, r: rowNum - 1 };
}

function toNumber(v: EvalVal): { ok: true; n: number } | { ok: false; err: string } {
  if (v.kind === 'err') return { ok: false, err: v.code };
  if (v.kind === 'num') return { ok: true, n: v.v };
  if (v.kind === 'empty') return { ok: true, n: 0 };
  if (v.kind === 'bool') return { ok: true, n: v.v ? 1 : 0 };
  const s = v.v.trim();
  if (!s) return { ok: true, n: 0 };
  const n = Number(s);
  if (Number.isFinite(n)) return { ok: true, n };
  return { ok: false, err: '#VALUE!' };
}

function toStr(v: EvalVal): string {
  if (v.kind === 'str') return v.v;
  if (v.kind === 'num') return String(v.v);
  if (v.kind === 'bool') return v.v ? 'TRUE' : 'FALSE';
  if (v.kind === 'err') return v.code;
  return '';
}

function toBool(v: EvalVal): { ok: true; b: boolean } | { ok: false; err: string } {
  if (v.kind === 'err') return { ok: false, err: v.code };
  if (v.kind === 'bool') return { ok: true, b: v.v };
  if (v.kind === 'num') return { ok: true, b: v.v !== 0 };
  if (v.kind === 'empty') return { ok: true, b: false };
  const upper = v.v.trim().toUpperCase();
  if (upper === 'TRUE') return { ok: true, b: true };
  if (upper === 'FALSE') return { ok: true, b: false };
  return { ok: false, err: '#VALUE!' };
}

function compareEvalVals(a: EvalVal, b: EvalVal): number {
  if (a.kind === 'num' && b.kind === 'num') return a.v < b.v ? -1 : a.v > b.v ? 1 : 0;
  if (a.kind === 'empty' && b.kind === 'num') return 0 < b.v ? -1 : 0 > b.v ? 1 : 0;
  if (a.kind === 'num' && b.kind === 'empty') return a.v < 0 ? -1 : a.v > 0 ? 1 : 0;
  if (a.kind === 'empty' && b.kind === 'empty') return 0;
  if (a.kind === 'empty' && b.kind === 'str') return '' < b.v ? -1 : '' > b.v ? 1 : 0;
  if (a.kind === 'str' && b.kind === 'empty') return a.v < '' ? -1 : a.v > '' ? 1 : 0;
  if (a.kind === 'str' && b.kind === 'str') {
    const la = a.v.toLowerCase();
    const lb = b.v.toLowerCase();
    return la < lb ? -1 : la > lb ? 1 : 0;
  }
  if (a.kind === 'bool' && b.kind === 'bool') return a.v === b.v ? 0 : a.v ? 1 : -1;
  // Excel type ordering: number < text < boolean
  const rank = (x: EvalVal) => (x.kind === 'num' || x.kind === 'empty' ? 0 : x.kind === 'str' ? 1 : 2);
  return rank(a) - rank(b);
}

function roundAwayFromZero(val: number, digits: number, mode: 'round' | 'up' | 'down'): number {
  const factor = Math.pow(10, digits);
  const scaled = val * factor;
  if (mode === 'up') {
    return (scaled >= 0 ? Math.ceil(scaled - 1e-12) : Math.floor(scaled + 1e-12)) / factor;
  }
  if (mode === 'down') {
    return Math.trunc(scaled + (scaled >= 0 ? 1e-12 : -1e-12)) / factor;
  }
  return (scaled >= 0 ? Math.floor(scaled + 0.5 + 1e-12) : Math.ceil(scaled - 0.5 - 1e-12)) / factor;
}

function wildcardToRegex(pattern: string): RegExp {
  let re = '^';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length) {
      const next = pattern[++i];
      re += next.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    } else if (ch === '*') {
      re += '[\\s\\S]*';
    } else if (ch === '?') {
      re += '[\\s\\S]';
    } else {
      re += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
  }
  re += '$';
  return new RegExp(re, 'i');
}

function matchesCriteria(val: EvalVal, criteria: EvalVal): boolean {
  if (val.kind === 'err') {
    return criteria.kind === 'err' && val.code === criteria.code;
  }
  if (criteria.kind === 'err') return false;
  if (criteria.kind === 'empty') {
    return val.kind === 'empty' || (val.kind === 'str' && val.v === '');
  }
  if (criteria.kind === 'num') {
    if (val.kind === 'num') return val.v === criteria.v;
    if (val.kind === 'str' && val.v.trim() !== '') {
      const n = Number(val.v.trim());
      return Number.isFinite(n) && n === criteria.v;
    }
    return false;
  }
  if (criteria.kind === 'bool') {
    return val.kind === 'bool' && val.v === criteria.v;
  }

  // String criteria (may contain operator prefix >=, <=, <>, >, <, =)
  const raw = criteria.v;
  const opMatch = raw.match(/^(<>|>=|<=|>|<|=)([\s\S]*)$/);
  const op = opMatch ? opMatch[1] : '=';
  const rhsText = opMatch ? opMatch[2] : raw;

  if (op === '=' && rhsText === '') {
    return val.kind === 'empty' || (val.kind === 'str' && val.v === '');
  }
  if (op === '<>' && rhsText === '') {
    return !(val.kind === 'empty' || (val.kind === 'str' && val.v === ''));
  }

  const rhsNum = rhsText.trim() !== '' ? Number(rhsText.trim()) : NaN;
  const rhsIsNum = Number.isFinite(rhsNum);

  if (op === '>' || op === '>=' || op === '<' || op === '<=') {
    if (val.kind === 'empty') return false;
    if (rhsIsNum && val.kind === 'num') {
      return op === '>' ? val.v > rhsNum : op === '>=' ? val.v >= rhsNum : op === '<' ? val.v < rhsNum : val.v <= rhsNum;
    }
    if (!rhsIsNum && val.kind === 'str') {
      const a = val.v.toLowerCase();
      const b = rhsText.toLowerCase();
      return op === '>' ? a > b : op === '>=' ? a >= b : op === '<' ? a < b : a <= b;
    }
    return false;
  }

  // Equality (=) or Inequality (<>)
  let equal = false;
  if (rhsIsNum) {
    if (val.kind === 'num') {
      equal = val.v === rhsNum;
    } else if (val.kind === 'str' && val.v.trim() !== '') {
      const n = Number(val.v.trim());
      equal = Number.isFinite(n) ? n === rhsNum : val.v.toLowerCase() === rhsText.toLowerCase();
    }
  } else if (rhsText.toUpperCase() === 'TRUE' || rhsText.toUpperCase() === 'FALSE') {
    if (val.kind === 'bool') equal = val.v === (rhsText.toUpperCase() === 'TRUE');
    else if (val.kind === 'str') equal = val.v.toLowerCase() === rhsText.toLowerCase();
  } else {
    const hasWildcard = /[*?~]/.test(rhsText);
    if (val.kind === 'str') {
      equal = hasWildcard ? wildcardToRegex(rhsText).test(val.v) : val.v.toLowerCase() === rhsText.toLowerCase();
    } else if (val.kind === 'num') {
      // Excel rule: wildcard criteria (* or ?) ONLY match text cells, NEVER numeric cells
      equal = !hasWildcard && String(val.v) === rhsText;
    }
  }
  return op === '<>' ? !equal : equal;
}

function excelWeekday(serial: number, returnType: number): number | null {
  const s = Math.floor(serial);
  if (s < 1) return null;
  // Excel 1900 date system: serial 1 = 1900-01-01 (Sunday); serial 60 = 1900-02-29 (fictional);
  // for serial >= 61, Date.UTC(1899, 11, 30) + s*86400000 gives exact Gregorian UTC day (0=Sun..6=Sat).
  // For serial 1..59, Day 1 is Sunday (0), so (s - 1) % 7 gives exact 0=Sun..6=Sat!
  const dowSun0 =
    s < 60
      ? (s - 1) % 7
      : new Date(Date.UTC(1899, 11, 30) + s * 86400000).getUTCDay();
  switch (returnType) {
    case 1:
    case 17:
      return dowSun0 + 1; // 1=Sun .. 7=Sat
    case 2:
    case 11:
      return dowSun0 === 0 ? 7 : dowSun0; // 1=Mon .. 7=Sun
    case 3:
      return dowSun0 === 0 ? 6 : dowSun0 - 1; // 0=Mon .. 6=Sun
    case 12:
      return ((dowSun0 + 5) % 7) + 1; // 1=Tue .. 7=Mon
    case 13:
      return ((dowSun0 + 4) % 7) + 1; // 1=Wed .. 7=Tue
    case 14:
      return ((dowSun0 + 3) % 7) + 1; // 1=Thu .. 7=Wed
    case 15:
      return ((dowSun0 + 2) % 7) + 1; // 1=Fri .. 7=Thu
    case 16:
      return ((dowSun0 + 1) % 7) + 1; // 1=Sat .. 7=Fri
    default:
      return null;
  }
}

/**
 * Evaluate all formula cells in an XLSX.WorkBook in-place so that stale or missing
 * cached values are recomputed identically to Excel / Apache POI FormulaEvaluator.
 *
 * Unsupported formulas are marked as `#UNSUPPORTED!` (with `__formulaError`) and are
 * strictly propagated through `IFERROR`, `COUNTIF`, `SUMIF`, etc., so unused auxiliary
 * columns never block processing while any actually-read unsupported formula fails loudly.
 */
export function evaluateWorkbookFormulas(wb: XLSXModule.WorkBook): void {
  const memo = new Map<string, EvalVal>();
  const visiting = new Set<string>();

  const evalCell = (sheetName: string, r: number, c: number): EvalVal => {
    const ws = wb.Sheets[sheetName];
    if (!ws) return { kind: 'err', code: '#REF!' };
    const addr = XLSX.utils.encode_cell({ r, c });
    const key = `${sheetName}!${addr}`;
    if (memo.has(key)) return memo.get(key)!;

    const cell = ws[addr] as FormulaErrorCell | undefined;
    if (!cell) return { kind: 'empty' };

    if (cell.f) {
      if (visiting.has(key)) return { kind: 'err', code: '#REF!' };
      visiting.add(key);
      try {
        const res = evalFormulaText(String(cell.f), sheetName);
        memo.set(key, res);
        applyEvalResultToCell(cell, res, sheetName, addr);
        return res;
      } catch (err) {
        if (err instanceof UnsupportedFormulaError) {
          const msg = `${sheetName}!${addr} 包含暂不支持自动求值的公式（=${cell.f}），请在 Excel 中重新计算并粘贴为数值后重试`;
          const errVal: EvalVal = { kind: 'err', code: UNSUPPORTED_ERR };
          memo.set(key, errVal);
          cell.t = 'e';
          cell.v = 0;
          cell.w = UNSUPPORTED_ERR;
          cell.__formulaError = msg;
          return errVal;
        }
        throw err;
      } finally {
        visiting.delete(key);
      }
    }

    const staticVal = readStaticCell(cell);
    memo.set(key, staticVal);
    return staticVal;
  };

  const readStaticCell = (cell: XLSXModule.CellObject): EvalVal => {
    if (cell.t === 'z') return { kind: 'empty' };
    if (cell.t === 'e') return { kind: 'err', code: String(cell.w || cell.v || '#VALUE!') };
    if (cell.t === 'n' && typeof cell.v === 'number') return { kind: 'num', v: cell.v };
    if (cell.t === 'b') return { kind: 'bool', v: Boolean(cell.v) };
    if (cell.t === 's' || typeof cell.v === 'string') return { kind: 'str', v: String(cell.v ?? '') };
    if (cell.v === undefined || cell.v === null) return { kind: 'empty' };
    return { kind: 'str', v: String(cell.v) };
  };

  const applyEvalResultToCell = (cell: FormulaErrorCell, res: EvalVal, sheetName: string, addr: string) => {
    if (res.kind === 'num') {
      cell.t = 'n';
      cell.v = res.v;
      const fmt = typeof cell.z === 'string' ? cell.z.trim() : '';
      if (fmt && fmt !== 'General' && fmt !== '@') {
        try {
          cell.w = XLSX.SSF.format(fmt, res.v);
        } catch {
          cell.w = fmt.includes('%') ? `${Number((res.v * 100).toPrecision(15))}%` : String(res.v);
        }
      } else {
        cell.w = String(res.v);
      }
    } else if (res.kind === 'str') {
      cell.t = 's';
      cell.v = res.v;
      cell.w = res.v;
    } else if (res.kind === 'bool') {
      cell.t = 'b';
      cell.v = res.v;
      cell.w = res.v ? 'TRUE' : 'FALSE';
    } else if (res.kind === 'empty') {
      cell.t = 's';
      cell.v = '';
      cell.w = '';
    } else if (res.kind === 'err') {
      cell.t = 'e';
      cell.v = 0;
      cell.w = res.code;
      if (res.code === UNSUPPORTED_ERR && !cell.__formulaError) {
        cell.__formulaError = `${sheetName}!${addr} 引用的公式包含暂不支持自动求值的函数（=${cell.f || ''}），请在 Excel 中重新计算并粘贴为数值后重试`;
      }
    }
  };

  const evalRange = (
    startRef: { sheet: string; r: number; c: number },
    endRef: { sheet: string; r: number; c: number }
  ): { vals: EvalVal[]; rows: number; cols: number; origin: { sheet: string; r: number; c: number } } => {
    const sheet = startRef.sheet;
    const r1 = Math.min(startRef.r, endRef.r);
    const r2 = Math.max(startRef.r, endRef.r);
    const c1 = Math.min(startRef.c, endRef.c);
    const c2 = Math.max(startRef.c, endRef.c);
    const rows = r2 - r1 + 1;
    const cols = c2 - c1 + 1;
    if (rows * cols > 200_000) {
      throw new UnsupportedFormulaError('公式引用区域过大');
    }
    const vals: EvalVal[] = [];
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        vals.push(evalCell(sheet, r, c));
      }
    }
    return { vals, rows, cols, origin: { sheet, r: r1, c: c1 } };
  };

  type CellOrigin = { sheet: string; r: number; c: number };
  type ArgVal =
    | { isRange: false; isRef: boolean; origin?: CellOrigin; val: EvalVal }
    | { isRange: true; origin: CellOrigin; vals: EvalVal[]; rows: number; cols: number; isWholeCol?: boolean; isWholeRow?: boolean };

  const evalFormulaText = (rawFormula: string, currentSheet: string): EvalVal => {
    const cleaned = rawFormula.trim().replace(/^=/, '');
    const tokens = tokenize(cleaned);
    let pos = 0;

    const peek = () => tokens[pos];
    const consume = () => tokens[pos++];

    const skipBalancedArg = () => {
      let depth = 0;
      while (peek().type !== 'eof') {
        const t = peek().type;
        if (t === 'lparen') {
          depth++;
          consume();
        } else if (t === 'rparen') {
          if (depth === 0) break;
          depth--;
          consume();
        } else if (t === 'comma' && depth === 0) {
          break;
        } else {
          consume();
        }
      }
    };

    const parseComparison = (): EvalVal => {
      let left = parseConcat();
      while (peek().type === 'op' && ['=', '<>', '<', '<=', '>', '>='].includes(peek().value)) {
        const op = consume().value;
        const right = parseConcat();
        if (left.kind === 'err') return left;
        if (right.kind === 'err') return right;
        const cmp = compareEvalVals(left, right);
        const b =
          op === '='
            ? cmp === 0
            : op === '<>'
              ? cmp !== 0
              : op === '<'
                ? cmp < 0
                : op === '<='
                  ? cmp <= 0
                  : op === '>'
                    ? cmp > 0
                    : cmp >= 0;
        left = { kind: 'bool', v: b };
      }
      return left;
    };

    const parseConcat = (): EvalVal => {
      let left = parseAddSub();
      while (peek().type === 'op' && peek().value === '&') {
        consume();
        const right = parseAddSub();
        if (left.kind === 'err') return left;
        if (right.kind === 'err') return right;
        left = { kind: 'str', v: toStr(left) + toStr(right) };
      }
      return left;
    };

    const parseAddSub = (): EvalVal => {
      let left = parseMulDiv();
      while (peek().type === 'op' && (peek().value === '+' || peek().value === '-')) {
        const op = consume().value;
        const right = parseMulDiv();
        const ln = toNumber(left);
        if (!ln.ok) return { kind: 'err', code: ln.err };
        const rn = toNumber(right);
        if (!rn.ok) return { kind: 'err', code: rn.err };
        const res = op === '+' ? ln.n + rn.n : ln.n - rn.n;
        left = { kind: 'num', v: Number(res.toPrecision(15)) };
      }
      return left;
    };

    const parseMulDiv = (): EvalVal => {
      let left = parsePower();
      while (peek().type === 'op' && (peek().value === '*' || peek().value === '/')) {
        const op = consume().value;
        const right = parsePower();
        const ln = toNumber(left);
        if (!ln.ok) return { kind: 'err', code: ln.err };
        const rn = toNumber(right);
        if (!rn.ok) return { kind: 'err', code: rn.err };
        if (op === '/') {
          if (rn.n === 0) return { kind: 'err', code: '#DIV/0!' };
          left = { kind: 'num', v: Number((ln.n / rn.n).toPrecision(15)) };
        } else {
          left = { kind: 'num', v: Number((ln.n * rn.n).toPrecision(15)) };
        }
      }
      return left;
    };

    const parsePower = (): EvalVal => {
      let left = parseUnary();
      if (peek().type === 'op' && peek().value === '^') {
        consume();
        const right = parsePower();
        const ln = toNumber(left);
        if (!ln.ok) return { kind: 'err', code: ln.err };
        const rn = toNumber(right);
        if (!rn.ok) return { kind: 'err', code: rn.err };
        const res = Math.pow(ln.n, rn.n);
        if (!Number.isFinite(res)) return { kind: 'err', code: '#NUM!' };
        left = { kind: 'num', v: res };
      }
      return left;
    };

    const parseUnary = (): EvalVal => {
      if (peek().type === 'op' && (peek().value === '+' || peek().value === '-')) {
        const op = consume().value;
        const val = parseUnary();
        const n = toNumber(val);
        if (!n.ok) return { kind: 'err', code: n.err };
        return { kind: 'num', v: op === '-' ? -n.n : n.n };
      }
      return parsePostfix();
    };

    const parsePostfix = (): EvalVal => {
      let val = parsePrimary();
      while (peek().type === 'op' && peek().value === '%') {
        consume();
        const n = toNumber(val);
        if (!n.ok) return { kind: 'err', code: n.err };
        val = { kind: 'num', v: Number((n.n / 100).toPrecision(15)) };
      }
      return val;
    };

    const parseArg = (): ArgVal => {
      // Check for range reference: A1:B5, $D:$D, Sheet1!$D:$D, 1:5, $1:$5
      if (
        (peek().type === 'ident' || peek().type === 'num') &&
        tokens[pos + 1]?.type === 'colon' &&
        (tokens[pos + 2]?.type === 'ident' || tokens[pos + 2]?.type === 'num')
      ) {
        const startTok = consume().value;
        consume(); // ':'
        const endTok = consume().value;

        // 1. Standard cell:cell range (A1:B5)
        const startRef = parseCellRef(startTok, currentSheet);
        const endRef = parseCellRef(endTok, startRef?.sheet || currentSheet);
        if (startRef && endRef) {
          const grid = evalRange(startRef, endRef);
          return { isRange: true, origin: grid.origin, vals: grid.vals, rows: grid.rows, cols: grid.cols };
        }

        // 2. Whole-column range ($D:$D, A:C, Sheet1!$D:$D): ALWAYS start at row 0 so cross-sheet columns stay row-aligned!
        const startCol = parseColOnlyRef(startTok, currentSheet);
        const endCol = parseColOnlyRef(endTok, startCol?.sheet || currentSheet);
        if (startCol && endCol) {
          const targetWs = wb.Sheets[startCol.sheet];
          const refRange = targetWs && targetWs['!ref'] ? XLSX.utils.decode_range(targetWs['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
          const grid = evalRange(
            { sheet: startCol.sheet, r: 0, c: startCol.c },
            { sheet: startCol.sheet, r: Math.max(0, refRange.e.r), c: endCol.c }
          );
          return { isRange: true, origin: grid.origin, vals: grid.vals, rows: grid.rows, cols: grid.cols, isWholeCol: true };
        }

        // 3. Whole-row range (1:5, $1:$5): ALWAYS start at col 0!
        const startRow = parseRowOnlyRef(startTok, currentSheet);
        const endRow = parseRowOnlyRef(endTok, startRow?.sheet || currentSheet);
        if (startRow && endRow) {
          const targetWs = wb.Sheets[startRow.sheet];
          const refRange = targetWs && targetWs['!ref'] ? XLSX.utils.decode_range(targetWs['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
          const grid = evalRange(
            { sheet: startRow.sheet, r: startRow.r, c: 0 },
            { sheet: startRow.sheet, r: endRow.r, c: Math.max(0, refRange.e.c) }
          );
          return { isRange: true, origin: grid.origin, vals: grid.vals, rows: grid.rows, cols: grid.cols, isWholeRow: true };
        }

        throw new UnsupportedFormulaError(`无效区域引用：${startTok}:${endTok}`);
      }

      // Check for standalone single-cell reference (e.g., SUM(A1), SUM(F7:H7, E7), SUMIF(H2:H3, D2, I2))
      if (
        peek().type === 'ident' &&
        (tokens[pos + 1]?.type === 'comma' || tokens[pos + 1]?.type === 'rparen')
      ) {
        const ref = parseCellRef(peek().value, currentSheet);
        if (ref) {
          consume();
          return { isRange: false, isRef: true, origin: ref, val: evalCell(ref.sheet, ref.r, ref.c) };
        }
      }
      return { isRange: false, isRef: false, val: parseComparison() };
    };

    const evalFunction = (rawName: string, args: ArgVal[]): EvalVal => {
      const name = rawName.toUpperCase();

      // Excel aggregation semantics: cell references (single cell or range) ignore text/bool/empty;
      // direct literal/expression arguments coerce to number.
      const flattenNums = (list: ArgVal[]): { ok: true; nums: number[] } | { ok: false; err: string } => {
        const nums: number[] = [];
        for (const a of list) {
          if (a.isRange) {
            for (const v of a.vals) {
              if (v.kind === 'err') return { ok: false, err: v.code };
              if (v.kind === 'num') nums.push(v.v);
            }
          } else if (a.isRef) {
            if (a.val.kind === 'err') return { ok: false, err: a.val.code };
            if (a.val.kind === 'num') nums.push(a.val.v);
          } else {
            if (a.val.kind === 'err') return { ok: false, err: a.val.code };
            if (a.val.kind === 'empty') continue;
            const n = toNumber(a.val);
            if (!n.ok) return { ok: false, err: n.err };
            nums.push(n.n);
          }
        }
        return { ok: true, nums };
      };

      const scalar = (idx: number): EvalVal => {
        const a = args[idx];
        if (!a) return { kind: 'empty' };
        return a.isRange ? (a.vals[0] ?? { kind: 'empty' }) : a.val;
      };

      const asRangeVals = (idx: number): EvalVal[] => {
        const a = args[idx];
        if (!a) return [];
        return a.isRange ? a.vals : [a.val];
      };

      const argShape = (a: ArgVal): { rows: number; cols: number } =>
        a.isRange ? { rows: a.rows, cols: a.cols } : { rows: 1, cols: 1 };

      // Helper to align multiple criteria/sum ranges when all are whole-column ranges of same column count
      const alignMultiRanges = (rangeArgs: ArgVal[]): { ok: true; aligned: EvalVal[][]; len: number } | { ok: false; err: string } => {
        if (rangeArgs.length === 0) return { ok: true, aligned: [], len: 0 };
        const allWholeCol =
          rangeArgs.every((a) => a.isRange && a.isWholeCol) &&
          new Set(rangeArgs.map((a) => (a.isRange ? a.cols : 1))).size === 1;
        if (allWholeCol) {
          const cols = rangeArgs[0].isRange ? rangeArgs[0].cols : 1;
          const maxRows = Math.max(...rangeArgs.map((a) => (a.isRange ? a.rows : 1)));
          const len = maxRows * cols;
          const aligned = rangeArgs.map((a) => {
            const vals = a.isRange ? a.vals : [a.val];
            if (vals.length === len) return vals;
            const padded = vals.slice();
            while (padded.length < len) padded.push({ kind: 'empty' });
            return padded;
          });
          return { ok: true, aligned, len };
        }
        const shapes = rangeArgs.map(argShape);
        const r0 = shapes[0].rows;
        const c0 = shapes[0].cols;
        if (shapes.some((s) => s.rows !== r0 || s.cols !== c0)) {
          return { ok: false, err: '#VALUE!' };
        }
        const aligned = rangeArgs.map((a) => (a.isRange ? a.vals : [a.val]));
        return { ok: true, aligned, len: r0 * c0 };
      };

      switch (name) {
        case 'SUM': {
          const r = flattenNums(args);
          if (!r.ok) return { kind: 'err', code: r.err };
          const sum = r.nums.reduce((acc, x) => acc + x, 0);
          return { kind: 'num', v: Number(sum.toPrecision(15)) };
        }
        case 'SUMPRODUCT': {
          if (args.length === 0) return { kind: 'err', code: '#VALUE!' };
          const shapes = args.map(argShape);
          const r0 = shapes[0].rows;
          const c0 = shapes[0].cols;
          // Excel requires all SUMPRODUCT arrays to have identical rows & cols dimensions
          if (shapes.some((s) => s.rows !== r0 || s.cols !== c0)) {
            return { kind: 'err', code: '#VALUE!' };
          }
          const arrays = args.map((_, idx) => asRangeVals(idx));
          const len = r0 * c0;
          let total = 0;
          for (let i = 0; i < len; i++) {
            let prod = 1;
            for (const arr of arrays) {
              const v = arr[i];
              if (v.kind === 'err') return v;
              prod *= v.kind === 'num' ? v.v : 0;
            }
            total += prod;
          }
          return { kind: 'num', v: Number(total.toPrecision(15)) };
        }
        case 'SUBTOTAL': {
          if (args.length < 2) return { kind: 'err', code: '#VALUE!' };
          const fnRes = toNumber(scalar(0));
          if (!fnRes.ok) return { kind: 'err', code: fnRes.err };
          const rawCode = Math.trunc(fnRes.n);
          const ignoreHidden = rawCode >= 100;
          const code = rawCode % 100;
          const fnMap: Record<number, string> = {
            1: 'AVERAGE',
            2: 'COUNT',
            3: 'COUNTA',
            4: 'MAX',
            5: 'MIN',
            6: 'PRODUCT',
            9: 'SUM',
          };
          const mapped = fnMap[code];
          if (!mapped) throw new UnsupportedFormulaError(`不支持的 SUBTOTAL 功能号：${fnRes.n}`);

          // Filter out nested SUBTOTAL(...) cells and (for 101-111) hidden rows
          const filteredVals: EvalVal[] = [];
          for (const a of args.slice(1)) {
            if (a.isRange) {
              const ws = wb.Sheets[a.origin.sheet];
              const rowsMeta = (ws?.['!rows'] || []) as Array<{ hidden?: boolean } | undefined>;
              for (let dr = 0; dr < a.rows; dr++) {
                const rr = a.origin.r + dr;
                if (ignoreHidden && rowsMeta[rr]?.hidden) continue;
                for (let dc = 0; dc < a.cols; dc++) {
                  const cc = a.origin.c + dc;
                  const cellObj = ws?.[XLSX.utils.encode_cell({ r: rr, c: cc })] as XLSXModule.CellObject | undefined;
                  if (cellObj?.f && /\bSUBTOTAL\s*\(/i.test(String(cellObj.f))) continue;
                  filteredVals.push(a.vals[dr * a.cols + dc]);
                }
              }
            } else if (a.isRef && a.origin) {
              const ws = wb.Sheets[a.origin.sheet];
              const rowsMeta = (ws?.['!rows'] || []) as Array<{ hidden?: boolean } | undefined>;
              if (ignoreHidden && rowsMeta[a.origin.r]?.hidden) continue;
              const cellObj = ws?.[XLSX.utils.encode_cell({ r: a.origin.r, c: a.origin.c })] as XLSXModule.CellObject | undefined;
              if (cellObj?.f && /\bSUBTOTAL\s*\(/i.test(String(cellObj.f))) continue;
              filteredVals.push(a.val);
            } else {
              filteredVals.push(a.val);
            }
          }
          const syntheticRange: ArgVal = {
            isRange: true,
            origin: { sheet: currentSheet, r: 0, c: 0 },
            vals: filteredVals,
            rows: filteredVals.length,
            cols: 1,
          };
          return evalFunction(mapped, [syntheticRange]);
        }
        case 'TODAY': {
          const now = new Date();
          const utcToday = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
          const excelEpoch = Date.UTC(1899, 11, 30);
          return { kind: 'num', v: Math.floor((utcToday - excelEpoch) / 86400000) };
        }
        case 'NOW': {
          const now = new Date();
          const utcNow = Date.UTC(
            now.getFullYear(),
            now.getMonth(),
            now.getDate(),
            now.getHours(),
            now.getMinutes(),
            now.getSeconds()
          );
          const excelEpoch = Date.UTC(1899, 11, 30);
          return { kind: 'num', v: (utcNow - excelEpoch) / 86400000 };
        }
        case 'WEEKDAY': {
          if (args.length < 1) return { kind: 'err', code: '#VALUE!' };
          const sRes = toNumber(scalar(0));
          if (!sRes.ok) return { kind: 'err', code: sRes.err };
          const rtRes = args.length >= 2 ? toNumber(scalar(1)) : { ok: true as const, n: 1 };
          if (!rtRes.ok) return { kind: 'err', code: rtRes.err };
          const wd = excelWeekday(sRes.n, Math.trunc(rtRes.n));
          return wd === null ? { kind: 'err', code: '#NUM!' } : { kind: 'num', v: wd };
        }
        case 'CEILING': {
          if (args.length < 2) return { kind: 'err', code: '#VALUE!' };
          const nRes = toNumber(scalar(0));
          const sRes = toNumber(scalar(1));
          if (!nRes.ok) return { kind: 'err', code: nRes.err };
          if (!sRes.ok) return { kind: 'err', code: sRes.err };
          const num = nRes.n;
          const sig = sRes.n;
          if (sig === 0) return { kind: 'num', v: 0 };
          if (num > 0 && sig < 0) return { kind: 'err', code: '#NUM!' };
          const q = num / sig;
          const res = (sig > 0 ? Math.ceil(q - 1e-12) : Math.ceil(q - 1e-12)) * sig;
          return { kind: 'num', v: Number(res.toPrecision(15)) };
        }
        case 'FLOOR': {
          if (args.length < 2) return { kind: 'err', code: '#VALUE!' };
          const nRes = toNumber(scalar(0));
          const sRes = toNumber(scalar(1));
          if (!nRes.ok) return { kind: 'err', code: nRes.err };
          if (!sRes.ok) return { kind: 'err', code: sRes.err };
          const num = nRes.n;
          const sig = sRes.n;
          if (sig === 0) return { kind: 'err', code: '#DIV/0!' };
          if (num > 0 && sig < 0) return { kind: 'err', code: '#NUM!' };
          const q = num / sig;
          const res = Math.floor(q + 1e-12) * sig;
          return { kind: 'num', v: Number(res.toPrecision(15)) };
        }
        case 'AVERAGE':
        case 'AVG': {
          const r = flattenNums(args);
          if (!r.ok) return { kind: 'err', code: r.err };
          if (r.nums.length === 0) return { kind: 'err', code: '#DIV/0!' };
          const avg = r.nums.reduce((acc, x) => acc + x, 0) / r.nums.length;
          return { kind: 'num', v: Number(avg.toPrecision(15)) };
        }
        case 'MEDIAN': {
          const r = flattenNums(args);
          if (!r.ok) return { kind: 'err', code: r.err };
          if (r.nums.length === 0) return { kind: 'err', code: '#NUM!' };
          const sorted = [...r.nums].sort((a, b) => a - b);
          const mid = Math.floor(sorted.length / 2);
          const med = sorted.length % 2 === 1 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
          return { kind: 'num', v: Number(med.toPrecision(15)) };
        }
        case 'LARGE':
        case 'SMALL': {
          if (args.length < 2) return { kind: 'err', code: '#VALUE!' };
          const r = flattenNums([args[0]]);
          if (!r.ok) return { kind: 'err', code: r.err };
          const kVal = toNumber(scalar(1));
          if (!kVal.ok) return { kind: 'err', code: kVal.err };
          const k = Math.trunc(kVal.n);
          if (k < 1 || k > r.nums.length) return { kind: 'err', code: '#NUM!' };
          const sorted = [...r.nums].sort((a, b) => (name === 'LARGE' ? b - a : a - b));
          return { kind: 'num', v: sorted[k - 1] };
        }
        case 'RANK':
        case 'RANK.EQ': {
          if (args.length < 2) return { kind: 'err', code: '#VALUE!' };
          const target = toNumber(scalar(0));
          if (!target.ok) return { kind: 'err', code: target.err };
          const refVals = asRangeVals(1);
          const nums: number[] = [];
          for (const v of refVals) {
            if (v.kind === 'err') return v;
            if (v.kind === 'num') nums.push(v.v);
          }
          const orderVal = args.length >= 3 ? toNumber(scalar(2)) : { ok: true as const, n: 0 };
          if (!orderVal.ok) return { kind: 'err', code: orderVal.err };
          const ascending = orderVal.n !== 0;
          if (!nums.some((x) => x === target.n)) return { kind: 'err', code: '#N/A' };
          let rank = 1;
          for (const x of nums) {
            if (ascending ? x < target.n : x > target.n) rank++;
          }
          return { kind: 'num', v: rank };
        }
        case 'COUNTIF': {
          if (args.length < 2) return { kind: 'err', code: '#VALUE!' };
          const vals = asRangeVals(0);
          const crit = scalar(1);
          if (isUnsupportedErr(crit)) return crit;
          let cnt = 0;
          for (const v of vals) {
            if (isUnsupportedErr(v)) return v;
            if (matchesCriteria(v, crit)) cnt++;
          }
          return { kind: 'num', v: cnt };
        }
        case 'COUNTIFS': {
          if (args.length < 2 || args.length % 2 !== 0) return { kind: 'err', code: '#VALUE!' };
          const rangeArgs: ArgVal[] = [];
          const crits: EvalVal[] = [];
          for (let i = 0; i < args.length; i += 2) {
            rangeArgs.push(args[i]);
            const cr = scalar(i + 1);
            if (isUnsupportedErr(cr)) return cr;
            crits.push(cr);
          }
          const al = alignMultiRanges(rangeArgs);
          if (!al.ok) return { kind: 'err', code: al.err };
          let cnt = 0;
          for (let idx = 0; idx < al.len; idx++) {
            let ok = true;
            for (let p = 0; p < al.aligned.length; p++) {
              const v = al.aligned[p][idx];
              if (isUnsupportedErr(v)) return v;
              if (!matchesCriteria(v, crits[p])) {
                ok = false;
                break;
              }
            }
            if (ok) cnt++;
          }
          return { kind: 'num', v: cnt };
        }
        case 'SUMIF': {
          if (args.length < 2) return { kind: 'err', code: '#VALUE!' };
          const critArg = args[0];
          const crit = scalar(1);
          if (isUnsupportedErr(crit)) return crit;

          let critVals = asRangeVals(0);
          let sumVals: EvalVal[];

          if (args.length >= 3) {
            const sumArg = args[2];
            if (critArg.isRange && critArg.isWholeCol && sumArg.isRange && sumArg.isWholeCol && critArg.cols === sumArg.cols) {
              // Both are whole-column references (possibly cross-sheet with different used end rows) -> align from row 0 to maxRows
              const maxRows = Math.max(critArg.rows, sumArg.rows);
              const len = maxRows * critArg.cols;
              critVals = critArg.vals.slice();
              while (critVals.length < len) critVals.push({ kind: 'empty' });
              sumVals = sumArg.vals.slice();
              while (sumVals.length < len) sumVals.push({ kind: 'empty' });
            } else if (sumArg.origin) {
              // Excel SUMIF rule: sum_range is anchored at its top-left cell and auto-resized to match criteria_range dimensions
              const critRows = critArg.isRange ? critArg.rows : 1;
              const critCols = critArg.isRange ? critArg.cols : 1;
              const expanded = evalRange(sumArg.origin, {
                sheet: sumArg.origin.sheet,
                r: sumArg.origin.r + critRows - 1,
                c: sumArg.origin.c + critCols - 1,
              });
              sumVals = expanded.vals;
            } else {
              sumVals = asRangeVals(2);
            }
          } else {
            sumVals = critVals;
          }

          let sum = 0;
          const len = Math.min(critVals.length, sumVals.length);
          for (let i = 0; i < len; i++) {
            const cv = critVals[i];
            if (isUnsupportedErr(cv)) return cv;
            if (matchesCriteria(cv, crit)) {
              const sv = sumVals[i];
              if (sv.kind === 'err') return sv;
              if (sv.kind === 'num') sum += sv.v;
            }
          }
          return { kind: 'num', v: Number(sum.toPrecision(15)) };
        }
        case 'SUMIFS': {
          if (args.length < 3 || args.length % 2 !== 1) return { kind: 'err', code: '#VALUE!' };
          const allRangeArgs: ArgVal[] = [args[0]];
          const crits: EvalVal[] = [];
          for (let i = 1; i < args.length; i += 2) {
            allRangeArgs.push(args[i]);
            const cr = scalar(i + 1);
            if (isUnsupportedErr(cr)) return cr;
            crits.push(cr);
          }
          const al = alignMultiRanges(allRangeArgs);
          if (!al.ok) return { kind: 'err', code: al.err };
          const sumVals = al.aligned[0];
          const critRanges = al.aligned.slice(1);
          let sum = 0;
          for (let idx = 0; idx < al.len; idx++) {
            let ok = true;
            for (let p = 0; p < critRanges.length; p++) {
              const cv = critRanges[p][idx];
              if (isUnsupportedErr(cv)) return cv;
              if (!matchesCriteria(cv, crits[p])) {
                ok = false;
                break;
              }
            }
            if (ok) {
              const sv = sumVals[idx];
              if (sv.kind === 'err') return sv;
              if (sv.kind === 'num') sum += sv.v;
            }
          }
          return { kind: 'num', v: Number(sum.toPrecision(15)) };
        }
        case 'VLOOKUP': {
          if (args.length < 3) return { kind: 'err', code: '#VALUE!' };
          const lookupVal = scalar(0);
          if (lookupVal.kind === 'err') return lookupVal;
          const tableArg = args[1];
          if (!tableArg || !tableArg.isRange) return { kind: 'err', code: '#VALUE!' };
          const colIdxRes = toNumber(scalar(2));
          if (!colIdxRes.ok) return { kind: 'err', code: colIdxRes.err };
          const colIdx = Math.trunc(colIdxRes.n);
          if (colIdx < 1) return { kind: 'err', code: '#VALUE!' };
          if (colIdx > tableArg.cols) return { kind: 'err', code: '#REF!' };
          const approxRes = args.length >= 4 ? toBool(scalar(3)) : { ok: true as const, b: true };
          if (!approxRes.ok) return { kind: 'err', code: approxRes.err };

          if (!approxRes.b) {
            for (let r = 0; r < tableArg.rows; r++) {
              const keyCell = tableArg.vals[r * tableArg.cols];
              if (isUnsupportedErr(keyCell)) return keyCell;
              if (keyCell.kind === 'err') continue;
              if (lookupVal.kind === 'str' && keyCell.kind === 'str') {
                if (wildcardToRegex(lookupVal.v).test(keyCell.v)) {
                  return tableArg.vals[r * tableArg.cols + (colIdx - 1)];
                }
              } else if (compareEvalVals(keyCell, lookupVal) === 0 && keyCell.kind === lookupVal.kind) {
                return tableArg.vals[r * tableArg.cols + (colIdx - 1)];
              }
            }
            return { kind: 'err', code: '#N/A' };
          }
          let bestRow = -1;
          for (let r = 0; r < tableArg.rows; r++) {
            const keyCell = tableArg.vals[r * tableArg.cols];
            if (isUnsupportedErr(keyCell)) return keyCell;
            if (keyCell.kind !== lookupVal.kind) continue;
            if (compareEvalVals(keyCell, lookupVal) <= 0) bestRow = r;
          }
          if (bestRow < 0) return { kind: 'err', code: '#N/A' };
          return tableArg.vals[bestRow * tableArg.cols + (colIdx - 1)];
        }
        case 'MATCH': {
          if (args.length < 2) return { kind: 'err', code: '#VALUE!' };
          const lookupVal = scalar(0);
          if (lookupVal.kind === 'err') return lookupVal;
          const arr = asRangeVals(1);
          const typeRes = args.length >= 3 ? toNumber(scalar(2)) : { ok: true as const, n: 1 };
          if (!typeRes.ok) return { kind: 'err', code: typeRes.err };
          const matchType = Math.trunc(typeRes.n);
          if (matchType === 0) {
            for (let i = 0; i < arr.length; i++) {
              const cellVal = arr[i];
              if (isUnsupportedErr(cellVal)) return cellVal;
              if (cellVal.kind === 'err') continue;
              if (lookupVal.kind === 'str' && cellVal.kind === 'str') {
                if (wildcardToRegex(lookupVal.v).test(cellVal.v)) return { kind: 'num', v: i + 1 };
              } else if (cellVal.kind === lookupVal.kind && compareEvalVals(cellVal, lookupVal) === 0) {
                return { kind: 'num', v: i + 1 };
              }
            }
            return { kind: 'err', code: '#N/A' };
          }
          let bestIdx = -1;
          for (let i = 0; i < arr.length; i++) {
            const cellVal = arr[i];
            if (isUnsupportedErr(cellVal)) return cellVal;
            if (cellVal.kind !== lookupVal.kind) continue;
            const cmp = compareEvalVals(cellVal, lookupVal);
            if (matchType > 0 ? cmp <= 0 : cmp >= 0) bestIdx = i;
          }
          return bestIdx >= 0 ? { kind: 'num', v: bestIdx + 1 } : { kind: 'err', code: '#N/A' };
        }
        case 'INDEX': {
          if (args.length < 2) return { kind: 'err', code: '#VALUE!' };
          const arrArg = args[0];
          const rowRes = toNumber(scalar(1));
          if (!rowRes.ok) return { kind: 'err', code: rowRes.err };
          const colRes = args.length >= 3 ? toNumber(scalar(2)) : { ok: true as const, n: 1 };
          if (!colRes.ok) return { kind: 'err', code: colRes.err };
          const rIdx = Math.trunc(rowRes.n);
          const cIdx = Math.trunc(colRes.n);
          if (!arrArg.isRange) {
            return rIdx === 1 && cIdx === 1 ? arrArg.val : { kind: 'err', code: '#REF!' };
          }
          if (arrArg.rows === 1 && args.length < 3) {
            if (rIdx < 1 || rIdx > arrArg.cols) return { kind: 'err', code: '#REF!' };
            return arrArg.vals[rIdx - 1];
          }
          if (rIdx < 1 || rIdx > arrArg.rows || cIdx < 1 || cIdx > arrArg.cols) {
            return { kind: 'err', code: '#REF!' };
          }
          return arrArg.vals[(rIdx - 1) * arrArg.cols + (cIdx - 1)];
        }
        case 'MIN': {
          const r = flattenNums(args);
          if (!r.ok) return { kind: 'err', code: r.err };
          return { kind: 'num', v: r.nums.length === 0 ? 0 : Math.min(...r.nums) };
        }
        case 'MAX': {
          const r = flattenNums(args);
          if (!r.ok) return { kind: 'err', code: r.err };
          return { kind: 'num', v: r.nums.length === 0 ? 0 : Math.max(...r.nums) };
        }
        case 'COUNT': {
          const r = flattenNums(args);
          if (!r.ok) return { kind: 'err', code: r.err };
          return { kind: 'num', v: r.nums.length };
        }
        case 'COUNTA': {
          let cnt = 0;
          for (const a of args) {
            const list = a.isRange ? a.vals : [a.val];
            for (const v of list) {
              if (isUnsupportedErr(v)) return v;
              if (v.kind !== 'empty' && !(v.kind === 'str' && v.v === '')) cnt++;
            }
          }
          return { kind: 'num', v: cnt };
        }
        case 'PRODUCT': {
          const r = flattenNums(args);
          if (!r.ok) return { kind: 'err', code: r.err };
          if (r.nums.length === 0) return { kind: 'num', v: 0 };
          return { kind: 'num', v: Number(r.nums.reduce((a, b) => a * b, 1).toPrecision(15)) };
        }
        case 'ABS': {
          const n = toNumber(scalar(0));
          return n.ok ? { kind: 'num', v: Math.abs(n.n) } : { kind: 'err', code: n.err };
        }
        case 'INT': {
          const n = toNumber(scalar(0));
          return n.ok ? { kind: 'num', v: Math.floor(n.n) } : { kind: 'err', code: n.err };
        }
        case 'SQRT': {
          const n = toNumber(scalar(0));
          if (!n.ok) return { kind: 'err', code: n.err };
          return n.n < 0 ? { kind: 'err', code: '#NUM!' } : { kind: 'num', v: Math.sqrt(n.n) };
        }
        case 'POWER': {
          const a = toNumber(scalar(0));
          const b = toNumber(scalar(1));
          if (!a.ok) return { kind: 'err', code: a.err };
          if (!b.ok) return { kind: 'err', code: b.err };
          const p = Math.pow(a.n, b.n);
          return Number.isFinite(p) ? { kind: 'num', v: p } : { kind: 'err', code: '#NUM!' };
        }
        case 'MOD': {
          const a = toNumber(scalar(0));
          const b = toNumber(scalar(1));
          if (!a.ok) return { kind: 'err', code: a.err };
          if (!b.ok) return { kind: 'err', code: b.err };
          if (b.n === 0) return { kind: 'err', code: '#DIV/0!' };
          return { kind: 'num', v: a.n - b.n * Math.floor(a.n / b.n) };
        }
        case 'ROUND':
        case 'ROUNDUP':
        case 'ROUNDDOWN': {
          const a = toNumber(scalar(0));
          const d = toNumber(scalar(1));
          if (!a.ok) return { kind: 'err', code: a.err };
          if (!d.ok) return { kind: 'err', code: d.err };
          const mode = name === 'ROUNDUP' ? 'up' : name === 'ROUNDDOWN' ? 'down' : 'round';
          return { kind: 'num', v: roundAwayFromZero(a.n, Math.trunc(d.n), mode) };
        }
        case 'IF': {
          const cond = toBool(scalar(0));
          if (!cond.ok) return { kind: 'err', code: cond.err };
          if (cond.b) return args.length >= 2 ? scalar(1) : { kind: 'bool', v: true };
          return args.length >= 3 ? scalar(2) : { kind: 'bool', v: false };
        }
        case 'IFERROR': {
          const first = scalar(0);
          // NEVER swallow #UNSUPPORTED! (evaluator limitation must propagate loudly, unlike real Excel errors like #DIV/0!)
          if (isUnsupportedErr(first)) return first;
          return first.kind === 'err' ? scalar(1) : first;
        }
        case 'IFNA': {
          const first = scalar(0);
          if (isUnsupportedErr(first)) return first;
          return first.kind === 'err' && first.code === '#N/A' ? scalar(1) : first;
        }
        case 'ISBLANK': {
          const v = scalar(0);
          if (isUnsupportedErr(v)) return v;
          return { kind: 'bool', v: v.kind === 'empty' };
        }
        case 'ISNUMBER': {
          const v = scalar(0);
          if (isUnsupportedErr(v)) return v;
          return { kind: 'bool', v: v.kind === 'num' };
        }
        case 'ISTEXT': {
          const v = scalar(0);
          if (isUnsupportedErr(v)) return v;
          return { kind: 'bool', v: v.kind === 'str' };
        }
        case 'ISERROR': {
          const v = scalar(0);
          if (isUnsupportedErr(v)) return v;
          return { kind: 'bool', v: v.kind === 'err' };
        }
        case 'ISNA': {
          const v = scalar(0);
          if (isUnsupportedErr(v)) return v;
          return { kind: 'bool', v: v.kind === 'err' && v.code === '#N/A' };
        }
        case 'AND': {
          let result = true;
          for (const a of args) {
            const list = a.isRange ? a.vals : [a.val];
            for (const v of list) {
              if (v.kind === 'empty') continue;
              const b = toBool(v);
              if (!b.ok) return { kind: 'err', code: b.err };
              if (!b.b) result = false;
            }
          }
          return { kind: 'bool', v: result };
        }
        case 'OR': {
          let result = false;
          for (const a of args) {
            const list = a.isRange ? a.vals : [a.val];
            for (const v of list) {
              if (v.kind === 'empty') continue;
              const b = toBool(v);
              if (!b.ok) return { kind: 'err', code: b.err };
              if (b.b) result = true;
            }
          }
          return { kind: 'bool', v: result };
        }
        case 'NOT': {
          const b = toBool(scalar(0));
          return b.ok ? { kind: 'bool', v: !b.b } : { kind: 'err', code: b.err };
        }
        case 'TRIM': {
          const v = scalar(0);
          if (v.kind === 'err') return v;
          return { kind: 'str', v: toStr(v).trim().replace(/\s+/g, ' ') };
        }
        case 'LEN': {
          const v = scalar(0);
          if (v.kind === 'err') return v;
          return { kind: 'num', v: toStr(v).length };
        }
        case 'LEFT': {
          const v = scalar(0);
          if (v.kind === 'err') return v;
          const cnt = args.length >= 2 ? toNumber(scalar(1)) : { ok: true as const, n: 1 };
          if (!cnt.ok) return { kind: 'err', code: cnt.err };
          return { kind: 'str', v: toStr(v).slice(0, Math.max(0, Math.trunc(cnt.n))) };
        }
        case 'RIGHT': {
          const v = scalar(0);
          if (v.kind === 'err') return v;
          const cnt = args.length >= 2 ? toNumber(scalar(1)) : { ok: true as const, n: 1 };
          if (!cnt.ok) return { kind: 'err', code: cnt.err };
          const s = toStr(v);
          const n = Math.max(0, Math.trunc(cnt.n));
          return { kind: 'str', v: n === 0 ? '' : s.slice(-n) };
        }
        case 'MID': {
          const v = scalar(0);
          if (v.kind === 'err') return v;
          const st = toNumber(scalar(1));
          const len = toNumber(scalar(2));
          if (!st.ok) return { kind: 'err', code: st.err };
          if (!len.ok) return { kind: 'err', code: len.err };
          return { kind: 'str', v: toStr(v).slice(Math.max(0, Math.trunc(st.n) - 1), Math.max(0, Math.trunc(st.n) - 1 + Math.trunc(len.n))) };
        }
        case 'SUBSTITUTE': {
          if (args.length < 3) return { kind: 'err', code: '#VALUE!' };
          const textVal = scalar(0);
          const oldVal = scalar(1);
          const newVal = scalar(2);
          if (textVal.kind === 'err') return textVal;
          if (oldVal.kind === 'err') return oldVal;
          if (newVal.kind === 'err') return newVal;
          const text = toStr(textVal);
          const oldText = toStr(oldVal);
          const newText = toStr(newVal);
          if (oldText === '') return { kind: 'str', v: text };
          if (args.length >= 4) {
            const instRes = toNumber(scalar(3));
            if (!instRes.ok) return { kind: 'err', code: instRes.err };
            const inst = Math.trunc(instRes.n);
            if (inst < 1) return { kind: 'err', code: '#VALUE!' };
            let count = 0;
            let idx = 0;
            while (true) {
              const found = text.indexOf(oldText, idx);
              if (found < 0) return { kind: 'str', v: text };
              count++;
              if (count === inst) {
                return { kind: 'str', v: text.slice(0, found) + newText + text.slice(found + oldText.length) };
              }
              idx = found + oldText.length;
            }
          }
          return { kind: 'str', v: text.split(oldText).join(newText) };
        }
        case 'REPLACE': {
          if (args.length < 4) return { kind: 'err', code: '#VALUE!' };
          const oldVal = scalar(0);
          const newVal = scalar(3);
          if (oldVal.kind === 'err') return oldVal;
          if (newVal.kind === 'err') return newVal;
          const st = toNumber(scalar(1));
          const cnt = toNumber(scalar(2));
          if (!st.ok) return { kind: 'err', code: st.err };
          if (!cnt.ok) return { kind: 'err', code: cnt.err };
          const s = toStr(oldVal);
          const startIdx = Math.trunc(st.n) - 1;
          const numChars = Math.trunc(cnt.n);
          if (startIdx < 0 || numChars < 0) return { kind: 'err', code: '#VALUE!' };
          return { kind: 'str', v: s.slice(0, startIdx) + toStr(newVal) + s.slice(startIdx + numChars) };
        }
        case 'FIND': {
          if (args.length < 2) return { kind: 'err', code: '#VALUE!' };
          const needle = scalar(0);
          const hay = scalar(1);
          if (needle.kind === 'err') return needle;
          if (hay.kind === 'err') return hay;
          const startRes = args.length >= 3 ? toNumber(scalar(2)) : { ok: true as const, n: 1 };
          if (!startRes.ok) return { kind: 'err', code: startRes.err };
          const startIdx = Math.trunc(startRes.n) - 1;
          if (startIdx < 0 || startIdx > toStr(hay).length) return { kind: 'err', code: '#VALUE!' };
          const found = toStr(hay).indexOf(toStr(needle), startIdx);
          return found < 0 ? { kind: 'err', code: '#VALUE!' } : { kind: 'num', v: found + 1 };
        }
        case 'SEARCH': {
          if (args.length < 2) return { kind: 'err', code: '#VALUE!' };
          const needle = scalar(0);
          const hay = scalar(1);
          if (needle.kind === 'err') return needle;
          if (hay.kind === 'err') return hay;
          const startRes = args.length >= 3 ? toNumber(scalar(2)) : { ok: true as const, n: 1 };
          if (!startRes.ok) return { kind: 'err', code: startRes.err };
          const startIdx = Math.trunc(startRes.n) - 1;
          const hayStr = toStr(hay);
          const needleStr = toStr(needle);
          if (startIdx < 0 || startIdx > hayStr.length) return { kind: 'err', code: '#VALUE!' };
          if (/[*?~]/.test(needleStr)) {
            const sub = hayStr.slice(startIdx);
            // Convert wildcard to unanchored regex
            let reSrc = '';
            for (let i = 0; i < needleStr.length; i++) {
              const ch = needleStr[i];
              if (ch === '~' && i + 1 < needleStr.length) {
                reSrc += needleStr[++i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
              } else if (ch === '*') {
                reSrc += '[\\s\\S]*?';
              } else if (ch === '?') {
                reSrc += '[\\s\\S]';
              } else {
                reSrc += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
              }
            }
            const m = new RegExp(reSrc, 'i').exec(sub);
            return m ? { kind: 'num', v: startIdx + m.index + 1 } : { kind: 'err', code: '#VALUE!' };
          }
          const found = hayStr.toLowerCase().indexOf(needleStr.toLowerCase(), startIdx);
          return found < 0 ? { kind: 'err', code: '#VALUE!' } : { kind: 'num', v: found + 1 };
        }
        case 'CONCAT':
        case 'CONCATENATE': {
          let out = '';
          for (const a of args) {
            const list = a.isRange ? a.vals : [a.val];
            for (const v of list) {
              if (v.kind === 'err') return v;
              out += toStr(v);
            }
          }
          return { kind: 'str', v: out };
        }
        case 'TEXT': {
          const v = scalar(0);
          const fmtVal = scalar(1);
          if (v.kind === 'err') return v;
          if (fmtVal.kind === 'err') return fmtVal;
          const fmt = toStr(fmtVal);
          if (v.kind === 'num') {
            try {
              return { kind: 'str', v: XLSX.SSF.format(fmt, v.v) };
            } catch {
              return { kind: 'str', v: String(v.v) };
            }
          }
          return { kind: 'str', v: toStr(v) };
        }
        case 'UPPER': {
          const v = scalar(0);
          return v.kind === 'err' ? v : { kind: 'str', v: toStr(v).toUpperCase() };
        }
        case 'LOWER': {
          const v = scalar(0);
          return v.kind === 'err' ? v : { kind: 'str', v: toStr(v).toLowerCase() };
        }
        case 'VALUE': {
          const n = toNumber(scalar(0));
          return n.ok ? { kind: 'num', v: n.n } : { kind: 'err', code: n.err };
        }
        default:
          throw new UnsupportedFormulaError(`不支持的函数：${rawName}`);
      }
    };

    const parsePrimary = (): EvalVal => {
      const tok = peek();
      if (tok.type === 'num') {
        consume();
        return { kind: 'num', v: Number(tok.value) };
      }
      if (tok.type === 'str') {
        consume();
        return { kind: 'str', v: tok.value };
      }
      if (tok.type === 'lparen') {
        consume();
        const val = parseComparison();
        if (peek().type !== 'rparen') throw new UnsupportedFormulaError('公式括号不匹配');
        consume();
        return val;
      }
      if (tok.type === 'ident') {
        consume();
        const upper = tok.value.toUpperCase();
        if (upper === 'TRUE') return { kind: 'bool', v: true };
        if (upper === 'FALSE') return { kind: 'bool', v: false };
        if (tok.value.startsWith('#')) return { kind: 'err', code: upper };

        // Lazy IF(cond, trueBranch, falseBranch) — only evaluates the selected branch!
        if (upper === 'IF' && peek().type === 'lparen') {
          consume();
          const condVal = parseComparison();
          const condBool = toBool(condVal);
          if (!condBool.ok) {
            while (peek().type === 'comma') {
              consume();
              skipBalancedArg();
            }
            if (peek().type !== 'rparen') throw new UnsupportedFormulaError('函数 IF 括号不匹配');
            consume();
            return { kind: 'err', code: condBool.err };
          }
          let result: EvalVal = { kind: 'bool', v: condBool.b };
          if (peek().type === 'comma') {
            consume();
            if (condBool.b) {
              result = parseComparison();
              if (peek().type === 'comma') {
                consume();
                skipBalancedArg();
              }
            } else {
              skipBalancedArg();
              if (peek().type === 'comma') {
                consume();
                result = parseComparison();
              }
            }
          }
          if (peek().type !== 'rparen') throw new UnsupportedFormulaError('函数 IF 括号不匹配');
          consume();
          return result;
        }

        // Function call
        if (peek().type === 'lparen') {
          consume();
          const args: ArgVal[] = [];
          if (peek().type !== 'rparen') {
            args.push(parseArg());
            while (peek().type === 'comma') {
              consume();
              args.push(parseArg());
            }
          }
          if (peek().type !== 'rparen') throw new UnsupportedFormulaError(`函数 ${tok.value} 括号不匹配`);
          consume();
          return evalFunction(tok.value, args);
        }

        // Cell reference
        const ref = parseCellRef(tok.value, currentSheet);
        if (ref) {
          return evalCell(ref.sheet, ref.r, ref.c);
        }
        throw new UnsupportedFormulaError(`无法识别的标识符：${tok.value}`);
      }
      throw new UnsupportedFormulaError(`公式语法无法解析：${tok.value || tok.type}`);
    };

    const result = parseComparison();
    if (peek().type !== 'eof') {
      throw new UnsupportedFormulaError(`公式存在未解析尾部：${peek().value}`);
    }
    return result;
  };

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws || !ws['!ref']) continue;
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr] as XLSXModule.CellObject | undefined;
        if (cell && cell.f) {
          evalCell(sheetName, r, c);
        }
      }
    }
  }
}
