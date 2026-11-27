import {
  Dec,
  buildStyledExcelBuffer,
  colsFromList,
  numericText,
  readWorkbookAllSheets,
} from './toolboxWorkbook.js';

export const PRICE_SCORE_HEADERS = [
  '分标名称',
  '分包名称',
  '轮次',
  '投标人名称',
  '投标价格（万元）',
  '折扣率',
  '是否参与计算',
];

export const PRICE_RULE_HEADERS = [
  '分标名称',
  '价格公式',
  'qz',
  'w1',
  'w2',
  'n1',
  'n2',
  'c',
  'F1',
  'F2',
  'a',
  'b',
  'm',
];

export const PRICE_METHODS = [
  '区间平均价浮动法',
  '算术平均值下浮法',
  '区间平均下浮双边曲线算法',
  '区间复合平均价法',
  '算数平均值浮动法',
  '区间复合平均价浮动法',
] as const;

interface Rule {
  method: number; // 1..6
  values: Record<string, Dec>;
}

interface Bid {
  section: string;
  pack: string;
  round: string;
  supplier: string;
  value: Dec;
  discount: boolean;
  source: string;
}

export interface PriceScoreRow {
  section: string;
  pack: string;
  round: string;
  supplier: string;
  quote_basis: string;
  quote_value: string;
  method_name: string;
  base_value: string;
  price_score: string;
  weighted_score: string;
  source: string;
}

export interface PriceScoreResult {
  quoteFiles: string[];
  ruleFile: string;
  roundFilter: string;
  totalRows: number;
  packageCount: number;
  rows: PriceScoreRow[];
  warnings: string[];
}

function parameterNames(method: number): string[] {
  switch (method) {
    case 1:
      return ['w1', 'w2', 'n1', 'n2', 'c', 'qz'];
    case 2:
      return ['c', 'n1', 'n2', 'qz'];
    case 3:
      return ['F1', 'F2', 'n1', 'n2', 'c', 'qz'];
    case 4:
      return ['b', 'c', 'n1', 'n2', 'qz'];
    case 5:
      return ['F1', 'F2', 'm', 'c', 'qz'];
    default:
      return ['a', 'b', 'n1', 'n2', 'c', 'qz'];
  }
}

export function priceRuleTemplateHeaders(method: string): string[] {
  const index = PRICE_METHODS.indexOf(method as (typeof PRICE_METHODS)[number]);
  if (index < 0) throw new Error('请选择有效价格公式');
  const used = new Set([...parameterNames(index + 1), '分标名称', '价格公式']);
  return PRICE_RULE_HEADERS.filter((h) => used.has(h));
}

function required(row: Record<string, string>, key: string): string {
  const s = (row[key] ?? '').trim();
  if (!s) throw new Error(`缺少${key}`);
  return s;
}

function cleanMissing(v: string | undefined): string {
  if (!v) return '';
  const t = v.trim();
  return ['/', '／', '—', '-'].includes(t) ? '' : t;
}

function parseNumber(s: string, label: string): Dec {
  try {
    return Dec.of(s.replace(/,/g, '').trim());
  } catch {
    throw new Error(`${label}不是有效数字：${s}`);
  }
}

function checkRange(
  r: Rule,
  key: string,
  low: number,
  high: number,
  closedLow: boolean,
  closedHigh: boolean
): void {
  const n = r.values[key];
  const l = n.cmp(Dec.of(low));
  const h = n.cmp(Dec.of(high));
  if (l < 0 || (!closedLow && l === 0) || h > 0 || (!closedHigh && h === 0)) {
    throw new Error(
      `参数${key}超出范围：${closedLow ? '[' : '('}${low},${high}${closedHigh ? ']' : ')'}`
    );
  }
}

function validateRule(r: Rule): void {
  checkRange(r, 'qz', 0, 100, true, true);
  for (const n of ['n1', 'n2']) {
    if (r.values[n] !== undefined && r.values[n].signum() <= 0) {
      throw new Error(`${n}必须大于0`);
    }
  }
  if (r.method === 4) {
    checkRange(r, 'b', -1, 0, false, false);
    checkRange(r, 'c', 0, 1, false, false);
  } else {
    checkRange(r, 'c', -0.1, 0.1, false, false);
  }
  if (r.method === 1) {
    checkRange(r, 'w1', -1, 0, false, false);
    checkRange(r, 'w2', 0, 1, false, false);
  }
  if (r.method === 3) {
    checkRange(r, 'F1', 0, 1, false, false);
    checkRange(r, 'F2', 1, 2, false, false);
  }
  if (r.method === 5) {
    checkRange(r, 'm', 0, 1, true, false);
    if (r.values.F1.signum() <= 0 || r.values.F2.signum() <= 0) {
      throw new Error('F1、F2必须大于0');
    }
  }
  if (r.method === 6) {
    checkRange(r, 'a', -1, 0, false, false);
    checkRange(r, 'b', 0, 1, false, false);
  }
}

function readRules(buffer: Buffer, fileName: string): Map<string, Rule> {
  const rules = new Map<string, Rule>();
  const optional = colsFromList(PRICE_RULE_HEADERS.slice(2));
  const tables = readWorkbookAllSheets(buffer, fileName, colsFromList(PRICE_RULE_HEADERS.slice(0, 2)), optional);
  for (const table of tables.values()) {
    for (const row of table.rows) {
      const section = required(row, '分标名称');
      const name = required(row, '价格公式');
      let method = PRICE_METHODS.indexOf(name as (typeof PRICE_METHODS)[number]) + 1;
      if (method === 0 && /^1001300[1-6]$/.test(name)) {
        method = parseInt(name.slice(7), 10);
      }
      if (method < 1 || method > 6) throw new Error(`未知价格公式：${name}`);
      const fields = parameterNames(method);
      const params: Record<string, Dec> = {};
      for (const f of fields) {
        params[f] = parseNumber(numericText(row, f), `参数${f}`);
      }
      const rule: Rule = { method, values: params };
      validateRule(rule);
      if (rules.has(section)) {
        throw new Error(`同一分标配置了多条价格公式：${section}`);
      }
      rules.set(section, rule);
    }
  }
  if (rules.size === 0) throw new Error('参数表没有分标配置');
  return rules;
}

function avg(list: Dec[]): Dec {
  let sum = Dec.ZERO;
  for (const v of list) sum = sum.add(v);
  return sum.div(Dec.of(list.length), 18, 'HALF_UP');
}

function filterInterval(list: Dec[], low: Dec, high: Dec, open: boolean): Dec[] {
  return list.filter((v) =>
    open ? v.cmp(low) > 0 && v.cmp(high) < 0 : v.cmp(low) >= 0 && v.cmp(high) <= 0
  );
}

function computeBase(r: Rule, prices: Dec[]): Dec {
  let a = avg(prices);
  const one = Dec.ONE;
  let selected: Dec[];
  switch (r.method) {
    case 1: {
      const sorted = [...prices].sort((x, y) => x.cmp(y));
      const n = prices.length;
      const low = n <= 5 ? 0 : n <= 20 ? 1 : n <= 30 ? 2 : 3;
      const high = n <= 5 ? 0 : n <= 10 ? 1 : n <= 20 ? 2 : n <= 30 ? 3 : 4;
      const remaining = sorted.slice(low, n - high);
      a = avg(remaining);
      selected = filterInterval(
        remaining,
        a.mul(one.add(r.values.w1)),
        a.mul(one.add(r.values.w2)),
        false
      );
      return (selected.length === 0 ? a : avg(selected)).mul(one.sub(r.values.c));
    }
    case 2:
      return a.mul(one.sub(r.values.c));
    case 3:
      selected = filterInterval(prices, a.mul(r.values.F1), a.mul(r.values.F2), false);
      return (selected.length === 0 ? a : avg(selected)).mul(one.add(r.values.c));
    case 4: {
      selected = filterInterval(
        prices,
        a.mul(one.add(r.values.b)),
        a.mul(one.add(r.values.c)),
        false
      );
      if (selected.length > 0) return avg(selected);
      let min = prices[0];
      for (const v of prices) if (v.cmp(min) < 0) min = v;
      const minCount = prices.filter((v) => v.eq(min)).length;
      if (minCount > 1) {
        throw new Error(
          '区间复合平均价法筛选为空且最低价并列：原版以并列最低价之和计算，口径有歧义，请确认采购公式后再处理'
        );
      }
      return a.add(min).div(Dec.of(2), 6, 'HALF_UP');
    }
    case 5: {
      const cutoff = a.mul(one.sub(r.values.m));
      selected = prices.filter((v) => v.cmp(cutoff) >= 0);
      return avg(selected).mul(one.add(r.values.c));
    }
    default: {
      if (prices.length <= 5) return a.mul(one.add(r.values.c));
      selected = filterInterval(
        prices,
        a.mul(one.add(r.values.a)),
        a.mul(one.add(r.values.b)),
        true
      );
      return (selected.length === 0 ? a : avg(selected)).mul(one.add(r.values.c));
    }
  }
}

function computeScore(r: Rule, price: Dec, base: Dec): Dec {
  const above = price.cmp(base) >= 0;
  const n = r.values[r.method === 5 ? (above ? 'F1' : 'F2') : above ? 'n1' : 'n2'];
  let result: Dec;
  if (r.method === 3) {
    const ratio = (above ? base : price).div(above ? price : base, 6, 'HALF_UP');
    const powVal = Math.pow(ratio.toNumber(), n.toNumber());
    result = Dec.HUNDRED.mul(Dec.of(String(powVal)));
  } else {
    const deduction = Dec.HUNDRED.mul(n)
      .mul(price.sub(base).abs())
      .div(base, 6, 'HALF_UP');
    result = Dec.HUNDRED.sub(deduction).max(Dec.ZERO);
  }
  return result.setScale(2, 'HALF_UP');
}

export function runPriceScoreCalculation(
  quoteFiles: { name: string; buffer: Buffer }[],
  ruleFile: { name: string; buffer: Buffer },
  roundFilter = ''
): PriceScoreResult {
  if (!quoteFiles.length || !ruleFile) {
    throw new Error('请选择报价文件及一份价格公式参数表');
  }
  const rules = readRules(ruleFile.buffer, ruleFile.name);
  const groups = new Map<string, Bid[]>();
  const round = roundFilter.trim();

  const reqCols = colsFromList(PRICE_SCORE_HEADERS.slice(0, 4));
  reqCols['分包名称'] = ['分包名称', '包名', '包名称'];
  reqCols['投标人名称'] = ['投标人名称', '投标人'];
  const optCols = colsFromList(PRICE_SCORE_HEADERS.slice(4, 7));

  for (const file of quoteFiles) {
    const tables = readWorkbookAllSheets(file.buffer, file.name, reqCols, optCols);
    for (const [sheetName, table] of tables.entries()) {
      for (const row of table.rows) {
        const participation = (row['是否参与计算'] ?? '').trim();
        if (participation === '否') continue;
        if (participation !== '' && participation !== '是') {
          throw new Error('是否参与计算只能填写是或否');
        }
        const r = required(row, '轮次');
        if (round && round !== r) continue;

        const source = `${file.name} / ${sheetName} 第${row.__row}行`;
        const total = cleanMissing(numericText(row, '投标价格（万元）'));
        const discount = cleanMissing(numericText(row, '折扣率'));
        if ((total === '') === (discount === '')) {
          throw new Error(`${source}：投标价格与折扣率必须恰好填写一项`);
        }
        const isDiscount = total === '';
        let value: Dec;
        if (isDiscount) {
          const display = row['折扣率'] ?? '';
          value = parseNumber(discount.replace(/[%％]/g, ''), source);
          if (display.endsWith('%') || display.endsWith('％')) {
            if (row['__numeric:折扣率'] !== undefined) {
              value = value.mul(Dec.HUNDRED);
            }
          } else {
            value = value.mul(Dec.HUNDRED);
          }
        } else {
          value = parseNumber(total, source);
        }
        if (value.signum() <= 0) throw new Error(`${source}：报价必须大于0`);

        const section = required(row, '分标名称');
        const pack = required(row, '分包名称');
        const supplier = required(row, '投标人名称');
        const gk = `${section}\u0000${pack}`;
        const arr = groups.get(gk) || [];
        arr.push({ section, pack, round: r, supplier, value, discount: isDiscount, source });
        groups.set(gk, arr);
      }
    }
  }

  if (groups.size === 0) throw new Error('所选轮次没有可计算的报价');

  const rows: PriceScoreRow[] = [];
  for (const bids of groups.values()) {
    const { section, pack } = bids[0];
    const rule = rules.get(section);
    if (!rule) throw new Error(`分标缺少公式参数：${section}`);
    if (new Set(bids.map((b) => b.round)).size !== 1) {
      throw new Error(`${section} / ${pack} 存在多个轮次，请填写要计算的轮次或仅选单一轮次文件`);
    }
    if (new Set(bids.map((b) => b.discount)).size !== 1) {
      throw new Error(`${section} / ${pack} 混有总价和折扣率，不可混算`);
    }
    if (new Set(bids.map((b) => b.supplier)).size !== bids.length) {
      throw new Error(`${section} / ${pack} 同一投标人有重复记录，请先确认唯一报价`);
    }

    const base = computeBase(
      rule,
      bids.map((b) => b.value)
    );
    if (base.signum() <= 0) throw new Error('基准值必须大于0');

    const methodName = PRICE_METHODS[rule.method - 1];
    const baseText = base.setScale(6, 'HALF_UP').fmt();

    for (const bid of bids) {
      const score = computeScore(rule, bid.value, base);
      const weighted = score.divExact(Dec.HUNDRED).mul(rule.values.qz).setScale(2, 'HALF_UP');
      rows.push({
        section,
        pack,
        round: bid.round,
        supplier: bid.supplier,
        quote_basis: bid.discount ? '折扣率（%）' : '总价（万元）',
        quote_value: bid.value.fmt(),
        method_name: methodName,
        base_value: baseText,
        price_score: score.fmt(),
        weighted_score: weighted.fmt(),
        source: bid.source,
      });
    }
  }

  return {
    quoteFiles: quoteFiles.map((f) => f.name),
    ruleFile: ruleFile.name,
    roundFilter: round,
    totalRows: rows.length,
    packageCount: groups.size,
    rows,
    warnings: [
      '参与计算的报价应为已确认有效、进入详评的报价；可在「是否参与计算」列填「否」排除，空白按参与计算处理；本工具不会自动判定否决。参数须按本批采购文件填写。',
    ],
  };
}

export async function exportPriceScoreExcel(result: PriceScoreResult): Promise<Buffer> {
  return buildStyledExcelBuffer([
    {
      name: '价格分',
      headers: [
        '分标名称',
        '分包名称',
        '轮次',
        '投标人名称',
        '报价口径',
        '报价',
        '价格公式',
        '基准值',
        '价格分',
        '加权价格分',
        '来源',
      ],
      rows: result.rows.map((r) => [
        r.section,
        r.pack,
        r.round,
        r.supplier,
        r.quote_basis,
        r.quote_value,
        r.method_name,
        r.base_value,
        r.price_score,
        r.weighted_score,
        r.source,
      ]),
    },
  ]);
}

export async function exportPriceScoreTemplate(kind: 'quotes' | 'rules', method?: string): Promise<Buffer> {
  if (kind === 'quotes') {
    return buildStyledExcelBuffer([
      { name: '报价文件', headers: PRICE_SCORE_HEADERS, rows: [] },
    ]);
  }
  const m = method || PRICE_METHODS[0];
  const headers = priceRuleTemplateHeaders(m);
  const sampleRow = headers.map((h) => (h === '价格公式' ? m : ''));
  return buildStyledExcelBuffer([
    { name: '价格公式参数表', headers, rows: [sampleRow] },
  ]);
}
