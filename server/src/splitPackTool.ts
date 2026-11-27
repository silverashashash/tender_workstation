import {
  Dec,
  buildStyledExcelBuffer,
  colsFromList,
  numericText,
  readWorkbookAllSheets,
} from './toolboxWorkbook.js';

export const SPLIT_MATERIAL_HEADERS = [
  '序号',
  '物料编码',
  '物料描述',
  '物料长描述',
  '计量单位名称',
  '技术规范书ID',
  '扩展描述',
  '数量',
  '单价限价',
  '分配粒度（可空）',
];

export const SPLIT_RULE_HEADERS = ['包号', '分配值'];

export type SplitPackMode = '比例' | '金额';

interface Material {
  key: string;
  code: string;
  extension: string;
  description: string;
  longDescription: string;
  unit: string;
  specification: string;
  quantity: Dec;
  price: Dec;
  factor: Dec;
  step: Dec;
  normalizedQuantity: Dec;
  normalizedPrice: Dec;
}

interface Pack {
  number: string;
  value: Dec;
  target: Dec;
}

export interface SplitSummaryRow {
  pack_no: string;
  item_count: number;
  total_qty: string;
  total_budget: string;
  target_amount: string;
  diff_amount: string;
}

export interface SplitDetailRow {
  seq: number;
  pack_no: string;
  material_code: string;
  description: string;
  long_description: string;
  specification: string;
  extension: string;
  quantity: string;
  unit: string;
  unit_price: string;
  total_price: string;
}

export interface SplitPackResult {
  materialFile: string;
  ruleFile: string;
  mode: SplitPackMode;
  weightStep: string;
  lengthStep: string;
  packCount: number;
  materialCount: number;
  summaryRows: SplitSummaryRow[];
  detailRows: SplitDetailRow[];
  warnings: string[];
}

function nonblank(s: string | undefined, label: string): string {
  if (!s || !s.trim()) throw new Error(`${label}不能为空`);
  return s;
}

function positiveDec(text: string | undefined, label: string): Dec {
  try {
    const n = Dec.of((text ?? '').trim());
    if (n.signum() <= 0) throw new Error();
    return n;
  } catch {
    throw new Error(`${label}必须为有效正数：${text ?? ''}`);
  }
}

function addQty(map: Map<string, Dec>, key: string, qty: Dec): void {
  if (qty.signum() > 0) {
    map.set(key, (map.get(key) ?? Dec.ZERO).add(qty));
  }
}

export function runSplitPack(
  materialFile: { name: string; buffer: Buffer },
  ruleFile: { name: string; buffer: Buffer },
  mode: SplitPackMode = '比例',
  weightStepStr = '1',
  lengthStepStr = '1'
): SplitPackResult {
  if (!['比例', '金额'].includes(mode)) {
    throw new Error('分配方式必须为比例或金额');
  }
  const weight = positiveDec(weightStepStr, '重量粒度（千克）');
  const length = positiveDec(lengthStepStr, '长度粒度（米）');

  const materials: Material[] = [];
  const keys = new Set<string>();
  let total = Dec.ZERO;

  const reqMat = colsFromList(SPLIT_MATERIAL_HEADERS.slice(0, 9));
  const optMat = { '分配粒度（可空）': ['分配粒度（可空）', '分配粒度'] };
  const matTables = readWorkbookAllSheets(materialFile.buffer, materialFile.name, reqMat, optMat);

  for (const [sheetName, table] of matTables.entries()) {
    for (const row of table.rows) {
      const location = `${sheetName} 第${row.__row}行`;
      const code = nonblank(row['物料编码'], `${location}物料编码`);
      const extension = row['扩展描述'] ?? '';
      const key = `${code}\u0000${extension}`;
      if (keys.has(key)) {
        throw new Error(`${location} 物料编码＋扩展描述重复，请先合并数量`);
      }
      keys.add(key);

      const unit = nonblank(row['计量单位名称'], `${location}计量单位名称`);
      const factor = unit === '吨' || unit === '千米' ? Dec.THOUSAND : Dec.ONE;
      const quantity = positiveDec(numericText(row, '数量'), `${location}数量`);
      const price = positiveDec(numericText(row, '单价限价'), `${location}单价限价`);
      const override = numericText(row, '分配粒度（可空）') ?? '';
      const step = !override.trim()
        ? unit === '吨' || unit === '千克'
          ? weight
          : unit === '米' || unit === '千米'
            ? length
            : Dec.ONE
        : positiveDec(override, `${location}分配粒度`).mul(factor);

      const normalizedQuantity = quantity.mul(factor);
      const normalizedPrice = price.divExact(factor);

      materials.push({
        key,
        code,
        extension,
        description: row['物料描述'] ?? '',
        longDescription: row['物料长描述'] ?? '',
        unit,
        specification: row['技术规范书ID'] ?? '',
        quantity,
        price,
        factor,
        step,
        normalizedQuantity,
        normalizedPrice,
      });
      total = total.add(quantity.mul(price));
    }
  }

  if (materials.length === 0) throw new Error('物料表没有数据');

  const packs: Pack[] = [];
  const packKeys = new Set<string>();
  let valueTotal = Dec.ZERO;
  const ruleTables = readWorkbookAllSheets(
    ruleFile.buffer,
    ruleFile.name,
    colsFromList(SPLIT_RULE_HEADERS)
  );
  for (const table of ruleTables.values()) {
    for (const row of table.rows) {
      const number = nonblank(row['包号'], '包号');
      if (packKeys.has(number)) throw new Error(`包号重复：${number}`);
      packKeys.add(number);
      const value = positiveDec(numericText(row, '分配值'), `包${number}分配值`);
      packs.push({ number, value, target: Dec.ZERO });
      valueTotal = valueTotal.add(value);
    }
  }
  if (packs.length === 0) throw new Error('分包规则表没有数据');
  if (mode === '金额' && valueTotal.cmp(total) > 0) {
    throw new Error('各包目标金额合计超过物料总金额');
  }

  for (let i = 0; i < packs.length; i++) {
    const p = packs[i];
    const target =
      mode === '比例' ? total.div(valueTotal, 4, 'DOWN').mul(p.value) : p.value;
    packs[i] = { number: p.number, value: p.value, target };
  }

  packs.sort((a, b) => {
    const t = b.target.cmp(a.target);
    if (t !== 0) return t;
    return a.number < b.number ? -1 : a.number > b.number ? 1 : 0;
  });

  const quantities = new Map<string, Map<string, Dec>>();
  const remaining = new Map<string, Dec>();
  const budget = new Map<string, Dec>();

  for (const m of materials) remaining.set(m.key, m.normalizedQuantity);
  for (const p of packs) {
    quantities.set(p.number, new Map());
    budget.set(p.number, p.target);
  }

  let remainingTarget = Dec.ZERO;
  for (const p of packs) remainingTarget = remainingTarget.add(p.target);

  // Pass 1: one allocation unit per material and package, even when target is exceeded.
  for (const p of packs) {
    const pMap = quantities.get(p.number)!;
    for (const m of materials) {
      const rem = remaining.get(m.key)!;
      const allocated = rem.min(m.step);
      addQty(pMap, m.key, allocated);
      remaining.set(m.key, rem.sub(allocated));
      const cost = allocated.mul(m.normalizedPrice);
      budget.set(p.number, budget.get(p.number)!.sub(cost).max(Dec.ZERO));
      remainingTarget = remainingTarget.sub(cost);
    }
  }

  // Pass 2: frozen remaining stock multiplied by each package's remaining-target share.
  const snapshot = new Map<string, Dec>(remaining);
  if (remainingTarget.signum() > 0) {
    for (const p of packs) {
      const pMap = quantities.get(p.number)!;
      const ratio = budget.get(p.number)!.div(remainingTarget, 4, 'HALF_DOWN');
      for (const m of materials) {
        const snap = snapshot.get(m.key)!;
        const rem = remaining.get(m.key)!;
        const amount = snap
          .mul(ratio)
          .setScale(0, 'HALF_DOWN')
          .div(m.step, 0, 'HALF_DOWN')
          .mul(m.step)
          .min(rem);
        addQty(pMap, m.key, amount);
        remaining.set(m.key, rem.sub(amount));
      }
    }
  }

  // Pass 3: expensive stock first; aggregate full round-robin cycles.
  const tail = [...materials].sort((a, b) => b.normalizedPrice.cmp(a.normalizedPrice));
  let cursor = 0;
  const count = BigInt(packs.length);

  for (const m of tail) {
    const stock = remaining.get(m.key)!;
    const [wholeChunks, remPart] = stock.divAndRem(m.step);
    const chunkCount = wholeChunks.toBigIntExact();
    const fullRounds = chunkCount / count;
    const leftoverSteps = Number(chunkCount % count);

    if (fullRounds > 0n) {
      const each = m.step.mul(Dec.of(fullRounds));
      for (const p of packs) {
        addQty(quantities.get(p.number)!, m.key, each);
      }
    }
    for (let i = 0; i < leftoverSteps; i++) {
      addQty(quantities.get(packs[cursor].number)!, m.key, m.step);
      cursor = (cursor + 1) % packs.length;
    }
    if (remPart.signum() > 0) {
      addQty(quantities.get(packs[cursor].number)!, m.key, remPart);
      cursor = (cursor + 1) % packs.length;
    }
    remaining.set(m.key, Dec.ZERO);
  }

  // Conservation check on quantities
  for (const m of materials) {
    let sum = Dec.ZERO;
    for (const p of packs) {
      sum = sum.add(quantities.get(p.number)!.get(m.key) ?? Dec.ZERO);
    }
    if (!sum.eq(m.normalizedQuantity)) {
      throw new Error('分配数量不守恒');
    }
  }

  const detailRows: SplitDetailRow[] = [];
  const summaryRows: SplitSummaryRow[] = [];
  let grandMoney = Dec.ZERO;
  let grandQuantity = Dec.ZERO;
  let grandTarget = Dec.ZERO;
  let grandItems = 0;

  const outputPacks = [...packs].sort((a, b) => (a.number < b.number ? -1 : a.number > b.number ? 1 : 0));

  for (const p of outputPacks) {
    let amount = Dec.ZERO;
    let qtySum = Dec.ZERO;
    let items = 0;
    const pMap = quantities.get(p.number)!;

    for (const m of materials) {
      const normalized = pMap.get(m.key) ?? Dec.ZERO;
      if (normalized.signum() === 0) continue;
      const q = normalized.divExact(m.factor);
      const money = q.mul(m.price);
      items++;
      amount = amount.add(money);
      qtySum = qtySum.add(q);
      detailRows.push({
        seq: detailRows.length + 1,
        pack_no: p.number,
        material_code: m.code,
        description: m.description,
        long_description: m.longDescription,
        specification: m.specification,
        extension: m.extension,
        quantity: q.fmt(),
        unit: m.unit,
        unit_price: m.price.fmt(),
        total_price: money.fmt(),
      });
    }

    summaryRows.push({
      pack_no: p.number,
      item_count: items,
      total_qty: qtySum.fmt(),
      total_budget: amount.fmt(),
      target_amount: p.target.fmt(),
      diff_amount: amount.sub(p.target).fmt(),
    });
    grandMoney = grandMoney.add(amount);
    grandQuantity = grandQuantity.add(qtySum);
    grandTarget = grandTarget.add(p.target);
    grandItems += items;
  }

  if (!grandMoney.eq(total)) {
    throw new Error('分配金额不守恒');
  }

  summaryRows.push({
    pack_no: '合计',
    item_count: grandItems,
    total_qty: grandQuantity.fmt(),
    total_budget: grandMoney.fmt(),
    target_amount: grandTarget.fmt(),
    diff_amount: grandMoney.sub(grandTarget).fmt(),
  });

  return {
    materialFile: materialFile.name,
    ruleFile: ruleFile.name,
    mode,
    weightStep: weight.fmt(),
    lengthStep: length.fmt(),
    packCount: packs.length,
    materialCount: materials.length,
    summaryRows,
    detailRows,
    warnings: [
      '分包结果为建议；目标金额不是硬限额。请核对实际金额与目标偏差；不同计量单位的数量合计仅作汇总。',
    ],
  };
}

export async function exportSplitPackExcel(result: SplitPackResult): Promise<Buffer> {
  return buildStyledExcelBuffer([
    {
      name: '汇总',
      headers: ['包号', '物料数量', '总数量', '总预算', '目标金额', '实际减目标'],
      rows: result.summaryRows.map((r) => [
        r.pack_no,
        String(r.item_count),
        r.total_qty,
        r.total_budget,
        r.target_amount,
        r.diff_amount,
      ]),
    },
    {
      name: '分包结果数据',
      headers: [
        '序号',
        '包号',
        '物料编码',
        '物料描述',
        '物料长描述',
        '技术规范书ID',
        '扩展描述',
        '数量',
        '计量单位名称',
        '预算单价',
        '预算总价',
      ],
      rows: result.detailRows.map((r) => [
        String(r.seq),
        r.pack_no,
        r.material_code,
        r.description,
        r.long_description,
        r.specification,
        r.extension,
        r.quantity,
        r.unit,
        r.unit_price,
        r.total_price,
      ]),
    },
  ]);
}

export async function exportSplitPackTemplate(kind: 'materials' | 'rules'): Promise<Buffer> {
  if (kind === 'materials') {
    return buildStyledExcelBuffer([{ name: '物料表', headers: SPLIT_MATERIAL_HEADERS, rows: [] }]);
  }
  return buildStyledExcelBuffer([{ name: '分包规则表', headers: SPLIT_RULE_HEADERS, rows: [] }]);
}
