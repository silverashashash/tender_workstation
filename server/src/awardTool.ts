import XLSXModule from 'xlsx';
import { evaluateWorkbookFormulas } from './excelFormulaEval.js';
import {
  Dec,
  buildStyledExcelBuffer,
  colsFromList,
  numericText,
  readCellVal,
  readWorkbookAllSheets,
} from './toolboxWorkbook.js';

const XLSX = (XLSXModule as unknown as { default?: typeof XLSXModule }).default ?? XLSXModule;

export const AWARD_MODES = ['无限授', '一轮一包', '一轮多包'] as const;
export type AwardMode = (typeof AWARD_MODES)[number];

export const AWARD_INPUT_HEADERS = ['分标编号', '分标名称', '包名称', '投标人', '排序', '投标总价'];
export const AWARD_GROUP_HEADERS = ['限授组', '分标编号', '包名称', '最大中包数', '授标顺序'];

interface Candidate {
  bidder: string;
  rank: number;
  price: Dec;
  source: string;
}

interface Pack {
  section: string;
  pack: string;
  name: string;
  order: number;
  candidates: Candidate[];
  group: string;
  note: string;
  status: string;
  winner: string;
  source: string;
  priority: number;
  limit: number;
  round: number;
}

interface Choice {
  pack: Pack;
  candidate: Candidate;
}

export interface AwardResultRow {
  section_code: string;
  section_name: string;
  pack_name: string;
  limit_group: string;
  winner: string;
  status: string;
  round: string;
  mode: string;
  note: string;
  source: string;
}

export interface AwardAllocationResult {
  files: string[];
  groupFile: string;
  mode: AwardMode;
  totalPacks: number;
  awardedCount: number;
  unallocatedCount: number;
  rows: AwardResultRow[];
}

function packKey(section: string, pack: string): string {
  return `${section}\u0000${pack}`;
}

function formatKeyLabel(section: string, pack: string): string {
  return `Key[section=${section}, pack=${pack}]`;
}

function required(row: Record<string, string>, key: string): string {
  const s = (row[key] ?? '').trim();
  if (!s) throw new Error(`第${row.__row}行缺少${key}`);
  return s;
}

function positiveInt(text: string): number {
  const n = Dec.of(text.trim()).intValueExact();
  if (n < 1) throw new Error('排序、限授数量和授标顺序必须为正整数');
  return n;
}

function addCandidate(
  packs: Map<string, Pack>,
  code: string,
  name: string,
  pack: string,
  supplier: string,
  rank: string,
  price: string,
  source: string
): void {
  if (!code.trim() || !name.trim() || !pack.trim()) {
    throw new Error(`${source} 缺少分标编号、分标名称或包名称`);
  }
  const c = code.trim();
  const n = name.trim();
  const pk = pack.trim();
  const k = packKey(c, pk);
  let p = packs.get(k);
  if (!p) {
    p = {
      section: c,
      pack: pk,
      name: n,
      order: packs.size,
      candidates: [],
      group: '',
      note: '',
      status: '',
      winner: '',
      source: '',
      priority: 1,
      limit: 0,
      round: 0,
    };
    packs.set(k, p);
  }
  if (p.name !== n) {
    throw new Error(`${source} 同一分标编号的名称不一致`);
  }
  if (!supplier.trim()) {
    if (rank.trim() || price.trim()) throw new Error(`${source} 缺少投标人`);
    return;
  }
  const sup = supplier.trim();
  if (p.candidates.some((x) => x.bidder === sup)) {
    throw new Error(`${source} 同一包投标人重复`);
  }
  const amount = Dec.of(price.trim().replace(/,/g, ''));
  if (amount.signum() < 0) throw new Error(`${source} 投标总价不能为负数`);
  p.candidates.push({ bidder: sup, rank: positiveInt(rank), price: amount, source });
}

function readRankingFile(file: { name: string; buffer: Buffer }, packs: Map<string, Pack>): void {
  let wb: XLSXModule.WorkBook;
  try {
    wb = XLSX.read(file.buffer, { type: 'buffer', cellDates: false, sheetStubs: true, cellNF: true, cellStyles: true });
  } catch (err) {
    throw new Error(`${file.name} 无法读取：${(err as Error).message}`);
  }
  evaluateWorkbookFormulas(wb);

  const firstSheet = wb.SheetNames.length > 0 ? wb.Sheets[wb.SheetNames[0]] : null;
  const isOriginal =
    firstSheet !== null && readCellVal(firstSheet, 5, 4, wb.SheetNames[0]).display.trim() === '投标人';

  if (!isOriginal) {
    const tables = readWorkbookAllSheets(file.buffer, file.name, colsFromList(AWARD_INPUT_HEADERS));
    for (const [sheetName, table] of tables.entries()) {
      for (const row of table.rows) {
        addCandidate(
          packs,
          required(row, '分标编号'),
          required(row, '分标名称'),
          required(row, '包名称'),
          row['投标人'] ?? '',
          numericText(row, '排序'),
          numericText(row, '投标总价'),
          `${file.name} / ${sheetName} 第${row.__row}行`
        );
      }
    }
    return;
  }

  const localKeys = new Set<string>();
  const codeRe = /分标编号\s*[:：]\s*([\s\S]*?)\s*价格单位/;
  const nameRe = /(?:设备\/材料名称|分标名称)\s*[:：]\s*([\s\S]*?)\s*日期/;

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const location = `${file.name} / ${sheetName}`;
    const meta = readCellVal(ws, 3, 0, sheetName).display.trim();
    const cm = meta.match(codeRe);
    const nm = meta.match(nameRe);
    if (!cm || !nm) {
      throw new Error(`${location} 第4行缺少分标编号/设备材料名称`);
    }
    const section = cm[1].trim();
    const sectionName = nm[1].trim();

    const h4 = readCellVal(ws, 5, 4, sheetName).display.trim();
    const h6 = readCellVal(ws, 5, 6, sheetName).display.trim();
    const h8 = readCellVal(ws, 5, 8, sheetName).display.trim();
    const h9 = readCellVal(ws, 5, 9, sheetName).display.trim();
    const h11 = readCellVal(ws, 5, 11, sheetName).display.trim();

    const full = h9 === '排序' && h11.includes('投标总价');
    const compact = h6 === '排序' && h8.includes('投标总价');
    if (h4 !== '投标人' || (!full && !compact)) {
      throw new Error(`${location} 第6行不是支持的经济排序表`);
    }

    let pack = readCellVal(ws, 6, 0, sheetName).display.trim();
    if (!pack || pack === '无' || pack === '无数据') pack = sheetName;
    const k = packKey(section, pack);
    if (packs.has(k) || localKeys.has(k)) {
      throw new Error(`${location} 重复包：${formatKeyLabel(section, pack)}`);
    }
    localKeys.add(k);
    addCandidate(packs, section, sectionName, pack, '', '', '', location);

    const ref = ws['!ref'];
    const lastRow = ref ? XLSX.utils.decode_range(ref).e.r : 5;
    for (let r = 6; r <= lastRow; r++) {
      const supplier = readCellVal(ws, r, 4, sheetName).display.trim();
      const rank = readCellVal(ws, r, full ? 9 : 6, sheetName).display.trim();
      const price = readCellVal(ws, r, full ? 11 : 8, sheetName).numericText;
      if (!supplier && !rank && !price.trim()) continue;
      addCandidate(packs, section, sectionName, pack, supplier, rank, price, `${location} 第${r + 1}行`);
    }
  }
}

function configureGroups(
  groupFile: { name: string; buffer: Buffer } | null,
  packs: Map<string, Pack>
): void {
  if (!groupFile) throw new Error('限授模式需要一份限授组配置表');
  const tables = readWorkbookAllSheets(
    groupFile.buffer,
    groupFile.name,
    colsFromList(AWARD_GROUP_HEADERS)
  );
  const seen = new Set<string>();
  const limits = new Map<string, number>();

  for (const table of tables.values()) {
    for (const row of table.rows) {
      const sec = required(row, '分标编号');
      const pk = required(row, '包名称');
      const k = packKey(sec, pk);
      const p = packs.get(k);
      if (!p) throw new Error(`配置中的包不存在：${formatKeyLabel(sec, pk)}`);
      if (seen.has(k)) throw new Error(`同一包重复配置限授组：${formatKeyLabel(sec, pk)}`);
      seen.add(k);

      p.group = required(row, '限授组');
      p.note = '请确认组内报价单位及口径一致；金额与折扣率不可混合比较';
      p.limit = positiveInt(numericText(row, '最大中包数'));
      p.priority = positiveInt(numericText(row, '授标顺序'));
      const prior = limits.get(p.group);
      if (prior !== undefined && prior !== p.limit) {
        throw new Error(`同一限授组最大中包数不一致：${p.group}`);
      }
      limits.set(p.group, p.limit);
    }
  }
  if (seen.size === 0) {
    throw new Error('限授配置为空；如不需要限授，请选择无限授');
  }
}

function compareChoices(a: Choice, b: Choice): number {
  const priceCmp = b.candidate.price.cmp(a.candidate.price);
  if (priceCmp !== 0) return priceCmp;
  if (a.pack.priority !== b.pack.priority) return a.pack.priority - b.pack.priority;
  return a.pack.order - b.pack.order;
}

function markTies(choices: Choice[]): void {
  const warning = '；报价及授标顺序相同，按输入文件/工作表顺序处理，请复核';
  for (let i = 1; i < choices.length; i++) {
    const a = choices[i - 1];
    const b = choices[i];
    if (a.candidate.price.eq(b.candidate.price) && a.pack.priority === b.pack.priority) {
      if (!a.pack.note.includes(warning)) a.pack.note += warning;
      if (!b.pack.note.includes(warning)) b.pack.note += warning;
    }
  }
}

function awardChoice(c: Choice, round: number): void {
  c.pack.winner = c.candidate.bidder;
  c.pack.status = '拟授标';
  c.pack.round = round;
  c.pack.source = c.candidate.source;
}

function allocateGroup(packs: Pack[], mode: AwardMode): void {
  const counts = new Map<string, number>();
  let round = 1;
  while (true) {
    const choices: Choice[] = [];
    for (const p of packs) {
      if (p.status !== '') continue;
      const cand = p.candidates.find((x) => (counts.get(x.bidder) ?? 0) < p.limit);
      if (cand) {
        choices.push({ pack: p, candidate: cand });
      } else {
        p.status = p.candidates.length === 0 ? '无候选' : '限授后无可授候选';
      }
    }
    if (choices.length === 0) break;

    if (mode === '一轮一包') {
      choices.sort(compareChoices);
      markTies(choices);
      const selected = choices[0];
      awardChoice(selected, round++);
      counts.set(selected.candidate.bidder, (counts.get(selected.candidate.bidder) ?? 0) + 1);
    } else {
      const byBidder = new Map<string, Choice[]>();
      for (const c of choices) {
        const list = byBidder.get(c.candidate.bidder) || [];
        list.push(c);
        byBidder.set(c.candidate.bidder, list);
      }
      for (const group of byBidder.values()) {
        group.sort(compareChoices);
        markTies(group);
        const bidder = group[0].candidate.bidder;
        const available = group[0].pack.limit - (counts.get(bidder) ?? 0);
        for (let i = 0; i < Math.min(available, group.length); i++) {
          const c = group[i];
          awardChoice(c, round);
          counts.set(bidder, (counts.get(bidder) ?? 0) + 1);
        }
      }
      round++;
    }
  }
}

export function runAwardAllocation(
  rankingFiles: { name: string; buffer: Buffer }[],
  groupFile: { name: string; buffer: Buffer } | null,
  mode: AwardMode = '无限授'
): AwardAllocationResult {
  if (!rankingFiles || rankingFiles.length === 0) {
    throw new Error('请选择技术经济排序表');
  }
  if (!AWARD_MODES.includes(mode)) {
    throw new Error('授标方式应为：无限授、一轮一包或一轮多包');
  }

  const packs = new Map<string, Pack>();
  for (const f of rankingFiles) {
    readRankingFile(f, packs);
  }
  if (packs.size === 0) throw new Error('没有可处理的包');

  for (const p of packs.values()) {
    p.candidates.sort((a, b) => a.rank - b.rank);
    const ranks = new Set<number>();
    for (const c of p.candidates) {
      if (ranks.has(c.rank)) {
        throw new Error(`${formatKeyLabel(p.section, p.pack)} 排序并列，请先确认唯一排序后授标`);
      }
      ranks.add(c.rank);
    }
  }

  if (mode !== '无限授') {
    configureGroups(groupFile, packs);
  }

  for (const p of packs.values()) {
    if (mode === '无限授' || p.group === '') {
      if (p.candidates.length === 0) {
        p.status = '无候选';
      } else {
        awardChoice({ pack: p, candidate: p.candidates[0] }, 1);
      }
    }
  }

  if (mode !== '无限授') {
    const groups = new Map<string, Pack[]>();
    for (const p of packs.values()) {
      if (p.group !== '') {
        const list = groups.get(p.group) || [];
        list.push(p);
        groups.set(p.group, list);
      }
    }
    for (const group of groups.values()) {
      allocateGroup(group, mode);
    }
  }

  const rows: AwardResultRow[] = [];
  let awardedCount = 0;
  for (const p of packs.values()) {
    if (p.status === '拟授标') awardedCount++;
    rows.push({
      section_code: p.section,
      section_name: p.name,
      pack_name: p.pack,
      limit_group: p.group,
      winner: p.winner,
      status: p.status,
      round: p.round === 0 ? '' : String(p.round),
      mode,
      note: p.note,
      source: p.source,
    });
  }

  return {
    files: rankingFiles.map((f) => f.name),
    groupFile: groupFile ? groupFile.name : '',
    mode,
    totalPacks: rows.length,
    awardedCount,
    unallocatedCount: rows.length - awardedCount,
    rows,
  };
}

export async function exportAwardAllocationExcel(result: AwardAllocationResult): Promise<Buffer> {
  return buildStyledExcelBuffer([
    {
      name: '授标结果',
      headers: ['分标编号', '分标名称', '包名称', '限授组', '拟授标人', '状态', '轮次', '授标方式', '说明', '来源'],
      rows: result.rows.map((r) => [
        r.section_code,
        r.section_name,
        r.pack_name,
        r.limit_group,
        r.winner,
        r.status,
        r.round,
        r.mode,
        r.note,
        r.source,
      ]),
    },
  ]);
}

export async function exportAwardTemplate(kind: 'ranking' | 'groups'): Promise<Buffer> {
  if (kind === 'ranking') {
    return buildStyledExcelBuffer([{ name: '技术经济排序表', headers: AWARD_INPUT_HEADERS, rows: [] }]);
  }
  return buildStyledExcelBuffer([{ name: '限授组配置表', headers: AWARD_GROUP_HEADERS, rows: [] }]);
}
