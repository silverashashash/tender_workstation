import XLSXModule from 'xlsx';
import { evaluateWorkbookFormulas } from './excelFormulaEval.js';
import { Dec, buildStyledExcelBuffer, readCellVal } from './toolboxWorkbook.js';

const XLSX = (XLSXModule as unknown as { default?: typeof XLSXModule }).default ?? XLSXModule;

interface Candidate {
  supplier: string;
  technical: Dec;
  business: Dec;
  priceScore: Dec;
  total: Dec;
  originalRank: number;
  bid: Dec;
}

export interface EconomicRankIssue {
  section: string;
  pack: string;
  supplier: string;
  original_rank: string;
  verified_rank: string;
  reason: string;
}

export interface EconomicRankResult {
  files: string[];
  sheetCount: number;
  examinedCount: number;
  issueCount: number;
  issues: EconomicRankIssue[];
}

const SECTION_RE = /设备\/材料名称[:：]\s*(.*?)\s*日期[:：]/;
const NON_NEG_NUM_RE = /^(?:[0-9]+|[0-9]{1,3}(?:,[0-9]{3})+)(?:\.[0-9]+)?%?$/;

function parseNonNegNumber(
  ws: XLSXModule.WorkSheet,
  r: number,
  c: number,
  sheetName: string
): Dec {
  const val = readCellVal(ws, r, c, sheetName);
  const text = val.numericText.trim();
  if (!NON_NEG_NUM_RE.test(text)) {
    throw new Error(`${sheetName} 第${r + 1}行第${c + 1}列不是有效的非负数：${text}`);
  }
  return Dec.of(text.replace(/,/g, '').replace(/%/g, ''));
}

export function runEconomicRankCheck(
  files: { name: string; buffer: Buffer }[]
): EconomicRankResult {
  if (!files || files.length === 0) {
    throw new Error('请选择经济排序表');
  }

  const issues: EconomicRankIssue[] = [];
  let examinedCount = 0;
  let sheetCount = 0;

  for (const file of files) {
    let wb: XLSXModule.WorkBook;
    try {
      wb = XLSX.read(file.buffer, { type: 'buffer', cellDates: false, sheetStubs: true, cellNF: true, cellStyles: true });
    } catch (err) {
      throw new Error(`${file.name} 无法读取：${(err as Error).message}`);
    }
    evaluateWorkbookFormulas(wb);

    for (const sheetName of wb.SheetNames) {
      sheetCount++;
      const ws = wb.Sheets[sheetName];
      const location = `${file.name} / ${sheetName}`;
      const metaText = readCellVal(ws, 3, 0, sheetName).display;
      const metaMatch = metaText.match(SECTION_RE);
      if (!metaMatch || !metaMatch[1].trim()) {
        throw new Error(`${location} 第4行缺少设备/材料名称及日期，无法确认分标`);
      }
      const section = metaMatch[1];
      const pack = sheetName;

      const requiredHeaders: [number, string][] = [
        [4, '投标人'],
        [5, '技术'],
        [6, '商务'],
        [7, '价格'],
        [8, '总分'],
        [9, '排序'],
      ];
      for (const [colIdx, expected] of requiredHeaders) {
        const actual = readCellVal(ws, 5, colIdx, sheetName).display.trim();
        if (actual !== expected) {
          throw new Error(`${location} 第6行表头与经济排序表格式不符：${expected}`);
        }
      }
      const totalBidHeader = readCellVal(ws, 5, 11, sheetName).display;
      if (!totalBidHeader.includes('投标总价')) {
        throw new Error(`${location} 缺少投标总价表头`);
      }

      const ref = ws['!ref'];
      const lastRow = ref ? XLSX.utils.decode_range(ref).e.r : 5;
      const candidates: Candidate[] = [];
      const suppliers = new Set<string>();

      for (let r = 6; r <= lastRow; r++) {
        if (readCellVal(ws, r, 0, sheetName).display === '无数据') continue;
        const supplier = readCellVal(ws, r, 4, sheetName).display;
        let blank = true;
        for (let c = 0; c < 14; c++) {
          if (readCellVal(ws, r, c, sheetName).display.trim() !== '') {
            blank = false;
          }
        }
        if (blank) continue;

        if (!supplier.trim() || suppliers.has(supplier)) {
          throw new Error(`${location} 第${r + 1}行投标人为空或重复，无法确定排序记录`);
        }
        suppliers.add(supplier);

        const tech = parseNonNegNumber(ws, r, 5, sheetName);
        const biz = parseNonNegNumber(ws, r, 6, sheetName);
        const price = parseNonNegNumber(ws, r, 7, sheetName);
        const total = parseNonNegNumber(ws, r, 8, sheetName);
        const bid = parseNonNegNumber(ws, r, 11, sheetName);
        const rank = parseNonNegNumber(ws, r, 9, sheetName).intValueExact();
        if (rank < 1) {
          throw new Error(`${location} 第${r + 1}行排序必须为正整数`);
        }
        candidates.push({
          supplier,
          technical: tech,
          business: biz,
          priceScore: price,
          total,
          originalRank: rank,
          bid,
        });
        examinedCount++;
      }

      let hasZero = false;
      for (const c of candidates) {
        const reasons: string[] = [];
        if (c.technical.signum() === 0) reasons.push('技术分为0');
        if (c.business.signum() === 0) reasons.push('商务分为0');
        if (c.priceScore.signum() === 0) reasons.push('价格分为0');
        if (reasons.length > 0) {
          hasZero = true;
          issues.push({
            section,
            pack,
            supplier: c.supplier,
            original_rank: String(c.originalRank),
            verified_rank: '',
            reason: reasons.join('；'),
          });
        }
      }
      if (hasZero) continue;

      candidates.sort((a, b) => {
        const t = b.total.cmp(a.total);
        if (t !== 0) return t;
        const tech = b.technical.cmp(a.technical);
        if (tech !== 0) return tech;
        const bid = a.bid.cmp(b.bid);
        if (bid !== 0) return bid;
        return b.business.cmp(a.business);
      });

      for (let i = 0; i < candidates.length; i++) {
        const c = candidates[i];
        if (c.originalRank !== i + 1) {
          issues.push({
            section,
            pack,
            supplier: c.supplier,
            original_rank: String(c.originalRank),
            verified_rank: String(i + 1),
            reason: '排序不一致',
          });
        }
      }
    }
  }

  if (examinedCount === 0) {
    throw new Error('未找到可核验的供应商排序记录');
  }

  return {
    files: files.map((f) => f.name),
    sheetCount,
    examinedCount,
    issueCount: issues.length,
    issues,
  };
}

export async function exportEconomicRankExcel(result: EconomicRankResult): Promise<Buffer> {
  return buildStyledExcelBuffer([
    {
      name: '核验结果',
      headers: ['分标名称', '包名称', '投标人', '原排序', '核验排序', '原因'],
      rows: result.issues.map((x) => [
        x.section,
        x.pack,
        x.supplier,
        x.original_rank,
        x.verified_rank,
        x.reason,
      ]),
    },
  ]);
}
