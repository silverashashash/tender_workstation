import fs from 'node:fs';
import * as XLSXModule from 'xlsx';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const XLSX: any = (XLSXModule as any).default ?? XLSXModule;

export type ServiceFeeCategory = '服务' | '工程' | '物资';

export interface ServiceFeeResult {
  amountWan: number;
  category: ServiceFeeCategory;
  discountRate: number;
  preDiscountYuan: number;
  postDiscountYuan: number;
  platformFeeYuan: number;
  bracketLabel: string;
}

export interface ClientAgreementSeed {
  group_name: string;
  client_name: string;
  contact_name: string;
  contact_phone: string;
  contract_expiry: string;
  discount_rate: number;
  discount_label: string;
  agreement_status: string;
  notes: string;
}

/**
 * Exact scaled BigInt decimal representation: value = intVal / 10^scale
 * Prevents IEEE-754 floating-point drift (e.g. 59.485 * 150 * 0.9 = 8030.475 -> 8030.48)
 * and implements Excel's exact ROUND(x, 2) (round half away from zero).
 */
interface Dec {
  v: bigint;
  s: number;
}

function toDec(input: number | string): Dec {
  let str = typeof input === 'number' ? input.toString() : String(input).trim();
  if (!str || !Number.isFinite(Number(str))) return { v: 0n, s: 0 };
  // Handle scientific notation if any
  if (/[eE]/.test(str)) {
    str = Number(str).toFixed(12).replace(/\.?0+$/, '');
  }
  const neg = str.startsWith('-');
  if (neg || str.startsWith('+')) str = str.slice(1);
  const [intPart, fracPart = ''] = str.split('.');
  const digits = `${intPart || '0'}${fracPart}`;
  const v = BigInt(digits) * (neg ? -1n : 1n);
  return { v, s: fracPart.length };
}

function pow10(n: number): bigint {
  return 10n ** BigInt(n);
}

function addDec(a: Dec, b: Dec): Dec {
  if (a.s === b.s) return { v: a.v + b.v, s: a.s };
  if (a.s > b.s) return { v: a.v + b.v * pow10(a.s - b.s), s: a.s };
  return { v: a.v * pow10(b.s - a.s) + b.v, s: b.s };
}

function subDec(a: Dec, b: Dec): Dec {
  return addDec(a, { v: -b.v, s: b.s });
}

function mulDec(a: Dec, b: Dec): Dec {
  return { v: a.v * b.v, s: a.s + b.s };
}

function roundHalfAwayFromZero(d: Dec, places: number): number {
  if (d.s <= places) {
    const scaled = d.v * pow10(places - d.s);
    return Number(scaled) / Math.pow(10, places);
  }
  const drop = d.s - places;
  const divisor = pow10(drop);
  const absV = d.v < 0n ? -d.v : d.v;
  const quot = absV / divisor;
  const rem = absV % divisor;
  const half = divisor / 2n;
  const roundedAbs = rem >= half ? quot + 1n : quot;
  const signed = d.v < 0n ? -roundedAbs : roundedAbs;
  return Number(signed) / Math.pow(10, places);
}

function decToNumber(d: Dec): number {
  return Number(d.v) / Math.pow(10, d.s);
}

export function normalizeCategory(raw?: string): ServiceFeeCategory {
  const s = String(raw || '').trim();
  if (/物资|货物/.test(s)) return '物资';
  if (/施工|工程/.test(s)) return '工程';
  return '服务';
}

interface Bracket {
  upper: number; // in 万元, Infinity for top bracket
  lower: number;
  rate: string; // exact decimal string
  baseWan: string; // cumulative fee in 万元 at lower bound
  label: string;
}

const BRACKETS: Record<ServiceFeeCategory, Bracket[]> = {
  物资: [
    { upper: 100, lower: 0, rate: '0.015', baseWan: '0', label: '≤100万 (1.5%)' },
    { upper: 500, lower: 100, rate: '0.011', baseWan: '1.5', label: '100~500万 (1.1%)' },
    { upper: 1000, lower: 500, rate: '0.008', baseWan: '5.9', label: '500~1000万 (0.8%)' },
    { upper: 5000, lower: 1000, rate: '0.005', baseWan: '9.9', label: '1000~5000万 (0.5%)' },
    { upper: 10000, lower: 5000, rate: '0.0025', baseWan: '29.9', label: '5000万~1亿 (0.25%)' },
    { upper: 50000, lower: 10000, rate: '0.0005', baseWan: '42.4', label: '1~5亿 (0.05%)' },
    { upper: 100000, lower: 50000, rate: '0.00035', baseWan: '62.4', label: '5~10亿 (0.035%)' },
    { upper: 500000, lower: 100000, rate: '0.00008', baseWan: '79.9', label: '10~50亿 (0.008%)' },
    { upper: 1000000, lower: 500000, rate: '0.00006', baseWan: '111.9', label: '50~100亿 (0.006%)' },
    { upper: Infinity, lower: 1000000, rate: '0.00004', baseWan: '141.9', label: '>100亿 (0.004%)' },
  ],
  服务: [
    { upper: 100, lower: 0, rate: '0.015', baseWan: '0', label: '≤100万 (1.5%)' },
    { upper: 500, lower: 100, rate: '0.008', baseWan: '1.5', label: '100~500万 (0.8%)' },
    { upper: 1000, lower: 500, rate: '0.0045', baseWan: '4.7', label: '500~1000万 (0.45%)' },
    { upper: 5000, lower: 1000, rate: '0.0025', baseWan: '6.95', label: '1000~5000万 (0.25%)' },
    { upper: 10000, lower: 5000, rate: '0.001', baseWan: '16.95', label: '5000万~1亿 (0.1%)' },
    { upper: 50000, lower: 10000, rate: '0.0005', baseWan: '21.95', label: '1~5亿 (0.05%)' },
    { upper: 100000, lower: 50000, rate: '0.00035', baseWan: '41.95', label: '5~10亿 (0.035%)' },
    { upper: 500000, lower: 100000, rate: '0.00008', baseWan: '59.45', label: '10~50亿 (0.008%)' },
    { upper: 1000000, lower: 500000, rate: '0.00006', baseWan: '91.45', label: '50~100亿 (0.006%)' },
    { upper: Infinity, lower: 1000000, rate: '0.00004', baseWan: '121.45', label: '>100亿 (0.004%)' },
  ],
  工程: [
    { upper: 100, lower: 0, rate: '0.01', baseWan: '0', label: '≤100万 (1.0%)' },
    { upper: 500, lower: 100, rate: '0.007', baseWan: '1', label: '100~500万 (0.7%)' },
    { upper: 1000, lower: 500, rate: '0.0055', baseWan: '3.8', label: '500~1000万 (0.55%)' },
    { upper: 5000, lower: 1000, rate: '0.0035', baseWan: '6.55', label: '1000~5000万 (0.35%)' },
    { upper: 10000, lower: 5000, rate: '0.002', baseWan: '20.55', label: '5000万~1亿 (0.2%)' },
    { upper: 50000, lower: 10000, rate: '0.0005', baseWan: '30.55', label: '1~5亿 (0.05%)' },
    { upper: 100000, lower: 50000, rate: '0.00035', baseWan: '50.55', label: '5~10亿 (0.035%)' },
    { upper: 500000, lower: 100000, rate: '0.00008', baseWan: '68.05', label: '10~50亿 (0.008%)' },
    { upper: 1000000, lower: 500000, rate: '0.00006', baseWan: '100.05', label: '50~100亿 (0.006%)' },
    { upper: Infinity, lower: 1000000, rate: '0.00004', baseWan: '130.05', label: '>100亿 (0.004%)' },
  ],
};

/**
 * Exact Excel-equivalent agency service fee calculation.
 * Mirrors `服务费计算表格.xlsx`:
 *   L = tieredFeeWan * 10000  (unrounded)
 *   M = ROUND(L * discountRate, 2)  (Excel half-away-from-zero rounding)
 *   N = clamp(amountWan * 10, 100, 5000)  (platform usage fee, 0.1% of amount in yuan)
 */
export function calculateServiceFee(
  amountWanInput: number | string,
  categoryRaw: string = '服务',
  discountRateInput: number | string = 1
): ServiceFeeResult | null {
  const numAmount = Number(String(amountWanInput ?? '').trim());
  if (!Number.isFinite(numAmount) || numAmount <= 0) return null;

  const category = normalizeCategory(categoryRaw);
  const numDiscount = Number(String(discountRateInput ?? '1').trim());
  const discountRate = Number.isFinite(numDiscount) && numDiscount > 0 ? numDiscount : 1;

  const brackets = BRACKETS[category];
  const bracket = brackets.find((b) => numAmount <= b.upper) || brackets[brackets.length - 1];

  const decAmount = toDec(amountWanInput);
  const decDiff = subDec(decAmount, toDec(bracket.lower));
  const decFeeWan = addDec(mulDec(decDiff, toDec(bracket.rate)), toDec(bracket.baseWan));
  const decPreYuan = mulDec(decFeeWan, toDec(10000));
  const decPostYuan = mulDec(decPreYuan, toDec(discountRateInput));

  const preDiscountYuan = decToNumber(decPreYuan);
  const postDiscountYuan = roundHalfAwayFromZero(decPostYuan, 2);

  // Platform usage fee: 0.1% of award amount in yuan (= amountWan * 10), min 100, max 5000
  const rawPlatform = decToNumber(mulDec(decAmount, toDec(10)));
  const platformFeeYuan = rawPlatform <= 100 ? 100 : rawPlatform >= 5000 ? 5000 : Number(rawPlatform.toFixed(6));

  return {
    amountWan: numAmount,
    category,
    discountRate,
    preDiscountYuan,
    postDiscountYuan,
    platformFeeYuan,
    bracketLabel: bracket.label,
  };
}

export function formatDiscountLabel(rate: number): string {
  if (Math.abs(rate - 1) < 1e-6) return '不打折 (100%)';
  const pct = Number((rate * 100).toFixed(2));
  const zhe = Number((rate * 10).toFixed(2));
  return `${zhe}折 (${pct}%)`;
}

/**
 * Default demo client agreement list (real agreements can be imported via Excel in the Award module)
 */
export const DEFAULT_CLIENT_AGREEMENTS: ClientAgreementSeed[] = [
  { group_name: '主业单位', client_name: '示例省级电力有限公司', contact_name: '张联系人', contact_phone: '010-00000001', contract_expiry: '2026-12-31', discount_rate: 1, discount_label: '不打折 (100%)', agreement_status: '完成签订', notes: '示例数据，请通过「导入客户协议清单」替换' },
  { group_name: '主业单位', client_name: '示例市供电公司 A（授权采购）', contact_name: '李联系人', contact_phone: '010-00000002', contract_expiry: '2026-12-31', discount_rate: 0.9, discount_label: '9折 (90%)', agreement_status: '完成签订', notes: '' },
  { group_name: '主业单位', client_name: '示例市供电公司 B', contact_name: '王联系人', contact_phone: '010-00000003', contract_expiry: '2026-12-31', discount_rate: 0.95, discount_label: '9.5折 (95%)', agreement_status: '完成签订', notes: '' },
  { group_name: '直属单位', client_name: '示例信息通信分公司', contact_name: '赵联系人', contact_phone: '010-00000004', contract_expiry: '2026-12-31', discount_rate: 0.9, discount_label: '9折 (90%)', agreement_status: '完成签订', notes: '' },
  { group_name: '直属单位', client_name: '示例综合能源服务公司', contact_name: '周联系人', contact_phone: '010-00000005', contract_expiry: '2026-12-31', discount_rate: 0.85, discount_label: '8.5折 (85%)', agreement_status: '完成签订', notes: '' },
];

function excelSerialToDateStr(val: unknown): string {
  if (typeof val === 'number' && val > 20000 && val < 80000) {
    const utc = (val - 25569) * 86400 * 1000;
    return new Date(utc).toISOString().slice(0, 10);
  }
  return String(val ?? '').trim();
}

/**
 * Parse an uploaded 客户协议清单 (.xls / .xlsx) to update the client discount table.
 */
export function parseClientAgreementsExcel(filePath: string): ClientAgreementSeed[] {
  const wb = XLSX.read(fs.readFileSync(filePath), { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as unknown[][];

  let headerIdx = -1;
  for (let i = 0; i < Math.min(aoa.length, 10); i++) {
    const joined = aoa[i].map((c) => String(c || '').trim()).join('|');
    if (joined.includes('委托人') && joined.includes('折扣')) {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx === -1) {
    throw new Error('未识别到客户协议清单表头（需包含「委托人」和「折扣」列）');
  }

  const headers = aoa[headerIdx].map((c) => String(c || '').replace(/\s+/g, ''));
  const colIdx = (re: RegExp) => headers.findIndex((h) => re.test(h));
  const iGroup = colIdx(/组别/);
  const iClient = colIdx(/委托人|业主单位|客户/);
  const iContact = colIdx(/联系人/);
  const iPhone = colIdx(/联系方式|电话/);
  const iExpiry = colIdx(/到期/);
  const iDiscount = colIdx(/折扣/);
  const iStatus = colIdx(/签订/);
  const iNotes = colIdx(/备注/);

  const results: ClientAgreementSeed[] = [];
  let currentGroup = '主业单位';

  for (let r = headerIdx + 1; r < aoa.length; r++) {
    const row = aoa[r];
    if (iGroup >= 0 && String(row[iGroup] || '').trim()) {
      currentGroup = String(row[iGroup]).trim();
    }
    const clientName = iClient >= 0 ? String(row[iClient] || '').trim() : '';
    if (!clientName) continue;

    const contactName = iContact >= 0 ? String(row[iContact] || '').trim() : '';
    const contactPhone = iPhone >= 0 ? String(row[iPhone] || '').replace(/\s+/g, ' ').trim() : '';
    const contractExpiry = iExpiry >= 0 ? excelSerialToDateStr(row[iExpiry]) : '';
    const status = iStatus >= 0 ? String(row[iStatus] || '').trim() : '';
    const notes = iNotes >= 0 ? String(row[iNotes] || '').trim() : '';
    const rawDiscount = iDiscount >= 0 ? row[iDiscount] : 1;

    if (typeof rawDiscount === 'number' && Number.isFinite(rawDiscount)) {
      const rate = rawDiscount > 1 && rawDiscount <= 100 ? rawDiscount / 100 : rawDiscount;
      results.push({
        group_name: currentGroup,
        client_name: clientName,
        contact_name: contactName,
        contact_phone: contactPhone,
        contract_expiry: contractExpiry,
        discount_rate: rate,
        discount_label: formatDiscountLabel(rate),
        agreement_status: status,
        notes,
      });
    } else {
      const text = String(rawDiscount || '').trim();
      // Handle multi-rate entries like "授权采购83%/客户工程53%"
      const parts = text.split(/[\/、;；]/).map((p) => p.trim()).filter(Boolean);
      const parsedMulti: Array<{ tag: string; rate: number }> = [];
      for (const p of parts) {
        const m = p.match(/^([^\d]*)(\d+(?:\.\d+)?)\s*%?$/);
        if (m) {
          const num = Number(m[2]);
          const rate = num > 1 ? num / 100 : num;
          parsedMulti.push({ tag: m[1].trim(), rate });
        }
      }
      if (parsedMulti.length > 1) {
        for (const item of parsedMulti) {
          results.push({
            group_name: currentGroup,
            client_name: item.tag ? `${clientName}（${item.tag}）` : clientName,
            contact_name: contactName,
            contact_phone: contactPhone,
            contract_expiry: contractExpiry,
            discount_rate: item.rate,
            discount_label: formatDiscountLabel(item.rate),
            agreement_status: status,
            notes: notes || text,
          });
        }
      } else if (parsedMulti.length === 1) {
        results.push({
          group_name: currentGroup,
          client_name: clientName,
          contact_name: contactName,
          contact_phone: contactPhone,
          contract_expiry: contractExpiry,
          discount_rate: parsedMulti[0].rate,
          discount_label: formatDiscountLabel(parsedMulti[0].rate),
          agreement_status: status,
          notes,
        });
      } else {
        results.push({
          group_name: currentGroup,
          client_name: clientName,
          contact_name: contactName,
          contact_phone: contactPhone,
          contract_expiry: contractExpiry,
          discount_rate: 1,
          discount_label: '不打折 (100%)',
          agreement_status: status,
          notes: notes || text,
        });
      }
    }
  }

  return results;
}
