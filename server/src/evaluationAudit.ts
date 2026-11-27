import { isSameReqAppNo, stripAllSpaces } from './biddingParser.js';

export type TaxCategory = '服务' | '物资';

export const SERVICE_VALID_TAX_RATES = [0, 1, 3, 5, 6, 9, 13];
export const GOODS_VALID_TAX_RATES = [13];
export const ALL_PROOF_CANDIDATE_RATES = [0, 1, 3, 5, 6, 9, 13];

export interface TaxAuditItem {
  type: 'tax_abnormal' | 'tax_proof' | 'tax_missing';
  label: string;
  lot_code: string;
  lot_name: string;
  package_no: string;
  round_label: string;
  req_app_no: string;
  project_name: string;
  supplier_name: string;
  tax_rate_raw: string;
  detail: string;
}

export interface SameQuoteItem {
  scope?: '整包' | '子包';
  lot_code: string;
  lot_name: string;
  package_no: string;
  round_label: string;
  req_app_no: string;
  project_name: string;
  quote_basis: '含税总价(万元)' | '折扣率(%)';
  quote_value: string;
  suppliers: string[];
  detail: string;
}

export interface RoundIncreaseItem {
  scope?: '整包' | '子包';
  lot_code: string;
  lot_name: string;
  package_no: string;
  req_app_no: string;
  project_name: string;
  supplier_name: string;
  quote_basis: '含税总价(万元)' | '折扣率(%)';
  round1_value: string;
  round2_value: string;
  increase_text: string;
  detail: string;
}

export interface LowQuoteItem {
  lot_code: string;
  lot_name: string;
  package_no: string;
  project_name: string;
  supplier_name: string;
  round_label: string;
  quote_basis: '含税总价(万元)' | '折扣率(%)';
  quote_value: string;
  avg_value: string;
  second_lowest_value: string;
  max_price_value: string;
  triggered_rules: string[];
  detail: string;
}

export interface QuoteTaxAuditResult {
  taxCategory: TaxCategory;
  proofRates: number[];
  validRates: number[];
  taxAnomalies: TaxAuditItem[];
  taxProofNeeded: TaxAuditItem[];
  taxMissing: TaxAuditItem[];
  sameQuotes: SameQuoteItem[];
  roundIncreases: RoundIncreaseItem[];
  lowQuotes: LowQuoteItem[];
  counts: {
    taxAbnormal: number;
    taxProof: number;
    taxMissing: number;
    sameQuote: number;
    roundIncrease: number;
    lowQuote: number;
  };
}

export interface ContactOverlapItem {
  lot_code: string;
  lot_name: string;
  package_no: string;
  field_key: string;
  field_label: string;
  matched_value: string;
  suppliers: string[];
  alert_label: string;
}

export interface ContactMissingItem {
  lot_code: string;
  lot_name: string;
  package_no: string;
  supplier_name: string;
  detail: string;
}

export interface ContactAuditResult {
  overlaps: ContactOverlapItem[];
  missingContacts: ContactMissingItem[];
  counts: {
    overlapCount: number;
    missingCount: number;
  };
}

export interface HardwareAlertItem {
  lot_code: string;
  lot_name: string;
  package_no: string;
  supplier_a: string;
  file_a: string;
  round_a: string;
  source_a: string;
  supplier_b: string;
  file_b: string;
  round_b: string;
  source_b: string;
  cpu: string;
  disk: string;
  mac: string;
  ip_a: string;
  ip_b: string;
  detail: string;
}

export interface HardwareManualReviewItem extends HardwareAlertItem {
  mac_a: string;
  mac_b: string;
  review_reason: string;
}

export interface HardwareInsufficientItem {
  lot_code: string;
  lot_name: string;
  package_no: string;
  round_label: string;
  supplier_name: string;
  file_type: string;
  source_location: string;
  ip: string;
  mac: string;
  cpu: string;
  disk: string;
  detail: string;
}

export interface HardwareAuditResult {
  totalRecords: number;
  alerts: HardwareAlertItem[];
  manualReview: HardwareManualReviewItem[];
  insufficient: HardwareInsufficientItem[];
  counts: {
    alertCount: number;
    manualReviewCount: number;
    insufficientCount: number;
  };
}

const pinyinCollator = new Intl.Collator('zh-Hans-CN-u-co-pinyin', { numeric: true, sensitivity: 'base' });

function extractPkgNum(pkgNo: string): number {
  const m = String(pkgNo || '').match(/(\d+)/);
  return m ? parseInt(m[1], 10) : Number.MAX_SAFE_INTEGER;
}

function sortByLotPkg<T extends { lot_name?: string; lot_code?: string; package_no?: string }>(arr: T[]): T[] {
  return [...arr].sort((a, b) => {
    const lc = pinyinCollator.compare(a.lot_name || a.lot_code || '', b.lot_name || b.lot_code || '');
    if (lc !== 0) return lc;
    return extractPkgNum(a.package_no || '') - extractPkgNum(b.package_no || '');
  });
}

function getRoundNum(label: string): number {
  const s = String(label || '');
  const m = s.match(/(\d+)/);
  if (m) return parseInt(m[1], 10);
  if (/二/.test(s)) return 2;
  if (/三/.test(s)) return 3;
  return 1;
}

function parseTaxRatePercent(raw: unknown): { missing: boolean; num: number | null; rawStr: string } {
  const s = String(raw ?? '').trim();
  if (!s || s === '/' || s === '-') return { missing: true, num: null, rawStr: s || '/' };
  const cleaned = s.replace(/[%％\s]/g, '');
  const n = Number(cleaned);
  if (!Number.isFinite(n)) return { missing: false, num: null, rawStr: s };
  // If stored as decimal like 0.06 -> 6% (note: 0 stays 0, 1 stays 1% because Excel tax rate column stores 1 for 1% and 0.06 for 6%)
  const pct = n > 0 && n < 1 && !s.includes('%') && !s.includes('％') ? Number((n * 100).toFixed(4)) : Number(n.toFixed(4));
  return { missing: false, num: pct, rawStr: s };
}

function parseDiscountNumeric(raw: unknown): number | null {
  const s = String(raw ?? '').trim();
  if (!s || s === '/' || s === '-') return null;
  const hasPct = s.endsWith('%') || s.endsWith('％');
  const n = Number(s.replace(/[%％,\s]/g, ''));
  if (!Number.isFinite(n) || n <= 0) return null;
  if (!hasPct && n <= 1.5) return Number((n * 100).toFixed(4));
  return Number(n.toFixed(4));
}

export function parseProofRatesCsv(raw: string | undefined | null): number[] {
  if (raw === undefined || raw === null) return [0, 1];
  const s = String(raw).trim();
  if (s === '') return [];
  return Array.from(
    new Set(
      s
        .split(/[,，\s]+/)
        .map((v) => Number(v.trim()))
        .filter((n) => Number.isFinite(n) && n >= 0)
    )
  ).sort((a, b) => a - b);
}

/**
 * 1. 报价与税率核验：
 * - 税率异常 / 税率证明待核实 / 税率缺失：基于《行信息报价基础数据表》
 * - 相同报价 / 二轮超一轮：
 *   整包各轮报价以《开标一览表》为主（未传时用《行信息报价基础数据表》兜底）；
 *   合并子包各轮报价以《行信息报价基础数据表》逐子包核验。
 */
export function runQuoteTaxAudit(params: {
  bids: Array<Record<string, unknown>>;
  lineQuotes: Array<Record<string, unknown>>;
  projects: Array<Record<string, unknown>>;
  taxCategory: TaxCategory;
  proofRates: number[];
}): QuoteTaxAuditResult {
  const { bids, lineQuotes, projects, taxCategory, proofRates } = params;
  const validRates = taxCategory === '物资' ? GOODS_VALID_TAX_RATES : SERVICE_VALID_TAX_RATES;
  const validSet = new Set(validRates);
  const proofSet = new Set(proofRates);

  // Lookup lot_name, project_name, quote_method, and sub-package count per (lot_code, package_no)
  const lotNameMap = new Map<string, string>();
  const projNameByReq = new Map<string, string>();
  const pkgMetaMap = new Map<string, {
    lot_name: string;
    project_names: string[];
    req_app_nos: string[];
    is_discount: boolean;
    sub_count: number;
  }>();

  for (const p of projects) {
    const lc = String(p.lot_code || '').trim();
    const ln = String(p.lot_name || '').trim();
    const pkgNo = String(p.package_no || '').trim();
    const rNo = stripAllSpaces(p.req_app_no);
    const pName = String(p.name || '').trim();
    const qMethod = String(p.quote_method || '').trim();
    const maxRaw = String(p.max_price_raw || '').trim();
    const isDisc = /折扣|下浮|费率|百分/.test(qMethod) || /%$/.test(maxRaw);
    if (lc && ln && !lotNameMap.has(lc)) lotNameMap.set(lc, ln);
    if (rNo && pName && !projNameByReq.has(rNo)) projNameByReq.set(rNo, pName);
    if (lc || pkgNo) {
      const pk = `${lc}::${pkgNo}`;
      let meta = pkgMetaMap.get(pk);
      if (!meta) {
        meta = { lot_name: ln || lc, project_names: [], req_app_nos: [], is_discount: isDisc, sub_count: 0 };
        pkgMetaMap.set(pk, meta);
      }
      meta.sub_count += 1;
      if (isDisc) meta.is_discount = true;
      if (pName && !meta.project_names.includes(pName)) meta.project_names.push(pName);
      if (rNo && !meta.req_app_nos.includes(rNo)) meta.req_app_nos.push(rNo);
    }
  }
  const resolveProjName = (reqAppNo: string, fallback: string) => {
    const cleanReq = stripAllSpaces(reqAppNo);
    if (cleanReq) {
      const exact = projNameByReq.get(cleanReq);
      if (exact) return exact;
      for (const [k, v] of projNameByReq.entries()) {
        if (isSameReqAppNo(k, cleanReq)) return v;
      }
    }
    return fallback;
  };

  const taxAnomalies: TaxAuditItem[] = [];
  const taxProofNeeded: TaxAuditItem[] = [];
  const taxMissing: TaxAuditItem[] = [];

  // Grouping for Same-Quote check: key = scope::lot_code::package_no::round_label::full_req_app_no::basis::numeric_val
  const sameQuoteGroups = new Map<string, {
    scope: '整包' | '子包';
    lot_code: string;
    lot_name: string;
    package_no: string;
    round_label: string;
    req_app_no: string;
    project_name: string;
    quote_basis: '含税总价(万元)' | '折扣率(%)';
    quote_value: string;
    suppliers: Set<string>;
  }>();

  // Grouping for Round 1 vs Round 2 increase check: key = scope::lot_code::package_no::full_req_app_no::supplier_name::basis
  const roundPairMap = new Map<string, {
    scope: '整包' | '子包';
    lot_code: string;
    lot_name: string;
    package_no: string;
    req_app_no: string;
    project_name: string;
    supplier_name: string;
    quote_basis: '含税总价(万元)' | '折扣率(%)';
    r1?: { num: number; display: string };
    r2?: { num: number; display: string };
  }>();

  const recordQuoteForAudit = (entry: {
    scope: '整包' | '子包';
    lotCode: string;
    lotName: string;
    pkgNo: string;
    roundLabel: string;
    roundNum: number;
    reqAppNo: string;
    projName: string;
    supplier: string;
    basis: '含税总价(万元)' | '折扣率(%)';
    num: number;
    display: string;
  }) => {
    const sqKey = `${entry.scope}::${entry.lotCode}::${entry.pkgNo}::${entry.roundLabel}::${entry.reqAppNo}::${entry.basis}::${entry.num}`;
    let grp = sameQuoteGroups.get(sqKey);
    if (!grp) {
      grp = {
        scope: entry.scope,
        lot_code: entry.lotCode,
        lot_name: entry.lotName,
        package_no: entry.pkgNo,
        round_label: entry.roundLabel,
        req_app_no: entry.reqAppNo,
        project_name: entry.projName,
        quote_basis: entry.basis,
        quote_value: entry.display,
        suppliers: new Set(),
      };
      sameQuoteGroups.set(sqKey, grp);
    }
    grp.suppliers.add(entry.supplier);

    if (entry.roundNum === 1 || entry.roundNum === 2) {
      const rpKey = `${entry.scope}::${entry.lotCode}::${entry.pkgNo}::${entry.reqAppNo}::${entry.supplier}::${entry.basis}`;
      let rp = roundPairMap.get(rpKey);
      if (!rp) {
        rp = {
          scope: entry.scope,
          lot_code: entry.lotCode,
          lot_name: entry.lotName,
          package_no: entry.pkgNo,
          req_app_no: entry.reqAppNo,
          project_name: entry.projName,
          supplier_name: entry.supplier,
          quote_basis: entry.basis,
        };
        roundPairMap.set(rpKey, rp);
      }
      if (entry.roundNum === 1) rp.r1 = { num: entry.num, display: entry.display };
      if (entry.roundNum === 2) rp.r2 = { num: entry.num, display: entry.display };
    }
  };

  // Track which (lot_code, package_no) exist in 开标一览表 and which suppliers actually submitted Round 2 in 开标一览表
  const bidOpeningPkgs = new Set<string>();
  const bidOpeningRound2Sups = new Set<string>();

  for (const b of bids) {
    const lotCode = String(b.lot_code || '').trim();
    const pkgNo = String(b.package_no || '').trim();
    const supplier = String(b.supplier_name || '').trim();
    if (!supplier) continue;
    const pk = `${lotCode}::${pkgNo}`;
    bidOpeningPkgs.add(pk);

    const roundLabel = String(b.round_label || '第1轮').trim();
    const roundNum = getRoundNum(roundLabel);
    if (roundNum >= 2) bidOpeningRound2Sups.add(`${pk}::${supplier}`);

    const meta = pkgMetaMap.get(pk);
    const lotName = String(b.lot_name || '').trim() || meta?.lot_name || lotNameMap.get(lotCode) || lotCode;
    const projName = meta?.project_names.join('；') || '';
    const reqAppNo = meta?.req_app_nos.join(' / ') || '';
    const isDiscountMode = meta?.is_discount || String(b.bid_price_raw || '').trim().endsWith('%');

    const rawPriceStr = String(b.bid_price_raw || '').trim();
    const rawWan = b.bid_price_wan !== null && b.bid_price_wan !== undefined ? Number(b.bid_price_wan) : null;
    const discStr = String(b.discount_rate || '').trim();

    if (isDiscountMode) {
      const discNum = parseDiscountNumeric(rawPriceStr && rawPriceStr !== '/' && rawPriceStr !== '-' ? rawPriceStr : discStr);
      if (discNum !== null) {
        recordQuoteForAudit({
          scope: '整包',
          lotCode, lotName, pkgNo, roundLabel, roundNum, reqAppNo, projName, supplier,
          basis: '折扣率(%)',
          num: discNum,
          display: `${discNum}%`,
        });
      }
    } else {
      let numVal: number | null = null;
      if (rawWan !== null && Number.isFinite(rawWan) && rawWan > 0) {
        numVal = Number(rawWan.toFixed(6));
      } else if (rawPriceStr && rawPriceStr !== '/' && rawPriceStr !== '-') {
        const parsed = Number(rawPriceStr.replace(/万元|万|元|,|，|\s/g, ''));
        if (Number.isFinite(parsed) && parsed > 0) numVal = Number(parsed.toFixed(6));
      }
      if (numVal !== null) {
        recordQuoteForAudit({
          scope: '整包',
          lotCode, lotName, pkgNo, roundLabel, roundNum, reqAppNo, projName, supplier,
          basis: '含税总价(万元)',
          num: numVal,
          display: `${numVal}`,
        });
      }
    }
  }

  // Group lineQuotes by (lotCode, pkgNo, supplier, roundLabel) for: (a) fallback whole-package audit when 开标一览表 missing, (b) merged sub-package audit
  const linePkgSupRounds = new Map<string, Map<string, {
    round_label: string;
    round_num: number;
    inc_sum: number | null;
    discount_rate: string;
    rows: Array<Record<string, unknown>>;
  }>>();

  for (const lq of lineQuotes) {
    const lotCode = String(lq.lot_code || '').trim();
    const lotName = lotNameMap.get(lotCode) || lotCode;
    const pkgNo = String(lq.package_no || '').trim();
    const roundLabel = String(lq.round_label || '第1轮').trim();
    const roundNum = getRoundNum(roundLabel);
    const reqAppNo = stripAllSpaces(lq.req_app_no);
    const projName = resolveProjName(reqAppNo, String(lq.project_name || '').trim());
    const supplier = String(lq.supplier_name || '').trim();
    if (!supplier) continue;

    // A. Tax checks (always run from 行信息报价基础数据表 regardless of whether ceiling matched)
    const taxInfo = parseTaxRatePercent(lq.tax_rate);
    if (taxInfo.missing) {
      taxMissing.push({
        type: 'tax_missing',
        label: '税率缺失',
        lot_code: lotCode,
        lot_name: lotName,
        package_no: pkgNo,
        round_label: roundLabel,
        req_app_no: reqAppNo,
        project_name: projName,
        supplier_name: supplier,
        tax_rate_raw: taxInfo.rawStr,
        detail: '税率为空或填写“/”，需核查响应文件',
      });
    } else if (taxInfo.num === null || !validSet.has(taxInfo.num)) {
      taxAnomalies.push({
        type: 'tax_abnormal',
        label: '税率异常',
        lot_code: lotCode,
        lot_name: lotName,
        package_no: pkgNo,
        round_label: roundLabel,
        req_app_no: reqAppNo,
        project_name: projName,
        supplier_name: supplier,
        tax_rate_raw: taxInfo.rawStr,
        detail: `${taxCategory}类合法税率为 ${validRates.join('、')}%，当前填报“${taxInfo.rawStr}”`,
      });
    } else if (proofSet.has(taxInfo.num)) {
      taxProofNeeded.push({
        type: 'tax_proof',
        label: '税率证明待核实',
        lot_code: lotCode,
        lot_name: lotName,
        package_no: pkgNo,
        round_label: roundLabel,
        req_app_no: reqAppNo,
        project_name: projName,
        supplier_name: supplier,
        tax_rate_raw: `${taxInfo.num}%`,
        detail: `命中独立核查税率（${taxInfo.num}%），请人工核查投标文件中的免税/小规模纳税人证明`,
      });
    }

    const psKey = `${lotCode}::${pkgNo}::${supplier}`;
    let rMap = linePkgSupRounds.get(psKey);
    if (!rMap) {
      rMap = new Map();
      linePkgSupRounds.set(psKey, rMap);
    }
    let rGrp = rMap.get(roundLabel);
    if (!rGrp) {
      rGrp = { round_label: roundLabel, round_num: roundNum, inc_sum: null, discount_rate: '', rows: [] };
      rMap.set(roundLabel, rGrp);
    }
    rGrp.rows.push(lq);
    const incTax = lq.price_inc_tax_wan !== null && lq.price_inc_tax_wan !== undefined ? Number(lq.price_inc_tax_wan) : null;
    if (incTax !== null && Number.isFinite(incTax) && incTax > 0) {
      rGrp.inc_sum = Number(((rGrp.inc_sum ?? 0) + incTax).toFixed(6));
    }
    const disc = String(lq.discount_rate || '').trim();
    if (disc && disc !== '/' && !rGrp.discount_rate) rGrp.discount_rate = disc;
  }

  // B. Process lineQuotes for: (1) fallback whole-package if not in 开标一览表; (2) per-sub-package if merged sub-packages
  for (const [psKey, rMap] of linePkgSupRounds.entries()) {
    const [lotCode, pkgNo, supplier] = psKey.split('::');
    const pk = `${lotCode}::${pkgNo}`;
    const meta = pkgMetaMap.get(pk);
    const lotName = meta?.lot_name || lotNameMap.get(lotCode) || lotCode;
    const hasBidOpening = bidOpeningPkgs.has(pk);
    const hasBidOpeningR2 = bidOpeningRound2Sups.has(`${pk}::${supplier}`);
    const isMergedSubPkg = (meta?.sub_count || 0) > 1;

    const r1Grp = Array.from(rMap.values()).find((g) => g.round_num === 1);

    for (const grp of rMap.values()) {
      // Detect ECP/ETP auto-carried Round 2+ rows when 开标一览表 has no Round 2 for this supplier
      if (grp.round_num >= 2 && !hasBidOpeningR2 && r1Grp) {
        const sumSame =
          (grp.inc_sum === null && r1Grp.inc_sum === null) ||
          (grp.inc_sum !== null && r1Grp.inc_sum !== null && Math.abs(grp.inc_sum - r1Grp.inc_sum) < 1e-6);
        const discSame = (grp.discount_rate || '') === (r1Grp.discount_rate || '');
        if (sumSame && discSame) {
          continue;
        }
      }

      // Fallback whole-package audit when 开标一览表 wasn't uploaded for this package
      if (!hasBidOpening) {
        const projName = meta?.project_names.join('；') || String(grp.rows[0]?.project_name || '').trim();
        const reqAppNo = meta?.req_app_nos.join(' / ') || stripAllSpaces(grp.rows[0]?.req_app_no);
        if (grp.inc_sum !== null && grp.inc_sum > 0) {
          recordQuoteForAudit({
            scope: '整包',
            lotCode, lotName, pkgNo, roundLabel: grp.round_label, roundNum: grp.round_num, reqAppNo, projName, supplier,
            basis: '含税总价(万元)',
            num: grp.inc_sum,
            display: `${grp.inc_sum}`,
          });
        }
        const discNum = parseDiscountNumeric(grp.discount_rate);
        if (discNum !== null) {
          recordQuoteForAudit({
            scope: '整包',
            lotCode, lotName, pkgNo, roundLabel: grp.round_label, roundNum: grp.round_num, reqAppNo, projName, supplier,
            basis: '折扣率(%)',
            num: discNum,
            display: `${discNum}%`,
          });
        }
      }

      // Merged sub-package audit: check each sub-package individually from 行信息报价基础数据表
      if (isMergedSubPkg) {
        for (const lq of grp.rows) {
          const reqAppNo = stripAllSpaces(lq.req_app_no);
          const projName = resolveProjName(reqAppNo, String(lq.project_name || '').trim());
          const incTax = lq.price_inc_tax_wan !== null && lq.price_inc_tax_wan !== undefined ? Number(lq.price_inc_tax_wan) : null;
          if (incTax !== null && Number.isFinite(incTax) && incTax > 0) {
            const normNum = Number(incTax.toFixed(6));
            recordQuoteForAudit({
              scope: '子包',
              lotCode, lotName, pkgNo, roundLabel: grp.round_label, roundNum: grp.round_num, reqAppNo, projName, supplier,
              basis: '含税总价(万元)',
              num: normNum,
              display: `${normNum}`,
            });
          }
          const discNum = parseDiscountNumeric(lq.discount_rate);
          if (discNum !== null) {
            recordQuoteForAudit({
              scope: '子包',
              lotCode, lotName, pkgNo, roundLabel: grp.round_label, roundNum: grp.round_num, reqAppNo, projName, supplier,
              basis: '折扣率(%)',
              num: discNum,
              display: `${discNum}%`,
            });
          }
        }
      }
    }
  }

  const sameQuotes: SameQuoteItem[] = [];
  for (const grp of sameQuoteGroups.values()) {
    if (grp.suppliers.size >= 2) {
      const supList = Array.from(grp.suppliers);
      sameQuotes.push({
        scope: grp.scope,
        lot_code: grp.lot_code,
        lot_name: grp.lot_name,
        package_no: grp.package_no,
        round_label: grp.round_label,
        req_app_no: grp.req_app_no,
        project_name: grp.project_name,
        quote_basis: grp.quote_basis,
        quote_value: grp.quote_value,
        suppliers: supList,
        detail: `[${grp.scope}] ${supList.length} 家供应商在 ${grp.round_label} 的${grp.quote_basis}完全相同（${grp.quote_value}）：${supList.join('、')}`,
      });
    }
  }

  const roundIncreases: RoundIncreaseItem[] = [];
  for (const rp of roundPairMap.values()) {
    if (rp.r1 && rp.r2 && rp.r2.num > rp.r1.num + 1e-6) {
      const diff = Number((rp.r2.num - rp.r1.num).toFixed(4));
      const incText = rp.quote_basis === '折扣率(%)' ? `+${diff}%` : `+${diff}万元`;
      roundIncreases.push({
        scope: rp.scope,
        lot_code: rp.lot_code,
        lot_name: rp.lot_name,
        package_no: rp.package_no,
        req_app_no: rp.req_app_no,
        project_name: rp.project_name,
        supplier_name: rp.supplier_name,
        quote_basis: rp.quote_basis,
        round1_value: rp.r1.display,
        round2_value: rp.r2.display,
        increase_text: incText,
        detail: `[${rp.scope}] 二轮超一轮（否决项）：第1轮 ${rp.r1.display} → 第2轮 ${rp.r2.display}（上涨 ${incText}）`,
      });
    }
  }

  return {
    taxCategory,
    proofRates,
    validRates,
    taxAnomalies: sortByLotPkg(taxAnomalies),
    taxProofNeeded: sortByLotPkg(taxProofNeeded),
    taxMissing: sortByLotPkg(taxMissing),
    sameQuotes: sortByLotPkg(sameQuotes),
    roundIncreases: sortByLotPkg(roundIncreases),
    lowQuotes: [],
    counts: {
      taxAbnormal: taxAnomalies.length,
      taxProof: taxProofNeeded.length,
      taxMissing: taxMissing.length,
      sameQuote: sameQuotes.length,
      roundIncrease: roundIncreases.length,
      lowQuote: 0,
    },
  };
}

/**
 * 1B. 报价过低（低于成本）核查：
 * 对每个标包内通过初步评审的应答人（未否决），判定是否满足以下任一情形：
 * 1. 包应答报价 < 全部通过初步评审应答人应答报价平均值 × 50%
 * 2. 包应答报价 < 通过初步评审的次低报价应答人应答报价 × 50%
 * 3. 应答报价 < 最高限价 × 45%
 */
export function runLowQuoteAudit(packages: Array<{
  lot_code: string;
  lot_name: string;
  package_no: string;
  project_name: string;
  quote_method: string;
  max_price_yuan: number | null;
  max_price_raw: string;
  supplier_bids: Array<{
    supplier_name: string;
    final_price: string;
    final_round: string;
    is_rejected: boolean;
  }>;
}>): LowQuoteItem[] {
  const results: LowQuoteItem[] = [];

  const fmtVal = (n: number, isDiscount: boolean) =>
    isDiscount ? `${Number(n.toFixed(4))}%` : `${Number(n.toFixed(4))}万元`;

  for (const pkg of packages) {
    const isDiscountMode =
      /折扣|下浮|费率|百分/.test(pkg.quote_method || '') || /%$/.test(String(pkg.max_price_raw || '').trim());
    const quoteBasis: '含税总价(万元)' | '折扣率(%)' = isDiscountMode ? '折扣率(%)' : '含税总价(万元)';

    const parseQuoteNum = (raw: string): number | null => {
      if (!raw || raw === '/' || raw === '-') return null;
      if (isDiscountMode) return parseDiscountNumeric(raw);
      const n = Number(String(raw).replace(/万元|万|元|,|，|\s/g, ''));
      return Number.isFinite(n) && n > 0 ? n : null;
    };

    const passingBids = (pkg.supplier_bids || [])
      .filter((sb) => !sb.is_rejected)
      .map((sb) => ({
        supplier_name: sb.supplier_name,
        round_label: sb.final_round || '第1轮',
        num: parseQuoteNum(sb.final_price),
      }))
      .filter((x): x is { supplier_name: string; round_label: string; num: number } => x.num !== null);

    if (passingBids.length === 0) continue;

    const avgNum = passingBids.reduce((s, b) => s + b.num, 0) / passingBids.length;
    const sortedAsc = [...passingBids].sort((a, b) => a.num - b.num);
    const secondLowestNum = sortedAsc.length >= 2 ? sortedAsc[1].num : null;
    const maxLimitNum = isDiscountMode
      ? parseDiscountNumeric(pkg.max_price_raw)
      : pkg.max_price_yuan !== null && pkg.max_price_yuan > 0
        ? pkg.max_price_yuan / 10000
        : null;

    for (const bid of passingBids) {
      const triggered: string[] = [];
      const details: string[] = [];

      // 情形 1：低于全部通过初步评审应答人应答报价平均值 50%
      if (passingBids.length >= 2) {
        const t1 = avgNum * 0.5;
        if (bid.num < t1 - 1e-6) {
          triggered.push('低于通过初评平均值50%');
          details.push(`报价 ${fmtVal(bid.num, isDiscountMode)} < 通过初评均值50%（${fmtVal(t1, isDiscountMode)}，均值 ${fmtVal(avgNum, isDiscountMode)}）`);
        }
      }

      // 情形 2：低于通过初步评审的次低报价应答人应答报价 50%
      if (secondLowestNum !== null) {
        const t2 = secondLowestNum * 0.5;
        if (bid.num < t2 - 1e-6) {
          triggered.push('低于次低报价50%');
          details.push(`报价 ${fmtVal(bid.num, isDiscountMode)} < 通过初评次低价50%（${fmtVal(t2, isDiscountMode)}，次低价 ${fmtVal(secondLowestNum, isDiscountMode)}）`);
        }
      }

      // 情形 3：低于最高限价 45%
      if (maxLimitNum !== null && maxLimitNum > 0) {
        const t3 = maxLimitNum * 0.45;
        if (bid.num < t3 - 1e-6) {
          triggered.push('低于最高限价45%');
          details.push(`报价 ${fmtVal(bid.num, isDiscountMode)} < 最高限价45%（${fmtVal(t3, isDiscountMode)}，限价 ${fmtVal(maxLimitNum, isDiscountMode)}）`);
        }
      }

      if (triggered.length > 0) {
        results.push({
          lot_code: pkg.lot_code,
          lot_name: pkg.lot_name || pkg.lot_code,
          package_no: pkg.package_no,
          project_name: pkg.project_name,
          supplier_name: bid.supplier_name,
          round_label: bid.round_label,
          quote_basis: quoteBasis,
          quote_value: fmtVal(bid.num, isDiscountMode),
          avg_value: fmtVal(avgNum, isDiscountMode),
          second_lowest_value: secondLowestNum !== null ? fmtVal(secondLowestNum, isDiscountMode) : '—',
          max_price_value: maxLimitNum !== null ? fmtVal(maxLimitNum, isDiscountMode) : '—',
          triggered_rules: triggered,
          detail: `疑似低于成本：${details.join('；')}`,
        });
      }
    }
  }

  return sortByLotPkg(results);
}

/**
 * 2. 同标包供应商联系信息围串标核验（开标一览表决定名单，供应商联系人信息表提供字段）
 */
export function runContactCollusionAudit(params: {
  bids: Array<Record<string, unknown>>;
  suppliers: Array<Record<string, unknown>>;
  projects: Array<Record<string, unknown>>;
}): ContactAuditResult {
  const { bids, suppliers, projects } = params;

  // Multi-record supplier contact resolver: preserves ALL contact rows per supplier,
  // and NEVER silently borrows another lot's contact record when the current lot/package has no match.
  const cleanKey = (v: unknown) => String(v || '').replace(/[\s\u3000]+/g, '').trim();
  const normPkgToken = (pkg: string) => {
    const s = cleanKey(pkg);
    if (!s) return '';
    const m = s.match(/^(?:标包|包)?(\d+)$/);
    return m ? `包${parseInt(m[1], 10)}` : s;
  };
  const pushUnique = (map: Map<string, Array<Record<string, unknown>>>, k: string, rec: Record<string, unknown>) => {
    let list = map.get(k);
    if (!list) {
      list = [];
      map.set(k, list);
    }
    if (!list.includes(rec)) list.push(rec);
  };

  const byLotPkg = new Map<string, Array<Record<string, unknown>>>();
  const byPkgNoLot = new Map<string, Array<Record<string, unknown>>>();
  const byGeneric = new Map<string, Array<Record<string, unknown>>>();
  const hasAnySupplierRecord = new Set<string>();

  for (const s of suppliers) {
    const sk = cleanKey(s.supplier_name);
    if (!sk) continue;
    hasAnySupplierRecord.add(sk);

    const rawPkg = String(s.package_no || '').trim();
    if (!rawPkg) {
      pushUnique(byGeneric, sk, s);
      continue;
    }
    const segments = rawPkg.split(/[;；\r\n]+/).map((seg) => seg.trim()).filter(Boolean);
    for (const seg of segments) {
      const colonIdx = seg.search(/[:：]/);
      if (colonIdx >= 0) {
        const lotPart = cleanKey(seg.slice(0, colonIdx));
        const pkgsPart = seg.slice(colonIdx + 1);
        const tokens = pkgsPart.split(/[,，、\s]+/).map(normPkgToken).filter(Boolean);
        for (const tok of tokens) {
          if (lotPart) {
            pushUnique(byLotPkg, `${sk}::${lotPart}::${tok}`, s);
          } else {
            pushUnique(byPkgNoLot, `${sk}::${tok}`, s);
          }
        }
      } else {
        const tokens = seg.split(/[,，、\s]+/).map(normPkgToken).filter(Boolean);
        for (const tok of tokens) {
          pushUnique(byPkgNoLot, `${sk}::${tok}`, s);
        }
      }
    }
  }

  const resolveContacts = (supplierName: string, lotName: string, packageNo: string): Array<Record<string, unknown>> => {
    const sk = cleanKey(supplierName);
    if (!sk) return [];
    const normLot = cleanKey(lotName);
    const normPkg = normPkgToken(packageNo);
    if (normLot && normPkg) {
      const exact = byLotPkg.get(`${sk}::${normLot}::${normPkg}`);
      if (exact && exact.length > 0) return exact;
    }
    if (normPkg) {
      const byPkg = byPkgNoLot.get(`${sk}::${normPkg}`);
      if (byPkg && byPkg.length > 0) return byPkg;
    }
    // Only fall back to generic contact rows (package_no is blank, e.g. ETP exports), never cross-lot records
    return byGeneric.get(sk) || [];
  };

  // Build bidirectional lot_code <-> lot_name lookup so rows missing lot_code still merge with same-lot rows
  const lotNameMap = new Map<string, string>();
  const lotNameToCodes = new Map<string, Set<string>>();
  const registerLotPair = (lc: string, ln: string) => {
    if (!lc || !ln) return;
    if (!lotNameMap.has(lc)) lotNameMap.set(lc, ln);
    let set = lotNameToCodes.get(ln);
    if (!set) {
      set = new Set();
      lotNameToCodes.set(ln, set);
    }
    set.add(lc);
  };
  for (const p of projects) {
    registerLotPair(String(p.lot_code || '').trim(), String(p.lot_name || '').trim());
  }
  for (const b of bids) {
    registerLotPair(String(b.lot_code || '').trim(), String(b.lot_name || '').trim());
  }
  const uniqueLotCodeByName = new Map<string, string>();
  for (const [ln, codes] of lotNameToCodes.entries()) {
    if (codes.size === 1) uniqueLotCodeByName.set(ln, Array.from(codes)[0]);
  }

  // Group unique suppliers per (lotKey, package_no) from 开标一览表
  const pkgSuppliers = new Map<string, {
    lot_code: string;
    lot_name: string;
    package_no: string;
    suppliers: Set<string>;
  }>();
  const ambiguousMissingSeen = new Set<string>();
  const ambiguousMissingList: ContactMissingItem[] = [];

  for (const b of bids) {
    const rawLotCode = String(b.lot_code || '').trim();
    const rawLotName = String(b.lot_name || '').trim();
    const isAmbiguousLotName = !rawLotCode && Boolean(rawLotName) && (lotNameToCodes.get(rawLotName)?.size || 0) > 1;
    const lotCode = rawLotCode || (rawLotName ? uniqueLotCodeByName.get(rawLotName) || '' : '');
    const lotName = rawLotName || lotNameMap.get(lotCode) || lotCode;
    const pkgNo = String(b.package_no || '').trim();
    const supName = String(b.supplier_name || '').trim();
    const lotKey = lotCode || (lotName ? `NAME:${lotName}` : '');
    if (!supName || !lotKey || !pkgNo) continue;
    if (isAmbiguousLotName) {
      const ambKey = `${rawLotName}::${pkgNo}::${supName}`;
      if (!ambiguousMissingSeen.has(ambKey)) {
        ambiguousMissingSeen.add(ambKey);
        ambiguousMissingList.push({
          lot_code: '',
          lot_name: rawLotName,
          package_no: pkgNo,
          supplier_name: supName,
          detail: '同名分标存在多个分标编号且本行缺少分标编号，无法自动归组比对，请人工核实',
        });
      }
      continue;
    }
    const k = `${lotKey}::${pkgNo}`;
    let entry = pkgSuppliers.get(k);
    if (!entry) {
      entry = { lot_code: lotCode, lot_name: lotName, package_no: pkgNo, suppliers: new Set() };
      pkgSuppliers.set(k, entry);
    }
    entry.suppliers.add(supName);
  }

  const CONTACT_FIELDS: Array<{
    key: string;
    label: string;
    alertPrefix: string;
    isAddress?: boolean;
    isPhone?: boolean;
    isEmail?: boolean;
  }> = [
    { key: 'contact_name', label: '联系人姓名', alertPrefix: '投标人联系人姓名相同' },
    { key: 'mobile', label: '手机号码', alertPrefix: '手机号码相同', isPhone: true },
    { key: 'phone', label: '固定电话', alertPrefix: '固定电话相同', isPhone: true },
    { key: 'email', label: '电子邮件', alertPrefix: '电子邮件相同', isEmail: true },
    { key: 'fax', label: '传真号码', alertPrefix: '传真号码相同', isPhone: true },
    { key: 'address', label: '邮政地址', alertPrefix: '邮政地址相同', isAddress: true },
  ];

  const PLACEHOLDER_RE = /^(?:[\/／\-－—–]+|无|无邮箱|无传真|无电话|暂无|null|none|n\/a|na)$/i;
  const overlaps: ContactOverlapItem[] = [];
  const missingContacts: ContactMissingItem[] = [...ambiguousMissingList];

  for (const pkg of pkgSuppliers.values()) {
    const supRecords: Array<{ name: string; recs: Array<Record<string, unknown>> }> = [];
    for (const supName of pkg.suppliers) {
      const recs = resolveContacts(supName, pkg.lot_name, pkg.package_no);
      if (recs.length === 0) {
        const sk = cleanKey(supName);
        missingContacts.push({
          lot_code: pkg.lot_code,
          lot_name: pkg.lot_name,
          package_no: pkg.package_no,
          supplier_name: supName,
          detail: hasAnySupplierRecord.has(sk)
            ? '本标包无联系人记录（仅在其他分标/包存在联系信息，未跨标借用，请人工核实）'
            : '未找到供应商联系人信息',
        });
      } else {
        supRecords.push({ name: supName, recs });
      }
    }

    if (supRecords.length < 2) continue;

    for (const f of CONTACT_FIELDS) {
      const valMap = new Map<string, { suppliers: Set<string>; rawSamples: Set<string> }>();
      for (const { name, recs } of supRecords) {
        for (const rec of recs) {
          const rawVal = String(rec[f.key] ?? '').trim();
          if (!rawVal || PLACEHOLDER_RE.test(rawVal)) continue;
          // 邮政地址需至少包含一个数字或“号”字才参与比较
          if (f.isAddress && !/[\d号]/.test(rawVal)) continue;
          let normKey = rawVal;
          if (f.isPhone) {
            // 号码类字段剥离非数字字符后比对，不足 7 位视为无效短号/占位值跳过
            const digits = rawVal.replace(/\D/g, '');
            if (digits.length < 7) continue;
            normKey = digits;
          } else if (f.isEmail) {
            // 邮箱需包含 @ 且域名部分包含 .，统一转小写比对
            const lower = rawVal.toLowerCase();
            const atIdx = lower.lastIndexOf('@');
            if (atIdx <= 0 || !lower.slice(atIdx + 1).includes('.')) continue;
            normKey = lower;
          }
          let entry = valMap.get(normKey);
          if (!entry) {
            entry = { suppliers: new Set(), rawSamples: new Set() };
            valMap.set(normKey, entry);
          }
          entry.suppliers.add(name);
          entry.rawSamples.add(rawVal);
        }
      }

      for (const [val, entry] of valMap.entries()) {
        if (entry.suppliers.size >= 2) {
          overlaps.push({
            lot_code: pkg.lot_code,
            lot_name: pkg.lot_name,
            package_no: pkg.package_no,
            field_key: f.key,
            field_label: f.label,
            matched_value: f.isPhone ? Array.from(entry.rawSamples).join(' / ') : val,
            suppliers: Array.from(entry.suppliers),
            alert_label: f.alertPrefix,
          });
        }
      }
    }
  }

  return {
    overlaps: sortByLotPkg(overlaps),
    missingContacts: sortByLotPkg(missingContacts),
    counts: {
      overlapCount: overlaps.length,
      missingCount: missingContacts.length,
    },
  };
}

/**
 * 3. 《投标文件IP地址一览表》硬件特征围串标核验
 * 规则：
 * - 同一分标、同一包内，不同投标人的技术/商务/价格/资格业绩四类文件跨类别、跨轮次两两比较（跳过同投标人）
 * - 先比 CPU 序列号和硬盘序列号，两项都相同再看网卡 MAC 地址；仅当 CPU + 硬盘 + MAC 三项全部一致才报「围串标预警」（IP 相同不单独触发）
 * - 某类文件四字段全空时直接跳过；只有 IP 或部分设备信息但缺 CPU / 硬盘任意一项，列为「资料不足，未完成核验」
 * - 硬件字段按整格文本严格比较（不做大小写或分隔符归一化）
 */
export function runHardwareCollusionAudit(ipRecords: Array<Record<string, unknown>>): HardwareAuditResult {
  interface FileEntry {
    supplier_name: string;
    round_label: string;
    file_type: string;
    ip: string;
    mac: string;
    cpu: string;
    disk: string;
    source_line: string;
  }

  const FILE_SPECS = [
    { label: '技术文件', ip: 'tech_ip', mac: 'tech_mac', cpu: 'tech_cpu', disk: 'tech_disk' },
    { label: '商务文件', ip: 'biz_ip', mac: 'biz_mac', cpu: 'biz_cpu', disk: 'biz_disk' },
    { label: '价格文件', ip: 'price_ip', mac: 'price_mac', cpu: 'price_cpu', disk: 'price_disk' },
    { label: '资格业绩文件', ip: 'qual_ip', mac: 'qual_mac', cpu: 'qual_cpu', disk: 'qual_disk' },
  ] as const;

  const pkgGroups = new Map<string, {
    lot_code: string;
    lot_name: string;
    package_no: string;
    entries: FileEntry[];
  }>();

  const insufficient: HardwareInsufficientItem[] = [];

  // 统一清洗半角/全角占位符（/ ／ - － — —— 无 null N/A）为空串；非占位符保留原样精确比对
  const PLACEHOLDER_RE = /^(?:[\/／\-－—–]+|无|null|n\/a|na)$/i;
  const exactVal = (v: unknown) => {
    const s = String(v ?? '').trim();
    return !s || PLACEHOLDER_RE.test(s) ? '' : s;
  };

  const formatSourceLine = (fileType: string, roundLabel: string, sourceFile: string, sheetName: string, rowNum: number) => {
    const posPart = sheetName ? `${sheetName} 第${rowNum}行` : `第${rowNum}行`;
    const locParts = [sourceFile, posPart].filter(Boolean).join(' / ');
    return `${fileType} / ${roundLabel} / ${locParts}`;
  };

  for (let rowIdx = 0; rowIdx < ipRecords.length; rowIdx++) {
    const r = ipRecords[rowIdx];
    const fallbackSeq = rowIdx + 1;
    const sourceFile = String(r.source_file ?? '').trim();
    const sheetName = String(r.sheet_name ?? '').trim();
    const rawSourceRow = r.source_row !== null && r.source_row !== undefined ? Number(r.source_row) : NaN;
    const excelRowNum = Number.isFinite(rawSourceRow) && rawSourceRow > 0 ? Math.trunc(rawSourceRow) : fallbackSeq;

    const lotCode = exactVal(r.lot_code);
    const lotName = exactVal(r.lot_name) || lotCode;
    const pkgNo = exactVal(r.package_no);
    const roundLabel = String(r.round_label || '第1轮').trim();
    const supplier = String(r.supplier_name || '').trim();
    if (!supplier) continue;

    // 分标定位键：优先分标编号，编号为空时用分标名称兜底防串标；分标与包号任一缺失则不参与跨供应商比对
    const lotKey = lotCode || (lotName ? `NAME:${lotName}` : '');
    const hasValidLocation = Boolean(lotKey && pkgNo);

    let pkg: { lot_code: string; lot_name: string; package_no: string; entries: FileEntry[] } | undefined;
    if (hasValidLocation) {
      const k = `${lotKey}::${pkgNo}`;
      pkg = pkgGroups.get(k);
      if (!pkg) {
        pkg = { lot_code: lotCode, lot_name: lotName, package_no: pkgNo, entries: [] };
        pkgGroups.set(k, pkg);
      }
    }

    for (const fs of FILE_SPECS) {
      const ip = exactVal(r[fs.ip]);
      const mac = exactVal(r[fs.mac]);
      const cpu = exactVal(r[fs.cpu]);
      const disk = exactVal(r[fs.disk]);

      // 四字段全空直接跳过
      if (!ip && !mac && !cpu && !disk) continue;

      const sourceLine = formatSourceLine(fs.label, roundLabel, sourceFile, sheetName, excelRowNum);

      // 标包定位信息不全（缺分标编号/名称或包号）-> 资料不足，禁止跨标段误分组
      if (!hasValidLocation || !pkg) {
        const missingLoc = [!lotKey ? '分标编号/分标名称' : '', !pkgNo ? '包号' : ''].filter(Boolean).join('、');
        insufficient.push({
          lot_code: lotCode,
          lot_name: lotName,
          package_no: pkgNo,
          round_label: roundLabel,
          supplier_name: supplier,
          file_type: fs.label,
          source_location: sourceLine,
          ip,
          mac,
          cpu,
          disk,
          detail: `资料不足，未完成核验（标包定位信息不全，缺少${missingLoc}）`,
        });
        continue;
      }

      // 缺 CPU 或硬盘任意一项 -> 资料不足，未完成核验
      if (!cpu || !disk) {
        const missingParts = [!cpu ? 'CPU序列号' : '', !disk ? '硬盘序列号' : ''].filter(Boolean).join('、');
        insufficient.push({
          lot_code: lotCode,
          lot_name: lotName,
          package_no: pkgNo,
          round_label: roundLabel,
          supplier_name: supplier,
          file_type: fs.label,
          source_location: sourceLine,
          ip,
          mac,
          cpu,
          disk,
          detail: `资料不足，未完成核验（缺少${missingParts}）`,
        });
        continue;
      }

      pkg.entries.push({
        supplier_name: supplier,
        round_label: roundLabel,
        file_type: fs.label,
        ip,
        mac,
        cpu,
        disk,
        source_line: sourceLine,
      });
    }
  }

  interface SupplierMacGroup {
    supplier_name: string;
    mac: string;
    cpu: string;
    disk: string;
    sources: Array<{ round_label: string; file_type: string; ip: string; source_line: string }>;
  }

  const summarizeSources = (sources: SupplierMacGroup['sources']) => {
    const rounds = Array.from(new Set(sources.map((s) => s.round_label).filter(Boolean)));
    const fileTypes = Array.from(new Set(sources.map((s) => s.file_type).filter(Boolean)));
    const ips = Array.from(new Set(sources.map((s) => s.ip).filter(Boolean)));
    const uniqueLines = Array.from(new Set(sources.map((s) => s.source_line).filter(Boolean)));
    return {
      round: rounds.join('、') || '第1轮',
      file: fileTypes.join('、'),
      sourceText: uniqueLines.join('\n'),
      ip: ips.join('、'),
    };
  };

  const alerts: HardwareAlertItem[] = [];
  const manualReview: HardwareManualReviewItem[] = [];

  for (const pkg of pkgGroups.values()) {
    const list = pkg.entries;
    if (list.length < 2) continue;

    // 1. 先按精确 (CPU, 硬盘) 分桶
    const cpuDiskBucket = new Map<string, FileEntry[]>();
    for (const e of list) {
      const cdKey = `${e.cpu}\u0000${e.disk}`;
      let arr = cpuDiskBucket.get(cdKey);
      if (!arr) {
        arr = [];
        cpuDiskBucket.set(cdKey, arr);
      }
      arr.push(e);
    }

    for (const group of cpuDiskBucket.values()) {
      if (group.length < 2) continue;

      // 2. 在 (CPU, 硬盘) 组内按 (投标人, MAC) 归集来源，彻底消除行序依赖与早退遮挡
      const supMacMap = new Map<string, SupplierMacGroup>();
      for (const e of group) {
        const smKey = `${e.supplier_name}\u0000${e.mac}`;
        let sm = supMacMap.get(smKey);
        if (!sm) {
          sm = {
            supplier_name: e.supplier_name,
            mac: e.mac,
            cpu: e.cpu,
            disk: e.disk,
            sources: [],
          };
          supMacMap.set(smKey, sm);
        }
        sm.sources.push({
          round_label: e.round_label,
          file_type: e.file_type,
          ip: e.ip,
          source_line: e.source_line,
        });
      }

      const smList = Array.from(supMacMap.values());
      for (let i = 0; i < smList.length; i++) {
        for (let j = i + 1; j < smList.length; j++) {
          const a = smList[i];
          const b = smList[j];
          if (a.supplier_name === b.supplier_name) continue;
          const [first, second] = a.supplier_name < b.supplier_name ? [a, b] : [b, a];
          const firstMeta = summarizeSources(first.sources);
          const secondMeta = summarizeSources(second.sources);
          const firstInline = firstMeta.sourceText.replace(/\n/g, '；');
          const secondInline = secondMeta.sourceText.replace(/\n/g, '；');

          // 2A. 双方均有 MAC：一致才报围串标预警，明确不同则跳过（且不遮挡该供应商的其他 MAC 记录）
          if (first.mac && second.mac) {
            if (first.mac !== second.mac) continue;
            alerts.push({
              lot_code: pkg.lot_code,
              lot_name: pkg.lot_name,
              package_no: pkg.package_no,
              supplier_a: first.supplier_name,
              file_a: firstMeta.file,
              round_a: firstMeta.round,
              source_a: firstMeta.sourceText,
              supplier_b: second.supplier_name,
              file_b: secondMeta.file,
              round_b: secondMeta.round,
              source_b: secondMeta.sourceText,
              cpu: first.cpu,
              disk: first.disk,
              mac: first.mac,
              ip_a: firstMeta.ip,
              ip_b: secondMeta.ip,
              detail: `围串标预警：${first.supplier_name}（${firstInline}）与 ${second.supplier_name}（${secondInline}）的 CPU、硬盘序列号及网卡 MAC 地址完全一致`,
            });
            continue;
          }

          // 2B. CPU + 硬盘一致，单方或双方未提供 MAC -> 待人工核实（与同供应商其他完整命中记录并存，互不遮挡）
          const bothMissingMac = !first.mac && !second.mac;
          const reviewReason = bothMissingMac
            ? 'CPU、硬盘一致，双方未提供 MAC'
            : '待人工核实（CPU、硬盘一致，单方未提供 MAC）';
          manualReview.push({
            lot_code: pkg.lot_code,
            lot_name: pkg.lot_name,
            package_no: pkg.package_no,
            supplier_a: first.supplier_name,
            file_a: firstMeta.file,
            round_a: firstMeta.round,
            source_a: firstMeta.sourceText,
            supplier_b: second.supplier_name,
            file_b: secondMeta.file,
            round_b: secondMeta.round,
            source_b: secondMeta.sourceText,
            cpu: first.cpu,
            disk: first.disk,
            mac: bothMissingMac ? '双方均未提供' : `${first.mac || '未提供'} / ${second.mac || '未提供'}`,
            mac_a: first.mac,
            mac_b: second.mac,
            ip_a: firstMeta.ip,
            ip_b: secondMeta.ip,
            review_reason: reviewReason,
            detail: `${reviewReason}：${first.supplier_name}（${firstInline}）与 ${second.supplier_name}（${secondInline}）`,
          });
        }
      }
    }
  }

  return {
    totalRecords: ipRecords.length,
    alerts: sortByLotPkg(alerts),
    manualReview: sortByLotPkg(manualReview),
    insufficient: sortByLotPkg(insufficient),
    counts: {
      alertCount: alerts.length,
      manualReviewCount: manualReview.length,
      insufficientCount: insufficient.length,
    },
  };
}
