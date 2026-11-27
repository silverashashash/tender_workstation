import fs from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';
import { Router } from 'express';
import multer from 'multer';
import { v4 as uuid } from 'uuid';
import { isSameReqAppNo, parseSupplierContacts } from '../biddingParser.js';
import { db, UPLOADS_DIR } from '../db.js';
import {
  parseProofRatesCsv,
  runContactCollusionAudit,
  runHardwareCollusionAudit,
  runLowQuoteAudit,
  runQuoteTaxAudit,
  type TaxCategory,
} from '../evaluationAudit.js';
import {
  normalizeBatchCode,
  parseBidOpeningSheet,
  parseIpHardwareSheet,
  parseLineQuotesSheet,
  parseRankingsSheet,
  parseRejectionsSheet,
} from '../evaluationParser.js';
import { parseSpreadsheetTables } from '../excelParser.js';
import {
  discoverObjectiveMethods,
  expandScoreFiles,
  runExpertObjectiveAudit,
  runExpertTotalAudit,
  type ExpertObjectiveAuditResult,
  type ExpertTotalAuditResult,
  type ObjectiveScoringMethod,
  type ObjectiveSelection,
  type ScoreUploadedFile,
} from '../expertScoreAudit.js';
import {
  discoverPrecheckMetadata,
  runExpertPrecheckAudit,
  type ExpertPrecheckResult,
  type PrecheckObjectiveSelection,
  type PrecheckRejectionRecord,
  type PrecheckSheetMeta,
} from '../expertPrecheckAudit.js';
import {
  buildInitialReportDocxBuffer,
  DEFAULT_COMPANY_SUFFIXES,
  generateInitialReportData,
  type InitialReportResult,
} from '../initialReportGenerator.js';
import {
  buildWarningFormsZipBuffer,
  renderWarningFormDocx,
  runWarningFormAudit,
  type WarningFormAuditResult,
} from '../warningFormGenerator.js';
import {
  buildMinutesResult,
  formatDefaultMeetingTime,
  parseMinutesExcel,
  renderInitialMinutesDocx,
  type InitialMinutesConfig,
  type InitialMinutesResult,
  type MinutesProcurementType,
} from '../initialMinutesGenerator.js';
import { toolboxToolsRouter } from './toolboxTools.js';

export const evaluationRouter = Router();
evaluationRouter.use(toolboxToolsRouter);

export const EVAL_SLOTS = [
  {
    key: 'bid_opening',
    order: 1,
    label: '开标一览表',
    description: '提取各标段响应家数与每轮报价（非公开支持多轮，公开招标为单轮报价）',
  },
  {
    key: 'line_quotes',
    order: 2,
    label: '行信息报价基础数据表',
    description: '关联采购申请号，提取含税/未含税总价、增值税率、同价与二轮涨价核验',
  },
  {
    key: 'ip_hardware',
    order: 3,
    label: '投标文件IP地址一览表',
    description: '核验同标包不同投标人技术/商务/价格/资格文件的 CPU、硬盘与 MAC 硬件特征',
  },
  {
    key: 'rejections',
    order: 4,
    label: '未进入详评情况统计表',
    description: '记录初评否决原因、文件条款依据与商务/技术分类，供公示与报告引用',
  },
  {
    key: 'rankings',
    order: 5,
    label: '总分排序表（最终版）',
    description: '多区块解析技术/商务/价格/总分、候选人排序、基准价并自动识别流标包',
  },
] as const;

const storage = multer.diskStorage({
  destination: (req, _file, cb) => {
    const batchId = String(req.params.batchId || 'common');
    const dir = path.join(UPLOADS_DIR, batchId, 'evaluation');
    fs.mkdirSync(dir, { recursive: true });
    cb(null, dir);
  },
  filename: (_req, file, cb) => {
    const ext = path.extname(file.originalname);
    cb(null, `${Date.now()}-${uuid().slice(0, 8)}${ext}`);
  },
});

const upload = multer({ storage, limits: { fileSize: 50 * 1024 * 1024 } });

function decodeName(raw: string): string {
  try {
    return Buffer.from(raw, 'latin1').toString('utf8');
  } catch {
    return raw;
  }
}

function normSupplierKey(name: unknown): string {
  return String(name ?? '')
    .replace(/[\s\u3000\u00A0]+/g, '')
    .replace(/（/g, '(')
    .replace(/）/g, ')')
    .trim();
}

function getRoundOrder(label: string): number {
  const s = String(label || '');
  if (/最终|终轮/.test(s)) return 99;
  const m = s.match(/(\d+)/);
  if (m) return parseInt(m[1], 10);
  if (/二/.test(s)) return 2;
  if (/三/.test(s)) return 3;
  return 1;
}

/**
 * Normalize percentage/discount-rate strings to percentage points (0..100+).
 * Examples: "100%" -> 100, "89.50%" -> 89.5, "88" -> 88, "0.95" -> 95.
 */
function parsePercentPoints(val: unknown): number | null {
  const s = String(val ?? '').trim();
  if (!s || s === '/' || s === '-') return null;
  const hasPct = s.endsWith('%');
  const num = Number(hasPct ? s.slice(0, -1).trim() : s);
  if (!Number.isFinite(num) || num <= 0) return null;
  if (!hasPct && num <= 1.5) return Number((num * 100).toFixed(4));
  return num;
}

/**
 * Check whether a supplier's final quote exceeds the package's max_price ceiling,
 * cleanly distinguishing 折扣率报价 (%) vs 总价报价 (万元).
 */
function checkOverLimit(
  finalQuote: string,
  quoteMethod: string,
  maxPriceYuan: number | null,
  maxPriceRaw: string
): { isOver: boolean; msg: string; limitDisplay: string; quoteDisplay: string; exceedText: string } {
  if (!finalQuote || finalQuote === '/' || finalQuote === '-') {
    return { isOver: false, msg: '', limitDisplay: '', quoteDisplay: '', exceedText: '' };
  }
  const isDiscountMode = /折扣|下浮|费率|百分/.test(quoteMethod) || /%$/.test(String(maxPriceRaw || '').trim());

  if (isDiscountMode) {
    const limitPct = parsePercentPoints(maxPriceRaw);
    const quotePct = parsePercentPoints(finalQuote);
    if (limitPct !== null && quotePct !== null && quotePct > limitPct + 1e-6) {
      const diff = Number((quotePct - limitPct).toFixed(4));
      return {
        isOver: true,
        msg: `超限价（报价 ${quotePct}% > 限价 ${limitPct}%）`,
        limitDisplay: `${limitPct}%`,
        quoteDisplay: `${quotePct}%`,
        exceedText: `+${diff}%`,
      };
    }
    return { isOver: false, msg: '', limitDisplay: '', quoteDisplay: '', exceedText: '' };
  }

  // 总价报价（万元）
  if (maxPriceYuan !== null && maxPriceYuan > 0) {
    const limitWan = maxPriceYuan / 10000;
    const quoteWan = Number(String(finalQuote).replace(/万元|万|元/g, '').trim());
    if (Number.isFinite(quoteWan) && quoteWan > 0 && quoteWan > limitWan + 1e-6) {
      const limStr = String(Number(limitWan.toFixed(4)));
      const diffWan = Number((quoteWan - limitWan).toFixed(4));
      const diffPct = Number((((quoteWan - limitWan) / limitWan) * 100).toFixed(2));
      return {
        isOver: true,
        msg: `超限价（报价 ${quoteWan} 万元 > 限价 ${limStr} 万元）`,
        limitDisplay: `${limStr} 万元`,
        quoteDisplay: `${quoteWan} 万元`,
        exceedText: `+${diffWan}万元 (+${diffPct}%)`,
      };
    }
  }
  return { isOver: false, msg: '', limitDisplay: '', quoteDisplay: '', exceedText: '' };
}

const evalPinyinCollator = new Intl.Collator('zh-Hans-CN-u-co-pinyin', { numeric: true, sensitivity: 'base' });

function extractEvalPkgNum(pkgNo: string): number {
  const m = String(pkgNo || '').match(/(\d+)/);
  return m ? parseInt(m[1], 10) : Number.MAX_SAFE_INTEGER;
}

function buildEvaluationWorkspaceData(batchId: string) {
  const batch = db.prepare('SELECT * FROM batches WHERE id = ?').get(batchId) as Record<string, unknown> | undefined;
  if (!batch) return null;

  const projects = db.prepare('SELECT * FROM batch_projects WHERE batch_id = ? ORDER BY seq ASC').all(batchId) as Array<Record<string, unknown>>;
  const sopItems = db.prepare('SELECT * FROM sop_items WHERE batch_id = ? AND is_template = 0 ORDER BY item_type DESC, sort_order ASC').all(batchId);
  const files = db.prepare("SELECT * FROM uploaded_files WHERE batch_id = ? AND module = 'evaluation' ORDER BY created_at DESC").all(batchId) as Array<Record<string, unknown>>;

  const bids = db.prepare('SELECT * FROM eval_bids WHERE batch_id = ?').all(batchId) as Array<Record<string, unknown>>;
  const lineQuotes = db.prepare('SELECT * FROM eval_line_quotes WHERE batch_id = ?').all(batchId) as Array<Record<string, unknown>>;
  const rejections = db.prepare('SELECT * FROM eval_rejections WHERE batch_id = ?').all(batchId) as Array<Record<string, unknown>>;
  const rankings = db.prepare('SELECT * FROM eval_rankings WHERE batch_id = ? ORDER BY rank_no ASC').all(batchId) as Array<Record<string, unknown>>;
  const suppliers = db.prepare('SELECT * FROM batch_suppliers WHERE batch_id = ? ORDER BY seq ASC').all(batchId) as Array<Record<string, unknown>>;
  const ipRecords = db.prepare('SELECT * FROM eval_ip_records WHERE batch_id = ?').all(batchId) as Array<Record<string, unknown>>;

  // 对历史已上传供应商联系人信息表、但尚未入库 fax 列的批次自动回填一次
  if (suppliers.length > 0 && suppliers.every((s) => !String(s.fax || '').trim())) {
    const supFile = db
      .prepare("SELECT stored_path FROM uploaded_files WHERE batch_id = ? AND module = 'bidding' AND slot_key = 'supplier_contacts' ORDER BY created_at DESC LIMIT 1")
      .get(batchId) as { stored_path?: string } | undefined;
    if (supFile?.stored_path) {
      const absSup = path.join(UPLOADS_DIR, supFile.stored_path);
      if (fs.existsSync(absSup)) {
        try {
          const parsedSup = parseSupplierContacts(absSup);
          const faxBySeq = new Map<number, string>();
          parsedSup.suppliers.forEach((ps) => {
            if (ps.fax) faxBySeq.set(ps.seq, ps.fax);
          });
          if (faxBySeq.size > 0) {
            const updFax = db.prepare('UPDATE batch_suppliers SET fax = ? WHERE id = ?');
            for (const s of suppliers) {
              const fx = faxBySeq.get(Number(s.seq));
              if (fx) {
                s.fax = fx;
                updFax.run(fx, String(s.id));
              }
            }
          }
        } catch {
          // ignore non-fatal backfill
        }
      }
    }
  }

  const formattedFiles = files.map((f) => ({
    ...f,
    slot_key: String(f.slot_key || ''),
    parsed: f.parsed_json ? JSON.parse(String(f.parsed_json)) : null,
  }));

  const slots = EVAL_SLOTS.map((s) => ({
    ...s,
    file: formattedFiles.find((f) => f.slot_key === s.key) || null,
  }));

  const otherFiles = formattedFiles.filter((f) => !f.slot_key);

  // Build unified per-package evaluation summary keyed by lot_code + package_no (or fallback to package_no)
  const pkgKey = (lotCode: string, pkgNo: string) => `${lotCode || ''}::${pkgNo || ''}`;
  interface SubPkgDef {
    req_app_no: string;
    name: string;
    max_price_yuan: number | null;
    max_price_raw: string;
    quote_method: string;
  }
  interface OverLimitDetail {
    scope: '整包' | '子包';
    round_label: string;
    req_app_no: string;
    sub_name: string;
    limit_display: string;
    quote_display: string;
    exceed_text: string;
    msg: string;
  }
  interface SubQuoteItem {
    req_app_no: string;
    sub_name: string;
    round_label: string;
    price_inc_tax_wan: number | null;
    discount_rate: string;
    limit_display: string;
    is_over_limit: boolean;
    over_limit_msg: string;
  }

  const pkgMap = new Map<string, {
    key: string;
    lot_code: string;
    lot_name: string;
    package_no: string;
    req_app_no: string;
    project_name: string;
    max_price_yuan: number | null;
    max_price_raw: string;
    quote_method: string;
    sub_packages: SubPkgDef[];
    over_limit_count: number;
    bidder_count: number;
    rejection_count: number;
    detailed_eval_count: number;
    is_failed_lot: boolean;
    failed_reason: string;
    eval_date: string;
    first_candidate: string;
    first_price_wan: string;
    first_score: number | null;
    discount_pct: string;
    bid_price_summary: string;
    supplier_bids: Array<{
      supplier_name: string;
      round1_price: string;
      round2_price: string;
      final_price: string;
      final_round: string;
      tax_rate: string;
      price_inc_tax_wan: number | null;
      price_ex_tax_wan: number | null;
      discount_rate: string;
      is_rejected: boolean;
      rejection_reason: string;
      is_over_limit: boolean;
      over_limit_msg: string;
      over_limit_details: OverLimitDetail[];
      sub_quotes: SubQuoteItem[];
    }>;
    candidates: Array<Record<string, unknown>>;
    bid_rounds: Array<Record<string, unknown>>;
    line_quotes: Array<Record<string, unknown>>;
    rejections: Array<Record<string, unknown>>;
  }>();

  const ensurePkg = (lotCode: string, lotName: string, pkgNo: string, projName = '', reqAppNo = '', maxYuan: number | null = null, maxRaw = '', quoteMethod = '') => {
    const k = pkgKey(lotCode, pkgNo);
    let entry = pkgMap.get(k);
    if (!entry) {
      entry = {
        key: k,
        lot_code: lotCode,
        lot_name: lotName,
        package_no: pkgNo,
        req_app_no: reqAppNo,
        project_name: projName,
        max_price_yuan: maxYuan,
        max_price_raw: maxRaw,
        quote_method: quoteMethod,
        sub_packages: [],
        over_limit_count: 0,
        bidder_count: 0,
        rejection_count: 0,
        detailed_eval_count: 0,
        is_failed_lot: false,
        failed_reason: '',
        eval_date: '',
        first_candidate: '',
        first_price_wan: '',
        first_score: null,
        discount_pct: '',
        bid_price_summary: '',
        supplier_bids: [],
        candidates: [],
        bid_rounds: [],
        line_quotes: [],
        rejections: [],
      };
      pkgMap.set(k, entry);
    } else {
      if (!entry.lot_name && lotName) entry.lot_name = lotName;
      if (!entry.project_name && projName) entry.project_name = projName;
      if (!entry.req_app_no && reqAppNo) entry.req_app_no = reqAppNo;
      if (entry.max_price_yuan === null && maxYuan !== null) entry.max_price_yuan = maxYuan;
    }
    return entry;
  };

  // Seed from Bidding stage projects first (preserve individual sub-package ceilings & sum whole-package ceiling)
  projects.forEach((p) => {
    if (p.lot_code || p.package_no) {
      const lotCode = String(p.lot_code || '');
      const pkgNo = String(p.package_no || '');
      const k = pkgKey(lotCode, pkgNo);
      const existed = pkgMap.has(k);
      const pMaxYuan =
        p.max_price_yuan !== null && p.max_price_yuan !== undefined
          ? Number(p.max_price_yuan)
          : p.budget_yuan !== null && p.budget_yuan !== undefined
            ? Number(p.budget_yuan)
            : null;
      const pMaxRaw = String(p.max_price_raw || '');
      const pQuoteMethod = String(p.quote_method || '');
      const pName = String(p.name || '');
      const reqNo = String(p.req_app_no || '');
      const entry = ensurePkg(
        lotCode,
        String(p.lot_name || ''),
        pkgNo,
        pName,
        reqNo,
        pMaxYuan,
        pMaxRaw,
        pQuoteMethod
      );
      entry.sub_packages.push({
        req_app_no: reqNo,
        name: pName,
        max_price_yuan: pMaxYuan,
        max_price_raw: pMaxRaw,
        quote_method: pQuoteMethod,
      });
      if (existed) {
        if (pMaxYuan !== null) {
          entry.max_price_yuan = (entry.max_price_yuan || 0) + pMaxYuan;
        }
        if (pName && !entry.project_name.includes(pName)) {
          entry.project_name = entry.project_name ? `${entry.project_name}；${pName}` : pName;
        }
        if (reqNo && !entry.req_app_no.includes(reqNo)) {
          entry.req_app_no = entry.req_app_no ? `${entry.req_app_no} / ${reqNo}` : reqNo;
        }
      }
    }
  });

  // Attach bids (deduplicated strictly by normalized supplier name per lot+package)
  const bidderSets = new Map<string, Set<string>>();
  bids.forEach((b) => {
    const entry = ensurePkg(String(b.lot_code || ''), String(b.lot_name || ''), String(b.package_no || ''));
    entry.bid_rounds.push(b);
    const sk = normSupplierKey(b.supplier_name);
    if (sk) {
      if (!bidderSets.has(entry.key)) bidderSets.set(entry.key, new Set());
      bidderSets.get(entry.key)!.add(sk);
    }
  });
  for (const [k, set] of bidderSets.entries()) {
    const entry = pkgMap.get(k);
    if (entry) entry.bidder_count = set.size;
  }

  // Attach line quotes (VAT tax rates & req_app_no)
  lineQuotes.forEach((lq) => {
    const entry = ensurePkg(String(lq.lot_code || ''), '', String(lq.package_no || ''), String(lq.project_name || ''), String(lq.req_app_no || ''));
    entry.line_quotes.push(lq);
  });

  // Attach rejections (match by lot_code + package_no, with lot_name + package_no fallback; deduplicate same supplier across biz/tech reviews)
  const rejSupplierSets = new Map<string, Set<string>>();
  rejections.forEach((rj) => {
    const lotCode = String(rj.lot_code || '');
    const lotName = String(rj.lot_name || '');
    const pkgNo = String(rj.package_no || '');
    let entry = lotCode ? pkgMap.get(pkgKey(lotCode, pkgNo)) : undefined;
    if (!entry && lotName && pkgNo) {
      for (const candidate of pkgMap.values()) {
        if (candidate.lot_name === lotName && candidate.package_no === pkgNo) {
          entry = candidate;
          break;
        }
      }
    }
    if (!entry) {
      entry = ensurePkg(lotCode, lotName, pkgNo);
    }
    entry.rejections.push(rj);
    const sk = normSupplierKey(rj.supplier_name);
    if (sk) {
      if (!rejSupplierSets.has(entry.key)) rejSupplierSets.set(entry.key, new Set());
      rejSupplierSets.get(entry.key)!.add(sk);
    }
    entry.rejection_count = rejSupplierSets.get(entry.key)?.size || 0;
  });

  // Attach rankings & compute winner + discount vs max_price (distinguishing 折扣率报价 vs 总价报价)
  rankings.forEach((rk) => {
    const entry = ensurePkg(String(rk.lot_code || ''), String(rk.lot_name || ''), String(rk.package_no || ''), String(rk.project_name || ''));
    if (rk.eval_date && !entry.eval_date) entry.eval_date = String(rk.eval_date);
    const isDiscountMode =
      /折扣|下浮|费率|百分/.test(entry.quote_method) || /%$/.test(String(entry.max_price_raw || '').trim());

    const fmtMaybePct = (raw: unknown): string => {
      const s = String(raw ?? '').trim();
      if (!s || s === '/' || s === '-') return s;
      if (!isDiscountMode || s.endsWith('%')) return s;
      const n = Number(s);
      if (Number.isFinite(n) && n > 0) return `${n}%`;
      return s;
    };

    if (Number(rk.is_failed_lot) === 1) {
      entry.is_failed_lot = true;
    } else {
      entry.candidates.push(
        isDiscountMode
          ? {
              ...rk,
              base_price: fmtMaybePct(rk.base_price),
              bid_price: fmtMaybePct(rk.bid_price),
              eval_price: fmtMaybePct(rk.eval_price),
            }
          : rk
      );
      if (Number(rk.rank_no) === 1) {
        entry.first_candidate = String(rk.supplier_name || '');
        const rawFirstPrice = String(rk.eval_price || rk.bid_price || '').trim();
        entry.first_price_wan = fmtMaybePct(rawFirstPrice);
        entry.first_score = rk.total_score !== null ? Number(rk.total_score) : null;

        if (isDiscountMode) {
          const limitPct = parsePercentPoints(entry.max_price_raw);
          const quotePct = parsePercentPoints(rawFirstPrice);
          if (limitPct !== null && quotePct !== null && limitPct > 0) {
            const pct = ((limitPct - quotePct) / limitPct) * 100;
            entry.discount_pct = `${pct >= 0 ? '下浮 ' : '超限 '}${Math.abs(pct).toFixed(2)}%`;
          }
        } else {
          const evalWan = Number(rawFirstPrice);
          if (entry.max_price_yuan && Number.isFinite(evalWan) && evalWan > 0) {
            const maxWan = entry.max_price_yuan / 10000;
            const pct = ((maxWan - evalWan) / maxWan) * 100;
            entry.discount_pct = `${pct >= 0 ? '下浮 ' : '超限 '}${Math.abs(pct).toFixed(2)}%`;
          }
        }
      }
    }
  });

  // Helper: extract numeric round order so 第2轮 > 第1轮 (always treat latest round as final quote)
  const getRoundOrder = (label: string): number => {
    const s = String(label || '');
    if (/最终|终轮/.test(s)) return 99;
    const m = s.match(/(\d+)/);
    if (m) return parseInt(m[1], 10);
    if (/二/.test(s)) return 2;
    if (/三/.test(s)) return 3;
    return 1;
  };

  const pickQuoteStr = (raw: unknown, num: unknown, disc: unknown): string => {
    const r = String(raw || '').trim();
    if (r && r !== '/' && r !== '-') return r;
    if (num !== null && num !== undefined && Number.isFinite(Number(num))) return String(num);
    const d = String(disc || '').trim();
    if (d && d !== '/' && d !== '-') return d;
    return '';
  };

  const isSameQuoteVal = (a: string, b: string): boolean => {
    const sa = String(a || '').trim();
    const sb = String(b || '').trim();
    if (!sa || !sb) return false;
    if (sa === sb) return true;
    const na = Number(sa.replace(/[%万元万,，\s]/g, ''));
    const nb = Number(sb.replace(/[%万元万,，\s]/g, ''));
    if (Number.isFinite(na) && Number.isFinite(nb)) {
      return Math.abs(na - nb) < 1e-6;
    }
    return false;
  };

  // Compute remaining detailed evaluation count, early failed-lot detection, and unified per-supplier bids (2nd round = final)
  // 公开招标：进入详评至少 3 家；非公开招标：进入详评至少 2 家
  const isPublicBidding = String(batch.procurement_type || '') === '公开招标';
  const minDetailedEval = isPublicBidding ? 3 : 2;
  for (const entry of pkgMap.values()) {
    entry.detailed_eval_count = Math.max(0, entry.bidder_count - entry.rejection_count);
    if (entry.is_failed_lot) {
      entry.failed_reason =
        entry.bidder_count > 0 && entry.detailed_eval_count < minDetailedEval
          ? `详评不足${minDetailedEval}家`
          : '排序表流标';
    } else if (!entry.first_candidate && entry.bidder_count > 0 && entry.detailed_eval_count < minDetailedEval) {
      entry.is_failed_lot = true;
      entry.failed_reason = `详评不足${minDetailedEval}家`;
    }

    // Merge bid_rounds (开标一览表) + line_quotes (行信息报价表) per supplier, preferring latest/2nd round as final quote
    const rejMap = new Map<string, string>();
    for (const rj of entry.rejections) {
      const name = String(rj.supplier_name || '').trim();
      if (name && !rejMap.has(name)) {
        rejMap.set(name, String(rj.reason_category || rj.review_type || '否决'));
      }
    }

    const isMergedSubPkg = entry.sub_packages.length > 1;

    const supMap = new Map<string, {
      supplier_name: string;
      round1_price: string;
      round2_price: string;
      final_price: string;
      final_round: string;
      final_round_order: number;
      tax_rate: string;
      price_inc_tax_wan: number | null;
      price_ex_tax_wan: number | null;
      discount_rate: string;
      line_round_order: number;
      whole_rounds: Map<string, string>; // round_label -> whole-package quote string
      line_by_round: Map<string, {
        round_label: string;
        round_order: number;
        tax_rate: string;
        discount_rate: string;
        inc_sum: number | null;
        ex_sum: number | null;
        rows: Array<Record<string, unknown>>;
      }>;
    }>();

    const getSup = (name: string) => {
      let s = supMap.get(name);
      if (!s) {
        s = {
          supplier_name: name,
          round1_price: '',
          round2_price: '',
          final_price: '',
          final_round: '',
          final_round_order: 0,
          tax_rate: '',
          price_inc_tax_wan: null,
          price_ex_tax_wan: null,
          discount_rate: '',
          line_round_order: 0,
          whole_rounds: new Map(),
          line_by_round: new Map(),
        };
        supMap.set(name, s);
      }
      return s;
    };

    for (const b of entry.bid_rounds) {
      const name = String(b.supplier_name || '').trim();
      if (!name) continue;
      const s = getSup(name);
      const rLabel = String(b.round_label || '第1轮');
      const rOrd = getRoundOrder(rLabel);
      const qStr = pickQuoteStr(b.bid_price_raw, b.bid_price_wan, b.discount_rate);
      if (qStr) s.whole_rounds.set(rLabel, qStr);
      if (rOrd === 1 && qStr) s.round1_price = qStr;
      if (rOrd >= 2 && qStr) s.round2_price = qStr;
      if (qStr && rOrd >= s.final_round_order) {
        s.final_price = qStr;
        s.final_round = rLabel;
        s.final_round_order = rOrd;
      }
      const dStr = String(b.discount_rate || '').trim();
      if (dStr && dStr !== '/' && !s.discount_rate) s.discount_rate = dStr;
    }

    // Group line_quotes by (supplier, round_label) so multi-sub-package rows are summed into whole-package round totals
    for (const lq of entry.line_quotes) {
      const name = String(lq.supplier_name || '').trim();
      if (!name) continue;
      const s = getSup(name);
      const rLabel = String(lq.round_label || '第1轮');
      const rOrd = getRoundOrder(rLabel);
      let grp = s.line_by_round.get(rLabel);
      if (!grp) {
        grp = {
          round_label: rLabel,
          round_order: rOrd,
          tax_rate: '',
          discount_rate: '',
          inc_sum: null,
          ex_sum: null,
          rows: [],
        };
        s.line_by_round.set(rLabel, grp);
      }
      grp.rows.push(lq);
      const incTax = lq.price_inc_tax_wan !== null && lq.price_inc_tax_wan !== undefined ? Number(lq.price_inc_tax_wan) : null;
      const exTax = lq.price_ex_tax_wan !== null && lq.price_ex_tax_wan !== undefined ? Number(lq.price_ex_tax_wan) : null;
      const disc = String(lq.discount_rate || '').trim();
      const tax = String(lq.tax_rate || '').trim();
      if (tax && !grp.tax_rate) grp.tax_rate = tax;
      if (disc && disc !== '/' && !grp.discount_rate) grp.discount_rate = disc;
      if (incTax !== null) grp.inc_sum = Number(((grp.inc_sum ?? 0) + incTax).toFixed(6));
      if (exTax !== null) grp.ex_sum = Number(((grp.ex_sum ?? 0) + exTax).toFixed(6));
    }

    // Helper: detect ECP/ETP auto-carried Round 2+ rows in 行信息报价基础数据表 (when supplier did NOT submit Round 2 and system copies Round 1 verbatim)
    const isCarriedOverRound = (
      s: { round1_price: string; round2_price: string; line_by_round: Map<string, { round_order: number; inc_sum: number | null; discount_rate: string; rows: Array<Record<string, unknown>> }> },
      grp: { round_order: number; inc_sum: number | null; discount_rate: string; rows: Array<Record<string, unknown>> },
      hasBidOpeningRound2: boolean
    ): boolean => {
      if (grp.round_order < 2 || hasBidOpeningRound2) return false;
      const qStr = pickQuoteStr('', grp.inc_sum, grp.discount_rate);
      const r1Grp = Array.from(s.line_by_round.values()).find((g) => g.round_order === 1);
      const r1Str = s.round1_price || (r1Grp ? pickQuoteStr('', r1Grp.inc_sum, r1Grp.discount_rate) : '');
      if (!r1Str || !isSameQuoteVal(qStr, r1Str)) return false;
      // For merged sub-packages, also verify every sub-package quote matches Round 1
      if (isMergedSubPkg && r1Grp) {
        for (const row of grp.rows) {
          const reqNo = String(row.req_app_no || '').trim();
          const r1Row = r1Grp.rows.find((x) => isSameReqAppNo(x.req_app_no, reqNo));
          if (!r1Row) return false;
          const curQ = pickQuoteStr('', row.price_inc_tax_wan, row.discount_rate);
          const prevQ = pickQuoteStr('', r1Row.price_inc_tax_wan, r1Row.discount_rate);
          if (!isSameQuoteVal(curQ, prevQ)) return false;
        }
      }
      return true;
    };

    // Populate package-level round totals from grouped line_quotes (NEVER let a single sub-package overwrite whole-package round2/final,
    // and ignore auto-carried Round 2 rows that are identical to Round 1 when 开标一览表 has no Round 2)
    const carriedOverRoundLabels = new Map<string, Set<string>>();
    for (const s of supMap.values()) {
      const hasBidOpeningRound2 = Boolean(s.round2_price);
      const sortedLineRounds = Array.from(s.line_by_round.values()).sort((a, b) => a.round_order - b.round_order);
      for (const grp of sortedLineRounds) {
        const qStr = pickQuoteStr('', grp.inc_sum, grp.discount_rate);
        const carriedOver = isCarriedOverRound(s, grp, hasBidOpeningRound2);
        if (carriedOver) {
          if (!carriedOverRoundLabels.has(s.supplier_name)) carriedOverRoundLabels.set(s.supplier_name, new Set());
          carriedOverRoundLabels.get(s.supplier_name)!.add(grp.round_label);
        }
        if (!carriedOver && grp.round_order >= s.line_round_order) {
          s.line_round_order = grp.round_order;
          if (grp.tax_rate) s.tax_rate = grp.tax_rate;
          if (grp.inc_sum !== null) s.price_inc_tax_wan = grp.inc_sum;
          if (grp.ex_sum !== null) s.price_ex_tax_wan = grp.ex_sum;
          if (grp.discount_rate) s.discount_rate = grp.discount_rate;
        } else if (carriedOver && !s.tax_rate && grp.tax_rate) {
          s.tax_rate = grp.tax_rate;
        }
        if (carriedOver) continue;
        if (qStr && !s.whole_rounds.has(grp.round_label)) {
          s.whole_rounds.set(grp.round_label, qStr);
        }
        if (grp.round_order === 1 && !s.round1_price && qStr) s.round1_price = qStr;
        if (grp.round_order >= 2 && !s.round2_price && qStr) s.round2_price = qStr;
        if (qStr && grp.round_order > s.final_round_order) {
          s.final_price = qStr;
          s.final_round = grp.round_label;
          s.final_round_order = grp.round_order;
        } else if (!s.final_price && qStr) {
          s.final_price = qStr;
          s.final_round = grp.round_label;
          s.final_round_order = grp.round_order;
        }
      }
    }

    entry.supplier_bids = Array.from(supMap.values()).map((s) => {
      const overLimitDetails: OverLimitDetail[] = [];
      const subQuotes: SubQuoteItem[] = [];
      const skippedRounds = carriedOverRoundLabels.get(s.supplier_name);

      // 1. Check whole-package ceiling across ALL rounds
      const wholeRoundEntries = Array.from(s.whole_rounds.entries()).sort(
        ([a], [b]) => getRoundOrder(a) - getRoundOrder(b)
      );
      for (const [rLabel, qStr] of wholeRoundEntries) {
        const ov = checkOverLimit(qStr, entry.quote_method, entry.max_price_yuan, entry.max_price_raw);
        if (ov.isOver) {
          overLimitDetails.push({
            scope: '整包',
            round_label: rLabel,
            req_app_no: entry.req_app_no,
            sub_name: entry.project_name,
            limit_display: ov.limitDisplay,
            quote_display: ov.quoteDisplay,
            exceed_text: ov.exceedText,
            msg: `[整包·${rLabel}] ${ov.msg}`,
          });
        }
      }

      // 2. For merged sub-packages, check EVERY sub-package across EVERY actual round in line_quotes against its individual sub-package ceiling
      if (isMergedSubPkg) {
        const sortedLineRounds = Array.from(s.line_by_round.values()).sort((a, b) => a.round_order - b.round_order);
        for (const grp of sortedLineRounds) {
          if (skippedRounds?.has(grp.round_label)) continue;
          for (const lq of grp.rows) {
            const lqReq = String(lq.req_app_no || '').trim();
            const sp = entry.sub_packages.find((p) => isSameReqAppNo(p.req_app_no, lqReq));
            const incTax = lq.price_inc_tax_wan !== null && lq.price_inc_tax_wan !== undefined ? Number(lq.price_inc_tax_wan) : null;
            const disc = String(lq.discount_rate || '').trim();
            const qStr = pickQuoteStr('', incTax, disc);
            const spMethod = sp?.quote_method || entry.quote_method;
            const spMaxYuan = sp?.max_price_yuan ?? null;
            const spMaxRaw = sp?.max_price_raw || '';
            const spLimitDisp = spMaxRaw || (spMaxYuan !== null ? `${Number((spMaxYuan / 10000).toFixed(4))} 万元` : '—');
            const ov = sp ? checkOverLimit(qStr, spMethod, spMaxYuan, spMaxRaw) : { isOver: false, msg: '', limitDisplay: '', quoteDisplay: '', exceedText: '' };

            subQuotes.push({
              req_app_no: sp?.req_app_no || lqReq,
              sub_name: sp?.name || String(lq.project_name || lqReq),
              round_label: grp.round_label,
              price_inc_tax_wan: incTax,
              discount_rate: disc,
              limit_display: spLimitDisp,
              is_over_limit: ov.isOver,
              over_limit_msg: ov.msg,
            });

            if (ov.isOver && sp) {
              overLimitDetails.push({
                scope: '子包',
                round_label: grp.round_label,
                req_app_no: sp.req_app_no,
                sub_name: sp.name,
                limit_display: ov.limitDisplay,
                quote_display: ov.quoteDisplay,
                exceed_text: ov.exceedText,
                msg: `[子包·${grp.round_label}] ${sp.name}：${ov.msg}`,
              });
            }
          }
        }
      }

      return {
        supplier_name: s.supplier_name,
        round1_price: s.round1_price,
        round2_price: s.round2_price,
        final_price: s.final_price,
        final_round: s.final_round,
        tax_rate: s.tax_rate,
        price_inc_tax_wan: s.price_inc_tax_wan,
        price_ex_tax_wan: s.price_ex_tax_wan,
        discount_rate: s.discount_rate,
        is_rejected: rejMap.has(s.supplier_name),
        rejection_reason: rejMap.get(s.supplier_name) || '',
        is_over_limit: overLimitDetails.length > 0,
        over_limit_msg: overLimitDetails.map((d) => d.msg).join('；'),
        over_limit_details: overLimitDetails,
        sub_quotes: subQuotes,
      };
    });
    entry.over_limit_count = entry.supplier_bids.filter((s) => s.is_over_limit).length;

    // Compute bid opening price summary (min~max of final_price) so overview shows prices even before 总分排序表 is uploaded
    if (entry.supplier_bids.length > 0) {
      const nums = entry.supplier_bids
        .map((s) => Number(s.final_price))
        .filter((n) => Number.isFinite(n) && n > 0);
      if (nums.length > 0) {
        const minP = Math.min(...nums);
        const maxP = Math.max(...nums);
        entry.bid_price_summary = minP === maxP ? `${minP}` : `${minP} ~ ${maxP}`;
      } else {
        const rawSamples = Array.from(new Set(entry.supplier_bids.map((s) => s.final_price).filter(Boolean)));
        if (rawSamples.length > 0) entry.bid_price_summary = rawSamples.slice(0, 2).join(' / ');
      }
    }
  }

  const taxCategory: TaxCategory = String(batch.eval_tax_category || batch.award_category || '服务').includes('物资')
    ? '物资'
    : '服务';
  const proofRates = parseProofRatesCsv(
    batch.eval_tax_proof_rates !== undefined && batch.eval_tax_proof_rates !== null
      ? String(batch.eval_tax_proof_rates)
      : '0,1'
  );

  const packageList = Array.from(pkgMap.values());
  const quoteTaxAudit = runQuoteTaxAudit({
    bids,
    lineQuotes,
    projects,
    taxCategory,
    proofRates,
  });
  const lowQuotes = runLowQuoteAudit(packageList);
  quoteTaxAudit.lowQuotes = lowQuotes;
  quoteTaxAudit.counts.lowQuote = lowQuotes.length;

  const contactAudit = runContactCollusionAudit({
    bids,
    suppliers,
    projects,
  });
  const hardwareAudit = runHardwareCollusionAudit(ipRecords);

  const rawNotices = db
    .prepare('SELECT * FROM batch_change_notices WHERE batch_id = ? ORDER BY notice_date DESC, created_at DESC')
    .all(batchId) as Array<Record<string, unknown>>;
  const projById = new Map(projects.map((p) => [String(p.id), p]));
  const changeNotices = rawNotices.map((n) => {
    let change_types: string[] = [];
    let project_ids: string[] = [];
    try { change_types = JSON.parse(String(n.change_types_json || '[]')); } catch { change_types = []; }
    try { project_ids = JSON.parse(String(n.project_ids_json || '[]')); } catch { project_ids = []; }
    const affected_packages = project_ids
      .map((pid) => projById.get(pid))
      .filter((p): p is Record<string, unknown> => Boolean(p))
      .map((p) => ({
        lot_code: String(p.lot_code || ''),
        lot_name: String(p.lot_name || ''),
        package_no: String(p.package_no || ''),
        project_name: String(p.name || ''),
      }));
    return {
      id: String(n.id),
      batch_id: String(n.batch_id),
      title: String(n.title || ''),
      notice_date: String(n.notice_date || ''),
      change_types,
      summary: String(n.summary || ''),
      project_ids,
      affected_packages,
      file_original_name: String(n.file_original_name || ''),
      file_mime_type: String(n.file_mime_type || ''),
      has_file: Boolean(n.file_stored_path),
      created_at: String(n.created_at || ''),
    };
  });

  return {
    batch,
    slots,
    otherFiles,
    sopItems,
    packages: packageList,
    rejections,
    lineQuotes,
    changeNotices,
    auditConfig: {
      tax_category: taxCategory,
      proof_rates: proofRates,
    },
    quoteTaxAudit,
    contactAudit,
    hardwareAudit,
    stats: {
      bidRowCount: bids.length,
      lineQuoteCount: lineQuotes.length,
      rejectionCount: rejections.length,
      rankingRowCount: rankings.length,
      ipRowCount: ipRecords.length,
    },
  };
}

function appendStyledAuditSheet(
  wb: ExcelJS.Workbook,
  sheetName: string,
  headers: string[],
  widths: number[],
  rowsData: unknown[][],
  centerCols: Set<number>
) {
  const ws = wb.addWorksheet(sheetName, {
    pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  });
  ws.columns = headers.map((header, i) => ({ header, key: `c${i}`, width: widths[i] }));
  const borderThin: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FF000000' } },
  };
  const headerRow = ws.getRow(1);
  headerRow.height = 22;
  for (let c = 1; c <= headers.length; c++) {
    const cell = headerRow.getCell(c);
    cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF000000' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = borderThin;
  }
  for (const rVals of rowsData) {
    const row = ws.addRow(rVals);
    const maxLines = rVals.reduce<number>((acc, val) => {
      if (val === null || val === undefined) return acc;
      return Math.max(acc, String(val).split('\n').length);
    }, 1);
    row.height = Math.max(20, maxLines * 17);
    for (let c = 1; c <= headers.length; c++) {
      const cell = row.getCell(c);
      cell.font = { name: 'Calibri', size: 11, color: { argb: 'FF000000' } };
      cell.alignment = {
        vertical: 'middle',
        horizontal: centerCols.has(c) ? 'center' : 'left',
        wrapText: true,
      };
      cell.border = borderThin;
    }
  }
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: headers.length } };
  ws.pageSetup.printTitlesRow = '1:1';
}

// GET /api/evaluation/batch/:batchId/workspace
evaluationRouter.get('/batch/:batchId/workspace', (req, res) => {
  const data = buildEvaluationWorkspaceData(req.params.batchId);
  if (!data) return res.status(404).json({ error: '未找到该批次' });
  res.json(data);
});

// PUT /api/evaluation/batch/:batchId/audit-config — persist tax category & proof-check tax rates
evaluationRouter.put('/batch/:batchId/audit-config', (req, res) => {
  const batchId = req.params.batchId;
  const { tax_category, proof_rates } = req.body as {
    tax_category?: string;
    proof_rates?: number[] | string;
  };
  const existing = db.prepare('SELECT id FROM batches WHERE id = ?').get(batchId);
  if (!existing) return res.status(404).json({ error: '未找到该批次' });

  const cleanCat: TaxCategory = String(tax_category || '服务').includes('物资') ? '物资' : '服务';
  const ratesCsv = Array.isArray(proof_rates)
    ? proof_rates.map((n) => Number(n)).filter((n) => Number.isFinite(n) && n >= 0).join(',')
    : String(proof_rates ?? '0,1');

  db.prepare('UPDATE batches SET eval_tax_category = ?, eval_tax_proof_rates = ?, updated_at = ? WHERE id = ?').run(
    cleanCat,
    ratesCsv,
    new Date().toISOString(),
    batchId
  );
  res.json({ ok: true, tax_category: cleanCat, proof_rates: parseProofRatesCsv(ratesCsv) });
});

// GET /api/evaluation/batch/:batchId/quote-audit/export — export 报价核验表 (超限价 + 报价过低 + 二轮超一轮 + 相同报价)
evaluationRouter.get('/batch/:batchId/quote-audit/export', async (req, res) => {
  try {
    const data = buildEvaluationWorkspaceData(req.params.batchId);
    if (!data) return res.status(404).json({ error: '未找到该批次' });
    const qt = data.quoteTaxAudit;
    const wb = new ExcelJS.Workbook();
    wb.creator = '国网招标代理 SOP 工作台';

    const sortedPkgs = [...data.packages]
      .filter((p) => (p.over_limit_count || 0) > 0)
      .sort((a, b) => {
        const lotCmp = evalPinyinCollator.compare(a.lot_name || a.lot_code, b.lot_name || b.lot_code);
        if (lotCmp !== 0) return lotCmp;
        return extractEvalPkgNum(a.package_no) - extractEvalPkgNum(b.package_no);
      });
    const overLimitExportRows: unknown[][] = [];
    let olSeq = 0;
    for (const pkg of sortedPkgs) {
      const isDiscountMode = /折扣|下浮|费率|百分/.test(pkg.quote_method) || /%$/.test(String(pkg.max_price_raw || '').trim());
      const quoteMethodLabel = pkg.quote_method || (isDiscountMode ? '折扣率报价' : '总价报价');
      for (const sb of pkg.supplier_bids) {
        if (!sb.is_over_limit) continue;
        for (const d of sb.over_limit_details || []) {
          olSeq++;
          overLimitExportRows.push([
            olSeq, d.scope, d.round_label || '第1轮', pkg.lot_code || '', pkg.lot_name || '', pkg.package_no || '',
            d.req_app_no || pkg.req_app_no || '', d.sub_name || pkg.project_name || '', quoteMethodLabel,
            d.limit_display || '—', sb.supplier_name, d.quote_display || '—', d.exceed_text || '—',
            sb.is_rejected ? `已否决（${sb.rejection_reason || '否决'}）` : '未否决（需核查）', d.msg,
          ]);
        }
      }
    }
    appendStyledAuditSheet(
      wb,
      '超限价',
      ['序号', '核验维度', '超限轮次', '分标编号', '分标名称', '包号', '采购申请号', '项目/子包名称', '报价方式', '最高限价', '超限供应商名称', '超限报价', '超出金额/幅度', '初评否决状态', '超限核验说明'],
      [7, 10, 10, 20, 22, 10, 22, 34, 14, 16, 30, 16, 20, 22, 40],
      overLimitExportRows,
      new Set([1, 2, 3, 4, 6, 10, 12, 13, 14])
    );

    appendStyledAuditSheet(
      wb,
      '报价过低',
      ['序号', '分标编号', '分标名称', '包号', '项目名称', '投标人名称', '轮次', '报价口径', '应答报价', '通过初评均值', '通过初评次低价', '最高限价', '触发规则', '核验说明'],
      [7, 20, 22, 10, 32, 30, 10, 14, 14, 14, 14, 14, 28, 48],
      (qt.lowQuotes || []).map((lq, i) => [i + 1, lq.lot_code, lq.lot_name, lq.package_no, lq.project_name, lq.supplier_name, lq.round_label, lq.quote_basis, lq.quote_value, lq.avg_value, lq.second_lowest_value, lq.max_price_value, lq.triggered_rules.join('；'), lq.detail]),
      new Set([1, 2, 4, 7, 8, 9, 10, 11, 12])
    );

    appendStyledAuditSheet(
      wb,
      '二轮超一轮',
      ['序号', '核验维度', '分标编号', '分标名称', '包号', '采购申请号', '项目名称', '投标人名称', '报价口径', '第1轮报价', '第2轮报价', '上涨幅度', '说明'],
      [7, 10, 20, 22, 10, 22, 32, 30, 16, 14, 14, 16, 42],
      qt.roundIncreases.map((ri, i) => [i + 1, ri.scope || '整包', ri.lot_code, ri.lot_name, ri.package_no, ri.req_app_no, ri.project_name, ri.supplier_name, ri.quote_basis, ri.round1_value, ri.round2_value, ri.increase_text, ri.detail]),
      new Set([1, 2, 3, 5, 6, 9, 10, 11, 12])
    );

    appendStyledAuditSheet(
      wb,
      '相同报价',
      ['序号', '核验维度', '分标编号', '分标名称', '包号', '轮次', '采购申请号', '项目名称', '报价口径', '相同报价', '涉及供应商', '说明'],
      [7, 10, 20, 22, 10, 10, 22, 32, 16, 16, 40, 42],
      qt.sameQuotes.map((sq, i) => [i + 1, sq.scope || '整包', sq.lot_code, sq.lot_name, sq.package_no, sq.round_label, sq.req_app_no, sq.project_name, sq.quote_basis, sq.quote_value, sq.suppliers.join('、'), sq.detail]),
      new Set([1, 2, 3, 5, 6, 7, 9, 10])
    );

    const safeCode = String(data.batch.code || 'BATCH').replace(/[^\w\u4e00-\u9fa5-]/g, '_');
    const filename = `${safeCode}_报价核验.xlsx`;
    const buf = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.send(Buffer.from(buf));
  } catch (err) {
    res.status(500).json({ error: `导出报价核验失败：${(err as Error).message}` });
  }
});

// GET /api/evaluation/batch/:batchId/tax-audit/export — export 税率核验表 (税率异常 + 税率缺失 + 税率证明待核实)
evaluationRouter.get('/batch/:batchId/tax-audit/export', async (req, res) => {
  try {
    const data = buildEvaluationWorkspaceData(req.params.batchId);
    if (!data) return res.status(404).json({ error: '未找到该批次' });
    const qt = data.quoteTaxAudit;
    const wb = new ExcelJS.Workbook();
    wb.creator = '国网招标代理 SOP 工作台';

    const taxHeaders = ['序号', '核验结论', '分标编号', '分标名称', '包号', '轮次', '采购申请号', '项目名称', '投标人名称', '填报税率', '说明'];
    const taxWidths = [7, 16, 20, 22, 10, 10, 22, 32, 30, 12, 42];
    const taxCenter = new Set([1, 2, 3, 5, 6, 7, 10]);

    appendStyledAuditSheet(
      wb,
      '税率异常',
      taxHeaders,
      taxWidths,
      qt.taxAnomalies.map((t, i) => [i + 1, t.label, t.lot_code, t.lot_name, t.package_no, t.round_label, t.req_app_no, t.project_name, t.supplier_name, t.tax_rate_raw, t.detail]),
      taxCenter
    );
    appendStyledAuditSheet(
      wb,
      '税率缺失',
      taxHeaders,
      taxWidths,
      qt.taxMissing.map((t, i) => [i + 1, t.label, t.lot_code, t.lot_name, t.package_no, t.round_label, t.req_app_no, t.project_name, t.supplier_name, t.tax_rate_raw, t.detail]),
      taxCenter
    );
    appendStyledAuditSheet(
      wb,
      '税率证明待核实',
      taxHeaders,
      taxWidths,
      qt.taxProofNeeded.map((t, i) => [i + 1, t.label, t.lot_code, t.lot_name, t.package_no, t.round_label, t.req_app_no, t.project_name, t.supplier_name, t.tax_rate_raw, t.detail]),
      taxCenter
    );

    const safeCode = String(data.batch.code || 'BATCH').replace(/[^\w\u4e00-\u9fa5-]/g, '_');
    const filename = `${safeCode}_税率核验.xlsx`;
    const buf = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.send(Buffer.from(buf));
  } catch (err) {
    res.status(500).json({ error: `导出税率核验失败：${(err as Error).message}` });
  }
});

// GET /api/evaluation/batch/:batchId/quote-tax-audit/export — legacy combined export
evaluationRouter.get('/batch/:batchId/quote-tax-audit/export', async (req, res) => {
  try {
    const data = buildEvaluationWorkspaceData(req.params.batchId);
    if (!data) return res.status(404).json({ error: '未找到该批次' });
    const qt = data.quoteTaxAudit;
    const wb = new ExcelJS.Workbook();
    wb.creator = '国网招标代理 SOP 工作台';

    const taxAll = [...qt.taxAnomalies, ...qt.taxMissing, ...qt.taxProofNeeded];
    appendStyledAuditSheet(
      wb,
      '税率异常与证明核查',
      ['序号', '核验结论', '分标编号', '分标名称', '包号', '轮次', '采购申请号', '项目名称', '投标人名称', '填报税率', '说明'],
      [7, 16, 20, 22, 10, 10, 22, 32, 30, 12, 42],
      taxAll.map((t, i) => [i + 1, t.label, t.lot_code, t.lot_name, t.package_no, t.round_label, t.req_app_no, t.project_name, t.supplier_name, t.tax_rate_raw, t.detail]),
      new Set([1, 2, 3, 5, 6, 7, 10])
    );

    appendStyledAuditSheet(
      wb,
      '相同报价核验',
      ['序号', '核验维度', '分标编号', '分标名称', '包号', '轮次', '采购申请号', '项目名称', '报价口径', '相同报价', '涉及供应商', '说明'],
      [7, 10, 20, 22, 10, 10, 22, 32, 16, 16, 40, 42],
      qt.sameQuotes.map((sq, i) => [i + 1, sq.scope || '整包', sq.lot_code, sq.lot_name, sq.package_no, sq.round_label, sq.req_app_no, sq.project_name, sq.quote_basis, sq.quote_value, sq.suppliers.join('、'), sq.detail]),
      new Set([1, 2, 3, 5, 6, 7, 9, 10])
    );

    appendStyledAuditSheet(
      wb,
      '二轮超一轮',
      ['序号', '核验维度', '分标编号', '分标名称', '包号', '采购申请号', '项目名称', '投标人名称', '报价口径', '第1轮报价', '第2轮报价', '上涨幅度', '说明'],
      [7, 10, 20, 22, 10, 22, 32, 30, 16, 14, 14, 16, 40],
      qt.roundIncreases.map((ri, i) => [i + 1, ri.scope || '整包', ri.lot_code, ri.lot_name, ri.package_no, ri.req_app_no, ri.project_name, ri.supplier_name, ri.quote_basis, ri.round1_value, ri.round2_value, ri.increase_text, ri.detail]),
      new Set([1, 2, 3, 5, 6, 9, 10, 11, 12])
    );

    appendStyledAuditSheet(
      wb,
      '报价过低',
      ['序号', '分标编号', '分标名称', '包号', '项目名称', '投标人名称', '轮次', '报价口径', '应答报价', '通过初评均值', '通过初评次低价', '最高限价', '触发规则', '核验说明'],
      [7, 20, 22, 10, 32, 30, 10, 14, 14, 14, 14, 14, 28, 48],
      (qt.lowQuotes || []).map((lq, i) => [i + 1, lq.lot_code, lq.lot_name, lq.package_no, lq.project_name, lq.supplier_name, lq.round_label, lq.quote_basis, lq.quote_value, lq.avg_value, lq.second_lowest_value, lq.max_price_value, lq.triggered_rules.join('；'), lq.detail]),
      new Set([1, 2, 4, 7, 8, 9, 10, 11, 12])
    );

    const safeCode = String(data.batch.code || 'BATCH').replace(/[^\w\u4e00-\u9fa5-]/g, '_');
    const filename = `${safeCode}_报价与税率合规核验.xlsx`;
    const buf = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.send(Buffer.from(buf));
  } catch (err) {
    res.status(500).json({ error: `导出报价与税率核验失败：${(err as Error).message}` });
  }
});

// GET /api/evaluation/batch/:batchId/contact-audit/export — export 同标包供应商联系信息核验
evaluationRouter.get('/batch/:batchId/contact-audit/export', async (req, res) => {
  try {
    const data = buildEvaluationWorkspaceData(req.params.batchId);
    if (!data) return res.status(404).json({ error: '未找到该批次' });
    const ca = data.contactAudit;
    const wb = new ExcelJS.Workbook();
    wb.creator = '国网招标代理 SOP 工作台';

    appendStyledAuditSheet(
      wb,
      '同标包联系信息重合预警',
      ['序号', '分标编号', '分标名称', '包号', '预警类型', '重合字段', '重合内容', '涉及供应商（同标包）'],
      [7, 20, 22, 10, 22, 14, 34, 48],
      ca.overlaps.map((o, i) => [i + 1, o.lot_code, o.lot_name, o.package_no, o.alert_label, o.field_label, o.matched_value, o.suppliers.join('、')]),
      new Set([1, 2, 4, 5, 6])
    );

    appendStyledAuditSheet(
      wb,
      '未找到联系人信息',
      ['序号', '分标编号', '分标名称', '包号', '投标人名称', '提示'],
      [7, 20, 22, 10, 36, 30],
      ca.missingContacts.map((m, i) => [i + 1, m.lot_code, m.lot_name, m.package_no, m.supplier_name, m.detail]),
      new Set([1, 2, 4])
    );

    const safeCode = String(data.batch.code || 'BATCH').replace(/[^\w\u4e00-\u9fa5-]/g, '_');
    const filename = `${safeCode}_同标包供应商联系信息核验.xlsx`;
    const buf = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.send(Buffer.from(buf));
  } catch (err) {
    res.status(500).json({ error: `导出联系信息核验失败：${(err as Error).message}` });
  }
});

// GET /api/evaluation/batch/:batchId/hardware-audit/export — export 投标文件硬件特征核验
evaluationRouter.get('/batch/:batchId/hardware-audit/export', async (req, res) => {
  try {
    const data = buildEvaluationWorkspaceData(req.params.batchId);
    if (!data) return res.status(404).json({ error: '未找到该批次' });
    const ha = data.hardwareAudit;
    const wb = new ExcelJS.Workbook();
    wb.creator = '国网招标代理 SOP 工作台';

    appendStyledAuditSheet(
      wb,
      '围串标硬件特征预警',
      ['序号', '分标编号', '分标名称', '包号', '投标人A', '甲方文件及来源', '投标人B', '乙方文件及来源', 'CPU序列号', '硬盘序列号', '网卡MAC地址', '提交IP(A / B)', '核验结论'],
      [7, 20, 20, 10, 28, 44, 28, 44, 22, 32, 20, 28, 42],
      ha.alerts.map((a, i) => [
        i + 1,
        a.lot_code,
        a.lot_name,
        a.package_no,
        a.supplier_a,
        a.source_a || `${a.round_a} ${a.file_a}`,
        a.supplier_b,
        a.source_b || `${a.round_b} ${a.file_b}`,
        a.cpu,
        a.disk,
        a.mac,
        `${a.ip_a || '/'} / ${a.ip_b || '/'}`,
        '围串标预警（CPU+硬盘+MAC全部一致）',
      ]),
      new Set([1, 2, 4, 11])
    );

    appendStyledAuditSheet(
      wb,
      '待人工核实(缺MAC)',
      ['序号', '分标编号', '分标名称', '包号', '投标人A', '甲方文件及来源', '投标人B', '乙方文件及来源', 'CPU序列号', '硬盘序列号', '网卡MAC(A / B)', '提交IP(A / B)', '核验结论'],
      [7, 20, 20, 10, 28, 44, 28, 44, 22, 32, 26, 28, 42],
      (ha.manualReview || []).map((m, i) => [
        i + 1,
        m.lot_code,
        m.lot_name,
        m.package_no,
        m.supplier_a,
        m.source_a || `${m.round_a} ${m.file_a}`,
        m.supplier_b,
        m.source_b || `${m.round_b} ${m.file_b}`,
        m.cpu,
        m.disk,
        m.mac,
        `${m.ip_a || '/'} / ${m.ip_b || '/'}`,
        m.review_reason,
      ]),
      new Set([1, 2, 4, 11])
    );

    appendStyledAuditSheet(
      wb,
      '资料不足未完成核验',
      ['序号', '分标编号', '分标名称', '包号', '轮次', '投标人名称', '文件类别', '来源位置', '提交IP', 'MAC地址', 'CPU序列号', '硬盘序列号', '核验结论'],
      [7, 20, 20, 10, 10, 30, 14, 44, 18, 20, 22, 28, 34],
      ha.insufficient.map((ins, i) => [
        i + 1,
        ins.lot_code,
        ins.lot_name,
        ins.package_no,
        ins.round_label,
        ins.supplier_name,
        ins.file_type,
        ins.source_location || '',
        ins.ip || '/',
        ins.mac || '/',
        ins.cpu || '/',
        ins.disk || '/',
        ins.detail,
      ]),
      new Set([1, 2, 4, 5, 7])
    );

    const safeCode = String(data.batch.code || 'BATCH').replace(/[^\w\u4e00-\u9fa5-]/g, '_');
    const filename = `${safeCode}_投标文件硬件特征核验.xlsx`;
    const buf = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.send(Buffer.from(buf));
  } catch (err) {
    res.status(500).json({ error: `导出硬件特征核验失败：${(err as Error).message}` });
  }
});

// GET /api/evaluation/batch/:batchId/over-limit/export — export all over-limit packages/suppliers as formatted .xlsx
evaluationRouter.get('/batch/:batchId/over-limit/export', async (req, res) => {
  try {
    const data = buildEvaluationWorkspaceData(req.params.batchId);
    if (!data) return res.status(404).json({ error: '未找到该批次' });

    const sortedPkgs = [...data.packages]
      .filter((p) => (p.over_limit_count || 0) > 0)
      .sort((a, b) => {
        const lotCmp = evalPinyinCollator.compare(a.lot_name || a.lot_code, b.lot_name || b.lot_code);
        if (lotCmp !== 0) return lotCmp;
        return extractEvalPkgNum(a.package_no) - extractEvalPkgNum(b.package_no);
      });

    const rowsData: unknown[][] = [];
    let seq = 0;

    for (const pkg of sortedPkgs) {
      const isDiscountMode =
        /折扣|下浮|费率|百分/.test(pkg.quote_method) || /%$/.test(String(pkg.max_price_raw || '').trim());
      const quoteMethodLabel = pkg.quote_method || (isDiscountMode ? '折扣率报价' : '总价报价');

      for (const sb of pkg.supplier_bids) {
        if (!sb.is_over_limit) continue;
        const details = sb.over_limit_details || [];
        for (const d of details) {
          seq++;
          rowsData.push([
            seq,
            d.scope,
            d.round_label || '第1轮',
            pkg.lot_code || '',
            pkg.lot_name || '',
            pkg.package_no || '',
            d.req_app_no || pkg.req_app_no || '',
            d.sub_name || pkg.project_name || '',
            quoteMethodLabel,
            d.limit_display || '—',
            sb.supplier_name,
            d.quote_display || '—',
            d.exceed_text || '—',
            sb.is_rejected ? `已否决（${sb.rejection_reason || '否决'}）` : '未否决（需核查）',
            d.msg,
          ]);
        }
      }
    }

    if (rowsData.length === 0) {
      return res.status(400).json({ error: '当前批次暂无超限价记录可导出' });
    }

    const wb = new ExcelJS.Workbook();
    wb.creator = '国网招标代理 SOP 工作台';
    const ws = wb.addWorksheet('超限价核验明细', {
      pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    });

    const headers = [
      '序号',
      '核验维度',
      '超限轮次',
      '分标编号',
      '分标名称',
      '包号',
      '采购申请号',
      '项目/子包名称',
      '报价方式',
      '最高限价',
      '超限供应商名称',
      '超限报价',
      '超出金额/幅度',
      '初评否决状态',
      '超限核验说明',
    ];
    const widths = [7, 10, 10, 20, 22, 10, 22, 34, 14, 16, 30, 16, 20, 22, 40];
    ws.columns = headers.map((header, i) => ({ header, key: `c${i}`, width: widths[i] }));

    const borderThin: Partial<ExcelJS.Borders> = {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } },
    };

    const headerRow = ws.getRow(1);
    headerRow.height = 22;
    for (let c = 1; c <= headers.length; c++) {
      const cell = headerRow.getCell(c);
      cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF000000' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      cell.border = borderThin;
    }

    const centerCols = new Set([1, 2, 4, 5, 7, 8, 10, 11, 12, 13, 14]);
    for (const rVals of rowsData) {
      const row = ws.addRow(rVals);
      row.height = 20;
      for (let c = 1; c <= headers.length; c++) {
        const cell = row.getCell(c);
        cell.font = { name: 'Calibri', size: 11, color: { argb: 'FF000000' } };
        cell.alignment = {
          vertical: 'middle',
          horizontal: centerCols.has(c) ? 'center' : 'left',
          wrapText: true,
        };
        cell.border = borderThin;
      }
    }

    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: headers.length } };
    ws.pageSetup.printTitlesRow = '1:1';

    const safeCode = String(data.batch.code || 'BATCH').replace(/[^\w\u4e00-\u9fa5-]/g, '_');
    const filename = `${safeCode}_超限价核验明细.xlsx`;
    const buf = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.send(Buffer.from(buf));
  } catch (err) {
    res.status(500).json({ error: `导出超限价明细失败：${(err as Error).message}` });
  }
});

// POST /api/evaluation/batch/:batchId/upload-slot/:slotKey
evaluationRouter.post('/batch/:batchId/upload-slot/:slotKey', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: '请选择要上传的开评标 Excel 文件' });

  const batchId = req.params.batchId;
  const slotKey = req.params.slotKey;
  const force = String(req.body?.force || '') === 'true';
  const slotDef = EVAL_SLOTS.find((s) => s.key === slotKey);
  if (!slotDef) {
    if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    return res.status(400).json({ error: '无效的开评标文件槽位' });
  }

  const batch = db.prepare('SELECT * FROM batches WHERE id = ?').get(batchId) as Record<string, string> | undefined;
  if (!batch) {
    if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    return res.status(404).json({ error: '批次不存在' });
  }

  const currentBatchCode = batch.code ? normalizeBatchCode(batch.code) : '';
  const originalName = decodeName(req.file.originalname);

  try {
    let detectedBatchCode = '';
    let summary = '';
    const warnings: string[] = [];

    // Pre-parse to extract detectedBatchCode before committing
    let parsedBid: ReturnType<typeof parseBidOpeningSheet> | null = null;
    let parsedLine: ReturnType<typeof parseLineQuotesSheet> | null = null;
    let parsedIp: ReturnType<typeof parseIpHardwareSheet> | null = null;
    let parsedRej: ReturnType<typeof parseRejectionsSheet> | null = null;
    let parsedRank: ReturnType<typeof parseRankingsSheet> | null = null;

    if (slotKey === 'bid_opening') {
      parsedBid = parseBidOpeningSheet(req.file.path);
      detectedBatchCode = parsedBid.detectedBatchCode;
    } else if (slotKey === 'line_quotes') {
      parsedLine = parseLineQuotesSheet(req.file.path);
      detectedBatchCode = parsedLine.detectedBatchCode;
    } else if (slotKey === 'ip_hardware') {
      parsedIp = parseIpHardwareSheet(req.file.path, originalName);
      detectedBatchCode = parsedIp.detectedBatchCode;
    } else if (slotKey === 'rejections') {
      parsedRej = parseRejectionsSheet(req.file.path);
      detectedBatchCode = parsedRej.detectedBatchCode;
    } else if (slotKey === 'rankings') {
      parsedRank = parseRankingsSheet(req.file.path);
      detectedBatchCode = parsedRank.detectedBatchCode;
    }

    // Cross-Batch Mismatch Guardrail!
    if (currentBatchCode && detectedBatchCode && currentBatchCode !== detectedBatchCode && !force) {
      if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
      return res.status(409).json({
        mismatch: true,
        currentBatchCode,
        fileBatchCode: detectedBatchCode,
        error: `批次不匹配！当前工作台批次编号为「${currentBatchCode}」，但上传文件内容属于「${detectedBatchCode}」，请确认是否传错批次。`,
      });
    }

    if (currentBatchCode && detectedBatchCode && currentBatchCode !== detectedBatchCode && force) {
      warnings.push(`注意：已强制导入不同批次文件（当前批次 ${currentBatchCode} ← 文件批次 ${detectedBatchCode}）`);
    }
    const relPath = path.relative(UPLOADS_DIR, req.file.path);
    const now = new Date().toISOString();

    let parsedPreviewJson: string | null = null;
    try {
      const tables = parseSpreadsheetTables(req.file.path);
      if (tables) parsedPreviewJson = JSON.stringify(tables);
    } catch {
      // ignore
    }

    const fileId = uuid();
    db.prepare(`
      INSERT INTO uploaded_files (id, batch_id, module, slot_key, category, original_name, stored_path, mime_type, file_size, parsed_json, notes, created_at)
      VALUES (?, ?, 'evaluation', ?, ?, ?, ?, ?, ?, ?, '', ?)
    `).run(
      fileId, batchId, slotKey, slotDef.label, originalName, relPath,
      req.file.mimetype || 'application/octet-stream', req.file.size, parsedPreviewJson, now
    );

    // Helper: run immediate ALL-ROUND over-limit check (supporting both whole-package and per-sub-package ceilings)
    const runUploadOverLimitAudit = (
      quoteRows: Array<{ lot_code: string; package_no: string; round_label: string; supplier_name: string; req_app_no?: string; quoteStr: string }>,
      isLineLevel: boolean
    ) => {
      const projs = db.prepare('SELECT lot_code, lot_name, package_no, req_app_no, name, quote_method, budget_yuan, max_price_yuan, max_price_raw FROM batch_projects WHERE batch_id = ?').all(batchId) as Array<Record<string, unknown>>;
      const pkgLimitMap = new Map<string, {
        lotName: string;
        quoteMethod: string;
        maxYuan: number | null;
        maxRaw: string;
        subs: Array<{ req_app_no: string; name: string; quoteMethod: string; maxYuan: number | null; maxRaw: string }>;
      }>();
      for (const p of projs) {
        const k = `${String(p.lot_code || '')}::${String(p.package_no || '')}`;
        const pMax = p.max_price_yuan !== null && p.max_price_yuan !== undefined ? Number(p.max_price_yuan) : (p.budget_yuan !== null && p.budget_yuan !== undefined ? Number(p.budget_yuan) : null);
        const sub = {
          req_app_no: String(p.req_app_no || ''),
          name: String(p.name || ''),
          quoteMethod: String(p.quote_method || ''),
          maxYuan: pMax,
          maxRaw: String(p.max_price_raw || ''),
        };
        const cur = pkgLimitMap.get(k);
        if (!cur) {
          pkgLimitMap.set(k, {
            lotName: String(p.lot_name || p.lot_code || ''),
            quoteMethod: sub.quoteMethod,
            maxYuan: pMax,
            maxRaw: sub.maxRaw,
            subs: [sub],
          });
        } else {
          if (pMax !== null) cur.maxYuan = (cur.maxYuan || 0) + pMax;
          cur.subs.push(sub);
        }
      }
      const overList: string[] = [];
      for (const q of quoteRows) {
        if (!q.supplier_name || !q.quoteStr) continue;
        const lim = pkgLimitMap.get(`${q.lot_code}::${q.package_no}`);
        if (!lim) continue;
        if (isLineLevel && lim.subs.length > 1 && q.req_app_no) {
          const sp = lim.subs.find((s) => isSameReqAppNo(s.req_app_no, q.req_app_no));
          if (!sp) continue;
          const resCheck = checkOverLimit(q.quoteStr, sp.quoteMethod || lim.quoteMethod, sp.maxYuan, sp.maxRaw);
          if (resCheck.isOver) {
            overList.push(`[${lim.lotName} ${q.package_no} · 子包·${q.round_label}] ${q.supplier_name} (${sp.name}) — ${resCheck.msg}`);
          }
        } else {
          const resCheck = checkOverLimit(q.quoteStr, lim.quoteMethod, lim.maxYuan, lim.maxRaw);
          if (resCheck.isOver) {
            overList.push(`[${lim.lotName} ${q.package_no} · ${q.round_label}] ${q.supplier_name} — ${resCheck.msg}`);
          }
        }
      }
      if (overList.length > 0) {
        warnings.push(`超限价核查预警：共发现 ${overList.length} 条各轮次/子包报价超过最高限价（请前往「报价与税率核验 → 超限价」查看明细或导出 Excel）`);
        overList.slice(0, 8).forEach((line) => warnings.push(line));
        if (overList.length > 8) {
          warnings.push(`……其余 ${overList.length - 8} 条超限价记录请在「报价与税率核验 → 超限价」子版块查看或导出`);
        }
      }
    };

    if (parsedBid) {
      db.prepare('DELETE FROM eval_bids WHERE batch_id = ?').run(batchId);
      const ins = db.prepare(`
        INSERT INTO eval_bids (id, batch_id, lot_code, lot_name, package_no, round_label, supplier_name, discount_rate, bid_price_wan, bid_price_raw, remark, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `);
      parsedBid.rows.forEach((r) => {
        ins.run(uuid(), batchId, r.lot_code, r.lot_name, r.package_no, r.round_label, r.supplier_name, r.discount_rate, r.bid_price_wan, r.bid_price_raw, r.remark, now);
      });
      summary = `已解析开标报价 ${parsedBid.rows.length} 条（包含轮次：${parsedBid.rounds.join('、')}，自动清理 ${parsedBid.spacesCleaned} 处供应商空格）`;
      if (batch.procurement_type === '公开招标' && parsedBid.rounds.length > 1) {
        warnings.push('当前批次为「公开招标」，但表中检测到多轮报价，请核对采购类型设置是否正确');
      }
      runUploadOverLimitAudit(
        parsedBid.rows.map((r) => ({
          lot_code: r.lot_code,
          package_no: r.package_no,
          round_label: r.round_label,
          supplier_name: r.supplier_name,
          quoteStr: (r.bid_price_raw && r.bid_price_raw !== '/') ? r.bid_price_raw : (r.bid_price_wan !== null ? String(r.bid_price_wan) : r.discount_rate),
        })),
        false
      );
    } else if (parsedLine) {
      db.prepare('DELETE FROM eval_line_quotes WHERE batch_id = ?').run(batchId);
      const ins = db.prepare(`
        INSERT INTO eval_line_quotes (id, batch_id, lot_code, package_no, round_label, supplier_name, req_app_no, project_name, tax_rate, price_ex_tax_wan, price_inc_tax_wan, discount_rate, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `);
      parsedLine.rows.forEach((r) => {
        ins.run(uuid(), batchId, r.lot_code, r.package_no, r.round_label, r.supplier_name, r.req_app_no, r.project_name, r.tax_rate, r.price_ex_tax_wan, r.price_inc_tax_wan, r.discount_rate, now);
      });
      summary = `已解析行信息报价 ${parsedLine.rows.length} 条（已提取采购申请号与增值税率，清理 ${parsedLine.spacesCleaned} 处空格）`;
      runUploadOverLimitAudit(
        parsedLine.rows.map((r) => ({
          lot_code: r.lot_code,
          package_no: r.package_no,
          round_label: r.round_label,
          supplier_name: r.supplier_name,
          req_app_no: r.req_app_no,
          quoteStr: r.price_inc_tax_wan !== null ? String(r.price_inc_tax_wan) : r.discount_rate,
        })),
        true
      );
    } else if (parsedIp) {
      db.prepare('DELETE FROM eval_ip_records WHERE batch_id = ?').run(batchId);
      const ins = db.prepare(`
        INSERT INTO eval_ip_records (
          id, batch_id, lot_code, lot_name, package_no, round_label, supplier_name, download_ip,
          tech_ip, biz_ip, price_ip, qual_ip,
          tech_mac, biz_mac, price_mac, qual_mac,
          tech_cpu, biz_cpu, price_cpu, qual_cpu,
          tech_disk, biz_disk, price_disk, qual_disk,
          source_file, sheet_name, source_row, created_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `);
      parsedIp.rows.forEach((r) => {
        ins.run(
          uuid(), batchId, r.lot_code, r.lot_name, r.package_no, r.round_label, r.supplier_name, r.download_ip,
          r.tech_ip, r.biz_ip, r.price_ip, r.qual_ip,
          r.tech_mac, r.biz_mac, r.price_mac, r.qual_mac,
          r.tech_cpu, r.biz_cpu, r.price_cpu, r.qual_cpu,
          r.tech_disk, r.biz_disk, r.price_disk, r.qual_disk,
          r.source_file, r.sheet_name, r.source_row, now
        );
      });
      const hwAudit = runHardwareCollusionAudit(parsedIp.rows as unknown as Array<Record<string, unknown>>);
      summary = `已解析投标文件硬件/IP 记录 ${parsedIp.rows.length} 条（清理 ${parsedIp.spacesCleaned} 处供应商空格）；围串标预警 ${hwAudit.alerts.length} 条，待人工核实 ${hwAudit.manualReview.length} 条，资料不足 ${hwAudit.insufficient.length} 条`;
      if (parsedIp.warnings?.length) {
        parsedIp.warnings.forEach((w) => warnings.push(w));
      }
      if (hwAudit.alerts.length > 0) {
        warnings.push(`围串标硬件特征预警：共发现 ${hwAudit.alerts.length} 组同标包不同投标人 CPU+硬盘+MAC 全部一致（详见「围串标风控核验」子版块）`);
      }
      if (hwAudit.manualReview.length > 0) {
        warnings.push(`硬件特征待人工核实：共发现 ${hwAudit.manualReview.length} 组 CPU+硬盘一致但单方/双方未提供 MAC（详见「围串标风控核验」子版块）`);
      }
      if (hwAudit.insufficient.length > 0) {
        warnings.push(`硬件资料不足：共 ${hwAudit.insufficient.length} 条文件记录缺少 CPU/硬盘序列号或标包定位信息，未完成核验`);
      }
    } else if (parsedRej) {
      db.prepare('DELETE FROM eval_rejections WHERE batch_id = ?').run(batchId);
      const ins = db.prepare(`
        INSERT INTO eval_rejections (id, batch_id, lot_code, lot_name, package_no, supplier_name, reason_category, rejection_tpl, doc_requirement, bid_response, review_type, notes, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `);
      parsedRej.rows.forEach((r) => {
        ins.run(uuid(), batchId, r.lot_code, r.lot_name, r.package_no, r.supplier_name, r.reason_category, r.rejection_tpl, r.doc_requirement, r.bid_response, r.review_type, r.notes, now);
      });
      summary = `已解析未进入详评（否决）记录 ${parsedRej.rows.length} 条`;
    } else if (parsedRank) {
      db.prepare('DELETE FROM eval_rankings WHERE batch_id = ?').run(batchId);
      const ins = db.prepare(`
        INSERT INTO eval_rankings (id, batch_id, lot_code, lot_name, package_no, project_name, eval_date, supplier_name, tech_score, biz_score, price_score, total_score, rank_no, base_price, bid_price, eval_price, is_failed_lot, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `);
      parsedRank.rows.forEach((r) => {
        ins.run(
          uuid(), batchId, r.lot_code, r.lot_name, r.package_no, r.project_name, r.eval_date,
          r.supplier_name, r.tech_score, r.biz_score, r.price_score, r.total_score,
          r.rank_no, r.base_price, r.bid_price, r.eval_price, r.is_failed_lot, now
        );
      });
      summary = `已解析总分排序记录 ${parsedRank.rows.length} 条（清理 ${parsedRank.spacesCleaned} 处空格）`;
      if (parsedRank.failedLotsCount > 0) {
        warnings.push(`检测到 ${parsedRank.failedLotsCount} 个标段为「流标（无数据）」，已在结果总览中自动标红`);
      }
    }

    res.status(201).json({
      ok: true,
      fileId,
      report: { slotLabel: slotDef.label, summary, warnings },
    });
  } catch (err) {
    if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    res.status(500).json({ error: `解析开评标表格失败：${(err as Error).message}` });
  }
});

// DELETE /api/evaluation/batch/:batchId/slot/:slotKey — clear parsed data & uploaded file(s) for a specific evaluation slot
evaluationRouter.delete('/batch/:batchId/slot/:slotKey', (req, res) => {
  const { batchId, slotKey } = req.params;
  const slotDef = EVAL_SLOTS.find((s) => s.key === slotKey);
  if (!slotDef) return res.status(400).json({ error: '无效的开评标文件槽位' });

  const tableMap: Record<string, string> = {
    bid_opening: 'eval_bids',
    line_quotes: 'eval_line_quotes',
    ip_hardware: 'eval_ip_records',
    rejections: 'eval_rejections',
    rankings: 'eval_rankings',
  };
  const targetTable = tableMap[slotKey];
  if (targetTable) {
    db.prepare(`DELETE FROM ${targetTable} WHERE batch_id = ?`).run(batchId);
  }

  const slotFiles = db
    .prepare("SELECT id, stored_path FROM uploaded_files WHERE batch_id = ? AND module = 'evaluation' AND slot_key = ?")
    .all(batchId, slotKey) as Array<{ id: string; stored_path: string }>;

  for (const f of slotFiles) {
    if (f.stored_path) {
      const absPath = path.join(UPLOADS_DIR, f.stored_path);
      if (fs.existsSync(absPath)) fs.unlinkSync(absPath);
    }
  }

  db.prepare("DELETE FROM uploaded_files WHERE batch_id = ? AND module = 'evaluation' AND slot_key = ?").run(batchId, slotKey);

  res.json({ ok: true, slotLabel: slotDef.label });
});

// ==================== 专家打分结果校验 (工具 8 总分校验 & 工具 9 客观分校验) ====================

const scoreMemoryUpload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 80 * 1024 * 1024, files: 100 },
});

interface PersistedScoreFileMeta {
  originalName: string;
  storedPath: string;
  size: number;
}

function saveExpandedScoreFiles(batchId: string, subDir: 'score_total' | 'score_objective' | 'score_precheck' | 'initial_report' | 'warning_form' | 'initial_minutes', expanded: ScoreUploadedFile[]): PersistedScoreFileMeta[] {
  const targetDir = path.join(UPLOADS_DIR, batchId, 'evaluation', subDir);
  if (fs.existsSync(targetDir)) {
    fs.rmSync(targetDir, { recursive: true, force: true });
  }
  fs.mkdirSync(targetDir, { recursive: true });
  const metas: PersistedScoreFileMeta[] = [];
  expanded.forEach((f, idx) => {
    const ext = path.extname(f.originalName) || '.xlsx';
    const safeDiskName = `${String(idx + 1).padStart(3, '0')}-${uuid().slice(0, 8)}${ext}`;
    const absPath = path.join(targetDir, safeDiskName);
    fs.writeFileSync(absPath, f.buffer);
    metas.push({
      originalName: f.originalName,
      storedPath: path.relative(UPLOADS_DIR, absPath),
      size: f.buffer.length,
    });
  });
  return metas;
}

function loadPersistedScoreFiles(metas: PersistedScoreFileMeta[]): ScoreUploadedFile[] {
  const list: ScoreUploadedFile[] = [];
  for (const m of metas) {
    const absPath = path.join(UPLOADS_DIR, m.storedPath);
    if (fs.existsSync(absPath)) {
      list.push({ originalName: m.originalName, buffer: fs.readFileSync(absPath) });
    }
  }
  return list;
}

function getScoreCheckRow(batchId: string, checkType: 'total' | 'objective' | 'precheck' | 'initial_report' | 'warning_form' | 'initial_minutes') {
  const row = db
    .prepare('SELECT * FROM eval_score_checks WHERE batch_id = ? AND check_type = ?')
    .get(batchId, checkType) as Record<string, unknown> | undefined;
  if (!row) return null;
  return {
    files: JSON.parse(String(row.files_json || '[]')) as PersistedScoreFileMeta[],
    config: JSON.parse(String(row.config_json || '{}')) as Record<string, unknown>,
    methods: JSON.parse(String(row.methods_json || '[]')) as unknown[],
    result: row.result_json ? JSON.parse(String(row.result_json)) : null,
    updated_at: String(row.updated_at || ''),
  };
}

function upsertScoreCheckRow(
  batchId: string,
  checkType: 'total' | 'objective' | 'precheck' | 'initial_report' | 'warning_form' | 'initial_minutes',
  payload: {
    files: PersistedScoreFileMeta[];
    config: Record<string, unknown>;
    methods?: unknown[];
    result: unknown;
  }
) {
  const now = new Date().toISOString();
  db.prepare(`
    INSERT INTO eval_score_checks (batch_id, check_type, files_json, config_json, methods_json, result_json, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(batch_id, check_type) DO UPDATE SET
      files_json = excluded.files_json,
      config_json = excluded.config_json,
      methods_json = excluded.methods_json,
      result_json = excluded.result_json,
      updated_at = excluded.updated_at
  `).run(
    batchId,
    checkType,
    JSON.stringify(payload.files),
    JSON.stringify(payload.config),
    JSON.stringify(payload.methods || []),
    payload.result ? JSON.stringify(payload.result) : null,
    now
  );
  return now;
}

// GET /api/evaluation/batch/:batchId/score-check
evaluationRouter.get('/batch/:batchId/score-check', (req, res) => {
  const { batchId } = req.params;
  const batch = db.prepare('SELECT id FROM batches WHERE id = ?').get(batchId);
  if (!batch) return res.status(404).json({ error: '未找到该批次' });
  const total = getScoreCheckRow(batchId, 'total');
  const objective = getScoreCheckRow(batchId, 'objective');
  res.json({
    total: total
      ? {
          files: total.files.map((f) => ({ originalName: f.originalName, size: f.size })),
          config: total.config,
          result: total.result as ExpertTotalAuditResult | null,
          updated_at: total.updated_at,
        }
      : null,
    objective: objective
      ? {
          files: objective.files.map((f) => ({ originalName: f.originalName, size: f.size })),
          config: objective.config,
          methods: objective.methods,
          result: objective.result as ExpertObjectiveAuditResult | null,
          updated_at: objective.updated_at,
        }
      : null,
  });
});

// POST /api/evaluation/batch/:batchId/score-check/total — upload & run 专家总分校验
evaluationRouter.post('/batch/:batchId/score-check/total', scoreMemoryUpload.array('files', 100), (req, res) => {
  try {
    const { batchId } = req.params;
    const batch = db.prepare('SELECT id FROM batches WHERE id = ?').get(batchId);
    if (!batch) return res.status(404).json({ error: '未找到该批次' });

    const uploaded = (req.files as Express.Multer.File[] | undefined) || [];
    if (uploaded.length === 0) return res.status(400).json({ error: '请至少选择一份专家总分统计表（.xlsx / .xls / .zip）' });

    const rawExpertNum = String(req.body?.expertNum ?? '').trim();
    const expertNum = rawExpertNum && Number(rawExpertNum) > 0 ? Number(rawExpertNum) : null;
    const rawDiff = String(req.body?.maxScoreDiff ?? '8').trim();
    const maxScoreDiff = rawDiff === '' ? 8 : Number(rawDiff);
    const password = String(req.body?.password ?? '');

    const scoreFiles: ScoreUploadedFile[] = uploaded.map((f) => ({
      originalName: decodeName(f.originalname),
      buffer: f.buffer,
    }));
    const expanded = expandScoreFiles(scoreFiles);
    const result = runExpertTotalAudit({ files: expanded, expertNum, maxScoreDiff, password });
    const metas = saveExpandedScoreFiles(batchId, 'score_total', expanded);
    const config = { expertNum, maxScoreDiff };
    const updated_at = upsertScoreCheckRow(batchId, 'total', { files: metas, config, result });

    res.status(201).json({
      files: metas.map((f) => ({ originalName: f.originalName, size: f.size })),
      config,
      result,
      updated_at,
    });
  } catch (err) {
    res.status(400).json({ error: (err as Error).message });
  }
});

// POST /api/evaluation/batch/:batchId/score-check/total/rerun — rerun with updated params on persisted files
evaluationRouter.post('/batch/:batchId/score-check/total/rerun', (req, res) => {
  try {
    const { batchId } = req.params;
    const existing = getScoreCheckRow(batchId, 'total');
    if (!existing || existing.files.length === 0) {
      return res.status(400).json({ error: '尚未上传专家总分工作簿，请先上传文件' });
    }
    const { expertNum: rawExp, maxScoreDiff: rawDiff, password } = req.body as {
      expertNum?: number | string | null;
      maxScoreDiff?: number | string;
      password?: string;
    };
    const expertNum = rawExp !== undefined && rawExp !== null && String(rawExp).trim() !== '' && Number(rawExp) > 0 ? Number(rawExp) : null;
    const maxScoreDiff = rawDiff !== undefined && String(rawDiff).trim() !== '' ? Number(rawDiff) : 8;

    const files = loadPersistedScoreFiles(existing.files);
    const result = runExpertTotalAudit({ files, expertNum, maxScoreDiff, password: String(password || '') });
    const config = { expertNum, maxScoreDiff };
    const updated_at = upsertScoreCheckRow(batchId, 'total', { files: existing.files, config, result });
    res.json({
      files: existing.files.map((f) => ({ originalName: f.originalName, size: f.size })),
      config,
      result,
      updated_at,
    });
  } catch (err) {
    res.status(400).json({ error: (err as Error).message });
  }
});

// DELETE /api/evaluation/batch/:batchId/score-check/total
evaluationRouter.delete('/batch/:batchId/score-check/total', (req, res) => {
  const { batchId } = req.params;
  const targetDir = path.join(UPLOADS_DIR, batchId, 'evaluation', 'score_total');
  if (fs.existsSync(targetDir)) fs.rmSync(targetDir, { recursive: true, force: true });
  db.prepare('DELETE FROM eval_score_checks WHERE batch_id = ? AND check_type = ?').run(batchId, 'total');
  res.json({ ok: true });
});

// GET /api/evaluation/batch/:batchId/score-check/total/export — 专家总分核验.xlsx
evaluationRouter.get('/batch/:batchId/score-check/total/export', async (req, res) => {
  try {
    const { batchId } = req.params;
    const batch = db.prepare('SELECT code FROM batches WHERE id = ?').get(batchId) as { code?: string } | undefined;
    if (!batch) return res.status(404).json({ error: '未找到该批次' });
    const existing = getScoreCheckRow(batchId, 'total');
    const result = existing?.result as ExpertTotalAuditResult | null;
    if (!result) return res.status(400).json({ error: '当前批次尚未执行专家总分校验' });

    const wb = new ExcelJS.Workbook();
    wb.creator = '国网招标代理 SOP 工作台';
    appendStyledAuditSheet(
      wb,
      '核验结果',
      ['分标名称', '包名称', '投标人', '超过半数同分', '分差超过阈值', '来源文件', '投标人位置'],
      [26, 12, 32, 38, 48, 34, 18],
      result.items.map((it) => [
        it.section,
        it.package_name,
        it.supplier,
        it.majority_same,
        it.diff_exceeded,
        it.source_file,
        it.supplier_location,
      ]),
      new Set([2, 7])
    );

    const safeCode = String(batch.code || 'BATCH').replace(/[^\w\u4e00-\u9fa5-]/g, '_');
    const filename = `${safeCode}_专家总分核验.xlsx`;
    const buf = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.send(Buffer.from(buf));
  } catch (err) {
    res.status(500).json({ error: `导出专家总分核验失败：${(err as Error).message}` });
  }
});

// POST /api/evaluation/batch/:batchId/score-check/objective/upload — upload & discover scoring methods
evaluationRouter.post('/batch/:batchId/score-check/objective/upload', scoreMemoryUpload.array('files', 100), (req, res) => {
  try {
    const { batchId } = req.params;
    const batch = db.prepare('SELECT id FROM batches WHERE id = ?').get(batchId);
    if (!batch) return res.status(404).json({ error: '未找到该批次' });

    const uploaded = (req.files as Express.Multer.File[] | undefined) || [];
    if (uploaded.length === 0) return res.status(400).json({ error: '请选择专家评分明细表（.xlsx / .xls / .zip）' });

    const scoreFiles: ScoreUploadedFile[] = uploaded.map((f) => ({
      originalName: decodeName(f.originalname),
      buffer: f.buffer,
    }));
    const { expandedFiles, methods } = discoverObjectiveMethods(scoreFiles);
    const metas = saveExpandedScoreFiles(batchId, 'score_objective', expandedFiles);
    const config = { selections: [], sectionOverride: '' };
    const updated_at = upsertScoreCheckRow(batchId, 'objective', {
      files: metas,
      config,
      methods,
      result: null,
    });

    res.status(201).json({
      files: metas.map((f) => ({ originalName: f.originalName, size: f.size })),
      config,
      methods,
      result: null,
      updated_at,
    });
  } catch (err) {
    res.status(400).json({ error: (err as Error).message });
  }
});

// POST /api/evaluation/batch/:batchId/score-check/objective/run — run 专家客观分校验
evaluationRouter.post('/batch/:batchId/score-check/objective/run', (req, res) => {
  try {
    const { batchId } = req.params;
    const existing = getScoreCheckRow(batchId, 'objective');
    if (!existing || existing.files.length === 0) {
      return res.status(400).json({ error: '尚未上传专家评分明细表，请先上传文件' });
    }
    const { selections, sectionOverride } = req.body as {
      selections?: ObjectiveSelection[];
      sectionOverride?: string;
    };
    const files = loadPersistedScoreFiles(existing.files);
    const result = runExpertObjectiveAudit({
      files,
      selections: selections || [],
      sectionOverride: String(sectionOverride || ''),
    });
    const config = { selections: selections || [], sectionOverride: String(sectionOverride || '') };
    const updated_at = upsertScoreCheckRow(batchId, 'objective', {
      files: existing.files,
      config,
      methods: existing.methods,
      result,
    });
    res.json({
      files: existing.files.map((f) => ({ originalName: f.originalName, size: f.size })),
      config,
      methods: existing.methods,
      result,
      updated_at,
    });
  } catch (err) {
    res.status(400).json({ error: (err as Error).message });
  }
});

// DELETE /api/evaluation/batch/:batchId/score-check/objective
evaluationRouter.delete('/batch/:batchId/score-check/objective', (req, res) => {
  const { batchId } = req.params;
  const targetDir = path.join(UPLOADS_DIR, batchId, 'evaluation', 'score_objective');
  if (fs.existsSync(targetDir)) fs.rmSync(targetDir, { recursive: true, force: true });
  db.prepare('DELETE FROM eval_score_checks WHERE batch_id = ? AND check_type = ?').run(batchId, 'objective');
  res.json({ ok: true });
});

// GET /api/evaluation/batch/:batchId/score-check/objective/export — 专家客观分核验.xlsx
evaluationRouter.get('/batch/:batchId/score-check/objective/export', async (req, res) => {
  try {
    const { batchId } = req.params;
    const batch = db.prepare('SELECT code FROM batches WHERE id = ?').get(batchId) as { code?: string } | undefined;
    if (!batch) return res.status(404).json({ error: '未找到该批次' });
    const existing = getScoreCheckRow(batchId, 'objective');
    const result = existing?.result as ExpertObjectiveAuditResult | null;
    if (!result) return res.status(400).json({ error: '当前批次尚未执行专家客观分校验' });

    const wb = new ExcelJS.Workbook();
    wb.creator = '国网招标代理 SOP 工作台';
    appendStyledAuditSheet(
      wb,
      '核验结果',
      ['分标名称', '包名称', '投标人', '客观评分细则', '问题类型', '说明', '来源位置'],
      [24, 12, 30, 38, 16, 44, 42],
      result.items.map((it) => [
        it.section,
        it.package_name,
        it.supplier,
        it.criterion,
        it.issue_type,
        it.detail,
        it.source_location,
      ]),
      new Set([2, 5])
    );

    const safeCode = String(batch.code || 'BATCH').replace(/[^\w\u4e00-\u9fa5-]/g, '_');
    const filename = `${safeCode}_专家客观分核验.xlsx`;
    const buf = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.send(Buffer.from(buf));
  } catch (err) {
    res.status(500).json({ error: `导出专家客观分核验失败：${(err as Error).message}` });
  }
});

// ==================== 专家打分预校验 (线下打分模板预校验：总分 + 客观分) ====================

function getBatchPrecheckRejections(batchId: string): PrecheckRejectionRecord[] {
  const rows = db
    .prepare(
      'SELECT lot_code, lot_name, package_no, supplier_name, review_type, reason_category FROM eval_rejections WHERE batch_id = ?'
    )
    .all(batchId) as Array<Record<string, unknown>>;
  return rows.map((r) => ({
    lot_code: String(r.lot_code || ''),
    lot_name: String(r.lot_name || ''),
    package_no: String(r.package_no || ''),
    supplier_name: String(r.supplier_name || ''),
    review_type: String(r.review_type || ''),
    reason_category: String(r.reason_category || ''),
  }));
}

function hasBatchRejectionFile(batchId: string): boolean {
  const row = db
    .prepare(
      "SELECT 1 FROM uploaded_files WHERE batch_id = ? AND module = 'evaluation' AND slot_key = 'rejections' LIMIT 1"
    )
    .get(batchId);
  return Boolean(row);
}

// GET /api/evaluation/batch/:batchId/score-precheck
evaluationRouter.get('/batch/:batchId/score-precheck', (req, res) => {
  const { batchId } = req.params;
  const batch = db.prepare('SELECT id FROM batches WHERE id = ?').get(batchId);
  if (!batch) return res.status(404).json({ error: '未找到该批次' });
  const batchRejectionCount = getBatchPrecheckRejections(batchId).length;
  const hasRejectionFile = hasBatchRejectionFile(batchId) || batchRejectionCount > 0;
  const row = getScoreCheckRow(batchId, 'precheck');
  if (!row) {
    return res.json({
      files: [],
      templateFile: null,
      sheets: [],
      batchRejectionCount,
      hasRejectionFile,
      config: { maxScoreDiff: 8, excludeRejected: true, objectiveSelections: [] },
      result: null,
      updated_at: '',
    });
  }
  const cfg = row.config as {
    templateName?: string;
    maxScoreDiff?: number;
    excludeRejected?: boolean;
    objectiveSelections?: PrecheckObjectiveSelection[];
  };
  const templateName = String(cfg?.templateName || '');
  res.json({
    files: row.files.map((f) => ({ originalName: f.originalName, size: f.size })),
    templateFile: templateName ? { originalName: templateName } : null,
    sheets: row.methods as PrecheckSheetMeta[],
    batchRejectionCount,
    hasRejectionFile,
    config: {
      maxScoreDiff: cfg?.maxScoreDiff ?? 8,
      excludeRejected: cfg?.excludeRejected !== false,
      objectiveSelections: cfg?.objectiveSelections || [],
    },
    result: row.result as ExpertPrecheckResult | null,
    updated_at: row.updated_at,
  });
});

// POST /api/evaluation/batch/:batchId/score-precheck/upload — upload expert workbooks (+ optional template) & run precheck
evaluationRouter.post(
  '/batch/:batchId/score-precheck/upload',
  scoreMemoryUpload.fields([
    { name: 'experts', maxCount: 100 },
    { name: 'template', maxCount: 1 },
  ]),
  async (req, res) => {
    try {
      const { batchId } = req.params;
      const batch = db.prepare('SELECT id FROM batches WHERE id = ?').get(batchId);
      if (!batch) return res.status(404).json({ error: '未找到该批次' });

      const fieldFiles = (req.files || {}) as Record<string, Express.Multer.File[]>;
      const uploadedExperts = fieldFiles.experts || [];
      if (uploadedExperts.length === 0) {
        return res.status(400).json({ error: '请至少选择一份专家线下打分表（.xlsx / .xls / .zip）' });
      }
      const uploadedTemplate = fieldFiles.template?.[0] || null;

      const rawDiff = String(req.body?.maxScoreDiff ?? '8').trim();
      const maxScoreDiff = rawDiff === '' ? 8 : Number(rawDiff);
      const excludeRejected = String(req.body?.excludeRejected ?? 'true') !== 'false';
      const isAppend = String(req.body?.append ?? '') === 'true';
      const batchRejections = getBatchPrecheckRejections(batchId);

      const newExpertFiles: ScoreUploadedFile[] = uploadedExperts.map((f) => ({
        originalName: decodeName(f.originalname),
        buffer: f.buffer,
      }));
      let expertFiles: ScoreUploadedFile[] = newExpertFiles;
      let prevTemplateName = '';
      let prevSelections: PrecheckObjectiveSelection[] | null = null;

      if (isAppend) {
        const existing = getScoreCheckRow(batchId, 'precheck');
        if (existing && existing.files.length > 0) {
          const prevLoaded = loadPersistedScoreFiles(existing.files);
          const newNameSet = new Set(newExpertFiles.map((f) => f.originalName));
          expertFiles = [
            ...prevLoaded.filter((f) => !newNameSet.has(f.originalName)),
            ...newExpertFiles,
          ];
          const prevCfg = existing.config as {
            templateName?: string;
            objectiveSelections?: PrecheckObjectiveSelection[];
          };
          prevTemplateName = String(prevCfg?.templateName || '');
          if (Array.isArray(prevCfg?.objectiveSelections) && prevCfg.objectiveSelections.length > 0) {
            prevSelections = prevCfg.objectiveSelections;
          }
        }
      }

      const templateFile: ScoreUploadedFile | null = uploadedTemplate
        ? { originalName: decodeName(uploadedTemplate.originalname), buffer: uploadedTemplate.buffer }
        : null;

      const { expandedExperts, sheets, defaultSelections } = await discoverPrecheckMetadata({
        expertFiles,
        templateFile,
      });

      const effectiveSelections =
        isAppend && !templateFile && prevSelections ? prevSelections : defaultSelections;
      const effectiveTemplateName = templateFile ? templateFile.originalName : prevTemplateName;

      const result = runExpertPrecheckAudit({
        expertFiles: expandedExperts,
        sheetNames: sheets.map((s) => s.sheetName),
        maxScoreDiff,
        objectiveSelections: effectiveSelections,
        excludeRejected,
        rejections: batchRejections,
      });

      const metas = saveExpandedScoreFiles(batchId, 'score_precheck', expandedExperts);
      const config = {
        maxScoreDiff,
        excludeRejected,
        objectiveSelections: effectiveSelections,
        templateName: effectiveTemplateName,
      };
      const updated_at = upsertScoreCheckRow(batchId, 'precheck', {
        files: metas,
        config,
        methods: sheets,
        result,
      });

      res.status(201).json({
        files: metas.map((f) => ({ originalName: f.originalName, size: f.size })),
        templateFile: templateFile ? { originalName: templateFile.originalName } : null,
        sheets,
        batchRejectionCount: batchRejections.length,
        hasRejectionFile: hasBatchRejectionFile(batchId) || batchRejections.length > 0,
        config: { maxScoreDiff, excludeRejected, objectiveSelections: defaultSelections },
        result,
        updated_at,
      });
    } catch (err) {
      res.status(400).json({ error: (err as Error).message });
    }
  }
);

// POST /api/evaluation/batch/:batchId/score-precheck/rerun — rerun with updated threshold, excludeRejected & objective selections
evaluationRouter.post('/batch/:batchId/score-precheck/rerun', (req, res) => {
  try {
    const { batchId } = req.params;
    const existing = getScoreCheckRow(batchId, 'precheck');
    if (!existing || existing.files.length === 0) {
      return res.status(400).json({ error: '尚未上传专家线下打分表，请先上传文件' });
    }
    const files = loadPersistedScoreFiles(existing.files);
    if (files.length !== existing.files.length) {
      return res.status(400).json({ error: '已保存的专家线下打分表原文件缺失，请重新上传后再校验' });
    }

    const body = req.body as {
      maxScoreDiff?: number | string;
      excludeRejected?: boolean;
      objectiveSelections?: PrecheckObjectiveSelection[];
    };
    const rawDiff = body.maxScoreDiff !== undefined ? String(body.maxScoreDiff).trim() : '';
    const maxScoreDiff = rawDiff === '' ? 8 : Number(rawDiff);
    const excludeRejected = body.excludeRejected !== false;
    const objectiveSelections = Array.isArray(body.objectiveSelections) ? body.objectiveSelections : [];
    const sheets = existing.methods as PrecheckSheetMeta[];
    const batchRejections = getBatchPrecheckRejections(batchId);

    const result = runExpertPrecheckAudit({
      expertFiles: files,
      sheetNames: sheets.map((s) => s.sheetName),
      maxScoreDiff,
      objectiveSelections,
      excludeRejected,
      rejections: batchRejections,
    });

    const templateName = String((existing.config as { templateName?: string })?.templateName || '');
    const config = { maxScoreDiff, excludeRejected, objectiveSelections, templateName };
    const updated_at = upsertScoreCheckRow(batchId, 'precheck', {
      files: existing.files,
      config,
      methods: sheets,
      result,
    });

    res.json({
      files: existing.files.map((f) => ({ originalName: f.originalName, size: f.size })),
      templateFile: templateName ? { originalName: templateName } : null,
      sheets,
      batchRejectionCount: batchRejections.length,
      hasRejectionFile: hasBatchRejectionFile(batchId) || batchRejections.length > 0,
      config: { maxScoreDiff, excludeRejected, objectiveSelections },
      result,
      updated_at,
    });
  } catch (err) {
    res.status(400).json({ error: (err as Error).message });
  }
});

// POST /api/evaluation/batch/:batchId/score-precheck/remove-files — remove selected uploaded expert files & rerun
evaluationRouter.post('/batch/:batchId/score-precheck/remove-files', (req, res) => {
  try {
    const { batchId } = req.params;
    const existing = getScoreCheckRow(batchId, 'precheck');
    if (!existing || existing.files.length === 0) {
      return res.status(400).json({ error: '当前批次暂无已上传的专家线下打分表' });
    }
    const body = req.body as {
      fileNames?: string[];
      maxScoreDiff?: number | string;
      excludeRejected?: boolean;
      objectiveSelections?: PrecheckObjectiveSelection[];
    };
    const removeSet = new Set((body.fileNames || []).map((s) => String(s)));
    if (removeSet.size === 0) {
      return res.status(400).json({ error: '请先选择要删除的文件' });
    }

    const removedMetas = existing.files.filter((f) => removeSet.has(f.originalName));
    const remainingMetas = existing.files.filter((f) => !removeSet.has(f.originalName));
    for (const rm of removedMetas) {
      const absPath = path.join(UPLOADS_DIR, rm.storedPath);
      if (fs.existsSync(absPath)) fs.unlinkSync(absPath);
    }

    const batchRejections = getBatchPrecheckRejections(batchId);
    const hasRejectionFile = hasBatchRejectionFile(batchId) || batchRejections.length > 0;

    if (remainingMetas.length === 0) {
      const targetDir = path.join(UPLOADS_DIR, batchId, 'evaluation', 'score_precheck');
      if (fs.existsSync(targetDir)) fs.rmSync(targetDir, { recursive: true, force: true });
      db.prepare('DELETE FROM eval_score_checks WHERE batch_id = ? AND check_type = ?').run(batchId, 'precheck');
      return res.json({
        files: [],
        templateFile: null,
        sheets: [],
        batchRejectionCount: batchRejections.length,
        hasRejectionFile,
        config: { maxScoreDiff: 8, excludeRejected: true, objectiveSelections: [] },
        result: null,
        updated_at: '',
      });
    }

    const remainingFiles = loadPersistedScoreFiles(remainingMetas);
    if (remainingFiles.length !== remainingMetas.length) {
      return res.status(400).json({ error: '剩余专家打分表原文件缺失，请重新上传' });
    }

    const prevCfg = existing.config as {
      templateName?: string;
      maxScoreDiff?: number;
      excludeRejected?: boolean;
      objectiveSelections?: PrecheckObjectiveSelection[];
    };
    const rawDiff = body.maxScoreDiff !== undefined ? String(body.maxScoreDiff).trim() : '';
    const maxScoreDiff = rawDiff === '' ? (prevCfg?.maxScoreDiff ?? 8) : Number(rawDiff);
    const excludeRejected =
      body.excludeRejected !== undefined ? body.excludeRejected : prevCfg?.excludeRejected !== false;
    const objectiveSelections = Array.isArray(body.objectiveSelections)
      ? body.objectiveSelections
      : prevCfg?.objectiveSelections || [];
    const sheets = existing.methods as PrecheckSheetMeta[];

    const result = runExpertPrecheckAudit({
      expertFiles: remainingFiles,
      sheetNames: sheets.map((s) => s.sheetName),
      maxScoreDiff,
      objectiveSelections,
      excludeRejected,
      rejections: batchRejections,
    });

    const templateName = String(prevCfg?.templateName || '');
    const config = { maxScoreDiff, excludeRejected, objectiveSelections, templateName };
    const updated_at = upsertScoreCheckRow(batchId, 'precheck', {
      files: remainingMetas,
      config,
      methods: sheets,
      result,
    });

    res.json({
      files: remainingMetas.map((f) => ({ originalName: f.originalName, size: f.size })),
      templateFile: templateName ? { originalName: templateName } : null,
      sheets,
      batchRejectionCount: batchRejections.length,
      hasRejectionFile,
      config: { maxScoreDiff, excludeRejected, objectiveSelections },
      result,
      updated_at,
    });
  } catch (err) {
    res.status(400).json({ error: (err as Error).message });
  }
});

// DELETE /api/evaluation/batch/:batchId/score-precheck
evaluationRouter.delete('/batch/:batchId/score-precheck', (req, res) => {
  const { batchId } = req.params;
  const targetDir = path.join(UPLOADS_DIR, batchId, 'evaluation', 'score_precheck');
  if (fs.existsSync(targetDir)) fs.rmSync(targetDir, { recursive: true, force: true });
  db.prepare('DELETE FROM eval_score_checks WHERE batch_id = ? AND check_type = ?').run(batchId, 'precheck');
  res.json({ ok: true });
});

// GET /api/evaluation/batch/:batchId/score-precheck/export — 专家打分预校验.xlsx
evaluationRouter.get('/batch/:batchId/score-precheck/export', async (req, res) => {
  try {
    const { batchId } = req.params;
    const batch = db.prepare('SELECT code FROM batches WHERE id = ?').get(batchId) as { code?: string } | undefined;
    if (!batch) return res.status(404).json({ error: '未找到该批次' });
    const existing = getScoreCheckRow(batchId, 'precheck');
    const result = existing?.result as ExpertPrecheckResult | null;
    if (!result) return res.status(400).json({ error: '当前批次尚未执行专家打分预校验' });

    const wb = new ExcelJS.Workbook();
    wb.creator = '国网招标代理 SOP 工作台';

    appendStyledAuditSheet(
      wb,
      '总分预校验结果',
      ['工作表', '标段名称', '包号', '应答人', '超过半数同分', '分差超过阈值', '各专家总分', '来源位置'],
      [16, 24, 12, 30, 38, 48, 36, 36],
      result.totalItems.map((it) => [
        it.sheet_name,
        it.section,
        it.package_name,
        it.supplier,
        it.majority_same,
        it.diff_exceeded,
        it.expert_scores,
        it.source_location,
      ]),
      new Set([1, 3])
    );

    appendStyledAuditSheet(
      wb,
      '客观分预校验结果',
      ['工作表', '标段名称', '包号', '应答人', '客观评分细则', '问题类型', '各专家打分说明', '来源位置'],
      [16, 24, 12, 30, 32, 16, 42, 42],
      result.objectiveItems.map((it) => [
        it.sheet_name,
        it.section,
        it.package_name,
        it.supplier,
        it.criterion,
        it.issue_type,
        it.detail,
        it.source_location,
      ]),
      new Set([1, 3, 6])
    );

    if (result.excludedRejectedItems && result.excludedRejectedItems.length > 0) {
      appendStyledAuditSheet(
        wb,
        '已排除否决供应商',
        ['工作表', '标段名称', '包号', '应答人', '否决类别', '未进入详评原因'],
        [16, 24, 12, 32, 18, 42],
        result.excludedRejectedItems.map((it) => [
          it.sheet_name,
          it.section,
          it.package_name,
          it.supplier,
          it.review_type,
          it.reason_category,
        ]),
        new Set([1, 3, 5])
      );
    }

    const safeCode = String(batch.code || 'BATCH').replace(/[^\w\u4e00-\u9fa5-]/g, '_');
    const filename = `${safeCode}_专家打分预校验.xlsx`;
    const buf = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.send(Buffer.from(buf));
  } catch (err) {
    res.status(500).json({ error: `导出专家打分预校验失败：${(err as Error).message}` });
  }
});

// ==================== 工具 6：初评汇报生成 (Excel 模板 -> 初评汇报单.docx) ====================

// GET /api/evaluation/batch/:batchId/initial-report
evaluationRouter.get('/batch/:batchId/initial-report', (req, res) => {
  const { batchId } = req.params;
  const batch = db.prepare('SELECT id FROM batches WHERE id = ?').get(batchId);
  if (!batch) return res.status(404).json({ error: '未找到该批次' });
  const row = getScoreCheckRow(batchId, 'initial_report');
  if (!row) {
    return res.json({
      file: null,
      config: {
        department: '',
        clarification: '',
        simplifyNames: true,
        companySuffixes: DEFAULT_COMPANY_SUFFIXES,
      },
      result: null,
      updated_at: '',
    });
  }
  res.json({
    file: row.files[0] ? { originalName: row.files[0].originalName, size: row.files[0].size } : null,
    config: row.config,
    result: row.result as InitialReportResult | null,
    updated_at: row.updated_at,
  });
});

// POST /api/evaluation/batch/:batchId/initial-report/upload
evaluationRouter.post('/batch/:batchId/initial-report/upload', scoreMemoryUpload.single('file'), (req, res) => {
  try {
    const { batchId } = req.params;
    const batch = db.prepare('SELECT id FROM batches WHERE id = ?').get(batchId);
    if (!batch) return res.status(404).json({ error: '未找到该批次' });
    if (!req.file) return res.status(400).json({ error: '请选择初评情况汇报模板 Excel 文件（.xlsx / .xls）' });

    const originalName = decodeName(req.file.originalname);
    const department = String(req.body?.department ?? '').trim();
    const clarification = String(req.body?.clarification ?? '').trim();
    const simplifyNames = String(req.body?.simplifyNames ?? 'true') !== 'false';
    const companySuffixes = req.body?.companySuffixes !== undefined ? String(req.body.companySuffixes) : DEFAULT_COMPANY_SUFFIXES;

    const result = generateInitialReportData(req.file.buffer, originalName, {
      department,
      clarification,
      simplifyNames,
      companySuffixes,
    });
    const metas = saveExpandedScoreFiles(batchId, 'initial_report', [{ originalName, buffer: req.file.buffer }]);
    const config = { department, clarification, simplifyNames, companySuffixes };
    const updated_at = upsertScoreCheckRow(batchId, 'initial_report', { files: metas, config, result });

    res.status(201).json({
      file: metas[0] ? { originalName: metas[0].originalName, size: metas[0].size } : null,
      config,
      result,
      updated_at,
    });
  } catch (err) {
    res.status(400).json({ error: (err as Error).message });
  }
});

// POST /api/evaluation/batch/:batchId/initial-report/regenerate
evaluationRouter.post('/batch/:batchId/initial-report/regenerate', (req, res) => {
  try {
    const { batchId } = req.params;
    const existing = getScoreCheckRow(batchId, 'initial_report');
    if (!existing || existing.files.length === 0) {
      return res.status(400).json({ error: '尚未上传初评情况汇报模板，请先上传 Excel 文件' });
    }
    const { department, clarification, simplifyNames, companySuffixes } = req.body as {
      department?: string;
      clarification?: string;
      simplifyNames?: boolean;
      companySuffixes?: string;
    };
    const files = loadPersistedScoreFiles(existing.files);
    if (files.length === 0) {
      return res.status(400).json({ error: '初评模板原文件已丢失，请重新上传' });
    }
    const config = {
      department: String(department ?? '').trim(),
      clarification: String(clarification ?? '').trim(),
      simplifyNames: simplifyNames !== undefined ? Boolean(simplifyNames) : true,
      companySuffixes: companySuffixes !== undefined ? String(companySuffixes) : DEFAULT_COMPANY_SUFFIXES,
    };
    const result = generateInitialReportData(files[0].buffer, files[0].originalName, config);
    const updated_at = upsertScoreCheckRow(batchId, 'initial_report', { files: existing.files, config, result });
    res.json({
      file: existing.files[0] ? { originalName: existing.files[0].originalName, size: existing.files[0].size } : null,
      config,
      result,
      updated_at,
    });
  } catch (err) {
    res.status(400).json({ error: (err as Error).message });
  }
});

// DELETE /api/evaluation/batch/:batchId/initial-report
evaluationRouter.delete('/batch/:batchId/initial-report', (req, res) => {
  const { batchId } = req.params;
  const targetDir = path.join(UPLOADS_DIR, batchId, 'evaluation', 'initial_report');
  if (fs.existsSync(targetDir)) fs.rmSync(targetDir, { recursive: true, force: true });
  db.prepare('DELETE FROM eval_score_checks WHERE batch_id = ? AND check_type = ?').run(batchId, 'initial_report');
  res.json({ ok: true });
});

// GET /api/evaluation/batch/:batchId/initial-report/export — 初评汇报单.docx
evaluationRouter.get('/batch/:batchId/initial-report/export', (req, res) => {
  try {
    const { batchId } = req.params;
    const batch = db.prepare('SELECT code FROM batches WHERE id = ?').get(batchId) as { code?: string } | undefined;
    if (!batch) return res.status(404).json({ error: '未找到该批次' });
    const existing = getScoreCheckRow(batchId, 'initial_report');
    const result = existing?.result as InitialReportResult | null;
    if (!result) return res.status(400).json({ error: '当前批次尚未生成初评汇报单' });

    const buf = buildInitialReportDocxBuffer(result.paragraphs);
    const safeCode = String(batch.code || 'BATCH').replace(/[^\w\u4e00-\u9fa5-]/g, '_');
    const filename = `${safeCode}_初评汇报单.docx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.send(buf);
  } catch (err) {
    res.status(500).json({ error: `导出初评汇报单失败：${(err as Error).message}` });
  }
});

// ==================== 工具 1：预警事项确认单 (风险提示表 + 标包分组表 + 开标表 -> Word / ZIP) ====================

// GET /api/evaluation/batch/:batchId/warning-form
evaluationRouter.get('/batch/:batchId/warning-form', (req, res) => {
  const { batchId } = req.params;
  const batch = db.prepare('SELECT id FROM batches WHERE id = ?').get(batchId);
  if (!batch) return res.status(404).json({ error: '未找到该批次' });
  const row = getScoreCheckRow(batchId, 'warning_form');
  if (!row) {
    return res.json({
      files: { risk: null, groups: null, bids: null },
      config: { date: '' },
      result: null,
      updated_at: '',
    });
  }
  const fRisk = row.files[0] ? { originalName: row.files[0].originalName, size: row.files[0].size } : null;
  const fGroups = row.files[1] ? { originalName: row.files[1].originalName, size: row.files[1].size } : null;
  const fBids = row.files[2] ? { originalName: row.files[2].originalName, size: row.files[2].size } : null;
  res.json({
    files: { risk: fRisk, groups: fGroups, bids: fBids },
    config: row.config,
    result: row.result as WarningFormAuditResult | null,
    updated_at: row.updated_at,
  });
});

// POST /api/evaluation/batch/:batchId/warning-form/run
evaluationRouter.post(
  '/batch/:batchId/warning-form/run',
  scoreMemoryUpload.fields([
    { name: 'risk', maxCount: 1 },
    { name: 'groups', maxCount: 1 },
    { name: 'bids', maxCount: 1 },
  ]),
  (req, res) => {
    try {
      const { batchId } = req.params;
      const batch = db.prepare('SELECT id FROM batches WHERE id = ?').get(batchId);
      if (!batch) return res.status(404).json({ error: '未找到该批次' });

      const fieldFiles = (req.files || {}) as Record<string, Express.Multer.File[]>;
      const existing = getScoreCheckRow(batchId, 'warning_form');
      const prevFiles = existing ? loadPersistedScoreFiles(existing.files) : [];

      const riskFile: ScoreUploadedFile | null = fieldFiles.risk?.[0]
        ? { originalName: decodeName(fieldFiles.risk[0].originalname), buffer: fieldFiles.risk[0].buffer }
        : prevFiles[0] || null;
      const groupFile: ScoreUploadedFile | null = fieldFiles.groups?.[0]
        ? { originalName: decodeName(fieldFiles.groups[0].originalname), buffer: fieldFiles.groups[0].buffer }
        : prevFiles[1] || null;
      const bidFile: ScoreUploadedFile | null = fieldFiles.bids?.[0]
        ? { originalName: decodeName(fieldFiles.bids[0].originalname), buffer: fieldFiles.bids[0].buffer }
        : prevFiles[2] || null;

      if (!riskFile) return res.status(400).json({ error: '请先上传《围标串标风险提示表》' });
      if (!groupFile) return res.status(400).json({ error: '请先上传《标包分组表》' });
      if (!bidFile) return res.status(400).json({ error: '请先上传《开标项目信息表》' });

      const date = String(req.body?.date ?? '').trim();
      const result = runWarningFormAudit({ riskFile, groupFile, bidFile, date });
      const metas = saveExpandedScoreFiles(batchId, 'warning_form', [riskFile, groupFile, bidFile]);
      const config = { date };
      const updated_at = upsertScoreCheckRow(batchId, 'warning_form', { files: metas, config, result });

      res.status(201).json({
        files: {
          risk: { originalName: metas[0].originalName, size: metas[0].size },
          groups: { originalName: metas[1].originalName, size: metas[1].size },
          bids: { originalName: metas[2].originalName, size: metas[2].size },
        },
        config,
        result,
        updated_at,
      });
    } catch (err) {
      res.status(400).json({ error: (err as Error).message });
    }
  }
);

// POST /api/evaluation/batch/:batchId/warning-form/rerun
evaluationRouter.post('/batch/:batchId/warning-form/rerun', (req, res) => {
  try {
    const { batchId } = req.params;
    const existing = getScoreCheckRow(batchId, 'warning_form');
    if (!existing || existing.files.length < 3) {
      return res.status(400).json({ error: '请先完整上传风险提示表、标包分组表与开标项目信息表' });
    }
    const files = loadPersistedScoreFiles(existing.files);
    if (files.length < 3) {
      return res.status(400).json({ error: '已上传原表文件不完整，请重新上传' });
    }
    const date = String((req.body as { date?: string })?.date ?? '').trim();
    const result = runWarningFormAudit({
      riskFile: files[0],
      groupFile: files[1],
      bidFile: files[2],
      date,
    });
    const config = { date };
    const updated_at = upsertScoreCheckRow(batchId, 'warning_form', { files: existing.files, config, result });
    res.json({
      files: {
        risk: { originalName: existing.files[0].originalName, size: existing.files[0].size },
        groups: { originalName: existing.files[1].originalName, size: existing.files[1].size },
        bids: { originalName: existing.files[2].originalName, size: existing.files[2].size },
      },
      config,
      result,
      updated_at,
    });
  } catch (err) {
    res.status(400).json({ error: (err as Error).message });
  }
});

// DELETE /api/evaluation/batch/:batchId/warning-form
evaluationRouter.delete('/batch/:batchId/warning-form', (req, res) => {
  const { batchId } = req.params;
  const targetDir = path.join(UPLOADS_DIR, batchId, 'evaluation', 'warning_form');
  if (fs.existsSync(targetDir)) fs.rmSync(targetDir, { recursive: true, force: true });
  db.prepare('DELETE FROM eval_score_checks WHERE batch_id = ? AND check_type = ?').run(batchId, 'warning_form');
  res.json({ ok: true });
});

// GET /api/evaluation/batch/:batchId/warning-form/export — 预警事项确认单.zip
evaluationRouter.get('/batch/:batchId/warning-form/export', (req, res) => {
  try {
    const { batchId } = req.params;
    const batch = db.prepare('SELECT code FROM batches WHERE id = ?').get(batchId) as { code?: string } | undefined;
    if (!batch) return res.status(404).json({ error: '未找到该批次' });
    const existing = getScoreCheckRow(batchId, 'warning_form');
    const result = existing?.result as WarningFormAuditResult | null;
    if (!result) return res.status(400).json({ error: '当前批次尚未生成预警事项确认单' });

    const buf = buildWarningFormsZipBuffer(result);
    const safeCode = String(batch.code || 'BATCH').replace(/[^\w\u4e00-\u9fa5-]/g, '_');
    const filename = `${safeCode}_预警事项确认单.zip`;
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.send(buf);
  } catch (err) {
    res.status(500).json({ error: `导出预警事项确认单失败：${(err as Error).message}` });
  }
});

// GET /api/evaluation/batch/:batchId/warning-form/export-single/:idx — single .docx
evaluationRouter.get('/batch/:batchId/warning-form/export-single/:idx', (req, res) => {
  try {
    const { batchId, idx } = req.params;
    const existing = getScoreCheckRow(batchId, 'warning_form');
    const result = existing?.result as WarningFormAuditResult | null;
    if (!result) return res.status(400).json({ error: '当前批次尚未生成预警事项确认单' });
    const i = Number(idx);
    if (!Number.isInteger(i) || i < 0 || i >= result.forms.length) {
      return res.status(404).json({ error: '未找到对应确认单' });
    }
    const item = result.forms[i];
    const buf = renderWarningFormDocx(item);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(item.docxFileName)}`);
    res.send(buf);
  } catch (err) {
    res.status(500).json({ error: `导出单份确认单失败：${(err as Error).message}` });
  }
});

// ==================== 初评会议纪要生成 (公开 / 非公开 Word 模板) ====================

// GET /api/evaluation/batch/:batchId/initial-minutes
evaluationRouter.get('/batch/:batchId/initial-minutes', (req, res) => {
  const { batchId } = req.params;
  const batch = db.prepare('SELECT id, name, procurement_type FROM batches WHERE id = ?').get(batchId) as
    | { id: string; name: string; procurement_type?: string }
    | undefined;
  if (!batch) return res.status(404).json({ error: '未找到该批次' });
  const defaultType: MinutesProcurementType = batch.procurement_type === '公开招标' ? '公开' : '非公开';
  const row = getScoreCheckRow(batchId, 'initial_minutes');
  if (!row) {
    return res.json({
      file: null,
      config: {
        procurementType: defaultType,
        batchName: batch.name || '',
        meetingTime: formatDefaultMeetingTime(),
        groupName: '商务组',
        clarificationNote: defaultType === '公开' ? '本次评标无澄清事项。' : '本次评审无澄清事项。',
        reviewTypeFilter: 'all',
      },
      result: null,
      updated_at: '',
    });
  }
  res.json({
    file: row.files[0] ? { originalName: row.files[0].originalName, size: row.files[0].size } : null,
    config: row.config,
    result: row.result as InitialMinutesResult | null,
    updated_at: row.updated_at,
  });
});

// POST /api/evaluation/batch/:batchId/initial-minutes/upload
evaluationRouter.post('/batch/:batchId/initial-minutes/upload', scoreMemoryUpload.single('file'), (req, res) => {
  try {
    const { batchId } = req.params;
    const batch = db.prepare('SELECT id, name, procurement_type FROM batches WHERE id = ?').get(batchId) as
      | { id: string; name: string; procurement_type?: string }
      | undefined;
    if (!batch) return res.status(404).json({ error: '未找到该批次' });
    if (!req.file) return res.status(400).json({ error: '请先上传包含否决明细的 Excel 文件（《初评情况汇报模板》或《未进入详评情况统计表》）' });

    const originalName = decodeName(req.file.originalname);
    const parsed = parseMinutesExcel(req.file.buffer, originalName);

    const rawType = String(req.body?.procurementType ?? '').trim();
    const procurementType: MinutesProcurementType =
      rawType === '公开' || rawType === '非公开'
        ? rawType
        : batch.procurement_type === '公开招标'
          ? '公开'
          : '非公开';
    const batchName = String(req.body?.batchName ?? '').trim() || parsed.detectedBatchName || batch.name || '';
    const meetingTime = String(req.body?.meetingTime ?? '').trim() || formatDefaultMeetingTime();
    const groupName = String(req.body?.groupName ?? '').trim() || parsed.detectedGroupName || '商务组';
    const clarificationNote = String(req.body?.clarificationNote ?? '').trim();
    const rawFilter = String(req.body?.reviewTypeFilter ?? 'all').trim();
    const reviewTypeFilter: 'all' | '商务' | '技术' = rawFilter === '商务' || rawFilter === '技术' ? rawFilter : 'all';

    const config: InitialMinutesConfig = {
      procurementType,
      batchName,
      meetingTime,
      groupName,
      clarificationNote,
      reviewTypeFilter,
    };
    const result = buildMinutesResult(parsed, originalName, config);
    const metas = saveExpandedScoreFiles(batchId, 'initial_minutes', [{ originalName, buffer: req.file.buffer }]);
    const updated_at = upsertScoreCheckRow(batchId, 'initial_minutes', {
      files: metas,
      config: result.config as unknown as Record<string, unknown>,
      result,
    });

    res.status(201).json({
      file: metas[0] ? { originalName: metas[0].originalName, size: metas[0].size } : null,
      config: result.config,
      result,
      updated_at,
    });
  } catch (err) {
    res.status(400).json({ error: (err as Error).message });
  }
});

// POST /api/evaluation/batch/:batchId/initial-minutes/regenerate
evaluationRouter.post('/batch/:batchId/initial-minutes/regenerate', (req, res) => {
  try {
    const { batchId } = req.params;
    const existing = getScoreCheckRow(batchId, 'initial_minutes');
    if (!existing || existing.files.length === 0) {
      return res.status(400).json({ error: '尚未上传否决明细 Excel，请先上传文件' });
    }
    const files = loadPersistedScoreFiles(existing.files);
    if (files.length === 0) {
      return res.status(400).json({ error: '已上传的 Excel 原文件丢失，请重新上传' });
    }
    const body = req.body as Partial<InitialMinutesConfig>;
    const parsed = parseMinutesExcel(files[0].buffer, files[0].originalName);
    const procurementType: MinutesProcurementType = body.procurementType === '公开' ? '公开' : '非公开';
    const reviewTypeFilter: 'all' | '商务' | '技术' =
      body.reviewTypeFilter === '商务' || body.reviewTypeFilter === '技术' ? body.reviewTypeFilter : 'all';

    const config: InitialMinutesConfig = {
      procurementType,
      batchName: String(body.batchName ?? '').trim(),
      meetingTime: String(body.meetingTime ?? '').trim(),
      groupName: String(body.groupName ?? '').trim(),
      clarificationNote: String(body.clarificationNote ?? '').trim(),
      reviewTypeFilter,
    };
    const result = buildMinutesResult(parsed, files[0].originalName, config);
    const updated_at = upsertScoreCheckRow(batchId, 'initial_minutes', {
      files: existing.files,
      config: result.config as unknown as Record<string, unknown>,
      result,
    });
    res.json({
      file: existing.files[0] ? { originalName: existing.files[0].originalName, size: existing.files[0].size } : null,
      config: result.config,
      result,
      updated_at,
    });
  } catch (err) {
    res.status(400).json({ error: (err as Error).message });
  }
});

// DELETE /api/evaluation/batch/:batchId/initial-minutes
evaluationRouter.delete('/batch/:batchId/initial-minutes', (req, res) => {
  const { batchId } = req.params;
  const targetDir = path.join(UPLOADS_DIR, batchId, 'evaluation', 'initial_minutes');
  if (fs.existsSync(targetDir)) fs.rmSync(targetDir, { recursive: true, force: true });
  db.prepare('DELETE FROM eval_score_checks WHERE batch_id = ? AND check_type = ?').run(batchId, 'initial_minutes');
  res.json({ ok: true });
});

// GET /api/evaluation/batch/:batchId/initial-minutes/export — 初评会议纪要.docx
evaluationRouter.get('/batch/:batchId/initial-minutes/export', (req, res) => {
  try {
    const { batchId } = req.params;
    const batch = db.prepare('SELECT code FROM batches WHERE id = ?').get(batchId) as { code?: string } | undefined;
    if (!batch) return res.status(404).json({ error: '未找到该批次' });
    const existing = getScoreCheckRow(batchId, 'initial_minutes');
    const result = existing?.result as InitialMinutesResult | null;
    if (!result) return res.status(400).json({ error: '当前批次尚未生成初评会议纪要' });

    const buf = renderInitialMinutesDocx(result);
    const safeCode = String(batch.code || 'BATCH').replace(/[^\w\u4e00-\u9fa5-]/g, '_');
    const safeGrp = String(result.config.groupName || '评审组').replace(/[^\w\u4e00-\u9fa5-]/g, '_');
    const filename = `【${result.config.procurementType}】${safeCode}_初评会议纪要（${safeGrp}）.docx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.send(buf);
  } catch (err) {
    res.status(500).json({ error: `导出初评会议纪要失败：${(err as Error).message}` });
  }
});
