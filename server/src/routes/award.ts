import fs from 'node:fs';
import path from 'node:path';
import { Router } from 'express';
import multer from 'multer';
import { v4 as uuid } from 'uuid';
import { db, UPLOADS_DIR } from '../db.js';
import {
  generateAwardResultNoticeDocx,
  generateAwardResultTableExcel,
  generateNoticeHandoverExcel,
  generatePreAwardPublicityDocx,
  generateProjectHandoverExcel,
  generateServiceFeeExcel,
  type AwardBidderDetailRow,
  type AwardDeliverablePayload,
  type AwardPackageRow,
} from '../awardGenerator.js';
import { parseErpPlan, stripAllSpaces } from '../biddingParser.js';
import { importMaterialCatalogToDb, lookupMaterialDescMap } from '../materialCatalog.js';
import {
  calculateServiceFee,
  formatDiscountLabel,
  normalizeCategory,
  parseClientAgreementsExcel,
  type ServiceFeeCategory,
} from '../serviceFee.js';

export const awardRouter = Router();

const upload = multer({
  dest: path.join(UPLOADS_DIR, 'tmp'),
  limits: { fileSize: 20 * 1024 * 1024 },
});

const pinyinCollator = new Intl.Collator('zh-Hans-CN-u-co-pinyin', { numeric: true, sensitivity: 'base' });

function extractPkgNum(pkgNo: string): number {
  const m = String(pkgNo || '').match(/(\d+)/);
  return m ? parseInt(m[1], 10) : Number.MAX_SAFE_INTEGER;
}

export const AWARD_TEMPLATES = [
  {
    key: 'pre_award_publicity',
    title: '预成交公示（中标候选人公示）',
    file_type: 'docx',
    description: '按原模板保留固定条款正文，自动替换标黄批次名称、编号、日期及各包推荐候选人/流标状态。',
    sources: ['招标模块：批次名称、编号、分标包号', '开评标模块：第一推荐候选人与流标状态'],
  },
  {
    key: 'award_result_notice',
    title: '成交结果公告（中标结果公告）',
    file_type: 'docx',
    description: '按原模板保留固定公告正文，自动填充标黄批次信息、采购人名称、成交人与最终报价（万元/%）。',
    sources: ['定标配置：业主单位（采购人）', '开评标模块：成交人与最终报价'],
  },
  {
    key: 'notice_handover',
    title: '成交通知书交接单',
    file_type: 'xlsx',
    description: '标准 19 列交接单，自动关联成交厂家联系人/电话/邮箱/邮政地址、项目概算与精确服务费（元）。',
    sources: ['招标模块：供应商联系方式与概算', '定标配置：业主单位费率与精确服务费'],
  },
  {
    key: 'project_handover',
    title: '项目交接信息表',
    file_type: 'xlsx',
    description: '标准 34 列全量交接台账，含项目概况、批次经理、开标地点、成交价格、服务费及流标标记。',
    sources: ['招标模块：需求一览表项目概况与预算', '开评标+定标：成交结果与服务费'],
  },
  {
    key: 'award_result_table',
    title: '成交结果表（全量投标人明细）',
    file_type: 'xlsx',
    description: '标准 22 列明细表，自动拉取各包全部投标人（含未通过符合性者）、得分排序及后台留存的申报人/电话。',
    sources: ['招标模块：申报人/电话与供应商联系信息', '开评标模块：总分排序与否决明细'],
  },
  {
    key: 'service_fee_sheet',
    title: '服务费计算表（含 Excel 公式）',
    file_type: 'xlsx',
    description: '与原《服务费计算表格》算法 100% 一致，含折扣前/折扣后服务费、平台使用费及动态合计公式。',
    sources: ['客户协议清单：业主单位折扣率', '开评标模块：各包成交价格'],
  },
] as const;

interface ClientAgreementRow {
  id: string;
  group_name: string;
  client_name: string;
  contact_name: string;
  contact_phone: string;
  contract_expiry: string;
  discount_rate: number;
  discount_label: string;
  agreement_status: string;
  notes: string;
  sort_order: number;
}

function matchDefaultClient(clients: ClientAgreementRow[], batch: Record<string, unknown>): ClientAgreementRow | null {
  const savedId = String(batch.award_client_id || '');
  if (savedId) {
    const exact = clients.find((c) => c.id === savedId);
    if (exact) return exact;
  }
  const unitText = `${String(batch.project_unit || '')} ${String(batch.name || '')}`.replace(/\s+/g, '');
  if (!unitText) return null;

  // Exact or substring match
  const exactUnit = clients.find((c) => c.client_name === String(batch.project_unit || '').trim());
  if (exactUnit) return exactUnit;

  // Keyword fuzzy match (e.g., "综合能源" -> "国网安徽综合能源服务有限公司")
  const keywords = [
    { re: /综合能源/, target: '国网安徽综合能源服务有限公司' },
    { re: /明生/, target: '安徽明生电力投资集团有限公司' },
    { re: /电动汽车/, target: '国网安徽电动汽车服务有限公司' },
    { re: /信通|信息通信/, target: '国网安徽省电力有限公司信息通信分公司' },
    { re: /超高压/, target: '国网安徽省电力有限公司超高压分公司' },
    { re: /经研院|经济技术研究院/, target: '国网安徽省电力有限公司经济技术研究院' },
    { re: /电科院|电力科学研究院/, target: '国网安徽省电力有限公司电力科学研究院' },
    { re: /送变电/, target: '安徽送变电工程有限公司' },
    { re: /合肥供电/, target: '国网安徽省电力有限公司合肥供电公司（授权采购）' },
  ];
  for (const kw of keywords) {
    if (kw.re.test(unitText)) {
      const found = clients.find((c) => c.client_name === kw.target);
      if (found) return found;
    }
  }
  for (const c of clients) {
    const shortName = c.client_name
      .replace(/国网安徽省电力有限公司|国网安徽|有限公司|有限责任公司/g, '')
      .replace(/（.*）/, '');
    if (shortName.length >= 2 && unitText.includes(shortName)) return c;
  }
  return null;
}

function cleanSupplierKey(name: string): string {
  return String(name || '').replace(/[\s\u3000]+/g, '').trim();
}

function normalizePkgToken(pkg: string): string {
  const s = cleanSupplierKey(pkg);
  if (!s) return '';
  const m = s.match(/^(?:标包|包)?(\d+)$/);
  return m ? `包${parseInt(m[1], 10)}` : s;
}

/**
 * Build multi-level supplier contact resolver supporting ECP's "分标包号" column
 * (e.g. "零星服务-铜陵公司:包4,包5,包6") as well as ETP exports without package_no.
 */
export function buildSupplierContactResolver(suppliers: Array<Record<string, unknown>>) {
  const byLotPkg = new Map<string, Record<string, unknown>>();
  const byPkgOnly = new Map<string, Record<string, unknown>>();
  const byNameOnly = new Map<string, Record<string, unknown>>();

  for (const s of suppliers) {
    const sk = cleanSupplierKey(String(s.supplier_name || ''));
    if (!sk) continue;
    if (!byNameOnly.has(sk)) byNameOnly.set(sk, s);

    const rawPkg = String(s.package_no || '').trim();
    if (!rawPkg) continue;

    const segments = rawPkg.split(/[;；\r\n]+/).map((seg) => seg.trim()).filter(Boolean);
    for (const seg of segments) {
      const colonIdx = seg.search(/[:：]/);
      if (colonIdx >= 0) {
        const lotPart = cleanSupplierKey(seg.slice(0, colonIdx));
        const pkgsPart = seg.slice(colonIdx + 1);
        const tokens = pkgsPart.split(/[,，、\s]+/).map(normalizePkgToken).filter(Boolean);
        for (const tok of tokens) {
          if (lotPart) {
            const k1 = `${sk}::${lotPart}::${tok}`;
            if (!byLotPkg.has(k1)) byLotPkg.set(k1, s);
          }
          const k2 = `${sk}::${tok}`;
          if (!byPkgOnly.has(k2)) byPkgOnly.set(k2, s);
        }
      } else {
        const tokens = seg.split(/[,，、\s]+/).map(normalizePkgToken).filter(Boolean);
        for (const tok of tokens) {
          const k2 = `${sk}::${tok}`;
          if (!byPkgOnly.has(k2)) byPkgOnly.set(k2, s);
        }
      }
    }
  }

  return (supplierName: string, lotName: string, packageNo: string): Record<string, unknown> | undefined => {
    const sk = cleanSupplierKey(supplierName);
    if (!sk) return undefined;
    const normLot = cleanSupplierKey(lotName);
    const normPkg = normalizePkgToken(packageNo);
    if (normLot && normPkg) {
      const exact = byLotPkg.get(`${sk}::${normLot}::${normPkg}`);
      if (exact) return exact;
    }
    if (normPkg) {
      const byPkg = byPkgOnly.get(`${sk}::${normPkg}`);
      if (byPkg) return byPkg;
    }
    return byNameOnly.get(sk);
  };
}

const REJECTION_TAG_ORDER = ['商务否决', '技术否决', '资格否决', '形式否决', '报价否决'];

function extractRejectionEntry(rj: Record<string, unknown>): { tag: string; detail: string } {
  const reviewType = String(rj.review_type || '').trim();
  const reason = String(rj.reason_category || '').trim();
  const tpl = String(rj.rejection_tpl || '').trim();
  const detail = reason || tpl;
  const combined = `${reviewType} ${reason}`;
  let tag = '未进入详评';
  if (/商务/.test(combined)) tag = '商务否决';
  else if (/技术/.test(combined)) tag = '技术否决';
  else if (/资格/.test(combined)) tag = '资格否决';
  else if (/形式/.test(combined)) tag = '形式否决';
  else if (/报价|价格|限价/.test(combined)) tag = '报价否决';
  else if (reviewType) tag = reviewType.endsWith('否决') ? reviewType : `${reviewType.replace(/类$/, '')}否决`;
  return { tag, detail: detail === tag ? '' : detail };
}

function joinRejectionEntries(grouped: Map<string, Set<string>> | undefined): string {
  if (!grouped || grouped.size === 0) return '';
  return Array.from(grouped.entries())
    .sort(([a], [b]) => {
      const ia = REJECTION_TAG_ORDER.indexOf(a);
      const ib = REJECTION_TAG_ORDER.indexOf(b);
      return (ia === -1 ? 99 : ia) - (ib === -1 ? 99 : ib);
    })
    .map(([tag, details]) => {
      const list = Array.from(details).filter(Boolean);
      return list.length > 0 ? `${tag}：${list.join('、')}` : tag;
    })
    .join('；');
}

function buildAwardContext(batchId: string, overrideOpts?: { category?: string; discountRate?: number; ownerUnit?: string }) {
  const batch = db.prepare('SELECT * FROM batches WHERE id = ?').get(batchId) as Record<string, unknown> | undefined;
  if (!batch) return null;

  const clients = db.prepare('SELECT * FROM client_agreements ORDER BY sort_order ASC, client_name ASC').all() as unknown as ClientAgreementRow[];
  const selectedClient = matchDefaultClient(clients, batch);

  const category: ServiceFeeCategory = normalizeCategory(
    overrideOpts?.category || String(batch.award_category || '服务')
  );
  const discountRate =
    overrideOpts?.discountRate !== undefined && Number.isFinite(overrideOpts.discountRate) && overrideOpts.discountRate > 0
      ? overrideOpts.discountRate
      : batch.award_discount_rate !== null && batch.award_discount_rate !== undefined && Number(batch.award_discount_rate) > 0
        ? Number(batch.award_discount_rate)
        : selectedClient
          ? Number(selectedClient.discount_rate)
          : 1;
  const ownerUnit =
    overrideOpts?.ownerUnit ||
    (selectedClient ? selectedClient.client_name.replace(/（[^）]*）$/, '') : String(batch.project_unit || ''));

  const projects = db.prepare('SELECT * FROM batch_projects WHERE batch_id = ? ORDER BY seq ASC').all(batchId) as Array<Record<string, unknown>>;
  const suppliers = db.prepare('SELECT * FROM batch_suppliers WHERE batch_id = ?').all(batchId) as Array<Record<string, unknown>>;

  // 对历史已上传过 ERP 采购计划（含主业+产业多份）、但尚未入库 material_code 的批次自动回填一次
  if (projects.length > 0 && projects.some((p) => !String(p.material_code || '').trim())) {
    const erpFiles = db
      .prepare("SELECT stored_path FROM uploaded_files WHERE batch_id = ? AND module = 'bidding' AND slot_key = 'erp_plan' ORDER BY created_at ASC")
      .all(batchId) as Array<{ stored_path?: string }>;
    if (erpFiles.length > 0) {
      const byReq = new Map<string, { code: string; desc: string }>();
      const byProjCode = new Map<string, { code: string; desc: string }>();
      const byName = new Map<string, { code: string; desc: string }>();
      for (const erpFile of erpFiles) {
        if (!erpFile?.stored_path) continue;
        const absErpPath = path.join(UPLOADS_DIR, erpFile.stored_path);
        if (!fs.existsSync(absErpPath)) continue;
        try {
          const erpParsed = parseErpPlan(absErpPath);
          for (const er of erpParsed.rows) {
            const mc = String(er.material_code || '').trim();
            const md = String(er.material_desc || '').trim();
            if (!mc && !md) continue;
            const rNo = stripAllSpaces(er.req_app_no);
            const pCode = stripAllSpaces(er.project_code);
            const pName = stripAllSpaces(er.name);
            if (rNo) {
              if (!byReq.has(rNo)) byReq.set(rNo, { code: mc, desc: md });
              if (rNo.length === 15 && !byReq.has(rNo.slice(0, 10))) {
                byReq.set(rNo.slice(0, 10), { code: mc, desc: md });
              }
            }
            if (pCode && !byProjCode.has(pCode)) byProjCode.set(pCode, { code: mc, desc: md });
            if (pName && !byName.has(pName)) byName.set(pName, { code: mc, desc: md });
          }
        } catch {
          // ignore non-fatal backfill errors
        }
      }
      if (byReq.size > 0 || byProjCode.size > 0 || byName.size > 0) {
        const updMat = db.prepare('UPDATE batch_projects SET material_code = ?, material_desc = ? WHERE id = ?');
        for (const p of projects) {
          if (String(p.material_code || '').trim()) continue;
          const pReq = stripAllSpaces(p.req_app_no);
          const hit =
            byReq.get(pReq) ||
            (pReq.length === 15 ? byReq.get(pReq.slice(0, 10)) : undefined) ||
            byProjCode.get(stripAllSpaces(p.project_code)) ||
            byName.get(stripAllSpaces(p.name));
          if (hit) {
            p.material_code = hit.code;
            p.material_desc = hit.desc;
            updMat.run(hit.code, hit.desc, String(p.id));
          }
        }
      }
    }
  }

  const materialDescMap = lookupMaterialDescMap(projects.map((p) => String(p.material_code || '')));
  const rankings = db.prepare('SELECT * FROM eval_rankings WHERE batch_id = ? ORDER BY rank_no ASC').all(batchId) as Array<Record<string, unknown>>;
  const rejections = db.prepare('SELECT * FROM eval_rejections WHERE batch_id = ?').all(batchId) as Array<Record<string, unknown>>;
  const bids = db.prepare('SELECT * FROM eval_bids WHERE batch_id = ?').all(batchId) as Array<Record<string, unknown>>;

  // Supplier contact lookup: 供应商 + 分标名称 + 包号 精确匹配 -> 供应商 + 包号 -> 供应商名称兜底（兼容 ETP）
  const resolveSupplierContact = buildSupplierContactResolver(suppliers);

  // Merge packages keyed by lot_code::package_no
  const pkgKey = (lotCode: string, pkgNo: string) => `${lotCode || ''}::${pkgNo || ''}`;
  const pkgMap = new Map<string, AwardPackageRow>();

  const ensurePkg = (lotCode: string, lotName: string, pkgNo: string) => {
    const k = pkgKey(lotCode, pkgNo);
    let entry = pkgMap.get(k);
    if (!entry) {
      entry = {
        seq: pkgMap.size + 1,
        lot_code: lotCode,
        lot_name: lotName,
        package_no: pkgNo,
        project_name: '',
        project_overview: '',
        req_app_no: '',
        budget_wan: null,
        max_price_wan: null,
        max_price_raw: '',
        declarant_info: '',
        material_desc: '',
        is_failed_lot: false,
        winner_name: '',
        winner_price_wan: null,
        winner_price_raw: '',
        winner_score: null,
        winner_contact: '',
        winner_mobile: '',
        winner_email: '',
        winner_address: '',
        service_fee_pre_yuan: null,
        service_fee_yuan: null,
        platform_fee_yuan: null,
        category,
        discount_rate: discountRate,
        remark: '',
      };
      pkgMap.set(k, entry);
    } else if (!entry.lot_name && lotName) {
      entry.lot_name = lotName;
    }
    return entry;
  };

  for (const p of projects) {
    const lotCode = String(p.lot_code || p.code || '');
    const lotName = String(p.lot_name || '');
    const pkgNo = String(p.package_no || '');
    if (!lotCode && !pkgNo && !p.name) continue;
    const existed = pkgMap.has(pkgKey(lotCode, pkgNo));
    const entry = ensurePkg(lotCode, lotName, pkgNo);
    const pName = String(p.name || '');
    const pOverview = String(p.project_overview || '');
    const reqNo = String(p.req_app_no || '');
    const budgetYuan = p.budget_yuan !== null && p.budget_yuan !== undefined ? Number(p.budget_yuan) : null;
    const maxYuan = p.max_price_yuan !== null && p.max_price_yuan !== undefined ? Number(p.max_price_yuan) : null;
    const addBudgetWan = budgetYuan !== null ? Number((budgetYuan / 10000).toFixed(6)) : null;
    const addMaxWan = maxYuan !== null ? Number((maxYuan / 10000).toFixed(6)) : addBudgetWan;

    const matCode = stripAllSpaces(p.material_code);
    const resolvedMatDesc = (matCode ? materialDescMap.get(matCode) : '') || String(p.material_desc || '').trim();

    if (!existed) {
      entry.project_name = pName || entry.project_name;
      entry.project_overview = pOverview;
      entry.req_app_no = reqNo;
      entry.budget_wan = addBudgetWan;
      entry.max_price_wan = addMaxWan;
      entry.max_price_raw = String(p.max_price_raw || '');
      const dName = String(p.declarant_name || '').trim();
      const dPhone = String(p.declarant_phone || '').trim();
      entry.declarant_info = dName && dPhone ? `${dName} ${dPhone}` : dName || dPhone;
      entry.material_desc = resolvedMatDesc;
    } else {
      // 合并成包（多个子包共用同一分标编号+包号）：累加预算并拼接子包项目名、申请号与物料描述
      if (pName && !entry.project_name.includes(pName)) {
        entry.project_name = entry.project_name ? `${entry.project_name}；${pName}` : pName;
      }
      if (resolvedMatDesc && !entry.material_desc.includes(resolvedMatDesc)) {
        entry.material_desc = entry.material_desc ? `${entry.material_desc}；${resolvedMatDesc}` : resolvedMatDesc;
      }
      if (pOverview && !entry.project_overview.includes(pOverview)) {
        entry.project_overview = entry.project_overview ? `${entry.project_overview}；${pOverview}` : pOverview;
      }
      if (reqNo && !entry.req_app_no.includes(reqNo)) {
        entry.req_app_no = entry.req_app_no ? `${entry.req_app_no} / ${reqNo}` : reqNo;
      }
      if (addBudgetWan !== null) {
        entry.budget_wan = Number(((entry.budget_wan || 0) + addBudgetWan).toFixed(6));
      }
      if (addMaxWan !== null) {
        entry.max_price_wan = Number(((entry.max_price_wan || 0) + addMaxWan).toFixed(6));
      }
    }
  }

  // Track bidders & rejections per package for early failed-lot detection
  const bidderSets = new Map<string, Set<string>>();
  for (const b of bids) {
    const entry = ensurePkg(String(b.lot_code || ''), String(b.lot_name || ''), String(b.package_no || ''));
    const k = pkgKey(entry.lot_code, entry.package_no);
    if (!bidderSets.has(k)) bidderSets.set(k, new Set());
    bidderSets.get(k)!.add(cleanSupplierKey(String(b.supplier_name || '')));
  }

  const rejSets = new Map<string, Set<string>>();
  for (const rj of rejections) {
    const lotCode = String(rj.lot_code || '');
    const lotName = String(rj.lot_name || '');
    const pkgNo = String(rj.package_no || '');
    let target = lotCode ? pkgMap.get(pkgKey(lotCode, pkgNo)) : undefined;
    if (!target && !lotCode && lotName && pkgNo) {
      for (const c of pkgMap.values()) {
        if (c.lot_name === lotName && c.package_no === pkgNo) {
          target = c;
          break;
        }
      }
    }
    if (!target) target = ensurePkg(lotCode, lotName, pkgNo);
    const k = pkgKey(target.lot_code, target.package_no);
    if (!rejSets.has(k)) rejSets.set(k, new Set());
    rejSets.get(k)!.add(cleanSupplierKey(String(rj.supplier_name || '')));
  }

  // Attach rankings (rank 1 = winner), handling 折扣率报价 (e.g. 88 -> 88%, monetary equivalent = budget_wan * 88%)
  for (const rk of rankings) {
    const entry = ensurePkg(String(rk.lot_code || ''), String(rk.lot_name || ''), String(rk.package_no || ''));
    if (!entry.project_name && rk.project_name) entry.project_name = String(rk.project_name);
    if (Number(rk.is_failed_lot) === 1) {
      entry.is_failed_lot = true;
    } else if (Number(rk.rank_no) === 1) {
      entry.winner_name = String(rk.supplier_name || '');
      const rawPrice = String(rk.eval_price || rk.bid_price || '').trim();
      const isDiscountMode = /%$/.test(entry.max_price_raw.trim());
      if (isDiscountMode) {
        const cleanedNum = Number(rawPrice.replace(/%$/, '').trim());
        const pctPoints = Number.isFinite(cleanedNum) && cleanedNum > 0
          ? (cleanedNum <= 1.5 && !rawPrice.endsWith('%') ? Number((cleanedNum * 100).toFixed(4)) : cleanedNum)
          : null;
        entry.winner_price_raw = pctPoints !== null ? `${pctPoints}%` : rawPrice;
        // For service fee calculation on discount-rate packages: budget_wan * (pct / 100)
        entry.winner_price_wan =
          pctPoints !== null && entry.budget_wan !== null && entry.budget_wan > 0
            ? Number(((entry.budget_wan * pctPoints) / 100).toFixed(6))
            : null;
        if (!entry.remark) entry.remark = '折扣率';
      } else {
        entry.winner_price_raw = rawPrice;
        const numPrice = Number(rawPrice);
        entry.winner_price_wan = Number.isFinite(numPrice) && numPrice > 0 ? numPrice : null;
      }
      entry.winner_score = rk.total_score !== null ? Number(rk.total_score) : null;
    }
  }

  // Sort packages by 分标名称 Pinyin + 包号 integer
  const packages = Array.from(pkgMap.values()).sort((a, b) => {
    const lotCmp = pinyinCollator.compare(a.lot_name || a.lot_code, b.lot_name || b.lot_code);
    if (lotCmp !== 0) return lotCmp;
    return extractPkgNum(a.package_no) - extractPkgNum(b.package_no);
  });

  const minDetailedEval = String(batch.procurement_type || '') === '公开招标' ? 3 : 2;
  let totalAwardWan = 0;
  let totalServiceFeePreYuan = 0;
  let totalServiceFeeYuan = 0;
  let totalPlatformFeeYuan = 0;

  packages.forEach((p, idx) => {
    p.seq = idx + 1;
    const k = pkgKey(p.lot_code, p.package_no);
    const bCount = bidderSets.get(k)?.size || 0;
    const rCount = rejSets.get(k)?.size || 0;
    if (!p.winner_name && bCount > 0 && Math.max(0, bCount - rCount) < minDetailedEval) {
      p.is_failed_lot = true;
    }

    if (!p.is_failed_lot && p.winner_name) {
      const sup = resolveSupplierContact(p.winner_name, p.lot_name, p.package_no);
      if (sup) {
        p.winner_contact = String(sup.contact_name || '');
        p.winner_mobile = String(sup.mobile || sup.phone || '');
        p.winner_email = String(sup.email || '');
        p.winner_address = String(sup.address || '');
      }
      if (p.winner_price_wan !== null && p.winner_price_wan > 0) {
        totalAwardWan += p.winner_price_wan;
        const fee = calculateServiceFee(p.winner_price_wan, p.category, p.discount_rate);
        if (fee) {
          p.service_fee_pre_yuan = fee.preDiscountYuan;
          p.service_fee_yuan = fee.postDiscountYuan;
          p.platform_fee_yuan = fee.platformFeeYuan;
          totalServiceFeePreYuan += fee.preDiscountYuan;
          totalServiceFeeYuan += fee.postDiscountYuan;
          totalPlatformFeeYuan += fee.platformFeeYuan;
        }
      }
    }
  });

  // Build all-bidder detail rows for 成交结果表 (22 cols)
  // Anchored on ALL bidders from 开标一览表 (eval_bids) merged with 总分排序表 (eval_rankings) & 未进入详评表 (eval_rejections)
  const bidderDetails: AwardBidderDetailRow[] = [];
  for (const p of packages) {
    const pkgRankings = rankings.filter(
      (r) => String(r.lot_code || '') === p.lot_code && String(r.package_no || '') === p.package_no && Number(r.is_failed_lot) === 0
    );
    const pkgRejs = rejections.filter((rj) => {
      const lc = String(rj.lot_code || '');
      const ln = String(rj.lot_name || '');
      const pn = String(rj.package_no || '');
      return pn === p.package_no && (lc ? lc === p.lot_code : ln === p.lot_name);
    });
    const pkgBids = bids.filter((b) => {
      const lc = String(b.lot_code || '');
      const ln = String(b.lot_name || '');
      const pn = String(b.package_no || '');
      return pn === p.package_no && (lc && p.lot_code ? lc === p.lot_code : ln === p.lot_name);
    });

    const rejReasonMap = new Map<string, Map<string, Set<string>>>();
    for (const rj of pkgRejs) {
      const sk = cleanSupplierKey(String(rj.supplier_name || ''));
      if (!sk) continue;
      if (!rejReasonMap.has(sk)) rejReasonMap.set(sk, new Map());
      const byTag = rejReasonMap.get(sk)!;
      const { tag, detail } = extractRejectionEntry(rj);
      if (!byTag.has(tag)) byTag.set(tag, new Set());
      if (detail) byTag.get(tag)!.add(detail);
    }

    const addedSuppliers = new Set<string>();
    const maxDisp = p.max_price_raw || (p.max_price_wan !== null ? String(p.max_price_wan) : '');

    // 1. Ranked bidders first (ordered by rank_no 1, 2, 3...)
    for (const rk of pkgRankings) {
      const sName = String(rk.supplier_name || '');
      const sk = cleanSupplierKey(sName);
      if (!sk || addedSuppliers.has(sk)) continue;
      addedSuppliers.add(sk);
      const sup = resolveSupplierContact(sName, p.lot_name, p.package_no);
      bidderDetails.push({
        seq: bidderDetails.length + 1,
        batch_code: String(batch.code || ''),
        lot_name: p.lot_name,
        package_no: p.package_no,
        max_price_display: maxDisp,
        department: ownerUnit,
        declarant_info: p.declarant_info,
        project_name: p.project_name,
        material_desc: p.material_desc,
        supplier_name: sName,
        contact_name: String(sup?.contact_name || ''),
        mobile: String(sup?.mobile || sup?.phone || ''),
        email: String(sup?.email || ''),
        is_compliant: '是',
        tech_score: rk.tech_score !== null && rk.tech_score !== undefined ? Number(rk.tech_score).toFixed(2) : '/',
        biz_score: rk.biz_score !== null && rk.biz_score !== undefined ? Number(rk.biz_score).toFixed(2) : '/',
        price_score: rk.price_score !== null && rk.price_score !== undefined ? Number(rk.price_score).toFixed(2) : '/',
        total_score: rk.total_score !== null && rk.total_score !== undefined ? Number(rk.total_score).toFixed(2) : '/',
        rank_no: rk.rank_no !== null && rk.rank_no !== undefined ? String(rk.rank_no) : '/',
        base_price: String(rk.base_price || '/'),
        eval_price: String(rk.eval_price || rk.bid_price || '/'),
        remark: '',
      });
    }

    // 2. All remaining bidders from 开标一览表 (eval_bids) — covers rejected bidders & failed lots
    for (const b of pkgBids) {
      const sName = String(b.supplier_name || '');
      const sk = cleanSupplierKey(sName);
      if (!sk || addedSuppliers.has(sk)) continue;
      addedSuppliers.add(sk);
      const sup = resolveSupplierContact(sName, p.lot_name, p.package_no);
      const rejReason = joinRejectionEntries(rejReasonMap.get(sk));
      const hasRankingsOrRej = pkgRankings.length > 0 || pkgRejs.length > 0 || p.is_failed_lot;
      bidderDetails.push({
        seq: bidderDetails.length + 1,
        batch_code: String(batch.code || ''),
        lot_name: p.lot_name,
        package_no: p.package_no,
        max_price_display: maxDisp,
        department: ownerUnit,
        declarant_info: p.declarant_info,
        project_name: p.project_name,
        material_desc: p.material_desc,
        supplier_name: sName,
        contact_name: String(sup?.contact_name || ''),
        mobile: String(sup?.mobile || sup?.phone || ''),
        email: String(sup?.email || ''),
        is_compliant: hasRankingsOrRej ? '否' : '是',
        tech_score: '/',
        biz_score: '/',
        price_score: '/',
        total_score: '/',
        rank_no: '/',
        base_price: '/',
        eval_price: hasRankingsOrRej ? '/' : String(b.bid_price_raw || (b.bid_price_wan ?? '/')),
        remark: p.is_failed_lot ? (rejReason ? `流标（${rejReason}）` : '流标') : rejReason,
      });
    }

    // 3. Any extra rejected bidders not present in eval_bids
    for (const rj of pkgRejs) {
      const sName = String(rj.supplier_name || '');
      const sk = cleanSupplierKey(sName);
      if (!sk || addedSuppliers.has(sk)) continue;
      addedSuppliers.add(sk);
      const sup = resolveSupplierContact(sName, p.lot_name, p.package_no);
      const rejReason = joinRejectionEntries(rejReasonMap.get(sk));
      bidderDetails.push({
        seq: bidderDetails.length + 1,
        batch_code: String(batch.code || ''),
        lot_name: p.lot_name,
        package_no: p.package_no,
        max_price_display: maxDisp,
        department: ownerUnit,
        declarant_info: p.declarant_info,
        project_name: p.project_name,
        material_desc: p.material_desc,
        supplier_name: sName,
        contact_name: String(sup?.contact_name || ''),
        mobile: String(sup?.mobile || sup?.phone || ''),
        email: String(sup?.email || ''),
        is_compliant: '否',
        tech_score: '/',
        biz_score: '/',
        price_score: '/',
        total_score: '/',
        rank_no: '/',
        base_price: '/',
        eval_price: '/',
        remark: p.is_failed_lot ? (rejReason ? `流标（${rejReason}）` : '流标') : rejReason,
      });
    }
  }

  return {
    batch,
    clients,
    selectedClient,
    awardConfig: {
      client_id: selectedClient?.id || '',
      owner_unit: ownerUnit,
      category,
      discount_rate: discountRate,
      discount_label: formatDiscountLabel(discountRate),
      batch_manager: String(batch.award_batch_manager || ''),
      bid_location: String(batch.award_bid_location || '华亿'),
      publicity_date: String(batch.award_publicity_date || ''),
      announcement_date: String(batch.award_announcement_date || ''),
    },
    packages,
    bidderDetails,
    totals: {
      winningCount: packages.filter((p) => !p.is_failed_lot && p.winner_name).length,
      failedCount: packages.filter((p) => p.is_failed_lot).length,
      totalAwardWan: Number(totalAwardWan.toFixed(6)),
      totalServiceFeePreYuan: Number(totalServiceFeePreYuan.toFixed(2)),
      totalServiceFeeYuan: Number(totalServiceFeeYuan.toFixed(2)),
      totalPlatformFeeYuan: Number(totalPlatformFeeYuan.toFixed(2)),
    },
  };
}

// GET /api/award/clients
awardRouter.get('/clients', (_req, res) => {
  const clients = db.prepare('SELECT * FROM client_agreements ORDER BY sort_order ASC, client_name ASC').all();
  res.json(clients);
});

// POST /api/award/clients/import — upload updated 客户协议清单 (.xls / .xlsx)
awardRouter.post('/clients/import', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: '请选择客户协议清单 Excel 文件' });
  try {
    const parsed = parseClientAgreementsExcel(req.file.path);
    if (parsed.length === 0) {
      return res.status(400).json({ error: '未在表格中解析到有效客户协议记录' });
    }
    db.prepare('DELETE FROM client_agreements').run();
    const ins = db.prepare(`
      INSERT INTO client_agreements (id, group_name, client_name, contact_name, contact_phone, contract_expiry, discount_rate, discount_label, agreement_status, notes, sort_order)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);
    parsed.forEach((c, idx) => {
      ins.run(uuid(), c.group_name, c.client_name, c.contact_name, c.contact_phone, c.contract_expiry, c.discount_rate, c.discount_label, c.agreement_status, c.notes, idx + 1);
    });
    res.json({ ok: true, count: parsed.length, clients: db.prepare('SELECT * FROM client_agreements ORDER BY sort_order ASC').all() });
  } catch (err) {
    res.status(400).json({ error: `解析客户协议清单失败：${(err as Error).message}` });
  } finally {
    if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
  }
});

// POST /api/award/materials/import — upload updated 服务物料编码--分标名称（授权采购）.xlsx
awardRouter.post('/materials/import', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: '请选择服务物料编码对照表 Excel 文件' });
  try {
    const count = importMaterialCatalogToDb(req.file.path);
    res.json({ ok: true, count });
  } catch (err) {
    res.status(400).json({ error: `解析物料编码对照表失败：${(err as Error).message}` });
  } finally {
    if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
  }
});

// PUT /api/award/batch/:batchId/config — persist selected owner unit, category, discount, manager, location, publicity/announcement dates
awardRouter.put('/batch/:batchId/config', (req, res) => {
  const batchId = req.params.batchId;
  const { client_id, category, discount_rate, batch_manager, bid_location, publicity_date, announcement_date } = req.body as {
    client_id?: string;
    category?: string;
    discount_rate?: number | null;
    batch_manager?: string;
    bid_location?: string;
    publicity_date?: string;
    announcement_date?: string;
  };

  const existing = db.prepare('SELECT * FROM batches WHERE id = ?').get(batchId) as Record<string, unknown> | undefined;
  if (!existing) return res.status(404).json({ error: '未找到该批次' });

  const normCat = category !== undefined ? normalizeCategory(category) : normalizeCategory(String(existing.award_category || '服务'));
  const cleanRate =
    discount_rate !== undefined
      ? discount_rate !== null && Number.isFinite(Number(discount_rate)) && Number(discount_rate) > 0
        ? Number(discount_rate)
        : null
      : (existing.award_discount_rate as number | null);

  db.prepare(`
    UPDATE batches
    SET award_client_id = ?, award_category = ?, award_discount_rate = ?, award_batch_manager = ?, award_bid_location = ?, award_publicity_date = ?, award_announcement_date = ?, updated_at = ?
    WHERE id = ?
  `).run(
    client_id !== undefined ? String(client_id) : String(existing.award_client_id || ''),
    normCat,
    cleanRate,
    batch_manager !== undefined ? String(batch_manager) : String(existing.award_batch_manager || ''),
    bid_location !== undefined ? String(bid_location) : String(existing.award_bid_location || '华亿'),
    publicity_date !== undefined ? String(publicity_date) : String(existing.award_publicity_date || ''),
    announcement_date !== undefined ? String(announcement_date) : String(existing.award_announcement_date || ''),
    new Date().toISOString(),
    batchId
  );

  res.json({ ok: true });
});

// GET /api/award/batch/:batchId
awardRouter.get('/batch/:batchId', (req, res) => {
  const batchId = req.params.batchId;
  const ctx = buildAwardContext(batchId);
  if (!ctx) return res.status(404).json({ error: '未找到该批次' });

  const biddingFiles = db.prepare("SELECT id, category, original_name, parsed_json, created_at FROM uploaded_files WHERE batch_id = ? AND module = 'bidding' ORDER BY created_at DESC").all(batchId) as Record<string, unknown>[];
  const evalFiles = db.prepare("SELECT id, category, original_name, parsed_json, created_at FROM uploaded_files WHERE batch_id = ? AND module = 'evaluation' ORDER BY created_at DESC").all(batchId) as Record<string, unknown>[];
  const outputs = db.prepare('SELECT * FROM award_outputs WHERE batch_id = ? ORDER BY created_at DESC').all(batchId);

  res.json({
    readiness: {
      hasBatchInfo: Boolean(ctx.batch.name && ctx.batch.project_unit),
      projectCount: ctx.packages.length,
      biddingFileCount: biddingFiles.length,
      evalFileCount: evalFiles.length,
    },
    batch: ctx.batch,
    clients: ctx.clients,
    awardConfig: ctx.awardConfig,
    packages: ctx.packages,
    totals: ctx.totals,
    templates: AWARD_TEMPLATES,
    outputs,
  });
});

// POST /api/award/batch/:batchId/generate
awardRouter.post('/batch/:batchId/generate', async (req, res) => {
  try {
    const batchId = req.params.batchId;
    const { template_key, overrides = {}, sync_to_files = true } = req.body as {
      template_key?: string;
      overrides?: Record<string, string>;
      sync_to_files?: boolean;
    };

    const template = AWARD_TEMPLATES.find((t) => t.key === template_key);
    if (!template) return res.status(400).json({ error: '无效的交付物模板' });

    const ctx = buildAwardContext(batchId, {
      category: overrides.category,
      discountRate: overrides.discountRate ? Number(overrides.discountRate) : undefined,
      ownerUnit: overrides.ownerUnit,
    });
    if (!ctx) return res.status(404).json({ error: '批次不存在' });

    const verRow = db.prepare('SELECT COUNT(*) as c FROM award_outputs WHERE batch_id = ? AND template_key = ?').get(batchId, template.key) as { c: number };
    const version = (verRow.c || 0) + 1;

    const isPublicBatch = String(ctx.batch.procurement_type || '') === '公开招标';
    const batchCode = String(overrides.batchCode || ctx.batch.code || 'BATCH');
    const safePrefix = batchCode.replace(/[^\w\u4e00-\u9fa5-]/g, '_');
    const shortTitle =
      template.key === 'pre_award_publicity'
        ? isPublicBatch
          ? '推荐的中标候选人公示'
          : '预成交公示'
        : template.key === 'award_result_notice'
          ? isPublicBatch
            ? '中标公告'
            : '成交结果公告'
          : template.title.replace(/（.*）/, '');
    const filename = `${safePrefix}_${shortTitle}_v${version}.${template.file_type}`;
    const relPath = path.join(batchId, 'award', `${Date.now()}-${uuid().slice(0, 6)}.${template.file_type}`);
    const absPath = path.join(UPLOADS_DIR, relPath);

    const payload: AwardDeliverablePayload = {
      batch: {
        id: String(ctx.batch.id),
        name: String(ctx.batch.name || ''),
        code: String(ctx.batch.code || ''),
        project_unit: String(ctx.batch.project_unit || ''),
        procurement_type: String(ctx.batch.procurement_type || '非公开招标'),
        unit_nature: String(ctx.batch.unit_nature || '产业'),
        announcement_time: String(ctx.batch.announcement_time || ''),
        bid_opening_time: String(ctx.batch.bid_opening_time || ''),
        contact_phone: String(ctx.batch.contact_phone || ''),
        contact_email: String(ctx.batch.contact_email || ''),
      },
      ownerUnit: ctx.awardConfig.owner_unit,
      category: ctx.awardConfig.category,
      discountRate: ctx.awardConfig.discount_rate,
      discountLabel: ctx.awardConfig.discount_label,
      packages: ctx.packages,
      bidderDetails: ctx.bidderDetails,
      overrides: {
        batchManager: ctx.awardConfig.batch_manager,
        bidLocation: ctx.awardConfig.bid_location,
        publicityDate: ctx.awardConfig.publicity_date,
        announcementDate: ctx.awardConfig.announcement_date,
        ...overrides,
      },
    };

    if (template.key === 'pre_award_publicity' && overrides.publicityDate) {
      db.prepare('UPDATE batches SET award_publicity_date = ?, updated_at = ? WHERE id = ?').run(
        String(overrides.publicityDate),
        new Date().toISOString(),
        batchId
      );
    } else if (template.key === 'award_result_notice' && overrides.announcementDate) {
      db.prepare('UPDATE batches SET award_announcement_date = ?, updated_at = ? WHERE id = ?').run(
        String(overrides.announcementDate),
        new Date().toISOString(),
        batchId
      );
    }

    if (template.key === 'pre_award_publicity') {
      generatePreAwardPublicityDocx(payload, absPath);
    } else if (template.key === 'award_result_notice') {
      generateAwardResultNoticeDocx(payload, absPath);
    } else if (template.key === 'notice_handover') {
      await generateNoticeHandoverExcel(payload, absPath);
    } else if (template.key === 'project_handover') {
      await generateProjectHandoverExcel(payload, absPath);
    } else if (template.key === 'award_result_table') {
      await generateAwardResultTableExcel(payload, absPath);
    } else if (template.key === 'service_fee_sheet') {
      await generateServiceFeeExcel(payload, absPath);
    }

    const id = uuid();
    const now = new Date().toISOString();
    db.prepare(`
      INSERT INTO award_outputs (id, batch_id, template_key, title, file_type, filename, stored_path, snapshot_json, version, created_at)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).run(id, batchId, template.key, template.title, template.file_type, filename, relPath, JSON.stringify(overrides), version, now);

    if (sync_to_files) {
      const stat = fs.statSync(absPath);
      db.prepare(`
        INSERT INTO uploaded_files (id, batch_id, module, category, original_name, stored_path, mime_type, file_size, parsed_json, notes, created_at)
        VALUES (?, ?, 'award', ?, ?, ?, ?, ?, NULL, ?, ?)
      `).run(
        uuid(),
        batchId,
        shortTitle,
        filename,
        relPath,
        template.file_type === 'xlsx'
          ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
          : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        stat.size,
        `定标模块自动生成 v${version}`,
        now
      );
    }

    res.status(201).json(db.prepare('SELECT * FROM award_outputs WHERE id = ?').get(id));
  } catch (err) {
    res.status(500).json({ error: `生成交付物失败: ${(err as Error).message}` });
  }
});

// GET /api/award/outputs/:id/download
awardRouter.get('/outputs/:id/download', (req, res) => {
  const row = db.prepare('SELECT * FROM award_outputs WHERE id = ?').get(req.params.id) as Record<string, string> | undefined;
  if (!row) return res.status(404).json({ error: '交付物不存在' });
  const absPath = path.join(UPLOADS_DIR, row.stored_path);
  if (!fs.existsSync(absPath)) return res.status(404).json({ error: '文件不存在' });
  res.download(absPath, row.filename);
});

// DELETE /api/award/outputs/:id
awardRouter.delete('/outputs/:id', (req, res) => {
  const row = db.prepare('SELECT * FROM award_outputs WHERE id = ?').get(req.params.id) as Record<string, string> | undefined;
  if (row?.stored_path) {
    const absPath = path.join(UPLOADS_DIR, row.stored_path);
    if (fs.existsSync(absPath)) fs.unlinkSync(absPath);
    db.prepare("DELETE FROM uploaded_files WHERE module = 'award' AND stored_path = ?").run(row.stored_path);
  }
  db.prepare('DELETE FROM award_outputs WHERE id = ?').run(req.params.id);
  res.json({ ok: true });
});

// POST /api/award/outputs/batch-delete
awardRouter.post('/outputs/batch-delete', (req, res) => {
  const { ids } = req.body as { ids?: string[] };
  if (!Array.isArray(ids) || ids.length === 0) {
    return res.status(400).json({ error: '请选择至少一条要删除的交付文件记录' });
  }
  const cleanIds = ids.map((x) => String(x || '').trim()).filter(Boolean);
  let deleted = 0;
  const sel = db.prepare('SELECT id, stored_path FROM award_outputs WHERE id = ?');
  const delOut = db.prepare('DELETE FROM award_outputs WHERE id = ?');
  const delFile = db.prepare("DELETE FROM uploaded_files WHERE module = 'award' AND stored_path = ?");
  for (const id of cleanIds) {
    const row = sel.get(id) as { id: string; stored_path?: string } | undefined;
    if (!row) continue;
    if (row.stored_path) {
      const absPath = path.join(UPLOADS_DIR, row.stored_path);
      if (fs.existsSync(absPath)) fs.unlinkSync(absPath);
      delFile.run(row.stored_path);
    }
    delOut.run(id);
    deleted++;
  }
  res.json({ ok: true, deleted });
});
