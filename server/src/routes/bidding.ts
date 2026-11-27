import fs from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';
import { Router } from 'express';
import multer from 'multer';
import { v4 as uuid } from 'uuid';
import {
  detectCollusionRisks,
  formatYuanAsWan,
  isSameReqAppNo,
  parseAnyProjectSpreadsheet,
  parseDeclarationPlan,
  parseErpPlan,
  parseRequirementList,
  parseSupplierContacts,
  parseVariableTemplate,
  stripAllSpaces,
} from '../biddingParser.js';
import { db, UPLOADS_DIR } from '../db.js';
import { parseSpreadsheetTables } from '../excelParser.js';

export const biddingRouter = Router();

export const BIDDING_SLOTS = [
  {
    key: 'declaration_plan',
    order: 1,
    label: '批次申报计划表',
    stage: '文件审查会',
    description: '提取项目管理部门、项目编码、资金来源及申报人姓名/电话（供定标成交结果表调用）',
  },
  {
    key: 'erp_plan',
    order: 2,
    label: '采购计划（ERP 导出）',
    stage: 'ERP 推送',
    description: '自动识别执行系统（ECP/ETP）、采购方式及采购申请标识，打通申报表与需求表',
  },
  {
    key: 'variable_template',
    order: 3,
    label: '采购需求变量模板',
    stage: '编制文件',
    description: '一键回填批次名称、编号、公告日期、获取文件时间、截止/开标时间与代理联系方式',
  },
  {
    key: 'requirement_list',
    order: 4,
    label: '需求一览表',
    stage: '挂网公告',
    description: '合并分标编号、包号、权重比例（自动校验 Excel 时间小数截断）及详评模板',
  },
  {
    key: 'supplier_contacts',
    order: 5,
    label: '供应商联系人信息表',
    stage: '系统报名后导出',
    description: '自动去除文本空格入库供应商台账，执行围串标风险初筛并供交接单调用',
  },
] as const;

const storage = multer.diskStorage({
  destination: (req, _file, cb) => {
    const batchId = String(req.params.batchId || 'common');
    const dir = path.join(UPLOADS_DIR, batchId, 'bidding');
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

/**
 * Find matching existing project row in batch_projects by:
 * 1. req_app_no (采购申请号 - primary key once generated in ERP)
 * 2. tech_spec_code (技术规范书编码)
 * 3. project_code (项目编码)
 * 4. space-stripped project name (项目名称 - ONLY as fallback when early offline sheets lack 采购申请号,
 *    and NEVER merging two rows that have different non-empty 采购申请号 or different 分标+包号!)
 */
function hasConflict(existingRow: Record<string, unknown>, incoming: Record<string, unknown>): boolean {
  const eReq = stripAllSpaces(existingRow.req_app_no);
  const iReq = stripAllSpaces(incoming.req_app_no);
  if (eReq && iReq && !isSameReqAppNo(eReq, iReq)) return true;

  const eLot = stripAllSpaces(existingRow.lot_code);
  const iLot = stripAllSpaces(incoming.lot_code);
  const ePkg = stripAllSpaces(existingRow.package_no);
  const iPkg = stripAllSpaces(incoming.package_no);
  if (eLot && iLot && ePkg && iPkg && (eLot !== iLot || ePkg !== iPkg)) return true;

  return false;
}

function findMatchingProject(
  existing: Array<Record<string, unknown>>,
  incoming: Record<string, unknown>
): Record<string, unknown> | undefined {
  const reqNo = stripAllSpaces(incoming.req_app_no);
  const techCode = stripAllSpaces(incoming.tech_spec_code);
  const projCode = stripAllSpaces(incoming.project_code);
  const normName = stripAllSpaces(incoming.name);

  if (reqNo) {
    const exact = existing.find((e) => stripAllSpaces(e.req_app_no) === reqNo);
    if (exact) return exact;
    const compat = existing.find((e) => isSameReqAppNo(e.req_app_no, reqNo));
    if (compat) return compat;
  }
  if (techCode) {
    const m = existing.find((e) => stripAllSpaces(e.tech_spec_code) === techCode && !hasConflict(e, incoming));
    if (m) return m;
  }
  if (projCode) {
    const m = existing.find((e) => stripAllSpaces(e.project_code) === projCode && !hasConflict(e, incoming));
    if (m) return m;
  }
  if (normName) {
    return existing.find((e) => stripAllSpaces(e.name) === normName && !hasConflict(e, incoming));
  }
  return undefined;
}

interface ErpConflictFieldDiff {
  field: string;
  reqVal: string;
  erpVal: string;
}

interface ErpConflictItem {
  type: 'erp_extra' | 'req_missing' | 'field_conflict';
  req_app_no: string;
  lot_info: string;
  project_name: string;
  fields?: ErpConflictFieldDiff[];
  summary: string;
}

function normalizeTextForCompare(val: unknown): string {
  return String(val ?? '')
    .replace(/[\s\u3000\u00A0]+/g, '')
    .replace(/[，,；;：:。.、（）()\[\]【】“”"'‘’]/g, '')
    .toLowerCase();
}

function fmtYuanShort(yuan: unknown): string {
  if (yuan === null || yuan === undefined || yuan === '') return '—';
  const n = Number(yuan);
  if (!Number.isFinite(n)) return String(yuan);
  return `${Number((n / 10000).toFixed(4))}万元`;
}

function hasUploadedSlot(batchId: string, slotKey: string): boolean {
  const row = db
    .prepare("SELECT id FROM uploaded_files WHERE batch_id = ? AND module = 'bidding' AND slot_key = ? LIMIT 1")
    .get(batchId, slotKey) as { id?: string } | undefined;
  return Boolean(row?.id);
}

/**
 * Compare authoritative 需求一览表 against all uploaded ERP 采购计划 files (主业 + 产业).
 * Persists structured conflicts into batches.erp_conflicts_json and returns concise banner warnings.
 */
function auditRequirementVsErp(batchId: string, reqRowsOverride?: Array<Record<string, unknown>>): {
  conflicts: ErpConflictItem[];
  warnings: string[];
} {
  const erpFiles = db
    .prepare("SELECT stored_path FROM uploaded_files WHERE batch_id = ? AND module = 'bidding' AND slot_key = 'erp_plan' ORDER BY created_at ASC")
    .all(batchId) as Array<{ stored_path?: string }>;

  let reqRows = reqRowsOverride;
  if (!reqRows) {
    const reqFile = db
      .prepare("SELECT stored_path FROM uploaded_files WHERE batch_id = ? AND module = 'bidding' AND slot_key = 'requirement_list' ORDER BY created_at DESC LIMIT 1")
      .get(batchId) as { stored_path?: string } | undefined;
    if (reqFile?.stored_path) {
      const absReq = path.join(UPLOADS_DIR, reqFile.stored_path);
      if (fs.existsSync(absReq)) {
        try {
          reqRows = parseRequirementList(absReq).rows;
        } catch {
          reqRows = undefined;
        }
      }
    }
  }

  if (!reqRows || reqRows.length === 0 || erpFiles.length === 0) {
    db.prepare("UPDATE batches SET erp_conflicts_json = '[]' WHERE id = ?").run(batchId);
    return { conflicts: [], warnings: [] };
  }

  const erpRows: Array<Record<string, unknown>> = [];
  for (const ef of erpFiles) {
    if (!ef.stored_path) continue;
    const absErp = path.join(UPLOADS_DIR, ef.stored_path);
    if (!fs.existsSync(absErp)) continue;
    try {
      erpRows.push(...parseErpPlan(absErp).rows);
    } catch {
      // ignore broken file
    }
  }
  if (erpRows.length === 0) {
    db.prepare("UPDATE batches SET erp_conflicts_json = '[]' WHERE id = ?").run(batchId);
    return { conflicts: [], warnings: [] };
  }

  const conflicts: ErpConflictItem[] = [];
  const matchedErpIdx = new Set<number>();

  // Build lookup index on ERP rows by exact req_app_no + 10-digit SAP prefix
  const erpByExact = new Map<string, number>();
  const erpByPrefix = new Map<string, number>();
  erpRows.forEach((e, idx) => {
    const pk = stripAllSpaces(e.req_app_no);
    if (!pk) return;
    if (!erpByExact.has(pk)) erpByExact.set(pk, idx);
    if (pk.length === 15 && !erpByPrefix.has(pk.slice(0, 10))) {
      erpByPrefix.set(pk.slice(0, 10), idx);
    }
  });

  for (const r of reqRows) {
    const rPk = stripAllSpaces(r.req_app_no);
    const lotInfo = [r.lot_name || r.lot_code, r.package_no].filter(Boolean).join(' · ');
    const rName = String(r.name || '');

    let eIdx: number | undefined;
    if (rPk) {
      eIdx = erpByExact.get(rPk) ?? erpByPrefix.get(rPk) ?? (rPk.length === 15 ? erpByExact.get(rPk.slice(0, 10)) : undefined);
    }
    if (eIdx === undefined) {
      conflicts.push({
        type: 'req_missing',
        req_app_no: rPk || '无采购申请号',
        lot_info: lotInfo,
        project_name: rName,
        summary: '需求一览表包含此项目，但已上传的 ERP 采购计划中未找到对应采购申请号',
      });
      continue;
    }

    matchedErpIdx.add(eIdx);
    const e = erpRows[eIdx];
    const diffs: ErpConflictFieldDiff[] = [];

    // 1. 项目名称
    if (rName && String(e.name || '') && normalizeTextForCompare(rName) !== normalizeTextForCompare(e.name)) {
      diffs.push({ field: '项目名称', reqVal: rName, erpVal: String(e.name || '') });
    }
    // 2. 项目概况（仅双方均非空时比对，忽略中英文标点与空格差异）
    const rOv = String(r.project_overview || '').trim();
    const eOv = String(e.project_overview || '').trim();
    if (rOv && eOv && normalizeTextForCompare(rOv) !== normalizeTextForCompare(eOv)) {
      diffs.push({ field: '项目概况', reqVal: rOv, erpVal: eOv });
    }
    // 3. 预算
    const rBudget = r.budget_yuan !== null && r.budget_yuan !== undefined ? Number(r.budget_yuan) : null;
    const eBudget = e.budget_yuan !== null && e.budget_yuan !== undefined ? Number(e.budget_yuan) : null;
    if (rBudget !== null && eBudget !== null && Math.abs(rBudget - eBudget) > 0.01) {
      diffs.push({ field: '预算金额', reqVal: fmtYuanShort(rBudget), erpVal: fmtYuanShort(eBudget) });
    }
    // 4. 限价（非折扣率报价时比对，若与预算冲突已报且限价等于预算则不重复刷屏）
    const rMaxRaw = String(r.max_price_raw || '').trim();
    const rMax = r.max_price_yuan !== null && r.max_price_yuan !== undefined ? Number(r.max_price_yuan) : null;
    if (!rMaxRaw && rMax !== null && eBudget !== null && Math.abs(rMax - eBudget) > 0.01) {
      if (rBudget === null || Math.abs(rMax - rBudget) > 0.01) {
        diffs.push({ field: '最高限价', reqVal: fmtYuanShort(rMax), erpVal: fmtYuanShort(eBudget) });
      }
    }
    // 5. 技术规范书编码
    const rTech = stripAllSpaces(r.tech_spec_code).toUpperCase();
    const eTech = stripAllSpaces(e.tech_spec_code).toUpperCase();
    if (rTech && eTech && rTech !== eTech) {
      diffs.push({ field: '技术规范书编码', reqVal: rTech, erpVal: eTech });
    }

    if (diffs.length > 0) {
      conflicts.push({
        type: 'field_conflict',
        req_app_no: rPk,
        lot_info: lotInfo,
        project_name: rName,
        fields: diffs,
        summary: diffs.map((d) => `${d.field}不一致`).join('、'),
      });
    }
  }

  // ERP extra rows not present in 需求一览表
  erpRows.forEach((e, idx) => {
    if (matchedErpIdx.has(idx)) return;
    conflicts.push({
      type: 'erp_extra',
      req_app_no: stripAllSpaces(e.req_app_no) || '无采购申请号',
      lot_info: String(e.project_unit || ''),
      project_name: String(e.name || ''),
      summary: 'ERP 采购计划多出此项目（未在最终需求一览表中，已自动过滤）',
    });
  });

  // 排序：数量差异（ERP缺失 -> ERP多出）置顶，字段冲突排在后面
  const typePriority: Record<ErpConflictItem['type'], number> = {
    req_missing: 0,
    erp_extra: 1,
    field_conflict: 2,
  };
  conflicts.sort((a, b) => typePriority[a.type] - typePriority[b.type]);

  db.prepare('UPDATE batches SET erp_conflicts_json = ? WHERE id = ?').run(JSON.stringify(conflicts), batchId);

  const warnings: string[] = [];
  const extraCount = conflicts.filter((c) => c.type === 'erp_extra').length;
  const missingCount = conflicts.filter((c) => c.type === 'req_missing').length;
  const fieldCount = conflicts.filter((c) => c.type === 'field_conflict').length;
  if (extraCount > 0) {
    warnings.push(`数量差异：ERP 采购计划多出 ${extraCount} 条不在《需求一览表》中的项目，已以《需求一览表》为准自动过滤（可在下方差异预警面板查看明细）`);
  }
  if (missingCount > 0) {
    warnings.push(`数量差异：《需求一览表》中有 ${missingCount} 条项目在已上传的 ERP 采购计划中未匹配到采购申请号（若属联合采购，请确认主业/产业两份 ERP 均已上传）`);
  }
  if (fieldCount > 0) {
    warnings.push(`字段冲突：检测到 ${fieldCount} 条项目在《需求一览表》与 ERP 采购计划之间存在关键信息差异（项目名称/概况/预算/限价/技术规范书编码），当前均已以《需求一览表》为准展示`);
  }
  return { conflicts, warnings };
}

function upsertMergedProjects(
  batchId: string,
  incomingRows: Array<Record<string, unknown>>,
  opts?: { mode?: 'merge' | 'enrich_only' | 'authoritative' }
) {
  const mode = opts?.mode || 'merge';
  const existing = db.prepare('SELECT * FROM batch_projects WHERE batch_id = ? ORDER BY seq ASC').all(batchId) as Array<Record<string, unknown>>;
  const now = new Date().toISOString();

  if (mode === 'authoritative') {
    // 以《需求一览表》为权威基准：保留其精确项目数量与核心字段，仅从既有申报表/ERP 记录继承补充字段
    db.prepare('DELETE FROM batch_projects WHERE batch_id = ?').run(batchId);
    const usedExistingIds = new Set<string>();
    let seq = 0;
    let inserted = 0;
    let updated = 0;

    for (const row of incomingRows) {
      seq++;
      const match = findMatchingProject(existing, row);
      if (match?.id) usedExistingIds.add(String(match.id));
      const budgetYuan = (row.budget_yuan as number | null | undefined) ?? (match?.budget_yuan as number | null | undefined) ?? null;
      const maxPriceRaw = String(row.max_price_raw || '');
      const maxPriceYuan = maxPriceRaw
        ? null
        : (row.max_price_yuan as number | null | undefined) ?? budgetYuan ?? (match?.max_price_yuan as number | null | undefined) ?? null;
      const budgetText = budgetYuan !== null ? formatYuanAsWan(budgetYuan) : String(match?.budget || '');
      const codeStr = row.lot_code
        ? `${row.lot_code}-${row.package_no || ''}`
        : String(match?.code || row.project_code || '');
      const newId = uuid();
      db.prepare(`
        INSERT INTO batch_projects (
          id, batch_id, seq, name, code, budget, procurement_method, department, notes, created_at,
          project_code, req_app_no, tech_spec_code, lot_code, lot_name, package_no,
          project_unit, project_overview, qualification_req, performance_req, staff_req, quote_method, subcontract_rule,
          budget_yuan, max_price_yuan, max_price_raw, declarant_name, declarant_phone,
          weight_ratio, weight_warning, weight_suggested, price_formula, price_params,
          biz_eval_tpl, tech_eval_tpl, contract_version, funding_source, duration, erp_sync_status, material_code, material_desc
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, '', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `).run(
        newId, batchId, seq,
        String(row.name || match?.name || ''),
        codeStr,
        budgetText,
        String(row.procurement_method || match?.procurement_method || '竞争性谈判'),
        String(row.department || match?.department || ''),
        now,
        String(row.project_code || match?.project_code || ''),
        String(row.req_app_no || match?.req_app_no || ''),
        String(row.tech_spec_code || match?.tech_spec_code || ''),
        String(row.lot_code || ''),
        String(row.lot_name || ''),
        String(row.package_no || ''),
        String(row.project_unit || match?.project_unit || ''),
        String(row.project_overview || match?.project_overview || ''),
        String(row.qualification_req || match?.qualification_req || ''),
        String(row.performance_req || match?.performance_req || ''),
        String(row.staff_req || match?.staff_req || ''),
        String(row.quote_method || match?.quote_method || ''),
        String(row.subcontract_rule || match?.subcontract_rule || ''),
        budgetYuan, maxPriceYuan, maxPriceRaw,
        String(row.declarant_name || match?.declarant_name || ''),
        String(row.declarant_phone || match?.declarant_phone || ''),
        String(row.weight_ratio ?? ''),
        String(row.weight_warning ?? ''),
        String(row.weight_suggested ?? ''),
        String(row.price_formula || ''),
        String(row.price_params || ''),
        String(row.biz_eval_tpl || ''),
        String(row.tech_eval_tpl || ''),
        String(row.contract_version || match?.contract_version || ''),
        String(row.funding_source || match?.funding_source || ''),
        String(row.duration || match?.duration || ''),
        String(row.erp_sync_status || match?.erp_sync_status || ''),
        String(row.material_code || match?.material_code || ''),
        String(row.material_desc || match?.material_desc || '')
      );
      if (match) updated++;
      else inserted++;
    }
    const pruned = Math.max(0, existing.length - usedExistingIds.size);
    return { inserted, updated, pruned };
  }

  let maxSeq = existing.reduce((m, r) => Math.max(m, Number(r.seq || 0)), 0);
  let inserted = 0;
  let updated = 0;
  let skippedExtra = 0;

  for (const row of incomingRows) {
    const match = findMatchingProject(existing, row);
    if (mode === 'enrich_only') {
      if (!match) {
        skippedExtra++;
        continue;
      }
      // 已有《需求一览表》时，仅补充缺失字段与 ERP/申报专属字段，绝不覆盖《需求一览表》的核心字段
      const budgetYuan = (match.budget_yuan as number | null | undefined) ?? (row.budget_yuan as number | null | undefined) ?? null;
      const maxPriceRaw = String(match.max_price_raw || row.max_price_raw || '');
      const maxPriceYuan = maxPriceRaw
        ? null
        : (match.max_price_yuan as number | null | undefined) ?? (row.max_price_yuan as number | null | undefined) ?? budgetYuan;
      const budgetText = String(match.budget || '') || (budgetYuan !== null ? formatYuanAsWan(budgetYuan) : '');
      db.prepare(`
        UPDATE batch_projects SET
          department = ?,
          project_code = ?,
          req_app_no = ?,
          tech_spec_code = ?,
          project_unit = ?,
          project_overview = ?,
          budget = ?,
          budget_yuan = ?,
          max_price_yuan = ?,
          max_price_raw = ?,
          declarant_name = ?,
          declarant_phone = ?,
          funding_source = ?,
          duration = ?,
          erp_sync_status = ?,
          material_code = ?,
          material_desc = ?
        WHERE id = ?
      `).run(
        String(match.department || row.department || ''),
        String(match.project_code || row.project_code || ''),
        String(match.req_app_no || row.req_app_no || ''),
        String(match.tech_spec_code || row.tech_spec_code || ''),
        String(match.project_unit || row.project_unit || ''),
        String(match.project_overview || row.project_overview || ''),
        budgetText,
        budgetYuan,
        maxPriceYuan,
        maxPriceRaw,
        String(match.declarant_name || row.declarant_name || ''),
        String(match.declarant_phone || row.declarant_phone || ''),
        String(match.funding_source || row.funding_source || ''),
        String(match.duration || row.duration || ''),
        String(row.erp_sync_status || match.erp_sync_status || ''),
        String(row.material_code || match.material_code || ''),
        String(row.material_desc || match.material_desc || ''),
        String(match.id)
      );
      updated++;
      continue;
    }

    const budgetYuan = (row.budget_yuan as number | null | undefined) ?? (match?.budget_yuan as number | null | undefined) ?? null;
    const maxPriceRaw = String(row.max_price_raw || match?.max_price_raw || '');
    const maxPriceYuan = maxPriceRaw
      ? null
      : (row.max_price_yuan as number | null | undefined) ?? (match?.max_price_yuan as number | null | undefined) ?? budgetYuan;
    const budgetText = budgetYuan !== null ? formatYuanAsWan(budgetYuan) : String(match?.budget || '');

    if (match) {
      db.prepare(`
        UPDATE batch_projects SET
          name = ?,
          code = ?,
          budget = ?,
          procurement_method = ?,
          department = ?,
          project_code = ?,
          req_app_no = ?,
          tech_spec_code = ?,
          lot_code = ?,
          lot_name = ?,
          package_no = ?,
          project_unit = ?,
          project_overview = ?,
          qualification_req = ?,
          performance_req = ?,
          staff_req = ?,
          quote_method = ?,
          subcontract_rule = ?,
          budget_yuan = ?,
          max_price_yuan = ?,
          max_price_raw = ?,
          declarant_name = ?,
          declarant_phone = ?,
          weight_ratio = ?,
          weight_warning = ?,
          weight_suggested = ?,
          price_formula = ?,
          price_params = ?,
          biz_eval_tpl = ?,
          tech_eval_tpl = ?,
          contract_version = ?,
          funding_source = ?,
          duration = ?,
          erp_sync_status = ?,
          material_code = ?,
          material_desc = ?
        WHERE id = ?
      `).run(
        String(row.name || match.name),
        String(row.lot_code ? `${row.lot_code}-${row.package_no || ''}` : match.code || row.project_code || ''),
        budgetText,
        String(row.procurement_method || match.procurement_method || '竞争性谈判'),
        String(row.department || match.department || ''),
        String(row.project_code || match.project_code || ''),
        String(row.req_app_no || match.req_app_no || ''),
        String(row.tech_spec_code || match.tech_spec_code || ''),
        String(row.lot_code || match.lot_code || ''),
        String(row.lot_name || match.lot_name || ''),
        String(row.package_no || match.package_no || ''),
        String(row.project_unit || match.project_unit || ''),
        String(row.project_overview || match.project_overview || ''),
        String(row.qualification_req || match.qualification_req || ''),
        String(row.performance_req || match.performance_req || ''),
        String(row.staff_req || match.staff_req || ''),
        String(row.quote_method || match.quote_method || ''),
        String(row.subcontract_rule || match.subcontract_rule || ''),
        budgetYuan,
        maxPriceYuan,
        String(row.max_price_raw || match.max_price_raw || ''),
        String(row.declarant_name || match.declarant_name || ''),
        String(row.declarant_phone || match.declarant_phone || ''),
        String(row.weight_ratio !== undefined ? row.weight_ratio : match.weight_ratio || ''),
        String(row.weight_warning !== undefined ? row.weight_warning : match.weight_warning || ''),
        String(row.weight_suggested !== undefined ? row.weight_suggested : match.weight_suggested || ''),
        String(row.price_formula || match.price_formula || ''),
        String(row.price_params || match.price_params || ''),
        String(row.biz_eval_tpl || match.biz_eval_tpl || ''),
        String(row.tech_eval_tpl || match.tech_eval_tpl || ''),
        String(row.contract_version || match.contract_version || ''),
        String(row.funding_source || match.funding_source || ''),
        String(row.duration || match.duration || ''),
        String(row.erp_sync_status || match.erp_sync_status || ''),
        String(row.material_code || match.material_code || ''),
        String(row.material_desc || match.material_desc || ''),
        String(match.id)
      );
      updated++;
    } else {
      maxSeq++;
      const newId = uuid();
      const codeStr = row.lot_code
        ? `${row.lot_code}-${row.package_no || ''}`
        : String(row.project_code || '');
      db.prepare(`
        INSERT INTO batch_projects (
          id, batch_id, seq, name, code, budget, procurement_method, department, notes, created_at,
          project_code, req_app_no, tech_spec_code, lot_code, lot_name, package_no,
          project_unit, project_overview, qualification_req, performance_req, staff_req, quote_method, subcontract_rule,
          budget_yuan, max_price_yuan, max_price_raw, declarant_name, declarant_phone,
          weight_ratio, weight_warning, weight_suggested, price_formula, price_params,
          biz_eval_tpl, tech_eval_tpl, contract_version, funding_source, duration, erp_sync_status, material_code, material_desc
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, '', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `).run(
        newId, batchId, maxSeq, String(row.name || ''), codeStr, budgetText,
        String(row.procurement_method || '竞争性谈判'), String(row.department || ''), now,
        String(row.project_code || ''), String(row.req_app_no || ''), String(row.tech_spec_code || ''),
        String(row.lot_code || ''), String(row.lot_name || ''), String(row.package_no || ''),
        String(row.project_unit || ''), String(row.project_overview || ''), String(row.qualification_req || ''),
        String(row.performance_req || ''), String(row.staff_req || ''), String(row.quote_method || ''),
        String(row.subcontract_rule || ''),
        budgetYuan, maxPriceYuan, String(row.max_price_raw || ''), String(row.declarant_name || ''), String(row.declarant_phone || ''),
        String(row.weight_ratio || ''), String(row.weight_warning || ''), String(row.weight_suggested || ''),
        String(row.price_formula || ''), String(row.price_params || ''),
        String(row.biz_eval_tpl || ''), String(row.tech_eval_tpl || ''),
        String(row.contract_version || ''), String(row.funding_source || ''), String(row.duration || ''),
        String(row.erp_sync_status || ''), String(row.material_code || ''), String(row.material_desc || '')
      );
      existing.push({ id: newId, ...row });
      inserted++;
    }
  }

  return { inserted, updated, pruned: skippedExtra };
}

// GET /api/bidding/batch/:batchId/workspace
biddingRouter.get('/batch/:batchId/workspace', (req, res) => {
  const batchId = req.params.batchId;
  const batch = db.prepare('SELECT * FROM batches WHERE id = ?').get(batchId) as Record<string, unknown> | undefined;
  if (!batch) return res.status(404).json({ error: '未找到该批次' });

  const projects = db.prepare('SELECT * FROM batch_projects WHERE batch_id = ? ORDER BY seq ASC, created_at ASC').all(batchId);
  const files = db.prepare("SELECT * FROM uploaded_files WHERE batch_id = ? AND module = 'bidding' ORDER BY created_at DESC").all(batchId) as Array<Record<string, unknown>>;
  const suppliers = db.prepare('SELECT * FROM batch_suppliers WHERE batch_id = ? ORDER BY seq ASC').all(batchId) as Array<{
    id: string;
    supplier_name: string;
    contact_name: string;
    mobile: string;
    phone: string;
    email: string;
    address: string;
  }>;

  const formattedFiles = files.map((f) => ({
    ...f,
    slot_key: String(f.slot_key || ''),
    parsed: f.parsed_json ? JSON.parse(String(f.parsed_json)) : null,
  }));

  const slots = BIDDING_SLOTS.map((s) => {
    const slotFiles = formattedFiles.filter((f) => f.slot_key === s.key);
    return {
      ...s,
      file: slotFiles[0] || null,
      files: slotFiles,
    };
  });

  const collusionRisks = detectCollusionRisks(suppliers);
  let extraFields = [];
  try {
    extraFields = JSON.parse(String(batch.extra_fields_json || '[]'));
  } catch {
    extraFields = [];
  }

  let erpConflicts: ErpConflictItem[] = [];
  try {
    erpConflicts = JSON.parse(String(batch.erp_conflicts_json || '[]'));
  } catch {
    erpConflicts = [];
  }

  const rawNotices = db
    .prepare('SELECT * FROM batch_change_notices WHERE batch_id = ? ORDER BY notice_date DESC, created_at DESC')
    .all(batchId) as Array<Record<string, unknown>>;
  const changeNotices = rawNotices.map((n) => {
    let change_types: string[] = [];
    let project_ids: string[] = [];
    try { change_types = JSON.parse(String(n.change_types_json || '[]')); } catch { change_types = []; }
    try { project_ids = JSON.parse(String(n.project_ids_json || '[]')); } catch { project_ids = []; }
    return {
      id: String(n.id),
      batch_id: String(n.batch_id),
      title: String(n.title || ''),
      notice_date: String(n.notice_date || ''),
      change_types,
      summary: String(n.summary || ''),
      project_ids,
      file_original_name: String(n.file_original_name || ''),
      file_mime_type: String(n.file_mime_type || ''),
      has_file: Boolean(n.file_stored_path),
      created_at: String(n.created_at || ''),
    };
  });

  res.json({
    batch: { ...batch, extra_fields: extraFields },
    slots,
    files: formattedFiles,
    projects,
    suppliers,
    collusionRisks,
    erpConflicts,
    changeNotices,
  });
});

// POST /api/bidding/batch/:batchId/upload-slot/:slotKey
biddingRouter.post('/batch/:batchId/upload-slot/:slotKey', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: '请选择要上传的 Excel 文件' });

  const batchId = req.params.batchId;
  const slotKey = req.params.slotKey;
  const slotDef = BIDDING_SLOTS.find((s) => s.key === slotKey);
  if (!slotDef) return res.status(400).json({ error: '无效的文件槽位' });

  const originalName = decodeName(req.file.originalname);
  const relPath = path.relative(UPLOADS_DIR, req.file.path);

  let parsedPreviewJson: string | null = null;
  try {
    const tables = parseSpreadsheetTables(req.file.path);
    if (tables) parsedPreviewJson = JSON.stringify(tables);
  } catch {
    // non-fatal
  }

  // Save uploaded file record
  const fileId = uuid();
  const now = new Date().toISOString();
  db.prepare(`
    INSERT INTO uploaded_files (id, batch_id, module, slot_key, category, original_name, stored_path, mime_type, file_size, parsed_json, notes, created_at)
    VALUES (?, ?, 'bidding', ?, ?, ?, ?, ?, ?, ?, '', ?)
  `).run(
    fileId, batchId, slotKey, slotDef.label, originalName, relPath,
    req.file.mimetype || 'application/octet-stream', req.file.size, parsedPreviewJson, now
  );

  // Execute slot-specific automated extraction & data cleaning
  const report: {
    slotLabel: string;
    summary: string;
    warnings: string[];
  } = {
    slotLabel: slotDef.label,
    summary: '',
    warnings: [],
  };

  try {
    if (slotKey === 'variable_template') {
      const parsed = parseVariableTemplate(req.file.path);
      const existing = db.prepare('SELECT * FROM batches WHERE id = ?').get(batchId) as Record<string, string>;
      const p = parsed.batchPatch;
      db.prepare(`
        UPDATE batches SET
          name = ?, code = ?, announcement_time = ?, doc_fetch_time = ?, bid_opening_time = ?,
          contact_phone = ?, contact_email = ?, procurement_type = ?, updated_at = ?
        WHERE id = ?
      `).run(
        p.name || existing.name,
        p.code || existing.code,
        p.announcement_time || existing.announcement_time,
        p.doc_fetch_time || existing.doc_fetch_time,
        p.bid_opening_time || existing.bid_opening_time,
        p.contact_phone || existing.contact_phone,
        p.contact_email || existing.contact_email,
        p.procurement_type || existing.procurement_type,
        now,
        batchId
      );
      report.summary = `已自动回填批次名称、编号(${p.code || '-'})、公告与截止时间及代理联系方式`;
      report.warnings = parsed.warnings;
    } else if (slotKey === 'declaration_plan') {
      const parsed = parseDeclarationPlan(req.file.path);
      const hasReq = hasUploadedSlot(batchId, 'requirement_list');
      const { inserted, updated, pruned } = upsertMergedProjects(batchId, parsed.rows, {
        mode: hasReq ? 'enrich_only' : 'merge',
      });
      report.summary = hasReq
        ? `已解析 ${parsed.rows.length} 条申报计划（按《需求一览表》为准补充合并 ${updated} 条，忽略表外 ${pruned} 条）`
        : `已解析 ${parsed.rows.length} 条申报计划（新增 ${inserted} 条、合并更新 ${updated} 条），并完成申报人姓名/电话拆分与价格单位对齐`;
    } else if (slotKey === 'erp_plan') {
      const parsed = parseErpPlan(req.file.path);
      if (parsed.batchMeta.code) {
        const existing = db.prepare('SELECT * FROM batches WHERE id = ?').get(batchId) as Record<string, string>;
        // 联合采购批次若已是主业(ECP)执行，上传产业ERP时保持主业(ECP)平台不变
        const nextUnitNature =
          existing.unit_nature === '主业' || parsed.batchMeta.unit_nature === '主业'
            ? '主业'
            : parsed.batchMeta.unit_nature || existing.unit_nature;
        db.prepare(`
          UPDATE batches SET
            code = ?, name = ?, project_unit = ?, unit_nature = ?, procurement_type = ?, updated_at = ?
          WHERE id = ?
        `).run(
          parsed.batchMeta.code || existing.code,
          parsed.batchMeta.name || existing.name,
          existing.project_unit || parsed.batchMeta.project_unit,
          nextUnitNature,
          parsed.batchMeta.procurement_type || existing.procurement_type,
          now,
          batchId
        );
      }
      const hasReq = hasUploadedSlot(batchId, 'requirement_list');
      const { inserted, updated, pruned } = upsertMergedProjects(batchId, parsed.rows, {
        mode: hasReq ? 'enrich_only' : 'merge',
      });
      const audit = auditRequirementVsErp(batchId);
      report.summary = hasReq
        ? `已解析${parsed.erpSource} ERP 计划 ${parsed.rows.length} 条（以《需求一览表》为准：匹配补充 ${updated} 条，自动忽略表外 ${pruned} 条）`
        : `已解析${parsed.erpSource} ERP 计划 ${parsed.rows.length} 条（新增 ${inserted}、合并 ${updated}），执行平台：${parsed.batchMeta.unit_nature === '主业' ? 'ECP 2.0' : 'ETP'}`;
      report.warnings = [...parsed.warnings, ...audit.warnings];
    } else if (slotKey === 'requirement_list') {
      const parsed = parseRequirementList(req.file.path);
      const { inserted, updated, pruned } = upsertMergedProjects(batchId, parsed.rows, { mode: 'authoritative' });
      const audit = auditRequirementVsErp(batchId, parsed.rows);
      const mergeNote = parsed.mergesExpanded > 0 ? `，自动展开 ${parsed.mergesExpanded} 处合并单元格（含子包）` : '';
      const pruneNote = pruned > 0 ? `，已自动过滤前期表中不在需求一览表的 ${pruned} 条项目` : '';
      report.summary = `已以《需求一览表》为准载入 ${parsed.rows.length} 条项目（继承合并 ${updated} 条、新增 ${inserted} 条${mergeNote}${pruneNote}）`;
      report.warnings = [...parsed.warnings, ...audit.warnings];
    } else if (slotKey === 'supplier_contacts') {
      const parsed = parseSupplierContacts(req.file.path);
      db.prepare('DELETE FROM batch_suppliers WHERE batch_id = ?').run(batchId);
      const ins = db.prepare(`
        INSERT INTO batch_suppliers (
          id, batch_id, seq, supplier_name, package_no, contact_name, mobile, phone, email, fax, address, postal_code, tax_no, invoice_title, bank_name, bank_account, created_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `);
      parsed.suppliers.forEach((s) => {
        ins.run(
          uuid(), batchId, s.seq, s.supplier_name, s.package_no, s.contact_name, s.mobile, s.phone,
          s.email, s.fax, s.address, s.postal_code, s.tax_no, s.invoice_title, s.bank_name, s.bank_account, now
        );
      });
      report.summary = `已清洗并入库 ${parsed.suppliers.length} 家报名供应商（自动去除 ${parsed.spacesCleanedCount} 处多余空格）`;
      if (parsed.risks.length > 0) {
        report.warnings = parsed.risks.map((rk) => `${rk.label}（${rk.value}）：${rk.suppliers.join('、')}`);
      }
    }
  } catch (err) {
    report.warnings.push(`自动提取时遇到格式异常：${(err as Error).message}`);
  }

  res.status(201).json({ ok: true, fileId, report });
});

// DELETE /api/bidding/batch/:batchId/slot/:slotKey — clear parsed data & uploaded file(s) for a specific bidding slot
biddingRouter.delete('/batch/:batchId/slot/:slotKey', (req, res) => {
  const { batchId, slotKey } = req.params;
  const slotDef = BIDDING_SLOTS.find((s) => s.key === slotKey);
  if (!slotDef) return res.status(400).json({ error: '无效的招标文件槽位' });

  const slotFiles = db
    .prepare("SELECT id, stored_path FROM uploaded_files WHERE batch_id = ? AND module = 'bidding' AND slot_key = ?")
    .all(batchId, slotKey) as Array<{ id: string; stored_path: string }>;

  for (const f of slotFiles) {
    if (f.stored_path) {
      const absPath = path.join(UPLOADS_DIR, f.stored_path);
      if (fs.existsSync(absPath)) fs.unlinkSync(absPath);
    }
  }
  db.prepare("DELETE FROM uploaded_files WHERE batch_id = ? AND module = 'bidding' AND slot_key = ?").run(batchId, slotKey);

  if (slotKey === 'supplier_contacts') {
    db.prepare('DELETE FROM batch_suppliers WHERE batch_id = ?').run(batchId);
  } else if (slotKey === 'variable_template') {
    db.prepare(`
      UPDATE batches
      SET announcement_time = '', doc_fetch_time = '', bid_opening_time = '', contact_phone = '', contact_email = '', updated_at = ?
      WHERE id = ?
    `).run(new Date().toISOString(), batchId);
  } else if (slotKey === 'declaration_plan' || slotKey === 'erp_plan' || slotKey === 'requirement_list') {
    // Clear batch_projects and rebuild from any remaining uploaded project-source slots in SOP order
    db.prepare('DELETE FROM batch_projects WHERE batch_id = ?').run(batchId);
    const orderedProjectSlots = ['declaration_plan', 'erp_plan', 'requirement_list'] as const;
    for (const sk of orderedProjectSlots) {
      const filesToReplay =
        sk === 'erp_plan'
          ? (db
              .prepare("SELECT stored_path FROM uploaded_files WHERE batch_id = ? AND module = 'bidding' AND slot_key = ? ORDER BY created_at ASC")
              .all(batchId, sk) as Array<{ stored_path?: string }>)
          : (db
              .prepare("SELECT stored_path FROM uploaded_files WHERE batch_id = ? AND module = 'bidding' AND slot_key = ? ORDER BY created_at DESC LIMIT 1")
              .all(batchId, sk) as Array<{ stored_path?: string }>);
      for (const f of filesToReplay) {
        if (!f?.stored_path) continue;
        const absPath = path.join(UPLOADS_DIR, f.stored_path);
        if (!fs.existsSync(absPath)) continue;
        try {
          if (sk === 'declaration_plan') {
            upsertMergedProjects(batchId, parseDeclarationPlan(absPath).rows);
          } else if (sk === 'erp_plan') {
            upsertMergedProjects(batchId, parseErpPlan(absPath).rows);
          } else if (sk === 'requirement_list') {
            upsertMergedProjects(batchId, parseRequirementList(absPath).rows, { mode: 'authoritative' });
          }
        } catch {
          // ignore non-fatal rebuild errors
        }
      }
    }
    auditRequirementVsErp(batchId);
  }

  res.json({ ok: true, slotLabel: slotDef.label });
});

// GET /api/bidding/batch/:batchId/erp-conflicts/export — export ALL 需求一览表 vs ERP 差异明细 as formatted .xlsx
biddingRouter.get('/batch/:batchId/erp-conflicts/export', async (req, res) => {
  try {
    const batchId = req.params.batchId;
    const batch = db.prepare('SELECT code, name, erp_conflicts_json FROM batches WHERE id = ?').get(batchId) as
      | { code?: string; name?: string; erp_conflicts_json?: string }
      | undefined;
    if (!batch) return res.status(404).json({ error: '未找到该批次' });

    let conflicts: ErpConflictItem[] = [];
    try {
      conflicts = JSON.parse(String(batch.erp_conflicts_json || '[]'));
    } catch {
      conflicts = [];
    }
    if (conflicts.length === 0) {
      return res.status(400).json({ error: '当前批次暂无可导出的差异记录' });
    }

    // 无论历史缓存顺序如何，导出时强制将数量差异（ERP缺失 -> ERP多出）排在最前面
    const priority: Record<ErpConflictItem['type'], number> = {
      req_missing: 0,
      erp_extra: 1,
      field_conflict: 2,
    };
    const sortedConflicts = [...conflicts].sort((a, b) => (priority[a.type] ?? 9) - (priority[b.type] ?? 9));
    const countDiffs = sortedConflicts.filter((c) => c.type === 'req_missing' || c.type === 'erp_extra');
    const fieldDiffs = sortedConflicts.filter((c) => c.type === 'field_conflict');

    const wb = new ExcelJS.Workbook();
    wb.creator = '国网招标代理 SOP 工作台';

    const borderThin: Partial<ExcelJS.Borders> = {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } },
    };

    const styleSheet = (
      ws: ExcelJS.Worksheet,
      headers: string[],
      widths: number[],
      rowsData: unknown[][],
      centerCols: Set<number>
    ) => {
      ws.columns = headers.map((header, i) => ({ header, key: `c${i}`, width: widths[i] }));
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
    };

    const typeLabel = (t: ErpConflictItem['type']) =>
      t === 'erp_extra' ? 'ERP多出（未在需求一览表）' : t === 'req_missing' ? 'ERP缺失（需求一览表独有）' : '字段冲突';

    // Sheet 1: 全部差异汇总（数量差异置顶）
    const allRows: unknown[][] = [];
    let allSeq = 0;
    for (const item of sortedConflicts) {
      if (item.type === 'field_conflict' && item.fields && item.fields.length > 0) {
        for (const f of item.fields) {
          allSeq++;
          allRows.push([
            allSeq,
            typeLabel(item.type),
            item.req_app_no,
            item.lot_info || '',
            item.project_name || '',
            f.field,
            f.reqVal,
            f.erpVal,
            '已以《需求一览表》为准展示',
          ]);
        }
      } else {
        allSeq++;
        allRows.push([
          allSeq,
          typeLabel(item.type),
          item.req_app_no,
          item.lot_info || '',
          item.project_name || '',
          '项目数量差异',
          item.type === 'req_missing' ? '有此项目' : '无此项目',
          item.type === 'erp_extra' ? '有此项目' : '无此项目',
          item.summary,
        ]);
      }
    }
    const wsAll = wb.addWorksheet('全部差异汇总', {
      pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    });
    styleSheet(
      wsAll,
      ['序号', '差异类型', '采购申请号', '分标/单位', '项目名称', '冲突字段', '需求一览表（当前展示）', 'ERP采购计划', '说明'],
      [8, 22, 22, 24, 34, 16, 36, 36, 38],
      allRows,
      new Set([1, 2, 3, 6])
    );

    // Sheet 2: 数量差异（多出与缺失）
    const countRows: unknown[][] = countDiffs.map((item, idx) => [
      idx + 1,
      typeLabel(item.type),
      item.req_app_no,
      item.lot_info || '',
      item.project_name || '',
      item.summary,
    ]);
    const wsCount = wb.addWorksheet('数量差异（多出与缺失）', {
      pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    });
    styleSheet(
      wsCount,
      ['序号', '差异类型', '采购申请号', '分标/单位', '项目名称', '说明'],
      [8, 24, 22, 26, 40, 44],
      countRows,
      new Set([1, 2, 3])
    );

    // Sheet 3: 关键字段冲突明细
    const fieldRows: unknown[][] = [];
    let fieldSeq = 0;
    for (const item of fieldDiffs) {
      for (const f of item.fields || []) {
        fieldSeq++;
        fieldRows.push([
          fieldSeq,
          item.req_app_no,
          item.lot_info || '',
          item.project_name || '',
          f.field,
          f.reqVal,
          f.erpVal,
          '已以《需求一览表》为准展示',
        ]);
      }
    }
    const wsField = wb.addWorksheet('关键字段冲突明细', {
      pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    });
    styleSheet(
      wsField,
      ['序号', '采购申请号', '分标/单位', '项目名称', '冲突字段', '需求一览表（当前展示）', 'ERP采购计划', '说明'],
      [8, 22, 24, 34, 16, 38, 38, 28],
      fieldRows,
      new Set([1, 2, 5])
    );

    const safeCode = String(batch.code || 'BATCH').replace(/[^\w\u4e00-\u9fa5-]/g, '_');
    const filename = `${safeCode}_需求一览表与ERP差异明细.xlsx`;
    const buf = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.send(Buffer.from(buf));
  } catch (err) {
    res.status(500).json({ error: `导出差异明细失败：${(err as Error).message}` });
  }
});

// POST /api/bidding/batch/:batchId/import-projects - Incremental import from Excel with strict deduplication
biddingRouter.post('/batch/:batchId/import-projects', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: '请选择包含新增项目的 Excel 文件' });
  const batchId = req.params.batchId;

  try {
    const { detectedType, rows, warnings } = parseAnyProjectSpreadsheet(req.file.path);
    const existing = db.prepare('SELECT * FROM batch_projects WHERE batch_id = ? ORDER BY seq ASC').all(batchId) as Array<Record<string, unknown>>;

    const newOnlyRows: Array<Record<string, unknown>> = [];
    let skippedDuplicates = 0;

    for (const row of rows) {
      const match = findMatchingProject(existing, row) || findMatchingProject(newOnlyRows, row);
      if (match) {
        skippedDuplicates++;
      } else {
        newOnlyRows.push(row);
      }
    }

    const { inserted } = newOnlyRows.length > 0 ? upsertMergedProjects(batchId, newOnlyRows) : { inserted: 0 };
    if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);

    res.status(201).json({
      ok: true,
      report: {
        slotLabel: `表格增量导入（识别为：${detectedType}）`,
        summary: `共读取 ${rows.length} 条记录：成功新增 ${inserted} 条新项目，已自动跳过 ${skippedDuplicates} 条重复项目`,
        warnings,
      },
    });
  } catch (err) {
    if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    res.status(500).json({ error: `增量导入失败：${(err as Error).message}` });
  }
});

// POST /api/bidding/projects/batch-delete - Batch delete selected projects
biddingRouter.post('/projects/batch-delete', (req, res) => {
  const { ids } = req.body as { ids?: string[] };
  if (!Array.isArray(ids) || ids.length === 0) {
    return res.status(400).json({ error: '请先勾选要删除的项目' });
  }
  const placeholders = ids.map(() => '?').join(',');
  db.prepare(`DELETE FROM batch_projects WHERE id IN (${placeholders})`).run(...ids);
  res.json({ ok: true, deleted: ids.length });
});

// POST /api/bidding/projects/:projectId/fix-weight - Apply one-click weight ratio fix
biddingRouter.post('/projects/:projectId/fix-weight', (req, res) => {
  const { weight_ratio } = req.body as { weight_ratio?: string };
  if (!weight_ratio) return res.status(400).json({ error: '缺少修正后的权重比例' });
  db.prepare(`
    UPDATE batch_projects SET weight_ratio = ?, weight_warning = '', weight_suggested = '' WHERE id = ?
  `).run(weight_ratio.trim(), req.params.projectId);
  res.json({ ok: true });
});

// ==================== 变更公告管理 (PDF/附件 + 涉及标包关联提醒) ====================

// POST /api/bidding/batch/:batchId/change-notices
biddingRouter.post('/batch/:batchId/change-notices', upload.single('file'), (req, res) => {
  const { batchId } = req.params;
  const batch = db.prepare('SELECT id FROM batches WHERE id = ?').get(batchId);
  if (!batch) {
    if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    return res.status(404).json({ error: '未找到该批次' });
  }

  const title = String(req.body?.title || '').trim();
  const noticeDate = String(req.body?.notice_date || '').trim();
  const summary = String(req.body?.summary || '').trim();
  let changeTypes: string[] = [];
  let projectIds: string[] = [];
  try { changeTypes = JSON.parse(String(req.body?.change_types || '[]')); } catch { changeTypes = []; }
  try { projectIds = JSON.parse(String(req.body?.project_ids || '[]')); } catch { projectIds = []; }

  const fileOriginalName = req.file ? decodeName(req.file.originalname) : '';
  const fileStoredPath = req.file ? path.relative(UPLOADS_DIR, req.file.path).replace(/\\/g, '/') : '';
  const fileMimeType = req.file ? (req.file.mimetype || 'application/pdf') : '';

  const effectiveTitle = title || fileOriginalName.replace(/\.[^.]+$/, '') || '变更公告';
  if (!effectiveTitle && !summary && !req.file) {
    return res.status(400).json({ error: '请至少填写变更说明或上传变更公告文件' });
  }

  const id = uuid();
  const now = new Date().toISOString();
  db.prepare(`
    INSERT INTO batch_change_notices (
      id, batch_id, title, notice_date, change_types_json, summary, project_ids_json,
      file_original_name, file_stored_path, file_mime_type, created_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).run(
    id,
    batchId,
    effectiveTitle,
    noticeDate,
    JSON.stringify(Array.isArray(changeTypes) ? changeTypes : []),
    summary,
    JSON.stringify(Array.isArray(projectIds) ? projectIds : []),
    fileOriginalName,
    fileStoredPath,
    fileMimeType,
    now
  );

  res.status(201).json({
    ok: true,
    notice: {
      id,
      batch_id: batchId,
      title: effectiveTitle,
      notice_date: noticeDate,
      change_types: Array.isArray(changeTypes) ? changeTypes : [],
      summary,
      project_ids: Array.isArray(projectIds) ? projectIds : [],
      file_original_name: fileOriginalName,
      file_mime_type: fileMimeType,
      has_file: Boolean(fileStoredPath),
      created_at: now,
    },
  });
});

// PUT /api/bidding/change-notices/:noticeId — edit existing change notice & optionally replace/remove attachment
biddingRouter.put('/change-notices/:noticeId', upload.single('file'), (req, res) => {
  const existing = db
    .prepare('SELECT * FROM batch_change_notices WHERE id = ?')
    .get(req.params.noticeId) as Record<string, unknown> | undefined;
  if (!existing) {
    if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    return res.status(404).json({ error: '未找到该变更公告' });
  }

  const title = req.body?.title !== undefined ? String(req.body.title).trim() : String(existing.title || '');
  const noticeDate =
    req.body?.notice_date !== undefined ? String(req.body.notice_date).trim() : String(existing.notice_date || '');
  const summary = req.body?.summary !== undefined ? String(req.body.summary).trim() : String(existing.summary || '');
  const removeFile = String(req.body?.remove_file || '') === 'true';

  let changeTypes: string[] = [];
  let projectIds: string[] = [];
  try {
    changeTypes =
      req.body?.change_types !== undefined
        ? JSON.parse(String(req.body.change_types))
        : JSON.parse(String(existing.change_types_json || '[]'));
  } catch {
    changeTypes = [];
  }
  try {
    projectIds =
      req.body?.project_ids !== undefined
        ? JSON.parse(String(req.body.project_ids))
        : JSON.parse(String(existing.project_ids_json || '[]'));
  } catch {
    projectIds = [];
  }

  let fileOriginalName = String(existing.file_original_name || '');
  let fileStoredPath = String(existing.file_stored_path || '');
  let fileMimeType = String(existing.file_mime_type || '');

  if (req.file) {
    // Delete old file if present
    if (fileStoredPath) {
      const oldAbs = path.join(UPLOADS_DIR, fileStoredPath);
      if (fs.existsSync(oldAbs)) {
        try { fs.unlinkSync(oldAbs); } catch { /* ignore */ }
      }
    }
    fileOriginalName = decodeName(req.file.originalname);
    fileStoredPath = path.relative(UPLOADS_DIR, req.file.path).replace(/\\/g, '/');
    fileMimeType = req.file.mimetype || 'application/pdf';
  } else if (removeFile && fileStoredPath) {
    const oldAbs = path.join(UPLOADS_DIR, fileStoredPath);
    if (fs.existsSync(oldAbs)) {
      try { fs.unlinkSync(oldAbs); } catch { /* ignore */ }
    }
    fileOriginalName = '';
    fileStoredPath = '';
    fileMimeType = '';
  }

  const effectiveTitle = title || fileOriginalName.replace(/\.[^.]+$/, '') || '变更公告';

  db.prepare(`
    UPDATE batch_change_notices
    SET title = ?, notice_date = ?, change_types_json = ?, summary = ?, project_ids_json = ?,
        file_original_name = ?, file_stored_path = ?, file_mime_type = ?
    WHERE id = ?
  `).run(
    effectiveTitle,
    noticeDate,
    JSON.stringify(Array.isArray(changeTypes) ? changeTypes : []),
    summary,
    JSON.stringify(Array.isArray(projectIds) ? projectIds : []),
    fileOriginalName,
    fileStoredPath,
    fileMimeType,
    req.params.noticeId
  );

  res.json({
    ok: true,
    notice: {
      id: String(existing.id),
      batch_id: String(existing.batch_id),
      title: effectiveTitle,
      notice_date: noticeDate,
      change_types: Array.isArray(changeTypes) ? changeTypes : [],
      summary,
      project_ids: Array.isArray(projectIds) ? projectIds : [],
      file_original_name: fileOriginalName,
      file_mime_type: fileMimeType,
      has_file: Boolean(fileStoredPath),
      created_at: String(existing.created_at || ''),
    },
  });
});

// DELETE /api/bidding/change-notices/:noticeId
biddingRouter.delete('/change-notices/:noticeId', (req, res) => {
  const row = db
    .prepare('SELECT id, file_stored_path FROM batch_change_notices WHERE id = ?')
    .get(req.params.noticeId) as { id: string; file_stored_path?: string } | undefined;
  if (!row) return res.status(404).json({ error: '未找到该变更公告' });
  if (row.file_stored_path) {
    const absPath = path.join(UPLOADS_DIR, row.file_stored_path);
    if (fs.existsSync(absPath)) {
      try { fs.unlinkSync(absPath); } catch { /* ignore */ }
    }
  }
  db.prepare('DELETE FROM batch_change_notices WHERE id = ?').run(req.params.noticeId);
  res.json({ ok: true });
});

// GET /api/bidding/change-notices/:noticeId/file — inline view PDF or download attachment
biddingRouter.get('/change-notices/:noticeId/file', (req, res) => {
  const row = db
    .prepare('SELECT file_original_name, file_stored_path, file_mime_type FROM batch_change_notices WHERE id = ?')
    .get(req.params.noticeId) as { file_original_name?: string; file_stored_path?: string; file_mime_type?: string } | undefined;
  if (!row || !row.file_stored_path) return res.status(404).json({ error: '该变更公告未附带文件' });
  const absPath = path.join(UPLOADS_DIR, row.file_stored_path);
  if (!fs.existsSync(absPath)) return res.status(404).json({ error: '附件文件不存在' });
  const filename = row.file_original_name || path.basename(absPath);
  const isInline = /pdf|image/i.test(row.file_mime_type || '') || /\.(pdf|png|jpe?g|webp)$/i.test(filename);
  res.setHeader('Content-Type', row.file_mime_type || 'application/octet-stream');
  res.setHeader(
    'Content-Disposition',
    `${isInline ? 'inline' : 'attachment'}; filename*=UTF-8''${encodeURIComponent(filename)}`
  );
  res.sendFile(absPath);
});
