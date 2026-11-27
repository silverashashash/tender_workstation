import fs from 'node:fs';
import * as XLSXNamespace from 'xlsx';
import { evaluateWorkbookFormulas, getCellFormulaError } from './excelFormulaEval.js';

const XLSX = (XLSXNamespace as unknown as { default?: typeof XLSXNamespace }).default ?? XLSXNamespace;

/**
 * Strip ALL whitespace (half-width spaces, full-width spaces, tabs, newlines)
 * — used for supplier names, contacts, phones, emails, tax numbers, bank accounts.
 */
export function stripAllSpaces(val: unknown): string {
  if (typeof val === 'object' && val !== null && '__cellError' in val) {
    throw new Error(String((val as { __cellError: string }).__cellError));
  }
  return String(val ?? '').replace(/[\s\u3000\u00A0]+/g, '');
}

/**
 * Compare two primary keys (采购申请号 / 采购申请标识 / 采购申请唯一标识 / 采购申请).
 * Supports exact match as well as SAP 10-digit PR <-> 15-digit 采购申请唯一标识 (10-digit + 5-digit line item).
 */
export function isSameReqAppNo(a: unknown, b: unknown): boolean {
  const sa = stripAllSpaces(a);
  const sb = stripAllSpaces(b);
  if (!sa || !sb) return false;
  if (sa === sb) return true;
  if (sa.length === 15 && sb.length === 10 && sa.startsWith(sb)) return true;
  if (sb.length === 15 && sa.length === 10 && sb.startsWith(sa)) return true;
  return false;
}

/**
 * Clean general text by trimming and collapsing internal newlines/extra spaces.
 */
export function cleanText(val: unknown): string {
  if (typeof val === 'object' && val !== null && '__cellError' in val) {
    throw new Error(String((val as { __cellError: string }).__cellError));
  }
  return String(val ?? '').replace(/\r?\n+/g, ' ').replace(/\s+/g, ' ').trim();
}

/**
 * Normalize monetary price values to exact Yuan (元), supporting both 元 and 万元 inputs.
 */
export function parsePriceToYuan(rawVal: unknown, headerHint = ''): number | null {
  if (rawVal === null || rawVal === undefined || rawVal === '') return null;
  if (typeof rawVal === 'number' && Number.isFinite(rawVal)) {
    const isWanHeader = /万元|万/.test(headerHint);
    return Math.round((isWanHeader ? rawVal * 10000 : rawVal) * 100) / 100;
  }
  const str = stripAllSpaces(rawVal).replace(/,/g, '');
  const match = str.match(/^([+-]?\d+(?:\.\d+)?)(万元|万|元)?$/);
  if (!match) return null;
  const num = Number(match[1]);
  if (!Number.isFinite(num)) return null;
  const unit = match[2] || '';
  const isWan = unit.includes('万') || (!unit && /万元|万/.test(headerHint));
  return Math.round((isWan ? num * 10000 : num) * 100) / 100;
}

export function formatYuanAsWan(yuan: number | null | undefined): string {
  if (yuan === null || yuan === undefined || !Number.isFinite(yuan)) return '';
  const wan = yuan / 10000;
  return `${Number(wan.toFixed(4))} 万元`;
}

/**
 * Parse max price supporting BOTH monetary amounts (总价报价 -> 元)
 * AND percentage/discount rates (折扣率/下浮率/费率报价 -> %).
 */
export function parseMaxPrice(
  rawVal: unknown,
  headerHint = '',
  quoteMethod = ''
): { yuan: number | null; rawDisplay: string } {
  if (rawVal === null || rawVal === undefined || String(rawVal).trim() === '') {
    return { yuan: null, rawDisplay: '' };
  }
  const str = stripAllSpaces(rawVal);
  const isDiscountMode = /折扣|下浮|费率|百分/.test(quoteMethod);

  // Explicit percentage string e.g. "95%" or "95.5%"
  if (/%$/.test(str)) {
    return { yuan: null, rawDisplay: str };
  }
  // Numeric decimal stored by Excel for percentage cell (e.g. 0.95 -> "95%") when quote method is discount rate
  if (isDiscountMode) {
    const num = Number(str);
    if (Number.isFinite(num)) {
      if (num > 0 && num <= 1) {
        return { yuan: null, rawDisplay: `${Number((num * 100).toFixed(4))}%` };
      }
      if (num > 1 && num <= 100) {
        return { yuan: null, rawDisplay: `${Number(num.toFixed(4))}%` };
      }
    }
    return { yuan: null, rawDisplay: cleanText(rawVal) };
  }

  const yuan = parsePriceToYuan(rawVal, headerHint);
  return { yuan, rawDisplay: yuan !== null ? '' : cleanText(rawVal) };
}

/**
 * Parse and validate Technical:Business:Price weight ratio (技:商:价).
 * Handles Excel's notorious conversion of "60:10:30" into time-serial floats (e.g. 2.50729166...)
 * or truncated 24-hour modulo strings (e.g. "12:10:30" where 60 % 24 = 12).
 */
export function parseAndValidateWeightRatio(rawVal: unknown): {
  normalized: string;
  valid: boolean;
  warning: string;
  suggestedFix: string;
} {
  if (rawVal === null || rawVal === undefined || String(rawVal).trim() === '') {
    return { normalized: '', valid: true, warning: '', suggestedFix: '' };
  }

  let parts: number[] | null = null;
  let convertedFromDecimal = false;

  if (typeof rawVal === 'number' && Number.isFinite(rawVal) && rawVal > 0 && rawVal < 10) {
    const totalSec = Math.round(rawVal * 86400);
    const h = Math.floor(totalSec / 3600);
    const m = Math.floor((totalSec % 3600) / 60);
    const s = totalSec % 60;
    parts = [h, m, s];
    convertedFromDecimal = true;
  } else {
    const text = stripAllSpaces(rawVal).replace(/：/g, ':');
    if (/^\d+\.\d+$/.test(text)) {
      const num = Number(text);
      if (num > 0 && num < 10) {
        const totalSec = Math.round(num * 86400);
        const h = Math.floor(totalSec / 3600);
        const m = Math.floor((totalSec % 3600) / 60);
        const s = totalSec % 60;
        parts = [h, m, s];
        convertedFromDecimal = true;
      }
    } else {
      const m = text.match(/^(\d+):(\d+):(\d+)$/);
      if (m) {
        parts = [Number(m[1]), Number(m[2]), Number(m[3])];
      }
    }
  }

  if (!parts) {
    return {
      normalized: cleanText(rawVal),
      valid: false,
      warning: `权重格式无法识别（原始值：${cleanText(rawVal)}），请核对为“技:商:价”三段整数`,
      suggestedFix: '',
    };
  }

  const [a, b, c] = parts;
  const normalized = `${a}:${b}:${c}`;
  const sum = a + b + c;

  if (sum === 100) {
    return {
      normalized,
      valid: true,
      warning: convertedFromDecimal ? '已自动从 Excel 小数时间格式还原为权重比例' : '',
      suggestedFix: '',
    };
  }

  const expectedA = 100 - b - c;
  let suggestedFix = '';
  let warning = `三项权重之和为 ${sum}（应为 100）`;
  if (expectedA > 0 && expectedA < 100 && expectedA % 24 === a % 24) {
    suggestedFix = `${expectedA}:${b}:${c}`;
    warning = `疑似 Excel 时间格式截断（${normalized} 之和=${sum}），建议还原为 ${suggestedFix}`;
  }

  return { normalized, valid: false, warning, suggestedFix };
}

/**
 * Split declarant string like "张三13800000001" or "李四\n13900000002" into { name, phone }.
 */
export function splitDeclarant(raw: unknown): { name: string; phone: string } {
  const compact = stripAllSpaces(raw);
  if (!compact) return { name: '', phone: '' };
  const phoneMatch = compact.match(/(1\d{10}|\d{3,4}-?\d{7,8})/);
  if (!phoneMatch) return { name: compact, phone: '' };
  const phone = phoneMatch[1];
  const name = compact.replace(phone, '').replace(/[()（）/:：,-]/g, '').trim();
  return { name, phone };
}

interface CellErrorSentinel {
  __cellError: string;
}

function assertCellVal(v: unknown): unknown {
  if (typeof v === 'object' && v !== null && '__cellError' in v) {
    throw new Error((v as CellErrorSentinel).__cellError);
  }
  return v;
}

function readSheetRawAoa(filePath: string, raw = true): { aoa: unknown[][]; mergesExpanded: number } {
  const buf = fs.readFileSync(filePath);
  const wb = XLSX.read(buf, { type: 'buffer', raw: true, sheetStubs: true, cellNF: true, cellStyles: true });
  evaluateWorkbookFormulas(wb);
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: '', raw });
  // Overlay error sentinels so read key cells fail loudly instead of silently becoming empty strings
  if (ws && ws['!ref']) {
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let r = range.s.r; r <= range.e.r; r++) {
      const rowArr = aoa[r - range.s.r];
      if (!rowArr) continue;
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr] as XLSXNamespace.CellObject | undefined;
        if (cell && cell.t === 'e') {
          const msg = getCellFormulaError(cell) || `${sheetName}!${addr} 为错误单元格（${cell.w || cell.v || '#ERROR'}）`;
          rowArr[c - range.s.c] = { __cellError: msg } as CellErrorSentinel;
        }
      }
    }
  }
  const merges = (ws['!merges'] || []) as Array<{ s: { r: number; c: number }; e: { r: number; c: number } }>;

  // Automatically expand all merged cell ranges (e.g. 合并成包 / 子包 rows sharing lot_code, package_no, qualifications, etc.)
  for (const m of merges) {
    const topVal = aoa[m.s.r]?.[m.s.c];
    if (topVal === undefined || topVal === null || String(topVal).trim() === '') continue;
    for (let r = m.s.r; r <= m.e.r; r++) {
      if (!aoa[r]) continue;
      for (let c = m.s.c; c <= m.e.c; c++) {
        const cur = aoa[r][c];
        if (cur === '' || cur === undefined || cur === null) {
          aoa[r][c] = topVal;
        }
      }
    }
  }

  return { aoa, mergesExpanded: merges.length };
}

function findHeaderRowIndex(aoa: unknown[][], keywords: RegExp[]): number {
  for (let i = 0; i < Math.min(aoa.length, 6); i++) {
    const row = aoa[i] || [];
    const hits = row.filter((c) => keywords.some((k) => k.test(String(c ?? '')))).length;
    if (hits >= 2) return i;
  }
  return 0;
}

function colIndex(headers: string[], patterns: RegExp[]): number {
  return headers.findIndex((h) => patterns.some((p) => p.test(h)));
}

// 1. Parse 采购需求变量模板.xlsx
export function parseVariableTemplate(filePath: string) {
  const { aoa } = readSheetRawAoa(filePath);
  const hIdx = findHeaderRowIndex(aoa, [/项目名称|文件名称/, /采购编号|招标编号/, /截止时间|公告日期/]);
  const headers = (aoa[hIdx] || []).map((v) => cleanText(v));
  const dataRow = aoa[hIdx + 1] || [];

  const get = (patterns: RegExp[]) => {
    const idx = colIndex(headers, patterns);
    return idx >= 0 ? cleanText(dataRow[idx]) : '';
  };

  const name = get([/项目名称/, /文件名称/]);
  const code = stripAllSpaces(get([/采购编号|招标编号|批次编号/]));
  const docFetchTime = get([/获取.*文件时间/]);
  const bidOpeningTime = get([/应答截止时间|投标截止时间|开标时间/]);
  const announcementTime = get([/公告日期|挂网时间/]);
  const contactPhone = stripAllSpaces(get([/联系人电话|联系电话/]));
  const contactEmail = stripAllSpaces(get([/电子邮件|邮箱/]));

  const warnings: string[] = [];
  for (const [label, val] of [
    ['获取文件时间', docFetchTime],
    ['截止/开标时间', bidOpeningTime],
    ['公告日期', announcementTime],
  ]) {
    if (val && /(^|[^2])02\d年/.test(val)) {
      warnings.push(`${label}中检测到疑似年份缺位：“${val}”，请核对是否漏写“2”（如 2026年）`);
    }
  }

  let procurementType: '公开招标' | '非公开招标' | undefined;
  if (/公开招标/.test(name) && !/竞争性谈判|非公开/.test(name)) {
    procurementType = '公开招标';
  } else if (/竞争性谈判|非公开|询价|单一来源/.test(name)) {
    procurementType = '非公开招标';
  }

  return {
    batchPatch: {
      name,
      code,
      announcement_time: announcementTime,
      doc_fetch_time: docFetchTime,
      bid_opening_time: bidOpeningTime,
      contact_phone: contactPhone,
      contact_email: contactEmail,
      ...(procurementType ? { procurement_type: procurementType } : {}),
    },
    warnings,
  };
}

// 2. Parse 附件7：批次申报计划表.xlsx
export function parseDeclarationPlan(filePath: string) {
  const { aoa } = readSheetRawAoa(filePath);
  const hIdx = findHeaderRowIndex(aoa, [/项目名称/, /项目编码|技术规范书编码|技术规范书ID/i, /申报人/]);
  const headers = (aoa[hIdx] || []).map((v) => cleanText(v));

  const idxDept = colIndex(headers, [/项目管理部门/]);
  const idxUnit = colIndex(headers, [/项目单位/]);
  const idxName = colIndex(headers, [/项目名称/]);
  const idxReqAppNo = colIndex(headers, [/采购申请号/]);
  const idxOverview = colIndex(headers, [/项目概况/]);
  const idxBudget = colIndex(headers, [/含税预算金额|预算金额|概算金额|含税概算|概算总价/]);
  const idxMaxPrice = colIndex(headers, [/最高限价/]);
  const idxDuration = colIndex(headers, [/工期|工程进度安排|计划开.*竣工/]);
  const idxTechSpec = colIndex(headers, [/技术规范书编码|技术规范书ID/i]);
  const idxDeclarant = colIndex(headers, [/申报人及联系方式|申报人/]);
  const idxQual = colIndex(headers, [/资质条件|资质要求/]);
  const idxPerf = colIndex(headers, [/业绩要求/]);
  const idxStaff = colIndex(headers, [/人员要求/]);
  const idxQuote = colIndex(headers, [/报价方式/]);
  const idxSubcontract = colIndex(headers, [/分包原则/]);
  const idxContract = colIndex(headers, [/合同模板名称|合同版本|合同名称/]);
  const idxFunding = colIndex(headers, [/资金来源/]);
  const idxProjCode = colIndex(headers, [/项目编码/]);

  const budgetHeader = idxBudget >= 0 ? headers[idxBudget] : '';
  const maxPriceHeader = idxMaxPrice >= 0 ? headers[idxMaxPrice] : '';

  const rows: Array<Record<string, unknown>> = [];
  for (let i = hIdx + 1; i < aoa.length; i++) {
    const r = aoa[i] || [];
    const projName = idxName >= 0 ? cleanText(r[idxName]) : '';
    const firstCell = cleanText(r[0]);
    if (!projName || /^分管领导|^注[:：]/.test(firstCell)) continue;

    const quoteMethod = idxQuote >= 0 ? cleanText(r[idxQuote]) : '';
    const declarant = idxDeclarant >= 0 ? splitDeclarant(r[idxDeclarant]) : { name: '', phone: '' };
    const budgetYuan = idxBudget >= 0 ? parsePriceToYuan(r[idxBudget], budgetHeader) : null;
    const maxP = idxMaxPrice >= 0
      ? parseMaxPrice(r[idxMaxPrice], maxPriceHeader, quoteMethod)
      : { yuan: budgetYuan, rawDisplay: '' };

    rows.push({
      department: idxDept >= 0 ? cleanText(r[idxDept]) : '',
      project_unit: idxUnit >= 0 ? cleanText(r[idxUnit]) : '',
      name: projName,
      req_app_no: idxReqAppNo >= 0 ? stripAllSpaces(r[idxReqAppNo]) : '',
      project_overview: idxOverview >= 0 ? cleanText(r[idxOverview]) : '',
      project_code: idxProjCode >= 0 ? stripAllSpaces(r[idxProjCode]) : '',
      tech_spec_code: idxTechSpec >= 0 ? stripAllSpaces(r[idxTechSpec]) : '',
      declarant_name: declarant.name,
      declarant_phone: declarant.phone,
      qualification_req: idxQual >= 0 ? cleanText(r[idxQual]) : '',
      performance_req: idxPerf >= 0 ? cleanText(r[idxPerf]) : '',
      staff_req: idxStaff >= 0 ? cleanText(r[idxStaff]) : '',
      quote_method: quoteMethod,
      subcontract_rule: idxSubcontract >= 0 ? cleanText(r[idxSubcontract]) : '',
      budget_yuan: budgetYuan,
      max_price_yuan: maxP.yuan,
      max_price_raw: maxP.rawDisplay,
      duration: idxDuration >= 0 ? cleanText(r[idxDuration]) : '',
      contract_version: idxContract >= 0 ? cleanText(r[idxContract]) : '',
      funding_source: idxFunding >= 0 ? cleanText(r[idxFunding]) : '',
    });
  }

  return { rows };
}

// 3. Parse 采购计划-ERP导出 (compatible with BOTH 主业 SAP/ECP ERP & 产业 ETP/ECP ERP exports)
export function parseErpPlan(filePath: string) {
  const { aoa } = readSheetRawAoa(filePath);
  const hIdx = findHeaderRowIndex(aoa, [
    /批次编码|招标计划编号|批次名称|招标计划名称/,
    /工程项目名称|^项目名称$|规范项目名称|工程项目编码|^项目编码$/,
    /采购申请唯一标识|采购申请标识|^采购申请号$|^采购申请$|技术规范书ID|技术规范书编码|物料编码/i,
  ]);
  const headers = (aoa[hIdx] || []).map((v) => cleanText(v));

  const isMainGridErp = headers.some((h) =>
    /^(招标计划编号|采购申请唯一标识|网省物资审批标识|工厂简称|概算总价)$/.test(h)
  );
  const erpSource: '主业' | '产业' = isMainGridErp ? '主业' : '产业';

  const idxBatchCode = colIndex(headers, [/^批次编码$/, /^招标计划编号$/]);
  const idxBatchName = colIndex(headers, [/^批次名称$/, /^招标计划名称$/]);
  const idxExecSys = colIndex(headers, [/批次执行系统/]);
  const idxUnit = colIndex(headers, [/^需求单位名称$/, /^工厂简称$/, /^工厂名称$/, /^实施主体$/, /^计划提报公司$/]);
  const idxDept = colIndex(headers, [/^主管部门$/, /^计划提报部门$/, /^创建部门$/]);
  const idxProjCode = colIndex(headers, [/^工程项目编码$/, /^项目编码$/]);
  const idxProjName = colIndex(headers, [/^工程项目名称$/, /^项目名称$/, /^规范项目名称$/]);
  const idxAltProjName = colIndex(headers, [/^规范项目名称$/, /^大工程名称$/, /^年度计划单体工程名称$/]);
  const idxOverview = colIndex(headers, [/^项目概况$/, /^工程规模$/]);
  const idxDuration = colIndex(headers, [/^工期$/]);
  const idxPrice = colIndex(headers, [/^评估总价$/, /^概算总价$/, /^评估单价$/, /^概算价格$/]);
  const idxContact = colIndex(headers, [/^联系人$/, /^计划提报人$/, /^创建人$/]);
  const idxPhone = colIndex(headers, [/^联系信息$/, /^联系电话$/]);
  const idxTechSpec = colIndex(headers, [/技术规范书ID|技术规范书编码/i]);
  const idxMethod = colIndex(headers, [/^批次采购方式$/, /^采购方式$/]);
  const idxSyncStatus = colIndex(headers, [/采购申请同步状态/, /网省退回原因/, /ECP返回消息/]);
  const idxReqUid = colIndex(headers, [/^采购申请唯一标识$/, /^采购申请标识$/, /^采购申请号$/]);
  const idxReqPr = colIndex(headers, [/^采购申请$/]);
  const idxFundingName = colIndex(headers, [/^资金来源名称$/]);
  const idxFunding = colIndex(headers, [/^资金来源$/]);
  const idxMaterialCode = colIndex(headers, [/^物料编码$/]);
  const idxMaterialName = colIndex(headers, [/^物料名称$|^物料描述$/]);

  const priceHeader = idxPrice >= 0 ? headers[idxPrice] : '元';

  let batchMeta: Record<string, string> = {};
  const warnings: string[] = [];
  const rows: Array<Record<string, unknown>> = [];

  const cleanOptionalContact = (raw: unknown) => {
    const s = stripAllSpaces(raw);
    return !s || s === '无' || s === '-' || s === '/' ? '' : s;
  };

  for (let i = hIdx + 1; i < aoa.length; i++) {
    const r = aoa[i] || [];
    const projName =
      (idxProjName >= 0 ? cleanText(r[idxProjName]) : '') ||
      (idxAltProjName >= 0 ? cleanText(r[idxAltProjName]) : '');
    const reqUid = idxReqUid >= 0 ? stripAllSpaces(r[idxReqUid]) : '';
    const reqPr = idxReqPr >= 0 ? stripAllSpaces(r[idxReqPr]) : '';
    const reqAppNo = reqUid || reqPr;
    if (!projName && !reqAppNo) continue;

    const rawMethod = idxMethod >= 0 ? cleanText(r[idxMethod]) : '';
    const bName = idxBatchName >= 0 ? cleanText(r[idxBatchName]) : '';
    const methodText = rawMethod || (/公开招标/.test(bName) ? '公开招标' : /竞争性谈判|公开谈判/.test(bName) ? '竞争性谈判' : '');

    if (!batchMeta.code && idxBatchCode >= 0) {
      const execSys = idxExecSys >= 0 ? stripAllSpaces(r[idxExecSys]).toUpperCase() : '';
      batchMeta = {
        code: stripAllSpaces(r[idxBatchCode]),
        name: bName,
        project_unit: idxUnit >= 0 ? cleanText(r[idxUnit]) : '',
        unit_nature: isMainGridErp || execSys.includes('ECP') ? '主业' : execSys.includes('ETP') ? '产业' : '产业',
        procurement_type: /公开招标/.test(methodText || bName) ? '公开招标' : '非公开招标',
      };
    }

    const syncStatus = idxSyncStatus >= 0 ? cleanText(r[idxSyncStatus]) : '';
    if (/退回/.test(syncStatus)) {
      warnings.push(`项目「${projName || reqAppNo}」在 ERP 中状态为“${syncStatus}”，请确认是否已重新同步`);
    }

    const fundingName = idxFundingName >= 0 ? cleanText(r[idxFundingName]) : '';
    const fundingRaw = idxFunding >= 0 ? cleanText(r[idxFunding]) : '';
    const fundingSource = fundingName || (fundingRaw && fundingRaw !== '0' ? fundingRaw : '');

    rows.push({
      name: projName,
      project_unit: idxUnit >= 0 ? cleanText(r[idxUnit]) : '',
      department: idxDept >= 0 ? cleanText(r[idxDept]) : '',
      project_overview: idxOverview >= 0 ? cleanText(r[idxOverview]) : '',
      duration: idxDuration >= 0 ? cleanText(r[idxDuration]) : '',
      project_code: idxProjCode >= 0 ? stripAllSpaces(r[idxProjCode]) : '',
      req_app_no: reqAppNo,
      tech_spec_code: idxTechSpec >= 0 ? stripAllSpaces(r[idxTechSpec]) : '',
      budget_yuan: idxPrice >= 0 ? parsePriceToYuan(r[idxPrice], priceHeader) : null,
      declarant_name: idxContact >= 0 ? cleanOptionalContact(r[idxContact]) : '',
      declarant_phone: idxPhone >= 0 ? cleanOptionalContact(r[idxPhone]) : '',
      procurement_method: methodText || '竞争性谈判',
      funding_source: fundingSource,
      erp_sync_status: syncStatus,
      material_code: idxMaterialCode >= 0 ? stripAllSpaces(r[idxMaterialCode]) : '',
      material_desc: idxMaterialName >= 0 ? cleanText(r[idxMaterialName]) : '',
    });
  }

  return { batchMeta, erpSource, rows, warnings };
}

// 4. Parse 需求一览表.xlsx
export function parseRequirementList(filePath: string) {
  const { aoa, mergesExpanded } = readSheetRawAoa(filePath);
  const hIdx = findHeaderRowIndex(aoa, [/采购申请号/, /分标编号/, /项目名称/]);
  const headers = (aoa[hIdx] || []).map((v) => cleanText(v));

  const idxReqAppNo = colIndex(headers, [/采购申请号/]);
  const idxLotCode = colIndex(headers, [/分标编号/]);
  const idxLotName = colIndex(headers, [/分标名称/]);
  const idxPkg = colIndex(headers, [/包号/]);
  const idxUnit = colIndex(headers, [/项目单位/]);
  const idxName = colIndex(headers, [/项目名称/]);
  const idxOverview = colIndex(headers, [/项目概况/]);
  const idxDuration = colIndex(headers, [/工期|工程进度安排|计划开.*竣工/]);
  const idxQual = colIndex(headers, [/资质条件|资质要求/]);
  const idxPerf = colIndex(headers, [/业绩要求/]);
  const idxStaff = colIndex(headers, [/人员要求/]);
  const idxBudget = colIndex(headers, [/含税预算金额|预算金额|概算金额|含税概算|概算总价/]);
  const idxMaxPrice = colIndex(headers, [/最高限价/]);
  const idxQuote = colIndex(headers, [/报价方式/]);
  const idxSubcontract = colIndex(headers, [/分包原则/]);
  const idxTechSpec = colIndex(headers, [/技术规范书编码|技术规范书ID/i]);
  const idxWeight = colIndex(headers, [/权重比例/]);
  const idxFormula = colIndex(headers, [/价格公式/]);
  const idxParams = colIndex(headers, [/参数选择/]);
  const idxBizTpl = colIndex(headers, [/商务详评模板|商务细则/]);
  const idxTechTpl = colIndex(headers, [/技术详评模板|技术细则/]);
  const idxContract = colIndex(headers, [/合同版本|合同名称|合同模板/]);
  const idxMaterialCode = colIndex(headers, [/物料编码/]);
  const idxMaterialDesc = colIndex(headers, [/物料描述|物料名称/]);

  const budgetHeader = idxBudget >= 0 ? headers[idxBudget] : '元';
  const maxPriceHeader = idxMaxPrice >= 0 ? headers[idxMaxPrice] : '元';

  const warnings: string[] = [];
  const rows: Array<Record<string, unknown>> = [];

  for (let i = hIdx + 1; i < aoa.length; i++) {
    const r = aoa[i] || [];
    const projName = idxName >= 0 ? cleanText(r[idxName]) : '';
    if (!projName) continue;

    const quoteMethod = idxQuote >= 0 ? cleanText(r[idxQuote]) : '';
    const maxP = idxMaxPrice >= 0
      ? parseMaxPrice(r[idxMaxPrice], maxPriceHeader, quoteMethod)
      : { yuan: null, rawDisplay: '' };

    const weightRes = idxWeight >= 0
      ? parseAndValidateWeightRatio(r[idxWeight])
      : { normalized: '', valid: true, warning: '', suggestedFix: '' };

    if (!weightRes.valid && weightRes.warning) {
      const lotInfo = `${idxLotCode >= 0 ? cleanText(r[idxLotCode]) : ''} ${idxPkg >= 0 ? cleanText(r[idxPkg]) : ''}`.trim();
      warnings.push(`[${lotInfo || projName}] ${weightRes.warning}`);
    }

    rows.push({
      req_app_no: idxReqAppNo >= 0 ? stripAllSpaces(r[idxReqAppNo]) : '',
      lot_code: idxLotCode >= 0 ? stripAllSpaces(r[idxLotCode]) : '',
      lot_name: idxLotName >= 0 ? cleanText(r[idxLotName]) : '',
      package_no: idxPkg >= 0 ? stripAllSpaces(r[idxPkg]) : '',
      project_unit: idxUnit >= 0 ? cleanText(r[idxUnit]) : '',
      name: projName,
      project_overview: idxOverview >= 0 ? cleanText(r[idxOverview]) : '',
      duration: idxDuration >= 0 ? cleanText(r[idxDuration]) : '',
      qualification_req: idxQual >= 0 ? cleanText(r[idxQual]) : '',
      performance_req: idxPerf >= 0 ? cleanText(r[idxPerf]) : '',
      staff_req: idxStaff >= 0 ? cleanText(r[idxStaff]) : '',
      budget_yuan: idxBudget >= 0 ? parsePriceToYuan(r[idxBudget], budgetHeader) : null,
      max_price_yuan: maxP.yuan,
      max_price_raw: maxP.rawDisplay,
      quote_method: quoteMethod,
      subcontract_rule: idxSubcontract >= 0 ? cleanText(r[idxSubcontract]) : '',
      tech_spec_code: idxTechSpec >= 0 ? stripAllSpaces(r[idxTechSpec]) : '',
      weight_ratio: weightRes.normalized,
      weight_warning: weightRes.valid ? '' : weightRes.warning,
      weight_suggested: weightRes.suggestedFix,
      price_formula: idxFormula >= 0 ? cleanText(r[idxFormula]) : '',
      price_params: idxParams >= 0 ? cleanText(r[idxParams]) : '',
      biz_eval_tpl: idxBizTpl >= 0 ? stripAllSpaces(r[idxBizTpl]) : '',
      tech_eval_tpl: idxTechTpl >= 0 ? stripAllSpaces(r[idxTechTpl]) : '',
      contract_version: idxContract >= 0 ? cleanText(r[idxContract]) : '',
      material_code: idxMaterialCode >= 0 ? stripAllSpaces(r[idxMaterialCode]) : '',
      material_desc: idxMaterialDesc >= 0 ? cleanText(r[idxMaterialDesc]) : '',
    });
  }

  return { rows, warnings, mergesExpanded };
}

// 5. Parse 供应商联系人信息.xlsx (compatible with BOTH ETP & ECP 2.0 headers)
export interface ParsedSupplier {
  seq: number;
  supplier_name: string;
  package_no: string;
  contact_name: string;
  mobile: string;
  phone: string;
  email: string;
  fax: string;
  address: string;
  postal_code: string;
  tax_no: string;
  invoice_title: string;
  bank_name: string;
  bank_account: string;
}

export interface CollusionRisk {
  type: 'shared_mobile' | 'shared_email';
  label: string;
  value: string;
  suppliers: string[];
}

export function parseSupplierContacts(filePath: string) {
  // Use formatted display values (raw: false) so numeric codes with leading zeros (e.g. fax 012345 with format 000000) stay intact
  const { aoa } = readSheetRawAoa(filePath, false);
  const hIdx = findHeaderRowIndex(aoa, [/投标人名称|供应商名称/, /联系人/, /手机号码|电话/]);
  const headers = (aoa[hIdx] || []).map((v) => cleanText(v));

  const idxName = colIndex(headers, [/投标人名称|响应供应商|供应商名称/]);
  const idxPkg = colIndex(headers, [/包号|分标|标段/]);
  const idxContact = colIndex(headers, [/联系人姓名|联系人/]);
  const idxMobile = colIndex(headers, [/手机号码|手机号/]);
  const idxAddr = colIndex(headers, [/邮政地址|地址/]);
  const idxZip = colIndex(headers, [/邮编/]);
  const idxPhone = colIndex(headers, [/固定电话|固话/]);
  const idxEmail = colIndex(headers, [/电子邮件|邮箱/]);
  const idxFax = colIndex(headers, [/传真号码|传真/]);
  const idxTax = colIndex(headers, [/税号/]);
  const idxInvoice = colIndex(headers, [/发票抬头/]);
  const idxBank = colIndex(headers, [/开户银行/]);
  const idxAccount = colIndex(headers, [/银行账号/]);

  const suppliers: ParsedSupplier[] = [];
  let spacesCleanedCount = 0;

  for (let i = hIdx + 1; i < aoa.length; i++) {
    const r = aoa[i] || [];
    const rawName = idxName >= 0 ? String(assertCellVal(r[idxName]) ?? '') : '';
    const cleanedName = stripAllSpaces(rawName);
    if (!cleanedName) continue;

    const cleanField = (idx: number) => {
      if (idx < 0) return '';
      const raw = String(assertCellVal(r[idx]) ?? '');
      const stripped = stripAllSpaces(raw);
      if (raw !== stripped) spacesCleanedCount++;
      return stripped;
    };

    if (rawName !== cleanedName) spacesCleanedCount++;

    const rawPkgCell = idxPkg >= 0 ? String(assertCellVal(r[idxPkg]) ?? '') : '';
    const cleanedPkg = stripAllSpaces(rawPkgCell.replace(/[\r\n]+/g, ';'));
    if (idxPkg >= 0 && rawPkgCell !== cleanedPkg) spacesCleanedCount++;

    suppliers.push({
      seq: suppliers.length + 1,
      supplier_name: cleanedName,
      package_no: cleanedPkg,
      contact_name: cleanField(idxContact),
      mobile: cleanField(idxMobile),
      phone: cleanField(idxPhone),
      email: cleanField(idxEmail).toLowerCase(),
      fax: cleanField(idxFax),
      address: cleanField(idxAddr),
      postal_code: cleanField(idxZip),
      tax_no: cleanField(idxTax).toUpperCase(),
      invoice_title: cleanField(idxInvoice),
      bank_name: cleanField(idxBank),
      bank_account: cleanField(idxAccount),
    });
  }

  return {
    suppliers,
    spacesCleanedCount,
    risks: detectCollusionRisks(suppliers),
  };
}

/**
 * Preliminary collusion screening:
 * Only flags DIFFERENT suppliers sharing the same mobile number or email address.
 * (Same supplier appearing multiple times is valid — e.g. bidding on multiple packages in ETP exports without package_no.)
 */
export function detectCollusionRisks(suppliers: Array<{
  supplier_name: string;
  mobile: string;
  email: string;
}>): CollusionRisk[] {
  const risks: CollusionRisk[] = [];

  const checkShared = (
    field: 'mobile' | 'email',
    type: CollusionRisk['type'],
    label: string
  ) => {
    const map = new Map<string, Set<string>>();
    for (const s of suppliers) {
      const val = (s[field] || '').trim();
      if (!val || val.length < 5) continue;
      if (!map.has(val)) map.set(val, new Set());
      map.get(val)!.add(s.supplier_name);
    }
    for (const [val, set] of map.entries()) {
      if (set.size > 1) {
        risks.push({ type, label, value: val, suppliers: Array.from(set) });
      }
    }
  };

  checkShared('mobile', 'shared_mobile', '不同供应商共用手机号');
  checkShared('email', 'shared_email', '不同供应商共用电子邮件');

  return risks;
}

/**
 * Auto-detect spreadsheet type (需求一览表 / ERP 采购计划 / 批次申报计划表)
 * for incremental project import with deduplication.
 */
export function parseAnyProjectSpreadsheet(filePath: string): {
  detectedType: string;
  rows: Array<Record<string, unknown>>;
  warnings: string[];
} {
  const { aoa } = readSheetRawAoa(filePath);
  const sampleText = aoa
    .slice(0, 6)
    .flat()
    .map((c) => cleanText(c))
    .join(' ');

  if (/分标编号|分标名称|权重比例|商务详评模板/.test(sampleText)) {
    const res = parseRequirementList(filePath);
    return { detectedType: '需求一览表', rows: res.rows, warnings: res.warnings };
  }
  if (/采购申请唯一标识|采购申请标识|招标计划编号|网省物资审批标识|批次执行系统|工程项目编码|评估总价|概算总价/.test(sampleText)) {
    const res = parseErpPlan(filePath);
    return { detectedType: `${res.erpSource} ERP 采购计划表`, rows: res.rows, warnings: res.warnings };
  }
  const res = parseDeclarationPlan(filePath);
  return { detectedType: '批次申报计划表 / 通用项目表', rows: res.rows, warnings: [] };
}
