import fs from 'node:fs';
import path from 'node:path';
import * as XLSXNamespace from 'xlsx';
import { cleanText, stripAllSpaces } from './biddingParser.js';
import { evaluateWorkbookFormulas, getCellFormulaError } from './excelFormulaEval.js';

const XLSX = (XLSXNamespace as unknown as { default?: typeof XLSXNamespace }).default ?? XLSXNamespace;

interface CellErrorSentinel {
  __cellError: string;
}

function isCellErrorSentinel(v: unknown): v is CellErrorSentinel {
  return typeof v === 'object' && v !== null && '__cellError' in v;
}

function assertCellValue(v: unknown): unknown {
  if (isCellErrorSentinel(v)) {
    throw new Error(v.__cellError);
  }
  return v;
}

function wsToAoaPreserveFormatAndErrors(ws: XLSXNamespace.WorkSheet, sheetName: string): unknown[][] {
  const ref = ws['!ref'];
  if (!ref) return [];
  const range = XLSX.utils.decode_range(ref);
  const aoa: unknown[][] = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row: unknown[] = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr] as XLSXNamespace.CellObject | undefined;
      if (!cell || cell.t === 'z') {
        row.push('');
      } else if (cell.t === 'e') {
        const msg = getCellFormulaError(cell) || `${sheetName}!${addr} 为错误单元格（${cell.w || cell.v || '#ERROR'}）`;
        row.push({ __cellError: msg } as CellErrorSentinel);
      } else if (cell.t === 'n' && typeof cell.v === 'number') {
        const fmt = typeof cell.z === 'string' ? cell.z.trim() : '';
        if (fmt && fmt !== 'General' && fmt !== '@' && cell.w !== undefined && cell.w !== null) {
          row.push(String(cell.w));
        } else {
          row.push(cell.v);
        }
      } else if (cell.w !== undefined && cell.w !== null) {
        row.push(String(cell.w));
      } else if (cell.v !== undefined && cell.v !== null) {
        row.push(cell.v);
      } else {
        row.push('');
      }
    }
    aoa.push(row);
  }
  return aoa;
}

function readFirstSheetAoa(filePath: string): unknown[][] {
  const buf = fs.readFileSync(filePath);
  const wb = XLSX.read(buf, { type: 'buffer', raw: true, sheetStubs: true, cellNF: true, cellStyles: true });
  evaluateWorkbookFormulas(wb);
  const firstName = wb.SheetNames[0];
  if (!firstName) return [];
  return wsToAoaPreserveFormatAndErrors(wb.Sheets[firstName], firstName);
}

function readAllSheetsAoa(filePath: string): unknown[][] {
  const buf = fs.readFileSync(filePath);
  const wb = XLSX.read(buf, { type: 'buffer', raw: true, sheetStubs: true, cellNF: true, cellStyles: true });
  evaluateWorkbookFormulas(wb);
  const merged: unknown[][] = [];
  for (const name of wb.SheetNames) {
    const rows = wsToAoaPreserveFormatAndErrors(wb.Sheets[name], name);
    merged.push(...rows, []);
  }
  return merged;
}

function readRawAoa(filePath: string): unknown[][] {
  return readAllSheetsAoa(filePath);
}

function colIndex(headers: string[], patterns: RegExp[]): number {
  return headers.findIndex((h) => patterns.some((p) => p.test(h)));
}

function toNum(val: unknown): number | null {
  if (val === null || val === undefined || val === '' || String(val).trim() === '/') return null;
  const n = Number(String(val).replace(/[,，\s%]/g, ''));
  return Number.isFinite(n) ? n : null;
}

/**
 * Strip lot suffix "-01", "-02" from a lot code like "CY1226SZN029-02" -> "CY1226SZN029".
 */
export function normalizeBatchCode(raw: string): string {
  const s = stripAllSpaces(raw).toUpperCase();
  // Strip one or more trailing numeric lot segments: CY1226SZN029-02 -> CY1226SZN029, 12EJ02-901100-9904 -> 12EJ02
  return s.replace(/(?:-\d+)+$/, '');
}

/**
 * Extract batch code from meta rows or data columns for cross-batch mismatch detection.
 */
function detectBatchCodeInAoa(aoa: unknown[][]): string {
  for (let i = 0; i < Math.min(aoa.length, 12); i++) {
    const line = (aoa[i] || []).map((c) => cleanText(c)).join(' ');
    const m1 = line.match(/(?:招标编号|采购项目编号|批次编号|分标编号)\s*[:：]\s*([A-Z0-9-]{6,})/i);
    if (m1) return normalizeBatchCode(m1[1]);
    const m2 = line.match(/\b([A-Z]{2}\d{4}[A-Z0-9]{4,})(?:-\d{2})?\b/);
    if (m2) return normalizeBatchCode(m2[1]);
  }
  return '';
}

/**
 * 1. Parse 开标一览表（Excel）(支持单轮/多轮)
 */
export function parseBidOpeningSheet(filePath: string) {
  // 开标一览表通常含两个 Sheet（开标记录报表 / 开标全量报表），严格仅以第一个 Sheet 为准
  const aoa = readFirstSheetAoa(filePath);
  let hIdx = 0;
  for (let i = 0; i < Math.min(aoa.length, 6); i++) {
    const rowText = (aoa[i] || []).map((c) => cleanText(c)).join(' ');
    if (/投标人名称|响应供应商/.test(rowText) && /分包名称|包号|分标编号/.test(rowText)) {
      hIdx = i;
      break;
    }
  }
  const headers = (aoa[hIdx] || []).map((v) => cleanText(v));
  const idxBatchNo = colIndex(headers, [/采购项目编号|招标编号|批次编号/]);
  const idxLotName = colIndex(headers, [/分标名称/]);
  const idxLotCode = colIndex(headers, [/分标编号/]);
  const idxPkg = colIndex(headers, [/分包名称|包号|包名/]);
  const idxRound = colIndex(headers, [/轮次/]);
  const idxSupplier = colIndex(headers, [/投标人名称|响应供应商|投标人/]);
  const idxDiscount = colIndex(headers, [/折扣率/]);
  const idxPrice = colIndex(headers, [/投标价格|报价|应答价格/]);
  const idxRemark = colIndex(headers, [/开标备注|备注/]);

  let detectedBatchCode = '';
  let spacesCleaned = 0;
  const roundsSet = new Set<string>();
  const rows: Array<{
    lot_code: string;
    lot_name: string;
    package_no: string;
    round_label: string;
    supplier_name: string;
    discount_rate: string;
    bid_price_wan: number | null;
    bid_price_raw: string;
    remark: string;
  }> = [];

  for (let i = hIdx + 1; i < aoa.length; i++) {
    const r = aoa[i] || [];
    const rawSup = idxSupplier >= 0 ? String(assertCellValue(r[idxSupplier]) ?? '') : '';
    const supplier = stripAllSpaces(rawSup);
    if (!supplier || /^(投标人名称|投标人|响应供应商|供应商名称)$/.test(supplier)) continue;
    const lotCodeCheck = idxLotCode >= 0 ? stripAllSpaces(r[idxLotCode]) : '';
    if (lotCodeCheck === '分标编号') continue;
    if (rawSup !== supplier) spacesCleaned++;

    if (!detectedBatchCode && idxBatchNo >= 0 && r[idxBatchNo]) {
      detectedBatchCode = normalizeBatchCode(String(assertCellValue(r[idxBatchNo])));
    }
    const lotCode = idxLotCode >= 0 ? stripAllSpaces(r[idxLotCode]) : '';
    if (!detectedBatchCode && lotCode) {
      detectedBatchCode = normalizeBatchCode(lotCode);
    }

    const roundLabel = idxRound >= 0 ? cleanText(r[idxRound]) || '第1轮' : '第1轮';
    roundsSet.add(roundLabel);

    const rawPrice = idxPrice >= 0 ? cleanText(r[idxPrice]) : '';
    const priceWan = toNum(rawPrice);

    rows.push({
      lot_code: lotCode,
      lot_name: idxLotName >= 0 ? cleanText(r[idxLotName]) : '',
      package_no: idxPkg >= 0 ? stripAllSpaces(r[idxPkg]) : '',
      round_label: roundLabel,
      supplier_name: supplier,
      discount_rate: idxDiscount >= 0 ? cleanText(r[idxDiscount]) : '',
      bid_price_wan: priceWan,
      bid_price_raw: rawPrice,
      remark: idxRemark >= 0 ? cleanText(r[idxRemark]) : '',
    });
  }

  return {
    detectedBatchCode: detectedBatchCode || detectBatchCodeInAoa(aoa),
    rows,
    rounds: Array.from(roundsSet),
    spacesCleaned,
  };
}

/**
 * 2. Parse 行信息报价基础数据表.xlsx (multi-block by lot)
 */
export function parseLineQuotesSheet(filePath: string) {
  const aoa = readRawAoa(filePath);
  let detectedBatchCode = '';
  let currentLotCode = '';
  let spacesCleaned = 0;

  let colMap = {
    pkg: -1,
    round: -1,
    supplier: -1,
    reqAppNo: -1,
    projName: -1,
    taxRate: -1,
    exTax: -1,
    incTax: -1,
    discount: -1,
  };

  const rows: Array<{
    lot_code: string;
    package_no: string;
    round_label: string;
    supplier_name: string;
    req_app_no: string;
    project_name: string;
    tax_rate: string;
    price_ex_tax_wan: number | null;
    price_inc_tax_wan: number | null;
    discount_rate: string;
  }> = [];

  for (let i = 0; i < aoa.length; i++) {
    const r = aoa[i] || [];
    const lineText = r.map((c) => cleanText(c)).join(' ').trim();
    if (!lineText) continue;

    // Block meta line: 招标编号：CY1226SZN027    分标编号：CY1226SZN027-01
    if (/招标编号\s*[:：]/.test(lineText) || /分标编号\s*[:：]/.test(lineText)) {
      const mBatch = lineText.match(/招标编号\s*[:：]\s*([A-Z0-9-]+)/i);
      const mLot = lineText.match(/分标编号\s*[:：]\s*([A-Z0-9-]+)/i);
      if (mBatch && !detectedBatchCode) detectedBatchCode = normalizeBatchCode(mBatch[1]);
      if (mLot) {
        currentLotCode = stripAllSpaces(mLot[1]);
        if (!detectedBatchCode) detectedBatchCode = normalizeBatchCode(currentLotCode);
      }
      continue;
    }

    // Header row inside each block
    if (/采购申请号/.test(lineText) && /投标人|响应供应商/.test(lineText)) {
      const headers = r.map((v) => cleanText(v));
      colMap = {
        pkg: colIndex(headers, [/包名|包号|分包名称/]),
        round: colIndex(headers, [/轮次/]),
        supplier: colIndex(headers, [/投标人|响应供应商/]),
        reqAppNo: colIndex(headers, [/采购申请号/]),
        projName: colIndex(headers, [/工程项目名称|项目名称/]),
        taxRate: colIndex(headers, [/增值税率/]),
        exTax: colIndex(headers, [/^未含税总价/]),
        incTax: colIndex(headers, [/^含税总价/]),
        discount: colIndex(headers, [/折扣率/]),
      };
      continue;
    }

    if (colMap.supplier >= 0) {
      const rawSup = String(r[colMap.supplier] ?? '');
      const supplier = stripAllSpaces(rawSup);
      if (!supplier || supplier === '投标人' || supplier === '无数据') continue;
      if (rawSup !== supplier) spacesCleaned++;

      rows.push({
        lot_code: currentLotCode,
        package_no: colMap.pkg >= 0 ? stripAllSpaces(r[colMap.pkg]) : '',
        round_label: colMap.round >= 0 ? cleanText(r[colMap.round]) || '第1轮' : '第1轮',
        supplier_name: supplier,
        req_app_no: colMap.reqAppNo >= 0 ? stripAllSpaces(r[colMap.reqAppNo]) : '',
        project_name: colMap.projName >= 0 ? cleanText(r[colMap.projName]) : '',
        tax_rate: colMap.taxRate >= 0 ? cleanText(r[colMap.taxRate]) : '',
        price_ex_tax_wan: colMap.exTax >= 0 ? toNum(r[colMap.exTax]) : null,
        price_inc_tax_wan: colMap.incTax >= 0 ? toNum(r[colMap.incTax]) : null,
        discount_rate: colMap.discount >= 0 ? cleanText(r[colMap.discount]) : '',
      });
    }
  }

  return {
    detectedBatchCode: detectedBatchCode || detectBatchCodeInAoa(aoa),
    rows,
    spacesCleaned,
  };
}

/**
 * 3. Parse 未进入详评情况统计表.xlsx
 */
export function parseRejectionsSheet(filePath: string) {
  const aoa = readRawAoa(filePath);
  let detectedBatchCode = detectBatchCodeInAoa(aoa);
  let singleLotCode = '';
  const lotNameToCode = new Map<string, string>();
  let hIdx = 0;

  for (let i = 0; i < Math.min(aoa.length, 8); i++) {
    const line = (aoa[i] || []).map((c) => cleanText(c)).join(' ');
    // Support multi-lot banner: 分标编号：A,B,C   分标名称:X,Y,Z
    const mCodes = line.match(/分标编号\s*[:：]\s*([^\s]+(?:\s*,\s*[^\s]+)*)/);
    const mNames = line.match(/分标名称\s*[:：]\s*(.+)$/);
    if (mCodes) {
      const codes = mCodes[1].split(/[,，]/).map((s) => stripAllSpaces(s)).filter(Boolean);
      if (codes.length > 0 && !detectedBatchCode) {
        detectedBatchCode = normalizeBatchCode(codes[0]);
      }
      if (codes.length === 1) {
        singleLotCode = codes[0];
      } else if (mNames) {
        const names = mNames[1].split(/[,，]/).map((s) => cleanText(s)).filter(Boolean);
        const len = Math.min(codes.length, names.length);
        for (let k = 0; k < len; k++) {
          lotNameToCode.set(names[k], codes[k]);
        }
      }
    }
    if (/未进入详评原因|否决原因/.test(line) && /投标人|响应供应商/.test(line)) {
      hIdx = i;
      break;
    }
  }

  const headers = (aoa[hIdx] || []).map((v) => cleanText(v));
  const idxLotCode = colIndex(headers, [/分标编号/]);
  const idxLotName = colIndex(headers, [/分标名称/]);
  const idxSupplier = colIndex(headers, [/投标人|响应供应商/]);
  const idxPkg = colIndex(headers, [/包号|子包号|分包/]);
  const idxReason = colIndex(headers, [/未进入详评原因|否决原因/]);
  const idxTpl = colIndex(headers, [/否决模板/]);
  const idxDocReq = colIndex(headers, [/招标文件要求|采购文件要求/]);
  const idxBidResp = colIndex(headers, [/投标文件响应|响应文件/]);
  const idxType = colIndex(headers, [/商务类|技术类/]);
  const idxNotes = colIndex(headers, [/备注/]);

  let spacesCleaned = 0;
  const rows: Array<{
    lot_code: string;
    lot_name: string;
    package_no: string;
    supplier_name: string;
    reason_category: string;
    rejection_tpl: string;
    doc_requirement: string;
    bid_response: string;
    review_type: string;
    notes: string;
  }> = [];

  for (let i = hIdx + 1; i < aoa.length; i++) {
    const r = aoa[i] || [];
    const rawSup = idxSupplier >= 0 ? String(r[idxSupplier] ?? '') : '';
    const supplier = stripAllSpaces(rawSup);
    if (!supplier || /^(投标人名称|投标人|响应供应商|供应商名称)$/.test(supplier)) continue;
    if (rawSup !== supplier) spacesCleaned++;

    const lotName = idxLotName >= 0 ? cleanText(r[idxLotName]) : '';
    const rowLotCode = idxLotCode >= 0 ? stripAllSpaces(r[idxLotCode]) : '';
    const resolvedLotCode = rowLotCode || lotNameToCode.get(lotName) || singleLotCode;

    rows.push({
      lot_code: resolvedLotCode,
      lot_name: lotName,
      package_no: idxPkg >= 0 ? stripAllSpaces(r[idxPkg]) : '',
      supplier_name: supplier,
      reason_category: idxReason >= 0 ? cleanText(r[idxReason]) : '',
      rejection_tpl: idxTpl >= 0 ? cleanText(r[idxTpl]) : '',
      doc_requirement: idxDocReq >= 0 ? cleanText(r[idxDocReq]) : '',
      bid_response: idxBidResp >= 0 ? cleanText(r[idxBidResp]) : '',
      review_type: idxType >= 0 ? cleanText(r[idxType]) : '',
      notes: idxNotes >= 0 ? cleanText(r[idxNotes]) : '',
    });
  }

  return {
    detectedBatchCode,
    rows,
    spacesCleaned,
  };
}

/**
 * 4. Parse 总分排序表 (最终版).xlsx (multi-block 表2-9 parser, supports merged cells & 流标 "无数据")
 */
export function parseRankingsSheet(filePath: string) {
  const aoa = readRawAoa(filePath);
  let detectedBatchCode = '';
  let currentLotCode = '';
  let currentLotName = '';
  let currentPkg = '';
  let currentEvalDate = '';
  let currentProjectName = '';
  let spacesCleaned = 0;

  let colMap = {
    pkg: -1,
    projName: -1,
    supplier: -1,
    tech: -1,
    biz: -1,
    price: -1,
    total: -1,
    rank: -1,
    basePrice: -1,
    bidPrice: -1,
    evalPrice: -1,
  };

  const rows: Array<{
    lot_code: string;
    lot_name: string;
    package_no: string;
    project_name: string;
    eval_date: string;
    supplier_name: string;
    tech_score: number | null;
    biz_score: number | null;
    price_score: number | null;
    total_score: number | null;
    rank_no: number | null;
    base_price: string;
    bid_price: string;
    eval_price: string;
    is_failed_lot: number;
  }> = [];

  for (let i = 0; i < aoa.length; i++) {
    const r = aoa[i] || [];
    const lineText = r.map((c) => cleanText(c)).join(' ').trim();
    if (!lineText) continue;

    // Block meta line: 设备/材料名称:   招标编号: CY1226SZN029-02  包号: 包1_综合服务  项目单位: 无;  日期:2026-10-21
    if (/招标编号\s*[:：]/.test(lineText) && /包号\s*[:：]/.test(lineText)) {
      const mLot = lineText.match(/招标编号\s*[:：]\s*([A-Z0-9-]+)/i);
      const mPkg = lineText.match(/包号\s*[:：]\s*([^\s]+)/);
      const mDate = lineText.match(/日期\s*[:：]\s*([\d-]+)/);
      if (mLot) {
        currentLotCode = stripAllSpaces(mLot[1]);
        if (!detectedBatchCode) detectedBatchCode = normalizeBatchCode(currentLotCode);
      }
      if (mPkg) {
        const rawPkg = mPkg[1];
        const parts = rawPkg.split('_');
        currentPkg = stripAllSpaces(parts[0]);
        currentLotName = parts[1] ? cleanText(parts[1]) : '';
      }
      currentEvalDate = mDate ? mDate[1] : currentEvalDate;
      currentProjectName = '';
      continue;
    }

    // Header row inside each block
    if (/包号\/子包号/.test(lineText) && /投标人|响应供应商/.test(lineText)) {
      const headers = r.map((v) => cleanText(v));
      colMap = {
        pkg: colIndex(headers, [/包号\/子包号|包号/]),
        projName: colIndex(headers, [/项目名称/]),
        supplier: colIndex(headers, [/投标人|响应供应商/]),
        tech: colIndex(headers, [/^技术$/]),
        biz: colIndex(headers, [/^商务$/]),
        price: colIndex(headers, [/^价格$/]),
        total: colIndex(headers, [/^总分$/]),
        rank: colIndex(headers, [/^排序$/]),
        basePrice: colIndex(headers, [/基准价/]),
        bidPrice: colIndex(headers, [/投标总价|应答总价/]),
        evalPrice: colIndex(headers, [/评标总价|评审总价/]),
      };
      continue;
    }

    // Handle failed lot ("无数据")
    if (/^无数据/.test(cleanText(r[0])) && currentLotCode && currentPkg) {
      rows.push({
        lot_code: currentLotCode,
        lot_name: currentLotName,
        package_no: currentPkg,
        project_name: currentProjectName,
        eval_date: currentEvalDate,
        supplier_name: '',
        tech_score: null,
        biz_score: null,
        price_score: null,
        total_score: null,
        rank_no: null,
        base_price: '',
        bid_price: '',
        eval_price: '',
        is_failed_lot: 1,
      });
      continue;
    }

    if (colMap.supplier >= 0 && currentLotCode) {
      const rawSup = String(r[colMap.supplier] ?? '');
      const supplier = stripAllSpaces(rawSup);
      if (!supplier || supplier === '投标人') continue;
      if (rawSup !== supplier) spacesCleaned++;

      const rowProjName = colMap.projName >= 0 ? cleanText(r[colMap.projName]).replace(/[;；]+$/, '') : '';
      if (rowProjName) currentProjectName = rowProjName;

      const rankNum = colMap.rank >= 0 ? toNum(r[colMap.rank]) : null;
      if (rankNum === null && colMap.total >= 0 && toNum(r[colMap.total]) === null) continue;

      rows.push({
        lot_code: currentLotCode,
        lot_name: currentLotName,
        package_no: currentPkg,
        project_name: currentProjectName,
        eval_date: currentEvalDate,
        supplier_name: supplier,
        tech_score: colMap.tech >= 0 ? toNum(r[colMap.tech]) : null,
        biz_score: colMap.biz >= 0 ? toNum(r[colMap.biz]) : null,
        price_score: colMap.price >= 0 ? toNum(r[colMap.price]) : null,
        total_score: colMap.total >= 0 ? toNum(r[colMap.total]) : null,
        rank_no: rankNum,
        base_price: colMap.basePrice >= 0 ? cleanText(r[colMap.basePrice]) : '',
        bid_price: colMap.bidPrice >= 0 ? cleanText(r[colMap.bidPrice]) : '',
        eval_price: colMap.evalPrice >= 0 ? cleanText(r[colMap.evalPrice]) : '',
        is_failed_lot: 0,
      });
    }
  }

  const failedLotsCount = rows.filter((r) => r.is_failed_lot === 1).length;
  return {
    detectedBatchCode: detectedBatchCode || detectBatchCodeInAoa(aoa),
    rows,
    failedLotsCount,
    spacesCleaned,
  };
}

/**
 * 5. Parse 投标文件IP地址一览表.xlsx (hardware & IP records for collusion verification)
 * Preserves exact raw trimmed text for MAC / CPU / Disk per exact-match rule,
 * and reads sheet-by-sheet with blankrows: true to preserve exact Excel sheet name & 1-based row number.
 */
export function parseIpHardwareSheet(filePath: string, originalFileName?: string) {
  const buf = fs.readFileSync(filePath);
  const wb = XLSX.read(buf, { type: 'buffer', raw: true, sheetStubs: true, cellNF: true, cellStyles: true });
  evaluateWorkbookFormulas(wb);
  const sourceFile = (originalFileName && originalFileName.trim()) || path.basename(filePath);

  let detectedBatchCode = '';
  let fallbackBatchCode = '';
  let spacesCleaned = 0;

  const rows: Array<{
    lot_code: string;
    lot_name: string;
    package_no: string;
    round_label: string;
    supplier_name: string;
    download_ip: string;
    tech_ip: string;
    biz_ip: string;
    price_ip: string;
    qual_ip: string;
    tech_mac: string;
    biz_mac: string;
    price_mac: string;
    qual_mac: string;
    tech_cpu: string;
    biz_cpu: string;
    price_cpu: string;
    qual_cpu: string;
    tech_disk: string;
    biz_disk: string;
    price_disk: string;
    qual_disk: string;
    source_file: string;
    sheet_name: string;
    source_row: number;
  }> = [];

  const PLACEHOLDER_RE = /^(?:[\/／\-－—–]+|无|null|n\/a|na)$/i;
  const exactCell = (r: unknown[], idx: number): string => {
    if (idx < 0) return '';
    const v = String(r[idx] ?? '').trim();
    return !v || PLACEHOLDER_RE.test(v) ? '' : v;
  };

  const warnings: string[] = [];
  let headerMatched = false;

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws || !ws['!ref']) continue;

    const rangeStartRow = XLSX.utils.decode_range(ws['!ref']).s.r;
    const aoa = XLSX.utils.sheet_to_json<unknown[]>(ws, {
      header: 1,
      defval: '',
      raw: true,
      blankrows: true,
    });
    if (!fallbackBatchCode) {
      fallbackBatchCode = detectBatchCodeInAoa(aoa);
    }

    let colMap = {
      lotName: -1,
      lotCode: -1,
      pkg: -1,
      round: -1,
      supplier: -1,
      downloadIp: -1,
      techIp: -1,
      bizIp: -1,
      priceIp: -1,
      qualIp: -1,
      techMac: -1,
      bizMac: -1,
      priceMac: -1,
      qualMac: -1,
      techCpu: -1,
      bizCpu: -1,
      priceCpu: -1,
      qualCpu: -1,
      techDisk: -1,
      bizDisk: -1,
      priceDisk: -1,
      qualDisk: -1,
    };

    for (let i = 0; i < aoa.length; i++) {
      const excelRowNum = rangeStartRow + i + 1;
      const r = aoa[i] || [];
      const rowText = r.map((c) => cleanText(c)).join(' ').trim();
      if (!rowText) continue;

      if (/分标编号|分标名称/.test(rowText) && /投标人|响应供应商/.test(rowText) && /MAC|CPU|硬盘|IP/i.test(rowText)) {
        const headers = r.map((v) => cleanText(v));
        headerMatched = true;
        colMap = {
          lotName: colIndex(headers, [/^分标名称$/]),
          lotCode: colIndex(headers, [/^分标编号$/]),
          pkg: colIndex(headers, [/分包名称|包号|包名/]),
          round: colIndex(headers, [/^轮次$/]),
          supplier: colIndex(headers, [/^投标人$|^投标人名称$|^响应供应商$|^供应商名称$/]),
          downloadIp: colIndex(headers, [/下载.*IP/i]),
          techIp: colIndex(headers, [/技术文件.*IP|技术.*提交.*IP/i]),
          bizIp: colIndex(headers, [/商务文件.*IP|商务.*提交.*IP/i]),
          priceIp: colIndex(headers, [/价格文件.*IP|价格.*提交.*IP/i]),
          qualIp: colIndex(headers, [/(?:资格|资质)业绩.*提交.*(?:IP|地址)|(?:资格|资质)业绩.*IP/i]),
          techMac: colIndex(headers, [/技术文件.*MAC/i]),
          bizMac: colIndex(headers, [/商务文件.*MAC/i]),
          priceMac: colIndex(headers, [/价格文件.*MAC/i]),
          qualMac: colIndex(headers, [/(?:资格|资质)业绩.*MAC/i]),
          techCpu: colIndex(headers, [/技术文件.*CPU/i]),
          bizCpu: colIndex(headers, [/商务文件.*CPU/i]),
          priceCpu: colIndex(headers, [/价格文件.*CPU/i]),
          qualCpu: colIndex(headers, [/(?:资格|资质)业绩.*CPU/i]),
          techDisk: colIndex(headers, [/技术文件.*硬盘/i]),
          bizDisk: colIndex(headers, [/商务文件.*硬盘/i]),
          priceDisk: colIndex(headers, [/价格文件.*硬盘/i]),
          qualDisk: colIndex(headers, [/(?:资格|资质)业绩.*硬盘/i]),
        };
        const groupChecks = [
          { label: '技术文件', cols: [colMap.techCpu, colMap.techDisk, colMap.techMac] },
          { label: '商务文件', cols: [colMap.bizCpu, colMap.bizDisk, colMap.bizMac] },
          { label: '价格文件', cols: [colMap.priceCpu, colMap.priceDisk, colMap.priceMac] },
          { label: '资格/资质业绩文件', cols: [colMap.qualCpu, colMap.qualDisk, colMap.qualMac] },
        ];
        const missingGroups = groupChecks.filter((g) => g.cols.every((idx) => idx < 0)).map((g) => g.label);
        if (missingGroups.length > 0) {
          const msg = `表头未识别到以下类别的硬件特征列（CPU/硬盘/MAC）：${missingGroups.join('、')}，如原表含此类列请检查表头命名`;
          if (!warnings.includes(msg)) warnings.push(msg);
        }
        continue;
      }

      if (colMap.supplier >= 0) {
        const rawSup = String(r[colMap.supplier] ?? '');
        const supplier = stripAllSpaces(rawSup);
        if (!supplier || /^(投标人|投标人名称|响应供应商|供应商名称)$/.test(supplier)) continue;
        if (rawSup !== supplier) spacesCleaned++;

        const lotCode = colMap.lotCode >= 0 ? stripAllSpaces(r[colMap.lotCode]) : '';
        if (lotCode === '分标编号') continue;
        if (!detectedBatchCode && lotCode) detectedBatchCode = normalizeBatchCode(lotCode);

        const rawRound = colMap.round >= 0 ? cleanText(r[colMap.round]) : '';
        const roundLabel = /^\d+$/.test(rawRound) ? `第${rawRound}轮` : rawRound || '第1轮';

        rows.push({
          lot_code: lotCode,
          lot_name: colMap.lotName >= 0 ? cleanText(r[colMap.lotName]) : '',
          package_no: colMap.pkg >= 0 ? stripAllSpaces(r[colMap.pkg]) : '',
          round_label: roundLabel,
          supplier_name: supplier,
          download_ip: exactCell(r, colMap.downloadIp),
          tech_ip: exactCell(r, colMap.techIp),
          biz_ip: exactCell(r, colMap.bizIp),
          price_ip: exactCell(r, colMap.priceIp),
          qual_ip: exactCell(r, colMap.qualIp),
          tech_mac: exactCell(r, colMap.techMac),
          biz_mac: exactCell(r, colMap.bizMac),
          price_mac: exactCell(r, colMap.priceMac),
          qual_mac: exactCell(r, colMap.qualMac),
          tech_cpu: exactCell(r, colMap.techCpu),
          biz_cpu: exactCell(r, colMap.bizCpu),
          price_cpu: exactCell(r, colMap.priceCpu),
          qual_cpu: exactCell(r, colMap.qualCpu),
          tech_disk: exactCell(r, colMap.techDisk),
          biz_disk: exactCell(r, colMap.bizDisk),
          price_disk: exactCell(r, colMap.priceDisk),
          qual_disk: exactCell(r, colMap.qualDisk),
          source_file: sourceFile,
          sheet_name: sheetName,
          source_row: excelRowNum,
        });
      }
    }
  }

  if (!headerMatched) {
    warnings.push('未识别到标准硬件特征表头行（需包含分标编号/分标名称、投标人及 MAC/CPU/硬盘/IP 列）');
  }
  return {
    detectedBatchCode: detectedBatchCode || fallbackBatchCode,
    rows,
    spacesCleaned,
    warnings,
  };
}
