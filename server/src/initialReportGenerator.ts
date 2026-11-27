import PizZip from 'pizzip';
import XLSX from 'xlsx';
import { evaluateWorkbookFormulas, getCellFormulaError } from './excelFormulaEval.js';

export const DEFAULT_COMPANY_SUFFIXES =
  '有限责任公司,有限公司,公司,集团,股份有限责任公司,园区,股份有限公司';

export interface InitialReportOptionsInput {
  department?: string;
  clarification?: string;
  simplifyNames?: boolean;
  companySuffixes?: string;
}

export interface InitialReportSectionSummary {
  section: string;
  reason: string;
  items: Array<{ company: string; displayName: string; packs: string }>;
  sentence: string;
}

export interface InitialReportResult {
  fileName: string;
  dept: string;
  leader: string;
  members: string;
  sectionCount: number;
  packCount: number;
  bidderCount: number;
  documentCount: number;
  rejectionRowCount: number;
  clarification: string;
  simplifyNames: boolean;
  companySuffixes: string;
  paragraphs: string[];
  groupedRejections: InitialReportSectionSummary[];
  warnings: string[];
  sheetNames: {
    detailSheet: string;
    summarySheet: string;
    coverSheet: string;
  };
}

export function parseSuffixes(raw: string): string[] {
  const seen = new Set<string>();
  const list: string[] = [];
  for (const token of String(raw ?? '').split(/[,，、\r\n]/)) {
    const t = token.trim();
    if (!t || seen.has(t)) continue;
    seen.add(t);
    list.push(t);
  }
  return list.sort((a, b) => b.length - a.length);
}

export function computeShortCompanyName(full: string, suffixes: string[]): string {
  let value = full.trim();
  let changed = true;
  while (changed) {
    changed = false;
    for (const suffix of suffixes) {
      if (suffix && value.endsWith(suffix)) {
        const next = value.slice(0, value.length - suffix.length).trim();
        if (!next) return full;
        value = next;
        changed = true;
        break;
      }
    }
  }
  return value;
}

export function buildCompanyDisplayNames(fullNames: string[], suffixes: string[]): Map<string, string> {
  const result = new Map<string, string>();
  const counts = new Map<string, number>();
  for (const full of new Set(fullNames)) {
    const shortName = computeShortCompanyName(full, suffixes);
    result.set(full, shortName);
    counts.set(shortName, (counts.get(shortName) || 0) + 1);
  }
  for (const [full, shortName] of result.entries()) {
    if ((counts.get(shortName) || 0) > 1) {
      result.set(full, full);
    }
  }
  return result;
}

/**
 * Merge package numbers like ["包1", "包2", "包3", "包5"] -> "包1至包3，包5"
 * Supports "包1", "包1至包6", and comma/dunhao separated lists.
 */
export function mergePackageIntervals(values: Iterable<string>): string {
  const intervals: Array<{ start: number; end: number }> = [];
  const packRe = /^包?\s*(\d+)\s*(?:至\s*包?\s*(\d+))?$/;

  for (const raw of values) {
    for (const token of String(raw ?? '').split(/[,，、]/)) {
      const trimmed = token.trim();
      if (!trimmed) continue;
      const m = trimmed.match(packRe);
      if (!m) {
        throw new Error(`无法识别包号：${raw}；请使用“包1”、“包1至包3”或逗号分隔的包号`);
      }
      const a = parseInt(m[1], 10);
      const b = m[2] ? parseInt(m[2], 10) : a;
      if (!Number.isFinite(a) || !Number.isFinite(b) || a <= 0 || b <= 0) {
        throw new Error(`包号须为正整数：${raw}`);
      }
      intervals.push({ start: Math.min(a, b), end: Math.max(a, b) });
    }
  }

  intervals.sort((x, y) => (x.start !== y.start ? x.start - y.start : x.end - y.end));
  const merged: Array<{ start: number; end: number }> = [];
  for (const next of intervals) {
    if (merged.length === 0) {
      merged.push({ ...next });
      continue;
    }
    const prev = merged[merged.length - 1];
    if (next.start <= prev.end + 1) {
      prev.end = Math.max(prev.end, next.end);
    } else {
      merged.push({ ...next });
    }
  }

  return merged.map((r) => (r.start === r.end ? `包${r.start}` : `包${r.start}至包${r.end}`)).join('，');
}

function getCellText(ws: XLSX.WorkSheet, sheetName: string, r: number, c: number): string {
  const cell = ws[XLSX.utils.encode_cell({ r, c })] as XLSX.CellObject | undefined;
  if (!cell || cell.t === 'z') return '';
  if (cell.t === 'e') {
    const fErr = getCellFormulaError(cell);
    throw new Error(fErr || `${sheetName}!${XLSX.utils.encode_cell({ r, c })} 为错误单元格`);
  }
  if (cell.w !== undefined && cell.w !== null) return String(cell.w);
  if (cell.v === undefined || cell.v === null) return '';
  return String(cell.v);
}

function getCellNumericText(ws: XLSX.WorkSheet, sheetName: string, r: number, c: number): string {
  const cell = ws[XLSX.utils.encode_cell({ r, c })] as XLSX.CellObject | undefined;
  if (!cell || cell.t === 'z') return '';
  if (cell.t === 'e') {
    const fErr = getCellFormulaError(cell);
    throw new Error(fErr || `${sheetName}!${XLSX.utils.encode_cell({ r, c })} 为错误单元格`);
  }
  if (cell.t === 'n' && typeof cell.v === 'number') {
    return Number(cell.v.toFixed(10)).toString();
  }
  if (cell.v !== undefined && cell.v !== null) return String(cell.v).trim();
  return cell.w ? String(cell.w).trim() : '';
}

function getLastRow(ws: XLSX.WorkSheet): number {
  const ref = ws['!ref'];
  if (!ref) return -1;
  return XLSX.utils.decode_range(ref).e.r;
}

function parseNonNegInt(value: string): number {
  const trimmed = value.trim();
  if (!trimmed) return 0;
  const n = Number(trimmed);
  if (!Number.isFinite(n) || n < 0 || !Number.isInteger(n)) {
    throw new Error(`数量应为非负整数：${value}`);
  }
  return n;
}

/**
 * Resolve the 3 sheets cleanly regardless of tab order:
 * - detailSheet: 附件/附表/否决明细 (skips 封面1/封面2)
 * - summarySheet: contains "一、基本情况" (usually 封面二)
 * - coverSheet: contains "组长：" or named 封面一/封面1
 */
const AUX_SHEET_NAME_RE = /示例|样例|参考|说明|填表|流转|指引|草稿|历史记录|签收|收件|清册/i;
const DETAIL_SHEET_NAME_RE = /否决|初评|未进入详评|附件|附表/i;
const PACK_CELL_RE = /^包?\s*\d+/;

function hasRejectionDetailHeader(ws: XLSX.WorkSheet, name: string): boolean {
  const lastR = getLastRow(ws);
  if (lastR < 0) return false;
  const headerMaxR = Math.min(lastR, 2);
  for (let r = 0; r <= headerMaxR; r++) {
    const b = getCellText(ws, name, r, 1).trim();
    const c = getCellText(ws, name, r, 2).trim();
    const d = getCellText(ws, name, r, 3).trim();
    const e = getCellText(ws, name, r, 4).trim();
    let hits = 0;
    if (/分标|标段/.test(b)) hits++;
    if (/问题|原因|否决|描述|情况/.test(c)) hits++;
    if (/投标人|供应商|企业|单位|名称/.test(d)) hits++;
    if (/包/.test(e)) hits++;
    if (hits >= 3) return true;
  }
  return false;
}

function hasRejectionDetailDataRows(ws: XLSX.WorkSheet, name: string): boolean {
  const lastR = getLastRow(ws);
  if (lastR < 2) return false;
  const dataMaxR = Math.min(lastR, 50);
  for (let r = 2; r <= dataMaxR; r++) {
    const b = getCellText(ws, name, r, 1).trim();
    const c = getCellText(ws, name, r, 2).trim();
    const d = getCellText(ws, name, r, 3).trim();
    const e = getCellText(ws, name, r, 4).trim();
    if (b && c && d && e && PACK_CELL_RE.test(e)) return true;
  }
  return false;
}

function resolveReportSheets(wb: XLSX.WorkBook): {
  detailName: string;
  summaryName: string;
  coverName: string;
} {
  if (wb.SheetNames.length < 3) {
    throw new Error('初评工作簿需要至少 3 张工作表（初评情况汇报附件、封面二/基本情况、封面一/组长信息）');
  }

  // Exclude hidden sheets from automatic selection
  const hiddenSheets = new Set<string>();
  const sheetsMeta = wb.Workbook?.Sheets || [];
  wb.SheetNames.forEach((name, idx) => {
    const meta = sheetsMeta[idx] as { Hidden?: number } | undefined;
    if (meta && meta.Hidden && meta.Hidden > 0) {
      hiddenSheets.add(name);
    }
  });
  const visibleNames = wb.SheetNames.filter((n) => !hiddenSheets.has(n));
  const activeNames = visibleNames.length >= 3 ? visibleNames : wb.SheetNames;

  let summaryName = '';
  let coverName = '';
  let detailName = '';

  for (const name of activeNames) {
    const ws = wb.Sheets[name];
    const lastR = Math.min(getLastRow(ws), 40);
    let hasBasic = false;
    let hasLeader = false;
    for (let r = 0; r <= lastR; r++) {
      const a = getCellText(ws, name, r, 0).trim();
      if (a === '一、基本情况') hasBasic = true;
      if (a.includes('组长：') || a.includes('组长:')) hasLeader = true;
    }
    if (hasBasic && !summaryName) summaryName = name;
    else if (hasLeader && !coverName) coverName = name;
  }

  const isCoverName = (n: string) => /封面\s*[12一二]/i.test(n);
  if (!summaryName) {
    summaryName = activeNames.find((n) => n !== coverName) || activeNames[1];
  }
  if (!coverName) {
    coverName = activeNames.find((n) => n !== summaryName) || activeNames[2];
  }

  const candidates = activeNames.filter((n) => n !== summaryName && n !== coverName);
  const primaryCandidates = candidates.filter((n) => !AUX_SHEET_NAME_RE.test(n) && !isCoverName(n));

  // Priority 1: visible non-auxiliary sheet(s) with standard rejection detail B..E header
  const headerMatches = primaryCandidates.filter((n) => hasRejectionDetailHeader(wb.Sheets[n], n));
  if (headerMatches.length === 1) {
    detailName = headerMatches[0];
  } else if (headerMatches.length > 1) {
    throw new Error(
      `工作簿中存在多张疑似否决明细的工作表（${headerMatches.join('、')}），无法唯一确定，请仅保留 1 张否决明细表后重试`
    );
  } else {
    // Priority 2: visible sheet whose name identifies it as the formal rejection/attachment detail sheet (even when 0 rejections)
    const namedMatches = primaryCandidates.filter((n) => DETAIL_SHEET_NAME_RE.test(n));
    if (namedMatches.length === 1) {
      detailName = namedMatches[0];
    } else {
      // Priority 3: visible non-auxiliary sheet(s) with B..E data rows where col E matches package format
      const dataMatches = primaryCandidates.filter((n) => hasRejectionDetailDataRows(wb.Sheets[n], n));
      if (dataMatches.length === 1) {
        detailName = dataMatches[0];
      } else if (dataMatches.length > 1) {
        throw new Error(
          `工作簿中存在多张疑似否决明细的工作表（${dataMatches.join('、')}），无法唯一确定，请仅保留 1 张否决明细表后重试`
        );
      } else if (primaryCandidates.length === 1) {
        detailName = primaryCandidates[0];
      } else {
        throw new Error(
          `未能从候选工作表（${candidates.join('、') || '无'}）中识别出否决明细表（需在 B~E 列包含分标、否决原因、投标人、包号），请检查工作簿结构`
        );
      }
    }
  }

  return { detailName, summaryName, coverName };
}

export function generateInitialReportData(
  buffer: Buffer,
  fileName: string,
  options: InitialReportOptionsInput = {}
): InitialReportResult {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: false, sheetStubs: true, cellNF: true, cellStyles: true });
  evaluateWorkbookFormulas(wb);
  const { detailName, summaryName, coverName } = resolveReportSheets(wb);

  const reasonsWs = wb.Sheets[detailName];
  const baseWs = wb.Sheets[summaryName];
  const coverWs = wb.Sheets[coverName];

  // 1. Parse rejection details from 附件 sheet (row 3+ / 0-based r >= 2, cols B..E)
  // sections -> reason -> company -> Set<pack>
  const sections = new Map<string, Map<string, Map<string, Set<string>>>>();
  let rejectionRowCount = 0;
  const reasonsLastRow = getLastRow(reasonsWs);
  for (let r = 2; r <= reasonsLastRow; r++) {
    const section = getCellText(reasonsWs, detailName, r, 1).trim();
    const reason = getCellText(reasonsWs, detailName, r, 2).trim();
    const company = getCellText(reasonsWs, detailName, r, 3).trim();
    const pack = getCellText(reasonsWs, detailName, r, 4).trim();
    if (!section && !reason && !company && !pack) continue;
    if (!section || !reason || !company || !pack) {
      throw new Error(`「${detailName}」第 ${r + 1} 行 B 至 E 列（分标名称/问题描述/投标人/包号）存在空值`);
    }
    rejectionRowCount++;
    let reasonMap = sections.get(section);
    if (!reasonMap) {
      reasonMap = new Map();
      sections.set(section, reasonMap);
    }
    let companyMap = reasonMap.get(reason);
    if (!companyMap) {
      companyMap = new Map();
      reasonMap.set(reason, companyMap);
    }
    let packSet = companyMap.get(company);
    if (!packSet) {
      packSet = new Set();
      companyMap.set(company, packSet);
    }
    packSet.add(pack);
  }

  // 2. Parse basic stats from 封面二 sheet ("一、基本情况" -> "二、澄清事项")
  const baseLastRow = getLastRow(baseWs);
  let start = -1;
  let end = -1;
  for (let r = 0; r <= baseLastRow; r++) {
    const val = getCellText(baseWs, summaryName, r, 0).trim();
    if (val === '一、基本情况') start = r;
    if (val === '二、澄清事项') {
      end = r;
      break;
    }
  }
  if (start < 0 || end < start + 2) {
    throw new Error(`「${summaryName}」缺少或错置“一、基本情况”/“二、澄清事项”标记`);
  }

  let sectionCount = 0;
  let packCount = 0;
  let bidderCount = 0;
  let documentCount = 0;
  for (let r = start + 2; r < end; r++) {
    const code = getCellText(baseWs, summaryName, r, 0).trim();
    const name = getCellText(baseWs, summaryName, r, 1).trim();
    const pCnt = getCellNumericText(baseWs, summaryName, r, 2);
    const bCnt = getCellNumericText(baseWs, summaryName, r, 3);
    const dCnt = getCellNumericText(baseWs, summaryName, r, 4);
    if (!code && !name) {
      if (!pCnt && !bCnt && !dCnt) continue;
      throw new Error(`「${summaryName}」第 ${r + 1} 行有数量但缺少分标编号和分标名称`);
    }
    sectionCount++;
    packCount += parseNonNegInt(pCnt);
    bidderCount += parseNonNegInt(bCnt);
    documentCount += parseNonNegInt(dCnt);
  }

  // 3. Parse group & leader info from 封面一 sheet (A2 & A11, with smart scan fallback)
  let groupRaw = getCellText(coverWs, coverName, 1, 0).trim();
  let peopleRaw = getCellText(coverWs, coverName, 10, 0).trim();
  if (!peopleRaw.includes('组长：') && !peopleRaw.includes('组长:')) {
    const coverLast = Math.min(getLastRow(coverWs), 30);
    for (let r = 0; r <= coverLast; r++) {
      const t = getCellText(coverWs, coverName, r, 0).trim();
      if (t.includes('组长：') || t.includes('组长:')) {
        peopleRaw = t;
        break;
      }
    }
  }

  let dept = String(options.department ?? '').trim();
  if (!dept) {
    const parenMatch = groupRaw.match(/[（(]([^()（）]*组)[）)]/);
    if (parenMatch && parenMatch[1].trim()) {
      dept = parenMatch[1].trim();
    } else if (groupRaw.length >= 5) {
      dept = groupRaw.slice(groupRaw.length - 5, groupRaw.length - 2).trim();
    } else {
      throw new Error(`「${coverName}」A2 缺少组别说明，请在界面填写组别参数`);
    }
  }
  if (!dept.endsWith('组')) dept += '组';

  const normalizedPeople = peopleRaw.replace('组长:', '组长：');
  const marker = normalizedPeople.indexOf('组长：');
  const sep = normalizedPeople.indexOf('、', Math.max(0, marker));
  if (marker < 0 || sep <= marker + 3 || sep === normalizedPeople.length - 1) {
    throw new Error(`「${coverName}」A11 应包含“组长：姓名、姓名”格式（当前为：${peopleRaw || '空'}）`);
  }
  const leader = normalizedPeople.slice(marker + 3, sep).trim();
  const members = normalizedPeople.slice(sep + 1).trim();

  // 4. Build paragraphs & simplified company display names
  const simplifyNames = options.simplifyNames !== undefined ? Boolean(options.simplifyNames) : true;
  const companySuffixes =
    options.companySuffixes !== undefined ? String(options.companySuffixes) : DEFAULT_COMPANY_SUFFIXES;
  const suffixList = simplifyNames ? parseSuffixes(companySuffixes) : [];

  const allCompanies: string[] = [];
  for (const reasonMap of sections.values()) {
    for (const companyMap of reasonMap.values()) {
      for (const c of companyMap.keys()) allCompanies.push(c);
    }
  }
  const displayNames = buildCompanyDisplayNames(allCompanies, suffixList);

  const warnings: string[] = [];
  let clarification = String(options.clarification ?? '').trim();
  if (!clarification) {
    clarification = '【待核实填写】';
    warnings.push('澄清事项尚未填写（输出已标注“【待核实填写】”），请核实后在界面补充再导出正式汇报单。');
  }

  const paragraphs: string[] = [
    `各位领导好，我们是${dept}的组长，我是${leader}，这位是${members}。下面由我向评委会汇报${dept}初评情况。`,
    `本组所评共${sectionCount}个标段、${packCount}个包，在投标截止时间前共${bidderCount}个投标人成功递交了${documentCount}份投标文件，各标段基本情况见输入表格。`,
    `澄清事项：${clarification}。`,
    '建议否决投标事项：',
  ];

  const groupedRejections: InitialReportSectionSummary[] = [];
  if (sections.size === 0) {
    paragraphs.push('输入明细中未列出否决事项。');
  } else {
    for (const [section, reasonMap] of sections.entries()) {
      for (const [reason, companyMap] of reasonMap.entries()) {
        const items: Array<{ company: string; displayName: string; packs: string }> = [];
        const parts: string[] = [];
        for (const [company, packSet] of companyMap.entries()) {
          const displayName = displayNames.get(company) || company;
          const packsStr = mergePackageIntervals(packSet);
          items.push({ company, displayName, packs: packsStr });
          parts.push(`${displayName}的${packsStr}`);
        }
        const sentence = `${section}标段，因${reason}而建议否决的有：${parts.join('；')}。`;
        paragraphs.push(sentence);
        groupedRejections.push({ section, reason, items, sentence });
      }
    }
  }

  paragraphs.push('汇报完毕！');

  return {
    fileName,
    dept,
    leader,
    members,
    sectionCount,
    packCount,
    bidderCount,
    documentCount,
    rejectionRowCount,
    clarification,
    simplifyNames,
    companySuffixes,
    paragraphs,
    groupedRejections,
    warnings,
    sheetNames: {
      detailSheet: detailName,
      summarySheet: summaryName,
      coverSheet: coverName,
    },
  };
}

function escapeXml(str: string): string {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Build `初评汇报单.docx` matching TenderToolbox Desktop's exact XWPF styling:
 * - Centered bold 18pt (w:sz="36") 宋体 title: "初评汇报单"
 * - 14pt (w:sz="28") 宋体 body paragraphs with 560 twips first-line indent and 160 twips after spacing
 */
export function buildInitialReportDocxBuffer(paragraphs: string[]): Buffer {
  const titleXml =
    `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>` +
    `<w:r><w:rPr><w:rFonts w:ascii="宋体" w:hAnsi="宋体" w:cs="宋体" w:eastAsia="宋体"/><w:b w:val="true"/><w:sz w:val="36"/></w:rPr>` +
    `<w:t>初评汇报单</w:t></w:r></w:p>`;

  const bodyXml = paragraphs
    .map(
      (text) =>
        `<w:p><w:pPr><w:spacing w:after="160"/><w:ind w:firstLine="560"/></w:pPr>` +
        `<w:r><w:rPr><w:rFonts w:ascii="宋体" w:hAnsi="宋体" w:cs="宋体" w:eastAsia="宋体"/><w:sz w:val="28"/></w:rPr>` +
        `<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`
    )
    .join('');

  const zip = new PizZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>${titleXml}${bodyXml}<w:sectPr/></w:body>
</w:document>`
  );

  return zip.generate({ type: 'nodebuffer' });
}
