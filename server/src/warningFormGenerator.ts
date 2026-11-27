import fs from 'node:fs';
import path from 'node:path';
import PizZip from 'pizzip';
import XLSX from 'xlsx';
import { evaluateWorkbookFormulas, getCellFormulaError } from './excelFormulaEval.js';

export const RISK_HEADERS = [
  '分标名称',
  '分包名称',
  '投标人',
  '注册地址',
  '企业联系方式',
  '邮箱',
  '注册地址是否相同',
  '企业联系方式是否相同',
  '邮箱是否相同',
] as const;

export const GROUP_HEADERS = ['分标名称', '分包名称', '技术组', '商务组'] as const;

export const BID_HEADERS = [
  '采购项目名称',
  '采购项目编号',
  '分标名称',
  '分标编号',
  '分包名称',
  '投标人名称',
] as const;

export interface WarningFormItem {
  prjName: string;
  prjNo: string;
  labelNo: string;
  section: string;
  group: string;
  packs: string;
  suppliers: string[];
  flags: [boolean, boolean, boolean];
  values: [string, string, string];
  warnMatters: string;
  reasonLabels: string[];
  docxPath: string;
  docxFileName: string;
  date: string;
}

export interface WarningFormAuditResult {
  date: string;
  riskRowCount: number;
  groupRowCount: number;
  bidRowCount: number;
  formCount: number;
  forms: WarningFormItem[];
  warnings: string[];
  readmeText: string;
}

interface ParsedTable {
  rows: Array<Record<string, string>>;
  headerRow: number;
}

function getCellDisplay(ws: XLSX.WorkSheet, sheetName: string, r: number, c: number): string {
  const cell = ws[XLSX.utils.encode_cell({ r, c })] as XLSX.CellObject | undefined;
  if (!cell || cell.t === 'z') return '';
  if (cell.t === 'e') {
    const fErr = getCellFormulaError(cell);
    throw new Error(fErr || `${sheetName}!${XLSX.utils.encode_cell({ r, c })} 为错误单元格`);
  }
  if (cell.w !== undefined && cell.w !== null) return String(cell.w).trim();
  if (cell.v === undefined || cell.v === null) return '';
  return String(cell.v).trim();
}

/**
 * Exact port of WorkbookTable.readSheet (including vertical-only header merge inheritance).
 */
function readWorkbookTable(
  buffer: Buffer,
  fileName: string,
  columns: Record<string, string[]>
): ParsedTable {
  let wb: XLSX.WorkBook;
  try {
    wb = XLSX.read(buffer, { type: 'buffer', cellDates: false, sheetStubs: true, cellNF: true, cellStyles: true });
  } catch (err) {
    throw new Error(`${fileName} 无法读取：${(err as Error).message}`);
  }
  evaluateWorkbookFormulas(wb);
  if (wb.SheetNames.length === 0) {
    throw new Error(`工作簿没有工作表：${fileName}`);
  }
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const ref = ws['!ref'];
  if (!ref) {
    throw new Error(`${sheetName} 为空工作表`);
  }
  const range = XLSX.utils.decode_range(ref);
  const merges = ws['!merges'] || [];

  const getHeaderCells = (rowIndex: number): Array<{ col: number; text: string }> => {
    const map = new Map<number, string>();
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r: rowIndex, c });
      if (ws[addr] !== undefined && ws[addr].t !== 'z') {
        map.set(c, getCellDisplay(ws, sheetName, rowIndex, c));
      }
    }
    // Only vertical header merges inherit their anchor (exact WorkbookTable.java rule)
    for (const m of merges) {
      if (m.s.c === m.e.c && m.s.r <= rowIndex && m.e.r >= rowIndex) {
        const anchorAddr = XLSX.utils.encode_cell({ r: m.s.r, c: m.s.c });
        if (ws[anchorAddr] !== undefined && ws[anchorAddr].t !== 'z') {
          map.set(m.s.c, getCellDisplay(ws, sheetName, m.s.r, m.s.c));
        }
      }
    }
    return Array.from(map.entries())
      .sort((a, b) => a[0] - b[0])
      .map(([col, text]) => ({ col, text }));
  };

  const colEntries = Object.entries(columns);
  let indexes: Map<string, number> | null = null;
  let headerRow = -1;
  let bestMatched = new Set<string>();
  let bestRow = -1;

  for (let r = range.s.r; r <= range.e.r; r++) {
    const hCells = getHeaderCells(r);
    if (hCells.length === 0) continue;
    const found = new Map<string, number>();
    for (const [key, aliases] of colEntries) {
      for (const hc of hCells) {
        if (aliases.includes(hc.text)) {
          if (found.has(key)) {
            throw new Error(`表头重复：${key}，行${r + 1}`);
          }
          found.set(key, hc.col);
        }
      }
    }
    if (found.size > bestMatched.size) {
      bestMatched = new Set(found.keys());
      bestRow = r + 1;
    }
    if (found.size === colEntries.length) {
      indexes = found;
      headerRow = r;
      break;
    }
  }

  if (!indexes) {
    const missing = colEntries.map(([k]) => k).filter((k) => !bestMatched.has(k));
    throw new Error(
      `${sheetName} 缺少必要表头：${missing.join('、')}${bestRow > 0 ? `（最接近的表头在第${bestRow}行）` : ''}`
    );
  }

  // Data starts below every selected vertical header merge
  const selectedCols = new Set(indexes.values());
  let dataStart = headerRow + 1;
  for (const m of merges) {
    if (m.s.c === m.e.c && selectedCols.has(m.s.c) && m.s.r <= headerRow && m.e.r >= headerRow) {
      dataStart = Math.max(dataStart, m.e.r + 1);
    }
  }

  const rows: Array<Record<string, string>> = [];
  for (let r = dataStart; r <= range.e.r; r++) {
    const vals: Record<string, string> = {};
    let empty = true;
    for (const [key, colIdx] of indexes.entries()) {
      const v = getCellDisplay(ws, sheetName, r, colIdx);
      vals[key] = v;
      if (v !== '') empty = false;
    }
    if (!empty) {
      vals.__row = String(r + 1);
      rows.push(vals);
    }
  }

  return { rows, headerRow: headerRow + 1 };
}

function requireFields(row: Record<string, string>, fields: readonly string[]) {
  for (const f of fields) {
    if (!String(row[f] || '').trim()) {
      throw new Error(`第${row.__row}行缺少${f}`);
    }
  }
}

function comparePacks(a: string, b: string): number {
  const ma = a.replace(/^包/, '');
  const mb = b.replace(/^包/, '');
  if (/^\d+$/.test(ma) && /^\d+$/.test(mb)) {
    const na = BigInt(ma);
    const nb = BigInt(mb);
    if (na !== nb) return na < nb ? -1 : 1;
  }
  return a < b ? -1 : a > b ? 1 : 0;
}

function safePathSegment(value: string): string {
  return value.replace(/[\\/:*?"<>|]/g, '_').replace(/\.\./g, '_');
}

function formatTodayChinese(): string {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  return `${y}年${m}月${d}日`;
}

/**
 * Exact port of WarningFormTool.java execute() logic.
 */
export function runWarningFormAudit(params: {
  riskFile: { originalName: string; buffer: Buffer };
  groupFile: { originalName: string; buffer: Buffer };
  bidFile: { originalName: string; buffer: Buffer };
  date?: string;
}): WarningFormAuditResult {
  const riskColumns: Record<string, string[]> = {};
  for (const h of RISK_HEADERS) {
    if (h === '分包名称') riskColumns[h] = [h, '包名', '包名称', '分包编号'];
    else if (h === '投标人') riskColumns[h] = [h, '投标人名称', '供应商名称', '投标厂商名称'];
    else riskColumns[h] = [h];
  }

  const groupColumns: Record<string, string[]> = {};
  for (const h of GROUP_HEADERS) {
    if (h === '分包名称') groupColumns[h] = [h, '包名', '包名称'];
    else if (h === '技术组') groupColumns[h] = [h, '技术评审组'];
    else if (h === '商务组') groupColumns[h] = [h, '商务评审组'];
    else groupColumns[h] = [h];
  }

  const bidColumns: Record<string, string[]> = {};
  for (const h of BID_HEADERS) {
    if (h === '分包名称') bidColumns[h] = [h, '包名', '包名称'];
    else bidColumns[h] = [h];
  }

  const risks = readWorkbookTable(params.riskFile.buffer, params.riskFile.originalName, riskColumns);
  const grouping = readWorkbookTable(params.groupFile.buffer, params.groupFile.originalName, groupColumns);
  const bids = readWorkbookTable(params.bidFile.buffer, params.bidFile.originalName, bidColumns);

  const packKey = (section: string, pack: string) => `${section}\u0000${pack}`;

  const projects = new Map<string, { name: string; id: string; sectionId: string }>();
  const participants = new Set<string>();
  const groups = new Map<string, string[]>();

  for (const row of bids.rows) {
    requireFields(row, BID_HEADERS);
    const pk = packKey(row['分标名称'], row['分包名称']);
    const proj = {
      name: row['采购项目名称'],
      id: row['采购项目编号'],
      sectionId: row['分标编号'],
    };
    const old = projects.get(pk);
    if (old) {
      if (old.name !== proj.name || old.id !== proj.id || old.sectionId !== proj.sectionId) {
        throw new Error(`同标包项目或编号冲突：${row['分标名称']} / ${row['分包名称']}`);
      }
    } else {
      projects.set(pk, proj);
    }
    participants.add(`${pk}\u0000${row['投标人名称']}`);
  }

  for (const row of grouping.rows) {
    requireFields(row, GROUP_HEADERS);
    const pk = packKey(row['分标名称'], row['分包名称']);
    const grpVal = Array.from(new Set([row['技术组'], row['商务组']]));
    const old = groups.get(pk);
    if (old) {
      if (old.length !== grpVal.length || old.some((v, idx) => v !== grpVal[idx])) {
        throw new Error(`同标包评审组冲突：${row['分标名称']} / ${row['分包名称']}`);
      }
    } else {
      groups.set(pk, grpVal);
    }
  }

  const placeholders = new Set(['/', '／', '-', '—']);
  const byPack = new Map<
    string,
    {
      section: string;
      pack: string;
      list: Array<{ supplier: string; values: [string, string, string]; flags: [boolean, boolean, boolean] }>;
    }
  >();

  for (const row of risks.rows) {
    requireFields(row, ['分标名称', '分包名称', '投标人']);
    const flags: boolean[] = [];
    const values: string[] = [];
    for (let i = 3; i < 6; i++) {
      const flag = row[RISK_HEADERS[i + 3]];
      if (flag !== '是' && flag !== '否') {
        throw new Error(`风险表第${row.__row}行：是否相同须填写是或否`);
      }
      const enabled = flag === '是';
      const val = row[RISK_HEADERS[i]];
      if (enabled && (!val || placeholders.has(val))) {
        throw new Error(`风险表第${row.__row}行已标记相同但缺少${RISK_HEADERS[i]}`);
      }
      flags.push(enabled);
      values.push(val);
    }
    if (!flags.includes(true)) continue;

    const section = row['分标名称'];
    const pack = row['分包名称'];
    const supplier = row['投标人'];
    const pk = packKey(section, pack);
    if (!participants.has(`${pk}\u0000${supplier}`)) {
      throw new Error(`开标表未找到风险记录对应的投标人：${section} / ${pack} / ${supplier}`);
    }
    if (!groups.has(pk)) {
      throw new Error(`缺少标包分组：${section} / ${pack}`);
    }
    let entry = byPack.get(pk);
    if (!entry) {
      entry = { section, pack, list: [] };
      byPack.set(pk, entry);
    }
    entry.list.push({
      supplier,
      values: [values[0], values[1], values[2]],
      flags: [flags[0], flags[1], flags[2]],
    });
  }

  interface FormAgg {
    project: { name: string; id: string; sectionId: string };
    section: string;
    group: string;
    suppliers: string[];
    flags: [boolean, boolean, boolean];
    values: [string, string, string];
    packs: Set<string>;
  }

  const formMap = new Map<string, FormAgg>();

  for (const [pk, packEntry] of byPack.entries()) {
    const cohorts = new Map<
      string,
      { flags: [boolean, boolean, boolean]; values: [string, string, string]; suppliers: Set<string> }
    >();
    for (const risk of packEntry.list) {
      const matchedValues: [string, string, string] = [
        risk.flags[0] ? risk.values[0] : '',
        risk.flags[1] ? risk.values[1] : '',
        risk.flags[2] ? risk.values[2] : '',
      ];
      const sigKey = JSON.stringify([risk.flags, matchedValues]);
      let cohort = cohorts.get(sigKey);
      if (!cohort) {
        cohort = { flags: risk.flags, values: matchedValues, suppliers: new Set() };
        cohorts.set(sigKey, cohort);
      }
      cohort.suppliers.add(risk.supplier);
    }

    for (const cohort of cohorts.values()) {
      if (cohort.suppliers.size < 2) continue;
      // Java TreeSet<String> sorts by UTF-16 code unit order
      const sortedSuppliers = Array.from(cohort.suppliers).sort((a, b) => (a < b ? -1 : a > b ? 1 : 0));
      const proj = projects.get(pk)!;
      for (const grp of groups.get(pk)!) {
        const fKey = JSON.stringify([proj, packEntry.section, grp, sortedSuppliers, cohort.flags, cohort.values]);
        let agg = formMap.get(fKey);
        if (!agg) {
          agg = {
            project: proj,
            section: packEntry.section,
            group: grp,
            suppliers: sortedSuppliers,
            flags: cohort.flags,
            values: cohort.values,
            packs: new Set(),
          };
          formMap.set(fKey, agg);
        }
        agg.packs.add(packEntry.pack);
      }
    }
  }

  const documentDate = String(params.date ?? '').trim() || formatTodayChinese();
  const warnings: string[] = [];
  if (formMap.size === 0) {
    warnings.push('未找到标记组合一致且具体信息相同的不同投标人；不代表已经人工确认无异常');
  }

  const shortReasonLabels = ['注册地址相同', '联系方式相同', '邮箱相同'] as const;
  const usedNames = new Set<string>();
  const forms: WarningFormItem[] = [];

  for (const agg of formMap.values()) {
    const sortedPacks = Array.from(agg.packs).sort(comparePacks);
    const packsStr = sortedPacks.join('、');

    const reasonLabels: string[] = [];
    const reasonDetails: string[] = [];
    for (let f = 0; f < 3; f++) {
      if (agg.flags[f]) {
        reasonLabels.push(shortReasonLabels[f]);
        reasonDetails.push(`${RISK_HEADERS[f + 3]}相同（${agg.values[f]}）`);
      }
    }

    const groupSuffix = agg.group.endsWith('组') ? agg.group : `${agg.group}组`;
    const baseName = safePathSegment(
      `预警事项确认单(${reasonLabels.join('、')})-${agg.section}-${packsStr}-${agg.suppliers.join('、')}-${agg.project.id}${groupSuffix}`
    );
    const folder = safePathSegment(agg.group);
    let candidate = `${folder}/${baseName}.docx`;
    let dup = 1;
    while (usedNames.has(candidate)) {
      dup++;
      candidate = `${folder}/${baseName}(${dup}).docx`;
    }
    usedNames.add(candidate);

    forms.push({
      prjName: agg.project.name,
      prjNo: agg.project.id,
      labelNo: agg.project.sectionId,
      section: agg.section,
      group: agg.group,
      packs: packsStr,
      suppliers: agg.suppliers,
      flags: agg.flags,
      values: agg.values,
      warnMatters: reasonDetails.join('；'),
      reasonLabels,
      docxPath: candidate,
      docxFileName: path.basename(candidate),
      date: documentDate,
    });
  }

  const readmeText = `共生成${forms.length}份确认单。按风险表中相同的预警标记组合及对应值匹配同标包的不同投标人，同一事项的企业合并为一单，按技术/商务组分别生成；仅企业名单完全一致时归并包号。使用原版Word模板及固定复核文字，签字前请专家核对实际情况。\n`;

  return {
    date: documentDate,
    riskRowCount: risks.rows.length,
    groupRowCount: grouping.rows.length,
    bidRowCount: bids.rows.length,
    formCount: forms.length,
    forms,
    warnings,
    readmeText,
  };
}

function escapeXml(str: string): string {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function resolveTemplatePath(): string {
  const candidates = [
    path.resolve(process.cwd(), 'templates', '预警事项确认单.docx'),
    path.resolve(process.cwd(), 'server', 'templates', '预警事项确认单.docx'),
  ];
  for (const p of candidates) {
    if (fs.existsSync(p)) return p;
  }
  throw new Error('缺少原版预警事项确认单 Word 模板文件（templates/预警事项确认单.docx）');
}

/**
 * Render a single 预警事项确认单.docx from the bundled original Word template,
 * replacing only <w:t> text nodes and preserving all runs, fonts, spacing, and page layout.
 */
export function renderWarningFormDocx(item: WarningFormItem): Buffer {
  const tplBuf = fs.readFileSync(resolveTemplatePath());
  const zip = new PizZip(tplBuf);
  const docFile = zip.file('word/document.xml');
  if (!docFile) throw new Error('原版预警事项确认单模板损坏（缺少 word/document.xml）');
  const xml = docFile.asText();

  const values: Record<string, string> = {
    prjName: item.prjName,
    prjNo: item.prjNo,
    group: item.group,
    lableName: item.section,
    labelNo: item.labelNo,
    packageName: item.packs,
    corpNum: String(item.suppliers.length),
    corporateNames: item.suppliers.join('、'),
    warnMatters: item.warnMatters,
    date: item.date,
  };

  const replacedXml = xml.replace(/(<w:t(?:\s[^>]*)?>)([\s\S]*?)(<\/w:t>)/g, (_m, open, inner, close) => {
    let text = String(inner);
    if (item.suppliers.length !== 2) {
      text = text.replace(/该两家应答人/g, `该${item.suppliers.length}家应答人`);
    }
    text = text.replace(/\{\{([A-Za-z]+)\}\}/g, (_ph, key: string) => {
      if (!(key in values)) {
        throw new Error(`未知模板字段：${key}`);
      }
      return escapeXml(values[key]);
    });
    return `${open}${text}${close}`;
  });

  zip.file('word/document.xml', replacedXml);
  return zip.generate({ type: 'nodebuffer' });
}

/**
 * Package all rendered confirmation forms + 使用说明.txt into a single ZIP buffer.
 */
export function buildWarningFormsZipBuffer(result: WarningFormAuditResult): Buffer {
  const zip = new PizZip();
  for (const item of result.forms) {
    const docxBuf = renderWarningFormDocx(item);
    zip.file(item.docxPath, docxBuf);
  }
  zip.file('使用说明.txt', result.readmeText);
  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
}
