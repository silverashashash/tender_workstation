import fs from 'node:fs';
import path from 'node:path';
import PizZip from 'pizzip';
import XLSX from 'xlsx';

export type MinutesProcurementType = '公开' | '非公开';

export interface MinutesRejectionRow {
  seq: number;
  lot_name: string;
  package_no: string;
  supplier_name: string;
  reason: string;
  bid_response: string;
  review_type: string;
}

export interface InitialMinutesConfig {
  procurementType: MinutesProcurementType;
  batchName: string;
  meetingTime: string;
  groupName: string;
  clarificationNote: string;
  reviewTypeFilter?: 'all' | '商务' | '技术';
}

export interface InitialMinutesResult {
  fileName: string;
  sheetName: string;
  headerRow: number;
  config: InitialMinutesConfig;
  totalRows: number;
  hasReviewTypeCol: boolean;
  rows: MinutesRejectionRow[];
  previewParagraphs: string[];
}

const COL_ALIASES: Record<string, string[]> = {
  lot_name: ['分标名称', '分标分包名称', '标段名称', '分标'],
  package_no: ['包号', '包号/子包号', '分包名称', '包名', '包名称', '分包编号'],
  supplier_name: ['投标人', '投标人名称', '响应人', '响应人名称', '应答人', '应答人名称', '供应商', '供应商名称', '投标厂商名称'],
  reason: ['未进入详评原因', '否决原因', '否决模板', '未通过原因', '问题描述', '原因'],
  bid_response: ['投标文件响应', '应答文件响应', '响应文件响应', '投标文件响应情况', '应答文件响应情况', '响应情况', '说明'],
  review_type: ['商务类/技术类', '评审类别', '审查类型', '原因类别', '类别'],
};

function getCellText(ws: XLSX.WorkSheet, r: number, c: number): string {
  const cell = ws[XLSX.utils.encode_cell({ r, c })] as XLSX.CellObject | undefined;
  if (!cell) return '';
  if (cell.w !== undefined && cell.w !== null) return String(cell.w).trim();
  if (cell.v === undefined || cell.v === null) return '';
  return String(cell.v).trim();
}

function normalizeHeader(raw: string): string {
  return String(raw ?? '')
    .replace(/[\s\u3000\u00A0]+/g, '')
    .replace(/（/g, '(')
    .replace(/）/g, ')')
    .trim();
}

/**
 * Dynamically locate header row (scanning rows 1..25 across non-cover sheets first)
 * so both 《初评情况汇报模板》 (row 2 header) and 《未进入详评情况统计表》 (row 4 header)
 * or any custom group Excel table parse cleanly.
 */
export function parseMinutesExcel(buffer: Buffer, fileName: string): {
  sheetName: string;
  headerRow: number;
  rows: MinutesRejectionRow[];
  hasReviewTypeCol: boolean;
  detectedBatchName: string;
  detectedGroupName: string;
} {
  let wb: XLSX.WorkBook;
  try {
    wb = XLSX.read(buffer, { type: 'buffer', cellDates: false });
  } catch (err) {
    throw new Error(`无法读取 Excel 文件「${fileName}」：${(err as Error).message}`);
  }
  if (wb.SheetNames.length === 0) {
    throw new Error(`工作簿「${fileName}」没有工作表`);
  }

  const isCoverName = (n: string) => /封面\s*[12一二]/i.test(n);
  const orderedSheets = [
    ...wb.SheetNames.filter((n) => !isCoverName(n)),
    ...wb.SheetNames.filter((n) => isCoverName(n)),
  ];

  const requiredKeys = ['lot_name', 'package_no', 'supplier_name', 'reason'] as const;
  let matchedSheet = '';
  let matchedHeaderRow = -1;
  let matchedIndexes: Record<string, number> = {};
  let bestCount = 0;
  let bestInfo = '';

  for (const sheetName of orderedSheets) {
    const ws = wb.Sheets[sheetName];
    const ref = ws['!ref'];
    if (!ref) continue;
    const range = XLSX.utils.decode_range(ref);
    const maxScanRow = Math.min(range.e.r, 25);

    for (let r = range.s.r; r <= maxScanRow; r++) {
      const rowCells: { c: number; text: string }[] = [];
      for (let c = range.s.c; c <= Math.min(range.e.c, 30); c++) {
        const text = getCellText(ws, r, c);
        if (text) rowCells.push({ c, text });
      }
      if (rowCells.length === 0) continue;

      const idxMap: Record<string, number> = {};
      for (const [key, aliases] of Object.entries(COL_ALIASES)) {
        // Match aliases in priority order so exact primary headers (e.g. 未进入详评原因)
        // always take precedence over fallback aliases (e.g. 问题描述)
        for (const alias of aliases) {
          const hit = rowCells.find((cell) => normalizeHeader(cell.text) === normalizeHeader(alias));
          if (hit) {
            idxMap[key] = hit.c;
            break;
          }
        }
      }
      const reqHit = requiredKeys.filter((k) => idxMap[k] !== undefined).length;
      if (reqHit > bestCount) {
        bestCount = reqHit;
        bestInfo = `工作表「${sheetName}」第 ${r + 1} 行`;
      }
      if (reqHit === requiredKeys.length) {
        matchedSheet = sheetName;
        matchedHeaderRow = r;
        matchedIndexes = idxMap;
        break;
      }
    }
    if (matchedSheet) break;
  }

  if (!matchedSheet) {
    throw new Error(
      `未在「${fileName}」中找到包含必要表头（分标名称、包号、投标人/响应人、未进入详评原因）的数据表${
        bestInfo ? `（最接近位置：${bestInfo}）` : ''
      }`
    );
  }

  const ws = wb.Sheets[matchedSheet];
  const range = XLSX.utils.decode_range(ws['!ref']!);

  // Try to auto-detect batch name & group name from title rows above header
  let detectedBatchName = '';
  let detectedGroupName = '';
  for (let r = 0; r < matchedHeaderRow; r++) {
    const line = getCellText(ws, r, 0);
    if (!line) continue;
    if (!detectedGroupName) {
      const gm = line.match(/((?:商务|技术)[一二三四五六七八九十\d]*组)/);
      if (gm) detectedGroupName = gm[1];
    }
    if (!detectedBatchName && /国网|采购|招标/.test(line) && !line.startsWith('分标编号')) {
      detectedBatchName = line
        .replace(/\s*(?:商务|技术)[一二三四五六七八九十\d]*组.*$/, '')
        .replace(/初评情况汇报.*$/, '')
        .trim();
    }
  }

  const rows: MinutesRejectionRow[] = [];
  let seq = 0;
  for (let r = matchedHeaderRow + 1; r <= range.e.r; r++) {
    const lotName = getCellText(ws, r, matchedIndexes.lot_name);
    const pkgNo = getCellText(ws, r, matchedIndexes.package_no);
    const supplier = getCellText(ws, r, matchedIndexes.supplier_name);
    const reason = getCellText(ws, r, matchedIndexes.reason);
    const bidResponse =
      matchedIndexes.bid_response !== undefined ? getCellText(ws, r, matchedIndexes.bid_response) : '';
    const reviewType =
      matchedIndexes.review_type !== undefined ? getCellText(ws, r, matchedIndexes.review_type) : '';

    if (!lotName && !pkgNo && !supplier && !reason && !bidResponse) continue;
    // Skip repeated header rows
    if (normalizeHeader(supplier) === '投标人' || normalizeHeader(supplier) === '投标人名称') continue;

    seq++;
    rows.push({
      seq,
      lot_name: lotName,
      package_no: pkgNo,
      supplier_name: supplier,
      reason,
      bid_response: bidResponse,
      review_type: reviewType,
    });
  }

  return {
    sheetName: matchedSheet,
    headerRow: matchedHeaderRow + 1,
    rows,
    hasReviewTypeCol: matchedIndexes.review_type !== undefined,
    detectedBatchName,
    detectedGroupName,
  };
}

export function buildMinutesResult(
  parsed: ReturnType<typeof parseMinutesExcel>,
  fileName: string,
  config: InitialMinutesConfig
): InitialMinutesResult {
  const filter = config.reviewTypeFilter || 'all';
  const filtered = parsed.rows
    .filter((r) => {
      if (filter === 'all' || !parsed.hasReviewTypeCol) return true;
      return r.review_type.includes(filter);
    })
    .map((r, idx) => ({ ...r, seq: idx + 1 }));

  const isPublic = config.procurementType === '公开';
  const committee = isPublic ? '评标委员会' : '评审委员会';
  const defaultClarif = isPublic ? '本次评标无澄清事项。' : '本次评审无澄清事项。';
  const clarifText = config.clarificationNote.trim() || defaultClarif;
  const basisText = isPublic
    ? '根据招标文件第三章评标办法前附表之二否决事项规定，不满足本次招标文件要求，其投标文件作否决处理。'
    : '根据采购文件第二章评审办法否决事项规定，不满足本次采购文件要求，其应答文件作否决处理。';
  const summaryText = isPublic
    ? '与会人员发表意见，评标委员会主任作总结发言。经审核，批准同意评标组提交的初评情况报告。'
    : '与会人员发表意见，评审委员会主任作总结发言。经审核，批准同意评审组提交的初评情况报告。';

  const groupLabel = config.groupName.trim() || '商务组';
  const timeLabel = config.meetingTime.trim() || formatDefaultMeetingTime();
  const batchLabel = config.batchName.trim() || '【待填写批次名称】';

  const previewParagraphs = [
    `${timeLabel}，${committee}主持召开${batchLabel}${groupLabel}初评会议。${committee}主任、法律保障组、监督组等有关人员参加了会议。现将会议主要内容纪要如下：`,
    '一、会议情况',
    `会议听取了${groupLabel}的初评情况报告，对${groupLabel}提出的澄清、否决情况进行审核。（${clarifText}）基本情况及否决依据如下：`,
    basisText,
    `[否决明细表格：共 ${filtered.length} 条记录]`,
    '二、会议总结',
    summaryText,
    '参会单位、人员：',
  ];

  return {
    fileName,
    sheetName: parsed.sheetName,
    headerRow: parsed.headerRow,
    config: {
      ...config,
      batchName: batchLabel,
      meetingTime: timeLabel,
      groupName: groupLabel,
      clarificationNote: clarifText,
      reviewTypeFilter: filter,
    },
    totalRows: filtered.length,
    hasReviewTypeCol: parsed.hasReviewTypeCol,
    rows: filtered,
    previewParagraphs,
  };
}

export function formatDefaultMeetingTime(): string {
  const now = new Date();
  const y = now.getFullYear();
  const m = now.getMonth() + 1;
  const d = now.getDate();
  const hh = String(now.getHours()).padStart(2, '0');
  const mm = String(now.getMinutes()).padStart(2, '0');
  return `${y}年${m}月${d}日${hh}时${mm}分`;
}

function escapeXml(str: string): string {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function resolveMinutesTemplatePath(type: MinutesProcurementType): string {
  const file = type === '公开' ? '初评会议纪要_公开.docx' : '初评会议纪要_非公开.docx';
  const candidates = [
    path.resolve(process.cwd(), 'templates', file),
    path.resolve(process.cwd(), 'server', 'templates', file),
  ];
  for (const p of candidates) {
    if (fs.existsSync(p)) return p;
  }
  throw new Error(`缺少原版初评会议纪要模板文件（templates/${file}）`);
}

function buildStyledTableXml(rows: MinutesRejectionRow[], isPublic: boolean): string {
  // 6 columns totaling ~8500 dxa (matches A4 with standard margins from original template)
  const colWidths = [680, 1300, 920, 1750, 1450, 2400];
  const headers = isPublic
    ? ['序号', '分标名称', '包号', '投标人', '未进入详评原因', '投标文件响应']
    : ['序号', '分标名称', '包号', '响应人', '未进入详评原因', '应答文件响应'];

  const borderXml =
    '<w:tcBorders>' +
    '<w:top w:val="single" w:color="000000" w:sz="4" w:space="0"/>' +
    '<w:left w:val="single" w:color="000000" w:sz="4" w:space="0"/>' +
    '<w:bottom w:val="single" w:color="000000" w:sz="4" w:space="0"/>' +
    '<w:right w:val="single" w:color="000000" w:sz="4" w:space="0"/>' +
    '</w:tcBorders>';

  const makeCell = (text: string, width: number, isHeader: boolean, align: 'center' | 'left') => {
    const sz = isHeader ? 24 : 21; // 12pt header, 10.5pt body (五号宋体)
    const boldXml = isHeader ? '<w:b/><w:bCs/>' : '';
    const lines = String(text ?? '').split(/\r?\n/);
    const runsXml = lines
      .map(
        (ln, i) =>
          (i > 0 ? '<w:br/>' : '') +
          `<w:t xml:space="preserve">${escapeXml(ln)}</w:t>`
      )
      .join('');
    return (
      `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${borderXml}` +
      `<w:shd w:val="clear" w:color="auto" w:fill="auto"/><w:vAlign w:val="center"/></w:tcPr>` +
      `<w:p><w:pPr><w:jc w:val="${align}"/><w:spacing w:before="40" w:after="40" w:line="280" w:lineRule="auto"/></w:pPr>` +
      `<w:r><w:rPr><w:rFonts w:hint="eastAsia" w:ascii="SimSun" w:hAnsi="SimSun" w:eastAsia="SimSun" w:cs="SimSun"/>${boldXml}<w:color w:val="000000"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr>` +
      `${runsXml}</w:r></w:p></w:tc>`
    );
  };

  const gridXml = `<w:tblGrid>${colWidths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;

  const headerRowXml =
    `<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>` +
    headers.map((h, i) => makeCell(h, colWidths[i], true, 'center')).join('') +
    `</w:tr>`;

  const dataRowsXml =
    rows.length > 0
      ? rows
          .map(
            (r) =>
              `<w:tr><w:trPr><w:cantSplit/></w:trPr>` +
              makeCell(String(r.seq), colWidths[0], false, 'center') +
              makeCell(r.lot_name, colWidths[1], false, 'center') +
              makeCell(r.package_no, colWidths[2], false, 'center') +
              makeCell(r.supplier_name, colWidths[3], false, 'left') +
              makeCell(r.reason, colWidths[4], false, 'center') +
              makeCell(r.bid_response, colWidths[5], false, 'left') +
              `</w:tr>`
          )
          .join('')
      : `<w:tr><w:trPr><w:cantSplit/></w:trPr>` +
        `<w:tc><w:tcPr><w:tcW w:w="8500" w:type="dxa"/><w:gridSpan w:val="6"/>${borderXml}<w:vAlign w:val="center"/></w:tcPr>` +
        `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="60" w:after="60"/></w:pPr>` +
        `<w:r><w:rPr><w:rFonts w:hint="eastAsia" w:ascii="SimSun" w:hAnsi="SimSun" w:eastAsia="SimSun" w:cs="SimSun"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>` +
        `<w:t>本次初评无未进入详评（否决）事项</w:t></w:r></w:p></w:tc></w:tr>`;

  return (
    `<w:tbl><w:tblPr><w:tblW w:w="8500" w:type="dxa"/><w:jc w:val="center"/>` +
    `<w:tblBorders>` +
    `<w:top w:val="single" w:color="000000" w:sz="4" w:space="0"/>` +
    `<w:left w:val="single" w:color="000000" w:sz="4" w:space="0"/>` +
    `<w:bottom w:val="single" w:color="000000" w:sz="4" w:space="0"/>` +
    `<w:right w:val="single" w:color="000000" w:sz="4" w:space="0"/>` +
    `<w:insideH w:val="single" w:color="000000" w:sz="4" w:space="0"/>` +
    `<w:insideV w:val="single" w:color="000000" w:sz="4" w:space="0"/>` +
    `</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>` +
    `${gridXml}${headerRowXml}${dataRowsXml}</w:tbl>`
  );
}

/**
 * Rebuild a paragraph's runs cleanly using the first run's rPr (with yellow highlight stripped)
 * while preserving the original paragraph's exact <w:pPr> (font, indent, line spacing, alignment).
 */
function replaceParagraphTextPreservingStyle(paraXml: string, newText: string): string {
  const pOpenMatch = paraXml.match(/^<w:p(?:\s[^>]*)?>/);
  const pOpen = pOpenMatch ? pOpenMatch[0] : '<w:p>';
  const pPrMatch = paraXml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
  const pPr = pPrMatch ? pPrMatch[0] : '';

  // Extract first <w:rPr> as template and strip yellow highlight/shading
  const rPrMatch = paraXml.match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
  let rPr = rPrMatch ? rPrMatch[0] : '';
  rPr = rPr
    .replace(/<w:highlight\b[^>]*\/>/gi, '')
    .replace(/<w:shd\b[^>]*w:fill="FFFF00"[^>]*\/>/gi, '');

  return `${pOpen}${pPr}<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(newText)}</w:t></w:r></w:p>`;
}

/**
 * Render the final 初评会议纪要 .docx from the corresponding original template (公开 or 非公开).
 */
export function renderInitialMinutesDocx(result: InitialMinutesResult): Buffer {
  const isPublic = result.config.procurementType === '公开';
  const tplBuf = fs.readFileSync(resolveMinutesTemplatePath(result.config.procurementType));
  const zip = new PizZip(tplBuf);
  const docFile = zip.file('word/document.xml');
  if (!docFile) throw new Error('初评会议纪要模板损坏（缺少 word/document.xml）');
  const xml = docFile.asText();

  const tableMatch = xml.match(/<w:tbl[\s\S]*?<\/w:tbl>/);
  if (!tableMatch) {
    throw new Error('初评会议纪要模板缺少否决明细表格');
  }

  const beforeTable = xml.slice(0, tableMatch.index!);
  const afterTable = xml.slice(tableMatch.index! + tableMatch[0].length);

  // Replace dynamic paragraphs in beforeTable (P1 = opening meeting info, P3 = group report & clarification)
  const committee = isPublic ? '评标委员会' : '评审委员会';
  const p1Text = `${result.config.meetingTime}，${committee}主持召开${result.config.batchName}${result.config.groupName}初评会议。${committee}主任、法律保障组、监督组等有关人员参加了会议。现将会议主要内容纪要如下：`;
  const p3Text = `会议听取了${result.config.groupName}的初评情况报告，对${result.config.groupName}提出的澄清、否决情况进行审核。（${result.config.clarificationNote}）基本情况及否决依据如下：`;

  const updatedBefore = beforeTable.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g, (paraXml) => {
    const plain = [...paraXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((m) => m[1]).join('');
    if (plain.includes('主持召开') && plain.includes('初评会议')) {
      return replaceParagraphTextPreservingStyle(paraXml, p1Text);
    }
    if (plain.includes('会议听取了') && plain.includes('初评情况报告')) {
      return replaceParagraphTextPreservingStyle(paraXml, p3Text);
    }
    return paraXml;
  });

  // Also ensure the closing summary paragraph aligns terminology (公开 -> 评标委员会/评标组, 非公开 -> 评审委员会/评审组)
  const summaryText = isPublic
    ? '与会人员发表意见，评标委员会主任作总结发言。经审核，批准同意评标组提交的初评情况报告。'
    : '与会人员发表意见，评审委员会主任作总结发言。经审核，批准同意评审组提交的初评情况报告。';

  const updatedAfter = afterTable.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g, (paraXml) => {
    const plain = [...paraXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((m) => m[1]).join('');
    if (plain.includes('与会人员发表意见') && plain.includes('初评情况报告')) {
      return replaceParagraphTextPreservingStyle(paraXml, summaryText);
    }
    return paraXml;
  });

  const newTableXml = buildStyledTableXml(result.rows, isPublic);
  zip.file('word/document.xml', `${updatedBefore}${newTableXml}${updatedAfter}`);
  return zip.generate({ type: 'nodebuffer' });
}
