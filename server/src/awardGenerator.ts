import fs from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';
import PizZip from 'pizzip';
import { calculateServiceFee, type ServiceFeeCategory } from './serviceFee.js';

export interface AwardPackageRow {
  seq: number;
  lot_code: string;
  lot_name: string;
  package_no: string;
  project_name: string;
  project_overview: string;
  req_app_no: string;
  budget_wan: number | null;
  max_price_wan: number | null;
  max_price_raw: string;
  declarant_info: string;
  material_desc: string;
  is_failed_lot: boolean;
  winner_name: string;
  winner_price_wan: number | null;
  winner_price_raw: string;
  winner_score: number | null;
  winner_contact: string;
  winner_mobile: string;
  winner_email: string;
  winner_address: string;
  service_fee_pre_yuan: number | null;
  service_fee_yuan: number | null;
  platform_fee_yuan: number | null;
  category: ServiceFeeCategory;
  discount_rate: number;
  remark: string;
}

export interface AwardBidderDetailRow {
  seq: number;
  batch_code: string;
  lot_name: string;
  package_no: string;
  max_price_display: string;
  department: string;
  declarant_info: string;
  project_name: string;
  material_desc: string;
  supplier_name: string;
  contact_name: string;
  mobile: string;
  email: string;
  is_compliant: '是' | '否';
  tech_score: string;
  biz_score: string;
  price_score: string;
  total_score: string;
  rank_no: string;
  base_price: string;
  eval_price: string;
  remark: string;
}

export interface AwardDeliverablePayload {
  batch: {
    id: string;
    name: string;
    code: string;
    project_unit: string;
    procurement_type: string;
    unit_nature: string;
    announcement_time: string;
    bid_opening_time: string;
    contact_phone: string;
    contact_email: string;
  };
  ownerUnit: string;
  category: ServiceFeeCategory;
  discountRate: number;
  discountLabel: string;
  packages: AwardPackageRow[];
  bidderDetails: AwardBidderDetailRow[];
  overrides: Record<string, string>;
}

function escapeXml(str: string): string {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function formatChineseDate(dateStr?: string): string {
  const s = String(dateStr || '').trim();
  const m = s.match(/^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})/);
  if (m) {
    return `${m[1]}年${m[2].padStart(2, '0')}月${m[3].padStart(2, '0')}日`;
  }
  const now = new Date();
  return `${now.getFullYear()}年${String(now.getMonth() + 1).padStart(2, '0')}月${String(now.getDate()).padStart(2, '0')}日`;
}

/**
 * 清理价格末尾多余的 0（例如 100.000 -> 100，100.000万元 -> 100万元，88.50% -> 88.5%）
 */
function formatCleanPrice(val: string | number | null | undefined, emptyFallback = '/'): string {
  if (val === null || val === undefined) return emptyFallback;
  if (typeof val === 'number') {
    if (!Number.isFinite(val)) return emptyFallback;
    return Number(val.toFixed(6)).toString();
  }
  const s = String(val).trim();
  if (!s || s === '/' || s === '-' || s === '—') return s || emptyFallback;
  const m = s.match(/^([+-]?\d+\.\d+)(.*)$/);
  if (m) {
    const cleanedNum = m[1].replace(/(\.\d*?[1-9])0+$/, '$1').replace(/\.0+$/, '');
    return `${cleanedNum}${m[2]}`;
  }
  return s;
}

const EXCEL_ALL_BORDERS: Partial<ExcelJS.Borders> = {
  top: { style: 'thin', color: { argb: 'FF000000' } },
  bottom: { style: 'thin', color: { argb: 'FF000000' } },
  left: { style: 'thin', color: { argb: 'FF000000' } },
  right: { style: 'thin', color: { argb: 'FF000000' } },
};

function styleHeaderRow(row: ExcelJS.Row) {
  row.height = 22;
  const colCount = row.cellCount;
  row.worksheet.pageSetup.printTitlesRow = `${row.number}:${row.number}`;
  if (colCount > 0) {
    row.worksheet.autoFilter = {
      from: { row: row.number, column: 1 },
      to: { row: row.number, column: colCount },
    };
  }
  for (let c = 1; c <= colCount; c++) {
    const cell = row.getCell(c);
    cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF000000' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = EXCEL_ALL_BORDERS;
  }
}

function styleDataRow(row: ExcelJS.Row, colCount?: number) {
  row.height = 18;
  const totalCols = colCount || row.worksheet.getRow(1).cellCount || row.cellCount;
  for (let c = 1; c <= totalCols; c++) {
    const cell = row.getCell(c);
    cell.font = { name: 'Calibri', size: 11, color: { argb: 'FF000000' } };
    cell.alignment = { vertical: 'middle', wrapText: true };
    cell.border = EXCEL_ALL_BORDERS;
  }
}

function buildDocxXmlTable(headers: string[], rows: string[][], colWidthsDxa: number[]): string {
  const gridCols = colWidthsDxa.map((w) => `<w:gridCol w:w="${w}"/>`).join('');

  const makeRow = (cells: string[], isHeader: boolean) => {
    const trPr = isHeader
      ? '<w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>'
      : '<w:trPr><w:cantSplit/></w:trPr>';
    const tcs = cells
      .map((val, idx) => {
        const w = colWidthsDxa[idx] || 1500;
        const bold = isHeader ? '<w:b/>' : '';
        return `<w:tc>
          <w:tcPr>
            <w:tcW w:w="${w}" w:type="dxa"/>
            <w:vAlign w:val="center"/>
          </w:tcPr>
          <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
            <w:r>
              <w:rPr><w:rFonts w:ascii="宋体" w:eastAsia="宋体" w:hAnsi="宋体"/>${bold}<w:sz w:val="18"/></w:rPr>
              <w:t xml:space="preserve">${escapeXml(val)}</w:t>
            </w:r>
          </w:p>
        </w:tc>`;
      })
      .join('');
    return `<w:tr>${trPr}${tcs}</w:tr>`;
  };

  return `<w:tbl>
    <w:tblPr>
      <w:tblW w:w="9000" w:type="dxa"/>
      <w:jc w:val="center"/>
      <w:tblBorders>
        <w:top w:val="single" w:sz="6" w:space="0" w:color="333333"/>
        <w:left w:val="single" w:sz="6" w:space="0" w:color="333333"/>
        <w:bottom w:val="single" w:sz="6" w:space="0" w:color="333333"/>
        <w:right w:val="single" w:sz="6" w:space="0" w:color="333333"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="666666"/>
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="666666"/>
      </w:tblBorders>
      <w:tblCellMar>
        <w:top w:w="25" w:type="dxa"/>
        <w:left w:w="80" w:type="dxa"/>
        <w:bottom w:w="25" w:type="dxa"/>
        <w:right w:w="80" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
    <w:tblGrid>${gridCols}</w:tblGrid>
    ${makeRow(headers, true)}
    ${rows.map((r) => makeRow(r, false)).join('')}
  </w:tbl>`;
}

function makePara(
  text: string,
  opts: {
    sizeHalfPt?: number;
    bold?: boolean;
    align?: 'left' | 'center' | 'right';
    firstLineChars?: number;
    spaceAfter?: number;
    spaceBefore?: number;
  } = {}
): string {
  const { sizeHalfPt = 24, bold = false, align = 'left', firstLineChars = 0, spaceAfter = 120, spaceBefore = 0 } = opts;
  const jc = align !== 'left' ? `<w:jc w:val="${align}"/>` : '';
  const ind = firstLineChars > 0 ? `<w:ind w:firstLineChars="${firstLineChars * 100}"/>` : '';
  const b = bold ? '<w:b/>' : '';
  return `<w:p>
    <w:pPr>${jc}${ind}<w:spacing w:before="${spaceBefore}" w:after="${spaceAfter}" w:line="360" w:lineRule="auto"/></w:pPr>
    <w:r>
      <w:rPr><w:rFonts w:ascii="仿宋_GB2312" w:eastAsia="仿宋_GB2312" w:hAnsi="仿宋_GB2312"/>${b}<w:sz w:val="${sizeHalfPt}"/></w:rPr>
      <w:t xml:space="preserve">${escapeXml(text)}</w:t>
    </w:r>
  </w:p>`;
}

function writeDocxFromBody(bodyXml: string, outputPath: string) {
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
  <w:body>
    ${bodyXml}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>`
  );

  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, zip.generate({ type: 'nodebuffer' }));
}

interface DocxTableColumnOptions {
  includeLotCode: boolean;
  includeLotName: boolean;
  includeProjectName: boolean;
  includePrice: boolean;
}

function resolveBoolOverride(val: string | undefined, defaultVal: boolean): boolean {
  if (val === 'true') return true;
  if (val === 'false') return false;
  return defaultVal;
}

/**
 * 动态构建公示/公告 Word 表格：固定「序号」「包号」「候选人/成交人」，
 * 自由勾选「分标编号」「分标名称」「项目名称」「价格」，自动按比例分配总宽 9000 dxa，居中无底色、紧凑行高、跨页重复表头。
 */
function buildAwardWordTable(
  packages: AwardPackageRow[],
  opts: DocxTableColumnOptions,
  labels: {
    winnerHeader: string;
    priceHeader: string;
    lotNameHeader?: string;
  }
): string {
  const cols: Array<{ header: string; weight: number; getValue: (p: AwardPackageRow, idx: number) => string }> = [
    { header: '序号', weight: 7, getValue: (_p, idx) => String(idx + 1) },
  ];
  if (opts.includeLotCode) {
    cols.push({ header: '分标编号', weight: 18, getValue: (p) => p.lot_code || '-' });
  }
  if (opts.includeLotName) {
    cols.push({ header: labels.lotNameHeader || '分标名称', weight: 13, getValue: (p) => p.lot_name || '-' });
  }
  cols.push({ header: '包号', weight: 8, getValue: (p) => p.package_no || '-' });
  if (opts.includeProjectName) {
    cols.push({ header: '项目名称', weight: 28, getValue: (p) => p.project_name || '-' });
  }
  cols.push({
    header: labels.winnerHeader,
    weight: 26,
    getValue: (p) => (p.is_failed_lot ? '流标' : p.winner_name || '待确定'),
  });
  if (opts.includePrice) {
    cols.push({
      header: labels.priceHeader,
      weight: 14,
      getValue: (p) => (p.is_failed_lot ? '/' : formatCleanPrice(p.winner_price_raw || p.winner_price_wan)),
    });
  }

  const totalWeight = cols.reduce((sum, c) => sum + c.weight, 0);
  const widths = cols.map((c) => Math.round((c.weight / totalWeight) * 9000));
  const diff = 9000 - widths.reduce((sum, w) => sum + w, 0);
  let maxIdx = 0;
  for (let i = 1; i < widths.length; i++) {
    if (widths[i] > widths[maxIdx]) maxIdx = i;
  }
  widths[maxIdx] += diff;

  const headers = cols.map((c) => c.header);
  const rows = packages.map((p, idx) => cols.map((c) => c.getValue(p, idx)));
  return buildDocxXmlTable(headers, rows, widths);
}

/**
 * 1. 公示文件 (.docx) — 区分公开招标与非公开招标正文，表格列支持自由勾选：
 *    可选项：分标编号、分标名称、项目名称、价格
 */
export function generatePreAwardPublicityDocx(payload: AwardDeliverablePayload, outputPath: string): void {
  const isPublic = payload.batch.procurement_type === '公开招标';
  const batchName = payload.overrides.batchName || payload.batch.name;
  const batchCode = payload.overrides.batchCode || payload.batch.code;
  const publicityDays = payload.overrides.publicityDays || (isPublic ? '3' : '1');
  const objectionEmail = payload.overrides.objectionEmail || payload.batch.contact_email || 'mszblw@ahdlzb.com';
  const objectionPhone = payload.overrides.objectionPhone || payload.batch.contact_phone || process.env.DEFAULT_OBJECTION_PHONE || '000-00000000';
  const docDate = formatChineseDate(payload.overrides.publicityDate || payload.overrides.docDate);

  const colOpts: DocxTableColumnOptions = {
    includeLotCode: resolveBoolOverride(payload.overrides.includeLotCode, !isPublic),
    includeLotName: resolveBoolOverride(payload.overrides.includeLotName, true),
    includeProjectName: resolveBoolOverride(payload.overrides.includeProjectName, isPublic),
    includePrice: resolveBoolOverride(payload.overrides.includePrice, isPublic),
  };

  if (isPublic) {
    const tableXml = buildAwardWordTable(payload.packages, colOpts, {
      lotNameHeader: '分标名称',
      winnerHeader: '推荐中标人',
      priceHeader: '评标总价(万元/%)',
    });

    const bodyParts = [
      makePara(batchName, { sizeHalfPt: 30, bold: true, align: 'center', spaceAfter: 40 }),
      makePara(`招标编号：${batchCode}`, { sizeHalfPt: 24, align: 'center', spaceAfter: 40 }),
      makePara('推荐的中标候选人公示', { sizeHalfPt: 30, bold: true, align: 'center', spaceAfter: 160 }),
      makePara('各相关投标人：', { sizeHalfPt: 24, spaceAfter: 100 }),
      makePara(
        `${batchName}（招标编号：${batchCode}）评标工作已经结束。依据《中华人民共和国招标投标法实施条例》、《招标公告和公示信息发布管理办法》对中标候选人公示的规定，现将评标委员会推荐的中标候选人予以公示，公示期${publicityDays}日。投标人或者其他利害关系人若对评标结果有异议的，请在中标候选人公示期间以书面形式（邮件）提出。`,
        { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 160 }
      ),
      tableXml,
      makePara(`投标人对以上结果如有异议，可以自公示之日起${publicityDays}日内，向招标人提出。`, {
        sizeHalfPt: 24,
        firstLineChars: 2,
        spaceBefore: 160,
        spaceAfter: 100,
      }),
      makePara(
        '根据招标投标法及相关法规规定，投标人和直接参与并且与招投标活动有着直接利害关系的当事人有异议的，有权依法进行异议，提出异议时应注意以下事项：',
        { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 80 }
      ),
      makePara('1.异议必须在中标结果公示结束前以书面形式邮件提出。', { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 60 }),
      makePara('2.应当提交异议书，并包括下列内容：', { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 60 }),
      makePara('（1）异议人的名称、地址、联系人及有效联系方式；', { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 60 }),
      makePara('（2）被异议人的名称；', { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 60 }),
      makePara('（3）异议事项的基本事实；', { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 60 }),
      makePara('（4）有效线索和相关证明材料。', { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 60 }),
      makePara(
        '3.异议人为法人的，异议书必须由其法定代表人或者授权代表签字并加盖公章，同时还需提交授权委托书；异议人为个人的，异议书必须由异议人本人签字，并附有效身份证明，由本人提交。',
        { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 60 }
      ),
      makePara('4.下列异议将不予接收：', { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 60 }),
      makePara('（1）在中标结果公示结束后提出的；', { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 60 }),
      makePara('（2）异议人不能证明是所异议招标投标活动的投标人和直接参与并且与招投标活动有着直接利害关系的当事人；', {
        sizeHalfPt: 24,
        firstLineChars: 2,
        spaceAfter: 60,
      }),
      makePara('（3）异议事项不具体，且未提供有效线索，难以查证的；', { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 60 }),
      makePara('（4）对异议事项已经答复，且异议人没有提出新的证据的。', { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 60 }),
      makePara('5. 异议人不得以异议为名排挤竞争对手，进行虚假、恶意异议，阻碍招标投标活动的正常进行。', {
        sizeHalfPt: 24,
        firstLineChars: 2,
        spaceAfter: 200,
      }),
      makePara(docDate, { sizeHalfPt: 24, align: 'right', spaceAfter: 80 }),
    ];

    writeDocxFromBody(bodyParts.join(''), outputPath);
    return;
  }

  const tableXml = buildAwardWordTable(payload.packages, colOpts, {
    lotNameHeader: '分标名称',
    winnerHeader: '推荐成交候选人',
    priceHeader: '最终报价(万元/%)',
  });

  const bodyParts = [
    makePara(`${batchName}预成交公示`, { sizeHalfPt: 30, bold: true, align: 'center', spaceAfter: 80 }),
    makePara(`（采购项目编号：${batchCode}）`, { sizeHalfPt: 24, align: 'center', spaceAfter: 160 }),
    makePara('各相关响应人：', { sizeHalfPt: 24, spaceAfter: 100 }),
    makePara(
      `${batchName}的评审工作已经结束，现将评审委员会推荐的成交候选人予以公示，公示期 ${publicityDays} 日，响应人或者其他利害关系人若对评审结果有异议的，请在成交候选人公示期间以书面形式提出。`,
      { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 160 }
    ),
    tableXml,
    makePara(`响应人对以上结果如有异议，可以自公示之日起 ${publicityDays} 日内，向采购人提出。`, {
      sizeHalfPt: 24,
      firstLineChars: 2,
      spaceBefore: 160,
      spaceAfter: 100,
    }),
    makePara(
      '根据招标投标法及相关法规规定，响应人和直接参与并且与采购活动有着直接利害关系的当事人有异议的，有权依法进行异议，提出异议时应注意以下事项：',
      { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 80 }
    ),
    makePara('1.异议必须在成交候选人公示结束前以书面形式提出。', { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 60 }),
    makePara('2.应当提交异议书，并包括下列内容：', { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 60 }),
    makePara('（1）异议人的名称、地址、联系人及有效联系方式；', { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 60 }),
    makePara('（2）被异议人的名称；', { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 60 }),
    makePara('（3）异议事项的基本事实；', { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 60 }),
    makePara('（4）有效线索和相关证明材料。', { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 60 }),
    makePara(
      '3.异议人为法人的，异议书必须由其法定代表人或者授权代表签字并加盖公章，同时还需提交授权委托书；异议人为个人的，异议书必须由异议人本人签字，并附有效身份证明，由本人提交。',
      { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 60 }
    ),
    makePara('4.下列异议将不予接收：', { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 60 }),
    makePara('（1）在成交候选人公示结束后提出的；', { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 60 }),
    makePara('（2）异议人不能证明是所异议采购活动的响应人和直接参与并且与采购活动有着直接利害关系的当事人；', {
      sizeHalfPt: 24,
      firstLineChars: 2,
      spaceAfter: 60,
    }),
    makePara('（3）异议事项不具体，且未提供有效线索，难以查证的；', { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 60 }),
    makePara('（4）对异议事项已经答复，且异议人没有提出新的证据的。', { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 60 }),
    makePara('5.异议人不得以异议为名排挤竞争对手，进行虚假、恶意异议，阻碍采购活动的正常进行。', {
      sizeHalfPt: 24,
      firstLineChars: 2,
      spaceAfter: 60,
    }),
    makePara(`6.评审结果异议联系方式：邮箱：${objectionEmail}`, { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 60 }),
    makePara(`联系电话：${objectionPhone}`, { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 200 }),
    makePara(docDate, { sizeHalfPt: 24, align: 'right', spaceAfter: 80 }),
  ];

  writeDocxFromBody(bodyParts.join(''), outputPath);
}

/**
 * 2. 结果公告 (.docx) — 区分公开招标与非公开招标正文，表格列支持自由勾选：
 *    可选项：分标编号、分标名称、项目名称、价格
 */
export function generateAwardResultNoticeDocx(payload: AwardDeliverablePayload, outputPath: string): void {
  const isPublic = payload.batch.procurement_type === '公开招标';
  const batchName = payload.overrides.batchName || payload.batch.name;
  const batchCode = payload.overrides.batchCode || payload.batch.code;
  const ownerUnit = payload.overrides.ownerUnit || payload.ownerUnit || payload.batch.project_unit;
  const agencyName = payload.overrides.agencyName || process.env.DEFAULT_AGENCY_NAME || 'XX工程咨询有限公司招标代理机构';
  const docDate = formatChineseDate(payload.overrides.announcementDate || payload.overrides.docDate);

  const colOpts: DocxTableColumnOptions = {
    includeLotCode: resolveBoolOverride(payload.overrides.includeLotCode, isPublic),
    includeLotName: resolveBoolOverride(payload.overrides.includeLotName, !isPublic),
    includeProjectName: resolveBoolOverride(payload.overrides.includeProjectName, !isPublic),
    includePrice: resolveBoolOverride(payload.overrides.includePrice, !isPublic),
  };

  if (isPublic) {
    const tableXml = buildAwardWordTable(payload.packages, colOpts, {
      lotNameHeader: '分标名称',
      winnerHeader: '中标候选人',
      priceHeader: '中标金额(万元/%)',
    });

    const bodyParts = [
      makePara(batchName, { sizeHalfPt: 30, bold: true, align: 'center', spaceAfter: 40 }),
      makePara('中标公告', { sizeHalfPt: 30, bold: true, align: 'center', spaceAfter: 80 }),
      makePara(`（招标编号：${batchCode}）`, { sizeHalfPt: 24, align: 'center', spaceAfter: 160 }),
      makePara(`${batchName}的中标候选人公示活动已经结束。现将中标人名单公告如下：`, {
        sizeHalfPt: 24,
        firstLineChars: 2,
        spaceAfter: 160,
      }),
      tableXml,
      makePara(`招标代理机构：${agencyName}`, { sizeHalfPt: 24, align: 'right', spaceBefore: 240, spaceAfter: 60 }),
      makePara(`招标人：${ownerUnit}`, { sizeHalfPt: 24, align: 'right', spaceAfter: 60 }),
      makePara(docDate, { sizeHalfPt: 24, align: 'right', spaceAfter: 80 }),
    ];

    writeDocxFromBody(bodyParts.join(''), outputPath);
    return;
  }

  const tableXml = buildAwardWordTable(payload.packages, colOpts, {
    lotNameHeader: '分标名称',
    winnerHeader: '成交人',
    priceHeader: '最终报价(万元/%)',
  });

  const bodyParts = [
    makePara(batchName, { sizeHalfPt: 30, bold: true, align: 'center', spaceAfter: 40 }),
    makePara('成交结果公告', { sizeHalfPt: 30, bold: true, align: 'center', spaceAfter: 80 }),
    makePara(`（采购编号：${batchCode}）`, { sizeHalfPt: 24, align: 'center', spaceAfter: 160 }),
    makePara('各相关应答人：', { sizeHalfPt: 24, spaceAfter: 100 }),
    makePara(
      `${batchName}（采购编号：${batchCode}）的评审工作已结束，经评审委员会评审并报国家电网有限公司招投标工作领导小组批准，现将成交人名单公告如下。`,
      { sizeHalfPt: 24, firstLineChars: 2, spaceAfter: 160 }
    ),
    tableXml,
    makePara(`代理机构：${agencyName}`, { sizeHalfPt: 24, align: 'right', spaceBefore: 240, spaceAfter: 60 }),
    makePara(`采购人：${ownerUnit}`, { sizeHalfPt: 24, align: 'right', spaceAfter: 60 }),
    makePara(docDate, { sizeHalfPt: 24, align: 'right', spaceAfter: 80 }),
  ];

  writeDocxFromBody(bodyParts.join(''), outputPath);
}

/**
 * 3. 成交通知书交接单 (.xlsx) — exact 19 columns matching `成交通知书交接单-综合能源029.xlsx`
 */
export async function generateNoticeHandoverExcel(payload: AwardDeliverablePayload, outputPath: string): Promise<void> {
  const wb = new ExcelJS.Workbook();
  wb.creator = '招标代理 SOP 工作台';
  const ws = wb.addWorksheet('成交通知书交接单');

  const headers = [
    '序号', '业主单位', '项目编号', '分标编号', '项目名称', '项目类别', '开标时间',
    '分标名称', '包号', '成交厂家', '成交价格（万元）', '服务费（元）',
    '联系人', '电话', '邮箱', '邮政地址', '份数', '项目概算（万元）', '备注',
  ];
  styleHeaderRow(ws.addRow(headers));

  const ownerUnit = payload.overrides.ownerUnit || payload.ownerUnit || payload.batch.project_unit;
  const bidOpenTime = payload.overrides.bidOpeningTime || payload.batch.bid_opening_time;

  payload.packages.forEach((p, idx) => {
    const failed = p.is_failed_lot;
    const row = ws.addRow([
      idx + 1,
      ownerUnit,
      payload.overrides.batchCode || payload.batch.code,
      p.lot_code,
      p.project_name,
      p.category,
      bidOpenTime,
      p.lot_name,
      p.package_no,
      failed ? '流标' : p.winner_name,
      failed ? '/' : p.winner_price_wan !== null ? Number(p.winner_price_wan.toFixed(6)) : formatCleanPrice(p.winner_price_raw, ''),
      failed ? '' : p.service_fee_yuan !== null ? p.service_fee_yuan : '',
      failed ? '/' : p.winner_contact,
      failed ? '/' : p.winner_mobile,
      failed ? '/' : p.winner_email,
      failed ? '/' : p.winner_address,
      failed ? '/' : '',
      p.budget_wan !== null ? p.budget_wan : p.max_price_wan ?? '',
      failed ? p.remark || '流标' : p.remark,
    ]);
    styleDataRow(row, headers.length);
    if (!failed && p.service_fee_yuan !== null) {
      row.getCell(12).numFmt = '0.00';
    }
  });

  ws.columns = [
    { width: 6 }, { width: 26 }, { width: 16 }, { width: 18 }, { width: 34 }, { width: 10 }, { width: 14 },
    { width: 12 }, { width: 8 }, { width: 28 }, { width: 16 }, { width: 14 },
    { width: 12 }, { width: 15 }, { width: 22 }, { width: 36 }, { width: 8 }, { width: 16 }, { width: 16 },
  ];

  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  await wb.xlsx.writeFile(outputPath);
}

/**
 * 4. 项目交接信息表 (.xlsx) — exact 34 columns matching `项目交接信息-CY1226SZN029.xlsx`
 */
export async function generateProjectHandoverExcel(payload: AwardDeliverablePayload, outputPath: string): Promise<void> {
  const wb = new ExcelJS.Workbook();
  wb.creator = '招标代理 SOP 工作台';
  const ws = wb.addWorksheet('项目交接信息');

  const headers = [
    '序号', '内部编号', '匹配时间', '业主单位', '所属单位', '单位类型（主业/产业/系统外）',
    '采购层级（省公司/区域联合授权/授权）', '项目名称（匹配单为准）', '项目概况',
    '采购类别（货物/服务/工程）', '采购组织方式（公开招标/竞争性谈判/单一来源)', '采购组织形式(常规批次/框架协议）',
    '项目预算（万元）', '备注', '批次经理', '批次名称', '项目名称', '项目编号', '标段号', '包号', '子包号',
    '挂网时间', '实际开标时间', '开标地点', '成交挂网日期', '成交通知书日期', '成交通知书推送日期',
    '成交厂家', '成交价格(万元）', '成交服务费（元）', '成交厂家联系人', '成交厂家联系电话', '成交厂家联系邮箱',
    '备注（填延期、项目成交共几家、流标原因等需要特别声明的情况）',
  ];
  styleHeaderRow(ws.addRow(headers));

  const ownerUnit = payload.overrides.ownerUnit || payload.ownerUnit || payload.batch.project_unit;
  const batchManager = payload.overrides.batchManager || '';
  const bidLocation = payload.overrides.bidLocation || process.env.DEFAULT_BID_LOCATION || '评标基地';
  const hangingDate = payload.overrides.announcementTime || payload.batch.announcement_time || '';
  const matchDate = payload.overrides.matchDate || hangingDate;
  const announcementDate = payload.overrides.announcementDate || '';
  const noticeDate = payload.overrides.noticeDate || announcementDate;
  const noticePushDate = payload.overrides.noticePushDate || announcementDate;
  const orgWay = payload.batch.procurement_type === '公开招标' ? '公开招标' : '竞争性谈判';
  const catDisplay = payload.category === '物资' ? '货物' : payload.category;

  payload.packages.forEach((p, idx) => {
    const failed = p.is_failed_lot;
    const row = ws.addRow([
      idx + 1,
      '',
      matchDate,
      ownerUnit,
      ownerUnit,
      payload.batch.unit_nature || '产业',
      '授权',
      p.project_name,
      p.project_overview,
      catDisplay,
      orgWay,
      '常规批次',
      p.budget_wan !== null ? p.budget_wan : p.max_price_wan ?? '',
      '',
      batchManager,
      payload.overrides.batchName || payload.batch.name,
      p.project_name,
      payload.overrides.batchCode || payload.batch.code,
      p.lot_code,
      p.package_no,
      '/',
      payload.overrides.announcementTime || payload.batch.announcement_time,
      payload.overrides.bidOpeningTime || payload.batch.bid_opening_time,
      bidLocation,
      failed ? '/' : announcementDate,
      failed ? '/' : noticeDate,
      failed ? '/' : noticePushDate,
      failed ? '/' : p.winner_name,
      failed ? '/' : p.winner_price_wan !== null ? Number(p.winner_price_wan.toFixed(6)) : formatCleanPrice(p.winner_price_raw, ''),
      failed ? '/' : p.service_fee_yuan !== null ? p.service_fee_yuan : '',
      failed ? '/' : p.winner_contact,
      failed ? '/' : p.winner_mobile,
      failed ? '/' : p.winner_email,
      failed ? '流标' : p.remark,
    ]);
    styleDataRow(row);
    if (!failed && p.service_fee_yuan !== null) {
      row.getCell(30).numFmt = '0.00';
    }
  });

  ws.columns = headers.map((h) => ({ width: h.length > 12 ? 24 : 15 }));
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  await wb.xlsx.writeFile(outputPath);
}

/**
 * 5. 成交结果表 (.xlsx) — exact 22 columns matching `成交结果表029.xlsx` (all bidders per package + declarant)
 */
export async function generateAwardResultTableExcel(payload: AwardDeliverablePayload, outputPath: string): Promise<void> {
  const wb = new ExcelJS.Workbook();
  wb.creator = '招标代理 SOP 工作台';
  const ws = wb.addWorksheet('成交结果表');

  const headers = [
    '序号', '采购项目编号', '分标名称', '分包名称', '最高限价（万元/%）', '项目部门', '申报人',
    '项目名称', '物料描述', '投标人名称（同一个项目所有的投标人都列出来）',
    '投标人联系人姓名', '手机号码', '电子邮件', '符合性',
    '技术', '商务', '价格', '总分', '排序', '基准价(万元/%)', '评标总价(万元/%)', '备注',
  ];
  styleHeaderRow(ws.addRow(headers));

  payload.bidderDetails.forEach((b, idx) => {
    const row = ws.addRow([
      idx + 1,
      b.batch_code,
      b.lot_name,
      b.package_no,
      formatCleanPrice(b.max_price_display, ''),
      b.department,
      b.declarant_info,
      b.project_name,
      b.material_desc,
      b.supplier_name,
      b.contact_name,
      b.mobile,
      b.email,
      b.is_compliant,
      b.tech_score,
      b.biz_score,
      b.price_score,
      b.total_score,
      b.rank_no,
      formatCleanPrice(b.base_price),
      formatCleanPrice(b.eval_price),
      b.remark,
    ]);
    styleDataRow(row);
  });

  ws.columns = [
    { width: 6 }, { width: 16 }, { width: 12 }, { width: 10 }, { width: 14 }, { width: 24 }, { width: 18 },
    { width: 32 }, { width: 22 }, { width: 28 }, { width: 14 }, { width: 15 }, { width: 22 }, { width: 8 },
    { width: 10 }, { width: 10 }, { width: 10 }, { width: 10 }, { width: 8 }, { width: 14 }, { width: 14 }, { width: 16 },
  ];

  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  await wb.xlsx.writeFile(outputPath);
}

/**
 * 6. 服务费计算表 (.xlsx) — live Excel formulas + exact verified values matching `服务费计算表格.xlsx`
 */
export async function generateServiceFeeExcel(payload: AwardDeliverablePayload, outputPath: string): Promise<void> {
  const wb = new ExcelJS.Workbook();
  wb.creator = '招标代理 SOP 工作台';
  const ws = wb.addWorksheet('服务费测算表');

  const headers = [
    '序号', '业主单位', '分标编号', '分标名称', '包号', '项目类别', '中标（成交）厂家',
    '成交价格(万元)', '计费区间', '折扣率', '折扣前服务费(元)', '折扣后服务费(元)', '平台使用费(元)', '备注',
  ];
  styleHeaderRow(ws.addRow(headers));

  const ownerUnit = payload.overrides.ownerUnit || payload.ownerUnit || payload.batch.project_unit;

  let totalPre = 0;
  let totalPost = 0;
  let totalPlatform = 0;

  payload.packages.forEach((p, idx) => {
    const failed = p.is_failed_lot || p.winner_price_wan === null;
    const fee = !failed && p.winner_price_wan !== null
      ? calculateServiceFee(p.winner_price_wan, p.category, p.discount_rate)
      : null;

    if (fee) {
      totalPre += fee.preDiscountYuan;
      totalPost += fee.postDiscountYuan;
      totalPlatform += fee.platformFeeYuan;
    }

    const rIdx = idx + 2;
    const row = ws.addRow([
      idx + 1,
      ownerUnit,
      p.lot_code,
      p.lot_name,
      p.package_no,
      p.category,
      p.is_failed_lot ? '流标' : p.winner_name || '—',
      failed ? '/' : p.winner_price_wan,
      fee ? fee.bracketLabel : '/',
      payload.discountLabel,
      fee ? fee.preDiscountYuan : '',
      fee ? { formula: `ROUND(K${rIdx}*${p.discount_rate},2)`, result: fee.postDiscountYuan } : '',
      fee ? fee.platformFeeYuan : '',
      p.is_failed_lot ? '流标不计费' : p.remark,
    ]);
    styleDataRow(row);
    if (fee) {
      row.getCell(12).numFmt = '0.00';
      row.getCell(13).numFmt = '0.00';
    }
  });

  const sumRowIdx = payload.packages.length + 2;
  const sumRow = ws.addRow([
    '合计', '', '', '', '', '', `成交 ${payload.packages.filter((p) => !p.is_failed_lot && p.winner_name).length} 包`,
    '', '', '',
    { formula: `SUM(K2:K${sumRowIdx - 1})`, result: Number(totalPre.toFixed(4)) },
    { formula: `ROUND(SUM(L2:L${sumRowIdx - 1}),2)`, result: Number(totalPost.toFixed(2)) },
    { formula: `ROUND(SUM(M2:M${sumRowIdx - 1}),2)`, result: Number(totalPlatform.toFixed(2)) },
    '',
  ]);
  styleDataRow(sumRow, headers.length);
  for (let c = 1; c <= headers.length; c++) {
    sumRow.getCell(c).font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF000000' } };
  }
  sumRow.getCell(12).numFmt = '0.00';
  sumRow.getCell(13).numFmt = '0.00';

  ws.columns = [
    { width: 6 }, { width: 26 }, { width: 18 }, { width: 12 }, { width: 8 }, { width: 10 }, { width: 28 },
    { width: 15 }, { width: 18 }, { width: 14 }, { width: 16 }, { width: 16 }, { width: 15 }, { width: 16 },
  ];

  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  await wb.xlsx.writeFile(outputPath);
}
