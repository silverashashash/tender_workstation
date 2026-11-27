import fs from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';
import PizZip from 'pizzip';

export interface AwardBatchData {
  batch: {
    id: string;
    name: string;
    code: string;
    project_unit: string;
    announcement_time: string;
    bid_opening_time: string;
    status: string;
    notes: string;
  };
  projects: Array<{
    seq: number;
    name: string;
    code: string;
    budget: string;
    procurement_method: string;
    department: string;
    notes: string;
  }>;
  overrides: Record<string, string>;
}

/**
 * Generate a formatted .xlsx deliverable for the Award stage.
 */
export async function generateAwardExcel(
  templateKey: string,
  data: AwardBatchData,
  outputPath: string
): Promise<void> {
  const wb = new ExcelJS.Workbook();
  wb.creator = '招标代理 SOP 工作台';
  wb.created = new Date();

  const ws = wb.addWorksheet(
    templateKey === 'candidate_publicity' ? '预成交候选人公示表' : '定标结果汇总表'
  );

  const titleText =
    templateKey === 'candidate_publicity'
      ? `${data.overrides.batchName || data.batch.name} — 预成交/中标候选人公示表`
      : `${data.overrides.batchName || data.batch.name} — 定标结果汇总表`;

  ws.mergeCells('A1:G1');
  const titleCell = ws.getCell('A1');
  titleCell.value = titleText;
  titleCell.font = { name: 'PingFang SC', size: 14, bold: true, color: { argb: 'FF1E293B' } };
  titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
  ws.getRow(1).height = 30;

  ws.mergeCells('A2:G2');
  const metaCell = ws.getCell('A2');
  metaCell.value = `批次编号：${data.overrides.batchCode || data.batch.code}   |   项目单位：${data.overrides.projectUnit || data.batch.project_unit}   |   开标时间：${data.overrides.bidOpeningTime || data.batch.bid_opening_time}   |   生成日期：${new Date().toISOString().slice(0, 10)}`;
  metaCell.font = { name: 'PingFang SC', size: 10, color: { argb: 'FF475569' } };
  metaCell.alignment = { vertical: 'middle', horizontal: 'left' };
  ws.getRow(2).height = 22;

  const headers =
    templateKey === 'candidate_publicity'
      ? ['序号', '标段/项目编号', '具体项目名称', '采购方式', '最高限价/预算', '拟成交候选人', '备注']
      : ['序号', '标段/项目编号', '具体项目名称', '需求部门', '采购方式', '预算金额', '定标说明'];

  const headerRow = ws.addRow(headers);
  headerRow.height = 22;
  ws.pageSetup.printTitlesRow = '3:3';
  ws.autoFilter = {
    from: { row: 3, column: 1 },
    to: { row: 3, column: headers.length },
  };
  const thinBlackBorder = {
    top: { style: 'thin' as const, color: { argb: 'FF000000' } },
    bottom: { style: 'thin' as const, color: { argb: 'FF000000' } },
    left: { style: 'thin' as const, color: { argb: 'FF000000' } },
    right: { style: 'thin' as const, color: { argb: 'FF000000' } },
  };
  for (let c = 1; c <= headers.length; c++) {
    const cell = headerRow.getCell(c);
    cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF000000' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = thinBlackBorder;
  }

  data.projects.forEach((p, idx) => {
    const row = ws.addRow(
      templateKey === 'candidate_publicity'
        ? [
            p.seq || idx + 1,
            p.code,
            p.name,
            p.procurement_method,
            p.budget,
            data.overrides[`winner_${idx}`] || '待填入中标候选人',
            p.notes,
          ]
        : [
            p.seq || idx + 1,
            p.code,
            p.name,
            p.department || data.batch.project_unit,
            p.procurement_method,
            p.budget,
            data.overrides.awardSummary || p.notes || '按评审报告推荐排序定标',
          ]
    );
    row.height = 18;
    for (let c = 1; c <= headers.length; c++) {
      const cell = row.getCell(c);
      cell.font = { name: 'Calibri', size: 11, color: { argb: 'FF000000' } };
      cell.alignment = { vertical: 'middle', wrapText: true };
      cell.border = thinBlackBorder;
    }
  });

  ws.columns = [
    { width: 8 },
    { width: 20 },
    { width: 34 },
    { width: 18 },
    { width: 16 },
    { width: 24 },
    { width: 28 },
  ];

  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  await wb.xlsx.writeFile(outputPath);
}

function escapeXml(str: string): string {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Generate a clean, valid Word (.docx) document using minimal OOXML packaging via PizZip.
 */
export function generateAwardDocx(
  templateKey: string,
  data: AwardBatchData,
  outputPath: string
): void {
  const isNotice = templateKey === 'award_notice';
  const docTitle = isNotice ? '成交 / 中标通知书（草稿）' : '采购定标审批与情况说明';

  const paragraphs = [
    docTitle,
    `批次名称：${data.overrides.batchName || data.batch.name}`,
    `批次编号：${data.overrides.batchCode || data.batch.code}`,
    `项目单位：${data.overrides.projectUnit || data.batch.project_unit}`,
    `挂网时间：${data.overrides.announcementTime || data.batch.announcement_time}    开标时间：${data.overrides.bidOpeningTime || data.batch.bid_opening_time}`,
    `代理机构/经办说明：${data.overrides.awardSummary || '本批次已完成评审委员会评审及结果公示，拟按评审报告推荐意见办理定标及成交通知手续。'}`,
    '涵盖具体招标项目 / 标段清单：',
    ...data.projects.map(
      (p, idx) =>
        `${idx + 1}. ${p.name}（编号：${p.code || '-'}，方式：${p.procurement_method}，限价/预算：${p.budget || '-'}）`
    ),
    `生成时间：${new Date().toISOString().slice(0, 10)}（由招标代理 SOP 工作台自动汇总生成，后续可替换为正式 Word 模板）`,
  ];

  const bodyXml = paragraphs
    .map((text, idx) => {
      const isHeading = idx === 0;
      const sz = isHeading ? 32 : 22;
      const bold = isHeading ? '<w:b/>' : '';
      const jc = isHeading ? '<w:jc w:val="center"/>' : '';
      return `<w:p><w:pPr>${jc}<w:spacing w:after="140"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="PingFang SC" w:eastAsia="PingFang SC"/>${bold}<w:sz w:val="${sz}"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
    })
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
  <w:body>${bodyXml}<w:sectPr/></w:body>
</w:document>`
  );

  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  const buffer = zip.generate({ type: 'nodebuffer' });
  fs.writeFileSync(outputPath, buffer);
}
