import crypto from 'node:crypto';
import path from 'node:path';
import PizZip from 'pizzip';
import XLSX from 'xlsx';
import { evaluateWorkbookFormulas, getCellFormulaError } from './excelFormulaEval.js';
import { Dec } from './toolboxWorkbook.js';

export interface ScoreUploadedFile {
  originalName: string;
  buffer: Buffer;
}

export interface ExpertTotalIssue {
  section: string;
  package_name: string;
  supplier: string;
  majority_same: string;
  diff_exceeded: string;
  source_file: string;
  supplier_location: string;
}

export interface ExpertTotalAuditResult {
  checkedCount: number;
  fileCount: number;
  sheetCount: number;
  expertNum: number | null;
  maxScoreDiff: number;
  items: ExpertTotalIssue[];
  counts: {
    totalIssues: number;
    majoritySame: number;
    diffExceeded: number;
  };
}

export interface ObjectiveCriterion {
  name: string;
  project: string;
}

export interface ObjectiveScoringMethod {
  id: string;
  criteria: ObjectiveCriterion[];
  packages: string[];
}

export interface ObjectiveSelection {
  methodId: string;
  criterion: string;
  step?: number | null;
}

export interface ObjectiveAuditIssue {
  section: string;
  package_name: string;
  supplier: string;
  criterion: string;
  issue_type: '分数不一致' | '分值步长不符' | '专家记录缺失' | '专家数量不足' | '细则未找到';
  detail: string;
  source_location: string;
}

export interface ExpertObjectiveAuditResult {
  checkedCount: number;
  fileCount: number;
  items: ObjectiveAuditIssue[];
  warnings: string[];
  counts: {
    totalIssues: number;
    scoreMismatch: number;
    stepMismatch: number;
    expertMissing: number;
    expertInsufficient: number;
    criterionMissing: number;
  };
}

/**
 * Decode zip entry filenames safely (handles both UTF-8 and Windows GBK/GB18030 archives).
 */
function decodeZipEntryName(rawName: string, rawBytes?: Uint8Array): string {
  if (rawBytes && rawBytes.length > 0) {
    try {
      const utf8 = new TextDecoder('utf-8', { fatal: true }).decode(rawBytes);
      return utf8;
    } catch {
      try {
        return new TextDecoder('gb18030').decode(rawBytes);
      } catch {
        // fallback
      }
    }
  }
  return rawName;
}

/**
 * Expand uploaded files: pass through .xlsx/.xls and automatically extract .xlsx/.xls from .zip archives.
 */
export function expandScoreFiles(files: ScoreUploadedFile[]): ScoreUploadedFile[] {
  const expanded: ScoreUploadedFile[] = [];
  for (const f of files) {
    const ext = path.extname(f.originalName).toLowerCase();
    if (ext === '.zip') {
      let zip: PizZip;
      try {
        zip = new PizZip(f.buffer);
      } catch (err) {
        throw new Error(`压缩包「${f.originalName}」解压失败：${(err as Error).message}`);
      }
      for (const [entryKey, entry] of Object.entries(zip.files)) {
        if (entry.dir) continue;
        // Access raw filename bytes if available in PizZip entry metadata for GBK decoding
        const rawBytes = (entry as unknown as { unsafeOriginalName?: Uint8Array }).unsafeOriginalName;
        const decodedName = decodeZipEntryName(entryKey, rawBytes);
        const baseName = path.basename(decodedName);
        if (!baseName || baseName.startsWith('._') || decodedName.includes('__MACOSX/')) continue;
        const innerExt = path.extname(baseName).toLowerCase();
        if (innerExt !== '.xlsx' && innerExt !== '.xls') continue;
        const buf = Buffer.from(entry.asUint8Array());
        expanded.push({ originalName: baseName, buffer: buf });
      }
    } else if (ext === '.xlsx' || ext === '.xls') {
      expanded.push(f);
    }
  }
  return expanded;
}

interface SheetReader {
  lastRow: number;
  lastCol: number;
  text(r: number, c: number): string;
  numericText(r: number, c: number): string;
  rowLastCol(r: number): number;
  merges: Array<{ s: { r: number; c: number }; e: { r: number; c: number } }>;
}

function createSheetReader(ws: XLSX.WorkSheet): SheetReader {
  const ref = ws['!ref'];
  const range = ref ? XLSX.utils.decode_range(ref) : { s: { r: 0, c: 0 }, e: { r: -1, c: -1 } };
  const merges = (ws['!merges'] || []).map((m) => ({
    s: { r: m.s.r, c: m.s.c },
    e: { r: m.e.r, c: m.e.c },
  }));

  const getCell = (r: number, c: number): XLSX.CellObject | undefined => {
    if (r < 0 || c < 0) return undefined;
    return ws[XLSX.utils.encode_cell({ r, c })];
  };

  const text = (r: number, c: number): string => {
    const cell = getCell(r, c);
    if (!cell || cell.t === 'z') return '';
    if (cell.t === 'e') {
      const fErr = getCellFormulaError(cell);
      throw new Error(fErr || `第${r + 1}行第${c + 1}列存在错误单元格（${cell.w || cell.v || '#ERROR'}）`);
    }
    if (cell.w !== undefined && cell.w !== null) return String(cell.w);
    if (cell.v === undefined || cell.v === null) return '';
    return String(cell.v);
  };

  const numericText = (r: number, c: number): string => {
    const cell = getCell(r, c);
    if (!cell || cell.t === 'z') return '';
    if (cell.t === 'e') {
      const fErr = getCellFormulaError(cell);
      throw new Error(fErr || `第${r + 1}行第${c + 1}列存在错误单元格（${cell.w || cell.v || '#ERROR'}）`);
    }
    if (cell.t === 'n' && typeof cell.v === 'number') {
      return Dec.of(String(cell.v)).fmt();
    }
    if (cell.v !== undefined && cell.v !== null) {
      return String(cell.v).trim();
    }
    return cell.w ? String(cell.w).trim() : '';
  };

  const rowLastCol = (r: number): number => {
    let maxC = -1;
    for (let c = range.e.c; c >= 0; c--) {
      const cObj = ws[XLSX.utils.encode_cell({ r, c })] as XLSX.CellObject | undefined;
      if (cObj !== undefined && cObj.t !== 'z') {
        maxC = c;
        break;
      }
    }
    return maxC + 1;
  };

  return {
    lastRow: range.e.r,
    lastCol: range.e.c + 1,
    text,
    numericText,
    rowLastCol,
    merges,
  };
}

function readWorkbookSafe(buffer: Buffer, fileName: string, password?: string): XLSX.WorkBook {
  let wb: XLSX.WorkBook;
  try {
    wb = XLSX.read(buffer, {
      type: 'buffer',
      password: password || undefined,
      cellDates: false,
      sheetStubs: true,
      cellNF: true,
      cellStyles: true,
    });
  } catch (err) {
    const msg = (err as Error).message || '';
    if (/password|encrypt/i.test(msg)) {
      throw new Error(`工作簿「${fileName}」已加密：未提供正确口令或加密格式不支持`);
    }
    throw new Error(`读取工作簿「${fileName}」失败：${msg}`);
  }
  evaluateWorkbookFormulas(wb);
  return wb;
}

/**
 * Tool 8: 专家总分校验 (migrated from ExpertTotalTool.java)
 * - Supports explicit expertNum OR auto-detection (expertNum = null / 0) per sheet from the "平均分" row
 * - Checks:
 *   1. 超过半数专家同分 (count > Math.floor(experts / 2))
 *   2. 任意两名专家分差严格大于阈值 (diff > threshold, default 8)
 */
export function runExpertTotalAudit(params: {
  files: ScoreUploadedFile[];
  expertNum?: number | null;
  maxScoreDiff?: number;
  password?: string;
}): ExpertTotalAuditResult {
  const expandedFiles = expandScoreFiles(params.files);
  if (expandedFiles.length === 0) {
    throw new Error('请选择至少一份专家总分工作簿（支持 .xlsx / .xls / .zip）');
  }

  const explicitExperts =
    params.expertNum !== undefined && params.expertNum !== null && params.expertNum > 0
      ? Math.floor(params.expertNum)
      : null;
  if (explicitExperts !== null && (explicitExperts < 1 || explicitExperts > 1_048_569)) {
    throw new Error('专家人数超出有效范围');
  }

  const threshold = params.maxScoreDiff !== undefined && Number.isFinite(params.maxScoreDiff) ? params.maxScoreDiff : 8;
  const thresholdDec = Dec.of(String(threshold));
  if (thresholdDec.signum() < 0) {
    throw new Error('分差阈值不能为负数');
  }

  const items: ExpertTotalIssue[] = [];
  let checkedCount = 0;
  let sheetCount = 0;
  let detectedExperts: number | null = explicitExperts;

  for (const file of expandedFiles) {
    const wb = readWorkbookSafe(file.buffer, file.originalName, params.password);
    const commonTitle =
      wb.SheetNames.length > 0 ? createSheetReader(wb.Sheets[wb.SheetNames[0]]).text(1, 0) : '';

    for (const sheetName of wb.SheetNames) {
      sheetCount++;
      const ws = wb.Sheets[sheetName];
      const reader = createSheetReader(ws);
      let title = reader.text(1, 0);
      if (!title.trim()) title = commonTitle;

      const match = title.match(/材料名称[:：]\s*(.*?)\s*日期[:：]/s);
      if (!match || !match[1].trim()) {
        throw new Error(`${file.originalName} / ${sheetName} 缺少材料名称及日期`);
      }
      const section = match[1].trim();

      // Determine experts count for this sheet
      let experts = explicitExperts;
      if (experts === null) {
        let foundAvgRow = -1;
        for (let r = 4; r <= Math.min(reader.lastRow, 100); r++) {
          if (reader.text(r, 0).includes('平均分')) {
            foundAvgRow = r;
            break;
          }
        }
        if (foundAvgRow < 5) {
          throw new Error(`${file.originalName} / ${sheetName} 无法自动识别专家人数（未找到平均分行），请手动指定专家人数`);
        }
        experts = foundAvgRow - 4;
        if (detectedExperts === null) detectedExperts = experts;
      }

      for (let base = 0; base <= reader.lastRow; base += experts + 7) {
        let columns = reader.rowLastCol(base + 3);
        let hasScores = false;
        let hasAnySupplierName = false;
        for (let c = 1; c < columns; c++) {
          if (reader.text(base + 3, c).trim()) hasAnySupplierName = true;
        }
        for (let r = base + 4; r <= Math.min(reader.lastRow, base + experts + 3); r++) {
          const rCols = reader.rowLastCol(r);
          columns = Math.max(columns, rCols);
          for (let c = 1; c < rCols; c++) {
            if (reader.text(r, c).trim()) hasScores = true;
          }
        }

        if (!hasAnySupplierName && !hasScores) {
          continue;
        }
        if (!hasAnySupplierName && hasScores) {
          throw new Error(`${file.originalName} / ${sheetName} 第${base + 4}行缺少供应商名称，但本页存在评分`);
        }

        if (!reader.text(base + experts + 4, 0).includes('平均分')) {
          throw new Error(`${file.originalName} / ${sheetName} 第${base + 1}行起专家人数/分页格式与设置不符（当前专家人数：${experts}）`);
        }

        const expertLabels: string[] = [];
        for (let j = 0; j < experts; j++) {
          let label = reader.text(base + 4 + j, 0).trim();
          const sep = label.indexOf('_');
          if (sep >= 0) label = label.slice(sep + 1).trim();
          expertLabels.push(label || '专家姓名未填写');
        }

        for (let col = 1; col < columns; col++) {
          const supplier = reader.text(base + 3, col).trim();
          if (!supplier) {
            for (let j = 0; j < experts; j++) {
              if (reader.text(base + 4 + j, col).trim()) {
                throw new Error(`${file.originalName} / ${sheetName} 有分数但缺少投标人名称，列${col + 1}`);
              }
            }
            continue;
          }

          const scores: Dec[] = [];
          for (let j = 0; j < experts; j++) {
            const numStr = reader.numericText(base + 4 + j, col).trim();
            let scoreDec: Dec;
            try {
              scoreDec = Dec.of(numStr);
              if (scoreDec.signum() < 0) throw new Error('negative');
            } catch {
              throw new Error(`${file.originalName} / ${sheetName} 第${base + 5 + j}行第${col + 1}列分数为空或无效`);
            }
            scores.push(scoreDec.stripTrailingZeros());
          }

          checkedCount++;

          // 1. 超过半数同分（基于精确 Dec 比对，与 Java BigDecimal 1:1 一致）
          const countsMap = new Map<string, { score: Dec; count: number }>();
          for (const sc of scores) {
            const key = sc.fmt();
            const cur = countsMap.get(key);
            if (cur) cur.count++;
            else countsMap.set(key, { score: sc, count: 1 });
          }
          const majorityParts: string[] = [];
          for (const { score: scoreVal, count: cnt } of countsMap.values()) {
            if (cnt > Math.floor(experts / 2)) {
              const namesAtScore: string[] = [];
              for (let k = 0; k < scores.length; k++) {
                if (scores[k].eq(scoreVal)) namesAtScore.push(expertLabels[k]);
              }
              majorityParts.push(`${cnt}个相同分数：${scoreVal.fmt()}；专家：${namesAtScore.join('、')}`);
            }
          }

          // 2. 分差严格大于阈值（基于精确 Dec 运算，与 Java BigDecimal 1:1 一致）
          const diffSet = new Set<string>();
          for (let i = 0; i < scores.length; i++) {
            for (let j = i + 1; j < scores.length; j++) {
              const high = scores[i].max(scores[j]);
              const low = scores[i].min(scores[j]);
              const diff = high.sub(low);
              if (diff.cmp(thresholdDec) > 0) {
                const highIdx = scores[i].cmp(scores[j]) > 0 ? i : j;
                const lowIdx = highIdx === i ? j : i;
                diffSet.add(
                  `分差${diff.fmt()}分：高分${high.fmt()}（${expertLabels[highIdx]}），低分${low.fmt()}（${expertLabels[lowIdx]}）`
                );
              }
            }
          }

          if (majorityParts.length > 0 || diffSet.size > 0) {
            items.push({
              section,
              package_name: sheetName,
              supplier,
              majority_same: majorityParts.join('；'),
              diff_exceeded: Array.from(diffSet).join('；\n'),
              source_file: file.originalName,
              supplier_location: `第${base + 4}行第${col + 1}列`,
            });
          }
        }
      }
    }
  }

  if (checkedCount === 0) {
    throw new Error('未找到可核验的专家评分');
  }

  return {
    checkedCount,
    fileCount: expandedFiles.length,
    sheetCount,
    expertNum: detectedExperts,
    maxScoreDiff: threshold,
    items,
    counts: {
      totalIssues: items.length,
      majoritySame: items.filter((x) => Boolean(x.majority_same)).length,
      diffExceeded: items.filter((x) => Boolean(x.diff_exceeded)).length,
    },
  };
}

// ==================== Tool 9: 专家客观分校验 (migrated from ObjectiveTool.java) ====================

function isCriteriaHeader(text: string): boolean {
  const t = text.trim();
  return t === '商务评分细则' || t === '技术评分细则';
}

function extractObjectiveSectionName(text: string): string {
  const match = text.match(/材料名称[:：]\s*(.*?)\s*(?:设备\/材料名称|材料名称|分包编号|项目单位|工程名称|日期)[:：]/s);
  return match ? match[1].trim() : '';
}

function extractSheetCriteria(sheetName: string, reader: SheetReader): ObjectiveCriterion[] {
  let definition: ObjectiveCriterion[] | null = null;
  for (let r = 0; r <= reader.lastRow; r++) {
    if (isCriteriaHeader(reader.text(r, 1))) {
      const page: ObjectiveCriterion[] = [];
      let foundTotal = false;
      for (r += 2; r <= reader.lastRow; r++) {
        if (reader.text(r, 0).trim() === '合计' || reader.text(r, 1).trim() === '合计') {
          foundTotal = true;
          break;
        }
        const item = reader.text(r, 1).trim();
        if (!item) continue;
        let projectRow = r;
        for (const m of reader.merges) {
          if (r >= m.s.r && r <= m.e.r && 0 >= m.s.c && 0 <= m.e.c) {
            projectRow = m.s.r;
            break;
          }
        }
        const project = reader.text(projectRow, 0).trim() || '未填写项目名称';
        page.push({ name: item, project });
      }
      if (!foundTotal) {
        throw new Error(`${sheetName} 缺少合计行`);
      }
      if (definition === null) {
        definition = page;
      } else {
        const same =
          definition.length === page.length &&
          definition.every((d, idx) => d.name === page[idx].name && d.project === page[idx].project);
        if (!same) {
          throw new Error(`${sheetName} 同一工作表分页评分办法不一致`);
        }
      }
    }
  }
  if (!definition || definition.length === 0) {
    throw new Error(`${sheetName} 未找到评分细则`);
  }
  return definition;
}

function computeMethodId(criteria: ObjectiveCriterion[]): string {
  let data = '';
  for (const item of criteria) {
    data += `${item.project.length}:${item.project}${item.name.length}:${item.name}`;
  }
  return crypto.createHash('sha256').update(data, 'utf8').digest('hex');
}

export function discoverObjectiveMethods(files: ScoreUploadedFile[]): {
  expandedFiles: ScoreUploadedFile[];
  methods: ObjectiveScoringMethod[];
} {
  const expandedFiles = expandScoreFiles(files);
  if (expandedFiles.length === 0) {
    throw new Error('请先选择专家评分明细表（支持 .xlsx / .xls / .zip）');
  }

  const definitions = new Map<string, ObjectiveCriterion[]>();
  const packagesMap = new Map<string, Set<string>>();

  for (const file of expandedFiles) {
    const wb = readWorkbookSafe(file.buffer, file.originalName);
    for (const sheetName of wb.SheetNames) {
      const reader = createSheetReader(wb.Sheets[sheetName]);
      const items = extractSheetCriteria(sheetName, reader);
      const id = computeMethodId(items);
      if (!definitions.has(id)) definitions.set(id, items);
      let section = extractObjectiveSectionName(reader.text(1, 0));
      if (!section) section = file.originalName;
      if (!packagesMap.has(id)) packagesMap.set(id, new Set());
      packagesMap.get(id)!.add(`${section} / ${sheetName}`);
    }
  }

  if (definitions.size === 0) {
    throw new Error('没有读取到评分细则');
  }

  const methods: ObjectiveScoringMethod[] = [];
  for (const [id, criteria] of definitions.entries()) {
    methods.push({
      id,
      criteria,
      packages: Array.from(packagesMap.get(id) || []),
    });
  }

  return { expandedFiles, methods };
}

function isMultipleOfStep(value: Dec, step: Dec): boolean {
  const [, rem] = value.divAndRem(step);
  return rem.signum() === 0;
}

export function runExpertObjectiveAudit(params: {
  files: ScoreUploadedFile[];
  selections: ObjectiveSelection[];
  sectionOverride?: string;
}): ExpertObjectiveAuditResult {
  const expandedFiles = expandScoreFiles(params.files);
  if (expandedFiles.length === 0) {
    throw new Error('请先上传专家评分明细表');
  }
  if (!params.selections || params.selections.length === 0) {
    throw new Error('请至少勾选一项客观评分细则');
  }

  // Group rules by methodId ("" means global)
  const ruleSets = new Map<string, Map<string, Dec | null>>();
  for (const sel of params.selections) {
    const methodKey = String(sel.methodId || '').trim();
    const item = String(sel.criterion || '').trim();
    if (!item) throw new Error('客观评分细则名称不能为空');
    let rules = ruleSets.get(methodKey);
    if (!rules) {
      rules = new Map();
      ruleSets.set(methodKey, rules);
    }
    if (rules.has(item)) throw new Error(`规则细则重复：${item}`);
    const rawStep = sel.step !== undefined && sel.step !== null ? String(sel.step).trim() : '';
    let stepVal: Dec | null = null;
    if (rawStep !== '') {
      try {
        stepVal = Dec.of(rawStep);
      } catch {
        throw new Error(`分值步长必须大于 0：${item}`);
      }
      if (stepVal.signum() <= 0) {
        throw new Error(`分值步长必须大于 0：${item}`);
      }
    }
    rules.set(item, stepVal);
  }

  const matchedMethods = new Set<string>();
  const scopeMethods = new Map<string, string>();
  const scopeRules = new Map<string, Map<string, Dec | null>>();
  const scoresMap = new Map<
    string,
    {
      section: string;
      pack: string;
      supplier: string;
      criterion: string;
      values: Map<string, { value: Dec; source: string }>;
    }
  >();
  const expertsByScope = new Map<string, Set<string>>();

  const outputs: ObjectiveAuditIssue[] = [];
  const warnings: string[] = [];

  for (const file of expandedFiles) {
    const stem = file.originalName.replace(/\.xlsx?$/i, '');
    const parts = stem.split('_');
    if (parts.length < 3 || !parts[1].trim() || !parts.slice(2).join('_').trim()) {
      throw new Error(`文件名须为「评分表_专家编号_专家姓名.xlsx」格式：${file.originalName}`);
    }
    const expert = `${parts[1].trim()}_${parts.slice(2).join('_').trim()}`;

    const wb = readWorkbookSafe(file.buffer, file.originalName);
    for (const sheetName of wb.SheetNames) {
      const reader = createSheetReader(wb.Sheets[sheetName]);
      let section = String(params.sectionOverride || '').trim();
      if (!section) {
        section = extractObjectiveSectionName(reader.text(1, 0));
        if (!section) {
          throw new Error(`${file.originalName} / ${sheetName} 无法提取分标名称，请在界面填写分标名称覆盖`);
        }
      }

      const method = computeMethodId(extractSheetCriteria(sheetName, reader));
      matchedMethods.add(method);

      const rules = new Map<string, Dec | null>(ruleSets.get('') || []);
      for (const [k, v] of (ruleSets.get(method) || new Map()).entries()) {
        if (rules.has(k)) throw new Error(`同一细则同时配置全局与评分办法规则：${k}`);
        rules.set(k, v);
      }

      const scopeKey = `${section}::${sheetName}`;
      const prevMethod = scopeMethods.get(scopeKey);
      if (prevMethod && prevMethod !== method) {
        throw new Error(`同标包不同专家的评分办法不一致：${section} / ${sheetName}`);
      }
      scopeMethods.set(scopeKey, method);

      if (rules.size === 0) continue;
      scopeRules.set(scopeKey, rules);

      if (!expertsByScope.has(scopeKey)) expertsByScope.set(scopeKey, new Set());
      if (expertsByScope.get(scopeKey)!.has(expert)) {
        throw new Error(`同标包专家重复导入：${expert} / ${section} / ${sheetName}`);
      }
      expertsByScope.get(scopeKey)!.add(expert);

      let count = 0;
      for (let r = 4; r <= reader.lastRow; r++) {
        if (reader.text(r, 0).trim() === '合计' || reader.text(r, 1).trim() === '合计') {
          count = r - 4;
          break;
        }
      }
      if (count < 1) {
        throw new Error(`${file.originalName} / ${sheetName} 未找到评分项后的合计行`);
      }

      let stride = count + 8;
      if (/^第\s*\d+\s*页\s*[/／]\s*共\s*\d+\s*页$/.test(reader.text(stride, 0).trim())) {
        stride++;
      }

      const foundItems = new Set<string>();
      let anyPage = false;

      for (let base = 0; base <= reader.lastRow; base += stride) {
        const namesLastCol = reader.rowLastCol(base + 3);
        let hasNameRow = false;
        for (let c = 0; c < namesLastCol; c++) {
          if (reader.text(base + 3, c).trim()) {
            hasNameRow = true;
            break;
          }
        }
        if (!hasNameRow) {
          let meaningful = false;
          for (let r = base; r <= reader.lastRow; r++) {
            const rc = reader.rowLastCol(r);
            for (let c = 0; c < rc; c++) {
              if (reader.text(r, c).trim()) {
                meaningful = true;
                break;
              }
            }
            if (meaningful) break;
          }
          if (meaningful) {
            throw new Error(`分页不完整：${file.originalName} / ${sheetName} 第${base + 1}行`);
          }
          continue;
        }

        if (!isCriteriaHeader(reader.text(base + 2, 1))) {
          throw new Error(`评分细则表头不符：${file.originalName} / ${sheetName} 第${base + 3}行B列`);
        }
        if (reader.text(base + count + 4, 0).trim() !== '合计' && reader.text(base + count + 4, 1).trim() !== '合计') {
          throw new Error(`分页评分项数不一致：${file.originalName} / ${sheetName}`);
        }
        anyPage = true;

        for (let r = base + 4; r < base + count + 4; r++) {
          const criterion = reader.text(r, 1).trim();
          if (!rules.has(criterion)) continue;
          foundItems.add(criterion);

          const lastCol = Math.max(namesLastCol, reader.rowLastCol(r));
          for (let col = 3; col < lastCol; col++) {
            const supplier = reader.text(base + 3, col).trim();
            const numStr = reader.numericText(r, col).trim();
            const location = `${file.originalName} / ${sheetName} / ${r + 1}行${col + 1}列`;
            if (!supplier) {
              if (numStr) throw new Error(`${location} 有分数但缺少供应商名称`);
              continue;
            }
            if (!numStr) throw new Error(`${location} 缺少客观分`);
            let valDec: Dec;
            try {
              valDec = Dec.of(numStr).stripTrailingZeros();
            } catch {
              throw new Error(`${location} 客观分无效：${numStr}`);
            }

            const scoreKey = `${section}::${sheetName}::${supplier}::${criterion}`;
            let entry = scoresMap.get(scoreKey);
            if (!entry) {
              entry = { section, pack: sheetName, supplier, criterion, values: new Map() };
              scoresMap.set(scoreKey, entry);
            }
            if (entry.values.has(expert)) {
              throw new Error(`${location} 同一专家/供应商/细则重复，不能覆盖`);
            }
            entry.values.set(expert, { value: valDec, source: location });
          }
        }
      }

      if (!anyPage) {
        throw new Error(`没有完整评分页：${file.originalName} / ${sheetName}`);
      }

      for (const item of rules.keys()) {
        if (!foundItems.has(item)) {
          outputs.push({
            section,
            package_name: sheetName,
            supplier: '',
            criterion: item,
            issue_type: '细则未找到',
            detail: `${expert}：未找到完全一致的细则名称`,
            source_location: file.originalName,
          });
          warnings.push(`${file.originalName} / ${sheetName} 未找到客观项：${item}`);
        }
      }
    }
  }

  for (const methodKey of ruleSets.keys()) {
    if (methodKey && !matchedMethods.has(methodKey)) {
      throw new Error('所选评分办法未匹配当前文件，请重新读取评分项并勾选');
    }
  }

  for (const entry of scoresMap.values()) {
    const scopeKey = `${entry.section}::${entry.pack}`;
    const valuesArr = Array.from(entry.values.entries());
    const detailText = valuesArr.map(([exp, v]) => `${exp}：${v.value.fmt()}分`).join('\n');
    const sourcesText = valuesArr.map(([, v]) => v.source).join('\n');

    const expected = expertsByScope.get(scopeKey) || new Set<string>();
    const missingExperts = Array.from(expected).filter((e) => !entry.values.has(e));
    if (missingExperts.length > 0) {
      outputs.push({
        section: entry.section,
        package_name: entry.pack,
        supplier: entry.supplier,
        criterion: entry.criterion,
        issue_type: '专家记录缺失',
        detail: `未找到：${missingExperts.join('、')}`,
        source_location: sourcesText,
      });
    }

    if (entry.values.size < 2) {
      outputs.push({
        section: entry.section,
        package_name: entry.pack,
        supplier: entry.supplier,
        criterion: entry.criterion,
        issue_type: '专家数量不足',
        detail: `仅1位专家有此项评分，无法核验一致性；${detailText}`,
        source_location: sourcesText,
      });
    }

    const distinctScores = new Set(valuesArr.map(([, v]) => v.value.fmt()));
    if (distinctScores.size > 1) {
      outputs.push({
        section: entry.section,
        package_name: entry.pack,
        supplier: entry.supplier,
        criterion: entry.criterion,
        issue_type: '分数不一致',
        detail: detailText,
        source_location: sourcesText,
      });
    }

    const step = scopeRules.get(scopeKey)?.get(entry.criterion) ?? null;
    if (step !== null) {
      for (const [exp, v] of valuesArr) {
        if (!isMultipleOfStep(v.value, step)) {
          outputs.push({
            section: entry.section,
            package_name: entry.pack,
            supplier: entry.supplier,
            criterion: entry.criterion,
            issue_type: '分值步长不符',
            detail: `${exp}：${v.value.fmt()}分，不是${step.fmt()}的倍数`,
            source_location: v.source,
          });
        }
      }
    }
  }

  if (scoresMap.size === 0) {
    warnings.push('没有找到可核验的客观项评分，请核对所勾选的评分细则');
  }

  return {
    checkedCount: scoresMap.size,
    fileCount: expandedFiles.length,
    items: outputs,
    warnings,
    counts: {
      totalIssues: outputs.length,
      scoreMismatch: outputs.filter((x) => x.issue_type === '分数不一致').length,
      stepMismatch: outputs.filter((x) => x.issue_type === '分值步长不符').length,
      expertMissing: outputs.filter((x) => x.issue_type === '专家记录缺失').length,
      expertInsufficient: outputs.filter((x) => x.issue_type === '专家数量不足').length,
      criterionMissing: outputs.filter((x) => x.issue_type === '细则未找到').length,
    },
  };
}
