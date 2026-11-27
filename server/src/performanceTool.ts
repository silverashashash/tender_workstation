import {
  Dec,
  buildStyledExcelBuffer,
  colsFromList,
  numericText,
  readWorkbookFirstSheet,
} from './toolboxWorkbook.js';

export const PERF_BID_HEADERS = [
  '采购项目编号',
  '分标名称',
  '分标编号',
  '分包名称',
  '分包编号',
  '项目单位',
  '投标人名称',
  '单位名称',
  '项目归口管理部门',
];

export const PERF_SCORE_HEADERS = [
  '单位名称',
  '关联类型',
  '所属部门',
  '服务类型/物料名称',
  '供应商名称',
  '综合评价',
  '备注',
];

interface Rating {
  score: number;
  note: string;
  applies: boolean;
}

export interface PerformanceRow {
  project_code: string;
  lot_name: string;
  lot_code: string;
  pack_name: string;
  pack_code: string;
  project_unit: string;
  bidder_name: string;
  unit_name: string;
  department: string;
  evaluation: string;
  score: number;
  remark: string;
  match_note: string;
}

export interface PerformanceResult {
  bidFile: string;
  scoreFile: string;
  totalRows: number;
  matchedCount: number;
  reviewCount: number;
  rows: PerformanceRow[];
  warnings: string[];
}

const NAMED_RATINGS = new Map<string, number>([
  ['优秀', 5],
  ['A', 5],
  ['良好', 4],
  ['B', 4],
  ['一般', 3],
  ['C', 3],
  ['及格', 2],
  ['D', 2],
  ['不及格', 1],
  ['E', 1],
]);

function rateEvaluation(rawText: string): Rating {
  const text = rawText.trim();
  if (text === '' || text === '/' || text === '／' || text === '0') {
    return { score: 4, note: '评价为未知标记（空、/或0），按默认4分；请复核', applies: false };
  }
  const named = NAMED_RATINGS.get(text);
  if (named !== undefined) {
    return { score: named, note: '', applies: true };
  }
  let num: Dec;
  try {
    num = Dec.of(text);
  } catch {
    return { score: 4, note: '评价无法识别，沿用原版默认4分；请复核', applies: false };
  }
  if (num.signum() < 0) {
    throw new Error('履约评价不能为负数');
  }
  if (num.cmp(Dec.of(90)) >= 0) return { score: 5, note: '', applies: true };
  if (num.cmp(Dec.of(80)) > 0) return { score: 4, note: '', applies: true };
  if (num.cmp(Dec.of(60)) > 0) return { score: 3, note: '', applies: true };
  if (num.cmp(Dec.of(30)) > 0) return { score: 2, note: '', applies: true };
  if (num.cmp(Dec.of(30)) < 0) return { score: 1, note: '', applies: true };
  return { score: 4, note: '评价恰好30，原版未覆盖该边界，沿用默认4分；请复核', applies: false };
}

export function runPerformanceScore(
  bidFile: { name: string; buffer: Buffer },
  scoreFile: { name: string; buffer: Buffer }
): PerformanceResult {
  const bids = readWorkbookFirstSheet(bidFile.buffer, bidFile.name, colsFromList(PERF_BID_HEADERS));
  const evaluations = readWorkbookFirstSheet(scoreFile.buffer, scoreFile.name, colsFromList(PERF_SCORE_HEADERS));

  if (bids.rows.length === 0) throw new Error('开标表没有数据');

  const validLinkTypes = new Set(['供应商', '供应商+标段', '供应商+所属部门']);
  for (const row of evaluations.rows) {
    const type = row['关联类型'] ?? '';
    if (!validLinkTypes.has(type)) {
      throw new Error(`履约评价第${row.__row}行关联类型不支持：${type}`);
    }
    for (const key of ['单位名称', '供应商名称']) {
      if (!(row[key] ?? '').trim()) {
        throw new Error(`履约评价第${row.__row}行缺少${key}`);
      }
    }
    if (type === '供应商+标段' && !(row['服务类型/物料名称'] ?? '').trim()) {
      throw new Error('按标段关联时必须填写服务类型/物料名称');
    }
    if (type === '供应商+所属部门' && !(row['所属部门'] ?? '').trim()) {
      throw new Error('按部门关联时必须填写所属部门');
    }
  }

  const rows: PerformanceRow[] = [];
  const warnings: string[] = [];
  let matchedCount = 0;

  for (const bid of bids.rows) {
    for (const key of ['单位名称', '投标人名称', '分标名称']) {
      if (!(bid[key] ?? '').trim()) {
        throw new Error(`开标表第${bid.__row}行缺少${key}`);
      }
    }

    let match: Record<string, string> | null = null;
    let matches = 0;
    let rating: Rating | null = null;
    const skipped: string[] = [];
    let validRating = false;

    for (const row of evaluations.rows) {
      if (row['单位名称'] !== bid['单位名称'] || row['供应商名称'] !== bid['投标人名称']) continue;
      if (row['关联类型'] === '供应商+标段' && row['服务类型/物料名称'] !== bid['分标名称']) continue;
      if (row['关联类型'] === '供应商+所属部门' && row['所属部门'] !== bid['项目归口管理部门']) continue;

      matches++;
      const candidate = rateEvaluation(numericText(row, '综合评价'));
      if (candidate.applies) {
        match = row;
        rating = candidate;
        validRating = true;
      } else {
        const evalRaw = (row['综合评价'] ?? '').trim();
        const isBlankMarker = evalRaw === '' || evalRaw === '/' || evalRaw === '／' || evalRaw === '0';
        skipped.push(isBlankMarker ? `评价“${evalRaw || '空'}”为空或占位标记，未覆盖有效评价` : `评价“${row['综合评价']}”未被原版分段覆盖`);
        if (!validRating) {
          match = row;
          rating = candidate;
        }
      }
    }

    if (!rating) {
      rating = { score: 4, note: '未匹配评价，按原版默认4分', applies: false };
    }
    let note = rating.note;
    if (validRating && skipped.length > 0) {
      note = [note, ...skipped, '保留末条有效评价，请复核'].filter(Boolean).join('；');
    }
    if (matches > 1) {
      note = note
        ? `匹配${matches}条，采用末条有效评价（无有效评价时默认4）；${note}`
        : `匹配${matches}条，采用末条有效评价（无有效评价时默认4）`;
    }
    if (note) {
      warnings.push(`第${bid.__row}行 ${bid['投标人名称']}：${note}`);
    } else {
      matchedCount++;
    }

    rows.push({
      project_code: bid['采购项目编号'] ?? '',
      lot_name: bid['分标名称'] ?? '',
      lot_code: bid['分标编号'] ?? '',
      pack_name: bid['分包名称'] ?? '',
      pack_code: bid['分包编号'] ?? '',
      project_unit: bid['项目单位'] ?? '',
      bidder_name: bid['投标人名称'] ?? '',
      unit_name: bid['单位名称'] ?? '',
      department: bid['项目归口管理部门'] ?? '',
      evaluation: match ? match['综合评价'] ?? '' : '',
      score: rating.score,
      remark: match ? match['备注'] ?? '' : '',
      match_note: note || '已匹配',
    });
  }

  return {
    bidFile: bidFile.name,
    scoreFile: scoreFile.name,
    totalRows: rows.length,
    matchedCount,
    reviewCount: warnings.length,
    rows,
    warnings,
  };
}

export async function exportPerformanceExcel(result: PerformanceResult): Promise<Buffer> {
  return buildStyledExcelBuffer([
    {
      name: '履约分',
      headers: [...PERF_BID_HEADERS, '综合评价', '履约分结果', '备注', '匹配说明'],
      rows: result.rows.map((r) => [
        r.project_code,
        r.lot_name,
        r.lot_code,
        r.pack_name,
        r.pack_code,
        r.project_unit,
        r.bidder_name,
        r.unit_name,
        r.department,
        r.evaluation,
        String(r.score),
        r.remark,
        r.match_note,
      ]),
    },
  ]);
}

export async function exportPerformanceTemplate(kind: 'bids' | 'scores'): Promise<Buffer> {
  if (kind === 'bids') {
    return buildStyledExcelBuffer([{ name: '开标表', headers: PERF_BID_HEADERS, rows: [] }]);
  }
  return buildStyledExcelBuffer([{ name: '履约评价表', headers: PERF_SCORE_HEADERS, rows: [] }]);
}
