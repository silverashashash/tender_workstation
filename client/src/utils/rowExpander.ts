import * as XLSX from 'xlsx';

export type RowExpandMode =
  | 'simple'
  | 'two_level'
  | 'merge'
  | 'merge_two'
  | 'contacts'
  | 'contacts_existing';

export interface RowExpandOptions {
  mode: RowExpandMode;
  fileName: string;
  colName: string;
  sep: string;
  prefixSep: string;
  prefixName: string;
  itemName: string;
  groupSep: string;
  groupCols: string;
  newColPos: string;
  supplierCol: string;
}

export interface ContactsStats {
  input_rows: number;
  skipped_blank_rows: number;
  expanded_rows: number;
  safe_keys: number;
  conflict_keys: number;
  review_rows: number;
  duplicate_rows: number;
}

export interface RowExpandResult {
  workbook: XLSX.WorkBook;
  outputFileName: string;
  srcRows: number;
  outRows: number;
  logs: string[];
  sheetPreviews: Array<{
    name: string;
    headers: string[];
    rows: string[][];
    totalRows: number;
  }>;
  contactsStats?: ContactsStats;
}

const CONTACT_ALIASES: Record<string, string[]> = {
  供应商: ['供应商', '供应商名称', '投标人名称', '应答人', '应答人名称'],
  分标: ['分标', '标段', '标段名称', '分标名称'],
  包号: ['包号', '分包名称'],
  联系人姓名: [
    '联系人姓名',
    '投标人联系人姓名',
    '供应商联系人姓名',
    '应答人联系人姓名',
    '投标人联系人',
    '供应商联系人',
    '应答人联系人',
    '联系人',
    '姓名',
  ],
  手机号: ['手机号', '手机号码', '手机', '移动电话'],
  地址: ['地址', '邮政地址', '联系地址', '通讯地址', '办公地址', '注册地址'],
  固定电话: ['固定电话', '联系电话', '电话', '办公电话', '座机'],
  邮箱: ['邮箱', '电子邮箱', '电子邮件', '电子邮件地址', '联系邮箱'],
};

const CONTACT_FIELDS = ['联系人姓名', '手机号', '地址', '固定电话', '邮箱'] as const;

function cellToText(cell?: XLSX.CellObject | null): string {
  if (!cell || cell.v === undefined || cell.v === null) return '';
  return String(cell.v).trim();
}

function cellToRawValue(cell?: XLSX.CellObject | null): string | number | boolean | null {
  if (!cell || cell.v === undefined || cell.v === null) return null;
  if (typeof cell.v === 'string' || typeof cell.v === 'number' || typeof cell.v === 'boolean') {
    return cell.v;
  }
  return String(cell.v);
}

function readSheetMatrix(ws: XLSX.WorkSheet): {
  headers: string[];
  dataRows: Array<{ rowNum: number; cells: Array<XLSX.CellObject | undefined>; values: Array<string | number | boolean | null> }>;
  colCount: number;
} {
  const ref = ws['!ref'];
  if (!ref) return { headers: [], dataRows: [], colCount: 0 };
  const range = XLSX.utils.decode_range(ref);
  const colCount = range.e.c + 1;
  const headers: string[] = [];
  for (let c = 0; c < colCount; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: 0, c })] as XLSX.CellObject | undefined;
    headers.push(cellToText(cell));
  }

  const dataRows: Array<{
    rowNum: number;
    cells: Array<XLSX.CellObject | undefined>;
    values: Array<string | number | boolean | null>;
  }> = [];
  for (let r = 1; r <= range.e.r; r++) {
    const cells: Array<XLSX.CellObject | undefined> = [];
    const values: Array<string | number | boolean | null> = [];
    for (let c = 0; c < colCount; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })] as XLSX.CellObject | undefined;
      cells.push(cell);
      values.push(cellToRawValue(cell));
    }
    dataRows.push({ rowNum: r + 1, cells, values });
  }
  return { headers, dataRows, colCount };
}

export function extractWorkbookHeaders(wb: XLSX.WorkBook): {
  sheetName: string;
  headers: string[];
  rowCount: number;
} {
  const sheetName = wb.SheetNames[0] || '';
  const ws = wb.Sheets[sheetName];
  if (!ws?.['!ref']) {
    return { sheetName, headers: [], rowCount: 0 };
  }
  const range = XLSX.utils.decode_range(ws['!ref']);
  const headers: string[] = [];
  for (let c = 0; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: 0, c })] as XLSX.CellObject | undefined;
    const text = cellToText(cell);
    if (text) headers.push(text);
  }
  return { sheetName, headers, rowCount: Math.max(0, range.e.r) };
}

function buildLiteralSheet(
  rows: Array<Array<string | number | boolean | null>>,
  options?: { forceText?: boolean; freezeTop?: boolean; autoFilter?: boolean; colWidths?: number[] }
): XLSX.WorkSheet {
  const ws: XLSX.WorkSheet = {};
  const maxCols = rows.reduce((m, r) => Math.max(m, r.length), 0);

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    for (let c = 0; c < row.length; c++) {
      const val = row[c];
      if (val === null || val === undefined || val === '') continue;
      const addr = XLSX.utils.encode_cell({ r, c });
      if (options?.forceText || typeof val === 'string') {
        ws[addr] = { t: 's', v: String(val) };
      } else if (typeof val === 'number') {
        ws[addr] = { t: 'n', v: val };
      } else if (typeof val === 'boolean') {
        ws[addr] = { t: 'b', v: val };
      } else {
        ws[addr] = { t: 's', v: String(val) };
      }
    }
  }

  if (rows.length > 0 && maxCols > 0) {
    ws['!ref'] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: rows.length - 1, c: maxCols - 1 },
    });
    if (options?.autoFilter) {
      ws['!autofilter'] = { ref: ws['!ref'] };
    }
  }

  if (options?.colWidths) {
    ws['!cols'] = options.colWidths.map((wch) => ({ wch }));
  } else if (maxCols > 0) {
    ws['!cols'] = Array.from({ length: maxCols }, (_, c) => ({
      wch: Math.min(
        48,
        Math.max(
          10,
          ...rows
            .slice(0, 100)
            .map((r) => String(r[c] ?? '').replace(/[^\x00-\xff]/g, 'xx').length + 2)
        )
      ),
    }));
  }

  return ws;
}

function makePreview(
  name: string,
  rows: Array<Array<string | number | boolean | null>>,
  limit = 30
) {
  const headers = (rows[0] || []).map((v) => (v === null || v === undefined ? '' : String(v)));
  const data = rows.slice(1, 1 + limit).map((r) =>
    headers.map((_, i) => (r[i] === null || r[i] === undefined ? '' : String(r[i])))
  );
  return {
    name,
    headers,
    rows: data,
    totalRows: Math.max(0, rows.length - 1),
  };
}

function resolveInsertPos(newColPos: string, header: string[]): number {
  const s = (newColPos || '').trim();
  const n = header.length;
  if (!s) return n;
  if (/^\d+$/.test(s)) {
    return Math.max(0, Math.min(parseInt(s, 10) - 1, n));
  }
  const byName = header.indexOf(s);
  if (byName >= 0) return byName + 1;
  if (/^[A-Za-z]+$/.test(s)) {
    const idx1 = XLSX.utils.decode_col(s.toUpperCase()) + 1;
    if (idx1 >= 1 && idx1 <= n) return idx1;
  }
  throw new Error(
    `新列位置「${s}」无法识别：请填列名、列号(数字)、列字母，或留空表示放最后。当前表头：${header.join('、')}`
  );
}

function splitFirst(str: string, sep: string): [string, boolean, string] {
  const idx = str.indexOf(sep);
  if (idx < 0) return [str, false, ''];
  return [str.slice(0, idx), true, str.slice(idx + sep.length)];
}

function stripExt(fileName: string): string {
  return fileName.replace(/\.(xlsx?|xlsm)$/i, '') || '数据表';
}

export function runRowExpander(wb: XLSX.WorkBook, opts: RowExpandOptions): RowExpandResult {
  const sheetName = wb.SheetNames[0];
  if (!sheetName) throw new Error('工作簿中没有工作表');
  const ws = wb.Sheets[sheetName];
  if (!ws?.['!ref']) throw new Error('首个工作表为空');

  const { headers, dataRows } = readSheetMatrix(ws);
  const logs: string[] = [
    `读取工作表：${sheetName}，共 ${dataRows.length + 1} 行 / ${headers.length} 列`,
  ];

  const stem = stripExt(opts.fileName);

  // ==================== 1. 简单展开 ====================
  if (opts.mode === 'simple') {
    const colName = opts.colName.trim();
    if (!colName) throw new Error('请填写目标列名');
    if (!opts.sep) throw new Error('请填写分隔符');
    const colIdx = headers.indexOf(colName);
    if (colIdx < 0) {
      throw new Error(`未在表头中找到列「${colName}」。当前表头：${headers.join('、')}`);
    }
    logs.push(`简单展开：列「${colName}」（第 ${colIdx + 1} 列），分隔符 ${JSON.stringify(opts.sep)}`);

    const outRows: Array<Array<string | number | boolean | null>> = [headers];
    for (const row of dataRows) {
      const raw = row.values[colIdx] === null ? '' : String(row.values[colIdx]);
      let items = raw
        .split(opts.sep)
        .map((p) => p.trim())
        .filter(Boolean);
      if (items.length === 0) items = [''];
      for (const item of items) {
        outRows.push([...row.values.slice(0, colIdx), item, ...row.values.slice(colIdx + 1)]);
      }
    }

    const outWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(outWb, buildLiteralSheet(outRows), sheetName);
    logs.push(`完成：${dataRows.length} 行 → ${outRows.length - 1} 行`);
    return {
      workbook: outWb,
      outputFileName: `${stem}_展开.xlsx`,
      srcRows: dataRows.length,
      outRows: outRows.length - 1,
      logs,
      sheetPreviews: [makePreview(sheetName, outRows)],
    };
  }

  // ==================== 2. 两级拆分 ====================
  if (opts.mode === 'two_level') {
    const colName = opts.colName.trim();
    if (!colName) throw new Error('请填写目标列名');
    if (!opts.sep) throw new Error('请填写列表分隔符');
    if (!opts.prefixSep) throw new Error('请填写前缀分隔符');
    const colIdx = headers.indexOf(colName);
    if (colIdx < 0) {
      throw new Error(`未在表头中找到列「${colName}」。当前表头：${headers.join('、')}`);
    }
    const prefixName = opts.prefixName.trim() || '分标';
    const itemName = opts.itemName.trim() || '包号';
    logs.push(
      `两级拆分：列「${colName}」（第 ${colIdx + 1} 列）→「${prefixName}」+「${itemName}」，前缀分隔符 ${JSON.stringify(opts.prefixSep)}，列表分隔符 ${JSON.stringify(opts.sep)}`
    );

    const newHeader = [...headers.slice(0, colIdx), prefixName, itemName, ...headers.slice(colIdx + 1)];
    const outRows: Array<Array<string | number | boolean | null>> = [newHeader];
    for (const row of dataRows) {
      const raw = row.values[colIdx] === null ? '' : String(row.values[colIdx]);
      const [prefixRaw, , listPart] = splitFirst(raw, opts.prefixSep);
      const prefix = prefixRaw.trim();
      let items = listPart
        .split(opts.sep)
        .map((b) => b.trim())
        .filter(Boolean);
      if (items.length === 0) items = [''];
      for (const item of items) {
        outRows.push([...row.values.slice(0, colIdx), prefix, item, ...row.values.slice(colIdx + 1)]);
      }
    }

    const outWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(outWb, buildLiteralSheet(outRows), sheetName);
    logs.push(`完成：${dataRows.length} 行 → ${outRows.length - 1} 行`);
    return {
      workbook: outWb,
      outputFileName: `${stem}_展开.xlsx`,
      srcRows: dataRows.length,
      outRows: outRows.length - 1,
      logs,
      sheetPreviews: [makePreview(sheetName, outRows)],
    };
  }

  // ==================== 3. 一级合并 ====================
  if (opts.mode === 'merge') {
    const colName = opts.colName.trim();
    if (!colName) throw new Error('请填写目标列名');
    if (!opts.sep) throw new Error('请填写连接分隔符');
    const colIdx = headers.indexOf(colName);
    if (colIdx < 0) {
      throw new Error(`未在表头中找到列「${colName}」。当前表头：${headers.join('、')}`);
    }

    const groupCols = opts.groupCols
      .replace(/，/g, ',')
      .split(',')
      .map((s) => s.trim())
      .filter(Boolean);

    if (groupCols.length > 0) {
      for (const name of groupCols) {
        if (!headers.includes(name)) {
          throw new Error(`未在表头中找到分组列「${name}」。当前表头：${headers.join('、')}`);
        }
        if (name === colName) {
          throw new Error(`分组列不能包含目标列「${colName}」`);
        }
      }
      const groupIdx = groupCols.map((c) => headers.indexOf(c));
      const naturalCollator = new Intl.Collator('zh-Hans-CN', { numeric: true, sensitivity: 'base' });
      logs.push(
        `一级合并（按指定分组列 [${groupCols.join('、')}]）：目标列「${colName}」去重并自然排序拼接，连接符 ${JSON.stringify(opts.sep)}`
      );

      const groups = new Map<
        string,
        { keyVals: Array<string | number | boolean | null>; items: Set<string> }
      >();
      for (const row of dataRows) {
        const keyVals = groupIdx.map((g) => row.values[g]);
        const keyStr = JSON.stringify(keyVals.map((v) => (v === null ? '' : String(v).trim())));
        const val = row.values[colIdx] === null ? '' : String(row.values[colIdx]).trim();
        let grp = groups.get(keyStr);
        if (!grp) {
          grp = { keyVals, items: new Set() };
          groups.set(keyStr, grp);
        }
        if (val) grp.items.add(val);
      }

      const outRows: Array<Array<string | number | boolean | null>> = [[...groupCols, colName]];
      for (const grp of groups.values()) {
        const sortedItems = Array.from(grp.items).sort((a, b) => naturalCollator.compare(a, b));
        outRows.push([...grp.keyVals, sortedItems.join(opts.sep)]);
      }

      const outWb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(outWb, buildLiteralSheet(outRows), sheetName);
      logs.push(`完成：${dataRows.length} 行 → ${outRows.length - 1} 行`);
      return {
        workbook: outWb,
        outputFileName: `${stem}_分组汇总.xlsx`,
        srcRows: dataRows.length,
        outRows: outRows.length - 1,
        logs,
        sheetPreviews: [makePreview(sheetName, outRows)],
      };
    }

    logs.push(`一级合并：目标列「${colName}」（第 ${colIdx + 1} 列），连接符 ${JSON.stringify(opts.sep)}`);

    const groups = new Map<
      string,
      { keyParts: Array<string | number | boolean | null>; items: string[] }
    >();
    for (const row of dataRows) {
      const keyParts = [...row.values.slice(0, colIdx), ...row.values.slice(colIdx + 1)];
      const keyStr = JSON.stringify(keyParts);
      const val = row.values[colIdx] === null ? '' : String(row.values[colIdx]).trim();
      let grp = groups.get(keyStr);
      if (!grp) {
        grp = { keyParts, items: [] };
        groups.set(keyStr, grp);
      }
      if (val) grp.items.push(val);
    }

    const outRows: Array<Array<string | number | boolean | null>> = [headers];
    for (const grp of groups.values()) {
      const merged = grp.items.join(opts.sep);
      outRows.push([...grp.keyParts.slice(0, colIdx), merged, ...grp.keyParts.slice(colIdx)]);
    }

    const outWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(outWb, buildLiteralSheet(outRows), sheetName);
    logs.push(`完成：${dataRows.length} 行 → ${outRows.length - 1} 行`);
    return {
      workbook: outWb,
      outputFileName: `${stem}_合并.xlsx`,
      srcRows: dataRows.length,
      outRows: outRows.length - 1,
      logs,
      sheetPreviews: [makePreview(sheetName, outRows)],
    };
  }

  // ==================== 4. 两级合并 ====================
  if (opts.mode === 'merge_two') {
    if (!opts.sep) throw new Error('请填写列表分隔符');
    if (!opts.prefixSep) throw new Error('请填写前缀分隔符');
    if (!opts.groupSep) throw new Error('请填写分组分隔符');
    const prefixCol = opts.prefixName.trim() || '分标';
    const itemCol = opts.itemName.trim() || '包号';
    const outColName = opts.colName.trim() || '分标包号';
    const groupCols = opts.groupCols
      .replace(/，/g, ',')
      .split(',')
      .map((s) => s.trim())
      .filter(Boolean);
    if (groupCols.length === 0) {
      throw new Error('请至少指定一个分组列（如：供应商名称）');
    }

    for (const name of [...groupCols, prefixCol, itemCol]) {
      if (!headers.includes(name)) {
        throw new Error(`未在表头中找到列「${name}」。当前表头：${headers.join('、')}`);
      }
    }
    const groupIdx = groupCols.map((c) => headers.indexOf(c));
    const pi = headers.indexOf(prefixCol);
    const ii = headers.indexOf(itemCol);
    if (pi === ii) throw new Error('分标列与包号列不能是同一列');

    const insertAt = resolveInsertPos(opts.newColPos, headers);
    logs.push(
      `两级合并：分组列 [${groupCols.join('、')}] →「${prefixCol}」+「${itemCol}」重组为新列「${outColName}」（第 ${insertAt + 1} 列）`
    );

    const newHeader = [...headers];
    newHeader.splice(insertAt, 0, outColName);

    const blocks = new Map<
      string,
      { firstRow: Array<string | number | boolean | null>; prefixMap: Map<string, string[]> }
    >();
    for (const row of dataRows) {
      const keyStr = JSON.stringify(groupIdx.map((g) => row.values[g]));
      const prefix = row.values[pi] === null ? '' : String(row.values[pi]).trim();
      const item = row.values[ii] === null ? '' : String(row.values[ii]).trim();
      let blk = blocks.get(keyStr);
      if (!blk) {
        blk = { firstRow: row.values, prefixMap: new Map() };
        blocks.set(keyStr, blk);
      }
      let items = blk.prefixMap.get(prefix);
      if (!items) {
        items = [];
        blk.prefixMap.set(prefix, items);
      }
      if (item) items.push(item);
    }

    const outRows: Array<Array<string | number | boolean | null>> = [newHeader];
    for (const blk of blocks.values()) {
      const parts: string[] = [];
      for (const [prefix, items] of blk.prefixMap.entries()) {
        const itemsStr = items.join(opts.sep);
        if (prefix) {
          parts.push(itemsStr ? `${prefix}${opts.prefixSep}${itemsStr}` : prefix);
        } else if (itemsStr) {
          parts.push(itemsStr);
        }
      }
      const combined = parts.filter(Boolean).join(opts.groupSep);
      const nextRow = [...blk.firstRow];
      nextRow.splice(insertAt, 0, combined);
      outRows.push(nextRow);
    }

    const outWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(outWb, buildLiteralSheet(outRows), sheetName);
    logs.push(`完成：${dataRows.length} 行 → ${outRows.length - 1} 行`);
    return {
      workbook: outWb,
      outputFileName: `${stem}_合并.xlsx`,
      srcRows: dataRows.length,
      outRows: outRows.length - 1,
      logs,
      sheetPreviews: [makePreview(sheetName, outRows)],
    };
  }

  // ==================== 5 & 6. 联系人展开并检查 / 已展开联系人检查 ====================
  const isExpanded = opts.mode === 'contacts_existing';
  if (!opts.sep || !opts.prefixSep || opts.sep === opts.prefixSep) {
    throw new Error('列表分隔符和前缀分隔符必须非空且不同');
  }

  const locate = (name: string, explicit?: string): number => {
    const exp = explicit?.trim();
    const aliases = CONTACT_ALIASES[name] || [name];
    let candidates = exp ? [exp] : aliases;
    let found: number[] = [];
    headers.forEach((h, i) => {
      if (candidates.includes(h)) found.push(i);
    });
    // 若传入的是默认列名（如“分标”“包号”）且表头用的是其他常见别名（如“分标名称”），自动回退到别名库匹配
    if (found.length === 0 && exp && aliases.includes(exp)) {
      candidates = aliases;
      headers.forEach((h, i) => {
        if (candidates.includes(h)) found.push(i);
      });
    }
    if (found.length !== 1) {
      throw new Error(`字段「${name}」缺失或有歧义；请检查表头或手动指定列名。当前表头：${headers.join('、')}`);
    }
    return found[0];
  };

  const supplierIdx = locate('供应商', opts.supplierCol || undefined);
  const contactIdxs = CONTACT_FIELDS.map((f) => locate(f));
  let sectionIdx = -1;
  let packageIdx = -1;
  let compoundIdx = -1;
  let inputIdxs: number[] = [];

  if (isExpanded) {
    sectionIdx = locate('分标', opts.prefixName || undefined);
    packageIdx = locate('包号', opts.itemName || undefined);
    inputIdxs = [supplierIdx, sectionIdx, packageIdx, ...contactIdxs];
  } else {
    const colName = opts.colName.trim() || '分标包号';
    compoundIdx = locate(colName, colName);
    inputIdxs = [supplierIdx, compoundIdx, ...contactIdxs];
  }

  if (new Set(inputIdxs).size !== inputIdxs.length) {
    throw new Error('关键字段不能使用同一列');
  }

  interface ContactRec {
    key: [string, string, string];
    keyStr: string;
    contact: string[];
    rowNum: number;
    raw: Array<string | number | boolean | null>;
    reasons: string[];
  }

  const records: ContactRec[] = [];
  let inputCount = 0;
  let skippedBlank = 0;

  for (const row of dataRows) {
    if (row.values.every((v) => v === null || v === '')) {
      skippedBlank++;
      continue;
    }
    inputCount++;
    const reasons: string[] = [];
    if (inputIdxs.some((idx) => Boolean(row.cells[idx]?.f) || row.cells[idx]?.t === 'e')) {
      reasons.push('关键字段含公式或错误值，需核实原值');
    }

    const supplier = cellToText(row.cells[supplierIdx]);
    const contact = contactIdxs.map((idx) => cellToText(row.cells[idx]));

    let section = '';
    let packages: string[] = [];
    if (isExpanded) {
      section = cellToText(row.cells[sectionIdx]);
      packages = [cellToText(row.cells[packageIdx])];
    } else {
      const compound = cellToText(row.cells[compoundIdx]);
      const [secRaw, hasMarker, remainder] = splitFirst(compound, opts.prefixSep);
      section = secRaw.trim();
      packages = remainder.split(opts.sep).map((p) => p.trim());
      if (!hasMarker || !section || packages.some((p) => !p)) {
        reasons.push('分标包号结构不完整');
      }
      if (remainder.includes(opts.prefixSep) || compound.includes(';') || compound.includes('；')) {
        reasons.push('包含多标段或不支持的分组结构，需核实');
      }
    }

    for (const pkg of packages) {
      const rowReasons = [...reasons];
      if (!section || !pkg || !supplier) {
        rowReasons.push('标段、包号或供应商缺失');
      }
      const key: [string, string, string] = [section, pkg, supplier];
      records.push({
        key,
        keyStr: JSON.stringify(key),
        contact,
        rowNum: row.rowNum,
        raw: row.values,
        reasons: rowReasons,
      });
    }
  }

  const groups = new Map<string, ContactRec[]>();
  for (const rec of records) {
    let list = groups.get(rec.keyStr);
    if (!list) {
      list = [];
      groups.set(rec.keyStr, list);
    }
    list.push(rec);
  }

  const standard = ['标段', '包号', '供应商', ...CONTACT_FIELDS];
  const safeRows: Array<Array<string | number | boolean | null>> = [
    [...standard, '来源文件', '来源工作表', '来源行号'],
  ];
  const auditHeader = [
    ...standard,
    '检查状态',
    '原因',
    '来源文件',
    '来源工作表',
    '来源行号',
    ...headers.map((h) => `原始:${h}`),
  ];
  const reviewRows: Array<Array<string | number | boolean | null>> = [auditHeader];
  const detailRows: Array<Array<string | number | boolean | null>> = [auditHeader];

  const stats: ContactsStats = {
    input_rows: inputCount,
    skipped_blank_rows: skippedBlank,
    expanded_rows: records.length,
    safe_keys: 0,
    conflict_keys: 0,
    review_rows: 0,
    duplicate_rows: 0,
  };

  for (const members of groups.values()) {
    const conflictFields: string[] = [];
    CONTACT_FIELDS.forEach((field, i) => {
      const distinct = new Set(members.map((m) => m.contact[i]));
      if (distinct.size > 1) conflictFields.push(field);
    });
    if (conflictFields.length > 0) stats.conflict_keys++;

    const blocked = conflictFields.length > 0 || members.some((m) => m.reasons.length > 0);
    if (!blocked) {
      stats.safe_keys++;
      stats.duplicate_rows += members.length - 1;
      const uniqueRowNums = Array.from(new Set(members.map((m) => m.rowNum)));
      safeRows.push([
        ...members[0].key,
        ...members[0].contact,
        opts.fileName,
        sheetName,
        uniqueRowNums.join(','),
      ]);
    }

    for (const rec of members) {
      const recReasons = [...rec.reasons];
      if (conflictFields.length > 0) {
        recReasons.push(`同键信息不一致：${conflictFields.join('、')}`);
      }
      if (blocked && recReasons.length === 0) {
        recReasons.push('同键其他记录需核实，整组暂停匹配');
      }
      const status = blocked ? '待核实' : members.length > 1 ? '一致重复' : '可匹配';
      const values: Array<string | number | boolean | null> = [
        ...rec.key,
        ...rec.contact,
        status,
        recReasons.join('；'),
        opts.fileName,
        sheetName,
        rec.rowNum,
        ...rec.raw,
      ];
      detailRows.push(values);
      if (blocked) {
        stats.review_rows++;
        reviewRows.push(values);
      }
    }
  }

  const summaryRows: Array<Array<string | number | boolean | null>> = [
    ['指标', '数量或说明'],
    ['input_rows（有效输入行）', stats.input_rows],
    ['skipped_blank_rows（跳过空行）', stats.skipped_blank_rows],
    ['expanded_rows（展开后记录数）', stats.expanded_rows],
    ['safe_keys（可匹配键）', stats.safe_keys],
    ['conflict_keys（冲突键）', stats.conflict_keys],
    ['review_rows（待核实行）', stats.review_rows],
    ['duplicate_rows（一致重复合并行）', stats.duplicate_rows],
    ['使用说明', '仅「可匹配数据」供下游使用；「待核实记录」更正源数据后重跑。'],
    ['空值规则', '五项联系字段允许空值；空值与非空值视为不一致。不验证联系信息真实性。'],
    ['数字限制', '源文件已按数字保存而丢失的前导零无法恢复；请在源表使用文本。'],
  ];

  const outWb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(outWb, buildLiteralSheet(safeRows, { forceText: true, autoFilter: true }), '可匹配数据');
  XLSX.utils.book_append_sheet(outWb, buildLiteralSheet(reviewRows, { forceText: true, autoFilter: true }), '待核实记录');
  XLSX.utils.book_append_sheet(outWb, buildLiteralSheet(detailRows, { forceText: true, autoFilter: true }), '展开明细');
  XLSX.utils.book_append_sheet(outWb, buildLiteralSheet(summaryRows, { forceText: true, autoFilter: true }), '检查摘要');

  logs.push(
    `可匹配键 ${stats.safe_keys}；冲突键 ${stats.conflict_keys}；待核实行 ${stats.review_rows}；一致重复 ${stats.duplicate_rows}`
  );
  logs.push('下游只能使用「可匹配数据」工作表；检查报告不代表联系方式真实有效。');

  return {
    workbook: outWb,
    outputFileName: `${stem}_联系人检查.xlsx`,
    srcRows: stats.input_rows,
    outRows: stats.expanded_rows,
    logs,
    sheetPreviews: [
      makePreview('可匹配数据', safeRows),
      makePreview('待核实记录', reviewRows),
      makePreview('展开明细', detailRows),
      makePreview('检查摘要', summaryRows),
    ],
    contactsStats: stats,
  };
}
