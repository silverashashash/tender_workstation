import * as XLSX from 'xlsx';

export interface SheetInfoItem {
  name: string;
  hidden: boolean;
  isEmpty: boolean;
  rowCount: number;
  colCount: number;
}

export interface SheetMergeOptions {
  names?: string[];
  headerRow?: number; // 1-based
  startRow?: number; // 1-based
  fillMerged?: boolean;
  source?: boolean;
  alignHeaders?: boolean;
}

export interface SheetMergeCount {
  name: string;
  count: number;
}

export interface SheetMergeResult {
  workbook: XLSX.WorkBook;
  rows: Array<Array<string | number | boolean | null>>;
  counts: SheetMergeCount[];
  warnings: string[];
}

const isBlankCell = (c?: XLSX.CellObject | null): boolean =>
  !c || ((c.v === undefined || c.v === null || c.v === '') && !c.f);

export function getWorkbookSheetInfos(wb: XLSX.WorkBook): SheetInfoItem[] {
  return wb.SheetNames.map((name, i) => {
    const hidden = Boolean(wb.Workbook?.Sheets?.[i]?.Hidden);
    const sheet = wb.Sheets[name];
    const ref = sheet?.['!ref'];
    if (!ref) {
      return { name, hidden, isEmpty: true, rowCount: 0, colCount: 0 };
    }
    const range = XLSX.utils.decode_range(ref);
    return {
      name,
      hidden,
      isEmpty: false,
      rowCount: range.e.r + 1,
      colCount: range.e.c + 1,
    };
  });
}

export function getRawSheetPreview(
  wb: XLSX.WorkBook,
  sheetName: string,
  maxRows = 15,
  maxCols = 30
): { colHeaders: string[]; rows: Array<{ rowNum: number; cells: string[] }>; isEmpty: boolean } {
  const sheet = wb.Sheets[sheetName];
  if (!sheet?.['!ref']) {
    return { colHeaders: [], rows: [], isEmpty: true };
  }
  const range = XLSX.utils.decode_range(sheet['!ref']);
  const colCount = Math.min(range.e.c + 1, maxCols);
  const colHeaders = Array.from({ length: colCount }, (_, c) => XLSX.utils.encode_col(c));
  const rowLimit = Math.min(range.e.r, maxRows - 1);
  const rows: Array<{ rowNum: number; cells: string[] }> = [];

  for (let r = 0; r <= rowLimit; r++) {
    const cells: string[] = [];
    for (let c = 0; c < colCount; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })] as XLSX.CellObject | undefined;
      cells.push(cell ? XLSX.utils.format_cell(cell) : '');
    }
    rows.push({ rowNum: r + 1, cells });
  }
  return { colHeaders, rows, isEmpty: false };
}

/**
 * Merge multiple worksheets from a single workbook into one consolidated worksheet (`合并结果`),
 * preserving cell types, number formats, cached formula values, date1904 mode, and optional merged-cell fill.
 */
export function mergeWorkbook(input: XLSX.WorkBook, options: SheetMergeOptions = {}): SheetMergeResult {
  const {
    headerRow = 1,
    startRow = 2,
    fillMerged = true,
    source = true,
    alignHeaders = true,
  } = options;

  if (
    !Number.isInteger(headerRow) ||
    !Number.isInteger(startRow) ||
    headerRow < 1 ||
    startRow <= headerRow ||
    startRow > 1048576
  ) {
    throw new Error('请填写有效行号：数据起始行必须大于表头行，且不超过 1048576。');
  }

  const chosen = options.names ?? input.SheetNames;
  if (!chosen.length) {
    throw new Error('请至少选择一个工作表。');
  }
  for (const name of chosen) {
    if (!input.SheetNames.includes(name)) {
      throw new Error(`工作表不存在：${name}`);
    }
  }

  // Always preserve original workbook sheet order
  const names = input.SheetNames.filter((n) => chosen.includes(n));
  const rows: Array<Array<string | number | boolean | null>> = [];
  const counts: SheetMergeCount[] = [];
  const out: XLSX.WorkSheet = {};
  const warnings: string[] = [];

  let headers: Array<string | number | boolean | null> | undefined;
  let width = 0;
  let outputRow = 0;
  let scanned = 0;

  const prepared = names.map((name) => {
    const sheet = input.Sheets[name];
    if (!sheet?.['!ref']) {
      throw new Error(`${name} 是空工作表，请取消勾选。`);
    }
    const range = XLSX.utils.decode_range(sheet['!ref']);
    if (headerRow - 1 > range.e.r) {
      throw new Error(`${name} 没有第 ${headerRow} 行表头。`);
    }
    let last = -1;
    for (let c = 0; c <= range.e.c; c++) {
      if (!isBlankCell(sheet[XLSX.utils.encode_cell({ r: headerRow - 1, c })])) {
        last = c;
      }
    }
    if (last < 0) {
      throw new Error(`${name} 的第 ${headerRow} 行表头为空，请调整表头行。`);
    }
    const labels = Array.from({ length: last + 1 }, (_, c) => {
      const cell = sheet[XLSX.utils.encode_cell({ r: headerRow - 1, c })] as XLSX.CellObject | undefined;
      return (cell?.v as string | number | boolean | undefined) ?? null;
    });

    if (alignHeaders) {
      if (labels.some((v) => v === null || String(v).trim() === '')) {
        throw new Error(`${name} 的表头有空白列名，无法按列名对齐。`);
      }
      if (new Set(labels.map(String)).size !== labels.length) {
        throw new Error(`${name} 的表头有重复列名，无法按列名对齐。`);
      }
    }
    return { name, labels };
  });

  if (alignHeaders) {
    headers = [...new Set(prepared.flatMap((p) => p.labels))];
    width = headers.length;
    if (width + (source ? 1 : 0) > 16384) {
      throw new Error('列数超出 Excel 限制。');
    }
    headers.forEach((v, c) => {
      out[XLSX.utils.encode_cell({ r: 0, c })] = { t: 's', v: String(v) };
    });
    if (source) {
      out[XLSX.utils.encode_cell({ r: 0, c: width })] = { t: 's', v: '来源 sheet' };
    }
    rows.push(source ? [...headers, '来源 sheet'] : [...headers]);
    outputRow = 1;

    for (const p of prepared) {
      const missing = headers.filter((h) => !p.labels.includes(h));
      if (missing.length) {
        warnings.push(`${p.name} 缺少「${missing.join('、')}」，已留空`);
      }
    }
  }

  function copyCell(
    cell: XLSX.CellObject | undefined,
    name: string,
    address: string
  ): XLSX.CellObject | null {
    if (cell?.f && (cell.v === undefined || cell.v === null)) {
      throw new Error(
        `${name} 的 ${address} 公式没有缓存结果，请在 WPS / Excel 中重新计算并保存后再试。`
      );
    }
    if (isBlankCell(cell) || !cell) return null;
    const copied: XLSX.CellObject = { t: cell.t, v: cell.v };
    if (cell.z) copied.z = cell.z;
    if (cell.t === 'e') {
      warnings.push(`${name}!${address} 含原始错误值`);
    }
    return copied;
  }

  for (const name of names) {
    const sheet = input.Sheets[name];
    if (!sheet?.['!ref']) {
      throw new Error(`${name} 是空工作表，请取消勾选。`);
    }
    const range = XLSX.utils.decode_range(sheet['!ref']);
    if (headerRow - 1 > range.e.r) {
      throw new Error(`${name} 没有第 ${headerRow} 行表头。`);
    }
    const merges = sheet['!merges'] || [];

    const getCell = (
      r: number,
      c: number,
      fill: boolean
    ): [XLSX.CellObject | undefined, string] => {
      const address = XLSX.utils.encode_cell({ r, c });
      if (fill) {
        const m = merges.find(
          (rng) => r >= rng.s.r && r <= rng.e.r && c >= rng.s.c && c <= rng.e.c
        );
        if (m) {
          const originAddr = XLSX.utils.encode_cell(m.s);
          return [sheet[originAddr], originAddr];
        }
      }
      return [sheet[address], address];
    };

    // Ignore formatting-only cells when determining the header width
    let last = -1;
    for (let c = 0; c <= range.e.c; c++) {
      if (!isBlankCell(sheet[XLSX.utils.encode_cell({ r: headerRow - 1, c })])) {
        last = c;
      }
    }
    if (last < 0) {
      throw new Error(`${name} 的第 ${headerRow} 行表头为空，请调整表头行。`);
    }

    const hCells: Array<XLSX.CellObject | null> = [];
    for (let c = 0; c <= last; c++) {
      const [cell, address] = getCell(headerRow - 1, c, false);
      hCells.push(copyCell(cell, name, address));
    }
    const labels = hCells.map(
      (c) => (c?.v as string | number | boolean | undefined) ?? null
    );

    if (!headers) {
      headers = labels;
      width = hCells.length;
      if (width + (source ? 1 : 0) > 16384) {
        throw new Error('列数超出 Excel 限制。');
      }
      hCells.forEach((c, i) => {
        if (c) out[XLSX.utils.encode_cell({ r: 0, c: i })] = c;
      });
      if (source) {
        out[XLSX.utils.encode_cell({ r: 0, c: width })] = { t: 's', v: '来源 sheet' };
      }
      rows.push(source ? [...labels, '来源 sheet'] : labels);
      outputRow = 1;
    } else if (!alignHeaders && JSON.stringify(headers) !== JSON.stringify(labels)) {
      throw new Error(
        `${name} 的表头与 ${names[0]} 不一致，请检查第 ${headerRow} 行的列名和顺序；也可开启按列名对齐。`
      );
    }

    const mapping = alignHeaders
      ? headers.map((h) => labels.indexOf(h))
      : headers.map((_, i) => i);

    const area = Math.max(0, range.e.r - startRow + 2) * (range.e.c + 1);
    scanned += area;
    if (scanned > 3000000) {
      throw new Error('所选范围超过 300 万个单元格，请减少工作表或清理多余格式后重试。');
    }

    let count = 0;
    for (let r = startRow - 1; r <= range.e.r; r++) {
      for (let c = labels.length; c <= range.e.c; c++) {
        if (!isBlankCell(sheet[XLSX.utils.encode_cell({ r, c })])) {
          throw new Error(`${name} 第 ${r + 1} 行存在超出表头范围的数据，请检查表头。`);
        }
      }

      const cells: Array<XLSX.CellObject | null> = [];
      for (let c = 0; c < width; c++) {
        if (mapping[c] < 0) {
          cells.push(null);
          continue;
        }
        const [cell, address] = getCell(r, mapping[c], fillMerged);
        cells.push(copyCell(cell, name, address));
      }
      if (cells.every((c) => !c)) continue;
      if (outputRow >= 1048576) {
        throw new Error('合并结果超出 Excel 最大行数。');
      }

      cells.forEach((cell, c) => {
        if (cell) out[XLSX.utils.encode_cell({ r: outputRow, c })] = cell;
      });
      if (source) {
        out[XLSX.utils.encode_cell({ r: outputRow, c: width })] = { t: 's', v: name };
      }
      const values = cells.map(
        (c) => (c?.v as string | number | boolean | undefined) ?? null
      );
      rows.push(source ? [...values, name] : values);
      count++;
      outputRow++;
    }
    counts.push({ name, count });
  }

  out['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: outputRow - 1, c: width - 1 + (source ? 1 : 0) },
  });
  out['!autofilter'] = { ref: out['!ref'] };
  out['!cols'] = Array.from({ length: width + (source ? 1 : 0) }, (_, c) => ({
    wch: Math.min(
      48,
      Math.max(
        12,
        ...rows
          .slice(0, 100)
          .map((r) => String(r[c] ?? '').replace(/[^\x00-\xff]/g, 'xx').length + 2)
      )
    ),
  }));

  const workbook = XLSX.utils.book_new();
  workbook.Workbook = { WBProps: { date1904: Boolean(input.Workbook?.WBProps?.date1904) } };
  XLSX.utils.book_append_sheet(workbook, out, '合并结果');

  return {
    workbook,
    rows,
    counts,
    warnings: [...new Set(warnings)],
  };
}
