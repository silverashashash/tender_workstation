import { useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  AlertTriangle,
  CheckCircle2,
  Download,
  Eye,
  FileSpreadsheet,
  Info,
  Layers,
  Play,
  RotateCcw,
  Upload,
} from 'lucide-react';
import {
  getRawSheetPreview,
  getWorkbookSheetInfos,
  mergeWorkbook,
  type SheetInfoItem,
  type SheetMergeResult,
} from '../utils/sheetMerger';

const MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024; // 50 MB

export function SheetMergerPanel() {
  const readVersionRef = useRef(0);

  const [fileName, setFileName] = useState('');
  const [fileSize, setFileSize] = useState(0);
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [sheetInfos, setSheetInfos] = useState<SheetInfoItem[]>([]);
  const [selectedNames, setSelectedNames] = useState<Set<string>>(new Set());
  const [previewSheetName, setPreviewSheetName] = useState('');

  // Merge options
  const [headerRowInput, setHeaderRowInput] = useState('1');
  const [startRowInput, setStartRowInput] = useState('2');
  const [fillMerged, setFillMerged] = useState(true);
  const [includeSource, setIncludeSource] = useState(true);
  const [alignHeaders, setAlignHeaders] = useState(true);

  // Status & Result
  const [reading, setReading] = useState(false);
  const [statusMsg, setStatusMsg] = useState('');
  const [errorMsg, setErrorMsg] = useState('');
  const [result, setResult] = useState<SheetMergeResult | null>(null);

  const parsedHeaderRow = Number(headerRowInput);
  const parsedStartRow = Number(startRowInput);

  function invalidateResult() {
    setResult(null);
    setStatusMsg('');
    setErrorMsg('');
  }

  async function handleFileSelect(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;

    const token = ++readVersionRef.current;
    setWorkbook(null);
    setSheetInfos([]);
    setSelectedNames(new Set());
    setPreviewSheetName('');
    setFileName('');
    setFileSize(0);
    invalidateResult();

    if (!/\.xlsx?$/i.test(file.name)) {
      setErrorMsg('请选择 Excel 文件（支持 .xlsx / .xls，推荐使用 .xlsx）。');
      return;
    }
    if (file.size > MAX_FILE_SIZE_BYTES) {
      setErrorMsg('文件超过 50 MB，请拆分后处理。');
      return;
    }

    setReading(true);
    try {
      const data = await file.arrayBuffer();
      if (token !== readVersionRef.current) return;
      const parsed = XLSX.read(data, { type: 'array', cellNF: true, cellFormula: true });
      if (!parsed.SheetNames.length) {
        throw new Error('文件中没有工作表。');
      }
      const infos = getWorkbookSheetInfos(parsed);
      const defaultSelected = new Set(
        infos.filter((s) => !s.hidden && !s.isEmpty).map((s) => s.name)
      );
      setWorkbook(parsed);
      setFileName(file.name);
      setFileSize(file.size);
      setSheetInfos(infos);
      setSelectedNames(defaultSelected);
      setPreviewSheetName(infos[0]?.name || '');
    } catch (err) {
      setErrorMsg(`无法读取文件：${(err as Error).message}`);
    } finally {
      if (token === readVersionRef.current) {
        setReading(false);
      }
    }
  }

  function handleReset() {
    readVersionRef.current++;
    setWorkbook(null);
    setSheetInfos([]);
    setSelectedNames(new Set());
    setPreviewSheetName('');
    setFileName('');
    setFileSize(0);
    invalidateResult();
  }

  function toggleSheet(name: string) {
    invalidateResult();
    setSelectedNames((prev) => {
      const next = new Set(prev);
      if (next.has(name)) next.delete(name);
      else next.add(name);
      return next;
    });
  }

  function selectAllSheets(checked: boolean) {
    invalidateResult();
    if (checked) {
      setSelectedNames(new Set(sheetInfos.map((s) => s.name)));
    } else {
      setSelectedNames(new Set());
    }
  }

  function applyRowPreset(hRow: number, sRow: number) {
    invalidateResult();
    setHeaderRowInput(String(hRow));
    setStartRowInput(String(sRow));
  }

  function handleHeaderRowChange(val: string) {
    invalidateResult();
    setHeaderRowInput(val);
    const n = Number(val);
    if (Number.isInteger(n) && n >= 1) {
      const curStart = Number(startRowInput);
      if (!Number.isInteger(curStart) || curStart <= n) {
        setStartRowInput(String(n + 1));
      }
    }
  }

  function handleMerge() {
    if (!workbook) return;
    invalidateResult();
    const chosen = sheetInfos.filter((s) => selectedNames.has(s.name)).map((s) => s.name);
    try {
      const merged = mergeWorkbook(workbook, {
        names: chosen,
        headerRow: Number(headerRowInput),
        startRow: Number(startRowInput),
        fillMerged,
        source: includeSource,
        alignHeaders,
      });
      setResult(merged);
      setStatusMsg(
        `合并完成：共合并 ${merged.counts.length} 张工作表、${merged.rows.length - 1} 条数据记录（表头保留 1 行），请预览后下载。`
      );
    } catch (err) {
      setErrorMsg((err as Error).message);
    }
  }

  function handleDownload() {
    if (!result) return;
    try {
      const outName = `${fileName.replace(/\.xlsx?$/i, '') || '工作表'}_合并.xlsx`;
      XLSX.writeFile(result.workbook, outName, { bookType: 'xlsx', compression: true });
      setStatusMsg(`已开始下载「${outName}」，请查看浏览器下载列表。`);
    } catch (err) {
      setErrorMsg(`下载失败：${(err as Error).message}`);
    }
  }

  const rawPreview = useMemo(() => {
    if (!workbook || !previewSheetName) return null;
    return getRawSheetPreview(workbook, previewSheetName, 15, 30);
  }, [workbook, previewSheetName]);

  const resultPreviewRows = useMemo(() => {
    if (!result) return [];
    const sheet = result.workbook.Sheets['合并结果'];
    if (!sheet) return [];
    return result.rows.slice(0, 31).map((row, r) =>
      row.map((_, c) => {
        const cell = sheet[XLSX.utils.encode_cell({ r, c })] as XLSX.CellObject | undefined;
        return cell ? XLSX.utils.format_cell(cell) : '';
      })
    );
  }, [result]);

  return (
    <div className="space-y-4">
      {/* Intro Banner */}
      <div className="rounded-xl border border-slate-200 bg-white px-4 py-3 flex flex-wrap items-center justify-between gap-2">
        <div className="flex items-start gap-2 text-xs text-slate-600">
          <Info className="h-4 w-4 shrink-0 text-[#3B5B7E] mt-0.5" />
          <span>
            将同一 Excel 工作簿中的多个工作表（Sheet）纵向合并为一张总表，仅保留一次表头。全程在浏览器本地处理，不上传服务器，原文件不会被修改。
          </span>
        </div>
        {workbook && (
          <button
            type="button"
            onClick={handleReset}
            className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-slate-50 px-2.5 py-1 text-xs font-medium text-slate-600 hover:border-[#A94442]/40 hover:text-[#A94442]"
          >
            <RotateCcw className="h-3.5 w-3.5" />
            <span>清空重选</span>
          </button>
        )}
      </div>

      {/* Main 2-column setup */}
      <div className="grid grid-cols-1 gap-4 lg:grid-cols-12">
        {/* Left: Step 1 — File & Sheet selection */}
        <div className="lg:col-span-6 rounded-xl border border-slate-200 bg-white p-4 flex flex-col gap-3">
          <div className="flex items-center justify-between gap-2">
            <div className="flex items-center gap-2">
              <span className="inline-flex h-5 w-5 items-center justify-center rounded-md bg-[#EBF2F7] text-[11px] font-semibold text-[#3B5B7E]">
                1
              </span>
              <h2 className="text-sm font-semibold text-slate-800">选择文件与工作表</h2>
            </div>
            <span className="text-[11px] text-slate-400 tabular-nums">
              {workbook ? `已选 ${selectedNames.size} / ${sheetInfos.length} 张` : '尚未选择文件'}
            </span>
          </div>

          <label className="flex flex-col items-center justify-center gap-1.5 rounded-lg border border-dashed border-[#3B5B7E]/35 bg-slate-50/70 px-4 py-5 text-center cursor-pointer hover:border-[#3B5B7E] hover:bg-[#EBF2F7]/30 transition-colors">
            <Upload className="h-5 w-5 text-[#3B5B7E]" />
            <div className="text-xs font-semibold text-slate-700">
              {reading
                ? '正在读取工作簿...'
                : fileName
                  ? fileName
                  : '点击选择需要合并的 Excel 工作簿'}
            </div>
            <div className="text-[11px] text-slate-400">
              {fileName
                ? `${(fileSize / 1024).toFixed(1)} KB · 点击可更换文件`
                : '支持 .xlsx / .xls（最大 50 MB）'}
            </div>
            <input
              type="file"
              accept=".xlsx,.xls"
              onChange={handleFileSelect}
              className="hidden"
              disabled={reading}
            />
          </label>

          {/* Sheet selection toolbar */}
          <div className="flex items-center justify-between gap-2">
            <div className="flex items-center gap-1.5">
              <button
                type="button"
                disabled={!workbook}
                onClick={() => selectAllSheets(true)}
                className="rounded-md border border-slate-200 bg-white px-2.5 py-1 text-[11px] font-medium text-slate-600 hover:bg-slate-50 disabled:opacity-40"
              >
                全选
              </button>
              <button
                type="button"
                disabled={!workbook}
                onClick={() => selectAllSheets(false)}
                className="rounded-md border border-slate-200 bg-white px-2.5 py-1 text-[11px] font-medium text-slate-600 hover:bg-slate-50 disabled:opacity-40"
              >
                全不选
              </button>
            </div>
            <span className="text-[11px] text-slate-400">隐藏表与空表默认不勾选</span>
          </div>

          {/* Sheet list */}
          <div className="max-h-60 overflow-y-auto rounded-lg border border-slate-200 divide-y divide-slate-100">
            {sheetInfos.length === 0 ? (
              <div className="py-10 text-center text-xs text-slate-400">
                选择 Excel 文件后，这里会列出全部工作表
              </div>
            ) : (
              sheetInfos.map((info) => {
                const checked = selectedNames.has(info.name);
                return (
                  <label
                    key={info.name}
                    className={`flex items-center justify-between gap-2 px-3 py-2 text-xs cursor-pointer transition-colors ${
                      checked ? 'bg-[#EBF2F7]/35' : 'hover:bg-slate-50'
                    }`}
                  >
                    <div className="flex items-center gap-2 min-w-0">
                      <input
                        type="checkbox"
                        checked={checked}
                        onChange={() => toggleSheet(info.name)}
                        className="rounded border-slate-300 cursor-pointer shrink-0"
                      />
                      <span className="truncate font-medium text-slate-700" title={info.name}>
                        {info.name}
                      </span>
                    </div>
                    <div className="flex items-center gap-1.5 shrink-0 text-[11px] text-slate-400 tabular-nums">
                      {info.hidden && (
                        <span className="rounded bg-slate-100 px-1.5 py-0.5 text-[10px] text-slate-500">
                          隐藏
                        </span>
                      )}
                      <span>{info.isEmpty ? '空表' : `${info.rowCount} 行`}</span>
                    </div>
                  </label>
                );
              })
            )}
          </div>
        </div>

        {/* Right: Step 2 — Merge rules */}
        <div className="lg:col-span-6 rounded-xl border border-slate-200 bg-white p-4 flex flex-col justify-between gap-4">
          <div className="space-y-3.5">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <div className="flex items-center gap-2">
                <span className="inline-flex h-5 w-5 items-center justify-center rounded-md bg-[#EBF2F7] text-[11px] font-semibold text-[#3B5B7E]">
                  2
                </span>
                <h2 className="text-sm font-semibold text-slate-800">设置合并规则</h2>
              </div>

              {/* Quick row presets */}
              <div className="flex items-center gap-1.5 text-[11px]">
                <span className="text-slate-400">快捷行号：</span>
                <button
                  type="button"
                  onClick={() => applyRowPreset(1, 2)}
                  className={`rounded border px-2 py-0.5 font-medium transition-colors ${
                    headerRowInput === '1' && startRowInput === '2'
                      ? 'border-[#3B5B7E] bg-[#EBF2F7] text-[#3B5B7E]'
                      : 'border-slate-200 bg-slate-50 text-slate-600 hover:bg-slate-100'
                  }`}
                >
                  标准表 (1/2)
                </button>
                <button
                  type="button"
                  onClick={() => applyRowPreset(6, 7)}
                  className={`rounded border px-2 py-0.5 font-medium transition-colors ${
                    headerRowInput === '6' && startRowInput === '7'
                      ? 'border-[#3B5B7E] bg-[#EBF2F7] text-[#3B5B7E]'
                      : 'border-slate-200 bg-slate-50 text-slate-600 hover:bg-slate-100'
                  }`}
                >
                  排序表 (6/7)
                </button>
              </div>
            </div>

            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  表头行（Excel 行号）
                </label>
                <input
                  type="number"
                  min={1}
                  max={1048575}
                  value={headerRowInput}
                  onChange={(e) => handleHeaderRowChange(e.target.value)}
                  className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs tabular-nums focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  数据起始行（从此行开始合并）
                </label>
                <input
                  type="number"
                  min={2}
                  max={1048576}
                  value={startRowInput}
                  onChange={(e) => {
                    invalidateResult();
                    setStartRowInput(e.target.value);
                  }}
                  className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs tabular-nums focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>
            </div>

            <div className="rounded-lg bg-slate-50 border border-slate-200/70 px-3 py-2 text-[11px] text-slate-500 leading-relaxed">
              行号与 Excel 左侧行号一致，所有选中的工作表共用同一组设置。可在下方「原表预览」中核对蓝色（表头行）与绿色（数据起始行）高亮位置。
            </div>

            <div className="space-y-2 pt-1 text-xs text-slate-700">
              <label className="flex items-center gap-2 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={fillMerged}
                  onChange={(e) => {
                    invalidateResult();
                    setFillMerged(e.target.checked);
                  }}
                  className="rounded border-slate-300 cursor-pointer"
                />
                <span>将合并单元格的内容自动填充到每条记录</span>
              </label>

              <label className="flex items-center gap-2 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={includeSource}
                  onChange={(e) => {
                    invalidateResult();
                    setIncludeSource(e.target.checked);
                  }}
                  className="rounded border-slate-300 cursor-pointer"
                />
                <span>在末列添加「来源 sheet」标记列</span>
              </label>

              <label className="flex items-center gap-2 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={alignHeaders}
                  onChange={(e) => {
                    invalidateResult();
                    setAlignHeaders(e.target.checked);
                  }}
                  className="rounded border-slate-300 cursor-pointer"
                />
                <span>按列名自动对齐（某表缺少字段时留空并提示；关闭则要求列序完全一致）</span>
              </label>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2.5 border-t border-slate-100 pt-3">
            <button
              type="button"
              onClick={handleMerge}
              disabled={!workbook || selectedNames.size === 0}
              className="inline-flex items-center gap-1.5 rounded-lg bg-[#3B5B7E] px-4 py-2 text-xs font-medium text-white hover:bg-[#2E4866] disabled:opacity-45 transition-colors"
            >
              <Play className="h-3.5 w-3.5" />
              <span>检查表头并合并</span>
            </button>

            <button
              type="button"
              onClick={handleDownload}
              disabled={!result}
              className="inline-flex items-center gap-1.5 rounded-lg border border-[#2F6F5E]/35 bg-[#ECFDF5] px-4 py-2 text-xs font-medium text-[#2F6F5E] hover:bg-[#D1FAE5] disabled:opacity-45 transition-colors"
            >
              <Download className="h-3.5 w-3.5" />
              <span>下载合并结果 (.xlsx)</span>
            </button>
          </div>
        </div>
      </div>

      {/* Error / Status / Warnings */}
      {errorMsg && (
        <div className="flex items-start gap-2 rounded-xl border border-[#A94442]/30 bg-[#FEF2F2] px-4 py-3 text-xs text-[#A94442]">
          <AlertTriangle className="h-4 w-4 shrink-0 mt-0.5" />
          <span className="whitespace-pre-line">{errorMsg}</span>
        </div>
      )}

      {statusMsg && !errorMsg && (
        <div className="flex items-start gap-2 rounded-xl border border-[#2F6F5E]/30 bg-[#ECFDF5] px-4 py-3 text-xs text-[#2F6F5E]">
          <CheckCircle2 className="h-4 w-4 shrink-0 mt-0.5" />
          <span>{statusMsg}</span>
        </div>
      )}

      {result && result.warnings.length > 0 && (
        <div className="rounded-xl border border-[#9A6B26]/30 bg-[#FEF9EC] px-4 py-3 text-xs text-[#9A6B26] space-y-1">
          <div className="font-semibold">合并提示（{result.warnings.length} 条）：</div>
          {result.warnings.map((w, i) => (
            <div key={i}>• {w}</div>
          ))}
        </div>
      )}

      {/* Raw Sheet Preview (first 15 rows with headerRow & startRow highlights) */}
      {workbook && rawPreview && (
        <div className="rounded-xl border border-slate-200 bg-white overflow-hidden">
          <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 bg-slate-50/70 px-4 py-2.5">
            <div className="flex items-center gap-2">
              <Eye className="h-4 w-4 text-[#3B5B7E]" />
              <h3 className="text-xs font-semibold text-slate-700">原表预览（前 15 行）</h3>
              <span className="inline-flex items-center gap-1 rounded bg-[#EBF2F7] px-1.5 py-0.5 text-[10px] font-medium text-[#3B5B7E]">
                蓝色 = 表头行（第 {Number.isInteger(parsedHeaderRow) ? parsedHeaderRow : '?'} 行）
              </span>
              <span className="inline-flex items-center gap-1 rounded bg-[#ECFDF5] px-1.5 py-0.5 text-[10px] font-medium text-[#2F6F5E]">
                绿色 = 起始数据行（第 {Number.isInteger(parsedStartRow) ? parsedStartRow : '?'} 行起）
              </span>
            </div>

            <div className="flex items-center gap-2 text-xs">
              <span className="text-slate-500">切换预览工作表：</span>
              <select
                value={previewSheetName}
                onChange={(e) => setPreviewSheetName(e.target.value)}
                className="rounded-md border border-slate-200 bg-white px-2.5 py-1 text-xs focus:border-[#3B5B7E] focus:outline-none max-w-[240px]"
              >
                {sheetInfos.map((s) => (
                  <option key={s.name} value={s.name}>
                    {s.name}
                  </option>
                ))}
              </select>
            </div>
          </div>

          <div className="max-h-80 overflow-auto">
            {rawPreview.isEmpty ? (
              <div className="py-8 text-center text-xs text-slate-400">当前工作表为空</div>
            ) : (
              <table className="w-full border-collapse text-left text-[11px]">
                <thead className="sticky top-0 bg-slate-100 text-slate-500 border-b border-slate-200">
                  <tr>
                    <th className="px-2.5 py-1.5 font-mono font-medium w-14 text-center border-r border-slate-200">
                      行号
                    </th>
                    {rawPreview.colHeaders.map((col) => (
                      <th
                        key={col}
                        className="px-2.5 py-1.5 font-mono font-medium whitespace-nowrap border-r border-slate-200/70"
                      >
                        {col}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {rawPreview.rows.map((row) => {
                    const isHeader = row.rowNum === parsedHeaderRow;
                    const isStart = row.rowNum === parsedStartRow;
                    const rowCls = isHeader
                      ? 'bg-[#EBF2F7] font-semibold text-[#3B5B7E]'
                      : isStart
                        ? 'bg-[#ECFDF5]/70 text-[#2F6F5E]'
                        : 'hover:bg-slate-50/70 text-slate-700';
                    return (
                      <tr key={row.rowNum} className={rowCls}>
                        <td className="px-2.5 py-1.5 font-mono text-center border-r border-slate-200 tabular-nums whitespace-nowrap">
                          {row.rowNum}
                        </td>
                        {row.cells.map((val, idx) => (
                          <td
                            key={idx}
                            className="px-2.5 py-1.5 max-w-[260px] truncate border-r border-slate-100"
                            title={val}
                          >
                            {val}
                          </td>
                        ))}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            )}
          </div>
        </div>
      )}

      {/* Merged Result Preview */}
      {result && (
        <div className="rounded-xl border border-slate-200 bg-white overflow-hidden">
          <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 bg-slate-50/70 px-4 py-2.5">
            <div className="flex items-center gap-2">
              <Layers className="h-4 w-4 text-[#2F6F5E]" />
              <h3 className="text-xs font-semibold text-slate-800">合并结果预览</h3>
              <span className="rounded-full bg-[#ECFDF5] px-2 py-0.5 text-[11px] font-medium text-[#2F6F5E] tabular-nums">
                {result.counts.length} 张工作表 · {result.rows.length - 1} 条记录
              </span>
            </div>
            <span className="text-[11px] text-slate-400">
              仅预览前 30 条记录，下载包含全部 {result.rows.length - 1} 条数据
            </span>
          </div>

          {/* Per-sheet counts */}
          <div className="flex flex-wrap gap-1.5 border-b border-slate-100 bg-slate-50/40 px-4 py-2">
            {result.counts.map((c) => (
              <span
                key={c.name}
                className="inline-flex items-center gap-1 rounded bg-white border border-slate-200 px-2 py-0.5 text-[11px] text-slate-600 tabular-nums"
              >
                <FileSpreadsheet className="h-3 w-3 text-[#3B5B7E]" />
                <span>{c.name}：</span>
                <strong className="text-slate-800">{c.count}</strong> 条
              </span>
            ))}
          </div>

          <div className="max-h-96 overflow-auto">
            <table className="w-full border-collapse text-left text-xs">
              {resultPreviewRows.length > 0 && (
                <thead className="sticky top-0 bg-slate-100 text-slate-600 border-b border-slate-200">
                  <tr>
                    {resultPreviewRows[0].map((h, i) => (
                      <th key={i} className="px-3 py-2 font-medium whitespace-nowrap">
                        {h}
                      </th>
                    ))}
                  </tr>
                </thead>
              )}
              <tbody className="divide-y divide-slate-100">
                {resultPreviewRows.length <= 1 ? (
                  <tr>
                    <td
                      colSpan={Math.max(1, resultPreviewRows[0]?.length || 1)}
                      className="py-8 text-center text-xs text-slate-400"
                    >
                      所选起始行以下无有效数据记录
                    </td>
                  </tr>
                ) : (
                  resultPreviewRows.slice(1).map((row, rIdx) => (
                    <tr key={rIdx} className="hover:bg-slate-50/70">
                      {row.map((val, cIdx) => (
                        <td
                          key={cIdx}
                          className="px-3 py-1.5 max-w-[280px] truncate text-slate-700"
                          title={val}
                        >
                          {val}
                        </td>
                      ))}
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}
