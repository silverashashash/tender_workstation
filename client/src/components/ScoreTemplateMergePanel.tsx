import { useState } from 'react';
import {
  AlertTriangle,
  Award,
  CheckCircle2,
  Download,
  FileSpreadsheet,
  Info,
  Play,
  RotateCcw,
  TableProperties,
  Trash2,
  Upload,
  X,
} from 'lucide-react';
import { api } from '../api/client';

type MergeResponse = Awaited<ReturnType<typeof api.runScoreTemplateMerge>>;

export function ScoreTemplateMergePanel() {
  const [scoreFiles, setScoreFiles] = useState<File[]>([]);
  const [selectedScoreIndices, setSelectedScoreIndices] = useState<Set<number>>(new Set());
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [successListFile, setSuccessListFile] = useState<File | null>(null);

  const [loading, setLoading] = useState(false);
  const [errorMsg, setErrorMsg] = useState('');
  const [result, setResult] = useState<MergeResponse | null>(null);
  const [activePreviewIdx, setActivePreviewIdx] = useState(0);

  function handleAddScoreFiles(e: React.ChangeEvent<HTMLInputElement>, replace = false) {
    const picked = Array.from(e.target.files || []).filter(
      (f) => /\.(xlsx|zip)$/i.test(f.name) && !f.name.startsWith('~$')
    );
    e.target.value = '';
    if (picked.length === 0) return;
    setErrorMsg('');
    setResult(null);

    if (replace) {
      setScoreFiles(picked);
      return;
    }
    setScoreFiles((prev) => {
      const seen = new Set(prev.map((f) => `${f.name}:${f.size}`));
      const next = [...prev];
      for (const f of picked) {
        const key = `${f.name}:${f.size}`;
        if (!seen.has(key)) {
          next.push(f);
          seen.add(key);
        }
      }
      return next;
    });
  }

  function removeScoreFile(idx: number) {
    setScoreFiles((prev) => prev.filter((_, i) => i !== idx));
    setSelectedScoreIndices(new Set());
    setResult(null);
  }

  function toggleScoreSelect(idx: number) {
    setSelectedScoreIndices((prev) => {
      const next = new Set(prev);
      if (next.has(idx)) next.delete(idx);
      else next.add(idx);
      return next;
    });
  }

  function toggleSelectAllScores() {
    if (selectedScoreIndices.size === scoreFiles.length) {
      setSelectedScoreIndices(new Set());
    } else {
      setSelectedScoreIndices(new Set(scoreFiles.map((_, i) => i)));
    }
  }

  function removeSelectedScoreFiles() {
    if (selectedScoreIndices.size === 0) return;
    setScoreFiles((prev) => prev.filter((_, i) => !selectedScoreIndices.has(i)));
    setSelectedScoreIndices(new Set());
    setResult(null);
  }

  function handleResetAll() {
    setScoreFiles([]);
    setSelectedScoreIndices(new Set());
    setTemplateFile(null);
    setSuccessListFile(null);
    setErrorMsg('');
    setResult(null);
    setActivePreviewIdx(0);
  }

  async function handleRun() {
    if (scoreFiles.length === 0) {
      setErrorMsg('请先选择至少一份专家打分表（支持多选 .xlsx 或 .zip 压缩包）');
      return;
    }
    setLoading(true);
    setErrorMsg('');
    setResult(null);
    setActivePreviewIdx(0);
    try {
      const res = await api.runScoreTemplateMerge({
        files: scoreFiles,
        template: templateFile,
        successList: successListFile,
      });
      setResult(res);
    } catch (err) {
      setErrorMsg((err as Error).message);
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="space-y-4">
      {/* Intro */}
      <div className="rounded-xl border border-slate-200 bg-white px-4 py-3 flex flex-wrap items-center justify-between gap-2">
        <div className="flex items-start gap-2 text-xs text-slate-600">
          <Info className="h-4 w-4 shrink-0 text-[#3B5B7E] mt-0.5" />
          <span>
            将多位专家分别填写的线下打分表按<strong>原始模板</strong>汇总合并为一张完整汇总表：自动识别模板初始分基线并忽略未调整/已清空行，按<strong>标段 + 包号 + 应答人</strong>对齐，不同专家分数冲突时自动以 <code className="rounded bg-slate-100 px-1 text-[11px]">A | B</code> 标黄提示；未传模板时自动按众数骨架启用旧版合并规则；可选上传《优秀名单》自动高亮对应应答人。
          </span>
        </div>
        {(scoreFiles.length > 0 || templateFile || successListFile) && (
          <button
            type="button"
            onClick={handleResetAll}
            disabled={loading}
            className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-slate-50 px-2.5 py-1 text-xs font-medium text-slate-600 hover:border-[#A94442]/40 hover:text-[#A94442]"
          >
            <RotateCcw className="h-3.5 w-3.5" />
            <span>清空重选</span>
          </button>
        )}
      </div>

      {/* Main Upload Grid */}
      <div className="grid grid-cols-1 gap-4 lg:grid-cols-12">
        {/* Left: Expert score files */}
        <div className="lg:col-span-7 rounded-xl border border-slate-200 bg-white p-4 space-y-3">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="flex items-center gap-2">
              <TableProperties className="h-4 w-4 text-[#3B5B7E]" />
              <h2 className="text-sm font-semibold text-slate-800">1. 专家打分表 *（支持多选或 ZIP）</h2>
              {scoreFiles.length > 0 && (
                <span className="rounded-full bg-[#EBF2F7] px-2 py-0.5 text-[11px] font-medium text-[#3B5B7E] tabular-nums">
                  已选 {scoreFiles.length} 份
                </span>
              )}
            </div>
            <div className="flex items-center gap-1.5">
              <label className="inline-flex cursor-pointer items-center gap-1 rounded-lg border border-[#3B5B7E]/35 bg-[#EBF2F7]/60 px-3 py-1.5 text-xs font-medium text-[#3B5B7E] hover:bg-[#EBF2F7] transition-colors">
                <Upload className="h-3.5 w-3.5" />
                <span>{scoreFiles.length > 0 ? '追加打分表' : '选择打分表 / ZIP'}</span>
                <input
                  type="file"
                  multiple
                  accept=".xlsx,.zip"
                  onChange={(e) => handleAddScoreFiles(e, false)}
                  className="hidden"
                  disabled={loading}
                />
              </label>
            </div>
          </div>

          {scoreFiles.length === 0 ? (
            <div className="rounded-lg border border-dashed border-slate-300 bg-slate-50/60 py-8 text-center text-xs text-slate-400">
              尚未选择专家打分表（支持多选 .xlsx 文件或打包好的 .zip 压缩包，自动跳过 ~$ 临时文件）
            </div>
          ) : (
            <div className="rounded-lg border border-slate-200 overflow-hidden">
              <div className="flex items-center justify-between gap-2 border-b border-slate-200 bg-slate-50/80 px-3 py-1.5 text-[11px]">
                <label className="inline-flex items-center gap-1.5 cursor-pointer select-none text-slate-600">
                  <input
                    type="checkbox"
                    checked={scoreFiles.length > 0 && selectedScoreIndices.size === scoreFiles.length}
                    onChange={toggleSelectAllScores}
                    disabled={loading}
                    className="h-3.5 w-3.5 rounded border-slate-300 accent-[#3B5B7E]"
                  />
                  <span>
                    {selectedScoreIndices.size > 0
                      ? `已勾选 ${selectedScoreIndices.size} / ${scoreFiles.length} 份`
                      : '全选'}
                  </span>
                </label>
                <button
                  type="button"
                  onClick={removeSelectedScoreFiles}
                  disabled={loading || selectedScoreIndices.size === 0}
                  className="inline-flex items-center gap-1 rounded px-2 py-0.5 text-[11px] font-medium text-[#A94442] hover:bg-[#FEF2F2] disabled:opacity-40 disabled:hover:bg-transparent transition-colors"
                >
                  <Trash2 className="h-3 w-3" />
                  <span>批量删除{selectedScoreIndices.size > 0 ? ` (${selectedScoreIndices.size})` : ''}</span>
                </button>
              </div>
              <div className="max-h-56 overflow-y-auto divide-y divide-slate-100">
                {scoreFiles.map((f, idx) => {
                  const checked = selectedScoreIndices.has(idx);
                  return (
                    <div
                      key={`${f.name}-${idx}`}
                      className={`flex items-center justify-between gap-2 px-3 py-1.5 text-xs transition-colors ${
                        checked ? 'bg-[#EBF2F7]/50' : 'hover:bg-slate-50'
                      }`}
                    >
                      <label className="flex items-center gap-2 min-w-0 flex-1 cursor-pointer">
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => toggleScoreSelect(idx)}
                          disabled={loading}
                          className="h-3.5 w-3.5 rounded border-slate-300 accent-[#3B5B7E] shrink-0"
                        />
                        <FileSpreadsheet className="h-3.5 w-3.5 text-[#3B5B7E] shrink-0" />
                        <span className="truncate text-slate-700" title={f.name}>
                          {f.name}
                        </span>
                        <span className="text-[10px] text-slate-400 tabular-nums shrink-0">
                          {(f.size / 1024).toFixed(0)} KB
                        </span>
                      </label>
                      <button
                        type="button"
                        onClick={() => removeScoreFile(idx)}
                        disabled={loading}
                        className="rounded p-0.5 text-slate-400 hover:text-[#A94442]"
                        title="移除"
                      >
                        <X className="h-3.5 w-3.5" />
                      </button>
                    </div>
                  );
                })}
              </div>
            </div>
          )}
        </div>

        {/* Right: Optional template + Optional success list + Run */}
        <div className="lg:col-span-5 rounded-xl border border-slate-200 bg-white p-4 flex flex-col justify-between gap-4">
          <div className="space-y-3.5">
            {/* 2. Original Template */}
            <div className="space-y-1.5">
              <label className="block text-xs font-semibold text-slate-800">
                2. 原始空白/含初始分模板（推荐）
              </label>
              <p className="text-[11px] text-slate-500 leading-relaxed">
                用于识别模板初始分基线与校验表头骨架；含预填初始分的新模板请务必上传。留空则用众数法按旧版规则合并。
              </p>
              <div className="flex items-center gap-2">
                <label className="flex-1 inline-flex cursor-pointer items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:border-[#3B5B7E] hover:text-[#3B5B7E] transition-colors min-w-0">
                  <FileSpreadsheet className="h-3.5 w-3.5 shrink-0" />
                  <span className="truncate">
                    {templateFile ? templateFile.name : '选择原始模板 (.xlsx)'}
                  </span>
                  <input
                    type="file"
                    accept=".xlsx"
                    onChange={(e) => {
                      const f = e.target.files?.[0] || null;
                      e.target.value = '';
                      setTemplateFile(f);
                      setResult(null);
                    }}
                    className="hidden"
                    disabled={loading}
                  />
                </label>
                {templateFile && (
                  <button
                    type="button"
                    onClick={() => {
                      setTemplateFile(null);
                      setResult(null);
                    }}
                    disabled={loading}
                    className="text-xs text-slate-400 hover:text-[#A94442] shrink-0"
                  >
                    清除
                  </button>
                )}
              </div>
            </div>

            {/* 3. Optional Success List */}
            <div className="space-y-1.5 border-t border-slate-100 pt-3">
              <label className="block text-xs font-semibold text-slate-800">
                3. 优秀名单文件（可选）
              </label>
              <p className="text-[11px] text-slate-500 leading-relaxed">
                需含「分标名称/标段」「分包名称/包号」「应答人」列，命中的应答人单元格在汇总表中自动标黄高亮。
              </p>
              <div className="flex items-center gap-2">
                <label className="flex-1 inline-flex cursor-pointer items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:border-[#3B5B7E] hover:text-[#3B5B7E] transition-colors min-w-0">
                  <Award className="h-3.5 w-3.5 shrink-0 text-[#9A6B26]" />
                  <span className="truncate">
                    {successListFile ? successListFile.name : '选择优秀名单 (.xlsx)'}
                  </span>
                  <input
                    type="file"
                    accept=".xlsx"
                    onChange={(e) => {
                      const f = e.target.files?.[0] || null;
                      e.target.value = '';
                      setSuccessListFile(f);
                      setResult(null);
                    }}
                    className="hidden"
                    disabled={loading}
                  />
                </label>
                {successListFile && (
                  <button
                    type="button"
                    onClick={() => {
                      setSuccessListFile(null);
                      setResult(null);
                    }}
                    disabled={loading}
                    className="text-xs text-slate-400 hover:text-[#A94442] shrink-0"
                  >
                    清除
                  </button>
                )}
              </div>
            </div>
          </div>

          <div className="border-t border-slate-100 pt-3 flex flex-wrap items-center gap-2">
            <button
              type="button"
              onClick={handleRun}
              disabled={loading || scoreFiles.length === 0}
              className="inline-flex items-center gap-1.5 rounded-lg bg-[#3B5B7E] px-4 py-2 text-xs font-medium text-white hover:bg-[#2E4866] disabled:opacity-45 transition-colors"
            >
              <Play className="h-3.5 w-3.5" />
              <span>{loading ? '正在合并中...' : '开始按模板合并'}</span>
            </button>
          </div>
        </div>
      </div>

      {/* Error */}
      {errorMsg && (
        <div className="flex items-start gap-2 rounded-xl border border-[#A94442]/30 bg-[#FEF2F2] px-4 py-3 text-xs text-[#A94442]">
          <AlertTriangle className="h-4 w-4 shrink-0 mt-0.5" />
          <span className="whitespace-pre-line">{errorMsg}</span>
        </div>
      )}

      {/* Results Summary & Log */}
      {result && (
        <div className="space-y-4">
          <div className="rounded-xl border border-[#2F6F5E]/30 bg-[#ECFDF5]/60 p-4 space-y-3">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <div className="flex items-center gap-2 text-xs font-semibold text-[#2F6F5E]">
                <CheckCircle2 className="h-4 w-4 shrink-0" />
                <span>
                  合并完成（{result.usedTemplate ? '原始模板基线模式' : '众数骨架旧版模式'}）：已合并{' '}
                  {result.validFileCount} / {result.totalInputFiles} 份打分表 → {result.outputFileName}
                </span>
              </div>
              <a
                href={`/api/toolbox/score-merge/${result.jobId}/download`}
                className="inline-flex items-center gap-1.5 rounded-lg bg-[#2F6F5E] px-3.5 py-1.5 text-xs font-medium text-white hover:bg-[#24594B] transition-colors"
              >
                <Download className="h-3.5 w-3.5" />
                <span>下载 {result.outputFileName}</span>
              </a>
            </div>

            <div className="flex flex-wrap gap-2 text-[11px] tabular-nums">
              {result.usedTemplate && (
                <>
                  <span className="rounded bg-white border border-slate-200 px-2.5 py-1 text-slate-700">
                    有效评分行：<strong>{result.activeRows}</strong>
                  </span>
                  <span className="rounded bg-white border border-slate-200 px-2.5 py-1 text-slate-500">
                    忽略未调整/空行：<strong>{result.ignoredRows}</strong>
                  </span>
                </>
              )}
              <span
                className={`rounded border px-2.5 py-1 ${
                  result.conflictCount > 0
                    ? 'bg-[#FEF9EC] border-[#9A6B26]/30 text-[#9A6B26] font-semibold'
                    : 'bg-white border-slate-200 text-[#2F6F5E]'
                }`}
              >
                分数冲突（已标黄）：<strong>{result.conflictCount}</strong> 处
              </span>
              <span className="rounded bg-white border border-slate-200 px-2.5 py-1 text-slate-700">
                优秀名单高亮：<strong>{result.highlightCount}</strong> 处
              </span>
              {result.skippedFiles.length > 0 && (
                <span className="rounded bg-[#FEF2F2] border border-[#A94442]/30 px-2.5 py-1 text-[#A94442] font-medium">
                  跳过不匹配文件：<strong>{result.skippedFiles.length}</strong> 份
                </span>
              )}
            </div>
          </div>

          {/* Preview Table */}
          {result.sheetPreviews && result.sheetPreviews.length > 0 && (() => {
            const curPreview =
              result.sheetPreviews[Math.min(activePreviewIdx, result.sheetPreviews.length - 1)];
            const colLabel = (idx: number) => {
              let n = idx + 1;
              let s = '';
              while (n > 0) {
                const rem = (n - 1) % 26;
                s = String.fromCharCode(65 + rem) + s;
                n = Math.floor((n - 1) / 26);
              }
              return s;
            };
            return (
              <div className="rounded-xl border border-slate-200 bg-white overflow-hidden">
                <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 bg-slate-50/70 px-4 py-2.5">
                  <div className="flex flex-wrap items-center gap-1.5">
                    <TableProperties className="h-4 w-4 text-[#3B5B7E]" />
                    <span className="text-xs font-semibold text-slate-700 mr-1">结果预览：</span>
                    {result.sheetPreviews.map((sp, idx) => (
                      <button
                        key={sp.name}
                        type="button"
                        onClick={() => setActivePreviewIdx(idx)}
                        className={`rounded-md px-2.5 py-1 text-[11px] font-medium tabular-nums transition-colors ${
                          activePreviewIdx === idx
                            ? 'bg-[#3B5B7E] text-white'
                            : 'bg-white border border-slate-200 text-slate-600 hover:bg-slate-100'
                        }`}
                      >
                        {sp.name} ({sp.totalRows}行)
                      </button>
                    ))}
                  </div>
                  <span className="text-[11px] text-slate-400">
                    仅预览前 {curPreview.rows.length} 行（黄底为冲突/高亮单元格），下载包含全部 {curPreview.totalRows} 行
                  </span>
                </div>

                <div className="max-h-96 overflow-auto">
                  <table className="w-full border-collapse text-left text-xs">
                    <thead className="sticky top-0 z-10 bg-slate-100 text-slate-500 border-b border-slate-200">
                      <tr>
                        <th className="px-2 py-1.5 text-center font-mono text-[10px] border-r border-slate-200 w-10">
                          #
                        </th>
                        {Array.from({ length: curPreview.colCount }, (_, cIdx) => (
                          <th
                            key={cIdx}
                            className="px-2.5 py-1.5 font-mono text-[10px] font-medium border-r border-slate-200/70 whitespace-nowrap"
                          >
                            {colLabel(cIdx)}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {curPreview.rows.length === 0 ? (
                        <tr>
                          <td
                            colSpan={curPreview.colCount + 1}
                            className="py-8 text-center text-xs text-slate-400"
                          >
                            该工作表暂无数据
                          </td>
                        </tr>
                      ) : (
                        curPreview.rows.map((row) => {
                          const isHeader = row.rowNum <= curPreview.headerRows;
                          return (
                            <tr
                              key={row.rowNum}
                              className={isHeader ? 'bg-slate-50/90 font-medium' : 'hover:bg-slate-50/70'}
                            >
                              <td className="px-2 py-1.5 text-center font-mono text-[10px] text-slate-400 border-r border-slate-200 bg-slate-50/70">
                                {row.rowNum}
                              </td>
                              {row.cells.map((cell, cIdx) => (
                                <td
                                  key={cIdx}
                                  title={cell.text}
                                  className={`px-2.5 py-1.5 max-w-[220px] truncate border-r border-slate-100 ${
                                    cell.highlighted
                                      ? 'bg-[#FEF3C7] text-[#92400E] font-semibold'
                                      : isHeader
                                        ? 'text-slate-700'
                                        : 'text-slate-600'
                                  }`}
                                >
                                  {cell.text}
                                </td>
                              ))}
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            );
          })()}

          {/* Run Logs */}
          <div className="rounded-xl border border-slate-200 bg-white overflow-hidden">
            <div className="border-b border-slate-200 bg-slate-50/70 px-4 py-2 text-xs font-semibold text-slate-600">
              运行日志
            </div>
            <pre className="max-h-64 overflow-auto p-3.5 text-[11px] font-mono leading-relaxed text-slate-700 bg-slate-50/40 whitespace-pre-wrap">
              {result.logs.join('\n')}
            </pre>
          </div>
        </div>
      )}
    </div>
  );
}
