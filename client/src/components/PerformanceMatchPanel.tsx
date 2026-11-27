import { useState } from 'react';
import {
  AlertTriangle,
  Archive,
  CheckCircle2,
  Download,
  FileCheck2,
  FileSpreadsheet,
  Info,
  Play,
  RotateCcw,
  Upload,
  X,
} from 'lucide-react';
import { api } from '../api/client';

type PerfMatchResponse = Awaited<ReturnType<typeof api.runPerformanceMatch>>;

export function PerformanceMatchPanel() {
  const [lyfFile, setLyfFile] = useState<File | null>(null);
  const [templateFiles, setTemplateFiles] = useState<File[]>([]);
  const [defaultScoreInput, setDefaultScoreInput] = useState('4');

  const [loading, setLoading] = useState(false);
  const [errorMsg, setErrorMsg] = useState('');
  const [result, setResult] = useState<PerfMatchResponse | null>(null);

  function handleAddTemplates(e: React.ChangeEvent<HTMLInputElement>) {
    const picked = Array.from(e.target.files || []).filter(
      (f) => /\.(xlsx|zip)$/i.test(f.name) && !f.name.startsWith('~$')
    );
    e.target.value = '';
    if (picked.length === 0) return;
    setErrorMsg('');
    setResult(null);
    setTemplateFiles((prev) => {
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

  function removeTemplate(idx: number) {
    setTemplateFiles((prev) => prev.filter((_, i) => i !== idx));
    setResult(null);
  }

  function handleReset() {
    setLyfFile(null);
    setTemplateFiles([]);
    setErrorMsg('');
    setResult(null);
  }

  async function handleRun() {
    if (!lyfFile) {
      setErrorMsg('请先选择《履约分》Excel 文件（.xlsx）');
      return;
    }
    if (templateFiles.length === 0) {
      setErrorMsg('请至少选择一份详评打分模板（支持多选 .xlsx 或 .zip）');
      return;
    }
    const defaultScore = defaultScoreInput.trim() === '' ? 4 : Number(defaultScoreInput.trim());
    if (!Number.isFinite(defaultScore) || defaultScore < 0) {
      setErrorMsg('未匹配默认分必须为大于或等于 0 的数字');
      return;
    }

    setLoading(true);
    setErrorMsg('');
    setResult(null);
    try {
      const res = await api.runPerformanceMatch({
        lyf: lyfFile,
        templates: templateFiles,
        defaultScore,
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
            按<strong>分标名称 + 包号 + 投标人</strong>三元组将《履约分》自动匹配并填入一份或多份《详评打分模板》的<strong>「绩效评价」</strong>列，完整保留原模板样式、合并单元格与合计公式；未匹配行自动填默认分（默认 4 分）并生成未匹配明细报告。
          </span>
        </div>
        {(lyfFile || templateFiles.length > 0) && (
          <button
            type="button"
            onClick={handleReset}
            disabled={loading}
            className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-slate-50 px-2.5 py-1 text-xs font-medium text-slate-600 hover:border-[#A94442]/40 hover:text-[#A94442]"
          >
            <RotateCcw className="h-3.5 w-3.5" />
            <span>清空重选</span>
          </button>
        )}
      </div>

      {/* Main 2-column setup */}
      <div className="grid grid-cols-1 gap-4 lg:grid-cols-12">
        {/* Left: 1. 履约分表 + 默认分 */}
        <div className="lg:col-span-5 rounded-xl border border-slate-200 bg-white p-4 flex flex-col justify-between gap-4">
          <div className="space-y-3.5">
            <div className="space-y-1.5">
              <div className="flex items-center gap-2">
                <FileCheck2 className="h-4 w-4 text-[#3B5B7E]" />
                <h2 className="text-sm font-semibold text-slate-800">1. 选择履约分表 *</h2>
              </div>
              <p className="text-[11px] text-slate-500 leading-relaxed">
                自动识别「分标名称 / 分包名称 / 投标人 / 履约分」列（兼容标准 B~E 列布局与本工作台「履约分生成」导出表），批量处理多个模板时仅读取一次。
              </p>
              <div className="flex items-center gap-2 pt-1">
                <label className="flex-1 inline-flex cursor-pointer items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-3 py-2 text-xs font-medium text-slate-700 hover:border-[#3B5B7E] hover:text-[#3B5B7E] transition-colors min-w-0">
                  <Upload className="h-3.5 w-3.5 shrink-0" />
                  <span className="truncate">
                    {lyfFile ? lyfFile.name : '选择履约分表 (.xlsx)'}
                  </span>
                  <input
                    type="file"
                    accept=".xlsx"
                    onChange={(e) => {
                      const f = e.target.files?.[0] || null;
                      e.target.value = '';
                      setLyfFile(f);
                      setResult(null);
                    }}
                    className="hidden"
                    disabled={loading}
                  />
                </label>
                {lyfFile && (
                  <button
                    type="button"
                    onClick={() => {
                      setLyfFile(null);
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

            <div className="border-t border-slate-100 pt-3">
              <label className="block text-xs font-medium text-slate-700 mb-1">
                未匹配时默认填入分值（未参加绩效评价）
              </label>
              <div className="flex items-center gap-2">
                <input
                  type="number"
                  min={0}
                  step="0.5"
                  value={defaultScoreInput}
                  onChange={(e) => setDefaultScoreInput(e.target.value)}
                  disabled={loading}
                  className="w-24 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs tabular-nums focus:border-[#3B5B7E] focus:outline-none"
                />
                <span className="text-[11px] text-slate-400">分（默认 4 分）</span>
              </div>
            </div>
          </div>

          <div className="border-t border-slate-100 pt-3">
            <button
              type="button"
              onClick={handleRun}
              disabled={loading || !lyfFile || templateFiles.length === 0}
              className="w-full inline-flex items-center justify-center gap-1.5 rounded-lg bg-[#3B5B7E] px-4 py-2 text-xs font-medium text-white hover:bg-[#2E4866] disabled:opacity-45 transition-colors"
            >
              <Play className="h-3.5 w-3.5" />
              <span>{loading ? '正在匹配填入...' : '开始匹配并填入履约分'}</span>
            </button>
          </div>
        </div>

        {/* Right: 2. 详评打分模板列表（支持批量多选） */}
        <div className="lg:col-span-7 rounded-xl border border-slate-200 bg-white p-4 space-y-3">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="flex items-center gap-2">
              <FileSpreadsheet className="h-4 w-4 text-[#3B5B7E]" />
              <h2 className="text-sm font-semibold text-slate-800">2. 详评打分模板 *（支持批量多选或 ZIP）</h2>
              {templateFiles.length > 0 && (
                <span className="rounded-full bg-[#EBF2F7] px-2 py-0.5 text-[11px] font-medium text-[#3B5B7E] tabular-nums">
                  已选 {templateFiles.length} 份
                </span>
              )}
            </div>
            <label className="inline-flex cursor-pointer items-center gap-1.5 rounded-lg border border-[#3B5B7E]/35 bg-[#EBF2F7]/60 px-3 py-1.5 text-xs font-medium text-[#3B5B7E] hover:bg-[#EBF2F7] transition-colors">
              <Upload className="h-3.5 w-3.5" />
              <span>{templateFiles.length > 0 ? '继续添加模板' : '选择模板 / ZIP'}</span>
              <input
                type="file"
                multiple
                accept=".xlsx,.zip"
                onChange={handleAddTemplates}
                className="hidden"
                disabled={loading}
              />
            </label>
          </div>

          {templateFiles.length === 0 ? (
            <div className="rounded-lg border border-dashed border-slate-300 bg-slate-50/60 py-10 text-center text-xs text-slate-400">
              尚未选择详评打分模板（支持一次添加多个组的模板，自动按 Sheet 探测「标段/包号/应答人」与「绩效评价」列）
            </div>
          ) : (
            <div className="max-h-56 overflow-y-auto rounded-lg border border-slate-200 divide-y divide-slate-100">
              {templateFiles.map((f, idx) => (
                <div key={`${f.name}-${idx}`} className="flex items-center justify-between gap-2 px-3 py-1.5 text-xs hover:bg-slate-50">
                  <div className="flex items-center gap-2 min-w-0">
                    <FileSpreadsheet className="h-3.5 w-3.5 text-[#3B5B7E] shrink-0" />
                    <span className="truncate text-slate-700" title={f.name}>
                      {f.name}
                    </span>
                    <span className="text-[10px] text-slate-400 tabular-nums shrink-0">
                      {(f.size / 1024).toFixed(0)} KB
                    </span>
                  </div>
                  <button
                    type="button"
                    onClick={() => removeTemplate(idx)}
                    disabled={loading}
                    className="rounded p-0.5 text-slate-400 hover:text-[#A94442]"
                    title="移除"
                  >
                    <X className="h-3.5 w-3.5" />
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      {/* Error */}
      {errorMsg && (
        <div className="flex items-start gap-2 rounded-xl border border-[#A94442]/30 bg-[#FEF2F2] px-4 py-3 text-xs text-[#A94442]">
          <AlertTriangle className="h-4 w-4 shrink-0 mt-0.5" />
          <span className="whitespace-pre-line">{errorMsg}</span>
        </div>
      )}

      {/* Results */}
      {result && (
        <div className="space-y-4">
          <div className="rounded-xl border border-[#2F6F5E]/30 bg-white overflow-hidden">
            <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 bg-[#ECFDF5]/60 px-4 py-3">
              <div className="flex items-center gap-2 text-xs font-semibold text-[#2F6F5E]">
                <CheckCircle2 className="h-4 w-4 shrink-0" />
                <span>
                  匹配完成：共处理 {result.templates.length} 个模板 · 载入履约分 {result.lyfRecordCount} 条 · 匹配{' '}
                  {result.totalMatched}/{result.totalRows} 行
                  {result.totalDefaulted > 0 ? ` · 默认填 ${result.defaultScore} 分 ${result.totalDefaulted} 条` : ''}
                </span>
              </div>
              <a
                href={`/api/toolbox/perf-match/${result.jobId}/download-all`}
                className="inline-flex items-center gap-1.5 rounded-lg bg-[#2F6F5E] px-3.5 py-1.5 text-xs font-medium text-white hover:bg-[#24594B] transition-colors"
              >
                <Archive className="h-3.5 w-3.5" />
                <span>全部打包下载 (.zip 含匹配报告)</span>
              </a>
            </div>

            <div className="divide-y divide-slate-100">
              {result.templates.map((t) => (
                <div key={t.key} className="px-4 py-3 flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
                  <div className="min-w-0 space-y-0.5">
                    <div className="text-xs font-medium text-slate-800 truncate">{t.originalName}</div>
                    <div className="text-[11px] text-slate-500 tabular-nums">
                      处理 {t.sheetsProcessed} 个 Sheet · 匹配 <strong className="text-[#2F6F5E]">{t.matched}</strong> /{' '}
                      {t.totalRows} 行
                      {t.defaulted > 0 && (
                        <>
                          {' · '}未匹配默认填 {result.defaultScore} 分：
                          <strong className="text-[#9A6B26]">{t.defaulted}</strong> 条
                        </>
                      )}
                      {t.sheetsSkipped.length > 0 && ` · 跳过 ${t.sheetsSkipped.length} 个 Sheet`}
                    </div>
                  </div>
                  <a
                    href={`/api/toolbox/perf-match/${result.jobId}/download/${t.key}`}
                    className="inline-flex shrink-0 items-center gap-1 rounded-md border border-[#3B5B7E]/30 bg-[#EBF2F7]/60 px-2.5 py-1.5 text-[11px] font-medium text-[#3B5B7E] hover:bg-[#DCE8F2]"
                  >
                    <Download className="h-3.5 w-3.5" />
                    <span className="max-w-[260px] truncate">{t.outputFileName}</span>
                  </a>
                </div>
              ))}
            </div>
          </div>

          {/* Unmatched Table if any */}
          {result.unmatchedItems.length > 0 && (
            <div className="rounded-xl border border-[#9A6B26]/30 bg-white overflow-hidden">
              <div className="border-b border-[#9A6B26]/20 bg-[#FEF9EC]/70 px-4 py-2 text-xs font-semibold text-[#9A6B26]">
                未匹配明细（共 {result.unmatchedItems.length} 条，已自动按默认 {result.defaultScore} 分填入）
              </div>
              <div className="max-h-60 overflow-auto">
                <table className="w-full border-collapse text-left text-xs">
                  <thead className="sticky top-0 bg-slate-50 text-slate-500 border-b border-slate-200">
                    <tr>
                      <th className="px-3 py-1.5 font-medium">模板</th>
                      <th className="px-3 py-1.5 font-medium">工作表</th>
                      <th className="px-3 py-1.5 font-medium">标段</th>
                      <th className="px-3 py-1.5 font-medium">包号</th>
                      <th className="px-3 py-1.5 font-medium">应答人</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {result.unmatchedItems.slice(0, 100).map((u, idx) => (
                      <tr key={idx} className="hover:bg-slate-50/70">
                        <td className="px-3 py-1.5 text-slate-600">{u.templateName}</td>
                        <td className="px-3 py-1.5 text-slate-600">{u.sheetName}</td>
                        <td className="px-3 py-1.5 text-slate-700">{u.section}</td>
                        <td className="px-3 py-1.5 text-slate-600">{u.packageNo}</td>
                        <td className="px-3 py-1.5 font-medium text-slate-800">{u.supplier}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* Run Logs */}
          <div className="rounded-xl border border-slate-200 bg-white overflow-hidden">
            <div className="border-b border-slate-200 bg-slate-50/70 px-4 py-2 text-xs font-semibold text-slate-600">
              运行日志
            </div>
            <pre className="max-h-56 overflow-auto p-3.5 text-[11px] font-mono leading-relaxed text-slate-700 bg-slate-50/40 whitespace-pre-wrap">
              {result.logs.join('\n')}
            </pre>
          </div>
        </div>
      )}
    </div>
  );
}
