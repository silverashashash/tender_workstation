import { useState } from 'react';
import {
  AlertTriangle,
  Archive,
  CheckCircle2,
  Download,
  FileSpreadsheet,
  Info,
  Play,
  RotateCcw,
  Shuffle,
  Upload,
} from 'lucide-react';
import { api } from '../api/client';

type RandomScoreResponse = Awaited<ReturnType<typeof api.runRandomScoreGenerator>>;
type NameMode = 'manual' | 'excel' | 'auto';

export function RandomScorePanel() {
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [nameMode, setNameMode] = useState<NameMode>('auto');
  const [namesFile, setNamesFile] = useState<File | null>(null);
  const [manualNames, setManualNames] = useState('');
  const [nExpertsInput, setNExpertsInput] = useState('5');
  const [maxDiffInput, setMaxDiffInput] = useState('7');

  const [loading, setLoading] = useState(false);
  const [errorMsg, setErrorMsg] = useState('');
  const [result, setResult] = useState<RandomScoreResponse | null>(null);

  function handleReset() {
    setTemplateFile(null);
    setNamesFile(null);
    setResult(null);
    setErrorMsg('');
  }

  async function handleRun() {
    if (!templateFile) {
      setErrorMsg('请先选择原始专家打分表（.xlsx）');
      return;
    }
    if (nameMode === 'excel' && !namesFile) {
      setErrorMsg('请选择专家姓名表（.xlsx，第一列为专家姓名）');
      return;
    }
    if (nameMode === 'manual' && !manualNames.trim()) {
      setErrorMsg('请填写要生成的专家姓名（逗号或换行分隔）');
      return;
    }
    const maxTotalDiff = maxDiffInput.trim() === '' ? 7 : Number(maxDiffInput.trim());
    if (!Number.isFinite(maxTotalDiff) || maxTotalDiff < 0) {
      setErrorMsg('总分最大差值必须为大于或等于 0 的数字');
      return;
    }

    setLoading(true);
    setErrorMsg('');
    setResult(null);
    try {
      const res = await api.runRandomScoreGenerator({
        template: templateFile,
        nameMode,
        namesFile: nameMode === 'excel' ? namesFile : null,
        manualNames: nameMode === 'manual' ? manualNames : undefined,
        nExperts: nameMode === 'auto' ? Number(nExpertsInput) : undefined,
        maxTotalDiff,
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
            根据一份已打分的专家打分表，自动识别表头结构与<strong>黄色高亮客观分列</strong>，为其余 <code className="rounded bg-slate-100 px-1 text-[11px]">n-1</code> 位专家随机生成主观分并保持客观分一致，严格校验<strong>总分差 ≤ 阈值（默认 7 分）</strong>与<strong>同分人数 ≤ ⌊n/2⌋</strong>，自动跳过未通过符合性评审或异常分数的行。
          </span>
        </div>
        {templateFile && (
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

      {/* Main Config Card */}
      <div className="grid grid-cols-1 gap-4 lg:grid-cols-12">
        {/* Left: Template + Threshold */}
        <div className="lg:col-span-5 rounded-xl border border-slate-200 bg-white p-4 space-y-4">
          <div className="space-y-2">
            <div className="flex items-center gap-2">
              <Shuffle className="h-4 w-4 text-[#3B5B7E]" />
              <h2 className="text-sm font-semibold text-slate-800">1. 选择已打分的源专家表 *</h2>
            </div>
            <p className="text-[11px] text-slate-500 leading-relaxed">
              需包含「合计」列及上方的最高分/最低分/级差行；表头标黄列自动锁定为客观分，不作随机改动。
            </p>
            <div className="flex items-center gap-2 pt-1">
              <label className="flex-1 inline-flex cursor-pointer items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-3 py-2 text-xs font-medium text-slate-700 hover:border-[#3B5B7E] hover:text-[#3B5B7E] transition-colors min-w-0">
                <Upload className="h-3.5 w-3.5 shrink-0" />
                <span className="truncate">
                  {templateFile ? templateFile.name : '选择源专家打分表 (.xlsx)'}
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

          <div className="border-t border-slate-100 pt-3">
            <label className="block text-xs font-medium text-slate-700 mb-1">
              总分最大差值上限（默认 7 分）
            </label>
            <div className="flex items-center gap-2">
              <input
                type="number"
                min={0}
                step="0.5"
                value={maxDiffInput}
                onChange={(e) => setMaxDiffInput(e.target.value)}
                disabled={loading}
                className="w-24 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs tabular-nums focus:border-[#3B5B7E] focus:outline-none"
              />
              <span className="text-[11px] text-slate-400">分（同分专家数自动限制 ≤ ⌊n/2⌋）</span>
            </div>
          </div>
        </div>

        {/* Right: Expert Names Setup */}
        <div className="lg:col-span-7 rounded-xl border border-slate-200 bg-white p-4 flex flex-col justify-between gap-4">
          <div className="space-y-3">
            <h2 className="text-sm font-semibold text-slate-800">2. 设置待生成的其余专家名单</h2>

            <div className="flex flex-wrap gap-1.5">
              {([
                { key: 'auto', label: '自动命名 (专家2~n)' },
                { key: 'manual', label: '手动输入姓名' },
                { key: 'excel', label: '从 Excel 导入姓名表' },
              ] as const).map((m) => (
                <button
                  key={m.key}
                  type="button"
                  onClick={() => {
                    setNameMode(m.key);
                    setErrorMsg('');
                  }}
                  className={`rounded-lg px-3 py-1.5 text-xs font-medium transition-colors ${
                    nameMode === m.key
                      ? 'bg-[#3B5B7E] text-white'
                      : 'bg-slate-100 text-slate-600 hover:bg-slate-200/70'
                  }`}
                >
                  {m.label}
                </button>
              ))}
            </div>

            {nameMode === 'auto' && (
              <div className="flex items-center gap-2 pt-1">
                <span className="text-xs text-slate-600">专家总数 n =</span>
                <input
                  type="number"
                  min={2}
                  max={30}
                  value={nExpertsInput}
                  onChange={(e) => setNExpertsInput(e.target.value)}
                  disabled={loading}
                  className="w-20 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs tabular-nums focus:border-[#3B5B7E] focus:outline-none"
                />
                <span className="text-[11px] text-slate-400">
                  （含已有的 1 位专家，将自动生成专家2 ~ 专家{Math.max(2, Number(nExpertsInput) || 2)}）
                </span>
              </div>
            )}

            {nameMode === 'manual' && (
              <div className="space-y-1">
                <label className="block text-[11px] text-slate-500">
                  每行一个或用逗号/顿号分隔（仅填写需要新生成的专家姓名，不含源表已有专家）
                </label>
                <textarea
                  rows={3}
                  value={manualNames}
                  onChange={(e) => setManualNames(e.target.value)}
                  disabled={loading}
                  placeholder="张三、李四、王五、赵六"
                  className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>
            )}

            {nameMode === 'excel' && (
              <div className="space-y-1.5">
                <div className="text-[11px] text-slate-500">
                  读取第一列姓名（自动跳过「姓名 / 专家姓名 / 专家名称」表头，不含源表已有专家）
                </div>
                <div className="flex items-center gap-2">
                  <label className="inline-flex cursor-pointer items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:border-[#3B5B7E] hover:text-[#3B5B7E]">
                    <FileSpreadsheet className="h-3.5 w-3.5" />
                    <span>{namesFile ? namesFile.name : '选择专家姓名表 (.xlsx)'}</span>
                    <input
                      type="file"
                      accept=".xlsx"
                      onChange={(e) => {
                        const f = e.target.files?.[0] || null;
                        e.target.value = '';
                        setNamesFile(f);
                      }}
                      className="hidden"
                      disabled={loading}
                    />
                  </label>
                  {namesFile && (
                    <button
                      type="button"
                      onClick={() => setNamesFile(null)}
                      className="text-xs text-slate-400 hover:text-[#A94442]"
                    >
                      清除
                    </button>
                  )}
                </div>
              </div>
            )}
          </div>

          <div className="border-t border-slate-100 pt-3 flex flex-wrap items-center gap-2">
            <button
              type="button"
              onClick={handleRun}
              disabled={loading || !templateFile}
              className="inline-flex items-center gap-1.5 rounded-lg bg-[#3B5B7E] px-4 py-2 text-xs font-medium text-white hover:bg-[#2E4866] disabled:opacity-45 transition-colors"
            >
              <Play className="h-3.5 w-3.5" />
              <span>{loading ? '正在随机生成并校验...' : '开始生成专家打分表'}</span>
            </button>

            {result && (
              <a
                href={`/api/toolbox/random-score/${result.jobId}/download-all`}
                className="inline-flex items-center gap-1.5 rounded-lg bg-[#2F6F5E] px-4 py-2 text-xs font-medium text-white hover:bg-[#24594B] transition-colors"
              >
                <Archive className="h-3.5 w-3.5" />
                <span>全部打包下载 (.zip)</span>
              </a>
            )}
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

      {/* Results */}
      {result && (
        <div className="space-y-4">
          <div className="rounded-xl border border-[#2F6F5E]/30 bg-white overflow-hidden">
            <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 bg-[#ECFDF5]/60 px-4 py-3">
              <div className="flex items-center gap-2 text-xs font-semibold text-[#2F6F5E]">
                <CheckCircle2 className="h-4 w-4 shrink-0" />
                <span>
                  已生成 {result.outputs.length} 份打分表（共 {result.nExperts} 位专家 · 核验{' '}
                  {result.totalCheckedGroups} 组 ·{' '}
                  {result.violations === 0 ? '规则全部通过 ✓' : `${result.violations} 处违规`}）
                </span>
              </div>
              <a
                href={`/api/toolbox/random-score/${result.jobId}/download-all`}
                className="inline-flex items-center gap-1.5 rounded-lg bg-[#2F6F5E] px-3.5 py-1.5 text-xs font-medium text-white hover:bg-[#24594B] transition-colors"
              >
                <Archive className="h-3.5 w-3.5" />
                <span>全部打包下载 (.zip)</span>
              </a>
            </div>

            <div className="p-4 flex flex-wrap gap-2">
              {result.outputs.map((out) => (
                <a
                  key={out.key}
                  href={`/api/toolbox/random-score/${result.jobId}/download/${out.key}`}
                  className="inline-flex items-center gap-1.5 rounded-lg border border-[#3B5B7E]/30 bg-[#EBF2F7]/60 px-3 py-1.5 text-xs font-medium text-[#3B5B7E] hover:bg-[#DCE8F2]"
                >
                  <Download className="h-3.5 w-3.5" />
                  <span>{out.filename}</span>
                </a>
              ))}
            </div>
          </div>

          {result.skippedRows.length > 0 && (
            <div className="rounded-xl border border-[#9A6B26]/30 bg-white overflow-hidden">
              <div className="border-b border-[#9A6B26]/20 bg-[#FEF9EC]/70 px-4 py-2 text-xs font-semibold text-[#9A6B26]">
                以下 {result.skippedRows.length} 行已自动跳过（保持原始分数不变）
              </div>
              <div className="max-h-56 overflow-auto">
                <table className="w-full border-collapse text-left text-xs">
                  <thead className="sticky top-0 bg-slate-50 text-slate-500 border-b border-slate-200">
                    <tr>
                      <th className="px-3 py-1.5 font-medium">工作表</th>
                      <th className="px-3 py-1.5 font-medium">行号</th>
                      <th className="px-3 py-1.5 font-medium">标段 / 包号 / 应答人</th>
                      <th className="px-3 py-1.5 font-medium">跳过原因</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {result.skippedRows.map((s, idx) => (
                      <tr key={idx}>
                        <td className="px-3 py-1.5 text-slate-600">{s.sheetName}</td>
                        <td className="px-3 py-1.5 font-mono text-slate-500">{s.row}</td>
                        <td className="px-3 py-1.5 text-slate-800">
                          {[s.section, s.packageNo, s.supplier].filter(Boolean).join(' · ')}
                        </td>
                        <td className="px-3 py-1.5 text-[#9A6B26]">{s.reason}</td>
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
