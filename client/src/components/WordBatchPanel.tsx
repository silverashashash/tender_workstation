import { useEffect, useState } from 'react';
import {
  AlertTriangle,
  Archive,
  CheckCircle2,
  Download,
  Edit3,
  FileSpreadsheet,
  FileText,
  Info,
  Play,
  Plus,
  RotateCcw,
  Trash2,
  Upload,
  X,
} from 'lucide-react';
import { api } from '../api/client';

interface ReplacePair {
  find: string;
  replace: string;
}

interface StoredConfig {
  modifyEnabled: boolean;
  modifyTemplate: string;
  multiplier: string;
  deleteEnabled: boolean;
  deleteTemplate: string;
  replaceEnabled: boolean;
  replaceUseRegex: boolean;
  replacePairs: ReplacePair[];
  extractEnabled: boolean;
}

const STORAGE_KEY = 'tender_word_batch_cfg_v1';

const DEFAULT_CFG: StoredConfig = {
  modifyEnabled: true,
  modifyTemplate: '最高应答限价为*万元',
  multiplier: '10000',
  deleteEnabled: false,
  deleteTemplate: '有效报价平均值为*万元',
  replaceEnabled: false,
  replaceUseRegex: false,
  replacePairs: [],
  extractEnabled: false,
};

function loadStoredConfig(): StoredConfig {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return DEFAULT_CFG;
    const parsed = JSON.parse(raw) as Partial<StoredConfig>;
    return {
      ...DEFAULT_CFG,
      ...parsed,
      replacePairs: Array.isArray(parsed.replacePairs) ? parsed.replacePairs : [],
    };
  } catch {
    return DEFAULT_CFG;
  }
}

type ProcessResponse = Awaited<ReturnType<typeof api.processWordBatch>>;

export function WordBatchPanel() {
  const [files, setFiles] = useState<File[]>([]);

  const initCfg = loadStoredConfig();
  const [modifyEnabled, setModifyEnabled] = useState(initCfg.modifyEnabled);
  const [modifyTemplate, setModifyTemplate] = useState(initCfg.modifyTemplate);
  const [multiplier, setMultiplier] = useState(initCfg.multiplier);

  const [deleteEnabled, setDeleteEnabled] = useState(initCfg.deleteEnabled);
  const [deleteTemplate, setDeleteTemplate] = useState(initCfg.deleteTemplate);

  const [replaceEnabled, setReplaceEnabled] = useState(initCfg.replaceEnabled);
  const [replaceUseRegex, setReplaceUseRegex] = useState(initCfg.replaceUseRegex);
  const [replacePairs, setReplacePairs] = useState<ReplacePair[]>(initCfg.replacePairs);
  const [findInput, setFindInput] = useState('');
  const [replaceInput, setReplaceInput] = useState('');

  const [extractEnabled, setExtractEnabled] = useState(initCfg.extractEnabled);

  const [loading, setLoading] = useState(false);
  const [errorMsg, setErrorMsg] = useState('');
  const [result, setResult] = useState<ProcessResponse | null>(null);

  // Persist settings to localStorage whenever they change
  useEffect(() => {
    try {
      const cfg: StoredConfig = {
        modifyEnabled,
        modifyTemplate,
        multiplier,
        deleteEnabled,
        deleteTemplate,
        replaceEnabled,
        replaceUseRegex,
        replacePairs,
        extractEnabled,
      };
      localStorage.setItem(STORAGE_KEY, JSON.stringify(cfg));
    } catch {
      // ignore storage quota errors
    }
  }, [
    modifyEnabled,
    modifyTemplate,
    multiplier,
    deleteEnabled,
    deleteTemplate,
    replaceEnabled,
    replaceUseRegex,
    replacePairs,
    extractEnabled,
  ]);

  function handleSelectFiles(e: React.ChangeEvent<HTMLInputElement>) {
    const picked = Array.from(e.target.files || []).filter((f) => {
      const name = f.name;
      return /\.(docx?)$/i.test(name) && !name.startsWith('~$');
    });
    e.target.value = '';
    if (picked.length === 0) return;
    setErrorMsg('');
    setFiles((prev) => {
      const existingNames = new Set(prev.map((f) => `${f.name}:${f.size}`));
      const merged = [...prev];
      for (const f of picked) {
        const key = `${f.name}:${f.size}`;
        if (!existingNames.has(key)) {
          merged.push(f);
          existingNames.add(key);
        }
      }
      return merged;
    });
  }

  function removeFile(idx: number) {
    setFiles((prev) => prev.filter((_, i) => i !== idx));
  }

  function handleAddPair() {
    if (!findInput) {
      setErrorMsg('「查找」内容不能为空');
      return;
    }
    setErrorMsg('');
    setReplacePairs((prev) => [...prev, { find: findInput, replace: replaceInput }]);
    setFindInput('');
    setReplaceInput('');
  }

  function handleEditPair(idx: number) {
    const target = replacePairs[idx];
    if (!target) return;
    setFindInput(target.find);
    setReplaceInput(target.replace);
    setReplacePairs((prev) => prev.filter((_, i) => i !== idx));
  }

  function handleRemovePair(idx: number) {
    setReplacePairs((prev) => prev.filter((_, i) => i !== idx));
  }

  async function handleRun() {
    if (files.length === 0) {
      setErrorMsg('请先选择至少一个 Word 文件（.docx / .doc）');
      return;
    }
    if (!modifyEnabled && !deleteEnabled && !replaceEnabled && !extractEnabled) {
      setErrorMsg('请至少勾选启用「① 批量修改」「② 批量删除」「③ 查找替换」「④ 提取评标表格」中的一项');
      return;
    }
    if (replaceEnabled && replacePairs.length === 0) {
      setErrorMsg('已启用「查找替换」，请先在下方添加至少一组替换规则');
      return;
    }

    setLoading(true);
    setErrorMsg('');
    setResult(null);
    try {
      const res = await api.processWordBatch(files, {
        modifyEnabled,
        modifyTemplate,
        multiplier,
        deleteEnabled,
        deleteTemplate,
        replaceEnabled,
        replaceUseRegex,
        replacePairs,
        extractEnabled,
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
      {/* Intro Banner */}
      <div className="rounded-xl border border-slate-200 bg-white px-4 py-3 flex flex-wrap items-center justify-between gap-2">
        <div className="flex items-start gap-2 text-xs text-slate-600">
          <Info className="h-4 w-4 shrink-0 text-[#3B5B7E] mt-0.5" />
          <span>
            批量处理 Word 评标/评审报告中的数字与文字：支持<strong>模板数字按倍数修改</strong>、<strong>整段删除</strong>、<strong>多组查找替换（支持正则）</strong>及<strong>评标评分表一键提取为 Excel</strong>。原文件不会被修改，均生成新的 <code className="rounded bg-slate-100 px-1 text-[11px]">*_modified.docx</code> 或 <code className="rounded bg-slate-100 px-1 text-[11px]">*_表格.xlsx</code>。
          </span>
        </div>
        {files.length > 0 && (
          <button
            type="button"
            onClick={() => {
              setFiles([]);
              setResult(null);
              setErrorMsg('');
            }}
            className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-slate-50 px-2.5 py-1 text-xs font-medium text-slate-600 hover:border-[#A94442]/40 hover:text-[#A94442]"
          >
            <RotateCcw className="h-3.5 w-3.5" />
            <span>清空文件</span>
          </button>
        )}
      </div>

      {/* Step 0: File Picker */}
      <div className="rounded-xl border border-slate-200 bg-white p-4 space-y-3">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div className="flex items-center gap-2">
            <FileText className="h-4 w-4 text-[#3B5B7E]" />
            <h2 className="text-sm font-semibold text-slate-800">选择待处理的 Word 文档（支持多选批量处理）</h2>
            {files.length > 0 && (
              <span className="rounded-full bg-[#EBF2F7] px-2 py-0.5 text-[11px] font-medium text-[#3B5B7E] tabular-nums">
                已选 {files.length} 个文件
              </span>
            )}
          </div>
          <label className="inline-flex cursor-pointer items-center gap-1.5 rounded-lg border border-[#3B5B7E]/35 bg-[#EBF2F7]/50 px-3 py-1.5 text-xs font-medium text-[#3B5B7E] hover:bg-[#EBF2F7] transition-colors">
            <Upload className="h-3.5 w-3.5" />
            <span>{files.length > 0 ? '继续添加 Word 文件' : '选择 .docx / .doc 文件'}</span>
            <input
              type="file"
              multiple
              accept=".docx,.doc"
              onChange={handleSelectFiles}
              className="hidden"
              disabled={loading}
            />
          </label>
        </div>

        {files.length > 0 ? (
          <div className="flex flex-wrap gap-1.5">
            {files.map((f, idx) => (
              <span
                key={`${f.name}-${idx}`}
                className="inline-flex items-center gap-1.5 rounded-md border border-slate-200 bg-slate-50 px-2.5 py-1 text-xs text-slate-700"
              >
                <FileText className="h-3.5 w-3.5 text-[#3B5B7E] shrink-0" />
                <span className="max-w-[280px] truncate" title={f.name}>
                  {f.name}
                </span>
                <span className="text-[10px] text-slate-400 tabular-nums">
                  {(f.size / 1024).toFixed(0)} KB
                </span>
                <button
                  type="button"
                  onClick={() => removeFile(idx)}
                  disabled={loading}
                  className="text-slate-400 hover:text-[#A94442]"
                  title="移除"
                >
                  <X className="h-3 w-3" />
                </button>
              </span>
            ))}
          </div>
        ) : (
          <div className="text-[11px] text-slate-400">
            支持新版 .docx（纯内置解析，无需安装 Office）与旧版 .doc（按文件头自动识别，本机装有 LibreOffice 或 Windows Word/WPS 时自动转换）。
          </div>
        )}
      </div>

      {/* Rules Grid: ① Modify & ② Delete side by side */}
      <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
        {/* ① 批量修改：数字 ×倍数 */}
        <div
          className={`rounded-xl border p-4 space-y-3 transition-colors ${
            modifyEnabled ? 'border-[#3B5B7E]/40 bg-white' : 'border-slate-200 bg-slate-50/50'
          }`}
        >
          <div className="flex items-center justify-between gap-2">
            <label className="inline-flex items-center gap-2 cursor-pointer select-none text-sm font-semibold text-slate-800">
              <input
                type="checkbox"
                checked={modifyEnabled}
                onChange={(e) => setModifyEnabled(e.target.checked)}
                className="rounded border-slate-300 cursor-pointer"
                disabled={loading}
              />
              <span>① 批量修改：数字 ×倍数</span>
            </label>
            <span className="text-[10px] text-slate-400">精确十进制，自动去末尾零</span>
          </div>

          <div className="grid grid-cols-1 gap-2.5 sm:grid-cols-3">
            <div className="sm:col-span-2">
              <label className="block text-[11px] font-medium text-slate-600 mb-1">
                匹配模板（用 *、xxx 或“数字”占位）
              </label>
              <input
                value={modifyTemplate}
                onChange={(e) => setModifyTemplate(e.target.value)}
                placeholder="例如：最高应答限价为*万元"
                disabled={!modifyEnabled || loading}
                className="w-full rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none disabled:bg-slate-100 disabled:text-slate-400"
              />
            </div>
            <div>
              <label className="block text-[11px] font-medium text-slate-600 mb-1">× 倍数</label>
              <input
                value={multiplier}
                onChange={(e) => setMultiplier(e.target.value)}
                placeholder="10000"
                disabled={!modifyEnabled || loading}
                className="w-full rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs tabular-nums focus:border-[#3B5B7E] focus:outline-none disabled:bg-slate-100 disabled:text-slate-400"
              />
            </div>
          </div>
        </div>

        {/* ② 批量删除：整段删除 */}
        <div
          className={`rounded-xl border p-4 space-y-3 transition-colors ${
            deleteEnabled ? 'border-[#3B5B7E]/40 bg-white' : 'border-slate-200 bg-slate-50/50'
          }`}
        >
          <div className="flex items-center justify-between gap-2">
            <label className="inline-flex items-center gap-2 cursor-pointer select-none text-sm font-semibold text-slate-800">
              <input
                type="checkbox"
                checked={deleteEnabled}
                onChange={(e) => setDeleteEnabled(e.target.checked)}
                className="rounded border-slate-300 cursor-pointer"
                disabled={loading}
              />
              <span>② 批量删除：整段删除</span>
            </label>
            <span className="text-[10px] text-slate-400">自动连带删除紧邻的一个标点</span>
          </div>

          <div>
            <label className="block text-[11px] font-medium text-slate-600 mb-1">
              删除模板（用 *、xxx 或“数字”占位）
            </label>
            <input
              value={deleteTemplate}
              onChange={(e) => setDeleteTemplate(e.target.value)}
              placeholder="例如：有效报价平均值为*万元"
              disabled={!deleteEnabled || loading}
              className="w-full rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none disabled:bg-slate-100 disabled:text-slate-400"
            />
          </div>
        </div>
      </div>

      {/* ③ 查找替换（可追加多组） */}
      <div
        className={`rounded-xl border p-4 space-y-3 transition-colors ${
          replaceEnabled ? 'border-[#3B5B7E]/40 bg-white' : 'border-slate-200 bg-slate-50/50'
        }`}
      >
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div className="flex flex-wrap items-center gap-4">
            <label className="inline-flex items-center gap-2 cursor-pointer select-none text-sm font-semibold text-slate-800">
              <input
                type="checkbox"
                checked={replaceEnabled}
                onChange={(e) => setReplaceEnabled(e.target.checked)}
                className="rounded border-slate-300 cursor-pointer"
                disabled={loading}
              />
              <span>③ 批量查找替换（支持多组 A→B）</span>
            </label>

            <label className="inline-flex items-center gap-1.5 cursor-pointer select-none text-xs text-slate-600">
              <input
                type="checkbox"
                checked={replaceUseRegex}
                onChange={(e) => setReplaceUseRegex(e.target.checked)}
                className="rounded border-slate-300 cursor-pointer"
                disabled={!replaceEnabled || loading}
              />
              <span>使用正则表达式（替换内容支持 \1 或 $1 引用分组）</span>
            </label>
          </div>
          <span className="text-[11px] text-slate-400 tabular-nums">已添加 {replacePairs.length} 组</span>
        </div>

        <div className="flex flex-col gap-2 sm:flex-row sm:items-center">
          <input
            value={findInput}
            onChange={(e) => setFindInput(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                e.preventDefault();
                handleAddPair();
              }
            }}
            placeholder="查找内容 *"
            disabled={!replaceEnabled || loading}
            className="flex-1 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none disabled:bg-slate-100"
          />
          <span className="text-xs text-slate-400 hidden sm:inline">→</span>
          <input
            value={replaceInput}
            onChange={(e) => setReplaceInput(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                e.preventDefault();
                handleAddPair();
              }
            }}
            placeholder="替换为（留空即删除查找内容）"
            disabled={!replaceEnabled || loading}
            className="flex-1 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none disabled:bg-slate-100"
          />
          <button
            type="button"
            onClick={handleAddPair}
            disabled={!replaceEnabled || loading || !findInput}
            className="inline-flex shrink-0 items-center justify-center gap-1 rounded-lg border border-[#3B5B7E]/35 bg-[#EBF2F7] px-3 py-1.5 text-xs font-medium text-[#3B5B7E] hover:bg-[#DCE8F2] disabled:opacity-45"
          >
            <Plus className="h-3.5 w-3.5" />
            <span>添加规则</span>
          </button>
        </div>

        {replacePairs.length > 0 && (
          <div className="overflow-hidden rounded-lg border border-slate-200">
            <table className="w-full border-collapse text-left text-xs">
              <thead>
                <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                  <th className="px-3 py-1.5 font-medium w-10">#</th>
                  <th className="px-3 py-1.5 font-medium">查找</th>
                  <th className="px-3 py-1.5 font-medium">替换为</th>
                  <th className="px-3 py-1.5 font-medium w-24 text-right">操作</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 bg-white">
                {replacePairs.map((pair, idx) => (
                  <tr key={idx} className="hover:bg-slate-50/70">
                    <td className="px-3 py-1.5 text-slate-400 tabular-nums">{idx + 1}</td>
                    <td className="px-3 py-1.5 font-mono text-slate-800">{pair.find}</td>
                    <td className="px-3 py-1.5 font-mono text-slate-600">
                      {pair.replace === '' ? <span className="text-slate-400">（删空）</span> : pair.replace}
                    </td>
                    <td className="px-3 py-1.5 text-right whitespace-nowrap">
                      <button
                        type="button"
                        onClick={() => handleEditPair(idx)}
                        disabled={!replaceEnabled || loading}
                        className="rounded p-1 text-slate-400 hover:text-[#3B5B7E] mr-1"
                        title="取回编辑"
                      >
                        <Edit3 className="h-3.5 w-3.5" />
                      </button>
                      <button
                        type="button"
                        onClick={() => handleRemovePair(idx)}
                        disabled={!replaceEnabled || loading}
                        className="rounded p-1 text-slate-400 hover:text-[#A94442]"
                        title="删除"
                      >
                        <Trash2 className="h-3.5 w-3.5" />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* ④ 提取评标表格到 Excel + Run Button */}
      <div className="rounded-xl border border-slate-200 bg-white p-4 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
        <div className="space-y-1">
          <label className="inline-flex items-center gap-2 cursor-pointer select-none text-sm font-semibold text-slate-800">
            <input
              type="checkbox"
              checked={extractEnabled}
              onChange={(e) => setExtractEnabled(e.target.checked)}
              className="rounded border-slate-300 cursor-pointer"
              disabled={loading}
            />
            <span>④ 提取评标表格到 Excel（输出同名 _表格.xlsx）</span>
          </label>
          <p className="text-[11px] text-slate-500 pl-6">
            自动识别按「分标-分包」排布的候选人评分表（含流标包标记），合并导出为含「分标名称 / 分包 / 序号 / 应答人 / 应答报价（万元）/ 技术·价格·商务加权分 / 总分 / 排序 / 推荐的成交候选人」的 Excel。
          </p>
        </div>

        <button
          type="button"
          onClick={handleRun}
          disabled={loading || files.length === 0}
          className="inline-flex shrink-0 items-center justify-center gap-1.5 rounded-lg bg-[#3B5B7E] px-5 py-2.5 text-xs font-medium text-white hover:bg-[#2E4866] disabled:opacity-45 transition-colors"
        >
          <Play className="h-3.5 w-3.5" />
          <span>{loading ? '正在处理中...' : '开始批量处理'}</span>
        </button>
      </div>

      {/* Error Banner */}
      {errorMsg && (
        <div className="flex items-start gap-2 rounded-xl border border-[#A94442]/30 bg-[#FEF2F2] px-4 py-3 text-xs text-[#A94442]">
          <AlertTriangle className="h-4 w-4 shrink-0 mt-0.5" />
          <span className="whitespace-pre-line">{errorMsg}</span>
        </div>
      )}

      {/* Results & Downloads */}
      {result && (
        <div className="space-y-4">
          <div className="rounded-xl border border-[#2F6F5E]/30 bg-white overflow-hidden">
            <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 bg-[#ECFDF5]/60 px-4 py-3">
              <div className="flex items-center gap-2 text-xs font-semibold text-[#2F6F5E]">
                <CheckCircle2 className="h-4 w-4 shrink-0" />
                <span>
                  处理完成：成功 {result.okFiles} / {result.totalFiles} 个文件
                </span>
                {Object.keys(result.grandTotals).length > 0 && (
                  <span className="font-normal text-slate-600">
                    （
                    {Object.entries(result.grandTotals)
                      .map(([k, v]) => `${k} ${v} ${k === '提取表格' ? '行' : '处'}`)
                      .join('，')}
                    ）
                  </span>
                )}
              </div>

              {result.hasOutputs && (
                <a
                  href={`/api/toolbox/word-batch/${result.jobId}/download-all`}
                  className="inline-flex items-center gap-1.5 rounded-lg bg-[#2F6F5E] px-3.5 py-1.5 text-xs font-medium text-white hover:bg-[#24594B] transition-colors"
                >
                  <Archive className="h-3.5 w-3.5" />
                  <span>全部打包下载 (.zip 含日志)</span>
                </a>
              )}
            </div>

            <div className="divide-y divide-slate-100">
              {result.files.map((item, idx) => (
                <div key={idx} className="px-4 py-3 flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
                  <div className="min-w-0 space-y-0.5">
                    <div className="flex items-center gap-2 text-xs font-medium text-slate-800">
                      <span
                        className={`inline-block h-2 w-2 rounded-full shrink-0 ${
                          item.ok ? 'bg-[#2F6F5E]' : 'bg-[#A94442]'
                        }`}
                      />
                      <span className="truncate">{item.originalName}</span>
                      {item.convertedFromDoc && (
                        <span className="rounded bg-[#EBF2F7] px-1.5 py-0.5 text-[10px] font-medium text-[#3B5B7E]">
                          .doc 自动转换
                        </span>
                      )}
                    </div>
                    {item.ok ? (
                      <div className="text-[11px] text-slate-500 pl-4">
                        {Object.entries(item.counts)
                          .map(([k, v]) => `${k} ${v} ${k === '提取表格' ? '行' : '处'}`)
                          .join(' · ')}
                      </div>
                    ) : (
                      <div className="text-[11px] text-[#A94442] pl-4">{item.error}</div>
                    )}
                  </div>

                  {item.outputs.length > 0 && (
                    <div className="flex flex-wrap items-center gap-1.5 shrink-0 pl-4 sm:pl-0">
                      {item.outputs.map((out) => (
                        <a
                          key={out.key}
                          href={`/api/toolbox/word-batch/${result.jobId}/download/${out.key}`}
                          className={`inline-flex items-center gap-1 rounded-md border px-2.5 py-1 text-[11px] font-medium transition-colors ${
                            out.kind === 'xlsx'
                              ? 'border-[#2F6F5E]/30 bg-[#ECFDF5]/60 text-[#2F6F5E] hover:bg-[#D1FAE5]'
                              : 'border-[#3B5B7E]/30 bg-[#EBF2F7]/60 text-[#3B5B7E] hover:bg-[#DCE8F2]'
                          }`}
                        >
                          {out.kind === 'xlsx' ? (
                            <FileSpreadsheet className="h-3.5 w-3.5" />
                          ) : (
                            <Download className="h-3.5 w-3.5" />
                          )}
                          <span className="max-w-[220px] truncate">{out.filename}</span>
                        </a>
                      ))}
                    </div>
                  )}
                </div>
              ))}
            </div>
          </div>

          {/* Processing Log */}
          <div className="rounded-xl border border-slate-200 bg-white overflow-hidden">
            <div className="border-b border-slate-200 bg-slate-50/70 px-4 py-2 text-xs font-semibold text-slate-600">
              处理日志
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
