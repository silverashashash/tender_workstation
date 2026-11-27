import { useState } from 'react';
import {
  AlertTriangle,
  Archive,
  CheckCircle2,
  Download,
  FileSearch,
  Info,
  Play,
  RotateCcw,
  Search,
  Upload,
  X,
} from 'lucide-react';
import { api } from '../api/client';
import { paginate, PaginationBar, usePagination } from './PaginationBar';

type ComplianceResponse = Awaited<ReturnType<typeof api.runChecklistCompliance>>;
type StatusFilter = 'all' | 'OK' | '缺失' | '异常' | '解压失败' | '注意';

export function ChecklistCompliancePanel() {
  const [files, setFiles] = useState<File[]>([]);
  const [keywordsInput, setKeywordsInput] = useState('工作量清单, 工程量清单');
  const [loading, setLoading] = useState(false);
  const [errorMsg, setErrorMsg] = useState('');
  const [result, setResult] = useState<ComplianceResponse | null>(null);

  const [statusFilter, setStatusFilter] = useState<StatusFilter>('all');
  const [search, setSearch] = useState('');
  const pager = usePagination(`checklist-${statusFilter}-${search}`);

  function handleSelectFiles(e: React.ChangeEvent<HTMLInputElement>) {
    const picked = Array.from(e.target.files || []).filter(
      (f) => f.name.toLowerCase().endsWith('.zip') && !f.name.startsWith('._')
    );
    e.target.value = '';
    if (picked.length === 0) return;
    setErrorMsg('');
    setResult(null);
    setFiles((prev) => {
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

  function removeFile(idx: number) {
    setFiles((prev) => prev.filter((_, i) => i !== idx));
    setResult(null);
  }

  function handleReset() {
    setFiles([]);
    setResult(null);
    setErrorMsg('');
    setStatusFilter('all');
    setSearch('');
  }

  async function handleRun() {
    if (files.length === 0) {
      setErrorMsg('请先选择至少一个 .zip 压缩包（支持批次总包或直接多选项目技术规范书压缩包）');
      return;
    }
    setLoading(true);
    setErrorMsg('');
    setResult(null);
    setStatusFilter('all');
    try {
      const res = await api.runChecklistCompliance({
        files,
        keywords: keywordsInput,
      });
      setResult(res);
    } catch (err) {
      setErrorMsg((err as Error).message);
    } finally {
      setLoading(false);
    }
  }

  const q = search.trim().toLowerCase();
  const filteredItems = (result?.items || []).filter((it) => {
    if (statusFilter !== 'all' && it.status !== statusFilter) return false;
    if (q) {
      const hay = `${it.batch} ${it.specId} ${it.status} ${it.fileName} ${it.note}`.toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });

  function badgeClass(status: string) {
    if (status === 'OK') return 'bg-[#ECFDF5] text-[#2F6F5E]';
    if (status === '注意') return 'bg-[#FEF9EC] text-[#9A6B26]';
    return 'bg-[#FEF2F2] text-[#A94442]';
  }

  return (
    <div className="space-y-4">
      {/* Intro */}
      <div className="rounded-xl border border-slate-200 bg-white px-4 py-3 flex flex-wrap items-center justify-between gap-2">
        <div className="flex items-start gap-2 text-xs text-slate-600">
          <Info className="h-4 w-4 shrink-0 text-[#3B5B7E] mt-0.5" />
          <span>
            递归解压招标项目压缩包（自动修复 Windows/WPS 压缩包内中文 GBK 乱码，支持多层嵌套 ZIP 及 7z/RAR），查找含指定关键词的清单文件并实际打开校验其完整性（拦截小于 1KB 空文件、损坏文件、空表及非 Excel 格式），一键生成带状态着色的<strong>《清单合规性检验报告.xlsx》</strong>。
          </span>
        </div>
        {files.length > 0 && (
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

      {/* Upload & Config */}
      <div className="rounded-xl border border-slate-200 bg-white p-4 space-y-3.5">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div className="flex items-center gap-2">
            <FileSearch className="h-4 w-4 text-[#3B5B7E]" />
            <h2 className="text-sm font-semibold text-slate-800">选择批次压缩包或项目技术规范书压缩包 (.zip)</h2>
            {files.length > 0 && (
              <span className="rounded-full bg-[#EBF2F7] px-2 py-0.5 text-[11px] font-medium text-[#3B5B7E] tabular-nums">
                已选 {files.length} 个压缩包
              </span>
            )}
          </div>
          <label className="inline-flex cursor-pointer items-center gap-1.5 rounded-lg border border-[#3B5B7E]/35 bg-[#EBF2F7]/60 px-3 py-1.5 text-xs font-medium text-[#3B5B7E] hover:bg-[#EBF2F7] transition-colors">
            <Upload className="h-3.5 w-3.5" />
            <span>{files.length > 0 ? '继续添加 ZIP' : '选择 .zip 压缩包'}</span>
            <input
              type="file"
              multiple
              accept=".zip"
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
                <Archive className="h-3.5 w-3.5 text-[#3B5B7E] shrink-0" />
                <span className="max-w-[260px] truncate" title={f.name}>
                  {f.name}
                </span>
                <span className="text-[10px] text-slate-400 tabular-nums">
                  {(f.size / (1024 * 1024)).toFixed(2)} MB
                </span>
                <button
                  type="button"
                  onClick={() => removeFile(idx)}
                  disabled={loading}
                  className="text-slate-400 hover:text-[#A94442]"
                >
                  <X className="h-3 w-3" />
                </button>
              </span>
            ))}
          </div>
        ) : (
          <div className="text-[11px] text-slate-400">
            支持两种上传方式：① 直接上传整个批次的总压缩包（如 <code className="rounded bg-slate-100 px-1">122617.zip</code>，自动向下探测内层项目 .zip 列表）；② 直接多选多个项目技术规范书 .zip 文件。
          </div>
        )}

        <div className="flex flex-col gap-3 border-t border-slate-100 pt-3 sm:flex-row sm:items-end sm:justify-between">
          <div className="flex-1 max-w-md">
            <label className="block text-[11px] font-medium text-slate-600 mb-1">
              清单文件名关键词（逗号或顿号分隔）
            </label>
            <input
              value={keywordsInput}
              onChange={(e) => setKeywordsInput(e.target.value)}
              disabled={loading}
              placeholder="工作量清单, 工程量清单"
              className="w-full rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
            />
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              type="button"
              onClick={handleRun}
              disabled={loading || files.length === 0}
              className="inline-flex items-center gap-1.5 rounded-lg bg-[#3B5B7E] px-4 py-2 text-xs font-medium text-white hover:bg-[#2E4866] disabled:opacity-45 transition-colors"
            >
              <Play className="h-3.5 w-3.5" />
              <span>{loading ? '正在递归解压并检验...' : '开始合规性检验'}</span>
            </button>

            {result && (
              <a
                href={`/api/toolbox/checklist-compliance/${result.jobId}/download`}
                className="inline-flex items-center gap-1.5 rounded-lg border border-[#2F6F5E]/35 bg-[#ECFDF5] px-4 py-2 text-xs font-medium text-[#2F6F5E] hover:bg-[#D1FAE5] transition-colors"
              >
                <Download className="h-3.5 w-3.5" />
                <span>下载检验报告 (.xlsx)</span>
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

      {/* Results Table */}
      {result && (
        <div className="space-y-4">
          <div className="rounded-xl border border-slate-200 bg-white overflow-hidden">
            <div className="flex flex-col gap-2 border-b border-slate-200 bg-slate-50/70 px-4 py-3 sm:flex-row sm:items-center sm:justify-between">
              <div className="flex flex-wrap items-center gap-1.5">
                {([
                  { key: 'all', label: `全部 (${result.counts.total})` },
                  { key: 'OK', label: `OK (${result.counts.ok})` },
                  { key: '缺失', label: `缺失 (${result.counts.missing})` },
                  { key: '异常', label: `异常 (${result.counts.abnormal})` },
                  { key: '解压失败', label: `解压失败 (${result.counts.unpackFailed})` },
                  { key: '注意', label: `注意 (${result.counts.notice})` },
                ] as const).map((tab) => (
                  <button
                    key={tab.key}
                    type="button"
                    onClick={() => setStatusFilter(tab.key)}
                    className={`rounded-md px-2.5 py-1 text-xs font-medium tabular-nums transition-colors ${
                      statusFilter === tab.key
                        ? tab.key === 'OK'
                          ? 'bg-[#2F6F5E] text-white'
                          : tab.key === 'all'
                            ? 'bg-slate-800 text-white'
                            : tab.key === '注意'
                              ? 'bg-[#9A6B26] text-white'
                              : 'bg-[#A94442] text-white'
                        : 'bg-white border border-slate-200 text-slate-600 hover:bg-slate-100'
                    }`}
                  >
                    {tab.label}
                  </button>
                ))}
              </div>

              <div className="relative w-full sm:w-64">
                <Search className="pointer-events-none absolute left-2.5 top-2 h-3.5 w-3.5 text-slate-400" />
                <input
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  placeholder="搜索批次、规范书ID或清单名..."
                  className="w-full rounded-lg border border-slate-200 bg-white py-1.5 pl-8 pr-3 text-xs focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full border-collapse text-left text-xs">
                <thead>
                  <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                    <th className="px-3 py-2 font-medium whitespace-nowrap">批次编号</th>
                    <th className="px-3 py-2 font-medium whitespace-nowrap">技术规范书 ID</th>
                    <th className="px-3 py-2 font-medium whitespace-nowrap">检验结果</th>
                    <th className="px-3 py-2 font-medium min-w-[200px]">清单文件名</th>
                    <th className="px-3 py-2 font-medium whitespace-nowrap">文件大小</th>
                    <th className="px-3 py-2 font-medium min-w-[180px]">说明</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {filteredItems.length === 0 ? (
                    <tr>
                      <td colSpan={6} className="py-8 text-center text-xs text-slate-400">
                        <span className="inline-flex items-center gap-1.5">
                          <CheckCircle2 className="h-4 w-4 text-[#2F6F5E]" />
                          当前筛选条件下无记录
                        </span>
                      </td>
                    </tr>
                  ) : (
                    paginate(filteredItems, pager.page, pager.pageSize).map((it, idx) => (
                      <tr key={idx} className="hover:bg-slate-50/70">
                        <td className="px-3 py-2 font-mono text-slate-700 whitespace-nowrap">{it.batch}</td>
                        <td className="px-3 py-2 font-mono font-medium text-slate-800 whitespace-nowrap">
                          {it.specId}
                        </td>
                        <td className="px-3 py-2 whitespace-nowrap">
                          <span className={`rounded px-2 py-0.5 text-[11px] font-medium ${badgeClass(it.status)}`}>
                            {it.status}
                          </span>
                        </td>
                        <td className="px-3 py-2 text-slate-700">
                          <div>{it.fileName || '—'}</div>
                          {it.relativePath && (
                            <div className="text-[10px] text-slate-400 truncate max-w-[360px]" title={it.relativePath}>
                              {it.relativePath}
                            </div>
                          )}
                        </td>
                        <td className="px-3 py-2 text-slate-500 tabular-nums whitespace-nowrap">
                          {it.fileSize !== null ? `${(it.fileSize / 1024).toFixed(1)} KB` : '—'}
                        </td>
                        <td className="px-3 py-2 text-slate-600">{it.note}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
            {filteredItems.length > 20 && (
              <PaginationBar
                total={filteredItems.length}
                page={pager.page}
                pageSize={pager.pageSize}
                onPageChange={pager.setPage}
                onPageSizeChange={pager.setPageSize}
              />
            )}
          </div>
        </div>
      )}
    </div>
  );
}
