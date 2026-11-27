import { useCallback, useEffect, useState } from 'react';
import {
  AlertTriangle,
  Award,
  Download,
  FileSpreadsheet,
  Play,
  Search,
  Trash2,
  Upload,
} from 'lucide-react';
import { api } from '../api/client';
import { paginate, PaginationBar, usePagination } from './PaginationBar';
import type { PerformanceState } from '../types';

export function PerformanceScorePanel({ batchId }: { batchId: string }) {
  const [state, setState] = useState<PerformanceState | null>(null);
  const [bidFile, setBidFile] = useState<File | null>(null);
  const [scoreFile, setScoreFile] = useState<File | null>(null);
  const [search, setSearch] = useState('');
  const [onlyReview, setOnlyReview] = useState(false);
  const [running, setRunning] = useState(false);
  const [error, setError] = useState('');

  const pager = usePagination(`${batchId}-perf-${search}-${onlyReview}`);

  const load = useCallback(async () => {
    try {
      const data = await api.getPerformanceState(batchId);
      setState(data);
    } catch {
      // ignore
    }
  }, [batchId]);

  useEffect(() => {
    load();
    setBidFile(null);
    setScoreFile(null);
    setError('');
  }, [load]);

  async function handleRun() {
    setRunning(true);
    setError('');
    try {
      const res = await api.runPerformanceScore(batchId, {
        bids: bidFile || undefined,
        scores: scoreFile || undefined,
      });
      setState(res);
      setBidFile(null);
      setScoreFile(null);
    } catch (err) {
      setError((err as Error).message);
    } finally {
      setRunning(false);
    }
  }

  async function handleClear() {
    if (!window.confirm('确认清除当前批次已保存的履约分生成文件与结果吗？')) return;
    await api.clearPerformanceScore(batchId);
    setBidFile(null);
    setScoreFile(null);
    setError('');
    await load();
  }

  const rows = state?.result?.rows || [];
  const q = search.trim().toLowerCase();
  const filtered = rows.filter((r) => {
    if (onlyReview && r.match_note === '已匹配') return false;
    if (!q) return true;
    return `${r.lot_name} ${r.pack_name} ${r.bidder_name} ${r.unit_name} ${r.department} ${r.match_note}`
      .toLowerCase()
      .includes(q);
  });

  return (
    <div className="space-y-4">
      <div className="rounded-xl border border-slate-200 bg-white p-4 space-y-4">
        <div className="flex flex-wrap items-start justify-between gap-2">
          <div>
            <h3 className="text-sm font-semibold text-slate-800 flex items-center gap-1.5">
              <Award className="h-4 w-4 text-[#3B5B7E]" />
              <span>履约分生成</span>
            </h3>
            <p className="mt-0.5 text-xs text-slate-500">
              严格按原工具箱逻辑关联开标表与履约评价表（支持「供应商」「供应商+标段」「供应商+所属部门」三类关联），按等级/分数区间映射 1–5 分并标记需复核项。
            </p>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <a
              href={`/api/evaluation/batch/${batchId}/performance-score/template?kind=bids`}
              className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-medium text-slate-600 hover:border-[#3B5B7E] hover:text-[#3B5B7E]"
            >
              <FileSpreadsheet className="h-3.5 w-3.5" />
              <span>下载开标表模板</span>
            </a>
            <a
              href={`/api/evaluation/batch/${batchId}/performance-score/template?kind=scores`}
              className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-medium text-slate-600 hover:border-[#3B5B7E] hover:text-[#3B5B7E]"
            >
              <FileSpreadsheet className="h-3.5 w-3.5" />
              <span>下载履约评价模板</span>
            </a>
          </div>
        </div>

        <div className="grid grid-cols-1 gap-3 md:grid-cols-3">
          <div className="rounded-lg border border-slate-200 bg-slate-50/50 p-3 space-y-2">
            <div className="text-xs font-semibold text-slate-700">1. 开标表（含单位名称/归口部门）</div>
            <div className="text-[11px] text-slate-500 truncate">
              {bidFile
                ? `待上传：${bidFile.name}`
                : state?.files.bids
                  ? `已上传：${state.files.bids.originalName}`
                  : '尚未选择开标表'}
            </div>
            <label className="inline-flex cursor-pointer items-center gap-1 rounded-md border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:border-[#3B5B7E] hover:text-[#3B5B7E]">
              <Upload className="h-3.5 w-3.5" />
              <span>选择开标表</span>
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) => setBidFile(e.target.files?.[0] || null)}
                className="hidden"
              />
            </label>
          </div>

          <div className="rounded-lg border border-slate-200 bg-slate-50/50 p-3 space-y-2">
            <div className="text-xs font-semibold text-slate-700">2. 履约评价表</div>
            <div className="text-[11px] text-slate-500 truncate">
              {scoreFile
                ? `待上传：${scoreFile.name}`
                : state?.files.scores
                  ? `已上传：${state.files.scores.originalName}`
                  : '尚未选择履约评价表'}
            </div>
            <label className="inline-flex cursor-pointer items-center gap-1 rounded-md border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:border-[#3B5B7E] hover:text-[#3B5B7E]">
              <Upload className="h-3.5 w-3.5" />
              <span>选择履约评价表</span>
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) => setScoreFile(e.target.files?.[0] || null)}
                className="hidden"
              />
            </label>
          </div>

          <div className="rounded-lg border border-slate-200 bg-slate-50/50 p-3 flex flex-col justify-between space-y-2">
            <div className="text-[11px] text-slate-500 leading-relaxed">
              评分映射：优秀/A/≥90 → 5分；良好/B/(80,90) → 4分；一般/C/(60,80] → 3分；及格/D/(30,60] → 2分；不及格/E/&lt;30 → 1分；未匹配或边界值默认 4 分并提示复核。
            </div>
            <div className="flex items-center gap-2">
              <button
                type="button"
                disabled={running}
                onClick={handleRun}
                className="inline-flex flex-1 items-center justify-center gap-1.5 rounded-md bg-[#3B5B7E] px-3 py-1.5 text-xs font-medium text-white hover:bg-[#2E4866] disabled:opacity-50"
              >
                <Play className="h-3.5 w-3.5" />
                <span>{running ? '生成中...' : '生成履约分'}</span>
              </button>
              {state?.result && (
                <button
                  type="button"
                  onClick={handleClear}
                  className="inline-flex items-center gap-1 rounded-md border border-[#A94442]/25 bg-[#FEF2F2] px-2.5 py-1.5 text-xs font-medium text-[#A94442] hover:bg-[#FEE2E2]"
                >
                  <Trash2 className="h-3.5 w-3.5" />
                </button>
              )}
            </div>
          </div>
        </div>

        {error && (
          <div className="rounded-lg border border-[#A94442]/30 bg-[#FEF2F2] px-3 py-2 text-xs text-[#A94442] flex items-center gap-2">
            <AlertTriangle className="h-4 w-4 shrink-0" />
            <span>{error}</span>
          </div>
        )}
      </div>

      {state?.result && (
        <div className="rounded-xl border border-slate-200 bg-white p-4 space-y-3">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="flex flex-wrap items-center gap-2 text-xs">
              <span className="font-semibold text-slate-800">履约分匹配结果</span>
              <span className="rounded bg-[#ECFDF5] px-2 py-0.5 font-medium text-[#2F6F5E] tabular-nums">
                精确匹配 {state.result.matchedCount} 条
              </span>
              {state.result.reviewCount > 0 && (
                <span className="rounded bg-[#FEF9EC] px-2 py-0.5 font-medium text-[#9A6B26] tabular-nums">
                  需复核 / 默认4分 {state.result.reviewCount} 条
                </span>
              )}
            </div>
            <div className="flex flex-wrap items-center gap-2">
              <label className="inline-flex items-center gap-1.5 text-xs text-slate-600 cursor-pointer">
                <input
                  type="checkbox"
                  checked={onlyReview}
                  onChange={(e) => setOnlyReview(e.target.checked)}
                  className="rounded border-slate-300"
                />
                <span>仅看需复核</span>
              </label>
              <div className="relative w-52">
                <Search className="pointer-events-none absolute left-2.5 top-2 h-3.5 w-3.5 text-slate-400" />
                <input
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  placeholder="搜索投标人、分标或单位..."
                  className="w-full rounded-lg border border-slate-200 py-1.5 pl-8 pr-3 text-xs focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>
              <a
                href={`/api/evaluation/batch/${batchId}/performance-score/export`}
                className="inline-flex items-center gap-1 rounded-md bg-[#2F6F5E] px-3 py-1.5 text-xs font-medium text-white hover:bg-[#24594B]"
              >
                <Download className="h-3.5 w-3.5" />
                <span>导出履约分 (.xlsx)</span>
              </a>
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full border-collapse text-left text-xs">
              <thead>
                <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                  <th className="px-2.5 py-2 font-medium whitespace-nowrap">分标 / 包名</th>
                  <th className="px-2.5 py-2 font-medium">投标人名称</th>
                  <th className="px-2.5 py-2 font-medium whitespace-nowrap">单位 / 归口部门</th>
                  <th className="px-2.5 py-2 font-medium whitespace-nowrap">综合评价</th>
                  <th className="px-2.5 py-2 font-medium whitespace-nowrap">履约分</th>
                  <th className="px-2.5 py-2 font-medium">匹配说明</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {paginate(filtered, pager.page, pager.pageSize).map((r, idx) => (
                  <tr
                    key={idx}
                    className={r.match_note !== '已匹配' ? 'bg-[#FEF9EC]/40 hover:bg-[#FEF9EC]/70' : 'hover:bg-slate-50/70'}
                  >
                    <td className="px-2.5 py-2 whitespace-nowrap">
                      <div className="font-medium text-slate-800">{r.lot_name}</div>
                      <div className="text-[11px] text-slate-400">{r.pack_name}</div>
                    </td>
                    <td className="px-2.5 py-2 font-medium text-slate-800">{r.bidder_name}</td>
                    <td className="px-2.5 py-2 whitespace-nowrap text-slate-600">
                      <div>{r.unit_name}</div>
                      <div className="text-[11px] text-slate-400">{r.department || '—'}</div>
                    </td>
                    <td className="px-2.5 py-2 whitespace-nowrap text-slate-700">{r.evaluation || '—'}</td>
                    <td className="px-2.5 py-2 whitespace-nowrap tabular-nums font-semibold text-[#3B5B7E]">{r.score}</td>
                    <td className="px-2.5 py-2 text-slate-600">
                      {r.match_note === '已匹配' ? (
                        <span className="rounded bg-[#ECFDF5] px-1.5 py-0.5 text-[11px] font-medium text-[#2F6F5E]">
                          已匹配
                        </span>
                      ) : (
                        <span className="text-[#9A6B26]">{r.match_note}</span>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {filtered.length > 20 && (
            <div className="overflow-hidden rounded-lg border border-slate-200">
              <PaginationBar
                total={filtered.length}
                page={pager.page}
                pageSize={pager.pageSize}
                onPageChange={pager.setPage}
                onPageSizeChange={pager.setPageSize}
              />
            </div>
          )}
        </div>
      )}
    </div>
  );
}
