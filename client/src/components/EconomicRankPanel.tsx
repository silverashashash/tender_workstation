import { useCallback, useEffect, useState } from 'react';
import {
  AlertTriangle,
  CheckCircle2,
  Download,
  ListChecks,
  Play,
  Search,
  Trash2,
  Upload,
} from 'lucide-react';
import { api } from '../api/client';
import { paginate, PaginationBar, usePagination } from './PaginationBar';
import type { EconomicRankState } from '../types';

export function EconomicRankPanel({ batchId }: { batchId: string }) {
  const [state, setState] = useState<EconomicRankState | null>(null);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [search, setSearch] = useState('');
  const [running, setRunning] = useState(false);
  const [error, setError] = useState('');

  const pager = usePagination(`${batchId}-er-${search}`);

  const load = useCallback(async () => {
    try {
      const data = await api.getEconomicRankState(batchId);
      setState(data);
    } catch {
      // ignore
    }
  }, [batchId]);

  useEffect(() => {
    load();
    setSelectedFiles([]);
    setError('');
  }, [load]);

  async function handleRun() {
    if (selectedFiles.length === 0) {
      setError('请先选择至少一份经济排序表（.xlsx / .xls）');
      return;
    }
    setRunning(true);
    setError('');
    try {
      const res = await api.runEconomicRankCheck(batchId, selectedFiles);
      setState(res);
      setSelectedFiles([]);
    } catch (err) {
      setError((err as Error).message);
    } finally {
      setRunning(false);
    }
  }

  async function handleClear() {
    if (!window.confirm('确认清除当前批次已保存的经济排序表核验记录吗？')) return;
    await api.clearEconomicRankCheck(batchId);
    setSelectedFiles([]);
    setError('');
    await load();
  }

  const issues = state?.result?.issues || [];
  const q = search.trim().toLowerCase();
  const filtered = issues.filter(
    (x) =>
      !q ||
      `${x.section} ${x.pack} ${x.supplier} ${x.reason}`.toLowerCase().includes(q)
  );

  return (
    <div className="space-y-4">
      <div className="rounded-xl border border-slate-200 bg-white p-4 space-y-4">
        <div>
          <h3 className="text-sm font-semibold text-slate-800 flex items-center gap-1.5">
            <ListChecks className="h-4 w-4 text-[#3B5B7E]" />
            <span>经济排序表核验</span>
          </h3>
          <p className="mt-0.5 text-xs text-slate-500">
            校验系统导出的经济排序表：先排查技术/商务/价格分为 0 的异常记录，再严格按「总分降序 → 技术分降序 → 投标总价升序 → 商务分降序」复核排序一致性。
          </p>
        </div>

        <div className="flex flex-col gap-3 rounded-lg border border-slate-200 bg-slate-50/50 p-3.5 sm:flex-row sm:items-center sm:justify-between">
          <div className="space-y-1 min-w-0">
            <div className="text-xs font-semibold text-slate-700">选择经济排序表（支持多选）</div>
            <div className="text-[11px] text-slate-500 truncate">
              {selectedFiles.length > 0
                ? `待核验 ${selectedFiles.length} 份：${selectedFiles.map((f) => f.name).join('、')}`
                : state?.files.length
                  ? `已核验 ${state.files.length} 份：${state.files.map((f) => f.originalName).join('、')}`
                  : '尚未选择文件'}
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2 shrink-0">
            <label className="inline-flex cursor-pointer items-center gap-1 rounded-md border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:border-[#3B5B7E] hover:text-[#3B5B7E]">
              <Upload className="h-3.5 w-3.5" />
              <span>选择排序表</span>
              <input
                type="file"
                multiple
                accept=".xlsx,.xls"
                onChange={(e) => setSelectedFiles(Array.from(e.target.files || []))}
                className="hidden"
              />
            </label>
            <button
              type="button"
              disabled={running || selectedFiles.length === 0}
              onClick={handleRun}
              className="inline-flex items-center gap-1.5 rounded-md bg-[#3B5B7E] px-3.5 py-1.5 text-xs font-medium text-white hover:bg-[#2E4866] disabled:opacity-50"
            >
              <Play className="h-3.5 w-3.5" />
              <span>{running ? '核验中...' : '开始核验'}</span>
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
              <span className="font-semibold text-slate-800">核验结论</span>
              <span className="rounded bg-slate-100 px-2 py-0.5 text-slate-600 tabular-nums">
                核验 {state.result.sheetCount} 个包 · {state.result.examinedCount} 条记录
              </span>
              {state.result.issueCount === 0 ? (
                <span className="inline-flex items-center gap-1 rounded bg-[#ECFDF5] px-2 py-0.5 font-medium text-[#2F6F5E]">
                  <CheckCircle2 className="h-3.5 w-3.5" /> 全部排序一致，无零分异常
                </span>
              ) : (
                <span className="rounded bg-[#FEF2F2] px-2 py-0.5 font-medium text-[#A94442] tabular-nums">
                  发现 {state.result.issueCount} 项异常
                </span>
              )}
            </div>

            <div className="flex flex-wrap items-center gap-2">
              {state.result.issueCount > 0 && (
                <div className="relative w-52">
                  <Search className="pointer-events-none absolute left-2.5 top-2 h-3.5 w-3.5 text-slate-400" />
                  <input
                    value={search}
                    onChange={(e) => setSearch(e.target.value)}
                    placeholder="搜索分标、包号或投标人..."
                    className="w-full rounded-lg border border-slate-200 py-1.5 pl-8 pr-3 text-xs focus:border-[#3B5B7E] focus:outline-none"
                  />
                </div>
              )}
              <a
                href={`/api/evaluation/batch/${batchId}/economic-rank/export`}
                className="inline-flex items-center gap-1 rounded-md bg-[#2F6F5E] px-3 py-1.5 text-xs font-medium text-white hover:bg-[#24594B]"
              >
                <Download className="h-3.5 w-3.5" />
                <span>导出经济排序核验 (.xlsx)</span>
              </a>
            </div>
          </div>

          {state.result.issueCount > 0 && (
            <>
              <div className="overflow-x-auto">
                <table className="w-full border-collapse text-left text-xs">
                  <thead>
                    <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                      <th className="px-2.5 py-2 font-medium whitespace-nowrap">分标名称</th>
                      <th className="px-2.5 py-2 font-medium whitespace-nowrap">包名称</th>
                      <th className="px-2.5 py-2 font-medium">投标人</th>
                      <th className="px-2.5 py-2 font-medium whitespace-nowrap">原排序</th>
                      <th className="px-2.5 py-2 font-medium whitespace-nowrap">核验排序</th>
                      <th className="px-2.5 py-2 font-medium">原因</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {paginate(filtered, pager.page, pager.pageSize).map((it, idx) => (
                      <tr key={idx} className="bg-[#FEF2F2]/30 hover:bg-[#FEF2F2]/55">
                        <td className="px-2.5 py-2 font-medium text-slate-800 whitespace-nowrap">{it.section}</td>
                        <td className="px-2.5 py-2 whitespace-nowrap text-slate-600">{it.pack}</td>
                        <td className="px-2.5 py-2 font-medium text-slate-800">{it.supplier}</td>
                        <td className="px-2.5 py-2 whitespace-nowrap tabular-nums text-slate-600">{it.original_rank}</td>
                        <td className="px-2.5 py-2 whitespace-nowrap tabular-nums font-semibold text-[#A94442]">
                          {it.verified_rank || '—'}
                        </td>
                        <td className="px-2.5 py-2">
                          <span className="rounded bg-[#FEF2F2] px-1.5 py-0.5 text-[11px] font-medium text-[#A94442]">
                            {it.reason}
                          </span>
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
            </>
          )}
        </div>
      )}
    </div>
  );
}
