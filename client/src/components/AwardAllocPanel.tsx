import { useCallback, useEffect, useState } from 'react';
import {
  AlertTriangle,
  BookOpen,
  Download,
  FileSpreadsheet,
  GitBranch,
  Play,
  Search,
  Trash2,
  Upload,
} from 'lucide-react';
import { api } from '../api/client';
import { paginate, PaginationBar, usePagination } from './PaginationBar';
import type { AwardAllocState, AwardMode } from '../types';

const MODES: AwardMode[] = ['无限授', '一轮一包', '一轮多包'];

function getAwardHelpText(mode: AwardMode): string {
  const common =
    '输入｜采用排序表中的“排序”，不使用推荐排序，不重新计算评分；同包并列排名需先确定唯一排序。兼容原版经济排序表及标准 6 列排序表。';
  if (mode === '无限授') {
    return `${common}\n原则｜每个包选择排名最靠前的供应商，不限制同一供应商的中包数量；此模式不使用限授组配置表。`;
  }
  const rule =
    mode === '一轮一包'
      ? '每轮各包选择未达上限的最高排名候选，优先授出候选报价最高的一个包；更新中包数后重新选择候选。'
      : '每轮各包选择未达上限的最高排名候选，按供应商归并其候选包；每家按优先顺序一次获得不超过剩余名额的包，再进入下一轮。';
  return `${common}\n原则｜${rule}\n包优先级｜当前候选投标总价从高到低 → 授标顺序从小到大 → 输入文件/工作表顺序（报价与顺序相同时输出复核提示）。\n限授组配置｜每行一个包；同组最大中包数须一致（累计上限，各组独立计数）；未配置的包按无限授处理。限授组内报价单位与口径必须一致。`;
}

export function AwardAllocPanel({ batchId }: { batchId: string }) {
  const [state, setState] = useState<AwardAllocState | null>(null);
  const [rankingFiles, setRankingFiles] = useState<File[]>([]);
  const [groupFile, setGroupFile] = useState<File | null>(null);
  const [mode, setMode] = useState<AwardMode>('无限授');
  const [showHelp, setShowHelp] = useState(false);
  const [search, setSearch] = useState('');
  const [running, setRunning] = useState(false);
  const [error, setError] = useState('');

  const pager = usePagination(`${batchId}-aw-${search}`);

  const load = useCallback(async () => {
    try {
      const data = await api.getAwardAllocState(batchId);
      setState(data);
      if (data.config?.mode) setMode(data.config.mode);
    } catch {
      // ignore
    }
  }, [batchId]);

  useEffect(() => {
    load();
    setRankingFiles([]);
    setGroupFile(null);
    setError('');
  }, [load]);

  async function handleRun() {
    setRunning(true);
    setError('');
    try {
      const res = await api.runAwardAlloc(batchId, {
        files: rankingFiles.length ? rankingFiles : undefined,
        groups: groupFile || undefined,
        mode,
      });
      setState(res);
      setRankingFiles([]);
      setGroupFile(null);
    } catch (err) {
      setError((err as Error).message);
    } finally {
      setRunning(false);
    }
  }

  async function handleClear() {
    if (!window.confirm('确认清除当前批次已保存的授标分配文件与结果吗？')) return;
    await api.clearAwardAlloc(batchId);
    setRankingFiles([]);
    setGroupFile(null);
    setError('');
    await load();
  }

  const rows = state?.result?.rows || [];
  const q = search.trim().toLowerCase();
  const filtered = rows.filter(
    (r) =>
      !q ||
      `${r.section_code} ${r.section_name} ${r.pack_name} ${r.limit_group} ${r.winner} ${r.status}`
        .toLowerCase()
        .includes(q)
  );

  return (
    <div className="space-y-4">
      <div className="rounded-xl border border-slate-200 bg-white p-4 space-y-4">
        <div className="flex flex-wrap items-start justify-between gap-2">
          <div>
            <h3 className="text-sm font-semibold text-slate-800 flex items-center gap-1.5">
              <GitBranch className="h-4 w-4 text-[#3B5B7E]" />
              <span>授标分配（无限授 / 一轮一包 / 一轮多包）</span>
            </h3>
            <p className="mt-0.5 text-xs text-slate-500">
              严格以输入技术经济排序表的「排序」为权威依据进行授标推演，支持原版经济排序表与标准 6 列排序表。
            </p>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <button
              type="button"
              onClick={() => setShowHelp((v) => !v)}
              className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-slate-50 px-2.5 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-100"
            >
              <BookOpen className="h-3.5 w-3.5 text-[#3B5B7E]" />
              <span>{showHelp ? '收起规则说明' : '授标规则说明'}</span>
            </button>
            <a
              href={`/api/evaluation/batch/${batchId}/award-alloc/template?kind=ranking`}
              className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-medium text-slate-600 hover:border-[#3B5B7E] hover:text-[#3B5B7E]"
            >
              <FileSpreadsheet className="h-3.5 w-3.5" />
              <span>下载排序表模板</span>
            </a>
            {mode !== '无限授' && (
              <a
                href={`/api/evaluation/batch/${batchId}/award-alloc/template?kind=groups`}
                className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-medium text-slate-600 hover:border-[#3B5B7E] hover:text-[#3B5B7E]"
              >
                <FileSpreadsheet className="h-3.5 w-3.5" />
                <span>下载限授配置模板</span>
              </a>
            )}
          </div>
        </div>

        {showHelp && (
          <div className="rounded-lg border border-[#3B5B7E]/20 bg-[#EBF2F7]/40 p-3 text-xs text-slate-700 whitespace-pre-line leading-relaxed">
            {getAwardHelpText(mode)}
          </div>
        )}

        <div className="grid grid-cols-1 gap-3 md:grid-cols-3">
          <div className="rounded-lg border border-slate-200 bg-slate-50/50 p-3 space-y-2">
            <div className="text-xs font-semibold text-slate-700">1. 授标方式</div>
            <div className="flex flex-wrap gap-1.5">
              {MODES.map((m) => (
                <button
                  key={m}
                  type="button"
                  onClick={() => setMode(m)}
                  className={`rounded-md px-2.5 py-1.5 text-xs font-medium transition-colors ${
                    mode === m
                      ? 'bg-[#3B5B7E] text-white'
                      : 'border border-slate-200 bg-white text-slate-600 hover:bg-slate-100'
                  }`}
                >
                  {m}
                </button>
              ))}
            </div>
          </div>

          <div className="rounded-lg border border-slate-200 bg-slate-50/50 p-3 space-y-2">
            <div className="text-xs font-semibold text-slate-700">2. 技术经济排序表（支持多选）</div>
            <div className="text-[11px] text-slate-500 truncate">
              {rankingFiles.length > 0
                ? `待上传 ${rankingFiles.length} 份：${rankingFiles.map((f) => f.name).join('、')}`
                : state?.files.rankings.length
                  ? `已上传 ${state.files.rankings.length} 份：${state.files.rankings.map((f) => f.originalName).join('、')}`
                  : '尚未选择排序表'}
            </div>
            <label className="inline-flex cursor-pointer items-center gap-1 rounded-md border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:border-[#3B5B7E] hover:text-[#3B5B7E]">
              <Upload className="h-3.5 w-3.5" />
              <span>选择排序表</span>
              <input
                type="file"
                multiple
                accept=".xlsx,.xls"
                onChange={(e) => setRankingFiles(Array.from(e.target.files || []))}
                className="hidden"
              />
            </label>
          </div>

          <div className="rounded-lg border border-slate-200 bg-slate-50/50 p-3 space-y-2">
            <div className="text-xs font-semibold text-slate-700">
              3. 限授组配置表{mode === '无限授' ? '（无限授模式无需上传）' : '（必填 1 份）'}
            </div>
            {mode !== '无限授' && (
              <div className="flex items-center justify-between gap-2">
                <span className="text-[11px] text-slate-500 truncate">
                  {groupFile
                    ? `待上传：${groupFile.name}`
                    : state?.files.groups
                      ? `已上传：${state.files.groups.originalName}`
                      : '尚未选择配置表'}
                </span>
                <label className="inline-flex shrink-0 cursor-pointer items-center gap-1 rounded-md border border-slate-200 bg-white px-2.5 py-1 text-xs font-medium text-slate-700 hover:border-[#3B5B7E] hover:text-[#3B5B7E]">
                  <Upload className="h-3 w-3" />
                  <span>选择</span>
                  <input
                    type="file"
                    accept=".xlsx,.xls"
                    onChange={(e) => setGroupFile(e.target.files?.[0] || null)}
                    className="hidden"
                  />
                </label>
              </div>
            )}
            <div className="flex items-center gap-2 pt-0.5">
              <button
                type="button"
                disabled={running}
                onClick={handleRun}
                className="inline-flex flex-1 items-center justify-center gap-1.5 rounded-md bg-[#3B5B7E] px-3 py-1.5 text-xs font-medium text-white hover:bg-[#2E4866] disabled:opacity-50"
              >
                <Play className="h-3.5 w-3.5" />
                <span>{running ? '推演中...' : '执行授标分配'}</span>
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
              <span className="font-semibold text-slate-800">授标分配结果（{state.result.mode}）</span>
              <span className="rounded bg-[#ECFDF5] px-2 py-0.5 font-medium text-[#2F6F5E] tabular-nums">
                拟授标 {state.result.awardedCount} 包
              </span>
              {state.result.unallocatedCount > 0 && (
                <span className="rounded bg-[#FEF2F2] px-2 py-0.5 font-medium text-[#A94442] tabular-nums">
                  未授标 / 无候选 {state.result.unallocatedCount} 包
                </span>
              )}
            </div>
            <div className="flex flex-wrap items-center gap-2">
              <div className="relative w-52">
                <Search className="pointer-events-none absolute left-2.5 top-2 h-3.5 w-3.5 text-slate-400" />
                <input
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  placeholder="搜索分标、包名、限授组或拟授标人..."
                  className="w-full rounded-lg border border-slate-200 py-1.5 pl-8 pr-3 text-xs focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>
              <a
                href={`/api/evaluation/batch/${batchId}/award-alloc/export`}
                className="inline-flex items-center gap-1 rounded-md bg-[#2F6F5E] px-3 py-1.5 text-xs font-medium text-white hover:bg-[#24594B]"
              >
                <Download className="h-3.5 w-3.5" />
                <span>导出授标结果 (.xlsx)</span>
              </a>
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full border-collapse text-left text-xs">
              <thead>
                <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                  <th className="px-2.5 py-2 font-medium whitespace-nowrap">分标 / 包名</th>
                  <th className="px-2.5 py-2 font-medium whitespace-nowrap">限授组</th>
                  <th className="px-2.5 py-2 font-medium">拟授标人</th>
                  <th className="px-2.5 py-2 font-medium whitespace-nowrap">状态</th>
                  <th className="px-2.5 py-2 font-medium whitespace-nowrap">轮次</th>
                  <th className="px-2.5 py-2 font-medium">说明</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {paginate(filtered, pager.page, pager.pageSize).map((r, idx) => (
                  <tr
                    key={idx}
                    className={r.status !== '拟授标' ? 'bg-[#FEF2F2]/30 hover:bg-[#FEF2F2]/55' : 'hover:bg-slate-50/70'}
                  >
                    <td className="px-2.5 py-2 whitespace-nowrap">
                      <div className="font-mono text-slate-700">{r.section_code}</div>
                      <div className="text-[11px] text-slate-400">
                        {r.section_name} · {r.pack_name}
                      </div>
                    </td>
                    <td className="px-2.5 py-2 whitespace-nowrap text-slate-600">{r.limit_group || '—'}</td>
                    <td className="px-2.5 py-2 font-medium text-slate-800">{r.winner || '—'}</td>
                    <td className="px-2.5 py-2 whitespace-nowrap">
                      <span
                        className={`rounded px-1.5 py-0.5 text-[11px] font-medium ${
                          r.status === '拟授标'
                            ? 'bg-[#ECFDF5] text-[#2F6F5E]'
                            : 'bg-[#FEF2F2] text-[#A94442]'
                        }`}
                      >
                        {r.status}
                      </span>
                    </td>
                    <td className="px-2.5 py-2 whitespace-nowrap tabular-nums text-slate-600">{r.round || '—'}</td>
                    <td className="px-2.5 py-2 text-[11px] text-slate-500">{r.note || '—'}</td>
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
