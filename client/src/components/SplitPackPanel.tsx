import { useCallback, useEffect, useState } from 'react';
import {
  AlertTriangle,
  Boxes,
  Download,
  FileSpreadsheet,
  Play,
  Search,
  Trash2,
  Upload,
} from 'lucide-react';
import { api } from '../api/client';
import { paginate, PaginationBar, usePagination } from './PaginationBar';
import type { SplitPackMode, SplitPackState } from '../types';

export function SplitPackPanel({ batchId }: { batchId: string }) {
  const [state, setState] = useState<SplitPackState | null>(null);
  const [matFile, setMatFile] = useState<File | null>(null);
  const [ruleFile, setRuleFile] = useState<File | null>(null);
  const [mode, setMode] = useState<SplitPackMode>('比例');
  const [weightStep, setWeightStep] = useState('1');
  const [lengthStep, setLengthStep] = useState('1');
  const [viewTab, setViewTab] = useState<'summary' | 'detail'>('summary');
  const [search, setSearch] = useState('');
  const [running, setRunning] = useState(false);
  const [error, setError] = useState('');

  const detailPager = usePagination(`${batchId}-sp-${search}`);

  const load = useCallback(async () => {
    try {
      const data = await api.getSplitPackState(batchId);
      setState(data);
      if (data.config?.mode) setMode(data.config.mode);
      if (data.config?.weightStep) setWeightStep(data.config.weightStep);
      if (data.config?.lengthStep) setLengthStep(data.config.lengthStep);
    } catch {
      // ignore
    }
  }, [batchId]);

  useEffect(() => {
    load();
    setMatFile(null);
    setRuleFile(null);
    setError('');
  }, [load]);

  async function handleRun() {
    setRunning(true);
    setError('');
    try {
      const res = await api.runSplitPack(batchId, {
        materials: matFile || undefined,
        rules: ruleFile || undefined,
        mode,
        weightStep,
        lengthStep,
      });
      setState(res);
      setMatFile(null);
      setRuleFile(null);
    } catch (err) {
      setError((err as Error).message);
    } finally {
      setRunning(false);
    }
  }

  async function handleClear() {
    if (!window.confirm('确认清除当前批次已保存的智能分包文件与结果吗？')) return;
    await api.clearSplitPack(batchId);
    setMatFile(null);
    setRuleFile(null);
    setError('');
    await load();
  }

  const detailRows = state?.result?.detailRows || [];
  const q = search.trim().toLowerCase();
  const filteredDetails = detailRows.filter(
    (r) =>
      !q ||
      `${r.pack_no} ${r.material_code} ${r.description} ${r.extension} ${r.specification}`
        .toLowerCase()
        .includes(q)
  );

  return (
    <div className="space-y-4">
      <div className="rounded-xl border border-slate-200 bg-white p-4 space-y-4">
        <div className="flex flex-wrap items-start justify-between gap-2">
          <div>
            <h3 className="text-sm font-semibold text-slate-800 flex items-center gap-1.5">
              <Boxes className="h-4 w-4 text-[#3B5B7E]" />
              <span>智能分包（三阶段数量拆分守恒算法）</span>
            </h3>
            <p className="mt-0.5 text-xs text-slate-500">
              严格复刻原工具箱三阶段分包逻辑（保底单位分配 → 冻结剩余按比例切分 → 高单价优先轮询兜底），确保数量与预算总额绝对守恒，支持按「比例」或「金额」目标分包。
            </p>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <a
              href={`/api/evaluation/batch/${batchId}/split-pack/template?kind=materials`}
              className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-medium text-slate-600 hover:border-[#3B5B7E] hover:text-[#3B5B7E]"
            >
              <FileSpreadsheet className="h-3.5 w-3.5" />
              <span>下载物料表模板</span>
            </a>
            <a
              href={`/api/evaluation/batch/${batchId}/split-pack/template?kind=rules`}
              className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-medium text-slate-600 hover:border-[#3B5B7E] hover:text-[#3B5B7E]"
            >
              <FileSpreadsheet className="h-3.5 w-3.5" />
              <span>下载分包规则模板</span>
            </a>
          </div>
        </div>

        <div className="grid grid-cols-1 gap-3 md:grid-cols-3">
          <div className="rounded-lg border border-slate-200 bg-slate-50/50 p-3 space-y-2">
            <div className="text-xs font-semibold text-slate-700">1. 物料表（1 份）</div>
            <div className="text-[11px] text-slate-500 truncate">
              {matFile
                ? `待上传：${matFile.name}`
                : state?.files.materials
                  ? `已上传：${state.files.materials.originalName}`
                  : '尚未选择物料表'}
            </div>
            <label className="inline-flex cursor-pointer items-center gap-1 rounded-md border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:border-[#3B5B7E] hover:text-[#3B5B7E]">
              <Upload className="h-3.5 w-3.5" />
              <span>选择物料表</span>
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) => setMatFile(e.target.files?.[0] || null)}
                className="hidden"
              />
            </label>
          </div>

          <div className="rounded-lg border border-slate-200 bg-slate-50/50 p-3 space-y-2">
            <div className="text-xs font-semibold text-slate-700">2. 分包规则表（包号 + 分配值）</div>
            <div className="text-[11px] text-slate-500 truncate">
              {ruleFile
                ? `待上传：${ruleFile.name}`
                : state?.files.rules
                  ? `已上传：${state.files.rules.originalName}`
                  : '尚未选择分包规则表'}
            </div>
            <label className="inline-flex cursor-pointer items-center gap-1 rounded-md border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:border-[#3B5B7E] hover:text-[#3B5B7E]">
              <Upload className="h-3.5 w-3.5" />
              <span>选择规则表</span>
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) => setRuleFile(e.target.files?.[0] || null)}
                className="hidden"
              />
            </label>
          </div>

          <div className="rounded-lg border border-slate-200 bg-slate-50/50 p-3 space-y-2">
            <div className="flex items-center justify-between gap-2">
              <span className="text-xs font-semibold text-slate-700">3. 分配方式与粒度</span>
              <div className="inline-flex rounded border border-slate-200 bg-white p-0.5">
                {(['比例', '金额'] as const).map((m) => (
                  <button
                    key={m}
                    type="button"
                    onClick={() => setMode(m)}
                    className={`rounded px-2 py-0.5 text-[11px] font-medium ${
                      mode === m ? 'bg-[#3B5B7E] text-white' : 'text-slate-600'
                    }`}
                  >
                    按{m}
                  </button>
                ))}
              </div>
            </div>
            <div className="grid grid-cols-2 gap-2">
              <div>
                <label className="block text-[10px] text-slate-500 mb-0.5">重量粒度(千克)</label>
                <input
                  value={weightStep}
                  onChange={(e) => setWeightStep(e.target.value)}
                  className="w-full rounded border border-slate-200 bg-white px-2 py-1 text-xs"
                />
              </div>
              <div>
                <label className="block text-[10px] text-slate-500 mb-0.5">长度粒度(米)</label>
                <input
                  value={lengthStep}
                  onChange={(e) => setLengthStep(e.target.value)}
                  className="w-full rounded border border-slate-200 bg-white px-2 py-1 text-xs"
                />
              </div>
            </div>
            <div className="flex items-center gap-2">
              <button
                type="button"
                disabled={running}
                onClick={handleRun}
                className="inline-flex flex-1 items-center justify-center gap-1.5 rounded-md bg-[#3B5B7E] px-3 py-1.5 text-xs font-medium text-white hover:bg-[#2E4866] disabled:opacity-50"
              >
                <Play className="h-3.5 w-3.5" />
                <span>{running ? '分包中...' : '生成分包方案'}</span>
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
            <div className="flex flex-wrap items-center gap-1.5">
              <button
                type="button"
                onClick={() => setViewTab('summary')}
                className={`rounded-md px-2.5 py-1 text-xs font-medium ${
                  viewTab === 'summary' ? 'bg-[#3B5B7E] text-white' : 'bg-slate-100 text-slate-600'
                }`}
              >
                分包汇总 ({state.result.packCount} 包)
              </button>
              <button
                type="button"
                onClick={() => setViewTab('detail')}
                className={`rounded-md px-2.5 py-1 text-xs font-medium ${
                  viewTab === 'detail' ? 'bg-[#3B5B7E] text-white' : 'bg-slate-100 text-slate-600'
                }`}
              >
                分包结果数据 ({state.result.detailRows.length} 条)
              </button>
            </div>
            <div className="flex flex-wrap items-center gap-2">
              {viewTab === 'detail' && (
                <div className="relative w-52">
                  <Search className="pointer-events-none absolute left-2.5 top-2 h-3.5 w-3.5 text-slate-400" />
                  <input
                    value={search}
                    onChange={(e) => setSearch(e.target.value)}
                    placeholder="搜索包号、物料编码或描述..."
                    className="w-full rounded-lg border border-slate-200 py-1.5 pl-8 pr-3 text-xs focus:border-[#3B5B7E] focus:outline-none"
                  />
                </div>
              )}
              <a
                href={`/api/evaluation/batch/${batchId}/split-pack/export`}
                className="inline-flex items-center gap-1 rounded-md bg-[#2F6F5E] px-3 py-1.5 text-xs font-medium text-white hover:bg-[#24594B]"
              >
                <Download className="h-3.5 w-3.5" />
                <span>导出分包结果数据 (.xlsx)</span>
              </a>
            </div>
          </div>

          {state.result.warnings.length > 0 && (
            <div className="rounded-lg border border-[#9A6B26]/25 bg-[#FEF9EC] px-3 py-2 text-[11px] text-[#9A6B26]">
              {state.result.warnings.join(' ')}
            </div>
          )}

          {viewTab === 'summary' ? (
            <div className="overflow-x-auto">
              <table className="w-full border-collapse text-left text-xs">
                <thead>
                  <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                    <th className="px-2.5 py-2 font-medium">包号</th>
                    <th className="px-2.5 py-2 font-medium">物料数量</th>
                    <th className="px-2.5 py-2 font-medium">总数量</th>
                    <th className="px-2.5 py-2 font-medium">总预算</th>
                    <th className="px-2.5 py-2 font-medium">目标金额</th>
                    <th className="px-2.5 py-2 font-medium">实际减目标</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {state.result.summaryRows.map((r, idx) => (
                    <tr
                      key={idx}
                      className={r.pack_no === '合计' ? 'bg-slate-50 font-semibold text-slate-800' : 'hover:bg-slate-50/70'}
                    >
                      <td className="px-2.5 py-2">{r.pack_no}</td>
                      <td className="px-2.5 py-2 tabular-nums">{r.item_count}</td>
                      <td className="px-2.5 py-2 tabular-nums">{r.total_qty}</td>
                      <td className="px-2.5 py-2 tabular-nums text-[#3B5B7E] font-medium">{r.total_budget}</td>
                      <td className="px-2.5 py-2 tabular-nums text-slate-600">{r.target_amount}</td>
                      <td className="px-2.5 py-2 tabular-nums text-slate-600">{r.diff_amount}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <>
              <div className="overflow-x-auto">
                <table className="w-full border-collapse text-left text-xs">
                  <thead>
                    <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                      <th className="px-2.5 py-2 font-medium whitespace-nowrap">序号</th>
                      <th className="px-2.5 py-2 font-medium whitespace-nowrap">包号</th>
                      <th className="px-2.5 py-2 font-medium whitespace-nowrap">物料编码</th>
                      <th className="px-2.5 py-2 font-medium">物料描述 / 扩展描述</th>
                      <th className="px-2.5 py-2 font-medium whitespace-nowrap">数量 / 单位</th>
                      <th className="px-2.5 py-2 font-medium whitespace-nowrap">预算单价</th>
                      <th className="px-2.5 py-2 font-medium whitespace-nowrap">预算总价</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {paginate(filteredDetails, detailPager.page, detailPager.pageSize).map((r) => (
                      <tr key={r.seq} className="hover:bg-slate-50/70">
                        <td className="px-2.5 py-2 tabular-nums text-slate-400">{r.seq}</td>
                        <td className="px-2.5 py-2 whitespace-nowrap font-medium text-slate-700">{r.pack_no}</td>
                        <td className="px-2.5 py-2 whitespace-nowrap font-mono text-slate-700">{r.material_code}</td>
                        <td className="px-2.5 py-2">
                          <div className="text-slate-800">{r.description}</div>
                          {r.extension && <div className="text-[11px] text-slate-400">{r.extension}</div>}
                        </td>
                        <td className="px-2.5 py-2 whitespace-nowrap tabular-nums text-slate-700">
                          {r.quantity} {r.unit}
                        </td>
                        <td className="px-2.5 py-2 whitespace-nowrap tabular-nums text-slate-600">{r.unit_price}</td>
                        <td className="px-2.5 py-2 whitespace-nowrap tabular-nums font-medium text-[#3B5B7E]">{r.total_price}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              {filteredDetails.length > 20 && (
                <div className="overflow-hidden rounded-lg border border-slate-200">
                  <PaginationBar
                    total={filteredDetails.length}
                    page={detailPager.page}
                    pageSize={detailPager.pageSize}
                    onPageChange={detailPager.setPage}
                    onPageSizeChange={detailPager.setPageSize}
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
