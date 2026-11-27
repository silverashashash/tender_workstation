import { useCallback, useEffect, useState } from 'react';
import {
  AlertTriangle,
  BookOpen,
  Calculator,
  Download,
  FileSpreadsheet,
  Play,
  Search,
  Trash2,
  Upload,
} from 'lucide-react';
import { api } from '../api/client';
import { paginate, PaginationBar, usePagination } from './PaginationBar';
import type { PriceScoreState } from '../types';

const FORMULA_DOCS = [
  {
    name: '区间平均价浮动法',
    steps: '有效报价 → 按数量去极值 → 均价T → 区间筛选 → 基准价B',
    base: '≤5家不去极值；6–10家去1高1低；11–20家去2高1低；21–30家去3高2低；>30家去4高3低（按条数剔除）。剩余均价为T，保留 [(1+w1)×T, (1+w2)×T]，B=保留均值×(1−c)（区间为空时用T）。',
    score: '价格分 = max(0, 100 − 100×k×|P−B|÷B)，P≥B 取 n1，P<B 取 n2。',
    params: 'w1 ∈ (-1,0)，w2 ∈ (0,1)，c ∈ (-0.1,0.1)，n1/n2 > 0，qz ∈ [0,100]（如 30 表示 30%）。',
  },
  {
    name: '算术平均值下浮法',
    steps: '有效报价 → 均价A → 下浮 → 基准价B',
    base: 'A = 全部有效报价算术平均值；B = A×(1−c)。',
    score: '价格分 = max(0, 100 − 100×k×|P−B|÷B)，P≥B 取 n1，P<B 取 n2。',
    params: 'c ∈ (-0.1,0.1)，n1/n2 > 0，qz ∈ [0,100]。',
  },
  {
    name: '区间平均下浮双边曲线算法',
    steps: '有效报价 → 均价A → 区间筛选 → 浮动 → 双边曲线评分',
    base: '在 [F1×A, F2×A] 内筛选，B = 保留均值×(1+c)（区间为空时用A）。',
    score: 'P≥B：价格分 = 100×(B/P)^n1；P<B：价格分 = 100×(P/B)^n2。',
    params: 'F1 ∈ (0,1)，F2 ∈ (1,2)，c ∈ (-0.1,0.1)，n1/n2 > 0，qz ∈ [0,100]。',
  },
  {
    name: '区间复合平均价法',
    steps: '有效报价 → 均价A → 区间筛选 → 复合基准价B',
    base: '在 [(1+b)×A, (1+c)×A] 内筛选，B = 保留均值；区间为空且最低价唯一时 B=(A+最低价)÷2（并列最低价报错待确认）。',
    score: '价格分 = max(0, 100 − 100×k×|P−B|÷B)，P≥B 取 n1，P<B 取 n2。',
    params: 'b ∈ (-1,0)，c ∈ (0,1)（注意本公式 c 为上界系数），n1/n2 > 0，qz ∈ [0,100]。',
  },
  {
    name: '算数平均值浮动法',
    steps: '有效报价 → 均价A → 剔除低报价 → 浮动 → 基准价B',
    base: '保留 P ≥ A×(1−m) 的报价，B = 保留均值×(1+c)。',
    score: '价格分 = max(0, 100 − 100×k×|P−B|÷B)，P≥B 取 F1，P<B 取 F2。',
    params: 'm ∈ [0,1)，c ∈ (-0.1,0.1)，F1/F2 > 0，qz ∈ [0,100]。',
  },
  {
    name: '区间复合平均价浮动法',
    steps: '有效报价 → 按家数分支 → 区间筛选 → 浮动 → 基准价B',
    base: '≤5家：B = A×(1+c)；>5家：保留 (1+a)×A < P < (1+b)×A，B = 保留均值×(1+c)（区间为空时用A）。',
    score: '价格分 = max(0, 100 − 100×k×|P−B|÷B)，P≥B 取 n1，P<B 取 n2。',
    params: 'a ∈ (-1,0)，b ∈ (0,1)，c ∈ (-0.1,0.1)，n1/n2 > 0，qz ∈ [0,100]。',
  },
];

export function PriceScorePanel({ batchId }: { batchId: string }) {
  const [state, setState] = useState<PriceScoreState | null>(null);
  const [quoteFiles, setQuoteFiles] = useState<File[]>([]);
  const [ruleFile, setRuleFile] = useState<File | null>(null);
  const [roundFilter, setRoundFilter] = useState('');
  const [selectedFormula, setSelectedFormula] = useState(FORMULA_DOCS[0].name);
  const [showDocs, setShowDocs] = useState(false);
  const [search, setSearch] = useState('');
  const [running, setRunning] = useState(false);
  const [error, setError] = useState('');

  const pager = usePagination(`${batchId}-ps-${search}`);

  const load = useCallback(async () => {
    try {
      const data = await api.getPriceScoreState(batchId);
      setState(data);
      setRoundFilter(data.config?.round || '');
    } catch {
      // ignore
    }
  }, [batchId]);

  useEffect(() => {
    load();
    setQuoteFiles([]);
    setRuleFile(null);
    setError('');
  }, [load]);

  async function handleRun() {
    setRunning(true);
    setError('');
    try {
      const res = await api.runPriceScore(batchId, {
        quotes: quoteFiles.length ? quoteFiles : undefined,
        rules: ruleFile || undefined,
        round: roundFilter,
      });
      setState(res);
      setQuoteFiles([]);
      setRuleFile(null);
    } catch (err) {
      setError((err as Error).message);
    } finally {
      setRunning(false);
    }
  }

  async function handleClear() {
    if (!window.confirm('确认清除当前批次已保存的价格分计算文件与结果吗？')) return;
    await api.clearPriceScore(batchId);
    setQuoteFiles([]);
    setRuleFile(null);
    setRoundFilter('');
    setError('');
    await load();
  }

  const rows = state?.result?.rows || [];
  const q = search.trim().toLowerCase();
  const filtered = rows.filter(
    (r) =>
      !q ||
      `${r.section} ${r.pack} ${r.supplier} ${r.method_name} ${r.round}`.toLowerCase().includes(q)
  );
  const activeDoc = FORMULA_DOCS.find((d) => d.name === selectedFormula) || FORMULA_DOCS[0];

  return (
    <div className="space-y-4">
      <div className="rounded-xl border border-slate-200 bg-white p-4 space-y-4">
        <div className="flex flex-wrap items-start justify-between gap-2">
          <div>
            <h3 className="text-sm font-semibold text-slate-800 flex items-center gap-1.5">
              <Calculator className="h-4 w-4 text-[#3B5B7E]" />
              <span>价格分计算（国网 6 种标准公式）</span>
            </h3>
            <p className="mt-0.5 text-xs text-slate-500">
              严格按工具箱高精度十进制算法复核基准价、价格分与加权价格分，支持总价（万元）与折扣率（%）口径。
            </p>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <button
              type="button"
              onClick={() => setShowDocs((v) => !v)}
              className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-slate-50 px-2.5 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-100"
            >
              <BookOpen className="h-3.5 w-3.5 text-[#3B5B7E]" />
              <span>{showDocs ? '收起公式说明' : '查看 6 种公式说明'}</span>
            </button>
            <a
              href={`/api/evaluation/batch/${batchId}/price-score/template?kind=quotes`}
              className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-medium text-slate-600 hover:border-[#3B5B7E] hover:text-[#3B5B7E]"
            >
              <FileSpreadsheet className="h-3.5 w-3.5" />
              <span>下载报价模板</span>
            </a>
            <a
              href={`/api/evaluation/batch/${batchId}/price-score/template?kind=rules&method=${encodeURIComponent(selectedFormula)}`}
              className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-medium text-slate-600 hover:border-[#3B5B7E] hover:text-[#3B5B7E]"
            >
              <FileSpreadsheet className="h-3.5 w-3.5" />
              <span>下载参数模板</span>
            </a>
          </div>
        </div>

        {showDocs && (
          <div className="rounded-lg border border-[#3B5B7E]/20 bg-[#EBF2F7]/40 p-3.5 space-y-2.5 text-xs">
            <div className="flex flex-wrap items-center gap-2">
              <span className="font-semibold text-slate-700">公式选择（同步参数模板下载）：</span>
              <select
                value={selectedFormula}
                onChange={(e) => setSelectedFormula(e.target.value)}
                className="rounded-md border border-slate-200 bg-white px-2.5 py-1 text-xs"
              >
                {FORMULA_DOCS.map((d, idx) => (
                  <option key={d.name} value={d.name}>
                    {idx + 1}. {d.name}
                  </option>
                ))}
              </select>
            </div>
            <div className="text-[#3B5B7E] font-medium">{activeDoc.steps}</div>
            <div className="space-y-1 text-slate-600 leading-relaxed">
              <div><strong>基准价计算：</strong>{activeDoc.base}</div>
              <div><strong>价格分计算：</strong>{activeDoc.score} 加权价格分 = 价格分 ÷ 100 × qz（均保留两位小数）。</div>
              <div><strong>参数填写：</strong>{activeDoc.params}</div>
            </div>
          </div>
        )}

        <div className="grid grid-cols-1 gap-3 md:grid-cols-3">
          <div className="rounded-lg border border-slate-200 bg-slate-50/50 p-3 space-y-2">
            <div className="text-xs font-semibold text-slate-700">1. 报价文件（支持多选 .xlsx/.xls）</div>
            <div className="text-[11px] text-slate-500 truncate">
              {quoteFiles.length > 0
                ? `待上传 ${quoteFiles.length} 份：${quoteFiles.map((f) => f.name).join('、')}`
                : state?.files.quotes.length
                  ? `已上传 ${state.files.quotes.length} 份：${state.files.quotes.map((f) => f.originalName).join('、')}`
                  : '尚未选择报价文件'}
            </div>
            <label className="inline-flex cursor-pointer items-center gap-1 rounded-md border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:border-[#3B5B7E] hover:text-[#3B5B7E]">
              <Upload className="h-3.5 w-3.5" />
              <span>选择报价文件</span>
              <input
                type="file"
                multiple
                accept=".xlsx,.xls"
                onChange={(e) => setQuoteFiles(Array.from(e.target.files || []))}
                className="hidden"
              />
            </label>
          </div>

          <div className="rounded-lg border border-slate-200 bg-slate-50/50 p-3 space-y-2">
            <div className="text-xs font-semibold text-slate-700">2. 价格公式参数表（1 份）</div>
            <div className="text-[11px] text-slate-500 truncate">
              {ruleFile
                ? `待上传：${ruleFile.name}`
                : state?.files.rules
                  ? `已上传：${state.files.rules.originalName}`
                  : '尚未选择公式参数表'}
            </div>
            <label className="inline-flex cursor-pointer items-center gap-1 rounded-md border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:border-[#3B5B7E] hover:text-[#3B5B7E]">
              <Upload className="h-3.5 w-3.5" />
              <span>选择参数表</span>
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) => setRuleFile(e.target.files?.[0] || null)}
                className="hidden"
              />
            </label>
          </div>

          <div className="rounded-lg border border-slate-200 bg-slate-50/50 p-3 space-y-2">
            <div className="text-xs font-semibold text-slate-700">3. 轮次筛选（可选）</div>
            <input
              value={roundFilter}
              onChange={(e) => setRoundFilter(e.target.value)}
              placeholder="留空计算单一轮次；多轮时填如 第1轮"
              className="w-full rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
            />
            <div className="flex items-center gap-2 pt-0.5">
              <button
                type="button"
                disabled={running}
                onClick={handleRun}
                className="inline-flex flex-1 items-center justify-center gap-1.5 rounded-md bg-[#3B5B7E] px-3 py-1.5 text-xs font-medium text-white hover:bg-[#2E4866] disabled:opacity-50"
              >
                <Play className="h-3.5 w-3.5" />
                <span>{running ? '计算中...' : '开始计算'}</span>
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
            <div className="flex flex-wrap items-center gap-2 text-xs text-slate-600">
              <span className="font-semibold text-slate-800">计算结果明细</span>
              <span className="rounded bg-[#EBF2F7] px-2 py-0.5 font-medium text-[#3B5B7E] tabular-nums">
                {state.result.packageCount} 个标包 · {state.result.totalRows} 条报价
              </span>
            </div>
            <div className="flex flex-wrap items-center gap-2">
              <div className="relative w-56">
                <Search className="pointer-events-none absolute left-2.5 top-2 h-3.5 w-3.5 text-slate-400" />
                <input
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  placeholder="搜索分标、包号或投标人..."
                  className="w-full rounded-lg border border-slate-200 py-1.5 pl-8 pr-3 text-xs focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>
              <a
                href={`/api/evaluation/batch/${batchId}/price-score/export`}
                className="inline-flex items-center gap-1 rounded-md bg-[#2F6F5E] px-3 py-1.5 text-xs font-medium text-white hover:bg-[#24594B]"
              >
                <Download className="h-3.5 w-3.5" />
                <span>导出价格分计算 (.xlsx)</span>
              </a>
            </div>
          </div>

          {state.result.warnings.length > 0 && (
            <div className="rounded-lg border border-[#9A6B26]/25 bg-[#FEF9EC] px-3 py-2 text-[11px] text-[#9A6B26]">
              {state.result.warnings.join(' ')}
            </div>
          )}

          <div className="overflow-x-auto">
            <table className="w-full border-collapse text-left text-xs">
              <thead>
                <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                  <th className="px-2.5 py-2 font-medium whitespace-nowrap">分标 / 包号</th>
                  <th className="px-2.5 py-2 font-medium whitespace-nowrap">轮次</th>
                  <th className="px-2.5 py-2 font-medium">投标人名称</th>
                  <th className="px-2.5 py-2 font-medium whitespace-nowrap">口径 / 报价</th>
                  <th className="px-2.5 py-2 font-medium whitespace-nowrap">价格公式</th>
                  <th className="px-2.5 py-2 font-medium whitespace-nowrap">基准值</th>
                  <th className="px-2.5 py-2 font-medium whitespace-nowrap">价格分</th>
                  <th className="px-2.5 py-2 font-medium whitespace-nowrap">加权价格分</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {paginate(filtered, pager.page, pager.pageSize).map((r, idx) => (
                  <tr key={idx} className="hover:bg-slate-50/70">
                    <td className="px-2.5 py-2 whitespace-nowrap">
                      <div className="font-medium text-slate-800">{r.section}</div>
                      <div className="text-[11px] text-slate-400">{r.pack}</div>
                    </td>
                    <td className="px-2.5 py-2 whitespace-nowrap text-slate-600">{r.round}</td>
                    <td className="px-2.5 py-2 font-medium text-slate-800">{r.supplier}</td>
                    <td className="px-2.5 py-2 whitespace-nowrap tabular-nums">
                      <span className="font-medium text-slate-700">{r.quote_value}</span>
                      <span className="ml-1 text-[10px] text-slate-400">{r.quote_basis}</span>
                    </td>
                    <td className="px-2.5 py-2 whitespace-nowrap text-slate-600">{r.method_name}</td>
                    <td className="px-2.5 py-2 whitespace-nowrap tabular-nums text-slate-600">{r.base_value}</td>
                    <td className="px-2.5 py-2 whitespace-nowrap tabular-nums font-semibold text-[#3B5B7E]">{r.price_score}</td>
                    <td className="px-2.5 py-2 whitespace-nowrap tabular-nums font-semibold text-[#2F6F5E]">{r.weighted_score}</td>
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
