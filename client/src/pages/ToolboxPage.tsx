import React, { useMemo, useState } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import {
  ArrowLeft,
  FileCheck2,
  FileSearch,
  FileSpreadsheet,
  FileText,
  Layers,
  Search,
  Shuffle,
  Sparkles,
  SplitSquareVertical,
  TableProperties,
  Wrench,
} from 'lucide-react';
import { ChecklistCompliancePanel } from '../components/ChecklistCompliancePanel';
import { PerformanceMatchPanel } from '../components/PerformanceMatchPanel';
import { RandomScorePanel } from '../components/RandomScorePanel';
import { RowExpanderPanel } from '../components/RowExpanderPanel';
import { ScoreTemplateMergePanel } from '../components/ScoreTemplateMergePanel';
import { SheetMergerPanel } from '../components/SheetMergerPanel';
import { WordBatchPanel } from '../components/WordBatchPanel';

export type ToolboxCategory = '表格处理' | '文档处理' | '通用工具';

interface ToolboxItem {
  key: string;
  title: string;
  category: ToolboxCategory;
  status: 'ready' | 'coming_soon';
  summary: string;
  tags: string[];
  icon: React.ComponentType<{ className?: string }>;
}

const TOOLBOX_ITEMS: ToolboxItem[] = [
  {
    key: 'sheet-merger',
    title: '多工作表合并 (Sheet Merger)',
    category: '表格处理',
    status: 'ready',
    summary:
      '将同一 Excel 工作簿中多个表头一致的工作表（Sheet）纵向合并为一张总表，可指定表头行与数据起始行，支持合并单元格展开、按列名对齐及来源标记。',
    tags: ['Excel 合并', '多 Sheet 汇总', '指定起始行'],
    icon: Layers,
  },
  {
    key: 'row-expander',
    title: '单行数据展开与合并',
    category: '表格处理',
    status: 'ready',
    summary:
      '按分隔符将单行多值展开为多行、两级拆分「分标:包号」、逆向一级/两级合并，并支持供应商联系人展开与冲突核查（输出 4 表核查报告）。',
    tags: ['Excel 展开', '分标包号拆分', '逆向合并', '联系人核查'],
    icon: SplitSquareVertical,
  },
  {
    key: 'word-batch',
    title: 'Word 文字与数字批量处理',
    category: '文档处理',
    status: 'ready',
    summary:
      '批量处理 Word 评审报告中的数字与文字：支持模板数字按倍数修改、整段删除、多组查找替换（支持正则）及按分标-分包提取评标表格为 Excel。',
    tags: ['Word 批量处理', '数字倍增', '查找替换', '评标表提取'],
    icon: FileText,
  },
  {
    key: 'score-merge',
    title: '专家打分表按模板合并',
    category: '表格处理',
    status: 'ready',
    summary:
      '上传多位专家分别填写的线下打分表与原始模板，自动忽略未调整/清空行、按标段+包号+应答人对齐汇总，冲突标黄并支持按《优秀名单》自动高亮。',
    tags: ['Excel 合并', '专家打分', '冲突标黄', '优秀名单高亮'],
    icon: TableProperties,
  },
  {
    key: 'perf-match',
    title: '履约分匹配填入',
    category: '表格处理',
    status: 'ready',
    summary:
      '按「分标名称 + 包号 + 投标人」自动把履约分批量填入多份详评打分模板的「绩效评价」列，保留原模板样式与公式，未匹配默认填 4 分并输出报告。',
    tags: ['履约分', '详评模板填充', '批量匹配'],
    icon: FileCheck2,
  },
  {
    key: 'checklist-compliance',
    title: '清单合规性检验',
    category: '表格处理',
    status: 'ready',
    summary:
      '递归解压批次或技术规范书压缩包（修复中文乱码），自动检索「工作量清单/工程量清单」并校验文件大小、损坏与空表，生成着色检验报告。',
    tags: ['压缩包递归解压', '清单检验', 'GBK 乱码修复'],
    icon: FileSearch,
  },
  {
    key: 'random-score',
    title: '专家随机打分',
    category: '表格处理',
    status: 'ready',
    summary:
      '根据一份源专家打分表，自动锁定黄色客观分列，在最高/最低/级差范围内为其余 n-1 位专家随机生成合规主观分（总分差 ≤7、同分 ≤⌊n/2⌋）。',
    tags: ['专家打分', '客观分锁定', '合规校验'],
    icon: Shuffle,
  },
];

const CATEGORIES: Array<'全部' | ToolboxCategory> = ['全部', '表格处理', '文档处理', '通用工具'];

export function ToolboxPage() {
  const { toolKey } = useParams<{ toolKey?: string }>();
  const navigate = useNavigate();
  const [category, setCategory] = useState<'全部' | ToolboxCategory>('全部');
  const [search, setSearch] = useState('');

  const activeTool = useMemo(
    () => (toolKey ? TOOLBOX_ITEMS.find((t) => t.key === toolKey) || null : null),
    [toolKey]
  );

  const filteredTools = useMemo(() => {
    const q = search.trim().toLowerCase();
    return TOOLBOX_ITEMS.filter((item) => {
      if (category !== '全部' && item.category !== category) return false;
      if (!q) return true;
      const hay = [item.title, item.category, item.summary, ...item.tags].join(' ').toLowerCase();
      return hay.includes(q);
    });
  }, [category, search]);

  // Single-tool focused view
  if (activeTool) {
    const Icon = activeTool.icon;
    return (
      <div className="space-y-5">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-3 min-w-0">
            <button
              type="button"
              onClick={() => navigate('/toolbox')}
              className="inline-flex shrink-0 items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-600 hover:border-[#3B5B7E] hover:text-[#3B5B7E] transition-colors"
            >
              <ArrowLeft className="h-3.5 w-3.5" />
              <span>返回实用工具箱</span>
            </button>
            <div className="min-w-0">
              <div className="flex items-center gap-2">
                <Icon className="h-4 w-4 text-[#3B5B7E] shrink-0" />
                <h1 className="text-base font-semibold text-slate-800 truncate">{activeTool.title}</h1>
                <span className="rounded bg-[#EBF2F7] px-2 py-0.5 text-[11px] font-medium text-[#3B5B7E]">
                  {activeTool.category}
                </span>
              </div>
            </div>
          </div>
        </div>

        {activeTool.key === 'sheet-merger' ? (
          <SheetMergerPanel />
        ) : activeTool.key === 'row-expander' ? (
          <RowExpanderPanel />
        ) : activeTool.key === 'word-batch' ? (
          <WordBatchPanel />
        ) : activeTool.key === 'score-merge' ? (
          <ScoreTemplateMergePanel />
        ) : activeTool.key === 'perf-match' ? (
          <PerformanceMatchPanel />
        ) : activeTool.key === 'checklist-compliance' ? (
          <ChecklistCompliancePanel />
        ) : activeTool.key === 'random-score' ? (
          <RandomScorePanel />
        ) : (
          <div className="rounded-xl border border-slate-200 bg-white p-8 text-center space-y-3">
            <div className="mx-auto flex h-12 w-12 items-center justify-center rounded-xl bg-[#EBF2F7] text-[#3B5B7E]">
              <FileSpreadsheet className="h-6 w-6" />
            </div>
            <h2 className="text-sm font-semibold text-slate-800">{activeTool.title} · 工具入口已就绪</h2>
            <p className="mx-auto max-w-lg text-xs text-slate-500 leading-relaxed">
              {activeTool.summary}
            </p>
            <div className="pt-1">
              <span className="inline-flex items-center gap-1.5 rounded-full border border-[#9A6B26]/30 bg-[#FEF9EC] px-3 py-1 text-[11px] font-medium text-[#9A6B26]">
                <Sparkles className="h-3 w-3" />
                <span>待提供样表与合并规则后即可接入具体处理逻辑</span>
              </span>
            </div>
          </div>
        )}
      </div>
    );
  }

  // Toolbox Card Wall Home
  return (
    <div className="space-y-5">
      {/* Header */}
      <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <div className="flex flex-wrap items-center gap-2">
            <Wrench className="h-5 w-5 text-[#3B5B7E]" />
            <h1 className="text-lg font-semibold text-slate-800">实用工具箱</h1>
            <span className="rounded-md bg-slate-100 px-2 py-0.5 text-xs font-medium text-slate-500">
              跨阶段独立工具 · 即开即用，不绑定当前批次
            </span>
          </div>
          <p className="mt-1 text-xs text-slate-500">
            收纳招投标全流程中常用的零散表格、文档与批处理小工具；不依赖批次上下文，上传文件处理完即可直接导出。
          </p>
        </div>
      </div>

      {/* Category Filter + Search Bar */}
      <div className="flex flex-col gap-2.5 sm:flex-row sm:items-center sm:justify-between rounded-xl border border-slate-200 bg-white p-3.5">
        <div className="flex flex-wrap items-center gap-1.5">
          {CATEGORIES.map((cat) => {
            const count =
              cat === '全部'
                ? TOOLBOX_ITEMS.length
                : TOOLBOX_ITEMS.filter((t) => t.category === cat).length;
            const active = category === cat;
            return (
              <button
                key={cat}
                type="button"
                onClick={() => setCategory(cat)}
                className={`rounded-lg px-3 py-1.5 text-xs font-medium transition-colors tabular-nums ${
                  active
                    ? 'bg-[#3B5B7E] text-white'
                    : 'bg-slate-100 text-slate-600 hover:bg-slate-200/70'
                }`}
              >
                {cat} ({count})
              </button>
            );
          })}
        </div>
        <div className="relative w-full sm:w-64">
          <Search className="pointer-events-none absolute left-2.5 top-2 h-3.5 w-3.5 text-slate-400" />
          <input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="搜索工具名称或用途关键词..."
            className="w-full rounded-lg border border-slate-200 bg-slate-50/70 py-1.5 pl-8 pr-3 text-xs focus:border-[#3B5B7E] focus:bg-white focus:outline-none"
          />
        </div>
      </div>

      {/* Tool Cards Grid */}
      {filteredTools.length === 0 ? (
        <div className="rounded-xl border border-slate-200 bg-white py-12 text-center text-xs text-slate-400">
          该分类或关键词下暂无工具
        </div>
      ) : (
        <div className="grid grid-cols-1 gap-4 md:grid-cols-2 lg:grid-cols-3">
          {filteredTools.map((tool) => {
            const Icon = tool.icon;
            return (
              <button
                key={tool.key}
                type="button"
                onClick={() => navigate(`/toolbox/${tool.key}`)}
                className="group flex flex-col justify-between rounded-xl border border-slate-200 bg-white p-4 text-left transition-all hover:border-[#3B5B7E]/50 hover:shadow-sm"
              >
                <div className="space-y-2.5">
                  <div className="flex items-start justify-between gap-2">
                    <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-[#EBF2F7] text-[#3B5B7E] group-hover:bg-[#3B5B7E] group-hover:text-white transition-colors">
                      <Icon className="h-4 w-4" />
                    </div>
                    <div className="flex items-center gap-1.5">
                      <span className="rounded bg-slate-100 px-2 py-0.5 text-[10px] font-medium text-slate-500">
                        {tool.category}
                      </span>
                      {tool.status === 'coming_soon' ? (
                        <span className="rounded bg-[#FEF9EC] border border-[#9A6B26]/25 px-2 py-0.5 text-[10px] font-medium text-[#9A6B26]">
                          待接入逻辑
                        </span>
                      ) : (
                        <span className="rounded bg-[#ECFDF5] px-2 py-0.5 text-[10px] font-medium text-[#2F6F5E]">
                          可用
                        </span>
                      )}
                    </div>
                  </div>
                  <div>
                    <h3 className="text-sm font-semibold text-slate-800 group-hover:text-[#3B5B7E] transition-colors">
                      {tool.title}
                    </h3>
                    <p className="mt-1 text-xs text-slate-500 leading-relaxed">
                      {tool.summary}
                    </p>
                  </div>
                </div>

                <div className="mt-4 flex flex-wrap items-center justify-between gap-2 border-t border-slate-100 pt-2.5">
                  <div className="flex flex-wrap gap-1">
                    {tool.tags.map((tg) => (
                      <span
                        key={tg}
                        className="rounded bg-slate-50 px-1.5 py-0.5 text-[10px] text-slate-500"
                      >
                        {tg}
                      </span>
                    ))}
                  </div>
                  <span className="text-[11px] font-medium text-[#3B5B7E] group-hover:translate-x-0.5 transition-transform">
                    打开 →
                  </span>
                </div>
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}
