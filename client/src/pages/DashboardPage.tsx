import React, { useCallback, useEffect, useMemo, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  CalendarClock,
  FolderKanban,
  AlertTriangle,
  CheckCircle2,
  ArrowRight,
  FileSpreadsheet,
  ClipboardCheck,
  Award,
  FolderArchive,
  Search,
  LayoutList,
  LayoutGrid,
  X,
  ChevronDown,
  ChevronUp,
  Coins,
  Trash2,
  RotateCcw,
} from 'lucide-react';
import { api } from '../api/client';
import { BatchDeleteModal } from '../components/BatchDeleteModal';
import { paginate, PaginationBar, usePagination } from '../components/PaginationBar';
import {
  StatusDonutChart,
  BidCalendarHeatmap,
  BatchTimelineChart,
  ProgressGapChart,
  STATUS_META,
  STATUS_ORDER,
  getCountdownLabel,
  parseDateOnly,
  formatDateKey,
} from '../components/DashboardCharts';
import { useBatch } from '../context/BatchContext';
import type { Batch } from '../types';

type QuickFilter = 'all' | 'urgent' | 'missing_archive';
type ViewMode = 'table' | 'cards';

export function DashboardPage() {
  const { batches, setCurrentBatchId, refreshBatches, authEnabled } = useBatch();
  const navigate = useNavigate();

  // Interactive filter & view state
  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState<string>('all');
  const [quickFilter, setQuickFilter] = useState<QuickFilter>('all');
  const [dateFilter, setDateFilter] = useState<string | null>(null);
  const [hideCompleted, setHideCompleted] = useState(false);
  const [viewMode, setViewMode] = useState<ViewMode>('table');
  const [expandedBatchId, setExpandedBatchId] = useState<string | null>(null);

  // Recycle Bin (soft-delete) & Hard Delete modal state
  const [trashBatches, setTrashBatches] = useState<Batch[]>([]);
  const [showTrash, setShowTrash] = useState(false);
  const [pendingHardDelete, setPendingHardDelete] = useState<Batch | null>(null);
  const [recentTrashed, setRecentTrashed] = useState<{ id: string; name: string } | null>(null);
  const [busyBatchId, setBusyBatchId] = useState<string | null>(null);

  const loadTrash = useCallback(async () => {
    try {
      const list = await api.listTrashBatches();
      setTrashBatches(list);
    } catch {
      // ignore if offline
    }
  }, []);

  useEffect(() => {
    loadTrash();
  }, [loadTrash]);

  async function handleMoveToTrash(b: Batch) {
    if (!window.confirm(`确认将「${b.name}」移入回收站吗？\n\n移入后将从总览和左侧列表隐藏，不会立即删除数据，可随时在回收站一键还原。`)) {
      return;
    }
    setBusyBatchId(b.id);
    try {
      await api.trashBatch(b.id);
      if (expandedBatchId === b.id) setExpandedBatchId(null);
      setRecentTrashed({ id: b.id, name: b.name });
      await Promise.all([refreshBatches(), loadTrash()]);
    } catch (err) {
      alert((err as Error).message);
    } finally {
      setBusyBatchId(null);
    }
  }

  async function handleRestoreBatch(id: string) {
    setBusyBatchId(id);
    try {
      await api.restoreBatch(id);
      if (recentTrashed?.id === id) setRecentTrashed(null);
      await Promise.all([refreshBatches(), loadTrash()]);
    } catch (err) {
      alert((err as Error).message);
    } finally {
      setBusyBatchId(null);
    }
  }

  async function handleConfirmHardDelete(payload: { confirmText: string; password?: string }) {
    if (!pendingHardDelete) return;
    await api.deleteBatch(pendingHardDelete.id, payload);
    setPendingHardDelete(null);
    await Promise.all([refreshBatches(), loadTrash()]);
  }

  // KPI metrics
  const totalBatches = batches.length;
  const activeBatches = batches.filter((b) => b.status !== '已完成').length;
  const upcomingBatches = batches.filter((b) => getCountdownLabel(b.bid_opening_time)?.urgent);
  const upcomingCount = upcomingBatches.length;
  const totalProjects = batches.reduce((acc, b) => acc + (b.stats?.projectCount || 0), 0);
  const missingReqBatches = batches.filter((b) => (b.stats?.archiveReqMissing || 0) > 0).length;
  const totalMissingItems = batches.reduce((acc, b) => acc + (b.stats?.archiveReqMissing || 0), 0);
  const totalAwardOutputs = batches.reduce((acc, b) => acc + (b.stats?.awardOutputCount || 0), 0);
  const totalBudgetWan = batches.reduce((acc, b) => acc + (b.stats?.totalBudgetWan || 0), 0);

  // Filtered batch list for the bottom ledger panel
  const filteredBatches = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    return batches.filter((b) => {
      if (hideCompleted && b.status === '已完成') return false;
      if (statusFilter !== 'all' && b.status !== statusFilter) return false;
      if (quickFilter === 'urgent' && !getCountdownLabel(b.bid_opening_time)?.urgent) return false;
      if (quickFilter === 'missing_archive' && (b.stats?.archiveReqMissing || 0) <= 0) return false;
      if (dateFilter) {
        const openKey = parseDateOnly(b.bid_opening_time) ? formatDateKey(parseDateOnly(b.bid_opening_time)!) : '';
        const annKey = parseDateOnly(b.announcement_time) ? formatDateKey(parseDateOnly(b.announcement_time)!) : '';
        const creKey = parseDateOnly(b.created_at) ? formatDateKey(parseDateOnly(b.created_at)!) : '';
        if (openKey !== dateFilter && annKey !== dateFilter && creKey !== dateFilter) return false;
      }
      if (q) {
        const hay = `${b.name} ${b.code} ${b.project_unit} ${b.status}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
  }, [batches, searchQuery, statusFilter, quickFilter, dateFilter, hideCompleted]);

  const batchPager = usePagination(
    `${filteredBatches.length}-${statusFilter}-${quickFilter}-${dateFilter || ''}-${hideCompleted}-${searchQuery}`
  );

  const hasActiveFilter =
    statusFilter !== 'all' ||
    quickFilter !== 'all' ||
    Boolean(dateFilter) ||
    hideCompleted ||
    searchQuery.trim() !== '';

  function clearAllFilters() {
    setSearchQuery('');
    setStatusFilter('all');
    setQuickFilter('all');
    setDateFilter(null);
    setHideCompleted(false);
  }

  function goToModule(batchId: string, path: string) {
    setCurrentBatchId(batchId);
    navigate(path);
  }

  return (
    <div className="space-y-5">
      {/* Page Header */}
      <div className="flex flex-wrap items-end justify-between gap-2">
        <div>
          <h1 className="text-lg font-semibold text-slate-800">工作台总览</h1>
          <p className="mt-0.5 text-xs text-slate-500">
            全局掌控各招标批次阶段分布、排期热力、时间节点与归档缺口，点击任意图表或指标可联动筛选下方台账。
          </p>
        </div>
      </div>

      {/* Undo Banner after moving a batch to Recycle Bin */}
      {recentTrashed && (
        <div className="flex flex-wrap items-center justify-between gap-2 rounded-xl border border-[#9A6B26]/30 bg-[#FEF9EC] px-4 py-2.5 text-xs text-[#9A6B26]">
          <span>
            已将批次<strong>「{recentTrashed.name}」</strong>移入回收站（数据已保留，未永久删除）。
          </span>
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => handleRestoreBatch(recentTrashed.id)}
              className="inline-flex items-center gap-1 rounded-md bg-white border border-[#9A6B26]/30 px-2.5 py-1 text-[11px] font-medium text-[#9A6B26] hover:bg-[#FEF3C7]"
            >
              <RotateCcw className="h-3 w-3" />
              <span>撤销还原</span>
            </button>
            <button
              type="button"
              onClick={() => setRecentTrashed(null)}
              className="text-slate-400 hover:text-slate-600"
            >
              <X className="h-3.5 w-3.5" />
            </button>
          </div>
        </div>
      )}

      {/* Zone 1: Top KPI Cards (Clickable quick filters) */}
      <div className="grid grid-cols-2 gap-3 lg:grid-cols-4">
        <button
          type="button"
          onClick={() => {
            setShowTrash(false);
            setQuickFilter('all');
            setHideCompleted((v) => !v);
          }}
          className={`rounded-xl border p-4 text-left transition-all ${
            hideCompleted && quickFilter === 'all' && !showTrash
              ? 'border-[#3B5B7E] bg-[#EBF2F7]/40 ring-1 ring-[#3B5B7E]/30'
              : 'border-slate-200 bg-white hover:border-slate-300'
          }`}
        >
          <div className="flex items-center justify-between text-xs text-slate-500">
            <span>在办招标批次</span>
            <FolderKanban className="h-4 w-4 text-[#3B5B7E]" />
          </div>
          <div className="mt-2 flex items-baseline gap-2">
            <span className="text-2xl font-semibold text-slate-800 tabular-nums">{activeBatches}</span>
            <span className="text-xs text-slate-400 tabular-nums">
              共 {totalBatches} 批 · {totalProjects} 标段
            </span>
          </div>
        </button>

        <button
          type="button"
          onClick={() => {
            setShowTrash(false);
            setQuickFilter((v) => (v === 'urgent' ? 'all' : 'urgent'));
          }}
          className={`rounded-xl border p-4 text-left transition-all ${
            quickFilter === 'urgent' && !showTrash
              ? 'border-[#9A6B26] bg-[#FEF9EC]/60 ring-1 ring-[#9A6B26]/30'
              : 'border-slate-200 bg-white hover:border-slate-300'
          }`}
        >
          <div className="flex items-center justify-between text-xs text-slate-500">
            <span>7日内即将开标</span>
            <CalendarClock className="h-4 w-4 text-[#9A6B26]" />
          </div>
          <div className="mt-2 flex items-baseline gap-2">
            <span className="text-2xl font-semibold text-slate-800 tabular-nums">{upcomingCount}</span>
            <span className="text-xs text-slate-400">
              {upcomingCount > 0 ? '点击聚焦临近批次' : '近期排期平稳'}
            </span>
          </div>
        </button>

        <button
          type="button"
          onClick={() => {
            setShowTrash(false);
            setQuickFilter((v) => (v === 'missing_archive' ? 'all' : 'missing_archive'));
          }}
          className={`rounded-xl border p-4 text-left transition-all ${
            quickFilter === 'missing_archive' && !showTrash
              ? 'border-[#A94442] bg-[#FEF2F2]/60 ring-1 ring-[#A94442]/30'
              : 'border-slate-200 bg-white hover:border-slate-300'
          }`}
        >
          <div className="flex items-center justify-between text-xs text-slate-500">
            <span>归档必须项待补</span>
            <AlertTriangle className="h-4 w-4 text-[#A94442]" />
          </div>
          <div className="mt-2 flex items-baseline gap-2">
            <span className="text-2xl font-semibold text-[#A94442] tabular-nums">{missingReqBatches}</span>
            <span className="text-xs text-slate-400 tabular-nums">
              个批次 · 累计缺 {totalMissingItems} 项
            </span>
          </div>
        </button>

        <div className="rounded-xl border border-slate-200 bg-white p-4">
          <div className="flex items-center justify-between text-xs text-slate-500">
            <span>{totalBudgetWan > 0 ? '管理预算规模' : '累计定标交付'}</span>
            {totalBudgetWan > 0 ? (
              <Coins className="h-4 w-4 text-[#2F6F5E]" />
            ) : (
              <CheckCircle2 className="h-4 w-4 text-[#2F6F5E]" />
            )}
          </div>
          <div className="mt-2 flex items-baseline gap-2">
            {totalBudgetWan > 0 ? (
              <>
                <span className="text-2xl font-semibold text-[#2F6F5E] tabular-nums">
                  {totalBudgetWan >= 10000
                    ? `${(totalBudgetWan / 10000).toFixed(2)}亿`
                    : `${Math.round(totalBudgetWan).toLocaleString()}万`}
                </span>
                <span className="text-xs text-slate-400 tabular-nums">
                  已交付 {totalAwardOutputs} 份成果
                </span>
              </>
            ) : (
              <>
                <span className="text-2xl font-semibold text-[#2F6F5E] tabular-nums">
                  {totalAwardOutputs}
                </span>
                <span className="text-xs text-slate-400">份定标交付文件</span>
              </>
            )}
          </div>
        </div>
      </div>

      {/* Zone 2: Visual Analytics Bento Grid (4 Interactive Charts) */}
      <div className="grid grid-cols-1 gap-4 lg:grid-cols-12">
        <div className="lg:col-span-5">
          <StatusDonutChart
            batches={batches}
            selectedStatus={statusFilter}
            onSelectStatus={(s) => {
              setShowTrash(false);
              setStatusFilter(s);
            }}
          />
        </div>
        <div className="lg:col-span-7">
          <BidCalendarHeatmap
            batches={batches}
            selectedDateKey={dateFilter}
            onSelectDateKey={(d) => {
              setShowTrash(false);
              setDateFilter(d);
            }}
            onOpenBatch={goToModule}
          />
        </div>
        <div className="lg:col-span-7">
          <BatchTimelineChart batches={batches} onOpenBatch={goToModule} />
        </div>
        <div className="lg:col-span-5">
          <ProgressGapChart batches={batches} onOpenBatch={goToModule} />
        </div>
      </div>

      {/* Zone 3: Contained Batch Ledger Panel (+ Recycle Bin View) */}
      <div className="overflow-hidden rounded-xl border border-slate-200 bg-white">
        {/* Ledger Header & Filter Bar */}
        <div className="flex flex-col gap-3 border-b border-slate-200 bg-slate-50/70 px-4 py-3 lg:flex-row lg:items-center lg:justify-between">
          <div className="flex flex-wrap items-center gap-2">
            <h2 className="text-sm font-semibold text-slate-800">
              {showTrash ? '批次回收站' : '招标批次台账'}
            </h2>
            <span className="rounded-full bg-white border border-slate-200 px-2 py-0.5 text-[11px] font-medium text-slate-600 tabular-nums">
              {showTrash ? `${trashBatches.length} 项待清理` : `${filteredBatches.length} / ${totalBatches}`}
            </span>

            {/* Active Filter Pills (hidden in trash view) */}
            {!showTrash && statusFilter !== 'all' && (
              <button
                type="button"
                onClick={() => setStatusFilter('all')}
                className="inline-flex items-center gap-1 rounded-full bg-[#EBF2F7] px-2 py-0.5 text-[11px] font-medium text-[#3B5B7E] hover:bg-[#DCE8F2]"
              >
                <span>状态：{statusFilter}</span>
                <X className="h-3 w-3" />
              </button>
            )}
            {!showTrash && quickFilter === 'urgent' && (
              <button
                type="button"
                onClick={() => setQuickFilter('all')}
                className="inline-flex items-center gap-1 rounded-full bg-[#FEF9EC] px-2 py-0.5 text-[11px] font-medium text-[#9A6B26]"
              >
                <span>7日内开标</span>
                <X className="h-3 w-3" />
              </button>
            )}
            {!showTrash && quickFilter === 'missing_archive' && (
              <button
                type="button"
                onClick={() => setQuickFilter('all')}
                className="inline-flex items-center gap-1 rounded-full bg-[#FEF2F2] px-2 py-0.5 text-[11px] font-medium text-[#A94442]"
              >
                <span>归档有缺口</span>
                <X className="h-3 w-3" />
              </button>
            )}
            {!showTrash && dateFilter && (
              <button
                type="button"
                onClick={() => setDateFilter(null)}
                className="inline-flex items-center gap-1 rounded-full bg-[#FEF9EC] px-2 py-0.5 text-[11px] font-medium text-[#9A6B26] font-mono"
              >
                <span>日期：{dateFilter}</span>
                <X className="h-3 w-3" />
              </button>
            )}
            {!showTrash && hasActiveFilter && (
              <button
                type="button"
                onClick={clearAllFilters}
                className="text-[11px] text-slate-400 underline hover:text-slate-600"
              >
                重置筛选
              </button>
            )}
          </div>

          <div className="flex flex-wrap items-center gap-2">
            {!showTrash && (
              <>
                {/* Search */}
                <div className="relative w-full sm:w-52">
                  <Search className="pointer-events-none absolute left-2.5 top-2 h-3.5 w-3.5 text-slate-400" />
                  <input
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    placeholder="搜索批次名、编号或单位..."
                    className="w-full rounded-lg border border-slate-200 bg-white py-1.5 pl-8 pr-3 text-xs focus:border-[#3B5B7E] focus:outline-none"
                  />
                </div>

                {/* Status Select */}
                <select
                  value={statusFilter}
                  onChange={(e) => setStatusFilter(e.target.value)}
                  className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs text-slate-600 focus:border-[#3B5B7E] focus:outline-none"
                >
                  <option value="all">全部状态</option>
                  {STATUS_ORDER.map((st) => (
                    <option key={st} value={st}>
                      {st}
                    </option>
                  ))}
                </select>

                {/* Hide Completed Toggle */}
                <label className="inline-flex cursor-pointer select-none items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs text-slate-600 hover:bg-slate-50">
                  <input
                    type="checkbox"
                    checked={hideCompleted}
                    onChange={(e) => setHideCompleted(e.target.checked)}
                    className="h-3.5 w-3.5 rounded border-slate-300 accent-[#3B5B7E]"
                  />
                  <span>隐藏已完成</span>
                </label>

                {/* View Switcher */}
                <div className="flex items-center rounded-lg border border-slate-200 bg-white p-0.5">
                  <button
                    type="button"
                    onClick={() => setViewMode('table')}
                    title="紧凑台账视图"
                    className={`inline-flex items-center gap-1 rounded-md px-2 py-1 text-xs font-medium transition-colors ${
                      viewMode === 'table'
                        ? 'bg-[#3B5B7E] text-white'
                        : 'text-slate-500 hover:text-slate-700'
                    }`}
                  >
                    <LayoutList className="h-3.5 w-3.5" />
                    <span>台账</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => setViewMode('cards')}
                    title="展开卡片视图"
                    className={`inline-flex items-center gap-1 rounded-md px-2 py-1 text-xs font-medium transition-colors ${
                      viewMode === 'cards'
                        ? 'bg-[#3B5B7E] text-white'
                        : 'text-slate-500 hover:text-slate-700'
                    }`}
                  >
                    <LayoutGrid className="h-3.5 w-3.5" />
                    <span>卡片</span>
                  </button>
                </div>
              </>
            )}

            {/* Recycle Bin Toggle Button */}
            <button
              type="button"
              onClick={() => setShowTrash((v) => !v)}
              className={`inline-flex items-center gap-1.5 rounded-lg border px-2.5 py-1.5 text-xs font-medium transition-colors ${
                showTrash
                  ? 'border-[#3B5B7E] bg-[#3B5B7E] text-white'
                  : trashBatches.length > 0
                    ? 'border-[#A94442]/30 bg-[#FEF2F2]/60 text-[#A94442] hover:bg-[#FEE2E2]'
                    : 'border-slate-200 bg-white text-slate-500 hover:bg-slate-50'
              }`}
            >
              <Trash2 className="h-3.5 w-3.5" />
              <span>{showTrash ? '返回台账' : `回收站${trashBatches.length > 0 ? ` (${trashBatches.length})` : ''}`}</span>
            </button>
          </div>
        </div>

        {/* Recycle Bin Panel */}
        {showTrash ? (
          trashBatches.length === 0 ? (
            <div className="py-12 text-center text-xs text-slate-400 space-y-1.5">
              <Trash2 className="mx-auto h-6 w-6 text-slate-300" />
              <div>回收站为空</div>
              <div className="text-[11px]">在台账中移入回收站的批次会暂存在这里，支持随时一键还原或带备份确认彻底清理</div>
            </div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full border-collapse text-left text-xs">
                <thead>
                  <tr className="border-b border-slate-200 bg-slate-50/50 text-[11px] text-slate-500">
                    <th className="px-4 py-2.5 font-medium">已移入回收站的批次</th>
                    <th className="px-3 py-2.5 font-medium whitespace-nowrap">项目单位</th>
                    <th className="px-3 py-2.5 font-medium whitespace-nowrap">包含数据规模</th>
                    <th className="px-3 py-2.5 font-medium whitespace-nowrap">移入时间</th>
                    <th className="px-4 py-2.5 font-medium whitespace-nowrap text-right">操作</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {trashBatches.map((b) => {
                    const st = b.stats;
                    const fileTotal = (st?.biddingFileCount || 0) + (st?.evalFileCount || 0);
                    return (
                      <tr key={b.id} className="hover:bg-slate-50/70">
                        <td className="px-4 py-3">
                          <div className="font-medium text-slate-800">{b.name}</div>
                          <div className="font-mono text-[10px] text-slate-400">
                            {b.code || '未编号'} · 原状态：{b.status}
                          </div>
                        </td>
                        <td className="px-3 py-3 text-slate-600 whitespace-nowrap">
                          {b.project_unit || '—'}
                        </td>
                        <td className="px-3 py-3 text-[11px] text-slate-500 whitespace-nowrap tabular-nums">
                          {st?.projectCount || 0} 标段 · {fileTotal} 附件 · {st?.awardOutputCount || 0} 交付件
                        </td>
                        <td className="px-3 py-3 font-mono text-[11px] text-slate-400 whitespace-nowrap tabular-nums">
                          {b.deleted_at ? b.deleted_at.slice(0, 16).replace('T', ' ') : '—'}
                        </td>
                        <td className="px-4 py-3 whitespace-nowrap text-right">
                          <div className="inline-flex items-center gap-2">
                            <button
                              type="button"
                              disabled={busyBatchId === b.id}
                              onClick={() => handleRestoreBatch(b.id)}
                              className="inline-flex items-center gap-1 rounded-md border border-[#2F6F5E]/30 bg-[#ECFDF5]/70 px-2.5 py-1 text-[11px] font-medium text-[#2F6F5E] hover:bg-[#ECFDF5] disabled:opacity-50"
                            >
                              <RotateCcw className="h-3 w-3" />
                              <span>一键还原</span>
                            </button>
                            <button
                              type="button"
                              disabled={busyBatchId === b.id}
                              onClick={() => setPendingHardDelete(b)}
                              className="inline-flex items-center gap-1 rounded-md border border-[#A94442]/30 bg-[#FEF2F2] px-2.5 py-1 text-[11px] font-medium text-[#A94442] hover:bg-[#FEE2E2] disabled:opacity-50"
                            >
                              <Trash2 className="h-3 w-3" />
                              <span>彻底删除</span>
                            </button>
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )
        ) : filteredBatches.length === 0 ? (
          <div className="py-12 text-center text-xs text-slate-400">
            {totalBatches === 0 ? (
              <span>暂无招标批次，请在左侧边栏新建第一个招标批次</span>
            ) : (
              <span>
                当前筛选条件下无匹配批次，
                <button
                  type="button"
                  onClick={clearAllFilters}
                  className="ml-1 text-[#3B5B7E] underline"
                >
                  清空筛选查看全部
                </button>
              </span>
            )}
          </div>
        ) : viewMode === 'table' ? (
          /* Compact Table Ledger Mode */
          <div className="overflow-x-auto">
            <table className="w-full border-collapse text-left text-xs">
              <thead>
                <tr className="border-b border-slate-200 bg-slate-50/50 text-[11px] text-slate-500">
                  <th className="px-3.5 py-2.5 font-medium whitespace-nowrap">状态</th>
                  <th className="px-3 py-2.5 font-medium min-w-[220px]">招标批次 / 编号</th>
                  <th className="px-3 py-2.5 font-medium whitespace-nowrap">项目单位</th>
                  <th className="px-3 py-2.5 font-medium whitespace-nowrap">标段 / 预算</th>
                  <th className="px-3 py-2.5 font-medium whitespace-nowrap">开标节点</th>
                  <th className="px-3 py-2.5 font-medium whitespace-nowrap w-32">SOP 进度</th>
                  <th className="px-3 py-2.5 font-medium whitespace-nowrap">归档必须项</th>
                  <th className="px-3.5 py-2.5 font-medium whitespace-nowrap text-right">快捷进入</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {paginate(filteredBatches, batchPager.page, batchPager.pageSize).map((b) => {
                  const st = b.stats;
                  const countdown = getCountdownLabel(b.bid_opening_time);
                  const sopPct = st.sopTotal ? Math.round((st.sopDone / st.sopTotal) * 100) : 0;
                  const reqPct = st.archiveReqTotal
                    ? Math.round((st.archiveReqPresent / st.archiveReqTotal) * 100)
                    : 0;
                  const meta = STATUS_META[b.status] || STATUS_META['筹备中'];
                  const isExpanded = expandedBatchId === b.id;

                  return (
                    <React.Fragment key={b.id}>
                      <tr className="hover:bg-slate-50/80 transition-colors">
                        <td className="px-3.5 py-2.5 whitespace-nowrap">
                          <span
                            className={`inline-flex items-center gap-1 rounded-md px-2 py-0.5 text-[11px] font-medium ${meta.bgClass} ${meta.textClass}`}
                          >
                            <span className={`h-1.5 w-1.5 rounded-full ${meta.dotClass}`} />
                            {b.status}
                          </span>
                        </td>
                        <td className="px-3 py-2.5">
                          <button
                            type="button"
                            onClick={() => setExpandedBatchId(isExpanded ? null : b.id)}
                            className="group flex items-start gap-1.5 text-left"
                          >
                            <span className="mt-0.5 text-slate-400 group-hover:text-[#3B5B7E]">
                              {isExpanded ? (
                                <ChevronUp className="h-3.5 w-3.5" />
                              ) : (
                                <ChevronDown className="h-3.5 w-3.5" />
                              )}
                            </span>
                            <div className="min-w-0">
                              <div className="font-medium text-slate-800 group-hover:text-[#3B5B7E] line-clamp-1">
                                {b.name}
                              </div>
                              <div className="font-mono text-[10px] text-slate-400">
                                {b.code || '未编号'}
                              </div>
                            </div>
                          </button>
                        </td>
                        <td className="px-3 py-2.5 text-slate-600 whitespace-nowrap max-w-[140px] truncate" title={b.project_unit}>
                          {b.project_unit || '—'}
                        </td>
                        <td className="px-3 py-2.5 whitespace-nowrap tabular-nums">
                          <div className="text-slate-700">{st.projectCount} 个标段</div>
                          {st.totalBudgetWan && st.totalBudgetWan > 0 ? (
                            <div className="text-[10px] text-slate-400">
                              {st.totalBudgetWan.toLocaleString()} 万元
                            </div>
                          ) : null}
                        </td>
                        <td className="px-3 py-2.5 whitespace-nowrap tabular-nums">
                          <div className="text-slate-700">{b.bid_opening_time || '待定'}</div>
                          {countdown && (
                            <span
                              className={`inline-block mt-0.5 rounded px-1.5 py-0.2 text-[10px] font-medium ${
                                countdown.urgent
                                  ? 'bg-[#FEF9EC] text-[#9A6B26]'
                                  : 'bg-slate-100 text-slate-500'
                              }`}
                            >
                              {countdown.text}
                            </span>
                          )}
                        </td>
                        <td className="px-3 py-2.5 whitespace-nowrap">
                          <div className="flex items-center gap-2">
                            <div className="h-1.5 w-16 overflow-hidden rounded-full bg-slate-200">
                              <div
                                className="h-full rounded-full bg-[#3B5B7E]"
                                style={{ width: `${sopPct}%` }}
                              />
                            </div>
                            <span className="font-mono text-[11px] text-slate-600 tabular-nums">
                              {sopPct}%
                            </span>
                          </div>
                          <div className="mt-0.5 text-[10px] text-slate-400 tabular-nums">
                            {st.sopDone}/{st.sopTotal} 步
                          </div>
                        </td>
                        <td className="px-3 py-2.5 whitespace-nowrap">
                          {st.archiveReqMissing > 0 ? (
                            <button
                              type="button"
                              onClick={() => goToModule(b.id, '/archive')}
                              className="inline-flex items-center gap-1 rounded-md bg-[#FEF2F2] px-2 py-0.5 text-[11px] font-medium text-[#A94442] hover:bg-[#FEE2E2] tabular-nums"
                            >
                              <AlertTriangle className="h-3 w-3" />
                              <span>缺 {st.archiveReqMissing} 项</span>
                            </button>
                          ) : (
                            <span className="inline-flex items-center gap-1 rounded-md bg-[#ECFDF5] px-2 py-0.5 text-[11px] font-medium text-[#2F6F5E]">
                              <CheckCircle2 className="h-3 w-3" />
                              <span>已齐 ({reqPct}%)</span>
                            </span>
                          )}
                        </td>
                        <td className="px-3.5 py-2.5 whitespace-nowrap text-right">
                          <div className="inline-flex items-center gap-1">
                            <button
                              type="button"
                              onClick={() => goToModule(b.id, '/bidding')}
                              title="进入招标阶段"
                              className="rounded border border-slate-200 bg-white px-2 py-1 text-[11px] font-medium text-slate-600 hover:border-[#3B5B7E] hover:text-[#3B5B7E]"
                            >
                              招
                            </button>
                            <button
                              type="button"
                              onClick={() => goToModule(b.id, '/evaluation')}
                              title="进入开评标 SOP"
                              className="rounded border border-slate-200 bg-white px-2 py-1 text-[11px] font-medium text-slate-600 hover:border-[#3B5B7E] hover:text-[#3B5B7E]"
                            >
                              评
                            </button>
                            <button
                              type="button"
                              onClick={() => goToModule(b.id, '/award')}
                              title="进入定标交付"
                              className="rounded border border-slate-200 bg-white px-2 py-1 text-[11px] font-medium text-slate-600 hover:border-[#3B5B7E] hover:text-[#3B5B7E]"
                            >
                              定
                            </button>
                            <button
                              type="button"
                              onClick={() => goToModule(b.id, '/archive')}
                              title="进入归档核查"
                              className="rounded border border-slate-200 bg-white px-2 py-1 text-[11px] font-medium text-slate-600 hover:border-[#3B5B7E] hover:text-[#3B5B7E]"
                            >
                              档
                            </button>
                          </div>
                        </td>
                      </tr>

                      {/* Inline Expandable 4-Stage Summary Strip + Safe Soft-Delete */}
                      {isExpanded && (
                        <tr className="bg-slate-50/70">
                          <td colSpan={8} className="px-4 py-3 space-y-2.5">
                            <div className="grid grid-cols-2 gap-2.5 sm:grid-cols-4">
                              <button
                                type="button"
                                onClick={() => goToModule(b.id, '/bidding')}
                                className="group flex items-center justify-between rounded-lg border border-slate-200 bg-white px-3 py-2 text-left hover:border-[#3B5B7E]/40"
                              >
                                <div>
                                  <div className="flex items-center gap-1.5 text-xs font-medium text-slate-700 group-hover:text-[#3B5B7E]">
                                    <FileSpreadsheet className="h-3.5 w-3.5" />
                                    1. 招标阶段
                                  </div>
                                  <div className="mt-1 text-[11px] text-slate-400 tabular-nums">
                                    {st.projectCount} 标段 · {st.biddingFileCount} 份前期表
                                  </div>
                                </div>
                                <ArrowRight className="h-3.5 w-3.5 text-slate-300 group-hover:text-[#3B5B7E]" />
                              </button>

                              <button
                                type="button"
                                onClick={() => goToModule(b.id, '/evaluation')}
                                className="group flex items-center justify-between rounded-lg border border-slate-200 bg-white px-3 py-2 text-left hover:border-[#3B5B7E]/40"
                              >
                                <div>
                                  <div className="flex items-center gap-1.5 text-xs font-medium text-slate-700 group-hover:text-[#3B5B7E]">
                                    <ClipboardCheck className="h-3.5 w-3.5" />
                                    2. 开评标 SOP
                                  </div>
                                  <div className="mt-1 text-[11px] text-slate-400 tabular-nums">
                                    已完成 {st.sopDone}/{st.sopTotal} · {st.evalFileCount} 份文件
                                  </div>
                                </div>
                                <ArrowRight className="h-3.5 w-3.5 text-slate-300 group-hover:text-[#3B5B7E]" />
                              </button>

                              <button
                                type="button"
                                onClick={() => goToModule(b.id, '/award')}
                                className="group flex items-center justify-between rounded-lg border border-slate-200 bg-white px-3 py-2 text-left hover:border-[#3B5B7E]/40"
                              >
                                <div>
                                  <div className="flex items-center gap-1.5 text-xs font-medium text-slate-700 group-hover:text-[#3B5B7E]">
                                    <Award className="h-3.5 w-3.5" />
                                    3. 定标交付
                                  </div>
                                  <div className="mt-1 text-[11px] text-slate-400 tabular-nums">
                                    已生成 {st.awardOutputCount} 份交付成果
                                  </div>
                                </div>
                                <ArrowRight className="h-3.5 w-3.5 text-slate-300 group-hover:text-[#3B5B7E]" />
                              </button>

                              <button
                                type="button"
                                onClick={() => goToModule(b.id, '/archive')}
                                className="group flex items-center justify-between rounded-lg border border-slate-200 bg-white px-3 py-2 text-left hover:border-[#3B5B7E]/40"
                              >
                                <div>
                                  <div className="flex items-center gap-1.5 text-xs font-medium text-slate-700 group-hover:text-[#3B5B7E]">
                                    <FolderArchive className="h-3.5 w-3.5" />
                                    4. 归档必须项
                                  </div>
                                  <div className="mt-1 text-[11px] text-slate-400 tabular-nums">
                                    必须项 {st.archiveReqPresent}/{st.archiveReqTotal} 已齐
                                  </div>
                                </div>
                                <ArrowRight className="h-3.5 w-3.5 text-slate-300 group-hover:text-[#3B5B7E]" />
                              </button>
                            </div>

                            <div className="flex items-center justify-end border-t border-slate-200/70 pt-2">
                              <button
                                type="button"
                                disabled={busyBatchId === b.id}
                                onClick={() => handleMoveToTrash(b)}
                                className="inline-flex items-center gap-1 rounded-md px-2 py-1 text-[11px] text-slate-400 hover:bg-[#FEF2F2] hover:text-[#A94442] transition-colors"
                              >
                                <Trash2 className="h-3 w-3" />
                                <span>移入回收站</span>
                              </button>
                            </div>
                          </td>
                        </tr>
                      )}
                    </React.Fragment>
                  );
                })}
              </tbody>
            </table>
          </div>
        ) : (
          /* Expanded Card Grid Mode (Contained inside the Ledger Panel) */
          <div className="grid grid-cols-1 gap-3 p-4 md:grid-cols-2">
            {paginate(filteredBatches, batchPager.page, batchPager.pageSize).map((b) => {
              const st = b.stats;
              const countdown = getCountdownLabel(b.bid_opening_time);
              const sopPct = st.sopTotal ? Math.round((st.sopDone / st.sopTotal) * 100) : 0;
              const reqPct = st.archiveReqTotal
                ? Math.round((st.archiveReqPresent / st.archiveReqTotal) * 100)
                : 0;
              const meta = STATUS_META[b.status] || STATUS_META['筹备中'];

              return (
                <div
                  key={b.id}
                  className="rounded-xl border border-slate-200 bg-slate-50/40 p-4 transition-colors hover:border-slate-300 hover:bg-white"
                >
                  <div className="flex items-start justify-between gap-2">
                    <div className="min-w-0 space-y-1">
                      <div className="flex flex-wrap items-center gap-1.5">
                        <span
                          className={`inline-flex items-center rounded px-2 py-0.5 text-[11px] font-medium ${meta.bgClass} ${meta.textClass}`}
                        >
                          {b.status}
                        </span>
                        <span className="font-mono text-[11px] text-slate-400">
                          {b.code || '未编号'}
                        </span>
                        {countdown && (
                          <span
                            className={`inline-flex items-center rounded px-1.5 py-0.5 text-[10px] font-medium ${
                              countdown.urgent
                                ? 'bg-[#FEF9EC] text-[#9A6B26]'
                                : 'bg-slate-100 text-slate-500'
                            }`}
                          >
                            {countdown.text}
                          </span>
                        )}
                      </div>
                      <h3 className="text-sm font-semibold text-slate-800 line-clamp-1">{b.name}</h3>
                      <div className="flex flex-wrap gap-x-3 gap-y-0.5 text-[11px] text-slate-500">
                        <span>{b.project_unit || '未填单位'}</span>
                        <span className="tabular-nums">{st.projectCount} 个标段</span>
                        <span className="tabular-nums">开标：{b.bid_opening_time || '待定'}</span>
                      </div>
                    </div>

                    <div className="flex items-center gap-1.5 shrink-0">
                      {st.archiveReqMissing > 0 ? (
                        <button
                          type="button"
                          onClick={() => goToModule(b.id, '/archive')}
                          className="inline-flex items-center gap-1 rounded-md bg-[#FEF2F2] px-2 py-1 text-[11px] font-medium text-[#A94442] hover:bg-[#FEE2E2]"
                        >
                          <AlertTriangle className="h-3 w-3" />
                          <span>缺 {st.archiveReqMissing}</span>
                        </button>
                      ) : (
                        <span className="inline-flex items-center gap-1 rounded-md bg-[#ECFDF5] px-2 py-1 text-[11px] font-medium text-[#2F6F5E]">
                          <CheckCircle2 className="h-3 w-3" />
                          <span>已齐</span>
                        </span>
                      )}
                      <button
                        type="button"
                        disabled={busyBatchId === b.id}
                        onClick={() => handleMoveToTrash(b)}
                        title="移入回收站"
                        className="rounded p-1 text-slate-300 hover:bg-[#FEF2F2] hover:text-[#A94442] transition-colors"
                      >
                        <Trash2 className="h-3.5 w-3.5" />
                      </button>
                    </div>
                  </div>

                  <div className="mt-3 grid grid-cols-4 gap-1.5 border-t border-slate-200/70 pt-3">
                    <button
                      type="button"
                      onClick={() => goToModule(b.id, '/bidding')}
                      className="rounded-lg border border-slate-200 bg-white px-2 py-1.5 text-left hover:border-[#3B5B7E]/40"
                    >
                      <div className="text-[11px] font-medium text-slate-700">1. 招标</div>
                      <div className="mt-0.5 text-[10px] text-slate-400 tabular-nums">
                        {st.biddingFileCount} 表
                      </div>
                    </button>
                    <button
                      type="button"
                      onClick={() => goToModule(b.id, '/evaluation')}
                      className="rounded-lg border border-slate-200 bg-white px-2 py-1.5 text-left hover:border-[#3B5B7E]/40"
                    >
                      <div className="text-[11px] font-medium text-slate-700">2. 评标</div>
                      <div className="mt-0.5 text-[10px] text-[#3B5B7E] tabular-nums">{sopPct}%</div>
                    </button>
                    <button
                      type="button"
                      onClick={() => goToModule(b.id, '/award')}
                      className="rounded-lg border border-slate-200 bg-white px-2 py-1.5 text-left hover:border-[#3B5B7E]/40"
                    >
                      <div className="text-[11px] font-medium text-slate-700">3. 定标</div>
                      <div className="mt-0.5 text-[10px] text-slate-400 tabular-nums">
                        {st.awardOutputCount} 份
                      </div>
                    </button>
                    <button
                      type="button"
                      onClick={() => goToModule(b.id, '/archive')}
                      className="rounded-lg border border-slate-200 bg-white px-2 py-1.5 text-left hover:border-[#3B5B7E]/40"
                    >
                      <div className="text-[11px] font-medium text-slate-700">4. 归档</div>
                      <div className="mt-0.5 text-[10px] text-[#2F6F5E] tabular-nums">{reqPct}%</div>
                    </button>
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {/* Pagination Footer */}
        {!showTrash && filteredBatches.length > 10 && (
          <div className="border-t border-slate-200">
            <PaginationBar
              total={filteredBatches.length}
              page={batchPager.page}
              pageSize={batchPager.pageSize}
              onPageChange={batchPager.setPage}
              onPageSizeChange={batchPager.setPageSize}
            />
          </div>
        )}
      </div>

      {/* Permanent Hard-Delete Confirmation Modal (with impact summary + backup regret pill + hard gate) */}
      {pendingHardDelete && (
        <BatchDeleteModal
          batch={pendingHardDelete}
          authEnabled={authEnabled}
          onCancel={() => setPendingHardDelete(null)}
          onConfirmDelete={handleConfirmHardDelete}
        />
      )}
    </div>
  );
}
