import React, { useEffect, useState } from 'react';
import { NavLink, Outlet, useLocation, useNavigate } from 'react-router-dom';
import {
  LayoutDashboard,
  FileSpreadsheet,
  ClipboardCheck,
  Award,
  FolderArchive,
  Plus,
  Download,
  Upload,
  Package,
  LogOut,
  Database,
  PanelLeftClose,
  PanelLeftOpen,
  X,
  ChevronDown,
  ChevronRight,
  Trophy,
  Scale,
  ShieldAlert,
  Calculator,
  FileText,
  ListChecks,
  Wrench,
} from 'lucide-react';
import { api } from '../api/client';
import { DateTimePickerField } from './DateTimePickerField';
import { SearchableBatchSelect } from './SearchableBatchSelect';
import { SearchableClientSelect } from './SearchableClientSelect';
import { useBatch } from '../context/BatchContext';
import type { ClientAgreement } from '../types';

const NAV_ITEMS = [
  { to: '/', label: '工作台总览', shortLabel: '总览', icon: LayoutDashboard, end: true },
  { to: '/bidding', label: '招标模块', shortLabel: '招标', icon: FileSpreadsheet, end: false },
  { to: '/evaluation', label: '开评标模块', shortLabel: '开评标', icon: ClipboardCheck, end: false },
  { to: '/award', label: '定标模块', shortLabel: '定标', icon: Award, end: false },
  { to: '/archive', label: '资料整理', shortLabel: '归档', icon: FolderArchive, end: false },
];

const EVAL_SUB_ITEMS = [
  { to: '/evaluation/results', label: '评审总览与数据', icon: Trophy },
  { to: '/evaluation/quote-tax', label: '报价与税率核验', icon: Scale },
  { to: '/evaluation/collusion', label: '围串标风控核验', icon: ShieldAlert },
  { to: '/evaluation/tools', label: '现场核验与算分', icon: Calculator },
  { to: '/evaluation/reports', label: '现场汇报与报告', icon: FileText },
  { to: '/evaluation/sop', label: 'SOP 与阶段附件', icon: ListChecks },
];

export function Layout() {
  const {
    batches,
    currentBatchId,
    setCurrentBatchId,
    refreshBatches,
    storageMode,
    authEnabled,
    setAuthenticated,
  } = useBatch();
  const navigate = useNavigate();
  const location = useLocation();
  const isEvalRoute = location.pathname.startsWith('/evaluation');

  const [collapsed, setCollapsed] = useState(false);
  const [evalExpanded, setEvalExpanded] = useState(true);
  const [showNewModal, setShowNewModal] = useState(false);

  useEffect(() => {
    if (isEvalRoute) setEvalExpanded(true);
  }, [isEvalRoute]);
  const [clients, setClients] = useState<ClientAgreement[]>([]);
  const [newName, setNewName] = useState('');
  const [newCode, setNewCode] = useState('');
  const [newUnit, setNewUnit] = useState('');
  const [newOpenTime, setNewOpenTime] = useState('');
  const [creating, setCreating] = useState(false);
  const [importingBatch, setImportingBatch] = useState(false);
  const [importingFull, setImportingFull] = useState(false);

  useEffect(() => {
    if (showNewModal && clients.length === 0) {
      api.listClients().then(setClients).catch(() => {});
    }
  }, [showNewModal, clients.length]);

  async function handleCreateBatch(e: React.FormEvent) {
    e.preventDefault();
    if (!newName.trim()) return;
    setCreating(true);
    try {
      const created = await api.createBatch({
        name: newName.trim(),
        code: newCode.trim(),
        project_unit: newUnit.trim(),
        bid_opening_time: newOpenTime.trim(),
        status: '筹备中',
      });
      await refreshBatches();
      setCurrentBatchId(created.id);
      setShowNewModal(false);
      setNewName('');
      setNewCode('');
      setNewUnit('');
      setNewOpenTime('');
      navigate('/bidding');
    } finally {
      setCreating(false);
    }
  }

  async function handleLogout() {
    await api.logout();
    setAuthenticated(false);
  }

  async function handleImportBatchZip(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;
    setImportingBatch(true);
    try {
      const res = await api.importBatchPackage(file);
      await refreshBatches();
      setCurrentBatchId(res.batch.id);
      const verb = res.action === 'updated' ? '已无损更新' : '已新增导入';
      const label = [res.batch.code, res.batch.name].filter(Boolean).join(' ');
      alert(
        `${verb}批次「${label}」！\n` +
          `项目 ${res.counts.projects} 条 · 附件记录 ${res.counts.files} 份 · 定标产物 ${res.counts.awardOutputs} 份 · 恢复磁盘文件 ${res.counts.extractedFiles} 个。\n` +
          `（工作台内其他批次均不受影响）`
      );
    } catch (err) {
      alert((err as Error).message || '导入批次包失败');
    } finally {
      setImportingBatch(false);
    }
  }

  async function handleImportFullBackupZip(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;
    const confirmed = window.confirm(
      '确认导入整库备份吗？\n\n此操作将用备份包内的完整数据库与附件覆盖当前工作台的全部内容（所有批次、客户协议、物料库与文件）。\n\n如果只想无损合并单个批次，请取消并改用「导入批次包 (.zip)」。'
    );
    if (!confirmed) return;
    setImportingFull(true);
    try {
      const res = await api.importFullBackup(file);
      await refreshBatches();
      if (res.firstBatchId) {
        setCurrentBatchId(res.firstBatchId);
      }
      alert(
        `整库备份已恢复完成！\n` +
          `共恢复 ${res.counts.batches} 个批次 · ${res.counts.projects} 条项目 · ${res.counts.files} 份附件记录 · ${res.counts.extractedFiles} 个磁盘文件。`
      );
      window.location.reload();
    } catch (err) {
      alert((err as Error).message || '恢复整库备份失败');
    } finally {
      setImportingFull(false);
    }
  }

  return (
    <div className="min-h-screen flex bg-[#F7F8FA] text-[#1E293B]">
      {/* PC Left Sidebar (Sticky full viewport height) */}
      <aside
        className={`hidden md:flex flex-col shrink-0 sticky top-0 h-screen border-r border-slate-200 bg-white transition-all duration-200 ${
          collapsed ? 'w-[68px]' : 'w-[228px]'
        }`}
      >
        <div className="flex h-14 items-center justify-between border-b border-slate-200 px-4">
          {!collapsed && (
            <div className="flex items-center gap-2 min-w-0">
              <div className="flex h-7 w-7 shrink-0 items-center justify-center rounded-lg bg-[#3B5B7E] text-white text-xs font-bold">
                招
              </div>
              <span className="truncate text-sm font-semibold tracking-tight text-slate-800">
                招标代理 SOP 工作台
              </span>
            </div>
          )}
          <button
            onClick={() => setCollapsed((v) => !v)}
            className="rounded-md p-1.5 text-slate-400 hover:bg-slate-100 hover:text-slate-600"
            title={collapsed ? '展开侧边栏' : '收起侧边栏'}
          >
            {collapsed ? <PanelLeftOpen className="h-4 w-4" /> : <PanelLeftClose className="h-4 w-4" />}
          </button>
        </div>

        <nav className="flex-1 overflow-y-auto space-y-1 p-2.5">
          {NAV_ITEMS.map((item) => {
            const Icon = item.icon;
            if (item.to === '/evaluation') {
              return (
                <div key={item.to} className="space-y-0.5">
                  <div
                    className={`flex items-center justify-between rounded-lg transition-colors ${
                      isEvalRoute
                        ? 'bg-[#3B5B7E]/10 text-[#3B5B7E]'
                        : 'text-slate-600 hover:bg-slate-100/80 hover:text-slate-900'
                    }`}
                  >
                    <button
                      type="button"
                      onClick={() => {
                        if (collapsed) {
                          navigate('/evaluation/results');
                        } else {
                          if (!isEvalRoute) navigate('/evaluation/results');
                          setEvalExpanded((v) => (isEvalRoute ? !v : true));
                        }
                      }}
                      className="flex flex-1 items-center gap-3 px-3 py-2.5 text-sm font-medium text-left min-w-0"
                      title={collapsed ? item.label : undefined}
                    >
                      <Icon className="h-4 w-4 shrink-0" />
                      {!collapsed && <span className="truncate">{item.label}</span>}
                    </button>
                    {!collapsed && (
                      <button
                        type="button"
                        onClick={() => setEvalExpanded((v) => !v)}
                        className="mr-1.5 rounded p-1 text-slate-400 hover:bg-slate-200/60 hover:text-slate-600"
                        title={evalExpanded ? '收起子模块' : '展开子模块'}
                      >
                        {evalExpanded ? <ChevronDown className="h-3.5 w-3.5" /> : <ChevronRight className="h-3.5 w-3.5" />}
                      </button>
                    )}
                  </div>
                  {!collapsed && evalExpanded && (
                    <div className="ml-3.5 space-y-0.5 border-l border-slate-200 pl-2.5 py-0.5">
                      {EVAL_SUB_ITEMS.map((sub) => {
                        const SubIcon = sub.icon;
                        const isDefaultResults =
                          sub.to === '/evaluation/results' &&
                          (location.pathname === '/evaluation' || location.pathname === '/evaluation/rejections');
                        return (
                          <NavLink
                            key={sub.to}
                            to={sub.to}
                            className={({ isActive }) =>
                              `flex items-center gap-2 rounded-md px-2.5 py-1.5 text-xs font-medium transition-colors ${
                                isActive || isDefaultResults
                                  ? 'bg-[#3B5B7E] text-white'
                                  : 'text-slate-600 hover:bg-slate-100 hover:text-slate-900'
                              }`
                            }
                          >
                            <SubIcon className="h-3.5 w-3.5 shrink-0" />
                            <span className="truncate">{sub.label}</span>
                          </NavLink>
                        );
                      })}
                    </div>
                  )}
                </div>
              );
            }
            return (
              <NavLink
                key={item.to}
                to={item.to}
                end={item.end}
                className={({ isActive }) =>
                  `flex items-center gap-3 rounded-lg px-3 py-2.5 text-sm font-medium transition-colors ${
                    isActive
                      ? 'bg-[#3B5B7E]/10 text-[#3B5B7E]'
                      : 'text-slate-600 hover:bg-slate-100/80 hover:text-slate-900'
                  }`
                }
                title={collapsed ? item.label : undefined}
              >
                <Icon className="h-4 w-4 shrink-0" />
                {!collapsed && <span className="truncate">{item.label}</span>}
              </NavLink>
            );
          })}

          <div className="pt-2 mt-2 border-t border-slate-200/80">
            {!collapsed && (
              <div className="px-3 pb-1 text-[10px] font-semibold tracking-wide text-slate-400">
                跨阶段辅助
              </div>
            )}
            <NavLink
              to="/toolbox"
              className={({ isActive }) =>
                `flex items-center gap-3 rounded-lg px-3 py-2.5 text-sm font-medium transition-colors ${
                  isActive
                    ? 'bg-[#3B5B7E]/10 text-[#3B5B7E]'
                    : 'text-slate-600 hover:bg-slate-100/80 hover:text-slate-900'
                }`
              }
              title={collapsed ? '实用工具箱' : undefined}
            >
              <Wrench className="h-4 w-4 shrink-0" />
              {!collapsed && <span className="truncate">实用工具箱</span>}
            </NavLink>
          </div>
        </nav>

        <div className="border-t border-slate-200 p-3 space-y-2">
          {!collapsed && (
            <div className="flex items-center gap-2 rounded-lg bg-slate-50 px-2.5 py-2 text-xs text-slate-500">
              <Database className="h-3.5 w-3.5 shrink-0 text-[#3B5B7E]" />
              <span className="truncate">
                {storageMode === 'sqlite' ? 'SQLite 本地持久化' : 'IndexedDB 离线模式'}
              </span>
            </div>
          )}
          {storageMode === 'sqlite' && (
            <div className="space-y-1.5">
              {currentBatchId && (
                <a
                  href={`/api/backup/export-batch/${currentBatchId}`}
                  className="flex items-center justify-center gap-2 rounded-lg border border-[#3B5B7E]/30 bg-[#EBF2F7]/60 px-2.5 py-1.5 text-xs font-medium text-[#3B5B7E] hover:bg-[#EBF2F7] transition-colors"
                  title="仅导出当前批次数据库记录与附件 (.zip)，适合带出开标现场"
                >
                  <Package className="h-3.5 w-3.5 shrink-0" />
                  {!collapsed && <span>导出当前批次 (.zip)</span>}
                </a>
              )}
              <label
                className="flex cursor-pointer items-center justify-center gap-2 rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-medium text-slate-600 hover:border-[#3B5B7E] hover:text-[#3B5B7E] transition-colors"
                title="无损合并导入单批次 .zip 包（不影响其他批次）"
              >
                <Upload className="h-3.5 w-3.5 shrink-0" />
                {!collapsed && <span>{importingBatch ? '导入中...' : '导入批次包 (.zip)'}</span>}
                <input
                  type="file"
                  accept=".zip"
                  onChange={handleImportBatchZip}
                  disabled={importingBatch}
                  className="hidden"
                />
              </label>
              <a
                href="/api/backup/export"
                className="flex items-center justify-center gap-2 rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-medium text-slate-500 hover:bg-slate-50 hover:text-slate-700 transition-colors"
                title="导出完整数据库与全部附件 (ZIP)"
              >
                <Download className="h-3.5 w-3.5 shrink-0" />
                {!collapsed && <span>导出整库备份</span>}
              </a>
              <label
                className="flex cursor-pointer items-center justify-center gap-2 rounded-lg border border-amber-200 bg-amber-50/50 px-2.5 py-1.5 text-xs font-medium text-amber-700 hover:bg-amber-50 transition-colors"
                title="导入完整工作台备份 .zip（将覆盖当前全部数据，适合新机器一键同步镜像）"
              >
                <Database className="h-3.5 w-3.5 shrink-0" />
                {!collapsed && <span>{importingFull ? '整库恢复中...' : '恢复整库备份 (.zip)'}</span>}
                <input
                  type="file"
                  accept=".zip"
                  onChange={handleImportFullBackupZip}
                  disabled={importingFull}
                  className="hidden"
                />
              </label>
            </div>
          )}
          {authEnabled && (
            <button
              onClick={handleLogout}
              className="flex w-full items-center justify-center gap-2 rounded-lg px-2.5 py-2 text-xs font-medium text-slate-500 hover:bg-slate-100 hover:text-slate-700"
            >
              <LogOut className="h-3.5 w-3.5 shrink-0" />
              {!collapsed && <span>锁定退出</span>}
            </button>
          )}
        </div>
      </aside>

      {/* Main Column */}
      <div className="flex flex-1 flex-col min-w-0 pb-16 md:pb-0">
        {/* Top Global Context Bar */}
        <header className="sticky top-0 z-20 flex h-14 items-center justify-between gap-3 border-b border-slate-200 bg-white/95 px-4 backdrop-blur-sm md:px-6">
          <div className="flex items-center gap-2.5 min-w-0 flex-1">
            <span className="hidden sm:inline text-xs font-medium text-slate-400 shrink-0">当前批次</span>
            <SearchableBatchSelect
              batches={batches}
              value={currentBatchId}
              onChange={setCurrentBatchId}
            />
          </div>

          <button
            onClick={() => setShowNewModal(true)}
            className="inline-flex shrink-0 items-center gap-1.5 rounded-lg bg-[#3B5B7E] px-3 py-1.5 text-xs font-medium text-white hover:bg-[#2E4866] transition-colors"
          >
            <Plus className="h-3.5 w-3.5" />
            <span className="hidden sm:inline">新建批次</span>
            <span className="sm:hidden">新建</span>
          </button>
        </header>

        {/* Page Content */}
        <main className="flex-1 p-4 md:p-6 max-w-7xl w-full mx-auto">
          <Outlet />
        </main>
      </div>

      {/* Mobile Bottom Navigation */}
      <nav className="fixed bottom-0 left-0 right-0 z-30 grid grid-cols-6 border-t border-slate-200 bg-white md:hidden">
        {NAV_ITEMS.map((item) => {
          const Icon = item.icon;
          return (
            <NavLink
              key={item.to}
              to={item.to}
              end={item.end}
              className={({ isActive }) =>
                `flex flex-col items-center justify-center py-2 text-[11px] font-medium transition-colors ${
                  isActive ? 'text-[#3B5B7E]' : 'text-slate-500'
                }`
              }
            >
              <Icon className="h-4 w-4 mb-0.5" />
              <span>{item.shortLabel}</span>
            </NavLink>
          );
        })}
        <NavLink
          to="/toolbox"
          className={({ isActive }) =>
            `flex flex-col items-center justify-center py-2 text-[11px] font-medium transition-colors ${
              isActive ? 'text-[#3B5B7E]' : 'text-slate-500'
            }`
          }
        >
          <Wrench className="h-4 w-4 mb-0.5" />
          <span>工具箱</span>
        </NavLink>
      </nav>

      {/* New Batch Modal */}
      {showNewModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4 backdrop-blur-[1px]">
          <div className="w-full max-w-md rounded-xl border border-slate-200 bg-white p-5 shadow-xl">
            <div className="flex items-center justify-between mb-4">
              <h3 className="text-base font-semibold text-slate-800">新建招标批次</h3>
              <button
                onClick={() => setShowNewModal(false)}
                className="rounded-lg p-1 text-slate-400 hover:bg-slate-100"
              >
                <X className="h-4 w-4" />
              </button>
            </div>
            <form onSubmit={handleCreateBatch} className="space-y-3.5">
              <div>
                <label className="block text-xs font-medium text-slate-600 mb-1">批次名称 *</label>
                <input
                  required
                  value={newName}
                  onChange={(e) => setNewName(e.target.value)}
                  placeholder="例如：2026年12月某单位信息化运维采购批次"
                  className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1">批次编号</label>
                  <input
                    value={newCode}
                    onChange={(e) => setNewCode(e.target.value)}
                    placeholder="GW-AH-2026-XXXX"
                    className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm focus:border-[#3B5B7E] focus:outline-none"
                  />
                </div>
                <div className="col-span-2">
                  <label className="block text-xs font-medium text-slate-600 mb-1">预计开标时间（日期必选，具体时间可选）</label>
                  <DateTimePickerField
                    value={newOpenTime}
                    onChange={setNewOpenTime}
                    allowTime
                  />
                </div>
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-600 mb-1">项目单位（招标人）</label>
                <SearchableClientSelect
                  clients={clients}
                  value={newUnit}
                  valueMode="name"
                  allowCustom
                  placeholder="搜索并选择协议项目单位..."
                  onChange={(val) => setNewUnit(val)}
                />
              </div>
              <p className="text-xs text-slate-400">
                创建后将自动套用默认的【开评标 SOP 步骤】与【归档/重要保管文件清单】模板。
              </p>
              <div className="flex justify-end gap-2 pt-2">
                <button
                  type="button"
                  onClick={() => setShowNewModal(false)}
                  className="rounded-lg border border-slate-200 px-4 py-2 text-xs font-medium text-slate-600 hover:bg-slate-50"
                >
                  取消
                </button>
                <button
                  type="submit"
                  disabled={creating}
                  className="rounded-lg bg-[#3B5B7E] px-4 py-2 text-xs font-medium text-white hover:bg-[#2E4866] disabled:opacity-50"
                >
                  {creating ? '创建中...' : '确认创建'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}
