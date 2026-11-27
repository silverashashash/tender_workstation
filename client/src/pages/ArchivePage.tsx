import React, { useCallback, useEffect, useState } from 'react';
import {
  Upload,
  FileSpreadsheet,
  Download,
  Archive,
  CheckCircle2,
  AlertTriangle,
  MinusCircle,
  ChevronDown,
  ChevronUp,
  Paperclip,
  Link2,
  Plus,
  Trash2,
  Search,
  X,
} from 'lucide-react';
import { api } from '../api/client';
import { paginate, PaginationBar, usePagination } from '../components/PaginationBar';
import { useBatch } from '../context/BatchContext';
import type { ArchiveItem, ArchiveOverview, ParsedArchivePreviewRow } from '../types';

type FilterMode = 'all' | 'req_missing' | 'missing' | 'present' | 'na';

export function ArchivePage() {
  const { currentBatchId, refreshBatches } = useBatch();
  const [data, setData] = useState<ArchiveOverview | null>(null);
  const [activeSheet, setActiveSheet] = useState<string>('ALL');
  const [filterMode, setFilterMode] = useState<FilterMode>('all');
  const [search, setSearch] = useState('');
  const [expandedIds, setExpandedIds] = useState<Set<string>>(new Set());
  const archivePager = usePagination(`${currentBatchId}-${activeSheet}-${filterMode}-${search}`);

  // Import Excel preview modal state
  const [importing, setImporting] = useState(false);
  const [importPreview, setImportPreview] = useState<{
    filename: string;
    sheetNames: string[];
    items: ParsedArchivePreviewRow[];
    colorDetectedCount: number;
    format: string;
  } | null>(null);
  const importPager = usePagination(importPreview?.filename || '');
  const [importReplace, setImportReplace] = useState(true);
  const [importSaveTpl, setImportSaveTpl] = useState(true);
  const [committing, setCommitting] = useState(false);

  // Cross-module link modal state
  const [linkingItem, setLinkingItem] = useState<ArchiveItem | null>(null);

  // Add manual item modal state
  const [showAddModal, setShowAddModal] = useState(false);
  const [addSheet, setAddSheet] = useState('归档文件');
  const [addSeq, setAddSeq] = useState('');
  const [addName, setAddName] = useState('');
  const [addRequired, setAddRequired] = useState(true);
  const [addMedia, setAddMedia] = useState('纸质+电子');

  const loadData = useCallback(async () => {
    if (!currentBatchId) return;
    const res = await api.getArchiveOverview(currentBatchId);
    setData(res);
  }, [currentBatchId]);

  useEffect(() => {
    loadData();
  }, [loadData]);

  function toggleExpand(id: string) {
    setExpandedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  async function handleImportSelect(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file || !currentBatchId) return;
    setImporting(true);
    try {
      const preview = await api.previewArchiveImport(currentBatchId, file);
      setImportPreview(preview);
    } catch (err) {
      alert((err as Error).message);
    } finally {
      setImporting(false);
      e.target.value = '';
    }
  }

  async function handleCommitImport() {
    if (!currentBatchId || !importPreview) return;
    setCommitting(true);
    try {
      await api.commitArchiveImport(currentBatchId, importPreview.items, importReplace, importSaveTpl);
      setImportPreview(null);
      setActiveSheet('ALL');
      await loadData();
      await refreshBatches();
    } finally {
      setCommitting(false);
    }
  }

  async function updateStatus(item: ArchiveItem, status: 'present' | 'missing' | 'na') {
    await api.updateArchiveItem(item.id, {
      status,
      is_manual_marked: status === 'present' && item.linked_file_ids.length === 0,
    });
    await loadData();
    await refreshBatches();
  }

  async function toggleRequired(item: ArchiveItem) {
    await api.updateArchiveItem(item.id, { is_required: !item.is_required });
    await loadData();
    await refreshBatches();
  }

  async function updateNotes(item: ArchiveItem, notes: string) {
    await api.updateArchiveItem(item.id, { notes });
    await loadData();
  }

  async function handleItemFileUpload(item: ArchiveItem, e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file || !currentBatchId) return;
    await api.uploadArchiveItemFile(item.id, currentBatchId, file);
    await loadData();
    await refreshBatches();
    e.target.value = '';
  }

  async function toggleLinkedFile(item: ArchiveItem, fileId: string) {
    const has = item.linked_file_ids.includes(fileId);
    const nextIds = has
      ? item.linked_file_ids.filter((id) => id !== fileId)
      : [...item.linked_file_ids, fileId];
    await api.updateArchiveItem(item.id, {
      linked_file_ids: nextIds,
      status: nextIds.length > 0 ? 'present' : item.status,
    });
    await loadData();
    await refreshBatches();
    setLinkingItem((prev) => (prev ? { ...prev, linked_file_ids: nextIds } : null));
  }

  async function handleDeleteItem(id: string) {
    await api.deleteArchiveItem(id);
    await loadData();
    await refreshBatches();
  }

  async function handleAddManualItem(e: React.FormEvent) {
    e.preventDefault();
    if (!currentBatchId || !addName.trim()) return;
    await api.addArchiveItem(currentBatchId, {
      sheet_name: addSheet,
      seq: addSeq,
      doc_name: addName.trim(),
      is_required: addRequired,
      media_type: addMedia,
    });
    setShowAddModal(false);
    setAddSeq('');
    setAddName('');
    await loadData();
    await refreshBatches();
  }

  if (!data) {
    return <div className="py-12 text-center text-sm text-slate-400">正在加载归档资料清单...</div>;
  }

  const { summary, sheetNames, items, availableBatchFiles } = data;
  const reqPct = summary.requiredTotal
    ? Math.round((summary.requiredPresent / summary.requiredTotal) * 100)
    : 100;
  const allPct = summary.activeTotal
    ? Math.round((summary.present / summary.activeTotal) * 100)
    : 100;

  const filteredItems = items.filter((it) => {
    if (activeSheet !== 'ALL' && it.sheet_name !== activeSheet) return false;
    if (filterMode === 'req_missing' && !(it.is_required && it.status === 'missing')) return false;
    if (filterMode === 'missing' && it.status !== 'missing') return false;
    if (filterMode === 'present' && it.status !== 'present') return false;
    if (filterMode === 'na' && it.status !== 'na') return false;
    if (search.trim()) {
      const q = search.trim().toLowerCase();
      return (
        it.doc_name.toLowerCase().includes(q) ||
        it.seq.toLowerCase().includes(q) ||
        it.notes.toLowerCase().includes(q)
      );
    }
    return true;
  });

  return (
    <div className="space-y-5">
      {/* Header & Top Actions */}
      <div className="flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
        <div>
          <h1 className="text-lg font-semibold text-slate-800">4. 资料整理模块（归档进度追踪）</h1>
          <p className="mt-1 text-xs text-slate-500">
            支持导入多 Sheet 归档目录 Excel（自动识别 B 列颜色标注为“必须”），逐项上传文件、手动标记“已有”或引用前序模块文件。
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <label className="inline-flex cursor-pointer items-center gap-1.5 rounded-lg bg-[#3B5B7E] px-3 py-1.5 text-xs font-medium text-white hover:bg-[#2E4866] transition-colors">
            <FileSpreadsheet className="h-3.5 w-3.5" />
            <span>{importing ? '解析目录中...' : '导入归档目录 Excel'}</span>
            <input
              type="file"
              accept=".xls,.xlsx"
              onChange={handleImportSelect}
              className="hidden"
              disabled={importing}
            />
          </label>
          <button
            onClick={() => setShowAddModal(true)}
            className="inline-flex items-center gap-1 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50"
          >
            <Plus className="h-3.5 w-3.5" /> 新增条目
          </button>
          <a
            href={`/api/archive/batch/${currentBatchId}/export-xlsx`}
            className="inline-flex items-center gap-1 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50"
          >
            <Download className="h-3.5 w-3.5" /> 导出核对表
          </a>
          <a
            href={`/api/archive/batch/${currentBatchId}/export-zip`}
            className="inline-flex items-center gap-1 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50"
          >
            <Archive className="h-3.5 w-3.5" /> 打包电子件(ZIP)
          </a>
        </div>
      </div>

      {/* Progress Dashboard */}
      <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
        <div
          className={`rounded-xl border p-4 ${
            summary.requiredMissing > 0
              ? 'border-[#A94442]/30 bg-[#FEF2F2]/50'
              : 'border-[#2F6F5E]/30 bg-[#ECFDF5]/50'
          }`}
        >
          <div className="flex items-center justify-between text-xs">
            <span className="font-semibold text-slate-700">必须项归档进度（核心）</span>
            <span
              className={`font-semibold tabular-nums ${
                summary.requiredMissing > 0 ? 'text-[#A94442]' : 'text-[#2F6F5E]'
              }`}
            >
              {reqPct}%
            </span>
          </div>
          <div className="mt-2 flex items-baseline gap-2">
            <span className="text-2xl font-semibold text-slate-800 tabular-nums">
              {summary.requiredPresent} / {summary.requiredTotal}
            </span>
            {summary.requiredMissing > 0 ? (
              <span className="text-xs font-medium text-[#A94442]">还缺 {summary.requiredMissing} 项必须资料</span>
            ) : (
              <span className="text-xs font-medium text-[#2F6F5E]">必须资料已全部齐备</span>
            )}
          </div>
          <div className="mt-2.5 h-1.5 w-full overflow-hidden rounded-full bg-white/80">
            <div
              className={`h-full rounded-full ${summary.requiredMissing > 0 ? 'bg-[#A94442]' : 'bg-[#2F6F5E]'}`}
              style={{ width: `${reqPct}%` }}
            />
          </div>
        </div>

        <div className="rounded-xl border border-slate-200 bg-white p-4">
          <div className="flex items-center justify-between text-xs text-slate-500">
            <span>总体归档进度（不含“无”）</span>
            <span className="font-semibold text-[#3B5B7E] tabular-nums">{allPct}%</span>
          </div>
          <div className="mt-2 flex items-baseline gap-2">
            <span className="text-2xl font-semibold text-slate-800 tabular-nums">
              {summary.present} / {summary.activeTotal}
            </span>
            <span className="text-xs text-slate-400">共缺失 {summary.missing} 项</span>
          </div>
          <div className="mt-2.5 h-1.5 w-full overflow-hidden rounded-full bg-slate-100">
            <div className="h-full rounded-full bg-[#3B5B7E]" style={{ width: `${allPct}%` }} />
          </div>
        </div>

        <div className="rounded-xl border border-slate-200 bg-white p-4">
          <div className="text-xs text-slate-500">非必须项 & 不涉及（无）</div>
          <div className="mt-2 flex items-baseline gap-3">
            <div>
              <span className="text-2xl font-semibold text-slate-800 tabular-nums">
                {summary.optionalPresent}/{summary.optionalTotal}
              </span>
              <span className="ml-1 text-xs text-slate-400">非必须已齐</span>
            </div>
            <div className="text-xs text-slate-400 border-l border-slate-200 pl-3 tabular-nums">
              已标“无”：{summary.na} 项
            </div>
          </div>
          <p className="mt-2 text-[11px] text-slate-400">
            标记为“不涉及(无)”的条目不计入缺失预警分母
          </p>
        </div>
      </div>

      {/* Sheet Tabs + Filters */}
      <div className="rounded-xl border border-slate-200 bg-white p-4 space-y-3">
        <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
          {/* Multi-Sheet Tabs */}
          <div className="flex flex-wrap items-center gap-1.5">
            <button
              onClick={() => setActiveSheet('ALL')}
              className={`rounded-lg px-3 py-1.5 text-xs font-medium transition-colors ${
                activeSheet === 'ALL'
                  ? 'bg-[#3B5B7E] text-white'
                  : 'bg-slate-100 text-slate-600 hover:bg-slate-200/70'
              }`}
            >
              全部 ({items.length})
            </button>
            {sheetNames.map((sn) => (
              <button
                key={sn}
                onClick={() => setActiveSheet(sn)}
                className={`rounded-lg px-3 py-1.5 text-xs font-medium transition-colors ${
                  activeSheet === sn
                    ? 'bg-[#3B5B7E] text-white'
                    : 'bg-slate-100 text-slate-600 hover:bg-slate-200/70'
                }`}
              >
                {sn} ({items.filter((i) => i.sheet_name === sn).length})
              </button>
            ))}
          </div>

          {/* Search */}
          <div className="relative w-full md:w-64">
            <Search className="pointer-events-none absolute left-2.5 top-2 h-3.5 w-3.5 text-slate-400" />
            <input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="搜索资料名称、序号或备注..."
              className="w-full rounded-lg border border-slate-200 py-1.5 pl-8 pr-3 text-xs focus:border-[#3B5B7E] focus:outline-none"
            />
          </div>
        </div>

        {/* Status Filter Pills */}
        <div className="flex flex-wrap items-center gap-1.5 border-t border-slate-100 pt-3">
          {(
            [
              { key: 'all', label: '全部状态' },
              { key: 'req_missing', label: `只看缺失的必须项 (${summary.requiredMissing})`, highlight: true },
              { key: 'missing', label: `全部缺失 (${summary.missing})` },
              { key: 'present', label: `已有 (${summary.present})` },
              { key: 'na', label: `不涉及/无 (${summary.na})` },
            ] as Array<{ key: FilterMode; label: string; highlight?: boolean }>
          ).map((f) => (
            <button
              key={f.key}
              onClick={() => setFilterMode(f.key)}
              className={`rounded-md px-2.5 py-1 text-xs font-medium transition-colors ${
                filterMode === f.key
                  ? f.highlight
                    ? 'bg-[#A94442] text-white'
                    : 'bg-slate-800 text-white'
                  : f.highlight && summary.requiredMissing > 0
                    ? 'bg-[#FEF2F2] text-[#A94442] hover:bg-[#FEE2E2]'
                    : 'bg-slate-50 text-slate-600 hover:bg-slate-100'
              }`}
            >
              {f.label}
            </button>
          ))}
        </div>

        {/* Checklist Items */}
        <div className="divide-y divide-slate-100 border-t border-slate-200 pt-1">
          {filteredItems.length === 0 ? (
            <div className="py-10 text-center text-xs text-slate-400">当前筛选条件下暂无归档条目</div>
          ) : (
            paginate(filteredItems, archivePager.page, archivePager.pageSize).map((item) => {
              const expanded = expandedIds.has(item.id);
              const isReqMissing = item.is_required && item.status === 'missing';

              return (
                <div
                  key={item.id}
                  className={`py-3 px-2 rounded-lg transition-colors ${
                    isReqMissing ? 'bg-[#FEF2F2]/35' : 'hover:bg-slate-50/70'
                  }`}
                >
                  <div className="flex flex-col gap-2.5 lg:flex-row lg:items-center lg:justify-between">
                    {/* Left Info */}
                    <div className="flex items-start gap-2.5 min-w-0 flex-1">
                      <span className="mt-0.5 inline-flex h-5 min-w-[26px] shrink-0 items-center justify-center rounded bg-slate-100 px-1.5 text-[11px] font-mono text-slate-500 tabular-nums">
                        {item.seq || '-'}
                      </span>
                      <div className="min-w-0 flex-1">
                        <div className="flex flex-wrap items-center gap-1.5">
                          <span className="text-sm font-medium text-slate-800">{item.doc_name}</span>
                          <button
                            onClick={() => toggleRequired(item)}
                            title="点击切换必须/非必须"
                            className={`rounded px-1.5 py-0.5 text-[11px] font-medium transition-colors ${
                              item.is_required
                                ? 'bg-[#FEF9EC] text-[#9A6B26] border border-[#9A6B26]/20'
                                : 'bg-slate-100 text-slate-500'
                            }`}
                          >
                            {item.is_required ? '必须' : '非必须'}
                          </button>
                          <span className="rounded bg-slate-100/80 px-1.5 py-0.5 text-[11px] text-slate-500">
                            {item.media_type}
                          </span>
                          {activeSheet === 'ALL' && (
                            <span className="text-[11px] text-slate-400">[{item.sheet_name}]</span>
                          )}
                        </div>

                        {/* Linked files & inline notes */}
                        <div className="mt-1.5 flex flex-wrap items-center gap-2">
                          {item.linked_files.map((lf) => (
                            <a
                              key={lf.id}
                              href={`/api/files/${lf.id}/download`}
                              className="inline-flex items-center gap-1 rounded bg-[#EBF2F7] px-2 py-0.5 text-[11px] text-[#3B5B7E] hover:underline"
                            >
                              <Paperclip className="h-3 w-3" />
                              <span className="max-w-[180px] truncate">{lf.original_name}</span>
                            </a>
                          ))}
                          <input
                            defaultValue={item.notes}
                            onBlur={(e) => {
                              if (e.target.value !== item.notes) updateNotes(item, e.target.value);
                            }}
                            placeholder="流转备注（如：已归还 / 缺签字）..."
                            className="w-48 rounded border border-transparent bg-transparent px-1.5 py-0.5 text-xs text-slate-500 hover:border-slate-200 focus:border-[#3B5B7E] focus:bg-white focus:outline-none"
                          />
                        </div>
                      </div>
                    </div>

                    {/* Right Actions */}
                    <div className="flex flex-wrap items-center gap-1.5 shrink-0 pl-8 lg:pl-0">
                      {/* 3-State Toggle */}
                      <div className="inline-flex rounded-lg border border-slate-200 bg-white p-0.5">
                        <button
                          onClick={() => updateStatus(item, 'present')}
                          className={`inline-flex items-center gap-1 rounded-md px-2 py-1 text-[11px] font-medium transition-colors ${
                            item.status === 'present'
                              ? 'bg-[#ECFDF5] text-[#2F6F5E]'
                              : 'text-slate-400 hover:text-slate-700'
                          }`}
                        >
                          <CheckCircle2 className="h-3 w-3" /> 已有
                        </button>
                        <button
                          onClick={() => updateStatus(item, 'missing')}
                          className={`inline-flex items-center gap-1 rounded-md px-2 py-1 text-[11px] font-medium transition-colors ${
                            item.status === 'missing'
                              ? item.is_required
                                ? 'bg-[#FEF2F2] text-[#A94442]'
                                : 'bg-slate-200/80 text-slate-700'
                              : 'text-slate-400 hover:text-slate-700'
                          }`}
                        >
                          <AlertTriangle className="h-3 w-3" /> 缺失
                        </button>
                        <button
                          onClick={() => updateStatus(item, 'na')}
                          className={`inline-flex items-center gap-1 rounded-md px-2 py-1 text-[11px] font-medium transition-colors ${
                            item.status === 'na'
                              ? 'bg-slate-100 text-slate-600'
                              : 'text-slate-400 hover:text-slate-700'
                          }`}
                        >
                          <MinusCircle className="h-3 w-3" /> 无
                        </button>
                      </div>

                      {/* Upload file to this item */}
                      <label
                        className="inline-flex cursor-pointer items-center gap-1 rounded-lg border border-slate-200 bg-white px-2 py-1.5 text-[11px] font-medium text-slate-600 hover:bg-slate-50"
                        title="上传电子文件（上传后自动标为已有）"
                      >
                        <Upload className="h-3 w-3" />
                        <span>上传</span>
                        <input
                          type="file"
                          onChange={(e) => handleItemFileUpload(item, e)}
                          className="hidden"
                        />
                      </label>

                      {/* Cross-module link */}
                      <button
                        onClick={() => setLinkingItem(item)}
                        className="inline-flex items-center gap-1 rounded-lg border border-slate-200 bg-white px-2 py-1.5 text-[11px] font-medium text-slate-600 hover:bg-slate-50"
                        title="引用招标/开评标/定标模块已有文件"
                      >
                        <Link2 className="h-3 w-3" /> 引用
                      </button>

                      {/* Expand details */}
                      <button
                        onClick={() => toggleExpand(item.id)}
                        className="rounded-lg p-1.5 text-slate-400 hover:bg-slate-100 hover:text-slate-600"
                        title="查看纸质/电子/制度要求"
                      >
                        {expanded ? <ChevronUp className="h-3.5 w-3.5" /> : <ChevronDown className="h-3.5 w-3.5" />}
                      </button>

                      <button
                        onClick={() => handleDeleteItem(item.id)}
                        className="rounded-lg p-1.5 text-slate-300 hover:bg-red-50 hover:text-[#A94442]"
                        title="删除条目"
                      >
                        <Trash2 className="h-3.5 w-3.5" />
                      </button>
                    </div>
                  </div>

                  {/* Expanded Requirements */}
                  {expanded && (
                    <div className="mt-2.5 ml-8 grid grid-cols-1 gap-2 rounded-lg border border-slate-200/70 bg-slate-50/80 p-3 text-xs sm:grid-cols-2">
                      <div>
                        <span className="font-medium text-slate-500">纸质要求：</span>
                        <span className="text-slate-700">{item.paper_req || '无特殊要求'}</span>
                      </div>
                      <div>
                        <span className="font-medium text-slate-500">电子要求：</span>
                        <span className="text-slate-700">{item.electronic_req || '无特殊要求'}</span>
                      </div>
                      <div>
                        <span className="font-medium text-slate-500">制度依据：</span>
                        <span className="text-slate-600">{item.policy_basis || '—'}</span>
                      </div>
                      <div>
                        <span className="font-medium text-slate-500">模板来源：</span>
                        <span className="text-slate-600">{item.template_source || '—'}</span>
                      </div>
                    </div>
                  )}
                </div>
              );
            })
          )}
        </div>
        {filteredItems.length > 20 && (
          <div className="overflow-hidden rounded-lg border border-slate-200">
            <PaginationBar
              total={filteredItems.length}
              page={archivePager.page}
              pageSize={archivePager.pageSize}
              onPageChange={archivePager.setPage}
              onPageSizeChange={archivePager.setPageSize}
            />
          </div>
        )}
      </div>

      {/* Import Preview Modal (B-column color detection confirmation) */}
      {importPreview && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4 backdrop-blur-[1px]">
          <div className="flex max-h-[85vh] w-full max-w-4xl flex-col overflow-hidden rounded-xl border border-slate-200 bg-white shadow-xl">
            <div className="flex items-center justify-between border-b border-slate-200 px-5 py-3.5">
              <div>
                <h3 className="text-sm font-semibold text-slate-800">
                  归档目录解析确认 — {importPreview.filename}
                </h3>
                <p className="mt-0.5 text-xs text-slate-500">
                  识别到 {importPreview.sheetNames.length} 个工作表（{importPreview.sheetNames.join('、')}），共{' '}
                  {importPreview.items.length} 条记录；其中通过 B 列颜色自动识别出{' '}
                  <strong className="text-[#9A6B26]">{importPreview.colorDetectedCount}</strong> 项必须文件（可点击标签微调）。
                </p>
              </div>
              <button
                onClick={() => setImportPreview(null)}
                className="rounded-lg p-1 text-slate-400 hover:bg-slate-100"
              >
                <X className="h-4 w-4" />
              </button>
            </div>

            <div className="flex-1 overflow-auto p-4">
              <table className="w-full border-collapse text-left text-xs">
                <thead>
                  <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                    <th className="px-2.5 py-2">工作表</th>
                    <th className="px-2.5 py-2 w-12">序号</th>
                    <th className="px-2.5 py-2">文件名称</th>
                    <th className="px-2.5 py-2 w-24">必须/非必须</th>
                    <th className="px-2.5 py-2 w-24">保存介质</th>
                    <th className="px-2.5 py-2">备注</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {paginate(importPreview.items, importPager.page, importPager.pageSize).map((row, localIdx) => {
                    const realIdx = importPager.pageSize === 0 ? localIdx : (importPager.page - 1) * importPager.pageSize + localIdx;
                    return (
                      <tr key={realIdx} className="hover:bg-slate-50">
                        <td className="px-2.5 py-2 text-slate-400 whitespace-nowrap">{row.sheetName}</td>
                        <td className="px-2.5 py-2 font-mono text-slate-500">{row.seq}</td>
                        <td className="px-2.5 py-2 font-medium text-slate-800">{row.docName}</td>
                        <td className="px-2.5 py-2">
                          <button
                            type="button"
                            onClick={() => {
                              const next = [...importPreview.items];
                              next[realIdx] = { ...next[realIdx], isRequired: !next[realIdx].isRequired };
                              setImportPreview({ ...importPreview, items: next });
                            }}
                            className={`rounded px-2 py-0.5 text-[11px] font-medium ${
                              row.isRequired
                                ? 'bg-[#FEF9EC] text-[#9A6B26]'
                                : 'bg-slate-100 text-slate-500'
                            }`}
                          >
                            {row.isRequired ? '必须' : '非必须'}
                          </button>
                        </td>
                        <td className="px-2.5 py-2 text-slate-500">{row.mediaType}</td>
                        <td className="px-2.5 py-2 text-slate-400">{row.notes || '—'}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            <PaginationBar
              total={importPreview.items.length}
              page={importPager.page}
              pageSize={importPager.pageSize}
              onPageChange={importPager.setPage}
              onPageSizeChange={importPager.setPageSize}
            />

            <div className="flex flex-col gap-3 border-t border-slate-200 bg-slate-50 px-5 py-3.5 sm:flex-row sm:items-center sm:justify-between">
              <div className="flex flex-wrap items-center gap-4 text-xs text-slate-600">
                <label className="inline-flex items-center gap-1.5 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={importReplace}
                    onChange={(e) => setImportReplace(e.target.checked)}
                    className="rounded border-slate-300"
                  />
                  <span>覆盖当前批次已有清单</span>
                </label>
                <label className="inline-flex items-center gap-1.5 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={importSaveTpl}
                    onChange={(e) => setImportSaveTpl(e.target.checked)}
                    className="rounded border-slate-300"
                  />
                  <span>保存为全局常用归档模板</span>
                </label>
              </div>
              <div className="flex justify-end gap-2">
                <button
                  onClick={() => setImportPreview(null)}
                  className="rounded-lg border border-slate-200 bg-white px-4 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"
                >
                  取消
                </button>
                <button
                  onClick={handleCommitImport}
                  disabled={committing}
                  className="rounded-lg bg-[#3B5B7E] px-4 py-1.5 text-xs font-medium text-white hover:bg-[#2E4866] disabled:opacity-50"
                >
                  {committing ? '导入中...' : `确认导入 (${importPreview.items.length} 条)`}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Cross-Module File Link Modal */}
      {linkingItem && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4 backdrop-blur-[1px]">
          <div className="w-full max-w-lg rounded-xl border border-slate-200 bg-white p-5 shadow-xl">
            <div className="flex items-center justify-between mb-3">
              <div>
                <h3 className="text-sm font-semibold text-slate-800">从工作台引用文件</h3>
                <p className="text-xs text-slate-400 mt-0.5">归档条目：{linkingItem.doc_name}</p>
              </div>
              <button
                onClick={() => setLinkingItem(null)}
                className="rounded-lg p-1 text-slate-400 hover:bg-slate-100"
              >
                <X className="h-4 w-4" />
              </button>
            </div>

            <div className="max-h-72 overflow-y-auto divide-y divide-slate-100 border border-slate-200 rounded-lg">
              {availableBatchFiles.length === 0 ? (
                <div className="py-8 text-center text-xs text-slate-400">
                  当前批次尚未上传或生成任何文件
                </div>
              ) : (
                availableBatchFiles.map((bf) => {
                  const checked = linkingItem.linked_file_ids.includes(bf.id);
                  return (
                    <label
                      key={bf.id}
                      className="flex cursor-pointer items-center justify-between gap-3 p-3 hover:bg-slate-50"
                    >
                      <div className="flex items-center gap-2.5 min-w-0">
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => toggleLinkedFile(linkingItem, bf.id)}
                          className="rounded border-slate-300"
                        />
                        <div className="min-w-0">
                          <div className="truncate text-xs font-medium text-slate-800">{bf.original_name}</div>
                          <div className="text-[11px] text-slate-400">
                            来源模块：{bf.module} · 分类：{bf.category}
                          </div>
                        </div>
                      </div>
                    </label>
                  );
                })
              )}
            </div>

            <div className="mt-4 flex justify-end">
              <button
                onClick={() => setLinkingItem(null)}
                className="rounded-lg bg-[#3B5B7E] px-4 py-1.5 text-xs font-medium text-white hover:bg-[#2E4866]"
              >
                完成
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Manual Add Archive Item Modal */}
      {showAddModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4 backdrop-blur-[1px]">
          <div className="w-full max-w-md rounded-xl border border-slate-200 bg-white p-5 shadow-xl">
            <div className="flex items-center justify-between mb-4">
              <h3 className="text-sm font-semibold text-slate-800">新增归档条目</h3>
              <button onClick={() => setShowAddModal(false)} className="rounded-lg p-1 text-slate-400 hover:bg-slate-100">
                <X className="h-4 w-4" />
              </button>
            </div>
            <form onSubmit={handleAddManualItem} className="space-y-3">
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1">归属目录</label>
                  <select
                    value={addSheet}
                    onChange={(e) => setAddSheet(e.target.value)}
                    className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs"
                  >
                    <option value="归档文件">归档文件</option>
                    <option value="重要保管文件">重要保管文件</option>
                  </select>
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1">序号</label>
                  <input
                    value={addSeq}
                    onChange={(e) => setAddSeq(e.target.value)}
                    placeholder="如 20"
                    className="w-full rounded-lg border border-slate-200 px-2.5 py-1.5 text-xs"
                  />
                </div>
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-600 mb-1">文件名称 *</label>
                <input
                  required
                  value={addName}
                  onChange={(e) => setAddName(e.target.value)}
                  placeholder="输入归档或保管文件名称"
                  className="w-full rounded-lg border border-slate-200 px-2.5 py-1.5 text-xs"
                />
              </div>
              <div className="grid grid-cols-2 gap-3 items-center">
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1">保存介质</label>
                  <select
                    value={addMedia}
                    onChange={(e) => setAddMedia(e.target.value)}
                    className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs"
                  >
                    <option value="纸质+电子">纸质+电子</option>
                    <option value="纸质">纸质</option>
                    <option value="电子">电子</option>
                    <option value="纸质或电子">纸质或电子</option>
                  </select>
                </div>
                <label className="inline-flex items-center gap-2 pt-4 text-xs text-slate-700 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={addRequired}
                    onChange={(e) => setAddRequired(e.target.checked)}
                    className="rounded border-slate-300"
                  />
                  <span>标记为必须项</span>
                </label>
              </div>
              <div className="flex justify-end gap-2 pt-2">
                <button
                  type="button"
                  onClick={() => setShowAddModal(false)}
                  className="rounded-lg border border-slate-200 px-4 py-1.5 text-xs font-medium text-slate-600"
                >
                  取消
                </button>
                <button
                  type="submit"
                  className="rounded-lg bg-[#3B5B7E] px-4 py-1.5 text-xs font-medium text-white hover:bg-[#2E4866]"
                >
                  确认添加
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}
