import React, { useCallback, useEffect, useState } from 'react';
import {
  Save,
  Plus,
  Pencil,
  Trash2,
  Upload,
  Eye,
  Download,
  Check,
  CheckCircle2,
  AlertTriangle,
  ShieldAlert,
  Search,
  Wrench,
  Users,
  Layers,
  SlidersHorizontal,
  ChevronDown,
  ChevronUp,
  ArrowUpDown,
  ArrowUp,
  ArrowDown,
} from 'lucide-react';
import { api } from '../api/client';
import { ChangeNoticeBanner } from '../components/ChangeNoticeBanner';
import { DateTimePickerField } from '../components/DateTimePickerField';
import { ExcelPreviewModal } from '../components/ExcelPreviewModal';
import { paginate, PaginationBar, usePagination } from '../components/PaginationBar';
import { ProjectEditModal } from '../components/ProjectEditModal';
import { SearchableClientSelect } from '../components/SearchableClientSelect';
import { useBatch } from '../context/BatchContext';
import type {
  BatchProject,
  BiddingWorkspace,
  ClientAgreement,
  ExtraField,
  ParsedSheetTable,
} from '../types';
import { sortProjects, type ProjectSortKey, type SortDirection } from '../utils/projectSort';
import { formatMoney, getTerminology } from '../utils/terminology';

const BATCH_STATUSES = ['筹备中', '已挂网', '开评标中', '归档中', '已完成'];
const PROCUREMENT_METHODS = ['竞争性谈判', '公开招标', '询价采购', '单一来源', '竞争性磋商'];

type ActiveTab = 'projects' | 'suppliers' | 'batch_info';

export function BiddingPage() {
  const { currentBatchId, refreshBatches } = useBatch();
  const [workspace, setWorkspace] = useState<BiddingWorkspace | null>(null);
  const [activeTab, setActiveTab] = useState<ActiveTab>('projects');
  const [priceUnit, setPriceUnit] = useState<'wan' | 'yuan'>('wan');
  const [tableView, setTableView] = useState<'compact' | 'full'>('compact');
  const [expandedProjectIds, setExpandedProjectIds] = useState<Set<string>>(new Set());
  const [selectedProjectIds, setSelectedProjectIds] = useState<Set<string>>(new Set());
  const [projectSearch, setProjectSearch] = useState('');
  const [onlyChangedProjects, setOnlyChangedProjects] = useState(false);
  const [editingProject, setEditingProject] = useState<BatchProject | null>(null);
  const [sortKey, setSortKey] = useState<ProjectSortKey>('lot_package');
  const [sortDir, setSortDir] = useState<SortDirection>('asc');
  const [projPage, setProjPage] = useState(1);
  const [projPageSize, setProjPageSize] = useState(20);
  const [importingProjects, setImportingProjects] = useState(false);
  const [supplierSearch, setSupplierSearch] = useState('');
  const [showConflictDetails, setShowConflictDetails] = useState(false);
  const [conflictFilter, setConflictFilter] = useState<'all' | 'erp_extra' | 'req_missing' | 'field_conflict'>('all');
  const conflictPager = usePagination(`${currentBatchId}-${conflictFilter}`);
  const supplierPager = usePagination(`${currentBatchId}-${supplierSearch}`);

  // Batch editable fields
  const [clients, setClients] = useState<ClientAgreement[]>([]);
  const [name, setName] = useState('');
  const [code, setCode] = useState('');
  const [projectUnit, setProjectUnit] = useState('');
  const [procurementType, setProcurementType] = useState<'公开招标' | '非公开招标'>('非公开招标');
  const [unitNature, setUnitNature] = useState<'主业' | '产业'>('产业');
  const [announcementTime, setAnnouncementTime] = useState('');
  const [docFetchTime, setDocFetchTime] = useState('');
  const [bidOpeningTime, setBidOpeningTime] = useState('');
  const [contactPhone, setContactPhone] = useState('');
  const [contactEmail, setContactEmail] = useState('');
  const [status, setStatus] = useState('筹备中');
  const [notes, setNotes] = useState('');
  const [extraFields, setExtraFields] = useState<ExtraField[]>([]);

  const [saving, setSaving] = useState(false);
  const [savedToast, setSavedToast] = useState(false);

  // Slot upload state + Cleaning Report banner
  const [uploadingSlot, setUploadingSlot] = useState<string | null>(null);
  const [lastReport, setLastReport] = useState<{
    slotLabel: string;
    summary: string;
    warnings: string[];
  } | null>(null);
  const [previewData, setPreviewData] = useState<{ title: string; sheets: ParsedSheetTable[] } | null>(null);

  // Manual project add state
  const [newProjName, setNewProjName] = useState('');
  const [newProjCode, setNewProjCode] = useState('');
  const [newProjBudget, setNewProjBudget] = useState('');
  const [newProjMethod, setNewProjMethod] = useState('竞争性谈判');
  const [newProjDept, setNewProjDept] = useState('');

  const loadData = useCallback(async () => {
    if (!currentBatchId) return;
    try {
      const ws = await api.getBiddingWorkspace(currentBatchId);
      setWorkspace(ws);
      const b = ws.batch;
      setName(b.name || '');
      setCode(b.code || '');
      setProjectUnit(b.project_unit || '');
      setProcurementType(b.procurement_type === '公开招标' ? '公开招标' : '非公开招标');
      setUnitNature(b.unit_nature === '主业' ? '主业' : '产业');
      setAnnouncementTime(b.announcement_time || '');
      setDocFetchTime(b.doc_fetch_time || '');
      setBidOpeningTime(b.bid_opening_time || '');
      setContactPhone(b.contact_phone || '');
      setContactEmail(b.contact_email || '');
      setStatus(b.status || '筹备中');
      setNotes(b.notes || '');
      setExtraFields(b.extra_fields || []);
    } catch {
      // ignore if offline
    }
  }, [currentBatchId]);

  useEffect(() => {
    api.listClients().then(setClients).catch(() => {});
  }, []);

  useEffect(() => {
    loadData();
    setLastReport(null);
    setSelectedProjectIds(new Set());
    setProjPage(1);
  }, [loadData]);

  useEffect(() => {
    setProjPage(1);
  }, [projectSearch, sortKey, sortDir]);

  const terms = getTerminology(procurementType, unitNature);

  function toggleExpandProject(id: string) {
    setExpandedProjectIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  async function handleSaveBatch(e: React.FormEvent) {
    e.preventDefault();
    if (!currentBatchId) return;
    setSaving(true);
    try {
      await api.updateBatch(currentBatchId, {
        name,
        code,
        project_unit: projectUnit,
        procurement_type: procurementType,
        unit_nature: unitNature,
        announcement_time: announcementTime,
        doc_fetch_time: docFetchTime,
        bid_opening_time: bidOpeningTime,
        contact_phone: contactPhone,
        contact_email: contactEmail,
        status,
        notes,
        extra_fields: extraFields.filter((f) => f.key.trim() !== ''),
      });
      await loadData();
      await refreshBatches();
      setSavedToast(true);
      setTimeout(() => setSavedToast(false), 2000);
    } finally {
      setSaving(false);
    }
  }

  async function handleSlotUpload(slotKey: string, e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file || !currentBatchId) return;
    setUploadingSlot(slotKey);
    try {
      const res = await api.uploadBiddingSlot(currentBatchId, slotKey, file);
      setLastReport(res.report);
      await loadData();
      await refreshBatches();
      if (slotKey === 'supplier_contacts') setActiveTab('suppliers');
      else if (slotKey === 'variable_template') setActiveTab('batch_info');
      else setActiveTab('projects');
    } catch (err) {
      alert((err as Error).message);
    } finally {
      setUploadingSlot(null);
      e.target.value = '';
    }
  }

  async function handleClearSlot(slotKey: string, slotLabel: string) {
    if (!currentBatchId) return;
    if (!window.confirm(`确认清除【${slotLabel}】已导入的全部解析数据与原表文件吗？`)) return;
    setUploadingSlot(slotKey);
    try {
      await api.clearBiddingSlot(currentBatchId, slotKey);
      setLastReport({
        slotLabel,
        summary: `已清除【${slotLabel}】的全部解析数据与原表记录，可随时重新上传。`,
        warnings: [],
      });
      await loadData();
      await refreshBatches();
    } catch (err) {
      alert((err as Error).message);
    } finally {
      setUploadingSlot(null);
    }
  }

  async function handleFixWeight(projectId: string, suggested: string) {
    await api.fixProjectWeight(projectId, suggested);
    await loadData();
  }

  async function handleImportProjects(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file || !currentBatchId) return;
    setImportingProjects(true);
    try {
      const res = await api.importProjectsFromExcel(currentBatchId, file);
      setLastReport(res.report);
      await loadData();
      await refreshBatches();
    } catch (err) {
      alert((err as Error).message);
    } finally {
      setImportingProjects(false);
      e.target.value = '';
    }
  }

  async function handleAddProject(e: React.FormEvent) {
    e.preventDefault();
    if (!currentBatchId || !newProjName.trim()) return;
    await api.createProject(currentBatchId, {
      name: newProjName.trim(),
      code: newProjCode.trim(),
      budget: newProjBudget.trim(),
      procurement_method: newProjMethod,
      department: newProjDept.trim(),
    });
    setNewProjName('');
    setNewProjCode('');
    setNewProjBudget('');
    setNewProjDept('');
    await loadData();
    await refreshBatches();
  }

  async function handleDeleteProject(projectId: string) {
    await api.deleteProject(projectId);
    setSelectedProjectIds((prev) => {
      const next = new Set(prev);
      next.delete(projectId);
      return next;
    });
    await loadData();
    await refreshBatches();
  }

  function toggleSelectProject(id: string) {
    setSelectedProjectIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  function toggleSelectAllVisible() {
    const visibleIds = sortedProjects.map((p) => p.id);
    const allSelected = visibleIds.length > 0 && visibleIds.every((id) => selectedProjectIds.has(id));
    if (allSelected) {
      setSelectedProjectIds((prev) => {
        const next = new Set(prev);
        visibleIds.forEach((id) => next.delete(id));
        return next;
      });
    } else {
      setSelectedProjectIds((prev) => {
        const next = new Set(prev);
        visibleIds.forEach((id) => next.add(id));
        return next;
      });
    }
  }

  async function handleBatchDeleteProjects() {
    const ids = Array.from(selectedProjectIds);
    if (ids.length === 0) return;
    if (!window.confirm(`确认批量删除已勾选的 ${ids.length} 个项目吗？删除后无法恢复。`)) return;
    await api.batchDeleteProjects(ids);
    setSelectedProjectIds(new Set());
    await loadData();
    await refreshBatches();
  }

  const projects = workspace?.projects || [];
  const suppliers = workspace?.suppliers || [];
  const collusionRisks = workspace?.collusionRisks || [];
  const erpConflicts = workspace?.erpConflicts || [];
  const changeNotices = workspace?.changeNotices || [];
  const slots = workspace?.slots || [];

  const projectChangeMap = new Map<string, typeof changeNotices>();
  changeNotices.forEach((cn) => {
    cn.project_ids.forEach((pid) => {
      const list = projectChangeMap.get(pid) || [];
      if (!list.some((x) => x.id === cn.id)) list.push(cn);
      projectChangeMap.set(pid, list);
    });
  });

  const weightAnomalies = projects.filter((p) => p.weight_warning);
  const filteredProjects = projects.filter((p) => {
    if (onlyChangedProjects && !projectChangeMap.has(p.id)) return false;
    if (!projectSearch.trim()) return true;
    const tokens = projectSearch.trim().toLowerCase().split(/\s+/);
    const haystack = [
      p.lot_name,
      p.req_app_no,
      p.name,
      p.lot_code,
      p.package_no,
      p.project_code,
      p.tech_spec_code,
      p.project_overview,
      p.project_unit,
      p.department,
      p.declarant_name,
      p.declarant_phone,
      p.biz_eval_tpl,
      p.tech_eval_tpl,
      p.contract_version,
      p.quote_method,
      p.funding_source,
      p.qualification_req,
      p.performance_req,
      p.staff_req,
    ]
      .filter(Boolean)
      .join(' ')
      .toLowerCase();
    return tokens.every((tk) => haystack.includes(tk));
  });
  const sortedProjects = sortProjects(filteredProjects, sortKey, sortDir);
  const pagedProjects = paginate(sortedProjects, projPage, projPageSize);
  const pageOffset = projPageSize === 0 ? 0 : (Math.min(projPage, Math.max(1, Math.ceil(sortedProjects.length / projPageSize))) - 1) * projPageSize;

  function handleSort(key: ProjectSortKey) {
    if (sortKey === key) {
      setSortDir((d) => (d === 'asc' ? 'desc' : 'asc'));
    } else {
      setSortKey(key);
      setSortDir('asc');
    }
  }

  function renderSortBtn(label: string, key: ProjectSortKey) {
    const active = sortKey === key;
    return (
      <button
        type="button"
        onClick={() => handleSort(key)}
        className={`inline-flex items-center gap-1 font-medium transition-colors hover:text-[#3B5B7E] ${
          active ? 'text-[#3B5B7E]' : 'text-slate-500'
        }`}
      >
        <span>{label}</span>
        {active ? (
          sortDir === 'asc' ? <ArrowUp className="h-3 w-3" /> : <ArrowDown className="h-3 w-3" />
        ) : (
          <ArrowUpDown className="h-3 w-3 opacity-40" />
        )}
      </button>
    );
  }

  const filteredSuppliers = suppliers.filter((s) => {
    if (!supplierSearch.trim()) return true;
    const q = supplierSearch.trim().toLowerCase();
    return (
      s.supplier_name.toLowerCase().includes(q) ||
      s.contact_name.toLowerCase().includes(q) ||
      s.mobile.includes(q) ||
      s.tax_no.toLowerCase().includes(q)
    );
  });

  // 总预算严格优先按「含税预算金额」汇总（兼容折扣率报价行最高限价为 100% 的情况），仅当含税预算缺失时回退到数值型最高限价
  const totalBudgetYuan = projects.reduce((sum, p) => {
    const b = p.budget_yuan !== null && p.budget_yuan !== undefined ? Number(p.budget_yuan) : null;
    if (b !== null && Number.isFinite(b) && b > 0) return sum + b;
    const m = p.max_price_yuan !== null && p.max_price_yuan !== undefined ? Number(p.max_price_yuan) : null;
    return sum + (m !== null && Number.isFinite(m) && m > 0 ? m : 0);
  }, 0);
  const discountRateCount = projects.filter((p) => p.quote_method?.includes('折扣') || Boolean(p.max_price_raw)).length;
  const mergedSubPkgCount = projects.filter((p) => p.subcontract_rule?.includes('合并') || p.name?.includes('子包')).length;

  return (
    <div className="space-y-5">
      {/* Header with Live Terminology & Platform Badges */}
      <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <div className="flex flex-wrap items-center gap-2">
            <h1 className="text-lg font-semibold text-slate-800">1. 招标模块</h1>
            <span className="rounded-md bg-[#EBF2F7] px-2 py-0.5 text-xs font-medium text-[#3B5B7E]">
              {terms.procurementType === '公开招标' ? '公开招标' : '非公开招标（竞争性谈判）'}
            </span>
            <span className="rounded-md bg-slate-100 px-2 py-0.5 text-xs font-medium text-slate-600">
              {terms.platformFull}
            </span>
          </div>
          <p className="mt-1 text-xs text-slate-500">
            按 SOP 上传 5 类前期表格，系统以「采购申请号 → 技术规范书编码 → 项目编码 → 项目名称」四级匹配融合全量字段，并自动清洗空格、对齐元/万元及校验权重比例。
          </p>
        </div>
      </div>

      {/* Data Cleaning & Validation Report Banner */}
      {lastReport && (
        <div
          className={`rounded-xl border p-4 space-y-2 ${
            lastReport.warnings.length > 0
              ? 'border-[#9A6B26]/30 bg-[#FEF9EC]/70'
              : 'border-[#2F6F5E]/30 bg-[#ECFDF5]/70'
          }`}
        >
          <div className="flex items-start justify-between gap-2">
            <div className="flex items-center gap-2 text-xs font-semibold text-slate-800">
              {lastReport.warnings.length > 0 ? (
                <AlertTriangle className="h-4 w-4 text-[#9A6B26]" />
              ) : (
                <CheckCircle2 className="h-4 w-4 text-[#2F6F5E]" />
              )}
              <span>【{lastReport.slotLabel}】解析与数据清洗完成</span>
            </div>
            <button
              onClick={() => setLastReport(null)}
              className="text-xs text-slate-400 hover:text-slate-600"
            >
              关闭
            </button>
          </div>
          <p className="text-xs text-slate-600">{lastReport.summary}</p>
          {lastReport.warnings.length > 0 && (
            <ul className="space-y-1 border-t border-[#9A6B26]/15 pt-2 text-xs text-[#9A6B26]">
              {lastReport.warnings.map((w, i) => (
                <li key={i}>• {w}</li>
              ))}
            </ul>
          )}
        </div>
      )}

      {/* Change Notice Reminder Banner & Registration */}
      {currentBatchId && (
        <ChangeNoticeBanner
          batchId={currentBatchId}
          notices={changeNotices}
          projects={projects}
          canManage
          onlyChanged={onlyChangedProjects}
          onToggleOnlyChanged={() => setOnlyChangedProjects((v) => !v)}
          onChanged={async () => {
            await loadData();
            await refreshBatches();
          }}
        />
      )}

      {/* 5-Slot SOP Upload Matrix */}
      <div className="rounded-xl border border-slate-200 bg-white p-4 space-y-3">
        <div className="flex items-center justify-between">
          <h2 className="text-sm font-semibold text-slate-700">招标阶段标准文件槽位（按 SOP 流程顺序上传）</h2>
          <span className="text-xs text-slate-400 tabular-nums">
            已就绪 {slots.filter((s) => s.file).length} / {slots.length} 份
          </span>
        </div>

        <div className="grid grid-cols-1 gap-3 md:grid-cols-2 lg:grid-cols-5">
          {slots.map((slot) => {
            const isUploading = uploadingSlot === slot.key;
            const hasFile = Boolean(slot.file);

            return (
              <div
                key={slot.key}
                className={`flex flex-col justify-between rounded-lg border p-3 transition-colors ${
                  hasFile
                    ? 'border-[#2F6F5E]/30 bg-[#ECFDF5]/25'
                    : 'border-slate-200 bg-slate-50/40'
                }`}
              >
                <div>
                  <div className="flex items-center justify-between gap-1">
                    <span className="inline-flex h-5 w-5 items-center justify-center rounded-full bg-[#3B5B7E]/10 text-[11px] font-semibold text-[#3B5B7E] tabular-nums">
                      {slot.order}
                    </span>
                    <span className="rounded bg-white px-1.5 py-0.5 text-[10px] font-medium text-slate-500 border border-slate-200/70">
                      {slot.stage}
                    </span>
                  </div>
                  <div className="mt-2 text-xs font-semibold text-slate-800">{slot.label}</div>
                  <p className="mt-1 text-[11px] text-slate-500 leading-relaxed line-clamp-2" title={slot.description}>
                    {slot.description}
                  </p>
                </div>

                <div className="mt-3 pt-2 border-t border-slate-200/60 space-y-2">
                  {slot.file ? (
                    <div className="space-y-1">
                      <div className="flex items-center justify-between gap-1 text-[11px] text-[#2F6F5E]">
                        <span className="truncate font-medium" title={slot.file.original_name}>
                          {slot.file.original_name}
                        </span>
                        <div className="flex items-center gap-0.5 shrink-0">
                          {slot.file.parsed && (
                            <button
                              onClick={() =>
                                setPreviewData({ title: slot.file!.original_name, sheets: slot.file!.parsed!.sheets })
                              }
                              className="rounded p-1 text-slate-400 hover:bg-white hover:text-[#3B5B7E]"
                              title="预览原表"
                            >
                              <Eye className="h-3.5 w-3.5" />
                            </button>
                          )}
                          {slot.file.stored_path && (
                            <a
                              href={`/api/files/${slot.file.id}/download`}
                              className="rounded p-1 text-slate-400 hover:bg-white hover:text-slate-700"
                              title="下载原表"
                            >
                              <Download className="h-3.5 w-3.5" />
                            </a>
                          )}
                        </div>
                      </div>
                      {slot.key === 'erp_plan' && (slot.files?.length || 0) > 1 && (
                        <div className="text-[10px] text-[#3B5B7E] font-medium">
                          已合并 {slot.files!.length} 份 ERP 计划（支持主业+产业联合上传）
                        </div>
                      )}
                    </div>
                  ) : (
                    <div className="text-[11px] text-slate-400">
                      {slot.key === 'erp_plan' ? '支持主业/产业 ERP（可分次上传）' : '待上传 Excel'}
                    </div>
                  )}

                  <div className="flex items-center gap-1.5">
                    <label className="flex flex-1 cursor-pointer items-center justify-center gap-1 rounded-md border border-slate-200 bg-white py-1.5 text-[11px] font-medium text-slate-700 hover:border-[#3B5B7E] hover:text-[#3B5B7E] transition-colors">
                      <Upload className="h-3 w-3" />
                      <span>
                        {isUploading
                          ? '处理中...'
                          : hasFile
                            ? slot.key === 'erp_plan'
                              ? '追加/重传'
                              : '重新上传'
                            : '上传并自动提取'}
                      </span>
                      <input
                        type="file"
                        accept=".xlsx,.xls,.csv"
                        onChange={(e) => handleSlotUpload(slot.key, e)}
                        className="hidden"
                        disabled={isUploading}
                      />
                    </label>
                    {hasFile && (
                      <button
                        type="button"
                        onClick={() => handleClearSlot(slot.key, slot.label)}
                        disabled={isUploading}
                        className="inline-flex shrink-0 items-center gap-1 rounded-md border border-[#A94442]/25 bg-[#FEF2F2]/70 px-2 py-1.5 text-[11px] font-medium text-[#A94442] hover:bg-[#FEE2E2] disabled:opacity-50 transition-colors"
                        title="清除该槽位已导入的数据"
                      >
                        <Trash2 className="h-3 w-3" />
                        <span>清除</span>
                      </button>
                    )}
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      </div>

      {/* Main Workspace Tabs */}
      <div className="rounded-xl border border-slate-200 bg-white overflow-hidden">
        <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 bg-slate-50/70 px-4 py-2.5">
          <div className="flex flex-wrap gap-1">
            <button
              onClick={() => setActiveTab('projects')}
              className={`inline-flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-xs font-medium transition-colors ${
                activeTab === 'projects'
                  ? 'bg-[#3B5B7E] text-white'
                  : 'text-slate-600 hover:bg-slate-200/60'
              }`}
            >
              <Layers className="h-3.5 w-3.5" />
              <span>标段/项目主数据 ({projects.length})</span>
              {weightAnomalies.length > 0 && (
                <span className="rounded-full bg-[#A94442] px-1.5 py-0.2 text-[10px] text-white">
                  {weightAnomalies.length} 异常
                </span>
              )}
            </button>

            <button
              onClick={() => setActiveTab('suppliers')}
              className={`inline-flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-xs font-medium transition-colors ${
                activeTab === 'suppliers'
                  ? 'bg-[#3B5B7E] text-white'
                  : 'text-slate-600 hover:bg-slate-200/60'
              }`}
            >
              <Users className="h-3.5 w-3.5" />
              <span>报名{terms.bidder}与围串标初筛 ({suppliers.length})</span>
              {collusionRisks.length > 0 && (
                <span className="rounded-full bg-[#9A6B26] px-1.5 py-0.2 text-[10px] text-white">
                  {collusionRisks.length} 提醒
                </span>
              )}
            </button>

            <button
              onClick={() => setActiveTab('batch_info')}
              className={`inline-flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-xs font-medium transition-colors ${
                activeTab === 'batch_info'
                  ? 'bg-[#3B5B7E] text-white'
                  : 'text-slate-600 hover:bg-slate-200/60'
              }`}
            >
              <SlidersHorizontal className="h-3.5 w-3.5" />
              <span>批次基本信息与时间节点</span>
            </button>
          </div>

          {activeTab === 'projects' && (
            <div className="flex flex-wrap items-center gap-2">
              {totalBudgetYuan > 0 && (
                <span className="text-xs text-slate-500 tabular-nums" title="按各项目「含税预算金额」列合计，含折扣率报价项目">
                  含税预算合计：<strong className="text-slate-800">{formatMoney(totalBudgetYuan, priceUnit)}</strong>
                </span>
              )}
              {discountRateCount > 0 && (
                <span className="rounded bg-[#EBF2F7] px-1.5 py-0.5 text-[11px] text-[#3B5B7E] tabular-nums">
                  折扣率 {discountRateCount} 项
                </span>
              )}
              {mergedSubPkgCount > 0 && (
                <span className="rounded bg-[#FEF9EC] px-1.5 py-0.5 text-[11px] text-[#9A6B26] tabular-nums">
                  合并成包 {mergedSubPkgCount} 条
                </span>
              )}
              {/* Compact vs Full-Field View Toggle */}
              <div className="inline-flex rounded-lg border border-slate-200 bg-white p-0.5 text-[11px]">
                <button
                  onClick={() => setTableView('compact')}
                  className={`rounded-md px-2 py-0.5 font-medium ${
                    tableView === 'compact' ? 'bg-slate-800 text-white' : 'text-slate-500'
                  }`}
                >
                  精简视图
                </button>
                <button
                  onClick={() => setTableView('full')}
                  className={`rounded-md px-2 py-0.5 font-medium ${
                    tableView === 'full' ? 'bg-slate-800 text-white' : 'text-slate-500'
                  }`}
                >
                  全字段视图(23列)
                </button>
              </div>
              {/* Price Unit Toggle */}
              <div className="inline-flex rounded-lg border border-slate-200 bg-white p-0.5 text-[11px]">
                <button
                  onClick={() => setPriceUnit('wan')}
                  className={`rounded-md px-2 py-0.5 font-medium ${
                    priceUnit === 'wan' ? 'bg-[#3B5B7E] text-white' : 'text-slate-500'
                  }`}
                >
                  万元
                </button>
                <button
                  onClick={() => setPriceUnit('yuan')}
                  className={`rounded-md px-2 py-0.5 font-medium ${
                    priceUnit === 'yuan' ? 'bg-[#3B5B7E] text-white' : 'text-slate-500'
                  }`}
                >
                  元
                </button>
              </div>
            </div>
          )}
        </div>

        {/* TAB 1: Merged Master Projects Table */}
        {activeTab === 'projects' && (
          <div className="p-4 space-y-4">
            {/* Search & Incremental Excel Import Toolbar */}
            <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
              <div className="relative flex-1 max-w-md">
                <Search className="pointer-events-none absolute left-2.5 top-2 h-3.5 w-3.5 text-slate-400" />
                <input
                  value={projectSearch}
                  onChange={(e) => setProjectSearch(e.target.value)}
                  placeholder="搜索分标名称、采购申请号、项目名称或任意字段（空格支持多词）..."
                  className="w-full rounded-lg border border-slate-200 py-1.5 pl-8 pr-14 text-xs focus:border-[#3B5B7E] focus:outline-none"
                />
                {projectSearch && (
                  <button
                    type="button"
                    onClick={() => setProjectSearch('')}
                    className="absolute right-2 top-1.5 rounded px-1.5 py-0.5 text-[10px] text-slate-400 hover:bg-slate-100 hover:text-slate-600"
                  >
                    清空
                  </button>
                )}
              </div>
              <div className="flex flex-wrap items-center gap-2">
                {projectSearch && (
                  <span className="text-xs text-slate-500 tabular-nums">
                    匹配 {filteredProjects.length} / {projects.length} 条
                  </span>
                )}
                {selectedProjectIds.size > 0 && (
                  <button
                    type="button"
                    onClick={handleBatchDeleteProjects}
                    className="inline-flex shrink-0 items-center gap-1.5 rounded-lg border border-[#A94442]/30 bg-[#FEF2F2] px-3 py-1.5 text-xs font-medium text-[#A94442] hover:bg-[#FEE2E2] transition-colors"
                  >
                    <Trash2 className="h-3.5 w-3.5" />
                    <span>批量删除 ({selectedProjectIds.size})</span>
                  </button>
                )}
                <label className="inline-flex shrink-0 cursor-pointer items-center gap-1.5 rounded-lg border border-[#3B5B7E]/30 bg-[#EBF2F7]/50 px-3 py-1.5 text-xs font-medium text-[#3B5B7E] hover:bg-[#EBF2F7] transition-colors">
                  <Upload className="h-3.5 w-3.5" />
                  <span>{importingProjects ? '增量去重导入中...' : '表格增量导入项目（自动防重）'}</span>
                  <input
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    onChange={handleImportProjects}
                    className="hidden"
                    disabled={importingProjects}
                  />
                </label>
              </div>
            </div>

            {erpConflicts.length > 0 && (() => {
              const priority: Record<string, number> = { req_missing: 0, erp_extra: 1, field_conflict: 2 };
              const sortedConflicts = [...erpConflicts].sort((a, b) => (priority[a.type] ?? 9) - (priority[b.type] ?? 9));
              const extraList = sortedConflicts.filter((c) => c.type === 'erp_extra');
              const missingList = sortedConflicts.filter((c) => c.type === 'req_missing');
              const fieldList = sortedConflicts.filter((c) => c.type === 'field_conflict');
              const visibleConflicts = conflictFilter === 'all'
                ? sortedConflicts
                : sortedConflicts.filter((c) => c.type === conflictFilter);
              return (
                <div className="rounded-lg border border-[#9A6B26]/30 bg-[#FEF9EC]/70 p-3 text-xs space-y-2">
                  <div className="flex flex-wrap items-center justify-between gap-2">
                    <div className="flex items-center gap-2 text-[#9A6B26] font-semibold">
                      <AlertTriangle className="h-4 w-4 shrink-0" />
                      <span>
                        已以《需求一览表》为准展示 · 检测到 {erpConflicts.length} 项与 ERP 采购计划的差异
                      </span>
                    </div>
                    <div className="flex items-center gap-1.5">
                      <a
                        href={`/api/bidding/batch/${currentBatchId}/erp-conflicts/export`}
                        className="inline-flex items-center gap-1 rounded-md border border-[#3B5B7E]/30 bg-white px-2.5 py-1 text-[11px] font-medium text-[#3B5B7E] hover:bg-[#EBF2F7]/60"
                        title="导出全部差异明细为 Excel（含数量差异与字段冲突分表）"
                      >
                        <Download className="h-3 w-3" />
                        <span>导出全部 (.xlsx)</span>
                      </a>
                      <button
                        type="button"
                        onClick={() => setShowConflictDetails((v) => !v)}
                        className="inline-flex items-center gap-1 rounded-md border border-[#9A6B26]/30 bg-white px-2.5 py-1 text-[11px] font-medium text-[#9A6B26] hover:bg-[#FEF9EC]"
                      >
                        <span>{showConflictDetails ? '收起差异明细' : '展开差异明细'}</span>
                        {showConflictDetails ? <ChevronUp className="h-3 w-3" /> : <ChevronDown className="h-3 w-3" />}
                      </button>
                    </div>
                  </div>
                  <div className="flex flex-wrap items-center gap-2 text-[11px] text-slate-600 tabular-nums">
                    <button
                      type="button"
                      onClick={() => { setConflictFilter('all'); setShowConflictDetails(true); }}
                      className={`rounded px-2 py-0.5 border transition-colors ${conflictFilter === 'all' ? 'bg-slate-800 text-white border-slate-800' : 'bg-white border-slate-200 hover:bg-slate-50'}`}
                    >
                      全部 ({sortedConflicts.length})
                    </button>
                    {extraList.length > 0 && (
                      <button
                        type="button"
                        onClick={() => { setConflictFilter('erp_extra'); setShowConflictDetails(true); }}
                        className={`rounded px-2 py-0.5 border transition-colors ${conflictFilter === 'erp_extra' ? 'bg-[#9A6B26] text-white border-[#9A6B26]' : 'bg-white border-[#9A6B26]/25 hover:bg-[#FEF9EC]'}`}
                      >
                        ERP 多出（已过滤）：<strong>{extraList.length}</strong> 条
                      </button>
                    )}
                    {missingList.length > 0 && (
                      <button
                        type="button"
                        onClick={() => { setConflictFilter('req_missing'); setShowConflictDetails(true); }}
                        className={`rounded px-2 py-0.5 border transition-colors ${conflictFilter === 'req_missing' ? 'bg-[#A94442] text-white border-[#A94442]' : 'bg-white border-[#A94442]/25 hover:bg-[#FEF2F2]'}`}
                      >
                        需求一览表缺失 ERP 匹配：<strong>{missingList.length}</strong> 条
                      </button>
                    )}
                    {fieldList.length > 0 && (
                      <button
                        type="button"
                        onClick={() => { setConflictFilter('field_conflict'); setShowConflictDetails(true); }}
                        className={`rounded px-2 py-0.5 border transition-colors ${conflictFilter === 'field_conflict' ? 'bg-[#9A6B26] text-white border-[#9A6B26]' : 'bg-white border-[#9A6B26]/25 hover:bg-[#FEF9EC]'}`}
                      >
                        关键字段不一致：<strong>{fieldList.length}</strong> 条
                      </button>
                    )}
                  </div>
                  {showConflictDetails && (
                    <div className="overflow-hidden rounded-md border border-[#9A6B26]/20 bg-white">
                      <div className="overflow-x-auto">
                        <table className="w-full border-collapse text-left text-[11px]">
                          <thead className="bg-slate-50 border-b border-slate-200 text-slate-500">
                            <tr>
                              <th className="px-2.5 py-1.5 font-medium whitespace-nowrap">类型</th>
                              <th className="px-2.5 py-1.5 font-medium whitespace-nowrap">采购申请号</th>
                              <th className="px-2.5 py-1.5 font-medium whitespace-nowrap">分标/单位</th>
                              <th className="px-2.5 py-1.5 font-medium">项目名称</th>
                              <th className="px-2.5 py-1.5 font-medium">冲突明细（当前已以需求一览表为准）</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100">
                            {paginate(visibleConflicts, conflictPager.page, conflictPager.pageSize).map((c, idx) => (
                              <tr key={`${c.type}-${c.req_app_no}-${idx}`} className="hover:bg-slate-50/70">
                                <td className="px-2.5 py-1.5 whitespace-nowrap">
                                  {c.type === 'erp_extra' ? (
                                    <span className="rounded bg-slate-100 px-1.5 py-0.5 text-[10px] font-medium text-slate-600">ERP多出</span>
                                  ) : c.type === 'req_missing' ? (
                                    <span className="rounded bg-[#FEF2F2] px-1.5 py-0.5 text-[10px] font-medium text-[#A94442]">ERP缺失</span>
                                  ) : (
                                    <span className="rounded bg-[#FEF9EC] px-1.5 py-0.5 text-[10px] font-medium text-[#9A6B26]">字段冲突</span>
                                  )}
                                </td>
                                <td className="px-2.5 py-1.5 font-mono text-slate-600 whitespace-nowrap">{c.req_app_no}</td>
                                <td className="px-2.5 py-1.5 text-slate-500 whitespace-nowrap">{c.lot_info || '—'}</td>
                                <td className="px-2.5 py-1.5 text-slate-800 max-w-[200px] truncate" title={c.project_name}>{c.project_name || '—'}</td>
                                <td className="px-2.5 py-1.5 text-slate-600">
                                  {c.fields && c.fields.length > 0 ? (
                                    <div className="space-y-0.5">
                                      {c.fields.map((f, fi) => (
                                        <div key={fi}>
                                          <strong className="text-slate-700">{f.field}：</strong>
                                          <span className="text-[#2F6F5E]">需求「{f.reqVal}」</span>
                                          <span className="mx-1 text-slate-400">vs</span>
                                          <span className="text-[#9A6B26]">ERP「{f.erpVal}」</span>
                                        </div>
                                      ))}
                                    </div>
                                  ) : (
                                    <span>{c.summary}</span>
                                  )}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      <PaginationBar
                        total={visibleConflicts.length}
                        page={conflictPager.page}
                        pageSize={conflictPager.pageSize}
                        onPageChange={conflictPager.setPage}
                        onPageSizeChange={conflictPager.setPageSize}
                      />
                    </div>
                  )}
                </div>
              );
            })()}

            {weightAnomalies.length > 0 && (
              <div className="rounded-lg border border-[#A94442]/25 bg-[#FEF2F2]/60 p-3 text-xs text-[#A94442] flex items-start gap-2">
                <AlertTriangle className="h-4 w-4 shrink-0 mt-0.5" />
                <div>
                  检测到 <strong>{weightAnomalies.length}</strong> 条标段的「技:商:价权重比例」之和不等于 100（常见原因为 Excel 将 <code className="font-mono">60:10:30</code> 误转为时间小数/24小时制截断），请点击下方「一键修正」修复。
                </div>
              </div>
            )}

            <div className="overflow-x-auto">
              {tableView === 'compact' ? (
                <table className="w-full border-collapse text-left text-xs">
                  <thead>
                    <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                      <th className="px-2 py-2.5 w-8">
                        <input
                          type="checkbox"
                          checked={sortedProjects.length > 0 && sortedProjects.every((p) => selectedProjectIds.has(p.id))}
                          onChange={toggleSelectAllVisible}
                          className="rounded border-slate-300"
                          title="全选/取消全选当前列表"
                        />
                      </th>
                      <th className="px-1.5 py-2.5 font-medium w-7"></th>
                      <th className="px-2 py-2.5 font-medium w-10">{renderSortBtn('#', 'seq')}</th>
                      <th className="px-2.5 py-2.5 font-medium whitespace-nowrap">{renderSortBtn('采购申请号（主键）', 'req_app_no')}</th>
                      <th className="px-2.5 py-2.5 font-medium whitespace-nowrap">{renderSortBtn('分标 / 包号', 'lot_package')}</th>
                      <th className="px-2.5 py-2.5 font-medium min-w-[230px]">{renderSortBtn('项目名称与项目概况', 'name')}</th>
                      <th className="px-2.5 py-2.5 font-medium whitespace-nowrap">{renderSortBtn(`最高限价 (${priceUnit === 'wan' ? '万元' : '元'} / %)`, 'max_price_yuan')}</th>
                      <th className="px-2.5 py-2.5 font-medium whitespace-nowrap">{renderSortBtn('权重 (技:商:价)', 'weight_ratio')}</th>
                      <th className="px-2.5 py-2.5 font-medium whitespace-nowrap">{renderSortBtn('详评模板', 'biz_eval_tpl')}</th>
                      <th className="px-2 py-2.5 font-medium w-10 text-right">操作</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {sortedProjects.length === 0 ? (
                      <tr>
                        <td colSpan={10} className="py-8 text-center text-xs text-slate-400">
                          {projectSearch ? '未找到匹配的项目记录，请尝试其他关键词' : '暂无项目数据，可上传表格或手动添加'}
                        </td>
                      </tr>
                    ) : pagedProjects.map((p, idx) => {
                      const expanded = expandedProjectIds.has(p.id);
                      return (
                        <React.Fragment key={p.id}>
                          <tr
                            className={selectedProjectIds.has(p.id) ? 'bg-[#EBF2F7]/35' : p.weight_warning ? 'bg-[#FEF2F2]/30' : 'hover:bg-slate-50/70'}
                          >
                            <td className="px-2 py-2.5">
                              <input
                                type="checkbox"
                                checked={selectedProjectIds.has(p.id)}
                                onChange={() => toggleSelectProject(p.id)}
                                className="rounded border-slate-300"
                              />
                            </td>
                            <td className="px-1.5 py-2.5">
                              <button
                                onClick={() => toggleExpandProject(p.id)}
                                className="rounded p-0.5 text-slate-400 hover:bg-slate-200/70 hover:text-slate-700"
                                title="展开/收起需求一览表完整字段"
                              >
                                {expanded ? <ChevronUp className="h-3.5 w-3.5" /> : <ChevronDown className="h-3.5 w-3.5" />}
                              </button>
                            </td>
                            <td className="px-2 py-2.5 text-slate-400 tabular-nums">{pageOffset + idx + 1}</td>
                            <td className="px-2.5 py-2.5 whitespace-nowrap">
                              {p.req_app_no ? (
                                <span className="font-mono text-[11px] font-medium text-[#3B5B7E] bg-[#EBF2F7]/70 px-1.5 py-0.5 rounded">
                                  {p.req_app_no}
                                </span>
                              ) : (
                                <span className="text-[11px] text-[#9A6B26] bg-[#FEF9EC] px-1.5 py-0.5 rounded">
                                  待 ERP 回填
                                </span>
                              )}
                            </td>
                            <td className="px-2.5 py-2.5 whitespace-nowrap">
                              {p.lot_code ? (
                                <div>
                                  <div className="font-mono text-slate-700 font-medium">{p.lot_code}</div>
                                  <div className="text-[11px] text-slate-400 flex items-center gap-1">
                                    <span>{p.lot_name} · {p.package_no}</span>
                                    {(p.subcontract_rule?.includes('合并') || p.name?.includes('子包')) && (
                                      <span className="rounded bg-[#FEF9EC] px-1 py-0.2 text-[10px] font-medium text-[#9A6B26]">
                                        合并成包
                                      </span>
                                    )}
                                  </div>
                                </div>
                              ) : (
                                <span className="font-mono text-slate-500">{p.project_code || p.code || '—'}</span>
                              )}
                            </td>
                            <td className="px-2.5 py-2.5">
                              <div className="flex flex-wrap items-center gap-1.5">
                                <span className="font-medium text-slate-800">{p.name}</span>
                                {(projectChangeMap.get(p.id) || []).map((cn) =>
                                  cn.has_file ? (
                                    <a
                                      key={cn.id}
                                      href={`/api/bidding/change-notices/${cn.id}/file`}
                                      target="_blank"
                                      rel="noreferrer"
                                      title={`变更公告：${cn.title}${cn.summary ? ` — ${cn.summary}` : ''}（点击查看原件）`}
                                      className="inline-flex items-center rounded bg-[#FEF9EC] border border-[#9A6B26]/35 px-1.5 py-0.5 text-[10px] font-semibold text-[#9A6B26] hover:bg-[#9A6B26] hover:text-white transition-colors"
                                    >
                                      变更{cn.change_types[0] ? ` · ${cn.change_types[0]}` : ''}
                                    </a>
                                  ) : (
                                    <span
                                      key={cn.id}
                                      title={`变更公告：${cn.title}${cn.summary ? ` — ${cn.summary}` : ''}`}
                                      className="inline-flex items-center rounded bg-[#FEF9EC] border border-[#9A6B26]/35 px-1.5 py-0.5 text-[10px] font-semibold text-[#9A6B26]"
                                    >
                                      变更{cn.change_types[0] ? ` · ${cn.change_types[0]}` : ''}
                                    </span>
                                  )
                                )}
                              </div>
                              {p.project_overview && (
                                <div className="mt-0.5 text-[11px] text-slate-500 line-clamp-2" title={p.project_overview}>
                                  概况：{p.project_overview}
                                </div>
                              )}
                              <div className="mt-0.5 text-[10px] text-slate-400 font-mono">
                                {[p.project_code, p.tech_spec_code].filter(Boolean).join(' · ')}
                              </div>
                            </td>
                            <td className="px-2.5 py-2.5 whitespace-nowrap tabular-nums font-medium text-slate-700">
                              {p.max_price_raw ? (
                                <div className="space-y-0.5">
                                  <span className="inline-block rounded bg-[#EBF2F7] px-1.5 py-0.5 text-[#3B5B7E]">
                                    限价 {p.max_price_raw}
                                  </span>
                                  {p.budget_yuan !== null && p.budget_yuan !== undefined && (
                                    <div className="text-[10px] font-normal text-slate-500">
                                      预算 {formatMoney(p.budget_yuan, priceUnit)}
                                    </div>
                                  )}
                                </div>
                              ) : (
                                formatMoney(p.max_price_yuan ?? p.budget_yuan, priceUnit, p.budget || '—')
                              )}
                            </td>
                            <td className="px-2.5 py-2.5 whitespace-nowrap">
                              {p.weight_ratio ? (
                                <div className="space-y-1">
                                  <span
                                    className={`inline-block rounded px-1.5 py-0.5 font-mono text-[11px] font-medium ${
                                      p.weight_warning
                                        ? 'bg-[#FEF2F2] text-[#A94442] border border-[#A94442]/25'
                                        : 'bg-slate-100 text-slate-700'
                                    }`}
                                  >
                                    {p.weight_ratio}
                                  </span>
                                  {p.weight_suggested && (
                                    <button
                                      onClick={() => handleFixWeight(p.id, p.weight_suggested!)}
                                      className="block rounded bg-[#3B5B7E] px-2 py-0.5 text-[10px] font-medium text-white hover:bg-[#2E4866]"
                                    >
                                      <Wrench className="inline h-2.5 w-2.5 mr-0.5" />修正为 {p.weight_suggested}
                                    </button>
                                  )}
                                </div>
                              ) : (
                                <span className="text-slate-400">—</span>
                              )}
                            </td>
                            <td className="px-2.5 py-2.5 whitespace-nowrap text-[11px] font-mono text-slate-500">
                              {p.biz_eval_tpl || p.tech_eval_tpl
                                ? `${p.biz_eval_tpl || '-'} / ${p.tech_eval_tpl || '-'}`
                                : '—'}
                            </td>
                            <td className="px-2 py-2.5 text-right whitespace-nowrap">
                              <div className="inline-flex items-center gap-0.5">
                                <button
                                  type="button"
                                  onClick={() => setEditingProject(p)}
                                  className="rounded p-1 text-slate-400 hover:bg-[#EBF2F7] hover:text-[#3B5B7E]"
                                  title="编辑项目主数据（同步开评标与定标）"
                                >
                                  <Pencil className="h-3.5 w-3.5" />
                                </button>
                                <button
                                  type="button"
                                  onClick={() => handleDeleteProject(p.id)}
                                  className="rounded p-1 text-slate-400 hover:bg-red-50 hover:text-[#A94442]"
                                  title="删除项目"
                                >
                                  <Trash2 className="h-3.5 w-3.5" />
                                </button>
                              </div>
                            </td>
                          </tr>
                          {expanded && (
                            <tr className="bg-slate-50/80">
                              <td colSpan={10} className="px-6 py-3 space-y-2.5">
                                {(projectChangeMap.get(p.id) || []).length > 0 && (
                                  <div className="rounded-lg border border-[#9A6B26]/30 bg-[#FEF9EC]/85 px-3 py-2 text-xs space-y-1">
                                    <div className="font-semibold text-[#9A6B26]">关联变更公告提醒</div>
                                    {(projectChangeMap.get(p.id) || []).map((cn) => (
                                      <div key={cn.id} className="flex flex-wrap items-center justify-between gap-2 text-slate-700">
                                        <span>
                                          <strong>{cn.title}</strong>
                                          {cn.notice_date ? `（${cn.notice_date}）` : ''}
                                          {cn.summary ? `：${cn.summary}` : ''}
                                        </span>
                                        {cn.has_file && (
                                          <a
                                            href={`/api/bidding/change-notices/${cn.id}/file`}
                                            target="_blank"
                                            rel="noreferrer"
                                            className="shrink-0 text-[11px] font-medium text-[#3B5B7E] hover:underline"
                                          >
                                            查看原件
                                          </a>
                                        )}
                                      </div>
                                    ))}
                                  </div>
                                )}
                                <div className="grid grid-cols-1 gap-2.5 text-xs sm:grid-cols-2 lg:grid-cols-3">
                                  <div><span className="text-slate-400">项目单位：</span><span className="text-slate-700">{p.project_unit || '—'}</span></div>
                                  <div><span className="text-slate-400">项目管理部门：</span><span className="text-slate-700">{p.department || '—'}</span></div>
                                  <div><span className="text-slate-400">工期：</span><span className="text-slate-700">{p.duration || '—'}</span></div>
                                  <div><span className="text-slate-400">含税预算金额：</span><span className="text-slate-700 tabular-nums">{formatMoney(p.budget_yuan, priceUnit)}</span></div>
                                  <div><span className="text-slate-400">报价方式 / 分包原则：</span><span className="text-slate-700">{[p.quote_method, p.subcontract_rule].filter(Boolean).join(' · ') || '—'}</span></div>
                                  <div><span className="text-slate-400">价格公式与参数：</span><span className="text-slate-700">{[p.price_formula, p.price_params].filter(Boolean).join(' | ') || '—'}</span></div>
                                  <div><span className="text-slate-400">合同版本：</span><span className="text-slate-700">{p.contract_version || '—'}</span></div>
                                  <div><span className="text-slate-400">资金来源：</span><span className="text-slate-700">{p.funding_source || '—'}</span></div>
                                  <div><span className="text-slate-400">ERP 同步状态：</span><span className="text-slate-700">{p.erp_sync_status || '正常'}</span></div>
                                  <div className="sm:col-span-2 lg:col-span-3"><span className="text-slate-400">资质条件：</span><span className="text-slate-700">{p.qualification_req || '—'}</span></div>
                                  <div className="sm:col-span-2 lg:col-span-3"><span className="text-slate-400">业绩要求：</span><span className="text-slate-700">{p.performance_req || '—'}</span></div>
                                  <div className="sm:col-span-2 lg:col-span-3"><span className="text-slate-400">人员要求：</span><span className="text-slate-700">{p.staff_req || '—'}</span></div>
                                </div>
                              </td>
                            </tr>
                          )}
                        </React.Fragment>
                      );
                    })}
                  </tbody>
                </table>
              ) : (
                /* FULL 23-COLUMN VIEW */
                <table className="w-full border-collapse text-left text-xs whitespace-nowrap">
                  <thead>
                    <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                      <th className="px-2 py-2 w-8">
                        <input
                          type="checkbox"
                          checked={sortedProjects.length > 0 && sortedProjects.every((p) => selectedProjectIds.has(p.id))}
                          onChange={toggleSelectAllVisible}
                          className="rounded border-slate-300"
                        />
                      </th>
                      <th className="px-2 py-2 font-medium">操作</th>
                      <th className="px-2.5 py-2 font-medium">{renderSortBtn('#', 'seq')}</th>
                      <th className="px-2.5 py-2 font-medium">{renderSortBtn('采购申请号', 'req_app_no')}</th>
                      <th className="px-2.5 py-2 font-medium">{renderSortBtn('分标编号', 'lot_code')}</th>
                      <th className="px-2.5 py-2 font-medium">{renderSortBtn('分标名称', 'lot_name')}</th>
                      <th className="px-2.5 py-2 font-medium">{renderSortBtn('包号', 'package_no')}</th>
                      <th className="px-2.5 py-2 font-medium">{renderSortBtn('项目单位', 'project_unit')}</th>
                      <th className="px-2.5 py-2 font-medium">{renderSortBtn('项目名称', 'name')}</th>
                      <th className="px-2.5 py-2 font-medium">项目概况</th>
                      <th className="px-2.5 py-2 font-medium">{renderSortBtn('工期', 'duration')}</th>
                      <th className="px-2.5 py-2 font-medium">资质条件</th>
                      <th className="px-2.5 py-2 font-medium">业绩要求</th>
                      <th className="px-2.5 py-2 font-medium">人员要求</th>
                      <th className="px-2.5 py-2 font-medium">{renderSortBtn(`含税预算(${priceUnit === 'wan' ? '万元' : '元'})`, 'budget_yuan')}</th>
                      <th className="px-2.5 py-2 font-medium">{renderSortBtn(`最高限价(${priceUnit === 'wan' ? '万元' : '元'})`, 'max_price_yuan')}</th>
                      <th className="px-2.5 py-2 font-medium">报价方式</th>
                      <th className="px-2.5 py-2 font-medium">分包原则</th>
                      <th className="px-2.5 py-2 font-medium">技术规范书编码</th>
                      <th className="px-2.5 py-2 font-medium">{renderSortBtn('权重(技:商:价)', 'weight_ratio')}</th>
                      <th className="px-2.5 py-2 font-medium">价格公式</th>
                      <th className="px-2.5 py-2 font-medium">参数选择</th>
                      <th className="px-2.5 py-2 font-medium">{renderSortBtn('商务模板', 'biz_eval_tpl')}</th>
                      <th className="px-2.5 py-2 font-medium">技术模板</th>
                      <th className="px-2.5 py-2 font-medium">{renderSortBtn('合同版本', 'contract_version')}</th>
                      <th className="px-2.5 py-2 font-medium">申报人/电话</th>
                      <th className="px-2.5 py-2 font-medium">{renderSortBtn('管理部门', 'department')}</th>
                      <th className="px-2.5 py-2 font-medium">项目编码</th>
                      <th className="px-2.5 py-2 font-medium">{renderSortBtn('资金来源', 'funding_source')}</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {sortedProjects.length === 0 ? (
                      <tr>
                        <td colSpan={29} className="py-8 text-center text-xs text-slate-400">
                          未找到匹配的项目记录
                        </td>
                      </tr>
                    ) : pagedProjects.map((p, idx) => (
                      <tr key={p.id} className={selectedProjectIds.has(p.id) ? 'bg-[#EBF2F7]/35' : 'hover:bg-slate-50/70'}>
                        <td className="px-2 py-2">
                          <input
                            type="checkbox"
                            checked={selectedProjectIds.has(p.id)}
                            onChange={() => toggleSelectProject(p.id)}
                            className="rounded border-slate-300"
                          />
                        </td>
                        <td className="px-2 py-2">
                          <button
                            type="button"
                            onClick={() => setEditingProject(p)}
                            className="rounded p-1 text-slate-400 hover:bg-[#EBF2F7] hover:text-[#3B5B7E]"
                            title="编辑项目主数据"
                          >
                            <Pencil className="h-3.5 w-3.5" />
                          </button>
                        </td>
                        <td className="px-2.5 py-2 text-slate-400 tabular-nums">{pageOffset + idx + 1}</td>
                        <td className="px-2.5 py-2 font-mono text-[#3B5B7E]">{p.req_app_no || '待回填'}</td>
                        <td className="px-2.5 py-2 font-mono">{p.lot_code || '—'}</td>
                        <td className="px-2.5 py-2">{p.lot_name || '—'}</td>
                        <td className="px-2.5 py-2">{p.package_no || '—'}</td>
                        <td className="px-2.5 py-2">{p.project_unit || '—'}</td>
                        <td className="px-2.5 py-2 font-medium text-slate-800">
                          <div className="flex items-center gap-1.5">
                            <span>{p.name}</span>
                            {(projectChangeMap.get(p.id) || []).length > 0 && (
                              <span className="rounded bg-[#FEF9EC] border border-[#9A6B26]/35 px-1.5 py-0.5 text-[10px] font-semibold text-[#9A6B26]">
                                变更
                              </span>
                            )}
                          </div>
                        </td>
                        <td className="px-2.5 py-2 max-w-[220px] truncate" title={p.project_overview}>{p.project_overview || '—'}</td>
                        <td className="px-2.5 py-2">{p.duration || '—'}</td>
                        <td className="px-2.5 py-2 max-w-[200px] truncate" title={p.qualification_req}>{p.qualification_req || '—'}</td>
                        <td className="px-2.5 py-2 max-w-[200px] truncate" title={p.performance_req}>{p.performance_req || '—'}</td>
                        <td className="px-2.5 py-2 max-w-[160px] truncate" title={p.staff_req}>{p.staff_req || '—'}</td>
                        <td className="px-2.5 py-2 tabular-nums">{formatMoney(p.budget_yuan, priceUnit)}</td>
                        <td className="px-2.5 py-2 tabular-nums font-medium">{p.max_price_raw || formatMoney(p.max_price_yuan ?? p.budget_yuan, priceUnit)}</td>
                        <td className="px-2.5 py-2">{p.quote_method || '—'}</td>
                        <td className="px-2.5 py-2">{p.subcontract_rule || '—'}</td>
                        <td className="px-2.5 py-2 font-mono">{p.tech_spec_code || '—'}</td>
                        <td className="px-2.5 py-2 font-mono">{p.weight_ratio || '—'}</td>
                        <td className="px-2.5 py-2">{p.price_formula || '—'}</td>
                        <td className="px-2.5 py-2 font-mono">{p.price_params || '—'}</td>
                        <td className="px-2.5 py-2 font-mono">{p.biz_eval_tpl || '—'}</td>
                        <td className="px-2.5 py-2 font-mono">{p.tech_eval_tpl || '—'}</td>
                        <td className="px-2.5 py-2">{p.contract_version || '—'}</td>
                        <td className="px-2.5 py-2">{[p.declarant_name, p.declarant_phone].filter(Boolean).join(' ') || '—'}</td>
                        <td className="px-2.5 py-2">{p.department || '—'}</td>
                        <td className="px-2.5 py-2 font-mono">{p.project_code || '—'}</td>
                        <td className="px-2.5 py-2">{p.funding_source || '—'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>

            {sortedProjects.length > 20 && (
              <div className="overflow-hidden rounded-lg border border-slate-200">
                <PaginationBar
                  total={sortedProjects.length}
                  page={projPage}
                  pageSize={projPageSize}
                  onPageChange={setProjPage}
                  onPageSizeChange={setProjPageSize}
                />
              </div>
            )}

            {/* Inline Manual Add Project */}
            <form onSubmit={handleAddProject} className="grid grid-cols-1 gap-2 rounded-lg bg-slate-50 p-3 sm:grid-cols-6">
              <input
                required
                value={newProjName}
                onChange={(e) => setNewProjName(e.target.value)}
                placeholder="手动新增项目/标段名称 *"
                className="sm:col-span-2 rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
              />
              <input
                value={newProjCode}
                onChange={(e) => setNewProjCode(e.target.value)}
                placeholder="标段/项目编号"
                className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
              />
              <input
                value={newProjBudget}
                onChange={(e) => setNewProjBudget(e.target.value)}
                placeholder="最高限价（如 60万元）"
                className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
              />
              <select
                value={newProjMethod}
                onChange={(e) => setNewProjMethod(e.target.value)}
                className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
              >
                {PROCUREMENT_METHODS.map((m) => (
                  <option key={m} value={m}>
                    {m}
                  </option>
                ))}
              </select>
              <div className="flex gap-2">
                <input
                  value={newProjDept}
                  onChange={(e) => setNewProjDept(e.target.value)}
                  placeholder="管理部门"
                  className="min-w-0 flex-1 rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                />
                <button
                  type="submit"
                  className="shrink-0 inline-flex items-center gap-1 rounded-lg bg-[#3B5B7E] px-3 py-1.5 text-xs font-medium text-white hover:bg-[#2E4866]"
                >
                  <Plus className="h-3.5 w-3.5" /> 添加
                </button>
              </div>
            </form>
          </div>
        )}

        {/* TAB 2: Cleaned Supplier Contacts & Automatic Collusion Screening */}
        {activeTab === 'suppliers' && (
          <div className="p-4 space-y-4">
            {suppliers.length > 0 && (
              <div
                className={`rounded-lg border p-3.5 space-y-2 ${
                  collusionRisks.length > 0
                    ? 'border-[#9A6B26]/30 bg-[#FEF9EC]/60'
                    : 'border-[#2F6F5E]/25 bg-[#ECFDF5]/50'
                }`}
              >
                <div className="flex items-center gap-2 text-xs font-semibold text-slate-800">
                  {collusionRisks.length > 0 ? (
                    <ShieldAlert className="h-4 w-4 text-[#9A6B26]" />
                  ) : (
                    <CheckCircle2 className="h-4 w-4 text-[#2F6F5E]" />
                  )}
                  <span>
                    {terms.platformShort} 报名{terms.bidder}围串标风险初筛：
                    {collusionRisks.length > 0
                      ? `发现 ${collusionRisks.length} 项跨供应商重复手机/邮箱线索（已自动去除空格后比对）`
                      : '未发现不同供应商共用相同手机号或电子邮件的情况（同单位报多包不计入异常）'}
                  </span>
                </div>
                {collusionRisks.length > 0 && (
                  <div className="grid grid-cols-1 gap-2 sm:grid-cols-2 pt-1">
                    {collusionRisks.map((rk, i) => (
                      <div key={i} className="rounded-md bg-white/90 border border-[#9A6B26]/20 p-2.5 text-xs">
                        <div className="flex items-center justify-between gap-2">
                          <span className="font-semibold text-[#9A6B26]">{rk.label}</span>
                          <span className="font-mono text-[11px] text-slate-500">{rk.value}</span>
                        </div>
                        <div className="mt-1 text-slate-700">{rk.suppliers.join('  /  ')}</div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}

            <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
              <div className="text-xs text-slate-500">
                所有{terms.bidder}名称、联系人、电话、邮箱、税号及银行账号均已自动去除多余空格，供定标《成交供应商交接单》直接调用。
              </div>
              <div className="relative w-full sm:w-64">
                <Search className="pointer-events-none absolute left-2.5 top-2 h-3.5 w-3.5 text-slate-400" />
                <input
                  value={supplierSearch}
                  onChange={(e) => setSupplierSearch(e.target.value)}
                  placeholder={`搜索${terms.bidder}、联系人、手机或税号...`}
                  className="w-full rounded-lg border border-slate-200 py-1.5 pl-8 pr-3 text-xs focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>
            </div>

            {suppliers.length === 0 ? (
              <div className="py-10 text-center text-xs text-slate-400">
                请先在上方第 5 号槽位上传《供应商联系人信息表》
              </div>
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full border-collapse text-left text-xs">
                  <thead>
                    <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                      <th className="px-2.5 py-2 font-medium w-10">#</th>
                      <th className="px-2.5 py-2 font-medium">{terms.bidder}名称</th>
                      <th className="px-2.5 py-2 font-medium">联系人</th>
                      <th className="px-2.5 py-2 font-medium">手机 / 固话</th>
                      <th className="px-2.5 py-2 font-medium">电子邮件</th>
                      <th className="px-2.5 py-2 font-medium">税号</th>
                      <th className="px-2.5 py-2 font-medium">开户银行及账号</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {paginate(filteredSuppliers, supplierPager.page, supplierPager.pageSize).map((s) => (
                      <tr key={s.id} className="hover:bg-slate-50/70">
                        <td className="px-2.5 py-2 text-slate-400 tabular-nums">{s.seq}</td>
                        <td className="px-2.5 py-2 font-medium text-slate-800">{s.supplier_name}</td>
                        <td className="px-2.5 py-2 text-slate-700 whitespace-nowrap">{s.contact_name || '—'}</td>
                        <td className="px-2.5 py-2 font-mono text-slate-600 whitespace-nowrap">
                          <div>{s.mobile || '—'}</div>
                          {s.phone && <div className="text-[10px] text-slate-400">{s.phone}</div>}
                        </td>
                        <td className="px-2.5 py-2 text-slate-500">{s.email || '—'}</td>
                        <td className="px-2.5 py-2 font-mono text-[11px] text-slate-500">{s.tax_no || '—'}</td>
                        <td className="px-2.5 py-2 text-[11px] text-slate-500">
                          <div className="truncate max-w-[200px]" title={s.bank_name}>{s.bank_name || '—'}</div>
                          <div className="font-mono text-slate-400">{s.bank_account}</div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {filteredSuppliers.length > 20 && (
                  <div className="mt-3 overflow-hidden rounded-lg border border-slate-200">
                    <PaginationBar
                      total={filteredSuppliers.length}
                      page={supplierPager.page}
                      pageSize={supplierPager.pageSize}
                      onPageChange={supplierPager.setPage}
                      onPageSizeChange={supplierPager.setPageSize}
                    />
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {/* TAB 3: Batch Basic Info & Terminology Settings */}
        {activeTab === 'batch_info' && (
          <form onSubmit={handleSaveBatch} className="p-5 space-y-4">
            <div className="flex items-center justify-between">
              <div className="text-xs text-slate-500">
                选择「采购类型」与「单位性质」后，全站术语（{terms.initiator}/{terms.bidder}/{terms.tenderDoc}/{terms.responseDoc}）与平台（{terms.platformShort}）将自动联动。
              </div>
              <button
                type="submit"
                disabled={saving}
                className="inline-flex shrink-0 items-center gap-1.5 rounded-lg bg-[#3B5B7E] px-3.5 py-1.5 text-xs font-medium text-white hover:bg-[#2E4866] disabled:opacity-50"
              >
                {savedToast ? <Check className="h-3.5 w-3.5" /> : <Save className="h-3.5 w-3.5" />}
                <span>{savedToast ? '已保存' : saving ? '保存中...' : '保存修改'}</span>
              </button>
            </div>

            <div className="grid grid-cols-1 gap-4 md:grid-cols-2 lg:grid-cols-3">
              <div>
                <label className="block text-xs font-medium text-slate-500 mb-1">采购类型（决定专业术语）</label>
                <select
                  value={procurementType}
                  onChange={(e) => setProcurementType(e.target.value as '公开招标' | '非公开招标')}
                  className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm focus:border-[#3B5B7E] focus:outline-none"
                >
                  <option value="非公开招标">非公开招标（竞争性谈判采购）</option>
                  <option value="公开招标">公开招标</option>
                </select>
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-500 mb-1">项目单位性质（决定执行平台）</label>
                <select
                  value={unitNature}
                  onChange={(e) => setUnitNature(e.target.value as '主业' | '产业')}
                  className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm focus:border-[#3B5B7E] focus:outline-none"
                >
                  <option value="产业">产业（ETP 电工交易专区）</option>
                  <option value="主业">主业（ECP 2.0）</option>
                </select>
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-500 mb-1">当前阶段状态</label>
                <select
                  value={status}
                  onChange={(e) => setStatus(e.target.value)}
                  className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm focus:border-[#3B5B7E] focus:outline-none"
                >
                  {BATCH_STATUSES.map((s) => (
                    <option key={s} value={s}>
                      {s}
                    </option>
                  ))}
                </select>
              </div>

              <div className="md:col-span-2">
                <label className="block text-xs font-medium text-slate-500 mb-1">批次名称 *</label>
                <input
                  required
                  value={name}
                  onChange={(e) => setName(e.target.value)}
                  className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-500 mb-1">采购 / 批次编号</label>
                <input
                  value={code}
                  onChange={(e) => setCode(e.target.value)}
                  className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm font-mono focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-500 mb-1">
                  项目单位（{terms.initiator}）
                </label>
                <SearchableClientSelect
                  clients={clients}
                  value={projectUnit}
                  valueMode="name"
                  allowCustom
                  placeholder="搜索并选择协议项目单位..."
                  onChange={(val) => setProjectUnit(val)}
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-500 mb-1">公告日期（挂网时间）</label>
                <DateTimePickerField
                  value={announcementTime}
                  onChange={setAnnouncementTime}
                  allowTime
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-500 mb-1">{terms.fetchDocLabel}</label>
                <DateTimePickerField
                  mode="range"
                  value={docFetchTime}
                  onChange={setDocFetchTime}
                  allowTime
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-500 mb-1">{terms.deadlineLabel}</label>
                <DateTimePickerField
                  value={bidOpeningTime}
                  onChange={setBidOpeningTime}
                  allowTime
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-500 mb-1">代理联系人电话</label>
                <input
                  value={contactPhone}
                  onChange={(e) => setContactPhone(e.target.value)}
                  className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm font-mono focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-500 mb-1">代理联系邮箱</label>
                <input
                  value={contactEmail}
                  onChange={(e) => setContactEmail(e.target.value)}
                  className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm font-mono focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>
              <div className="md:col-span-2 lg:col-span-3">
                <label className="block text-xs font-medium text-slate-500 mb-1">批次备注</label>
                <input
                  value={notes}
                  onChange={(e) => setNotes(e.target.value)}
                  className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>
            </div>

            {/* Extensible Key-Value Extra Fields */}
            <div className="border-t border-slate-100 pt-3">
              <div className="flex items-center justify-between mb-2">
                <span className="text-xs font-medium text-slate-500">自定义扩展属性</span>
                <button
                  type="button"
                  onClick={() => setExtraFields((prev) => [...prev, { key: '', value: '' }])}
                  className="inline-flex items-center gap-1 text-xs font-medium text-[#3B5B7E] hover:underline"
                >
                  <Plus className="h-3 w-3" /> 添加字段
                </button>
              </div>
              <div className="grid grid-cols-1 gap-2 sm:grid-cols-2">
                {extraFields.map((ef, idx) => (
                  <div key={idx} className="flex items-center gap-2">
                    <input
                      value={ef.key}
                      onChange={(e) => {
                        const next = [...extraFields];
                        next[idx] = { ...next[idx], key: e.target.value };
                        setExtraFields(next);
                      }}
                      placeholder="属性名"
                      className="w-36 rounded-lg border border-slate-200 px-2.5 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                    />
                    <input
                      value={ef.value}
                      onChange={(e) => {
                        const next = [...extraFields];
                        next[idx] = { ...next[idx], value: e.target.value };
                        setExtraFields(next);
                      }}
                      placeholder="属性值"
                      className="flex-1 rounded-lg border border-slate-200 px-2.5 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                    />
                    <button
                      type="button"
                      onClick={() => setExtraFields((prev) => prev.filter((_, i) => i !== idx))}
                      className="rounded p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-600"
                    >
                      <Trash2 className="h-3.5 w-3.5" />
                    </button>
                  </div>
                ))}
              </div>
            </div>
          </form>
        )}
      </div>

      {editingProject && (
        <ProjectEditModal
          project={editingProject}
          onClose={() => setEditingProject(null)}
          onSaved={async () => {
            await loadData();
            await refreshBatches();
          }}
        />
      )}

      {previewData && (
        <ExcelPreviewModal
          title={previewData.title}
          sheets={previewData.sheets}
          onClose={() => setPreviewData(null)}
        />
      )}
    </div>
  );
}
