import React, { useCallback, useEffect, useState } from 'react';
import { Link } from 'react-router-dom';
import {
  CheckCircle2,
  AlertCircle,
  FileSpreadsheet,
  FileText,
  Sparkles,
  Download,
  Trash2,
  X,
  ArrowRight,
  Calculator,
  Upload,
  Save,
  Building2,
} from 'lucide-react';
import { api } from '../api/client';
import { DateTimePickerField } from '../components/DateTimePickerField';
import { paginate, PaginationBar } from '../components/PaginationBar';
import { SearchableClientSelect } from '../components/SearchableClientSelect';
import { useBatch } from '../context/BatchContext';
import type { AwardOverview, AwardTemplate } from '../types';
import { getTerminology } from '../utils/terminology';

const CATEGORY_OPTIONS: Array<{ value: '服务' | '工程' | '物资'; label: string }> = [
  { value: '服务', label: '服务类' },
  { value: '工程', label: '施工 / 工程类' },
  { value: '物资', label: '物资 / 货物类' },
];

interface DocxColumnOptions {
  includeLotCode: boolean;
  includeLotName: boolean;
  includeProjectName: boolean;
  includePrice: boolean;
}

const DOCX_OPTIONAL_COLUMNS: Array<{ key: keyof DocxColumnOptions; label: string }> = [
  { key: 'includeLotCode', label: '分标编号' },
  { key: 'includeLotName', label: '分标名称' },
  { key: 'includeProjectName', label: '项目名称' },
  { key: 'includePrice', label: '价格' },
];

export function AwardPage() {
  const { currentBatchId, refreshBatches } = useBatch();
  const [data, setData] = useState<AwardOverview | null>(null);
  const [activeTemplate, setActiveTemplate] = useState<AwardTemplate | null>(null);
  const [overrides, setOverrides] = useState<Record<string, string>>({});
  const [generating, setGenerating] = useState(false);
  const [exportingKey, setExportingKey] = useState<string | null>(null);
  const [pkgPage, setPkgPage] = useState(1);
  const [pkgPageSize, setPkgPageSize] = useState(20);
  const OUTPUT_PAGE_SIZE = 5;
  const [outputPage, setOutputPage] = useState(1);
  const [selectedOutputIds, setSelectedOutputIds] = useState<Set<string>>(new Set());
  const [batchDeletingOutputs, setBatchDeletingOutputs] = useState(false);
  const [publicityCols, setPublicityCols] = useState<DocxColumnOptions>({
    includeLotCode: true,
    includeLotName: true,
    includeProjectName: false,
    includePrice: false,
  });
  const [noticeCols, setNoticeCols] = useState<DocxColumnOptions>({
    includeLotCode: false,
    includeLotName: true,
    includeProjectName: true,
    includePrice: true,
  });

  // Config form states
  const [clientId, setClientId] = useState('');
  const [category, setCategory] = useState<'服务' | '工程' | '物资'>('服务');
  const [discountRate, setDiscountRate] = useState('1');
  const [batchManager, setBatchManager] = useState('');
  const [bidLocation, setBidLocation] = useState('评标基地');
  const [publicityDate, setPublicityDate] = useState('');
  const [announcementDate, setAnnouncementDate] = useState('');
  const [savingCfg, setSavingCfg] = useState(false);
  const [cfgSavedMsg, setCfgSavedMsg] = useState('');
  const [importingClients, setImportingClients] = useState(false);
  const [importingMaterials, setImportingMaterials] = useState(false);

  const loadData = useCallback(async () => {
    if (!currentBatchId) return;
    const res = await api.getAwardOverview(currentBatchId);
    setData(res);
    setClientId(res.awardConfig.client_id || '');
    setCategory(res.awardConfig.category || '服务');
    setDiscountRate(String(res.awardConfig.discount_rate ?? 1));
    setBatchManager(res.awardConfig.batch_manager || '');
    setBidLocation(res.awardConfig.bid_location || '评标基地');
    setPublicityDate(res.awardConfig.publicity_date || '');
    setAnnouncementDate(res.awardConfig.announcement_date || '');
    const isPub = res.batch.procurement_type === '公开招标';
    setPublicityCols({
      includeLotCode: !isPub,
      includeLotName: true,
      includeProjectName: isPub,
      includePrice: isPub,
    });
    setNoticeCols({
      includeLotCode: isPub,
      includeLotName: !isPub,
      includeProjectName: !isPub,
      includePrice: !isPub,
    });
  }, [currentBatchId]);

  useEffect(() => {
    loadData();
    setPkgPage(1);
    setOutputPage(1);
    setSelectedOutputIds(new Set());
  }, [loadData]);

  async function handleClientChange(nextClientId: string) {
    setClientId(nextClientId);
    if (!data || !currentBatchId) return;
    const found = data.clients.find((c) => c.id === nextClientId);
    const nextRate = found ? found.discount_rate : 1;
    setDiscountRate(String(nextRate));
    // Immediately persist & recalculate
    setSavingCfg(true);
    try {
      await api.saveAwardConfig(currentBatchId, {
        client_id: nextClientId,
        category,
        discount_rate: nextRate,
        batch_manager: batchManager,
        bid_location: bidLocation,
      });
      await loadData();
    } finally {
      setSavingCfg(false);
    }
  }

  async function handleCategoryChange(nextCat: '服务' | '工程' | '物资') {
    setCategory(nextCat);
    if (!currentBatchId) return;
    setSavingCfg(true);
    try {
      await api.saveAwardConfig(currentBatchId, {
        client_id: clientId,
        category: nextCat,
        discount_rate: Number(discountRate) || 1,
        batch_manager: batchManager,
        bid_location: bidLocation,
      });
      await loadData();
    } finally {
      setSavingCfg(false);
    }
  }

  async function handleSaveConfig(e?: React.FormEvent) {
    if (e) e.preventDefault();
    if (!currentBatchId) return;
    setSavingCfg(true);
    try {
      await api.saveAwardConfig(currentBatchId, {
        client_id: clientId,
        category,
        discount_rate: Number(discountRate) || 1,
        batch_manager: batchManager,
        bid_location: bidLocation,
      });
      await loadData();
      setCfgSavedMsg('已保存并重算');
      setTimeout(() => setCfgSavedMsg(''), 2000);
    } finally {
      setSavingCfg(false);
    }
  }

  async function handleImportClients(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    e.target.value = '';
    setImportingClients(true);
    try {
      const res = await api.importClientAgreements(file);
      alert(`客户协议清单更新成功，共导入 ${res.count} 条业主单位费率。`);
      await loadData();
    } catch (err) {
      alert((err as Error).message);
    } finally {
      setImportingClients(false);
    }
  }

  async function handleImportMaterials(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    e.target.value = '';
    setImportingMaterials(true);
    try {
      const res = await api.importMaterialCatalog(file);
      alert(`服务物料编码对照表更新成功，共导入 ${res.count} 条物料描述映射。`);
    } catch (err) {
      alert((err as Error).message);
    } finally {
      setImportingMaterials(false);
    }
  }

  async function handleAwardDateChange(field: 'publicity_date' | 'announcement_date', val: string) {
    if (field === 'publicity_date') setPublicityDate(val);
    else setAnnouncementDate(val);
    if (!currentBatchId) return;
    try {
      await api.saveAwardConfig(currentBatchId, { [field]: val });
    } catch {
      // ignore transient save errors
    }
  }

  function buildDefaultOverrides(tplKey?: string): Record<string, string> {
    if (!data) return {};
    const isPublic = data.batch.procurement_type === '公开招标';
    const today = new Date().toISOString().slice(0, 10);
    const cols = tplKey === 'pre_award_publicity' ? publicityCols : noticeCols;
    const effPublicityDate = publicityDate || today;
    const effAnnouncementDate = announcementDate || today;
    return {
      batchName: data.batch.name,
      batchCode: data.batch.code,
      ownerUnit: data.awardConfig.owner_unit || data.batch.project_unit,
      agencyName: 'XX工程咨询有限公司招标代理机构',
      announcementTime: data.batch.announcement_time,
      bidOpeningTime: data.batch.bid_opening_time,
      publicityDays: isPublic ? '3' : '1',
      publicityDate: effPublicityDate,
      announcementDate: effAnnouncementDate,
      objectionEmail: data.batch.contact_email || 'agency@example.com',
      objectionPhone: data.batch.contact_phone || '000-00000000',
      batchManager: batchManager,
      bidLocation: bidLocation || '评标基地',
      matchDate: data.batch.announcement_time || '',
      noticeDate: effAnnouncementDate,
      noticePushDate: effAnnouncementDate,
      category,
      discountRate,
      includeLotCode: cols.includeLotCode ? 'true' : 'false',
      includeLotName: cols.includeLotName ? 'true' : 'false',
      includeProjectName: cols.includeProjectName ? 'true' : 'false',
      includePrice: cols.includePrice ? 'true' : 'false',
    };
  }

  function triggerDownload(outputId: string) {
    const a = document.createElement('a');
    a.href = `/api/award/outputs/${outputId}/download`;
    a.download = '';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }

  function openGenerateModal(tpl: AwardTemplate) {
    if (!data) return;
    setActiveTemplate(tpl);
    setOverrides(buildDefaultOverrides(tpl.key));
  }

  async function handleQuickExport(tpl: AwardTemplate) {
    if (!currentBatchId || !data) return;
    setExportingKey(tpl.key);
    try {
      const out = await api.generateAwardDoc(currentBatchId, tpl.key, buildDefaultOverrides(tpl.key));
      await loadData();
      await refreshBatches();
      triggerDownload(out.id);
    } catch (err) {
      alert((err as Error).message);
    } finally {
      setExportingKey(null);
    }
  }

  async function handleGenerate(e: React.FormEvent) {
    e.preventDefault();
    if (!currentBatchId || !activeTemplate) return;
    setGenerating(true);
    try {
      const out = await api.generateAwardDoc(currentBatchId, activeTemplate.key, overrides);
      setActiveTemplate(null);
      await loadData();
      await refreshBatches();
      triggerDownload(out.id);
    } catch (err) {
      alert((err as Error).message);
    } finally {
      setGenerating(false);
    }
  }

  async function handleDeleteOutput(id: string) {
    await api.deleteAwardOutput(id);
    setSelectedOutputIds((prev) => {
      const next = new Set(prev);
      next.delete(id);
      return next;
    });
    await loadData();
    await refreshBatches();
  }

  function toggleSelectOutput(id: string) {
    setSelectedOutputIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  async function handleBatchDeleteOutputs() {
    const ids = Array.from(selectedOutputIds);
    if (ids.length === 0) return;
    if (!window.confirm(`确认批量删除已勾选的 ${ids.length} 份定标交付文件记录吗？`)) return;
    setBatchDeletingOutputs(true);
    try {
      await api.batchDeleteAwardOutputs(ids);
      setSelectedOutputIds(new Set());
      await loadData();
      await refreshBatches();
    } catch (err) {
      alert((err as Error).message);
    } finally {
      setBatchDeletingOutputs(false);
    }
  }

  if (!data) {
    return <div className="py-12 text-center text-sm text-slate-400">正在加载定标数据与客户协议费率库...</div>;
  }

  const terms = getTerminology(data.batch.procurement_type, data.batch.unit_nature);
  const selectedClientObj = data.clients.find((c) => c.id === clientId);

  const readinessItems = [
    {
      label: '批次与项目主数据',
      ok: data.readiness.projectCount > 0,
      detail: `已聚合 ${data.readiness.projectCount} 个标包`,
      link: '/bidding',
    },
    {
      label: '开评标排序与结果',
      ok: data.totals.winningCount > 0 || data.totals.failedCount > 0,
      detail: `成交 ${data.totals.winningCount} 包 · 流标 ${data.totals.failedCount} 包`,
      link: '/evaluation',
    },
    {
      label: '业主协议费率定位',
      ok: Boolean(selectedClientObj),
      detail: selectedClientObj ? `${selectedClientObj.discount_label}` : `当前 ${data.awardConfig.discount_label}`,
      link: '#fee-config',
    },
    {
      label: '已生成定标交付物',
      ok: data.outputs.length > 0,
      detail: `已产出 ${data.outputs.length} 份文件`,
      link: '#award-outputs',
    },
  ];

  const wordTemplates = data.templates.filter((t) => t.file_type === 'docx');
  const excelTemplates = data.templates.filter((t) => t.file_type === 'xlsx');

  return (
    <div className="space-y-5">
      {/* Header */}
      <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <div className="flex flex-wrap items-center gap-2">
            <h1 className="text-lg font-semibold text-slate-800">3. 定标模块（服务费精确测算与交付物生成）</h1>
            <span className="rounded-md bg-[#EBF2F7] px-2 py-0.5 text-xs font-medium text-[#3B5B7E]">
              {terms.procurementType}
            </span>
            <span className="rounded-md bg-slate-100 px-2 py-0.5 text-xs font-medium text-slate-600">
              Excel 精确四舍五入算法（0 误差）
            </span>
          </div>
          <p className="mt-1 text-xs text-slate-500">
            选择业主单位自动匹配《客户协议清单》服务费折扣，一键生成预成交公示、成交结果公告、交接单及服务费计算表。
          </p>
        </div>
      </div>

      {/* Upstream Readiness Strip */}
      <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-4">
        {readinessItems.map((item) => (
          <Link
            key={item.label}
            to={item.link.startsWith('#') ? '/award' : item.link}
            className="flex items-start justify-between rounded-xl border border-slate-200 bg-white p-3.5 hover:border-[#3B5B7E]/40 transition-colors"
          >
            <div>
              <div className="flex items-center gap-1.5 text-xs font-semibold text-slate-700">
                {item.ok ? (
                  <CheckCircle2 className="h-3.5 w-3.5 text-[#2F6F5E]" />
                ) : (
                  <AlertCircle className="h-3.5 w-3.5 text-[#9A6B26]" />
                )}
                <span>{item.label}</span>
              </div>
              <div className="mt-1 text-xs text-slate-500 tabular-nums">{item.detail}</div>
            </div>
            <ArrowRight className="h-3.5 w-3.5 text-slate-300 mt-0.5" />
          </Link>
        ))}
      </div>

      {/* Owner Unit & Service Fee Config Panel */}
      <form
        id="fee-config"
        onSubmit={handleSaveConfig}
        className="rounded-xl border border-slate-200 bg-white p-5 space-y-4"
      >
        <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-100 pb-3">
          <div className="flex items-center gap-2">
            <Building2 className="h-4 w-4 text-[#3B5B7E]" />
            <h2 className="text-sm font-semibold text-slate-800">业主单位协议费率与项目类别配置</h2>
            {savingCfg && <span className="text-xs text-[#3B5B7E]">正在重算服务费...</span>}
            {cfgSavedMsg && <span className="text-xs text-[#2F6F5E] font-medium">{cfgSavedMsg}</span>}
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <label className="inline-flex cursor-pointer items-center gap-1.5 rounded-lg border border-slate-200 bg-slate-50 px-3 py-1.5 text-xs font-medium text-slate-600 hover:border-[#3B5B7E] hover:text-[#3B5B7E] transition-colors">
              <Upload className="h-3.5 w-3.5" />
              <span>{importingClients ? '导入中...' : '更新客户协议清单 (.xls/.xlsx)'}</span>
              <input
                type="file"
                accept=".xls,.xlsx"
                onChange={handleImportClients}
                className="hidden"
                disabled={importingClients}
              />
            </label>
            <label className="inline-flex cursor-pointer items-center gap-1.5 rounded-lg border border-slate-200 bg-slate-50 px-3 py-1.5 text-xs font-medium text-slate-600 hover:border-[#3B5B7E] hover:text-[#3B5B7E] transition-colors">
              <Upload className="h-3.5 w-3.5" />
              <span>{importingMaterials ? '导入中...' : '更新物料编码库 (.xlsx)'}</span>
              <input
                type="file"
                accept=".xls,.xlsx"
                onChange={handleImportMaterials}
                className="hidden"
                disabled={importingMaterials}
              />
            </label>
          </div>
        </div>

        <div className="grid grid-cols-1 gap-4 md:grid-cols-12">
          {/* Owner Unit Selector */}
          <div className="md:col-span-5">
            <label className="block text-xs font-medium text-slate-600 mb-1">
              选择业主单位（共 {data.clients.length} 家协议客户，自动定位折扣）
            </label>
            <SearchableClientSelect
              clients={data.clients}
              value={clientId}
              valueMode="id"
              placeholder="搜索并选择业主单位（未选默认不打折）..."
              onChange={(id) => handleClientChange(id)}
            />
            {selectedClientObj && (
              <div className="mt-1.5 flex flex-wrap items-center gap-2 text-[11px] text-slate-500">
                <span>联系人：{selectedClientObj.contact_name || '—'} {selectedClientObj.contact_phone}</span>
                <span>·</span>
                <span>协议到期：{selectedClientObj.contract_expiry || '—'}</span>
                {selectedClientObj.agreement_status && (
                  <span className="rounded bg-[#ECFDF5] px-1.5 py-0.5 text-[10px] text-[#2F6F5E]">
                    {selectedClientObj.agreement_status}
                  </span>
                )}
              </div>
            )}
          </div>

          {/* Category */}
          <div className="md:col-span-2">
            <label className="block text-xs font-medium text-slate-600 mb-1">项目类别</label>
            <select
              value={category}
              onChange={(e) => handleCategoryChange(e.target.value as '服务' | '工程' | '物资')}
              className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs font-medium text-slate-800 focus:border-[#3B5B7E] focus:outline-none"
            >
              {CATEGORY_OPTIONS.map((opt) => (
                <option key={opt.value} value={opt.value}>
                  {opt.label}
                </option>
              ))}
            </select>
          </div>

          {/* Discount Rate Override */}
          <div className="md:col-span-2">
            <label className="block text-xs font-medium text-slate-600 mb-1">服务费折扣系数</label>
            <input
              type="number"
              step="0.01"
              min="0.01"
              max="2"
              value={discountRate}
              onChange={(e) => setDiscountRate(e.target.value)}
              className="w-full rounded-lg border border-slate-200 px-3 py-2 text-xs font-mono font-semibold text-[#3B5B7E] focus:border-[#3B5B7E] focus:outline-none"
            />
            <div className="mt-1 text-[10px] text-slate-400">1=不打折，0.9=9折</div>
          </div>

          {/* Batch Manager & Location */}
          <div className="md:col-span-3 flex gap-2 items-start">
            <div className="flex-1">
              <label className="block text-xs font-medium text-slate-600 mb-1">批次经理</label>
              <input
                value={batchManager}
                onChange={(e) => setBatchManager(e.target.value)}
                placeholder="交接表用"
                className="w-full rounded-lg border border-slate-200 px-2.5 py-2 text-xs focus:border-[#3B5B7E] focus:outline-none"
              />
            </div>
            <div className="w-20">
              <label className="block text-xs font-medium text-slate-600 mb-1">开标地点</label>
              <input
                value={bidLocation}
                onChange={(e) => setBidLocation(e.target.value)}
                placeholder="华亿"
                className="w-full rounded-lg border border-slate-200 px-2.5 py-2 text-xs focus:border-[#3B5B7E] focus:outline-none"
              />
            </div>
            <div className="pt-5">
              <button
                type="submit"
                disabled={savingCfg}
                className="inline-flex items-center gap-1 rounded-lg bg-[#3B5B7E] px-3 py-2 text-xs font-medium text-white hover:bg-[#2E4866] disabled:opacity-50"
                title="保存并重算服务费"
              >
                <Save className="h-3.5 w-3.5" />
                <span>应用</span>
              </button>
            </div>
          </div>
        </div>
      </form>

      {/* Per-Package Award & Exact Service Fee Preview Table */}
      <div className="rounded-xl border border-slate-200 bg-white overflow-hidden">
        <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 bg-slate-50/70 px-4 py-3">
          <div className="flex items-center gap-2">
            <Calculator className="h-4 w-4 text-[#3B5B7E]" />
            <h2 className="text-sm font-semibold text-slate-700">
              各标包成交结果与服务费试算明细（{data.awardConfig.category}类 · {data.awardConfig.discount_label}）
            </h2>
          </div>
          <div className="flex flex-wrap items-center gap-3 text-xs text-slate-500 tabular-nums">
            <span>
              成交总额 <strong className="text-slate-800">{data.totals.totalAwardWan.toLocaleString()}</strong> 万元
            </span>
            <span>·</span>
            <span>
              折后服务费合计 <strong className="text-[#3B5B7E]">{data.totals.totalServiceFeeYuan.toLocaleString('zh-CN', { minimumFractionDigits: 2 })}</strong> 元
            </span>
            <span>·</span>
            <span>
              平台使用费合计 <strong className="text-slate-700">{data.totals.totalPlatformFeeYuan.toLocaleString('zh-CN', { minimumFractionDigits: 2 })}</strong> 元
            </span>
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full border-collapse text-left text-xs">
            <thead>
              <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                <th className="px-3 py-2.5 font-medium whitespace-nowrap">序号</th>
                <th className="px-3 py-2.5 font-medium whitespace-nowrap">分标 / 包号</th>
                <th className="px-3 py-2.5 font-medium min-w-[200px]">项目名称</th>
                <th className="px-3 py-2.5 font-medium whitespace-nowrap">成交厂家</th>
                <th className="px-3 py-2.5 font-medium whitespace-nowrap text-right">成交价格(万元)</th>
                <th className="px-3 py-2.5 font-medium whitespace-nowrap text-right">折前服务费(元)</th>
                <th className="px-3 py-2.5 font-medium whitespace-nowrap text-right">折后服务费(元)</th>
                <th className="px-3 py-2.5 font-medium whitespace-nowrap text-right">平台使用费(元)</th>
                <th className="px-3 py-2.5 font-medium whitespace-nowrap">联系人 / 电话</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {data.packages.length === 0 ? (
                <tr>
                  <td colSpan={9} className="py-10 text-center text-xs text-slate-400">
                    暂无标段成交数据，请先在【招标模块】与【开评标模块】导入本批次表格
                  </td>
                </tr>
              ) : (
                paginate(data.packages, pkgPage, pkgPageSize).map((p) => (
                  <tr key={`${p.lot_code}-${p.package_no}`} className={p.is_failed_lot ? 'bg-[#FEF2F2]/35' : 'hover:bg-slate-50/60'}>
                    <td className="px-3 py-2.5 tabular-nums text-slate-400">{p.seq}</td>
                    <td className="px-3 py-2.5 whitespace-nowrap">
                      <div className="font-mono font-medium text-slate-700">{p.lot_code || '—'}</div>
                      <div className="text-[11px] text-slate-400">{p.lot_name} · {p.package_no}</div>
                    </td>
                    <td className="px-3 py-2.5 font-medium text-slate-800">{p.project_name || '—'}</td>
                    <td className="px-3 py-2.5 whitespace-nowrap">
                      {p.is_failed_lot ? (
                        <span className="rounded bg-[#FEF2F2] px-2 py-0.5 text-[11px] font-medium text-[#A94442]">流标</span>
                      ) : (
                        <span className="text-slate-800">{p.winner_name || '待导入排序表'}</span>
                      )}
                    </td>
                    <td className="px-3 py-2.5 whitespace-nowrap text-right tabular-nums text-slate-700">
                      {p.is_failed_lot ? '/' : p.winner_price_wan !== null ? p.winner_price_wan : p.winner_price_raw || '—'}
                    </td>
                    <td className="px-3 py-2.5 whitespace-nowrap text-right tabular-nums text-slate-500">
                      {p.service_fee_pre_yuan !== null ? p.service_fee_pre_yuan.toFixed(2) : '/'}
                    </td>
                    <td className="px-3 py-2.5 whitespace-nowrap text-right tabular-nums font-semibold text-[#3B5B7E]">
                      {p.service_fee_yuan !== null ? p.service_fee_yuan.toFixed(2) : '/'}
                    </td>
                    <td className="px-3 py-2.5 whitespace-nowrap text-right tabular-nums text-slate-600">
                      {p.platform_fee_yuan !== null ? p.platform_fee_yuan.toFixed(2) : '/'}
                    </td>
                    <td className="px-3 py-2.5 whitespace-nowrap text-[11px] text-slate-500">
                      {p.is_failed_lot ? '/' : p.winner_contact ? `${p.winner_contact} ${p.winner_mobile}` : '—'}
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
        <PaginationBar
          total={data.packages.length}
          page={pkgPage}
          pageSize={pkgPageSize}
          onPageChange={setPkgPage}
          onPageSizeChange={setPkgPageSize}
        />
      </div>

      {/* Deliverable Generation Center */}
      <div className="space-y-4">
        <div className="space-y-2.5">
          <div className="flex items-center justify-between">
            <h2 className="text-sm font-semibold text-slate-700">公示与公告（Word 模板 · 自动替换标黄内容与表格）</h2>
            <span className="text-xs text-slate-400">固定合规条款正文原样保留</span>
          </div>
          <div className="grid grid-cols-1 gap-3.5 md:grid-cols-2">
            {wordTemplates.map((tpl) => {
              const isPub = data.batch.procurement_type === '公开招标';
              const isPublicity = tpl.key === 'pre_award_publicity';
              const colsState = isPublicity ? publicityCols : noticeCols;
              const setColsState = isPublicity ? setPublicityCols : setNoticeCols;
              const selectedColCount = 3 + Object.values(colsState).filter(Boolean).length;
              const cardTitle = isPublicity
                ? isPub
                  ? `推荐的中标候选人公示（公示期 3 日 · ${selectedColCount} 列表）`
                  : `预成交公示（公示期 1 日 · ${selectedColCount} 列表）`
                : tpl.key === 'award_result_notice'
                  ? isPub
                    ? `中标公告（公开招标 · ${selectedColCount} 列表）`
                    : `成交结果公告（非公开 · ${selectedColCount} 列表）`
                  : tpl.title;
              const cardDesc = isPublicity
                ? isPub
                  ? '公开招标专用模板：含《招标投标法实施条例》固定正文，表格固定含序号/包号/推荐中标人，下方可自由勾选可选列。'
                  : '非公开招标专用模板：含固定合规正文与异议联系方式，表格固定含序号/包号/推荐成交候选人，下方可自由勾选可选列。'
                : tpl.key === 'award_result_notice'
                  ? isPub
                    ? '公开招标专用中标公告模板：表格居中无底色，固定含序号/包号/中标候选人，下方可按项目单位要求自由勾选可选列。'
                    : '非公开招标专用成交结果公告模板：表格居中无底色，固定含序号/包号/成交人，下方可按项目单位要求自由勾选可选列。'
                  : tpl.description;
              return (
                <div key={tpl.key} className="flex flex-col justify-between rounded-xl border border-slate-200 bg-white p-4">
                  <div>
                    <div className="flex items-center justify-between gap-2">
                      <div className="flex items-center gap-2">
                        <FileText className="h-4 w-4 text-[#3B5B7E]" />
                        <h3 className="text-sm font-semibold text-slate-800">{cardTitle}</h3>
                      </div>
                      <span className="rounded bg-[#EBF2F7] px-2 py-0.5 text-[10px] font-mono uppercase text-[#3B5B7E]">
                        .docx
                      </span>
                    </div>
                    <p className="mt-1.5 text-xs text-slate-500 leading-relaxed">{cardDesc}</p>

                    {/* Persistent Date Constant Picker */}
                    <div className="mt-3 flex flex-wrap items-center justify-between gap-2 rounded-lg border border-slate-100 bg-slate-50/80 px-3 py-2">
                      <span className="text-[11px] font-medium text-slate-600">
                        {isPublicity ? '公示日期（批次常量）：' : '公告日期（同步成交挂网/通知书/推送日期）：'}
                      </span>
                      <input
                        type="date"
                        value={isPublicity ? publicityDate : announcementDate}
                        onChange={(e) =>
                          handleAwardDateChange(isPublicity ? 'publicity_date' : 'announcement_date', e.target.value)
                        }
                        className="rounded-md border border-slate-200 bg-white px-2.5 py-1 text-xs tabular-nums text-slate-700 focus:border-[#3B5B7E] focus:outline-none"
                      />
                    </div>

                    {/* Optional Table Columns Selector */}
                    <div className="mt-2 flex flex-wrap items-center gap-x-3.5 gap-y-1.5 rounded-lg border border-slate-100 bg-slate-50/80 px-3 py-2">
                      <span className="text-[11px] font-medium text-slate-500">表格可选列：</span>
                      {DOCX_OPTIONAL_COLUMNS.map((opt) => (
                        <label
                          key={opt.key}
                          className="inline-flex cursor-pointer items-center gap-1.5 text-xs text-slate-700 select-none"
                        >
                          <input
                            type="checkbox"
                            checked={colsState[opt.key]}
                            onChange={(e) =>
                              setColsState((prev) => ({ ...prev, [opt.key]: e.target.checked }))
                            }
                            className="h-3.5 w-3.5 rounded border-slate-300 text-[#3B5B7E] focus:ring-[#3B5B7E]"
                          />
                          <span>{opt.label}</span>
                        </label>
                      ))}
                    </div>
                  </div>
                  <div className="mt-3 pt-3 border-t border-slate-100 flex items-center justify-end gap-2">
                    <button
                      onClick={() => openGenerateModal(tpl)}
                      className="inline-flex items-center gap-1 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50 transition-colors"
                    >
                      <Sparkles className="h-3.5 w-3.5 text-[#3B5B7E]" />
                      <span>校对标黄字段</span>
                    </button>
                    <button
                      onClick={() => handleQuickExport(tpl)}
                      disabled={exportingKey === tpl.key}
                      className="inline-flex items-center gap-1.5 rounded-lg bg-[#3B5B7E] px-3.5 py-1.5 text-xs font-medium text-white hover:bg-[#2E4866] disabled:opacity-50 transition-colors"
                    >
                      <Download className="h-3.5 w-3.5" />
                      <span>{exportingKey === tpl.key ? '导出中...' : '直接导出 Word'}</span>
                    </button>
                  </div>
                </div>
              );
            })}
          </div>
        </div>

        <div className="space-y-2.5">
          <div className="flex items-center justify-between">
            <h2 className="text-sm font-semibold text-slate-700">交接单、成交结果表与服务费测算（Excel 模板）</h2>
            <span className="text-xs text-slate-400">服务费与原 Excel 表格保持分位一致</span>
          </div>
          <div className="grid grid-cols-1 gap-3.5 md:grid-cols-2">
            {excelTemplates.map((tpl) => (
              <div key={tpl.key} className="flex flex-col justify-between rounded-xl border border-slate-200 bg-white p-4">
                <div>
                  <div className="flex items-center justify-between gap-2">
                    <div className="flex items-center gap-2">
                      <FileSpreadsheet className="h-4 w-4 text-[#2F6F5E]" />
                      <h3 className="text-sm font-semibold text-slate-800">{tpl.title}</h3>
                    </div>
                    <span className="rounded bg-[#ECFDF5] px-2 py-0.5 text-[10px] font-mono uppercase text-[#2F6F5E]">
                      .xlsx
                    </span>
                  </div>
                  <p className="mt-1.5 text-xs text-slate-500 leading-relaxed">{tpl.description}</p>
                </div>
                <div className="mt-3 pt-3 border-t border-slate-100 flex items-center justify-end gap-2">
                  <button
                    onClick={() => openGenerateModal(tpl)}
                    className="inline-flex items-center gap-1 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50 transition-colors"
                  >
                    <Sparkles className="h-3.5 w-3.5 text-[#3B5B7E]" />
                    <span>核对参数</span>
                  </button>
                  <button
                    onClick={() => handleQuickExport(tpl)}
                    disabled={exportingKey === tpl.key}
                    className="inline-flex items-center gap-1.5 rounded-lg bg-[#2F6F5E] px-3.5 py-1.5 text-xs font-medium text-white hover:bg-[#25594B] disabled:opacity-50 transition-colors"
                  >
                    <Download className="h-3.5 w-3.5" />
                    <span>{exportingKey === tpl.key ? '导出中...' : '直接导出 Excel'}</span>
                  </button>
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* Generated Outputs History */}
      {(() => {
        const totalOutputs = data.outputs.length;
        const totalOutputPages = Math.max(1, Math.ceil(totalOutputs / OUTPUT_PAGE_SIZE));
        const safeOutputPage = Math.min(outputPage, totalOutputPages);
        const startIdx = (safeOutputPage - 1) * OUTPUT_PAGE_SIZE;
        const pageOutputs = data.outputs.slice(startIdx, startIdx + OUTPUT_PAGE_SIZE);
        const allPageSelected = pageOutputs.length > 0 && pageOutputs.every((o) => selectedOutputIds.has(o.id));
        const allSelected = totalOutputs > 0 && data.outputs.every((o) => selectedOutputIds.has(o.id));

        return (
          <div id="award-outputs" className="rounded-xl border border-slate-200 bg-white p-5 space-y-3">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <div className="flex flex-wrap items-center gap-3">
                <h2 className="text-sm font-semibold text-slate-700">
                  已生成的定标交付文件记录 ({totalOutputs})
                </h2>
                {totalOutputs > 0 && (
                  <label className="inline-flex cursor-pointer items-center gap-1.5 text-xs text-slate-600 select-none">
                    <input
                      type="checkbox"
                      checked={allPageSelected}
                      onChange={(e) => {
                        setSelectedOutputIds((prev) => {
                          const next = new Set(prev);
                          if (e.target.checked) pageOutputs.forEach((o) => next.add(o.id));
                          else pageOutputs.forEach((o) => next.delete(o.id));
                          return next;
                        });
                      }}
                      className="h-3.5 w-3.5 rounded border-slate-300 text-[#3B5B7E] focus:ring-[#3B5B7E]"
                    />
                    <span>全选本页</span>
                  </label>
                )}
                {totalOutputs > OUTPUT_PAGE_SIZE && (
                  <button
                    type="button"
                    onClick={() => {
                      if (allSelected) setSelectedOutputIds(new Set());
                      else setSelectedOutputIds(new Set(data.outputs.map((o) => o.id)));
                    }}
                    className="text-xs font-medium text-[#3B5B7E] hover:underline"
                  >
                    {allSelected ? '取消全选全部' : `全选全部 (${totalOutputs})`}
                  </button>
                )}
              </div>
              <div className="flex flex-wrap items-center gap-2">
                {selectedOutputIds.size > 0 ? (
                  <button
                    type="button"
                    disabled={batchDeletingOutputs}
                    onClick={handleBatchDeleteOutputs}
                    className="inline-flex items-center gap-1 rounded-lg border border-[#A94442]/30 bg-[#FEF2F2] px-3 py-1.5 text-xs font-medium text-[#A94442] hover:bg-[#FEE2E2] disabled:opacity-50 transition-colors"
                  >
                    <Trash2 className="h-3.5 w-3.5" />
                    <span>{batchDeletingOutputs ? '删除中...' : `批量删除 (${selectedOutputIds.size})`}</span>
                  </button>
                ) : (
                  <span className="text-xs text-slate-400">生成后自动同步至文件库，可在【资料整理】中直接关联归档</span>
                )}
              </div>
            </div>

            <div className="divide-y divide-slate-100 rounded-lg border border-slate-200">
              {totalOutputs === 0 ? (
                <div className="py-8 text-center text-xs text-slate-400">
                  本批次暂未生成定标交付文件，请点击上方卡片一键生成
                </div>
              ) : (
                pageOutputs.map((out) => {
                  const checked = selectedOutputIds.has(out.id);
                  return (
                    <div
                      key={out.id}
                      className={`flex flex-col gap-2 p-3.5 sm:flex-row sm:items-center sm:justify-between transition-colors ${
                        checked ? 'bg-[#EBF2F7]/40' : ''
                      }`}
                    >
                      <div className="flex items-center gap-3 min-w-0">
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => toggleSelectOutput(out.id)}
                          className="h-3.5 w-3.5 shrink-0 rounded border-slate-300 text-[#3B5B7E] focus:ring-[#3B5B7E] cursor-pointer"
                        />
                        {out.file_type === 'xlsx' ? (
                          <FileSpreadsheet className="h-5 w-5 shrink-0 text-[#2F6F5E]" />
                        ) : (
                          <FileText className="h-5 w-5 shrink-0 text-[#3B5B7E]" />
                        )}
                        <div className="min-w-0">
                          <div className="truncate text-sm font-medium text-slate-800">{out.filename}</div>
                          <div className="mt-0.5 text-xs text-slate-400 tabular-nums">
                            版本 v{out.version} · 生成于 {out.created_at.slice(0, 16).replace('T', ' ')}
                          </div>
                        </div>
                      </div>
                      <div className="flex items-center gap-2 shrink-0">
                        <a
                          href={`/api/award/outputs/${out.id}/download`}
                          className="inline-flex items-center gap-1 rounded-lg border border-slate-200 px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50"
                        >
                          <Download className="h-3.5 w-3.5" /> 下载文件
                        </a>
                        <button
                          onClick={() => handleDeleteOutput(out.id)}
                          className="rounded-lg p-1.5 text-slate-400 hover:bg-red-50 hover:text-[#A94442]"
                          title="删除记录"
                        >
                          <Trash2 className="h-3.5 w-3.5" />
                        </button>
                      </div>
                    </div>
                  );
                })
              )}
            </div>

            {totalOutputs > OUTPUT_PAGE_SIZE && (
              <div className="flex flex-wrap items-center justify-between gap-2 pt-1 text-xs text-slate-500 tabular-nums">
                <span>
                  显示第 {startIdx + 1}–{Math.min(startIdx + OUTPUT_PAGE_SIZE, totalOutputs)} 份，共 {totalOutputs} 份（每页 {OUTPUT_PAGE_SIZE} 份）
                </span>
                <div className="flex items-center gap-1.5">
                  <button
                    type="button"
                    disabled={safeOutputPage <= 1}
                    onClick={() => setOutputPage(safeOutputPage - 1)}
                    className="rounded-md border border-slate-200 bg-white px-2.5 py-1 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:opacity-40"
                  >
                    上一页
                  </button>
                  <span className="px-2 text-slate-600">
                    第 {safeOutputPage} / {totalOutputPages} 页
                  </span>
                  <button
                    type="button"
                    disabled={safeOutputPage >= totalOutputPages}
                    onClick={() => setOutputPage(safeOutputPage + 1)}
                    className="rounded-md border border-slate-200 bg-white px-2.5 py-1 text-xs font-medium text-slate-600 hover:bg-slate-50 disabled:opacity-40"
                  >
                    下一页
                  </button>
                </div>
              </div>
            )}
          </div>
        );
      })()}

      {/* Pre-Generation Field Verification Modal */}
      {activeTemplate && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4 backdrop-blur-[1px]">
          <div className="flex max-h-[85vh] w-full max-w-lg flex-col overflow-hidden rounded-xl border border-slate-200 bg-white shadow-xl">
            <div className="flex items-center justify-between border-b border-slate-200 px-5 py-3.5">
              <div>
                <h3 className="text-sm font-semibold text-slate-800">生成前校对 — {activeTemplate.title}</h3>
                <p className="text-xs text-slate-400 mt-0.5">已自动填入上游数据与标黄字段，确认无误即可生成</p>
              </div>
              <button onClick={() => setActiveTemplate(null)} className="rounded-lg p-1 text-slate-400 hover:bg-slate-100">
                <X className="h-4 w-4" />
              </button>
            </div>

            <form onSubmit={handleGenerate} className="flex-1 overflow-y-auto p-5 space-y-3">
              <div>
                <label className="block text-xs font-medium text-slate-600 mb-1">批次名称</label>
                <input
                  value={overrides.batchName || ''}
                  onChange={(e) => setOverrides({ ...overrides, batchName: e.target.value })}
                  className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1">采购 / 项目编号</label>
                  <input
                    value={overrides.batchCode || ''}
                    onChange={(e) => setOverrides({ ...overrides, batchCode: e.target.value })}
                    className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-xs font-mono focus:border-[#3B5B7E] focus:outline-none"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1">业主单位 / 采购人</label>
                  <input
                    value={overrides.ownerUnit || ''}
                    onChange={(e) => setOverrides({ ...overrides, ownerUnit: e.target.value })}
                    className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                  />
                </div>
              </div>

              {activeTemplate.key === 'pre_award_publicity' && (
                <>
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-xs font-medium text-slate-600 mb-1">公示期（日）</label>
                      <input
                        value={overrides.publicityDays || '1'}
                        onChange={(e) => setOverrides({ ...overrides, publicityDays: e.target.value })}
                        className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-slate-600 mb-1">落款日期</label>
                      <input
                        type="date"
                        value={overrides.publicityDate || ''}
                        onChange={(e) => setOverrides({ ...overrides, publicityDate: e.target.value })}
                        className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                      />
                    </div>
                  </div>
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-xs font-medium text-slate-600 mb-1">异议联系邮箱</label>
                      <input
                        value={overrides.objectionEmail || ''}
                        onChange={(e) => setOverrides({ ...overrides, objectionEmail: e.target.value })}
                        className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-slate-600 mb-1">异议联系电话</label>
                      <input
                        value={overrides.objectionPhone || ''}
                        onChange={(e) => setOverrides({ ...overrides, objectionPhone: e.target.value })}
                        className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                      />
                    </div>
                  </div>
                </>
              )}

              {activeTemplate.key === 'award_result_notice' && (
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">代理机构名称</label>
                    <input
                      value={overrides.agencyName || ''}
                      onChange={(e) => setOverrides({ ...overrides, agencyName: e.target.value })}
                      className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">公告落款日期</label>
                    <input
                      type="date"
                      value={overrides.announcementDate || ''}
                      onChange={(e) => setOverrides({ ...overrides, announcementDate: e.target.value })}
                      className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                    />
                  </div>
                </div>
              )}

              {(activeTemplate.key === 'pre_award_publicity' || activeTemplate.key === 'award_result_notice') && (
                <div className="rounded-lg border border-slate-200 bg-slate-50/70 p-3 space-y-2">
                  <div className="text-xs font-medium text-slate-600">
                    表格可选列（固定包含：序号、包号、候选人/成交人）
                  </div>
                  <div className="flex flex-wrap gap-4">
                    {DOCX_OPTIONAL_COLUMNS.map((opt) => (
                      <label
                        key={opt.key}
                        className="inline-flex cursor-pointer items-center gap-1.5 text-xs text-slate-700 select-none"
                      >
                        <input
                          type="checkbox"
                          checked={overrides[opt.key] === 'true'}
                          onChange={(e) => {
                            const checked = e.target.checked;
                            if (activeTemplate.key === 'pre_award_publicity') {
                              setPublicityCols((prev) => ({ ...prev, [opt.key]: checked }));
                            } else {
                              setNoticeCols((prev) => ({ ...prev, [opt.key]: checked }));
                            }
                            setOverrides({ ...overrides, [opt.key]: checked ? 'true' : 'false' });
                          }}
                          className="h-3.5 w-3.5 rounded border-slate-300 text-[#3B5B7E] focus:ring-[#3B5B7E]"
                        />
                        <span>{opt.label}</span>
                      </label>
                    ))}
                  </div>
                </div>
              )}

              {(activeTemplate.key === 'notice_handover' || activeTemplate.key === 'project_handover') && (
                <div className="space-y-3">
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">实际开标时间</label>
                    <DateTimePickerField
                      compact
                      value={overrides.bidOpeningTime || ''}
                      onChange={(val) => setOverrides({ ...overrides, bidOpeningTime: val })}
                      allowTime
                    />
                  </div>
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-xs font-medium text-slate-600 mb-1">成交通知书日期（默认同结果公告）</label>
                      <input
                        type="date"
                        value={overrides.noticeDate || ''}
                        onChange={(e) => setOverrides({ ...overrides, noticeDate: e.target.value })}
                        className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                      />
                    </div>
                    {activeTemplate.key === 'project_handover' && (
                      <div>
                        <label className="block text-xs font-medium text-slate-600 mb-1">成交挂网 / 推送日期</label>
                        <input
                          type="date"
                          value={overrides.announcementDate || ''}
                          onChange={(e) =>
                            setOverrides({
                              ...overrides,
                              announcementDate: e.target.value,
                              noticePushDate: e.target.value,
                            })
                          }
                          className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                        />
                      </div>
                    )}
                  </div>
                </div>
              )}

              {activeTemplate.key === 'project_handover' && (
                <div className="space-y-3">
                  <div>
                    <label className="block text-xs font-medium text-slate-600 mb-1">匹配时间（默认自动取挂网时间）</label>
                    <DateTimePickerField
                      compact
                      value={overrides.matchDate || ''}
                      onChange={(val) => setOverrides({ ...overrides, matchDate: val })}
                      allowTime
                    />
                  </div>
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-xs font-medium text-slate-600 mb-1">批次经理</label>
                      <input
                        value={overrides.batchManager || ''}
                        onChange={(e) => setOverrides({ ...overrides, batchManager: e.target.value })}
                        className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-slate-600 mb-1">开标地点</label>
                      <input
                        value={overrides.bidLocation || '华亿'}
                        onChange={(e) => setOverrides({ ...overrides, bidLocation: e.target.value })}
                        className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                      />
                    </div>
                  </div>
                </div>
              )}

              <div className="rounded-lg bg-slate-50 px-3 py-2 text-[11px] text-slate-500">
                当前服务费计算基准：<strong>{category}类</strong> · 折扣系数 <strong>{discountRate}</strong> · 共 {data.packages.length} 个标包（成交 {data.totals.winningCount} / 流标 {data.totals.failedCount}）
              </div>

              <div className="flex justify-end gap-2 pt-2">
                <button
                  type="button"
                  onClick={() => setActiveTemplate(null)}
                  className="rounded-lg border border-slate-200 px-4 py-2 text-xs font-medium text-slate-600 hover:bg-slate-50"
                >
                  取消
                </button>
                <button
                  type="submit"
                  disabled={generating}
                  className="inline-flex items-center gap-1.5 rounded-lg bg-[#3B5B7E] px-4 py-2 text-xs font-medium text-white hover:bg-[#2E4866] disabled:opacity-50"
                >
                  <Sparkles className="h-3.5 w-3.5" />
                  <span>{generating ? '生成中...' : `确认生成 .${activeTemplate.file_type}`}</span>
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}
