import React, { useCallback, useEffect, useState } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import {
  CheckCircle2,
  Circle,
  AlertTriangle,
  Info,
  Plus,
  Trash2,
  Upload,
  FileSpreadsheet,
  Eye,
  Download,
  BookmarkCheck,
  ShieldAlert,
  ChevronDown,
  ChevronUp,
  Search,
  Trophy,
  XCircle,
  ClipboardCheck,
  X,
  Scale,
  Cpu,
  Users,
  FileText,
  Calculator,
} from 'lucide-react';
import { api } from '../api/client';
import { AwardAllocPanel } from '../components/AwardAllocPanel';
import { ChangeNoticeBanner } from '../components/ChangeNoticeBanner';
import { EconomicRankPanel } from '../components/EconomicRankPanel';
import { ExcelPreviewModal } from '../components/ExcelPreviewModal';
import { ExpertScoreCheckPanel } from '../components/ExpertScoreCheckPanel';
import { ExpertScorePrecheckPanel } from '../components/ExpertScorePrecheckPanel';
import { InitialMinutesPanel } from '../components/InitialMinutesPanel';
import { PerformanceScorePanel } from '../components/PerformanceScorePanel';
import { PriceScorePanel } from '../components/PriceScorePanel';
import { SopGuidesPanel } from '../components/SopGuidesPanel';
import { SplitPackPanel } from '../components/SplitPackPanel';
import { InitialReportPanel } from '../components/InitialReportPanel';
import { WarningFormPanel } from '../components/WarningFormPanel';
import { paginate, PaginationBar, usePagination } from '../components/PaginationBar';
import { useBatch } from '../context/BatchContext';
import type { EvaluationWorkspace, ParsedSheetTable, SopItem } from '../types';
import { extractPackageNumber } from '../utils/projectSort';
import { getTerminology } from '../utils/terminology';

const STAGE_GROUPS = ['开标前准备', '内网开标', '评审过程', '评审结束与资料带出'];
const OTHER_CATEGORIES = ['评审报告', '应答文件上传情况统计', '澄清/多轮报价截图', '会议签到表', '其他阶段附件'];
const PROOF_RATE_OPTIONS = [0, 1, 3, 5, 6, 9, 13];

const REPORT_ROADMAP_ITEMS = [
  {
    key: 'legal_report',
    title: '法律报告生成',
    badge: '待接入模板',
    desc: '围绕否决投标依据、异议处理及程序合规节点，自动套用批次与标段数据生成现场法律意见书/法律报告。',
  },
  {
    key: 'supervision_report',
    title: '现场监督报告生成',
    badge: '待接入模板',
    desc: '汇总开评标时间节点、响应家数、否决与流标情况及异常核查记录，一键生成现场监督报告。',
  },
  {
    key: 'eval_report',
    title: '评审报告生成',
    badge: '待接入模板',
    desc: '整合各标段总分排序、基准价、候选人推荐与否决说明，自动生成最终版评标/评审报告。',
  },
];

const pinyinCollator = new Intl.Collator('zh-Hans-CN-u-co-pinyin', { numeric: true, sensitivity: 'base' });

type ActiveTab = 'results' | 'quote_tax' | 'collusion' | 'rejections' | 'tools' | 'reports' | 'sop';
type QuoteTaxSubTab = 'over_limit' | 'low_quote' | 'round_increase' | 'same_quote' | 'tax_abnormal' | 'tax_missing' | 'tax_proof';
type CollusionSubTab = 'contact' | 'hardware';
type ToolFocusKey = 'score_precheck' | 'score_check' | 'price_score' | 'performance_score' | 'economic_rank' | 'award_alloc' | 'split_pack';
type ReportFocusKey = 'warning_form' | 'initial_report' | 'initial_minutes' | 'roadmap';

function resolveTabFromSection(section?: string): ActiveTab {
  switch (section) {
    case 'quote-tax':
      return 'quote_tax';
    case 'collusion':
      return 'collusion';
    case 'rejections':
      return 'rejections';
    case 'tools':
      return 'tools';
    case 'reports':
      return 'reports';
    case 'sop':
      return 'sop';
    case 'results':
    default:
      return 'results';
  }
}

export function EvaluationPage() {
  const { section, tool } = useParams<{ section?: string; tool?: string }>();
  const navigate = useNavigate();
  const { currentBatchId, refreshBatches } = useBatch();
  const [ws, setWs] = useState<EvaluationWorkspace | null>(null);
  const activeTab: ActiveTab = resolveTabFromSection(section);
  const [slotsOpen, setSlotsOpen] = useState<boolean>(activeTab === 'results');
  const [activeToolKey, setActiveToolKey] = useState<ToolFocusKey>('score_check');
  const [activeReportKey, setActiveReportKey] = useState<ReportFocusKey>('warning_form');
  const [expandedPkgKeys, setExpandedPkgKeys] = useState<Set<string>>(new Set());

  const setActiveTab = useCallback(
    (next: ActiveTab) => {
      const slug = next === 'quote_tax' ? 'quote-tax' : next;
      navigate(`/evaluation/${slug}`);
    },
    [navigate]
  );

  useEffect(() => {
    setSlotsOpen(activeTab === 'results');
  }, [activeTab]);

  useEffect(() => {
    if (
      activeTab === 'tools' &&
      (tool === 'score_precheck' ||
        tool === 'score_check' ||
        tool === 'price_score' ||
        tool === 'performance_score' ||
        tool === 'economic_rank' ||
        tool === 'award_alloc' ||
        tool === 'split_pack')
    ) {
      setActiveToolKey(tool);
    }
    if (activeTab === 'reports' && (tool === 'warning_form' || tool === 'initial_report' || tool === 'initial_minutes' || tool === 'roadmap')) {
      setActiveReportKey(tool);
    }
  }, [activeTab, tool]);

  // Results & Rejections filter + Pagination
  const [detailCountFilter, setDetailCountFilter] = useState<'all' | 'lt3' | 'lt2' | 'eq0' | 'failed'>('all');
  const [pkgSearch, setPkgSearch] = useState('');
  const [pkgPage, setPkgPage] = useState(1);
  const [pkgPageSize, setPkgPageSize] = useState(20);
  const [rejFilter, setRejFilter] = useState<'all' | '商务' | '技术'>('all');
  const [rejSearch, setRejSearch] = useState('');

  // Quote & Tax Audit sub-tab + config saving
  const [qtSubTab, setQtSubTab] = useState<QuoteTaxSubTab>('over_limit');
  const [qtSearch, setQtSearch] = useState('');
  const [savingAuditCfg, setSavingAuditCfg] = useState(false);

  // Collusion sub-tab + search
  const [collusionSubTab, setCollusionSubTab] = useState<CollusionSubTab>('contact');
  const [collusionSearch, setCollusionSearch] = useState('');

  // Unified 20/50/全部 pagination across all sub-modules
  const qtPager = usePagination(`${currentBatchId}-${qtSubTab}-${qtSearch}`);
  const contactOverlapPager = usePagination(`${currentBatchId}-co-${collusionSearch}`);
  const contactMissingPager = usePagination(`${currentBatchId}-cm-${collusionSearch}`);
  const hwAlertPager = usePagination(`${currentBatchId}-ha-${collusionSearch}`);
  const hwManualPager = usePagination(`${currentBatchId}-hm-${collusionSearch}`);
  const hwInsuffPager = usePagination(`${currentBatchId}-hi-${collusionSearch}`);
  const rejPager = usePagination(`${currentBatchId}-${rejFilter}-${rejSearch}`);

  // Upload & Cross-batch mismatch confirmation
  const [uploadingSlot, setUploadingSlot] = useState<string | null>(null);
  const [lastReport, setLastReport] = useState<{ slotLabel: string; summary: string; warnings: string[] } | null>(null);
  const [mismatchModal, setMismatchModal] = useState<{
    slotKey: string;
    file: File;
    currentBatchCode: string;
    fileBatchCode: string;
    error: string;
  } | null>(null);
  const [previewData, setPreviewData] = useState<{ title: string; sheets: ParsedSheetTable[] } | null>(null);

  // SOP controls
  const [savedTplMsg, setSavedTplMsg] = useState('');
  const [newStepStage, setNewStepStage] = useState(STAGE_GROUPS[0]);
  const [newStepTitle, setNewStepTitle] = useState('');
  const [newStepDesc, setNewStepDesc] = useState('');
  const [syncNewStepToTemplate, setSyncNewStepToTemplate] = useState(true);

  const [newNoticePriority, setNewNoticePriority] = useState<'critical' | 'normal'>('critical');
  const [newNoticeTitle, setNewNoticeTitle] = useState('');
  const [newNoticeDesc, setNewNoticeDesc] = useState('');

  const [otherCat, setOtherCat] = useState(OTHER_CATEGORIES[0]);
  const [uploadingOther, setUploadingOther] = useState(false);

  const loadData = useCallback(async () => {
    if (!currentBatchId) return;
    try {
      const data = await api.getEvaluationWorkspace(currentBatchId);
      setWs(data);
    } catch {
      // ignore if offline
    }
  }, [currentBatchId]);

  useEffect(() => {
    loadData();
    setLastReport(null);
    setPkgPage(1);
  }, [loadData]);

  useEffect(() => {
    setPkgPage(1);
  }, [detailCountFilter, pkgSearch]);

  const batch = ws?.batch;
  const terms = getTerminology(batch?.procurement_type, batch?.unit_nature);
  const isPublicBidding = terms.procurementType === '公开招标';

  const sopItems = ws?.sopItems || [];
  const steps = sopItems.filter((i) => i.item_type === 'step');
  const notices = sopItems.filter((i) => i.item_type === 'notice');
  const doneCount = steps.filter((s) => s.is_completed === 1).length;
  const sopPct = steps.length ? Math.round((doneCount / steps.length) * 100) : 0;

  const qtAudit = ws?.quoteTaxAudit;
  const contactAudit = ws?.contactAudit;
  const hardwareAudit = ws?.hardwareAudit;
  const taxCategory = ws?.auditConfig?.tax_category || '服务';
  const proofRates = ws?.auditConfig?.proof_rates || [0, 1];

  const collusionTotalAlerts =
    (contactAudit?.counts.overlapCount || 0) +
    (hardwareAudit?.counts.alertCount || 0) +
    (hardwareAudit?.counts.manualReviewCount || 0);

  // Sort packages by 分标名称 Pinyin + 包号 integer
  const sortedPackages = [...(ws?.packages || [])].sort((a, b) => {
    const lotCmp = pinyinCollator.compare(a.lot_name || a.lot_code, b.lot_name || b.lot_code);
    if (lotCmp !== 0) return lotCmp;
    return extractPackageNumber(a.package_no) - extractPackageNumber(b.package_no);
  });

  const filteredRejections = (ws?.rejections || []).filter((r) => {
    if (rejFilter !== 'all' && !r.review_type.includes(rejFilter)) return false;
    if (rejSearch.trim()) {
      const q = rejSearch.trim().toLowerCase();
      return (
        r.supplier_name.toLowerCase().includes(q) ||
        r.package_no.toLowerCase().includes(q) ||
        r.reason_category.toLowerCase().includes(q) ||
        r.bid_response.toLowerCase().includes(q)
      );
    }
    return true;
  });

  function toggleExpandPkg(key: string) {
    setExpandedPkgKeys((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  }

  async function handleUpdateAuditConfig(nextCategory: '服务' | '物资', nextProofRates: number[]) {
    if (!currentBatchId) return;
    setSavingAuditCfg(true);
    try {
      await api.saveEvaluationAuditConfig(currentBatchId, {
        tax_category: nextCategory,
        proof_rates: nextProofRates,
      });
      await loadData();
    } catch (err) {
      alert((err as Error).message);
    } finally {
      setSavingAuditCfg(false);
    }
  }

  function handleToggleProofRate(rate: number) {
    const exists = proofRates.includes(rate);
    const next = exists ? proofRates.filter((r) => r !== rate) : [...proofRates, rate].sort((a, b) => a - b);
    handleUpdateAuditConfig(taxCategory, next);
  }

  async function doSlotUpload(slotKey: string, file: File, force = false) {
    if (!currentBatchId) return;
    setUploadingSlot(slotKey);
    try {
      const res = await api.uploadEvaluationSlot(currentBatchId, slotKey, file, force);
      if ('mismatch' in res && res.mismatch) {
        setMismatchModal({
          slotKey,
          file,
          currentBatchCode: res.currentBatchCode,
          fileBatchCode: res.fileBatchCode,
          error: res.error,
        });
        return;
      }
      if ('report' in res) {
        setLastReport(res.report);
        await loadData();
        await refreshBatches();
        if (slotKey === 'rejections') setActiveTab('rejections');
        else if (slotKey === 'ip_hardware') {
          setActiveTab('collusion');
          setCollusionSubTab('hardware');
        } else {
          setActiveTab('results');
        }
      }
    } catch (err) {
      alert((err as Error).message);
    } finally {
      setUploadingSlot(null);
    }
  }

  function handleSlotSelect(slotKey: string, e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    e.target.value = '';
    doSlotUpload(slotKey, file, false);
  }

  async function handleConfirmForceUpload() {
    if (!mismatchModal) return;
    const { slotKey, file } = mismatchModal;
    setMismatchModal(null);
    await doSlotUpload(slotKey, file, true);
  }

  async function handleClearSlot(slotKey: string, slotLabel: string) {
    if (!currentBatchId) return;
    if (!window.confirm(`确认清除【${slotLabel}】已导入的全部解析数据与原表文件吗？`)) return;
    setUploadingSlot(slotKey);
    try {
      await api.clearEvaluationSlot(currentBatchId, slotKey);
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

  async function toggleStep(item: SopItem) {
    await api.updateSop(item.id, { is_completed: item.is_completed ? 0 : 1 });
    await loadData();
    await refreshBatches();
  }

  async function updateStepNotes(item: SopItem, notes: string) {
    await api.updateSop(item.id, { notes });
    await loadData();
  }

  async function handleAddStep(e: React.FormEvent) {
    e.preventDefault();
    if (!currentBatchId || !newStepTitle.trim()) return;
    await api.createSop(currentBatchId, {
      item_type: 'step',
      stage_group: newStepStage,
      title: newStepTitle.trim(),
      description: newStepDesc.trim(),
      sync_to_template: syncNewStepToTemplate,
    });
    setNewStepTitle('');
    setNewStepDesc('');
    await loadData();
    await refreshBatches();
  }

  async function handleAddNotice(e: React.FormEvent) {
    e.preventDefault();
    if (!currentBatchId || !newNoticeTitle.trim()) return;
    await api.createSop(currentBatchId, {
      item_type: 'notice',
      stage_group: '评审过程',
      priority: newNoticePriority,
      title: newNoticeTitle.trim(),
      description: newNoticeDesc.trim(),
      sync_to_template: true,
    });
    setNewNoticeTitle('');
    setNewNoticeDesc('');
    await loadData();
  }

  async function handleDeleteSop(id: string) {
    await api.deleteSop(id);
    await loadData();
    await refreshBatches();
  }

  async function handleSaveAsTemplate() {
    if (!currentBatchId) return;
    const res = await api.saveSopAsTemplate(currentBatchId);
    setSavedTplMsg(`已将 ${res.count} 条同步为全局默认模板`);
    setTimeout(() => setSavedTplMsg(''), 2500);
  }

  async function handleOtherUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file || !currentBatchId) return;
    setUploadingOther(true);
    try {
      await api.uploadFile(currentBatchId, file, 'evaluation', otherCat);
      await loadData();
      await refreshBatches();
    } finally {
      setUploadingOther(false);
      e.target.value = '';
    }
  }

  async function handleDeleteFile(id: string) {
    await api.deleteFile(id);
    await loadData();
    await refreshBatches();
  }

  const slots = ws?.slots || [];
  const failedCount = sortedPackages.filter((p) => p.is_failed_lot).length;
  const recommendedCount = sortedPackages.filter((p) => !p.is_failed_lot && p.first_candidate).length;
  const lt3Count = sortedPackages.filter((p) => p.bidder_count > 0 && p.detailed_eval_count < 3).length;
  const lt2Count = sortedPackages.filter((p) => p.bidder_count > 0 && p.detailed_eval_count < 2).length;
  const eq0Count = sortedPackages.filter((p) => p.bidder_count > 0 && p.detailed_eval_count === 0).length;

  // Flatten over-limit items for the Quote & Tax Audit sub-module (supports both 整包 and 合并子包 across all rounds)
  const overLimitRows = sortedPackages.flatMap((pkg) =>
    (pkg.supplier_bids || []).flatMap((sb) => {
      if (!sb.is_over_limit) return [];
      const details = sb.over_limit_details || [];
      if (details.length > 0) {
        return details.map((d) => ({
          scope: d.scope,
          round_label: d.round_label || '第1轮',
          lot_code: pkg.lot_code,
          lot_name: pkg.lot_name || pkg.lot_code,
          package_no: pkg.package_no,
          req_app_no: d.req_app_no || pkg.req_app_no,
          sub_name: d.sub_name || pkg.project_name,
          supplier_name: sb.supplier_name,
          limit_display: d.limit_display || '—',
          quote_display: d.quote_display || sb.final_price || '—',
          exceed_text: d.exceed_text || '—',
          is_rejected: sb.is_rejected,
          rejection_reason: sb.rejection_reason,
          msg: d.msg,
        }));
      }
      return [
        {
          scope: '整包' as const,
          round_label: sb.final_round || '第1轮',
          lot_code: pkg.lot_code,
          lot_name: pkg.lot_name || pkg.lot_code,
          package_no: pkg.package_no,
          req_app_no: pkg.req_app_no,
          sub_name: pkg.project_name,
          supplier_name: sb.supplier_name,
          limit_display: pkg.max_price_raw || (pkg.max_price_yuan !== null ? `${Number((pkg.max_price_yuan / 10000).toFixed(4))} 万元` : '—'),
          quote_display: sb.final_price || '—',
          exceed_text: '超限',
          is_rejected: sb.is_rejected,
          rejection_reason: sb.rejection_reason,
          msg: sb.over_limit_msg || '超过最高限价',
        },
      ];
    })
  );
  const overLimitTotalSuppliers = sortedPackages.reduce((s, p) => s + (p.over_limit_count || 0), 0);

  const qtTotalIssues =
    overLimitRows.length +
    (qtAudit?.counts.lowQuote || 0) +
    (qtAudit?.counts.taxAbnormal || 0) +
    (qtAudit?.counts.taxMissing || 0) +
    (qtAudit?.counts.taxProof || 0) +
    (qtAudit?.counts.sameQuote || 0) +
    (qtAudit?.counts.roundIncrease || 0);

  // Quick lookup sets for colored risk tags in 开标与行信息报价明细
  const flagKey = (lotCode: string, pkgNo: string, supName: string) => `${lotCode}::${pkgNo}::${supName}`;
  const changeNotices = ws?.changeNotices || [];
  const pkgChangeNoticeMap = new Map<string, typeof changeNotices>();
  changeNotices.forEach((cn) => {
    (cn.affected_packages || []).forEach((ap) => {
      const k = `${ap.lot_code || ''}::${ap.package_no || ''}`;
      const list = pkgChangeNoticeMap.get(k) || [];
      if (!list.some((x) => x.id === cn.id)) list.push(cn);
      pkgChangeNoticeMap.set(k, list);
    });
  });

  const lowQuoteSet = new Set((qtAudit?.lowQuotes || []).map((x) => flagKey(x.lot_code, x.package_no, x.supplier_name)));
  const roundIncreaseSet = new Set((qtAudit?.roundIncreases || []).map((x) => flagKey(x.lot_code, x.package_no, x.supplier_name)));
  const taxAbnormalSet = new Set((qtAudit?.taxAnomalies || []).map((x) => flagKey(x.lot_code, x.package_no, x.supplier_name)));
  const taxMissingSet = new Set((qtAudit?.taxMissing || []).map((x) => flagKey(x.lot_code, x.package_no, x.supplier_name)));

  const displayedPackages = sortedPackages.filter((p) => {
    if (detailCountFilter === 'lt3' && !(p.bidder_count > 0 && p.detailed_eval_count < 3)) return false;
    if (detailCountFilter === 'lt2' && !(p.bidder_count > 0 && p.detailed_eval_count < 2)) return false;
    if (detailCountFilter === 'eq0' && !(p.bidder_count > 0 && p.detailed_eval_count === 0)) return false;
    if (detailCountFilter === 'failed' && !p.is_failed_lot) return false;
    if (pkgSearch.trim()) {
      const q = pkgSearch.trim().toLowerCase();
      const hay = `${p.lot_code} ${p.lot_name} ${p.package_no} ${p.project_name} ${p.first_candidate} ${p.req_app_no}`.toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });

  // Filtered Quote/Tax audit lists
  const qtQuery = qtSearch.trim().toLowerCase();
  const filteredOverLimits = overLimitRows.filter(
    (x) =>
      !qtQuery ||
      `${x.lot_code} ${x.lot_name} ${x.package_no} ${x.req_app_no} ${x.sub_name} ${x.supplier_name} ${x.scope} ${x.round_label}`
        .toLowerCase()
        .includes(qtQuery)
  );
  const filterTaxList = (list: NonNullable<typeof qtAudit>['taxAnomalies']) =>
    !qtQuery
      ? list
      : list.filter((x) =>
          `${x.lot_code} ${x.lot_name} ${x.package_no} ${x.req_app_no} ${x.project_name} ${x.supplier_name} ${x.tax_rate_raw}`
            .toLowerCase()
            .includes(qtQuery)
        );
  const filteredTaxAbnormal = filterTaxList(qtAudit?.taxAnomalies || []);
  const filteredTaxMissing = filterTaxList(qtAudit?.taxMissing || []);
  const filteredTaxProof = filterTaxList(qtAudit?.taxProofNeeded || []);
  const filteredSameQuotes = (qtAudit?.sameQuotes || []).filter(
    (x) =>
      !qtQuery ||
      `${x.lot_code} ${x.lot_name} ${x.package_no} ${x.req_app_no} ${x.project_name} ${x.quote_value} ${x.suppliers.join(' ')}`
        .toLowerCase()
        .includes(qtQuery)
  );
  const filteredRoundIncreases = (qtAudit?.roundIncreases || []).filter(
    (x) =>
      !qtQuery ||
      `${x.lot_code} ${x.lot_name} ${x.package_no} ${x.req_app_no} ${x.project_name} ${x.supplier_name}`
        .toLowerCase()
        .includes(qtQuery)
  );
  const filteredLowQuotes = (qtAudit?.lowQuotes || []).filter(
    (x) =>
      !qtQuery ||
      `${x.lot_code} ${x.lot_name} ${x.package_no} ${x.project_name} ${x.supplier_name} ${x.triggered_rules.join(' ')}`
        .toLowerCase()
        .includes(qtQuery)
  );

  // Filtered Collusion lists
  const colQuery = collusionSearch.trim().toLowerCase();
  const filteredContactOverlaps = (contactAudit?.overlaps || []).filter(
    (x) =>
      !colQuery ||
      `${x.lot_code} ${x.lot_name} ${x.package_no} ${x.field_label} ${x.matched_value} ${x.suppliers.join(' ')}`
        .toLowerCase()
        .includes(colQuery)
  );
  const filteredContactMissing = (contactAudit?.missingContacts || []).filter(
    (x) =>
      !colQuery ||
      `${x.lot_code} ${x.lot_name} ${x.package_no} ${x.supplier_name}`.toLowerCase().includes(colQuery)
  );
  const filteredHardwareAlerts = (hardwareAudit?.alerts || []).filter(
    (x) =>
      !colQuery ||
      `${x.lot_code} ${x.lot_name} ${x.package_no} ${x.supplier_a} ${x.supplier_b} ${x.source_a || ''} ${x.source_b || ''} ${x.cpu} ${x.disk} ${x.mac}`
        .toLowerCase()
        .includes(colQuery)
  );
  const filteredHardwareManual = (hardwareAudit?.manualReview || []).filter(
    (x) =>
      !colQuery ||
      `${x.lot_code} ${x.lot_name} ${x.package_no} ${x.supplier_a} ${x.supplier_b} ${x.source_a || ''} ${x.source_b || ''} ${x.cpu} ${x.disk} ${x.mac} ${x.review_reason}`
        .toLowerCase()
        .includes(colQuery)
  );
  const filteredHardwareInsufficient = (hardwareAudit?.insufficient || []).filter(
    (x) =>
      !colQuery ||
      `${x.lot_code} ${x.lot_name} ${x.package_no} ${x.supplier_name} ${x.file_type} ${x.source_location || ''}`.toLowerCase().includes(colQuery)
  );

  return (
    <div className="space-y-5">
      {/* Header with Procurement Type Rule Reminder */}
      <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <div className="flex flex-wrap items-center gap-2">
            <h1 className="text-lg font-semibold text-slate-800">2. 开评标模块（内网离线协同）</h1>
            <span className="rounded-md bg-[#EBF2F7] px-2 py-0.5 text-xs font-medium text-[#3B5B7E]">
              {isPublicBidding ? '公开招标 · 单轮报价（开标须做价格公示）' : '非公开招标 · 支持多轮报价'}
            </span>
            <span className="rounded-md bg-slate-100 px-2 py-0.5 text-xs font-medium text-slate-600">
              {terms.platformFull}
            </span>
          </div>
          <p className="mt-1 text-xs text-slate-500">
            上传内网导出的 5 类核心表格，按子版块分区完成评审总览、报价与税率核验、围串标风控核验（联系信息/硬件特征）及现场报告准备。
          </p>
        </div>
      </div>

      {/* Batch Change Notice Reminder Banner (read-only mirror from Bidding) */}
      {currentBatchId && changeNotices.length > 0 && (
        <ChangeNoticeBanner batchId={currentBatchId} notices={changeNotices} canManage={false} />
      )}

      {/* Cleaning / Validation Report Banner */}
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
              <span>【{lastReport.slotLabel}】解析与入库完成</span>
            </div>
            <button onClick={() => setLastReport(null)} className="text-xs text-slate-400 hover:text-slate-600">
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

      {/* 5-Slot Standard Evaluation Data Upload Matrix (collapsible on sub-views) */}
      <div className="rounded-xl border border-slate-200 bg-white p-3.5 space-y-3">
        <div className="flex items-center justify-between gap-2">
          <div className="flex items-center gap-2">
            <h2 className="text-sm font-semibold text-slate-700">开评标核心数据槽位</h2>
            <span className="rounded-full bg-slate-100 px-2 py-0.5 text-[11px] font-medium text-slate-500 tabular-nums">
              已上传 {slots.filter((s) => s.file).length} / {slots.length} 份
            </span>
          </div>
          <button
            type="button"
            onClick={() => setSlotsOpen((v) => !v)}
            className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-slate-50 px-2.5 py-1 text-[11px] font-medium text-slate-600 hover:bg-slate-100"
          >
            {slotsOpen ? <ChevronUp className="h-3.5 w-3.5" /> : <ChevronDown className="h-3.5 w-3.5" />}
            <span>{slotsOpen ? '收起槽位' : '展开上传'}</span>
          </button>
        </div>

        {slotsOpen && (
        <div className="grid grid-cols-1 gap-3 md:grid-cols-2 lg:grid-cols-5">
          {slots.map((slot) => {
            const isUploading = uploadingSlot === slot.key;
            const hasFile = Boolean(slot.file);

            return (
              <div
                key={slot.key}
                className={`flex flex-col justify-between rounded-lg border p-3 transition-colors ${
                  hasFile ? 'border-[#2F6F5E]/30 bg-[#ECFDF5]/25' : 'border-slate-200 bg-slate-50/40'
                }`}
              >
                <div>
                  <div className="flex items-center justify-between gap-1">
                    <span className="inline-flex h-5 w-5 items-center justify-center rounded-full bg-[#3B5B7E]/10 text-[11px] font-semibold text-[#3B5B7E] tabular-nums">
                      {slot.order}
                    </span>
                    <span className="text-[10px] text-slate-400">
                      {hasFile ? '已解析' : '待上传'}
                    </span>
                  </div>
                  <div className="mt-2 text-xs font-semibold text-slate-800">{slot.label}</div>
                  <p className="mt-1 text-[11px] text-slate-500 leading-relaxed line-clamp-2" title={slot.description}>
                    {slot.description}
                  </p>
                </div>

                <div className="mt-3 pt-2 border-t border-slate-200/60 space-y-2">
                  {slot.file ? (
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
                  ) : (
                    <div className="text-[11px] text-slate-400">支持 .xlsx / .xls</div>
                  )}

                  <div className="flex items-center gap-1.5">
                    <label className="flex flex-1 cursor-pointer items-center justify-center gap-1 rounded-md border border-slate-200 bg-white py-1.5 text-[11px] font-medium text-slate-700 hover:border-[#3B5B7E] hover:text-[#3B5B7E] transition-colors">
                      <Upload className="h-3 w-3" />
                      <span>{isUploading ? '处理中...' : hasFile ? '重新上传' : '上传并自动提取'}</span>
                      <input
                        type="file"
                        accept=".xlsx,.xls"
                        onChange={(e) => handleSlotSelect(slot.key, e)}
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
        )}
      </div>

      {/* Active Sub-Section Workspace (driven by left sidebar secondary nav) */}
      <div className="rounded-xl border border-slate-200 bg-white overflow-hidden">
        <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 bg-slate-50/70 px-4 py-2.5">
          <div className="flex flex-wrap items-center gap-2">
            {(activeTab === 'results' || activeTab === 'rejections') && (
              <div className="flex flex-wrap gap-1">
                <button
                  type="button"
                  onClick={() => setActiveTab('results')}
                  className={`inline-flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-xs font-medium transition-colors ${
                    activeTab === 'results' ? 'bg-[#3B5B7E] text-white' : 'text-slate-600 hover:bg-slate-200/60'
                  }`}
                >
                  <Trophy className="h-3.5 w-3.5" />
                  <span>标段评审结果总览 ({sortedPackages.length})</span>
                  {failedCount > 0 && (
                    <span className="rounded-full bg-[#A94442] px-1.5 text-[10px] text-white">
                      {failedCount} 流标
                    </span>
                  )}
                </button>
                <button
                  type="button"
                  onClick={() => setActiveTab('rejections')}
                  className={`inline-flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-xs font-medium transition-colors ${
                    activeTab === 'rejections' ? 'bg-[#3B5B7E] text-white' : 'text-slate-600 hover:bg-slate-200/60'
                  }`}
                >
                  <XCircle className="h-3.5 w-3.5" />
                  <span>未进入详评 / 否决 ({ws?.rejections?.length || 0})</span>
                </button>
              </div>
            )}

            {activeTab === 'quote_tax' && (
              <div className="flex items-center gap-2 text-sm font-semibold text-slate-800">
                <Scale className="h-4 w-4 text-[#3B5B7E]" />
                <span>报价与税率核验</span>
                {qtTotalIssues > 0 && (
                  <span className="rounded-full bg-[#A94442] px-2 py-0.5 text-[10px] font-medium text-white tabular-nums">
                    {qtTotalIssues} 项异常/待核
                  </span>
                )}
              </div>
            )}

            {activeTab === 'collusion' && (
              <div className="flex items-center gap-2 text-sm font-semibold text-slate-800">
                <ShieldAlert className="h-4 w-4 text-[#3B5B7E]" />
                <span>围串标风控核验</span>
                {collusionTotalAlerts > 0 && (
                  <span className="rounded-full bg-[#A94442] px-2 py-0.5 text-[10px] font-medium text-white tabular-nums">
                    {collusionTotalAlerts} 项预警
                  </span>
                )}
              </div>
            )}

            {activeTab === 'tools' && (
              <div className="flex flex-wrap items-center gap-1.5">
                <span className="inline-flex items-center gap-1.5 text-xs font-semibold text-slate-500 mr-1">
                  <Calculator className="h-3.5 w-3.5 text-[#3B5B7E]" />
                  <span>现场核验与算分：</span>
                </span>
                {([
                  { key: 'score_precheck', label: '专家打分预校验' },
                  { key: 'score_check', label: '专家打分结果校验' },
                  { key: 'price_score', label: '价格分计算' },
                  { key: 'performance_score', label: '履约分生成' },
                  { key: 'economic_rank', label: '经济排序表核验' },
                  { key: 'award_alloc', label: '授标分配' },
                  { key: 'split_pack', label: '智能分包' },
                ] as const).map((item) => (
                  <button
                    key={item.key}
                    type="button"
                    onClick={() => setActiveToolKey(item.key)}
                    className={`rounded-lg px-3 py-1.5 text-xs font-medium transition-colors ${
                      activeToolKey === item.key ? 'bg-[#3B5B7E] text-white' : 'text-slate-600 hover:bg-slate-200/60'
                    }`}
                  >
                    {item.label}
                  </button>
                ))}
              </div>
            )}

            {activeTab === 'reports' && (
              <div className="flex flex-wrap items-center gap-1.5">
                <span className="inline-flex items-center gap-1.5 text-xs font-semibold text-slate-500 mr-1">
                  <FileText className="h-3.5 w-3.5 text-[#3B5B7E]" />
                  <span>现场汇报与报告：</span>
                </span>
                <button
                  type="button"
                  onClick={() => setActiveReportKey('warning_form')}
                  className={`rounded-lg px-3 py-1.5 text-xs font-medium transition-colors ${
                    activeReportKey === 'warning_form' ? 'bg-[#3B5B7E] text-white' : 'text-slate-600 hover:bg-slate-200/60'
                  }`}
                >
                  预警事项确认单
                </button>
                <button
                  type="button"
                  onClick={() => setActiveReportKey('initial_report')}
                  className={`rounded-lg px-3 py-1.5 text-xs font-medium transition-colors ${
                    activeReportKey === 'initial_report' ? 'bg-[#3B5B7E] text-white' : 'text-slate-600 hover:bg-slate-200/60'
                  }`}
                >
                  初评汇报生成
                </button>
                <button
                  type="button"
                  onClick={() => setActiveReportKey('initial_minutes')}
                  className={`rounded-lg px-3 py-1.5 text-xs font-medium transition-colors ${
                    activeReportKey === 'initial_minutes' ? 'bg-[#3B5B7E] text-white' : 'text-slate-600 hover:bg-slate-200/60'
                  }`}
                >
                  初评会议纪要
                </button>
                <button
                  type="button"
                  onClick={() => setActiveReportKey('roadmap')}
                  className={`rounded-lg px-3 py-1.5 text-xs font-medium transition-colors ${
                    activeReportKey === 'roadmap' ? 'bg-[#3B5B7E] text-white' : 'text-slate-600 hover:bg-slate-200/60'
                  }`}
                >
                  待接入模板 ({REPORT_ROADMAP_ITEMS.length})
                </button>
              </div>
            )}

            {activeTab === 'sop' && (
              <div className="flex items-center gap-2 text-sm font-semibold text-slate-800">
                <ClipboardCheck className="h-4 w-4 text-[#3B5B7E]" />
                <span>SOP 流程与阶段附件 ({doneCount}/{steps.length})</span>
              </div>
            )}
          </div>

          {activeTab === 'results' && sortedPackages.length > 0 && (
            <div className="flex flex-wrap items-center gap-2 text-xs text-slate-500 tabular-nums">
              <span>
                已推荐候选人 <strong className="text-[#2F6F5E]">{recommendedCount}</strong> 个 · 流标{' '}
                <strong className="text-[#A94442]">{failedCount}</strong> 个
              </span>
            </div>
          )}
        </div>

        {/* TAB 1: Per-Package Evaluation Results */}
        {activeTab === 'results' && (
          <div className="p-4 space-y-3">
            {sortedPackages.length > 0 && (
              <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
                <div className="flex flex-wrap items-center gap-1.5">
                  {([
                    { key: 'all', label: `全部 (${sortedPackages.length})` },
                    { key: 'lt3', label: `详评 < 3 家 (${lt3Count})` },
                    { key: 'lt2', label: `详评 < 2 家 (${lt2Count})` },
                    { key: 'eq0', label: `详评 = 0 家 (${eq0Count})` },
                    { key: 'failed', label: `流标 (${failedCount})` },
                  ] as const).map((item) => (
                    <button
                      key={item.key}
                      type="button"
                      onClick={() => setDetailCountFilter(item.key)}
                      className={`rounded-md px-2.5 py-1 text-[11px] font-medium tabular-nums transition-colors ${
                        detailCountFilter === item.key
                          ? item.key === 'all'
                            ? 'bg-slate-800 text-white'
                            : 'bg-[#A94442] text-white'
                          : 'bg-slate-100 text-slate-600 hover:bg-slate-200/70'
                      }`}
                    >
                      {item.label}
                    </button>
                  ))}
                </div>
                <div className="relative w-full sm:w-64">
                  <Search className="pointer-events-none absolute left-2.5 top-2 h-3.5 w-3.5 text-slate-400" />
                  <input
                    value={pkgSearch}
                    onChange={(e) => setPkgSearch(e.target.value)}
                    placeholder="搜索分标、包号、项目名或候选人..."
                    className="w-full rounded-lg border border-slate-200 py-1.5 pl-8 pr-3 text-xs focus:border-[#3B5B7E] focus:outline-none"
                  />
                </div>
              </div>
            )}
            <div className="overflow-x-auto">
              <table className="w-full border-collapse text-left text-xs">
                <thead>
                  <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                    <th className="px-2 py-2.5 w-8"></th>
                    <th className="px-2.5 py-2.5 font-medium whitespace-nowrap">分标 / 包号</th>
                    <th className="px-2.5 py-2.5 font-medium min-w-[200px]">具体项目名称</th>
                    <th className="px-2.5 py-2.5 font-medium whitespace-nowrap">响应 / 否决 / 详评</th>
                    <th className="px-2.5 py-2.5 font-medium whitespace-nowrap">评审状态</th>
                    <th className="px-2.5 py-2.5 font-medium whitespace-nowrap">
                      第一{isPublicBidding ? '中标' : '成交'}候选人
                    </th>
                    <th className="px-2.5 py-2.5 font-medium whitespace-nowrap">最终报价 (万元/%)</th>
                    <th className="px-2.5 py-2.5 font-medium whitespace-nowrap">综合得分</th>
                    <th className="px-2.5 py-2.5 font-medium whitespace-nowrap">相对限价</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {sortedPackages.length === 0 ? (
                    <tr>
                      <td colSpan={9} className="py-10 text-center text-xs text-slate-400">
                        暂无标段评审数据，请在上方上传《开标一览表》与《总分排序表》
                      </td>
                    </tr>
                  ) : (
                    paginate(displayedPackages, pkgPage, pkgPageSize).map((pkg) => {
                      const expanded = expandedPkgKeys.has(pkg.key);
                      return (
                        <React.Fragment key={pkg.key}>
                          <tr className={pkg.is_failed_lot ? 'bg-[#FEF2F2]/35' : 'hover:bg-slate-50/70'}>
                            <td className="px-2 py-2.5">
                              <button
                                onClick={() => toggleExpandPkg(pkg.key)}
                                className="rounded p-0.5 text-slate-400 hover:bg-slate-200/70 hover:text-slate-700"
                                title="展开候选人排序与多轮报价详情"
                              >
                                {expanded ? <ChevronUp className="h-3.5 w-3.5" /> : <ChevronDown className="h-3.5 w-3.5" />}
                              </button>
                            </td>
                            <td className="px-2.5 py-2.5 whitespace-nowrap">
                              <div className="font-mono font-medium text-slate-700">{pkg.lot_code || '—'}</div>
                              <div className="text-[11px] text-slate-400">
                                {pkg.lot_name} · {pkg.package_no}
                              </div>
                            </td>
                            <td className="px-2.5 py-2.5">
                              <div className="flex flex-wrap items-center gap-1.5">
                                <span className="font-medium text-slate-800">{pkg.project_name || '—'}</span>
                                {(pkgChangeNoticeMap.get(`${pkg.lot_code || ''}::${pkg.package_no || ''}`) || []).map((cn) =>
                                  cn.has_file ? (
                                    <a
                                      key={cn.id}
                                      href={`/api/bidding/change-notices/${cn.id}/file`}
                                      target="_blank"
                                      rel="noreferrer"
                                      title={`变更公告：${cn.title}${cn.summary ? ` — ${cn.summary}` : ''}（点击查看原件）`}
                                      className="inline-flex items-center rounded bg-[#FEF9EC] border border-[#9A6B26]/35 px-1.5 py-0.5 text-[10px] font-semibold text-[#9A6B26] hover:bg-[#9A6B26] hover:text-white transition-colors"
                                    >
                                      有变更{cn.change_types[0] ? ` · ${cn.change_types[0]}` : ''}
                                    </a>
                                  ) : (
                                    <span
                                      key={cn.id}
                                      title={`变更公告：${cn.title}${cn.summary ? ` — ${cn.summary}` : ''}`}
                                      className="inline-flex items-center rounded bg-[#FEF9EC] border border-[#9A6B26]/35 px-1.5 py-0.5 text-[10px] font-semibold text-[#9A6B26]"
                                    >
                                      有变更{cn.change_types[0] ? ` · ${cn.change_types[0]}` : ''}
                                    </span>
                                  )
                                )}
                              </div>
                              {pkg.req_app_no && (
                                <div className="text-[10px] font-mono text-slate-400 mt-0.5">
                                  申请号：{pkg.req_app_no}
                                </div>
                              )}
                            </td>
                            <td className="px-2.5 py-2.5 whitespace-nowrap tabular-nums">
                              <div className="flex items-center gap-1.5">
                                <span className="inline-flex min-w-[46px] justify-center rounded bg-slate-100 px-1.5 py-0.5 text-[10px] font-medium text-slate-700">
                                  响应 {pkg.bidder_count}
                                </span>
                                <span
                                  className={`inline-flex min-w-[46px] justify-center rounded px-1.5 py-0.5 text-[10px] font-medium ${
                                    pkg.rejection_count > 0
                                      ? 'bg-[#FEF2F2] text-[#A94442]'
                                      : 'bg-slate-100 text-slate-400'
                                  }`}
                                >
                                  否决 {pkg.rejection_count}
                                </span>
                                <span
                                  className={`inline-flex min-w-[46px] justify-center rounded px-1.5 py-0.5 text-[10px] font-medium ${
                                    pkg.bidder_count === 0
                                      ? 'bg-slate-100 text-slate-400'
                                      : pkg.detailed_eval_count < (isPublicBidding ? 3 : 2)
                                        ? 'bg-[#FEF2F2] text-[#A94442]'
                                        : 'bg-[#EBF2F7] text-[#3B5B7E]'
                                  }`}
                                >
                                  详评 {pkg.detailed_eval_count}
                                </span>
                              </div>
                            </td>
                            <td className="px-2.5 py-2.5 whitespace-nowrap">
                              {pkg.is_failed_lot ? (
                                <div className="inline-flex items-center gap-1">
                                  <span className="rounded bg-[#FEF2F2] px-2 py-0.5 text-[11px] font-medium text-[#A94442]">
                                    流标
                                  </span>
                                  {pkg.failed_reason && (
                                    <span className="text-[10px] text-[#A94442]/80">{pkg.failed_reason}</span>
                                  )}
                                </div>
                              ) : pkg.first_candidate ? (
                                <span className="rounded bg-[#ECFDF5] px-2 py-0.5 text-[11px] font-medium text-[#2F6F5E]">
                                  正常推荐
                                </span>
                              ) : (
                                <span className="rounded bg-slate-100 px-2 py-0.5 text-[11px] text-slate-500">
                                  待导入排序表
                                </span>
                              )}
                            </td>
                            <td className="px-2.5 py-2.5 font-medium text-slate-800">
                              {pkg.is_failed_lot ? '—' : pkg.first_candidate || '—'}
                            </td>
                            <td className="px-2.5 py-2.5 whitespace-nowrap tabular-nums font-medium text-[#3B5B7E]">
                              {pkg.first_price_wan ? (
                                pkg.first_price_wan
                              ) : pkg.bid_price_summary ? (
                                <span
                                  className="text-[11px] font-normal text-slate-500"
                                  title="以最新轮次（二轮优先）开标报价统计，点击左侧箭头展开查看各供应商明细"
                                >
                                  开标 {pkg.bid_price_summary}
                                </span>
                              ) : (
                                '—'
                              )}
                            </td>
                            <td className="px-2.5 py-2.5 whitespace-nowrap tabular-nums font-semibold text-slate-700">
                              {pkg.first_score !== null ? pkg.first_score.toFixed(2) : '—'}
                            </td>
                            <td className="px-2.5 py-2.5 whitespace-nowrap tabular-nums text-[11px]">
                              {pkg.discount_pct ? (
                                <span
                                  className={
                                    pkg.discount_pct.startsWith('超限')
                                      ? 'text-[#A94442] font-semibold'
                                      : 'text-[#2F6F5E]'
                                  }
                                >
                                  {pkg.discount_pct}
                                </span>
                              ) : (
                                '—'
                              )}
                            </td>
                          </tr>

                          {expanded && (
                            <tr className="bg-slate-50/75">
                              <td colSpan={9} className="px-6 py-3.5 space-y-3">
                                <div className="flex flex-wrap items-center gap-x-4 gap-y-1 rounded-lg border border-slate-200/80 bg-white px-3 py-2 text-[11px] text-slate-600 tabular-nums">
                                  <span>
                                    报价方式：
                                    <strong className="text-slate-800">
                                      {pkg.quote_method || (pkg.max_price_raw ? '折扣率报价' : '总价报价')}
                                    </strong>
                                  </span>
                                  <span>
                                    本包最高限价：
                                    {pkg.max_price_raw ? (
                                      <strong className="text-[#3B5B7E]">{pkg.max_price_raw}</strong>
                                    ) : pkg.max_price_yuan !== null ? (
                                      <strong className="text-[#3B5B7E]">
                                        {Number((pkg.max_price_yuan / 10000).toFixed(4))} 万元
                                      </strong>
                                    ) : (
                                      <span className="text-slate-400">未导入需求一览表限价</span>
                                    )}
                                  </span>
                                  {pkg.max_price_raw && pkg.max_price_yuan !== null && (
                                    <span>
                                      含税预算：
                                      <strong className="text-slate-700">
                                        {Number((pkg.max_price_yuan / 10000).toFixed(4))} 万元
                                      </strong>
                                    </span>
                                  )}
                                  {pkg.sub_packages && pkg.sub_packages.length > 1 && (
                                    <div className="w-full border-t border-slate-100 pt-1.5 mt-0.5 flex flex-wrap gap-2 text-[10px] text-slate-500">
                                      <span className="font-semibold text-[#9A6B26]">合并子包独立限价：</span>
                                      {pkg.sub_packages.map((sp, i) => (
                                        <span key={sp.req_app_no || i} className="rounded bg-slate-50 border border-slate-200/70 px-1.5 py-0.5">
                                          {sp.name}（限价{' '}
                                          <strong className="text-slate-700">
                                            {sp.max_price_raw || (sp.max_price_yuan !== null ? `${Number((sp.max_price_yuan / 10000).toFixed(4))}万元` : '—')}
                                          </strong>
                                          ）
                                        </span>
                                      ))}
                                    </div>
                                  )}
                                </div>

                                {pkg.candidates.length > 0 && (
                                  <div>
                                    <div className="text-[11px] font-semibold text-slate-600 mb-1.5">
                                      综合评分排序明细（评审日期：{pkg.eval_date || '—'}）
                                    </div>
                                    <table className="w-full border-collapse text-left text-[11px] bg-white rounded-lg overflow-hidden border border-slate-200/80">
                                      <thead>
                                        <tr className="bg-slate-100/80 text-slate-500">
                                          <th className="px-2.5 py-1.5">排序</th>
                                          <th className="px-2.5 py-1.5">{terms.bidder}</th>
                                          <th className="px-2.5 py-1.5">技术分</th>
                                          <th className="px-2.5 py-1.5">商务分</th>
                                          <th className="px-2.5 py-1.5">价格分</th>
                                          <th className="px-2.5 py-1.5">总分</th>
                                          <th className="px-2.5 py-1.5">基准价</th>
                                          <th className="px-2.5 py-1.5">最终评标/评审总价</th>
                                        </tr>
                                      </thead>
                                      <tbody className="divide-y divide-slate-100">
                                        {pkg.candidates.map((c) => (
                                          <tr key={c.id} className={c.rank_no === 1 ? 'bg-[#ECFDF5]/40' : ''}>
                                            <td className="px-2.5 py-1.5 font-mono font-semibold">{c.rank_no}</td>
                                            <td className="px-2.5 py-1.5 font-medium text-slate-800">{c.supplier_name}</td>
                                            <td className="px-2.5 py-1.5 tabular-nums">{c.tech_score ?? '—'}</td>
                                            <td className="px-2.5 py-1.5 tabular-nums">{c.biz_score ?? '—'}</td>
                                            <td className="px-2.5 py-1.5 tabular-nums">{c.price_score ?? '—'}</td>
                                            <td className="px-2.5 py-1.5 tabular-nums font-semibold text-[#3B5B7E]">{c.total_score ?? '—'}</td>
                                            <td className="px-2.5 py-1.5 tabular-nums text-slate-500">{c.base_price || '—'}</td>
                                            <td className="px-2.5 py-1.5 tabular-nums font-medium">{c.eval_price || c.bid_price || '—'}</td>
                                          </tr>
                                        ))}
                                      </tbody>
                                    </table>
                                  </div>
                                )}

                                {pkg.supplier_bids && pkg.supplier_bids.length > 0 && (
                                  <div>
                                    <div className="text-[11px] font-semibold text-slate-600 mb-1.5">
                                      开标与行信息报价明细（共 {pkg.supplier_bids.length} 家 · 有二轮报价时以二轮为最终报价）
                                    </div>
                                    <table className="w-full border-collapse text-left text-[11px] bg-white rounded-lg overflow-hidden border border-slate-200/80">
                                      <thead>
                                        <tr className="bg-slate-100/80 text-slate-500">
                                          <th className="px-2.5 py-1.5 w-8">#</th>
                                          <th className="px-2.5 py-1.5">{terms.bidder}</th>
                                          <th className="px-2.5 py-1.5">第1轮报价</th>
                                          <th className="px-2.5 py-1.5">第2轮报价</th>
                                          <th className="px-2.5 py-1.5">最终报价 (万元/%)</th>
                                          <th className="px-2.5 py-1.5">增值税率</th>
                                          <th className="px-2.5 py-1.5">行报价 (含税/未含税)</th>
                                          <th className="px-2.5 py-1.5">初评状态</th>
                                        </tr>
                                      </thead>
                                      <tbody className="divide-y divide-slate-100">
                                        {pkg.supplier_bids.map((sb, sbIdx) => (
                                          <React.Fragment key={sb.supplier_name}>
                                            <tr className={sb.is_rejected ? 'bg-[#FEF2F2]/35' : ''}>
                                              <td className="px-2.5 py-1.5 text-slate-400 tabular-nums">{sbIdx + 1}</td>
                                              <td className="px-2.5 py-1.5 font-medium text-slate-800">{sb.supplier_name}</td>
                                              <td className="px-2.5 py-1.5 tabular-nums text-slate-500">{sb.round1_price || '—'}</td>
                                              <td className="px-2.5 py-1.5 tabular-nums text-slate-600">{sb.round2_price || '—'}</td>
                                              <td className="px-2.5 py-1.5 whitespace-nowrap tabular-nums font-medium text-[#3B5B7E]">
                                                <div className="flex flex-wrap items-center gap-1">
                                                  <span className="font-semibold">{sb.final_price || '—'}</span>
                                                  {sb.round2_price && (
                                                    <span className="rounded bg-[#EBF2F7] px-1 py-0.2 text-[9px] font-medium text-[#3B5B7E]">
                                                      二轮
                                                    </span>
                                                  )}
                                                  {sb.is_over_limit && (
                                                    <span className="rounded bg-[#A94442] px-1.5 py-0.5 text-[10px] font-medium text-white">
                                                      超限价
                                                    </span>
                                                  )}
                                                  {lowQuoteSet.has(flagKey(pkg.lot_code, pkg.package_no, sb.supplier_name)) && (
                                                    <span className="rounded bg-[#C2410C] px-1.5 py-0.5 text-[10px] font-medium text-white">
                                                      报价过低
                                                    </span>
                                                  )}
                                                  {roundIncreaseSet.has(flagKey(pkg.lot_code, pkg.package_no, sb.supplier_name)) && (
                                                    <span className="rounded bg-[#991B1B] px-1.5 py-0.5 text-[10px] font-medium text-white">
                                                      二轮超一轮
                                                    </span>
                                                  )}
                                                </div>
                                              </td>
                                              <td className="px-2.5 py-1.5 whitespace-nowrap tabular-nums text-slate-600">
                                                <div className="flex flex-wrap items-center gap-1">
                                                  <span>{sb.tax_rate ? `${sb.tax_rate}%` : '—'}</span>
                                                  {taxAbnormalSet.has(flagKey(pkg.lot_code, pkg.package_no, sb.supplier_name)) && (
                                                    <span className="rounded bg-[#9A6B26] px-1.5 py-0.5 text-[10px] font-medium text-white">
                                                      税率异常
                                                    </span>
                                                  )}
                                                  {taxMissingSet.has(flagKey(pkg.lot_code, pkg.package_no, sb.supplier_name)) && (
                                                    <span className="rounded bg-slate-600 px-1.5 py-0.5 text-[10px] font-medium text-white">
                                                      税率缺失
                                                    </span>
                                                  )}
                                                </div>
                                              </td>
                                              <td className="px-2.5 py-1.5 tabular-nums text-slate-500">
                                                {sb.price_inc_tax_wan !== null || sb.price_ex_tax_wan !== null
                                                  ? `${sb.price_inc_tax_wan ?? '—'} / ${sb.price_ex_tax_wan ?? '—'}`
                                                  : '—'}
                                              </td>
                                              <td className="px-2.5 py-1.5 whitespace-nowrap">
                                                {sb.is_rejected ? (
                                                  <span className="rounded bg-[#FEF2F2] px-1.5 py-0.5 text-[10px] text-[#A94442]">
                                                    否决
                                                  </span>
                                                ) : (
                                                  <span className="rounded bg-[#ECFDF5] px-1.5 py-0.5 text-[10px] text-[#2F6F5E]">
                                                    正常
                                                  </span>
                                                )}
                                              </td>
                                            </tr>
                                            {sb.sub_quotes && sb.sub_quotes.length > 0 && (
                                              <tr className="bg-slate-50/60">
                                                <td></td>
                                                <td colSpan={7} className="px-2.5 py-1.5 text-[10px] text-slate-500">
                                                  <div className="flex flex-wrap gap-1.5">
                                                    <span className="font-medium text-slate-600">子包各轮报价：</span>
                                                    {sb.sub_quotes.map((sq, sqIdx) => (
                                                      <span
                                                        key={`${sq.req_app_no}-${sq.round_label}-${sqIdx}`}
                                                        className="rounded px-1.5 py-0.5 border border-slate-200 bg-white text-slate-600 tabular-nums"
                                                        title={`${sq.sub_name}（限价 ${sq.limit_display}）`}
                                                      >
                                                        [{sq.round_label}] {sq.sub_name}：<strong>{sq.price_inc_tax_wan ?? sq.discount_rate ?? '—'}</strong>（限{sq.limit_display}）
                                                      </span>
                                                    ))}
                                                  </div>
                                                </td>
                                              </tr>
                                            )}
                                          </React.Fragment>
                                        ))}
                                      </tbody>
                                    </table>
                                  </div>
                                )}

                                {pkg.rejections.length > 0 && (
                                  <div>
                                    <div className="text-[11px] font-semibold text-[#A94442] mb-1">
                                      本包未进入详评（否决）情况（{pkg.rejections.length} 条）
                                    </div>
                                    <ul className="space-y-1 text-[11px] text-slate-600">
                                      {pkg.rejections.map((rj) => (
                                        <li key={rj.id}>
                                          • <strong>{rj.supplier_name}</strong> [{rj.review_type} / {rj.reason_category}]：{rj.bid_response}
                                        </li>
                                      ))}
                                    </ul>
                                  </div>
                                )}
                              </td>
                            </tr>
                          )}
                        </React.Fragment>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>
            {displayedPackages.length > 20 && (
              <div className="overflow-hidden rounded-lg border border-slate-200">
                <PaginationBar
                  total={displayedPackages.length}
                  page={pkgPage}
                  pageSize={pkgPageSize}
                  onPageChange={setPkgPage}
                  onPageSizeChange={setPkgPageSize}
                />
              </div>
            )}
          </div>
        )}

        {/* TAB 2: Quote & Tax Compliance Audit (报价与税率核验) */}
        {activeTab === 'quote_tax' && (
          <div className="p-4 space-y-4">
            {/* Config & Export Bar */}
            <div className="flex flex-col gap-3 rounded-lg border border-slate-200 bg-slate-50/70 p-3.5">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <div className="flex flex-wrap items-center gap-3 text-xs">
                  <span className="font-semibold text-slate-700">税率校验口径：</span>
                  <div className="inline-flex rounded-md border border-slate-200 bg-white p-0.5">
                    {(['服务', '物资'] as const).map((cat) => (
                      <button
                        key={cat}
                        type="button"
                        disabled={savingAuditCfg}
                        onClick={() => handleUpdateAuditConfig(cat, proofRates)}
                        className={`rounded px-2.5 py-1 text-[11px] font-medium transition-colors ${
                          taxCategory === cat
                            ? 'bg-[#3B5B7E] text-white'
                            : 'text-slate-600 hover:bg-slate-100'
                        }`}
                      >
                        {cat === '服务' ? '服务类（0/1/3/5/6/9/13%）' : '物资类（仅 13%）'}
                      </button>
                    ))}
                  </div>
                </div>

                <div className="flex flex-wrap items-center gap-2">
                  <a
                    href={`/api/evaluation/batch/${currentBatchId}/quote-audit/export`}
                    className="inline-flex items-center gap-1 rounded-md border border-[#A94442]/30 bg-[#FEF2F2]/60 px-2.5 py-1.5 text-[11px] font-medium text-[#A94442] hover:bg-[#FEE2E2] transition-colors"
                    title="导出超限价、报价过低、二轮超一轮、相同报价明细"
                  >
                    <Download className="h-3.5 w-3.5" />
                    <span>导出报价核验 (.xlsx)</span>
                  </a>
                  <a
                    href={`/api/evaluation/batch/${currentBatchId}/tax-audit/export`}
                    className="inline-flex items-center gap-1 rounded-md border border-[#3B5B7E]/30 bg-white px-2.5 py-1.5 text-[11px] font-medium text-[#3B5B7E] hover:bg-[#EBF2F7]/60 transition-colors"
                    title="导出税率异常、税率缺失、税率证明待核实明细"
                  >
                    <Download className="h-3.5 w-3.5" />
                    <span>导出税率核验 (.xlsx)</span>
                  </a>
                </div>
              </div>

              <div className="flex flex-wrap items-center gap-2 border-t border-slate-200/70 pt-2.5 text-xs">
                <span className="font-semibold text-slate-700">证明核查税率（独立判断）：</span>
                <div className="flex flex-wrap items-center gap-1.5">
                  {PROOF_RATE_OPTIONS.map((rate) => {
                    const checked = proofRates.includes(rate);
                    return (
                      <button
                        key={rate}
                        type="button"
                        disabled={savingAuditCfg}
                        onClick={() => handleToggleProofRate(rate)}
                        className={`rounded-md border px-2 py-0.5 text-[11px] font-medium tabular-nums transition-colors ${
                          checked
                            ? 'border-[#9A6B26] bg-[#FEF9EC] text-[#9A6B26]'
                            : 'border-slate-200 bg-white text-slate-500 hover:border-slate-300'
                        }`}
                      >
                        {checked ? '✓ ' : ''}
                        {rate}%
                      </button>
                    );
                  })}
                </div>
                <span className="text-[11px] text-slate-400">
                  （命中即列为「税率证明待核实」，供人工查阅投标文件证明；未匹配限价的行同样校验税率）
                </span>
              </div>
            </div>

            {/* Sub-filter tabs + Search */}
            <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
              <div className="flex flex-wrap gap-1.5">
                {([
                  { key: 'over_limit', label: `超限价 (${overLimitRows.length})`, danger: true },
                  { key: 'low_quote', label: `报价过低 (${qtAudit?.counts.lowQuote || 0})`, danger: true },
                  { key: 'round_increase', label: `二轮超一轮 (${qtAudit?.counts.roundIncrease || 0})`, danger: true },
                  { key: 'same_quote', label: `不同供应商相同报价 (${qtAudit?.counts.sameQuote || 0})`, danger: true },
                  { key: 'tax_abnormal', label: `税率异常 (${qtAudit?.counts.taxAbnormal || 0})`, danger: true },
                  { key: 'tax_missing', label: `税率缺失 (${qtAudit?.counts.taxMissing || 0})`, danger: false },
                  { key: 'tax_proof', label: `税率证明待核实 (${qtAudit?.counts.taxProof || 0})`, danger: false },
                ] as const).map((t) => (
                  <button
                    key={t.key}
                    type="button"
                    onClick={() => setQtSubTab(t.key)}
                    className={`rounded-md px-2.5 py-1.5 text-xs font-medium tabular-nums transition-colors ${
                      qtSubTab === t.key
                        ? t.danger
                          ? 'bg-[#A94442] text-white'
                          : 'bg-[#3B5B7E] text-white'
                        : 'bg-slate-100 text-slate-600 hover:bg-slate-200/70'
                    }`}
                  >
                    {t.label}
                  </button>
                ))}
              </div>
              <div className="relative w-full sm:w-64">
                <Search className="pointer-events-none absolute left-2.5 top-2 h-3.5 w-3.5 text-slate-400" />
                <input
                  value={qtSearch}
                  onChange={(e) => setQtSearch(e.target.value)}
                  placeholder="搜索分标、包号、申请号或供应商..."
                  className="w-full rounded-lg border border-slate-200 py-1.5 pl-8 pr-3 text-xs focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>
            </div>

            {/* Sub-table: 超限价核验（整包 + 合并子包各轮次） */}
            {qtSubTab === 'over_limit' && (
              <div className="space-y-2">
                <div className="rounded-lg border border-slate-200 bg-slate-50/70 px-3 py-2 text-[11px] text-slate-600 flex items-start gap-2">
                  <Info className="h-3.5 w-3.5 shrink-0 text-[#3B5B7E] mt-0.5" />
                  <span>
                    超限价核验规则：按《需求一览表》最高限价逐轮核验各投标人报价（同时兼容总价与折扣率）；存在合并成包（子包）时，自动对每个子包的每一轮报价单独比对子包限价。
                  </span>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse text-left text-xs">
                    <thead>
                      <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">维度 / 轮次</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">分标 / 包号</th>
                        <th className="px-2.5 py-2 font-medium min-w-[180px]">采购申请号 / 项目（子包）名称</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">{terms.bidder}</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">最高限价</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">超限报价</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">超出幅度</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">初评状态</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {filteredOverLimits.length === 0 ? (
                        <tr>
                          <td colSpan={8} className="py-8 text-center text-xs text-slate-400">
                            {ws?.stats.bidRowCount || ws?.stats.lineQuoteCount ? '未发现超过最高限价的报价记录' : '请先上传《开标一览表》或《行信息报价基础数据表》'}
                          </td>
                        </tr>
                      ) : (
                        paginate(filteredOverLimits, qtPager.page, qtPager.pageSize).map((item, idx) => (
                          <tr key={idx} className="bg-[#FEF2F2]/30 hover:bg-[#FEF2F2]/60">
                            <td className="px-2.5 py-2 whitespace-nowrap">
                              <span
                                className={`rounded px-1.5 py-0.5 text-[10px] font-medium ${
                                  item.scope === '子包'
                                    ? 'bg-[#FEF9EC] text-[#9A6B26]'
                                    : 'bg-slate-100 text-slate-600'
                                }`}
                              >
                                {item.scope}
                              </span>
                              <span className="ml-1 text-[11px] text-slate-500">{item.round_label}</span>
                            </td>
                            <td className="px-2.5 py-2 whitespace-nowrap">
                              <div className="font-mono font-medium text-slate-700">{item.lot_code}</div>
                              <div className="text-[11px] text-slate-400">{item.lot_name} · {item.package_no}</div>
                            </td>
                            <td className="px-2.5 py-2">
                              {item.req_app_no && <div className="font-mono text-[11px] text-slate-500">{item.req_app_no}</div>}
                              <div className="text-slate-700">{item.sub_name || '—'}</div>
                            </td>
                            <td className="px-2.5 py-2 font-medium text-slate-800 whitespace-nowrap">{item.supplier_name}</td>
                            <td className="px-2.5 py-2 whitespace-nowrap tabular-nums text-slate-600">{item.limit_display}</td>
                            <td className="px-2.5 py-2 whitespace-nowrap tabular-nums font-semibold text-[#A94442]">{item.quote_display}</td>
                            <td className="px-2.5 py-2 whitespace-nowrap tabular-nums">
                              <span className="rounded bg-[#A94442] px-1.5 py-0.5 text-[10px] font-medium text-white">
                                {item.exceed_text}
                              </span>
                            </td>
                            <td className="px-2.5 py-2 whitespace-nowrap">
                              {item.is_rejected ? (
                                <span className="rounded bg-slate-100 px-1.5 py-0.5 text-[10px] text-slate-600">
                                  已否决
                                </span>
                              ) : (
                                <span className="rounded bg-[#FEF2F2] px-1.5 py-0.5 text-[10px] font-medium text-[#A94442]">
                                  未否决（需核查）
                                </span>
                              )}
                            </td>
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>
                {filteredOverLimits.length > 20 && (
                  <div className="overflow-hidden rounded-lg border border-slate-200">
                    <PaginationBar
                      total={filteredOverLimits.length}
                      page={qtPager.page}
                      pageSize={qtPager.pageSize}
                      onPageChange={qtPager.setPage}
                      onPageSizeChange={qtPager.setPageSize}
                    />
                  </div>
                )}
              </div>
            )}

            {/* Sub-table: 报价过低（低于成本） */}
            {qtSubTab === 'low_quote' && (
              <div className="space-y-2">
                <div className="rounded-lg border border-slate-200 bg-slate-50/70 px-3 py-2 text-[11px] text-slate-600 flex items-start gap-2">
                  <Info className="h-3.5 w-3.5 shrink-0 text-[#3B5B7E] mt-0.5" />
                  <span>
                    低于成本判定标准（满足任一情形即预警，基于同标包通过初步评审应答人最终报价）：① 包应答报价 &lt; 全部通过初评应答人报价平均值 × 50%；② 包应答报价 &lt; 通过初评次低报价应答人报价 × 50%；③ 应答报价 &lt; 最高限价 × 45%。
                  </span>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse text-left text-xs">
                    <thead>
                      <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">分标 / 包号</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">{terms.bidder}</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">应答报价</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">初评均值 / 次低价 / 限价</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">触发规则</th>
                        <th className="px-2.5 py-2 font-medium">核验说明</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {filteredLowQuotes.length === 0 ? (
                        <tr>
                          <td colSpan={6} className="py-8 text-center text-xs text-slate-400">
                            {ws?.stats.bidRowCount || ws?.stats.lineQuoteCount ? '未发现低于成本（报价过低）预警记录' : '请先上传《开标一览表》或《行信息报价基础数据表》'}
                          </td>
                        </tr>
                      ) : (
                        paginate(filteredLowQuotes, qtPager.page, qtPager.pageSize).map((item, idx) => (
                          <tr key={idx} className="bg-[#FEF2F2]/30 hover:bg-[#FEF2F2]/60">
                            <td className="px-2.5 py-2 whitespace-nowrap">
                              <div className="font-mono font-medium text-slate-700">{item.lot_code}</div>
                              <div className="text-[11px] text-slate-400">{item.lot_name} · {item.package_no}</div>
                            </td>
                            <td className="px-2.5 py-2">
                              <div className="font-medium text-slate-800 whitespace-nowrap">{item.supplier_name}</div>
                              <div className="text-[10px] text-slate-400">{item.round_label} · {item.quote_basis}</div>
                            </td>
                            <td className="px-2.5 py-2 whitespace-nowrap tabular-nums font-semibold text-[#A94442]">{item.quote_value}</td>
                            <td className="px-2.5 py-2 whitespace-nowrap tabular-nums text-[11px] text-slate-600">
                              均值 {item.avg_value} · 次低 {item.second_lowest_value} · 限价 {item.max_price_value}
                            </td>
                            <td className="px-2.5 py-2">
                              <div className="flex flex-wrap gap-1">
                                {item.triggered_rules.map((r) => (
                                  <span key={r} className="rounded bg-[#A94442] px-1.5 py-0.5 text-[10px] font-medium text-white whitespace-nowrap">
                                    {r}
                                  </span>
                                ))}
                              </div>
                            </td>
                            <td className="px-2.5 py-2 text-slate-600">{item.detail}</td>
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>
                {filteredLowQuotes.length > 20 && (
                  <div className="overflow-hidden rounded-lg border border-slate-200">
                    <PaginationBar
                      total={filteredLowQuotes.length}
                      page={qtPager.page}
                      pageSize={qtPager.pageSize}
                      onPageChange={qtPager.setPage}
                      onPageSizeChange={qtPager.setPageSize}
                    />
                  </div>
                )}
              </div>
            )}

            {/* Sub-table: 二轮超一轮 */}
            {qtSubTab === 'round_increase' && (
              <div className="space-y-2">
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse text-left text-xs">
                    <thead>
                      <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">维度</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">分标 / 包号</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">采购申请号 / 项目名称</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">{terms.bidder}</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">报价口径</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">第1轮</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">第2轮</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">上涨幅度</th>
                        <th className="px-2.5 py-2 font-medium">核验说明</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {filteredRoundIncreases.length === 0 ? (
                        <tr>
                          <td colSpan={9} className="py-8 text-center text-xs text-slate-400">
                            {ws?.stats.bidRowCount || ws?.stats.lineQuoteCount ? '未发现同一供应商二轮超一轮（否决项）' : '请先上传《开标一览表》或《行信息报价基础数据表》'}
                          </td>
                        </tr>
                      ) : (
                        paginate(filteredRoundIncreases, qtPager.page, qtPager.pageSize).map((item, idx) => (
                          <tr key={idx} className="bg-[#FEF2F2]/30 hover:bg-[#FEF2F2]/60">
                            <td className="px-2.5 py-2 whitespace-nowrap">
                              <span
                                className={`rounded px-1.5 py-0.5 text-[10px] font-medium ${
                                  item.scope === '子包'
                                    ? 'bg-[#FEF9EC] text-[#9A6B26]'
                                    : 'bg-slate-100 text-slate-600'
                                }`}
                              >
                                {item.scope || '整包'}
                              </span>
                            </td>
                            <td className="px-2.5 py-2 whitespace-nowrap">
                              <div className="font-mono font-medium text-slate-700">{item.lot_code}</div>
                              <div className="text-[11px] text-slate-400">{item.lot_name} · {item.package_no}</div>
                            </td>
                            <td className="px-2.5 py-2">
                              <div className="font-mono text-[11px] text-slate-600">{item.req_app_no || '—'}</div>
                              <div className="text-slate-700">{item.project_name || '—'}</div>
                            </td>
                            <td className="px-2.5 py-2 font-medium text-slate-800 whitespace-nowrap">{item.supplier_name}</td>
                            <td className="px-2.5 py-2 whitespace-nowrap text-slate-600">{item.quote_basis}</td>
                            <td className="px-2.5 py-2 whitespace-nowrap tabular-nums text-slate-600">{item.round1_value}</td>
                            <td className="px-2.5 py-2 whitespace-nowrap tabular-nums font-semibold text-[#A94442]">{item.round2_value}</td>
                            <td className="px-2.5 py-2 whitespace-nowrap tabular-nums">
                              <span className="rounded bg-[#A94442] px-1.5 py-0.5 text-[10px] font-medium text-white">
                                {item.increase_text}
                              </span>
                            </td>
                            <td className="px-2.5 py-2 text-slate-600">{item.detail}</td>
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>
                {filteredRoundIncreases.length > 20 && (
                  <div className="overflow-hidden rounded-lg border border-slate-200">
                    <PaginationBar
                      total={filteredRoundIncreases.length}
                      page={qtPager.page}
                      pageSize={qtPager.pageSize}
                      onPageChange={qtPager.setPage}
                      onPageSizeChange={qtPager.setPageSize}
                    />
                  </div>
                )}
              </div>
            )}

            {/* Sub-table: 不同供应商相同报价 */}
            {qtSubTab === 'same_quote' && (
              <div className="space-y-2">
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse text-left text-xs">
                    <thead>
                      <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">维度</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">分标 / 包号 / 轮次</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">采购申请号 / 项目名称</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">报价口径</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">相同数值</th>
                        <th className="px-2.5 py-2 font-medium min-w-[240px]">涉及供应商</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {filteredSameQuotes.length === 0 ? (
                        <tr>
                          <td colSpan={6} className="py-8 text-center text-xs text-slate-400">
                            {ws?.stats.bidRowCount || ws?.stats.lineQuoteCount ? '未发现同标包同轮次相同报价' : '请先上传《开标一览表》或《行信息报价基础数据表》'}
                          </td>
                        </tr>
                      ) : (
                        paginate(filteredSameQuotes, qtPager.page, qtPager.pageSize).map((item, idx) => (
                          <tr key={idx} className="hover:bg-slate-50/70">
                            <td className="px-2.5 py-2 whitespace-nowrap">
                              <span
                                className={`rounded px-1.5 py-0.5 text-[10px] font-medium ${
                                  item.scope === '子包'
                                    ? 'bg-[#FEF9EC] text-[#9A6B26]'
                                    : 'bg-slate-100 text-slate-600'
                                }`}
                              >
                                {item.scope || '整包'}
                              </span>
                            </td>
                            <td className="px-2.5 py-2 whitespace-nowrap">
                              <div className="font-mono font-medium text-slate-700">{item.lot_code}</div>
                              <div className="text-[11px] text-slate-400">
                                {item.lot_name} · {item.package_no} · {item.round_label}
                              </div>
                            </td>
                            <td className="px-2.5 py-2">
                              <div className="font-mono text-[11px] text-slate-600">{item.req_app_no || '—'}</div>
                              <div className="text-slate-700">{item.project_name || '—'}</div>
                            </td>
                            <td className="px-2.5 py-2 whitespace-nowrap text-slate-600">{item.quote_basis}</td>
                            <td className="px-2.5 py-2 whitespace-nowrap tabular-nums font-semibold text-[#A94442]">
                              {item.quote_value}
                            </td>
                            <td className="px-2.5 py-2 text-slate-800">
                              <div className="flex flex-wrap gap-1">
                                {item.suppliers.map((s) => (
                                  <span key={s} className="rounded bg-[#FEF2F2] px-1.5 py-0.5 text-[11px] font-medium text-[#A94442]">
                                    {s}
                                  </span>
                                ))}
                              </div>
                            </td>
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>
                {filteredSameQuotes.length > 20 && (
                  <div className="overflow-hidden rounded-lg border border-slate-200">
                    <PaginationBar
                      total={filteredSameQuotes.length}
                      page={qtPager.page}
                      pageSize={qtPager.pageSize}
                      onPageChange={qtPager.setPage}
                      onPageSizeChange={qtPager.setPageSize}
                    />
                  </div>
                )}
              </div>
            )}

            {/* Sub-table: 税率异常 / 税率缺失 / 税率证明待核实 */}
            {(qtSubTab === 'tax_abnormal' || qtSubTab === 'tax_missing' || qtSubTab === 'tax_proof') && (() => {
              const rows =
                qtSubTab === 'tax_abnormal'
                  ? filteredTaxAbnormal
                  : qtSubTab === 'tax_missing'
                    ? filteredTaxMissing
                    : filteredTaxProof;
              return (
                <div className="space-y-2">
                  <div className="overflow-x-auto">
                    <table className="w-full border-collapse text-left text-xs">
                      <thead>
                        <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                          <th className="px-2.5 py-2 font-medium whitespace-nowrap">分标 / 包号 / 轮次</th>
                          <th className="px-2.5 py-2 font-medium whitespace-nowrap">采购申请号 / 项目名称</th>
                          <th className="px-2.5 py-2 font-medium whitespace-nowrap">{terms.bidder}</th>
                          <th className="px-2.5 py-2 font-medium whitespace-nowrap">填报税率</th>
                          <th className="px-2.5 py-2 font-medium whitespace-nowrap">核验结论</th>
                          <th className="px-2.5 py-2 font-medium">说明</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {rows.length === 0 ? (
                          <tr>
                            <td colSpan={6} className="py-8 text-center text-xs text-slate-400">
                              {ws?.stats.lineQuoteCount ? '该项税率核验无异常记录' : '请先上传《行信息报价基础数据表》'}
                            </td>
                          </tr>
                        ) : (
                          paginate(rows, qtPager.page, qtPager.pageSize).map((item, idx) => (
                            <tr key={idx} className="hover:bg-slate-50/70">
                              <td className="px-2.5 py-2 whitespace-nowrap">
                                <div className="font-mono font-medium text-slate-700">{item.lot_code}</div>
                                <div className="text-[11px] text-slate-400">
                                  {item.lot_name} · {item.package_no} · {item.round_label}
                                </div>
                              </td>
                              <td className="px-2.5 py-2">
                                <div className="font-mono text-[11px] text-slate-600">{item.req_app_no || '—'}</div>
                                <div className="text-slate-700">{item.project_name || '—'}</div>
                              </td>
                              <td className="px-2.5 py-2 font-medium text-slate-800 whitespace-nowrap">{item.supplier_name}</td>
                              <td className="px-2.5 py-2 whitespace-nowrap tabular-nums font-mono font-semibold text-slate-700">
                                {item.tax_rate_raw}
                              </td>
                              <td className="px-2.5 py-2 whitespace-nowrap">
                                <span
                                  className={`rounded px-1.5 py-0.5 text-[11px] font-medium ${
                                    item.type === 'tax_abnormal'
                                      ? 'bg-[#FEF2F2] text-[#A94442]'
                                      : item.type === 'tax_missing'
                                        ? 'bg-slate-200 text-slate-700'
                                        : 'bg-[#FEF9EC] text-[#9A6B26]'
                                  }`}
                                >
                                  {item.label}
                                </span>
                              </td>
                              <td className="px-2.5 py-2 text-slate-600">{item.detail}</td>
                            </tr>
                          ))
                        )}
                      </tbody>
                    </table>
                  </div>
                  {rows.length > 20 && (
                    <div className="overflow-hidden rounded-lg border border-slate-200">
                      <PaginationBar
                        total={rows.length}
                        page={qtPager.page}
                        pageSize={qtPager.pageSize}
                        onPageChange={qtPager.setPage}
                        onPageSizeChange={qtPager.setPageSize}
                      />
                    </div>
                  )}
                </div>
              );
            })()}
          </div>
        )}

        {/* TAB 3: Collusion Risk Audit (围串标风控核验：联系信息 + 硬件特征) */}
        {activeTab === 'collusion' && (
          <div className="p-4 space-y-4">
            <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
              <div className="flex flex-wrap gap-1.5">
                <button
                  type="button"
                  onClick={() => setCollusionSubTab('contact')}
                  className={`inline-flex items-center gap-1.5 rounded-md px-3 py-1.5 text-xs font-medium transition-colors ${
                    collusionSubTab === 'contact'
                      ? 'bg-[#3B5B7E] text-white'
                      : 'bg-slate-100 text-slate-600 hover:bg-slate-200/70'
                  }`}
                >
                  <Users className="h-3.5 w-3.5" />
                  <span>同标包供应商联系信息核验</span>
                  {(contactAudit?.counts.overlapCount || 0) > 0 && (
                    <span className="rounded-full bg-[#A94442] px-1.5 text-[10px] text-white tabular-nums">
                      {contactAudit!.counts.overlapCount}
                    </span>
                  )}
                </button>
                <button
                  type="button"
                  onClick={() => setCollusionSubTab('hardware')}
                  className={`inline-flex items-center gap-1.5 rounded-md px-3 py-1.5 text-xs font-medium transition-colors ${
                    collusionSubTab === 'hardware'
                      ? 'bg-[#3B5B7E] text-white'
                      : 'bg-slate-100 text-slate-600 hover:bg-slate-200/70'
                  }`}
                >
                  <Cpu className="h-3.5 w-3.5" />
                  <span>投标文件硬件特征核验 (CPU/硬盘/MAC)</span>
                  {(hardwareAudit?.counts.alertCount || 0) > 0 && (
                    <span className="rounded-full bg-[#A94442] px-1.5 text-[10px] text-white tabular-nums">
                      {hardwareAudit!.counts.alertCount}
                    </span>
                  )}
                </button>
              </div>

              <div className="flex flex-wrap items-center gap-2">
                <div className="relative w-full sm:w-56">
                  <Search className="pointer-events-none absolute left-2.5 top-2 h-3.5 w-3.5 text-slate-400" />
                  <input
                    value={collusionSearch}
                    onChange={(e) => setCollusionSearch(e.target.value)}
                    placeholder="搜索分标、包号、供应商或特征..."
                    className="w-full rounded-lg border border-slate-200 py-1.5 pl-8 pr-3 text-xs focus:border-[#3B5B7E] focus:outline-none"
                  />
                </div>
                {collusionSubTab === 'contact' ? (
                  <a
                    href={`/api/evaluation/batch/${currentBatchId}/contact-audit/export`}
                    className="inline-flex items-center gap-1 rounded-md border border-[#3B5B7E]/30 bg-white px-2.5 py-1.5 text-[11px] font-medium text-[#3B5B7E] hover:bg-[#EBF2F7]/60 transition-colors"
                  >
                    <Download className="h-3.5 w-3.5" />
                    <span>导出联系信息核验 (.xlsx)</span>
                  </a>
                ) : (
                  <a
                    href={`/api/evaluation/batch/${currentBatchId}/hardware-audit/export`}
                    className="inline-flex items-center gap-1 rounded-md border border-[#3B5B7E]/30 bg-white px-2.5 py-1.5 text-[11px] font-medium text-[#3B5B7E] hover:bg-[#EBF2F7]/60 transition-colors"
                  >
                    <Download className="h-3.5 w-3.5" />
                    <span>导出硬件特征核验 (.xlsx)</span>
                  </a>
                )}
              </div>
            </div>

            {collusionSubTab === 'contact' && (
              <div className="space-y-4">
                <div className="rounded-lg border border-slate-200 bg-slate-50/70 px-3.5 py-2.5 text-[11px] text-slate-600 flex items-start gap-2">
                  <Info className="h-3.5 w-3.5 shrink-0 text-[#3B5B7E] mt-0.5" />
                  <span>
                    核验规则：严格限定<strong>同一分标 + 同一包</strong>内不同供应商（核验名单来自《开标一览表》，联系字段来自招标模块《供应商联系人信息表》）；逐项比对联系人姓名、手机号码、固定电话、电子邮件、传真号码、邮政地址（邮政地址须至少含一个数字或“号”字才参与比对）；跨包不触发预警。
                  </span>
                </div>

                <div>
                  <h3 className="text-xs font-semibold text-slate-700 mb-2">
                    同标包联系信息重合预警（{filteredContactOverlaps.length} 条）
                  </h3>
                  <div className="overflow-hidden rounded-lg border border-slate-200">
                    <div className="overflow-x-auto">
                      <table className="w-full border-collapse text-left text-xs">
                        <thead>
                          <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                            <th className="px-2.5 py-2 font-medium whitespace-nowrap">分标 / 包号</th>
                            <th className="px-2.5 py-2 font-medium whitespace-nowrap">预警类型</th>
                            <th className="px-2.5 py-2 font-medium whitespace-nowrap">重合字段</th>
                            <th className="px-2.5 py-2 font-medium min-w-[160px]">重合内容</th>
                            <th className="px-2.5 py-2 font-medium min-w-[260px]">涉及供应商（同标包）</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {filteredContactOverlaps.length === 0 ? (
                            <tr>
                              <td colSpan={5} className="py-7 text-center text-xs text-slate-400">
                                {ws?.stats.bidRowCount
                                  ? '未发现同标包不同供应商联系信息重合'
                                  : '请先在上方上传《开标一览表》（并确保招标模块已上传《供应商联系人信息表》）'}
                              </td>
                            </tr>
                          ) : (
                            paginate(filteredContactOverlaps, contactOverlapPager.page, contactOverlapPager.pageSize).map((item, idx) => (
                              <tr key={idx} className="bg-[#FEF2F2]/25 hover:bg-[#FEF2F2]/50">
                                <td className="px-2.5 py-2 whitespace-nowrap">
                                  <div className="font-mono font-medium text-slate-700">{item.lot_code}</div>
                                  <div className="text-[11px] text-slate-400">{item.lot_name} · {item.package_no}</div>
                                </td>
                                <td className="px-2.5 py-2 whitespace-nowrap">
                                  <span className="rounded bg-[#FEF2F2] px-1.5 py-0.5 text-[11px] font-medium text-[#A94442]">
                                    {item.alert_label}
                                  </span>
                                </td>
                                <td className="px-2.5 py-2 whitespace-nowrap text-slate-600">{item.field_label}</td>
                                <td className="px-2.5 py-2 font-mono text-slate-800">{item.matched_value}</td>
                                <td className="px-2.5 py-2">
                                  <div className="flex flex-wrap gap-1">
                                    {item.suppliers.map((s) => (
                                      <span key={s} className="rounded border border-[#A94442]/25 bg-white px-1.5 py-0.5 text-[11px] font-medium text-slate-800">
                                        {s}
                                      </span>
                                    ))}
                                  </div>
                                </td>
                              </tr>
                            ))
                          )}
                        </tbody>
                      </table>
                    </div>
                    <PaginationBar
                      total={filteredContactOverlaps.length}
                      page={contactOverlapPager.page}
                      pageSize={contactOverlapPager.pageSize}
                      onPageChange={contactOverlapPager.setPage}
                      onPageSizeChange={contactOverlapPager.setPageSize}
                    />
                  </div>
                </div>

                {(contactAudit?.counts.missingCount || 0) > 0 && (
                  <div>
                    <h3 className="text-xs font-semibold text-[#9A6B26] mb-2">
                      未找到供应商联系人信息（{filteredContactMissing.length} 条）
                    </h3>
                    <div className="overflow-hidden rounded-lg border border-[#9A6B26]/25">
                      <div className="overflow-x-auto">
                        <table className="w-full border-collapse text-left text-xs">
                          <thead>
                            <tr className="border-b border-[#9A6B26]/20 bg-[#FEF9EC]/70 text-[#9A6B26]">
                              <th className="px-2.5 py-2 font-medium whitespace-nowrap">分标 / 包号</th>
                              <th className="px-2.5 py-2 font-medium whitespace-nowrap">{terms.bidder}</th>
                              <th className="px-2.5 py-2 font-medium">提示</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 bg-white">
                            {paginate(filteredContactMissing, contactMissingPager.page, contactMissingPager.pageSize).map((m, idx) => (
                              <tr key={idx}>
                                <td className="px-2.5 py-2 whitespace-nowrap">
                                  <span className="font-mono text-slate-700">{m.lot_code}</span>
                                  <span className="ml-1.5 text-[11px] text-slate-400">{m.lot_name} · {m.package_no}</span>
                                </td>
                                <td className="px-2.5 py-2 font-medium text-slate-800">{m.supplier_name}</td>
                                <td className="px-2.5 py-2 text-[#9A6B26]">{m.detail}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      <PaginationBar
                        total={filteredContactMissing.length}
                        page={contactMissingPager.page}
                        pageSize={contactMissingPager.pageSize}
                        onPageChange={contactMissingPager.setPage}
                        onPageSizeChange={contactMissingPager.setPageSize}
                      />
                    </div>
                  </div>
                )}
              </div>
            )}

            {collusionSubTab === 'hardware' && (
              <div className="space-y-4">
                <div className="rounded-lg border border-slate-200 bg-slate-50/70 px-3.5 py-2.5 text-[11px] text-slate-600 flex items-start gap-2">
                  <Info className="h-3.5 w-3.5 shrink-0 text-[#3B5B7E] mt-0.5" />
                  <span>
                    核验规则：同一分标 + 同一包内（缺分标编号时按分标名称区分，定位不全归资料不足），不同投标人的技术、商务、价格、资格/资质业绩四类文件<strong>跨类别、跨轮次</strong>比对（同一投标人自身跳过，全角/半角占位符自动按空处理）；先比 CPU 与硬盘序列号：<strong>三项全一致</strong>报「围串标预警」，<strong>CPU + 硬盘一致但单方或双方缺 MAC</strong> 报「待人工核实」，双方 MAC 明确不同不预警；缺 CPU 或硬盘（但填有其他字段）列为「资料不足，未完成核验」，四项全空跳过。
                  </span>
                </div>

                <div>
                  <h3 className="text-xs font-semibold text-slate-700 mb-2">
                    围串标预警 — CPU + 硬盘 + MAC 全部一致（{filteredHardwareAlerts.length} 组）
                  </h3>
                  <div className="overflow-hidden rounded-lg border border-slate-200">
                    <div className="overflow-x-auto">
                      <table className="w-full border-collapse text-left text-xs">
                        <thead>
                          <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                            <th className="px-2.5 py-2 font-medium whitespace-nowrap">分标 / 包号</th>
                            <th className="px-2.5 py-2 font-medium whitespace-nowrap">投标人 A（文件及来源）</th>
                            <th className="px-2.5 py-2 font-medium whitespace-nowrap">投标人 B（文件及来源）</th>
                            <th className="px-2.5 py-2 font-medium whitespace-nowrap">CPU 序列号</th>
                            <th className="px-2.5 py-2 font-medium whitespace-nowrap">硬盘序列号</th>
                            <th className="px-2.5 py-2 font-medium whitespace-nowrap">网卡 MAC 地址</th>
                            <th className="px-2.5 py-2 font-medium whitespace-nowrap">提交 IP (A / B)</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {filteredHardwareAlerts.length === 0 ? (
                            <tr>
                              <td colSpan={7} className="py-7 text-center text-xs text-slate-400">
                                {(hardwareAudit?.totalRecords || 0) > 0
                                  ? `已核验 ${hardwareAudit!.totalRecords} 条记录，未发现 CPU+硬盘+MAC 全部一致的围串标预警`
                                  : '请在上方第 3 槽位上传《投标文件IP地址一览表》'}
                              </td>
                            </tr>
                          ) : (
                            paginate(filteredHardwareAlerts, hwAlertPager.page, hwAlertPager.pageSize).map((a, idx) => (
                              <tr key={idx} className="bg-[#FEF2F2]/30 hover:bg-[#FEF2F2]/55">
                                <td className="px-2.5 py-2 whitespace-nowrap">
                                  <div className="font-mono font-medium text-slate-700">{a.lot_code}</div>
                                  <div className="text-[11px] text-slate-400">{a.lot_name} · {a.package_no}</div>
                                </td>
                                <td className="px-2.5 py-2">
                                  <div className="font-medium text-slate-800">{a.supplier_a}</div>
                                  <div className="mt-0.5 text-[10px] font-mono text-[#A94442] whitespace-pre-line leading-relaxed">
                                    {a.source_a || `${a.round_a} · ${a.file_a}`}
                                  </div>
                                </td>
                                <td className="px-2.5 py-2">
                                  <div className="font-medium text-slate-800">{a.supplier_b}</div>
                                  <div className="mt-0.5 text-[10px] font-mono text-[#A94442] whitespace-pre-line leading-relaxed">
                                    {a.source_b || `${a.round_b} · ${a.file_b}`}
                                  </div>
                                </td>
                                <td className="px-2.5 py-2 font-mono text-[11px] text-slate-800 whitespace-nowrap">{a.cpu}</td>
                                <td className="px-2.5 py-2 font-mono text-[11px] text-slate-800 whitespace-nowrap">{a.disk}</td>
                                <td className="px-2.5 py-2 font-mono text-[11px] font-semibold text-[#A94442] whitespace-nowrap">{a.mac}</td>
                                <td className="px-2.5 py-2 font-mono text-[11px] text-slate-500 whitespace-nowrap">
                                  {a.ip_a || '/'} / {a.ip_b || '/'}
                                </td>
                              </tr>
                            ))
                          )}
                        </tbody>
                      </table>
                    </div>
                    <PaginationBar
                      total={filteredHardwareAlerts.length}
                      page={hwAlertPager.page}
                      pageSize={hwAlertPager.pageSize}
                      onPageChange={hwAlertPager.setPage}
                      onPageSizeChange={hwAlertPager.setPageSize}
                    />
                  </div>
                </div>

                {(hardwareAudit?.counts.manualReviewCount || 0) > 0 && (
                  <div>
                    <h3 className="text-xs font-semibold text-[#9A6B26] mb-2">
                      待人工核实 — CPU + 硬盘一致，单方或双方未提供 MAC（{filteredHardwareManual.length} 组）
                    </h3>
                    <div className="overflow-hidden rounded-lg border border-[#9A6B26]/30">
                      <div className="overflow-x-auto">
                        <table className="w-full border-collapse text-left text-xs">
                          <thead>
                            <tr className="border-b border-[#9A6B26]/20 bg-[#FEF9EC]/75 text-[#9A6B26]">
                              <th className="px-2.5 py-2 font-medium whitespace-nowrap">分标 / 包号</th>
                              <th className="px-2.5 py-2 font-medium whitespace-nowrap">投标人 A（文件及来源）</th>
                              <th className="px-2.5 py-2 font-medium whitespace-nowrap">投标人 B（文件及来源）</th>
                              <th className="px-2.5 py-2 font-medium whitespace-nowrap">CPU 序列号</th>
                              <th className="px-2.5 py-2 font-medium whitespace-nowrap">硬盘序列号</th>
                              <th className="px-2.5 py-2 font-medium whitespace-nowrap">网卡 MAC (A / B)</th>
                              <th className="px-2.5 py-2 font-medium">核验结论</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 bg-white">
                            {paginate(filteredHardwareManual, hwManualPager.page, hwManualPager.pageSize).map((m, idx) => (
                              <tr key={idx} className="hover:bg-[#FEF9EC]/40">
                                <td className="px-2.5 py-2 whitespace-nowrap">
                                  <div className="font-mono font-medium text-slate-700">{m.lot_code || '—'}</div>
                                  <div className="text-[11px] text-slate-400">{m.lot_name} · {m.package_no}</div>
                                </td>
                                <td className="px-2.5 py-2">
                                  <div className="font-medium text-slate-800">{m.supplier_a}</div>
                                  <div className="mt-0.5 text-[10px] font-mono text-[#9A6B26] whitespace-pre-line leading-relaxed">
                                    {m.source_a || `${m.round_a} · ${m.file_a}`}
                                  </div>
                                </td>
                                <td className="px-2.5 py-2">
                                  <div className="font-medium text-slate-800">{m.supplier_b}</div>
                                  <div className="mt-0.5 text-[10px] font-mono text-[#9A6B26] whitespace-pre-line leading-relaxed">
                                    {m.source_b || `${m.round_b} · ${m.file_b}`}
                                  </div>
                                </td>
                                <td className="px-2.5 py-2 font-mono text-[11px] text-slate-800 whitespace-nowrap">{m.cpu}</td>
                                <td className="px-2.5 py-2 font-mono text-[11px] text-slate-800 whitespace-nowrap">{m.disk}</td>
                                <td className="px-2.5 py-2 font-mono text-[11px] text-slate-500 whitespace-nowrap">{m.mac}</td>
                                <td className="px-2.5 py-2 font-medium text-[#9A6B26]">{m.review_reason}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      <PaginationBar
                        total={filteredHardwareManual.length}
                        page={hwManualPager.page}
                        pageSize={hwManualPager.pageSize}
                        onPageChange={hwManualPager.setPage}
                        onPageSizeChange={hwManualPager.setPageSize}
                      />
                    </div>
                  </div>
                )}

                {(hardwareAudit?.counts.insufficientCount || 0) > 0 && (
                  <div>
                    <h3 className="text-xs font-semibold text-[#9A6B26] mb-2">
                      资料不足，未完成核验（{filteredHardwareInsufficient.length} 条）
                    </h3>
                    <div className="overflow-hidden rounded-lg border border-[#9A6B26]/25">
                      <div className="overflow-x-auto">
                        <table className="w-full border-collapse text-left text-xs">
                          <thead>
                            <tr className="border-b border-[#9A6B26]/20 bg-[#FEF9EC]/70 text-[#9A6B26]">
                              <th className="px-2.5 py-2 font-medium whitespace-nowrap">分标 / 包号 / 轮次</th>
                              <th className="px-2.5 py-2 font-medium whitespace-nowrap">{terms.bidder}</th>
                              <th className="px-2.5 py-2 font-medium whitespace-nowrap">文件类别与来源位置</th>
                              <th className="px-2.5 py-2 font-medium whitespace-nowrap">IP / MAC / CPU / 硬盘</th>
                              <th className="px-2.5 py-2 font-medium">核验结论</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 bg-white">
                            {paginate(filteredHardwareInsufficient, hwInsuffPager.page, hwInsuffPager.pageSize).map((ins, idx) => (
                              <tr key={idx}>
                                <td className="px-2.5 py-2 whitespace-nowrap">
                                  <div className="font-mono text-slate-700">{ins.lot_code}</div>
                                  <div className="text-[11px] text-slate-400">{ins.lot_name} · {ins.package_no} · {ins.round_label}</div>
                                </td>
                                <td className="px-2.5 py-2 font-medium text-slate-800 whitespace-nowrap">{ins.supplier_name}</td>
                                <td className="px-2.5 py-2">
                                  <div className="text-slate-700 font-medium">{ins.file_type}</div>
                                  {ins.source_location && (
                                    <div className="text-[10px] font-mono text-slate-500 mt-0.5">{ins.source_location}</div>
                                  )}
                                </td>
                                <td className="px-2.5 py-2 font-mono text-[10px] text-slate-500">
                                  IP:{ins.ip || '/'} | MAC:{ins.mac || '/'} | CPU:{ins.cpu || '/'} | 硬盘:{ins.disk || '/'}
                                </td>
                                <td className="px-2.5 py-2 text-[#9A6B26]">{ins.detail}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      <PaginationBar
                        total={filteredHardwareInsufficient.length}
                        page={hwInsuffPager.page}
                        pageSize={hwInsuffPager.pageSize}
                        onPageChange={hwInsuffPager.setPage}
                        onPageSizeChange={hwInsuffPager.setPageSize}
                      />
                    </div>
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {/* TAB 4: Rejections Detail Table */}
        {activeTab === 'rejections' && (
          <div className="p-4 space-y-3">
            <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
              <div className="flex gap-1.5">
                {(['all', '商务', '技术'] as const).map((k) => (
                  <button
                    key={k}
                    onClick={() => setRejFilter(k)}
                    className={`rounded-md px-2.5 py-1 text-xs font-medium ${
                      rejFilter === k ? 'bg-slate-800 text-white' : 'bg-slate-100 text-slate-600 hover:bg-slate-200/70'
                    }`}
                  >
                    {k === 'all' ? '全部否决' : `${k}原因`}
                  </button>
                ))}
              </div>
              <div className="relative w-full sm:w-64">
                <Search className="pointer-events-none absolute left-2.5 top-2 h-3.5 w-3.5 text-slate-400" />
                <input
                  value={rejSearch}
                  onChange={(e) => setRejSearch(e.target.value)}
                  placeholder={`搜索包号、${terms.bidder}或否决原因...`}
                  className="w-full rounded-lg border border-slate-200 py-1.5 pl-8 pr-3 text-xs focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full border-collapse text-left text-xs">
                <thead>
                  <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                    <th className="px-2.5 py-2 font-medium whitespace-nowrap">分标 / 包号</th>
                    <th className="px-2.5 py-2 font-medium whitespace-nowrap">{terms.bidder}</th>
                    <th className="px-2.5 py-2 font-medium whitespace-nowrap">类别 / 原因</th>
                    <th className="px-2.5 py-2 font-medium min-w-[220px]">{terms.tenderDoc}要求</th>
                    <th className="px-2.5 py-2 font-medium min-w-[220px]">{terms.responseDoc}响应情况</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {filteredRejections.length === 0 ? (
                    <tr>
                      <td colSpan={5} className="py-8 text-center text-xs text-slate-400">
                        暂无未进入详评记录
                      </td>
                    </tr>
                  ) : (
                    paginate(filteredRejections, rejPager.page, rejPager.pageSize).map((r) => (
                      <tr key={r.id} className="hover:bg-slate-50/70">
                        <td className="px-2.5 py-2.5 whitespace-nowrap">
                          <div className="font-medium text-slate-700">{r.lot_name || r.lot_code}</div>
                          <div className="text-[11px] text-slate-400">{r.package_no}</div>
                        </td>
                        <td className="px-2.5 py-2.5 font-medium text-slate-800 whitespace-nowrap">{r.supplier_name}</td>
                        <td className="px-2.5 py-2.5 whitespace-nowrap">
                          <span
                            className={`inline-block rounded px-1.5 py-0.5 text-[11px] font-medium ${
                              r.review_type.includes('商务')
                                ? 'bg-[#FEF9EC] text-[#9A6B26]'
                                : 'bg-[#FEF2F2] text-[#A94442]'
                            }`}
                          >
                            {r.review_type || '否决'} · {r.reason_category}
                          </span>
                        </td>
                        <td className="px-2.5 py-2.5 text-slate-600 leading-relaxed">{r.doc_requirement || '—'}</td>
                        <td className="px-2.5 py-2.5 text-slate-700 leading-relaxed">{r.bid_response || '—'}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
            {filteredRejections.length > 20 && (
              <div className="overflow-hidden rounded-lg border border-slate-200">
                <PaginationBar
                  total={filteredRejections.length}
                  page={rejPager.page}
                  pageSize={rejPager.pageSize}
                  onPageChange={rejPager.setPage}
                  onPageSizeChange={rejPager.setPageSize}
                />
              </div>
            )}
          </div>
        )}

        {/* TAB 5A: On-Site Verification & Scoring Tools (单工具聚焦视图) */}
        {activeTab === 'tools' && (
          <div className="p-4 space-y-4">
            {activeToolKey === 'score_precheck' && currentBatchId && (
              <ExpertScorePrecheckPanel
                batchId={currentBatchId}
                rejectionCount={ws ? ws.rejections.length : undefined}
                hasRejectionFile={
                  ws ? Boolean(ws.slots.find((s) => s.key === 'rejections')?.file) : undefined
                }
              />
            )}
            {activeToolKey === 'score_check' && currentBatchId && (
              <ExpertScoreCheckPanel batchId={currentBatchId} />
            )}
            {activeToolKey === 'price_score' && currentBatchId && (
              <PriceScorePanel batchId={currentBatchId} />
            )}
            {activeToolKey === 'performance_score' && currentBatchId && (
              <PerformanceScorePanel batchId={currentBatchId} />
            )}
            {activeToolKey === 'economic_rank' && currentBatchId && (
              <EconomicRankPanel batchId={currentBatchId} />
            )}
            {activeToolKey === 'award_alloc' && currentBatchId && (
              <AwardAllocPanel batchId={currentBatchId} />
            )}
            {activeToolKey === 'split_pack' && currentBatchId && (
              <SplitPackPanel batchId={currentBatchId} />
            )}
          </div>
        )}

        {/* TAB 5B: On-Site Reports & Documents (单工具聚焦视图) */}
        {activeTab === 'reports' && (
          <div className="p-4 space-y-4">
            {activeReportKey === 'warning_form' && currentBatchId && (
              <WarningFormPanel batchId={currentBatchId} />
            )}
            {activeReportKey === 'initial_report' && currentBatchId && (
              <InitialReportPanel batchId={currentBatchId} />
            )}
            {activeReportKey === 'initial_minutes' && currentBatchId && (
              <InitialMinutesPanel batchId={currentBatchId} />
            )}
            {activeReportKey === 'roadmap' && (
              <div className="grid grid-cols-1 gap-3 md:grid-cols-3">
                {REPORT_ROADMAP_ITEMS.map((item) => (
                  <div
                    key={item.key}
                    className="flex flex-col justify-between rounded-xl border border-dashed border-slate-300 bg-slate-50/50 p-4 space-y-2"
                  >
                    <div className="flex items-center justify-between gap-2">
                      <h4 className="text-xs font-semibold text-slate-800">{item.title}</h4>
                      <span className="rounded-full bg-slate-200/80 px-2 py-0.5 text-[10px] font-medium text-slate-600">
                        {item.badge}
                      </span>
                    </div>
                    <p className="text-[11px] text-slate-500 leading-relaxed">{item.desc}</p>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        {/* TAB 6: SOP Checklist, Risk Notices & Other Attachments */}
        {activeTab === 'sop' && (
          <div className="p-4 space-y-5">
            <SopGuidesPanel />

            <div className="flex flex-wrap items-center justify-between gap-2 rounded-lg bg-slate-50 p-3">
              <div className="flex items-center gap-3 text-xs">
                <span className="font-medium text-slate-700">SOP 完成度：</span>
                <div className="h-2 w-36 overflow-hidden rounded-full bg-slate-200">
                  <div className="h-full rounded-full bg-[#3B5B7E]" style={{ width: `${sopPct}%` }} />
                </div>
                <span className="font-semibold text-[#3B5B7E] tabular-nums">
                  {doneCount}/{steps.length} ({sopPct}%)
                </span>
              </div>
              <button
                onClick={handleSaveAsTemplate}
                className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"
              >
                <BookmarkCheck className="h-3.5 w-3.5 text-[#3B5B7E]" />
                <span>{savedTplMsg || '覆盖保存当前 SOP 为全局默认模板'}</span>
              </button>
            </div>

            {isPublicBidding && (
              <div className="rounded-lg border border-[#3B5B7E]/25 bg-[#EBF2F7]/50 p-3 text-xs text-[#3B5B7E] flex items-center gap-2">
                <Info className="h-4 w-4 shrink-0" />
                <span>
                  <strong>公开招标特别提示</strong>：公开招标仅有一轮报价，开标操作环节务必完成<strong>价格公示</strong>。
                </span>
              </div>
            )}

            <div className="grid grid-cols-1 gap-5 lg:grid-cols-12">
              <div className="lg:col-span-7 space-y-4">
                {STAGE_GROUPS.map((group) => {
                  const groupSteps = steps.filter((s) => s.stage_group === group);
                  return (
                    <div key={group} className="rounded-xl border border-slate-200 bg-white p-4 space-y-2.5">
                      <div className="flex items-center justify-between border-b border-slate-100 pb-2">
                        <h3 className="text-sm font-semibold text-slate-700">{group}</h3>
                        <span className="text-xs text-slate-400 tabular-nums">
                          {groupSteps.filter((s) => s.is_completed === 1).length}/{groupSteps.length}
                        </span>
                      </div>

                      <div className="space-y-2">
                        {groupSteps.length === 0 ? (
                          <div className="text-xs text-slate-400 py-1">该子阶段暂无步骤，可在下方添加</div>
                        ) : (
                          groupSteps.map((s) => (
                            <div
                              key={s.id}
                              className={`rounded-lg border p-3 ${
                                s.is_completed
                                  ? 'border-[#2F6F5E]/20 bg-[#ECFDF5]/40'
                                  : 'border-slate-200/80 bg-white'
                              }`}
                            >
                              <div className="flex items-start gap-2.5">
                                <button onClick={() => toggleStep(s)} className="mt-0.5 shrink-0 text-slate-400 hover:text-[#3B5B7E]">
                                  {s.is_completed ? (
                                    <CheckCircle2 className="h-4 w-4 text-[#2F6F5E]" />
                                  ) : (
                                    <Circle className="h-4 w-4" />
                                  )}
                                </button>
                                <div className="flex-1 min-w-0">
                                  <div className="flex items-start justify-between gap-2">
                                    <span className={`text-xs font-medium ${s.is_completed ? 'text-slate-400 line-through' : 'text-slate-800'}`}>
                                      {s.title}
                                    </span>
                                    <button
                                      onClick={() => handleDeleteSop(s.id)}
                                      className="shrink-0 rounded p-1 text-slate-300 hover:bg-red-50 hover:text-[#A94442]"
                                    >
                                      <Trash2 className="h-3.5 w-3.5" />
                                    </button>
                                  </div>
                                  {s.description && <p className="mt-1 text-[11px] text-slate-500">{s.description}</p>}
                                  <div className="mt-2 flex items-center gap-2">
                                    <input
                                      defaultValue={s.notes}
                                      onBlur={(e) => {
                                        if (e.target.value !== s.notes) updateStepNotes(s, e.target.value);
                                      }}
                                      placeholder="填写执行备注..."
                                      className="flex-1 rounded border border-slate-200/80 bg-white px-2 py-1 text-[11px] focus:border-[#3B5B7E] focus:outline-none"
                                    />
                                    {s.completed_at && (
                                      <span className="text-[10px] text-[#2F6F5E] tabular-nums">{s.completed_at}</span>
                                    )}
                                  </div>
                                </div>
                              </div>
                            </div>
                          ))
                        )}
                      </div>
                    </div>
                  );
                })}

                <form onSubmit={handleAddStep} className="rounded-xl border border-slate-200 bg-slate-50/50 p-4 space-y-2.5">
                  <h3 className="text-xs font-semibold text-slate-700">新增 SOP 步骤</h3>
                  <div className="grid grid-cols-1 gap-2 sm:grid-cols-3">
                    <select
                      value={newStepStage}
                      onChange={(e) => setNewStepStage(e.target.value)}
                      className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs"
                    >
                      {STAGE_GROUPS.map((g) => (
                        <option key={g} value={g}>
                          {g}
                        </option>
                      ))}
                    </select>
                    <input
                      required
                      value={newStepTitle}
                      onChange={(e) => setNewStepTitle(e.target.value)}
                      placeholder="步骤标题 *"
                      className="sm:col-span-2 rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs"
                    />
                  </div>
                  <input
                    value={newStepDesc}
                    onChange={(e) => setNewStepDesc(e.target.value)}
                    placeholder="详细操作说明（可选）"
                    className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs"
                  />
                  <div className="flex flex-wrap items-center justify-between gap-2 pt-1">
                    <label className="inline-flex items-center gap-1.5 text-xs text-slate-600 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={syncNewStepToTemplate}
                        onChange={(e) => setSyncNewStepToTemplate(e.target.checked)}
                        className="rounded border-slate-300"
                      />
                      <span>同步加入全局默认模板（后续新建批次自动可见）</span>
                    </label>
                    <button
                      type="submit"
                      className="inline-flex items-center gap-1 rounded-lg bg-[#3B5B7E] px-3 py-1.5 text-xs font-medium text-white hover:bg-[#2E4866]"
                    >
                      <Plus className="h-3.5 w-3.5" /> 添加步骤
                    </button>
                  </div>
                </form>
              </div>

              <div className="lg:col-span-5 space-y-4">
                <div className="rounded-xl border border-slate-200 bg-white p-4 space-y-3">
                  <div className="flex items-center gap-2 border-b border-slate-100 pb-2">
                    <ShieldAlert className="h-4 w-4 text-[#9A6B26]" />
                    <h3 className="text-sm font-semibold text-slate-700">注意事项与避坑提醒</h3>
                  </div>
                  <div className="space-y-2">
                    {notices.map((n) => (
                      <div
                        key={n.id}
                        className={`rounded-lg border p-2.5 ${
                          n.priority === 'critical'
                            ? 'border-[#9A6B26]/25 bg-[#FEF9EC]/70'
                            : 'border-slate-200 bg-slate-50/70'
                        }`}
                      >
                        <div className="flex items-start justify-between gap-2">
                          <div className="text-xs font-semibold text-slate-800">{n.title}</div>
                          <button
                            onClick={() => handleDeleteSop(n.id)}
                            className="rounded p-0.5 text-slate-300 hover:text-[#A94442]"
                          >
                            <Trash2 className="h-3 w-3" />
                          </button>
                        </div>
                        {n.description && <p className="mt-1 text-[11px] text-slate-600">{n.description}</p>}
                      </div>
                    ))}
                  </div>
                  <form onSubmit={handleAddNotice} className="space-y-2 border-t border-slate-100 pt-2.5">
                    <div className="flex gap-2">
                      <select
                        value={newNoticePriority}
                        onChange={(e) => setNewNoticePriority(e.target.value as 'critical' | 'normal')}
                        className="rounded-lg border border-slate-200 bg-white px-2 py-1 text-xs"
                      >
                        <option value="critical">关键风控</option>
                        <option value="normal">普通提示</option>
                      </select>
                      <input
                        required
                        value={newNoticeTitle}
                        onChange={(e) => setNewNoticeTitle(e.target.value)}
                        placeholder="新增注意事项 *"
                        className="flex-1 rounded-lg border border-slate-200 px-2.5 py-1 text-xs"
                      />
                    </div>
                    <div className="flex gap-2">
                      <input
                        value={newNoticeDesc}
                        onChange={(e) => setNewNoticeDesc(e.target.value)}
                        placeholder="补充说明（可选）"
                        className="flex-1 rounded-lg border border-slate-200 px-2.5 py-1 text-xs"
                      />
                      <button
                        type="submit"
                        className="rounded-lg border border-slate-200 bg-slate-50 px-3 py-1 text-xs font-medium text-slate-700 hover:bg-slate-100"
                      >
                        添加
                      </button>
                    </div>
                  </form>
                </div>

                <div className="rounded-xl border border-slate-200 bg-white p-4 space-y-3">
                  <div className="border-b border-slate-100 pb-2 space-y-2">
                    <h3 className="text-sm font-semibold text-slate-700">其他阶段附件（评审报告 / 签到表等）</h3>
                    <div className="flex gap-2">
                      <select
                        value={otherCat}
                        onChange={(e) => setOtherCat(e.target.value)}
                        className="flex-1 rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs"
                      >
                        {OTHER_CATEGORIES.map((c) => (
                          <option key={c} value={c}>
                            {c}
                          </option>
                        ))}
                      </select>
                      <label className="inline-flex shrink-0 cursor-pointer items-center gap-1 rounded-lg bg-[#3B5B7E] px-3 py-1.5 text-xs font-medium text-white hover:bg-[#2E4866]">
                        <Upload className="h-3.5 w-3.5" />
                        <span>{uploadingOther ? '上传中...' : '上传'}</span>
                        <input type="file" onChange={handleOtherUpload} className="hidden" disabled={uploadingOther} />
                      </label>
                    </div>
                  </div>
                  <div className="divide-y divide-slate-100">
                    {(ws?.otherFiles || []).length === 0 ? (
                      <div className="py-5 text-center text-xs text-slate-400">暂无其他阶段附件</div>
                    ) : (
                      (ws?.otherFiles || []).map((f) => (
                        <div key={f.id} className="py-2 flex items-center justify-between gap-2">
                          <div className="flex items-center gap-2 min-w-0">
                            <FileSpreadsheet className="h-4 w-4 shrink-0 text-[#3B5B7E]" />
                            <div className="min-w-0">
                              <div className="truncate text-xs font-medium text-slate-800">{f.original_name}</div>
                              <div className="text-[10px] text-slate-400">{f.category}</div>
                            </div>
                          </div>
                          <div className="flex items-center gap-1 shrink-0">
                            {f.stored_path && (
                              <a href={`/api/files/${f.id}/download`} className="rounded p-1 text-slate-400 hover:text-slate-700">
                                <Download className="h-3.5 w-3.5" />
                              </a>
                            )}
                            <button onClick={() => handleDeleteFile(f.id)} className="rounded p-1 text-slate-400 hover:text-[#A94442]">
                              <Trash2 className="h-3.5 w-3.5" />
                            </button>
                          </div>
                        </div>
                      ))
                    )}
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Cross-Batch Mismatch Confirmation Modal */}
      {mismatchModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4 backdrop-blur-[1px]">
          <div className="w-full max-w-md rounded-xl border border-[#A94442]/30 bg-white p-5 shadow-xl space-y-4">
            <div className="flex items-start justify-between gap-2">
              <div className="flex items-center gap-2 text-[#A94442]">
                <AlertTriangle className="h-5 w-5 shrink-0" />
                <h3 className="text-sm font-semibold">检测到跨批次文件，已拦截导入</h3>
              </div>
              <button onClick={() => setMismatchModal(null)} className="rounded p-1 text-slate-400 hover:bg-slate-100">
                <X className="h-4 w-4" />
              </button>
            </div>
            <p className="text-xs text-slate-600 leading-relaxed">{mismatchModal.error}</p>
            <div className="rounded-lg bg-slate-50 p-3 text-xs space-y-1 font-mono">
              <div>当前工作台批次：<strong className="text-slate-800">{mismatchModal.currentBatchCode}</strong></div>
              <div>上传文件所属批次：<strong className="text-[#A94442]">{mismatchModal.fileBatchCode}</strong></div>
            </div>
            <div className="flex justify-end gap-2">
              <button
                onClick={() => setMismatchModal(null)}
                className="rounded-lg bg-[#3B5B7E] px-4 py-1.5 text-xs font-medium text-white hover:bg-[#2E4866]"
              >
                取消上传（防止传错）
              </button>
              <button
                onClick={handleConfirmForceUpload}
                className="rounded-lg border border-[#A94442]/30 bg-[#FEF2F2] px-3.5 py-1.5 text-xs font-medium text-[#A94442] hover:bg-[#FEE2E2]"
              >
                忽略预警，强制导入测试
              </button>
            </div>
          </div>
        </div>
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
