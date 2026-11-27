import { useCallback, useEffect, useState } from 'react';
import {
  AlertTriangle,
  CheckCircle2,
  ChevronDown,
  ChevronUp,
  ClipboardCheck,
  Download,
  FileSpreadsheet,
  Filter,
  Info,
  Play,
  RefreshCw,
  Search,
  Trash2,
  Upload,
  X,
} from 'lucide-react';
import { api } from '../api/client';
import { paginate, PaginationBar, usePagination } from './PaginationBar';
import type { ExpertPrecheckState, PrecheckObjectiveSelection } from '../types';

type ResultTab = 'total' | 'objective';
type ObjIssueFilter = 'all' | '分数不一致' | '分值步长不符' | '专家记录缺失' | '专家数量不足' | '细则未找到';

interface Props {
  batchId: string;
  rejectionCount?: number;
  hasRejectionFile?: boolean;
}

export function ExpertScorePrecheckPanel({ batchId, rejectionCount, hasRejectionFile }: Props) {
  const [state, setState] = useState<ExpertPrecheckState | null>(null);
  const [loading, setLoading] = useState(false);
  const [errorMsg, setErrorMsg] = useState('');

  // Upload queue (allows picking optional template before/with expert sheets)
  const [pendingExperts, setPendingExperts] = useState<File[]>([]);
  const [selectedPendingIndices, setSelectedPendingIndices] = useState<Set<number>>(new Set());
  const [selectedUploadedNames, setSelectedUploadedNames] = useState<Set<string>>(new Set());
  const [pendingTemplate, setPendingTemplate] = useState<File | null>(null);
  const [maxDiffInput, setMaxDiffInput] = useState('8');
  const [excludeRejected, setExcludeRejected] = useState(true);
  const [excludedExpanded, setExcludedExpanded] = useState(false);

  // Per-sheet objective selection map: `${sheetName}::${criterion}` -> { checked, step }
  const [activeSheetIdx, setActiveSheetIdx] = useState(0);
  const [selectionMap, setSelectionMap] = useState<Record<string, { checked: boolean; step: string }>>({});

  // Result view tabs & filters
  const [resultTab, setResultTab] = useState<ResultTab>('total');
  const [totalFilter, setTotalFilter] = useState<'all' | 'majority' | 'diff'>('all');
  const [totalSearch, setTotalSearch] = useState('');
  const [objFilter, setObjFilter] = useState<ObjIssueFilter>('all');
  const [objSearch, setObjSearch] = useState('');

  const totalPager = usePagination(`${batchId}-precheck-total-${totalFilter}-${totalSearch}`);
  const objPager = usePagination(`${batchId}-precheck-obj-${objFilter}-${objSearch}`);

  const syncFromState = useCallback((data: ExpertPrecheckState) => {
    setState(data);
    setMaxDiffInput(
      data.config?.maxScoreDiff !== undefined ? String(data.config.maxScoreDiff) : '8'
    );
    setExcludeRejected(data.config?.excludeRejected !== false);
    const nextMap: Record<string, { checked: boolean; step: string }> = {};
    for (const sel of data.config?.objectiveSelections || []) {
      nextMap[`${sel.sheetName}::${sel.criterion}`] = {
        checked: true,
        step: sel.step !== undefined && sel.step !== null ? String(sel.step) : '',
      };
    }
    setSelectionMap(nextMap);
    setActiveSheetIdx(0);
  }, []);

  const loadState = useCallback(async () => {
    if (!batchId) return;
    try {
      const data = await api.getScorePrecheckState(batchId);
      syncFromState(data);
    } catch {
      // ignore if offline
    }
  }, [batchId, syncFromState]);

  useEffect(() => {
    loadState();
    setErrorMsg('');
    setPendingExperts([]);
    setSelectedPendingIndices(new Set());
    setPendingTemplate(null);
  }, [loadState]);

  function handleAddPendingExperts(files: File[]) {
    const filtered = files.filter((f) => /\.(xlsx?|zip)$/i.test(f.name) && !f.name.startsWith('~$'));
    if (filtered.length === 0) return;
    setErrorMsg('');
    setPendingExperts((prev) => {
      const seen = new Set(prev.map((f) => `${f.name}:${f.size}`));
      const next = [...prev];
      for (const f of filtered) {
        const key = `${f.name}:${f.size}`;
        if (!seen.has(key)) {
          next.push(f);
          seen.add(key);
        }
      }
      return next;
    });
  }

  function removePendingExpert(idx: number) {
    setPendingExperts((prev) => prev.filter((_, i) => i !== idx));
    setSelectedPendingIndices(new Set());
  }

  function togglePendingSelect(idx: number) {
    setSelectedPendingIndices((prev) => {
      const next = new Set(prev);
      if (next.has(idx)) next.delete(idx);
      else next.add(idx);
      return next;
    });
  }

  function toggleSelectAllPending() {
    if (selectedPendingIndices.size === pendingExperts.length) {
      setSelectedPendingIndices(new Set());
    } else {
      setSelectedPendingIndices(new Set(pendingExperts.map((_, i) => i)));
    }
  }

  function removeSelectedPendingExperts() {
    if (selectedPendingIndices.size === 0) return;
    setPendingExperts((prev) => prev.filter((_, i) => !selectedPendingIndices.has(i)));
    setSelectedPendingIndices(new Set());
  }

  async function handleUploadAndRun() {
    if (pendingExperts.length === 0) {
      setErrorMsg('请先选择至少一份专家线下打分表（支持多选 .xlsx / .xls 或 .zip 压缩包）');
      return;
    }
    const maxScoreDiff = maxDiffInput.trim() === '' ? 8 : Number(maxDiffInput.trim());
    if (!Number.isFinite(maxScoreDiff) || maxScoreDiff < 0) {
      setErrorMsg('分差阈值必须为大于或等于 0 的数字');
      return;
    }

    setLoading(true);
    setErrorMsg('');
    try {
      const res = await api.uploadAndRunScorePrecheck(batchId, {
        experts: pendingExperts,
        template: pendingTemplate,
        maxScoreDiff,
        excludeRejected,
      });
      syncFromState(res);
      setPendingExperts([]);
      setSelectedPendingIndices(new Set());
      setPendingTemplate(null);
    } catch (err) {
      setErrorMsg((err as Error).message);
    } finally {
      setLoading(false);
    }
  }

  function buildSelectionsFromMap(): PrecheckObjectiveSelection[] | null {
    const sheets = state?.sheets || [];
    const selections: PrecheckObjectiveSelection[] = [];
    for (const s of sheets) {
      const allCols = Array.from(new Set([...s.scoringColumns, ...s.templateMarkedColumns]));
      for (const col of allCols) {
        const key = `${s.sheetName}::${col}`;
        const entry = selectionMap[key];
        if (!entry?.checked) continue;
        const stepStr = entry.step.trim();
        let stepNum: number | null = null;
        if (stepStr !== '') {
          const parsed = Number(stepStr);
          if (!Number.isFinite(parsed) || parsed <= 0) {
            setErrorMsg(`「${s.sheetName} / ${col}」的分值步长须留空或填写大于 0 的数字`);
            return null;
          }
          stepNum = parsed;
        }
        selections.push({ sheetName: s.sheetName, criterion: col, step: stepNum });
      }
    }
    return selections;
  }

  async function handleRerun() {
    const maxScoreDiff = maxDiffInput.trim() === '' ? 8 : Number(maxDiffInput.trim());
    if (!Number.isFinite(maxScoreDiff) || maxScoreDiff < 0) {
      setErrorMsg('分差阈值必须为大于或等于 0 的数字');
      return;
    }
    const selections = buildSelectionsFromMap();
    if (selections === null) return;

    setLoading(true);
    setErrorMsg('');
    try {
      const res = await api.rerunScorePrecheck(batchId, {
        maxScoreDiff,
        excludeRejected,
        objectiveSelections: selections,
      });
      syncFromState(res);
    } catch (err) {
      setErrorMsg((err as Error).message);
    } finally {
      setLoading(false);
    }
  }

  function toggleUploadedSelect(name: string) {
    setSelectedUploadedNames((prev) => {
      const next = new Set(prev);
      if (next.has(name)) next.delete(name);
      else next.add(name);
      return next;
    });
  }

  function toggleSelectAllUploaded() {
    const uploaded = state?.files || [];
    if (uploaded.length > 0 && selectedUploadedNames.size === uploaded.length) {
      setSelectedUploadedNames(new Set());
    } else {
      setSelectedUploadedNames(new Set(uploaded.map((f) => f.originalName)));
    }
  }

  async function handleRemoveUploadedFiles(fileNames: string[]) {
    if (fileNames.length === 0) return;
    const maxScoreDiff = maxDiffInput.trim() === '' ? 8 : Number(maxDiffInput.trim());
    if (!Number.isFinite(maxScoreDiff) || maxScoreDiff < 0) {
      setErrorMsg('分差阈值必须为大于或等于 0 的数字');
      return;
    }
    const selections = buildSelectionsFromMap();
    if (selections === null) return;

    setLoading(true);
    setErrorMsg('');
    try {
      const res = await api.removeScorePrecheckFiles(batchId, {
        fileNames,
        maxScoreDiff,
        excludeRejected,
        objectiveSelections: selections,
      });
      syncFromState(res);
      setSelectedUploadedNames(new Set());
    } catch (err) {
      setErrorMsg((err as Error).message);
    } finally {
      setLoading(false);
    }
  }

  async function handleAppendUploadedExperts(files: File[]) {
    const filtered = files.filter((f) => /\.(xlsx?|zip)$/i.test(f.name) && !f.name.startsWith('~$'));
    if (filtered.length === 0) return;
    const maxScoreDiff = maxDiffInput.trim() === '' ? 8 : Number(maxDiffInput.trim());
    if (!Number.isFinite(maxScoreDiff) || maxScoreDiff < 0) {
      setErrorMsg('分差阈值必须为大于或等于 0 的数字');
      return;
    }
    setLoading(true);
    setErrorMsg('');
    try {
      const res = await api.uploadAndRunScorePrecheck(batchId, {
        experts: filtered,
        maxScoreDiff,
        excludeRejected,
        append: true,
      });
      syncFromState(res);
      setSelectedUploadedNames(new Set());
    } catch (err) {
      setErrorMsg((err as Error).message);
    } finally {
      setLoading(false);
    }
  }

  async function handleClear() {
    if (!window.confirm('确认清空当前批次已上传的专家线下打分表与预校验结果吗？')) return;
    setLoading(true);
    setErrorMsg('');
    try {
      await api.clearScorePrecheck(batchId);
      setState(null);
      setSelectionMap({});
      setPendingExperts([]);
      setPendingTemplate(null);
      setActiveSheetIdx(0);
    } catch (err) {
      setErrorMsg((err as Error).message);
    } finally {
      setLoading(false);
    }
  }

  function toggleCriterion(sheetName: string, criterion: string) {
    const key = `${sheetName}::${criterion}`;
    setSelectionMap((prev) => ({
      ...prev,
      [key]: { checked: !prev[key]?.checked, step: prev[key]?.step || '' },
    }));
  }

  function setCriterionStep(sheetName: string, criterion: string, step: string) {
    const key = `${sheetName}::${criterion}`;
    setSelectionMap((prev) => ({
      ...prev,
      [key]: { checked: prev[key]?.checked ?? true, step },
    }));
  }

  function setSheetAllChecked(sheetName: string, columns: string[], checked: boolean) {
    setSelectionMap((prev) => {
      const next = { ...prev };
      for (const col of columns) {
        const key = `${sheetName}::${col}`;
        next[key] = { checked, step: next[key]?.step || '' };
      }
      return next;
    });
  }

  const sheets = state?.sheets || [];
  const result = state?.result || null;
  const totalSelectedCount = Object.values(selectionMap).filter((v) => v.checked).length;
  const effectiveRejectionCount = rejectionCount ?? state?.batchRejectionCount ?? 0;
  const effectiveHasRejectionFile =
    hasRejectionFile ?? state?.hasRejectionFile ?? effectiveRejectionCount > 0;

  const totalQuery = totalSearch.trim().toLowerCase();
  const filteredTotalItems = (result?.totalItems || []).filter((it) => {
    if (totalFilter === 'majority' && !it.majority_same) return false;
    if (totalFilter === 'diff' && !it.diff_exceeded) return false;
    if (totalQuery) {
      const hay = `${it.sheet_name} ${it.section} ${it.package_name} ${it.supplier} ${it.majority_same} ${it.diff_exceeded} ${it.expert_scores}`.toLowerCase();
      if (!hay.includes(totalQuery)) return false;
    }
    return true;
  });

  const objQuery = objSearch.trim().toLowerCase();
  const filteredObjItems = (result?.objectiveItems || []).filter((it) => {
    if (objFilter !== 'all' && it.issue_type !== objFilter) return false;
    if (objQuery) {
      const hay = `${it.sheet_name} ${it.section} ${it.package_name} ${it.supplier} ${it.criterion} ${it.issue_type} ${it.detail} ${it.source_location}`.toLowerCase();
      if (!hay.includes(objQuery)) return false;
    }
    return true;
  });

  return (
    <div className="rounded-xl border border-slate-200 bg-white overflow-hidden">
      {/* Header */}
      <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 bg-slate-50/80 px-4 py-2.5">
        <div className="flex items-center gap-2">
          <ClipboardCheck className="h-4 w-4 text-[#3B5B7E]" />
          <h3 className="text-sm font-semibold text-slate-800">专家打分预校验</h3>
          <span className="rounded bg-[#EBF2F7] px-1.5 py-0.5 text-[10px] font-medium text-[#3B5B7E]">
            线下打分表上传前预检
          </span>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          {state && state.files.length > 0 && (
            <button
              type="button"
              onClick={handleClear}
              disabled={loading}
              className="inline-flex items-center gap-1 rounded-md border border-[#A94442]/25 bg-[#FEF2F2]/70 px-2.5 py-1.5 text-xs font-medium text-[#A94442] hover:bg-[#FEE2E2] disabled:opacity-50"
            >
              <Trash2 className="h-3.5 w-3.5" />
              <span>清空</span>
            </button>
          )}
        </div>
      </div>

      {/* Error Banner */}
      {errorMsg && (
        <div className="mx-4 mt-3 flex items-start justify-between gap-2 rounded-lg border border-[#A94442]/30 bg-[#FEF2F2] px-3 py-2 text-xs text-[#A94442]">
          <div className="flex items-start gap-2">
            <AlertTriangle className="h-4 w-4 shrink-0 mt-0.5" />
            <span>{errorMsg}</span>
          </div>
          <button type="button" onClick={() => setErrorMsg('')} className="text-[11px] underline shrink-0">
            关闭
          </button>
        </div>
      )}

      <div className="p-4 space-y-4">
        {/* Step 1: Upload offline expert workbooks + optional base template */}
        <div className="rounded-lg border border-slate-200 bg-slate-50/70 p-3.5 space-y-3">
          <div className="flex items-start gap-2 text-[11px] text-slate-600">
            <Info className="h-3.5 w-3.5 shrink-0 text-[#3B5B7E] mt-0.5" />
            <span>
              在专家将线下打分表上传至 ECP/ETP 系统之前进行预校验。校验口径与线上<strong>「专家打分结果校验」</strong>保持严格一致：按应答人主键对齐、精确十进制运算（拦截空分/负数/公式错误/合计不一致）、总分校验<strong>超半数同分</strong>与<strong>分差严格大于阈值</strong>（默认 8 分）、客观分校验<strong>一致性 / 步长 / 缺漏</strong>。若上传了带表头填充色的基础模板，系统自动预勾选标色客观项。
            </span>
          </div>

          <div className="flex flex-wrap items-end gap-3">
            <div>
              <label className="block text-[11px] font-medium text-slate-600 mb-1">
                1. 专家线下打分表 *（多选或 ZIP）
              </label>
              <label className="inline-flex cursor-pointer items-center gap-1.5 rounded-md border border-slate-300 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:border-[#3B5B7E] hover:text-[#3B5B7E] transition-colors">
                <Upload className="h-3.5 w-3.5" />
                <span>
                  {pendingExperts.length > 0
                    ? `追加打分表（已选 ${pendingExperts.length} 份）`
                    : '选择专家打分表 / ZIP'}
                </span>
                <input
                  type="file"
                  multiple
                  accept=".xlsx,.xls,.zip"
                  onChange={(e) => {
                    const list = Array.from(e.target.files || []);
                    e.target.value = '';
                    if (list.length > 0) handleAddPendingExperts(list);
                  }}
                  className="hidden"
                  disabled={loading}
                />
              </label>
            </div>

            <div>
              <label className="block text-[11px] font-medium text-slate-600 mb-1">
                2. 基础模板（可选，自动识别标色客观项）
              </label>
              <div className="flex items-center gap-1.5">
                <label className="inline-flex cursor-pointer items-center gap-1.5 rounded-md border border-slate-300 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 hover:border-[#3B5B7E] hover:text-[#3B5B7E] transition-colors">
                  <FileSpreadsheet className="h-3.5 w-3.5" />
                  <span className="max-w-[180px] truncate">
                    {pendingTemplate ? pendingTemplate.name : '选择基础模板 (.xlsx)'}
                  </span>
                  <input
                    type="file"
                    accept=".xlsx"
                    onChange={(e) => {
                      const f = e.target.files?.[0] || null;
                      e.target.value = '';
                      setPendingTemplate(f);
                    }}
                    className="hidden"
                    disabled={loading}
                  />
                </label>
                {pendingTemplate && (
                  <button
                    type="button"
                    onClick={() => setPendingTemplate(null)}
                    className="text-[11px] text-slate-400 hover:text-[#A94442]"
                  >
                    移除
                  </button>
                )}
              </div>
            </div>

            <div>
              <label className="block text-[11px] font-medium text-slate-600 mb-1">
                分差阈值（严格大于触发）
              </label>
              <input
                type="number"
                min={0}
                step="0.5"
                value={maxDiffInput}
                onChange={(e) => setMaxDiffInput(e.target.value)}
                className="w-28 rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs tabular-nums focus:border-[#3B5B7E] focus:outline-none"
              />
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <button
                type="button"
                onClick={handleUploadAndRun}
                disabled={loading || pendingExperts.length === 0}
                className="inline-flex items-center gap-1.5 rounded-md bg-[#3B5B7E] px-3.5 py-1.5 text-xs font-medium text-white hover:bg-[#2E4866] disabled:opacity-50 transition-colors"
              >
                <Play className="h-3.5 w-3.5" />
                <span>{loading ? '预校验中...' : '上传并开始预校验'}</span>
              </button>

              {state && state.files.length > 0 && (
                <button
                  type="button"
                  onClick={handleRerun}
                  disabled={loading}
                  className="inline-flex items-center gap-1 rounded-md border border-[#3B5B7E]/30 bg-white px-2.5 py-1.5 text-xs font-medium text-[#3B5B7E] hover:bg-[#EBF2F7]/60 disabled:opacity-50"
                >
                  <RefreshCw className="h-3.5 w-3.5" />
                  <span>按当前参数重算</span>
                </button>
              )}
            </div>
          </div>

          {/* Pending Expert Files List with Single & Batch Delete */}
          {pendingExperts.length > 0 && (
            <div className="rounded-lg border border-slate-200 bg-white overflow-hidden">
              <div className="flex items-center justify-between gap-2 border-b border-slate-200 bg-slate-50/80 px-3 py-1.5 text-[11px]">
                <label className="inline-flex items-center gap-1.5 cursor-pointer select-none text-slate-600">
                  <input
                    type="checkbox"
                    checked={pendingExperts.length > 0 && selectedPendingIndices.size === pendingExperts.length}
                    onChange={toggleSelectAllPending}
                    disabled={loading}
                    className="h-3.5 w-3.5 rounded border-slate-300 accent-[#3B5B7E]"
                  />
                  <span>
                    {selectedPendingIndices.size > 0
                      ? `待上传列表：已勾选 ${selectedPendingIndices.size} / ${pendingExperts.length} 份`
                      : `待上传专家打分表（共 ${pendingExperts.length} 份）· 全选`}
                  </span>
                </label>
                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={removeSelectedPendingExperts}
                    disabled={loading || selectedPendingIndices.size === 0}
                    className="inline-flex items-center gap-1 rounded px-2 py-0.5 text-[11px] font-medium text-[#A94442] hover:bg-[#FEF2F2] disabled:opacity-40 disabled:hover:bg-transparent transition-colors"
                  >
                    <Trash2 className="h-3 w-3" />
                    <span>批量删除{selectedPendingIndices.size > 0 ? ` (${selectedPendingIndices.size})` : ''}</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      setPendingExperts([]);
                      setSelectedPendingIndices(new Set());
                    }}
                    disabled={loading}
                    className="text-[11px] text-slate-400 hover:text-[#A94442]"
                  >
                    清空列表
                  </button>
                </div>
              </div>
              <div className="max-h-48 overflow-y-auto divide-y divide-slate-100">
                {pendingExperts.map((f, idx) => {
                  const checked = selectedPendingIndices.has(idx);
                  return (
                    <div
                      key={`${f.name}-${f.size}-${idx}`}
                      className={`flex items-center justify-between gap-2 px-3 py-1.5 text-xs transition-colors ${
                        checked ? 'bg-[#EBF2F7]/50' : 'hover:bg-slate-50'
                      }`}
                    >
                      <label className="flex items-center gap-2 min-w-0 flex-1 cursor-pointer">
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => togglePendingSelect(idx)}
                          disabled={loading}
                          className="h-3.5 w-3.5 rounded border-slate-300 accent-[#3B5B7E] shrink-0"
                        />
                        <FileSpreadsheet className="h-3.5 w-3.5 text-[#3B5B7E] shrink-0" />
                        <span className="truncate text-slate-700" title={f.name}>
                          {f.name}
                        </span>
                        <span className="text-[10px] text-slate-400 tabular-nums shrink-0">
                          {(f.size / 1024).toFixed(0)} KB
                        </span>
                      </label>
                      <button
                        type="button"
                        onClick={() => removePendingExpert(idx)}
                        disabled={loading}
                        className="rounded p-0.5 text-slate-400 hover:text-[#A94442]"
                        title="移除"
                      >
                        <X className="h-3.5 w-3.5" />
                      </button>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* Auto-exclude initially rejected suppliers toggle */}
          <div className="flex flex-wrap items-center gap-2 border-t border-slate-200/70 pt-2.5 text-xs">
            <label className="inline-flex items-center gap-2 cursor-pointer select-none font-medium text-slate-700">
              <input
                type="checkbox"
                checked={excludeRejected}
                onChange={(e) => setExcludeRejected(e.target.checked)}
                className="rounded border-slate-300 cursor-pointer"
                disabled={loading}
              />
              <span>自动排除初评已否决供应商（命中《未进入详评情况统计表》标段+包号+供应商时跳过核验）</span>
            </label>
            {effectiveRejectionCount > 0 ? (
              <span className="rounded-full bg-[#EBF2F7] px-2 py-0.5 text-[10px] font-medium text-[#3B5B7E] tabular-nums">
                当前批次已入库 {effectiveRejectionCount} 条否决记录
              </span>
            ) : effectiveHasRejectionFile ? (
              <span className="rounded-full bg-[#ECFDF5] px-2 py-0.5 text-[10px] font-medium text-[#2F6F5E]">
                已上传《未进入详评情况统计表》（本批次 0 条否决记录）
              </span>
            ) : (
              <span className="rounded-full bg-slate-100 px-2 py-0.5 text-[10px] text-slate-500">
                当前批次尚未在上方槽位 4 上传《未进入详评情况统计表》
              </span>
            )}
          </div>

          {state && state.files.length > 0 && (
            <div className="rounded-lg border border-slate-200 bg-white overflow-hidden">
              <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 bg-slate-50/80 px-3 py-1.5 text-[11px]">
                <div className="flex flex-wrap items-center gap-2">
                  <label className="inline-flex items-center gap-1.5 cursor-pointer select-none text-slate-600">
                    <input
                      type="checkbox"
                      checked={state.files.length > 0 && selectedUploadedNames.size === state.files.length}
                      onChange={toggleSelectAllUploaded}
                      disabled={loading}
                      className="h-3.5 w-3.5 rounded border-slate-300 accent-[#3B5B7E]"
                    />
                    <span>
                      {selectedUploadedNames.size > 0
                        ? `已上传专家表：已勾选 ${selectedUploadedNames.size} / ${state.files.length} 份`
                        : `已上传专家打分表（共 ${state.files.length} 份）· 全选`}
                    </span>
                  </label>
                  <span className="text-slate-400">
                    {state.templateFile ? `模板：${state.templateFile.originalName}` : '自动识别标黄客观项'}
                    {result ? ` · ${result.sheetCount} 个工作表` : ''}
                  </span>
                </div>
                <div className="flex items-center gap-2">
                  <label className="inline-flex cursor-pointer items-center gap-1 rounded border border-[#3B5B7E]/30 bg-[#EBF2F7]/50 px-2 py-0.5 text-[11px] font-medium text-[#3B5B7E] hover:bg-[#EBF2F7] transition-colors">
                    <Upload className="h-3 w-3" />
                    <span>追加打分表并重算</span>
                    <input
                      type="file"
                      multiple
                      accept=".xlsx,.xls,.zip"
                      onChange={(e) => {
                        const list = Array.from(e.target.files || []);
                        e.target.value = '';
                        if (list.length > 0) handleAppendUploadedExperts(list);
                      }}
                      className="hidden"
                      disabled={loading}
                    />
                  </label>
                  <button
                    type="button"
                    onClick={() => handleRemoveUploadedFiles(Array.from(selectedUploadedNames))}
                    disabled={loading || selectedUploadedNames.size === 0}
                    className="inline-flex items-center gap-1 rounded px-2 py-0.5 text-[11px] font-medium text-[#A94442] hover:bg-[#FEF2F2] disabled:opacity-40 disabled:hover:bg-transparent transition-colors"
                  >
                    <Trash2 className="h-3 w-3" />
                    <span>删除选中并重算{selectedUploadedNames.size > 0 ? ` (${selectedUploadedNames.size})` : ''}</span>
                  </button>
                </div>
              </div>
              <div className="max-h-48 overflow-y-auto divide-y divide-slate-100">
                {state.files.map((f) => {
                  const checked = selectedUploadedNames.has(f.originalName);
                  return (
                    <div
                      key={f.originalName}
                      className={`flex items-center justify-between gap-2 px-3 py-1.5 text-xs transition-colors ${
                        checked ? 'bg-[#EBF2F7]/50' : 'hover:bg-slate-50'
                      }`}
                    >
                      <label className="flex items-center gap-2 min-w-0 flex-1 cursor-pointer">
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => toggleUploadedSelect(f.originalName)}
                          disabled={loading}
                          className="h-3.5 w-3.5 rounded border-slate-300 accent-[#3B5B7E] shrink-0"
                        />
                        <FileSpreadsheet className="h-3.5 w-3.5 text-[#3B5B7E] shrink-0" />
                        <span className="truncate text-slate-700" title={f.originalName}>
                          {f.originalName}
                        </span>
                        <span className="text-[10px] text-slate-400 tabular-nums shrink-0">
                          {(f.size / 1024).toFixed(0)} KB
                        </span>
                      </label>
                      <button
                        type="button"
                        onClick={() => handleRemoveUploadedFiles([f.originalName])}
                        disabled={loading}
                        className="rounded p-0.5 text-slate-400 hover:text-[#A94442]"
                        title="删除此文件并重新校验"
                      >
                        <X className="h-3.5 w-3.5" />
                      </button>
                    </div>
                  );
                })}
              </div>
            </div>
          )}
        </div>

        {/* Step 2: Per-Sheet Objective Criteria & Step Configuration */}
        {sheets.length > 0 && (
          <div className="rounded-lg border border-slate-200 bg-white overflow-hidden">
            <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 bg-slate-50/70 px-3.5 py-2">
              <div className="flex flex-wrap items-center gap-1.5">
                <span className="text-xs font-semibold text-slate-700 mr-1">
                  客观分核验项配置（已选 {totalSelectedCount} 项）：
                </span>
                {sheets.map((s, idx) => {
                  const allCols = Array.from(new Set([...s.scoringColumns, ...s.templateMarkedColumns]));
                  const checkedInSheet = allCols.filter((c) => selectionMap[`${s.sheetName}::${c}`]?.checked).length;
                  return (
                    <button
                      key={s.sheetName}
                      type="button"
                      onClick={() => setActiveSheetIdx(idx)}
                      className={`rounded-md px-2.5 py-1 text-[11px] font-medium tabular-nums transition-colors ${
                        activeSheetIdx === idx
                          ? 'bg-[#3B5B7E] text-white'
                          : 'bg-white border border-slate-200 text-slate-600 hover:bg-slate-100'
                      }`}
                    >
                      {s.sheetName}（{allCols.length}项{checkedInSheet > 0 ? ` · 已选${checkedInSheet}` : ''}）
                    </button>
                  );
                })}
              </div>

              <button
                type="button"
                onClick={handleRerun}
                disabled={loading}
                className="inline-flex items-center gap-1 rounded-md bg-[#2F6F5E] px-3 py-1.5 text-xs font-medium text-white hover:bg-[#24594B] disabled:opacity-50 transition-colors"
              >
                <RefreshCw className="h-3.5 w-3.5" />
                <span>{loading ? '校验中...' : '应用配置并重新校验'}</span>
              </button>
            </div>

            {(() => {
              const curSheet = sheets[Math.min(activeSheetIdx, sheets.length - 1)];
              if (!curSheet) return null;
              const allCols = Array.from(new Set([...curSheet.scoringColumns, ...curSheet.templateMarkedColumns]));
              const markedSet = new Set(curSheet.templateMarkedColumns);
              const detailMap = new Map(
                (curSheet.scoringColumnDetails || []).map((d) => [d.key, d])
              );
              const hasLevel2 = (curSheet.scoringColumnDetails || []).some((d) => Boolean(d.level2));
              const hasNumericMeta = (curSheet.scoringColumnDetails || []).some((d) =>
                Boolean(d.maxScore !== undefined || d.minScore !== undefined || d.stepHint !== undefined)
              );
              return (
                <div className="p-3 space-y-2.5">
                  <div className="flex flex-wrap items-center justify-between gap-2 text-[11px]">
                    <div className="text-slate-500">
                      总分汇总列：<strong className="text-slate-700">{curSheet.totalColumn}</strong>
                      {' · '}表头第 <strong className="text-slate-700">{curSheet.headerRow + 1}</strong> 行
                      {' · '}
                      <strong className="text-slate-700">
                        {hasLevel2 ? '一/二级要素展开' : '仅一级要素'}
                      </strong>
                      {hasNumericMeta && '（含最高/最低/极差行）'}
                      {curSheet.templateMarkedColumns.length > 0 && (
                        <>
                          {' · '}自动标色{' '}
                          <strong className="text-[#9A6B26]">{curSheet.templateMarkedColumns.length}</strong> 项
                        </>
                      )}
                    </div>
                    <div className="flex items-center gap-1.5 shrink-0">
                      <button
                        type="button"
                        onClick={() => setSheetAllChecked(curSheet.sheetName, allCols, true)}
                        className="rounded border border-slate-200 bg-slate-50 px-2 py-0.5 text-[11px] text-slate-600 hover:bg-slate-100"
                      >
                        本页全选
                      </button>
                      <button
                        type="button"
                        onClick={() => setSheetAllChecked(curSheet.sheetName, allCols, false)}
                        className="rounded border border-slate-200 bg-slate-50 px-2 py-0.5 text-[11px] text-slate-600 hover:bg-slate-100"
                      >
                        本页全不选
                      </button>
                    </div>
                  </div>

                  <div className="max-h-72 overflow-y-auto rounded-lg border border-slate-200">
                    <table className="w-full border-collapse text-left text-xs">
                      <thead className="sticky top-0 bg-slate-50 text-slate-500 border-b border-slate-200 z-10">
                        <tr>
                          <th className="px-2.5 py-1.5 w-11 text-center font-medium">核验</th>
                          <th className="px-2 py-1.5 w-12 text-center font-medium whitespace-nowrap">列</th>
                          <th className={`px-2.5 py-1.5 font-medium ${hasLevel2 ? 'w-[30%]' : ''}`}>
                            {hasLevel2 ? '一级要素' : '评分要素'}
                          </th>
                          {hasLevel2 && (
                            <th className="px-2.5 py-1.5 font-medium">二级要素 / 评分细则</th>
                          )}
                          {hasNumericMeta && (
                            <th className="px-2.5 py-1.5 w-32 font-medium whitespace-nowrap">分值 / 极差</th>
                          )}
                          <th className="px-2.5 py-1.5 w-20 font-medium whitespace-nowrap">标色</th>
                          <th className="px-2.5 py-1.5 w-32 font-medium whitespace-nowrap">分值步长（可空）</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {allCols.map((col) => {
                          const key = `${curSheet.sheetName}::${col}`;
                          const entry = selectionMap[key];
                          const checked = Boolean(entry?.checked);
                          const isMarked = markedSet.has(col);
                          const meta = detailMap.get(col);
                          const rangeParts: string[] = [];
                          if (meta?.minScore !== undefined && meta?.maxScore !== undefined) {
                            rangeParts.push(`${meta.minScore}~${meta.maxScore}分`);
                          } else if (meta?.maxScore !== undefined) {
                            rangeParts.push(`最高${meta.maxScore}`);
                          }
                          if (meta?.stepHint !== undefined) {
                            rangeParts.push(`极差${meta.stepHint}`);
                          }
                          return (
                            <tr key={col} className={checked ? 'bg-[#EBF2F7]/40' : 'hover:bg-slate-50/70'}>
                              <td className="px-2.5 py-1.5 text-center align-top">
                                <input
                                  type="checkbox"
                                  checked={checked}
                                  onChange={() => toggleCriterion(curSheet.sheetName, col)}
                                  className="mt-0.5 rounded border-slate-300 cursor-pointer"
                                />
                              </td>
                              <td
                                className="px-2 py-1.5 text-center align-top cursor-pointer"
                                onClick={() => toggleCriterion(curSheet.sheetName, col)}
                              >
                                <span className="inline-block rounded bg-slate-100 px-1.5 py-0.5 font-mono text-[10px] font-medium text-slate-600">
                                  {meta?.colLetter || '—'}
                                </span>
                              </td>
                              <td
                                className="px-2.5 py-1.5 text-slate-800 font-medium cursor-pointer align-top"
                                onClick={() => toggleCriterion(curSheet.sheetName, col)}
                              >
                                {meta?.level1 || col}
                              </td>
                              {hasLevel2 && (
                                <td
                                  className="px-2.5 py-1.5 text-slate-700 cursor-pointer align-top"
                                  onClick={() => toggleCriterion(curSheet.sheetName, col)}
                                  title={[meta?.level2, meta?.detail].filter(Boolean).join('\n')}
                                >
                                  <div className="line-clamp-2 leading-snug">
                                    {meta?.level2 || <span className="text-slate-300">—</span>}
                                  </div>
                                  {meta?.detail && (
                                    <div className="mt-0.5 line-clamp-1 text-[11px] text-slate-400">
                                      {meta.detail}
                                    </div>
                                  )}
                                </td>
                              )}
                              {hasNumericMeta && (
                                <td className="px-2.5 py-1.5 text-[11px] text-slate-500 tabular-nums whitespace-nowrap align-top">
                                  {rangeParts.length > 0 ? rangeParts.join(' · ') : '—'}
                                </td>
                              )}
                              <td className="px-2.5 py-1.5 align-top whitespace-nowrap">
                                {isMarked ? (
                                  <span className="rounded bg-[#FEF9EC] border border-[#9A6B26]/30 px-1.5 py-0.5 text-[10px] font-medium text-[#9A6B26]">
                                    标色
                                  </span>
                                ) : (
                                  <span className="text-[10px] text-slate-400">未标色</span>
                                )}
                              </td>
                              <td className="px-2.5 py-1.5 align-top">
                                <input
                                  type="number"
                                  min={0}
                                  step="0.5"
                                  value={entry?.step || ''}
                                  onChange={(e) => setCriterionStep(curSheet.sheetName, col, e.target.value)}
                                  placeholder={meta?.stepHint ? `表头极差 ${meta.stepHint}` : '如 1 或 0.5'}
                                  className="w-28 rounded border border-slate-200 bg-white px-2 py-1 text-[11px] tabular-nums focus:border-[#3B5B7E] focus:outline-none"
                                />
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              );
            })()}
          </div>
        )}

        {/* Auto-Excluded Rejected Suppliers Banner */}
        {result && (result.excludedRejectedCount || 0) > 0 && (
          <div className="rounded-lg border border-[#3B5B7E]/25 bg-[#EBF2F7]/50 overflow-hidden">
            <button
              type="button"
              onClick={() => setExcludedExpanded((v) => !v)}
              className="flex w-full items-center justify-between gap-2 px-3 py-2 text-left text-xs font-medium text-[#3B5B7E] hover:bg-[#EBF2F7]/80 transition-colors"
            >
              <div className="flex items-center gap-1.5">
                <Filter className="h-3.5 w-3.5 shrink-0" />
                <span>
                  已根据《未进入详评情况统计表》自动排除 <strong>{result.excludedRejectedCount}</strong> 条初评已否决应答人记录（空白/默认打分不触发异常，与上传系统后口径一致）
                </span>
              </div>
              <span className="inline-flex items-center gap-0.5 text-[11px] shrink-0">
                {excludedExpanded ? '收起明细' : '查看明细'}
                {excludedExpanded ? <ChevronUp className="h-3.5 w-3.5" /> : <ChevronDown className="h-3.5 w-3.5" />}
              </span>
            </button>
            {excludedExpanded && result.excludedRejectedItems && result.excludedRejectedItems.length > 0 && (
              <div className="border-t border-[#3B5B7E]/15 bg-white overflow-x-auto">
                <table className="w-full border-collapse text-left text-[11px]">
                  <thead>
                    <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                      <th className="px-2.5 py-1.5 font-medium whitespace-nowrap">工作表</th>
                      <th className="px-2.5 py-1.5 font-medium whitespace-nowrap">标段 / 包号</th>
                      <th className="px-2.5 py-1.5 font-medium whitespace-nowrap">已排除应答人</th>
                      <th className="px-2.5 py-1.5 font-medium whitespace-nowrap">否决类别</th>
                      <th className="px-2.5 py-1.5 font-medium">未进入详评原因</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {result.excludedRejectedItems.map((ex, i) => (
                      <tr key={i} className="hover:bg-slate-50/60">
                        <td className="px-2.5 py-1.5 text-slate-700 font-medium whitespace-nowrap">{ex.sheet_name}</td>
                        <td className="px-2.5 py-1.5 text-slate-600 whitespace-nowrap">
                          {[ex.section, ex.package_name].filter(Boolean).join(' · ') || '—'}
                        </td>
                        <td className="px-2.5 py-1.5 font-medium text-slate-800 whitespace-nowrap">{ex.supplier}</td>
                        <td className="px-2.5 py-1.5 whitespace-nowrap">
                          <span className="rounded bg-slate-100 px-1.5 py-0.5 text-[10px] text-slate-600">
                            {ex.review_type}
                          </span>
                        </td>
                        <td className="px-2.5 py-1.5 text-slate-500">{ex.reason_category}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}

        {/* Warnings */}
        {result && result.warnings.length > 0 && (
          <div className="rounded-lg border border-[#9A6B26]/30 bg-[#FEF9EC] px-3 py-2 text-xs text-[#9A6B26] space-y-1">
            <div className="font-semibold">核验提示与一致性预警（{result.warnings.length} 条）：</div>
            {result.warnings.map((w, i) => (
              <div key={i}>• {w}</div>
            ))}
          </div>
        )}

        {/* Step 3: Results Tabs (Total Score & Objective Score) */}
        {result && (
          <div className="space-y-3">
            <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 pb-2">
              <div className="flex flex-wrap gap-1.5">
                <button
                  type="button"
                  onClick={() => setResultTab('total')}
                  className={`inline-flex items-center gap-1.5 rounded-md px-3 py-1.5 text-xs font-medium transition-colors ${
                    resultTab === 'total'
                      ? 'bg-[#3B5B7E] text-white'
                      : 'bg-slate-100 text-slate-600 hover:bg-slate-200/70'
                  }`}
                >
                  <span>1. 总分预校验结果</span>
                  <span
                    className={`rounded-full px-1.5 text-[10px] tabular-nums ${
                      result.counts.totalIssues > 0
                        ? 'bg-[#A94442] text-white'
                        : resultTab === 'total'
                          ? 'bg-white/20 text-white'
                          : 'bg-[#ECFDF5] text-[#2F6F5E]'
                    }`}
                  >
                    {result.counts.totalIssues}
                  </span>
                </button>

                <button
                  type="button"
                  onClick={() => setResultTab('objective')}
                  className={`inline-flex items-center gap-1.5 rounded-md px-3 py-1.5 text-xs font-medium transition-colors ${
                    resultTab === 'objective'
                      ? 'bg-[#3B5B7E] text-white'
                      : 'bg-slate-100 text-slate-600 hover:bg-slate-200/70'
                  }`}
                >
                  <span>2. 客观分预校验结果</span>
                  <span
                    className={`rounded-full px-1.5 text-[10px] tabular-nums ${
                      result.counts.objectiveIssues > 0
                        ? 'bg-[#A94442] text-white'
                        : resultTab === 'objective'
                          ? 'bg-white/20 text-white'
                          : 'bg-[#ECFDF5] text-[#2F6F5E]'
                    }`}
                  >
                    {result.counts.objectiveIssues}
                  </span>
                </button>
              </div>

              <div className="flex flex-wrap items-center gap-2.5">
                <div className="text-[11px] text-slate-500 tabular-nums">
                  已核验总分 <strong className="text-slate-700">{result.totalCheckedCount}</strong> 条 · 客观分{' '}
                  <strong className="text-slate-700">{result.objectiveCheckedCount}</strong> 组 · 参与专家：
                  {result.expertNames.join('、')}
                </div>
                <a
                  href={`/api/evaluation/batch/${batchId}/score-precheck/export`}
                  className="inline-flex items-center gap-1.5 rounded-md bg-[#2F6F5E] px-3 py-1.5 text-xs font-medium text-white hover:bg-[#24594B] transition-colors shrink-0"
                >
                  <Download className="h-3.5 w-3.5" />
                  <span>导出预校验报告 (.xlsx)</span>
                </a>
              </div>
            </div>

            {/* Result Tab 1: Total Score */}
            {resultTab === 'total' && (
              <div className="space-y-3">
                <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
                  <div className="flex flex-wrap gap-1.5">
                    {([
                      { key: 'all', label: `全部异常 (${result.counts.totalIssues})` },
                      { key: 'majority', label: `超过半数同分 (${result.counts.majoritySame})` },
                      { key: 'diff', label: `分差超过阈值 (${result.counts.diffExceeded})` },
                    ] as const).map((t) => (
                      <button
                        key={t.key}
                        type="button"
                        onClick={() => setTotalFilter(t.key)}
                        className={`rounded-md px-2.5 py-1 text-xs font-medium tabular-nums transition-colors ${
                          totalFilter === t.key
                            ? t.key === 'all'
                              ? 'bg-slate-800 text-white'
                              : 'bg-[#A94442] text-white'
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
                      value={totalSearch}
                      onChange={(e) => setTotalSearch(e.target.value)}
                      placeholder="搜索工作表、标段、包号或应答人..."
                      className="w-full rounded-lg border border-slate-200 py-1.5 pl-8 pr-3 text-xs focus:border-[#3B5B7E] focus:outline-none"
                    />
                  </div>
                </div>

                <div className="overflow-hidden rounded-lg border border-slate-200">
                  <div className="overflow-x-auto">
                    <table className="w-full border-collapse text-left text-xs">
                      <thead>
                        <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                          <th className="px-2.5 py-2 font-medium whitespace-nowrap">工作表 / 标段 / 包号</th>
                          <th className="px-2.5 py-2 font-medium whitespace-nowrap">应答人</th>
                          <th className="px-2.5 py-2 font-medium min-w-[180px]">超过半数同分</th>
                          <th className="px-2.5 py-2 font-medium min-w-[240px]">分差超过阈值</th>
                          <th className="px-2.5 py-2 font-medium min-w-[180px]">各专家精确总分</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {filteredTotalItems.length === 0 ? (
                          <tr>
                            <td colSpan={5} className="py-8 text-center text-xs text-[#2F6F5E]">
                              <span className="inline-flex items-center gap-1.5">
                                <CheckCircle2 className="h-4 w-4" />
                                {result.counts.totalIssues === 0
                                  ? `已核验 ${result.totalCheckedCount} 条应答人总分记录，未发现超半数同分或分差严格大于 ${result.maxScoreDiff} 分的异常`
                                  : '当前筛选条件下无匹配记录'}
                              </span>
                            </td>
                          </tr>
                        ) : (
                          paginate(filteredTotalItems, totalPager.page, totalPager.pageSize).map((item, idx) => (
                            <tr key={idx} className="hover:bg-slate-50/70 align-top">
                              <td className="px-2.5 py-2 whitespace-nowrap">
                                <div className="font-medium text-slate-800">{item.sheet_name}</div>
                                <div className="text-[11px] text-slate-400">
                                  {[item.section, item.package_name].filter(Boolean).join(' · ') || '—'}
                                </div>
                              </td>
                              <td className="px-2.5 py-2 font-medium text-slate-800">{item.supplier}</td>
                              <td className="px-2.5 py-2 text-[#9A6B26] whitespace-pre-line">
                                {item.majority_same || <span className="text-slate-300">—</span>}
                              </td>
                              <td className="px-2.5 py-2 text-[#A94442] whitespace-pre-line leading-relaxed">
                                {item.diff_exceeded || <span className="text-slate-300">—</span>}
                              </td>
                              <td className="px-2.5 py-2 font-mono text-[11px] text-slate-600">
                                {item.expert_scores}
                              </td>
                            </tr>
                          ))
                        )}
                      </tbody>
                    </table>
                  </div>
                  {filteredTotalItems.length > 0 && (
                    <PaginationBar
                      total={filteredTotalItems.length}
                      page={totalPager.page}
                      pageSize={totalPager.pageSize}
                      onPageChange={totalPager.setPage}
                      onPageSizeChange={totalPager.setPageSize}
                    />
                  )}
                </div>
              </div>
            )}

            {/* Result Tab 2: Objective Score */}
            {resultTab === 'objective' && (
              <div className="space-y-3">
                <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
                  <div className="flex flex-wrap gap-1.5">
                    {([
                      { key: 'all', label: `全部问题 (${result.counts.objectiveIssues})` },
                      { key: '分数不一致', label: `分数不一致 (${result.counts.scoreMismatch})` },
                      { key: '分值步长不符', label: `分值步长不符 (${result.counts.stepMismatch})` },
                      { key: '专家记录缺失', label: `专家记录缺失 (${result.counts.expertMissing})` },
                      { key: '专家数量不足', label: `专家数量不足 (${result.counts.expertInsufficient})` },
                      { key: '细则未找到', label: `细则未找到 (${result.counts.criterionMissing})` },
                    ] as const).map((t) => (
                      <button
                        key={t.key}
                        type="button"
                        onClick={() => setObjFilter(t.key)}
                        className={`rounded-md px-2.5 py-1 text-xs font-medium tabular-nums transition-colors ${
                          objFilter === t.key
                            ? t.key === 'all'
                              ? 'bg-slate-800 text-white'
                              : 'bg-[#A94442] text-white'
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
                      value={objSearch}
                      onChange={(e) => setObjSearch(e.target.value)}
                      placeholder="搜索工作表、应答人或客观项..."
                      className="w-full rounded-lg border border-slate-200 py-1.5 pl-8 pr-3 text-xs focus:border-[#3B5B7E] focus:outline-none"
                    />
                  </div>
                </div>

                <div className="overflow-hidden rounded-lg border border-slate-200">
                  <div className="overflow-x-auto">
                    <table className="w-full border-collapse text-left text-xs">
                      <thead>
                        <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                          <th className="px-2.5 py-2 font-medium whitespace-nowrap">工作表 / 标段 / 包号</th>
                          <th className="px-2.5 py-2 font-medium whitespace-nowrap">应答人</th>
                          <th className="px-2.5 py-2 font-medium min-w-[160px]">客观评分细则</th>
                          <th className="px-2.5 py-2 font-medium whitespace-nowrap">问题类型</th>
                          <th className="px-2.5 py-2 font-medium min-w-[200px]">各专家打分说明</th>
                          <th className="px-2.5 py-2 font-medium min-w-[180px]">来源位置</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {filteredObjItems.length === 0 ? (
                          <tr>
                            <td colSpan={6} className="py-8 text-center text-xs text-[#2F6F5E]">
                              <span className="inline-flex items-center gap-1.5">
                                <CheckCircle2 className="h-4 w-4" />
                                {totalSelectedCount === 0
                                  ? '当前尚未勾选任何客观评分项，请在上方勾选客观项后点击「应用配置并重新校验」'
                                  : result.counts.objectiveIssues === 0
                                    ? `已核验 ${result.objectiveCheckedCount} 组客观评分项，各专家打分完全一致且符合步长要求`
                                    : '当前筛选条件下无匹配记录'}
                              </span>
                            </td>
                          </tr>
                        ) : (
                          paginate(filteredObjItems, objPager.page, objPager.pageSize).map((item, idx) => (
                            <tr key={idx} className="hover:bg-slate-50/70 align-top">
                              <td className="px-2.5 py-2 whitespace-nowrap">
                                <div className="font-medium text-slate-800">{item.sheet_name}</div>
                                <div className="text-[11px] text-slate-400">
                                  {[item.section, item.package_name].filter(Boolean).join(' · ') || '—'}
                                </div>
                              </td>
                              <td className="px-2.5 py-2 font-medium text-slate-800">{item.supplier || '—'}</td>
                              <td className="px-2.5 py-2 text-slate-700 font-medium">{item.criterion}</td>
                              <td className="px-2.5 py-2 whitespace-nowrap">
                                <span
                                  className={`rounded px-1.5 py-0.5 text-[11px] font-medium ${
                                    item.issue_type === '分数不一致' || item.issue_type === '分值步长不符'
                                      ? 'bg-[#FEF2F2] text-[#A94442]'
                                      : 'bg-[#FEF9EC] text-[#9A6B26]'
                                  }`}
                                >
                                  {item.issue_type}
                                </span>
                              </td>
                              <td className="px-2.5 py-2 text-slate-700 whitespace-pre-line font-mono text-[11px] leading-relaxed">
                                {item.detail}
                              </td>
                              <td className="px-2.5 py-2 text-slate-400 whitespace-pre-line font-mono text-[10px] leading-relaxed">
                                {item.source_location}
                              </td>
                            </tr>
                          ))
                        )}
                      </tbody>
                    </table>
                  </div>
                  {filteredObjItems.length > 0 && (
                    <PaginationBar
                      total={filteredObjItems.length}
                      page={objPager.page}
                      pageSize={objPager.pageSize}
                      onPageChange={objPager.setPage}
                      onPageSizeChange={objPager.setPageSize}
                    />
                  )}
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
