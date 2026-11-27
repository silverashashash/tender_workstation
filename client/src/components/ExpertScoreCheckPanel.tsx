import { useCallback, useEffect, useState } from 'react';
import {
  AlertTriangle,
  CheckCircle2,
  ClipboardCheck,
  Download,
  FileSpreadsheet,
  Info,
  Play,
  RefreshCw,
  Search,
  Trash2,
  Upload,
} from 'lucide-react';
import { api } from '../api/client';
import { paginate, PaginationBar, usePagination } from './PaginationBar';
import type { ObjectiveSelection, ScoreCheckState } from '../types';

type SubMode = 'total' | 'objective';
type ObjIssueFilter = 'all' | '分数不一致' | '分值步长不符' | '专家记录缺失' | '专家数量不足' | '细则未找到';

interface Props {
  batchId: string;
}

export function ExpertScoreCheckPanel({ batchId }: Props) {
  const [subMode, setSubMode] = useState<SubMode>('total');
  const [state, setState] = useState<ScoreCheckState>({ total: null, objective: null });
  const [loading, setLoading] = useState(false);
  const [errorMsg, setErrorMsg] = useState('');

  // --- Tool 8: Total Score states ---
  const [expertNumInput, setExpertNumInput] = useState(''); // empty = 自动识别
  const [maxDiffInput, setMaxDiffInput] = useState('8');
  const [passwordInput, setPasswordInput] = useState('');
  const [totalFilter, setTotalFilter] = useState<'all' | 'majority' | 'diff'>('all');
  const [totalSearch, setTotalSearch] = useState('');
  const totalPager = usePagination(`${batchId}-total-${totalFilter}-${totalSearch}`);

  // --- Tool 9: Objective Score states ---
  const [activeMethodIdx, setActiveMethodIdx] = useState(0);
  // key = `${methodId}::${criterion}` -> { checked, step }
  const [selectionMap, setSelectionMap] = useState<Record<string, { checked: boolean; step: string }>>({});
  const [sectionOverride, setSectionOverride] = useState('');
  const [objFilter, setObjFilter] = useState<ObjIssueFilter>('all');
  const [objSearch, setObjSearch] = useState('');
  const objPager = usePagination(`${batchId}-obj-${objFilter}-${objSearch}`);

  const syncFromState = useCallback((data: ScoreCheckState) => {
    setState(data);
    if (data.total?.config) {
      setExpertNumInput(data.total.config.expertNum ? String(data.total.config.expertNum) : '');
      setMaxDiffInput(
        data.total.config.maxScoreDiff !== undefined ? String(data.total.config.maxScoreDiff) : '8'
      );
    }
    if (data.objective?.config) {
      setSectionOverride(data.objective.config.sectionOverride || '');
      const nextMap: Record<string, { checked: boolean; step: string }> = {};
      for (const s of data.objective.config.selections || []) {
        nextMap[`${s.methodId}::${s.criterion}`] = {
          checked: true,
          step: s.step !== undefined && s.step !== null ? String(s.step) : '',
        };
      }
      setSelectionMap(nextMap);
    } else {
      setSelectionMap({});
      setSectionOverride('');
    }
    setActiveMethodIdx(0);
  }, []);

  const loadState = useCallback(async () => {
    if (!batchId) return;
    try {
      const data = await api.getScoreCheckState(batchId);
      syncFromState(data);
    } catch {
      // ignore if offline
    }
  }, [batchId, syncFromState]);

  useEffect(() => {
    loadState();
    setErrorMsg('');
  }, [loadState]);

  // --- Total Score Handlers ---
  async function handleTotalUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const files = Array.from(e.target.files || []);
    e.target.value = '';
    if (files.length === 0) return;
    setLoading(true);
    setErrorMsg('');
    try {
      const expertNum = expertNumInput.trim() ? Number(expertNumInput.trim()) : null;
      const maxScoreDiff = maxDiffInput.trim() === '' ? 8 : Number(maxDiffInput.trim());
      const res = await api.uploadAndRunTotalScoreCheck(batchId, files, {
        expertNum,
        maxScoreDiff,
        password: passwordInput,
      });
      setState((prev) => ({ ...prev, total: res }));
    } catch (err) {
      setErrorMsg((err as Error).message);
    } finally {
      setLoading(false);
    }
  }

  async function handleTotalRerun() {
    setLoading(true);
    setErrorMsg('');
    try {
      const expertNum = expertNumInput.trim() ? Number(expertNumInput.trim()) : null;
      const maxScoreDiff = maxDiffInput.trim() === '' ? 8 : Number(maxDiffInput.trim());
      const res = await api.rerunTotalScoreCheck(batchId, {
        expertNum,
        maxScoreDiff,
        password: passwordInput,
      });
      setState((prev) => ({ ...prev, total: res }));
    } catch (err) {
      setErrorMsg((err as Error).message);
    } finally {
      setLoading(false);
    }
  }

  async function handleTotalClear() {
    if (!window.confirm('确认清空当前批次已上传的专家总分统计表与校验结果吗？')) return;
    setLoading(true);
    setErrorMsg('');
    try {
      await api.clearTotalScoreCheck(batchId);
      setState((prev) => ({ ...prev, total: null }));
    } catch (err) {
      setErrorMsg((err as Error).message);
    } finally {
      setLoading(false);
    }
  }

  // --- Objective Score Handlers ---
  async function handleObjectiveUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const files = Array.from(e.target.files || []);
    e.target.value = '';
    if (files.length === 0) return;
    setLoading(true);
    setErrorMsg('');
    try {
      const res = await api.uploadObjectiveScoreFiles(batchId, files);
      setState((prev) => ({ ...prev, objective: res }));
      setSelectionMap({});
      setActiveMethodIdx(0);
    } catch (err) {
      setErrorMsg((err as Error).message);
    } finally {
      setLoading(false);
    }
  }

  function toggleCriterion(methodId: string, criterion: string) {
    const key = `${methodId}::${criterion}`;
    setSelectionMap((prev) => {
      const cur = prev[key];
      return {
        ...prev,
        [key]: { checked: !cur?.checked, step: cur?.step || '' },
      };
    });
  }

  function setCriterionStep(methodId: string, criterion: string, step: string) {
    const key = `${methodId}::${criterion}`;
    setSelectionMap((prev) => ({
      ...prev,
      [key]: { checked: prev[key]?.checked ?? true, step },
    }));
  }

  function setMethodAllChecked(methodId: string, criteriaNames: string[], checked: boolean) {
    setSelectionMap((prev) => {
      const next = { ...prev };
      for (const name of criteriaNames) {
        const key = `${methodId}::${name}`;
        next[key] = { checked, step: next[key]?.step || '' };
      }
      return next;
    });
  }

  async function handleObjectiveRun() {
    const methods = state.objective?.methods || [];
    const selections: ObjectiveSelection[] = [];
    for (const m of methods) {
      for (const c of m.criteria) {
        const key = `${m.id}::${c.name}`;
        const entry = selectionMap[key];
        if (!entry?.checked) continue;
        const stepStr = entry.step.trim();
        let stepNum: number | null = null;
        if (stepStr !== '') {
          const parsed = Number(stepStr);
          if (!Number.isFinite(parsed) || parsed <= 0) {
            setErrorMsg(`「${c.name.slice(0, 20)}…」的分值步长须留空或填写大于 0 的数字`);
            return;
          }
          stepNum = parsed;
        }
        selections.push({ methodId: m.id, criterion: c.name, step: stepNum });
      }
    }
    if (selections.length === 0) {
      setErrorMsg('请至少勾选一项客观评分细则后再执行校验');
      return;
    }

    setLoading(true);
    setErrorMsg('');
    try {
      const res = await api.runObjectiveScoreCheck(batchId, {
        selections,
        sectionOverride: sectionOverride.trim(),
      });
      setState((prev) => ({ ...prev, objective: res }));
    } catch (err) {
      setErrorMsg((err as Error).message);
    } finally {
      setLoading(false);
    }
  }

  async function handleObjectiveClear() {
    if (!window.confirm('确认清空当前批次已上传的专家评分明细表与客观分校验结果吗？')) return;
    setLoading(true);
    setErrorMsg('');
    try {
      await api.clearObjectiveScoreCheck(batchId);
      setState((prev) => ({ ...prev, objective: null }));
      setSelectionMap({});
      setActiveMethodIdx(0);
    } catch (err) {
      setErrorMsg((err as Error).message);
    } finally {
      setLoading(false);
    }
  }

  // Filtered total score results
  const totalResult = state.total?.result || null;
  const totalQuery = totalSearch.trim().toLowerCase();
  const filteredTotalItems = (totalResult?.items || []).filter((it) => {
    if (totalFilter === 'majority' && !it.majority_same) return false;
    if (totalFilter === 'diff' && !it.diff_exceeded) return false;
    if (totalQuery) {
      const hay = `${it.section} ${it.package_name} ${it.supplier} ${it.majority_same} ${it.diff_exceeded} ${it.source_file}`.toLowerCase();
      if (!hay.includes(totalQuery)) return false;
    }
    return true;
  });

  // Filtered objective score results
  const objMethods = state.objective?.methods || [];
  const objResult = state.objective?.result || null;
  const objQuery = objSearch.trim().toLowerCase();
  const filteredObjItems = (objResult?.items || []).filter((it) => {
    if (objFilter !== 'all' && it.issue_type !== objFilter) return false;
    if (objQuery) {
      const hay = `${it.section} ${it.package_name} ${it.supplier} ${it.criterion} ${it.issue_type} ${it.detail} ${it.source_location}`.toLowerCase();
      if (!hay.includes(objQuery)) return false;
    }
    return true;
  });

  const totalSelectedCriteriaCount = Object.values(selectionMap).filter((v) => v.checked).length;

  return (
    <div className="rounded-xl border border-slate-200 bg-white overflow-hidden">
      {/* Header + Sub-mode Switcher */}
      <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 bg-slate-50/80 px-4 py-2.5">
        <div className="flex items-center gap-2">
          <ClipboardCheck className="h-4 w-4 text-[#3B5B7E]" />
          <h3 className="text-sm font-semibold text-slate-800">专家打分结果校验</h3>
          <span className="rounded bg-[#ECFDF5] px-1.5 py-0.5 text-[10px] font-medium text-[#2F6F5E]">
            已上线
          </span>
        </div>

        <div className="flex flex-wrap gap-1.5">
          <button
            type="button"
            onClick={() => {
              setSubMode('total');
              setErrorMsg('');
            }}
            className={`inline-flex items-center gap-1.5 rounded-md px-3 py-1.5 text-xs font-medium transition-colors ${
              subMode === 'total' ? 'bg-[#3B5B7E] text-white' : 'bg-white border border-slate-200 text-slate-600 hover:bg-slate-100'
            }`}
          >
            <span>1. 专家总分校验</span>
            {totalResult && (
              <span
                className={`rounded-full px-1.5 text-[10px] tabular-nums ${
                  totalResult.counts.totalIssues > 0
                    ? 'bg-[#A94442] text-white'
                    : subMode === 'total'
                      ? 'bg-white/20 text-white'
                      : 'bg-[#ECFDF5] text-[#2F6F5E]'
                }`}
              >
                {totalResult.counts.totalIssues}
              </span>
            )}
          </button>

          <button
            type="button"
            onClick={() => {
              setSubMode('objective');
              setErrorMsg('');
            }}
            className={`inline-flex items-center gap-1.5 rounded-md px-3 py-1.5 text-xs font-medium transition-colors ${
              subMode === 'objective' ? 'bg-[#3B5B7E] text-white' : 'bg-white border border-slate-200 text-slate-600 hover:bg-slate-100'
            }`}
          >
            <span>2. 专家客观分校验</span>
            {objResult && (
              <span
                className={`rounded-full px-1.5 text-[10px] tabular-nums ${
                  objResult.counts.totalIssues > 0
                    ? 'bg-[#A94442] text-white'
                    : subMode === 'objective'
                      ? 'bg-white/20 text-white'
                      : 'bg-[#ECFDF5] text-[#2F6F5E]'
                }`}
              >
                {objResult.counts.totalIssues}
              </span>
            )}
          </button>
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

      {/* ==================== SUB-MODE 1: 专家总分校验 (工具 8) ==================== */}
      {subMode === 'total' && (
        <div className="p-4 space-y-4">
          <div className="rounded-lg border border-slate-200 bg-slate-50/70 p-3.5 space-y-3">
            <div className="flex items-start gap-2 text-[11px] text-slate-600">
              <Info className="h-3.5 w-3.5 shrink-0 text-[#3B5B7E] mt-0.5" />
              <span>
                上传一份或多份《商务/技术评分统计表（专家维度）》（支持多选 <code>.xlsx/.xls</code> 或 <code>.zip</code> 压缩包）。核验每标包每位投标人的专家打分是否存在<strong>超过半数专家同分</strong>或<strong>任意两名专家分差严格大于阈值</strong>（默认 8 分）。
              </span>
            </div>

            <div className="flex flex-wrap items-end gap-3">
              <div>
                <label className="block text-[11px] font-medium text-slate-600 mb-1">
                  专家人数（留空自动识别）
                </label>
                <input
                  type="number"
                  min={1}
                  value={expertNumInput}
                  onChange={(e) => setExpertNumInput(e.target.value)}
                  placeholder="自动识别"
                  className="w-32 rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs tabular-nums focus:border-[#3B5B7E] focus:outline-none"
                />
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

              <div>
                <label className="block text-[11px] font-medium text-slate-600 mb-1">
                  工作簿口令（加密表可选）
                </label>
                <input
                  type="password"
                  value={passwordInput}
                  onChange={(e) => setPasswordInput(e.target.value)}
                  placeholder="无加密请留空"
                  className="w-36 rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>

              <div className="flex flex-wrap items-center gap-2">
                <label className="inline-flex cursor-pointer items-center gap-1.5 rounded-md bg-[#3B5B7E] px-3 py-1.5 text-xs font-medium text-white hover:bg-[#2E4866] transition-colors">
                  <Upload className="h-3.5 w-3.5" />
                  <span>{loading ? '核验中...' : state.total?.files.length ? '重新上传并校验' : '上传总分表并校验'}</span>
                  <input
                    type="file"
                    multiple
                    accept=".xlsx,.xls,.zip"
                    onChange={handleTotalUpload}
                    className="hidden"
                    disabled={loading}
                  />
                </label>

                {state.total && state.total.files.length > 0 && (
                  <>
                    <button
                      type="button"
                      onClick={handleTotalRerun}
                      disabled={loading}
                      className="inline-flex items-center gap-1 rounded-md border border-[#3B5B7E]/30 bg-white px-2.5 py-1.5 text-xs font-medium text-[#3B5B7E] hover:bg-[#EBF2F7]/60 disabled:opacity-50"
                    >
                      <RefreshCw className="h-3.5 w-3.5" />
                      <span>按新参数重算</span>
                    </button>

                    {totalResult && (
                      <a
                        href={`/api/evaluation/batch/${batchId}/score-check/total/export`}
                        className="inline-flex items-center gap-1 rounded-md border border-[#2F6F5E]/30 bg-[#ECFDF5]/70 px-2.5 py-1.5 text-xs font-medium text-[#2F6F5E] hover:bg-[#D1FAE5]"
                      >
                        <Download className="h-3.5 w-3.5" />
                        <span>导出专家总分核验 (.xlsx)</span>
                      </a>
                    )}

                    <button
                      type="button"
                      onClick={handleTotalClear}
                      disabled={loading}
                      className="inline-flex items-center gap-1 rounded-md border border-[#A94442]/25 bg-[#FEF2F2]/70 px-2.5 py-1.5 text-xs font-medium text-[#A94442] hover:bg-[#FEE2E2] disabled:opacity-50"
                    >
                      <Trash2 className="h-3.5 w-3.5" />
                      <span>清空</span>
                    </button>
                  </>
                )}
              </div>
            </div>

            {state.total && state.total.files.length > 0 && (
              <div className="flex flex-wrap items-center gap-2 border-t border-slate-200/70 pt-2 text-[11px] text-slate-500 tabular-nums">
                <FileSpreadsheet className="h-3.5 w-3.5 text-[#3B5B7E]" />
                <span>
                  已载入 <strong className="text-slate-700">{state.total.files.length}</strong> 份工作簿
                  {totalResult && (
                    <>
                      {' · '}共 <strong className="text-slate-700">{totalResult.sheetCount}</strong> 个标包页
                      {' · '}核验 <strong className="text-slate-700">{totalResult.checkedCount}</strong> 条投标人总分
                      {totalResult.expertNum ? ` · 识别专家 ${totalResult.expertNum} 位` : ''}
                      {' · '}阈值 {totalResult.maxScoreDiff} 分
                    </>
                  )}
                </span>
              </div>
            )}
          </div>

          {totalResult && (
            <div className="space-y-3">
              {/* Filter & Search */}
              <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
                <div className="flex flex-wrap gap-1.5">
                  {([
                    { key: 'all', label: `全部异常 (${totalResult.counts.totalIssues})` },
                    { key: 'majority', label: `超过半数同分 (${totalResult.counts.majoritySame})` },
                    { key: 'diff', label: `分差超过阈值 (${totalResult.counts.diffExceeded})` },
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
                    placeholder="搜索分标、包号、投标人或专家..."
                    className="w-full rounded-lg border border-slate-200 py-1.5 pl-8 pr-3 text-xs focus:border-[#3B5B7E] focus:outline-none"
                  />
                </div>
              </div>

              <div className="overflow-hidden rounded-lg border border-slate-200">
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse text-left text-xs">
                    <thead>
                      <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">分标 / 包名</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">投标人</th>
                        <th className="px-2.5 py-2 font-medium min-w-[200px]">超过半数同分</th>
                        <th className="px-2.5 py-2 font-medium min-w-[260px]">分差超过阈值</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">来源文件 / 位置</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {filteredTotalItems.length === 0 ? (
                        <tr>
                          <td colSpan={5} className="py-8 text-center text-xs text-[#2F6F5E]">
                            <span className="inline-flex items-center gap-1.5">
                              <CheckCircle2 className="h-4 w-4" />
                              {totalResult.counts.totalIssues === 0
                                ? `已核验 ${totalResult.checkedCount} 条投标人总分记录，未发现超半数同分或分差超阈值异常`
                                : '当前筛选条件下无匹配记录'}
                            </span>
                          </td>
                        </tr>
                      ) : (
                        paginate(filteredTotalItems, totalPager.page, totalPager.pageSize).map((item, idx) => (
                          <tr key={idx} className="hover:bg-slate-50/70 align-top">
                            <td className="px-2.5 py-2 whitespace-nowrap">
                              <div className="font-medium text-slate-800">{item.section}</div>
                              <div className="text-[11px] text-slate-400">{item.package_name}</div>
                            </td>
                            <td className="px-2.5 py-2 font-medium text-slate-800">{item.supplier}</td>
                            <td className="px-2.5 py-2 text-[#9A6B26] whitespace-pre-line">
                              {item.majority_same || <span className="text-slate-300">—</span>}
                            </td>
                            <td className="px-2.5 py-2 text-[#A94442] whitespace-pre-line leading-relaxed">
                              {item.diff_exceeded || <span className="text-slate-300">—</span>}
                            </td>
                            <td className="px-2.5 py-2 text-[11px] text-slate-500">
                              <div className="max-w-[220px] truncate" title={item.source_file}>
                                {item.source_file}
                              </div>
                              <div className="font-mono text-[10px] text-slate-400">{item.supplier_location}</div>
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
        </div>
      )}

      {/* ==================== SUB-MODE 2: 专家客观分校验 (工具 9) ==================== */}
      {subMode === 'objective' && (
        <div className="p-4 space-y-4">
          {/* Step 1: Upload Expert Detail Sheets / ZIP */}
          <div className="rounded-lg border border-slate-200 bg-slate-50/70 p-3.5 space-y-3">
            <div className="flex items-start gap-2 text-[11px] text-slate-600">
              <Info className="h-3.5 w-3.5 shrink-0 text-[#3B5B7E] mt-0.5" />
              <span>
                <strong>步骤 1：</strong>上传同一评审组多位专家的《商务/技术评审打分表》（文件名格式须为 <code>评分表_专家编号_专家姓名.xlsx</code>，支持直接上传导出的 <strong>.zip 压缩包</strong> 或多选 <strong>.xlsx/.xls</strong>）。系统自动识别所有评分办法后，勾选需要核验的客观评分细则并执行校验。
              </span>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <label className="inline-flex cursor-pointer items-center gap-1.5 rounded-md bg-[#3B5B7E] px-3 py-1.5 text-xs font-medium text-white hover:bg-[#2E4866] transition-colors">
                <Upload className="h-3.5 w-3.5" />
                <span>{loading ? '读取中...' : state.objective?.files.length ? '重新上传评分明细表 / ZIP' : '上传评分明细表 / ZIP'}</span>
                <input
                  type="file"
                  multiple
                  accept=".xlsx,.xls,.zip"
                  onChange={handleObjectiveUpload}
                  className="hidden"
                  disabled={loading}
                />
              </label>

              {state.objective && state.objective.files.length > 0 && (
                <>
                  {objResult && (
                    <a
                      href={`/api/evaluation/batch/${batchId}/score-check/objective/export`}
                      className="inline-flex items-center gap-1 rounded-md border border-[#2F6F5E]/30 bg-[#ECFDF5]/70 px-2.5 py-1.5 text-xs font-medium text-[#2F6F5E] hover:bg-[#D1FAE5]"
                    >
                      <Download className="h-3.5 w-3.5" />
                      <span>导出专家客观分核验 (.xlsx)</span>
                    </a>
                  )}

                  <button
                    type="button"
                    onClick={handleObjectiveClear}
                    disabled={loading}
                    className="inline-flex items-center gap-1 rounded-md border border-[#A94442]/25 bg-[#FEF2F2]/70 px-2.5 py-1.5 text-xs font-medium text-[#A94442] hover:bg-[#FEE2E2] disabled:opacity-50"
                  >
                    <Trash2 className="h-3.5 w-3.5" />
                    <span>清空</span>
                  </button>
                </>
              )}
            </div>

            {state.objective && state.objective.files.length > 0 && (
              <div className="flex flex-wrap items-center gap-1.5 border-t border-slate-200/70 pt-2 text-[11px] text-slate-500">
                <FileSpreadsheet className="h-3.5 w-3.5 text-[#3B5B7E] shrink-0" />
                <span>已载入 <strong className="text-slate-700">{state.objective.files.length}</strong> 位专家评分表：</span>
                {state.objective.files.slice(0, 8).map((f) => (
                  <span key={f.originalName} className="rounded bg-white border border-slate-200 px-1.5 py-0.5 text-[10px] text-slate-600">
                    {f.originalName}
                  </span>
                ))}
                {state.objective.files.length > 8 && (
                  <span className="text-[10px] text-slate-400">等共 {state.objective.files.length} 份</span>
                )}
              </div>
            )}
          </div>

          {/* Step 2: Multi-Scoring-Method Criteria Selection (ported from ObjectiveSelectionDialog.java) */}
          {objMethods.length > 0 && (
            <div className="rounded-lg border border-slate-200 bg-white overflow-hidden">
              <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 bg-slate-50/70 px-3.5 py-2">
                <div className="flex flex-wrap items-center gap-1.5">
                  <span className="text-xs font-semibold text-slate-700 mr-1">
                    步骤 2：勾选客观评分细则（已选 {totalSelectedCriteriaCount} 项）
                  </span>
                  {objMethods.map((m, idx) => {
                    const checkedInMethod = m.criteria.filter(
                      (c) => selectionMap[`${m.id}::${c.name}`]?.checked
                    ).length;
                    return (
                      <button
                        key={m.id}
                        type="button"
                        onClick={() => setActiveMethodIdx(idx)}
                        className={`rounded-md px-2.5 py-1 text-[11px] font-medium tabular-nums transition-colors ${
                          activeMethodIdx === idx
                            ? 'bg-[#3B5B7E] text-white'
                            : 'bg-white border border-slate-200 text-slate-600 hover:bg-slate-100'
                        }`}
                      >
                        评分办法{idx + 1}（{m.criteria.length}项 / {m.packages.length}包
                        {checkedInMethod > 0 ? ` · 已选${checkedInMethod}` : ''}）
                      </button>
                    );
                  })}
                </div>

                <div className="flex flex-wrap items-center gap-2">
                  <input
                    value={sectionOverride}
                    onChange={(e) => setSectionOverride(e.target.value)}
                    placeholder="分标名称覆盖（通常留空自动提取）"
                    className="w-52 rounded-md border border-slate-200 bg-white px-2.5 py-1 text-[11px] focus:border-[#3B5B7E] focus:outline-none"
                  />
                  <button
                    type="button"
                    onClick={handleObjectiveRun}
                    disabled={loading || totalSelectedCriteriaCount === 0}
                    className="inline-flex items-center gap-1 rounded-md bg-[#2F6F5E] px-3 py-1.5 text-xs font-medium text-white hover:bg-[#24594B] disabled:opacity-50 transition-colors"
                  >
                    <Play className="h-3.5 w-3.5" />
                    <span>{loading ? '校验中...' : '开始客观分校验'}</span>
                  </button>
                </div>
              </div>

              {(() => {
                const currentMethod = objMethods[Math.min(activeMethodIdx, objMethods.length - 1)];
                if (!currentMethod) return null;
                return (
                  <div className="p-3 space-y-2.5">
                    <div className="flex flex-wrap items-center justify-between gap-2 text-[11px]">
                      <div className="text-slate-500">
                        <strong className="text-slate-700">适用标包范围（{currentMethod.packages.length} 包）：</strong>
                        <span>{currentMethod.packages.join('、')}</span>
                      </div>
                      <div className="flex items-center gap-1.5 shrink-0">
                        <button
                          type="button"
                          onClick={() =>
                            setMethodAllChecked(
                              currentMethod.id,
                              currentMethod.criteria.map((c) => c.name),
                              true
                            )
                          }
                          className="rounded border border-slate-200 bg-slate-50 px-2 py-0.5 text-[11px] text-slate-600 hover:bg-slate-100"
                        >
                          本页全选
                        </button>
                        <button
                          type="button"
                          onClick={() =>
                            setMethodAllChecked(
                              currentMethod.id,
                              currentMethod.criteria.map((c) => c.name),
                              false
                            )
                          }
                          className="rounded border border-slate-200 bg-slate-50 px-2 py-0.5 text-[11px] text-slate-600 hover:bg-slate-100"
                        >
                          本页全不选
                        </button>
                      </div>
                    </div>

                    <div className="max-h-72 overflow-y-auto rounded-lg border border-slate-200">
                      <table className="w-full border-collapse text-left text-xs">
                        <thead className="sticky top-0 bg-slate-50 text-slate-500 border-b border-slate-200">
                          <tr>
                            <th className="px-2.5 py-1.5 w-12 text-center font-medium">核验</th>
                            <th className="px-2.5 py-1.5 w-52 font-medium">项目名称</th>
                            <th className="px-2.5 py-1.5 font-medium">评分细则（原表完整名称）</th>
                            <th className="px-2.5 py-1.5 w-36 font-medium whitespace-nowrap">分值步长（可空）</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {currentMethod.criteria.map((c) => {
                            const key = `${currentMethod.id}::${c.name}`;
                            const entry = selectionMap[key];
                            const checked = Boolean(entry?.checked);
                            return (
                              <tr
                                key={c.name}
                                className={checked ? 'bg-[#EBF2F7]/40' : 'hover:bg-slate-50/70'}
                              >
                                <td className="px-2.5 py-1.5 text-center">
                                  <input
                                    type="checkbox"
                                    checked={checked}
                                    onChange={() => toggleCriterion(currentMethod.id, c.name)}
                                    className="rounded border-slate-300 cursor-pointer"
                                  />
                                </td>
                                <td
                                  className="px-2.5 py-1.5 text-slate-700 font-medium cursor-pointer"
                                  onClick={() => toggleCriterion(currentMethod.id, c.name)}
                                >
                                  {c.project}
                                </td>
                                <td
                                  className="px-2.5 py-1.5 text-slate-600 leading-relaxed cursor-pointer"
                                  onClick={() => toggleCriterion(currentMethod.id, c.name)}
                                >
                                  {c.name}
                                </td>
                                <td className="px-2.5 py-1.5">
                                  <input
                                    type="number"
                                    min={0}
                                    step="0.5"
                                    value={entry?.step || ''}
                                    onChange={(e) => setCriterionStep(currentMethod.id, c.name, e.target.value)}
                                    placeholder="如 1 或 0.5"
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

          {/* Step 3: Objective Score Audit Results */}
          {objResult && (
            <div className="space-y-3">
              {objResult.warnings.length > 0 && (
                <div className="rounded-lg border border-[#9A6B26]/30 bg-[#FEF9EC] px-3 py-2 text-xs text-[#9A6B26] space-y-1">
                  {objResult.warnings.map((w, i) => (
                    <div key={i}>• {w}</div>
                  ))}
                </div>
              )}

              <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
                <div className="flex flex-wrap gap-1.5">
                  {([
                    { key: 'all', label: `全部问题 (${objResult.counts.totalIssues})` },
                    { key: '分数不一致', label: `分数不一致 (${objResult.counts.scoreMismatch})` },
                    { key: '分值步长不符', label: `分值步长不符 (${objResult.counts.stepMismatch})` },
                    { key: '专家记录缺失', label: `专家记录缺失 (${objResult.counts.expertMissing})` },
                    { key: '专家数量不足', label: `专家数量不足 (${objResult.counts.expertInsufficient})` },
                    { key: '细则未找到', label: `细则未找到 (${objResult.counts.criterionMissing})` },
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
                    placeholder="搜索分标、包名、投标人或细则..."
                    className="w-full rounded-lg border border-slate-200 py-1.5 pl-8 pr-3 text-xs focus:border-[#3B5B7E] focus:outline-none"
                  />
                </div>
              </div>

              <div className="overflow-hidden rounded-lg border border-slate-200">
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse text-left text-xs">
                    <thead>
                      <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">分标 / 包名</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">投标人</th>
                        <th className="px-2.5 py-2 font-medium min-w-[200px]">客观评分细则</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">问题类型</th>
                        <th className="px-2.5 py-2 font-medium min-w-[220px]">各专家打分说明</th>
                        <th className="px-2.5 py-2 font-medium min-w-[200px]">来源位置</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {filteredObjItems.length === 0 ? (
                        <tr>
                          <td colSpan={6} className="py-8 text-center text-xs text-[#2F6F5E]">
                            <span className="inline-flex items-center gap-1.5">
                              <CheckCircle2 className="h-4 w-4" />
                              {objResult.counts.totalIssues === 0
                                ? `已核验 ${objResult.checkedCount} 组客观评分项，各专家打分完全一致且符合步长要求`
                                : '当前筛选条件下无匹配记录'}
                            </span>
                          </td>
                        </tr>
                      ) : (
                        paginate(filteredObjItems, objPager.page, objPager.pageSize).map((item, idx) => (
                          <tr key={idx} className="hover:bg-slate-50/70 align-top">
                            <td className="px-2.5 py-2 whitespace-nowrap">
                              <div className="font-medium text-slate-800">{item.section}</div>
                              <div className="text-[11px] text-slate-400">{item.package_name}</div>
                            </td>
                            <td className="px-2.5 py-2 font-medium text-slate-800">{item.supplier || '—'}</td>
                            <td className="px-2.5 py-2 text-slate-600 leading-relaxed max-w-[280px]">{item.criterion}</td>
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
  );
}
