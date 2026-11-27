import { useCallback, useEffect, useState } from 'react';
import {
  AlertTriangle,
  Download,
  FileSpreadsheet,
  FileText,
  Info,
  RefreshCw,
  Search,
  Trash2,
  Upload,
} from 'lucide-react';
import { api } from '../api/client';
import { paginate, PaginationBar, usePagination } from './PaginationBar';
import type { InitialMinutesState, MinutesProcurementType } from '../types';

interface Props {
  batchId: string;
}

export function InitialMinutesPanel({ batchId }: Props) {
  const [state, setState] = useState<InitialMinutesState | null>(null);
  const [loading, setLoading] = useState(false);
  const [errorMsg, setErrorMsg] = useState('');

  const [procurementType, setProcurementType] = useState<MinutesProcurementType>('非公开');
  const [batchName, setBatchName] = useState('');
  const [meetingTime, setMeetingTime] = useState('');
  const [groupName, setGroupName] = useState('商务组');
  const [clarificationNote, setClarificationNote] = useState('');
  const [reviewTypeFilter, setReviewTypeFilter] = useState<'all' | '商务' | '技术'>('all');
  const [searchText, setSearchText] = useState('');

  const pager = usePagination(`${batchId}-im-${state?.updated_at || ''}-${searchText}`);

  const syncFromState = useCallback((data: InitialMinutesState) => {
    setState(data);
    setProcurementType(data.config?.procurementType === '公开' ? '公开' : '非公开');
    setBatchName(data.config?.batchName || '');
    setMeetingTime(data.config?.meetingTime || '');
    setGroupName(data.config?.groupName || '商务组');
    setClarificationNote(data.config?.clarificationNote || '');
    setReviewTypeFilter(
      data.config?.reviewTypeFilter === '商务' || data.config?.reviewTypeFilter === '技术'
        ? data.config.reviewTypeFilter
        : 'all'
    );
  }, []);

  const loadState = useCallback(async () => {
    if (!batchId) return;
    try {
      const data = await api.getInitialMinutesState(batchId);
      syncFromState(data);
    } catch {
      // ignore if offline
    }
  }, [batchId, syncFromState]);

  useEffect(() => {
    loadState();
    setErrorMsg('');
  }, [loadState]);

  function handleTypeSwitch(nextType: MinutesProcurementType) {
    setProcurementType(nextType);
    // Automatically align default clarification wording if user hasn't customized it
    if (
      !clarificationNote.trim() ||
      clarificationNote.trim() === '本次评标无澄清事项。' ||
      clarificationNote.trim() === '本次评审无澄清事项。'
    ) {
      setClarificationNote(nextType === '公开' ? '本次评标无澄清事项。' : '本次评审无澄清事项。');
    }
  }

  async function handleUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;
    setLoading(true);
    setErrorMsg('');
    try {
      const res = await api.uploadInitialMinutesExcel(batchId, file, {
        procurementType,
        batchName,
        meetingTime,
        groupName,
        clarificationNote,
        reviewTypeFilter,
      });
      syncFromState(res);
    } catch (err) {
      setErrorMsg((err as Error).message);
    } finally {
      setLoading(false);
    }
  }

  async function handleRegenerate() {
    setLoading(true);
    setErrorMsg('');
    try {
      const res = await api.regenerateInitialMinutes(batchId, {
        procurementType,
        batchName,
        meetingTime,
        groupName,
        clarificationNote,
        reviewTypeFilter,
      });
      syncFromState(res);
    } catch (err) {
      setErrorMsg((err as Error).message);
    } finally {
      setLoading(false);
    }
  }

  async function handleClear() {
    if (!window.confirm('确认清空当前批次已上传的初评会议纪要 Excel 与生成记录吗？')) return;
    setLoading(true);
    setErrorMsg('');
    try {
      await api.clearInitialMinutes(batchId);
      await loadState();
    } catch (err) {
      setErrorMsg((err as Error).message);
    } finally {
      setLoading(false);
    }
  }

  const result = state?.result || null;
  const q = searchText.trim().toLowerCase();
  const filteredRows = (result?.rows || []).filter(
    (r) =>
      !q ||
      `${r.lot_name} ${r.package_no} ${r.supplier_name} ${r.reason} ${r.bid_response}`.toLowerCase().includes(q)
  );

  return (
    <div className="rounded-xl border border-slate-200 bg-white overflow-hidden">
      {/* Header */}
      <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 bg-slate-50/80 px-4 py-2.5">
        <div className="flex items-center gap-2">
          <FileText className="h-4 w-4 text-[#3B5B7E]" />
          <h3 className="text-sm font-semibold text-slate-800">初评会议纪要生成</h3>
          <span className="rounded bg-[#ECFDF5] px-1.5 py-0.5 text-[10px] font-medium text-[#2F6F5E]">
            已上线
          </span>
        </div>
        {result && (
          <div className="text-[11px] text-slate-500 tabular-nums">
            模板：<strong className="text-[#3B5B7E]">{result.config.procurementType}招标</strong> ·{' '}
            <strong className="text-slate-700">{result.config.groupName}</strong> · 共{' '}
            <strong className="text-[#A94442]">{result.totalRows}</strong> 条否决明细
          </div>
        )}
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
        {/* Config & Upload Card */}
        <div className="rounded-lg border border-slate-200 bg-slate-50/70 p-3.5 space-y-3">
          <div className="flex items-start gap-2 text-[11px] text-slate-600">
            <Info className="h-3.5 w-3.5 shrink-0 text-[#3B5B7E] mt-0.5" />
            <span>
              支持上传<strong>《初评情况汇报模板》</strong>（单组明细）或<strong>《未进入详评情况统计表》</strong>（自动寻找表头行并匹配分标名称、包号、投标人/响应人、未进入详评原因、投标/应答文件响应），按所选<strong>公开 / 非公开</strong>原版 Word 模板替换标黄字段并生成规范 6 列表格。
            </span>
          </div>

          {/* Procurement Type Toggle + Basic Fields */}
          <div className="grid grid-cols-1 gap-3 md:grid-cols-2 lg:grid-cols-12">
            <div className="lg:col-span-3">
              <label className="block text-[11px] font-medium text-slate-600 mb-1">招标性质（Word 模板）</label>
              <div className="inline-flex w-full rounded-md border border-slate-200 bg-white p-0.5">
                {(['公开', '非公开'] as const).map((t) => (
                  <button
                    key={t}
                    type="button"
                    onClick={() => handleTypeSwitch(t)}
                    className={`flex-1 rounded px-2.5 py-1 text-xs font-medium transition-colors ${
                      procurementType === t ? 'bg-[#3B5B7E] text-white' : 'text-slate-600 hover:bg-slate-100'
                    }`}
                  >
                    {t === '公开' ? '公开招标' : '非公开招标'}
                  </button>
                ))}
              </div>
            </div>

            <div className="lg:col-span-3">
              <label className="block text-[11px] font-medium text-slate-600 mb-1">开会时间</label>
              <input
                value={meetingTime}
                onChange={(e) => setMeetingTime(e.target.value)}
                placeholder="如：2025年12月30日13时00分"
                className="w-full rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
              />
            </div>

            <div className="lg:col-span-2">
              <label className="block text-[11px] font-medium text-slate-600 mb-1">评审组别</label>
              <input
                value={groupName}
                onChange={(e) => setGroupName(e.target.value)}
                placeholder="如：商务一组 / 技术组"
                className="w-full rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
              />
            </div>

            <div className="lg:col-span-4">
              <label className="block text-[11px] font-medium text-slate-600 mb-1">澄清说明（括号内文字）</label>
              <input
                value={clarificationNote}
                onChange={(e) => setClarificationNote(e.target.value)}
                placeholder={procurementType === '公开' ? '本次评标无澄清事项。' : '本次评审无澄清事项。'}
                className="w-full rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
              />
            </div>

            <div className="md:col-span-2 lg:col-span-9">
              <label className="block text-[11px] font-medium text-slate-600 mb-1">批次名称（留空自动读取或使用当前批次名）</label>
              <input
                value={batchName}
                onChange={(e) => setBatchName(e.target.value)}
                placeholder="例如：国网安徽电力2026年第一次服务授权竞争性谈判采购"
                className="w-full rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
              />
            </div>

            <div className="lg:col-span-3">
              <label className="block text-[11px] font-medium text-slate-600 mb-1">
                按商务/技术筛选（全量统计表适用）
              </label>
              <select
                value={reviewTypeFilter}
                onChange={(e) => setReviewTypeFilter(e.target.value as 'all' | '商务' | '技术')}
                className="w-full rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
              >
                <option value="all">全部（不筛选）</option>
                <option value="商务">仅商务类原因</option>
                <option value="技术">仅技术类原因</option>
              </select>
            </div>
          </div>

          {/* Action Bar */}
          <div className="flex flex-wrap items-center gap-2 pt-1">
            <label className="inline-flex cursor-pointer items-center gap-1.5 rounded-md bg-[#3B5B7E] px-3 py-1.5 text-xs font-medium text-white hover:bg-[#2E4866] transition-colors">
              <Upload className="h-3.5 w-3.5" />
              <span>{loading ? '解析中...' : state?.file ? '重新上传否决明细 Excel' : '上传否决明细 Excel (.xlsx)'}</span>
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={handleUpload}
                className="hidden"
                disabled={loading}
              />
            </label>

            {state?.file && (
              <>
                <button
                  type="button"
                  onClick={handleRegenerate}
                  disabled={loading}
                  className="inline-flex items-center gap-1 rounded-md border border-[#3B5B7E]/30 bg-white px-2.5 py-1.5 text-xs font-medium text-[#3B5B7E] hover:bg-[#EBF2F7]/60 disabled:opacity-50"
                >
                  <RefreshCw className="h-3.5 w-3.5" />
                  <span>应用配置并刷新预览</span>
                </button>

                {result && (
                  <a
                    href={`/api/evaluation/batch/${batchId}/initial-minutes/export`}
                    className="inline-flex items-center gap-1 rounded-md border border-[#2F6F5E]/30 bg-[#ECFDF5]/70 px-3 py-1.5 text-xs font-medium text-[#2F6F5E] hover:bg-[#D1FAE5]"
                  >
                    <Download className="h-3.5 w-3.5" />
                    <span>导出初评会议纪要 (.docx)</span>
                  </a>
                )}

                <button
                  type="button"
                  onClick={handleClear}
                  disabled={loading}
                  className="inline-flex items-center gap-1 rounded-md border border-[#A94442]/25 bg-[#FEF2F2]/70 px-2.5 py-1.5 text-xs font-medium text-[#A94442] hover:bg-[#FEE2E2] disabled:opacity-50"
                >
                  <Trash2 className="h-3.5 w-3.5" />
                  <span>清空</span>
                </button>
              </>
            )}
          </div>

          {state?.file && result && (
            <div className="flex flex-wrap items-center gap-2 border-t border-slate-200/70 pt-2 text-[11px] text-slate-500">
              <FileSpreadsheet className="h-3.5 w-3.5 text-[#3B5B7E]" />
              <span>
                已载入 <strong className="text-slate-700">{state.file.originalName}</strong> · 自动定位工作表「
                {result.sheetName}」第 <strong>{result.headerRow}</strong> 行表头
              </span>
            </div>
          )}
        </div>

        {/* Preview: Document Summary + 6-Column Rejection Table */}
        {result && (
          <div className="space-y-4">
            <div className="rounded-lg border border-slate-200 bg-slate-50/40 p-4 space-y-2 text-xs text-slate-700 leading-relaxed">
              <div className="text-center text-sm font-bold text-slate-900">会议纪要（正文预览）</div>
              {result.previewParagraphs.map((p, idx) => (
                <p
                  key={idx}
                  className={
                    p.startsWith('一、') || p.startsWith('二、')
                      ? 'font-semibold text-slate-800 pt-1'
                      : p.startsWith('[否决明细表格')
                        ? 'rounded bg-[#EBF2F7]/70 px-3 py-1.5 text-center text-[11px] font-medium text-[#3B5B7E]'
                        : 'indent-6'
                  }
                >
                  {p}
                </p>
              ))}
            </div>

            <div className="space-y-2">
              <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
                <div className="text-xs font-semibold text-slate-700">
                  纪要附表预览（共 {filteredRows.length} 条 · 将写入 Word 表格）
                </div>
                <div className="relative w-full sm:w-64">
                  <Search className="pointer-events-none absolute left-2.5 top-2 h-3.5 w-3.5 text-slate-400" />
                  <input
                    value={searchText}
                    onChange={(e) => setSearchText(e.target.value)}
                    placeholder="搜索分标、包号、投标人或原因..."
                    className="w-full rounded-lg border border-slate-200 py-1.5 pl-8 pr-3 text-xs focus:border-[#3B5B7E] focus:outline-none"
                  />
                </div>
              </div>

              <div className="overflow-hidden rounded-lg border border-slate-200">
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse text-left text-xs">
                    <thead>
                      <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                        <th className="px-2.5 py-2 font-medium w-12 text-center">序号</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">分标名称</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">包号</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">
                          {result.config.procurementType === '公开' ? '投标人' : '响应人'}
                        </th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">未进入详评原因</th>
                        <th className="px-2.5 py-2 font-medium min-w-[260px]">
                          {result.config.procurementType === '公开' ? '投标文件响应' : '应答文件响应'}
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {filteredRows.length === 0 ? (
                        <tr>
                          <td colSpan={6} className="py-7 text-center text-xs text-slate-400">
                            当前筛选条件下无否决明细记录
                          </td>
                        </tr>
                      ) : (
                        paginate(filteredRows, pager.page, pager.pageSize).map((row) => (
                          <tr key={row.seq} className="hover:bg-slate-50/70 align-top">
                            <td className="px-2.5 py-2 text-center tabular-nums text-slate-400">{row.seq}</td>
                            <td className="px-2.5 py-2 font-medium text-slate-800 whitespace-nowrap">{row.lot_name}</td>
                            <td className="px-2.5 py-2 whitespace-nowrap text-slate-600">{row.package_no}</td>
                            <td className="px-2.5 py-2 font-medium text-slate-800">{row.supplier_name}</td>
                            <td className="px-2.5 py-2 whitespace-nowrap">
                              <span className="rounded bg-[#FEF2F2] px-1.5 py-0.5 text-[11px] font-medium text-[#A94442]">
                                {row.reason}
                              </span>
                            </td>
                            <td className="px-2.5 py-2 text-slate-600 leading-relaxed">{row.bid_response || '—'}</td>
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>
                {filteredRows.length > 0 && (
                  <PaginationBar
                    total={filteredRows.length}
                    page={pager.page}
                    pageSize={pager.pageSize}
                    onPageChange={pager.setPage}
                    onPageSizeChange={pager.setPageSize}
                  />
                )}
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
