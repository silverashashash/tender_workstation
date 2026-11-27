import { useCallback, useEffect, useState } from 'react';
import {
  AlertTriangle,
  Copy,
  Check,
  Download,
  FileSpreadsheet,
  FileText,
  Info,
  RefreshCw,
  Trash2,
  Upload,
} from 'lucide-react';
import { api } from '../api/client';
import { paginate, PaginationBar, usePagination } from './PaginationBar';
import type { InitialReportState } from '../types';

const DEFAULT_SUFFIXES = '有限责任公司,有限公司,公司,集团,股份有限责任公司,园区,股份有限公司';

interface Props {
  batchId: string;
}

export function InitialReportPanel({ batchId }: Props) {
  const [state, setState] = useState<InitialReportState | null>(null);
  const [loading, setLoading] = useState(false);
  const [errorMsg, setErrorMsg] = useState('');
  const [copied, setCopied] = useState(false);

  const [department, setDepartment] = useState('');
  const [clarification, setClarification] = useState('');
  const [simplifyNames, setSimplifyNames] = useState(true);
  const [companySuffixes, setCompanySuffixes] = useState(DEFAULT_SUFFIXES);
  const [previewTab, setPreviewTab] = useState<'doc' | 'grouped'>('doc');

  const groupPager = usePagination(`${batchId}-ir-${state?.updated_at || ''}`);

  const syncFromState = useCallback((data: InitialReportState) => {
    setState(data);
    setDepartment(data.config?.department || '');
    setClarification(data.config?.clarification || '');
    setSimplifyNames(data.config?.simplifyNames !== undefined ? Boolean(data.config.simplifyNames) : true);
    setCompanySuffixes(
      data.config?.companySuffixes !== undefined ? data.config.companySuffixes : DEFAULT_SUFFIXES
    );
  }, []);

  const loadState = useCallback(async () => {
    if (!batchId) return;
    try {
      const data = await api.getInitialReportState(batchId);
      syncFromState(data);
    } catch {
      // ignore if offline
    }
  }, [batchId, syncFromState]);

  useEffect(() => {
    loadState();
    setErrorMsg('');
  }, [loadState]);

  async function handleUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;
    setLoading(true);
    setErrorMsg('');
    try {
      const res = await api.uploadInitialReportTemplate(batchId, file, {
        department,
        clarification,
        simplifyNames,
        companySuffixes,
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
      const res = await api.regenerateInitialReport(batchId, {
        department,
        clarification,
        simplifyNames,
        companySuffixes,
      });
      syncFromState(res);
    } catch (err) {
      setErrorMsg((err as Error).message);
    } finally {
      setLoading(false);
    }
  }

  async function handleClear() {
    if (!window.confirm('确认清空当前批次已上传的初评情况汇报模板与生成的汇报单吗？')) return;
    setLoading(true);
    setErrorMsg('');
    try {
      await api.clearInitialReport(batchId);
      syncFromState({
        file: null,
        config: {
          department: '',
          clarification: '',
          simplifyNames: true,
          companySuffixes: DEFAULT_SUFFIXES,
        },
        result: null,
        updated_at: '',
      });
    } catch (err) {
      setErrorMsg((err as Error).message);
    } finally {
      setLoading(false);
    }
  }

  async function handleCopyText() {
    if (!state?.result) return;
    const text = ['初评汇报单', ...state.result.paragraphs].join('\n\n');
    try {
      await navigator.clipboard.writeText(text);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch {
      // ignore clipboard errors
    }
  }

  const result = state?.result || null;

  return (
    <div className="rounded-xl border border-slate-200 bg-white overflow-hidden">
      {/* Header */}
      <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 bg-slate-50/80 px-4 py-2.5">
        <div className="flex items-center gap-2">
          <FileText className="h-4 w-4 text-[#3B5B7E]" />
          <h3 className="text-sm font-semibold text-slate-800">初评汇报生成</h3>
          <span className="rounded bg-[#ECFDF5] px-1.5 py-0.5 text-[10px] font-medium text-[#2F6F5E]">
            已上线
          </span>
        </div>
        {result && (
          <div className="flex flex-wrap items-center gap-2 text-[11px] text-slate-500 tabular-nums">
            <span>
              <strong className="text-slate-700">{result.dept}</strong>（组长：{result.leader}、{result.members}）
            </span>
            <span>·</span>
            <span>
              {result.sectionCount} 标段 / {result.packCount} 包 / {result.bidderCount} 家 / {result.documentCount} 份
            </span>
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
        {/* Upload & Config Card */}
        <div className="rounded-lg border border-slate-200 bg-slate-50/70 p-3.5 space-y-3">
          <div className="flex items-start gap-2 text-[11px] text-slate-600">
            <Info className="h-3.5 w-3.5 shrink-0 text-[#3B5B7E] mt-0.5" />
            <span>
              上传《初评情况汇报模板.xlsx》：自动识别<strong>初评情况汇报附件</strong>（否决明细，跳过封面页）、<strong>封面二</strong>（基本情况统计）与<strong>封面一</strong>（组别及组长信息），自动按标段/原因/企业归并连续包号区间并生成标准格式 Word <strong>初评汇报单 (.docx)</strong>。
            </span>
          </div>

          <div className="grid grid-cols-1 gap-3 md:grid-cols-2 lg:grid-cols-4">
            <div>
              <label className="block text-[11px] font-medium text-slate-600 mb-1">
                评审组别覆盖（留空自动读取封面一）
              </label>
              <input
                value={department}
                onChange={(e) => setDepartment(e.target.value)}
                placeholder="如：商务二组"
                className="w-full rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
              />
            </div>

            <div className="md:col-span-1 lg:col-span-3">
              <label className="block text-[11px] font-medium text-slate-600 mb-1">
                澄清事项（留空则标注「【待核实填写】」提示人工核查）
              </label>
              <input
                value={clarification}
                onChange={(e) => setClarification(e.target.value)}
                placeholder="如：无 / 详见附表 / 本组共发起 2 项商务澄清..."
                className="w-full rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
              />
            </div>
          </div>

          {/* Company name simplification config */}
          <div className="flex flex-col gap-2 border-t border-slate-200/70 pt-2.5 sm:flex-row sm:items-center">
            <label className="inline-flex shrink-0 items-center gap-1.5 text-xs font-medium text-slate-700 cursor-pointer">
              <input
                type="checkbox"
                checked={simplifyNames}
                onChange={(e) => setSimplifyNames(e.target.checked)}
                className="rounded border-slate-300"
              />
              <span>简化公司名称</span>
            </label>
            <input
              value={companySuffixes}
              onChange={(e) => setCompanySuffixes(e.target.value)}
              disabled={!simplifyNames}
              placeholder="逗号分隔的末尾后缀（较长优先连续删除，简称冲突时自动保留全称）"
              className="flex-1 rounded-md border border-slate-200 bg-white px-2.5 py-1 text-[11px] text-slate-600 disabled:bg-slate-100 disabled:text-slate-400 focus:border-[#3B5B7E] focus:outline-none"
            />
            <button
              type="button"
              onClick={() => setCompanySuffixes(DEFAULT_SUFFIXES)}
              disabled={!simplifyNames || companySuffixes === DEFAULT_SUFFIXES}
              className="shrink-0 rounded border border-slate-200 bg-white px-2 py-1 text-[11px] text-slate-500 hover:bg-slate-100 disabled:opacity-40"
            >
              恢复默认后缀
            </button>
          </div>

          {/* Action Buttons */}
          <div className="flex flex-wrap items-center gap-2 pt-1">
            <label className="inline-flex cursor-pointer items-center gap-1.5 rounded-md bg-[#3B5B7E] px-3 py-1.5 text-xs font-medium text-white hover:bg-[#2E4866] transition-colors">
              <Upload className="h-3.5 w-3.5" />
              <span>{loading ? '解析中...' : state?.file ? '重新上传初评汇报模板' : '上传初评汇报模板 (.xlsx)'}</span>
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
                  <>
                    <a
                      href={`/api/evaluation/batch/${batchId}/initial-report/export`}
                      className="inline-flex items-center gap-1 rounded-md border border-[#2F6F5E]/30 bg-[#ECFDF5]/70 px-3 py-1.5 text-xs font-medium text-[#2F6F5E] hover:bg-[#D1FAE5]"
                    >
                      <Download className="h-3.5 w-3.5" />
                      <span>导出初评汇报单 (.docx)</span>
                    </a>

                    <button
                      type="button"
                      onClick={handleCopyText}
                      className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-100"
                    >
                      {copied ? <Check className="h-3.5 w-3.5 text-[#2F6F5E]" /> : <Copy className="h-3.5 w-3.5" />}
                      <span>{copied ? '已复制全文' : '复制汇报全文'}</span>
                    </button>
                  </>
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
                已载入 <strong className="text-slate-700">{state.file.originalName}</strong> · 读取附件表「
                {result.sheetNames.detailSheet}」（共 {result.rejectionRowCount} 条否决明细 → 归并为{' '}
                {result.groupedRejections.length} 条汇报段落）
              </span>
            </div>
          )}
        </div>

        {/* Warnings */}
        {result && result.warnings.length > 0 && (
          <div className="rounded-lg border border-[#9A6B26]/30 bg-[#FEF9EC] px-3 py-2 text-xs text-[#9A6B26] space-y-1">
            {result.warnings.map((w, i) => (
              <div key={i} className="flex items-start gap-1.5">
                <AlertTriangle className="h-3.5 w-3.5 shrink-0 mt-0.5" />
                <span>{w}</span>
              </div>
            ))}
          </div>
        )}

        {/* Result Preview */}
        {result && (
          <div className="space-y-3">
            <div className="flex items-center gap-1.5">
              <button
                type="button"
                onClick={() => setPreviewTab('doc')}
                className={`rounded-md px-2.5 py-1 text-xs font-medium transition-colors ${
                  previewTab === 'doc' ? 'bg-slate-800 text-white' : 'bg-slate-100 text-slate-600 hover:bg-slate-200/70'
                }`}
              >
                汇报单正文预览（{result.paragraphs.length} 段）
              </button>
              <button
                type="button"
                onClick={() => setPreviewTab('grouped')}
                className={`rounded-md px-2.5 py-1 text-xs font-medium transition-colors ${
                  previewTab === 'grouped' ? 'bg-slate-800 text-white' : 'bg-slate-100 text-slate-600 hover:bg-slate-200/70'
                }`}
              >
                否决归并明细表（{result.groupedRejections.length} 组）
              </button>
            </div>

            {previewTab === 'doc' && (
              <div className="rounded-lg border border-slate-200 bg-white p-6 space-y-3 max-h-[520px] overflow-y-auto shadow-inner">
                <h4 className="text-center text-base font-bold text-slate-900 tracking-wide">初评汇报单</h4>
                {result.paragraphs.map((p, idx) => (
                  <p key={idx} className="text-[13px] leading-7 text-slate-800 indent-7">
                    {p}
                  </p>
                ))}
              </div>
            )}

            {previewTab === 'grouped' && (
              <div className="overflow-hidden rounded-lg border border-slate-200">
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse text-left text-xs">
                    <thead>
                      <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">分标名称</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">否决原因（问题描述）</th>
                        <th className="px-2.5 py-2 font-medium whitespace-nowrap">涉及家数</th>
                        <th className="px-2.5 py-2 font-medium min-w-[300px]">归并后企业与包号区间</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {result.groupedRejections.length === 0 ? (
                        <tr>
                          <td colSpan={4} className="py-6 text-center text-xs text-slate-400">
                            输入附件明细中未列出否决事项
                          </td>
                        </tr>
                      ) : (
                        paginate(result.groupedRejections, groupPager.page, groupPager.pageSize).map((g, idx) => (
                          <tr key={idx} className="hover:bg-slate-50/70 align-top">
                            <td className="px-2.5 py-2 font-medium text-slate-800 whitespace-nowrap">{g.section}</td>
                            <td className="px-2.5 py-2 text-[#A94442] whitespace-nowrap">{g.reason}</td>
                            <td className="px-2.5 py-2 tabular-nums text-slate-600">{g.items.length} 家</td>
                            <td className="px-2.5 py-2 text-slate-700 leading-relaxed">
                              <div className="flex flex-wrap gap-1">
                                {g.items.map((it) => (
                                  <span
                                    key={it.company}
                                    title={it.company}
                                    className="rounded border border-slate-200 bg-slate-50 px-1.5 py-0.5 text-[11px]"
                                  >
                                    {it.displayName}（<strong>{it.packs}</strong>）
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
                {result.groupedRejections.length > 0 && (
                  <PaginationBar
                    total={result.groupedRejections.length}
                    page={groupPager.page}
                    pageSize={groupPager.pageSize}
                    onPageChange={groupPager.setPage}
                    onPageSizeChange={groupPager.setPageSize}
                  />
                )}
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
