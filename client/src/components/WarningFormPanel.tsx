import { useCallback, useEffect, useState } from 'react';
import {
  AlertTriangle,
  CheckCircle2,
  Download,
  FileArchive,
  FileSpreadsheet,
  Info,
  Play,
  RefreshCw,
  ShieldAlert,
  Trash2,
  Upload,
} from 'lucide-react';
import { api } from '../api/client';
import { paginate, PaginationBar, usePagination } from './PaginationBar';
import type { WarningFormState } from '../types';

interface Props {
  batchId: string;
}

export function WarningFormPanel({ batchId }: Props) {
  const [state, setState] = useState<WarningFormState | null>(null);
  const [pendingFiles, setPendingFiles] = useState<{ risk?: File; groups?: File; bids?: File }>({});
  const [dateInput, setDateInput] = useState('');
  const [loading, setLoading] = useState(false);
  const [errorMsg, setErrorMsg] = useState('');

  const pager = usePagination(`${batchId}-wf-${state?.updated_at || ''}`);

  const loadState = useCallback(async () => {
    if (!batchId) return;
    try {
      const data = await api.getWarningFormState(batchId);
      setState(data);
      setDateInput(data.config?.date || '');
    } catch {
      // ignore if offline
    }
  }, [batchId]);

  useEffect(() => {
    loadState();
    setPendingFiles({});
    setErrorMsg('');
  }, [loadState]);

  function handleSelectFile(key: 'risk' | 'groups' | 'bids', e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;
    setPendingFiles((prev) => ({ ...prev, [key]: file }));
    setErrorMsg('');
  }

  const hasRisk = Boolean(pendingFiles.risk || state?.files.risk);
  const hasGroups = Boolean(pendingFiles.groups || state?.files.groups);
  const hasBids = Boolean(pendingFiles.bids || state?.files.bids);
  const readyToRun = hasRisk && hasGroups && hasBids;

  async function handleRun() {
    if (!readyToRun) {
      setErrorMsg('请先完整选择风险提示表、标包分组表与开标项目信息表');
      return;
    }
    setLoading(true);
    setErrorMsg('');
    try {
      const hasNewUpload = Boolean(pendingFiles.risk || pendingFiles.groups || pendingFiles.bids);
      const res = hasNewUpload
        ? await api.runWarningForm(batchId, {
            risk: pendingFiles.risk,
            groups: pendingFiles.groups,
            bids: pendingFiles.bids,
            date: dateInput.trim(),
          })
        : await api.rerunWarningForm(batchId, dateInput.trim());
      setState(res);
      setPendingFiles({});
    } catch (err) {
      setErrorMsg((err as Error).message);
    } finally {
      setLoading(false);
    }
  }

  async function handleClear() {
    if (!window.confirm('确认清空当前批次已上传的预警事项确认单输入表与生成结果吗？')) return;
    setLoading(true);
    setErrorMsg('');
    try {
      await api.clearWarningForm(batchId);
      setState({
        files: { risk: null, groups: null, bids: null },
        config: { date: '' },
        result: null,
        updated_at: '',
      });
      setPendingFiles({});
      setDateInput('');
    } catch (err) {
      setErrorMsg((err as Error).message);
    } finally {
      setLoading(false);
    }
  }

  const result = state?.result || null;

  const slots: Array<{ key: 'risk' | 'groups' | 'bids'; order: number; label: string; hint: string }> = [
    {
      key: 'risk',
      order: 1,
      label: '围标串标风险提示表',
      hint: '支持单行或双行纵向合并表头（含注册地址/企业联系方式/邮箱及是否相同列）',
    },
    {
      key: 'groups',
      order: 2,
      label: '标包分组表',
      hint: '需含分标名称、分包名称、技术组（技术评审组）、商务组（商务评审组）',
    },
    {
      key: 'bids',
      order: 3,
      label: '开标项目信息表',
      hint: '需含采购项目名称、采购项目编号、分标名称、分标编号、分包名称、投标人名称',
    },
  ];

  return (
    <div className="rounded-xl border border-slate-200 bg-white overflow-hidden">
      {/* Header */}
      <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 bg-slate-50/80 px-4 py-2.5">
        <div className="flex items-center gap-2">
          <ShieldAlert className="h-4 w-4 text-[#9A6B26]" />
          <h3 className="text-sm font-semibold text-slate-800">预警事项确认单生成</h3>
          <span className="rounded bg-[#ECFDF5] px-1.5 py-0.5 text-[10px] font-medium text-[#2F6F5E]">
            已上线
          </span>
        </div>
        {result && (
          <div className="text-[11px] text-slate-500 tabular-nums">
            已生成 <strong className="text-[#3B5B7E]">{result.formCount}</strong> 份确认单 · 落款日期{' '}
            <strong className="text-slate-700">{result.date}</strong>
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
        <div className="rounded-lg border border-slate-200 bg-slate-50/70 p-3.5 space-y-3">
          <div className="flex items-start gap-2 text-[11px] text-slate-600">
            <Info className="h-3.5 w-3.5 shrink-0 text-[#3B5B7E] mt-0.5" />
            <span>
              严格按原工具逻辑：同标包内按三项预警标记（注册地址、企业联系方式、邮箱）组合及具体值完全一致匹配 ≥2 家不同投标人，使用原版《预警事项确认单.docx》模板按<strong>技术组 / 商务组</strong>分别生成 Word 并打包 ZIP（跨包仅在企业名单、分标、组别及原因完全一致时归并包号）。
            </span>
          </div>

          {/* 3-File Input Slots */}
          <div className="grid grid-cols-1 gap-3 md:grid-cols-3">
            {slots.map((slot) => {
              const pending = pendingFiles[slot.key];
              const persisted = state?.files[slot.key] || null;
              const activeName = pending?.name || persisted?.originalName || '';
              return (
                <div
                  key={slot.key}
                  className={`flex flex-col justify-between rounded-lg border p-3 ${
                    activeName ? 'border-[#2F6F5E]/30 bg-white' : 'border-slate-200 bg-white/70'
                  }`}
                >
                  <div>
                    <div className="flex items-center justify-between gap-1">
                      <span className="text-xs font-semibold text-slate-800">
                        {slot.order}. {slot.label}
                      </span>
                      <span className="text-[10px] text-slate-400">
                        {pending ? '待生成' : persisted ? '已上传' : '未选择'}
                      </span>
                    </div>
                    <p className="mt-1 text-[11px] text-slate-500 leading-relaxed">{slot.hint}</p>
                  </div>

                  <div className="mt-2.5 pt-2 border-t border-slate-100 space-y-1.5">
                    {activeName && (
                      <div className="flex items-center gap-1 text-[11px] text-[#2F6F5E]">
                        <FileSpreadsheet className="h-3.5 w-3.5 shrink-0" />
                        <span className="truncate font-medium" title={activeName}>
                          {activeName}
                        </span>
                      </div>
                    )}
                    <label className="flex cursor-pointer items-center justify-center gap-1 rounded-md border border-slate-200 bg-slate-50 py-1.5 text-[11px] font-medium text-slate-700 hover:border-[#3B5B7E] hover:text-[#3B5B7E] transition-colors">
                      <Upload className="h-3 w-3" />
                      <span>{activeName ? '更换文件' : '选择 Excel 文件'}</span>
                      <input
                        type="file"
                        accept=".xlsx,.xls"
                        onChange={(e) => handleSelectFile(slot.key, e)}
                        className="hidden"
                        disabled={loading}
                      />
                    </label>
                  </div>
                </div>
              );
            })}
          </div>

          {/* Date + Action Bar */}
          <div className="flex flex-wrap items-end justify-between gap-3 pt-1">
            <div>
              <label className="block text-[11px] font-medium text-slate-600 mb-1">
                确认单日期（留空默认当天，如 2026年03月16日）
              </label>
              <input
                value={dateInput}
                onChange={(e) => setDateInput(e.target.value)}
                placeholder="默认当天（yyyy年MM月dd日）"
                className="w-56 rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
              />
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <button
                type="button"
                onClick={handleRun}
                disabled={loading || !readyToRun}
                className="inline-flex items-center gap-1.5 rounded-md bg-[#3B5B7E] px-3.5 py-1.5 text-xs font-medium text-white hover:bg-[#2E4866] disabled:opacity-50 transition-colors"
              >
                {result ? <RefreshCw className="h-3.5 w-3.5" /> : <Play className="h-3.5 w-3.5" />}
                <span>{loading ? '生成中...' : result ? '重新生成确认单' : '生成预警事项确认单'}</span>
              </button>

              {result && (
                <a
                  href={`/api/evaluation/batch/${batchId}/warning-form/export`}
                  className="inline-flex items-center gap-1.5 rounded-md border border-[#2F6F5E]/30 bg-[#ECFDF5]/70 px-3 py-1.5 text-xs font-medium text-[#2F6F5E] hover:bg-[#D1FAE5]"
                >
                  <FileArchive className="h-3.5 w-3.5" />
                  <span>导出全部确认单 (.zip)</span>
                </a>
              )}

              {(state?.files.risk || state?.files.groups || state?.files.bids || result) && (
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

        {/* Result List */}
        {result && (
          <div className="space-y-2.5">
            <div className="text-[11px] text-slate-500 tabular-nums">
              已读取风险提示表 <strong>{result.riskRowCount}</strong> 行 · 标包分组表{' '}
              <strong>{result.groupRowCount}</strong> 行 · 开标信息表 <strong>{result.bidRowCount}</strong> 行
            </div>

            <div className="overflow-hidden rounded-lg border border-slate-200">
              <div className="overflow-x-auto">
                <table className="w-full border-collapse text-left text-xs">
                  <thead>
                    <tr className="border-b border-slate-200 bg-slate-50 text-slate-500">
                      <th className="px-2.5 py-2 font-medium whitespace-nowrap">评审组</th>
                      <th className="px-2.5 py-2 font-medium whitespace-nowrap">分标 / 包号</th>
                      <th className="px-2.5 py-2 font-medium min-w-[220px]">涉及投标人</th>
                      <th className="px-2.5 py-2 font-medium min-w-[200px]">预警事项明细</th>
                      <th className="px-2.5 py-2 font-medium whitespace-nowrap">单份下载</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {result.forms.length === 0 ? (
                      <tr>
                        <td colSpan={5} className="py-7 text-center text-xs text-[#2F6F5E]">
                          <span className="inline-flex items-center gap-1.5">
                            <CheckCircle2 className="h-4 w-4" />
                            未找到标记组合一致且具体信息相同的不同投标人（仍可导出 ZIP 查看使用说明）
                          </span>
                        </td>
                      </tr>
                    ) : (
                      paginate(result.forms, pager.page, pager.pageSize).map((item, idx) => {
                        const globalIdx = (pager.page - 1) * (pager.pageSize === 0 ? result.forms.length : pager.pageSize) + idx;
                        return (
                          <tr key={item.docxPath} className="hover:bg-slate-50/70 align-top">
                            <td className="px-2.5 py-2 whitespace-nowrap">
                              <span className="rounded bg-[#EBF2F7] px-2 py-0.5 text-[11px] font-medium text-[#3B5B7E]">
                                {item.group}
                              </span>
                            </td>
                            <td className="px-2.5 py-2 whitespace-nowrap">
                              <div className="font-medium text-slate-800">{item.section}</div>
                              <div className="text-[11px] text-slate-400">
                                {item.packs} · {item.labelNo}
                              </div>
                            </td>
                            <td className="px-2.5 py-2">
                              <div className="flex flex-wrap gap-1">
                                {item.suppliers.map((s) => (
                                  <span
                                    key={s}
                                    className="rounded border border-[#A94442]/25 bg-[#FEF2F2]/50 px-1.5 py-0.5 text-[11px] font-medium text-slate-800"
                                  >
                                    {s}
                                  </span>
                                ))}
                              </div>
                            </td>
                            <td className="px-2.5 py-2 text-[#A94442] leading-relaxed">{item.warnMatters}</td>
                            <td className="px-2.5 py-2 whitespace-nowrap">
                              <a
                                href={`/api/evaluation/batch/${batchId}/warning-form/export-single/${globalIdx}`}
                                className="inline-flex items-center gap-1 rounded border border-slate-200 bg-white px-2 py-1 text-[11px] font-medium text-[#3B5B7E] hover:bg-[#EBF2F7]/60"
                                title={item.docxFileName}
                              >
                                <Download className="h-3 w-3" />
                                <span>.docx</span>
                              </a>
                            </td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
              {result.forms.length > 0 && (
                <PaginationBar
                  total={result.forms.length}
                  page={pager.page}
                  pageSize={pager.pageSize}
                  onPageChange={pager.setPage}
                  onPageSizeChange={pager.setPageSize}
                />
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
