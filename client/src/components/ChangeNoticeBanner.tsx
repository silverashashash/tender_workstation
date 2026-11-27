import React, { useMemo, useState } from 'react';
import {
  AlertCircle,
  ChevronDown,
  ChevronUp,
  ExternalLink,
  FileText,
  Filter,
  Pencil,
  Plus,
  Search,
  Trash2,
  Upload,
  X,
} from 'lucide-react';
import { api } from '../api/client';
import { DateTimePickerField } from './DateTimePickerField';
import type { BatchChangeNotice, BatchProject } from '../types';

export const CHANGE_TYPE_OPTIONS = [
  '最高限价/预算',
  '资质/业绩',
  '人员要求',
  '开标延期',
  '项目终止/取消',
  '其他',
] as const;

interface ChangeNoticeBannerProps {
  batchId: string;
  notices: BatchChangeNotice[];
  projects?: BatchProject[];
  canManage?: boolean;
  onlyChanged?: boolean;
  onToggleOnlyChanged?: () => void;
  onChanged?: () => Promise<void> | void;
}

export function ChangeNoticeBanner({
  batchId,
  notices,
  projects = [],
  canManage = false,
  onlyChanged = false,
  onToggleOnlyChanged,
  onChanged,
}: ChangeNoticeBannerProps) {
  const [expanded, setExpanded] = useState(false);
  const [showModal, setShowModal] = useState(false);
  const [editingNotice, setEditingNotice] = useState<BatchChangeNotice | null>(null);

  // Form state
  const [title, setTitle] = useState('');
  const [noticeDate, setNoticeDate] = useState(() => new Date().toISOString().slice(0, 10));
  const [selectedTypes, setSelectedTypes] = useState<string[]>(['最高限价/预算']);
  const [summary, setSummary] = useState('');
  const [selectedProjectIds, setSelectedProjectIds] = useState<Set<string>>(new Set());
  const [projFilter, setProjFilter] = useState('');
  const [file, setFile] = useState<File | null>(null);
  const [removeExistingFile, setRemoveExistingFile] = useState(false);
  const [submitting, setSubmitting] = useState(false);

  const projectById = useMemo(
    () => new Map(projects.map((p) => [p.id, p])),
    [projects]
  );

  const affectedProjectCount = useMemo(() => {
    const ids = new Set<string>();
    notices.forEach((n) => n.project_ids.forEach((id) => ids.add(id)));
    return ids.size;
  }, [notices]);

  const filteredProjects = useMemo(() => {
    const q = projFilter.trim().toLowerCase();
    if (!q) return projects;
    const tokens = q.split(/\s+/);
    return projects.filter((p) => {
      const hay = [p.lot_code, p.lot_name, p.package_no, p.name, p.req_app_no]
        .filter(Boolean)
        .join(' ')
        .toLowerCase();
      return tokens.every((tk) => hay.includes(tk));
    });
  }, [projects, projFilter]);

  function toggleType(t: string) {
    setSelectedTypes((prev) =>
      prev.includes(t) ? prev.filter((x) => x !== t) : [...prev, t]
    );
  }

  function toggleProject(pid: string) {
    setSelectedProjectIds((prev) => {
      const next = new Set(prev);
      if (next.has(pid)) next.delete(pid);
      else next.add(pid);
      return next;
    });
  }

  function openCreateModal() {
    setEditingNotice(null);
    setTitle('');
    setNoticeDate(new Date().toISOString().slice(0, 10));
    setSelectedTypes(['最高限价/预算']);
    setSummary('');
    setSelectedProjectIds(new Set());
    setProjFilter('');
    setFile(null);
    setRemoveExistingFile(false);
    setShowModal(true);
  }

  function openEditModal(notice: BatchChangeNotice) {
    setEditingNotice(notice);
    setTitle(notice.title || '');
    setNoticeDate(notice.notice_date || '');
    setSelectedTypes(notice.change_types || []);
    setSummary(notice.summary || '');
    setSelectedProjectIds(new Set(notice.project_ids || []));
    setProjFilter('');
    setFile(null);
    setRemoveExistingFile(false);
    setShowModal(true);
  }

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    const hasAttachment = Boolean(file) || (Boolean(editingNotice?.has_file) && !removeExistingFile);
    if (!title.trim() && !summary.trim() && !hasAttachment) {
      alert('请至少填写公告标题/变更说明，或上传变更公告附件');
      return;
    }
    setSubmitting(true);
    try {
      if (editingNotice) {
        await api.updateChangeNotice(editingNotice.id, {
          title: title.trim(),
          notice_date: noticeDate.trim(),
          change_types: selectedTypes,
          summary: summary.trim(),
          project_ids: Array.from(selectedProjectIds),
          file,
          remove_file: removeExistingFile,
        });
      } else {
        await api.createChangeNotice(batchId, {
          title: title.trim(),
          notice_date: noticeDate.trim(),
          change_types: selectedTypes,
          summary: summary.trim(),
          project_ids: Array.from(selectedProjectIds),
          file,
        });
      }
      setShowModal(false);
      setEditingNotice(null);
      setExpanded(true);
      await onChanged?.();
    } catch (err) {
      alert((err as Error).message || '保存变更公告失败');
    } finally {
      setSubmitting(false);
    }
  }

  async function handleDelete(notice: BatchChangeNotice) {
    if (!window.confirm(`确认删除变更公告「${notice.title}」吗？`)) return;
    try {
      await api.deleteChangeNotice(notice.id);
      await onChanged?.();
    } catch (err) {
      alert((err as Error).message || '删除失败');
    }
  }

  if (notices.length === 0 && !canManage) {
    return null;
  }

  return (
    <>
      {notices.length > 0 ? (
        <div className="rounded-xl border border-[#9A6B26]/35 bg-[#FEF9EC]/85 px-4 py-3 space-y-2.5">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="flex items-center gap-2 text-xs font-semibold text-[#9A6B26]">
              <AlertCircle className="h-4 w-4 shrink-0" />
              <span>
                变更公告提醒：本批次存在 <strong>{notices.length}</strong> 份变更公告
                {affectedProjectCount > 0
                  ? `（涉及 ${affectedProjectCount} 个指定标包/项目，请核对最新限价与要求）`
                  : '（涉及整批次公告事项，请留意查看）'}
              </span>
            </div>

            <div className="flex flex-wrap items-center gap-1.5">
              {onToggleOnlyChanged && affectedProjectCount > 0 && (
                <button
                  type="button"
                  onClick={onToggleOnlyChanged}
                  className={`inline-flex items-center gap-1 rounded-md border px-2.5 py-1 text-[11px] font-medium transition-colors ${
                    onlyChanged
                      ? 'border-[#9A6B26] bg-[#9A6B26] text-white'
                      : 'border-[#9A6B26]/30 bg-white text-[#9A6B26] hover:bg-[#FEF9EC]'
                  }`}
                >
                  <Filter className="h-3 w-3" />
                  <span>{onlyChanged ? '显示全部项目' : `只看有变更 (${affectedProjectCount})`}</span>
                </button>
              )}
              <button
                type="button"
                onClick={() => setExpanded((v) => !v)}
                className="inline-flex items-center gap-1 rounded-md border border-[#9A6B26]/30 bg-white px-2.5 py-1 text-[11px] font-medium text-[#9A6B26] hover:bg-[#FEF9EC]"
              >
                <span>{expanded ? '收起明细' : '查看变更明细'}</span>
                {expanded ? <ChevronUp className="h-3 w-3" /> : <ChevronDown className="h-3 w-3" />}
              </button>
              {canManage && (
                <button
                  type="button"
                  onClick={openCreateModal}
                  className="inline-flex items-center gap-1 rounded-md bg-[#9A6B26] px-2.5 py-1 text-[11px] font-medium text-white hover:bg-[#7D561E]"
                >
                  <Plus className="h-3 w-3" />
                  <span>登记变更公告</span>
                </button>
              )}
            </div>
          </div>

          {/* Compact preview when collapsed */}
          {!expanded && notices[0] && (
            <div className="flex flex-wrap items-center justify-between gap-2 text-xs text-slate-700 border-t border-[#9A6B26]/15 pt-2">
              <div className="flex flex-wrap items-center gap-1.5 min-w-0">
                {notices[0].notice_date && (
                  <span className="font-mono text-[11px] text-slate-500">{notices[0].notice_date}</span>
                )}
                {notices[0].change_types.map((t) => (
                  <span
                    key={t}
                    className="rounded bg-[#9A6B26]/12 px-1.5 py-0.5 text-[10px] font-medium text-[#9A6B26]"
                  >
                    {t}
                  </span>
                ))}
                <span className="font-medium text-slate-800 truncate">{notices[0].title}</span>
                {notices[0].summary && (
                  <span className="text-slate-600 truncate max-w-md">— {notices[0].summary}</span>
                )}
              </div>
              <div className="flex items-center gap-2 shrink-0">
                {notices[0].has_file && (
                  <a
                    href={`/api/bidding/change-notices/${notices[0].id}/file`}
                    target="_blank"
                    rel="noreferrer"
                    className="inline-flex items-center gap-1 text-[11px] font-medium text-[#3B5B7E] hover:underline"
                  >
                    <ExternalLink className="h-3 w-3" />
                    <span>查看原件</span>
                  </a>
                )}
                {canManage && (
                  <button
                    type="button"
                    onClick={() => openEditModal(notices[0])}
                    className="inline-flex items-center gap-0.5 text-[11px] font-medium text-slate-500 hover:text-[#3B5B7E]"
                  >
                    <Pencil className="h-3 w-3" />
                    <span>修改</span>
                  </button>
                )}
              </div>
            </div>
          )}

          {/* Expanded list of all change notices */}
          {expanded && (
            <div className="space-y-2 border-t border-[#9A6B26]/15 pt-2.5">
              {notices.map((n) => {
                const affectedLabels =
                  n.affected_packages && n.affected_packages.length > 0
                    ? n.affected_packages.map((p) =>
                        [p.lot_code || p.lot_name, p.package_no, p.project_name].filter(Boolean).join(' · ')
                      )
                    : n.project_ids
                        .map((pid) => projectById.get(pid))
                        .filter((p): p is BatchProject => Boolean(p))
                        .map((p) => [p.lot_code || p.lot_name, p.package_no, p.name].filter(Boolean).join(' · '));

                return (
                  <div
                    key={n.id}
                    className="rounded-lg border border-[#9A6B26]/20 bg-white p-3 text-xs space-y-1.5"
                  >
                    <div className="flex flex-wrap items-start justify-between gap-2">
                      <div className="flex flex-wrap items-center gap-1.5">
                        <FileText className="h-3.5 w-3.5 text-[#9A6B26] shrink-0" />
                        <span className="font-semibold text-slate-800">{n.title}</span>
                        {n.notice_date && (
                          <span className="rounded bg-slate-100 px-1.5 py-0.5 font-mono text-[10px] text-slate-500">
                            {n.notice_date}
                          </span>
                        )}
                        {n.change_types.map((t) => (
                          <span
                            key={t}
                            className="rounded bg-[#FEF9EC] border border-[#9A6B26]/25 px-1.5 py-0.5 text-[10px] font-medium text-[#9A6B26]"
                          >
                            {t}
                          </span>
                        ))}
                      </div>
                      <div className="flex items-center gap-1.5 shrink-0">
                        {n.has_file && (
                          <a
                            href={`/api/bidding/change-notices/${n.id}/file`}
                            target="_blank"
                            rel="noreferrer"
                            className="inline-flex items-center gap-1 rounded border border-[#3B5B7E]/30 bg-[#EBF2F7]/60 px-2 py-0.5 text-[11px] font-medium text-[#3B5B7E] hover:bg-[#EBF2F7]"
                          >
                            <ExternalLink className="h-3 w-3" />
                            <span>查看原件 ({n.file_original_name || 'PDF'})</span>
                          </a>
                        )}
                        {canManage && (
                          <>
                            <button
                              type="button"
                              onClick={() => openEditModal(n)}
                              className="rounded p-1 text-slate-400 hover:bg-[#EBF2F7] hover:text-[#3B5B7E]"
                              title="修改该变更公告"
                            >
                              <Pencil className="h-3.5 w-3.5" />
                            </button>
                            <button
                              type="button"
                              onClick={() => handleDelete(n)}
                              className="rounded p-1 text-slate-400 hover:bg-red-50 hover:text-[#A94442]"
                              title="删除该变更公告"
                            >
                              <Trash2 className="h-3.5 w-3.5" />
                            </button>
                          </>
                        )}
                      </div>
                    </div>

                    {n.summary && (
                      <p className="text-slate-600 leading-relaxed whitespace-pre-wrap">{n.summary}</p>
                    )}

                    <div className="flex flex-wrap items-center gap-1.5 pt-0.5 text-[11px]">
                      <span className="text-slate-400">涉及范围：</span>
                      {affectedLabels.length > 0 ? (
                        affectedLabels.map((lbl, idx) => (
                          <span
                            key={idx}
                            className="rounded bg-slate-100 px-1.5 py-0.5 text-slate-700 max-w-full truncate"
                            title={lbl}
                          >
                            {lbl}
                          </span>
                        ))
                      ) : (
                        <span className="rounded bg-slate-100 px-1.5 py-0.5 text-slate-600">整批次</span>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      ) : (
        canManage && (
          <div className="flex justify-end">
            <button
              type="button"
              onClick={openCreateModal}
              className="inline-flex items-center gap-1.5 rounded-lg border border-[#9A6B26]/35 bg-[#FEF9EC]/70 px-3 py-1.5 text-xs font-medium text-[#9A6B26] hover:bg-[#FEF9EC] transition-colors"
            >
              <AlertCircle className="h-3.5 w-3.5" />
              <span>登记变更公告（上传 PDF / 标记涉及标包）</span>
            </button>
          </div>
        )
      )}

      {/* Create / Edit Change Notice Modal */}
      {showModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4 backdrop-blur-[1px]">
          <div className="w-full max-w-xl max-h-[90vh] flex flex-col rounded-xl border border-slate-200 bg-white shadow-xl overflow-hidden">
            <div className="flex items-center justify-between border-b border-slate-200 px-5 py-3.5">
              <div className="flex items-center gap-2">
                <AlertCircle className="h-4 w-4 text-[#9A6B26]" />
                <h3 className="text-sm font-semibold text-slate-800">
                  {editingNotice ? '修改变更公告' : '登记变更公告'}
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setShowModal(false)}
                className="rounded-lg p-1 text-slate-400 hover:bg-slate-100"
              >
                <X className="h-4 w-4" />
              </button>
            </div>

            <form onSubmit={handleSubmit} className="flex-1 overflow-y-auto p-5 space-y-4">
              <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
                <div className="sm:col-span-2">
                  <label className="block text-xs font-medium text-slate-600 mb-1">变更公告标题</label>
                  <input
                    value={title}
                    onChange={(e) => setTitle(e.target.value)}
                    placeholder="例如：部分标包最高限价变更公告（可留空自动取文件名）"
                    className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-600 mb-1">公告日期</label>
                  <DateTimePickerField value={noticeDate} onChange={setNoticeDate} />
                </div>
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-600 mb-1.5">变更类型（可多选）</label>
                <div className="flex flex-wrap gap-1.5">
                  {CHANGE_TYPE_OPTIONS.map((t) => {
                    const active = selectedTypes.includes(t);
                    return (
                      <button
                        key={t}
                        type="button"
                        onClick={() => toggleType(t)}
                        className={`rounded-md border px-2.5 py-1 text-xs font-medium transition-colors ${
                          active
                            ? 'border-[#9A6B26] bg-[#9A6B26] text-white'
                            : 'border-slate-200 bg-white text-slate-600 hover:bg-slate-50'
                        }`}
                      >
                        {t}
                      </button>
                    );
                  })}
                </div>
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-600 mb-1">变更核心内容摘要</label>
                <textarea
                  rows={3}
                  value={summary}
                  onChange={(e) => setSummary(e.target.value)}
                  placeholder="例如：分标 122617-9012008-0001 包28 最高限价由 51 万元改为 1098 万元"
                  className="w-full rounded-lg border border-slate-200 px-3 py-2 text-xs focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-600 mb-1">
                  变更公告原件（可选，支持 PDF / Word / 图片）
                </label>
                {editingNotice?.has_file && !file && !removeExistingFile && (
                  <div className="mb-1.5 flex items-center justify-between gap-2 rounded-lg border border-slate-200 bg-slate-50 px-3 py-1.5 text-xs">
                    <span className="truncate text-slate-700">
                      当前原件：<strong>{editingNotice.file_original_name || '已上传附件'}</strong>
                    </span>
                    <button
                      type="button"
                      onClick={() => setRemoveExistingFile(true)}
                      className="shrink-0 text-[11px] text-[#A94442] hover:underline"
                    >
                      移除原件
                    </button>
                  </div>
                )}
                <label className="flex cursor-pointer items-center justify-between gap-2 rounded-lg border border-dashed border-slate-300 bg-slate-50/70 px-3 py-2 text-xs text-slate-600 hover:border-[#3B5B7E] hover:bg-[#EBF2F7]/30">
                  <div className="flex items-center gap-2 min-w-0">
                    <Upload className="h-3.5 w-3.5 shrink-0 text-[#3B5B7E]" />
                    <span className="truncate">
                      {file
                        ? `将替换为：${file.name}`
                        : editingNotice?.has_file && !removeExistingFile
                          ? '点击上传新文件替换当前原件...'
                          : '点击选择变更公告 PDF / Word / 图片文件...'}
                    </span>
                  </div>
                  {file && (
                    <button
                      type="button"
                      onClick={(e) => {
                        e.preventDefault();
                        setFile(null);
                      }}
                      className="text-[11px] text-slate-400 hover:text-slate-600"
                    >
                      取消
                    </button>
                  )}
                  <input
                    type="file"
                    accept=".pdf,.doc,.docx,.png,.jpg,.jpeg,.webp"
                    onChange={(e) => {
                      const f = e.target.files?.[0] || null;
                      setFile(f);
                      if (f) setRemoveExistingFile(false);
                    }}
                    className="hidden"
                  />
                </label>
              </div>

              {projects.length > 0 && (
                <div>
                  <div className="flex items-center justify-between mb-1">
                    <label className="text-xs font-medium text-slate-600">
                      关联具体标包/项目（不勾选则视为「整批次变更」）
                    </label>
                    <span className="text-[11px] text-[#9A6B26] font-medium tabular-nums">
                      已选 {selectedProjectIds.size} 项
                    </span>
                  </div>
                  <div className="relative mb-1.5">
                    <Search className="pointer-events-none absolute left-2.5 top-2 h-3.5 w-3.5 text-slate-400" />
                    <input
                      value={projFilter}
                      onChange={(e) => setProjFilter(e.target.value)}
                      placeholder="搜索分标编号、包号或项目名称快速勾选..."
                      className="w-full rounded-lg border border-slate-200 py-1.5 pl-8 pr-3 text-xs focus:border-[#3B5B7E] focus:outline-none"
                    />
                  </div>
                  <div className="max-h-40 overflow-y-auto rounded-lg border border-slate-200 divide-y divide-slate-100">
                    {filteredProjects.length === 0 ? (
                      <div className="py-4 text-center text-xs text-slate-400">无匹配标包</div>
                    ) : (
                      filteredProjects.map((p) => {
                        const checked = selectedProjectIds.has(p.id);
                        return (
                          <label
                            key={p.id}
                            className={`flex cursor-pointer items-center gap-2 px-2.5 py-1.5 text-xs ${
                              checked ? 'bg-[#FEF9EC]/70' : 'hover:bg-slate-50'
                            }`}
                          >
                            <input
                              type="checkbox"
                              checked={checked}
                              onChange={() => toggleProject(p.id)}
                              className="rounded border-slate-300"
                            />
                            <span className="font-mono text-[11px] text-slate-500 shrink-0">
                              {[p.lot_code, p.package_no].filter(Boolean).join(' ')}
                            </span>
                            <span className="truncate text-slate-700">{p.name}</span>
                          </label>
                        );
                      })
                    )}
                  </div>
                </div>
              )}

              <div className="flex justify-end gap-2 border-t border-slate-100 pt-3">
                <button
                  type="button"
                  onClick={() => setShowModal(false)}
                  className="rounded-lg border border-slate-200 px-4 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"
                >
                  取消
                </button>
                <button
                  type="submit"
                  disabled={submitting}
                  className="rounded-lg bg-[#9A6B26] px-4 py-1.5 text-xs font-medium text-white hover:bg-[#7D561E] disabled:opacity-50"
                >
                  {submitting ? '保存中...' : editingNotice ? '保存修改' : '保存变更公告'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </>
  );
}
