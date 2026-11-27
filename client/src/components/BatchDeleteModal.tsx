import { useEffect, useState } from 'react';
import {
  AlertTriangle,
  Archive,
  CheckCircle2,
  Download,
  Lock,
  Trash2,
  X,
} from 'lucide-react';
import type { Batch } from '../types';

interface Props {
  batch: Batch;
  authEnabled: boolean;
  onCancel: () => void;
  onConfirmDelete: (payload: { confirmText: string; password?: string }) => Promise<void>;
}

export function BatchDeleteModal({ batch, authEnabled, onCancel, onConfirmDelete }: Props) {
  const expectedToken = (batch.code || '').trim() || batch.name.trim();
  const tokenLabel = (batch.code || '').trim() ? '批次编号' : '批次名称';

  const [confirmInput, setConfirmInput] = useState('');
  const [passwordInput, setPasswordInput] = useState('');
  const [backedUp, setBackedUp] = useState(false);
  const [cooldown, setCooldown] = useState(authEnabled ? 0 : 3);
  const [submitting, setSubmitting] = useState(false);
  const [errorMsg, setErrorMsg] = useState('');

  useEffect(() => {
    if (cooldown <= 0) return;
    const timer = setInterval(() => {
      setCooldown((c) => (c > 0 ? c - 1 : 0));
    }, 1000);
    return () => clearInterval(timer);
  }, [cooldown]);

  const tokenMatches = confirmInput.trim() === expectedToken;
  const passwordReady = !authEnabled || passwordInput.trim().length > 0;
  const canSubmit = tokenMatches && passwordReady && cooldown === 0 && !submitting;

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (!canSubmit) return;
    setSubmitting(true);
    setErrorMsg('');
    try {
      await onConfirmDelete({
        confirmText: confirmInput.trim(),
        password: authEnabled ? passwordInput : undefined,
      });
    } catch (err) {
      setErrorMsg((err as Error).message);
      setSubmitting(false);
    }
  }

  const st = batch.stats;

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-4 backdrop-blur-[1px]"
      onClick={onCancel}
    >
      <div
        className="w-full max-w-lg overflow-hidden rounded-xl border border-[#A94442]/30 bg-white shadow-2xl"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Header */}
        <div className="flex items-start justify-between gap-3 border-b border-[#A94442]/15 bg-[#FEF2F2]/80 px-5 py-3.5">
          <div className="flex items-start gap-2.5">
            <AlertTriangle className="mt-0.5 h-5 w-5 shrink-0 text-[#A94442]" />
            <div>
              <h3 className="text-sm font-semibold text-[#A94442]">永久彻底删除招标批次</h3>
              <p className="mt-0.5 text-[11px] text-[#A94442]/80">
                此操作不可撤销，将永久清空数据库记录及该批次已上传的所有物理文件。
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onCancel}
            disabled={submitting}
            className="rounded p-1 text-slate-400 hover:bg-white hover:text-slate-600"
          >
            <X className="h-4 w-4" />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="space-y-4 p-5">
          {/* Batch Identity & Impact Scope */}
          <div className="rounded-lg border border-slate-200 bg-slate-50/70 p-3.5 space-y-2.5">
            <div>
              <div className="text-xs font-semibold text-slate-800">{batch.name}</div>
              <div className="mt-0.5 flex flex-wrap gap-x-3 text-[11px] text-slate-500">
                <span className="font-mono">编号：{batch.code || '未编号'}</span>
                <span>单位：{batch.project_unit || '未填写'}</span>
              </div>
            </div>

            <div className="border-t border-slate-200/80 pt-2.5">
              <div className="mb-1.5 text-[11px] font-medium text-slate-600">
                即将一并永久清除的数据范围：
              </div>
              <div className="grid grid-cols-2 gap-2 sm:grid-cols-4">
                <div className="rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-center">
                  <div className="text-sm font-semibold text-slate-800 tabular-nums">
                    {st?.projectCount || 0}
                  </div>
                  <div className="text-[10px] text-slate-400">标段/项目</div>
                </div>
                <div className="rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-center">
                  <div className="text-sm font-semibold text-slate-800 tabular-nums">
                    {st?.biddingFileCount || 0}
                  </div>
                  <div className="text-[10px] text-slate-400">招标前期表</div>
                </div>
                <div className="rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-center">
                  <div className="text-sm font-semibold text-slate-800 tabular-nums">
                    {st?.evalFileCount || 0}
                  </div>
                  <div className="text-[10px] text-slate-400">开评标附件</div>
                </div>
                <div className="rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-center">
                  <div className="text-sm font-semibold text-slate-800 tabular-nums">
                    {st?.awardOutputCount || 0}
                  </div>
                  <div className="text-[10px] text-slate-400">定标交付件</div>
                </div>
              </div>
            </div>
          </div>

          {/* Safety Net: Backup Before Delete ("后悔药") */}
          <div className="flex flex-wrap items-center justify-between gap-2 rounded-lg border border-[#3B5B7E]/25 bg-[#EBF2F7]/50 px-3.5 py-2.5">
            <div className="flex items-start gap-2 min-w-0 flex-1">
              <Archive className="mt-0.5 h-4 w-4 shrink-0 text-[#3B5B7E]" />
              <div className="text-[11px] text-slate-600 leading-relaxed">
                <strong className="text-[#3B5B7E]">建议先留底备份（后悔药）</strong>：导出单批次无损 ZIP 包，后续随时可在侧边栏一键导入还原全部数据与原表。
              </div>
            </div>
            <a
              href={`/api/backup/export-batch/${batch.id}`}
              onClick={() => setBackedUp(true)}
              className="inline-flex shrink-0 items-center gap-1 rounded-md border border-[#3B5B7E]/30 bg-white px-2.5 py-1.5 text-xs font-medium text-[#3B5B7E] hover:bg-[#EBF2F7] transition-colors"
            >
              {backedUp ? (
                <>
                  <CheckCircle2 className="h-3.5 w-3.5 text-[#2F6F5E]" />
                  <span className="text-[#2F6F5E]">已触发备份下载</span>
                </>
              ) : (
                <>
                  <Download className="h-3.5 w-3.5" />
                  <span>先备份再删 (.zip)</span>
                </>
              )}
            </a>
          </div>

          {/* Hard Confirmation Input */}
          <div className="space-y-3">
            <div>
              <label className="block text-xs text-slate-600 mb-1.5">
                请完整输入该批次的{tokenLabel}{' '}
                <code className="rounded bg-slate-100 px-1.5 py-0.5 font-mono text-[11px] font-semibold text-[#A94442] select-all">
                  {expectedToken}
                </code>{' '}
                以确认删除：
              </label>
              <input
                autoFocus
                value={confirmInput}
                onChange={(e) => setConfirmInput(e.target.value)}
                placeholder={`请输入：${expectedToken}`}
                className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-xs font-mono focus:border-[#A94442] focus:outline-none"
                disabled={submitting}
              />
            </div>

            {authEnabled && (
              <div>
                <label className="flex items-center gap-1 text-xs text-slate-600 mb-1.5">
                  <Lock className="h-3.5 w-3.5 text-slate-400" />
                  <span>请输入工作台访问密码进行二次校验：</span>
                </label>
                <input
                  type="password"
                  value={passwordInput}
                  onChange={(e) => setPasswordInput(e.target.value)}
                  placeholder="输入工作台访问密码"
                  className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-xs focus:border-[#A94442] focus:outline-none"
                  disabled={submitting}
                />
              </div>
            )}
          </div>

          {errorMsg && (
            <div className="rounded-lg border border-[#A94442]/30 bg-[#FEF2F2] px-3 py-2 text-xs text-[#A94442]">
              {errorMsg}
            </div>
          )}

          {/* Footer Actions */}
          <div className="flex items-center justify-end gap-2 border-t border-slate-100 pt-3">
            <button
              type="button"
              onClick={onCancel}
              disabled={submitting}
              className="rounded-lg border border-slate-200 bg-white px-3.5 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"
            >
              取消
            </button>
            <button
              type="submit"
              disabled={!canSubmit}
              className="inline-flex items-center gap-1.5 rounded-lg bg-[#A94442] px-4 py-1.5 text-xs font-medium text-white hover:bg-[#8E3533] disabled:cursor-not-allowed disabled:opacity-45 transition-colors"
            >
              <Trash2 className="h-3.5 w-3.5" />
              <span>
                {submitting
                  ? '正在彻底清理...'
                  : cooldown > 0
                    ? `请确认影响范围 (${cooldown}s)`
                    : '永久彻底删除'}
              </span>
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}
