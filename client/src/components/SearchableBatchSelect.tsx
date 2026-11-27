import { useEffect, useMemo, useRef, useState } from 'react';
import { Check, ChevronDown, Search, X } from 'lucide-react';
import type { Batch } from '../types';

interface SearchableBatchSelectProps {
  batches: Batch[];
  value: string;
  onChange: (batchId: string) => void;
}

export function SearchableBatchSelect({
  batches,
  value,
  onChange,
}: SearchableBatchSelectProps) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState('');
  const containerRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    function handleOutside(e: MouseEvent) {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    }
    document.addEventListener('mousedown', handleOutside);
    return () => document.removeEventListener('mousedown', handleOutside);
  }, []);

  useEffect(() => {
    if (open) {
      setTimeout(() => inputRef.current?.focus(), 20);
    } else {
      setQuery('');
    }
  }, [open]);

  const selectedBatch = useMemo(
    () => batches.find((b) => b.id === value) || null,
    [batches, value]
  );

  const filtered = useMemo(() => {
    const trimmed = query.trim().toLowerCase();
    if (!trimmed) return batches;
    const tokens = trimmed.split(/\s+/);
    return batches.filter((b) => {
      const haystack = [b.code, b.name, b.project_unit, b.status]
        .filter(Boolean)
        .join(' ')
        .toLowerCase();
      return tokens.every((tk) => haystack.includes(tk));
    });
  }, [batches, query]);

  function handleSelect(batchId: string) {
    onChange(batchId);
    setOpen(false);
  }

  const triggerTitle = selectedBatch
    ? [selectedBatch.code, selectedBatch.name, selectedBatch.status ? `[${selectedBatch.status}]` : '']
        .filter(Boolean)
        .join(' ')
    : '选择批次';

  return (
    <div ref={containerRef} className="relative min-w-0 flex-1 max-w-xl">
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        title={triggerTitle}
        className={`flex w-full items-center justify-between gap-2 rounded-lg border px-2.5 py-1.5 text-left transition-colors ${
          open
            ? 'border-[#3B5B7E] bg-white ring-1 ring-[#3B5B7E]/20'
            : 'border-slate-200 bg-slate-50/70 hover:border-slate-300 hover:bg-white'
        }`}
      >
        <div className="flex min-w-0 flex-1 items-center gap-2">
          {selectedBatch ? (
            <>
              {selectedBatch.code && (
                <span className="shrink-0 rounded bg-[#EBF2F7] px-1.5 py-0.5 font-mono text-xs font-semibold text-[#3B5B7E]">
                  {selectedBatch.code}
                </span>
              )}
              <span className="truncate text-sm font-medium text-slate-800">
                {selectedBatch.name}
              </span>
              {(selectedBatch.stats?.changeNoticeCount || 0) > 0 && (
                <span
                  className="shrink-0 rounded bg-[#FEF9EC] border border-[#9A6B26]/30 px-1.5 py-0.5 text-[10px] font-semibold text-[#9A6B26] tabular-nums"
                  title={`该批次有 ${selectedBatch.stats.changeNoticeCount} 份变更公告`}
                >
                  变更×{selectedBatch.stats.changeNoticeCount}
                </span>
              )}
              {selectedBatch.status && (
                <span className="hidden sm:inline-block shrink-0 rounded bg-slate-100 px-1.5 py-0.5 text-[11px] font-medium text-slate-500">
                  {selectedBatch.status}
                </span>
              )}
            </>
          ) : (
            <span className="truncate text-sm text-slate-400">请选择或搜索批次...</span>
          )}
        </div>
        <ChevronDown
          className={`h-4 w-4 shrink-0 text-slate-400 transition-transform ${open ? 'rotate-180' : ''}`}
        />
      </button>

      {open && (
        <div className="absolute left-0 z-40 mt-1.5 w-[min(92vw,560px)] overflow-hidden rounded-xl border border-slate-200 bg-white shadow-xl">
          {/* Search Input */}
          <div className="border-b border-slate-100 p-2.5">
            <div className="relative">
              <Search className="pointer-events-none absolute left-2.5 top-2 h-3.5 w-3.5 text-slate-400" />
              <input
                ref={inputRef}
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Escape') {
                    setOpen(false);
                  } else if (e.key === 'Enter' && filtered.length > 0) {
                    e.preventDefault();
                    handleSelect(filtered[0].id);
                  }
                }}
                placeholder="搜索批次编号或名称关键词（支持空格多词，如：CY1226 第六次）..."
                className="w-full rounded-lg border border-slate-200 bg-slate-50/70 py-1.5 pl-8 pr-8 text-xs text-slate-800 placeholder:text-slate-400 focus:border-[#3B5B7E] focus:bg-white focus:outline-none"
              />
              {query && (
                <button
                  type="button"
                  onClick={() => setQuery('')}
                  className="absolute right-2 top-1.5 rounded p-0.5 text-slate-400 hover:bg-slate-200/70 hover:text-slate-600"
                  title="清空搜索"
                >
                  <X className="h-3.5 w-3.5" />
                </button>
              )}
            </div>
            <div className="mt-1.5 flex items-center justify-between px-0.5 text-[11px] text-slate-400 tabular-nums">
              <span>
                {query.trim()
                  ? `匹配 ${filtered.length} / ${batches.length} 个批次`
                  : `共 ${batches.length} 个批次`}
              </span>
              <span>按 Enter 快速选中首项</span>
            </div>
          </div>

          {/* Batch Options List */}
          <div className="max-h-80 overflow-y-auto divide-y divide-slate-100">
            {filtered.length === 0 ? (
              <div className="px-4 py-6 text-center text-xs text-slate-400">
                未找到匹配“{query.trim()}”的批次
              </div>
            ) : (
              filtered.map((b) => {
                const isSelected = b.id === value;
                return (
                  <button
                    key={b.id}
                    type="button"
                    onClick={() => handleSelect(b.id)}
                    className={`flex w-full items-start justify-between gap-2.5 px-3 py-2.5 text-left transition-colors ${
                      isSelected ? 'bg-[#EBF2F7]/80' : 'hover:bg-slate-50'
                    }`}
                  >
                    <div className="min-w-0 flex-1 space-y-1">
                      <div className="flex flex-wrap items-center gap-1.5">
                        <span
                          className={`rounded px-1.5 py-0.5 font-mono text-[11px] font-semibold ${
                            b.code
                              ? isSelected
                                ? 'bg-[#3B5B7E] text-white'
                                : 'bg-[#EBF2F7] text-[#3B5B7E]'
                              : 'bg-slate-100 text-slate-400'
                          }`}
                        >
                          {b.code || '未编号'}
                        </span>
                        {(b.stats?.changeNoticeCount || 0) > 0 && (
                          <span className="rounded bg-[#FEF9EC] border border-[#9A6B26]/30 px-1.5 py-0.5 text-[10px] font-semibold text-[#9A6B26] tabular-nums">
                            变更×{b.stats.changeNoticeCount}
                          </span>
                        )}
                        {b.status && (
                          <span className="rounded bg-slate-100 px-1.5 py-0.5 text-[10px] font-medium text-slate-500">
                            {b.status}
                          </span>
                        )}
                      </div>
                      <div
                        className={`text-xs leading-snug break-words ${
                          isSelected ? 'font-semibold text-[#2E4866]' : 'text-slate-700'
                        }`}
                      >
                        {b.name}
                      </div>
                    </div>
                    {isSelected && (
                      <Check className="mt-0.5 h-4 w-4 shrink-0 text-[#3B5B7E]" />
                    )}
                  </button>
                );
              })
            )}
          </div>
        </div>
      )}
    </div>
  );
}
