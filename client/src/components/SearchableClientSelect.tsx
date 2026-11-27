import { useEffect, useRef, useState } from 'react';
import { ChevronDown, Search, Check, X } from 'lucide-react';
import type { ClientAgreement } from '../types';

interface SearchableClientSelectProps {
  clients: ClientAgreement[];
  /** Controlled value: either client.id (when valueMode="id") or client_name (when valueMode="name") */
  value: string;
  valueMode?: 'id' | 'name';
  placeholder?: string;
  allowCustom?: boolean;
  onChange: (val: string, matchedClient: ClientAgreement | null) => void;
}

export function SearchableClientSelect({
  clients,
  value,
  valueMode = 'name',
  placeholder = '请选择或搜索项目单位...',
  allowCustom = false,
  onChange,
}: SearchableClientSelectProps) {
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

  const selectedClient =
    valueMode === 'id'
      ? clients.find((c) => c.id === value) || null
      : clients.find((c) => c.client_name === value || c.client_name.replace(/（[^）]*）$/, '') === value) || null;

  const displayLabel =
    valueMode === 'id'
      ? selectedClient
        ? `${selectedClient.client_name} — ${selectedClient.discount_label}`
        : ''
      : value || '';

  const q = query.trim().toLowerCase();
  const filtered = clients.filter((c) => {
    if (!q) return true;
    return (
      c.client_name.toLowerCase().includes(q) ||
      c.group_name.toLowerCase().includes(q) ||
      c.contact_name.toLowerCase().includes(q) ||
      c.discount_label.toLowerCase().includes(q)
    );
  });

  const groups = [
    { label: '直属单位', items: filtered.filter((c) => c.group_name !== '主业单位') },
    { label: '主业单位', items: filtered.filter((c) => c.group_name === '主业单位') },
  ].filter((g) => g.items.length > 0);

  function handleSelect(c: ClientAgreement) {
    if (valueMode === 'id') {
      onChange(c.id, c);
    } else {
      onChange(c.client_name, c);
    }
    setOpen(false);
  }

  return (
    <div ref={containerRef} className="relative">
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        className={`flex w-full items-center justify-between gap-2 rounded-lg border bg-white px-3 py-2 text-left text-xs transition-colors ${
          open ? 'border-[#3B5B7E] ring-1 ring-[#3B5B7E]/20' : 'border-slate-200 hover:border-slate-300'
        }`}
      >
        <span className={`truncate ${displayLabel ? 'font-medium text-slate-800' : 'text-slate-400'}`}>
          {displayLabel || placeholder}
        </span>
        <div className="flex items-center gap-1 shrink-0">
          {displayLabel && (
            <span
              role="button"
              tabIndex={-1}
              onClick={(e) => {
                e.stopPropagation();
                onChange('', null);
              }}
              className="rounded p-0.5 text-slate-300 hover:bg-slate-100 hover:text-slate-500"
              title="清空选择"
            >
              <X className="h-3 w-3" />
            </span>
          )}
          <ChevronDown className={`h-3.5 w-3.5 text-slate-400 transition-transform ${open ? 'rotate-180' : ''}`} />
        </div>
      </button>

      {open && (
        <div className="absolute left-0 right-0 z-40 mt-1 overflow-hidden rounded-xl border border-slate-200 bg-white shadow-lg">
          {/* Search Input */}
          <div className="border-b border-slate-100 p-2">
            <div className="relative">
              <Search className="pointer-events-none absolute left-2.5 top-2 h-3.5 w-3.5 text-slate-400" />
              <input
                ref={inputRef}
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                placeholder="输入关键词快速筛选（如：铜陵、合肥、综合能源）..."
                className="w-full rounded-lg border border-slate-200 bg-slate-50/70 py-1.5 pl-8 pr-2.5 text-xs focus:border-[#3B5B7E] focus:bg-white focus:outline-none"
              />
            </div>
          </div>

          {/* Options List */}
          <div className="max-h-64 overflow-y-auto py-1">
            {groups.length === 0 ? (
              <div className="px-3 py-4 text-center text-xs text-slate-400">
                未找到匹配的协议项目单位
              </div>
            ) : (
              groups.map((grp) => (
                <div key={grp.label}>
                  <div className="sticky top-0 bg-slate-50/95 px-3 py-1 text-[10px] font-semibold text-slate-400">
                    {grp.label} ({grp.items.length})
                  </div>
                  {grp.items.map((c) => {
                    const isSelected =
                      valueMode === 'id' ? value === c.id : value === c.client_name;
                    return (
                      <button
                        key={c.id}
                        type="button"
                        onClick={() => handleSelect(c)}
                        className={`flex w-full items-center justify-between gap-2 px-3 py-2 text-left text-xs transition-colors ${
                          isSelected
                            ? 'bg-[#EBF2F7] text-[#3B5B7E] font-medium'
                            : 'text-slate-700 hover:bg-slate-50'
                        }`}
                      >
                        <div className="min-w-0 flex-1">
                          <div className="truncate">{c.client_name}</div>
                        </div>
                        <div className="flex items-center gap-1.5 shrink-0">
                          <span className="rounded bg-slate-100 px-1.5 py-0.5 text-[10px] text-slate-500 tabular-nums">
                            {c.discount_label}
                          </span>
                          {isSelected && <Check className="h-3.5 w-3.5 text-[#3B5B7E]" />}
                        </div>
                      </button>
                    );
                  })}
                </div>
              ))
            )}

            {allowCustom && query.trim() && !clients.some((c) => c.client_name === query.trim()) && (
              <button
                type="button"
                onClick={() => {
                  onChange(query.trim(), null);
                  setOpen(false);
                }}
                className="flex w-full items-center gap-1.5 border-t border-slate-100 px-3 py-2 text-left text-xs text-[#3B5B7E] hover:bg-[#EBF2F7]/50"
              >
                <span>使用自定义单位名称：</span>
                <strong className="truncate">{query.trim()}</strong>
              </button>
            )}
          </div>
        </div>
      )}
    </div>
  );
}
