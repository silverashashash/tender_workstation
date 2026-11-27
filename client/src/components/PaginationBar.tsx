import { useEffect, useState } from 'react';
import { ChevronLeft, ChevronRight, ChevronsLeft, ChevronsRight } from 'lucide-react';

export const PAGE_SIZE_OPTIONS = [20, 50, 0] as const; // 0 = 全部

interface PaginationBarProps {
  total: number;
  page: number;
  pageSize: number; // 0 means show all
  onPageChange: (page: number) => void;
  onPageSizeChange: (size: number) => void;
}

export function PaginationBar({
  total,
  page,
  pageSize,
  onPageChange,
  onPageSizeChange,
}: PaginationBarProps) {
  if (total <= 20) return null;

  const isAll = pageSize === 0;
  const totalPages = isAll ? 1 : Math.max(1, Math.ceil(total / pageSize));
  const safePage = Math.min(Math.max(1, page), totalPages);
  const start = total === 0 ? 0 : isAll ? 1 : (safePage - 1) * pageSize + 1;
  const end = isAll ? total : Math.min(total, safePage * pageSize);

  return (
    <div className="flex flex-col gap-2 border-t border-slate-200 bg-slate-50/60 px-4 py-2.5 text-xs text-slate-600 sm:flex-row sm:items-center sm:justify-between">
      <div className="flex flex-wrap items-center gap-2 tabular-nums">
        <span>
          共 <strong className="text-slate-800">{total}</strong> 项 · 当前第{' '}
          <strong className="text-slate-800">
            {start}–{end}
          </strong>{' '}
          项
        </span>
        <span className="text-slate-300">|</span>
        <div className="inline-flex items-center gap-1">
          <span className="text-slate-500">每页展示</span>
          <div className="inline-flex rounded-md border border-slate-200 bg-white p-0.5">
            {PAGE_SIZE_OPTIONS.map((opt) => (
              <button
                key={opt}
                type="button"
                onClick={() => {
                  onPageSizeChange(opt);
                  onPageChange(1);
                }}
                className={`rounded px-2 py-0.5 text-[11px] font-medium transition-colors ${
                  pageSize === opt
                    ? 'bg-[#3B5B7E] text-white'
                    : 'text-slate-600 hover:bg-slate-100'
                }`}
              >
                {opt === 0 ? '全部' : `${opt} 项`}
              </button>
            ))}
          </div>
        </div>
      </div>

      {!isAll && totalPages > 1 && (
        <div className="flex items-center gap-1 tabular-nums">
          <button
            type="button"
            onClick={() => onPageChange(1)}
            disabled={safePage <= 1}
            className="rounded-md border border-slate-200 bg-white p-1 text-slate-500 hover:bg-slate-100 disabled:opacity-40"
            title="首页"
          >
            <ChevronsLeft className="h-3.5 w-3.5" />
          </button>
          <button
            type="button"
            onClick={() => onPageChange(safePage - 1)}
            disabled={safePage <= 1}
            className="inline-flex items-center gap-0.5 rounded-md border border-slate-200 bg-white px-2 py-1 text-xs font-medium text-slate-600 hover:bg-slate-100 disabled:opacity-40"
          >
            <ChevronLeft className="h-3.5 w-3.5" /> 上一页
          </button>
          <span className="px-2 text-xs text-slate-500">
            第 <strong className="text-slate-800">{safePage}</strong> / {totalPages} 页
          </span>
          <button
            type="button"
            onClick={() => onPageChange(safePage + 1)}
            disabled={safePage >= totalPages}
            className="inline-flex items-center gap-0.5 rounded-md border border-slate-200 bg-white px-2 py-1 text-xs font-medium text-slate-600 hover:bg-slate-100 disabled:opacity-40"
          >
            下一页 <ChevronRight className="h-3.5 w-3.5" />
          </button>
          <button
            type="button"
            onClick={() => onPageChange(totalPages)}
            disabled={safePage >= totalPages}
            className="rounded-md border border-slate-200 bg-white p-1 text-slate-500 hover:bg-slate-100 disabled:opacity-40"
            title="末页"
          >
            <ChevronsRight className="h-3.5 w-3.5" />
          </button>
        </div>
      )}
    </div>
  );
}

export function paginate<T>(items: T[], page: number, pageSize: number): T[] {
  if (pageSize === 0) return items;
  const totalPages = Math.max(1, Math.ceil(items.length / pageSize));
  const safePage = Math.min(Math.max(1, page), totalPages);
  const start = (safePage - 1) * pageSize;
  return items.slice(start, start + pageSize);
}

export function usePagination(resetKey?: string | number) {
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState<number>(20);

  useEffect(() => {
    setPage(1);
  }, [resetKey]);

  return { page, setPage, pageSize, setPageSize };
}
