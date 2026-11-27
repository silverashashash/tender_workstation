import React from 'react';
import { X, FileSpreadsheet } from 'lucide-react';
import type { ParsedSheetTable } from '../types';

interface Props {
  title: string;
  sheets: ParsedSheetTable[];
  onClose: () => void;
}

export function ExcelPreviewModal({ title, sheets, onClose }: Props) {
  const [activeIdx, setActiveIdx] = React.useState(0);
  const sheet = sheets[activeIdx];

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4 backdrop-blur-[1px]">
      <div className="flex max-h-[85vh] w-full max-w-5xl flex-col overflow-hidden rounded-xl border border-slate-200 bg-white shadow-xl">
        <div className="flex items-center justify-between border-b border-slate-200 px-5 py-3.5">
          <div className="flex items-center gap-2.5 min-w-0">
            <FileSpreadsheet className="h-5 w-5 shrink-0 text-[#3B5B7E]" />
            <h3 className="truncate text-sm font-semibold text-slate-800">{title}</h3>
            {sheet && (
              <span className="shrink-0 rounded bg-slate-100 px-2 py-0.5 text-xs text-slate-500 tabular-nums">
                共 {sheet.totalRows} 行
              </span>
            )}
          </div>
          <button
            onClick={onClose}
            className="rounded-lg p-1.5 text-slate-400 hover:bg-slate-100 hover:text-slate-600"
          >
            <X className="h-4 w-4" />
          </button>
        </div>

        {sheets.length > 1 && (
          <div className="flex gap-1 border-b border-slate-200 bg-slate-50 px-5 py-2 overflow-x-auto">
            {sheets.map((s, idx) => (
              <button
                key={s.name}
                onClick={() => setActiveIdx(idx)}
                className={`rounded-md px-3 py-1 text-xs font-medium transition-colors whitespace-nowrap ${
                  idx === activeIdx
                    ? 'bg-[#3B5B7E] text-white'
                    : 'text-slate-600 hover:bg-slate-200/70'
                }`}
              >
                {s.name} ({s.totalRows})
              </button>
            ))}
          </div>
        )}

        <div className="flex-1 overflow-auto p-4">
          {sheet && sheet.rows.length > 0 ? (
            <table className="w-full border-collapse text-left text-xs">
              <thead>
                <tr className="border-b border-slate-200 bg-slate-50 text-slate-600">
                  <th className="px-3 py-2 font-medium w-10">#</th>
                  {sheet.headers.map((h) => (
                    <th key={h} className="px-3 py-2 font-medium whitespace-nowrap">
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {sheet.rows.map((row, rIdx) => (
                  <tr key={rIdx} className="hover:bg-slate-50/80">
                    <td className="px-3 py-2 text-slate-400 tabular-nums">{rIdx + 1}</td>
                    {sheet.headers.map((h) => (
                      <td key={h} className="px-3 py-2 text-slate-700 max-w-xs truncate" title={row[h]}>
                        {row[h] || '—'}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          ) : (
            <div className="py-12 text-center text-sm text-slate-400">该工作表暂无数据</div>
          )}
        </div>
      </div>
    </div>
  );
}
