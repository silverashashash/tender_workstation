import { Clock, X } from 'lucide-react';

interface ParsedDateTime {
  date: string; // YYYY-MM-DD
  time: string; // HH:MM (24-hour) or ''
}

const HOURS_24 = Array.from({ length: 24 }, (_, i) => String(i).padStart(2, '0'));
const MINUTES_60 = Array.from({ length: 60 }, (_, i) => String(i).padStart(2, '0'));

function parseSingleDateTime(raw: string, fallbackYear?: string): ParsedDateTime {
  const s = String(raw || '').trim();
  if (!s) return { date: '', time: '' };

  let date = '';
  const fullMatch = s.match(/(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})/);
  if (fullMatch) {
    date = `${fullMatch[1]}-${fullMatch[2].padStart(2, '0')}-${fullMatch[3].padStart(2, '0')}`;
  } else if (fallbackYear) {
    const shortMatch = s.match(/(\d{1,2})[-/.月](\d{1,2})/);
    if (shortMatch) {
      date = `${fallbackYear}-${shortMatch[1].padStart(2, '0')}-${shortMatch[2].padStart(2, '0')}`;
    }
  }

  let time = '';
  const timeMatch = s.match(/(\d{1,2})[:：](\d{2})/);
  if (timeMatch) {
    const hh = Math.min(23, Math.max(0, parseInt(timeMatch[1], 10)));
    const mm = Math.min(59, Math.max(0, parseInt(timeMatch[2], 10)));
    time = `${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}`;
  }
  return { date, time };
}

function formatChineseDateTime(date: string, time: string): string {
  if (!date) return '';
  const m = date.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return date;
  const base = `${m[1]}年${m[2]}月${m[3]}日`;
  return time ? `${base} ${time}` : base;
}

function parseRange(raw: string): { start: ParsedDateTime; end: ParsedDateTime } {
  const s = String(raw || '').trim();
  if (!s) {
    return { start: { date: '', time: '' }, end: { date: '', time: '' } };
  }
  // Split range by "至", "~", " - ", or "-" after "日"/"时"
  const parts = s.split(/\s*(?:至|~|\s-\s|(?<=[日時时])\s*-\s*)\s*/);
  const start = parseSingleDateTime(parts[0] || '');
  const fallbackYear = start.date ? start.date.slice(0, 4) : undefined;
  const end = parts.length > 1 ? parseSingleDateTime(parts[1] || '', fallbackYear) : { date: '', time: '' };
  return { start, end };
}

function Time24Select({
  value,
  onChange,
  compact,
}: {
  value: string;
  onChange: (nextTime: string) => void;
  compact: boolean;
}) {
  const [rawH = '09', rawM = '00'] = value.split(':');
  const hh = HOURS_24.includes(rawH.padStart(2, '0')) ? rawH.padStart(2, '0') : '09';
  const mm = MINUTES_60.includes(rawM.padStart(2, '0')) ? rawM.padStart(2, '0') : '00';

  const selectCls = compact
    ? 'bg-transparent py-1 pl-1.5 pr-1 text-xs tabular-nums text-slate-700 focus:outline-none cursor-pointer'
    : 'bg-transparent py-1.5 pl-2 pr-1 text-xs tabular-nums text-slate-700 focus:outline-none cursor-pointer';

  return (
    <div
      className="inline-flex items-center rounded-lg border border-slate-200 bg-white focus-within:border-[#3B5B7E]"
      title="24小时制时间选择"
    >
      <select
        aria-label="小时（24小时制）"
        value={hh}
        onChange={(e) => onChange(`${e.target.value}:${mm}`)}
        className={selectCls}
      >
        {HOURS_24.map((h) => (
          <option key={h} value={h}>
            {h}
          </option>
        ))}
      </select>
      <span className="text-xs font-medium text-slate-400 select-none">:</span>
      <select
        aria-label="分钟"
        value={mm}
        onChange={(e) => onChange(`${hh}:${e.target.value}`)}
        className={selectCls}
      >
        {MINUTES_60.map((m) => (
          <option key={m} value={m}>
            {m}
          </option>
        ))}
      </select>
    </div>
  );
}

interface DateTimePickerFieldProps {
  value: string;
  onChange: (nextVal: string) => void;
  mode?: 'single' | 'range';
  allowTime?: boolean;
  compact?: boolean;
}

export function DateTimePickerField({
  value,
  onChange,
  mode = 'single',
  allowTime = true,
  compact = false,
}: DateTimePickerFieldProps) {
  const inputCls = compact
    ? 'rounded-lg border border-slate-200 bg-white px-2 py-1 text-xs tabular-nums focus:border-[#3B5B7E] focus:outline-none'
    : 'rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs tabular-nums focus:border-[#3B5B7E] focus:outline-none';

  if (mode === 'range') {
    const { start, end } = parseRange(value);
    const emitRange = (nextStart: ParsedDateTime, nextEnd: ParsedDateTime) => {
      const sText = formatChineseDateTime(nextStart.date, nextStart.time);
      const eText = formatChineseDateTime(nextEnd.date, nextEnd.time);
      if (sText && eText) onChange(`${sText} - ${eText}`);
      else onChange(sText || eText || '');
    };

    return (
      <div className="space-y-1.5">
        <div className="flex flex-wrap items-center gap-1.5">
          <span className="text-[11px] text-slate-400 shrink-0">起</span>
          <input
            type="date"
            value={start.date}
            onChange={(e) => emitRange({ ...start, date: e.target.value }, end)}
            className={`${inputCls} flex-1 min-w-[126px]`}
          />
          {allowTime &&
            (start.time ? (
              <div className="inline-flex items-center gap-1">
                <Time24Select
                  value={start.time}
                  onChange={(t) => emitRange({ ...start, time: t }, end)}
                  compact={compact}
                />
                <button
                  type="button"
                  onClick={() => emitRange({ ...start, time: '' }, end)}
                  className="rounded p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-600"
                  title="移除开始时间"
                >
                  <X className="h-3 w-3" />
                </button>
              </div>
            ) : (
              <button
                type="button"
                disabled={!start.date}
                onClick={() => emitRange({ ...start, time: '09:00' }, end)}
                className="inline-flex items-center gap-0.5 rounded-md border border-dashed border-slate-200 px-1.5 py-1 text-[11px] text-slate-400 hover:border-[#3B5B7E] hover:text-[#3B5B7E] disabled:opacity-40"
                title="添加可选具体时间（24小时制）"
              >
                <Clock className="h-3 w-3" />
                <span>时间</span>
              </button>
            ))}
        </div>
        <div className="flex flex-wrap items-center gap-1.5">
          <span className="text-[11px] text-slate-400 shrink-0">止</span>
          <input
            type="date"
            value={end.date}
            onChange={(e) => emitRange(start, { ...end, date: e.target.value })}
            className={`${inputCls} flex-1 min-w-[126px]`}
          />
          {allowTime &&
            (end.time ? (
              <div className="inline-flex items-center gap-1">
                <Time24Select
                  value={end.time}
                  onChange={(t) => emitRange(start, { ...end, time: t })}
                  compact={compact}
                />
                <button
                  type="button"
                  onClick={() => emitRange(start, { ...end, time: '' })}
                  className="rounded p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-600"
                  title="移除截止时间"
                >
                  <X className="h-3 w-3" />
                </button>
              </div>
            ) : (
              <button
                type="button"
                disabled={!end.date}
                onClick={() => emitRange(start, { ...end, time: '12:00' })}
                className="inline-flex items-center gap-0.5 rounded-md border border-dashed border-slate-200 px-1.5 py-1 text-[11px] text-slate-400 hover:border-[#3B5B7E] hover:text-[#3B5B7E] disabled:opacity-40"
                title="添加可选具体时间（24小时制）"
              >
                <Clock className="h-3 w-3" />
                <span>时间</span>
              </button>
            ))}
        </div>
      </div>
    );
  }

  const parsed = parseSingleDateTime(value);
  const emitSingle = (nextDate: string, nextTime: string) => {
    onChange(formatChineseDateTime(nextDate, nextTime));
  };

  return (
    <div className="flex flex-wrap items-center gap-1.5">
      <input
        type="date"
        value={parsed.date}
        onChange={(e) => emitSingle(e.target.value, parsed.time)}
        className={`${inputCls} flex-1 min-w-[132px]`}
      />
      {allowTime &&
        (parsed.time ? (
          <div className="inline-flex items-center gap-1">
            <Time24Select
              value={parsed.time}
              onChange={(t) => emitSingle(parsed.date, t)}
              compact={compact}
            />
            <button
              type="button"
              onClick={() => emitSingle(parsed.date, '')}
              className="rounded p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-600"
              title="清除具体时间（仅保留日期）"
            >
              <X className="h-3 w-3" />
            </button>
          </div>
        ) : (
          <button
            type="button"
            disabled={!parsed.date}
            onClick={() => emitSingle(parsed.date, '08:30')}
            className="inline-flex items-center gap-0.5 rounded-md border border-dashed border-slate-200 px-2 py-1.5 text-[11px] text-slate-400 hover:border-[#3B5B7E] hover:text-[#3B5B7E] disabled:opacity-40"
            title="可选添加具体时间（24小时制）"
          >
            <Clock className="h-3 w-3" />
            <span>+时间</span>
          </button>
        ))}
    </div>
  );
}
