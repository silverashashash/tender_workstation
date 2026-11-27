import { useMemo, useState } from 'react';
import {
  CalendarDays,
  Clock,
  Flame,
  PieChart as PieChartIcon,
  TrendingUp,
  ArrowRight,
  AlertTriangle,
  CheckCircle2,
} from 'lucide-react';
import type { Batch } from '../types';

export const STATUS_ORDER = ['筹备中', '已挂网', '开评标中', '归档中', '已完成'] as const;

export const STATUS_META: Record<
  string,
  { color: string; bgClass: string; textClass: string; dotClass: string }
> = {
  筹备中: {
    color: '#94A3B8',
    bgClass: 'bg-slate-100',
    textClass: 'text-slate-600',
    dotClass: 'bg-slate-400',
  },
  已挂网: {
    color: '#D49A3A',
    bgClass: 'bg-[#FEF9EC]',
    textClass: 'text-[#9A6B26]',
    dotClass: 'bg-[#D49A3A]',
  },
  开评标中: {
    color: '#3B5B7E',
    bgClass: 'bg-[#EBF2F7]',
    textClass: 'text-[#3B5B7E]',
    dotClass: 'bg-[#3B5B7E]',
  },
  归档中: {
    color: '#2F6F5E',
    bgClass: 'bg-[#ECFDF5]',
    textClass: 'text-[#2F6F5E]',
    dotClass: 'bg-[#2F6F5E]',
  },
  已完成: {
    color: '#10B981',
    bgClass: 'bg-emerald-50',
    textClass: 'text-emerald-700',
    dotClass: 'bg-emerald-500',
  },
};

export function parseDateOnly(dateStr?: string | null): Date | null {
  if (!dateStr) return null;
  const cleaned = String(dateStr).trim();
  if (!cleaned) return null;
  const match = cleaned.match(/(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})/);
  if (match) {
    const y = Number(match[1]);
    const m = Number(match[2]) - 1;
    const d = Number(match[3]);
    const dt = new Date(y, m, d);
    if (!Number.isNaN(dt.getTime())) return dt;
  }
  const parsed = new Date(cleaned.replace(' ', 'T'));
  if (Number.isNaN(parsed.getTime())) return null;
  return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

export function formatDateKey(dt: Date): string {
  const y = dt.getFullYear();
  const m = String(dt.getMonth() + 1).padStart(2, '0');
  const d = String(dt.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

export function getCountdownLabel(dateStr: string): { text: string; urgent: boolean; diffDays: number } | null {
  const parsed = parseDateOnly(dateStr);
  if (!parsed) return null;
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const diffDays = Math.round((parsed.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
  if (diffDays > 0 && diffDays <= 7) {
    return { text: `${diffDays} 天后开标`, urgent: true, diffDays };
  }
  if (diffDays > 7) {
    return { text: `${diffDays} 天后开标`, urgent: false, diffDays };
  }
  if (diffDays === 0) {
    return { text: '今日开标', urgent: true, diffDays: 0 };
  }
  return { text: '已开标', urgent: false, diffDays };
}

/* ==================== 1. Status Donut Chart ==================== */
interface StatusDonutProps {
  batches: Batch[];
  selectedStatus: string;
  onSelectStatus: (status: string) => void;
}

export function StatusDonutChart({ batches, selectedStatus, onSelectStatus }: StatusDonutProps) {
  const [hovered, setHovered] = useState<string | null>(null);

  const segments = useMemo(() => {
    const counts = new Map<string, number>();
    for (const st of STATUS_ORDER) counts.set(st, 0);
    for (const b of batches) {
      const k = b.status || '筹备中';
      counts.set(k, (counts.get(k) || 0) + 1);
    }
    const total = batches.length;
    const radius = 44;
    const circumference = 2 * Math.PI * radius;
    let offset = 0;

    return Array.from(counts.entries()).map(([status, count]) => {
      const pct = total > 0 ? count / total : 0;
      const dash = pct * circumference;
      const seg = {
        status,
        count,
        pct: Math.round(pct * 100),
        color: STATUS_META[status]?.color || '#94A3B8',
        dash,
        gap: circumference - dash,
        offset,
      };
      offset += dash;
      return seg;
    });
  }, [batches]);

  const activeFocus = hovered || (selectedStatus !== 'all' ? selectedStatus : null);
  const focusSeg = activeFocus ? segments.find((s) => s.status === activeFocus) : null;
  const inProgressCount = batches.filter((b) => b.status !== '已完成').length;

  return (
    <div className="flex h-full flex-col rounded-xl border border-slate-200 bg-white p-4">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-1.5">
          <PieChartIcon className="h-4 w-4 text-[#3B5B7E]" />
          <h2 className="text-sm font-semibold text-slate-800">批次阶段分布</h2>
        </div>
        <span className="text-[11px] text-slate-400">点击状态联动筛选</span>
      </div>

      <div className="mt-4 flex flex-1 flex-col items-center justify-center gap-4 sm:flex-row sm:gap-6">
        {/* Donut SVG */}
        <div className="relative h-32 w-32 shrink-0">
          <svg viewBox="0 0 120 120" className="h-full w-full -rotate-90">
            <circle
              cx="60"
              cy="60"
              r="44"
              fill="none"
              stroke="#F1F5F9"
              strokeWidth="14"
            />
            {batches.length > 0 &&
              segments
                .filter((s) => s.count > 0)
                .map((s) => {
                  const isSelected = selectedStatus === s.status;
                  const isDimmed = selectedStatus !== 'all' && !isSelected;
                  return (
                    <circle
                      key={s.status}
                      cx="60"
                      cy="60"
                      r="44"
                      fill="none"
                      stroke={s.color}
                      strokeWidth={isSelected || hovered === s.status ? 17 : 14}
                      strokeDasharray={`${s.dash} ${s.gap}`}
                      strokeDashoffset={-s.offset}
                      strokeLinecap="butt"
                      opacity={isDimmed ? 0.35 : 1}
                      className="cursor-pointer transition-all duration-200"
                      onMouseEnter={() => setHovered(s.status)}
                      onMouseLeave={() => setHovered(null)}
                      onClick={() =>
                        onSelectStatus(selectedStatus === s.status ? 'all' : s.status)
                      }
                    />
                  );
                })}
          </svg>
          <div className="pointer-events-none absolute inset-0 flex flex-col items-center justify-center text-center">
            <span className="text-xl font-semibold text-slate-800 tabular-nums">
              {focusSeg ? focusSeg.count : inProgressCount}
            </span>
            <span className="text-[10px] text-slate-400">
              {focusSeg ? `${focusSeg.status} (${focusSeg.pct}%)` : '在办批次'}
            </span>
          </div>
        </div>

        {/* Interactive Legend */}
        <div className="w-full flex-1 space-y-1.5">
          {segments.map((s) => {
            const isSelected = selectedStatus === s.status;
            return (
              <button
                key={s.status}
                type="button"
                onMouseEnter={() => setHovered(s.status)}
                onMouseLeave={() => setHovered(null)}
                onClick={() => onSelectStatus(isSelected ? 'all' : s.status)}
                className={`flex w-full items-center justify-between rounded-lg px-2.5 py-1.5 text-xs transition-colors ${
                  isSelected
                    ? 'bg-[#EBF2F7] text-[#3B5B7E] font-medium'
                    : 'hover:bg-slate-50 text-slate-600'
                }`}
              >
                <span className="flex items-center gap-2">
                  <span
                    className="h-2.5 w-2.5 rounded-full shrink-0"
                    style={{ backgroundColor: s.color }}
                  />
                  <span>{s.status}</span>
                </span>
                <span className="flex items-center gap-2 tabular-nums">
                  <span className="font-semibold text-slate-800">{s.count}</span>
                  <span className="w-8 text-right text-[11px] text-slate-400">{s.pct}%</span>
                </span>
              </button>
            );
          })}
        </div>
      </div>
    </div>
  );
}

/* ==================== 2. Calendar Activity / Bid-Opening Heatmap ==================== */
interface HeatmapProps {
  batches: Batch[];
  selectedDateKey: string | null;
  onSelectDateKey: (dateKey: string | null) => void;
  onOpenBatch: (batchId: string, path: string) => void;
}

export function BidCalendarHeatmap({
  batches,
  selectedDateKey,
  onSelectDateKey,
  onOpenBatch,
}: HeatmapProps) {
  const [mode, setMode] = useState<'opening' | 'all'>('opening');
  const [hoveredKey, setHoveredKey] = useState<string | null>(null);

  const { weeks, monthLabels, todayKey, dateMap, totalHits } = useMemo(() => {
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const tKey = formatDateKey(today);

    // Build 18 weeks: 9 weeks back + current week + 8 weeks ahead (aligned to Monday)
    const dow = (today.getDay() + 6) % 7; // Mon=0..Sun=6
    const startMonday = new Date(today);
    startMonday.setDate(today.getDate() - dow - 9 * 7);

    const map = new Map<string, Array<{ batch: Batch; tag: string }>>();
    const addHit = (dt: Date | null, batch: Batch, tag: string) => {
      if (!dt) return;
      const k = formatDateKey(dt);
      const list = map.get(k) || [];
      if (!list.some((x) => x.batch.id === batch.id && x.tag === tag)) {
        list.push({ batch, tag });
        map.set(k, list);
      }
    };

    for (const b of batches) {
      addHit(parseDateOnly(b.bid_opening_time), b, '开标');
      if (mode === 'all') {
        addHit(parseDateOnly(b.announcement_time), b, '公告');
        addHit(parseDateOnly(b.created_at), b, '建档');
      }
    }

    const wks: Array<Array<{ date: Date; key: string; diffFromToday: number }>> = [];
    const mLabels: Array<{ weekIdx: number; label: string }> = [];
    let lastMonth = -1;
    let hitsInWindow = 0;

    for (let w = 0; w < 18; w++) {
      const col: Array<{ date: Date; key: string; diffFromToday: number }> = [];
      for (let d = 0; d < 7; d++) {
        const cur = new Date(startMonday);
        cur.setDate(startMonday.getDate() + w * 7 + d);
        const k = formatDateKey(cur);
        const diff = Math.round((cur.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
        if (d === 0 && cur.getMonth() !== lastMonth) {
          lastMonth = cur.getMonth();
          mLabels.push({ weekIdx: w, label: `${cur.getMonth() + 1}月` });
        }
        hitsInWindow += map.get(k)?.length || 0;
        col.push({ date: cur, key: k, diffFromToday: diff });
      }
      wks.push(col);
    }

    return { weeks: wks, monthLabels: mLabels, todayKey: tKey, dateMap: map, totalHits: hitsInWindow };
  }, [batches, mode]);

  const activeKey = hoveredKey || selectedDateKey;
  const activeEntries = activeKey ? dateMap.get(activeKey) || [] : [];

  function cellColor(count: number, diffFromToday: number): string {
    if (count === 0) return 'bg-slate-100 hover:bg-slate-200/80';
    const isSoon = diffFromToday >= 0 && diffFromToday <= 7;
    if (isSoon) {
      if (count >= 2) return 'bg-[#B45309] hover:bg-[#92400E]';
      return 'bg-[#D97706] hover:bg-[#B45309]';
    }
    if (count === 1) return 'bg-[#B8D0E4] hover:bg-[#9BBCD8]';
    if (count === 2) return 'bg-[#658CB2] hover:bg-[#4F779E]';
    return 'bg-[#3B5B7E] hover:bg-[#2E4866]';
  }

  return (
    <div className="flex h-full flex-col rounded-xl border border-slate-200 bg-white p-4">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div className="flex items-center gap-1.5">
          <Flame className="h-4 w-4 text-[#9A6B26]" />
          <h2 className="text-sm font-semibold text-slate-800">排期节点热力图</h2>
          <span className="rounded bg-slate-100 px-1.5 py-0.5 text-[10px] text-slate-500 tabular-nums">
            前后 18 周 · {totalHits} 个节点
          </span>
        </div>

        <div className="flex items-center gap-1 rounded-lg bg-slate-100 p-0.5 text-[11px]">
          <button
            type="button"
            onClick={() => setMode('opening')}
            className={`rounded-md px-2 py-0.5 font-medium transition-colors ${
              mode === 'opening' ? 'bg-white text-[#3B5B7E] shadow-xs' : 'text-slate-500 hover:text-slate-700'
            }`}
          >
            仅开标日
          </button>
          <button
            type="button"
            onClick={() => setMode('all')}
            className={`rounded-md px-2 py-0.5 font-medium transition-colors ${
              mode === 'all' ? 'bg-white text-[#3B5B7E] shadow-xs' : 'text-slate-500 hover:text-slate-700'
            }`}
          >
            全部节点
          </button>
        </div>
      </div>

      {/* Heatmap Matrix */}
      <div className="mt-3 overflow-x-auto pb-1">
        <div className="min-w-[440px]">
          {/* Month Row */}
          <div className="mb-1 flex pl-7 text-[10px] text-slate-400">
            {weeks.map((_, wIdx) => {
              const m = monthLabels.find((x) => x.weekIdx === wIdx);
              return (
                <div key={wIdx} className="flex-1 text-left">
                  {m ? m.label : ''}
                </div>
              );
            })}
          </div>

          <div className="flex gap-1.5">
            {/* Weekday Labels */}
            <div className="flex w-5 shrink-0 flex-col justify-between py-0.5 text-[10px] text-slate-400">
              <span>一</span>
              <span>三</span>
              <span>五</span>
              <span>日</span>
            </div>

            {/* 18 Columns x 7 Rows */}
            <div className="grid flex-1 grid-cols-[repeat(18,minmax(0,1fr))] gap-1">
              {weeks.map((week, wIdx) => (
                <div key={wIdx} className="flex flex-col gap-1">
                  {week.map((day) => {
                    const hits = dateMap.get(day.key) || [];
                    const isToday = day.key === todayKey;
                    const isSelected = selectedDateKey === day.key;
                    return (
                      <button
                        key={day.key}
                        type="button"
                        onMouseEnter={() => setHoveredKey(day.key)}
                        onMouseLeave={() => setHoveredKey(null)}
                        onClick={() => {
                          if (hits.length === 0) return;
                          onSelectDateKey(isSelected ? null : day.key);
                        }}
                        title={`${day.key}${isToday ? ' (今日)' : ''}${
                          hits.length > 0 ? ` · ${hits.length} 个节点` : ' · 无排期'
                        }`}
                        className={`aspect-square w-full rounded-[3px] transition-all ${cellColor(
                          hits.length,
                          day.diffFromToday
                        )} ${isToday ? 'ring-2 ring-[#3B5B7E] ring-offset-1' : ''} ${
                          isSelected ? 'ring-2 ring-[#9A6B26] ring-offset-1' : ''
                        } ${hits.length > 0 ? 'cursor-pointer' : 'cursor-default'}`}
                      />
                    );
                  })}
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>

      {/* Legend + Hover/Selection Preview Bar */}
      <div className="mt-auto flex flex-wrap items-center justify-between gap-2 border-t border-slate-100 pt-2.5 text-[11px]">
        <div className="min-w-0 flex-1 truncate text-slate-500">
          {activeKey ? (
            activeEntries.length > 0 ? (
              <span className="inline-flex flex-wrap items-center gap-1.5">
                <strong className="font-mono text-slate-700">{activeKey}</strong>
                {activeEntries.slice(0, 2).map((item, i) => (
                  <button
                    key={`${item.batch.id}-${i}`}
                    type="button"
                    onClick={() => onOpenBatch(item.batch.id, '/evaluation')}
                    className="inline-flex items-center gap-1 rounded bg-[#EBF2F7] px-1.5 py-0.5 text-[10px] font-medium text-[#3B5B7E] hover:bg-[#DCE8F2]"
                  >
                    <span>{item.tag}</span>
                    <span className="max-w-[140px] truncate">{item.batch.name}</span>
                  </button>
                ))}
                {activeEntries.length > 2 && (
                  <span className="text-slate-400">等 {activeEntries.length} 项</span>
                )}
              </span>
            ) : (
              <span>
                <span className="font-mono text-slate-600">{activeKey}</span> 无排期节点
              </span>
            )
          ) : (
            <span>悬浮查看当日排期，点击有色方格联动筛选下方台账</span>
          )}
        </div>

        <div className="flex items-center gap-1.5 shrink-0 text-[10px] text-slate-400">
          <span>少</span>
          <span className="h-2.5 w-2.5 rounded-[2px] bg-slate-100" />
          <span className="h-2.5 w-2.5 rounded-[2px] bg-[#B8D0E4]" />
          <span className="h-2.5 w-2.5 rounded-[2px] bg-[#658CB2]" />
          <span className="h-2.5 w-2.5 rounded-[2px] bg-[#3B5B7E]" />
          <span className="ml-1 h-2.5 w-2.5 rounded-[2px] bg-[#D97706]" />
          <span>7日内</span>
        </div>
      </div>
    </div>
  );
}

/* ==================== 3. Project Milestone Timeline Chart ==================== */
interface TimelineProps {
  batches: Batch[];
  onOpenBatch: (batchId: string, path: string) => void;
}

export function BatchTimelineChart({ batches, onOpenBatch }: TimelineProps) {
  const timeline = useMemo(() => {
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

    const items = batches
      .map((b) => {
        const startDt = parseDateOnly(b.announcement_time) || parseDateOnly(b.created_at);
        const openDt = parseDateOnly(b.bid_opening_time);
        const anchor = openDt || startDt;
        if (!anchor) return null;
        const cd = getCountdownLabel(b.bid_opening_time);
        return {
          batch: b,
          startDt: startDt || openDt!,
          endDt: openDt || startDt!,
          hasSpan: Boolean(startDt && openDt && openDt.getTime() > startDt.getTime()),
          anchorMs: anchor.getTime(),
          countdown: cd,
        };
      })
      .filter((x): x is NonNullable<typeof x> => x !== null)
      .sort((a, b) => {
        // Prioritize in-progress batches with upcoming opening dates, then most recent
        const aDone = a.batch.status === '已完成' ? 1 : 0;
        const bDone = b.batch.status === '已完成' ? 1 : 0;
        if (aDone !== bDone) return aDone - bDone;
        return Math.abs(a.anchorMs - today.getTime()) - Math.abs(b.anchorMs - today.getTime());
      })
      .slice(0, 7);

    if (items.length === 0) {
      return { items: [], minMs: today.getTime(), maxMs: today.getTime(), todayPct: 50, ticks: [] as Array<{ pct: number; label: string }> };
    }

    const allMs = [today.getTime()];
    for (const it of items) {
      allMs.push(it.startDt.getTime(), it.endDt.getTime());
    }
    const rawMin = Math.min(...allMs);
    const rawMax = Math.max(...allMs);
    const pad = Math.max((rawMax - rawMin) * 0.08, 4 * 86400000);
    const minMs = rawMin - pad;
    const maxMs = rawMax + pad;
    const span = Math.max(maxMs - minMs, 86400000);

    const todayPct = ((today.getTime() - minMs) / span) * 100;
    const ticks = [0, 0.25, 0.5, 0.75, 1].map((r) => {
      const dt = new Date(minMs + span * r);
      return {
        pct: r * 100,
        label: `${dt.getMonth() + 1}/${dt.getDate()}`,
      };
    });

    return { items, minMs, maxMs, todayPct, ticks };
  }, [batches]);

  const spanMs = Math.max(timeline.maxMs - timeline.minMs, 86400000);

  return (
    <div className="flex h-full flex-col rounded-xl border border-slate-200 bg-white p-4">
      <div className="flex items-center justify-between gap-2">
        <div className="flex items-center gap-1.5">
          <Clock className="h-4 w-4 text-[#3B5B7E]" />
          <h2 className="text-sm font-semibold text-slate-800">近期批次时间图</h2>
          <span className="text-[11px] text-slate-400">公告 / 建档 → 开标节点</span>
        </div>
        <span className="inline-flex items-center gap-1 text-[10px] text-[#A94442]">
          <span className="inline-block h-2.5 w-0.5 bg-[#A94442]/70" />
          今日参考线
        </span>
      </div>

      {timeline.items.length === 0 ? (
        <div className="flex flex-1 flex-col items-center justify-center py-10 text-xs text-slate-400">
          <CalendarDays className="mb-1.5 h-5 w-5 text-slate-300" />
          <span>暂无包含公告或开标日期的批次</span>
        </div>
      ) : (
        <div className="mt-3 flex flex-1 flex-col justify-between">
          {/* Timeline Rows */}
          <div className="relative space-y-2.5 py-1">
            {/* Vertical Today Line */}
            <div
              className="pointer-events-none absolute top-0 bottom-0 z-10 w-px border-l border-dashed border-[#A94442]/60"
              style={{ left: `calc(34% + ${(100 - 34 - 16) * (timeline.todayPct / 100)}%)` }}
            />

            {timeline.items.map((it) => {
              const leftPct = Math.max(0, Math.min(100, ((it.startDt.getTime() - timeline.minMs) / spanMs) * 100));
              const rightPct = Math.max(leftPct, Math.min(100, ((it.endDt.getTime() - timeline.minMs) / spanMs) * 100));
              const widthPct = Math.max(rightPct - leftPct, 1.5);
              const color = STATUS_META[it.batch.status]?.color || '#3B5B7E';
              const targetPath =
                it.batch.status === '归档中' || it.batch.status === '已完成'
                  ? '/archive'
                  : it.batch.status === '开评标中'
                    ? '/evaluation'
                    : '/bidding';

              return (
                <button
                  key={it.batch.id}
                  type="button"
                  onClick={() => onOpenBatch(it.batch.id, targetPath)}
                  className="group flex w-full items-center gap-2 rounded-lg px-1.5 py-1 text-left hover:bg-slate-50 transition-colors"
                  title={`${it.batch.name}\n公告/建档：${formatDateKey(it.startDt)}\n开标日期：${it.batch.bid_opening_time || '待定'}`}
                >
                  {/* Batch Name (34%) */}
                  <div className="w-[34%] shrink-0 pr-1">
                    <div className="truncate text-xs font-medium text-slate-700 group-hover:text-[#3B5B7E]">
                      {it.batch.name}
                    </div>
                    <div className="truncate text-[10px] text-slate-400">
                      {it.batch.status} · {it.batch.code || it.batch.project_unit || '—'}
                    </div>
                  </div>

                  {/* Bar Track (50%) */}
                  <div className="relative h-5 flex-1 rounded bg-slate-100/80">
                    {it.hasSpan ? (
                      <div
                        className="absolute top-1.5 h-2 rounded-full transition-all group-hover:opacity-90"
                        style={{
                          left: `${leftPct}%`,
                          width: `${widthPct}%`,
                          backgroundColor: color,
                        }}
                      />
                    ) : (
                      <div
                        className="absolute top-1 h-3 w-3 -translate-x-1/2 rotate-45 rounded-[2px]"
                        style={{
                          left: `${rightPct}%`,
                          backgroundColor: color,
                        }}
                      />
                    )}
                    {/* Opening Milestone Dot */}
                    {it.batch.bid_opening_time && (
                      <div
                        className="absolute top-1 h-3 w-3 -translate-x-1/2 rounded-full border-2 border-white shadow-xs"
                        style={{
                          left: `${rightPct}%`,
                          backgroundColor: it.countdown?.urgent ? '#D97706' : color,
                        }}
                      />
                    )}
                  </div>

                  {/* Countdown (16%) */}
                  <div className="w-[16%] shrink-0 text-right">
                    {it.countdown ? (
                      <span
                        className={`inline-block rounded px-1.5 py-0.5 text-[10px] font-medium tabular-nums ${
                          it.countdown.urgent
                            ? 'bg-[#FEF9EC] text-[#9A6B26]'
                            : 'bg-slate-100 text-slate-500'
                        }`}
                      >
                        {it.countdown.text}
                      </span>
                    ) : (
                      <span className="text-[10px] text-slate-400">待定</span>
                    )}
                  </div>
                </button>
              );
            })}
          </div>

          {/* X-Axis Ticks */}
          <div className="mt-2 flex items-center border-t border-slate-100 pt-2">
            <div className="w-[34%] shrink-0 text-[10px] text-slate-400">
              展示近期待办 Top {timeline.items.length}
            </div>
            <div className="relative flex-1 h-3">
              {timeline.ticks.map((tk, idx) => (
                <span
                  key={idx}
                  className="absolute -translate-x-1/2 font-mono text-[10px] text-slate-400 tabular-nums"
                  style={{ left: `${tk.pct}%` }}
                >
                  {tk.label}
                </span>
              ))}
            </div>
            <div className="w-[16%] shrink-0" />
          </div>
        </div>
      )}
    </div>
  );
}

/* ==================== 4. SOP & Archive Progress Comparison Chart ==================== */
interface ProgressGapProps {
  batches: Batch[];
  onOpenBatch: (batchId: string, path: string) => void;
}

export function ProgressGapChart({ batches, onOpenBatch }: ProgressGapProps) {
  const focusBatches = useMemo(() => {
    return [...batches]
      .sort((a, b) => {
        const aDone = a.status === '已完成' ? 1 : 0;
        const bDone = b.status === '已完成' ? 1 : 0;
        if (aDone !== bDone) return aDone - bDone;
        const gapDiff = (b.stats?.archiveReqMissing || 0) - (a.stats?.archiveReqMissing || 0);
        if (gapDiff !== 0) return gapDiff;
        const aSop = a.stats?.sopTotal ? a.stats.sopDone / a.stats.sopTotal : 1;
        const bSop = b.stats?.sopTotal ? b.stats.sopDone / b.stats.sopTotal : 1;
        return aSop - bSop;
      })
      .slice(0, 6);
  }, [batches]);

  return (
    <div className="flex h-full flex-col rounded-xl border border-slate-200 bg-white p-4">
      <div className="flex items-center justify-between gap-2">
        <div className="flex items-center gap-1.5">
          <TrendingUp className="h-4 w-4 text-[#2F6F5E]" />
          <h2 className="text-sm font-semibold text-slate-800">SOP 与归档齐套率</h2>
        </div>
        <div className="flex items-center gap-2.5 text-[10px] text-slate-500">
          <span className="inline-flex items-center gap-1">
            <span className="h-2 w-2 rounded-full bg-[#3B5B7E]" />
            SOP
          </span>
          <span className="inline-flex items-center gap-1">
            <span className="h-2 w-2 rounded-full bg-[#2F6F5E]" />
            归档齐套
          </span>
        </div>
      </div>

      {focusBatches.length === 0 ? (
        <div className="flex flex-1 items-center justify-center py-10 text-xs text-slate-400">
          暂无批次数据
        </div>
      ) : (
        <div className="mt-3 flex-1 space-y-3">
          {focusBatches.map((b) => {
            const st = b.stats;
            const sopPct = st?.sopTotal ? Math.round((st.sopDone / st.sopTotal) * 100) : 0;
            const reqPct = st?.archiveReqTotal
              ? Math.round((st.archiveReqPresent / st.archiveReqTotal) * 100)
              : 0;
            const hasMissing = (st?.archiveReqMissing || 0) > 0;

            return (
              <button
                key={b.id}
                type="button"
                onClick={() => onOpenBatch(b.id, hasMissing ? '/archive' : '/evaluation')}
                className="group block w-full rounded-lg px-2 py-1.5 text-left hover:bg-slate-50 transition-colors"
              >
                <div className="flex items-center justify-between gap-2 text-xs">
                  <span className="truncate font-medium text-slate-700 group-hover:text-[#3B5B7E]">
                    {b.name}
                  </span>
                  {hasMissing ? (
                    <span className="inline-flex items-center gap-0.5 shrink-0 text-[10px] font-medium text-[#A94442] tabular-nums">
                      <AlertTriangle className="h-3 w-3" />
                      缺 {st.archiveReqMissing} 项
                    </span>
                  ) : (
                    <span className="inline-flex items-center gap-0.5 shrink-0 text-[10px] font-medium text-[#2F6F5E]">
                      <CheckCircle2 className="h-3 w-3" />
                      已齐
                    </span>
                  )}
                </div>

                {/* Dual Slim Bars */}
                <div className="mt-1.5 space-y-1">
                  <div className="flex items-center gap-2">
                    <div className="h-1.5 flex-1 overflow-hidden rounded-full bg-slate-100">
                      <div
                        className="h-full rounded-full bg-[#3B5B7E] transition-all"
                        style={{ width: `${sopPct}%` }}
                      />
                    </div>
                    <span className="w-8 text-right font-mono text-[10px] text-slate-500 tabular-nums">
                      {sopPct}%
                    </span>
                  </div>
                  <div className="flex items-center gap-2">
                    <div className="h-1.5 flex-1 overflow-hidden rounded-full bg-slate-100">
                      <div
                        className={`h-full rounded-full transition-all ${
                          hasMissing ? 'bg-[#A94442]/75' : 'bg-[#2F6F5E]'
                        }`}
                        style={{ width: `${reqPct}%` }}
                      />
                    </div>
                    <span className="w-8 text-right font-mono text-[10px] text-slate-500 tabular-nums">
                      {reqPct}%
                    </span>
                  </div>
                </div>
              </button>
            );
          })}
        </div>
      )}

      {batches.length > 6 && (
        <div className="mt-2 flex items-center justify-between border-t border-slate-100 pt-2 text-[10px] text-slate-400">
          <span>按归档缺口与待办优先展示 Top 6</span>
          <span className="inline-flex items-center gap-0.5 text-[#3B5B7E]">
            点击直达处理 <ArrowRight className="h-2.5 w-2.5" />
          </span>
        </div>
      )}
    </div>
  );
}
