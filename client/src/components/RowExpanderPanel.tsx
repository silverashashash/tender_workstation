import { useState } from 'react';
import * as XLSX from 'xlsx';
import {
  AlertTriangle,
  CheckCircle2,
  ChevronDown,
  ChevronRight,
  Download,
  FileSpreadsheet,
  Info,
  Play,
  RotateCcw,
  SplitSquareVertical,
  Upload,
} from 'lucide-react';
import {
  extractWorkbookHeaders,
  runRowExpander,
  type RowExpandMode,
  type RowExpandResult,
} from '../utils/rowExpander';

interface ModeOption {
  key: RowExpandMode;
  label: string;
  desc: string;
  badge?: string;
}

const PRIMARY_MODES: ModeOption[] = [
  {
    key: 'two_level',
    label: '两级拆分',
    badge: '常用',
    desc: '把「分标:包1,包2」拆成独立分标与包号两列并展开为多行，方便筛选与匹配中标人',
  },
  {
    key: 'contacts',
    label: '联系人展开并检查',
    desc: '展开「分标包号」并按 (标段,包号,供应商) 校验联系方式冲突，输出 4 表核查报告',
  },
  {
    key: 'simple',
    label: '简单展开',
    desc: '通用单列拆分：按分隔符把单格多值拆成多行，值保留在原列',
  },
  {
    key: 'contacts_existing',
    label: '已展开联系人检查',
    desc: '对已含分标/包号两列的联系人表直接做冲突与完整性检查，输出 4 表核查报告',
  },
];

const ADVANCED_MODES: ModeOption[] = [
  {
    key: 'merge',
    label: '一级合并（支持分组汇总）',
    desc: '可指定分组列（如 投标人名称,分标）按自然顺序去重拼接目标列；留空则按其余全部列完全一致合并',
  },
  {
    key: 'merge_two',
    label: '两级合并（两级分类拼回单格）',
    desc: '按分组列把「一级分类列 + 二级明细列」重组为「分类1:项1,项2;分类2:项3」并插入新列',
  },
];

export function RowExpanderPanel() {
  const [fileName, setFileName] = useState('');
  const [fileSize, setFileSize] = useState(0);
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [headers, setHeaders] = useState<string[]>([]);
  const [sheetMeta, setSheetMeta] = useState<{ sheetName: string; rowCount: number } | null>(null);

  const [mode, setMode] = useState<RowExpandMode>('two_level');
  const [showAdvancedModes, setShowAdvancedModes] = useState(false);
  const [headerPickTarget, setHeaderPickTarget] = useState<'colName' | 'groupCols' | 'prefixName' | 'itemName'>('colName');
  const [colName, setColName] = useState('分标包号');
  const [sep, setSep] = useState(',');
  const [prefixSep, setPrefixSep] = useState(':');
  const [prefixName, setPrefixName] = useState('分标');
  const [itemName, setItemName] = useState('包号');
  const [groupSep, setGroupSep] = useState(';');
  const [groupCols, setGroupCols] = useState('');
  const [newColPos, setNewColPos] = useState('');
  const [supplierCol, setSupplierCol] = useState('');

  const [reading, setReading] = useState(false);
  const [errorMsg, setErrorMsg] = useState('');
  const [result, setResult] = useState<RowExpandResult | null>(null);
  const [activePreviewIdx, setActivePreviewIdx] = useState(0);

  async function handleFileSelect(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;

    setWorkbook(null);
    setHeaders([]);
    setSheetMeta(null);
    setFileName('');
    setFileSize(0);
    setErrorMsg('');
    setResult(null);

    if (!/\.(xlsx?|xlsm)$/i.test(file.name)) {
      setErrorMsg('请选择 Excel 文件（支持 .xlsx / .xlsm / .xls）');
      return;
    }

    setReading(true);
    try {
      const buf = await file.arrayBuffer();
      const parsed = XLSX.read(buf, { type: 'array', cellFormula: true, cellNF: true });
      const meta = extractWorkbookHeaders(parsed);
      setWorkbook(parsed);
      setFileName(file.name);
      setFileSize(file.size);
      setHeaders(meta.headers);
      setSheetMeta({ sheetName: meta.sheetName, rowCount: meta.rowCount });
    } catch (err) {
      setErrorMsg(`无法读取文件：${(err as Error).message}`);
    } finally {
      setReading(false);
    }
  }

  function handleReset() {
    setWorkbook(null);
    setHeaders([]);
    setSheetMeta(null);
    setFileName('');
    setFileSize(0);
    setErrorMsg('');
    setResult(null);
    setActivePreviewIdx(0);
  }

  function handleRun() {
    if (!workbook) return;
    setErrorMsg('');
    setResult(null);
    setActivePreviewIdx(0);
    try {
      const res = runRowExpander(workbook, {
        mode,
        fileName,
        colName,
        sep,
        prefixSep,
        prefixName,
        itemName,
        groupSep,
        groupCols,
        newColPos,
        supplierCol,
      });
      setResult(res);
    } catch (err) {
      setErrorMsg((err as Error).message);
    }
  }

  function handleDownload() {
    if (!result) return;
    try {
      XLSX.writeFile(result.workbook, result.outputFileName, { bookType: 'xlsx', compression: true });
    } catch (err) {
      setErrorMsg(`下载失败：${(err as Error).message}`);
    }
  }

  const showTwoLevel =
    mode === 'two_level' || mode === 'merge_two' || mode === 'contacts' || mode === 'contacts_existing';
  const showMergeTwo = mode === 'merge_two';
  const showContacts = mode === 'contacts' || mode === 'contacts_existing';

  const curPreview = result?.sheetPreviews[Math.min(activePreviewIdx, (result?.sheetPreviews.length || 1) - 1)];

  return (
    <div className="space-y-4">
      {/* Intro */}
      <div className="rounded-xl border border-slate-200 bg-white px-4 py-3 flex flex-wrap items-center justify-between gap-2">
        <div className="flex items-start gap-2 text-xs text-slate-600">
          <Info className="h-4 w-4 shrink-0 text-[#3B5B7E] mt-0.5" />
          <span>
            对 Excel 表格执行<strong>单行按分隔符展开</strong>、<strong>分标/包号两级拆分</strong>、<strong>一级/两级逆向合并</strong>，以及<strong>供应商联系人展开与冲突核查</strong>。全程浏览器本地处理，原文件不会被修改。
          </span>
        </div>
        {workbook && (
          <button
            type="button"
            onClick={handleReset}
            className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-slate-50 px-2.5 py-1 text-xs font-medium text-slate-600 hover:border-[#A94442]/40 hover:text-[#A94442]"
          >
            <RotateCcw className="h-3.5 w-3.5" />
            <span>清空重选</span>
          </button>
        )}
      </div>

      {/* File Upload + Detected Headers */}
      <div className="rounded-xl border border-slate-200 bg-white p-4 space-y-3">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div className="flex items-center gap-2">
            <FileSpreadsheet className="h-4 w-4 text-[#3B5B7E]" />
            <h2 className="text-sm font-semibold text-slate-800">1. 选择待处理的 Excel 文件</h2>
            {sheetMeta && (
              <span className="rounded-full bg-[#EBF2F7] px-2 py-0.5 text-[11px] font-medium text-[#3B5B7E] tabular-nums">
                {sheetMeta.sheetName} · {sheetMeta.rowCount} 行数据 · {headers.length} 列
              </span>
            )}
          </div>
          <label className="inline-flex cursor-pointer items-center gap-1.5 rounded-lg border border-[#3B5B7E]/35 bg-[#EBF2F7]/50 px-3 py-1.5 text-xs font-medium text-[#3B5B7E] hover:bg-[#EBF2F7] transition-colors">
            <Upload className="h-3.5 w-3.5" />
            <span>{reading ? '读取中...' : fileName ? '更换文件' : '选择 Excel (.xlsx / .xls)'}</span>
            <input
              type="file"
              accept=".xlsx,.xlsm,.xls"
              onChange={handleFileSelect}
              className="hidden"
              disabled={reading}
            />
          </label>
        </div>

        {fileName && (
          <div className="space-y-2">
            <div className="text-xs text-slate-600">
              当前文件：<strong className="text-slate-800">{fileName}</strong>
              <span className="ml-2 text-[11px] text-slate-400 tabular-nums">
                ({(fileSize / 1024).toFixed(1)} KB)
              </span>
            </div>
            {headers.length > 0 && (() => {
              const effectiveTarget: 'colName' | 'groupCols' | 'prefixName' | 'itemName' =
                mode === 'merge_two'
                  ? headerPickTarget === 'prefixName' || headerPickTarget === 'itemName'
                    ? headerPickTarget
                    : 'groupCols'
                  : mode === 'merge'
                    ? headerPickTarget === 'groupCols'
                      ? 'groupCols'
                      : 'colName'
                    : 'colName';
              const selectedGroupList = groupCols
                .replace(/，/g, ',')
                .split(',')
                .map((s) => s.trim())
                .filter(Boolean);
              return (
                <div className="flex flex-wrap items-center gap-1.5 text-[11px]">
                  <span className="text-slate-400">识别到表头（点击填入</span>
                  {mode === 'merge' ? (
                    <span className="inline-flex items-center rounded-md border border-slate-200 bg-slate-100 p-0.5">
                      <button
                        type="button"
                        onClick={() => setHeaderPickTarget('colName')}
                        className={`rounded px-1.5 py-0.5 text-[10px] font-medium transition-colors ${
                          effectiveTarget === 'colName'
                            ? 'bg-[#3B5B7E] text-white'
                            : 'text-slate-600 hover:text-slate-800'
                        }`}
                      >
                        目标列名
                      </button>
                      <button
                        type="button"
                        onClick={() => setHeaderPickTarget('groupCols')}
                        className={`rounded px-1.5 py-0.5 text-[10px] font-medium transition-colors ${
                          effectiveTarget === 'groupCols'
                            ? 'bg-[#2F6F5E] text-white'
                            : 'text-slate-600 hover:text-slate-800'
                        }`}
                      >
                        分组列
                      </button>
                    </span>
                  ) : mode === 'merge_two' ? (
                    <span className="inline-flex items-center rounded-md border border-slate-200 bg-slate-100 p-0.5">
                      <button
                        type="button"
                        onClick={() => setHeaderPickTarget('groupCols')}
                        className={`rounded px-1.5 py-0.5 text-[10px] font-medium transition-colors ${
                          effectiveTarget === 'groupCols'
                            ? 'bg-[#2F6F5E] text-white'
                            : 'text-slate-600 hover:text-slate-800'
                        }`}
                      >
                        分组列
                      </button>
                      <button
                        type="button"
                        onClick={() => setHeaderPickTarget('prefixName')}
                        className={`rounded px-1.5 py-0.5 text-[10px] font-medium transition-colors ${
                          effectiveTarget === 'prefixName'
                            ? 'bg-[#3B5B7E] text-white'
                            : 'text-slate-600 hover:text-slate-800'
                        }`}
                      >
                        一级分类列
                      </button>
                      <button
                        type="button"
                        onClick={() => setHeaderPickTarget('itemName')}
                        className={`rounded px-1.5 py-0.5 text-[10px] font-medium transition-colors ${
                          effectiveTarget === 'itemName'
                            ? 'bg-[#9A6B26] text-white'
                            : 'text-slate-600 hover:text-slate-800'
                        }`}
                      >
                        二级明细列
                      </button>
                    </span>
                  ) : (
                    <span className="text-slate-500 font-medium">目标列名</span>
                  )}
                  <span className="text-slate-400">）：</span>
                  {headers.map((h) => {
                    const isColMatch = mode !== 'merge_two' && mode !== 'contacts_existing' && colName.trim() === h;
                    const isGroupMatch = (mode === 'merge' || mode === 'merge_two') && selectedGroupList.includes(h);
                    const isPrefixMatch = mode === 'merge_two' && prefixName.trim() === h;
                    const isItemMatch = mode === 'merge_two' && itemName.trim() === h;
                    return (
                      <button
                        key={h}
                        type="button"
                        onClick={() => {
                          setResult(null);
                          if (effectiveTarget === 'groupCols') {
                            setGroupCols((prev) => {
                              const list = prev
                                .replace(/，/g, ',')
                                .split(',')
                                .map((s) => s.trim())
                                .filter(Boolean);
                              if (list.includes(h)) {
                                return list.filter((x) => x !== h).join(',');
                              }
                              return [...list, h].join(',');
                            });
                          } else if (effectiveTarget === 'prefixName') {
                            setPrefixName(h);
                          } else if (effectiveTarget === 'itemName') {
                            setItemName(h);
                          } else {
                            setColName(h);
                          }
                        }}
                        className={`rounded border px-2 py-0.5 transition-colors ${
                          isColMatch || isPrefixMatch
                            ? 'border-[#3B5B7E] bg-[#EBF2F7] text-[#3B5B7E] font-medium'
                            : isGroupMatch
                              ? 'border-[#2F6F5E]/60 bg-[#ECFDF5] text-[#2F6F5E] font-medium'
                              : isItemMatch
                                ? 'border-[#9A6B26]/50 bg-[#FEF9EC] text-[#9A6B26] font-medium'
                                : 'border-slate-200 bg-slate-50 text-slate-600 hover:border-[#3B5B7E] hover:text-[#3B5B7E]'
                        }`}
                      >
                        {h}
                      </button>
                    );
                  })}
                </div>
              );
            })()}
          </div>
        )}
      </div>

      {/* Mode Selection + Parameters */}
      <div className="rounded-xl border border-slate-200 bg-white p-4 space-y-4">
        <div className="space-y-2.5">
          <h2 className="text-sm font-semibold text-slate-800">2. 选择操作模式与参数</h2>
          <div className="grid grid-cols-1 gap-2 sm:grid-cols-2">
            {PRIMARY_MODES.map((m) => {
              const active = mode === m.key;
              return (
                <button
                  key={m.key}
                  type="button"
                  onClick={() => {
                    setMode(m.key);
                    setResult(null);
                    setErrorMsg('');
                  }}
                  className={`rounded-lg border p-2.5 text-left transition-colors ${
                    active
                      ? 'border-[#3B5B7E] bg-[#EBF2F7]/50'
                      : 'border-slate-200 bg-white hover:bg-slate-50'
                  }`}
                >
                  <div className="flex items-center gap-1.5">
                    <span className={`text-xs font-semibold ${active ? 'text-[#3B5B7E]' : 'text-slate-800'}`}>
                      {m.label}
                    </span>
                    {m.badge && (
                      <span className="rounded bg-[#EBF2F7] px-1.5 py-0.5 text-[10px] font-medium text-[#3B5B7E]">
                        {m.badge}
                      </span>
                    )}
                  </div>
                  <div className="mt-0.5 text-[11px] text-slate-500 leading-relaxed">{m.desc}</div>
                </button>
              );
            })}
          </div>

          <div>
            <button
              type="button"
              onClick={() => setShowAdvancedModes((prev) => !prev)}
              className="inline-flex items-center gap-1 text-[11px] font-medium text-slate-500 hover:text-[#3B5B7E] transition-colors"
            >
              {showAdvancedModes || mode === 'merge' || mode === 'merge_two' ? (
                <ChevronDown className="h-3.5 w-3.5" />
              ) : (
                <ChevronRight className="h-3.5 w-3.5" />
              )}
              <span>高级 / 逆向合并模式（一级合并、两级合并）</span>
            </button>

            {(showAdvancedModes || mode === 'merge' || mode === 'merge_two') && (
              <div className="mt-2 grid grid-cols-1 gap-2 sm:grid-cols-2">
                {ADVANCED_MODES.map((m) => {
                  const active = mode === m.key;
                  return (
                    <button
                      key={m.key}
                      type="button"
                      onClick={() => {
                        setMode(m.key);
                        setResult(null);
                        setErrorMsg('');
                      }}
                      className={`rounded-lg border p-2.5 text-left transition-colors ${
                        active
                          ? 'border-[#3B5B7E] bg-[#EBF2F7]/50'
                          : 'border-slate-200 bg-slate-50/50 hover:bg-slate-50'
                      }`}
                    >
                      <div className={`text-xs font-semibold ${active ? 'text-[#3B5B7E]' : 'text-slate-700'}`}>
                        {m.label}
                      </div>
                      <div className="mt-0.5 text-[11px] text-slate-500 leading-relaxed">{m.desc}</div>
                    </button>
                  );
                })}
              </div>
            )}
          </div>
        </div>

        {showMergeTwo && (
          <div className="rounded-lg border border-[#3B5B7E]/20 bg-[#EBF2F7]/40 px-3 py-2 text-[11px] text-slate-600 leading-relaxed">
            <strong className="text-[#3B5B7E]">示例：</strong>
            按「分组列」（如公司/供应商）归并多行，把同组内的「一级分类列」（如地区/分标）与「二级明细列」（如门店/包号）拼成单格
            <code className="mx-1 rounded bg-white px-1.5 py-0.5 text-[#3B5B7E] border border-slate-200">北京:朝阳,海淀;上海:浦东</code>
            （招标场景即 <code className="rounded bg-white px-1.5 py-0.5 text-[#3B5B7E] border border-slate-200">分标1:包1,包2;分标2:包3</code>）并插入为新列。
          </div>
        )}

        {/* Dynamic Parameter Inputs */}
        <div className="grid grid-cols-1 gap-3 border-t border-slate-100 pt-3 sm:grid-cols-2 lg:grid-cols-4">
          {mode !== 'contacts_existing' && (
            <div>
              <label className="block text-[11px] font-medium text-slate-600 mb-1">
                {mode === 'merge_two' ? '输出汇总列名' : '目标列名'}
              </label>
              <input
                value={colName}
                onFocus={() => setHeaderPickTarget('colName')}
                onChange={(e) => {
                  setColName(e.target.value);
                  setResult(null);
                }}
                placeholder={mode === 'merge_two' ? '例如：分类汇总 / 分标包号' : '分标包号'}
                className="w-full rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
              />
            </div>
          )}

          <div>
            <label className="block text-[11px] font-medium text-slate-600 mb-1">
              {mode === 'merge' ? '连接分隔符' : mode === 'merge_two' ? '明细项分隔符' : '列表分隔符'}
            </label>
            <input
              value={sep}
              onChange={(e) => {
                setSep(e.target.value);
                setResult(null);
              }}
              placeholder=","
              className="w-full rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-mono focus:border-[#3B5B7E] focus:outline-none"
            />
          </div>

          {mode === 'merge' && (
            <div className="sm:col-span-2">
              <div className="flex items-center justify-between mb-1">
                <label className="block text-[11px] font-medium text-slate-600">
                  分组列（可选；留空=按其余全部列相同合并，填写如 投标人名称,分标 则输出分组汇总表）
                </label>
                {groupCols && (
                  <button
                    type="button"
                    onClick={() => {
                      setGroupCols('');
                      setResult(null);
                    }}
                    className="text-[10px] text-slate-400 hover:text-[#A94442]"
                  >
                    清空
                  </button>
                )}
              </div>
              <input
                value={groupCols}
                onFocus={() => setHeaderPickTarget('groupCols')}
                onChange={(e) => {
                  setGroupCols(e.target.value);
                  setResult(null);
                }}
                placeholder="例如：投标人名称,分标（聚焦后点击上方表头快速添加/取消）"
                className="w-full rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
              />
            </div>
          )}

          {showTwoLevel && (
            <>
              <div>
                <label className="block text-[11px] font-medium text-slate-600 mb-1">
                  {mode === 'merge_two' ? '一二级连接符' : '前缀分隔符'}
                </label>
                <input
                  value={prefixSep}
                  onChange={(e) => {
                    setPrefixSep(e.target.value);
                    setResult(null);
                  }}
                  placeholder=":"
                  className="w-full rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-mono focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-[11px] font-medium text-slate-600 mb-1">
                  {mode === 'two_level'
                    ? '拆分后前缀列名'
                    : mode === 'merge_two'
                      ? '一级分类列（前缀列）'
                      : '分标列名'}
                </label>
                <input
                  value={prefixName}
                  onFocus={() => mode === 'merge_two' && setHeaderPickTarget('prefixName')}
                  onChange={(e) => {
                    setPrefixName(e.target.value);
                    setResult(null);
                  }}
                  placeholder={mode === 'merge_two' ? '如：地区 / 分标' : '分标'}
                  className="w-full rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-[11px] font-medium text-slate-600 mb-1">
                  {mode === 'two_level'
                    ? '拆分后列表列名'
                    : mode === 'merge_two'
                      ? '二级明细列（列表列）'
                      : '包号列名'}
                </label>
                <input
                  value={itemName}
                  onFocus={() => mode === 'merge_two' && setHeaderPickTarget('itemName')}
                  onChange={(e) => {
                    setItemName(e.target.value);
                    setResult(null);
                  }}
                  placeholder={mode === 'merge_two' ? '如：门店 / 包号' : '包号'}
                  className="w-full rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>
            </>
          )}

          {showMergeTwo && (
            <>
              <div className="sm:col-span-2">
                <div className="flex items-center justify-between mb-1">
                  <label className="block text-[11px] font-medium text-slate-600">
                    分组列 *（按哪些列合并为一行，逗号分隔）
                  </label>
                  {groupCols && (
                    <button
                      type="button"
                      onClick={() => {
                        setGroupCols('');
                        setResult(null);
                      }}
                      className="text-[10px] text-slate-400 hover:text-[#A94442]"
                    >
                      清空
                    </button>
                  )}
                </div>
                <input
                  value={groupCols}
                  onFocus={() => setHeaderPickTarget('groupCols')}
                  onChange={(e) => {
                    setGroupCols(e.target.value);
                    setResult(null);
                  }}
                  placeholder="例如：公司名称 / 投标人名称（聚焦后点击上方表头快速添加）"
                  className="w-full rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-[11px] font-medium text-slate-600 mb-1">
                  新列位置（留空放最后）
                </label>
                <input
                  value={newColPos}
                  onChange={(e) => {
                    setNewColPos(e.target.value);
                    setResult(null);
                  }}
                  placeholder="列名 / 列号 / 列字母"
                  className="w-full rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-[11px] font-medium text-slate-600 mb-1">不同分类分隔符</label>
                <input
                  value={groupSep}
                  onChange={(e) => {
                    setGroupSep(e.target.value);
                    setResult(null);
                  }}
                  placeholder=";"
                  className="w-full rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-mono focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>
            </>
          )}

          {showContacts && (
            <div className="sm:col-span-2">
              <label className="block text-[11px] font-medium text-slate-600 mb-1">
                供应商列名（留空按常见表头自动识别；有歧义时手动指定）
              </label>
              <input
                value={supplierCol}
                onChange={(e) => {
                  setSupplierCol(e.target.value);
                  setResult(null);
                }}
                placeholder="例如：供应商名称"
                className="w-full rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
              />
            </div>
          )}
        </div>

        <div className="flex flex-wrap items-center gap-2.5 border-t border-slate-100 pt-3">
          <button
            type="button"
            onClick={handleRun}
            disabled={!workbook}
            className="inline-flex items-center gap-1.5 rounded-lg bg-[#3B5B7E] px-4 py-2 text-xs font-medium text-white hover:bg-[#2E4866] disabled:opacity-45 transition-colors"
          >
            <Play className="h-3.5 w-3.5" />
            <span>开始处理</span>
          </button>

          <button
            type="button"
            onClick={handleDownload}
            disabled={!result}
            className="inline-flex items-center gap-1.5 rounded-lg border border-[#2F6F5E]/35 bg-[#ECFDF5] px-4 py-2 text-xs font-medium text-[#2F6F5E] hover:bg-[#D1FAE5] disabled:opacity-45 transition-colors"
          >
            <Download className="h-3.5 w-3.5" />
            <span>{result ? `下载 ${result.outputFileName}` : '下载结果 (.xlsx)'}</span>
          </button>
        </div>
      </div>

      {/* Error Banner */}
      {errorMsg && (
        <div className="flex items-start gap-2 rounded-xl border border-[#A94442]/30 bg-[#FEF2F2] px-4 py-3 text-xs text-[#A94442]">
          <AlertTriangle className="h-4 w-4 shrink-0 mt-0.5" />
          <span className="whitespace-pre-line">{errorMsg}</span>
        </div>
      )}

      {/* Results */}
      {result && (
        <div className="space-y-4">
          <div className="rounded-xl border border-[#2F6F5E]/30 bg-[#ECFDF5]/70 px-4 py-3 space-y-1.5 text-xs text-[#2F6F5E]">
            <div className="flex flex-wrap items-center gap-2 font-semibold">
              <CheckCircle2 className="h-4 w-4 shrink-0" />
              <span>
                处理完成：{result.srcRows} 行 → {result.outRows} 行（输出文件：{result.outputFileName}）
              </span>
            </div>
            {result.contactsStats && (
              <div className="flex flex-wrap gap-2 pl-6 text-[11px] tabular-nums">
                <span className="rounded bg-white px-2 py-0.5 text-[#2F6F5E] border border-[#2F6F5E]/20">
                  可匹配键：{result.contactsStats.safe_keys}
                </span>
                <span className="rounded bg-white px-2 py-0.5 text-[#A94442] border border-[#A94442]/20">
                  冲突键：{result.contactsStats.conflict_keys}
                </span>
                <span className="rounded bg-white px-2 py-0.5 text-[#9A6B26] border border-[#9A6B26]/20">
                  待核实行：{result.contactsStats.review_rows}
                </span>
                <span className="rounded bg-white px-2 py-0.5 text-slate-600 border border-slate-200">
                  一致重复：{result.contactsStats.duplicate_rows}
                </span>
              </div>
            )}
          </div>

          {/* Preview Table */}
          {curPreview && (
            <div className="rounded-xl border border-slate-200 bg-white overflow-hidden">
              <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 bg-slate-50/70 px-4 py-2.5">
                <div className="flex flex-wrap items-center gap-1.5">
                  <SplitSquareVertical className="h-4 w-4 text-[#3B5B7E]" />
                  <span className="text-xs font-semibold text-slate-700 mr-1">结果预览：</span>
                  {result.sheetPreviews.map((sp, idx) => (
                    <button
                      key={sp.name}
                      type="button"
                      onClick={() => setActivePreviewIdx(idx)}
                      className={`rounded-md px-2.5 py-1 text-[11px] font-medium tabular-nums transition-colors ${
                        activePreviewIdx === idx
                          ? 'bg-[#3B5B7E] text-white'
                          : 'bg-white border border-slate-200 text-slate-600 hover:bg-slate-100'
                      }`}
                    >
                      {sp.name} ({sp.totalRows})
                    </button>
                  ))}
                </div>
                <span className="text-[11px] text-slate-400">
                  仅预览前 30 行，下载包含全部 {curPreview.totalRows} 行
                </span>
              </div>

              <div className="max-h-96 overflow-auto">
                <table className="w-full border-collapse text-left text-xs">
                  <thead className="sticky top-0 bg-slate-100 text-slate-600 border-b border-slate-200">
                    <tr>
                      {curPreview.headers.map((h, i) => (
                        <th key={i} className="px-3 py-2 font-medium whitespace-nowrap">
                          {h}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {curPreview.rows.length === 0 ? (
                      <tr>
                        <td
                          colSpan={Math.max(1, curPreview.headers.length)}
                          className="py-8 text-center text-xs text-slate-400"
                        >
                          该工作表暂无数据行
                        </td>
                      </tr>
                    ) : (
                      curPreview.rows.map((row, rIdx) => (
                        <tr key={rIdx} className="hover:bg-slate-50/70">
                          {row.map((val, cIdx) => (
                            <td
                              key={cIdx}
                              className="px-3 py-1.5 max-w-[260px] truncate text-slate-700"
                              title={val}
                            >
                              {val}
                            </td>
                          ))}
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* Logs */}
          <div className="rounded-xl border border-slate-200 bg-white overflow-hidden">
            <div className="border-b border-slate-200 bg-slate-50/70 px-4 py-2 text-xs font-semibold text-slate-600">
              处理日志
            </div>
            <pre className="p-3.5 text-[11px] font-mono leading-relaxed text-slate-700 bg-slate-50/40 whitespace-pre-wrap">
              {result.logs.join('\n')}
            </pre>
          </div>
        </div>
      )}
    </div>
  );
}
