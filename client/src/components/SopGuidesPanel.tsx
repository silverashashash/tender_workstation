import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  BookOpen,
  Download,
  Edit3,
  ExternalLink,
  FileCode2,
  FileText,
  ImagePlus,
  Plus,
  Save,
  Search,
  Trash2,
  Upload,
  X,
  ZoomIn,
} from 'lucide-react';
import { api } from '../api/client';
import type { SopGuide, SopGuideImage } from '../types';

const GUIDE_STAGES = ['开标前准备', '内网开标', '评审过程', '评审结束与资料带出', '系统通用操作'];

/* ==================== Lightweight Safe Markdown Renderer ==================== */
function renderInline(text: string, onZoomImage: (src: string, alt: string) => void): React.ReactNode[] {
  const nodes: React.ReactNode[] = [];
  // Matches: ![alt](url) | [text](url) | `code` | **bold** | *italic*
  const tokenRe = /!\[([^\]]*)\]\(([^)\s]+)\)|\[([^\]]+)\]\((https?:\/\/[^)\s]+)\)|`([^`]+)`|\*\*([^*]+)\*\*|\*([^*]+)\*/g;
  let lastIdx = 0;
  let match: RegExpExecArray | null;
  let key = 0;

  while ((match = tokenRe.exec(text)) !== null) {
    if (match.index > lastIdx) {
      nodes.push(text.slice(lastIdx, match.index));
    }
    if (match[1] !== undefined && match[2]) {
      const alt = match[1] || '操作截图';
      const src = match[2];
      nodes.push(
        <span key={`img-${key++}`} className="my-2 block">
          <img
            src={src}
            alt={alt}
            onClick={() => onZoomImage(src, alt)}
            className="max-h-96 cursor-zoom-in rounded-lg border border-slate-200 bg-slate-50 object-contain shadow-xs transition-opacity hover:opacity-95"
          />
          {alt && alt !== '操作截图' && (
            <span className="mt-1 block text-[11px] text-slate-400">{alt}</span>
          )}
        </span>
      );
    } else if (match[3] && match[4]) {
      nodes.push(
        <a
          key={`lnk-${key++}`}
          href={match[4]}
          target="_blank"
          rel="noreferrer"
          className="text-[#3B5B7E] underline hover:text-[#2E4866]"
        >
          {match[3]}
        </a>
      );
    } else if (match[5]) {
      nodes.push(
        <code
          key={`code-${key++}`}
          className="rounded bg-slate-100 px-1.5 py-0.5 font-mono text-[11px] text-[#3B5B7E]"
        >
          {match[5]}
        </code>
      );
    } else if (match[6]) {
      nodes.push(
        <strong key={`b-${key++}`} className="font-semibold text-slate-900">
          {match[6]}
        </strong>
      );
    } else if (match[7]) {
      nodes.push(
        <em key={`i-${key++}`} className="italic text-slate-700">
          {match[7]}
        </em>
      );
    }
    lastIdx = match.index + match[0].length;
  }
  if (lastIdx < text.length) {
    nodes.push(text.slice(lastIdx));
  }
  return nodes;
}

function MarkdownView({
  content,
  onZoomImage,
}: {
  content: string;
  onZoomImage: (src: string, alt: string) => void;
}) {
  const blocks = useMemo(() => {
    const lines = content.replace(/\r\n/g, '\n').split('\n');
    const out: React.ReactNode[] = [];
    let i = 0;
    let key = 0;

    while (i < lines.length) {
      const line = lines[i];

      // Blank line
      if (!line.trim()) {
        i++;
        continue;
      }

      // Fenced code block
      if (line.trim().startsWith('```')) {
        const codeLines: string[] = [];
        i++;
        while (i < lines.length && !lines[i].trim().startsWith('```')) {
          codeLines.push(lines[i]);
          i++;
        }
        i++;
        out.push(
          <pre
            key={`pre-${key++}`}
            className="my-2 overflow-x-auto rounded-lg bg-slate-900 p-3 font-mono text-[11px] leading-relaxed text-slate-100"
          >
            <code>{codeLines.join('\n')}</code>
          </pre>
        );
        continue;
      }

      // Headings (H1–H6)
      const hMatch = line.match(/^\s*(#{1,6})\s+(.+)$/);
      if (hMatch) {
        const level = hMatch[1].length;
        const text = hMatch[2].trim();
        const cls =
          level === 1
            ? 'mt-4 mb-2 border-b border-slate-100 pb-1.5 text-base font-bold text-slate-900 first:mt-0'
            : level === 2
              ? 'mt-3.5 mb-1.5 text-sm font-semibold text-[#3B5B7E] first:mt-0'
              : 'mt-2.5 mb-1 text-xs font-semibold text-slate-800';
        out.push(
          <div key={`h-${key++}`} className={cls}>
            {renderInline(text, onZoomImage)}
          </div>
        );
        i++;
        continue;
      }

      // Horizontal rule
      if (/^\s*[-*_]{3,}\s*$/.test(line)) {
        out.push(<hr key={`hr-${key++}`} className="my-3 border-slate-200" />);
        i++;
        continue;
      }

      // Blockquote
      if (line.trim().startsWith('>')) {
        const qLines: string[] = [];
        while (i < lines.length && lines[i].trim().startsWith('>')) {
          qLines.push(lines[i].replace(/^\s*>\s?/, ''));
          i++;
        }
        out.push(
          <blockquote
            key={`bq-${key++}`}
            className="my-2 border-l-3 border-[#3B5B7E]/50 bg-[#EBF2F7]/40 px-3 py-2 text-xs text-slate-700 leading-relaxed"
          >
            {qLines.map((ql, idx) => (
              <div key={idx}>{renderInline(ql, onZoomImage)}</div>
            ))}
          </blockquote>
        );
        continue;
      }

      // Markdown table
      if (
        line.includes('|') &&
        i + 1 < lines.length &&
        /^\s*\|?[\s:-]+\|[\s|:-]+\s*$/.test(lines[i + 1])
      ) {
        const splitRow = (r: string) =>
          r
            .trim()
            .replace(/^\|/, '')
            .replace(/\|$/, '')
            .split('|')
            .map((c) => c.trim());
        const headers = splitRow(line);
        i += 2;
        const bodyRows: string[][] = [];
        while (i < lines.length && lines[i].includes('|') && lines[i].trim() !== '') {
          bodyRows.push(splitRow(lines[i]));
          i++;
        }
        out.push(
          <div key={`tbl-${key++}`} className="my-2.5 overflow-x-auto rounded-lg border border-slate-200">
            <table className="w-full border-collapse text-left text-xs">
              <thead>
                <tr className="border-b border-slate-200 bg-slate-50 text-slate-600">
                  {headers.map((h, idx) => (
                    <th key={idx} className="px-2.5 py-1.5 font-medium">
                      {renderInline(h, onZoomImage)}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {bodyRows.map((row, rIdx) => (
                  <tr key={rIdx} className="hover:bg-slate-50/60">
                    {headers.map((_, cIdx) => (
                      <td key={cIdx} className="px-2.5 py-1.5 text-slate-700">
                        {renderInline(row[cIdx] || '', onZoomImage)}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        );
        continue;
      }

      // Unordered list
      if (/^\s*[-*+]\s+/.test(line)) {
        const items: string[] = [];
        while (i < lines.length && /^\s*[-*+]\s+/.test(lines[i])) {
          items.push(lines[i].replace(/^\s*[-*+]\s+/, ''));
          i++;
        }
        out.push(
          <ul key={`ul-${key++}`} className="my-1.5 list-disc space-y-1 pl-5 text-xs text-slate-700 leading-relaxed">
            {items.map((it, idx) => (
              <li key={idx}>{renderInline(it, onZoomImage)}</li>
            ))}
          </ul>
        );
        continue;
      }

      // Ordered list
      if (/^\s*\d+[.)、]\s+/.test(line)) {
        const items: string[] = [];
        while (i < lines.length && /^\s*\d+[.)、]\s+/.test(lines[i])) {
          items.push(lines[i].replace(/^\s*\d+[.)、]\s+/, ''));
          i++;
        }
        out.push(
          <ol key={`ol-${key++}`} className="my-1.5 list-decimal space-y-1 pl-5 text-xs text-slate-700 leading-relaxed">
            {items.map((it, idx) => (
              <li key={idx}>{renderInline(it, onZoomImage)}</li>
            ))}
          </ol>
        );
        continue;
      }

      // Paragraph (always consumes at least the current line so an infinite loop is impossible)
      const pLines: string[] = [lines[i]];
      i++;
      while (
        i < lines.length &&
        lines[i].trim() !== '' &&
        !/^\s*#{1,6}\s+/.test(lines[i]) &&
        !lines[i].trim().startsWith('```') &&
        !lines[i].trim().startsWith('>') &&
        !/^\s*[-*_]{3,}\s*$/.test(lines[i]) &&
        !/^\s*[-*+]\s+/.test(lines[i]) &&
        !/^\s*\d+[.)、]\s+/.test(lines[i])
      ) {
        pLines.push(lines[i]);
        i++;
      }
      out.push(
        <p key={`p-${key++}`} className="my-1.5 text-xs text-slate-700 leading-relaxed whitespace-pre-wrap">
          {renderInline(pLines.join('\n'), onZoomImage)}
        </p>
      );
    }

    return out;
  }, [content, onZoomImage]);

  return <div>{blocks}</div>;
}

/* ==================== Main SOP Guides Panel ==================== */
export function SopGuidesPanel() {
  const [guides, setGuides] = useState<SopGuide[]>([]);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [stageFilter, setStageFilter] = useState<string>('all');
  const [search, setSearch] = useState('');
  const [loading, setLoading] = useState(false);
  const [uploadingFile, setUploadingFile] = useState(false);
  const [uploadingImg, setUploadingImg] = useState(false);
  const [errorMsg, setErrorMsg] = useState('');

  // Create / Edit state
  const [editingMode, setEditingMode] = useState<'none' | 'create' | 'edit'>('none');
  const [draftTitle, setDraftTitle] = useState('');
  const [draftStage, setDraftStage] = useState(GUIDE_STAGES[1]);
  const [draftMd, setDraftMd] = useState('');
  const [draftImages, setDraftImages] = useState<SopGuideImage[]>([]);
  const textareaRef = useRef<HTMLTextAreaElement | null>(null);
  const readerContentRef = useRef<HTMLDivElement | null>(null);

  // Lightbox zoom modal
  const [zoomedImg, setZoomedImg] = useState<{ src: string; alt: string } | null>(null);

  const loadGuides = useCallback(async (preferId?: string) => {
    try {
      const list = await api.listSopGuides();
      setGuides(list);
      setSelectedId((prev) => {
        if (preferId && list.some((g) => g.id === preferId)) return preferId;
        if (prev && list.some((g) => g.id === prev)) return prev;
        return list[0]?.id || null;
      });
    } catch {
      // ignore offline
    }
  }, []);

  useEffect(() => {
    loadGuides();
  }, [loadGuides]);

  const filteredGuides = useMemo(() => {
    const q = search.trim().toLowerCase();
    return guides.filter((g) => {
      if (stageFilter !== 'all' && g.stage_group !== stageFilter) return false;
      if (q) {
        const hay = `${g.title} ${g.stage_group} ${g.file_original_name} ${g.content_md}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
  }, [guides, stageFilter, search]);

  const activeGuide = useMemo(
    () => guides.find((g) => g.id === selectedId) || filteredGuides[0] || null,
    [guides, selectedId, filteredGuides]
  );

  function startCreate() {
    setErrorMsg('');
    setEditingMode('create');
    setDraftTitle('');
    setDraftStage(stageFilter !== 'all' ? stageFilter : GUIDE_STAGES[1]);
    setDraftMd('');
    setDraftImages([]);
  }

  function startEdit(guide: SopGuide) {
    setErrorMsg('');
    setEditingMode('edit');
    setDraftTitle(guide.title);
    setDraftStage(guide.stage_group || GUIDE_STAGES[1]);
    setDraftMd(guide.content_md || '');
    setDraftImages(guide.images || []);
  }

  async function handleImportFile(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;
    setUploadingFile(true);
    setErrorMsg('');
    try {
      const created = await api.importSopGuideFile(file, {
        stage_group: stageFilter !== 'all' ? stageFilter : '内网开标',
      });
      setEditingMode('none');
      await loadGuides(created.id);
    } catch (err) {
      setErrorMsg((err as Error).message);
    } finally {
      setUploadingFile(false);
    }
  }

  async function uploadAndAttachImages(files: File[], insertAtCursor = false) {
    if (files.length === 0) return;
    setUploadingImg(true);
    setErrorMsg('');
    try {
      const uploaded: SopGuideImage[] = [];
      for (const f of files) {
        const img = await api.uploadSopGuideImage(f);
        uploaded.push(img);
      }
      setDraftImages((prev) => [...prev, ...uploaded]);
      if (insertAtCursor) {
        const mdSnippets = uploaded
          .map((img) => `\n![${img.originalName.replace(/\.[^.]+$/, '')}](${img.url})\n`)
          .join('');
        const el = textareaRef.current;
        if (el) {
          const start = el.selectionStart ?? draftMd.length;
          const end = el.selectionEnd ?? draftMd.length;
          setDraftMd((prev) => prev.slice(0, start) + mdSnippets + prev.slice(end));
        } else {
          setDraftMd((prev) => prev + mdSnippets);
        }
      }
    } catch (err) {
      setErrorMsg((err as Error).message);
    } finally {
      setUploadingImg(false);
    }
  }

  function handlePaste(e: React.ClipboardEvent<HTMLTextAreaElement>) {
    const items = Array.from(e.clipboardData?.items || []);
    const imgFiles: File[] = [];
    for (const item of items) {
      if (item.type.startsWith('image/')) {
        const f = item.getAsFile();
        if (f) imgFiles.push(f);
      }
    }
    if (imgFiles.length > 0) {
      e.preventDefault();
      uploadAndAttachImages(imgFiles, true);
    }
  }

  async function handleSaveDraft(e: React.FormEvent) {
    e.preventDefault();
    if (!draftTitle.trim()) {
      setErrorMsg('请填写操作指南标题');
      return;
    }
    setLoading(true);
    setErrorMsg('');
    try {
      if (editingMode === 'create') {
        const created = await api.createSopGuide({
          stage_group: draftStage,
          title: draftTitle.trim(),
          content_md: draftMd,
          images: draftImages,
        });
        setEditingMode('none');
        await loadGuides(created.id);
      } else if (editingMode === 'edit' && activeGuide) {
        const updated = await api.updateSopGuide(activeGuide.id, {
          stage_group: draftStage,
          title: draftTitle.trim(),
          content_md: activeGuide.source_type === 'pdf' ? activeGuide.content_md : draftMd,
          images: draftImages,
        });
        setEditingMode('none');
        await loadGuides(updated.id);
      }
    } catch (err) {
      setErrorMsg((err as Error).message);
    } finally {
      setLoading(false);
    }
  }

  async function handleDelete(guide: SopGuide) {
    if (!window.confirm(`确认删除操作指南「${guide.title}」吗？`)) return;
    setLoading(true);
    try {
      await api.deleteSopGuide(guide.id);
      if (editingMode !== 'none') setEditingMode('none');
      await loadGuides();
    } catch (err) {
      setErrorMsg((err as Error).message);
    } finally {
      setLoading(false);
    }
  }

  function insertImageToMarkdown(img: SopGuideImage) {
    const snippet = `\n![${img.caption || img.originalName.replace(/\.[^.]+$/, '')}](${img.url})\n`;
    setDraftMd((prev) => prev + snippet);
  }

  function sanitizeFileName(name: string): string {
    return (name || '操作指南').replace(/[\\/:*?"<>|]/g, '_').trim() || '操作指南';
  }

  function handleExportMd(guide: SopGuide) {
    const trimmed = (guide.content_md || '').trim();
    const hasH1 = /^\s*#\s+/.test(trimmed);
    const origin = window.location.origin;
    const parts: string[] = [];
    if (!hasH1) parts.push(`# ${guide.title}`);
    if (trimmed) parts.push(trimmed);
    const extraImgs = (guide.images || []).filter((img) => !trimmed.includes(img.url));
    if (extraImgs.length > 0) {
      parts.push('## 操作截图附图');
      for (const img of extraImgs) {
        const absUrl = img.url.startsWith('http') ? img.url : `${origin}${img.url}`;
        parts.push(`![${img.caption || img.originalName}](${absUrl})`);
      }
    }
    const mdText = parts.join('\n\n') + '\n';
    const blob = new Blob([mdText], { type: 'text/markdown;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${sanitizeFileName(guide.title)}.md`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function handleExportPdf(guide: SopGuide) {
    if (guide.source_type === 'pdf' && guide.has_file) {
      window.location.href = `/api/sop/guides/${guide.id}/file?download=1`;
      return;
    }
    const bodyHtml = readerContentRef.current?.innerHTML || '';
    const safeTitle = guide.title.replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const safeStage = guide.stage_group.replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const exportDate = new Date().toISOString().slice(0, 10);

    const iframe = document.createElement('iframe');
    iframe.style.position = 'fixed';
    iframe.style.right = '0';
    iframe.style.bottom = '0';
    iframe.style.width = '0';
    iframe.style.height = '0';
    iframe.style.border = '0';
    document.body.appendChild(iframe);

    const doc = iframe.contentDocument || iframe.contentWindow?.document;
    if (!doc) {
      document.body.removeChild(iframe);
      return;
    }

    doc.open();
    doc.write(`<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <title>${safeTitle}</title>
  <style>
    @page { size: A4; margin: 16mm; }
    body { font-family: -apple-system, BlinkMacSystemFont, "PingFang SC", "Microsoft YaHei", sans-serif; color: #1e293b; font-size: 13px; line-height: 1.7; margin: 0; }
    .doc-header { border-bottom: 2px solid #3B5B7E; padding-bottom: 10px; margin-bottom: 16px; }
    .doc-meta { font-size: 11px; color: #64748b; margin-bottom: 4px; }
    .doc-title { font-size: 20px; font-weight: 700; color: #0f172a; margin: 0; }
    h1, h2, h3, h4 { page-break-after: avoid; }
    p, li { margin: 4px 0; }
    ul, ol { padding-left: 20px; margin: 6px 0; }
    blockquote { border-left: 3px solid #3B5B7E; background: #f1f5f9; padding: 8px 12px; margin: 10px 0; }
    pre { background: #0f172a; color: #f8fafc; padding: 10px 12px; border-radius: 6px; font-size: 11px; white-space: pre-wrap; word-break: break-all; }
    code { background: #f1f5f9; padding: 1px 4px; border-radius: 3px; font-size: 11px; color: #3B5B7E; }
    pre code { background: transparent; color: inherit; padding: 0; }
    table { width: 100%; border-collapse: collapse; margin: 10px 0; font-size: 12px; }
    th, td { border: 1px solid #cbd5e1; padding: 6px 8px; text-align: left; }
    th { background: #f8fafc; font-weight: 600; }
    img { max-width: 100%; max-height: 520px; height: auto; border: 1px solid #e2e8f0; border-radius: 6px; margin: 8px 0; page-break-inside: avoid; display: block; }
    hr { border: none; border-top: 1px solid #e2e8f0; margin: 12px 0; }
    .print-hide { display: none !important; }
  </style>
</head>
<body>
  <div class="doc-header">
    <div class="doc-meta">开评标系统操作 SOP · ${safeStage} · 导出日期 ${exportDate}</div>
    <h1 class="doc-title">${safeTitle}</h1>
  </div>
  <div>${bodyHtml}</div>
</body>
</html>`);
    doc.close();

    const imgs = Array.from(doc.querySelectorAll('img'));
    const triggerPrint = () => {
      const win = iframe.contentWindow;
      if (!win) {
        if (iframe.parentNode) iframe.parentNode.removeChild(iframe);
        return;
      }
      win.focus();
      win.print();
      setTimeout(() => {
        if (iframe.parentNode) iframe.parentNode.removeChild(iframe);
      }, 1500);
    };

    if (imgs.length === 0) {
      setTimeout(triggerPrint, 120);
    } else {
      let remaining = imgs.length;
      const done = () => {
        remaining--;
        if (remaining <= 0) setTimeout(triggerPrint, 120);
      };
      for (const img of imgs) {
        if (img.complete) {
          done();
        } else {
          img.onload = done;
          img.onerror = done;
        }
      }
    }
  }

  return (
    <div className="rounded-xl border border-slate-200 bg-white overflow-hidden">
      {/* Header Bar */}
      <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 bg-slate-50/80 px-4 py-2.5">
        <div className="flex flex-wrap items-center gap-2">
          <BookOpen className="h-4 w-4 text-[#3B5B7E]" />
          <h3 className="text-sm font-semibold text-slate-800">开评标系统操作 SOP 指南库</h3>
          <span className="rounded bg-[#EBF2F7] px-1.5 py-0.5 text-[10px] font-medium text-[#3B5B7E]">
            全批次共享 · 图文 / MD / PDF
          </span>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <label className="inline-flex cursor-pointer items-center gap-1.5 rounded-md border border-[#3B5B7E]/30 bg-white px-2.5 py-1.5 text-xs font-medium text-[#3B5B7E] hover:bg-[#EBF2F7]/60 transition-colors">
            <Upload className="h-3.5 w-3.5" />
            <span>{uploadingFile ? '读取中...' : '导入 MD / PDF'}</span>
            <input
              type="file"
              accept=".md,.markdown,.txt,.pdf"
              onChange={handleImportFile}
              className="hidden"
              disabled={uploadingFile}
            />
          </label>
          <button
            type="button"
            onClick={startCreate}
            className="inline-flex items-center gap-1 rounded-md bg-[#3B5B7E] px-3 py-1.5 text-xs font-medium text-white hover:bg-[#2E4866] transition-colors"
          >
            <Plus className="h-3.5 w-3.5" />
            <span>新建图文 SOP</span>
          </button>
        </div>
      </div>

      {errorMsg && (
        <div className="mx-4 mt-3 flex items-center justify-between rounded-lg border border-[#A94442]/30 bg-[#FEF2F2] px-3 py-2 text-xs text-[#A94442]">
          <span>{errorMsg}</span>
          <button type="button" onClick={() => setErrorMsg('')} className="underline text-[11px]">
            关闭
          </button>
        </div>
      )}

      {/* Stage Filter Pills */}
      <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-100 px-4 py-2 bg-white">
        <div className="flex flex-wrap items-center gap-1">
          <button
            type="button"
            onClick={() => setStageFilter('all')}
            className={`rounded-md px-2 py-1 text-[11px] font-medium transition-colors ${
              stageFilter === 'all' ? 'bg-slate-800 text-white' : 'bg-slate-100 text-slate-600 hover:bg-slate-200/70'
            }`}
          >
            全部 ({guides.length})
          </button>
          {GUIDE_STAGES.map((st) => {
            const cnt = guides.filter((g) => g.stage_group === st).length;
            return (
              <button
                key={st}
                type="button"
                onClick={() => setStageFilter(st)}
                className={`rounded-md px-2 py-1 text-[11px] font-medium transition-colors ${
                  stageFilter === st ? 'bg-[#3B5B7E] text-white' : 'bg-slate-100 text-slate-600 hover:bg-slate-200/70'
                }`}
              >
                {st}
                {cnt > 0 ? ` (${cnt})` : ''}
              </button>
            );
          })}
        </div>

        <div className="relative w-full sm:w-52">
          <Search className="pointer-events-none absolute left-2.5 top-2 h-3.5 w-3.5 text-slate-400" />
          <input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="搜索操作指南..."
            className="w-full rounded-lg border border-slate-200 py-1 pl-8 pr-3 text-xs focus:border-[#3B5B7E] focus:outline-none"
          />
        </div>
      </div>

      {/* Main Split View: Left Catalog + Right Reader/Editor */}
      <div className="grid grid-cols-1 lg:grid-cols-12">
        {/* Left Catalog */}
        <div className="border-b border-slate-200 lg:col-span-4 lg:border-r lg:border-b-0">
          {filteredGuides.length === 0 ? (
            <div className="p-6 text-center text-xs text-slate-400 space-y-1.5">
              <BookOpen className="mx-auto h-6 w-6 text-slate-300" />
              <div>暂无系统操作指南</div>
              <div className="text-[11px]">点击右上角「导入 MD / PDF」或「新建图文 SOP」</div>
            </div>
          ) : (
            <div className="max-h-[560px] divide-y divide-slate-100 overflow-y-auto">
              {filteredGuides.map((g) => {
                const isActive = editingMode === 'none' && activeGuide?.id === g.id;
                return (
                  <button
                    key={g.id}
                    type="button"
                    onClick={() => {
                      setEditingMode('none');
                      setSelectedId(g.id);
                    }}
                    className={`flex w-full items-start gap-2.5 px-3.5 py-2.5 text-left transition-colors ${
                      isActive ? 'bg-[#EBF2F7]/70' : 'hover:bg-slate-50'
                    }`}
                  >
                    {g.source_type === 'pdf' ? (
                      <FileText className="mt-0.5 h-4 w-4 shrink-0 text-[#A94442]" />
                    ) : (
                      <FileCode2 className="mt-0.5 h-4 w-4 shrink-0 text-[#3B5B7E]" />
                    )}
                    <div className="min-w-0 flex-1">
                      <div className="truncate text-xs font-medium text-slate-800">{g.title}</div>
                      <div className="mt-1 flex flex-wrap items-center gap-1.5 text-[10px] text-slate-400">
                        <span className="rounded bg-slate-100 px-1.5 py-0.2 text-slate-600">
                          {g.stage_group}
                        </span>
                        <span>{g.source_type === 'pdf' ? 'PDF 手册' : '图文 MD'}</span>
                        {g.images.length > 0 && <span>· {g.images.length} 图</span>}
                      </div>
                    </div>
                  </button>
                );
              })}
            </div>
          )}
        </div>

        {/* Right Reader or Editor */}
        <div className="lg:col-span-8 p-4">
          {editingMode !== 'none' ? (
            <form onSubmit={handleSaveDraft} className="space-y-3">
              <div className="flex items-center justify-between border-b border-slate-100 pb-2">
                <h4 className="text-xs font-semibold text-slate-800">
                  {editingMode === 'create' ? '新建开评标系统图文 SOP' : '编辑系统操作 SOP'}
                </h4>
                <button
                  type="button"
                  onClick={() => setEditingMode('none')}
                  className="text-xs text-slate-400 hover:text-slate-600"
                >
                  取消
                </button>
              </div>

              <div className="grid grid-cols-1 gap-2 sm:grid-cols-3">
                <select
                  value={draftStage}
                  onChange={(e) => setDraftStage(e.target.value)}
                  className="rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                >
                  {GUIDE_STAGES.map((st) => (
                    <option key={st} value={st}>
                      {st}
                    </option>
                  ))}
                </select>
                <input
                  required
                  value={draftTitle}
                  onChange={(e) => setDraftTitle(e.target.value)}
                  placeholder="操作指南标题（如：ECP 内网开标解密操作步骤）*"
                  className="sm:col-span-2 rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>

              {!(editingMode === 'edit' && activeGuide?.source_type === 'pdf') && (
                <>
                  <div className="space-y-1.5">
                    <div className="flex flex-wrap items-center justify-between gap-2 text-[11px] text-slate-500">
                      <span>支持 Markdown 语法，可直接在文本框内按 Ctrl+V / Cmd+V 粘贴截图</span>
                      <label className="inline-flex cursor-pointer items-center gap-1 rounded border border-slate-200 bg-slate-50 px-2 py-0.5 text-[11px] font-medium text-[#3B5B7E] hover:bg-[#EBF2F7]">
                        <ImagePlus className="h-3.5 w-3.5" />
                        <span>{uploadingImg ? '上传中...' : '上传并插入截图'}</span>
                        <input
                          type="file"
                          multiple
                          accept="image/*"
                          onChange={(e) => {
                            const files = Array.from(e.target.files || []);
                            e.target.value = '';
                            uploadAndAttachImages(files, true);
                          }}
                          className="hidden"
                          disabled={uploadingImg}
                        />
                      </label>
                    </div>
                    <textarea
                      ref={textareaRef}
                      rows={11}
                      value={draftMd}
                      onChange={(e) => setDraftMd(e.target.value)}
                      onPaste={handlePaste}
                      placeholder={
                        '支持 Markdown 排版：\n## 1. 登录内网开标系统\n1. 插入 CA 证书并打开浏览器...\n2. 点击「开标大厅」进入对应标段...\n\n（直接粘贴截图会自动上传并插入到当前位置）'
                      }
                      className="w-full rounded-lg border border-slate-200 bg-white p-3 font-mono text-xs leading-relaxed focus:border-[#3B5B7E] focus:outline-none"
                    />
                  </div>

                  {/* Attached Images Strip */}
                  {draftImages.length > 0 && (
                    <div className="space-y-1.5 rounded-lg border border-slate-200 bg-slate-50/70 p-2.5">
                      <div className="text-[11px] font-medium text-slate-600">
                        已上传附图（{draftImages.length} 张 · 点击「插入正文」可插到 Markdown 中）
                      </div>
                      <div className="flex flex-wrap gap-2">
                        {draftImages.map((img, idx) => (
                          <div
                            key={`${img.filename}-${idx}`}
                            className="group relative w-28 overflow-hidden rounded-lg border border-slate-200 bg-white"
                          >
                            <img
                              src={img.url}
                              alt={img.originalName}
                              className="h-16 w-full object-cover cursor-zoom-in"
                              onClick={() => setZoomedImg({ src: img.url, alt: img.originalName })}
                            />
                            <div className="flex items-center justify-between gap-1 px-1.5 py-1 text-[10px]">
                              <button
                                type="button"
                                onClick={() => insertImageToMarkdown(img)}
                                className="text-[#3B5B7E] hover:underline"
                              >
                                插入正文
                              </button>
                              <button
                                type="button"
                                onClick={() =>
                                  setDraftImages((prev) => prev.filter((_, i) => i !== idx))
                                }
                                className="text-slate-400 hover:text-[#A94442]"
                                title="移除"
                              >
                                <X className="h-3 w-3" />
                              </button>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                </>
              )}

              <div className="flex justify-end gap-2 pt-1">
                <button
                  type="button"
                  onClick={() => setEditingMode('none')}
                  className="rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs text-slate-600 hover:bg-slate-50"
                >
                  取消
                </button>
                <button
                  type="submit"
                  disabled={loading}
                  className="inline-flex items-center gap-1 rounded-lg bg-[#2F6F5E] px-3.5 py-1.5 text-xs font-medium text-white hover:bg-[#24594B] disabled:opacity-50"
                >
                  <Save className="h-3.5 w-3.5" />
                  <span>{loading ? '保存中...' : '保存操作指南'}</span>
                </button>
              </div>
            </form>
          ) : activeGuide ? (
            <div className="space-y-3">
              {/* Guide Reader Header */}
              <div className="flex flex-wrap items-start justify-between gap-2 border-b border-slate-100 pb-2.5">
                <div className="min-w-0">
                  <div className="flex flex-wrap items-center gap-2">
                    <span className="rounded bg-[#EBF2F7] px-1.5 py-0.5 text-[10px] font-medium text-[#3B5B7E]">
                      {activeGuide.stage_group}
                    </span>
                    <span className="text-[11px] text-slate-400">
                      {activeGuide.source_type === 'pdf'
                        ? `PDF 文件 · ${activeGuide.file_original_name}`
                        : activeGuide.file_original_name
                          ? `导入自 ${activeGuide.file_original_name}`
                          : '图文 Markdown'}
                    </span>
                  </div>
                  <h4 className="mt-1 text-sm font-semibold text-slate-800">{activeGuide.title}</h4>
                </div>

                <div className="flex flex-wrap items-center gap-1.5 shrink-0">
                  {activeGuide.source_type === 'pdf' && activeGuide.has_file ? (
                    <>
                      <a
                        href={`/api/sop/guides/${activeGuide.id}/file`}
                        target="_blank"
                        rel="noreferrer"
                        className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2.5 py-1 text-[11px] font-medium text-slate-600 hover:border-[#3B5B7E] hover:text-[#3B5B7E]"
                      >
                        <ExternalLink className="h-3 w-3" />
                        <span>全屏打开</span>
                      </a>
                      <a
                        href={`/api/sop/guides/${activeGuide.id}/file?download=1`}
                        className="inline-flex items-center gap-1 rounded-md border border-[#2F6F5E]/30 bg-[#ECFDF5]/60 px-2.5 py-1 text-[11px] font-medium text-[#2F6F5E] hover:bg-[#ECFDF5]"
                      >
                        <Download className="h-3 w-3" />
                        <span>导出 PDF</span>
                      </a>
                    </>
                  ) : (
                    <>
                      <button
                        type="button"
                        onClick={() => handleExportMd(activeGuide)}
                        className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2.5 py-1 text-[11px] font-medium text-slate-600 hover:border-[#3B5B7E] hover:text-[#3B5B7E]"
                      >
                        <Download className="h-3 w-3" />
                        <span>导出 MD</span>
                      </button>
                      <button
                        type="button"
                        onClick={() => handleExportPdf(activeGuide)}
                        className="inline-flex items-center gap-1 rounded-md border border-[#2F6F5E]/30 bg-[#ECFDF5]/60 px-2.5 py-1 text-[11px] font-medium text-[#2F6F5E] hover:bg-[#ECFDF5]"
                      >
                        <FileText className="h-3 w-3" />
                        <span>导出 PDF</span>
                      </button>
                    </>
                  )}
                  <button
                    type="button"
                    onClick={() => startEdit(activeGuide)}
                    className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2.5 py-1 text-[11px] font-medium text-slate-600 hover:border-[#3B5B7E] hover:text-[#3B5B7E]"
                  >
                    <Edit3 className="h-3 w-3" />
                    <span>{activeGuide.source_type === 'pdf' ? '改标题/阶段' : '编辑/加图'}</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => handleDelete(activeGuide)}
                    className="inline-flex items-center gap-1 rounded-md border border-[#A94442]/25 bg-[#FEF2F2]/70 px-2 py-1 text-[11px] font-medium text-[#A94442] hover:bg-[#FEE2E2]"
                  >
                    <Trash2 className="h-3 w-3" />
                  </button>
                </div>
              </div>

              {/* Content Body */}
              {activeGuide.source_type === 'pdf' ? (
                <div className="overflow-hidden rounded-lg border border-slate-200 bg-slate-100">
                  <iframe
                    src={`/api/sop/guides/${activeGuide.id}/file`}
                    title={activeGuide.title}
                    className="h-[520px] w-full"
                  />
                </div>
              ) : (
                <div ref={readerContentRef} className="space-y-4">
                  <div className="max-h-[460px] overflow-y-auto pr-1">
                    {activeGuide.content_md.trim() ? (
                      <MarkdownView
                        content={activeGuide.content_md}
                        onZoomImage={(src, alt) => setZoomedImg({ src, alt })}
                      />
                    ) : (
                      <div className="py-4 text-xs text-slate-400">暂无正文文字，点击右上角「编辑/加图」补充说明</div>
                    )}
                  </div>

                  {/* Extra Images Gallery (shows images not already embedded inline in markdown) */}
                  {(() => {
                    const extraImgs = (activeGuide.images || []).filter(
                      (img) => !activeGuide.content_md.includes(img.url)
                    );
                    if (extraImgs.length === 0) return null;
                    return (
                      <div className="border-t border-slate-100 pt-3 space-y-2">
                        <div className="text-[11px] font-semibold text-slate-600">
                          操作截图附图（点击放大）
                        </div>
                        <div className="grid grid-cols-2 gap-2.5 sm:grid-cols-3">
                          {extraImgs.map((img, idx) => (
                            <button
                              key={`${img.filename}-${idx}`}
                              type="button"
                              onClick={() => setZoomedImg({ src: img.url, alt: img.originalName })}
                              className="group relative overflow-hidden rounded-lg border border-slate-200 bg-slate-50 text-left"
                            >
                              <img
                                src={img.url}
                                alt={img.originalName}
                                className="h-28 w-full object-cover transition-transform group-hover:scale-[1.02]"
                              />
                              <div className="absolute right-1.5 top-1.5 rounded bg-slate-900/60 p-1 text-white opacity-0 transition-opacity group-hover:opacity-100">
                                <ZoomIn className="h-3 w-3" />
                              </div>
                              <div className="truncate px-2 py-1 text-[10px] text-slate-500">
                                {img.caption || img.originalName}
                              </div>
                            </button>
                          ))}
                        </div>
                      </div>
                    );
                  })()}
                </div>
              )}
            </div>
          ) : (
            <div className="flex h-56 flex-col items-center justify-center text-center text-xs text-slate-400 space-y-2">
              <BookOpen className="h-7 w-7 text-slate-300" />
              <p>可以把 ECP/ETP 系统开标、解密、多轮报价、评标导表等图文指南集中沉淀在这里</p>
              <p className="text-[11px] text-slate-400">
                支持直接上传 <strong>.md</strong> 自动排版、上传 <strong>.pdf</strong> 在线翻阅，或在线编写并 <strong>Ctrl+V 粘贴截图</strong>
              </p>
            </div>
          )}
        </div>
      </div>

      {/* Fullscreen Image Lightbox */}
      {zoomedImg && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/80 p-4 backdrop-blur-[1px]"
          onClick={() => setZoomedImg(null)}
        >
          <div
            className="relative max-h-[90vh] max-w-[92vw] overflow-hidden rounded-xl bg-white p-2 shadow-2xl"
            onClick={(e) => e.stopPropagation()}
          >
            <button
              type="button"
              onClick={() => setZoomedImg(null)}
              className="absolute right-3 top-3 rounded-full bg-slate-900/70 p-1.5 text-white hover:bg-slate-900"
            >
              <X className="h-4 w-4" />
            </button>
            <img
              src={zoomedImg.src}
              alt={zoomedImg.alt}
              className="max-h-[82vh] max-w-[88vw] object-contain"
            />
            {zoomedImg.alt && (
              <div className="px-2 pt-1.5 text-center text-xs text-slate-500">{zoomedImg.alt}</div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}
