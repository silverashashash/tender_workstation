import fs from 'node:fs';
import path from 'node:path';
import archiver from 'archiver';
import { Router } from 'express';
import multer from 'multer';
import { v4 as uuid } from 'uuid';
import { UPLOADS_DIR } from '../db.js';
import { runChecklistCompliance } from '../checklistCompliance.js';
import { runPerformanceMatch } from '../performanceMatcher.js';
import { parseExpertNamesFromExcel, runRandomScoreGenerator } from '../randomScoreGenerator.js';
import { runScoreTemplateMerge } from '../scoreTemplateMerger.js';
import {
  buildWordBatchOperations,
  processSingleWordFile,
  type WordBatchConfig,
} from '../wordBatchProcessor.js';

export const standaloneToolboxRouter = Router();

const wordBatchUpload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 80 * 1024 * 1024, files: 50 },
});

const WORD_BATCH_ROOT = path.join(UPLOADS_DIR, '_toolbox_word_batch');
const JOB_TTL_MS = 2 * 60 * 60 * 1000; // 2 hours

interface StoredJobManifest {
  jobId: string;
  createdAt: string;
  logText: string;
  files: Array<{
    key: string;
    filename: string;
    diskName: string;
    kind: 'docx' | 'xlsx';
    size: number;
  }>;
}

function decodeName(raw: string): string {
  try {
    return Buffer.from(raw, 'latin1').toString('utf8');
  } catch {
    return raw;
  }
}

function pruneExpiredJobs() {
  if (!fs.existsSync(WORD_BATCH_ROOT)) return;
  const now = Date.now();
  try {
    for (const entry of fs.readdirSync(WORD_BATCH_ROOT)) {
      const dir = path.join(WORD_BATCH_ROOT, entry);
      const stat = fs.statSync(dir);
      if (stat.isDirectory() && now - stat.mtimeMs > JOB_TTL_MS) {
        fs.rmSync(dir, { recursive: true, force: true });
      }
    }
  } catch {
    // ignore cleanup errors
  }
}

function loadManifest(jobId: string): { dir: string; manifest: StoredJobManifest } | null {
  if (!/^[a-zA-Z0-9-]+$/.test(jobId)) return null;
  const dir = path.join(WORD_BATCH_ROOT, jobId);
  const manifestPath = path.join(dir, 'manifest.json');
  if (!fs.existsSync(manifestPath)) return null;
  try {
    const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8')) as StoredJobManifest;
    return { dir, manifest };
  } catch {
    return null;
  }
}

// POST /api/toolbox/word-batch/process
standaloneToolboxRouter.post(
  '/word-batch/process',
  wordBatchUpload.array('files', 50),
  async (req, res) => {
    try {
      pruneExpiredJobs();
      const uploaded = (req.files as Express.Multer.File[] | undefined) || [];
      if (uploaded.length === 0) {
        return res.status(400).json({ error: '请至少选择一个 Word 文件（.docx / .doc）' });
      }

      let cfg: WordBatchConfig;
      try {
        cfg = JSON.parse(String(req.body?.config || '{}')) as WordBatchConfig;
      } catch {
        return res.status(400).json({ error: '配置参数格式错误' });
      }

      const operations = buildWordBatchOperations(cfg);
      const extractEnabled = Boolean(cfg.extractEnabled);

      // Filter out Office temp lock files (~$*)
      const targetFiles = uploaded
        .map((f) => ({ originalName: decodeName(f.originalname), buffer: f.buffer }))
        .filter((f) => {
          const base = path.basename(f.originalName);
          return base && !base.startsWith('~$');
        });

      if (targetFiles.length === 0) {
        return res.status(400).json({ error: '没有可处理的 .doc / .docx 文件' });
      }

      const jobId = uuid();
      const jobDir = path.join(WORD_BATCH_ROOT, jobId);
      fs.mkdirSync(jobDir, { recursive: true });

      const opNames = operations.map((o) => o.name);
      if (extractEnabled) opNames.push('提取表格');

      const nowStr = new Date().toLocaleString('zh-CN', { hour12: false });
      const allLogs: string[] = [
        `==== 开始处理（${nowStr}）====`,
        `操作：${opNames.join('、')}`,
        `共 ${targetFiles.length} 个文件：`,
      ];

      const grandTotals: Record<string, number> = {};
      const manifestFiles: StoredJobManifest['files'] = [];
      const fileSummaries: Array<{
        originalName: string;
        ok: boolean;
        convertedFromDoc: boolean;
        counts: Record<string, number>;
        outputs: Array<{ key: string; filename: string; kind: 'docx' | 'xlsx'; size: number }>;
        error?: string;
      }> = [];

      let okCount = 0;
      let fileIdx = 0;

      for (const f of targetFiles) {
        const result = await processSingleWordFile(
          f.originalName,
          f.buffer,
          operations,
          extractEnabled
        );
        allLogs.push(...result.logs);

        const outItems: Array<{ key: string; filename: string; kind: 'docx' | 'xlsx'; size: number }> = [];
        if (result.ok) {
          okCount++;
          for (const [k, v] of Object.entries(result.counts)) {
            grandTotals[k] = (grandTotals[k] || 0) + v;
          }
          for (const out of result.outputs) {
            fileIdx++;
            const key = `f${fileIdx}`;
            const diskName = `${key}.${out.kind}`;
            fs.writeFileSync(path.join(jobDir, diskName), out.buffer);
            manifestFiles.push({
              key,
              filename: out.filename,
              diskName,
              kind: out.kind,
              size: out.buffer.length,
            });
            outItems.push({
              key,
              filename: out.filename,
              kind: out.kind,
              size: out.buffer.length,
            });
          }
        }

        fileSummaries.push({
          originalName: result.originalName,
          ok: result.ok,
          convertedFromDoc: result.convertedFromDoc,
          counts: result.counts,
          outputs: outItems,
          error: result.error,
        });
      }

      const summaryText = Object.entries(grandTotals)
        .map(([k, v]) => `${k} 共 ${v} ${k === '提取表格' ? '行' : '处'}`)
        .join('，');
      allLogs.push(
        `==== 完成：成功 ${okCount}/${targetFiles.length} 个文件${summaryText ? `；${summaryText}` : ''} ====`
      );

      const logText = allLogs.join('\n');
      const manifest: StoredJobManifest = {
        jobId,
        createdAt: new Date().toISOString(),
        logText,
        files: manifestFiles,
      };
      fs.writeFileSync(path.join(jobDir, 'manifest.json'), JSON.stringify(manifest, null, 2), 'utf8');

      res.status(201).json({
        jobId,
        totalFiles: targetFiles.length,
        okFiles: okCount,
        grandTotals,
        files: fileSummaries,
        logs: allLogs,
        hasOutputs: manifestFiles.length > 0,
      });
    } catch (err) {
      res.status(400).json({ error: (err as Error).message });
    }
  }
);

// GET /api/toolbox/word-batch/:jobId/download/:fileKey
standaloneToolboxRouter.get('/word-batch/:jobId/download/:fileKey', (req, res) => {
  const { jobId, fileKey } = req.params;
  const loaded = loadManifest(jobId);
  if (!loaded) return res.status(404).json({ error: '处理结果已过期或不存在，请重新执行处理' });

  const item = loaded.manifest.files.find((f) => f.key === fileKey);
  if (!item) return res.status(404).json({ error: '未找到对应的输出文件' });

  const absPath = path.join(loaded.dir, item.diskName);
  if (!fs.existsSync(absPath)) return res.status(404).json({ error: '输出文件已失效，请重新处理' });

  const mime =
    item.kind === 'docx'
      ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  res.setHeader('Content-Type', mime);
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(item.filename)}`);
  res.send(fs.readFileSync(absPath));
});

// GET /api/toolbox/word-batch/:jobId/download-all
standaloneToolboxRouter.get('/word-batch/:jobId/download-all', (req, res) => {
  const { jobId } = req.params;
  const loaded = loadManifest(jobId);
  if (!loaded) return res.status(404).json({ error: '处理结果已过期或不存在，请重新执行处理' });
  if (loaded.manifest.files.length === 0) {
    return res.status(400).json({ error: '当前任务没有可下载的输出文件' });
  }

  const stamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const zipName = `Word批量处理结果_${stamp}.zip`;
  res.setHeader('Content-Type', 'application/zip');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(zipName)}`);

  const archive = archiver('zip', { zlib: { level: 6 } });
  archive.on('error', (err) => {
    if (!res.headersSent) {
      res.status(500).json({ error: `打包 ZIP 失败：${err.message}` });
    }
  });
  archive.pipe(res);

  const usedNames = new Set<string>();
  for (const item of loaded.manifest.files) {
    const absPath = path.join(loaded.dir, item.diskName);
    if (!fs.existsSync(absPath)) continue;
    let entryName = item.filename;
    if (usedNames.has(entryName)) {
      const ext = path.extname(entryName);
      const stem = entryName.slice(0, -ext.length);
      let seq = 2;
      while (usedNames.has(`${stem}_${seq}${ext}`)) seq++;
      entryName = `${stem}_${seq}${ext}`;
    }
    usedNames.add(entryName);
    archive.file(absPath, { name: entryName });
  }

  archive.append(loaded.manifest.logText + '\n', { name: `处理日志_${stamp}.txt` });
  archive.finalize();
});

// ==================== 专家打分表按模板合并 (合并表格并标记) ====================

const SCORE_MERGE_ROOT = path.join(UPLOADS_DIR, '_toolbox_score_merge');

const scoreMergeUpload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 80 * 1024 * 1024, files: 100 },
});

function pruneExpiredScoreMergeJobs() {
  if (!fs.existsSync(SCORE_MERGE_ROOT)) return;
  const now = Date.now();
  try {
    for (const entry of fs.readdirSync(SCORE_MERGE_ROOT)) {
      const dir = path.join(SCORE_MERGE_ROOT, entry);
      const stat = fs.statSync(dir);
      if (stat.isDirectory() && now - stat.mtimeMs > JOB_TTL_MS) {
        fs.rmSync(dir, { recursive: true, force: true });
      }
    }
  } catch {
    // ignore
  }
}

// POST /api/toolbox/score-merge/run
standaloneToolboxRouter.post(
  '/score-merge/run',
  scoreMergeUpload.fields([
    { name: 'files', maxCount: 100 },
    { name: 'template', maxCount: 1 },
    { name: 'successList', maxCount: 1 },
  ]),
  async (req, res) => {
    try {
      pruneExpiredScoreMergeJobs();
      const fieldFiles = (req.files || {}) as Record<string, Express.Multer.File[]>;
      const uploadedScores = fieldFiles.files || [];
      if (uploadedScores.length === 0) {
        return res.status(400).json({ error: '请至少选择一份专家打分表（.xlsx 或 .zip）' });
      }
      const uploadedTemplate = fieldFiles.template?.[0] || null;
      const uploadedSuccessList = fieldFiles.successList?.[0] || null;

      const result = await runScoreTemplateMerge({
        scoreFiles: uploadedScores.map((f) => ({
          originalName: decodeName(f.originalname),
          buffer: f.buffer,
        })),
        templateFile: uploadedTemplate
          ? { originalName: decodeName(uploadedTemplate.originalname), buffer: uploadedTemplate.buffer }
          : null,
        successListFile: uploadedSuccessList
          ? { originalName: decodeName(uploadedSuccessList.originalname), buffer: uploadedSuccessList.buffer }
          : null,
      });

      const jobId = uuid();
      const jobDir = path.join(SCORE_MERGE_ROOT, jobId);
      fs.mkdirSync(jobDir, { recursive: true });
      fs.writeFileSync(path.join(jobDir, 'output.xlsx'), result.buffer);
      fs.writeFileSync(
        path.join(jobDir, 'meta.json'),
        JSON.stringify({ outputFileName: result.outputFileName }),
        'utf8'
      );

      res.status(201).json({
        jobId,
        outputFileName: result.outputFileName,
        totalInputFiles: result.totalInputFiles,
        validFileCount: result.validFileCount,
        skippedFiles: result.skippedFiles,
        usedTemplate: result.usedTemplate,
        activeRows: result.activeRows,
        ignoredRows: result.ignoredRows,
        unmatchedRows: result.unmatchedRows,
        conflictCount: result.conflictCount,
        highlightCount: result.highlightCount,
        headerSummary: result.headerSummary,
        sheetPreviews: result.sheetPreviews,
        logs: result.logs,
      });
    } catch (err) {
      res.status(400).json({ error: (err as Error).message });
    }
  }
);

// GET /api/toolbox/score-merge/:jobId/download
standaloneToolboxRouter.get('/score-merge/:jobId/download', (req, res) => {
  const { jobId } = req.params;
  if (!/^[a-zA-Z0-9-]+$/.test(jobId)) {
    return res.status(400).json({ error: '无效的任务 ID' });
  }
  const jobDir = path.join(SCORE_MERGE_ROOT, jobId);
  const outPath = path.join(jobDir, 'output.xlsx');
  const metaPath = path.join(jobDir, 'meta.json');
  if (!fs.existsSync(outPath) || !fs.existsSync(metaPath)) {
    return res.status(404).json({ error: '合并结果已过期或不存在，请重新执行合并' });
  }
  const meta = JSON.parse(fs.readFileSync(metaPath, 'utf8')) as { outputFileName: string };
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader(
    'Content-Disposition',
    `attachment; filename*=UTF-8''${encodeURIComponent(meta.outputFileName)}`
  );
  res.send(fs.readFileSync(outPath));
});

// ==================== 履约分匹配填入 (履约分表 -> 详评打分模板绩效评价列) ====================

const PERF_MATCH_ROOT = path.join(UPLOADS_DIR, '_toolbox_perf_match');

function pruneExpiredPerfMatchJobs() {
  if (!fs.existsSync(PERF_MATCH_ROOT)) return;
  const now = Date.now();
  try {
    for (const entry of fs.readdirSync(PERF_MATCH_ROOT)) {
      const dir = path.join(PERF_MATCH_ROOT, entry);
      const stat = fs.statSync(dir);
      if (stat.isDirectory() && now - stat.mtimeMs > JOB_TTL_MS) {
        fs.rmSync(dir, { recursive: true, force: true });
      }
    }
  } catch {
    // ignore
  }
}

interface PerfMatchManifest {
  reportText: string;
  files: Array<{ key: string; filename: string; diskName: string }>;
}

// POST /api/toolbox/perf-match/run
standaloneToolboxRouter.post(
  '/perf-match/run',
  scoreMergeUpload.fields([
    { name: 'lyf', maxCount: 1 },
    { name: 'templates', maxCount: 50 },
  ]),
  async (req, res) => {
    try {
      pruneExpiredPerfMatchJobs();
      const fieldFiles = (req.files || {}) as Record<string, Express.Multer.File[]>;
      const lyfUploaded = fieldFiles.lyf?.[0];
      const tplUploaded = fieldFiles.templates || [];
      if (!lyfUploaded) return res.status(400).json({ error: '请先上传《履约分》Excel 文件（.xlsx）' });
      if (tplUploaded.length === 0) return res.status(400).json({ error: '请至少上传一份详评打分模板（.xlsx / .zip）' });

      const rawDefault = String(req.body?.defaultScore ?? '4').trim();
      const defaultScore = rawDefault === '' ? 4 : Number(rawDefault);

      const result = await runPerformanceMatch({
        lyfFile: { originalName: decodeName(lyfUploaded.originalname), buffer: lyfUploaded.buffer },
        templateFiles: tplUploaded.map((f) => ({
          originalName: decodeName(f.originalname),
          buffer: f.buffer,
        })),
        defaultScore,
      });

      const jobId = uuid();
      const jobDir = path.join(PERF_MATCH_ROOT, jobId);
      fs.mkdirSync(jobDir, { recursive: true });

      const manifestFiles: PerfMatchManifest['files'] = [];
      const templatesPayload = result.templates.map((t, idx) => {
        const key = `t${idx + 1}`;
        const diskName = `${key}.xlsx`;
        fs.writeFileSync(path.join(jobDir, diskName), t.buffer);
        manifestFiles.push({ key, filename: t.outputFileName, diskName });
        return {
          key,
          originalName: t.originalName,
          outputFileName: t.outputFileName,
          sheetsProcessed: t.sheetsProcessed,
          sheetsSkipped: t.sheetsSkipped,
          totalRows: t.totalRows,
          matched: t.matched,
          defaulted: t.defaulted,
        };
      });

      const manifest: PerfMatchManifest = {
        reportText: result.reportText,
        files: manifestFiles,
      };
      fs.writeFileSync(path.join(jobDir, 'manifest.json'), JSON.stringify(manifest, null, 2), 'utf8');

      res.status(201).json({
        jobId,
        lyfFileName: result.lyfFileName,
        lyfRecordCount: result.lyfRecordCount,
        lyfDuplicates: result.lyfDuplicates,
        defaultScore: result.defaultScore,
        templates: templatesPayload,
        totalRows: result.totalRows,
        totalMatched: result.totalMatched,
        totalDefaulted: result.totalDefaulted,
        totalSkippedSheets: result.totalSkippedSheets,
        unmatchedItems: result.templates.flatMap((t) => t.unmatched),
        logs: result.logs,
      });
    } catch (err) {
      res.status(400).json({ error: (err as Error).message });
    }
  }
);

// GET /api/toolbox/perf-match/:jobId/download/:fileKey
standaloneToolboxRouter.get('/perf-match/:jobId/download/:fileKey', (req, res) => {
  const { jobId, fileKey } = req.params;
  if (!/^[a-zA-Z0-9-]+$/.test(jobId)) return res.status(400).json({ error: '无效的任务 ID' });
  const jobDir = path.join(PERF_MATCH_ROOT, jobId);
  const manifestPath = path.join(jobDir, 'manifest.json');
  if (!fs.existsSync(manifestPath)) return res.status(404).json({ error: '结果已过期，请重新执行匹配' });
  const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8')) as PerfMatchManifest;
  const item = manifest.files.find((f) => f.key === fileKey);
  if (!item) return res.status(404).json({ error: '未找到对应模板文件' });
  const absPath = path.join(jobDir, item.diskName);
  if (!fs.existsSync(absPath)) return res.status(404).json({ error: '输出文件已失效' });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(item.filename)}`);
  res.send(fs.readFileSync(absPath));
});

// GET /api/toolbox/perf-match/:jobId/download-all
standaloneToolboxRouter.get('/perf-match/:jobId/download-all', (req, res) => {
  const { jobId } = req.params;
  if (!/^[a-zA-Z0-9-]+$/.test(jobId)) return res.status(400).json({ error: '无效的任务 ID' });
  const jobDir = path.join(PERF_MATCH_ROOT, jobId);
  const manifestPath = path.join(jobDir, 'manifest.json');
  if (!fs.existsSync(manifestPath)) return res.status(404).json({ error: '结果已过期，请重新执行匹配' });
  const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8')) as PerfMatchManifest;
  const stamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const zipName = `履约分匹配结果_${stamp}.zip`;
  res.setHeader('Content-Type', 'application/zip');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(zipName)}`);

  const archive = archiver('zip', { zlib: { level: 6 } });
  archive.on('error', (err) => {
    if (!res.headersSent) res.status(500).json({ error: `打包失败：${err.message}` });
  });
  archive.pipe(res);

  const usedNames = new Set<string>();
  for (const item of manifest.files) {
    const absPath = path.join(jobDir, item.diskName);
    if (!fs.existsSync(absPath)) continue;
    let entryName = item.filename;
    if (usedNames.has(entryName)) {
      const ext = path.extname(entryName);
      const stem = entryName.slice(0, -ext.length);
      let seq = 2;
      while (usedNames.has(`${stem}_${seq}${ext}`)) seq++;
      entryName = `${stem}_${seq}${ext}`;
    }
    usedNames.add(entryName);
    archive.file(absPath, { name: entryName });
  }
  archive.append(manifest.reportText + '\n', { name: '匹配报告.txt' });
  archive.finalize();
});

// ==================== 清单合规性检验 (递归解压 ZIP/7z/RAR + 校验清单 Excel) ====================

const CHECKLIST_ROOT = path.join(UPLOADS_DIR, '_toolbox_checklist_compliance');

const checklistUpload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 200 * 1024 * 1024, files: 100 },
});

function pruneExpiredChecklistJobs() {
  if (!fs.existsSync(CHECKLIST_ROOT)) return;
  const now = Date.now();
  try {
    for (const entry of fs.readdirSync(CHECKLIST_ROOT)) {
      const dir = path.join(CHECKLIST_ROOT, entry);
      const stat = fs.statSync(dir);
      if (stat.isDirectory() && now - stat.mtimeMs > JOB_TTL_MS) {
        fs.rmSync(dir, { recursive: true, force: true });
      }
    }
  } catch {
    // ignore
  }
}

// POST /api/toolbox/checklist-compliance/run
standaloneToolboxRouter.post(
  '/checklist-compliance/run',
  checklistUpload.array('files', 100),
  async (req, res) => {
    try {
      pruneExpiredChecklistJobs();
      const uploaded = (req.files as Express.Multer.File[] | undefined) || [];
      if (uploaded.length === 0) {
        return res.status(400).json({ error: '请至少上传一个 .zip 压缩包' });
      }
      const rawKeywords = String(req.body?.keywords ?? '').trim();
      const keywords = rawKeywords
        ? rawKeywords
            .replace(/，|、/g, ',')
            .split(',')
            .map((s) => s.trim())
            .filter(Boolean)
        : undefined;

      const result = await runChecklistCompliance({
        files: uploaded.map((f) => ({
          originalName: decodeName(f.originalname),
          buffer: f.buffer,
        })),
        keywords,
      });

      const jobId = uuid();
      const jobDir = path.join(CHECKLIST_ROOT, jobId);
      fs.mkdirSync(jobDir, { recursive: true });
      fs.writeFileSync(path.join(jobDir, 'report.xlsx'), result.reportBuffer);

      res.status(201).json({
        jobId,
        keywords: result.keywords,
        counts: result.counts,
        items: result.items,
        reportFileName: result.reportFileName,
        logs: result.logs,
      });
    } catch (err) {
      res.status(400).json({ error: (err as Error).message });
    }
  }
);

// GET /api/toolbox/checklist-compliance/:jobId/download
standaloneToolboxRouter.get('/checklist-compliance/:jobId/download', (req, res) => {
  const { jobId } = req.params;
  if (!/^[a-zA-Z0-9-]+$/.test(jobId)) return res.status(400).json({ error: '无效的任务 ID' });
  const reportPath = path.join(CHECKLIST_ROOT, jobId, 'report.xlsx');
  if (!fs.existsSync(reportPath)) return res.status(404).json({ error: '检验报告已过期，请重新执行检验' });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader(
    'Content-Disposition',
    `attachment; filename*=UTF-8''${encodeURIComponent('清单合规性检验报告.xlsx')}`
  );
  res.send(fs.readFileSync(reportPath));
});

// ==================== 专家随机打分 (根据 1 份源表随机生成 n-1 份合规打分表) ====================

const RANDOM_SCORE_ROOT = path.join(UPLOADS_DIR, '_toolbox_random_score');

function pruneExpiredRandomScoreJobs() {
  if (!fs.existsSync(RANDOM_SCORE_ROOT)) return;
  const now = Date.now();
  try {
    for (const entry of fs.readdirSync(RANDOM_SCORE_ROOT)) {
      const dir = path.join(RANDOM_SCORE_ROOT, entry);
      const stat = fs.statSync(dir);
      if (stat.isDirectory() && now - stat.mtimeMs > JOB_TTL_MS) {
        fs.rmSync(dir, { recursive: true, force: true });
      }
    }
  } catch {
    // ignore
  }
}

interface RandomScoreManifest {
  files: Array<{ key: string; filename: string; diskName: string; expertName: string }>;
}

// POST /api/toolbox/random-score/run
standaloneToolboxRouter.post(
  '/random-score/run',
  scoreMergeUpload.fields([
    { name: 'template', maxCount: 1 },
    { name: 'namesFile', maxCount: 1 },
  ]),
  async (req, res) => {
    try {
      pruneExpiredRandomScoreJobs();
      const fieldFiles = (req.files || {}) as Record<string, Express.Multer.File[]>;
      const tpl = fieldFiles.template?.[0];
      if (!tpl) return res.status(400).json({ error: '请先上传源专家打分表（.xlsx）' });

      const nameMode = String(req.body?.nameMode || 'auto');
      const maxTotalDiff = req.body?.maxTotalDiff !== undefined ? Number(req.body.maxTotalDiff) : 7;
      let expertNames: string[] | null = null;
      const nExperts = Number(req.body?.nExperts || 5);

      if (nameMode === 'excel') {
        const nf = fieldFiles.namesFile?.[0];
        if (!nf) return res.status(400).json({ error: '请上传专家姓名表（.xlsx）' });
        expertNames = await parseExpertNamesFromExcel(nf.buffer);
        if (expertNames.length === 0) {
          return res.status(400).json({ error: '未在专家姓名表第一列找到有效姓名' });
        }
      } else if (nameMode === 'manual') {
        const raw = String(req.body?.manualNames || '');
        expertNames = raw
          .replace(/[，、；;\r\n]+/g, ',')
          .split(',')
          .map((s) => s.trim())
          .filter(Boolean);
        if (expertNames.length === 0) {
          return res.status(400).json({ error: '请输入至少一位专家姓名' });
        }
      }

      const result = await runRandomScoreGenerator({
        templateFile: { originalName: decodeName(tpl.originalname), buffer: tpl.buffer },
        expertNames,
        nExperts,
        maxTotalDiff,
      });

      const jobId = uuid();
      const jobDir = path.join(RANDOM_SCORE_ROOT, jobId);
      fs.mkdirSync(jobDir, { recursive: true });

      const manifestFiles: RandomScoreManifest['files'] = result.outputs.map((o, idx) => {
        const key = `e${idx + 1}`;
        const diskName = `${key}.xlsx`;
        fs.writeFileSync(path.join(jobDir, diskName), o.buffer);
        return { key, filename: o.filename, diskName, expertName: o.expertName };
      });
      fs.writeFileSync(
        path.join(jobDir, 'manifest.json'),
        JSON.stringify({ files: manifestFiles }, null, 2),
        'utf8'
      );

      res.status(201).json({
        jobId,
        nExperts: result.nExperts,
        maxTotalDiff: result.maxTotalDiff,
        maxSameCount: result.maxSameCount,
        expertNames: result.expertNames,
        sheetSummaries: result.sheetSummaries,
        skippedRows: result.skippedRows,
        totalCheckedGroups: result.totalCheckedGroups,
        violations: result.violations,
        outputs: manifestFiles.map((f) => ({ key: f.key, filename: f.filename, expertName: f.expertName })),
        logs: result.logs,
      });
    } catch (err) {
      res.status(400).json({ error: (err as Error).message });
    }
  }
);

// GET /api/toolbox/random-score/:jobId/download/:fileKey
standaloneToolboxRouter.get('/random-score/:jobId/download/:fileKey', (req, res) => {
  const { jobId, fileKey } = req.params;
  if (!/^[a-zA-Z0-9-]+$/.test(jobId)) return res.status(400).json({ error: '无效的任务 ID' });
  const jobDir = path.join(RANDOM_SCORE_ROOT, jobId);
  const manifestPath = path.join(jobDir, 'manifest.json');
  if (!fs.existsSync(manifestPath)) return res.status(404).json({ error: '结果已过期，请重新生成' });
  const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8')) as RandomScoreManifest;
  const item = manifest.files.find((f) => f.key === fileKey);
  if (!item) return res.status(404).json({ error: '未找到对应打分表' });
  const absPath = path.join(jobDir, item.diskName);
  if (!fs.existsSync(absPath)) return res.status(404).json({ error: '输出文件已失效' });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(item.filename)}`);
  res.send(fs.readFileSync(absPath));
});

// GET /api/toolbox/random-score/:jobId/download-all
standaloneToolboxRouter.get('/random-score/:jobId/download-all', (req, res) => {
  const { jobId } = req.params;
  if (!/^[a-zA-Z0-9-]+$/.test(jobId)) return res.status(400).json({ error: '无效的任务 ID' });
  const jobDir = path.join(RANDOM_SCORE_ROOT, jobId);
  const manifestPath = path.join(jobDir, 'manifest.json');
  if (!fs.existsSync(manifestPath)) return res.status(404).json({ error: '结果已过期，请重新生成' });
  const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8')) as RandomScoreManifest;
  const stamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const zipName = `专家随机打分表_${stamp}.zip`;
  res.setHeader('Content-Type', 'application/zip');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(zipName)}`);

  const archive = archiver('zip', { zlib: { level: 6 } });
  archive.on('error', (err) => {
    if (!res.headersSent) res.status(500).json({ error: `打包失败：${err.message}` });
  });
  archive.pipe(res);
  for (const item of manifest.files) {
    const absPath = path.join(jobDir, item.diskName);
    if (fs.existsSync(absPath)) archive.file(absPath, { name: item.filename });
  }
  archive.finalize();
});
