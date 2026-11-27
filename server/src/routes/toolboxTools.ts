import fs from 'node:fs';
import path from 'node:path';
import { Router } from 'express';
import multer from 'multer';
import { v4 as uuid } from 'uuid';
import {
  AWARD_MODES,
  exportAwardAllocationExcel,
  exportAwardTemplate,
  runAwardAllocation,
  type AwardAllocationResult,
  type AwardMode,
} from '../awardTool.js';
import { db, UPLOADS_DIR } from '../db.js';
import {
  exportEconomicRankExcel,
  runEconomicRankCheck,
  type EconomicRankResult,
} from '../economicRankTool.js';
import {
  exportPerformanceExcel,
  exportPerformanceTemplate,
  runPerformanceScore,
  type PerformanceResult,
} from '../performanceTool.js';
import {
  exportPriceScoreExcel,
  exportPriceScoreTemplate,
  runPriceScoreCalculation,
  type PriceScoreResult,
} from '../priceScoreTool.js';
import {
  exportSplitPackExcel,
  exportSplitPackTemplate,
  runSplitPack,
  type SplitPackMode,
  type SplitPackResult,
} from '../splitPackTool.js';

export const toolboxToolsRouter = Router();

const memUpload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 80 * 1024 * 1024, files: 100 },
});

function decodeName(raw: string): string {
  try {
    return Buffer.from(raw, 'latin1').toString('utf8');
  } catch {
    return raw;
  }
}

interface PersistedFileMeta {
  slot: string;
  originalName: string;
  storedPath: string;
  size: number;
}

type ToolboxCheckType =
  | 'price_score'
  | 'performance_score'
  | 'economic_rank'
  | 'award_alloc'
  | 'split_pack';

function saveFiles(
  batchId: string,
  subDir: ToolboxCheckType,
  items: { slot: string; name: string; buffer: Buffer }[]
): PersistedFileMeta[] {
  const targetDir = path.join(UPLOADS_DIR, batchId, 'evaluation', subDir);
  if (fs.existsSync(targetDir)) fs.rmSync(targetDir, { recursive: true, force: true });
  fs.mkdirSync(targetDir, { recursive: true });
  return items.map((item, idx) => {
    const ext = path.extname(item.name) || '.xlsx';
    const diskName = `${String(idx + 1).padStart(3, '0')}-${item.slot}-${uuid().slice(0, 8)}${ext}`;
    const absPath = path.join(targetDir, diskName);
    fs.writeFileSync(absPath, item.buffer);
    return {
      slot: item.slot,
      originalName: item.name,
      storedPath: path.relative(UPLOADS_DIR, absPath),
      size: item.buffer.length,
    };
  });
}

function loadFilesBySlot(metas: PersistedFileMeta[], slot: string): { name: string; buffer: Buffer }[] {
  const out: { name: string; buffer: Buffer }[] = [];
  const missing: string[] = [];
  for (const m of metas) {
    if (m.slot !== slot) continue;
    const abs = path.join(UPLOADS_DIR, m.storedPath);
    if (!fs.existsSync(abs)) {
      missing.push(m.originalName);
      continue;
    }
    out.push({ name: m.originalName, buffer: fs.readFileSync(abs) });
  }
  if (missing.length > 0) {
    throw new Error(`已保存的附件「${missing.join('、')}」已丢失或无法读取，请重新上传后再计算`);
  }
  return out;
}

function getRow(batchId: string, checkType: ToolboxCheckType) {
  const row = db
    .prepare('SELECT * FROM eval_score_checks WHERE batch_id = ? AND check_type = ?')
    .get(batchId, checkType) as Record<string, unknown> | undefined;
  if (!row) return null;
  return {
    files: JSON.parse(String(row.files_json || '[]')) as PersistedFileMeta[],
    config: JSON.parse(String(row.config_json || '{}')) as Record<string, unknown>,
    result: row.result_json ? JSON.parse(String(row.result_json)) : null,
    updated_at: String(row.updated_at || ''),
  };
}

function upsertRow(
  batchId: string,
  checkType: ToolboxCheckType,
  payload: { files: PersistedFileMeta[]; config: Record<string, unknown>; result: unknown }
): string {
  const now = new Date().toISOString();
  db.prepare(`
    INSERT INTO eval_score_checks (batch_id, check_type, files_json, config_json, methods_json, result_json, updated_at)
    VALUES (?, ?, ?, ?, '[]', ?, ?)
    ON CONFLICT(batch_id, check_type) DO UPDATE SET
      files_json = excluded.files_json,
      config_json = excluded.config_json,
      result_json = excluded.result_json,
      updated_at = excluded.updated_at
  `).run(
    batchId,
    checkType,
    JSON.stringify(payload.files),
    JSON.stringify(payload.config),
    payload.result ? JSON.stringify(payload.result) : null,
    now
  );
  return now;
}

function clearRow(batchId: string, checkType: ToolboxCheckType) {
  const targetDir = path.join(UPLOADS_DIR, batchId, 'evaluation', checkType);
  if (fs.existsSync(targetDir)) fs.rmSync(targetDir, { recursive: true, force: true });
  db.prepare('DELETE FROM eval_score_checks WHERE batch_id = ? AND check_type = ?').run(batchId, checkType);
}

function sendExcel(res: import('express').Response, buf: Buffer, filename: string) {
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
  res.send(buf);
}

function safeBatchPrefix(batchId: string): string {
  const batch = db.prepare('SELECT code FROM batches WHERE id = ?').get(batchId) as { code?: string } | undefined;
  return String(batch?.code || 'BATCH').replace(/[^\w\u4e00-\u9fa5-]/g, '_');
}

// ==================== 1. 价格分计算 ====================

toolboxToolsRouter.get('/batch/:batchId/price-score', (req, res) => {
  const { batchId } = req.params;
  const row = getRow(batchId, 'price_score');
  if (!row) {
    return res.json({ files: { quotes: [], rules: null }, config: { round: '' }, result: null, updated_at: '' });
  }
  const quotes = row.files.filter((f) => f.slot === 'quotes').map((f) => ({ originalName: f.originalName, size: f.size }));
  const rule = row.files.find((f) => f.slot === 'rules');
  res.json({
    files: { quotes, rules: rule ? { originalName: rule.originalName, size: rule.size } : null },
    config: row.config,
    result: row.result as PriceScoreResult | null,
    updated_at: row.updated_at,
  });
});

toolboxToolsRouter.post(
  '/batch/:batchId/price-score/run',
  memUpload.fields([
    { name: 'quotes', maxCount: 50 },
    { name: 'rules', maxCount: 1 },
  ]),
  (req, res) => {
    try {
      const { batchId } = req.params;
      const fields = (req.files || {}) as Record<string, Express.Multer.File[]>;
      const existing = getRow(batchId, 'price_score');

      const quoteFiles = fields.quotes?.length
        ? fields.quotes.map((f) => ({ name: decodeName(f.originalname), buffer: f.buffer }))
        : existing
          ? loadFilesBySlot(existing.files, 'quotes')
          : [];
      const ruleFile = fields.rules?.[0]
        ? { name: decodeName(fields.rules[0].originalname), buffer: fields.rules[0].buffer }
        : existing
          ? loadFilesBySlot(existing.files, 'rules')[0] || null
          : null;

      if (!quoteFiles.length) return res.status(400).json({ error: '请先上传报价文件（可多选）' });
      if (!ruleFile) return res.status(400).json({ error: '请先上传一份价格公式参数表' });

      const round = String(req.body?.round ?? '').trim();
      const result = runPriceScoreCalculation(quoteFiles, ruleFile, round);
      const metas = saveFiles(batchId, 'price_score', [
        ...quoteFiles.map((f) => ({ slot: 'quotes', name: f.name, buffer: f.buffer })),
        { slot: 'rules', name: ruleFile.name, buffer: ruleFile.buffer },
      ]);
      const config = { round };
      const updated_at = upsertRow(batchId, 'price_score', { files: metas, config, result });

      res.status(201).json({
        files: {
          quotes: metas.filter((m) => m.slot === 'quotes').map((m) => ({ originalName: m.originalName, size: m.size })),
          rules: { originalName: ruleFile.name, size: ruleFile.buffer.length },
        },
        config,
        result,
        updated_at,
      });
    } catch (err) {
      res.status(400).json({ error: (err as Error).message });
    }
  }
);

toolboxToolsRouter.delete('/batch/:batchId/price-score', (req, res) => {
  clearRow(req.params.batchId, 'price_score');
  res.json({ ok: true });
});

toolboxToolsRouter.get('/batch/:batchId/price-score/export', async (req, res) => {
  try {
    const row = getRow(req.params.batchId, 'price_score');
    const result = row?.result as PriceScoreResult | null;
    if (!result) return res.status(400).json({ error: '尚未生成价格分计算结果' });
    const buf = await exportPriceScoreExcel(result);
    sendExcel(res, buf, `${safeBatchPrefix(req.params.batchId)}_价格分计算.xlsx`);
  } catch (err) {
    res.status(500).json({ error: (err as Error).message });
  }
});

toolboxToolsRouter.get('/batch/:batchId/price-score/template', async (req, res) => {
  try {
    const kind = req.query.kind === 'quotes' ? 'quotes' : 'rules';
    const method = typeof req.query.method === 'string' ? req.query.method : undefined;
    const buf = await exportPriceScoreTemplate(kind, method);
    const fname = kind === 'quotes' ? '报价文件-填写模板.xlsx' : `价格公式参数表-${method || '区间平均价浮动法'}-填写模板.xlsx`;
    sendExcel(res, buf, fname);
  } catch (err) {
    res.status(400).json({ error: (err as Error).message });
  }
});

// ==================== 2. 履约分生成 ====================

toolboxToolsRouter.get('/batch/:batchId/performance-score', (req, res) => {
  const row = getRow(req.params.batchId, 'performance_score');
  if (!row) {
    return res.json({ files: { bids: null, scores: null }, result: null, updated_at: '' });
  }
  const bids = row.files.find((f) => f.slot === 'bids');
  const scores = row.files.find((f) => f.slot === 'scores');
  res.json({
    files: {
      bids: bids ? { originalName: bids.originalName, size: bids.size } : null,
      scores: scores ? { originalName: scores.originalName, size: scores.size } : null,
    },
    result: row.result as PerformanceResult | null,
    updated_at: row.updated_at,
  });
});

toolboxToolsRouter.post(
  '/batch/:batchId/performance-score/run',
  memUpload.fields([
    { name: 'bids', maxCount: 1 },
    { name: 'scores', maxCount: 1 },
  ]),
  (req, res) => {
    try {
      const { batchId } = req.params;
      const fields = (req.files || {}) as Record<string, Express.Multer.File[]>;
      const existing = getRow(batchId, 'performance_score');

      const bidFile = fields.bids?.[0]
        ? { name: decodeName(fields.bids[0].originalname), buffer: fields.bids[0].buffer }
        : existing
          ? loadFilesBySlot(existing.files, 'bids')[0] || null
          : null;
      const scoreFile = fields.scores?.[0]
        ? { name: decodeName(fields.scores[0].originalname), buffer: fields.scores[0].buffer }
        : existing
          ? loadFilesBySlot(existing.files, 'scores')[0] || null
          : null;

      if (!bidFile) return res.status(400).json({ error: '请先上传开标表' });
      if (!scoreFile) return res.status(400).json({ error: '请先上传履约评价表' });

      const result = runPerformanceScore(bidFile, scoreFile);
      const metas = saveFiles(batchId, 'performance_score', [
        { slot: 'bids', name: bidFile.name, buffer: bidFile.buffer },
        { slot: 'scores', name: scoreFile.name, buffer: scoreFile.buffer },
      ]);
      const updated_at = upsertRow(batchId, 'performance_score', { files: metas, config: {}, result });

      res.status(201).json({
        files: {
          bids: { originalName: bidFile.name, size: bidFile.buffer.length },
          scores: { originalName: scoreFile.name, size: scoreFile.buffer.length },
        },
        result,
        updated_at,
      });
    } catch (err) {
      res.status(400).json({ error: (err as Error).message });
    }
  }
);

toolboxToolsRouter.delete('/batch/:batchId/performance-score', (req, res) => {
  clearRow(req.params.batchId, 'performance_score');
  res.json({ ok: true });
});

toolboxToolsRouter.get('/batch/:batchId/performance-score/export', async (req, res) => {
  try {
    const row = getRow(req.params.batchId, 'performance_score');
    const result = row?.result as PerformanceResult | null;
    if (!result) return res.status(400).json({ error: '尚未生成履约分结果' });
    const buf = await exportPerformanceExcel(result);
    sendExcel(res, buf, `${safeBatchPrefix(req.params.batchId)}_履约分.xlsx`);
  } catch (err) {
    res.status(500).json({ error: (err as Error).message });
  }
});

toolboxToolsRouter.get('/batch/:batchId/performance-score/template', async (req, res) => {
  try {
    const kind = req.query.kind === 'bids' ? 'bids' : 'scores';
    const buf = await exportPerformanceTemplate(kind);
    sendExcel(res, buf, kind === 'bids' ? '开标表-填写模板.xlsx' : '履约评价表-填写模板.xlsx');
  } catch (err) {
    res.status(400).json({ error: (err as Error).message });
  }
});

// ==================== 3. 经济排序表核验 ====================

toolboxToolsRouter.get('/batch/:batchId/economic-rank', (req, res) => {
  const row = getRow(req.params.batchId, 'economic_rank');
  if (!row) {
    return res.json({ files: [], result: null, updated_at: '' });
  }
  res.json({
    files: row.files.map((f) => ({ originalName: f.originalName, size: f.size })),
    result: row.result as EconomicRankResult | null,
    updated_at: row.updated_at,
  });
});

toolboxToolsRouter.post('/batch/:batchId/economic-rank/run', memUpload.array('files', 50), (req, res) => {
  try {
    const { batchId } = req.params;
    const uploaded = (req.files as Express.Multer.File[] | undefined) || [];
    if (uploaded.length === 0) return res.status(400).json({ error: '请至少上传一份经济排序表' });

    const files = uploaded.map((f) => ({ name: decodeName(f.originalname), buffer: f.buffer }));
    const result = runEconomicRankCheck(files);
    const metas = saveFiles(
      batchId,
      'economic_rank',
      files.map((f) => ({ slot: 'files', name: f.name, buffer: f.buffer }))
    );
    const updated_at = upsertRow(batchId, 'economic_rank', { files: metas, config: {}, result });

    res.status(201).json({
      files: metas.map((m) => ({ originalName: m.originalName, size: m.size })),
      result,
      updated_at,
    });
  } catch (err) {
    res.status(400).json({ error: (err as Error).message });
  }
});

toolboxToolsRouter.delete('/batch/:batchId/economic-rank', (req, res) => {
  clearRow(req.params.batchId, 'economic_rank');
  res.json({ ok: true });
});

toolboxToolsRouter.get('/batch/:batchId/economic-rank/export', async (req, res) => {
  try {
    const row = getRow(req.params.batchId, 'economic_rank');
    const result = row?.result as EconomicRankResult | null;
    if (!result) return res.status(400).json({ error: '尚未执行经济排序表核验' });
    const buf = await exportEconomicRankExcel(result);
    sendExcel(res, buf, `${safeBatchPrefix(req.params.batchId)}_经济排序核验.xlsx`);
  } catch (err) {
    res.status(500).json({ error: (err as Error).message });
  }
});

// ==================== 4. 授标分配 ====================

toolboxToolsRouter.get('/batch/:batchId/award-alloc', (req, res) => {
  const row = getRow(req.params.batchId, 'award_alloc');
  if (!row) {
    return res.json({
      files: { rankings: [], groups: null },
      config: { mode: '无限授' },
      result: null,
      updated_at: '',
    });
  }
  const rankings = row.files.filter((f) => f.slot === 'rankings').map((f) => ({ originalName: f.originalName, size: f.size }));
  const group = row.files.find((f) => f.slot === 'groups');
  res.json({
    files: { rankings, groups: group ? { originalName: group.originalName, size: group.size } : null },
    config: row.config,
    result: row.result as AwardAllocationResult | null,
    updated_at: row.updated_at,
  });
});

toolboxToolsRouter.post(
  '/batch/:batchId/award-alloc/run',
  memUpload.fields([
    { name: 'files', maxCount: 50 },
    { name: 'groups', maxCount: 1 },
  ]),
  (req, res) => {
    try {
      const { batchId } = req.params;
      const fields = (req.files || {}) as Record<string, Express.Multer.File[]>;
      const existing = getRow(batchId, 'award_alloc');

      const rankingFiles = fields.files?.length
        ? fields.files.map((f) => ({ name: decodeName(f.originalname), buffer: f.buffer }))
        : existing
          ? loadFilesBySlot(existing.files, 'rankings')
          : [];
      const groupFile = fields.groups?.[0]
        ? { name: decodeName(fields.groups[0].originalname), buffer: fields.groups[0].buffer }
        : existing
          ? loadFilesBySlot(existing.files, 'groups')[0] || null
          : null;

      if (!rankingFiles.length) return res.status(400).json({ error: '请先上传技术经济排序表（可多选）' });

      const rawMode = String(req.body?.mode ?? '无限授').trim() as AwardMode;
      const mode: AwardMode = AWARD_MODES.includes(rawMode) ? rawMode : '无限授';

      const result = runAwardAllocation(rankingFiles, mode === '无限授' ? null : groupFile, mode);
      const toSave = [
        ...rankingFiles.map((f) => ({ slot: 'rankings', name: f.name, buffer: f.buffer })),
        ...(groupFile ? [{ slot: 'groups', name: groupFile.name, buffer: groupFile.buffer }] : []),
      ];
      const metas = saveFiles(batchId, 'award_alloc', toSave);
      const config = { mode };
      const updated_at = upsertRow(batchId, 'award_alloc', { files: metas, config, result });

      res.status(201).json({
        files: {
          rankings: metas.filter((m) => m.slot === 'rankings').map((m) => ({ originalName: m.originalName, size: m.size })),
          groups: groupFile ? { originalName: groupFile.name, size: groupFile.buffer.length } : null,
        },
        config,
        result,
        updated_at,
      });
    } catch (err) {
      res.status(400).json({ error: (err as Error).message });
    }
  }
);

toolboxToolsRouter.delete('/batch/:batchId/award-alloc', (req, res) => {
  clearRow(req.params.batchId, 'award_alloc');
  res.json({ ok: true });
});

toolboxToolsRouter.get('/batch/:batchId/award-alloc/export', async (req, res) => {
  try {
    const row = getRow(req.params.batchId, 'award_alloc');
    const result = row?.result as AwardAllocationResult | null;
    if (!result) return res.status(400).json({ error: '尚未生成授标分配结果' });
    const buf = await exportAwardAllocationExcel(result);
    sendExcel(res, buf, `${safeBatchPrefix(req.params.batchId)}_授标结果.xlsx`);
  } catch (err) {
    res.status(500).json({ error: (err as Error).message });
  }
});

toolboxToolsRouter.get('/batch/:batchId/award-alloc/template', async (req, res) => {
  try {
    const kind = req.query.kind === 'groups' ? 'groups' : 'ranking';
    const buf = await exportAwardTemplate(kind);
    sendExcel(res, buf, kind === 'groups' ? '限授组配置表-填写模板.xlsx' : '技术经济排序表-填写模板.xlsx');
  } catch (err) {
    res.status(400).json({ error: (err as Error).message });
  }
});

// ==================== 5. 智能分包 ====================

toolboxToolsRouter.get('/batch/:batchId/split-pack', (req, res) => {
  const row = getRow(req.params.batchId, 'split_pack');
  if (!row) {
    return res.json({
      files: { materials: null, rules: null },
      config: { mode: '比例', weightStep: '1', lengthStep: '1' },
      result: null,
      updated_at: '',
    });
  }
  const mat = row.files.find((f) => f.slot === 'materials');
  const rule = row.files.find((f) => f.slot === 'rules');
  res.json({
    files: {
      materials: mat ? { originalName: mat.originalName, size: mat.size } : null,
      rules: rule ? { originalName: rule.originalName, size: rule.size } : null,
    },
    config: row.config,
    result: row.result as SplitPackResult | null,
    updated_at: row.updated_at,
  });
});

toolboxToolsRouter.post(
  '/batch/:batchId/split-pack/run',
  memUpload.fields([
    { name: 'materials', maxCount: 1 },
    { name: 'rules', maxCount: 1 },
  ]),
  (req, res) => {
    try {
      const { batchId } = req.params;
      const fields = (req.files || {}) as Record<string, Express.Multer.File[]>;
      const existing = getRow(batchId, 'split_pack');

      const matFile = fields.materials?.[0]
        ? { name: decodeName(fields.materials[0].originalname), buffer: fields.materials[0].buffer }
        : existing
          ? loadFilesBySlot(existing.files, 'materials')[0] || null
          : null;
      const ruleFile = fields.rules?.[0]
        ? { name: decodeName(fields.rules[0].originalname), buffer: fields.rules[0].buffer }
        : existing
          ? loadFilesBySlot(existing.files, 'rules')[0] || null
          : null;

      if (!matFile) return res.status(400).json({ error: '请先上传物料表' });
      if (!ruleFile) return res.status(400).json({ error: '请先上传分包规则表' });

      const mode: SplitPackMode = req.body?.mode === '金额' ? '金额' : '比例';
      const weightStep = req.body?.weightStep === undefined ? '1' : String(req.body.weightStep).trim();
      const lengthStep = req.body?.lengthStep === undefined ? '1' : String(req.body.lengthStep).trim();

      const result = runSplitPack(matFile, ruleFile, mode, weightStep, lengthStep);
      const metas = saveFiles(batchId, 'split_pack', [
        { slot: 'materials', name: matFile.name, buffer: matFile.buffer },
        { slot: 'rules', name: ruleFile.name, buffer: ruleFile.buffer },
      ]);
      const config = { mode, weightStep, lengthStep };
      const updated_at = upsertRow(batchId, 'split_pack', { files: metas, config, result });

      res.status(201).json({
        files: {
          materials: { originalName: matFile.name, size: matFile.buffer.length },
          rules: { originalName: ruleFile.name, size: ruleFile.buffer.length },
        },
        config,
        result,
        updated_at,
      });
    } catch (err) {
      res.status(400).json({ error: (err as Error).message });
    }
  }
);

toolboxToolsRouter.delete('/batch/:batchId/split-pack', (req, res) => {
  clearRow(req.params.batchId, 'split_pack');
  res.json({ ok: true });
});

toolboxToolsRouter.get('/batch/:batchId/split-pack/export', async (req, res) => {
  try {
    const row = getRow(req.params.batchId, 'split_pack');
    const result = row?.result as SplitPackResult | null;
    if (!result) return res.status(400).json({ error: '尚未生成智能分包结果' });
    const buf = await exportSplitPackExcel(result);
    sendExcel(res, buf, `${safeBatchPrefix(req.params.batchId)}_分包结果数据.xlsx`);
  } catch (err) {
    res.status(500).json({ error: (err as Error).message });
  }
});

toolboxToolsRouter.get('/batch/:batchId/split-pack/template', async (req, res) => {
  try {
    const kind = req.query.kind === 'rules' ? 'rules' : 'materials';
    const buf = await exportSplitPackTemplate(kind);
    sendExcel(res, buf, kind === 'rules' ? '分包规则表-填写模板.xlsx' : '物料表-填写模板.xlsx');
  } catch (err) {
    res.status(400).json({ error: (err as Error).message });
  }
});
