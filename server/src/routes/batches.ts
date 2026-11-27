import fs from 'node:fs';
import path from 'node:path';
import { Router } from 'express';
import { v4 as uuid } from 'uuid';
import { isAuthEnabled, verifyAdminPassword } from '../auth.js';
import { db, UPLOADS_DIR } from '../db.js';

export const batchesRouter = Router();

// Helper to compute summary metrics for a batch
function enrichBatch(batch: Record<string, unknown>) {
  const batchId = String(batch.id);
  const projCount = (db.prepare('SELECT COUNT(*) as c FROM batch_projects WHERE batch_id = ?').get(batchId) as { c: number }).c;
  const biddingFiles = (db.prepare("SELECT COUNT(*) as c FROM uploaded_files WHERE batch_id = ? AND module = 'bidding'").get(batchId) as { c: number }).c;
  const evalFiles = (db.prepare("SELECT COUNT(*) as c FROM uploaded_files WHERE batch_id = ? AND module = 'evaluation'").get(batchId) as { c: number }).c;
  const awardOutputs = (db.prepare('SELECT COUNT(*) as c FROM award_outputs WHERE batch_id = ?').get(batchId) as { c: number }).c;
  const changeNotices = (db.prepare('SELECT COUNT(*) as c FROM batch_change_notices WHERE batch_id = ?').get(batchId) as { c: number }).c;

  const sopStats = db.prepare(`
    SELECT COUNT(*) as total, SUM(CASE WHEN is_completed = 1 THEN 1 ELSE 0 END) as done
    FROM sop_items WHERE batch_id = ? AND is_template = 0 AND item_type = 'step'
  `).get(batchId) as { total: number; done: number | null };

  const archiveStats = db.prepare(`
    SELECT
      COUNT(*) as total,
      SUM(CASE WHEN status = 'present' THEN 1 ELSE 0 END) as present,
      SUM(CASE WHEN status = 'na' THEN 1 ELSE 0 END) as na,
      SUM(CASE WHEN is_required = 1 AND status != 'na' THEN 1 ELSE 0 END) as req_total,
      SUM(CASE WHEN is_required = 1 AND status = 'present' THEN 1 ELSE 0 END) as req_present,
      SUM(CASE WHEN is_required = 1 AND status = 'missing' THEN 1 ELSE 0 END) as req_missing
    FROM archive_items WHERE batch_id = ? AND is_template = 0
  `).get(batchId) as {
    total: number;
    present: number | null;
    na: number | null;
    req_total: number | null;
    req_present: number | null;
    req_missing: number | null;
  };

  let extraFields = [];
  try {
    extraFields = JSON.parse(String(batch.extra_fields_json || '[]'));
  } catch {
    extraFields = [];
  }

  const budgetRow = db.prepare('SELECT SUM(budget_yuan) as total_yuan FROM batch_projects WHERE batch_id = ?').get(batchId) as {
    total_yuan: number | null;
  };
  const totalBudgetWan =
    budgetRow?.total_yuan && Number.isFinite(budgetRow.total_yuan)
      ? Math.round((budgetRow.total_yuan / 10000) * 100) / 100
      : 0;

  return {
    ...batch,
    extra_fields: extraFields,
    stats: {
      projectCount: projCount,
      totalBudgetWan,
      biddingFileCount: biddingFiles,
      evalFileCount: evalFiles,
      awardOutputCount: awardOutputs,
      changeNoticeCount: changeNotices,
      sopTotal: sopStats.total || 0,
      sopDone: sopStats.done || 0,
      archiveTotal: archiveStats.total || 0,
      archivePresent: archiveStats.present || 0,
      archiveNa: archiveStats.na || 0,
      archiveReqTotal: archiveStats.req_total || 0,
      archiveReqPresent: archiveStats.req_present || 0,
      archiveReqMissing: archiveStats.req_missing || 0,
    },
  };
}

// GET /api/batches - List active (non-trashed) batches with enriched progress metrics
batchesRouter.get('/', (_req, res) => {
  const rows = db
    .prepare("SELECT * FROM batches WHERE deleted_at IS NULL OR deleted_at = '' ORDER BY created_at DESC")
    .all() as Record<string, unknown>[];
  res.json(rows.map(enrichBatch));
});

// GET /api/batches/trash - List soft-deleted batches in recycle bin
batchesRouter.get('/trash', (_req, res) => {
  const rows = db
    .prepare("SELECT * FROM batches WHERE deleted_at IS NOT NULL AND deleted_at != '' ORDER BY deleted_at DESC")
    .all() as Record<string, unknown>[];
  res.json(rows.map(enrichBatch));
});

// GET /api/batches/:id - Get single batch with projects
batchesRouter.get('/:id', (req, res) => {
  const row = db.prepare('SELECT * FROM batches WHERE id = ?').get(req.params.id) as Record<string, unknown> | undefined;
  if (!row) return res.status(404).json({ error: '未找到该招标批次' });
  const projects = db.prepare('SELECT * FROM batch_projects WHERE batch_id = ? ORDER BY seq ASC, created_at ASC').all(req.params.id);
  res.json({ ...enrichBatch(row), projects });
});

// POST /api/batches - Create new batch (auto-copies default SOP & Archive templates)
batchesRouter.post('/', (req, res) => {
  const {
    name,
    code = '',
    project_unit = '',
    announcement_time = '',
    bid_opening_time = '',
    status = '筹备中',
    notes = '',
    extra_fields = [],
    copy_templates = true,
  } = req.body as Record<string, unknown>;

  if (!name || !String(name).trim()) {
    return res.status(400).json({ error: '批次名称不能为空' });
  }

  const id = uuid();
  const now = new Date().toISOString();

  db.prepare(`
    INSERT INTO batches (id, name, code, project_unit, announcement_time, bid_opening_time, status, notes, extra_fields_json, created_at, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).run(
    id,
    String(name).trim(),
    String(code || ''),
    String(project_unit || ''),
    String(announcement_time || ''),
    String(bid_opening_time || ''),
    String(status || '筹备中'),
    String(notes || ''),
    JSON.stringify(Array.isArray(extra_fields) ? extra_fields : []),
    now,
    now
  );

  if (copy_templates) {
    // Copy default SOP templates
    const sopTemplates = db.prepare('SELECT * FROM sop_items WHERE is_template = 1 ORDER BY sort_order ASC').all() as Record<string, unknown>[];
    const insSop = db.prepare(`
      INSERT INTO sop_items (id, batch_id, is_template, item_type, stage_group, priority, title, description, is_completed, completed_at, notes, sort_order)
      VALUES (?, ?, 0, ?, ?, ?, ?, ?, 0, NULL, '', ?)
    `);
    sopTemplates.forEach((t) => {
      insSop.run(uuid(), id, String(t.item_type), String(t.stage_group), String(t.priority), String(t.title), String(t.description), Number(t.sort_order));
    });

    // Copy default Archive templates
    const arcTemplates = db.prepare('SELECT * FROM archive_items WHERE is_template = 1 ORDER BY sort_order ASC').all() as Record<string, unknown>[];
    const insArc = db.prepare(`
      INSERT INTO archive_items (id, batch_id, is_template, sheet_name, seq, doc_name, is_required, required_source, media_type, paper_req, electronic_req, policy_basis, template_source, status, is_manual_marked, linked_file_ids_json, notes, sort_order)
      VALUES (?, ?, 0, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'missing', 0, '[]', '', ?)
    `);
    arcTemplates.forEach((t) => {
      insArc.run(
        uuid(), id, String(t.sheet_name), String(t.seq), String(t.doc_name),
        Number(t.is_required), String(t.required_source), String(t.media_type),
        String(t.paper_req), String(t.electronic_req), String(t.policy_basis),
        String(t.template_source), Number(t.sort_order)
      );
    });
  }

  const created = db.prepare('SELECT * FROM batches WHERE id = ?').get(id) as Record<string, unknown>;
  res.status(201).json(enrichBatch(created));
});

// PUT /api/batches/:id - Update batch basic info
batchesRouter.put('/:id', (req, res) => {
  const existing = db.prepare('SELECT * FROM batches WHERE id = ?').get(req.params.id) as Record<string, unknown> | undefined;
  if (!existing) return res.status(404).json({ error: '未找到该招标批次' });

  const {
    name = existing.name,
    code = existing.code,
    project_unit = existing.project_unit,
    procurement_type = existing.procurement_type || '非公开招标',
    unit_nature = existing.unit_nature || '产业',
    announcement_time = existing.announcement_time,
    doc_fetch_time = existing.doc_fetch_time || '',
    bid_opening_time = existing.bid_opening_time,
    contact_phone = existing.contact_phone || '',
    contact_email = existing.contact_email || '',
    status = existing.status,
    notes = existing.notes,
    extra_fields,
  } = req.body as Record<string, unknown>;

  const extraJson = extra_fields !== undefined
    ? JSON.stringify(Array.isArray(extra_fields) ? extra_fields : [])
    : String(existing.extra_fields_json || '[]');

  db.prepare(`
    UPDATE batches
    SET name = ?, code = ?, project_unit = ?, procurement_type = ?, unit_nature = ?, announcement_time = ?, doc_fetch_time = ?, bid_opening_time = ?, contact_phone = ?, contact_email = ?, status = ?, notes = ?, extra_fields_json = ?, updated_at = ?
    WHERE id = ?
  `).run(
    String(name).trim(),
    String(code || ''),
    String(project_unit || ''),
    String(procurement_type || '非公开招标'),
    String(unit_nature || '产业'),
    String(announcement_time || ''),
    String(doc_fetch_time || ''),
    String(bid_opening_time || ''),
    String(contact_phone || ''),
    String(contact_email || ''),
    String(status || '筹备中'),
    String(notes || ''),
    extraJson,
    new Date().toISOString(),
    req.params.id
  );

  const updated = db.prepare('SELECT * FROM batches WHERE id = ?').get(req.params.id) as Record<string, unknown>;
  res.json(enrichBatch(updated));
});

// POST /api/batches/:id/trash - Soft-delete batch into recycle bin
batchesRouter.post('/:id/trash', (req, res) => {
  const existing = db.prepare('SELECT id FROM batches WHERE id = ?').get(req.params.id) as { id: string } | undefined;
  if (!existing) return res.status(404).json({ error: '未找到该招标批次' });
  const now = new Date().toISOString();
  db.prepare('UPDATE batches SET deleted_at = ?, updated_at = ? WHERE id = ?').run(now, now, req.params.id);
  res.json({ ok: true, deleted_at: now });
});

// POST /api/batches/:id/restore - Restore batch from recycle bin
batchesRouter.post('/:id/restore', (req, res) => {
  const existing = db.prepare('SELECT id FROM batches WHERE id = ?').get(req.params.id) as { id: string } | undefined;
  if (!existing) return res.status(404).json({ error: '未找到该招标批次' });
  const now = new Date().toISOString();
  db.prepare('UPDATE batches SET deleted_at = NULL, updated_at = ? WHERE id = ?').run(now, req.params.id);
  res.json({ ok: true });
});

// DELETE /api/batches/:id - Permanently delete batch (with confirmText + optional admin password check + disk cleanup)
batchesRouter.delete('/:id', (req, res) => {
  const existing = db
    .prepare('SELECT id, code, name FROM batches WHERE id = ?')
    .get(req.params.id) as { id: string; code: string; name: string } | undefined;
  if (!existing) return res.status(404).json({ error: '未找到该招标批次' });

  const body = (req.body || {}) as { confirmText?: string; password?: string };
  const expectedToken = String(existing.code || '').trim() || String(existing.name || '').trim();
  if (expectedToken && String(body.confirmText || '').trim() !== expectedToken) {
    return res.status(400).json({ error: `请输入「${expectedToken}」以确认永久删除` });
  }
  if (isAuthEnabled() && !verifyAdminPassword(body.password)) {
    return res.status(403).json({ error: '工作台访问密码不正确，无法执行永久删除' });
  }

  db.prepare('DELETE FROM batches WHERE id = ?').run(req.params.id);

  // Clean up physical upload directory for this batch
  const batchUploadDir = path.resolve(UPLOADS_DIR, req.params.id);
  const uploadsRoot = path.resolve(UPLOADS_DIR);
  if (batchUploadDir.startsWith(uploadsRoot + path.sep) && fs.existsSync(batchUploadDir)) {
    try {
      fs.rmSync(batchUploadDir, { recursive: true, force: true });
    } catch {
      // non-fatal
    }
  }

  res.json({ ok: true });
});

// --- One-to-many Specific Bidding Projects / Packages CRUD ---

batchesRouter.get('/:batchId/projects', (req, res) => {
  const rows = db.prepare('SELECT * FROM batch_projects WHERE batch_id = ? ORDER BY seq ASC, created_at ASC').all(req.params.batchId);
  res.json(rows);
});

batchesRouter.post('/:batchId/projects', (req, res) => {
  const { name, code = '', budget = '', procurement_method = '公开招标', department = '', notes = '', seq } = req.body as Record<string, unknown>;
  if (!name || !String(name).trim()) {
    return res.status(400).json({ error: '具体项目/标段名称不能为空' });
  }
  const maxSeqRow = db.prepare('SELECT MAX(seq) as m FROM batch_projects WHERE batch_id = ?').get(req.params.batchId) as { m: number | null };
  const nextSeq = seq !== undefined ? Number(seq) : (maxSeqRow.m || 0) + 1;
  const id = uuid();
  const now = new Date().toISOString();

  db.prepare(`
    INSERT INTO batch_projects (id, batch_id, seq, name, code, budget, procurement_method, department, notes, created_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).run(
    id, req.params.batchId, nextSeq, String(name).trim(), String(code || ''),
    String(budget || ''), String(procurement_method || '公开招标'), String(department || ''), String(notes || ''), now
  );

  const created = db.prepare('SELECT * FROM batch_projects WHERE id = ?').get(id);
  res.status(201).json(created);
});

batchesRouter.put('/projects/:projectId', (req, res) => {
  const existing = db.prepare('SELECT * FROM batch_projects WHERE id = ?').get(req.params.projectId) as Record<string, unknown> | undefined;
  if (!existing) return res.status(404).json({ error: '未找到该项目/标段' });

  const body = req.body as Record<string, unknown>;
  const pickStr = (k: string) =>
    body[k] !== undefined ? String(body[k] ?? '').trim() : String(existing[k] ?? '');
  const pickNumOrNull = (k: string): number | null => {
    if (body[k] === undefined) {
      const cur = existing[k];
      return cur !== null && cur !== undefined && Number.isFinite(Number(cur)) ? Number(cur) : null;
    }
    if (body[k] === null || body[k] === '') return null;
    const n = Number(body[k]);
    return Number.isFinite(n) ? n : null;
  };

  const seq = body.seq !== undefined ? Number(body.seq) : Number(existing.seq ?? 1);
  const name = pickStr('name');
  if (!name) return res.status(400).json({ error: '项目名称不能为空' });

  const lotCode = pickStr('lot_code');
  const lotName = pickStr('lot_name');
  const packageNo = pickStr('package_no');
  const reqAppNo = pickStr('req_app_no');
  const quoteMethod = pickStr('quote_method');
  const budgetYuan = pickNumOrNull('budget_yuan');
  const maxPriceRaw = pickStr('max_price_raw');
  const maxPriceYuan = maxPriceRaw ? null : pickNumOrNull('max_price_yuan');
  const budgetText =
    body.budget !== undefined
      ? String(body.budget ?? '').trim()
      : budgetYuan !== null
        ? `${Number((budgetYuan / 10000).toFixed(4))} 万元`
        : String(existing.budget ?? '');
  const code =
    body.code !== undefined
      ? String(body.code ?? '').trim()
      : lotCode
        ? `${lotCode}-${packageNo}`
        : String(existing.code ?? '');

  db.prepare(`
    UPDATE batch_projects
    SET seq = ?, name = ?, code = ?, budget = ?, procurement_method = ?, department = ?, notes = ?,
        lot_code = ?, lot_name = ?, package_no = ?, req_app_no = ?, project_overview = ?,
        quote_method = ?, budget_yuan = ?, max_price_yuan = ?, max_price_raw = ?,
        duration = ?, qualification_req = ?, performance_req = ?, staff_req = ?,
        weight_ratio = ?, contract_version = ?
    WHERE id = ?
  `).run(
    seq,
    name,
    code,
    budgetText,
    pickStr('procurement_method'),
    pickStr('department'),
    pickStr('notes'),
    lotCode,
    lotName,
    packageNo,
    reqAppNo,
    pickStr('project_overview'),
    quoteMethod,
    budgetYuan,
    maxPriceYuan,
    maxPriceRaw,
    pickStr('duration'),
    pickStr('qualification_req'),
    pickStr('performance_req'),
    pickStr('staff_req'),
    pickStr('weight_ratio'),
    pickStr('contract_version'),
    req.params.projectId
  );

  res.json(db.prepare('SELECT * FROM batch_projects WHERE id = ?').get(req.params.projectId));
});

batchesRouter.delete('/projects/:projectId', (req, res) => {
  db.prepare('DELETE FROM batch_projects WHERE id = ?').run(req.params.projectId);
  res.json({ ok: true });
});
