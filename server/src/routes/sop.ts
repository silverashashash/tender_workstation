import fs from 'node:fs';
import path from 'node:path';
import { Router } from 'express';
import multer from 'multer';
import { v4 as uuid } from 'uuid';
import { db, UPLOADS_DIR } from '../db.js';

export const sopRouter = Router();

const SOP_GUIDES_DIR = path.join(UPLOADS_DIR, 'sop-guides');
fs.mkdirSync(SOP_GUIDES_DIR, { recursive: true });

function decodeOriginalName(raw: string): string {
  try {
    return Buffer.from(raw, 'latin1').toString('utf8');
  } catch {
    return raw;
  }
}

const guideUpload = multer({
  storage: multer.diskStorage({
    destination: (_req, _file, cb) => {
      fs.mkdirSync(SOP_GUIDES_DIR, { recursive: true });
      cb(null, SOP_GUIDES_DIR);
    },
    filename: (_req, file, cb) => {
      const ext = path.extname(decodeOriginalName(file.originalname)).toLowerCase() || '.bin';
      cb(null, `${Date.now()}-${uuid().slice(0, 8)}${ext}`);
    },
  }),
  limits: { fileSize: 50 * 1024 * 1024 },
});

interface SopGuideImage {
  filename: string;
  url: string;
  originalName: string;
  caption?: string;
}

function formatGuideRow(row: Record<string, unknown>) {
  let images: SopGuideImage[] = [];
  try {
    const parsed = JSON.parse(String(row.images_json || '[]'));
    if (Array.isArray(parsed)) images = parsed;
  } catch {
    images = [];
  }
  return {
    id: String(row.id),
    stage_group: String(row.stage_group || '内网开标'),
    title: String(row.title || ''),
    source_type: String(row.source_type || 'markdown') as 'markdown' | 'pdf',
    content_md: String(row.content_md || ''),
    images,
    file_original_name: String(row.file_original_name || ''),
    file_mime_type: String(row.file_mime_type || ''),
    file_size: Number(row.file_size || 0),
    has_file: Boolean(row.file_stored_path),
    sort_order: Number(row.sort_order || 0),
    created_at: String(row.created_at || ''),
    updated_at: String(row.updated_at || ''),
  };
}

// ==================== Global System Operation SOP Guides (Text / Image / MD / PDF) ====================

// GET /api/sop/guides
sopRouter.get('/guides', (_req, res) => {
  const rows = db
    .prepare('SELECT * FROM sop_guides ORDER BY sort_order ASC, updated_at DESC')
    .all() as Record<string, unknown>[];
  res.json(rows.map(formatGuideRow));
});

// POST /api/sop/guides - Create manual text / Markdown + images SOP guide
sopRouter.post('/guides', (req, res) => {
  const {
    stage_group = '内网开标',
    title,
    content_md = '',
    images = [],
  } = req.body as Record<string, unknown>;

  if (!title || !String(title).trim()) {
    return res.status(400).json({ error: '指南标题不能为空' });
  }

  const maxRow = db.prepare('SELECT MAX(sort_order) as m FROM sop_guides').get() as { m: number | null };
  const id = uuid();
  const now = new Date().toISOString();
  const safeImages = Array.isArray(images) ? images : [];

  db.prepare(`
    INSERT INTO sop_guides (id, stage_group, title, source_type, content_md, images_json, file_original_name, file_stored_path, file_mime_type, file_size, sort_order, created_at, updated_at)
    VALUES (?, ?, ?, 'markdown', ?, ?, '', '', '', 0, ?, ?, ?)
  `).run(
    id,
    String(stage_group || '内网开标'),
    String(title).trim(),
    String(content_md || ''),
    JSON.stringify(safeImages),
    (maxRow.m || 0) + 1,
    now,
    now
  );

  const created = db.prepare('SELECT * FROM sop_guides WHERE id = ?').get(id) as Record<string, unknown>;
  res.status(201).json(formatGuideRow(created));
});

// POST /api/sop/guides/upload-image - Upload screenshot/image for SOP guides
sopRouter.post('/guides/upload-image', guideUpload.single('image'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: '请选择要上传的图片' });
  const originalName = decodeOriginalName(req.file.originalname);
  const ext = path.extname(originalName).toLowerCase();
  const allowed = new Set(['.png', '.jpg', '.jpeg', '.webp', '.gif', '.bmp', '.svg']);
  if (!allowed.has(ext) && !req.file.mimetype.startsWith('image/')) {
    try {
      fs.unlinkSync(req.file.path);
    } catch {
      // ignore
    }
    return res.status(400).json({ error: '仅支持上传 PNG / JPG / WebP / GIF 图片文件' });
  }

  const filename = path.basename(req.file.path);
  res.status(201).json({
    filename,
    url: `/api/sop/guides/media/${encodeURIComponent(filename)}`,
    originalName,
    caption: '',
  });
});

// GET /api/sop/guides/media/:filename - Serve uploaded SOP images
sopRouter.get('/guides/media/:filename', (req, res) => {
  const safeName = path.basename(req.params.filename);
  const absPath = path.join(SOP_GUIDES_DIR, safeName);
  if (!absPath.startsWith(SOP_GUIDES_DIR + path.sep) || !fs.existsSync(absPath)) {
    return res.status(404).json({ error: '图片不存在' });
  }
  res.sendFile(absPath);
});

// POST /api/sop/guides/import - Import a .md / .markdown / .txt or .pdf file directly
sopRouter.post('/guides/import', guideUpload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: '请选择要导入的 MD 或 PDF 文件' });

  const originalName = decodeOriginalName(req.file.originalname);
  const ext = path.extname(originalName).toLowerCase();
  const stageGroup = String(req.body?.stage_group || '内网开标');
  const customTitle = String(req.body?.title || '').trim();
  const stem = path.basename(originalName, ext).trim() || originalName;

  const isMd = ext === '.md' || ext === '.markdown' || ext === '.txt';
  const isPdf = ext === '.pdf';

  if (!isMd && !isPdf) {
    try {
      fs.unlinkSync(req.file.path);
    } catch {
      // ignore
    }
    return res.status(400).json({ error: '仅支持导入 Markdown (.md / .markdown / .txt) 或 PDF (.pdf) 文件' });
  }

  const maxRow = db.prepare('SELECT MAX(sort_order) as m FROM sop_guides').get() as { m: number | null };
  const id = uuid();
  const now = new Date().toISOString();

  if (isMd) {
    let contentMd = '';
    try {
      contentMd = fs.readFileSync(req.file.path, 'utf8');
    } catch (err) {
      return res.status(400).json({ error: `读取 Markdown 文件失败：${(err as Error).message}` });
    } finally {
      try {
        fs.unlinkSync(req.file.path);
      } catch {
        // ignore
      }
    }

    // Auto-extract first H1 heading if no custom title provided
    let title = customTitle || stem;
    if (!customTitle) {
      const h1 = contentMd.match(/^\s*#\s+(.+?)\s*$/m);
      if (h1 && h1[1].trim()) {
        title = h1[1].trim();
      }
    }

    db.prepare(`
      INSERT INTO sop_guides (id, stage_group, title, source_type, content_md, images_json, file_original_name, file_stored_path, file_mime_type, file_size, sort_order, created_at, updated_at)
      VALUES (?, ?, ?, 'markdown', ?, '[]', ?, '', 'text/markdown', ?, ?, ?, ?)
    `).run(id, stageGroup, title, contentMd, originalName, req.file.size, (maxRow.m || 0) + 1, now, now);
  } else {
    const relPath = path.relative(UPLOADS_DIR, req.file.path).replace(/\\/g, '/');
    const title = customTitle || stem;
    db.prepare(`
      INSERT INTO sop_guides (id, stage_group, title, source_type, content_md, images_json, file_original_name, file_stored_path, file_mime_type, file_size, sort_order, created_at, updated_at)
      VALUES (?, ?, ?, 'pdf', '', '[]', ?, ?, 'application/pdf', ?, ?, ?, ?)
    `).run(id, stageGroup, title, originalName, relPath, req.file.size, (maxRow.m || 0) + 1, now, now);
  }

  const created = db.prepare('SELECT * FROM sop_guides WHERE id = ?').get(id) as Record<string, unknown>;
  res.status(201).json(formatGuideRow(created));
});

// GET /api/sop/guides/:id/file - Stream PDF inline (or ?download=1)
sopRouter.get('/guides/:id/file', (req, res) => {
  const row = db.prepare('SELECT * FROM sop_guides WHERE id = ?').get(req.params.id) as Record<string, unknown> | undefined;
  if (!row || !row.file_stored_path) {
    return res.status(404).json({ error: '未找到对应的 PDF 文件' });
  }
  const absPath = path.resolve(UPLOADS_DIR, String(row.file_stored_path));
  if (!absPath.startsWith(path.resolve(UPLOADS_DIR) + path.sep) || !fs.existsSync(absPath)) {
    return res.status(404).json({ error: '物理文件已丢失' });
  }
  const fileName = String(row.file_original_name || `${row.title || 'guide'}.pdf`);
  const disposition = req.query.download === '1' ? 'attachment' : 'inline';
  res.setHeader('Content-Type', String(row.file_mime_type || 'application/pdf'));
  res.setHeader('Content-Disposition', `${disposition}; filename*=UTF-8''${encodeURIComponent(fileName)}`);
  res.sendFile(absPath);
});

// PUT /api/sop/guides/:id - Update SOP guide
sopRouter.put('/guides/:id', (req, res) => {
  const existing = db.prepare('SELECT * FROM sop_guides WHERE id = ?').get(req.params.id) as Record<string, unknown> | undefined;
  if (!existing) return res.status(404).json({ error: '操作指南不存在' });

  const {
    stage_group = existing.stage_group,
    title = existing.title,
    content_md = existing.content_md,
    images,
  } = req.body as Record<string, unknown>;

  const nextTitle = String(title || '').trim();
  if (!nextTitle) return res.status(400).json({ error: '指南标题不能为空' });

  const nextImagesJson =
    images !== undefined
      ? JSON.stringify(Array.isArray(images) ? images : [])
      : String(existing.images_json || '[]');

  db.prepare(`
    UPDATE sop_guides
    SET stage_group = ?, title = ?, content_md = ?, images_json = ?, updated_at = ?
    WHERE id = ?
  `).run(
    String(stage_group || '内网开标'),
    nextTitle,
    String(content_md ?? ''),
    nextImagesJson,
    new Date().toISOString(),
    req.params.id
  );

  const updated = db.prepare('SELECT * FROM sop_guides WHERE id = ?').get(req.params.id) as Record<string, unknown>;
  res.json(formatGuideRow(updated));
});

// DELETE /api/sop/guides/:id - Delete SOP guide and associated stored files
sopRouter.delete('/guides/:id', (req, res) => {
  const existing = db.prepare('SELECT * FROM sop_guides WHERE id = ?').get(req.params.id) as Record<string, unknown> | undefined;
  if (existing) {
    if (existing.file_stored_path) {
      const absPdf = path.resolve(UPLOADS_DIR, String(existing.file_stored_path));
      if (absPdf.startsWith(path.resolve(UPLOADS_DIR) + path.sep) && fs.existsSync(absPdf)) {
        try {
          fs.unlinkSync(absPdf);
        } catch {
          // ignore
        }
      }
    }
    try {
      const imgs = JSON.parse(String(existing.images_json || '[]')) as SopGuideImage[];
      if (Array.isArray(imgs)) {
        for (const img of imgs) {
          if (!img?.filename) continue;
          const absImg = path.join(SOP_GUIDES_DIR, path.basename(img.filename));
          if (absImg.startsWith(SOP_GUIDES_DIR + path.sep) && fs.existsSync(absImg)) {
            fs.unlinkSync(absImg);
          }
        }
      }
    } catch {
      // ignore
    }
  }
  db.prepare('DELETE FROM sop_guides WHERE id = ?').run(req.params.id);
  res.json({ ok: true });
});

// ==================== Per-Batch SOP Checklist Items ====================

// GET /api/sop/batch/:batchId
sopRouter.get('/batch/:batchId', (req, res) => {
  const rows = db.prepare(`
    SELECT * FROM sop_items
    WHERE batch_id = ? AND is_template = 0
    ORDER BY item_type DESC, sort_order ASC
  `).all(req.params.batchId);
  res.json(rows);
});

// POST /api/sop/batch/:batchId
sopRouter.post('/batch/:batchId', (req, res) => {
  const {
    item_type = 'step',
    stage_group = '开标前准备',
    priority = 'normal',
    title,
    description = '',
    notes = '',
    sync_to_template = false,
  } = req.body as Record<string, unknown>;

  if (!title || !String(title).trim()) {
    return res.status(400).json({ error: '标题不能为空' });
  }

  const maxRow = db.prepare(`
    SELECT MAX(sort_order) as m FROM sop_items WHERE batch_id = ? AND item_type = ?
  `).get(req.params.batchId, String(item_type)) as { m: number | null };

  const id = uuid();
  db.prepare(`
    INSERT INTO sop_items (id, batch_id, is_template, item_type, stage_group, priority, title, description, is_completed, completed_at, notes, sort_order)
    VALUES (?, ?, 0, ?, ?, ?, ?, ?, 0, NULL, ?, ?)
  `).run(
    id, req.params.batchId, String(item_type), String(stage_group), String(priority),
    String(title).trim(), String(description || ''), String(notes || ''), (maxRow.m || 0) + 1
  );

  // Sync new step/notice to global default template so future batches inherit it automatically
  if (sync_to_template) {
    const tplMax = db.prepare(`
      SELECT MAX(sort_order) as m FROM sop_items WHERE is_template = 1 AND item_type = ?
    `).get(String(item_type)) as { m: number | null };
    db.prepare(`
      INSERT INTO sop_items (id, batch_id, is_template, item_type, stage_group, priority, title, description, is_completed, completed_at, notes, sort_order)
      VALUES (?, NULL, 1, ?, ?, ?, ?, ?, 0, NULL, '', ?)
    `).run(
      uuid(), String(item_type), String(stage_group), String(priority),
      String(title).trim(), String(description || ''), (tplMax.m || 0) + 1
    );
  }

  res.status(201).json(db.prepare('SELECT * FROM sop_items WHERE id = ?').get(id));
});

// PUT /api/sop/:id
sopRouter.put('/:id', (req, res) => {
  const existing = db.prepare('SELECT * FROM sop_items WHERE id = ?').get(req.params.id) as Record<string, unknown> | undefined;
  if (!existing) return res.status(404).json({ error: '条目不存在' });

  const {
    stage_group = existing.stage_group,
    priority = existing.priority,
    title = existing.title,
    description = existing.description,
    is_completed,
    notes = existing.notes,
  } = req.body as Record<string, unknown>;

  let nextCompleted = Number(existing.is_completed);
  let completedAt = existing.completed_at ? String(existing.completed_at) : null;

  if (is_completed !== undefined) {
    nextCompleted = is_completed ? 1 : 0;
    completedAt = nextCompleted ? new Date().toISOString().slice(0, 16).replace('T', ' ') : null;
  }

  db.prepare(`
    UPDATE sop_items
    SET stage_group = ?, priority = ?, title = ?, description = ?, is_completed = ?, completed_at = ?, notes = ?
    WHERE id = ?
  `).run(
    String(stage_group), String(priority), String(title).trim(), String(description || ''),
    nextCompleted, completedAt, String(notes || ''), req.params.id
  );

  res.json(db.prepare('SELECT * FROM sop_items WHERE id = ?').get(req.params.id));
});

// DELETE /api/sop/:id
sopRouter.delete('/:id', (req, res) => {
  db.prepare('DELETE FROM sop_items WHERE id = ?').run(req.params.id);
  res.json({ ok: true });
});

// POST /api/sop/batch/:batchId/save-as-template
sopRouter.post('/batch/:batchId/save-as-template', (req, res) => {
  const items = db.prepare('SELECT * FROM sop_items WHERE batch_id = ? AND is_template = 0 ORDER BY sort_order ASC').all(req.params.batchId) as Record<string, unknown>[];
  db.prepare('DELETE FROM sop_items WHERE is_template = 1').run();
  const ins = db.prepare(`
    INSERT INTO sop_items (id, batch_id, is_template, item_type, stage_group, priority, title, description, is_completed, completed_at, notes, sort_order)
    VALUES (?, NULL, 1, ?, ?, ?, ?, ?, 0, NULL, '', ?)
  `);
  items.forEach((it, idx) => {
    ins.run(uuid(), String(it.item_type), String(it.stage_group), String(it.priority), String(it.title), String(it.description), idx + 1);
  });
  res.json({ ok: true, count: items.length });
});
