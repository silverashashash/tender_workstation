import fs from 'node:fs';
import path from 'node:path';
import { Router } from 'express';
import multer from 'multer';
import { v4 as uuid } from 'uuid';
import { db, UPLOADS_DIR } from '../db.js';
import { parseSpreadsheetTables } from '../excelParser.js';

export const filesRouter = Router();

const storage = multer.diskStorage({
  destination: (req, _file, cb) => {
    const batchId = String(req.params.batchId || 'common');
    const mod = String(req.body?.module || 'general');
    const dir = path.join(UPLOADS_DIR, batchId, mod);
    fs.mkdirSync(dir, { recursive: true });
    cb(null, dir);
  },
  filename: (_req, file, cb) => {
    const ext = path.extname(file.originalname);
    cb(null, `${Date.now()}-${uuid().slice(0, 8)}${ext}`);
  },
});

const upload = multer({ storage, limits: { fileSize: 50 * 1024 * 1024 } });

function decodeOriginalName(raw: string): string {
  try {
    // Handle latin1-encoded UTF-8 filenames from multer
    return Buffer.from(raw, 'latin1').toString('utf8');
  } catch {
    return raw;
  }
}

// GET /api/files/batch/:batchId?module=bidding
filesRouter.get('/batch/:batchId', (req, res) => {
  const mod = req.query.module ? String(req.query.module) : null;
  const rows = mod
    ? db.prepare('SELECT * FROM uploaded_files WHERE batch_id = ? AND module = ? ORDER BY created_at DESC').all(req.params.batchId, mod)
    : db.prepare('SELECT * FROM uploaded_files WHERE batch_id = ? ORDER BY created_at DESC').all(req.params.batchId);

  const formatted = (rows as Record<string, unknown>[]).map((r) => ({
    ...r,
    parsed: r.parsed_json ? JSON.parse(String(r.parsed_json)) : null,
  }));
  res.json(formatted);
});

// POST /api/files/batch/:batchId/upload
filesRouter.post('/batch/:batchId/upload', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: '请选择要上传的文件' });

  const batchId = req.params.batchId;
  const mod = String(req.body.module || 'bidding');
  const category = String(req.body.category || '其他资料');
  const notes = String(req.body.notes || '');
  const originalName = decodeOriginalName(req.file.originalname);
  const relPath = path.relative(UPLOADS_DIR, req.file.path);

  let parsedJson: string | null = null;
  try {
    const parsed = parseSpreadsheetTables(req.file.path);
    if (parsed) {
      parsedJson = JSON.stringify(parsed);
    }
  } catch {
    // Non-fatal if spreadsheet parse fails
  }

  const id = uuid();
  const now = new Date().toISOString();
  db.prepare(`
    INSERT INTO uploaded_files (id, batch_id, module, category, original_name, stored_path, mime_type, file_size, parsed_json, notes, created_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).run(
    id, batchId, mod, category, originalName, relPath,
    req.file.mimetype || 'application/octet-stream', req.file.size, parsedJson, notes, now
  );

  const created = db.prepare('SELECT * FROM uploaded_files WHERE id = ?').get(id) as Record<string, unknown>;
  res.status(201).json({
    ...created,
    parsed: created.parsed_json ? JSON.parse(String(created.parsed_json)) : null,
  });
});

// GET /api/files/:id/download
filesRouter.get('/:id/download', (req, res) => {
  const row = db.prepare('SELECT * FROM uploaded_files WHERE id = ?').get(req.params.id) as Record<string, unknown> | undefined;
  if (!row) return res.status(404).json({ error: '文件不存在' });
  if (!row.stored_path) {
    return res.status(404).json({ error: '该条记录为预置演示数据，暂无物理文件，请上传真实文件后下载' });
  }
  const absPath = path.join(UPLOADS_DIR, String(row.stored_path));
  if (!fs.existsSync(absPath)) {
    return res.status(404).json({ error: '物理文件已丢失' });
  }
  res.download(absPath, String(row.original_name));
});

// DELETE /api/files/:id
filesRouter.delete('/:id', (req, res) => {
  const row = db.prepare('SELECT * FROM uploaded_files WHERE id = ?').get(req.params.id) as Record<string, unknown> | undefined;
  if (row && row.stored_path) {
    const absPath = path.join(UPLOADS_DIR, String(row.stored_path));
    if (fs.existsSync(absPath)) {
      fs.unlinkSync(absPath);
    }
  }
  db.prepare('DELETE FROM uploaded_files WHERE id = ?').run(req.params.id);
  res.json({ ok: true });
});
