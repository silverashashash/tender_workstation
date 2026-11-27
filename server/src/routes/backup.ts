import fs from 'node:fs';
import path from 'node:path';
import { DatabaseSync } from 'node:sqlite';
import { pipeline } from 'node:stream/promises';
import archiver from 'archiver';
import { Router } from 'express';
import multer from 'multer';
import unzipper from 'unzipper';
import { v4 as uuid } from 'uuid';
import { DATA_DIR, db, DB_PATH, UPLOADS_DIR } from '../db.js';

export const backupRouter = Router();

const BATCH_CHILD_TABLES = [
  'batch_projects',
  'batch_suppliers',
  'eval_bids',
  'eval_line_quotes',
  'eval_rejections',
  'eval_rankings',
  'eval_ip_records',
  'eval_score_checks',
  'uploaded_files',
  'sop_items',
  'award_outputs',
  'archive_items',
  'batch_change_notices',
] as const;

const importTmpDir = path.join(DATA_DIR, 'tmp_imports');
fs.mkdirSync(importTmpDir, { recursive: true });

const uploadBatchZip = multer({
  dest: importTmpDir,
  limits: { fileSize: 1024 * 1024 * 1024 },
});

function toPosixPath(p: string): string {
  return String(p || '').replace(/\\/g, '/').replace(/^\/+/, '');
}

function getTableColumns(tableName: string): string[] {
  const rows = db.prepare(`PRAGMA table_info(${tableName})`).all() as Array<{ name: string }>;
  return rows.map((r) => r.name);
}

function normalizeRowPaths(tableName: string, row: Record<string, unknown>): Record<string, unknown> {
  const copy = { ...row };
  if ((tableName === 'uploaded_files' || tableName === 'award_outputs') && typeof copy.stored_path === 'string') {
    copy.stored_path = toPosixPath(copy.stored_path);
  }
  if (tableName === 'batch_change_notices' && typeof copy.file_stored_path === 'string') {
    copy.file_stored_path = toPosixPath(copy.file_stored_path);
  }
  return copy;
}

// GET /api/backup/export - Export full SQLite DB + uploads folder as ZIP
backupRouter.get('/export', (_req, res) => {
  try {
    db.exec('PRAGMA wal_checkpoint(TRUNCATE);');
  } catch {
    // ignore
  }

  const stamp = new Date().toISOString().slice(0, 10);
  const zipName = `tender-workstation-backup-${stamp}.zip`;
  res.setHeader('Content-Type', 'application/zip');
  res.setHeader('Content-Disposition', `attachment; filename="${zipName}"`);

  const archive = archiver('zip', { zlib: { level: 6 } });
  archive.pipe(res);

  if (fs.existsSync(DB_PATH)) {
    archive.file(DB_PATH, { name: 'workstation.db' });
  }
  if (fs.existsSync(UPLOADS_DIR)) {
    archive.directory(UPLOADS_DIR, 'uploads');
  }
  archive.finalize();
});

// GET /api/backup/export-batch/:batchId - Export a SINGLE batch (DB records + associated files) as a portable ZIP
backupRouter.get('/export-batch/:batchId', (req, res) => {
  const { batchId } = req.params;
  const batch = db.prepare('SELECT * FROM batches WHERE id = ?').get(batchId) as Record<string, unknown> | undefined;
  if (!batch) {
    return res.status(404).json({ error: '未找到该批次' });
  }

  const tables: Record<string, Array<Record<string, unknown>>> = {};
  const counts: Record<string, number> = {};

  for (const table of BATCH_CHILD_TABLES) {
    const query =
      table === 'sop_items' || table === 'archive_items'
        ? `SELECT * FROM ${table} WHERE batch_id = ? AND is_template = 0`
        : `SELECT * FROM ${table} WHERE batch_id = ?`;
    const rows = (db.prepare(query).all(batchId) as Array<Record<string, unknown>>).map((r) =>
      normalizeRowPaths(table, r)
    );
    tables[table] = rows;
    counts[table] = rows.length;
  }

  const stamp = new Date().toISOString().slice(0, 10);
  const safeCode = String(batch.code || '').replace(/[^\w\u4e00-\u9fa5-]/g, '_');
  const safeName = String(batch.name || 'batch')
    .replace(/[^\w\u4e00-\u9fa5-]/g, '_')
    .slice(0, 36);
  const zipName = [safeCode, safeName, '批次包', stamp].filter(Boolean).join('_') + '.zip';

  const manifest = {
    format: 'tender-workstation-batch-v1',
    exportedAt: new Date().toISOString(),
    batch: {
      id: batch.id,
      code: batch.code || '',
      name: batch.name || '',
      status: batch.status || '',
    },
    counts,
  };

  res.setHeader('Content-Type', 'application/zip');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(zipName)}`);

  const archive = archiver('zip', { zlib: { level: 6 } });
  archive.pipe(res);

  archive.append(JSON.stringify(manifest, null, 2), { name: 'manifest.json' });
  archive.append(JSON.stringify({ batch, tables }, null, 2), { name: 'batch-data.json' });

  // Include all files under uploads/<batchId>/
  const addedRelPaths = new Set<string>();
  const batchUploadDir = path.join(UPLOADS_DIR, batchId);
  if (fs.existsSync(batchUploadDir)) {
    archive.directory(batchUploadDir, `uploads/${batchId}`);
  }

  // Also include any referenced stored_path outside uploads/<batchId>/ just in case
  const referencedPaths: string[] = [];
  for (const row of [...(tables.uploaded_files || []), ...(tables.award_outputs || [])]) {
    if (typeof row.stored_path === 'string' && row.stored_path.trim()) {
      referencedPaths.push(toPosixPath(row.stored_path));
    }
  }
  for (const rel of referencedPaths) {
    if (rel.startsWith(`${batchId}/`) || addedRelPaths.has(rel)) continue;
    const abs = path.resolve(UPLOADS_DIR, rel);
    if (!abs.startsWith(path.resolve(UPLOADS_DIR) + path.sep)) continue;
    if (fs.existsSync(abs) && fs.statSync(abs).isFile()) {
      archive.file(abs, { name: `uploads/${rel}` });
      addedRelPaths.add(rel);
    }
  }

  archive.finalize();
});

// POST /api/backup/import-batch - Non-destructively merge/import a single batch ZIP into current DB
backupRouter.post('/import-batch', uploadBatchZip.single('file'), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: '请选择要导入的批次包 (.zip)' });
  }

  const tmpZipPath = req.file.path;
  try {
    const directory = await unzipper.Open.file(tmpZipPath);
    const dataEntry = directory.files.find((f) => toPosixPath(f.path) === 'batch-data.json');
    if (!dataEntry) {
      return res.status(400).json({
        error: '无法识别该压缩包：缺少 batch-data.json。请上传通过「导出当前批次」生成的单批次 ZIP 包。',
      });
    }

    const rawJson = (await dataEntry.buffer()).toString('utf8');
    const parsed = JSON.parse(rawJson) as {
      batch?: Record<string, unknown>;
      tables?: Record<string, Array<Record<string, unknown>>>;
    };

    const incomingBatch = parsed.batch;
    const incomingTables = parsed.tables || {};
    if (!incomingBatch || typeof incomingBatch.id !== 'string' || !incomingBatch.id.trim()) {
      return res.status(400).json({ error: '批次包数据格式无效：缺少批次 ID' });
    }

    const batchId = incomingBatch.id.trim();
    const existingBatch = db.prepare('SELECT id, code, name FROM batches WHERE id = ?').get(batchId) as
      | { id: string; code: string; name: string }
      | undefined;

    const batchCols = getTableColumns('batches');
    const validBatchCols = batchCols.filter((c) => Object.prototype.hasOwnProperty.call(incomingBatch, c));

    db.exec('BEGIN');
    try {
      if (existingBatch) {
        const updateCols = validBatchCols.filter((c) => c !== 'id');
        if (updateCols.length > 0) {
          const setClause = updateCols.map((c) => `${c} = ?`).join(', ');
          const values = updateCols.map((c) => incomingBatch[c] as never);
          db.prepare(`UPDATE batches SET ${setClause} WHERE id = ?`).run(...values, batchId);
        }
        // Clear ONLY this batch's child rows; leave all other batches and global templates untouched
        for (const table of BATCH_CHILD_TABLES) {
          if (table === 'sop_items' || table === 'archive_items') {
            db.prepare(`DELETE FROM ${table} WHERE batch_id = ? AND is_template = 0`).run(batchId);
          } else {
            db.prepare(`DELETE FROM ${table} WHERE batch_id = ?`).run(batchId);
          }
        }
      } else {
        const placeholders = validBatchCols.map(() => '?').join(', ');
        const values = validBatchCols.map((c) => incomingBatch[c] as never);
        db.prepare(`INSERT INTO batches (${validBatchCols.join(', ')}) VALUES (${placeholders})`).run(...values);
      }

      for (const table of BATCH_CHILD_TABLES) {
        const rows = incomingTables[table] || [];
        if (rows.length === 0) continue;
        const tableCols = getTableColumns(table);
        for (const rawRow of rows) {
          const row = normalizeRowPaths(table, { ...rawRow, batch_id: batchId });
          if ((table === 'sop_items' || table === 'archive_items') && !row.id) {
            row.id = uuid();
          }
          const cols = tableCols.filter((c) => Object.prototype.hasOwnProperty.call(row, c));
          if (cols.length === 0) continue;
          const placeholders = cols.map(() => '?').join(', ');
          const values = cols.map((c) => row[c] as never);
          db.prepare(`INSERT OR REPLACE INTO ${table} (${cols.join(', ')}) VALUES (${placeholders})`).run(...values);
        }
      }

      db.exec('COMMIT');
    } catch (txErr) {
      db.exec('ROLLBACK');
      throw txErr;
    }

    // Extract uploads/* entries safely into UPLOADS_DIR
    const resolvedUploadsRoot = path.resolve(UPLOADS_DIR);
    let extractedFiles = 0;
    for (const entry of directory.files) {
      if (entry.type !== 'File') continue;
      const posix = toPosixPath(entry.path);
      if (!posix.startsWith('uploads/')) continue;
      const rel = posix.slice('uploads/'.length);
      if (!rel || rel.includes('..')) continue;
      const destPath = path.resolve(resolvedUploadsRoot, rel);
      if (!destPath.startsWith(resolvedUploadsRoot + path.sep)) continue;
      fs.mkdirSync(path.dirname(destPath), { recursive: true });
      await pipeline(entry.stream(), fs.createWriteStream(destPath));
      extractedFiles++;
    }

    res.status(201).json({
      ok: true,
      action: existingBatch ? 'updated' : 'created',
      batch: {
        id: batchId,
        code: String(incomingBatch.code || ''),
        name: String(incomingBatch.name || ''),
        status: String(incomingBatch.status || ''),
      },
      counts: {
        projects: (incomingTables.batch_projects || []).length,
        files: (incomingTables.uploaded_files || []).length,
        evalBids: (incomingTables.eval_bids || []).length,
        awardOutputs: (incomingTables.award_outputs || []).length,
        extractedFiles,
      },
    });
  } catch (err) {
    res.status(500).json({ error: `导入批次包失败：${(err as Error).message}` });
  } finally {
    if (fs.existsSync(tmpZipPath)) {
      fs.unlinkSync(tmpZipPath);
    }
  }
});

// POST /api/backup/import-full - Restore a FULL workstation backup ZIP (workstation.db + uploads/) in-place
backupRouter.post('/import-full', uploadBatchZip.single('file'), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: '请选择要恢复的整库备份包 (.zip)' });
  }

  const tmpZipPath = req.file.path;
  const tmpDbPath = path.join(importTmpDir, `restore-${uuid()}.db`);
  try {
    const directory = await unzipper.Open.file(tmpZipPath);
    const dbEntry = directory.files.find((f) => {
      const p = toPosixPath(f.path);
      return p === 'workstation.db' || p.endsWith('/workstation.db');
    });
    if (!dbEntry) {
      const isSingleBatch = directory.files.some((f) => toPosixPath(f.path) === 'batch-data.json');
      return res.status(400).json({
        error: isSingleBatch
          ? '该文件是「单批次导出包」，请点击「导入批次包 (.zip)」进行无损合并，而非整库恢复。'
          : '无法识别该压缩包：缺少 workstation.db。请上传通过「导出整库备份」生成的完整备份 ZIP。',
      });
    }

    // Extract workstation.db to a temporary file and copy all tables inside an atomic transaction
    // (avoids Windows EBUSY file-lock issues and requires zero server restart)
    await pipeline(dbEntry.stream(), fs.createWriteStream(tmpDbPath));

    const srcDb = new DatabaseSync(tmpDbPath, { readOnly: true });
    const targetTables = (
      db.prepare("SELECT name FROM sqlite_master WHERE type = 'table' AND name NOT LIKE 'sqlite_%'").all() as Array<{
        name: string;
      }>
    ).map((r) => r.name);
    const srcTablesSet = new Set(
      (
        srcDb.prepare("SELECT name FROM sqlite_master WHERE type = 'table' AND name NOT LIKE 'sqlite_%'").all() as Array<{
          name: string;
        }>
      ).map((r) => r.name)
    );

    const tableRowCounts: Record<string, number> = {};
    try {
      db.exec('PRAGMA foreign_keys = OFF;');
      db.exec('BEGIN;');
      try {
        for (const table of targetTables) {
          if (!srcTablesSet.has(table)) continue;
          const targetCols = getTableColumns(table);
          const srcCols = new Set(
            (srcDb.prepare(`PRAGMA table_info(${table})`).all() as Array<{ name: string }>).map((r) => r.name)
          );
          const commonCols = targetCols.filter((c) => srcCols.has(c));
          if (commonCols.length === 0) continue;

          db.prepare(`DELETE FROM ${table}`).run();
          const rows = srcDb.prepare(`SELECT ${commonCols.join(', ')} FROM ${table}`).all() as Array<
            Record<string, unknown>
          >;
          if (rows.length > 0) {
            const placeholders = commonCols.map(() => '?').join(', ');
            const insertStmt = db.prepare(
              `INSERT INTO ${table} (${commonCols.join(', ')}) VALUES (${placeholders})`
            );
            for (const rawRow of rows) {
              const row = normalizeRowPaths(table, rawRow);
              const vals = commonCols.map((c) => row[c] as never);
              insertStmt.run(...vals);
            }
          }
          tableRowCounts[table] = rows.length;
        }
        db.exec('COMMIT;');
      } catch (txErr) {
        db.exec('ROLLBACK;');
        throw txErr;
      } finally {
        db.exec('PRAGMA foreign_keys = ON;');
      }
    } finally {
      srcDb.close();
    }

    // Replace uploads/ contents cleanly with files from the backup ZIP
    const resolvedUploadsRoot = path.resolve(UPLOADS_DIR);
    fs.mkdirSync(resolvedUploadsRoot, { recursive: true });
    for (const item of fs.readdirSync(resolvedUploadsRoot)) {
      fs.rmSync(path.join(resolvedUploadsRoot, item), { recursive: true, force: true });
    }

    let extractedFiles = 0;
    for (const entry of directory.files) {
      if (entry.type !== 'File') continue;
      const posix = toPosixPath(entry.path);
      const uploadsIdx = posix.indexOf('uploads/');
      if (uploadsIdx < 0) continue;
      const rel = posix.slice(uploadsIdx + 'uploads/'.length);
      if (!rel || rel.includes('..')) continue;
      const destPath = path.resolve(resolvedUploadsRoot, rel);
      if (!destPath.startsWith(resolvedUploadsRoot + path.sep)) continue;
      fs.mkdirSync(path.dirname(destPath), { recursive: true });
      await pipeline(entry.stream(), fs.createWriteStream(destPath));
      extractedFiles++;
    }

    const firstBatch = db
      .prepare('SELECT id FROM batches ORDER BY updated_at DESC, created_at DESC LIMIT 1')
      .get() as { id?: string } | undefined;

    res.status(201).json({
      ok: true,
      firstBatchId: firstBatch?.id || '',
      counts: {
        batches: tableRowCounts.batches || 0,
        projects: tableRowCounts.batch_projects || 0,
        files: tableRowCounts.uploaded_files || 0,
        awardOutputs: tableRowCounts.award_outputs || 0,
        extractedFiles,
      },
    });
  } catch (err) {
    res.status(500).json({ error: `恢复整库备份失败：${(err as Error).message}` });
  } finally {
    if (fs.existsSync(tmpDbPath)) {
      try {
        fs.unlinkSync(tmpDbPath);
      } catch {
        // ignore
      }
    }
    if (fs.existsSync(tmpZipPath)) {
      try {
        fs.unlinkSync(tmpZipPath);
      } catch {
        // ignore
      }
    }
  }
});
