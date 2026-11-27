import fs from 'node:fs';
import path from 'node:path';
import archiver from 'archiver';
import ExcelJS from 'exceljs';
import { Router } from 'express';
import multer from 'multer';
import { v4 as uuid } from 'uuid';
import { db, UPLOADS_DIR } from '../db.js';
import { parseArchiveCatalogExcel, type ParsedArchiveRow } from '../excelParser.js';

export const archiveRouter = Router();

const tempStorage = multer.diskStorage({
  destination: (_req, _file, cb) => {
    const dir = path.join(UPLOADS_DIR, '_tmp');
    fs.mkdirSync(dir, { recursive: true });
    cb(null, dir);
  },
  filename: (_req, file, cb) => {
    const ext = path.extname(file.originalname);
    cb(null, `${Date.now()}-${uuid().slice(0, 8)}${ext}`);
  },
});

const itemFileStorage = multer.diskStorage({
  destination: (req, _file, cb) => {
    const batchId = String(req.body?.batchId || 'archive');
    const dir = path.join(UPLOADS_DIR, batchId, 'archive');
    fs.mkdirSync(dir, { recursive: true });
    cb(null, dir);
  },
  filename: (_req, file, cb) => {
    const ext = path.extname(file.originalname);
    cb(null, `${Date.now()}-${uuid().slice(0, 8)}${ext}`);
  },
});

const uploadTemp = multer({ storage: tempStorage, limits: { fileSize: 30 * 1024 * 1024 } });
const uploadItemFile = multer({ storage: itemFileStorage, limits: { fileSize: 50 * 1024 * 1024 } });

function decodeName(raw: string): string {
  try {
    return Buffer.from(raw, 'latin1').toString('utf8');
  } catch {
    return raw;
  }
}

// GET /api/archive/batch/:batchId
archiveRouter.get('/batch/:batchId', (req, res) => {
  const batchId = req.params.batchId;
  const rows = db.prepare(`
    SELECT * FROM archive_items
    WHERE batch_id = ? AND is_template = 0
    ORDER BY sheet_name ASC, sort_order ASC
  `).all(batchId) as Record<string, unknown>[];

  const batchFiles = db.prepare(`
    SELECT id, module, category, original_name, stored_path, file_size, created_at
    FROM uploaded_files
    WHERE batch_id = ?
    ORDER BY created_at DESC
  `).all(batchId) as Array<Record<string, unknown>>;

  const fileMap = new Map(batchFiles.map((f) => [String(f.id), f]));

  const items = rows.map((r) => {
    let linkedIds: string[] = [];
    try {
      linkedIds = JSON.parse(String(r.linked_file_ids_json || '[]'));
    } catch {
      linkedIds = [];
    }
    const linkedFiles = linkedIds.map((fid) => fileMap.get(fid)).filter(Boolean);
    return {
      ...r,
      sheet_name: String(r.sheet_name || '归档文件'),
      status: String(r.status || 'missing'),
      is_required: Boolean(r.is_required),
      is_manual_marked: Boolean(r.is_manual_marked),
      linked_file_ids: linkedIds,
      linked_files: linkedFiles,
    };
  });

  const sheetNames = Array.from(new Set(items.map((i) => String(i.sheet_name))));
  const activeItems = items.filter((i) => i.status !== 'na');
  const requiredActive = activeItems.filter((i) => i.is_required);
  const optionalActive = activeItems.filter((i) => !i.is_required);

  res.json({
    sheetNames,
    items,
    availableBatchFiles: batchFiles,
    summary: {
      total: items.length,
      activeTotal: activeItems.length,
      present: activeItems.filter((i) => i.status === 'present').length,
      missing: activeItems.filter((i) => i.status === 'missing').length,
      na: items.filter((i) => i.status === 'na').length,
      requiredTotal: requiredActive.length,
      requiredPresent: requiredActive.filter((i) => i.status === 'present').length,
      requiredMissing: requiredActive.filter((i) => i.status === 'missing').length,
      optionalTotal: optionalActive.length,
      optionalPresent: optionalActive.filter((i) => i.status === 'present').length,
    },
  });
});

// POST /api/archive/batch/:batchId/import-preview
archiveRouter.post('/batch/:batchId/import-preview', uploadTemp.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: '请上传 Excel 归档清单文件' });
  try {
    const parsed = await parseArchiveCatalogExcel(req.file.path);
    fs.unlinkSync(req.file.path);
    res.json({
      filename: decodeName(req.file.originalname),
      ...parsed,
    });
  } catch (err) {
    if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    res.status(500).json({ error: `解析归档目录失败: ${(err as Error).message}` });
  }
});

// POST /api/archive/batch/:batchId/import-commit
archiveRouter.post('/batch/:batchId/import-commit', (req, res) => {
  const batchId = req.params.batchId;
  const { items = [], replace = true, saveAsTemplate = false } = req.body as {
    items?: ParsedArchiveRow[];
    replace?: boolean;
    saveAsTemplate?: boolean;
  };

  if (!Array.isArray(items) || items.length === 0) {
    return res.status(400).json({ error: '没有可导入的归档条目' });
  }

  if (replace) {
    db.prepare('DELETE FROM archive_items WHERE batch_id = ? AND is_template = 0').run(batchId);
  }

  const ins = db.prepare(`
    INSERT INTO archive_items (id, batch_id, is_template, sheet_name, seq, doc_name, is_required, required_source, media_type, paper_req, electronic_req, policy_basis, template_source, status, is_manual_marked, linked_file_ids_json, notes, sort_order)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '[]', ?, ?)
  `);

  items.forEach((it, idx) => {
    ins.run(
      uuid(),
      batchId,
      0,
      it.sheetName || '归档文件',
      String(it.seq ?? idx + 1),
      String(it.docName || '').trim(),
      it.isRequired ? 1 : 0,
      it.requiredSource || 'manual',
      it.mediaType || '纸质+电子',
      it.paperReq || '',
      it.electronicReq || '',
      it.policyBasis || '',
      it.templateSource || '',
      it.status || 'missing',
      it.status === 'present' ? 1 : 0,
      it.notes || '',
      idx + 1
    );
  });

  if (saveAsTemplate) {
    db.prepare('DELETE FROM archive_items WHERE is_template = 1').run();
    items.forEach((it, idx) => {
      ins.run(
        uuid(),
        null,
        1,
        it.sheetName || '归档文件',
        String(it.seq ?? idx + 1),
        String(it.docName || '').trim(),
        it.isRequired ? 1 : 0,
        it.requiredSource || 'manual',
        it.mediaType || '纸质+电子',
        it.paperReq || '',
        it.electronicReq || '',
        it.policyBasis || '',
        it.templateSource || '',
        'missing',
        0,
        '',
        idx + 1
      );
    });
  }

  res.json({ ok: true, count: items.length });
});

// POST /api/archive/batch/:batchId/items - Add manual item
archiveRouter.post('/batch/:batchId/items', (req, res) => {
  const batchId = req.params.batchId;
  const {
    sheet_name = '归档文件',
    seq = '',
    doc_name,
    is_required = true,
    media_type = '纸质+电子',
    paper_req = '',
    electronic_req = '',
    policy_basis = '',
    template_source = '',
    notes = '',
  } = req.body as Record<string, unknown>;

  if (!doc_name || !String(doc_name).trim()) {
    return res.status(400).json({ error: '资料名称不能为空' });
  }

  const maxRow = db.prepare('SELECT MAX(sort_order) as m FROM archive_items WHERE batch_id = ?').get(batchId) as { m: number | null };
  const id = uuid();
  db.prepare(`
    INSERT INTO archive_items (id, batch_id, is_template, sheet_name, seq, doc_name, is_required, required_source, media_type, paper_req, electronic_req, policy_basis, template_source, status, is_manual_marked, linked_file_ids_json, notes, sort_order)
    VALUES (?, ?, 0, ?, ?, ?, ?, 'manual', ?, ?, ?, ?, ?, 'missing', 0, '[]', ?, ?)
  `).run(
    id, batchId, String(sheet_name), String(seq || (maxRow.m || 0) + 1), String(doc_name).trim(),
    is_required ? 1 : 0, String(media_type), String(paper_req), String(electronic_req),
    String(policy_basis), String(template_source), String(notes), (maxRow.m || 0) + 1
  );

  res.status(201).json(db.prepare('SELECT * FROM archive_items WHERE id = ?').get(id));
});

// PUT /api/archive/items/:id
archiveRouter.put('/items/:id', (req, res) => {
  const existing = db.prepare('SELECT * FROM archive_items WHERE id = ?').get(req.params.id) as Record<string, unknown> | undefined;
  if (!existing) return res.status(404).json({ error: '归档条目不存在' });

  const {
    doc_name = existing.doc_name,
    is_required,
    media_type = existing.media_type,
    paper_req = existing.paper_req,
    electronic_req = existing.electronic_req,
    policy_basis = existing.policy_basis,
    template_source = existing.template_source,
    status = existing.status,
    is_manual_marked,
    linked_file_ids,
    notes = existing.notes,
  } = req.body as Record<string, unknown>;

  const nextRequired = is_required !== undefined ? (is_required ? 1 : 0) : Number(existing.is_required);
  const nextManual = is_manual_marked !== undefined ? (is_manual_marked ? 1 : 0) : Number(existing.is_manual_marked);
  const nextLinkedJson = linked_file_ids !== undefined
    ? JSON.stringify(Array.isArray(linked_file_ids) ? linked_file_ids : [])
    : String(existing.linked_file_ids_json || '[]');

  db.prepare(`
    UPDATE archive_items
    SET doc_name = ?, is_required = ?, media_type = ?, paper_req = ?, electronic_req = ?, policy_basis = ?, template_source = ?, status = ?, is_manual_marked = ?, linked_file_ids_json = ?, notes = ?
    WHERE id = ?
  `).run(
    String(doc_name).trim(), nextRequired, String(media_type), String(paper_req),
    String(electronic_req), String(policy_basis), String(template_source),
    String(status), nextManual, nextLinkedJson, String(notes), req.params.id
  );

  res.json(db.prepare('SELECT * FROM archive_items WHERE id = ?').get(req.params.id));
});

// POST /api/archive/items/:id/upload - Upload file directly to an archive item
archiveRouter.post('/items/:id/upload', uploadItemFile.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: '请选择要上传的归档文件' });
  const item = db.prepare('SELECT * FROM archive_items WHERE id = ?').get(req.params.id) as Record<string, unknown> | undefined;
  if (!item) return res.status(404).json({ error: '归档条目不存在' });

  const batchId = String(item.batch_id);
  const originalName = decodeName(req.file.originalname);
  const relPath = path.relative(UPLOADS_DIR, req.file.path);
  const fileId = uuid();
  const now = new Date().toISOString();

  db.prepare(`
    INSERT INTO uploaded_files (id, batch_id, module, category, original_name, stored_path, mime_type, file_size, parsed_json, notes, created_at)
    VALUES (?, ?, 'archive', ?, ?, ?, ?, ?, NULL, ?, ?)
  `).run(
    fileId, batchId, String(item.doc_name), originalName, relPath,
    req.file.mimetype || 'application/octet-stream', req.file.size, `归档条目：${item.doc_name}`, now
  );

  let linked: string[] = [];
  try {
    linked = JSON.parse(String(item.linked_file_ids_json || '[]'));
  } catch {
    linked = [];
  }
  linked.push(fileId);

  db.prepare(`
    UPDATE archive_items SET status = 'present', linked_file_ids_json = ? WHERE id = ?
  `).run(JSON.stringify(linked), req.params.id);

  res.status(201).json({ ok: true, fileId, originalName });
});

// DELETE /api/archive/items/:id
archiveRouter.delete('/items/:id', (req, res) => {
  db.prepare('DELETE FROM archive_items WHERE id = ?').run(req.params.id);
  res.json({ ok: true });
});

// GET /api/archive/batch/:batchId/export-xlsx - Export progress sheet
archiveRouter.get('/batch/:batchId/export-xlsx', async (req, res) => {
  const batchId = req.params.batchId;
  const batch = db.prepare('SELECT name, code FROM batches WHERE id = ?').get(batchId) as { name: string; code: string } | undefined;
  const items = db.prepare('SELECT * FROM archive_items WHERE batch_id = ? AND is_template = 0 ORDER BY sheet_name ASC, sort_order ASC').all(batchId) as Array<Record<string, unknown>>;

  const wb = new ExcelJS.Workbook();
  const sheetNames = Array.from(new Set(items.map((i) => String(i.sheet_name))));

  for (const sName of sheetNames) {
    const ws = wb.addWorksheet(sName);
    const headers = ['序号', '文件名称', '是否必须', '当前状态', '保存介质', '纸质要求', '电子要求', '流转备注', '制度依据', '模板来源'];
    const headerRow = ws.addRow(headers);
    headerRow.height = 22;
    ws.pageSetup.printTitlesRow = '1:1';
    ws.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: headers.length },
    };
    const allBorders = {
      top: { style: 'thin' as const, color: { argb: 'FF000000' } },
      bottom: { style: 'thin' as const, color: { argb: 'FF000000' } },
      left: { style: 'thin' as const, color: { argb: 'FF000000' } },
      right: { style: 'thin' as const, color: { argb: 'FF000000' } },
    };
    for (let c = 1; c <= headers.length; c++) {
      const cell = headerRow.getCell(c);
      cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF000000' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      cell.border = allBorders;
    }

    items.filter((i) => String(i.sheet_name) === sName).forEach((it) => {
      const statusText = it.status === 'present' ? '已有' : it.status === 'na' ? '不涉及(无)' : '缺失';
      const row = ws.addRow([
        it.seq,
        it.doc_name,
        it.is_required ? '必须' : '非必须',
        statusText,
        it.media_type,
        it.paper_req,
        it.electronic_req,
        it.notes,
        it.policy_basis,
        it.template_source,
      ]);
      row.height = 18;
      for (let c = 1; c <= headers.length; c++) {
        const cell = row.getCell(c);
        cell.font = { name: 'Calibri', size: 11, color: { argb: 'FF000000' } };
        cell.alignment = { vertical: 'middle', wrapText: true };
        cell.border = allBorders;
      }
      if (it.is_required) {
        row.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
      }
      if (it.status === 'missing' && it.is_required) {
        row.getCell(4).font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFA94442' } };
      } else if (it.status === 'present') {
        row.getCell(4).font = { name: 'Calibri', size: 11, color: { argb: 'FF2F6F5E' } };
      }
    });

    ws.columns = [
      { width: 8 }, { width: 38 }, { width: 10 }, { width: 12 }, { width: 14 },
      { width: 28 }, { width: 28 }, { width: 24 }, { width: 30 }, { width: 26 },
    ];
  }

  const filename = `${(batch?.code || 'archive')}_归档进度核对表.xlsx`;
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
  await wb.xlsx.write(res);
  res.end();
});

// GET /api/archive/batch/:batchId/export-zip - Package all linked physical files
archiveRouter.get('/batch/:batchId/export-zip', (req, res) => {
  const batchId = req.params.batchId;
  const batch = db.prepare('SELECT name, code FROM batches WHERE id = ?').get(batchId) as { name: string; code: string } | undefined;
  const items = db.prepare('SELECT * FROM archive_items WHERE batch_id = ? AND is_template = 0 ORDER BY sheet_name ASC, sort_order ASC').all(batchId) as Array<Record<string, unknown>>;
  const files = db.prepare('SELECT * FROM uploaded_files WHERE batch_id = ?').all(batchId) as Array<Record<string, unknown>>;
  const fileMap = new Map(files.map((f) => [String(f.id), f]));

  const zipName = `${(batch?.code || 'archive')}_电子归档文件包.zip`;
  res.setHeader('Content-Type', 'application/zip');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(zipName)}`);

  const archive = archiver('zip', { zlib: { level: 6 } });
  archive.pipe(res);

  let addedCount = 0;
  for (const it of items) {
    let linkedIds: string[] = [];
    try {
      linkedIds = JSON.parse(String(it.linked_file_ids_json || '[]'));
    } catch {
      linkedIds = [];
    }
    const safeDoc = String(it.doc_name).replace(/[/\\:*?"<>|]/g, '_').slice(0, 40);
    for (const fid of linkedIds) {
      const f = fileMap.get(fid);
      if (!f || !f.stored_path) continue;
      const absPath = path.join(UPLOADS_DIR, String(f.stored_path));
      if (fs.existsSync(absPath)) {
        archive.file(absPath, {
          name: `${String(it.sheet_name)}/${String(it.seq)}_${safeDoc}/${String(f.original_name)}`,
        });
        addedCount++;
      }
    }
  }

  if (addedCount === 0) {
    archive.append(
      `批次：${batch?.name || batchId}\n当前暂无已上传的物理电子归档文件，请先在工作台上传具体电子文件后再打包。\n`,
      { name: '说明.txt' }
    );
  }

  archive.finalize();
});
