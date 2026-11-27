import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import XLSX from 'xlsx';
import type { ScoreUploadedFile } from './expertScoreAudit.js';

export type ChecklistStatus = 'OK' | '缺失' | '异常' | '解压失败' | '注意';

export interface ChecklistItemResult {
  batch: string;
  specId: string;
  status: ChecklistStatus;
  fileName: string;
  fileSize: number | null;
  relativePath: string;
  note: string;
}

export interface ChecklistComplianceRunResult {
  keywords: string[];
  items: ChecklistItemResult[];
  counts: {
    total: number;
    ok: number;
    missing: number;
    abnormal: number;
    unpackFailed: number;
    notice: number;
  };
  reportBuffer: Buffer;
  reportFileName: string;
  logs: string[];
}

const DEFAULT_KEYWORDS = ['工作量清单', '工程量清单'];
const ARCHIVE_EXTS = ['.zip', '.7z', '.rar'];

// ==================== Pure-Node ZIP Extractor with GBK/UTF-8 Filename Repair ====================

const gbkDecoder = new TextDecoder('gb18030');
const utf8Decoder = new TextDecoder('utf-8', { fatal: true });

function decodeZipFilename(rawBytes: Buffer, isUtf8Flag: boolean): string {
  if (isUtf8Flag) {
    return rawBytes.toString('utf8');
  }
  try {
    return utf8Decoder.decode(rawBytes);
  } catch {
    try {
      return gbkDecoder.decode(rawBytes);
    } catch {
      return rawBytes.toString('latin1');
    }
  }
}

function safeJoin(base: string, relName: string): string {
  const normalized = relName.replace(/\\/g, '/');
  const parts = normalized.split('/').filter((p) => p && p !== '.' && p !== '..');
  return path.join(base, ...parts);
}

/**
 * Extract a ZIP buffer into `dest`, properly decoding GBK/GB18030 Chinese filenames
 * created on Windows/WPS/ECP systems where UTF-8 flag (bit 11) is unset.
 */
import zlib from 'node:zlib';

function extractZipBuffer(zipBuf: Buffer, dest: string): void {
  fs.mkdirSync(dest, { recursive: true });

  // Locate End of Central Directory (EOCD) signature 0x06054b50 scanning backwards
  let eocdOffset = -1;
  const minOffset = Math.max(0, zipBuf.length - 65557);
  for (let i = zipBuf.length - 22; i >= minOffset; i--) {
    if (
      zipBuf[i] === 0x50 &&
      zipBuf[i + 1] === 0x4b &&
      zipBuf[i + 2] === 0x05 &&
      zipBuf[i + 3] === 0x06
    ) {
      eocdOffset = i;
      break;
    }
  }
  if (eocdOffset < 0) {
    throw new Error('无效的 ZIP 压缩包（未找到中央目录结束标记）');
  }

  const cdEntries = zipBuf.readUInt16LE(eocdOffset + 10);
  const cdOffset = zipBuf.readUInt32LE(eocdOffset + 16);
  let pos = cdOffset;

  for (let i = 0; i < cdEntries; i++) {
    if (pos + 46 > zipBuf.length || zipBuf.readUInt32LE(pos) !== 0x02014b50) {
      break;
    }
    const flags = zipBuf.readUInt16LE(pos + 8);
    const method = zipBuf.readUInt16LE(pos + 10);
    const compSize = zipBuf.readUInt32LE(pos + 20);
    const nameLen = zipBuf.readUInt16LE(pos + 28);
    const extraLen = zipBuf.readUInt16LE(pos + 30);
    const commentLen = zipBuf.readUInt16LE(pos + 32);
    const localHeaderOffset = zipBuf.readUInt32LE(pos + 42);

    const rawNameBytes = zipBuf.subarray(pos + 46, pos + 46 + nameLen);
    const isUtf8 = (flags & 0x0800) !== 0;
    const entryName = decodeZipFilename(rawNameBytes, isUtf8);
    pos += 46 + nameLen + extraLen + commentLen;

    // Skip macOS metadata folders
    if (entryName.startsWith('__MACOSX/') || path.basename(entryName).startsWith('._')) {
      continue;
    }
    if (entryName.endsWith('/')) {
      fs.mkdirSync(safeJoin(dest, entryName), { recursive: true });
      continue;
    }

    if (
      localHeaderOffset + 30 > zipBuf.length ||
      zipBuf.readUInt32LE(localHeaderOffset) !== 0x04034b50
    ) {
      continue;
    }
    const lhNameLen = zipBuf.readUInt16LE(localHeaderOffset + 26);
    const lhExtraLen = zipBuf.readUInt16LE(localHeaderOffset + 28);
    const dataStart = localHeaderOffset + 30 + lhNameLen + lhExtraLen;
    const compData = zipBuf.subarray(dataStart, dataStart + compSize);

    const targetPath = safeJoin(dest, entryName);
    fs.mkdirSync(path.dirname(targetPath), { recursive: true });

    try {
      let plain: Buffer;
      if (method === 0) {
        plain = compData;
      } else if (method === 8) {
        plain = zlib.inflateRawSync(compData);
      } else {
        throw new Error(`不支持的压缩算法(${method})`);
      }
      fs.writeFileSync(targetPath, plain);
    } catch (err) {
      fs.writeFileSync(`${targetPath}.__READ_ERROR__`, String((err as Error).message));
    }
  }
}

function findCliTool(candidates: string[]): string | null {
  for (const c of candidates) {
    if (path.isAbsolute(c)) {
      if (fs.existsSync(c)) return c;
    } else {
      try {
        execFileSync(process.platform === 'win32' ? 'where' : 'which', [c], { stdio: 'ignore' });
        return c;
      } catch {
        // continue
      }
    }
  }
  return null;
}

function extractArchiveFile(archivePath: string, dest: string, notes: string[]): boolean {
  const ext = path.extname(archivePath).toLowerCase();
  fs.mkdirSync(dest, { recursive: true });
  try {
    if (ext === '.zip') {
      extractZipBuffer(fs.readFileSync(archivePath), dest);
      return true;
    }
    if (ext === '.rar' || ext === '.7z') {
      const unar = findCliTool(['/opt/homebrew/bin/unar', '/usr/local/bin/unar', 'unar']);
      if (unar) {
        execFileSync(unar, ['-q', '-D', '-o', dest, archivePath], {
          stdio: 'pipe',
          timeout: 60_000,
        });
        return true;
      }
      const sevenZip = findCliTool([
        '/opt/homebrew/bin/7z',
        '/usr/local/bin/7z',
        '7z',
        '7za',
        'C:\\Program Files\\7-Zip\\7z.exe',
      ]);
      if (sevenZip) {
        execFileSync(sevenZip, ['x', '-y', `-o${dest}`, archivePath], {
          stdio: 'pipe',
          timeout: 60_000,
        });
        return true;
      }
      notes.push(`解压失败(${path.basename(archivePath)}): 服务器未安装 unar 或 7z，无法解压 ${ext}`);
      return false;
    }
  } catch (err) {
    notes.push(`解压失败(${path.basename(archivePath)}): ${(err as Error).message.slice(0, 160)}`);
    return false;
  }
  return false;
}

function walkFiles(dir: string): string[] {
  const out: string[] = [];
  if (!fs.existsSync(dir)) return out;
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      if (entry.name === '__MACOSX') continue;
      out.push(...walkFiles(full));
    } else if (entry.isFile()) {
      out.push(full);
    }
  }
  return out;
}

function recursiveUnpack(root: string, notes: string[], maxPasses = 10): void {
  for (let pass = 0; pass < maxPasses; pass++) {
    const allFiles = walkFiles(root);
    const archives = allFiles.filter((fp) => {
      const lower = fp.toLowerCase();
      return ARCHIVE_EXTS.some((ext) => lower.endsWith(ext));
    });
    if (archives.length === 0) break;

    for (const ap of archives) {
      const stem = path.basename(ap, path.extname(ap));
      const subDir = path.join(path.dirname(ap), stem);
      const ok = extractArchiveFile(ap, subDir, notes);
      if (ok) {
        try {
          fs.unlinkSync(ap);
        } catch {
          // ignore
        }
      } else {
        try {
          fs.renameSync(ap, `${ap}.__UNPACK_FAILED__`);
        } catch {
          // ignore
        }
      }
    }
  }
}

// ==================== Excel Content Validator ====================

function validateChecklistFile(filePath: string): { status: ChecklistStatus; note: string; size: number } {
  const size = fs.statSync(filePath).size;
  const ext = path.extname(filePath).toLowerCase();
  if (size < 1024) {
    return { status: '异常', note: `文件过小(${size}字节), 疑似空文件`, size };
  }
  if (ext === '.xlsx' || ext === '.xlsm' || ext === '.xls') {
    let wb: XLSX.WorkBook;
    try {
      wb = XLSX.read(fs.readFileSync(filePath), { type: 'buffer', cellDates: false });
    } catch (err) {
      return { status: '异常', note: `无法打开(坏文件): ${(err as Error).name}`, size };
    }
    let hasContent = false;
    for (const sName of wb.SheetNames) {
      const ws = wb.Sheets[sName];
      if (!ws?.['!ref']) continue;
      const range = XLSX.utils.decode_range(ws['!ref']);
      const maxR = Math.min(range.e.r, 50);
      const maxC = Math.min(range.e.c, 30);
      for (let r = range.s.r; r <= maxR; r++) {
        for (let c = range.s.c; c <= maxC; c++) {
          const cell = ws[XLSX.utils.encode_cell({ r, c })] as XLSX.CellObject | undefined;
          if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '') {
            hasContent = true;
            break;
          }
        }
        if (hasContent) break;
      }
      if (hasContent) break;
    }
    if (!hasContent) {
      return { status: '异常', note: '可打开但内容为空', size };
    }
    return { status: 'OK', note: '正常', size };
  }
  return { status: '注意', note: `非Excel格式(${ext || '无扩展名'})`, size };
}

/**
 * Find the directory inside an unpacked batch zip that holds project `.zip` files (BFS depth <= 4).
 */
function findProjectZipDir(rootDir: string): string | null {
  const queue: Array<{ dir: string; depth: number }> = [{ dir: rootDir, depth: 0 }];
  while (queue.length > 0) {
    const { dir, depth } = queue.shift()!;
    let entries: fs.Dirent[];
    try {
      entries = fs.readdirSync(dir, { withFileTypes: true });
    } catch {
      continue;
    }
    if (
      entries.some(
        (e) =>
          e.isFile() &&
          e.name.toLowerCase().endsWith('.zip') &&
          !e.name.startsWith('.') &&
          !e.name.startsWith('._')
      )
    ) {
      return dir;
    }
    if (depth < 4) {
      for (const e of entries) {
        if (e.isDirectory() && e.name !== '__MACOSX' && !e.name.startsWith('.')) {
          queue.push({ dir: path.join(dir, e.name), depth: depth + 1 });
        }
      }
    }
  }
  return null;
}

// ==================== Report Excel Builder ====================

async function buildComplianceReportBuffer(items: ChecklistItemResult[]): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = '国网招标代理 SOP 工作台';

  const ws = wb.addWorksheet('清单合规性检验报告', {
    views: [{ state: 'frozen', ySplit: 1 }],
  });
  const headers = [
    '批次编号',
    '技术规范书ID',
    '检验结果',
    '清单文件名',
    '清单文件大小(字节)',
    '清单文件相对路径',
    '说明',
  ];
  const hRow = ws.addRow(headers);
  hRow.height = 24;
  const headerFill: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
  hRow.eachCell((c) => {
    c.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = headerFill;
    c.alignment = { horizontal: 'center', vertical: 'middle' };
  });

  const okFill: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } };
  const badFill: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
  const warnFill: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB9C' } };

  for (const it of items) {
    const row = ws.addRow([
      it.batch,
      it.specId,
      it.status,
      it.fileName,
      it.fileSize ?? '',
      it.relativePath,
      it.note,
    ]);
    const fill =
      it.status === 'OK'
        ? okFill
        : it.status === '缺失' || it.status === '异常' || it.status === '解压失败'
          ? badFill
          : warnFill;
    row.eachCell((c) => {
      c.fill = fill;
      c.font = { name: 'Calibri', size: 11 };
      c.alignment = { vertical: 'middle', wrapText: true };
    });
  }

  [42, 22, 12, 55, 18, 70, 30].forEach((w, idx) => {
    ws.getColumn(idx + 1).width = w;
  });

  // Sheet 2: 汇总统计
  const ws2 = wb.addWorksheet('汇总统计');
  const shRow = ws2.addRow(['检验结果', '项目数']);
  shRow.eachCell((c) => {
    c.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = headerFill;
  });
  const countMap = new Map<ChecklistStatus, number>();
  for (const it of items) {
    countMap.set(it.status, (countMap.get(it.status) || 0) + 1);
  }
  for (const k of ['OK', '缺失', '异常', '解压失败', '注意'] as const) {
    const v = countMap.get(k) || 0;
    if (v > 0) ws2.addRow([k, v]);
  }
  ws2.addRow(['合计', items.length]);
  ws2.getColumn(1).width = 14;
  ws2.getColumn(2).width = 12;

  return Buffer.from(await wb.xlsx.writeBuffer());
}

// ==================== Main Entry ====================

export async function runChecklistCompliance(params: {
  files: ScoreUploadedFile[];
  keywords?: string[];
}): Promise<ChecklistComplianceRunResult> {
  const keywords = (params.keywords || [])
    .map((k) => k.trim())
    .filter(Boolean);
  const activeKeywords = keywords.length > 0 ? keywords : DEFAULT_KEYWORDS;

  const zipFiles = params.files.filter(
    (f) => f.originalName.toLowerCase().endsWith('.zip') && !path.basename(f.originalName).startsWith('._')
  );
  if (zipFiles.length === 0) {
    throw new Error('请至少上传一个 .zip 压缩包（支持批次总压缩包，或直接多选项目技术规范书 .zip）');
  }

  const tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'checklist_compliance_'));
  const items: ChecklistItemResult[] = [];
  const logs: string[] = [`关键词：${activeKeywords.join('、')}`, `上传压缩包：${zipFiles.length} 个`];

  try {
    const inspectProjectZip = (batchName: string, specId: string, zipBuf: Buffer, workDir: string) => {
      const notes: string[] = [];
      try {
        extractZipBuffer(zipBuf, workDir);
        recursiveUnpack(workDir, notes);
      } catch (err) {
        const msg = `${(err as Error).name}: ${(err as Error).message}`;
        items.push({
          batch: batchName,
          specId,
          status: '解压失败',
          fileName: '',
          fileSize: null,
          relativePath: '',
          note: msg,
        });
        logs.push(`  [${batchName}] ${specId}: 解压失败 (${msg})`);
        return;
      }

      const allFiles = walkFiles(workDir);
      const hits = allFiles.filter((fp) => {
        const base = path.basename(fp);
        if (base.startsWith('.') || base.startsWith('~$')) return false;
        return activeKeywords.some((kw) => base.includes(kw));
      });
      const noteStr = notes.join('; ');

      if (hits.length === 0) {
        const msg = noteStr ? `未找到清单文件 | ${noteStr}` : '未找到清单文件';
        items.push({
          batch: batchName,
          specId,
          status: '缺失',
          fileName: '',
          fileSize: null,
          relativePath: '',
          note: msg,
        });
        logs.push(`  [${batchName}] ${specId}: 缺失 (${msg})`);
        return;
      }

      for (const h of hits) {
        const v = validateChecklistFile(h);
        const msg = noteStr ? `${v.note} | ${noteStr}` : v.note;
        const rel = path.relative(workDir, h);
        items.push({
          batch: batchName,
          specId,
          status: v.status,
          fileName: path.basename(h),
          fileSize: v.size,
          relativePath: `${batchName}/${specId}/${rel}`,
          note: msg,
        });
        logs.push(`  [${batchName}] ${specId}: [${v.status}] ${path.basename(h)} (${v.size}字节) - ${msg}`);
      }
    };

    for (let i = 0; i < zipFiles.length; i++) {
      const zf = zipFiles[i];
      const stem = path.basename(zf.originalName, path.extname(zf.originalName)).replace(/_+$/, '');
      const probeDir = path.join(tmpRoot, `in_${i}`);

      try {
        extractZipBuffer(zf.buffer, probeDir);
      } catch (err) {
        const msg = `${(err as Error).name}: ${(err as Error).message}`;
        items.push({
          batch: stem,
          specId: stem,
          status: '解压失败',
          fileName: '',
          fileSize: null,
          relativePath: '',
          note: msg,
        });
        logs.push(`  [${stem}] 解压失败: ${msg}`);
        continue;
      }

      // Auto-detect: if this zip contains inner project .zip files (batch archive like 122617.zip),
      // inspect each inner .zip as an individual 技术规范书ID; otherwise treat this zip itself as one project!
      const innerZipDir = findProjectZipDir(probeDir);
      if (innerZipDir) {
        const innerZips = fs
          .readdirSync(innerZipDir)
          .filter((fn) => fn.toLowerCase().endsWith('.zip') && !fn.startsWith('.') && !fn.startsWith('._'))
          .sort();
        logs.push(`\n批次 [${stem}] 检测到 ${innerZips.length} 个项目压缩包`);
        for (let j = 0; j < innerZips.length; j++) {
          const fn = innerZips[j];
          const specId = path.basename(fn, path.extname(fn)).replace(/_+$/, '');
          const projDir = path.join(tmpRoot, `proj_${i}_${j}`);
          inspectProjectZip(stem, specId, fs.readFileSync(path.join(innerZipDir, fn)), projDir);
        }
      } else {
        logs.push(`\n项目压缩包 [${stem}]`);
        const projDir = path.join(tmpRoot, `single_${i}`);
        inspectProjectZip('直接上传', stem, zf.buffer, projDir);
      }
    }

    const reportBuffer = await buildComplianceReportBuffer(items);
    const counts = {
      total: items.length,
      ok: items.filter((x) => x.status === 'OK').length,
      missing: items.filter((x) => x.status === '缺失').length,
      abnormal: items.filter((x) => x.status === '异常').length,
      unpackFailed: items.filter((x) => x.status === '解压失败').length,
      notice: items.filter((x) => x.status === '注意').length,
    };

    logs.push(
      `\n✅ 检验完成：共 ${counts.total} 项（OK ${counts.ok} · 缺失 ${counts.missing} · 异常 ${counts.abnormal} · 解压失败 ${counts.unpackFailed} · 注意 ${counts.notice}）`
    );

    return {
      keywords: activeKeywords,
      items,
      counts,
      reportBuffer,
      reportFileName: '清单合规性检验报告.xlsx',
      logs,
    };
  } finally {
    fs.rmSync(tmpRoot, { recursive: true, force: true });
  }
}
