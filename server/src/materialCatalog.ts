import fs from 'node:fs';
import path from 'node:path';
import * as XLSXNamespace from 'xlsx';
import { cleanText, stripAllSpaces } from './biddingParser.js';
import { db } from './db.js';

const XLSX = (XLSXNamespace as unknown as { default?: typeof XLSXNamespace }).default ?? XLSXNamespace;

export interface MaterialCatalogEntry {
  material_code: string;
  material_desc: string;
  lot_name: string;
}

export function parseMaterialCatalogExcel(filePath: string): MaterialCatalogEntry[] {
  const buf = fs.readFileSync(filePath);
  const wb = XLSX.read(buf, { type: 'buffer', raw: true });
  const targetSheetName =
    wb.SheetNames.find((n) => n.includes('物料编码')) || wb.SheetNames[0];
  const ws = wb.Sheets[targetSheetName];
  const aoa = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: '', raw: true });

  let hIdx = 0;
  for (let i = 0; i < Math.min(aoa.length, 6); i++) {
    const row = (aoa[i] || []).map((c) => cleanText(c));
    if (row.some((c) => /物料编码/.test(c)) && row.some((c) => /物料描述/.test(c))) {
      hIdx = i;
      break;
    }
  }

  const headers = (aoa[hIdx] || []).map((c) => cleanText(c));
  const idxCode = headers.findIndex((h) => /^物料编码$/.test(h) || /物料编码/.test(h));
  const idxDesc = headers.findIndex((h) => /^物料描述$/.test(h) || /物料描述/.test(h));
  const idxLot = headers.findIndex((h) => /^分标名称$/.test(h) || /分标名称/.test(h));

  if (idxCode < 0 || idxDesc < 0) {
    throw new Error('未在表格中找到「物料编码」和「物料描述」列');
  }

  const dedup = new Map<string, MaterialCatalogEntry>();
  for (let i = hIdx + 1; i < aoa.length; i++) {
    const r = aoa[i] || [];
    const code = stripAllSpaces(r[idxCode]);
    const desc = cleanText(r[idxDesc]);
    if (!code || !desc || code === '物料编码' || desc === '物料描述') continue;
    if (!dedup.has(code)) {
      dedup.set(code, {
        material_code: code,
        material_desc: desc,
        lot_name: idxLot >= 0 ? cleanText(r[idxLot]) : '',
      });
    }
  }

  return Array.from(dedup.values());
}

export function importMaterialCatalogToDb(filePath: string): number {
  const entries = parseMaterialCatalogExcel(filePath);
  if (entries.length === 0) {
    throw new Error('未解析到有效的物料编码记录');
  }
  db.exec('BEGIN');
  try {
    db.prepare('DELETE FROM material_catalog').run();
    const ins = db.prepare(
      'INSERT OR REPLACE INTO material_catalog (material_code, material_desc, lot_name) VALUES (?, ?, ?)'
    );
    for (const e of entries) {
      ins.run(e.material_code, e.material_desc, e.lot_name);
    }
    db.exec('COMMIT');
  } catch (err) {
    db.exec('ROLLBACK');
    throw err;
  }
  return entries.length;
}

export function seedMaterialCatalogIfEmpty(): void {
  const row = db.prepare('SELECT COUNT(*) as c FROM material_catalog').get() as { c: number };
  if (row && row.c > 0) return;

  const candidates = [
    path.resolve(process.cwd(), '../sample_docs/3.award/其他文件/服务物料编码--分标名称（授权采购）.xlsx'),
    path.resolve(process.cwd(), 'sample_docs/3.award/其他文件/服务物料编码--分标名称（授权采购）.xlsx'),
  ];
  for (const fp of candidates) {
    if (fs.existsSync(fp)) {
      try {
        const count = importMaterialCatalogToDb(fp);
        console.log(`[materialCatalog] Seeded ${count} material codes from ${path.basename(fp)}`);
      } catch (err) {
        console.warn(`[materialCatalog] Failed to seed from ${fp}:`, (err as Error).message);
      }
      return;
    }
  }
}

export function lookupMaterialDescMap(codes: string[]): Map<string, string> {
  const unique = Array.from(new Set(codes.map((c) => stripAllSpaces(c)).filter(Boolean)));
  const result = new Map<string, string>();
  if (unique.length === 0) return result;

  // Query in chunks of 400 to stay well within SQLite parameter limit
  for (let i = 0; i < unique.length; i += 400) {
    const chunk = unique.slice(i, i + 400);
    const placeholders = chunk.map(() => '?').join(',');
    const rows = db
      .prepare(`SELECT material_code, material_desc FROM material_catalog WHERE material_code IN (${placeholders})`)
      .all(...chunk) as Array<{ material_code: string; material_desc: string }>;
    for (const r of rows) {
      if (r.material_code && r.material_desc) {
        result.set(r.material_code, r.material_desc);
      }
    }
  }
  return result;
}
