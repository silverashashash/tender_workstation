import fs from 'node:fs';
import path from 'node:path';
import { DatabaseSync } from 'node:sqlite';

const defaultBaseDir = path.resolve(process.cwd());
const customDataDir = process.env.DATA_DIR?.trim();

export const DATA_DIR = customDataDir
  ? path.resolve(customDataDir)
  : path.join(defaultBaseDir, 'data');

export const UPLOADS_DIR = customDataDir
  ? path.join(path.resolve(customDataDir), 'uploads')
  : path.join(defaultBaseDir, 'uploads');

fs.mkdirSync(DATA_DIR, { recursive: true });
fs.mkdirSync(UPLOADS_DIR, { recursive: true });

export const DB_PATH = path.join(DATA_DIR, 'workstation.db');

export const db = new DatabaseSync(DB_PATH);

db.exec('PRAGMA journal_mode = WAL;');
db.exec('PRAGMA foreign_keys = ON;');

function safeAddColumn(table: string, colDef: string) {
  try {
    db.exec(`ALTER TABLE ${table} ADD COLUMN ${colDef};`);
  } catch {
    // Column already exists
  }
}

export function initSchema() {
  db.exec(`
    CREATE TABLE IF NOT EXISTS batches (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL,
      code TEXT NOT NULL DEFAULT '',
      project_unit TEXT NOT NULL DEFAULT '',
      procurement_type TEXT NOT NULL DEFAULT '非公开招标',
      unit_nature TEXT NOT NULL DEFAULT '产业',
      announcement_time TEXT NOT NULL DEFAULT '',
      doc_fetch_time TEXT NOT NULL DEFAULT '',
      bid_opening_time TEXT NOT NULL DEFAULT '',
      contact_phone TEXT NOT NULL DEFAULT '',
      contact_email TEXT NOT NULL DEFAULT '',
      status TEXT NOT NULL DEFAULT '筹备中',
      notes TEXT NOT NULL DEFAULT '',
      extra_fields_json TEXT NOT NULL DEFAULT '[]',
      created_at TEXT NOT NULL,
      updated_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS batch_projects (
      id TEXT PRIMARY KEY,
      batch_id TEXT NOT NULL,
      seq INTEGER NOT NULL DEFAULT 1,
      name TEXT NOT NULL,
      code TEXT NOT NULL DEFAULT '',
      budget TEXT NOT NULL DEFAULT '',
      procurement_method TEXT NOT NULL DEFAULT '竞争性谈判',
      department TEXT NOT NULL DEFAULT '',
      notes TEXT NOT NULL DEFAULT '',
      created_at TEXT NOT NULL,
      FOREIGN KEY (batch_id) REFERENCES batches(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS batch_suppliers (
      id TEXT PRIMARY KEY,
      batch_id TEXT NOT NULL,
      seq INTEGER NOT NULL DEFAULT 1,
      supplier_name TEXT NOT NULL,
      package_no TEXT NOT NULL DEFAULT '',
      contact_name TEXT NOT NULL DEFAULT '',
      mobile TEXT NOT NULL DEFAULT '',
      phone TEXT NOT NULL DEFAULT '',
      email TEXT NOT NULL DEFAULT '',
      address TEXT NOT NULL DEFAULT '',
      postal_code TEXT NOT NULL DEFAULT '',
      tax_no TEXT NOT NULL DEFAULT '',
      invoice_title TEXT NOT NULL DEFAULT '',
      bank_name TEXT NOT NULL DEFAULT '',
      bank_account TEXT NOT NULL DEFAULT '',
      created_at TEXT NOT NULL,
      FOREIGN KEY (batch_id) REFERENCES batches(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS eval_bids (
      id TEXT PRIMARY KEY,
      batch_id TEXT NOT NULL,
      lot_code TEXT NOT NULL DEFAULT '',
      lot_name TEXT NOT NULL DEFAULT '',
      package_no TEXT NOT NULL DEFAULT '',
      round_label TEXT NOT NULL DEFAULT '第1轮',
      supplier_name TEXT NOT NULL,
      discount_rate TEXT NOT NULL DEFAULT '',
      bid_price_wan REAL,
      bid_price_raw TEXT NOT NULL DEFAULT '',
      remark TEXT NOT NULL DEFAULT '',
      created_at TEXT NOT NULL,
      FOREIGN KEY (batch_id) REFERENCES batches(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS eval_line_quotes (
      id TEXT PRIMARY KEY,
      batch_id TEXT NOT NULL,
      lot_code TEXT NOT NULL DEFAULT '',
      package_no TEXT NOT NULL DEFAULT '',
      round_label TEXT NOT NULL DEFAULT '第1轮',
      supplier_name TEXT NOT NULL,
      req_app_no TEXT NOT NULL DEFAULT '',
      project_name TEXT NOT NULL DEFAULT '',
      tax_rate TEXT NOT NULL DEFAULT '',
      price_ex_tax_wan REAL,
      price_inc_tax_wan REAL,
      discount_rate TEXT NOT NULL DEFAULT '',
      created_at TEXT NOT NULL,
      FOREIGN KEY (batch_id) REFERENCES batches(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS eval_rejections (
      id TEXT PRIMARY KEY,
      batch_id TEXT NOT NULL,
      lot_code TEXT NOT NULL DEFAULT '',
      lot_name TEXT NOT NULL DEFAULT '',
      package_no TEXT NOT NULL DEFAULT '',
      supplier_name TEXT NOT NULL,
      reason_category TEXT NOT NULL DEFAULT '',
      rejection_tpl TEXT NOT NULL DEFAULT '',
      doc_requirement TEXT NOT NULL DEFAULT '',
      bid_response TEXT NOT NULL DEFAULT '',
      review_type TEXT NOT NULL DEFAULT '',
      notes TEXT NOT NULL DEFAULT '',
      created_at TEXT NOT NULL,
      FOREIGN KEY (batch_id) REFERENCES batches(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS eval_rankings (
      id TEXT PRIMARY KEY,
      batch_id TEXT NOT NULL,
      lot_code TEXT NOT NULL DEFAULT '',
      lot_name TEXT NOT NULL DEFAULT '',
      package_no TEXT NOT NULL DEFAULT '',
      project_name TEXT NOT NULL DEFAULT '',
      eval_date TEXT NOT NULL DEFAULT '',
      supplier_name TEXT NOT NULL DEFAULT '',
      tech_score REAL,
      biz_score REAL,
      price_score REAL,
      total_score REAL,
      rank_no INTEGER,
      base_price TEXT NOT NULL DEFAULT '',
      bid_price TEXT NOT NULL DEFAULT '',
      eval_price TEXT NOT NULL DEFAULT '',
      is_failed_lot INTEGER NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL,
      FOREIGN KEY (batch_id) REFERENCES batches(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS uploaded_files (
      id TEXT PRIMARY KEY,
      batch_id TEXT NOT NULL,
      module TEXT NOT NULL,
      slot_key TEXT NOT NULL DEFAULT '',
      category TEXT NOT NULL DEFAULT '其他',
      original_name TEXT NOT NULL,
      stored_path TEXT NOT NULL,
      mime_type TEXT NOT NULL DEFAULT 'application/octet-stream',
      file_size INTEGER NOT NULL DEFAULT 0,
      parsed_json TEXT,
      notes TEXT NOT NULL DEFAULT '',
      created_at TEXT NOT NULL,
      FOREIGN KEY (batch_id) REFERENCES batches(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS sop_items (
      id TEXT PRIMARY KEY,
      batch_id TEXT,
      is_template INTEGER NOT NULL DEFAULT 0,
      item_type TEXT NOT NULL DEFAULT 'step',
      stage_group TEXT NOT NULL DEFAULT '开标前准备',
      priority TEXT NOT NULL DEFAULT 'normal',
      title TEXT NOT NULL,
      description TEXT NOT NULL DEFAULT '',
      is_completed INTEGER NOT NULL DEFAULT 0,
      completed_at TEXT,
      notes TEXT NOT NULL DEFAULT '',
      sort_order INTEGER NOT NULL DEFAULT 0,
      FOREIGN KEY (batch_id) REFERENCES batches(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS sop_guides (
      id TEXT PRIMARY KEY,
      stage_group TEXT NOT NULL DEFAULT '内网开标',
      title TEXT NOT NULL,
      source_type TEXT NOT NULL DEFAULT 'markdown',
      content_md TEXT NOT NULL DEFAULT '',
      images_json TEXT NOT NULL DEFAULT '[]',
      file_original_name TEXT NOT NULL DEFAULT '',
      file_stored_path TEXT NOT NULL DEFAULT '',
      file_mime_type TEXT NOT NULL DEFAULT '',
      file_size INTEGER NOT NULL DEFAULT 0,
      sort_order INTEGER NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL,
      updated_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS award_outputs (
      id TEXT PRIMARY KEY,
      batch_id TEXT NOT NULL,
      template_key TEXT NOT NULL,
      title TEXT NOT NULL,
      file_type TEXT NOT NULL DEFAULT 'xlsx',
      filename TEXT NOT NULL,
      stored_path TEXT NOT NULL,
      snapshot_json TEXT NOT NULL DEFAULT '{}',
      version INTEGER NOT NULL DEFAULT 1,
      created_at TEXT NOT NULL,
      FOREIGN KEY (batch_id) REFERENCES batches(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS archive_items (
      id TEXT PRIMARY KEY,
      batch_id TEXT,
      is_template INTEGER NOT NULL DEFAULT 0,
      sheet_name TEXT NOT NULL DEFAULT '归档文件',
      seq TEXT NOT NULL DEFAULT '',
      doc_name TEXT NOT NULL,
      is_required INTEGER NOT NULL DEFAULT 0,
      required_source TEXT NOT NULL DEFAULT 'manual',
      media_type TEXT NOT NULL DEFAULT '纸质+电子',
      paper_req TEXT NOT NULL DEFAULT '',
      electronic_req TEXT NOT NULL DEFAULT '',
      policy_basis TEXT NOT NULL DEFAULT '',
      template_source TEXT NOT NULL DEFAULT '',
      status TEXT NOT NULL DEFAULT 'missing',
      is_manual_marked INTEGER NOT NULL DEFAULT 0,
      linked_file_ids_json TEXT NOT NULL DEFAULT '[]',
      notes TEXT NOT NULL DEFAULT '',
      sort_order INTEGER NOT NULL DEFAULT 0,
      FOREIGN KEY (batch_id) REFERENCES batches(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS client_agreements (
      id TEXT PRIMARY KEY,
      group_name TEXT NOT NULL DEFAULT '主业单位',
      client_name TEXT NOT NULL,
      contact_name TEXT NOT NULL DEFAULT '',
      contact_phone TEXT NOT NULL DEFAULT '',
      contract_expiry TEXT NOT NULL DEFAULT '',
      discount_rate REAL NOT NULL DEFAULT 1,
      discount_label TEXT NOT NULL DEFAULT '不打折 (100%)',
      agreement_status TEXT NOT NULL DEFAULT '',
      notes TEXT NOT NULL DEFAULT '',
      sort_order INTEGER NOT NULL DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS material_catalog (
      material_code TEXT PRIMARY KEY,
      material_desc TEXT NOT NULL DEFAULT '',
      lot_name TEXT NOT NULL DEFAULT ''
    );

    CREATE TABLE IF NOT EXISTS eval_ip_records (
      id TEXT PRIMARY KEY,
      batch_id TEXT NOT NULL,
      lot_code TEXT NOT NULL DEFAULT '',
      lot_name TEXT NOT NULL DEFAULT '',
      package_no TEXT NOT NULL DEFAULT '',
      round_label TEXT NOT NULL DEFAULT '1',
      supplier_name TEXT NOT NULL,
      download_ip TEXT NOT NULL DEFAULT '',
      tech_ip TEXT NOT NULL DEFAULT '',
      biz_ip TEXT NOT NULL DEFAULT '',
      price_ip TEXT NOT NULL DEFAULT '',
      qual_ip TEXT NOT NULL DEFAULT '',
      tech_mac TEXT NOT NULL DEFAULT '',
      biz_mac TEXT NOT NULL DEFAULT '',
      price_mac TEXT NOT NULL DEFAULT '',
      qual_mac TEXT NOT NULL DEFAULT '',
      tech_cpu TEXT NOT NULL DEFAULT '',
      biz_cpu TEXT NOT NULL DEFAULT '',
      price_cpu TEXT NOT NULL DEFAULT '',
      qual_cpu TEXT NOT NULL DEFAULT '',
      tech_disk TEXT NOT NULL DEFAULT '',
      biz_disk TEXT NOT NULL DEFAULT '',
      price_disk TEXT NOT NULL DEFAULT '',
      qual_disk TEXT NOT NULL DEFAULT '',
      source_file TEXT NOT NULL DEFAULT '',
      sheet_name TEXT NOT NULL DEFAULT '',
      source_row INTEGER,
      created_at TEXT NOT NULL,
      FOREIGN KEY (batch_id) REFERENCES batches(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS eval_score_checks (
      batch_id TEXT NOT NULL,
      check_type TEXT NOT NULL,
      files_json TEXT NOT NULL DEFAULT '[]',
      config_json TEXT NOT NULL DEFAULT '{}',
      methods_json TEXT NOT NULL DEFAULT '[]',
      result_json TEXT,
      updated_at TEXT NOT NULL,
      PRIMARY KEY (batch_id, check_type),
      FOREIGN KEY (batch_id) REFERENCES batches(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS batch_change_notices (
      id TEXT PRIMARY KEY,
      batch_id TEXT NOT NULL,
      title TEXT NOT NULL DEFAULT '',
      notice_date TEXT NOT NULL DEFAULT '',
      change_types_json TEXT NOT NULL DEFAULT '[]',
      summary TEXT NOT NULL DEFAULT '',
      project_ids_json TEXT NOT NULL DEFAULT '[]',
      file_original_name TEXT NOT NULL DEFAULT '',
      file_stored_path TEXT NOT NULL DEFAULT '',
      file_mime_type TEXT NOT NULL DEFAULT '',
      created_at TEXT NOT NULL,
      FOREIGN KEY (batch_id) REFERENCES batches(id) ON DELETE CASCADE
    );
  `);

  // Safe additive migrations for existing SQLite databases
  safeAddColumn('batches', "procurement_type TEXT NOT NULL DEFAULT '非公开招标'");
  safeAddColumn('batches', "unit_nature TEXT NOT NULL DEFAULT '产业'");
  safeAddColumn('batches', "doc_fetch_time TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batches', "contact_phone TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batches', "contact_email TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batches', "award_client_id TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batches', "award_category TEXT NOT NULL DEFAULT '服务'");
  safeAddColumn('batches', 'award_discount_rate REAL');
  safeAddColumn('batches', "award_batch_manager TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batches', "award_bid_location TEXT NOT NULL DEFAULT '华亿'");
  safeAddColumn('batches', "award_publicity_date TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batches', "award_announcement_date TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batches', "erp_conflicts_json TEXT NOT NULL DEFAULT '[]'");
  safeAddColumn('batches', "eval_tax_category TEXT NOT NULL DEFAULT '服务'");
  safeAddColumn('batches', "eval_tax_proof_rates TEXT NOT NULL DEFAULT '0,1'");
  safeAddColumn('batches', 'deleted_at TEXT');

  safeAddColumn('uploaded_files', "slot_key TEXT NOT NULL DEFAULT ''");

  safeAddColumn('batch_projects', "project_code TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "req_app_no TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "tech_spec_code TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "lot_code TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "lot_name TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "package_no TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "project_unit TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "project_overview TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "qualification_req TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "performance_req TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "staff_req TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "quote_method TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "subcontract_rule TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', 'budget_yuan REAL');
  safeAddColumn('batch_projects', 'max_price_yuan REAL');
  safeAddColumn('batch_projects', "max_price_raw TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "declarant_name TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "declarant_phone TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "weight_ratio TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "weight_warning TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "weight_suggested TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "price_formula TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "price_params TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "biz_eval_tpl TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "tech_eval_tpl TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "contract_version TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "funding_source TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "duration TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "erp_sync_status TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "material_code TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_projects', "material_desc TEXT NOT NULL DEFAULT ''");

  safeAddColumn('batch_suppliers', "package_no TEXT NOT NULL DEFAULT ''");
  safeAddColumn('batch_suppliers', "fax TEXT NOT NULL DEFAULT ''");

  safeAddColumn('eval_ip_records', "source_file TEXT NOT NULL DEFAULT ''");
  safeAddColumn('eval_ip_records', "sheet_name TEXT NOT NULL DEFAULT ''");
  safeAddColumn('eval_ip_records', 'source_row INTEGER');

  // Clean up deprecated default "开标前准备" SOP template items per user preference
  db.prepare(`
    DELETE FROM sop_items
    WHERE is_template = 1
      AND stage_group = '开标前准备'
      AND (title LIKE '核验 CA 证书%' OR title LIKE '核对最高限价审批原件%')
  `).run();

  // Clean up any accidental repeated-sheet header rows in eval_bids / eval_rejections
  db.prepare(`
    DELETE FROM eval_bids
    WHERE supplier_name IN ('投标人名称', '投标人', '响应供应商', '供应商名称')
       OR lot_code = '分标编号'
  `).run();
  db.prepare(`
    DELETE FROM eval_rejections
    WHERE supplier_name IN ('投标人名称', '投标人', '响应供应商', '供应商名称')
  `).run();
}
