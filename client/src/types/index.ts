export interface ExtraField {
  key: string;
  value: string;
}

export interface BatchStats {
  projectCount: number;
  totalBudgetWan?: number;
  biddingFileCount: number;
  evalFileCount: number;
  awardOutputCount: number;
  changeNoticeCount?: number;
  sopTotal: number;
  sopDone: number;
  archiveTotal: number;
  archivePresent: number;
  archiveNa: number;
  archiveReqTotal: number;
  archiveReqPresent: number;
  archiveReqMissing: number;
}

export interface Batch {
  id: string;
  name: string;
  code: string;
  project_unit: string;
  procurement_type?: '公开招标' | '非公开招标';
  unit_nature?: '主业' | '产业';
  announcement_time: string;
  doc_fetch_time?: string;
  bid_opening_time: string;
  contact_phone?: string;
  contact_email?: string;
  status: string;
  notes: string;
  extra_fields: ExtraField[];
  deleted_at?: string | null;
  created_at: string;
  updated_at: string;
  stats: BatchStats;
  projects?: BatchProject[];
}

export interface BatchProject {
  id: string;
  batch_id: string;
  seq: number;
  name: string;
  code: string;
  budget: string;
  procurement_method: string;
  department: string;
  notes: string;
  created_at: string;
  // Bidding v2 enriched fields (all 23 columns from 需求一览表 + 申报表/ERP)
  project_code?: string;
  req_app_no?: string;
  tech_spec_code?: string;
  lot_code?: string;
  lot_name?: string;
  package_no?: string;
  project_unit?: string;
  project_overview?: string;
  qualification_req?: string;
  performance_req?: string;
  staff_req?: string;
  quote_method?: string;
  subcontract_rule?: string;
  budget_yuan?: number | null;
  max_price_yuan?: number | null;
  max_price_raw?: string;
  declarant_name?: string;
  declarant_phone?: string;
  weight_ratio?: string;
  weight_warning?: string;
  weight_suggested?: string;
  price_formula?: string;
  price_params?: string;
  biz_eval_tpl?: string;
  tech_eval_tpl?: string;
  contract_version?: string;
  funding_source?: string;
  duration?: string;
  erp_sync_status?: string;
}

export interface BatchSupplier {
  id: string;
  batch_id: string;
  seq: number;
  supplier_name: string;
  package_no?: string;
  contact_name: string;
  mobile: string;
  phone: string;
  email: string;
  address: string;
  postal_code: string;
  tax_no: string;
  invoice_title: string;
  bank_name: string;
  bank_account: string;
}

export interface CollusionRisk {
  type: 'shared_mobile' | 'shared_email';
  label: string;
  value: string;
  suppliers: string[];
}

export interface BiddingSlotInfo {
  key: string;
  order: number;
  label: string;
  stage: string;
  description: string;
  file: UploadedFile | null;
  files?: UploadedFile[];
}

export interface ErpConflictFieldDiff {
  field: string;
  reqVal: string;
  erpVal: string;
}

export interface ErpConflictItem {
  type: 'erp_extra' | 'req_missing' | 'field_conflict';
  req_app_no: string;
  lot_info: string;
  project_name: string;
  fields?: ErpConflictFieldDiff[];
  summary: string;
}

export interface BatchChangeNotice {
  id: string;
  batch_id: string;
  title: string;
  notice_date: string;
  change_types: string[];
  summary: string;
  project_ids: string[];
  affected_packages?: Array<{
    lot_code: string;
    lot_name: string;
    package_no: string;
    project_name: string;
  }>;
  file_original_name: string;
  file_mime_type: string;
  has_file: boolean;
  created_at: string;
}

export interface BiddingWorkspace {
  batch: Batch;
  slots: BiddingSlotInfo[];
  files: UploadedFile[];
  projects: BatchProject[];
  suppliers: BatchSupplier[];
  collusionRisks: CollusionRisk[];
  erpConflicts?: ErpConflictItem[];
  changeNotices?: BatchChangeNotice[];
}

export interface ParsedSheetTable {
  name: string;
  headers: string[];
  rows: Record<string, string>[];
  totalRows: number;
}

export interface UploadedFile {
  id: string;
  batch_id: string;
  module: 'bidding' | 'evaluation' | 'award' | 'archive';
  slot_key?: string;
  category: string;
  original_name: string;
  stored_path: string;
  mime_type: string;
  file_size: number;
  parsed?: {
    sheets: ParsedSheetTable[];
    parsedAt: string;
  } | null;
  notes: string;
  created_at: string;
}

export interface SopItem {
  id: string;
  batch_id: string | null;
  is_template: number;
  item_type: 'step' | 'notice';
  stage_group: string;
  priority: 'normal' | 'critical';
  title: string;
  description: string;
  is_completed: number;
  completed_at: string | null;
  notes: string;
  sort_order: number;
}

export interface SopGuideImage {
  filename: string;
  url: string;
  originalName: string;
  caption?: string;
}

export interface SopGuide {
  id: string;
  stage_group: string;
  title: string;
  source_type: 'markdown' | 'pdf';
  content_md: string;
  images: SopGuideImage[];
  file_original_name: string;
  file_mime_type: string;
  file_size: number;
  has_file: boolean;
  sort_order: number;
  created_at: string;
  updated_at: string;
}

export interface EvalSlotInfo {
  key: string;
  order: number;
  label: string;
  description: string;
  file: UploadedFile | null;
}

export interface EvalRejectionRow {
  id: string;
  lot_code: string;
  lot_name: string;
  package_no: string;
  supplier_name: string;
  reason_category: string;
  rejection_tpl: string;
  doc_requirement: string;
  bid_response: string;
  review_type: string;
  notes: string;
}

export interface EvalPackageSummary {
  key: string;
  lot_code: string;
  lot_name: string;
  package_no: string;
  req_app_no: string;
  project_name: string;
  max_price_yuan: number | null;
  max_price_raw: string;
  quote_method?: string;
  sub_packages?: Array<{
    req_app_no: string;
    name: string;
    max_price_yuan: number | null;
    max_price_raw: string;
    quote_method: string;
  }>;
  over_limit_count?: number;
  bidder_count: number;
  rejection_count: number;
  detailed_eval_count: number;
  is_failed_lot: boolean;
  failed_reason?: string;
  eval_date: string;
  first_candidate: string;
  first_price_wan: string;
  first_score: number | null;
  discount_pct: string;
  bid_price_summary?: string;
  supplier_bids?: Array<{
    supplier_name: string;
    round1_price: string;
    round2_price: string;
    final_price: string;
    final_round: string;
    tax_rate: string;
    price_inc_tax_wan: number | null;
    price_ex_tax_wan: number | null;
    discount_rate: string;
    is_rejected: boolean;
    rejection_reason: string;
    is_over_limit?: boolean;
    over_limit_msg?: string;
    over_limit_details?: Array<{
      scope: '整包' | '子包';
      round_label: string;
      req_app_no: string;
      sub_name: string;
      limit_display: string;
      quote_display: string;
      exceed_text: string;
      msg: string;
    }>;
    sub_quotes?: Array<{
      req_app_no: string;
      sub_name: string;
      round_label: string;
      price_inc_tax_wan: number | null;
      discount_rate: string;
      limit_display: string;
      is_over_limit: boolean;
      over_limit_msg: string;
    }>;
  }>;
  candidates: Array<{
    id: string;
    supplier_name: string;
    tech_score: number | null;
    biz_score: number | null;
    price_score: number | null;
    total_score: number | null;
    rank_no: number | null;
    base_price: string;
    bid_price: string;
    eval_price: string;
  }>;
  bid_rounds: Array<{
    id: string;
    supplier_name: string;
    round_label: string;
    bid_price_wan: number | null;
    bid_price_raw: string;
    discount_rate: string;
  }>;
  line_quotes: Array<{
    id: string;
    supplier_name: string;
    round_label: string;
    req_app_no: string;
    tax_rate: string;
    price_ex_tax_wan: number | null;
    price_inc_tax_wan: number | null;
    discount_rate: string;
  }>;
  rejections: EvalRejectionRow[];
}

export interface TaxAuditItem {
  type: 'tax_abnormal' | 'tax_proof' | 'tax_missing';
  label: string;
  lot_code: string;
  lot_name: string;
  package_no: string;
  round_label: string;
  req_app_no: string;
  project_name: string;
  supplier_name: string;
  tax_rate_raw: string;
  detail: string;
}

export interface SameQuoteItem {
  scope?: '整包' | '子包';
  lot_code: string;
  lot_name: string;
  package_no: string;
  round_label: string;
  req_app_no: string;
  project_name: string;
  quote_basis: '含税总价(万元)' | '折扣率(%)';
  quote_value: string;
  suppliers: string[];
  detail: string;
}

export interface RoundIncreaseItem {
  scope?: '整包' | '子包';
  lot_code: string;
  lot_name: string;
  package_no: string;
  req_app_no: string;
  project_name: string;
  supplier_name: string;
  quote_basis: '含税总价(万元)' | '折扣率(%)';
  round1_value: string;
  round2_value: string;
  increase_text: string;
  detail: string;
}

export interface LowQuoteItem {
  lot_code: string;
  lot_name: string;
  package_no: string;
  project_name: string;
  supplier_name: string;
  round_label: string;
  quote_basis: '含税总价(万元)' | '折扣率(%)';
  quote_value: string;
  avg_value: string;
  second_lowest_value: string;
  max_price_value: string;
  triggered_rules: string[];
  detail: string;
}

export interface QuoteTaxAuditResult {
  taxCategory: '服务' | '物资';
  proofRates: number[];
  validRates: number[];
  taxAnomalies: TaxAuditItem[];
  taxProofNeeded: TaxAuditItem[];
  taxMissing: TaxAuditItem[];
  sameQuotes: SameQuoteItem[];
  roundIncreases: RoundIncreaseItem[];
  lowQuotes: LowQuoteItem[];
  counts: {
    taxAbnormal: number;
    taxProof: number;
    taxMissing: number;
    sameQuote: number;
    roundIncrease: number;
    lowQuote: number;
  };
}

export interface ContactOverlapItem {
  lot_code: string;
  lot_name: string;
  package_no: string;
  field_key: string;
  field_label: string;
  matched_value: string;
  suppliers: string[];
  alert_label: string;
}

export interface ContactMissingItem {
  lot_code: string;
  lot_name: string;
  package_no: string;
  supplier_name: string;
  detail: string;
}

export interface ContactAuditResult {
  overlaps: ContactOverlapItem[];
  missingContacts: ContactMissingItem[];
  counts: {
    overlapCount: number;
    missingCount: number;
  };
}

export interface HardwareAlertItem {
  lot_code: string;
  lot_name: string;
  package_no: string;
  supplier_a: string;
  file_a: string;
  round_a: string;
  source_a?: string;
  supplier_b: string;
  file_b: string;
  round_b: string;
  source_b?: string;
  cpu: string;
  disk: string;
  mac: string;
  ip_a: string;
  ip_b: string;
  detail: string;
}

export interface HardwareInsufficientItem {
  lot_code: string;
  lot_name: string;
  package_no: string;
  round_label: string;
  supplier_name: string;
  file_type: string;
  source_location?: string;
  ip: string;
  mac: string;
  cpu: string;
  disk: string;
  detail: string;
}

export interface HardwareManualReviewItem extends HardwareAlertItem {
  mac_a: string;
  mac_b: string;
  review_reason: string;
}

export interface HardwareAuditResult {
  totalRecords: number;
  alerts: HardwareAlertItem[];
  manualReview?: HardwareManualReviewItem[];
  insufficient: HardwareInsufficientItem[];
  counts: {
    alertCount: number;
    manualReviewCount?: number;
    insufficientCount: number;
  };
}

export interface EvaluationWorkspace {
  batch: Batch;
  slots: EvalSlotInfo[];
  otherFiles: UploadedFile[];
  sopItems: SopItem[];
  packages: EvalPackageSummary[];
  rejections: EvalRejectionRow[];
  changeNotices?: BatchChangeNotice[];
  auditConfig?: {
    tax_category: '服务' | '物资';
    proof_rates: number[];
  };
  quoteTaxAudit?: QuoteTaxAuditResult;
  contactAudit?: ContactAuditResult;
  hardwareAudit?: HardwareAuditResult;
  stats: {
    bidRowCount: number;
    lineQuoteCount: number;
    rejectionCount: number;
    rankingRowCount: number;
    ipRowCount?: number;
  };
}

export interface ExpertTotalIssue {
  section: string;
  package_name: string;
  supplier: string;
  majority_same: string;
  diff_exceeded: string;
  source_file: string;
  supplier_location: string;
}

export interface ExpertTotalAuditResult {
  checkedCount: number;
  fileCount: number;
  sheetCount: number;
  expertNum: number | null;
  maxScoreDiff: number;
  items: ExpertTotalIssue[];
  counts: {
    totalIssues: number;
    majoritySame: number;
    diffExceeded: number;
  };
}

export interface ObjectiveCriterion {
  name: string;
  project: string;
}

export interface ObjectiveScoringMethod {
  id: string;
  criteria: ObjectiveCriterion[];
  packages: string[];
}

export interface ObjectiveSelection {
  methodId: string;
  criterion: string;
  step?: number | null;
}

export interface ObjectiveAuditIssue {
  section: string;
  package_name: string;
  supplier: string;
  criterion: string;
  issue_type: '分数不一致' | '分值步长不符' | '专家记录缺失' | '专家数量不足' | '细则未找到';
  detail: string;
  source_location: string;
}

export interface ExpertObjectiveAuditResult {
  checkedCount: number;
  fileCount: number;
  items: ObjectiveAuditIssue[];
  warnings: string[];
  counts: {
    totalIssues: number;
    scoreMismatch: number;
    stepMismatch: number;
    expertMissing: number;
    expertInsufficient: number;
    criterionMissing: number;
  };
}

export interface ScoreCheckFileItem {
  originalName: string;
  size: number;
}

export interface ScoreCheckState {
  total: {
    files: ScoreCheckFileItem[];
    config: { expertNum?: number | null; maxScoreDiff?: number };
    result: ExpertTotalAuditResult | null;
    updated_at: string;
  } | null;
  objective: {
    files: ScoreCheckFileItem[];
    config: { selections?: ObjectiveSelection[]; sectionOverride?: string };
    methods: ObjectiveScoringMethod[];
    result: ExpertObjectiveAuditResult | null;
    updated_at: string;
  } | null;
}

export interface PrecheckScoringColumnMeta {
  key: string;
  colLetter: string;
  level1: string;
  level2?: string;
  detail?: string;
  maxScore?: string;
  minScore?: string;
  stepHint?: string;
}

export interface PrecheckSheetMeta {
  sheetName: string;
  headerRow: number;
  scoringColumns: string[];
  scoringColumnDetails?: PrecheckScoringColumnMeta[];
  totalColumn: string;
  templateMarkedColumns: string[];
}

export interface PrecheckObjectiveSelection {
  sheetName: string;
  criterion: string;
  step?: number | null;
}

export interface PrecheckTotalIssue {
  sheet_name: string;
  section: string;
  package_name: string;
  supplier: string;
  majority_same: string;
  diff_exceeded: string;
  expert_scores: string;
  source_location: string;
}

export interface PrecheckObjectiveIssue {
  sheet_name: string;
  section: string;
  package_name: string;
  supplier: string;
  criterion: string;
  issue_type: '分数不一致' | '分值步长不符' | '专家记录缺失' | '专家数量不足' | '细则未找到';
  detail: string;
  source_location: string;
}

export interface PrecheckExcludedItem {
  sheet_name: string;
  section: string;
  package_name: string;
  supplier: string;
  review_type: string;
  reason_category: string;
}

export interface ExpertPrecheckResult {
  fileCount: number;
  sheetCount: number;
  expertNames: string[];
  maxScoreDiff: number;
  excludeRejected?: boolean;
  excludedRejectedCount?: number;
  excludedRejectedItems?: PrecheckExcludedItem[];
  totalCheckedCount: number;
  objectiveCheckedCount: number;
  totalItems: PrecheckTotalIssue[];
  objectiveItems: PrecheckObjectiveIssue[];
  warnings: string[];
  counts: {
    totalIssues: number;
    majoritySame: number;
    diffExceeded: number;
    objectiveIssues: number;
    scoreMismatch: number;
    stepMismatch: number;
    expertMissing: number;
    expertInsufficient: number;
    criterionMissing: number;
  };
}

export interface ExpertPrecheckState {
  files: ScoreCheckFileItem[];
  templateFile: { originalName: string } | null;
  sheets: PrecheckSheetMeta[];
  batchRejectionCount?: number;
  hasRejectionFile?: boolean;
  config: {
    maxScoreDiff?: number;
    excludeRejected?: boolean;
    objectiveSelections?: PrecheckObjectiveSelection[];
  };
  result: ExpertPrecheckResult | null;
  updated_at: string;
}

export interface InitialReportSectionSummary {
  section: string;
  reason: string;
  items: Array<{ company: string; displayName: string; packs: string }>;
  sentence: string;
}

export interface InitialReportResult {
  fileName: string;
  dept: string;
  leader: string;
  members: string;
  sectionCount: number;
  packCount: number;
  bidderCount: number;
  documentCount: number;
  rejectionRowCount: number;
  clarification: string;
  simplifyNames: boolean;
  companySuffixes: string;
  paragraphs: string[];
  groupedRejections: InitialReportSectionSummary[];
  warnings: string[];
  sheetNames: {
    detailSheet: string;
    summarySheet: string;
    coverSheet: string;
  };
}

export interface InitialReportState {
  file: ScoreCheckFileItem | null;
  config: {
    department?: string;
    clarification?: string;
    simplifyNames?: boolean;
    companySuffixes?: string;
  };
  result: InitialReportResult | null;
  updated_at: string;
}

export interface WarningFormItem {
  prjName: string;
  prjNo: string;
  labelNo: string;
  section: string;
  group: string;
  packs: string;
  suppliers: string[];
  flags: [boolean, boolean, boolean];
  values: [string, string, string];
  warnMatters: string;
  reasonLabels: string[];
  docxPath: string;
  docxFileName: string;
  date: string;
}

export interface WarningFormAuditResult {
  date: string;
  riskRowCount: number;
  groupRowCount: number;
  bidRowCount: number;
  formCount: number;
  forms: WarningFormItem[];
  warnings: string[];
  readmeText: string;
}

export interface WarningFormState {
  files: {
    risk: ScoreCheckFileItem | null;
    groups: ScoreCheckFileItem | null;
    bids: ScoreCheckFileItem | null;
  };
  config: { date?: string };
  result: WarningFormAuditResult | null;
  updated_at: string;
}

export type MinutesProcurementType = '公开' | '非公开';

export interface MinutesRejectionRow {
  seq: number;
  lot_name: string;
  package_no: string;
  supplier_name: string;
  reason: string;
  bid_response: string;
  review_type: string;
}

export interface InitialMinutesConfig {
  procurementType: MinutesProcurementType;
  batchName: string;
  meetingTime: string;
  groupName: string;
  clarificationNote: string;
  reviewTypeFilter?: 'all' | '商务' | '技术';
}

export interface InitialMinutesResult {
  fileName: string;
  sheetName: string;
  headerRow: number;
  config: InitialMinutesConfig;
  totalRows: number;
  hasReviewTypeCol: boolean;
  rows: MinutesRejectionRow[];
  previewParagraphs: string[];
}

export interface InitialMinutesState {
  file: ScoreCheckFileItem | null;
  config: InitialMinutesConfig;
  result: InitialMinutesResult | null;
  updated_at: string;
}

export interface PriceScoreRow {
  section: string;
  pack: string;
  round: string;
  supplier: string;
  quote_basis: string;
  quote_value: string;
  method_name: string;
  base_value: string;
  price_score: string;
  weighted_score: string;
  source: string;
}

export interface PriceScoreState {
  files: { quotes: ScoreCheckFileItem[]; rules: ScoreCheckFileItem | null };
  config: { round?: string };
  result: {
    quoteFiles: string[];
    ruleFile: string;
    roundFilter: string;
    totalRows: number;
    packageCount: number;
    rows: PriceScoreRow[];
    warnings: string[];
  } | null;
  updated_at: string;
}

export interface PerformanceRow {
  project_code: string;
  lot_name: string;
  lot_code: string;
  pack_name: string;
  pack_code: string;
  project_unit: string;
  bidder_name: string;
  unit_name: string;
  department: string;
  evaluation: string;
  score: number;
  remark: string;
  match_note: string;
}

export interface PerformanceState {
  files: { bids: ScoreCheckFileItem | null; scores: ScoreCheckFileItem | null };
  result: {
    bidFile: string;
    scoreFile: string;
    totalRows: number;
    matchedCount: number;
    reviewCount: number;
    rows: PerformanceRow[];
    warnings: string[];
  } | null;
  updated_at: string;
}

export interface EconomicRankIssue {
  section: string;
  pack: string;
  supplier: string;
  original_rank: string;
  verified_rank: string;
  reason: string;
}

export interface EconomicRankState {
  files: ScoreCheckFileItem[];
  result: {
    files: string[];
    sheetCount: number;
    examinedCount: number;
    issueCount: number;
    issues: EconomicRankIssue[];
  } | null;
  updated_at: string;
}

export type AwardMode = '无限授' | '一轮一包' | '一轮多包';

export interface AwardResultRow {
  section_code: string;
  section_name: string;
  pack_name: string;
  limit_group: string;
  winner: string;
  status: string;
  round: string;
  mode: string;
  note: string;
  source: string;
}

export interface AwardAllocState {
  files: { rankings: ScoreCheckFileItem[]; groups: ScoreCheckFileItem | null };
  config: { mode?: AwardMode };
  result: {
    files: string[];
    groupFile: string;
    mode: AwardMode;
    totalPacks: number;
    awardedCount: number;
    unallocatedCount: number;
    rows: AwardResultRow[];
  } | null;
  updated_at: string;
}

export type SplitPackMode = '比例' | '金额';

export interface SplitSummaryRow {
  pack_no: string;
  item_count: number;
  total_qty: string;
  total_budget: string;
  target_amount: string;
  diff_amount: string;
}

export interface SplitDetailRow {
  seq: number;
  pack_no: string;
  material_code: string;
  description: string;
  long_description: string;
  specification: string;
  extension: string;
  quantity: string;
  unit: string;
  unit_price: string;
  total_price: string;
}

export interface SplitPackState {
  files: { materials: ScoreCheckFileItem | null; rules: ScoreCheckFileItem | null };
  config: { mode?: SplitPackMode; weightStep?: string; lengthStep?: string };
  result: {
    materialFile: string;
    ruleFile: string;
    mode: SplitPackMode;
    weightStep: string;
    lengthStep: string;
    packCount: number;
    materialCount: number;
    summaryRows: SplitSummaryRow[];
    detailRows: SplitDetailRow[];
    warnings: string[];
  } | null;
  updated_at: string;
}

export interface ClientAgreement {
  id: string;
  group_name: string;
  client_name: string;
  contact_name: string;
  contact_phone: string;
  contract_expiry: string;
  discount_rate: number;
  discount_label: string;
  agreement_status: string;
  notes: string;
  sort_order: number;
}

export interface AwardConfig {
  client_id: string;
  owner_unit: string;
  category: '服务' | '工程' | '物资';
  discount_rate: number;
  discount_label: string;
  batch_manager: string;
  bid_location: string;
  publicity_date: string;
  announcement_date: string;
}

export interface AwardPackageItem {
  seq: number;
  lot_code: string;
  lot_name: string;
  package_no: string;
  project_name: string;
  project_overview: string;
  req_app_no: string;
  budget_wan: number | null;
  max_price_wan: number | null;
  max_price_raw: string;
  declarant_info: string;
  material_desc: string;
  is_failed_lot: boolean;
  winner_name: string;
  winner_price_wan: number | null;
  winner_price_raw: string;
  winner_score: number | null;
  winner_contact: string;
  winner_mobile: string;
  winner_email: string;
  winner_address: string;
  service_fee_pre_yuan: number | null;
  service_fee_yuan: number | null;
  platform_fee_yuan: number | null;
  category: '服务' | '工程' | '物资';
  discount_rate: number;
  remark: string;
}

export interface AwardTemplate {
  key: string;
  title: string;
  file_type: 'xlsx' | 'docx';
  description: string;
  sources: string[];
}

export interface AwardOutput {
  id: string;
  batch_id: string;
  template_key: string;
  title: string;
  file_type: string;
  filename: string;
  version: number;
  created_at: string;
}

export interface AwardOverview {
  readiness: {
    hasBatchInfo: boolean;
    projectCount: number;
    biddingFileCount: number;
    evalFileCount: number;
  };
  batch: Batch;
  clients: ClientAgreement[];
  awardConfig: AwardConfig;
  packages: AwardPackageItem[];
  totals: {
    winningCount: number;
    failedCount: number;
    totalAwardWan: number;
    totalServiceFeePreYuan: number;
    totalServiceFeeYuan: number;
    totalPlatformFeeYuan: number;
  };
  templates: AwardTemplate[];
  outputs: AwardOutput[];
}

export interface ArchiveItem {
  id: string;
  batch_id: string;
  sheet_name: string;
  seq: string;
  doc_name: string;
  is_required: boolean;
  required_source: string;
  media_type: string;
  paper_req: string;
  electronic_req: string;
  policy_basis: string;
  template_source: string;
  status: 'present' | 'missing' | 'na';
  is_manual_marked: boolean;
  linked_file_ids: string[];
  linked_files: Array<{
    id: string;
    module: string;
    category: string;
    original_name: string;
    file_size: number;
    created_at: string;
  }>;
  notes: string;
  sort_order: number;
}

export interface ArchiveOverview {
  sheetNames: string[];
  items: ArchiveItem[];
  availableBatchFiles: Array<{
    id: string;
    module: string;
    category: string;
    original_name: string;
    file_size: number;
    created_at: string;
  }>;
  summary: {
    total: number;
    activeTotal: number;
    present: number;
    missing: number;
    na: number;
    requiredTotal: number;
    requiredPresent: number;
    requiredMissing: number;
    optionalTotal: number;
    optionalPresent: number;
  };
}

export interface ParsedArchivePreviewRow {
  sheetName: string;
  seq: string;
  docName: string;
  isRequired: boolean;
  requiredSource: 'color' | 'keyword' | 'manual';
  highlightColor?: string;
  mediaType: string;
  paperReq: string;
  electronicReq: string;
  policyBasis: string;
  templateSource: string;
  status: 'present' | 'missing' | 'na';
  notes: string;
}
