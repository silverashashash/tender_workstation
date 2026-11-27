import { v4 as uuid } from 'uuid';
import { db } from './db.js';
import { DEFAULT_CLIENT_AGREEMENTS } from './serviceFee.js';

export const DEFAULT_SOP_STEPS = [
  {
    item_type: 'step',
    stage_group: '开标前准备',
    priority: 'normal',
    title: '打印评审纪律、廉洁保密承诺书与会议签到表',
    description: '按身份分别准备三份《廉洁保密承诺书》：评审委员会、现场监督人员、代理及工作人员。',
  },
  {
    item_type: 'step',
    stage_group: '内网开标',
    priority: 'normal',
    title: '收取手机及通信工具并登记保管记录',
    description: '填写《采购通信工具保管记录》（现场签署，留存 6 个月后销毁）。',
  },
  {
    item_type: 'step',
    stage_group: '内网开标',
    priority: 'normal',
    title: 'ECP2.0 解密应答文件并导出送达与密封情况表',
    description: '导出后增加“密封情况”列（密封完好），同时导出《批次行信息报价基础数据表》。',
  },
  {
    item_type: 'step',
    stage_group: '评审过程',
    priority: 'normal',
    title: '核查国网及省公司供应商不良行为与信用记录',
    description: '整理当前批次生效的不良行为清单，填写《项目信用情况查询记录表》并由查询人签字。',
  },
  {
    item_type: 'step',
    stage_group: '评审过程',
    priority: 'normal',
    title: '跟进初评、澄清发出及多轮报价截图留痕',
    description: '凡发出澄清或多轮报价，务必在 ECP2.0 截全图留存；涉及否决项须同步生成《否决审批单》。',
  },
  {
    item_type: 'step',
    stage_group: '评审过程',
    priority: 'normal',
    title: '核对技术/商务评分一览表、评分统计表与初评小结',
    description: '逐项复核专家打分汇总无算术错误，确认专家签字页无遗漏。',
  },
  {
    item_type: 'step',
    stage_group: '评审结束与资料带出',
    priority: 'normal',
    title: '导出正式评审报告、技术经济排序表与授标结果排序表',
    description: '回收评审报告原件（含全体评委签字页）、监督记录表/报告、法律保障工作报告。',
  },
  {
    item_type: 'step',
    stage_group: '评审结束与资料带出',
    priority: 'normal',
    title: '刻录全量应答文件光盘并清理内网临时目录',
    description: '成交人与未成交人应答文件分开刻录；按保密要求清除评审终端缓存，完成专家评价截图。',
  },
];

export const DEFAULT_SOP_NOTICES = [
  {
    item_type: 'notice',
    stage_group: '内网开标',
    priority: 'critical',
    title: '内网离线数据带出前必须逐项核对完整性',
    description: '开评标在内网完成且无法随时回查，离开现场前务必确认报价基础数据表、评分统计表、澄清截图及排序表已完整导出并拷贝。',
  },
  {
    item_type: 'notice',
    stage_group: '评审过程',
    priority: 'critical',
    title: '严防专家签字页遗漏（初评小结 / 评分统计表 / 评审报告）',
    description: '离场前逐页清点全体评委及监督人员签字，避免出现“缺某专家签字”导致归档卡壳。',
  },
  {
    item_type: 'notice',
    stage_group: '评审过程',
    priority: 'critical',
    title: '否决投标条款必须与采购文件实质性条款严格对应',
    description: '每一项否决均须在 ECP2.0 导出《否决审批单》，并在预成交公示中准确披露否决依据。',
  },
  {
    item_type: 'notice',
    stage_group: '开标前准备',
    priority: 'normal',
    title: '三类人员《廉洁保密承诺书》需分表签署',
    description: '评审委员会、现场监督人员、代理及技术支持会务人员须分别签署原件，不可混签。',
  },
  {
    item_type: 'notice',
    stage_group: '评审结束与资料带出',
    priority: 'normal',
    title: '6 个月销毁类过程资料单独存放',
    description: '通信工具保管记录、现场电话/传真记录及音视频监控资料需按 6 个月周期单独台账管理并留存销毁记录。',
  },
];

export interface ArchiveTemplateSeed {
  sheet_name: string;
  seq: string;
  doc_name: string;
  is_required: number;
  media_type: string;
  paper_req: string;
  electronic_req: string;
  policy_basis: string;
  template_source: string;
  default_notes?: string;
}

export const DEFAULT_ARCHIVE_TEMPLATES: ArchiveTemplateSeed[] = [
  // Sheet 1: 归档文件 (20 items from user's Excel)
  { sheet_name: '归档文件', seq: '0', doc_name: '封面+备考表+目录', is_required: 1, media_type: '纸质', paper_req: '包含批次基本信息，版式完全一致', electronic_req: '', policy_basis: '', template_source: '按照省公司档案管理要求制作封面' },
  { sheet_name: '归档文件', seq: '1', doc_name: '采购代理服务合同', is_required: 1, media_type: '纸质+电子', paper_req: '原件复印', electronic_req: '原件扫描', policy_basis: '国网（物资/4）243-2024 第十六条（一）', template_source: '数字化法治企业建设平台采购代理服务合同模板' },
  { sheet_name: '归档文件', seq: '2', doc_name: '采购计划及审批文件（年度集中采购目录清单和批次安排）', is_required: 1, media_type: '纸质+电子', paper_req: '1.正式文件（带文号）打印；2.批次安排表加盖部门公章。', electronic_req: '1.正式文件（带文号）；2.原件扫描。', policy_basis: '国网（物资/4）243-2024 第十六条（二）', template_source: '' },
  { sheet_name: '归档文件', seq: '3', doc_name: '授权采购“一事一授权”审批单（如有）', is_required: 0, media_type: '纸质+电子', paper_req: '原件（一站式平台）', electronic_req: '原件扫描', policy_basis: '国网（物资/4）243-2024 第十六条（二）', template_source: '国网（物资/2）123-2024 附件1' },
  { sheet_name: '归档文件', seq: '4', doc_name: '发布采购公告（或邀请函）的公司内部审批单', is_required: 1, media_type: '纸质+电子', paper_req: '原件', electronic_req: '原件扫描', policy_basis: '国网（物资/4）243-2024 第十六条（三）', template_source: '国网（物资/4）239-2024 附件7采购文件审批单' },
  { sheet_name: '归档文件', seq: '5', doc_name: '采购文件', is_required: 1, media_type: '纸质+电子', paper_req: '1.正式发布版采购文件(盖封面章)，双面打印，不含技术规范书；2.技术规范书(光盘)；3.合同及其它附件(光盘)。', electronic_req: '1.正式发布版采购文件（带封面章，PDF）；2.正式发布版技术规范书。', policy_basis: '国网（物资/4）243-2024 第十六条（四）', template_source: '采购文件正文应用国网范本' },
  { sheet_name: '归档文件', seq: '6', doc_name: '采购文件澄清、修改函件（如有）', is_required: 0, media_type: '纸质+电子', paper_req: '正式发布版澄清、修改函件(盖封面章)', electronic_req: '带封面章PDF', policy_basis: '国网（物资/4）243-2024 第十六条（四）', template_source: '' },
  { sheet_name: '归档文件', seq: '7', doc_name: '采购文件澄清、修改函件的内部审批单（如有）', is_required: 0, media_type: '纸质+电子', paper_req: '原件', electronic_req: '原件扫描', policy_basis: '国网（物资/4）243-2024 第十六条（五）', template_source: '国网（物资/4）239-2024 附件8' },
  { sheet_name: '归档文件', seq: '8', doc_name: '项目管理部门或项目单位出具的最高限价（如有）', is_required: 1, media_type: '纸质+电子', paper_req: '原件（三级审批+部门公章）', electronic_req: '原件扫描', policy_basis: '国网（物资/4）243-2024 第十六条（六）', template_source: '国网（物资/4）239-2024 附件6' },
  { sheet_name: '归档文件', seq: '9', doc_name: '应答文件的送达（上传交易平台信息系统）时间和密封（加密）情况', is_required: 1, media_type: '纸质+电子', paper_req: 'ECP2.0导出打印(盖封面章)，双面打印', electronic_req: '原件扫描', policy_basis: '国网（物资/4）243-2024 第十六条（七）', template_source: 'ECP2.0导出，增加“密封情况”列（密封完好）' },
  { sheet_name: '归档文件', seq: '10', doc_name: '评审小组组建名单', is_required: 1, media_type: '纸质+电子', paper_req: '正式行文的评审小组组建文件（带文号）打印', electronic_req: '正式行文的评审小组组建文件（带文号）', policy_basis: '国网（物资/4）243-2024 第十六条（八）', template_source: '国网（物资/4）239-2024 附件9' },
  { sheet_name: '归档文件', seq: '11', doc_name: '评审小组组建审批单', is_required: 1, media_type: '纸质+电子', paper_req: '原件', electronic_req: '原件扫描', policy_basis: '国网（物资/4）243-2024 第十六条（八）', template_source: '物资招标〔2023〕23号 附件12' },
  { sheet_name: '归档文件', seq: '12', doc_name: '评审纪律', is_required: 1, media_type: '纸质+电子', paper_req: '带批次名称打印(盖封面章)', electronic_req: '原件扫描', policy_basis: '国网（物资/4）243-2024 第十六条（九）', template_source: '物资招标〔2023〕23号 附件13' },
  { sheet_name: '归档文件', seq: '13', doc_name: '评审报告', is_required: 1, media_type: '纸质+电子', paper_req: '正式评审报告(盖封面章)，双面打印；评委会签字页为原件', electronic_req: '正式评审报告（带章封面及评委会签字页扫描，PDF）', policy_basis: '国网（物资/4）243-2024 第十六条（十）', template_source: '国网（物资/4）239-2024 附件12' },
  { sheet_name: '归档文件', seq: '14', doc_name: '定标审批单', is_required: 1, media_type: '纸质+电子', paper_req: '原件', electronic_req: '原件扫描', policy_basis: '国网（物资/4）243-2024 第十六条（十一）', template_source: '物资招标〔2023〕23号 附件20' },
  { sheet_name: '归档文件', seq: '15', doc_name: '会议纪要（招投标/授权采购工作领导小组）', is_required: 1, media_type: '纸质+电子', paper_req: '正式行文的领导小组会议纪要（带文号）打印', electronic_req: '正式行文的会议纪要（带文号）', policy_basis: '国网（物资/4）243-2024 第十六条（十一）', template_source: '物资招标〔2023〕23号 附件20' },
  { sheet_name: '归档文件', seq: '16', doc_name: '预成交公示', is_required: 1, media_type: '纸质+电子', paper_req: '正式预成交公示及否决事项(盖封面章)；系统截图(盖封面章)', electronic_req: '原件扫描', policy_basis: '国网（物资/4）243-2024 第十六条（十二）', template_source: '省公司自拟预成交公示及否决事项模板' },
  { sheet_name: '归档文件', seq: '17', doc_name: '预成交公示异议和答复（如有）', is_required: 0, media_type: '纸质+电子', paper_req: '异议和答复原件；异议属实转流标审批单', electronic_req: '原件扫描', policy_basis: '国网（物资/4）243-2024 第十六条（十二）', template_source: '国网（物资/4）239-2024 附件13' },
  { sheet_name: '归档文件', seq: '18', doc_name: '成交通知书', is_required: 1, media_type: '纸质+电子', paper_req: 'ECP2.0导出打印（服务类包括分项报价），须包括采购人及代理机构章', electronic_req: 'ECP2.0导出', policy_basis: '国网（物资/4）243-2024 第十六条（十三）', template_source: 'ECP2.0导出' },
  { sheet_name: '归档文件', seq: '19', doc_name: '成交人的应答文件及其澄清回复函件', is_required: 1, media_type: '纸质+电子', paper_req: '用于现场评审全量应答文件及其澄清回复函件（光盘）', electronic_req: '全量应答文件及澄清回复函件', policy_basis: '国网（物资/4）243-2024 第十六条（十四）', template_source: '' },

  // Sheet 2: 重要保管文件 (Representative key items from user's Excel 0~36)
  { sheet_name: '重要保管文件', seq: '0', doc_name: '封面+备考表+目录', is_required: 1, media_type: '纸质', paper_req: '包含批次基本信息，版式完全一致', electronic_req: '', policy_basis: '', template_source: '按照省公司档案管理要求制作封面' },
  { sheet_name: '重要保管文件', seq: '1', doc_name: '采购工作方案（如有）', is_required: 0, media_type: '纸质+电子', paper_req: '正式采购方案(盖封面章)，双面打印', electronic_req: '带章封面扫描PDF', policy_basis: '国网（物资/4）243-2024 第二十条（一）', template_source: '省公司自拟方案模板' },
  { sheet_name: '重要保管文件', seq: '2', doc_name: '采购文件审查会议纪要', is_required: 1, media_type: '纸质+电子', paper_req: '审查会议纪要(盖封面章)及签到表原件', electronic_req: '带章封面及签到表PDF', policy_basis: '国网（物资/4）243-2024 第二十条（二）', template_source: '省公司自拟审查会议纪要模板' },
  { sheet_name: '重要保管文件', seq: '5', doc_name: '廉洁保密承诺书', is_required: 1, media_type: '纸质+电子', paper_req: '原件（评委会/现场监督/代理工作人员分签）', electronic_req: '原件扫描', policy_basis: '国网（物资/4）243-2024 第二十条（五）', template_source: '物资监察〔2024〕9号' },
  { sheet_name: '重要保管文件', seq: '6', doc_name: '监督记录表、监督报告', is_required: 1, media_type: '纸质+电子', paper_req: '原件', electronic_req: '原件扫描', policy_basis: '国网（物资/4）243-2024 第二十条（六）', template_source: '国网（物资/4）143-2018' },
  { sheet_name: '重要保管文件', seq: '7', doc_name: '未成交人的应答文件及其澄清回复函件', is_required: 1, media_type: '纸质+电子', paper_req: '全量应答文件光盘（与归档文件分开刻录）', electronic_req: '全量应答文件及澄清函件', policy_basis: '国网（物资/4）243-2024 第二十条（七）', template_source: '' },
  { sheet_name: '重要保管文件', seq: '8', doc_name: '下浮系数抽取表（如有）', is_required: 0, media_type: '纸质+电子', paper_req: '原件', electronic_req: '原件扫描', policy_basis: '过程重点保管资料', template_source: '省公司自拟模板' },
  { sheet_name: '重要保管文件', seq: '9', doc_name: '批次行信息报价基础数据表', is_required: 1, media_type: '电子', paper_req: '', electronic_req: 'ECP2.0导出（带章封面扫描PDF）', policy_basis: '过程重点保管资料', template_source: 'ECP2.0导出' },
  { sheet_name: '重要保管文件', seq: '10', doc_name: '采购回避承诺书', is_required: 1, media_type: '纸质+电子', paper_req: '原件', electronic_req: '原件扫描', policy_basis: '过程重点保管资料', template_source: '' },
  { sheet_name: '重要保管文件', seq: '12', doc_name: '供应商不良行为处理文件', is_required: 1, media_type: '电子', paper_req: '', electronic_req: '国网不良行为、省公司不良行为当前批次发布内容', policy_basis: '过程重点保管资料', template_source: 'ECP2.0导出' },
  { sheet_name: '重要保管文件', seq: '13', doc_name: '技术（商务）组供应商应答响应情况一览表', is_required: 1, media_type: '纸质或电子', paper_req: '原件带专家签字', electronic_req: '原件扫描或带专家电子签名', policy_basis: '过程重点保管资料', template_source: '省公司自拟模板' },
  { sheet_name: '重要保管文件', seq: '14', doc_name: '项目信用情况查询记录表', is_required: 1, media_type: '纸质或电子', paper_req: '原件带查询人签字', electronic_req: '原件扫描或带查询人电子签名', policy_basis: '过程重点保管资料', template_source: '省公司自拟模板' },
  { sheet_name: '重要保管文件', seq: '16', doc_name: '技术（商务）X组评分一览表', is_required: 1, media_type: '电子', paper_req: '', electronic_req: 'ECP2.0导出转pdf', policy_basis: '过程重点保管资料', template_source: 'ECP2.0导出' },
  { sheet_name: '重要保管文件', seq: '17', doc_name: '技术（商务）X组评分统计表', is_required: 1, media_type: '纸质+电子', paper_req: '原件带专家签字', electronic_req: '原件扫描或带专家电子签名', policy_basis: '过程重点保管资料', template_source: 'ECP2.0导出' },
  { sheet_name: '重要保管文件', seq: '18', doc_name: '技术（商务）X组初评小结', is_required: 1, media_type: '纸质或电子', paper_req: '原件带专家签字', electronic_req: '原件扫描或带专家电子签名', policy_basis: '过程重点保管资料', template_source: '省公司自拟模板' },
  { sheet_name: '重要保管文件', seq: '19', doc_name: '法律保障工作报告', is_required: 1, media_type: '纸质+电子', paper_req: '原件', electronic_req: '原件扫描', policy_basis: '过程重点保管资料', template_source: '省公司自拟模板' },
  { sheet_name: '重要保管文件', seq: '22', doc_name: '技术经济排序表和授标结果排序表', is_required: 1, media_type: '电子', paper_req: '', electronic_req: 'ECP2.0导出转pdf', policy_basis: '过程重点保管资料', template_source: 'ECP2.0导出' },
  { sheet_name: '重要保管文件', seq: '23', doc_name: '招投标（授权采购）工作领导小组会议签到表', is_required: 1, media_type: '纸质', paper_req: '原件', electronic_req: '', policy_basis: '过程重点保管资料', template_source: '省公司自拟会议签到表模板' },
  { sheet_name: '重要保管文件', seq: '24', doc_name: '成交结果公告、成交结果公告发布确认', is_required: 1, media_type: '纸质+电子', paper_req: '正式中标结果公告(盖封面章)；系统截图(盖封面章)', electronic_req: '原件扫描', policy_basis: '过程重点保管资料', template_source: '系统导出' },
  { sheet_name: '重要保管文件', seq: '29', doc_name: '采购公告发布确认单', is_required: 1, media_type: '纸质+电子', paper_req: '包含ECP2.0、公共服务平台截图及负责人复核人签字', electronic_req: '原件扫描', policy_basis: '过程重点保管资料', template_source: '省公司自拟确认单模板' },
  { sheet_name: '重要保管文件', seq: '34', doc_name: '采购通信工具保管记录', is_required: 0, media_type: '纸质', paper_req: '现场签署资料（6个月销毁）', electronic_req: '', policy_basis: '过程重点保管资料', template_source: '省公司自拟模板' },
  { sheet_name: '重要保管文件', seq: '36', doc_name: '现场音视频监控资料', is_required: 0, media_type: '电子', paper_req: '代理机构统一保存（6个月销毁，保存销毁记录）', electronic_req: '音视频为电子，销毁记录扫描', policy_basis: '过程重点保管资料', template_source: '监控导出' },
];

export function seedIfEmpty() {
  // Always ensure client_agreements table is populated (even for existing DBs)
  const clientCnt = db.prepare('SELECT COUNT(*) as cnt FROM client_agreements').get() as { cnt: number };
  if (clientCnt.cnt === 0) {
    const insClient = db.prepare(`
      INSERT INTO client_agreements (id, group_name, client_name, contact_name, contact_phone, contract_expiry, discount_rate, discount_label, agreement_status, notes, sort_order)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);
    DEFAULT_CLIENT_AGREEMENTS.forEach((c, idx) => {
      insClient.run(uuid(), c.group_name, c.client_name, c.contact_name, c.contact_phone, c.contract_expiry, c.discount_rate, c.discount_label, c.agreement_status, c.notes, idx + 1);
    });
  }

  const countRow = db.prepare('SELECT COUNT(*) as cnt FROM batches').get() as { cnt: number };
  if (countRow.cnt > 0) return;

  const now = new Date().toISOString();

  // 1. Seed global SOP templates
  const insertSop = db.prepare(`
    INSERT INTO sop_items (id, batch_id, is_template, item_type, stage_group, priority, title, description, is_completed, completed_at, notes, sort_order)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);

  DEFAULT_SOP_STEPS.forEach((s, idx) => {
    insertSop.run(uuid(), null, 1, s.item_type, s.stage_group, s.priority, s.title, s.description, 0, null, '', idx + 1);
  });
  DEFAULT_SOP_NOTICES.forEach((n, idx) => {
    insertSop.run(uuid(), null, 1, n.item_type, n.stage_group, n.priority, n.title, n.description, 0, null, '', idx + 1);
  });

  // 2. Seed global Archive templates
  const insertArchive = db.prepare(`
    INSERT INTO archive_items (id, batch_id, is_template, sheet_name, seq, doc_name, is_required, required_source, media_type, paper_req, electronic_req, policy_basis, template_source, status, is_manual_marked, linked_file_ids_json, notes, sort_order)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);

  DEFAULT_ARCHIVE_TEMPLATES.forEach((a, idx) => {
    insertArchive.run(
      uuid(), null, 1, a.sheet_name, a.seq, a.doc_name, a.is_required,
      a.is_required ? 'color' : 'manual', a.media_type, a.paper_req, a.electronic_req,
      a.policy_basis, a.template_source, 'missing', 0, '[]', '', idx + 1
    );
  });

  // 3. Seed 3 realistic batches at different stages
  const insertBatch = db.prepare(`
    INSERT INTO batches (id, name, code, project_unit, announcement_time, bid_opening_time, status, notes, extra_fields_json, created_at, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);
  const insertProject = db.prepare(`
    INSERT INTO batch_projects (id, batch_id, seq, name, code, budget, procurement_method, department, notes, created_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);
  const insertFile = db.prepare(`
    INSERT INTO uploaded_files (id, batch_id, module, category, original_name, stored_path, mime_type, file_size, parsed_json, notes, created_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);

  // --- Batch 1: In Archiving stage ---
  const b1 = uuid();
  insertBatch.run(
    b1,
    '2026年第四季度信息通信运维服务授权采购批次',
    'GW-AH-2026-0412',
    '国网安徽省电力有限公司信通分公司',
    '2026-10-18 10:00',
    '2026-11-05 09:00',
    '归档中',
    '已完成内网评审与结果公示，正在推进归档资料回收与装订。',
    JSON.stringify([
      { key: '采购大类', value: '直属单位服务-授权采购' },
      { key: '代理负责人', value: '张明' },
    ]),
    '2026-10-15T08:30:00.000Z',
    now
  );

  insertProject.run(uuid(), b1, 1, '省公司数据中心机房基础设施年度运维服务', 'XT-2026-0412-01', '145.00 万元', '竞争性谈判', '运检中心', '含7×24小时驻场', now);
  insertProject.run(uuid(), b1, 2, '电力调度通信传输网光缆巡检与应急抢修服务', 'XT-2026-0412-02', '98.50 万元', '竞争性谈判', '通信运检部', '覆盖16个地市节点', now);
  insertProject.run(uuid(), b1, 3, '信息系统安全等级保护测评与渗透测试服务', 'XT-2026-0412-03', '62.00 万元', '询价采购', '网络安全部', '需具备等保测评资质', now);

  const f1 = uuid();
  insertFile.run(
    f1, b1, 'bidding', '批次项目明细表', '2026Q4信通运维采购需求明细表.xlsx', '',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 48210,
    JSON.stringify({
      parsedAt: '2026-10-18T09:00:00.000Z',
      sheets: [{
        name: '需求明细',
        headers: ['标段编号', '项目名称', '需求部门', '最高限价(万元)', '采购方式'],
        totalRows: 3,
        rows: [
          { '标段编号': 'XT-2026-0412-01', '项目名称': '省公司数据中心机房基础设施年度运维服务', '需求部门': '运检中心', '最高限价(万元)': '145.00', '采购方式': '竞争性谈判' },
          { '标段编号': 'XT-2026-0412-02', '项目名称': '电力调度通信传输网光缆巡检与应急抢修服务', '需求部门': '通信运检部', '最高限价(万元)': '98.50', '采购方式': '竞争性谈判' },
          { '标段编号': 'XT-2026-0412-03', '项目名称': '信息系统安全等级保护测评与渗透测试服务', '需求部门': '网络安全部', '最高限价(万元)': '62.00', '采购方式': '询价采购' },
        ],
      }],
    }),
    '已核对最高限价与项目编号', '2026-10-18T09:00:00.000Z'
  );

  const f2 = uuid();
  insertFile.run(
    f2, b1, 'evaluation', '评分统计表', 'ECP2.0导出-2026Q4信通运维授标结果排序表.xlsx', '',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 39450,
    JSON.stringify({
      parsedAt: '2026-11-06T14:20:00.000Z',
      sheets: [{
        name: '授标排序',
        headers: ['标段编号', '第一成交候选人', '最终报价(万元)', '综合得分'],
        totalRows: 3,
        rows: [
          { '标段编号': 'XT-2026-0412-01', '第一成交候选人': '中电信息系统运维有限公司', '最终报价(万元)': '138.60', '综合得分': '93.4' },
          { '标段编号': 'XT-2026-0412-02', '第一成交候选人': '安徽电力通信工程有限责任公司', '最终报价(万元)': '94.20', '综合得分': '91.8' },
          { '标段编号': 'XT-2026-0412-03', '第一成交候选人': '国网信安检测技术研究院', '最终报价(万元)': '58.90', '综合得分': '95.1' },
        ],
      }],
    }),
    '内网评审结束后导出，已核对专家签字', '2026-11-06T14:20:00.000Z'
  );

  // Batch 1 SOP: all completed
  DEFAULT_SOP_STEPS.forEach((s, idx) => {
    insertSop.run(uuid(), b1, 0, s.item_type, s.stage_group, s.priority, s.title, s.description, 1, '2026-11-05 17:30', idx === 7 ? '全部专家签字核对完毕' : '', idx + 1);
  });
  DEFAULT_SOP_NOTICES.forEach((n, idx) => {
    insertSop.run(uuid(), b1, 0, n.item_type, n.stage_group, n.priority, n.title, n.description, 0, null, '', idx + 1);
  });

  // Batch 1 Archive: mostly present, 2 na, 4 required missing to demonstrate alerts
  DEFAULT_ARCHIVE_TEMPLATES.forEach((a, idx) => {
    let status: 'present' | 'missing' | 'na' = 'present';
    let notes = '';
    let linkedIds: string[] = [];

    if (/（如有）/.test(a.doc_name) && !a.is_required) {
      status = 'na';
      notes = '无';
    } else if (a.sheet_name === '归档文件' && ['14', '15', '18'].includes(a.seq)) {
      status = 'missing';
      notes = a.seq === '18' ? '待ECP2.0导出盖章' : '待领导小组签批返回';
    } else if (a.sheet_name === '重要保管文件' && a.seq === '18') {
      status = 'missing';
      notes = '缺技术组专家补签字';
    } else if (a.sheet_name === '重要保管文件' && a.seq === '5') {
      notes = '已归还原件并扫描';
    } else if (a.sheet_name === '重要保管文件' && a.seq === '22') {
      linkedIds = [f2];
      notes = '已关联开评标导出表';
    } else if (a.sheet_name === '归档文件' && a.seq === '5') {
      linkedIds = [f1];
    }

    insertArchive.run(
      uuid(), b1, 0, a.sheet_name, a.seq, a.doc_name, a.is_required,
      a.is_required ? 'color' : 'manual', a.media_type, a.paper_req, a.electronic_req,
      a.policy_basis, a.template_source, status, status === 'present' && linkedIds.length === 0 ? 1 : 0,
      JSON.stringify(linkedIds), notes, idx + 1
    );
  });

  // --- Batch 2: In Evaluation stage ---
  const b2 = uuid();
  insertBatch.run(
    b2,
    '2026年配网工程勘察设计服务集中采购批次',
    'GW-AH-2026-0418',
    '国网合肥供电公司',
    '2026-11-01 09:00',
    '2026-11-19 09:30',
    '开评标中',
    '内网评审进行中，预计明日出具正式评审报告。',
    JSON.stringify([
      { key: '采购大类', value: '工程服务-公开招标' },
      { key: '评标地点', value: '省公司评标基地 302 室' },
    ]),
    '2026-10-28T10:00:00.000Z',
    now
  );

  insertProject.run(uuid(), b2, 1, '包河区10kV配网改造工程勘察设计（标段一）', 'HF-PW-2026-01', '210.00 万元', '公开招标', '发展策划部', '', now);
  insertProject.run(uuid(), b2, 2, '蜀山区配网自动化升级勘察设计（标段二）', 'HF-PW-2026-02', '185.00 万元', '公开招标', '发展策划部', '', now);
  insertProject.run(uuid(), b2, 3, '肥西县农网巩固提升工程设计（标段三）', 'HF-PW-2026-03', '160.00 万元', '公开招标', '农电管理部', '', now);
  insertProject.run(uuid(), b2, 4, '巢湖市配网防灾抗冰专项设计（标段四）', 'HF-PW-2026-04', '128.00 万元', '公开招标', '运检部', '', now);

  insertFile.run(
    uuid(), b2, 'bidding', '批次项目明细表', '合肥配网勘察设计标段限价清单.xlsx', '',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 31200,
    JSON.stringify({
      parsedAt: '2026-11-01T11:00:00.000Z',
      sheets: [{
        name: '标段清单',
        headers: ['标段编号', '标段名称', '最高限价(万元)', '报名家数'],
        totalRows: 4,
        rows: [
          { '标段编号': 'HF-PW-2026-01', '标段名称': '包河区10kV配网改造工程勘察设计', '最高限价(万元)': '210.00', '报名家数': '7' },
          { '标段编号': 'HF-PW-2026-02', '标段名称': '蜀山区配网自动化升级勘察设计', '最高限价(万元)': '185.00', '报名家数': '6' },
          { '标段编号': 'HF-PW-2026-03', '标段名称': '肥西县农网巩固提升工程设计', '最高限价(万元)': '160.00', '报名家数': '8' },
          { '标段编号': 'HF-PW-2026-04', '标段名称': '巢湖市配网防灾抗冰专项设计', '最高限价(万元)': '128.00', '报名家数': '5' },
        ],
      }],
    }),
    '前期已核对', '2026-11-01T11:00:00.000Z'
  );

  // Batch 2 SOP: 6/10 completed
  DEFAULT_SOP_STEPS.forEach((s, idx) => {
    const done = idx < 6 ? 1 : 0;
    insertSop.run(uuid(), b2, 0, s.item_type, s.stage_group, s.priority, s.title, s.description, done, done ? '2026-11-19 11:20' : null, idx === 5 ? '已核查国网最新不良行为公告' : '', idx + 1);
  });
  DEFAULT_SOP_NOTICES.forEach((n, idx) => {
    insertSop.run(uuid(), b2, 0, n.item_type, n.stage_group, n.priority, n.title, n.description, 0, null, '', idx + 1);
  });

  // Batch 2 Archive: early stage items present
  DEFAULT_ARCHIVE_TEMPLATES.forEach((a, idx) => {
    const earlyDone = a.sheet_name === '归档文件' && ['1', '2', '4', '5', '8', '9'].includes(a.seq);
    insertArchive.run(
      uuid(), b2, 0, a.sheet_name, a.seq, a.doc_name, a.is_required,
      a.is_required ? 'color' : 'manual', a.media_type, a.paper_req, a.electronic_req,
      a.policy_basis, a.template_source, earlyDone ? 'present' : 'missing', earlyDone ? 1 : 0, '[]', '', idx + 1
    );
  });

  // --- Batch 3: Newly Announced (Upcoming bid opening in 5 days) ---
  const b3 = uuid();
  insertBatch.run(
    b3,
    '2026年12月直属单位后勤及物业专项服务采购批次',
    'GW-AH-2026-0425',
    '国网安徽电力综合服务中心',
    '2026-11-16 14:00',
    '2026-11-25 09:30',
    '已挂网',
    '公告已发布至 ECP2.0 及公共服务平台，下周三上午开标。',
    JSON.stringify([
      { key: '采购大类', value: '直属单位服务-授权采购' },
      { key: '公告截图', value: '已留存待签字' },
    ]),
    '2026-11-15T09:00:00.000Z',
    now
  );

  insertProject.run(uuid(), b3, 1, '综合办公楼2027年度物业安保及保洁外包服务', 'ZH-2026-0425-01', '118.00 万元', '竞争性谈判', '后勤保障部', '', now);
  insertProject.run(uuid(), b3, 2, '职工食堂食材集中配送与食品安全检测服务', 'ZH-2026-0425-02', '86.00 万元', '竞争性谈判', '膳食管理科', '', now);

  DEFAULT_SOP_STEPS.forEach((s, idx) => {
    const done = idx < 1 ? 1 : 0;
    insertSop.run(uuid(), b3, 0, s.item_type, s.stage_group, s.priority, s.title, s.description, done, done ? '2026-11-20 10:00' : null, '', idx + 1);
  });
  DEFAULT_SOP_NOTICES.forEach((n, idx) => {
    insertSop.run(uuid(), b3, 0, n.item_type, n.stage_group, n.priority, n.title, n.description, 0, null, '', idx + 1);
  });

  DEFAULT_ARCHIVE_TEMPLATES.forEach((a, idx) => {
    const earlyDone = a.sheet_name === '归档文件' && ['1', '2', '4', '5'].includes(a.seq);
    insertArchive.run(
      uuid(), b3, 0, a.sheet_name, a.seq, a.doc_name, a.is_required,
      a.is_required ? 'color' : 'manual', a.media_type, a.paper_req, a.electronic_req,
      a.policy_basis, a.template_source, earlyDone ? 'present' : 'missing', earlyDone ? 1 : 0, '[]', '', idx + 1
    );
  });
}
