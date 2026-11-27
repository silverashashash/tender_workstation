import Dexie, { type Table } from 'dexie';
import type { ArchiveItem, Batch, BatchProject, SopItem } from '../types';

class WorkstationIDB extends Dexie {
  batches!: Table<Batch, string>;
  projects!: Table<BatchProject, string>;
  sopItems!: Table<SopItem, string>;
  archiveItems!: Table<ArchiveItem, string>;

  constructor() {
    super('TenderWorkstationFallbackDB');
    this.version(1).stores({
      batches: 'id, created_at',
      projects: 'id, batch_id',
      sopItems: 'id, batch_id, item_type',
      archiveItems: 'id, batch_id, sheet_name',
    });
  }
}

export const idb = new WorkstationIDB();

export async function ensureIdbSeeded() {
  const count = await idb.batches.count();
  if (count > 0) return;

  const now = new Date().toISOString();
  const b1 = 'idb-batch-1';
  await idb.batches.add({
    id: b1,
    name: '2026年第四季度信息通信运维服务授权采购批次（离线兜底）',
    code: 'GW-AH-2026-0412',
    project_unit: '国网安徽省电力有限公司信通分公司',
    announcement_time: '2026-10-18 10:00',
    bid_opening_time: '2026-11-05 09:00',
    status: '归档中',
    notes: '当前为浏览器 IndexedDB 离线持久化模式，刷新数据不丢失。',
    extra_fields: [{ key: '采购大类', value: '直属单位服务-授权采购' }],
    created_at: now,
    updated_at: now,
    stats: {
      projectCount: 2,
      biddingFileCount: 0,
      evalFileCount: 0,
      awardOutputCount: 0,
      sopTotal: 2,
      sopDone: 1,
      archiveTotal: 2,
      archivePresent: 1,
      archiveNa: 0,
      archiveReqTotal: 2,
      archiveReqPresent: 1,
      archiveReqMissing: 1,
    },
  });

  await idb.projects.bulkAdd([
    {
      id: 'idb-p1',
      batch_id: b1,
      seq: 1,
      name: '省公司数据中心机房基础设施年度运维服务',
      code: 'XT-2026-0412-01',
      budget: '145.00 万元',
      procurement_method: '竞争性谈判',
      department: '运检中心',
      notes: '',
      created_at: now,
    },
    {
      id: 'idb-p2',
      batch_id: b1,
      seq: 2,
      name: '电力调度通信传输网光缆巡检服务',
      code: 'XT-2026-0412-02',
      budget: '98.50 万元',
      procurement_method: '竞争性谈判',
      department: '通信运检部',
      notes: '',
      created_at: now,
    },
  ]);

  await idb.sopItems.bulkAdd([
    {
      id: 'idb-s1',
      batch_id: b1,
      is_template: 0,
      item_type: 'step',
      stage_group: '开标前准备',
      priority: 'normal',
      title: '核验 CA 证书与内网 ECP2.0 账号权限',
      description: '提前确认 CA 锁驱动正常与批次权限。',
      is_completed: 1,
      completed_at: '2026-11-05 08:30',
      notes: '',
      sort_order: 1,
    },
    {
      id: 'idb-s2',
      batch_id: b1,
      is_template: 0,
      item_type: 'notice',
      stage_group: '评审过程',
      priority: 'critical',
      title: '离场前核对全体评委签字页完整',
      description: '避免初评小结或评审报告遗漏签字。',
      is_completed: 0,
      completed_at: null,
      notes: '',
      sort_order: 2,
    },
  ]);

  await idb.archiveItems.bulkAdd([
    {
      id: 'idb-a1',
      batch_id: b1,
      sheet_name: '归档文件',
      seq: '1',
      doc_name: '采购代理服务合同',
      is_required: true,
      required_source: 'color',
      media_type: '纸质+电子',
      paper_req: '原件复印',
      electronic_req: '原件扫描',
      policy_basis: '国网（物资/4）243-2024',
      template_source: '',
      status: 'present',
      is_manual_marked: true,
      linked_file_ids: [],
      linked_files: [],
      notes: '已归档',
      sort_order: 1,
    },
    {
      id: 'idb-a2',
      batch_id: b1,
      sheet_name: '归档文件',
      seq: '13',
      doc_name: '评审报告',
      is_required: true,
      required_source: 'color',
      media_type: '纸质+电子',
      paper_req: '正式评审报告(盖封面章)，评委会签字页原件',
      electronic_req: '带章封面及签字页扫描PDF',
      policy_basis: '国网（物资/4）243-2024',
      template_source: '',
      status: 'missing',
      is_manual_marked: false,
      linked_file_ids: [],
      linked_files: [],
      notes: '待补签字',
      sort_order: 2,
    },
  ]);
}
