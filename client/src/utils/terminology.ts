export type ProcurementType = '公开招标' | '非公开招标';
export type UnitNature = '主业' | '产业';

export interface TerminologySet {
  procurementType: ProcurementType;
  unitNature: UnitNature;
  initiator: string;      // 招标人 vs 采购人
  bidder: string;         // 投标人 vs 响应供应商
  tenderDoc: string;      // 招标文件 vs 采购文件
  responseDoc: string;    // 投标文件 vs 响应文件
  deadlineLabel: string;  // 投标截止时间 vs 应答截止时间
  fetchDocLabel: string;  // 获取招标文件时间 vs 获取采购文件时间
  platformShort: string;  // ECP 2.0 vs ETP
  platformFull: string;   // ECP 2.0 vs ETP（电工交易专区）
}

export function getTerminology(
  procurementType?: string,
  unitNature?: string
): TerminologySet {
  const isPublic = procurementType === '公开招标';
  const isMain = unitNature === '主业';

  return {
    procurementType: isPublic ? '公开招标' : '非公开招标',
    unitNature: isMain ? '主业' : '产业',
    initiator: isPublic ? '招标人' : '采购人',
    bidder: isPublic ? '投标人' : '响应供应商',
    tenderDoc: isPublic ? '招标文件' : '采购文件',
    responseDoc: isPublic ? '投标文件' : '响应文件',
    deadlineLabel: isPublic ? '投标截止（开标）时间' : '应答截止（开标）时间',
    fetchDocLabel: isPublic ? '获取招标文件时间' : '获取采购文件时间',
    platformShort: isMain ? 'ECP 2.0' : 'ETP',
    platformFull: isMain ? '主业 · ECP 2.0' : '产业 · ETP（电工交易专区）',
  };
}

export function formatMoney(yuan: number | null | undefined, unit: 'yuan' | 'wan', fallback = '—'): string {
  if (yuan === null || yuan === undefined || !Number.isFinite(yuan)) {
    return fallback;
  }
  if (unit === 'wan') {
    const wan = yuan / 10000;
    return `${wan.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 4 })} 万元`;
  }
  return `${yuan.toLocaleString('zh-CN', { maximumFractionDigits: 2 })} 元`;
}
