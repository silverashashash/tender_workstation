import type { BatchProject } from '../types';

export type ProjectSortKey =
  | 'seq'
  | 'req_app_no'
  | 'lot_package'
  | 'lot_code'
  | 'lot_name'
  | 'package_no'
  | 'name'
  | 'project_unit'
  | 'budget_yuan'
  | 'max_price_yuan'
  | 'weight_ratio'
  | 'biz_eval_tpl'
  | 'duration'
  | 'contract_version'
  | 'department'
  | 'funding_source';

export type SortDirection = 'asc' | 'desc';

const pinyinCollator = new Intl.Collator('zh-Hans-CN-u-co-pinyin', {
  numeric: true,
  sensitivity: 'base',
});

/**
 * Extract the integer after "包" (e.g. "包1" -> 1, "包2" -> 2, "包10" -> 10).
 */
export function extractPackageNumber(pkg?: string | null): number {
  const s = String(pkg ?? '').trim();
  if (!s) return Number.POSITIVE_INFINITY;
  const m = s.match(/包\s*(\d+)/) || s.match(/(\d+)/);
  return m ? Number(m[1]) : Number.POSITIVE_INFINITY;
}

/**
 * Compare two numeric-like values (supports large integers like 采购申请号 and floats like prices).
 */
function compareNumeric(a: unknown, b: unknown): number {
  const sa = String(a ?? '').trim();
  const sb = String(b ?? '').trim();
  if (!sa && !sb) return 0;
  if (!sa) return 1;
  if (!sb) return -1;

  // Big integer strings (e.g. 19-digit 采购申请号)
  if (/^\d{12,}$/.test(sa) && /^\d{12,}$/.test(sb)) {
    const ba = BigInt(sa);
    const bb = BigInt(sb);
    return ba < bb ? -1 : ba > bb ? 1 : 0;
  }

  const na = Number(sa.replace(/[^\d.-]/g, ''));
  const nb = Number(sb.replace(/[^\d.-]/g, ''));
  const aValid = Number.isFinite(na) && sa !== '';
  const bValid = Number.isFinite(nb) && sb !== '';
  if (aValid && bValid) return na - nb;
  if (aValid) return -1;
  if (bValid) return 1;
  return pinyinCollator.compare(sa, sb);
}

/**
 * Primary sort by 分标名称 (Chinese Pinyin initials),
 * Secondary sort by 包号 number (number after "包", so 包2 < 包10).
 */
export function compareLotAndPackage(a: BatchProject, b: BatchProject): number {
  const lotA = (a.lot_name || a.lot_code || '').trim();
  const lotB = (b.lot_name || b.lot_code || '').trim();
  const lotCmp = pinyinCollator.compare(lotA, lotB);
  if (lotCmp !== 0) return lotCmp;

  const pkgNumA = extractPackageNumber(a.package_no);
  const pkgNumB = extractPackageNumber(b.package_no);
  if (pkgNumA !== pkgNumB) return pkgNumA - pkgNumB;

  return pinyinCollator.compare(a.lot_code || '', b.lot_code || '');
}

export function sortProjects(
  list: BatchProject[],
  sortKey: ProjectSortKey | null,
  sortDir: SortDirection
): BatchProject[] {
  if (!sortKey) return list;
  const factor = sortDir === 'asc' ? 1 : -1;

  return [...list].sort((a, b) => {
    let cmp = 0;
    switch (sortKey) {
      case 'seq':
        cmp = (a.seq || 0) - (b.seq || 0);
        break;
      case 'req_app_no':
        cmp = compareNumeric(a.req_app_no, b.req_app_no);
        break;
      case 'lot_package':
      case 'lot_name':
      case 'lot_code':
        cmp = compareLotAndPackage(a, b);
        break;
      case 'package_no': {
        const pa = extractPackageNumber(a.package_no);
        const pb = extractPackageNumber(b.package_no);
        cmp = pa !== pb ? pa - pb : compareLotAndPackage(a, b);
        break;
      }
      case 'budget_yuan':
        cmp = compareNumeric(a.budget_yuan, b.budget_yuan);
        break;
      case 'max_price_yuan':
        cmp = compareNumeric(
          a.max_price_yuan ?? a.budget_yuan ?? a.max_price_raw,
          b.max_price_yuan ?? b.budget_yuan ?? b.max_price_raw
        );
        break;
      case 'weight_ratio': {
        const wa = Number((a.weight_ratio || '').split(':')[0]);
        const wb = Number((b.weight_ratio || '').split(':')[0]);
        cmp = Number.isFinite(wa) && Number.isFinite(wb)
          ? wa - wb
          : pinyinCollator.compare(a.weight_ratio || '', b.weight_ratio || '');
        break;
      }
      default:
        cmp = pinyinCollator.compare(
          String(a[sortKey] ?? ''),
          String(b[sortKey] ?? '')
        );
    }
    return cmp * factor;
  });
}
