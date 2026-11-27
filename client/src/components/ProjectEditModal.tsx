import React, { useState } from 'react';
import { Pencil, Save, X } from 'lucide-react';
import { api } from '../api/client';
import type { BatchProject } from '../types';

interface ProjectEditModalProps {
  project: BatchProject;
  onClose: () => void;
  onSaved: () => Promise<void> | void;
}

function yuanToWanInput(yuan: number | null | undefined): string {
  if (yuan === null || yuan === undefined || !Number.isFinite(Number(yuan))) return '';
  return String(Number((Number(yuan) / 10000).toFixed(6)));
}

export function ProjectEditModal({ project, onClose, onSaved }: ProjectEditModalProps) {
  const initialIsDiscount =
    /折扣|下浮|费率|百分/.test(project.quote_method || '') || Boolean(project.max_price_raw);

  const [name, setName] = useState(project.name || '');
  const [lotCode, setLotCode] = useState(project.lot_code || '');
  const [lotName, setLotName] = useState(project.lot_name || '');
  const [packageNo, setPackageNo] = useState(project.package_no || '');
  const [reqAppNo, setReqAppNo] = useState(project.req_app_no || '');
  const [quoteMethod, setQuoteMethod] = useState(
    project.quote_method || (initialIsDiscount ? '折扣率报价' : '总价报价')
  );
  const [budgetWan, setBudgetWan] = useState(yuanToWanInput(project.budget_yuan));
  const [maxPriceInput, setMaxPriceInput] = useState(
    project.max_price_raw || yuanToWanInput(project.max_price_yuan)
  );
  const [weightRatio, setWeightRatio] = useState(project.weight_ratio || '');
  const [duration, setDuration] = useState(project.duration || '');
  const [department, setDepartment] = useState(project.department || '');
  const [contractVersion, setContractVersion] = useState(project.contract_version || '');
  const [projectOverview, setProjectOverview] = useState(project.project_overview || '');
  const [qualificationReq, setQualificationReq] = useState(project.qualification_req || '');
  const [performanceReq, setPerformanceReq] = useState(project.performance_req || '');
  const [staffReq, setStaffReq] = useState(project.staff_req || '');
  const [saving, setSaving] = useState(false);

  const isDiscountMode = /折扣|下浮|费率|百分/.test(quoteMethod) || /%$/.test(maxPriceInput.trim());

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (!name.trim()) {
      alert('项目名称不能为空');
      return;
    }

    // Parse budget (万元 -> 元)
    let budgetYuan: number | null = null;
    const cleanBudget = budgetWan.trim().replace(/万元|万|,/g, '');
    if (cleanBudget !== '') {
      const n = Number(cleanBudget);
      if (!Number.isFinite(n) || n < 0) {
        alert('含税预算金额格式不正确，请输入数字（单位：万元）');
        return;
      }
      budgetYuan = Math.round(n * 10000 * 100) / 100;
    }

    // Parse max price (supports both 万元 number and 折扣率 %)
    let maxPriceYuan: number | null = null;
    let maxPriceRaw = '';
    const cleanMax = maxPriceInput.trim().replace(/,/g, '');
    if (cleanMax !== '') {
      if (/%$/.test(cleanMax) || /折扣|下浮|费率|百分/.test(quoteMethod)) {
        maxPriceRaw = cleanMax.endsWith('%') ? cleanMax : `${cleanMax}%`;
        maxPriceYuan = null;
      } else {
        const n = Number(cleanMax.replace(/万元|万/g, ''));
        if (!Number.isFinite(n) || n < 0) {
          alert('最高限价格式不正确，总价请输入数字（万元），折扣率请输入百分比（如 95%）');
          return;
        }
        maxPriceYuan = Math.round(n * 10000 * 100) / 100;
        maxPriceRaw = '';
      }
    }

    setSaving(true);
    try {
      await api.updateProject(project.id, {
        name: name.trim(),
        lot_code: lotCode.trim(),
        lot_name: lotName.trim(),
        package_no: packageNo.trim(),
        req_app_no: reqAppNo.trim(),
        quote_method: quoteMethod.trim(),
        budget_yuan: budgetYuan,
        max_price_yuan: maxPriceYuan,
        max_price_raw: maxPriceRaw,
        weight_ratio: weightRatio.trim(),
        duration: duration.trim(),
        department: department.trim(),
        contract_version: contractVersion.trim(),
        project_overview: projectOverview.trim(),
        qualification_req: qualificationReq.trim(),
        performance_req: performanceReq.trim(),
        staff_req: staffReq.trim(),
      });
      await onSaved();
      onClose();
    } catch (err) {
      alert((err as Error).message || '保存项目修改失败');
    } finally {
      setSaving(false);
    }
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4 backdrop-blur-[1px]">
      <div className="w-full max-w-2xl max-h-[90vh] flex flex-col rounded-xl border border-slate-200 bg-white shadow-xl overflow-hidden">
        <div className="flex items-center justify-between border-b border-slate-200 px-5 py-3.5">
          <div className="flex items-center gap-2">
            <Pencil className="h-4 w-4 text-[#3B5B7E]" />
            <h3 className="text-sm font-semibold text-slate-800">编辑标段 / 项目主数据</h3>
            <span className="rounded bg-[#EBF2F7] px-2 py-0.5 text-[10px] font-medium text-[#3B5B7E]">
              保存后自动同步开评标核验与定标导出
            </span>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="rounded-lg p-1 text-slate-400 hover:bg-slate-100"
          >
            <X className="h-4 w-4" />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="flex-1 overflow-y-auto p-5 space-y-4">
          {/* Core Price & Quote Fields (highlighted) */}
          <div className="rounded-lg border border-[#3B5B7E]/20 bg-[#EBF2F7]/30 p-3.5 space-y-3">
            <div className="text-xs font-semibold text-[#3B5B7E]">价格与限价核心字段（联动开评标超限价 / 报价过低核验与定标预算）</div>
            <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
              <div>
                <label className="block text-xs font-medium text-slate-600 mb-1">报价方式</label>
                <select
                  value={quoteMethod}
                  onChange={(e) => setQuoteMethod(e.target.value)}
                  className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                >
                  <option value="总价报价">总价报价</option>
                  <option value="折扣率报价">折扣率报价</option>
                  <option value="单价报价">单价报价</option>
                  <option value="费率报价">费率报价</option>
                </select>
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-600 mb-1">含税预算 / 概算金额（万元）</label>
                <input
                  value={budgetWan}
                  onChange={(e) => setBudgetWan(e.target.value)}
                  placeholder="例如：1098"
                  className="w-full rounded-lg border border-slate-200 bg-white px-2.5 py-1.5 text-xs font-mono focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-600 mb-1">
                  最高限价（{isDiscountMode ? '折扣率，如 100%' : '万元，如 1098'}）
                </label>
                <input
                  value={maxPriceInput}
                  onChange={(e) => setMaxPriceInput(e.target.value)}
                  placeholder={isDiscountMode ? '例如：100%' : '例如：1098'}
                  className="w-full rounded-lg border border-[#3B5B7E]/40 bg-white px-2.5 py-1.5 text-xs font-mono font-semibold text-[#3B5B7E] focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>
            </div>
          </div>

          {/* Basic Identification */}
          <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
            <div className="sm:col-span-2">
              <label className="block text-xs font-medium text-slate-600 mb-1">项目名称 *</label>
              <input
                required
                value={name}
                onChange={(e) => setName(e.target.value)}
                className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-600 mb-1">分标编号</label>
              <input
                value={lotCode}
                onChange={(e) => setLotCode(e.target.value)}
                className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-xs font-mono focus:border-[#3B5B7E] focus:outline-none"
              />
            </div>
            <div className="grid grid-cols-2 gap-2">
              <div>
                <label className="block text-xs font-medium text-slate-600 mb-1">分标名称</label>
                <input
                  value={lotName}
                  onChange={(e) => setLotName(e.target.value)}
                  className="w-full rounded-lg border border-slate-200 px-2.5 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-600 mb-1">包号</label>
                <input
                  value={packageNo}
                  onChange={(e) => setPackageNo(e.target.value)}
                  className="w-full rounded-lg border border-slate-200 px-2.5 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-600 mb-1">采购申请号</label>
              <input
                value={reqAppNo}
                onChange={(e) => setReqAppNo(e.target.value)}
                className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-xs font-mono focus:border-[#3B5B7E] focus:outline-none"
              />
            </div>
            <div className="grid grid-cols-2 gap-2">
              <div>
                <label className="block text-xs font-medium text-slate-600 mb-1">权重比例 (技:商:价)</label>
                <input
                  value={weightRatio}
                  onChange={(e) => setWeightRatio(e.target.value)}
                  placeholder="如 60:10:30"
                  className="w-full rounded-lg border border-slate-200 px-2.5 py-1.5 text-xs font-mono focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-600 mb-1">工期 / 进度安排</label>
                <input
                  value={duration}
                  onChange={(e) => setDuration(e.target.value)}
                  className="w-full rounded-lg border border-slate-200 px-2.5 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
                />
              </div>
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-600 mb-1">项目管理部门</label>
              <input
                value={department}
                onChange={(e) => setDepartment(e.target.value)}
                className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-600 mb-1">合同版本 / 名称</label>
              <input
                value={contractVersion}
                onChange={(e) => setContractVersion(e.target.value)}
                className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
              />
            </div>
          </div>

          {/* Requirements */}
          <div className="space-y-3">
            <div>
              <label className="block text-xs font-medium text-slate-600 mb-1">项目概况</label>
              <textarea
                rows={2}
                value={projectOverview}
                onChange={(e) => setProjectOverview(e.target.value)}
                className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-600 mb-1">资质条件 / 要求</label>
              <textarea
                rows={2}
                value={qualificationReq}
                onChange={(e) => setQualificationReq(e.target.value)}
                className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-600 mb-1">业绩要求</label>
              <textarea
                rows={2}
                value={performanceReq}
                onChange={(e) => setPerformanceReq(e.target.value)}
                className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-600 mb-1">人员要求</label>
              <textarea
                rows={2}
                value={staffReq}
                onChange={(e) => setStaffReq(e.target.value)}
                className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-xs focus:border-[#3B5B7E] focus:outline-none"
              />
            </div>
          </div>

          <div className="flex justify-end gap-2 border-t border-slate-100 pt-3">
            <button
              type="button"
              onClick={onClose}
              className="rounded-lg border border-slate-200 px-4 py-1.5 text-xs font-medium text-slate-600 hover:bg-slate-50"
            >
              取消
            </button>
            <button
              type="submit"
              disabled={saving}
              className="inline-flex items-center gap-1.5 rounded-lg bg-[#3B5B7E] px-4 py-1.5 text-xs font-medium text-white hover:bg-[#2E4866] disabled:opacity-50"
            >
              <Save className="h-3.5 w-3.5" />
              <span>{saving ? '保存中...' : '保存并全局同步'}</span>
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}
