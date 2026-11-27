import type {
  ArchiveOverview,
  AwardOutput,
  AwardOverview,
  Batch,
  BatchProject,
  BiddingWorkspace,
  ParsedArchivePreviewRow,
  SopItem,
  UploadedFile,
} from '../types';
import { ensureIdbSeeded, idb } from './idbFallback';

export type StorageMode = 'sqlite' | 'indexeddb';

let currentStorageMode: StorageMode = 'sqlite';

export function getStorageMode(): StorageMode {
  return currentStorageMode;
}

async function request<T>(url: string, options?: RequestInit): Promise<T> {
  const res = await fetch(url, {
    credentials: 'include',
    ...options,
    headers: {
      ...(options?.body instanceof FormData ? {} : { 'Content-Type': 'application/json' }),
      ...(options?.headers || {}),
    },
  });
  if (res.status === 401) {
    window.dispatchEvent(new CustomEvent('tw:unauthorized'));
    const data = await res.json().catch(() => ({}));
    throw new Error(data.error || '请先登录');
  }
  if (!res.ok) {
    const data = await res.json().catch(() => ({}));
    throw new Error(data.error || `请求失败 (${res.status})`);
  }
  return res.json() as Promise<T>;
}

export const api = {
  async checkHealth(): Promise<{ storage: StorageMode; authEnabled: boolean; authenticated: boolean }> {
    try {
      await request<{ ok: boolean }>('/api/health');
      const auth = await request<{ authEnabled: boolean; authenticated: boolean }>('/api/auth/status');
      currentStorageMode = 'sqlite';
      return { storage: 'sqlite', ...auth };
    } catch {
      currentStorageMode = 'indexeddb';
      await ensureIdbSeeded();
      return { storage: 'indexeddb', authEnabled: false, authenticated: true };
    }
  },

  async login(password: string) {
    return request<{ ok: boolean; authenticated: boolean }>('/api/auth/login', {
      method: 'POST',
      body: JSON.stringify({ password }),
    });
  },

  async logout() {
    return request<{ ok: boolean }>('/api/auth/logout', { method: 'POST' });
  },

  // --- Batches ---
  async listBatches(): Promise<Batch[]> {
    if (currentStorageMode === 'indexeddb') {
      return idb.batches.toArray();
    }
    return request<Batch[]>('/api/batches');
  },

  async getBatch(id: string): Promise<Batch> {
    if (currentStorageMode === 'indexeddb') {
      const b = await idb.batches.get(id);
      if (!b) throw new Error('批次不存在');
      const projects = await idb.projects.where('batch_id').equals(id).toArray();
      return { ...b, projects };
    }
    return request<Batch>(`/api/batches/${id}`);
  },

  async createBatch(payload: Partial<Batch>): Promise<Batch> {
    if (currentStorageMode === 'indexeddb') {
      const now = new Date().toISOString();
      const b: Batch = {
        id: crypto.randomUUID(),
        name: payload.name || '未命名批次',
        code: payload.code || '',
        project_unit: payload.project_unit || '',
        procurement_type: payload.procurement_type || '非公开招标',
        unit_nature: payload.unit_nature || '产业',
        announcement_time: payload.announcement_time || '',
        bid_opening_time: payload.bid_opening_time || '',
        status: payload.status || '筹备中',
        notes: payload.notes || '',
        extra_fields: payload.extra_fields || [],
        created_at: now,
        updated_at: now,
        stats: {
          projectCount: 0, biddingFileCount: 0, evalFileCount: 0, awardOutputCount: 0,
          sopTotal: 0, sopDone: 0, archiveTotal: 0, archivePresent: 0, archiveNa: 0,
          archiveReqTotal: 0, archiveReqPresent: 0, archiveReqMissing: 0,
        },
      };
      await idb.batches.add(b);
      return b;
    }
    return request<Batch>('/api/batches', {
      method: 'POST',
      body: JSON.stringify(payload),
    });
  },

  async updateBatch(id: string, payload: Partial<Batch>): Promise<Batch> {
    if (currentStorageMode === 'indexeddb') {
      await idb.batches.update(id, { ...payload, updated_at: new Date().toISOString() });
      const b = await idb.batches.get(id);
      if (!b) throw new Error('批次不存在');
      return b;
    }
    return request<Batch>(`/api/batches/${id}`, {
      method: 'PUT',
      body: JSON.stringify(payload),
    });
  },

  async listTrashBatches(): Promise<Batch[]> {
    if (currentStorageMode === 'indexeddb') return [];
    return request<Batch[]>('/api/batches/trash');
  },

  async trashBatch(id: string): Promise<{ ok: boolean; deleted_at: string }> {
    return request<{ ok: boolean; deleted_at: string }>(`/api/batches/${id}/trash`, {
      method: 'POST',
    });
  },

  async restoreBatch(id: string): Promise<{ ok: boolean }> {
    return request<{ ok: boolean }>(`/api/batches/${id}/restore`, {
      method: 'POST',
    });
  },

  async deleteBatch(id: string, payload?: { confirmText?: string; password?: string }): Promise<void> {
    if (currentStorageMode === 'indexeddb') {
      await idb.batches.delete(id);
      return;
    }
    await request(`/api/batches/${id}`, {
      method: 'DELETE',
      body: JSON.stringify(payload || {}),
    });
  },

  // --- Projects ---
  async createProject(batchId: string, payload: Partial<BatchProject>): Promise<BatchProject> {
    if (currentStorageMode === 'indexeddb') {
      const p: BatchProject = {
        id: crypto.randomUUID(),
        batch_id: batchId,
        seq: payload.seq || 1,
        name: payload.name || '',
        code: payload.code || '',
        budget: payload.budget || '',
        procurement_method: payload.procurement_method || '竞争性谈判',
        department: payload.department || '',
        notes: payload.notes || '',
        created_at: new Date().toISOString(),
      };
      await idb.projects.add(p);
      return p;
    }
    return request<BatchProject>(`/api/batches/${batchId}/projects`, {
      method: 'POST',
      body: JSON.stringify(payload),
    });
  },

  async updateProject(projectId: string, payload: Partial<BatchProject>): Promise<BatchProject> {
    if (currentStorageMode === 'indexeddb') {
      await idb.projects.update(projectId, payload);
      return (await idb.projects.get(projectId))!;
    }
    return request<BatchProject>(`/api/batches/projects/${projectId}`, {
      method: 'PUT',
      body: JSON.stringify(payload),
    });
  },

  async deleteProject(projectId: string): Promise<void> {
    if (currentStorageMode === 'indexeddb') {
      await idb.projects.delete(projectId);
      return;
    }
    await request(`/api/batches/projects/${projectId}`, { method: 'DELETE' });
  },

  // --- Files ---
  async listFiles(batchId: string, module?: string): Promise<UploadedFile[]> {
    if (currentStorageMode === 'indexeddb') return [];
    const q = module ? `?module=${encodeURIComponent(module)}` : '';
    return request<UploadedFile[]>(`/api/files/batch/${batchId}${q}`);
  },

  async uploadFile(batchId: string, file: File, module: string, category: string, notes = ''): Promise<UploadedFile> {
    const fd = new FormData();
    fd.append('module', module);
    fd.append('category', category);
    fd.append('notes', notes);
    fd.append('file', file);
    return request<UploadedFile>(`/api/files/batch/${batchId}/upload`, {
      method: 'POST',
      body: fd,
    });
  },

  async deleteFile(fileId: string): Promise<void> {
    await request(`/api/files/${fileId}`, { method: 'DELETE' });
  },

  // --- Bidding v2 Workspace & 5-Slot SOP Upload ---
  async getBiddingWorkspace(batchId: string): Promise<BiddingWorkspace> {
    return request<BiddingWorkspace>(`/api/bidding/batch/${batchId}/workspace`);
  },

  async uploadBiddingSlot(batchId: string, slotKey: string, file: File) {
    const fd = new FormData();
    fd.append('file', file);
    return request<{
      ok: boolean;
      fileId: string;
      report: { slotLabel: string; summary: string; warnings: string[] };
    }>(`/api/bidding/batch/${batchId}/upload-slot/${slotKey}`, {
      method: 'POST',
      body: fd,
    });
  },

  async clearBiddingSlot(batchId: string, slotKey: string): Promise<{ ok: boolean; slotLabel: string }> {
    return request<{ ok: boolean; slotLabel: string }>(`/api/bidding/batch/${batchId}/slot/${slotKey}`, {
      method: 'DELETE',
    });
  },

  async fixProjectWeight(projectId: string, weightRatio: string) {
    return request<{ ok: boolean }>(`/api/bidding/projects/${projectId}/fix-weight`, {
      method: 'POST',
      body: JSON.stringify({ weight_ratio: weightRatio }),
    });
  },

  async createChangeNotice(
    batchId: string,
    payload: {
      title: string;
      notice_date: string;
      change_types: string[];
      summary: string;
      project_ids: string[];
      file?: File | null;
    }
  ) {
    const fd = new FormData();
    fd.append('title', payload.title);
    fd.append('notice_date', payload.notice_date);
    fd.append('change_types', JSON.stringify(payload.change_types));
    fd.append('summary', payload.summary);
    fd.append('project_ids', JSON.stringify(payload.project_ids));
    if (payload.file) fd.append('file', payload.file);
    return request<{ ok: boolean; notice: import('../types').BatchChangeNotice }>(
      `/api/bidding/batch/${batchId}/change-notices`,
      { method: 'POST', body: fd }
    );
  },

  async updateChangeNotice(
    noticeId: string,
    payload: {
      title: string;
      notice_date: string;
      change_types: string[];
      summary: string;
      project_ids: string[];
      file?: File | null;
      remove_file?: boolean;
    }
  ) {
    const fd = new FormData();
    fd.append('title', payload.title);
    fd.append('notice_date', payload.notice_date);
    fd.append('change_types', JSON.stringify(payload.change_types));
    fd.append('summary', payload.summary);
    fd.append('project_ids', JSON.stringify(payload.project_ids));
    if (payload.file) fd.append('file', payload.file);
    if (payload.remove_file) fd.append('remove_file', 'true');
    return request<{ ok: boolean; notice: import('../types').BatchChangeNotice }>(
      `/api/bidding/change-notices/${noticeId}`,
      { method: 'PUT', body: fd }
    );
  },

  async deleteChangeNotice(noticeId: string) {
    return request<{ ok: boolean }>(`/api/bidding/change-notices/${noticeId}`, {
      method: 'DELETE',
    });
  },

  async importProjectsFromExcel(batchId: string, file: File) {
    const fd = new FormData();
    fd.append('file', file);
    return request<{
      ok: boolean;
      report: { slotLabel: string; summary: string; warnings: string[] };
    }>(`/api/bidding/batch/${batchId}/import-projects`, {
      method: 'POST',
      body: fd,
    });
  },

  async batchDeleteProjects(ids: string[]) {
    if (currentStorageMode === 'indexeddb') {
      await idb.projects.bulkDelete(ids);
      return { ok: true, deleted: ids.length };
    }
    return request<{ ok: boolean; deleted: number }>('/api/bidding/projects/batch-delete', {
      method: 'POST',
      body: JSON.stringify({ ids }),
    });
  },

  // --- Evaluation v2 Workspace & SOP ---
  async getEvaluationWorkspace(batchId: string) {
    return request<import('../types').EvaluationWorkspace>(`/api/evaluation/batch/${batchId}/workspace`);
  },

  async uploadEvaluationSlot(batchId: string, slotKey: string, file: File, force = false): Promise<
    | { ok: true; fileId: string; report: { slotLabel: string; summary: string; warnings: string[] } }
    | { mismatch: true; currentBatchCode: string; fileBatchCode: string; error: string }
  > {
    const fd = new FormData();
    fd.append('file', file);
    if (force) fd.append('force', 'true');
    const res = await fetch(`/api/evaluation/batch/${batchId}/upload-slot/${slotKey}`, {
      method: 'POST',
      credentials: 'include',
      body: fd,
    });
    const data = await res.json().catch(() => ({}));
    if (res.status === 409 && data.mismatch) {
      return data;
    }
    if (!res.ok) {
      throw new Error(data.error || `上传失败 (${res.status})`);
    }
    return data;
  },

  async clearEvaluationSlot(batchId: string, slotKey: string): Promise<{ ok: boolean; slotLabel: string }> {
    return request<{ ok: boolean; slotLabel: string }>(`/api/evaluation/batch/${batchId}/slot/${slotKey}`, {
      method: 'DELETE',
    });
  },

  async saveEvaluationAuditConfig(
    batchId: string,
    payload: { tax_category: '服务' | '物资'; proof_rates: number[] }
  ): Promise<{ ok: boolean; tax_category: '服务' | '物资'; proof_rates: number[] }> {
    return request(`/api/evaluation/batch/${batchId}/audit-config`, {
      method: 'PUT',
      body: JSON.stringify(payload),
    });
  },

  async getScoreCheckState(batchId: string) {
    return request<import('../types').ScoreCheckState>(`/api/evaluation/batch/${batchId}/score-check`);
  },

  async uploadAndRunTotalScoreCheck(
    batchId: string,
    files: File[],
    params: { expertNum?: number | null; maxScoreDiff?: number; password?: string }
  ) {
    const fd = new FormData();
    files.forEach((f) => fd.append('files', f));
    if (params.expertNum !== undefined && params.expertNum !== null && params.expertNum > 0) {
      fd.append('expertNum', String(params.expertNum));
    }
    if (params.maxScoreDiff !== undefined) {
      fd.append('maxScoreDiff', String(params.maxScoreDiff));
    }
    if (params.password) {
      fd.append('password', params.password);
    }
    return request<NonNullable<import('../types').ScoreCheckState['total']>>(
      `/api/evaluation/batch/${batchId}/score-check/total`,
      { method: 'POST', body: fd }
    );
  },

  async rerunTotalScoreCheck(
    batchId: string,
    params: { expertNum?: number | null; maxScoreDiff?: number; password?: string }
  ) {
    return request<NonNullable<import('../types').ScoreCheckState['total']>>(
      `/api/evaluation/batch/${batchId}/score-check/total/rerun`,
      { method: 'POST', body: JSON.stringify(params) }
    );
  },

  async clearTotalScoreCheck(batchId: string) {
    return request<{ ok: boolean }>(`/api/evaluation/batch/${batchId}/score-check/total`, {
      method: 'DELETE',
    });
  },

  async uploadObjectiveScoreFiles(batchId: string, files: File[]) {
    const fd = new FormData();
    files.forEach((f) => fd.append('files', f));
    return request<NonNullable<import('../types').ScoreCheckState['objective']>>(
      `/api/evaluation/batch/${batchId}/score-check/objective/upload`,
      { method: 'POST', body: fd }
    );
  },

  async runObjectiveScoreCheck(
    batchId: string,
    payload: { selections: import('../types').ObjectiveSelection[]; sectionOverride?: string }
  ) {
    return request<NonNullable<import('../types').ScoreCheckState['objective']>>(
      `/api/evaluation/batch/${batchId}/score-check/objective/run`,
      { method: 'POST', body: JSON.stringify(payload) }
    );
  },

  async clearObjectiveScoreCheck(batchId: string) {
    return request<{ ok: boolean }>(`/api/evaluation/batch/${batchId}/score-check/objective`, {
      method: 'DELETE',
    });
  },

  // --- 专家打分预校验（线下打分模板预校验） ---
  async getScorePrecheckState(batchId: string) {
    return request<import('../types').ExpertPrecheckState>(`/api/evaluation/batch/${batchId}/score-precheck`);
  },

  async uploadAndRunScorePrecheck(
    batchId: string,
    payload: {
      experts: File[];
      template?: File | null;
      maxScoreDiff?: number;
      excludeRejected?: boolean;
      append?: boolean;
    }
  ) {
    const fd = new FormData();
    payload.experts.forEach((f) => fd.append('experts', f));
    if (payload.template) fd.append('template', payload.template);
    if (payload.maxScoreDiff !== undefined) {
      fd.append('maxScoreDiff', String(payload.maxScoreDiff));
    }
    if (payload.excludeRejected !== undefined) {
      fd.append('excludeRejected', String(payload.excludeRejected));
    }
    if (payload.append) {
      fd.append('append', 'true');
    }
    return request<import('../types').ExpertPrecheckState>(
      `/api/evaluation/batch/${batchId}/score-precheck/upload`,
      { method: 'POST', body: fd }
    );
  },

  async removeScorePrecheckFiles(
    batchId: string,
    payload: {
      fileNames: string[];
      maxScoreDiff?: number;
      excludeRejected?: boolean;
      objectiveSelections?: import('../types').PrecheckObjectiveSelection[];
    }
  ) {
    return request<import('../types').ExpertPrecheckState>(
      `/api/evaluation/batch/${batchId}/score-precheck/remove-files`,
      { method: 'POST', body: JSON.stringify(payload) }
    );
  },

  async rerunScorePrecheck(
    batchId: string,
    payload: {
      maxScoreDiff?: number;
      excludeRejected?: boolean;
      objectiveSelections: import('../types').PrecheckObjectiveSelection[];
    }
  ) {
    return request<import('../types').ExpertPrecheckState>(
      `/api/evaluation/batch/${batchId}/score-precheck/rerun`,
      { method: 'POST', body: JSON.stringify(payload) }
    );
  },

  async clearScorePrecheck(batchId: string) {
    return request<{ ok: boolean }>(`/api/evaluation/batch/${batchId}/score-precheck`, {
      method: 'DELETE',
    });
  },

  async getInitialReportState(batchId: string) {
    return request<import('../types').InitialReportState>(`/api/evaluation/batch/${batchId}/initial-report`);
  },

  async uploadInitialReportTemplate(
    batchId: string,
    file: File,
    params: { department?: string; clarification?: string; simplifyNames?: boolean; companySuffixes?: string }
  ) {
    const fd = new FormData();
    fd.append('file', file);
    if (params.department !== undefined) fd.append('department', params.department);
    if (params.clarification !== undefined) fd.append('clarification', params.clarification);
    if (params.simplifyNames !== undefined) fd.append('simplifyNames', String(params.simplifyNames));
    if (params.companySuffixes !== undefined) fd.append('companySuffixes', params.companySuffixes);
    return request<import('../types').InitialReportState>(
      `/api/evaluation/batch/${batchId}/initial-report/upload`,
      { method: 'POST', body: fd }
    );
  },

  async regenerateInitialReport(
    batchId: string,
    params: { department?: string; clarification?: string; simplifyNames?: boolean; companySuffixes?: string }
  ) {
    return request<import('../types').InitialReportState>(
      `/api/evaluation/batch/${batchId}/initial-report/regenerate`,
      { method: 'POST', body: JSON.stringify(params) }
    );
  },

  async clearInitialReport(batchId: string) {
    return request<{ ok: boolean }>(`/api/evaluation/batch/${batchId}/initial-report`, {
      method: 'DELETE',
    });
  },

  async getWarningFormState(batchId: string) {
    return request<import('../types').WarningFormState>(`/api/evaluation/batch/${batchId}/warning-form`);
  },

  async runWarningForm(
    batchId: string,
    payload: { risk?: File; groups?: File; bids?: File; date?: string }
  ) {
    const fd = new FormData();
    if (payload.risk) fd.append('risk', payload.risk);
    if (payload.groups) fd.append('groups', payload.groups);
    if (payload.bids) fd.append('bids', payload.bids);
    if (payload.date !== undefined) fd.append('date', payload.date);
    return request<import('../types').WarningFormState>(
      `/api/evaluation/batch/${batchId}/warning-form/run`,
      { method: 'POST', body: fd }
    );
  },

  async rerunWarningForm(batchId: string, date: string) {
    return request<import('../types').WarningFormState>(
      `/api/evaluation/batch/${batchId}/warning-form/rerun`,
      { method: 'POST', body: JSON.stringify({ date }) }
    );
  },

  async clearWarningForm(batchId: string) {
    return request<{ ok: boolean }>(`/api/evaluation/batch/${batchId}/warning-form`, {
      method: 'DELETE',
    });
  },

  async getInitialMinutesState(batchId: string) {
    return request<import('../types').InitialMinutesState>(`/api/evaluation/batch/${batchId}/initial-minutes`);
  },

  async uploadInitialMinutesExcel(
    batchId: string,
    file: File,
    params: import('../types').InitialMinutesConfig
  ) {
    const fd = new FormData();
    fd.append('file', file);
    fd.append('procurementType', params.procurementType);
    fd.append('batchName', params.batchName);
    fd.append('meetingTime', params.meetingTime);
    fd.append('groupName', params.groupName);
    fd.append('clarificationNote', params.clarificationNote);
    if (params.reviewTypeFilter) fd.append('reviewTypeFilter', params.reviewTypeFilter);
    return request<import('../types').InitialMinutesState>(
      `/api/evaluation/batch/${batchId}/initial-minutes/upload`,
      { method: 'POST', body: fd }
    );
  },

  async regenerateInitialMinutes(
    batchId: string,
    params: import('../types').InitialMinutesConfig
  ) {
    return request<import('../types').InitialMinutesState>(
      `/api/evaluation/batch/${batchId}/initial-minutes/regenerate`,
      { method: 'POST', body: JSON.stringify(params) }
    );
  },

  async clearInitialMinutes(batchId: string) {
    return request<{ ok: boolean }>(`/api/evaluation/batch/${batchId}/initial-minutes`, {
      method: 'DELETE',
    });
  },

  // --- 价格分计算 ---
  async getPriceScoreState(batchId: string) {
    return request<import('../types').PriceScoreState>(`/api/evaluation/batch/${batchId}/price-score`);
  },
  async runPriceScore(batchId: string, payload: { quotes?: File[]; rules?: File; round?: string }) {
    const fd = new FormData();
    (payload.quotes || []).forEach((f) => fd.append('quotes', f));
    if (payload.rules) fd.append('rules', payload.rules);
    if (payload.round !== undefined) fd.append('round', payload.round);
    return request<import('../types').PriceScoreState>(`/api/evaluation/batch/${batchId}/price-score/run`, {
      method: 'POST',
      body: fd,
    });
  },
  async clearPriceScore(batchId: string) {
    return request<{ ok: boolean }>(`/api/evaluation/batch/${batchId}/price-score`, { method: 'DELETE' });
  },

  // --- 履约分生成 ---
  async getPerformanceState(batchId: string) {
    return request<import('../types').PerformanceState>(`/api/evaluation/batch/${batchId}/performance-score`);
  },
  async runPerformanceScore(batchId: string, payload: { bids?: File; scores?: File }) {
    const fd = new FormData();
    if (payload.bids) fd.append('bids', payload.bids);
    if (payload.scores) fd.append('scores', payload.scores);
    return request<import('../types').PerformanceState>(`/api/evaluation/batch/${batchId}/performance-score/run`, {
      method: 'POST',
      body: fd,
    });
  },
  async clearPerformanceScore(batchId: string) {
    return request<{ ok: boolean }>(`/api/evaluation/batch/${batchId}/performance-score`, { method: 'DELETE' });
  },

  // --- 经济排序表核验 ---
  async getEconomicRankState(batchId: string) {
    return request<import('../types').EconomicRankState>(`/api/evaluation/batch/${batchId}/economic-rank`);
  },
  async runEconomicRankCheck(batchId: string, files: File[]) {
    const fd = new FormData();
    files.forEach((f) => fd.append('files', f));
    return request<import('../types').EconomicRankState>(`/api/evaluation/batch/${batchId}/economic-rank/run`, {
      method: 'POST',
      body: fd,
    });
  },
  async clearEconomicRankCheck(batchId: string) {
    return request<{ ok: boolean }>(`/api/evaluation/batch/${batchId}/economic-rank`, { method: 'DELETE' });
  },

  // --- 授标分配 ---
  async getAwardAllocState(batchId: string) {
    return request<import('../types').AwardAllocState>(`/api/evaluation/batch/${batchId}/award-alloc`);
  },
  async runAwardAlloc(batchId: string, payload: { files?: File[]; groups?: File; mode: import('../types').AwardMode }) {
    const fd = new FormData();
    (payload.files || []).forEach((f) => fd.append('files', f));
    if (payload.groups) fd.append('groups', payload.groups);
    fd.append('mode', payload.mode);
    return request<import('../types').AwardAllocState>(`/api/evaluation/batch/${batchId}/award-alloc/run`, {
      method: 'POST',
      body: fd,
    });
  },
  async clearAwardAlloc(batchId: string) {
    return request<{ ok: boolean }>(`/api/evaluation/batch/${batchId}/award-alloc`, { method: 'DELETE' });
  },

  // --- 智能分包 ---
  async getSplitPackState(batchId: string) {
    return request<import('../types').SplitPackState>(`/api/evaluation/batch/${batchId}/split-pack`);
  },
  async runSplitPack(
    batchId: string,
    payload: { materials?: File; rules?: File; mode: import('../types').SplitPackMode; weightStep: string; lengthStep: string }
  ) {
    const fd = new FormData();
    if (payload.materials) fd.append('materials', payload.materials);
    if (payload.rules) fd.append('rules', payload.rules);
    fd.append('mode', payload.mode);
    fd.append('weightStep', payload.weightStep);
    fd.append('lengthStep', payload.lengthStep);
    return request<import('../types').SplitPackState>(`/api/evaluation/batch/${batchId}/split-pack/run`, {
      method: 'POST',
      body: fd,
    });
  },
  async clearSplitPack(batchId: string) {
    return request<{ ok: boolean }>(`/api/evaluation/batch/${batchId}/split-pack`, { method: 'DELETE' });
  },

  async listSop(batchId: string): Promise<SopItem[]> {
    if (currentStorageMode === 'indexeddb') {
      return idb.sopItems.where('batch_id').equals(batchId).toArray();
    }
    return request<SopItem[]>(`/api/sop/batch/${batchId}`);
  },

  async createSop(batchId: string, payload: Partial<SopItem> & { sync_to_template?: boolean }): Promise<SopItem> {
    return request<SopItem>(`/api/sop/batch/${batchId}`, {
      method: 'POST',
      body: JSON.stringify(payload),
    });
  },

  async updateSop(id: string, payload: Partial<SopItem>): Promise<SopItem> {
    return request<SopItem>(`/api/sop/${id}`, {
      method: 'PUT',
      body: JSON.stringify(payload),
    });
  },

  async deleteSop(id: string): Promise<void> {
    await request(`/api/sop/${id}`, { method: 'DELETE' });
  },

  async saveSopAsTemplate(batchId: string): Promise<{ ok: boolean; count: number }> {
    return request(`/api/sop/batch/${batchId}/save-as-template`, { method: 'POST' });
  },

  // --- Global System Operation SOP Guides (Text / Image / MD / PDF) ---
  async listSopGuides(): Promise<import('../types').SopGuide[]> {
    if (currentStorageMode === 'indexeddb') return [];
    return request<import('../types').SopGuide[]>('/api/sop/guides');
  },

  async createSopGuide(payload: {
    stage_group: string;
    title: string;
    content_md: string;
    images: import('../types').SopGuideImage[];
  }): Promise<import('../types').SopGuide> {
    return request<import('../types').SopGuide>('/api/sop/guides', {
      method: 'POST',
      body: JSON.stringify(payload),
    });
  },

  async updateSopGuide(
    id: string,
    payload: Partial<{
      stage_group: string;
      title: string;
      content_md: string;
      images: import('../types').SopGuideImage[];
    }>
  ): Promise<import('../types').SopGuide> {
    return request<import('../types').SopGuide>(`/api/sop/guides/${id}`, {
      method: 'PUT',
      body: JSON.stringify(payload),
    });
  },

  async deleteSopGuide(id: string): Promise<void> {
    await request(`/api/sop/guides/${id}`, { method: 'DELETE' });
  },

  async uploadSopGuideImage(file: File): Promise<import('../types').SopGuideImage> {
    const fd = new FormData();
    fd.append('image', file);
    return request<import('../types').SopGuideImage>('/api/sop/guides/upload-image', {
      method: 'POST',
      body: fd,
    });
  },

  async importSopGuideFile(
    file: File,
    options?: { stage_group?: string; title?: string }
  ): Promise<import('../types').SopGuide> {
    const fd = new FormData();
    fd.append('file', file);
    if (options?.stage_group) fd.append('stage_group', options.stage_group);
    if (options?.title) fd.append('title', options.title);
    return request<import('../types').SopGuide>('/api/sop/guides/import', {
      method: 'POST',
      body: fd,
    });
  },

  // --- Award & Client Agreements ---
  async listClients(): Promise<import('../types').ClientAgreement[]> {
    if (currentStorageMode === 'indexeddb') return [];
    return request<import('../types').ClientAgreement[]>('/api/award/clients');
  },

  async getAwardOverview(batchId: string): Promise<AwardOverview> {
    return request<AwardOverview>(`/api/award/batch/${batchId}`);
  },

  async saveAwardConfig(
    batchId: string,
    payload: {
      client_id?: string;
      category?: string;
      discount_rate?: number | null;
      batch_manager?: string;
      bid_location?: string;
      publicity_date?: string;
      announcement_date?: string;
    }
  ): Promise<{ ok: boolean }> {
    return request<{ ok: boolean }>(`/api/award/batch/${batchId}/config`, {
      method: 'PUT',
      body: JSON.stringify(payload),
    });
  },

  async importClientAgreements(file: File): Promise<{ ok: boolean; count: number }> {
    const fd = new FormData();
    fd.append('file', file);
    return request<{ ok: boolean; count: number }>('/api/award/clients/import', {
      method: 'POST',
      body: fd,
    });
  },

  async importMaterialCatalog(file: File): Promise<{ ok: boolean; count: number }> {
    const fd = new FormData();
    fd.append('file', file);
    return request<{ ok: boolean; count: number }>('/api/award/materials/import', {
      method: 'POST',
      body: fd,
    });
  },

  async generateAwardDoc(batchId: string, template_key: string, overrides: Record<string, string>): Promise<AwardOutput> {
    return request<AwardOutput>(`/api/award/batch/${batchId}/generate`, {
      method: 'POST',
      body: JSON.stringify({ template_key, overrides, sync_to_files: true }),
    });
  },

  async deleteAwardOutput(id: string): Promise<void> {
    await request(`/api/award/outputs/${id}`, { method: 'DELETE' });
  },

  async batchDeleteAwardOutputs(ids: string[]): Promise<{ ok: boolean; deleted: number }> {
    return request<{ ok: boolean; deleted: number }>('/api/award/outputs/batch-delete', {
      method: 'POST',
      body: JSON.stringify({ ids }),
    });
  },

  // --- Archive ---
  async getArchiveOverview(batchId: string): Promise<ArchiveOverview> {
    return request<ArchiveOverview>(`/api/archive/batch/${batchId}`);
  },

  async previewArchiveImport(batchId: string, file: File): Promise<{
    filename: string;
    sheetNames: string[];
    items: ParsedArchivePreviewRow[];
    colorDetectedCount: number;
    format: string;
  }> {
    const fd = new FormData();
    fd.append('file', file);
    return request(`/api/archive/batch/${batchId}/import-preview`, {
      method: 'POST',
      body: fd,
    });
  },

  async commitArchiveImport(
    batchId: string,
    items: ParsedArchivePreviewRow[],
    replace: boolean,
    saveAsTemplate: boolean
  ): Promise<{ ok: boolean; count: number }> {
    return request(`/api/archive/batch/${batchId}/import-commit`, {
      method: 'POST',
      body: JSON.stringify({ items, replace, saveAsTemplate }),
    });
  },

  async addArchiveItem(batchId: string, payload: Record<string, unknown>) {
    return request(`/api/archive/batch/${batchId}/items`, {
      method: 'POST',
      body: JSON.stringify(payload),
    });
  },

  async updateArchiveItem(id: string, payload: Record<string, unknown>) {
    return request(`/api/archive/items/${id}`, {
      method: 'PUT',
      body: JSON.stringify(payload),
    });
  },

  async uploadArchiveItemFile(itemId: string, batchId: string, file: File) {
    const fd = new FormData();
    fd.append('batchId', batchId);
    fd.append('file', file);
    return request(`/api/archive/items/${itemId}/upload`, {
      method: 'POST',
      body: fd,
    });
  },

  async deleteArchiveItem(id: string) {
    return request(`/api/archive/items/${id}`, { method: 'DELETE' });
  },

  // --- Backup & Single-Batch Migration ---
  async importBatchPackage(file: File): Promise<{
    ok: boolean;
    action: 'created' | 'updated';
    batch: { id: string; code: string; name: string; status: string };
    counts: {
      projects: number;
      files: number;
      evalBids: number;
      awardOutputs: number;
      extractedFiles: number;
    };
  }> {
    const fd = new FormData();
    fd.append('file', file);
    return request('/api/backup/import-batch', {
      method: 'POST',
      body: fd,
    });
  },

  async importFullBackup(file: File): Promise<{
    ok: boolean;
    firstBatchId: string;
    counts: {
      batches: number;
      projects: number;
      files: number;
      awardOutputs: number;
      extractedFiles: number;
    };
  }> {
    const fd = new FormData();
    fd.append('file', file);
    return request('/api/backup/import-full', {
      method: 'POST',
      body: fd,
    });
  },

  // --- Standalone Toolbox: Word 文字与数字批量处理 ---
  async processWordBatch(
    files: File[],
    config: {
      modifyEnabled: boolean;
      modifyTemplate: string;
      multiplier: string;
      deleteEnabled: boolean;
      deleteTemplate: string;
      replaceEnabled: boolean;
      replaceUseRegex: boolean;
      replacePairs: Array<{ find: string; replace: string }>;
      extractEnabled: boolean;
    }
  ): Promise<{
    jobId: string;
    totalFiles: number;
    okFiles: number;
    grandTotals: Record<string, number>;
    files: Array<{
      originalName: string;
      ok: boolean;
      convertedFromDoc: boolean;
      counts: Record<string, number>;
      outputs: Array<{ key: string; filename: string; kind: 'docx' | 'xlsx'; size: number }>;
      error?: string;
    }>;
    logs: string[];
    hasOutputs: boolean;
  }> {
    const fd = new FormData();
    files.forEach((f) => fd.append('files', f));
    fd.append('config', JSON.stringify(config));
    return request('/api/toolbox/word-batch/process', {
      method: 'POST',
      body: fd,
    });
  },

  // --- Standalone Toolbox: 专家打分表按模板合并 ---
  async runScoreTemplateMerge(payload: {
    files: File[];
    template?: File | null;
    successList?: File | null;
  }): Promise<{
    jobId: string;
    outputFileName: string;
    totalInputFiles: number;
    validFileCount: number;
    skippedFiles: Array<{ fileName: string; reason: string }>;
    usedTemplate: boolean;
    activeRows: number;
    ignoredRows: number;
    unmatchedRows: number;
    conflictCount: number;
    highlightCount: number;
    headerSummary: Array<{ sheetName: string; headerRows: number }>;
    sheetPreviews?: Array<{
      name: string;
      totalRows: number;
      colCount: number;
      headerRows: number;
      rows: Array<{
        rowNum: number;
        cells: Array<{ text: string; highlighted?: boolean }>;
      }>;
    }>;
    logs: string[];
  }> {
    const fd = new FormData();
    payload.files.forEach((f) => fd.append('files', f));
    if (payload.template) fd.append('template', payload.template);
    if (payload.successList) fd.append('successList', payload.successList);
    return request('/api/toolbox/score-merge/run', {
      method: 'POST',
      body: fd,
    });
  },

  // --- Standalone Toolbox: 履约分匹配填入 ---
  async runPerformanceMatch(payload: {
    lyf: File;
    templates: File[];
    defaultScore?: number;
  }): Promise<{
    jobId: string;
    lyfFileName: string;
    lyfRecordCount: number;
    lyfDuplicates: number;
    defaultScore: number;
    templates: Array<{
      key: string;
      originalName: string;
      outputFileName: string;
      sheetsProcessed: number;
      sheetsSkipped: string[];
      totalRows: number;
      matched: number;
      defaulted: number;
    }>;
    totalRows: number;
    totalMatched: number;
    totalDefaulted: number;
    totalSkippedSheets: number;
    unmatchedItems: Array<{
      templateName: string;
      sheetName: string;
      section: string;
      packageNo: string;
      supplier: string;
    }>;
    logs: string[];
  }> {
    const fd = new FormData();
    fd.append('lyf', payload.lyf);
    payload.templates.forEach((f) => fd.append('templates', f));
    if (payload.defaultScore !== undefined) {
      fd.append('defaultScore', String(payload.defaultScore));
    }
    return request('/api/toolbox/perf-match/run', {
      method: 'POST',
      body: fd,
    });
  },

  // --- Standalone Toolbox: 清单合规性检验 ---
  async runChecklistCompliance(payload: {
    files: File[];
    keywords?: string;
  }): Promise<{
    jobId: string;
    keywords: string[];
    counts: {
      total: number;
      ok: number;
      missing: number;
      abnormal: number;
      unpackFailed: number;
      notice: number;
    };
    items: Array<{
      batch: string;
      specId: string;
      status: 'OK' | '缺失' | '异常' | '解压失败' | '注意';
      fileName: string;
      fileSize: number | null;
      relativePath: string;
      note: string;
    }>;
    reportFileName: string;
    logs: string[];
  }> {
    const fd = new FormData();
    payload.files.forEach((f) => fd.append('files', f));
    if (payload.keywords !== undefined) fd.append('keywords', payload.keywords);
    return request('/api/toolbox/checklist-compliance/run', {
      method: 'POST',
      body: fd,
    });
  },

  // --- Standalone Toolbox: 专家随机打分 ---
  async runRandomScoreGenerator(payload: {
    template: File;
    nameMode: 'manual' | 'excel' | 'auto';
    namesFile?: File | null;
    manualNames?: string;
    nExperts?: number;
    maxTotalDiff?: number;
  }): Promise<{
    jobId: string;
    nExperts: number;
    maxTotalDiff: number;
    maxSameCount: number;
    expertNames: string[];
    sheetSummaries: Array<{
      sheetName: string;
      scoreColRange: string;
      objCount: number;
      subjCount: number;
      dataRowCount: number;
      skipCount: number;
    }>;
    skippedRows: Array<{
      sheetName: string;
      row: number;
      section: string;
      packageNo: string;
      supplier: string;
      reason: string;
    }>;
    totalCheckedGroups: number;
    violations: number;
    outputs: Array<{ key: string; filename: string; expertName: string }>;
    logs: string[];
  }> {
    const fd = new FormData();
    fd.append('template', payload.template);
    fd.append('nameMode', payload.nameMode);
    if (payload.namesFile) fd.append('namesFile', payload.namesFile);
    if (payload.manualNames !== undefined) fd.append('manualNames', payload.manualNames);
    if (payload.nExperts !== undefined) fd.append('nExperts', String(payload.nExperts));
    if (payload.maxTotalDiff !== undefined) fd.append('maxTotalDiff', String(payload.maxTotalDiff));
    return request('/api/toolbox/random-score/run', {
      method: 'POST',
      body: fd,
    });
  },
};
