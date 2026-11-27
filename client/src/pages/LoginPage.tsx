import React, { useState } from 'react';
import { Lock } from 'lucide-react';
import { api } from '../api/client';
import { useBatch } from '../context/BatchContext';

export function LoginPage() {
  const { setAuthenticated, refreshBatches } = useBatch();
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    setError('');
    setLoading(true);
    try {
      await api.login(password);
      setAuthenticated(true);
      await refreshBatches();
    } catch (err) {
      setError((err as Error).message || '解锁失败');
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="min-h-screen flex items-center justify-center bg-[#F7F8FA] p-4">
      <div className="w-full max-w-sm rounded-xl border border-slate-200 bg-white p-6 shadow-sm">
        <div className="flex flex-col items-center text-center mb-6">
          <div className="flex h-11 w-11 items-center justify-center rounded-xl bg-[#3B5B7E]/10 text-[#3B5B7E] mb-3">
            <Lock className="h-5 w-5" />
          </div>
          <h1 className="text-base font-semibold text-slate-800">招标代理 SOP 工作台</h1>
          <p className="mt-1 text-xs text-slate-400">公网访问已开启保护，请输入访问密码解锁</p>
        </div>

        <form onSubmit={handleSubmit} className="space-y-4">
          <div>
            <input
              type="password"
              required
              autoFocus
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              placeholder="输入访问密码"
              className="w-full rounded-lg border border-slate-200 px-3.5 py-2.5 text-sm focus:border-[#3B5B7E] focus:outline-none"
            />
          </div>
          {error && <div className="text-xs text-[#A94442]">{error}</div>}
          <button
            type="submit"
            disabled={loading}
            className="w-full rounded-lg bg-[#3B5B7E] py-2.5 text-sm font-medium text-white hover:bg-[#2E4866] disabled:opacity-50 transition-colors"
          >
            {loading ? '验证中...' : '解锁工作台'}
          </button>
        </form>
      </div>
    </div>
  );
}
