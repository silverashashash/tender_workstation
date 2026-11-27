import React, { createContext, useCallback, useContext, useEffect, useState } from 'react';
import { api, type StorageMode } from '../api/client';
import type { Batch } from '../types';

interface BatchContextValue {
  batches: Batch[];
  currentBatchId: string;
  currentBatch: Batch | null;
  setCurrentBatchId: (id: string) => void;
  refreshBatches: () => Promise<void>;
  loading: boolean;
  storageMode: StorageMode;
  authEnabled: boolean;
  authenticated: boolean;
  setAuthenticated: (val: boolean) => void;
}

const BatchContext = createContext<BatchContextValue | null>(null);
const STORAGE_KEY = 'tw_current_batch_id';

export function BatchProvider({ children }: { children: React.ReactNode }) {
  const [batches, setBatches] = useState<Batch[]>([]);
  const [currentBatchId, setCurrentBatchIdState] = useState<string>(() => localStorage.getItem(STORAGE_KEY) || '');
  const [loading, setLoading] = useState(true);
  const [storageMode, setStorageMode] = useState<StorageMode>('sqlite');
  const [authEnabled, setAuthEnabled] = useState(false);
  const [authenticated, setAuthenticated] = useState(true);

  const setCurrentBatchId = useCallback((id: string) => {
    setCurrentBatchIdState(id);
    localStorage.setItem(STORAGE_KEY, id);
  }, []);

  const refreshBatches = useCallback(async () => {
    try {
      const list = await api.listBatches();
      setBatches(list);
      setCurrentBatchIdState((prev) => {
        if (prev && list.some((b) => b.id === prev)) return prev;
        const first = list[0]?.id || '';
        if (first) localStorage.setItem(STORAGE_KEY, first);
        return first;
      });
    } catch {
      // Unauthorized handled via event
    }
  }, []);

  useEffect(() => {
    const onUnauth = () => setAuthenticated(false);
    window.addEventListener('tw:unauthorized', onUnauth);
    return () => window.removeEventListener('tw:unauthorized', onUnauth);
  }, []);

  useEffect(() => {
    (async () => {
      setLoading(true);
      const status = await api.checkHealth();
      setStorageMode(status.storage);
      setAuthEnabled(status.authEnabled);
      setAuthenticated(status.authenticated);
      if (status.authenticated) {
        await refreshBatches();
      }
      setLoading(false);
    })();
  }, [refreshBatches]);

  const currentBatch = batches.find((b) => b.id === currentBatchId) || batches[0] || null;

  return (
    <BatchContext.Provider
      value={{
        batches,
        currentBatchId: currentBatch?.id || '',
        currentBatch,
        setCurrentBatchId,
        refreshBatches,
        loading,
        storageMode,
        authEnabled,
        authenticated,
        setAuthenticated,
      }}
    >
      {children}
    </BatchContext.Provider>
  );
}

export function useBatch() {
  const ctx = useContext(BatchContext);
  if (!ctx) throw new Error('useBatch must be used within BatchProvider');
  return ctx;
}
