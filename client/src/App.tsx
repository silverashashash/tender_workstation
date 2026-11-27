import React from 'react';
import { BrowserRouter, Navigate, Route, Routes } from 'react-router-dom';
import { Layout } from './components/Layout';
import { BatchProvider, useBatch } from './context/BatchContext';
import { ArchivePage } from './pages/ArchivePage';
import { AwardPage } from './pages/AwardPage';
import { BiddingPage } from './pages/BiddingPage';
import { DashboardPage } from './pages/DashboardPage';
import { EvaluationPage } from './pages/EvaluationPage';
import { LoginPage } from './pages/LoginPage';
import { ToolboxPage } from './pages/ToolboxPage';

function AppRoutes() {
  const { loading, authenticated } = useBatch();

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-[#F7F8FA] text-sm text-slate-400">
        正在初始化招标代理工作台...
      </div>
    );
  }

  if (!authenticated) {
    return <LoginPage />;
  }

  return (
    <Routes>
      <Route path="/" element={<Layout />}>
        <Route index element={<DashboardPage />} />
        <Route path="bidding" element={<BiddingPage />} />
        <Route path="evaluation" element={<EvaluationPage />} />
        <Route path="evaluation/:section" element={<EvaluationPage />} />
        <Route path="evaluation/:section/:tool" element={<EvaluationPage />} />
        <Route path="award" element={<AwardPage />} />
        <Route path="archive" element={<ArchivePage />} />
        <Route path="toolbox" element={<ToolboxPage />} />
        <Route path="toolbox/:toolKey" element={<ToolboxPage />} />
        <Route path="*" element={<Navigate to="/" replace />} />
      </Route>
    </Routes>
  );
}

export default function App() {
  return (
    <BrowserRouter>
      <BatchProvider>
        <AppRoutes />
      </BatchProvider>
    </BrowserRouter>
  );
}
