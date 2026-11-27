import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import cookieParser from 'cookie-parser';
import cors from 'cors';
import dotenv from 'dotenv';
import express from 'express';
import { handleAuthStatus, handleLogin, handleLogout, requireAuth } from './auth.js';
import { DATA_DIR, initSchema } from './db.js';
import { seedMaterialCatalogIfEmpty } from './materialCatalog.js';
import { seedIfEmpty } from './seed.js';
import { archiveRouter } from './routes/archive.js';
import { awardRouter } from './routes/award.js';
import { backupRouter } from './routes/backup.js';
import { batchesRouter } from './routes/batches.js';
import { biddingRouter } from './routes/bidding.js';
import { evaluationRouter } from './routes/evaluation.js';
import { filesRouter } from './routes/files.js';
import { sopRouter } from './routes/sop.js';
import { standaloneToolboxRouter } from './routes/standaloneToolbox.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
// Load .env from project root first, then server directory
dotenv.config({ path: path.resolve(__dirname, '../../.env') });
dotenv.config({ path: path.resolve(__dirname, '../.env') });

initSchema();
seedIfEmpty();
seedMaterialCatalogIfEmpty();

const app = express();
const PORT = Number(process.env.PORT || 3001);

app.use(cors({ origin: true, credentials: true }));
app.use(express.json({ limit: '20mb' }));
app.use(cookieParser());

// Health & Auth routes (public)
app.get('/api/health', (_req, res) => {
  res.json({ ok: true, storage: 'sqlite', dataDir: DATA_DIR });
});
app.get('/api/auth/status', handleAuthStatus);
app.post('/api/auth/login', handleLogin);
app.post('/api/auth/logout', handleLogout);

// Protected API routes
app.use('/api/batches', requireAuth, batchesRouter);
app.use('/api/bidding', requireAuth, biddingRouter);
app.use('/api/evaluation', requireAuth, evaluationRouter);
app.use('/api/files', requireAuth, filesRouter);
app.use('/api/sop', requireAuth, sopRouter);
app.use('/api/award', requireAuth, awardRouter);
app.use('/api/archive', requireAuth, archiveRouter);
app.use('/api/backup', requireAuth, backupRouter);
app.use('/api/toolbox', requireAuth, standaloneToolboxRouter);

// Serve built frontend in production (single-port personal website deployment)
const clientDist = path.resolve(__dirname, '../../client/dist');
if (fs.existsSync(clientDist)) {
  app.use(express.static(clientDist));
  app.get('*', (req, res, next) => {
    if (req.path.startsWith('/api')) return next();
    res.sendFile(path.join(clientDist, 'index.html'));
  });
}

app.listen(PORT, () => {
  console.log(`[tender-workstation] API server running at http://localhost:${PORT}`);
  console.log(`[tender-workstation] SQLite data directory: ${DATA_DIR}`);
});
