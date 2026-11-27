import crypto from 'node:crypto';
import type { Request, Response, NextFunction } from 'express';

const COOKIE_NAME = 'tw_session';
const SESSION_TTL_MS = 7 * 24 * 60 * 60 * 1000; // 7 days

// Simple in-memory rate limiter for login attempts
const loginAttempts = new Map<string, { count: number; resetAt: number }>();

export function isAuthEnabled(): boolean {
  return process.env.AUTH_ENABLED === 'true';
}

function getSecret(): string {
  return process.env.AUTH_SECRET || 'tender-workstation-dev-secret-change-in-prod';
}

function getPassword(): string {
  return process.env.AUTH_PASSWORD || 'admin123';
}

export function verifyAdminPassword(password?: string): boolean {
  if (!isAuthEnabled()) return true;
  return Boolean(password && password === getPassword());
}

function signToken(exp: number): string {
  const payload = `${exp}`;
  const sig = crypto.createHmac('sha256', getSecret()).update(payload).digest('hex');
  return `${payload}.${sig}`;
}

export function verifyToken(token: string | undefined): boolean {
  if (!token) return false;
  const parts = token.split('.');
  if (parts.length !== 2) return false;
  const [expStr, sig] = parts;
  const exp = Number(expStr);
  if (!Number.isFinite(exp) || Date.now() > exp) return false;
  const expected = crypto.createHmac('sha256', getSecret()).update(expStr).digest('hex');
  if (sig.length !== expected.length) return false;
  return crypto.timingSafeEqual(Buffer.from(sig), Buffer.from(expected));
}

export function requireAuth(req: Request, res: Response, next: NextFunction) {
  if (!isAuthEnabled()) {
    return next();
  }
  const token = req.cookies?.[COOKIE_NAME];
  if (verifyToken(token)) {
    return next();
  }
  return res.status(401).json({ error: '未登录或会话已过期，请先输入访问密码解锁工作台' });
}

export function handleAuthStatus(req: Request, res: Response) {
  const enabled = isAuthEnabled();
  const authenticated = !enabled || verifyToken(req.cookies?.[COOKIE_NAME]);
  res.json({ authEnabled: enabled, authenticated });
}

export function handleLogin(req: Request, res: Response) {
  if (!isAuthEnabled()) {
    return res.json({ ok: true, authEnabled: false });
  }
  const ip = req.ip || 'local';
  const now = Date.now();
  const record = loginAttempts.get(ip);
  if (record && now < record.resetAt && record.count >= 10) {
    return res.status(429).json({ error: '登录尝试过于频繁，请 15 分钟后再试' });
  }

  const { password } = req.body as { password?: string };
  if (!password || password !== getPassword()) {
    const nextCount = record && now < record.resetAt ? record.count + 1 : 1;
    loginAttempts.set(ip, { count: nextCount, resetAt: now + 15 * 60 * 1000 });
    return res.status(401).json({ error: '密码不正确，请重试' });
  }

  loginAttempts.delete(ip);
  const exp = now + SESSION_TTL_MS;
  const token = signToken(exp);
  res.cookie(COOKIE_NAME, token, {
    httpOnly: true,
    sameSite: 'lax',
    secure: process.env.NODE_ENV === 'production' && req.secure,
    maxAge: SESSION_TTL_MS,
    path: '/',
  });
  return res.json({ ok: true, authenticated: true });
}

export function handleLogout(_req: Request, res: Response) {
  res.clearCookie(COOKIE_NAME, { path: '/' });
  res.json({ ok: true });
}
