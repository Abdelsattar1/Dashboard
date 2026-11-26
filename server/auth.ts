import crypto from 'crypto';
import type { NextFunction, Request, Response } from 'express';
import type Database from 'better-sqlite3';
import type { UserProfile, UserRole } from '../src/types/bi.ts';

const EPHEMERAL_SECRET = crypto.randomBytes(32).toString('hex');

function getSessionSecret(): string {
  const envSecret = process.env.SESSION_SECRET?.trim();
  return envSecret && envSecret.length >= 16 ? envSecret : EPHEMERAL_SECRET;
}

export function hashPassword(password: string): string {
  const salt = crypto.randomBytes(16).toString('hex');
  const derived = crypto.scryptSync(password, salt, 64).toString('hex');
  return `${salt}:${derived}`;
}

export function verifyPassword(password: string, storedHash: string): boolean {
  const parts = storedHash.split(':');
  if (parts.length !== 2) return false;
  const [salt, originalHex] = parts;
  if (!salt || !originalHex) return false;
  const derivedBuf = crypto.scryptSync(password, salt, 64);
  const originalBuf = Buffer.from(originalHex, 'hex');
  if (derivedBuf.length !== originalBuf.length) return false;
  return crypto.timingSafeEqual(derivedBuf, originalBuf);
}

interface TokenPayload {
  userId: number;
  username: string;
  role: UserRole;
  exp: number;
}

export function signToken(user: { id: number; username: string; role: UserRole }): string {
  const payload: TokenPayload = {
    userId: user.id,
    username: user.username,
    role: user.role,
    exp: Date.now() + 1000 * 60 * 60 * 12,
  };
  const payloadB64 = Buffer.from(JSON.stringify(payload), 'utf8').toString('base64url');
  const sig = crypto
    .createHmac('sha256', getSessionSecret())
    .update(payloadB64)
    .digest('base64url');
  return `${payloadB64}.${sig}`;
}

export function verifyToken(token: string): TokenPayload | null {
  const parts = token.split('.');
  if (parts.length !== 2) return null;
  const [payloadB64, sigB64] = parts;
  if (!payloadB64 || !sigB64) return null;

  const expectedSig = crypto
    .createHmac('sha256', getSessionSecret())
    .update(payloadB64)
    .digest('base64url');

  const sigBuf = Buffer.from(sigB64);
  const expectedBuf = Buffer.from(expectedSig);
  if (sigBuf.length !== expectedBuf.length) return null;
  if (!crypto.timingSafeEqual(sigBuf, expectedBuf)) return null;

  try {
    const parsed = JSON.parse(Buffer.from(payloadB64, 'base64url').toString('utf8')) as TokenPayload;
    if (!parsed.exp || parsed.exp < Date.now()) return null;
    return parsed;
  } catch {
    return null;
  }
}

export interface DbUserRow {
  id: number;
  username: string;
  full_name: string;
  email: string | null;
  role: UserRole;
  allowed_sectors: string;
  password_hash: string;
  must_change_password: number;
  is_active: number;
  created_at: string;
}

export function mapDbUserToProfile(row: DbUserRow): UserProfile {
  let allowedSectors: string[] = [];
  try {
    const parsed = JSON.parse(row.allowed_sectors);
    if (Array.isArray(parsed)) {
      allowedSectors = parsed.map((s) => String(s));
    }
  } catch {
    allowedSectors = [];
  }

  return {
    id: row.id,
    username: row.username,
    fullName: row.full_name,
    email: row.email,
    role: row.role,
    allowedSectors,
    isActive: row.is_active === 1,
    mustChangePassword: row.must_change_password === 1,
    createdAt: row.created_at,
  };
}

export interface AuthenticatedRequest extends Request {
  user?: UserProfile;
}

export function createAuthMiddleware(db: Database.Database) {
  return (req: AuthenticatedRequest, res: Response, next: NextFunction) => {
    const authHeader = req.headers.authorization;
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
      res.status(401).json({ error: 'مطلوب تسجيل الدخول للوصول إلى هذه البيانات' });
      return;
    }

    const token = authHeader.slice('Bearer '.length).trim();
    const payload = verifyToken(token);
    if (!payload) {
      res.status(401).json({ error: 'انتهت صلاحية الجلسة أو رمز التحقق غير صالح' });
      return;
    }

    const row = db
      .prepare('SELECT * FROM users WHERE id = ? AND is_active = 1')
      .get(payload.userId) as DbUserRow | undefined;

    if (!row) {
      res.status(401).json({ error: 'الحساب غير موجود أو موقوف' });
      return;
    }

    req.user = mapDbUserToProfile(row);
    next();
  };
}

export function requireAdmin(req: AuthenticatedRequest, res: Response, next: NextFunction) {
  if (!req.user || req.user.role !== 'admin') {
    res.status(403).json({ error: 'هذه العملية متاحة لمدير النظام فقط' });
    return;
  }
  next();
}
