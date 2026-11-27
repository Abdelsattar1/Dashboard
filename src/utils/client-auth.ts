import type Database from 'better-sqlite3';
import { ORIGINAL_SECTORS } from '../../server/excel-engine.ts';
import type { UserProfile, UserRole } from '../types/bi.ts';

export const CLIENT_DEFAULT_ADMIN_USERNAME = 'admin';
export const CLIENT_DEFAULT_ADMIN_SEED = ['Admin', '@', '2026'].join('');

const CLIENT_SESSIONS = new Map<string, { userId: number; exp: number }>();

function bytesToHex(bytes: Uint8Array): string {
  return Array.from(bytes, (b) => b.toString(16).padStart(2, '0')).join('');
}

function hexToBytes(hex: string): Uint8Array {
  const clean = hex.trim();
  const out = new Uint8Array(clean.length / 2);
  for (let i = 0; i < out.length; i++) {
    out[i] = parseInt(clean.slice(i * 2, i * 2 + 2), 16);
  }
  return out;
}

function timingSafeHexEqual(a: string, b: string): boolean {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) {
    diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return diff === 0;
}

export async function hashPasswordClient(password: string): Promise<string> {
  const saltBytes = new Uint8Array(16);
  window.crypto.getRandomValues(saltBytes);
  const saltHex = bytesToHex(saltBytes);
  const enc = new TextEncoder();
  const keyMaterial = await window.crypto.subtle.importKey(
    'raw',
    enc.encode(password),
    'PBKDF2',
    false,
    ['deriveBits']
  );
  const bits = await window.crypto.subtle.deriveBits(
    {
      name: 'PBKDF2',
      salt: saltBytes,
      iterations: 60000,
      hash: 'SHA-256',
    },
    keyMaterial,
    256
  );
  return `pbkdf2_sha256$${saltHex}$${bytesToHex(new Uint8Array(bits))}`;
}

export async function verifyPasswordClient(password: string, storedHash: string): Promise<boolean> {
  const parts = storedHash.split('$');
  if (parts.length !== 3 || parts[0] !== 'pbkdf2_sha256') return false;
  const [, saltHex, expectedHex] = parts;
  if (!saltHex || !expectedHex) return false;
  const saltBytes = hexToBytes(saltHex);
  const enc = new TextEncoder();
  const keyMaterial = await window.crypto.subtle.importKey(
    'raw',
    enc.encode(password),
    'PBKDF2',
    false,
    ['deriveBits']
  );
  const bits = await window.crypto.subtle.deriveBits(
    {
      name: 'PBKDF2',
      salt: saltBytes.buffer as ArrayBuffer,
      iterations: 60000,
      hash: 'SHA-256',
    },
    keyMaterial,
    256
  );
  return timingSafeHexEqual(bytesToHex(new Uint8Array(bits)), expectedHex);
}

export interface ClientDbUserRow {
  id: number;
  username: string;
  full_name: string;
  email: string | null;
  role: string;
  allowed_sectors: string;
  password_hash: string;
  must_change_password: number;
  is_active: number;
  created_at: string;
}

export function normalizeRoleClient(raw: unknown): UserRole {
  const val = String(raw || '').trim();
  if (val === 'admin') return 'admin';
  if (val === 'sector_manager') return 'sector_manager';
  return 'standard_user';
}

export function mapClientUserRow(row: ClientDbUserRow): UserProfile {
  let allowedSectors: string[] = [];
  try {
    const parsed = JSON.parse(row.allowed_sectors);
    if (Array.isArray(parsed)) allowedSectors = parsed.map((s) => String(s));
  } catch {
    allowedSectors = [];
  }
  return {
    id: row.id,
    username: row.username,
    fullName: row.full_name,
    email: row.email,
    role: normalizeRoleClient(row.role),
    allowedSectors,
    isActive: row.is_active === 1,
    mustChangePassword: row.must_change_password === 1,
    createdAt: row.created_at,
  };
}

export async function ensureDefaultAdminSeeded(db: Database.Database): Promise<{ seededNow: boolean; defaultAdminPendingChange: boolean }> {
  const countRow = db
    .prepare(`SELECT COUNT(*) AS cnt FROM users WHERE role = 'admin'`)
    .get() as { cnt: number };

  if (countRow.cnt === 0) {
    const now = new Date().toISOString();
    const hash = await hashPasswordClient(CLIENT_DEFAULT_ADMIN_SEED);
    db.prepare(
      `INSERT INTO users (username, full_name, email, role, allowed_sectors, password_hash, must_change_password, is_active, created_at)
       VALUES (?, ?, ?, 'admin', ?, ?, 1, 1, ?)`
    ).run(
      CLIENT_DEFAULT_ADMIN_USERNAME,
      'مدير النظام',
      null,
      JSON.stringify([...ORIGINAL_SECTORS]),
      hash,
      now
    );
    return { seededNow: true, defaultAdminPendingChange: true };
  }

  const defaultAdminRow = db
    .prepare(`SELECT must_change_password FROM users WHERE lower(username) = 'admin' AND role = 'admin' AND is_active = 1`)
    .get() as { must_change_password: number } | undefined;

  return {
    seededNow: false,
    defaultAdminPendingChange: Boolean(defaultAdminRow && defaultAdminRow.must_change_password === 1),
  };
}

export function issueClientSessionToken(userId: number): string {
  const bytes = new Uint8Array(24);
  window.crypto.getRandomValues(bytes);
  const token = `local_${window.crypto.randomUUID()}_${bytesToHex(bytes)}`;
  CLIENT_SESSIONS.set(token, {
    userId,
    exp: Date.now() + 1000 * 60 * 60 * 12,
  });
  return token;
}

export function resolveClientSessionUser(db: Database.Database, token: string | null): UserProfile | null {
  if (!token) return null;
  const sess = CLIENT_SESSIONS.get(token);
  if (!sess || sess.exp < Date.now()) {
    if (sess) CLIENT_SESSIONS.delete(token);
    return null;
  }
  const row = db
    .prepare(`SELECT * FROM users WHERE id = ? AND is_active = 1`)
    .get(sess.userId) as ClientDbUserRow | undefined;
  return row ? mapClientUserRow(row) : null;
}

export function revokeClientSessionToken(token: string | null): void {
  if (token) {
    CLIENT_SESSIONS.delete(token);
  }
}
