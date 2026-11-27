import { revokeClientSessionToken } from './client-auth.ts';
import { ClientHttpError, handleClientApiRequest, handleClientDownload } from './client-api.ts';
import { isSupabaseConfigured } from './supabase-sync.ts';

let authToken: string | null = null;
let onUnauthorizedCallback: (() => void) | null = null;

type ApiExecutionMode = 'server' | 'client';
let detectedMode: ApiExecutionMode | null = null;
let detectPromise: Promise<ApiExecutionMode> | null = null;

export function setAuthToken(token: string | null) {
  if (!token && authToken) {
    revokeClientSessionToken(authToken);
  }
  authToken = token;
}

export function getAuthToken(): string | null {
  return authToken;
}

export function setOnUnauthorized(cb: (() => void) | null) {
  onUnauthorizedCallback = cb;
}

async function resolveExecutionMode(): Promise<ApiExecutionMode> {
  if (isSupabaseConfigured()) {
    detectedMode = 'client';
    return 'client';
  }
  if (detectedMode) return detectedMode;
  if (detectPromise) return detectPromise;

  detectPromise = (async (): Promise<ApiExecutionMode> => {
    try {
      const res = await fetch('/api/health', { method: 'GET', cache: 'no-store' });
      const contentType = res.headers.get('content-type') || '';
      if (res.ok && contentType.includes('application/json')) {
        const data = (await res.json()) as { status?: string };
        if (data && data.status === 'ok') {
          detectedMode = 'server';
          return 'server';
        }
      }
    } catch {
      // Static deploy or offline host -> transparently use browser-local client engine
    }
    detectedMode = 'client';
    return 'client';
  })();

  return detectPromise;
}

async function invokeClientApi<T>(url: string, options: RequestInit): Promise<T> {
  try {
    return await handleClientApiRequest<T>(url, options, authToken);
  } catch (err) {
    if (err instanceof ClientHttpError && err.status === 401) {
      authToken = null;
      if (onUnauthorizedCallback) {
        onUnauthorizedCallback();
      }
    }
    throw err;
  }
}

export async function apiRequest<T>(url: string, options: RequestInit = {}): Promise<T> {
  const mode = await resolveExecutionMode();

  if (mode === 'client') {
    return invokeClientApi<T>(url, options);
  }

  const headers = new Headers(options.headers || {});
  if (authToken) {
    headers.set('Authorization', `Bearer ${authToken}`);
  }
  if (!(options.body instanceof FormData) && !headers.has('Content-Type') && options.body) {
    headers.set('Content-Type', 'application/json');
  }

  let res: Response;
  try {
    res = await fetch(url, {
      ...options,
      headers,
    });
  } catch {
    // Fallback gracefully to browser-local engine if server becomes unreachable on static hosting
    detectedMode = 'client';
    return invokeClientApi<T>(url, options);
  }

  const contentType = res.headers.get('content-type') || '';
  if (res.status === 404 && !contentType.includes('application/json')) {
    // Vercel static NOT_FOUND returns non-JSON 404 -> switch transparently to client engine
    detectedMode = 'client';
    return invokeClientApi<T>(url, options);
  }

  if (res.status === 401) {
    authToken = null;
    if (onUnauthorizedCallback) {
      onUnauthorizedCallback();
    }
  }

  if (contentType.includes('application/json')) {
    const data = await res.json();
    if (!res.ok) {
      throw new Error(data?.error || `خطأ في الطلب (${res.status})`);
    }
    return data as T;
  }

  const text = await res.text();
  if (!res.ok) {
    throw new Error(text || `خطأ في الطلب (${res.status})`);
  }

  try {
    return JSON.parse(text) as T;
  } catch {
    throw new Error('استجابة غير صالحة من الخادم (ليست بصيغة JSON)');
  }
}

export async function downloadProtectedFile(url: string, fallbackFilename: string): Promise<void> {
  const mode = await resolveExecutionMode();

  if (mode === 'client') {
    const { blob, filename } = await handleClientDownload(url, fallbackFilename, authToken);
    const objectUrl = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = objectUrl;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(objectUrl);
    return;
  }

  const headers = new Headers();
  if (authToken) {
    headers.set('Authorization', `Bearer ${authToken}`);
  }
  const res = await fetch(url, { headers });
  const contentType = res.headers.get('content-type') || '';
  if (res.status === 404 && !contentType.includes('application/json')) {
    detectedMode = 'client';
    const { blob, filename } = await handleClientDownload(url, fallbackFilename, authToken);
    const objectUrl = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = objectUrl;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(objectUrl);
    return;
  }

  if (!res.ok) {
    let msg = `تعذر تنزيل الملف (${res.status})`;
    try {
      const json = await res.json();
      if (json.error) msg = json.error;
    } catch {
      // ignore
    }
    throw new Error(msg);
  }
  const blob = await res.blob();
  const objectUrl = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = objectUrl;
  a.download = fallbackFilename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(objectUrl);
}
