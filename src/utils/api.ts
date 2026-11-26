let authToken: string | null = null;
let onUnauthorizedCallback: (() => void) | null = null;

export function setAuthToken(token: string | null) {
  authToken = token;
}

export function getAuthToken(): string | null {
  return authToken;
}

export function setOnUnauthorized(cb: (() => void) | null) {
  onUnauthorizedCallback = cb;
}

export async function apiRequest<T>(url: string, options: RequestInit = {}): Promise<T> {
  const headers = new Headers(options.headers || {});
  if (authToken) {
    headers.set('Authorization', `Bearer ${authToken}`);
  }
  if (!(options.body instanceof FormData) && !headers.has('Content-Type') && options.body) {
    headers.set('Content-Type', 'application/json');
  }

  const res = await fetch(url, {
    ...options,
    headers,
  });

  if (res.status === 401) {
    authToken = null;
    if (onUnauthorizedCallback) {
      onUnauthorizedCallback();
    }
  }

  const contentType = res.headers.get('content-type') || '';
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
  const headers = new Headers();
  if (authToken) {
    headers.set('Authorization', `Bearer ${authToken}`);
  }
  const res = await fetch(url, { headers });
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
