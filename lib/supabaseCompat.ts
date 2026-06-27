/* eslint-disable @typescript-eslint/no-explicit-any */
type ApiResult<T = any> = { data: T; error: null; count?: number | null } | { data: null; error: { message: string }; count?: number | null };

type Filter =
  | { type: "eq"; column: string; value: any }
  | { type: "in"; column: string; value: any[] }
  | { type: "is"; column: string; value: any }
  | { type: "textSearch"; column: string; value: string };

type QueryPayload = {
  table: string;
  action: "select" | "insert" | "update" | "delete";
  select?: string;
  payload?: any;
  filters: Filter[];
  orders: Array<{ column: string; ascending: boolean; nullsFirst?: boolean }>;
  limit?: number;
  single?: boolean;
  maybeSingle?: boolean;
  head?: boolean;
  count?: "exact" | "planned" | "estimated" | null;
};

const API_BASE_URL =
  process.env.NEXT_PUBLIC_API_BASE_URL ||
  process.env.EXPO_PUBLIC_API_BASE_URL ||
  "";

const AUTH_TOKEN_KEY = "auth_token";

type AuthSession = {
  access_token: string;
  user: any;
};

type AuthEvent = "SIGNED_IN" | "SIGNED_OUT";
type AuthListener = (event: AuthEvent, session: AuthSession | null) => void;

const authListeners = new Set<AuthListener>();

function getClientToken() {
  if (typeof window === "undefined") return null;
  return localStorage.getItem(AUTH_TOKEN_KEY);
}

function setClientToken(token: string | null) {
  if (typeof window === "undefined") return;
  if (!token) localStorage.removeItem(AUTH_TOKEN_KEY);
  else localStorage.setItem(AUTH_TOKEN_KEY, token);
}

function authHeaders() {
  const token = getClientToken();
  if (!token) return {};
  return { Authorization: `Bearer ${token}` };
}

async function apiFetch(path: string, init?: RequestInit) {
  const res = await fetch(`${API_BASE_URL}${path}`, {
    ...init,
    credentials: "include",
    headers: {
      ...(init?.headers || {}),
      ...authHeaders(),
    },
  });
  return res;
}

function normalizeError(e: any) {
  return { message: e?.message || String(e) || "Unknown error" };
}

class QueryBuilder {
  private payload: QueryPayload;

  constructor(table: string) {
    this.payload = {
      table,
      action: "select",
      filters: [],
      orders: [],
      count: null,
    };
  }

  select(columns = "*", options?: { head?: boolean; count?: "exact" | "planned" | "estimated" }) {
    this.payload.action = "select";
    this.payload.select = columns;
    this.payload.head = !!options?.head;
    this.payload.count = options?.count || null;
    return this;
  }

  insert(values: any) {
    this.payload.action = "insert";
    this.payload.payload = values;
    return this;
  }

  update(values: any) {
    this.payload.action = "update";
    this.payload.payload = values;
    return this;
  }

  delete() {
    this.payload.action = "delete";
    return this;
  }

  eq(column: string, value: any) {
    this.payload.filters.push({ type: "eq", column, value });
    return this;
  }

  in(column: string, value: any[]) {
    this.payload.filters.push({ type: "in", column, value });
    return this;
  }

  is(column: string, value: any) {
    this.payload.filters.push({ type: "is", column, value });
    return this;
  }

  textSearch(column: string, value: string) {
    this.payload.filters.push({ type: "textSearch", column, value });
    return this;
  }

  order(column: string, opts?: { ascending?: boolean; nullsFirst?: boolean }) {
    this.payload.orders.push({
      column,
      ascending: opts?.ascending ?? true,
      nullsFirst: opts?.nullsFirst,
    });
    return this;
  }

  limit(n: number) {
    this.payload.limit = n;
    return this;
  }

  single() {
    this.payload.single = true;
    this.payload.maybeSingle = false;
    return this;
  }

  maybeSingle() {
    this.payload.maybeSingle = true;
    this.payload.single = false;
    return this;
  }

  async execute(): Promise<ApiResult<any>> {
    try {
      const res = await apiFetch("/api/db/query", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(this.payload),
      });
      const json = await res.json().catch(() => ({}));
      if (!res.ok) {
        return { data: null, error: { message: json?.error || `Request failed (${res.status})` } };
      }
      return {
        data: json?.data ?? null,
        error: null,
        count: json?.count ?? null,
      };
    } catch (e: any) {
      return { data: null, error: normalizeError(e) };
    }
  }

  then<TResult1 = ApiResult<any>, TResult2 = never>(
    onfulfilled?: ((value: ApiResult<any>) => TResult1 | PromiseLike<TResult1>) | null,
    onrejected?: ((reason: any) => TResult2 | PromiseLike<TResult2>) | null
  ): Promise<TResult1 | TResult2> {
    return this.execute().then(onfulfilled as any, onrejected as any);
  }
}

const auth = {
  async signInWithPassword(params: { email: string; password: string }) {
    try {
      const res = await apiFetch("/api/auth/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(params),
      });
      const json = await res.json().catch(() => ({}));
      if (!res.ok) {
        return { data: { session: null, user: null }, error: { message: json?.error || "Login failed" } };
      }

      const token = json?.token || null;
      if (token) setClientToken(token);
      const session: AuthSession = { access_token: token || "", user: json?.user };

      for (const listener of authListeners) listener("SIGNED_IN", session);
      return { data: { session, user: json?.user }, error: null };
    } catch (e: any) {
      return { data: { session: null, user: null }, error: normalizeError(e) };
    }
  },

  async getSession() {
    try {
      const res = await apiFetch("/api/auth/me", { method: "GET" });
      const json = await res.json().catch(() => ({}));
      if (!res.ok || !json?.user) return { data: { session: null }, error: null };
      const token = json?.token || getClientToken() || "";
      if (token) setClientToken(token);
      return {
        data: {
          session: {
            access_token: token,
            user: json.user,
          },
        },
        error: null,
      };
    } catch (e: any) {
      return { data: { session: null }, error: normalizeError(e) };
    }
  },

  async getUser() {
    const sessionRes = await auth.getSession();
    if (sessionRes.error) return { data: { user: null }, error: sessionRes.error };
    return { data: { user: sessionRes.data.session?.user || null }, error: null };
  },

  async signOut() {
    try {
      await apiFetch("/api/auth/logout", { method: "POST" });
      setClientToken(null);
      for (const listener of authListeners) listener("SIGNED_OUT", null);
      return { error: null };
    } catch (e: any) {
      return { error: normalizeError(e) };
    }
  },

  onAuthStateChange(callback: AuthListener) {
    authListeners.add(callback);
    return {
      data: {
        subscription: {
          unsubscribe: () => authListeners.delete(callback),
        },
      },
    };
  },
};

const storage = {
  from(bucket: string) {
    return {
      async upload(path: string, file: File, options?: { contentType?: string }) {
        try {
          const form = new FormData();
          form.append("bucket", bucket);
          form.append("path", path);
          form.append("file", file);
          if (options?.contentType) form.append("contentType", options.contentType);

          const res = await apiFetch("/api/upload", {
            method: "POST",
            body: form,
          });
          const json = await res.json().catch(() => ({}));
          if (!res.ok) return { data: null, error: { message: json?.error || "Upload failed" } };
          return {
            data: {
              path: json?.path,
              fullPath: json?.path,
            },
            error: null,
          };
        } catch (e: any) {
          return { data: null, error: normalizeError(e) };
        }
      },

      getPublicUrl(path: string) {
        const base = (process.env.NEXT_PUBLIC_S3_BUCKET_URL || "").replace(/\/+$/, "");
        const cleanPath = String(path || "").replace(/^\/+/, "");
        return {
          data: {
            publicUrl: base ? `${base}/${cleanPath}` : `${API_BASE_URL}/api/storage/download?bucket=${encodeURIComponent(bucket)}&path=${encodeURIComponent(cleanPath)}`,
          },
        };
      },

      async createSignedUrl(path: string, expiresIn = 600) {
        try {
          const res = await apiFetch("/api/storage/signed-url", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ bucket, path, expiresIn }),
          });
          const json = await res.json().catch(() => ({}));
          if (!res.ok) return { data: null, error: { message: json?.error || "Failed to generate signed URL" } };
          return { data: { signedUrl: json?.url }, error: null };
        } catch (e: any) {
          return { data: null, error: normalizeError(e) };
        }
      },

      async download(path: string) {
        try {
          const res = await apiFetch(
            `/api/storage/download?bucket=${encodeURIComponent(bucket)}&path=${encodeURIComponent(path)}`,
            { method: "GET" }
          );
          if (!res.ok) {
            const txt = await res.text().catch(() => "Download failed");
            return { data: null, error: { message: txt || "Download failed" } };
          }
          const blob = await res.blob();
          return { data: blob, error: null };
        } catch (e: any) {
          return { data: null, error: normalizeError(e) };
        }
      },
    };
  },

  async listBuckets() {
    try {
      const res = await apiFetch("/api/storage/buckets", { method: "GET" });
      const json = await res.json().catch(() => ({}));
      if (!res.ok) return { data: null, error: { message: json?.error || "Failed to list buckets" } };
      return { data: json?.buckets || [], error: null };
    } catch (e: any) {
      return { data: null, error: normalizeError(e) };
    }
  },
};

const functions = {
  async invoke(name: string, opts?: { body?: any }) {
    try {
      if (name === "report-docx") {
        const reportId = opts?.body?.reportId;
        if (!reportId) return { data: null, error: { message: "reportId is required" } };
        const res = await apiFetch(`/api/reports/${encodeURIComponent(reportId)}/docx`, {
          method: "GET",
        });
        const json = await res.json().catch(() => ({}));
        if (!res.ok) return { data: null, error: { message: json?.error || "DOCX generation failed" } };
        return { data: json, error: null };
      }
      return { data: null, error: { message: `Unknown function: ${name}` } };
    } catch (e: any) {
      return { data: null, error: normalizeError(e) };
    }
  },
};

export const supabaseCompat = {
  from(table: string) {
    return new QueryBuilder(table);
  },
  auth,
  storage,
  functions,
};
