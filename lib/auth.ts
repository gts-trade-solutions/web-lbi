import jwt, { JwtPayload } from "jsonwebtoken";

export const AUTH_COOKIE_NAME = "auth_token";

type SafeUser = {
  id: string;
  email: string;
  name?: string | null;
  role?: string | null;
};

export type AuthTokenPayload = JwtPayload & {
  id: string;
  email: string;
  name?: string | null;
  role?: string | null;
};

function cookieOpts(maxAgeSeconds: number) {
  return [
    "Path=/",
    "HttpOnly",
    "SameSite=Lax",
    process.env.NODE_ENV === "production" ? "Secure" : "",
    `Max-Age=${Math.max(0, maxAgeSeconds)}`,
  ]
    .filter(Boolean)
    .join("; ");
}

function parseExpiresToSeconds(value: string) {
  const v = value.trim();
  if (/^\d+$/.test(v)) return Number(v);
  const m = v.match(/^(\d+)([smhd])$/i);
  if (!m) return 7 * 24 * 60 * 60;
  const n = Number(m[1]);
  const unit = m[2].toLowerCase();
  if (unit === "s") return n;
  if (unit === "m") return n * 60;
  if (unit === "h") return n * 60 * 60;
  return n * 24 * 60 * 60;
}

function getJwtSecret() {
  const secret = process.env.JWT_SECRET_KEY;
  if (!secret) throw new Error("JWT_SECRET_KEY is not configured");
  return secret;
}

function getTokenExpiry() {
  return process.env.AUTH_TOKEN_EXPIRES_IN || "7d";
}

export function signToken(user: SafeUser) {
  const payload: AuthTokenPayload = {
    id: user.id,
    email: user.email,
    name: user.name || null,
    role: user.role || null,
  };
  return jwt.sign(payload, getJwtSecret(), {
    expiresIn: getTokenExpiry() as jwt.SignOptions["expiresIn"],
  });
}

export function verifyToken(token: string): AuthTokenPayload | null {
  try {
    return jwt.verify(token, getJwtSecret()) as AuthTokenPayload;
  } catch {
    return null;
  }
}

function parseCookieHeader(cookieHeader: string | null) {
  const out: Record<string, string> = {};
  if (!cookieHeader) return out;
  for (const part of cookieHeader.split(";")) {
    const [k, ...rest] = part.trim().split("=");
    if (!k) continue;
    out[k] = decodeURIComponent(rest.join("=") || "");
  }
  return out;
}

export function extractTokenFromRequest(request: Request) {
  const xAuthToken = request.headers.get("x-auth-token");
  if (xAuthToken && xAuthToken.trim()) return xAuthToken.trim();

  const auth = request.headers.get("authorization") || "";
  if (auth.toLowerCase().startsWith("bearer ")) {
    return auth.slice(7).trim();
  }

  const cookieMap = parseCookieHeader(request.headers.get("cookie"));
  return cookieMap[AUTH_COOKIE_NAME] || null;
}

export function getAuthUser(request: Request) {
  const token = extractTokenFromRequest(request);
  if (!token) return null;
  return verifyToken(token);
}

export function requireAuth(request: Request) {
  const user = getAuthUser(request);
  if (!user) throw new Error("Unauthorized");
  return user;
}

export function makeAuthCookieHeader(token: string) {
  const maxAge = parseExpiresToSeconds(getTokenExpiry());
  return `${AUTH_COOKIE_NAME}=${encodeURIComponent(token)}; ${cookieOpts(maxAge)}`;
}

export function clearAuthCookieHeader() {
  return `${AUTH_COOKIE_NAME}=; ${cookieOpts(0)}`;
}
