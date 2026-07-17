"use client";

import React, { useEffect, useState } from "react";
import { usePathname, useRouter } from "next/navigation";

export default function AuthGate({ children }: { children: React.ReactNode }) {
  const router = useRouter();
  const pathname = usePathname();
  const [ready, setReady] = useState(false);

  useEffect(() => {
    let cancelled = false;

    // Public: the login/auth flows, all API routes, the password-gated share
    // landing (/share/...), and the map opened via a share link (?share=...) —
    // external clients reach these without an app login. The query check reads
    // window (client-only, inside the effect) to avoid a useSearchParams()
    // Suspense de-opt on statically rendered pages.
    const publicPaths = ["/login", "/auth", "/api", "/share"];
    const isShareView =
      typeof window !== "undefined" &&
      new URLSearchParams(window.location.search).get("share") != null;
    const isPublic = isShareView || publicPaths.some((p) => pathname?.startsWith(p));

    async function checkAuth() {
      try {
        if (isPublic) {
          if (!cancelled) setReady(true);
          return;
        }

        const token =
          typeof window !== "undefined"
            ? localStorage.getItem("auth_token")
            : null;

        const res = await fetch("/api/auth/me", {
          method: "GET",
          credentials: "include",
          headers: token ? { Authorization: `Bearer ${token}` } : {},
        });

        if (!cancelled && !res.ok) {
          setReady(true);
          router.replace("/login");
          return;
        }

        if (!cancelled) setReady(true);
      } catch (error) {
        console.error("[AuthGate] auth check failed:", error);
        if (!cancelled) {
          setReady(true);
          router.replace("/login");
        }
      }
    }

    checkAuth();

    return () => {
      cancelled = true;
    };
  }, [router, pathname]);

  if (!ready) return null;
  return <>{children}</>;
}
