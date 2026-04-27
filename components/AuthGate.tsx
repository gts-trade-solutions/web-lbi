"use client";

import React, { useEffect, useState } from "react";
import { usePathname, useRouter } from "next/navigation";

export default function AuthGate({ children }: { children: React.ReactNode }) {
  const router = useRouter();
  const pathname = usePathname();
  const [ready, setReady] = useState(false);

  useEffect(() => {
    let cancelled = false;

    const publicPaths = ["/login", "/auth", "/api"];
    const isPublic = publicPaths.some((p) => pathname?.startsWith(p));

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
