"use client";

import Link from "next/link";
import { useEffect, useState } from "react";
import { supabase } from "../../lib/supabase";
import type { Report } from "@/types/db";

export default function ReportsPage() {
  const [reports, setReports] = useState<Report[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    (async () => {
      setLoading(true);

      const { data, error } = await supabase
        .from("reports")
        .select("*")
        .order("created_at", { ascending: false });

      if (!error && data) setReports(data as Report[]);
      setLoading(false);
    })();
  }, []);

  return (
    <div style={{ padding: 20, fontFamily: "system-ui" }}>
      <h2>Reports</h2>

      {loading ? <p>Loading...</p> : null}

      <div style={{ display: "grid", gap: 10 }}>
        {reports.map((r) => (
          <Link
            key={r.id}
            href={`/reports/${r.id}`}
            style={{
              border: "1px solid #ddd",
              padding: 12,
              borderRadius: 10,
              textDecoration: "none",
              color: "inherit",
            }}
          >
            <div style={{ fontWeight: 700 }}>{r.category || "Report"}</div>
            <div style={{ fontSize: 12, color: "#666" }}>
              {new Date(r.created_at).toLocaleString()}
            </div>
            <div style={{ marginTop: 6 }}>
              {(r.description || "").slice(0, 120)}
            </div>
          </Link>
        ))}
      </div>
    </div>
  );
}
