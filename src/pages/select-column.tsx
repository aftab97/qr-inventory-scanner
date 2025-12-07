import React, { useEffect, useState } from "react";
import { useNavigate } from "react-router-dom";
import { isLocalHost } from "../api";

const API_BASE = isLocalHost ? "http://localhost:3000" : 'http://qr-scanner-api.us-east-1.elasticbeanstalk.com'
const LS_SELECTED_COLUMN = "scan_selected_column";
const LS_COLUMNS_CACHE = "columns_cache_v1";
const CACHE_TTL_MS = 1000 * 60 * 2; // 2 minutes

// Columns to exclude from selection
const EXCLUDE = new Set([
  "id",
  "nom",
  "nombre",
  "pierres",
  "piedras",
  "total",
  "quantité totale",
  "quantité total",
  "quantite totale",
  "quantite total",
  "fotos",
  "foto",
  "pictures",
  "photos",
]);

type ColumnEntry = { norm: string; orig: string };

function getCachedColumns(): { cols: ColumnEntry[]; ts: number } | null {
  try {
    const raw = localStorage.getItem(LS_COLUMNS_CACHE);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed?.cols || !Array.isArray(parsed.cols)) return null;
    return parsed;
  } catch {
    return null;
  }
}

function setCachedColumns(cols: ColumnEntry[]) {
  try {
    localStorage.setItem(LS_COLUMNS_CACHE, JSON.stringify({ cols, ts: Date.now() }));
  } catch {
    // ignore
  }
}

async function fetchColumnsFast(signal?: AbortSignal): Promise<ColumnEntry[]> {
  // If server has a fast endpoint, use it first
  try {
    const res = await fetch(`${API_BASE}/columns`, { signal });
    if (res.ok) {
      const json = await res.json();
      const exclude = EXCLUDE;
      const map = new Map<string, string>(); // norm -> orig
      for (const raw of (json.columns || [])) {
        const orig = String(raw || "").trim();
        if (!orig) continue;
        const norm = orig.toLowerCase();
        if (!norm || exclude.has(norm)) continue;
        if (!map.has(norm)) map.set(norm, orig);
      }
      const out = Array.from(map.entries()).map(([norm, orig]) => ({ norm, orig }));
      out.sort((a, b) => a.orig.localeCompare(b.orig, undefined, { sensitivity: "base" }));
      return out;
    }
  } catch {
    // fall back to client aggregation
  }

  // Fallback: aggregate via /categories + /category/:name with limited concurrency
  const catRes = await fetch(`${API_BASE}/categories`, { signal });
  if (!catRes.ok) throw new Error(await catRes.text());
  const catJson = await catRes.json();
  const cats: string[] = (catJson.categories || []).map((c: any) => String(c.name));

  const exclude = EXCLUDE;
  const map = new Map<string, string>();

  const CONCURRENCY = 6;
  let i = 0;
  async function worker() {
    while (i < cats.length) {
      const idx = i++;
      const cat = cats[idx];
      if (signal?.aborted) return;
      try {
        const res = await fetch(`${API_BASE}/category/${encodeURIComponent(cat)}`, { signal });
        if (!res.ok) continue;
        const json = await res.json();
        const schema = json.schema || null;
        let headers: string[] | null = null;
        if (schema?.headerOriginalOrder?.length) {
          headers = schema.headerOriginalOrder;
        } else if ((json.items || []).length > 0) {
          headers = Object.keys(json.items[0]).filter((k) => k && k !== "category" && k !== "ID");
        }
        if (headers) {
          for (const h of headers) {
            const orig = String(h || "").trim();
            if (!orig) continue;
            const norm = orig.toLowerCase();
            if (!norm || exclude.has(norm)) continue;
            if (!map.has(norm)) map.set(norm, orig);
          }
        }
      } catch {
        // skip errors, continue
      }
    }
  }

  const workers = Array.from({ length: CONCURRENCY }, () => worker());
  await Promise.all(workers);

  const out = Array.from(map.entries()).map(([norm, orig]) => ({ norm, orig }));
  out.sort((a, b) => a.orig.localeCompare(b.orig, undefined, { sensitivity: "base" }));
  return out;
}

export default function SelectColumn(): JSX.Element {
  const nav = useNavigate();
  const [columns, setColumns] = useState<ColumnEntry[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  const [current, setCurrent] = useState<string>(() => {
    try {
      return window.localStorage.getItem(LS_SELECTED_COLUMN) ?? "";
    } catch {
      return "";
    }
  });

  useEffect(() => {
    const ac = new AbortController();

    // 1) Try cache first for instant render
    const cached = getCachedColumns();
    if (cached && Date.now() - cached.ts < CACHE_TTL_MS) {
      setColumns(cached.cols);
      setLoading(false);
      // 2) Background refresh the cache
      fetchColumnsFast(ac.signal)
        .then((fresh) => {
          setColumns(fresh);
          setCachedColumns(fresh);
        })
        .catch(() => {});
      return () => ac.abort();
    }

    // 3) No cache or stale -> fetch and cache
    setLoading(true);
    fetchColumnsFast(ac.signal)
      .then((cols) => {
        setColumns(cols);
        setCachedColumns(cols);
        setError(null);
      })
      .catch((err) => {
        setError("Failed to load columns. Please try again.");
        console.error(err);
      })
      .finally(() => setLoading(false));

    return () => ac.abort();
  }, []);

  const choose = (norm: string) => {
    try {
      window.localStorage.setItem(LS_SELECTED_COLUMN, norm);
    } catch {}
    setCurrent(norm);
    // Return to scan page (history back with fallback)
    try {
      nav(-1);
    } catch {
      nav("/scan", { replace: true });
    }
  };

  return (
    <div className="app-container p-4">
      <header className="flex items-center justify-between mb-4">
        <button
          onClick={() => {
            try { nav(-1); } catch { nav("/scan", { replace: true }); }
          }}
          className="py-2 px-3 rounded bg-gray-100"
        >
          Back
        </button>
        <h1 className="text-lg font-semibold">Choose column to increment</h1>
        <div style={{ width: 56 }} />
      </header>

      {loading ? (
        <div className="p-6 bg-gray-50 rounded text-center">Loading columns…</div>
      ) : error ? (
        <div className="p-4 bg-red-50 text-red-700 rounded">{error}</div>
      ) : columns.length === 0 ? (
        <div className="p-4 bg-yellow-50 text-yellow-800 rounded">No candidate columns found.</div>
      ) : (
        <div className="space-y-3 h-[70vh] overflow-auto">
          {columns.map((c) => {
            const active = current === c.norm;
            return (
              <button
                key={c.norm}
                onClick={() => choose(c.norm)}
                className={`w-full text-left p-5 rounded-lg shadow-sm ${active ? "bg-blue-600 text-white" : "bg-white hover:bg-gray-50"}`}
                style={{ minHeight: 96 }}
                aria-label={`Select column ${c.orig}`}
                type="button"
              >
                <div className="flex items-center justify-between">
                  <div>
                    <div className="text-lg font-medium">{c.orig}</div>
                    <div className="text-xs text-gray-500 mt-1">key: {c.norm}</div>
                  </div>
                  <div className="text-sm text-gray-400">{active ? "Selected" : "Tap to select"}</div>
                </div>
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}