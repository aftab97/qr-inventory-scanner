import React, { useEffect, useMemo, useState } from "react";
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  PointElement,
  BarElement,
  ArcElement,
  Title,
  Tooltip,
  Legend,
} from "chart.js";
import { Scatter } from "react-chartjs-2";
import { Bar } from "react-chartjs-2";
import { Doughnut } from "react-chartjs-2";
import { Bubble } from "react-chartjs-2";
import { isLocalHost } from "../api";

ChartJS.register(
  CategoryScale,
  LinearScale,
  PointElement,
  BarElement,
  ArcElement,
  Title,
  Tooltip,
  Legend
);

const API_BASE = isLocalHost ? "http://localhost:3000" : 'http://qr-scanner-api.us-east-1.elasticbeanstalk.com/'

function Spinner({ className = "h-4 w-4 text-white" }) {
  return (
    <svg className={`animate-spin ${className}`} viewBox="0 0 24 24" fill="none" stroke="currentColor" aria-hidden>
      <circle className="opacity-25" cx="12" cy="12" r="10" strokeWidth="3" />
      <path className="opacity-75" d="M4 12a8 8 0 018-8" strokeWidth="3" strokeLinecap="round" />
    </svg>
  );
}

function computeTotal(item, headerNorm) {
  const excluded = new Set([
    "id",
    "nombre",
    "nom",
    "fotos",
    "pierres",
    "piedras",
    "pictures",
    "photos",
    "total",
    "quantité totale",
    "quantité total",
  ]);
  let sum = 0;
  for (const key of headerNorm) {
    const k = String(key || "").toLowerCase();
    if (excluded.has(k) || k === "total") continue;
    const raw = item[k];
    if (raw === undefined || raw === null || raw === "") continue;
    const n = Number(String(raw).replace(",", "."));
    if (Number.isFinite(n)) sum += n;
  }
  return sum;
}

export default function ViewerWithCharts() {
  const [categories, setCategories] = useState([]);
  const [activeCategory, setActiveCategory] = useState(null);
  const [schema, setSchema] = useState(null);
  const [items, setItems] = useState([]);
  const [loadingCats, setLoadingCats] = useState(false);
  const [loadingItems, setLoadingItems] = useState(false);
  const [error, setError] = useState(null);

  // UI state
  const [chartType, setChartType] = useState("scatter"); // scatter | bar | bubble | donut
  const [topN, setTopN] = useState(100);
  const [sortBy, setSortBy] = useState("total_desc"); // total_desc, total_asc, label

  // New: value column selection & filters
  // 'total' means use client-side computed total (default)
  const [valueColumn, setValueColumn] = useState("total");
  const [minFilter, setMinFilter] = useState("");
  const [maxFilter, setMaxFilter] = useState("");

  // load categories
  useEffect(() => {
    let cancelled = false;
    async function loadCats() {
      setLoadingCats(true);
      try {
        const res = await fetch(`${API_BASE}/categories`);
        if (!res.ok) throw new Error(await res.text());
        const json = await res.json();
        if (cancelled) return;
        setCategories(json.categories || []);
        if (json.categories && json.categories.length > 0 && !activeCategory) {
          setActiveCategory(json.categories[0].name);
        }
      } catch (err) {
        console.error("load categories", err);
        setError("Failed to load categories");
      } finally {
        if (!cancelled) setLoadingCats(false);
      }
    }
    loadCats();
    return () => {
      cancelled = true;
    };
  }, []); // on mount

  // load items for activeCategory
  useEffect(() => {
    if (!activeCategory) {
      setItems([]);
      setSchema(null);
      return;
    }
    let cancelled = false;
    async function loadItems() {
      setLoadingItems(true);
      setError(null);
      try {
        const res = await fetch(`${API_BASE}/category/${encodeURIComponent(activeCategory)}`);
        if (!res.ok) throw new Error(await res.text());
        const json = await res.json();
        if (cancelled) return;
        setSchema(json.schema || null);

        const headerNorm =
          (json.schema && json.schema.headerNormalizedOrder) ||
          (json.schema && json.schema.headerOriginalOrder && json.schema.headerOriginalOrder.map((h) => String(h).toLowerCase())) ||
          [];

        const itemsLoaded = (json.items || []).map((it) => {
          const normalized = {};
          Object.keys(it || {}).forEach((k) => {
            normalized[String(k).toLowerCase()] = it[k];
          });
          const normList = headerNorm.length ? headerNorm : Object.keys(normalized);
          const total = computeTotal(normalized, normList);
          normalized._total = total;
          return normalized;
        });
        setItems(itemsLoaded);
      } catch (err) {
        console.error("load category items", err);
        setError("Failed to load category items");
        setItems([]);
        setSchema(null);
      } finally {
        if (!cancelled) setLoadingItems(false);
      }
    }
    loadItems();
    return () => {
      cancelled = true;
    };
  }, [activeCategory]);

  // derive available numeric columns for filtering/selection (exclude unwanted names)
  const availableColumns = useMemo(() => {
    // column names from schema.headerNormalizedOrder or derived from items
    const origList = (schema && schema.headerOriginalOrder) || [];
    const normList = (schema && schema.headerNormalizedOrder) || origList.map((h) => String(h).toLowerCase());
    const excludeNames = new Set(["total", "quantité totale", "quantité total", "quantite totale", "quantite total"]);
    // also exclude earlier protected names
    ["id", "nombre", "nom", "fotos", "pierres", "piedras", "pictures", "photos"].forEach((n) => excludeNames.add(n));

    const candidates = [];
    for (let i = 0; i < normList.length; i++) {
      const norm = String(normList[i] || "").toLowerCase();
      const orig = origList[i] || norm;
      if (!norm) continue;
      if (excludeNames.has(norm)) continue;
      // check if column contains numeric values in items (at least one numeric)
      let numericFound = false;
      for (const it of items) {
        const v = it[norm];
        if (v === undefined || v === null || v === "") continue;
        const n = Number(String(v).replace(",", "."));
        if (Number.isFinite(n)) {
          numericFound = true;
          break;
        }
      }
      if (numericFound) candidates.push({ norm, orig });
    }
    return candidates;
  }, [schema, items]);

  // Ensure valueColumn defaults to 'total' or first available numeric column
  useEffect(() => {
    if (valueColumn === "total") return;
    // if current valueColumn is not available (e.g., schema changed), reset to total
    const ok = valueColumn === "total" || availableColumns.some((c) => c.norm === valueColumn);
    if (!ok) {
      setValueColumn("total");
    }
  }, [availableColumns, valueColumn]);

  // Prepare chart data derived from items and applied filters
  const prepared = useMemo(() => {
    const headerNorm = schema && schema.headerNormalizedOrder ? schema.headerNormalizedOrder : [];
    const parsedMin = minFilter === "" ? null : Number(minFilter);
    const parsedMax = maxFilter === "" ? null : Number(maxFilter);

    const rows = items
      .map((it, idx) => {
        const idOrig = it.id ?? it.ID ?? String(idx + 1);
        const nom = it.nom ?? it.nombre ?? "";
        const pierres = it.pierres ?? it.piedras ?? "";
        let label = "";
        if (nom && pierres) label = `${nom} - ${pierres}`;
        else if (nom) label = `${nom}`;
        else if (pierres) label = `${pierres}`;
        else label = String(idOrig);

        const total = typeof it._total === "number" ? it._total : computeTotal(it, headerNorm);

        // value for chart derived from selected valueColumn
        let value;
        if (valueColumn === "total") value = Number(total);
        else {
          const raw = it[valueColumn];
          value = raw === undefined || raw === null || raw === "" ? NaN : Number(String(raw).replace(",", "."));
        }

        return {
          id: String(idOrig),
          label,
          nom,
          pierres,
          total: Number(total),
          value,
          raw: it,
        };
      })
      // filter out items where selected value is NaN (when not using total)
      .filter((r) => {
        if (valueColumn === "total") {
          // total always numeric
          return true;
        }
        // allow inclusion if value is finite
        return Number.isFinite(r.value);
      })
      // apply range filters if provided
      .filter((r) => {
        if (parsedMin !== null && r.value < parsedMin) return false;
        if (parsedMax !== null && r.value > parsedMax) return false;
        return true;
      });

    let sorted = [...rows];
    if (sortBy === "total_desc") sorted.sort((a, b) => b.total - a.total);
    else if (sortBy === "total_asc") sorted.sort((a, b) => a.total - b.total);
    else if (sortBy === "label") sorted.sort((a, b) => (a.label > b.label ? 1 : a.label < b.label ? -1 : 0));

    if (typeof topN === "number" && topN > 0) sorted = sorted.slice(0, topN);

    return { headerNorm, rows: sorted };
  }, [items, schema, sortBy, topN, valueColumn, minFilter, maxFilter]);

  // Chart builders using prepared.rows and chosen value
  const buildScatterData = () => {
    const dataPoints = prepared.rows.map((r, i) => ({
      x: i + 1,
      y: valueColumn === "total" ? r.total : r.value,
      meta: r,
    }));
    return {
      datasets: [
        {
          label: `Value (${valueColumn === "total" ? "Total" : valueColumn}) per item`,
          data: dataPoints,
          backgroundColor: "rgba(59,130,246,0.85)",
          pointRadius: 6,
        },
      ],
    };
  };

  const buildBarData = () => {
    const labels = prepared.rows.map((r) => r.label);
    const data = prepared.rows.map((r) => (valueColumn === "total" ? r.total : r.value));
    return {
      labels,
      datasets: [
        {
          label: valueColumn === "total" ? "Total" : valueColumn,
          data,
          backgroundColor: prepared.rows.map(() => "rgba(99,102,241,0.85)"),
        },
      ],
    };
  };

  const buildBubbleData = () => {
    const data = prepared.rows.map((r, i) => {
      const val = valueColumn === "total" ? r.total : r.value;
      const rSize = Math.max(4, Math.min(40, Math.sqrt(Math.abs(val || 0))));
      return { x: i + 1, y: val, r: rSize, meta: r };
    });
    return {
      datasets: [
        {
          label: valueColumn === "total" ? "Totals (bubble)" : `${valueColumn} (bubble)`,
          data,
          backgroundColor: "rgba(16,185,129,0.85)",
        },
      ],
    };
  };

  const buildDonutData = () => {
    const labels = prepared.rows.map((r) => r.label);
    const data = prepared.rows.map((r) => (valueColumn === "total" ? r.total : r.value));
    const colors = prepared.rows.map((_, i) => `hsl(${(i * 47) % 360} 70% 50% / 0.85)`);
    return {
      labels,
      datasets: [
        {
          data,
          backgroundColor: colors,
          hoverOffset: 6,
        },
      ],
    };
  };

  // Tooltip callbacks adjusted to include nom/pierres and show chosen value column
  const scatterOptions = {
    responsive: true,
    plugins: {
      legend: { display: false },
      tooltip: {
        callbacks: {
          title: (ctx) => {
            const raw = ctx[0]?.raw;
            const meta = raw?.meta ?? null;
            return meta ? meta.label : "";
          },
          label: (ctx) => {
            const raw = ctx.raw;
            const meta = raw?.meta ?? null;
            if (meta) {
              const val = valueColumn === "total" ? meta.total : meta.value;
              const lines = [`${valueColumn === "total" ? "Total" : valueColumn}: ${val}`];
              if (meta.nom) lines.push(`Nom: ${meta.nom}`);
              if (meta.pierres) lines.push(`Pierres: ${meta.pierres}`);
              return lines;
            }
            return `Value: ${ctx.raw?.y ?? ctx.raw}`;
          },
        },
      },
      title: { display: true, text: `Scatter: ${valueColumn === "total" ? "Total" : valueColumn} per item (${prepared.rows.length})` },
    },
    scales: {
      x: { title: { display: true, text: "Item (ordinal index)" } },
      y: { title: { display: true, text: valueColumn === "total" ? "Total" : valueColumn } },
    },
  };

  const barOptions = {
    responsive: true,
    plugins: {
      legend: { display: false },
      tooltip: {
        callbacks: {
          title: (ctx) => ctx[0]?.label ?? "",
          label: (ctx) => {
            const idx = ctx.dataIndex;
            const meta = prepared.rows[idx];
            const val = valueColumn === "total" ? meta.total : meta.value;
            const lines = [`${valueColumn === "total" ? "Total" : valueColumn}: ${val}`];
            if (meta.nom) lines.push(`Nom: ${meta.nom}`);
            if (meta.pierres) lines.push(`Pierres: ${meta.pierres}`);
            return lines;
          },
        },
      },
      title: { display: true, text: `Bar: ${valueColumn === "total" ? "Total" : valueColumn} per item (${prepared.rows.length})` },
    },
    scales: {
      x: { title: { display: true, text: "Item (nom - pierres)" }, ticks: { maxRotation: 45, minRotation: 0 } },
      y: { title: { display: true, text: valueColumn === "total" ? "Total" : valueColumn } },
    },
  };

  const bubbleOptions = {
    responsive: true,
    plugins: {
      legend: { display: false },
      tooltip: {
        callbacks: {
          title: (ctx) => {
            const raw = ctx[0]?.raw;
            const meta = raw?.meta ?? null;
            return meta ? meta.label : "";
          },
          label: (ctx) => {
            const raw = ctx.raw;
            const meta = raw?.meta ?? null;
            if (meta) {
              const val = valueColumn === "total" ? meta.total : meta.value;
              const lines = [`${valueColumn === "total" ? "Total" : valueColumn}: ${val}`, `Size ≈ ${Math.round(Math.sqrt(Math.abs(val || 0)))}`];
              if (meta.nom) lines.push(`Nom: ${meta.nom}`);
              if (meta.pierres) lines.push(`Pierres: ${meta.pierres}`);
              return lines;
            }
            return `Value: ${ctx.raw?.y ?? ctx.raw}`;
          },
        },
      },
      title: { display: true, text: `Bubble: ${valueColumn === "total" ? "Total" : valueColumn} (size ≈ value)` },
    },
    scales: {
      x: { title: { display: true, text: "Item (ordinal index)" } },
      y: { title: { display: true, text: valueColumn === "total" ? "Total" : valueColumn } },
    },
  };

  const donutOptions = {
    responsive: true,
    plugins: {
      legend: { position: "right" },
      tooltip: {
        callbacks: {
          label: (ctx) => {
            const idx = ctx.dataIndex;
            const meta = prepared.rows[idx];
            const val = valueColumn === "total" ? meta.total : meta.value;
            const lines = [`${valueColumn === "total" ? "Total" : valueColumn}: ${val}`];
            if (meta.nom) lines.push(`Nom: ${meta.nom}`);
            if (meta.pierres) lines.push(`Pierres: ${meta.pierres}`);
            return lines;
          },
        },
      },
      title: { display: true, text: `Donut: ${valueColumn === "total" ? "Total" : valueColumn} share (${prepared.rows.length} items)` },
    },
  };

  return (
    <div className="p-4 max-w-7xl mx-auto">
      <header className="mb-4 flex items-center justify-between">
        <h2 className="text-xl font-semibold">Inventory Charts — {activeCategory ?? "No category"}</h2>

        <div className="flex items-center gap-3">
          <select value={chartType} onChange={(e) => setChartType(e.target.value)} className="p-1 border rounded">
            <option value="scatter">Scatter</option>
            <option value="bar">Bar</option>
            <option value="bubble">Bubble</option>
            <option value="donut">Donut</option>
          </select>

          <select value={valueColumn} onChange={(e) => setValueColumn(e.target.value)} className="p-1 border rounded">
            <option value="total">Client-side Total</option>
            {availableColumns.map((c) => (
              <option key={c.norm} value={c.norm}>
                {c.orig}
              </option>
            ))}
          </select>
        </div>
      </header>

      <div className="flex gap-4">
        <aside className="w-64">
          <div className="mb-4">
            <h4 className="text-sm font-medium mb-2">Categories</h4>
            {loadingCats ? (
              <div className="text-sm text-gray-500">
                <Spinner className="h-4 w-4 text-gray-600" /> Loading...
              </div>
            ) : (
              <div className="flex flex-col gap-2">
                {categories.map((c) => (
                  <button
                    key={c.name}
                    onClick={() => setActiveCategory(c.name)}
                    className={`text-left px-2 py-1 rounded ${activeCategory === c.name ? "bg-blue-600 text-white" : "hover:bg-gray-100"}`}
                  >
                    {c.name}
                  </button>
                ))}
              </div>
            )}
          </div>

          <div className="mb-4">
            <h4 className="text-sm font-medium mb-2">Chart Controls</h4>
            <div className="space-y-2 text-sm">
              <label className="block">Top N items</label>
              <input type="number" value={topN} min={1} onChange={(e) => setTopN(Number(e.target.value) || 10)} className="w-full p-1 border rounded" />

              <label className="block mt-2">Sort</label>
              <select value={sortBy} onChange={(e) => setSortBy(e.target.value)} className="w-full p-1 border rounded">
                <option value="total_desc">Total — descending</option>
                <option value="total_asc">Total — ascending</option>
                <option value="label">Label (alpha)</option>
              </select>

              <div className="mt-3">
                <label className="block text-sm font-medium">Value filters (optional)</label>
                <input type="text" placeholder="min" value={minFilter} onChange={(e) => setMinFilter(e.target.value)} className="w-full p-1 border rounded mt-1" />
                <input type="text" placeholder="max" value={maxFilter} onChange={(e) => setMaxFilter(e.target.value)} className="w-full p-1 border rounded mt-2" />
                <div className="mt-2 text-xs text-gray-500">Leave empty to disable min/max. Filters apply to selected value column.</div>
              </div>
            </div>
          </div>

          <div>
            <h4 className="text-sm font-medium mb-2">Status</h4>
            <div className="text-sm text-gray-600">
              <div>Active: <span className="font-medium">{activeCategory ?? "—"}</span></div>
              <div>Displayed items: <span className="font-medium">{prepared.rows.length}</span></div>
              <div>Loading items: <span className="font-medium">{loadingItems ? "Yes" : "No"}</span></div>
            </div>
          </div>
        </aside>

        <main className="flex-1">
          {loadingItems ? (
            <div className="p-6 bg-gray-50 rounded text-center">
              <Spinner className="h-6 w-6 text-gray-600" /> Loading items…
            </div>
          ) : (
            <div className="p-2 border rounded">
              {chartType === "scatter" && <Scatter data={buildScatterData()} options={scatterOptions} />}
              {chartType === "bar" && <Bar data={buildBarData()} options={barOptions} />}
              {chartType === "bubble" && <Bubble data={buildBubbleData()} options={bubbleOptions} />}
              {chartType === "donut" && <Doughnut data={buildDonutData()} options={donutOptions} />}
            </div>
          )}
        </main>
      </div>
    </div>
  );
}