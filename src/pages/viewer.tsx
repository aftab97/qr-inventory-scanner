import React, { useEffect, useRef, useState } from "react";
import { Trash2 } from "lucide-react";
import * as XLSX from "xlsx";
import { isLocalHost } from "../api";

const API_BASE = isLocalHost ? "http://localhost:3000" : 'http://qr-scanner-api.us-east-1.elasticbeanstalk.com/'

// LocalStorage keys
const LS_ROWS_KEY = "qr_viewer_row_selection_v1";
const LS_COLS_KEY = "qr_viewer_col_selection_v1";

/**
 * Viewer.jsx
 *
 * Full, self-contained Viewer component:
 * - Loads categories and category items from backend (GET /categories, GET /category/:name).
 * - Per-category column selection (persisted in localStorage).
 * - Per-category row selection (persisted in localStorage, merged on load so user choices survive navigation).
 * - Client-side "Total" column and optional totals row for exports.
 * - Per-column delete (protected names) with server endpoints (DELETE /category/:name/column).
 * - Per-row delete (DELETE /item/:id) with confirmation modal; first data row protected.
 * - Exports (single category and all) honoring selected columns, selected rows, and include-totals option.
 * - Hover effects and small inline spinners for long operations.
 *
 * Tailwind CSS classes are used for styling; replace if you don't use Tailwind.
 */

function Spinner({ className = "h-4 w-4 text-white" }) {
  return (
    <svg className={`animate-spin ${className}`} viewBox="0 0 24 24" fill="none" stroke="currentColor">
      <circle className="opacity-25" cx="12" cy="12" r="10" strokeWidth="3"></circle>
      <path className="opacity-75" d="M4 12a8 8 0 018-8" strokeWidth="3" strokeLinecap="round"></path>
    </svg>
  );
}

export default function Viewer() {
  // Data
  const [categories, setCategories] = useState([]);
  const [active, setActive] = useState(null);
  const [items, setItems] = useState([]);
  const [schema, setSchema] = useState(null);

  // UI state / flags
  const [loadingCats, setLoadingCats] = useState(false);
  const [loadingItems, setLoadingItems] = useState(false);
  const [replacing, setReplacing] = useState(false);
  const [deletingColumn, setDeletingColumn] = useState(false);
  const [deletingRows, setDeletingRows] = useState(false);
  const [deletingCategory, setDeletingCategory] = useState(false);
  const [exportingCategory, setExportingCategory] = useState(false);
  const [exportingAll, setExportingAll] = useState(false);
  const [error, setError] = useState(null);
  const [modalError, setModalError] = useState(null);

  // Editing state
  const [editingCell, setEditingCell] = useState(null); // { id, normKey, original }
  const [cellValue, setCellValue] = useState("");
  const [savingCellKey, setSavingCellKey] = useState(null);

  // File/replace
  const fileInputRef = useRef(null);
  const [pendingReplaceFile, setPendingReplaceFile] = useState(null);
  const [confirmReplaceOpen, setConfirmReplaceOpen] = useState(false);

  // Delete column modal
  const [pendingDeleteColumn, setPendingDeleteColumn] = useState(null); // { columnOrig, columnNorm }
  // Delete rows modal
  const [pendingDeleteRowIds, setPendingDeleteRowIds] = useState([]);
  const [confirmDeleteRowsOpen, setConfirmDeleteRowsOpen] = useState(false);
  // Delete category modal
  const [pendingDeleteCategory, setPendingDeleteCategory] = useState(null);

  // Per-category selections (persisted to localStorage)
  const [columnSelectionByCategory, setColumnSelectionByCategory] = useState(() => {
    try {
      const raw = window.localStorage.getItem(LS_COLS_KEY);
      return raw ? JSON.parse(raw) : {};
    } catch {
      return {};
    }
  });
  const [rowSelectionByCategory, setRowSelectionByCategory] = useState(() => {
    try {
      const raw = window.localStorage.getItem(LS_ROWS_KEY);
      return raw ? JSON.parse(raw) : {};
    } catch {
      return {};
    }
  });

  // Export options per category (in-memory only)
  const [exportSelectedRowsOnlyByCategory, setExportSelectedRowsOnlyByCategory] = useState({});
  const [includeTotalsRowByCategory, setIncludeTotalsRowByCategory] = useState({});

  // Persist selections when they change
  useEffect(() => {
    try {
      window.localStorage.setItem(LS_COLS_KEY, JSON.stringify(columnSelectionByCategory));
    } catch {}
  }, [columnSelectionByCategory]);
  useEffect(() => {
    try {
      window.localStorage.setItem(LS_ROWS_KEY, JSON.stringify(rowSelectionByCategory));
    } catch {}
  }, [rowSelectionByCategory]);

  // Load categories on mount
  useEffect(() => { loadCategories(); /* eslint-disable-next-line */ }, []);

  // Load items whenever active changes
  useEffect(() => { if (active) loadCategoryItems(active); /* eslint-disable-next-line */ }, [active]);

  // ----------------- Backend calls -----------------
  async function loadCategories() {
    setLoadingCats(true);
    setError(null);
    try {
      const res = await fetch(`${API_BASE}/categories`);
      if (!res.ok) throw new Error(await res.text());
      const json = await res.json();
      setCategories(json.categories || []);
      if (json.categories && json.categories.length > 0) {
        if (!active || !json.categories.find((c) => c.name === active)) setActive(json.categories[0].name);
      } else {
        setActive(null);
      }
    } catch (err) {
      console.error("loadCategories", err);
      setError((err && err.message) || "Failed to load categories");
    } finally {
      setLoadingCats(false);
    }
  }

  async function loadCategoryItems(category) {
    setLoadingItems(true);
    setItems([]);
    setSchema(null);
    setError(null);
    try {
      const res = await fetch(`${API_BASE}/category/${encodeURIComponent(category)}`);
      if (!res.ok) throw new Error(await res.text());
      const json = await res.json();
      const loadedItems = json.items || [];
      setSchema(json.schema || null);
      setItems(loadedItems);
      initializeSelectionForCategory(category, json.schema || null, loadedItems);
      // ensure export options defaults exist
      setExportSelectedRowsOnlyByCategory((prev) => (prev && typeof prev[category] !== "undefined") ? prev : { ...(prev || {}), [category]: false });
      setIncludeTotalsRowByCategory((prev) => (prev && typeof prev[category] !== "undefined") ? prev : { ...(prev || {}), [category]: true });
    } catch (err) {
      console.error("loadCategoryItems", err);
      setError((err && err.message) || "Failed to load items");
    } finally {
      setLoadingItems(false);
    }
  }

  // ----------------- Header helpers -----------------
  function headerOriginalOrder() {
    if (schema && schema.headerOriginalOrder && schema.headerOriginalOrder.length) return schema.headerOriginalOrder;
    if (!items || items.length === 0) return [];
    return Object.keys(items[0]).filter((k) => k !== "category" && k !== "ID");
  }
  function headerNormalizedOrder() {
    if (schema && schema.headerNormalizedOrder && schema.headerNormalizedOrder.length) return schema.headerNormalizedOrder;
    return headerOriginalOrder().map((h) => String(h).toLowerCase());
  }

  // ----------------- Initialize selections (merge with saved) -----------------
  function initializeSelectionForCategory(category, schemaObj, itemsArr) {
    // Columns: if no saved columns for this category, initialize default (select all)
    setColumnSelectionByCategory((prev) => {
      if (prev && prev[category]) return prev;
      const next = { ...(prev || {}) };
      const orig = (schemaObj && Array.isArray(schemaObj.headerOriginalOrder) && schemaObj.headerOriginalOrder.length)
        ? schemaObj.headerOriginalOrder.slice()
        : (itemsArr && itemsArr.length ? Object.keys(itemsArr[0]).filter((k) => k !== "category" && k !== "ID") : []);
      const normalized = orig.map((h) => String(h).toLowerCase());
      const map = { id: true, total: true };
      normalized.forEach((n) => (map[n] = true));
      next[category] = map;
      return next;
    });

    // Rows: merge saved row selection map with current item IDs
    setRowSelectionByCategory((prev) => {
      const existing = (prev && prev[category]) || {};
      const next = { ...(prev || {}) };

      const currentIds = new Set((itemsArr || []).map((it) => it.ID));

      // keep existing entries for IDs still present
      const merged = {};
      for (const id of Object.keys(existing)) {
        if (currentIds.has(id)) merged[id] = !!existing[id];
      }
      // for any current ID not in merged, default to true (selected)
      for (const it of itemsArr || []) {
        if (merged[it.ID] === undefined) merged[it.ID] = true;
      }
      next[category] = merged;
      return next;
    });
  }

  // ----------------- Column toggles -----------------
  function toggleColumnForActive(normKey) {
    if (!active) return;
    setColumnSelectionByCategory((prev) => {
      const prevMap = (prev && prev[active]) || {};
      const next = { ...(prev || {}) };
      next[active] = { ...prevMap, [normKey]: !prevMap[normKey] };
      return next;
    });
  }

  // ----------------- Row selection helpers -----------------
  function isRowSelected(category, id) {
    const map = (rowSelectionByCategory && rowSelectionByCategory[category]) || {};
    // map should exist due to initialization, but default to true if missing
    if (!map) return true;
    return !!map[id];
  }

  function toggleRowSelection(id) {
    if (!active) return;
    setRowSelectionByCategory((prev) => {
      const prevMap = (prev && prev[active]) || {};
      const next = { ...(prev || {}) };
      next[active] = { ...prevMap, [id]: !prevMap[id] };
      return next;
    });
  }

  function setAllRowsSelectedForActive(selectAll) {
    if (!active) return;
    setRowSelectionByCategory((prev) => {
      const next = { ...(prev || {}) };
      const map = {};
      for (const it of items || []) map[it.ID] = !!selectAll;
      next[active] = map;
      return next;
    });
  }

  function areAllRowsSelected() {
    if (!active) return false;
    const map = (rowSelectionByCategory && rowSelectionByCategory[active]) || {};
    if (!items || items.length === 0) return false;
    return items.every((it) => !!map[it.ID]);
  }

  function anyRowSelected() {
    if (!active) return false;
    const map = (rowSelectionByCategory && rowSelectionByCategory[active]) || {};
    return (items || []).some((it) => !!map[it.ID]);
  }

  // ----------------- Delete rows -----------------
  function onRequestDeleteSelectedRows() {
    if (!active) return;
    const map = (rowSelectionByCategory && rowSelectionByCategory[active]) || {};
    const ids = (items || []).filter((it) => !!map[it.ID]).map((it) => it.ID);
    if (!ids.length) {
      setModalError({ title: "No rows selected", message: "Select rows to delete first." });
      return;
    }
    const firstId = items[0] ? items[0].ID : null;
    if (ids.includes(firstId)) {
      setModalError({ title: "Cannot delete first row", message: "The first data row is protected and cannot be deleted." });
      return;
    }
    setPendingDeleteRowIds(ids);
    setConfirmDeleteRowsOpen(true);
  }

  async function confirmDeleteSelectedRows() {
    if (!active || !pendingDeleteRowIds.length) {
      setPendingDeleteRowIds([]); setConfirmDeleteRowsOpen(false); return;
    }
    setDeletingRows(true);
    try {
      // Send DELETE per-id; for many rows prefer adding a server batch-delete endpoint
      const promises = pendingDeleteRowIds.map((id) => fetch(`${API_BASE}/item/${encodeURIComponent(id)}`, { method: "DELETE" }));
      const responses = await Promise.all(promises);
      const failed = [];
      for (let i = 0; i < responses.length; i++) {
        if (!responses[i].ok) {
          const txt = await responses[i].text().catch(() => "");
          failed.push({ id: pendingDeleteRowIds[i], text: txt, status: responses[i].status });
        }
      }
      if (failed.length) {
        console.error("some deletes failed", failed);
        setModalError({ title: "Delete incomplete", message: `Failed to delete ${failed.length} rows. See console.` });
      }
      // reload items (this also reinitializes/merges row selections)
      await loadCategoryItems(active);
      setPendingDeleteRowIds([]); setConfirmDeleteRowsOpen(false);
    } catch (err) {
      console.error("confirmDeleteSelectedRows", err);
      setModalError({ title: "Delete failed", message: (err && err.message) || "Unable to delete rows." });
    } finally {
      setDeletingRows(false);
    }
  }

  // ----------------- Column delete -----------------
  function isProtectedColumn(orig) {
    if (!orig) return false;
    const n = String(orig).trim().toLowerCase();
    return n === "nom" || n === "nombre" || n === "pierres" || n === "piedras";
  }

  async function confirmDeleteColumn() {
    if (!active || !pendingDeleteColumn) { setPendingDeleteColumn(null); return; }
    const col = pendingDeleteColumn.columnOrig;
    if (isProtectedColumn(col)) {
      setModalError({ title: "Protected column", message: `The column "${col}" cannot be deleted.` });
      setPendingDeleteColumn(null);
      return;
    }
    setDeletingColumn(true);
    try {
      let res = await fetch(`${API_BASE}/category/${encodeURIComponent(active)}/column/${encodeURIComponent(col)}`, { method: "DELETE" });
      if (!res.ok) {
        res = await fetch(`${API_BASE}/category/${encodeURIComponent(active)}/column`, {
          method: "DELETE",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ column: col }),
        });
      }
      if (!res.ok) {
        const txt = await res.text().catch(() => "");
        throw new Error(txt || `Server responded ${res.status}`);
      }
      await loadCategoryItems(active);
      setPendingDeleteColumn(null);
    } catch (err) {
      console.error("confirmDeleteColumn", err);
      setModalError({ title: "Delete column failed", message: (err && err.message) || "Unable to delete column." });
    } finally {
      setDeletingColumn(false);
    }
  }

  // ----------------- Totals + export -----------------
  function computeTotalForItem(item, headerNorms) {
    const excluded = new Set(["id", "nombre", "nom", "fotos", "pierres", "piedras", "pictures", "photos"]);
    let sum = 0;
    for (const k of headerNorms) {
      const nk = String(k || "").toLowerCase();
      if (excluded.has(nk) || nk === "total") continue;
      const v = item[nk];
      if (v === undefined || v === null || v === "") continue;
      const n = Number(String(v).replace(",", "."));
      if (Number.isFinite(n)) sum += n;
    }
    return sum;
  }

  function computeColumnTotalsForItems(itemsArr, headerNorms) {
    const totals = {};
    const excluded = new Set(["id", "nom", "nombre", "pierres", "piedras", "total"]);
    for (const nk of headerNorms) {
      const key = String(nk || "").toLowerCase();
      if (excluded.has(key)) { totals[key] = null; continue; }
      let anyNumeric = false;
      let sum = 0;
      for (const it of itemsArr) {
        const v = it[key];
        if (v === undefined || v === null || v === "") continue;
        const n = Number(String(v).replace(",", "."));
        if (Number.isFinite(n)) { sum += n; anyNumeric = true; }
      }
      totals[key] = anyNumeric ? sum : null;
    }
    return totals;
  }

  function getSelectionForCategory(category, headerNorms) {
    const saved = columnSelectionByCategory[category];
    if (saved) return saved;
    const map = {};
    map["id"] = true;
    headerNorms.forEach((n) => (map[n] = true));
    map["total"] = true;
    return map;
  }

  function buildRowsForExport(itemsArr, headerOrig, headerNorm, selection, rowSelectionMap, exportSelectedOnly, includeTotalsRow) {
    let filteredItems = itemsArr;
    if (exportSelectedOnly && rowSelectionMap) filteredItems = itemsArr.filter((it) => !!rowSelectionMap[it.ID]);

    const rows = [];
    const headers = [];
    if (selection["id"]) headers.push("ID");
    for (let i = 0; i < headerOrig.length; i++) {
      const orig = headerOrig[i];
      const norm = headerNorm[i];
      if (selection[norm]) headers.push(orig);
    }
    if (selection["total"]) headers.push("Total");
    rows.push(headers);

    for (const it of filteredItems) {
      const r = [];
      if (selection["id"]) r.push(it.ID);
      for (let i = 0; i < headerOrig.length; i++) {
        const norm = headerNorm[i];
        if (selection[norm]) r.push(it[norm] === undefined || it[norm] === null ? "" : it[norm]);
      }
      if (selection["total"]) r.push(computeTotalForItem(it, headerNorm));
      rows.push(r);
    }

    if (includeTotalsRow) {
      const colTotals = computeColumnTotalsForItems(filteredItems, headerNorm);
      const totalsRow = [];
      if (selection["id"]) totalsRow.push("Totals");
      for (let i = 0; i < headerOrig.length; i++) {
        const norm = headerNorm[i];
        if (selection[norm]) {
          const val = colTotals[norm];
          totalsRow.push(val === null || val === undefined ? "" : val);
        }
      }
      if (selection["total"]) {
        const totOfTotals = filteredItems.reduce((acc, it) => acc + computeTotalForItem(it, headerNorm), 0);
        totalsRow.push(totOfTotals);
      }
      rows.push(totalsRow);
    }

    return rows;
  }

  async function exportCategoryWithSelection(category) {
    setExportingCategory(true);
    try {
      let useItems = items;
      let useSchema = schema;
      if (category !== active) {
        const res = await fetch(`${API_BASE}/category/${encodeURIComponent(category)}`);
        if (!res.ok) throw new Error(await res.text());
        const json = await res.json();
        useItems = json.items || [];
        useSchema = json.schema || null;
      }
      const headerOrig = (useSchema && useSchema.headerOriginalOrder && useSchema.headerOriginalOrder.length)
        ? useSchema.headerOriginalOrder
        : (useItems.length ? Object.keys(useItems[0]).filter((k) => k !== "category" && k !== "ID") : []);
      const headerNorm = headerOrig.map((h) => String(h).toLowerCase());
      const selection = getSelectionForCategory(category, headerNorm);
      const rowSelectionMap = (rowSelectionByCategory && rowSelectionByCategory[category]) || null;
      const exportSelectedOnly = !!exportSelectedRowsOnlyByCategory[category];
      const includeTotalsRow = !!includeTotalsRowByCategory[category];

      const rows = buildRowsForExport(useItems, headerOrig, headerNorm, selection, rowSelectionMap, exportSelectedOnly, includeTotalsRow);
      const ws = XLSX.utils.aoa_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, (category || "category").substring(0, 31));
      const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([wbout], { type: "application/octet-stream" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${category}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error("exportCategoryWithSelection", err);
      setModalError({ title: "Export failed", message: (err && err.message) || "Unable to export category." });
    } finally {
      setExportingCategory(false);
    }
  }

  async function exportAllWithSelections() {
    setExportingAll(true);
    try {
      const wb = XLSX.utils.book_new();
      const catNames = categories.map((c) => c.name);
      for (const catName of catNames) {
        const res = await fetch(`${API_BASE}/category/${encodeURIComponent(catName)}`);
        if (!res.ok) { console.warn(`Skipping category ${catName}: failed to fetch items`); continue; }
        const json = await res.json();
        const useItems = json.items || [];
        const useSchema = json.schema || null;
        const headerOrig = (useSchema && useSchema.headerOriginalOrder && useSchema.headerOriginalOrder.length)
          ? useSchema.headerOriginalOrder
          : (useItems.length ? Object.keys(useItems[0]).filter((k) => k !== "category" && k !== "ID") : []);
        const headerNorm = headerOrig.map((h) => String(h).toLowerCase());
        const selection = getSelectionForCategory(catName, headerNorm);

        const rowSelectionMap = (rowSelectionByCategory && rowSelectionByCategory[catName]) || null;
        const exportSelectedOnly = !!exportSelectedRowsOnlyByCategory[catName];
        const includeTotalsRow = !!includeTotalsRowByCategory[catName];

        const rows = buildRowsForExport(useItems, headerOrig, headerNorm, selection, rowSelectionMap, exportSelectedOnly, includeTotalsRow);
        const ws = XLSX.utils.aoa_to_sheet(rows);
        XLSX.utils.book_append_sheet(wb, ws, (catName || "category").substring(0, 31));
      }
      const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([wbout], { type: "application/octet-stream" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `all_categories.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error("exportAllWithSelections", err);
      setModalError({ title: "Export failed", message: (err && err.message) || "Unable to export all categories." });
    } finally {
      setExportingAll(false);
    }
  }

  // ----------------- Render helpers -----------------
  function headerCheckboxRow(selection, norm) {
    return (
      <tr>
        <th className="p-1 border-b text-center">
          <input type="checkbox" checked={areAllRowsSelected()} onChange={() => setAllRowsSelectedForActive(!areAllRowsSelected())} className="cursor-pointer" />
        </th>

        <th className="p-1 border-b text-center">
          <input type="checkbox" checked={Boolean(selection["id"])} onChange={() => toggleColumnForActive("id")} className="cursor-pointer" />
        </th>

        {headerOriginalOrder().map((h, i) => {
          const protectedCol = isProtectedColumn(h);
          return (
            <th key={`chk-${i}`} className="p-1 border-b text-center">
              <div className="flex items-center justify-center gap-2">
                <input type="checkbox" checked={Boolean(selection[norm[i]])} onChange={() => toggleColumnForActive(norm[i])} className="cursor-pointer" />
                <button
                  onClick={(e) => { e.stopPropagation(); if (!protectedCol) setPendingDeleteColumn({ columnOrig: h, columnNorm: norm[i] }); }}
                  title={protectedCol ? `Cannot delete protected column "${h}"` : `Delete column "${h}"`}
                  className={`ml-1 inline-flex items-center justify-center w-6 h-6 rounded-full ${protectedCol ? 'text-gray-400 cursor-not-allowed' : 'text-red-600 hover:bg-red-100'}`}
                  disabled={protectedCol || deletingColumn}
                  type="button"
                >
                  {deletingColumn && pendingDeleteColumn && pendingDeleteColumn.columnOrig === h ? <Spinner className="h-4 w-4 text-red-600" /> : <Trash2 className="h-4 w-4" />}
                </button>
              </div>
            </th>
          );
        })}

        <th className="p-1 border-b text-center">
          <input type="checkbox" checked={Boolean(selection["total"])} onChange={() => toggleColumnForActive("total")} className="cursor-pointer" />
        </th>
      </tr>
    );
  }

  function renderTable() {
    if (!items || !items.length) return <div className="text-sm text-gray-500">No items found</div>;
    const orig = headerOriginalOrder();
    const norm = headerNormalizedOrder();
    const selection = getSelectionForCategory(active, norm);
    const totals = computeColumnTotalsForItems(items, norm);

    return (
      <div className="overflow-auto border border-gray-200 rounded shadow-sm">
        <table className="min-w-full table-fixed text-sm">
          <thead className="bg-gray-50 sticky top-0 z-10">
            {headerCheckboxRow(selection, norm)}
            <tr>
              <th className="p-2 border-b border-r text-left w-12">Sel</th>
              <th className="p-2 border-b border-r text-left w-56">ID</th>
              {orig.map((h) => (<th key={h} className="p-2 border-b border-r text-left">{h}</th>))}
              <th className="p-2 border-b text-left">Total</th>
            </tr>
          </thead>

          <tbody>
            {items.map((it) => {
              const map = (rowSelectionByCategory && rowSelectionByCategory[active]) || {};
              const selected = map[it.ID] === undefined ? true : !!map[it.ID];
              return (
                <tr key={it.ID} className="group odd:bg-white even:bg-gray-50 hover:bg-blue-50 transition-colors">
                  <td className="p-2 border-r text-center">
                    <input type="checkbox" checked={selected} onChange={() => toggleRowSelection(it.ID)} className="cursor-pointer" />
                  </td>

                  <td className="relative p-2 border-r align-top font-mono text-xs bg-gray-100 group-hover:bg-blue-50">
                    <div className="truncate">{it.ID}</div>
                  </td>

                  {orig.map((h, colIndex) => {
                    const nkey = norm[colIndex] || String(h).toLowerCase();
                    const value = it[nkey];
                    const isEditing = editingCell && editingCell.id === it.ID && editingCell.normKey === nkey;
                    const savingKey = `${it.ID}|${nkey}`;
                    return (
                      <td key={nkey} className="relative p-2 border-r align-top whitespace-nowrap overflow-hidden" onDoubleClick={() => startEdit(it.ID, nkey, value)}>
                        {!isEditing && (
                          <div className="text-xs leading-7 h-8 overflow-hidden truncate">
                            {savingCellKey === savingKey ? <span className="text-indigo-600">Saving…</span> : (value === undefined || value === null ? "" : String(value))}
                          </div>
                        )}
                        {isEditing && (
                          <input autoFocus value={cellValue} onChange={(e) => setCellValue(e.target.value)} onKeyDown={(e) => onInputKeyDown(e, it.ID, nkey)} onBlur={() => {
                            const original = editingCell ? (editingCell.original == null ? "" : String(editingCell.original)) : "";
                            const incoming = cellValue == null ? "" : String(cellValue);
                            if (incoming === original) cancelEdit(); else saveEdit(it.ID, nkey, cellValue);
                          }} className="absolute inset-0 w-full h-full px-1 text-sm bg-white focus:outline-none box-border" style={{ padding: "4px 6px", border: "1px solid rgba(0,0,0,0.08)", borderRadius: 4 }} />
                        )}
                      </td>
                    );
                  })}

                  <td className="p-2 align-top text-sm">{computeTotalForItem(it, norm)}</td>
                </tr>
              );
            })}
          </tbody>

          <tfoot className="bg-gray-100">
            <tr>
              <td className="p-2 border-t text-left font-medium">Totals</td>
              <td className="p-2 border-t border-r font-medium"></td>
              {orig.map((h, i) => {
                const nkey = norm[i] || String(h).toLowerCase();
                const v = totals[nkey];
                return <td key={`tot-${nkey}`} className="p-2 border-t border-r text-sm font-semibold">{v == null ? "" : String(v)}</td>;
              })}
              <td className="p-2 border-t text-sm font-semibold"></td>
            </tr>
          </tfoot>
        </table>

        {/* Row actions */}
        <div className="p-3 flex items-center gap-3">
          <button onClick={onRequestDeleteSelectedRows} disabled={!anyRowSelected()} className={`px-3 py-1 rounded bg-red-600 text-white ${!anyRowSelected() ? "opacity-60 cursor-not-allowed" : "hover:bg-red-700"}`}>Delete selected rows</button>
        </div>
      </div>
    );
  }

  // Editing save helper
  async function saveEdit(id, normKey, newValue) {
    const original = editingCell ? (editingCell.original == null ? "" : String(editingCell.original)) : "";
    const incoming = newValue == null ? "" : String(newValue);
    if (incoming === original) {
      cancelEdit();
      return;
    }
    const savingKeyLocal = `${id}|${normKey}`;
    setSavingCellKey(savingKeyLocal);
    try {
      const body = { attribute: normKey, value: incoming, category: active };
      const res = await fetch(`${API_BASE}/item/${encodeURIComponent(id)}`, { method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
      if (!res.ok) {
        const txt = await res.text().catch(() => "");
        throw new Error(txt || `Server ${res.status}`);
      }
      const json = await res.json();
      setItems((prev) => prev.map((it) => (it.ID === id ? { ...it, [normKey]: incoming, ...(json.updated || {}) } : it)));
      cancelEdit();
    } catch (err) {
      console.error("saveEdit", err);
      setModalError({ title: "Save failed", message: (err && err.message) || "Unable to save change." });
    } finally {
      setSavingCellKey(null);
    }
  }

  function startEdit(id, normKey, initial) {
    setEditingCell({ id, normKey, original: initial == null ? "" : String(initial) });
    setCellValue(initial == null ? "" : String(initial));
  }
  function cancelEdit() { setEditingCell(null); setCellValue(""); }

  function onInputKeyDown(e, id, normKey) {
    if (e.key === "Enter") { e.preventDefault(); saveEdit(id, normKey, cellValue); }
    else if (e.key === "Escape") cancelEdit();
  }

  // ----------------- Small UI pieces (replace/delete/export buttons) -----------------
  function headerControls() {
    return (
      <div className="flex gap-2 items-center">
        <button onClick={exportAllWithSelections} className={`px-3 py-1 bg-indigo-600 text-white rounded flex items-center gap-2 ${exportingAll ? "opacity-70 cursor-wait" : "hover:bg-indigo-700"}`} disabled={exportingAll || exportingCategory || deletingColumn || deletingRows || replacing}>
          {exportingAll ? <Spinner className="h-4 w-4 text-white" /> : null}
          <span>{exportingAll ? "Preparing…" : "Download All"}</span>
        </button>

        <button onClick={() => active && exportCategoryWithSelection(active)} disabled={!active || exportingCategory || exportingAll || deletingColumn || deletingRows || replacing} className={`px-3 py-1 bg-green-600 text-white rounded flex items-center gap-2 ${(!active || exportingCategory) ? "opacity-70 cursor-not-allowed" : "hover:bg-green-700"}`}>
          {exportingCategory ? <Spinner className="h-4 w-4 text-white" /> : null}
          <span>{exportingCategory ? "Preparing…" : "Download Category"}</span>
        </button>

        <button onClick={onReplaceAllClick} disabled={!active || replacing || deletingColumn || deletingRows} className={`px-3 py-1 bg-red-600 text-white rounded ${(!active || replacing) ? "opacity-70 cursor-not-allowed" : "hover:bg-red-700"}`}>
          {replacing ? <Spinner className="h-4 w-4 text-white" /> : null}
          <span>{replacing ? "Replacing…" : "Replace Category"}</span>
        </button>

        <input ref={fileInputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => { const f = e.target.files?.[0] ?? null; if (!f) return; setPendingReplaceFile(f); setConfirmReplaceOpen(true); e.target.value = ""; }} />
      </div>
    );
  }

  function onReplaceAllClick() {
    if (!active) { setModalError({ title: "No category selected", message: "Select a category before replacing." }); return; }
    fileInputRef.current?.click();
  }

  // Column delete trigger already uses pendingDeleteColumn -> confirmDeleteColumn

  // ----------------- JSX -----------------
  return (
    <div className="max-w-7xl mx-auto p-4">
      <header className="mb-6 flex items-start justify-between">
        <div>
          <h1 className="text-2xl font-semibold">Inventory Viewer</h1>
          <p className="text-sm text-gray-500 mt-1">Browse categories, edit inline, select columns/rows and export.</p>
        </div>
        {headerControls()}
      </header>

      <div className="mb-4 flex items-center gap-3">
        <div className="font-medium">Categories:</div>
        {loadingCats ? <div className="text-sm text-gray-500">Loading...</div> : (
          <div className="flex gap-3 overflow-auto">
            {categories.map((c) => (
              <div key={c.name} className={`flex items-center justify-between gap-2 px-3 py-1 rounded ${active === c.name ? "bg-blue-600 text-white" : "bg-gray-100 text-gray-800"} hover:shadow-sm`} style={{ minWidth: 160, maxWidth: 320 }}>
                <button onClick={() => setActive(c.name)} className="flex-1 text-left truncate" title={`Select category ${c.name}`}>{c.name}</button>
                <button onClick={(e) => { e.stopPropagation(); setPendingDeleteCategory(c.name); }} className="ml-2 inline-flex items-center justify-center w-8 h-8 rounded-full bg-red-100 hover:bg-red-200" title={`Delete category ${c.name}`}>
                  <Trash2 className="h-4 w-4 text-red-600" />
                </button>
              </div>
            ))}
          </div>
        )}
      </div>

      {error && <div className="mb-4 text-sm text-red-600">{error}</div>}

      <main>
        <div className="mb-4">
          <h2 className="text-lg font-medium mb-2">{active || "No category selected"}</h2>
          {loadingItems ? <div className="text-sm text-gray-500">Loading items...</div> : renderTable()}
        </div>
      </main>

      {/* Modals: Delete Column */}
      {pendingDeleteColumn && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40">
          <div className="bg-white rounded-lg shadow-lg p-6 max-w-md w-full">
            <h3 className="text-lg font-semibold mb-2">Delete column "{pendingDeleteColumn.columnOrig}"?</h3>
            <p className="text-sm text-gray-700 mb-4">This will remove the column <strong>{pendingDeleteColumn.columnOrig}</strong> from all items in category <strong>{active}</strong>. This cannot be undone.</p>
            <div className="flex gap-2 justify-end">
              <button onClick={() => setPendingDeleteColumn(null)} className="px-3 py-1 rounded bg-gray-200 hover:bg-gray-300">Cancel</button>
              <button onClick={confirmDeleteColumn} className="px-3 py-1 rounded bg-red-600 text-white flex items-center gap-2 hover:bg-red-700" disabled={deletingColumn}>
                {deletingColumn ? <Spinner className="h-4 w-4 text-white" /> : null}
                <span>{deletingColumn ? "Deleting…" : "Confirm Delete"}</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Delete selected rows modal */}
      {confirmDeleteRowsOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40">
          <div className="bg-white rounded-lg shadow-lg p-6 max-w-md w-full">
            <h3 className="text-lg font-semibold mb-2">Delete selected rows?</h3>
            <p className="text-sm text-gray-700 mb-4">This will permanently delete {pendingDeleteRowIds.length} rows from category <strong>{active}</strong>. The first data row and the totals row cannot be deleted.</p>
            <div className="flex gap-2 justify-end">
              <button onClick={() => { setConfirmDeleteRowsOpen(false); setPendingDeleteRowIds([]); }} className="px-3 py-1 rounded bg-gray-200 hover:bg-gray-300">Cancel</button>
              <button onClick={confirmDeleteSelectedRows} disabled={deletingRows} className="px-3 py-1 rounded bg-red-600 text-white flex items-center gap-2 hover:bg-red-700">
                {deletingRows ? <Spinner className="h-4 w-4 text-white" /> : null}
                <span>{deletingRows ? "Deleting…" : `Delete ${pendingDeleteRowIds.length} rows`}</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Replace modal */}
      {confirmReplaceOpen && pendingReplaceFile && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40">
          <div className="bg-white rounded-lg shadow-lg p-6 max-w-lg w-full">
            <h3 className="text-lg font-semibold mb-2">Replace category "{active}"?</h3>
            <p className="text-sm text-gray-700 mb-4">This will DELETE all items in category <strong>{active}</strong> and replace them with the contents of:</p>
            <div className="mb-4">
              <div className="text-sm font-medium">{pendingReplaceFile.name}</div>
              <div className="text-xs text-gray-500">{Math.round(pendingReplaceFile.size / 1024)} KB</div>
            </div>
            <div className="flex gap-2 justify-end">
              <button onClick={() => { setConfirmReplaceOpen(false); setPendingReplaceFile(null); }} className="px-3 py-1 rounded bg-gray-200 hover:bg-gray-300">Cancel</button>
              <button onClick={confirmReplace} className="px-3 py-1 rounded bg-red-600 text-white flex items-center gap-2 hover:bg-red-700" disabled={replacing}>
                {replacing ? <Spinner className="h-4 w-4 text-white" /> : null}
                <span>{replacing ? "Replacing…" : "Confirm Replace"}</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Delete category modal */}
      {pendingDeleteCategory && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40">
          <div className="bg-white rounded-lg shadow-lg p-6 max-w-md w-full">
            <h3 className="text-lg font-semibold mb-2">Delete category "{pendingDeleteCategory}"?</h3>
            <p className="text-sm text-gray-700 mb-4">This will DELETE all items in category <strong>{pendingDeleteCategory}</strong>. This action cannot be undone.</p>
            <div className="flex gap-2 justify-end">
              <button onClick={() => setPendingDeleteCategory(null)} className="px-3 py-1 rounded bg-gray-200 hover:bg-gray-300">Cancel</button>
              <button onClick={async () => {
                setDeletingCategory(true);
                try {
                  const res = await fetch(`${API_BASE}/category/${encodeURIComponent(pendingDeleteCategory)}`, { method: 'DELETE' });
                  if (!res.ok) throw new Error(await res.text());
                  await loadCategories();
                  setItems([]); setSchema(null); setPendingDeleteCategory(null);
                } catch (err) {
                  console.error("delete category", err);
                  setModalError({ title: "Delete failed", message: (err && err.message) || "Unable to delete category." });
                } finally {
                  setDeletingCategory(false);
                }
              }} className="px-3 py-1 rounded bg-red-600 text-white flex items-center gap-2 hover:bg-red-700" disabled={deletingCategory}>
                {deletingCategory ? <Spinner className="h-4 w-4 text-white" /> : null}
                <span>{deletingCategory ? "Deleting…" : "Confirm Delete"}</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Error modal */}
      {modalError && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40" onClick={() => setModalError(null)}>
          <div className="bg-white rounded-lg shadow-lg p-4 max-w-lg w-full" onClick={(e) => e.stopPropagation()}>
            <h3 className="font-semibold text-lg mb-2">{modalError.title}</h3>
            <div className="text-sm text-gray-700 mb-4">{modalError.message}</div>
            <div className="flex justify-end"><button onClick={() => setModalError(null)} className="px-3 py-1 rounded bg-gray-200 hover:bg-gray-300">Close</button></div>
          </div>
        </div>
      )}
    </div>
  );
}