import React, { useEffect, useRef, useState } from "react";
import { Trash2, Plus, Image as ImageIcon } from "lucide-react";
import * as XLSX from "xlsx"; // kept for non-image export paths if needed
import { isLocalHost } from "../api";

const API_BASE = isLocalHost
  ? "http://localhost:8080"
  : "https://qr-inventory-scanner-backend.vercel.app";

// LocalStorage keys
const LS_ROWS_KEY = "qr_viewer_row_selection_v1";
const LS_COLS_KEY = "qr_viewer_col_selection_v1";

/**
 * Viewer.jsx
 *
 * Right-click:
 * - Header cell: "Add column (left/right)" (not on ID).
 * - Body cell: "Add row (top/bottom)", "Delete row", "Insert image".
 *   Image insertion allowed on ANY column.
 */

function Spinner({ className = "h-4 w-4 text-white" }) {
  return (
    <svg
      className={`animate-spin ${className}`}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
    >
      <circle className="opacity-25" cx="12" cy="12" r="10" strokeWidth="3" />
      <path
        className="opacity-75"
        d="M4 12a8 8 0 018-8"
        strokeWidth="3"
        strokeLinecap="round"
      />
    </svg>
  );
}

// UUID v4
function uuidv4() {
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
    const r = crypto.getRandomValues(new Uint8Array(1))[0] & 0xf;
    const v = c === "x" ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });
}

// Try parse image JSON
function parseImageValue(val) {
  if (!val || typeof val !== "string") return null;
  try {
    const obj = JSON.parse(val);
    if (obj && obj.type === "image" && obj.src) return obj;
  } catch {}
  return null;
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
  const [editingCell, setEditingCell] = useState(null);
  const [cellValue, setCellValue] = useState("");
  const [savingCellKey, setSavingCellKey] = useState(null);

  // File/replace
  const fileInputRef = useRef(null);
  const [pendingReplaceFile, setPendingReplaceFile] = useState(null);
  const [confirmReplaceOpen, setConfirmReplaceOpen] = useState(false);

  // Delete column modal
  const [pendingDeleteColumn, setPendingDeleteColumn] = useState(null);
  // Delete rows modal
  const [pendingDeleteRowIds, setPendingDeleteRowIds] = useState([]);
  const [confirmDeleteRowsOpen, setConfirmDeleteRowsOpen] = useState(false);
  // Delete category modal
  const [pendingDeleteCategory, setPendingDeleteCategory] = useState(null);

  // Insert column modal state
  const [insertColumnOpen, setInsertColumnOpen] = useState(false);
  const [newColumnName, setNewColumnName] = useState("");
  const [insertColumnIndex, setInsertColumnIndex] = useState(0);
  const [insertingColumn, setInsertingColumn] = useState(false);

  // Insert row modal state
  const [insertRowOpen, setInsertRowOpen] = useState(false);
  const [insertRowIndex, setInsertRowIndex] = useState(0);
  const [insertingRow, setInsertingRow] = useState(false);
  const [newRowId, setNewRowId] = useState("");
  const [newRowValues, setNewRowValues] = useState({});
  const [newRowIdChecking, setNewRowIdChecking] = useState(false);
  const [newRowIdExists, setNewRowIdExists] = useState(false);

  // Insert image modal state
  const [insertImageOpen, setInsertImageOpen] = useState(false);
  const [insertImageTarget, setInsertImageTarget] = useState(null); // { id, normKey }
  const [imageUploadMode, setImageUploadMode] = useState("upload"); // 'upload' | 'url'
  const [imageFile, setImageFile] = useState(null);
  const [imageUrl, setImageUrl] = useState("");
  const [imageAlt, setImageAlt] = useState("");
  const [imageWidth, setImageWidth] = useState(64);
  const [imageHeight, setImageHeight] = useState(64);
  const [uploadingImage, setUploadingImage] = useState(false);
  const [imagePreviewUrl, setImagePreviewUrl] = useState("");

  // Context menus
  const [ctxMenu, setCtxMenu] = useState({
    open: false,
    x: 0,
    y: 0,
    type: null, // 'column' | 'row'
    columnIndex: null,
    columnName: null,
    rowIndex: null,
    rowId: null,
    isHeaderCell: false,
    cellNormKey: null,
  });

  // Close context menu on Escape or global click
  useEffect(() => {
    const onKey = (e) => {
      if (e.key === "Escape") setCtxMenu((p) => ({ ...p, open: false }));
    };
    const onDocClick = () => setCtxMenu((p) => ({ ...p, open: false }));
    if (ctxMenu.open) {
      document.addEventListener("keydown", onKey);
      document.addEventListener("click", onDocClick);
      return () => {
        document.removeEventListener("keydown", onKey);
        document.removeEventListener("click", onDocClick);
      };
    }
  }, [ctxMenu.open]);

  // Persisted selections
  const [columnSelectionByCategory, setColumnSelectionByCategory] = useState(
    () => {
      try {
        const raw = window.localStorage.getItem(LS_COLS_KEY);
        return raw ? JSON.parse(raw) : {};
      } catch {
        return {};
      }
    }
  );
  const [rowSelectionByCategory, setRowSelectionByCategory] = useState(() => {
    try {
      const raw = window.localStorage.getItem(LS_ROWS_KEY);
      return raw ? JSON.parse(raw) : {};
    } catch {
      return {};
    }
  });
  const [
    exportSelectedRowsOnlyByCategory,
    setExportSelectedRowsOnlyByCategory,
  ] = useState({});
  const [includeTotalsRowByCategory, setIncludeTotalsRowByCategory] = useState(
    {}
  );

  useEffect(() => {
    try {
      window.localStorage.setItem(
        LS_COLS_KEY,
        JSON.stringify(columnSelectionByCategory)
      );
    } catch {}
  }, [columnSelectionByCategory]);
  useEffect(() => {
    try {
      window.localStorage.setItem(
        LS_ROWS_KEY,
        JSON.stringify(rowSelectionByCategory)
      );
    } catch {}
  }, [rowSelectionByCategory]);

  // Load categories/items
  useEffect(() => {
    loadCategories();
  }, []);
  useEffect(() => {
    if (active) loadCategoryItems(active);
  }, [active]);

  // ----------------- Backend calls -----------------
  async function loadCategories() {
    setLoadingCats(true);
    setError(null);
    try {
      const res = await fetch(`${API_BASE}/categories`);
      if (!res.ok) throw new Error(await res.text());
      const json = await res.json();
      setCategories(json.categories || []);
      if (json.categories?.length) {
        if (!active || !json.categories.find((c) => c.name === active))
          setActive(json.categories[0].name);
      } else {
        setActive(null);
      }
    } catch (err) {
      setError(err?.message || "Failed to load categories");
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
      const res = await fetch(
        `${API_BASE}/category/${encodeURIComponent(category)}`
      );
      if (!res.ok) throw new Error(await res.text());
      const json = await res.json();
      setSchema(json.schema || null);
      setItems(json.items || []);
      initializeSelectionForCategory(
        category,
        json.schema || null,
        json.items || []
      );
      setExportSelectedRowsOnlyByCategory((prev) =>
        prev && typeof prev[category] !== "undefined"
          ? prev
          : { ...(prev || {}), [category]: false }
      );
      setIncludeTotalsRowByCategory((prev) =>
        prev && typeof prev[category] !== "undefined"
          ? prev
          : { ...(prev || {}), [category]: true }
      );
    } catch (err) {
      setError(err?.message || "Failed to load items");
    } finally {
      setLoadingItems(false);
    }
  }

  async function checkIdExists(id) {
    if (!id || !String(id).trim()) {
      setNewRowIdExists(false);
      return false;
    }
    setNewRowIdChecking(true);
    try {
      const res = await fetch(
        `${API_BASE}/item/exists/${encodeURIComponent(id)}`
      );
      const json = res.ok ? await res.json() : {};
      const exists = !!json.exists;
      setNewRowIdExists(exists);
      return exists;
    } catch {
      setNewRowIdExists(false);
      return false;
    } finally {
      setNewRowIdChecking(false);
    }
  }

  async function presignUpload(filename, contentType, meta) {
    const res = await fetch(`${API_BASE}/uploads/presign`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ filename, contentType, ...meta }),
    });
    if (!res.ok) throw new Error(await res.text());
    return await res.json(); // { uploadUrl, finalUrl, s3Key, expires }
  }

  // ----------------- Header helpers -----------------
  function headerOriginalOrder() {
    if (schema?.headerOriginalOrder?.length) return schema.headerOriginalOrder;
    if (!items?.length) return [];
    return Object.keys(items[0]).filter((k) => k !== "category" && k !== "ID");
  }
  function headerNormalizedOrder() {
    if (schema?.headerNormalizedOrder?.length)
      return schema.headerNormalizedOrder;
    return headerOriginalOrder().map((h) => String(h).toLowerCase());
  }
  function norm(s) {
    return String(s || "")
      .trim()
      .toLowerCase();
  }

  // ----------------- Initialize selections -----------------
  function initializeSelectionForCategory(category, schemaObj, itemsArr) {
    setColumnSelectionByCategory((prev) => {
      if (prev?.[category]) return prev;
      const next = { ...(prev || {}) };
      const orig = schemaObj?.headerOriginalOrder?.length
        ? schemaObj.headerOriginalOrder.slice()
        : itemsArr?.length
        ? Object.keys(itemsArr[0]).filter((k) => k !== "category" && k !== "ID")
        : [];
      const normalized = orig.map((h) => String(h).toLowerCase());
      const map = { id: true, total: true };
      normalized.forEach((n) => (map[n] = true));
      next[category] = map;
      return next;
    });

    setRowSelectionByCategory((prev) => {
      const existing = (prev && prev[category]) || {};
      const next = { ...(prev || {}) };
      const currentIds = new Set((itemsArr || []).map((it) => it.ID));
      const merged = {};
      for (const id of Object.keys(existing)) {
        if (currentIds.has(id)) merged[id] = !!existing[id];
      }
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
    const map =
      (rowSelectionByCategory && rowSelectionByCategory[active]) || {};
    if (!items?.length) return false;
    return items.every((it) => !!map[it.ID]);
  }
  function anyRowSelected() {
    if (!active) return false;
    const map =
      (rowSelectionByCategory && rowSelectionByCategory[active]) || {};
    return (items || []).some((it) => !!map[it.ID]);
  }

  // ----------------- Delete rows bulk -----------------
  function onRequestDeleteSelectedRows() {
    if (!active) return;
    const map =
      (rowSelectionByCategory && rowSelectionByCategory[active]) || {};
    const ids = (items || []).filter((it) => !!map[it.ID]).map((it) => it.ID);
    if (!ids.length) {
      setModalError({
        title: "No rows selected",
        message: "Select rows to delete first.",
      });
      return;
    }
    const firstId = items[0] ? items[0].ID : null;
    if (ids.includes(firstId)) {
      setModalError({
        title: "Cannot delete first row",
        message: "The first data row is protected and cannot be deleted.",
      });
      return;
    }
    setPendingDeleteRowIds(ids);
    setConfirmDeleteRowsOpen(true);
  }
  async function confirmDeleteSelectedRows() {
    if (!active || !pendingDeleteRowIds.length) {
      setPendingDeleteRowIds([]);
      setConfirmDeleteRowsOpen(false);
      return;
    }
    setDeletingRows(true);
    try {
      const responses = await Promise.all(
        pendingDeleteRowIds.map((id) =>
          fetch(`${API_BASE}/item/${encodeURIComponent(id)}`, {
            method: "DELETE",
          })
        )
      );
      const failed = [];
      for (let i = 0; i < responses.length; i++) {
        if (!responses[i].ok) {
          const txt = await responses[i].text().catch(() => "");
          failed.push({
            id: pendingDeleteRowIds[i],
            text: txt,
            status: responses[i].status,
          });
        }
      }
      if (failed.length)
        setModalError({
          title: "Delete incomplete",
          message: `Failed to delete ${failed.length} rows. See console.`,
        });
      await loadCategoryItems(active);
      setPendingDeleteRowIds([]);
      setConfirmDeleteRowsOpen(false);
    } catch (err) {
      setModalError({
        title: "Delete failed",
        message: err?.message || "Unable to delete rows.",
      });
    } finally {
      setDeletingRows(false);
    }
  }

  // ----------------- Column delete -----------------
  function isProtectedColumn(orig) {
    const n = String(orig || "")
      .trim()
      .toLowerCase();
    return n === "nom" || n === "nombre" || n === "pierres" || n === "piedras";
  }
  async function confirmDeleteColumn() {
    if (!active || !pendingDeleteColumn) {
      setPendingDeleteColumn(null);
      return;
    }
    const col = pendingDeleteColumn.columnOrig;
    if (isProtectedColumn(col)) {
      setModalError({
        title: "Protected column",
        message: `The column "${col}" cannot be deleted.`,
      });
      setPendingDeleteColumn(null);
      return;
    }
    setDeletingColumn(true);
    try {
      let res = await fetch(
        `${API_BASE}/category/${encodeURIComponent(
          active
        )}/column/${encodeURIComponent(col)}`,
        { method: "DELETE" }
      );
      if (!res.ok) {
        res = await fetch(
          `${API_BASE}/category/${encodeURIComponent(active)}/column`,
          {
            method: "DELETE",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ column: col }),
          }
        );
      }
      if (!res.ok) {
        const txt = await res.text().catch(() => "");
        throw new Error(txt || `Server responded ${res.status}`);
      }
      await loadCategoryItems(active);
      setPendingDeleteColumn(null);
    } catch (err) {
      setModalError({
        title: "Delete column failed",
        message: err?.message || "Unable to delete column.",
      });
    } finally {
      setDeletingColumn(false);
    }
  }

  // ----------------- Insert column (modal workflow) -----------------
  function openInsertColumnAt(index) {
    setInsertColumnIndex(index);
    setNewColumnName("");
    setInsertColumnOpen(true);
  }
  async function confirmInsertColumn() {
    if (!active) return;
    const name = newColumnName.trim();
    if (!name) {
      setModalError({
        title: "Column name required",
        message: "Please enter a column name.",
      });
      return;
    }
    setInsertingColumn(true);
    try {
      const res = await fetch(
        `${API_BASE}/category/${encodeURIComponent(active)}/column`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            columnName: name,
            insertIndex: insertColumnIndex,
            defaultValue: null,
          }),
        }
      );
      if (!res.ok) throw new Error(await res.text());
      const nName = norm(name);
      setSchema((prev) => {
        const orig = prev?.headerOriginalOrder
          ? prev.headerOriginalOrder.slice()
          : headerOriginalOrder().slice();
        const normArr = prev?.headerNormalizedOrder
          ? prev.headerNormalizedOrder.slice()
          : headerNormalizedOrder().slice();
        orig.splice(insertColumnIndex, 0, name);
        normArr.splice(insertColumnIndex, 0, nName);
        return {
          ...(prev || {}),
          headerOriginalOrder: orig,
          headerNormalizedOrder: normArr,
          updatedAt: new Date().toISOString(),
        };
      });
      setItems((prev) => prev.map((it) => ({ ...it, [nName]: null })));
      setColumnSelectionByCategory((prev) => {
        const map = (prev && prev[active]) || {};
        return { ...prev, [active]: { ...map, [nName]: true } };
      });
      setInsertColumnOpen(false);
    } catch (err) {
      setModalError({
        title: "Add column failed",
        message: err?.message || "Unable to add column.",
      });
    } finally {
      setInsertingColumn(false);
    }
  }

  // ----------------- Context menus -----------------
  function onHeaderCellContextMenu(e, idx, label) {
    e.preventDefault();
    const isId =
      String(label || "")
        .trim()
        .toLowerCase() === "id";
    if (isId) return;
    setCtxMenu({
      open: true,
      x: e.clientX + 2,
      y: e.clientY + 2,
      type: "column",
      columnIndex: idx,
      columnName: label,
      rowIndex: null,
      rowId: null,
      isHeaderCell: true,
      cellNormKey: null,
    });
  }
  function ColumnContextMenu() {
    if (!ctxMenu.open || ctxMenu.type !== "column") return null;
    const idx = ctxMenu.columnIndex ?? 0;
    const colName = String(ctxMenu.columnName || "");
    if (colName.trim().toLowerCase() === "id") return null;

    return (
      <div
        className="fixed z-50 bg-white border border-gray-300 rounded shadow-lg"
        style={{ top: ctxMenu.y, left: ctxMenu.x, minWidth: 220 }}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="px-3 py-2 text-xs text-gray-500">Column: {colName}</div>
        <button
          className="w-full text-left px-3 py-2 hover:bg-gray-100"
          onClick={() => {
            setCtxMenu((p) => ({ ...p, open: false }));
            openInsertColumnAt(idx);
          }}
        >
          Add column (left)
        </button>
        <button
          className="w-full text-left px-3 py-2 hover:bg-gray-100"
          onClick={() => {
            setCtxMenu((p) => ({ ...p, open: false }));
            openInsertColumnAt(idx + 1);
          }}
        >
          Add column (right)
        </button>
      </div>
    );
  }

  function onBodyCellContextMenu(e, rowIdx, rowId, cellNormKey) {
    e.preventDefault();
    setCtxMenu({
      open: true,
      x: e.clientX + 2,
      y: e.clientY + 2,
      type: "row",
      columnIndex: null,
      columnName: null,
      rowIndex: rowIdx,
      rowId,
      isHeaderCell: false,
      cellNormKey,
    });
  }

  function openInsertRowAt(index) {
    setInsertRowIndex(index);
    const headerNorm = headerNormalizedOrder();
    const initValues = {};
    for (const nk of headerNorm) {
      if (nk === "total" || nk === "id" || nk === "category") continue;
      initValues[nk] = "";
    }
    const defaultId = uuidv4();
    setNewRowValues(initValues);
    setNewRowId(defaultId);
    setInsertRowOpen(true);
    checkIdExists(defaultId);
  }

  async function confirmInsertRow() {
    if (!active) return;
    const id = String(newRowId || "").trim();
    if (!id) {
      setModalError({
        title: "ID required",
        message: "Please enter an ID for the new row.",
      });
      return;
    }
    const exists = await checkIdExists(id);
    if (exists) {
      setModalError({
        title: "ID already exists",
        message: `Another item with ID "${id}" already exists. Choose a different ID.`,
      });
      return;
    }
    setInsertingRow(true);
    try {
      const headerNorm = headerNormalizedOrder();
      const values = {};
      for (const nk of headerNorm) {
        if (nk === "total" || nk === "id" || nk === "category") continue;
        const v = newRowValues[nk];
        values[nk] = v === "" ? null : v;
      }

      const res = await fetch(
        `${API_BASE}/category/${encodeURIComponent(active)}/row`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ insertIndex: insertRowIndex, id, values }),
        }
      );
      if (!res.ok) {
        const txt = await res.text().catch(() => "");
        throw new Error(txt || `Server responded ${res.status}`);
      }
      const json = await res.json();
      const created = json.item || null;
      if (!created) {
        await loadCategoryItems(active);
        setInsertRowOpen(false);
        return;
      }

      setItems((prev) => {
        const next = prev.slice();
        next.splice(insertRowIndex, 0, created);
        return next;
      });
      setRowSelectionByCategory((prev) => {
        const map = (prev && prev[active]) || {};
        return { ...prev, [active]: { ...map, [created.ID]: true } };
      });
      setInsertRowOpen(false);
    } catch (err) {
      setModalError({
        title: "Add row failed",
        message: err?.message || "Unable to add row.",
      });
    } finally {
      setInsertingRow(false);
    }
  }

  async function deleteRowById(id) {
    if (!id) return;
    const firstId = items[0] ? items[0].ID : null;
    if (id === firstId) {
      setModalError({
        title: "Cannot delete first row",
        message: "The first data row is protected and cannot be deleted.",
      });
      return;
    }
    setDeletingRows(true);
    try {
      const res = await fetch(`${API_BASE}/item/${encodeURIComponent(id)}`, {
        method: "DELETE",
      });
      if (!res.ok) {
        const txt = await res.text().catch(() => "");
        throw new Error(txt || `Server responded ${res.status}`);
      }
      setItems((prev) => prev.filter((it) => it.ID !== id));
      setRowSelectionByCategory((prev) => {
        const map = (prev && prev[active]) || {};
        const nextMap = { ...map };
        delete nextMap[id];
        return { ...prev, [active]: nextMap };
      });
    } catch (err) {
      setModalError({
        title: "Delete failed",
        message: err?.message || "Unable to delete row.",
      });
    } finally {
      setDeletingRows(false);
      setCtxMenu((p) => ({ ...p, open: false }));
    }
  }

  function RowContextMenu() {
    if (!ctxMenu.open || ctxMenu.type !== "row" || ctxMenu.isHeaderCell)
      return null;
    const idx = ctxMenu.rowIndex ?? 0;
    const rowId = ctxMenu.rowId;
    const cellNormKey = ctxMenu.cellNormKey || null;
    return (
      <div
        className="fixed z-50 bg-white border border-gray-300 rounded shadow-lg"
        style={{ top: ctxMenu.y, left: ctxMenu.x, minWidth: 240 }}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="px-3 py-2 text-xs text-gray-500">Row #{idx + 1}</div>
        <button
          className="w-full text-left px-3 py-2 hover:bg-gray-100"
          onClick={() => {
            setCtxMenu((p) => ({ ...p, open: false }));
            openInsertRowAt(idx);
          }}
          disabled={insertingRow}
        >
          Add row (top)
        </button>
        <button
          className="w-full text-left px-3 py-2 hover:bg-gray-100"
          onClick={() => {
            setCtxMenu((p) => ({ ...p, open: false }));
            openInsertRowAt(idx + 1);
          }}
          disabled={insertingRow}
        >
          Add row (bottom)
        </button>
        <div className="border-t my-1" />
        <button
          className="w-full text-left px-3 py-2 hover:bg-red-50 text-red-700"
          onClick={() => deleteRowById(rowId)}
          disabled={deletingRows}
        >
          {deletingRows ? "Deleting…" : "Delete row"}
        </button>
        <div className="border-t my-1" />
        <button
          className="w-full text-left px-3 py-2 hover:bg-gray-100 flex items-center gap-2"
          onClick={() => {
            setCtxMenu((p) => ({ ...p, open: false }));
            setInsertImageTarget({ id: rowId, normKey: cellNormKey });
            setImageUploadMode("upload");
            setImageFile(null);
            setImageUrl("");
            setImageAlt("");
            setImageWidth(64);
            setImageHeight(64);
            setImagePreviewUrl("");
            setInsertImageOpen(true);
          }}
        >
          <ImageIcon className="h-4 w-4" /> Insert image
        </button>
      </div>
    );
  }

  // ----------------- Image Insert Modal -----------------
  const InsertImageModal = insertImageOpen ? (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/40"
      onClick={() => !uploadingImage && setInsertImageOpen(false)}
    >
      <div
        className="bg-white rounded-lg shadow-lg p-6 max-w-lg w-full"
        onClick={(e) => e.stopPropagation()}
      >
        <h3 className="text-lg font-semibold mb-2">Insert image</h3>
        <p className="text-sm text-gray-700 mb-4">
          Set image for item <strong>{insertImageTarget?.id}</strong>, column{" "}
          <strong>{insertImageTarget?.normKey}</strong>.
        </p>

        <div className="mb-3">
          <div className="flex gap-2">
            <button
              className={`px-3 py-1 rounded ${
                imageUploadMode === "upload"
                  ? "bg-indigo-600 text-white"
                  : "bg-gray-200"
              }`}
              onClick={() => setImageUploadMode("upload")}
              disabled={uploadingImage}
            >
              Upload file
            </button>
            <button
              className={`px-3 py-1 rounded ${
                imageUploadMode === "url"
                  ? "bg-indigo-600 text-white"
                  : "bg-gray-200"
              }`}
              onClick={() => setImageUploadMode("url")}
              disabled={uploadingImage}
            >
              Paste URL
            </button>
          </div>
        </div>

        {imageUploadMode === "upload" ? (
          <div className="mb-3">
            <input
              type="file"
              accept="image/*"
              onChange={(e) => {
                const f = e.target.files?.[0] ?? null;
                setImageFile(f);
                if (f) {
                  const localUrl = URL.createObjectURL(f);
                  setImagePreviewUrl(localUrl);
                } else setImagePreviewUrl("");
              }}
            />
          </div>
        ) : (
          <div className="mb-3">
            <label className="block text-sm font-medium mb-1">Image URL</label>
            <input
              className="w-full border rounded px-2 py-2"
              placeholder="https://example.com/image.jpg"
              value={imageUrl}
              onChange={(e) => {
                setImageUrl(e.target.value);
                setImagePreviewUrl(e.target.value);
              }}
            />
          </div>
        )}

        <div className="grid grid-cols-2 gap-3 mb-3">
          <div>
            <label className="block text-sm font-medium mb-1">Width (px)</label>
            <input
              type="number"
              min={16}
              max={1024}
              className="w-full border rounded px-2 py-2"
              value={imageWidth}
              onChange={(e) => setImageWidth(Number(e.target.value || 64))}
            />
          </div>
          <div>
            <label className="block text-sm font-medium mb-1">
              Height (px)
            </label>
            <input
              type="number"
              min={16}
              max={1024}
              className="w-full border rounded px-2 py-2"
              value={imageHeight}
              onChange={(e) => setImageHeight(Number(e.target.value || 64))}
            />
          </div>
        </div>

        <div className="mb-3">
          <label className="block text-sm font-medium mb-1">
            Alt text (optional)
          </label>
          <input
            className="w-full border rounded px-2 py-2"
            placeholder="Description"
            value={imageAlt}
            onChange={(e) => setImageAlt(e.target.value)}
          />
        </div>

        {imagePreviewUrl ? (
          <div className="mb-4">
            <div className="text-xs text-gray-600 mb-1">Preview</div>
            <div
              className="border rounded p-2 flex items-center justify-center"
              style={{ maxHeight: 200 }}
            >
              <img
                src={imagePreviewUrl}
                alt={imageAlt || "preview"}
                style={{
                  maxWidth: "100%",
                  maxHeight: "180px",
                  objectFit: "contain",
                }}
              />
            </div>
          </div>
        ) : null}

        <div className="flex gap-2 justify-end">
          <button
            onClick={() => setInsertImageOpen(false)}
            className="px-3 py-1 rounded bg-gray-200 hover:bg-gray-300"
          >
            Cancel
          </button>
          <button
            onClick={async () => {
              if (
                !insertImageTarget?.id ||
                !insertImageTarget?.normKey ||
                !active
              )
                return;
              setUploadingImage(true);
              try {
                let finalUrl = "";
                let s3Key = "";
                if (imageUploadMode === "upload") {
                  if (!imageFile)
                    throw new Error("Please choose an image file.");
                  const contentType =
                    imageFile.type || "application/octet-stream";
                  const meta = {
                    category: active,
                    itemId: insertImageTarget.id,
                    column: insertImageTarget.normKey,
                  };
                  const presigned = await presignUpload(
                    imageFile.name,
                    contentType,
                    meta
                  );
                  const putRes = await fetch(presigned.uploadUrl, {
                    method: "PUT",
                    headers: { "Content-Type": contentType },
                    body: imageFile,
                  });
                  if (!putRes.ok) throw new Error(await putRes.text());
                  finalUrl = presigned.finalUrl || "";
                  s3Key = presigned.s3Key || "";
                } else {
                  if (!imageUrl || !/^https?:\/\//i.test(imageUrl))
                    throw new Error("Please enter a valid image URL.");
                  finalUrl = imageUrl.trim();
                }

                const payload = {
                  type: "image",
                  src: finalUrl,
                  alt: imageAlt || "",
                  width: Number(imageWidth || 64),
                  height: Number(imageHeight || 64),
                  ...(s3Key ? { s3Key } : {}),
                };
                const valueStr = JSON.stringify(payload);

                // Save to cell
                const res = await fetch(
                  `${API_BASE}/item/${encodeURIComponent(
                    insertImageTarget.id
                  )}`,
                  {
                    method: "PATCH",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({
                      attribute: insertImageTarget.normKey,
                      value: valueStr,
                      category: active,
                    }),
                  }
                );
                if (!res.ok) throw new Error(await res.text());
                const json = await res.json();
                const updated = json.updated || null;

                setItems((prev) =>
                  prev.map((it) =>
                    it.ID === insertImageTarget.id
                      ? {
                          ...it,
                          [insertImageTarget.normKey]: valueStr,
                          ...(updated || {}),
                        }
                      : it
                  )
                );
                setInsertImageOpen(false);
              } catch (err) {
                setModalError({
                  title: "Insert image failed",
                  message: err?.message || "Unable to insert image.",
                });
              } finally {
                setUploadingImage(false);
              }
            }}
            disabled={
              uploadingImage ||
              (imageUploadMode === "upload" && !imageFile) ||
              (imageUploadMode === "url" && !imageUrl)
            }
            className="px-3 py-1 rounded bg-indigo-600 text-white hover:bg-indigo-700"
          >
            {uploadingImage ? <Spinner className="h-4 w-4 text-white" /> : null}
            <span>{uploadingImage ? "Uploading…" : "Insert image"}</span>
          </button>
        </div>
      </div>
    </div>
  ) : null;

  // ----------------- Totals + export -----------------
  function computeTotalForItem(item, headerNorms) {
    const excluded = new Set([
      "id",
      "nombre",
      "nom",
      "fotos",
      "pierres",
      "piedras",
      "pictures",
      "photos",
    ]);
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
    const excluded = new Set([
      "id",
      "nom",
      "nombre",
      "pierres",
      "piedras",
      "total",
    ]);
    for (const nk of headerNorms) {
      const key = String(nk || "").toLowerCase();
      if (excluded.has(key)) {
        totals[key] = null;
        continue;
      }
      let anyNumeric = false;
      let sum = 0;
      for (const it of itemsArr) {
        const v = it[key];
        if (v === undefined || v === null || v === "") continue;
        const n = Number(String(v).replace(",", "."));
        if (Number.isFinite(n)) {
          sum += n;
          anyNumeric = true;
        }
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
  function buildRowsForExport(
    itemsArr,
    headerOrig,
    headerNorm,
    selection,
    rowSelectionMap,
    exportSelectedOnly,
    includeTotalsRow
  ) {
    let filteredItems = itemsArr;
    if (exportSelectedOnly && rowSelectionMap)
      filteredItems = itemsArr.filter((it) => !!rowSelectionMap[it.ID]);

    const rows = [];
    const headers = [];
    if (selection["id"]) headers.push("ID");
    for (let i = 0; i < headerOrig.length; i++) {
      const orig = headerOrig[i];
      const normKey = headerNorm[i];
      if (selection[normKey]) headers.push(orig);
    }
    if (selection["total"]) headers.push("Total");
    rows.push(headers);

    for (const it of filteredItems) {
      const r = [];
      if (selection["id"]) r.push(it.ID);
      for (let i = 0; i < headerOrig.length; i++) {
        const normKey = headerNorm[i];
        if (selection[normKey])
          r.push(
            it[normKey] === undefined || it[normKey] === null ? "" : it[normKey]
          );
      }
      if (selection["total"]) r.push(computeTotalForItem(it, headerNorm));
      rows.push(r);
    }

    if (includeTotalsRow) {
      const colTotals = computeColumnTotalsForItems(filteredItems, headerNorm);
      const totalsRow = [];
      if (selection["id"]) totalsRow.push("Totals");
      for (let i = 0; i < headerOrig.length; i++) {
        const normKey = headerNorm[i];
        if (selection[normKey]) {
          const val = colTotals[normKey];
          totalsRow.push(val === null || val === undefined ? "" : val);
        }
      }
      if (selection["total"]) {
        const totOfTotals = filteredItems.reduce(
          (acc, it) => acc + computeTotalForItem(it, headerNorm),
          0
        );
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
        const res = await fetch(
          `${API_BASE}/category/${encodeURIComponent(category)}`
        );
        if (!res.ok) throw new Error(await res.text());
        const json = await res.json();
        useItems = json.items || [];
        useSchema = json.schema || null;
      }
      const headerOrig = useSchema?.headerOriginalOrder?.length
        ? useSchema.headerOriginalOrder
        : useItems.length
        ? Object.keys(useItems[0]).filter((k) => k !== "category" && k !== "ID")
        : [];
      const headerNorm = headerOrig.map((h) => String(h).toLowerCase());
      const selection = getSelectionForCategory(category, headerNorm);
      const rowSelectionMap =
        (rowSelectionByCategory && rowSelectionByCategory[category]) || null;
      const exportSelectedOnly = !!exportSelectedRowsOnlyByCategory[category];
      const includeTotalsRow = !!includeTotalsRowByCategory[category];

      const rows = buildRowsForExport(
        useItems,
        headerOrig,
        headerNorm,
        selection,
        rowSelectionMap,
        exportSelectedOnly,
        includeTotalsRow
      );
      const ws = XLSX.utils.aoa_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(
        wb,
        ws,
        (category || "category").substring(0, 31)
      );
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
      setModalError({
        title: "Export failed",
        message: err?.message || "Unable to export category.",
      });
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
        const res = await fetch(
          `${API_BASE}/category/${encodeURIComponent(catName)}`
        );
        if (!res.ok) {
          console.warn(`Skipping category ${catName}: failed to fetch items`);
          continue;
        }
        const json = await res.json();
        const useItems = json.items || [];
        const useSchema = json.schema || null;
        const headerOrig = useSchema?.headerOriginalOrder?.length
          ? useSchema.headerOriginalOrder
          : useItems.length
          ? Object.keys(useItems[0]).filter(
              (k) => k !== "category" && k !== "ID"
            )
          : [];
        const headerNorm = headerOrig.map((h) => String(h).toLowerCase());
        const selection = getSelectionForCategory(catName, headerNorm);

        const rowSelectionMap =
          (rowSelectionByCategory && rowSelectionByCategory[catName]) || null;
        const exportSelectedOnly = !!exportSelectedRowsOnlyByCategory[catName];
        const includeTotalsRow = !!includeTotalsRowByCategory[catName];

        const rows = buildRowsForExport(
          useItems,
          headerOrig,
          headerNorm,
          selection,
          rowSelectionMap,
          exportSelectedOnly,
          includeTotalsRow
        );
        const ws = XLSX.utils.aoa_to_sheet(rows);
        XLSX.utils.book_append_sheet(
          wb,
          ws,
          (catName || "category").substring(0, 31)
        );
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
      setModalError({
        title: "Export failed",
        message: err?.message || "Unable to export all categories.",
      });
    } finally {
      setExportingAll(false);
    }
  }

  function headerCheckboxRow(selection, normArr) {
    return (
      <tr>
        <th className="p-1 border-b text-center">
          <input
            type="checkbox"
            checked={areAllRowsSelected()}
            onChange={() => setAllRowsSelectedForActive(!areAllRowsSelected())}
            className="cursor-pointer"
          />
        </th>
        <th className="p-1 border-b text-center">
          <input
            type="checkbox"
            checked={Boolean(selection["id"])}
            onChange={() => toggleColumnForActive("id")}
            className="cursor-pointer"
          />
        </th>
        {headerOriginalOrder().map((h, i) => {
          const protectedCol = isProtectedColumn(h);
          const normArrLocal = headerNormalizedOrder();
          return (
            <th
              key={`chk-${i}`}
              className="p-1 border-b text-center"
              onContextMenu={(e) => onHeaderCellContextMenu(e, i, h)}
            >
              <div className="flex items-center justify-center gap-2">
                <input
                  type="checkbox"
                  checked={Boolean(selection[normArrLocal[i]])}
                  onChange={() => toggleColumnForActive(normArrLocal[i])}
                  className="cursor-pointer"
                />
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    if (!protectedCol)
                      setPendingDeleteColumn({
                        columnOrig: h,
                        columnNorm: normArrLocal[i],
                      });
                  }}
                  title={
                    protectedCol
                      ? `Cannot delete protected column "${h}"`
                      : `Delete column "${h}"`
                  }
                  className={`ml-1 inline-flex items-center justify-center w-6 h-6 rounded-full ${
                    protectedCol
                      ? "text-gray-400 cursor-not-allowed"
                      : "text-red-600 hover:bg-red-100"
                  }`}
                  disabled={protectedCol || deletingColumn}
                  type="button"
                >
                  {deletingColumn &&
                  pendingDeleteColumn &&
                  pendingDeleteColumn.columnOrig === h ? (
                    <Spinner className="h-4 w-4 text-red-600" />
                  ) : (
                    <Trash2 className="h-4 w-4" />
                  )}
                </button>
              </div>
            </th>
          );
        })}
        <th className="p-1 border-b text-center">
          <input
            type="checkbox"
            checked={Boolean(selection["total"])}
            onChange={() => toggleColumnForActive("total")}
            className="cursor-pointer"
          />
        </th>
      </tr>
    );
  }

  function renderTable() {
    if (!items?.length)
      return <div className="text-sm text-gray-500">No items found</div>;
    const orig = headerOriginalOrder();
    const normArr = headerNormalizedOrder();
    const selection = getSelectionForCategory(active, normArr);
    const totals = computeColumnTotalsForItems(items, normArr);

    return (
      <div className="overflow-auto border border-gray-200 rounded shadow-sm relative">
        <table className="min-w-full table-fixed text-sm">
          <thead className="bg-gray-50 sticky top-0 z-10">
            {headerCheckboxRow(selection, normArr)}
            <tr>
              <th className="p-2 border-b border-r text-left w-12">Sel</th>
              <th
                className="p-2 border-b border-r text-left w-56"
                onContextMenu={(e) => e.preventDefault()}
              >
                ID
              </th>
              {orig.map((h, idx) => (
                <th
                  key={h}
                  className="p-2 border-b border-r text-left"
                  onContextMenu={(e) => onHeaderCellContextMenu(e, idx, h)}
                  title={h}
                >
                  {h}
                </th>
              ))}
              <th
                className="p-2 border-b text-left"
                onContextMenu={(e) =>
                  onHeaderCellContextMenu(e, orig.length, "Total")
                }
              >
                Total
              </th>
            </tr>
          </thead>

          <tbody>
            {items.map((it, rowIdx) => {
              const map =
                (rowSelectionByCategory && rowSelectionByCategory[active]) ||
                {};
              const selected = map[it.ID] === undefined ? true : !!map[it.ID];
              return (
                <tr
                  key={it.ID}
                  className="group odd:bg-white even:bg-gray-50 hover:bg-blue-50 transition-colors"
                >
                  <td className="p-2 border-r text-center">
                    <input
                      type="checkbox"
                      checked={selected}
                      onChange={() => toggleRowSelection(it.ID)}
                      className="cursor-pointer"
                    />
                  </td>

                  <td className="relative p-2 border-r align-top font-mono text-xs bg-gray-100 group-hover:bg-blue-50">
                    <div className="truncate">{it.ID}</div>
                  </td>

                  {orig.map((h, colIndex) => {
                    const nkey = normArr[colIndex] || String(h).toLowerCase();
                    const value = it[nkey];
                    const imgObj = parseImageValue(value);
                    const isEditing =
                      editingCell &&
                      editingCell.id === it.ID &&
                      editingCell.normKey === nkey;
                    const savingKey = `${it.ID}|${nkey}`;
                    return (
                      <td
                        key={nkey}
                        className="relative p-2 border-r align-top whitespace-nowrap overflow-hidden"
                        onDoubleClick={() => startEdit(it.ID, nkey, value)}
                        onContextMenu={(e) =>
                          onBodyCellContextMenu(e, rowIdx, it.ID, nkey)
                        }
                        title={h}
                      >
                        {!isEditing && (
                          <div className="text-xs leading-7 h-8 overflow-hidden truncate flex items-center gap-2">
                            {imgObj ? (
                              <>
                                <img
                                  src={imgObj.src}
                                  alt={imgObj.alt || h}
                                  style={{
                                    width: Math.min(64, imgObj.width || 64),
                                    height: Math.min(64, imgObj.height || 64),
                                    objectFit: "cover",
                                    borderRadius: 4,
                                  }}
                                />
                              </>
                            ) : savingCellKey === savingKey ? (
                              <span className="text-indigo-600">Saving…</span>
                            ) : value === undefined || value === null ? (
                              ""
                            ) : (
                              String(value)
                            )}
                          </div>
                        )}
                        {/* CHANGE: Always allow inline editing editor, even if the cell currently contains an image JSON */}
                        {isEditing && (
                          <input
                            autoFocus
                            value={cellValue}
                            onChange={(e) => setCellValue(e.target.value)}
                            onKeyDown={(e) => onInputKeyDown(e, it.ID, nkey)}
                            onBlur={() => {
                              const original = editingCell
                                ? editingCell.original == null
                                  ? ""
                                  : String(editingCell.original)
                                : "";
                              const incoming =
                                cellValue == null ? "" : String(cellValue);
                              if (incoming === original) {
                                cancelEdit();
                              } else {
                                // Save plain text to the cell (overwrites image JSON if present)
                                saveEdit(it.ID, nkey, cellValue);
                              }
                            }}
                            className="absolute inset-0 w-full h-full px-1 text-sm bg-white focus:outline-none box-border"
                            style={{
                              padding: "4px 6px",
                              border: "1px solid rgba(0,0,0,0.08)",
                              borderRadius: 4,
                            }}
                          />
                        )}
                      </td>
                    );
                  })}

                  <td className="p-2 align-top text-sm">
                    {computeTotalForItem(it, normArr)}
                  </td>
                </tr>
              );
            })}
          </tbody>

          <tfoot className="bg-gray-100">
            <tr>
              <td className="p-2 border-t text-left font-medium">Totals</td>
              <td className="p-2 border-t border-r font-medium"></td>
              {orig.map((h, i) => {
                const nkey = normArr[i] || String(h).toLowerCase();
                const v = totals[nkey];
                return (
                  <td
                    key={`tot-${nkey}`}
                    className="p-2 border-t border-r text-sm font-semibold"
                  >
                    {v == null ? "" : String(v)}
                  </td>
                );
              })}
              <td className="p-2 border-t text-sm font-semibold"></td>
            </tr>
          </tfoot>
        </table>

        {/* Row actions */}
        <div className="p-3 flex items-center gap-3">
          <button
            onClick={onRequestDeleteSelectedRows}
            disabled={!anyRowSelected()}
            className={`px-3 py-1 rounded bg-red-600 text-white ${
              !anyRowSelected()
                ? "opacity-60 cursor-not-allowed"
                : "hover:bg-red-700"
            }`}
          >
            Delete selected rows
          </button>
        </div>

        {/* Context menus */}
        <ColumnContextMenu />
        <RowContextMenu />

        {/* Modals */}
        {InsertImageModal}
      </div>
    );
  }

  // Editing save helper
  async function saveEdit(id, normKey, newValue) {
    const original = editingCell
      ? editingCell.original == null
        ? ""
        : String(editingCell.original)
      : "";
    const incoming = newValue == null ? "" : String(newValue);
    if (incoming === original) {
      cancelEdit();
      return;
    }
    const savingKeyLocal = `${id}|${normKey}`;
    setSavingCellKey(savingKeyLocal);
    try {
      const body = { attribute: normKey, value: incoming, category: active };
      const res = await fetch(`${API_BASE}/item/${encodeURIComponent(id)}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
      if (!res.ok) {
        const txt = await res.text().catch(() => "");
        throw new Error(txt || `Server ${res.status}`);
      }
      const json = await res.json();
      setItems((prev) =>
        prev.map((it) =>
          it.ID === id
            ? { ...it, [normKey]: incoming, ...(json.updated || {}) }
            : it
        )
      );
      cancelEdit();
    } catch (err) {
      setModalError({
        title: "Save failed",
        message: err?.message || "Unable to save change.",
      });
    } finally {
      setSavingCellKey(null);
    }
  }
  function startEdit(id, normKey, initial) {
    setEditingCell({
      id,
      normKey,
      original: initial == null ? "" : String(initial),
    });
    setCellValue(initial == null ? "" : String(initial));
  }
  function cancelEdit() {
    setEditingCell(null);
    setCellValue("");
  }
  function onInputKeyDown(e, id, normKey) {
    if (e.key === "Enter") {
      e.preventDefault();
      saveEdit(id, normKey, cellValue);
    } else if (e.key === "Escape") cancelEdit();
  }

  // ----------------- Small header controls -----------------
  function headerControls() {
    return (
      <div className="flex gap-2 items-center">
        <button
          onClick={exportAllWithSelections}
          className={`px-3 py-1 bg-indigo-600 text-white rounded flex items-center gap-2 ${
            exportingAll ? "opacity-70 cursor-wait" : "hover:bg-indigo-700"
          }`}
          disabled={
            exportingAll ||
            exportingCategory ||
            deletingColumn ||
            deletingRows ||
            replacing
          }
        >
          {exportingAll ? <Spinner className="h-4 w-4 text-white" /> : null}
          <span>{exportingAll ? "Preparing…" : "Download All"}</span>
        </button>

        <button
          onClick={() => active && exportCategoryWithSelection(active)}
          disabled={
            !active ||
            exportingCategory ||
            exportingAll ||
            deletingColumn ||
            deletingRows ||
            replacing
          }
          className={`px-3 py-1 bg-green-600 text-white rounded flex items-center gap-2 ${
            !active || exportingCategory
              ? "opacity-70 cursor-not-allowed"
              : "hover:bg-green-700"
          }`}
        >
          {exportingCategory ? (
            <Spinner className="h-4 w-4 text-white" />
          ) : null}
          <span>{exportingCategory ? "Preparing…" : "Download Category"}</span>
        </button>

        <button
          onClick={onReplaceAllClick}
          disabled={!active || replacing || deletingColumn || deletingRows}
          className={`px-3 py-1 bg-red-600 text-white rounded ${
            !active || replacing
              ? "opacity-70 cursor-not-allowed"
              : "hover:bg-red-700"
          }`}
        >
          {replacing ? <Spinner className="h-4 w-4 text-white" /> : null}
          <span>{replacing ? "Replacing…" : "Replace Category"}</span>
        </button>

        {/* Quick add column at end */}
        <button
          onClick={() => openInsertColumnAt(headerOriginalOrder().length)}
          className="px-3 py-1 bg-indigo-600 text-white rounded flex items-center gap-2 hover:bg-indigo-700"
          disabled={!active}
        >
          <Plus className="h-4 w-4" />
          <span>Add column (end)</span>
        </button>

        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx,.xls"
          className="hidden"
          onChange={(e) => {
            const f = e.target.files?.[0] ?? null;
            if (!f) return;
            setPendingReplaceFile(f);
            setConfirmReplaceOpen(true);
            e.target.value = "";
          }}
        />
      </div>
    );
  }
  function onReplaceAllClick() {
    if (!active) {
      setModalError({
        title: "No category selected",
        message: "Select a category before replacing.",
      });
      return;
    }
    fileInputRef.current?.click();
  }

  // ----------------- Modals -----------------
  const InsertColumnModal = insertColumnOpen ? (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/40"
      onClick={() => !insertingColumn && setInsertColumnOpen(false)}
    >
      <div
        className="bg-white rounded-lg shadow-lg p-6 max-w-md w-full"
        onClick={(e) => e.stopPropagation()}
      >
        <h3 className="text-lg font-semibold mb-2">Insert column</h3>
        <p className="text-sm text-gray-700 mb-4">
          Add a new column at position {insertColumnIndex} for category{" "}
          <strong>{active}</strong>.
        </p>
        <label className="block text-sm font-medium mb-1">Column name</label>
        <input
          value={newColumnName}
          onChange={(e) => setNewColumnName(e.target.value)}
          className="w-full border rounded px-2 py-2 mb-3"
          placeholder="e.g., NewQuantity"
        />
        <div className="flex gap-2 justify-end">
          <button
            onClick={() => setInsertColumnOpen(false)}
            className="px-3 py-1 rounded bg-gray-200 hover:bg-gray-300"
          >
            Cancel
          </button>
          <button
            onClick={confirmInsertColumn}
            disabled={insertingColumn}
            className="px-3 py-1 rounded bg-indigo-600 text-white flex items-center gap-2 hover:bg-indigo-700"
          >
            {insertingColumn ? (
              <Spinner className="h-4 w-4 text-white" />
            ) : null}
            <span>{insertingColumn ? "Adding…" : "Insert column"}</span>
          </button>
        </div>
      </div>
    </div>
  ) : null;

  const InsertRowModal = insertRowOpen ? (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/40"
      onClick={() => !insertingRow && setInsertRowOpen(false)}
    >
      <div
        className="bg-white rounded-lg shadow-lg p-6 max-w-xl w-full"
        onClick={(e) => e.stopPropagation()}
      >
        <h3 className="text-lg font-semibold mb-2">Insert row</h3>
        <p className="text-sm text-gray-700 mb-4">
          Add a new row at position {insertRowIndex} for category{" "}
          <strong>{active}</strong>.
        </p>

        {/* ID field */}
        <label className="block text-sm font-medium mb-1">ID</label>
        <div className="flex items-center gap-2 mb-3">
          <input
            value={newRowId}
            onChange={(e) => {
              const v = e.target.value;
              setNewRowId(v);
              checkIdExists(v);
            }}
            placeholder="Enter a unique ID (UUID pre-filled)"
            className="flex-1 border rounded px-2 py-2"
          />
          <button
            type="button"
            onClick={() => {
              const v = uuidv4();
              setNewRowId(v);
              checkIdExists(v);
            }}
            className="px-2 py-2 bg-gray-200 rounded hover:bg-gray-300 text-sm"
            title="Generate UUID"
          >
            Generate
          </button>
        </div>
        <div className="text-xs mb-4">
          {newRowIdChecking ? (
            <span className="text-gray-600">Checking ID…</span>
          ) : newRowIdExists ? (
            <span className="text-red-600">
              This ID already exists; choose a different ID.
            </span>
          ) : newRowId ? (
            <span className="text-green-700">ID available</span>
          ) : (
            <span className="text-gray-600">ID is required</span>
          )}
        </div>

        {/* Field inputs */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mb-4">
          {headerNormalizedOrder()
            .filter((nk) => !["total", "id", "category"].includes(nk))
            .map((nk) => {
              const idx = headerNormalizedOrder().indexOf(nk);
              const label = headerOriginalOrder()[idx] || nk;
              return (
                <div key={nk}>
                  <label className="block text-xs font-medium mb-1">
                    {label}
                  </label>
                  <input
                    value={newRowValues[nk] ?? ""}
                    onChange={(e) =>
                      setNewRowValues((prev) => ({
                        ...prev,
                        [nk]: e.target.value,
                      }))
                    }
                    className="w-full border rounded px-2 py-2"
                    placeholder={`Enter ${label}`}
                  />
                </div>
              );
            })}
        </div>

        <div className="flex gap-2 justify-end">
          <button
            onClick={() => setInsertRowOpen(false)}
            className="px-3 py-1 rounded bg-gray-200 hover:bg-gray-300"
          >
            Cancel
          </button>
          <button
            onClick={confirmInsertRow}
            disabled={insertingRow || !newRowId || newRowIdExists}
            className="px-3 py-1 rounded bg-indigo-600 text-white flex items-center gap-2 hover:bg-indigo-700 disabled:opacity-70"
          >
            {insertingRow ? <Spinner className="h-4 w-4 text-white" /> : null}
            <span>{insertingRow ? "Adding…" : "Insert row"}</span>
          </button>
        </div>
      </div>
    </div>
  ) : null;

  // ----------------- JSX -----------------
  return (
    <div className="max-w-7xl mx-auto p-4">
      <header className="mb-6 flex items-start justify-between">
        <div>
          <h1 className="text-2xl font-semibold">Inventory Viewer</h1>
          <p className="text-sm text-gray-500 mt-1">
            Right-click headers to Add column (left/right). Right-click a data
            cell to Add row (top/bottom), Delete row, or Insert image.
          </p>
        </div>
        {headerControls()}
      </header>

      <div className="mb-4 flex items-center gap-3">
        <div className="font-medium">Categories:</div>
        {loadingCats ? (
          <div className="text-sm text-gray-500">Loading...</div>
        ) : (
          <div className="flex gap-3 overflow-auto">
            {categories.map((c) => (
              <div
                key={c.name}
                className={`flex items-center justify-between gap-2 px-3 py-1 rounded ${
                  active === c.name
                    ? "bg-blue-600 text-white"
                    : "bg-gray-100 text-gray-800"
                } hover:shadow-sm`}
                style={{ minWidth: 160, maxWidth: 320 }}
              >
                <button
                  onClick={() => setActive(c.name)}
                  className="flex-1 text-left truncate"
                  title={`Select category ${c.name}`}
                >
                  {c.name}
                </button>
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    setPendingDeleteCategory(c.name);
                  }}
                  className="ml-2 inline-flex items-center justify-center w-8 h-8 rounded-full bg-red-100 hover:bg-red-200"
                  title={`Delete category ${c.name}`}
                >
                  <Trash2 className="h-4 w-4 text-red-600" />
                </button>
              </div>
            ))}
          </div>
        )}
      </div>

      {error && <div className="mb-4 text-sm text-red-600">{error}</div>}

      <main>
        <div className="mb-4 relative">
          <h2 className="text-lg font-medium mb-2">
            {active || "No category selected"}
          </h2>
          {loadingItems ? (
            <div className="text-sm text-gray-500">Loading items...</div>
          ) : (
            renderTable()
          )}

          {/* Modals */}
          {InsertColumnModal}
          {InsertRowModal}
          {InsertImageModal}

          {/* Context menus */}
          <ColumnContextMenu />
          <RowContextMenu />
        </div>
      </main>

      {/* Delete Column Modal */}
      {pendingDeleteColumn && (
        <div className="fixed inset-0 z-40 flex items-center justify-center bg-black/40">
          <div className="bg-white rounded-lg shadow-lg p-6 max-w-md w-full">
            <h3 className="text-lg font-semibold mb-2">
              Delete column "{pendingDeleteColumn.columnOrig}"?
            </h3>
            <p className="text-sm text-gray-700 mb-4">
              This will remove the column{" "}
              <strong>{pendingDeleteColumn.columnOrig}</strong> from all items
              in category <strong>{active}</strong>. This cannot be undone.
            </p>
            <div className="flex gap-2 justify-end">
              <button
                onClick={() => setPendingDeleteColumn(null)}
                className="px-3 py-1 rounded bg-gray-200 hover:bg-gray-300"
              >
                Cancel
              </button>
              <button
                onClick={confirmDeleteColumn}
                className="px-3 py-1 rounded bg-red-600 text-white flex items-center gap-2 hover:bg-red-700"
                disabled={deletingColumn}
              >
                {deletingColumn ? (
                  <Spinner className="h-4 w-4 text-white" />
                ) : null}
                <span>{deletingColumn ? "Deleting…" : "Confirm Delete"}</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Bulk Delete Rows Modal */}
      {confirmDeleteRowsOpen && (
        <div className="fixed inset-0 z-40 flex items-center justify-center bg-black/40">
          <div className="bg-white rounded-lg shadow-lg p-6 max-w-md w-full">
            <h3 className="text-lg font-semibold mb-2">
              Delete selected rows?
            </h3>
            <p className="text-sm text-gray-700 mb-4">
              This will permanently delete {pendingDeleteRowIds.length} rows
              from category <strong>{active}</strong>. The first data row and
              the totals row cannot be deleted.
            </p>
            <div className="flex gap-2 justify-end">
              <button
                onClick={() => {
                  setConfirmDeleteRowsOpen(false);
                  setPendingDeleteRowIds([]);
                }}
                className="px-3 py-1 rounded bg-gray-200 hover:bg-gray-300"
              >
                Cancel
              </button>
              <button
                onClick={confirmDeleteSelectedRows}
                disabled={deletingRows}
                className="px-3 py-1 rounded bg-red-600 text-white flex items-center gap-2 hover:bg-red-700"
              >
                {deletingRows ? (
                  <Spinner className="h-4 w-4 text-white" />
                ) : null}
                <span>
                  {deletingRows
                    ? "Deleting…"
                    : `Delete ${pendingDeleteRowIds.length} rows`}
                </span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Error modal */}
      {modalError && (
        <div
          className="fixed inset-0 z-40 flex items-center justify-center bg-black/40"
          onClick={() => setModalError(null)}
        >
          <div
            className="bg-white rounded-lg shadow-lg p-4 max-w-lg w-full"
            onClick={(e) => e.stopPropagation()}
          >
            <h3 className="font-semibold text-lg mb-2">{modalError.title}</h3>
            <div className="text-sm text-gray-700 mb-4">
              {modalError.message}
            </div>
            <div className="flex justify-end">
              <button
                onClick={() => setModalError(null)}
                className="px-3 py-1 rounded bg-gray-200 hover:bg-gray-300"
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}