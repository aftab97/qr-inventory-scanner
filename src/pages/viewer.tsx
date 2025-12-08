import React, {
  useEffect,
  useRef,
  useState,
  useMemo,
  useCallback,
} from "react";
import {
  Trash2,
  Plus,
  Image as ImageIcon,
  Minus,
  Plus as PlusIcon,
} from "lucide-react";
// You can remove SheetJS entirely if not used elsewhere
// import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { isLocalHost } from "../api";

/* Performance refactor v2 (height scaling via transform):
   - Column width: smooth via CSS variables and requestAnimationFrame (as before).
   - Image column height: animated via transform: scaleY on a fixed-height container (GPU-friendly, avoids layout reflow).
   - Debounced commit: after interaction, set the column's base height CSS var to the target height and reset scale to 1; persist to React state/localStorage without jank.
   - Images auto-fit (width/height "auto"); <img> uses decoding="async" and loading="lazy".
   - Existing functionality preserved (context menus, insert/delete, export, etc.).
*/

const API_BASE = isLocalHost
  ? "http://localhost:8080"
  : "https://qr-inventory-scanner-backend.vercel.app";

// LocalStorage keys
const LS_ROWS_KEY = "qr_viewer_row_selection_v1";
const LS_COLS_KEY = "qr_viewer_col_selection_v1";
const LS_COL_WIDTHS_KEY = "qr_viewer_col_widths_v2";
const LS_ROW_HEIGHTS_KEY = "qr_viewer_row_heights_v2";
const LS_IMAGE_COL_HEIGHTS_KEY = "qr_viewer_image_col_heights_v1";

// Sizing defaults
const SIZING = {
  colMin: 60,
  colMax: 1600,
  rowMin: 28,
  rowMax: 1200,
  defaults: {
    sel: 48,
    id: 224,
    data: 180,
    total: 140,
    row: 56,
    imageColHeight: 128,
  },
};
const CELL_PADDING = 8;
const STEP = { width: 40, height: 40 };

// UUID v4
function uuidv4() {
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
    const r = crypto.getRandomValues(new Uint8Array(1))[0] & 0xf;
    const v = c === "x" ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });
}

// Parse image JSON
type ImageCell = {
  type: "image";
  src: string;
  alt?: string;
  width?: string;
  height?: string;
  s3Key?: string;
};
function parseImageValue(val: unknown): ImageCell | null {
  if (!val || typeof val !== "string") return null;
  try {
    const obj = JSON.parse(val);
    if (obj && obj.type === "image" && typeof obj.src === "string") return obj;
  } catch {}
  return null;
}

// Utils
function norm(s: unknown) {
  return String(s ?? "")
    .trim()
    .toLowerCase();
}

// Debounce helper
function useDebouncedCallback<T extends (...args: any[]) => void>(
  fn: T,
  delay: number
) {
  const timerRef = useRef<number | null>(null);
  return useCallback(
    (...args: Parameters<T>) => {
      if (timerRef.current) window.clearTimeout(timerRef.current);
      timerRef.current = window.setTimeout(() => {
        timerRef.current = null;
        fn(...args);
      }, delay);
    },
    [fn, delay]
  );
}

// Numeric coercion for Excel cells
function toNumberIfNumeric(val: unknown): number | string | null {
  if (val === null || val === undefined) return null;
  if (typeof val === "number" && Number.isFinite(val)) return val;
  const s = String(val).trim();
  if (s === "") return "";
  // Allow comma decimal separator (e.g., "12,34")
  const normalized = s.replace(",", ".");
  const n = Number(normalized);
  return Number.isFinite(n) ? n : s;
}

// Pixels per inch and min image width rules
const PX_PER_INCH = 96;
const MIN_IMAGE_WIDTH_INCHES = 2.0;
const MIN_IMAGE_WIDTH_PX = Math.round(MIN_IMAGE_WIDTH_INCHES * PX_PER_INCH);

// Excel column width conversions (~7.5 px per "character" width)
function charsToPx(chars: number) {
  return Math.round((chars || 10) * 7.5);
}
function pxToChars(px: number) {
  return Math.max(6, Math.round(px / 7.5));
}

// Decode image intrinsic dimensions using createImageBitmap for aspect ratio
async function getImageDimensionsFromBuffer(
  ab: ArrayBuffer
): Promise<{ width: number; height: number } | null> {
  try {
    const blob = new Blob([ab]);
    const bitmap = await createImageBitmap(blob);
    const dims = { width: bitmap.width, height: bitmap.height };
    bitmap.close?.();
    return dims;
  } catch (e) {
    console.warn("getImageDimensionsFromBuffer failed", e);
    return null;
  }
}

export default function Viewer() {
  // Data
  const [categories, setCategories] = useState<Array<{ name: string }>>([]);
  const [active, setActive] = useState<string | null>(null);
  const [items, setItems] = useState<any[]>([]);
  const [schema, setSchema] = useState<any | null>(null);

  // UI flags
  const [loadingCats, setLoadingCats] = useState(false);
  const [loadingItems, setLoadingItems] = useState(false);
  const [replacing, setReplacing] = useState(false);
  const [deletingColumn, setDeletingColumn] = useState(false);
  const [deletingRows, setDeletingRows] = useState(false);
  const [deletingCategory, setDeletingCategory] = useState<string | false>(
    false
  );
  const [exportingCategory, setExportingCategory] = useState(false);
  const [exportingAll, setExportingAll] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [modalError, setModalError] = useState<{
    title: string;
    message: string;
  } | null>(null);

  // Editing state
  const [editingCell, setEditingCell] = useState<{
    id: string;
    normKey: string;
    original: string;
  } | null>(null);
  const [cellValue, setCellValue] = useState("");
  const [savingCellKey, setSavingCellKey] = useState<string | null>(null);

  // File/replace
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [pendingReplaceFile, setPendingReplaceFile] = useState<File | null>(
    null
  );
  const [confirmReplaceOpen, setConfirmReplaceOpen] = useState(false);

  // Column/Row delete modals
  const [pendingDeleteColumn, setPendingDeleteColumn] = useState<{
    columnOrig: string;
    columnNorm: string;
  } | null>(null);
  const [pendingDeleteRowIds, setPendingDeleteRowIds] = useState<string[]>([]);
  const [confirmDeleteRowsOpen, setConfirmDeleteRowsOpen] = useState(false);

  // Insert column/row modals
  const [insertColumnOpen, setInsertColumnOpen] = useState(false);
  const [newColumnName, setNewColumnName] = useState("");
  const [insertColumnIndex, setInsertColumnIndex] = useState(0);
  const [insertingColumn, setInsertingColumn] = useState(false);

  const [insertRowOpen, setInsertRowOpen] = useState(false);
  const [insertRowIndex, setInsertRowIndex] = useState(0);
  const [insertingRow, setInsertingRow] = useState(false);
  const [newRowId, setNewRowId] = useState("");
  const [newRowValues, setNewRowValues] = useState<Record<string, string>>({});
  const [newRowIdChecking, setNewRowIdChecking] = useState(false);
  const [newRowIdExists, setNewRowIdExists] = useState(false);

  // Insert image modal
  const [insertImageOpen, setInsertImageOpen] = useState(false);
  const [insertImageTarget, setInsertImageTarget] = useState<{
    id: string;
    normKey: string;
  } | null>(null); // { id, normKey }
  const [imageUploadMode, setImageUploadMode] = useState<"upload" | "url">(
    "upload"
  );
  const [imageFile, setImageFile] = useState<File | null>(null);
  const [imageUrl, setImageUrl] = useState("");
  const [imageAlt, setImageAlt] = useState("");
  const [uploadingImage, setUploadingImage] = useState(false);
  const [imagePreviewUrl, setImagePreviewUrl] = useState("");
  const [imageWidth, setImageWidth] = useState(64);
  const [imageHeight, setImageHeight] = useState(64);

  // Image columns autodetect (typed)
  const imageColumns = useMemo<Set<string>>(() => {
    const norm = headerNormalizedOrder();
    const set = new Set<string>();
    for (const nk of norm) {
      for (const it of items || []) {
        const img = parseImageValue(it[nk]);
        if (img) {
          set.add(nk);
          break;
        }
      }
    }
    return set;
  }, [items, schema]);

  // Context menu
  const [ctxMenu, setCtxMenu] = useState<{
    open: boolean;
    x: number;
    y: number;
    type: "column" | "row" | null;
    columnIndex: number | null;
    columnName: string | null;
    rowIndex: number | null;
    rowId: string | null;
    isHeaderCell: boolean;
    cellNormKey: string | null;
  }>({
    open: false,
    x: 0,
    y: 0,
    type: null,
    columnIndex: null,
    columnName: null,
    rowIndex: null,
    rowId: null,
    isHeaderCell: false,
    cellNormKey: null,
  });

  // Selections
  const [columnSelectionByCategory, setColumnSelectionByCategory] = useState<
    Record<string, Record<string, boolean>>
  >(() => {
    try {
      const raw = window.localStorage.getItem(LS_COLS_KEY);
      return raw ? JSON.parse(raw) : {};
    } catch {
      return {};
    }
  });
  const [rowSelectionByCategory, setRowSelectionByCategory] = useState<
    Record<string, Record<string, boolean>>
  >(() => {
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
  ] = useState<Record<string, boolean>>({});
  const [includeTotalsRowByCategory, setIncludeTotalsRowByCategory] = useState<
    Record<string, boolean>
  >({});

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

  // Column widths and row heights (persisted)
  const [colWidthsByCategory, setColWidthsByCategory] = useState<
    Record<string, number[]>
  >(() => {
    try {
      const raw = window.localStorage.getItem(LS_COL_WIDTHS_KEY);
      return raw ? JSON.parse(raw) : {};
    } catch {
      return {};
    }
  });
  const [rowHeightsByCategory, setRowHeightsByCategory] = useState<
    Record<string, Record<string, number>>
  >(() => {
    try {
      const raw = window.localStorage.getItem(LS_ROW_HEIGHTS_KEY);
      return raw ? JSON.parse(raw) : {};
    } catch {
      return {};
    }
  });
  const [imageColHeightsByCategory, setImageColHeightsByCategory] = useState<
    Record<string, Record<string, number>>
  >(() => {
    try {
      const raw = window.localStorage.getItem(LS_IMAGE_COL_HEIGHTS_KEY);
      return raw ? JSON.parse(raw) : {};
    } catch {
      return {};
    }
  });

  const persistColWidthsDebounced = useDebouncedCallback(
    (next: Record<string, number[]>) => {
      try {
        window.localStorage.setItem(LS_COL_WIDTHS_KEY, JSON.stringify(next));
      } catch {}
    },
    200
  );

  const persistRowHeightsDebounced = useDebouncedCallback(
    (next: Record<string, Record<string, number>>) => {
      try {
        window.localStorage.setItem(LS_ROW_HEIGHTS_KEY, JSON.stringify(next));
      } catch {}
    },
    200
  );

  const persistImageColHeightsDebounced = useDebouncedCallback(
    (next: Record<string, Record<string, number>>) => {
      try {
        window.localStorage.setItem(
          LS_IMAGE_COL_HEIGHTS_KEY,
          JSON.stringify(next)
        );
      } catch {}
    },
    200
  );

  useEffect(() => {
    persistColWidthsDebounced(colWidthsByCategory);
  }, [colWidthsByCategory, persistColWidthsDebounced]);
  useEffect(() => {
    persistRowHeightsDebounced(rowHeightsByCategory);
  }, [rowHeightsByCategory, persistRowHeightsDebounced]);
  useEffect(() => {
    persistImageColHeightsDebounced(imageColHeightsByCategory);
  }, [imageColHeightsByCategory, persistImageColHeightsDebounced]);

  // Helpers: header orders
  function headerOriginalOrder(): string[] {
    if (schema?.headerOriginalOrder?.length) return schema.headerOriginalOrder;
    if (!items?.length) return [];
    return Object.keys(items[0]).filter((k) => k !== "category" && k !== "ID");
  }
  function headerNormalizedOrder(): string[] {
    if (schema?.headerNormalizedOrder?.length)
      return schema.headerNormalizedOrder;
    return headerOriginalOrder().map((h) => String(h).toLowerCase());
  }
  function getVisualColumnCount(): number {
    return 1 + 1 + headerOriginalOrder().length + 1; // Sel + ID + data + Total (UI only)
  }

  // Ensure widths array shape
  function ensureColWidthsForActive() {
    if (!active) return;
    const count = getVisualColumnCount();
    setColWidthsByCategory((prev) => {
      const existing = prev[active];
      if (existing && existing.length === count) return prev;
      const nextArr = new Array<number>(count);
      let i = 0;
      nextArr[i++] = SIZING.defaults.sel;
      nextArr[i++] = SIZING.defaults.id;
      const dataCount = headerOriginalOrder().length;
      for (let d = 0; d < dataCount; d++) nextArr[i++] = SIZING.defaults.data;
      nextArr[i++] = SIZING.defaults.total;
      if (existing && existing.length) {
        for (let j = 0; j < Math.min(existing.length, count); j++) {
          if (typeof existing[j] === "number" && existing[j] > 0)
            nextArr[j] = existing[j];
        }
      }
      return { ...prev, [active]: nextArr };
    });
  }
  function getColWidths(): number[] {
    if (!active) return [];
    const arr = colWidthsByCategory[active];
    if (!arr || arr.length !== getVisualColumnCount()) return [];
    return arr;
  }
  function setColWidthAt(index: number, width: number) {
    if (!active) return;
    const clamped = Math.max(
      SIZING.colMin,
      Math.min(SIZING.colMax, Math.round(width))
    );
    setColWidthsByCategory((prev) => {
      const arr = (prev[active] || []).slice();
      arr[index] = clamped;
      return { ...prev, [active]: arr };
    });
    // Also set CSS variable for immediate visual update
    setCssVar(`--col-${index}-w`, `${clamped}px`);
  }
  function getRowHeightById(id: string): number {
    if (!active || !id) return SIZING.defaults.row;
    const map = rowHeightsByCategory[active] || {};
    return Math.max(
      SIZING.rowMin,
      Math.min(SIZING.rowMax, Number(map[id] || SIZING.defaults.row))
    );
  }
  function setRowHeightById(id: string, height: number) {
    if (!active || !id) return;
    const clamped = Math.max(
      SIZING.rowMin,
      Math.min(SIZING.rowMax, Math.round(height))
    );
    setRowHeightsByCategory((prev) => {
      const map = { ...(prev[active] || {}) };
      map[id] = clamped;
      return { ...prev, [active]: map };
    });
  }

  // Image column height: base height from state; transform scaleY for live updates
  function getImageColHeight(normKey: string): number {
    if (!active || !normKey) return SIZING.defaults.imageColHeight;
    const perCat = imageColHeightsByCategory[active] || {};
    const h = perCat[normKey];
    return Math.max(
      SIZING.rowMin,
      Math.min(SIZING.rowMax, Number(h || SIZING.defaults.imageColHeight))
    );
  }
  const persistImgColHeightDebounced = useDebouncedCallback(
    (normKey: string, height: number) => {
      if (!active || !normKey) return;
      const clamped = Math.max(
        SIZING.rowMin,
        Math.min(SIZING.rowMax, Math.round(height))
      );
      setImageColHeightsByCategory((prev) => {
        const perCat = { ...(prev[active] || {}) };
        perCat[normKey] = clamped;
        return { ...prev, [active]: perCat };
      });
    },
    120
  );
  function setImageColHeightCssVar(normKey: string, heightPx: number) {
    const clamped = Math.max(
      SIZING.rowMin,
      Math.min(SIZING.rowMax, Math.round(heightPx))
    );
    setCssVar(`--imgcol-${normKey}-h`, `${clamped}px`);
  }

  // Scale vars for image columns
  const imgColScaleRef = useRef<Record<string, number>>({});
  function setImageColScaleCssVar(normKey: string, scale: number) {
    const s = Math.max(0.1, Math.min(10, scale)); // clamp
    imgColScaleRef.current[normKey] = s;
    setCssVar(`--imgcol-${normKey}-scale`, String(s));
  }
  const commitImgColHeightDebounced = useDebouncedCallback(
    (normKey: string, targetHeightPx: number) => {
      setImageColHeightCssVar(normKey, targetHeightPx);
      setImageColScaleCssVar(normKey, 1);
      persistImgColHeightDebounced(normKey, targetHeightPx);
    },
    120
  );

  // requestAnimationFrame scheduling for smooth CSS var updates
  const rafIdRef = useRef<number | null>(null);
  const pendingUpdateRef = useRef<Record<string, string>>({});

  function setCssVar(name: string, value: string) {
    pendingUpdateRef.current[name] = value;
    if (rafIdRef.current) return;
    rafIdRef.current = window.requestAnimationFrame(() => {
      const entries = Object.entries(pendingUpdateRef.current);
      entries.forEach(([n, v]) => {
        document.documentElement.style.setProperty(n, v);
      });
      pendingUpdateRef.current = {};
      rafIdRef.current = null;
    });
  }

  // Drag state (column/row)
  const [draggingCol, setDraggingCol] = useState<{
    index: number;
    startX: number;
    startWidth: number;
  } | null>(null);
  const [draggingRow, setDraggingRow] = useState<{
    id: string;
    startY: number;
    startHeight: number;
  } | null>(null);

  useEffect(() => {
    function onMove(e: MouseEvent) {
      if (draggingCol) {
        const delta = e.clientX - draggingCol.startX;
        const next = draggingCol.startWidth + delta;
        const clamped = Math.max(SIZING.colMin, Math.min(SIZING.colMax, next));
        // CSS var only during drag
        setCssVar(`--col-${draggingCol.index}-w`, `${clamped}px`);
      } else if (draggingRow) {
        const delta = e.clientY - draggingRow.startY;
        const next = draggingRow.startHeight + delta;
        const clamped = Math.max(SIZING.rowMin, Math.min(SIZING.rowMax, next));
        // Update row height state (affects only that row)
        setRowHeightById(draggingRow.id, clamped);
      }
    }
    function onUp() {
      if (draggingCol) {
        // Persist final column width to React state
        const computed = getComputedStyle(document.documentElement)
          .getPropertyValue(`--col-${draggingCol.index}-w`)
          .trim();
        const val = parseInt(computed || `${draggingCol.startWidth}`, 10);
        setColWidthAt(draggingCol.index, val || draggingCol.startWidth);
      }
      setDraggingCol(null);
      setDraggingRow(null);
    }
    if (draggingCol || draggingRow) {
      window.addEventListener("mousemove", onMove);
      window.addEventListener("mouseup", onUp);
      return () => {
        window.removeEventListener("mousemove", onMove);
        window.removeEventListener("mouseup", onUp);
      };
    }
  }, [draggingCol, draggingRow]);

  // Context menu global close
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
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

  // Load categories/items
  useEffect(() => {
    loadCategories();
  }, []);
  useEffect(() => {
    if (active) loadCategoryItems(active);
  }, [active]);

  useEffect(() => {
    ensureColWidthsForActive();
    // Initialize CSS vars
    const widths = getColWidths();
    widths.forEach((w, i) => setCssVar(`--col-${i}-w`, `${w}px`));
    headerNormalizedOrder().forEach((nk) => {
      setCssVar(`--imgcol-${nk}-h`, `${getImageColHeight(nk)}px`);
      setCssVar(`--imgcol-${nk}-scale`, "1");
    });
  }, [active, schema, items]);

  // Backend calls
  async function loadCategories() {
    setLoadingCats(true);
    setError(null);
    try {
      const res = await fetch(`${API_BASE}/categories`);
      if (!res.ok) throw new Error(await res.text());
      const json = await res.json();
      setCategories(json.categories || []);
      if (json.categories?.length) {
        if (!active || !json.categories.find((c: any) => c.name === active))
          setActive(json.categories[0].name);
      } else {
        setActive(null);
      }
    } catch (err: any) {
      setError(err?.message || "Failed to load categories");
    } finally {
      setLoadingCats(false);
    }
  }

  async function loadCategoryItems(category: string) {
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
    } catch (err: any) {
      setError(err?.message || "Failed to load items");
    } finally {
      setLoadingItems(false);
    }
  }

  async function checkIdExists(id: string) {
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

  async function presignUpload(
    filename: string,
    contentType: string,
    meta: Record<string, string>
  ) {
    const res = await fetch(`${API_BASE}/uploads/presign`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ filename, contentType, ...meta }),
    });
    if (!res.ok) throw new Error(await res.text());
    return await res.json();
  }

  // Initialize selections
  function initializeSelectionForCategory(
    category: string,
    schemaObj: any,
    itemsArr: any[]
  ) {
    setColumnSelectionByCategory((prev) => {
      if (prev?.[category]) return prev;
      const next = { ...(prev || {}) };
      const orig = schemaObj?.headerOriginalOrder?.length
        ? schemaObj.headerOriginalOrder.slice()
        : itemsArr?.length
        ? Object.keys(itemsArr[0]).filter((k) => k !== "category" && k !== "ID")
        : [];
      const normalized = orig.map((h: string) => String(h).toLowerCase());
      const map: Record<string, boolean> = { id: true, total: true };
      normalized.forEach((n: string) => (map[n] = true));
      next[category] = map;
      return next;
    });

    setRowSelectionByCategory((prev) => {
      const existing = (prev && prev[category]) || {};
      const next = { ...(prev || {}) };
      const currentIds = new Set((itemsArr || []).map((it: any) => it.ID));
      const merged: Record<string, boolean> = {};
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

  // Column toggles
  function toggleColumnForActive(normKey: string) {
    if (!active) return;
    setColumnSelectionByCategory((prev) => {
      const prevMap = (prev && prev[active]) || {};
      const next = { ...(prev || {}) };
      next[active] = { ...prevMap, [normKey]: !prevMap[normKey] };
      return next;
    });
  }

  // Row selection helpers
  function toggleRowSelection(id: string) {
    if (!active) return;
    setRowSelectionByCategory((prev) => {
      const prevMap = (prev && prev[active]) || {};
      const next = { ...(prev || {}) };
      next[active] = { ...prevMap, [id]: !prevMap[id] };
      return next;
    });
  }
  function setAllRowsSelectedForActive(selectAll: boolean) {
    if (!active) return;
    setRowSelectionByCategory((prev) => {
      const next = { ...(prev || {}) };
      const map: Record<string, boolean> = {};
      for (const it of items || []) map[it.ID] = !!selectAll;
      next[active] = map;
      return next;
    });
  }
  function areAllRowsSelected(): boolean {
    if (!active) return false;
    const map =
      (rowSelectionByCategory && rowSelectionByCategory[active]) || {};
    if (!items?.length) return false;
    return items.every((it) => !!map[it.ID]);
  }
  function anyRowSelected(): boolean {
    if (!active) return false;
    const map =
      (rowSelectionByCategory && rowSelectionByCategory[active]) || {};
    return (items || []).some((it) => !!map[it.ID]);
  }

  // Bulk delete rows
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
    if (firstId && ids.includes(firstId)) {
      setModalError({
        title: "Cannot delete first row",
        message: "The first data row is protected and cannot be deleted.",
      });
      return;
    }
    setPendingDeleteRowIds(ids);
    setConfirmDeleteRowsOpen(true);
  }
  async function deleteCategoryConfirmed() {
    if (!deletingCategory) return;
    try {
      // Optional: confirmation guard. Remove if you prefer no prompt
      const ok = window.confirm(
        `Delete category "${deletingCategory}"? This will remove all items in that category.`
      );
      if (!ok) {
        setDeletingCategory(false);
        return;
      }

      const res = await fetch(
        `${API_BASE}/category/${encodeURIComponent(deletingCategory)}`,
        { method: "DELETE" }
      );
      if (!res.ok) {
        const txt = await res.text().catch(() => "");
        throw new Error(txt || `Server responded ${res.status}`);
      }

      // If the deleted category was active, clear active selection
      setActive((prev) => (prev === deletingCategory ? null : prev));

      // Refresh categories to reflect deletion
      await loadCategories();

      // Close modal/state
      setDeletingCategory(false);
    } catch (err: any) {
      setModalError({
        title: "Delete category failed",
        message: err?.message || "Unable to delete category.",
      });
      setDeletingCategory(false);
    }
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
      const failed: Array<{ id: string; text: string; status: number }> = [];
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
      await loadCategoryItems(active!);
      setPendingDeleteRowIds([]);
      setConfirmDeleteRowsOpen(false);
    } catch (err: any) {
      setModalError({
        title: "Delete failed",
        message: err?.message || "Unable to delete rows.",
      });
    } finally {
      setDeletingRows(false);
    }
  }

  // Column delete
  function isProtectedColumn(orig: string) {
    const n = String(orig || "")
      .trim()
      .toLowerCase();
    return false;
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
          active!
        )}/column/${encodeURIComponent(col)}`,
        { method: "DELETE" }
      );
      if (!res.ok) {
        res = await fetch(
          `${API_BASE}/category/${encodeURIComponent(active!)}/column`,
          {
            method: "DELETE",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ column: col }),
          }
        );
      }
      if (!res.ok) throw new Error(await res.text());
      await loadCategoryItems(active!);
      setPendingDeleteColumn(null);
    } catch (err: any) {
      setModalError({
        title: "Delete column failed",
        message: err?.message || "Unable to delete column.",
      });
    } finally {
      setDeletingColumn(false);
    }
  }

  // Insert column
  function openInsertColumnAt(index: number) {
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
        `${API_BASE}/category/${encodeURIComponent(active!)}/column`,
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
      const nName = norm(name) as string;
      setSchema((prev: any) => {
        const orig = prev?.headerOriginalOrder
          ? prev.headerOriginalOrder.slice()
          : headerOriginalOrder().slice();
        const normArr = prev?.headerNormalizedOrder
          ? prev.headerNormalizedOrder.slice()
          : headerNormalizedOrder().slice();
        const idx = insertColumnIndex;
        orig.splice(idx, 0, name);
        normArr.splice(idx, 0, nName);
        return {
          ...(prev || {}),
          headerOriginalOrder: orig,
          headerNormalizedOrder: normArr,
          updatedAt: new Date().toISOString(),
        };
      });
      setItems((prev) => prev.map((it) => ({ ...it, [nName]: null })));
      setColumnSelectionByCategory((prev) => {
        const map = (prev && prev[active as string]) || {};
        return { ...prev, [active as string]: { ...map, [nName]: true } };
      });
      setInsertColumnOpen(false);
      setTimeout(() => ensureColWidthsForActive(), 0);
    } catch (err: any) {
      setModalError({
        title: "Add column failed",
        message: err?.message || "Unable to add column.",
      });
    } finally {
      setInsertingColumn(false);
    }
  }

  // Context menus
  function onHeaderCellContextMenu(
    e: React.MouseEvent,
    idx: number,
    label: string
  ) {
    e.preventDefault();
    if (
      String(label || "")
        .trim()
        .toLowerCase() === "id"
    )
      return;
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

  function onBodyCellContextMenu(
    e: React.MouseEvent,
    rowIdx: number,
    rowId: string,
    cellNormKey: string
  ) {
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

  function openInsertRowAt(index: number) {
    setInsertRowIndex(index);
    const headerNorm = headerNormalizedOrder();
    const initValues: Record<string, string> = {};
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
      const values: Record<string, string | null> = {};
      for (const nk of headerNorm) {
        if (nk === "total" || nk === "id" || nk === "category") continue;
        const v = newRowValues[nk];
        values[nk] = v === "" ? null : v;
      }

      const res = await fetch(
        `${API_BASE}/category/${encodeURIComponent(active!)}/row`,
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
        await loadCategoryItems(active!);
        setInsertRowOpen(false);
        return;
      }

      setItems((prev) => {
        const next = prev.slice();
        next.splice(insertRowIndex, 0, created);
        return next;
      });
      setRowSelectionByCategory((prev) => {
        const map = (prev && prev[active as string]) || {};
        return { ...prev, [active as string]: { ...map, [created.ID]: true } };
      });
      setInsertRowOpen(false);
    } catch (err: any) {
      setModalError({
        title: "Add row failed",
        message: err?.message || "Unable to add row.",
      });
    } finally {
      setInsertingRow(false);
    }
  }

  async function deleteRowById(id: string) {
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
        const map = (prev && prev[active as string]) || {};
        const nextMap = { ...map };
        delete nextMap[id];
        return { ...prev, [active as string]: nextMap };
      });
    } catch (err: any) {
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
    const rowId = ctxMenu.rowId!;
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
            setInsertImageTarget({ id: rowId, normKey: cellNormKey! });
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

  // Export helpers
  function computeTotalForItem(item: any, headerNorms: string[]) {
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
  function computeColumnTotalsForItems(itemsArr: any[], headerNorms: string[]) {
    const totals: Record<string, number | null> = {};
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
  function getSelectionForCategory(
    category: string,
    headerNorms: string[]
  ): Record<string, boolean> {
    const saved = columnSelectionByCategory[category];
    if (saved) return saved;
    const map: Record<string, boolean> = {};
    map["id"] = true;
    headerNorms.forEach((n) => (map[n] = true));
    map["total"] = true;
    return map;
  }

  // Client-side ExcelJS helpers
  function excelColWidthFromPx(px: number) {
    return Math.max(6, Math.round(px / 10));
  }
  function detectImageExtensionFromUrlOrCT(
    url: string,
    contentType: string | null
  ) {
    const u = (url || "").toLowerCase();
    if (contentType?.includes("png") || u.endsWith(".png")) return "png";
    if (contentType?.includes("webp") || u.endsWith(".webp")) return "webp";
    return "jpeg";
  }

  async function presignGetIfNeeded(
    imgObj: ImageCell
  ): Promise<{ url: string; contentType: string | null }> {
    if (imgObj?.s3Key) {
      try {
        const res = await fetch(
          `${API_BASE}/uploads/presign-get?s3Key=${encodeURIComponent(
            imgObj.s3Key
          )}`
        );
        if (!res.ok) return { url: imgObj.src || "", contentType: null };
        const json = await res.json();
        return { url: json.url || imgObj.src || "", contentType: null };
      } catch {
        return { url: imgObj.src || "", contentType: null };
      }
    }
    return { url: imgObj?.src || "", contentType: null };
  }

  async function fetchImageArrayBufferWithCT(
    url: string
  ): Promise<{ ab: ArrayBuffer | null; contentType: string | null }> {
    try {
      const res = await fetch(url, { mode: "cors" });
      if (!res.ok) throw new Error(`Image fetch failed: ${res.status}`);
      const ct = res.headers.get("content-type") || null;
      const ab = await res.arrayBuffer();
      return { ab, contentType: ct };
    } catch (e) {
      console.warn("fetchImageArrayBufferWithCT error for", url, e);
      return { ab: null, contentType: null };
    }
  }

  function colNumberToLetter(n: number) {
    // 1 -> A, 2 -> B ...
    let s = "";
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  async function addSheetWithImages({
    workbook,
    sheetName,
    itemsArr,
    headerOrig,
    headerNorm,
    selection,
    rowSelectionMap,
    exportSelectedOnly,
    includeTotalsRow,
    colWidthsPx,
  }: {
    workbook: ExcelJS.Workbook;
    sheetName: string;
    itemsArr: any[];
    headerOrig: string[];
    headerNorm: string[];
    selection: Record<string, boolean>;
    rowSelectionMap: Record<string, boolean> | null;
    exportSelectedOnly: boolean;
    includeTotalsRow: boolean;
    colWidthsPx: number[];
  }) {
    const worksheet = workbook.addWorksheet(sheetName.substring(0, 31));

    // Build headers for Excel: DO NOT include "Sel"
    const headers: string[] = [];
    const excelColIndexes: number[] = []; // maps data index -> excel column index
    let excelColCounter = 0;

    if (selection["id"]) {
      headers.push("ID");
      excelColCounter++;
    }

    for (let i = 0; i < headerOrig.length; i++) {
      if (selection[headerNorm[i]]) {
        headers.push(headerOrig[i]);
        excelColIndexes[i] = ++excelColCounter;
      } else {
        excelColIndexes[i] = -1; // skipped
      }
    }

    let totalColIdx = -1;
    if (selection["total"]) {
      headers.push("Total");
      totalColIdx = ++excelColCounter;
    }

    worksheet.addRow(headers);

    // Column widths from UI widths (skip "Sel" since not in Excel)
    const columns: ExcelJS.Column[] = [];
    if (selection["id"]) {
      columns.push({
        width: excelColWidthFromPx(colWidthsPx?.[1] || SIZING.defaults.id),
      });
    }
    for (let i = 0; i < headerOrig.length; i++) {
      if (selection[headerNorm[i]]) {
        const vIdx = 2 + i; // visual UI index
        columns.push({
          width: excelColWidthFromPx(
            colWidthsPx?.[vIdx] || SIZING.defaults.data
          ),
        });
      }
    }
    if (selection["total"]) {
      const vIdx = 2 + headerOrig.length;
      columns.push({
        width: excelColWidthFromPx(
          colWidthsPx?.[vIdx] || SIZING.defaults.total
        ),
      });
    }
    worksheet.columns = columns;

    // Filtered items according to row selection
    let filteredItems = itemsArr;
    if (exportSelectedOnly && rowSelectionMap) {
      filteredItems = itemsArr.filter((it) => !!rowSelectionMap[it.ID]);
    }

    // Helper: cells to sum for a given row
    function getRowSumCellRefs(rowNumber: number): string[] {
      const refs: string[] = [];
      for (let i = 0; i < headerOrig.length; i++) {
        if (!selection[headerNorm[i]]) continue;
        const excelCol = excelColIndexes[i];
        if (excelCol && excelCol > 0) {
          refs.push(`${colNumberToLetter(excelCol)}${rowNumber}`);
        }
      }
      // Exclude ID and Total columns from the row SUM (we add only data columns above)
      return refs;
    }

    // Add rows and collect image cells for embedding
    const dataRowStart = 2; // first data row after header (header is row 1)
    let currentRow = dataRowStart;

    for (const it of filteredItems) {
      const rowVals: Array<number | string | null> = [];

      // ID first (as text)
      if (selection["id"]) {
        rowVals.push(String(it.ID));
      }

      const imageCells: Array<{ colIndexInExcel: number; imgObj: ImageCell }> =
        [];

      for (let i = 0; i < headerOrig.length; i++) {
        const nk = headerNorm[i];
        if (!selection[nk]) continue;

        const rawVal = it[nk];
        const imgObj = parseImageValue(rawVal);
        if (imgObj?.src) {
          rowVals.push(""); // placeholder
          // colIndexInExcel = rowVals.length (1-based)
          imageCells.push({ colIndexInExcel: rowVals.length, imgObj });
        } else {
          const coerced = toNumberIfNumeric(rawVal);
          rowVals.push(coerced === undefined ? "" : coerced);
        }
      }

      // push placeholder for Total; we'll set a formula after adding row
      if (selection["total"]) {
        rowVals.push(""); // placeholder for formula
      }

      const newRow = worksheet.addRow(rowVals);
      newRow.height = Math.max(20, Math.round(getRowHeightById(it.ID) / 1.6));

      // Compute and set SUM formula for the Total cell of this row
      if (selection["total"] && totalColIdx > 0) {
        const sumRefs = getRowSumCellRefs(newRow.number);
        if (sumRefs.length > 0) {
          const firstRef = sumRefs[0];
          const lastRef = sumRefs[sumRefs.length - 1];
          // If columns are contiguous, we can use A2:K2. If not, join with SUM(A2,B2,...) is fine.
          // To be robust with skipped columns, use SUM with a comma-separated list.
          const formula = `SUM(${sumRefs.join(",")})`;
          worksheet.getCell(newRow.number, totalColIdx).value = { formula };
        } else {
          worksheet.getCell(newRow.number, totalColIdx).value = "";
        }
      }

      // Track tallest image for row to bump row height afterwards
      let maxImageHeightForRowPx = 0;

      // Embed images (relative to current row) with min width, aspect ratio, and auto column widening
      for (const cell of imageCells) {
        const { imgObj } = cell;
        const { url } = await presignGetIfNeeded(imgObj);
        const { ab, contentType } = await fetchImageArrayBufferWithCT(url);
        if (!ab) continue;

        const dims = await getImageDimensionsFromBuffer(ab);
        const intrinsicW = dims?.width ?? 1;
        const intrinsicH = dims?.height ?? 1;
        const aspectRatio = intrinsicH / intrinsicW;

        const ext = detectImageExtensionFromUrlOrCT(imgObj.src, contentType);
        const imageId = workbook.addImage({
          buffer: ab,
          extension: ext as any,
        });

        const colIdx1Based = cell.colIndexInExcel;
        const rowIdx1Based = newRow.number;

        const excelCol = worksheet.columns[colIdx1Based - 1];
        const currentChars = excelCol?.width || 10;
        const currentPxWidth = charsToPx(currentChars);

        const targetPxWidth = Math.max(MIN_IMAGE_WIDTH_PX, currentPxWidth);
        const finalHeightPx = Math.max(
          24,
          Math.round(targetPxWidth * aspectRatio)
        );

        if (currentPxWidth < targetPxWidth) {
          worksheet.columns[colIdx1Based - 1].width = pxToChars(targetPxWidth);
        }

        worksheet.addImage(imageId, {
          tl: { col: colIdx1Based - 1, row: rowIdx1Based - 1 },
          ext: { width: Math.max(24, targetPxWidth), height: finalHeightPx },
          editAs: "oneCell",
        });

        maxImageHeightForRowPx = Math.max(
          maxImageHeightForRowPx,
          finalHeightPx
        );
      }

      if (maxImageHeightForRowPx > 0) {
        const pointsPerPx = 0.75;
        const neededPoints = Math.round(maxImageHeightForRowPx * pointsPerPx);
        newRow.height = Math.max(newRow.height || 20, neededPoints);
      }

      currentRow++;
    }

    // Totals row (footer): write a SUM down the Total column, and for each data column write SUM down that column
    if (includeTotalsRow) {
      const totalsRow = worksheet.addRow([]);
      totalsRow.font = { bold: true };
      totalsRow.height = Math.max(18, Math.round(SIZING.defaults.row / 1.8));

      const firstDataRow = dataRowStart;
      const lastDataRow = currentRow - 1;

      // ID totals blank
      let colCursor = 1;
      if (selection["id"]) {
        worksheet.getCell(totalsRow.number, colCursor).value = "";
        colCursor++;
      }

      // For each selected data column, SUM down the column
      for (let i = 0; i < headerOrig.length; i++) {
        if (!selection[headerNorm[i]]) continue;
        const excelCol = excelColIndexes[i];
        const colLetter = colNumberToLetter(excelCol);
        const formula = `SUM(${colLetter}${firstDataRow}:${colLetter}${lastDataRow})`;
        worksheet.getCell(totalsRow.number, excelCol).value = { formula };
      }

      // Total column: SUM down the Total column
      if (selection["total"] && totalColIdx > 0) {
        const totalColLetter = colNumberToLetter(totalColIdx);
        const formula = `SUM(${totalColLetter}${firstDataRow}:${totalColLetter}${lastDataRow})`;
        worksheet.getCell(totalsRow.number, totalColIdx).value = { formula };
      }
    }

    // Style header row
    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.alignment = { vertical: "middle" };
    headerRow.height = Math.max(18, Math.round(SIZING.defaults.row / 1.8));
  }

  // Client-side ExcelJS export (Category)
  async function exportCategoryWithSelection(category: string) {
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
      const headerNorm = headerOrig.map((h: string) => String(h).toLowerCase());
      const selection = getSelectionForCategory(category, headerNorm);
      const rowSelectionMap =
        (rowSelectionByCategory && rowSelectionByCategory[category]) || null;
      const exportSelectedOnly = !!exportSelectedRowsOnlyByCategory[category];
      const includeTotalsRow = !!includeTotalsRowByCategory[category];

      const wb = new ExcelJS.Workbook();
      wb.creator = "Inventory Viewer";
      wb.created = new Date();

      await addSheetWithImages({
        workbook: wb,
        sheetName: category || "category",
        itemsArr: useItems,
        headerOrig,
        headerNorm,
        selection,
        rowSelectionMap,
        exportSelectedOnly,
        includeTotalsRow,
        colWidthsPx: getColWidths(),
      });

      const buffer = await wb.xlsx.writeBuffer();
      saveAs(
        new Blob([buffer], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }),
        `${category}.xlsx`
      );
    } catch (err: any) {
      console.error("exportCategoryWithSelection(exceljs)", err);
      setModalError({
        title: "Export failed",
        message: err?.message || "Unable to export category.",
      });
    } finally {
      setExportingCategory(false);
    }
  }

  // Client-side ExcelJS export (All)
  async function exportAllWithSelections() {
    setExportingAll(true);
    try {
      const wb = new ExcelJS.Workbook();
      wb.creator = "Inventory Viewer";
      wb.created = new Date();

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
        const headerNorm = headerOrig.map((h: string) =>
          String(h).toLowerCase()
        );
        const selection = getSelectionForCategory(catName, headerNorm);

        const rowSelectionMap =
          (rowSelectionByCategory && rowSelectionByCategory[catName]) || null;
        const exportSelectedOnly = !!exportSelectedRowsOnlyByCategory[catName];
        const includeTotalsRow = !!includeTotalsRowByCategory[catName];

        await addSheetWithImages({
          workbook: wb,
          sheetName: catName || "category",
          itemsArr: useItems,
          headerOrig,
          headerNorm,
          selection,
          rowSelectionMap,
          exportSelectedOnly,
          includeTotalsRow,
          colWidthsPx: getColWidths(), // you can persist per-category if desired
        });
      }

      const buffer = await wb.xlsx.writeBuffer();
      saveAs(
        new Blob([buffer], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }),
        `inventory.xlsx`
      );
    } catch (err: any) {
      console.error("exportAllWithSelections(exceljs)", err);
      setModalError({
        title: "Export failed",
        message: err?.message || "Unable to export all categories.",
      });
    } finally {
      setExportingAll(false);
    }
  }

  // Editing save helper
  async function saveEdit(id: string, normKey: string, newValue: string) {
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
      if (!res.ok) throw new Error(await res.text());
      const json = await res.json();
      setItems((prev) =>
        prev.map((it) =>
          it.ID === id
            ? { ...it, [normKey]: incoming, ...(json.updated || {}) }
            : it
        )
      );
      cancelEdit();
    } catch (err: any) {
      setModalError({
        title: "Save failed",
        message: err?.message || "Unable to save change.",
      });
    } finally {
      setSavingCellKey(null);
    }
  }
  function startEdit(id: string, normKey: string, initial: any) {
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
  function onInputKeyDown(e: React.KeyboardEvent, id: string, normKey: string) {
    if (e.key === "Enter") {
      e.preventDefault();
      saveEdit(id, normKey, cellValue);
    } else if (e.key === "Escape") cancelEdit();
  }

  // Header controls
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

  // Header checkbox row (UI only; Sel is for UI, not in Excel)
  function headerCheckboxRow(selection: Record<string, boolean>) {
    const orig = headerOriginalOrder();
    const normArr = headerNormalizedOrder();
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
        {orig.map((h, i) => {
          const protectedCol = isProtectedColumn(h);
          const nk = normArr[i];
          return (
            <th
              key={`chk-${i}`}
              className="p-1 border-b text-center"
              onContextMenu={(e) => onHeaderCellContextMenu(e, i, h)}
            >
              <div className="flex items-center justify-center gap-2">
                <input
                  type="checkbox"
                  checked={Boolean(selection[nk])}
                  onChange={() => toggleColumnForActive(nk)}
                  className="cursor-pointer"
                />
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    if (!protectedCol)
                      setPendingDeleteColumn({ columnOrig: h, columnNorm: nk });
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
                    <span className="inline-block">
                      <svg
                        className="animate-spin h-4 w-4 text-red-600"
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="currentColor"
                      >
                        <circle
                          className="opacity-25"
                          cx="12"
                          cy="12"
                          r="10"
                          strokeWidth="3"
                        />
                        <path
                          className="opacity-75"
                          d="M4 12a8 8 0 018-8"
                          strokeWidth="3"
                          strokeLinecap="round"
                        />
                      </svg>
                    </span>
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

  // Render table (UI includes Sel column for selection). Make it fill width.
  function renderTable() {
    if (!items?.length)
      return <div className="text-sm text-gray-500">No items found</div>;
    const orig = headerOriginalOrder();
    const normArr = headerNormalizedOrder();
    const selection = getSelectionForCategory(active as string, normArr);
    const totals = computeColumnTotalsForItems(items, normArr);
    const colWidths = getColWidths();
    const visualColCount = getVisualColumnCount();

    return (
      <div className="overflow-auto border border-gray-200 rounded shadow-sm relative w-full">
        <table className="min-w-full table-fixed text-sm w-full">
          <colgroup>
            {Array.from({ length: visualColCount }).map((_, i) => (
              <col
                key={`col-${i}`}
                style={{ width: `var(--col-${i}-w, ${colWidths[i] || 100}px)` }}
              />
            ))}
          </colgroup>

          <thead className="bg-gray-50 sticky top-0 z-10">
            {headerCheckboxRow(selection)}
            <tr>
              {/* Sel */}
              <th className="p-2 border-b border-r text-left relative">Sel</th>

              {/* ID */}
              <th
                className="p-2 border-b border-r text-left relative"
                onContextMenu={(e) => e.preventDefault()}
              >
                <div className="flex items-center justify-between">
                  <span>ID</span>
                </div>
                {/* Resizer */}
                <div
                  onMouseDown={(e) => {
                    e.preventDefault();
                    const idx = 1;
                    const startW = colWidths[idx] || SIZING.defaults.id;
                    setDraggingCol({
                      index: idx,
                      startX: e.clientX,
                      startWidth: startW,
                    });
                  }}
                  style={{
                    position: "absolute",
                    top: 0,
                    right: -3,
                    width: 6,
                    height: "100%",
                    cursor: "col-resize",
                  }}
                />
              </th>

              {/* Data headers */}
              {orig.map((h, idx) => {
                const vIdx = 2 + idx;
                const normKey = normArr[idx] || String(h).toLowerCase();
                const isImgCol = imageColumns.has(normKey);
                const imgColBaseH = getImageColHeight(normKey);
                const colWidth = colWidths[vIdx] || SIZING.defaults.data;

                return (
                  <th
                    key={h}
                    className="p-2 border-b border-r text-left relative"
                    onContextMenu={(e) => onHeaderCellContextMenu(e, idx, h)}
                    title={h}
                  >
                    <div className="flex items-center justify-between gap-2">
                      <span>{h}</span>
                      {isImgCol ? (
                        <div className="flex items-center gap-1">
                          <button
                            className="px-1 py-0.5 rounded bg-gray-200 hover:bg-gray-300"
                            title="Decrease image/cell size"
                            onClick={(e) => {
                              e.stopPropagation();
                              const nextW = colWidth - STEP.width;
                              setCssVar(
                                `--col-${vIdx}-w`,
                                `${Math.max(SIZING.colMin, nextW)}px`
                              );

                              const baseHStr = getComputedStyle(
                                document.documentElement
                              )
                                .getPropertyValue(`--imgcol-${normKey}-h`)
                                .trim();
                              const baseH = parseInt(
                                baseHStr || `${imgColBaseH}`,
                                10
                              );
                              const currentScaleStr = getComputedStyle(
                                document.documentElement
                              )
                                .getPropertyValue(`--imgcol-${normKey}-scale`)
                                .trim();
                              const currentScale =
                                parseFloat(currentScaleStr || "1") || 1;
                              const currentH = baseH * currentScale;
                              const targetH = Math.max(
                                SIZING.rowMin,
                                currentH - STEP.height
                              );
                              const nextScale = targetH / baseH;

                              setImageColScaleCssVar(normKey, nextScale);
                              commitImgColHeightDebounced(normKey, targetH);
                              setColWidthAt(vIdx, nextW);
                            }}
                          >
                            <Minus className="h-3 w-3" />
                          </button>
                          <button
                            className="px-1 py-0.5 rounded bg-gray-200 hover:bg-gray-300"
                            title="Increase image/cell size"
                            onClick={(e) => {
                              e.stopPropagation();
                              const nextW = colWidth + STEP.width;
                              setCssVar(
                                `--col-${vIdx}-w`,
                                `${Math.min(SIZING.colMax, nextW)}px`
                              );

                              const baseHStr = getComputedStyle(
                                document.documentElement
                              )
                                .getPropertyValue(`--imgcol-${normKey}-h`)
                                .trim();
                              const baseH = parseInt(
                                baseHStr || `${imgColBaseH}`,
                                10
                              );
                              const currentScaleStr = getComputedStyle(
                                document.documentElement
                              )
                                .getPropertyValue(`--imgcol-${normKey}-scale`)
                                .trim();
                              const currentScale =
                                parseFloat(currentScaleStr || "1") || 1;
                              const currentH = baseH * currentScale;
                              const targetH = Math.min(
                                SIZING.rowMax,
                                currentH + STEP.height
                              );
                              const nextScale = targetH / baseH;

                              setImageColScaleCssVar(normKey, nextScale);
                              commitImgColHeightDebounced(normKey, targetH);
                              setColWidthAt(vIdx, nextW);
                            }}
                          >
                            <PlusIcon className="h-3 w-3" />
                          </button>
                        </div>
                      ) : null}
                    </div>

                    {/* Drag resizer */}
                    <div
                      onMouseDown={(e) => {
                        e.preventDefault();
                        setDraggingCol({
                          index: vIdx,
                          startX: e.clientX,
                          startWidth: colWidth,
                        });
                      }}
                      style={{
                        position: "absolute",
                        top: 0,
                        right: -3,
                        width: 6,
                        height: "100%",
                        cursor: "col-resize",
                      }}
                    />
                  </th>
                );
              })}

              {/* Total */}
              <th
                className="p-2 border-b text-left relative"
                onContextMenu={(e) =>
                  onHeaderCellContextMenu(e, orig.length, "Total")
                }
              >
                <div className="flex items-center justify-between">
                  <span>Total</span>
                </div>
                <div
                  onMouseDown={(e) => {
                    e.preventDefault();
                    const idx = 2 + orig.length;
                    const startW = colWidths[idx] || SIZING.defaults.total;
                    setDraggingCol({
                      index: idx,
                      startX: e.clientX,
                      startWidth: startW,
                    });
                  }}
                  style={{
                    position: "absolute",
                    top: 0,
                    right: -3,
                    width: 6,
                    height: "100%",
                    cursor: "col-resize",
                  }}
                />
              </th>
            </tr>
          </thead>

          <tbody>
            {items.map((it, rowIdx) => {
              const map =
                (rowSelectionByCategory &&
                  rowSelectionByCategory[active as string]) ||
                {};
              const selected = map[it.ID] === undefined ? true : !!map[it.ID];

              const rowHeight = getRowHeightById(it.ID);
              const contentMaxH = Math.max(0, rowHeight - CELL_PADDING * 2);

              return (
                <tr
                  key={it.ID}
                  className="group odd:bg-white even:bg-gray-50 hover:bg-blue-50 transition-colors"
                  style={{ height: rowHeight }}
                >
                  {/* Sel */}
                  <td className="p-2 border-r text-center relative">
                    <input
                      type="checkbox"
                      checked={selected}
                      onChange={() => toggleRowSelection(it.ID)}
                      className="cursor-pointer"
                    />
                    {/* Row resizer handle */}
                    <div
                      onMouseDown={(e) => {
                        e.preventDefault();
                        setDraggingRow({
                          id: it.ID,
                          startY: e.clientY,
                          startHeight: rowHeight,
                        });
                      }}
                      title="Drag to resize row height"
                      style={{
                        position: "absolute",
                        left: 0,
                        right: 0,
                        bottom: -3,
                        height: 6,
                        cursor: "row-resize",
                      }}
                    />
                  </td>

                  {/* ID */}
                  <td className="relative p-2 border-r align-top font-mono text-xs bg-gray-100 group-hover:bg-blue-50">
                    <div
                      className="truncate"
                      style={{ maxHeight: contentMaxH }}
                    >
                      {it.ID}
                    </div>
                  </td>

                  {/* Data cells */}
                  {orig.map((h, colIndex) => {
                    const nkey = normArr[colIndex] || String(h).toLowerCase();
                    const value = it[nkey];
                    const imgObj = parseImageValue(value);
                    const isEditing =
                      editingCell &&
                      editingCell.id === it.ID &&
                      editingCell.normKey === nkey;
                    const savingKey = `${it.ID}|${nkey}`;

                    const visualIdx = 2 + colIndex;
                    const cellWidthVar = `var(--col-${visualIdx}-w, ${
                      getColWidths()[visualIdx] || SIZING.defaults.data
                    }px)`;
                    const cellWidth =
                      (getColWidths()[visualIdx] || SIZING.defaults.data) -
                      CELL_PADDING * 2;

                    const isImgCol = imageColumns.has(nkey);
                    const imgColHeightVar = `var(--imgcol-${nkey}-h, ${SIZING.defaults.imageColHeight}px)`;
                    const imgColScaleVar = `var(--imgcol-${nkey}-scale, 1)`;

                    const textCellHeight = Math.max(
                      0,
                      getRowHeightById(it.ID) - CELL_PADDING * 2
                    );

                    return (
                      <td
                        key={nkey}
                        className="relative p-2 border-r align-top overflow-hidden"
                        style={{ width: cellWidthVar }}
                        onDoubleClick={() => startEdit(it.ID, nkey, value)}
                        onContextMenu={(e) =>
                          onBodyCellContextMenu(e, rowIdx, it.ID, nkey)
                        }
                        title={h}
                      >
                        {!isEditing && (
                          <div
                            className="text-xs flex items-center gap-2"
                            style={{
                              width: "100%",
                              maxHeight: isImgCol
                                ? imgColHeightVar
                                : textCellHeight,
                              overflow: "hidden",
                            }}
                          >
                            {imgObj ? (
                              <div
                                className="img-box"
                                style={{
                                  width: "100%",
                                  height: imgColHeightVar,
                                  transformOrigin: "top",
                                  transform: `scaleY(${imgColScaleVar})`,
                                  willChange: "transform",
                                  contain: "layout paint",
                                }}
                              >
                                <img
                                  src={imgObj.src}
                                  alt={imgObj.alt || h}
                                  decoding="async"
                                  loading="lazy"
                                  style={{
                                    maxWidth: "100%",
                                    maxHeight: "100%",
                                    objectFit: "contain",
                                    display: "block",
                                  }}
                                />
                              </div>
                            ) : savingCellKey === savingKey ? (
                              <span className="text-indigo-600">Saving…</span>
                            ) : value === undefined || value === null ? (
                              ""
                            ) : (
                              <div
                                className="truncate"
                                style={{
                                  maxWidth: cellWidth,
                                  maxHeight: textCellHeight,
                                }}
                              >
                                {String(value)}
                              </div>
                            )}
                          </div>
                        )}
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

                  {/* Total */}
                  <td className="p-2 align-top text-sm">
                    <div style={{ maxHeight: contentMaxH, overflow: "hidden" }}>
                      {computeTotalForItem(it, normArr)}
                    </div>
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
        {/* Row context menu uses Insert image modal button */}
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
      if (!res.ok) throw new Error(await res.text());
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

  // Header controls
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
                setImagePreviewUrl(f ? URL.createObjectURL(f) : "");
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
              style={{ maxHeight: 240 }}
            >
              <img
                src={imagePreviewUrl}
                alt={imageAlt || "preview"}
                decoding="async"
                loading="lazy"
                style={{
                  maxWidth: "100%",
                  maxHeight: "220px",
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

                // Always auto width/height for images
                const payload = {
                  type: "image",
                  src: finalUrl,
                  alt: imageAlt || "",
                  width: "auto",
                  height: "auto",
                  ...(s3Key ? { s3Key } : {}),
                };
                const valueStr = JSON.stringify(payload);

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
  return (
    <div className="w-full px-4">
      <header className="mb-4 flex items-start justify-between">
        <div>
          <h1 className="text-2xl font-semibold">Inventory Viewer</h1>
          <p className="text-sm text-gray-500 mt-1">
            Right-click headers to Add column (left/right). Right-click a data
            cell to Add row (top/bottom), Delete row, or Insert image. Drag
            header edges to resize columns, and drag the small bar under the
            left cell of a row to resize height. Image columns show +/-
            controls; width changes are immediate, height changes animate
            smoothly.
          </p>
        </div>
        {headerControls()}
      </header>

      <div className="mb-3 flex items-center gap-2">
        <div className="font-medium">Categories:</div>
        {loadingCats ? (
          <div className="text-sm text-gray-500">Loading...</div>
        ) : (
          <div className="flex gap-2 overflow-auto w-full">
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
                    setDeletingCategory(c.name);
                  }}
                  className="ml-2 inline-flex items-center justify-center w-8 h-8 rounded-full bg-red-100 hover:bg-red-200"
                  title={`Delete category ${c.name}`}
                  disabled={!!deletingCategory} // disable while confirm modal is open
                >
                  <Trash2 className="h-4 w-4 text-red-600" />
                </button>
              </div>
            ))}
          </div>
        )}
      </div>

      {error && <div className="mb-3 text-sm text-red-600">{error}</div>}

      <main>
        <div className="mb-3 relative">
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

      {deletingCategory && typeof deletingCategory === "string" && (
        <div
          className="fixed inset-0 z-40 flex items-center justify-center bg-black/40"
          onClick={() => setDeletingCategory(false)}
        >
          <div
            className="bg-white rounded-lg shadow-lg p-4 max-w-sm w-full"
            onClick={(e) => e.stopPropagation()}
          >
            <h3 className="font-semibold text-lg mb-2">Delete category</h3>
            <div className="text-sm text-gray-700 mb-4">
              Are you sure you want to delete the category "
              <strong>{deletingCategory}</strong>"? This will remove all of its
              items.
            </div>
            <div className="flex justify-end gap-2">
              <button
                onClick={() => setDeletingCategory(false)}
                className="px-3 py-1 rounded bg-gray-200 hover:bg-gray-300"
              >
                Cancel
              </button>
              <button
                onClick={deleteCategoryConfirmed}
                className="px-3 py-1 rounded bg-red-600 text-white hover:bg-red-700"
                disabled={
                  deletingRows ||
                  deletingColumn ||
                  exportingAll ||
                  exportingCategory ||
                  replacing
                }
              >
                Delete
              </button>
            </div>
          </div>
        </div>
      )}

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

// Tiny inline spinner component (used above)
function Spinner({ className = "h-4 w-4 text-white" }: { className?: string }) {
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
