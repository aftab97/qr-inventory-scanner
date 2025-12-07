import React, { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { isLocalHost } from "../api";

const API_BASE = isLocalHost ? "http://localhost:3000" : 'http://qr-scanner-api.us-east-1.elasticbeanstalk.com/'

function Spinner({ className = "h-4 w-4 text-white" }) {
  return (
    <svg className={`animate-spin ${className}`} viewBox="0 0 24 24" fill="none" stroke="currentColor" aria-hidden>
      <circle className="opacity-25" cx="12" cy="12" r="10" strokeWidth="3"></circle>
      <path className="opacity-75" d="M4 12a8 8 0 018-8" strokeWidth="3" strokeLinecap="round"></path>
    </svg>
  );
}

/**
 * UploadCsv.jsx
 *
 * Fix included:
 * - When a file is selected, we parse it immediately.
 * - After parsing completes (success or failure), the hidden <input> value is cleared so selecting
 *   the same file again will fire the change event and trigger a re-parse.
 *
 * This prevents the "re-upload same file doesn't refetch" issue.
 *
 * All other functionality is the same as prior: visible Choose file button, automatic parse,
 * auto/manual category mode, apply, popups, spinners, etc.
 */

export default function UploadCsv() {
  const [file, setFile] = useState(null);

  const [operation, setOperation] = useState("update"); // update | replace

  const [categoryMode, setCategoryMode] = useState("manual"); // "auto" | "manual"
  const [canAutoCategory, setCanAutoCategory] = useState(false);

  const [datasets, setDatasets] = useState([]); // { sheetName, headerOriginalOrder, rows, warnings }
  const [selected, setSelected] = useState({}); // sheetName -> bool
  const [categoryOverrides, setCategoryOverrides] = useState({}); // sheetName -> manual category name

  const [previewing, setPreviewing] = useState(false); // parsing/preview in progress
  const [loadingApply, setLoadingApply] = useState(false); // apply in progress
  const [responseJson, setResponseJson] = useState(null);
  const [error, setError] = useState(null);

  // Modal state for popup messages
  const [modalOpen, setModalOpen] = useState(false);
  const [modalTitle, setModalTitle] = useState("");
  const [modalMessage, setModalMessage] = useState("");
  const [modalSuccess, setModalSuccess] = useState(false);

  const fileInputRef = useRef(null);

  // When datasets change ensure selected map has entries
  useEffect(() => {
    if (datasets.length > 0) {
      setSelected((prev) => {
        const next = { ...(prev || {}) };
        datasets.forEach((d) => {
          if (!(d.sheetName in next)) next[d.sheetName] = true;
        });
        return next;
      });
    }
  }, [datasets]);

  // When switching to manual via UI, pre-fill overrides with sheet names for easy editing
  useEffect(() => {
    if (categoryMode === "manual" && datasets.length > 0) {
      setCategoryOverrides((prev) => {
        const next = { ...(prev || {}) };
        datasets.forEach((d) => {
          if (!(d.sheetName in next) || !next[d.sheetName]) {
            next[d.sheetName] = d.sheetName; // default to sheet name for easy edit
          }
        });
        return next;
      });
    } else if (categoryMode === "auto") {
      setCategoryOverrides({});
    }
  }, [categoryMode, datasets]);

  // Clear form
  const clearAll = () => {
    setFile(null);
    setDatasets([]);
    setSelected({});
    setCategoryOverrides({});
    setResponseJson(null);
    setError(null);
    setCanAutoCategory(false);
    setCategoryMode("manual");
  };

  // helper to show modal
  const showModal = (success, title, message) => {
    setModalSuccess(!!success);
    setModalTitle(title);
    setModalMessage(message);
    setModalOpen(true);
  };

  // Inspect file to determine auto-possible and then parse automatically
  const handleFileButtonClick = () => {
    fileInputRef.current?.click();
  };

  const handleFileChange = async (ev) => {
    setError(null);
    setResponseJson(null);
    setDatasets([]);
    setSelected({});
    setCategoryOverrides({});
    setCanAutoCategory(false);

    const f = ev.target.files?.[0] ?? null;
    if (!f) {
      setFile(null);
      return;
    }
    setFile(f);

    const name = (f.name || "").toLowerCase();
    const isXlsx = name.endsWith(".xlsx") || name.endsWith(".xls");
    if (isXlsx) {
      try {
        const ab = await f.arrayBuffer();
        const wb = XLSX.read(ab, { type: "array" });
        const sheetCount = Array.isArray(wb.SheetNames) ? wb.SheetNames.length : 0;
        const hasTabs = sheetCount > 1;
        setCanAutoCategory(hasTabs);
        setCategoryMode(hasTabs ? "auto" : "manual");
      } catch (err) {
        console.warn("Failed to inspect workbook:", err);
        setCanAutoCategory(false);
        setCategoryMode("manual");
      }
    } else {
      // CSV: cannot auto
      setCanAutoCategory(false);
      setCategoryMode("manual");
    }

    // parse automatically. Ensure input is cleared after parse so re-selecting same file fires change.
    try {
      await parsePreviewWithAuto(f);
    } finally {
      // CLEAR the actual <input> value so selecting the same file again triggers onChange.
      // We keep `file` state so UI still shows selected file name.
      try {
        if (fileInputRef.current) fileInputRef.current.value = "";
      } catch (e) {
        // ignore
      }
    }
  };

  // parse/preview by calling backend /parse
  const parsePreviewWithAuto = async (explicitFile) => {
    setError(null);
    setResponseJson(null);
    setDatasets([]);
    setSelected({});
    setCategoryOverrides({});
    const f = explicitFile || file;

    if (!f) {
      const msg = "Please select a file.";
      setError(msg);
      showModal(false, "Parse failed", msg);
      return;
    }

    setPreviewing(true);
    try {
      const form = new FormData();
      form.append("file", f);
      const res = await fetch(`${API_BASE}/parse`, { method: "POST", body: form });
      if (!res.ok) {
        const txt = await res.text().catch(() => "Server error");
        throw new Error(txt || `Server responded ${res.status}`);
      }
      const json = await res.json();
      setResponseJson(json);

      if (json.datasets && Array.isArray(json.datasets)) {
        const ds = json.datasets.map((d) => {
          const sheetName = d.sheetName || d.category || "dataset";
          const headers = Array.isArray(d.headerOriginalOrder) ? d.headerOriginalOrder : [];
          return {
            sheetName,
            headerOriginalOrder: headers,
            rows: d.rows || 0,
            warnings: d.warnings || [],
          };
        });
        setDatasets(ds);

        // mark all as selected
        const sel = {};
        ds.forEach((d) => (sel[d.sheetName] = true));
        setSelected(sel);

        if (categoryMode === "manual") {
          const ov = {};
          ds.forEach((d) => (ov[d.sheetName] = d.sheetName));
          setCategoryOverrides(ov);
        } else {
          setCategoryOverrides({});
        }
      } else {
        setDatasets([]);
      }
    } catch (err) {
      console.error("parsePreview error:", err);
      const msg = (err && err.message) || "Preview failed";
      setError(msg);
      showModal(false, "Parsing failed", String(msg));
    } finally {
      setPreviewing(false);
    }
  };

  // toggles and setters
  const toggleDatasetIncluded = (sheetName) => setSelected((s) => ({ ...(s || {}), [sheetName]: !s[sheetName] }));
  const setCategoryOverride = (sheetName, value) => setCategoryOverrides((s) => ({ ...(s || {}), [sheetName]: value }));

  // validation & mapping
  const validateBeforeApply = () => {
    if (!datasets.length) return { ok: false, message: "No datasets available to apply." };
    const anySelected = Object.values(selected || {}).some(Boolean);
    if (!anySelected) return { ok: false, message: "No datasets selected to apply." };
    if (categoryMode === "manual") {
      for (const d of datasets) {
        if (!selected[d.sheetName]) continue;
        const cat = (categoryOverrides[d.sheetName] || "").trim();
        if (!cat) return { ok: false, message: `Category name is required for sheet "${d.sheetName}".` };
      }
    }
    return { ok: true };
  };

  const buildCategoriesMapping = () => {
    const mapping = {};
    for (const d of datasets) {
      if (!selected[d.sheetName]) continue;
      const cat = categoryMode === "auto" ? d.sheetName : (categoryOverrides[d.sheetName] || "").trim();
      mapping[d.sheetName] = cat;
    }
    return mapping;
  };

  // apply (update/replace)
  const apply = async ({ previewOnly = false }) => {
    setError(null);
    setResponseJson(null);

    const v = validateBeforeApply();
    if (!v.ok) {
      setError(v.message);
      showModal(false, "Validation failed", v.message);
      return;
    }

    if (!file) {
      const msg = "No file selected";
      setError(msg);
      showModal(false, "Apply failed", msg);
      return;
    }

    setLoadingApply(true);
    try {
      const categoriesMap = buildCategoriesMapping();
      const params = new URLSearchParams();
      if (previewOnly) params.set("dry", "true");
      if (operation === "replace") params.set("replace", "true");
      else params.set("update", "true");

      const url = `${API_BASE}/parse?${params.toString()}`;
      const form = new FormData();
      form.append("file", file);
      form.append("categories", JSON.stringify(categoriesMap));
      const res = await fetch(url, { method: "POST", body: form });
      if (!res.ok) {
        const txt = await res.text().catch(() => "Server error");
        throw new Error(txt || `Server responded ${res.status}`);
      }
      const json = await res.json();
      setResponseJson(json);

      // success — show popup and on OK clear form
      showModal(true, "Apply succeeded", "The operation completed successfully.");
    } catch (err) {
      console.error("apply error:", err);
      const msg = (err && err.message) || "Apply failed";
      setError(msg);
      showModal(false, "Apply failed", String(msg));
    } finally {
      setLoadingApply(false);
    }
  };

  const handleApplyClick = async () => {
    if (operation === "replace") {
      const ok = window.confirm("You are about to DELETE items for the selected categories and replace them. Continue?");
      if (!ok) return;
    }
    await apply({ previewOnly: false });
  };

  // When modal OK is clicked: if success then clear the form
  const onModalOk = () => {
    setModalOpen(false);
    if (modalSuccess) {
      clearAll();
    }
  };

  // "Edit categories" helper: switch to manual and prefill overrides
  const onEditCategories = () => {
    setCategoryMode("manual");
    const ov = {};
    datasets.forEach((d) => {
      ov[d.sheetName] = d.sheetName;
    });
    setCategoryOverrides(ov);
  };

  const applyDisabled =
    loadingApply ||
    previewing ||
    datasets.length === 0 ||
    !Object.values(selected || {}).some(Boolean) ||
    (categoryMode === "manual" && datasets.some((d) => selected[d.sheetName] && !(categoryOverrides[d.sheetName] || "").trim()));

  return (
    <div className="app-container max-w-3xl mx-auto p-4">
      <header className="mb-6">
        <h1 className="text-2xl font-semibold">Upload CSV / XLSX (category-aware)</h1>
        <p className="text-sm text-gray-500 mt-1">Click "Choose file" to upload. Files parse automatically on upload.</p>
      </header>

      <section className="bg-white rounded-lg shadow p-4 space-y-4">
        {/* File input as hidden with visible button */}
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-2">Choose CSV / XLSX file</label>

          <div className="flex items-center gap-3">
            <button
              type="button"
              onClick={handleFileButtonClick}
              className="inline-flex items-center gap-2 px-4 py-2 rounded bg-blue-600 text-white hover:bg-blue-700 focus:outline-none"
            >
              <svg className="h-4 w-4" viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M12 5v14M5 12h14" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" /></svg>
              Choose file
            </button>

            <div className="text-sm text-gray-700">
              {file ? (
                <>
                  <span className="font-medium">{file.name}</span>
                  <span className="ml-2 text-gray-500">({Math.round(file.size / 1024)} KB)</span>
                </>
              ) : (
                <span className="text-gray-500">No file selected</span>
              )}
            </div>
          </div>

          <input
            ref={fileInputRef}
            type="file"
            accept=".csv,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
            onChange={handleFileChange}
            className="hidden"
          />

          <div className="mt-2 text-xs text-gray-500">Parsing & preview runs automatically after file selection.</div>
        </div>

        {/* Operation */}
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-2">Operation</label>
          <div className="flex gap-3">
            <label className={`flex-1 p-3 rounded-lg text-center cursor-pointer ${operation === "update" ? "bg-green-600 text-white" : "bg-gray-100 text-gray-800"}`}>
              <input type="radio" name="operation" checked={operation === "update"} onChange={() => setOperation("update")} className="hidden" />
              Update (upsert rows)
            </label>
            <label className={`flex-1 p-3 rounded-lg text-center cursor-pointer ${operation === "replace" ? "bg-red-600 text-white" : "bg-gray-100 text-gray-800"}`}>
              <input type="radio" name="operation" checked={operation === "replace"} onChange={() => setOperation("replace")} className="hidden" />
              Replace (delete all then write)
            </label>
          </div>
        </div>

        {/* Category mode */}
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-2">Category mode</label>
          <div className="flex gap-3 items-center">
            {canAutoCategory ? (
              <label className={`p-3 rounded-lg cursor-pointer ${categoryMode === "auto" ? "bg-indigo-600 text-white" : "bg-gray-100 text-gray-800"}`}>
                <input type="radio" name="categoryMode" checked={categoryMode === "auto"} onChange={() => setCategoryMode("auto")} className="hidden" />
                Auto (use sheet/tab names)
              </label>
            ) : null}

            <label className={`p-3 rounded-lg cursor-pointer ${categoryMode === "manual" ? "bg-indigo-600 text-white" : "bg-gray-100 text-gray-800"}`}>
              <input type="radio" name="categoryMode" checked={categoryMode === "manual"} onChange={() => setCategoryMode("manual")} className="hidden" />
              Manual (edit category names)
            </label>

            {datasets.length > 0 && categoryMode === "auto" && (
              <button onClick={onEditCategories} className="ml-4 px-3 py-2 rounded bg-yellow-500 text-white hover:bg-yellow-600">Edit categories</button>
            )}
          </div>

          <div className="text-xs text-gray-500 mt-2">
            {canAutoCategory ? "Auto will use sheet names as categories when multiple sheets are present; use Edit categories to override." : "Auto not available for this file type — use Manual and provide category names."}
          </div>
        </div>

        {/* Parsing indicator */}
        {previewing && (
          <div className="p-3 bg-gray-50 rounded text-sm flex items-center gap-2">
            <Spinner className="h-4 w-4 text-gray-700" />
            Parsing file and preparing preview…
          </div>
        )}

        {/* Datasets */}
        {datasets.length > 0 && (
          <div className="mt-4">
            <h3 className="text-lg font-medium">Datasets found</h3>
            <p className="text-sm text-gray-500 mb-2">Uncheck to skip a dataset during apply. Manual mode requires category names for each selected dataset.</p>

            <div className="space-y-3">
              {datasets.map((d) => {
                const overrideValue = categoryOverrides[d.sheetName] ?? "";
                return (
                  <div key={d.sheetName} className="p-3 border rounded flex flex-col gap-3">
                    <div className="flex items-start gap-3">
                      <input type="checkbox" checked={!!selected[d.sheetName]} onChange={() => toggleDatasetIncluded(d.sheetName)} className="mt-1" />
                      <div className="flex-1">
                        <div className="flex gap-3 items-center">
                          <div className="font-medium">{d.sheetName}</div>
                          <div className="text-xs text-gray-500">({d.rows} rows)</div>
                          {d.warnings && d.warnings.length > 0 && (
                            <div className="ml-2 text-xs text-yellow-700 bg-yellow-100 px-2 py-0.5 rounded">Warnings: {d.warnings.length}</div>
                          )}
                        </div>

                        <div className="mt-2 flex gap-2 items-center">
                          <label className="text-sm text-gray-700">Category:</label>

                          {categoryMode === "auto" ? (
                            <input value={d.sheetName} readOnly className="p-1 border rounded text-sm bg-gray-50" />
                          ) : (
                            <input
                              value={overrideValue}
                              onChange={(e) => setCategoryOverride(d.sheetName, e.target.value)}
                              placeholder="category name (required)"
                              className="p-1 border rounded text-sm"
                            />
                          )}

                          <div className="ml-4 text-xs text-gray-600">Headers: {d.headerOriginalOrder?.slice(0, 6).join(", ")}{(d.headerOriginalOrder?.length || 0) > 6 ? "…" : ""}</div>
                        </div>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {error && <div className="text-sm text-red-600">{error}</div>}

        {responseJson && (
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">Server response / Preview</label>
            <div className="max-h-96 overflow-auto bg-gray-50 border rounded p-3 text-sm">
              <pre className="whitespace-pre-wrap text-xs">{JSON.stringify(responseJson, null, 2)}</pre>
            </div>
            <div className="mt-2 flex gap-2">
              <button onClick={() => navigator.clipboard.writeText(JSON.stringify(responseJson))} className="py-2 px-3 rounded bg-blue-600 text-white text-sm">Copy JSON</button>
              <a className="py-2 px-3 rounded bg-indigo-600 text-white text-sm" href={`data:application/json;charset=utf-8,${encodeURIComponent(JSON.stringify(responseJson))}`} download="parsed.json">Download JSON</a>
            </div>
          </div>
        )}

        {/* Apply area */}
        <div className="mt-6 border-t pt-4">
          <div className="flex items-center justify-between gap-3">
            <div>
              <div className="text-sm text-gray-700">After confirming datasets and categories, click Apply to run the selected operation on the server.</div>
              {categoryMode === "manual" && datasets.some((d) => selected[d.sheetName] && !(categoryOverrides[d.sheetName] || "").trim()) ? (
                <div className="mt-2 text-xs text-red-600">Please fill all category names for selected datasets before applying.</div>
              ) : null}
            </div>

            <div className="flex gap-2">
              <button onClick={clearAll} className="px-3 py-2 rounded bg-gray-100">Clear</button>
              <button
                onClick={handleApplyClick}
                disabled={applyDisabled}
                className={`px-4 py-2 rounded ${applyDisabled ? "bg-gray-300 text-gray-700 cursor-not-allowed" : "bg-blue-600 text-white hover:bg-blue-700"}`}
              >
                {loadingApply ? (
                  <span className="inline-flex items-center gap-2"><Spinner className="h-4 w-4 text-white" /> Applying…</span>
                ) : (
                  <span>{operation === "replace" ? "Apply Replace" : "Apply Update"}</span>
                )}
              </button>
            </div>
          </div>
        </div>
      </section>

      {/* Modal Popup */}
      {modalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40">
          <div className="bg-white rounded-lg shadow-lg p-6 max-w-lg w-full">
            <div className="flex items-start gap-3">
              <div className={`p-2 rounded ${modalSuccess ? "bg-green-100" : "bg-red-100"}`}>
                {modalSuccess ? (
                  <svg className="h-6 w-6 text-green-600" viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M5 13l4 4L19 7" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" /></svg>
                ) : (
                  <svg className="h-6 w-6 text-red-600" viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M6 18L18 6M6 6l12 12" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" /></svg>
                )}
              </div>
              <div className="flex-1">
                <h3 className="text-lg font-semibold">{modalTitle}</h3>
                <div className="text-sm text-gray-700 mt-2">{modalMessage}</div>
              </div>
            </div>

            <div className="mt-6 flex justify-end">
              <button onClick={() => setModalOpen(false)} className="px-3 py-1 rounded bg-gray-200 hover:bg-gray-300">Close</button>
              <button onClick={onModalOk} className="ml-2 px-3 py-1 rounded bg-blue-600 text-white hover:bg-blue-700">OK</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}