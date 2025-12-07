import React, { useEffect, useMemo, useRef, useState } from "react";
import { Link, useNavigate, useLocation } from "react-router-dom";
import { Camera, Check, X, ChevronRight, Minus, Plus } from "lucide-react";
import QRScanner, { type QRScannerHandle } from "../components/qr/qr-scanner";
import { isLocalHost } from "../api";

const API_BASE = isLocalHost ? "http://localhost:3000" : 'http://qr-scanner-api.us-east-1.elasticbeanstalk.com/'
const LS_SELECTED_COLUMN = "scan_selected_column";

type RawItem = Record<string, any>;

export default function Scan(): JSX.Element {
  const nav = useNavigate();
  const location = useLocation();
  const scannerRef = useRef<QRScannerHandle | null>(null);

  // Categories needed to search by UUID (no user selection; we resolve automatically)
  const [categories, setCategories] = useState<string[]>([]);
  const [catsLoading, setCatsLoading] = useState<boolean>(false);

  // Column selection (from SelectColumn page)
  const [selectedColumn, setSelectedColumn] = useState<string>(() => {
    try { return window.localStorage.getItem(LS_SELECTED_COLUMN) ?? ""; } catch { return ""; }
  });

  // Scanner state
  const [scanning, setScanning] = useState<boolean>(false);
  const [decoded, setDecoded] = useState<string | null>(null);
  const [searching, setSearching] = useState<boolean>(false);
  const [foundItem, setFoundItem] = useState<RawItem | null>(null);
  const [foundCategory, setFoundCategory] = useState<string | null>(null);

  // Operation state
  const [updating, setUpdating] = useState<boolean>(false);

  // Quantity to increment (+/-) — defaults to 1, cannot go below 1
  const [quantity, setQuantity] = useState<number>(1);

  // Snapshot of the amount applied/shown in result screen (so UI shows the correct number even if quantity resets)
  const [appliedQuantity, setAppliedQuantity] = useState<number>(1);

  // Full-screen result view state (mobile-friendly replacement of content)
  const [resultView, setResultView] = useState<{
    open: boolean;
    success: boolean;
    title: string;
    message: string;
    details?: string;
  }>({ open: false, success: false, title: "", message: "", details: "" });

  const showResultView = (success: boolean, title: string, message: string, details?: string) => {
    setResultView({ open: true, success, title, message, details });
    try { scannerRef.current?.stop(); } catch {}
    setScanning(false);
  };

  const closeResultAndReset = () => {
    setResultView({ open: false, success: false, title: "", message: "", details: "" });
    // Reset entire scanner page to initial state
    setDecoded(null);
    setFoundItem(null);
    setFoundCategory(null);
    setScanning(false);
    setUpdating(false);
    // Keep quantity reset to 1 for next operation
    setQuantity(1);
  };

  // Load categories once (we will search through these)
  useEffect(() => {
    let cancelled = false;
    async function loadCategories() {
      setCatsLoading(true);
      try {
        const res = await fetch(`${API_BASE}/categories`);
        if (!res.ok) throw new Error(await res.text());
        const json = await res.json();
        if (cancelled) return;
        const list = (json.categories || []).map((c: any) => String(c.name));
        setCategories(list);
      } catch (err) {
        console.error("Failed to load categories", err);
      } finally {
        if (!cancelled) setCatsLoading(false);
      }
    }
    loadCategories();
    return () => { cancelled = true; };
  }, []);

  // Helper: apply selected column from localStorage and reset UI
  const refreshSelectedColumnFromStorage = () => {
    try {
      const val = window.localStorage.getItem(LS_SELECTED_COLUMN) ?? "";
      setSelectedColumn(val);
      // Reset to camera start
      setDecoded(null);
      setFoundItem(null);
      setFoundCategory(null);
      setScanning(false);
      setResultView({ open: false, success: false, title: "", message: "" });
      setQuantity(1);
      setAppliedQuantity(1);
    } catch { /* ignore */ }
  };

  // React to returning from select page and storage changes
  useEffect(() => {
    const onFocus = () => refreshSelectedColumnFromStorage();
    const onStorage = (e: StorageEvent) => {
      if (e.key === LS_SELECTED_COLUMN) refreshSelectedColumnFromStorage();
    };
    window.addEventListener("focus", onFocus);
    window.addEventListener("storage", onStorage);
    return () => {
      window.removeEventListener("focus", onFocus);
      window.removeEventListener("storage", onStorage);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // If select page sends ?col=... back, apply it and clean URL
  useEffect(() => {
    const qs = new URLSearchParams(location.search);
    const col = qs.get("col");
    if (col) {
      try { window.localStorage.setItem(LS_SELECTED_COLUMN, col); } catch {}
      refreshSelectedColumnFromStorage();
      nav(location.pathname, { replace: true });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [location.search]);

  // Scan result: search all categories sequentially for the UUID
  const onScanResult = async (text: string) => {
    setDecoded(text);
    setFoundItem(null);
    setFoundCategory(null);
    setSearching(true);

    const target = String(text).trim();
    if (!target) {
      // snapshot the current chosen amount so result screen shows correct number
      setAppliedQuantity(Math.max(1, quantity));
      showResultView(false, "Invalid scan", "Scanned value was empty. Please try again.");
      setSearching(false);
      return;
    }

    try {
      let found: RawItem | null = null;
      let foundCat: string | null = null;

      for (const cat of categories) {
        try {
          const res = await fetch(`${API_BASE}/category/${encodeURIComponent(cat)}`);
          if (!res.ok) continue;
          const json = await res.json();
          const items: RawItem[] = json.items || [];

          for (const it of items) {
            let matched = false;
            for (const key of Object.keys(it)) {
              if (/^id$/i.test(key) && String(it[key]) === target) { matched = true; break; }
            }
            if (!matched) {
              if ((it.ID && String(it.ID) === target) || (it.id && String(it.id) === target)) matched = true;
            }
            if (!matched) {
              for (const key of Object.keys(it)) {
                const v = it[key];
                if (v !== undefined && v !== null && String(v) === target) { matched = true; break; }
              }
            }
            if (matched) { found = it; foundCat = cat; break; }
          }
        } catch { /* ignore category fetch errors and continue */ }
        if (found) break;
      }

      if (!found) {
        setAppliedQuantity(Math.max(1, quantity));
        showResultView(false, "Unable to add", "Item does not exist. Please try again.");
      } else {
        const normalized: RawItem = {};
        for (const k of Object.keys(found)) normalized[String(k).toLowerCase()] = found[k];
        setFoundItem(normalized);
        setFoundCategory(foundCat);

        // Immediately check selected column existence on this item (client-side early guard)
        const hasColumn = selectedColumn && Object.prototype.hasOwnProperty.call(normalized, selectedColumn);
        if (!selectedColumn || !hasColumn) {
          setAppliedQuantity(Math.max(1, quantity));
          showResultView(false, "Column not available", "The selected column does not exist for this item. Please select another column and try again.");
        }
      }
    } catch (err) {
      console.error("Search error", err);
      setAppliedQuantity(Math.max(1, quantity));
      showResultView(false, "Lookup failed", "We couldn't verify the item. Please try again.");
    } finally {
      setSearching(false);
      setScanning(false);
    }
  };

  // Confirm increment: treat null/empty as 0 and PATCH the item attribute (server enforces column exists in schema)
  const handleConfirmIncrement = async () => {
    const qtyToApply = Math.max(1, quantity); // snapshot the amount immediately
    setAppliedQuantity(qtyToApply);

    if (!decoded) { showResultView(false, "Invalid scan", "Nothing scanned. Please try again."); return; }
    if (!foundItem || !foundCategory) { showResultView(false, "Unable to add", "Item does not exist. Please try again."); return; }
    if (!selectedColumn || !Object.prototype.hasOwnProperty.call(foundItem, selectedColumn)) {
      showResultView(false, "Column not available", "The selected column is not available for this item. Please choose another column.");
      return;
    }

    setUpdating(true);
    try {
      const idKey = Object.keys(foundItem).find((k) => /^id$/i.test(k)) ?? "id";
      const idVal = foundItem[idKey] ?? decoded;

      const currentRaw = foundItem[selectedColumn];
      const current = Number(String(currentRaw ?? "").replace(",", ".")) || 0;
      const newValue = current + qtyToApply;

      const body = { attribute: selectedColumn, value: newValue, category: foundCategory };
      const res = await fetch(`${API_BASE}/item/${encodeURIComponent(String(idVal))}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });

      if (!res.ok) {
        const txt = await res.text().catch(() => "");
        let payload: any = null;
        try { payload = JSON.parse(txt); } catch {}
        const msg = payload?.error || txt || `Server responded ${res.status}`;

        if (payload?.code === "COLUMN_NOT_IN_SCHEMA") {
          showResultView(false, "Column not available", payload.details || "Selected column does not exist for this category.");
          setUpdating(false);
          return;
        }
        if (res.status === 404) {
          showResultView(false, "Unable to add", "Item does not exist.");
          setUpdating(false);
          return;
        }
        if (res.status === 409 && payload?.code === "CATEGORY_CONDITION_FAILED") {
          showResultView(false, "Unable to add", "Item category mismatch. Please rescan.");
          setUpdating(false);
          return;
        }

        throw new Error(msg);
      }

      await res.json();
      setFoundItem((prev) => (prev ? { ...prev, [selectedColumn]: newValue } : prev));
      showResultView(true, "Successfully added", `The item has been updated (+${qtyToApply}).`);
    } catch (err) {
      console.error("Update failed", err);
      showResultView(false, "Unable to add", "We couldn't update the item. Please try again.");
    } finally {
      setUpdating(false);
      // Keep quantity as-is until user dismisses; it will reset on closeResultAndReset
    }
  };

  const handleCancel = () => {
    setDecoded(null);
    setFoundItem(null);
    setFoundCategory(null);
    setScanning(false);
    setResultView({ open: false, success: false, title: "", message: "" });
    setQuantity(1);
    setAppliedQuantity(1);
  };

  const goSelectColumn = () => {
    nav("/select-column");
  };

  const displayName = useMemo(() => {
    if (!foundItem) return null;
    const nom = foundItem.nom ?? foundItem.nombre ?? "";
    const pierres = foundItem.pierres ?? foundItem.piedras ?? "";
    if (nom && pierres) return `${nom} - ${pierres}`;
    if (nom) return String(nom);
    if (pierres) return String(pierres);
    const idKey = Object.keys(foundItem).find((k) => /^id$/i.test(k));
    return (idKey && (foundItem[idKey] ?? decoded)) ?? decoded;
  }, [foundItem, decoded]);

  // Full-screen result component
  const ResultFullScreen = () => {
    if (!resultView.open) return null;
    return (
      <div className={`fixed inset-0 z-50 flex flex-col ${resultView.success ? "bg-green-600" : "bg-red-600"} text-white`}>
        <header className="px-4 pt-6 pb-4 flex items-center justify-between">
          <button onClick={() => nav("/")} className="py-2 px-3 rounded-md bg-white/20 hover:bg-white/30">Home</button>
          <h2 className="text-lg font-semibold">{resultView.success ? "Success" : "Error"}</h2>
          <div style={{ width: 56 }} />
        </header>

        <main className="flex-1 px-6 py-8 flex flex-col items-center justify-center text-center">
          <div className="mb-6">
            {resultView.success ? (
              <Check className="h-20 w-20" />
            ) : (
              <X className="h-20 w-20" />
            )}
          </div>
          <h3 className="text-2xl font-bold">{resultView.title}</h3>
          <p className="mt-3 text-base opacity-90">{resultView.message}</p>
          {resultView.details ? <p className="mt-2 text-sm opacity-80">{resultView.details}</p> : null}

          {/* Helpful contextual info */}
          {(decoded || foundCategory || selectedColumn) && (
            <div className="mt-6 bg-white/10 rounded-lg p-4 w-full max-w-md text-left">
              {decoded && (
                <>
                  <div className="text-sm opacity-80">UUID</div>
                  <div className="mt-1 font-mono break-words">{decoded}</div>
                </>
              )}
              {foundCategory && (
                <>
                  <div className="mt-4 text-sm opacity-80">Category</div>
                  <div className="mt-1">{foundCategory}</div>
                </>
              )}
              {selectedColumn && (
                <>
                  <div className="mt-4 text-sm opacity-80">Column</div>
                  <div className="mt-1">{selectedColumn}</div>
                </>
              )}
              <div className="mt-4 text-sm opacity-80">Amount</div>
              <div className="mt-1">+{Math.max(1, appliedQuantity)}</div>
            </div>
          )}

          <div className="mt-8 flex flex-col gap-3 w-full max-w-md">
            <button
              onClick={closeResultAndReset}
              className="py-4 rounded-lg bg-white text-black font-semibold"
            >
              OK
            </button>
            {!resultView.success && (
              <button
                onClick={() => {
                  closeResultAndReset();
                  setScanning(true); // immediately return to scanner
                }}
                className="py-4 rounded-lg bg-black/20 font-semibold"
              >
                Try again
              </button>
            )}
          </div>
        </main>
      </div>
    );
  };

  // Small loader component
  const InlineLoader = ({ text }: { text: string }) => (
    <div className="mt-3">
      <div className="w-full h-2 bg-gray-200 rounded overflow-hidden">
        <div className="h-2 bg-indigo-600 animate-pulse w-1/2 rounded" />
      </div>
      <div className="mt-2 text-sm text-gray-600 flex items-center gap-2">
        <svg className="animate-spin h-4 w-4 text-indigo-600" viewBox="0 0 24 24" fill="none">
          <circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" strokeOpacity="0.2" />
          <path d="M4 12a8 8 0 018-8" stroke="currentColor" strokeWidth="4" />
        </svg>
        {text}
      </div>
    </div>
  );

  // Quantity picker UI
  const QuantityPicker = () => (
    <div className="mt-2 flex items-center gap-3">
      <button
        type="button"
        aria-label="Decrease amount"
        onClick={() => setQuantity((q) => Math.max(1, q - 1))}
        disabled={updating}
        className="p-3 rounded-lg bg-gray-200 text-gray-800 disabled:opacity-50"
      >
        <Minus className="w-5 h-5" />
      </button>
      <div className="min-w-[64px] text-center py-2 px-4 rounded-lg bg-gray-50 border font-medium">
        {Math.max(1, quantity)}
      </div>
      <button
        type="button"
        aria-label="Increase amount"
        onClick={() => setQuantity((q) => q + 1)}
        disabled={updating}
        className="p-3 rounded-lg bg-gray-200 text-gray-800 disabled:opacity-50"
      >
        <Plus className="w-5 h-5" />
      </button>
    </div>
  );

  return (
    <div className="app-container px-4 pb-8">
      {/* Full-screen success/error view */}
      <ResultFullScreen />

      {!resultView.open && (
        <>
          <header className="flex items-center justify-between mt-6 mb-4">
            <button onClick={() => nav("/")} className="py-2 px-3 rounded-md bg-gray-100 text-gray-800">
              Back
            </button>
            <h2 className="text-lg font-medium">Scan Item</h2>
            <div style={{ width: 56 }} />
          </header>

          {/* Top button: choose column */}
          <div className="mb-4">
            <div className="flex gap-2">
              <button
                onClick={goSelectColumn}
                className="flex-1 px-4 py-3 rounded-lg bg-indigo-600 text-white flex items-center gap-2"
                aria-label="Choose column"
                title="Choose which column to increment after scan"
              >
                <span className="text-sm font-medium truncate">{selectedColumn ? selectedColumn : "Choose column"}</span>
                <ChevronRight className="w-4 h-4" />
              </button>
            </div>
            <div className="text-xs text-gray-500 mt-2">
              The UUID determines the category automatically. Selected column will be increased by the amount you choose.
            </div>
          </div>

          <main className="space-y-6">
            <div className="bg-white rounded-lg shadow p-4">
              {/* Start camera button */}
              {!scanning && !decoded && (
                <div className="flex flex-col items-center gap-4">
                  <button
                    onClick={() => {
                      setDecoded(null);
                      setFoundItem(null);
                      setFoundCategory(null);
                      setResultView({ open: false, success: false, title: "", message: "" });
                      setQuantity(1);
                      setAppliedQuantity(1);
                      setScanning(true);
                    }}
                    className="w-64 h-64 rounded-full bg-blue-600 flex flex-col items-center justify-center text-white shadow-lg"
                    aria-label="Start camera"
                  >
                    <Camera className="h-16 w-16" />
                    <div className="mt-3 text-lg font-medium">Start Camera</div>
                  </button>
                  <div className="text-sm text-gray-500 text-center px-4">Point the camera at the QR code. Capture runs automatically.</div>
                </div>
              )}

              {/* scanner view */}
              {scanning && !decoded && (
                <div className="flex flex-col items-stretch gap-3">
                  <div className="rounded overflow-hidden">
                    <QRScanner
                      ref={scannerRef}
                      onResult={onScanResult}
                      onStop={() => setScanning(false)}
                    />
                  </div>
                  <div className="flex gap-3">
                    <button
                      onClick={() => {
                        scannerRef.current?.stop();
                        setScanning(false);
                      }}
                      className="flex-1 py-3 rounded-lg bg-gray-200 text-gray-800"
                    >
                      Cancel
                    </button>
                    <div className="flex-1" />
                  </div>
                </div>
              )}

              {/* result / actions */}
              {decoded && (
                <div className="flex flex-col gap-4 items-stretch">
                  {/* Searching loader */}
                  {searching && <InlineLoader text="Searching for item…" />}

                  {!searching && (
                    <>
                      <div>
                        <div className="text-xs text-gray-500">Detected UUID</div>
                        <div className="mt-2 p-3 bg-gray-50 rounded text-base font-mono break-words">{decoded}</div>
                      </div>

                      <div>
                        <div className="text-xs text-gray-500">Found in category</div>
                        <div className="mt-2 p-3 bg-gray-50 rounded text-base">{foundCategory ?? "Not found"}</div>
                      </div>

                      <div>
                        <div className="text-xs text-gray-500">Item</div>
                        <div className="mt-2 p-3 bg-gray-50 rounded text-base">
                          {foundItem ? (
                            <div>
                              <div className="text-lg font-medium">{displayName}</div>
                              <div className="text-sm text-gray-500 mt-1">ID: {foundItem.id ?? foundItem.ID ?? decoded}</div>
                            </div>
                          ) : (
                            <div className="text-sm text-gray-500">Invalid or unknown ID</div>
                          )}
                        </div>
                      </div>

                      <div>
                        <div className="text-xs text-gray-500">Column to increment</div>
                        <div className="mt-2 p-3 bg-gray-50 rounded text-base">
                          {selectedColumn ? <div className="font-medium">{selectedColumn}</div> : <div className="text-sm text-gray-500">No column selected — tap Choose column</div>}
                        </div>
                      </div>

                      {/* Quantity picker */}
                      <div>
                        <label className="block text-sm text-gray-600">Amount to add</label>
                        <QuantityPicker />
                        <div className="mt-1 text-xs text-gray-500">Minimum 1. Defaults to 1.</div>
                      </div>

                      <div className="flex gap-3">
                        <button
                          onClick={handleConfirmIncrement}
                          disabled={!foundItem || !selectedColumn || updating}
                          className={`flex-1 py-4 rounded-lg ${(!foundItem || !selectedColumn || updating) ? "bg-gray-300 text-gray-700 cursor-not-allowed" : "bg-green-600 text-white hover:bg-green-700"}`}
                        >
                          {updating ? (
                            <span className="inline-flex items-center gap-2">
                              <svg className="animate-spin h-4 w-4 text-white" viewBox="0 0 24 24" fill="none">
                                <circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" strokeOpacity="0.4" />
                                <path d="M4 12a8 8 0 018-8" stroke="currentColor" strokeWidth="4" />
                              </svg>
                              Updating…
                            </span>
                          ) : (
                            <span className="inline-flex items-center justify-center gap-2">
                              <Check className="h-5 w-5" /> Confirm (+{Math.max(1, quantity)})
                            </span>
                          )}
                        </button>

                        <button onClick={handleCancel} disabled={updating} className="w-24 py-4 rounded-lg bg-red-600 text-white hover:bg-red-700 flex items-center justify-center">
                          <X className="h-5 w-5" />
                        </button>
                      </div>
                    </>
                  )}
                </div>
              )}
            </div>

            <div className="bg-white rounded-lg shadow p-4">
              <p className="text-sm text-gray-600">Quick actions</p>
              <div className="mt-3 flex flex-col gap-3">
                <Link to="/manual" className="block text-center py-3 rounded-lg bg-yellow-500 text-white font-medium">
                  Enter item manually
                </Link>
                <button
                  onClick={() => {
                    setDecoded(null);
                    setFoundItem(null);
                    setFoundCategory(null);
                    setScanning(false);
                    setQuantity(1);
                    setAppliedQuantity(1);
                  }}
                  className="block py-3 rounded-lg bg-gray-200 text-gray-800"
                >
                  Reset
                </button>
              </div>
            </div>
          </main>
        </>
      )}
    </div>
  );
}