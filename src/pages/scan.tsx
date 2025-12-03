import React, { useState } from "react";
import { Link, useNavigate } from "react-router-dom";
import QRScanner from "../components/qr/qr-scanner";

export default function Scan() {
  const [lastDecoded, setLastDecoded] = useState(null);
  const nav = useNavigate();

  return (
    <div className="app-container">
      <header className="flex items-center justify-between mt-6 mb-4">
        <button
          onClick={() => nav("/")}
          className="py-2 px-3 rounded-md bg-gray-100 text-gray-800"
        >
          Back
        </button>
        <h2 className="text-lg font-medium">Scan Item</h2>
        <div style={{ width: 56 }} />
      </header>

      <main className="space-y-4">
        <div className="bg-white rounded-lg shadow p-4">
          <QRScanner onResult={(text) => setLastDecoded(text)} />
        </div>

        <div className="bg-white rounded-lg shadow p-4">
          <p className="text-sm text-gray-600">Decoded result:</p>
          <div className="mt-2 p-3 bg-gray-50 rounded">{lastDecoded ?? <em className="text-gray-400">None yet</em>}</div>

          <div className="mt-4 flex gap-3">
            <Link to="/manual" className="flex-1 py-3 text-center rounded-lg bg-yellow-500 text-white font-medium">
              Enter item manually
            </Link>
            <button onClick={() => nav("/")} className="flex-1 py-3 text-center rounded-lg bg-gray-200 text-gray-800">
              Home
            </button>
          </div>
        </div>
      </main>
    </div>
  );
}