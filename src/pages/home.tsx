import React from "react";
import { useNavigate } from "react-router-dom";

export default function Home() {
  const nav = useNavigate();
  return (
    <div className="app-container max-w-3xl mx-auto px-4">
      <header className="flex items-center justify-center mt-8 mb-10">
        <h1 className="text-2xl font-semibold">QR Inventory</h1>
      </header>

      <main className="space-y-6">
        <div className="bg-white rounded-lg shadow p-6">
          <p className="text-gray-700 mb-4">Quick actions</p>

          <button
            onClick={() => nav("/scan")}
            className="w-full py-4 rounded-lg bg-blue-600 text-white text-lg font-medium shadow hover:bg-blue-700 focus:outline-none"
            aria-label="Scan item"
            type="button"
          >
            Scan item
          </button>

          <button
            onClick={() => nav("/manual")}
            className="w-full mt-4 py-4 rounded-lg bg-gray-200 text-gray-800 text-lg font-medium hover:bg-gray-300 focus:outline-none"
            aria-label="Enter item manually"
            type="button"
          >
            Enter item manually
          </button>

          <button
            onClick={() => nav("/viewer")}
            className="w-full mt-4 py-4 rounded-lg bg-green-600 text-white text-lg font-medium shadow hover:bg-green-700 focus:outline-none"
            aria-label="View inventory"
            type="button"
          >
            View inventory
          </button>

          <button
            onClick={() => nav("/csv")}
            className="w-full mt-4 py-4 rounded-lg bg-indigo-600 text-white text-lg font-medium shadow hover:bg-indigo-700 focus:outline-none"
            aria-label="Add inventory (via Excel)"
            type="button"
          >
            Add inventory (via Excel)
          </button>

          <button
            onClick={() => nav("/charts")}
            className="w-full mt-4 py-4 rounded-lg bg-green-600 text-white text-lg font-medium shadow hover:bg-green-700 focus:outline-none"
            aria-label="View inventory"
            type="button"
          >
            View inventory (Charts)
          </button>
        </div>

        <div className="text-sm text-gray-500">
          <p>Tips:</p>
          <ul className="list-disc ml-6">
            <li>Grant camera permission when prompted.</li>
            <li>Use the rear camera, enable torch if available for dim lighting.</li>
            <li>If scanning fails, use manual entry.</li>
          </ul>
        </div>
      </main>
    </div>
  );
}