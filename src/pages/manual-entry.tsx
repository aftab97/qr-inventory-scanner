/* eslint-disable @typescript-eslint/no-explicit-any */
import { useState } from "react";
import { useNavigate } from "react-router-dom";

export default function ManualEntry() {
  const [value, setValue] = useState("");
  const nav = useNavigate();

  const submit = (e: any) => {
    e.preventDefault();
    // replace with real submit logic (API call / state update)
    console.log("Manual entry submitted:", value);
    alert("Submitted: " + value);
    setValue("");
    nav("/");
  };

  return (
    <div className="app-container">
      <header className="flex items-center justify-between mt-6 mb-4">
        <button onClick={() => nav(-1)} className="py-2 px-3 rounded-md bg-gray-100 text-gray-800">Back</button>
        <h2 className="text-lg font-medium">Enter Item</h2>
        <div style={{ width: 56 }} />
      </header>

      <main>
        <form onSubmit={submit} className="bg-white rounded-lg shadow p-6 space-y-4">
          <label className="block text-sm font-medium text-gray-700">Item code or name</label>
          <input
            className="w-full p-4 border rounded-lg text-lg focus:ring-2 focus:ring-blue-400"
            value={value}
            onChange={(e) => setValue(e.target.value)}
            placeholder="Type item code or name"
            inputMode="text"
            required
          />

          <button type="submit" className="w-full py-4 rounded-lg bg-green-600 text-white text-lg font-medium">
            Submit
          </button>
        </form>
      </main>
    </div>
  );
}