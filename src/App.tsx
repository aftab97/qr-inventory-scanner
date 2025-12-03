import { Routes, Route } from "react-router-dom";
import Home from "./pages/home";
import Scan from "./pages/scan";
import ManualEntry from "./pages/manual-entry";

export default function App() {
  return (
    <div className="min-h-screen bg-gray-50">
      <Routes>
        <Route path="/" element={<Home />} />
        <Route path="/scan" element={<Scan />} />
        <Route path="/manual" element={<ManualEntry />} />
      </Routes>
    </div>
  );
}