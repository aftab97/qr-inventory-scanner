import { Routes, Route } from "react-router-dom";
import Home from "./pages/home";
import Scan from "./pages/scan";
import ManualEntry from "./pages/manual-entry";
import UploadCsv from "./pages/upload-csv";
import Viewer from "./pages/viewer";
import ViewerWithCharts from "./pages/viewer-with-charts";
import SelectColumn from "./pages/select-column";

export default function App() {
  return (
    <div className="min-h-screen bg-gray-50">
      <Routes>
        <Route path="/" element={<Home />} />
        <Route path="/scan" element={<Scan />} />
        <Route path="/manual" element={<ManualEntry />} />
        <Route path="/csv" element={<UploadCsv />} />
        <Route path="/viewer" element={<Viewer />} />
        <Route path="/charts" element={<ViewerWithCharts />} />
        <Route path="/select-column" element={<SelectColumn />} />
      </Routes>
    </div>
  );
}