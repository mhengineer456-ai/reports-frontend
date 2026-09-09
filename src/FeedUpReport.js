import React, { useState, useEffect, useMemo, useRef } from "react";
import { useHistory } from "react-router-dom";
import { fetchSheetDataFromBackend } from "./config";

const SPREADSHEET_ID = "1IMhmYlJ3s2PPRgEQs1Ikd4O1OBXK4EYL1oV_-kWAkyg";
const SHEET_NAME = "FeedUp";

function findCol(headers, keywords) {
  if (!Array.isArray(headers)) return -1;
  const clean = keywords.map((k) => k.toLowerCase().trim());
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || "").toLowerCase().trim();
    if (clean.includes(h)) return i;
  }
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || "").toLowerCase().trim();
    if (clean.some((k) => h.includes(k))) return i;
  }
  return -1;
}

// Reusable Multi-Select Dropdown Component
function MultiSelectDropdown({ label, options, selectedValues, onChange, themeColor = "#0284c7" }) {
  const [isOpen, setIsOpen] = useState(false);
  const [filterSearch, setFilterSearch] = useState("");
  const dropdownRef = useRef(null);

  useEffect(() => {
    const handleClickOutside = (event) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target)) {
        setIsOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  const toggleOption = (val) => {
    if (val === "all") {
      onChange([]);
      return;
    }
    if (selectedValues.includes(val)) {
      onChange(selectedValues.filter((v) => v !== val));
    } else {
      onChange([...selectedValues, val]);
    }
  };

  const isAllSelected = selectedValues.length === 0;

  const getDisplayText = () => {
    if (isAllSelected) return "All";
    if (selectedValues.length === 1) return selectedValues[0];
    return `${selectedValues.length} Selected`;
  };

  const filteredOptions = options.filter((opt) =>
    opt.label.toLowerCase().includes(filterSearch.toLowerCase())
  );

  return (
    <div ref={dropdownRef} style={{ position: "relative" }}>
      <label style={{ fontSize: "0.72rem", fontWeight: 800, color: "#64748b", textTransform: "uppercase", letterSpacing: "0.4px", marginBottom: "4px", display: "block" }}>
        {label}
      </label>
      <div
        onClick={() => setIsOpen(!isOpen)}
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          userSelect: "none",
          padding: "8px 12px",
          borderRadius: "10px",
          border: `1.5px solid ${selectedValues.length > 0 ? themeColor : "#cbd5e1"}`,
          background: selectedValues.length > 0 ? "rgba(2, 132, 199, 0.08)" : "#ffffff",
          fontWeight: selectedValues.length > 0 ? "800" : "600",
          color: selectedValues.length > 0 ? themeColor : "#1e293b",
          cursor: "pointer",
          fontSize: "0.85rem",
          transition: "all 0.18s"
        }}
      >
        <span style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: "140px" }}>
          {getDisplayText()}
        </span>
        <span style={{ fontSize: "0.75rem", transform: isOpen ? "rotate(180deg)" : "rotate(0deg)", transition: "transform 0.2s", opacity: 0.7 }}>
          ▼
        </span>
      </div>

      {isOpen && (
        <div
          style={{
            position: "absolute",
            top: "calc(100% + 4px)",
            left: 0,
            zIndex: 100,
            background: "#ffffff",
            border: "1.5px solid #cbd5e1",
            borderRadius: "14px",
            boxShadow: "0 10px 25px rgba(0, 0, 0, 0.15)",
            minWidth: "200px",
            maxHeight: "260px",
            overflowY: "auto",
            padding: "8px"
          }}
        >
          {options.length > 6 && (
            <input
              type="text"
              placeholder={`Search ${label}...`}
              value={filterSearch}
              onChange={(e) => setFilterSearch(e.target.value)}
              onClick={(e) => e.stopPropagation()}
              style={{
                width: "100%",
                padding: "6px 10px",
                borderRadius: "8px",
                border: "1px solid #cbd5e1",
                fontSize: "0.8rem",
                marginBottom: "6px",
                boxSizing: "border-box",
                outline: "none"
              }}
            />
          )}

          {/* Option: Select All / Clear */}
          <div
            onClick={() => onChange([])}
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              padding: "6px 10px",
              borderRadius: "8px",
              cursor: "pointer",
              fontSize: "0.8rem",
              fontWeight: 800,
              color: isAllSelected ? themeColor : "#64748b",
              background: isAllSelected ? "rgba(2, 132, 199, 0.08)" : "transparent",
              borderBottom: "1px solid #f1f5f9",
              marginBottom: "4px"
            }}
          >
            <span>Select All (Clear)</span>
            {isAllSelected && <span>✓</span>}
          </div>

          {filteredOptions.map((opt) => {
            const isChecked = selectedValues.includes(opt.value);
            return (
              <div
                key={opt.value}
                onClick={() => toggleOption(opt.value)}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: "8px",
                  padding: "6px 10px",
                  borderRadius: "8px",
                  cursor: "pointer",
                  fontSize: "0.82rem",
                  background: isChecked ? "rgba(2, 132, 199, 0.06)" : "transparent",
                  color: isChecked ? themeColor : "#334155",
                  fontWeight: isChecked ? "700" : "500",
                  transition: "background 0.15s"
                }}
              >
                <input
                  type="checkbox"
                  checked={isChecked}
                  onChange={() => {}}
                  style={{ accentColor: themeColor, cursor: "pointer" }}
                />
                <span style={{ flex: 1 }}>{opt.label}</span>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

export default function FeedUpReport() {
  const history = useHistory();
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  // Filter States
  const [searchTerm, setSearchTerm] = useState("");
  const [selectedStatuses, setSelectedStatuses] = useState([]);
  const [selectedBrands, setSelectedBrands] = useState([]);
  const [selectedGarments, setSelectedGarments] = useState([]);
  const [selectedFabrics, setSelectedFabrics] = useState([]);
  const [selectedSupervisors, setSelectedSupervisors] = useState([]);
  const [dateFilter, setDateFilter] = useState("");

  const parseCompletionDate = (raw) => {
    if (!raw || raw === "[]" || raw === "-") return "";
    const s = String(raw).trim();
    if (s.startsWith("[") || s.startsWith("{")) {
      try {
        const parsed = JSON.parse(s);
        if (Array.isArray(parsed) && parsed.length > 0) {
          const last = parsed[parsed.length - 1];
          return last.timestamp || last.date || last.completionDate || last.completedDate || "";
        } else if (parsed && typeof parsed === "object") {
          return parsed.timestamp || parsed.date || parsed.completionDate || "";
        }
      } catch (e) {
        // use raw string
      }
    }
    return s;
  };

  const parseWipRemarks = (raw) => {
    if (!raw || raw === "[]") return "In Progress";
    const s = String(raw).trim();
    if (s.startsWith("[") || s.startsWith("{")) {
      try {
        const parsed = JSON.parse(s);
        if (Array.isArray(parsed) && parsed.length > 0) {
          const last = parsed[parsed.length - 1];
          return last.remarks || last.status || last.updateType || "In Progress";
        }
      } catch (e) {
        // use raw string
      }
    }
    return s;
  };

  const fetchData = async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await fetchSheetDataFromBackend(SPREADSHEET_ID, `'${SHEET_NAME}'!A:Z`);
      if (!res.ok || !Array.isArray(res.values) || res.values.length === 0) {
        setData([]);
        setLoading(false);
        return;
      }

      const rows = res.values;
      const headers = rows[0] || [];

      const tsIdx = findCol(headers, ["timestamp", "time"]);
      const lotIdx = findCol(headers, ["lot number", "lot no", "lot"]);
      const garmentIdx = findCol(headers, ["garment type", "garment"]);
      const fabricIdx = findCol(headers, ["fabric"]);
      const styleIdx = findCol(headers, ["style"]);
      const brandIdx = findCol(headers, ["brand"]);
      const supIdx = findCol(headers, ["feed up supervisor", "supervisor"]);
      const dateIdx = findCol(headers, ["feed up date", "date"]);
      const pcsIdx = findCol(headers, ["total pcs", "pcs", "quantity"]);
      const wipIdx = findCol(headers, ["wip feed up", "wip", "remarks"]);
      const compIdx = findCol(headers, ["feed up completed", "feed up complete", "complete", "completed"]);

      const parsed = [];
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const lotNo = String(row[lotIdx !== -1 ? lotIdx : 1] || "").trim();
        if (!lotNo) continue;

        const pcs = parseInt(row[pcsIdx !== -1 ? pcsIdx : 8], 10) || 0;
        const rawComp = String(row[compIdx !== -1 ? compIdx : 10] || "").trim();
        const rawWip = String(row[wipIdx !== -1 ? wipIdx : 9] || "").trim();

        const completeDate = parseCompletionDate(rawComp);
        const isCompleted = !!completeDate && completeDate !== "-" && !completeDate.toLowerCase().includes("pending");
        const wipText = parseWipRemarks(rawWip);

        parsed.push({
          id: i,
          timestamp: row[tsIdx !== -1 ? tsIdx : 0] || "",
          lotNumber: lotNo,
          garmentType: row[garmentIdx !== -1 ? garmentIdx : 2] || "",
          fabric: row[fabricIdx !== -1 ? fabricIdx : 3] || "",
          style: row[styleIdx !== -1 ? styleIdx : 4] || "",
          brand: row[brandIdx !== -1 ? brandIdx : 5] || "",
          supervisor: row[supIdx !== -1 ? supIdx : 6] || "Feed Up Department",
          feedUpDate: row[dateIdx !== -1 ? dateIdx : 7] || "",
          totalPcs: pcs,
          wip: wipText,
          completeDate,
          isCompleted
        });
      }

      setData(parsed);
    } catch (err) {
      console.error("Error loading Feed Up data:", err);
      setError(err.message || "Failed to load sheet data");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchData();
  }, []);

  // Multi-Select Options
  const statusOptions = [
    { value: "pending", label: "⏳ In Progress / Pending" },
    { value: "completed", label: "✓ Completed" }
  ];

  const brandOptions = useMemo(() => {
    const set = new Set(data.map((d) => d.brand).filter(Boolean));
    return Array.from(set).sort().map((b) => ({ value: b, label: b }));
  }, [data]);

  const garmentOptions = useMemo(() => {
    const set = new Set(data.map((d) => d.garmentType).filter(Boolean));
    return Array.from(set).sort().map((g) => ({ value: g, label: g }));
  }, [data]);

  const fabricOptions = useMemo(() => {
    const set = new Set(data.map((d) => d.fabric).filter(Boolean));
    return Array.from(set).sort().map((f) => ({ value: f, label: f }));
  }, [data]);

  const supervisorOptions = useMemo(() => {
    const set = new Set(data.map((d) => d.supervisor).filter(Boolean));
    return Array.from(set).sort().map((s) => ({ value: s, label: s }));
  }, [data]);

  // Filtered dataset
  const filteredData = useMemo(() => {
    return data.filter((item) => {
      const term = searchTerm.toLowerCase().trim();
      const matchesSearch =
        !term ||
        item.lotNumber.toLowerCase().includes(term) ||
        item.style.toLowerCase().includes(term) ||
        item.fabric.toLowerCase().includes(term) ||
        item.brand.toLowerCase().includes(term) ||
        item.garmentType.toLowerCase().includes(term) ||
        item.supervisor.toLowerCase().includes(term) ||
        item.wip.toLowerCase().includes(term);

      const matchesStatus =
        selectedStatuses.length === 0 ||
        (selectedStatuses.includes("completed") && item.isCompleted) ||
        (selectedStatuses.includes("pending") && !item.isCompleted);

      const matchesBrand = selectedBrands.length === 0 || selectedBrands.some((b) => b.toLowerCase() === item.brand.toLowerCase());
      const matchesGarment = selectedGarments.length === 0 || selectedGarments.some((g) => g.toLowerCase() === item.garmentType.toLowerCase());
      const matchesFabric = selectedFabrics.length === 0 || selectedFabrics.some((f) => f.toLowerCase() === item.fabric.toLowerCase());
      const matchesSupervisor = selectedSupervisors.length === 0 || selectedSupervisors.some((s) => s.toLowerCase() === item.supervisor.toLowerCase());
      const matchesDate = !dateFilter || item.feedUpDate.includes(dateFilter);

      return matchesSearch && matchesStatus && matchesBrand && matchesGarment && matchesFabric && matchesSupervisor && matchesDate;
    });
  }, [data, searchTerm, selectedStatuses, selectedBrands, selectedGarments, selectedFabrics, selectedSupervisors, dateFilter]);

  const resetFilters = () => {
    setSearchTerm("");
    setSelectedStatuses([]);
    setSelectedBrands([]);
    setSelectedGarments([]);
    setSelectedFabrics([]);
    setSelectedSupervisors([]);
    setDateFilter("");
  };

  const hasActiveFilters =
    searchTerm !== "" ||
    selectedStatuses.length > 0 ||
    selectedBrands.length > 0 ||
    selectedGarments.length > 0 ||
    selectedFabrics.length > 0 ||
    selectedSupervisors.length > 0 ||
    dateFilter !== "";

  // Summary Metrics
  const totalLots = filteredData.length;
  const totalPcs = filteredData.reduce((sum, item) => sum + item.totalPcs, 0);
  const completedLots = filteredData.filter((item) => item.isCompleted).length;
  const pendingLots = totalLots - completedLots;

  // CSV Export
  const exportToCSV = () => {
    if (filteredData.length === 0) return;
    const headers = ["Timestamp", "Lot Number", "Garment Type", "Fabric", "Style", "Brand", "Feed Up Supervisor", "Feed Up Date", "Total Pcs", "WIP Remarks", "Status", "Completion Date"];
    const rows = filteredData.map((d) => [
      `"${d.timestamp}"`,
      `"${d.lotNumber}"`,
      `"${d.garmentType}"`,
      `"${d.fabric}"`,
      `"${d.style}"`,
      `"${d.brand}"`,
      `"${d.supervisor}"`,
      `"${d.feedUpDate}"`,
      d.totalPcs,
      `"${d.wip}"`,
      d.isCompleted ? "Completed" : "Pending",
      `"${d.completeDate}"`
    ]);
    const csvContent = "data:text/csv;charset=utf-8," + [headers.join(","), ...rows.map((r) => r.join(","))].join("\n");
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", `Feed_Up_Report_${new Date().toISOString().slice(0, 10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const formatDisplayDate = (val) => {
    if (!val) return "—";
    try {
      const d = new Date(val);
      if (!isNaN(d.getTime())) {
        const dd = String(d.getDate()).padStart(2, "0");
        const mm = String(d.getMonth() + 1).padStart(2, "0");
        const yy = String(d.getFullYear()).slice(-2);
        return `${dd}/${mm}/${yy}`;
      }
    } catch (e) {}
    return val;
  };

  return (
    <>
      <style>{`
        .feedup-report-container {
          padding: 28px 32px 80px;
          background-color: #f8fafc;
          min-height: 100vh;
          font-family: 'Plus Jakarta Sans', 'Inter', system-ui, -apple-system, sans-serif;
          color: #0f172a;
          box-sizing: border-box;
        }
        .report-header-box {
          background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%);
          border-radius: 20px;
          padding: 24px 32px;
          display: flex;
          justify-content: space-between;
          align-items: center;
          color: white;
          box-shadow: 0 10px 25px -5px rgba(2, 132, 199, 0.25);
          margin-bottom: 24px;
        }
        .header-btn {
          background: rgba(255, 255, 255, 0.15);
          backdrop-filter: blur(8px);
          border: 1px solid rgba(255, 255, 255, 0.25);
          color: white;
          padding: 9px 18px;
          border-radius: 12px;
          font-weight: 700;
          font-size: 0.85rem;
          cursor: pointer;
          transition: all 0.2s ease;
          display: inline-flex;
          align-items: center;
          gap: 6px;
        }
        .header-btn:hover {
          background: rgba(255, 255, 255, 0.28);
          transform: translateY(-1px);
        }
        .stat-cards-grid {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
          gap: 16px;
          margin-bottom: 24px;
        }
        .stat-card {
          background: white;
          border-radius: 16px;
          padding: 18px 22px;
          border: 1px solid #e2e8f0;
          box-shadow: 0 2px 4px rgba(0, 0, 0, 0.02);
          transition: transform 0.2s;
        }
        .stat-card:hover {
          transform: translateY(-2px);
          box-shadow: 0 8px 16px rgba(0, 0, 0, 0.05);
        }
        .stat-label {
          font-size: 0.76rem;
          font-weight: 700;
          color: #64748b;
          text-transform: uppercase;
          letter-spacing: 0.5px;
          margin-bottom: 6px;
        }
        .stat-value {
          font-size: 1.8rem;
          font-weight: 800;
          line-height: 1.2;
        }
        .filter-panel-card {
          background: white;
          border-radius: 16px;
          padding: 20px 24px;
          border: 1px solid #e2e8f0;
          box-shadow: 0 2px 8px rgba(0, 0, 0, 0.02);
          margin-bottom: 24px;
        }
        .filter-grid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(170px, 1fr));
          gap: 14px;
          margin-top: 14px;
        }
        .filter-input {
          width: 100%;
          padding: 8px 12px;
          border-radius: 10px;
          border: 1.5px solid #cbd5e1;
          font-size: 0.85rem;
          outline: none;
          box-sizing: border-box;
          font-weight: 600;
          color: #1e293b;
        }
        .filter-input:focus {
          border-color: #0284c7;
        }
        .filter-label {
          font-size: 0.72rem;
          font-weight: 800;
          color: #64748b;
          text-transform: uppercase;
          letter-spacing: 0.4px;
          margin-bottom: 4px;
          display: block;
        }
        .data-table-card {
          background: white;
          border-radius: 16px;
          border: 1px solid #e2e8f0;
          overflow: hidden;
          box-shadow: 0 4px 12px rgba(0, 0, 0, 0.03);
        }
        .feedup-table {
          width: 100%;
          border-collapse: collapse;
          text-align: left;
          font-size: 0.84rem;
        }
        .feedup-table th {
          background: #f1f5f9;
          padding: 12px 14px;
          font-size: 0.74rem;
          font-weight: 800;
          color: #475569;
          text-transform: uppercase;
          letter-spacing: 0.5px;
          border-bottom: 2px solid #e2e8f0;
          white-space: nowrap;
        }
        .feedup-table td {
          padding: 12px 14px;
          border-bottom: 1px solid #f1f5f9;
          color: #1e293b;
          font-weight: 600;
        }
        .feedup-table tbody tr:hover {
          background-color: #f8fafc;
        }
        .status-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 3px 9px;
          border-radius: 20px;
          font-size: 0.72rem;
          font-weight: 800;
          text-transform: uppercase;
        }
        .status-badge.completed {
          background: #dcfce7;
          color: #15803d;
          border: 1px solid #bbf7d0;
        }
        .status-badge.pending {
          background: #fef3c7;
          color: #b45309;
          border: 1px solid #fde68a;
        }
      `}</style>

      <div className="feedup-report-container">
        {/* Header Box */}
        <div className="report-header-box">
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
              <span style={{ fontSize: "1.8rem" }}>🧵</span>
              <h1 style={{ margin: 0, fontSize: "1.45rem", fontWeight: 800, color: "#ffffff" }}>
                Feed Up Department Report
              </h1>
              <span style={{ background: "#0284c7", color: "white", padding: "2px 8px", borderRadius: "8px", fontSize: "0.72rem", fontWeight: 800, border: "1px solid rgba(255,255,255,0.4)" }}>
                LIVE GOOGLE SHEET
              </span>
            </div>
            <p style={{ margin: "6px 0 0 0", color: "#e0f2fe", fontSize: "0.85rem" }}>
              Live production data linked from <strong>OVERLOCK..FOLDING..KAJBUTTON</strong> (`FeedUp` Tab)
            </p>
          </div>

          <div style={{ display: "flex", gap: "10px" }}>
            <button className="header-btn" onClick={() => history.push("/dashboard")}>
              ← Dashboard
            </button>
            <button className="header-btn" onClick={() => history.push("/stitching-complete-lot")}>
              📊 Stitching Report
            </button>
            <button className="header-btn" onClick={fetchData}>
              ↻ Refresh
            </button>
            <button className="header-btn" onClick={exportToCSV}>
              📥 Export CSV
            </button>
            <button className="header-btn" onClick={() => window.print()}>
              🖨️ Print
            </button>
          </div>
        </div>

        {/* Stats Grid */}
        <div className="stat-cards-grid">
          <div className="stat-card">
            <div className="stat-label">Total Feed Up Lots</div>
            <div className="stat-value" style={{ color: "#0284c7" }}>{totalLots}</div>
          </div>
          <div className="stat-card">
            <div className="stat-label">Total Pieces</div>
            <div className="stat-value" style={{ color: "#3b82f6" }}>{totalPcs.toLocaleString()}</div>
          </div>
          <div className="stat-card">
            <div className="stat-label">Completed Lots</div>
            <div className="stat-value" style={{ color: "#10b981" }}>{completedLots}</div>
          </div>
          <div className="stat-card">
            <div className="stat-label">Pending In Feed Up</div>
            <div className="stat-value" style={{ color: "#f59e0b" }}>{pendingLots}</div>
          </div>
        </div>

        {/* Filters Toolbar with Multi-Select Dropdowns */}
        <div className="filter-panel-card">
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "12px" }}>
            <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
              <span style={{ fontSize: "1.1rem" }}>🔍</span>
              <span style={{ fontWeight: 800, fontSize: "0.95rem", color: "#0f172a" }}>Multi-Select Filter & Search Toolbar</span>
            </div>
            {hasActiveFilters && (
              <button
                onClick={resetFilters}
                style={{
                  background: "#fee2e2",
                  color: "#991b1b",
                  border: "1px solid #fecaca",
                  padding: "4px 12px",
                  borderRadius: "8px",
                  fontSize: "0.78rem",
                  fontWeight: 800,
                  cursor: "pointer"
                }}
              >
                ✕ Clear All Filters
              </button>
            )}
          </div>

          <div className="filter-grid">
            {/* Quick Search */}
            <div>
              <label className="filter-label">Search Keyword</label>
              <input
                type="text"
                placeholder="Search Lot, Style, Fabric..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="filter-input"
              />
            </div>

            {/* Status Filter (Multi-Select) */}
            <MultiSelectDropdown
              label="Status"
              options={statusOptions}
              selectedValues={selectedStatuses}
              onChange={setSelectedStatuses}
            />

            {/* Brand Filter (Multi-Select) */}
            <MultiSelectDropdown
              label="Brand"
              options={brandOptions}
              selectedValues={selectedBrands}
              onChange={setSelectedBrands}
            />

            {/* Garment Type Filter (Multi-Select) */}
            <MultiSelectDropdown
              label="Garment Type"
              options={garmentOptions}
              selectedValues={selectedGarments}
              onChange={setSelectedGarments}
            />

            {/* Fabric Filter (Multi-Select) */}
            <MultiSelectDropdown
              label="Fabric"
              options={fabricOptions}
              selectedValues={selectedFabrics}
              onChange={setSelectedFabrics}
            />

            {/* Supervisor Filter (Multi-Select) */}
            <MultiSelectDropdown
              label="Supervisor"
              options={supervisorOptions}
              selectedValues={selectedSupervisors}
              onChange={setSelectedSupervisors}
            />

            {/* Feed Up Date Filter */}
            <div>
              <label className="filter-label">Feed Up Date</label>
              <input
                type="date"
                value={dateFilter}
                onChange={(e) => setDateFilter(e.target.value)}
                className="filter-input"
              />
            </div>
          </div>
        </div>

        {/* Data Table */}
        <div className="data-table-card">
          <div style={{ padding: "16px 20px", borderBottom: "1px solid #e2e8f0", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span style={{ fontWeight: 800, fontSize: "1.05rem" }}>Feed Up Lots</span>
            <span style={{ fontSize: "0.82rem", fontWeight: 700, color: "#64748b" }}>
              Showing {filteredData.length} of {data.length} lots
            </span>
          </div>

          {loading ? (
            <div style={{ padding: "40px", textAlign: "center" }}>
              <div style={{ fontWeight: 700, color: "#64748b" }}>Loading Feed Up records from Google Sheet...</div>
            </div>
          ) : error ? (
            <div style={{ padding: "40px", textAlign: "center", color: "#dc2626" }}>
              <div style={{ fontWeight: 700 }}>Error loading Feed Up data</div>
              <div style={{ fontSize: "0.82rem", marginTop: "4px" }}>{error}</div>
            </div>
          ) : filteredData.length === 0 ? (
            <div style={{ padding: "40px", textAlign: "center", color: "#64748b" }}>
              <div style={{ fontWeight: 700, fontSize: "1.1rem" }}>No matching lots found</div>
              <div style={{ fontSize: "0.82rem", marginTop: "4px" }}>Try adjusting your search or filters</div>
            </div>
          ) : (
            <div style={{ overflowX: "auto" }}>
              <table className="feedup-table">
                <thead>
                  <tr>
                    <th>#</th>
                    <th>Lot Number</th>
                    <th>Garment Type</th>
                    <th>Fabric</th>
                    <th>Style</th>
                    <th>Brand</th>
                    <th>Feed Up Supervisor</th>
                    <th>Feed Up Date</th>
                    <th>Total Pcs</th>
                    <th>WIP Status / Remarks</th>
                    <th>Status</th>
                    <th>Completion Date</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredData.map((item, idx) => (
                    <tr key={item.id || idx}>
                      <td style={{ color: "#94a3b8" }}>{idx + 1}</td>
                      <td>
                        <span style={{
                          background: "#e0f2fe",
                          color: "#0369a1",
                          padding: "3px 8px",
                          borderRadius: "6px",
                          fontWeight: 800
                        }}>
                          {item.lotNumber}
                        </span>
                      </td>
                      <td>{item.garmentType || "—"}</td>
                      <td>{item.fabric || "—"}</td>
                      <td style={{ fontWeight: 700 }}>{item.style || "—"}</td>
                      <td>{item.brand || "—"}</td>
                      <td>
                        <span style={{
                          display: "inline-flex",
                          alignItems: "center",
                          gap: "4px",
                          color: "#334155"
                        }}>
                          👤 {item.supervisor}
                        </span>
                      </td>
                      <td>{formatDisplayDate(item.feedUpDate)}</td>
                      <td>
                        <span style={{ fontWeight: 800, color: "#0f172a" }}>
                          {item.totalPcs.toLocaleString()}
                        </span>
                      </td>
                      <td>
                        <span style={{
                          fontSize: "0.78rem",
                          color: item.wip.toLowerCase().includes("issue") || item.wip.toLowerCase().includes("hold") ? "#b91c1c" : "#475569"
                        }}>
                          {item.wip}
                        </span>
                      </td>
                      <td>
                        {item.isCompleted ? (
                          <span className="status-badge completed">
                            ✓ Done
                          </span>
                        ) : (
                          <span className="status-badge pending">
                            ⏳ WIP
                          </span>
                        )}
                      </td>
                      <td>
                        {item.isCompleted ? (
                          <span style={{ color: "#15803d", fontWeight: 700 }}>
                            {formatDisplayDate(item.completeDate)}
                          </span>
                        ) : (
                          <span style={{ color: "#94a3b8" }}>—</span>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
    </>
  );
}
