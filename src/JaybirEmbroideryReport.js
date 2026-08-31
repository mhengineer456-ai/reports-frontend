import React, { useState, useEffect, useMemo, useRef } from "react";
import { useHistory } from "react-router-dom";
import { fetchSheetDataFromBackend } from "./config";

const SPREADSHEET_ID = "1IMhmYlJ3s2PPRgEQs1Ikd4O1OBXK4EYL1oV_-kWAkyg";
const SHEET_NAME = "Jaybir Embroidery";

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
function MultiSelectDropdown({ label, options, selectedValues, onChange, themeColor = "#8b5cf6" }) {
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
          background: selectedValues.length > 0 ? "rgba(139, 92, 246, 0.08)" : "#ffffff",
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
              background: isAllSelected ? "rgba(139, 92, 246, 0.08)" : "transparent",
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
                  fontWeight: isChecked ? "800" : "500",
                  color: isChecked ? themeColor : "#1e293b",
                  background: isChecked ? "rgba(139, 92, 246, 0.08)" : "transparent",
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

export default function JaybirEmbroideryReport() {
  const history = useHistory();
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  // Multi-Select Filters (Array of strings)
  const [searchTerm, setSearchTerm] = useState("");
  const [selectedStatuses, setSelectedStatuses] = useState([]);
  const [selectedBrands, setSelectedBrands] = useState([]);
  const [selectedGarments, setSelectedGarments] = useState([]);
  const [selectedFabrics, setSelectedFabrics] = useState([]);
  const [selectedSupervisors, setSelectedSupervisors] = useState([]);
  const [dateFilter, setDateFilter] = useState("");

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
      const supIdx = findCol(headers, ["embroidery supervisor", "supervisor"]);
      const dateIdx = findCol(headers, ["embroidery date", "date"]);
      const pcsIdx = findCol(headers, ["total pcs", "pcs", "quantity"]);
      const wipIdx = findCol(headers, ["wip jaybir embroidery", "wip", "remarks"]);
      const compIdx = findCol(headers, ["jaybir embroidery complete", "complete", "completed"]);

      const parsed = [];
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const lotNo = String(row[lotIdx !== -1 ? lotIdx : 1] || "").trim();
        if (!lotNo) continue;

        const pcs = parseInt(row[pcsIdx !== -1 ? pcsIdx : 8], 10) || 0;
        const compVal = String(row[compIdx !== -1 ? compIdx : 10] || "").trim();
        const wipVal = String(row[wipIdx !== -1 ? wipIdx : 9] || "").trim();
        const isCompleted = !!compVal && compVal !== "[]" && compVal !== "-" && !compVal.toLowerCase().includes("pending");

        parsed.push({
          id: i,
          timestamp: row[tsIdx !== -1 ? tsIdx : 0] || "",
          lotNumber: lotNo,
          garmentType: row[garmentIdx !== -1 ? garmentIdx : 2] || "",
          fabric: row[fabricIdx !== -1 ? fabricIdx : 3] || "",
          style: row[styleIdx !== -1 ? styleIdx : 4] || "",
          brand: row[brandIdx !== -1 ? brandIdx : 5] || "",
          supervisor: row[supIdx !== -1 ? supIdx : 6] || "Jaybir Embroidery",
          embroideryDate: row[dateIdx !== -1 ? dateIdx : 7] || "",
          totalPcs: pcs,
          wip: wipVal === "[]" ? "In Progress" : wipVal || "In Progress",
          completeDate: compVal === "[]" ? "" : compVal,
          isCompleted
        });
      }

      setData(parsed);
    } catch (err) {
      console.error("Error loading Jaybir Embroidery data:", err);
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
      const matchesDate = !dateFilter || item.embroideryDate.includes(dateFilter);

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

  // Summary Metrics for filtered data
  const totalLots = filteredData.length;
  const totalPcs = filteredData.reduce((sum, item) => sum + item.totalPcs, 0);
  const completedLots = filteredData.filter((item) => item.isCompleted).length;
  const pendingLots = totalLots - completedLots;

  // CSV Export
  const exportToCSV = () => {
    if (filteredData.length === 0) return;
    const headers = ["Timestamp", "Lot Number", "Garment Type", "Fabric", "Style", "Brand", "Supervisor", "Embroidery Date", "Total Pcs", "WIP Remarks", "Status", "Completion Date"];
    const rows = filteredData.map((d) => [
      `"${d.timestamp}"`,
      `"${d.lotNumber}"`,
      `"${d.garmentType}"`,
      `"${d.fabric}"`,
      `"${d.style}"`,
      `"${d.brand}"`,
      `"${d.supervisor}"`,
      `"${d.embroideryDate}"`,
      d.totalPcs,
      `"${d.wip}"`,
      d.isCompleted ? "Completed" : "Pending",
      `"${d.completeDate}"`
    ]);
    const csvContent = "data:text/csv;charset=utf-8," + [headers.join(","), ...rows.map((r) => r.join(","))].join("\n");
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", `Jaybir_Embroidery_Report_${new Date().toISOString().slice(0, 10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  return (
    <>
      <style>{`
        .jaybir-report-container {
          padding: 28px 32px 80px;
          background-color: #f8fafc;
          min-height: 100vh;
          font-family: 'Plus Jakarta Sans', 'Inter', system-ui, -apple-system, sans-serif;
          color: #0f172a;
          box-sizing: border-box;
        }
        .report-header-box {
          background: linear-gradient(135deg, #2e1065 0%, #5b21b6 100%);
          border-radius: 20px;
          padding: 24px 32px;
          color: white;
          margin-bottom: 24px;
          box-shadow: 0 10px 25px -5px rgba(46, 16, 101, 0.2);
          display: flex;
          justify-content: space-between;
          align-items: center;
          flex-wrap: wrap;
          gap: 16px;
        }
        .header-btn {
          background: rgba(255, 255, 255, 0.15);
          border: 1px solid rgba(255, 255, 255, 0.25);
          color: white;
          padding: 8px 16px;
          border-radius: 10px;
          cursor: pointer;
          font-weight: 700;
          font-size: 0.85rem;
          display: flex;
          align-items: center;
          gap: 6px;
          transition: all 0.2s;
        }
        .header-btn:hover {
          background: rgba(255, 255, 255, 0.25);
        }
        .stat-cards-grid {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
          gap: 16px;
          margin-bottom: 24px;
        }
        .stat-card {
          background: white;
          padding: 20px 24px;
          border-radius: 16px;
          border: 1px solid #e2e8f0;
          box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        }
        .stat-label {
          font-size: 0.78rem;
          color: #64748b;
          font-weight: 800;
          text-transform: uppercase;
          letter-spacing: 0.5px;
          margin-bottom: 6px;
        }
        .stat-value {
          font-size: 1.7rem;
          font-weight: 900;
          color: #0f172a;
        }
        .filter-panel-card {
          background: white;
          border-radius: 18px;
          border: 1.5px solid #e2e8f0;
          box-shadow: 0 4px 10px rgba(0, 0, 0, 0.03);
          padding: 20px 24px;
          margin-bottom: 24px;
        }
        .filter-grid {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
          gap: 12px;
          margin-top: 14px;
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
        .filter-input {
          width: 100%;
          padding: 8px 12px;
          border-radius: 10px;
          border: 1.5px solid #cbd5e1;
          font-size: 0.85rem;
          font-weight: 600;
          color: #1e293b;
          background: #ffffff;
          outline: none;
          box-sizing: border-box;
          transition: border-color 0.2s;
        }
        .filter-input:focus {
          border-color: #8b5cf6;
        }
        .table-card {
          background: white;
          border-radius: 20px;
          border: 1px solid #e2e8f0;
          box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.05);
          overflow: hidden;
        }
        .table-header {
          padding: 18px 24px;
          border-bottom: 1px solid #f1f5f9;
          display: flex;
          justify-content: space-between;
          align-items: center;
          flex-wrap: wrap;
          gap: 12px;
        }
        .table-container {
          overflow: auto;
          max-height: calc(100vh - 420px);
          min-height: 320px;
        }
        .data-table {
          width: 100%;
          border-collapse: separate;
          border-spacing: 0;
        }
        .data-table th {
          background: #2e1065;
          color: white;
          padding: 14px 18px;
          text-align: left;
          font-weight: 800;
          font-size: 12px;
          text-transform: uppercase;
          letter-spacing: 0.6px;
          position: sticky;
          top: 0;
          z-index: 10;
        }
        .data-table td {
          padding: 13px 18px;
          border-bottom: 1px solid #f1f5f9;
          font-size: 13.5px;
          color: #1e293b;
        }
        .status-badge {
          display: inline-flex;
          align-items: center;
          gap: 4px;
          padding: 3px 10px;
          border-radius: 6px;
          font-size: 0.75rem;
          font-weight: 800;
        }
        .active-tag-pill {
          display: inline-flex;
          align-items: center;
          gap: 6px;
          background: rgba(139, 92, 246, 0.1);
          color: #6d28d9;
          border: 1px solid #ddd6fe;
          padding: 3px 10px;
          border-radius: 8px;
          font-size: 0.75rem;
          font-weight: 700;
        }
      `}</style>

      <div className="jaybir-report-container">
        {/* Header Box */}
        <div className="report-header-box">
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
              <span style={{ fontSize: "1.8rem" }}>🧵</span>
              <h1 style={{ margin: 0, fontSize: "1.45rem", fontWeight: 800, color: "#ffffff" }}>
                Jaybir Embroidery Report
              </h1>
              <span style={{ background: "#8b5cf6", color: "white", padding: "2px 8px", borderRadius: "8px", fontSize: "0.72rem", fontWeight: 800 }}>
                LIVE GOOGLE SHEET
              </span>
            </div>
            <p style={{ margin: "6px 0 0 0", color: "#cbd5e1", fontSize: "0.85rem" }}>
              Live production data linked from <strong>OVERLOCK..FOLDING..KAJBUTTON</strong> (`Jaybir Embroidery` Tab)
            </p>
          </div>

          <div style={{ display: "flex", gap: "10px" }}>
            <button className="header-btn" onClick={() => history.push("/dashboard")}>
              ← Dashboard
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
            <div className="stat-label">Total Embroidery Lots</div>
            <div className="stat-value" style={{ color: "#8b5cf6" }}>{totalLots}</div>
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
            <div className="stat-label">Pending In Embroidery</div>
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
              themeColor="#8b5cf6"
            />

            {/* Brand Filter (Multi-Select) */}
            <MultiSelectDropdown
              label="Brand"
              options={brandOptions}
              selectedValues={selectedBrands}
              onChange={setSelectedBrands}
              themeColor="#8b5cf6"
            />

            {/* Garment Type Filter (Multi-Select) */}
            <MultiSelectDropdown
              label="Garment Type"
              options={garmentOptions}
              selectedValues={selectedGarments}
              onChange={setSelectedGarments}
              themeColor="#8b5cf6"
            />

            {/* Fabric Filter (Multi-Select) */}
            <MultiSelectDropdown
              label="Fabric"
              options={fabricOptions}
              selectedValues={selectedFabrics}
              onChange={setSelectedFabrics}
              themeColor="#8b5cf6"
            />

            {/* Supervisor Filter (Multi-Select) */}
            <MultiSelectDropdown
              label="Supervisor"
              options={supervisorOptions}
              selectedValues={selectedSupervisors}
              onChange={setSelectedSupervisors}
              themeColor="#8b5cf6"
            />

            {/* Date Filter */}
            <div>
              <label className="filter-label">Embroidery Date</label>
              <input
                type="date"
                value={dateFilter}
                onChange={(e) => setDateFilter(e.target.value)}
                className="filter-input"
              />
            </div>
          </div>

          {/* Active Tags Pill Row */}
          {hasActiveFilters && (
            <div style={{ display: "flex", gap: "8px", flexWrap: "wrap", marginTop: "14px", paddingTop: "12px", borderTop: "1px solid #f1f5f9", alignItems: "center" }}>
              <span style={{ fontSize: "0.75rem", fontWeight: 800, color: "#64748b", textTransform: "uppercase" }}>Active:</span>
              {selectedStatuses.map((s) => (
                <span key={s} className="active-tag-pill">
                  Status: {s}
                  <span onClick={() => setSelectedStatuses(selectedStatuses.filter((x) => x !== s))} style={{ cursor: "pointer", fontWeight: 900 }}>×</span>
                </span>
              ))}
              {selectedBrands.map((b) => (
                <span key={b} className="active-tag-pill">
                  Brand: {b}
                  <span onClick={() => setSelectedBrands(selectedBrands.filter((x) => x !== b))} style={{ cursor: "pointer", fontWeight: 900 }}>×</span>
                </span>
              ))}
              {selectedGarments.map((g) => (
                <span key={g} className="active-tag-pill">
                  Garment: {g}
                  <span onClick={() => setSelectedGarments(selectedGarments.filter((x) => x !== g))} style={{ cursor: "pointer", fontWeight: 900 }}>×</span>
                </span>
              ))}
              {selectedFabrics.map((f) => (
                <span key={f} className="active-tag-pill">
                  Fabric: {f}
                  <span onClick={() => setSelectedFabrics(selectedFabrics.filter((x) => x !== f))} style={{ cursor: "pointer", fontWeight: 900 }}>×</span>
                </span>
              ))}
              {selectedSupervisors.map((s) => (
                <span key={s} className="active-tag-pill">
                  Supervisor: {s}
                  <span onClick={() => setSelectedSupervisors(selectedSupervisors.filter((x) => x !== s))} style={{ cursor: "pointer", fontWeight: 900 }}>×</span>
                </span>
              ))}
              {dateFilter && (
                <span className="active-tag-pill">
                  Date: {dateFilter}
                  <span onClick={() => setDateFilter("")} style={{ cursor: "pointer", fontWeight: 900 }}>×</span>
                </span>
              )}
            </div>
          )}
        </div>

        {/* Table Card */}
        <div className="table-card">
          <div className="table-header">
            <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
              <span style={{ fontWeight: 800, fontSize: "1.05rem" }}>Jaybir Embroidery Lots</span>
              <span style={{ background: "#f1f5f9", padding: "3px 10px", borderRadius: "6px", fontSize: "0.8rem", fontWeight: 800, color: "#334155" }}>
                Showing {filteredData.length} of {data.length} lots ({totalPcs.toLocaleString()} pcs)
              </span>
            </div>
          </div>

          <div className="table-container">
            {loading ? (
              <div style={{ textAlign: "center", padding: "60px" }}>
                <div style={{ fontSize: "2rem", marginBottom: "8px" }}>⚡</div>
                <div style={{ fontWeight: 700, color: "#64748b" }}>Loading Jaybir Embroidery records from Google Sheet...</div>
              </div>
            ) : filteredData.length === 0 ? (
              <div style={{ textAlign: "center", padding: "60px", color: "#64748b" }}>
                <div style={{ fontSize: "2.5rem", marginBottom: "8px" }}>📭</div>
                <div style={{ fontWeight: 800, fontSize: "1.1rem", color: "#0f172a" }}>No Embroidery Lots Found</div>
                <div style={{ fontSize: "0.85rem", marginTop: "4px" }}>
                  {hasActiveFilters ? "Try clearing or adjusting your multi-select filters." : "New entries from the Google Sheet tab will appear here."}
                </div>
              </div>
            ) : (
              <table className="data-table">
                <thead>
                  <tr>
                    <th>Timestamp</th>
                    <th>Lot #</th>
                    <th>Garment</th>
                    <th>Fabric</th>
                    <th>Style</th>
                    <th>Brand</th>
                    <th>Supervisor</th>
                    <th>Embroidery Date</th>
                    <th>Total Pcs</th>
                    <th>WIP Remarks</th>
                    <th>Status / Complete</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredData.map((item) => (
                    <tr key={item.id}>
                      <td style={{ fontSize: "0.8rem", color: "#64748b" }}>{item.timestamp}</td>
                      <td>
                        <strong style={{ color: "#5b21b6", fontSize: "0.95rem" }}>
                          #{item.lotNumber}
                        </strong>
                      </td>
                      <td>{item.garmentType || "-"}</td>
                      <td>{item.fabric || "-"}</td>
                      <td>{item.style || "-"}</td>
                      <td>{item.brand || "-"}</td>
                      <td>{item.supervisor || "-"}</td>
                      <td style={{ fontWeight: 700 }}>{item.embroideryDate || "-"}</td>
                      <td>
                        <span style={{ background: "#ede9fe", color: "#6d28d9", padding: "3px 8px", borderRadius: "6px", fontWeight: 800 }}>
                          {item.totalPcs}
                        </span>
                      </td>
                      <td style={{ fontSize: "0.82rem", color: "#64748b" }}>{item.wip}</td>
                      <td>
                        {item.isCompleted ? (
                          <span className="status-badge" style={{ background: "#d1fae5", color: "#047857" }}>
                            ✓ {item.completeDate || "Completed"}
                          </span>
                        ) : (
                          <span className="status-badge" style={{ background: "#fef3c7", color: "#b45309" }}>
                            ⏳ In Progress
                          </span>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </div>
      </div>
    </>
  );
}
