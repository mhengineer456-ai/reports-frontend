// src/HoldLotsReport.js
import React, { useState, useEffect, useMemo, useRef } from "react";
import { useHistory } from "react-router-dom";
import * as XLSX from "xlsx-js-style";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import { SPREADSHEET_IDS, fetchSheetDataFromBackend, BACKEND_URL, HOLD_LOTS_WEBHOOK_URL } from "./config";
import { getCurrentUser } from "./auth";

/** Date formatting helper */
function formatDate(dateStr) {
  if (!dateStr || dateStr === "-" || dateStr === "N/A" || dateStr === "—") return "—";
  const str = String(dateStr).trim();
  if (str.includes("T")) {
    const d = new Date(str);
    if (!isNaN(d.getTime())) {
      const day = String(d.getDate()).padStart(2, "0");
      const month = String(d.getMonth() + 1).padStart(2, "0");
      const year = String(d.getFullYear()).slice(-2);
      const hours = String(d.getHours()).padStart(2, "0");
      const mins = String(d.getMinutes()).padStart(2, "0");
      return `${day}/${month}/${year} ${hours}:${mins}`;
    }
  }
  const parts = str.split(/[\/\-\.]/);
  if (parts.length === 3) {
    let day = parseInt(parts[0], 10);
    let month = parseInt(parts[1], 10);
    let year = parseInt(parts[2], 10);
    if (day > 1000) {
      const tmp = day; day = year; year = tmp;
    }
    if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
      const fullYear = year < 100 ? 2000 + year : year;
      return `${String(day).padStart(2, "0")}/${String(month).padStart(2, "0")}/${String(fullYear).slice(-2)}`;
    }
  }
  return str;
}

/** Reusable Multi-Select Dropdown Component */
function MultiSelectDropdown({ label, options, selectedValues, onChange, themeColor = "#b45309" }) {
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
    String(opt.label || "").toLowerCase().includes(filterSearch.toLowerCase())
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
          background: selectedValues.length > 0 ? "rgba(180, 83, 9, 0.08)" : "#ffffff",
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
            minWidth: "210px",
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
              background: isAllSelected ? "rgba(180, 83, 9, 0.08)" : "transparent",
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
                  background: isChecked ? "rgba(180, 83, 9, 0.08)" : "transparent",
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

export default function HoldLotsReport() {
  const history = useHistory();
  const currentUser = getCurrentUser();

  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [selectedLotModal, setSelectedLotModal] = useState(null);

  // Filters State
  const [searchTerm, setSearchTerm] = useState("");
  const [statusFilter, setStatusFilter] = useState("ALL"); // ALL, ON HOLD, RELEASED
  const [selectedDepts, setSelectedDepts] = useState([]);
  const [selectedPriorities, setSelectedPriorities] = useState([]);
  const [selectedBrands, setSelectedBrands] = useState([]);
  const [selectedStyles, setSelectedStyles] = useState([]);
  const [selectedFabrics, setSelectedFabrics] = useState([]);
  const [selectedParties, setSelectedParties] = useState([]);
  const [selectedLocations, setSelectedLocations] = useState([]);
  const [selectedReasons, setSelectedReasons] = useState([]);
  const [selectedHoldBy, setSelectedHoldBy] = useState([]);
  const [selectedApprovedBy, setSelectedApprovedBy] = useState([]);

  // Fetch Hold Lots Data
  const fetchData = async () => {
    setLoading(true);
    setError(null);
    try {
      // Step 1: Try backend /api/sheets/hold-lots
      try {
        const backendRes = await fetch(`${BACKEND_URL}/api/sheets/hold-lots?refresh=true`);
        if (backendRes.ok) {
          const json = await backendRes.json();
          if (json.success && Array.isArray(json.data) && json.data.length > 0) {
            setData(json.data);
            setLoading(false);
            return;
          }
        }
      } catch (backendErr) {
        console.warn("Backend hold lots fetch failed, using direct sheet fallback:", backendErr.message);
      }

      // Step 2: Fallback to direct sheet fetch across all candidate tabs
      const holdSpreadsheetId = SPREADSHEET_IDS.HOLD_LOTS || "1oBetbe44z2lUXngctvk3J31WBiWTv07NIgx5jFlylOs";
      const candidateTabs = ["All Holds", "Sheet1", "Stitching", "Cutting", "Hold lot", "Hold Lot", "Hold Lots", "Hold", "General"];

      const holdsMap = new Map();

      for (const tab of candidateTabs) {
        try {
          const sheetRes = await fetchSheetDataFromBackend(holdSpreadsheetId, `${tab}!A1:Z3000`, true);
          if (sheetRes && sheetRes.ok && Array.isArray(sheetRes.values) && sheetRes.values.length > 1) {
            const rows = sheetRes.values;
            const headers = rows[0].map((h) => String(h || "").trim());
            const normalize = (s) => String(s || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");

            const headerMap = {};
            headers.forEach((h, idx) => {
              headerMap[normalize(h)] = idx;
            });

            const getVal = (row, key) => {
              const idx = headerMap[normalize(key)];
              return idx !== undefined && row[idx] !== undefined ? String(row[idx]).trim() : "";
            };

            for (let i = 1; i < rows.length; i++) {
              const row = rows[i];
              if (!row || row.length === 0) continue;

              let lotNo = getVal(row, "Lot Number") || getVal(row, "Lot No") || getVal(row, "Lot") || (row[3] ? String(row[3]).trim() : "");
              if (!lotNo && !row[0] && !row[1]) continue;
              if (String(lotNo).toLowerCase() === "lot number" || String(row[0]).toLowerCase() === "timestamp") continue;

              const holdId = getVal(row, "Hold ID") || getVal(row, "ID") || row[1] || `HOLD-${i}`;
              const qtyNum = parseInt(getVal(row, "Quantity") || getVal(row, "Qty") || row[10] || 0, 10) || 0;

              let location = getVal(row, "Location") || getVal(row, "Hold Location") || getVal(row, "Placement") || getVal(row, "Floor") || "";
              let holdBy = getVal(row, "Hold By") || getVal(row, "Held By") || "";
              let approvedBy = getVal(row, "Approved By") || getVal(row, "Approval By") || "";
              let priority = getVal(row, "Priority") || getVal(row, "Urgency") || "";
              let status = getVal(row, "Status") || "";
              let releasedAt = getVal(row, "Released At") || getVal(row, "Release Date") || "";
              let releasedBy = getVal(row, "Released By") || getVal(row, "Release By") || "";

              for (let colIdx = 14; colIdx < row.length; colIdx++) {
                const val = String(row[colIdx] || "").trim();
                if (!val) continue;
                const valUpper = val.toUpperCase();
                if (valUpper === "ON HOLD" || valUpper === "RELEASED" || valUpper === "HOLD" || valUpper === "ACTIVE") {
                  if (!status) status = valUpper === "RELEASED" ? "RELEASED" : "ON HOLD";
                } else if (
                  valUpper.includes("ALERT") ||
                  valUpper.includes("STANDARD") ||
                  valUpper.includes("CRITICAL") ||
                  valUpper.includes("URGENT") ||
                  valUpper.includes("HIGH") ||
                  valUpper.includes("NORMAL")
                ) {
                  if (!priority) priority = val;
                } else if (valUpper.includes("FLOOR") || valUpper.includes("FLR") || valUpper.includes("RACK") || valUpper.includes("TABLE")) {
                  if (!location) location = val;
                }
              }

              if (!location && row[15]) location = String(row[15]).trim();
              if (!holdBy && row[16]) holdBy = String(row[16]).trim();
              if (!approvedBy && row[17]) approvedBy = String(row[17]).trim();
              if (!priority) priority = row[18] ? String(row[18]).trim() : "Standard Update";
              if (!status) status = row[19] ? String(row[19]).trim() : "ON HOLD";

              const isHold = status.toUpperCase().includes("HOLD");
              const key = `${holdId}_${lotNo}`.toLowerCase();

              if (!holdsMap.has(key)) {
                holdsMap.set(key, {
                  id: holdId,
                  timestamp: getVal(row, "Timestamp") || row[0] || "",
                  department: getVal(row, "Department") || getVal(row, "Dept") || row[2] || tab,
                  lotNumber: lotNo,
                  garmentType: getVal(row, "Garment Type") || getVal(row, "Garment") || getVal(row, "Item") || "—",
                  jobOrderNo: getVal(row, "Job Order No") || getVal(row, "JO") || row[4] || "—",
                  poDate: getVal(row, "PO Date") || getVal(row, "Date") || row[5] || "",
                  partyName: getVal(row, "Party Name") || getVal(row, "Party") || row[6] || "—",
                  brand: getVal(row, "Brand") || row[7] || "—",
                  style: getVal(row, "Style") || row[8] || "—",
                  fabric: getVal(row, "Fabric") || row[9] || "—",
                  section: getVal(row, "Section") || "—",
                  season: getVal(row, "Season") || "—",
                  directStitching: getVal(row, "Direct Stitching") || getVal(row, "Design Work") || "—",
                  quantity: qtyNum,
                  unit: getVal(row, "Unit") || row[11] || "PCS",
                  shade: getVal(row, "Shade") || getVal(row, "Color") || row[12] || "—",
                  size: getVal(row, "Size") || row[13] || "—",
                  reason: getVal(row, "Hold Reason") || getVal(row, "Reason") || getVal(row, "Remarks") || row[14] || "Under Review",
                  location: location || "—",
                  holdBy: holdBy || "—",
                  approvedBy: approvedBy || "—",
                  priority: priority || "Standard Update",
                  status: isHold ? "ON HOLD" : "RELEASED",
                  releasedAt: releasedAt || (row[20] ? String(row[20]).trim() : ""),
                  releasedBy: releasedBy || (row[21] ? String(row[21]).trim() : "")
                });
              }
            }
          }
        } catch (tabErr) {
          // Continue to next tab
        }
      }

      const allData = Array.from(holdsMap.values());
      setData(allData.reverse());
    } catch (err) {
      console.error("Error loading Hold Lots:", err);
      setError(err.message || "Failed to load hold lots data");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchData();
  }, []);

  // Options for Multi-Selects
  const departmentOptions = useMemo(() => {
    const set = new Set(data.map((d) => d.department).filter(Boolean));
    return Array.from(set).sort().map((d) => ({ value: d, label: d }));
  }, [data]);

  const priorityOptions = useMemo(() => {
    const set = new Set(data.map((d) => d.priority).filter(Boolean));
    return Array.from(set).sort().map((p) => ({ value: p, label: p }));
  }, [data]);

  const brandOptions = useMemo(() => {
    const set = new Set(data.map((d) => d.brand).filter((b) => b && b !== "—"));
    return Array.from(set).sort().map((b) => ({ value: b, label: b }));
  }, [data]);

  const styleOptions = useMemo(() => {
    const set = new Set(data.map((d) => d.style).filter((s) => s && s !== "—"));
    return Array.from(set).sort().map((s) => ({ value: s, label: s }));
  }, [data]);

  const fabricOptions = useMemo(() => {
    const set = new Set(data.map((d) => d.fabric).filter((f) => f && f !== "—"));
    return Array.from(set).sort().map((f) => ({ value: f, label: f }));
  }, [data]);

  const partyOptions = useMemo(() => {
    const set = new Set(data.map((d) => d.partyName).filter((p) => p && p !== "—"));
    return Array.from(set).sort().map((p) => ({ value: p, label: p }));
  }, [data]);

  const locationOptions = useMemo(() => {
    const set = new Set(data.map((d) => d.location).filter((l) => l && l !== "—"));
    return Array.from(set).sort().map((l) => ({ value: l, label: l }));
  }, [data]);

  const reasonOptions = useMemo(() => {
    const set = new Set(data.map((d) => d.reason).filter(Boolean));
    return Array.from(set).sort().map((r) => ({ value: r, label: r }));
  }, [data]);

  const holdByOptions = useMemo(() => {
    const set = new Set(data.map((d) => d.holdBy).filter(Boolean));
    return Array.from(set).sort().map((h) => ({ value: h, label: h }));
  }, [data]);

  const approvedByOptions = useMemo(() => {
    const set = new Set(data.map((d) => d.approvedBy).filter(Boolean));
    return Array.from(set).sort().map((a) => ({ value: a, label: a }));
  }, [data]);

  // Filtered Data Computation
  const filteredData = useMemo(() => {
    return data.filter((item) => {
      const term = searchTerm.toLowerCase().trim();
      const matchesSearch =
        !term ||
        item.lotNumber.toLowerCase().includes(term) ||
        item.jobOrderNo.toLowerCase().includes(term) ||
        item.brand.toLowerCase().includes(term) ||
        item.style.toLowerCase().includes(term) ||
        item.fabric.toLowerCase().includes(term) ||
        item.partyName.toLowerCase().includes(term) ||
        item.department.toLowerCase().includes(term) ||
        item.reason.toLowerCase().includes(term) ||
        item.location.toLowerCase().includes(term) ||
        item.holdBy.toLowerCase().includes(term) ||
        item.approvedBy.toLowerCase().includes(term) ||
        item.id.toLowerCase().includes(term);

      const matchesStatus =
        statusFilter === "ALL" ||
        (statusFilter === "ON HOLD" && item.status === "ON HOLD") ||
        (statusFilter === "RELEASED" && item.status === "RELEASED");

      const matchesDept = selectedDepts.length === 0 || selectedDepts.some((d) => d.toLowerCase() === item.department.toLowerCase());
      const matchesPriority = selectedPriorities.length === 0 || selectedPriorities.some((p) => p.toLowerCase() === item.priority.toLowerCase());
      const matchesBrand = selectedBrands.length === 0 || selectedBrands.some((b) => b.toLowerCase() === item.brand.toLowerCase());
      const matchesStyle = selectedStyles.length === 0 || selectedStyles.some((s) => s.toLowerCase() === item.style.toLowerCase());
      const matchesFabric = selectedFabrics.length === 0 || selectedFabrics.some((f) => f.toLowerCase() === item.fabric.toLowerCase());
      const matchesParty = selectedParties.length === 0 || selectedParties.some((p) => p.toLowerCase() === item.partyName.toLowerCase());
      const matchesLocation = selectedLocations.length === 0 || selectedLocations.some((l) => l.toLowerCase() === item.location.toLowerCase());
      const matchesReason = selectedReasons.length === 0 || selectedReasons.some((r) => r.toLowerCase() === item.reason.toLowerCase());
      const matchesHoldBy = selectedHoldBy.length === 0 || selectedHoldBy.some((h) => h.toLowerCase() === item.holdBy.toLowerCase());
      const matchesApprovedBy = selectedApprovedBy.length === 0 || selectedApprovedBy.some((a) => a.toLowerCase() === item.approvedBy.toLowerCase());

      return matchesSearch && matchesStatus && matchesDept && matchesPriority && matchesBrand && matchesStyle && matchesFabric && matchesParty && matchesLocation && matchesReason && matchesHoldBy && matchesApprovedBy;
    });
  }, [data, searchTerm, statusFilter, selectedDepts, selectedPriorities, selectedBrands, selectedStyles, selectedFabrics, selectedParties, selectedLocations, selectedReasons, selectedHoldBy, selectedApprovedBy]);

  const resetFilters = () => {
    setSearchTerm("");
    setStatusFilter("ALL");
    setSelectedDepts([]);
    setSelectedPriorities([]);
    setSelectedBrands([]);
    setSelectedStyles([]);
    setSelectedFabrics([]);
    setSelectedParties([]);
    setSelectedLocations([]);
    setSelectedReasons([]);
    setSelectedHoldBy([]);
    setSelectedApprovedBy([]);
  };

  const hasActiveFilters =
    searchTerm !== "" ||
    statusFilter !== "ALL" ||
    selectedDepts.length > 0 ||
    selectedPriorities.length > 0 ||
    selectedBrands.length > 0 ||
    selectedStyles.length > 0 ||
    selectedFabrics.length > 0 ||
    selectedParties.length > 0 ||
    selectedLocations.length > 0 ||
    selectedReasons.length > 0 ||
    selectedHoldBy.length > 0 ||
    selectedApprovedBy.length > 0;

  // KPI Metrics
  const totalHoldRecords = filteredData.length;
  const activeHolds = filteredData.filter((d) => d.status === "ON HOLD").length;
  const releasedHolds = filteredData.filter((d) => d.status === "RELEASED").length;
  const totalHoldQty = filteredData.reduce((sum, d) => sum + (d.quantity || 0), 0);
  const highAlertHolds = filteredData.filter((d) => d.status === "ON HOLD" && String(d.priority).toLowerCase().includes("high")).length;

  // CSV Export
  const exportToCSV = () => {
    if (filteredData.length === 0) return;
    const headers = [
      "Timestamp", "Hold ID", "Status", "Lot Number", "Garment Type", "Style", "Fabric", "Brand",
      "Quantity", "Section", "Season", "Party Name", "Direct Stitching", "Priority", "Department",
      "Job Order No", "PO Date", "Unit", "Shade", "Size", "Hold Reason", "Location",
      "Hold By", "Approved By", "Released At", "Released By"
    ];
    const rows = filteredData.map((d) => [
      `"${d.timestamp}"`,
      `"${d.id}"`,
      `"${d.status}"`,
      `"${d.lotNumber}"`,
      `"${d.garmentType || '—'}"`,
      `"${d.style}"`,
      `"${d.fabric}"`,
      `"${d.brand}"`,
      d.quantity,
      `"${d.section || '—'}"`,
      `"${d.season || '—'}"`,
      `"${d.partyName}"`,
      `"${d.directStitching || '—'}"`,
      `"${d.priority}"`,
      `"${d.department}"`,
      `"${d.jobOrderNo}"`,
      `"${d.poDate}"`,
      `"${d.unit}"`,
      `"${d.shade}"`,
      `"${d.size}"`,
      `"${d.reason}"`,
      `"${d.location}"`,
      `"${d.holdBy}"`,
      `"${d.approvedBy}"`,
      `"${d.releasedAt || ''}"`,
      `"${d.releasedBy || ''}"`
    ]);

    const csvContent = "data:text/csv;charset=utf-8," + [headers.join(","), ...rows.map((r) => r.join(","))].join("\n");
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", `Hold_Lots_Report_${new Date().toISOString().slice(0, 10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  // Excel Export with rich styling
  const exportToExcel = () => {
    if (filteredData.length === 0) return;

    const headers = [
      "Sr No.", "Lot Number", "Garment Type", "Style", "Fabric", "Brand",
      "Quantity", "Section", "Season", "Party Name", "Direct Stitching",
      "Hold ID", "Status", "Priority", "Department",
      "Job Order No", "PO Date", "Unit", "Shade", "Size", "Hold Reason", "Location",
      "Hold By", "Approved By", "Released At", "Released By"
    ];

    const rows = filteredData.map((d, index) => [
      index + 1,
      d.lotNumber,
      d.garmentType || "—",
      d.style,
      d.fabric,
      d.brand,
      d.quantity,
      d.section || "—",
      d.season || "—",
      d.partyName,
      d.directStitching || "—",
      d.id,
      d.status,
      d.priority,
      d.department,
      d.jobOrderNo,
      d.poDate ? formatDate(d.poDate) : "",
      d.unit,
      d.shade,
      d.size,
      d.reason,
      d.location,
      d.holdBy,
      d.approvedBy,
      d.releasedAt ? formatDate(d.releasedAt) : "",
      d.releasedBy || ""
    ]);

    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);

    // Auto-fit column widths
    const colWidths = headers.map((h, i) => {
      let maxLen = h.length;
      rows.forEach((r) => {
        const val = r[i] ? String(r[i]) : "";
        if (val.length > maxLen) maxLen = val.length;
      });
      return { wch: Math.min(Math.max(maxLen + 3, 10), 40) };
    });
    worksheet["!cols"] = colWidths;

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Hold Lots");

    const today = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(workbook, `Hold_Lots_Master_Report_${today}.xlsx`);
  };

  // PDF Export
  const exportToPDF = () => {
    if (filteredData.length === 0) {
      alert("No data available to download PDF.");
      return;
    }

    const doc = new jsPDF({
      orientation: "landscape",
      unit: "mm",
      format: "a3",
    });

    const headerColor = [124, 45, 18]; // Deep Amber / Maroon (#7c2d12)
    const textColor = [17, 24, 39];
    const borderColor = [226, 232, 240];
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const margin = 8;
    const contentWidth = pageWidth - margin * 2;

    const nowStr = new Date().toLocaleDateString("en-GB", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
    }) + " " + new Date().toLocaleTimeString("en-US", { hour: "2-digit", minute: "2-digit" });

    const drawHeader = () => {
      doc.setFillColor(124, 45, 18);
      doc.rect(margin, 6, contentWidth, 18, "F");

      doc.setFontSize(14);
      doc.setFont("helvetica", "bold");
      doc.setTextColor(255, 255, 255);
      doc.text("FACTORY SUITE PRO — HOLD LOTS AUDIT REPORT", margin + 6, 13);

      doc.setFontSize(9);
      doc.setFont("helvetica", "normal");
      doc.setTextColor(254, 243, 199);
      doc.text(
        `Generated: ${nowStr}  |  Total Records: ${totalHoldRecords}  |  Active On Hold: ${activeHolds}  |  Released: ${releasedHolds}  |  Total Hold Pcs: ${totalHoldQty.toLocaleString()}  |  High Alert: ${highAlertHolds}`,
        margin + 6,
        19.5
      );
    };

    const tableHeaders = [
      [
        { content: "Sr", styles: { halign: "center", fontStyle: "bold" } },
        { content: "Lot #", styles: { halign: "center", fontStyle: "bold" } },
        { content: "Garment Type", styles: { halign: "center", fontStyle: "bold" } },
        { content: "Style", styles: { halign: "center", fontStyle: "bold" } },
        { content: "Fabric", styles: { halign: "center", fontStyle: "bold" } },
        { content: "Brand", styles: { halign: "center", fontStyle: "bold" } },
        { content: "Qty", styles: { halign: "center", fontStyle: "bold" } },
        { content: "Section", styles: { halign: "center", fontStyle: "bold" } },
        { content: "Season", styles: { halign: "center", fontStyle: "bold" } },
        { content: "Party Name", styles: { halign: "center", fontStyle: "bold" } },
        { content: "Direct Stitching", styles: { halign: "center", fontStyle: "bold" } },
        { content: "Status", styles: { halign: "center", fontStyle: "bold" } },
        { content: "Priority", styles: { halign: "center", fontStyle: "bold" } },
        { content: "Dept", styles: { halign: "center", fontStyle: "bold" } },
        { content: "Hold Reason", styles: { halign: "center", fontStyle: "bold" } },
        { content: "Location", styles: { halign: "center", fontStyle: "bold" } },
        { content: "Hold By", styles: { halign: "center", fontStyle: "bold" } },
        { content: "Approved By", styles: { halign: "center", fontStyle: "bold" } },
        { content: "Hold Date", styles: { halign: "center", fontStyle: "bold" } },
        { content: "Released At", styles: { halign: "center", fontStyle: "bold" } },
      ],
    ];

    const tableBody = filteredData.map((item, idx) => {
      const rowBgColor = idx % 2 === 0 ? [255, 255, 255] : [254, 252, 246];
      const isHold = item.status === "ON HOLD";

      return [
        { content: (idx + 1).toString(), styles: { halign: "center", fillColor: rowBgColor, fontSize: 8 } },
        { content: item.lotNumber || "—", styles: { halign: "center", fontStyle: "bold", textColor: [185, 28, 28], fillColor: rowBgColor, fontSize: 9 } },
        { content: item.garmentType || "—", styles: { halign: "center", fillColor: rowBgColor, fontSize: 8 } },
        { content: item.style || "—", styles: { halign: "center", fillColor: rowBgColor, fontSize: 8 } },
        { content: item.fabric || "—", styles: { halign: "center", fillColor: rowBgColor, fontSize: 8 } },
        { content: item.brand || "—", styles: { halign: "center", fontStyle: "bold", fillColor: rowBgColor, fontSize: 8 } },
        { content: `${item.quantity || 0} ${item.unit || ''}`, styles: { halign: "center", fontStyle: "bold", textColor: [180, 83, 9], fillColor: rowBgColor, fontSize: 8.5 } },
        { content: item.section || "—", styles: { halign: "center", fillColor: rowBgColor, fontSize: 8 } },
        { content: item.season || "—", styles: { halign: "center", fillColor: rowBgColor, fontSize: 8 } },
        { content: item.partyName || "—", styles: { halign: "center", fillColor: rowBgColor, fontSize: 8 } },
        { content: item.directStitching || "—", styles: { halign: "center", fillColor: rowBgColor, fontSize: 8 } },
        {
          content: item.status,
          styles: {
            halign: "center",
            fontStyle: "bold",
            fillColor: isHold ? [254, 226, 226] : [220, 252, 231],
            textColor: isHold ? [185, 28, 28] : [21, 128, 61],
            fontSize: 8,
          },
        },
        { content: item.priority || "—", styles: { halign: "center", fontStyle: "bold", fillColor: rowBgColor, fontSize: 7.5 } },
        { content: item.department || "—", styles: { halign: "center", fillColor: rowBgColor, fontSize: 8 } },
        { content: item.reason || "—", styles: { halign: "left", fillColor: rowBgColor, fontSize: 8 } },
        { content: item.location || "—", styles: { halign: "center", fontStyle: "bold", fillColor: rowBgColor, textColor: [30, 41, 59], fontSize: 8 } },
        { content: item.holdBy || "—", styles: { halign: "center", fillColor: rowBgColor, fontSize: 7.5 } },
        { content: item.approvedBy || "—", styles: { halign: "center", fillColor: rowBgColor, fontSize: 7.5 } },
        { content: item.timestamp ? formatDate(item.timestamp) : "—", styles: { halign: "center", fillColor: rowBgColor, fontSize: 7.5 } },
        { content: item.releasedAt ? formatDate(item.releasedAt) : "—", styles: { halign: "center", fillColor: rowBgColor, fontSize: 7.5 } },
      ];
    });

    autoTable(doc, {
      head: tableHeaders,
      body: tableBody,
      startY: 28,
      theme: "grid",
      headStyles: {
        fillColor: headerColor,
        textColor: [255, 255, 255],
        fontSize: 8.5,
        fontStyle: "bold",
        halign: "center",
        valign: "middle",
        cellPadding: { top: 3.5, right: 1.5, bottom: 3.5, left: 1.5 },
      },
      bodyStyles: {
        textColor: textColor,
        fontSize: 8,
        halign: "center",
        valign: "middle",
        cellPadding: { top: 2.8, right: 1.5, bottom: 2.8, left: 1.5 },
        lineColor: borderColor,
        overflow: "linebreak",
      },
      margin: { top: 28, left: margin, right: margin, bottom: 10 },
      tableWidth: contentWidth,
      showHead: "everyPage",
      didDrawPage: function () {
        drawHeader();
        const str = `Page ${doc.internal.getNumberOfPages()}`;
        doc.setFontSize(8);
        doc.setTextColor(148, 163, 184);
        doc.text(str, pageWidth - margin - 15, pageHeight - 4);
        doc.text("Factory Suite Pro — Hold Lots Management", margin, pageHeight - 4);
      },
    });

    const fileDate = new Date().toISOString().slice(0, 10);
    doc.save(`Hold_Lots_Report_${fileDate}.pdf`);
  };

  return (
    <>
      <style>{`
        .hold-report-container {
          padding: 28px 32px 80px;
          background-color: #f8fafc;
          min-height: 100vh;
          font-family: 'Plus Jakarta Sans', 'Inter', system-ui, -apple-system, sans-serif;
          color: #0f172a;
          box-sizing: border-box;
        }
        .report-header-box {
          background: linear-gradient(135deg, #7c2d12 0%, #b45309 50%, #dc2626 100%);
          border-radius: 20px;
          padding: 24px 32px;
          color: white;
          margin-bottom: 24px;
          box-shadow: 0 12px 28px -6px rgba(180, 83, 9, 0.35);
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
        .header-btn-action {
          background: #ffffff !important;
          color: #7c2d12 !important;
          border-color: #ffffff !important;
          font-weight: 800 !important;
          box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
        }
        .header-btn-action:hover {
          background: #fef3c7 !important;
        }
        .header-btn-pdf {
          background: #dc2626 !important;
          border-color: #b91c1c !important;
          box-shadow: 0 4px 12px rgba(220, 38, 38, 0.35);
        }
        .header-btn-excel {
          background: #059669 !important;
          border-color: #047857 !important;
          box-shadow: 0 4px 12px rgba(5, 150, 105, 0.35);
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
          position: relative;
          overflow: hidden;
        }
        .stat-card::before {
          content: '';
          position: absolute;
          top: 0;
          left: 0;
          width: 5px;
          height: 100%;
        }
        .stat-card.amber::before { background: #f59e0b; }
        .stat-card.red::before { background: #ef4444; }
        .stat-card.green::before { background: #10b981; }
        .stat-card.blue::before { background: #3b82f6; }
        .stat-card.purple::before { background: #8b5cf6; }

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
          grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
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
          border-color: #b45309;
        }
        .status-segment-group {
          display: flex;
          background: #f1f5f9;
          padding: 3px;
          border-radius: 10px;
          gap: 4px;
        }
        .status-segment-btn {
          flex: 1;
          padding: 6px 12px;
          border-radius: 8px;
          border: none;
          background: transparent;
          font-size: 0.8rem;
          font-weight: 700;
          color: #64748b;
          cursor: pointer;
          transition: all 0.2s;
          white-space: nowrap;
        }
        .status-segment-btn.active-all {
          background: #ffffff;
          color: #0f172a;
          box-shadow: 0 2px 4px rgba(0, 0, 0, 0.08);
        }
        .status-segment-btn.active-hold {
          background: #fee2e2;
          color: #b91c1c;
          box-shadow: 0 2px 4px rgba(239, 68, 68, 0.2);
        }
        .status-segment-btn.active-rel {
          background: #dcfce7;
          color: #15803d;
          box-shadow: 0 2px 4px rgba(34, 197, 94, 0.2);
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
          min-width: 1300px;
        }
        .data-table th {
          background: #7c2d12;
          color: white;
          padding: 14px 16px;
          text-align: center;
          font-weight: 800;
          font-size: 11.5px;
          text-transform: uppercase;
          letter-spacing: 0.6px;
          position: sticky;
          top: 0;
          z-index: 10;
          white-space: nowrap;
          border-bottom: 2px solid rgba(255, 255, 255, 0.2);
        }
        .data-table td {
          padding: 12px 14px;
          border-bottom: 1px solid #f1f5f9;
          font-size: 13px;
          color: #1e293b;
          text-align: center;
          vertical-align: middle;
        }
        .data-table tr:hover td {
          background-color: #fffbeb !important;
        }
        .status-badge-hold {
          background: #fee2e2;
          color: #b91c1c;
          border: 1px solid #fca5a5;
          padding: 3px 10px;
          border-radius: 6px;
          font-weight: 800;
          font-size: 0.75rem;
          display: inline-flex;
          align-items: center;
          gap: 4px;
        }
        .status-badge-released {
          background: #dcfce7;
          color: #15803d;
          border: 1px solid #86efac;
          padding: 3px 10px;
          border-radius: 6px;
          font-weight: 800;
          font-size: 0.75rem;
          display: inline-flex;
          align-items: center;
          gap: 4px;
        }
        .priority-badge-high {
          background: #fef2f2;
          color: #dc2626;
          border: 1px solid #fecaca;
          padding: 2px 8px;
          border-radius: 6px;
          font-size: 0.72rem;
          font-weight: 800;
        }
        .priority-badge-std {
          background: #f8fafc;
          color: #475569;
          border: 1px solid #e2e8f0;
          padding: 2px 8px;
          border-radius: 6px;
          font-size: 0.72rem;
          font-weight: 700;
        }
        .location-pill {
          background: #fef3c7;
          color: #92400e;
          border: 1px solid #fde68a;
          padding: 3px 8px;
          border-radius: 6px;
          font-size: 0.78rem;
          font-weight: 800;
          display: inline-flex;
          align-items: center;
          gap: 4px;
        }
        .active-tag-pill {
          display: inline-flex;
          align-items: center;
          gap: 6px;
          background: rgba(180, 83, 9, 0.1);
          color: #b45309;
          border: 1px solid #fde68a;
          padding: 3px 10px;
          border-radius: 8px;
          font-size: 0.75rem;
          font-weight: 700;
        }
      `}</style>

      <div className="hold-report-container">
        {/* Header Box */}
        <div className="report-header-box">
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
              <span style={{ fontSize: "1.8rem" }}>🚨</span>
              <h1 style={{ margin: 0, fontSize: "1.45rem", fontWeight: 800, color: "#ffffff" }}>
                Hold Lots Master Audit Report
              </h1>
              <span style={{ background: "#fbbf24", color: "#7c2d12", padding: "2px 8px", borderRadius: "8px", fontSize: "0.72rem", fontWeight: 800 }}>
                LIVE HOLD LOGS
              </span>
            </div>
            <p style={{ margin: "6px 0 0 0", color: "#fef3c7", fontSize: "0.85rem" }}>
              Comprehensive tracking of all lots put on hold across all factory departments with location and audit approval details
            </p>
          </div>

          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", alignItems: "center" }}>
            <button className="header-btn" onClick={() => history.push("/dashboard")}>
              ← Dashboard
            </button>
            <button className="header-btn header-btn-action" onClick={() => history.push("/hold-lot")}>
              ⚡ Put Lot on Hold / Action
            </button>
            <button className="header-btn" onClick={fetchData}>
              ↻ Refresh
            </button>
            <button className="header-btn header-btn-excel" onClick={exportToExcel} disabled={filteredData.length === 0}>
              📊 Excel Export
            </button>
            <button className="header-btn header-btn-pdf" onClick={exportToPDF} disabled={filteredData.length === 0}>
              📄 PDF Export
            </button>
            <button className="header-btn" onClick={exportToCSV}>
              📥 CSV
            </button>
          </div>
        </div>

        {/* Stats Grid */}
        <div className="stat-cards-grid">
          <div className="stat-card amber">
            <div className="stat-label">Total Hold Records</div>
            <div className="stat-value" style={{ color: "#b45309" }}>{totalHoldRecords}</div>
          </div>
          <div className="stat-card red">
            <div className="stat-label">Active On Hold</div>
            <div className="stat-value" style={{ color: "#dc2626" }}>{activeHolds}</div>
          </div>
          <div className="stat-card green">
            <div className="stat-label">Released Lots</div>
            <div className="stat-value" style={{ color: "#16a34a" }}>{releasedHolds}</div>
          </div>
          <div className="stat-card blue">
            <div className="stat-label">Total Hold Pieces</div>
            <div className="stat-value" style={{ color: "#2563eb" }}>{totalHoldQty.toLocaleString()}</div>
          </div>
          <div className="stat-card purple">
            <div className="stat-label">High Alert Holds</div>
            <div className="stat-value" style={{ color: "#9333ea" }}>{highAlertHolds}</div>
          </div>
        </div>

        {/* Filters Toolbar */}
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
              <label className="filter-label">Quick Search</label>
              <input
                type="text"
                placeholder="Search Lot, Style, Location..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="filter-input"
              />
            </div>

            {/* Status Segment */}
            <div>
              <label className="filter-label">Status</label>
              <div className="status-segment-group">
                <button
                  type="button"
                  onClick={() => setStatusFilter("ALL")}
                  className={`status-segment-btn ${statusFilter === "ALL" ? "active-all" : ""}`}
                >
                  All ({data.length})
                </button>
                <button
                  type="button"
                  onClick={() => setStatusFilter("ON HOLD")}
                  className={`status-segment-btn ${statusFilter === "ON HOLD" ? "active-hold" : ""}`}
                >
                  On Hold
                </button>
                <button
                  type="button"
                  onClick={() => setStatusFilter("RELEASED")}
                  className={`status-segment-btn ${statusFilter === "RELEASED" ? "active-rel" : ""}`}
                >
                  Released
                </button>
              </div>
            </div>

            {/* Department Multi-Select */}
            <MultiSelectDropdown
              label="Department"
              options={departmentOptions}
              selectedValues={selectedDepts}
              onChange={setSelectedDepts}
            />

            {/* Priority Multi-Select */}
            <MultiSelectDropdown
              label="Priority"
              options={priorityOptions}
              selectedValues={selectedPriorities}
              onChange={setSelectedPriorities}
            />

            {/* Location Multi-Select */}
            <MultiSelectDropdown
              label="Location"
              options={locationOptions}
              selectedValues={selectedLocations}
              onChange={setSelectedLocations}
            />

            {/* Brand Multi-Select */}
            <MultiSelectDropdown
              label="Brand"
              options={brandOptions}
              selectedValues={selectedBrands}
              onChange={setSelectedBrands}
            />

            {/* Style Multi-Select */}
            <MultiSelectDropdown
              label="Style"
              options={styleOptions}
              selectedValues={selectedStyles}
              onChange={setSelectedStyles}
            />

            {/* Fabric Multi-Select */}
            <MultiSelectDropdown
              label="Fabric"
              options={fabricOptions}
              selectedValues={selectedFabrics}
              onChange={setSelectedFabrics}
            />

            {/* Party Name Multi-Select */}
            <MultiSelectDropdown
              label="Party Name"
              options={partyOptions}
              selectedValues={selectedParties}
              onChange={setSelectedParties}
            />

            {/* Hold Reason Multi-Select */}
            <MultiSelectDropdown
              label="Hold Reason"
              options={reasonOptions}
              selectedValues={selectedReasons}
              onChange={setSelectedReasons}
            />

            {/* Hold By Multi-Select */}
            <MultiSelectDropdown
              label="Hold By"
              options={holdByOptions}
              selectedValues={selectedHoldBy}
              onChange={setSelectedHoldBy}
            />

            {/* Approved By Multi-Select */}
            <MultiSelectDropdown
              label="Approved By"
              options={approvedByOptions}
              selectedValues={selectedApprovedBy}
              onChange={setSelectedApprovedBy}
            />
          </div>

          {/* Active Filter Pills */}
          {hasActiveFilters && (
            <div style={{ display: "flex", gap: "8px", flexWrap: "wrap", marginTop: "14px", paddingTop: "12px", borderTop: "1px solid #f1f5f9", alignItems: "center" }}>
              <span style={{ fontSize: "0.75rem", fontWeight: 800, color: "#64748b", textTransform: "uppercase" }}>Active:</span>
              {statusFilter !== "ALL" && (
                <span className="active-tag-pill">
                  Status: {statusFilter}
                  <span onClick={() => setStatusFilter("ALL")} style={{ cursor: "pointer", fontWeight: 900 }}>×</span>
                </span>
              )}
              {selectedDepts.map((d) => (
                <span key={d} className="active-tag-pill">
                  Dept: {d}
                  <span onClick={() => setSelectedDepts(selectedDepts.filter((x) => x !== d))} style={{ cursor: "pointer", fontWeight: 900 }}>×</span>
                </span>
              ))}
              {selectedLocations.map((l) => (
                <span key={l} className="active-tag-pill">
                  📍 {l}
                  <span onClick={() => setSelectedLocations(selectedLocations.filter((x) => x !== l))} style={{ cursor: "pointer", fontWeight: 900 }}>×</span>
                </span>
              ))}
              {selectedBrands.map((b) => (
                <span key={b} className="active-tag-pill">
                  Brand: {b}
                  <span onClick={() => setSelectedBrands(selectedBrands.filter((x) => x !== b))} style={{ cursor: "pointer", fontWeight: 900 }}>×</span>
                </span>
              ))}
            </div>
          )}
        </div>

        {/* Table Card */}
        <div className="table-card">
          <div className="table-header">
            <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
              <span style={{ fontWeight: 800, fontSize: "1.05rem" }}>Hold Lots Registry</span>
              <span style={{ background: "#fef3c7", padding: "3px 10px", borderRadius: "6px", fontSize: "0.8rem", fontWeight: 800, color: "#92400e" }}>
                Showing {filteredData.length} of {data.length} records ({totalHoldQty.toLocaleString()} pcs)
              </span>
            </div>
          </div>

          <div className="table-container">
            {loading ? (
              <div style={{ textAlign: "center", padding: "60px" }}>
                <div style={{ fontSize: "2rem", marginBottom: "8px" }}>⚡</div>
                <div style={{ fontWeight: 700, color: "#64748b" }}>Loading Hold Lots records from Google Sheet...</div>
              </div>
            ) : filteredData.length === 0 ? (
              <div style={{ textAlign: "center", padding: "60px", color: "#64748b" }}>
                <div style={{ fontSize: "2.5rem", marginBottom: "8px" }}>📭</div>
                <div style={{ fontWeight: 800, fontSize: "1.1rem", color: "#0f172a" }}>No Hold Lots Found</div>
                <div style={{ fontSize: "0.85rem", marginTop: "4px" }}>
                  {hasActiveFilters ? "Try adjusting or clearing your filters." : "New hold lots added via the Hold Action Center will appear here."}
                </div>
              </div>
            ) : (
              <table className="data-table">
                <thead>
                  <tr>
                    <th>Sr</th>
                    <th>Lot Number</th>
                    <th>Garment Type</th>
                    <th>Style</th>
                    <th>Fabric</th>
                    <th>Brand</th>
                    <th>Qty</th>
                    <th>Section</th>
                    <th>Season</th>
                    <th>Party Name</th>
                    <th>Direct Stitching</th>
                    <th>Status</th>
                    <th>Priority</th>
                    <th>Department</th>
                    <th>Location</th>
                    <th>Hold Reason</th>
                    <th>Job Order</th>
                    <th>Hold By</th>
                    <th>Approved By</th>
                    <th>Hold Date</th>
                    <th>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredData.map((item, idx) => {
                    const isHold = item.status === "ON HOLD";
                    const isHigh = String(item.priority).toLowerCase().includes("high");

                    return (
                      <tr key={item.id || idx}>
                        <td style={{ color: "#64748b", fontWeight: 600 }}>{idx + 1}</td>
                        <td>
                          <strong style={{ color: "#dc2626", fontSize: "0.95rem" }}>
                            #{item.lotNumber}
                          </strong>
                        </td>
                        <td>{item.garmentType || "—"}</td>
                        <td>{item.style || "—"}</td>
                        <td>{item.fabric || "—"}</td>
                        <td><strong>{item.brand || "—"}</strong></td>
                        <td>
                          <span style={{ background: "#fef3c7", color: "#92400e", padding: "2px 6px", borderRadius: "6px", fontWeight: 800, fontSize: "0.82rem" }}>
                            {item.quantity} {item.unit}
                          </span>
                        </td>
                        <td>{item.section || "—"}</td>
                        <td>{item.season || "—"}</td>
                        <td>{item.partyName || "—"}</td>
                        <td>{item.directStitching || "—"}</td>
                        <td>
                          <span className={isHold ? "status-badge-hold" : "status-badge-released"}>
                            {isHold ? "⏸️ ON HOLD" : "✓ RELEASED"}
                          </span>
                        </td>
                        <td>
                          <span className={isHigh ? "priority-badge-high" : "priority-badge-std"}>
                            {item.priority}
                          </span>
                        </td>
                        <td>
                          <span style={{ background: "#f1f5f9", padding: "2px 8px", borderRadius: "6px", fontSize: "0.8rem", fontWeight: 700, color: "#334155" }}>
                            {item.department}
                          </span>
                        </td>
                        <td>
                          {item.location && item.location !== "—" ? (
                            <span className="location-pill">
                              📍 {item.location}
                            </span>
                          ) : (
                            <span style={{ color: "#94a3b8" }}>—</span>
                          )}
                        </td>
                        <td style={{ textAlign: "left", maxWidth: "220px", fontSize: "0.82rem", fontWeight: 600, color: "#991b1b" }}>
                          {item.reason}
                        </td>
                        <td>{item.jobOrderNo || "—"}</td>
                        <td style={{ fontSize: "0.82rem" }}>{item.holdBy || "—"}</td>
                        <td style={{ fontSize: "0.82rem", fontWeight: 600 }}>{item.approvedBy || "—"}</td>
                        <td style={{ fontSize: "0.8rem", color: "#64748b" }}>{item.timestamp ? formatDate(item.timestamp) : "—"}</td>
                        <td>
                          <button
                            onClick={() => setSelectedLotModal(item)}
                            style={{
                              background: "#f1f5f9",
                              border: "1px solid #cbd5e1",
                              padding: "4px 10px",
                              borderRadius: "6px",
                              fontSize: "0.78rem",
                              fontWeight: 700,
                              cursor: "pointer",
                              color: "#334155"
                            }}
                          >
                            👁️ View
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            )}
          </div>
        </div>
      </div>

      {/* Lot Details Modal */}
      {selectedLotModal && (
        <div
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: "rgba(15, 23, 42, 0.75)",
            zIndex: 99999,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            backdropFilter: "blur(4px)",
            padding: "20px"
          }}
          onClick={() => setSelectedLotModal(null)}
        >
          <div
            style={{
              background: "#ffffff",
              borderRadius: "20px",
              maxWidth: "600px",
              width: "100%",
              padding: "24px",
              boxShadow: "0 25px 50px -12px rgba(0, 0, 0, 0.35)",
              border: "1px solid #cbd5e1"
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "16px", borderBottom: "1px solid #f1f5f9", paddingBottom: "12px" }}>
              <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                <span style={{ fontSize: "1.5rem" }}>🚨</span>
                <div>
                  <h3 style={{ margin: 0, fontSize: "1.15rem", fontWeight: 800, color: "#7c2d12" }}>
                    Lot #{selectedLotModal.lotNumber} Details
                  </h3>
                  <span style={{ fontSize: "0.78rem", color: "#64748b" }}>Hold ID: {selectedLotModal.id}</span>
                </div>
              </div>
              <button
                onClick={() => setSelectedLotModal(null)}
                style={{
                  background: "#f1f5f9",
                  border: "none",
                  borderRadius: "50%",
                  width: "30px",
                  height: "30px",
                  cursor: "pointer",
                  fontWeight: "bold",
                  color: "#64748b"
                }}
              >
                ✕
              </button>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: "12px", fontSize: "0.85rem", marginBottom: "20px" }}>
              <div><strong>Status:</strong> <span style={{ color: selectedLotModal.status === "ON HOLD" ? "#dc2626" : "#16a34a", fontWeight: 800 }}>{selectedLotModal.status}</span></div>
              <div><strong>Priority:</strong> {selectedLotModal.priority}</div>
              <div><strong>Department:</strong> {selectedLotModal.department}</div>
              <div><strong>Location:</strong> 📍 {selectedLotModal.location || "Not Specified"}</div>
              <div><strong>Brand:</strong> {selectedLotModal.brand}</div>
              <div><strong>Style:</strong> {selectedLotModal.style}</div>
              <div><strong>Fabric:</strong> {selectedLotModal.fabric}</div>
              <div><strong>Quantity:</strong> {selectedLotModal.quantity} {selectedLotModal.unit}</div>
              <div><strong>Party:</strong> {selectedLotModal.partyName}</div>
              <div><strong>Job Order:</strong> {selectedLotModal.jobOrderNo}</div>
              <div><strong>Hold By:</strong> {selectedLotModal.holdBy}</div>
              <div><strong>Approved By:</strong> {selectedLotModal.approvedBy}</div>
              <div><strong>Hold Date:</strong> {selectedLotModal.timestamp ? formatDate(selectedLotModal.timestamp) : "—"}</div>
              <div><strong>Released:</strong> {selectedLotModal.releasedAt ? `${formatDate(selectedLotModal.releasedAt)} (${selectedLotModal.releasedBy || ''})` : "Not Released"}</div>
            </div>

            <div style={{ background: "#fef2f2", border: "1px solid #fecaca", borderRadius: "10px", padding: "12px", marginBottom: "16px" }}>
              <div style={{ fontSize: "0.75rem", fontWeight: 800, color: "#991b1b", textTransform: "uppercase", marginBottom: "4px" }}>Hold Reason / Issue Remark:</div>
              <div style={{ fontSize: "0.9rem", color: "#7f1d1d", fontWeight: 600 }}>{selectedLotModal.reason}</div>
            </div>

            <div style={{ display: "flex", justifyContent: "flex-end", gap: "10px" }}>
              <button
                onClick={() => setSelectedLotModal(null)}
                style={{
                  background: "#f1f5f9",
                  border: "1px solid #cbd5e1",
                  padding: "8px 16px",
                  borderRadius: "8px",
                  fontWeight: 700,
                  cursor: "pointer"
                }}
              >
                Close
              </button>
              <button
                onClick={() => {
                  setSelectedLotModal(null);
                  history.push("/hold-lot");
                }}
                style={{
                  background: "#7c2d12",
                  color: "white",
                  border: "none",
                  padding: "8px 16px",
                  borderRadius: "8px",
                  fontWeight: 700,
                  cursor: "pointer"
                }}
              >
                Open in Hold Action Center →
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}
