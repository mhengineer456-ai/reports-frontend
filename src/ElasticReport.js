import React, { useState, useEffect, useMemo, useRef } from "react";
import { useHistory } from "react-router-dom";
import { SPREADSHEET_IDS, fetchSheetDataFromBackend } from "./config";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

const SPREADSHEET_ID = "1IMhmYlJ3s2PPRgEQs1Ikd4O1OBXK4EYL1oV_-kWAkyg";
const SHEET_NAME = "Elastic";
const INDEX_SPREADSHEET_ID = "1Hj3JeJEKB43aYYWv8gk2UhdU6BWuEQfCg5pBlTdBMNA";

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

function normalizeLot(lot) {
  return String(lot || "").trim().toLowerCase();
}

// Reusable Multi-Select Dropdown Component
function MultiSelectDropdown({ label, options, selectedValues, onChange, themeColor = "#6366f1" }) {
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
    <div ref={dropdownRef} style={{ position: "relative", minWidth: "140px" }}>
      <label
        style={{
          fontSize: "0.72rem",
          fontWeight: 800,
          color: "#64748b",
          textTransform: "uppercase",
          letterSpacing: "0.4px",
          marginBottom: "4px",
          display: "block",
        }}
      >
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
          background: selectedValues.length > 0 ? "rgba(99, 102, 241, 0.08)" : "#ffffff",
          fontWeight: selectedValues.length > 0 ? "800" : "600",
          color: selectedValues.length > 0 ? themeColor : "#1e293b",
          cursor: "pointer",
          fontSize: "0.85rem",
          transition: "all 0.18s",
          boxSizing: "border-box",
        }}
      >
        <span style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: "130px" }}>
          {getDisplayText()}
        </span>
        <span
          style={{
            fontSize: "0.75rem",
            transform: isOpen ? "rotate(180deg)" : "rotate(0deg)",
            transition: "transform 0.2s",
            opacity: 0.7,
            marginLeft: "6px",
          }}
        >
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
            minWidth: "220px",
            maxHeight: "260px",
            overflowY: "auto",
            padding: "8px",
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
                outline: "none",
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
              background: isAllSelected ? "rgba(99, 102, 241, 0.08)" : "transparent",
              borderBottom: "1px solid #f1f5f9",
              marginBottom: "4px",
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
                  background: isChecked ? "rgba(99, 102, 241, 0.06)" : "transparent",
                  color: isChecked ? themeColor : "#334155",
                  fontWeight: isChecked ? "700" : "500",
                  transition: "background 0.15s",
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

export default function ElasticReport() {
  const history = useHistory();
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  // Filter States - Default to pending lots
  const [searchTerm, setSearchTerm] = useState("");
  const [selectedStatuses, setSelectedStatuses] = useState(["pending"]);
  const [selectedBrands, setSelectedBrands] = useState([]);
  const [selectedGarments, setSelectedGarments] = useState([]);
  const [selectedFabrics, setSelectedFabrics] = useState([]);
  const [selectedStyles, setSelectedStyles] = useState([]);
  const [selectedSeasons, setSelectedSeasons] = useState([]);
  const [selectedSections, setSelectedSections] = useState([]);
  const [selectedSupervisors, setSelectedSupervisors] = useState([]);
  const [selectedStitchingSupervisors, setSelectedStitchingSupervisors] = useState([]);
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");

  // Pagination & Sorting
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(25);
  const [sortField, setSortField] = useState("id");
  const [sortAsc, setSortAsc] = useState(false);
  const [copiedLot, setCopiedLot] = useState(null);

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
        // fallback to raw
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
        // fallback to raw
      }
    }
    return s;
  };

  const fetchData = async () => {
    setLoading(true);
    setError(null);
    try {
      // Fetch Elastic Sheet, Index Sheet, and JobOrder Sheet data in parallel
      const [elasticRes, indexRes, jobOrderRes] = await Promise.allSettled([
        fetchSheetDataFromBackend(SPREADSHEET_ID, `'${SHEET_NAME}'!A:Z`),
        fetchSheetDataFromBackend(INDEX_SPREADSHEET_ID, `'Index'!A:AG`),
        fetchSheetDataFromBackend(SPREADSHEET_IDS.JOBORDER, `'JobOrder'!A:AZ`),
      ]);

      // Parse JobOrder Sheet for Brand / Party / Season / Section lookup against lots
      const lotToJobInfo = {};
      if (jobOrderRes.status === "fulfilled" && jobOrderRes.value?.ok && Array.isArray(jobOrderRes.value.values)) {
        const jRows = jobOrderRes.value.values;
        const jHeaders = jRows[0] || [];
        const jLotIdx = findCol(jHeaders, ["lot number", "lot no", "lot #", "lot"]);
        const jBrandIdx = findCol(jHeaders, ["brand", "brand name", "buyer"]);
        const jPartyIdx = findCol(jHeaders, ["party name", "party", "customer"]);
        const jGarmentIdx = findCol(jHeaders, ["garment type", "garment"]);
        const jStyleIdx = findCol(jHeaders, ["style"]);
        const jFabricIdx = findCol(jHeaders, ["fabric"]);
        const jSeasonIdx = findCol(jHeaders, ["season"]);
        const jSectionIdx = findCol(jHeaders, ["section"]);
        const jDirectIdx = findCol(jHeaders, ["direct stitching", "direct", "design work"]);

        if (jLotIdx !== -1) {
          for (let i = 1; i < jRows.length; i++) {
            const r = jRows[i];
            const lotKey = normalizeLot(r[jLotIdx]);
            if (lotKey) {
              lotToJobInfo[lotKey] = {
                brand: (jBrandIdx !== -1 ? String(r[jBrandIdx] || "").trim() : "") || (jPartyIdx !== -1 ? String(r[jPartyIdx] || "").trim() : ""),
                party: jPartyIdx !== -1 ? String(r[jPartyIdx] || "").trim() : "",
                garment: jGarmentIdx !== -1 ? String(r[jGarmentIdx] || "").trim() : "",
                style: jStyleIdx !== -1 ? String(r[jStyleIdx] || "").trim() : "",
                fabric: jFabricIdx !== -1 ? String(r[jFabricIdx] || "").trim() : "",
                season: jSeasonIdx !== -1 ? String(r[jSeasonIdx] || "").trim() : "",
                section: jSectionIdx !== -1 ? String(r[jSectionIdx] || "").trim() : "",
                directStitching: jDirectIdx !== -1 ? String(r[jDirectIdx] || "").trim() : "",
              };
            }
          }
        }
      }

      // Parse Index Sheet for Stitching Supervisor lookup
      const lotToIndexInfo = {};
      if (indexRes.status === "fulfilled" && indexRes.value?.ok && Array.isArray(indexRes.value.values)) {
        const indexRows = indexRes.value.values;
        const indexHeaders = indexRows[0] || [];
        const iLotIdx = findCol(indexHeaders, ["lot number", "lot no", "lot #", "lot"]);
        const iSupIdx = findCol(indexHeaders, ["supervisor", "stitching supervisor", "stiching supervisor"]);

        if (iLotIdx !== -1) {
          for (let i = 1; i < indexRows.length; i++) {
            const row = indexRows[i];
            const lot = normalizeLot(row[iLotIdx]);
            const sup = iSupIdx !== -1 ? String(row[iSupIdx] || "").trim() : "";
            if (lot) {
              lotToIndexInfo[lot] = { supervisor: sup };
            }
          }
        }
      }

      if (
        elasticRes.status !== "fulfilled" ||
        !elasticRes.value?.ok ||
        !Array.isArray(elasticRes.value.values) ||
        elasticRes.value.values.length === 0
      ) {
        setData([]);
        setLoading(false);
        return;
      }

      const rows = elasticRes.value.values;
      const headers = rows[0] || [];

      const tsIdx = findCol(headers, ["timestamp", "time"]);
      const lotIdx = findCol(headers, ["lot number", "lot no", "lot"]);
      const garmentIdx = findCol(headers, ["garment type", "garment"]);
      const fabricIdx = findCol(headers, ["fabric"]);
      const styleIdx = findCol(headers, ["style"]);
      const brandIdx = findCol(headers, ["brand", "brand name"]);
      const seasonIdx = findCol(headers, ["season"]);
      const sectionIdx = findCol(headers, ["section"]);
      const supIdx = findCol(headers, ["elastic supervisor", "supervisor"]);
      const stitchSupIdx = findCol(headers, ["stitching supervisor", "stiching supervisor", "stitching sup"]);
      const dateIdx = findCol(headers, ["elastic date", "date", "issue date"]);
      const pcsIdx = findCol(headers, ["total pcs", "pcs", "quantity"]);
      const wipIdx = findCol(headers, ["wip elastic", "wip", "remarks"]);
      const compIdx = findCol(headers, ["elastic complete", "elastic completed", "complete", "completed"]);
      const manpowerIdx = findCol(headers, ["total manpower", "manpower"]);

      const parsed = [];
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const lotNo = String(row[lotIdx !== -1 ? lotIdx : 1] || "").trim();
        if (!lotNo) continue;

        const pcs = parseInt(row[pcsIdx !== -1 ? pcsIdx : 8], 10) || 0;
        const rawComp = String(row[compIdx !== -1 ? compIdx : 10] || "").trim();
        const rawWip = String(row[wipIdx !== -1 ? wipIdx : 9] || "").trim();
        const manpower = parseInt(row[manpowerIdx !== -1 ? manpowerIdx : 11], 10) || 0;

        const completeDate = parseCompletionDate(rawComp);
        const isCompleted = !!completeDate && completeDate !== "-" && !completeDate.toLowerCase().includes("pending");
        const wipText = isCompleted ? "Done" : parseWipRemarks(rawWip);

        // Fetch Stitching Supervisor & JobOrder details (Brand, Season, Section, Party, Direct)
        const rowStitchSup = stitchSupIdx !== -1 ? String(row[stitchSupIdx] || "").trim() : "";
        const normLot = normalizeLot(lotNo);
        const indexInfo = lotToIndexInfo[normLot] || {};
        const jobInfo = lotToJobInfo[normLot] || {};
        const stitchingSupervisor = rowStitchSup || indexInfo.supervisor || "—";
        const sheetBrand = brandIdx !== -1 ? String(row[brandIdx] || "").trim() : "";
        const finalBrand = sheetBrand && sheetBrand !== "-" ? sheetBrand : (jobInfo.brand || jobInfo.party || "—");
        const sheetSeason = seasonIdx !== -1 ? String(row[seasonIdx] || "").trim() : "";
        const finalSeason = sheetSeason && sheetSeason !== "-" ? sheetSeason : (jobInfo.season || "—");
        const sheetSection = sectionIdx !== -1 ? String(row[sectionIdx] || "").trim() : "";
        const finalSection = sheetSection && sheetSection !== "-" ? sheetSection : (jobInfo.section || "—");

        parsed.push({
          id: i,
          timestamp: row[tsIdx !== -1 ? tsIdx : 0] || "",
          lotNumber: lotNo,
          garmentType: row[garmentIdx !== -1 ? garmentIdx : 2] || jobInfo.garment || "",
          style: row[styleIdx !== -1 ? styleIdx : 4] || jobInfo.style || "",
          fabric: row[fabricIdx !== -1 ? fabricIdx : 3] || jobInfo.fabric || "",
          brand: finalBrand,
          totalPcs: pcs,
          section: finalSection,
          season: finalSeason,
          partyName: jobInfo.party || "—",
          directStitching: jobInfo.directStitching || "—",
          stitchingSupervisor: stitchingSupervisor,
          elasticDate: row[dateIdx !== -1 ? dateIdx : 7] || "",
          wip: wipText,
          supervisor: row[supIdx !== -1 ? supIdx : 6] || "Elastic Department",
          completeDate,
          isCompleted,
          manpower,
        });
      }

      setData(parsed);
    } catch (err) {
      console.error("Error loading Elastic report data:", err);
      setError(err.message || "Failed to load sheet data");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchData();
  }, []);

  // Multi-Select Filter Options
  const statusOptions = [
    { value: "pending", label: "⏳ In Progress / Pending" },
    { value: "completed", label: "✓ Completed" },
  ];

  const brandOptions = useMemo(() => {
    const set = new Set(data.map((d) => d.brand).filter((b) => b && b !== "—" && b !== "-"));
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

  const styleOptions = useMemo(() => {
    const set = new Set(data.map((d) => d.style).filter(Boolean));
    return Array.from(set).sort().map((s) => ({ value: s, label: s }));
  }, [data]);

  const seasonOptions = useMemo(() => {
    const set = new Set(data.map((d) => d.season).filter((s) => s && s !== "—" && s !== "-"));
    return Array.from(set).sort().map((s) => ({ value: s, label: s }));
  }, [data]);

  const sectionOptions = useMemo(() => {
    const set = new Set(data.map((d) => d.section).filter((s) => s && s !== "—" && s !== "-"));
    return Array.from(set).sort().map((s) => ({ value: s, label: s }));
  }, [data]);

  const supervisorOptions = useMemo(() => {
    const set = new Set(data.map((d) => d.supervisor).filter(Boolean));
    return Array.from(set).sort().map((s) => ({ value: s, label: s }));
  }, [data]);

  const stitchingSupervisorOptions = useMemo(() => {
    const set = new Set(
      data
        .map((d) => d.stitchingSupervisor)
        .filter((s) => s && s !== "—" && s !== "-")
    );
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
        (item.season && item.season.toLowerCase().includes(term)) ||
        (item.section && item.section.toLowerCase().includes(term)) ||
        item.garmentType.toLowerCase().includes(term) ||
        item.supervisor.toLowerCase().includes(term) ||
        (item.stitchingSupervisor && item.stitchingSupervisor.toLowerCase().includes(term)) ||
        item.wip.toLowerCase().includes(term);

      const matchesStatus =
        selectedStatuses.length === 0 ||
        (selectedStatuses.includes("completed") && item.isCompleted) ||
        (selectedStatuses.includes("pending") && !item.isCompleted);

      const matchesBrand =
        selectedBrands.length === 0 || selectedBrands.some((b) => b.toLowerCase() === item.brand.toLowerCase());
      const matchesGarment =
        selectedGarments.length === 0 ||
        selectedGarments.some((g) => g.toLowerCase() === item.garmentType.toLowerCase());
      const matchesFabric =
        selectedFabrics.length === 0 || selectedFabrics.some((f) => f.toLowerCase() === item.fabric.toLowerCase());
      const matchesStyle =
        selectedStyles.length === 0 || selectedStyles.some((s) => s.toLowerCase() === item.style.toLowerCase());
      const matchesSeason =
        selectedSeasons.length === 0 || selectedSeasons.some((s) => s.toLowerCase() === (item.season || "").toLowerCase());
      const matchesSection =
        selectedSections.length === 0 || selectedSections.some((s) => s.toLowerCase() === (item.section || "").toLowerCase());
      const matchesSupervisor =
        selectedSupervisors.length === 0 ||
        selectedSupervisors.some((s) => s.toLowerCase() === item.supervisor.toLowerCase());
      const matchesStitchingSupervisor =
        selectedStitchingSupervisors.length === 0 ||
        selectedStitchingSupervisors.some((s) => s.toLowerCase() === (item.stitchingSupervisor || "").toLowerCase());

      let matchesDate = true;
      if (startDate && item.elasticDate) {
        matchesDate = matchesDate && item.elasticDate >= startDate;
      }
      if (endDate && item.elasticDate) {
        matchesDate = matchesDate && item.elasticDate <= endDate;
      }

      return (
        matchesSearch &&
        matchesStatus &&
        matchesBrand &&
        matchesGarment &&
        matchesFabric &&
        matchesStyle &&
        matchesSeason &&
        matchesSection &&
        matchesSupervisor &&
        matchesStitchingSupervisor &&
        matchesDate
      );
    });
  }, [
    data,
    searchTerm,
    selectedStatuses,
    selectedBrands,
    selectedGarments,
    selectedFabrics,
    selectedStyles,
    selectedSeasons,
    selectedSections,
    selectedSupervisors,
    selectedStitchingSupervisors,
    startDate,
    endDate,
  ]);

  // Sorted data
  const sortedData = useMemo(() => {
    return [...filteredData].sort((a, b) => {
      let valA = a[sortField];
      let valB = b[sortField];

      if (typeof valA === "string") valA = valA.toLowerCase();
      if (typeof valB === "string") valB = valB.toLowerCase();

      if (valA < valB) return sortAsc ? -1 : 1;
      if (valA > valB) return sortAsc ? 1 : -1;
      return 0;
    });
  }, [filteredData, sortField, sortAsc]);

  // Paginated data
  const totalPages = Math.max(1, Math.ceil(sortedData.length / pageSize));
  const paginatedData = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return sortedData.slice(start, start + pageSize);
  }, [sortedData, currentPage, pageSize]);

  const handleSort = (field) => {
    if (sortField === field) {
      setSortAsc(!sortAsc);
    } else {
      setSortField(field);
      setSortAsc(true);
    }
  };

  const copyLot = (lot) => {
    navigator.clipboard.writeText(lot);
    setCopiedLot(lot);
    setTimeout(() => setCopiedLot(null), 1800);
  };

  const resetFilters = () => {
    setSearchTerm("");
    setSelectedStatuses(["pending"]);
    setSelectedBrands([]);
    setSelectedGarments([]);
    setSelectedFabrics([]);
    setSelectedStyles([]);
    setSelectedSeasons([]);
    setSelectedSections([]);
    setSelectedSupervisors([]);
    setSelectedStitchingSupervisors([]);
    setStartDate("");
    setEndDate("");
    setCurrentPage(1);
  };

  const hasActiveFilters =
    searchTerm !== "" ||
    (selectedStatuses.length !== 1 || selectedStatuses[0] !== "pending") ||
    selectedBrands.length > 0 ||
    selectedGarments.length > 0 ||
    selectedFabrics.length > 0 ||
    selectedStyles.length > 0 ||
    selectedSeasons.length > 0 ||
    selectedSections.length > 0 ||
    selectedSupervisors.length > 0 ||
    selectedStitchingSupervisors.length > 0 ||
    startDate !== "" ||
    endDate !== "";

  // Summary Metrics
  const totalLots = filteredData.length;
  const completedLots = filteredData.filter((item) => item.isCompleted).length;
  const pendingLots = totalLots - completedLots;
  const pendingPcs = filteredData.filter((item) => !item.isCompleted).reduce((sum, item) => sum + item.totalPcs, 0);
  const totalManpower = filteredData.reduce((sum, item) => sum + (item.manpower || 0), 0);
  const completionRate = totalLots > 0 ? Math.round((completedLots / totalLots) * 100) : 0;

  const formatDisplayDate = (val) => {
    if (!val || val === "-" || val === "—") return "—";
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

  // PDF Export Function - 15 columns spacious & polished layout
  // --- Professional PDF Export (A3 Landscape - Matching Factory Suite Pro Standard) ---
  const handleDownloadPDF = () => {
    if (filteredData.length === 0) {
      alert("No data available to download PDF.");
      return;
    }

    try {
      const doc = new jsPDF({
        orientation: "landscape",
        unit: "pt",
        format: "a3"
      });

      const pageW = doc.internal.pageSize.getWidth();
      const pageH = doc.internal.pageSize.getHeight();
      const totalPieces = filteredData.reduce((sum, item) => sum + (item.totalPcs || 0), 0);
      const totalLots = filteredData.length;

      // Grouping for 4-Column Side-by-Side Executive Summary
      const garmentMap = {};
      const supervisorMap = {};
      const seasonMap = {};
      let goodAgingLots = 0, goodAgingPcs = 0;
      let warnAgingLots = 0, warnAgingPcs = 0;
      let critAgingLots = 0, critAgingPcs = 0;
      let completedLotsCount = 0, completedPcsCount = 0;
      let pendingLotsCount = 0, pendingPcsCount = 0;

      filteredData.forEach(item => {
        const pcs = item.totalPcs || 0;
        const gType = (item.garmentType || 'Unknown').trim();
        const sup = (item.supervisor || 'Unassigned').trim();
        const season = (item.season || 'Other / NA').trim();
        const isComp = item.isCompleted;

        if (!garmentMap[gType]) garmentMap[gType] = { totalLots: 0, totalPcs: 0 };
        garmentMap[gType].totalLots += 1;
        garmentMap[gType].totalPcs += pcs;

        if (!supervisorMap[sup]) supervisorMap[sup] = { totalLots: 0, totalPcs: 0 };
        supervisorMap[sup].totalLots += 1;
        supervisorMap[sup].totalPcs += pcs;

        if (!seasonMap[season]) seasonMap[season] = { totalLots: 0, totalPcs: 0 };
        seasonMap[season].totalLots += 1;
        seasonMap[season].totalPcs += pcs;

        if (isComp) {
          completedLotsCount += 1;
          completedPcsCount += pcs;
        } else {
          pendingLotsCount += 1;
          pendingPcsCount += pcs;
        }
      });

      const sortedGarments = Object.keys(garmentMap).map(name => ({
        name,
        totalLots: garmentMap[name].totalLots,
        totalPcs: garmentMap[name].totalPcs
      })).sort((a, b) => b.totalPcs - a.totalPcs);

      const sortedSupervisors = Object.keys(supervisorMap).map(name => ({
        name,
        totalLots: supervisorMap[name].totalLots,
        totalPcs: supervisorMap[name].totalPcs
      })).sort((a, b) => b.totalPcs - a.totalPcs);

      const sortedSeasons = Object.keys(seasonMap).map(name => ({
        name,
        totalLots: seasonMap[name].totalLots,
        totalPcs: seasonMap[name].totalPcs
      })).sort((a, b) => b.totalPcs - a.totalPcs);

      // 1. Main Header Block
      doc.setFillColor(15, 23, 42); // Dark Navy #0F172A
      doc.rect(15, 12, pageW - 30, 48, 'F');

      doc.setFontSize(14);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(255, 255, 255);
      const isPendingView = selectedStatuses.length === 1 && selectedStatuses[0] === "pending";
      const isCompletedView = selectedStatuses.length === 1 && selectedStatuses[0] === "completed";
      let reportTitle = "FACTORY SUITE PRO - ELASTIC DEPARTMENT REPORT";
      if (isPendingView) reportTitle = "FACTORY SUITE PRO - ELASTIC DEPARTMENT REPORT (PENDING LOTS)";
      else if (isCompletedView) reportTitle = "FACTORY SUITE PRO - ELASTIC DEPARTMENT REPORT (COMPLETED LOTS)";
      doc.text(reportTitle, pageW / 2, 30, { align: 'center' });

      doc.setFontSize(8.5);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(199, 210, 254);
      const subText = `Total Lots: ${totalLots}   |   Total Pieces: ${totalPieces.toLocaleString()}   |   Pending Pieces: ${pendingPcsCount.toLocaleString()}   |   Completed Lots: ${completedLotsCount}   |   Pending Lots: ${pendingLotsCount}   |   Supervisors: ${sortedSupervisors.length}`;
      doc.text(subText, pageW / 2, 48, { align: 'center' });

      // 2. Filter Banner
      doc.setFillColor(241, 245, 249);
      doc.rect(15, 63, pageW - 30, 16, 'F');
      doc.setFontSize(8);
      doc.setFont('helvetica', 'italic');
      doc.setTextColor(0, 0, 0);
      const filterSummary = `Filters: Status: ${selectedStatuses.length ? selectedStatuses.join(', ') : 'All'} | Brands: ${selectedBrands.length ? selectedBrands.join(', ') : 'All'} | Garments: ${selectedGarments.length ? selectedGarments.join(', ') : 'All'} | Fabrics: ${selectedFabrics.length ? selectedFabrics.join(', ') : 'All'} | Seasons: ${selectedSeasons.length ? selectedSeasons.join(', ') : 'All'} | Sections: ${selectedSections.length ? selectedSections.join(', ') : 'All'} | Supervisors: ${selectedSupervisors.length ? selectedSupervisors.join(', ') : 'All'} | Stitching Sups: ${selectedStitchingSupervisors.length ? selectedStitchingSupervisors.join(', ') : 'All'} | Search: ${searchTerm || 'None'}`;
      doc.text(filterSummary, pageW / 2, 74, { align: 'center' });

      // 3. Main Data Table
      const tableColumns = [
        '#',
        'Lot Number',
        'Garment Type',
        'Style',
        'Fabric',
        'Brand',
        'Total Pcs',
        'Section',
        'Season',
        'Party Name',
        'Direct Stitching',
        'Stitching Sup',
        'Elastic Date',
        'WIP Remarks',
        'Elastic Sup',
        'Status',
        'Completed At'
      ];

      const tableBody = filteredData.map((item, idx) => {
        const isComp = item.isCompleted;
        return [
          (idx + 1).toString(),
          item.lotNumber || '—',
          item.garmentType || '—',
          item.style || '—',
          item.fabric || '—',
          item.brand || '—',
          (item.totalPcs || 0).toLocaleString(),
          item.section || '—',
          item.season || '—',
          item.partyName || '—',
          item.directStitching || '—',
          item.stitchingSupervisor || '—',
          formatDisplayDate(item.elasticDate) || '—',
          isComp ? "Done" : (item.wip || '—'),
          item.supervisor || '—',
          isComp ? 'Completed' : 'In Progress',
          isComp ? (formatDisplayDate(item.completeDate) || 'Completed') : '—'
        ];
      });

      // Total Row
      tableBody.push([
        '',
        `TOTAL (${totalLots})`,
        '',
        '',
        '',
        '',
        totalPieces.toLocaleString(),
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        `${sortedSupervisors.length} Sups`,
        `${completedLotsCount} Comp | ${pendingLotsCount} Pend`,
        ''
      ]);

      const columnStyles = {
        0: { cellWidth: 25, halign: 'center' },
        1: { cellWidth: 65, halign: 'center', fontStyle: 'bold' },
        2: { cellWidth: 80, halign: 'center' },
        3: { cellWidth: 80, halign: 'center' },
        4: { cellWidth: 80, halign: 'center' },
        5: { cellWidth: 65, halign: 'center' },
        6: { cellWidth: 55, halign: 'center', fontStyle: 'bold' },
        7: { cellWidth: 45, halign: 'center' },
        8: { cellWidth: 55, halign: 'center' },
        9: { cellWidth: 80, halign: 'center' },
        10: { cellWidth: 50, halign: 'center' },
        11: { cellWidth: 65, halign: 'center' },
        12: { cellWidth: 65, halign: 'center' },
        13: { cellWidth: 145, halign: 'center' },
        14: { cellWidth: 75, halign: 'center' },
        15: { cellWidth: 65, halign: 'center' },
        16: { cellWidth: 65, halign: 'center' }
      };

      autoTable(doc, {
        head: [tableColumns],
        body: tableBody,
        startY: 85,
        tableWidth: pageW - 30,
        margin: { top: 85, right: 15, bottom: 25, left: 15 },
        theme: "grid",
        styles: {
          fontSize: 8.5,
          cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
          overflow: "linebreak",
          valign: 'middle',
          halign: 'center',
          textColor: [0, 0, 0], // Pure Black
          lineColor: [0, 0, 0], // Black grid lines
          lineWidth: 0.3,
          fontStyle: 'normal',
          minCellHeight: 12,
        },
        headStyles: {
          fillColor: [15, 23, 42],
          textColor: [255, 255, 255],
          fontStyle: "bold",
          lineColor: [0, 0, 0],
          lineWidth: 0.5,
          halign: 'center',
          fontSize: 9,
          valign: 'middle',
          cellPadding: { top: 5, right: 3, bottom: 5, left: 3 },
        },
        alternateRowStyles: {
          fillColor: [248, 250, 252],
        },
        columnStyles,
        didParseCell: function (data) {
          if (data.section === 'body') {
            const rowIndex = data.row.index;
            const isTotalRow = rowIndex === tableBody.length - 1;

            if (isTotalRow) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [226, 232, 240];
              data.cell.styles.textColor = [0, 0, 0];
              data.cell.styles.halign = 'center';
              return;
            }

            const item = filteredData[rowIndex];
            if (!item) return;

            // Lot number styling
            if (data.column.index === 1) {
              data.cell.styles.textColor = [220, 38, 38];
              data.cell.styles.fontStyle = 'bold';
            }

            // Total Pcs styling
            if (data.column.index === 6) {
              data.cell.styles.textColor = [220, 38, 38];
              data.cell.styles.fontStyle = 'bold';
            }

            // Status styling
            if (data.column.index === 15) {
              if (item.isCompleted) {
                data.cell.styles.fillColor = [220, 252, 231];
                data.cell.styles.textColor = [21, 128, 61];
                data.cell.styles.fontStyle = 'bold';
              } else {
                data.cell.styles.fillColor = [254, 243, 199];
                data.cell.styles.textColor = [180, 83, 9];
                data.cell.styles.fontStyle = 'bold';
              }
            }
          }
        }
      });

      // --- 4-COLUMN SIDE-BY-SIDE EXECUTIVE SUMMARY ---
      const gBody = sortedGarments.map(item => {
        const pct = totalPieces > 0 ? ((item.totalPcs / totalPieces) * 100).toFixed(1) : "0.0";
        return [item.name, item.totalLots.toString(), item.totalPcs.toLocaleString(), `${pct}%`];
      });
      gBody.push(["TOTAL", totalLots.toString(), totalPieces.toLocaleString(), "100.0%"]);

      const supBody = sortedSupervisors.map(item => {
        const pct = totalPieces > 0 ? ((item.totalPcs / totalPieces) * 100).toFixed(1) : "0.0";
        return [item.name, item.totalLots.toString(), item.totalPcs.toLocaleString(), `${pct}%`];
      });
      supBody.push(["TOTAL", totalLots.toString(), totalPieces.toLocaleString(), "100.0%"]);

      const seasonBody = sortedSeasons.map(item => {
        const pct = totalPieces > 0 ? ((item.totalPcs / totalPieces) * 100).toFixed(1) : "0.0";
        return [item.name, item.totalLots.toString(), item.totalPcs.toLocaleString(), `${pct}%`];
      });
      seasonBody.push(["TOTAL", totalLots.toString(), totalPieces.toLocaleString(), "100.0%"]);

      const statBody = [
        ["Completed Lots", completedLotsCount.toString(), completedPcsCount.toLocaleString(), `${totalPieces > 0 ? ((completedPcsCount / totalPieces) * 100).toFixed(1) : 0}%`],
        ["Pending Lots", pendingLotsCount.toString(), pendingPcsCount.toLocaleString(), `${totalPieces > 0 ? ((pendingPcsCount / totalPieces) * 100).toFixed(1) : 0}%`],
        ["TOTAL", totalLots.toString(), totalPieces.toLocaleString(), "100.0%"]
      ];

      const maxRows = Math.max(gBody.length, supBody.length, seasonBody.length, statBody.length);
      const approxSummaryHeight = 55 + (maxRows * 18);

      let summaryStartY = doc.lastAutoTable.finalY + 22;
      const neededSpace = approxSummaryHeight + 35;
      if (summaryStartY + neededSpace > pageH - 30) {
        doc.addPage();
        summaryStartY = 40;
      } else {
        doc.setDrawColor(203, 213, 225);
        doc.setLineWidth(0.8);
        doc.line(15, summaryStartY - 8, pageW - 15, summaryStartY - 8);
      }

      // Title & KPI Subtitle
      doc.setFontSize(14);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(0, 0, 0);
      doc.text("EXECUTIVE SUMMARY & PRODUCTION BREAKDOWN", pageW / 2, summaryStartY + 4, { align: 'center' });

      doc.setFontSize(9);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(0, 0, 0);
      const summarySub = `Total Lots: ${totalLots}   |   Total Pieces: ${totalPieces.toLocaleString()} Pcs   |   Supervisors: ${sortedSupervisors.length}   |   Garments: ${sortedGarments.length}   |   Seasons: ${sortedSeasons.length}`;
      doc.text(summarySub, pageW / 2, summaryStartY + 16, { align: 'center' });

      const sectionTitleY = summaryStartY + 30;
      const tableStartY = sectionTitleY + 6;

      // 4 Columns Side-by-Side Configuration (Exactly matching full page width)
      const colWidth = 278;
      const gap = 16;
      const col1X = 15;
      const col2X = col1X + colWidth + gap; // 309
      const col3X = col2X + colWidth + gap; // 603
      const col4X = col3X + colWidth + gap; // 897

      doc.setFontSize(9.5);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(0, 0, 0);
      doc.text("1. GARMENT BREAKDOWN", col1X, sectionTitleY);
      doc.text("2. SUPERVISOR BREAKDOWN", col2X, sectionTitleY);
      doc.text("3. SEASON BREAKDOWN", col3X, sectionTitleY);
      doc.text("4. STATUS DISTRIBUTION", col4X, sectionTitleY);

      const summaryColStyles = {
        0: { cellWidth: 110, halign: 'center' },
        1: { cellWidth: 45, halign: 'center' },
        2: { cellWidth: 68, halign: 'center' },
        3: { cellWidth: 55, halign: 'center' },
      };

      // Column 1 Table: Garment Breakdown
      autoTable(doc, {
        head: [['Garment Type', 'Lots', 'Total Pcs', 'Share %']],
        body: gBody,
        startY: tableStartY,
        tableWidth: colWidth,
        margin: { left: col1X, right: pageW - (col1X + colWidth) },
        theme: "grid",
        styles: {
          fontSize: 8.5,
          cellPadding: { top: 3.5, right: 2, bottom: 3.5, left: 2 },
          overflow: "linebreak",
          valign: 'middle',
          halign: 'center',
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          lineWidth: 0.3,
        },
        headStyles: {
          fillColor: [15, 118, 110], // Teal
          textColor: [255, 255, 255],
          fontStyle: "bold",
          fontSize: 8.5,
          halign: 'center',
          cellPadding: { top: 4, right: 2, bottom: 4, left: 2 },
        },
        columnStyles: summaryColStyles,
        didParseCell: function (data) {
          if (data.section === 'body') {
            data.cell.styles.textColor = [0, 0, 0];
            data.cell.styles.halign = 'center';
            if (data.row.index === gBody.length - 1) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });
      const endY1 = doc.lastAutoTable.finalY;

      // Column 2 Table: Supervisor Breakdown
      autoTable(doc, {
        head: [['Supervisor', 'Lots', 'Total Pcs', 'Share %']],
        body: supBody,
        startY: tableStartY,
        tableWidth: colWidth,
        margin: { left: col2X, right: pageW - (col2X + colWidth) },
        theme: "grid",
        styles: {
          fontSize: 8.5,
          cellPadding: { top: 3.5, right: 2, bottom: 3.5, left: 2 },
          overflow: "linebreak",
          valign: 'middle',
          halign: 'center',
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          lineWidth: 0.3,
        },
        headStyles: {
          fillColor: [67, 56, 202], // Indigo
          textColor: [255, 255, 255],
          fontStyle: "bold",
          fontSize: 8.5,
          halign: 'center',
          cellPadding: { top: 4, right: 2, bottom: 4, left: 2 },
        },
        columnStyles: summaryColStyles,
        didParseCell: function (data) {
          if (data.section === 'body') {
            data.cell.styles.textColor = [0, 0, 0];
            data.cell.styles.halign = 'center';
            if (data.row.index === supBody.length - 1) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });
      const endY2 = doc.lastAutoTable.finalY;

      // Column 3 Table: Season Breakdown
      autoTable(doc, {
        head: [['Season', 'Lots', 'Total Pcs', 'Share %']],
        body: seasonBody,
        startY: tableStartY,
        tableWidth: colWidth,
        margin: { left: col3X, right: pageW - (col3X + colWidth) },
        theme: "grid",
        styles: {
          fontSize: 8.5,
          cellPadding: { top: 3.5, right: 2, bottom: 3.5, left: 2 },
          overflow: "linebreak",
          valign: 'middle',
          halign: 'center',
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          lineWidth: 0.3,
        },
        headStyles: {
          fillColor: [30, 64, 175], // Royal Blue
          textColor: [255, 255, 255],
          fontStyle: "bold",
          fontSize: 8.5,
          halign: 'center',
          cellPadding: { top: 4, right: 2, bottom: 4, left: 2 },
        },
        columnStyles: summaryColStyles,
        didParseCell: function (data) {
          if (data.section === 'body') {
            data.cell.styles.textColor = [0, 0, 0];
            data.cell.styles.halign = 'center';
            if (data.row.index === seasonBody.length - 1) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });
      const endY3 = doc.lastAutoTable.finalY;

      // Column 4 Table: Status Breakdown
      autoTable(doc, {
        head: [['Status Category', 'Lots', 'Total Pcs', 'Share %']],
        body: statBody,
        startY: tableStartY,
        tableWidth: colWidth,
        margin: { left: col4X, right: pageW - (col4X + colWidth) },
        theme: "grid",
        styles: {
          fontSize: 8.5,
          cellPadding: { top: 3.5, right: 2, bottom: 3.5, left: 2 },
          overflow: "linebreak",
          valign: 'middle',
          halign: 'center',
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          lineWidth: 0.3,
        },
        headStyles: {
          fillColor: [180, 83, 9], // Amber
          textColor: [255, 255, 255],
          fontStyle: "bold",
          fontSize: 8.5,
          halign: 'center',
          cellPadding: { top: 4, right: 2, bottom: 4, left: 2 },
        },
        columnStyles: summaryColStyles,
        didParseCell: function (data) {
          if (data.section === 'body') {
            data.cell.styles.textColor = [0, 0, 0];
            data.cell.styles.halign = 'center';
            if (data.row.index === 0) {
              data.cell.styles.fillColor = [220, 252, 231]; // Soft Green
              data.cell.styles.fontStyle = 'bold';
            } else if (data.row.index === 1) {
              data.cell.styles.fillColor = [254, 243, 199]; // Soft Amber
              data.cell.styles.fontStyle = 'bold';
            } else if (data.row.index === 2) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });
      const endY4 = doc.lastAutoTable.finalY;

      const maxEndY = Math.max(endY1, endY2, endY3, endY4);
      const finalY = maxEndY + 16;
      if (finalY <= pageH - 22) {
        doc.setDrawColor(0, 0, 0);
        doc.setLineWidth(0.5);
        doc.line(15, finalY, pageW - 15, finalY);

        doc.setFontSize(8.5);
        doc.setFont('helvetica', 'italic');
        doc.setTextColor(0, 0, 0);
        doc.text("Elastic Department Report — Factory Suite Pro", 15, finalY + 12);
      }

      // Page Numbering Loop
      const totalPages = doc.internal.getNumberOfPages();
      const timeStr = `${new Date().toLocaleDateString('en-IN')} ${new Date().toLocaleTimeString('en-IN')}`;
      for (let i = 1; i <= totalPages; i++) {
        doc.setPage(i);
        doc.setFontSize(8.5);
        doc.setFont('helvetica', 'normal');
        doc.setTextColor(0, 0, 0);

        doc.setDrawColor(203, 213, 225);
        doc.setLineWidth(0.5);
        doc.line(15, pageH - 22, pageW - 15, pageH - 22);

        doc.text(`Page ${i} of ${totalPages}`, pageW / 2, pageH - 12, { align: 'center' });
        doc.text(`Generated: ${timeStr}`, pageW - 18, pageH - 12, { align: 'right' });
      }

      const fileDate = new Date().toISOString().slice(0, 10);
      doc.save(`Elastic_Department_Report_${fileDate}.pdf`);

    } catch (error) {
      console.error("Error generating Elastic Report PDF:", error);
      alert("Failed to generate PDF. Please try again.");
    }
  };

  // Professional Multi-Sheet Excel Export (matching PendingIssue format)
  const exportToExcel = async () => {
    if (filteredData.length === 0) {
      alert("No data available to export.");
      return;
    }

    try {
      const workbook = new ExcelJS.Workbook();
      workbook.creator = "Factory Suite Pro";
      workbook.created = new Date();

      const totalLotsCount = filteredData.length;
      const totalPiecesCount = filteredData.reduce((sum, item) => sum + (item.totalPcs || 0), 0);

      // Grouping for Executive Summary
      const garmentMap = {};
      const supervisorMap = {};
      const seasonMap = {};
      let compCount = 0, compPcs = 0;
      let pendCount = 0, pendPcs = 0;

      filteredData.forEach((item) => {
        const pcs = item.totalPcs || 0;
        const gType = (item.garmentType || "Unknown").trim();
        const sup = (item.supervisor || "Unassigned").trim();
        const season = (item.season || "Other / NA").trim();

        if (!garmentMap[gType]) garmentMap[gType] = { totalLots: 0, totalPcs: 0 };
        garmentMap[gType].totalLots += 1;
        garmentMap[gType].totalPcs += pcs;

        if (!supervisorMap[sup]) supervisorMap[sup] = { totalLots: 0, totalPcs: 0 };
        supervisorMap[sup].totalLots += 1;
        supervisorMap[sup].totalPcs += pcs;

        if (!seasonMap[season]) seasonMap[season] = { totalLots: 0, totalPcs: 0 };
        seasonMap[season].totalLots += 1;
        seasonMap[season].totalPcs += pcs;

        if (item.isCompleted) {
          compCount += 1;
          compPcs += pcs;
        } else {
          pendCount += 1;
          pendPcs += pcs;
        }
      });

      const sortedGarments = Object.keys(garmentMap)
        .map((name) => ({
          name,
          totalLots: garmentMap[name].totalLots,
          totalPcs: garmentMap[name].totalPcs,
        }))
        .sort((a, b) => b.totalPcs - a.totalPcs);

      const sortedSupervisors = Object.keys(supervisorMap)
        .map((name) => ({
          name,
          totalLots: supervisorMap[name].totalLots,
          totalPcs: supervisorMap[name].totalPcs,
        }))
        .sort((a, b) => b.totalPcs - a.totalPcs);

      const thinBorder = {
        top: { style: "thin", color: { argb: "FFCBD5E1" } },
        bottom: { style: "thin", color: { argb: "FFCBD5E1" } },
        left: { style: "thin", color: { argb: "FFCBD5E1" } },
        right: { style: "thin", color: { argb: "FFCBD5E1" } },
      };

      // ================= SHEET 1: MAIN DATA =================
      const ws1 = workbook.addWorksheet("Elastic Report", {
        views: [{ showGridLines: true }],
      });

      // Title Banner
      ws1.mergeCells("A1:Q1");
      const titleCell = ws1.getCell("A1");
      titleCell.value = "FACTORY SUITE PRO - ELASTIC ATTACHMENT REPORT";
      titleCell.font = { name: "Segoe UI", size: 14, bold: true, color: { argb: "FFFFFFFF" } };
      titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F172A" } };
      titleCell.alignment = { horizontal: "center", vertical: "middle" };
      ws1.getRow(1).height = 32;

      // Subtitle
      ws1.mergeCells("A2:Q2");
      const subCell = ws1.getCell("A2");
      subCell.value = `Report Date: ${new Date().toLocaleDateString("en-IN")}  |  Total Lots: ${totalLotsCount}  |  Total Pieces: ${totalPiecesCount.toLocaleString()}  |  Completed Lots: ${compCount}  |  Pending Lots: ${pendCount}`;
      subCell.font = { name: "Segoe UI", size: 9.5, color: { argb: "FFCBD5E1" } };
      subCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E293B" } };
      subCell.alignment = { horizontal: "center", vertical: "middle" };
      ws1.getRow(2).height = 22;

      // Spacer
      ws1.addRow([]);

      // Table Headers
      const headers1 = [
        "Sr No.",
        "Lot Number",
        "Garment Type",
        "Style",
        "Fabric",
        "Brand",
        "Total Pcs",
        "Section",
        "Season",
        "Party Name",
        "Direct Stitching",
        "Elastic Date",
        "Elastic Supervisor",
        "Stitching Supervisor",
        "Status",
        "WIP Remarks",
        "Completed At",
      ];
      const headerRow = ws1.addRow(headers1);
      headerRow.height = 25;
      headerRow.eachCell((cell) => {
        cell.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E293B" } };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = thinBorder;
      });

      // Data Rows
      filteredData.forEach((item, idx) => {
        const isComp = item.isCompleted;

        const row = ws1.addRow([
          idx + 1,
          item.lotNumber || "—",
          item.garmentType || "—",
          item.style || "—",
          item.fabric || "—",
          item.brand || "—",
          item.totalPcs || 0,
          item.section || "—",
          item.season || "—",
          item.party || "—",
          item.directStitching || "—",
          item.elasticDate || "—",
          item.supervisor || "—",
          item.stitchingSupervisor || "—",
          isComp ? "Completed" : "In Progress",
          item.wip || "—",
          isComp ? formatDisplayDate(item.completeDate) : "—",
        ]);

        row.height = 20;
        row.eachCell((cell) => {
          cell.font = { name: "Segoe UI", size: 9.5 };
          cell.border = thinBorder;
          cell.alignment = { horizontal: "center", vertical: "middle" };
          if (idx % 2 === 1) {
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
          }
        });

        row.getCell(2).font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFDC2626" } };
        row.getCell(7).numFmt = "#,##0";
        row.getCell(7).font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFDC2626" } };

        if (isComp) {
          row.getCell(15).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDCFCE7" } };
          row.getCell(15).font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FF15803D" } };
        } else {
          row.getCell(15).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE0E7FF" } };
          row.getCell(15).font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FF3730A3" } };
        }
      });

      // Total Row
      const totalRow1 = ws1.addRow([
        "",
        `TOTAL (${totalLotsCount} Lots)`,
        "",
        "",
        "",
        "",
        totalPiecesCount,
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        `${compCount} Comp | ${pendCount} Pending`,
        "",
        "",
      ]);
      totalRow1.height = 24;
      totalRow1.eachCell((cell) => {
        cell.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF000000" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE2E8F0" } };
        cell.border = {
          top: { style: "thin", color: { argb: "FF000000" } },
          bottom: { style: "double", color: { argb: "FF000000" } },
          left: { style: "thin", color: { argb: "FFCBD5E1" } },
          right: { style: "thin", color: { argb: "FFCBD5E1" } },
        };
        cell.alignment = { horizontal: "center", vertical: "middle" };
      });
      totalRow1.getCell(7).numFmt = "#,##0";

      const colWidths = [8, 16, 20, 20, 20, 16, 14, 12, 14, 20, 16, 16, 22, 20, 16, 30, 18];
      colWidths.forEach((w, i) => {
        ws1.getColumn(i + 1).width = w;
      });

      // ================= SHEET 2: EXECUTIVE SUMMARY =================
      const ws2 = workbook.addWorksheet("Executive Summary", {
        views: [{ showGridLines: true }],
      });

      // Section 1: Garment Type Breakdown
      ws2.mergeCells("A1:D1");
      const gTitle = ws2.getCell("A1");
      gTitle.value = "1. GARMENT TYPE BREAKDOWN (LOTS & PIECES DISTRIBUTION)";
      gTitle.font = { name: "Segoe UI", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      gTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F766E" } };
      gTitle.alignment = { horizontal: "left", vertical: "middle" };
      ws2.getRow(1).height = 26;

      const gHeader = ws2.addRow(["Garment Type", "Total Lots", "Total Pieces (Qty)", "Share %"]);
      gHeader.height = 22;
      gHeader.eachCell((c) => {
        c.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFFFFFFF" } };
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF134E4A" } };
        c.alignment = { horizontal: "center", vertical: "middle" };
        c.border = thinBorder;
      });

      sortedGarments.forEach((item, idx) => {
        const pct = totalPiecesCount > 0 ? item.totalPcs / totalPiecesCount : 0;
        const r = ws2.addRow([item.name, item.totalLots, item.totalPcs, pct]);
        r.height = 19;
        r.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
        r.getCell(2).alignment = { horizontal: "center", vertical: "middle" };
        r.getCell(3).alignment = { horizontal: "center", vertical: "middle" };
        r.getCell(3).numFmt = "#,##0";
        r.getCell(4).alignment = { horizontal: "center", vertical: "middle" };
        r.getCell(4).numFmt = "0.0%";
        r.eachCell((c) => {
          c.font = { name: "Segoe UI", size: 9 };
          c.border = thinBorder;
          if (idx % 2 === 1) c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
        });
      });

      const gTotalRow = ws2.addRow(["TOTAL", totalLotsCount, totalPiecesCount, 1]);
      gTotalRow.height = 22;
      gTotalRow.eachCell((c) => {
        c.font = { name: "Segoe UI", size: 9.5, bold: true };
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE2E8F0" } };
        c.border = { top: { style: "thin" }, bottom: { style: "double" }, left: { style: "thin" }, right: { style: "thin" } };
        c.alignment = { horizontal: "center", vertical: "middle" };
      });
      gTotalRow.getCell(3).numFmt = "#,##0";
      gTotalRow.getCell(4).numFmt = "0.0%";

      // Spacer
      ws2.addRow([]);

      // Section 2: Supervisor Breakdown
      const supStartRow = ws2.rowCount + 1;
      ws2.mergeCells(`A${supStartRow}:D${supStartRow}`);
      const supTitle = ws2.getCell(`A${supStartRow}`);
      supTitle.value = "2. ELASTIC SUPERVISOR WORKLOAD";
      supTitle.font = { name: "Segoe UI", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      supTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E40AF" } };
      supTitle.alignment = { horizontal: "left", vertical: "middle" };
      ws2.getRow(supStartRow).height = 26;

      const supHeader = ws2.addRow(["Supervisor Name", "Total Lots", "Total Pieces (Qty)", "Share %"]);
      supHeader.height = 22;
      supHeader.eachCell((c) => {
        c.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFFFFFFF" } };
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E3A8A" } };
        c.alignment = { horizontal: "center", vertical: "middle" };
        c.border = thinBorder;
      });

      sortedSupervisors.forEach((item, idx) => {
        const pct = totalPiecesCount > 0 ? item.totalPcs / totalPiecesCount : 0;
        const r = ws2.addRow([item.name, item.totalLots, item.totalPcs, pct]);
        r.height = 19;
        r.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
        r.getCell(2).alignment = { horizontal: "center", vertical: "middle" };
        r.getCell(3).alignment = { horizontal: "center", vertical: "middle" };
        r.getCell(3).numFmt = "#,##0";
        r.getCell(4).alignment = { horizontal: "center", vertical: "middle" };
        r.getCell(4).numFmt = "0.0%";
        r.eachCell((c) => {
          c.font = { name: "Segoe UI", size: 9 };
          c.border = thinBorder;
          if (idx % 2 === 1) c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
        });
      });

      const supTotalRow = ws2.addRow(["TOTAL", totalLotsCount, totalPiecesCount, 1]);
      supTotalRow.height = 22;
      supTotalRow.eachCell((c) => {
        c.font = { name: "Segoe UI", size: 9.5, bold: true };
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE2E8F0" } };
        c.border = { top: { style: "thin" }, bottom: { style: "double" }, left: { style: "thin" }, right: { style: "thin" } };
        c.alignment = { horizontal: "center", vertical: "middle" };
      });
      supTotalRow.getCell(3).numFmt = "#,##0";
      supTotalRow.getCell(4).numFmt = "0.0%";

      // Spacer
      ws2.addRow([]);

      // Section 3: Status Breakdown
      const staStartRow = ws2.rowCount + 1;
      ws2.mergeCells(`A${staStartRow}:D${staStartRow}`);
      const staTitle = ws2.getCell(`A${staStartRow}`);
      staTitle.value = "3. COMPLETION STATUS SUMMARY";
      staTitle.font = { name: "Segoe UI", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      staTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4338CA" } };
      staTitle.alignment = { horizontal: "left", vertical: "middle" };
      ws2.getRow(staStartRow).height = 26;

      const staHeader = ws2.addRow(["Status", "Total Lots", "Total Pieces (Qty)", "Share %"]);
      staHeader.height = 22;
      staHeader.eachCell((c) => {
        c.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFFFFFFF" } };
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF312E81" } };
        c.alignment = { horizontal: "center", vertical: "middle" };
        c.border = thinBorder;
      });

      const statusData = [
        { name: "Completed", lots: compCount, pcs: compPcs, bg: "FFDCFCE7", fg: "FF15803D" },
        { name: "In Progress / Pending", lots: pendCount, pcs: pendPcs, bg: "FFE0E7FF", fg: "FF3730A3" },
      ];

      statusData.forEach((item) => {
        const pct = totalPiecesCount > 0 ? item.pcs / totalPiecesCount : 0;
        const r = ws2.addRow([item.name, item.lots, item.pcs, pct]);
        r.height = 20;
        r.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
        r.getCell(2).alignment = { horizontal: "center", vertical: "middle" };
        r.getCell(3).alignment = { horizontal: "center", vertical: "middle" };
        r.getCell(3).numFmt = "#,##0";
        r.getCell(4).alignment = { horizontal: "center", vertical: "middle" };
        r.getCell(4).numFmt = "0.0%";
        r.eachCell((c) => {
          c.font = { name: "Segoe UI", size: 9, bold: true, color: { argb: item.fg } };
          c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: item.bg } };
          c.border = thinBorder;
        });
      });

      const staTotalRow = ws2.addRow(["TOTAL", totalLotsCount, totalPiecesCount, 1]);
      staTotalRow.height = 22;
      staTotalRow.eachCell((c) => {
        c.font = { name: "Segoe UI", size: 9.5, bold: true };
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE2E8F0" } };
        c.border = { top: { style: "thin" }, bottom: { style: "double" }, left: { style: "thin" }, right: { style: "thin" } };
        c.alignment = { horizontal: "center", vertical: "middle" };
      });
      staTotalRow.getCell(3).numFmt = "#,##0";
      staTotalRow.getCell(4).numFmt = "0.0%";

      ws2.getColumn(1).width = 34;
      ws2.getColumn(2).width = 16;
      ws2.getColumn(3).width = 22;
      ws2.getColumn(4).width = 16;

      // ================= SHEET 3: APPLIED FILTERS =================
      const ws3 = workbook.addWorksheet("Applied Filters", {
        views: [{ showGridLines: true }],
      });

      ws3.mergeCells("A1:B1");
      const fTitle = ws3.getCell("A1");
      fTitle.value = "APPLIED FILTERS & METADATA";
      fTitle.font = { name: "Segoe UI", size: 12, bold: true, color: { argb: "FFFFFFFF" } };
      fTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E293B" } };
      fTitle.alignment = { horizontal: "left", vertical: "middle" };
      ws3.getRow(1).height = 28;

      const filterItems = [
        ["Report Generated", `${new Date().toLocaleDateString("en-IN")} ${new Date().toLocaleTimeString("en-IN")}`],
        ["Total Lots Exported", totalLotsCount],
        ["Total Pieces Exported", totalPiecesCount.toLocaleString()],
        ["Status Filter", selectedStatuses.length ? selectedStatuses.join(", ") : "All Statuses"],
        ["Supervisor Filter", selectedSupervisors.length ? selectedSupervisors.join(", ") : "All Supervisors"],
        ["Stitching Supervisor Filter", selectedStitchingSupervisors.length ? selectedStitchingSupervisors.join(", ") : "All"],
        ["Garment Type Filter", selectedGarments.length ? selectedGarments.join(", ") : "All Types"],
        ["Fabric Filter", selectedFabrics.length ? selectedFabrics.join(", ") : "All Fabrics"],
        ["Brand Filter", selectedBrands.length ? selectedBrands.join(", ") : "All Brands"],
        ["Search Query", searchTerm || "None"],
      ];

      filterItems.forEach(([k, v], idx) => {
        const r = ws3.addRow([k, v]);
        r.height = 20;
        r.getCell(1).font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF1E293B" } };
        r.getCell(2).font = { name: "Segoe UI", size: 9.5, color: { argb: "FF334155" } };
        r.getCell(1).alignment = { horizontal: "left", vertical: "middle" };
        r.getCell(2).alignment = { horizontal: "left", vertical: "middle" };
        r.eachCell((c) => {
          c.border = thinBorder;
          if (idx % 2 === 1) c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
        });
      });

      ws3.getColumn(1).width = 24;
      ws3.getColumn(2).width = 50;

      const buffer = await workbook.xlsx.writeBuffer();
      const ts = new Date().toISOString().slice(0, 10);
      saveAs(new Blob([buffer]), `Elastic_Attachment_Report_${ts}.xlsx`);
    } catch (e) {
      console.error("Excel export error:", e);
      alert(`Excel export failed: ${e.message}`);
    }
  };

  // CSV Export
  const exportToCSV = () => {
    if (filteredData.length === 0) return;
    const headers = [
      "Sr No.",
      "Lot Number",
      "Garment Type",
      "Style",
      "Brand",
      "Fabric",
      "Season",
      "Section",
      "Pcs",
      "Stitching Supervisor",
      "Elastic Date",
      "Wip Remarks",
      "Elastic Sup",
      "Status",
      "Completed At",
    ];
    const rows = filteredData.map((d, idx) => [
      idx + 1,
      `"${d.lotNumber}"`,
      `"${d.garmentType}"`,
      `"${d.style}"`,
      `"${d.brand}"`,
      `"${d.fabric}"`,
      `"${d.season}"`,
      `"${d.section}"`,
      d.totalPcs,
      `"${d.stitchingSupervisor}"`,
      `"${d.elasticDate}"`,
      `"${d.wip}"`,
      `"${d.supervisor}"`,
      d.isCompleted ? "Completed" : "In Progress",
      `"${d.completeDate}"`,
    ]);
    const csvContent =
      "data:text/csv;charset=utf-8," + [headers.join(","), ...rows.map((r) => r.join(","))].join("\n");
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", `Elastic_Report_${new Date().toISOString().slice(0, 10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const handlePrint = () => {
    window.print();
  };

  return (
    <>
      <style>{`
        .elastic-report-container {
          padding: 24px 28px 80px;
          background-color: #f8fafc;
          min-height: 100vh;
          font-family: 'Plus Jakarta Sans', 'Inter', system-ui, -apple-system, sans-serif;
          color: #0f172a;
          box-sizing: border-box;
        }
        .report-header-box {
          background: linear-gradient(135deg, #4f46e5 0%, #3730a3 100%);
          border-radius: 20px;
          padding: 22px 28px;
          display: flex;
          justify-content: space-between;
          align-items: center;
          color: white;
          box-shadow: 0 10px 25px -5px rgba(79, 70, 229, 0.25);
          margin-bottom: 22px;
          flex-wrap: wrap;
          gap: 16px;
        }
        .header-btn {
          background: rgba(255, 255, 255, 0.15);
          backdrop-filter: blur(8px);
          border: 1px solid rgba(255, 255, 255, 0.25);
          color: white;
          padding: 9px 16px;
          border-radius: 12px;
          font-weight: 700;
          font-size: 0.85rem;
          cursor: pointer;
          transition: all 0.2s ease;
          display: inline-flex;
          align-items: center;
          gap: 6px;
        }
        .header-btn:hover:not(:disabled) {
          background: rgba(255, 255, 255, 0.28);
          transform: translateY(-1px);
        }
        .header-btn-pdf {
          background: #ef4444 !important;
          border-color: #dc2626 !important;
          box-shadow: 0 4px 12px rgba(239, 68, 68, 0.35);
        }
        .header-btn-pdf:hover:not(:disabled) {
          background: #dc2626 !important;
          box-shadow: 0 6px 16px rgba(220, 38, 38, 0.45);
        }
        .stat-cards-grid {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
          gap: 14px;
          margin-bottom: 22px;
        }
        .stat-card {
          background: white;
          border-radius: 16px;
          padding: 16px 18px;
          border: 1px solid #e2e8f0;
          box-shadow: 0 2px 4px rgba(0, 0, 0, 0.02);
          transition: transform 0.2s;
        }
        .stat-card:hover {
          transform: translateY(-2px);
          box-shadow: 0 8px 16px rgba(0, 0, 0, 0.05);
        }
        .stat-card-highlight {
          border: 2px solid #f59e0b;
          background: linear-gradient(180deg, #fffbeb 0%, #ffffff 100%);
        }
        .stat-label {
          font-size: 0.74rem;
          font-weight: 800;
          color: #64748b;
          text-transform: uppercase;
          letter-spacing: 0.5px;
          margin-bottom: 4px;
        }
        .stat-value {
          font-size: 1.75rem;
          font-weight: 800;
          line-height: 1.2;
          color: #1e1b4b;
        }
        .stat-subtext {
          font-size: 0.75rem;
          color: #94a3b8;
          font-weight: 600;
          margin-top: 4px;
        }

        /* Filter Panel */
        .filter-panel {
          background: white;
          border-radius: 20px;
          padding: 18px 22px;
          border: 1px solid #e2e8f0;
          margin-bottom: 22px;
          box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.02);
        }
        .filter-grid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(160px, 1fr));
          gap: 12px;
          align-items: flex-end;
        }
        .search-input {
          width: 100%;
          padding: 8px 14px;
          border: 1.5px solid #cbd5e1;
          border-radius: 10px;
          font-size: 0.85rem;
          font-weight: 600;
          outline: none;
          box-sizing: border-box;
          transition: border-color 0.2s;
        }
        .search-input:focus {
          border-color: #6366f1;
        }
        .date-input {
          width: 100%;
          padding: 7px 10px;
          border: 1.5px solid #cbd5e1;
          border-radius: 10px;
          font-size: 0.82rem;
          font-weight: 600;
          outline: none;
          box-sizing: border-box;
        }

        /* Table Card */
        .table-card {
          background: white;
          border-radius: 20px;
          border: 1px solid #e2e8f0;
          overflow: hidden;
          box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.03);
          margin-bottom: 24px;
        }
        .table-header-bar {
          padding: 16px 24px;
          display: flex;
          justify-content: space-between;
          align-items: center;
          border-bottom: 1px solid #f1f5f9;
        }
        .table-wrapper {
          overflow-x: auto;
          max-height: 68vh;
        }
        .report-table {
          width: 100%;
          border-collapse: separate;
          border-spacing: 0;
          font-size: 0.85rem;
        }
        .report-table th {
          background: #f8fafc;
          padding: 12px 14px;
          font-weight: 800;
          font-size: 0.72rem;
          text-transform: uppercase;
          letter-spacing: 0.5px;
          color: #475569;
          border-bottom: 1.5px solid #e2e8f0;
          position: sticky;
          top: 0;
          z-index: 10;
          cursor: pointer;
          user-select: none;
          white-space: nowrap;
          text-align: left;
        }
        .report-table th:hover {
          background: #f1f5f9;
          color: #0f172a;
        }
        .report-table td {
          padding: 12px 14px;
          border-bottom: 1px solid #f1f5f9;
          vertical-align: middle;
          color: #334155;
          font-weight: 500;
        }
        .report-table tr:hover td {
          background: #faf5ff;
        }
        .report-table tfoot td {
          padding: 14px 14px;
          border-top: 2px solid #cbd5e1;
          vertical-align: middle;
          background: #f8fafc;
        }
        .lot-badge {
          display: inline-flex;
          align-items: center;
          gap: 6px;
          padding: 3px 8px;
          border-radius: 8px;
          background: #fee2e2;
          color: #ef4444;
          font-weight: 800;
          font-size: 0.86rem;
          cursor: pointer;
          border: 1px solid #fca5a5;
          transition: all 0.15s ease;
        }
        .lot-badge:hover {
          background: #fecaca;
          transform: scale(1.03);
        }
        .status-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 3px 10px;
          border-radius: 20px;
          font-size: 0.75rem;
          font-weight: 800;
          letter-spacing: 0.2px;
          white-space: nowrap;
        }
        .status-completed {
          background: #dcfce7;
          color: #15803d;
          border: 1px solid #86efac;
        }
        .status-pending {
          background: #fef3c7;
          color: #b45309;
          border: 1px solid #fde68a;
        }
        .tag-pill {
          display: inline-block;
          padding: 2px 7px;
          border-radius: 6px;
          font-size: 0.74rem;
          font-weight: 700;
          background: #f1f5f9;
          color: #475569;
        }
        .manpower-badge {
          display: inline-flex;
          align-items: center;
          gap: 4px;
          padding: 2px 8px;
          border-radius: 12px;
          font-size: 0.75rem;
          font-weight: 700;
          background: #ede9fe;
          color: #6d28d9;
          border: 1px solid #ddd6fe;
        }

        /* Pagination Bar */
        .pagination-bar {
          padding: 14px 24px;
          display: flex;
          justify-content: space-between;
          align-items: center;
          border-top: 1px solid #f1f5f9;
          background: #ffffff;
        }
        .page-btn {
          padding: 6px 12px;
          border-radius: 8px;
          border: 1.5px solid #cbd5e1;
          background: white;
          font-weight: 700;
          font-size: 0.8rem;
          cursor: pointer;
          transition: all 0.15s;
        }
        .page-btn:hover:not(:disabled) {
          border-color: #6366f1;
          color: #4f46e5;
          background: #eef2ff;
        }
        .page-btn:disabled {
          opacity: 0.4;
          cursor: not-allowed;
        }

        @media print {
          .no-print {
            display: none !important;
          }
          .elastic-report-container {
            padding: 0 !important;
            background: white !important;
          }
          .report-table th, .report-table td {
            padding: 6px 8px !important;
            font-size: 8pt !important;
          }
        }
      `}</style>

      <div className="elastic-report-container">
        {/* Header Box */}
        <div className="report-header-box no-print">
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
              <span style={{ fontSize: "1.8rem" }}></span>
              <div>
                <h1 style={{ margin: 0, fontSize: "1.5rem", fontWeight: 800, letterSpacing: "-0.02em" }}>
                  Elastic Attachment Report
                </h1>
                <p style={{ margin: "2px 0 0", fontSize: "0.82rem", color: "#c7d2fe", fontWeight: 600 }}>
                  Tracking daily elastic issuance, pending pieces, stitching supervisor & complete lots
                </p>
              </div>
            </div>
          </div>

          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", alignItems: "center" }}>
            <button type="button" className="header-btn" onClick={() => history.push("/dashboard")}>
              ← Dashboard
            </button>
            <button type="button" className="header-btn" onClick={fetchData} disabled={loading}>
              🔄 {loading ? "Loading..." : "Refresh"}
            </button>
            <button
              type="button"
              className="header-btn header-btn-pdf"
              onClick={handleDownloadPDF}
              disabled={filteredData.length === 0}
              title="Download clean PDF report matching Stitching Report style"
            >
              📄 Download PDF
            </button>
            <button type="button" className="header-btn" onClick={exportToExcel} disabled={filteredData.length === 0}>
              📊 Export Excel
            </button>
            <button type="button" className="header-btn" onClick={exportToCSV} disabled={filteredData.length === 0}>
              📥 CSV
            </button>
            <button type="button" className="header-btn" onClick={handlePrint} disabled={filteredData.length === 0}>
              🖨️ Print
            </button>
          </div>
        </div>

        {/* Top KPI Stat Cards */}
        <div className="stat-cards-grid">
          {/* 1. Total Lots */}
          <div className="stat-card">
            <div className="stat-label">Total Elastic Lots</div>
            <div className="stat-value" style={{ color: "#4f46e5" }}>
              {totalLots}
            </div>
            <div className="stat-subtext">Issued Lots in System</div>
          </div>

          {/* 2. Pending Pieces (Hero Metric) */}
          <div className="stat-card stat-card-highlight">
            <div className="stat-label" style={{ color: "#d97706" }}>
              Pending Pieces (Qty)
            </div>
            <div className="stat-value" style={{ color: "#d97706" }}>
              {pendingPcs.toLocaleString()}
            </div>
            <div className="stat-subtext">Across {pendingLots} Pending Lots</div>
          </div>

          {/* 3. In Progress / Pending Lots */}
          <div className="stat-card">
            <div className="stat-label">⏳ In Progress / Pending</div>
            <div className="stat-value" style={{ color: "#b45309" }}>
              {pendingLots}
            </div>
            <div className="stat-subtext">Lots under attachment</div>
          </div>

          {/* 4. Completed Lots */}
          <div className="stat-card">
            <div className="stat-label">✓ Completed Lots</div>
            <div className="stat-value" style={{ color: "#16a34a" }}>
              {completedLots}
            </div>
            <div className="stat-subtext">{completionRate}% Completion Rate</div>
          </div>

          {/* 5. Stitching Supervisors */}
          <div className="stat-card">
            <div className="stat-label">🧵 Stitching Supervisors</div>
            <div className="stat-value" style={{ color: "#6d28d9" }}>
              {stitchingSupervisorOptions.length}
            </div>
            <div className="stat-subtext">
              {stitchingSupervisorOptions.map((s) => s.label).slice(0, 3).join(", ") || "—"}
            </div>
          </div>

          {/* 6. Active Elastic Supervisors */}
          <div className="stat-card">
            <div className="stat-label">👤 Elastic Supervisor</div>
            <div className="stat-value" style={{ color: "#2563eb" }}>
              {supervisorOptions.length}
            </div>
            <div className="stat-subtext">
              {supervisorOptions.map((s) => s.label).join(", ") || "None"}
            </div>
          </div>

          {/* 7. Total Manpower */}
          <div className="stat-card">
            <div className="stat-label">⚙️ Total Manpower</div>
            <div className="stat-value" style={{ color: "#7c3aed" }}>
              {totalManpower}
            </div>
            <div className="stat-subtext">Allocated Operators</div>
          </div>
        </div>

        {/* Filter Panel */}
        <div className="filter-panel no-print">
          <div className="filter-grid">
            {/* Search Input */}
            <div>
              <label
                style={{
                  fontSize: "0.72rem",
                  fontWeight: 800,
                  color: "#64748b",
                  textTransform: "uppercase",
                  letterSpacing: "0.4px",
                  marginBottom: "4px",
                  display: "block",
                }}
              >
                Search Filter
              </label>
              <input
                type="text"
                placeholder="Search Lot, Supervisor, Fabric, Season..."
                value={searchTerm}
                onChange={(e) => {
                  setSearchTerm(e.target.value);
                  setCurrentPage(1);
                }}
                className="search-input"
              />
            </div>

            {/* Status Filter */}
            <MultiSelectDropdown
              label="Status"
              options={statusOptions}
              selectedValues={selectedStatuses}
              onChange={(v) => {
                setSelectedStatuses(v);
                setCurrentPage(1);
              }}
              themeColor="#4f46e5"
            />

            {/* Elastic Supervisor Filter */}
            <MultiSelectDropdown
              label="Elastic Supervisor"
              options={supervisorOptions}
              selectedValues={selectedSupervisors}
              onChange={(v) => {
                setSelectedSupervisors(v);
                setCurrentPage(1);
              }}
              themeColor="#2563eb"
            />

            {/* Stitching Supervisor Filter */}
            <MultiSelectDropdown
              label="Stitching Supervisor"
              options={stitchingSupervisorOptions}
              selectedValues={selectedStitchingSupervisors}
              onChange={(v) => {
                setSelectedStitchingSupervisors(v);
                setCurrentPage(1);
              }}
              themeColor="#7c3aed"
            />

            {/* Brand Filter */}
            <MultiSelectDropdown
              label="Brand"
              options={brandOptions}
              selectedValues={selectedBrands}
              onChange={(v) => {
                setSelectedBrands(v);
                setCurrentPage(1);
              }}
              themeColor="#4f46e5"
            />

            {/* Garment Type Filter */}
            <MultiSelectDropdown
              label="Garment Type"
              options={garmentOptions}
              selectedValues={selectedGarments}
              onChange={(v) => {
                setSelectedGarments(v);
                setCurrentPage(1);
              }}
              themeColor="#4f46e5"
            />

            {/* Fabric Filter */}
            <MultiSelectDropdown
              label="Fabric"
              options={fabricOptions}
              selectedValues={selectedFabrics}
              onChange={(v) => {
                setSelectedFabrics(v);
                setCurrentPage(1);
              }}
              themeColor="#4f46e5"
            />

            {/* Style Filter */}
            <MultiSelectDropdown
              label="Style"
              options={styleOptions}
              selectedValues={selectedStyles}
              onChange={(v) => {
                setSelectedStyles(v);
                setCurrentPage(1);
              }}
              themeColor="#4f46e5"
            />

            {/* Season Filter */}
            <MultiSelectDropdown
              label="Season"
              options={seasonOptions}
              selectedValues={selectedSeasons}
              onChange={(v) => {
                setSelectedSeasons(v);
                setCurrentPage(1);
              }}
              themeColor="#0891b2"
            />

            {/* Section Filter */}
            <MultiSelectDropdown
              label="Section"
              options={sectionOptions}
              selectedValues={selectedSections}
              onChange={(v) => {
                setSelectedSections(v);
                setCurrentPage(1);
              }}
              themeColor="#0891b2"
            />

            {/* Date Range Filters */}
            <div>
              <label
                style={{
                  fontSize: "0.72rem",
                  fontWeight: 800,
                  color: "#64748b",
                  textTransform: "uppercase",
                  letterSpacing: "0.4px",
                  marginBottom: "4px",
                  display: "block",
                }}
              >
                From Date
              </label>
              <input
                type="date"
                value={startDate}
                onChange={(e) => {
                  setStartDate(e.target.value);
                  setCurrentPage(1);
                }}
                className="date-input"
              />
            </div>

            <div>
              <label
                style={{
                  fontSize: "0.72rem",
                  fontWeight: 800,
                  color: "#64748b",
                  textTransform: "uppercase",
                  letterSpacing: "0.4px",
                  marginBottom: "4px",
                  display: "block",
                }}
              >
                To Date
              </label>
              <input
                type="date"
                value={endDate}
                onChange={(e) => {
                  setEndDate(e.target.value);
                  setCurrentPage(1);
                }}
                className="date-input"
              />
            </div>

            {/* Clear Filters Button */}
            {hasActiveFilters && (
              <div>
                <button
                  type="button"
                  onClick={resetFilters}
                  style={{
                    padding: "8px 14px",
                    borderRadius: "10px",
                    border: "1.5px solid #fca5a5",
                    background: "#fef2f2",
                    color: "#dc2626",
                    fontWeight: 700,
                    fontSize: "0.82rem",
                    cursor: "pointer",
                    height: "36px",
                    width: "100%",
                  }}
                >
                  ✕ Reset
                </button>
              </div>
            )}
          </div>
        </div>

        {/* Error Message */}
        {error && (
          <div
            style={{
              background: "#fee2e2",
              border: "1px solid #f87171",
              borderRadius: "12px",
              padding: "14px 20px",
              color: "#991b1b",
              fontWeight: 600,
              fontSize: "0.9rem",
              marginBottom: "20px",
            }}
          >
            ⚠️ {error}
          </div>
        )}

        {/* Main Table Card */}
        <div className="table-card">
          <div className="table-header-bar no-print">
            <div style={{ fontWeight: 800, fontSize: "0.95rem", color: "#1e293b", display: "flex", alignItems: "center", gap: "12px" }}>
              <span>Showing {filteredData.length} of {data.length} Lots</span>
              <span style={{ fontSize: "0.85rem", color: "#b45309", fontWeight: 800, background: "#fef3c7", border: "1px solid #fde68a", padding: "2px 8px", borderRadius: "6px" }}>
                Pending: {pendingPcs.toLocaleString()} Pcs
              </span>
            </div>

            <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
              <span style={{ fontSize: "0.8rem", color: "#64748b", fontWeight: 700 }}>Rows per page:</span>
              <select
                value={pageSize}
                onChange={(e) => {
                  setPageSize(Number(e.target.value));
                  setCurrentPage(1);
                }}
                style={{
                  padding: "4px 8px",
                  borderRadius: "8px",
                  border: "1px solid #cbd5e1",
                  fontSize: "0.82rem",
                  fontWeight: 700,
                  outline: "none",
                }}
              >
                <option value={10}>10</option>
                <option value={25}>25</option>
                <option value={50}>50</option>
                <option value={100}>100</option>
                <option value={1000}>All</option>
              </select>
            </div>
          </div>

          <div className="table-wrapper">
            <table className="report-table">
              <thead>
                <tr>
                  {/* 15 Ordered Headers:
                      1. Sr No.
                      2. Lot Number
                      3. Garment Type
                      4. Style
                      5. Brand
                      6. Fabric
                      7. Season
                      8. Section
                      9. Pcs
                      10. Stitching Supervisor
                      11. Elastic Date
                      12. Wip Remarks
                      13. Elastic Sup
                      14. Status
                      15. Completed At */}
                  <th onClick={() => handleSort("id")}>
                    Sr No. {sortField === "id" && (sortAsc ? "▲" : "▼")}
                  </th>
                  <th onClick={() => handleSort("lotNumber")}>
                    Lot Number {sortField === "lotNumber" && (sortAsc ? "▲" : "▼")}
                  </th>
                  <th onClick={() => handleSort("garmentType")}>
                    Garment Type {sortField === "garmentType" && (sortAsc ? "▲" : "▼")}
                  </th>
                  <th onClick={() => handleSort("style")}>
                    Style {sortField === "style" && (sortAsc ? "▲" : "▼")}
                  </th>
                  <th onClick={() => handleSort("fabric")}>
                    Fabric {sortField === "fabric" && (sortAsc ? "▲" : "▼")}
                  </th>
                  <th onClick={() => handleSort("brand")}>
                    Brand {sortField === "brand" && (sortAsc ? "▲" : "▼")}
                  </th>
                  <th onClick={() => handleSort("totalPcs")}>
                    Pcs {sortField === "totalPcs" && (sortAsc ? "▲" : "▼")}
                  </th>
                  <th onClick={() => handleSort("section")}>
                    Section {sortField === "section" && (sortAsc ? "▲" : "▼")}
                  </th>
                  <th onClick={() => handleSort("season")}>
                    Season {sortField === "season" && (sortAsc ? "▲" : "▼")}
                  </th>
                  <th>Party Name</th>
                  <th>Direct Stitching</th>
                  <th onClick={() => handleSort("stitchingSupervisor")}>
                    Stitching Supervisor {sortField === "stitchingSupervisor" && (sortAsc ? "▲" : "▼")}
                  </th>
                  <th onClick={() => handleSort("elasticDate")}>
                    Elastic Date {sortField === "elasticDate" && (sortAsc ? "▲" : "▼")}
                  </th>
                  <th>Wip Remarks</th>
                  <th onClick={() => handleSort("supervisor")}>
                    Elastic Sup {sortField === "supervisor" && (sortAsc ? "▲" : "▼")}
                  </th>
                  <th onClick={() => handleSort("isCompleted")}>
                    Status {sortField === "isCompleted" && (sortAsc ? "▲" : "▼")}
                  </th>
                  <th onClick={() => handleSort("completeDate")}>
                    Completed At {sortField === "completeDate" && (sortAsc ? "▲" : "▼")}
                  </th>
                </tr>
              </thead>
              <tbody>
                {loading ? (
                  <tr>
                    <td colSpan={17} style={{ textAlign: "center", padding: "40px", color: "#64748b" }}>
                      ⏳ Loading Elastic Report & Stitching Supervisor data from Google Sheets...
                    </td>
                  </tr>
                ) : paginatedData.length === 0 ? (
                  <tr>
                    <td colSpan={17} style={{ textAlign: "center", padding: "40px", color: "#94a3b8" }}>
                      No lots found matching your filter criteria.
                    </td>
                  </tr>
                ) : (
                  paginatedData.map((row, idx) => (
                    <tr key={row.id}>
                      {/* 1. Sr No. */}
                      <td style={{ color: "#94a3b8", fontSize: "0.78rem" }}>
                        {(currentPage - 1) * pageSize + idx + 1}
                      </td>

                      {/* 2. Lot Number */}
                      <td>
                        <span
                          className="lot-badge"
                          onClick={() => copyLot(row.lotNumber)}
                          title="Click to copy lot number"
                        >
                          {row.lotNumber}
                          <span style={{ fontSize: "10px" }}>{copiedLot === row.lotNumber ? "✓" : "📋"}</span>
                        </span>
                      </td>

                      {/* 3. Garment Type */}
                      <td>
                        <span className="tag-pill">{row.garmentType || "—"}</span>
                      </td>

                      {/* 4. Style */}
                      <td style={{ maxWidth: "140px", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={row.style}>
                        {row.style || "—"}
                      </td>

                      {/* 5. Fabric */}
                      <td style={{ maxWidth: "150px", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={row.fabric}>
                        {row.fabric || "—"}
                      </td>

                      {/* 6. Brand */}
                      <td>
                        <span style={{ fontWeight: 700, color: "#475569" }}>{row.brand || "—"}</span>
                      </td>

                      {/* 7. Pcs */}
                      <td style={{ fontWeight: 800, color: "#ef4444", fontSize: "0.9rem" }}>
                        {row.totalPcs.toLocaleString()}
                      </td>

                      {/* 8. Section */}
                      <td>
                        <span className="tag-pill" style={{ background: "#fef3c7", color: "#92400e" }}>{row.section || "—"}</span>
                      </td>

                      {/* 9. Season */}
                      <td>
                        <span className="tag-pill" style={{ background: "#e0f2fe", color: "#0369a1" }}>{row.season || "—"}</span>
                      </td>

                      {/* 10. Party Name */}
                      <td style={{ fontWeight: 600, color: "#1e293b" }}>
                        {row.partyName || "—"}
                      </td>

                      {/* 11. Direct Stitching */}
                      <td>
                        <span className="tag-pill">{row.directStitching || "—"}</span>
                      </td>

                      {/* 12. Stitching Supervisor */}
                      <td>
                        <span
                          style={{
                            fontWeight: 700,
                            color: "#6b21a8",
                            background: "#f3e8ff",
                            border: "1px solid #e9d5ff",
                            padding: "3px 8px",
                            borderRadius: "6px",
                            fontSize: "0.78rem",
                            display: "inline-block",
                            whiteSpace: "nowrap",
                          }}
                        >
                          🧵 {row.stitchingSupervisor || "—"}
                        </span>
                      </td>

                      {/* 11. Elastic Date */}
                      <td style={{ whiteSpace: "nowrap", fontWeight: 700 }}>
                        {formatDisplayDate(row.elasticDate)}
                      </td>

                      {/* 12. Wip Remarks */}
                      <td style={{ maxWidth: "160px", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={row.isCompleted ? "Done" : row.wip}>
                        {row.isCompleted ? (
                          <span style={{ fontSize: "0.78rem", fontWeight: 700, color: "#16a34a" }}>Done</span>
                        ) : (
                          <span style={{ fontSize: "0.78rem", color: "#64748b" }}>{row.wip || "—"}</span>
                        )}
                      </td>

                      {/* 13. Elastic Sup */}
                      <td>
                        <span
                          style={{
                            fontWeight: 700,
                            color: "#1e1b4b",
                            background: "#e0e7ff",
                            padding: "3px 8px",
                            borderRadius: "6px",
                            fontSize: "0.78rem",
                            display: "inline-block",
                            whiteSpace: "nowrap",
                          }}
                        >
                          👤 {row.supervisor}
                        </span>
                      </td>

                      {/* 14. Status */}
                      <td>
                        {row.isCompleted ? (
                          <span className="status-badge status-completed">✓ Completed</span>
                        ) : (
                          <span className="status-badge status-pending">⏳ In Progress</span>
                        )}
                      </td>

                      {/* 15. Completed At */}
                      <td style={{ whiteSpace: "nowrap", fontSize: "0.78rem", color: row.isCompleted ? "#16a34a" : "#94a3b8" }}>
                        {row.completeDate ? formatDisplayDate(row.completeDate) : "—"}
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
              {filteredData.length > 0 && (
                <tfoot>
                  <tr style={{ background: "#f8fafc", fontWeight: 800, borderTop: "2px solid #cbd5e1" }}>
                    <td colSpan={2} style={{ textAlign: "center", color: "#1e293b", fontSize: "0.9rem", fontWeight: 800 }}>
                      TOTAL ({filteredData.length} LOTS)
                    </td>
                    <td colSpan={6} style={{ textAlign: "right", color: "#b45309", textTransform: "uppercase", fontSize: "0.82rem", letterSpacing: "0.5px", fontWeight: 800 }}>
                      PENDING PIECES:
                    </td>
                    <td style={{ fontWeight: 800, color: "#d97706", fontSize: "1.05rem" }}>
                      {pendingPcs.toLocaleString()}
                    </td>
                    <td colSpan={6}>
                      <span style={{ fontSize: "0.8rem", color: "#64748b", fontWeight: 700 }}>
                        <b style={{ color: "#d97706" }}>{pendingLots}</b> Pending Lots ({pendingPcs.toLocaleString()} pcs)  |  <b style={{ color: "#16a34a" }}>{completedLots}</b> Completed Lots
                      </span>
                    </td>
                  </tr>
                </tfoot>
              )}
            </table>
          </div>

          {/* Pagination Bar */}
          <div className="pagination-bar no-print">
            <div style={{ fontSize: "0.82rem", color: "#64748b", fontWeight: 600 }}>
              Page {currentPage} of {totalPages} ({filteredData.length} total items)
            </div>

            <div style={{ display: "flex", gap: "6px" }}>
              <button
                type="button"
                className="page-btn"
                onClick={() => setCurrentPage(1)}
                disabled={currentPage === 1}
              >
                « First
              </button>
              <button
                type="button"
                className="page-btn"
                onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                disabled={currentPage === 1}
              >
                ‹ Prev
              </button>

              <span
                style={{
                  display: "inline-flex",
                  alignItems: "center",
                  padding: "0 10px",
                  fontSize: "0.82rem",
                  fontWeight: 700,
                  color: "#334155",
                }}
              >
                {currentPage} / {totalPages}
              </span>

              <button
                type="button"
                className="page-btn"
                onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                disabled={currentPage === totalPages}
              >
                Next ›
              </button>
              <button
                type="button"
                className="page-btn"
                onClick={() => setCurrentPage(totalPages)}
                disabled={currentPage === totalPages}
              >
                Last »
              </button>
            </div>
          </div>
        </div>
      </div>
    </>
  );
}
