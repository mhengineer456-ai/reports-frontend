import React, { useState, useEffect, useMemo, useRef } from "react";
import { useHistory } from "react-router-dom";
import { SPREADSHEET_IDS, fetchSheetDataFromBackend } from "./config";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

const SPREADSHEET_ID = "1IMhmYlJ3s2PPRgEQs1Ikd4O1OBXK4EYL1oV_-kWAkyg";
const SHEET_NAME = "Washing";

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

// Function to format general date strings as DD/MM/YY
function formatDateToDisplay(dateString) {
  if (!dateString || (typeof dateString !== "string" && typeof dateString !== "number")) {
    return "";
  }

  try {
    let clean = String(dateString).trim().replace(/^['"\s]+|['"\s]+$/g, "");
    if (!clean || clean === "-" || clean === "N/A" || clean === "—" || clean === "[]") return "";

    // Match ISO string YYYY-MM-DD or YYYY/MM/DD
    const isoMatch = clean.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
    if (isoMatch) {
      const y = isoMatch[1].slice(-2);
      const m = isoMatch[2].padStart(2, "0");
      const d = isoMatch[3].padStart(2, "0");
      return `${d}/${m}/${y}`;
    }

    // Match DD/MM/YYYY or DD-MM-YYYY or DD.MM.YYYY
    const parts = clean.split(/[\/\-\.]/);
    if (parts.length === 3) {
      let day = parseInt(parts[0], 10);
      let month = parseInt(parts[1], 10);
      let year = parseInt(parts[2], 10);

      if (day > 1000) {
        const tmp = day;
        day = year;
        year = tmp;
      }

      if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
        const fullYear = year < 100 ? 2000 + year : year;
        return `${String(day).padStart(2, "0")}/${String(month).padStart(2, "0")}/${String(fullYear).slice(-2)}`;
      }
    }

    const date = new Date(clean);
    if (!isNaN(date.getTime())) {
      const day = String(date.getDate()).padStart(2, "0");
      const month = String(date.getMonth() + 1).padStart(2, "0");
      const year = String(date.getFullYear()).slice(-2);
      return `${day}/${month}/${year}`;
    }

    return clean;
  } catch {
    return String(dateString);
  }
}

// Function to parse completion status / JSON and extract ONLY the date
function formatCompletionDate(raw) {
  if (!raw || raw === "[]" || raw === "-" || raw === "—" || raw === "N/A") return "";
  let clean = String(raw).trim();

  // If JSON array or object
  if (clean.startsWith("[") || clean.startsWith("{")) {
    try {
      const parsed = JSON.parse(clean);
      const arr = Array.isArray(parsed) ? parsed : [parsed];
      if (arr.length > 0) {
        // Look for the last valid entry with date or timestamp
        for (let idx = arr.length - 1; idx >= 0; idx--) {
          const item = arr[idx];
          if (typeof item === "string" && item.trim()) {
            clean = item.trim();
            break;
          } else if (item && typeof item === "object") {
            const extracted = item.timestamp || item.date || item.completionDate || item.completedDate || item.time || item.dateTime;
            if (extracted) {
              clean = String(extracted).trim();
              break;
            }
          }
        }
      }
    } catch (e) {
      // Regex extraction fallback if JSON parsing errors
      const match = clean.match(/"(?:timestamp|date|completionDate|completedDate|time)":\s*"([^"]+)"/i);
      if (match) {
        clean = match[1];
      }
    }
  }

  if (!clean || clean === "-" || clean === "—" || clean === "[]" || clean === "N/A") return "";
  return formatDateToDisplay(clean);
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
                  fontWeight: isChecked ? "800" : "500",
                  color: isChecked ? themeColor : "#1e293b",
                  background: isChecked ? "rgba(2, 132, 199, 0.08)" : "transparent",
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

export default function WashingReport() {
  const history = useHistory();
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  // Multi-Select Filters
  const [searchTerm, setSearchTerm] = useState("");
  const [selectedStatuses, setSelectedStatuses] = useState([]);
  const [selectedBrands, setSelectedBrands] = useState([]);
  const [selectedGarments, setSelectedGarments] = useState([]);
  const [selectedFabrics, setSelectedFabrics] = useState([]);
  const [selectedSeasons, setSelectedSeasons] = useState([]);
  const [selectedSections, setSelectedSections] = useState([]);
  const [selectedPlants, setSelectedPlants] = useState([]);
  const [selectedSupervisors, setSelectedSupervisors] = useState([]);
  const [dateFilter, setDateFilter] = useState("");

  const fetchData = async () => {
    setLoading(true);
    setError(null);
    try {
      const [washingRes, jobOrderRes] = await Promise.allSettled([
        fetchSheetDataFromBackend(SPREADSHEET_ID, `'${SHEET_NAME}'!A:Z`),
        fetchSheetDataFromBackend(SPREADSHEET_IDS.JOBORDER, `'JobOrder'!A:AZ`),
      ]);

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
        const jPlantIdx = findCol(jHeaders, ["washing plant", "washing-plant", "washingplant", "plant", "washing unit", "wash plant", "washing vendor", "vendor"]);
        const jDirectIdx = findCol(jHeaders, ["direct stitching", "directstitch", "direct"]);

        if (jLotIdx !== -1) {
          for (let i = 1; i < jRows.length; i++) {
            const r = jRows[i];
            const lotKey = String(r[jLotIdx] || "").trim().toLowerCase();
            if (lotKey) {
              lotToJobInfo[lotKey] = {
                brand: (jBrandIdx !== -1 ? String(r[jBrandIdx] || "").trim() : "") || (jPartyIdx !== -1 ? String(r[jPartyIdx] || "").trim() : ""),
                party: jPartyIdx !== -1 ? String(r[jPartyIdx] || "").trim() : "",
                garment: jGarmentIdx !== -1 ? String(r[jGarmentIdx] || "").trim() : "",
                style: jStyleIdx !== -1 ? String(r[jStyleIdx] || "").trim() : "",
                fabric: jFabricIdx !== -1 ? String(r[jFabricIdx] || "").trim() : "",
                season: jSeasonIdx !== -1 ? String(r[jSeasonIdx] || "").trim() : "",
                section: jSectionIdx !== -1 ? String(r[jSectionIdx] || "").trim() : "",
                washingPlant: jPlantIdx !== -1 ? String(r[jPlantIdx] || "").trim() : "",
                directStitching: jDirectIdx !== -1 ? String(r[jDirectIdx] || "").trim() : "",
              };
            }
          }
        }
      }

      if (
        washingRes.status !== "fulfilled" ||
        !washingRes.value?.ok ||
        !Array.isArray(washingRes.value.values) ||
        washingRes.value.values.length === 0
      ) {
        setData([]);
        setLoading(false);
        return;
      }

      const rows = washingRes.value.values;
      const headers = rows[0] || [];

      const tsIdx = findCol(headers, ["timestamp", "time"]);
      const lotIdx = findCol(headers, ["lot number", "lot no", "lot"]);
      const garmentIdx = findCol(headers, ["garment type", "garment"]);
      const fabricIdx = findCol(headers, ["fabric"]);
      const styleIdx = findCol(headers, ["style"]);
      const brandIdx = findCol(headers, ["brand", "brand name"]);
      const seasonIdx = findCol(headers, ["season"]);
      const sectionIdx = findCol(headers, ["section"]);
      const plantIdx = findCol(headers, ["washing plant", "washing-plant", "washingplant", "plant", "washing unit", "wash plant", "washing vendor", "vendor", "unit", "washing party", "plant name"]);
      const supIdx = findCol(headers, ["washing supervisor", "supervisor"]);
      const dateIdx = findCol(headers, ["washing date", "date"]);
      const pcsIdx = findCol(headers, ["total pcs", "pcs", "quantity"]);
      const wipIdx = findCol(headers, ["wip washing", "wip", "remarks"]);
      const compIdx = findCol(headers, ["washing complete", "complete", "completed"]);

      const parsed = [];
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const lotNo = String(row[lotIdx !== -1 ? lotIdx : 1] || "").trim();
        if (!lotNo) continue;

        const pcs = parseInt(row[pcsIdx !== -1 ? pcsIdx : 8], 10) || 0;
        const rawComp = row[compIdx !== -1 ? compIdx : 10] || "";
        const formattedComp = formatCompletionDate(rawComp);
        const wipVal = String(row[wipIdx !== -1 ? wipIdx : 9] || "").trim();
        const isCompleted = !!formattedComp || (!!rawComp && rawComp !== "[]" && rawComp !== "-" && !String(rawComp).toLowerCase().includes("pending"));

        const normLot = lotNo.toLowerCase();
        const jobInfo = lotToJobInfo[normLot] || {};
        const sheetBrand = brandIdx !== -1 ? String(row[brandIdx] || "").trim() : "";
        const finalBrand = sheetBrand && sheetBrand !== "-" ? sheetBrand : (jobInfo.brand || jobInfo.party || "—");
        const sheetSeason = seasonIdx !== -1 ? String(row[seasonIdx] || "").trim() : "";
        const finalSeason = sheetSeason && sheetSeason !== "-" ? sheetSeason : (jobInfo.season || "—");
        const sheetSection = sectionIdx !== -1 ? String(row[sectionIdx] || "").trim() : "";
        const finalSection = sheetSection && sheetSection !== "-" ? sheetSection : (jobInfo.section || "—");
        const sheetPlant = plantIdx !== -1 ? String(row[plantIdx] || "").trim() : "";
        const finalPlant = sheetPlant && sheetPlant !== "-" ? sheetPlant : (jobInfo.washingPlant || "—");
        const cleanWashingDate = formatDateToDisplay(row[dateIdx !== -1 ? dateIdx : 7] || "");

        parsed.push({
          id: i,
          timestamp: row[tsIdx !== -1 ? tsIdx : 0] || "",
          lotNumber: lotNo,
          garmentType: row[garmentIdx !== -1 ? garmentIdx : 2] || jobInfo.garment || "",
          fabric: row[fabricIdx !== -1 ? fabricIdx : 3] || jobInfo.fabric || "",
          style: row[styleIdx !== -1 ? styleIdx : 4] || jobInfo.style || "",
          brand: finalBrand,
          season: finalSeason,
          section: finalSection,
          partyName: jobInfo.party || "—",
          directStitching: jobInfo.directStitching || "—",
          washingPlant: finalPlant,
          supervisor: row[supIdx !== -1 ? supIdx : 6] || "Washing Department",
          washingDate: cleanWashingDate || row[dateIdx !== -1 ? dateIdx : 7] || "",
          totalPcs: pcs,
          wip: wipVal === "[]" ? "In Progress" : wipVal || "In Progress",
          completeDate: formattedComp || (isCompleted ? "Completed" : ""),
          isCompleted
        });
      }

      setData(parsed);
    } catch (err) {
      console.error("Error loading Washing data:", err);
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

  const seasonOptions = useMemo(() => {
    const set = new Set(data.map((d) => d.season).filter((s) => s && s !== "—" && s !== "-"));
    return Array.from(set).sort().map((s) => ({ value: s, label: s }));
  }, [data]);

  const sectionOptions = useMemo(() => {
    const set = new Set(data.map((d) => d.section).filter((s) => s && s !== "—" && s !== "-"));
    return Array.from(set).sort().map((s) => ({ value: s, label: s }));
  }, [data]);

  const plantOptions = useMemo(() => {
    const set = new Set(data.map((d) => d.washingPlant).filter((p) => p && p !== "—" && p !== "-"));
    return Array.from(set).sort().map((p) => ({ value: p, label: p }));
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
        (item.season && item.season.toLowerCase().includes(term)) ||
        (item.section && item.section.toLowerCase().includes(term)) ||
        (item.washingPlant && item.washingPlant.toLowerCase().includes(term)) ||
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
      const matchesSeason = selectedSeasons.length === 0 || selectedSeasons.some((s) => s.toLowerCase() === (item.season || "").toLowerCase());
      const matchesSection = selectedSections.length === 0 || selectedSections.some((s) => s.toLowerCase() === (item.section || "").toLowerCase());
      const matchesPlant = selectedPlants.length === 0 || selectedPlants.some((p) => p.toLowerCase() === (item.washingPlant || "").toLowerCase());
      const matchesSupervisor = selectedSupervisors.length === 0 || selectedSupervisors.some((s) => s.toLowerCase() === item.supervisor.toLowerCase());
      const matchesDate = !dateFilter || item.washingDate.includes(dateFilter);

      return matchesSearch && matchesStatus && matchesBrand && matchesGarment && matchesFabric && matchesSeason && matchesSection && matchesPlant && matchesSupervisor && matchesDate;
    });
  }, [data, searchTerm, selectedStatuses, selectedBrands, selectedGarments, selectedFabrics, selectedSeasons, selectedSections, selectedPlants, selectedSupervisors, dateFilter]);

  const resetFilters = () => {
    setSearchTerm("");
    setSelectedStatuses([]);
    setSelectedBrands([]);
    setSelectedGarments([]);
    setSelectedFabrics([]);
    setSelectedSeasons([]);
    setSelectedSections([]);
    setSelectedPlants([]);
    setSelectedSupervisors([]);
    setDateFilter("");
  };

  const hasActiveFilters =
    searchTerm !== "" ||
    selectedStatuses.length > 0 ||
    selectedBrands.length > 0 ||
    selectedGarments.length > 0 ||
    selectedFabrics.length > 0 ||
    selectedSeasons.length > 0 ||
    selectedSections.length > 0 ||
    selectedPlants.length > 0 ||
    selectedSupervisors.length > 0 ||
    dateFilter !== "";

  // Summary Metrics for filtered data
  const totalLots = filteredData.length;
  const totalPcs = filteredData.reduce((sum, item) => sum + item.totalPcs, 0);
  const completedLots = filteredData.filter((item) => item.isCompleted).length;
  const pendingLots = totalLots - completedLots;

  // Professional Multi-Sheet Excel Export (Matching Factory Suite Pro Standard)
  const exportToExcel = async () => {
    if (filteredData.length === 0) {
      alert("No data available to export.");
      return;
    }

    try {
      const workbook = new ExcelJS.Workbook();
      workbook.creator = "Factory Suite Pro";
      workbook.created = new Date();

      // Executive Summary aggregations
      const plantMap = {};
      const garmentMap = {};
      const supervisorMap = {};

      filteredData.forEach((item) => {
        const plant = (item.washingPlant || "Unassigned").trim();
        const garment = (item.garmentType || "Unknown").trim();
        const sup = (item.supervisor || "Unassigned").trim();
        const pcs = item.totalPcs || 0;
        const isComp = item.isCompleted;

        if (!plantMap[plant]) plantMap[plant] = { lots: 0, pcs: 0, compLots: 0, pendLots: 0 };
        plantMap[plant].lots += 1;
        plantMap[plant].pcs += pcs;
        if (isComp) plantMap[plant].compLots += 1;
        else plantMap[plant].pendLots += 1;

        if (!garmentMap[garment]) garmentMap[garment] = { lots: 0, pcs: 0, compLots: 0, pendLots: 0 };
        garmentMap[garment].lots += 1;
        garmentMap[garment].pcs += pcs;
        if (isComp) garmentMap[garment].compLots += 1;
        else garmentMap[garment].pendLots += 1;

        if (!supervisorMap[sup]) supervisorMap[sup] = { lots: 0, pcs: 0, compLots: 0, pendLots: 0 };
        supervisorMap[sup].lots += 1;
        supervisorMap[sup].pcs += pcs;
        if (isComp) supervisorMap[sup].compLots += 1;
        else supervisorMap[sup].pendLots += 1;
      });

      const sortedPlants = Object.keys(plantMap).map(k => ({ name: k, ...plantMap[k] })).sort((a, b) => b.pcs - a.pcs);
      const sortedGarments = Object.keys(garmentMap).map(k => ({ name: k, ...garmentMap[k] })).sort((a, b) => b.pcs - a.pcs);
      const sortedSupervisors = Object.keys(supervisorMap).map(k => ({ name: k, ...supervisorMap[k] })).sort((a, b) => b.pcs - a.pcs);

      const thinBorder = {
        top: { style: "thin", color: { argb: "FFCBD5E1" } },
        bottom: { style: "thin", color: { argb: "FFCBD5E1" } },
        left: { style: "thin", color: { argb: "FFCBD5E1" } },
        right: { style: "thin", color: { argb: "FFCBD5E1" } }
      };

      // ================= SHEET 1: WASHING DATA =================
      const ws1 = workbook.addWorksheet("Washing Report", {
        views: [{ showGridLines: true }]
      });

      // Title Banner
      ws1.mergeCells("A1:Q1");
      const titleCell = ws1.getCell("A1");
      titleCell.value = "FACTORY SUITE PRO - WASHING DEPARTMENT REPORT";
      titleCell.font = { name: "Segoe UI", size: 14, bold: true, color: { argb: "FFFFFFFF" } };
      titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F172A" } };
      titleCell.alignment = { horizontal: "center", vertical: "middle" };
      ws1.getRow(1).height = 32;

      // Subtitle
      ws1.mergeCells("A2:Q2");
      const subCell = ws1.getCell("A2");
      subCell.value = `Report Date: ${new Date().toLocaleDateString("en-IN")}  |  Total Lots: ${totalLots}  |  Total Pieces: ${totalPcs.toLocaleString()}  |  Completed Lots: ${completedLots}  |  Pending Lots: ${pendingLots}`;
      subCell.font = { name: "Segoe UI", size: 9.5, color: { argb: "FFCBD5E1" } };
      subCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E293B" } };
      subCell.alignment = { horizontal: "center", vertical: "middle" };
      ws1.getRow(2).height = 22;

      // Spacer
      ws1.addRow([]);

      // Table Headers
      const headers1 = [
        "Sr No.", "Lot Number", "Garment Type", "Style", "Fabric", "Brand",
        "Total Pcs", "Section", "Season", "Party Name", "Direct Stitching",
        "Washing Plant", "Washing Date", "WIP Remarks", "Washing Supervisor",
        "Status", "Completion Date"
      ];
      const headerRow = ws1.addRow(headers1);
      headerRow.height = 25;
      headerRow.eachCell((cell) => {
        cell.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F172A" } };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = thinBorder;
      });

      // Data Rows
      filteredData.forEach((d, idx) => {
        const rowData = [
          idx + 1,
          d.lotNumber || "—",
          d.garmentType || "—",
          d.style || "—",
          d.fabric || "—",
          d.brand || "—",
          d.totalPcs || 0,
          d.section || "—",
          d.season || "—",
          d.partyName || "—",
          d.directStitching || "—",
          d.washingPlant || "—",
          d.washingDate || "—",
          d.isCompleted ? "Done" : (d.wip || "—"),
          d.supervisor || "—",
          d.isCompleted ? "Completed" : "Pending",
          d.isCompleted ? (d.completeDate || "Completed") : "—"
        ];

        const row = ws1.addRow(rowData);
        row.height = 20;

        const isEven = idx % 2 === 0;
        const rowBgColor = isEven ? "FFFFFFFF" : "FFF8FAFC";

        row.eachCell((cell, colNumber) => {
          cell.border = thinBorder;
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: rowBgColor } };
          cell.alignment = { horizontal: "center", vertical: "middle" };
          cell.font = { name: "Segoe UI", size: 9, color: { argb: "FF1E293B" } };

          // Lot Number styling
          if (colNumber === 2) {
            cell.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFDC2626" } };
          }
          // Brand styling
          if (colNumber === 6) {
            cell.font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FF1E293B" } };
          }
          // Total Pcs styling
          if (colNumber === 7) {
            cell.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFDC2626" } };
            cell.numFmt = "#,##0";
          }
          // Washing Plant styling
          if (colNumber === 12 && d.washingPlant && d.washingPlant !== "—") {
            cell.font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FF0369A1" } };
          }
          // Supervisor styling
          if (colNumber === 15) {
            cell.font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FF1E293B" } };
          }
          // Status column badge
          if (colNumber === 16) {
            if (d.isCompleted) {
              cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDCFCE7" } };
              cell.font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FF15803D" } };
            } else {
              cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEF3C7" } };
              cell.font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FFB45309" } };
            }
          }
        });
      });

      // Total Summary Row
      const totalRow = ws1.addRow([
        "",
        "TOTAL",
        `${totalLots} Lots`,
        "",
        "",
        "TOTAL PCS:",
        totalPcs,
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        `${pendingLots} Pending Lots`,
        `${completedLots} Completed Lots`,
        ""
      ]);
      totalRow.height = 24;
      totalRow.eachCell((cell, colNumber) => {
        cell.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF0F172A" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE2E8F0" } };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = {
          top: { style: "medium", color: { argb: "FF0F172A" } },
          bottom: { style: "double", color: { argb: "FF0F172A" } },
          left: { style: "thin", color: { argb: "FFCBD5E1" } },
          right: { style: "thin", color: { argb: "FFCBD5E1" } }
        };
        if (colNumber === 7) {
          cell.numFmt = "#,##0";
          cell.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFDC2626" } };
        }
      });

      // Auto-fit column widths
      ws1.columns.forEach((col) => {
        let maxLen = 12;
        col.eachCell({ includeEmpty: false }, (cell) => {
          const valStr = cell.value ? cell.value.toString() : "";
          if (valStr.length > maxLen) {
            maxLen = Math.min(valStr.length, 36);
          }
        });
        col.width = Math.max(maxLen + 3, 11);
      });

      // ================= SHEET 2: EXECUTIVE SUMMARY =================
      const ws2 = workbook.addWorksheet("Executive Summary", {
        views: [{ showGridLines: true }]
      });

      // Summary Header
      ws2.mergeCells("A1:G1");
      const sumTitle = ws2.getCell("A1");
      sumTitle.value = "WASHING DEPARTMENT - EXECUTIVE KPI & WORKLOAD BREAKDOWN";
      sumTitle.font = { name: "Segoe UI", size: 13, bold: true, color: { argb: "FFFFFFFF" } };
      sumTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F172A" } };
      sumTitle.alignment = { horizontal: "center", vertical: "middle" };
      ws2.getRow(1).height = 30;

      // Section 1: Washing Plant Summary
      let curRow = 3;
      ws2.mergeCells(`A${curRow}:F${curRow}`);
      const plantSec = ws2.getCell(`A${curRow}`);
      plantSec.value = "🏭 WASHING PLANT WORKLOAD BREAKDOWN";
      plantSec.font = { name: "Segoe UI", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      plantSec.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0369A1" } };
      plantSec.alignment = { horizontal: "left", vertical: "middle", indent: 1 };
      ws2.getRow(curRow).height = 24;

      curRow++;
      const pHeaders = ["Washing Plant", "Total Lots", "Total Pieces", "Completed Lots", "Pending Lots", "% Share"];
      const pHeaderRow = ws2.addRow(pHeaders);
      pHeaderRow.height = 22;
      pHeaderRow.eachCell((c) => {
        c.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFFFFFFF" } };
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E293B" } };
        c.alignment = { horizontal: "center", vertical: "middle" };
        c.border = thinBorder;
      });

      sortedPlants.forEach((p) => {
        curRow++;
        const share = totalPcs > 0 ? (p.pcs / totalPcs) * 100 : 0;
        const r = ws2.addRow([
          p.name,
          p.lots,
          p.pcs,
          p.compLots,
          p.pendLots,
          `${share.toFixed(1)}%`
        ]);
        r.height = 19;
        r.eachCell((c, colIdx) => {
          c.border = thinBorder;
          c.alignment = { horizontal: "center", vertical: "middle" };
          c.font = { name: "Segoe UI", size: 9 };
          if (colIdx === 1) {
            c.alignment = { horizontal: "left", vertical: "middle", indent: 1 };
            c.font = { name: "Segoe UI", size: 9, bold: true };
          }
          if (colIdx === 3) {
            c.numFmt = "#,##0";
            c.font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FFDC2626" } };
          }
        });
      });

      // Section 2: Garment Type Summary
      curRow += 3;
      ws2.mergeCells(`A${curRow}:F${curRow}`);
      const gSec = ws2.getCell(`A${curRow}`);
      gSec.value = "👕 GARMENT TYPE DISTRIBUTION";
      gSec.font = { name: "Segoe UI", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      gSec.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F766E" } };
      gSec.alignment = { horizontal: "left", vertical: "middle", indent: 1 };
      ws2.getRow(curRow).height = 24;

      curRow++;
      const gHeaders = ["Garment Type", "Total Lots", "Total Pieces", "Completed Lots", "Pending Lots", "% Share"];
      const gHeaderRow = ws2.addRow(gHeaders);
      gHeaderRow.height = 22;
      gHeaderRow.eachCell((c) => {
        c.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFFFFFFF" } };
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E293B" } };
        c.alignment = { horizontal: "center", vertical: "middle" };
        c.border = thinBorder;
      });

      sortedGarments.forEach((g) => {
        curRow++;
        const share = totalPcs > 0 ? (g.pcs / totalPcs) * 100 : 0;
        const r = ws2.addRow([
          g.name,
          g.lots,
          g.pcs,
          g.compLots,
          g.pendLots,
          `${share.toFixed(1)}%`
        ]);
        r.height = 19;
        r.eachCell((c, colIdx) => {
          c.border = thinBorder;
          c.alignment = { horizontal: "center", vertical: "middle" };
          c.font = { name: "Segoe UI", size: 9 };
          if (colIdx === 1) {
            c.alignment = { horizontal: "left", vertical: "middle", indent: 1 };
            c.font = { name: "Segoe UI", size: 9, bold: true };
          }
          if (colIdx === 3) {
            c.numFmt = "#,##0";
            c.font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FFDC2626" } };
          }
        });
      });

      // Section 3: Supervisor Summary
      curRow += 3;
      ws2.mergeCells(`A${curRow}:E${curRow}`);
      const sSec = ws2.getCell(`A${curRow}`);
      sSec.value = "👤 SUPERVISOR ALLOCATION";
      sSec.font = { name: "Segoe UI", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      sSec.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E40AF" } };
      sSec.alignment = { horizontal: "left", vertical: "middle", indent: 1 };
      ws2.getRow(curRow).height = 24;

      curRow++;
      const sHeaders = ["Supervisor", "Total Lots", "Total Pieces", "Completed Lots", "Pending Lots"];
      const sHeaderRow = ws2.addRow(sHeaders);
      sHeaderRow.height = 22;
      sHeaderRow.eachCell((c) => {
        c.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFFFFFFF" } };
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E293B" } };
        c.alignment = { horizontal: "center", vertical: "middle" };
        c.border = thinBorder;
      });

      sortedSupervisors.forEach((s) => {
        curRow++;
        const r = ws2.addRow([
          s.name,
          s.lots,
          s.pcs,
          s.compLots,
          s.pendLots
        ]);
        r.height = 19;
        r.eachCell((c, colIdx) => {
          c.border = thinBorder;
          c.alignment = { horizontal: "center", vertical: "middle" };
          c.font = { name: "Segoe UI", size: 9 };
          if (colIdx === 1) {
            c.alignment = { horizontal: "left", vertical: "middle", indent: 1 };
            c.font = { name: "Segoe UI", size: 9, bold: true };
          }
          if (colIdx === 3) {
            c.numFmt = "#,##0";
            c.font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FFDC2626" } };
          }
        });
      });

      // Column widths for Sheet 2
      ws2.columns = [
        { width: 28 },
        { width: 15 },
        { width: 16 },
        { width: 18 },
        { width: 16 },
        { width: 15 },
        { width: 15 }
      ];

      // ================= SHEET 3: APPLIED FILTERS & METADATA =================
      const ws3 = workbook.addWorksheet("Applied Filters", {
        views: [{ showGridLines: true }]
      });

      ws3.mergeCells("A1:D1");
      const fTitle = ws3.getCell("A1");
      fTitle.value = "FACTORY SUITE PRO - REPORT PARAMETERS & METADATA";
      fTitle.font = { name: "Segoe UI", size: 12, bold: true, color: { argb: "FFFFFFFF" } };
      fTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F172A" } };
      fTitle.alignment = { horizontal: "center", vertical: "middle" };
      ws3.getRow(1).height = 28;

      const metadata = [
        ["Report Type", "Washing Department Live Report"],
        ["Generated At", new Date().toLocaleString("en-IN")],
        ["Total Lots Count", totalLots],
        ["Total Pieces Count", totalPcs],
        ["Completed Lots", completedLots],
        ["Pending Lots", pendingLots],
        ["Search Keyword", searchTerm || "(None)"],
        ["Status Filter", selectedStatuses.length > 0 ? selectedStatuses.join(", ") : "All Statuses"],
        ["Brand Filter", selectedBrands.length > 0 ? selectedBrands.join(", ") : "All Brands"],
        ["Garment Type Filter", selectedGarments.length > 0 ? selectedGarments.join(", ") : "All Garments"],
        ["Fabric Filter", selectedFabrics.length > 0 ? selectedFabrics.join(", ") : "All Fabrics"],
        ["Season Filter", selectedSeasons.length > 0 ? selectedSeasons.join(", ") : "All Seasons"],
        ["Section Filter", selectedSections.length > 0 ? selectedSections.join(", ") : "All Sections"],
        ["Washing Plant Filter", selectedPlants.length > 0 ? selectedPlants.join(", ") : "All Washing Plants"],
        ["Supervisor Filter", selectedSupervisors.length > 0 ? selectedSupervisors.join(", ") : "All Supervisors"],
        ["Date Filter", dateFilter || "(None)"]
      ];

      metadata.forEach(([k, v]) => {
        const r = ws3.addRow([k, v]);
        r.height = 20;
        r.getCell(1).font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF0F172A" } };
        r.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } };
        r.getCell(1).border = thinBorder;
        r.getCell(2).font = { name: "Segoe UI", size: 9.5, color: { argb: "FF334155" } };
        r.getCell(2).border = thinBorder;
      });

      ws3.columns = [{ width: 25 }, { width: 45 }, { width: 15 }, { width: 15 }];

      // Download file
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      saveAs(blob, `Washing_Department_Report_${new Date().toISOString().slice(0, 10)}.xlsx`);
    } catch (err) {
      console.error("Error generating Washing Report Excel:", err);
      alert("Failed to export Excel file. Please try again.");
    }
  };

  // CSV Export
  const exportToCSV = () => {
    if (filteredData.length === 0) return;
    const headers = ["Timestamp", "Lot Number", "Garment Type", "Fabric", "Style", "Brand", "Season", "Section", "Washing Plant", "Supervisor", "Washing Date", "Total Pcs", "WIP Remarks", "Status", "Completion Date"];
    const rows = filteredData.map((d) => [
      `"${d.timestamp}"`,
      `"${d.lotNumber}"`,
      `"${d.garmentType}"`,
      `"${d.fabric}"`,
      `"${d.style}"`,
      `"${d.brand}"`,
      `"${d.season || ''}"`,
      `"${d.section || ''}"`,
      `"${d.washingPlant || ''}"`,
      `"${d.supervisor}"`,
      `"${d.washingDate}"`,
      d.totalPcs,
      `"${d.wip}"`,
      d.isCompleted ? "Completed" : "Pending",
      `"${d.completeDate || ''}"`
    ]);
    const csvContent = "data:text/csv;charset=utf-8," + [headers.join(","), ...rows.map((r) => r.join(","))].join("\n");
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", `Washing_Department_Report_${new Date().toISOString().slice(0, 10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  // PDF Export Function - 15 columns spacious & centered layout
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
      const isPendingView = selectedStatuses.length === 1 && selectedStatuses[0] === "pending";
      const isCompletedView = selectedStatuses.length === 1 && selectedStatuses[0] === "completed";

      let reportTitle = "FACTORY SUITE PRO - WASHING DEPARTMENT PRODUCTION REPORT";
      if (isPendingView) {
        reportTitle = "FACTORY SUITE PRO - WASHING DEPARTMENT PENDING LOTS REPORT";
      } else if (isCompletedView) {
        reportTitle = "FACTORY SUITE PRO - WASHING DEPARTMENT COMPLETED LOTS REPORT";
      }

      // Aggregations for 4-Column Side-by-Side Executive Summary
      const plantMap = {};
      const garmentMap = {};
      const supervisorMap = {};

      filteredData.forEach((item) => {
        const plant = (item.washingPlant || "Unassigned").trim();
        const garment = (item.garmentType || "Unknown").trim();
        const sup = (item.supervisor || "Unassigned").trim();
        const pcs = item.totalPcs || 0;
        const isComp = item.isCompleted;

        if (!plantMap[plant]) plantMap[plant] = { totalLots: 0, totalPcs: 0, compLots: 0, pendLots: 0 };
        plantMap[plant].totalLots += 1;
        plantMap[plant].totalPcs += pcs;
        if (isComp) plantMap[plant].compLots += 1;
        else plantMap[plant].pendLots += 1;

        if (!garmentMap[garment]) garmentMap[garment] = { totalLots: 0, totalPcs: 0, compLots: 0, pendLots: 0 };
        garmentMap[garment].totalLots += 1;
        garmentMap[garment].totalPcs += pcs;
        if (isComp) garmentMap[garment].compLots += 1;
        else garmentMap[garment].pendLots += 1;

        if (!supervisorMap[sup]) supervisorMap[sup] = { totalLots: 0, totalPcs: 0, compLots: 0, pendLots: 0 };
        supervisorMap[sup].totalLots += 1;
        supervisorMap[sup].totalPcs += pcs;
        if (isComp) supervisorMap[sup].compLots += 1;
        else supervisorMap[sup].pendLots += 1;
      });

      const sortedPlants = Object.keys(plantMap).map(name => ({
        name,
        totalLots: plantMap[name].totalLots,
        totalPcs: plantMap[name].totalPcs,
        compLots: plantMap[name].compLots,
        pendLots: plantMap[name].pendLots
      })).sort((a, b) => b.totalPcs - a.totalPcs);

      const sortedGarments = Object.keys(garmentMap).map(name => ({
        name,
        totalLots: garmentMap[name].totalLots,
        totalPcs: garmentMap[name].totalPcs,
        compLots: garmentMap[name].compLots,
        pendLots: garmentMap[name].pendLots
      })).sort((a, b) => b.totalPcs - a.totalPcs);

      const sortedSupervisors = Object.keys(supervisorMap).map(name => ({
        name,
        totalLots: supervisorMap[name].totalLots,
        totalPcs: supervisorMap[name].totalPcs,
        compLots: supervisorMap[name].compLots,
        pendLots: supervisorMap[name].pendLots
      })).sort((a, b) => b.totalPcs - a.totalPcs);

      // 1. Main Header Block
      doc.setFillColor(15, 23, 42); // Dark Navy #0F172A
      doc.rect(15, 12, pageW - 30, 48, 'F');

      doc.setFontSize(14);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(255, 255, 255);
      doc.text(reportTitle, pageW / 2, 30, { align: 'center' });

      doc.setFontSize(8.5);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(199, 210, 254);
      const subText = `Total Lots: ${totalLots}   |   Total Pieces: ${totalPcs.toLocaleString()}   |   Completed Lots: ${completedLots}   |   Pending Lots: ${pendingLots}   |   Plants: ${sortedPlants.length}   |   Supervisors: ${sortedSupervisors.length}`;
      doc.text(subText, pageW / 2, 48, { align: 'center' });

      // 2. Filter Banner
      doc.setFillColor(241, 245, 249);
      doc.rect(15, 63, pageW - 30, 16, 'F');
      doc.setFontSize(8);
      doc.setFont('helvetica', 'italic');
      doc.setTextColor(0, 0, 0);
      const filterSummary = `Filters: Status: ${selectedStatuses.length ? selectedStatuses.join(', ') : 'All'} | Plants: ${selectedPlants.length ? selectedPlants.join(', ') : 'All'} | Supervisors: ${selectedSupervisors.length ? selectedSupervisors.join(', ') : 'All'} | Garments: ${selectedGarments.length ? selectedGarments.join(', ') : 'All'} | Fabrics: ${selectedFabrics.length ? selectedFabrics.join(', ') : 'All'} | Brands: ${selectedBrands.length ? selectedBrands.join(', ') : 'All'} | Seasons: ${selectedSeasons.length ? selectedSeasons.join(', ') : 'All'} | Search: ${searchTerm || 'None'}`;
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
        'Washing Plant',
        'Washing Date',
        'WIP Remarks',
        'Supervisor',
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
          item.washingPlant || '—',
          item.washingDate || '—',
          isComp ? "Done" : (item.wip || '—'),
          item.supervisor || '—',
          isComp ? 'Completed' : 'In Progress',
          isComp ? (item.completeDate || 'Completed') : '—'
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
        totalPcs.toLocaleString(),
        '',
        '',
        '',
        '',
        `${sortedPlants.length} Plants`,
        '',
        '',
        `${sortedSupervisors.length} Sups`,
        `${completedLots} Comp | ${pendingLots} Pend`,
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
        11: { cellWidth: 75, halign: 'center' },
        12: { cellWidth: 65, halign: 'center' },
        13: { cellWidth: 150, halign: 'center' },
        14: { cellWidth: 75, halign: 'center' },
        15: { cellWidth: 55, halign: 'center' },
        16: { cellWidth: 60, halign: 'center' }
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

            // Completion date styling
            if (data.column.index === 16 && item.isCompleted) {
              data.cell.styles.fillColor = [220, 252, 231];
              data.cell.styles.textColor = [21, 128, 61];
              data.cell.styles.fontStyle = 'bold';
            }
          }
        }
      });

      // --- 4-COLUMN SIDE-BY-SIDE EXECUTIVE SUMMARY ---
      const plantBody = sortedPlants.map(item => {
        const pct = totalPcs > 0 ? ((item.totalPcs / totalPcs) * 100).toFixed(1) : "0.0";
        return [item.name, item.totalLots.toString(), item.totalPcs.toLocaleString(), `${pct}%`];
      });
      plantBody.push(["TOTAL", totalLots.toString(), totalPcs.toLocaleString(), "100.0%"]);

      const gBody = sortedGarments.map(item => {
        const pct = totalPcs > 0 ? ((item.totalPcs / totalPcs) * 100).toFixed(1) : "0.0";
        return [item.name, item.totalLots.toString(), item.totalPcs.toLocaleString(), `${pct}%`];
      });
      gBody.push(["TOTAL", totalLots.toString(), totalPcs.toLocaleString(), "100.0%"]);

      const supBody = sortedSupervisors.map(item => {
        const pct = totalPcs > 0 ? ((item.totalPcs / totalPcs) * 100).toFixed(1) : "0.0";
        return [item.name, item.totalLots.toString(), item.totalPcs.toLocaleString(), `${pct}%`];
      });
      supBody.push(["TOTAL", totalLots.toString(), totalPcs.toLocaleString(), "100.0%"]);

      const compPcs = filteredData.filter(i => i.isCompleted).reduce((s, i) => s + (i.totalPcs || 0), 0);
      const pendPcs = totalPcs - compPcs;
      const statusBody = [
        ["Completed Lots", completedLots.toString(), compPcs.toLocaleString(), `${totalPcs > 0 ? ((compPcs / totalPcs) * 100).toFixed(1) : 0}%`],
        ["Pending Lots", pendingLots.toString(), pendPcs.toLocaleString(), `${totalPcs > 0 ? ((pendPcs / totalPcs) * 100).toFixed(1) : 0}%`],
        ["TOTAL", totalLots.toString(), totalPcs.toLocaleString(), "100.0%"]
      ];

      const maxRows = Math.max(plantBody.length, gBody.length, supBody.length, statusBody.length);
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
      const summarySub = `Total Lots: ${totalLots}   |   Total Pieces: ${totalPcs.toLocaleString()} Pcs   |   Washing Plants: ${sortedPlants.length}   |   Garments: ${sortedGarments.length}   |   Supervisors: ${sortedSupervisors.length}`;
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
      doc.text("1. WASHING PLANT BREAKDOWN", col1X, sectionTitleY);
      doc.text("2. GARMENT BREAKDOWN", col2X, sectionTitleY);
      doc.text("3. SUPERVISOR BREAKDOWN", col3X, sectionTitleY);
      doc.text("4. STATUS & PROGRESS BREAKDOWN", col4X, sectionTitleY);

      const summaryColStyles = {
        0: { cellWidth: 110, halign: 'center' },
        1: { cellWidth: 45, halign: 'center' },
        2: { cellWidth: 68, halign: 'center' },
        3: { cellWidth: 55, halign: 'center' },
      };

      // Column 1 Table: Washing Plant Breakdown
      autoTable(doc, {
        head: [['Washing Plant', 'Lots', 'Total Pcs', 'Share %']],
        body: plantBody,
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
            if (data.row.index === plantBody.length - 1) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });
      const endY1 = doc.lastAutoTable.finalY;

      // Column 2 Table: Garment Breakdown
      autoTable(doc, {
        head: [['Garment Type', 'Lots', 'Total Pcs', 'Share %']],
        body: gBody,
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
            if (data.row.index === gBody.length - 1) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });
      const endY2 = doc.lastAutoTable.finalY;

      // Column 3 Table: Supervisor Breakdown
      autoTable(doc, {
        head: [['Supervisor', 'Lots', 'Total Pcs', 'Share %']],
        body: supBody,
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
            if (data.row.index === supBody.length - 1) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });
      const endY3 = doc.lastAutoTable.finalY;

      // Column 4 Table: Status & Progress Breakdown
      autoTable(doc, {
        head: [['Status Category', 'Lots', 'Total Pcs', 'Share %']],
        body: statusBody,
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
        doc.text("Washing Department Live Report — Factory Suite Pro", 15, finalY + 12);
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
      const fileName = `Washing_Department_Report_${fileDate}.pdf`;
      doc.save(fileName);
    } catch (err) {
      console.error("Error generating Washing Report PDF:", err);
      alert("Failed to generate PDF. Please try again.");
    }
  };

  return (
    <>
      <style>{`
        .washing-report-container {
          padding: 28px 32px 80px;
          background-color: #f8fafc;
          min-height: 100vh;
          font-family: 'Plus Jakarta Sans', 'Inter', system-ui, -apple-system, sans-serif;
          color: #0f172a;
          box-sizing: border-box;
        }
        .report-header-box {
          background: linear-gradient(135deg, #0369a1 0%, #0284c7 100%);
          border-radius: 20px;
          padding: 24px 32px;
          color: white;
          margin-bottom: 24px;
          box-shadow: 0 10px 25px -5px rgba(2, 132, 199, 0.25);
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
        .header-btn-pdf {
          background: #ef4444 !important;
          border-color: #dc2626 !important;
          box-shadow: 0 4px 12px rgba(239, 68, 68, 0.35);
        }
        .header-btn-pdf:hover {
          background: #dc2626 !important;
          box-shadow: 0 6px 16px rgba(220, 38, 38, 0.45);
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
          border-color: #0284c7;
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
          background: #0369a1;
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
          background: rgba(2, 132, 199, 0.1);
          color: #0369a1;
          border: 1px solid #bae6fd;
          padding: 3px 10px;
          border-radius: 8px;
          font-size: 0.75rem;
          font-weight: 700;
        }
      `}</style>

      <div className="washing-report-container">
        {/* Header Box */}
        <div className="report-header-box">
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
              <span style={{ fontSize: "1.8rem" }}>🧼</span>
              <h1 style={{ margin: 0, fontSize: "1.45rem", fontWeight: 800, color: "#ffffff" }}>
                Washing Department Report
              </h1>
              <span style={{ background: "#0ea5e9", color: "white", padding: "2px 8px", borderRadius: "8px", fontSize: "0.72rem", fontWeight: 800 }}>
                LIVE GOOGLE SHEET
              </span>
            </div>
            <p style={{ margin: "6px 0 0 0", color: "#e0f2fe", fontSize: "0.85rem" }}>
              Live production data linked from <strong>OVERLOCK..FOLDING..KAJBUTTON</strong> (`Washing` Tab)
            </p>
          </div>

          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", alignItems: "center" }}>
            <button className="header-btn" onClick={() => history.push("/dashboard")}>
              ← Dashboard
            </button>
            <button className="header-btn" onClick={fetchData}>
              ↻ Refresh
            </button>
            <button
              className="header-btn header-btn-pdf"
              onClick={handleDownloadPDF}
              disabled={filteredData.length === 0}
            >
              📄 Download PDF
            </button>
            <button
              className="header-btn"
              onClick={exportToExcel}
              disabled={filteredData.length === 0}
            >
              📊 Export Excel
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
            <div className="stat-label">Total Washing Lots</div>
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
            <div className="stat-label">Pending In Washing</div>
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
                placeholder="Search Lot, Style, Plant..."
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
              themeColor="#0284c7"
            />

            {/* Brand Filter (Multi-Select) */}
            <MultiSelectDropdown
              label="Brand"
              options={brandOptions}
              selectedValues={selectedBrands}
              onChange={setSelectedBrands}
              themeColor="#0284c7"
            />

            {/* Garment Type Filter (Multi-Select) */}
            <MultiSelectDropdown
              label="Garment Type"
              options={garmentOptions}
              selectedValues={selectedGarments}
              onChange={setSelectedGarments}
              themeColor="#0284c7"
            />

            {/* Fabric Filter (Multi-Select) */}
            <MultiSelectDropdown
              label="Fabric"
              options={fabricOptions}
              selectedValues={selectedFabrics}
              onChange={setSelectedFabrics}
              themeColor="#0284c7"
            />

            {/* Season Filter (Multi-Select) */}
            <MultiSelectDropdown
              label="Season"
              options={seasonOptions}
              selectedValues={selectedSeasons}
              onChange={setSelectedSeasons}
              themeColor="#0284c7"
            />

            {/* Section Filter (Multi-Select) */}
            <MultiSelectDropdown
              label="Section"
              options={sectionOptions}
              selectedValues={selectedSections}
              onChange={setSelectedSections}
              themeColor="#0284c7"
            />

            {/* Washing Plant Filter (Multi-Select) */}
            <MultiSelectDropdown
              label="Washing Plant"
              options={plantOptions}
              selectedValues={selectedPlants}
              onChange={setSelectedPlants}
              themeColor="#0284c7"
            />

            {/* Supervisor Filter (Multi-Select) */}
            <MultiSelectDropdown
              label="Supervisor"
              options={supervisorOptions}
              selectedValues={selectedSupervisors}
              onChange={setSelectedSupervisors}
              themeColor="#0284c7"
            />

            {/* Date Filter */}
            <div>
              <label className="filter-label">Washing Date</label>
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
              {selectedSeasons.map((s) => (
                <span key={s} className="active-tag-pill">
                  Season: {s}
                  <span onClick={() => setSelectedSeasons(selectedSeasons.filter((x) => x !== s))} style={{ cursor: "pointer", fontWeight: 900 }}>×</span>
                </span>
              ))}
              {selectedSections.map((sec) => (
                <span key={sec} className="active-tag-pill">
                  Section: {sec}
                  <span onClick={() => setSelectedSections(selectedSections.filter((x) => x !== sec))} style={{ cursor: "pointer", fontWeight: 900 }}>×</span>
                </span>
              ))}
              {selectedPlants.map((p) => (
                <span key={p} className="active-tag-pill">
                  Plant: {p}
                  <span onClick={() => setSelectedPlants(selectedPlants.filter((x) => x !== p))} style={{ cursor: "pointer", fontWeight: 900 }}>×</span>
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
              <span style={{ fontWeight: 800, fontSize: "1.05rem" }}>Washing Lots & Batches</span>
              <span style={{ background: "#f1f5f9", padding: "3px 10px", borderRadius: "6px", fontSize: "0.8rem", fontWeight: 800, color: "#334155" }}>
                Showing {filteredData.length} of {data.length} lots ({totalPcs.toLocaleString()} pcs)
              </span>
            </div>
          </div>

          <div className="table-container">
            {loading ? (
              <div style={{ textAlign: "center", padding: "60px" }}>
                <div style={{ fontSize: "2rem", marginBottom: "8px" }}>⚡</div>
                <div style={{ fontWeight: 700, color: "#64748b" }}>Loading Washing records from Google Sheet...</div>
              </div>
            ) : filteredData.length === 0 ? (
              <div style={{ textAlign: "center", padding: "60px", color: "#64748b" }}>
                <div style={{ fontSize: "2.5rem", marginBottom: "8px" }}>📭</div>
                <div style={{ fontWeight: 800, fontSize: "1.1rem", color: "#0f172a" }}>No Washing Lots Found</div>
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
                    <th>Garment Type</th>
                    <th>Style</th>
                    <th>Fabric</th>
                    <th>Brand</th>
                    <th>Total Pcs</th>
                    <th>Section</th>
                    <th>Season</th>
                    <th>Party Name</th>
                    <th>Direct Stitching</th>
                    <th>Washing Plant</th>
                    <th>Supervisor</th>
                    <th>Washing Date</th>
                    <th>WIP Remarks</th>
                    <th>Status / Complete</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredData.map((item) => (
                    <tr key={item.id}>
                      <td style={{ fontSize: "0.8rem", color: "#64748b" }}>{item.timestamp}</td>
                      <td>
                        <strong style={{ color: "#0369a1", fontSize: "0.95rem" }}>
                          #{item.lotNumber}
                        </strong>
                      </td>
                      <td>{item.garmentType || "-"}</td>
                      <td>{item.style || "-"}</td>
                      <td>{item.fabric || "-"}</td>
                      <td>{item.brand || "-"}</td>
                      <td>
                        <span style={{ background: "#e0f2fe", color: "#0369a1", padding: "3px 8px", borderRadius: "6px", fontWeight: 800 }}>
                          {item.totalPcs}
                        </span>
                      </td>
                      <td>
                        <span style={{ background: "#f1f5f9", padding: "2px 8px", borderRadius: "6px", fontSize: "0.8rem", fontWeight: 700, color: "#475569" }}>
                          {item.section || "—"}
                        </span>
                      </td>
                      <td>
                        <span style={{ background: "#f1f5f9", padding: "2px 8px", borderRadius: "6px", fontSize: "0.8rem", fontWeight: 700, color: "#475569" }}>
                          {item.season || "—"}
                        </span>
                      </td>
                      <td>{item.partyName || "-"}</td>
                      <td>{item.directStitching || "-"}</td>
                      <td>
                        {item.washingPlant && item.washingPlant !== "—" && item.washingPlant !== "-" ? (
                          <span style={{ background: "#e0f2fe", color: "#0369a1", padding: "2px 8px", borderRadius: "6px", fontSize: "0.8rem", fontWeight: 700, border: "1px solid #bae6fd" }}>
                            {item.washingPlant}
                          </span>
                        ) : (
                          <span style={{ color: "#94a3b8" }}>—</span>
                        )}
                      </td>
                      <td>{item.supervisor || "-"}</td>
                      <td style={{ fontWeight: 700 }}>{item.washingDate || "-"}</td>
                      <td style={{ fontSize: "0.82rem", color: "#64748b" }}>{item.wip}</td>
                      <td>
                        {item.isCompleted ? (
                          <span className="status-badge" style={{ background: "#d1fae5", color: "#047857", border: "1px solid #a7f3d0" }}>
                            ✓ {item.completeDate || "Completed"}
                          </span>
                        ) : (
                          <span className="status-badge" style={{ background: "#fef3c7", color: "#b45309", border: "1px solid #fde68a" }}>
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
