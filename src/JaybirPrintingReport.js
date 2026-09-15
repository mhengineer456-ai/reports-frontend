import React, { useState, useEffect, useMemo, useRef } from "react";
import { useHistory } from "react-router-dom";
import { SPREADSHEET_IDS, fetchSheetDataFromBackend } from "./config";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";

const SPREADSHEET_ID = "1IMhmYlJ3s2PPRgEQs1Ikd4O1OBXK4EYL1oV_-kWAkyg";
const SHEET_NAME = "Jaybir Printing";

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

// Parse various date string formats (YYYY-MM-DD, DD/MM/YYYY, ISO, timestamp)
function parseFlexibleDate(dateStr) {
  if (!dateStr) return null;
  const str = String(dateStr).trim();
  if (!str || str === "-" || str === "—" || str === "[]" || str.toLowerCase() === "pending") return null;

  // Try direct date parse for ISO strings
  const direct = new Date(str);
  if (!isNaN(direct.getTime()) && str.includes("-") && str.split("-")[0].length === 4) {
    return direct;
  }

  // Handle DD/MM/YYYY or MM/DD/YYYY or YYYY-MM-DD
  const cleanStr = str.split(" ")[0];
  const parts = cleanStr.split(/[\/\-\.]/);
  if (parts.length === 3) {
    let day, month, year;
    if (parts[0].length === 4) {
      year = parseInt(parts[0], 10);
      month = parseInt(parts[1], 10) - 1;
      day = parseInt(parts[2], 10);
    } else {
      let p1 = parseInt(parts[0], 10);
      let p2 = parseInt(parts[1], 10);
      year = parseInt(parts[2], 10);
      if (year < 100) year += 2000;

      if (p1 > 12) {
        day = p1;
        month = p2 - 1;
      } else if (p2 > 12) {
        month = p1 - 1;
        day = p2;
      } else {
        // Standard Indian sheet format DD/MM/YYYY
        day = p1;
        month = p2 - 1;
      }
    }

    if (!isNaN(year) && !isNaN(month) && !isNaN(day) && month >= 0 && month <= 11 && day >= 1 && day <= 31) {
      const d = new Date(year, month, day);
      if (!isNaN(d.getTime())) return d;
    }
  }

  if (!isNaN(direct.getTime()) && direct.getFullYear() >= 2000 && direct.getFullYear() <= 2100) {
    return direct;
  }

  return null;
}

// Parse Completion Data from raw JSON string or date string
function parseCompletionData(compVal) {
  if (!compVal || compVal === "[]" || compVal === "-" || compVal === "—" || String(compVal).trim().toLowerCase() === "pending") {
    return {
      isCompleted: false,
      completedDateText: "",
      completedDateObj: null,
      supervisor: "",
      remarks: ""
    };
  }

  const raw = String(compVal).trim();
  let parsedObj = null;

  if (raw.startsWith("[") || raw.startsWith("{")) {
    try {
      const json = JSON.parse(raw);
      if (Array.isArray(json) && json.length > 0) {
        parsedObj = json[json.length - 1];
      } else if (typeof json === "object" && json !== null) {
        parsedObj = json;
      }
    } catch (e) {
      // not JSON
    }
  }

  if (parsedObj) {
    const dateStr = parsedObj.date || "";
    const timeStr = parsedObj.time || "";
    const tsStr = parsedObj.timestamp || "";
    const supervisor = parsedObj.supervisor || "";
    const remarks = parsedObj.remarks || "";

    let dateObj = null;
    if (tsStr) {
      const d = new Date(tsStr);
      if (!isNaN(d.getTime())) dateObj = d;
    }
    if (!dateObj && dateStr) {
      dateObj = parseFlexibleDate(dateStr);
    }

    let formattedDisplay = "";
    if (dateStr && timeStr) {
      formattedDisplay = `${dateStr} ${timeStr}`;
    } else if (dateStr) {
      formattedDisplay = dateStr;
    } else if (dateObj) {
      formattedDisplay = dateObj.toLocaleDateString("en-GB") + (timeStr ? ` ${timeStr}` : "");
    } else {
      formattedDisplay = "Completed";
    }

    return {
      isCompleted: true,
      completedDateText: formattedDisplay,
      completedDateObj: dateObj,
      supervisor,
      remarks,
      rawDate: dateStr,
      rawTime: timeStr
    };
  }

  // Plain string date or status
  const dateObj = parseFlexibleDate(raw);
  return {
    isCompleted: true,
    completedDateText: raw,
    completedDateObj: dateObj,
    supervisor: "",
    remarks: ""
  };
}

// Calculate Aging Days:
// If Pending: Today - Issue Date
// If Completed: Completion Date - Issue Date
function calculateAgingDays(issueDateStr, timestampStr, completedDateObj, isCompleted) {
  const issueDateObj = parseFlexibleDate(issueDateStr) || parseFlexibleDate(timestampStr);
  if (!issueDateObj) return null;

  const issueMidnight = new Date(issueDateObj.getFullYear(), issueDateObj.getMonth(), issueDateObj.getDate()).getTime();

  let targetTime;
  if (isCompleted && completedDateObj) {
    targetTime = new Date(completedDateObj.getFullYear(), completedDateObj.getMonth(), completedDateObj.getDate()).getTime();
  } else {
    const today = new Date();
    targetTime = new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
  }

  const diffMs = targetTime - issueMidnight;
  const days = Math.floor(diffMs / (1000 * 60 * 60 * 24));
  return Math.max(0, days);
}

// Reusable Multi-Select Dropdown Component
function MultiSelectDropdown({ label, options, selectedValues, onChange, themeColor = "#ec4899" }) {
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
          background: selectedValues.length > 0 ? "rgba(236, 72, 153, 0.08)" : "#ffffff",
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
              background: isAllSelected ? "rgba(236, 72, 153, 0.08)" : "transparent",
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
                  background: isChecked ? "rgba(236, 72, 153, 0.08)" : "transparent",
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

export default function JaybirPrintingReport() {
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
  const [selectedSeasons, setSelectedSeasons] = useState([]);
  const [selectedSections, setSelectedSections] = useState([]);
  const [selectedSupervisors, setSelectedSupervisors] = useState([]);
  const [dateFilter, setDateFilter] = useState("");

  const fetchData = async () => {
    setLoading(true);
    setError(null);
    try {
      const [printingRes, jobOrderRes] = await Promise.allSettled([
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
                directStitching: jDirectIdx !== -1 ? String(r[jDirectIdx] || "").trim() : "",
              };
            }
          }
        }
      }

      if (
        printingRes.status !== "fulfilled" ||
        !printingRes.value?.ok ||
        !Array.isArray(printingRes.value.values) ||
        printingRes.value.values.length === 0
      ) {
        setData([]);
        setLoading(false);
        return;
      }

      const rows = printingRes.value.values;
      const headers = rows[0] || [];

      const tsIdx = findCol(headers, ["timestamp", "time"]);
      const lotIdx = findCol(headers, ["lot number", "lot no", "lot"]);
      const garmentIdx = findCol(headers, ["garment type", "garment"]);
      const fabricIdx = findCol(headers, ["fabric"]);
      const styleIdx = findCol(headers, ["style"]);
      const brandIdx = findCol(headers, ["brand", "brand name"]);
      const seasonIdx = findCol(headers, ["season"]);
      const sectionIdx = findCol(headers, ["section"]);
      const supIdx = findCol(headers, ["printing supervisor", "supervisor"]);
      const dateIdx = findCol(headers, ["printing date", "date"]);
      const pcsIdx = findCol(headers, ["total pcs", "pcs", "quantity"]);
      const wipIdx = findCol(headers, ["wip jaybir printing", "wip", "remarks"]);
      const compIdx = findCol(headers, ["jaybir printing complete", "complete", "completed"]);

      const parsed = [];
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const lotNo = String(row[lotIdx !== -1 ? lotIdx : 1] || "").trim();
        if (!lotNo) continue;

        const pcs = parseInt(row[pcsIdx !== -1 ? pcsIdx : 8], 10) || 0;
        const compVal = String(row[compIdx !== -1 ? compIdx : 10] || "").trim();
        const wipVal = String(row[wipIdx !== -1 ? wipIdx : 9] || "").trim();
        const printDate = row[dateIdx !== -1 ? dateIdx : 7] || "";
        const ts = row[tsIdx !== -1 ? tsIdx : 0] || "";

        // Parse completion JSON / Date
        const compData = parseCompletionData(compVal);
        const agingDays = calculateAgingDays(printDate, ts, compData.completedDateObj, compData.isCompleted);

        const normLot = lotNo.toLowerCase();
        const jobInfo = lotToJobInfo[normLot] || {};
        const sheetBrand = brandIdx !== -1 ? String(row[brandIdx] || "").trim() : "";
        const finalBrand = sheetBrand && sheetBrand !== "-" ? sheetBrand : (jobInfo.brand || jobInfo.party || "—");
        const sheetSeason = seasonIdx !== -1 ? String(row[seasonIdx] || "").trim() : "";
        const finalSeason = sheetSeason && sheetSeason !== "-" ? sheetSeason : (jobInfo.season || "—");
        const sheetSection = sectionIdx !== -1 ? String(row[sectionIdx] || "").trim() : "";
        const finalSection = sheetSection && sheetSection !== "-" ? sheetSection : (jobInfo.section || "—");

        const isDone = compData.isCompleted || (wipVal.toLowerCase().includes("done") || wipVal.toLowerCase().includes("complete"));
        const finalWip = isDone
          ? "Done"
          : (wipVal === "[]" ? "In Progress" : wipVal || "In Progress");

        parsed.push({
          id: i,
          timestamp: ts,
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
          supervisor: row[supIdx !== -1 ? supIdx : 6] || "Jaybir Printing",
          printingDate: printDate,
          wip: finalWip,
          completeDate: compData.completedDateText,
          completedDateObj: compData.completedDateObj,
          completionSupervisor: compData.supervisor,
          completionRemarks: compData.remarks,
          isCompleted: compData.isCompleted,
          agingDays: agingDays
        });
      }

      setData(parsed);
    } catch (err) {
      console.error("Error loading Jaybir Printing data:", err);
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
        item.garmentType.toLowerCase().includes(term) ||
        item.supervisor.toLowerCase().includes(term) ||
        item.wip.toLowerCase().includes(term) ||
        (item.completeDate && item.completeDate.toLowerCase().includes(term));

      const matchesStatus =
        selectedStatuses.length === 0 ||
        (selectedStatuses.includes("completed") && item.isCompleted) ||
        (selectedStatuses.includes("pending") && !item.isCompleted);

      const matchesBrand = selectedBrands.length === 0 || selectedBrands.some((b) => b.toLowerCase() === item.brand.toLowerCase());
      const matchesGarment = selectedGarments.length === 0 || selectedGarments.some((g) => g.toLowerCase() === item.garmentType.toLowerCase());
      const matchesFabric = selectedFabrics.length === 0 || selectedFabrics.some((f) => f.toLowerCase() === item.fabric.toLowerCase());
      const matchesSeason = selectedSeasons.length === 0 || selectedSeasons.some((s) => s.toLowerCase() === (item.season || "").toLowerCase());
      const matchesSection = selectedSections.length === 0 || selectedSections.some((s) => s.toLowerCase() === (item.section || "").toLowerCase());
      const matchesSupervisor = selectedSupervisors.length === 0 || selectedSupervisors.some((s) => s.toLowerCase() === item.supervisor.toLowerCase());
      const matchesDate = !dateFilter || item.printingDate.includes(dateFilter);

      return matchesSearch && matchesStatus && matchesBrand && matchesGarment && matchesFabric && matchesSeason && matchesSection && matchesSupervisor && matchesDate;
    });
  }, [data, searchTerm, selectedStatuses, selectedBrands, selectedGarments, selectedFabrics, selectedSeasons, selectedSections, selectedSupervisors, dateFilter]);

  const resetFilters = () => {
    setSearchTerm("");
    setSelectedStatuses([]);
    setSelectedBrands([]);
    setSelectedGarments([]);
    setSelectedFabrics([]);
    setSelectedSeasons([]);
    setSelectedSections([]);
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
    selectedSupervisors.length > 0 ||
    dateFilter !== "";

  // Summary Metrics for filtered data
  const totalLots = filteredData.length;
  const totalPcs = filteredData.reduce((sum, item) => sum + item.totalPcs, 0);
  const completedLots = filteredData.filter((item) => item.isCompleted).length;
  const pendingLots = totalLots - completedLots;
  const pendingPcs = filteredData.filter((item) => !item.isCompleted).reduce((sum, item) => sum + (item.totalPcs || 0), 0);
  const validAgingList = filteredData.map((d) => d.agingDays).filter((a) => a !== null && !isNaN(a));
  const avgAging = validAgingList.length > 0 ? Math.round(validAgingList.reduce((a, b) => a + b, 0) / validAgingList.length) : 0;

  // Excel (.xlsx) Multi-Sheet Export with ExcelJS (Matching PendingIssue.js Structure)
  const exportToExcel = async () => {
    if (filteredData.length === 0) {
      alert("No data available to export.");
      return;
    }

    try {
      const workbook = new ExcelJS.Workbook();
      workbook.creator = "Factory Suite Pro";
      workbook.created = new Date();

      const totalPieces = filteredData.reduce((sum, item) => sum + (item.totalPcs || 0), 0);
      const totalLots = filteredData.length;
      const completedCount = filteredData.filter((d) => d.isCompleted).length;
      const pendingCount = totalLots - completedCount;

      // Grouping for Summary
      const garmentMap = {};
      const seasonMap = {};
      const brandMap = {};
      let normalLots = 0, normalPcs = 0;
      let urgentLots = 0, urgentPcs = 0;
      let criticalLots = 0, criticalPcs = 0;

      filteredData.forEach((d) => {
        const pcs = d.totalPcs || 0;
        const gType = (d.garmentType || "Unknown").trim();
        const season = (d.season || "Other / NA").trim();
        const brand = (d.brand || "Unknown").trim();
        const days = typeof d.agingDays === "number" ? d.agingDays : 0;

        if (!garmentMap[gType]) garmentMap[gType] = { totalLots: 0, totalPcs: 0 };
        garmentMap[gType].totalLots += 1;
        garmentMap[gType].totalPcs += pcs;

        if (!seasonMap[season]) seasonMap[season] = { totalLots: 0, totalPcs: 0 };
        seasonMap[season].totalLots += 1;
        seasonMap[season].totalPcs += pcs;

        if (!brandMap[brand]) brandMap[brand] = { totalLots: 0, totalPcs: 0 };
        brandMap[brand].totalLots += 1;
        brandMap[brand].totalPcs += pcs;

        if (days <= 7) {
          normalLots += 1;
          normalPcs += pcs;
        } else if (days <= 14) {
          urgentLots += 1;
          urgentPcs += pcs;
        } else {
          criticalLots += 1;
          criticalPcs += pcs;
        }
      });

      const sortedGarments = Object.keys(garmentMap)
        .map((name) => ({ name, totalLots: garmentMap[name].totalLots, totalPcs: garmentMap[name].totalPcs }))
        .sort((a, b) => b.totalPcs - a.totalPcs);

      const sortedSeasons = Object.keys(seasonMap)
        .map((name) => ({ name, totalLots: seasonMap[name].totalLots, totalPcs: seasonMap[name].totalPcs }))
        .sort((a, b) => b.totalPcs - a.totalPcs);

      const sortedBrands = Object.keys(brandMap)
        .map((name) => ({ name, totalLots: brandMap[name].totalLots, totalPcs: brandMap[name].totalPcs }))
        .sort((a, b) => b.totalPcs - a.totalPcs);

      const thinBorder = {
        top: { style: "thin", color: { argb: "FFCBD5E1" } },
        bottom: { style: "thin", color: { argb: "FFCBD5E1" } },
        left: { style: "thin", color: { argb: "FFCBD5E1" } },
        right: { style: "thin", color: { argb: "FFCBD5E1" } },
      };

      // ================= SHEET 1: MASTER PRINTING LOTS =================
      const ws1 = workbook.addWorksheet("Printing Lots", {
        views: [{ showGridLines: true }],
      });

      // Banner Row 1: Title
      ws1.mergeCells("A1:Q1");
      const titleCell = ws1.getCell("A1");
      titleCell.value = "FACTORY SUITE PRO - JAYBIR PRINTING REPORT";
      titleCell.font = { name: "Segoe UI", size: 14, bold: true, color: { argb: "FFFFFFFF" } };
      titleCell.alignment = { horizontal: "center", vertical: "middle" };
      titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F172A" } };
      ws1.getRow(1).height = 32;

      // Banner Row 2: Subtitle & KPI
      ws1.mergeCells("A2:Q2");
      const subCell = ws1.getCell("A2");
      subCell.value = `Total Lots: ${totalLots}   |   Total Pieces: ${totalPieces.toLocaleString()}   |   Completed: ${completedCount}   |   Pending: ${pendingCount}   |   Normal (<=7d): ${normalLots}   |   Urgent (8-14d): ${urgentLots}   |   Critical (>14d): ${criticalLots}`;
      subCell.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFC7D2FE" } };
      subCell.alignment = { horizontal: "center", vertical: "middle" };
      subCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E293B" } };
      ws1.getRow(2).height = 24;

      // Banner Row 3: Filter Info
      ws1.mergeCells("A3:Q3");
      const filterCell = ws1.getCell("A3");
      const filterSummary = `Brand: ${selectedBrands.length ? selectedBrands.join(", ") : "All"} | Garment: ${selectedGarments.length ? selectedGarments.join(", ") : "All"} | Fabric: ${selectedFabrics.length ? selectedFabrics.join(", ") : "All"} | Season: ${selectedSeasons.length ? selectedSeasons.join(", ") : "All"} | Section: ${selectedSections.length ? selectedSections.join(", ") : "All"} | Supervisor: ${selectedSupervisors.length ? selectedSupervisors.join(", ") : "All"} | Status: ${selectedStatuses.length ? selectedStatuses.join(", ") : "All"} | Date: ${dateFilter || "All"}`;
      filterCell.value = `Applied Filters: ${filterSummary}`;
      filterCell.font = { name: "Segoe UI", size: 8.5, italic: true, color: { argb: "FF475569" } };
      filterCell.alignment = { horizontal: "center", vertical: "middle" };
      filterCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } };
      ws1.getRow(3).height = 20;

      // Row 4: Spacer
      ws1.getRow(4).height = 10;

      // Table Headers (Row 5)
      const headers1 = [
        "#",
        "Timestamp",
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
        "Supervisor",
        "Printing Date",
        "Aging (Days)",
        "WIP Remarks",
        "Status",
        "Completion Date",
      ];
      const headerRow1 = ws1.addRow(headers1);
      headerRow1.height = 28;
      headerRow1.eachCell((cell) => {
        cell.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFFFFFFF" } };
        cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E293B" } };
        cell.border = {
          top: { style: "medium", color: { argb: "FF0F172A" } },
          bottom: { style: "medium", color: { argb: "FF0F172A" } },
          left: { style: "thin", color: { argb: "FF334155" } },
          right: { style: "thin", color: { argb: "FF334155" } },
        };
      });

      // Data Rows
      filteredData.forEach((item, idx) => {
        const days = typeof item.agingDays === "number" ? item.agingDays : null;
        const rowData = [
          idx + 1,
          item.timestamp || "—",
          item.lotNumber,
          item.garmentType || "—",
          item.style || "—",
          item.fabric || "—",
          item.brand || "—",
          item.totalPcs || 0,
          item.section || "—",
          item.season || "—",
          item.partyName || "—",
          item.directStitching || "—",
          item.supervisor || "—",
          item.printingDate || "—",
          days !== null ? `${days} Days` : "—",
          item.isCompleted ? "Done" : (item.wip || "In Progress"),
          item.isCompleted ? "Completed" : "In Progress",
          item.isCompleted ? (item.completeDate || "Completed") : "—",
        ];

        const r = ws1.addRow(rowData);
        r.height = 20;

        for (let col = 1; col <= 18; col++) {
          r.getCell(col).alignment = { horizontal: "center", vertical: "middle" };
          r.getCell(col).font = { name: "Segoe UI", size: 9 };
          r.getCell(col).border = thinBorder;
          if (idx % 2 === 1) {
            r.getCell(col).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
          }
        }

        r.getCell(8).numFmt = "#,##0";
        r.getCell(3).font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FF5B21B6" } };

        // Days Pending badge color
        if (days !== null) {
          if (days > 14) {
            r.getCell(15).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEE2E2" } };
            r.getCell(15).font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FFDC2626" } };
          } else if (days > 7) {
            r.getCell(15).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEF3C7" } };
            r.getCell(15).font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FFB45309" } };
          } else {
            r.getCell(15).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDCFCE7" } };
            r.getCell(15).font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FF15803D" } };
          }
        }

        // Status badge color
        if (item.isCompleted) {
          r.getCell(17).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDCFCE7" } };
          r.getCell(17).font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FF15803D" } };
        } else {
          r.getCell(17).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEF3C7" } };
          r.getCell(17).font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FFB45309" } };
        }
      });

      // Total Row
      const totalRow1 = ws1.addRow([
        "",
        `TOTAL (${totalLots} Lots)`,
        "",
        "",
        "",
        "",
        "",
        totalPieces,
        "",
        "",
        "",
        "",
        "",
        "",
        `Avg: ${avgAging}d`,
        "",
        `${completedCount} Done | ${pendingCount} Pending`,
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
      });
      totalRow1.getCell(2).alignment = { horizontal: "center", vertical: "middle" };
      totalRow1.getCell(8).alignment = { horizontal: "center", vertical: "middle" };
      totalRow1.getCell(8).numFmt = "#,##0";
      totalRow1.getCell(15).alignment = { horizontal: "center", vertical: "middle" };
      totalRow1.getCell(17).alignment = { horizontal: "center", vertical: "middle" };

      // Set column widths
      const colWidths1 = [6, 18, 16, 18, 22, 22, 18, 14, 14, 14, 18, 16, 18, 16, 14, 16, 16, 22];
      colWidths1.forEach((w, i) => {
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
        const pct = totalPieces > 0 ? item.totalPcs / totalPieces : 0;
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

      const gTotalRow = ws2.addRow(["TOTAL", totalLots, totalPieces, 1]);
      gTotalRow.height = 22;
      gTotalRow.eachCell((c) => {
        c.font = { name: "Segoe UI", size: 9.5, bold: true };
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE2E8F0" } };
        c.border = { top: { style: "thin" }, bottom: { style: "double" }, left: { style: "thin" }, right: { style: "thin" } };
      });
      gTotalRow.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
      gTotalRow.getCell(2).alignment = { horizontal: "center", vertical: "middle" };
      gTotalRow.getCell(3).alignment = { horizontal: "center", vertical: "middle" };
      gTotalRow.getCell(3).numFmt = "#,##0";
      gTotalRow.getCell(4).alignment = { horizontal: "center", vertical: "middle" };
      gTotalRow.getCell(4).numFmt = "0.0%";

      // Spacer
      ws2.addRow([]);

      // Section 2: Season Breakdown
      const seasonStartRow = ws2.rowCount + 1;
      ws2.mergeCells(`A${seasonStartRow}:D${seasonStartRow}`);
      const sTitle = ws2.getCell(`A${seasonStartRow}`);
      sTitle.value = "2. SEASON WISE BREAKDOWN (LOTS & PIECES DISTRIBUTION)";
      sTitle.font = { name: "Segoe UI", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      sTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4338CA" } };
      sTitle.alignment = { horizontal: "left", vertical: "middle" };
      ws2.getRow(seasonStartRow).height = 26;

      const sHeader = ws2.addRow(["Season", "Total Lots", "Total Pieces (Qty)", "Share %"]);
      sHeader.height = 22;
      sHeader.eachCell((c) => {
        c.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFFFFFFF" } };
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF312E81" } };
        c.alignment = { horizontal: "center", vertical: "middle" };
        c.border = thinBorder;
      });

      sortedSeasons.forEach((item, idx) => {
        const pct = totalPieces > 0 ? item.totalPcs / totalPieces : 0;
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

      const sTotalRow = ws2.addRow(["TOTAL", totalLots, totalPieces, 1]);
      sTotalRow.height = 22;
      sTotalRow.eachCell((c) => {
        c.font = { name: "Segoe UI", size: 9.5, bold: true };
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE2E8F0" } };
        c.border = { top: { style: "thin" }, bottom: { style: "double" }, left: { style: "thin" }, right: { style: "thin" } };
      });
      sTotalRow.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
      sTotalRow.getCell(2).alignment = { horizontal: "center", vertical: "middle" };
      sTotalRow.getCell(3).alignment = { horizontal: "center", vertical: "middle" };
      sTotalRow.getCell(3).numFmt = "#,##0";
      sTotalRow.getCell(4).alignment = { horizontal: "center", vertical: "middle" };
      sTotalRow.getCell(4).numFmt = "0.0%";

      // Spacer
      ws2.addRow([]);

      // Section 3: Brand Summary
      const brandStartRow = ws2.rowCount + 1;
      ws2.mergeCells(`A${brandStartRow}:D${brandStartRow}`);
      const pTitle = ws2.getCell(`A${brandStartRow}`);
      pTitle.value = "3. BRAND SUMMARY & DISTRIBUTION";
      pTitle.font = { name: "Segoe UI", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      pTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E40AF" } };
      pTitle.alignment = { horizontal: "left", vertical: "middle" };
      ws2.getRow(brandStartRow).height = 26;

      const pHeader = ws2.addRow(["Brand Name", "Total Lots", "Total Pieces (Qty)", "Share %"]);
      pHeader.height = 22;
      pHeader.eachCell((c) => {
        c.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFFFFFFF" } };
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E3A8A" } };
        c.alignment = { horizontal: "center", vertical: "middle" };
        c.border = thinBorder;
      });

      sortedBrands.forEach((item, idx) => {
        const pct = totalPieces > 0 ? item.totalPcs / totalPieces : 0;
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

      const pTotalRow = ws2.addRow(["TOTAL", totalLots, totalPieces, 1]);
      pTotalRow.height = 22;
      pTotalRow.eachCell((c) => {
        c.font = { name: "Segoe UI", size: 9.5, bold: true };
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE2E8F0" } };
        c.border = { top: { style: "thin" }, bottom: { style: "double" }, left: { style: "thin" }, right: { style: "thin" } };
      });
      pTotalRow.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
      pTotalRow.getCell(2).alignment = { horizontal: "center", vertical: "middle" };
      pTotalRow.getCell(3).alignment = { horizontal: "center", vertical: "middle" };
      pTotalRow.getCell(3).numFmt = "#,##0";
      pTotalRow.getCell(4).alignment = { horizontal: "center", vertical: "middle" };
      pTotalRow.getCell(4).numFmt = "0.0%";

      // Spacer
      ws2.addRow([]);

      // Section 4: SLA & Aging Breakdown
      const slaStartRow = ws2.rowCount + 1;
      ws2.mergeCells(`A${slaStartRow}:D${slaStartRow}`);
      const slaTitle = ws2.getCell(`A${slaStartRow}`);
      slaTitle.value = "4. AGING & SLA BREAKDOWN";
      slaTitle.font = { name: "Segoe UI", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      slaTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFB45309" } };
      slaTitle.alignment = { horizontal: "left", vertical: "middle" };
      ws2.getRow(slaStartRow).height = 26;

      const slaHeader = ws2.addRow(["Aging Status", "Total Lots", "Total Pieces (Qty)", "Share %"]);
      slaHeader.height = 22;
      slaHeader.eachCell((c) => {
        c.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFFFFFFF" } };
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF78350F" } };
        c.alignment = { horizontal: "center", vertical: "middle" };
        c.border = thinBorder;
      });

      const slaData = [
        { name: "<= 7 Days (On-Time / Normal)", lots: normalLots, pcs: normalPcs, bg: "FFDCFCE7", fg: "FF15803D" },
        { name: "8 - 14 Days (Urgent Zone)", lots: urgentLots, pcs: urgentPcs, bg: "FFFEF3C7", fg: "FFB45309" },
        { name: "> 14 Days (Critical / Delayed)", lots: criticalLots, pcs: criticalPcs, bg: "FFFEE2E2", fg: "FFDC2626" },
      ];

      slaData.forEach((item) => {
        const pct = totalPieces > 0 ? item.pcs / totalPieces : 0;
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

      const slaTotalRow = ws2.addRow(["TOTAL", totalLots, totalPieces, 1]);
      slaTotalRow.height = 22;
      slaTotalRow.eachCell((c) => {
        c.font = { name: "Segoe UI", size: 9.5, bold: true };
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE2E8F0" } };
        c.border = { top: { style: "thin" }, bottom: { style: "double" }, left: { style: "thin" }, right: { style: "thin" } };
      });
      slaTotalRow.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
      slaTotalRow.getCell(2).alignment = { horizontal: "center", vertical: "middle" };
      slaTotalRow.getCell(3).alignment = { horizontal: "center", vertical: "middle" };
      slaTotalRow.getCell(3).numFmt = "#,##0";
      slaTotalRow.getCell(4).alignment = { horizontal: "center", vertical: "middle" };
      slaTotalRow.getCell(4).numFmt = "0.0%";

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
        ["Department", "Jaybir Printing"],
        ["Total Lots Exported", totalLots],
        ["Total Pieces Exported", totalPieces.toLocaleString()],
        ["Completed Lots", completedCount],
        ["Pending Lots", pendingCount],
        ["Average Aging", `${avgAging} Days`],
        ["Brand Filter", selectedBrands.length ? selectedBrands.join(", ") : "All Brands"],
        ["Garment Type Filter", selectedGarments.length ? selectedGarments.join(", ") : "All Garments"],
        ["Fabric Filter", selectedFabrics.length ? selectedFabrics.join(", ") : "All Fabrics"],
        ["Season Filter", selectedSeasons.length ? selectedSeasons.join(", ") : "All Seasons"],
        ["Section Filter", selectedSections.length ? selectedSections.join(", ") : "All Sections"],
        ["Supervisor Filter", selectedSupervisors.length ? selectedSupervisors.join(", ") : "All Supervisors"],
        ["Status Filter", selectedStatuses.length ? selectedStatuses.join(", ") : "All Statuses"],
        ["Printing Date Filter", dateFilter || "All Dates"],
        ["Search Keyword", searchTerm || "None"],
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

      // Save using file-saver
      const buffer = await workbook.xlsx.writeBuffer();
      const ts = new Date().toISOString().slice(0, 10);
      saveAs(new Blob([buffer]), `Jaybir_Printing_Report_${ts}.xlsx`);
    } catch (e) {
      console.error("Excel export error:", e);
      alert(`Excel export failed: ${e.message}`);
    }
  };

  // CSV Export
  const exportToCSV = () => {
    if (filteredData.length === 0) return;
    const headers = ["Sr No.", "Timestamp", "Lot Number", "Garment Type", "Fabric", "Style", "Brand", "Season", "Section", "Party Name", "Direct Stitching", "Supervisor", "Printing Date", "Total Pcs", "Aging (Days)", "WIP Remarks", "Status", "Completion Date"];
    const rows = filteredData.map((d, index) => [
      index + 1,
      `"${d.timestamp}"`,
      `"${d.lotNumber}"`,
      `"${d.garmentType}"`,
      `"${d.fabric}"`,
      `"${d.style}"`,
      `"${d.brand}"`,
      `"${d.season || ''}"`,
      `"${d.section || ''}"`,
      `"${d.partyName || ''}"`,
      `"${d.directStitching || ''}"`,
      `"${d.supervisor}"`,
      `"${d.printingDate}"`,
      d.totalPcs,
      d.agingDays !== null ? d.agingDays : "",
      `"${d.wip}"`,
      d.isCompleted ? "Completed" : "Pending",
      `"${d.isCompleted ? (d.completeDate || 'Completed') : ''}"`
    ]);
    const csvContent = "data:text/csv;charset=utf-8," + [headers.join(","), ...rows.map((r) => r.join(","))].join("\n");
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", `Jaybir_Printing_Report_${new Date().toISOString().slice(0, 10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  // Professional A3 Landscape PDF Export (Matching PendingIssue.js Structure)
  const handleDownloadPDF = () => {
    if (filteredData.length === 0) {
      alert("No data available to download PDF.");
      return;
    }

    try {
      const doc = new jsPDF({
        orientation: "landscape",
        unit: "pt",
        format: "a3",
      });

      const pageW = doc.internal.pageSize.getWidth();
      const pageH = doc.internal.pageSize.getHeight();
      const tableWidth = 1140;
      const startX = (pageW - tableWidth) / 2;

      const totalPieces = filteredData.reduce((sum, item) => sum + (item.totalPcs || 0), 0);
      const totalLots = filteredData.length;
      const completedCount = filteredData.filter((d) => d.isCompleted).length;
      const pendingCount = totalLots - completedCount;

      // Grouping for Executive Summary
      const garmentMap = {};
      const seasonMap = {};
      const brandMap = {};
      let normalLots = 0, normalPcs = 0;
      let urgentLots = 0, urgentPcs = 0;
      let criticalLots = 0, criticalPcs = 0;

      filteredData.forEach((d) => {
        const pcs = d.totalPcs || 0;
        const gType = (d.garmentType || "Unknown").trim();
        const season = (d.season || "Other / NA").trim();
        const brand = (d.brand || "Unknown").trim();
        const days = typeof d.agingDays === "number" ? d.agingDays : 0;

        if (!garmentMap[gType]) garmentMap[gType] = { totalLots: 0, totalPcs: 0 };
        garmentMap[gType].totalLots += 1;
        garmentMap[gType].totalPcs += pcs;

        if (!seasonMap[season]) seasonMap[season] = { totalLots: 0, totalPcs: 0 };
        seasonMap[season].totalLots += 1;
        seasonMap[season].totalPcs += pcs;

        if (!brandMap[brand]) brandMap[brand] = { totalLots: 0, totalPcs: 0 };
        brandMap[brand].totalLots += 1;
        brandMap[brand].totalPcs += pcs;

        if (days <= 7) {
          normalLots += 1;
          normalPcs += pcs;
        } else if (days <= 14) {
          urgentLots += 1;
          urgentPcs += pcs;
        } else {
          criticalLots += 1;
          criticalPcs += pcs;
        }
      });

      const sortedGarments = Object.keys(garmentMap)
        .map((name) => ({ name, totalLots: garmentMap[name].totalLots, totalPcs: garmentMap[name].totalPcs }))
        .sort((a, b) => b.totalPcs - a.totalPcs);

      const sortedSeasons = Object.keys(seasonMap)
        .map((name) => ({ name, totalLots: seasonMap[name].totalLots, totalPcs: seasonMap[name].totalPcs }))
        .sort((a, b) => b.totalPcs - a.totalPcs);

      const sortedBrands = Object.keys(brandMap)
        .map((name) => ({ name, totalLots: brandMap[name].totalLots, totalPcs: brandMap[name].totalPcs }))
        .sort((a, b) => b.totalPcs - a.totalPcs);

      // Main Header Block (Centered)
      doc.setFillColor(15, 23, 42); // Dark Navy
      doc.rect(startX, 12, tableWidth, 48, "F");

      doc.setFontSize(14);
      doc.setFont("helvetica", "bold");
      doc.setTextColor(255, 255, 255);
      doc.text("FACTORY SUITE PRO - JAYBIR PRINTING REPORT", pageW / 2, 30, { align: "center" });

      doc.setFontSize(8.5);
      doc.setFont("helvetica", "normal");
      doc.setTextColor(199, 210, 254);
      const subText = `Total Lots: ${totalLots}   |   Total Pieces: ${totalPieces.toLocaleString()}   |   Completed: ${completedCount}   |   Pending: ${pendingCount}   |   Avg Aging: ${avgAging} Days   |   Normal (<=7d): ${normalLots}   |   Urgent (8-14d): ${urgentLots}   |   Critical (>14d): ${criticalLots}`;
      doc.text(subText, pageW / 2, 48, { align: "center" });

      // Filter Banner (Centered)
      doc.setFillColor(241, 245, 249);
      doc.rect(startX, 63, tableWidth, 16, "F");
      doc.setFontSize(8);
      doc.setFont("helvetica", "italic");
      doc.setTextColor(0, 0, 0);
      const filterSummary = `Filters: Brand: ${selectedBrands.length ? selectedBrands.join(", ") : "All"} | Garment: ${selectedGarments.length ? selectedGarments.join(", ") : "All"} | Fabric: ${selectedFabrics.length ? selectedFabrics.join(", ") : "All"} | Season: ${selectedSeasons.length ? selectedSeasons.join(", ") : "All"} | Section: ${selectedSections.length ? selectedSections.join(", ") : "All"} | Supervisor: ${selectedSupervisors.length ? selectedSupervisors.join(", ") : "All"} | Status: ${selectedStatuses.length ? selectedStatuses.join(", ") : "All"} | Date: ${dateFilter || "All"}`;
      doc.text(filterSummary, pageW / 2, 74, { align: "center" });

      // Table columns & rows
      const tableColumns = [
        "#",
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
        "Supervisor",
        "Printing Date",
        "Aging",
        "WIP Remarks",
        "Status",
        "Completed Date",
      ];

      const tableBody = filteredData.map((item, idx) => {
        const isComp = item.isCompleted;
        const days = typeof item.agingDays === "number" ? item.agingDays : null;
        return [
          (idx + 1).toString(),
          item.lotNumber || "—",
          item.garmentType || "—",
          item.style || "—",
          item.fabric || "—",
          item.brand || "—",
          (item.totalPcs || 0).toLocaleString(),
          item.section || "—",
          item.season || "—",
          item.partyName || "—",
          item.directStitching || "—",
          item.supervisor || "—",
          item.printingDate || "—",
          days !== null ? `${days} Days` : "—",
          isComp ? "Done" : (item.wip || "In Progress"),
          isComp ? "Completed" : "In Progress",
          isComp ? (item.completeDate || "Completed") : "—",
        ];
      });

      // Add Total Row
      tableBody.push([
        "",
        `TOTAL (${totalLots})`,
        "",
        "",
        "",
        "",
        totalPieces.toLocaleString(),
        "",
        "",
        "",
        "",
        "",
        "",
        `Avg: ${avgAging}d`,
        "",
        `${completedCount} Done | ${pendingCount} Pending`,
        "",
      ]);

      const columnStyles = {
        0: { cellWidth: 25, halign: "center" },
        1: { cellWidth: 65, halign: "center" },
        2: { cellWidth: 75, halign: "center" },
        3: { cellWidth: 80, halign: "center" },
        4: { cellWidth: 80, halign: "center" },
        5: { cellWidth: 70, halign: "center" },
        6: { cellWidth: 55, halign: "center", fontStyle: "bold" },
        7: { cellWidth: 50, halign: "center" },
        8: { cellWidth: 50, halign: "center" },
        9: { cellWidth: 80, halign: "center" },
        10: { cellWidth: 55, halign: "center" },
        11: { cellWidth: 75, halign: "center" },
        12: { cellWidth: 70, halign: "center" },
        13: { cellWidth: 55, halign: "center" },
        14: { cellWidth: 65, halign: "center" },
        15: { cellWidth: 65, halign: "center" },
        16: { cellWidth: 80, halign: "center" },
      };

      autoTable(doc, {
        head: [tableColumns],
        body: tableBody,
        startY: 85,
        tableWidth: tableWidth,
        margin: { top: 85, right: startX, bottom: 25, left: startX },
        theme: "grid",
        styles: {
          fontSize: 8.5,
          cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
          overflow: "linebreak",
          valign: "middle",
          halign: "center",
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          lineWidth: 0.3,
          fontStyle: "normal",
          minCellHeight: 12,
        },
        headStyles: {
          fillColor: [15, 23, 42],
          textColor: [255, 255, 255],
          fontStyle: "bold",
          lineColor: [0, 0, 0],
          lineWidth: 0.5,
          halign: "center",
          fontSize: 9,
          valign: "middle",
          cellPadding: { top: 5, right: 3, bottom: 5, left: 3 },
        },
        alternateRowStyles: {
          fillColor: [248, 250, 252],
        },
        columnStyles,
        didParseCell: function (data) {
          if (data.section === "body") {
            const rowIndex = data.row.index;
            const isTotalRow = rowIndex === tableBody.length - 1;

            if (isTotalRow) {
              data.cell.styles.fontStyle = "bold";
              data.cell.styles.fillColor = [226, 232, 240];
              data.cell.styles.textColor = [0, 0, 0];
              data.cell.styles.halign = "center";
              return;
            }

            const item = filteredData[rowIndex];
            if (!item) return;

            // Lot number styling
            if (data.column.index === 1) {
              data.cell.styles.textColor = [91, 33, 182];
              data.cell.styles.fontStyle = "bold";
            }

            // Days pending / Aging styling
            if (data.column.index === 13) {
              const days = typeof item.agingDays === "number" ? item.agingDays : null;
              if (days !== null) {
                if (days > 14) {
                  data.cell.styles.fillColor = [239, 68, 68];
                  data.cell.styles.textColor = [255, 255, 255];
                  data.cell.styles.fontStyle = "bold";
                } else if (days > 7) {
                  data.cell.styles.fillColor = [254, 243, 199];
                  data.cell.styles.textColor = [180, 83, 9];
                  data.cell.styles.fontStyle = "bold";
                } else {
                  data.cell.styles.fillColor = [220, 252, 231];
                  data.cell.styles.textColor = [21, 128, 61];
                  data.cell.styles.fontStyle = "bold";
                }
              }
            }

            // WIP styling
            if (data.column.index === 14) {
              if (item.isCompleted) {
                data.cell.styles.fillColor = [220, 252, 231];
                data.cell.styles.textColor = [21, 128, 61];
                data.cell.styles.fontStyle = "bold";
              }
            }

            // Status styling
            if (data.column.index === 15) {
              if (item.isCompleted) {
                data.cell.styles.fillColor = [220, 252, 231];
                data.cell.styles.textColor = [21, 128, 61];
                data.cell.styles.fontStyle = "bold";
              } else {
                data.cell.styles.fillColor = [254, 243, 199];
                data.cell.styles.textColor = [180, 83, 9];
                data.cell.styles.fontStyle = "bold";
              }
            }
          }
        },
      });

      // --- 4-COLUMN SIDE-BY-SIDE EXECUTIVE SUMMARY ---
      // 1. Garment Body
      const gBody = sortedGarments.map((item) => {
        const pct = totalPieces > 0 ? ((item.totalPcs / totalPieces) * 100).toFixed(1) : "0.0";
        return [item.name, item.totalLots.toString(), item.totalPcs.toLocaleString(), `${pct}%`];
      });
      gBody.push(["TOTAL", totalLots.toString(), totalPieces.toLocaleString(), "100.0%"]);

      // 2. Season Body
      const sBody = sortedSeasons.map((item) => {
        const pct = totalPieces > 0 ? ((item.totalPcs / totalPieces) * 100).toFixed(1) : "0.0";
        return [item.name, item.totalLots.toString(), item.totalPcs.toLocaleString(), `${pct}%`];
      });
      sBody.push(["TOTAL", totalLots.toString(), totalPieces.toLocaleString(), "100.0%"]);

      // 3. Brand Body
      const bBody = sortedBrands.map((brand) => {
        const pct = totalPieces > 0 ? ((brand.totalPcs / totalPieces) * 100).toFixed(1) : "0.0";
        return [brand.name, brand.totalLots.toString(), brand.totalPcs.toLocaleString(), `${pct}%`];
      });
      bBody.push(["TOTAL", totalLots.toString(), totalPieces.toLocaleString(), "100.0%"]);

      // 4. SLA Body
      const aBody = [
        ["<= 7 Days (Normal)", normalLots.toString(), normalPcs.toLocaleString(), `${totalLots > 0 ? ((normalLots / totalLots) * 100).toFixed(1) : 0}%`],
        ["8 - 14 Days (Urgent)", urgentLots.toString(), urgentPcs.toLocaleString(), `${totalLots > 0 ? ((urgentLots / totalLots) * 100).toFixed(1) : 0}%`],
        ["> 14 Days (Critical)", criticalLots.toString(), criticalPcs.toLocaleString(), `${totalLots > 0 ? ((criticalLots / totalLots) * 100).toFixed(1) : 0}%`],
        ["TOTAL", totalLots.toString(), totalPieces.toLocaleString(), "100.0%"],
      ];

      const maxRows = Math.max(gBody.length, sBody.length, bBody.length, aBody.length);
      const approxSummaryHeight = 55 + maxRows * 18;

      let summaryStartY = doc.lastAutoTable.finalY + 22;
      const neededSpace = approxSummaryHeight + 35;
      if (summaryStartY + neededSpace > pageH - 30) {
        doc.addPage();
        summaryStartY = 40;
      } else {
        doc.setDrawColor(203, 213, 225);
        doc.setLineWidth(0.8);
        doc.line(startX, summaryStartY - 8, startX + tableWidth, summaryStartY - 8);
      }

      // Title Banner (Centered)
      doc.setFillColor(15, 23, 42);
      doc.rect(startX, summaryStartY, tableWidth, 20, "F");
      doc.setFontSize(10);
      doc.setFont("helvetica", "bold");
      doc.setTextColor(255, 255, 255);
      doc.text("EXECUTIVE DISTRIBUTION & AGING SUMMARY", pageW / 2, summaryStartY + 14, { align: "center" });

      const sumTableTop = summaryStartY + 28;
      const gap = 12;
      const colW = (tableWidth - gap * 3) / 4;

      const sumTableStyles = {
        theme: "grid",
        styles: {
          fontSize: 7.5,
          cellPadding: { top: 2.5, right: 2, bottom: 2.5, left: 2 },
          overflow: "linebreak",
          valign: "middle",
          halign: "center",
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          lineWidth: 0.25,
          minCellHeight: 9,
        },
        headStyles: {
          textColor: [255, 255, 255],
          fontStyle: "bold",
          lineColor: [0, 0, 0],
          lineWidth: 0.35,
          halign: "center",
          fontSize: 8,
          cellPadding: { top: 3.5, right: 2, bottom: 3.5, left: 2 },
        },
        alternateRowStyles: { fillColor: [248, 250, 252] },
        columnStyles: {
          0: { cellWidth: colW * 0.44, halign: "center", fontStyle: "bold" },
          1: { cellWidth: colW * 0.18, halign: "center" },
          2: { cellWidth: colW * 0.22, halign: "center" },
          3: { cellWidth: colW * 0.16, halign: "center" },
        },
      };

      const makeTotalStyler = (bodyLength) => ({
        didParseCell: (d) => {
          if (d.section === "body" && d.row.index === bodyLength - 1) {
            d.cell.styles.fontStyle = "bold";
            d.cell.styles.fillColor = [226, 232, 240];
            d.cell.styles.textColor = [0, 0, 0];
          }
        },
      });

      // 1. Garment Table
      autoTable(doc, {
        ...sumTableStyles,
        head: [["Garment", "Lots", "Pieces", "%"]],
        body: gBody,
        startY: sumTableTop,
        margin: { left: startX, right: pageW - startX - colW },
        tableWidth: colW,
        headStyles: { ...sumTableStyles.headStyles, fillColor: [15, 118, 110] },
        ...makeTotalStyler(gBody.length),
      });

      // 2. Season Table
      autoTable(doc, {
        ...sumTableStyles,
        head: [["Season", "Lots", "Pieces", "%"]],
        body: sBody,
        startY: sumTableTop,
        margin: { left: startX + colW + gap, right: pageW - (startX + colW * 2 + gap) },
        tableWidth: colW,
        headStyles: { ...sumTableStyles.headStyles, fillColor: [67, 56, 202] },
        ...makeTotalStyler(sBody.length),
      });

      // 3. Brand Table
      autoTable(doc, {
        ...sumTableStyles,
        head: [["Brand", "Lots", "Pieces", "%"]],
        body: bBody,
        startY: sumTableTop,
        margin: { left: startX + (colW + gap) * 2, right: pageW - (startX + (colW + gap) * 3 - gap) },
        tableWidth: colW,
        headStyles: { ...sumTableStyles.headStyles, fillColor: [30, 64, 175] },
        ...makeTotalStyler(bBody.length),
      });

      // 4. Aging Table
      autoTable(doc, {
        ...sumTableStyles,
        head: [["Aging Status", "Lots", "Pieces", "%"]],
        body: aBody,
        startY: sumTableTop,
        margin: { left: startX + (colW + gap) * 3, right: startX },
        tableWidth: colW,
        headStyles: { ...sumTableStyles.headStyles, fillColor: [180, 83, 9] },
        didParseCell: (d) => {
          if (d.section === "body") {
            if (d.row.index === aBody.length - 1) {
              d.cell.styles.fontStyle = "bold";
              d.cell.styles.fillColor = [226, 232, 240];
              d.cell.styles.textColor = [0, 0, 0];
            } else if (d.row.index === 0) {
              d.cell.styles.fillColor = [220, 252, 231];
              d.cell.styles.textColor = [21, 128, 61];
            } else if (d.row.index === 1) {
              d.cell.styles.fillColor = [254, 243, 199];
              d.cell.styles.textColor = [180, 83, 9];
            } else if (d.row.index === 2) {
              d.cell.styles.fillColor = [254, 226, 226];
              d.cell.styles.textColor = [220, 38, 38];
            }
          }
        },
      });

      const totalPages = doc.internal.getNumberOfPages();
      for (let p = 1; p <= totalPages; p++) {
        doc.setPage(p);
        doc.setFontSize(8);
        doc.setFont("helvetica", "normal");
        doc.setTextColor(148, 163, 184);
        doc.text("Factory Suite Pro - Jaybir Printing Department", startX, pageH - 10);
        doc.text(`Page ${p} of ${totalPages}`, pageW - startX - 45, pageH - 10);
      }

      const fileDate = new Date().toISOString().slice(0, 10);
      const fileName = `Jaybir_Printing_Report_${fileDate}.pdf`;
      doc.save(fileName);
    } catch (err) {
      console.error("PDF generation error:", err);
      alert(`PDF Generation failed: ${err.message}`);
    }
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
          background: linear-gradient(135deg, #1e1b4b 0%, #312e81 100%);
          border-radius: 20px;
          padding: 24px 32px;
          color: white;
          margin-bottom: 24px;
          box-shadow: 0 10px 25px -5px rgba(30, 27, 75, 0.2);
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
        .header-btn-excel {
          background: #10b981 !important;
          border-color: #059669 !important;
          box-shadow: 0 4px 12px rgba(16, 185, 129, 0.35);
        }
        .header-btn-excel:hover {
          background: #059669 !important;
          box-shadow: 0 6px 16px rgba(5, 150, 105, 0.45);
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
          grid-template-columns: repeat(auto-fit, minmax(190px, 1fr));
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
          border-color: #ec4899;
        }
        .table-card {
          background: white;
          border-radius: 16px;
          border: 1.5px solid #cbd5e1;
          box-shadow: 0 4px 14px rgba(0, 0, 0, 0.04);
          overflow: hidden;
        }
        .table-header {
          padding: 14px 20px;
          border-bottom: 1.5px solid #cbd5e1;
          background: #f8fafc;
          display: flex;
          justify-content: space-between;
          align-items: center;
          flex-wrap: wrap;
          gap: 12px;
        }
        .table-container {
          overflow: auto;
          max-height: calc(100vh - 400px);
          min-height: 320px;
        }
        .data-table {
          width: 100%;
          border-collapse: collapse;
          border: 1.5px solid #cbd5e1;
          font-size: 0.82rem;
        }
        .data-table th {
          background: #1e1b4b;
          color: #ffffff;
          padding: 11px 8px;
          text-align: center;
          vertical-align: middle;
          font-weight: 800;
          font-size: 0.76rem;
          text-transform: uppercase;
          letter-spacing: 0.5px;
          position: sticky;
          top: 0;
          z-index: 10;
          border: 1px solid #312e81;
          white-space: nowrap;
        }
        .data-table td {
          padding: 9px 8px;
          border: 1px solid #cbd5e1;
          font-size: 0.82rem;
          color: #1e293b;
          text-align: center;
          vertical-align: middle;
        }
        .data-table tbody tr:nth-child(even) {
          background-color: #f8fafc;
        }
        .data-table tbody tr:hover {
          background-color: #f1f5f9;
        }
        .data-table tfoot td {
          background-color: #e2e8f0;
          border: 1.5px solid #94a3b8;
          font-weight: 800;
          font-size: 0.82rem;
          color: #0f172a;
          padding: 11px 8px;
          text-align: center;
          vertical-align: middle;
        }
        .status-badge {
          display: inline-flex;
          align-items: center;
          gap: 4px;
          padding: 4px 10px;
          border-radius: 6px;
          font-size: 0.78rem;
          font-weight: 800;
        }
        .active-tag-pill {
          display: inline-flex;
          align-items: center;
          gap: 6px;
          background: rgba(236, 72, 153, 0.1);
          color: #be185d;
          border: 1px solid #fbcfe8;
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
              <span style={{ fontSize: "1.8rem" }}>🖨️</span>
              <h1 style={{ margin: 0, fontSize: "1.45rem", fontWeight: 800, color: "#ffffff" }}>
                Jaybir Printing Report
              </h1>
              <span style={{ background: "#ec4899", color: "white", padding: "2px 8px", borderRadius: "8px", fontSize: "0.72rem", fontWeight: 800 }}>
                LIVE GOOGLE SHEET
              </span>
            </div>
            <p style={{ margin: "6px 0 0 0", color: "#cbd5e1", fontSize: "0.85rem" }}>
              Live production data linked from <strong>OVERLOCK..FOLDING..KAJBUTTON</strong> (`Jaybir Printing` Tab)
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
              className="header-btn header-btn-excel"
              onClick={exportToExcel}
              disabled={filteredData.length === 0}
            >
              📊 Export Excel (.xlsx)
            </button>
            <button
              className="header-btn header-btn-pdf"
              onClick={handleDownloadPDF}
              disabled={filteredData.length === 0}
            >
              📄 Download PDF
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
            <div className="stat-label">Total Printing Lots</div>
            <div className="stat-value" style={{ color: "#ec4899" }}>{totalLots}</div>
          </div>
          <div className="stat-card">
            <div className="stat-label">Pending Pieces</div>
            <div className="stat-value" style={{ color: "#3b82f6" }}>{pendingPcs.toLocaleString()}</div>
          </div>
          <div className="stat-card">
            <div className="stat-label">Completed Lots</div>
            <div className="stat-value" style={{ color: "#10b981" }}>{completedLots}</div>
          </div>
          <div className="stat-card">
            <div className="stat-label">Pending In Printing</div>
            <div className="stat-value" style={{ color: "#f59e0b" }}>{pendingLots}</div>
          </div>
          <div className="stat-card">
            <div className="stat-label">Average Aging</div>
            <div className="stat-value" style={{ color: "#6366f1" }}>{avgAging} <span style={{ fontSize: "1rem", fontWeight: 700 }}>Days</span></div>
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
              themeColor="#ec4899"
            />

            {/* Brand Filter (Multi-Select) */}
            <MultiSelectDropdown
              label="Brand"
              options={brandOptions}
              selectedValues={selectedBrands}
              onChange={setSelectedBrands}
              themeColor="#ec4899"
            />

            {/* Garment Type Filter (Multi-Select) */}
            <MultiSelectDropdown
              label="Garment Type"
              options={garmentOptions}
              selectedValues={selectedGarments}
              onChange={setSelectedGarments}
              themeColor="#ec4899"
            />

            {/* Fabric Filter (Multi-Select) */}
            <MultiSelectDropdown
              label="Fabric"
              options={fabricOptions}
              selectedValues={selectedFabrics}
              onChange={setSelectedFabrics}
              themeColor="#ec4899"
            />

            {/* Season Filter (Multi-Select) */}
            <MultiSelectDropdown
              label="Season"
              options={seasonOptions}
              selectedValues={selectedSeasons}
              onChange={setSelectedSeasons}
              themeColor="#ec4899"
            />

            {/* Section Filter (Multi-Select) */}
            <MultiSelectDropdown
              label="Section"
              options={sectionOptions}
              selectedValues={selectedSections}
              onChange={setSelectedSections}
              themeColor="#ec4899"
            />

            {/* Supervisor Filter (Multi-Select) */}
            <MultiSelectDropdown
              label="Supervisor"
              options={supervisorOptions}
              selectedValues={selectedSupervisors}
              onChange={setSelectedSupervisors}
              themeColor="#ec4899"
            />

            {/* Date Filter */}
            <div>
              <label className="filter-label">Printing Date</label>
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
              <span style={{ fontWeight: 800, fontSize: "1.05rem", color: "#0f172a" }}>Jaybir Printing Lots</span>
              <span style={{ background: "#fce7f3", padding: "3px 10px", borderRadius: "6px", fontSize: "0.8rem", fontWeight: 800, color: "#be185d" }}>
                Showing {filteredData.length} of {data.length} lots ({totalPcs.toLocaleString()} pcs)
              </span>
            </div>
          </div>

          <div className="table-container">
            {loading ? (
              <div style={{ textAlign: "center", padding: "60px" }}>
                <div style={{ fontSize: "2rem", marginBottom: "8px" }}>⚡</div>
                <div style={{ fontWeight: 700, color: "#64748b" }}>Loading Jaybir Printing records from Google Sheet...</div>
              </div>
            ) : filteredData.length === 0 ? (
              <div style={{ textAlign: "center", padding: "60px", color: "#64748b" }}>
                <div style={{ fontSize: "2.5rem", marginBottom: "8px" }}>📭</div>
                <div style={{ fontWeight: 800, fontSize: "1.1rem", color: "#0f172a" }}>No Printing Lots Found</div>
                <div style={{ fontSize: "0.85rem", marginTop: "4px" }}>
                  {hasActiveFilters ? "Try clearing or adjusting your multi-select filters." : "New entries from the Google Sheet tab will appear here."}
                </div>
              </div>
            ) : (
              <table className="data-table">
                <thead>
                  <tr>
                    <th>#</th>
                    <th>Timestamp</th>
                    <th>Lot Number</th>
                    <th>Garment Type</th>
                    <th>Style</th>
                    <th>Fabric</th>
                    <th>Brand</th>
                    <th>Total Pcs</th>
                    <th>Section</th>
                    <th>Season</th>
                    <th>Party Name</th>
                    <th>Direct Stitching</th>
                    <th>Supervisor</th>
                    <th>Printing Date</th>
                    <th>Aging</th>
                    <th>WIP Remarks</th>
                    <th>Status / Complete</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredData.map((item, idx) => (
                    <tr key={item.id}>
                      <td style={{ color: "#64748b", fontWeight: 700, background: "#f8fafc" }}>{idx + 1}</td>
                      <td style={{ fontSize: "0.78rem", color: "#64748b", whiteSpace: "nowrap" }}>{item.timestamp || "—"}</td>
                      <td>
                        <strong style={{ color: "#312e81", fontSize: "0.95rem" }}>
                          #{item.lotNumber}
                        </strong>
                      </td>
                      <td>{item.garmentType || "—"}</td>
                      <td>{item.style || "—"}</td>
                      <td>{item.fabric || "—"}</td>
                      <td style={{ fontWeight: 700 }}>{item.brand || "—"}</td>
                      <td>
                        <span style={{ background: "#e0f2fe", color: "#0369a1", padding: "3px 8px", borderRadius: "6px", fontWeight: 800 }}>
                          {(item.totalPcs || 0).toLocaleString()}
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
                      <td>{item.partyName || "—"}</td>
                      <td>{item.directStitching || "—"}</td>
                      <td>{item.supervisor || "—"}</td>
                      <td style={{ fontWeight: 700, whiteSpace: "nowrap" }}>{item.printingDate || "—"}</td>
                      <td>
                        {item.agingDays !== null ? (
                          <span
                            style={{
                              display: "inline-flex",
                              alignItems: "center",
                              justifyContent: "center",
                              padding: "3px 8px",
                              borderRadius: "6px",
                              fontWeight: 800,
                              fontSize: "0.78rem",
                              whiteSpace: "nowrap",
                              background: item.agingDays > 14 ? "#fee2e2" : item.agingDays > 7 ? "#fef3c7" : "#ecfdf5",
                              color: item.agingDays > 14 ? "#b91c1c" : item.agingDays > 7 ? "#b45309" : "#047857",
                              border: `1px solid ${item.agingDays > 14 ? "#fca5a5" : item.agingDays > 7 ? "#fde68a" : "#a7f3d0"}`
                            }}
                          >
                            {item.agingDays} {item.agingDays === 1 ? "Day" : "Days"}
                          </span>
                        ) : (
                          <span style={{ color: "#94a3b8" }}>—</span>
                        )}
                      </td>
                      <td>
                        {item.isCompleted ? (
                          <span style={{ color: "#047857", fontWeight: 800, background: "#d1fae5", padding: "3px 10px", borderRadius: "6px", fontSize: "0.78rem", border: "1px solid #a7f3d0" }}>
                            ✓ Done
                          </span>
                        ) : (
                          <span style={{ color: "#b45309", fontWeight: 700, background: "#fef3c7", padding: "3px 8px", borderRadius: "6px", fontSize: "0.78rem", border: "1px solid #fde68a" }}>
                            {item.wip || "In Progress"}
                          </span>
                        )}
                      </td>
                      <td>
                        {item.isCompleted ? (
                          <div style={{ display: "flex", flexDirection: "column", gap: "2px", alignItems: "center" }}>
                            <span className="status-badge" style={{ background: "#d1fae5", color: "#047857", border: "1px solid #a7f3d0" }}>
                              ✓ {item.completeDate || "Completed"}
                            </span>
                            {item.completionSupervisor && (
                              <span style={{ fontSize: "0.72rem", color: "#64748b" }}>
                                By: {item.completionSupervisor}
                              </span>
                            )}
                            {item.completionRemarks && (
                              <span style={{ fontSize: "0.72rem", color: "#475569", fontStyle: "italic" }}>
                                "{item.completionRemarks}"
                              </span>
                            )}
                          </div>
                        ) : (
                          <span className="status-badge" style={{ background: "#fef3c7", color: "#b45309", border: "1px solid #fde68a" }}>
                            ⏳ In Progress
                          </span>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr>
                    <td></td>
                    <td style={{ fontWeight: 800, color: "#1e1b4b" }}>TOTAL</td>
                    <td style={{ fontWeight: 800 }}>{totalLots} Lots</td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td style={{ fontWeight: 900, color: "#0369a1" }}>{totalPcs.toLocaleString()}</td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td style={{ fontWeight: 800, color: "#1e1b4b" }}>Avg: {avgAging}d</td>
                    <td style={{ fontWeight: 800, color: "#047857" }}>{completedLots} Done</td>
                    <td style={{ fontWeight: 800, color: "#b45309" }}>{pendingLots} Pending</td>
                  </tr>
                </tfoot>
              </table>
            )}
          </div>
        </div>
      </div>
    </>
  );
}
