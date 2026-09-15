import React, { useState, useEffect, useMemo, useRef } from "react";
import { useHistory } from "react-router-dom";
import { SPREADSHEET_IDS, fetchSheetDataFromBackend } from "./config";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

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
  const [selectedSeasons, setSelectedSeasons] = useState([]);
  const [selectedSections, setSelectedSections] = useState([]);
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
      const [feedUpRes, jobOrderRes] = await Promise.allSettled([
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
        feedUpRes.status !== "fulfilled" ||
        !feedUpRes.value?.ok ||
        !Array.isArray(feedUpRes.value.values) ||
        feedUpRes.value.values.length === 0
      ) {
        setData([]);
        setLoading(false);
        return;
      }

      const rows = feedUpRes.value.values;
      const headers = rows[0] || [];

      const tsIdx = findCol(headers, ["timestamp", "time"]);
      const lotIdx = findCol(headers, ["lot number", "lot no", "lot"]);
      const garmentIdx = findCol(headers, ["garment type", "garment"]);
      const fabricIdx = findCol(headers, ["fabric"]);
      const styleIdx = findCol(headers, ["style"]);
      const brandIdx = findCol(headers, ["brand", "brand name"]);
      const seasonIdx = findCol(headers, ["season"]);
      const sectionIdx = findCol(headers, ["section"]);
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

        const normLot = lotNo.toLowerCase();
        const jobInfo = lotToJobInfo[normLot] || {};
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
          supervisor: row[supIdx !== -1 ? supIdx : 6] || "Feed Up Department",
          feedUpDate: row[dateIdx !== -1 ? dateIdx : 7] || "",
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
      const matchesSupervisor = selectedSupervisors.length === 0 || selectedSupervisors.some((s) => s.toLowerCase() === item.supervisor.toLowerCase());
      const matchesDate = !dateFilter || item.feedUpDate.includes(dateFilter);

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

  // Summary Metrics
  const totalLots = filteredData.length;
  const totalPcs = filteredData.reduce((sum, item) => sum + item.totalPcs, 0);
  const completedLots = filteredData.filter((item) => item.isCompleted).length;
  const pendingLots = totalLots - completedLots;

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
      const ws1 = workbook.addWorksheet("Feed Up Report", {
        views: [{ showGridLines: true }],
      });

      // Title Banner
      ws1.mergeCells("A1:N1");
      const titleCell = ws1.getCell("A1");
      titleCell.value = "FACTORY SUITE PRO - FEED UP DEPARTMENT PRODUCTION REPORT";
      titleCell.font = { name: "Segoe UI", size: 14, bold: true, color: { argb: "FFFFFFFF" } };
      titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F172A" } };
      titleCell.alignment = { horizontal: "center", vertical: "middle" };
      ws1.getRow(1).height = 32;

      // Subtitle
      ws1.mergeCells("A2:N2");
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
        "Feed Up Date",
        "Feed Up Supervisor",
        "Status",
        "WIP Remarks",
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
          item.feedUpDate || "—",
          item.supervisor || "—",
          isComp ? "Completed" : "Pending",
          item.wip || "—",
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
          row.getCell(13).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDCFCE7" } };
          row.getCell(13).font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FF15803D" } };
        } else {
          row.getCell(13).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE0E7FF" } };
          row.getCell(13).font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FF3730A3" } };
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
        `${compCount} Comp | ${pendCount} Pending`,
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

      const colWidths = [8, 16, 20, 20, 20, 16, 14, 12, 14, 20, 16, 20, 16, 30];
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
      supTitle.value = "2. FEED UP SUPERVISOR WORKLOAD";
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
        { name: "Pending In Feed Up", lots: pendCount, pcs: pendPcs, bg: "FFE0E7FF", fg: "FF3730A3" },
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
        ["Garment Type Filter", selectedGarments.length ? selectedGarments.join(", ") : "All Types"],
        ["Fabric Filter", selectedFabrics.length ? selectedFabrics.join(", ") : "All Fabrics"],
        ["Brand Filter", selectedBrands.length ? selectedBrands.join(", ") : "All Brands"],
        ["Date Filter", dateFilter || "All Dates"],
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
      saveAs(new Blob([buffer]), `Feed_Up_Report_${ts}.xlsx`);
    } catch (e) {
      console.error("Excel export error:", e);
      alert(`Excel export failed: ${e.message}`);
    }
  };

  // CSV Export
  const exportToCSV = () => {
    if (filteredData.length === 0) return;
    const headers = ["Timestamp", "Lot Number", "Garment Type", "Fabric", "Style", "Brand", "Season", "Section", "Feed Up Supervisor", "Feed Up Date", "Total Pcs", "WIP Remarks", "Status", "Completion Date"];
    const rows = filteredData.map((d) => [
      `"${d.timestamp}"`,
      `"${d.lotNumber}"`,
      `"${d.garmentType}"`,
      `"${d.fabric}"`,
      `"${d.style}"`,
      `"${d.brand}"`,
      `"${d.season || ''}"`,
      `"${d.section || ''}"`,
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

  // PDF Export Function - 14 columns spacious & centered layout
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
      let reportTitle = "FACTORY SUITE PRO - FEED UP DEPARTMENT REPORT";
      if (isPendingView) reportTitle = "FACTORY SUITE PRO - FEED UP DEPARTMENT REPORT (PENDING LOTS)";
      else if (isCompletedView) reportTitle = "FACTORY SUITE PRO - FEED UP DEPARTMENT REPORT (COMPLETED LOTS)";
      doc.text(reportTitle, pageW / 2, 30, { align: 'center' });

      doc.setFontSize(8.5);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(199, 210, 254);
      const subText = `Total Lots: ${totalLots}   |   Total Pieces: ${totalPieces.toLocaleString()}   |   Completed Lots: ${completedLotsCount}   |   Pending Lots: ${pendingLotsCount}   |   Supervisors: ${sortedSupervisors.length}`;
      doc.text(subText, pageW / 2, 48, { align: 'center' });

      // 2. Filter Banner
      doc.setFillColor(241, 245, 249);
      doc.rect(15, 63, pageW - 30, 16, 'F');
      doc.setFontSize(8);
      doc.setFont('helvetica', 'italic');
      doc.setTextColor(0, 0, 0);
      const filterSummary = `Filters: Status: ${selectedStatuses.length ? selectedStatuses.join(', ') : 'All'} | Brands: ${selectedBrands.length ? selectedBrands.join(', ') : 'All'} | Garments: ${selectedGarments.length ? selectedGarments.join(', ') : 'All'} | Fabrics: ${selectedFabrics.length ? selectedFabrics.join(', ') : 'All'} | Seasons: ${selectedSeasons.length ? selectedSeasons.join(', ') : 'All'} | Sections: ${selectedSections.length ? selectedSections.join(', ') : 'All'} | Supervisors: ${selectedSupervisors.length ? selectedSupervisors.join(', ') : 'All'} | Search: ${searchTerm || 'None'}`;
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
        'Feed Up Date',
        'WIP Remarks',
        'Feed Up Sup',
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
          formatDisplayDate(item.feedUpDate) || '—',
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
        `${sortedSupervisors.length} Sups`,
        `${completedLotsCount} Comp | ${pendingLotsCount} Pend`,
        ''
      ]);

      const columnStyles = {
        0: { cellWidth: 25, halign: 'center' },
        1: { cellWidth: 70, halign: 'center', fontStyle: 'bold' },
        2: { cellWidth: 85, halign: 'center' },
        3: { cellWidth: 85, halign: 'center' },
        4: { cellWidth: 85, halign: 'center' },
        5: { cellWidth: 70, halign: 'center' },
        6: { cellWidth: 60, halign: 'center', fontStyle: 'bold' },
        7: { cellWidth: 50, halign: 'center' },
        8: { cellWidth: 55, halign: 'center' },
        9: { cellWidth: 85, halign: 'center' },
        10: { cellWidth: 55, halign: 'center' },
        11: { cellWidth: 70, halign: 'center' },
        12: { cellWidth: 175, halign: 'center' },
        13: { cellWidth: 80, halign: 'center' },
        14: { cellWidth: 70, halign: 'center' },
        15: { cellWidth: 80, halign: 'center' }
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
            if (data.column.index === 14) {
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
        doc.text("Feed Up Department Report — Factory Suite Pro", 15, finalY + 12);
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
      doc.save(`Feed_Up_Report_${fileDate}.pdf`);

    } catch (error) {
      console.error("Error generating Feed Up Report PDF:", error);
      alert("Failed to generate PDF. Please try again.");
    }
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
          padding: 8px 16px;
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
          background: rgba(255, 255, 255, 0.25);
          transform: translateY(-1px);
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

          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", alignItems: "center" }}>
            <button className="header-btn" onClick={() => history.push("/dashboard")}>
              ← Dashboard
            </button>
            <button className="header-btn" onClick={() => history.push("/stitching-complete-lot")}>
              📊 Stitching Report
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
            <button className="header-btn" onClick={exportToExcel} disabled={filteredData.length === 0}>
              📊 Export Excel
            </button>
            <button className="header-btn" onClick={exportToCSV} disabled={filteredData.length === 0}>
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

            {/* Season Filter (Multi-Select) */}
            <MultiSelectDropdown
              label="Season"
              options={seasonOptions}
              selectedValues={selectedSeasons}
              onChange={setSelectedSeasons}
            />

            {/* Section Filter (Multi-Select) */}
            <MultiSelectDropdown
              label="Section"
              options={sectionOptions}
              selectedValues={selectedSections}
              onChange={setSelectedSections}
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
                    <th>Style</th>
                    <th>Fabric</th>
                    <th>Brand</th>
                    <th>Total Pcs</th>
                    <th>Section</th>
                    <th>Season</th>
                    <th>Party Name</th>
                    <th>Direct Stitching</th>
                    <th>Feed Up Supervisor</th>
                    <th>Feed Up Date</th>
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
                      <td style={{ fontWeight: 700 }}>{item.style || "—"}</td>
                      <td>{item.fabric || "—"}</td>
                      <td>{item.brand || "—"}</td>
                      <td>
                        <span style={{ fontWeight: 800, color: "#0f172a" }}>
                          {item.totalPcs.toLocaleString()}
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
