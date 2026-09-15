// src/DailyStitchingIssue.js
import React, { useEffect, useMemo, useState, useRef } from "react";
import { Link, useHistory } from "react-router-dom";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { GOOGLE_API_KEY, SPREADSHEET_IDS, fetchSheetDataFromBackend } from "./config";
import { getCurrentUser, logoutUser } from "./auth";

/**
 * Factory Suite Pro - Daily Stitching Issue Tracker
 * Full Executive Standard:
 * - Perfectly aligned 13-column interactive data grid
 * - Parallel data sync from Index, JobOrder, Cutting matrix, and Historical Stitching Issues
 * - Comprehensive 3-Sheet Excel Export (.xlsx) via ExcelJS
 * - Direct A3 Landscape PDF Export (.pdf) with Pure Black Text & 4-Column Executive Summary
 * - Top Brand Bar with Navigation Links & Live Filter Audit
 */

// ====== CONFIG ======
const API_KEY = GOOGLE_API_KEY;
const BUDGET_SHEET_ID = SPREADSHEET_IDS.MAIN;
const JOB_SHEET_ID = SPREADSHEET_IDS.JOBORDER;
const INDEX_SHEET_NAME = "Index";
const CUTTING_SHEET_NAME = "Cutting";
const OLD_SHEET_ID = "18FzakygM7DVD29IRbpe68pDeCFQhFLj7t4C-XQ1MWWc";
const OLD_SHEET_NAME = "Stitching_Issues";

// Ranges
const INDEX_RANGE = `${INDEX_SHEET_NAME}!A:AG`;
const CUTTING_BIG_RANGE = `${CUTTING_SHEET_NAME}!A1:ZZ400000`;
const OLD_SHEET_RANGE = `${OLD_SHEET_NAME}!A:Q`;
const JOB_RANGE = "JobOrder!A:AZ";

// Display columns (13 canonical columns)
const DISPLAY_HEADERS = [
  "Sr. No",
  "Lot Number",
  "Garment Type",
  "Style",
  "Fabric",
  "Brand",
  "PCS",
  "Section",
  "Season",
  "Party Name",
  "Direct Stitching",
  "Supervisor",
  "Date of Issue",
];

const COLUMN_ICONS = {
  "Sr. No": "#️⃣",
  "Lot Number": "🏷️",
  "Garment Type": "👕",
  "Style": "🎨",
  "Fabric": "🧵",
  "Brand": "🏢",
  "PCS": "🔢",
  "Section": "👥",
  "Season": "🍂",
  "Party Name": "🤝",
  "Direct Stitching": "⚡",
  "Supervisor": "👨‍💼",
  "Date of Issue": "📅",
};

// Recent lot threshold (24 hours)
const RECENT_THRESHOLD_MS = 24 * 60 * 60 * 1000;

/* ---------- Utility Functions ---------- */
const norm = (s) => String(s || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");

const normalizeKey = (s = "") => {
  return norm(s);
};

function formatDisplayDate(d) {
  if (!d) return "";
  const parsed = new Date(d);
  if (isNaN(parsed.getTime())) return String(d);
  return parsed.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
}

// Parse date string to Date object
const parseDateString = (dateStr) => {
  if (!dateStr) return null;
  const str = String(dateStr).trim();

  // Try direct date parsing
  const d1 = new Date(str);
  if (!isNaN(d1.getTime()) && str.includes("-") && str.length >= 8) {
    return d1;
  }

  // DD/MM/YYYY or DD-MM-YYYY
  const dmy = str.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})/);
  if (dmy) {
    const day = parseInt(dmy[1], 10);
    const month = parseInt(dmy[2], 10) - 1;
    const year = parseInt(dmy[3], 10);
    const d = new Date(year, month, day);
    if (!isNaN(d.getTime())) return d;
  }

  // DD MMM YYYY (e.g. 15 Sep 2026 or 15 Sept 2026)
  const monthMap = {
    jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
    jul: 6, aug: 7, sep: 8, sept: 8, oct: 9, nov: 10, dec: 11
  };
  const dMonthY = str.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})/);
  if (dMonthY) {
    const day = parseInt(dMonthY[1], 10);
    const mKey = dMonthY[2].toLowerCase().slice(0, 4);
    const mIdx = monthMap[mKey] ?? monthMap[mKey.slice(0, 3)];
    if (mIdx !== undefined) {
      const year = parseInt(dMonthY[3], 10);
      const d = new Date(year, mIdx, day);
      if (!isNaN(d.getTime())) return d;
    }
  }

  return null;
};

const clean = (v) => (v == null ? "" : String(v).trim());

const titleCase = (s) =>
  String(s || "")
    .split(/\s+/)
    .filter(Boolean)
    .map((w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
    .join(" ");

const normalizeSupervisor = (sRaw) => {
  const s = clean(sRaw);
  if (!s) return { display: "—", key: "" };
  const key = norm(s);
  const display = titleCase(s);
  return { display, key };
};

/* ---------- Cutting Matrix PCS Calculation ---------- */
function sliceCuttingMatrix(bigValues, startRow, numRows) {
  if (!Array.isArray(bigValues) || bigValues.length === 0) return [];
  if (!(startRow > 0 && numRows > 0)) return [];
  const r0 = Math.max(0, startRow - 1);
  const r1 = Math.min(bigValues.length - 1, r0 + numRows - 1);
  return bigValues.slice(r0, r1 + 1);
}

function findHeaderRowIndex(windowValues, expectedSizesNorm) {
  const hasSizeToken = (rowSet) => expectedSizesNorm.some((sz) => rowSet.has(sz));
  for (let i = 0; i < windowValues.length; i++) {
    const row = windowValues[i] || [];
    const set = new Set(row.map((c) => norm(c)));
    const hasShadeHeader = set.has("color") || set.has("shade") || set.has("shades");
    if (hasShadeHeader && hasSizeToken(set)) return i;
  }
  for (let i = 0; i < windowValues.length; i++) {
    const row = windowValues[i] || [];
    const set = new Set(row.map((c) => norm(c)));
    let matches = 0;
    expectedSizesNorm.forEach((sz) => {
      if (set.has(sz)) matches++;
    });
    if (matches >= 2) return i;
  }
  return 0;
}

function calculateTotalPCS(windowValues, sizes = []) {
  if (!windowValues || windowValues.length === 0) return 0;

  const normalizedSizes = Array.from(new Set((sizes || []).map(norm).filter(Boolean)));
  const headerRowIdx = findHeaderRowIndex(windowValues, normalizedSizes);
  const header = windowValues[headerRowIdx] || [];

  const hIdx = {};
  header.forEach((h, i) => {
    const k = norm(h);
    if (k && !(k in hIdx)) hIdx[k] = i;
  });

  const nonSizeColumns = new Set([
    "color", "shade", "shades", "cuttingtable", "cutting", "table",
    "total", "totalpcs", "totals", "grandtotal", "sum", "lot", "style",
    "fabric", "garment", "partyname", "brand", "section", "season"
  ]);

  let sizeColIndices = [];
  header.forEach((h, i) => {
    const normalizedHeader = norm(h);
    if (normalizedHeader && !nonSizeColumns.has(normalizedHeader)) {
      sizeColIndices.push(i);
    }
  });

  if (sizeColIndices.length === 0) {
    normalizedSizes.forEach((ns) => {
      if (ns in hIdx) sizeColIndices.push(hIdx[ns]);
    });

    if (sizeColIndices.length === 0) {
      const ct = hIdx["cuttingtable"];
      if (ct != null && ct >= 0) {
        const guessStart = ct + 1;
        const guessed = [];
        for (let k = 0; k < normalizedSizes.length; k++) guessed.push(guessStart + k);
        sizeColIndices = Array.from(new Set(guessed.filter((g) => g < header.length)));
      }
    }
  }

  if (sizeColIndices.length === 0) return 0;

  let totalQty = 0;

  for (let r = headerRowIdx + 1; r < windowValues.length; r++) {
    const row = windowValues[r] || [];
    const rawShade = String(row[hIdx["color"] || hIdx["shade"] || 0] || "").trim();
    const shadeKey = norm(rawShade);

    if (!shadeKey || shadeKey === "total" || shadeKey === "totals" || shadeKey === "grandtotal") continue;

    sizeColIndices.forEach((c) => {
      const raw = row[c];
      if (raw != null && raw !== "") {
        const n = parseFloat(String(raw).replace(/,/g, ""));
        if (!isNaN(n) && n > 0) {
          totalQty += n;
        }
      }
    });
  }

  return totalQty;
}

/* ---------- Main Component ---------- */
export default function DailyStitchingIssue() {
  const history = useHistory();
  const currentUser = getCurrentUser();

  const handleLogout = () => {
    logoutUser();
    history.push("/");
  };

  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [refreshing, setRefreshing] = useState(false);
  const [loadingProgress, setLoadingProgress] = useState(0);
  const [loadingMessage, setLoadingMessage] = useState("Loading stitching issue data...");
  const [error, setError] = useState("");
  const [lastUpdated, setLastUpdated] = useState(null);

  // Filters and UI state
  const [filterLot, setFilterLot] = useState("");
  const [filterSupervisor, setFilterSupervisor] = useState("");
  const [filterIssueDate, setFilterIssueDate] = useState("");
  const [filterBrand, setFilterBrand] = useState("");
  const [filterGarment, setFilterGarment] = useState("");
  const [searchTerm, setSearchTerm] = useState("");
  const [showOnlyRecent, setShowOnlyRecent] = useState(false);

  // Date Range Filter State
  const [dateRangeFilter, setDateRangeFilter] = useState({
    startDate: "",
    endDate: "",
    enabled: false
  });

  const [sortConfig, setSortConfig] = useState({
    key: "Date of Issue",
    direction: "desc",
  });
  const [pageSize, setPageSize] = useState(25);
  const [page, setPage] = useState(1);

  const abortRef = useRef(null);

  const loadData = async (mode = "initial") => {
    if (mode === "initial") {
      setLoading(true);
      setLoadingProgress(0);
      setLoadingMessage("Loading stitching issue data...");
    } else {
      setRefreshing(true);
    }
    setError("");

    try {
      abortRef.current?.abort();
    } catch { }
    const ctrl = new AbortController();
    abortRef.current = ctrl;

    try {
      // 1. Fetch Index, JobOrder, Cutting, and Old Stitching in parallel
      setLoadingMessage("Fetching production sheets...");
      setLoadingProgress(15);

      const [idxRes, jobRes, cuttingRes, oldRes] = await Promise.all([
        fetchSheetDataFromBackend(BUDGET_SHEET_ID, INDEX_RANGE),
        fetchSheetDataFromBackend(JOB_SHEET_ID, JOB_RANGE),
        fetchSheetDataFromBackend(BUDGET_SHEET_ID, CUTTING_BIG_RANGE),
        fetchSheetDataFromBackend(OLD_SHEET_ID, OLD_SHEET_RANGE)
      ]);

      setLoadingProgress(50);
      setLoadingMessage("Processing Job Orders & Index maps...");

      // Parse JobOrder Sheet for authoritative master metadata
      const jobMap = new Map();
      const jobValues = jobRes?.values || [];
      if (jobValues.length > 0) {
        const jHeader = (jobValues[0] || []).map(h => normalizeKey(String(h || "")));
        const getJ = (row, key) => {
          const idx = jHeader.indexOf(key);
          return idx !== -1 ? (row[idx] ?? "") : "";
        };

        for (let i = 1; i < jobValues.length; i++) {
          const row = jobValues[i] || [];
          const lot = clean(getJ(row, "lotno") || getJ(row, "lotnumber") || getJ(row, "lot"));
          if (!lot) continue;

          jobMap.set(norm(lot), {
            lot,
            garmentType: clean(getJ(row, "garmenttype") || getJ(row, "garment")),
            style: clean(getJ(row, "style")),
            fabric: clean(getJ(row, "fabric")),
            brand: clean(getJ(row, "brand")),
            section: clean(getJ(row, "section") || getJ(row, "mwk")),
            season: clean(getJ(row, "season")),
            partyName: clean(getJ(row, "partyname") || getJ(row, "party")),
            directStitching: clean(getJ(row, "directstitching") || getJ(row, "direct")),
            jobOrderNo: clean(getJ(row, "joborderno") || getJ(row, "joborder")),
            date: clean(getJ(row, "date") || getJ(row, "podate"))
          });
        }
      }

      setLoadingProgress(70);
      setLoadingMessage("Calculating cutting quantities & index issues...");

      // Parse Index Sheet
      const idxValues = idxRes?.values || [];
      const idxHeader = (idxValues[0] || []).map(h => normalizeKey(String(h || "")));
      const getIdx = (row, key) => {
        const idx = idxHeader.indexOf(key);
        return idx !== -1 ? (row[idx] ?? "") : "";
      };

      const bigCuttingValues = cuttingRes?.values || [];
      const indexRows = [];

      for (let i = 1; i < idxValues.length; i++) {
        const row = idxValues[i] || [];
        const lot = clean(getIdx(row, "lotnumber") || getIdx(row, "lotno") || getIdx(row, "lot"));
        const supervisor = clean(getIdx(row, "supervisor"));
        const dateOfIssue = clean(getIdx(row, "dateofissue") || getIdx(row, "date"));

        if (!lot || !supervisor || !dateOfIssue) continue;

        const startRow = parseInt(getIdx(row, "startrow") || "0", 10);
        const numRows = parseInt(getIdx(row, "numrows") || "0", 10);
        const sizes = String(getIdx(row, "sizes") || "").split(",").map(s => s.trim()).filter(Boolean);

        const window = sliceCuttingMatrix(bigCuttingValues, startRow, numRows);
        const totalPCS = calculateTotalPCS(window, sizes);

        const jInfo = jobMap.get(norm(lot)) || {};
        const { display: supDisplay, key: supKey } = normalizeSupervisor(supervisor);
        const parsedDate = parseDateString(dateOfIssue);

        const garmentType = jInfo.garmentType || clean(getIdx(row, "garmenttype") || getIdx(row, "garment"));
        const style = jInfo.style || clean(getIdx(row, "style"));
        const fabric = jInfo.fabric || clean(getIdx(row, "fabric"));
        const brand = jInfo.brand || clean(getIdx(row, "brand"));
        const section = jInfo.section || clean(getIdx(row, "section") || getIdx(row, "mwk"));
        const season = jInfo.season || clean(getIdx(row, "season"));
        let partyName = jInfo.partyName || clean(getIdx(row, "partyname") || getIdx(row, "party"));
        if (partyName.toLowerCase().includes("mohit")) partyName = "MH (Mohit Hosiery)";
        const directStitching = jInfo.directStitching || clean(getIdx(row, "directstitching") || getIdx(row, "direct"));

        indexRows.push({
          "Date of Issue": formatDisplayDate(dateOfIssue),
          _rawDateOfIssue: dateOfIssue,
          _parsedDate: parsedDate,
          "Lot Number": lot,
          "Garment Type": garmentType || "—",
          "Style": style || "—",
          "Fabric": fabric || "—",
          "Brand": brand || "—",
          "PCS": totalPCS > 0 ? totalPCS : 0,
          "Section": section || "—",
          "Season": season || "—",
          "Party Name": partyName || "—",
          "Direct Stitching": directStitching ? (directStitching.toLowerCase() === "yes" ? "Yes" : "No") : "No",
          "Supervisor": supDisplay,
          _supKey: supKey,
          "Source Type": "Index",
          _isRecent: Date.now() - (parsedDate ? parsedDate.getTime() : Date.now()) < RECENT_THRESHOLD_MS
        });
      }

      setLoadingProgress(85);
      setLoadingMessage("Processing historical stitching logs...");

      // Parse Old Stitching Issues Sheet (Aggregated)
      const oldValues = oldRes?.values || [];
      const oldHeader = (oldValues[0] || []).map(h => normalizeKey(String(h || "")));
      const getOld = (row, key) => {
        const idx = oldHeader.indexOf(key);
        return idx !== -1 ? (row[idx] ?? "") : "";
      };

      const aggregationMap = new Map();

      for (let i = 1; i < oldValues.length; i++) {
        const row = oldValues[i] || [];
        const dateOfIssue = clean(getOld(row, "dateofissue") || getOld(row, "timestamp") || getOld(row, "date"));
        const lotNumber = clean(getOld(row, "lotnumber") || getOld(row, "lotno") || getOld(row, "lot"));
        const supervisor = clean(getOld(row, "supervisor"));
        const fabric = clean(getOld(row, "fabric"));
        const garmentType = clean(getOld(row, "garmenttype") || getOld(row, "garment"));
        const pcs = parseFloat(String(getOld(row, "pcs") || "0").replace(/,/g, "")) || 0;
        const brand = clean(getOld(row, "brand"));
        const style = clean(getOld(row, "style"));
        const section = clean(getOld(row, "section") || getOld(row, "mwk"));
        const season = clean(getOld(row, "season"));
        const partyName = clean(getOld(row, "partyname") || getOld(row, "party"));
        const directStitching = clean(getOld(row, "directstitching") || getOld(row, "direct"));

        if (!lotNumber || !supervisor || !dateOfIssue) continue;

        const { display: supDisplay, key: supKey } = normalizeSupervisor(supervisor);
        const aggKey = `${norm(lotNumber)}|${supKey}|${norm(dateOfIssue)}`;

        const jInfo = jobMap.get(norm(lotNumber)) || {};

        if (!aggregationMap.has(aggKey)) {
          const parsedDate = parseDateString(dateOfIssue);
          let pDisplay = jInfo.partyName || partyName || "—";
          if (pDisplay.toLowerCase().includes("mohit")) pDisplay = "MH (Mohit Hosiery)";

          aggregationMap.set(aggKey, {
            "Date of Issue": formatDisplayDate(dateOfIssue),
            _rawDateOfIssue: dateOfIssue,
            _parsedDate: parsedDate,
            "Lot Number": lotNumber,
            "Garment Type": jInfo.garmentType || garmentType || "—",
            "Style": jInfo.style || style || "—",
            "Fabric": jInfo.fabric || fabric || "—",
            "Brand": jInfo.brand || brand || "—",
            "PCS": pcs,
            "Section": jInfo.section || section || "—",
            "Season": jInfo.season || season || "—",
            "Party Name": pDisplay,
            "Direct Stitching": (jInfo.directStitching || directStitching) ? (String(jInfo.directStitching || directStitching).toLowerCase() === "yes" ? "Yes" : "No") : "No",
            "Supervisor": supDisplay,
            _supKey: supKey,
            "Source Type": "Old Lot",
            _isRecent: Date.now() - (parsedDate ? parsedDate.getTime() : Date.now()) < RECENT_THRESHOLD_MS
          });
        } else {
          const existing = aggregationMap.get(aggKey);
          existing.PCS += pcs;
        }
      }

      const oldRows = Array.from(aggregationMap.values());

      // Combine both datasets
      const combined = [...indexRows, ...oldRows];
      setRows(combined);
      setLastUpdated(new Date().toLocaleString());
      setLoadingProgress(100);

    } catch (e) {
      if (e?.name === "AbortError") {
        console.log("Request aborted");
      } else {
        console.error("Error loading data:", e);
        setError(e.message || "Failed to load data");
      }
    } finally {
      if (mode === "initial") {
        setTimeout(() => {
          setLoading(false);
          setLoadingProgress(0);
        }, 300);
      } else {
        setRefreshing(false);
      }
    }
  };

  useEffect(() => {
    loadData("initial");
    return () => {
      try {
        abortRef.current?.abort();
      } catch { }
    };
  }, []);

  /* ---------- Filter Options ---------- */
  const uniqueLots = useMemo(() => {
    return Array.from(new Set(rows.map((r) => r["Lot Number"]).filter(Boolean))).sort();
  }, [rows]);

  const uniqueSupervisors = useMemo(() => {
    const map = new Map();
    rows.forEach((r) => {
      if (r._supKey && r.Supervisor && r.Supervisor !== "—") {
        map.set(r._supKey, r.Supervisor);
      }
    });
    return Array.from(map.entries())
      .map(([key, label]) => ({ key, label }))
      .sort((a, b) => a.label.localeCompare(b.label));
  }, [rows]);

  const uniqueIssueDates = useMemo(() => {
    return Array.from(new Set(rows.map((r) => r["Date of Issue"]).filter(Boolean))).sort();
  }, [rows]);

  const uniqueBrands = useMemo(() => {
    return Array.from(new Set(rows.map((r) => r["Brand"]).filter(b => b && b !== "—"))).sort();
  }, [rows]);

  const uniqueGarments = useMemo(() => {
    return Array.from(new Set(rows.map((r) => r["Garment Type"]).filter(g => g && g !== "—"))).sort();
  }, [rows]);

  /* ---------- Enhanced Filtering with Date Range ---------- */
  const filteredRows = useMemo(() => {
    let filtered = rows;

    if (filterLot) filtered = filtered.filter((r) => r["Lot Number"] === filterLot);
    if (filterSupervisor) filtered = filtered.filter((r) => r._supKey === filterSupervisor);
    if (filterIssueDate) filtered = filtered.filter((r) => r["Date of Issue"] === filterIssueDate);
    if (filterBrand) filtered = filtered.filter((r) => r["Brand"] === filterBrand);
    if (filterGarment) filtered = filtered.filter((r) => r["Garment Type"] === filterGarment);
    if (showOnlyRecent) filtered = filtered.filter((r) => r._isRecent);

    // Apply date range filter
    if (dateRangeFilter.enabled && (dateRangeFilter.startDate || dateRangeFilter.endDate)) {
      filtered = filtered.filter((row) => {
        const rowDate = row._parsedDate;
        if (!rowDate) return false;

        const rowTime = rowDate.getTime();
        const startTime = dateRangeFilter.startDate ? new Date(dateRangeFilter.startDate).getTime() : null;
        const endTime = dateRangeFilter.endDate ? new Date(dateRangeFilter.endDate).getTime() + 86400000 : null;

        if (startTime && rowTime < startTime) return false;
        if (endTime && rowTime >= endTime) return false;

        return true;
      });
    }

    // Apply search across all canonical columns
    if (searchTerm.trim()) {
      const term = searchTerm.toLowerCase().trim();
      filtered = filtered.filter((r) =>
        DISPLAY_HEADERS.some((h) => String(r[h] ?? "").toLowerCase().includes(term))
      );
    }

    // Apply sorting
    if (sortConfig.key) {
      const { key, direction } = sortConfig;
      filtered = [...filtered].sort((a, b) => {
        if (key === "Date of Issue") {
          const da = a._parsedDate ? a._parsedDate.getTime() : 0;
          const db = b._parsedDate ? b._parsedDate.getTime() : 0;
          if (da < db) return direction === "asc" ? -1 : 1;
          if (da > db) return direction === "asc" ? 1 : -1;
          return 0;
        }

        if (key === "PCS") {
          const va = Number(a.PCS) || 0;
          const vb = Number(b.PCS) || 0;
          if (va < vb) return direction === "asc" ? -1 : 1;
          if (va > vb) return direction === "asc" ? 1 : -1;
          return 0;
        }

        const va = String(a[key] ?? "").toLowerCase();
        const vb = String(b[key] ?? "").toLowerCase();
        if (va < vb) return direction === "asc" ? -1 : 1;
        if (va > vb) return direction === "asc" ? 1 : -1;
        return 0;
      });
    }

    return filtered;
  }, [rows, filterLot, filterSupervisor, filterIssueDate, filterBrand, filterGarment, showOnlyRecent, dateRangeFilter, searchTerm, sortConfig]);

  /* ---------- Analytics ---------- */
  const analytics = useMemo(() => {
    const totalRecords = filteredRows.length;
    const totalPCS = filteredRows.reduce((sum, r) => sum + (Number(r.PCS) || 0), 0);
    const uniqueLotsSet = new Set(filteredRows.map((r) => r["Lot Number"]).filter(Boolean));
    const recentLotsCount = filteredRows.filter((r) => r._isRecent).length;
    const supervisorsSet = new Set(filteredRows.map((r) => r.Supervisor).filter(s => s && s !== "—"));

    return {
      totalRecords,
      totalPCS,
      uniqueLots: uniqueLotsSet.size,
      recentLots: recentLotsCount,
      supervisorsCount: supervisorsSet.size
    };
  }, [filteredRows]);

  const handleSort = (key) => {
    setSortConfig((prev) => ({
      key,
      direction: prev.key === key && prev.direction === "asc" ? "desc" : "asc",
    }));
  };

  const clearFilters = () => {
    setFilterLot("");
    setFilterSupervisor("");
    setFilterIssueDate("");
    setFilterBrand("");
    setFilterGarment("");
    setSearchTerm("");
    setShowOnlyRecent(false);
    setDateRangeFilter({ startDate: "", endDate: "", enabled: false });
    setSortConfig({ key: "Date of Issue", direction: "desc" });
    setPage(1);
  };

  const handleRefresh = () => loadData("refresh");

  const handleGoBack = () => {
    try {
      if (window.history.length > 1) {
        history.goBack();
        return;
      }
    } catch { }
    history.push("/dashboard");
  };

  const hasActiveFilters = Boolean(
    filterLot ||
    filterSupervisor ||
    filterIssueDate ||
    filterBrand ||
    filterGarment ||
    searchTerm ||
    showOnlyRecent ||
    (dateRangeFilter.enabled && (dateRangeFilter.startDate || dateRangeFilter.endDate))
  );

  /* ---------- Export Excel via ExcelJS (Factory Suite Pro 3-Sheet Workbook) ---------- */
  const handleExportExcel = async () => {
    try {
      if (filteredRows.length === 0) {
        alert("No data available to export.");
        return;
      }

      const now = new Date();
      const reportDateStr = `${String(now.getDate()).padStart(2, '0')}/${String(now.getMonth() + 1).padStart(2, '0')}/${now.getFullYear()} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;

      const totalLots = filteredRows.length;
      let totalQty = 0;
      const garmentAnalysis = {};
      const supervisorAnalysis = {};
      const partyAnalysis = {};
      const sectionAnalysis = {};

      filteredRows.forEach(row => {
        const qty = Number(row.PCS) || 0;
        totalQty += qty;

        const g = row["Garment Type"] || "Unspecified";
        if (!garmentAnalysis[g]) garmentAnalysis[g] = { lots: 0, qty: 0 };
        garmentAnalysis[g].lots += 1;
        garmentAnalysis[g].qty += qty;

        const sup = row.Supervisor || "Unassigned";
        if (!supervisorAnalysis[sup]) supervisorAnalysis[sup] = { lots: 0, qty: 0 };
        supervisorAnalysis[sup].lots += 1;
        supervisorAnalysis[sup].qty += qty;

        const party = row["Party Name"] || "Unassigned";
        if (!partyAnalysis[party]) partyAnalysis[party] = { lots: 0, qty: 0 };
        partyAnalysis[party].lots += 1;
        partyAnalysis[party].qty += qty;

        const sec = row.Section || "Unassigned";
        if (!sectionAnalysis[sec]) sectionAnalysis[sec] = { lots: 0, qty: 0 };
        sectionAnalysis[sec].lots += 1;
        sectionAnalysis[sec].qty += qty;
      });

      const workbook = new ExcelJS.Workbook();
      workbook.creator = "Factory Suite Pro";
      workbook.created = now;

      // Styling Helpers
      const thinBorder = {
        top: { style: 'thin', color: { argb: 'CBD5E1' } },
        left: { style: 'thin', color: { argb: 'CBD5E1' } },
        bottom: { style: 'thin', color: { argb: 'CBD5E1' } },
        right: { style: 'thin', color: { argb: 'CBD5E1' } }
      };

      const headerBorder = {
        top: { style: 'thin', color: { argb: '0F172A' } },
        left: { style: 'thin', color: { argb: '0F172A' } },
        bottom: { style: 'medium', color: { argb: '0F172A' } },
        right: { style: 'thin', color: { argb: '0F172A' } }
      };

      const totalBorder = {
        top: { style: 'thin', color: { argb: '0F172A' } },
        left: { style: 'thin', color: { argb: 'CBD5E1' } },
        bottom: { style: 'double', color: { argb: '0F172A' } },
        right: { style: 'thin', color: { argb: 'CBD5E1' } }
      };

      // ==========================================
      // SHEET 1: DAILY STITCHING ISSUE LOG
      // ==========================================
      const ws1 = workbook.addWorksheet("Stitching Issue Log", {
        views: [{ showGridLines: true, state: 'frozen', xSplit: 0, ySplit: 5 }]
      });

      const cols1 = [
        { header: "Sr. No", key: "sr", width: 8 },
        { header: "Lot Number", key: "lot", width: 16 },
        { header: "Garment Type", key: "garment", width: 18 },
        { header: "Style", key: "style", width: 18 },
        { header: "Fabric", key: "fabric", width: 22 },
        { header: "Brand", key: "brand", width: 16 },
        { header: "PCS", key: "pcs", width: 14 },
        { header: "Section", key: "section", width: 14 },
        { header: "Season", key: "season", width: 14 },
        { header: "Party Name", key: "party", width: 20 },
        { header: "Direct Stitching", key: "direct", width: 16 },
        { header: "Supervisor", key: "sup", width: 18 },
        { header: "Date of Issue", key: "date", width: 16 }
      ];

      const numCols1 = cols1.length;
      ws1.columns = cols1;

      // Row 1: Title Banner
      const titleRow1 = ws1.getRow(1);
      titleRow1.values = ["FACTORY SUITE PRO - DAILY STITCHING ISSUE REPORT"];
      ws1.mergeCells(1, 1, 1, numCols1);
      titleRow1.font = { name: "Segoe UI", size: 14, bold: true, color: { argb: "FFFFFFFF" } };
      titleRow1.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E1B4B" } };
      titleRow1.alignment = { vertical: "middle", horizontal: "center" };
      titleRow1.height = 34;

      // Row 2: Metadata Banner
      const metaRow1 = ws1.getRow(2);
      metaRow1.values = [`Exported on: ${reportDateStr}  |  Total Records: ${totalLots}  |  Total PCS: ${totalQty.toLocaleString()}  |  Active Supervisors: ${Object.keys(supervisorAnalysis).length}  |  Garments: ${Object.keys(garmentAnalysis).length}`];
      ws1.mergeCells(2, 1, 2, numCols1);
      metaRow1.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF1E293B" } };
      metaRow1.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } };
      metaRow1.alignment = { vertical: "middle", horizontal: "center" };
      metaRow1.height = 22;

      // Row 3: Blank Row
      const blankRow3 = ws1.getRow(3);
      blankRow3.values = [];
      blankRow3.height = 6;

      // Row 4: Blank Row
      const blankRow4 = ws1.getRow(4);
      blankRow4.values = [];
      blankRow4.height = 6;

      // Row 5: Table Headers Row
      const headerRow1 = ws1.getRow(5);
      headerRow1.values = cols1.map(c => c.header);
      headerRow1.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
      headerRow1.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF312E81" } };
      headerRow1.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      headerRow1.height = 28;
      for (let c = 1; c <= numCols1; c++) {
        headerRow1.getCell(c).border = headerBorder;
      }

      // Add Data Rows
      let rIdx1 = 6;
      filteredRows.forEach((row, index) => {
        const rowValues = [
          index + 1,
          row["Lot Number"] || "—",
          row["Garment Type"] || "—",
          row["Style"] || "—",
          row["Fabric"] || "—",
          row["Brand"] || "—",
          Number(row.PCS) || 0,
          row["Section"] || "—",
          row["Season"] || "—",
          row["Party Name"] || "—",
          row["Direct Stitching"] || "No",
          row.Supervisor || "—",
          row["Date of Issue"] || "—"
        ];

        const dataRow = ws1.getRow(rIdx1);
        dataRow.values = rowValues;
        dataRow.height = 22;

        const isEven = index % 2 === 0;
        const defaultBg = isEven ? "FFFFFFFF" : "FFF8FAFC";

        for (let c = 1; c <= numCols1; c++) {
          const cell = dataRow.getCell(c);
          cell.font = { name: "Segoe UI", size: 9.5, color: { argb: "FF0F172A" } };
          cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
          cell.border = thinBorder;
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: defaultBg } };

          // Lot No Bold
          if (c === 2) {
            cell.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF0F172A" } };
          }
          // PCS numeric format
          if (c === 7) {
            cell.numFmt = "#,##0";
            cell.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF0F172A" } };
          }
        }

        rIdx1++;
      });

      // Total Row for Sheet 1
      const totalRow1 = ws1.getRow(rIdx1);
      const totalValues1 = new Array(numCols1).fill("");
      totalValues1[0] = "TOTAL";
      totalValues1[6] = totalQty;
      totalValues1[numCols1 - 1] = `${totalLots} Lots`;
      totalRow1.values = totalValues1;
      totalRow1.height = 26;

      for (let c = 1; c <= numCols1; c++) {
        const cell = totalRow1.getCell(c);
        cell.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FF0F172A" } };
        cell.alignment = { vertical: "middle", horizontal: "center" };
        cell.border = totalBorder;
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE2E8F0" } };
        if (c === 7) cell.numFmt = "#,##0";
      }

      // ==========================================
      // SHEET 2: PRODUCTION SUMMARY & BREAKDOWNS
      // ==========================================
      const ws2 = workbook.addWorksheet("Issue Summary", {
        views: [{ showGridLines: true }]
      });

      ws2.columns = [
        { width: 32 },
        { width: 18 },
        { width: 18 },
        { width: 22 }
      ];

      // Sheet 2 Title Banner
      const titleRow2 = ws2.getRow(1);
      titleRow2.values = ["FACTORY SUITE PRO - DAILY STITCHING ISSUE BREAKDOWN"];
      ws2.mergeCells(1, 1, 1, 4);
      titleRow2.font = { name: "Segoe UI", size: 14, bold: true, color: { argb: "FFFFFFFF" } };
      titleRow2.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E1B4B" } };
      titleRow2.alignment = { vertical: "middle", horizontal: "center" };
      titleRow2.height = 32;

      let r2 = 3;

      // Section 1: Garment Type Breakdown
      const gTitle = ws2.getRow(r2);
      gTitle.values = ["1. GARMENT TYPE BREAKDOWN"];
      ws2.mergeCells(r2, 1, r2, 4);
      gTitle.font = { name: "Segoe UI", size: 10.5, bold: true, color: { argb: "FFFFFFFF" } };
      gTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F766E" } };
      gTitle.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
      gTitle.height = 24;
      r2++;

      const gHeaders = ws2.getRow(r2);
      gHeaders.values = ["Garment Type", "Total Lots", "Percentage (%)", "Total Quantity"];
      gHeaders.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFFFFFFF" } };
      gHeaders.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF14B8A6" } };
      gHeaders.alignment = { vertical: "middle", horizontal: "center" };
      gHeaders.height = 24;
      for (let c = 1; c <= 4; c++) gHeaders.getCell(c).border = headerBorder;
      r2++;

      const garmentArr = Object.entries(garmentAnalysis).map(([g, d]) => ({
        name: g,
        lots: d.lots,
        pct: totalLots > 0 ? Math.round((d.lots / totalLots) * 100) : 0,
        qty: d.qty
      })).sort((a, b) => b.lots - a.lots);

      garmentArr.forEach((g, idx) => {
        const row = ws2.getRow(r2);
        row.values = [g.name, g.lots, `${g.pct}%`, g.qty];
        row.height = 22;
        const bg = idx % 2 === 0 ? "FFFFFFFF" : "FFF8FAFC";
        for (let c = 1; c <= 4; c++) {
          const cell = row.getCell(c);
          cell.font = { name: "Segoe UI", size: 9.5, color: { argb: "FF0F172A" } };
          cell.alignment = { vertical: "middle", horizontal: c === 1 ? "left" : "center", indent: c === 1 ? 1 : 0 };
          cell.border = thinBorder;
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bg } };
          if (c === 4) cell.numFmt = "#,##0";
        }
        r2++;
      });
      r2 += 2;

      // Section 2: Supervisor Production Breakdown
      const sTitle = ws2.getRow(r2);
      sTitle.values = ["2. SUPERVISOR PRODUCTION BREAKDOWN"];
      ws2.mergeCells(r2, 1, r2, 4);
      sTitle.font = { name: "Segoe UI", size: 10.5, bold: true, color: { argb: "FFFFFFFF" } };
      sTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4338CA" } };
      sTitle.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
      sTitle.height = 24;
      r2++;

      const sHeaders = ws2.getRow(r2);
      sHeaders.values = ["Supervisor", "Total Lots", "Percentage (%)", "Total Quantity"];
      sHeaders.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFFFFFFF" } };
      sHeaders.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF6366F1" } };
      sHeaders.alignment = { vertical: "middle", horizontal: "center" };
      sHeaders.height = 24;
      for (let c = 1; c <= 4; c++) sHeaders.getCell(c).border = headerBorder;
      r2++;

      const supArr = Object.entries(supervisorAnalysis).map(([s, d]) => ({
        name: s,
        lots: d.lots,
        pct: totalLots > 0 ? Math.round((d.lots / totalLots) * 100) : 0,
        qty: d.qty
      })).sort((a, b) => b.lots - a.lots);

      supArr.forEach((s, idx) => {
        const row = ws2.getRow(r2);
        row.values = [s.name, s.lots, `${s.pct}%`, s.qty];
        row.height = 22;
        const bg = idx % 2 === 0 ? "FFFFFFFF" : "FFF8FAFC";
        for (let c = 1; c <= 4; c++) {
          const cell = row.getCell(c);
          cell.font = { name: "Segoe UI", size: 9.5, color: { argb: "FF0F172A" } };
          cell.alignment = { vertical: "middle", horizontal: c === 1 ? "left" : "center", indent: c === 1 ? 1 : 0 };
          cell.border = thinBorder;
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bg } };
          if (c === 4) cell.numFmt = "#,##0";
        }
        r2++;
      });
      r2 += 2;

      // ==========================================
      // SHEET 3: APPLIED FILTERS (AUDIT SHEET)
      // ==========================================
      const ws3 = workbook.addWorksheet("Applied Filters", {
        views: [{ showGridLines: true }]
      });
      ws3.columns = [{ width: 28 }, { width: 65 }];

      const fTitle = ws3.getRow(1);
      fTitle.values = ["REPORT FILTERS & AUDIT PARAMETERS"];
      ws3.mergeCells(1, 1, 1, 2);
      fTitle.font = { name: "Segoe UI", size: 12, bold: true, color: { argb: "FFFFFFFF" } };
      fTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF334155" } };
      fTitle.alignment = { vertical: "middle", horizontal: "center" };
      fTitle.height = 28;

      const fHead = ws3.getRow(2);
      fHead.values = ["Filter Parameter", "Selected Value / Criteria"];
      fHead.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
      fHead.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF475569" } };
      fHead.alignment = { vertical: "middle", horizontal: "center" };
      fHead.height = 24;
      fHead.getCell(1).border = headerBorder;
      fHead.getCell(2).border = headerBorder;

      const filterAudit = [
        ["Export Timestamp", reportDateStr],
        ["Lot Filter", filterLot || "All"],
        ["Supervisor Filter", filterSupervisor || "All"],
        ["Issue Date Filter", filterIssueDate || "All"],
        ["Brand Filter", filterBrand || "All"],
        ["Garment Filter", filterGarment || "All"],
        ["Date Range", dateRangeFilter.enabled ? `${dateRangeFilter.startDate || ""} to ${dateRangeFilter.endDate || ""}` : "All"],
        ["Search Keyword", searchTerm || "None"],
        ["Recent Lots Only", showOnlyRecent ? "Yes (Last 24h)" : "No"]
      ];

      filterAudit.forEach(([param, val], idx) => {
        const row = ws3.getRow(idx + 3);
        row.values = [param, val];
        row.height = 22;
        const bg = idx % 2 === 0 ? "FFFFFFFF" : "FFF8FAFC";
        for (let c = 1; c <= 2; c++) {
          const cell = row.getCell(c);
          cell.font = { name: "Segoe UI", size: 9.5, color: { argb: "FF0F172A" }, bold: c === 1 };
          cell.alignment = { vertical: "middle", horizontal: c === 1 ? "left" : "center", indent: c === 1 ? 1 : 0 };
          cell.border = thinBorder;
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bg } };
        }
      });

      const fileName = `Daily_Stitching_Issues_${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      saveAs(blob, `${fileName}.xlsx`);
    } catch (err) {
      console.error("Error exporting Excel:", err);
      alert(`Failed to export Excel: ${err.message}`);
    }
  };

  /* ---------- Direct PDF Export via jsPDF (Factory Suite Pro A3 Landscape) ---------- */
  const handleExportPDF = () => {
    try {
      if (filteredRows.length === 0) {
        alert("No data available to export.");
        return;
      }

      const totalLots = filteredRows.length;
      let totalQty = 0;
      const garmentMap = {};
      const supervisorMap = {};
      const partyMap = {};
      const sectionMap = {};

      filteredRows.forEach(row => {
        const qty = Number(row.PCS) || 0;
        totalQty += qty;

        const g = row["Garment Type"] || "Unspecified";
        if (!garmentMap[g]) garmentMap[g] = { lots: 0, qty: 0 };
        garmentMap[g].lots += 1;
        garmentMap[g].qty += qty;

        const s = row.Supervisor || "Unassigned";
        if (!supervisorMap[s]) supervisorMap[s] = { lots: 0, qty: 0 };
        supervisorMap[s].lots += 1;
        supervisorMap[s].qty += qty;

        const p = row["Party Name"] || "Unassigned";
        if (!partyMap[p]) partyMap[p] = { lots: 0, qty: 0 };
        partyMap[p].lots += 1;
        partyMap[p].qty += qty;

        const sec = row.Section || "Unassigned";
        if (!sectionMap[sec]) sectionMap[sec] = { lots: 0, qty: 0 };
        sectionMap[sec].lots += 1;
        sectionMap[sec].qty += qty;
      });

      const sortedGarments = Object.keys(garmentMap).map(name => ({
        name,
        lots: garmentMap[name].lots,
        qty: garmentMap[name].qty
      })).sort((a, b) => b.qty - a.qty);

      const sortedSupervisors = Object.keys(supervisorMap).map(name => ({
        name,
        lots: supervisorMap[name].lots,
        qty: supervisorMap[name].qty
      })).sort((a, b) => b.qty - a.qty);

      const sortedParties = Object.keys(partyMap).map(name => ({
        name,
        lots: partyMap[name].lots,
        qty: partyMap[name].qty
      })).sort((a, b) => b.qty - a.qty);

      const sortedSections = Object.keys(sectionMap).map(name => ({
        name,
        lots: sectionMap[name].lots,
        qty: sectionMap[name].qty
      })).sort((a, b) => b.qty - a.qty);

      // Create PDF in A3 Landscape
      const doc = new jsPDF({
        orientation: "landscape",
        unit: "pt",
        format: "a3"
      });

      const pageW = doc.internal.pageSize.getWidth();
      const pageH = doc.internal.pageSize.getHeight();

      // 1. Top Header Banner
      doc.setFillColor(30, 27, 75); // Dark Indigo #1E1B4B
      doc.rect(15, 12, pageW - 30, 48, 'F');

      doc.setFontSize(14);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(255, 255, 255);
      doc.text("FACTORY SUITE PRO - DAILY STITCHING ISSUE REPORT", pageW / 2, 30, { align: 'center' });

      doc.setFontSize(8.5);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(199, 210, 254);
      const subText = `Total Lots: ${totalLots}   |   Total PCS: ${totalQty.toLocaleString()} Pcs   |   Supervisors: ${sortedSupervisors.length}   |   Parties: ${sortedParties.length}   |   Garments: ${sortedGarments.length}`;
      doc.text(subText, pageW / 2, 48, { align: 'center' });

      // 2. Filter Banner
      doc.setFillColor(241, 245, 249);
      doc.rect(15, 63, pageW - 30, 16, 'F');
      doc.setFontSize(8);
      doc.setFont('helvetica', 'italic');
      doc.setTextColor(0, 0, 0);
      const filterSummary = `Filters: Lot: ${filterLot || 'All'} | Supervisor: ${filterSupervisor || 'All'} | Issue Date: ${filterIssueDate || 'All'} | Brand: ${filterBrand || 'All'} | Garment: ${filterGarment || 'All'} | Date Range: ${dateRangeFilter.enabled ? `${dateRangeFilter.startDate || ''} to ${dateRangeFilter.endDate || ''}` : 'All'} | Search: ${searchTerm || 'None'}`;
      doc.text(filterSummary, pageW / 2, 74, { align: 'center' });

      // 3. Main Data Table
      const tableColumns = [
        '#',
        'Lot Number',
        'Garment Type',
        'Style',
        'Fabric',
        'Brand',
        'PCS',
        'Section',
        'Season',
        'Party Name',
        'Direct Stitching',
        'Supervisor',
        'Date of Issue'
      ];

      const tableBody = filteredRows.map((row, idx) => {
        const lotNo = (row["Lot Number"] || "").toString().trim();
        const isRecent = row._isRecent;
        const lotDisplay = isRecent ? `* ${lotNo}` : (lotNo || "—");

        return [
          (idx + 1).toString(),
          lotDisplay,
          row["Garment Type"] || "—",
          row["Style"] || "—",
          row["Fabric"] || "—",
          row["Brand"] || "—",
          (Number(row.PCS) || 0).toLocaleString(),
          row["Section"] || "—",
          row["Season"] || "—",
          row["Party Name"] || "—",
          row["Direct Stitching"] || "No",
          row.Supervisor || "—",
          row["Date of Issue"] || "—"
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
        totalQty.toLocaleString(),
        '',
        '',
        `${sortedParties.length} Parties`,
        '',
        `${sortedSupervisors.length} Supervisors`,
        ''
      ]);

      const columnStyles = {
        0: { cellWidth: 35, halign: 'center' },
        1: { cellWidth: 85, halign: 'center', fontStyle: 'bold' },
        2: { cellWidth: 105, halign: 'center' },
        3: { cellWidth: 110, halign: 'center' },
        4: { cellWidth: 115, halign: 'center' },
        5: { cellWidth: 85, halign: 'center' },
        6: { cellWidth: 70, halign: 'center', fontStyle: 'bold' },
        7: { cellWidth: 75, halign: 'center' },
        8: { cellWidth: 75, halign: 'center' },
        9: { cellWidth: 105, halign: 'center' },
        10: { cellWidth: 75, halign: 'center' },
        11: { cellWidth: 100, halign: 'center' },
        12: { cellWidth: 85, halign: 'center' }
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
          textColor: [0, 0, 0], // Pure black text
          lineColor: [0, 0, 0],
          lineWidth: 0.3,
          fontStyle: 'normal',
          minCellHeight: 14,
        },
        headStyles: {
          fillColor: [30, 27, 75],
          textColor: [255, 255, 255],
          fontStyle: "bold",
          lineColor: [0, 0, 0],
          lineWidth: 0.5,
          halign: 'center',
          fontSize: 9,
          valign: 'middle',
          cellPadding: { top: 5, right: 2, bottom: 5, left: 2 },
        },
        alternateRowStyles: {
          fillColor: [248, 250, 252],
        },
        columnStyles,
        didParseCell: function (data) {
          if (data.section === 'body') {
            data.cell.styles.textColor = [0, 0, 0]; // Pure black

            const rowIndex = data.row.index;
            const isTotalRow = rowIndex === tableBody.length - 1;

            if (isTotalRow) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [226, 232, 240];
              data.cell.styles.textColor = [0, 0, 0];
              data.cell.styles.halign = 'center';
            }
          }
        }
      });

      // --- 4-COLUMN SIDE-BY-SIDE EXECUTIVE SUMMARY ---
      const gBody = sortedGarments.map(item => {
        const pct = totalQty > 0 ? ((item.qty / totalQty) * 100).toFixed(1) : "0.0";
        return [item.name, item.lots.toString(), item.qty.toLocaleString(), `${pct}%`];
      });
      gBody.push(["TOTAL", totalLots.toString(), totalQty.toLocaleString(), "100.0%"]);

      const supBody = sortedSupervisors.map(item => {
        const pct = totalQty > 0 ? ((item.qty / totalQty) * 100).toFixed(1) : "0.0";
        return [item.name, item.lots.toString(), item.qty.toLocaleString(), `${pct}%`];
      });
      supBody.push(["TOTAL", totalLots.toString(), totalQty.toLocaleString(), "100.0%"]);

      const pBody = sortedParties.map(item => {
        const pct = totalQty > 0 ? ((item.qty / totalQty) * 100).toFixed(1) : "0.0";
        return [item.name, item.lots.toString(), item.qty.toLocaleString(), `${pct}%`];
      });
      pBody.push(["TOTAL", totalLots.toString(), totalQty.toLocaleString(), "100.0%"]);

      const sBody = sortedSections.map(item => {
        const pct = totalQty > 0 ? ((item.qty / totalQty) * 100).toFixed(1) : "0.0";
        return [item.name, item.lots.toString(), item.qty.toLocaleString(), `${pct}%`];
      });
      sBody.push(["TOTAL", totalLots.toString(), totalQty.toLocaleString(), "100.0%"]);

      const maxRows = Math.max(gBody.length, supBody.length, pBody.length, sBody.length);
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
      doc.text("EXECUTIVE SUMMARY & STITCHING ISSUE BREAKDOWN", pageW / 2, summaryStartY + 4, { align: 'center' });

      doc.setFontSize(9);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(0, 0, 0);
      const summarySub = `Total Lots: ${totalLots}   |   Total Quantity: ${totalQty.toLocaleString()} Pcs   |   Supervisors: ${sortedSupervisors.length}   |   Garments: ${sortedGarments.length}   |   Parties: ${sortedParties.length}`;
      doc.text(summarySub, pageW / 2, summaryStartY + 16, { align: 'center' });

      const sectionTitleY = summaryStartY + 30;
      const tableStartY = sectionTitleY + 6;

      const colWidth = 278;
      const gap = 16;
      const col1X = 15;
      const col2X = col1X + colWidth + gap;
      const col3X = col2X + colWidth + gap;
      const col4X = col3X + colWidth + gap;

      doc.setFontSize(9.5);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(0, 0, 0);
      doc.text("1. GARMENT BREAKDOWN", col1X, sectionTitleY);
      doc.text("2. SUPERVISOR BREAKDOWN", col2X, sectionTitleY);
      doc.text("3. PARTY BREAKDOWN", col3X, sectionTitleY);
      doc.text("4. SECTION BREAKDOWN", col4X, sectionTitleY);

      const summaryColStyles = {
        0: { cellWidth: 110, halign: 'center' },
        1: { cellWidth: 45, halign: 'center' },
        2: { cellWidth: 68, halign: 'center' },
        3: { cellWidth: 55, halign: 'center' },
      };

      // Col 1: Garments
      autoTable(doc, {
        head: [['Garment Type', 'Lots', 'Total Qty', 'Share %']],
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

      // Col 2: Supervisors
      autoTable(doc, {
        head: [['Supervisor', 'Lots', 'Total Qty', 'Share %']],
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

      // Col 3: Parties
      autoTable(doc, {
        head: [['Party Name', 'Lots', 'Total Qty', 'Share %']],
        body: pBody,
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
            if (data.row.index === pBody.length - 1) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });

      // Col 4: Sections
      autoTable(doc, {
        head: [['Section', 'Lots', 'Total Qty', 'Share %']],
        body: sBody,
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
            if (data.row.index === sBody.length - 1) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });

      const fileName = `Daily_Stitching_Issues_${new Date().toISOString().split('T')[0]}.pdf`;
      doc.save(fileName);
    } catch (err) {
      console.error("Error generating PDF:", err);
      alert(`Failed to generate PDF: ${err.message}`);
    }
  };

  /* ---------- Pagination ---------- */
  const totalPages = Math.max(1, Math.ceil(filteredRows.length / pageSize));
  const pagedRows = useMemo(() => {
    const start = (page - 1) * pageSize;
    return filteredRows.slice(start, start + pageSize);
  }, [filteredRows, page, pageSize]);

  return (
    <div style={{ minHeight: "100vh", backgroundColor: "#f8fafc", color: "#0f172a", fontFamily: "'Inter', system-ui, -apple-system, sans-serif" }}>
      {/* Top Navbar */}
      <header style={{
        background: "#ffffff",
        borderBottom: "1px solid #e2e8f0",
        padding: "12px 28px",
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        boxShadow: "0 2px 8px rgba(0,0,0,0.04)",
        position: "sticky",
        top: 0,
        zIndex: 50
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: "14px" }}>
          <button
            onClick={handleGoBack}
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: "6px",
              background: "#f1f5f9",
              border: "1.5px solid #cbd5e1",
              color: "#1e293b",
              padding: "7px 14px",
              borderRadius: "8px",
              fontSize: "0.82rem",
              fontWeight: 800,
              cursor: "pointer",
              transition: "all 0.15s ease"
            }}
            title="Go back to previous page"
          >
            ← Back
          </button>
          <Link to="/dashboard" style={{ display: "flex", alignItems: "center", gap: "10px", textDecoration: "none" }}>
            <span style={{ fontSize: "1.6rem" }}>🏭</span>
            <div>
              <div style={{ fontSize: "1.05rem", fontWeight: 900, color: "#0f172a", letterSpacing: "-0.3px", display: "flex", alignItems: "center", gap: "6px" }}>
                Factory Suite Pro <span style={{ background: "#dcfce7", color: "#15803d", fontSize: "0.68rem", fontWeight: 800, padding: "2px 6px", borderRadius: "4px" }}>LIVE</span>
              </div>
              <div style={{ fontSize: "0.72rem", color: "#64748b", fontWeight: 600 }}>Daily Stitching Issue Console</div>
            </div>
          </Link>
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
          <Link to="/dashboard" style={{
            display: "inline-flex",
            alignItems: "center",
            gap: "6px",
            background: "#1e1b4b",
            color: "#ffffff",
            padding: "7px 14px",
            borderRadius: "8px",
            fontSize: "0.82rem",
            fontWeight: 700,
            textDecoration: "none"
          }}>
            🏢 Dashboard
          </Link>
          <Link to="/lot-logs" style={{
            display: "inline-flex",
            alignItems: "center",
            gap: "6px",
            background: "#f1f5f9",
            color: "#334155",
            padding: "7px 14px",
            borderRadius: "8px",
            fontSize: "0.82rem",
            fontWeight: 700,
            textDecoration: "none"
          }}>
            📋 Lot Logs
          </Link>
          <Link to="/production-flowchart" style={{
            display: "inline-flex",
            alignItems: "center",
            gap: "6px",
            background: "#f1f5f9",
            color: "#334155",
            padding: "7px 14px",
            borderRadius: "8px",
            fontSize: "0.82rem",
            fontWeight: 700,
            textDecoration: "none"
          }}>
            🗺️ Flow Poster
          </Link>
          <Link to="/lot-timeline" style={{
            display: "inline-flex",
            alignItems: "center",
            gap: "6px",
            background: "#f1f5f9",
            color: "#334155",
            padding: "7px 14px",
            borderRadius: "8px",
            fontSize: "0.82rem",
            fontWeight: 700,
            textDecoration: "none"
          }}>
            🚚 Lot Tracker
          </Link>
          {currentUser && (
            <button
              onClick={handleLogout}
              style={{
                background: "#fee2e2",
                border: "1px solid #fca5a5",
                color: "#dc2626",
                padding: "7px 14px",
                borderRadius: "8px",
                fontSize: "0.82rem",
                fontWeight: 700,
                cursor: "pointer"
              }}
            >
              🚪 Logout
            </button>
          )}
        </div>
      </header>

      {/* Main Container */}
      <main style={{ padding: "24px 32px 60px 32px", maxWidth: "100%", boxSizing: "border-box" }}>
        {/* Header Banner */}
        <div style={{
          background: "linear-gradient(135deg, #1e1b4b 0%, #312e81 40%, #4338ca 100%)",
          borderRadius: "20px",
          padding: "26px 32px",
          marginBottom: "24px",
          border: "1px solid rgba(255, 255, 255, 0.12)",
          boxShadow: "0 16px 32px -10px rgba(30, 27, 75, 0.3)",
          color: "#ffffff",
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          flexWrap: "wrap",
          gap: "20px"
        }}>
          <div>
            <h1 style={{ margin: 0, fontSize: "1.85rem", fontWeight: 900, letterSpacing: "-0.5px", display: "flex", alignItems: "center", gap: "10px" }}>
              <span>🪡</span> DAILY STITCHING ISSUE
            </h1>
            <p style={{ margin: "6px 0 0 0", color: "#c7d2fe", fontSize: "0.92rem", fontWeight: 500 }}>
              Live stitching line allotment records, supervisor targets & brand allocations
            </p>
            {lastUpdated && (
              <div style={{ color: "#a5b4fc", fontSize: "0.76rem", marginTop: "6px", fontWeight: 600 }}>
                Synced: {lastUpdated}
              </div>
            )}
          </div>

          {/* KPI Stat Cards */}
          <div style={{ display: "flex", gap: "14px", flexWrap: "wrap" }}>
            <div style={{
              background: "rgba(255, 255, 255, 0.12)",
              backdropFilter: "blur(10px)",
              padding: "10px 18px",
              borderRadius: "14px",
              border: "1px solid rgba(255, 255, 255, 0.2)",
              textAlign: "center",
              minWidth: "100px"
            }}>
              <div style={{ fontSize: "1.45rem", fontWeight: 900, color: "#ffffff", lineHeight: 1.1 }}>
                {analytics.totalRecords.toLocaleString()}
              </div>
              <div style={{ fontSize: "0.68rem", color: "#e0e7ff", fontWeight: 800, textTransform: "uppercase", marginTop: "2px" }}>
                Total Records
              </div>
            </div>
            {/* 
            <div style={{
              background: "rgba(255, 255, 255, 0.12)",
              backdropFilter: "blur(10px)",
              padding: "10px 18px",
              borderRadius: "14px",
              border: "1px solid rgba(255, 255, 255, 0.2)",
              textAlign: "center",
              minWidth: "100px"
            }}>
              <div style={{ fontSize: "1.45rem", fontWeight: 900, color: "#a7f3d0", lineHeight: 1.1 }}>
                {analytics.totalPCS.toLocaleString()}
              </div>
              <div style={{ fontSize: "0.68rem", color: "#e0e7ff", fontWeight: 800, textTransform: "uppercase", marginTop: "2px" }}>
                Total PCS
              </div>
            </div> */}

            <div style={{
              background: "rgba(255, 255, 255, 0.12)",
              backdropFilter: "blur(10px)",
              padding: "10px 18px",
              borderRadius: "14px",
              border: "1px solid rgba(255, 255, 255, 0.2)",
              textAlign: "center",
              minWidth: "100px"
            }}>
              <div style={{ fontSize: "1.45rem", fontWeight: 900, color: "#ffffff", lineHeight: 1.1 }}>
                {analytics.uniqueLots}
              </div>
              <div style={{ fontSize: "0.68rem", color: "#e0e7ff", fontWeight: 800, textTransform: "uppercase", marginTop: "2px" }}>
                Active Lots
              </div>
            </div>

            <div style={{
              background: "rgba(255, 255, 255, 0.12)",
              backdropFilter: "blur(10px)",
              padding: "10px 18px",
              borderRadius: "14px",
              border: "1px solid rgba(255, 255, 255, 0.2)",
              textAlign: "center",
              minWidth: "100px"
            }}>
              <div style={{ fontSize: "1.45rem", fontWeight: 900, color: "#fed7aa", lineHeight: 1.1 }}>
                {analytics.supervisorsCount}
              </div>
              <div style={{ fontSize: "0.68rem", color: "#e0e7ff", fontWeight: 800, textTransform: "uppercase", marginTop: "2px" }}>
                Supervisors
              </div>
            </div>
          </div>
        </div>

        {/* Toolbar & Filters */}
        <div style={{
          background: "#ffffff",
          borderRadius: "18px",
          padding: "20px 24px",
          marginBottom: "24px",
          border: "1px solid #e2e8f0",
          boxShadow: "0 8px 24px rgba(0, 0, 0, 0.03)"
        }}>
          {/* Search Box */}
          <div style={{ position: "relative", marginBottom: "16px" }}>
            <span style={{ position: "absolute", left: "14px", top: "50%", transform: "translateY(-50%)", color: "#64748b", fontSize: "16px" }}>🔍</span>
            <input
              type="text"
              placeholder="Search across lot numbers, garments, styles, fabrics, brands, supervisors..."
              value={searchTerm}
              onChange={(e) => { setSearchTerm(e.target.value); setPage(1); }}
              style={{
                width: "100%",
                padding: "11px 16px 11px 44px",
                border: "1.5px solid #cbd5e1",
                borderRadius: "10px",
                fontSize: "0.88rem",
                fontWeight: 600,
                color: "#0f172a",
                boxSizing: "border-box",
                outline: "none"
              }}
            />
          </div>

          {/* Filter Dropdowns Grid */}
          <div style={{
            display: "grid",
            gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))",
            gap: "14px",
            marginBottom: "16px"
          }}>
            <div>
              <label style={{ fontSize: "0.76rem", fontWeight: 800, color: "#475569", textTransform: "uppercase", marginBottom: "4px", display: "block" }}>
                Lot Number
              </label>
              <select
                value={filterLot}
                onChange={(e) => { setFilterLot(e.target.value); setPage(1); }}
                style={{ width: "100%", padding: "8px 12px", borderRadius: "8px", border: "1.5px solid #cbd5e1", fontSize: "0.82rem", fontWeight: 600, background: "#ffffff", color: "#0f172a" }}
              >
                <option value="">All Lots ({uniqueLots.length})</option>
                {uniqueLots.map((lot) => (
                  <option key={lot} value={lot}>{lot}</option>
                ))}
              </select>
            </div>

            <div>
              <label style={{ fontSize: "0.76rem", fontWeight: 800, color: "#475569", textTransform: "uppercase", marginBottom: "4px", display: "block" }}>
                Supervisor
              </label>
              <select
                value={filterSupervisor}
                onChange={(e) => { setFilterSupervisor(e.target.value); setPage(1); }}
                style={{ width: "100%", padding: "8px 12px", borderRadius: "8px", border: "1.5px solid #cbd5e1", fontSize: "0.82rem", fontWeight: 600, background: "#ffffff", color: "#0f172a" }}
              >
                <option value="">All Supervisors ({uniqueSupervisors.length})</option>
                {uniqueSupervisors.map(({ key, label }) => (
                  <option key={key} value={key}>{label}</option>
                ))}
              </select>
            </div>

            <div>
              <label style={{ fontSize: "0.76rem", fontWeight: 800, color: "#475569", textTransform: "uppercase", marginBottom: "4px", display: "block" }}>
                Garment Type
              </label>
              <select
                value={filterGarment}
                onChange={(e) => { setFilterGarment(e.target.value); setPage(1); }}
                style={{ width: "100%", padding: "8px 12px", borderRadius: "8px", border: "1.5px solid #cbd5e1", fontSize: "0.82rem", fontWeight: 600, background: "#ffffff", color: "#0f172a" }}
              >
                <option value="">All Garments ({uniqueGarments.length})</option>
                {uniqueGarments.map((g) => (
                  <option key={g} value={g}>{g}</option>
                ))}
              </select>
            </div>

            <div>
              <label style={{ fontSize: "0.76rem", fontWeight: 800, color: "#475569", textTransform: "uppercase", marginBottom: "4px", display: "block" }}>
                Brand
              </label>
              <select
                value={filterBrand}
                onChange={(e) => { setFilterBrand(e.target.value); setPage(1); }}
                style={{ width: "100%", padding: "8px 12px", borderRadius: "8px", border: "1.5px solid #cbd5e1", fontSize: "0.82rem", fontWeight: 600, background: "#ffffff", color: "#0f172a" }}
              >
                <option value="">All Brands ({uniqueBrands.length})</option>
                {uniqueBrands.map((b) => (
                  <option key={b} value={b}>{b}</option>
                ))}
              </select>
            </div>

            <div>
              <label style={{ fontSize: "0.76rem", fontWeight: 800, color: "#475569", textTransform: "uppercase", marginBottom: "4px", display: "block" }}>
                Issue Date
              </label>
              <select
                value={filterIssueDate}
                onChange={(e) => { setFilterIssueDate(e.target.value); setPage(1); }}
                style={{ width: "100%", padding: "8px 12px", borderRadius: "8px", border: "1.5px solid #cbd5e1", fontSize: "0.82rem", fontWeight: 600, background: "#ffffff", color: "#0f172a" }}
              >
                <option value="">All Dates ({uniqueIssueDates.length})</option>
                {uniqueIssueDates.map((d) => (
                  <option key={d} value={d}>{d}</option>
                ))}
              </select>
            </div>
          </div>

          {/* Action Buttons & Export Controls */}
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "12px", borderTop: "1px solid #f1f5f9", paddingTop: "14px" }}>
            <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
              <button
                onClick={handleExportExcel}
                style={{
                  background: "#059669",
                  color: "#ffffff",
                  border: "none",
                  padding: "9px 16px",
                  borderRadius: "8px",
                  fontSize: "0.84rem",
                  fontWeight: 800,
                  cursor: "pointer",
                  display: "inline-flex",
                  alignItems: "center",
                  gap: "6px"
                }}
              >
                📊 Export Excel (.xlsx)
              </button>

              <button
                onClick={handleExportPDF}
                style={{
                  background: "#dc2626",
                  color: "#ffffff",
                  border: "none",
                  padding: "9px 16px",
                  borderRadius: "8px",
                  fontSize: "0.84rem",
                  fontWeight: 800,
                  cursor: "pointer",
                  display: "inline-flex",
                  alignItems: "center",
                  gap: "6px"
                }}
              >
                📄 PDF Report (.pdf)
              </button>

              <button
                onClick={handleRefresh}
                style={{
                  background: "#4f46e5",
                  color: "#ffffff",
                  border: "none",
                  padding: "9px 16px",
                  borderRadius: "8px",
                  fontSize: "0.84rem",
                  fontWeight: 800,
                  cursor: "pointer",
                  display: "inline-flex",
                  alignItems: "center",
                  gap: "6px"
                }}
              >
                🔄 {refreshing ? "Refreshing..." : "Refresh Live"}
              </button>

              <button
                onClick={handleGoBack}
                style={{
                  background: "#334155",
                  color: "#ffffff",
                  border: "none",
                  padding: "9px 16px",
                  borderRadius: "8px",
                  fontSize: "0.84rem",
                  fontWeight: 800,
                  cursor: "pointer",
                  display: "inline-flex",
                  alignItems: "center",
                  gap: "6px"
                }}
                title="Go back to previous page"
              >
                ← Back
              </button>

              {hasActiveFilters && (
                <button
                  onClick={clearFilters}
                  style={{
                    background: "#f1f5f9",
                    color: "#475569",
                    border: "1px solid #cbd5e1",
                    padding: "9px 16px",
                    borderRadius: "8px",
                    fontSize: "0.84rem",
                    fontWeight: 800,
                    cursor: "pointer"
                  }}
                >
                  Clear Filters
                </button>
              )}
            </div>

            <div style={{ fontSize: "0.84rem", color: "#64748b", fontWeight: 700 }}>
              Showing {filteredRows.length === 0 ? 0 : (page - 1) * pageSize + 1} - {Math.min(page * pageSize, filteredRows.length)} of {filteredRows.length} lots
            </div>
          </div>
        </div>

        {/* Data Table */}
        <div style={{
          background: "#ffffff",
          borderRadius: "18px",
          border: "1px solid #e2e8f0",
          boxShadow: "0 8px 24px rgba(0, 0, 0, 0.03)",
          overflow: "hidden"
        }}>
          <div style={{ overflowX: "auto", width: "100%" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", textAlign: "center", fontSize: "0.84rem" }}>
              <thead>
                <tr style={{ background: "#1e1b4b", color: "#ffffff" }}>
                  {DISPLAY_HEADERS.map((h) => (
                    <th
                      key={h}
                      onClick={() => handleSort(h)}
                      style={{
                        padding: "12px 10px",
                        fontWeight: 800,
                        fontSize: "0.78rem",
                        textTransform: "uppercase",
                        letterSpacing: "0.5px",
                        borderRight: "1px solid rgba(255, 255, 255, 0.1)",
                        cursor: "pointer",
                        userSelect: "none",
                        whiteSpace: "nowrap"
                      }}
                    >
                      <div style={{ display: "inline-flex", alignItems: "center", gap: "6px" }}>
                        <span>{COLUMN_ICONS[h]}</span>
                        <span>{h}</span>
                        {sortConfig.key === h && (
                          <span style={{ color: "#38bdf8" }}>{sortConfig.direction === "asc" ? "▲" : "▼"}</span>
                        )}
                      </div>
                    </th>
                  ))}
                </tr>
              </thead>

              <tbody>
                {loading ? (
                  <tr>
                    <td colSpan={DISPLAY_HEADERS.length} style={{ padding: "60px 20px", textAlign: "center" }}>
                      <div style={{ fontSize: "2rem", marginBottom: "8px" }}>⏳</div>
                      <div style={{ fontSize: "1rem", fontWeight: 800, color: "#1e1b4b" }}>{loadingMessage}</div>
                      <div style={{ fontSize: "0.82rem", color: "#64748b" }}>Progress: {loadingProgress}%</div>
                    </td>
                  </tr>
                ) : pagedRows.length === 0 ? (
                  <tr>
                    <td colSpan={DISPLAY_HEADERS.length} style={{ padding: "60px 20px", textAlign: "center" }}>
                      <div style={{ fontSize: "2.5rem", marginBottom: "8px" }}>📭</div>
                      <div style={{ fontSize: "1.1rem", fontWeight: 800, color: "#0f172a" }}>No Stitching Issue Records Found</div>
                      <p style={{ color: "#64748b", margin: "4px 0 14px 0" }}>Try clearing some filters or searching for another lot number.</p>
                      {hasActiveFilters && (
                        <button
                          onClick={clearFilters}
                          style={{
                            background: "#4f46e5",
                            color: "#ffffff",
                            border: "none",
                            padding: "8px 18px",
                            borderRadius: "8px",
                            fontSize: "0.84rem",
                            fontWeight: 700,
                            cursor: "pointer"
                          }}
                        >
                          Clear all filters
                        </button>
                      )}
                    </td>
                  </tr>
                ) : (
                  pagedRows.map((row, idx) => {
                    const serialNumber = (page - 1) * pageSize + idx + 1;
                    const isEven = idx % 2 === 0;

                    return (
                      <tr
                        key={`${row["Lot Number"]}-${idx}-${row["Source Type"]}`}
                        style={{
                          background: row._isRecent ? "#f0fdf4" : (isEven ? "#ffffff" : "#f8fafc"),
                          borderBottom: "1px solid #e2e8f0",
                          transition: "background-color 0.15s"
                        }}
                        onMouseEnter={(e) => e.currentTarget.style.backgroundColor = "#f1f5f9"}
                        onMouseLeave={(e) => e.currentTarget.style.backgroundColor = row._isRecent ? "#f0fdf4" : (isEven ? "#ffffff" : "#f8fafc")}
                      >
                        {/* 1. Sr. No */}
                        <td style={{ padding: "10px 8px", fontWeight: 700, color: "#64748b", borderRight: "1px solid #f1f5f9" }}>
                          {serialNumber}
                        </td>

                        {/* 2. Lot Number */}
                        <td style={{ padding: "10px 8px", fontWeight: 800, color: "#0f172a", borderRight: "1px solid #f1f5f9", whiteSpace: "nowrap" }}>
                          <div style={{ display: "inline-flex", alignItems: "center", gap: "6px" }}>
                            <span>{row["Lot Number"]}</span>
                            {row._isRecent && (
                              <span style={{
                                background: "#dcfce7",
                                color: "#15803d",
                                border: "1px solid #86efac",
                                fontSize: "0.65rem",
                                fontWeight: 900,
                                padding: "1px 5px",
                                borderRadius: "4px"
                              }}>
                                NEW
                              </span>
                            )}
                          </div>
                        </td>

                        {/* 3. Garment Type */}
                        <td style={{ padding: "10px 8px", fontWeight: 600, color: "#334155", borderRight: "1px solid #f1f5f9", whiteSpace: "nowrap" }}>
                          {row["Garment Type"] || "—"}
                        </td>

                        {/* 4. Style */}
                        <td style={{ padding: "10px 8px", fontWeight: 600, color: "#334155", borderRight: "1px solid #f1f5f9", whiteSpace: "nowrap" }}>
                          {row["Style"] || "—"}
                        </td>

                        {/* 5. Fabric */}
                        <td style={{ padding: "10px 8px", fontWeight: 600, color: "#334155", borderRight: "1px solid #f1f5f9", whiteSpace: "nowrap" }}>
                          {row["Fabric"] || "—"}
                        </td>

                        {/* 6. Brand */}
                        <td style={{ padding: "10px 8px", fontWeight: 700, color: "#0f172a", borderRight: "1px solid #f1f5f9", whiteSpace: "nowrap" }}>
                          {row["Brand"] || "—"}
                        </td>

                        {/* 7. PCS */}
                        <td style={{ padding: "10px 8px", fontWeight: 800, color: "#0f172a", borderRight: "1px solid #f1f5f9", whiteSpace: "nowrap" }}>
                          {(Number(row.PCS) || 0).toLocaleString()}
                        </td>

                        {/* 8. Section */}
                        <td style={{ padding: "10px 8px", fontWeight: 600, color: "#475569", borderRight: "1px solid #f1f5f9", whiteSpace: "nowrap" }}>
                          {row["Section"] || "—"}
                        </td>

                        {/* 9. Season */}
                        <td style={{ padding: "10px 8px", fontWeight: 600, color: "#475569", borderRight: "1px solid #f1f5f9", whiteSpace: "nowrap" }}>
                          {row["Season"] || "—"}
                        </td>

                        {/* 10. Party Name */}
                        <td style={{ padding: "10px 8px", fontWeight: 700, color: "#0f172a", borderRight: "1px solid #f1f5f9", whiteSpace: "nowrap" }}>
                          {row["Party Name"] || "—"}
                        </td>

                        {/* 11. Direct Stitching */}
                        <td style={{ padding: "10px 8px", fontWeight: 700, borderRight: "1px solid #f1f5f9", whiteSpace: "nowrap" }}>
                          <span style={{
                            padding: "3px 8px",
                            borderRadius: "6px",
                            fontSize: "0.75rem",
                            background: String(row["Direct Stitching"]).toLowerCase() === "yes" ? "#dbeafe" : "#f1f5f9",
                            color: String(row["Direct Stitching"]).toLowerCase() === "yes" ? "#1d4ed8" : "#64748b",
                            border: String(row["Direct Stitching"]).toLowerCase() === "yes" ? "1px solid #bfdbfe" : "1px solid #e2e8f0"
                          }}>
                            {row["Direct Stitching"] || "No"}
                          </span>
                        </td>

                        {/* 12. Supervisor */}
                        <td style={{ padding: "10px 8px", fontWeight: 700, color: "#4338ca", borderRight: "1px solid #f1f5f9", whiteSpace: "nowrap" }}>
                          {row.Supervisor || "—"}
                        </td>

                        {/* 13. Date of Issue */}
                        <td style={{ padding: "10px 8px", fontWeight: 600, color: "#0f172a", whiteSpace: "nowrap" }}>
                          {row["Date of Issue"] || "—"}
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>

          {/* Pagination Footer */}
          <div style={{
            padding: "16px 24px",
            background: "#f8fafc",
            borderTop: "1px solid #e2e8f0",
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            flexWrap: "wrap",
            gap: "12px"
          }}>
            <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
              <span style={{ fontSize: "0.82rem", color: "#64748b", fontWeight: 600 }}>Rows per page:</span>
              <select
                value={pageSize}
                onChange={(e) => { setPageSize(Number(e.target.value)); setPage(1); }}
                style={{ padding: "5px 10px", borderRadius: "6px", border: "1px solid #cbd5e1", fontSize: "0.82rem", fontWeight: 700 }}
              >
                <option value={15}>15</option>
                <option value={25}>25</option>
                <option value={50}>50</option>
                <option value={100}>100</option>
              </select>
            </div>

            <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
              <button
                onClick={() => setPage(1)}
                disabled={page === 1}
                style={{ padding: "6px 12px", borderRadius: "6px", border: "1px solid #cbd5e1", background: page === 1 ? "#f1f5f9" : "#ffffff", cursor: page === 1 ? "not-allowed" : "pointer", fontWeight: 700, fontSize: "0.8rem" }}
              >
                « First
              </button>
              <button
                onClick={() => setPage(p => Math.max(1, p - 1))}
                disabled={page === 1}
                style={{ padding: "6px 12px", borderRadius: "6px", border: "1px solid #cbd5e1", background: page === 1 ? "#f1f5f9" : "#ffffff", cursor: page === 1 ? "not-allowed" : "pointer", fontWeight: 700, fontSize: "0.8rem" }}
              >
                ‹ Prev
              </button>

              <span style={{ padding: "6px 14px", fontSize: "0.84rem", fontWeight: 800, color: "#1e1b4b" }}>
                Page {page} of {totalPages}
              </span>

              <button
                onClick={() => setPage(p => Math.min(totalPages, p + 1))}
                disabled={page === totalPages}
                style={{ padding: "6px 12px", borderRadius: "6px", border: "1px solid #cbd5e1", background: page === totalPages ? "#f1f5f9" : "#ffffff", cursor: page === totalPages ? "not-allowed" : "pointer", fontWeight: 700, fontSize: "0.8rem" }}
              >
                Next ›
              </button>
              <button
                onClick={() => setPage(totalPages)}
                disabled={page === totalPages}
                style={{ padding: "6px 12px", borderRadius: "6px", border: "1px solid #cbd5e1", background: page === totalPages ? "#f1f5f9" : "#ffffff", cursor: page === totalPages ? "not-allowed" : "pointer", fontWeight: 700, fontSize: "0.8rem" }}
              >
                Last »
              </button>
            </div>
          </div>
        </div>
      </main>
    </div>
  );
}