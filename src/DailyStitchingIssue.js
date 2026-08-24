// src/components/DailyStitchingIssue.jsx
import React, { useEffect, useMemo, useState, useRef } from "react";

/**
 * Enhanced Daily Stitching Issue Tracker
 * - Smooth user experience with optimized rendering
 * - Unique classnames to prevent CSS conflicts
 * - Progressive loading with better feedback
 * - Date Range Filter for export functionality
 * - Includes Brand field from Index sheet
 */

// ====== CONFIG ======
const API_KEY = "AIzaSyAomDFBkOySlIxKWSKGHe6ATv9gvaBr7uk";
const BUDGET_SHEET_ID = "1Hj3JeJEKB43aYYWv8gk2UhdU6BWuEQfCg5pBlTdBMNA";
const INDEX_SHEET_NAME = "Index";
const CUTTING_SHEET_NAME = "Cutting";
const OLD_SHEET_ID = "18FzakygM7DVD29IRbpe68pDeCFQhFLj7t4C-XQ1MWWc";
const OLD_SHEET_NAME = "Stitching_Issues";

// Ranges
const INDEX_RANGE = `${INDEX_SHEET_NAME}!A:O`;
const CUTTING_BIG_RANGE = `${CUTTING_SHEET_NAME}!A1:ZZ200000`;
const OLD_SHEET_RANGE = `${OLD_SHEET_NAME}!A:Q`;

// Display columns
const DISPLAY_HEADERS = [
  "Sr. No",
  "Lot Number",
  "Fabric",
  "Garment Type",
  "Brand",
  "Supervisor",
  "Date of Issue",
  "PCS",
];

const COLUMN_ICONS = {
  "Sr. No": "#️⃣",
  "Lot Number": "🏷️",
  "Fabric": "🧵",
  "Garment Type": "👕",
  "Brand": "🏢",
  "Supervisor": "👨‍💼",
  "Date of Issue": "📅",
  "PCS": "🔢",
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
  if (isNaN(parsed)) return String(d);
  return parsed.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
}

// Parse date string to Date object
const parseDateString = (dateStr) => {
  if (!dateStr) return null;

  // Try different date formats
  const formats = [
    // ISO format
    () => new Date(dateStr),
    // DD/MM/YYYY
    () => {
      const parts = dateStr.split('/');
      if (parts.length === 3) {
        return new Date(parts[2], parts[1] - 1, parts[0]);
      }
      return null;
    },
    // MM/DD/YYYY
    () => {
      const parts = dateStr.split('/');
      if (parts.length === 3) {
        return new Date(parts[2], parts[0] - 1, parts[1]);
      }
      return null;
    },
    // DD-MM-YYYY
    () => {
      const parts = dateStr.split('-');
      if (parts.length === 3 && parts[2].length === 4) {
        return new Date(parts[2], parts[1] - 1, parts[0]);
      }
      return null;
    },
  ];

  for (const format of formats) {
    try {
      const date = format();
      if (date && !isNaN(date.getTime())) {
        return date;
      }
    } catch (e) {
      // Continue to next format
    }
  }

  return null;
};

const clean = (v) => (v == null ? "" : String(v).trim());

// Supervisor normalization
const SUPERVISOR_ALIASES = {};
const titleCase = (s) =>
  s
    .split(/\s+/)
    .filter(Boolean)
    .map((w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
    .join(" ");

const normalizeSupervisor = (sRaw) => {
  const s = clean(sRaw).toLowerCase();
  if (!s) return { display: "", key: "" };
  const key = s;
  const aliasDisplay = SUPERVISOR_ALIASES[s];
  const display = aliasDisplay ? aliasDisplay : titleCase(s);
  return { display, key };
};

/* ---------- Data Fetching ---------- */
const dataCache = {
  timestamp: null,
  data: null
};

async function fetchSheet({ sheetId, range, apiKey, signal }, { retries = 3, baseDelayMs = 400 } = {}) {
  const cacheKey = `${sheetId}-${range}`;

  // Check cache first (5 minute cache)
  if (dataCache[cacheKey] && Date.now() - dataCache.timestamp < 300000) {
    return dataCache[cacheKey];
  }

  const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(
    range
  )}?key=${apiKey}`;

  let attempt = 0;
  while (attempt <= retries) {
    try {
      const res = await fetch(url, { signal });
      if (res.ok) {
        const data = await res.json();
        dataCache[cacheKey] = data;
        dataCache.timestamp = Date.now();
        return data;
      }

      const text = await res.text();
      if ((res.status === 429 || res.status >= 500) && attempt < retries) {
        const delay = baseDelayMs * Math.pow(2, attempt);
        await new Promise((r) => setTimeout(r, delay));
        attempt++;
        continue;
      }
      throw new Error(`Sheets API error: ${res.status} ${text}`);
    } catch (error) {
      if (error.name === 'AbortError') throw error;
      if (attempt < retries) {
        const delay = baseDelayMs * Math.pow(2, attempt);
        await new Promise((r) => setTimeout(r, delay));
        attempt++;
        continue;
      }
      throw error;
    }
  }
}

/* ---------- Index Sheet Parser ---------- */
function parseIndexRow(header, row) {
  const hmap = {};
  header.forEach((h, i) => (hmap[normalizeKey(h)] = i));

  const get = (key) => {
    const i = hmap[key];
    return i == null || i < 0 ? "" : row[i] ?? "";
  };

  const lot = String(
    get("lotnumber") ||
    get("lot number") ||
    get("lotno") ||
    get("lot")
  ).trim();

  const supervisor = get("supervisor");
  const dateOfIssue = get("dateofissue");

  if (!lot || !supervisor || !dateOfIssue) return null;

  const startRow = parseInt(get("startrow") || "0", 10);
  const numRows = parseInt(get("numrows") || "0", 10);
  const fabric = get("fabric");
  const garmentType = get("garmenttype") || get("garment");
  const style = get("style");
  const brand = get("brand");

  const sizes = String(get("sizes") || "")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);

  return {
    lot,
    startRow,
    numRows,
    fabric,
    garmentType,
    style,
    sizes,
    supervisor,
    dateOfIssue,
    brand,
    sourceType: "Index"
  };
}

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

    let rowTotal = 0;

    sizeColIndices.forEach((c) => {
      const raw = row[c];
      if (raw != null && raw !== "") {
        const n = parseFloat(String(raw).replace(/,/g, ""));
        if (!isNaN(n) && n > 0) {
          rowTotal += n;
        }
      }
    });

    totalQty += rowTotal;
  }

  return totalQty;
}

/* ---------- Old Stitching Issues Parser - AGGREGATED ---------- */
function parseOldStitchingRows(header, allRows) {
  const hmap = {};
  header.forEach((h, i) => (hmap[normalizeKey(h)] = i));

  const get = (row, key) => {
    const i = hmap[key];
    return i == null || i < 0 ? "" : row[i] ?? "";
  };

  const aggregationMap = new Map();

  for (let i = 1; i < allRows.length; i++) {
    const row = allRows[i];

    const dateOfIssue = get(row, "dateofissue") || get(row, "timestamp");
    const lotNumber = get(row, "lotnumber") || get(row, "lot no");
    const supervisor = get(row, "supervisor");
    const fabric = get(row, "fabric");
    const garmentType = get(row, "garmenttype") || get(row, "garment");
    const pcs = get(row, "pcs");
    const brand = get(row, "brand");

    if (!lotNumber || !supervisor || !dateOfIssue) continue;

    const { display: supDisplay, key: supKey } = normalizeSupervisor(supervisor);

    const aggKey = `${clean(lotNumber)}|${supKey}|${clean(dateOfIssue)}`;

    if (!aggregationMap.has(aggKey)) {
      aggregationMap.set(aggKey, {
        "Date of Issue": formatDisplayDate(dateOfIssue),
        _rawDateOfIssue: dateOfIssue,
        _parsedDate: parseDateString(dateOfIssue),
        "Lot Number": clean(lotNumber),
        "Supervisor": supDisplay,
        _supKey: supKey,
        "Fabric": clean(fabric),
        "Garment Type": clean(garmentType),
        "Brand": clean(brand),
        "PCS": Number(pcs) || 0,
        "Source Type": "Old Lot"
      });
    } else {
      const existing = aggregationMap.get(aggKey);
      existing.PCS += Number(pcs) || 0;
    }
  }

  return Array.from(aggregationMap.values()).map(row => ({
    ...row,
    "PCS": String(row.PCS)
  }));
}

/* ---------- Loading Indicator Component ---------- */
const LoadingIndicator = ({ message = "Loading...", progress = null }) => {
  return (
    <div className="daily-stitching-loader" role="status" aria-live="polite">
      <div className="daily-stitching-spinner" />
      <div className="daily-stitching-loader-text">{message}</div>
      {progress !== null && (
        <div className="daily-stitching-loader-progress">
          <div className="daily-stitching-progress-bar">
            <div
              className="daily-stitching-progress-fill"
              style={{ width: `${progress}%` }}
            />
          </div>
          <div className="daily-stitching-progress-text">{progress}%</div>
        </div>
      )}
    </div>
  );
};

/* ---------- Main Component ---------- */
export default function DailyStitchingIssue() {
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
  const [pageSize, setPageSize] = useState(20);
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
      // 1. Fetch Index data
      setLoadingMessage("Fetching index data...");
      setLoadingProgress(10);
      const idxRes = await fetchSheet(
        {
          sheetId: BUDGET_SHEET_ID,
          range: INDEX_RANGE,
          apiKey: API_KEY,
          signal: ctrl.signal
        }
      );
      setLoadingProgress(30);
      const idxValues = idxRes.values || [];
      const idxHeader = idxValues[0] || [];
      const indexMap = new Map();

      for (let i = 1; i < idxValues.length; i++) {
        const entry = parseIndexRow(idxHeader, idxValues[i]);
        if (entry) {
          indexMap.set(entry.lot, entry);
        }
      }

      // 2. Fetch Cutting data for PCS calculation
      setLoadingMessage("Fetching cutting data...");
      setLoadingProgress(50);
      const cuttingRes = await fetchSheet(
        {
          sheetId: BUDGET_SHEET_ID,
          range: CUTTING_BIG_RANGE,
          apiKey: API_KEY,
          signal: ctrl.signal
        }
      );
      setLoadingProgress(70);
      const bigCuttingValues = cuttingRes.values || [];

      // 3. Process Index rows with PCS from cutting matrix
      setLoadingMessage("Processing index rows...");
      const indexRows = [];
      for (const [lot, indexData] of indexMap) {
        const window = sliceCuttingMatrix(bigCuttingValues, indexData.startRow, indexData.numRows);
        const totalPCS = calculateTotalPCS(window, indexData.sizes);
        const { display: supDisplay, key: supKey } = normalizeSupervisor(indexData.supervisor);
        const displayDate = indexData.dateOfIssue;
        const parsedDate = parseDateString(displayDate);

        indexRows.push({
          "Date of Issue": formatDisplayDate(displayDate),
          _rawDateOfIssue: displayDate,
          _parsedDate: parsedDate,
          "Lot Number": clean(lot),
          "Supervisor": supDisplay,
          _supKey: supKey,
          "Fabric": clean(indexData.fabric),
          "Garment Type": clean(indexData.garmentType),
          "Brand": clean(indexData.brand),
          "PCS": totalPCS > 0 ? String(totalPCS) : "",
          "Source Type": "Index",
          _isRecent: Date.now() - (parsedDate ? parsedDate.getTime() : Date.now()) < RECENT_THRESHOLD_MS
        });
      }

      // 4. Fetch and aggregate Old Stitching Issues data
      setLoadingMessage("Fetching old stitching issues...");
      setLoadingProgress(85);
      const oldRes = await fetchSheet(
        {
          sheetId: OLD_SHEET_ID,
          range: OLD_SHEET_RANGE,
          apiKey: API_KEY,
          signal: ctrl.signal
        }
      );
      setLoadingProgress(95);
      const oldValues = oldRes.values || [];
      const oldHeader = oldValues[0] || [];
      const oldRows = parseOldStitchingRows(oldHeader, oldValues).map(row => ({
        ...row,
        _isRecent: Date.now() - (row._parsedDate ? row._parsedDate.getTime() : Date.now()) < RECENT_THRESHOLD_MS
      }));

      // 5. Combine both data sources
      setLoadingMessage("Finalizing data...");
      const allRows = [...indexRows, ...oldRows];
      setRows(allRows);
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
        }, 500);
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

  /* ---------- Enhanced Filtering with Date Range ---------- */
  const filteredRows = useMemo(() => {
    let filtered = rows;

    // Apply filters
    if (filterLot) filtered = filtered.filter((r) => r["Lot Number"] === filterLot);
    if (filterSupervisor) filtered = filtered.filter((r) => r._supKey === filterSupervisor);
    if (filterIssueDate) filtered = filtered.filter((r) => r["Date of Issue"] === filterIssueDate);
    if (filterBrand) filtered = filtered.filter((r) => r["Brand"] === filterBrand);
    if (showOnlyRecent) filtered = filtered.filter((r) => r._isRecent);

    // Apply date range filter
    if (dateRangeFilter.enabled && (dateRangeFilter.startDate || dateRangeFilter.endDate)) {
      filtered = filtered.filter((row) => {
        const rowDate = row._parsedDate;
        if (!rowDate) return false;

        const rowTime = rowDate.getTime();
        const startTime = dateRangeFilter.startDate ? new Date(dateRangeFilter.startDate).getTime() : null;
        const endTime = dateRangeFilter.endDate ? new Date(dateRangeFilter.endDate).getTime() + 86400000 : null; // Add 1 day to include end date

        if (startTime && rowTime < startTime) return false;
        if (endTime && rowTime >= endTime) return false;

        return true;
      });
    }

    // Apply search
    if (searchTerm) {
      const term = searchTerm.toLowerCase();
      filtered = filtered.filter((r) =>
        DISPLAY_HEADERS.some((h) => String(r[h] ?? "").toLowerCase().includes(term))
      );
    }

    // Apply sorting
    if (sortConfig.key) {
      const { key, direction } = sortConfig;
      filtered = [...filtered].sort((a, b) => {
        if (key === "Date of Issue") {
          const da = a._parsedDate;
          const db = b._parsedDate;
          const av = da ? da.getTime() : 0;
          const bv = db ? db.getTime() : 0;
          if (av < bv) return direction === "asc" ? -1 : 1;
          if (av > bv) return direction === "asc" ? 1 : -1;
          return 0;
        }
        if (key === "Supervisor") {
          const av = a._supKey || "";
          const bv = b._supKey || "";
          if (av < bv) return direction === "asc" ? -1 : 1;
          if (av > bv) return direction === "asc" ? 1 : -1;
          return 0;
        }
        if (key === "Brand") {
          const av = String(a[key] ?? "").toLowerCase();
          const bv = String(b[key] ?? "").toLowerCase();
          if (av < bv) return direction === "asc" ? -1 : 1;
          if (av > bv) return direction === "asc" ? 1 : -1;
          return 0;
        }
        const av = String(a[key] ?? "");
        const bv = String(b[key] ?? "");
        if (av < bv) return direction === "asc" ? -1 : 1;
        if (av > bv) return direction === "asc" ? 1 : -1;
        return 0;
      });
    }

    return filtered;
  }, [rows, filterLot, filterSupervisor, filterIssueDate, filterBrand, searchTerm, sortConfig, showOnlyRecent, dateRangeFilter]);

  /* ---------- Analytics ---------- */
  const analytics = useMemo(() => {
    const recentLots = rows.filter(r => r._isRecent).length;
    const totalPCS = rows.reduce((sum, row) => sum + (parseInt(row.PCS) || 0), 0);
    const uniqueLots = new Set(rows.map(r => r["Lot Number"]).filter(Boolean)).size;
    const uniqueSupervisors = new Set(rows.map(r => r._supKey).filter(Boolean)).size;
    const uniqueBrands = new Set(rows.map(r => r["Brand"]).filter(Boolean)).size;

    return {
      recentLots,
      totalPCS,
      uniqueLots,
      uniqueSupervisors,
      uniqueBrands,
      totalRecords: rows.length
    };
  }, [rows]);

  /* ---------- Calculate Totals ---------- */
  const totalPCS = useMemo(() => {
    return rows.reduce((sum, row) => {
      const pcsValue = parseInt(row.PCS) || 0;
      return sum + pcsValue;
    }, 0);
  }, [rows]);

  const filteredTotalPCS = useMemo(() => {
    return filteredRows.reduce((sum, row) => {
      const pcsValue = parseInt(row.PCS) || 0;
      return sum + pcsValue;
    }, 0);
  }, [filteredRows]);

  /* ---------- Filter Options ---------- */
  const uniqueLots = useMemo(() => {
    const s = new Set(rows.map((r) => r["Lot Number"]).filter(Boolean));
    return Array.from(s).sort((a, b) =>
      a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" })
    );
  }, [rows]);

  const uniqueSupervisors = useMemo(() => {
    const map = new Map();
    for (const r of rows) {
      if (r._supKey) {
        if (!map.has(r._supKey)) map.set(r._supKey, r.Supervisor || "");
      }
    }
    return Array.from(map.entries())
      .map(([key, label]) => ({ key, label }))
      .sort((a, b) => a.label.localeCompare(b.label));
  }, [rows]);

  const uniqueIssueDates = useMemo(() => {
    const s = new Set(rows.map((r) => r["Date of Issue"]).filter(Boolean));
    return Array.from(s).sort((a, b) => {
      const da = new Date(a);
      const db = new Date(b);
      if (isNaN(da) || isNaN(db)) return b.localeCompare(a);
      return db - da;
    });
  }, [rows]);

  const uniqueBrands = useMemo(() => {
    const s = new Set(rows.map((r) => r["Brand"]).filter(Boolean));
    return Array.from(s).sort((a, b) => a.localeCompare(b));
  }, [rows]);

  /* ---------- Date Range Functions ---------- */
  const handleDateRangeChange = (field, value) => {
    setDateRangeFilter(prev => ({
      ...prev,
      [field]: value
    }));
    setPage(1);
  };

  const toggleDateRangeFilter = () => {
    setDateRangeFilter(prev => ({
      ...prev,
      enabled: !prev.enabled,
      startDate: !prev.enabled ? "" : prev.startDate,
      endDate: !prev.enabled ? "" : prev.endDate
    }));
    setPage(1);
  };

  const clearDateRangeFilter = () => {
    setDateRangeFilter({
      startDate: "",
      endDate: "",
      enabled: false
    });
    setPage(1);
  };

  const getDateRangeLabel = () => {
    if (!dateRangeFilter.enabled) return "No date range";
    if (dateRangeFilter.startDate && dateRangeFilter.endDate) {
      return `${formatDisplayDate(dateRangeFilter.startDate)} to ${formatDisplayDate(dateRangeFilter.endDate)}`;
    } else if (dateRangeFilter.startDate) {
      return `From ${formatDisplayDate(dateRangeFilter.startDate)}`;
    } else if (dateRangeFilter.endDate) {
      return `Until ${formatDisplayDate(dateRangeFilter.endDate)}`;
    }
    return "Custom date range";
  };

  /* ---------- Event Handlers ---------- */
  const handleSort = (key) => {
    let direction = "asc";
    if (sortConfig.key === key && sortConfig.direction === "asc") {
      direction = "desc";
    } else if (sortConfig.key !== key) {
      direction = key === "Date of Issue" ? "desc" : "asc";
    }
    setSortConfig({ key, direction });
    setPage(1);
  };

  const hasActiveFilters = filterLot || filterSupervisor || filterIssueDate || filterBrand || searchTerm || showOnlyRecent ||
    (dateRangeFilter.enabled && (dateRangeFilter.startDate || dateRangeFilter.endDate));

  const clearFilters = () => {
    setFilterLot("");
    setFilterSupervisor("");
    setFilterIssueDate("");
    setFilterBrand("");
    setSearchTerm("");
    setShowOnlyRecent(false);
    clearDateRangeFilter();
    setSortConfig({ key: "Date of Issue", direction: "desc" });
    setPage(1);
  };

  const handleRefresh = () => loadData("refresh");

  const goBack = () => {
    try {
      if (window.history.length > 1) {
        window.history.back();
        return;
      }
    } catch { }
    window.location.href = "/";
  };

  /* ---------- Export Functions with Date Range ---------- */
  const downloadExcel = (data, filename = "stitching-issues") => {
    // Remove "Sr. No" from headers for export since it's dynamic
    const exportHeaders = DISPLAY_HEADERS.filter(h => h !== "Sr. No");
    const headers = exportHeaders.join(",");
    const rows = data
      .map((row, index) =>
        exportHeaders.map((header) => {
          const value = row[header] || "";
          const escaped = String(value).replace(/"/g, '""');
          return escaped.includes(",") ? `"${escaped}"` : escaped;
        }).join(",")
      )
      .join("\n");

    let dateRangeInfo = "";
    if (dateRangeFilter.enabled && (dateRangeFilter.startDate || dateRangeFilter.endDate)) {
      dateRangeInfo = `Date Range: ${getDateRangeLabel()}\n`;
    }

    const csvContent = `${dateRangeInfo}${headers}\n${rows}`;
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    link.setAttribute("href", url);
    link.setAttribute(
      "download",
      `${filename}-${new Date().toISOString().split("T")[0]}.csv`
    );
    link.style.visibility = "hidden";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const downloadPDF = (data, filename = "stitching-issues") => {
    const printWindow = window.open("", "_blank");
    const currentDate = new Date().toLocaleDateString();
    const totalPCS = data.reduce((sum, row) => sum + (parseInt(row.PCS) || 0), 0);

    // Date range info for PDF
    let dateRangeHtml = "";
    if (dateRangeFilter.enabled && (dateRangeFilter.startDate || dateRangeFilter.endDate)) {
      dateRangeHtml = `<p style="font-size: 12px; color: #6b7280; margin: 2px 0;">Date Range: ${getDateRangeLabel()}</p>`;
    }

    // Remove "Sr. No" from headers for PDF export
    const pdfHeaders = DISPLAY_HEADERS.filter(h => h !== "Sr. No");

    const tableContent = `
    <html>
      <head>
        <title>${filename}</title>
        <style>
          body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Inter, Arial, sans-serif; margin: 20px; color:#111; }
          .daily-stitching-header { text-align: left; margin-bottom: 16px; border-bottom: 1px solid #e5e7eb; padding-bottom: 8px; }
          .daily-stitching-title { font-size: 20px; font-weight: 700; margin: 0; }
          .daily-stitching-subtitle { font-size: 12px; color: #6b7280; margin: 4px 0; }
          .daily-stitching-date { font-size: 12px; color: #6b7280; }
          .daily-stitching-total-pcs { font-size: 14px; color: #059669; font-weight: 600; margin: 8px 0; }
          .daily-stitching-table { width: 100%; border-collapse: collapse; margin-top: 14px; font-size: 12px; }
          .daily-stitching-th, .daily-stitching-td { border: 1px solid #e5e7eb; padding: 8px; text-align: left; }
          .daily-stitching-th { background:#f8fafc; font-weight: 600; }
          .daily-stitching-tr:nth-child(even){ background:#fcfcfd; }
          .daily-stitching-tr-recent { background: #f0f9ff !important; border-left: 3px solid #3b82f6; }
          @media print { body { margin: 0; } .daily-stitching-no-print { display: none; } }
        </style>
      </head>
      <body>
        <div class="daily-stitching-header">
          <h1 class="daily-stitching-title">Daily Stitching Issue Report</h1>
          <p class="daily-stitching-subtitle">Quality Control Tracking</p>
          <p class="daily-stitching-date">Generated on: ${currentDate} • Total Records: ${data.length}</p>
          ${dateRangeHtml}
          <p class="daily-stitching-total-pcs">Total PCS: ${totalPCS.toLocaleString()}</p>
        </div>
        <table class="daily-stitching-table">
          <thead>
            <tr>
              ${pdfHeaders.map((header) => `<th class="daily-stitching-th">${header}</th>`).join("")}
            </tr>
          </thead>
          <tbody>
            ${data
        .map(
          (row) => `
              <tr class="daily-stitching-tr ${row._isRecent ? 'daily-stitching-tr-recent' : ''}">
                ${pdfHeaders.map((header) => `<td class="daily-stitching-td">${row[header] || "-"}</td>`).join("")}
              </tr>
            `
        )
        .join("")}
          </tbody>
        </table>
        <div class="daily-stitching-no-print" style="margin-top: 16px;">
          <button onclick="window.print()" style="padding: 8px 12px; border:1px solid #e5e7eb; background:#fff; border-radius:6px; cursor:pointer;">Print / Save as PDF</button>
          <button onclick="window.close()" style="padding: 8px 12px; margin-left:8px; border:1px solid #e5e7eb; background:#fff; border-radius:6px; cursor:pointer;">Close</button>
        </div>
      </body>
    </html>`;
    printWindow.document.write(tableContent);
    printWindow.document.close();
  };

  /* ---------- Pagination ---------- */
  const totalPages = Math.max(1, Math.ceil(filteredRows.length / pageSize));
  const pagedRows = useMemo(() => {
    const start = (page - 1) * pageSize;
    return filteredRows.slice(start, start + pageSize);
  }, [filteredRows, page, pageSize]);

  return (
    <div className="daily-stitching-container">
      <style>{`
        .daily-stitching-container {
          min-height: 100vh;
          background-color: #f8fafc;
          background-image: 
            radial-gradient(at 0% 0%, rgba(99, 102, 241, 0.08) 0px, transparent 50%), 
            radial-gradient(at 100% 0%, rgba(236, 72, 153, 0.06) 0px, transparent 50%), 
            radial-gradient(at 50% 100%, rgba(16, 185, 129, 0.06) 0px, transparent 50%);
          font-family: 'Plus Jakarta Sans', 'Inter', system-ui, -apple-system, sans-serif;
          color: #0f172a;
          padding: 24px 32px;
        }

        .daily-stitching-main {
          max-width: 100%;
          margin: 0 auto;
        }

        .daily-stitching-header {
          background: linear-gradient(135deg, #1e1b4b 0%, #312e81 40%, #4338ca 100%);
          border-radius: 24px;
          padding: 32px 36px;
          margin-bottom: 28px;
          border: 1px solid rgba(255, 255, 255, 0.1);
          box-shadow: 0 20px 40px -15px rgba(30, 27, 75, 0.25);
          position: relative;
          overflow: hidden;
        }

        .daily-stitching-header-content {
          display: flex;
          justify-content: space-between;
          align-items: center;
          gap: 24px;
          position: relative;
          z-index: 1;
        }

        .daily-stitching-header-left {
          flex: 1;
        }

        .daily-stitching-title {
          margin: 0 0 6px 0;
          font-size: 2rem;
          font-weight: 800;
          color: #ffffff;
          letter-spacing: -0.02em;
          display: flex;
          align-items: center;
          gap: 12px;
        }

        .daily-stitching-subtitle {
          margin: 0;
          font-size: 0.95rem;
          color: #c7d2fe;
          font-weight: 500;
        }

        .daily-stitching-stats {
          display: flex;
          gap: 16px;
        }

        .daily-stitching-stat {
          text-align: center;
          padding: 14px 22px;
          background: rgba(255, 255, 255, 0.12);
          backdrop-filter: blur(12px);
          -webkit-backdrop-filter: blur(12px);
          border-radius: 18px;
          border: 1px solid rgba(255, 255, 255, 0.2);
          min-width: 120px;
          box-shadow: 0 8px 16px rgba(0, 0, 0, 0.1);
          transition: all 0.2s ease;
        }

        .daily-stitching-stat:hover {
          transform: translateY(-2px);
          background: rgba(255, 255, 255, 0.18);
        }

        .daily-stitching-stat-value {
          display: block;
          font-size: 1.8rem;
          font-weight: 800;
          margin-bottom: 2px;
          color: white;
          line-height: 1.1;
        }

        .daily-stitching-stat-label {
          font-size: 0.72rem;
          color: #e0e7ff;
          text-transform: uppercase;
          letter-spacing: 0.06em;
          font-weight: 700;
        }

        .daily-stitching-toolbar {
          background: #ffffff;
          border-radius: 20px;
          padding: 24px 28px;
          margin-bottom: 28px;
          border: 1px solid #e2e8f0;
          box-shadow: 0 10px 30px rgba(0, 0, 0, 0.03);
        }

        .daily-stitching-search {
          position: relative;
          margin-bottom: 20px;
        }

        .daily-stitching-search-icon {
          position: absolute;
          left: 14px;
          top: 50%;
          transform: translateY(-50%);
          color: #64748b;
          font-size: 16px;
        }

        .daily-stitching-search-input {
          padding: 12px 14px 12px 46px;
          border: 1.5px solid #cbd5e1;
          border-radius: 12px;
          background: #ffffff;
          color: #0f172a;
          font-size: 0.9rem;
          font-weight: 600;
          width: 100%;
          box-sizing: border-box;
          transition: all 0.2s ease;
        }

        .daily-stitching-search-input:focus {
          outline: none;
          border-color: #6366f1;
          box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.15);
        }

        .daily-stitching-filters {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(240px, 1fr));
          gap: 18px;
          margin-bottom: 20px;
        }

        .daily-stitching-filter-group {
          display: flex;
          flex-direction: column;
          gap: 6px;
        }

        .daily-stitching-filter-label {
          font-size: 0.82rem;
          color: #475569;
          font-weight: 700;
          text-transform: uppercase;
          letter-spacing: 0.03em;
        }

        .daily-stitching-filter-select, .daily-stitching-date-input {
          padding: 10px 14px;
          border: 1.5px solid #cbd5e1;
          border-radius: 12px;
          background: #ffffff;
          color: #0f172a;
          font-size: 0.88rem;
          font-weight: 600;
          transition: all 0.2s ease;
          width: 100%;
          box-sizing: border-box;
        }

        .daily-stitching-filter-select:focus, .daily-stitching-date-input:focus {
          outline: none;
          border-color: #6366f1;
          box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.15);
        }

        .daily-stitching-date-range-group {
          grid-column: span 2;
          background: #f8fafc;
          border-radius: 16px;
          padding: 18px;
          border: 1.5px solid #e2e8f0;
        }

        .daily-stitching-controls {
          display: flex;
          justify-content: space-between;
          align-items: center;
          gap: 16px;
          flex-wrap: wrap;
        }

        .daily-stitching-buttons {
          display: flex;
          gap: 12px;
          flex-wrap: wrap;
        }

        .daily-stitching-btn {
          display: inline-flex;
          align-items: center;
          gap: 8px;
          padding: 10px 20px;
          border: none;
          border-radius: 12px;
          font-size: 0.88rem;
          font-weight: 700;
          cursor: pointer;
          transition: all 0.2s ease;
        }

        .daily-stitching-btn:hover:not(:disabled) {
          transform: translateY(-2px);
        }

        .daily-stitching-btn-excel {
          background: linear-gradient(135deg, #10b981 0%, #059669 100%);
          color: #ffffff;
          box-shadow: 0 4px 12px rgba(16, 185, 129, 0.25);
        }

        .daily-stitching-btn-pdf {
          background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%);
          color: #ffffff;
          box-shadow: 0 4px 12px rgba(239, 68, 68, 0.25);
        }

        .daily-stitching-btn-refresh {
          background: linear-gradient(135deg, #6366f1 0%, #4338ca 100%);
          color: #ffffff;
          box-shadow: 0 4px 12px rgba(99, 102, 241, 0.25);
        }

        .daily-stitching-btn-back {
          background: #0f172a;
          color: #ffffff;
          box-shadow: 0 4px 12px rgba(15, 23, 42, 0.2);
        }

        .daily-stitching-btn-clear {
          background: #f1f5f9;
          color: #475569;
          border: 1px solid #cbd5e1;
        }

        .daily-stitching-table-container {
          background: #ffffff;
          border-radius: 24px;
          border: 1px solid #e2e8f0;
          box-shadow: 0 10px 30px rgba(0, 0, 0, 0.03);
          overflow: hidden;
          margin-bottom: 24px;
        }

        .daily-stitching-table-wrapper {
          overflow-x: auto;
          max-height: 70vh;
        }

        .daily-stitching-table {
          width: 100%;
          border-collapse: separate;
          border-spacing: 0;
          font-size: 0.88rem;
        }

        .daily-stitching-thead {
          position: sticky;
          top: 0;
          z-index: 20;
          background: linear-gradient(135deg, #1e1b4b 0%, #312e81 100%);
        }

        .daily-stitching-th {
          padding: 16px 14px;
          text-align: center;
          font-weight: 800;
          color: #ffffff;
          font-size: 0.75rem;
          text-transform: uppercase;
          letter-spacing: 0.06em;
          border-bottom: 2px solid #312e81;
          white-space: nowrap;
        }

        .daily-stitching-td {
          padding: 12px 14px;
          border-bottom: 1px solid #f1f5f9;
          border-right: 1px solid #f8fafc;
          background: #ffffff;
          color: #1e293b;
          text-align: center;
          vertical-align: middle;
        }

        .daily-stitching-tr:nth-child(even) .daily-stitching-td {
          background: #f8fafc;
        }

        .daily-stitching-tr:hover .daily-stitching-td {
          background: #e0e7ff;
        }

        .daily-stitching-badge-new {
          display: inline-flex;
          align-items: center;
          justify-content: center;
          padding: 3px 8px;
          background: linear-gradient(135deg, #ec4899 0%, #be185d 100%);
          color: #ffffff;
          border-radius: 20px;
          font-size: 0.68rem;
          font-weight: 800;
          letter-spacing: 0.06em;
          box-shadow: 0 2px 8px rgba(236, 72, 153, 0.4);
        }

        .daily-stitching-pagination {
          display: flex;
          justify-content: space-between;
          align-items: center;
          background: #ffffff;
          padding: 20px 24px;
          border-radius: 20px;
          border: 1px solid #e2e8f0;
          box-shadow: 0 10px 30px rgba(0, 0, 0, 0.03);
        }

        .daily-stitching-loader {
          position: fixed;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          background: rgba(15, 23, 42, 0.5);
          backdrop-filter: blur(6px);
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          z-index: 10000;
        }

        .daily-stitching-spinner {
          width: 56px;
          height: 56px;
          border-radius: 50%;
          background: conic-gradient(from 0deg, #6366f1, #ec4899, #10b981, #6366f1);
          animation: daily-stitching-spin 1.2s linear infinite;
          padding: 4px;
          mask: radial-gradient(farthest-side, transparent calc(100% - 4px), #fff 0);
          -webkit-mask: radial-gradient(farthest-side, transparent calc(100% - 4px), #fff 0);
          margin-bottom: 20px;
        }

        @keyframes daily-stitching-spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }

        .daily-stitching-loader-text {
          font-size: 1.1rem;
          color: #ffffff;
          margin-bottom: 16px;
          font-weight: 700;
          text-align: center;
        }

        .daily-stitching-loader-progress {
          width: 300px;
          max-width: 80%;
          margin-top: 10px;
        }

        .daily-stitching-progress-bar {
          width: 100%;
          height: 8px;
          background: rgba(67, 49, 168, 0.1);
          border-radius: 4px;
          overflow: hidden;
          margin-bottom: 8px;
        }

        .daily-stitching-progress-fill {
          height: 100%;
          background: linear-gradient(90deg, #4331a8, #6d5bd9);
          border-radius: 4px;
          transition: width 0.3s ease;
        }

        .daily-stitching-progress-text {
          font-size: 14px;
          color: #4331a8;
          font-weight: 500;
          text-align: center;
        }

        .daily-stitching-refresh-overlay {
          position: fixed;
          top: 20px;
          right: 20px;
          background: rgba(67, 49, 168, 0.9);
          color: white;
          padding: 12px 20px;
          border-radius: 8px;
          display: flex;
          align-items: center;
          gap: 10px;
          z-index: 1001;
          box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
          animation: daily-stitching-slide-in 0.3s ease;
        }

        .daily-stitching-refresh-spinner {
          width: 20px;
          height: 20px;
          border: 2px solid rgba(255, 255, 255, 0.3);
          border-top: 2px solid white;
          border-radius: 50%;
          animation: daily-stitching-spin 1s linear infinite;
        }

        @keyframes daily-stitching-spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }

        @keyframes daily-stitching-slide-in {
          from {
            transform: translateX(100%);
            opacity: 0;
          }
          to {
            transform: translateX(0);
            opacity: 1;
          }
        }

        .daily-stitching-pcs-info {
          background: #f0fdf4;
          border: 1px solid #bbf7d0;
          border-radius: 12px;
          padding: 16px 20px;
          margin-bottom: 16px;
          display: flex;
          justify-content: space-between;
          align-items: center;
        }

        .daily-stitching-pcs-info span {
          font-weight: 700;
          color: #059669;
        }

        .daily-stitching-empty {
          text-align: center;
          padding: 40px;
          color: #64748b;
        }

        .daily-stitching-empty-icon {
          font-size: 48px;
          margin-bottom: 16px;
        }

        .daily-stitching-empty-text {
          font-size: 16px;
          font-weight: 500;
        }

        @media (max-width: 768px) {
          .daily-stitching-main {
            padding: 16px;
          }
          
          .daily-stitching-header-content {
            flex-direction: column;
          }
          
          .daily-stitching-stats {
            min-width: auto;
            width: 100%;
          }
          
          .daily-stitching-controls {
            flex-direction: column;
            align-items: stretch;
          }
          
          .daily-stitching-buttons {
            justify-content: center;
          }
          
          .daily-stitching-date-range-group {
            grid-column: span 1;
          }
          
          .daily-stitching-date-range-fields {
            grid-template-columns: 1fr;
          }
          
          .daily-stitching-loader-progress {
            width: 250px;
          }
        }
      `}</style>

      {loading && (
        <LoadingIndicator
          message={loadingMessage}
          progress={loadingProgress}
        />
      )}

      {refreshing && (
        <div className="daily-stitching-refresh-overlay">
          <div className="daily-stitching-refresh-spinner" />
          <span>Refreshing data...</span>
        </div>
      )}

      <div className="daily-stitching-main" aria-busy={loading || refreshing}>
        {/* Header */}
        <div className="daily-stitching-header">
          <div className="daily-stitching-header-content">
            <div className="daily-stitching-header-left">
              <h1 className="daily-stitching-title">DAILY STITCHING ISSUE</h1>
              <p className="daily-stitching-subtitle">Daily Stitching Issue Tracker with Brand Information</p>
              {lastUpdated && (
                <p style={{ color: '#e2e8f0', fontSize: '12px', marginTop: '8px', opacity: 0.8 }}>
                  Last updated: {lastUpdated}
                </p>
              )}
            </div>
            <div className="daily-stitching-stats">
              <div className="daily-stitching-stat">
                <span className="daily-stitching-stat-value">{analytics.totalRecords}</span>
                <span className="daily-stitching-stat-label">Total Records</span>
              </div>
              <div className="daily-stitching-stat">
                <span className="daily-stitching-stat-value">{analytics.uniqueLots}</span>
                <span className="daily-stitching-stat-label">Active Lots</span>
              </div>
              <div className="daily-stitching-stat daily-stitching-stat-recent">
                <span className="daily-stitching-stat-value">{analytics.recentLots}</span>
                <span className="daily-stitching-stat-label">Recent Lots</span>
              </div>
              {/* <div className="daily-stitching-stat">
                <span className="daily-stitching-stat-value">{analytics.uniqueBrands}</span>
                <span className="daily-stitching-stat-label">Brands</span>
              </div> */}
              {/* <div className="daily-stitching-stat daily-stitching-stat-pcs">
                <span className="daily-stitching-stat-value">{analytics.totalPCS.toLocaleString()}</span>
                <span className="daily-stitching-stat-label">Total PCS</span>
              </div> */}
            </div>
          </div>
        </div>

        {/* Error Display */}
        {error && (
          <div style={{
            background: '#fef2f2',
            border: '1px solid #fecaca',
            borderRadius: '12px',
            padding: '16px',
            marginBottom: '24px',
            color: '#991b1b'
          }}>
            <strong>Error:</strong> {error}
            <button
              onClick={handleRefresh}
              style={{
                marginLeft: '16px',
                padding: '8px 16px',
                background: '#dc2626',
                color: 'white',
                border: 'none',
                borderRadius: '6px',
                cursor: 'pointer',
                fontWeight: '600'
              }}
            >
              Retry
            </button>
          </div>
        )}

        {/* Total PCS Info */}
        {hasActiveFilters && filteredTotalPCS !== totalPCS && (
          <div className="daily-stitching-pcs-info">
            <div>
              <strong>Filtered Total PCS:</strong> <span>{filteredTotalPCS.toLocaleString()}</span>
            </div>
            <div style={{ fontSize: '12px', color: '#6b7280' }}>
              (Overall Total: {totalPCS.toLocaleString()})
            </div>
          </div>
        )}

        {/* Toolbar */}
        <div className="daily-stitching-toolbar">
          <div className="daily-stitching-search">
            <span className="daily-stitching-search-icon">🔍</span>
            <input
              type="text"
              className="daily-stitching-search-input"
              placeholder="Search across all columns…"
              value={searchTerm}
              onChange={(e) => { setSearchTerm(e.target.value); setPage(1); }}
              disabled={loading}
            />
          </div>

          <div className="daily-stitching-filters">
            <div className="daily-stitching-filter-group">
              <label className="daily-stitching-filter-label">Lot Number</label>
              <select
                value={filterLot}
                onChange={(e) => { setFilterLot(e.target.value); setPage(1); }}
                className="daily-stitching-filter-select"
                disabled={loading}
              >
                <option value="">All Lots</option>
                {uniqueLots.map((lot) => (
                  <option key={lot} value={lot}>{lot}</option>
                ))}
              </select>
            </div>

            <div className="daily-stitching-filter-group">
              <label className="daily-stitching-filter-label">Supervisor</label>
              <select
                value={filterSupervisor}
                onChange={(e) => { setFilterSupervisor(e.target.value); setPage(1); }}
                className="daily-stitching-filter-select"
                disabled={loading}
              >
                <option value="">All Supervisors</option>
                {uniqueSupervisors.map(({ key, label }) => (
                  <option key={key} value={key}>{label}</option>
                ))}
              </select>
            </div>

            <div className="daily-stitching-filter-group">
              <label className="daily-stitching-filter-label">Issue Date</label>
              <select
                value={filterIssueDate}
                onChange={(e) => { setFilterIssueDate(e.target.value); setPage(1); }}
                className="daily-stitching-filter-select"
                disabled={loading}
              >
                <option value="">All Dates</option>
                {uniqueIssueDates.map((d) => (
                  <option key={d} value={d}>{d}</option>
                ))}
              </select>
            </div>

            <div className="daily-stitching-filter-group">
              <label className="daily-stitching-filter-label">Brand</label>
              <select
                value={filterBrand}
                onChange={(e) => { setFilterBrand(e.target.value); setPage(1); }}
                className="daily-stitching-filter-select"
                disabled={loading}
              >
                <option value="">All Brands</option>
                {uniqueBrands.map((brand) => (
                  <option key={brand} value={brand}>{brand || "No Brand"}</option>
                ))}
              </select>
            </div>

            {/* Date Range Filter */}
            <div className="daily-stitching-date-range-group">
              <div className="daily-stitching-date-range-header">
                <div className="daily-stitching-date-range-title">
                  <span>📅</span>
                  <span>Date Range Filter</span>
                </div>
                <div
                  className={`daily-stitching-date-range-toggle ${dateRangeFilter.enabled ? 'active' : ''}`}
                  onClick={toggleDateRangeFilter}
                  style={{ cursor: loading ? 'not-allowed' : 'pointer', opacity: loading ? 0.6 : 1 }}
                >
                  {dateRangeFilter.enabled ? '✓ Active' : 'Enable'}
                </div>
              </div>

              <div className="daily-stitching-date-range-fields">
                <div>
                  <label className="daily-stitching-filter-label">Start Date</label>
                  <input
                    type="date"
                    value={dateRangeFilter.startDate}
                    onChange={(e) => handleDateRangeChange('startDate', e.target.value)}
                    className="daily-stitching-date-input"
                    disabled={!dateRangeFilter.enabled || loading}
                  />
                </div>
                <div>
                  <label className="daily-stitching-filter-label">End Date</label>
                  <input
                    type="date"
                    value={dateRangeFilter.endDate}
                    onChange={(e) => handleDateRangeChange('endDate', e.target.value)}
                    className="daily-stitching-date-input"
                    disabled={!dateRangeFilter.enabled || loading}
                  />
                </div>
              </div>

              {dateRangeFilter.enabled && (
                <div className="daily-stitching-date-range-info">
                  <strong>Active Filter:</strong> {getDateRangeLabel()}
                  {dateRangeFilter.startDate && dateRangeFilter.endDate && (
                    <span> • Records in range: {filteredRows.length}</span>
                  )}
                </div>
              )}
            </div>

            <div className="daily-stitching-filter-group">
              <label className="daily-stitching-filter-label">Quick Filters</label>
              <div className="daily-stitching-checkbox">
                <input
                  type="checkbox"
                  id="daily-stitching-showRecent"
                  checked={showOnlyRecent}
                  onChange={(e) => { setShowOnlyRecent(e.target.checked); setPage(1); }}
                  disabled={loading}
                />
                <label htmlFor="daily-stitching-showRecent" style={{ fontSize: '14px', color: '#374151' }}>
                  Show only recent lots (last 24h)
                </label>
              </div>
            </div>
          </div>

          <div className="daily-stitching-controls">
            <div className="daily-stitching-buttons">
              <button
                onClick={() => downloadExcel(filteredRows, "stitching-quality-report")}
                className="daily-stitching-btn daily-stitching-btn-excel"
                disabled={filteredRows.length === 0 || loading}
                title="Export current filtered data to CSV"
              >
                📊 Export CSV
              </button>
              <button
                onClick={() => downloadPDF(filteredRows, "stitching-quality-report")}
                className="daily-stitching-btn daily-stitching-btn-pdf"
                disabled={filteredRows.length === 0 || loading}
                title="Generate PDF report with current filters"
              >
                📄 PDF Report
              </button>
              <button
                onClick={handleRefresh}
                className="daily-stitching-btn daily-stitching-btn-refresh"
                disabled={loading || refreshing}
              >
                🔄 {refreshing ? "Refreshing..." : "Refresh Data"}
              </button>
              <button
                onClick={goBack}
                className="daily-stitching-btn daily-stitching-btn-back"
                disabled={loading}
              >
                ← Back
              </button>
            </div>

            {hasActiveFilters && (
              <button
                onClick={clearFilters}
                className="daily-stitching-btn daily-stitching-btn-clear"
                disabled={loading}
              >
                🗑️ Clear All Filters
              </button>
            )}
          </div>
        </div>

        {/* Table */}
        <div className="daily-stitching-table-container">
          <div className="daily-stitching-table-wrapper">
            <table className="daily-stitching-table">
              <thead className="daily-stitching-thead">
                <tr>
                  {DISPLAY_HEADERS.map((h) => (
                    <th
                      key={h}
                      className="daily-stitching-th"
                      onClick={() => !loading && handleSort(h)}
                      style={{ cursor: loading ? 'default' : 'pointer' }}
                    >
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '6px' }}>
                        <span>{COLUMN_ICONS[h]}</span>
                        <span>{h}</span>
                        {sortConfig.key === h && (
                          <span>{sortConfig.direction === "asc" ? "↑" : "↓"}</span>
                        )}
                      </div>
                    </th>
                  ))}
                </tr>
              </thead>
              {/* Table Body */}
              <tbody>
                {pagedRows.length === 0 ? (
                  <tr>
                    <td colSpan={DISPLAY_HEADERS.length} className="daily-stitching-td">
                      <div className="daily-stitching-empty">
                        <div className="daily-stitching-empty-icon">📭</div>
                        <div className="daily-stitching-empty-text">
                          {rows.length === 0 ? 'No data available' : 'No records found matching your filters'}
                        </div>
                        {hasActiveFilters && (
                          <button
                            className="daily-stitching-btn daily-stitching-btn-clear"
                            onClick={clearFilters}
                            style={{ marginTop: '16px' }}
                            disabled={loading}
                          >
                            Clear all filters
                          </button>
                        )}
                      </div>
                    </td>
                  </tr>
                ) : (
                  pagedRows.map((row, idx) => {
                    const serialNumber = ((page - 1) * pageSize) + idx + 1;
                    return (
                      <tr
                        key={`${row["Lot Number"]}-${idx}-${row["Source Type"]}`}
                        className={`daily-stitching-tr ${row._isRecent ? 'daily-stitching-tr-recent' : ''}`}
                      >
                        {/* Sr. No */}
                        <td className="daily-stitching-td">
                          {serialNumber}
                        </td>
                        {/* Lot Number */}
                        <td className="daily-stitching-td">
                          <div style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', justifyContent: 'center' }}>
                            <span style={{ fontWeight: '700' }}>{row["Lot Number"]}</span>
                            {row._isRecent && (
                              <span className="daily-stitching-badge-new">NEW</span>
                            )}
                          </div>
                        </td>
                        {/* Fabric */}
                        <td className="daily-stitching-td">
                          {row["Fabric"] || <span style={{ color: '#9ca3af', fontStyle: 'italic' }}>—</span>}
                        </td>
                        {/* Garment Type */}
                        <td className="daily-stitching-td">
                          {row["Garment Type"] || <span style={{ color: '#9ca3af', fontStyle: 'italic' }}>—</span>}
                        </td>
                        {/* Brand */}
                        <td className="daily-stitching-td">
                          {row["Brand"] || <span style={{ color: '#9ca3af', fontStyle: 'italic' }}>—</span>}
                        </td>
                        {/* Supervisor */}
                        <td className="daily-stitching-td">
                          {row["Supervisor"] || <span style={{ color: '#9ca3af', fontStyle: 'italic' }}>—</span>}
                        </td>
                        {/* Date of Issue */}
                        <td className="daily-stitching-td">
                          {row["Date of Issue"] || <span style={{ color: '#9ca3af', fontStyle: 'italic' }}>—</span>}
                        </td>
                        {/* PCS */}
                        <td className="daily-stitching-td">
                          {row["PCS"] || <span style={{ color: '#9ca3af', fontStyle: 'italic' }}>—</span>}
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        </div>

        {/* Pagination */}
        {filteredRows.length > 0 && (
          <div className="daily-stitching-pagination">
            <div style={{ color: '#374151', fontSize: '14px', fontWeight: '500' }}>
              Showing {((page - 1) * pageSize) + 1} to {Math.min(page * pageSize, filteredRows.length)} of {filteredRows.length} entries
              {hasActiveFilters && (
                <span style={{ color: '#059669', marginLeft: '12px' }}>
                  • Filtered PCS: {filteredTotalPCS.toLocaleString()}
                </span>
              )}
              {dateRangeFilter.enabled && (
                <span style={{ color: '#4331a8', marginLeft: '12px' }}>
                  • Date Range: {getDateRangeLabel()}
                </span>
              )}
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <button
                className="daily-stitching-btn"
                onClick={() => setPage(1)}
                disabled={page <= 1 || loading}
                style={{ padding: '8px 12px', fontSize: '14px' }}
              >
                « First
              </button>
              <button
                className="daily-stitching-btn"
                onClick={() => setPage(p => Math.max(1, p - 1))}
                disabled={page <= 1 || loading}
                style={{ padding: '8px 12px', fontSize: '14px' }}
              >
                ‹ Prev
              </button>
              <span style={{ color: '#374151', fontSize: '14px', margin: '0 12px' }}>
                Page {page} of {totalPages}
              </span>
              <button
                className="daily-stitching-btn"
                onClick={() => setPage(p => Math.min(totalPages, p + 1))}
                disabled={page >= totalPages || loading}
                style={{ padding: '8px 12px', fontSize: '14px' }}
              >
                Next ›
              </button>
              <button
                className="daily-stitching-btn"
                onClick={() => setPage(totalPages)}
                disabled={page >= totalPages || loading}
                style={{ padding: '8px 12px', fontSize: '14px' }}
              >
                Last »
              </button>
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <span style={{ color: '#374151', fontSize: '14px' }}>Rows per page:</span>
              <select
                value={pageSize}
                onChange={(e) => { setPageSize(Number(e.target.value)); setPage(1); }}
                className="daily-stitching-filter-select"
                style={{ padding: '6px 8px', fontSize: '14px' }}
                disabled={loading}
              >
                {[20, 50, 100, 200].map((n) => (
                  <option key={n} value={n}>{n}</option>
                ))}
              </select>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}