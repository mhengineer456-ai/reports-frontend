import React, { useEffect, useMemo, useState, useRef } from "react";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import { autoTable } from "jspdf-autotable";
import { GOOGLE_API_KEY, SPREADSHEET_IDS, fetchSheetDataFromBackend } from "./config";

/** ====== CONFIG ====== */
const JOB_SHEET_ID = SPREADSHEET_IDS.JOBORDER;
const API_KEY = GOOGLE_API_KEY;
const JOB_RANGE = "JobOrder!A:AZ";

// Budget Report spreadsheet
const BUDGET_SHEET_ID = SPREADSHEET_IDS.MAIN;
const CUTTING_SHEET_NAME = "Cutting";
const INDEX_SHEET_NAME = "Index";
const INDEX_RANGE = `${INDEX_SHEET_NAME}!A:AG`;

// One big read to avoid 429s
const CUTTING_BIG_RANGE = `${CUTTING_SHEET_NAME}!A1:ZZ400000`;

// Canonical output columns (order)
const OUTPUT_COLS = [
  "Job Order No",
  "Date",
  "Fabric",
  "Brand",
  "Style",
  "Party Name",
  "Garment Type",
  "Section",
  "Season",
  "Priority",
  "Direct Stitching",
  "Lot No",
  "Image",
  "Days after PO issue",
  "Total Qty",
  "Pending Shade",
  "Cutting Date",
  "Remarks",
];


/** Header synonyms -> canonical */
const HEADER_ALIAS_TO_CANON = {
  joborderno: "Job Order No",
  "joborder no": "Job Order No",
  "job order no": "Job Order No",
  jobordeerno: "Job Order No",
  orderno: "Job Order No",
  "jo no": "Job Order No",
  fabric: "Fabric",
  brand: "Brand",
  style: "Style",
  partyname: "Party Name",
  party: "Party Name",
  garmenttype: "Garment Type",
  garment: "Garment Type",
  section: "Section",
  season: "Season",
  directstitching: "Direct Stitching",
  "direct stitching": "Direct Stitching",
  lotno: "Lot No",
  lotnumber: "Lot No",
  "lot number": "Lot No",
  date: "Date",
  status: "Status",
  priority: "Priority",
  cancellationtimestamp: "Cancellation Timestamp",
  "cancellation timestamp": "Cancellation Timestamp",
  cancellationdate: "Cancellation Timestamp",
  "cancellation date": "Cancellation Timestamp",
  cancellationtime: "Cancellation Timestamp",
};

/** ====== UTILS ====== */
const norm = (s) => String(s || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
function normalizeKey(s = "") {
  return norm(s);
}

function formatDateYMDToDDMMMYYYY(dateStr) {
  if (!dateStr) return "";

  const parts = String(dateStr).split("-");
  if (parts.length === 3) {
    const [year, month, day] = parts.map(p => parseInt(p, 10));
    if (!isNaN(year) && !isNaN(month) && !isNaN(day)) {
      const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
        "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      const monthName = monthNames[month - 1] || "";
      return `${day} ${monthName} ${year}`;
    }
  }

  return dateStr;
}

function formatSavedAtToYMD(savedAt) {
  if (!savedAt) return "";
  const d = new Date(savedAt);
  if (isNaN(d.getTime())) return "";

  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const day = d.getDate();
  const monthName = monthNames[d.getMonth()];
  const year = d.getFullYear();

  return `${day} ${monthName} ${year}`;
}

function daysAfter(poDateStr, cuttingDateVal) {
  if (!poDateStr) return "";

  const parseDate = (val) => {
    if (!val) return null;
    if (val instanceof Date && !isNaN(val.getTime())) {
      return new Date(val.getFullYear(), val.getMonth(), val.getDate());
    }
    const str = String(val).trim();
    if (!str) return null;

    const ymdParts = str.split("-");
    if (ymdParts.length === 3) {
      const [y, m, d] = ymdParts.map((p) => parseInt(p, 10));
      if (y && m && d) return new Date(y, m - 1, d);
    }

    const d = new Date(str);
    if (!isNaN(d.getTime())) {
      return new Date(d.getFullYear(), d.getMonth(), d.getDate());
    }
    return null;
  };

  const start = parseDate(poDateStr);
  if (!start) return "";

  let end = null;
  if (cuttingDateVal) {
    end = parseDate(cuttingDateVal);
  }

  // If cutting is NOT done for the lot, use Today's date
  if (!end) {
    const now = new Date();
    end = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  }

  const MS_PER_DAY = 24 * 60 * 60 * 1000;
  const diff = Math.floor((end.getTime() - start.getTime()) / MS_PER_DAY);
  return Number.isFinite(diff) ? String(diff) : "";
}

function formatCancellationDate(ts) {
  if (!ts) return "";
  const s = String(ts).trim();
  if (!s) return "";

  // Match YYYY-MM-DD or YYYY MM DD or YYYY/MM/DD
  const ymdMatch = s.match(/^(\d{4})[-/\s](\d{1,2})[-/\s](\d{1,2})/);
  if (ymdMatch) {
    const y = ymdMatch[1];
    const m = String(ymdMatch[2]).padStart(2, "0");
    const d = String(ymdMatch[3]).padStart(2, "0");
    return `${d}-${m}-${y}`;
  }

  // Match DD-MM-YYYY or DD/MM/YYYY
  const dmyMatch = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if (dmyMatch) {
    const d = String(dmyMatch[1]).padStart(2, "0");
    const m = String(dmyMatch[2]).padStart(2, "0");
    const y = dmyMatch[3];
    return `${d}-${m}-${y}`;
  }

  try {
    const d = new Date(s);
    if (!isNaN(d.getTime())) {
      const day = String(d.getDate()).padStart(2, "0");
      const month = String(d.getMonth() + 1).padStart(2, "0");
      const year = d.getFullYear();
      return `${day}-${month}-${year}`;
    }
  } catch (e) {}

  return s;
}

function convertValuesToObjects(values) {
  if (!values || values.length === 0) return [];
  const rawHeaders = values[0];

  const canonAtIndex = rawHeaders.map((h) => HEADER_ALIAS_TO_CANON[normalizeKey(h)] || null);
  const canonToIndex = {};
  canonAtIndex.forEach((canon, idx) => {
    if (canon && !(canon in canonToIndex)) canonToIndex[canon] = idx;
  });

  return values.slice(1).map((row) => {
    const obj = {};
    [
      "Job Order No",
      "Date",
      "Fabric",
      "Brand",
      "Style",
      "Party Name",
      "Garment Type",
      "Section",
      "Season",
      "Direct Stitching",
      "Lot No",
      "Status",
      "Priority",
      "Cancellation Timestamp",
    ].forEach((canonHeader) => {
      const idx = canonToIndex[canonHeader];
      let value = idx != null ? (row[idx] ?? "") : "";

      if (canonHeader === "Date" && value) {
        value = formatDateYMDToDDMMMYYYY(value);
      }

      obj[canonHeader] = value;
    });

    if (!obj["Cancellation Timestamp"]) {
      rawHeaders.forEach((h, idx) => {
        const normH = normalizeKey(h);
        if (normH.includes("cancellationtimestamp") || normH.includes("cancellationdate") || (normH.includes("cancellation") && normH.includes("time"))) {
          obj["Cancellation Timestamp"] = row[idx] ?? "";
        }
      });
    }

    if (canonToIndex["Date"] != null) {
      const originalDate = row[canonToIndex["Date"]] ?? "";
      obj["PO Date"] = originalDate;
    } else {
      obj["PO Date"] = "";
    }

    obj["Days after PO issue"] = "";
    obj["Total Qty"] = 0;
    obj["Pending Shade"] = "";
    obj["Remarks"] = "";
    obj["Remarks 2"] = "";
    obj["Remarks 3"] = "";
    obj["Cutting Date"] = "";
    return obj;
  });
}

// Helper function to parse date string for comparison
function parseDateString(dateStr) {
  if (!dateStr) return null;
  // Try parsing "DD MMM YYYY" format (e.g., "27 Aug 2025")
  const parts = dateStr.split(" ");
  if (parts.length === 3) {
    const day = parseInt(parts[0], 10);
    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
      "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const month = monthNames.indexOf(parts[1]);
    const year = parseInt(parts[2], 10);
    if (!isNaN(day) && month !== -1 && !isNaN(year)) {
      return new Date(year, month, day);
    }
  }
  return null;
}

/** ====== PDF THEME ====== */
const PDF_THEME = {
  brandName: "Cutting Report",
  primary: [30, 41, 59], // Slate-800
  muted: [71, 85, 105], // Slate-600
  border: [203, 213, 225], // Slate-300
  zebra: [248, 250, 252], // Slate-50
  badgeDoneBG: [240, 253, 244],
  badgeDoneText: [21, 128, 61],
  badgePendingBG: [255, 251, 235],
  badgePendingText: [180, 83, 9],
  badgeIssueBG: [254, 242, 242],
  badgeIssueText: [185, 28, 28],
};

function fmtNum(n) {
  if (n == null || n === "") return "";
  const num = Number(n);
  if (Number.isNaN(num)) return String(n);
  return num.toLocaleString("en-IN");
}

function todayYMD() {
  return new Date().toISOString().slice(0, 10);
}

const COL_INDEX = OUTPUT_COLS.reduce((acc, k, i) => ((acc[k] = i), acc), {});

/** ====== FETCH ====== */
async function fetchSheet({ sheetId, range }) {
  const res = await fetchSheetDataFromBackend(sheetId, range);
  if (res.ok) {
    return { values: res.values || [] };
  }
  throw new Error(`Failed to fetch sheet range: ${range}`);
}

/** ====== INDEX + CUTTING HELPERS ====== */
const getDirectImageUrl = (url) => {
  if (!url) return "";
  const reg = /id=([a-zA-Z0-9_-]+)/;
  const match = url.match(reg);
  if (match && match[1]) {
    return `https://lh3.googleusercontent.com/d/${match[1]}`;
  }
  const dreg = /\/d\/([a-zA-Z0-9_-]+)/;
  const dmatch = url.match(dreg);
  if (dmatch && dmatch[1]) {
    return `https://lh3.googleusercontent.com/d/${dmatch[1]}`;
  }
  return url;
};

function parseIndexRow(header, row, imageUrlIndex = -1) {
  const hmap = {};
  header.forEach((h, i) => (hmap[normalizeKey(h)] = i));
  const get = (key) => {
    const i = hmap[key];
    return i == null || i < 0 ? "" : row[i] ?? "";
  };

  const lot = String(get("lotnumber") || get("lot number") || get("lotno")).trim();
  if (!lot) return null;

  const startRow = parseInt(get("startrow") || "0", 10);
  const numRows = parseInt(get("numrows") || "0", 10);
  const headerCols = parseInt(get("headercols") || "0", 10);
  const fabric = get("fabric");
  const garmentType = get("garmenttype") || get("garment");
  const style = get("style");
  const savedAt = get("savedat");

  const imageUrlRaw = imageUrlIndex !== -1 && row[imageUrlIndex] ? row[imageUrlIndex] : "";
  const imageUrl = getDirectImageUrl(imageUrlRaw);

  const sizes = String(get("sizes") || "")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);

  const shades = String(get("shades") || "")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);

  return { lot, startRow, numRows, headerCols, fabric, garmentType, style, sizes, shades, savedAt, imageUrl };
}

function sliceCuttingMatrix(bigValues, startRow, numRows) {
  if (!Array.isArray(bigValues) || bigValues.length === 0) return [];
  if (!(startRow > 0 && numRows > 0)) return [];
  const r0 = Math.max(0, startRow - 1);
  const r1 = Math.min(bigValues.length - 1, r0 + numRows - 1);
  return bigValues.slice(r0, r1 + 1);
}

function findHeaderRowIndex(windowValues, expectedSizesNorm) {
  // First, look for a row that has "Color" and size columns
  for (let i = 0; i < Math.min(windowValues.length, 20); i++) {
    const row = windowValues[i] || [];
    const rowText = row.map(c => String(c || "").toLowerCase());

    const hasColor = rowText.some(c => c === "color" || c === "shade" || c === "shades");
    const hasCuttingTable = rowText.some(c => c === "cutting table" || c === "cuttingtable");
    const hasSizes = expectedSizesNorm.some(sz => rowText.includes(sz));

    if (hasColor && (hasCuttingTable || hasSizes)) {
      console.log("Found header row at index", i, "with:", rowText);
      return i;
    }
  }

  // Fallback: look for row with most size matches
  let bestMatchIdx = 0;
  let bestMatchCount = 0;

  for (let i = 0; i < Math.min(windowValues.length, 20); i++) {
    const row = windowValues[i] || [];
    const rowText = row.map(c => String(c || "").toLowerCase());
    let matchCount = 0;

    expectedSizesNorm.forEach(sz => {
      if (rowText.includes(sz)) matchCount++;
    });

    if (matchCount > bestMatchCount) {
      bestMatchCount = matchCount;
      bestMatchIdx = i;
    }
  }

  return bestMatchIdx;
}

function calculateTotalPCS(cuttingData, startRow, numRows, sizes = []) {
  if (!cuttingData || cuttingData.length === 0) return 0;
  if (!(startRow > 0 && numRows > 0)) return 0;

  const r0 = Math.max(0, startRow - 1);
  const r1 = Math.min(cuttingData.length - 1, r0 + numRows - 1);
  const windowValues = cuttingData.slice(r0, r1 + 1);

  if (windowValues.length === 0) return 0;

  const normalizedSizes = Array.from(
    new Set((sizes || []).map((s) => normalizeKey(s)).filter(Boolean))
  );

  const findHeaderRowIndex = (windowValues, expectedSizesNorm) => {
    const hasSizeToken = (rowSet) => expectedSizesNorm.some((sz) => rowSet.has(sz));

    for (let i = 0; i < windowValues.length; i++) {
      const row = windowValues[i] || [];
      const set = new Set(row.map((c) => normalizeKey(c)));
      const hasShadeHeader = set.has("color") || set.has("shade") || set.has("shades");
      if (hasShadeHeader && hasSizeToken(set)) return i;
    }

    for (let i = 0; i < windowValues.length; i++) {
      const row = windowValues[i] || [];
      const set = new Set(row.map((c) => normalizeKey(c)));
      let matches = 0;
      expectedSizesNorm.forEach((sz) => {
        if (set.has(sz)) matches++;
      });
      if (matches >= 2) return i;
    }

    return 0;
  };

  const headerRowIdx = findHeaderRowIndex(windowValues, normalizedSizes);
  const header = windowValues[headerRowIdx] || [];

  const hIdx = {};
  header.forEach((h, i) => {
    const k = normalizeKey(h);
    if (k && !(k in hIdx)) hIdx[k] = i;
  });

  const nonSizeColumns = new Set([
    "color",
    "shade",
    "shades",
    "cuttingtable",
    "cutting",
    "table",
    "total",
    "totalpcs",
    "totals",
    "grandtotal",
    "sum",
    "lot",
    "style",
    "fabric",
    "garment",
    "partyname",
    "brand",
    "section",
    "season",
  ]);

  let sizeColIndices = [];
  header.forEach((h, i) => {
    const normalizedHeader = normalizeKey(h);
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

  const shadeColIndex = hIdx["color"] ?? hIdx["shade"] ?? hIdx["shades"] ?? 0;
  let totalQty = 0;

  for (let r = headerRowIdx + 1; r < windowValues.length; r++) {
    const row = windowValues[r] || [];
    const rawShade = String(row[shadeColIndex] || "").trim();
    const shadeKey = normalizeKey(rawShade);

    if (!shadeKey || shadeKey === "total" || shadeKey === "totals" || shadeKey === "grandtotal") {
      continue;
    }

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

function computePendingShades(windowValues, sizes = [], shades = []) {
  if (!windowValues || windowValues.length === 0) {
    return new Set(shades.map(norm));
  }

  const normalizedSizes = Array.from(
    new Set((sizes || []).map((s) => normalizeKey(s)).filter(Boolean))
  );
  const headerRowIdx = findHeaderRowIndex(windowValues, normalizedSizes);
  const header = windowValues[headerRowIdx] || [];

  const hIdx = {};
  header.forEach((h, i) => {
    const k = normalizeKey(h);
    if (k && !(k in hIdx)) hIdx[k] = i;
  });

  const shadeColIndex = hIdx["color"] ?? hIdx["shade"] ?? hIdx["shades"] ?? 0;

  const nonSizeColumns = new Set([
    "color",
    "shade",
    "shades",
    "cuttingtable",
    "cutting",
    "table",
    "total",
    "totalpcs",
    "totals",
    "grandtotal",
    "sum",
    "lot",
    "style",
    "fabric",
    "garment",
    "partyname",
    "brand",
    "section",
    "season",
  ]);

  let sizeColIndices = [];
  header.forEach((h, i) => {
    const normalizedHeader = normalizeKey(h);
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

  if (sizeColIndices.length === 0) {
    return new Set(shades.map(norm));
  }

  const shadeStats = new Map();

  for (let r = headerRowIdx + 1; r < windowValues.length; r++) {
    const row = windowValues[r] || [];
    const rawShade = String(row[shadeColIndex] || "").trim();
    const shadeKey = normalizeKey(rawShade);

    if (!shadeKey || shadeKey === "total" || shadeKey === "totals" || shadeKey === "grandtotal") {
      continue;
    }

    let hasPositiveData = false;
    let hasAnyData = false;
    let totalForThisRow = 0;

    sizeColIndices.forEach((c) => {
      const raw = row[c];
      if (raw != null && raw !== "") {
        hasAnyData = true;
        const n = parseFloat(String(raw).replace(/,/g, ""));
        if (!isNaN(n)) {
          totalForThisRow += n;
          if (n > 0) {
            hasPositiveData = true;
          }
        }
      }
    });

    if (hasAnyData) {
      if (hasPositiveData) {
        shadeStats.set(shadeKey, "found-with-data");
      } else {
        if (!shadeStats.has(shadeKey) || shadeStats.get(shadeKey) === "not-found") {
          shadeStats.set(shadeKey, "found-all-zero");
        }
      }
    } else {
      if (!shadeStats.has(shadeKey)) {
        shadeStats.set(shadeKey, "found-no-data");
      }
    }
  }

  const pendingShadeKeys = new Set();
  const expectedShadeKeys = (shades || []).map((sh) => normalizeKey(sh));

  expectedShadeKeys.forEach((shadeKey) => {
    const status = shadeStats.get(shadeKey);

    if (!status || status === "found-no-data") {
      pendingShadeKeys.add(shadeKey);
    }
  });

  return pendingShadeKeys;
}

function extractCuttingTables(windowValues, sizes = []) {
  if (!windowValues || windowValues.length === 0) return [];

  const normalizedSizes = Array.from(
    new Set((sizes || []).map((s) => normalizeKey(s)).filter(Boolean))
  );
  const headerRowIdx = findHeaderRowIndex(windowValues, normalizedSizes);
  if (headerRowIdx >= windowValues.length) return [];

  const header = windowValues[headerRowIdx] || [];

  const hIdx = {};
  header.forEach((h, i) => {
    const k = normalizeKey(h);
    if (k && !(k in hIdx)) hIdx[k] = i;
  });

  // Look for cutting table column - exact matches for "Cutting Table"
  let tableColIndex = -1;
  for (let i = 0; i < header.length; i++) {
    const headerText = String(header[i] || "").trim();
    const normalizedHeader = normalizeKey(headerText);
    if (normalizedHeader === "cuttingtable" ||
      normalizedHeader === "cutting table" ||
      headerText === "Cutting Table") {
      tableColIndex = i;
      break;
    }
  }

  if (tableColIndex === -1) {
    console.log("Cutting Table column not found in headers:", header);
    return [];
  }

  // Find size columns (M, L, XL, XXL, etc.)
  const sizeColumns = [];
  const sizePatterns = ['m', 'l', 'xl', 'xxl', '2xl', '3xl', '4xl', 's', 'xs', 'xxs'];

  for (let i = 0; i < header.length; i++) {
    const headerText = String(header[i] || "").trim().toLowerCase();
    if (sizePatterns.includes(headerText) ||
      (headerText.match(/^[0-9]+$/) && parseInt(headerText) > 0) ||
      sizePatterns.some(pattern => headerText === pattern)) {
      sizeColumns.push(i);
    }
  }

  // Also check for any numeric headers that might be sizes
  for (let i = 0; i < header.length; i++) {
    const headerText = String(header[i] || "").trim();
    const num = parseFloat(headerText);
    if (!isNaN(num) && num > 0 && num < 100 && !sizeColumns.includes(i)) {
      sizeColumns.push(i);
    }
  }

  if (sizeColumns.length === 0) {
    console.log("No size columns found");
    return [];
  }

  const uniqueTables = new Set();

  // Start from row after header
  for (let r = headerRowIdx + 1; r < windowValues.length; r++) {
    const row = windowValues[r] || [];

    // Skip empty rows
    if (!row || row.length === 0) continue;

    const shadeValue = String(row[0] || "").trim().toLowerCase();
    // Skip total rows
    if (shadeValue === "total" || shadeValue === "totals" || shadeValue === "grand total") {
      continue;
    }

    // Check if this row has any positive quantity in size columns
    let hasPositiveData = false;
    for (const colIdx of sizeColumns) {
      if (colIdx < row.length) {
        const raw = row[colIdx];
        if (raw != null && raw !== "") {
          const n = parseFloat(String(raw).replace(/,/g, ""));
          if (!isNaN(n) && n > 0) {
            hasPositiveData = true;
            break;
          }
        }
      }
    }

    if (hasPositiveData && tableColIndex < row.length) {
      const tableValue = String(row[tableColIndex] || "").trim();
      if (tableValue && tableValue !== "" && tableValue !== "0") {
        // Handle cases where multiple tables might be in one cell (e.g., "1,2" or "1 2")
        const tables = tableValue.split(/[,\s]+/).filter(t => t && t !== "" && t !== "0");
        tables.forEach(t => uniqueTables.add(t));
      }
    }
  }

  const result = Array.from(uniqueTables).sort((a, b) => {
    const numA = parseFloat(a);
    const numB = parseFloat(b);
    if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
    return String(a).localeCompare(String(b));
  });

  console.log("Extracted cutting tables:", result); // Debug log
  return result;
}

// Multi-select dropdown component
const MultiSelectDropdown = ({ options, selectedValues, onChange, placeholder, disabled }) => {
  const [isOpen, setIsOpen] = useState(false);
  const dropdownRef = useRef(null);

  useEffect(() => {
    const handleClickOutside = (event) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const toggleOption = (value) => {
    if (selectedValues.includes(value)) {
      onChange(selectedValues.filter(v => v !== value));
    } else {
      onChange([...selectedValues, value]);
    }
  };

  const clearAll = () => {
    onChange([]);
  };

  return (
    <div className="multi-select-dropdown" ref={dropdownRef}>
      <div
        className={`multi-select-trigger ${disabled ? 'disabled' : ''}`}
        onClick={() => !disabled && setIsOpen(!isOpen)}
      >
        <div className="multi-select-values">
          {selectedValues.length === 0 ? (
            <span className="placeholder">{placeholder}</span>
          ) : (
            <span className="selected-count">{selectedValues.length} selected</span>
          )}
        </div>
        <span className="dropdown-arrow">{isOpen ? '▲' : '▼'}</span>
      </div>
      {isOpen && !disabled && (
        <div className="multi-select-options">
          <div className="multi-select-actions">
            <button type="button" onClick={clearAll} className="clear-all-btn">Clear All</button>
          </div>
          {options.map(option => (
            <label key={option} className="multi-select-option">
              <input
                type="checkbox"
                checked={selectedValues.includes(option)}
                onChange={() => toggleOption(option)}
              />
              <span>{option}</span>
            </label>
          ))}
        </div>
      )}
    </div>
  );
};

/* ================== FINANCIAL YEAR HELPERS ================== */
const getFinancialYearFromDate = (dateStr) => {
  if (!dateStr) return null;
  let d = null;
  if (dateStr instanceof Date) {
    d = dateStr;
  } else if (typeof dateStr === "string") {
    const s = dateStr.trim();
    if (!s) return null;
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
      const parts = s.split("-");
      d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    } else if (/^\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{4}/.test(s)) {
      const parts = s.split(/[\/\-\.]/);
      d = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
    } else {
      d = new Date(s);
    }
  }
  if (!d || isNaN(d.getTime())) return null;
  const month = d.getMonth();
  const year = d.getFullYear();
  return month >= 3 ? `${year}-${year + 1}` : `${year - 1}-${year}`;
};

const getCurrentFinancialYear = () => {
  return getFinancialYearFromDate(new Date());
};

/** ====== COMPONENT ====== */
export default function CuttingStatsReport() {
  const [rows, setRows] = useState([]);
  const [lotFilter, setLotFilter] = useState("");
  const [partyFilter, setPartyFilter] = useState("");
  const [garmentFilter, setGarmentFilter] = useState([]); // Changed to array for multi-select
  const [seasonFilter, setSeasonFilter] = useState("");
  const [sectionFilter, setSectionFilter] = useState("");
  const [brandFilter, setBrandFilter] = useState([]); // Changed to array for multi-select
  const [priorityFilter, setPriorityFilter] = useState("");
  const [styleFilter, setStyleFilter] = useState([]);
  const [fabricFilter, setFabricFilter] = useState([]);
  const [directStitchingFilter, setDirectStitchingFilter] = useState([]);
  const [financialYearFilter, setFinancialYearFilter] = useState(getCurrentFinancialYear());
  const [daysFilter, setDaysFilter] = useState("all");
  const [loading, setLoading] = useState(true);
  const [refreshing, setRefreshing] = useState(false);
  const [err, setErr] = useState("");
  const [lastUpdated, setLastUpdated] = useState(null);

  // Date range filter states
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");

  const [dialogOpen, setDialogOpen] = useState(false);
  const [dialogLot, setDialogLot] = useState("");
  const [dialogShades, setDialogShades] = useState([]);

  const [pendingListByLot, setPendingListByLot] = useState({});
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(25);
  const [selectedRemarks, setSelectedRemarks] = useState(new Set());
  const [viewImageSrc, setViewImageSrc] = useState(null);

  const abortRef = useRef(null);

  const splitRemarks = (raw) =>
    String(raw || "")
      .split("|")
      .map((p) => p.trim())
      .filter(Boolean);

  // Get distinct values for filters
  const distinctParties = useMemo(() => {
    const set = new Set();
    rows.forEach((r) => {
      const party = String(r["Party Name"] || "").trim();
      if (party) set.add(party);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [rows]);

  const distinctSections = useMemo(() => {
    const set = new Set();
    rows.forEach((r) => {
      const section = String(r["Section"] || "").trim();
      if (section) set.add(section);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [rows]);

  const distinctBrands = useMemo(() => {
    const set = new Set();
    rows.forEach((r) => {
      const brand = String(r["Brand"] || "").trim();
      if (brand) set.add(brand);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [rows]);

  const distinctGarments = useMemo(() => {
    const set = new Set();
    rows.forEach((r) => {
      const garment = String(r["Garment Type"] || "").trim();
      if (garment) set.add(garment);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [rows]);

  const distinctSeasons = useMemo(() => {
    const set = new Set();
    rows.forEach((r) => {
      const season = String(r["Season"] || "").trim();
      if (season) set.add(season);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [rows]);

  const distinctPriorities = useMemo(() => {
    const set = new Set();
    rows.forEach((r) => {
      const priority = String(r["Priority"] || "").trim();
      if (priority) set.add(priority);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [rows]);

  const distinctRemarks = useMemo(() => {
    const set = new Set();
    rows.forEach((r) => {
      splitRemarks(r.Remarks).forEach((t) => {
        if (norm(t).startsWith("cancel") || norm(t).includes("cancel")) {
          set.add("Cancel");
        } else {
          set.add(t);
        }
      });
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [rows]);

  const distinctStyles = useMemo(() => {
    const set = new Set();
    rows.forEach((r) => {
      const style = String(r["Style"] || "").trim();
      if (style) set.add(style);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [rows]);

  const distinctFabrics = useMemo(() => {
    const set = new Set();
    rows.forEach((r) => {
      const fabric = String(r["Fabric"] || "").trim();
      if (fabric) set.add(fabric);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [rows]);

  const distinctDirectStitching = useMemo(() => {
    const set = new Set();
    rows.forEach((r) => {
      const ds = String(r["Direct Stitching"] || "").trim();
      if (ds) set.add(ds);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [rows]);

  const availableFYs = useMemo(() => {
    const set = new Set();
    const currentFY = getCurrentFinancialYear();
    if (currentFY) set.add(currentFY);
    rows.forEach((r) => {
      const rowDate = r["Cutting Date"] || r["Date"];
      const fy = getFinancialYearFromDate(rowDate);
      if (fy) set.add(fy);
    });
    return Array.from(set).sort().reverse();
  }, [rows]);

  const toggleRemark = (label) => {
    setSelectedRemarks((prev) => {
      const next = new Set(prev);
      if (next.has(label)) next.delete(label);
      else next.add(label);
      return next;
    });
  };

  const clearRemarks = () => setSelectedRemarks(new Set());

  const clearAllFilters = () => {
    setLotFilter("");
    setPartyFilter("");
    setGarmentFilter([]);
    setSeasonFilter("");
    setSectionFilter("");
    setBrandFilter([]);
    setPriorityFilter("");
    setStyleFilter([]);
    setFabricFilter([]);
    setDirectStitchingFilter([]);
    setStartDate("");
    setEndDate("");
    setDaysFilter("all");
    clearRemarks();
  };

  const openPendingDialog = (lot) => {
    const list = pendingListByLot[lot] || [];
    setDialogLot(lot);
    setDialogShades(list);
    setDialogOpen(true);
  };

  const closeDialog = () => {
    setDialogOpen(false);
    setDialogLot("");
    setDialogShades([]);
  };

  const loadData = async (mode = "initial") => {
    if (mode === "initial") {
      setLoading(true);
    } else {
      setRefreshing(true);
    }
    setErr("");

    try {
      abortRef.current?.abort();
    } catch { }
    const ctrl = new AbortController();
    abortRef.current = ctrl;

    try {
      const job = await fetchSheet(
        {
          sheetId: JOB_SHEET_ID,
          range: JOB_RANGE,
          apiKey: API_KEY,
          signal: ctrl.signal,
        }
      );
      let jobRows = convertValuesToObjects(job.values);

      const isCancel = (s) => {
        if (!s) return false;
        const sn = norm(s);
        return sn.startsWith("cancel") || sn.includes("cancel");
      };

      const idxRes = await fetchSheet(
        {
          sheetId: BUDGET_SHEET_ID,
          range: INDEX_RANGE,
          apiKey: API_KEY,
          signal: ctrl.signal,
        }
      );
      const idxValues = idxRes.values || [];
      const idxHeader = idxValues[0] || [];
      const indexMap = new Map();
      let imageUrlIndex = -1;
      idxHeader.forEach((header, index) => {
        const headerLower = String(header || "").toLowerCase().trim();
        if (headerLower === 'image url' || headerLower === 'imageurl' || headerLower === 'image' || headerLower.includes('image')) {
          imageUrlIndex = index;
        }
      });

      for (let i = 1; i < idxValues.length; i++) {
        const entry = parseIndexRow(idxHeader, idxValues[i], imageUrlIndex);
        if (entry) indexMap.set(entry.lot, entry);
      }

      const cuttingRes = await fetchSheet(
        {
          sheetId: BUDGET_SHEET_ID,
          range: CUTTING_BIG_RANGE,
          apiKey: API_KEY,
          signal: ctrl.signal,
        }
      );
      const bigCuttingValues = cuttingRes.values || [];

      const lots = Array.from(
        new Set(jobRows.map((r) => String(r["Lot No"] || "").trim()).filter(Boolean))
      );
      const lotToSummary = new Map();
      const pendingListTmp = {};

      for (const lot of lots) {
        const ix = indexMap.get(lot);
        if (!ix) {
          lotToSummary.set(lot, {
            totalQty: 0,
            remarks: "",
            remarks2: "",
            remarks3: lot ? "Fabric Issue Pending" : "",
            cuttingDate: "",
            cuttingTables: [],
            imageUrl: ""
          });
          pendingListTmp[lot] = [];
          continue;
        }

        const cuttingDate = formatSavedAtToYMD(ix.savedAt);
        const totalQty = calculateTotalPCS(bigCuttingValues, ix.startRow, ix.numRows, ix.sizes);
        const window = sliceCuttingMatrix(bigCuttingValues, ix.startRow, ix.numRows);
        const pendingShadeKeys = computePendingShades(window, ix.sizes, ix.shades);

        // Extract cutting tables
        const cuttingTables = extractCuttingTables(window, ix.sizes);

        const shadeKeyToOriginal = new Map((ix.shades || []).map((sh) => [norm(sh), sh]));
        const pendingList = Array.from(pendingShadeKeys).map(
          (k) => shadeKeyToOriginal.get(k) || k
        );

        let remarks = "";
        let remarks2 = "";
        let remarks3 = "";

        if (pendingShadeKeys.size > 0) {
          remarks2 = "Colour Pending";
        } else {
          remarks = "Cutting Done";
        }

        lotToSummary.set(lot, {
          totalQty,
          remarks,
          remarks2,
          remarks3,
          cuttingDate,
          cuttingTables,
          imageUrl: ix.imageUrl || ""
        });
        pendingListTmp[lot] = pendingList;
      }
      const merged = jobRows.map((r) => {
        const lot = String(r["Lot No"] || "").trim();
        const statusVal = String(r.Status ?? r.status ?? "").trim();
        const isCancelled = isCancel(statusVal);

        const sum =
          lotToSummary.get(lot) ||
          {
            totalQty: 0,
            remarks: "",
            remarks2: "",
            remarks3: lot ? "Fabric Issue Pending" : "",
            cuttingDate: "",
            cuttingTables: [],
            imageUrl: ""
          };
        const days = daysAfter(r["PO Date"] || r["Date"], sum.cuttingDate);

        let mergedRemarks = "";
        if (isCancelled) {
          const cancelDate = formatCancellationDate(r["Cancellation Timestamp"]);
          mergedRemarks = cancelDate ? `Cancel(${cancelDate})` : "Cancel";
        } else {
          const remarksList = [sum.remarks, sum.remarks2, sum.remarks3].filter(Boolean);
          mergedRemarks = remarksList.join(" | ");
        }

        // Format cutting tables for display
        const cuttingTablesDisplay = sum.cuttingTables && sum.cuttingTables.length > 0
          ? sum.cuttingTables.join(", ")
          : "";

        return {
          ...r,
          "Days after PO issue": days ?? "",
          "Total Qty": isCancelled ? (r["Total Qty"] || r.Quantity || sum.totalQty || 0) : sum.totalQty,
          "Pending Shade": "",
          "Cutting Date": sum.cuttingDate || "",
          "Cutting Table": cuttingTablesDisplay,
          Remarks: mergedRemarks,
          imageUrl: sum.imageUrl || r.imageUrl || ""
        };
      });

      setRows(merged);
      setPendingListByLot(pendingListTmp);
      setLastUpdated(new Date().toLocaleString());
    } catch (e) {
      if (e?.name === "AbortError") {
        // do nothing
      } else {
        console.error(e);
        setErr(e.message || "Failed to load sheet");
      }
    } finally {
      if (mode === "initial") setLoading(false);
      else setRefreshing(false);
    }
  };

  useEffect(() => {
    loadData("initial");
    return () => {
      try {
        abortRef.current?.abort();
      } catch { }
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const filtered = useMemo(() => {
    const lotQ = lotFilter.trim().toLowerCase();
    const partyQ = partyFilter.trim().toLowerCase();
    const garmentQ = garmentFilter.map(g => g.toLowerCase());
    const seasonQ = seasonFilter.trim().toLowerCase();
    const sectionQ = sectionFilter.trim().toLowerCase();
    const brandQ = brandFilter.map(b => b.toLowerCase());
    const priorityQ = priorityFilter.trim().toLowerCase();
    const styleQ = styleFilter.map(s => s.toLowerCase());
    const fabricQ = fabricFilter.map(f => f.toLowerCase());
    const directStitchingQ = directStitchingFilter.map(d => d.toLowerCase());
    const selected = selectedRemarks;

    return rows.filter((r) => {
      if (daysFilter !== "all") {
        const daysVal = parseFloat(r["Days after PO issue"]);
        if (isNaN(daysVal)) return false;
        if (daysFilter === "red" && daysVal <= 2) return false;
        if (daysFilter === "green" && daysVal > 2) return false;
        if (daysFilter === "gt5" && daysVal <= 5) return false;
        if (daysFilter === "gt10" && daysVal <= 10) return false;
      }
      if (financialYearFilter && financialYearFilter !== "ALL") {
        const rowDate = r["Cutting Date"] || r["Date"];
        const rowFY = getFinancialYearFromDate(rowDate);
        if (rowFY !== financialYearFilter) return false;
      }
      if (lotQ) {
        const lot = String(r["Lot No"] || "").toLowerCase();
        if (!lot.includes(lotQ)) return false;
      }
      if (partyQ) {
        const party = String(r["Party Name"] || "").toLowerCase();
        if (!party.includes(partyQ)) return false;
      }
      if (garmentQ.length > 0) {
        const garment = String(r["Garment Type"] || "").toLowerCase();
        if (!garmentQ.includes(garment)) return false;
      }
      if (seasonQ) {
        const season = String(r["Season"] || "").toLowerCase();
        if (!season.includes(seasonQ)) return false;
      }
      if (sectionQ) {
        const section = String(r["Section"] || "").toLowerCase();
        if (!section.includes(sectionQ)) return false;
      }
      if (brandQ.length > 0) {
        const brand = String(r["Brand"] || "").toLowerCase();
        if (!brandQ.includes(brand)) return false;
      }
      if (priorityQ) {
        const priority = String(r["Priority"] || "").toLowerCase();
        if (!priority.includes(priorityQ)) return false;
      }
      if (styleQ.length > 0) {
        const style = String(r["Style"] || "").toLowerCase();
        if (!styleQ.includes(style)) return false;
      }
      if (fabricQ.length > 0) {
        const fabric = String(r["Fabric"] || "").toLowerCase();
        if (!fabricQ.includes(fabric)) return false;
      }
      if (directStitchingQ.length > 0) {
        const ds = String(r["Direct Stitching"] || "").toLowerCase();
        if (!directStitchingQ.includes(ds)) return false;
      }

      // Date range filter
      if (startDate || endDate) {
        const dateValue = r["Date"];
        const parsedDate = parseDateString(dateValue);

        if (parsedDate) {
          if (startDate) {
            const start = new Date(startDate);
            if (parsedDate < start) return false;
          }
          if (endDate) {
            const end = new Date(endDate);
            if (parsedDate > end) return false;
          }
        } else if (dateValue) {
          // If date parsing fails, try to extract from the string
          const yearMatch = dateValue.match(/\d{4}/);
          if (yearMatch) {
            const year = parseInt(yearMatch[0]);
            if (startDate && new Date(startDate).getFullYear() > year) return false;
            if (endDate && new Date(endDate).getFullYear() < year) return false;
          }
        }
      }

      if (selected.size > 0) {
        const tokens = splitRemarks(r.Remarks);
        const tokenSet = new Set(tokens);
        const hasCancel = tokens.some(t => norm(t).startsWith("cancel") || norm(t).includes("cancel"));

        let matchesAny = false;
        for (const sel of selected) {
          if (sel === "Cancel" && hasCancel) {
            matchesAny = true;
            break;
          }
          if (tokenSet.has(sel)) {
            matchesAny = true;
            break;
          }
        }
        if (!matchesAny) return false;
      }

      return true;
    });
  }, [
    rows,
    financialYearFilter,
    daysFilter,
    lotFilter,
    partyFilter,
    garmentFilter,
    seasonFilter,
    sectionFilter,
    brandFilter,
    priorityFilter,
    styleFilter,
    fabricFilter,
    directStitchingFilter,
    startDate,
    endDate,
    selectedRemarks,
  ]);

  const redZoneCount = useMemo(() => {
    return rows.filter((r) => {
      const val = parseFloat(r["Days after PO issue"]);
      return !isNaN(val) && val > 2;
    }).length;
  }, [rows]);

  useEffect(() => {
    setPage(1);
  }, [lotFilter, partyFilter, garmentFilter, seasonFilter, pageSize, rows, startDate, endDate]);

  const pageCount = Math.max(1, Math.ceil(filtered.length / pageSize));
  const currentPage = Math.min(page, pageCount);
  const startIdx = (currentPage - 1) * pageSize;
  const endIdx = startIdx + pageSize;
  const pagedRows = filtered.slice(startIdx, endIdx);

  const buildExportRows = (sourceRows = filtered) => {
    return sourceRows.map((r) => {
      const lot = String(r["Lot No"] || "").trim();
      const pending = (pendingListByLot[lot] || []).join(", ");
      const obj = {};
      OUTPUT_COLS.forEach((col) => {
        if (col === "Pending Shade") obj[col] = pending;
        else if (col === "Image") obj[col] = r.imageUrl || "";
        else obj[col] = r[col] ?? "";
      });
      return obj;
    });
  };

  const handleExportExcel = () => {
    const exportRows = buildExportRows(filtered);
    const ws = XLSX.utils.json_to_sheet(exportRows, { header: OUTPUT_COLS });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Cutting Stats");
    XLSX.writeFile(wb, `Cutting_Stats_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const handleExportPDF = () => {
    const exportRows = buildExportRows(filtered);
    const HEADER_ORDER = [
      "Job Order No", "Date", "Lot No", "Fabric", "Brand", "Style",
      "Garment Type", "Party Name", "Section", "Season", "Direct Stitching",
      "Days after PO issue", "Total Qty", "Pending Shade", "Cutting Date", "Remarks"
    ];
    const finalHeaders = HEADER_ORDER.filter(h => OUTPUT_COLS.includes(h));

    // Summary stats
    let totalQty = 0;
    let mohitCount = 0;
    let highPriorityCount = 0;
    const parties = new Set();
    exportRows.forEach(row => {
      totalQty += parseInt(row["Total Qty"]) || 0;
      const pn = (row["Party Name"] || "").toString().trim();
      if (pn) {
        parties.add(pn);
        if (pn.toLowerCase().includes("mohit")) mohitCount++;
      }
      if ((row["Priority"] || "").toString().toUpperCase().includes("REPEATED_LOT")) {
        highPriorityCount++;
      }
    });

    const generatedDate = new Date().toLocaleString("en-IN", {
      weekday: "short", year: "numeric", month: "short",
      day: "numeric", hour: "2-digit", minute: "2-digit"
    });

    // Build remarks badges HTML
    const buildRemarksBadges = (rawRemarks) => {
      if (!rawRemarks) return "";
      const parts = String(rawRemarks).split("|")
        .map(p => p.trim())
        .filter(p => p && !p.includes("undefined"));
      const seen = new Set();
      const unique = [];
      parts.forEach(p => {
        const k = p.toLowerCase();
        if (!seen.has(k)) { seen.add(k); unique.push(p); }
      });
      return unique.map(part => {
        let cls = "badge-default";
        if (/done|completed|finished/i.test(part)) cls = "badge-done";
        else if (/cancel/i.test(part)) cls = "badge-cancel";
        else if (/issue|problem|error|fabric/i.test(part)) cls = "badge-issue";
        else if (/pending|waiting|hold/i.test(part)) cls = "badge-pending";
        else if (/cutting/i.test(part)) cls = "badge-cutting";
        return `<span class="badge ${cls}">${part}</span>`;
      }).join("");
    };

    // Build table rows
    const tableRowsHtml = exportRows.map((row, i) => {
      const isRepeated = (row["Priority"] || "").toString().toUpperCase().includes("REPEATED_LOT");
      const partyRaw = (row["Party Name"] || "").toString().trim();
      const partyDisplay = partyRaw.toLowerCase().includes("mohit") ? "MH" : partyRaw;
      const lotNo = (row["Lot No"] || "").toString().trim();
      const pending = (pendingListByLot[lotNo] || []).join(", ");

      const cells = finalHeaders.map(h => {
        if (h === "Lot No") {
          const star = isRepeated ? `<span class="star">&#9733;</span> ` : "";
          return `<td class="lot-cell">${star}<strong style="color:#dc2626">${lotNo}</strong></td>`;
        }
        if (h === "Party Name") return `<td>${partyDisplay}</td>`;
        if (h === "Total Qty") return `<td><strong>${fmtNum(row[h])}</strong></td>`;
        if (h === "Pending Shade") return `<td class="shade-cell">${pending}</td>`;
        if (h === "Remarks") return `<td class="remarks-cell">${buildRemarksBadges(row[h])}</td>`;
        const val = String(row[h] ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
        return `<td>${val}</td>`;
      }).join("");

      return `<tr class="${i % 2 === 1 ? "alt-row" : ""}">${cells}</tr>`;
    }).join("");

    const headerRowHtml = finalHeaders.map(h => `<th>${h}</th>`).join("");

    const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<title>Cutting Report ${todayYMD()}</title>
<style>
@page { size: A3 landscape; margin: 10mm; }
*{ box-sizing:border-box; margin:0; padding:0; }
body{ font-family:'Times New Roman',Times,serif; font-size:8.5pt; color:#000; background:#fff; }

.report-header{ display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:6px; }
.report-title{ font-size:15pt; font-weight:bold; color:#1e2950; }
.report-date{ font-size:7.5pt; color:#475569; margin-top:2px; }

.summary-box{ display:flex; gap:24px; background:#f0f9ff; border:1px solid #bae6fd; border-radius:5px; padding:7px 14px; margin-bottom:7px; flex-wrap:wrap; align-items:center; }
.summary-item{ display:flex; flex-direction:column; }
.summary-label{ font-size:6.5pt; color:#64748b; text-transform:uppercase; letter-spacing:0.3px; }
.summary-value{ font-size:12pt; font-weight:bold; color:#1e40af; line-height:1.2; }
.legend{ margin-left:auto; font-size:7pt; color:#64748b; display:flex; align-items:center; gap:4px; }
.star{ color:#f59e0b; font-size:10pt; }

table{ width:100%; border-collapse:collapse; table-layout:fixed; }
thead{ display:table-header-group; }
tr{ page-break-inside:avoid; }

th{ background:#4f46e5; color:#fff; font-size:7pt; font-weight:bold; padding:5px 2px; text-align:center; border:0.4px solid #a5b4fc; word-break:break-word; }
td{ font-size:7.5pt; padding:3px 2px; text-align:center; vertical-align:middle; border:0.4px solid #cbd5e1; word-break:break-word; }
tr.alt-row td{ background:#f8fafc; }

th:nth-child(1), td:nth-child(1) { width:42px; }
th:nth-child(2), td:nth-child(2) { width:58px; }
th:nth-child(3), td:nth-child(3) { width:52px; }
th:nth-child(4), td:nth-child(4) { width:80px; }
th:nth-child(5), td:nth-child(5) { width:78px; }
th:nth-child(6), td:nth-child(6) { width:86px; }
th:nth-child(7), td:nth-child(7) { width:70px; }
th:nth-child(8), td:nth-child(8) { width:32px; }
th:nth-child(9), td:nth-child(9) { width:42px; }
th:nth-child(10),td:nth-child(10){ width:46px; }
th:nth-child(11),td:nth-child(11){ width:50px; }
th:nth-child(12),td:nth-child(12){ width:44px; }
th:nth-child(13),td:nth-child(13){ width:44px; }
th:nth-child(14),td:nth-child(14){ width:128px; }
th:nth-child(15),td:nth-child(15){ width:66px; }
th:nth-child(16),td:nth-child(16){ width:86px; vertical-align:top; }

.lot-cell{ vertical-align:middle; }
.shade-cell{ font-size:6.5pt; color:#b45309; text-align:left; padding:3px 3px; }
.remarks-cell{ vertical-align:top; text-align:left; padding:3px 3px; }

.badge{ display:inline-block; font-size:6pt; font-weight:bold; padding:1px 3px; border-radius:3px; margin:1px; white-space:nowrap; }
.badge-done   { background:#dcfce7; color:#166534; }
.badge-cancel { background:#fee2e2; color:#991b1b; font-weight:bold; }
.badge-issue  { background:#fee2e2; color:#991b1b; }
.badge-pending{ background:#fef9c3; color:#854d0e; }
.badge-cutting{ background:#dbeafe; color:#1e40af; }
.badge-default{ background:#f1f5f9; color:#334155; }

.footer{ margin-top:5px; display:flex; justify-content:space-between; font-size:6.5pt; color:#94a3b8; border-top:0.4px solid #cbd5e1; padding-top:3px; }

@media print{
  .no-print{ display:none !important; }
  body{ -webkit-print-color-adjust:exact; print-color-adjust:exact; }
}
.print-btn{ display:inline-flex; align-items:center; gap:6px; padding:8px 18px; background:#4f46e5; color:#fff; border:none; border-radius:6px; font-size:13px; font-weight:600; cursor:pointer; margin-bottom:10px; }
.print-btn:hover{ background:#4338ca; }
</style>
</head>
<body>
<div class="no-print" style="padding:10px;">
  <button class="print-btn" onclick="window.print()">&#128424; Print / Save as PDF</button>
  <span style="font-size:11px;color:#64748b;margin-left:8px;">
    In print dialog &rarr; <strong>Destination: Save as PDF</strong> &rarr; Paper: A3 &rarr; Layout: Landscape
  </span>
</div>

<div class="report-header">
  <div>
    <div class="report-title">Cutting Report</div>
    <div class="report-date">Generated: ${generatedDate}</div>
  </div>
  <div style="font-size:7.5pt;color:#475569;text-align:right;">Total Rows: ${exportRows.length}</div>
</div>

<div class="summary-box">
  <div class="summary-item"><span class="summary-label">Total Lots</span><span class="summary-value">${exportRows.length}</span></div>
  <div class="summary-item"><span class="summary-label">Total Quantity</span><span class="summary-value">${fmtNum(totalQty)}</span></div>
  <div class="summary-item"><span class="summary-label">Unique Parties</span><span class="summary-value">${parties.size}</span></div>
  <div class="summary-item"><span class="summary-label">MH Lots</span><span class="summary-value">${mohitCount}</span></div>
  <div class="legend"><span class="star">&#9733;</span> Repeated Lots (${highPriorityCount})</div>
</div>

<table>
  <thead><tr>${headerRowHtml}</tr></thead>
  <tbody>${tableRowsHtml}</tbody>
</table>

<div class="footer">
  <span>Cutting Stats Report</span>
  <span>${generatedDate}</span>
</div>
</body>
</html>`;

    const printWindow = window.open("", "_blank", "width=1400,height=900");
    if (!printWindow) {
      alert("Popup blocked! Please allow popups for this site and try again.");
      return;
    }
    printWindow.document.open();
    printWindow.document.write(html);
    printWindow.document.close();
    printWindow.onload = () => { setTimeout(() => printWindow.print(), 400); };
  };

  const goBack = () => {
    if (typeof window !== "undefined" && window.history.length > 1) {
      window.history.back();
    }
  };

  return (
    <div className="cutting-stats-container">
      <style>{`
        .cutting-stats-container {
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

        /* Loader */
        .loader-overlay {
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

        .loader-spinner {
          width: 56px;
          height: 56px;
          border-radius: 50%;
          background: conic-gradient(from 0deg, #6366f1, #ec4899, #10b981, #6366f1);
          animation: spin 1.2s linear infinite;
          padding: 4px;
          mask: radial-gradient(farthest-side, transparent calc(100% - 4px), #fff 0);
          -webkit-mask: radial-gradient(farthest-side, transparent calc(100% - 4px), #fff 0);
          margin-bottom: 20px;
        }

        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }

        .loader-text {
          font-size: 1.1rem;
          font-weight: 700;
          color: #ffffff;
          letter-spacing: -0.01em;
        }

        /* Main Layout */
        .main-content {
          max-width: 100%;
          margin: 0 auto;
        }

        /* Header Card */
        .header {
          background: linear-gradient(135deg, #1e1b4b 0%, #312e81 40%, #4338ca 100%);
          border-radius: 24px;
          padding: 32px 36px;
          margin-bottom: 28px;
          border: 1px solid rgba(255, 255, 255, 0.1);
          box-shadow: 0 20px 40px -15px rgba(30, 27, 75, 0.25);
          position: relative;
          overflow: hidden;
        }

        .header::before {
          content: '';
          position: absolute;
          top: -50%;
          right: -10%;
          width: 300px;
          height: 300px;
          background: radial-gradient(circle, rgba(99, 102, 241, 0.25) 0%, transparent 70%);
          pointer-events: none;
        }

        .header-content {
          display: flex;
          justify-content: space-between;
          align-items: center;
          position: relative;
          z-index: 1;
        }

        .header-title {
          margin: 0 0 6px 0;
          font-size: 2rem;
          font-weight: 800;
          color: #ffffff;
          letter-spacing: -0.02em;
          display: flex;
          align-items: center;
          gap: 12px;
        }

        .header-subtitle {
          margin: 0;
          font-size: 0.95rem;
          color: #c7d2fe;
          font-weight: 500;
        }

        .stats-grid {
          display: flex;
          gap: 16px;
        }

        .stat-card {
          text-align: center;
          padding: 14px 24px;
          background: rgba(255, 255, 255, 0.12);
          backdrop-filter: blur(12px);
          -webkit-backdrop-filter: blur(12px);
          border-radius: 18px;
          border: 1px solid rgba(255, 255, 255, 0.2);
          min-width: 130px;
          box-shadow: 0 8px 16px rgba(0, 0, 0, 0.1);
        }

        .stat-value {
          display: block;
          font-size: 1.8rem;
          font-weight: 800;
          color: #ffffff;
          line-height: 1.1;
          margin-bottom: 2px;
        }

        .stat-label {
          font-size: 0.72rem;
          color: #e0e7ff;
          text-transform: uppercase;
          letter-spacing: 0.08em;
          font-weight: 700;
        }

        /* Toolbar Card */
        .toolbar {
          background: #ffffff;
          border-radius: 20px;
          padding: 24px 28px;
          margin-bottom: 28px;
          border: 1px solid #e2e8f0;
          box-shadow: 0 10px 30px rgba(0, 0, 0, 0.03);
        }

        .toolbar-section {
          margin-bottom: 24px;
          padding-bottom: 24px;
          border-bottom: 1px solid #f1f5f9;
        }

        .toolbar-section:last-child {
          margin-bottom: 0;
          padding-bottom: 0;
          border-bottom: none;
        }

        .section-title {
          font-size: 1.05rem;
          font-weight: 800;
          color: #1e1b4b;
          margin-bottom: 18px;
          display: flex;
          align-items: center;
          gap: 10px;
          letter-spacing: -0.01em;
        }

        .section-title::before {
          content: '';
          width: 5px;
          height: 18px;
          background: linear-gradient(135deg, #6366f1, #4338ca);
          border-radius: 4px;
        }

        /* Filter Section Grid */
        .filter-section {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
          gap: 20px;
          align-items: start;
        }

        .filter-group {
          display: flex;
          flex-direction: column;
          gap: 6px;
          min-width: 0;
        }

        .filter-label {
          font-size: 0.82rem;
          color: #475569;
          font-weight: 700;
          display: flex;
          align-items: center;
          gap: 6px;
          text-transform: uppercase;
          letter-spacing: 0.03em;
        }

        .filter-input, .filter-select, .date-input {
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

        .filter-input:focus, .filter-select:focus, .date-input:focus {
          outline: none;
          border-color: #6366f1;
          box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.15);
        }

        .date-range-group {
          display: flex;
          gap: 10px;
        }

        /* Multi-select dropdown styles */
        .multi-select-dropdown {
          position: relative;
          width: 100%;
        }

        .multi-select-trigger {
          display: flex;
          justify-content: space-between;
          align-items: center;
          padding: 10px 14px;
          border: 1.5px solid #cbd5e1;
          border-radius: 12px;
          background: #ffffff;
          cursor: pointer;
          transition: all 0.2s ease;
          width: 100%;
          box-sizing: border-box;
          font-weight: 600;
          font-size: 0.88rem;
        }

        .multi-select-trigger:hover {
          border-color: #94a3b8;
        }

        .multi-select-trigger.disabled {
          opacity: 0.6;
          cursor: not-allowed;
          background: #f8fafc;
        }

        .multi-select-values {
          flex: 1;
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: nowrap;
        }

        .multi-select-values .placeholder {
          color: #94a3b8;
        }

        .multi-select-values .selected-count {
          color: #1e1b4b;
          font-weight: 700;
        }

        .dropdown-arrow {
          color: #64748b;
          font-size: 11px;
          margin-left: 8px;
        }

        .multi-select-options {
          position: absolute;
          top: 100%;
          left: 0;
          right: 0;
          background: #ffffff;
          border: 1.5px solid #e2e8f0;
          border-radius: 14px;
          margin-top: 6px;
          max-height: 240px;
          overflow-y: auto;
          z-index: 100;
          box-shadow: 0 15px 30px rgba(0, 0, 0, 0.1);
        }

        .multi-select-actions {
          padding: 10px 14px;
          border-bottom: 1px solid #f1f5f9;
          position: sticky;
          top: 0;
          background: #ffffff;
        }

        .clear-all-btn {
          background: none;
          border: none;
          color: #ef4444;
          font-size: 0.8rem;
          font-weight: 700;
          cursor: pointer;
          padding: 2px 6px;
        }

        .clear-all-btn:hover {
          text-decoration: underline;
        }

        .multi-select-option {
          display: flex;
          align-items: center;
          gap: 10px;
          padding: 10px 14px;
          cursor: pointer;
          transition: background 0.15s;
          font-size: 0.85rem;
          font-weight: 500;
          color: #1e293b;
        }

        .multi-select-option:hover {
          background: #f1f5f9;
        }

        .multi-select-option input {
          accent-color: #6366f1;
        }

        /* Remarks Chips */
        .remarks-grid {
          display: flex;
          flex-wrap: wrap;
          gap: 10px;
        }

        .remark-chip {
          display: inline-flex;
          align-items: center;
          gap: 8px;
          padding: 8px 16px;
          background: #f8fafc;
          border: 1.5px solid #e2e8f0;
          border-radius: 12px;
          font-size: 0.82rem;
          font-weight: 700;
          color: #475569;
          cursor: pointer;
          transition: all 0.2s ease;
        }

        .remark-chip:hover {
          background: #e0e7ff;
          border-color: #c7d2fe;
          color: #3730a3;
        }

        .remark-chip.selected {
          background: #4338ca;
          border-color: #312e81;
          color: #ffffff;
          box-shadow: 0 4px 12px rgba(67, 56, 202, 0.25);
        }

        .remark-checkbox {
          width: 14px;
          height: 14px;
          border-radius: 4px;
          border: 1.5px solid #cbd5e1;
          display: flex;
          align-items: center;
          justify-content: center;
        }

        .remark-chip.selected .remark-checkbox {
          background: white;
          border-color: white;
        }

        .remark-chip.selected .remark-checkbox::after {
          content: '✓';
          color: #4338ca;
          font-size: 10px;
          font-weight: bold;
        }

        /* Buttons Section */
        .controls-section {
          display: flex;
          justify-content: space-between;
          align-items: center;
          gap: 16px;
          padding-top: 10px;
        }

        .page-size-control {
          display: flex;
          align-items: center;
          gap: 10px;
          font-size: 0.88rem;
          color: #475569;
          font-weight: 600;
        }

        .page-size-select {
          padding: 8px 12px;
          border: 1.5px solid #cbd5e1;
          border-radius: 10px;
          background: #ffffff;
          color: #0f172a;
          font-size: 0.88rem;
          font-weight: 700;
        }

        .button-group {
          display: flex;
          gap: 12px;
          flex-wrap: wrap;
        }

        .btn {
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

        .btn:hover:not(:disabled) {
          transform: translateY(-2px);
        }

        .btn-clear {
          background: #f1f5f9;
          color: #475569;
          border: 1px solid #cbd5e1;
        }

        .btn-clear:hover:not(:disabled) {
          background: #e2e8f0;
        }

        .btn-excel {
          background: linear-gradient(135deg, #10b981 0%, #059669 100%);
          color: #ffffff;
          box-shadow: 0 4px 12px rgba(16, 185, 129, 0.25);
        }

        .btn-excel:hover:not(:disabled) {
          box-shadow: 0 6px 16px rgba(16, 185, 129, 0.35);
        }

        .btn-pdf {
          background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%);
          color: #ffffff;
          box-shadow: 0 4px 12px rgba(239, 68, 68, 0.25);
        }

        .btn-pdf:hover:not(:disabled) {
          box-shadow: 0 6px 16px rgba(239, 68, 68, 0.35);
        }

        .btn-refresh {
          background: linear-gradient(135deg, #6366f1 0%, #4338ca 100%);
          color: #ffffff;
          box-shadow: 0 4px 12px rgba(99, 102, 241, 0.25);
        }

        .btn-refresh:hover:not(:disabled) {
          box-shadow: 0 6px 16px rgba(99, 102, 241, 0.35);
        }

        .btn-back {
          background: #0f172a;
          color: #ffffff;
          box-shadow: 0 4px 12px rgba(15, 23, 42, 0.2);
        }

        .btn-back:hover:not(:disabled) {
          background: #1e293b;
        }

        /* Table Card */
        .table-container {
          background: #ffffff;
          border-radius: 24px;
          border: 1px solid #e2e8f0;
          box-shadow: 0 10px 30px rgba(0, 0, 0, 0.03);
          overflow: hidden;
          margin-bottom: 24px;
        }

        .table-wrapper {
          overflow-x: auto;
          max-height: 70vh;
        }

        .data-table {
          width: 100%;
          border-collapse: separate;
          border-spacing: 0;
          font-size: 0.88rem;
        }

        .table-header {
          position: sticky;
          top: 0;
          z-index: 20;
          background: linear-gradient(135deg, #1e1b4b 0%, #312e81 100%);
        }

        .table-header-cell {
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

        .table-cell {
          padding: 12px 14px;
          border-bottom: 1px solid #f1f5f9;
          border-right: 1px solid #f8fafc;
          background: #ffffff;
          color: #1e293b;
          text-align: center;
          vertical-align: middle;
        }

        .table-row:nth-child(even) .table-cell {
          background: #f8fafc;
        }

        .table-row:hover .table-cell {
          background: #e0e7ff;
        }

        /* Badges & Accessories */
        .cutting-table-badge {
          display: inline-block;
          padding: 4px 12px;
          background: linear-gradient(135deg, #4338ca 0%, #312e81 100%);
          color: white;
          border-radius: 20px;
          font-size: 0.75rem;
          font-weight: 700;
          box-shadow: 0 2px 6px rgba(49, 46, 129, 0.2);
        }

        .cutting-table-multiple {
          background: linear-gradient(135deg, #ec4899 0%, #be185d 100%);
        }

        .table-image {
          width: 40px;
          height: 40px;
          object-fit: cover;
          border-radius: 10px;
          border: 1px solid #cbd5e1;
          cursor: pointer;
          transition: transform 0.2s ease, box-shadow 0.2s ease;
        }

        .table-image:hover {
          transform: scale(1.12);
          box-shadow: 0 6px 12px rgba(0, 0, 0, 0.15);
        }

        /* Image Modal Lightbox */
        .image-modal-backdrop {
          position: fixed;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          background-color: rgba(15, 23, 42, 0.75);
          z-index: 99999;
          display: flex;
          align-items: center;
          justify-content: center;
          backdrop-filter: blur(6px);
        }

        .image-modal-content {
          position: relative;
          max-width: 90%;
          max-height: 90%;
          background-color: white;
          border-radius: 20px;
          padding: 20px;
          box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);
        }

        .image-modal-img {
          max-width: 100%;
          max-height: 80vh;
          border-radius: 12px;
          object-fit: contain;
        }

        .image-modal-close {
          position: absolute;
          top: -14px;
          right: -14px;
          background-color: #0f172a;
          color: white;
          border: none;
          width: 32px;
          height: 32px;
          border-radius: 50%;
          font-size: 16px;
          font-weight: bold;
          display: flex;
          align-items: center;
          justify-content: center;
          cursor: pointer;
          box-shadow: 0 4px 10px rgba(0, 0, 0, 0.3);
        }

        /* Dialog Modal */
        .dialog-backdrop {
          position: fixed;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          background: rgba(15, 23, 42, 0.65);
          backdrop-filter: blur(8px);
          display: flex;
          align-items: center;
          justify-content: center;
          z-index: 9999;
          padding: 20px;
        }

        .dialog {
          background: #ffffff;
          border-radius: 24px;
          max-width: 440px;
          width: 100%;
          border: 1px solid rgba(99, 102, 241, 0.2);
          box-shadow: 0 25px 50px -12px rgba(15, 23, 42, 0.35);
          overflow: hidden;
          animation: dialogPop 0.25s cubic-bezier(0.16, 1, 0.3, 1);
        }

        @keyframes dialogPop {
          0% { transform: scale(0.92); opacity: 0; }
          100% { transform: scale(1); opacity: 1; }
        }

        .dialog-header {
          background: linear-gradient(135deg, #1e1b4b 0%, #312e81 100%);
          padding: 20px 24px;
          display: flex;
          align-items: center;
          justify-content: space-between;
          color: white;
        }

        .dialog-title {
          font-size: 1.1rem;
          font-weight: 800;
          margin: 0;
          display: flex;
          align-items: center;
          gap: 10px;
        }

        .dialog-close {
          background: rgba(255, 255, 255, 0.15);
          border: none;
          color: white;
          width: 32px;
          height: 32px;
          border-radius: 50%;
          font-size: 1.2rem;
          font-weight: bold;
          cursor: pointer;
          display: flex;
          align-items: center;
          justify-content: center;
          transition: all 0.15s ease;
        }

        .dialog-close:hover {
          background: rgba(255, 255, 255, 0.3);
          transform: scale(1.08);
        }

        .dialog-body {
          padding: 24px;
          max-height: 300px;
          overflow-y: auto;
        }

        .shades-list {
          display: flex;
          flex-direction: column;
          gap: 10px;
        }

        .shade-item {
          display: flex;
          align-items: center;
          gap: 10px;
          padding: 10px 14px;
          background: #f8fafc;
          border: 1px solid #e2e8f0;
          border-radius: 12px;
          font-size: 0.92rem;
          font-weight: 700;
          color: #1e293b;
        }

        .shade-bullet {
          color: #f59e0b;
          font-size: 1.2rem;
        }

        .no-shades {
          text-align: center;
          padding: 20px;
        }

        .no-shades-icon {
          font-size: 2.5rem;
          margin-bottom: 8px;
        }

        .no-shades-text {
          font-weight: 700;
          color: #166534;
        }

        .dialog-footer {
          padding: 16px 24px;
          background: #f8fafc;
          border-top: 1px solid #e2e8f0;
          display: flex;
          justify-content: flex-end;
        }

        .dialog-action-btn {
          padding: 10px 24px;
          background: #4338ca;
          color: white;
          border: none;
          border-radius: 12px;
          font-weight: 700;
          font-size: 0.9rem;
          cursor: pointer;
          transition: all 0.2s ease;
        }

        .dialog-action-btn:hover {
          background: #312e81;
          transform: translateY(-1px);
        }

        /* Pending Shades Button */
        .pending-btn {
          display: inline-flex;
          align-items: center;
          gap: 6px;
          padding: 6px 12px;
          background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%);
          color: #92400e;
          border: 1px solid #f59e0b;
          border-radius: 20px;
          font-size: 0.78rem;
          font-weight: 800;
          cursor: pointer;
          transition: all 0.2s ease;
          box-shadow: 0 2px 6px rgba(245, 158, 11, 0.2);
        }

        .pending-btn:hover {
          transform: translateY(-1px);
          box-shadow: 0 4px 10px rgba(245, 158, 11, 0.35);
          background: linear-gradient(135deg, #fde68a 0%, #fcd34d 100%);
        }

        .pending-icon {
          font-size: 0.85rem;
        }

        .no-pending {
          color: #94a3b8;
          font-size: 0.85rem;
        }

        /* Priority Badges */
        .priority-badge {
          display: inline-block;
          padding: 4px 10px;
          border-radius: 12px;
          font-size: 0.72rem;
          font-weight: 800;
          letter-spacing: 0.04em;
        }

        .priority-high {
          background: #fee2e2;
          color: #dc2626;
          border: 1px solid #fca5a5;
        }

        .priority-medium {
          background: #fef3c7;
          color: #d97706;
          border: 1px solid #fcd34d;
        }

        .priority-low {
          background: #e0e7ff;
          color: #4338ca;
          border: 1px solid #c7d2fe;
        }

        /* Remarks Badges */
        .remarks-badges {
          display: flex;
          flex-wrap: wrap;
          gap: 4px;
          justify-content: center;
        }

        .remark-badge {
          display: inline-block;
          padding: 3px 8px;
          border-radius: 10px;
          font-size: 0.72rem;
          font-weight: 700;
          white-space: nowrap;
        }

        .done-badge {
          background: #dcfce7;
          color: #15803d;
          border: 1px solid #86efac;
        }

        .pending-badge {
          background: #fef3c7;
          color: #b45309;
          border: 1px solid #fde68a;
        }

        .cancel-badge {
          background: #fee2e2;
          color: #dc2626;
          border: 1px solid #fca5a5;
          font-weight: 800;
        }

        .issue-badge {
          background: #fee2e2;
          color: #dc2626;
          border: 1px solid #fca5a5;
        }

        .default-badge {
          background: #f1f5f9;
          color: #475569;
          border: 1px solid #cbd5e1;
        }

        /* Pagination */
        .pagination {
          display: flex;
          justify-content: space-between;
          align-items: center;
          padding: 16px 24px;
          background: #ffffff;
          border-radius: 20px;
          border: 1px solid #e2e8f0;
          box-shadow: 0 10px 30px rgba(0, 0, 0, 0.03);
          margin-bottom: 24px;
        }

        .pagination-info {
          font-size: 0.85rem;
          color: #64748b;
          font-weight: 600;
        }

        .pagination-controls {
          display: flex;
          align-items: center;
          gap: 8px;
        }

        .pagination-btn, .page-btn {
          padding: 6px 12px;
          border: 1.5px solid #cbd5e1;
          border-radius: 10px;
          background: #ffffff;
          color: #475569;
          font-size: 0.82rem;
          font-weight: 700;
          cursor: pointer;
          transition: all 0.15s ease;
        }

        .pagination-btn:hover:not(:disabled), .page-btn:hover:not(:disabled) {
          border-color: #6366f1;
          color: #4338ca;
          background: #f5f3ff;
        }

        .pagination-btn:disabled {
          opacity: 0.5;
          cursor: not-allowed;
        }

        .page-numbers {
          display: flex;
          gap: 4px;
        }

        .page-btn.active-page {
          background: #4338ca;
          color: #ffffff;
          border-color: #312e81;
          box-shadow: 0 4px 10px rgba(67, 56, 202, 0.25);
        }

        /* Meta Info Bar */
        .meta-info {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-bottom: 16px;
        }

        .meta-left {
          display: flex;
          gap: 16px;
        }

        .last-updated {
          display: inline-flex;
          align-items: center;
          gap: 6px;
          font-size: 0.82rem;
          color: #64748b;
          font-weight: 600;
          background: #ffffff;
          padding: 6px 14px;
          border-radius: 10px;
          border: 1px solid #e2e8f0;
        }

        .refresh-indicator {
          font-size: 0.82rem;
          color: #6366f1;
          font-weight: 700;
        }

        .error-bar {
          background: #fee2e2;
          border: 1px solid #fca5a5;
          border-radius: 14px;
          padding: 14px 20px;
          margin-bottom: 20px;
          color: #991b1b;
        }

        .error-content {
          display: flex;
          align-items: center;
          gap: 12px;
          font-weight: 600;
          font-size: 0.9rem;
        }

        .retry-btn {
          margin-left: auto;
          padding: 6px 14px;
          background: #dc2626;
          color: white;
          border: none;
          border-radius: 8px;
          font-weight: 700;
          font-size: 0.82rem;
          cursor: pointer;
        }
      `}</style>

      {loading && (
        <div className="loader-overlay" role="status" aria-live="polite">
          <div className="loader-spinner" />
          <div className="loader-text">Loading job orders…</div>
        </div>
      )}

      <div className="main-content" aria-busy={loading || refreshing}>
        {/* Header */}
        <div className="header">
          <div className="header-content">
            <div className="header-left">
              <h1 className="header-title">Cutting Statistics Dashboard</h1>
              <p className="header-subtitle">Track and manage job order cutting progress</p>
            </div>
            <div className="header-right">
              <div className="stats-grid">
                <div className="stat-card">
                  <span className="stat-value">{rows.length}</span>
                  <span className="stat-label">Total Jobs</span>
                </div>
                <div className="stat-card">
                  <span className="stat-value">{filtered.length}</span>
                  <span className="stat-label">Filtered</span>
                </div>
                <div
                  className="stat-card"
                  onClick={() => setDaysFilter(daysFilter === 'red' ? 'all' : 'red')}
                  style={{
                    cursor: 'pointer',
                    background: daysFilter === 'red' ? 'rgba(239, 68, 68, 0.35)' : 'rgba(255, 255, 255, 0.12)',
                    borderColor: daysFilter === 'red' ? '#ef4444' : 'rgba(255, 255, 255, 0.2)'
                  }}
                  title="Click to toggle Red Zone filter (> 2 Days)"
                >
                  <span className="stat-value" style={{ color: '#fca5a5' }}>{redZoneCount}</span>
                  <span className="stat-label" style={{ color: '#fee2e2' }}>🔴 Red Zone (&gt;2 Days)</span>
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Toolbar with Filters */}
        <div className="toolbar">
          <div className="toolbar-section">
            <div className="section-title">Filters</div>
            <div className="filter-section">
              <div className="filter-group">
                <label className="filter-label">📅 Financial Year</label>
                <select
                  className="filter-select"
                  value={financialYearFilter}
                  onChange={(e) => setFinancialYearFilter(e.target.value)}
                  disabled={loading}
                >
                  <option value="ALL">🌐 All Data (Whole Data)</option>
                  {availableFYs.map((fy) => (
                    <option key={fy} value={fy}>
                      FY {fy} {fy === getCurrentFinancialYear() ? "(Current)" : ""}
                    </option>
                  ))}
                </select>
              </div>

              {/* Days / Zone Filter */}
              <div className="filter-group">
                <label className="filter-label">⏱️ Days / Zone Filter</label>
                <select
                  className="filter-select"
                  value={daysFilter}
                  onChange={(e) => setDaysFilter(e.target.value)}
                  disabled={loading}
                >
                  <option value="all">🌐 All Days</option>
                  <option value="red">🔴 Red Zone (&gt; 2 Days)</option>
                  <option value="green">🟢 Green Zone (≤ 2 Days)</option>
                  <option value="gt5">🔴 &gt; 5 Days</option>
                  <option value="gt10">🔴 &gt; 10 Days</option>
                </select>
                <div style={{ display: 'flex', gap: '6px', marginTop: '4px' }}>
                  <button
                    type="button"
                    onClick={() => setDaysFilter('red')}
                    style={{
                      flex: 1,
                      padding: '4px 8px',
                      borderRadius: '6px',
                      border: '1px solid #fca5a5',
                      background: daysFilter === 'red' ? '#fee2e2' : '#ffffff',
                      color: '#dc2626',
                      fontSize: '11px',
                      fontWeight: '700',
                      cursor: 'pointer'
                    }}
                  >
                    🔴 Red Zone ({redZoneCount})
                  </button>
                  <button
                    type="button"
                    onClick={() => setDaysFilter('green')}
                    style={{
                      flex: 1,
                      padding: '4px 8px',
                      borderRadius: '6px',
                      border: '1px solid #86efac',
                      background: daysFilter === 'green' ? '#dcfce7' : '#ffffff',
                      color: '#15803d',
                      fontSize: '11px',
                      fontWeight: '700',
                      cursor: 'pointer'
                    }}
                  >
                    🟢 Green Zone
                  </button>
                </div>
              </div>

              <div className="filter-group">
                <label className="filter-label">Lot No</label>
                <input
                  className="filter-input"
                  value={lotFilter}
                  onChange={(e) => setLotFilter(e.target.value)}
                  placeholder="Filter by Lot No..."
                  disabled={loading}
                />
              </div>

              <div className="filter-group">
                <label className="filter-label">Party Name</label>
                <select
                  className="filter-select"
                  value={partyFilter}
                  onChange={(e) => setPartyFilter(e.target.value)}
                  disabled={loading}
                >
                  <option value="">All Parties</option>
                  {distinctParties.map((party) => (
                    <option key={party} value={party}>
                      {party}
                    </option>
                  ))}
                </select>
              </div>

              {/* Multi-select Brand Filter */}
              <div className="filter-group">
                <label className="filter-label">Brand (Multi-select)</label>
                <MultiSelectDropdown
                  options={distinctBrands}
                  selectedValues={brandFilter}
                  onChange={setBrandFilter}
                  placeholder="Select brands..."
                  disabled={loading}
                />
              </div>

              {/* Multi-select Garment Type Filter */}
              <div className="filter-group">
                <label className="filter-label">Garment Type (Multi-select)</label>
                <MultiSelectDropdown
                  options={distinctGarments}
                  selectedValues={garmentFilter}
                  onChange={setGarmentFilter}
                  placeholder="Select garment types..."
                  disabled={loading}
                />
              </div>

              <div className="filter-group">
                <label className="filter-label">Date Range (from Date column)</label>
                <div className="date-range-group">
                  <input
                    type="date"
                    className="date-input"
                    value={startDate}
                    onChange={(e) => setStartDate(e.target.value)}
                    disabled={loading}
                    placeholder="Start Date"
                  />
                  <input
                    type="date"
                    className="date-input"
                    value={endDate}
                    onChange={(e) => setEndDate(e.target.value)}
                    disabled={loading}
                    placeholder="End Date"
                  />
                </div>
              </div>

              <div className="filter-group">
                <label className="filter-label">Section</label>
                <select
                  className="filter-select"
                  value={sectionFilter}
                  onChange={(e) => setSectionFilter(e.target.value)}
                  disabled={loading}
                >
                  <option value="">All Sections</option>
                  {distinctSections.map((section) => (
                    <option key={section} value={section}>
                      {section}
                    </option>
                  ))}
                </select>
              </div>

              <div className="filter-group">
                <label className="filter-label">Season</label>
                <select
                  className="filter-select"
                  value={seasonFilter}
                  onChange={(e) => setSeasonFilter(e.target.value)}
                  disabled={loading}
                >
                  <option value="">All Seasons</option>
                  {distinctSeasons.map((season) => (
                    <option key={season} value={season}>
                      {season}
                    </option>
                  ))}
                </select>
              </div>

              <div className="filter-group">
                <label className="filter-label">Priority</label>
                <select
                  className="filter-select"
                  value={priorityFilter}
                  onChange={(e) => setPriorityFilter(e.target.value)}
                  disabled={loading}
                >
                  <option value="">All Priorities</option>
                  {distinctPriorities.map((priority) => (
                    <option key={priority} value={priority}>
                      {priority}
                    </option>
                  ))}
                </select>
              </div>

              {/* Style Multi-select filter */}
              <div className="filter-group">
                <label className="filter-label">Style (Multi-select)</label>
                <MultiSelectDropdown
                  options={distinctStyles}
                  selectedValues={styleFilter}
                  onChange={setStyleFilter}
                  placeholder="Select styles..."
                  disabled={loading}
                />
              </div>

              {/* Fabric Multi-select filter */}
              <div className="filter-group">
                <label className="filter-label">Fabric (Multi-select)</label>
                <MultiSelectDropdown
                  options={distinctFabrics}
                  selectedValues={fabricFilter}
                  onChange={setFabricFilter}
                  placeholder="Select fabrics..."
                  disabled={loading}
                />
              </div>

              {/* Direct Stitching Multi-select filter */}
              <div className="filter-group">
                <label className="filter-label">Direct Stitching (Multi-select)</label>
                <MultiSelectDropdown
                  options={distinctDirectStitching}
                  selectedValues={directStitchingFilter}
                  onChange={setDirectStitchingFilter}
                  placeholder="Select..."
                  disabled={loading}
                />
              </div>
            </div>
          </div>

          <div className="toolbar-section">
            <div className="section-title">Remarks Filter</div>
            <div className="remarks-grid">
              {distinctRemarks.length === 0 ? (
                <span style={{ color: "#64748b", fontSize: "14px" }}>
                  No remarks available
                </span>
              ) : (
                distinctRemarks.map((label) => (
                  <div
                    key={label}
                    className={`remark-chip ${selectedRemarks.has(label) ? "selected" : ""}`}
                    onClick={() => toggleRemark(label)}
                  >
                    <div className="remark-checkbox"></div>
                    {label}
                  </div>
                ))
              )}
            </div>
          </div>

          <div className="toolbar-section">
            <div className="controls-section">
              <div className="left-controls">
                <div className="page-size-control">
                  <label className="control-label">
                    Show
                    <select
                      className="page-size-select"
                      value={pageSize}
                      onChange={(e) => setPageSize(Number(e.target.value))}
                      disabled={loading}
                    >
                      {[10, 25, 50, 100, 200].map((n) => (
                        <option key={n} value={n}>
                          {n}
                        </option>
                      ))}
                    </select>
                    entries
                  </label>
                </div>
              </div>

              <div className="button-group">
                <button
                  type="button"
                  className="btn btn-clear"
                  onClick={clearAllFilters}
                  disabled={loading}
                >
                  <span className="btn-icon">🗑️</span>
                  Clear Filters
                </button>
                <button
                  type="button"
                  className="btn btn-excel"
                  onClick={handleExportExcel}
                  disabled={loading || refreshing}
                >
                  <span className="btn-icon">📊</span>
                  Export Excel
                </button>
                <button
                  type="button"
                  className="btn btn-pdf"
                  onClick={handleExportPDF}
                  disabled={loading || refreshing}
                >
                  <span className="btn-icon">📄</span>
                  Export PDF
                </button>
                <button
                  type="button"
                  className="btn btn-refresh"
                  onClick={() => loadData("refresh")}
                  disabled={loading || refreshing}
                >
                  <span className="btn-icon">🔄</span>
                  {refreshing ? "Refreshing..." : "Refresh Data"}
                </button>
                <button type="button" className="btn btn-back" onClick={goBack}>
                  <span className="btn-icon">←</span>
                  Back
                </button>
              </div>
            </div>
          </div>
        </div>

        {err && (
          <div className="error-bar" role="alert">
            <div className="error-content">
              <span className="error-icon">⚠️</span>
              <span>{err}</span>
              <button className="retry-btn" onClick={() => loadData("refresh")} disabled={refreshing}>
                Try Again
              </button>
            </div>
          </div>
        )}

        {/* Meta Info */}
        <div className="meta-info">
          <div className="meta-left">
            {lastUpdated && (
              <div className="last-updated">
                <span className="meta-icon">🕒</span>
                Last updated: {lastUpdated}
              </div>
            )}
            {(startDate || endDate) && (
              <div className="last-updated">
                <span className="meta-icon">📅</span>
                Date Range: {startDate || "Any"} to {endDate || "Any"}
              </div>
            )}
          </div>
          <div className="meta-right">
            {refreshing && <div className="refresh-indicator">Updating data...</div>}
          </div>
        </div>

        {/* Table */}
        <div className="table-container">
          <div className="table-wrapper">
            <table className="data-table">
              <thead className="table-header">
                <tr>
                  {OUTPUT_COLS.map((h) => (
                    <th key={h} className="table-header-cell">
                      <div className="table-header-content">
                        {h}
                        {h === "Total Qty" && <span className="column-icon">🔢</span>}
                        {h === "Pending Shade" && <span className="column-icon">⏳</span>}
                        {h === "Remarks" && <span className="column-icon">🏷️</span>}
                        {h === "Priority" && <span className="column-icon">⚡</span>}
                      </div>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {!loading && filtered.length === 0 && (
                  <tr>
                    <td className="no-data" colSpan={OUTPUT_COLS.length}>
                      <div className="no-data-content">
                        <div className="no-data-icon">📭</div>
                        <div className="no-data-text">No records found matching your filters</div>
                        {(lotFilter ||
                          partyFilter ||
                          garmentFilter.length > 0 ||
                          seasonFilter ||
                          sectionFilter ||
                          brandFilter.length > 0 ||
                          priorityFilter ||
                          selectedRemarks.size > 0 ||
                          daysFilter !== "all" ||
                          startDate ||
                          endDate) && (
                            <button className="btn btn-clear" onClick={clearAllFilters}>
                              Clear all filters
                            </button>
                          )}
                      </div>
                    </td>
                  </tr>
                )}

                {!loading &&
                  filtered.length > 0 &&
                  pagedRows.map((row, i) => (
                    <tr key={startIdx + i} className="table-row">
                      {OUTPUT_COLS.map((h) => {
                        if (h === "Image") {
                          const imageUrl = row.imageUrl;
                          return (
                            <td key={h} className="table-cell">
                              {imageUrl ? (
                                <img
                                  src={imageUrl}
                                  alt="Style"
                                  className="table-image"
                                  referrerPolicy="no-referrer"
                                  onClick={() => setViewImageSrc(imageUrl)}
                                />
                              ) : (
                                <span className="no-image-placeholder">No Image</span>
                              )}
                            </td>
                          );
                        }

                        if (h === "Pending Shade") {
                          const lot = String(row["Lot No"] || "").trim();
                          const list = pendingListByLot[lot] || [];
                          const isPending = list.length > 0;

                          return (
                            <td key={h} className="table-cell">
                              {isPending ? (
                                <button
                                  type="button"
                                  className="pending-btn"
                                  onClick={() => openPendingDialog(lot)}
                                >
                                  <span className="pending-icon">⏳</span>
                                  Pending ({list.length})
                                </button>
                              ) : (
                                <span className="no-pending">—</span>
                              )}
                            </td>
                          );
                        }

                        if (h === "Priority") {
                          const priority = String(row[h] || "").trim().toLowerCase();
                          let priorityClass = "default-badge";
                          if (priority.includes("high") || priority.includes("urgent")) {
                            priorityClass = "priority-high";
                          } else if (priority.includes("medium")) {
                            priorityClass = "priority-medium";
                          } else if (priority.includes("low")) {
                            priorityClass = "priority-low";
                          }

                          return (
                            <td key={h} className="table-cell">
                              {priority ? (
                                <span className={`priority-badge ${priorityClass}`}>
                                  {priority.toUpperCase()}
                                </span>
                              ) : (
                                "—"
                              )}
                            </td>
                          );
                        }

                        if (h === "Total Qty") {
                          return (
                            <td key={h} className="table-cell numeric-cell">
                              {fmtNum(row[h])}
                            </td>
                          );
                        }

                        if (h === "Days after PO issue") {
                          const daysVal = parseFloat(row[h]);
                          if (isNaN(daysVal)) {
                            return (
                              <td key={h} className="table-cell numeric-cell">
                                —
                              </td>
                            );
                          }
                          const isRed = daysVal > 2;
                          return (
                            <td key={h} className="table-cell numeric-cell">
                              <span style={{
                                display: 'inline-flex',
                                alignItems: 'center',
                                gap: '4px',
                                padding: '3px 10px',
                                borderRadius: '12px',
                                fontSize: '0.82rem',
                                fontWeight: '800',
                                background: isRed ? '#fee2e2' : '#dcfce7',
                                color: isRed ? '#dc2626' : '#15803d',
                                border: isRed ? '1px solid #fca5a5' : '1px solid #86efac'
                              }}>
                                {isRed ? '🔴' : '🟢'} {daysVal} {daysVal === 1 ? 'Day' : 'Days'}
                              </span>
                            </td>
                          );
                        }
                        if (h === "Cutting Table") {
                          const cuttingTableValue = row[h] || "";
                          const isMultiple = cuttingTableValue && cuttingTableValue.includes(',');
                          return (
                            <td key={h} className="table-cell">
                              {cuttingTableValue ? (
                                <span className={`cutting-table-badge ${isMultiple ? 'cutting-table-multiple' : ''}`}>
                                  🪚 {cuttingTableValue}
                                </span>
                              ) : (
                                <span style={{ color: '#9ca3af', fontSize: '12px' }}>—</span>
                              )}
                            </td>
                          );
                        }

                        if (h === "Remarks") {
                          const remarks = String(row[h] || "");
                          return (
                            <td key={h} className="table-cell">
                              <div className="remarks-badges">
                                {splitRemarks(remarks).map((remark, idx) => (
                                  <span
                                    key={idx}
                                    className={`remark-badge ${
                                      remark.includes("Cancel")
                                        ? "cancel-badge"
                                        : remark.includes("Done")
                                          ? "done-badge"
                                          : remark.includes("Pending")
                                            ? "pending-badge"
                                            : remark.includes("Issue")
                                              ? "issue-badge"
                                              : "default-badge"
                                    }`}
                                  >
                                    {remark}
                                  </span>
                                ))}
                              </div>
                            </td>
                          );
                        }

                        return (
                          <td key={h} className="table-cell">
                            {row[h]}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Pagination */}
        <div className="pagination">
          <div className="pagination-info">
            Showing {startIdx + 1} to {Math.min(endIdx, filtered.length)} of {filtered.length}{" "}
            entries
            {filtered.length !== rows.length &&
              ` (filtered from ${rows.length} total entries)`}
          </div>
          <div className="pagination-controls">
            <button
              type="button"
              className="pagination-btn"
              onClick={() => setPage(1)}
              disabled={loading || currentPage <= 1}
            >
              « First
            </button>
            <button
              type="button"
              className="pagination-btn"
              onClick={() => setPage((p) => Math.max(1, p - 1))}
              disabled={loading || currentPage <= 1}
            >
              ‹ Previous
            </button>

            <div className="page-numbers">
              {Array.from({ length: Math.min(5, pageCount) }, (_, i) => {
                let pageNum;
                if (pageCount <= 5) {
                  pageNum = i + 1;
                } else if (currentPage <= 3) {
                  pageNum = i + 1;
                } else if (currentPage >= pageCount - 2) {
                  pageNum = pageCount - 4 + i;
                } else {
                  pageNum = currentPage - 2 + i;
                }

                return (
                  <button
                    key={pageNum}
                    type="button"
                    className={`page-btn ${currentPage === pageNum ? "active-page" : ""}`}
                    onClick={() => setPage(pageNum)}
                    disabled={loading}
                  >
                    {pageNum}
                  </button>
                );
              })}
            </div>

            <button
              type="button"
              className="pagination-btn"
              onClick={() => setPage((p) => Math.min(pageCount, p + 1))}
              disabled={loading || currentPage >= pageCount}
            >
              Next ›
            </button>
            <button
              type="button"
              className="pagination-btn"
              onClick={() => setPage(pageCount)}
              disabled={loading || currentPage >= pageCount}
            >
              Last »
            </button>
          </div>
        </div>

        {/* Dialog */}
        {dialogOpen && (
          <div className="dialog-backdrop" onClick={closeDialog}>
            <div className="dialog" onClick={(e) => e.stopPropagation()}>
              <div className="dialog-header">
                <h3 className="dialog-title">
                  <span className="dialog-icon">⏳</span>
                  Pending Shades — Lot {dialogLot}
                </h3>
                <button className="dialog-close" onClick={closeDialog} type="button">
                  ×
                </button>
              </div>
              <div className="dialog-body">
                {dialogShades && dialogShades.length > 0 ? (
                  <div className="shades-list">
                    {dialogShades.map((sh, idx) => (
                      <div key={idx} className="shade-item">
                        <span className="shade-bullet">•</span>
                        {sh}
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className="no-shades">
                    <div className="no-shades-icon">✅</div>
                    <div className="no-shades-text">No pending shades for this lot</div>
                  </div>
                )}
              </div>
              <div className="dialog-footer">
                <button className="dialog-action-btn" onClick={closeDialog} type="button">
                  Close
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Image Lightbox Modal */}
        {viewImageSrc && (
          <div className="image-modal-backdrop" onClick={() => setViewImageSrc(null)}>
            <div className="image-modal-content" onClick={(e) => e.stopPropagation()}>
              <button className="image-modal-close" onClick={() => setViewImageSrc(null)}>&times;</button>
              <img src={viewImageSrc} alt="Full Preview" className="image-modal-img" referrerPolicy="no-referrer" />
            </div>
          </div>
        )}
      </div>
    </div>
  );
}