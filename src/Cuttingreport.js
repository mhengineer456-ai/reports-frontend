import React, { useEffect, useMemo, useState, useRef } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { GOOGLE_API_KEY, SPREADSHEET_IDS, fetchSheetDataFromBackend } from "./config";
import { store } from "./store.js";
import { fetchRemarksForTab, saveRemarkForLot } from "./embPrintRemarksService";

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

// Cutting Scans Spreadsheet
const CUTTING_SCANS_SHEET_ID = SPREADSHEET_IDS.CUTTING_SCANS || "1UU9P1IOjuUYpm3Ojx06V5L-r6K2QRb4mpFQmlGblgDA";
const CUTTING_SCANS_RANGE = "CuttingScanss!A:C";

// Canonical output columns (order)
const OUTPUT_COLS = [
  "Lot No",
  "Garment Type",
  "Style",
  "Fabric",
  "Brand",
  "Total Qty",
  "Section",
  "Season",
  "Party Name",
  "Direct Stitching",
  "Job Order No",
  "Date",
  "Image",
  "Days after PO issue",
  "Pending Shade",
  "Cutting Date",
  "Cutting Scanned",
  "Priority",
  "Remarks",
  "💬 User Remarks",
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

function daysAfter(poDateStr, cuttingDateVal, isCuttingDone = false) {
  if (!poDateStr) return "";

  const parseDate = (val) => {
    if (!val) return null;
    if (val instanceof Date && !isNaN(val.getTime())) {
      return new Date(val.getFullYear(), val.getMonth(), val.getDate());
    }
    if (typeof val === "number" && !isNaN(val)) {
      if (val > 30000 && val < 70000) {
        const d = new Date(Math.round((val - 25569) * 86400 * 1000));
        if (!isNaN(d.getTime())) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
      }
      const d = new Date(val);
      if (!isNaN(d.getTime())) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
    }
    const str = String(val).trim();
    if (!str) return null;

    // DD MMM YYYY (e.g., "27 Aug 2025" or "27-Aug-2025")
    const mmmMatch = str.match(/^(\d{1,2})[\s\-]+([A-Za-z]{3,9})[\s\-]+(\d{4})/);
    if (mmmMatch) {
      const day = parseInt(mmmMatch[1], 10);
      const monthStr = mmmMatch[2].toLowerCase().slice(0, 3);
      const year = parseInt(mmmMatch[3], 10);
      const months = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
      const monthIdx = months.indexOf(monthStr);
      if (monthIdx !== -1 && !isNaN(day) && !isNaN(year)) {
        return new Date(year, monthIdx, day);
      }
    }

    // YYYY-MM-DD or YYYY/MM/DD
    const ymdMatch = str.match(/^(\d{4})[-\/\s](\d{1,2})[-\/\s](\d{1,2})/);
    if (ymdMatch) {
      const y = parseInt(ymdMatch[1], 10);
      const m = parseInt(ymdMatch[2], 10);
      const d = parseInt(ymdMatch[3], 10);
      if (y && m && d) return new Date(y, m - 1, d);
    }

    // DD-MM-YYYY or DD/MM/YYYY
    const dmyMatch = str.match(/^(\d{1,2})[-\/\s](\d{1,2})[-\/\s](\d{4})/);
    if (dmyMatch) {
      const d = parseInt(dmyMatch[1], 10);
      const m = parseInt(dmyMatch[2], 10);
      const y = parseInt(dmyMatch[3], 10);
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
  // If cutting is DONE for the lot and Cutting Date is available, calculate days from Cut Date - Job Date
  if (isCuttingDone && cuttingDateVal) {
    end = parseDate(cuttingDateVal);
  }

  // If cutting is NOT done (or cutting date is missing), dynamically calculate days up to Today
  if (!end) {
    const now = new Date();
    end = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  }

  const MS_PER_DAY = 24 * 60 * 60 * 1000;
  const diff = Math.floor((end.getTime() - start.getTime()) / MS_PER_DAY);
  return Number.isFinite(diff) ? String(diff) : "";
}

function formatScannedDate(ts) {
  if (!ts) return "";
  const s = String(ts).trim();
  if (!s) return "";

  const slashParts = s.split(/[\s,T]+/)[0].split(/[-/]/);
  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

  if (slashParts.length === 3) {
    let p1 = parseInt(slashParts[0], 10);
    let p2 = parseInt(slashParts[1], 10);
    let p3 = parseInt(slashParts[2], 10);

    // YYYY-MM-DD
    if (slashParts[0].length === 4) {
      if (p2 >= 1 && p2 <= 12 && p3 >= 1 && p3 <= 31) {
        return `${p3} ${monthNames[p2 - 1]} ${p1}`;
      }
    }
    // M/D/YYYY (e.g. 7/16/2026)
    if (slashParts[2].length === 4 && p1 <= 12 && p2 > 12) {
      return `${p2} ${monthNames[p1 - 1]} ${p3}`;
    }
    // D/M/YYYY (e.g. 16/7/2026)
    if (slashParts[2].length === 4 && p1 > 12 && p2 <= 12) {
      return `${p1} ${monthNames[p2 - 1]} ${p3}`;
    }
    // Standard M/D/YYYY
    if (slashParts[2].length === 4 && p1 >= 1 && p1 <= 12 && p2 >= 1 && p2 <= 31) {
      return `${p2} ${monthNames[p1 - 1]} ${p3}`;
    }
  }

  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return `${d.getDate()} ${monthNames[d.getMonth()]} ${d.getFullYear()}`;
  }

  return s;
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
    obj["Cutting Scanned"] = "";
    obj["Cutting Scanned Status"] = "";
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

  // Remarks States & Realtime Synchronization
  const [remarksMap, setRemarksMap] = useState({});
  const [remarksModalOpen, setRemarksModalOpen] = useState(false);
  const [selectedRemarksLot, setSelectedRemarksLot] = useState(null);
  const [newRemarkInputText, setNewRemarkInputText] = useState("");
  const [savingRemark, setSavingRemark] = useState(false);
  const [manualRemarksFilter, setManualRemarksFilter] = useState("all");

  // Fetch Cutting Remarks from Google Sheets & Subscribe to updates
  useEffect(() => {
    fetchRemarksForTab('CUTTING').then(map => {
      if (map && typeof map === 'object') {
        setRemarksMap(map);
      }
    });

    const handleRemarkUpdated = (e) => {
      if (e.detail && (e.detail.tabType === 'CUTTING' || !e.detail.tabType)) {
        setRemarksMap(prev => ({
          ...prev,
          [e.detail.lotNumber]: e.detail.history
        }));
      }
    };

    window.addEventListener('emb_print_remark_updated', handleRemarkUpdated);
    return () => {
      window.removeEventListener('emb_print_remark_updated', handleRemarkUpdated);
    };
  }, []);

  const handleOpenRemarksModal = (item) => {
    setSelectedRemarksLot(item);
    setNewRemarkInputText('');
    setRemarksModalOpen(true);
  };

  const handleCloseRemarksModal = () => {
    setRemarksModalOpen(false);
    setSelectedRemarksLot(null);
    setNewRemarkInputText('');
  };

  const handleSaveRemark = async () => {
    if (!selectedRemarksLot || !newRemarkInputText.trim()) return;

    try {
      setSavingRemark(true);
      const lotNumber = String(selectedRemarksLot["Lot No"] || selectedRemarksLot.lotNumber || '').trim();
      const updatedHistory = await saveRemarkForLot({
        tabType: 'CUTTING',
        lotNumber: lotNumber,
        partyName: selectedRemarksLot["Party Name"] || '',
        fabric: selectedRemarksLot["Fabric"] || '',
        style: selectedRemarksLot["Style"] || selectedRemarksLot["Garment Type"] || '',
        remarkText: newRemarkInputText.trim()
      });

      if (updatedHistory) {
        setRemarksMap(prev => ({
          ...prev,
          [lotNumber]: updatedHistory
        }));
      }

      setNewRemarkInputText('');
      setRemarksModalOpen(false);
    } catch (err) {
      console.error('Error saving remark:', err);
      alert('Failed to save remark. Please try again.');
    } finally {
      setSavingRemark(false);
    }
  };

  const distinctManualRemarks = useMemo(() => {
    const list = Object.values(remarksMap).flat().map(r => r.text).filter(Boolean);
    return Array.from(new Set(list));
  }, [remarksMap]);

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
        } else if (norm(t).includes("fabricissued") || norm(t) === "fabricissued") {
          set.add("Fabric Issued");
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
    setManualRemarksFilter("all");
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
      const [job, idxRes, cuttingRes, fabricRes, scansRes] = await Promise.all([
        fetchSheet({
          sheetId: JOB_SHEET_ID,
          range: JOB_RANGE,
          apiKey: API_KEY,
          signal: ctrl.signal,
        }),
        fetchSheet({
          sheetId: BUDGET_SHEET_ID,
          range: INDEX_RANGE,
          apiKey: API_KEY,
          signal: ctrl.signal,
        }),
        fetchSheet({
          sheetId: BUDGET_SHEET_ID,
          range: CUTTING_BIG_RANGE,
          apiKey: API_KEY,
          signal: ctrl.signal,
        }),
        store.getDailyFabricIssuanceReport("", "").catch((err) => {
          console.warn("Fabric issuance fetch error:", err);
          return null;
        }),
        fetchSheetDataFromBackend(CUTTING_SCANS_SHEET_ID, CUTTING_SCANS_RANGE)
          .then(res => (res && res.ok && res.values && res.values.length > 0) ? res : fetchSheetDataFromBackend(CUTTING_SCANS_SHEET_ID, "CuttingScans!A:C"))
          .catch(err => {
            console.warn("Cutting scans fetch error:", err);
            return { ok: false, values: [] };
          }),
      ]);

      let jobRows = convertValuesToObjects(job.values);

      const isCancel = (s) => {
        if (!s) return false;
        const sn = norm(s);
        return sn.startsWith("cancel") || sn.includes("cancel");
      };

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

      const bigCuttingValues = cuttingRes.values || [];

      // Parse fabric issuance records
      let fabricIssuances = [];
      if (Array.isArray(fabricRes)) {
        fabricIssuances = fabricRes;
      } else if (fabricRes && Array.isArray(fabricRes.data)) {
        fabricIssuances = fabricRes.data;
      } else if (fabricRes && fabricRes.success && Array.isArray(fabricRes.data)) {
        fabricIssuances = fabricRes.data;
      }

      const fabricIssuedMap = new Map();
      fabricIssuances.forEach((item) => {
        const rawLot = String(item.lotNumber || item.lotNo || item.lot || "").trim();
        if (!rawLot) return;
        const keyNorm = norm(rawLot);
        const keyLower = rawLot.toLowerCase();
        const rawDate = item.date || item.issueDate || item.createdAt || item.timestamp || "";
        let formattedDate = "";
        if (rawDate) {
          if (/^\d{4}-\d{2}-\d{2}/.test(String(rawDate))) {
            formattedDate = formatDateYMDToDDMMMYYYY(String(rawDate).slice(0, 10));
          } else {
            formattedDate = formatSavedAtToYMD(rawDate) || String(rawDate);
          }
        }
        const dataObj = {
          lotNumber: rawLot,
          issueDate: formattedDate,
          rawDate: rawDate,
        };
        if (!fabricIssuedMap.has(keyNorm)) fabricIssuedMap.set(keyNorm, dataObj);
        if (!fabricIssuedMap.has(keyLower)) fabricIssuedMap.set(keyLower, dataObj);
      });

      // Parse cutting scans records
      const cuttingScansMap = new Map();
      const scansValues = scansRes?.values || [];
      if (scansValues.length > 1) {
        const sHeader = (scansValues[0] || []).map(h => normalizeKey(String(h || "")));
        let tsIdx = sHeader.findIndex(h => h.includes("timestamp") || h.includes("time") || h.includes("date"));
        let lotIdx = sHeader.findIndex(h => h.includes("lot") || h.includes("lotno") || h.includes("lotnumber"));
        let statusIdx = sHeader.findIndex(h => h.includes("status"));

        if (tsIdx === -1) tsIdx = 0;
        if (lotIdx === -1) lotIdx = 1;
        if (statusIdx === -1) statusIdx = 2;

        for (let i = 1; i < scansValues.length; i++) {
          const sRow = scansValues[i] || [];
          const rawLot = String(sRow[lotIdx] || "").trim();
          if (!rawLot) continue;
          const rawTs = sRow[tsIdx] || "";
          const rawStatus = sRow[statusIdx] || "";

          let formattedDate = formatScannedDate(rawTs);

          const scanObj = {
            lotNumber: rawLot,
            timestamp: rawTs,
            scannedDate: formattedDate,
            status: rawStatus
          };

          cuttingScansMap.set(norm(rawLot), scanObj);
          cuttingScansMap.set(rawLot.toLowerCase(), scanObj);
          cuttingScansMap.set(rawLot, scanObj);
        }
      }

      const lots = Array.from(
        new Set(jobRows.map((r) => String(r["Lot No"] || "").trim()).filter(Boolean))
      );
      const lotToSummary = new Map();
      const pendingListTmp = {};

      for (const lot of lots) {
        const ix = indexMap.get(lot);
        if (!ix) {
          const fabricInfo = fabricIssuedMap.get(norm(lot)) || fabricIssuedMap.get(lot.toLowerCase().trim());
          const isFabricIssued = !!fabricInfo;
          const issueDate = fabricInfo?.issueDate || "";

          lotToSummary.set(lot, {
            totalQty: 0,
            remarks: "",
            remarks2: "",
            remarks3: lot ? (isFabricIssued ? "Fabric Issued" : "Fabric Issue Pending") : "",
            isCuttingDone: false,
            fabricIssueDate: isFabricIssued ? issueDate : "",
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

        const isCuttingDone = pendingShadeKeys.size === 0 && totalQty > 0;

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
          isCuttingDone,
          fabricIssueDate: "",
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
            isCuttingDone: false,
            fabricIssueDate: "",
            cuttingDate: "",
            cuttingTables: [],
            imageUrl: ""
          };
        const isDone = sum.isCuttingDone || sum.remarks === "Cutting Done";
        const days = daysAfter(r["PO Date"] || r["Date"], sum.cuttingDate, isDone);

        const scanInfo = cuttingScansMap.get(norm(lot)) || cuttingScansMap.get(lot.toLowerCase()) || cuttingScansMap.get(lot);
        const cuttingScanned = scanInfo ? scanInfo.scannedDate : "";
        const cuttingScannedStatus = scanInfo ? scanInfo.status : "";

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
          "Cutting Scanned": cuttingScanned || "",
          "Cutting Scanned Status": cuttingScannedStatus || "",
          "Cutting Table": cuttingTablesDisplay,
          Remarks: mergedRemarks,
          fabricIssueDate: sum.fabricIssueDate || "",
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
        const hasFabricIssued = tokens.some(t => norm(t).includes("fabricissued") || norm(t) === "fabricissued");

        let matchesAny = false;
        for (const sel of selected) {
          if (sel === "Cancel" && hasCancel) {
            matchesAny = true;
            break;
          }
          if (sel === "Fabric Issued" && hasFabricIssued) {
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

      // Manual Remarks filter
      if (manualRemarksFilter && manualRemarksFilter !== "all") {
        const lot = String(r["Lot No"] || '').trim();
        const lotRemarks = remarksMap[lot] || [];
        const hasRemark = lotRemarks.length > 0 && Boolean(lotRemarks[lotRemarks.length - 1].text);
        const latestText = hasRemark ? lotRemarks[lotRemarks.length - 1].text : '';

        if (manualRemarksFilter === 'WITH_REMARKS') {
          if (!hasRemark) return false;
        } else if (manualRemarksFilter === 'WITHOUT_REMARKS') {
          if (hasRemark) return false;
        } else {
          const match = latestText === manualRemarksFilter || lotRemarks.some(rem => rem.text === manualRemarksFilter);
          if (!match) return false;
        }
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
    manualRemarksFilter,
    remarksMap
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
      const lotRemarks = remarksMap[lot] || [];
      const latestRemark = lotRemarks.length > 0 ? lotRemarks[lotRemarks.length - 1] : null;
      const obj = {};
      OUTPUT_COLS.forEach((col) => {
        if (col === "Pending Shade") obj[col] = pending;
        else if (col === "Image") obj[col] = r.imageUrl || "";
        else if (col === "💬 User Remarks") obj[col] = latestRemark ? latestRemark.text : "—";
        else obj[col] = r[col] ?? "";
      });
      return obj;
    });
  };

  const handleExportExcel = async () => {
    try {
      setLoading(true);
      const exportRows = buildExportRows(filtered);
      if (exportRows.length === 0) {
        alert("No data available to export.");
        return;
      }

      const now = new Date();
      const reportDateStr = `${String(now.getDate()).padStart(2, '0')}/${String(now.getMonth() + 1).padStart(2, '0')}/${now.getFullYear()} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;

      const totalLots = exportRows.length;
      let totalQty = 0;
      let mohitCount = 0;
      let highPriorityCount = 0;
      let redZoneLots = 0;
      let redZoneQty = 0;
      let normalLots = 0;
      let normalQty = 0;
      const parties = new Set();
      const garmentAnalysis = {};
      const partyAnalysis = {};
      const sectionAnalysis = {};

      exportRows.forEach(row => {
        const qty = parseInt(row["Total Qty"], 10) || 0;
        totalQty += qty;

        const pn = (row["Party Name"] || "").toString().trim();
        if (pn) {
          parties.add(pn);
          if (pn.toLowerCase().includes("mohit")) mohitCount++;
          partyAnalysis[pn] = (partyAnalysis[pn] || { lots: 0, qty: 0 });
          partyAnalysis[pn].lots += 1;
          partyAnalysis[pn].qty += qty;
        }

        const isRepeated = (row["Priority"] || "").toString().toUpperCase().includes("REPEATED_LOT");
        if (isRepeated) highPriorityCount++;

        const days = parseFloat(row["Days after PO issue"]);
        if (!isNaN(days) && days > 2) {
          redZoneLots++;
          redZoneQty += qty;
        } else {
          normalLots++;
          normalQty += qty;
        }

        const garment = (row["Garment Type"] || "Unassigned").toString().trim() || "Unassigned";
        garmentAnalysis[garment] = (garmentAnalysis[garment] || { lots: 0, qty: 0 });
        garmentAnalysis[garment].lots += 1;
        garmentAnalysis[garment].qty += qty;

        const section = (row["Section"] || "Unassigned").toString().trim() || "Unassigned";
        sectionAnalysis[section] = (sectionAnalysis[section] || { lots: 0, qty: 0 });
        sectionAnalysis[section].lots += 1;
        sectionAnalysis[section].qty += qty;
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
      // SHEET 1: CUTTING PRODUCTION STATS
      // ==========================================
      const ws1 = workbook.addWorksheet("Cutting Stats", {
        views: [{ showGridLines: true, state: 'frozen', xSplit: 0, ySplit: 6 }]
      });

      const cols1 = [
        { header: "Sr.No", key: "sr", width: 8 },
        { header: "Lot No", key: "lotNo", width: 16 },
        { header: "Garment Type", key: "garmentType", width: 18 },
        { header: "Style", key: "style", width: 16 },
        { header: "Fabric", key: "fabric", width: 22 },
        { header: "Brand", key: "brand", width: 16 },
        { header: "Total Qty", key: "totalQty", width: 14 },
        { header: "Section", key: "section", width: 14 },
        { header: "Season", key: "season", width: 12 },
        { header: "Party Name", key: "partyName", width: 20 },
        { header: "Direct Stitching", key: "directStitching", width: 16 },
        { header: "Job Order No", key: "jobOrderNo", width: 16 },
        { header: "Date", key: "date", width: 14 },
        { header: "Days after PO issue", key: "daysAfterPO", width: 18 },
        { header: "Pending Shade", key: "pendingShade", width: 24 },
        { header: "Cutting Date", key: "cuttingDate", width: 16 },
        { header: "Cutting Scanned", key: "cuttingScanned", width: 18 },
        { header: "Priority", key: "priority", width: 16 },
        { header: "Remarks", key: "remarks", width: 32 },
        { header: "Manual Remarks", key: "userRemarks", width: 28 }
      ];

      const numCols1 = cols1.length;
      ws1.columns = cols1;

      // Row 1: Title Banner
      const titleRow1 = ws1.getRow(1);
      titleRow1.values = ["FACTORY SUITE PRO - CUTTING PRODUCTION STATS REPORT"];
      ws1.mergeCells(1, 1, 1, numCols1);
      titleRow1.font = { name: "Segoe UI", size: 14, bold: true, color: { argb: "FFFFFFFF" } };
      titleRow1.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F4C81" } };
      titleRow1.alignment = { vertical: "middle", horizontal: "center" };
      titleRow1.height = 34;

      // Row 2: Metadata Banner
      const metaRow1 = ws1.getRow(2);
      metaRow1.values = [`Exported on: ${reportDateStr}  |  Total Lots: ${totalLots}  |  Total Quantity: ${totalQty.toLocaleString()}  |  Unique Parties: ${parties.size}  |  MH Lots: ${mohitCount}  |  Repeated Lots: ${highPriorityCount}`];
      ws1.mergeCells(2, 1, 2, numCols1);
      metaRow1.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF1E293B" } };
      metaRow1.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } };
      metaRow1.alignment = { vertical: "middle", horizontal: "center" };
      metaRow1.height = 22;

      // Row 3: Blank Row
      const blankRow3 = ws1.getRow(3);
      blankRow3.values = [];
      blankRow3.height = 6;

      // Row 4: KPI Summary Row
      const kpiRow1 = ws1.getRow(4);
      kpiRow1.values = [`KPI SUMMARY  |  TOTAL LOTS: ${totalLots}  |  TOTAL QTY: ${totalQty.toLocaleString()}  |  NORMAL (<=2d): ${normalLots} (${totalLots > 0 ? Math.round((normalLots / totalLots) * 100) : 0}%)  |  RED ZONE (>2d): ${redZoneLots} (${totalLots > 0 ? Math.round((redZoneLots / totalLots) * 100) : 0}%)  |  REPEATED LOTS: ${highPriorityCount}`];
      ws1.mergeCells(4, 1, 4, numCols1);
      kpiRow1.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FF0369A1" } };
      kpiRow1.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE0F2FE" } };
      kpiRow1.alignment = { vertical: "middle", horizontal: "center" };
      kpiRow1.height = 24;

      // Row 5: Blank Row
      const blankRow5 = ws1.getRow(5);
      blankRow5.values = [];
      blankRow5.height = 6;

      // Row 6: Table Headers Row
      const headerRow1 = ws1.getRow(6);
      headerRow1.values = cols1.map(c => c.header);
      headerRow1.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
      headerRow1.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F4C81" } };
      headerRow1.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      headerRow1.height = 30;
      for (let c = 1; c <= numCols1; c++) {
        headerRow1.getCell(c).border = headerBorder;
      }

      // Add Data Rows
      let rIdx1 = 7;
      exportRows.forEach((row, index) => {
        const lotNo = (row["Lot No"] || "").toString().trim();
        const partyRaw = (row["Party Name"] || "").toString().trim();
        const partyDisplay = partyRaw.toLowerCase().includes("mohit") ? "MH" : partyRaw;
        const totalQtyVal = parseInt(row["Total Qty"], 10) || 0;
        const days = parseFloat(row["Days after PO issue"]);
        const isRepeated = (row["Priority"] || "").toString().toUpperCase().includes("REPEATED_LOT");
        const lotRemarks = remarksMap[lotNo] || [];
        const latestRemark = lotRemarks.length > 0 ? lotRemarks[lotRemarks.length - 1] : null;

        const rowValues = [
          index + 1,
          lotNo || "—",
          row["Garment Type"] || "—",
          row["Style"] || "—",
          row["Fabric"] || "—",
          row["Brand"] || "—",
          totalQtyVal,
          row["Section"] || "—",
          row["Season"] || "—",
          partyDisplay || "—",
          row["Direct Stitching"] ? (String(row["Direct Stitching"]).toLowerCase() === "yes" ? "Yes" : "No") : "No",
          row["Job Order No"] || "—",
          row["Date"] || "—",
          !isNaN(days) ? days : "—",
          row["Pending Shade"] || "—",
          row["Cutting Date"] || "—",
          row["Cutting Scanned"] || "—",
          row["Priority"] || "—",
          row["Remarks"] || "—",
          latestRemark ? latestRemark.text : (row["💬 User Remarks"] || "—")
        ];

        const dataRow = ws1.getRow(rIdx1);
        dataRow.values = rowValues;
        dataRow.height = 24;

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
          // Total Qty format
          if (c === 7) {
            cell.numFmt = "#,##0";
            cell.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF0F172A" } };
          }
          // Days after PO issue highlight
          if (c === 14 && !isNaN(days)) {
            cell.font = { name: "Segoe UI", size: 10, bold: true };
            if (days > 2) {
              cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEE2E2" } };
              cell.font.color = { argb: "FF991B1B" };
            } else {
              cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDCFCE7" } };
              cell.font.color = { argb: "FF166534" };
            }
          }
          // Priority highlight
          if (c === 18 && isRepeated) {
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEF3C7" } };
            cell.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF92400E" } };
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
      // SHEET 2: CUTTING BREAKDOWN & SUMMARY
      // ==========================================
      const ws2 = workbook.addWorksheet("Cutting Summary", {
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
      titleRow2.values = ["FACTORY SUITE PRO - CUTTING EXECUTIVE SUMMARY & BREAKDOWN"];
      ws2.mergeCells(1, 1, 1, 4);
      titleRow2.font = { name: "Segoe UI", size: 14, bold: true, color: { argb: "FFFFFFFF" } };
      titleRow2.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E293B" } };
      titleRow2.alignment = { vertical: "middle", horizontal: "center" };
      titleRow2.height = 32;

      const subRow2 = ws2.getRow(2);
      subRow2.values = [`Report Date: ${reportDateStr}  |  Total Lots: ${totalLots}  |  Total Quantity: ${totalQty.toLocaleString()}`];
      ws2.mergeCells(2, 1, 2, 4);
      subRow2.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF1E293B" } };
      subRow2.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } };
      subRow2.alignment = { vertical: "middle", horizontal: "center" };
      subRow2.height = 20;

      let r2 = 4;

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

      const totalGRow = ws2.getRow(r2);
      totalGRow.values = ["TOTAL", totalLots, "100%", totalQty];
      totalGRow.height = 24;
      for (let c = 1; c <= 4; c++) {
        const cell = totalGRow.getCell(c);
        cell.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF115E59" } };
        cell.alignment = { vertical: "middle", horizontal: c === 1 ? "left" : "center", indent: c === 1 ? 1 : 0 };
        cell.border = totalBorder;
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFCCFBF1" } };
        if (c === 4) cell.numFmt = "#,##0";
      }
      r2 += 2;

      // Section 2: Days Aging Breakdown
      const aTitle = ws2.getRow(r2);
      aTitle.values = ["2. DAYS AFTER PO ISSUE (AGING & RED ZONE ANALYSIS)"];
      ws2.mergeCells(r2, 1, r2, 4);
      aTitle.font = { name: "Segoe UI", size: 10.5, bold: true, color: { argb: "FFFFFFFF" } };
      aTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFB45309" } };
      aTitle.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
      aTitle.height = 24;
      r2++;

      const aHeaders = ws2.getRow(r2);
      aHeaders.values = ["Aging Category", "Total Lots", "Percentage (%)", "Total Quantity"];
      aHeaders.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFFFFFFF" } };
      aHeaders.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF59E0B" } };
      aHeaders.alignment = { vertical: "middle", horizontal: "center" };
      aHeaders.height = 24;
      for (let c = 1; c <= 4; c++) aHeaders.getCell(c).border = headerBorder;
      r2++;

      const agingData = [
        { name: "Normal (<= 2 Days)", lots: normalLots, pct: totalLots > 0 ? Math.round((normalLots / totalLots) * 100) : 0, qty: normalQty, bg: "FFDCFCE7", fg: "FF166534" },
        { name: "Red Zone (> 2 Days)", lots: redZoneLots, pct: totalLots > 0 ? Math.round((redZoneLots / totalLots) * 100) : 0, qty: redZoneQty, bg: "FFFEE2E2", fg: "FF991B1B" }
      ];

      agingData.forEach(a => {
        const row = ws2.getRow(r2);
        row.values = [a.name, a.lots, `${a.pct}%`, a.qty];
        row.height = 22;
        for (let c = 1; c <= 4; c++) {
          const cell = row.getCell(c);
          cell.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: a.fg } };
          cell.alignment = { vertical: "middle", horizontal: c === 1 ? "left" : "center", indent: c === 1 ? 1 : 0 };
          cell.border = thinBorder;
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: a.bg } };
          if (c === 4) cell.numFmt = "#,##0";
        }
        r2++;
      });

      const totalARow = ws2.getRow(r2);
      totalARow.values = ["TOTAL", totalLots, "100%", totalQty];
      totalARow.height = 24;
      for (let c = 1; c <= 4; c++) {
        const cell = totalARow.getCell(c);
        cell.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF78350F" } };
        cell.alignment = { vertical: "middle", horizontal: c === 1 ? "left" : "center", indent: c === 1 ? 1 : 0 };
        cell.border = totalBorder;
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEF3C7" } };
        if (c === 4) cell.numFmt = "#,##0";
      }
      r2 += 2;

      // Section 3: Party Name Breakdown
      const pTitle = ws2.getRow(r2);
      pTitle.values = ["3. PARTY-WISE PRODUCTION SUMMARY"];
      ws2.mergeCells(r2, 1, r2, 4);
      pTitle.font = { name: "Segoe UI", size: 10.5, bold: true, color: { argb: "FFFFFFFF" } };
      pTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E40AF" } };
      pTitle.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
      pTitle.height = 24;
      r2++;

      const pHeaders = ws2.getRow(r2);
      pHeaders.values = ["Party Name", "Total Lots", "Percentage (%)", "Total Quantity"];
      pHeaders.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFFFFFFF" } };
      pHeaders.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF3B82F6" } };
      pHeaders.alignment = { vertical: "middle", horizontal: "center" };
      pHeaders.height = 24;
      for (let c = 1; c <= 4; c++) pHeaders.getCell(c).border = headerBorder;
      r2++;

      const partyArr = Object.entries(partyAnalysis).map(([p, d]) => ({
        name: p,
        lots: d.lots,
        pct: totalLots > 0 ? Math.round((d.lots / totalLots) * 100) : 0,
        qty: d.qty
      })).sort((a, b) => b.lots - a.lots);

      partyArr.forEach((p, idx) => {
        const row = ws2.getRow(r2);
        row.values = [p.name, p.lots, `${p.pct}%`, p.qty];
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

      const totalPRow = ws2.getRow(r2);
      totalPRow.values = ["TOTAL", totalLots, "100%", totalQty];
      totalPRow.height = 24;
      for (let c = 1; c <= 4; c++) {
        const cell = totalPRow.getCell(c);
        cell.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF1E3A8A" } };
        cell.alignment = { vertical: "middle", horizontal: c === 1 ? "left" : "center", indent: c === 1 ? 1 : 0 };
        cell.border = totalBorder;
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDBEAFE" } };
        if (c === 4) cell.numFmt = "#,##0";
      }

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
        ["Financial Year", financialYearFilter || "All"],
        ["Date Range", startDate || endDate ? `${startDate || ""} to ${endDate || ""}` : "All"],
        ["Lot Filter", lotFilter || "All"],
        ["Party Filter", Array.isArray(partyFilter) && partyFilter.length > 0 ? partyFilter.join(", ") : (partyFilter || "All")],
        ["Garment Filter", Array.isArray(garmentFilter) && garmentFilter.length > 0 ? garmentFilter.join(", ") : (garmentFilter || "All")],
        ["Season Filter", Array.isArray(seasonFilter) && seasonFilter.length > 0 ? seasonFilter.join(", ") : (seasonFilter || "All")],
        ["Section Filter", Array.isArray(sectionFilter) && sectionFilter.length > 0 ? sectionFilter.join(", ") : (sectionFilter || "All")],
        ["Brand Filter", Array.isArray(brandFilter) && brandFilter.length > 0 ? brandFilter.join(", ") : (brandFilter || "All")],
        ["Style Filter", Array.isArray(styleFilter) && styleFilter.length > 0 ? styleFilter.join(", ") : (styleFilter || "All")],
        ["Fabric Filter", Array.isArray(fabricFilter) && fabricFilter.length > 0 ? fabricFilter.join(", ") : (fabricFilter || "All")],
        ["Days Filter", daysFilter || "All"],
        ["Priority Filter", priorityFilter || "All"],
        ["Manual Remarks Filter", manualRemarksFilter || "All"]
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

      const fileName = `Cutting_Stats_${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      saveAs(blob, `${fileName}.xlsx`);
    } catch (err) {
      console.error("Error exporting Excel:", err);
      alert(`Failed to export Excel: ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  const handleExportPDF = () => {
    try {
      const exportRows = buildExportRows(filtered);
      if (!exportRows || exportRows.length === 0) {
        alert("No data available to export.");
        return;
      }

      const totalLots = exportRows.length;
      let totalQty = 0;
      let normalLots = 0;
      let redZoneLots = 0;
      let highPriorityCount = 0;

      const garmentMap = {};
      const partyMap = {};
      const sectionMap = {};

      exportRows.forEach(row => {
        const qty = parseInt(row["Total Qty"], 10) || 0;
        totalQty += qty;

        const days = parseFloat(row["Days after PO issue"]);
        if (!isNaN(days)) {
          if (days > 2) redZoneLots++;
          else normalLots++;
        }

        const isRepeated = (row["Priority"] || "").toString().toUpperCase().includes("REPEATED_LOT");
        if (isRepeated) highPriorityCount++;

        // Garment grouping
        const gName = (row["Garment Type"] || "Unspecified").toString().trim() || "Unspecified";
        if (!garmentMap[gName]) garmentMap[gName] = { lots: 0, qty: 0 };
        garmentMap[gName].lots += 1;
        garmentMap[gName].qty += qty;

        // Party grouping
        let pName = (row["Party Name"] || "Unassigned").toString().trim() || "Unassigned";
        if (pName.toLowerCase().includes("mohit")) pName = "MH (Mohit Hosiery)";
        if (!partyMap[pName]) partyMap[pName] = { lots: 0, qty: 0 };
        partyMap[pName].lots += 1;
        partyMap[pName].qty += qty;

        // Section grouping
        const sName = (row["Section"] || "Unassigned").toString().trim() || "Unassigned";
        if (!sectionMap[sName]) sectionMap[sName] = { lots: 0, qty: 0 };
        sectionMap[sName].lots += 1;
        sectionMap[sName].qty += qty;
      });

      const sortedGarments = Object.keys(garmentMap).map(name => ({
        name,
        lots: garmentMap[name].lots,
        qty: garmentMap[name].qty
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
      doc.setFillColor(15, 23, 42); // Dark Navy #0F172A
      doc.rect(15, 12, pageW - 30, 48, 'F');

      doc.setFontSize(14);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(255, 255, 255);
      doc.text("FACTORY SUITE PRO - CUTTING DEPARTMENT REPORT", pageW / 2, 30, { align: 'center' });

      doc.setFontSize(8.5);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(199, 210, 254);
      const subText = `Total Lots: ${totalLots}   |   Total Qty: ${totalQty.toLocaleString()} Pcs   |   Normal (<=2d): ${normalLots}   |   Red Zone (>2d): ${redZoneLots}   |   Repeated Lots: ${highPriorityCount}   |   Parties: ${sortedParties.length}`;
      doc.text(subText, pageW / 2, 48, { align: 'center' });

      // 2. Filter Banner
      doc.setFillColor(241, 245, 249);
      doc.rect(15, 63, pageW - 30, 16, 'F');
      doc.setFontSize(8);
      doc.setFont('helvetica', 'italic');
      doc.setTextColor(0, 0, 0);
      const filterSummary = `Filters: Financial Year: ${financialYearFilter || 'All'} | Garments: ${Array.isArray(garmentFilter) && garmentFilter.length ? garmentFilter.join(', ') : (garmentFilter || 'All')} | Fabrics: ${Array.isArray(fabricFilter) && fabricFilter.length ? fabricFilter.join(', ') : (fabricFilter || 'All')} | Brands: ${Array.isArray(brandFilter) && brandFilter.length ? brandFilter.join(', ') : (brandFilter || 'All')} | Parties: ${Array.isArray(partyFilter) && partyFilter.length ? partyFilter.join(', ') : (partyFilter || 'All')} | Sections: ${Array.isArray(sectionFilter) && sectionFilter.length ? sectionFilter.join(', ') : (sectionFilter || 'All')} | Seasons: ${Array.isArray(seasonFilter) && seasonFilter.length ? seasonFilter.join(', ') : (seasonFilter || 'All')} | Days: ${daysFilter || 'All'} | Search: ${lotFilter || 'None'}`;
      doc.text(filterSummary, pageW / 2, 74, { align: 'center' });

      // 3. Main Data Table
      const tableColumns = [
        '#',
        'Lot No',
        'Garment Type',
        'Style',
        'Fabric',
        'Brand',
        'Total Qty',
        'Section',
        'Season',
        'Party Name',
        'Direct Stitching',
        'Job Order No',
        'PO Date',
        'Days',
        'Pending Shade',
        'Cutting Date',
        'Cutting Scanned',
        'Remarks',
        'User Remarks'
      ];

      const cleanRemarksText = (rawRemarks) => {
        if (!rawRemarks) return '—';
        return String(rawRemarks)
          .split('|')
          .map(p => p.trim())
          .filter(p => p && !p.includes('undefined'))
          .join(', ') || '—';
      };

      const tableBody = exportRows.map((row, idx) => {
        const lotNo = (row["Lot No"] || "").toString().trim();
        const partyRaw = (row["Party Name"] || "").toString().trim();
        const partyDisplay = partyRaw.toLowerCase().includes("mohit") ? "MH" : (partyRaw || "—");
        const totalQtyVal = parseInt(row["Total Qty"], 10) || 0;
        const days = parseFloat(row["Days after PO issue"]);
        const pending = (pendingListByLot[lotNo] || []).join(", ") || (row["Pending Shade"] || "—");
        const lotRemarks = remarksMap[lotNo] || [];
        const latestRemark = lotRemarks.length > 0 ? lotRemarks[lotRemarks.length - 1] : null;
        const userRem = latestRemark ? latestRemark.text : (row["💬 User Remarks"] || "—");
        const isRepeated = (row["Priority"] || "").toString().toUpperCase().includes("REPEATED_LOT");
        const lotDisplay = isRepeated ? `* ${lotNo}` : (lotNo || "—");

        return [
          (idx + 1).toString(),
          lotDisplay,
          row["Garment Type"] || "—",
          row["Style"] || "—",
          row["Fabric"] || "—",
          row["Brand"] || "—",
          totalQtyVal.toLocaleString(),
          row["Section"] || "—",
          row["Season"] || "—",
          partyDisplay,
          row["Direct Stitching"] ? (String(row["Direct Stitching"]).toLowerCase() === "yes" ? "Yes" : "No") : "No",
          row["Job Order No"] || "—",
          row["Date"] || "—",
          !isNaN(days) ? days.toString() : "—",
          pending,
          row["Cutting Date"] || "—",
          row["Cutting Scanned"] || "—",
          cleanRemarksText(row["Remarks"]),
          userRem
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
        '',
        '',
        `${redZoneLots} Red | ${normalLots} Norm`,
        '',
        '',
        '',
        '',
        ''
      ]);

      const columnStyles = {
        0: { cellWidth: 22, halign: 'center' },
        1: { cellWidth: 60, halign: 'center', fontStyle: 'bold' },
        2: { cellWidth: 70, halign: 'center' },
        3: { cellWidth: 75, halign: 'center' },
        4: { cellWidth: 75, halign: 'center' },
        5: { cellWidth: 50, halign: 'center' },
        6: { cellWidth: 50, halign: 'center', fontStyle: 'bold' },
        7: { cellWidth: 45, halign: 'center' },
        8: { cellWidth: 45, halign: 'center' },
        9: { cellWidth: 55, halign: 'center' },
        10: { cellWidth: 45, halign: 'center' },
        11: { cellWidth: 55, halign: 'center' },
        12: { cellWidth: 55, halign: 'center' },
        13: { cellWidth: 38, halign: 'center' },
        14: { cellWidth: 80, halign: 'left' },
        15: { cellWidth: 60, halign: 'center' },
        16: { cellWidth: 65, halign: 'center' },
        17: { cellWidth: 100, halign: 'left' },
        18: { cellWidth: 100, halign: 'left' }
      };

      autoTable(doc, {
        head: [tableColumns],
        body: tableBody,
        startY: 85,
        tableWidth: pageW - 30,
        margin: { top: 85, right: 15, bottom: 25, left: 15 },
        theme: "grid",
        styles: {
          fontSize: 8,
          cellPadding: { top: 3.5, right: 2, bottom: 3.5, left: 2 },
          overflow: "linebreak",
          valign: 'middle',
          halign: 'center',
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
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
          fontSize: 8.5,
          valign: 'middle',
          cellPadding: { top: 5, right: 2, bottom: 5, left: 2 },
        },
        alternateRowStyles: {
          fillColor: [248, 250, 252],
        },
        columnStyles,
        didParseCell: function (data) {
          if (data.section === 'body') {
            data.cell.styles.textColor = [0, 0, 0]; // Pure black text

            const rowIndex = data.row.index;
            const isTotalRow = rowIndex === tableBody.length - 1;

            if (isTotalRow) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [226, 232, 240];
              data.cell.styles.textColor = [0, 0, 0];
              data.cell.styles.halign = 'center';
              return;
            }

            const item = exportRows[rowIndex];
            if (!item) return;

            // Lot number bold black
            if (data.column.index === 1) {
              data.cell.styles.textColor = [0, 0, 0];
              data.cell.styles.fontStyle = 'bold';
            }

            // Total Qty bold black
            if (data.column.index === 6) {
              data.cell.styles.textColor = [0, 0, 0];
              data.cell.styles.fontStyle = 'bold';
            }

            // Days after PO issue soft highlight with pure black text
            if (data.column.index === 13) {
              const days = parseFloat(item["Days after PO issue"]);
              if (!isNaN(days)) {
                data.cell.styles.textColor = [0, 0, 0];
                data.cell.styles.fontStyle = 'bold';
                if (days > 2) {
                  data.cell.styles.fillColor = [254, 226, 226]; // Soft red
                } else {
                  data.cell.styles.fillColor = [220, 252, 231]; // Soft green
                }
              }
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

      const aBody = [
        ["Normal (<= 2 Days)", normalLots.toString(), "—", `${totalLots > 0 ? ((normalLots / totalLots) * 100).toFixed(1) : 0}%`],
        ["Red Zone (> 2 Days)", redZoneLots.toString(), "—", `${totalLots > 0 ? ((redZoneLots / totalLots) * 100).toFixed(1) : 0}%`],
        ["Repeated Lots", highPriorityCount.toString(), "—", `${totalLots > 0 ? ((highPriorityCount / totalLots) * 100).toFixed(1) : 0}%`],
        ["TOTAL LOTS", totalLots.toString(), totalQty.toLocaleString(), "100.0%"]
      ];

      const maxRows = Math.max(gBody.length, pBody.length, sBody.length, aBody.length);
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
      const summarySub = `Total Lots: ${totalLots}   |   Total Quantity: ${totalQty.toLocaleString()} Pcs   |   Parties: ${sortedParties.length}   |   Garments: ${sortedGarments.length}   |   Sections: ${sortedSections.length}`;
      doc.text(summarySub, pageW / 2, summaryStartY + 16, { align: 'center' });

      const sectionTitleY = summaryStartY + 30;
      const tableStartY = sectionTitleY + 6;

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
      doc.text("2. PARTY BREAKDOWN", col2X, sectionTitleY);
      doc.text("3. SECTION BREAKDOWN", col3X, sectionTitleY);
      doc.text("4. AGING & PRIORITY BREAKDOWN", col4X, sectionTitleY);

      const summaryColStyles = {
        0: { cellWidth: 110, halign: 'center' },
        1: { cellWidth: 45, halign: 'center' },
        2: { cellWidth: 68, halign: 'center' },
        3: { cellWidth: 55, halign: 'center' },
      };

      // Column 1 Table: Garment Breakdown
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
      const endY1 = doc.lastAutoTable.finalY;

      // Column 2 Table: Party Breakdown
      autoTable(doc, {
        head: [['Party Name', 'Lots', 'Total Qty', 'Share %']],
        body: pBody,
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
            if (data.row.index === pBody.length - 1) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });
      const endY2 = doc.lastAutoTable.finalY;

      // Column 3 Table: Section Breakdown
      autoTable(doc, {
        head: [['Section', 'Lots', 'Total Qty', 'Share %']],
        body: sBody,
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
            if (data.row.index === sBody.length - 1) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });
      const endY3 = doc.lastAutoTable.finalY;

      // Column 4 Table: Aging & Priority Breakdown
      autoTable(doc, {
        head: [['Category', 'Lots', 'Total Qty', 'Share %']],
        body: aBody,
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
              data.cell.styles.fillColor = [254, 226, 226]; // Soft Red
              data.cell.styles.fontStyle = 'bold';
            } else if (data.row.index === 2) {
              data.cell.styles.fillColor = [254, 243, 199]; // Soft Amber
              data.cell.styles.fontStyle = 'bold';
            } else if (data.row.index === 3) {
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
        doc.text("Cutting Department Report — Factory Suite Pro", 15, finalY + 12);
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
      doc.save(`Cutting_Department_Report_${fileDate}.pdf`);
    } catch (error) {
      console.error("Error generating Cutting Report PDF:", error);
      alert(`Failed to generate PDF: ${error.message}`);
    }
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

        .fabric-issued-badge {
          background: #ecfdf5;
          color: #047857;
          border: 1px solid #a7f3d0;
          font-weight: 700;
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

        /* Remarks Table Cell & Inline UI */
        .remarks-table-cell {
          vertical-align: middle;
        }

        .remarks-cell-container {
          display: flex;
          flex-direction: column;
          align-items: flex-start;
          gap: 6px;
          min-width: 190px;
          max-width: 280px;
        }

        .remark-bubble {
          background: #f1f5f9;
          border: 1px solid #cbd5e1;
          border-left: 3.5px solid #6366f1;
          border-radius: 8px;
          padding: 6px 10px;
          font-size: 0.82rem;
          color: #1e293b;
          width: 100%;
          box-sizing: border-box;
          text-align: left;
          transition: all 0.2s ease;
          cursor: pointer;
        }

        .remark-bubble:hover {
          background: #eef2ff;
          border-color: #a5b4fc;
          border-left-color: #4f46e5;
          box-shadow: 0 2px 8px rgba(99, 102, 241, 0.15);
        }

        .remark-text-preview {
          font-weight: 600;
          line-height: 1.35;
          word-break: break-word;
          white-space: normal;
          display: -webkit-box;
          -webkit-line-clamp: 2;
          -webkit-box-orient: vertical;
          overflow: hidden;
        }

        .remark-meta-preview {
          display: flex;
          align-items: center;
          justify-content: space-between;
          margin-top: 4px;
          font-size: 0.72rem;
          color: #64748b;
          font-weight: 500;
        }

        .remarks-actions-row {
          display: flex;
          align-items: center;
          gap: 6px;
          width: 100%;
        }

        .btn-add-remark-inline {
          background: linear-gradient(135deg, #eef2ff 0%, #e0e7ff 100%);
          color: #4338ca;
          border: 1px dashed #818cf8;
          border-radius: 8px;
          padding: 5px 12px;
          font-size: 0.78rem;
          font-weight: 700;
          cursor: pointer;
          display: inline-flex;
          align-items: center;
          gap: 4px;
          transition: all 0.2s ease;
        }

        .btn-add-remark-inline:hover {
          background: #4338ca;
          color: #ffffff;
          border-style: solid;
          border-color: #4338ca;
          transform: translateY(-1px);
          box-shadow: 0 4px 10px rgba(67, 56, 202, 0.2);
        }

        .btn-remark-history-badge {
          background: #ffffff;
          color: #6366f1;
          border: 1px solid #c7d2fe;
          border-radius: 6px;
          padding: 3px 8px;
          font-size: 0.72rem;
          font-weight: 700;
          cursor: pointer;
          display: inline-flex;
          align-items: center;
          gap: 3px;
          transition: all 0.15s ease;
        }

        .btn-remark-history-badge:hover {
          background: #6366f1;
          color: #ffffff;
        }

        .no-remarks-placeholder {
          color: #94a3b8;
          font-size: 0.8rem;
          font-style: italic;
        }

        /* Remarks Modal Dialog */
        .remarks-modal-overlay {
          position: fixed;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          background-color: rgba(15, 23, 42, 0.65);
          backdrop-filter: blur(8px);
          display: flex;
          align-items: center;
          justify-content: center;
          z-index: 99999;
          padding: 16px;
          animation: modalFadeIn 0.2s ease;
        }

        @keyframes modalFadeIn {
          from { opacity: 0; }
          to { opacity: 1; }
        }

        .remarks-modal-box {
          background: #ffffff;
          border-radius: 24px;
          max-width: 580px;
          width: 100%;
          max-height: 88vh;
          display: flex;
          flex-direction: column;
          box-shadow: 0 25px 60px -15px rgba(15, 23, 42, 0.35);
          border: 1px solid rgba(226, 232, 240, 0.8);
          overflow: hidden;
          animation: modalSlideUp 0.25s cubic-bezier(0.16, 1, 0.3, 1);
        }

        @keyframes modalSlideUp {
          from {
            opacity: 0;
            transform: translateY(20px) scale(0.97);
          }
          to {
            opacity: 1;
            transform: translateY(0) scale(1);
          }
        }

        .remarks-modal-header {
          background: linear-gradient(135deg, #1e1b4b 0%, #312e81 40%, #4338ca 100%);
          color: #ffffff;
          padding: 20px 24px;
          display: flex;
          align-items: center;
          justify-content: space-between;
          border-bottom: 1px solid rgba(255, 255, 255, 0.1);
        }

        .remarks-modal-header-title {
          display: flex;
          align-items: center;
          gap: 10px;
        }

        .remarks-modal-header-title h3 {
          margin: 0;
          font-size: 1.25rem;
          font-weight: 800;
          color: #ffffff;
          letter-spacing: -0.01em;
        }

        .remarks-modal-lot-tag {
          background: rgba(255, 255, 255, 0.2);
          color: #ffd700;
          padding: 3px 10px;
          border-radius: 12px;
          font-size: 0.85rem;
          font-weight: 800;
          letter-spacing: 0.02em;
        }

        .remarks-modal-close-btn {
          background: rgba(255, 255, 255, 0.15);
          border: 1px solid rgba(255, 255, 255, 0.2);
          color: #ffffff;
          font-size: 1.1rem;
          border-radius: 12px;
          width: 36px;
          height: 36px;
          cursor: pointer;
          display: flex;
          align-items: center;
          justify-content: center;
          transition: all 0.2s ease;
        }

        .remarks-modal-close-btn:hover {
          background: #ef4444;
          border-color: #ef4444;
          transform: rotate(90deg);
        }

        .remarks-modal-body {
          padding: 22px 24px;
          overflow-y: auto;
          flex: 1;
          display: flex;
          flex-direction: column;
          gap: 18px;
        }

        .remarks-lot-card {
          background: #f8fafc;
          border: 1px solid #e2e8f0;
          border-radius: 16px;
          padding: 12px 16px;
          display: grid;
          grid-template-columns: repeat(3, 1fr);
          gap: 10px;
        }

        .remarks-lot-field {
          display: flex;
          flex-direction: column;
        }

        .remarks-lot-field-label {
          font-size: 0.68rem;
          font-weight: 800;
          text-transform: uppercase;
          color: #64748b;
          letter-spacing: 0.04em;
        }

        .remarks-lot-field-value {
          font-size: 0.86rem;
          font-weight: 700;
          color: #1e293b;
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
        }

        .remarks-history-section {
          display: flex;
          flex-direction: column;
          gap: 10px;
        }

        .remarks-history-title {
          font-size: 0.82rem;
          font-weight: 800;
          color: #475569;
          text-transform: uppercase;
          letter-spacing: 0.05em;
          display: flex;
          align-items: center;
          gap: 6px;
          margin: 0;
        }

        .remarks-history-list {
          display: flex;
          flex-direction: column;
          gap: 10px;
          max-height: 180px;
          overflow-y: auto;
          padding-right: 4px;
        }

        .remarks-history-item {
          background: #ffffff;
          border: 1px solid #e2e8f0;
          border-left: 4px solid #6366f1;
          border-radius: 12px;
          padding: 10px 14px;
          box-shadow: 0 2px 4px rgba(0, 0, 0, 0.02);
          display: flex;
          flex-direction: column;
          gap: 4px;
        }

        .remarks-history-item.latest {
          background: #f5f7ff;
          border-color: #c7d2fe;
          border-left-color: #4338ca;
        }

        .remarks-history-text {
          font-size: 0.88rem;
          color: #0f172a;
          font-weight: 600;
          line-height: 1.4;
          word-break: break-word;
        }

        .remarks-history-time {
          font-size: 0.72rem;
          color: #64748b;
          font-weight: 600;
          display: flex;
          align-items: center;
          gap: 4px;
        }

        .remarks-input-section {
          display: flex;
          flex-direction: column;
          gap: 10px;
        }

        .remarks-input-label {
          font-size: 0.82rem;
          font-weight: 800;
          color: #1e1b4b;
          text-transform: uppercase;
          letter-spacing: 0.05em;
          display: flex;
          align-items: center;
          gap: 6px;
          margin: 0;
        }

        .remarks-quick-presets {
          display: flex;
          flex-wrap: wrap;
          gap: 6px;
        }

        .remark-preset-tag {
          background: #f1f5f9;
          border: 1px solid #cbd5e1;
          color: #334155;
          padding: 4px 10px;
          border-radius: 16px;
          font-size: 0.74rem;
          font-weight: 700;
          cursor: pointer;
          transition: all 0.15s ease;
          user-select: none;
        }

        .remark-preset-tag:hover {
          background: #e0e7ff;
          color: #4338ca;
          border-color: #a5b4fc;
          transform: translateY(-1px);
        }

        .remarks-textarea {
          width: 100%;
          min-height: 85px;
          padding: 12px 14px;
          border: 1.5px solid #cbd5e1;
          border-radius: 14px;
          font-family: inherit;
          font-size: 0.9rem;
          color: #0f172a;
          box-sizing: border-box;
          resize: vertical;
          transition: all 0.2s ease;
          outline: none;
        }

        .remarks-textarea:focus {
          border-color: #6366f1;
          box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.15);
          background: #ffffff;
        }

        .remarks-textarea::placeholder {
          color: #94a3b8;
        }

        .remarks-modal-footer {
          padding: 16px 24px;
          background: #f8fafc;
          border-top: 1px solid #e2e8f0;
          display: flex;
          align-items: center;
          justify-content: flex-end;
          gap: 12px;
        }

        .btn-cancel-remark {
          padding: 10px 20px;
          background: #ffffff;
          color: #475569;
          border: 1.5px solid #cbd5e1;
          border-radius: 14px;
          font-size: 0.88rem;
          font-weight: 700;
          cursor: pointer;
          transition: all 0.2s;
        }

        .btn-cancel-remark:hover {
          background: #f1f5f9;
          color: #0f172a;
        }

        .btn-save-remark {
          padding: 10px 24px;
          background: linear-gradient(135deg, #4338ca 0%, #3730a3 100%);
          color: #ffffff;
          border: none;
          border-radius: 14px;
          font-size: 0.88rem;
          font-weight: 800;
          cursor: pointer;
          display: flex;
          align-items: center;
          gap: 8px;
          transition: all 0.2s;
          box-shadow: 0 4px 14px rgba(67, 56, 202, 0.3);
        }

        .btn-save-remark:hover:not(:disabled) {
          transform: translateY(-2px);
          box-shadow: 0 8px 20px rgba(67, 56, 202, 0.4);
        }

        .btn-save-remark:disabled {
          opacity: 0.5;
          cursor: not-allowed;
          transform: none;
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

              {/* Manual Remarks filter */}
              <div className="filter-group">
                <label className="filter-label">💬 Manual Remarks</label>
                <select
                  className="filter-select"
                  value={manualRemarksFilter}
                  onChange={(e) => setManualRemarksFilter(e.target.value)}
                  disabled={loading}
                >
                  <option value="all">All (With & Without Remarks)</option>
                  <option value="WITH_REMARKS">📌 With Remarks Only</option>
                  <option value="WITHOUT_REMARKS">⚪ Without Remarks Only</option>
                  {distinctManualRemarks.map((rem, idx) => (
                    <option key={idx} value={rem}>
                      💬 {rem.length > 35 ? rem.substring(0, 35) + "..." : rem}
                    </option>
                  ))}
                </select>
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
                        if (h === "Cutting Scanned") {
                          const scannedDate = row["Cutting Scanned"];
                          return (
                            <td key={h} className="table-cell">
                              {scannedDate ? (
                                <span style={{
                                  display: 'inline-flex',
                                  alignItems: 'center',
                                  gap: '4px',
                                  padding: '2px 8px',
                                  borderRadius: '6px',
                                  fontSize: '0.78rem',
                                  fontWeight: '700',
                                  background: '#dcfce7',
                                  color: '#15803d',
                                  border: '1px solid #86efac',
                                  whiteSpace: 'nowrap'
                                }} title={`Scanned: ${scannedDate}${row["Cutting Scanned Status"] ? ` (${row["Cutting Scanned Status"]})` : ''}`}>
                                  📷 {scannedDate}
                                </span>
                              ) : (
                                <span style={{ color: '#9ca3af', fontSize: '12px' }}>—</span>
                              )}
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
                          const fabricIssueDate = row.fabricIssueDate || "";
                          return (
                            <td key={h} className="table-cell">
                              <div className="remarks-badges">
                                {splitRemarks(remarks).map((remark, idx) => {
                                  const isFabricIssued = remark.toLowerCase().includes("fabric issued") || remark.toLowerCase() === "fabric issued";
                                  return (
                                    <div
                                      key={idx}
                                      style={{
                                        display: 'inline-flex',
                                        flexDirection: 'column',
                                        alignItems: 'center',
                                        gap: '2px'
                                      }}
                                    >
                                      <span
                                        className={`remark-badge ${
                                          remark.includes("Cancel")
                                            ? "cancel-badge"
                                            : isFabricIssued
                                              ? "fabric-issued-badge"
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
                                      {isFabricIssued && fabricIssueDate && (
                                        <span
                                          className="fabric-issue-date-sub"
                                          style={{
                                            fontSize: '0.68rem',
                                            fontWeight: '700',
                                            color: '#065f46',
                                            backgroundColor: '#d1fae5',
                                            border: '1px solid #6ee7b7',
                                            padding: '1px 6px',
                                            borderRadius: '6px',
                                            whiteSpace: 'nowrap',
                                            marginTop: '1px',
                                            display: 'inline-flex',
                                            alignItems: 'center',
                                            gap: '3px'
                                          }}
                                          title={`Fabric Issued on ${fabricIssueDate}`}
                                        >
                                          📅 {fabricIssueDate}
                                        </span>
                                      )}
                                    </div>
                                  );
                                })}
                              </div>
                            </td>
                          );
                        }

                        if (h === "💬 User Remarks") {
                          const lot = String(row["Lot No"] || "").trim();
                          const lotRemarks = remarksMap[lot] || [];
                          const latestRemark = lotRemarks.length > 0 ? lotRemarks[lotRemarks.length - 1] : null;

                          return (
                            <td key={h} className="table-cell remarks-table-cell">
                              <div className="remarks-cell-container">
                                {latestRemark ? (
                                  <div
                                    className="remark-bubble"
                                    onClick={() => handleOpenRemarksModal(row)}
                                    title="Click to view history or add new remark"
                                  >
                                    <div className="remark-text-preview">
                                      {latestRemark.text}
                                    </div>
                                    <div className="remark-meta-preview">
                                      <span>🕒 {latestRemark.timestamp}</span>
                                    </div>
                                  </div>
                                ) : (
                                  <span className="no-remarks-placeholder">No remarks yet</span>
                                )}

                                <div className="remarks-actions-row">
                                  <button
                                    type="button"
                                    className="btn-add-remark-inline"
                                    onClick={() => handleOpenRemarksModal(row)}
                                    title="Add or update remark for this lot"
                                  >
                                    {latestRemark ? '✏️ Remark' : '+ Add Remark'}
                                  </button>

                                  {lotRemarks.length > 1 && (
                                    <button
                                      type="button"
                                      className="btn-remark-history-badge"
                                      onClick={() => handleOpenRemarksModal(row)}
                                      title="View all remarks history"
                                    >
                                      📜 ({lotRemarks.length})
                                    </button>
                                  )}
                                </div>
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

        {/* Interactive Remarks Modal */}
        {remarksModalOpen && selectedRemarksLot && (
          <div className="remarks-modal-overlay" onClick={handleCloseRemarksModal}>
            <div className="remarks-modal-box" onClick={(e) => e.stopPropagation()}>
              <div className="remarks-modal-header">
                <div className="remarks-modal-header-title">
                  <span style={{ fontSize: '1.4rem' }}>💬</span>
                  <h3>Cutting Lot Remarks</h3>
                  <span className="remarks-modal-lot-tag">#{selectedRemarksLot["Lot No"]}</span>
                </div>
                <button
                  type="button"
                  className="remarks-modal-close-btn"
                  onClick={handleCloseRemarksModal}
                  title="Close"
                >
                  ✕
                </button>
              </div>

              <div className="remarks-modal-body">
                {/* Lot Information Summary */}
                <div className="remarks-lot-card">
                  <div className="remarks-lot-field">
                    <span className="remarks-lot-field-label">Party</span>
                    <span className="remarks-lot-field-value">{selectedRemarksLot["Party Name"] || '-'}</span>
                  </div>
                  <div className="remarks-lot-field">
                    <span className="remarks-lot-field-label">Fabric</span>
                    <span className="remarks-lot-field-value">{selectedRemarksLot["Fabric"] || '-'}</span>
                  </div>
                  <div className="remarks-lot-field">
                    <span className="remarks-lot-field-label">Style</span>
                    <span className="remarks-lot-field-value">{selectedRemarksLot["Style"] || selectedRemarksLot["Garment Type"] || '-'}</span>
                  </div>
                  <div className="remarks-lot-field">
                    <span className="remarks-lot-field-label">Total Qty</span>
                    <span className="remarks-lot-field-value" style={{ color: '#0f766e', fontWeight: '800' }}>
                      {fmtNum(selectedRemarksLot["Total Qty"]) || '0'} pcs
                    </span>
                  </div>
                  <div className="remarks-lot-field">
                    <span className="remarks-lot-field-label">Section</span>
                    <span className="remarks-lot-field-value">{selectedRemarksLot["Section"] || '-'}</span>
                  </div>
                  <div className="remarks-lot-field">
                    <span className="remarks-lot-field-label">Days after PO</span>
                    <span className="remarks-lot-field-value" style={{ color: '#dc2626' }}>
                      {selectedRemarksLot["Days after PO issue"] || 0} days
                    </span>
                  </div>
                </div>

                {/* Remarks History */}
                <div className="remarks-history-section">
                  <h4 className="remarks-history-title">
                    <span>📜</span> Remarks History ({(remarksMap[String(selectedRemarksLot["Lot No"] || '').trim()] || []).length})
                  </h4>
                  <div className="remarks-history-list">
                    {(() => {
                      const lotKey = String(selectedRemarksLot["Lot No"] || '').trim();
                      const history = remarksMap[lotKey] || [];
                      if (history.length === 0) {
                        return (
                          <div style={{ textAlign: 'center', padding: '16px', color: '#94a3b8', fontSize: '0.85rem' }}>
                            No previous remarks for this lot. Add the first remark below!
                          </div>
                        );
                      }
                      return history.map((item, i) => (
                        <div
                          key={i}
                          className={`remarks-history-item ${i === history.length - 1 ? 'latest' : ''}`}
                        >
                          <div className="remarks-history-text">{item.text}</div>
                          <div className="remarks-history-time">
                            <span>🕒</span> {item.timestamp}
                            {i === history.length - 1 && (
                              <span style={{ marginLeft: 'auto', color: '#4338ca', fontWeight: '700', fontSize: '0.7rem', background: '#e0e7ff', padding: '2px 6px', borderRadius: '4px' }}>
                                Latest
                              </span>
                            )}
                          </div>
                        </div>
                      ));
                    })()}
                  </div>
                </div>

                {/* Enter New Remark Input */}
                <div className="remarks-input-section">
                  <h4 className="remarks-input-label">
                    <span>✏️</span> Enter Your Remark
                  </h4>

                  {/* Quick Cutting Presets */}
                  <div className="remarks-quick-presets">
                    {[
                      "Cutting in Progress",
                      "Fabric Issue Pending",
                      "Fabric Shortage",
                      "Sample / Pattern Pending",
                      "Urgent Cutting Required",
                      "Hold for Quality Check",
                      "Cutting Completed",
                      "Ready for Numbering / Fusing"
                    ].map((preset) => (
                      <button
                        key={preset}
                        type="button"
                        className="remark-preset-tag"
                        onClick={() => setNewRemarkInputText(preset)}
                      >
                        + {preset}
                      </button>
                    ))}
                  </div>

                  <textarea
                    className="remarks-textarea"
                    placeholder="Type your custom cutting remark here..."
                    value={newRemarkInputText}
                    onChange={(e) => setNewRemarkInputText(e.target.value)}
                    rows={3}
                    autoFocus
                  />
                </div>
              </div>

              <div className="remarks-modal-footer">
                <button
                  type="button"
                  className="btn-cancel-remark"
                  onClick={handleCloseRemarksModal}
                  disabled={savingRemark}
                >
                  Cancel
                </button>
                <button
                  type="button"
                  className="btn-save-remark"
                  onClick={handleSaveRemark}
                  disabled={savingRemark || !newRemarkInputText.trim()}
                >
                  {savingRemark ? '⏳ Saving...' : '💾 Save Remark'}
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