import React, { useEffect, useMemo, useState, useRef } from "react";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { GOOGLE_API_KEY, SPREADSHEET_IDS, fetchSheetDataFromBackend } from "./config";

/** ====== CONFIG ====== */
const JOB_SHEET_ID = SPREADSHEET_IDS.JOBORDER;
const API_KEY = GOOGLE_API_KEY;
const JOB_RANGE = "JobOrder!A:AL";

// Budget Report spreadsheet
const BUDGET_SHEET_ID = SPREADSHEET_IDS.MAIN;
const CUTTING_SHEET_NAME = "Cutting";
const INDEX_SHEET_NAME = "Index";
const INDEX_RANGE = `${INDEX_SHEET_NAME}!A:K`;

// One big read to avoid 429s
const CUTTING_BIG_RANGE = `${CUTTING_SHEET_NAME}!A1:ZZ200000`;

// Canonical output columns (order)
const OUTPUT_COLS = [
  "Job Order No",
  "Fabric",
  "Brand",
  "Style",
  "Party Name",
  "Garment Type",
  "Section",
  "Season",
  "Direct Stitching",
  "Lot No",
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
  date: "PO Date",
  status: "Status", // NEW: read Status from JobOrder
};

/** ====== UTILS ====== */
const norm = (s) => String(s || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
function normalizeKey(s = "") {
  return norm(s);
}

// Formats "Fri Sep 19 2025 12:30:10 GMT+0530 (India Standard Time)" -> "2025-09-19"
function formatSavedAtToYMD(savedAt) {
  if (!savedAt) return "";
  const d = new Date(savedAt);
  if (isNaN(d.getTime())) return "";
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
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
      "Fabric",
      "Brand",
      "Style",
      "Party Name",
      "Garment Type",
      "Section",
      "Season",
      "Direct Stitching",
      "Lot No",
      "PO Date",
      "Status", // NEW
    ].forEach((canonHeader) => {
      const idx = canonToIndex[canonHeader];
      obj[canonHeader] = idx != null ? (row[idx] ?? "") : "";
    });
    obj["Days after PO issue"] = "";
    obj["Total Qty"] = 0;
    obj["Pending Shade"] = "";
    obj["Remarks"] = "";
    obj["Remarks 2"] = "";
    obj["Remarks 3"] = "";
    obj["Cutting Date"] = ""; // fill later
    return obj;
  });
}

/** ====== PDF THEME ====== */
const PDF_THEME = {
  brandName: "Cutting Report",
  primary: [24, 90, 188],
  muted: [90, 107, 130],
  border: [220, 226, 235],
  zebra: [248, 250, 253],
  badgeDoneBG: [227, 253, 240],
  badgeDoneText: [10, 94, 62],
  badgePendingBG: [255, 245, 233],
  badgePendingText: [120, 70, 0],
  badgeIssueBG: [255, 233, 233],
  badgeIssueText: [144, 0, 0],
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
function parseIndexRow(header, row) {
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
  const savedAt = get("savedat"); // "Saved At"

  const sizes = String(get("sizes") || "")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);

  const shades = String(get("shades") || "")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);

  return { lot, startRow, numRows, headerCols, fabric, garmentType, style, sizes, shades, savedAt };
}

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

function computeTotalsAndPending(windowValues, sizes = [], shades = []) {
  if (!windowValues || windowValues.length === 0) {
    return { totalQty: 0, pendingShadeKeys: new Set(shades.map(norm)) };
  }

  const normalizedSizes = Array.from(new Set((sizes || []).map(norm).filter(Boolean)));
  const headerRowIdx = findHeaderRowIndex(windowValues, normalizedSizes);
  const header = windowValues[headerRowIdx] || [];

  const hIdx = {};
  header.forEach((h, i) => {
    const k = norm(h);
    if (k && !(k in hIdx)) hIdx[k] = i;
  });

  const shadeColIndex = hIdx["color"] ?? hIdx["shade"] ?? hIdx["shades"] ?? 0;

  let sizeColIndices = [];
  normalizedSizes.forEach((ns) => {
    if (ns in hIdx) sizeColIndices.push(hIdx[ns]);
  });

  if (sizeColIndices.length === 0 || sizeColIndices.length !== normalizedSizes.length) {
    const ct = hIdx["cuttingtable"];
    if (ct != null && ct >= 0) {
      const guessStart = ct + 1;
      const guessed = [];
      for (let k = 0; k < normalizedSizes.length; k++) guessed.push(guessStart + k);
      sizeColIndices = Array.from(new Set(guessed.filter((g) => g < header.length)));
    }
  }

  if (sizeColIndices.length === 0) {
    return { totalQty: 0, pendingShadeKeys: new Set(shades.map(norm)) };
  }

  const shadeStats = new Map();
  let totalQty = 0;

  // Track which shades have ANY cutting data
  for (let r = headerRowIdx + 1; r < windowValues.length; r++) {
    const row = windowValues[r] || [];
    const rawShade = String(row[shadeColIndex] || "").trim();
    const shadeKey = norm(rawShade);

    if (!shadeKey || shadeKey === "total" || shadeKey === "totals" || shadeKey === "grandtotal") {
      continue;
    }

    // Initialize or get existing stats
    if (!shadeStats.has(shadeKey)) {
      shadeStats.set(shadeKey, {
        hasAnyData: false,
        totalCut: 0,
        sizeCount: 0,
      });
    }
    
    const stat = shadeStats.get(shadeKey);
    let rowHasData = false;

    // Check each size column
    sizeColIndices.forEach((c) => {
      const raw = row[c];
      if (raw != null && raw !== "") {
        const n = parseFloat(String(raw).replace(/,/g, ""));
        if (!isNaN(n)) {
          stat.hasAnyData = true;
          rowHasData = true;
          if (n > 0) {
            stat.totalCut += n;
            stat.sizeCount += 1;
          }
        }
      }
    });

    if (rowHasData) {
      totalQty += stat.totalCut;
    }
  }

  // Determine pending shades
  const pendingShadeKeys = new Set();
  const expectedShadeKeys = (shades || []).map(norm);

  expectedShadeKeys.forEach((shadeKey) => {
    const stat = shadeStats.get(shadeKey);

    if (!stat) {
      // Shade not found at all → PENDING
      pendingShadeKeys.add(shadeKey);
      return;
    }

    if (!stat.hasAnyData) {
      // Shade found but all cells are empty/blank → PENDING
      pendingShadeKeys.add(shadeKey);
      return;
    }

    if (stat.totalCut === 0) {
      // Shade found but all quantities are 0 → NOT PENDING (intentionally zero)
      return;
    }

    // Shade has some cutting data (>0) → NOT PENDING
    // Even if not all sizes are filled, it's still "cut"
  });

  return { totalQty, pendingShadeKeys };
}

/** ====== COMPONENT ====== */
export default function CuttingStatsReport() {
  const [rows, setRows] = useState([]);
  const [lotFilter, setLotFilter] = useState("");
  const [partyFilter, setPartyFilter] = useState("");
  const [garmentFilter, setGarmentFilter] = useState("");
  const [seasonFilter, setSeasonFilter] = useState("");
  const [loading, setLoading] = useState(true);
  const [refreshing, setRefreshing] = useState(false);
  const [err, setErr] = useState("");
  const [lastUpdated, setLastUpdated] = useState(null);

  const [dialogOpen, setDialogOpen] = useState(false);
  const [dialogLot, setDialogLot] = useState("");
  const [dialogShades, setDialogShades] = useState([]);

  const [pendingListByLot, setPendingListByLot] = useState({});
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(25);
  const [selectedRemarks, setSelectedRemarks] = useState(new Set());

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

  const distinctRemarks = useMemo(() => {
    const set = new Set();
    rows.forEach((r) => splitRemarks(r.Remarks).forEach((t) => set.add(t)));
    return Array.from(set).sort((a, b) => a.localeCompare(b));
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
    setGarmentFilter("");
    setSeasonFilter("");
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
    } catch {}
    const ctrl = new AbortController();
    abortRef.current = ctrl;

    try {
      // Job Orders
      const job = await fetchSheet(
        { sheetId: JOB_SHEET_ID, range: JOB_RANGE, apiKey: API_KEY, signal: ctrl.signal }
      );
      let jobRows = convertValuesToObjects(job.values);

      // Exclude lots whose status starts with "cancel" (Cancelled, Canceled, etc.)
      jobRows = jobRows.filter((r) => {
        const s = (r.Status ?? "").toString();
        const sn = norm(s);
        return !sn.startsWith("cancel");
      });

      // Index sheet (for sizes/shades/savedAt)
      const idxRes = await fetchSheet(
        { sheetId: BUDGET_SHEET_ID, range: INDEX_RANGE, apiKey: API_KEY, signal: ctrl.signal }
      );
      const idxValues = idxRes.values || [];
      const idxHeader = idxValues[0] || [];
      const indexMap = new Map();
      for (let i = 1; i < idxValues.length; i++) {
        const entry = parseIndexRow(idxHeader, idxValues[i]);
        if (entry) indexMap.set(entry.lot, entry);
      }

      // Cutting big read
      const cuttingRes = await fetchSheet(
        { sheetId: BUDGET_SHEET_ID, range: CUTTING_BIG_RANGE, apiKey: API_KEY, signal: ctrl.signal }
      );
      const bigCuttingValues = cuttingRes.values || [];

      // Per-lot summaries
      const lots = Array.from(new Set(jobRows.map((r) => String(r["Lot No"] || "").trim()).filter(Boolean)));
      const lotToSummary = new Map();
      const pendingListTmp = {};

      for (const lot of lots) {
        const ix = indexMap.get(lot);
        if (!ix) {
          lotToSummary.set(lot, {
            totalQty: 0,
            remarks: "",
            remarks2: "",
            remarks3: "Fabric Issue Pending",
            cuttingDate: "",
          });
          pendingListTmp[lot] = [];
          continue;
        }

        const window = sliceCuttingMatrix(bigCuttingValues, ix.startRow, ix.numRows);
        const { totalQty, pendingShadeKeys } = computeTotalsAndPending(window, ix.sizes, ix.shades);
        const cuttingDate = formatSavedAtToYMD(ix.savedAt);

        const shadeKeyToOriginal = new Map((ix.shades || []).map((sh) => [norm(sh), sh]));
        const pendingList = Array.from(pendingShadeKeys).map((k) => shadeKeyToOriginal.get(k) || k);

        let remarks = "";
        let remarks2 = "";
        let remarks3 = "";

        if (pendingShadeKeys.size > 0) {
          remarks2 = "Cutting Pending";
        } else {
          remarks = "Cutting Done";
        }

        lotToSummary.set(lot, { totalQty, remarks, remarks2, remarks3, cuttingDate });
        pendingListTmp[lot] = pendingList;
      }

      const merged = jobRows.map((r) => {
        const lot = String(r["Lot No"] || "").trim();
        const sum =
          lotToSummary.get(lot) ||
          {
            totalQty: 0,
            remarks: "",
            remarks2: "",
            remarks3: lot ? "Fabric Issue Pending" : "",
            cuttingDate: "",
          };
        const days = daysAfter(r["PO Date"] || r["Date"], sum.cuttingDate);

        const remarksList = [sum.remarks, sum.remarks2, sum.remarks3].filter(Boolean);
        const mergedRemarks = remarksList.join(" | ");

        return {
          ...r,
          "Days after PO issue": days ?? "",
          "Total Qty": sum.totalQty,
          "Pending Shade": "",
          "Cutting Date": sum.cuttingDate || "",
          Remarks: mergedRemarks,
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
      } catch {}
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const filtered = useMemo(() => {
    const lotQ = lotFilter.trim().toLowerCase();
    const partyQ = partyFilter.trim().toLowerCase();
    const garmentQ = garmentFilter.trim().toLowerCase();
    const seasonQ = seasonFilter.trim().toLowerCase();
    const selected = selectedRemarks;

    return rows.filter((r) => {
      if (lotQ) {
        const lot = String(r["Lot No"] || "").toLowerCase();
        if (!lot.includes(lotQ)) return false;
      }
      if (partyQ) {
        const party = String(r["Party Name"] || "").toLowerCase();
        if (!party.includes(partyQ)) return false;
      }
      if (garmentQ) {
        const garment = String(r["Garment Type"] || "").toLowerCase();
        if (!garment.includes(garmentQ)) return false;
      }
      if (seasonQ) {
        const season = String(r["Season"] || "").toLowerCase();
        if (!season.includes(seasonQ)) return false;
      }

      if (selected.size > 0) {
        const tokens = splitRemarks(r.Remarks);
        const tokenSet = new Set(tokens);
        let matchesAny = false;
        for (const sel of selected) {
          if (tokenSet.has(sel)) {
            matchesAny = true;
            break;
          }
        }
        if (!matchesAny) return false;
      }

      return true;
    });
  }, [rows, lotFilter, partyFilter, garmentFilter, seasonFilter, selectedRemarks]);

  useEffect(() => {
    setPage(1);
  }, [lotFilter, partyFilter, garmentFilter, seasonFilter, pageSize, rows]);

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

  // Define the new header order
  const HEADER_ORDER = [
    "Job Order No",
    "Lot No",
    "Fabric",
    "Brand", 
    "Style",
    "Garment Type",
    "Party Name",
    "Section",
    "Season",
    "Direct Stitching",
    "Days after PO issue",
    "Total Qty",
    "Pending Shade",
    "Cutting Date",
    "Remarks"
  ];

  // Filter OUTPUT_COLS to ensure we have all headers in the correct order
  const finalHeaders = HEADER_ORDER.filter(h => OUTPUT_COLS.includes(h));
  
  const head = [finalHeaders];
  
  // Prepare body data with all formatting done upfront
  const body = exportRows.map((row) => {
    return finalHeaders.map((h) => {
      // Special handling for "Party Name" - replace "Mohit Hosiery" with "MH"
      if (h === "Party Name" && row[h]) {
        const partyName = row[h].toString().trim();
        if (partyName.toLowerCase().includes("mohit")) {
          return "MH";
        }
        return partyName;
      }
      
      // Format Total Qty
      if (h === "Total Qty") {
        return fmtNum(row[h]);
      }
      
      // For Remarks - clean up and prepare for display
      if (h === "Remarks" && row[h]) {
        const rawRemarks = String(row[h]).trim();
        if (!rawRemarks) return "";
        
        // Split by pipe, clean, and remove duplicates
        const parts = rawRemarks.split("|")
          .map(p => p.trim())
          .filter(p => p && p !== "" && !p.includes("undefined"));
        
        // Remove duplicates (case insensitive)
        const uniqueParts = [];
        const seen = new Set();
        parts.forEach(part => {
          const lowerPart = part.toLowerCase();
          if (!seen.has(lowerPart)) {
            seen.add(lowerPart);
            uniqueParts.push(part);
          }
        });
        
        return uniqueParts.join(", ");
      }
      
      return row[h] ?? "";
    });
  });

  // Calculate total width to ensure it fits
  const columnWidths = {
    "Job Order No": 60,
    "Lot No": 50,
    "Fabric": 90,
    "Brand": 75,
    "Style": 80,
    "Garment Type": 85,
    "Party Name": 55,
    "Section": 45,
    "Season": 45,
    "Direct Stitching": 50,
    "Days after PO issue": 50,
    "Total Qty": 45,
    "Pending Shade": 120,
    "Cutting Date": 65,
    "Remarks": 80
  };

  const totalWidth = finalHeaders.reduce((sum, h) => sum + (columnWidths[h] || 60), 0);
  
  // Use A4 if content is smaller, otherwise use A3
  const pageFormat = totalWidth > 1000 ? "a3" : "a4";
  
  const doc = new jsPDF({ 
    orientation: "landscape", 
    unit: "pt", 
    format: pageFormat
  });
  
  const W = doc.internal.pageSize.getWidth();
  const H = doc.internal.pageSize.getHeight();

  const leftX = 24;
  const topY = 24;

  const { primary, muted, border, zebra, brandName } = PDF_THEME;

  // ===== Title block =====
  doc.setFont("helvetica", "bold");
  doc.setFontSize(13);
  doc.setTextColor(primary[0], primary[1], primary[2]);
  doc.text(brandName, leftX, topY);

  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  doc.setTextColor(muted[0], muted[1], muted[2]);
  doc.text(`Generated: ${new Date().toLocaleString()}`, leftX, topY + 12);

  // thin rule under subtitle
  doc.setDrawColor(border[0], border[1], border[2]);
  doc.setLineWidth(0.5);
  doc.line(leftX, topY + 18, W - leftX, topY + 18);

  // Define column styles with proper widths
  const columnStyles = {};
  
  finalHeaders.forEach((header, index) => {
    columnStyles[index] = {
      cellWidth: columnWidths[header] || 60,
      halign: header === "Total Qty" ? "right" : "center"
    };
    
    // Special styling for Lot No - make it red and bold
    if (header === "Lot No") {
      columnStyles[index].fontStyle = "bold";
      columnStyles[index].textColor = [255, 0, 0];
    }
    
    // Special styling for Remarks
    if (header === "Remarks") {
      columnStyles[index].halign = "left";
      columnStyles[index].fontSize = 7.5;
    }
  });

  // Generate table with autoTable
  autoTable(doc, {
    head: head,
    body: body,
    startY: topY + 28,
    
    styles: {
      font: "helvetica",
      fontSize: 8.5,
      cellPadding: {
        top: 4,
        right: 2,
        bottom: 4,
        left: 2
      },
      minCellHeight: 18,
      overflow: "linebreak",
      lineColor: border,
      lineWidth: 0.35,
      textColor: [0, 0, 0],
    },

    headStyles: {
      fontStyle: "bold",
      fontSize: 8.5,
      halign: "center",
      fillColor: [79, 70, 229],
      textColor: [255, 255, 255],
      cellPadding: {
        top: 4,
        right: 2,
        bottom: 4,
        left: 2
      },
      lineColor: border,
      lineWidth: 0.35,
    },

    alternateRowStyles: { 
      fillColor: zebra 
    },

    columnStyles: columnStyles,

    margin: { 
      left: leftX, 
      right: leftX, 
      top: 28, 
      bottom: 44 
    },

    tableWidth: "auto",
    
    // Handle page breaks
    didDrawPage: (hookData) => {
      // Simple border
      doc.setDrawColor(border[0], border[1], border[2]);
      doc.setLineWidth(0.6);
      doc.roundedRect(12, 12, W - 24, H - 24, 2, 2, "S");

      // Footer text
      doc.setFont("helvetica", "normal");
      doc.setFontSize(7.5);
      doc.setTextColor(muted[0], muted[1], muted[2]);
      doc.text("Cutting Stats", leftX, H - 20);

      // Page numbers
      const totalPages = doc.internal.getNumberOfPages();
      const pageNumber = hookData.pageNumber;
      const pageText = `Page ${pageNumber} of ${totalPages}`;
      const textWidth = doc.getTextWidth(pageText);
      doc.text(pageText, W - leftX - textWidth, H - 20);
    },
  });

  // Save with a simple filename
  doc.save(`Cutting_Stats_${todayYMD()}.pdf`);
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
          background: #ffffffff;
          font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
          color: #1e293b;
        }

        /* Loader */
        .loader-overlay {
          position: fixed;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          background: rgba(255, 255, 255, 0.95);
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          z-index: 1000;
        }

        .loader-spinner {
          width: 48px;
          height: 48px;
          border: 3px solid #e2e8f0;
          border-top: 3px solid #3b82f6;
          border-radius: 50%;
          animation: spin 1s linear infinite;
          margin-bottom: 16px;
        }

        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }

        .loader-text {
          font-size: 16px;
          font-weight: 500;
          color: #64748b;
        }

        /* Main Layout */
        .main-content {
          max-width: 2440px;
          margin: 0 auto;
          padding: 24px;
        }

        /* Professional Header */
        .header {
         background: #4331a8ff;
          border-radius: 12px;
          padding: 32px;
          margin-bottom: 24px;
          border: 1px solid #e2e8f0;
          box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
        }

        .header-content {
          display: flex;
          justify-content: space-between;
          align-items: center;
        }

        .header-title {
          margin: 0 0 8px 0;
          font-size: 28px;
          font-weight: 700;
          color: #ffffffff;
        }

        .header-subtitle {
          margin: 0;
          font-size: 16px;
          color: #ffffffff;
          font-weight: 400;
        }

        .stats-grid {
          display: flex;
          gap: 20px;
        }

        .stat-card {
          text-align: center;
          padding: 20px;
          background: white;
          border-radius: 8px;
          border: 1px solid #e2e8f0;
          min-width: 120px;
          box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05);
        }

        .stat-value {
          display: block;
          font-size: 28px;
          font-weight: 700;
          color: #3b82f6;
          margin-bottom: 4px;
        }

        .stat-label {
          font-size: 12px;
          color: #64748b;
          text-transform: uppercase;
          letter-spacing: 0.5px;
          font-weight: 500;
        }

        /* Professional Toolbar */
        .toolbar {
          background: white;
          border-radius: 12px;
          padding: 24px;
          margin-bottom: 24px;
          border: 1px solid #e2e8f0;
          box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
        }

        .toolbar-section {
          margin-bottom: 24px;
          padding-bottom: 24px;
          border-bottom: 1px solid #e2e8f0;
        }

        .toolbar-section:last-child {
          margin-bottom: 0;
          padding-bottom: 0;
          border-bottom: none;
        }

        .section-title {
          font-size: 16px;
          font-weight: 600;
          color: #1e293b;
          margin-bottom: 16px;
          display: flex;
          align-items: center;
          gap: 8px;
        }

        .section-title::before {
          content: '';
          width: 4px;
          height: 16px;
          background: #3b82f6;
          border-radius: 2px;
        }

        .filter-section {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
          gap: 16px;
        }

        .filter-group {
          display: flex;
          flex-direction: column;
          gap: 8px;
        }

        .filter-label {
          font-size: 14px;
          color: #374151;
          font-weight: 500;
        }

        .filter-input, .filter-select {
          padding: 10px 12px;
          border: 1px solid #d1d5db;
          border-radius: 6px;
          background: white;
          color: #1e293b;
          font-size: 14px;
          transition: all 0.2s ease;
        }

        .filter-input::placeholder {
          color: #9ca3af;
        }

        .filter-input:focus, .filter-select:focus {
          outline: none;
          border-color: #3b82f6;
          box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.1);
        }

        .remarks-grid {
          display: flex;
          flex-wrap: wrap;
          gap: 8px;
        }

        .remark-chip {
          display: flex;
          align-items: center;
          gap: 8px;
          padding: 8px 16px;
          background: #f8fafc;
          border: 1px solid #e2e8f0;
          border-radius: 20px;
          color: #374151;
          font-size: 14px;
          font-weight: 500;
          cursor: pointer;
          transition: all 0.2s ease;
        }

        .remark-chip:hover {
          background: #f1f5f9;
          border-color: #cbd5e1;
        }

        .remark-chip.selected {
          background: #3b82f6;
          border-color: #3b82f6;
          color: white;
        }

        .remark-checkbox {
          width: 16px;
          height: 16px;
          border-radius: 4px;
          border: 2px solid #d1d5db;
          display: flex;
          align-items: center;
          justify-content: center;
          transition: all 0.2s ease;
        }

        .remark-chip.selected .remark-checkbox {
          background: white;
          border-color: white;
        }

        .remark-chip.selected .remark-checkbox::after {
          content: '✓';
          color: #3b82f6;
          font-size: 12px;
          font-weight: bold;
        }

        .controls-section {
          display: flex;
          justify-content: space-between;
          align-items: center;
          gap: 16px;
          flex-wrap: wrap;
        }

        .left-controls {
          display: flex;
          align-items: center;
          gap: 20px;
        }

        .page-size-control {
          display: flex;
          align-items: center;
          gap: 12px;
          font-size: 14px;
          color: #374151;
        }

        .control-label {
          display: flex;
          align-items: center;
          gap: 8px;
          font-weight: 500;
        }

        .page-size-select {
          padding: 8px 12px;
          border: 1px solid #d1d5db;
          border-radius: 6px;
          background: white;
          color: #1e293b;
          font-size: 14px;
          transition: all 0.2s ease;
        }

        .page-size-select:focus {
          outline: none;
          border-color: #3b82f6;
        }

        .button-group {
          display: flex;
          gap: 12px;
          flex-wrap: wrap;
        }

        .btn {
          display: flex;
          align-items: center;
          gap: 8px;
          padding: 10px 20px;
          border: 1px solid;
          border-radius: 6px;
          font-size: 14px;
          font-weight: 500;
          cursor: pointer;
          transition: all 0.2s ease;
          text-decoration: none;
        }

        .btn:disabled {
          opacity: 0.6;
          cursor: not-allowed;
          transform: none !important;
        }

        .btn-excel {
          background: #ecfdf5;
          color: #065f46;
          border-color: #a7f3d0;
        }

        .btn-excel:hover:not(:disabled) {
          background: #d1fae5;
          transform: translateY(-1px);
        }

        .btn-pdf {
          background: #fef2f2;
          color: #991b1b;
          border-color: #fecaca;
        }

        .btn-pdf:hover:not(:disabled) {
          background: #fee2e2;
          transform: translateY(-1px);
        }

        .btn-refresh {
          background: #eff6ff;
          color: #1e40af;
          border-color: #bfdbfe;
        }

        .btn-refresh:hover:not(:disabled) {
          background: #dbeafe;
          transform: translateY(-1px);
        }

        .btn-back {
          background: #f8fafc;
          color: #374151;
          border-color: #e2e8f0;
        }

        .btn-back:hover {
          background: #f1f5f9;
          transform: translateY(-1px);
        }

        .btn-clear {
          background: #fffbeb;
          color: #92400e;
          border-color: #fed7aa;
        }

        .btn-clear:hover {
          background: #fef3c7;
          transform: translateY(-1px);
        }

        .btn-icon {
          font-size: 16px;
        }

        /* Error Bar */
        .error-bar {
          background: #fef2f2;
          color: #dc2626;
          padding: 16px 20px;
          border-radius: 8px;
          margin-bottom: 20px;
          border: 1px solid #fecaca;
        }

        .error-content {
          display: flex;
          align-items: center;
          gap: 12px;
        }

        .error-icon {
          font-size: 18px;
        }

        .retry-btn {
          background: #dc2626;
          color: white;
          border: none;
          padding: 8px 16px;
          border-radius: 6px;
          font-size: 13px;
          font-weight: 500;
          cursor: pointer;
          margin-left: auto;
          transition: background 0.2s;
        }

        .retry-btn:hover {
          background: #b91c1c;
        }

        /* Meta Info */
        .meta-info {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-bottom: 20px;
          color: #64748b;
          font-size: 14px;
        }

        .meta-left, .meta-right {
          display: flex;
          align-items: center;
          gap: 12px;
        }

        .last-updated {
          display: flex;
          align-items: center;
          gap: 8px;
        }

        .meta-icon {
          font-size: 16px;
        }

        .refresh-indicator {
          color: #3b82f6;
          font-weight: 500;
        }

        /* Professional Table */
        .table-container {
          background: white;
          border-radius: 12px;
          overflow: hidden;
          box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
          border: 1px solid #e2e8f0;
          margin-bottom: 24px;
        }

        .table-wrapper {
          overflow-x: auto;
        }

        .data-table {
          width: 100%;
          border-collapse: collapse;
          font-size: 14px;
        }

        .table-header {
        background: #4331a8ff;
        }

      .table-header-cell {
  padding: 16px 12px;
  text-align: center;
  font-weight: 600;
  color: #ffffffff;
  font-size: 12px;
  text-transform: uppercase;
  letter-spacing: 0.5px;
  border-width: 1px;
  border-style: solid;
  border-color: black;
  background: blue;
}

        .table-header-content {
          display: flex;
          align-items: center;
          gap: 6px;
        }

        .column-icon {
          font-size: 12px;
          opacity: 0.7;
        }

      .table-cell {
  padding: 16px 12px;
  border-width: 1px;
  border-style: solid;
  border-color: black;
  background: white;
  transition: all 0.2s ease;
  color: #000000ff;
  text-align: center;
}

        .table-row:hover .table-cell {
          background: #f8fafc;
        }

        .table-row:last-child .table-cell {
          border-bottom: none;
        }

        .numeric-cell {
          text-align: right;
          font-family: 'SF Mono', 'Monaco', 'Consolas', monospace;
          font-weight: 500;
        }

        .no-data {
          text-align: center;
          padding: 60px 20px;
          color: #64748b;
        }

        .no-data-content {
          display: flex;
          flex-direction: column;
          align-items: center;
          gap: 16px;
        }

        .no-data-icon {
          font-size: 48px;
          opacity: 0.5;
        }

        .no-data-text {
          font-size: 16px;
          font-weight: 500;
        }

        /* Pending Button */
        .pending-btn {
          background: #fffbeb;
          color: #d97706;
          border: 1px solid #fed7aa;
          padding: 6px 12px;
          border-radius: 6px;
          font-size: 12px;
          font-weight: 500;
          cursor: pointer;
          display: flex;
          align-items: center;
          gap: 6px;
          transition: all 0.2s ease;
        }

        .pending-btn:hover {
          background: #fef3c7;
          transform: translateY(-1px);
        }

        .pending-icon {
          font-size: 14px;
        }

        .no-pending {
          color: #9ca3af;
          font-style: italic;
        }

        /* Remarks Badges */
        .remarks-badges {
          display: flex;
          flex-wrap: wrap;
          gap: 6px;
        }

        .remark-badge {
          padding: 4px 8px;
          border-radius: 6px;
          font-size: 11px;
          font-weight: 600;
          white-space: nowrap;
        }

        .done-badge {
          background: #ecfdf5;
          color: #065f46;
          border: 1px solid #a7f3d0;
        }

        .pending-badge {
          background: #fffbeb;
          color: #92400e;
          border: 1px solid #fed7aa;
        }

        .issue-badge {
          background: #fef2f2;
          color: #991b1b;
          border: 1px solid #fecaca;
        }

        .default-badge {
          background: #f8fafc;
          color: #374151;
          border: 1px solid #e2e8f0;
        }

        /* Professional Pagination */
        .pagination {
          display: flex;
          justify-content: space-between;
          align-items: center;
          background: white;
          padding: 20px 24px;
          border-radius: 12px;
          box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
          border: 1px solid #e2e8f0;
        }

        .pagination-info {
          color: #374151;
          font-size: 14px;
          font-weight: 500;
        }

        .pagination-controls {
          display: flex;
          align-items: center;
          gap: 8px;
        }

        .pagination-btn, .page-btn {
          padding: 8px 12px;
          border: 1px solid #d1d5db;
          background: white;
          color: #374151;
          cursor: pointer;
          border-radius: 6px;
          font-size: 14px;
          font-weight: 500;
          transition: all 0.2s ease;
          min-width: 40px;
          text-align: center;
        }

        .pagination-btn:hover:not(:disabled), .page-btn:hover:not(:disabled) {
          background: #f8fafc;
          border-color: #9ca3af;
        }

        .pagination-btn:disabled, .page-btn:disabled {
          opacity: 0.5;
          cursor: not-allowed;
        }

        .page-numbers {
          display: flex;
          gap: 4px;
        }

        .active-page {
          background: #3b82f6 !important;
          color: white !important;
          border-color: #3b82f6 !important;
        }

        /* Professional Dialog */
        .dialog-backdrop {
          position: fixed;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          background: rgba(0, 0, 0, 0.5);
          display: flex;
          align-items: center;
          justify-content: center;
          z-index: 1000;
          padding: 20px;
        }

        .dialog {
          background: white;
          border-radius: 12px;
          box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04);
          max-width: 500px;
          width: 100%;
          max-height: 80vh;
          overflow: hidden;
          display: flex;
          flex-direction: column;
          border: 1px solid #e2e8f0;
        }

        .dialog-header {
          padding: 20px 24px;
          border-bottom: 1px solid #e2e8f0;
          display: flex;
          justify-content: space-between;
          align-items: center;
          background: #f8fafc;
        }

        .dialog-title {
          margin: 0;
          font-size: 18px;
          font-weight: 600;
          color: #1e293b;
          display: flex;
          align-items: center;
          gap: 12px;
        }

        .dialog-icon {
          font-size: 20px;
          color: #d97706;
        }

        .dialog-close {
          background: none;
          border: none;
          color: #64748b;
          font-size: 24px;
          cursor: pointer;
          padding: 0;
          width: 32px;
          height: 32px;
          display: flex;
          align-items: center;
          justify-content: center;
          border-radius: 6px;
          transition: all 0.2s ease;
        }

        .dialog-close:hover {
          background: #f1f5f9;
          color: #374151;
        }

        .dialog-body {
          padding: 24px;
          flex: 1;
          overflow-y: auto;
        }

        .shades-list {
          display: flex;
          flex-direction: column;
          gap: 8px;
        }

        .shade-item {
          display: flex;
          align-items: center;
          gap: 12px;
          padding: 12px 16px;
          background: #f8fafc;
          border-radius: 8px;
          border: 1px solid #e2e8f0;
          color: #374151;
        }

        .shade-bullet {
          color: #3b82f6;
          font-weight: bold;
          font-size: 16px;
        }

        .no-shades {
          text-align: center;
          padding: 40px 20px;
          color: #64748b;
        }

        .no-shades-icon {
          font-size: 48px;
          margin-bottom: 16px;
          opacity: 0.5;
        }

        .no-shades-text {
          font-size: 16px;
          font-weight: 500;
        }

        .dialog-footer {
          padding: 20px 24px;
          border-top: 1px solid #e2e8f0;
          display: flex;
          justify-content: flex-end;
        }

        .dialog-action-btn {
          background: #3b82f6;
          color: white;
          border: 1px solid #3b82f6;
          padding: 10px 24px;
          border-radius: 6px;
          font-size: 14px;
          font-weight: 500;
          cursor: pointer;
          transition: all 0.2s ease;
        }

        .dialog-action-btn:hover {
          background: #2563eb;
          transform: translateY(-1px);
        }

        /* Responsive Design */
        @media (max-width: 1024px) {
          .main-content {
            padding: 16px;
          }

          .header-content {
            flex-direction: column;
            gap: 24px;
            text-align: center;
          }

          .filter-section {
            grid-template-columns: 1fr;
          }

          .controls-section {
            flex-direction: column;
            align-items: stretch;
          }

          .left-controls {
            justify-content: center;
          }

          .button-group {
            justify-content: center;
          }

          .pagination {
            flex-direction: column;
            gap: 16px;
            text-align: center;
          }

          .pagination-controls {
            flex-wrap: wrap;
            justify-content: center;
          }
        }

        @media (max-width: 768px) {
          .stats-grid {
            flex-direction: column;
            gap: 16px;
          }

          .header {
            padding: 24px;
          }

          .header-title {
            font-size: 24px;
          }

          .toolbar {
            padding: 20px;
          }
        }

        /* Print Styles */
        @media print {
          .cutting-stats-container {
            background: white;
          }

          .toolbar, .pagination, .btn {
            display: none;
          }

          .table-container {
            box-shadow: none;
            border: 1px solid #ddd;
          }
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

              <div className="filter-group">
                <label className="filter-label">Garment Type</label>
                <select
                  className="filter-select"
                  value={garmentFilter}
                  onChange={(e) => setGarmentFilter(e.target.value)}
                  disabled={loading}
                >
                  <option value="">All Garments</option>
                  {distinctGarments.map((garment) => (
                    <option key={garment} value={garment}>
                      {garment}
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
            </div>
          </div>

          <div className="toolbar-section">
            <div className="section-title">Remarks Filter</div>
            <div className="remarks-grid">
              {distinctRemarks.length === 0 ? (
                <span style={{ color: "#64748b", fontSize: "14px" }}>No remarks available</span>
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
                <button type="button" className="btn btn-clear" onClick={clearAllFilters} disabled={loading}>
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
                <button type="button" className="btn btn-pdf" onClick={handleExportPDF} disabled={loading || refreshing}>
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
          </div>
          <div className="meta-right">{refreshing && <div className="refresh-indicator">Updating data...</div>}</div>
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
                        {(lotFilter || partyFilter || garmentFilter || seasonFilter || selectedRemarks.size > 0) && (
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
                        if (h === "Pending Shade") {
                          const lot = String(row["Lot No"] || "").trim();
                          const list = pendingListByLot[lot] || [];
                          const isPending = list.length > 0;

                          return (
                            <td key={h} className="table-cell">
                              {isPending ? (
                                <button type="button" className="pending-btn" onClick={() => openPendingDialog(lot)}>
                                  <span className="pending-icon">⏳</span>
                                  Pending ({list.length})
                                </button>
                              ) : (
                                <span className="no-pending">—</span>
                              )}
                            </td>
                          );
                        }

                        if (h === "Total Qty" || h === "Days after PO issue") {
                          return (
                            <td key={h} className="table-cell numeric-cell">
                              {fmtNum(row[h])}
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
                                      remark.includes("Done")
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

                        return <td key={h} className="table-cell">{row[h]}</td>;
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
            Showing {startIdx + 1} to {Math.min(endIdx, filtered.length)} of {filtered.length} entries
            {filtered.length !== rows.length && ` (filtered from ${rows.length} total entries)`}
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
      </div>
    </div>
  );
}
