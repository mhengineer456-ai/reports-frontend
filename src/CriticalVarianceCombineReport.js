// src/CriticalVarianceCombineReport.js
import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import { Link, useHistory } from "react-router-dom";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { GOOGLE_API_KEY, SPREADSHEET_IDS, SHEET_NAMES, fetchSheetDataFromBackend } from "./config";
import { getCurrentUser, logoutUser } from "./auth";
import { fetchRemarksForTab } from "./embPrintRemarksService";

/**
 * Factory Suite Pro - Critical Red Zone & Variance Combined Report
 * Comprehensive Executive Command Center:
 * - Aggregates RED ZONE & PENDING lots across all 6 Most Critical Factory Stages:
 *   1. ✂️ Cutting Report
 *   2. 🧵 Embroidery Report
 *   3. 🖨️ Printing Report
 *   4. 📋 After EMB/Print Done (Pending Issue to Stitching)
 *   5. 🪡 Overall Stitching Operations
 *   6. 📦⏳ Pending Packing to Issue
 *
 * Full Feature Set:
 * - Direct deep links to jump straight into each stage's dedicated report
 * - Interactive multi-stage tab switching & severity filtering
 * - 3-Sheet Excel Export (.xlsx) via ExcelJS
 * - Direct A3 Landscape PDF Export (.pdf) with Pure Black Text & Executive KPI Summary
 * - Top Brand Navigation Bar with Back Button
 */

// ====== STAGE CONFIGURATION ======
const CRITICAL_STAGES = [
  { id: "all", label: "All Red Zone Bottlenecks", shortLabel: "All Red Zone", icon: "🔴", badgeColor: "#dc2626" },
  { id: "cutting", label: "1. Cutting Red Zone", shortLabel: "1. Cutting", icon: "✂️", badgeColor: "#059669", path: "/cutting-report" },
  { id: "embroidery", label: "2. Embroidery Red Zone", shortLabel: "2. Embroidery", icon: "🧵", badgeColor: "#7c3aed", path: "/embroidery" },
  { id: "printing", label: "3. Printing Red Zone", shortLabel: "3. Printing", icon: "🖨️", badgeColor: "#db2777", path: "/printing" },
  { id: "post_emb_print", label: "4. Post-EMB/Print Red Zone", shortLabel: "4. Post-EMB/Print", icon: "📋", badgeColor: "#ea580c", path: "/pending-issue-to-stitching" },
  { id: "stitching", label: "5. Stitching WIP Red Zone", shortLabel: "5. Stitching WIP", icon: "🪡", badgeColor: "#0284c7", path: "/stitching-complete-lot" },
  { id: "packing_handover", label: "6. Packing Pending to Issue Red Zone", shortLabel: "6. Packing Pending to Issue", icon: "📦⏳", badgeColor: "#c2410c", path: "/pending-packing-issue" },
];

const SEVERITY_CONFIG = {
  CRITICAL: { label: "Critical Red Zone (> 5 Days)", badge: "🔴 CRITICAL RED ZONE", color: "#dc2626", bg: "#fef2f2", border: "#fecaca" },
  HIGH: { label: "Red Zone (3-5 Days)", badge: "🔴 RED ZONE (3-5d)", color: "#b91c1c", bg: "#fff1f2", border: "#fecdd3" },
  MODERATE: { label: "Red Zone (> 2 Days)", badge: "🔴 RED ZONE (>2d)", color: "#991b1b", bg: "#fff5f5", border: "#fed7aa" },
};

// 11 Intermediate Finishing Process Sheets from DAILY_STITCHING spreadsheet
const FINISHING_PROCESS_SHEETS = [
  'Overlock',
  'Filling',
  'Press',
  'FeedUp',
  'Jaybir Printing',
  'Jaybir Embroidery',
  'Washing',
  'Bone',
  'Elastic',
  'Folding',
  'KajButton'
];

/* ---------- Utility Functions ---------- */
const norm = (s) => String(s || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
const clean = (v) => (v == null ? "" : String(v).trim());

const getDirectImageUrl = (url) => {
  if (!url) return "";
  if (url.includes("drive.google.com")) {
    const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (match && match[1]) {
      return `https://lh3.googleusercontent.com/u/0/d/${match[1]}`;
    }
    const idMatch = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (idMatch && idMatch[1]) {
      return `https://lh3.googleusercontent.com/u/0/d/${idMatch[1]}`;
    }
  }
  return url;
};

const titleCase = (s) =>
  String(s || "")
    .split(/\s+/)
    .filter(Boolean)
    .map((w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
    .join(" ");

const parseAnyDate = (dateVal) => {
  if (!dateVal) return null;
  if (dateVal instanceof Date) return isNaN(dateVal.getTime()) ? null : dateVal;

  const str = String(dateVal).trim();
  if (!str || str === "N/A" || str === "—" || str === "undefined" || str === "null") return null;

  // 1. Match DD-Mon-YYYY (e.g. "3 Aug 2026", "03-Aug-2026", "03 Aug 2026", "3-AUG-26")
  const textMatch = str.match(/^(\d{1,2})[-\s/]([a-zA-Z]{3,})[-\s/](\d{2,4})/);
  if (textMatch) {
    const [, dayStr, monStr, yearStr] = textMatch;
    const monthMap = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
    const m = monthMap[monStr.toLowerCase().slice(0, 3)];
    if (m !== undefined) {
      let y = parseInt(yearStr, 10);
      if (y < 100) y += (y < 50 ? 2000 : 1900);
      const dt = new Date(y, m, parseInt(dayStr, 10));
      if (!isNaN(dt.getTime())) return dt;
    }
  }

  // 2. Match "Wed Aug 03 2026" or "Aug 03 2026"
  const textMatch2 = str.match(/([a-zA-Z]{3,})\s+(\d{1,2})\s+(\d{4})/);
  if (textMatch2) {
    const [, monStr, dayStr, yearStr] = textMatch2;
    const monthMap = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
    const m = monthMap[monStr.toLowerCase().slice(0, 3)];
    if (m !== undefined) {
      const dt = new Date(parseInt(yearStr, 10), m, parseInt(dayStr, 10));
      if (!isNaN(dt.getTime())) return dt;
    }
  }

  // 3. Match ISO YYYY-MM-DD or YYYY/MM/DD
  const isoMatch = str.match(/^(\d{4})[-\/\.](\d{1,2})[-\/\.](\d{1,2})/);
  if (isoMatch) {
    const [, yyyy, mm, dd] = isoMatch;
    const dt = new Date(parseInt(yyyy, 10), parseInt(mm, 10) - 1, parseInt(dd, 10));
    if (!isNaN(dt.getTime())) return dt;
  }

  // 4. Match Indian standard DD/MM/YYYY or DD-MM-YYYY (NEVER use default JS new Date which parses as MM/DD/YYYY)
  const slashParts = str.split(/[\/\-\.]/);
  if (slashParts.length >= 3) {
    const p0 = parseInt(slashParts[0], 10);
    const p1 = parseInt(slashParts[1], 10);
    const p2 = parseInt(slashParts[2].split(" ")[0], 10);
    if (!isNaN(p0) && !isNaN(p1) && !isNaN(p2)) {
      const yr = p2 > 100 ? p2 : (p2 < 50 ? 2000 + p2 : 1900 + p2);
      const dt = new Date(yr, p1 - 1, p0);
      if (!isNaN(dt.getTime())) return dt;
    }
  }

  // 5. Fallback standard Date parse
  const d = new Date(str);
  if (!isNaN(d.getTime())) return d;

  return null;
};

/* ---------- Financial Year & Date Helpers ---------- */
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
      d = parseAnyDate(s);
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

const getEmbStatus = (challanHistoryJson) => {
  if (!challanHistoryJson) return { text: "Unknown", type: "neutral" };
  try {
    const history = JSON.parse(challanHistoryJson);
    if (!Array.isArray(history) || history.length === 0) return { text: "Emb Pending", type: "warning" };
    const allCompleted = history.every((e) => e && e.embCompleted === true);
    return allCompleted
      ? { text: "Emb Done", type: "success" }
      : { text: "Emb Pending", type: "warning" };
  } catch {
    return { text: "Emb Pending", type: "warning" };
  }
};

const getPrintingStatus = (challanHistoryJson) => {
  if (!challanHistoryJson) return { text: "Unknown", type: "neutral" };
  try {
    const history = JSON.parse(challanHistoryJson);
    if (!Array.isArray(history) || history.length === 0) return { text: "Printing Pending", type: "warning" };
    const allCompleted = history.every((e) => e && (e.embCompleted === true || e.printCompleted === true));
    return allCompleted
      ? { text: "Printing Done", type: "success" }
      : { text: "Printing Pending", type: "warning" };
  } catch {
    return { text: "Printing Pending", type: "warning" };
  }
};

const parseChallanHistory = (challanHistoryRaw) => {
  if (!challanHistoryRaw) return [];
  if (Array.isArray(challanHistoryRaw)) return challanHistoryRaw;
  try {
    const parsed = JSON.parse(challanHistoryRaw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
};

const getLastEmbDate = (challanHistory) => {
  const list = parseChallanHistory(challanHistory);
  const completed = list.filter((c) => c && (c.embCompleted === true || c.printCompleted === true));
  if (completed.length === 0) return null;
  const dates = completed
    .map((c) => parseAnyDate(c.embUpdatedAt || c.printUpdatedAt || c.receivedDate || c.date))
    .filter((d) => d && !isNaN(d.getTime()));
  if (dates.length === 0) return null;
  return new Date(Math.max(...dates.map((d) => d.getTime())));
};

const isLotCompleted = (completedStatus) => {
  if (!completedStatus) return false;
  if (Array.isArray(completedStatus)) {
    return completedStatus.some((entry) => entry && entry.status && norm(entry.status).includes("complete"));
  }
  const rawStr = String(completedStatus).trim();
  if (!rawStr || rawStr === "-" || rawStr === "[]") return false;
  const n = norm(rawStr);
  if (n.includes("complete") || n.includes("done")) return true;
  if (rawStr.startsWith("[") || rawStr.startsWith("{")) {
    try {
      const parsed = JSON.parse(rawStr);
      const arr = Array.isArray(parsed) ? parsed : [parsed];
      return arr.some((entry) => entry && entry.status && String(entry.status).toLowerCase().includes("complete"));
    } catch {
      return false;
    }
  }
  return false;
};

const calcDaysDiff = (dateObj) => {
  if (!dateObj || isNaN(dateObj.getTime())) return 0;
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  const target = new Date(dateObj);
  target.setHours(0, 0, 0, 0);
  const diffTime = Math.max(0, now.getTime() - target.getTime());
  return Math.floor(diffTime / (1000 * 60 * 60 * 24));
};

const calculatePendingDays = (completedDate) => {
  if (!completedDate || completedDate === "-" || completedDate === "—" || completedDate === "null" || completedDate === "undefined") return 0;
  try {
    const completed = parseAnyDate(completedDate);
    if (!completed || isNaN(completed.getTime())) return 0;

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    completed.setHours(0, 0, 0, 0);

    const diffTime = today - completed;
    const diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24));

    return diffDays > 0 ? diffDays : 0;
  } catch (error) {
    return 0;
  }
};

const extractDateFromValue = (val, fallbackDate) => {
  if (!val && !fallbackDate) return null;
  const str = String(val || "").trim();

  if (str.startsWith("[") || str.startsWith("{")) {
    try {
      const parsed = JSON.parse(str);
      if (Array.isArray(parsed) && parsed.length > 0) {
        const last = parsed[parsed.length - 1];
        if (last) {
          const possible = last.completionDate || last.completedDate || last.timestamp || last.date || last.updatedAt;
          if (possible) {
            const d = parseAnyDate(possible);
            if (d) return { dateObj: d, dateStr: String(possible) };
          }
        }
      } else if (parsed && typeof parsed === "object") {
        const possible = parsed.completionDate || parsed.completedDate || parsed.timestamp || parsed.date;
        if (possible) {
          const d = parseAnyDate(possible);
          if (d) return { dateObj: d, dateStr: String(possible) };
        }
      }
    } catch { }
  }

  if (str && str !== "-" && str.toLowerCase() !== "yes" && str.toLowerCase() !== "completed" && str.toLowerCase() !== "done") {
    const d = parseAnyDate(str);
    if (d) return { dateObj: d, dateStr: str };
  }

  if (fallbackDate) {
    const d = parseAnyDate(fallbackDate);
    if (d) return { dateObj: d, dateStr: String(fallbackDate) };
  }

  return null;
};

const transformIssuesData = (values) => {
  if (!values || values.length === 0) return { issuesData: [], lotMap: new Map() };

  const headers = values[0].map((header) => String(header || "").trim().toLowerCase());
  const rows = values.slice(1);

  const lotNumberIndex = headers.findIndex((h) => h.includes("lot number") || h.includes("lot"));
  const packingSupervisorIndex = headers.findIndex((h) => h.includes("packing supervisor"));
  const packingDateIndex = headers.findIndex((h) => h.includes("packing date"));
  const packingCompleteIndex = headers.findIndex((h) => h.includes("packing complete"));
  const totalPcsIndex = headers.findIndex((h) => h.includes("total pcs") || h.includes("pcs"));

  const issuesData = [];
  const lotMap = new Map();

  rows.forEach((row, index) => {
    if (row[lotNumberIndex]) {
      const lotNumber = String(row[lotNumberIndex] || "").trim();
      if (!lotNumber) return;

      const pDate = packingDateIndex !== -1 ? String(row[packingDateIndex] || "").trim() : "";
      const pcs = totalPcsIndex !== -1 ? parseFloat(String(row[totalPcsIndex] || "0").replace(/,/g, "")) || 0 : 0;
      const sup = packingSupervisorIndex !== -1 ? String(row[packingSupervisorIndex] || "").trim() : "";
      const comp = packingCompleteIndex !== -1 ? String(row[packingCompleteIndex] || "").trim() : "";

      const issueItem = {
        id: `issues-${index}`,
        lot: lotNumber,
        lotNumber: lotNumber,
        packingSupervisor: sup,
        supervisor: sup,
        packingDate: pDate,
        parsedPackingDate: parseAnyDate(pDate),
        packedPcs: pcs,
        packingComplete: comp
      };

      issuesData.push(issueItem);
      lotMap.set(lotNumber, issueItem);
      lotMap.set(lotNumber.toUpperCase(), issueItem);
      lotMap.set(norm(lotNumber), issueItem);
      const cleanKey = lotNumber.toUpperCase().replace(/[^A-Z0-9]/g, "");
      if (cleanKey) lotMap.set(cleanKey, issueItem);
    }
  });

  return { issuesData, lotMap };
};

const transformRawpackData = (values) => {
  if (!values || values.length === 0) return { rawpackData: [], rawpackLotMap: new Map() };

  let headerRowIdx = 0;
  for (let i = 0; i < Math.min(values.length, 5); i++) {
    const rowStr = (values[i] || []).join(" ").toLowerCase();
    if (rowStr.includes("lot no") || rowStr.includes("item") || rowStr.includes("packing person")) {
      headerRowIdx = i;
      break;
    }
  }

  const headers = (values[headerRowIdx] || []).map((h) => (h ? String(h).trim().toLowerCase() : ""));
  const rows = values.slice(headerRowIdx + 1);

  const getCol = (name) => headers.findIndex((h) => h.includes(name));

  const lotCol = getCol("lot no") >= 0 ? getCol("lot no") : getCol("lot");
  const lot2Col = getCol("lot no.2") >= 0 ? getCol("lot no.2") : getCol("lot2");
  const packingPersonCol = getCol("packing person");
  const supervisorCol = getCol("supervisior") >= 0 ? getCol("supervisior") : getCol("supervisor");
  const packingIssueDateCol = getCol("date of packing issue");
  const packingCompleteDateCol = getCol("date of packing complete");
  const reportFindCol = getCol("report find");
  const completedCol =
    headers.findIndex((h) => h && h.trim().toLowerCase() === "completed") >= 0
      ? headers.findIndex((h) => h && h.trim().toLowerCase() === "completed")
      : getCol("completed");

  const rawpackData = [];
  const rawpackLotMap = new Map();

  rows.forEach((row, idx) => {
    const lotVal1 = lotCol >= 0 ? String(row[lotCol] || "").trim() : "";
    const lotVal2 = lot2Col >= 0 ? String(row[lot2Col] || "").trim() : "";
    const lotNumber = lotVal1 || lotVal2;

    if (lotNumber && lotNumber !== "-" && lotNumber !== "0") {
      const packingPersonVal = packingPersonCol >= 0 ? String(row[packingPersonCol] || "").trim() : "";
      const packingIssueDateVal = packingIssueDateCol >= 0 ? String(row[packingIssueDateCol] || "").trim() : "";
      const packingCompleteDateVal = packingCompleteDateCol >= 0 ? String(row[packingCompleteDateCol] || "").trim() : "";
      const reportFindVal = reportFindCol >= 0 ? String(row[reportFindCol] || "").trim().toLowerCase() : "";
      const completedVal = completedCol >= 0 ? String(row[completedCol] || "").trim().toLowerCase() : "";

      const isRawpackCompleted =
        Boolean(packingCompleteDateVal && packingCompleteDateVal !== "-" && packingCompleteDateVal !== "#N/A" && packingCompleteDateVal !== "00/01/00") ||
        completedVal === "yes" ||
        completedVal === "complete" ||
        completedVal === "completed" ||
        reportFindVal === "complete" ||
        reportFindVal === "completed" ||
        reportFindVal === "yes";

      const isRawpackIssued = Boolean(packingPersonVal) || Boolean(packingIssueDateVal && packingIssueDateVal !== "-" && packingIssueDateVal !== "#N/A");

      const item = {
        id: `rawpack-${idx}`,
        lot: lotNumber,
        lotNumber: lotNumber,
        packingPerson: packingPersonVal,
        supervisor: supervisorCol >= 0 ? row[supervisorCol] || "" : "",
        packingIssueDate: packingIssueDateVal,
        packingCompleteDate: packingCompleteDateVal,
        completedColValue: completedVal,
        isCompleted: isRawpackCompleted,
        isIssued: isRawpackIssued,
        status: isRawpackCompleted ? "complete" : ""
      };

      rawpackData.push(item);
      rawpackLotMap.set(lotNumber, item);
      rawpackLotMap.set(lotNumber.toUpperCase(), item);
      rawpackLotMap.set(norm(lotNumber), item);
      const cleanKey = lotNumber.toUpperCase().replace(/[^A-Z0-9]/g, "");
      if (cleanKey) rawpackLotMap.set(cleanKey, item);
    }
  });

  return { rawpackData, rawpackLotMap };
};

const transformProcessSheetsData = (sheetResults, sheetNames) => {
  const runningLotsMap = new Map();
  const runningLotsSet = new Set();
  const lastProcessCompletedMap = new Map();

  sheetResults.forEach((sheetData, sIdx) => {
    const sheetName = sheetNames[sIdx];
    const values = sheetData?.values || [];
    if (!values || values.length === 0) return;

    const headers = (values[0] || []).map((h) => String(h || "").trim().toLowerCase());
    const normK = (s) => String(s || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
    const findCol = (kws) => headers.findIndex((h) => kws.some((k) => normK(h).includes(normK(k))));

    const lotIdx = findCol(["lot number", "lot no", "lot #", "lot"]);
    const compIdx = findCol(["complete", "completed", "completion date", "completion", "complete status", "status"]);
    const wipIdx = findCol(["wip", "remarks"]);
    const dateIdx = findCol(["date", "issue date", "timestamp"]);

    const actualLotIdx = lotIdx !== -1 ? lotIdx : 1;

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (!row || row.length === 0) continue;

      const rawLot = String(row[actualLotIdx] || row[0] || "").trim();
      if (!rawLot || rawLot === "-" || rawLot === "0") continue;

      const cleanLot = rawLot.toUpperCase();
      const cleanKey = cleanLot.replace(/[^A-Z0-9]/g, "");
      const rawComp = compIdx !== -1 && row[compIdx] ? String(row[compIdx]).trim() : "";
      const rawWip = wipIdx !== -1 && row[wipIdx] ? String(row[wipIdx]).trim() : "";
      const rawDate = dateIdx !== -1 && row[dateIdx] ? String(row[dateIdx]).trim() : "";

      // Determine if this process entry is completed
      let isCompleted = false;
      let completionDateResult = null;

      if (rawComp && rawComp !== "[]" && rawComp !== "-" && rawComp.toLowerCase() !== "null" && rawComp.toLowerCase() !== "undefined") {
        const compLower = rawComp.toLowerCase();
        if (rawComp.startsWith("[") || rawComp.startsWith("{")) {
          try {
            const parsed = JSON.parse(rawComp);
            if (Array.isArray(parsed) && parsed.length > 0) {
              const last = parsed[parsed.length - 1];
              if (last && (last.timestamp || last.date || last.status || last.completionDate || last.completedDate)) {
                isCompleted = true;
                completionDateResult = extractDateFromValue(rawComp, rawDate);
              }
            } else if (parsed && typeof parsed === "object" && Object.keys(parsed).length > 0) {
              isCompleted = true;
              completionDateResult = extractDateFromValue(rawComp, rawDate);
            }
          } catch (e) {
            isCompleted = true;
            completionDateResult = extractDateFromValue(rawComp, rawDate);
          }
        } else if (!compLower.includes("pending") && !compLower.includes("not") && compLower !== "no") {
          isCompleted = true;
          completionDateResult = extractDateFromValue(rawComp, rawDate);
        }
      }

      if (isCompleted) {
        const completedDateObj = completionDateResult?.dateObj || parseAnyDate(rawDate);
        if (completedDateObj && !isNaN(completedDateObj.getTime())) {
          const keysToSet = [cleanLot, rawLot, cleanKey, norm(rawLot)].filter(Boolean);
          keysToSet.forEach((key) => {
            const existing = lastProcessCompletedMap.get(key);
            if (!existing || completedDateObj.getTime() > existing.latestCompletedDate.getTime()) {
              lastProcessCompletedMap.set(key, {
                latestCompletedDate: completedDateObj,
                latestCompletedDateStr: completionDateResult?.dateStr || rawDate || formatDisplayDate(completedDateObj),
                sheetName: sheetName
              });
            }
          });
        }
      } else {
        // If not completed, this lot has an active/running process in this sheet
        runningLotsSet.add(cleanLot);
        runningLotsSet.add(rawLot);
        runningLotsSet.add(norm(rawLot));
        if (cleanKey) runningLotsSet.add(cleanKey);

        const existing = runningLotsMap.get(cleanLot) || runningLotsMap.get(rawLot) || [];
        const formattedIssueDate = rawDate;
        const alreadyAdded = existing.some((e) => e.sheetName === sheetName && e.issueDate === formattedIssueDate);
        if (!alreadyAdded) {
          existing.push({
            sheetName,
            issueDate: formattedIssueDate,
            wipRemark: rawWip
          });
        }
        runningLotsMap.set(cleanLot, existing);
        runningLotsMap.set(rawLot, existing);
        runningLotsMap.set(norm(rawLot), existing);
        if (cleanKey) runningLotsMap.set(cleanKey, existing);
      }
    }
  });

  return { runningLotsMap, runningLotsSet, lastProcessCompletedMap };
};

const formatDisplayDate = (d) => {
  if (!d) return "—";
  const parsed = d instanceof Date ? d : parseAnyDate(d);
  if (!parsed || isNaN(parsed.getTime())) return String(d);
  return parsed.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
};

export const formatLatestRemark = (rawVal, fallback = "—") => {
  if (rawVal === null || rawVal === undefined) return fallback;

  let val = rawVal;

  // 1. If it's a string, clean and test if it's stringified JSON
  if (typeof val === "string") {
    val = val.trim();
    if (!val || val === "—" || val === "N/A" || val === "-" || val === "null" || val === "undefined") {
      return fallback;
    }

    // Try parsing if it looks like JSON array or object
    if ((val.startsWith("[") && val.endsWith("]")) || (val.startsWith("{") && val.endsWith("}"))) {
      try {
        val = JSON.parse(val);
      } catch {
        try {
          val = JSON.parse(val.replace(/\\"/g, '"'));
        } catch {
          const matchRemark = val.match(/"(?:remarks|text|status|updateType)"\s*:\s*"([^"]+)"/i);
          if (matchRemark && matchRemark[1]) {
            return formatLatestRemark(matchRemark[1], fallback);
          }
        }
      }
    }
  }

  // 2. If it's an Array of items/updates/remarks
  if (Array.isArray(val)) {
    if (val.length === 0) return fallback;

    const validItems = val.filter((item) => item !== null && item !== undefined && item !== "");
    if (validItems.length === 0) return fallback;

    const hasTimestamp = validItems.some((item) => item && typeof item === "object" && item.timestamp);
    let sorted = validItems;
    if (hasTimestamp) {
      sorted = [...validItems].sort((a, b) => {
        const timeA = a?.timestamp ? new Date(a.timestamp).getTime() : 0;
        const timeB = b?.timestamp ? new Date(b.timestamp).getTime() : 0;
        if (!isNaN(timeA) && !isNaN(timeB) && (timeA > 0 || timeB > 0)) {
          return timeB - timeA;
        }
        return 0;
      });
    }

    const latestItem = (hasTimestamp && sorted[0]?.timestamp && !isNaN(new Date(sorted[0].timestamp).getTime()))
      ? sorted[0]
      : validItems[validItems.length - 1];

    return formatLatestRemark(latestItem, fallback);
  }

  // 3. If it's an Object (single update or remark object)
  if (val && typeof val === "object") {
    const rem = (val.remarks ?? val.remark ?? val.userRemarks ?? val.text ?? val.note ?? val.comment ?? "").toString().trim();
    const st = (val.status ?? val.wipStatus ?? val.updateType ?? val.compStatus ?? val.state ?? "").toString().trim();

    if (rem && rem !== "—" && rem !== "N/A" && rem !== "-") {
      return formatLatestRemark(rem, fallback);
    }
    if (st && st !== "—" && st !== "N/A" && st !== "-") {
      return formatLatestRemark(st, fallback);
    }
    for (const v of Object.values(val)) {
      if (typeof v === "string" && v.trim() && v !== "—" && v !== "N/A" && v !== "-") {
        return formatLatestRemark(v, fallback);
      }
    }
    return fallback;
  }

  // 4. Plain string value
  let finalStr = String(val).trim();
  if (!finalStr || finalStr === "—" || finalStr === "N/A" || finalStr === "-" || finalStr === "null" || finalStr === "undefined") {
    return fallback;
  }
  if (finalStr.includes("_")) {
    finalStr = finalStr.replace(/[_]+/g, " ").replace(/\s+/g, " ").trim();
    finalStr = finalStr
      .split(" ")
      .map((w) => {
        if (!w) return "";
        if (w.startsWith("(")) {
          return "(" + w.slice(1, 2).toUpperCase() + w.slice(2);
        }
        return w.charAt(0).toUpperCase() + w.slice(1);
      })
      .join(" ");
  }
  return finalStr;
};

const cleanRemarkUnderscores = (val) => {
  return formatLatestRemark(val, "—");
};

const getSectionFromMwk = (mwk) => {
  if (!mwk) return "";
  const m = String(mwk).trim().toUpperCase();
  if (m.startsWith("K") || m === "KIDS") return "KIDS";
  if (m.startsWith("M") || m === "MEN" || m === "MAN") return "GENTS";
  if (m.startsWith("W") || m === "WOMEN" || m === "LADIES") return "WOMEN";
  if (m.startsWith("G") || m === "GIRLS") return "GIRLS";
  if (m.startsWith("B") || m === "BOYS") return "BOYS";
  return m;
};

const abbreviateMWK = (mwkValue) => {
  if (!mwkValue || typeof mwkValue !== "string") return "—";
  const value = mwkValue.trim().toLowerCase();
  if (value.includes("gents") || value === "m" || value === "mens") return "M";
  if (value.includes("kids") || value === "k") return "K";
  if (value.includes("girls") || value === "g" || value.includes("girlish")) return "G";
  if (value.includes("women") || value.includes("womens") || value === "w") return "W";
  if (value.includes("boys") || value === "b") return "B";
  return value.charAt(0).toUpperCase();
};

const formatDateToDDMMYY = (dateString) => {
  if (!dateString || (typeof dateString !== "string" && typeof dateString !== "number")) return "—";
  try {
    const cleanStr = String(dateString).trim().replace(/^['"\s]+|['"\s]+$/g, "");
    if (!cleanStr || cleanStr === "-" || cleanStr === "N/A" || cleanStr === "—") return "—";

    const isoMatch = cleanStr.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
    if (isoMatch) {
      const y = isoMatch[1].slice(-2);
      const m = isoMatch[2].padStart(2, "0");
      const d = isoMatch[3].padStart(2, "0");
      return `${d}/${m}/${y}`;
    }

    const parts = cleanStr.split(/[\/\-\.]/);
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

    const date = new Date(cleanStr);
    if (!isNaN(date.getTime())) {
      const day = String(date.getDate()).padStart(2, "0");
      const month = String(date.getMonth() + 1).padStart(2, "0");
      const year = String(date.getFullYear()).slice(-2);
      return `${day}/${month}/${year}`;
    }
    return cleanStr;
  } catch {
    return String(dateString);
  }
};

const getEmbPrintDate = (challanHistory) => {
  if (!challanHistory || typeof challanHistory !== "string" || challanHistory.trim() === "") {
    return "—";
  }
  const trimmed = challanHistory.trim();
  try {
    if (trimmed.startsWith("[") && trimmed.endsWith("]")) {
      const historyArray = JSON.parse(trimmed);
      if (!Array.isArray(historyArray) || historyArray.length === 0) return "—";
      let latestDate = null;
      historyArray.forEach((entry) => {
        if (entry && (entry.embUpdatedAt || entry.printUpdatedAt || entry.date)) {
          const dateStr = entry.embUpdatedAt || entry.printUpdatedAt || entry.date;
          const currentDate = new Date(dateStr);
          if (!isNaN(currentDate.getTime())) {
            if (!latestDate || currentDate > latestDate) {
              latestDate = currentDate;
            }
          }
        }
      });
      if (latestDate) {
        return latestDate.toLocaleDateString("en-GB", {
          day: "2-digit",
          month: "2-digit",
          year: "2-digit",
        });
      }
    }
    return "—";
  } catch {
    return "—";
  }
};

const getLatestWipRemarks = (wipStatus, isCompleted = false) => {
  if (isCompleted) return "Done";
  return formatLatestRemark(wipStatus, "WIP");
};

const calculateStitchingDays = (dateOfIssue, completedStatus, isCompleted = false) => {
  if (!dateOfIssue || typeof dateOfIssue !== "string" || dateOfIssue.trim() === "" || dateOfIssue === "-" || dateOfIssue === "—" || dateOfIssue === "N/A") return 0;
  try {
    let issueDate = null;
    let testDate = new Date(dateOfIssue);
    if (!isNaN(testDate.getTime())) {
      issueDate = testDate;
    } else {
      const parts = dateOfIssue.split(/[\/\-]/);
      if (parts.length === 3) {
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10);
        const year = parseInt(parts[2], 10);
        const fullYear = year < 100 ? 2000 + year : year;
        testDate = new Date(fullYear, month - 1, day);
        if (!isNaN(testDate.getTime())) {
          issueDate = testDate;
        }
      }
    }

    if (!issueDate) return 0;

    let endDate = new Date();
    if (isCompleted && completedStatus) {
      try {
        if (typeof completedStatus === "string" && completedStatus.startsWith("[")) {
          const statusArray = JSON.parse(completedStatus);
          if (Array.isArray(statusArray) && statusArray.length > 0) {
            const completeEntry = statusArray.find((entry) => entry.status && entry.status.toLowerCase().includes("complete"));
            if (completeEntry && completeEntry.timestamp) {
              endDate = new Date(completeEntry.timestamp);
            }
          }
        }
        if (!endDate || isNaN(endDate.getTime())) {
          endDate = new Date();
        }
      } catch {
        endDate = new Date();
      }
    }

    const timeDiff = endDate.getTime() - issueDate.getTime();
    const days = Math.floor(timeDiff / (1000 * 3600 * 24));
    return Math.max(0, days);
  } catch {
    return 0;
  }
};

const extractLatestRemark = (remarksMap, lotStr, fallback = "") => {
  if (!lotStr) return formatLatestRemark(fallback, "—");
  const rawLot = String(lotStr).trim();
  const cleanLot = rawLot.replace(/★\s*/, "").trim();
  const nClean = norm(cleanLot);

  let rawFound = null;

  if (remarksMap && typeof remarksMap === "object") {
    if (remarksMap instanceof Map) {
      rawFound = remarksMap.get(cleanLot) || remarksMap.get(rawLot) || remarksMap.get(nClean);
      if (!rawFound) {
        for (const [k, v] of remarksMap.entries()) {
          if (norm(k) === nClean) {
            rawFound = v;
            break;
          }
        }
      }
    } else {
      rawFound = remarksMap[cleanLot] || remarksMap[rawLot] || remarksMap[nClean];
      if (!rawFound) {
        for (const [k, v] of Object.entries(remarksMap)) {
          if (norm(k) === nClean) {
            rawFound = v;
            break;
          }
        }
      }
    }
  }

  if (rawFound !== null && rawFound !== undefined) {
    const formatted = formatLatestRemark(rawFound, "");
    if (formatted && formatted !== "—" && formatted !== "N/A" && formatted !== "-") {
      return formatted;
    }
  }

  return formatLatestRemark(fallback, "—");
};

/* ---------- Multi-Select Dropdown Component ---------- */
const MultiSelectDropdown = ({ label, options, selectedValues, onChange, placeholder = "All" }) => {
  const [isOpen, setIsOpen] = useState(false);
  const [search, setSearch] = useState("");
  const dropdownRef = useRef(null);

  useEffect(() => {
    const handleClickOutside = (e) => {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target)) {
        setIsOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  const filteredOptions = useMemo(() => {
    if (!search.trim()) return options;
    return options.filter((opt) => opt.toLowerCase().includes(search.toLowerCase().trim()));
  }, [options, search]);

  const toggleOption = (val) => {
    if (selectedValues.includes(val)) {
      onChange(selectedValues.filter((v) => v !== val));
    } else {
      onChange([...selectedValues, val]);
    }
  };

  const displayText = useMemo(() => {
    if (selectedValues.length === 0) return placeholder;
    if (selectedValues.length === 1) return selectedValues[0];
    return `${selectedValues.length} Selected`;
  }, [selectedValues, placeholder]);

  return (
    <div ref={dropdownRef} style={{ position: "relative", minWidth: "150px", flex: "1 1 150px" }}>
      <label style={{ display: "block", fontSize: "0.72rem", fontWeight: 700, color: "#475569", marginBottom: "4px", textTransform: "uppercase", letterSpacing: "0.03em" }}>
        {label}
      </label>
      <div
        onClick={() => setIsOpen(!isOpen)}
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          padding: "8px 12px",
          background: selectedValues.length > 0 ? "#eff6ff" : "#ffffff",
          border: `1.5px solid ${selectedValues.length > 0 ? "#3b82f6" : "#cbd5e1"}`,
          borderRadius: "8px",
          cursor: "pointer",
          fontSize: "0.82rem",
          fontWeight: selectedValues.length > 0 ? 700 : 500,
          color: selectedValues.length > 0 ? "#1d4ed8" : "#334155",
          userSelect: "none",
          transition: "all 0.15s ease"
        }}
      >
        <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{displayText}</span>
        <span style={{ fontSize: "0.7rem", transform: isOpen ? "rotate(180deg)" : "rotate(0deg)", transition: "transform 0.2s" }}>▼</span>
      </div>

      {isOpen && (
        <div
          style={{
            position: "absolute",
            top: "calc(100% + 4px)",
            left: 0,
            right: 0,
            zIndex: 1000,
            background: "#ffffff",
            border: "1.5px solid #cbd5e1",
            borderRadius: "10px",
            boxShadow: "0 12px 28px rgba(0, 0, 0, 0.15)",
            maxHeight: "260px",
            overflowY: "auto",
            padding: "8px"
          }}
        >
          <div style={{ display: "flex", gap: "6px", marginBottom: "8px" }}>
            <input
              type="text"
              placeholder="Search..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              onClick={(e) => e.stopPropagation()}
              style={{
                width: "100%",
                padding: "6px 8px",
                fontSize: "0.78rem",
                borderRadius: "6px",
                border: "1px solid #cbd5e1",
                outline: "none"
              }}
            />
            {selectedValues.length > 0 && (
              <button
                type="button"
                onClick={(e) => {
                  e.stopPropagation();
                  onChange([]);
                }}
                style={{
                  background: "#fee2e2",
                  color: "#dc2626",
                  border: "none",
                  borderRadius: "6px",
                  padding: "4px 8px",
                  fontSize: "0.72rem",
                  fontWeight: 700,
                  cursor: "pointer",
                  whiteSpace: "nowrap"
                }}
              >
                Clear
              </button>
            )}
          </div>

          <div style={{ maxHeight: "180px", overflowY: "auto" }}>
            {filteredOptions.length === 0 ? (
              <div style={{ padding: "8px", fontSize: "0.78rem", color: "#94a3b8", textAlign: "center" }}>No options found</div>
            ) : (
              filteredOptions.map((opt) => {
                const isChecked = selectedValues.includes(opt);
                return (
                  <label
                    key={opt}
                    onClick={(e) => e.stopPropagation()}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: "8px",
                      padding: "6px 8px",
                      borderRadius: "6px",
                      cursor: "pointer",
                      fontSize: "0.8rem",
                      fontWeight: isChecked ? 700 : 500,
                      color: isChecked ? "#1d4ed8" : "#334155",
                      background: isChecked ? "#eff6ff" : "transparent"
                    }}
                  >
                    <input
                      type="checkbox"
                      checked={isChecked}
                      onChange={() => toggleOption(opt)}
                      style={{ cursor: "pointer" }}
                    />
                    <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{opt}</span>
                  </label>
                );
              })
            )}
          </div>
        </div>
      )}
    </div>
  );
};

/* ========================================================================= */
/* MAIN COMPONENT                                                            */
/* ========================================================================= */
export default function CriticalVarianceCombineReport() {
  const history = useHistory();
  const currentUser = getCurrentUser();

  const handleLogout = () => {
    logoutUser();
    history.push("/");
  };

  const handleGoBack = () => {
    try {
      if (window.history.length > 1) {
        history.goBack();
        return;
      }
    } catch { }
    history.push("/dashboard");
  };

  // State
  const [allBottleneckRows, setAllBottleneckRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [refreshing, setRefreshing] = useState(false);
  const [loadingProgress, setLoadingProgress] = useState(0);
  const [loadingMessage, setLoadingMessage] = useState("Initializing Critical Variance Engine...");
  const [error, setError] = useState("");
  const [lastUpdated, setLastUpdated] = useState(null);

  // Active Tab / Stage Selector
  const [activeStageTab, setActiveStageTab] = useState("all");

  // Filter States
  const [financialYearFilter, setFinancialYearFilter] = useState("ALL");
  const [filterSeverity, setFilterSeverity] = useState("all");
  const [filterGarments, setFilterGarments] = useState([]);
  const [filterBrands, setFilterBrands] = useState([]);
  const [filterFabrics, setFilterFabrics] = useState([]);
  const [filterParties, setFilterParties] = useState([]);
  const [filterSections, setFilterSections] = useState([]);
  const [filterSeasons, setFilterSeasons] = useState([]);
  const [searchTerm, setSearchTerm] = useState("");
  const [minVariancePcs, setMinVariancePcs] = useState("");
  const [minAgingDays, setMinAgingDays] = useState("");

  // Pagination & Sorting
  const [sortConfig, setSortConfig] = useState({ key: "agingDays", direction: "desc" });
  const [pageSize, setPageSize] = useState(25);
  const [currentPage, setCurrentPage] = useState(1);
  const [viewImageSrc, setViewImageSrc] = useState(null);

  const abortRef = useRef(null);

  /* ---------- MASTER DATA FETCHER & PIPELINE BOTTLENECK AGGREGATOR ---------- */
  const loadData = async (mode = "initial") => {
    if (mode === "initial") {
      setLoading(true);
      setLoadingProgress(0);
      setLoadingMessage("Connecting to factory production sheets...");
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
      setLoadingMessage("Fetching Job Orders, Index, Issues, RAWPACK, Working Updates, and Remarks data...");
      setLoadingProgress(15);

      const [
        jobRes,
        idxRes,
        issuesRes,
        rawPackRes,
        workingUpdatesRes,
        cuttingRemarksMap,
        embRemarksMap,
        printRemarksMap,
        pendingStitchingRemarksMap,
        packingRemarksMap,
        ...processSheetResults
      ] = await Promise.all([
        fetchSheetDataFromBackend(SPREADSHEET_IDS.JOBORDER, "JobOrder!A1:AZ20000"),
        fetchSheetDataFromBackend(SPREADSHEET_IDS.MAIN, "Index!A:AG"),
        fetchSheetDataFromBackend(SPREADSHEET_IDS.ISSUES, "Issues!A:Z"),
        fetchSheetDataFromBackend(SPREADSHEET_IDS.RAWPACK || "1xD8Uy1lUgvNTQ2RGRBI4ZjOrozbinUPRq2_UfIplP98", "RAWPACK!A:AZ"),
        fetchSheetDataFromBackend(SPREADSHEET_IDS.WORKING_UPDATES || "1Nh7XYE_MnAxtaTRUUntHvBpzctODwjnkbBYDiYLQgoc", "Working Updates!A1:ZZ20000").catch(() => null),
        fetchRemarksForTab("CUTTING").catch(() => ({})),
        fetchRemarksForTab("EMB").catch(() => ({})),
        fetchRemarksForTab("PRINT").catch(() => ({})),
        fetchRemarksForTab("PENDING_STITCHING").catch(() => ({})),
        fetchRemarksForTab("PACKING").catch(() => ({})),
        ...FINISHING_PROCESS_SHEETS.map((sheetName) =>
          fetchSheetDataFromBackend(SPREADSHEET_IDS.DAILY_STITCHING, `'${sheetName}'!A:Z`).catch(() => ({ values: [] }))
        )
      ]);

      setLoadingProgress(40);
      setLoadingMessage("Parsing Job Order, Challan, Working Updates, and Remarks records...");

      // 0. Process Working Updates (Pintu & EA Status)
      const workingUpdatesMap = new Map();
      const wuRows = workingUpdatesRes?.values || [];
      if (wuRows.length > 1) {
        wuRows.slice(1).forEach((row) => {
          const lotNumber = clean(row[0]);
          const updateHistory = clean(row[4] || "[]");
          if (!lotNumber) return;

          let pintuStatus = "—";
          let eaStatus = "—";

          try {
            let history = [];
            if (typeof updateHistory === "string" && (updateHistory.startsWith("[") || updateHistory.startsWith("{"))) {
              try {
                history = JSON.parse(updateHistory);
              } catch {
                history = [];
              }
            } else if (Array.isArray(updateHistory)) {
              history = updateHistory;
            }

            if (Array.isArray(history) && history.length > 0) {
              const pintuUpdates = history.filter((h) => h && h.updatedBy?.toLowerCase() === "pintu");
              if (pintuUpdates.length > 0) {
                const latestPintu = pintuUpdates[pintuUpdates.length - 1];
                pintuStatus = formatLatestRemark(latestPintu.remarks || latestPintu.updateType || latestPintu.status || latestPintu, "Updated");
              }

              const eaUpdates = history.filter((h) => h && (h.updatedBy?.toLowerCase() === "ea" || h.updatedBy?.toLowerCase() === "wa"));
              if (eaUpdates.length > 0) {
                const latestEA = eaUpdates[eaUpdates.length - 1];
                eaStatus = formatLatestRemark(latestEA.remarks || latestEA.updateType || latestEA.status || latestEA, "Updated");
              }

              // Fallback: If neither Pintu nor EA was matched by tag, but history has status updates
              if (pintuStatus === "—" && eaStatus === "—") {
                const generalLatest = formatLatestRemark(history, "—");
                if (generalLatest && generalLatest !== "—") {
                  pintuStatus = generalLatest;
                }
              }
            } else if (typeof updateHistory === "string" && updateHistory && updateHistory !== "[]") {
              pintuStatus = formatLatestRemark(updateHistory, "—");
            }
          } catch { }

          workingUpdatesMap.set(norm(lotNumber), {
            pintuStatus: formatLatestRemark(pintuStatus, "—"),
            eaStatus: formatLatestRemark(eaStatus, "—")
          });
        });
      }

      // 1. Process JobOrder Master Map & Dedicated Challans
      const jobMasterMap = new Map();
      const embChallanRows = [];
      const printChallanRows = [];
      const jobValues = jobRes?.values || [];

      if (jobValues.length > 0) {
        const jHeaders = (jobValues[0] || []).map((h) => String(h || "").trim());
        const getJ = (row, key) => {
          const exactIdx = jHeaders.findIndex((h) => norm(h) === norm(key));
          if (exactIdx !== -1) return clean(row[exactIdx]);
          const partialIdx = jHeaders.findIndex((h) => norm(h).includes(norm(key)));
          return partialIdx !== -1 ? clean(row[partialIdx]) : "";
        };

        for (let i = 1; i < jobValues.length; i++) {
          const row = jobValues[i] || [];
          const lot = getJ(row, "Lot Number") || getJ(row, "Lot No") || getJ(row, "Lot");
          if (!lot) continue;

          const challanNo = getJ(row, "Challan No") || getJ(row, "Challan Number");
          const challanDateStr = getJ(row, "Challan Date");
          const poDateStr = getJ(row, "Date") || challanDateStr;
          const challanQty = parseFloat(String(getJ(row, "Challan Total Qty") || getJ(row, "Challan Qty") || "0").replace(/,/g, "")) || 0;
          const poQty = parseFloat(String(getJ(row, "Quantity") || getJ(row, "Total Qty") || "0").replace(/,/g, "")) || 0;
          const challanHistoryJson = getJ(row, "Challan History JSON") || getJ(row, "Challan History");
          const garmentType = getJ(row, "Garment Type") || getJ(row, "Garment");
          const style = getJ(row, "Style");
          const fabric = getJ(row, "Fabric");
          const brand = getJ(row, "Brand");
          const section = getJ(row, "Section") || getJ(row, "M/W/K");
          const season = getJ(row, "Season");
          const partyName = getJ(row, "Party Name") || getJ(row, "Party");
          const directStitching = getJ(row, "Direct Stitching");
          const status = getJ(row, "Status");
          const priority = getJ(row, "Priority");
          const embVal = getJ(row, "Emb") || getJ(row, "Embroidery");
          const printVal = getJ(row, "Printing") || getJ(row, "Print");
          const jobOrderNo = getJ(row, "Job Order No") || getJ(row, "Order No") || getJ(row, "Job Order") || getJ(row, "JobOrder No") || getJ(row, "JO No");
          const remarks = getJ(row, "Remarks");
          const pendingShade = getJ(row, "Pending Challan Shade") || getJ(row, "Pending Shade") || getJ(row, "Pending");
          const cuttingDate = getJ(row, "Cutting Date");
          const cuttingScanned = getJ(row, "Cutting Scanned");
          const userRemarks = getJ(row, "User Remarks") || getJ(row, "💬 User Remarks") || getJ(row, "HOD Remarks");

          const parsedChallanDate = parseAnyDate(challanDateStr);
          const parsedJobDate = parseAnyDate(poDateStr);

          const entry = {
            lot,
            challanNo,
            challanDateStr,
            parsedChallanDate,
            qty: challanQty || poQty,
            challanQty,
            poQty,
            challanHistoryJson,
            garmentType,
            style,
            fabric,
            brand,
            section,
            season,
            partyName,
            directStitching,
            status,
            priority,
            emb: embVal,
            printing: printVal,
            jobOrderNo,
            remarks,
            pendingShade,
            cuttingDate,
            cuttingScanned,
            userRemarks,
            parsedJobDate,
            jobDate: poDateStr,
            quantity: poQty || challanQty
          };

          // Check if this is an Embroidery Challan (CH-EMB-)
          if (challanNo.trim().toUpperCase().startsWith("CH-EMB-")) {
            embChallanRows.push(entry);
          }
          // Check if this is a Printing Challan (CH-PRINT- or CH-PRN-)
          else if (challanNo.trim().toUpperCase().startsWith("CH-PRINT-") || challanNo.trim().toUpperCase().startsWith("CH-PRN-")) {
            printChallanRows.push(entry);
          }

          // Always register in master map if not yet set, or if this is base PO row
          if (!jobMasterMap.has(norm(lot)) || !challanNo) {
            jobMasterMap.set(norm(lot), entry);
          }
        }
      }

      setLoadingProgress(60);
      setLoadingMessage("Auditing Index Cutting & Stitching milestones...");

      // 2. Process Index Sheet Rows & Map
      const idxMap = new Map();
      const indexLotList = [];
      const idxValues = idxRes?.values || [];

      if (idxValues.length > 0) {
        const iHeaders = (idxValues[0] || []).map((h) => String(h || "").trim());

        const columnDefinitions = [
          { keys: ["Lot Number", "LotNumber", "lotNumber", "lot no", "lot", "lot#"], target: "lot" },
          { keys: ["Saved At", "Cutting Date", "Date", "SavedAt", "cut date", "cuttingdate"], target: "savedAt" },
          { keys: ["Date of Issue", "DateOfIssue", "Date of issue", "date of issue", "Issue Date", "date of issue to stitching"], target: "dateOfIssue" },
          { keys: ["Supervisor", "supervisor", "Supervisor Name", "stitching supervisor"], target: "supervisor" },
          { keys: ["DIRECT STITCHING", "Direct Stitching", "directstitching", "direct"], target: "directStitching" },
          { keys: ["CHALLAN HISTORY", "Challan History", "challanhistory", "challan history json"], target: "challanHistory" },
          { keys: ["Cutting Qty", "CuttingQty", "Total Pcs", "TotalPcs", "total pcs", "Total Plies", "Total"], target: "cuttingQty" },
          { keys: ["Stitching Issue Qty", "stitching issue qty", "StitchingQty", "stitching qty", "Stitch Qty"], target: "stitchingIssueQty" },
          { keys: ["WIP Status", "WIPStatus", "wip status", "wipstatus"], target: "wipStatus" },
          { keys: ["Completed Status", "CompletedStatus", "completed status", "comp status", "status"], target: "completedStatus" },
          { keys: ["Image URL", "ImageUrl", "image url", "Image", "image", "photo", "img"], target: "imageUrl" },
          { keys: ["Fabric", "FABRIC", "fabric", "material"], target: "fabric" },
          { keys: ["Garment Type", "GarmentType", "garment type", "garment", "item"], target: "garmentType" },
          { keys: ["Style", "style", "styledesc"], target: "style" },
          { keys: ["BRAND", "Brand", "brand", "brandname"], target: "brand" },
          { keys: ["M/W/K", "MWK", "m/w/k", "mwk", "m w k"], target: "mwk" },
          { keys: ["SECTION", "Section", "section", "sec"], target: "section" },
          { keys: ["SEASON", "Season", "season", "seasontype"], target: "season" },
          { keys: ["PARTY NAME", "Party Name", "PartyName", "party name", "party", "vendor"], target: "partyName" },
          { keys: ["Priority", "priority", "prioirty", "special"], target: "priority" }
        ];

        const colMap = {};
        columnDefinitions.forEach((def) => {
          for (const key of def.keys) {
            const idx = iHeaders.findIndex((h) => h && norm(h) === norm(key));
            if (idx !== -1) {
              colMap[def.target] = idx;
              break;
            }
          }
        });

        for (let i = 1; i < idxValues.length; i++) {
          const row = idxValues[i] || [];
          const lot = clean(colMap.lot !== undefined ? row[colMap.lot] : "");
          if (!lot) continue;

          const savedAt = clean(colMap.savedAt !== undefined ? row[colMap.savedAt] : "");
          const parsedCutDate = parseAnyDate(savedAt);

          const rawIssueDate = clean(colMap.dateOfIssue !== undefined ? row[colMap.dateOfIssue] : "");
          const validIssueDate = rawIssueDate && rawIssueDate !== "-" && rawIssueDate !== "—" && rawIssueDate !== "N/A" && rawIssueDate !== "null" && rawIssueDate !== "undefined" ? rawIssueDate : "";
          const parsedIssueDate = validIssueDate ? parseAnyDate(validIssueDate) : null;

          const rawSupervisor = clean(colMap.supervisor !== undefined ? row[colMap.supervisor] : "");
          const validSupervisor = rawSupervisor && rawSupervisor !== "-" && rawSupervisor !== "—" && rawSupervisor !== "N/A" && rawSupervisor.toLowerCase() !== "unassigned" ? rawSupervisor : "";

          const directStitchingRaw = clean(colMap.directStitching !== undefined ? row[colMap.directStitching] : "");
          const isDirect = directStitchingRaw.toLowerCase() === "yes" || directStitchingRaw.toLowerCase() === "direct" || directStitchingRaw.toLowerCase() === "y";
          const challanHistory = clean(colMap.challanHistory !== undefined ? row[colMap.challanHistory] : "");
          const cuttingQty = parseFloat(String((colMap.cuttingQty !== undefined ? row[colMap.cuttingQty] : "") || "0").replace(/,/g, "")) || 0;
          const wipStatus = clean(colMap.wipStatus !== undefined ? row[colMap.wipStatus] : "");
          const compStatus = clean(colMap.completedStatus !== undefined ? row[colMap.completedStatus] : "");

          const isCuttingDone = Boolean(savedAt || cuttingQty > 0);
          const isStitchingIssued = Boolean(validIssueDate || validSupervisor);
          const isStitchingDone = isLotCompleted(compStatus);

          let completedStatusDisplay = "—";
          if (compStatus) {
            if (compStatus.startsWith("[") || compStatus.startsWith("{")) {
              try {
                const parsed = JSON.parse(compStatus);
                const first = Array.isArray(parsed) ? parsed[0] : parsed;
                const ts = first?.timestamp || first?.completionDate || first?.date;
                if (ts) {
                  const d = parseAnyDate(ts);
                  if (d) {
                    completedStatusDisplay = `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`;
                  }
                }
              } catch { }
            } else {
              const d = parseAnyDate(compStatus);
              if (d) {
                completedStatusDisplay = `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`;
              } else if (compStatus.toLowerCase().includes("complete") || compStatus.toLowerCase().includes("done")) {
                completedStatusDisplay = "Completed";
              }
            }
          }

          const stitchingIssueQty = parseFloat(String((colMap.stitchingIssueQty !== undefined ? row[colMap.stitchingIssueQty] : "") || "0").replace(/,/g, "")) || cuttingQty;

          const idxEntry = {
            lot,
            savedAt,
            parsedCutDate,
            issueDate: validIssueDate,
            dateOfIssue: validIssueDate,
            parsedIssueDate,
            supervisor: validSupervisor,
            directStitching: isDirect ? "Yes" : "No",
            isDirect,
            challanHistory,
            cuttingQty,
            stitchingIssueQty,
            wipStatus,
            compStatus,
            completedStatusDisplay,
            imageUrl: clean(colMap.imageUrl !== undefined ? row[colMap.imageUrl] : ""),
            fabric: clean(colMap.fabric !== undefined ? row[colMap.fabric] : ""),
            garmentType: clean(colMap.garmentType !== undefined ? row[colMap.garmentType] : ""),
            style: clean(colMap.style !== undefined ? row[colMap.style] : ""),
            brand: clean(colMap.brand !== undefined ? row[colMap.brand] : ""),
            mwk: clean(colMap.mwk !== undefined ? row[colMap.mwk] : ""),
            section: clean(colMap.section !== undefined ? row[colMap.section] : ""),
            season: clean(colMap.season !== undefined ? row[colMap.season] : ""),
            partyName: clean(colMap.partyName !== undefined ? row[colMap.partyName] : ""),
            priority: clean(colMap.priority !== undefined ? row[colMap.priority] : ""),
            isCuttingDone,
            isStitchingIssued,
            isStitchingDone
          };

          indexLotList.push(idxEntry);
          idxMap.set(norm(lot), idxEntry);
        }
      }

      // 3. Process Issues Sheet (Packing Issues) - Exact 1:1 match with PackingPendingtoIssue.js
      const { issuesData: transformedIssuesData, lotMap: issuesLotMap } = transformIssuesData(issuesRes?.values || []);
      const issuesMap = issuesLotMap;

      // 4. Process RAWPACK Sheet (Carton & Finished Allotment) - Exact 1:1 match with PackingPendingtoIssue.js
      const { rawpackData: transformedRawpackData, rawpackLotMap } = transformRawpackData(rawPackRes?.values || []);
      const rawpackMap = rawpackLotMap;

      // 4b. Process 11 Intermediate Finishing Process Sheets - Exact 1:1 match with PackingPendingtoIssue.js
      const { runningLotsMap: processRunningLotsMap, runningLotsSet: processRunningLotsSet, lastProcessCompletedMap } = transformProcessSheetsData(processSheetResults, FINISHING_PROCESS_SHEETS);

      setLoadingProgress(80);
      setLoadingMessage("Evaluating Red Zone Bottlenecks across all 6 Departments...");

      // 5. MASTER BOTTLENECK EVALUATION ENGINE
      const bottlenecks = [];

      // =======================================================================
      // ✂️ STAGE 1: Cutting Report (Pending if Cutting NOT Done, Red Zone if > 2 Days)
      // Matches exact logic and schema of Cuttingreport.js
      // =======================================================================
      jobMasterMap.forEach((job) => {
        if (String(job.status || "").toLowerCase().includes("cancel")) return;
        const idx = idxMap.get(norm(job.lot));
        const isCuttingDone = idx ? idx.isCuttingDone : false;

        if (!isCuttingDone && job.parsedJobDate) {
          const agingDays = calcDaysDiff(job.parsedJobDate);
          if (agingDays > 2) {
            const severity = agingDays >= 5 ? "CRITICAL" : "HIGH";
            const rowFY = getFinancialYearFromDate(job.jobDate);
            const latestUserRemark = extractLatestRemark(cuttingRemarksMap, job.lot, job.userRemarks || "");

            bottlenecks.push({
              id: `${job.lot}-cutting`,
              lotNumber: job.lot,
              financialYear: rowFY,
              stageId: "cutting",
              stageName: "✂️ Cutting",
              stageRoute: "/cutting-report",
              reason: `Cutting Pending (PO Issued ${agingDays}d ago > 2d Standard)`,
              severity,
              agingDays,
              garmentType: job.garmentType || idx?.garmentType || "—",
              style: job.style || idx?.style || "—",
              fabric: job.fabric || idx?.fabric || "—",
              brand: job.brand || idx?.brand || "—",
              section: job.section || idx?.section || "—",
              season: job.season || idx?.season || "—",
              partyName: job.partyName || idx?.partyName || "—",
              directStitching: job.directStitching ? (String(job.directStitching).toLowerCase() === "yes" ? "Yes" : "No") : (idx?.isDirect ? "Yes" : "No"),
              jobOrderNo: job.jobOrderNo || "—",
              poQty: job.quantity,
              stageDoneQty: 0,
              pendingQty: job.quantity,
              variancePcs: job.quantity,
              variancePct: 100,
              stageDate: job.jobDate,
              pendingShade: job.pendingShade || idx?.pendingShade || "—",
              cuttingDate: idx?.savedAt || job.cuttingDate || "—",
              cuttingScanned: idx?.cuttingScanned || job.cuttingScanned || "—",
              userRemarks: latestUserRemark || "—",
              details: `PO Qty: ${job.quantity.toLocaleString()} Pcs • Job Date: ${formatDisplayDate(job.jobDate)} • Days: ${agingDays}d`
            });
          }
        }
      });

      // =======================================================================
      // 🧵 STAGE 2: Embroidery Report (Pending if EMB NOT Done, Red Zone if > 5 Days)
      // Matches exact logic and schema of EmbroideryChallan.js
      // =======================================================================
      embChallanRows.forEach((ch) => {
        if (String(ch.status || "").toLowerCase().includes("cancel")) return;
        const embStatus = getEmbStatus(ch.challanHistoryJson);
        const isCompleted = embStatus.text === "Emb Done";

        if (!isCompleted && ch.parsedChallanDate) {
          const idx = idxMap.get(norm(ch.lot)) || {};
          const agingDays = calcDaysDiff(ch.parsedChallanDate);
          if (agingDays > 5) {
            const severity = agingDays >= 8 ? "CRITICAL" : "HIGH";
            const rowFY = getFinancialYearFromDate(ch.challanDateStr);
            const embQty = ch.challanQty || ch.qty || idx.cuttingQty || 0;
            const latestEmbRemark = extractLatestRemark(embRemarksMap, ch.lot, ch.userRemarks || "");

            bottlenecks.push({
              id: `${ch.lot}-embroidery-${ch.challanNo || "ch"}`,
              lotNumber: ch.lot,
              financialYear: rowFY,
              stageId: "embroidery",
              stageName: "🧵 Embroidery",
              stageRoute: "/embroidery",
              reason: `Embroidery Pending (${agingDays}d since Challan > 5d Standard)`,
              severity,
              agingDays,
              garmentType: ch.garmentType || idx.garmentType || "—",
              style: ch.style || idx.style || "—",
              fabric: ch.fabric || idx.fabric || "—",
              brand: ch.brand || idx.brand || "—",
              section: ch.section || idx.section || "—",
              season: ch.season || idx.season || "—",
              partyName: ch.partyName || idx.partyName || "—",
              directStitching: ch.directStitching ? (String(ch.directStitching).toLowerCase() === "yes" ? "Yes" : "No") : (idx.isDirect ? "Yes" : "No"),
              embParty: ch.emb || ch.partyName || idx.partyName || "—",
              challanNo: ch.challanNo || "—",
              challanDate: ch.challanDateStr || "—",
              embStatus: embStatus.text || "Pending",
              pendingShade: ch.pendingShade || idx.pendingShade || "—",
              embDoneDate: embStatus.embDoneDate || "—",
              receivedDate: embStatus.receivedDate || "—",
              hodRemarks: latestEmbRemark || "—",
              poQty: embQty,
              stageDoneQty: 0,
              pendingQty: embQty,
              variancePcs: embQty,
              variancePct: 100,
              stageDate: ch.challanDateStr,
              details: `Challan: ${ch.challanNo || "CH-EMB"} • Qty: ${embQty.toLocaleString()} Pcs • EMB Aging: ${agingDays}d`
            });
          }
        }
      });

      // =======================================================================
      // 🖨️ STAGE 3: Printing Report (Pending if Printing NOT Done, Red Zone if > 5 Days)
      // Matches exact logic and schema of PrintingChallan.js
      // =======================================================================
      printChallanRows.forEach((ch) => {
        if (String(ch.status || "").toLowerCase().includes("cancel")) return;
        const printStatus = getPrintingStatus(ch.challanHistoryJson);
        const isCompleted = printStatus.text === "Printing Done";

        if (!isCompleted && ch.parsedChallanDate) {
          const idx = idxMap.get(norm(ch.lot)) || {};
          const agingDays = calcDaysDiff(ch.parsedChallanDate);
          if (agingDays > 5) {
            const severity = agingDays >= 8 ? "CRITICAL" : "HIGH";
            const rowFY = getFinancialYearFromDate(ch.challanDateStr);
            const printQty = ch.challanQty || ch.qty || idx.cuttingQty || 0;
            const latestPrintRemark = extractLatestRemark(printRemarksMap, ch.lot, ch.userRemarks || "");

            bottlenecks.push({
              id: `${ch.lot}-printing-${ch.challanNo || "ch"}`,
              lotNumber: ch.lot,
              financialYear: rowFY,
              stageId: "printing",
              stageName: "🖨️ Printing",
              stageRoute: "/printing",
              reason: `Screen Printing Pending (${agingDays}d since Challan > 5d Standard)`,
              severity,
              agingDays,
              garmentType: ch.garmentType || idx.garmentType || "—",
              style: ch.style || idx.style || "—",
              fabric: ch.fabric || idx.fabric || "—",
              brand: ch.brand || idx.brand || "—",
              section: ch.section || idx.section || "—",
              season: ch.season || idx.season || "—",
              partyName: ch.partyName || idx.partyName || "—",
              directStitching: ch.directStitching ? (String(ch.directStitching).toLowerCase() === "yes" ? "Yes" : "No") : (idx.isDirect ? "Yes" : "No"),
              printingParty: ch.printing || ch.partyName || idx.partyName || "—",
              challanNo: ch.challanNo || "—",
              challanDate: ch.challanDateStr || "—",
              printingStatus: printStatus.text || "Pending",
              pendingShade: ch.pendingShade || idx.pendingShade || "—",
              printingDoneDate: printStatus.printDoneDate || "—",
              hodRemarks: latestPrintRemark || "—",
              poQty: printQty,
              stageDoneQty: 0,
              pendingQty: printQty,
              variancePcs: printQty,
              variancePct: 100,
              stageDate: ch.challanDateStr,
              details: `Challan: ${ch.challanNo || "CH-PRINT"} • Qty: ${printQty.toLocaleString()} Pcs • Print Aging: ${agingDays}d`
            });
          }
        }
      });

      // =======================================================================
      // 📋 STAGE 4: Post-EMB/Print Handover (PendingIssue.js, Red Zone if > 2 Days)
      // Matches exact logic, fields, and schema of PendingIssue.js
      // =======================================================================
      indexLotList.forEach((idx) => {
        if (!idx.isCuttingDone || idx.isStitchingIssued) return;

        const isDirect = Boolean(
          idx.isDirect ||
          (idx.directStitching && (idx.directStitching.toLowerCase() === "yes" || idx.directStitching.toLowerCase() === "y" || idx.directStitching.toLowerCase() === "direct"))
        );

        const history = parseChallanHistory(idx.challanHistory);
        const validChallans = history.filter((c) => c && c.items && Array.isArray(c.items) && c.items.length > 0);
        const hasCompletedEmbChallans = validChallans.length > 0 && validChallans.every((c) => c.embCompleted === true && c.embUpdatedAt);
        const lastEmbDate = getLastEmbDate(idx.challanHistory);

        // PendingIssue.js qualification: Direct Stitching OR all EMB challans completed
        if (!isDirect && !hasCompletedEmbChallans && !lastEmbDate) return;

        const agingDays = isDirect
          ? calcDaysDiff(idx.parsedCutDate)
          : (lastEmbDate ? calcDaysDiff(lastEmbDate) : calcDaysDiff(idx.parsedCutDate));

        if (agingDays > 2) {
          const job = jobMasterMap.get(norm(idx.lot)) || {};
          const rowFY = getFinancialYearFromDate(lastEmbDate || idx.savedAt);
          const isRepeatedLot = Boolean(
            (idx.priority && idx.priority.toLowerCase().includes("repeat")) ||
            (job.priority && job.priority.toLowerCase().includes("repeat"))
          );
          const latestPostEmbRemark = extractLatestRemark(pendingStitchingRemarksMap, idx.lot, job.userRemarks || "");
          const pendingShadeText = idx.pendingShade || job.pendingShade || "No Colour Pending";
          const priorityText = idx.priority || job.priority || "Normal";

          bottlenecks.push({
            id: `${idx.lot}-post-emb-print`,
            lotNumber: idx.lot,
            financialYear: rowFY,
            stageId: "post_emb_print",
            stageName: "📋 Post-EMB/Print",
            stageRoute: "/pending-issue-to-stitching",
            reason: isDirect
              ? `Direct Cut Bundles Pending Line Issue (${agingDays}d > 2d Standard)`
              : `Embellished Panels Pending Sewing Line Issue (${agingDays}d > 2d Standard)`,
            severity: agingDays >= 5 ? "CRITICAL" : "HIGH",
            agingDays,
            garmentType: idx.garmentType || job.garmentType || "—",
            style: idx.style || job.style || "—",
            fabric: idx.fabric || job.fabric || "—",
            brand: idx.brand || job.brand || "—",
            section: idx.section || job.section || "—",
            season: idx.season || job.season || "—",
            partyName: idx.partyName || job.partyName || (isDirect ? "Direct Stitching" : "—"),
            directStitching: isDirect ? "yes" : "no",
            cuttingDate: formatDisplayDate(idx.parsedCutDate) || idx.savedAt || "—",
            embPrintDate: isDirect ? "Direct" : (lastEmbDate ? formatDisplayDate(lastEmbDate) : "—"),
            pendingShade: pendingShadeText,
            priority: priorityText,
            isRepeatedLot,
            userRemarks: latestPostEmbRemark || "—",
            poQty: idx.cuttingQty || job.quantity || 0,
            stageDoneQty: 0,
            pendingQty: idx.cuttingQty || job.quantity || 0,
            variancePcs: idx.cuttingQty || job.quantity || 0,
            variancePct: 100,
            stageDate: isDirect ? idx.savedAt : (lastEmbDate ? formatDisplayDate(lastEmbDate) : idx.savedAt),
            details: isDirect
              ? `Cut Qty: ${idx.cuttingQty.toLocaleString()} Pcs • Direct Stitching • Waiting Lag: ${agingDays}d`
              : `Cut Qty: ${idx.cuttingQty.toLocaleString()} Pcs • EMB/Print Finished • Transit Lag: ${agingDays}d`
          });
        }
      });

      // =======================================================================
      // 🪡 STAGE 5: Stitching WIP (StitchingCompleted.js, Red Zone if > 15 Days)
      // Matches exact logic, fields, and schema of StitchingCompleted.js
      // =======================================================================
      indexLotList.forEach((idx) => {
        if (!idx.isStitchingIssued || idx.isStitchingDone) return;

        // Calculate stitching days strictly from Date of Issue (like StitchingCompleted.js)
        const stitchAging = calculateStitchingDays(idx.issueDate, idx.compStatus, false);
        if (stitchAging > 15) {
          const job = jobMasterMap.get(norm(idx.lot)) || {};
          const severity = stitchAging >= 30 ? "CRITICAL" : "HIGH";
          const rowFY = getFinancialYearFromDate(idx.issueDate || idx.savedAt);
          const embPrintDateFormatted = getEmbPrintDate(idx.challanHistory);
          const wu = workingUpdatesMap.get(norm(idx.lot)) || {};
          const isDirect = Boolean(
            idx.isDirect ||
            (idx.directStitching && (idx.directStitching.toLowerCase() === "yes" || idx.directStitching.toLowerCase() === "y" || idx.directStitching.toLowerCase() === "direct"))
          );

          const mwkAbbr = abbreviateMWK(idx.mwk || job.mwk);
          const derivedSection = mwkAbbr === "M" ? "GENTS" : mwkAbbr === "W" ? "WOMEN" : mwkAbbr === "K" ? "KIDS" : mwkAbbr === "G" ? "GIRLS" : mwkAbbr === "B" ? "BOYS" : "";
          const sectionFormatted = derivedSection || (idx.section && idx.section !== "N/A" && idx.section !== "—" ? idx.section : (job.section || "—"));

          const mwkFormatted = abbreviateMWK(idx.mwk || idx.section || job.section);
          const formattedIssueDate = formatDateToDDMMYY(idx.issueDate);
          const latestWipRemark = getLatestWipRemarks(idx.wipStatus, false);

          bottlenecks.push({
            id: `${idx.lot}-stitching-wip`,
            lotNumber: idx.lot,
            financialYear: rowFY,
            stageId: "stitching",
            stageName: "🪡 Stitching WIP",
            stageRoute: "/stitching-complete-lot",
            reason: `Stitching Pending (${stitchAging}d on Line > 15d Standard)`,
            severity,
            agingDays: stitchAging,
            garmentType: idx.garmentType || job.garmentType || "—",
            style: idx.style || job.style || "—",
            fabric: idx.fabric || job.fabric || "—",
            brand: idx.brand || job.brand || "—",
            section: sectionFormatted,
            season: idx.season || job.season || "—",
            partyName: idx.partyName || job.partyName || "—",
            directStitching: isDirect ? "yes" : "no",
            supervisor: idx.supervisor || job.supervisor || "Unassigned",
            mwk: mwkFormatted,
            dateOfIssue: formattedIssueDate,
            stitchingDays: stitchAging,
            embPrintDate: embPrintDateFormatted,
            wipStatus: formatLatestRemark(latestWipRemark, "WIP"),
            pintu: formatLatestRemark(
              (wu.pintuStatus && wu.pintuStatus !== "—" && wu.pintuStatus !== "N/A" ? wu.pintuStatus : "") ||
              extractLatestRemark(pendingStitchingRemarksMap, idx.lot, ""),
              "—"
            ),
            ea: formatLatestRemark(
              (wu.eaStatus && wu.eaStatus !== "—" && wu.eaStatus !== "N/A" ? wu.eaStatus : "") ||
              extractLatestRemark(cuttingRemarksMap, idx.lot, ""),
              "—"
            ),
            completionDate: "Pending",
            status: "Pending",
            poQty: idx.cuttingQty || job.quantity || 0,
            stageDoneQty: 0,
            pendingQty: idx.cuttingQty || job.quantity || 0,
            variancePcs: idx.cuttingQty || job.quantity || 0,
            variancePct: 100,
            stageDate: idx.issueDate || idx.savedAt,
            details: `Supervisor: ${titleCase(idx.supervisor || "Line Leader")} • Issue Date: ${formattedIssueDate} • Line Days: ${stitchAging}d`
          });
        }
      });

      // =======================================================================
      // 📦⏳ STAGE 6: Packing Pending to Issue (PackingPendingtoIssue.js)
      // Exactly matches PackingPendingtoIssue.js ready for packing view (33 lots)
      // =======================================================================
      indexLotList.forEach((idx) => {
        const lotNumber = idx.lot?.toString().trim();
        if (!lotNumber) return;
        const normalizedLot = lotNumber.toUpperCase();
        const cleanKey = normalizedLot.replace(/[^A-Z0-9]/g, "");

        const hasCompletedStatus = idx.compStatus && idx.compStatus !== "" && idx.compStatus !== "-" && idx.compStatus !== "[]";
        const isNotInIssuesSheet = !issuesLotMap.has(lotNumber) && !issuesLotMap.has(normalizedLot) && (!cleanKey || !issuesLotMap.has(cleanKey));

        const rawpackInfo = rawpackLotMap.get(lotNumber) || rawpackLotMap.get(normalizedLot) || (cleanKey ? rawpackLotMap.get(cleanKey) : null);
        const isRawpackIssuedOrCompleted = rawpackInfo && (Boolean(rawpackInfo.packingPerson) || Boolean(rawpackInfo.packingIssueDate) || rawpackInfo.isCompleted);

        // Check if lot has running intermediate processes
        const isRunningInProcess = processRunningLotsSet.has(lotNumber) || processRunningLotsSet.has(normalizedLot) || (cleanKey && processRunningLotsSet.has(cleanKey));

        const isEligibleBase = hasCompletedStatus && isNotInIssuesSheet && !isRawpackIssuedOrCompleted;
        if (!isEligibleBase) return;

        // Stage 6: Ready for Packing (NOT running in any finishing process)
        if (!isRunningInProcess) {
          const agingDays = calculatePendingDays(idx.completedStatusDisplay);
          const job = jobMasterMap.get(norm(lotNumber)) || {};
          const rowFY = getFinancialYearFromDate(idx.issueDate || idx.savedAt);
          const packingRemarks = packingRemarksMap[lotNumber] || packingRemarksMap[norm(lotNumber)] || "";

          bottlenecks.push({
            id: `${lotNumber}-packing-handover`,
            lotNumber: lotNumber,
            financialYear: rowFY,
            stageId: "packing_handover",
            stageName: "📦⏳ Packing Pending to Issue",
            stageRoute: "/pending-packing-issue",
            reason: `Stitching Done, Ready for Packing (${agingDays}d Pending)`,
            severity: agingDays >= 5 ? "CRITICAL" : (agingDays >= 3 ? "HIGH" : "MODERATE"),
            agingDays,
            imageUrl: idx.imageUrl || job.imageUrl || "",
            garmentType: idx.garmentType || job.garmentType || "—",
            style: idx.style || job.style || "—",
            fabric: idx.fabric || job.fabric || "—",
            brand: idx.brand || job.brand || "—",
            section: idx.section || job.section || "—",
            mwk: idx.mwk || idx.section || job.section || "—",
            season: idx.season || job.season || "—",
            partyName: idx.partyName || job.partyName || "—",
            directStitching: idx.directStitching || "No",
            supervisor: idx.supervisor || "—",
            dateOfIssue: idx.dateOfIssue || idx.issueDate || idx.savedAt || "—",
            priority: idx.priority || "Normal",
            completedDate: idx.completedStatusDisplay || formatDisplayDate(idx.parsedIssueDate) || "—",
            userRemarks: formatLatestRemark(packingRemarks || idx.remarks || idx.userRemarks || "—"),
            poQty: idx.stitchingIssueQty || idx.cuttingQty || job.quantity || 0,
            stitchingQty: idx.stitchingIssueQty || idx.cuttingQty || job.quantity || 0,
            stageDoneQty: 0,
            pendingQty: idx.stitchingIssueQty || idx.cuttingQty || job.quantity || 0,
            variancePcs: idx.stitchingIssueQty || idx.cuttingQty || job.quantity || 0,
            variancePct: 100,
            stageDate: idx.issueDate || idx.savedAt,
            details: `Stitching Done • Awaiting Packing Issue • Pending: ${agingDays}d`
          });
        }
      });

      setAllBottleneckRows(bottlenecks);
      setLastUpdated(new Date().toLocaleString());
      setLoadingProgress(100);
    } catch (err) {
      if (err?.name === "AbortError") {
        console.log("Request aborted");
      } else {
        console.error("Error aggregating critical variance:", err);
        setError(err.message || "Failed to load critical variance data");
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

  /* ---------- Financial Year Options ---------- */
  const availableFinancialYears = useMemo(() => {
    const set = new Set();
    const currentFY = getCurrentFinancialYear();
    if (currentFY) set.add(currentFY);
    allBottleneckRows.forEach((r) => {
      if (r.financialYear) set.add(r.financialYear);
    });
    return Array.from(set).sort().reverse();
  }, [allBottleneckRows]);

  /* ---------- Base FY-Filtered Rows for Options ---------- */
  const fyBaseRows = useMemo(() => {
    if (!financialYearFilter || financialYearFilter === "ALL") return allBottleneckRows;
    return allBottleneckRows.filter((r) => r.financialYear === financialYearFilter);
  }, [allBottleneckRows, financialYearFilter]);

  /* ---------- Unique Options for Multi-Select Filters ---------- */
  const uniqueGarments = useMemo(() => Array.from(new Set(fyBaseRows.map((r) => r.garmentType).filter((g) => g && g !== "—"))).sort(), [fyBaseRows]);
  const uniqueBrands = useMemo(() => Array.from(new Set(fyBaseRows.map((r) => r.brand).filter((b) => b && b !== "—"))).sort(), [fyBaseRows]);
  const uniqueFabrics = useMemo(() => Array.from(new Set(fyBaseRows.map((r) => r.fabric).filter((f) => f && f !== "—"))).sort(), [fyBaseRows]);
  const uniqueParties = useMemo(() => Array.from(new Set(fyBaseRows.map((r) => r.partyName).filter((p) => p && p !== "—"))).sort(), [fyBaseRows]);
  const uniqueSections = useMemo(() => Array.from(new Set(fyBaseRows.map((r) => r.section).filter((s) => s && s !== "—"))).sort(), [fyBaseRows]);
  const uniqueSeasons = useMemo(() => Array.from(new Set(fyBaseRows.map((r) => r.season).filter((s) => s && s !== "—"))).sort(), [fyBaseRows]);

  /* ---------- Stage Counts for Tab Chips ---------- */
  const stageCounts = useMemo(() => {
    const counts = { all: fyBaseRows.length };
    CRITICAL_STAGES.forEach((stage) => {
      if (stage.id !== "all") {
        counts[stage.id] = fyBaseRows.filter((r) => r.stageId === stage.id).length;
      }
    });
    return counts;
  }, [fyBaseRows]);

  /* ---------- Filtered Bottleneck Dataset ---------- */
  const filteredRows = useMemo(() => {
    let list = fyBaseRows;

    // Filter by Active Stage Tab
    if (activeStageTab !== "all") {
      list = list.filter((r) => r.stageId === activeStageTab);
    }

    // Filter by Severity
    if (filterSeverity !== "all") {
      list = list.filter((r) => r.severity === filterSeverity);
    }

    // Multi-Select Dropdown Filters
    if (filterGarments.length > 0) list = list.filter((r) => filterGarments.includes(r.garmentType));
    if (filterBrands.length > 0) list = list.filter((r) => filterBrands.includes(r.brand));
    if (filterFabrics.length > 0) list = list.filter((r) => filterFabrics.includes(r.fabric));
    if (filterParties.length > 0) list = list.filter((r) => filterParties.includes(r.partyName));
    if (filterSections.length > 0) list = list.filter((r) => filterSections.includes(r.section));
    if (filterSeasons.length > 0) list = list.filter((r) => filterSeasons.includes(r.season));

    // Threshold Filters
    if (minAgingDays !== "") {
      const minDays = parseFloat(minAgingDays) || 0;
      list = list.filter((r) => r.agingDays >= minDays);
    }
    if (minVariancePcs !== "") {
      const minPcs = parseFloat(minVariancePcs) || 0;
      list = list.filter((r) => Math.abs(r.variancePcs) >= minPcs);
    }

    // Global Search Term
    if (searchTerm.trim()) {
      const term = searchTerm.toLowerCase().trim();
      list = list.filter((r) =>
        r.lotNumber.toLowerCase().includes(term) ||
        r.stageName.toLowerCase().includes(term) ||
        r.reason.toLowerCase().includes(term) ||
        r.garmentType.toLowerCase().includes(term) ||
        r.style.toLowerCase().includes(term) ||
        r.fabric.toLowerCase().includes(term) ||
        r.brand.toLowerCase().includes(term) ||
        r.partyName.toLowerCase().includes(term) ||
        r.section.toLowerCase().includes(term) ||
        r.season.toLowerCase().includes(term) ||
        r.details.toLowerCase().includes(term)
      );
    }

    // Sorting
    if (sortConfig.key) {
      const { key, direction } = sortConfig;
      list = [...list].sort((a, b) => {
        let va = a[key];
        let vb = b[key];

        if (typeof va === "number" && typeof vb === "number") {
          return direction === "asc" ? va - vb : vb - va;
        }

        va = String(va || "").toLowerCase();
        vb = String(vb || "").toLowerCase();
        if (va < vb) return direction === "asc" ? -1 : 1;
        if (va > vb) return direction === "asc" ? 1 : -1;
        return 0;
      });
    }

    return list;
  }, [
    fyBaseRows,
    activeStageTab,
    filterSeverity,
    filterGarments,
    filterBrands,
    filterFabrics,
    filterParties,
    filterSections,
    filterSeasons,
    minAgingDays,
    minVariancePcs,
    searchTerm,
    sortConfig
  ]);

  /* ---------- Executive Analytics KPI Calculations ---------- */
  const analytics = useMemo(() => {
    const totalLots = filteredRows.length;
    const criticalCount = filteredRows.filter((r) => r.severity === "CRITICAL").length;
    const highCount = filteredRows.filter((r) => r.severity === "HIGH").length;
    const moderateCount = filteredRows.filter((r) => r.severity === "MODERATE").length;

    const totalPendingPcs = filteredRows.reduce((sum, r) => sum + (r.pendingQty || 0), 0);
    const totalVariancePcs = filteredRows.reduce((sum, r) => sum + Math.abs(r.variancePcs || 0), 0);

    const maxAging = filteredRows.reduce((max, r) => Math.max(max, r.agingDays || 0), 0);
    const avgAging = totalLots > 0 ? Math.round(filteredRows.reduce((sum, r) => sum + (r.agingDays || 0), 0) / totalLots) : 0;

    // Department with most bottlenecks
    const stageCountsMap = {};
    filteredRows.forEach((r) => {
      stageCountsMap[r.stageName] = (stageCountsMap[r.stageName] || 0) + 1;
    });

    let topStage = "None";
    let topStageCount = 0;
    Object.entries(stageCountsMap).forEach(([st, cnt]) => {
      if (cnt > topStageCount) {
        topStage = st;
        topStageCount = cnt;
      }
    });

    return {
      totalLots,
      criticalCount,
      highCount,
      moderateCount,
      totalPendingPcs,
      totalVariancePcs,
      maxAging,
      avgAging,
      topStage,
      topStageCount
    };
  }, [filteredRows]);

  const handleSort = (key) => {
    setSortConfig((prev) => ({
      key,
      direction: prev.key === key && prev.direction === "asc" ? "desc" : "asc"
    }));
  };

  const clearAllFilters = () => {
    setActiveStageTab("all");
    setFinancialYearFilter(getCurrentFinancialYear() || "2026-2027");
    setFilterSeverity("all");
    setFilterGarments([]);
    setFilterBrands([]);
    setFilterFabrics([]);
    setFilterParties([]);
    setFilterSections([]);
    setFilterSeasons([]);
    setSearchTerm("");
    setMinAgingDays("");
    setMinVariancePcs("");
    setSortConfig({ key: "agingDays", direction: "desc" });
    setCurrentPage(1);
  };

  const hasActiveFilters = Boolean(
    activeStageTab !== "all" ||
    (financialYearFilter && financialYearFilter !== (getCurrentFinancialYear() || "2026-2027")) ||
    filterSeverity !== "all" ||
    filterGarments.length > 0 ||
    filterBrands.length > 0 ||
    filterFabrics.length > 0 ||
    filterParties.length > 0 ||
    filterSections.length > 0 ||
    filterSeasons.length > 0 ||
    searchTerm ||
    minAgingDays ||
    minVariancePcs
  );

  // Pagination Slice
  const totalPages = Math.ceil(filteredRows.length / pageSize) || 1;
  const paginatedRows = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return filteredRows.slice(start, start + pageSize);
  }, [filteredRows, currentPage, pageSize]);

  // Live Animated Tour Download State
  const [isAnimatedDownloading, setIsAnimatedDownloading] = useState(false);
  const [animProgress, setAnimProgress] = useState({
    stageId: "",
    title: "",
    shortTitle: "",
    icon: "🚀",
    color: "#059669",
    step: 0,
    totalSteps: 0,
    lotCount: 0,
    statusText: "",
    format: "excel",
    completed: false
  });
  const abortAnimationRef = useRef(false);

  /* ======================================================================= */
  /* STAGE REPORT CONFIGURATIONS & EXPORT HELPERS                            */
  /* ======================================================================= */
  const STAGE_REPORT_CONFIGS = [
    { id: "cutting", reportTitle: "REPORT 1 - CUTTING VARIANCE REPORT", shortTitle: "1. Cutting Variance", sheetName: "1. Cutting Variance", filePrefix: "Report_1_Cutting_Variance", color: "FF059669", rgb: [5, 150, 105], icon: "✂️" },
    { id: "embroidery", reportTitle: "REPORT 2 - EMBROIDERY VARIANCE REPORT", shortTitle: "2. Embroidery Variance", sheetName: "2. Embroidery Variance", filePrefix: "Report_2_Embroidery_Variance", color: "FF7C3AED", rgb: [124, 58, 237], icon: "🧵" },
    { id: "printing", reportTitle: "REPORT 3 - PRINTING VARIANCE REPORT", shortTitle: "3. Printing Variance", sheetName: "3. Printing Variance", filePrefix: "Report_3_Printing_Variance", color: "FFDB2777", rgb: [219, 39, 119], icon: "🖨️" },
    { id: "post_emb_print", reportTitle: "REPORT 4 - POST-EMB/PRINT VARIANCE REPORT", shortTitle: "4. Post-EMB Print", sheetName: "4. Post-EMB Print", filePrefix: "Report_4_Post_EMB_Print_Variance", color: "FFEA580C", rgb: [234, 88, 12], icon: "📋" },
    { id: "stitching", reportTitle: "REPORT 5 - STITCHING WIP VARIANCE REPORT", shortTitle: "5. Stitching WIP", sheetName: "5. Stitching WIP", filePrefix: "Report_5_Stitching_WIP_Variance", color: "FF0284C7", rgb: [2, 132, 199], icon: "🪡" },
    { id: "packing_handover", reportTitle: "REPORT 6 - PACKING PENDING TO ISSUE VARIANCE REPORT", shortTitle: "6. Packing Pending to Issue", sheetName: "6. Packing Pending to Issue", filePrefix: "Report_6_Packing_Pending_To_Issue_Variance", color: "FFC2410C", rgb: [194, 65, 12], icon: "📦⏳" }
  ];

  const borderThin = {
    top: { style: "thin", color: { argb: "CBD5E1" } },
    left: { style: "thin", color: { argb: "CBD5E1" } },
    bottom: { style: "thin", color: { argb: "CBD5E1" } },
    right: { style: "thin", color: { argb: "CBD5E1" } }
  };

  const masterCols = [
    { header: "Sr.", key: "sr", width: 6 },
    { header: "Lot Number", key: "lot", width: 14 },
    { header: "Critical Stage", key: "stage", width: 22 },
    { header: "Bottleneck / Variance Reason", key: "reason", width: 38 },
    { header: "Severity", key: "severity", width: 14 },
    { header: "Aging (Days)", key: "aging", width: 14 },
    { header: "Garment Type", key: "garment", width: 18 },
    { header: "Style", key: "style", width: 18 },
    { header: "Fabric", key: "fabric", width: 22 },
    { header: "Brand", key: "brand", width: 16 },
    { header: "Party Name", key: "party", width: 20 },
    { header: "Section", key: "section", width: 12 },
    { header: "Season", key: "season", width: 12 },
    { header: "PO / Cut Pcs", key: "poQty", width: 14 },
    { header: "Done Pcs", key: "doneQty", width: 12 },
    { header: "Pending Pcs", key: "pendingQty", width: 14 },
    { header: "Variance Pcs", key: "varPcs", width: 14 },
    { header: "Variance %", key: "varPct", width: 12 },
    { header: "Stage Date", key: "date", width: 14 }
  ];

  /* ----------------- DEDICATED STAGE EXCEL EXPORTERS ----------------- */

  // 1. Cutting Excel Export (Matches Cuttingreport.js schema)
  const exportCuttingExcel = async (stageRows) => {
    const now = new Date();
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Factory Suite Pro";
    workbook.created = now;

    const ws = workbook.addWorksheet("Cutting Red Zone Report", {
      views: [{ showGridLines: true, state: "frozen", xSplit: 0, ySplit: 4 }]
    });

    const cuttingCols = [
      { header: "S. No", key: "sr", width: 8 },
      { header: "Lot No", key: "lot", width: 14 },
      { header: "Garment Type", key: "garment", width: 18 },
      { header: "Style", key: "style", width: 18 },
      { header: "Fabric", key: "fabric", width: 22 },
      { header: "Brand", key: "brand", width: 16 },
      { header: "Total Qty", key: "poQty", width: 14 },
      { header: "Section", key: "section", width: 12 },
      { header: "Season", key: "season", width: 12 },
      { header: "Party Name", key: "party", width: 20 },
      { header: "Direct Stitching", key: "direct", width: 15 },
      { header: "Job Order No", key: "jobOrderNo", width: 16 },
      { header: "Date", key: "date", width: 14 },
      { header: "Days after PO issue", key: "days", width: 16 },
      { header: "Pending Shade", key: "pendingShade", width: 20 },
      { header: "Cutting Date", key: "cuttingDate", width: 15 },
      { header: "Cutting Scanned", key: "cuttingScanned", width: 15 },
      { header: "💬 User Remarks", key: "userRemarks", width: 28 }
    ];
    ws.columns = cuttingCols;

    const tRow = ws.getRow(1);
    tRow.values = ["FACTORY SUITE PRO — CUTTING RED ZONE REPORT (> 2 DAYS DELAYED)"];
    ws.mergeCells(1, 1, 1, cuttingCols.length);
    tRow.font = { name: "Segoe UI", size: 13, bold: true, color: { argb: "FFFFFFFF" } };
    tRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF059669" } };
    tRow.alignment = { vertical: "middle", horizontal: "center" };
    tRow.height = 30;

    const totalPcs = stageRows.reduce((s, r) => s + (r.poQty || 0), 0);
    const mRow = ws.getRow(2);
    mRow.values = [`Generated: ${now.toLocaleDateString()}  |  Total Red Zone Lots: ${stageRows.length}  |  Total Pending Pieces: ${totalPcs.toLocaleString()} Pcs  |  FY: ${financialYearFilter}`];
    ws.mergeCells(2, 1, 2, cuttingCols.length);
    mRow.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF1E293B" } };
    mRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
    mRow.alignment = { vertical: "middle", horizontal: "center" };
    mRow.height = 20;

    ws.getRow(3).height = 6;

    const hRow = ws.getRow(4);
    hRow.values = cuttingCols.map(c => c.header);
    hRow.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
    hRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F172A" } };
    hRow.alignment = { vertical: "middle", horizontal: "center" };
    hRow.height = 26;

    stageRows.forEach((r, idx) => {
      const row = ws.addRow({
        sr: idx + 1,
        lot: r.lotNumber,
        garment: r.garmentType || "—",
        style: r.style || "—",
        fabric: r.fabric || "—",
        brand: r.brand || "—",
        poQty: r.poQty || 0,
        section: r.section || "—",
        season: r.season || "—",
        party: r.partyName || "—",
        direct: r.directStitching || "No",
        jobOrderNo: r.jobOrderNo || "—",
        date: r.stageDate || "—",
        days: r.agingDays != null ? r.agingDays : "—",
        pendingShade: r.pendingShade || "—",
        cuttingDate: r.cuttingDate || "—",
        cuttingScanned: r.cuttingScanned || "—",
        userRemarks: r.userRemarks || "—"
      });
      row.font = { name: "Segoe UI", size: 9.5 };
      row.alignment = { vertical: "middle" };
      row.eachCell((c, cNum) => {
        c.border = borderThin;
        if ([1, 2, 8, 9, 11, 12, 13, 14, 16, 17].includes(cNum)) c.alignment = { horizontal: "center" };
        if (cNum === 7) {
          c.alignment = { horizontal: "right" };
          c.numFmt = "#,##0";
          c.font = { name: "Segoe UI", size: 9.5, bold: true };
        }
        if (cNum === 14) {
          c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEE2E2" } };
          c.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFDC2626" } };
        }
      });
    });

    const subTotalRow = ws.addRow({
      sr: "",
      lot: `TOTAL (${stageRows.length} Lots)`,
      garment: "",
      style: "",
      fabric: "",
      brand: "",
      poQty: totalPcs,
      section: "",
      season: "",
      party: "",
      direct: "",
      jobOrderNo: "",
      date: "",
      days: "",
      pendingShade: "",
      cuttingDate: "",
      cuttingScanned: "",
      userRemarks: ""
    });
    subTotalRow.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF0F172A" } };
    subTotalRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } };
    subTotalRow.eachCell((c, cNum) => {
      c.border = borderThin;
      if (cNum === 7) {
        c.alignment = { horizontal: "right" };
        c.numFmt = "#,##0";
      }
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    saveAs(blob, `Report_1_Cutting_Variance_${now.toISOString().slice(0, 10)}.xlsx`);
  };

  // 2. Embroidery Excel Export (Matches EmbroideryChallan.js schema)
  const exportEmbroideryExcel = async (stageRows) => {
    const now = new Date();
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Factory Suite Pro";
    workbook.created = now;

    const ws = workbook.addWorksheet("Embroidery Red Zone Report", {
      views: [{ showGridLines: true, state: "frozen", xSplit: 0, ySplit: 4 }]
    });

    const embCols = [
      { header: "S. No", key: "sr", width: 8 },
      { header: "Lot No.", key: "lot", width: 14 },
      { header: "Garment Type", key: "garment", width: 18 },
      { header: "Style", key: "style", width: 18 },
      { header: "Fabric", key: "fabric", width: 22 },
      { header: "Brand", key: "brand", width: 16 },
      { header: "Challan Qty", key: "poQty", width: 14 },
      { header: "Section (M/W/K)", key: "section", width: 14 },
      { header: "Season", key: "season", width: 12 },
      { header: "Party Name", key: "party", width: 20 },
      { header: "Direct Stitching", key: "direct", width: 15 },
      { header: "Emb Party", key: "embParty", width: 18 },
      { header: "Challan Date", key: "date", width: 14 },
      { header: "Emb Status", key: "embStatus", width: 15 },
      { header: "Pending Shade", key: "pendingShade", width: 20 },
      { header: "Emb Done Date", key: "embDoneDate", width: 15 },
      { header: "Received Date", key: "receivedDate", width: 15 },
      { header: "Days", key: "days", width: 12 },
      { header: "HOD Remarks", key: "hodRemarks", width: 28 }
    ];
    ws.columns = embCols;

    const tRow = ws.getRow(1);
    tRow.values = ["FACTORY SUITE PRO — EMBROIDERY RED ZONE REPORT (> 5 DAYS DELAYED)"];
    ws.mergeCells(1, 1, 1, embCols.length);
    tRow.font = { name: "Segoe UI", size: 13, bold: true, color: { argb: "FFFFFFFF" } };
    tRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF7C3AED" } };
    tRow.alignment = { vertical: "middle", horizontal: "center" };
    tRow.height = 30;

    const totalPcs = stageRows.reduce((s, r) => s + (r.poQty || 0), 0);
    const mRow = ws.getRow(2);
    mRow.values = [`Generated: ${now.toLocaleDateString()}  |  Total Red Zone Lots: ${stageRows.length}  |  Total Pending Pieces: ${totalPcs.toLocaleString()} Pcs  |  FY: ${financialYearFilter}`];
    ws.mergeCells(2, 1, 2, embCols.length);
    mRow.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF1E293B" } };
    mRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
    mRow.alignment = { vertical: "middle", horizontal: "center" };
    mRow.height = 20;

    ws.getRow(3).height = 6;

    const hRow = ws.getRow(4);
    hRow.values = embCols.map(c => c.header);
    hRow.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
    hRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F172A" } };
    hRow.alignment = { vertical: "middle", horizontal: "center" };
    hRow.height = 26;

    stageRows.forEach((r, idx) => {
      const row = ws.addRow({
        sr: idx + 1,
        lot: r.lotNumber,
        garment: r.garmentType || "—",
        style: r.style || "—",
        fabric: r.fabric || "—",
        brand: r.brand || "—",
        poQty: r.poQty || 0,
        section: r.section || "—",
        season: r.season || "—",
        party: r.partyName || "—",
        direct: r.directStitching || "No",
        embParty: r.embParty || r.partyName || "—",
        date: r.stageDate || "—",
        embStatus: r.embStatus || "Pending",
        pendingShade: r.pendingShade || "—",
        embDoneDate: r.embDoneDate || "—",
        receivedDate: r.receivedDate || "—",
        days: r.agingDays != null ? r.agingDays : "—",
        hodRemarks: r.hodRemarks || "—"
      });
      row.font = { name: "Segoe UI", size: 9.5 };
      row.alignment = { vertical: "middle" };
      row.eachCell((c, cNum) => {
        c.border = borderThin;
        if ([1, 2, 8, 9, 11, 12, 13, 14, 16, 17, 18].includes(cNum)) c.alignment = { horizontal: "center" };
        if (cNum === 7) {
          c.alignment = { horizontal: "right" };
          c.numFmt = "#,##0";
          c.font = { name: "Segoe UI", size: 9.5, bold: true };
        }
        if (cNum === 18) {
          c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEE2E2" } };
          c.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFDC2626" } };
        }
      });
    });

    const subTotalRow = ws.addRow({
      sr: "",
      lot: `TOTAL (${stageRows.length} Lots)`,
      garment: "",
      style: "",
      fabric: "",
      brand: "",
      poQty: totalPcs,
      section: "",
      season: "",
      party: "",
      direct: "",
      embParty: "",
      date: "",
      embStatus: "",
      pendingShade: "",
      embDoneDate: "",
      receivedDate: "",
      days: "",
      hodRemarks: ""
    });
    subTotalRow.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF0F172A" } };
    subTotalRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } };
    subTotalRow.eachCell((c, cNum) => {
      c.border = borderThin;
      if (cNum === 7) {
        c.alignment = { horizontal: "right" };
        c.numFmt = "#,##0";
      }
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    saveAs(blob, `Report_2_Embroidery_Variance_${now.toISOString().slice(0, 10)}.xlsx`);
  };

  // 3. Printing Excel Export (Matches PrintingChallan.js schema)
  const exportPrintingExcel = async (stageRows) => {
    const now = new Date();
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Factory Suite Pro";
    workbook.created = now;

    const ws = workbook.addWorksheet("Printing Red Zone Report", {
      views: [{ showGridLines: true, state: "frozen", xSplit: 0, ySplit: 4 }]
    });

    const printCols = [
      { header: "S. No", key: "sr", width: 8 },
      { header: "Lot No.", key: "lot", width: 14 },
      { header: "Garment Type", key: "garment", width: 18 },
      { header: "Style", key: "style", width: 18 },
      { header: "Fabric", key: "fabric", width: 22 },
      { header: "Brand", key: "brand", width: 16 },
      { header: "Challan Qty", key: "poQty", width: 14 },
      { header: "Section (M/W/K)", key: "section", width: 14 },
      { header: "Season", key: "season", width: 12 },
      { header: "Party Name", key: "party", width: 20 },
      { header: "Direct Stitching", key: "direct", width: 15 },
      { header: "Printing Party", key: "printingParty", width: 18 },
      { header: "Challan Date", key: "date", width: 14 },
      { header: "Printing Status", key: "printingStatus", width: 15 },
      { header: "Pending Shade", key: "pendingShade", width: 20 },
      { header: "Printing Done Date", key: "printingDoneDate", width: 15 },
      { header: "Days", key: "days", width: 12 },
      { header: "HOD Remarks", key: "hodRemarks", width: 28 }
    ];
    ws.columns = printCols;

    const tRow = ws.getRow(1);
    tRow.values = ["FACTORY SUITE PRO — PRINTING RED ZONE REPORT (> 5 DAYS DELAYED)"];
    ws.mergeCells(1, 1, 1, printCols.length);
    tRow.font = { name: "Segoe UI", size: 13, bold: true, color: { argb: "FFFFFFFF" } };
    tRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDB2777" } };
    tRow.alignment = { vertical: "middle", horizontal: "center" };
    tRow.height = 30;

    const totalPcs = stageRows.reduce((s, r) => s + (r.poQty || 0), 0);
    const mRow = ws.getRow(2);
    mRow.values = [`Generated: ${now.toLocaleDateString()}  |  Total Red Zone Lots: ${stageRows.length}  |  Total Pending Pieces: ${totalPcs.toLocaleString()} Pcs  |  FY: ${financialYearFilter}`];
    ws.mergeCells(2, 1, 2, printCols.length);
    mRow.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF1E293B" } };
    mRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
    mRow.alignment = { vertical: "middle", horizontal: "center" };
    mRow.height = 20;

    ws.getRow(3).height = 6;

    const hRow = ws.getRow(4);
    hRow.values = printCols.map(c => c.header);
    hRow.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
    hRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F172A" } };
    hRow.alignment = { vertical: "middle", horizontal: "center" };
    hRow.height = 26;

    stageRows.forEach((r, idx) => {
      const row = ws.addRow({
        sr: idx + 1,
        lot: r.lotNumber,
        garment: r.garmentType || "—",
        style: r.style || "—",
        fabric: r.fabric || "—",
        brand: r.brand || "—",
        poQty: r.poQty || 0,
        section: r.section || "—",
        season: r.season || "—",
        party: r.partyName || "—",
        direct: r.directStitching || "No",
        printingParty: r.printingParty || r.partyName || "—",
        date: r.stageDate || "—",
        printingStatus: r.printingStatus || "Pending",
        pendingShade: r.pendingShade || "—",
        printingDoneDate: r.printingDoneDate || "—",
        days: r.agingDays != null ? r.agingDays : "—",
        hodRemarks: r.hodRemarks || "—"
      });
      row.font = { name: "Segoe UI", size: 9.5 };
      row.alignment = { vertical: "middle" };
      row.eachCell((c, cNum) => {
        c.border = borderThin;
        if ([1, 2, 8, 9, 11, 12, 13, 14, 16, 17].includes(cNum)) c.alignment = { horizontal: "center" };
        if (cNum === 7) {
          c.alignment = { horizontal: "right" };
          c.numFmt = "#,##0";
          c.font = { name: "Segoe UI", size: 9.5, bold: true };
        }
        if (cNum === 17) {
          c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEE2E2" } };
          c.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFDC2626" } };
        }
      });
    });

    const subTotalRow = ws.addRow({
      sr: "",
      lot: `TOTAL (${stageRows.length} Lots)`,
      garment: "",
      style: "",
      fabric: "",
      brand: "",
      poQty: totalPcs,
      section: "",
      season: "",
      party: "",
      direct: "",
      printingParty: "",
      date: "",
      printingStatus: "",
      pendingShade: "",
      printingDoneDate: "",
      days: "",
      hodRemarks: ""
    });
    subTotalRow.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF0F172A" } };
    subTotalRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } };
    subTotalRow.eachCell((c, cNum) => {
      c.border = borderThin;
      if (cNum === 7) {
        c.alignment = { horizontal: "right" };
        c.numFmt = "#,##0";
      }
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    saveAs(blob, `Report_3_Printing_Variance_${now.toISOString().slice(0, 10)}.xlsx`);
  };

  // 4. Post-EMB/Print Excel Export (Matches PendingIssue.js schema)
  const exportPostEmbPrintExcel = async (stageRows) => {
    const now = new Date();
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Factory Suite Pro";
    workbook.created = now;

    const totalLots = stageRows.length;
    const totalPieces = stageRows.reduce((s, r) => s + (r.poQty || 0), 0);
    const colorPendingCount = stageRows.filter(r => r.pendingShade && !r.pendingShade.toLowerCase().includes("no col")).length;
    const repeatedLotCount = stageRows.filter(r => r.isRepeatedLot).length;

    // Production Breakdown Maps for Summary
    const garmentMap = {};
    const seasonMap = {};
    const partyMap = {};
    let normalLots = 0, normalPcs = 0;
    let urgentLots = 0, urgentPcs = 0;
    let criticalLots = 0, criticalPcs = 0;

    stageRows.forEach(issue => {
      const pcs = issue.poQty || 0;
      const gType = (issue.garmentType || 'Unknown').trim();
      const season = (issue.season || 'Other / NA').trim();
      const party = (issue.partyName || 'Unknown').trim();
      const days = typeof issue.agingDays === 'number' ? issue.agingDays : 0;

      if (!garmentMap[gType]) garmentMap[gType] = { totalLots: 0, totalPcs: 0 };
      garmentMap[gType].totalLots += 1;
      garmentMap[gType].totalPcs += pcs;

      if (!seasonMap[season]) seasonMap[season] = { totalLots: 0, totalPcs: 0 };
      seasonMap[season].totalLots += 1;
      seasonMap[season].totalPcs += pcs;

      if (!partyMap[party]) partyMap[party] = { totalLots: 0, totalPcs: 0 };
      partyMap[party].totalLots += 1;
      partyMap[party].totalPcs += pcs;

      if (days <= 2) {
        normalLots += 1;
        normalPcs += pcs;
      } else if (days <= 5) {
        urgentLots += 1;
        urgentPcs += pcs;
      } else {
        criticalLots += 1;
        criticalPcs += pcs;
      }
    });

    const sortedGarments = Object.keys(garmentMap).map(name => ({
      name,
      totalLots: garmentMap[name].totalLots,
      totalPcs: garmentMap[name].totalPcs
    })).sort((a, b) => b.totalPcs - a.totalPcs);

    const sortedSeasons = Object.keys(seasonMap).map(name => ({
      name,
      totalLots: seasonMap[name].totalLots,
      totalPcs: seasonMap[name].totalPcs
    })).sort((a, b) => b.totalPcs - a.totalPcs);

    const sortedParties = Object.keys(partyMap).map(name => ({
      name,
      totalLots: partyMap[name].totalLots,
      totalPcs: partyMap[name].totalPcs
    })).sort((a, b) => b.totalPcs - a.totalPcs);

    // ================= SHEET 1: PENDING ISSUES =================
    const ws1 = workbook.addWorksheet('Pending Issues', {
      views: [{ showGridLines: true, state: "frozen", xSplit: 0, ySplit: 4 }]
    });

    const postEmbCols = [
      { header: "#", key: "sr", width: 6 },
      { header: "Lot Number", key: "lot", width: 16 },
      { header: "Garment Type", key: "garment", width: 18 },
      { header: "Style", key: "style", width: 20 },
      { header: "Fabric", key: "fabric", width: 22 },
      { header: "Brand", key: "brand", width: 16 },
      { header: "Total Pcs", key: "poQty", width: 14 },
      { header: "M/W/K", key: "section", width: 10 },
      { header: "Season", key: "season", width: 12 },
      { header: "Party Name", key: "party", width: 20 },
      { header: "Direct Stitching", key: "direct", width: 15 },
      { header: "Cutting Date", key: "cuttingDate", width: 14 },
      { header: "Emb/Printing Date", key: "embPrintDate", width: 16 },
      { header: "Days Pending", key: "days", width: 14 },
      { header: "Color Status", key: "pendingShade", width: 22 },
      { header: "Priority", key: "priority", width: 14 },
      { header: "💬 User Remarks", key: "userRemarks", width: 28 }
    ];
    ws1.columns = postEmbCols;

    // Banner Row 1: Title
    ws1.mergeCells('A1:Q1');
    const titleCell = ws1.getCell('A1');
    titleCell.value = 'FACTORY SUITE PRO - PENDING ISSUES TO STITCHING (POST-EMB/PRINT RED ZONE)';
    titleCell.font = { name: 'Segoe UI', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEA580C' } };
    ws1.getRow(1).height = 30;

    // Banner Row 2: Subtitle & KPI
    ws1.mergeCells('A2:Q2');
    const subCell = ws1.getCell('A2');
    subCell.value = `Total Lots: ${totalLots}   |   Total Pieces: ${totalPieces.toLocaleString()}   |   Color Pending: ${colorPendingCount}   |   Repeated Lots: ${repeatedLotCount}   |   Red Zone (>2d): ${stageRows.length}   |   FY: ${financialYearFilter}`;
    subCell.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FF1E293B' } };
    subCell.alignment = { horizontal: 'center', vertical: 'middle' };
    subCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
    ws1.getRow(2).height = 20;

    ws1.getRow(3).height = 6;

    // Table Headers (Row 4)
    const headerRow1 = ws1.getRow(4);
    headerRow1.values = postEmbCols.map(c => c.header);
    headerRow1.height = 26;
    headerRow1.eachCell((cell) => {
      cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
      cell.border = borderThin;
    });

    // Data Rows
    stageRows.forEach((r, idx) => {
      const days = r.agingDays != null ? r.agingDays : null;
      const row = ws1.addRow({
        sr: idx + 1,
        lot: r.isRepeatedLot ? `★ ${r.lotNumber}` : r.lotNumber,
        garment: r.garmentType || '—',
        style: r.style || '—',
        fabric: r.fabric || '—',
        brand: r.brand || '—',
        poQty: r.poQty || 0,
        section: r.section || '—',
        season: r.season || '—',
        party: r.partyName || '—',
        direct: r.directStitching || 'No',
        cuttingDate: r.cuttingDate || '—',
        embPrintDate: r.embPrintDate || r.stageDate || '—',
        days: days !== null ? `${days} days` : '—',
        pendingShade: r.pendingShade || 'No Colour Pending',
        priority: r.priority || 'Normal',
        userRemarks: r.userRemarks || '—'
      });

      row.font = { name: 'Segoe UI', size: 9.5 };
      row.alignment = { vertical: 'middle' };
      row.eachCell((c, cNum) => {
        c.border = borderThin;
        if ([1, 2, 8, 9, 11, 12, 13, 14, 16].includes(cNum)) c.alignment = { horizontal: 'center' };
        if (cNum === 7) {
          c.alignment = { horizontal: 'right' };
          c.numFmt = '#,##0';
          c.font = { name: 'Segoe UI', size: 9.5, bold: true };
        }
      });

      // Highlight repeated lots
      if (r.isRepeatedLot) {
        row.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
        row.getCell(2).font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FF92400E' } };
      }

      // Highlight days pending
      if (days !== null) {
        if (days >= 5) {
          row.getCell(14).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
          row.getCell(14).font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFDC2626' } };
        } else if (days > 2) {
          row.getCell(14).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
          row.getCell(14).font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFB45309' } };
        } else {
          row.getCell(14).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } };
          row.getCell(14).font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FF15803D' } };
        }
      }
    });

    // Total Row
    const totalRow1 = ws1.addRow({
      sr: '',
      lot: `TOTAL (${totalLots} Lots)`,
      garment: '',
      style: '',
      fabric: '',
      brand: '',
      poQty: totalPieces,
      section: '',
      season: '',
      party: '',
      direct: '',
      cuttingDate: '',
      embPrintDate: '',
      days: '',
      pendingShade: `${colorPendingCount} Color Pending`,
      priority: `${repeatedLotCount} Repeated`,
      userRemarks: ''
    });
    totalRow1.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FF0F172A' } };
    totalRow1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
    totalRow1.eachCell((c, cNum) => {
      c.border = borderThin;
      if (cNum === 7) {
        c.alignment = { horizontal: 'right' };
        c.numFmt = '#,##0';
      }
      if ([15, 16].includes(cNum)) c.alignment = { horizontal: 'center' };
    });

    // ================= SHEET 2: EXECUTIVE SUMMARY =================
    const ws2 = workbook.addWorksheet('Executive Summary', { views: [{ showGridLines: true }] });

    // Section 1: Garment Type Breakdown
    ws2.mergeCells('A1:D1');
    const gTitle = ws2.getCell('A1');
    gTitle.value = '1. GARMENT TYPE BREAKDOWN (LOTS & PIECES DISTRIBUTION)';
    gTitle.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    gTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F766E' } };
    gTitle.alignment = { horizontal: 'left', vertical: 'middle' };
    ws2.getRow(1).height = 26;

    const gHeader = ws2.addRow(['Garment Type', 'Total Lots', 'Total Pieces (Qty)', 'Share %']);
    gHeader.height = 22;
    gHeader.eachCell(c => {
      c.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF134E4A' } };
      c.alignment = { horizontal: 'center', vertical: 'middle' };
      c.border = borderThin;
    });

    sortedGarments.forEach((item, idx) => {
      const pct = totalPieces > 0 ? (item.totalPcs / totalPieces) : 0;
      const r = ws2.addRow([item.name, item.totalLots, item.totalPcs, pct]);
      r.height = 19;
      r.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
      r.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
      r.getCell(3).alignment = { horizontal: 'right', vertical: 'middle' };
      r.getCell(3).numFmt = '#,##0';
      r.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
      r.getCell(4).numFmt = '0.0%';
      r.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9 };
        c.border = borderThin;
        if (idx % 2 === 1) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
      });
    });

    const gTotalRow = ws2.addRow(['TOTAL', totalLots, totalPieces, 1]);
    gTotalRow.height = 22;
    gTotalRow.eachCell(c => {
      c.font = { name: 'Segoe UI', size: 9.5, bold: true };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
      c.border = borderThin;
    });
    gTotalRow.getCell(3).alignment = { horizontal: 'right', vertical: 'middle' };
    gTotalRow.getCell(3).numFmt = '#,##0';
    gTotalRow.getCell(4).numFmt = '0.0%';

    // Section 2: Season Breakdown
    ws2.addRow([]);
    const seasonStartRow = ws2.rowCount + 1;
    ws2.mergeCells(`A${seasonStartRow}:D${seasonStartRow}`);
    const sTitle = ws2.getCell(`A${seasonStartRow}`);
    sTitle.value = '2. SEASON WISE BREAKDOWN (LOTS & PIECES DISTRIBUTION)';
    sTitle.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    sTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4338CA' } };
    sTitle.alignment = { horizontal: 'left', vertical: 'middle' };
    ws2.getRow(seasonStartRow).height = 26;

    const sHeader = ws2.addRow(['Season', 'Total Lots', 'Total Pieces (Qty)', 'Share %']);
    sHeader.height = 22;
    sHeader.eachCell(c => {
      c.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF312E81' } };
      c.alignment = { horizontal: 'center', vertical: 'middle' };
      c.border = borderThin;
    });

    sortedSeasons.forEach((item, idx) => {
      const pct = totalPieces > 0 ? (item.totalPcs / totalPieces) : 0;
      const r = ws2.addRow([item.name, item.totalLots, item.totalPcs, pct]);
      r.height = 19;
      r.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
      r.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
      r.getCell(3).alignment = { horizontal: 'right', vertical: 'middle' };
      r.getCell(3).numFmt = '#,##0';
      r.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
      r.getCell(4).numFmt = '0.0%';
      r.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9 };
        c.border = borderThin;
        if (idx % 2 === 1) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
      });
    });

    const sTotalRow = ws2.addRow(['TOTAL', totalLots, totalPieces, 1]);
    sTotalRow.height = 22;
    sTotalRow.eachCell(c => {
      c.font = { name: 'Segoe UI', size: 9.5, bold: true };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
      c.border = borderThin;
    });
    sTotalRow.getCell(3).alignment = { horizontal: 'right', vertical: 'middle' };
    sTotalRow.getCell(3).numFmt = '#,##0';
    sTotalRow.getCell(4).numFmt = '0.0%';

    // Section 3: Party Summary
    ws2.addRow([]);
    const partyStartRow = ws2.rowCount + 1;
    ws2.mergeCells(`A${partyStartRow}:D${partyStartRow}`);
    const pTitle = ws2.getCell(`A${partyStartRow}`);
    pTitle.value = '3. PARTY SUMMARY & WORKLOAD ALLOCATION';
    pTitle.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    pTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E40AF' } };
    pTitle.alignment = { horizontal: 'left', vertical: 'middle' };
    ws2.getRow(partyStartRow).height = 26;

    const pHeader = ws2.addRow(['Party Name', 'Total Lots', 'Total Pieces (Qty)', 'Share %']);
    pHeader.height = 22;
    pHeader.eachCell(c => {
      c.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
      c.alignment = { horizontal: 'center', vertical: 'middle' };
      c.border = borderThin;
    });

    sortedParties.forEach((item, idx) => {
      const pct = totalPieces > 0 ? (item.totalPcs / totalPieces) : 0;
      const r = ws2.addRow([item.name, item.totalLots, item.totalPcs, pct]);
      r.height = 19;
      r.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
      r.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
      r.getCell(3).alignment = { horizontal: 'right', vertical: 'middle' };
      r.getCell(3).numFmt = '#,##0';
      r.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
      r.getCell(4).numFmt = '0.0%';
      r.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9 };
        c.border = borderThin;
        if (idx % 2 === 1) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
      });
    });

    const pTotalRow = ws2.addRow(['TOTAL', totalLots, totalPieces, 1]);
    pTotalRow.height = 22;
    pTotalRow.eachCell(c => {
      c.font = { name: 'Segoe UI', size: 9.5, bold: true };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
      c.border = borderThin;
    });
    pTotalRow.getCell(3).alignment = { horizontal: 'right', vertical: 'middle' };
    pTotalRow.getCell(3).numFmt = '#,##0';
    pTotalRow.getCell(4).numFmt = '0.0%';

    ws2.getColumn(1).width = 32;
    ws2.getColumn(2).width = 16;
    ws2.getColumn(3).width = 22;
    ws2.getColumn(4).width = 16;

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    saveAs(blob, `Report_4_Post_EMB_Print_Variance_${now.toISOString().slice(0, 10)}.xlsx`);
  };

  // 6. Stitching WIP Excel Export (Matches StitchingCompleted.js schema)
  const exportStitchingExcel = async (stageRows) => {
    const now = new Date();
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Factory Suite Pro";
    workbook.created = now;

    const totalLots = stageRows.length;
    const totalPieces = stageRows.reduce((s, r) => s + (r.poQty || 0), 0);

    // Production Breakdown Maps for Summary
    const supervisorMap = {};
    const garmentMap = {};
    let redLots = 0, redPcs = 0;

    stageRows.forEach(issue => {
      const pcs = issue.poQty || 0;
      const sup = (issue.supervisor || 'Unassigned').trim();
      const gType = (issue.garmentType || 'Unknown').trim();
      const days = typeof issue.agingDays === 'number' ? issue.agingDays : 0;

      if (!supervisorMap[sup]) supervisorMap[sup] = { totalLots: 0, totalPcs: 0, redLots: 0 };
      supervisorMap[sup].totalLots += 1;
      supervisorMap[sup].totalPcs += pcs;

      if (!garmentMap[gType]) garmentMap[gType] = { totalLots: 0, totalPcs: 0 };
      garmentMap[gType].totalLots += 1;
      garmentMap[gType].totalPcs += pcs;

      if (days > 15) {
        redLots += 1;
        redPcs += pcs;
        supervisorMap[sup].redLots += 1;
      }
    });

    const sortedSupervisors = Object.keys(supervisorMap).map(name => ({
      name,
      totalLots: supervisorMap[name].totalLots,
      totalPcs: supervisorMap[name].totalPcs,
      redLots: supervisorMap[name].redLots
    })).sort((a, b) => b.totalPcs - a.totalPcs);

    const sortedGarments = Object.keys(garmentMap).map(name => ({
      name,
      totalLots: garmentMap[name].totalLots,
      totalPcs: garmentMap[name].totalPcs
    })).sort((a, b) => b.totalPcs - a.totalPcs);

    // ================= SHEET 1: STITCHING WIP RED ZONE =================
    const ws1 = workbook.addWorksheet('Stitching WIP Red Zone', {
      views: [{ showGridLines: true, state: "frozen", xSplit: 0, ySplit: 4 }]
    });

    const stitchingCols = [
      { header: "#", key: "sr", width: 6 },
      { header: "Lot Number", key: "lot", width: 16 },
      { header: "Garment Type", key: "garment", width: 18 },
      { header: "Style", key: "style", width: 20 },
      { header: "Fabric", key: "fabric", width: 22 },
      { header: "BRAND", key: "brand", width: 16 },
      { header: "Total PCS", key: "poQty", width: 14 },
      { header: "Section", key: "section", width: 12 },
      { header: "Season", key: "season", width: 12 },
      { header: "PARTY NAME", key: "party", width: 20 },
      { header: "Direct Stitching", key: "direct", width: 15 },
      { header: "Supervisor", key: "supervisor", width: 18 },
      { header: "M/W/K", key: "mwk", width: 10 },
      { header: "Date of Issue", key: "issueDate", width: 15 },
      { header: "Stitching Days", key: "days", width: 15 },
      { header: "Emb/Print Date", key: "embDate", width: 16 },
      { header: "WIP Status", key: "wipStatus", width: 22 },
      { header: "Pintu Remarks", key: "pintu", width: 22 },
      { header: "EA Remarks", key: "ea", width: 22 },
      { header: "Status", key: "status", width: 14 }
    ];
    ws1.columns = stitchingCols;

    // Banner Row 1: Title
    ws1.mergeCells('A1:T1');
    const titleCell = ws1.getCell('A1');
    titleCell.value = 'FACTORY SUITE PRO - STITCHING WIP RED ZONE REPORT (> 15 DAYS DELAYED)';
    titleCell.font = { name: 'Segoe UI', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0284C7' } };
    ws1.getRow(1).height = 30;

    // Banner Row 2: Subtitle & KPI
    ws1.mergeCells('A2:T2');
    const subCell = ws1.getCell('A2');
    subCell.value = `Total Lots: ${totalLots}   |   Total Pieces: ${totalPieces.toLocaleString()}   |   Critical (>15d): ${redLots} Lots   |   Supervisors: ${sortedSupervisors.length}   |   FY: ${financialYearFilter}`;
    subCell.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FF1E293B' } };
    subCell.alignment = { horizontal: 'center', vertical: 'middle' };
    subCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
    ws1.getRow(2).height = 20;

    ws1.getRow(3).height = 6;

    // Table Headers (Row 4)
    const headerRow1 = ws1.getRow(4);
    headerRow1.values = stitchingCols.map(c => c.header);
    headerRow1.height = 26;
    headerRow1.eachCell((cell) => {
      cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E1B4B' } };
      cell.border = borderThin;
    });

    // Data Rows
    stageRows.forEach((r, idx) => {
      const days = r.agingDays != null ? r.agingDays : null;
      const row = ws1.addRow({
        sr: idx + 1,
        lot: r.lotNumber,
        garment: r.garmentType || '—',
        style: r.style || '—',
        fabric: r.fabric || '—',
        brand: r.brand || '—',
        poQty: r.poQty || 0,
        section: r.section || '—',
        season: r.season || '—',
        party: r.partyName || '—',
        direct: r.directStitching || 'no',
        supervisor: r.supervisor || 'Unassigned',
        mwk: r.mwk || r.section || '—',
        issueDate: r.dateOfIssue || r.stageDate || '—',
        days: days !== null ? `${days} days` : '—',
        embDate: r.embPrintDate || '—',
        wipStatus: formatLatestRemark(r.wipStatus, 'WIP'),
        pintu: formatLatestRemark(r.pintu, '—'),
        ea: formatLatestRemark(r.ea, '—'),
        status: r.status || 'Pending'
      });

      row.font = { name: 'Segoe UI', size: 9.5 };
      row.alignment = { vertical: 'middle' };
      row.eachCell((c, cNum) => {
        c.border = borderThin;
        if ([1, 2, 8, 9, 11, 12, 13, 14, 15, 16, 20].includes(cNum)) c.alignment = { horizontal: 'center' };
        if (cNum === 7) {
          c.alignment = { horizontal: 'right' };
          c.numFmt = '#,##0';
          c.font = { name: 'Segoe UI', size: 9.5, bold: true };
        }
      });

      // Supervisor in bold blue
      const supCell = row.getCell(12);
      supCell.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FF2563EB' } };

      // Days cell highlight
      const daysCell = row.getCell(15);
      daysCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
      daysCell.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFDC2626' } };

      if (idx % 2 === 1) {
        row.eachCell((c, cNum) => {
          if (cNum !== 15) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        });
      }
    });

    // Total Row
    const totalRow = ws1.addRow({
      sr: '',
      lot: 'TOTAL',
      poQty: totalPieces
    });
    totalRow.height = 24;
    totalRow.eachCell((c, cNum) => {
      c.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0284C7' } };
      c.border = borderThin;
      if (cNum === 7) {
        c.alignment = { horizontal: 'right' };
        c.numFmt = '#,##0';
      }
    });

    // ================= SHEET 2: EXECUTIVE SUMMARY =================
    const ws2 = workbook.addWorksheet('Executive Summary', { views: [{ showGridLines: true }] });

    // Table 1: Supervisor Workload
    ws2.mergeCells('A1:D1');
    const sTitle = ws2.getCell('A1');
    sTitle.value = 'SUPERVISOR WORKLOAD & AGING BREAKDOWN';
    sTitle.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    sTitle.alignment = { horizontal: 'center', vertical: 'middle' };
    sTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0284C7' } };
    ws2.getRow(1).height = 24;

    const sHRow = ws2.getRow(2);
    sHRow.values = ['Supervisor', 'Total Lots', 'Total Pieces', 'Share %'];
    sHRow.height = 22;
    sHRow.eachCell(c => {
      c.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
      c.alignment = { horizontal: 'center', vertical: 'middle' };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
      c.border = borderThin;
    });

    sortedSupervisors.forEach((item, idx) => {
      const pct = totalPieces > 0 ? item.totalPcs / totalPieces : 0;
      const r = ws2.addRow([item.name, item.totalLots, item.totalPcs, pct]);
      r.height = 19;
      r.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
      r.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
      r.getCell(3).alignment = { horizontal: 'right', vertical: 'middle' };
      r.getCell(3).numFmt = '#,##0';
      r.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
      r.getCell(4).numFmt = '0.0%';
      r.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9 };
        c.border = borderThin;
        if (idx % 2 === 1) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
      });
    });

    const sTotalRow = ws2.addRow(['TOTAL', totalLots, totalPieces, 1]);
    sTotalRow.height = 22;
    sTotalRow.eachCell(c => {
      c.font = { name: 'Segoe UI', size: 9.5, bold: true };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
      c.border = borderThin;
    });
    sTotalRow.getCell(3).alignment = { horizontal: 'right', vertical: 'middle' };
    sTotalRow.getCell(3).numFmt = '#,##0';
    sTotalRow.getCell(4).numFmt = '0.0%';

    ws2.getColumn(1).width = 28;
    ws2.getColumn(2).width = 16;
    ws2.getColumn(3).width = 20;
    ws2.getColumn(4).width = 14;

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    saveAs(blob, `Report_6_Stitching_WIP_Variance_${now.toISOString().slice(0, 10)}.xlsx`);
  };

  // 6. Packing Pending to Issue Excel Export (Matches PendingPackingtoIssue.js schema)
  const exportPackingHandoverExcel = async (stageRows) => {
    const now = new Date();
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Factory Suite Pro";
    workbook.created = now;

    const totalLots = stageRows.length;
    const totalPieces = stageRows.reduce((s, r) => s + (r.poQty || r.stitchingQty || 0), 0);
    const directLotsCount = stageRows.filter(r => (r.directStitching || "").toLowerCase() === "yes").length;
    const highPriorityCount = stageRows.filter(r => (r.priority || "").toLowerCase() === "high" || (r.priority || "").toLowerCase() === "urgent").length;
    const avgPendingDays = totalLots > 0 ? Math.round(stageRows.reduce((s, r) => s + (r.agingDays || 0), 0) / totalLots) : 0;

    // Production Breakdown Maps for Summary
    const supervisorMap = {};
    const garmentMap = {};
    const seasonMap = {};
    const partyMap = {};
    const agingMap = {
      "0-7 Days": { totalLots: 0, totalQty: 0 },
      "8-15 Days": { totalLots: 0, totalQty: 0 },
      "16-30 Days": { totalLots: 0, totalQty: 0 },
      "30+ Days": { totalLots: 0, totalQty: 0 }
    };

    stageRows.forEach(item => {
      const qty = Number(item.poQty) || Number(item.stitchingQty) || 0;
      const gType = (item.garmentType || "Unknown").trim();
      const season = (item.season || "N/A").trim();
      const sup = (item.supervisor || "Unassigned").trim();
      const party = (item.partyName || (item.directStitching === "yes" ? "Direct Stitching" : "—")).trim();
      const days = item.agingDays || 0;

      if (!garmentMap[gType]) garmentMap[gType] = { totalLots: 0, totalQty: 0 };
      garmentMap[gType].totalLots += 1;
      garmentMap[gType].totalQty += qty;

      if (!seasonMap[season]) seasonMap[season] = { totalLots: 0, totalQty: 0 };
      seasonMap[season].totalLots += 1;
      seasonMap[season].totalQty += qty;

      if (!supervisorMap[sup]) supervisorMap[sup] = { totalLots: 0, totalQty: 0 };
      supervisorMap[sup].totalLots += 1;
      supervisorMap[sup].totalQty += qty;

      if (!partyMap[party]) partyMap[party] = { totalLots: 0, totalQty: 0 };
      partyMap[party].totalLots += 1;
      partyMap[party].totalQty += qty;

      if (days <= 7) {
        agingMap["0-7 Days"].totalLots += 1;
        agingMap["0-7 Days"].totalQty += qty;
      } else if (days <= 15) {
        agingMap["8-15 Days"].totalLots += 1;
        agingMap["8-15 Days"].totalQty += qty;
      } else if (days <= 30) {
        agingMap["16-30 Days"].totalLots += 1;
        agingMap["16-30 Days"].totalQty += qty;
      } else {
        agingMap["30+ Days"].totalLots += 1;
        agingMap["30+ Days"].totalQty += qty;
      }
    });

    const sortedGarments = Object.keys(garmentMap).map(k => ({ name: k, totalLots: garmentMap[k].totalLots, totalQty: garmentMap[k].totalQty })).sort((a, b) => b.totalQty - a.totalQty);
    const sortedSeasons = Object.keys(seasonMap).map(k => ({ name: k, totalLots: seasonMap[k].totalLots, totalQty: seasonMap[k].totalQty })).sort((a, b) => b.totalQty - a.totalQty);
    const sortedSupervisors = Object.keys(supervisorMap).map(k => ({ name: k, totalLots: supervisorMap[k].totalLots, totalQty: supervisorMap[k].totalQty })).sort((a, b) => b.totalQty - a.totalQty);

    // ================= SHEET 1: DATA TABLE =================
    const ws1 = workbook.addWorksheet("Completed Lots Ready", { views: [{ showGridLines: true }] });

    // Title Banner
    ws1.mergeCells("A1:R1");
    const titleCell = ws1.getCell("A1");
    titleCell.value = "MH FACTORY SUITE PRO - REPORT 6 - PACKING PENDING TO ISSUE VARIANCE REPORT";
    titleCell.font = { name: "Segoe UI", size: 13, bold: true, color: { argb: "FFFFFFFF" } };
    titleCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFC2410C" }
    };
    titleCell.alignment = { horizontal: "center", vertical: "middle" };
    ws1.getRow(1).height = 30;

    // Subtitle KPI Banner
    ws1.mergeCells("A2:R2");
    const subCell = ws1.getCell("A2");
    subCell.value = `Completed Lots: ${totalLots}   |   Total Stitching Qty: ${totalPieces.toLocaleString()}   |   Direct Lots: ${directLotsCount}   |   High Priority: ${highPriorityCount}   |   Avg Pending Days: ${avgPendingDays}d   |   Standard: >2d Delayed   |   Generated: ${new Date().toLocaleDateString("en-IN")} ${new Date().toLocaleTimeString("en-IN")}`;
    subCell.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFFFFFFF" } };
    subCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF9A3412" }
    };
    subCell.alignment = { horizontal: "center", vertical: "middle" };
    ws1.getRow(2).height = 22;

    // Table Header Row
    const tableHeaders = [
      "#", "Lot Number", "Garment Type", "Style", "Fabric", "Brand",
      "Stitching Qty", "M/W/K", "Season", "Party Name", "Direct Stitching",
      "Supervisor", "Date of Issue", "Priority", "Completed Date",
      "Pending Days", "Status", "Remarks"
    ];

    const headerRow = ws1.addRow(tableHeaders);
    headerRow.height = 24;
    headerRow.eachCell(cell => {
      cell.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF0F172A" }
      };
      cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      cell.border = borderThin;
    });

    // Data Rows
    stageRows.forEach((item, index) => {
      const lotNo = (item.lotNumber || "").toString().trim();
      const qty = Number(item.poQty) || Number(item.stitchingQty) || 0;
      const days = item.agingDays || 0;
      const isDirect = (item.directStitching || "").toLowerCase() === "yes";

      const r = ws1.addRow([
        index + 1,
        lotNo,
        item.garmentType || "—",
        item.style || "—",
        item.fabric || "—",
        item.brand || "—",
        qty,
        item.mwk || "—",
        item.season || "—",
        item.partyName || "—",
        isDirect ? "Yes" : "No",
        item.supervisor || "—",
        item.dateOfIssue || "—",
        item.priority || "Normal",
        item.completedDate || "—",
        days,
        "Ready for Packing",
        item.userRemarks || "—"
      ]);

      r.height = 20;

      r.eachCell(cell => {
        cell.font = { name: "Segoe UI", size: 9 };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = borderThin;
        if (index % 2 === 1) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
        }
      });

      // Stitching Qty styling
      r.getCell(7).numFmt = "#,##0";
      r.getCell(7).font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FFDC2626" } };

      // Days Aging formatting
      const daysCell = r.getCell(16);
      if (days <= 7) {
        daysCell.font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FF16A34A" } };
      } else if (days <= 15) {
        daysCell.font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FFD97706" } };
      } else {
        daysCell.font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FFDC2626" } };
        daysCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEE2E2" } };
      }

      // Direct badge
      if (isDirect) {
        r.getCell(11).font = { name: "Segoe UI", size: 9, bold: true, color: { argb: "FF15803D" } };
      }
    });

    // Auto fit column widths
    ws1.columns.forEach((column) => {
      let maxLen = 10;
      column.eachCell({ includeEmpty: true }, (cell) => {
        const val = cell.value ? cell.value.toString() : "";
        if (val.length > maxLen && val.length < 50) {
          maxLen = val.length;
        }
      });
      column.width = Math.max(maxLen + 3, 11);
    });
    ws1.getColumn(1).width = 6;
    ws1.getColumn(2).width = 14;
    ws1.getColumn(18).width = 30;

    // ================= SHEET 2: EXECUTIVE SUMMARY BREAKDOWN =================
    const ws2 = workbook.addWorksheet("Executive Summary", { views: [{ showGridLines: true }] });

    ws2.mergeCells("A1:D1");
    const sTitle = ws2.getCell("A1");
    sTitle.value = "MH FACTORY SUITE PRO - REPORT 6 - EXECUTIVE SUMMARY & WORKLOAD BREAKDOWN";
    sTitle.font = { name: "Segoe UI", size: 12, bold: true, color: { argb: "FFFFFFFF" } };
    sTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E1B4B" } };
    sTitle.alignment = { horizontal: "center", vertical: "middle" };
    ws2.getRow(1).height = 28;

    // Summary Section 1: Garment Type Breakdown
    let currentSummaryRow = 3;
    const gHeader = ws2.getRow(currentSummaryRow);
    gHeader.values = ["Garment Type", "Total Lots", "Total Stitching Qty (Pcs)", "Share %"];
    gHeader.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFFFFFFF" } };
    gHeader.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0284C7" } };
    gHeader.alignment = { horizontal: "center", vertical: "middle" };
    gHeader.height = 22;

    sortedGarments.forEach(g => {
      currentSummaryRow++;
      const row = ws2.addRow([g.name, g.totalLots, g.totalQty, totalPieces > 0 ? (g.totalQty / totalPieces) : 0]);
      row.height = 19;
      row.eachCell(c => {
        c.font = { name: "Segoe UI", size: 9 };
        c.border = borderThin;
      });
      row.getCell(2).alignment = { horizontal: "center", vertical: "middle" };
      row.getCell(3).alignment = { horizontal: "right", vertical: "middle" };
      row.getCell(3).numFmt = "#,##0";
      row.getCell(4).alignment = { horizontal: "right", vertical: "middle" };
      row.getCell(4).numFmt = "0.0%";
    });

    // Summary Section 2: Supervisor Workload
    currentSummaryRow += 2;
    const supHeader = ws2.getRow(currentSummaryRow);
    supHeader.values = ["Supervisor", "Total Lots", "Total Stitching Qty (Pcs)", "Share %"];
    supHeader.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFFFFFFF" } };
    supHeader.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4338CA" } };
    supHeader.alignment = { horizontal: "center", vertical: "middle" };
    supHeader.height = 22;

    sortedSupervisors.forEach(sup => {
      currentSummaryRow++;
      const row = ws2.addRow([sup.name, sup.totalLots, sup.totalQty, totalPieces > 0 ? (sup.totalQty / totalPieces) : 0]);
      row.height = 19;
      row.eachCell(c => {
        c.font = { name: "Segoe UI", size: 9 };
        c.border = borderThin;
      });
      row.getCell(2).alignment = { horizontal: "center", vertical: "middle" };
      row.getCell(3).alignment = { horizontal: "right", vertical: "middle" };
      row.getCell(3).numFmt = "#,##0";
      row.getCell(4).alignment = { horizontal: "right", vertical: "middle" };
      row.getCell(4).numFmt = "0.0%";
    });

    // Summary Section 3: Aging Analysis
    currentSummaryRow += 2;
    const ageHeader = ws2.getRow(currentSummaryRow);
    ageHeader.values = ["Aging Category", "Total Lots", "Total Stitching Qty (Pcs)", "Share %"];
    ageHeader.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFFFFFFF" } };
    ageHeader.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDC2626" } };
    ageHeader.alignment = { horizontal: "center", vertical: "middle" };
    ageHeader.height = 22;

    Object.entries(agingMap).forEach(([rangeName, data]) => {
      currentSummaryRow++;
      const row = ws2.addRow([rangeName, data.totalLots, data.totalQty, totalPieces > 0 ? (data.totalQty / totalPieces) : 0]);
      row.height = 19;
      row.eachCell(c => {
        c.font = { name: "Segoe UI", size: 9 };
        c.border = borderThin;
      });
      row.getCell(2).alignment = { horizontal: "center", vertical: "middle" };
      row.getCell(3).alignment = { horizontal: "right", vertical: "middle" };
      row.getCell(3).numFmt = "#,##0";
      row.getCell(4).alignment = { horizontal: "right", vertical: "middle" };
      row.getCell(4).numFmt = "0.0%";
    });

    ws2.getColumn(1).width = 30;
    ws2.getColumn(2).width = 16;
    ws2.getColumn(3).width = 24;
    ws2.getColumn(4).width = 14;

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    saveAs(blob, `Report_6_Packing_Pending_To_Issue_Variance_${now.toISOString().slice(0, 10)}.xlsx`);
  };

  /* Unified Single Stage Excel Exporter Router */
  const exportSingleStageExcel = async (cfg, stageRows) => {
    if (!stageRows || stageRows.length === 0) return;
    if (cfg.id === "cutting") {
      await exportCuttingExcel(stageRows);
    } else if (cfg.id === "embroidery") {
      await exportEmbroideryExcel(stageRows);
    } else if (cfg.id === "printing") {
      await exportPrintingExcel(stageRows);
    } else if (cfg.id === "post_emb_print") {
      await exportPostEmbPrintExcel(stageRows);
    } else if (cfg.id === "stitching") {
      await exportStitchingExcel(stageRows);
    } else if (cfg.id === "packing_handover") {
      await exportPackingHandoverExcel(stageRows);
    } else {
      const now = new Date();
      const reportDateStr = `${String(now.getDate()).padStart(2, "0")}/${String(now.getMonth() + 1).padStart(2, "0")}/${now.getFullYear()}`;
      const workbook = new ExcelJS.Workbook();
      workbook.creator = "Factory Suite Pro";
      workbook.created = now;

      const ws = workbook.addWorksheet(cfg.sheetName, {
        views: [{ showGridLines: true, state: "frozen", xSplit: 0, ySplit: 4 }]
      });
      ws.columns = masterCols;

      const tRow = ws.getRow(1);
      tRow.values = [`FACTORY SUITE PRO — ${cfg.reportTitle}`];
      ws.mergeCells(1, 1, 1, masterCols.length);
      tRow.font = { name: "Segoe UI", size: 13, bold: true, color: { argb: "FFFFFFFF" } };
      tRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: cfg.color } };
      tRow.alignment = { vertical: "middle", horizontal: "center" };
      tRow.height = 30;

      const stagePending = stageRows.reduce((s, r) => s + (r.pendingQty || 0), 0);
      const mRow = ws.getRow(2);
      mRow.values = [`Generated: ${reportDateStr}  |  Total Lots: ${stageRows.length}  |  Pending Quantity: ${stagePending.toLocaleString()} Pcs  |  FY: ${financialYearFilter}`];
      ws.mergeCells(2, 1, 2, masterCols.length);
      mRow.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF1E293B" } };
      mRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
      mRow.alignment = { vertical: "middle", horizontal: "center" };
      mRow.height = 20;

      ws.getRow(3).height = 6;

      const hRow = ws.getRow(4);
      hRow.values = masterCols.map((c) => c.header);
      hRow.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
      hRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F172A" } };
      hRow.alignment = { vertical: "middle", horizontal: "center" };
      hRow.height = 26;

      stageRows.forEach((r, idx) => {
        const row = ws.addRow({
          sr: idx + 1,
          lot: r.lotNumber,
          stage: r.stageName,
          reason: r.reason,
          severity: r.severity,
          aging: r.agingDays,
          garment: r.garmentType,
          style: r.style,
          fabric: r.fabric,
          brand: r.brand,
          party: r.partyName,
          section: r.section,
          season: r.season,
          poQty: r.poQty || 0,
          doneQty: r.stageDoneQty || 0,
          pendingQty: r.pendingQty || 0,
          varPcs: r.variancePcs || 0,
          varPct: `${r.variancePct || 0}%`,
          date: r.stageDate || "—"
        });
        row.font = { name: "Segoe UI", size: 9.5 };
        row.alignment = { vertical: "middle" };
        row.eachCell((c, cNum) => {
          c.border = borderThin;
          if ([1, 2, 5, 6, 12, 13, 19].includes(cNum)) c.alignment = { horizontal: "center" };
          if (cNum >= 14 && cNum <= 18) {
            c.alignment = { horizontal: "right" };
            if (cNum !== 18) c.numFmt = "#,##0";
          }
        });
      });

      const subTotalRow = ws.addRow({
        sr: "",
        lot: `TOTAL (${stageRows.length} Lots)`,
        stage: "",
        reason: "",
        severity: "",
        aging: "",
        garment: "",
        style: "",
        fabric: "",
        brand: "",
        party: "",
        section: "",
        season: "",
        poQty: stageRows.reduce((s, r) => s + (r.poQty || 0), 0),
        doneQty: stageRows.reduce((s, r) => s + (r.stageDoneQty || 0), 0),
        pendingQty: stagePending,
        varPcs: stageRows.reduce((s, r) => s + (r.variancePcs || 0), 0),
        varPct: "",
        date: ""
      });
      subTotalRow.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF0F172A" } };
      subTotalRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } };
      subTotalRow.eachCell((c, cNum) => {
        c.border = borderThin;
        if (cNum >= 14 && cNum <= 17) {
          c.alignment = { horizontal: "right" };
          c.numFmt = "#,##0";
        }
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      saveAs(blob, `${cfg.filePrefix}_${now.toISOString().slice(0, 10)}.xlsx`);
    }
  };

  /* ======================================================================= */
  /* DEDICATED PDF EXPORTERS MATCHING INDIVIDUAL REPORT LAYOUTS & COLUMNS    */
  /* ======================================================================= */

  const scaleColumnStyles = (baseStyles, targetWidth) => {
    const keys = Object.keys(baseStyles);
    const totalBase = keys.reduce((sum, k) => sum + (baseStyles[k].cellWidth || 50), 0);
    const result = {};
    let currentSum = 0;
    keys.forEach((k, idx) => {
      if (idx === keys.length - 1) {
        result[k] = { ...baseStyles[k], cellWidth: Math.max(10, Math.round((targetWidth - currentSum) * 10) / 10) };
      } else {
        const w = Math.round(((baseStyles[k].cellWidth || 50) / totalBase) * targetWidth * 10) / 10;
        result[k] = { ...baseStyles[k], cellWidth: w };
        currentSum += w;
      }
    });
    return result;
  };

  // 1. CUTTING REPORT PDF (Matches Cuttingreport.js exactly)
  const exportCuttingPDF = (stageRows) => {
    const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "A3" });
    const pageW = doc.internal.pageSize.getWidth();
    const pageH = doc.internal.pageSize.getHeight();
    const now = new Date();

    const marginX = 16;
    const tableW = pageW - marginX * 2;

    const totalLots = stageRows.length;
    const totalQty = stageRows.reduce((sum, r) => sum + (r.poQty || 0), 0);
    const redZoneLots = stageRows.filter((r) => r.agingDays > 2).length;
    const normalLots = totalLots - redZoneLots;

    const garmentMap = {};
    const partyMap = {};
    const sectionMap = {};

    stageRows.forEach((r) => {
      const g = r.garmentType || "Unassigned";
      const p = r.partyName || "Unassigned";
      const s = r.section || "Unassigned";
      const qty = r.poQty || 0;

      if (!garmentMap[g]) garmentMap[g] = { name: g, lots: 0, qty: 0 };
      garmentMap[g].lots += 1;
      garmentMap[g].qty += qty;

      if (!partyMap[p]) partyMap[p] = { name: p, lots: 0, qty: 0 };
      partyMap[p].lots += 1;
      partyMap[p].qty += qty;

      if (!sectionMap[s]) sectionMap[s] = { name: s, lots: 0, qty: 0 };
      sectionMap[s].lots += 1;
      sectionMap[s].qty += qty;
    });

    const sortedGarments = Object.values(garmentMap).sort((a, b) => b.qty - a.qty);
    const sortedParties = Object.values(partyMap).sort((a, b) => b.qty - a.qty);
    const sortedSections = Object.values(sectionMap).sort((a, b) => b.qty - a.qty);

    doc.setFontSize(20);
    doc.setTextColor(0, 0, 0);
    doc.setFont("helvetica", "bold");
    doc.text("REPORT 1 - CUTTING PRODUCTION VARIANCE REPORT", pageW / 2, 38, { align: "center" });

    doc.setFontSize(10);
    doc.setTextColor(50, 50, 50);
    doc.setFont("helvetica", "normal");
    doc.text(`Report Date: ${now.toLocaleDateString()} ${now.toLocaleTimeString()} | Factory Suite Pro`, pageW / 2, 54, { align: "center" });

    doc.setFillColor(239, 246, 255);
    doc.roundedRect(marginX, 64, tableW, 24, 6, 6, "F");
    doc.setDrawColor(191, 219, 254);
    doc.roundedRect(marginX, 64, tableW, 24, 6, 6, "D");

    doc.setFontSize(10.5);
    doc.setTextColor(0, 0, 0);
    doc.setFont("helvetica", "bold");
    const summaryText = `Total Records: ${totalLots}   |   Total Quantity: ${totalQty.toLocaleString()} Pcs   |   Parties: ${sortedParties.length}   |   Garments: ${sortedGarments.length}   |   Red Zone (>2 Days): ${redZoneLots}`;
    doc.text(summaryText, pageW / 2, 79, { align: "center" });

    const columns = [
      "#",
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
      "Days after PO issue",
      "Pending Shade",
      "Cutting Date",
      "Cutting Scanned",
      "User Remarks"
    ];

    const body = stageRows.map((r, idx) => [
      String(idx + 1),
      r.lotNumber,
      r.garmentType || "-",
      r.style || "-",
      r.fabric || "-",
      r.brand || "-",
      (r.poQty || 0).toLocaleString(),
      r.section || "-",
      r.season || "-",
      r.partyName || "-",
      r.directStitching || "No",
      r.jobOrderNo || "-",
      r.stageDate || "-",
      r.agingDays != null ? `${r.agingDays}` : "-",
      r.pendingShade || "-",
      r.cuttingDate || "-",
      r.cuttingScanned || "-",
      formatLatestRemark(r.userRemarks, "-")
    ]);

    body.push([
      "",
      `TOTAL (${totalLots})`,
      "",
      "",
      "",
      "",
      totalQty.toLocaleString(),
      "",
      "",
      `${sortedParties.length} Parties`,
      "",
      "",
      "",
      `${redZoneLots} Red | ${normalLots} Norm`,
      "",
      "",
      "",
      ""
    ]);

    const columnStyles = {
      0: { cellWidth: 22, halign: "center" },
      1: { cellWidth: 60, halign: "center", fontStyle: "bold" },
      2: { cellWidth: 70, halign: "center" },
      3: { cellWidth: 75, halign: "center" },
      4: { cellWidth: 75, halign: "center" },
      5: { cellWidth: 50, halign: "center" },
      6: { cellWidth: 50, halign: "center", fontStyle: "bold" },
      7: { cellWidth: 45, halign: "center" },
      8: { cellWidth: 45, halign: "center" },
      9: { cellWidth: 55, halign: "center" },
      10: { cellWidth: 45, halign: "center" },
      11: { cellWidth: 55, halign: "center" },
      12: { cellWidth: 55, halign: "center" },
      13: { cellWidth: 40, halign: "center" },
      14: { cellWidth: 75, halign: "left" },
      15: { cellWidth: 60, halign: "center" },
      16: { cellWidth: 60, halign: "center" },
      17: { cellWidth: 120, halign: "left" }
    };

    const scaledColumnStyles = scaleColumnStyles(columnStyles, tableW);

    autoTable(doc, {
      head: [columns],
      body,
      startY: 96,
      tableWidth: tableW,
      margin: { top: 96, right: marginX, bottom: 25, left: marginX },
      theme: "grid",
      styles: {
        fontSize: 8,
        cellPadding: { top: 3.5, right: 2, bottom: 3.5, left: 2 },
        overflow: "linebreak",
        valign: "middle",
        halign: "center",
        textColor: [0, 0, 0],
        lineColor: [0, 0, 0],
        lineWidth: 0.3
      },
      headStyles: {
        fillColor: [15, 23, 42],
        textColor: [255, 255, 255],
        fontStyle: "bold",
        lineColor: [0, 0, 0],
        lineWidth: 0.5,
        halign: "center",
        fontSize: 8.5,
        valign: "middle",
        cellPadding: { top: 5, right: 2, bottom: 5, left: 2 }
      },
      alternateRowStyles: {
        fillColor: [248, 250, 252]
      },
      columnStyles: scaledColumnStyles,
      didParseCell: function (data) {
        if (data.section === "body") {
          data.cell.styles.textColor = [0, 0, 0];
          const rowIndex = data.row.index;
          const isTotalRow = rowIndex === body.length - 1;

          if (isTotalRow) {
            data.cell.styles.fontStyle = "bold";
            data.cell.styles.fillColor = [226, 232, 240];
            data.cell.styles.textColor = [0, 0, 0];
            return;
          }

          if (data.column.index === 13) {
            const rawVal = parseFloat(data.cell.raw);
            if (!isNaN(rawVal)) {
              data.cell.styles.textColor = [0, 0, 0];
              data.cell.styles.fontStyle = "bold";
              if (rawVal > 2) {
                data.cell.styles.fillColor = [254, 226, 226];
              } else {
                data.cell.styles.fillColor = [220, 252, 231];
              }
            }
          }
        }
      },
      didDrawPage: () => {
        doc.setDrawColor(0, 0, 0);
        doc.setLineWidth(0.5);
        doc.roundedRect(8, 8, pageW - 16, pageH - 16, 2, 2, "S");
      }
    });

    const gBody = sortedGarments.map((item) => {
      const pct = totalQty > 0 ? ((item.qty / totalQty) * 100).toFixed(1) : "0.0";
      return [item.name, item.lots.toString(), item.qty.toLocaleString(), `${pct}%`];
    });
    gBody.push(["TOTAL", totalLots.toString(), totalQty.toLocaleString(), "100.0%"]);

    const pBody = sortedParties.map((item) => {
      const pct = totalQty > 0 ? ((item.qty / totalQty) * 100).toFixed(1) : "0.0";
      return [item.name, item.lots.toString(), item.qty.toLocaleString(), `${pct}%`];
    });
    pBody.push(["TOTAL", totalLots.toString(), totalQty.toLocaleString(), "100.0%"]);

    const sBody = sortedSections.map((item) => {
      const pct = totalQty > 0 ? ((item.qty / totalQty) * 100).toFixed(1) : "0.0";
      return [item.name, item.lots.toString(), item.qty.toLocaleString(), `${pct}%`];
    });
    sBody.push(["TOTAL", totalLots.toString(), totalQty.toLocaleString(), "100.0%"]);

    const aBody = [
      ["Normal (<= 2 Days)", normalLots.toString(), "—", `${totalLots > 0 ? ((normalLots / totalLots) * 100).toFixed(1) : 0}%`],
      ["Red Zone (> 2 Days)", redZoneLots.toString(), "—", `${totalLots > 0 ? ((redZoneLots / totalLots) * 100).toFixed(1) : 0}%`],
      ["TOTAL LOTS", totalLots.toString(), totalQty.toLocaleString(), "100.0%"]
    ];

    let summaryStartY = doc.lastAutoTable.finalY + 22;
    if (summaryStartY + 140 > pageH - 30) {
      doc.addPage();
      summaryStartY = 40;
    }

    doc.setFontSize(13);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(0, 0, 0);
    doc.text("EXECUTIVE SUMMARY & PRODUCTION BREAKDOWN", pageW / 2, summaryStartY + 4, { align: "center" });

    const sectionTitleY = summaryStartY + 26;
    const tableStartY = sectionTitleY + 6;
    const gap = 14;
    const colWidth = (tableW - 3 * gap) / 4;
    const col1X = marginX;
    const col2X = col1X + colWidth + gap;
    const col3X = col2X + colWidth + gap;
    const col4X = col3X + colWidth + gap;

    const summaryColStyles = {
      0: { cellWidth: 110, halign: "center" },
      1: { cellWidth: 45, halign: "center" },
      2: { cellWidth: 68, halign: "center" },
      3: { cellWidth: 55, halign: "center" }
    };

    const scaledSummaryColStyles = scaleColumnStyles(summaryColStyles, colWidth);

    [
      { body: gBody, left: col1X, fill: [15, 118, 110], label: "1. GARMENT BREAKDOWN", header: "Garment Type" },
      { body: pBody, left: col2X, fill: [30, 64, 175], label: "2. PARTY BREAKDOWN", header: "Party Name" },
      { body: sBody, left: col3X, fill: [109, 40, 217], label: "3. SECTION BREAKDOWN", header: "Section" },
      { body: aBody, left: col4X, fill: [180, 83, 9], label: "4. AGING BREAKDOWN", header: "Aging Status" }
    ].forEach((tbl) => {
      autoTable(doc, {
        head: [[tbl.header, "Lots", "Total Qty", "Share %"]],
        body: tbl.body,
        startY: tableStartY,
        tableWidth: colWidth,
        margin: { left: tbl.left, right: pageW - (tbl.left + colWidth) },
        theme: "grid",
        styles: {
          fontSize: 8.5,
          cellPadding: { top: 3.5, right: 2, bottom: 3.5, left: 2 },
          valign: "middle",
          halign: "center",
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          lineWidth: 0.3
        },
        headStyles: {
          fillColor: tbl.fill,
          textColor: [255, 255, 255],
          fontStyle: "bold",
          fontSize: 8.5,
          halign: "center"
        },
        columnStyles: scaledSummaryColStyles,
        didParseCell: (data) => {
          if (data.section === "body") {
            data.cell.styles.textColor = [0, 0, 0];
            if (data.row.index === tbl.body.length - 1) {
              data.cell.styles.fontStyle = "bold";
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });
    });

    doc.save(`Report_1_Cutting_Variance_${now.toISOString().slice(0, 10)}.pdf`);
  };

  // 2. EMBROIDERY CHALLAN REPORT PDF (Matches EmbroideryChallan.js exactly)
  const exportEmbroideryPDF = (stageRows) => {
    const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "A3" });
    const pageW = doc.internal.pageSize.getWidth();
    const pageH = doc.internal.pageSize.getHeight();
    const now = new Date();

    const marginX = 16;
    const tableW = pageW - marginX * 2;

    const totalLots = stageRows.length;
    const totalQty = stageRows.reduce((sum, r) => sum + (r.poQty || 0), 0);
    const redZoneLots = stageRows.filter((r) => r.agingDays > 5).length;
    const normalLots = totalLots - redZoneLots;

    const garmentMap = {};
    const embParties = {};

    stageRows.forEach((r) => {
      const g = r.garmentType || "Unassigned";
      const p = r.embParty || r.partyName || "Unassigned";
      const qty = r.poQty || 0;

      if (!garmentMap[g]) garmentMap[g] = { name: g, lots: 0, qty: 0 };
      garmentMap[g].lots += 1;
      garmentMap[g].qty += qty;

      if (!embParties[p]) embParties[p] = { name: p, lots: 0, qty: 0 };
      embParties[p].lots += 1;
      embParties[p].qty += qty;
    });

    const sortedGarments = Object.values(garmentMap).sort((a, b) => b.qty - a.qty);
    const sortedEmbParties = Object.values(embParties).sort((a, b) => b.qty - a.qty);

    doc.setFontSize(20);
    doc.setTextColor(0, 0, 0);
    doc.setFont("helvetica", "bold");
    doc.text("REPORT 2 - EMBROIDERY CHALLAN PRODUCTION REPORT", pageW / 2, 38, { align: "center" });

    doc.setFontSize(10);
    doc.setTextColor(50, 50, 50);
    doc.setFont("helvetica", "normal");
    doc.text(`Report Date: ${now.toLocaleDateString()} ${now.toLocaleTimeString()} | Factory Suite Pro`, pageW / 2, 54, { align: "center" });

    doc.setFillColor(239, 246, 255);
    doc.roundedRect(marginX, 64, tableW, 24, 6, 6, "F");
    doc.setDrawColor(191, 219, 254);
    doc.roundedRect(marginX, 64, tableW, 24, 6, 6, "D");

    doc.setFontSize(10.5);
    doc.setTextColor(0, 0, 0);
    doc.setFont("helvetica", "bold");
    const summaryText = `Total Records: ${totalLots}   |   Total Pcs / Qty: ${totalQty.toLocaleString()}   |   Garment Types: ${sortedGarments.length}   |   Embroidery Parties: ${sortedEmbParties.length}   |   Red Zone (>5d): ${redZoneLots}`;
    doc.text(summaryText, pageW / 2, 79, { align: "center" });

    const columns = [
      "S. No",
      "Lot No.",
      "Garment Type",
      "Style",
      "Fabric",
      "Brand",
      "Challan Qty",
      "Section (M/W/K)",
      "Season",
      "Party Name",
      "Direct Stitching",
      "Emb Party",
      "Challan Date",
      "Emb Status",
      "Pending Shade",
      "Emb Done Date",
      "Received Date",
      "Days",
      "HOD Remarks"
    ];

    const body = stageRows.map((r, idx) => [
      String(idx + 1),
      r.lotNumber,
      r.garmentType || "-",
      r.style || "-",
      r.fabric || "-",
      r.brand || "-",
      (r.poQty || 0).toLocaleString(),
      r.section || "-",
      r.season || "-",
      r.partyName || "-",
      r.directStitching || "No",
      r.embParty || r.partyName || "-",
      r.stageDate || "-",
      r.embStatus || "Pending",
      r.pendingShade || "-",
      r.embDoneDate || "-",
      r.receivedDate || "-",
      r.agingDays != null ? `${r.agingDays}` : "-",
      formatLatestRemark(r.hodRemarks, "-")
    ]);

    body.push([
      "",
      `TOTAL (${totalLots})`,
      "",
      "",
      "",
      "",
      totalQty.toLocaleString(),
      "",
      "",
      "",
      "",
      `${sortedEmbParties.length} Parties`,
      "",
      "",
      "",
      "",
      "",
      `${redZoneLots} Red Zone`,
      ""
    ]);

    const columnStyles = {
      0: { cellWidth: 26, halign: "center" },
      1: { cellWidth: 55, halign: "center", fontStyle: "bold" },
      2: { cellWidth: 65, halign: "center" },
      3: { cellWidth: 70, halign: "center" },
      4: { cellWidth: 60, halign: "center" },
      5: { cellWidth: 60, halign: "center" },
      6: { cellWidth: 50, halign: "center", fontStyle: "bold" },
      7: { cellWidth: 40, halign: "center" },
      8: { cellWidth: 45, halign: "center" },
      9: { cellWidth: 55, halign: "center" },
      10: { cellWidth: 45, halign: "center" },
      11: { cellWidth: 60, halign: "center" },
      12: { cellWidth: 55, halign: "center" },
      13: { cellWidth: 55, halign: "center" },
      14: { cellWidth: 70, halign: "left" },
      15: { cellWidth: 55, halign: "center" },
      16: { cellWidth: 55, halign: "center" },
      17: { cellWidth: 38, halign: "center" },
      18: { cellWidth: 110, halign: "left" }
    };

    const scaledColumnStyles = scaleColumnStyles(columnStyles, tableW);

    autoTable(doc, {
      head: [columns],
      body,
      startY: 96,
      tableWidth: tableW,
      margin: { top: 96, right: marginX, bottom: 25, left: marginX },
      theme: "grid",
      styles: {
        fontSize: 8.5,
        cellPadding: { top: 3.5, right: 2, bottom: 3.5, left: 2 },
        overflow: "linebreak",
        valign: "middle",
        halign: "center",
        textColor: [0, 0, 0],
        lineColor: [0, 0, 0],
        lineWidth: 0.5
      },
      headStyles: {
        fillColor: [15, 76, 129],
        textColor: [255, 255, 255],
        fontStyle: "bold",
        lineColor: [0, 0, 0],
        lineWidth: 0.5,
        halign: "center",
        fontSize: 9,
        valign: "middle",
        cellPadding: { top: 5, right: 2, bottom: 5, left: 2 }
      },
      alternateRowStyles: {
        fillColor: [248, 250, 252]
      },
      columnStyles: scaledColumnStyles,
      didParseCell: function (data) {
        if (data.section === "body") {
          data.cell.styles.textColor = [0, 0, 0];
          const rowIndex = data.row.index;
          const isTotalRow = rowIndex === body.length - 1;

          if (isTotalRow) {
            data.cell.styles.fontStyle = "bold";
            data.cell.styles.fillColor = [226, 232, 240];
            data.cell.styles.textColor = [0, 0, 0];
            return;
          }

          if (data.column.index === 17) {
            const rawVal = parseFloat(data.cell.raw);
            if (!isNaN(rawVal)) {
              data.cell.styles.fontStyle = "bold";
              if (rawVal > 5) {
                data.cell.styles.fillColor = [239, 68, 68];
                data.cell.styles.textColor = [255, 255, 255];
              } else {
                data.cell.styles.fillColor = [220, 252, 231];
                data.cell.styles.textColor = [21, 128, 61];
              }
            }
          }
        }
      },
      didDrawPage: () => {
        doc.setDrawColor(0, 0, 0);
        doc.setLineWidth(0.5);
        doc.roundedRect(8, 8, pageW - 16, pageH - 16, 2, 2, "S");
      }
    });

    const gBody = sortedGarments.map((item) => {
      const pct = totalQty > 0 ? ((item.qty / totalQty) * 100).toFixed(1) : "0.0";
      return [item.name, item.lots.toString(), `${pct}%`, item.qty.toLocaleString()];
    });
    gBody.push(["TOTAL", totalLots.toString(), "100.0%", totalQty.toLocaleString()]);

    const pBody = sortedEmbParties.map((item) => {
      const pct = totalQty > 0 ? ((item.qty / totalQty) * 100).toFixed(1) : "0.0";
      return [item.name, item.lots.toString(), `${pct}%`, item.qty.toLocaleString()];
    });
    pBody.push(["TOTAL", totalLots.toString(), "100.0%", totalQty.toLocaleString()]);

    const aBody = [
      ["<= 5 Days (On-Time / Normal)", normalLots.toString(), `${totalLots > 0 ? ((normalLots / totalLots) * 100).toFixed(1) : 0}%`, "—"],
      ["> 5 Days (Red Zone / Delayed)", redZoneLots.toString(), `${totalLots > 0 ? ((redZoneLots / totalLots) * 100).toFixed(1) : 0}%`, "—"],
      ["TOTAL LOTS", totalLots.toString(), "100.0%", totalQty.toLocaleString()]
    ];

    let summaryStartY = doc.lastAutoTable.finalY + 22;
    if (summaryStartY + 140 > pageH - 30) {
      doc.addPage();
      summaryStartY = 40;
    }

    doc.setFontSize(13);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(0, 0, 0);
    doc.text("EXECUTIVE SUMMARY & PRODUCTION BREAKDOWN", pageW / 2, summaryStartY + 4, { align: "center" });

    const sectionTitleY = summaryStartY + 26;
    const tableStartY = sectionTitleY + 6;
    const gap = 18;
    const colWidth = (tableW - 2 * gap) / 3;
    const col1X = marginX;
    const col2X = col1X + colWidth + gap;
    const col3X = col2X + colWidth + gap;

    const summaryColStyles = {
      0: { cellWidth: 140, halign: "center" },
      1: { cellWidth: 60, halign: "center" },
      2: { cellWidth: 70, halign: "center" },
      3: { cellWidth: 100, halign: "center" }
    };

    const scaledSummaryColStyles = scaleColumnStyles(summaryColStyles, colWidth);

    [
      { body: gBody, left: col1X, fill: [15, 118, 110], label: "1. GARMENT TYPE BREAKDOWN", header: "Garment Type" },
      { body: pBody, left: col2X, fill: [30, 64, 175], label: "2. EMBROIDERY PARTY ALLOCATION", header: "Emb Party" },
      { body: aBody, left: col3X, fill: [180, 83, 9], label: "3. SLA & DAYS AGING", header: "Aging Status" }
    ].forEach((tbl) => {
      autoTable(doc, {
        head: [[tbl.header, "Lots", "Share %", "Total Pieces"]],
        body: tbl.body,
        startY: tableStartY,
        tableWidth: colWidth,
        margin: { left: tbl.left, right: pageW - (tbl.left + colWidth) },
        theme: "grid",
        styles: {
          fontSize: 8.5,
          cellPadding: { top: 3.5, right: 2, bottom: 3.5, left: 2 },
          valign: "middle",
          halign: "center",
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          lineWidth: 0.3
        },
        headStyles: {
          fillColor: tbl.fill,
          textColor: [255, 255, 255],
          fontStyle: "bold",
          fontSize: 8.5,
          halign: "center"
        },
        columnStyles: scaledSummaryColStyles,
        didParseCell: (data) => {
          if (data.section === "body") {
            data.cell.styles.textColor = [0, 0, 0];
            if (data.row.index === tbl.body.length - 1) {
              data.cell.styles.fontStyle = "bold";
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });
    });

    doc.save(`Report_2_Embroidery_Variance_${now.toISOString().slice(0, 10)}.pdf`);
  };

  // 3. PRINTING CHALLAN REPORT PDF (Matches PrintingChallan.js exactly)
  const exportPrintingPDF = (stageRows) => {
    const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "A3" });
    const pageW = doc.internal.pageSize.getWidth();
    const pageH = doc.internal.pageSize.getHeight();
    const now = new Date();

    const marginX = 16;
    const tableW = pageW - marginX * 2;

    const totalLots = stageRows.length;
    const totalQty = stageRows.reduce((sum, r) => sum + (r.poQty || 0), 0);
    const redZoneLots = stageRows.filter((r) => r.agingDays > 5).length;
    const normalLots = totalLots - redZoneLots;

    const garmentMap = {};
    const printParties = {};

    stageRows.forEach((r) => {
      const g = r.garmentType || "Unassigned";
      const p = r.printingParty || r.partyName || "Unassigned";
      const qty = r.poQty || 0;

      if (!garmentMap[g]) garmentMap[g] = { name: g, lots: 0, qty: 0 };
      garmentMap[g].lots += 1;
      garmentMap[g].qty += qty;

      if (!printParties[p]) printParties[p] = { name: p, lots: 0, qty: 0 };
      printParties[p].lots += 1;
      printParties[p].qty += qty;
    });

    const sortedGarments = Object.values(garmentMap).sort((a, b) => b.qty - a.qty);
    const sortedPrintParties = Object.values(printParties).sort((a, b) => b.qty - a.qty);

    doc.setFontSize(20);
    doc.setTextColor(0, 0, 0);
    doc.setFont("helvetica", "bold");
    doc.text("REPORT 3 - PRINTING CHALLAN PRODUCTION REPORT", pageW / 2, 38, { align: "center" });

    doc.setFontSize(10);
    doc.setTextColor(50, 50, 50);
    doc.setFont("helvetica", "normal");
    doc.text(`Report Date: ${now.toLocaleDateString()} ${now.toLocaleTimeString()} | Factory Suite Pro`, pageW / 2, 54, { align: "center" });

    doc.setFillColor(239, 246, 255);
    doc.roundedRect(marginX, 64, tableW, 24, 6, 6, "F");
    doc.setDrawColor(191, 219, 254);
    doc.roundedRect(marginX, 64, tableW, 24, 6, 6, "D");

    doc.setFontSize(10.5);
    doc.setTextColor(0, 0, 0);
    doc.setFont("helvetica", "bold");
    const summaryText = `Total Records: ${totalLots}   |   Total Pcs / Qty: ${totalQty.toLocaleString()}   |   Garment Types: ${sortedGarments.length}   |   Printing Parties: ${sortedPrintParties.length}   |   Red Zone (>5d): ${redZoneLots}`;
    doc.text(summaryText, pageW / 2, 79, { align: "center" });

    const columns = [
      "S. No",
      "Lot No.",
      "Garment Type",
      "Style",
      "Fabric",
      "Brand",
      "Challan Qty",
      "Section (M/W/K)",
      "Season",
      "Party Name",
      "Direct Stitching",
      "Printing Party",
      "Challan Date",
      "Printing Status",
      "Pending Shade",
      "Printing Done Date",
      "Days",
      "HOD Remarks"
    ];

    const body = stageRows.map((r, idx) => [
      String(idx + 1),
      r.lotNumber,
      r.garmentType || "-",
      r.style || "-",
      r.fabric || "-",
      r.brand || "-",
      (r.poQty || 0).toLocaleString(),
      r.section || "-",
      r.season || "-",
      r.partyName || "-",
      r.directStitching || "No",
      r.printingParty || r.partyName || "-",
      r.stageDate || "-",
      r.printingStatus || "Pending",
      r.pendingShade || "-",
      r.printingDoneDate || "-",
      r.agingDays != null ? `${r.agingDays}` : "-",
      formatLatestRemark(r.hodRemarks, "-")
    ]);

    body.push([
      "",
      `TOTAL (${totalLots})`,
      "",
      "",
      "",
      "",
      totalQty.toLocaleString(),
      "",
      "",
      "",
      "",
      `${sortedPrintParties.length} Parties`,
      "",
      "",
      "",
      "",
      `${redZoneLots} Red Zone`,
      ""
    ]);

    const columnStyles = {
      0: { cellWidth: 26, halign: "center" },
      1: { cellWidth: 55, halign: "center", fontStyle: "bold" },
      2: { cellWidth: 65, halign: "center" },
      3: { cellWidth: 70, halign: "center" },
      4: { cellWidth: 60, halign: "center" },
      5: { cellWidth: 60, halign: "center" },
      6: { cellWidth: 50, halign: "center", fontStyle: "bold" },
      7: { cellWidth: 40, halign: "center" },
      8: { cellWidth: 45, halign: "center" },
      9: { cellWidth: 55, halign: "center" },
      10: { cellWidth: 45, halign: "center" },
      11: { cellWidth: 60, halign: "center" },
      12: { cellWidth: 55, halign: "center" },
      13: { cellWidth: 55, halign: "center" },
      14: { cellWidth: 70, halign: "left" },
      15: { cellWidth: 55, halign: "center" },
      16: { cellWidth: 38, halign: "center" },
      17: { cellWidth: 110, halign: "left" }
    };

    const scaledColumnStyles = scaleColumnStyles(columnStyles, tableW);

    autoTable(doc, {
      head: [columns],
      body,
      startY: 96,
      tableWidth: tableW,
      margin: { top: 96, right: marginX, bottom: 25, left: marginX },
      theme: "grid",
      styles: {
        fontSize: 8.5,
        cellPadding: { top: 3.5, right: 2, bottom: 3.5, left: 2 },
        overflow: "linebreak",
        valign: "middle",
        halign: "center",
        textColor: [0, 0, 0],
        lineColor: [0, 0, 0],
        lineWidth: 0.5
      },
      headStyles: {
        fillColor: [15, 76, 129],
        textColor: [255, 255, 255],
        fontStyle: "bold",
        lineColor: [0, 0, 0],
        lineWidth: 0.5,
        halign: "center",
        fontSize: 9,
        valign: "middle",
        cellPadding: { top: 5, right: 2, bottom: 5, left: 2 }
      },
      alternateRowStyles: {
        fillColor: [248, 250, 252]
      },
      columnStyles: scaledColumnStyles,
      didParseCell: function (data) {
        if (data.section === "body") {
          data.cell.styles.textColor = [0, 0, 0];
          const rowIndex = data.row.index;
          const isTotalRow = rowIndex === body.length - 1;

          if (isTotalRow) {
            data.cell.styles.fontStyle = "bold";
            data.cell.styles.fillColor = [226, 232, 240];
            data.cell.styles.textColor = [0, 0, 0];
            return;
          }

          if (data.column.index === 16) {
            const rawVal = parseFloat(data.cell.raw);
            if (!isNaN(rawVal)) {
              data.cell.styles.fontStyle = "bold";
              if (rawVal > 5) {
                data.cell.styles.fillColor = [239, 68, 68];
                data.cell.styles.textColor = [255, 255, 255];
              } else {
                data.cell.styles.fillColor = [220, 252, 231];
                data.cell.styles.textColor = [21, 128, 61];
              }
            }
          }
        }
      },
      didDrawPage: () => {
        doc.setDrawColor(0, 0, 0);
        doc.setLineWidth(0.5);
        doc.roundedRect(8, 8, pageW - 16, pageH - 16, 2, 2, "S");
      }
    });

    const gBody = sortedGarments.map((item) => {
      const pct = totalQty > 0 ? ((item.qty / totalQty) * 100).toFixed(1) : "0.0";
      return [item.name, item.lots.toString(), `${pct}%`, item.qty.toLocaleString()];
    });
    gBody.push(["TOTAL", totalLots.toString(), "100.0%", totalQty.toLocaleString()]);

    const pBody = sortedPrintParties.map((item) => {
      const pct = totalQty > 0 ? ((item.qty / totalQty) * 100).toFixed(1) : "0.0";
      return [item.name, item.lots.toString(), `${pct}%`, item.qty.toLocaleString()];
    });
    pBody.push(["TOTAL", totalLots.toString(), "100.0%", totalQty.toLocaleString()]);

    const aBody = [
      ["<= 5 Days (On-Time / Normal)", normalLots.toString(), `${totalLots > 0 ? ((normalLots / totalLots) * 100).toFixed(1) : 0}%`, "-"],
      ["> 5 Days (Red Zone / Delayed)", redZoneLots.toString(), `${totalLots > 0 ? ((redZoneLots / totalLots) * 100).toFixed(1) : 0}%`, "-"],
      ["TOTAL LOTS", totalLots.toString(), "100.0%", totalQty.toLocaleString()]
    ];

    let summaryStartY = doc.lastAutoTable.finalY + 22;
    if (summaryStartY + 140 > pageH - 30) {
      doc.addPage();
      summaryStartY = 40;
    }

    doc.setFontSize(13);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(0, 0, 0);
    doc.text("EXECUTIVE SUMMARY & PRODUCTION BREAKDOWN", pageW / 2, summaryStartY + 4, { align: "center" });

    const sectionTitleY = summaryStartY + 26;
    const tableStartY = sectionTitleY + 6;
    const gap = 18;
    const colWidth = (tableW - 2 * gap) / 3;
    const col1X = marginX;
    const col2X = col1X + colWidth + gap;
    const col3X = col2X + colWidth + gap;

    const summaryColStyles = {
      0: { cellWidth: 140, halign: "center" },
      1: { cellWidth: 60, halign: "center" },
      2: { cellWidth: 70, halign: "center" },
      3: { cellWidth: 100, halign: "center" }
    };

    const scaledSummaryColStyles = scaleColumnStyles(summaryColStyles, colWidth);

    [
      { body: gBody, left: col1X, fill: [15, 118, 110], label: "1. GARMENT TYPE BREAKDOWN", header: "Garment Type" },
      { body: pBody, left: col2X, fill: [30, 64, 175], label: "2. PRINTING PARTY ALLOCATION", header: "Printing Party" },
      { body: aBody, left: col3X, fill: [180, 83, 9], label: "3. SLA & DAYS AGING", header: "Aging Status" }
    ].forEach((tbl) => {
      autoTable(doc, {
        head: [[tbl.header, "Lots", "Share %", "Total Pieces"]],
        body: tbl.body,
        startY: tableStartY,
        tableWidth: colWidth,
        margin: { left: tbl.left, right: pageW - (tbl.left + colWidth) },
        theme: "grid",
        styles: {
          fontSize: 8.5,
          cellPadding: { top: 3.5, right: 2, bottom: 3.5, left: 2 },
          valign: "middle",
          halign: "center",
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          lineWidth: 0.3
        },
        headStyles: {
          fillColor: tbl.fill,
          textColor: [255, 255, 255],
          fontStyle: "bold",
          fontSize: 8.5,
          halign: "center"
        },
        columnStyles: scaledSummaryColStyles,
        didParseCell: (data) => {
          if (data.section === "body") {
            data.cell.styles.textColor = [0, 0, 0];
            if (data.row.index === tbl.body.length - 1) {
              data.cell.styles.fontStyle = "bold";
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });
    });

    doc.save(`Report_3_Printing_Variance_${now.toISOString().slice(0, 10)}.pdf`);
  };

  // 4. POST-EMB/PRINT CHALLAN REPORT PDF (Matches PendingIssue.js exactly)
  const exportPostEmbPrintPDF = (stageRows) => {
    const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "A3" });
    const pageW = doc.internal.pageSize.getWidth();
    const pageH = doc.internal.pageSize.getHeight();
    const now = new Date();

    const marginX = 16;
    const tableW = pageW - marginX * 2;

    const totalLots = stageRows.length;
    const totalQty = stageRows.reduce((sum, r) => sum + (r.poQty || 0), 0);
    const colorPendingLots = stageRows.filter((r) => r.pendingShade && !r.pendingShade.toLowerCase().includes("no col")).length;
    const repeatedLots = stageRows.filter((r) => r.isRepeatedLot).length;
    const redZoneLots = stageRows.filter((r) => r.agingDays > 2).length;
    const normalLots = totalLots - redZoneLots;

    const garmentMap = {};
    const seasonMap = {};
    const partyMap = {};

    stageRows.forEach((r) => {
      const g = r.garmentType || "Unassigned";
      const s = r.season || "Unassigned";
      const p = r.partyName || "Unassigned";
      const qty = r.poQty || 0;

      if (!garmentMap[g]) garmentMap[g] = { name: g, lots: 0, qty: 0 };
      garmentMap[g].lots += 1;
      garmentMap[g].qty += qty;

      if (!seasonMap[s]) seasonMap[s] = { name: s, lots: 0, qty: 0 };
      seasonMap[s].lots += 1;
      seasonMap[s].qty += qty;

      if (!partyMap[p]) partyMap[p] = { name: p, lots: 0, qty: 0 };
      partyMap[p].lots += 1;
      partyMap[p].qty += qty;
    });

    const sortedGarments = Object.values(garmentMap).sort((a, b) => b.qty - a.qty);
    const sortedSeasons = Object.values(seasonMap).sort((a, b) => b.qty - a.qty);
    const sortedParties = Object.values(partyMap).sort((a, b) => b.qty - a.qty);

    doc.setFontSize(20);
    doc.setTextColor(0, 0, 0);
    doc.setFont("helvetica", "bold");
    doc.text("REPORT 4 - POST-EMB/PRINT PENDING ISSUE TO STITCHING REPORT", pageW / 2, 38, { align: "center" });

    doc.setFontSize(10);
    doc.setTextColor(50, 50, 50);
    doc.setFont("helvetica", "normal");
    doc.text(`Report Date: ${now.toLocaleDateString()} ${now.toLocaleTimeString()} | Factory Suite Pro`, pageW / 2, 54, { align: "center" });

    doc.setFillColor(239, 246, 255);
    doc.roundedRect(marginX, 64, tableW, 24, 6, 6, "F");
    doc.setDrawColor(191, 219, 254);
    doc.roundedRect(marginX, 64, tableW, 24, 6, 6, "D");

    doc.setFontSize(10.5);
    doc.setTextColor(0, 0, 0);
    doc.setFont("helvetica", "bold");
    const summaryText = `Total Records: ${totalLots}   |   Total Pcs: ${totalQty.toLocaleString()}   |   Color Pending: ${colorPendingLots}   |   Repeated Lots: ${repeatedLots}   |   Red Zone (>2d): ${redZoneLots}`;
    doc.text(summaryText, pageW / 2, 79, { align: "center" });

    const columns = [
      "#",
      "Lot Number",
      "Garment Type",
      "Style",
      "Fabric",
      "Brand",
      "Total Pcs",
      "M/W/K",
      "Season",
      "Party Name",
      "Direct Stitching",
      "Cutting Date",
      "Emb/Print Date",
      "Days Pending",
      "Color Status",
      "Priority",
      "User Remarks"
    ];

    const body = stageRows.map((r, idx) => [
      String(idx + 1),
      r.isRepeatedLot ? `[Repeated] ${r.lotNumber}` : r.lotNumber,
      r.garmentType || "-",
      r.style || "-",
      r.fabric || "-",
      r.brand || "-",
      (r.poQty || 0).toLocaleString(),
      r.section || "-",
      r.season || "-",
      r.partyName || "-",
      r.directStitching || "No",
      r.cuttingDate || "-",
      r.embPrintDate || "-",
      r.agingDays != null ? `${r.agingDays} days` : "-",
      r.pendingShade || "No Colour Pending",
      r.priority || "Normal",
      formatLatestRemark(r.userRemarks, "-")
    ]);

    body.push([
      "",
      `TOTAL (${totalLots})`,
      "",
      "",
      "",
      "",
      totalQty.toLocaleString(),
      "",
      "",
      `${sortedParties.length} Parties`,
      "",
      "",
      "",
      `${redZoneLots} Red Zone`,
      `${colorPendingLots} Color Pending`,
      `${repeatedLots} Repeated`,
      ""
    ]);

    const columnStyles = {
      0: { cellWidth: 26, halign: "center" },
      1: { cellWidth: 60, halign: "center", fontStyle: "bold" },
      2: { cellWidth: 65, halign: "center" },
      3: { cellWidth: 70, halign: "center" },
      4: { cellWidth: 65, halign: "center" },
      5: { cellWidth: 60, halign: "center" },
      6: { cellWidth: 50, halign: "center", fontStyle: "bold" },
      7: { cellWidth: 40, halign: "center" },
      8: { cellWidth: 45, halign: "center" },
      9: { cellWidth: 60, halign: "center" },
      10: { cellWidth: 45, halign: "center" },
      11: { cellWidth: 55, halign: "center" },
      12: { cellWidth: 55, halign: "center" },
      13: { cellWidth: 50, halign: "center" },
      14: { cellWidth: 65, halign: "left" },
      15: { cellWidth: 45, halign: "center" },
      16: { cellWidth: 110, halign: "left" }
    };

    const scaledColumnStyles = scaleColumnStyles(columnStyles, tableW);

    autoTable(doc, {
      head: [columns],
      body,
      startY: 96,
      tableWidth: tableW,
      margin: { top: 96, right: marginX, bottom: 25, left: marginX },
      theme: "grid",
      styles: {
        fontSize: 8.5,
        cellPadding: { top: 3.5, right: 2, bottom: 3.5, left: 2 },
        overflow: "linebreak",
        valign: "middle",
        halign: "center",
        textColor: [0, 0, 0],
        lineColor: [0, 0, 0],
        lineWidth: 0.5
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
        cellPadding: { top: 5, right: 2, bottom: 5, left: 2 }
      },
      alternateRowStyles: {
        fillColor: [248, 250, 252]
      },
      columnStyles: scaledColumnStyles,
      didParseCell: function (data) {
        if (data.section === "body") {
          data.cell.styles.textColor = [0, 0, 0];
          const rowIndex = data.row.index;
          const isTotalRow = rowIndex === body.length - 1;

          if (isTotalRow) {
            data.cell.styles.fontStyle = "bold";
            data.cell.styles.fillColor = [226, 232, 240];
            data.cell.styles.textColor = [0, 0, 0];
            return;
          }

          if (data.column.index === 13) {
            const rawVal = parseFloat(data.cell.raw);
            if (!isNaN(rawVal)) {
              data.cell.styles.fontStyle = "bold";
              if (rawVal > 5) {
                data.cell.styles.fillColor = [239, 68, 68];
                data.cell.styles.textColor = [255, 255, 255];
              } else if (rawVal > 2) {
                data.cell.styles.fillColor = [254, 243, 199];
                data.cell.styles.textColor = [180, 83, 9];
              } else {
                data.cell.styles.fillColor = [220, 252, 231];
                data.cell.styles.textColor = [21, 128, 61];
              }
            }
          }
        }
      },
      didDrawPage: () => {
        doc.setDrawColor(0, 0, 0);
        doc.setLineWidth(0.5);
        doc.roundedRect(8, 8, pageW - 16, pageH - 16, 2, 2, "S");
      }
    });

    const gBody = sortedGarments.map((item) => {
      const pct = totalQty > 0 ? ((item.qty / totalQty) * 100).toFixed(1) : "0.0";
      return [item.name, item.lots.toString(), `${pct}%`, item.qty.toLocaleString()];
    });
    gBody.push(["TOTAL", totalLots.toString(), "100.0%", totalQty.toLocaleString()]);

    const sBody = sortedSeasons.map((item) => {
      const pct = totalQty > 0 ? ((item.qty / totalQty) * 100).toFixed(1) : "0.0";
      return [item.name, item.lots.toString(), `${pct}%`, item.qty.toLocaleString()];
    });
    sBody.push(["TOTAL", totalLots.toString(), "100.0%", totalQty.toLocaleString()]);

    const pBody = sortedParties.map((item) => {
      const pct = totalQty > 0 ? ((item.qty / totalQty) * 100).toFixed(1) : "0.0";
      return [item.name, item.lots.toString(), `${pct}%`, item.qty.toLocaleString()];
    });
    pBody.push(["TOTAL", totalLots.toString(), "100.0%", totalQty.toLocaleString()]);

    const aBody = [
      ["<= 2 Days (On-Time / Normal)", normalLots.toString(), `${totalLots > 0 ? ((normalLots / totalLots) * 100).toFixed(1) : 0}%`, "—"],
      ["> 2 Days (Red Zone / Delayed)", redZoneLots.toString(), `${totalLots > 0 ? ((redZoneLots / totalLots) * 100).toFixed(1) : 0}%`, "—"],
      ["TOTAL LOTS", totalLots.toString(), "100.0%", totalQty.toLocaleString()]
    ];

    let summaryStartY = doc.lastAutoTable.finalY + 22;
    if (summaryStartY + 140 > pageH - 30) {
      doc.addPage();
      summaryStartY = 40;
    }

    doc.setFontSize(13);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(0, 0, 0);
    doc.text("EXECUTIVE SUMMARY & PRODUCTION BREAKDOWN", pageW / 2, summaryStartY + 4, { align: "center" });

    const sectionTitleY = summaryStartY + 26;
    const tableStartY = sectionTitleY + 6;
    const gap = 14;
    const colWidth = (tableW - 3 * gap) / 4;
    const col1X = marginX;
    const col2X = col1X + colWidth + gap;
    const col3X = col2X + colWidth + gap;
    const col4X = col3X + colWidth + gap;

    const summaryColStyles = {
      0: { cellWidth: 110, halign: "center" },
      1: { cellWidth: 45, halign: "center" },
      2: { cellWidth: 68, halign: "center" },
      3: { cellWidth: 55, halign: "center" }
    };

    const scaledSummaryColStyles = scaleColumnStyles(summaryColStyles, colWidth);

    [
      { body: gBody, left: col1X, fill: [15, 118, 110], label: "1. GARMENT BREAKDOWN", header: "Garment Type" },
      { body: sBody, left: col2X, fill: [109, 40, 217], label: "2. SEASON BREAKDOWN", header: "Season" },
      { body: pBody, left: col3X, fill: [30, 64, 175], label: "3. PARTY BREAKDOWN", header: "Party Name" },
      { body: aBody, left: col4X, fill: [180, 83, 9], label: "4. AGING BREAKDOWN", header: "Aging Status" }
    ].forEach((tbl) => {
      autoTable(doc, {
        head: [[tbl.header, "Lots", "Share %", "Total Pieces"]],
        body: tbl.body,
        startY: tableStartY,
        tableWidth: colWidth,
        margin: { left: tbl.left, right: pageW - (tbl.left + colWidth) },
        theme: "grid",
        styles: {
          fontSize: 8.5,
          cellPadding: { top: 3.5, right: 2, bottom: 3.5, left: 2 },
          valign: "middle",
          halign: "center",
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          lineWidth: 0.3
        },
        headStyles: {
          fillColor: tbl.fill,
          textColor: [255, 255, 255],
          fontStyle: "bold",
          fontSize: 8.5,
          halign: "center"
        },
        columnStyles: scaledSummaryColStyles,
        didParseCell: (data) => {
          if (data.section === "body") {
            data.cell.styles.textColor = [0, 0, 0];
            if (data.row.index === tbl.body.length - 1) {
              data.cell.styles.fontStyle = "bold";
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });
    });

    doc.save(`Report_4_Post_EMB_Print_Variance_${now.toISOString().slice(0, 10)}.pdf`);
  };

  // 6. STITCHING WIP REPORT PDF (Exact 1:1 visual match with StitchingCompleted.js)
  const exportStitchingPDF = (stageRows) => {
    if (!stageRows || stageRows.length === 0) {
      alert("No data available to export.");
      return;
    }

    const exportData = stageRows;

    // Create PDF in landscape mode (A3)
    const doc = new jsPDF({
      orientation: 'landscape',
      unit: 'mm',
      format: 'a3'
    });

    // Colors matching StitchingCompleted.js exactly
    const headerColor = [15, 76, 129]; // Navy Blue
    const borderColor = [0, 0, 0]; // Black border
    const stageAnalysisColor = [59, 130, 246]; // Blue for stage analysis

    // Page dimensions
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const margin = 8;
    const contentWidth = pageWidth - (margin * 2);

    const isLotCompleted = (st) => {
      if (!st) return false;
      const s = String(st).toLowerCase().trim();
      return s === "completed" || s === "done" || s.startsWith("comp") || /^\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}/.test(s);
    };

    const normalizeSupervisorName = (name) => {
      if (!name || name.trim() === '') return 'Unassigned';
      const trimmedName = name.trim();
      return trimmedName.charAt(0).toUpperCase() + trimmedName.slice(1).toLowerCase();
    };

    const abbreviatePartyName = (partyName) => {
      if (!partyName || partyName.trim() === '') return 'N/A';
      const name = partyName.trim();
      if (name.toLowerCase().includes('mohit hosiery')) return 'MH';
      if (name.toLowerCase().includes('hosiery')) return name.split(' ')[0];
      const words = name.split(' ');
      if (words.length === 1) {
        return words[0].substring(0, 3).toUpperCase();
      } else if (words.length >= 2) {
        return (words[0].charAt(0) + words[1].charAt(0)).toUpperCase();
      }
      return name.substring(0, 4).toUpperCase();
    };

    const abbreviateMWKForPDF = (mwkValue) => {
      if (!mwkValue || typeof mwkValue !== 'string') return 'N/A';
      const value = mwkValue.trim().toLowerCase();
      if (value.includes('gents') || value === 'm' || value === 'mens') return 'M';
      if (value.includes('kids') || value === 'k') return 'K';
      if (value.includes('girls') || value === 'g' || value.includes('girlish')) return 'G';
      if (value.includes('women') || value.includes('womens') || value === 'w') return 'W';
      return value.charAt(0).toUpperCase();
    };

    const formatDateToDDMMYYForPDF = (dateString) => {
      if (!dateString || (typeof dateString !== 'string' && typeof dateString !== 'number')) {
        return '—';
      }
      try {
        const cleanVal = String(dateString).trim().replace(/^['"\s]+|['"\s]+$/g, '');
        if (!cleanVal || cleanVal === '-' || cleanVal === 'N/A' || cleanVal === '—') return '—';

        const isoMatch = cleanVal.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
        if (isoMatch) {
          const y = isoMatch[1].slice(-2);
          const m = isoMatch[2].padStart(2, '0');
          const d = isoMatch[3].padStart(2, '0');
          return `${d}/${m}/${y}`;
        }

        const parts = cleanVal.split(/[/\-.]/);
        if (parts.length === 3) {
          let day = parseInt(parts[0], 10);
          let month = parseInt(parts[1], 10);
          let year = parseInt(parts[2], 10);
          if (day > 1000) {
            const tmp = day; day = year; year = tmp;
          }
          if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
            const fullYear = year < 100 ? 2000 + year : year;
            return `${String(day).padStart(2, '0')}/${String(month).padStart(2, '0')}/${String(fullYear).slice(-2)}`;
          }
        }

        const date = new Date(cleanVal);
        if (!isNaN(date.getTime())) {
          const day = String(date.getDate()).padStart(2, '0');
          const month = String(date.getMonth() + 1).padStart(2, '0');
          const year = String(date.getFullYear()).slice(-2);
          return `${day}/${month}/${year}`;
        }
        return cleanVal;
      } catch {
        return String(dateString);
      }
    };

    const abbreviateSeason = (season) => {
      if (!season || season.trim() === '') return 'N/A';
      const seasonLower = season.trim().toLowerCase();
      if (seasonLower.includes('summer')) return 'S';
      if (seasonLower.includes('winter')) return 'W';
      if (seasonLower.includes('autumn')) return 'A';
      if (seasonLower.includes('spring')) return 'SP';
      return season.charAt(0).toUpperCase();
    };

    const formatDateOfIssue = (dateString) => {
      if (!dateString || dateString.trim() === '') return '—';
      return formatDateToDDMMYYForPDF(dateString);
    };

    const extractStageFromStatus = (wipRemarks) => {
      if (!wipRemarks || wipRemarks === 'N/A' || wipRemarks.trim() === '' || wipRemarks === '—') return 'Other';
      const remarks = wipRemarks.toLowerCase();
      if (remarks.includes('stitching done') && remarks.includes('overlock') && remarks.includes('folding')) {
        return 'Stitching Done Overlock and Folding Working';
      } else if (remarks.includes('stitching done') && remarks.includes('overlock')) {
        return 'Stitching Done Overlock Working';
      } else if (remarks.includes('stitching done') && remarks.includes('folding')) {
        return 'Stitching Done Folding Working';
      } else if (remarks.includes('stitching done')) {
        return 'Stitching Done';
      } else if (remarks.includes('overlock')) {
        return 'Overlock Working';
      } else if (remarks.includes('folding')) {
        return 'Folding Working';
      } else if (remarks.includes('cutting')) {
        return 'Cutting';
      } else if (remarks.includes('tailor working')) {
        return 'Tailor Working';
      } else if (remarks.includes('emb pending')) {
        return 'Emb Pending';
      } else {
        return 'On Stitching';
      }
    };

    const isStatusUpdatedToday = (wipStatus, status) => {
      if (status === 'Completed') {
        return true;
      }
      if (!wipStatus || wipStatus.trim() === '') {
        return false;
      }
      try {
        if (typeof wipStatus === 'string' && !wipStatus.startsWith('[')) {
          return false;
        }
        const statusArray = JSON.parse(wipStatus);
        if (!Array.isArray(statusArray) || statusArray.length === 0) {
          return false;
        }
        const sortedStatuses = [...statusArray].sort((a, b) => {
          const dateA = new Date(a.timestamp).getTime();
          const dateB = new Date(b.timestamp).getTime();
          return dateB - dateA;
        });
        const latestStatus = sortedStatuses[0];
        if (!latestStatus || !latestStatus.timestamp) {
          return false;
        }
        const today = new Date();
        const statusDate = new Date(latestStatus.timestamp);
        return statusDate.toDateString() === today.toDateString();
      } catch {
        return false;
      }
    };

    // Color coding legend helper matching StitchingCompleted.js
    const addColorLegend = (yPos) => {
      doc.setFontSize(8);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(0, 0, 0);
      doc.text('STITCHING DAYS COLOR CODING:', margin, yPos);

      const legendItems = [
        { color: [16, 185, 129], text: '1-6 Days: Good (Green)' },
        { color: [245, 158, 11], text: '7-15 Days: Average (Yellow)' },
        { color: [239, 68, 68], text: '15+ Days: Critical (Red)' }
      ];

      let legendX = margin + 70;
      legendItems.forEach((item) => {
        doc.setFillColor(...item.color);
        doc.rect(legendX, yPos - 3, 5, 5, 'F');
        doc.setTextColor(0, 0, 0);
        doc.setFont('helvetica', 'normal');
        doc.text(item.text, legendX + 7, yPos);
        legendX += 70;
      });

      return yPos + 8;
    };

    // Columns matching StitchingCompleted.js
    const activeCols = [
      { id: 'sr', label: 'Sr', baseWidth: 7 },
      { id: 'lotNo', label: 'Lot No', baseWidth: 17 },
      { id: 'fabric', label: 'Fabric', baseWidth: 22 },
      { id: 'garment', label: 'Garment', baseWidth: 18 },
      { id: 'style', label: 'Style', baseWidth: 18 },
      { id: 'brand', label: 'Brand', baseWidth: 15 },
      { id: 'party', label: 'Party', baseWidth: 20 },
      { id: 'supervisor', label: 'Supervisor', baseWidth: 18 },
      { id: 'season', label: 'Season', baseWidth: 13 },
      { id: 'mwk', label: 'M/W/K', baseWidth: 10 },
      { id: 'direct', label: 'Direct', baseWidth: 10 },
      { id: 'issueDate', label: 'Date of Issue', baseWidth: 16 },
      { id: 'days', label: 'Stitching Days', baseWidth: 16 },
      { id: 'totalPcs', label: 'Total PCS', baseWidth: 16 },
      { id: 'section', label: 'Section', baseWidth: 14 },
      { id: 'embPrint', label: 'Emb/Print', baseWidth: 16 },
      { id: 'wipStatus', label: 'WIP Status', baseWidth: 22 },
      { id: 'pintu', label: 'Pintu', baseWidth: 20 },
      { id: 'ea', label: 'EA', baseWidth: 20 },
      { id: 'lotStatus', label: 'Status', baseWidth: 14 }
    ];

    // Calculate scaled widths to perfectly fill contentWidth (404mm)
    const totalBaseWidth = activeCols.reduce((sum, c) => sum + (c.baseWidth || 15), 0);
    const columnWidths = {};
    const columnStyles = {};

    const tableHeaders = [
      activeCols.map((col, idx) => {
        const scaledWidth = Math.max(7, Math.round((col.baseWidth / totalBaseWidth) * contentWidth * 10) / 10);
        columnWidths[idx] = scaledWidth;
        columnStyles[idx] = {
          cellWidth: scaledWidth,
          halign: 'center',
          valign: 'middle'
        };
        return {
          content: col.label,
          styles: {
            fontStyle: 'bold',
            fillColor: headerColor,
            textColor: [255, 255, 255],
            cellWidth: scaledWidth,
            halign: 'center',
            fontSize: activeCols.length > 24 ? 7.5 : (activeCols.length > 18 ? 8.5 : 9.5),
            cellPadding: { top: 3, right: 1, bottom: 3, left: 1 }
          }
        };
      })
    ];

    // Identify list of supervisors to process
    const supSet = new Set();
    exportData.forEach(item => {
      const s = (item.supervisor || '').trim();
      if (s) supSet.add(normalizeSupervisorName(s));
      else supSet.add('Unassigned');
    });
    const supervisorsList = Array.from(supSet).sort();

    const cleanCellText = (text) => {
      if (text == null || text === '' || text === '-') return '—';
      if (text === 'N/A') return 'N/A';
      return formatLatestRemark(text, '—');
    };

    // Helper function to map data item to table cells
    const mapItemToRow = (item, rowIndex) => {
      const isCompleted = item.status === 'Completed' || isLotCompleted(item.completedStatus);
      const stitchingDays = typeof item.agingDays === 'number' ? item.agingDays : 0;
      const wipRemarks = formatLatestRemark(item.wipStatus, 'N/A');
      const embPrintDate = item.embPrintDate || item.embDate || '—';
      const abbreviatedParty = abbreviatePartyName(item.partyName);
      const abbreviatedSeason = abbreviateSeason(item.season);
      const totalPCS = item.poQty || item.totalPCS || 0;

      const pintuValue = formatLatestRemark(item.pintu, '—');
      const eaValue = formatLatestRemark(item.ea, '—');

      const isUpdatedToday = isStatusUpdatedToday(item.wipStatus, isCompleted ? 'Completed' : 'Pending');
      const wipDisplayValue = isCompleted
        ? 'Done'
        : (isUpdatedToday ? wipRemarks : 'Not updated');

      const lotStatus = isCompleted ? 'Completed' : 'Pending';
      const issueDate = formatDateOfIssue(item.dateOfIssue || item.stageDate);

      const formattedEmbPrintDate = embPrintDate !== '-' && embPrintDate !== '—' && embPrintDate !== 'N/A' ? formatDateToDDMMYYForPDF(embPrintDate) : '—';

      const rowBgColor = rowIndex % 2 === 0 ? [255, 255, 255] : [250, 250, 250];

      let stitchingDaysRGB = [240, 240, 240];
      if (stitchingDays <= 6) {
        stitchingDaysRGB = [220, 252, 231];
      } else if (stitchingDays <= 15) {
        stitchingDaysRGB = [254, 243, 199];
      } else {
        stitchingDaysRGB = [254, 226, 226];
      }

      let lotStatusColor = [254, 243, 199];
      if (isCompleted) {
        lotStatusColor = [220, 252, 231];
      }

      return activeCols.map((col, colIdx) => {
        const cellWidth = columnWidths[colIdx];
        const cellPad = { top: 2, right: 1, bottom: 2, left: 1 };
        const fontSz = activeCols.length > 24 ? 7.5 : (activeCols.length > 18 ? 8.5 : 9.5);

        if (col.id === 'sr') {
          return {
            content: (rowIndex + 1).toString(),
            styles: { cellWidth, fontSize: fontSz, halign: 'center', fillColor: rowBgColor, textColor: [0, 0, 0], fontStyle: 'normal', cellPadding: cellPad }
          };
        }
        if (col.id === 'lotNo') {
          return {
            content: cleanCellText(item.lotNumber),
            styles: { cellWidth, fontSize: Math.max(fontSz, 9.5), halign: 'center', fontStyle: 'bold', fillColor: rowBgColor, textColor: [0, 0, 0], cellPadding: cellPad }
          };
        }
        if (col.id === 'fabric') {
          return {
            content: cleanCellText(item.fabric),
            styles: { cellWidth, fontSize: fontSz, halign: 'center', fillColor: rowBgColor, fontStyle: 'bold', textColor: [0, 0, 0], cellPadding: cellPad }
          };
        }
        if (col.id === 'garment') {
          return {
            content: cleanCellText(item.garmentType),
            styles: { cellWidth, fontSize: fontSz, halign: 'center', fillColor: rowBgColor, fontStyle: 'bold', textColor: [0, 0, 0], cellPadding: cellPad }
          };
        }
        if (col.id === 'style') {
          return {
            content: cleanCellText(item.style),
            styles: { cellWidth, fontSize: fontSz, halign: 'center', fillColor: rowBgColor, fontStyle: 'bold', textColor: [0, 0, 0], cellPadding: cellPad }
          };
        }
        if (col.id === 'brand') {
          return {
            content: cleanCellText(item.brand),
            styles: { cellWidth, fontSize: fontSz, halign: 'center', fillColor: rowBgColor, fontStyle: 'bold', textColor: [0, 0, 0], cellPadding: cellPad }
          };
        }
        if (col.id === 'party') {
          return {
            content: abbreviatedParty,
            styles: { cellWidth, fontSize: fontSz + 1, halign: 'center', fontStyle: 'bold', fillColor: rowBgColor, textColor: [0, 0, 0], cellPadding: cellPad }
          };
        }
        if (col.id === 'supervisor') {
          return {
            content: cleanCellText(normalizeSupervisorName(item.supervisor)),
            styles: { cellWidth, fontSize: fontSz, halign: 'center', fillColor: rowBgColor, fontStyle: 'bold', textColor: [0, 0, 0], cellPadding: cellPad }
          };
        }
        if (col.id === 'season') {
          return {
            content: abbreviatedSeason,
            styles: { cellWidth, fontSize: fontSz, halign: 'center', fillColor: rowBgColor, textColor: [0, 0, 0], cellPadding: cellPad }
          };
        }
        if (col.id === 'mwk') {
          return {
            content: abbreviateMWKForPDF(item.mwk || item.section),
            styles: { cellWidth, fontSize: fontSz, halign: 'center', fillColor: rowBgColor, fontStyle: 'bold', textColor: [0, 0, 0], cellPadding: cellPad }
          };
        }
        if (col.id === 'direct') {
          return {
            content: item.directStitching ? (String(item.directStitching).toLowerCase() === 'yes' ? 'Y' : 'N') : 'N/A',
            styles: { cellWidth, fontSize: fontSz, halign: 'center', fillColor: rowBgColor, fontStyle: 'bold', textColor: [0, 0, 0], cellPadding: cellPad }
          };
        }
        if (col.id === 'issueDate') {
          return {
            content: cleanCellText(issueDate),
            styles: { cellWidth, fontSize: fontSz + 1, halign: 'center', fontStyle: 'bold', fillColor: rowBgColor, textColor: [0, 0, 0], cellPadding: cellPad }
          };
        }
        if (col.id === 'days') {
          return {
            content: stitchingDays.toString(),
            styles: { cellWidth, fontSize: fontSz + 1, halign: 'center', fontStyle: 'bold', fillColor: stitchingDaysRGB, textColor: [0, 0, 0], cellPadding: cellPad }
          };
        }
        if (col.id === 'totalPcs') {
          return {
            content: totalPCS > 0 ? totalPCS.toLocaleString() : 'N/A',
            styles: { cellWidth, fontSize: Math.max(fontSz, 9.5), halign: 'center', fontStyle: 'bold', fillColor: rowBgColor, textColor: [0, 0, 0], cellPadding: cellPad }
          };
        }
        if (col.id === 'section') {
          const mwkAbbr = abbreviateMWKForPDF(item.mwk);
          const derivedSection = mwkAbbr === 'M' ? 'GENTS' : mwkAbbr === 'W' ? 'WOMEN' : mwkAbbr === 'K' ? 'KIDS' : mwkAbbr === 'G' ? 'GIRLS' : mwkAbbr === 'B' ? 'BOYS' : '';
          const sectionVal = derivedSection || (item.section && item.section !== 'N/A' && item.section !== '—' ? item.section : '—');
          return {
            content: cleanCellText(sectionVal),
            styles: { cellWidth, fontSize: fontSz, halign: 'center', fontStyle: 'bold', fillColor: rowBgColor, textColor: [0, 0, 0], cellPadding: cellPad }
          };
        }
        if (col.id === 'embPrint') {
          return {
            content: cleanCellText(formattedEmbPrintDate),
            styles: { cellWidth, fontSize: fontSz, halign: 'center', fillColor: rowBgColor, fontStyle: formattedEmbPrintDate !== '-' && formattedEmbPrintDate !== '—' ? 'bold' : 'normal', textColor: [0, 0, 0], cellPadding: cellPad }
          };
        }
        if (col.id === 'wipStatus') {
          return {
            content: cleanCellText(wipDisplayValue),
            styles: { cellWidth, fontSize: fontSz, halign: 'center', fillColor: !isCompleted && !isUpdatedToday ? [255, 235, 235] : rowBgColor, fontStyle: isCompleted ? 'bold' : (wipRemarks !== 'N/A' ? 'bold' : 'normal'), textColor: [0, 0, 0], cellPadding: cellPad }
          };
        }
        if (col.id === 'pintu') {
          return {
            content: cleanCellText(pintuValue),
            styles: { cellWidth, fontSize: fontSz, halign: 'center', fontStyle: 'bold', fillColor: rowBgColor, textColor: [0, 0, 0], cellPadding: cellPad }
          };
        }
        if (col.id === 'ea') {
          return {
            content: cleanCellText(eaValue),
            styles: { cellWidth, fontSize: fontSz, halign: 'center', fontStyle: 'bold', fillColor: rowBgColor, textColor: [0, 0, 0], cellPadding: cellPad }
          };
        }
        if (col.id === 'lotStatus') {
          return {
            content: cleanCellText(lotStatus),
            styles: { cellWidth, fontSize: Math.max(6.5, fontSz - 1), halign: 'center', fontStyle: 'bold', fillColor: lotStatusColor, textColor: [0, 0, 0], cellPadding: cellPad }
          };
        }

        return {
          content: '—',
          styles: { cellWidth, fontSize: fontSz, halign: 'center', fillColor: rowBgColor, textColor: [0, 0, 0], cellPadding: cellPad }
        };
      });
    };

    // Helper function to render 3 side-by-side executive tables matching StitchingCompleted.js
    const renderThreeSideBySideSummary = (dataset, titlePrefix, startY) => {
      const stageAnalysis = {};
      const garmentAnalysis = {};
      let greenLots = 0, greenPcs = 0;
      let yellowLots = 0, yellowPcs = 0;
      let redLots = 0, redPcs = 0;

      dataset.forEach(item => {
        const pcs = item.poQty || item.totalPCS || 0;
        const stitchingDays = typeof item.agingDays === 'number' ? item.agingDays : 0;
        const isCompleted = item.status === 'Completed' || isLotCompleted(item.completedStatus);

        let stage = 'Other';
        if (isCompleted) {
          stage = 'Stitching Completed';
        } else {
          const wipRemarks = formatLatestRemark(item.wipStatus, 'N/A');
          stage = extractStageFromStatus(wipRemarks);
        }

        if (!stageAnalysis[stage]) {
          stageAnalysis[stage] = { lots: 0, pcs: 0, notUpdatedLots: 0 };
        }
        stageAnalysis[stage].lots += 1;
        stageAnalysis[stage].pcs += pcs;
        if (!isStatusUpdatedToday(item.wipStatus, isCompleted ? 'Completed' : 'Pending')) {
          stageAnalysis[stage].notUpdatedLots += 1;
        }

        const garment = (item.garmentType || 'Unassigned').trim() || 'Unassigned';
        if (!garmentAnalysis[garment]) {
          garmentAnalysis[garment] = { lots: 0, pcs: 0 };
        }
        garmentAnalysis[garment].lots += 1;
        garmentAnalysis[garment].pcs += pcs;

        if (stitchingDays <= 6) {
          greenLots += 1;
          greenPcs += pcs;
        } else if (stitchingDays <= 15) {
          yellowLots += 1;
          yellowPcs += pcs;
        } else {
          redLots += 1;
          redPcs += pcs;
        }
      });

      const stageArray = Object.entries(stageAnalysis)
        .map(([stage, data]) => ({
          stage,
          lots: data.lots,
          pcs: data.pcs,
          notUpdatedLots: data.notUpdatedLots
        }))
        .sort((a, b) => b.lots - a.lots);

      const totalStageLots = stageArray.reduce((sum, item) => sum + item.lots, 0);
      const totalStagePCS = stageArray.reduce((sum, item) => sum + item.pcs, 0);
      const totalStageNotUpdated = stageArray.reduce((sum, item) => sum + item.notUpdatedLots, 0);

      const garmentArray = Object.entries(garmentAnalysis)
        .map(([garment, data]) => ({
          garment,
          lots: data.lots,
          pcs: data.pcs
        }))
        .sort((a, b) => b.lots - a.lots);

      const totalGarmentLots = garmentArray.reduce((sum, item) => sum + item.lots, 0);
      const totalGarmentPCS = garmentArray.reduce((sum, item) => sum + item.pcs, 0);

      const totalDaysLots = greenLots + yellowLots + redLots;
      const totalDaysPCS = greenPcs + yellowPcs + redPcs;
      const greenPercent = totalDaysLots > 0 ? Math.round((greenLots / totalDaysLots) * 100) : 0;
      const yellowPercent = totalDaysLots > 0 ? Math.round((yellowLots / totalDaysLots) * 100) : 0;
      const redPercent = totalDaysLots > 0 ? Math.round((redLots / totalDaysLots) * 100) : 0;

      const gap = 4.5;
      const totalUsableWidth = contentWidth - (gap * 2);
      const width1 = totalUsableWidth * 0.37;
      const width2 = totalUsableWidth * 0.31;
      const width3 = totalUsableWidth * 0.32;

      let summaryStartY = startY;
      if (summaryStartY > pageHeight - 80) {
        doc.addPage();
        summaryStartY = 35;
      }

      doc.setFillColor(255, 255, 255);
      doc.rect(margin - 2, summaryStartY - 8, contentWidth + 4, 14, 'F');

      const t1Center = margin + (width1 / 2);
      const t2Center = (margin + width1 + gap) + (width2 / 2);
      const t3Center = (margin + width1 + gap + width2 + gap) + (width3 / 2);

      doc.setFontSize(10);
      doc.setFont('times', 'bold');
      doc.setTextColor(0, 0, 0);
      doc.text(`${titlePrefix.toUpperCase()} - STAGE ANALYSIS`, t1Center, summaryStartY, { align: 'center' });
      doc.text(`${titlePrefix.toUpperCase()} - GARMENT SUMMARY`, t2Center, summaryStartY, { align: 'center' });
      doc.text(`${titlePrefix.toUpperCase()} - DAYS-WISE AGING`, t3Center, summaryStartY, { align: 'center' });

      const stageBody = stageArray.map(item => {
        const percentage = totalStageLots > 0 ? Math.round((item.lots / totalStageLots) * 100) : 0;
        return [
          { content: item.stage, styles: { halign: 'left', fontSize: 8, cellPadding: { top: 2.2, right: 1.5, bottom: 2.2, left: 3 }, fontStyle: 'bold', fillColor: [240, 249, 255], textColor: [0, 0, 0] } },
          { content: item.lots.toString(), styles: { halign: 'center', fontSize: 8.5, fontStyle: 'bold', textColor: [0, 0, 0], cellPadding: { top: 2.2, right: 1.5, bottom: 2.2, left: 1.5 } } },
          { content: `${percentage}%`, styles: { halign: 'center', fontSize: 8, fontStyle: 'bold', textColor: [0, 0, 0], cellPadding: { top: 2.2, right: 1.5, bottom: 2.2, left: 1.5 } } },
          { content: item.pcs.toLocaleString(), styles: { halign: 'center', fontSize: 8.5, fontStyle: 'bold', textColor: [0, 0, 0], cellPadding: { top: 2.2, right: 1.5, bottom: 2.2, left: 1.5 } } },
          { content: item.notUpdatedLots > 0 ? item.notUpdatedLots.toString() : '-', styles: { halign: 'center', fontSize: 8.5, fontStyle: 'bold', textColor: [0, 0, 0], fillColor: item.notUpdatedLots > 0 ? [255, 235, 235] : [240, 249, 255], cellPadding: { top: 2.2, right: 1.5, bottom: 2.2, left: 1.5 } } }
        ];
      });

      stageBody.push([
        { content: 'TOTAL', styles: { halign: 'left', fontSize: 9, fontStyle: 'bold', fillColor: [225, 239, 255], textColor: [0, 0, 0], cellPadding: { top: 3, right: 1.5, bottom: 3, left: 3 } } },
        { content: totalStageLots.toString(), styles: { halign: 'center', fontSize: 9.5, fontStyle: 'bold', fillColor: [225, 239, 255], textColor: [0, 0, 0], cellPadding: { top: 3, right: 1.5, bottom: 3, left: 1.5 } } },
        { content: '100%', styles: { halign: 'center', fontSize: 8.5, fontStyle: 'bold', fillColor: [225, 239, 255], textColor: [0, 0, 0], cellPadding: { top: 3, right: 1.5, bottom: 3, left: 1.5 } } },
        { content: totalStagePCS.toLocaleString(), styles: { halign: 'center', fontSize: 9.5, fontStyle: 'bold', textColor: [0, 0, 0], fillColor: [225, 239, 255], cellPadding: { top: 3, right: 1.5, bottom: 3, left: 1.5 } } },
        { content: totalStageNotUpdated > 0 ? totalStageNotUpdated.toString() : '-', styles: { halign: 'center', fontSize: 9, fontStyle: 'bold', textColor: [0, 0, 0], fillColor: totalStageNotUpdated > 0 ? [255, 235, 235] : [225, 239, 255], cellPadding: { top: 3, right: 1.5, bottom: 3, left: 1.5 } } }
      ]);

      const stageColumnWidths = [width1 * 0.42, width1 * 0.13, width1 * 0.11, width1 * 0.18, width1 * 0.18];

      const garmentBody = garmentArray.map(item => [
        { content: item.garment, styles: { halign: 'left', fontSize: 8, cellPadding: { top: 2.2, right: 2, bottom: 2.2, left: 4 }, fontStyle: 'bold', fillColor: [240, 249, 255], textColor: [0, 0, 0] } },
        { content: item.lots.toString(), styles: { halign: 'center', fontSize: 8.5, fontStyle: 'bold', textColor: [0, 0, 0], cellPadding: { top: 2.2, right: 1.5, bottom: 2.2, left: 1.5 } } },
        { content: item.pcs.toLocaleString(), styles: { halign: 'center', fontSize: 8.5, fontStyle: 'bold', textColor: [0, 0, 0], cellPadding: { top: 2.2, right: 1.5, bottom: 2.2, left: 1.5 } } }
      ]);

      garmentBody.push([
        { content: 'TOTAL', styles: { halign: 'left', fontSize: 9, fontStyle: 'bold', fillColor: [225, 239, 255], textColor: [0, 0, 0], cellPadding: { top: 3, right: 2, bottom: 3, left: 4 } } },
        { content: totalGarmentLots.toString(), styles: { halign: 'center', fontSize: 9.5, fontStyle: 'bold', fillColor: [225, 239, 255], textColor: [0, 0, 0], cellPadding: { top: 3, right: 1.5, bottom: 3, left: 1.5 } } },
        { content: totalGarmentPCS.toLocaleString(), styles: { halign: 'center', fontSize: 9.5, fontStyle: 'bold', textColor: [0, 0, 0], fillColor: [225, 239, 255], cellPadding: { top: 3, right: 1.5, bottom: 3, left: 1.5 } } }
      ]);

      const garmentColumnWidths = [width2 * 0.50, width2 * 0.25, width2 * 0.25];

      const daysBody = [
        [
          { content: 'Green (1-6 Days)', styles: { halign: 'left', fontSize: 8, cellPadding: { top: 2.2, right: 2, bottom: 2.2, left: 4 }, fontStyle: 'bold', fillColor: [220, 252, 231], textColor: [0, 0, 0] } },
          { content: greenLots.toString(), styles: { halign: 'center', fontSize: 8.5, fontStyle: 'bold', fillColor: [220, 252, 231], textColor: [0, 0, 0], cellPadding: { top: 2.2, right: 1.5, bottom: 2.2, left: 1.5 } } },
          { content: `${greenPercent}%`, styles: { halign: 'center', fontSize: 8, fontStyle: 'bold', fillColor: [220, 252, 231], textColor: [0, 0, 0], cellPadding: { top: 2.2, right: 1.5, bottom: 2.2, left: 1.5 } } },
          { content: greenPcs.toLocaleString(), styles: { halign: 'center', fontSize: 8.5, fontStyle: 'bold', textColor: [0, 0, 0], cellPadding: { top: 2.2, right: 1.5, bottom: 2.2, left: 1.5 } } }
        ],
        [
          { content: 'Yellow (7-15 Days)', styles: { halign: 'left', fontSize: 8, cellPadding: { top: 2.2, right: 2, bottom: 2.2, left: 4 }, fontStyle: 'bold', fillColor: [254, 243, 199], textColor: [0, 0, 0] } },
          { content: yellowLots.toString(), styles: { halign: 'center', fontSize: 8.5, fontStyle: 'bold', fillColor: [254, 243, 199], textColor: [0, 0, 0], cellPadding: { top: 2.2, right: 1.5, bottom: 2.2, left: 1.5 } } },
          { content: `${yellowPercent}%`, styles: { halign: 'center', fontSize: 8, fontStyle: 'bold', fillColor: [254, 243, 199], textColor: [0, 0, 0], cellPadding: { top: 2.2, right: 1.5, bottom: 2.2, left: 1.5 } } },
          { content: yellowPcs.toLocaleString(), styles: { halign: 'center', fontSize: 8.5, fontStyle: 'bold', textColor: [0, 0, 0], cellPadding: { top: 2.2, right: 1.5, bottom: 2.2, left: 1.5 } } }
        ],
        [
          { content: 'Red (15+ Days)', styles: { halign: 'left', fontSize: 8, cellPadding: { top: 2.2, right: 2, bottom: 2.2, left: 4 }, fontStyle: 'bold', fillColor: [254, 226, 226], textColor: [0, 0, 0] } },
          { content: redLots.toString(), styles: { halign: 'center', fontSize: 8.5, fontStyle: 'bold', fillColor: [254, 226, 226], textColor: [0, 0, 0], cellPadding: { top: 2.2, right: 1.5, bottom: 2.2, left: 1.5 } } },
          { content: `${redPercent}%`, styles: { halign: 'center', fontSize: 8, fontStyle: 'bold', fillColor: [254, 226, 226], textColor: [0, 0, 0], cellPadding: { top: 2.2, right: 1.5, bottom: 2.2, left: 1.5 } } },
          { content: redPcs.toLocaleString(), styles: { halign: 'center', fontSize: 8.5, fontStyle: 'bold', textColor: [0, 0, 0], cellPadding: { top: 2.2, right: 1.5, bottom: 2.2, left: 1.5 } } }
        ],
        [
          { content: 'TOTAL', styles: { halign: 'left', fontSize: 9, fontStyle: 'bold', fillColor: [225, 239, 255], textColor: [0, 0, 0], cellPadding: { top: 3, right: 2, bottom: 3, left: 4 } } },
          { content: totalDaysLots.toString(), styles: { halign: 'center', fontSize: 9.5, fontStyle: 'bold', fillColor: [225, 239, 255], textColor: [0, 0, 0], cellPadding: { top: 3, right: 1.5, bottom: 3, left: 1.5 } } },
          { content: '100%', styles: { halign: 'center', fontSize: 8.5, fontStyle: 'bold', fillColor: [225, 239, 255], textColor: [0, 0, 0], cellPadding: { top: 3, right: 1.5, bottom: 3, left: 1.5 } } },
          { content: totalDaysPCS.toLocaleString(), styles: { halign: 'center', fontSize: 9.5, fontStyle: 'bold', textColor: [0, 0, 0], fillColor: [225, 239, 255], cellPadding: { top: 3, right: 1.5, bottom: 3, left: 1.5 } } }
        ]
      ];

      const daysColumnWidths = [width3 * 0.40, width3 * 0.17, width3 * 0.16, width3 * 0.27];

      // Render Table 1: Stage Analysis
      autoTable(doc, {
        startY: summaryStartY + 7,
        head: [[
          { content: 'WORK STAGE', styles: { halign: 'center', fontStyle: 'bold', cellWidth: stageColumnWidths[0], fillColor: stageAnalysisColor, textColor: [255, 255, 255] } },
          { content: 'LOTS', styles: { halign: 'center', fontStyle: 'bold', cellWidth: stageColumnWidths[1], fillColor: stageAnalysisColor, textColor: [255, 255, 255] } },
          { content: '%', styles: { halign: 'center', fontStyle: 'bold', cellWidth: stageColumnWidths[2], fillColor: stageAnalysisColor, textColor: [255, 255, 255] } },
          { content: 'TOTAL PCS', styles: { halign: 'center', fontStyle: 'bold', cellWidth: stageColumnWidths[3], fillColor: stageAnalysisColor, textColor: [255, 255, 255] } },
          { content: 'NOT UPDATED', styles: { halign: 'center', fontStyle: 'bold', cellWidth: stageColumnWidths[4], fillColor: [220, 38, 38], textColor: [255, 255, 255] } }
        ]],
        body: stageBody.map(row => row.map((cell, colIndex) => ({ ...cell, styles: { ...cell.styles, cellWidth: stageColumnWidths[colIndex] } }))),
        theme: 'grid',
        headStyles: { fillColor: stageAnalysisColor, textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 8, cellPadding: { top: 2.5, right: 1, bottom: 2.5, left: 1 }, lineWidth: 0.5, lineColor: stageAnalysisColor, halign: 'center', valign: 'middle' },
        bodyStyles: { fontSize: 8, cellPadding: { top: 2.2, right: 1, bottom: 2.2, left: 1 }, lineWidth: 0.3, lineColor: [220, 220, 220], textColor: [0, 0, 0], font: 'helvetica', valign: 'middle' },
        margin: { top: 35, left: margin, right: pageWidth - (margin + width1) },
        tableWidth: width1,
        showHead: 'everyPage',
        showFoot: false,
        pageBreak: 'auto',
        rowPageBreak: 'avoid'
      });
      const stageFinalY = doc.lastAutoTable ? doc.lastAutoTable.finalY : summaryStartY + 50;

      // Render Table 2: Garment Summary
      autoTable(doc, {
        startY: summaryStartY + 7,
        head: [[
          { content: 'GARMENT TYPE', styles: { halign: 'center', fontStyle: 'bold', cellWidth: garmentColumnWidths[0], fillColor: [15, 76, 129], textColor: [255, 255, 255] } },
          { content: 'TOTAL LOTS', styles: { halign: 'center', fontStyle: 'bold', cellWidth: garmentColumnWidths[1], fillColor: [15, 76, 129], textColor: [255, 255, 255] } },
          { content: 'TOTAL PCS', styles: { halign: 'center', fontStyle: 'bold', cellWidth: garmentColumnWidths[2], fillColor: [15, 76, 129], textColor: [255, 255, 255] } }
        ]],
        body: garmentBody.map(row => row.map((cell, colIndex) => ({ ...cell, styles: { ...cell.styles, cellWidth: garmentColumnWidths[colIndex] } }))),
        theme: 'grid',
        headStyles: { fillColor: [15, 76, 129], textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 8, cellPadding: { top: 2.5, right: 1, bottom: 2.5, left: 1 }, lineWidth: 0.5, lineColor: [15, 76, 129], halign: 'center', valign: 'middle' },
        bodyStyles: { fontSize: 8, cellPadding: { top: 2.2, right: 1, bottom: 2.2, left: 1 }, lineWidth: 0.3, lineColor: [220, 220, 220], textColor: [0, 0, 0], font: 'helvetica', valign: 'middle' },
        margin: { top: 35, left: margin + width1 + gap, right: pageWidth - (margin + width1 + gap + width2) },
        tableWidth: width2,
        showHead: 'everyPage',
        showFoot: false,
        pageBreak: 'auto',
        rowPageBreak: 'avoid'
      });
      const garmentFinalY = doc.lastAutoTable ? doc.lastAutoTable.finalY : summaryStartY + 50;

      // Render Table 3: Days-Wise Aging Summary
      autoTable(doc, {
        startY: summaryStartY + 7,
        head: [[
          { content: 'DAYS CATEGORY', styles: { halign: 'center', fontStyle: 'bold', cellWidth: daysColumnWidths[0], fillColor: [59, 130, 246], textColor: [255, 255, 255] } },
          { content: 'LOTS', styles: { halign: 'center', fontStyle: 'bold', cellWidth: daysColumnWidths[1], fillColor: [59, 130, 246], textColor: [255, 255, 255] } },
          { content: '%', styles: { halign: 'center', fontStyle: 'bold', cellWidth: daysColumnWidths[2], fillColor: [59, 130, 246], textColor: [255, 255, 255] } },
          { content: 'TOTAL PCS', styles: { halign: 'center', fontStyle: 'bold', cellWidth: daysColumnWidths[3], fillColor: [59, 130, 246], textColor: [255, 255, 255] } }
        ]],
        body: daysBody.map(row => row.map((cell, colIndex) => ({ ...cell, styles: { ...cell.styles, cellWidth: daysColumnWidths[colIndex] } }))),
        theme: 'grid',
        headStyles: { fillColor: [59, 130, 246], textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 8, cellPadding: { top: 2.5, right: 1, bottom: 2.5, left: 1 }, lineWidth: 0.5, lineColor: [59, 130, 246], halign: 'center', valign: 'middle' },
        bodyStyles: { fontSize: 8, cellPadding: { top: 2.2, right: 1, bottom: 2.2, left: 1 }, lineWidth: 0.3, lineColor: [220, 220, 220], textColor: [0, 0, 0], font: 'helvetica', valign: 'middle' },
        margin: { top: 35, left: margin + width1 + gap + width2 + gap, right: margin },
        tableWidth: width3,
        showHead: 'everyPage',
        showFoot: false,
        pageBreak: 'auto',
        rowPageBreak: 'avoid'
      });
      const daysFinalY = doc.lastAutoTable ? doc.lastAutoTable.finalY : summaryStartY + 50;

      return Math.max(stageFinalY, garmentFinalY, daysFinalY);
    };

    // Helper function to render Overall Supervisor Workload & Manpower / Attendance Summary Table matching StitchingCompleted.js
    const renderSupervisorWorkloadSummary = (dataset) => {
      doc.addPage();
      doc.setFillColor(255, 255, 255);
      doc.rect(0, 0, pageWidth, 30, 'F');

      const totalPCSAll = dataset.reduce((sum, item) => sum + (item.poQty || item.totalPCS || 0), 0);
      const totalLotsAll = dataset.length;
      const totalCompletedAll = dataset.filter(item => item.status === 'Completed' || isLotCompleted(item.completedStatus)).length;
      const totalPendingAll = totalLotsAll - totalCompletedAll;

      doc.setFontSize(18);
      doc.setTextColor(0, 0, 0);
      doc.setFont('Times New Roman', 'bold');
      doc.text('SUPERVISOR WORKLOAD & MANPOWER / ATTENDANCE SUMMARY', pageWidth / 2, 14, { align: 'center' });

      doc.setFontSize(11);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(0, 0, 0);

      const summaryDate = new Date();
      const summaryReportDate = `${String(summaryDate.getDate()).padStart(2, '0')}/${String(summaryDate.getMonth() + 1).padStart(2, '0')}/${String(summaryDate.getFullYear()).slice(-2)}`;
      doc.text(`Report Date: ${summaryReportDate}`, margin, 25);
      doc.text(`Total Lots: ${totalLotsAll} | Total PCS: ${totalPCSAll.toLocaleString()} | Completed: ${totalCompletedAll} | Pending: ${totalPendingAll}`,
        pageWidth / 2, 25, { align: 'center' });
      doc.text(`All Supervisors`, pageWidth - margin, 25, { align: 'right' });

      const supervisorWorkload = {};
      dataset.forEach(item => {
        const supervisor = normalizeSupervisorName(item.supervisor);
        const stitchingDays = typeof item.agingDays === 'number' ? item.agingDays : 0;
        const isCompleted = item.status === 'Completed' || isLotCompleted(item.completedStatus);

        if (!supervisorWorkload[supervisor]) {
          supervisorWorkload[supervisor] = {
            totalLots: 0,
            totalPCS: 0,
            greenLots: 0,
            greenPCS: 0,
            yellowLots: 0,
            yellowPCS: 0,
            redLots: 0,
            redPCS: 0,
            notUpdatedLots: 0,
            manpower: 0
          };
        }

        supervisorWorkload[supervisor].totalLots += 1;
        supervisorWorkload[supervisor].totalPCS += (item.poQty || item.totalPCS || 0);

        if (!isStatusUpdatedToday(item.wipStatus, isCompleted ? 'Completed' : 'Pending')) {
          supervisorWorkload[supervisor].notUpdatedLots += 1;
        }

        if (stitchingDays <= 6) {
          supervisorWorkload[supervisor].greenLots += 1;
          supervisorWorkload[supervisor].greenPCS += (item.poQty || item.totalPCS || 0);
        } else if (stitchingDays <= 15) {
          supervisorWorkload[supervisor].yellowLots += 1;
          supervisorWorkload[supervisor].yellowPCS += (item.poQty || item.totalPCS || 0);
        } else {
          supervisorWorkload[supervisor].redLots += 1;
          supervisorWorkload[supervisor].redPCS += (item.poQty || item.totalPCS || 0);
        }
      });

      const sortedSupervisors = Object.entries(supervisorWorkload).sort(([, a], [, b]) => b.totalLots - a.totalLots);

      const summaryBody = sortedSupervisors.map(([supervisorName, data]) => {
        const greenPercent = data.totalLots > 0 ? Math.round((data.greenLots / data.totalLots) * 100) : 0;
        const yellowPercent = data.totalLots > 0 ? Math.round((data.yellowLots / data.totalLots) * 100) : 0;
        const redPercent = data.totalLots > 0 ? Math.round((data.redLots / data.totalLots) * 100) : 0;
        const manpower = data.manpower || 0;
        const avgPCS = manpower > 0 ? Math.round(data.totalPCS / manpower) : 0;

        return [
          {
            content: supervisorName,
            styles: {
              halign: 'left',
              fontSize: 11,
              cellPadding: { top: 4, right: 3, bottom: 4, left: 6 },
              fontStyle: 'bold',
              fillColor: [240, 249, 255],
              textColor: [0, 0, 0]
            }
          },
          {
            content: manpower > 0 ? manpower.toString() : '-',
            styles: {
              halign: 'center',
              fontSize: 12,
              cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
              fontStyle: 'bold',
              fillColor: manpower > 0 ? [220, 252, 231] : [245, 245, 245],
              textColor: [0, 0, 0]
            }
          },
          {
            content: data.totalLots.toString(),
            styles: {
              halign: 'center',
              fontSize: 12,
              cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
              fontStyle: 'bold',
              textColor: [0, 0, 0]
            }
          },
          {
            content: data.totalPCS.toLocaleString(),
            styles: {
              halign: 'center',
              fontSize: 12,
              cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
              fontStyle: 'bold',
              textColor: [0, 0, 0]
            }
          },
          {
            content: avgPCS > 0 ? avgPCS.toLocaleString() : '-',
            styles: {
              halign: 'center',
              fontSize: 11,
              cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
              fontStyle: 'bold',
              fillColor: [255, 250, 240],
              textColor: [0, 0, 0]
            }
          },
          {
            content: data.notUpdatedLots > 0 ? data.notUpdatedLots.toString() : '-',
            styles: {
              halign: 'center',
              fontSize: 11,
              cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
              fontStyle: 'bold',
              textColor: [0, 0, 0],
              fillColor: data.notUpdatedLots > 0 ? [255, 235, 235] : [255, 255, 255]
            }
          },
          {
            content: `${data.greenLots} (${greenPercent}%)`,
            styles: {
              halign: 'center',
              fontSize: 12,
              cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
              fontStyle: 'bold',
              fillColor: [220, 252, 231],
              textColor: [0, 0, 0]
            }
          },
          {
            content: `${data.yellowLots} (${yellowPercent}%)`,
            styles: {
              halign: 'center',
              fontSize: 12,
              cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
              fontStyle: 'bold',
              fillColor: [254, 243, 199],
              textColor: [0, 0, 0]
            }
          },
          {
            content: `${data.redLots} (${redPercent}%)`,
            styles: {
              halign: 'center',
              fontSize: 12,
              cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
              fontStyle: 'bold',
              fillColor: [254, 226, 226],
              textColor: [0, 0, 0]
            }
          }
        ];
      });

      const totalManpowerAll = sortedSupervisors.reduce((sum, [, d]) => sum + (d.manpower || 0), 0);
      const totalGreenLotsAll = sortedSupervisors.reduce((sum, [, d]) => sum + d.greenLots, 0);
      const totalYellowLotsAll = sortedSupervisors.reduce((sum, [, d]) => sum + d.yellowLots, 0);
      const totalRedLotsAll = sortedSupervisors.reduce((sum, [, d]) => sum + d.redLots, 0);
      const totalNotUpdatedAll = sortedSupervisors.reduce((sum, [, d]) => sum + d.notUpdatedLots, 0);

      const greenPctAll = totalLotsAll > 0 ? Math.round((totalGreenLotsAll / totalLotsAll) * 100) : 0;
      const yellowPctAll = totalLotsAll > 0 ? Math.round((totalYellowLotsAll / totalLotsAll) * 100) : 0;
      const redPctAll = totalLotsAll > 0 ? Math.round((totalRedLotsAll / totalLotsAll) * 100) : 0;
      const avgPCSAll = totalManpowerAll > 0 ? Math.round(totalPCSAll / totalManpowerAll) : 0;

      const totalRowBg = [225, 239, 255];
      summaryBody.push([
        { content: 'OVERALL TOTALS', styles: { halign: 'left', fontSize: 11, fontStyle: 'bold', fillColor: totalRowBg, textColor: [0, 0, 0], cellPadding: { top: 5, right: 3, bottom: 5, left: 6 } } },
        { content: totalManpowerAll > 0 ? totalManpowerAll.toString() : '-', styles: { halign: 'center', fontSize: 12, fontStyle: 'bold', fillColor: totalRowBg, textColor: [0, 0, 0], cellPadding: { top: 5, right: 3, bottom: 5, left: 3 } } },
        { content: totalLotsAll.toString(), styles: { halign: 'center', fontSize: 12, fontStyle: 'bold', fillColor: totalRowBg, textColor: [0, 0, 0], cellPadding: { top: 5, right: 3, bottom: 5, left: 3 } } },
        { content: totalPCSAll.toLocaleString(), styles: { halign: 'center', fontSize: 12, fontStyle: 'bold', fillColor: totalRowBg, textColor: [0, 0, 0], cellPadding: { top: 5, right: 3, bottom: 5, left: 3 } } },
        { content: avgPCSAll > 0 ? avgPCSAll.toLocaleString() : '-', styles: { halign: 'center', fontSize: 11, fontStyle: 'bold', fillColor: totalRowBg, textColor: [0, 0, 0], cellPadding: { top: 5, right: 3, bottom: 5, left: 3 } } },
        { content: totalNotUpdatedAll > 0 ? totalNotUpdatedAll.toString() : '-', styles: { halign: 'center', fontSize: 11, fontStyle: 'bold', fillColor: totalRowBg, textColor: [0, 0, 0], cellPadding: { top: 5, right: 3, bottom: 5, left: 3 } } },
        { content: `${totalGreenLotsAll} (${greenPctAll}%)`, styles: { halign: 'center', fontSize: 12, fontStyle: 'bold', fillColor: [220, 252, 231], textColor: [0, 0, 0], cellPadding: { top: 5, right: 3, bottom: 5, left: 3 } } },
        { content: `${totalYellowLotsAll} (${yellowPctAll}%)`, styles: { halign: 'center', fontSize: 12, fontStyle: 'bold', fillColor: [254, 243, 199], textColor: [0, 0, 0], cellPadding: { top: 5, right: 3, bottom: 5, left: 3 } } },
        { content: `${totalRedLotsAll} (${redPctAll}%)`, styles: { halign: 'center', fontSize: 12, fontStyle: 'bold', fillColor: [254, 226, 226], textColor: [0, 0, 0], cellPadding: { top: 5, right: 3, bottom: 5, left: 3 } } }
      ]);

      const summaryCols = [
        contentWidth * 0.18,
        contentWidth * 0.12,
        contentWidth * 0.09,
        contentWidth * 0.11,
        contentWidth * 0.11,
        contentWidth * 0.11,
        contentWidth * 0.09,
        contentWidth * 0.09,
        contentWidth * 0.10
      ];

      autoTable(doc, {
        startY: 35,
        head: [[
          { content: 'SUPERVISOR', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryCols[0] } },
          { content: 'MANPOWER (ATTENDANCE)', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryCols[1], fillColor: [15, 76, 129], textColor: [255, 255, 255] } },
          { content: 'TOTAL LOTS', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryCols[2] } },
          { content: 'TOTAL PCS', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryCols[3] } },
          { content: 'AVG PCS / MANPOWER', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryCols[4] } },
          { content: 'NOT UPDATED', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryCols[5], fillColor: [220, 38, 38], textColor: [255, 255, 255] } },
          { content: 'GREEN (1-6 Days)', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryCols[6], fillColor: [220, 252, 231], textColor: [0, 0, 0] } },
          { content: 'YELLOW (7-15 Days)', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryCols[7], fillColor: [254, 243, 199], textColor: [0, 0, 0] } },
          { content: 'RED (15+ Days)', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryCols[8], fillColor: [254, 226, 226], textColor: [0, 0, 0] } }
        ]],
        body: summaryBody.map(row => row.map((cell, colIndex) => ({
          content: cell.content,
          styles: {
            ...cell.styles,
            cellWidth: summaryCols[colIndex]
          }
        }))),
        theme: 'grid',
        headStyles: {
          fillColor: [15, 76, 129],
          textColor: [255, 255, 255],
          fontStyle: 'bold',
          fontSize: 9,
          cellPadding: { top: 5, right: 3, bottom: 5, left: 3 },
          lineWidth: 0.5,
          lineColor: [15, 76, 129],
          halign: 'center',
          valign: 'middle'
        },
        bodyStyles: {
          fontSize: 10,
          cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
          lineWidth: 0.3,
          lineColor: [220, 220, 220],
          textColor: [0, 0, 0],
          font: 'helvetica',
          valign: 'middle'
        },
        margin: { top: 35, left: margin, right: margin },
        tableWidth: 'auto',
        showHead: 'everyPage',
        showFoot: false,
        pageBreak: 'auto',
        rowPageBreak: 'avoid'
      });
    };

    // Total PCS & Lots
    const totalPCSAll = exportData.reduce((sum, item) => sum + (item.poQty || item.totalPCS || 0), 0);
    const totalLotsAll = exportData.length;
    const totalCompletedAll = exportData.filter(item => item.status === 'Completed' || isLotCompleted(item.completedStatus)).length;
    const totalPendingAll = totalLotsAll - totalCompletedAll;

    const drawCombinedHeader = () => {
      doc.setFillColor(255, 255, 255);
      doc.rect(0, 0, pageWidth, 30, 'F');

      const title = 'FACTORY SUITE PRO - REPORT 6 - STITCHING PRODUCTION REPORT';

      doc.setFontSize(18);
      doc.setTextColor(0, 0, 0);
      doc.setFont('Times New Roman', 'bold');
      doc.text(title, pageWidth / 2, 12, { align: 'center' });

      doc.setFontSize(10);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(0, 0, 0);
      doc.text(`COMBINED ALL LOTS REPORT (Total Lots: ${totalLotsAll})`, pageWidth / 2, 18, { align: 'center' });

      // Key Metrics Row
      doc.setFontSize(11);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(0, 0, 0);

      const today = new Date();
      const reportDate = `${String(today.getDate()).padStart(2, '0')}/${String(today.getMonth() + 1).padStart(2, '0')}/${String(today.getFullYear()).slice(-2)}`;
      doc.text(`Report Date: ${reportDate}`, margin, 25);

      const centerX = pageWidth / 2;
      doc.text(`Lots: ${totalLotsAll} | Total PCS: ${totalPCSAll.toLocaleString()} | Completed: ${totalCompletedAll} | Pending: ${totalPendingAll} | Supervisors: ${supervisorsList.length}`,
        centerX, 25, { align: 'center' });

      doc.text(`Total Records: ${totalLotsAll}`, pageWidth - margin, 25, { align: 'right' });
    };

    drawCombinedHeader();
    const currentY = addColorLegend(32);

    // Build continuous table body for all lots
    const body = exportData.map((item, rowIndex) => mapItemToRow(item, rowIndex));

    // Add Grand Total Row at bottom
    const totalRow = activeCols.map((col, colIdx) => {
      const cellWidth = columnWidths[colIdx];
      const cellPad = { top: 3, right: 1, bottom: 3, left: 1 };
      const fontSz = activeCols.length > 24 ? 7.5 : (activeCols.length > 18 ? 8.5 : 9.5);

      if (col.id === 'sr') {
        return {
          content: '',
          styles: { cellWidth, fillColor: [225, 239, 255], cellPadding: cellPad }
        };
      }
      if (col.id === 'lotNo') {
        return {
          content: `TOTAL (${totalLotsAll})`,
          styles: { cellWidth, fontSize: Math.max(fontSz, 9.5), halign: 'center', fontStyle: 'bold', fillColor: [225, 239, 255], textColor: [0, 0, 0], cellPadding: cellPad }
        };
      }
      if (col.id === 'totalPcs') {
        return {
          content: totalPCSAll.toLocaleString(),
          styles: { cellWidth, fontSize: Math.max(fontSz, 9.5), halign: 'center', fontStyle: 'bold', fillColor: [225, 239, 255], textColor: [0, 0, 0], cellPadding: cellPad }
        };
      }
      if (col.id === 'supervisor') {
        return {
          content: `${supervisorsList.length} Sups`,
          styles: { cellWidth, fontSize: fontSz, halign: 'center', fontStyle: 'bold', fillColor: [225, 239, 255], textColor: [0, 0, 0], cellPadding: cellPad }
        };
      }
      if (col.id === 'lotStatus') {
        return {
          content: `${totalCompletedAll} Comp | ${totalPendingAll} Pend`,
          styles: { cellWidth, fontSize: Math.max(6.5, fontSz - 1), halign: 'center', fontStyle: 'bold', fillColor: [225, 239, 255], textColor: [0, 0, 0], cellPadding: cellPad }
        };
      }
      return {
        content: '',
        styles: { cellWidth, fillColor: [225, 239, 255], cellPadding: cellPad }
      };
    });

    body.push(totalRow);

    let lastAutoTableY = currentY;

    autoTable(doc, {
      startY: currentY,
      head: tableHeaders,
      body: body,
      theme: 'grid',
      headStyles: {
        fillColor: headerColor,
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        fontSize: activeCols.length > 24 ? 7.5 : (activeCols.length > 18 ? 8.5 : 9.5),
        cellPadding: { top: 3.5, right: 1.5, bottom: 3.5, left: 1.5 },
        lineWidth: 0.5,
        lineColor: headerColor,
        halign: 'center',
        valign: 'middle'
      },
      bodyStyles: {
        fontSize: activeCols.length > 24 ? 7.5 : (activeCols.length > 18 ? 8.5 : 9.5),
        cellPadding: { top: 2.5, right: 1.5, bottom: 2.5, left: 1.5 },
        lineWidth: 0.3,
        lineColor: borderColor,
        textColor: [0, 0, 0],
        fillColor: [255, 255, 255],
        font: 'helvetica',
        valign: 'middle',
        overflow: 'linebreak',
        minCellHeight: 6.5,
        lineHeight: 1.18
      },
      columnStyles: columnStyles,
      margin: { top: 35, left: margin, right: margin },
      tableWidth: contentWidth,
      showHead: 'everyPage',
      showFoot: false,
      pageBreak: 'auto',
      rowPageBreak: 'avoid',
      tableLineWidth: 0.5,
      tableLineColor: borderColor,
      didDrawPage: function (data) {
        drawCombinedHeader();
        if (data.cursor && data.cursor.y) {
          lastAutoTableY = data.cursor.y;
        }
      }
    });

    const summaryStartY = (doc.lastAutoTable ? doc.lastAutoTable.finalY : lastAutoTableY) + 12;
    renderThreeSideBySideSummary(exportData, 'COMBINED PRODUCTION', summaryStartY);

    // Overall supervisor summary page
    renderSupervisorWorkloadSummary(exportData);

    const todayDate = new Date();
    const formattedFileDate = `${String(todayDate.getDate()).padStart(2, '0')}-${String(todayDate.getMonth() + 1).padStart(2, '0')}-${todayDate.getFullYear()}`;
    doc.save(`Stitching_Production_Report_${formattedFileDate}.pdf`);
  };

  // 5. GENERIC STAGE PDF EXPORTER FOR OTHER STAGES (5, 7 TO 10)
  const exportGenericStagePDF = (cfg, stageRows) => {
    const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a3" });
    const pageWidth = doc.internal.pageSize.getWidth();
    const now = new Date();
    const reportDateStr = `${String(now.getDate()).padStart(2, "0")}/${String(now.getMonth() + 1).padStart(2, "0")}/${now.getFullYear()}`;

    const marginX = 14;
    const tableW = pageWidth - marginX * 2;

    doc.setFillColor(cfg.rgb[0], cfg.rgb[1], cfg.rgb[2]);
    doc.rect(0, 0, pageWidth, 24, "F");

    doc.setFont("helvetica", "bold");
    doc.setFontSize(16);
    doc.setTextColor(255, 255, 255);
    doc.text(`FACTORY SUITE PRO - ${cfg.reportTitle}`, 14, 11);

    const stagePending = stageRows.reduce((s, r) => s + (r.pendingQty || 0), 0);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.setTextColor(254, 202, 202);
    doc.text(`Generated: ${reportDateStr}  |  FY: ${financialYearFilter}  |  Total Lots: ${stageRows.length}  |  Pending Quantity: ${stagePending.toLocaleString()} Pcs`, 14, 18);

    const tableHeaders = [
      "#",
      "Lot #",
      "Bottleneck / Variance Reason",
      "Severity",
      "Aging",
      "Garment Type",
      "Style",
      "Fabric",
      "Brand",
      "Party",
      "PO / Cut",
      "Done",
      "Pending",
      "Variance"
    ];

    const tableData = stageRows.map((r, idx) => [
      idx + 1,
      r.lotNumber,
      r.reason,
      r.severity,
      `${r.agingDays}d`,
      r.garmentType,
      r.style,
      r.fabric,
      r.brand,
      r.partyName,
      (r.poQty || 0).toLocaleString(),
      (r.stageDoneQty || 0).toLocaleString(),
      (r.pendingQty || 0).toLocaleString(),
      `${r.variancePcs > 0 ? "+" : ""}${r.variancePcs} (${r.variancePct}%)`
    ]);

    const columnStyles = {
      0: { halign: "center", cellWidth: 10 },
      1: { halign: "center", fontStyle: "bold", cellWidth: 20 },
      2: { cellWidth: 70, fontStyle: "bold" },
      3: { halign: "center", fontStyle: "bold", cellWidth: 22 },
      4: { halign: "center", fontStyle: "bold", cellWidth: 14 },
      5: { cellWidth: 28 },
      6: { cellWidth: 28 },
      7: { cellWidth: 34 },
      8: { cellWidth: 24 },
      9: { cellWidth: 28 },
      10: { halign: "right", cellWidth: 20 },
      11: { halign: "right", cellWidth: 18 },
      12: { halign: "right", fontStyle: "bold", cellWidth: 20 },
      13: { halign: "right", cellWidth: 24 }
    };

    const scaledColumnStyles = scaleColumnStyles(columnStyles, tableW);

    autoTable(doc, {
      head: [tableHeaders],
      body: tableData,
      startY: 32,
      tableWidth: tableW,
      theme: "grid",
      styles: {
        fontSize: 8,
        cellPadding: 2.2,
        textColor: [0, 0, 0],
        font: "helvetica",
        valign: "middle"
      },
      headStyles: {
        fillColor: [15, 23, 42],
        textColor: [255, 255, 255],
        fontStyle: "bold",
        fontSize: 8.5,
        halign: "center"
      },
      alternateRowStyles: {
        fillColor: [248, 250, 252]
      },
      columnStyles: scaledColumnStyles,
      margin: { left: marginX, right: marginX, bottom: 16 }
    });

    doc.save(`${cfg.filePrefix}_${now.toISOString().slice(0, 10)}.pdf`);
  };

  // 6. Packing Pending to Issue PDF Export (Matches Reports 1-4 style exactly)
  const exportPackingHandoverPDF = (stageRows) => {
    const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "A3" });
    const pageW = doc.internal.pageSize.getWidth();
    const pageH = doc.internal.pageSize.getHeight();
    const now = new Date();

    const marginX = 16;
    const tableW = pageW - marginX * 2;

    const totalLots = stageRows.length;
    const totalQty = stageRows.reduce((sum, r) => sum + (r.poQty || r.stitchingQty || 0), 0);
    const redZoneLots = stageRows.filter((r) => r.agingDays > 2).length;
    const normalLots = totalLots - redZoneLots;
    const directLotsCount = stageRows.filter((r) => (r.directStitching || "").toLowerCase() === "yes").length;
    const highPriorityCount = stageRows.filter((r) => (r.priority || "").toLowerCase() === "high" || (r.priority || "").toLowerCase() === "urgent").length;
    const avgPendingDays = totalLots > 0 ? Math.round(stageRows.reduce((s, r) => s + (r.agingDays || 0), 0) / totalLots) : 0;

    const garmentMap = {};
    const supervisorMap = {};
    const partyMap = {};

    stageRows.forEach((r) => {
      const g = r.garmentType || "Unassigned";
      const sup = r.supervisor || "Unassigned";
      const p = r.partyName || "Unassigned";
      const qty = r.poQty || r.stitchingQty || 0;

      if (!garmentMap[g]) garmentMap[g] = { name: g, lots: 0, qty: 0 };
      garmentMap[g].lots += 1;
      garmentMap[g].qty += qty;

      if (!supervisorMap[sup]) supervisorMap[sup] = { name: sup, lots: 0, qty: 0 };
      supervisorMap[sup].lots += 1;
      supervisorMap[sup].qty += qty;

      if (!partyMap[p]) partyMap[p] = { name: p, lots: 0, qty: 0 };
      partyMap[p].lots += 1;
      partyMap[p].qty += qty;
    });

    const sortedGarments = Object.values(garmentMap).sort((a, b) => b.qty - a.qty);
    const sortedSupervisors = Object.values(supervisorMap).sort((a, b) => b.qty - a.qty);
    const sortedParties = Object.values(partyMap).sort((a, b) => b.qty - a.qty);

    doc.setFontSize(20);
    doc.setTextColor(0, 0, 0);
    doc.setFont("helvetica", "bold");
    doc.text("REPORT 6 - PACKING PENDING TO ISSUE VARIANCE REPORT", pageW / 2, 38, { align: "center" });

    doc.setFontSize(10);
    doc.setTextColor(50, 50, 50);
    doc.setFont("helvetica", "normal");
    doc.text(`Report Date: ${now.toLocaleDateString()} ${now.toLocaleTimeString()} | Factory Suite Pro`, pageW / 2, 54, { align: "center" });

    doc.setFillColor(239, 246, 255);
    doc.roundedRect(marginX, 64, tableW, 24, 6, 6, "F");
    doc.setDrawColor(191, 219, 254);
    doc.roundedRect(marginX, 64, tableW, 24, 6, 6, "D");

    doc.setFontSize(10.5);
    doc.setTextColor(0, 0, 0);
    doc.setFont("helvetica", "bold");
    const summaryText = `Total Records: ${totalLots}   |   Total Stitching Qty: ${totalQty.toLocaleString()} Pcs   |   Direct Lots: ${directLotsCount}   |   High Priority: ${highPriorityCount}   |   Supervisors: ${sortedSupervisors.length}   |   Red Zone (>2 Days): ${redZoneLots}`;
    doc.text(summaryText, pageW / 2, 79, { align: "center" });

    const columns = [
      "#",
      "Lot Number",
      "Garment Type",
      "Style",
      "Fabric",
      "Brand",
      "Stitching Qty",
      "M/W/K",
      "Season",
      "Party Name",
      "Direct Stitching",
      "Supervisor",
      "Date of Issue",
      "Priority",
      "Completed Date",
      "Pending Days",
      "User Remarks"
    ];

    const body = stageRows.map((r, idx) => [
      String(idx + 1),
      r.lotNumber,
      r.garmentType || "-",
      r.style || "-",
      r.fabric || "-",
      r.brand || "-",
      (r.poQty || r.stitchingQty || 0).toLocaleString(),
      r.mwk || r.section || "-",
      r.season || "-",
      r.partyName || "-",
      r.directStitching || "No",
      r.supervisor || "Unassigned",
      r.dateOfIssue || r.stageDate || "-",
      r.priority || "Normal",
      r.completedDate || "-",
      r.agingDays != null ? `${r.agingDays}` : "-",
      formatLatestRemark(r.userRemarks, "-")
    ]);

    body.push([
      "",
      `TOTAL (${totalLots})`,
      "",
      "",
      "",
      "",
      totalQty.toLocaleString(),
      "",
      "",
      `${sortedParties.length} Parties`,
      `${directLotsCount} Direct`,
      `${sortedSupervisors.length} Sups`,
      "",
      `${highPriorityCount} High`,
      "",
      `${redZoneLots} Red | ${normalLots} Norm`,
      ""
    ]);

    const columnStyles = {
      0: { cellWidth: 26, halign: "center" },
      1: { cellWidth: 60, halign: "center", fontStyle: "bold" },
      2: { cellWidth: 65, halign: "center" },
      3: { cellWidth: 70, halign: "center" },
      4: { cellWidth: 70, halign: "center" },
      5: { cellWidth: 55, halign: "center" },
      6: { cellWidth: 55, halign: "center", fontStyle: "bold" },
      7: { cellWidth: 40, halign: "center" },
      8: { cellWidth: 45, halign: "center" },
      9: { cellWidth: 60, halign: "center" },
      10: { cellWidth: 45, halign: "center" },
      11: { cellWidth: 55, halign: "center" },
      12: { cellWidth: 55, halign: "center" },
      13: { cellWidth: 45, halign: "center" },
      14: { cellWidth: 55, halign: "center" },
      15: { cellWidth: 45, halign: "center" },
      16: { cellWidth: 110, halign: "left" }
    };

    const scaledColumnStyles = scaleColumnStyles(columnStyles, tableW);

    autoTable(doc, {
      head: [columns],
      body,
      startY: 96,
      tableWidth: tableW,
      margin: { top: 96, right: marginX, bottom: 25, left: marginX },
      theme: "grid",
      styles: {
        fontSize: 8.5,
        cellPadding: { top: 3.5, right: 2, bottom: 3.5, left: 2 },
        overflow: "linebreak",
        valign: "middle",
        halign: "center",
        textColor: [0, 0, 0],
        lineColor: [0, 0, 0],
        lineWidth: 0.5
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
        cellPadding: { top: 5, right: 2, bottom: 5, left: 2 }
      },
      alternateRowStyles: {
        fillColor: [248, 250, 252]
      },
      columnStyles: scaledColumnStyles,
      didParseCell: function (data) {
        if (data.section === "body") {
          data.cell.styles.textColor = [0, 0, 0];
          const rowIndex = data.row.index;
          const isTotalRow = rowIndex === body.length - 1;

          if (isTotalRow) {
            data.cell.styles.fontStyle = "bold";
            data.cell.styles.fillColor = [226, 232, 240];
            data.cell.styles.textColor = [0, 0, 0];
            return;
          }

          if (data.column.index === 15) {
            const rawVal = parseFloat(data.cell.raw);
            if (!isNaN(rawVal)) {
              data.cell.styles.fontStyle = "bold";
              if (rawVal > 5) {
                data.cell.styles.fillColor = [239, 68, 68];
                data.cell.styles.textColor = [255, 255, 255];
              } else if (rawVal > 2) {
                data.cell.styles.fillColor = [254, 243, 199];
                data.cell.styles.textColor = [180, 83, 9];
              } else {
                data.cell.styles.fillColor = [220, 252, 231];
                data.cell.styles.textColor = [21, 128, 61];
              }
            }
          }
        }
      },
      didDrawPage: () => {
        doc.setDrawColor(0, 0, 0);
        doc.setLineWidth(0.5);
        doc.roundedRect(8, 8, pageW - 16, pageH - 16, 2, 2, "S");
      }
    });

    const gBody = sortedGarments.map((item) => {
      const pct = totalQty > 0 ? ((item.qty / totalQty) * 100).toFixed(1) : "0.0";
      return [item.name, item.lots.toString(), `${pct}%`, item.qty.toLocaleString()];
    });
    gBody.push(["TOTAL", totalLots.toString(), "100.0%", totalQty.toLocaleString()]);

    const supBody = sortedSupervisors.map((item) => {
      const pct = totalQty > 0 ? ((item.qty / totalQty) * 100).toFixed(1) : "0.0";
      return [item.name, item.lots.toString(), `${pct}%`, item.qty.toLocaleString()];
    });
    supBody.push(["TOTAL", totalLots.toString(), "100.0%", totalQty.toLocaleString()]);

    const aBody = [
      ["<= 2 Days (On-Time / Normal)", normalLots.toString(), `${totalLots > 0 ? ((normalLots / totalLots) * 100).toFixed(1) : 0}%`, "—"],
      ["> 2 Days (Red Zone / Delayed)", redZoneLots.toString(), `${totalLots > 0 ? ((redZoneLots / totalLots) * 100).toFixed(1) : 0}%`, "—"],
      ["TOTAL LOTS", totalLots.toString(), "100.0%", totalQty.toLocaleString()]
    ];

    let summaryStartY = doc.lastAutoTable.finalY + 22;
    if (summaryStartY + 140 > pageH - 30) {
      doc.addPage();
      summaryStartY = 40;
    }

    doc.setFontSize(13);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(0, 0, 0);
    doc.text("EXECUTIVE SUMMARY & PRODUCTION BREAKDOWN", pageW / 2, summaryStartY + 4, { align: "center" });

    const sectionTitleY = summaryStartY + 26;
    const tableStartY = sectionTitleY + 6;
    const gap = 18;
    const colWidth = (tableW - 2 * gap) / 3;
    const col1X = marginX;
    const col2X = col1X + colWidth + gap;
    const col3X = col2X + colWidth + gap;

    const summaryColStyles = {
      0: { cellWidth: 140, halign: "center" },
      1: { cellWidth: 60, halign: "center" },
      2: { cellWidth: 70, halign: "center" },
      3: { cellWidth: 100, halign: "center" }
    };

    const scaledSummaryColStyles = scaleColumnStyles(summaryColStyles, colWidth);

    [
      { body: gBody, left: col1X, fill: [15, 118, 110], label: "1. GARMENT TYPE BREAKDOWN", header: "Garment Type" },
      { body: supBody, left: col2X, fill: [30, 64, 175], label: "2. SUPERVISOR WORKLOAD", header: "Supervisor" },
      { body: aBody, left: col3X, fill: [180, 83, 9], label: "3. SLA & DAYS AGING", header: "Aging Status" }
    ].forEach((tbl) => {
      autoTable(doc, {
        head: [[tbl.header, "Lots", "Share %", "Total Pieces"]],
        body: tbl.body,
        startY: tableStartY,
        tableWidth: colWidth,
        margin: { left: tbl.left, right: pageW - (tbl.left + colWidth) },
        theme: "grid",
        styles: {
          fontSize: 8.5,
          cellPadding: { top: 3.5, right: 2, bottom: 3.5, left: 2 },
          valign: "middle",
          halign: "center",
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          lineWidth: 0.3
        },
        headStyles: {
          fillColor: tbl.fill,
          textColor: [255, 255, 255],
          fontStyle: "bold",
          fontSize: 8.5,
          halign: "center"
        },
        columnStyles: scaledSummaryColStyles,
        didParseCell: (data) => {
          if (data.section === "body") {
            data.cell.styles.textColor = [0, 0, 0];
            if (data.row.index === tbl.body.length - 1) {
              data.cell.styles.fontStyle = "bold";
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });
    });

    doc.save(`Report_6_Packing_Pending_To_Issue_Variance_${now.toISOString().slice(0, 10)}.pdf`);
  };

  /* Unified Single Stage PDF Exporter Router */
  const exportSingleStagePDF = (cfg, stageRows) => {
    if (!stageRows || stageRows.length === 0) return;
    if (cfg.id === "cutting") {
      exportCuttingPDF(stageRows);
    } else if (cfg.id === "embroidery") {
      exportEmbroideryPDF(stageRows);
    } else if (cfg.id === "printing") {
      exportPrintingPDF(stageRows);
    } else if (cfg.id === "post_emb_print") {
      exportPostEmbPrintPDF(stageRows);
    } else if (cfg.id === "stitching") {
      exportStitchingPDF(stageRows);
    } else if (cfg.id === "packing_handover") {
      exportPackingHandoverPDF(stageRows);
    } else {
      exportGenericStagePDF(cfg, stageRows);
    }
  };

  /* ======================================================================= */
  /* LIVE ANIMATED TOUR: SEQUENTIALLY OPENS STAGES & DOWNLOADS IN SEQUENCE   */
  /* ======================================================================= */
  const runLiveAnimatedBatchDownload = async (format = "excel") => {
    if (isAnimatedDownloading) return;
    abortAnimationRef.current = false;

    // Filter stages that actually have bottleneck rows in the current dataset
    const activeStages = STAGE_REPORT_CONFIGS.filter((cfg) => {
      const stageRows = filteredRows.filter((r) => r.stageId === cfg.id);
      return stageRows.length > 0;
    });

    if (activeStages.length === 0) {
      alert("No department records available to download.");
      return;
    }

    setIsAnimatedDownloading(true);

    try {
      for (let i = 0; i < activeStages.length; i++) {
        if (abortAnimationRef.current) break;
        const cfg = activeStages[i];
        const stageRows = filteredRows.filter((r) => r.stageId === cfg.id);

        // --- STEP 1: OPEN DEPARTMENT TAB ---
        setActiveStageTab(cfg.id);
        setCurrentPage(1);

        setAnimProgress({
          stageId: cfg.id,
          title: cfg.reportTitle,
          shortTitle: cfg.shortTitle,
          icon: cfg.icon || "📋",
          color: cfg.color ? `#${cfg.color.slice(2)}` : "#0f172a",
          step: i + 1,
          totalSteps: activeStages.length,
          lotCount: stageRows.length,
          statusText: `🚀 Step ${i + 1}/${activeStages.length}: Opening ${cfg.shortTitle} (${stageRows.length} Lots)...`,
          format,
          completed: false
        });

        // Pause so user sees the department view opening & table updating
        await new Promise((res) => setTimeout(res, 1000));
        if (abortAnimationRef.current) break;

        // --- STEP 2: DOWNLOAD REPORT ---
        setAnimProgress((prev) => ({
          ...prev,
          statusText: `📥 Downloading ${cfg.shortTitle} ${format.toUpperCase()}...`
        }));

        if (format === "excel") {
          await exportSingleStageExcel(cfg, stageRows);
        } else {
          exportSingleStagePDF(cfg, stageRows);
        }

        setAnimProgress((prev) => ({
          ...prev,
          statusText: `✅ ${cfg.shortTitle} Downloaded! Returning to Central Hub...`
        }));

        await new Promise((res) => setTimeout(res, 900));
        if (abortAnimationRef.current) break;

        // --- STEP 3: RETURN BACK TO ALL RED ZONE ---
        setActiveStageTab("all");
        setCurrentPage(1);

        if (i < activeStages.length - 1) {
          const nextCfg = activeStages[i + 1];
          setAnimProgress((prev) => ({
            ...prev,
            statusText: `↩️ Back at Overview. Next up: ${nextCfg.shortTitle}...`
          }));
          await new Promise((res) => setTimeout(res, 900));
        }
      }

      if (!abortAnimationRef.current) {
        // --- FINAL CELEBRATION STATE ---
        setActiveStageTab("all");
        setCurrentPage(1);

        setAnimProgress({
          stageId: "all",
          title: "All Department Reports Exported",
          shortTitle: "All Departments Done",
          icon: "🎉",
          color: "#059669",
          step: activeStages.length,
          totalSteps: activeStages.length,
          lotCount: filteredRows.length,
          statusText: `🎉 Tour Complete! All ${activeStages.length} Department Reports Downloaded Successfully!`,
          format,
          completed: true
        });

        await new Promise((res) => setTimeout(res, 2800));
      }
    } catch (e) {
      console.error("Live Animated Tour Download Error:", e);
      alert("An error occurred during the live download sequence.");
    } finally {
      setIsAnimatedDownloading(false);
    }
  };

  const handleStopAnimation = () => {
    abortAnimationRef.current = true;
    setIsAnimatedDownloading(false);
    setActiveStageTab("all");
  };

  /* ======================================================================= */
  /* EXCELJS EXPORT (COMBINED MULTI-STAGE WORKBOOK)                          */
  /* ======================================================================= */
  const handleExportExcel = async () => {
    try {
      if (filteredRows.length === 0) {
        alert("No bottleneck data available to export.");
        return;
      }

      const now = new Date();
      const reportDateStr = `${String(now.getDate()).padStart(2, "0")}/${String(now.getMonth() + 1).padStart(2, "0")}/${now.getFullYear()} ${String(now.getHours()).padStart(2, "0")}:${String(now.getMinutes()).padStart(2, "0")}`;

      const workbook = new ExcelJS.Workbook();
      workbook.creator = "Factory Suite Pro";
      workbook.created = now;

      // -----------------------------------------------------------------------
      // SHEET 1: MASTER COMBINED VARIANCE REPORT (SECTIONED WITH HEADINGS)
      // -----------------------------------------------------------------------
      const ws1 = workbook.addWorksheet("Combined Variance Report", {
        views: [{ showGridLines: true, state: "frozen", xSplit: 0, ySplit: 3 }]
      });
      ws1.columns = masterCols;

      // Row 1: Title Banner
      const titleRow = ws1.getRow(1);
      titleRow.values = ["FACTORY SUITE PRO — MULTI-STAGE VARIANCE COMBINED REPORT"];
      ws1.mergeCells(1, 1, 1, masterCols.length);
      titleRow.font = { name: "Segoe UI", size: 14, bold: true, color: { argb: "FFFFFFFF" } };
      titleRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF991B1B" } };
      titleRow.alignment = { vertical: "middle", horizontal: "center" };
      titleRow.height = 32;

      // Row 2: Metadata Banner
      const metaRow = ws1.getRow(2);
      metaRow.values = [`Generated: ${reportDateStr}  |  Financial Year: ${financialYearFilter}  |  Total Bottlenecks: ${analytics.totalLots} Lots  |  Total Pending: ${analytics.totalPendingPcs.toLocaleString()} Pcs  |  Max Aging: ${analytics.maxAging}d`];
      ws1.mergeCells(2, 1, 2, masterCols.length);
      metaRow.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF1E293B" } };
      metaRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEE2E2" } };
      metaRow.alignment = { vertical: "middle", horizontal: "center" };
      metaRow.height = 22;

      // Row 3: Blank Spacer
      ws1.getRow(3).height = 6;

      let rIdx = 4;
      let globalSr = 1;

      // Loop through each stage and write section headings
      STAGE_REPORT_CONFIGS.forEach((cfg) => {
        const stageRows = filteredRows.filter((r) => r.stageId === cfg.id);
        if (stageRows.length === 0) return;

        const stagePendingPcs = stageRows.reduce((sum, r) => sum + (r.pendingQty || 0), 0);

        // Section Heading Banner (e.g. REPORT 1: CUTTING VARIANCE REPORT)
        const secRow = ws1.getRow(rIdx);
        secRow.values = [`${cfg.reportTitle} — [ ${stageRows.length} Lots | ${stagePendingPcs.toLocaleString()} Pcs Pending ]`];
        ws1.mergeCells(rIdx, 1, rIdx, masterCols.length);
        secRow.font = { name: "Segoe UI", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
        secRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: cfg.color } };
        secRow.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
        secRow.height = 26;
        rIdx++;

        // Section Column Header Row
        const subHeaderRow = ws1.getRow(rIdx);
        subHeaderRow.values = masterCols.map((c) => c.header);
        subHeaderRow.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FFFFFFFF" } };
        subHeaderRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F172A" } };
        subHeaderRow.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
        subHeaderRow.height = 24;
        rIdx++;

        // Add Data Rows for this Stage
        stageRows.forEach((r, sIdx) => {
          const row = ws1.getRow(rIdx);
          row.values = [
            globalSr++,
            r.lotNumber,
            r.stageName,
            r.reason,
            r.severity,
            r.agingDays,
            r.garmentType,
            r.style,
            r.fabric,
            r.brand,
            r.partyName,
            r.section,
            r.season,
            r.poQty || 0,
            r.stageDoneQty || 0,
            r.pendingQty || 0,
            r.variancePcs || 0,
            `${r.variancePct || 0}%`,
            r.stageDate || "—"
          ];

          row.font = { name: "Segoe UI", size: 9.5, color: { argb: "FF0F172A" } };
          row.alignment = { vertical: "middle" };
          row.height = 20;

          const isEven = sIdx % 2 === 0;
          const bgFill = r.severity === "CRITICAL" ? "FFFFF1F2" : isEven ? "FFFFFFFF" : "FFF8FAFC";

          row.eachCell((cell, colNum) => {
            cell.border = borderThin;
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bgFill } };

            if ([1, 2, 5, 6, 12, 13, 19].includes(colNum)) {
              cell.alignment = { vertical: "middle", horizontal: "center" };
            }
            if (colNum >= 14 && colNum <= 18) {
              cell.alignment = { vertical: "middle", horizontal: "right" };
              if (colNum !== 18) cell.numFmt = "#,##0";
            }
            if (colNum === 5) {
              cell.font = {
                name: "Segoe UI",
                size: 9.5,
                bold: true,
                color: { argb: r.severity === "CRITICAL" ? "FFDC2626" : r.severity === "HIGH" ? "FFEA580C" : "FFD97706" }
              };
            }
          });
          rIdx++;
        });

        // Stage Subtotal Row
        const subTotalRow = ws1.getRow(rIdx);
        subTotalRow.values = [
          "",
          `SUBTOTAL (${stageRows.length} Lots)`,
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          stageRows.reduce((s, r) => s + (r.poQty || 0), 0),
          stageRows.reduce((s, r) => s + (r.stageDoneQty || 0), 0),
          stagePendingPcs,
          stageRows.reduce((s, r) => s + (r.variancePcs || 0), 0),
          "",
          ""
        ];
        subTotalRow.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF0F172A" } };
        subTotalRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } };
        subTotalRow.height = 22;
        subTotalRow.eachCell((cell, colNum) => {
          cell.border = borderThin;
          if (colNum >= 14 && colNum <= 17) {
            cell.alignment = { vertical: "middle", horizontal: "right" };
            cell.numFmt = "#,##0";
          }
        });
        rIdx++;

        // Spacer Row
        ws1.getRow(rIdx).height = 10;
        rIdx++;
      });

      // -----------------------------------------------------------------------
      // DEDICATED INDIVIDUAL SHEETS FOR KEY STAGES
      // -----------------------------------------------------------------------
      STAGE_REPORT_CONFIGS.slice(0, 3).forEach((cfg) => {
        const stageRows = filteredRows.filter((r) => r.stageId === cfg.id);
        if (stageRows.length === 0) return;

        const ws = workbook.addWorksheet(cfg.sheetName, {
          views: [{ showGridLines: true, state: "frozen", xSplit: 0, ySplit: 4 }]
        });
        ws.columns = masterCols;

        const tRow = ws.getRow(1);
        tRow.values = [`FACTORY SUITE PRO — ${cfg.reportTitle}`];
        ws.mergeCells(1, 1, 1, masterCols.length);
        tRow.font = { name: "Segoe UI", size: 13, bold: true, color: { argb: "FFFFFFFF" } };
        tRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: cfg.color } };
        tRow.alignment = { vertical: "middle", horizontal: "center" };
        tRow.height = 30;

        const stagePending = stageRows.reduce((s, r) => s + (r.pendingQty || 0), 0);
        const mRow = ws.getRow(2);
        mRow.values = [`Generated: ${reportDateStr}  |  Total Lots: ${stageRows.length}  |  Pending Quantity: ${stagePending.toLocaleString()} Pcs  |  FY: ${financialYearFilter}`];
        ws.mergeCells(2, 1, 2, masterCols.length);
        mRow.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF1E293B" } };
        mRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
        mRow.alignment = { vertical: "middle", horizontal: "center" };
        mRow.height = 20;

        ws.getRow(3).height = 6;

        const hRow = ws.getRow(4);
        hRow.values = masterCols.map((c) => c.header);
        hRow.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
        hRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F172A" } };
        hRow.alignment = { vertical: "middle", horizontal: "center" };
        hRow.height = 26;

        stageRows.forEach((r, idx) => {
          const row = ws.addRow({
            sr: idx + 1,
            lot: r.lotNumber,
            stage: r.stageName,
            reason: r.reason,
            severity: r.severity,
            aging: r.agingDays,
            garment: r.garmentType,
            style: r.style,
            fabric: r.fabric,
            brand: r.brand,
            party: r.partyName,
            section: r.section,
            season: r.season,
            poQty: r.poQty || 0,
            doneQty: r.stageDoneQty || 0,
            pendingQty: r.pendingQty || 0,
            varPcs: r.variancePcs || 0,
            varPct: `${r.variancePct || 0}%`,
            date: r.stageDate || "—"
          });
          row.font = { name: "Segoe UI", size: 9.5 };
          row.alignment = { vertical: "middle" };
          row.eachCell((c, cNum) => {
            c.border = borderThin;
            if ([1, 2, 5, 6, 12, 13, 19].includes(cNum)) c.alignment = { horizontal: "center" };
            if (cNum >= 14 && cNum <= 18) {
              c.alignment = { horizontal: "right" };
              if (cNum !== 18) c.numFmt = "#,##0";
            }
          });
        });
      });

      // -----------------------------------------------------------------------
      // SHEET: Stage & Severity Breakdown Matrix
      // -----------------------------------------------------------------------
      const wsMatrix = workbook.addWorksheet("Stage Analytics Matrix", {
        views: [{ showGridLines: true }]
      });

      wsMatrix.columns = [
        { header: "Report / Stage Name", key: "stage", width: 34 },
        { header: "Total Lots", key: "lots", width: 14 },
        { header: "Critical (>5d)", key: "crit", width: 16 },
        { header: "High (3-5d)", key: "high", width: 16 },
        { header: "Moderate (1-2d)", key: "mod", width: 16 },
        { header: "Pending Pieces", key: "pending", width: 18 },
        { header: "Max Aging (Days)", key: "maxAging", width: 16 }
      ];

      const t2 = wsMatrix.getRow(1);
      t2.values = ["MULTI-STAGE BOTTLENECK & VARIANCE EXECUTIVE MATRIX"];
      wsMatrix.mergeCells(1, 1, 1, 7);
      t2.font = { name: "Segoe UI", size: 12, bold: true, color: { argb: "FFFFFFFF" } };
      t2.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F172A" } };
      t2.alignment = { vertical: "middle", horizontal: "center" };
      t2.height = 28;

      const h2 = wsMatrix.getRow(2);
      h2.values = ["Report / Stage Name", "Total Lots", "Critical (>5d)", "High (3-5d)", "Moderate (1-2d)", "Pending Pieces", "Max Aging (Days)"];
      h2.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
      h2.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF334155" } };
      h2.height = 24;

      STAGE_REPORT_CONFIGS.forEach((cfg) => {
        const stageRows = filteredRows.filter((r) => r.stageId === cfg.id);
        if (stageRows.length === 0) return;

        const lots = stageRows.length;
        const crit = stageRows.filter((r) => r.severity === "CRITICAL").length;
        const high = stageRows.filter((r) => r.severity === "HIGH").length;
        const mod = stageRows.filter((r) => r.severity === "MODERATE").length;
        const pending = stageRows.reduce((s, r) => s + (r.pendingQty || 0), 0);
        const maxAging = stageRows.reduce((m, r) => Math.max(m, r.agingDays || 0), 0);

        const row = wsMatrix.addRow({
          stage: cfg.reportTitle,
          lots,
          crit,
          high,
          mod,
          pending,
          maxAging
        });
        row.font = { name: "Segoe UI", size: 9.5 };
        row.eachCell((c, i) => {
          c.border = borderThin;
          if (i > 1) {
            c.alignment = { horizontal: "right" };
            if (i === 6) c.numFmt = "#,##0";
          }
        });
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      saveAs(blob, `Factory_Suite_Pro_Combined_Variance_Report_${now.toISOString().slice(0, 10)}.xlsx`);
    } catch (e) {
      console.error("Excel Export Error:", e);
      alert("Failed to export Excel report.");
    }
  };

  /* ======================================================================= */
  /* PDF EXPORT (A3 LANDSCAPE SECTIONED BY REPORT 1, REPORT 2, REPORT 3...)  */
  /* ======================================================================= */
  const handleExportPDF = () => {
    try {
      if (filteredRows.length === 0) {
        alert("No bottleneck data available to export.");
        return;
      }

      const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a3" });
      const pageWidth = doc.internal.pageSize.getWidth();
      const pageHeight = doc.internal.pageSize.getHeight();
      const now = new Date();
      const reportDateStr = `${String(now.getDate()).padStart(2, "0")}/${String(now.getMonth() + 1).padStart(2, "0")}/${now.getFullYear()} ${String(now.getHours()).padStart(2, "0")}:${String(now.getMinutes()).padStart(2, "0")}`;

      // Master Top Header Banner (Deep Crimson Red)
      doc.setFillColor(153, 27, 27);
      doc.rect(0, 0, pageWidth, 24, "F");

      doc.setFont("helvetica", "bold");
      doc.setFontSize(16);
      doc.setTextColor(255, 255, 255);
      doc.text("FACTORY SUITE PRO — MULTI-STAGE VARIANCE COMBINED REPORT", 14, 11);

      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      doc.setTextColor(254, 202, 202);
      doc.text(
        `Generated: ${reportDateStr}  |  FY: ${financialYearFilter}  |  Total Bottlenecks: ${analytics.totalLots} Lots  |  Total Pending: ${analytics.totalPendingPcs.toLocaleString()} Pcs  |  Max Aging: ${analytics.maxAging} Days`,
        14,
        18
      );

      // KPI Metric Boxes (Side-by-Side Summary)
      const boxY = 28;
      const boxW = 58;
      const boxH = 16;
      const metrics = [
        { label: "TOTAL RED ZONE LOTS", val: `${analytics.totalLots} Lots`, color: [185, 28, 28] },
        { label: "CRITICAL SEVERITY (>5D)", val: `${analytics.criticalCount} Lots`, color: [220, 38, 38] },
        { label: "TOTAL PENDING PCS", val: `${analytics.totalPendingPcs.toLocaleString()} Pcs`, color: [234, 88, 12] },
        { label: "MAX PIPELINE AGING", val: `${analytics.maxAging} Days`, color: [180, 83, 9] },
        { label: "TOP BOTTLENECK STAGE", val: analytics.topStage, color: [30, 41, 59] }
      ];

      metrics.forEach((m, idx) => {
        const x = 14 + idx * (boxW + 6);
        doc.setFillColor(248, 250, 252);
        doc.setDrawColor(203, 213, 225);
        doc.roundedRect(x, boxY, boxW, boxH, 2, 2, "FD");

        doc.setFont("helvetica", "bold");
        doc.setFontSize(7.5);
        doc.setTextColor(100, 116, 139);
        doc.text(m.label, x + 4, boxY + 5.5);

        doc.setFontSize(11);
        doc.setTextColor(m.color[0], m.color[1], m.color[2]);
        doc.text(String(m.val), x + 4, boxY + 12);
      });

      let currentY = 50;
      const tableHeaders = [
        "#",
        "Lot #",
        "Bottleneck / Variance Reason",
        "Severity",
        "Aging",
        "Garment Type",
        "Style",
        "Fabric",
        "Brand",
        "Party",
        "PO / Cut",
        "Done",
        "Pending",
        "Variance"
      ];

      STAGE_REPORT_CONFIGS.forEach((cfg) => {
        const stageRows = filteredRows.filter((r) => r.stageId === cfg.id);
        if (stageRows.length === 0) return;

        const stagePendingPcs = stageRows.reduce((sum, r) => sum + (r.pendingQty || 0), 0);

        if (currentY > pageHeight - 40) {
          doc.addPage();
          currentY = 18;
        }

        doc.setFillColor(cfg.rgb[0], cfg.rgb[1], cfg.rgb[2]);
        doc.roundedRect(14, currentY, pageWidth - 28, 9, 1.5, 1.5, "F");

        doc.setFont("helvetica", "bold");
        doc.setFontSize(10.5);
        doc.setTextColor(255, 255, 255);
        doc.text(
          `${cfg.reportTitle} — [ ${stageRows.length} Lots | ${stagePendingPcs.toLocaleString()} Pcs Pending ]`,
          18,
          currentY + 6.2
        );

        currentY += 12;

        let secHeaders = [];
        let secData = [];
        let secColumnStyles = {};

        if (cfg.id === "cutting") {
          secHeaders = [
            "#",
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
            "Days after PO issue",
            "Pending Shade",
            "Cutting Date",
            "Cutting Scanned",
            "Remarks",
            "💬 User Remarks"
          ];
          secData = stageRows.map((r, idx) => [
            String(idx + 1),
            r.lotNumber,
            r.garmentType || "—",
            r.style || "—",
            r.fabric || "—",
            r.brand || "—",
            (r.poQty || 0).toLocaleString(),
            r.section || "—",
            r.season || "—",
            r.partyName || "—",
            r.directStitching || "No",
            r.jobOrderNo || "—",
            r.stageDate || "—",
            r.agingDays != null ? `${r.agingDays}` : "—",
            r.pendingShade || "—",
            r.cuttingDate || "—",
            r.cuttingScanned || "—",
            r.remarks || "—",
            r.userRemarks || "—"
          ]);
          secColumnStyles = {
            0: { cellWidth: 8, halign: "center" },
            1: { cellWidth: 20, halign: "center", fontStyle: "bold" },
            2: { cellWidth: 24, halign: "center" },
            3: { cellWidth: 26, halign: "center" },
            4: { cellWidth: 26, halign: "center" },
            5: { cellWidth: 18, halign: "center" },
            6: { cellWidth: 18, halign: "center", fontStyle: "bold" },
            7: { cellWidth: 15, halign: "center" },
            8: { cellWidth: 15, halign: "center" },
            9: { cellWidth: 20, halign: "center" },
            10: { cellWidth: 15, halign: "center" },
            11: { cellWidth: 20, halign: "center" },
            12: { cellWidth: 20, halign: "center" },
            13: { cellWidth: 15, halign: "center" },
            14: { cellWidth: 26, halign: "left" },
            15: { cellWidth: 20, halign: "center" },
            16: { cellWidth: 20, halign: "center" },
            17: { cellWidth: 32, halign: "left" },
            18: { cellWidth: 32, halign: "left" }
          };
        } else if (cfg.id === "embroidery") {
          secHeaders = [
            "S. No",
            "Lot No.",
            "Garment Type",
            "Style",
            "Fabric",
            "Brand",
            "Challan Qty",
            "Section (M/W/K)",
            "Season",
            "Party Name",
            "Direct Stitching",
            "Emb Party",
            "Challan Date",
            "Emb Status",
            "Pending Shade",
            "Remarks",
            "Emb Done Date",
            "Received Date",
            "Days",
            "HOD Remarks"
          ];
          secData = stageRows.map((r, idx) => [
            String(idx + 1),
            r.lotNumber,
            r.garmentType || "—",
            r.style || "—",
            r.fabric || "—",
            r.brand || "—",
            (r.poQty || 0).toLocaleString(),
            r.section || "—",
            r.season || "—",
            r.partyName || "—",
            r.directStitching || "No",
            r.embParty || r.partyName || "—",
            r.stageDate || "—",
            r.embStatus || "Pending",
            r.pendingShade || "—",
            r.remarks || "—",
            r.embDoneDate || "—",
            r.receivedDate || "—",
            r.agingDays != null ? `${r.agingDays}` : "—",
            r.hodRemarks || "—"
          ]);
          secColumnStyles = {
            0: { cellWidth: 8, halign: "center" },
            1: { cellWidth: 18, halign: "center", fontStyle: "bold" },
            2: { cellWidth: 22, halign: "center" },
            3: { cellWidth: 24, halign: "center" },
            4: { cellWidth: 22, halign: "center" },
            5: { cellWidth: 18, halign: "center" },
            6: { cellWidth: 18, halign: "center", fontStyle: "bold" },
            7: { cellWidth: 14, halign: "center" },
            8: { cellWidth: 15, halign: "center" },
            9: { cellWidth: 20, halign: "center" },
            10: { cellWidth: 15, halign: "center" },
            11: { cellWidth: 22, halign: "center" },
            12: { cellWidth: 20, halign: "center" },
            13: { cellWidth: 20, halign: "center" },
            14: { cellWidth: 26, halign: "left" },
            15: { cellWidth: 24, halign: "left" },
            16: { cellWidth: 20, halign: "center" },
            17: { cellWidth: 20, halign: "center" },
            18: { cellWidth: 14, halign: "center" },
            19: { cellWidth: 28, halign: "left" }
          };
        } else if (cfg.id === "printing") {
          secHeaders = [
            "S. No",
            "Lot No.",
            "Garment Type",
            "Style",
            "Fabric",
            "Brand",
            "Challan Qty",
            "Section (M/W/K)",
            "Season",
            "Party Name",
            "Direct Stitching",
            "Printing Party",
            "Challan Date",
            "Printing Status",
            "Pending Shade",
            "Remarks",
            "Printing Done Date",
            "Days",
            "HOD Remarks"
          ];
          secData = stageRows.map((r, idx) => [
            String(idx + 1),
            r.lotNumber,
            r.garmentType || "—",
            r.style || "—",
            r.fabric || "—",
            r.brand || "—",
            (r.poQty || 0).toLocaleString(),
            r.section || "—",
            r.season || "—",
            r.partyName || "—",
            r.directStitching || "No",
            r.printingParty || r.partyName || "—",
            r.stageDate || "—",
            r.printingStatus || "Pending",
            r.pendingShade || "—",
            r.remarks || "—",
            r.printingDoneDate || "—",
            r.agingDays != null ? `${r.agingDays}` : "—",
            r.hodRemarks || "—"
          ]);
          secColumnStyles = {
            0: { cellWidth: 8, halign: "center" },
            1: { cellWidth: 18, halign: "center", fontStyle: "bold" },
            2: { cellWidth: 22, halign: "center" },
            3: { cellWidth: 24, halign: "center" },
            4: { cellWidth: 22, halign: "center" },
            5: { cellWidth: 18, halign: "center" },
            6: { cellWidth: 18, halign: "center", fontStyle: "bold" },
            7: { cellWidth: 14, halign: "center" },
            8: { cellWidth: 15, halign: "center" },
            9: { cellWidth: 20, halign: "center" },
            10: { cellWidth: 15, halign: "center" },
            11: { cellWidth: 22, halign: "center" },
            12: { cellWidth: 20, halign: "center" },
            13: { cellWidth: 20, halign: "center" },
            14: { cellWidth: 26, halign: "left" },
            15: { cellWidth: 24, halign: "left" },
            16: { cellWidth: 20, halign: "center" },
            17: { cellWidth: 14, halign: "center" },
            18: { cellWidth: 32, halign: "left" }
          };
        } else if (cfg.id === "post_emb_print") {
          secHeaders = [
            "#",
            "Lot Number",
            "Garment Type",
            "Style",
            "Fabric",
            "Brand",
            "Total Pcs",
            "M/W/K",
            "Season",
            "Party Name",
            "Direct Stitching",
            "Cutting Date",
            "Emb/Print Date",
            "Days Pending",
            "Color Status",
            "Priority",
            "💬 User Remarks"
          ];
          secData = stageRows.map((r, idx) => [
            String(idx + 1),
            r.isRepeatedLot ? `★ ${r.lotNumber}` : r.lotNumber,
            r.garmentType || "—",
            r.style || "—",
            r.fabric || "—",
            r.brand || "—",
            (r.poQty || 0).toLocaleString(),
            r.section || "—",
            r.season || "—",
            r.partyName || "—",
            r.directStitching || "No",
            r.cuttingDate || "—",
            r.embPrintDate || "—",
            r.agingDays != null ? `${r.agingDays}d` : "—",
            r.pendingShade || "No Colour Pending",
            r.priority || "Normal",
            r.userRemarks || "—"
          ]);
          secColumnStyles = {
            0: { cellWidth: 8, halign: "center" },
            1: { cellWidth: 18, halign: "center", fontStyle: "bold" },
            2: { cellWidth: 22, halign: "center" },
            3: { cellWidth: 24, halign: "center" },
            4: { cellWidth: 22, halign: "center" },
            5: { cellWidth: 18, halign: "center" },
            6: { cellWidth: 18, halign: "center", fontStyle: "bold" },
            7: { cellWidth: 14, halign: "center" },
            8: { cellWidth: 15, halign: "center" },
            9: { cellWidth: 20, halign: "center" },
            10: { cellWidth: 16, halign: "center" },
            11: { cellWidth: 20, halign: "center" },
            12: { cellWidth: 20, halign: "center" },
            13: { cellWidth: 16, halign: "center" },
            14: { cellWidth: 26, halign: "left" },
            15: { cellWidth: 18, halign: "center" },
            16: { cellWidth: 32, halign: "left" }
          };
        } else if (cfg.id === "stitching") {
          secHeaders = [
            "#",
            "Lot Number",
            "Garment Type",
            "Style",
            "Fabric",
            "BRAND",
            "Total PCS",
            "Section",
            "Season",
            "PARTY NAME",
            "Direct Stitching",
            "Supervisor",
            "M/W/K",
            "Date of Issue",
            "Stitching Days",
            "Emb/Print Date",
            "WIP Status",
            "Pintu",
            "EA",
            "Status"
          ];
          secData = stageRows.map((r, idx) => [
            String(idx + 1),
            r.lotNumber,
            r.garmentType || "—",
            r.style || "—",
            r.fabric || "—",
            r.brand || "—",
            (r.poQty || 0).toLocaleString(),
            r.section || "—",
            r.season || "—",
            r.partyName || "—",
            r.directStitching || "no",
            r.supervisor || "Unassigned",
            r.mwk || r.section || "—",
            r.dateOfIssue || r.stageDate || "—",
            r.agingDays != null ? `${r.agingDays}d` : "—",
            r.embPrintDate || "—",
            r.wipStatus || "WIP",
            r.pintu || "—",
            r.ea || "—",
            r.status || "Pending"
          ]);
          secColumnStyles = {
            0: { cellWidth: 8, halign: "center" },
            1: { cellWidth: 18, halign: "center", fontStyle: "bold" },
            2: { cellWidth: 20, halign: "center" },
            3: { cellWidth: 22, halign: "center" },
            4: { cellWidth: 22, halign: "center" },
            5: { cellWidth: 18, halign: "center" },
            6: { cellWidth: 16, halign: "right", fontStyle: "bold" },
            7: { cellWidth: 15, halign: "center" },
            8: { cellWidth: 15, halign: "center" },
            9: { cellWidth: 20, halign: "center" },
            10: { cellWidth: 16, halign: "center" },
            11: { cellWidth: 20, halign: "center", fontStyle: "bold" },
            12: { cellWidth: 14, halign: "center" },
            13: { cellWidth: 18, halign: "center" },
            14: { cellWidth: 16, halign: "center", fontStyle: "bold" },
            15: { cellWidth: 18, halign: "center" },
            16: { cellWidth: 22, halign: "center" },
            17: { cellWidth: 20, halign: "center" },
            18: { cellWidth: 20, halign: "center" },
            19: { cellWidth: 16, halign: "center" }
          };
        } else {
          secHeaders = [
            "#",
            "Lot #",
            "Bottleneck / Variance Reason",
            "Severity",
            "Aging",
            "Garment Type",
            "Style",
            "Fabric",
            "Brand",
            "Party",
            "PO / Cut",
            "Done",
            "Pending",
            "Variance"
          ];
          secData = stageRows.map((r, idx) => [
            idx + 1,
            r.lotNumber,
            r.reason,
            r.severity,
            `${r.agingDays}d`,
            r.garmentType,
            r.style,
            r.fabric,
            r.brand,
            r.partyName,
            (r.poQty || 0).toLocaleString(),
            (r.stageDoneQty || 0).toLocaleString(),
            (r.pendingQty || 0).toLocaleString(),
            `${r.variancePcs > 0 ? "+" : ""}${r.variancePcs} (${r.variancePct}%)`
          ]);
          secColumnStyles = {
            0: { halign: "center", cellWidth: 10 },
            1: { halign: "center", fontStyle: "bold", cellWidth: 20 },
            2: { cellWidth: 70, fontStyle: "bold" },
            3: { halign: "center", fontStyle: "bold", cellWidth: 22 },
            4: { halign: "center", fontStyle: "bold", cellWidth: 14 },
            5: { cellWidth: 28 },
            6: { cellWidth: 28 },
            7: { cellWidth: 34 },
            8: { cellWidth: 24 },
            9: { cellWidth: 28 },
            10: { halign: "right", cellWidth: 20 },
            11: { halign: "right", cellWidth: 18 },
            12: { halign: "right", fontStyle: "bold", cellWidth: 20 },
            13: { halign: "right", cellWidth: 24 }
          };
        }

        const scaledSecColumnStyles = scaleColumnStyles(secColumnStyles, pageWidth - 28);

        autoTable(doc, {
          head: [secHeaders],
          body: secData,
          startY: currentY,
          tableWidth: pageWidth - 28,
          theme: "grid",
          styles: {
            fontSize: 7.5,
            cellPadding: 2,
            textColor: [0, 0, 0],
            font: "helvetica",
            valign: "middle"
          },
          headStyles: {
            fillColor: [15, 23, 42],
            textColor: [255, 255, 255],
            fontStyle: "bold",
            fontSize: 8,
            halign: "center"
          },
          alternateRowStyles: {
            fillColor: [248, 250, 252]
          },
          columnStyles: scaledSecColumnStyles,
          didParseCell: function (data) {
            if (data.section === "body") {
              data.cell.styles.textColor = [0, 0, 0];
              if (cfg.id === "cutting" && data.column.index === 13) {
                const n = parseFloat(data.cell.raw);
                if (!isNaN(n) && n > 2) {
                  data.cell.styles.fillColor = [254, 226, 226];
                }
              }
              if ((cfg.id === "embroidery" && data.column.index === 18) || (cfg.id === "printing" && data.column.index === 17)) {
                const n = parseFloat(data.cell.raw);
                if (!isNaN(n) && n > 5) {
                  data.cell.styles.fillColor = [239, 68, 68];
                  data.cell.styles.textColor = [255, 255, 255];
                  data.cell.styles.fontStyle = "bold";
                }
              }
            }
          },
          margin: { left: 14, right: 14, bottom: 16 }
        });

        currentY = doc.lastAutoTable.finalY + 10;
      });

      doc.save(`Factory_Suite_Pro_Combined_Variance_Report_${now.toISOString().slice(0, 10)}.pdf`);
    } catch (e) {
      console.error("PDF Export Error:", e);
      alert("Failed to export PDF report.");
    }
  };

  return (
    <div style={{ minHeight: "100vh", background: "#f8fafc", fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif" }}>
      {/* 🚀 LIVE ANIMATED TOUR FLOATING HUD OVERLAY */}
      {isAnimatedDownloading && (
        <div
          style={{
            position: "fixed",
            top: "20px",
            left: "50%",
            transform: "translateX(-50%)",
            zIndex: 9999,
            width: "92%",
            maxWidth: "660px",
            background: "rgba(15, 23, 42, 0.95)",
            backdropFilter: "blur(16px)",
            WebkitBackdropFilter: "blur(16px)",
            borderRadius: "16px",
            border: "2px solid #ef4444",
            boxShadow: "0 20px 50px rgba(0, 0, 0, 0.6), 0 0 30px rgba(239, 68, 68, 0.4)",
            padding: "18px 22px",
            color: "#ffffff",
            animation: "fadeInDown 0.3s ease-out"
          }}
        >
          {/* Header Row */}
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "12px" }}>
            <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
              <span style={{ fontSize: "1.8rem" }}>{animProgress.icon || "🚀"}</span>
              <div>
                <div style={{ fontSize: "0.72rem", fontWeight: 800, color: "#f87171", textTransform: "uppercase", letterSpacing: "0.08em" }}>
                  Live Factory Tour & Automated Downloader
                </div>
                <div style={{ fontSize: "1.05rem", fontWeight: 900, color: "#ffffff" }}>
                  {animProgress.title || "Processing Department Reports..."}
                </div>
              </div>
            </div>

            <button
              type="button"
              onClick={handleStopAnimation}
              style={{
                background: "rgba(239, 68, 68, 0.2)",
                border: "1px solid #ef4444",
                color: "#fca5a5",
                padding: "6px 12px",
                borderRadius: "8px",
                fontSize: "0.76rem",
                fontWeight: 800,
                cursor: "pointer",
                transition: "all 0.15s ease"
              }}
            >
              ✕ Cancel Tour
            </button>
          </div>

          {/* Live Progress Bar */}
          <div style={{ marginBottom: "12px" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: "0.75rem", fontWeight: 700, color: "#cbd5e1", marginBottom: "6px" }}>
              <span>
                Department {animProgress.step} of {animProgress.totalSteps}
              </span>
              <span>
                {Math.round(((animProgress.step) / (animProgress.totalSteps || 1)) * 100)}% Complete
              </span>
            </div>
            <div style={{ height: "8px", background: "rgba(255, 255, 255, 0.15)", borderRadius: "4px", overflow: "hidden" }}>
              <div
                style={{
                  width: `${Math.round(((animProgress.step) / (animProgress.totalSteps || 1)) * 100)}%`,
                  height: "100%",
                  background: "linear-gradient(90deg, #10b981 0%, #3b82f6 50%, #ec4899 100%)",
                  transition: "width 0.4s ease"
                }}
              />
            </div>
          </div>

          {/* Status Capsule */}
          <div
            style={{
              background: "rgba(255, 255, 255, 0.08)",
              border: "1px solid rgba(255, 255, 255, 0.12)",
              borderRadius: "10px",
              padding: "10px 14px",
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              fontSize: "0.82rem",
              fontWeight: 700,
              color: "#f1f5f9"
            }}
          >
            <span>{animProgress.statusText}</span>
            <span style={{ fontSize: "0.72rem", background: "#3b82f6", color: "#ffffff", padding: "2px 8px", borderRadius: "10px", textTransform: "uppercase", fontWeight: 800 }}>
              {animProgress.format.toUpperCase()}
            </span>
          </div>
        </div>
      )}

      {/* 1. TOP EXECUTIVE NAVBAR */}
      <header style={{ background: "linear-gradient(90deg, #7f1d1d 0%, #991b1b 50%, #b91c1c 100%)", color: "#ffffff", padding: "12px 24px", boxShadow: "0 4px 20px rgba(0,0,0,0.15)", position: "sticky", top: 0, zIndex: 100 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "12px" }}>
          {/* Left: Brand & Title */}
          <div style={{ display: "flex", alignItems: "center", gap: "14px" }}>
            <button
              type="button"
              onClick={handleGoBack}
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: "6px",
                background: "rgba(255, 255, 255, 0.15)",
                border: "1px solid rgba(255, 255, 255, 0.3)",
                color: "#ffffff",
                padding: "6px 14px",
                borderRadius: "8px",
                fontSize: "0.82rem",
                fontWeight: 700,
                cursor: "pointer",
                backdropFilter: "blur(6px)",
                transition: "all 0.15s ease"
              }}
            >
              ← Back
            </button>

            <div>
              <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                <span style={{ fontSize: "1.2rem" }}>🚨</span>
                <h1 style={{ margin: 0, fontSize: "1.2rem", fontWeight: 900, letterSpacing: "-0.02em", color: "#ffffff" }}>
                  Critical Red Zone & Variance Combined Report
                </h1>
                <span style={{ background: "#ef4444", color: "#ffffff", fontSize: "0.68rem", fontWeight: 800, padding: "2px 8px", borderRadius: "12px", letterSpacing: "0.04em", textTransform: "uppercase" }}>
                  Live Factory Command Center
                </span>
              </div>
              <p style={{ margin: "2px 0 0", fontSize: "0.76rem", color: "#fecaca" }}>
                Unified pipeline audit combining Red Zone & Pending lots across all 10 Most Critical Reports
              </p>
            </div>
          </div>

          {/* Right: Quick Links & Actions */}
          <div style={{ display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap" }}>
            <Link
              to="/dashboard"
              style={{ background: "rgba(255,255,255,0.12)", border: "1px solid rgba(255,255,255,0.25)", color: "#ffffff", padding: "6px 12px", borderRadius: "6px", fontSize: "0.78rem", fontWeight: 700, textDecoration: "none" }}
            >
              📊 Dashboard
            </Link>
            <Link
              to="/lot-logs"
              style={{ background: "rgba(255,255,255,0.12)", border: "1px solid rgba(255,255,255,0.25)", color: "#ffffff", padding: "6px 12px", borderRadius: "6px", fontSize: "0.78rem", fontWeight: 700, textDecoration: "none" }}
            >
              📜 Lot Logs
            </Link>
            <Link
              to="/production-flowchart"
              style={{ background: "rgba(255,255,255,0.12)", border: "1px solid rgba(255,255,255,0.25)", color: "#ffffff", padding: "6px 12px", borderRadius: "6px", fontSize: "0.78rem", fontWeight: 700, textDecoration: "none" }}
            >
              🗺️ Flow Poster
            </Link>

            <button
              type="button"
              onClick={() => loadData("refresh")}
              disabled={refreshing || loading}
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: "5px",
                background: "#ffffff",
                color: "#991b1b",
                border: "none",
                padding: "6px 14px",
                borderRadius: "6px",
                fontSize: "0.8rem",
                fontWeight: 800,
                cursor: refreshing ? "wait" : "pointer"
              }}
            >
              🔄 {refreshing ? "Refreshing..." : "Refresh"}
            </button>

            {/* Combined Excel Download */}
            <button
              type="button"
              onClick={handleExportExcel}
              title="Download 1 combined multi-sheet Excel file containing all stages"
              style={{ background: "#059669", color: "#ffffff", border: "none", padding: "6px 12px", borderRadius: "6px", fontSize: "0.78rem", fontWeight: 800, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: "4px" }}
            >
              📥 Combined Excel
            </button>

            {/* Live Animated Tour Download (Excel) */}
            <button
              type="button"
              onClick={() => runLiveAnimatedBatchDownload("excel")}
              disabled={isAnimatedDownloading}
              title="Live Visual Tour: Automatically opens Cutting, downloads report, goes to Embroidery, downloads, then Printing, etc. and returns back!"
              style={{
                background: "linear-gradient(135deg, #10b981 0%, #059669 100%)",
                color: "#ffffff",
                border: "1.5px solid #a7f3d0",
                padding: "6px 14px",
                borderRadius: "6px",
                fontSize: "0.8rem",
                fontWeight: 900,
                cursor: isAnimatedDownloading ? "wait" : "pointer",
                display: "inline-flex",
                alignItems: "center",
                gap: "6px",
                boxShadow: "0 4px 12px rgba(16, 185, 129, 0.4)",
                animation: isAnimatedDownloading ? "pulse 1.5s infinite" : "none"
              }}
            >
              🚀 Live Tour Excels ⚡
            </button>

            {/* Combined PDF Download */}
            <button
              type="button"
              onClick={handleExportPDF}
              title="Download 1 combined multi-section PDF file"
              style={{ background: "#dc2626", color: "#ffffff", border: "none", padding: "6px 12px", borderRadius: "6px", fontSize: "0.78rem", fontWeight: 800, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: "4px" }}
            >
              📄 Combined PDF
            </button>

            {/* Live Animated Tour Download (PDF) */}
            <button
              type="button"
              onClick={() => runLiveAnimatedBatchDownload("pdf")}
              disabled={isAnimatedDownloading}
              title="Live Visual Tour: Automatically opens Cutting, downloads PDF, goes to Embroidery, downloads, then Printing, etc. and returns back!"
              style={{
                background: "linear-gradient(135deg, #ef4444 0%, #b91c1c 100%)",
                color: "#ffffff",
                border: "1.5px solid #fca5a5",
                padding: "6px 14px",
                borderRadius: "6px",
                fontSize: "0.8rem",
                fontWeight: 900,
                cursor: isAnimatedDownloading ? "wait" : "pointer",
                display: "inline-flex",
                alignItems: "center",
                gap: "6px",
                boxShadow: "0 4px 12px rgba(220, 38, 38, 0.4)"
              }}
            >
              🚀 Live Tour PDFs ⚡
            </button>
          </div>
        </div>
      </header>

      {/* 2. MAIN CONTAINER */}
      <main style={{ maxWidth: "1800px", margin: "0 auto", padding: "20px" }}>
        {/* Loading / Refreshing Bar */}
        {(loading || refreshing) && (
          <div style={{ background: "#ffffff", padding: "16px", borderRadius: "12px", marginBottom: "20px", border: "1px solid #fed7aa", boxShadow: "0 4px 15px rgba(234, 88, 12, 0.08)" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "8px" }}>
              <span style={{ fontSize: "0.85rem", fontWeight: 700, color: "#c2410c" }}>⏳ {loadingMessage}</span>
              <span style={{ fontSize: "0.82rem", fontWeight: 800, color: "#ea580c" }}>{loadingProgress}%</span>
            </div>
            <div style={{ height: "6px", background: "#fed7aa", borderRadius: "3px", overflow: "hidden" }}>
              <div style={{ width: `${loadingProgress}%`, height: "100%", background: "linear-gradient(90deg, #ea580c, #dc2626)", transition: "width 0.3s ease" }} />
            </div>
          </div>
        )}

        {/* Error Alert */}
        {error && (
          <div style={{ background: "#fef2f2", border: "1.5px solid #fecaca", padding: "14px 18px", borderRadius: "10px", color: "#dc2626", fontWeight: 700, fontSize: "0.88rem", marginBottom: "20px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span>⚠️ {error}</span>
            <button type="button" onClick={() => loadData("refresh")} style={{ background: "#dc2626", color: "#ffffff", border: "none", padding: "4px 12px", borderRadius: "6px", fontSize: "0.78rem", fontWeight: 700, cursor: "pointer" }}>
              Retry
            </button>
          </div>
        )}

        {/* 3. EXECUTIVE KPI CARDS */}
        <section style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: "14px", marginBottom: "20px" }}>
          {/* Card 1: Total Red Zone Lots */}
          <div style={{ background: "#ffffff", border: "1.5px solid #fecaca", borderRadius: "14px", padding: "16px", boxShadow: "0 4px 14px rgba(220, 38, 38, 0.06)", position: "relative", overflow: "hidden" }}>
            <div style={{ position: "absolute", top: 0, left: 0, right: 0, height: "4px", background: "#dc2626" }} />
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
              <div>
                <div style={{ fontSize: "0.72rem", fontWeight: 800, color: "#991b1b", textTransform: "uppercase", letterSpacing: "0.04em" }}>
                  Total Red Zone Lots
                </div>
                <div style={{ fontSize: "1.8rem", fontWeight: 900, color: "#1e293b", marginTop: "4px" }}>
                  {analytics.totalLots}
                </div>
              </div>
              <span style={{ fontSize: "1.8rem" }}>🚨</span>
            </div>
            <div style={{ marginTop: "8px", fontSize: "0.75rem", color: "#64748b", fontWeight: 600 }}>
              🔴 {analytics.criticalCount} Critical &nbsp;•&nbsp; 🟠 {analytics.highCount} High &nbsp;•&nbsp; 🟡 {analytics.moderateCount} Mod
            </div>
          </div>

          {/* Card 2: Total Pending Pieces
          <div style={{ background: "#ffffff", border: "1.5px solid #fed7aa", borderRadius: "14px", padding: "16px", boxShadow: "0 4px 14px rgba(234, 88, 12, 0.06)", position: "relative", overflow: "hidden" }}>
            <div style={{ position: "absolute", top: 0, left: 0, right: 0, height: "4px", background: "#ea580c" }} />
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
              <div>
                <div style={{ fontSize: "0.72rem", fontWeight: 800, color: "#c2410c", textTransform: "uppercase", letterSpacing: "0.04em" }}>
                  Total Pending Pieces
                </div>
                <div style={{ fontSize: "1.8rem", fontWeight: 900, color: "#1e293b", marginTop: "4px" }}>
                  {analytics.totalPendingPcs.toLocaleString()}
                </div>
              </div>
              <span style={{ fontSize: "1.8rem" }}>⚠️</span>
            </div>
            <div style={{ marginTop: "8px", fontSize: "0.75rem", color: "#64748b", fontWeight: 600 }}>
              Blocked in bottleneck stages across pipeline
            </div>
          </div> */}

          {/* Card 3: Total Variance / Deficit */}
          {/* <div style={{ background: "#ffffff", border: "1.5px solid #ddd6fe", borderRadius: "14px", padding: "16px", boxShadow: "0 4px 14px rgba(124, 58, 237, 0.06)", position: "relative", overflow: "hidden" }}>
            <div style={{ position: "absolute", top: 0, left: 0, right: 0, height: "4px", background: "#7c3aed" }} />
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
              <div>
                <div style={{ fontSize: "0.72rem", fontWeight: 800, color: "#6d28d9", textTransform: "uppercase", letterSpacing: "0.04em" }}>
                  Total Piece Deficit / Variance
                </div>
                <div style={{ fontSize: "1.8rem", fontWeight: 900, color: "#1e293b", marginTop: "4px" }}>
                  {analytics.totalVariancePcs.toLocaleString()} Pcs
                </div>
              </div>
              <span style={{ fontSize: "1.8rem" }}>📉</span>
            </div>
            <div style={{ marginTop: "8px", fontSize: "0.75rem", color: "#64748b", fontWeight: 600 }}>
              Discrepancy vs PO & Cutting counts
            </div>
          </div> */}

          {/* Card 4: Max Pipeline Aging */}
          <div style={{ background: "#ffffff", border: "1.5px solid #fde68a", borderRadius: "14px", padding: "16px", boxShadow: "0 4px 14px rgba(217, 119, 6, 0.06)", position: "relative", overflow: "hidden" }}>
            <div style={{ position: "absolute", top: 0, left: 0, right: 0, height: "4px", background: "#d97706" }} />
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
              <div>
                <div style={{ fontSize: "0.72rem", fontWeight: 800, color: "#b45309", textTransform: "uppercase", letterSpacing: "0.04em" }}>
                  Pipeline Aging Delays
                </div>
                <div style={{ fontSize: "1.8rem", fontWeight: 900, color: "#1e293b", marginTop: "4px" }}>
                  {analytics.maxAging} <span style={{ fontSize: "1rem", fontWeight: 700, color: "#64748b" }}>Days Max</span>
                </div>
              </div>
              <span style={{ fontSize: "1.8rem" }}>⏱️</span>
            </div>
            <div style={{ marginTop: "8px", fontSize: "0.75rem", color: "#64748b", fontWeight: 600 }}>
              Average Pipeline Lag: {analytics.avgAging} Days
            </div>
          </div>

          {/* Card 5: Highest Bottleneck Stage */}
          <div style={{ background: "#ffffff", border: "1.5px solid #cbd5e1", borderRadius: "14px", padding: "16px", boxShadow: "0 4px 14px rgba(15, 23, 42, 0.06)", position: "relative", overflow: "hidden" }}>
            <div style={{ position: "absolute", top: 0, left: 0, right: 0, height: "4px", background: "#0f172a" }} />
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
              <div>
                <div style={{ fontSize: "0.72rem", fontWeight: 800, color: "#475569", textTransform: "uppercase", letterSpacing: "0.04em" }}>
                  Top Bottleneck Stage
                </div>
                <div style={{ fontSize: "1.15rem", fontWeight: 900, color: "#1e293b", marginTop: "6px", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                  {analytics.topStage}
                </div>
              </div>
              <span style={{ fontSize: "1.8rem" }}>🏢</span>
            </div>
            <div style={{ marginTop: "8px", fontSize: "0.75rem", color: "#dc2626", fontWeight: 700 }}>
              {analytics.topStageCount} Delayed Lots currently stalled
            </div>
          </div>
        </section>

        {/* 4. STAGE SELECTOR TABS (10 CRITICAL REPORTS BAR) */}
        <section style={{ marginBottom: "18px", background: "#ffffff", padding: "12px", borderRadius: "14px", border: "1px solid #e2e8f0", boxShadow: "0 2px 8px rgba(0,0,0,0.04)" }}>
          <div style={{ fontSize: "0.72rem", fontWeight: 800, color: "#64748b", textTransform: "uppercase", letterSpacing: "0.04em", marginBottom: "8px" }}>
            Filter by Critical Factory Stage ({CRITICAL_STAGES.length - 1} Departments)
          </div>
          <div style={{ display: "flex", gap: "8px", overflowX: "auto", paddingBottom: "4px" }}>
            {CRITICAL_STAGES.map((stage) => {
              const isActive = activeStageTab === stage.id;
              const count = stageCounts[stage.id] || 0;
              return (
                <button
                  key={stage.id}
                  type="button"
                  onClick={() => {
                    setActiveStageTab(stage.id);
                    setCurrentPage(1);
                  }}
                  style={{
                    display: "inline-flex",
                    alignItems: "center",
                    gap: "6px",
                    padding: "8px 14px",
                    borderRadius: "10px",
                    border: `1.5px solid ${isActive ? stage.badgeColor || "#dc2626" : "#e2e8f0"}`,
                    background: isActive ? stage.badgeColor || "#dc2626" : "#ffffff",
                    color: isActive ? "#ffffff" : "#334155",
                    fontSize: "0.8rem",
                    fontWeight: isActive ? 800 : 600,
                    cursor: "pointer",
                    whiteSpace: "nowrap",
                    transition: "all 0.15s ease",
                    boxShadow: isActive ? "0 4px 12px rgba(0,0,0,0.12)" : "none"
                  }}
                >
                  <span>{stage.icon}</span>
                  <span>{stage.shortLabel}</span>
                  <span
                    style={{
                      background: isActive ? "rgba(255, 255, 255, 0.25)" : "#f1f5f9",
                      color: isActive ? "#ffffff" : count > 0 ? "#dc2626" : "#64748b",
                      padding: "2px 6px",
                      borderRadius: "10px",
                      fontSize: "0.7rem",
                      fontWeight: 800
                    }}
                  >
                    {count}
                  </span>
                </button>
              );
            })}
          </div>
        </section>

        {/* 5. MULTI-DIMENSION FILTER TOOLBAR */}
        <section style={{ background: "#ffffff", padding: "16px", borderRadius: "14px", border: "1px solid #e2e8f0", marginBottom: "20px", boxShadow: "0 2px 8px rgba(0,0,0,0.04)" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "12px", flexWrap: "wrap", gap: "8px" }}>
            <div style={{ fontSize: "0.82rem", fontWeight: 800, color: "#1e293b", display: "flex", alignItems: "center", gap: "6px" }}>
              <span>🔍</span>
              <span>Filter Variance & Bottleneck Records</span>
              {hasActiveFilters && (
                <span style={{ background: "#fee2e2", color: "#dc2626", fontSize: "0.7rem", fontWeight: 800, padding: "2px 8px", borderRadius: "10px" }}>
                  Active Filters
                </span>
              )}
            </div>

            {hasActiveFilters && (
              <button
                type="button"
                onClick={clearAllFilters}
                style={{
                  background: "#fee2e2",
                  color: "#dc2626",
                  border: "1px solid #fecaca",
                  padding: "5px 12px",
                  borderRadius: "6px",
                  fontSize: "0.75rem",
                  fontWeight: 800,
                  cursor: "pointer"
                }}
              >
                Clear All Filters ✕
              </button>
            )}
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: "12px" }}>
            {/* Financial Year Filter */}
            <div style={{ minWidth: "160px" }}>
              <label style={{ display: "block", fontSize: "0.72rem", fontWeight: 700, color: "#475569", marginBottom: "4px", textTransform: "uppercase", letterSpacing: "0.03em" }}>
                Financial Year
              </label>
              <select
                value={financialYearFilter}
                onChange={(e) => {
                  setFinancialYearFilter(e.target.value);
                  setCurrentPage(1);
                }}
                style={{
                  width: "100%",
                  padding: "8px 12px",
                  fontSize: "0.82rem",
                  fontWeight: 700,
                  color: "#1d4ed8",
                  background: "#eff6ff",
                  borderRadius: "8px",
                  border: "1.5px solid #3b82f6",
                  outline: "none",
                  boxSizing: "border-box",
                  cursor: "pointer"
                }}
              >
                <option value="ALL">All Financial Years (Global)</option>
                {availableFinancialYears.map((fy) => (
                  <option key={fy} value={fy}>
                    FY {fy} {fy === getCurrentFinancialYear() ? "(Current)" : ""}
                  </option>
                ))}
              </select>
            </div>

            {/* Search Input */}
            <div style={{ minWidth: "180px", flex: "1 1 180px" }}>
              <label style={{ display: "block", fontSize: "0.72rem", fontWeight: 700, color: "#475569", marginBottom: "4px", textTransform: "uppercase", letterSpacing: "0.03em" }}>
                Search Lots / Style
              </label>
              <input
                type="text"
                placeholder="Search Lot, Style, Party..."
                value={searchTerm}
                onChange={(e) => {
                  setSearchTerm(e.target.value);
                  setCurrentPage(1);
                }}
                style={{
                  width: "100%",
                  padding: "8px 12px",
                  fontSize: "0.82rem",
                  borderRadius: "8px",
                  border: "1.5px solid #cbd5e1",
                  outline: "none",
                  boxSizing: "border-box"
                }}
              />
            </div>

            {/* Severity Filter */}
            <div style={{ minWidth: "150px" }}>
              <label style={{ display: "block", fontSize: "0.72rem", fontWeight: 700, color: "#475569", marginBottom: "4px", textTransform: "uppercase", letterSpacing: "0.03em" }}>
                Severity Level
              </label>
              <select
                value={filterSeverity}
                onChange={(e) => {
                  setFilterSeverity(e.target.value);
                  setCurrentPage(1);
                }}
                style={{
                  width: "100%",
                  padding: "8px 12px",
                  fontSize: "0.82rem",
                  fontWeight: filterSeverity !== "all" ? 700 : 500,
                  color: filterSeverity !== "all" ? "#dc2626" : "#334155",
                  background: filterSeverity !== "all" ? "#fef2f2" : "#ffffff",
                  borderRadius: "8px",
                  border: `1.5px solid ${filterSeverity !== "all" ? "#dc2626" : "#cbd5e1"}`,
                  outline: "none",
                  boxSizing: "border-box"
                }}
              >
                <option value="all">All Red Zone Lots</option>
                <option value="CRITICAL">🔴 Critical Red Zone (&gt; 5 Days)</option>
                <option value="HIGH">🔴 Red Zone (3-5 Days)</option>
                <option value="MODERATE">🔴 Red Zone (&gt; 2 Days)</option>
              </select>
            </div>

            {/* Garment Multi-Select */}
            <MultiSelectDropdown
              label="Garment Type"
              options={uniqueGarments}
              selectedValues={filterGarments}
              onChange={(vals) => {
                setFilterGarments(vals);
                setCurrentPage(1);
              }}
              placeholder="All Garments"
            />

            {/* Brand Multi-Select */}
            <MultiSelectDropdown
              label="Brand"
              options={uniqueBrands}
              selectedValues={filterBrands}
              onChange={(vals) => {
                setFilterBrands(vals);
                setCurrentPage(1);
              }}
              placeholder="All Brands"
            />

            {/* Fabric Multi-Select */}
            <MultiSelectDropdown
              label="Fabric"
              options={uniqueFabrics}
              selectedValues={filterFabrics}
              onChange={(vals) => {
                setFilterFabrics(vals);
                setCurrentPage(1);
              }}
              placeholder="All Fabrics"
            />

            {/* Party Name Multi-Select */}
            <MultiSelectDropdown
              label="Party Name"
              options={uniqueParties}
              selectedValues={filterParties}
              onChange={(vals) => {
                setFilterParties(vals);
                setCurrentPage(1);
              }}
              placeholder="All Parties"
            />

            {/* Min Aging Filter */}
            <div style={{ minWidth: "120px" }}>
              <label style={{ display: "block", fontSize: "0.72rem", fontWeight: 700, color: "#475569", marginBottom: "4px", textTransform: "uppercase", letterSpacing: "0.03em" }}>
                Min Aging (Days)
              </label>
              <input
                type="number"
                min="0"
                placeholder="e.g. 3"
                value={minAgingDays}
                onChange={(e) => {
                  setMinAgingDays(e.target.value);
                  setCurrentPage(1);
                }}
                style={{
                  width: "100%",
                  padding: "8px 12px",
                  fontSize: "0.82rem",
                  borderRadius: "8px",
                  border: "1.5px solid #cbd5e1",
                  outline: "none",
                  boxSizing: "border-box"
                }}
              />
            </div>
          </div>
        </section>

        {/* 6. MASTER BOTTLENECK & VARIANCE TABLE */}
        <section style={{ background: "#ffffff", borderRadius: "14px", border: "1px solid #e2e8f0", boxShadow: "0 4px 16px rgba(0,0,0,0.06)", overflow: "hidden" }}>
          <div style={{ padding: "14px 18px", borderBottom: "1px solid #e2e8f0", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "10px", background: "#f8fafc" }}>
            <div style={{ display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap" }}>
              <span style={{ fontSize: "1.1rem" }}>
                {activeStageTab === "cutting" ? "✂️" : activeStageTab === "embroidery" ? "🧵" : activeStageTab === "printing" ? "🖨️" : activeStageTab === "packing_handover" ? "📦⏳" : "📋"}
              </span>
              <span style={{ fontSize: "0.95rem", fontWeight: 800, color: "#0f172a" }}>
                {activeStageTab === "cutting"
                  ? "Cutting Production Red Zone Report (> 2 Days Delayed)"
                  : activeStageTab === "embroidery"
                    ? "Embroidery Challan Red Zone Report (> 5 Days Delayed)"
                    : activeStageTab === "printing"
                      ? "Printing Challan Red Zone Report (> 5 Days Delayed)"
                      : activeStageTab === "post_emb_print"
                        ? "Post-EMB/Print Handover Red Zone Report (Pending Issue to Stitching > 2 Days Delayed)"
                        : activeStageTab === "stitching"
                          ? "Stitching WIP Red Zone Report (Sewing Line Operations > 15 Days Delayed)"
                          : activeStageTab === "packing_handover"
                            ? "Packing Pending to Issue Red Zone Report (Completed Lots Ready for Packing > 2 Days Delayed)"
                            : "Master Critical Variance & Red Zone Command Center"}
              </span>
              <span style={{ background: "#dc2626", color: "#ffffff", fontSize: "0.74rem", fontWeight: 800, padding: "3px 10px", borderRadius: "12px" }}>
                {filteredRows.length} Red Zone Lots
              </span>
            </div>

            {/* Quick Department Export Buttons */}
            <div style={{ display: "flex", alignItems: "center", gap: "10px", flexWrap: "wrap" }}>
              {activeStageTab === "cutting" && (
                <>
                  <button
                    type="button"
                    onClick={() => exportCuttingExcel(filteredRows)}
                    style={{ background: "#059669", color: "#ffffff", border: "none", padding: "5px 12px", borderRadius: "6px", fontSize: "0.75rem", fontWeight: 800, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: "4px", boxShadow: "0 2px 6px rgba(5,150,105,0.3)" }}
                  >
                    📥 Export Cutting Excel
                  </button>
                  <button
                    type="button"
                    onClick={() => exportCuttingPDF(filteredRows)}
                    style={{ background: "#dc2626", color: "#ffffff", border: "none", padding: "5px 12px", borderRadius: "6px", fontSize: "0.75rem", fontWeight: 800, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: "4px", boxShadow: "0 2px 6px rgba(220,38,38,0.3)" }}
                  >
                    📄 Export Cutting PDF
                  </button>
                </>
              )}

              {activeStageTab === "embroidery" && (
                <>
                  <button
                    type="button"
                    onClick={() => exportEmbroideryExcel(filteredRows)}
                    style={{ background: "#7c3aed", color: "#ffffff", border: "none", padding: "5px 12px", borderRadius: "6px", fontSize: "0.75rem", fontWeight: 800, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: "4px", boxShadow: "0 2px 6px rgba(124,58,237,0.3)" }}
                  >
                    📥 Export Embroidery Excel
                  </button>
                  <button
                    type="button"
                    onClick={() => exportEmbroideryPDF(filteredRows)}
                    style={{ background: "#dc2626", color: "#ffffff", border: "none", padding: "5px 12px", borderRadius: "6px", fontSize: "0.75rem", fontWeight: 800, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: "4px", boxShadow: "0 2px 6px rgba(220,38,38,0.3)" }}
                  >
                    📄 Export Embroidery PDF
                  </button>
                </>
              )}

              {activeStageTab === "printing" && (
                <>
                  <button
                    type="button"
                    onClick={() => exportPrintingExcel(filteredRows)}
                    style={{ background: "#db2777", color: "#ffffff", border: "none", padding: "5px 12px", borderRadius: "6px", fontSize: "0.75rem", fontWeight: 800, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: "4px", boxShadow: "0 2px 6px rgba(219,39,119,0.3)" }}
                  >
                    📥 Export Printing Excel
                  </button>
                  <button
                    type="button"
                    onClick={() => exportPrintingPDF(filteredRows)}
                    style={{ background: "#dc2626", color: "#ffffff", border: "none", padding: "5px 12px", borderRadius: "6px", fontSize: "0.75rem", fontWeight: 800, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: "4px", boxShadow: "0 2px 6px rgba(220,38,38,0.3)" }}
                  >
                    📄 Export Printing PDF
                  </button>
                </>
              )}

              {activeStageTab === "post_emb_print" && (
                <>
                  <button
                    type="button"
                    onClick={() => exportPostEmbPrintExcel(filteredRows)}
                    style={{ background: "#ea580c", color: "#ffffff", border: "none", padding: "5px 12px", borderRadius: "6px", fontSize: "0.75rem", fontWeight: 800, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: "4px", boxShadow: "0 2px 6px rgba(234,88,12,0.3)" }}
                  >
                    📥 Export Post-EMB/Print Excel
                  </button>
                  <button
                    type="button"
                    onClick={() => exportPostEmbPrintPDF(filteredRows)}
                    style={{ background: "#dc2626", color: "#ffffff", border: "none", padding: "5px 12px", borderRadius: "6px", fontSize: "0.75rem", fontWeight: 800, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: "4px", boxShadow: "0 2px 6px rgba(220,38,38,0.3)" }}
                  >
                    📄 Export Post-EMB/Print PDF
                  </button>
                </>
              )}

              {activeStageTab === "stitching" && (
                <>
                  <button
                    type="button"
                    onClick={() => exportStitchingExcel(filteredRows)}
                    style={{ background: "#0284c7", color: "#ffffff", border: "none", padding: "5px 12px", borderRadius: "6px", fontSize: "0.75rem", fontWeight: 800, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: "4px", boxShadow: "0 2px 6px rgba(2,132,199,0.3)" }}
                  >
                    📥 Export Stitching Excel
                  </button>
                  <button
                    type="button"
                    onClick={() => exportStitchingPDF(filteredRows)}
                    style={{ background: "#dc2626", color: "#ffffff", border: "none", padding: "5px 12px", borderRadius: "6px", fontSize: "0.75rem", fontWeight: 800, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: "4px", boxShadow: "0 2px 6px rgba(220,38,38,0.3)" }}
                  >
                    📄 Export Stitching PDF
                  </button>
                </>
              )}

              {activeStageTab === "packing_handover" && (
                <>
                  <button
                    type="button"
                    onClick={() => exportPackingHandoverExcel(filteredRows)}
                    style={{ background: "#c2410c", color: "#ffffff", border: "none", padding: "5px 12px", borderRadius: "6px", fontSize: "0.75rem", fontWeight: 800, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: "4px", boxShadow: "0 2px 6px rgba(194,65,12,0.3)" }}
                  >
                    📥 Export Packing Pending to Issue Excel
                  </button>
                  <button
                    type="button"
                    onClick={() => exportPackingHandoverPDF(filteredRows)}
                    style={{ background: "#dc2626", color: "#ffffff", border: "none", padding: "5px 12px", borderRadius: "6px", fontSize: "0.75rem", fontWeight: 800, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: "4px", boxShadow: "0 2px 6px rgba(220,38,38,0.3)" }}
                  >
                    📄 Export Packing Pending to Issue PDF
                  </button>
                </>
              )}

              <div style={{ fontSize: "0.78rem", color: "#64748b", fontWeight: 600 }}>
                Rows per page:
                <select
                  value={pageSize}
                  onChange={(e) => {
                    setPageSize(Number(e.target.value));
                    setCurrentPage(1);
                  }}
                  style={{ marginLeft: "6px", padding: "4px 8px", borderRadius: "6px", border: "1px solid #cbd5e1", fontSize: "0.78rem", fontWeight: 700 }}
                >
                  <option value={15}>15</option>
                  <option value={25}>25</option>
                  <option value={50}>50</option>
                  <option value={100}>100</option>
                </select>
              </div>
            </div>
          </div>

          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", textAlign: "left", fontSize: "0.82rem" }}>
              {/* ========================================================= */}
              {/* TABLE HEADERS - DYNAMIC ACCORDING TO ACTIVE DEPARTMENT   */}
              {/* ========================================================= */}
              <thead>
                {/* 1. CUTTING REPORT HEADERS (18 COLUMNS) */}
                {activeStageTab === "cutting" ? (
                  <tr style={{ background: "#064e3b", color: "#ffffff", userSelect: "none" }}>
                    <th style={{ padding: "12px 8px", textAlign: "center", width: "35px", fontSize: "0.74rem", fontWeight: 800 }}>#</th>
                    <th onClick={() => handleSort("lotNumber")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800, whiteSpace: "nowrap" }}>
                      Lot No {sortConfig.key === "lotNumber" ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
                    </th>
                    <th onClick={() => handleSort("garmentType")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Garment Type</th>
                    <th onClick={() => handleSort("style")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Style</th>
                    <th onClick={() => handleSort("fabric")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Fabric</th>
                    <th onClick={() => handleSort("brand")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Brand</th>
                    <th onClick={() => handleSort("poQty")} style={{ padding: "12px 10px", textAlign: "right", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>
                      Total Qty {sortConfig.key === "poQty" ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
                    </th>
                    <th onClick={() => handleSort("section")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Section</th>
                    <th onClick={() => handleSort("season")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Season</th>
                    <th onClick={() => handleSort("partyName")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Party Name</th>
                    <th onClick={() => handleSort("directStitching")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Direct Stitching</th>
                    <th onClick={() => handleSort("jobOrderNo")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Job Order No</th>
                    <th onClick={() => handleSort("stageDate")} style={{ padding: "12px 10px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Date</th>
                    <th onClick={() => handleSort("agingDays")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800, background: "#047857", color: "#fef08a" }}>
                      Days after PO issue {sortConfig.key === "agingDays" ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
                    </th>
                    <th style={{ padding: "12px 10px", fontSize: "0.74rem", fontWeight: 800 }}>Pending Shade</th>
                    <th style={{ padding: "12px 10px", textAlign: "center", fontSize: "0.74rem", fontWeight: 800 }}>Cutting Date</th>
                    <th style={{ padding: "12px 10px", textAlign: "center", fontSize: "0.74rem", fontWeight: 800 }}>Cutting Scanned</th>
                    <th style={{ padding: "12px 10px", fontSize: "0.74rem", fontWeight: 800 }}>💬 User Remarks</th>
                    <th style={{ padding: "12px 10px", textAlign: "center", fontSize: "0.74rem", fontWeight: 800 }}>Action</th>
                  </tr>
                ) : activeStageTab === "embroidery" ? (
                  /* 2. EMBROIDERY REPORT HEADERS (19 COLUMNS) */
                  <tr style={{ background: "#4c1d95", color: "#ffffff", userSelect: "none" }}>
                    <th style={{ padding: "12px 8px", textAlign: "center", width: "35px", fontSize: "0.74rem", fontWeight: 800 }}>S. No</th>
                    <th onClick={() => handleSort("lotNumber")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800, whiteSpace: "nowrap" }}>
                      Lot No. {sortConfig.key === "lotNumber" ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
                    </th>
                    <th onClick={() => handleSort("garmentType")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Garment Type</th>
                    <th onClick={() => handleSort("style")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Style</th>
                    <th onClick={() => handleSort("fabric")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Fabric</th>
                    <th onClick={() => handleSort("brand")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Brand</th>
                    <th onClick={() => handleSort("poQty")} style={{ padding: "12px 10px", textAlign: "right", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>
                      Challan Qty {sortConfig.key === "poQty" ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
                    </th>
                    <th onClick={() => handleSort("section")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Section (M/W/K)</th>
                    <th onClick={() => handleSort("season")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Season</th>
                    <th onClick={() => handleSort("partyName")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Party Name</th>
                    <th onClick={() => handleSort("directStitching")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Direct Stitching</th>
                    <th onClick={() => handleSort("embParty")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Emb Party</th>
                    <th onClick={() => handleSort("stageDate")} style={{ padding: "12px 10px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Challan Date</th>
                    <th onClick={() => handleSort("embStatus")} style={{ padding: "12px 10px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Emb Status</th>
                    <th style={{ padding: "12px 10px", fontSize: "0.74rem", fontWeight: 800 }}>Pending Shade</th>
                    <th style={{ padding: "12px 10px", textAlign: "center", fontSize: "0.74rem", fontWeight: 800 }}>Emb Done Date</th>
                    <th style={{ padding: "12px 10px", textAlign: "center", fontSize: "0.74rem", fontWeight: 800 }}>Received Date</th>
                    <th onClick={() => handleSort("agingDays")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800, background: "#6d28d9", color: "#fef08a" }}>
                      Days {sortConfig.key === "agingDays" ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
                    </th>
                    <th style={{ padding: "12px 10px", fontSize: "0.74rem", fontWeight: 800 }}>HOD Remarks</th>
                    <th style={{ padding: "12px 10px", textAlign: "center", fontSize: "0.74rem", fontWeight: 800 }}>Action</th>
                  </tr>
                ) : activeStageTab === "printing" ? (
                  /* 3. PRINTING REPORT HEADERS (18 COLUMNS) */
                  <tr style={{ background: "#831843", color: "#ffffff", userSelect: "none" }}>
                    <th style={{ padding: "12px 8px", textAlign: "center", width: "35px", fontSize: "0.74rem", fontWeight: 800 }}>S. No</th>
                    <th onClick={() => handleSort("lotNumber")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800, whiteSpace: "nowrap" }}>
                      Lot No. {sortConfig.key === "lotNumber" ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
                    </th>
                    <th onClick={() => handleSort("garmentType")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Garment Type</th>
                    <th onClick={() => handleSort("style")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Style</th>
                    <th onClick={() => handleSort("fabric")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Fabric</th>
                    <th onClick={() => handleSort("brand")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Brand</th>
                    <th onClick={() => handleSort("poQty")} style={{ padding: "12px 10px", textAlign: "right", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>
                      Challan Qty {sortConfig.key === "poQty" ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
                    </th>
                    <th onClick={() => handleSort("section")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Section (M/W/K)</th>
                    <th onClick={() => handleSort("season")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Season</th>
                    <th onClick={() => handleSort("partyName")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Party Name</th>
                    <th onClick={() => handleSort("directStitching")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Direct Stitching</th>
                    <th onClick={() => handleSort("printingParty")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Printing Party</th>
                    <th onClick={() => handleSort("stageDate")} style={{ padding: "12px 10px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Challan Date</th>
                    <th onClick={() => handleSort("printingStatus")} style={{ padding: "12px 10px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Printing Status</th>
                    <th style={{ padding: "12px 10px", fontSize: "0.74rem", fontWeight: 800 }}>Pending Shade</th>
                    <th style={{ padding: "12px 10px", textAlign: "center", fontSize: "0.74rem", fontWeight: 800 }}>Printing Done Date</th>
                    <th onClick={() => handleSort("agingDays")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800, background: "#9d174d", color: "#fef08a" }}>
                      Days {sortConfig.key === "agingDays" ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
                    </th>
                    <th style={{ padding: "12px 10px", fontSize: "0.74rem", fontWeight: 800 }}>HOD Remarks</th>
                    <th style={{ padding: "12px 10px", textAlign: "center", fontSize: "0.74rem", fontWeight: 800 }}>Action</th>
                  </tr>
                ) : activeStageTab === "post_emb_print" ? (
                  /* 4. POST-EMB/PRINT REPORT HEADERS (18 COLUMNS - MATCHING PENDINGISSUE.JS) */
                  <tr style={{ background: "#7c2d12", color: "#ffffff", userSelect: "none" }}>
                    <th style={{ padding: "12px 8px", textAlign: "center", width: "35px", fontSize: "0.74rem", fontWeight: 800 }}>#</th>
                    <th onClick={() => handleSort("lotNumber")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800, whiteSpace: "nowrap" }}>
                      Lot Number {sortConfig.key === "lotNumber" ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
                    </th>
                    <th onClick={() => handleSort("garmentType")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Garment Type</th>
                    <th onClick={() => handleSort("style")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Style</th>
                    <th onClick={() => handleSort("fabric")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Fabric</th>
                    <th onClick={() => handleSort("brand")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Brand</th>
                    <th onClick={() => handleSort("poQty")} style={{ padding: "12px 10px", textAlign: "right", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>
                      Total Pcs {sortConfig.key === "poQty" ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
                    </th>
                    <th onClick={() => handleSort("section")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>M/W/K</th>
                    <th onClick={() => handleSort("season")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Season</th>
                    <th onClick={() => handleSort("partyName")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Party Name</th>
                    <th onClick={() => handleSort("directStitching")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Direct Stitching</th>
                    <th onClick={() => handleSort("cuttingDate")} style={{ padding: "12px 10px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Cutting Date</th>
                    <th onClick={() => handleSort("embPrintDate")} style={{ padding: "12px 10px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Emb/Printing Date</th>
                    <th onClick={() => handleSort("agingDays")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800, background: "#9a3412", color: "#fef08a" }}>
                      Days Pending {sortConfig.key === "agingDays" ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
                    </th>
                    <th style={{ padding: "12px 10px", fontSize: "0.74rem", fontWeight: 800 }}>Color Status</th>
                    <th style={{ padding: "12px 10px", textAlign: "center", fontSize: "0.74rem", fontWeight: 800 }}>Priority</th>
                    <th style={{ padding: "12px 10px", fontSize: "0.74rem", fontWeight: 800 }}>💬 User Remarks</th>
                    <th style={{ padding: "12px 10px", textAlign: "center", fontSize: "0.74rem", fontWeight: 800 }}>Action</th>
                  </tr>
                ) : activeStageTab === "stitching" ? (
                  /* 6. STITCHING WIP REPORT HEADERS (21 COLUMNS - MATCHING STITCHINGCOMPLETED.JS) */
                  <tr style={{ background: "#1e1b4b", color: "#ffffff", userSelect: "none" }}>
                    <th style={{ padding: "12px 8px", textAlign: "center", width: "35px", fontSize: "0.74rem", fontWeight: 800 }}>#</th>
                    <th onClick={() => handleSort("lotNumber")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800, whiteSpace: "nowrap" }}>
                      Lot Number {sortConfig.key === "lotNumber" ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
                    </th>
                    <th onClick={() => handleSort("garmentType")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Garment Type</th>
                    <th onClick={() => handleSort("style")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Style</th>
                    <th onClick={() => handleSort("fabric")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Fabric</th>
                    <th onClick={() => handleSort("brand")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>BRAND</th>
                    <th onClick={() => handleSort("poQty")} style={{ padding: "12px 10px", textAlign: "right", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>
                      Total PCS {sortConfig.key === "poQty" ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
                    </th>
                    <th onClick={() => handleSort("section")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Section</th>
                    <th onClick={() => handleSort("season")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Season</th>
                    <th onClick={() => handleSort("partyName")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>PARTY NAME</th>
                    <th onClick={() => handleSort("directStitching")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Direct Stitching</th>
                    <th onClick={() => handleSort("supervisor")} style={{ padding: "12px 10px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Supervisor</th>
                    <th onClick={() => handleSort("mwk")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>M/W/K</th>
                    <th onClick={() => handleSort("dateOfIssue")} style={{ padding: "12px 10px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Date of Issue</th>
                    <th onClick={() => handleSort("agingDays")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800, background: "#0369a1", color: "#fef08a" }}>
                      Stitching Days {sortConfig.key === "agingDays" ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
                    </th>
                    <th onClick={() => handleSort("embPrintDate")} style={{ padding: "12px 10px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>Emb/Print Date</th>
                    <th style={{ padding: "12px 10px", fontSize: "0.74rem", fontWeight: 800 }}>WIP Status</th>
                    <th style={{ padding: "12px 10px", fontSize: "0.74rem", fontWeight: 800 }}>Pintu</th>
                    <th style={{ padding: "12px 10px", fontSize: "0.74rem", fontWeight: 800 }}>EA</th>
                    <th style={{ padding: "12px 10px", textAlign: "center", fontSize: "0.74rem", fontWeight: 800 }}>Status</th>
                    <th style={{ padding: "12px 10px", textAlign: "center", fontSize: "0.74rem", fontWeight: 800 }}>Action</th>
                  </tr>
                ) : activeStageTab === "packing_handover" ? (
                  /* 7. FLOOR HANDOVER REPORT HEADERS (19 COLUMNS - MATCHING PENDINGPACKINGTOISSUE.JS) */
                  <tr style={{ background: "#7c2d12", color: "#ffffff", userSelect: "none" }}>
                    <th style={{ padding: "12px 8px", textAlign: "center", width: "35px", fontSize: "0.74rem", fontWeight: 800 }}>#</th>
                    <th style={{ padding: "12px 8px", textAlign: "center", fontSize: "0.74rem", fontWeight: 800 }}>IMAGE</th>
                    <th onClick={() => handleSort("lotNumber")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800, whiteSpace: "nowrap" }}>
                      LOT NUMBER {sortConfig.key === "lotNumber" ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
                    </th>
                    <th onClick={() => handleSort("garmentType")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>GARMENT TYPE</th>
                    <th onClick={() => handleSort("style")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>STYLE</th>
                    <th onClick={() => handleSort("fabric")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>FABRIC</th>
                    <th onClick={() => handleSort("brand")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>BRAND</th>
                    <th onClick={() => handleSort("poQty")} style={{ padding: "12px 10px", textAlign: "right", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>
                      STITCHING QTY {sortConfig.key === "poQty" ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
                    </th>
                    <th onClick={() => handleSort("mwk")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>M/W/K</th>
                    <th onClick={() => handleSort("season")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>SEASON</th>
                    <th onClick={() => handleSort("partyName")} style={{ padding: "12px 10px", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>PARTY NAME</th>
                    <th onClick={() => handleSort("directStitching")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>DIRECT STITCHING</th>
                    <th onClick={() => handleSort("supervisor")} style={{ padding: "12px 10px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>SUPERVISOR</th>
                    <th onClick={() => handleSort("dateOfIssue")} style={{ padding: "12px 10px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>DATE OF ISSUE</th>
                    <th onClick={() => handleSort("priority")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>PRIORITY</th>
                    <th onClick={() => handleSort("completedDate")} style={{ padding: "12px 10px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800 }}>COMPLETED DATE</th>
                    <th onClick={() => handleSort("agingDays")} style={{ padding: "12px 8px", textAlign: "center", cursor: "pointer", fontSize: "0.74rem", fontWeight: 800, background: "#9a3412", color: "#fef08a" }}>
                      PENDING DAYS {sortConfig.key === "agingDays" ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
                    </th>
                    <th style={{ padding: "12px 10px", fontSize: "0.74rem", fontWeight: 800 }}>💬 REMARKS</th>
                    <th style={{ padding: "12px 10px", textAlign: "center", fontSize: "0.74rem", fontWeight: 800 }}>ACTION</th>
                  </tr>
                ) : (
                  /* 5. MASTER EXECUTIVE OVERVIEW HEADERS (16 COLUMNS) */
                  <tr style={{ background: "#0f172a", color: "#ffffff", userSelect: "none" }}>
                    <th style={{ padding: "12px 10px", textAlign: "center", width: "40px", fontSize: "0.75rem", fontWeight: 800 }}>#</th>
                    <th onClick={() => handleSort("lotNumber")} style={{ padding: "12px 12px", cursor: "pointer", fontSize: "0.75rem", fontWeight: 800, whiteSpace: "nowrap" }}>
                      Lot Number {sortConfig.key === "lotNumber" ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
                    </th>
                    <th onClick={() => handleSort("stageName")} style={{ padding: "12px 12px", cursor: "pointer", fontSize: "0.75rem", fontWeight: 800, whiteSpace: "nowrap" }}>
                      Critical Stage {sortConfig.key === "stageName" ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
                    </th>
                    <th style={{ padding: "12px 14px", fontSize: "0.75rem", fontWeight: 800, minWidth: "240px" }}>
                      Bottleneck / Variance Reason
                    </th>
                    <th onClick={() => handleSort("severity")} style={{ padding: "12px 10px", textAlign: "center", cursor: "pointer", fontSize: "0.75rem", fontWeight: 800, whiteSpace: "nowrap" }}>
                      Severity {sortConfig.key === "severity" ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
                    </th>
                    <th onClick={() => handleSort("agingDays")} style={{ padding: "12px 10px", textAlign: "center", cursor: "pointer", fontSize: "0.75rem", fontWeight: 800, whiteSpace: "nowrap" }}>
                      Aging {sortConfig.key === "agingDays" ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
                    </th>
                    <th style={{ padding: "12px 12px", fontSize: "0.75rem", fontWeight: 800 }}>Garment Type</th>
                    <th style={{ padding: "12px 12px", fontSize: "0.75rem", fontWeight: 800 }}>Style</th>
                    <th style={{ padding: "12px 12px", fontSize: "0.75rem", fontWeight: 800 }}>Fabric</th>
                    <th style={{ padding: "12px 12px", fontSize: "0.75rem", fontWeight: 800 }}>Brand</th>
                    <th style={{ padding: "12px 12px", fontSize: "0.75rem", fontWeight: 800 }}>Party Name</th>
                    <th style={{ padding: "12px 10px", textAlign: "right", fontSize: "0.75rem", fontWeight: 800 }}>PO / Cut</th>
                    <th style={{ padding: "12px 10px", textAlign: "right", fontSize: "0.75rem", fontWeight: 800 }}>Done</th>
                    <th onClick={() => handleSort("pendingQty")} style={{ padding: "12px 10px", textAlign: "right", cursor: "pointer", fontSize: "0.75rem", fontWeight: 800 }}>
                      Pending {sortConfig.key === "pendingQty" ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
                    </th>
                    <th onClick={() => handleSort("variancePcs")} style={{ padding: "12px 10px", textAlign: "right", cursor: "pointer", fontSize: "0.75rem", fontWeight: 800 }}>
                      Variance {sortConfig.key === "variancePcs" ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
                    </th>
                    <th style={{ padding: "12px 12px", textAlign: "center", fontSize: "0.75rem", fontWeight: 800, minWidth: "120px" }}>Action</th>
                  </tr>
                )}
              </thead>

              {/* ========================================================= */}
              {/* TABLE BODY - DYNAMIC ACCORDING TO ACTIVE DEPARTMENT      */}
              {/* ========================================================= */}
              <tbody>
                {paginatedRows.length === 0 ? (
                  <tr>
                    <td colSpan={20} style={{ padding: "40px", textAlign: "center", color: "#64748b" }}>
                      <div style={{ fontSize: "2rem", marginBottom: "8px" }}>🎉</div>
                      <div style={{ fontSize: "1rem", fontWeight: 800, color: "#1e293b" }}>
                        No Red Zone Bottlenecks Found!
                      </div>
                      <div style={{ fontSize: "0.82rem", color: "#94a3b8", marginTop: "4px" }}>
                        All lots in this view are proceeding within standard timeline parameters.
                      </div>
                    </td>
                  </tr>
                ) : (
                  paginatedRows.map((r, idx) => {
                    const rowNumber = (currentPage - 1) * pageSize + idx + 1;
                    const sev = SEVERITY_CONFIG[r.severity] || SEVERITY_CONFIG.MODERATE;
                    const isEven = idx % 2 === 0;

                    // 1. CUTTING REPORT ROW (18 COLUMNS)
                    if (activeStageTab === "cutting") {
                      return (
                        <tr
                          key={r.id}
                          style={{
                            background: isEven ? "#ffffff" : "#f0fdf4",
                            borderBottom: "1px solid #e2e8f0",
                            transition: "background 0.15s ease"
                          }}
                        >
                          <td style={{ padding: "10px 8px", textAlign: "center", fontWeight: 700, color: "#64748b" }}>{rowNumber}</td>
                          <td style={{ padding: "10px 10px", fontWeight: 900, color: "#1e293b", whiteSpace: "nowrap" }}>
                            <span style={{ background: "#d1fae5", color: "#065f46", padding: "3px 8px", borderRadius: "6px", fontFamily: "monospace", fontSize: "0.85rem", fontWeight: 800 }}>
                              {r.lotNumber}
                            </span>
                          </td>
                          <td style={{ padding: "10px 10px", fontWeight: 600, color: "#334155" }}>{r.garmentType}</td>
                          <td style={{ padding: "10px 10px", fontWeight: 600, color: "#334155" }}>{r.style}</td>
                          <td style={{ padding: "10px 10px", color: "#475569", fontSize: "0.78rem" }}>{r.fabric}</td>
                          <td style={{ padding: "10px 10px", fontWeight: 700, color: "#1e293b" }}>{r.brand}</td>
                          <td style={{ padding: "10px 10px", textAlign: "right", fontWeight: 800, color: "#059669" }}>
                            {(r.poQty || 0).toLocaleString()}
                          </td>
                          <td style={{ padding: "10px 8px", textAlign: "center", color: "#475569" }}>{r.section}</td>
                          <td style={{ padding: "10px 8px", textAlign: "center", color: "#475569" }}>{r.season}</td>
                          <td style={{ padding: "10px 10px", color: "#334155", fontSize: "0.78rem" }}>{r.partyName}</td>
                          <td style={{ padding: "10px 8px", textAlign: "center", fontWeight: 700, color: r.directStitching === "Yes" ? "#059669" : "#64748b" }}>
                            {r.directStitching || "No"}
                          </td>
                          <td style={{ padding: "10px 10px", color: "#475569", fontFamily: "monospace", fontSize: "0.78rem" }}>{r.jobOrderNo || "—"}</td>
                          <td style={{ padding: "10px 10px", textAlign: "center", color: "#475569", fontSize: "0.78rem" }}>{r.stageDate || "—"}</td>
                          <td style={{ padding: "10px 8px", textAlign: "center", background: "#fee2e2", color: "#dc2626", fontWeight: 900, fontSize: "0.85rem", borderLeft: "2px solid #f87171", borderRight: "2px solid #f87171" }}>
                            {r.agingDays}d
                          </td>
                          <td style={{ padding: "10px 10px", color: "#475569", fontSize: "0.78rem" }}>{r.pendingShade || "—"}</td>
                          <td style={{ padding: "10px 10px", textAlign: "center", color: "#64748b", fontSize: "0.78rem" }}>{r.cuttingDate || "—"}</td>
                          <td style={{ padding: "10px 10px", textAlign: "center", color: "#64748b", fontSize: "0.78rem" }}>{r.cuttingScanned || "—"}</td>
                          <td style={{ padding: "10px 10px", fontSize: "0.75rem", maxWidth: "200px" }}>
                            {formatLatestRemark(r.userRemarks, "—") !== "—" ? (
                              <span style={{ display: "inline-block", background: "#f0fdf4", color: "#166534", padding: "4px 8px", borderRadius: "6px", border: "1px solid #bbf7d0", fontWeight: 700 }}>
                                💬 {formatLatestRemark(r.userRemarks)}
                              </span>
                            ) : (
                              <span style={{ color: "#94a3b8" }}>—</span>
                            )}
                          </td>
                          <td style={{ padding: "10px 10px", textAlign: "center" }}>
                            <Link
                              to={r.stageRoute}
                              style={{ display: "inline-flex", alignItems: "center", gap: "4px", background: "#065f46", color: "#ffffff", padding: "5px 10px", borderRadius: "6px", fontSize: "0.72rem", fontWeight: 800, textDecoration: "none", boxShadow: "0 2px 6px rgba(6,95,70,0.2)", whiteSpace: "nowrap" }}
                            >
                              Cutting ↗
                            </Link>
                          </td>
                        </tr>
                      );
                    }

                    // 2. EMBROIDERY REPORT ROW (19 COLUMNS)
                    if (activeStageTab === "embroidery") {
                      return (
                        <tr
                          key={r.id}
                          style={{
                            background: isEven ? "#ffffff" : "#faf5ff",
                            borderBottom: "1px solid #e2e8f0",
                            transition: "background 0.15s ease"
                          }}
                        >
                          <td style={{ padding: "10px 8px", textAlign: "center", fontWeight: 700, color: "#64748b" }}>{rowNumber}</td>
                          <td style={{ padding: "10px 10px", fontWeight: 900, color: "#1e293b", whiteSpace: "nowrap" }}>
                            <span style={{ background: "#ede9fe", color: "#5b21b6", padding: "3px 8px", borderRadius: "6px", fontFamily: "monospace", fontSize: "0.85rem", fontWeight: 800 }}>
                              {r.lotNumber}
                            </span>
                          </td>
                          <td style={{ padding: "10px 10px", fontWeight: 600, color: "#334155" }}>{r.garmentType}</td>
                          <td style={{ padding: "10px 10px", fontWeight: 600, color: "#334155" }}>{r.style}</td>
                          <td style={{ padding: "10px 10px", color: "#475569", fontSize: "0.78rem" }}>{r.fabric}</td>
                          <td style={{ padding: "10px 10px", fontWeight: 700, color: "#1e293b" }}>{r.brand}</td>
                          <td style={{ padding: "10px 10px", textAlign: "right", fontWeight: 800, color: "#7c3aed" }}>
                            {(r.poQty || 0).toLocaleString()}
                          </td>
                          <td style={{ padding: "10px 8px", textAlign: "center", color: "#475569" }}>{r.section}</td>
                          <td style={{ padding: "10px 8px", textAlign: "center", color: "#475569" }}>{r.season}</td>
                          <td style={{ padding: "10px 10px", color: "#334155", fontSize: "0.78rem" }}>{r.partyName}</td>
                          <td style={{ padding: "10px 8px", textAlign: "center", fontWeight: 700, color: r.directStitching === "Yes" ? "#059669" : "#64748b" }}>
                            {r.directStitching || "No"}
                          </td>
                          <td style={{ padding: "10px 10px", fontWeight: 700, color: "#6d28d9", fontSize: "0.78rem" }}>{r.embParty || r.partyName || "—"}</td>
                          <td style={{ padding: "10px 10px", textAlign: "center", color: "#475569", fontSize: "0.78rem" }}>{r.stageDate || "—"}</td>
                          <td style={{ padding: "10px 10px", textAlign: "center" }}>
                            <span style={{ background: "#fef3c7", color: "#92400e", padding: "2px 8px", borderRadius: "10px", fontSize: "0.72rem", fontWeight: 800 }}>
                              {r.embStatus || "Pending"}
                            </span>
                          </td>
                          <td style={{ padding: "10px 10px", color: "#475569", fontSize: "0.78rem" }}>{r.pendingShade || "—"}</td>
                          <td style={{ padding: "10px 10px", textAlign: "center", color: "#64748b", fontSize: "0.78rem" }}>{r.embDoneDate || "—"}</td>
                          <td style={{ padding: "10px 10px", textAlign: "center", color: "#64748b", fontSize: "0.78rem" }}>{r.receivedDate || "—"}</td>
                          <td style={{ padding: "10px 8px", textAlign: "center", background: "#fee2e2", color: "#dc2626", fontWeight: 900, fontSize: "0.85rem", borderLeft: "2px solid #f87171", borderRight: "2px solid #f87171" }}>
                            {r.agingDays}d
                          </td>
                          <td style={{ padding: "10px 10px", fontSize: "0.75rem", maxWidth: "200px" }}>
                            {formatLatestRemark(r.hodRemarks, "—") !== "—" ? (
                              <span style={{ display: "inline-block", background: "#f5f3ff", color: "#5b21b6", padding: "4px 8px", borderRadius: "6px", border: "1px solid #ddd6fe", fontWeight: 700 }}>
                                💬 {formatLatestRemark(r.hodRemarks)}
                              </span>
                            ) : (
                              <span style={{ color: "#94a3b8" }}>—</span>
                            )}
                          </td>
                          <td style={{ padding: "10px 10px", textAlign: "center" }}>
                            <Link
                              to={r.stageRoute}
                              style={{ display: "inline-flex", alignItems: "center", gap: "4px", background: "#6d28d9", color: "#ffffff", padding: "5px 10px", borderRadius: "6px", fontSize: "0.72rem", fontWeight: 800, textDecoration: "none", boxShadow: "0 2px 6px rgba(109,40,217,0.2)", whiteSpace: "nowrap" }}
                            >
                              Embroidery ↗
                            </Link>
                          </td>
                        </tr>
                      );
                    }

                    // 3. PRINTING REPORT ROW (18 COLUMNS)
                    if (activeStageTab === "printing") {
                      return (
                        <tr
                          key={r.id}
                          style={{
                            background: isEven ? "#ffffff" : "#fdf2f8",
                            borderBottom: "1px solid #e2e8f0",
                            transition: "background 0.15s ease"
                          }}
                        >
                          <td style={{ padding: "10px 8px", textAlign: "center", fontWeight: 700, color: "#64748b" }}>{rowNumber}</td>
                          <td style={{ padding: "10px 10px", fontWeight: 900, color: "#1e293b", whiteSpace: "nowrap" }}>
                            <span style={{ background: "#fce7f3", color: "#9d174d", padding: "3px 8px", borderRadius: "6px", fontFamily: "monospace", fontSize: "0.85rem", fontWeight: 800 }}>
                              {r.lotNumber}
                            </span>
                          </td>
                          <td style={{ padding: "10px 10px", fontWeight: 600, color: "#334155" }}>{r.garmentType}</td>
                          <td style={{ padding: "10px 10px", fontWeight: 600, color: "#334155" }}>{r.style}</td>
                          <td style={{ padding: "10px 10px", color: "#475569", fontSize: "0.78rem" }}>{r.fabric}</td>
                          <td style={{ padding: "10px 10px", fontWeight: 700, color: "#1e293b" }}>{r.brand}</td>
                          <td style={{ padding: "10px 10px", textAlign: "right", fontWeight: 800, color: "#db2777" }}>
                            {(r.poQty || 0).toLocaleString()}
                          </td>
                          <td style={{ padding: "10px 8px", textAlign: "center", color: "#475569" }}>{r.section}</td>
                          <td style={{ padding: "10px 8px", textAlign: "center", color: "#475569" }}>{r.season}</td>
                          <td style={{ padding: "10px 10px", color: "#334155", fontSize: "0.78rem" }}>{r.partyName}</td>
                          <td style={{ padding: "10px 8px", textAlign: "center", fontWeight: 700, color: r.directStitching === "Yes" ? "#059669" : "#64748b" }}>
                            {r.directStitching || "No"}
                          </td>
                          <td style={{ padding: "10px 10px", fontWeight: 700, color: "#be185d", fontSize: "0.78rem" }}>{r.printingParty || r.partyName || "—"}</td>
                          <td style={{ padding: "10px 10px", textAlign: "center", color: "#475569", fontSize: "0.78rem" }}>{r.stageDate || "—"}</td>
                          <td style={{ padding: "10px 10px", textAlign: "center" }}>
                            <span style={{ background: "#fef3c7", color: "#92400e", padding: "2px 8px", borderRadius: "10px", fontSize: "0.72rem", fontWeight: 800 }}>
                              {r.printingStatus || "Pending"}
                            </span>
                          </td>
                          <td style={{ padding: "10px 10px", color: "#475569", fontSize: "0.78rem" }}>{r.pendingShade || "—"}</td>
                          <td style={{ padding: "10px 10px", textAlign: "center", color: "#64748b", fontSize: "0.78rem" }}>{r.printingDoneDate || "—"}</td>
                          <td style={{ padding: "10px 8px", textAlign: "center", background: "#fee2e2", color: "#dc2626", fontWeight: 900, fontSize: "0.85rem", borderLeft: "2px solid #f87171", borderRight: "2px solid #f87171" }}>
                            {r.agingDays}d
                          </td>
                          <td style={{ padding: "10px 10px", fontSize: "0.75rem", maxWidth: "200px" }}>
                            {formatLatestRemark(r.hodRemarks, "—") !== "—" ? (
                              <span style={{ display: "inline-block", background: "#fdf2f8", color: "#9d174d", padding: "4px 8px", borderRadius: "6px", border: "1px solid #fbcfe8", fontWeight: 700 }}>
                                💬 {formatLatestRemark(r.hodRemarks)}
                              </span>
                            ) : (
                              <span style={{ color: "#94a3b8" }}>—</span>
                            )}
                          </td>
                          <td style={{ padding: "10px 10px", textAlign: "center" }}>
                            <Link
                              to={r.stageRoute}
                              style={{ display: "inline-flex", alignItems: "center", gap: "4px", background: "#be185d", color: "#ffffff", padding: "5px 10px", borderRadius: "6px", fontSize: "0.72rem", fontWeight: 800, textDecoration: "none", boxShadow: "0 2px 6px rgba(190,24,93,0.2)", whiteSpace: "nowrap" }}
                            >
                              Printing ↗
                            </Link>
                          </td>
                        </tr>
                      );
                    }

                    // 4. POST-EMB/PRINT REPORT ROW (18 COLUMNS - MATCHING PENDINGISSUE.JS)
                    if (activeStageTab === "post_emb_print") {
                      return (
                        <tr
                          key={r.id}
                          style={{
                            background: isEven ? "#ffffff" : "#fff7ed",
                            borderBottom: "1px solid #e2e8f0",
                            transition: "background 0.15s ease"
                          }}
                        >
                          <td style={{ padding: "10px 8px", textAlign: "center", fontWeight: 700, color: "#64748b" }}>
                            {r.isRepeatedLot ? "★ " : ""}{rowNumber}
                          </td>
                          <td style={{ padding: "10px 10px", fontWeight: 900, color: "#1e293b", whiteSpace: "nowrap" }}>
                            <span style={{ background: r.isRepeatedLot ? "#fef3c7" : "#ffedd5", color: r.isRepeatedLot ? "#92400e" : "#9a3412", padding: "3px 8px", borderRadius: "6px", fontFamily: "monospace", fontSize: "0.85rem", fontWeight: 800, border: r.isRepeatedLot ? "1px solid #fde68a" : "none" }}>
                              {r.lotNumber} {r.isRepeatedLot ? "★" : ""}
                            </span>
                          </td>
                          <td style={{ padding: "10px 10px", fontWeight: 600, color: "#334155" }}>{r.garmentType}</td>
                          <td style={{ padding: "10px 10px", fontWeight: 600, color: "#334155" }}>{r.style}</td>
                          <td style={{ padding: "10px 10px", color: "#475569", fontSize: "0.78rem" }}>{r.fabric}</td>
                          <td style={{ padding: "10px 10px", fontWeight: 700, color: "#1e293b" }}>{r.brand}</td>
                          <td style={{ padding: "10px 10px", textAlign: "right", fontWeight: 800, color: "#ea580c" }}>
                            {(r.poQty || 0).toLocaleString()}
                          </td>
                          <td style={{ padding: "10px 8px", textAlign: "center", color: "#475569" }}>{r.section}</td>
                          <td style={{ padding: "10px 8px", textAlign: "center", color: "#475569" }}>{r.season}</td>
                          <td style={{ padding: "10px 10px", color: "#334155", fontSize: "0.78rem" }}>{r.partyName}</td>
                          <td style={{ padding: "10px 8px", textAlign: "center", fontWeight: 700, color: (r.directStitching || "").toLowerCase() === "yes" ? "#059669" : "#64748b" }}>
                            {r.directStitching || "no"}
                          </td>
                          <td style={{ padding: "10px 10px", textAlign: "center", color: "#64748b", fontSize: "0.78rem" }}>{r.cuttingDate || "—"}</td>
                          <td style={{ padding: "10px 10px", textAlign: "center", color: "#64748b", fontSize: "0.78rem" }}>{r.embPrintDate || "—"}</td>
                          <td style={{ padding: "10px 8px", textAlign: "center", background: "#fee2e2", color: "#dc2626", fontWeight: 900, fontSize: "0.85rem", borderLeft: "2px solid #f87171", borderRight: "2px solid #f87171" }}>
                            {r.agingDays}d
                          </td>
                          <td style={{ padding: "10px 10px", color: "#475569", fontSize: "0.78rem" }}>{r.pendingShade || "—"}</td>
                          <td style={{ padding: "10px 8px", textAlign: "center", color: "#475569", fontSize: "0.78rem" }}>
                            <span style={{ background: r.isRepeatedLot ? "#fef3c7" : "#f1f5f9", color: r.isRepeatedLot ? "#b45309" : "#475569", padding: "2px 8px", borderRadius: "8px", fontWeight: 700 }}>
                              {r.priority || "Normal"}
                            </span>
                          </td>
                          <td style={{ padding: "10px 10px", fontSize: "0.75rem", maxWidth: "200px" }}>
                            {formatLatestRemark(r.userRemarks, "—") !== "—" ? (
                              <span style={{ display: "inline-block", background: "#fff7ed", color: "#9a3412", padding: "4px 8px", borderRadius: "6px", border: "1px solid #fed7aa", fontWeight: 700 }}>
                                💬 {formatLatestRemark(r.userRemarks)}
                              </span>
                            ) : (
                              <span style={{ color: "#94a3b8" }}>—</span>
                            )}
                          </td>
                          <td style={{ padding: "10px 10px", textAlign: "center" }}>
                            <Link
                              to={r.stageRoute}
                              style={{ display: "inline-flex", alignItems: "center", gap: "4px", background: "#c2410c", color: "#ffffff", padding: "5px 10px", borderRadius: "6px", fontSize: "0.72rem", fontWeight: 800, textDecoration: "none", boxShadow: "0 2px 6px rgba(194,65,12,0.2)", whiteSpace: "nowrap" }}
                            >
                              Post-EMB ↗
                            </Link>
                          </td>
                        </tr>
                      );
                    }

                    // 6. STITCHING WIP REPORT ROW (21 COLUMNS - MATCHING STITCHINGCOMPLETED.JS)
                    if (activeStageTab === "stitching") {
                      return (
                        <tr
                          key={r.id}
                          style={{
                            background: isEven ? "#ffffff" : "#f0f9ff",
                            borderBottom: "1px solid #e2e8f0",
                            transition: "background 0.15s ease"
                          }}
                        >
                          <td style={{ padding: "10px 8px", textAlign: "center", fontWeight: 700, color: "#64748b" }}>
                            {rowNumber}
                          </td>
                          <td style={{ padding: "10px 10px", fontWeight: 900, color: "#1e293b", whiteSpace: "nowrap" }}>
                            <span style={{ background: "#e0f2fe", color: "#0369a1", padding: "3px 8px", borderRadius: "6px", fontFamily: "monospace", fontSize: "0.85rem", fontWeight: 800 }}>
                              {r.lotNumber}
                            </span>
                          </td>
                          <td style={{ padding: "10px 10px", fontWeight: 600, color: "#334155" }}>{r.garmentType}</td>
                          <td style={{ padding: "10px 10px", fontWeight: 600, color: "#334155" }}>{r.style}</td>
                          <td style={{ padding: "10px 10px", color: "#475569", fontSize: "0.78rem" }}>{r.fabric}</td>
                          <td style={{ padding: "10px 10px", fontWeight: 700, color: "#1e293b" }}>{r.brand}</td>
                          <td style={{ padding: "10px 10px", textAlign: "right", fontWeight: 800, color: "#0284c7" }}>
                            {(r.poQty || 0).toLocaleString()}
                          </td>
                          <td style={{ padding: "10px 8px", textAlign: "center", fontWeight: 600, color: "#475569" }}>{r.section}</td>
                          <td style={{ padding: "10px 8px", textAlign: "center", color: "#475569" }}>{r.season}</td>
                          <td style={{ padding: "10px 10px", color: "#334155", fontSize: "0.78rem" }}>{r.partyName}</td>
                          <td style={{ padding: "10px 8px", textAlign: "center", fontWeight: 700, color: (r.directStitching || "").toLowerCase() === "yes" ? "#059669" : "#64748b" }}>
                            {r.directStitching || "no"}
                          </td>
                          <td style={{ padding: "10px 10px", textAlign: "center", fontWeight: 800, color: "#2563eb" }}>
                            {r.supervisor || "Unassigned"}
                          </td>
                          <td style={{ padding: "10px 8px", textAlign: "center", fontWeight: 800, color: r.mwk === "M" || r.mwk === "GENTS" ? "#3b82f6" : r.mwk === "W" || r.mwk === "WOMEN" ? "#ef4444" : "#10b981" }}>
                            {r.mwk || "—"}
                          </td>
                          <td style={{ padding: "10px 10px", textAlign: "center", color: "#64748b", fontSize: "0.78rem" }}>{r.dateOfIssue || "—"}</td>
                          <td style={{ padding: "10px 8px", textAlign: "center", background: "#fee2e2", color: "#dc2626", fontWeight: 900, fontSize: "0.85rem", borderLeft: "2px solid #f87171", borderRight: "2px solid #f87171" }}>
                            {r.agingDays}d
                          </td>
                          <td style={{ padding: "10px 10px", textAlign: "center", color: "#64748b", fontSize: "0.78rem" }}>{r.embPrintDate || "—"}</td>
                          <td style={{ padding: "10px 10px", fontWeight: 600, color: "#475569", fontSize: "0.78rem" }}>{formatLatestRemark(r.wipStatus, "WIP")}</td>
                          <td style={{ padding: "10px 10px", fontSize: "0.75rem", maxWidth: "160px" }}>
                            {formatLatestRemark(r.pintu, "—") !== "—" ? (
                              <span style={{ display: "inline-block", background: "#f3e8ff", color: "#7c3aed", padding: "3px 7px", borderRadius: "6px", fontWeight: 700 }}>
                                {formatLatestRemark(r.pintu)}
                              </span>
                            ) : (
                              <span style={{ color: "#94a3b8" }}>—</span>
                            )}
                          </td>
                          <td style={{ padding: "10px 10px", fontSize: "0.75rem", maxWidth: "160px" }}>
                            {formatLatestRemark(r.ea, "—") !== "—" ? (
                              <span style={{ display: "inline-block", background: "#dcfce7", color: "#15803d", padding: "3px 7px", borderRadius: "6px", fontWeight: 700 }}>
                                {formatLatestRemark(r.ea)}
                              </span>
                            ) : (
                              <span style={{ color: "#94a3b8" }}>—</span>
                            )}
                          </td>
                          <td style={{ padding: "10px 8px", textAlign: "center" }}>
                            <span style={{ background: "#fee2e2", color: "#b91c1c", padding: "3px 8px", borderRadius: "12px", fontSize: "0.72rem", fontWeight: 800 }}>
                              {r.status || "Pending"}
                            </span>
                          </td>
                          <td style={{ padding: "10px 10px", textAlign: "center" }}>
                            <Link
                              to={r.stageRoute}
                              style={{ display: "inline-flex", alignItems: "center", gap: "4px", background: "#0284c7", color: "#ffffff", padding: "5px 10px", borderRadius: "6px", fontSize: "0.72rem", fontWeight: 800, textDecoration: "none", boxShadow: "0 2px 6px rgba(2,132,199,0.2)", whiteSpace: "nowrap" }}
                            >
                              Stitching ↗
                            </Link>
                          </td>
                        </tr>
                      );
                    }

                    // 7. FLOOR HANDOVER REPORT ROW (19 COLUMNS - MATCHING PENDINGPACKINGTOISSUE.JS)
                    if (activeStageTab === "packing_handover") {
                      const directImgUrl = getDirectImageUrl(r.imageUrl);
                      return (
                        <tr
                          key={r.id}
                          style={{
                            background: isEven ? "#ffffff" : "#fff7ed",
                            borderBottom: "1px solid #e2e8f0",
                            transition: "background 0.15s ease"
                          }}
                        >
                          <td style={{ padding: "10px 8px", textAlign: "center", fontWeight: 700, color: "#64748b" }}>
                            {rowNumber}
                          </td>
                          <td style={{ padding: "6px 8px", textAlign: "center" }}>
                            {directImgUrl ? (
                              <img
                                src={directImgUrl}
                                alt={`Lot #${r.lotNumber}`}
                                onClick={() => setViewImageSrc(directImgUrl)}
                                style={{
                                  width: "42px",
                                  height: "42px",
                                  objectFit: "cover",
                                  borderRadius: "8px",
                                  border: "1px solid #cbd5e1",
                                  cursor: "pointer",
                                  boxShadow: "0 2px 6px rgba(0,0,0,0.1)"
                                }}
                                onError={(e) => { e.target.style.display = "none"; }}
                              />
                            ) : (
                              <span style={{ color: "#94a3b8", fontSize: "0.75rem" }}>—</span>
                            )}
                          </td>
                          <td style={{ padding: "10px 10px", fontWeight: 900, color: "#1e293b", whiteSpace: "nowrap" }}>
                            <span style={{ background: "#ffedd5", color: "#9a3412", padding: "3px 8px", borderRadius: "6px", fontFamily: "monospace", fontSize: "0.85rem", fontWeight: 800 }}>
                              {r.lotNumber}
                            </span>
                          </td>
                          <td style={{ padding: "10px 10px", fontWeight: 600, color: "#334155" }}>{r.garmentType}</td>
                          <td style={{ padding: "10px 10px", fontWeight: 600, color: "#334155" }}>{r.style}</td>
                          <td style={{ padding: "10px 10px", color: "#475569", fontSize: "0.78rem" }}>{r.fabric}</td>
                          <td style={{ padding: "10px 10px", fontWeight: 700, color: "#1e293b" }}>{r.brand}</td>
                          <td style={{ padding: "10px 10px", textAlign: "right", fontWeight: 800, color: "#c2410c" }}>
                            {(r.poQty || 0).toLocaleString()}
                          </td>
                          <td style={{ padding: "10px 8px", textAlign: "center", fontWeight: 700, color: r.mwk === "M" || r.mwk === "GENTS" ? "#3b82f6" : r.mwk === "W" || r.mwk === "WOMEN" ? "#ef4444" : "#10b981" }}>
                            {r.mwk || "—"}
                          </td>
                          <td style={{ padding: "10px 8px", textAlign: "center", color: "#475569" }}>{r.season}</td>
                          <td style={{ padding: "10px 10px", color: "#334155", fontSize: "0.78rem" }}>{r.partyName}</td>
                          <td style={{ padding: "10px 8px", textAlign: "center" }}>
                            <span style={{
                              background: (r.directStitching || "").toLowerCase() === "yes" ? "#dcfce7" : "#f1f5f9",
                              color: (r.directStitching || "").toLowerCase() === "yes" ? "#15803d" : "#64748b",
                              padding: "2px 8px",
                              borderRadius: "8px",
                              fontSize: "0.72rem",
                              fontWeight: 700
                            }}>
                              {r.directStitching || "no"}
                            </span>
                          </td>
                          <td style={{ padding: "10px 10px", textAlign: "center", fontWeight: 800, color: "#2563eb" }}>
                            {r.supervisor || "—"}
                          </td>
                          <td style={{ padding: "10px 10px", textAlign: "center", color: "#64748b", fontSize: "0.78rem" }}>
                            {r.dateOfIssue || "—"}
                          </td>
                          <td style={{ padding: "10px 8px", textAlign: "center" }}>
                            <span style={{
                              background: (r.priority || "Normal").toLowerCase() === "high" || (r.priority || "").toLowerCase() === "urgent" ? "#fee2e2" : "#f1f5f9",
                              color: (r.priority || "Normal").toLowerCase() === "high" || (r.priority || "").toLowerCase() === "urgent" ? "#b91c1c" : "#475569",
                              padding: "2px 8px",
                              borderRadius: "8px",
                              fontSize: "0.72rem",
                              fontWeight: 700
                            }}>
                              {r.priority || "Normal"}
                            </span>
                          </td>
                          <td style={{ padding: "10px 10px", textAlign: "center" }}>
                            <span style={{ background: "#dcfce7", color: "#166534", padding: "2px 8px", borderRadius: "6px", fontSize: "0.72rem", fontWeight: 700 }}>
                              {r.completedDate || "—"}
                            </span>
                          </td>
                          <td style={{ padding: "10px 8px", textAlign: "center", background: "#fee2e2", color: "#dc2626", fontWeight: 900, fontSize: "0.85rem", borderLeft: "2px solid #f87171", borderRight: "2px solid #f87171" }}>
                            {r.agingDays} {r.agingDays === 1 ? "day" : "days"}
                          </td>
                          <td style={{ padding: "10px 10px", fontSize: "0.75rem", maxWidth: "200px" }}>
                            {formatLatestRemark(r.userRemarks, "—") !== "—" ? (
                              <span style={{ display: "inline-block", background: "#fff7ed", color: "#9a3412", padding: "4px 8px", borderRadius: "6px", border: "1px solid #fed7aa", fontWeight: 700 }}>
                                💬 {formatLatestRemark(r.userRemarks)}
                              </span>
                            ) : (
                              <span style={{ color: "#94a3b8" }}>—</span>
                            )}
                          </td>
                          <td style={{ padding: "10px 10px", textAlign: "center" }}>
                            <Link
                              to={r.stageRoute}
                              style={{ display: "inline-flex", alignItems: "center", gap: "4px", background: "#c2410c", color: "#ffffff", padding: "5px 10px", borderRadius: "6px", fontSize: "0.72rem", fontWeight: 800, textDecoration: "none", boxShadow: "0 2px 6px rgba(194,65,12,0.2)", whiteSpace: "nowrap" }}
                            >
                              Packing ↗
                            </Link>
                          </td>
                        </tr>
                      );
                    }

                    // 5. MASTER EXECUTIVE OVERVIEW ROW (16 COLUMNS)
                    return (
                      <tr
                        key={r.id}
                        style={{
                          background: r.severity === "CRITICAL" ? "#fff1f2" : isEven ? "#ffffff" : "#f8fafc",
                          borderBottom: "1px solid #e2e8f0",
                          transition: "background 0.15s ease"
                        }}
                      >
                        {/* Sr. No */}
                        <td style={{ padding: "10px 8px", textAlign: "center", fontWeight: 700, color: "#64748b" }}>
                          {rowNumber}
                        </td>

                        {/* Lot Number */}
                        <td style={{ padding: "10px 12px", fontWeight: 900, color: "#1e293b", whiteSpace: "nowrap" }}>
                          <span
                            style={{
                              background: "#e2e8f0",
                              padding: "3px 8px",
                              borderRadius: "6px",
                              fontFamily: "monospace",
                              fontSize: "0.85rem",
                              color: "#0f172a"
                            }}
                          >
                            {r.lotNumber}
                          </span>
                        </td>

                        {/* Stage Name */}
                        <td style={{ padding: "10px 12px", whiteSpace: "nowrap" }}>
                          <span
                            style={{
                              display: "inline-flex",
                              alignItems: "center",
                              gap: "4px",
                              padding: "3px 8px",
                              borderRadius: "6px",
                              fontSize: "0.74rem",
                              fontWeight: 800,
                              background: "#f1f5f9",
                              color: "#1e293b",
                              border: "1px solid #cbd5e1"
                            }}
                          >
                            {r.stageName}
                          </span>
                        </td>

                        {/* Bottleneck / Variance Reason */}
                        <td style={{ padding: "10px 14px" }}>
                          <div style={{ fontWeight: 800, color: r.severity === "CRITICAL" ? "#b91c1c" : "#1e293b", fontSize: "0.82rem" }}>
                            {r.reason}
                          </div>
                          <div style={{ fontSize: "0.72rem", color: "#64748b", marginTop: "2px" }}>
                            {r.details}
                          </div>
                        </td>

                        {/* Severity */}
                        <td style={{ padding: "10px 10px", textAlign: "center", whiteSpace: "nowrap" }}>
                          <span
                            style={{
                              padding: "3px 8px",
                              borderRadius: "12px",
                              fontSize: "0.7rem",
                              fontWeight: 900,
                              background: sev.bg,
                              color: sev.color,
                              border: `1px solid ${sev.border}`
                            }}
                          >
                            {sev.badge}
                          </span>
                        </td>

                        {/* Aging */}
                        <td style={{ padding: "10px 10px", textAlign: "center", fontWeight: 900, color: r.agingDays >= 5 ? "#dc2626" : r.agingDays >= 3 ? "#ea580c" : "#d97706" }}>
                          {r.agingDays}d
                        </td>

                        {/* Garment Type */}
                        <td style={{ padding: "10px 12px", fontWeight: 600, color: "#334155" }}>
                          {r.garmentType}
                        </td>

                        {/* Style */}
                        <td style={{ padding: "10px 12px", fontWeight: 600, color: "#334155" }}>
                          {r.style}
                        </td>

                        {/* Fabric */}
                        <td style={{ padding: "10px 12px", color: "#475569", fontSize: "0.78rem" }}>
                          {r.fabric}
                        </td>

                        {/* Brand */}
                        <td style={{ padding: "10px 12px", fontWeight: 700, color: "#1e293b" }}>
                          {r.brand}
                        </td>

                        {/* Party Name */}
                        <td style={{ padding: "10px 12px", color: "#475569", fontSize: "0.78rem" }}>
                          {r.partyName}
                        </td>

                        {/* PO / Cut Qty */}
                        <td style={{ padding: "10px 10px", textAlign: "right", fontWeight: 700, color: "#334155" }}>
                          {(r.poQty || 0).toLocaleString()}
                        </td>

                        {/* Done Qty */}
                        <td style={{ padding: "10px 10px", textAlign: "right", fontWeight: 700, color: "#059669" }}>
                          {(r.stageDoneQty || 0).toLocaleString()}
                        </td>

                        {/* Pending Qty */}
                        <td style={{ padding: "10px 10px", textAlign: "right", fontWeight: 900, color: "#dc2626" }}>
                          {(r.pendingQty || 0).toLocaleString()}
                        </td>

                        {/* Variance */}
                        <td style={{ padding: "10px 10px", textAlign: "right", fontWeight: 800, color: r.variancePcs < 0 ? "#dc2626" : "#475569" }}>
                          {r.variancePcs > 0 ? `+${r.variancePcs}` : r.variancePcs}
                          <div style={{ fontSize: "0.7rem", color: "#64748b" }}>{r.variancePct}%</div>
                        </td>

                        {/* Action: Jump to Specific Stage Report */}
                        <td style={{ padding: "10px 12px", textAlign: "center" }}>
                          <Link
                            to={r.stageRoute}
                            style={{
                              display: "inline-flex",
                              alignItems: "center",
                              gap: "4px",
                              background: "#0f172a",
                              color: "#ffffff",
                              padding: "5px 10px",
                              borderRadius: "6px",
                              fontSize: "0.72rem",
                              fontWeight: 800,
                              textDecoration: "none",
                              boxShadow: "0 2px 6px rgba(0,0,0,0.12)",
                              whiteSpace: "nowrap"
                            }}
                          >
                            Open Report ↗
                          </Link>
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>

              {/* Table Subtotals / Footer summary */}
              {filteredRows.length > 0 && (
                <tfoot>
                  <tr style={{ background: "#f1f5f9", borderTop: "2px solid #cbd5e1", fontWeight: 800, color: "#0f172a" }}>
                    <td colSpan={activeStageTab === "cutting" || activeStageTab === "printing" ? 6 : activeStageTab === "embroidery" ? 6 : activeStageTab === "packing_handover" ? 7 : 11} style={{ padding: "12px 14px", textAlign: "right", fontSize: "0.82rem" }}>
                      TOTAL ({filteredRows.length} Red Zone Lots):
                    </td>
                    <td style={{ padding: "12px 10px", textAlign: "right", fontSize: "0.88rem", color: "#0f172a", fontWeight: 900 }}>
                      {filteredRows.reduce((sum, r) => sum + (r.poQty || 0), 0).toLocaleString()} Pcs
                    </td>
                    <td colSpan={activeStageTab === "cutting" ? 13 : activeStageTab === "embroidery" ? 14 : activeStageTab === "printing" ? 13 : activeStageTab === "packing_handover" ? 11 : 4} style={{ padding: "12px 14px", color: "#64748b", fontSize: "0.78rem" }}>
                      Showing {filteredRows.filter((r) => r.severity === "CRITICAL").length} Critical & {filteredRows.filter((r) => r.severity === "HIGH").length} High Severity Delayed Lots
                    </td>
                  </tr>
                </tfoot>
              )}
            </table>
          </div>

          {/* Pagination Footer */}
          <div style={{ padding: "12px 18px", borderTop: "1px solid #e2e8f0", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "10px", background: "#f8fafc" }}>
            <div style={{ fontSize: "0.8rem", color: "#64748b", fontWeight: 600 }}>
              Showing {filteredRows.length > 0 ? (currentPage - 1) * pageSize + 1 : 0} to {Math.min(currentPage * pageSize, filteredRows.length)} of {filteredRows.length} Bottleneck Lots
            </div>

            <div style={{ display: "flex", gap: "6px", alignItems: "center" }}>
              <button
                type="button"
                onClick={() => setCurrentPage(1)}
                disabled={currentPage === 1}
                style={{ padding: "5px 10px", borderRadius: "6px", border: "1px solid #cbd5e1", background: currentPage === 1 ? "#f1f5f9" : "#ffffff", cursor: currentPage === 1 ? "not-allowed" : "pointer", fontSize: "0.78rem", fontWeight: 700 }}
              >
                « First
              </button>
              <button
                type="button"
                onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                disabled={currentPage === 1}
                style={{ padding: "5px 10px", borderRadius: "6px", border: "1px solid #cbd5e1", background: currentPage === 1 ? "#f1f5f9" : "#ffffff", cursor: currentPage === 1 ? "not-allowed" : "pointer", fontSize: "0.78rem", fontWeight: 700 }}
              >
                ‹ Prev
              </button>

              <span style={{ padding: "5px 10px", fontSize: "0.8rem", fontWeight: 800, color: "#1e293b" }}>
                Page {currentPage} of {totalPages}
              </span>

              <button
                type="button"
                onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                disabled={currentPage === totalPages}
                style={{ padding: "5px 10px", borderRadius: "6px", border: "1px solid #cbd5e1", background: currentPage === totalPages ? "#f1f5f9" : "#ffffff", cursor: currentPage === totalPages ? "not-allowed" : "pointer", fontSize: "0.78rem", fontWeight: 700 }}
              >
                Next ›
              </button>
              <button
                type="button"
                onClick={() => setCurrentPage(totalPages)}
                disabled={currentPage === totalPages}
                style={{ padding: "5px 10px", borderRadius: "6px", border: "1px solid #cbd5e1", background: currentPage === totalPages ? "#f1f5f9" : "#ffffff", cursor: currentPage === totalPages ? "not-allowed" : "pointer", fontSize: "0.78rem", fontWeight: 700 }}
              >
                Last »
              </button>
            </div>
          </div>
        </section>
      </main>

      {/* Image Preview Modal */}
      {viewImageSrc && (
        <div
          onClick={() => setViewImageSrc(null)}
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: "rgba(0, 0, 0, 0.85)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            zIndex: 9999,
            padding: "20px"
          }}
        >
          <div
            style={{
              position: "relative",
              maxWidth: "90vw",
              maxHeight: "90vh",
              background: "#ffffff",
              borderRadius: "16px",
              overflow: "hidden",
              boxShadow: "0 25px 50px -12px rgba(0, 0, 0, 0.5)"
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <button
              type="button"
              onClick={() => setViewImageSrc(null)}
              style={{
                position: "absolute",
                top: "12px",
                right: "12px",
                width: "36px",
                height: "36px",
                borderRadius: "50%",
                background: "rgba(0,0,0,0.6)",
                color: "#ffffff",
                border: "none",
                fontSize: "1.2rem",
                cursor: "pointer",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                zIndex: 10
              }}
            >
              ✕
            </button>
            <img
              src={viewImageSrc}
              alt="Preview"
              style={{
                display: "block",
                maxWidth: "85vw",
                maxHeight: "85vh",
                objectFit: "contain"
              }}
            />
          </div>
        </div>
      )}
    </div>
  );
}

