import React, { useState, useEffect } from "react";
import { useHistory, useLocation } from "react-router-dom";
import { motion, AnimatePresence } from "framer-motion";
import { BACKEND_URL, SPREADSHEET_IDS, SHEET_NAMES, fetchSheetDataFromBackend } from "./config";

function normalizeLot(lot) {
  if (!lot) return "";
  return String(lot).trim().toUpperCase().replace(/^LOT[-#\s]*/i, "");
}

function findCol(headers, keywords) {
  if (!Array.isArray(headers)) return -1;
  const cleanKeywords = keywords.map((k) => k.toLowerCase().trim());
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || "").toLowerCase().trim();
    if (cleanKeywords.includes(h)) return i;
  }
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || "").toLowerCase().trim();
    if (cleanKeywords.some((k) => h.includes(k))) return i;
  }
  return -1;
}

function formatReadableDate(dateVal) {
  if (!dateVal) return "";
  const str = String(dateVal).trim();
  if (!str || str === "-" || str.toLowerCase() === "invalid date" || str.toLowerCase() === "pending") {
    return "";
  }
  try {
    if (str.includes("/")) {
      const parts = str.split(" ")[0].split("/");
      if (parts.length === 3) {
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10) - 1;
        let year = parseInt(parts[2], 10);
        if (year < 100) year = year <= 50 ? 2000 + year : 1900 + year;
        const d = new Date(year, month, day);
        if (!isNaN(d.getTime())) {
          return d.toLocaleDateString("en-US", { day: "numeric", month: "short", year: "numeric" });
        }
      }
    }
    if (str.includes("-")) {
      const parts = str.split(" ")[0].split("T")[0].split("-");
      if (parts.length === 3 && parts[0].length === 4) {
        const year = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10) - 1;
        const day = parseInt(parts[2], 10);
        const d = new Date(year, month, day);
        if (!isNaN(d.getTime())) {
          return d.toLocaleDateString("en-US", { day: "numeric", month: "short", year: "numeric" });
        }
      }
    }
    const d = new Date(str);
    if (!isNaN(d.getTime())) {
      return d.toLocaleDateString("en-US", { day: "numeric", month: "short", year: "numeric" });
    }
  } catch (e) {}
  return str;
}

function calculateDaysDiff(d1Str, d2Str) {
  if (!d1Str || !d2Str || d1Str === "Pending" || d2Str === "Pending" || d2Str === "In Progress") return null;
  try {
    const d1 = new Date(d1Str);
    const d2 = new Date(d2Str);
    if (isNaN(d1.getTime()) || isNaN(d2.getTime())) return null;
    const diff = Math.round((d2 - d1) / (1000 * 60 * 60 * 24));
    return diff >= 0 ? `+${diff} days` : `${diff} days`;
  } catch (e) {
    return null;
  }
}

export default function LotTimelineReport() {
  const history = useHistory();
  const location = useLocation();

  const queryParams = new URLSearchParams(location.search);
  const initialLot = queryParams.get("lot") || "76063";

  const [lotInput, setLotInput] = useState(initialLot);
  const [activeLot, setActiveLot] = useState(initialLot);
  const [loading, setLoading] = useState(false);
  const [timelineData, setTimelineData] = useState(null);
  const [error, setError] = useState(null);
  const [recentLots, setRecentLots] = useState(["76063", "65002", "11841", "12140", "11874"]);
  const [isDarkMode, setIsDarkMode] = useState(false); // Default: Light Color Tracker

  // Client-side fallback that queries Google Sheets directly if backend is offline
  const buildTimelineFromSheets = async (cleanLot) => {
    const mainId = SPREADSHEET_IDS.MAIN || "1Hj3JeJEKB43aYYWv8gk2UhdU6BWuEQfCg5pBlTdBMNA";
    const jobOrderId = SPREADSHEET_IDS.JOBORDER || "1fKSwGBIpzWEFk566WRQ4bzQ0anJlmasoY8TwrTLQHXI";
    const issuesId = SPREADSHEET_IDS.ISSUES || "1uo14nKO_yHu4AJ2rOgaJajuprcinj6xw1AUMFJ6_zYM";
    const workingUpdatesId = SPREADSHEET_IDS.WORKING_UPDATES || "1Nh7XYE_MnAxtaTRUUntHvBpzctODwjnkbBYDiYLQgoc";
    const barcodeId = SPREADSHEET_IDS.BARCODE || "1dOCjNFwaAel5qun0_ZJVIGmREqjI76CJBBFIjM3NHv8";
    const rawpackId = SPREADSHEET_IDS.RAWPACK || "1xD8Uy1lUgvNTQ2RGRBI4ZjOrozbinUPRq2_UfIplP98";

    const [indexRes, jobOrderRes, issuesRes, kajRes, barcodeRes, rawpackRes] = await Promise.all([
      fetchSheetDataFromBackend(mainId, "Index!A:AA"),
      fetchSheetDataFromBackend(jobOrderId, "JobOrder!A:Z"),
      fetchSheetDataFromBackend(issuesId, "Issues!A:R"),
      fetchSheetDataFromBackend(workingUpdatesId, "KajButton!B:O"),
      fetchSheetDataFromBackend(barcodeId, "LotBarcodeData!A:Z"),
      fetchSheetDataFromBackend(rawpackId, "RAWPACK!A:ZZ")
    ]);

    // 1. Dynamic JobOrder
    const jobRows = jobOrderRes.values || [];
    const jobHeaders = jobRows[0] || [];
    const jLotIdx = findCol(jobHeaders, ["lot number", "lot no", "lot #", "lot"]);
    const jNoIdx = findCol(jobHeaders, ["job order no", "job order", "job no"]);
    const jDateIdx = findCol(jobHeaders, ["date", "job date", "order date"]);
    const jPartyIdx = findCol(jobHeaders, ["party name", "party", "client"]);
    const jGarmentIdx = findCol(jobHeaders, ["garment type", "garment"]);
    const jFabricIdx = findCol(jobHeaders, ["fabric", "fabric type"]);
    const jStyleIdx = findCol(jobHeaders, ["style", "style no"]);
    const jQtyIdx = findCol(jobHeaders, ["quantity", "total pcs", "qty", "pcs"]);
    const jEmbIdx = findCol(jobHeaders, ["emb", "embroidery"]);
    const jPrintIdx = findCol(jobHeaders, ["printing", "print"]);

    let jobOrderMatch = null;
    if (jLotIdx !== -1) {
      for (let i = 1; i < jobRows.length; i++) {
        const row = jobRows[i];
        if (normalizeLot(row[jLotIdx]) === cleanLot) {
          jobOrderMatch = {
            jobOrderNo: row[jNoIdx] || "",
            lotNo: cleanLot,
            jobDate: formatReadableDate(row[jDateIdx]),
            party: row[jPartyIdx] || "",
            garment: row[jGarmentIdx] || "",
            fabric: row[jFabricIdx] || "",
            style: row[jStyleIdx] || "",
            totalPcs: row[jQtyIdx] || "",
            embRequired: String(row[jEmbIdx] || "").toLowerCase().includes("yes"),
            printRequired: String(row[jPrintIdx] || "").toLowerCase().includes("yes")
          };
          break;
        }
      }
    }

    // 2. Dynamic Index
    const indexRows = indexRes.values || [];
    const indexHeaders = indexRows[0] || [];
    const iLotIdx = findCol(indexHeaders, ["lot number", "lot no", "lot #", "lot"]);
    const iCutDateIdx = findCol(indexHeaders, ["saved at", "cut date", "cutting date", "date"]);
    const iCutQtyIdx = findCol(indexHeaders, ["cutting qty", "cutting pcs", "cut qty", "pcs", "quantity"]);
    const iPartyIdx = findCol(indexHeaders, ["party name", "party", "client"]);
    const iGarmentIdx = findCol(indexHeaders, ["garment type", "garment"]);
    const iFabricIdx = findCol(indexHeaders, ["fabric"]);
    const iStyleIdx = findCol(indexHeaders, ["style"]);
    const iChallanIdx = findCol(indexHeaders, ["challan history", "challan"]);
    const iIssueDateIdx = findCol(indexHeaders, ["date of issue", "issue date", "stitching issue"]);
    const iSupIdx = findCol(indexHeaders, ["supervisor", "stitching supervisor", "stiching supervisor"]);
    const iWipIdx = findCol(indexHeaders, ["wip status", "wip"]);
    const iCompIdx = findCol(indexHeaders, ["completed status", "complete status", "completed date", "status"]);
    const iImageIdx = findCol(indexHeaders, ["image", "image url", "photo"]);

    let indexMatch = null;
    if (iLotIdx !== -1) {
      for (let i = 1; i < indexRows.length; i++) {
        const row = indexRows[i];
        if (normalizeLot(row[iLotIdx]) === cleanLot) {
          indexMatch = {
            lotNo: cleanLot,
            cutDate: formatReadableDate(row[iCutDateIdx]),
            cuttingQty: row[iCutQtyIdx] || "",
            party: row[iPartyIdx] || "",
            garment: row[iGarmentIdx] || "",
            fabric: row[iFabricIdx] || "",
            style: row[iStyleIdx] || "",
            challanHistoryRaw: row[iChallanIdx] || "",
            dateOfIssue: formatReadableDate(row[iIssueDateIdx]),
            stitchingSupervisor: row[iSupIdx] || "",
            wipStatus: row[iWipIdx] || "",
            completedStatus: formatReadableDate(row[iCompIdx]),
            image: row[iImageIdx] || ""
          };
          break;
        }
      }
    }

    // 3. Dynamic Issues (Packing)
    const issuesRows = issuesRes.values || [];
    const issuesHeaders = issuesRows[0] || [];
    const pLotIdx = findCol(issuesHeaders, ["lot number", "lot no", "lot"]);
    const pDateIdx = findCol(issuesHeaders, ["packing date", "date", "pkg date", "date of issue"]);
    const pSupIdx = findCol(issuesHeaders, ["packing supervisor", "supervisor"]);
    const pWipIdx = findCol(issuesHeaders, ["wip packing", "remarks", "status"]);
    const pCompIdx = findCol(issuesHeaders, ["packing complete", "packing completed", "complete date", "completed date"]);

    let issuesMatch = null;
    if (pLotIdx !== -1) {
      for (let i = 1; i < issuesRows.length; i++) {
        const row = issuesRows[i];
        if (normalizeLot(row[pLotIdx]) === cleanLot) {
          issuesMatch = {
            lotNo: cleanLot,
            pkgDate: formatReadableDate(row[pDateIdx]),
            pkgSupervisor: row[pSupIdx] || "",
            wipPacking: row[pWipIdx] || "",
            packingComplete: formatReadableDate(row[pCompIdx])
          };
          break;
        }
      }
    }

    // 4. Dynamic KajButton
    const kajRows = kajRes.values || [];
    const kajHeaders = kajRows[0] || [];
    const kLotIdx = findCol(kajHeaders, ["lot number", "lot no", "lot"]);
    const kDateIdx = findCol(kajHeaders, ["kajbutton date", "date", "issue date"]);
    const kSupIdx = findCol(kajHeaders, ["kajbutton supervisor", "supervisor"]);
    const kPcsIdx = findCol(kajHeaders, ["total pcs", "pcs", "quantity"]);
    const kAgingIdx = findCol(kajHeaders, ["aging"]);
    const kStatusIdx = findCol(kajHeaders, ["status"]);
    const kRemarksIdx = findCol(kajHeaders, ["remarks", "recent remarks"]);
    const kStitchSupIdx = findCol(kajHeaders, ["stiching supervisor", "stitching supervisor"]);

    let kajMatch = null;
    if (kLotIdx !== -1) {
      for (let i = 1; i < kajRows.length; i++) {
        const row = kajRows[i];
        if (normalizeLot(row[kLotIdx]) === cleanLot) {
          kajMatch = {
            lotNo: cleanLot,
            kajDate: formatReadableDate(row[kDateIdx]),
            supervisor: row[kSupIdx] || "",
            totalPcs: row[kPcsIdx] || "",
            aging: row[kAgingIdx] || "",
            status: row[kStatusIdx] || "",
            remarks: row[kRemarksIdx] || "",
            stitchingSupervisor: row[kStitchSupIdx] || ""
          };
          break;
        }
      }
    }

    // 5. Dynamic Barcode
    const barcodeRows = barcodeRes.values || [];
    let barcodeMatch = null;
    if (barcodeRows.length > 0) {
      const bHeaders = barcodeRows[0] || [];
      const bLotIdx = findCol(bHeaders, ["lot number", "lot no", "lot"]);
      const bDateIdx = findCol(bHeaders, ["date", "timestamp", "scan date"]);
      const bCartonIdx = findCol(bHeaders, ["carton", "cartons", "count"]);
      if (bLotIdx !== -1) {
        for (let i = 1; i < barcodeRows.length; i++) {
          const row = barcodeRows[i];
          if (normalizeLot(row[bLotIdx]) === cleanLot) {
            barcodeMatch = {
              lotNo: cleanLot,
              barcodeDate: formatReadableDate(row[bDateIdx]),
              cartonCount: row[bCartonIdx] || ""
            };
            break;
          }
        }
      }
    }

    if (!jobOrderMatch && !indexMatch && !issuesMatch && !kajMatch && !barcodeMatch) {
      throw new Error(`No manufacturing records found for Lot #${cleanLot}`);
    }

    const bestParty = jobOrderMatch?.party || indexMatch?.party || "MH";
    const bestFabric = indexMatch?.fabric || jobOrderMatch?.fabric || "N/A";
    const bestStyle = indexMatch?.style || jobOrderMatch?.style || "Standard";
    const bestGarment = indexMatch?.garment || jobOrderMatch?.garment || "Garment";
    const bestTotalPcs = indexMatch?.cuttingQty || jobOrderMatch?.totalPcs || kajMatch?.totalPcs || "N/A";

    let embPrintChallans = [];
    let embPrintStatus = "Not Required / None";
    let embPrintIssueDate = "";
    let embPrintCompleteDate = "";
    const rawChallan = indexMatch?.challanHistoryRaw || "";
    if (rawChallan) {
      try {
        if (rawChallan.trim().startsWith("[")) {
          const parsed = JSON.parse(rawChallan);
          if (Array.isArray(parsed) && parsed.length > 0) {
            embPrintChallans = parsed;
            const first = parsed[0];
            embPrintIssueDate = formatReadableDate(first.date || first.dateOfIssue || "");
            const last = parsed[parsed.length - 1];
            if (last.embCompleted || last.embUpdatedAt) {
              embPrintCompleteDate = formatReadableDate(last.embUpdatedAt || last.date || "");
              embPrintStatus = "Completed";
            } else {
              embPrintStatus = "In Progress / Pending";
            }
          }
        } else {
          const chMatch = rawChallan.match(/CH-(EMB|PRINT)-\d+/gi);
          if (chMatch) {
            embPrintChallans = chMatch.map((c) => ({ number: c }));
            embPrintStatus = "Challan Created";
          }
        }
      } catch (e) {
        embPrintStatus = "Recorded";
      }
    }

    const milestones = [];

    // Stage 1: Order / Job Creation
    const hasJob = !!jobOrderMatch || !!indexMatch;
    const jobOrderIssueDate = jobOrderMatch?.jobDate || indexMatch?.cutDate || "N/A";
    const jobOrderCompDate = jobOrderMatch?.jobDate || indexMatch?.cutDate || "Confirmed";
    milestones.push({
      id: "job_order",
      stageNumber: 1,
      title: "Job Order Created",
      subtitle: `Fabric: ${bestFabric} • Style: ${bestStyle}`,
      icon: "📋",
      status: hasJob ? "completed" : "pending",
      issueDate: jobOrderIssueDate,
      completeDate: jobOrderCompDate,
      dwellDays: null,
      details: {
        jobOrderNo: jobOrderMatch?.jobOrderNo || "Assigned",
        party: bestParty,
        style: bestStyle,
        fabric: bestFabric,
        totalPcs: bestTotalPcs,
        embRequired: jobOrderMatch?.embRequired || false,
        printRequired: jobOrderMatch?.printRequired || false
      }
    });

    // Stage 2: Cutting Department
    const hasCut = !!indexMatch?.cutDate;
    const cutIssueDate = jobOrderMatch?.jobDate || indexMatch?.cutDate || "N/A";
    const cutCompleteDate = indexMatch?.cutDate || "Pending";
    milestones.push({
      id: "cutting",
      stageNumber: 2,
      title: "Fabric Cutting",
      subtitle: hasCut ? `Cut Quantity: ${bestTotalPcs} Pcs` : "Layers cut & bundled",
      icon: "✂️",
      status: hasCut ? "completed" : "pending",
      issueDate: cutIssueDate,
      completeDate: cutCompleteDate,
      dwellDays: calculateDaysDiff(cutIssueDate, cutCompleteDate),
      details: {
        cuttingQty: bestTotalPcs,
        partyInitials: bestParty
      }
    });

    // Stage 3: Embroidery / Printing
    const hasEmbPrint = embPrintChallans.length > 0 || jobOrderMatch?.embRequired || jobOrderMatch?.printRequired;
    if (hasEmbPrint || rawChallan) {
      const isEmbDone =
        embPrintStatus.toLowerCase().includes("done") ||
        embPrintStatus.toLowerCase().includes("complete") ||
        !!embPrintCompleteDate;
      const embIssueDate = embPrintIssueDate || indexMatch?.cutDate || "Pending";
      const embCompDate = embPrintCompleteDate || (isEmbDone ? embIssueDate : "Pending");
      milestones.push({
        id: "emb_print",
        stageNumber: 3,
        title: "Embroidery / Printing",
        subtitle: isEmbDone ? "Decorations & screens processed" : "Active in EMB/Print Unit",
        icon: "🎨",
        status: isEmbDone ? "completed" : hasCut ? "in_progress" : "pending",
        issueDate: embIssueDate,
        completeDate: embCompDate,
        dwellDays: calculateDaysDiff(embIssueDate, embCompDate),
        details: {
          challans: embPrintChallans,
          statusText: embPrintStatus
        }
      });
    }

    // Stage 4: Stitching Line Allocation & Floor Assembly
    const hasStitchIssue = !!indexMatch?.dateOfIssue;
    const hasStitchComplete =
      !!indexMatch?.completedStatus &&
      indexMatch.completedStatus !== "-" &&
      !indexMatch.completedStatus.toLowerCase().includes("pending");
    const stitchIssueDate = indexMatch?.dateOfIssue || "Pending";
    const stitchCompleteDate = indexMatch?.completedStatus || (hasStitchComplete ? "Completed" : "In Progress");

    milestones.push({
      id: "stitching",
      stageNumber: 4,
      title: "Stitching Department",
      subtitle: indexMatch?.stitchingSupervisor
        ? `Supervisor: ${indexMatch.stitchingSupervisor}`
        : "Floor assembly line queue",
      icon: "🪡",
      status: hasStitchComplete ? "completed" : hasStitchIssue ? "in_progress" : "pending",
      issueDate: stitchIssueDate,
      completeDate: stitchCompleteDate,
      dwellDays: calculateDaysDiff(stitchIssueDate, stitchCompleteDate),
      details: {
        supervisor: indexMatch?.stitchingSupervisor || "Not Assigned",
        wipStatus: indexMatch?.wipStatus || "In Progress"
      }
    });

    // Stage 5: Kaj Button & Secondary Work
    const hasKaj = !!kajMatch;
    const isKajDone =
      kajMatch?.status?.toLowerCase().includes("complete") ||
      kajMatch?.status?.toLowerCase().includes("done") ||
      !!issuesMatch?.pkgDate;
    const kajIssueDate = kajMatch?.kajDate || indexMatch?.completedStatus || "Pending";
    const kajCompDate = isKajDone ? (issuesMatch?.pkgDate || kajMatch?.kajDate || "Completed") : "In Progress";

    milestones.push({
      id: "kaj_button",
      stageNumber: 5,
      title: "Kaj Button & Secondary Work",
      subtitle: kajMatch?.supervisor ? `Supervisor: ${kajMatch.supervisor}` : "Button attachment & inspection",
      icon: "🔘",
      status: isKajDone ? "completed" : hasStitchComplete ? "in_progress" : "pending",
      issueDate: kajIssueDate,
      completeDate: kajCompDate,
      dwellDays: calculateDaysDiff(kajIssueDate, kajCompDate),
      details: {
        supervisor: kajMatch?.supervisor || "N/A",
        totalPcs: kajMatch?.totalPcs || bestTotalPcs,
        agingDays: kajMatch?.aging || "0",
        remarks: kajMatch?.remarks || "None",
        status: kajMatch?.status || (isKajDone ? "Completed" : "Pending")
      }
    });

    // Stage 6: Packing & Carton Boxing
    const hasPkgIssue = !!issuesMatch?.pkgDate;
    const hasPkgComplete =
      !!issuesMatch?.packingComplete &&
      issuesMatch.packingComplete !== "-" &&
      !issuesMatch.packingComplete.toLowerCase().includes("pending");
    const hasBarcode = !!barcodeMatch;
    const isFullyComplete = hasPkgComplete || hasBarcode;

    const pkgIssueDate = issuesMatch?.pkgDate || "Pending";
    const pkgCompDate = issuesMatch?.packingComplete || barcodeMatch?.barcodeDate || (isFullyComplete ? "Completed" : "In Progress");

    milestones.push({
      id: "packing",
      stageNumber: 6,
      title: "Packing & Finishing",
      subtitle: issuesMatch?.pkgSupervisor ? `Supervisor: ${issuesMatch.pkgSupervisor}` : "Carton boxing & sticker allocation",
      icon: "📦",
      status: isFullyComplete ? "completed" : hasPkgIssue ? "in_progress" : "pending",
      issueDate: pkgIssueDate,
      completeDate: pkgCompDate,
      dwellDays: calculateDaysDiff(pkgIssueDate, pkgCompDate),
      details: {
        supervisor: issuesMatch?.pkgSupervisor || "N/A",
        wipRemarks: issuesMatch?.wipPacking || "",
        barcodeScanned: hasBarcode,
        barcodeDate: barcodeMatch?.barcodeDate || "N/A"
      }
    });

    const completedCount = milestones.filter((m) => m.status === "completed").length;
    const progressPercent = Math.round((completedCount / milestones.length) * 100);

    let currentStage = "Order Received";
    const inProgressMilestone = milestones.find((m) => m.status === "in_progress");
    if (inProgressMilestone) {
      currentStage = inProgressMilestone.title;
    } else if (isFullyComplete) {
      currentStage = "Ready for Dispatch / Completed";
    } else {
      const lastCompleted = [...milestones].reverse().find((m) => m.status === "completed");
      currentStage = lastCompleted ? `Completed ${lastCompleted.title}` : "Pending Initial Operations";
    }

    return {
      success: true,
      lotNumber: cleanLot,
      summary: {
        party: bestParty,
        style: bestStyle,
        fabric: bestFabric,
        garment: bestGarment,
        totalPcs: bestTotalPcs,
        currentStage,
        progressPercent,
        isFullyComplete,
        totalLeadTimeDays: isFullyComplete ? "Verified Complete" : "In Progress",
        image: indexMatch?.image || ""
      },
      milestones
    };
  };

  // Fetch timeline data from backend endpoint with seamless Google Sheets fallback
  const fetchTimeline = async (lotNumber, forceRefresh = false) => {
    if (!lotNumber || !lotNumber.trim()) return;

    setLoading(true);
    setError(null);

    const cleanLot = normalizeLot(lotNumber);

    try {
      const baseUrl = (BACKEND_URL || "http://localhost:5000").replace(/\/$/, "");
      const res = await fetch(
        `${baseUrl}/api/timeline/lot/${cleanLot}${forceRefresh ? "?refresh=true" : ""}`
      );

      if (res.ok) {
        const json = await res.json();
        if (json.success) {
          setTimelineData(json);
          setActiveLot(cleanLot);
          setRecentLots((prev) => [cleanLot, ...prev.filter((l) => l !== cleanLot)].slice(0, 8));
          setLoading(false);
          return;
        }
      }
      throw new Error("Backend response error, attempting direct Google API fallback...");
    } catch (backendErr) {
      console.warn("Backend timeline endpoint unavailable, falling back to direct Google Sheets API:", backendErr.message);

      try {
        const directData = await buildTimelineFromSheets(cleanLot);
        setTimelineData(directData);
        setActiveLot(cleanLot);
        setRecentLots((prev) => [cleanLot, ...prev.filter((l) => l !== cleanLot)].slice(0, 8));
      } catch (sheetErr) {
        console.error("Error building timeline from Google Sheets fallback:", sheetErr);
        setError(sheetErr.message || `No records found for Lot #${cleanLot}`);
        setTimelineData(null);
      } finally {
        setLoading(false);
      }
    }
  };

  useEffect(() => {
    if (activeLot) {
      fetchTimeline(activeLot);
    }
  }, []);

  const handleSearch = (e) => {
    e.preventDefault();
    if (lotInput.trim()) {
      fetchTimeline(lotInput.trim());
    }
  };

  const handleSelectRecent = (lot) => {
    setLotInput(lot);
    fetchTimeline(lot);
  };

  // LIGHT THEME COLOR PALETTE (default) vs DARK THEME
  const theme = isDarkMode
    ? {
        bg: "#0b0f19",
        bgGradient: `
          radial-gradient(at 0% 0%, rgba(99, 102, 241, 0.12) 0px, transparent 50%),
          radial-gradient(at 100% 0%, rgba(236, 72, 153, 0.08) 0px, transparent 50%),
          linear-gradient(180deg, #0b0f19 0%, #111827 100%)
        `,
        textPrimary: "#f8fafc",
        textSecondary: "#94a3b8",
        cardBg: "#1e293b",
        cardBorder: "rgba(255, 255, 255, 0.12)",
        stepCardBg: "#111827",
        stepItemBg: "rgba(255, 255, 255, 0.03)",
        stepItemBorder: "rgba(255, 255, 255, 0.08)",
        trackLinePending: "rgba(255, 255, 255, 0.1)",
        inputBg: "rgba(15, 23, 42, 0.6)",
        inputBorder: "rgba(255, 255, 255, 0.2)",
        inputText: "#ffffff",
        heroBg: "linear-gradient(135deg, #1e1b4b 0%, #312e81 100%)",
        heroBorder: "rgba(255, 255, 255, 0.15)",
        heroText: "#ffffff",
        metricBg: "rgba(255, 255, 255, 0.05)",
        metricBorder: "rgba(255, 255, 255, 0.1)",
        tagBg: "rgba(255, 255, 255, 0.1)",
        tagBorder: "rgba(255, 255, 255, 0.15)",
        tagText: "#ffffff",
        datePillBg: "rgba(255, 255, 255, 0.06)",
        shadow: "0 20px 40px rgba(0, 0, 0, 0.4)"
      }
    : {
        bg: "#f8fafc",
        bgGradient: `
          radial-gradient(at 0% 0%, rgba(99, 102, 241, 0.05) 0px, transparent 50%),
          radial-gradient(at 100% 0%, rgba(245, 158, 11, 0.05) 0px, transparent 50%),
          linear-gradient(180deg, #f8fafc 0%, #f1f5f9 100%)
        `,
        textPrimary: "#0f172a",
        textSecondary: "#64748b",
        cardBg: "#ffffff",
        cardBorder: "#e2e8f0",
        stepCardBg: "#ffffff",
        stepItemBg: "#f8fafc",
        stepItemBorder: "#e2e8f0",
        trackLinePending: "#e2e8f0",
        inputBg: "#ffffff",
        inputBorder: "#cbd5e1",
        inputText: "#0f172a",
        heroBg: "linear-gradient(135deg, #ffffff 0%, #f8fafc 100%)",
        heroBorder: "#e2e8f0",
        heroText: "#0f172a",
        metricBg: "#f8fafc",
        metricBorder: "#e2e8f0",
        tagBg: "#f1f5f9",
        tagBorder: "#cbd5e1",
        tagText: "#334155",
        datePillBg: "#f1f5f9",
        shadow: "0 10px 30px -5px rgba(0, 0, 0, 0.06), 0 4px 6px -2px rgba(0, 0, 0, 0.02)"
      };

  return (
    <>
      <style>{`
        @media print {
          body {
            background: #ffffff !important;
            color: #000000 !important;
          }
          .no-print {
            display: none !important;
          }
          .timeline-wrapper {
            padding: 0 !important;
            background: #ffffff !important;
          }
          .timeline-card {
            box-shadow: none !important;
            border: 1px solid #cbd5e1 !important;
          }
        }
      `}</style>

      <div
        className="timeline-wrapper"
        style={{
          minHeight: "100vh",
          backgroundColor: theme.bg,
          backgroundImage: theme.bgGradient,
          color: theme.textPrimary,
          fontFamily: "'Plus Jakarta Sans', 'Inter', system-ui, -apple-system, sans-serif",
          padding: "32px 24px 80px",
          transition: "all 0.3s ease"
        }}
      >
        <div style={{ maxWidth: "1200px", margin: "0 auto" }}>
          {/* Top Bar Header */}
          <div
            className="no-print"
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              flexWrap: "wrap",
              gap: "16px",
              marginBottom: "28px"
            }}
          >
            <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
              <button
                onClick={() => history.push("/dashboard")}
                style={{
                  background: isDarkMode ? "rgba(255, 255, 255, 0.08)" : "#ffffff",
                  border: `1.5px solid ${theme.cardBorder}`,
                  color: theme.textPrimary,
                  padding: "8px 16px",
                  borderRadius: "12px",
                  cursor: "pointer",
                  fontWeight: 700,
                  fontSize: "0.85rem",
                  display: "flex",
                  alignItems: "center",
                  gap: "6px",
                  boxShadow: isDarkMode ? "none" : "0 2px 5px rgba(0,0,0,0.04)"
                }}
              >
                ← Dashboard
              </button>
              <div>
                <h1 style={{ margin: 0, fontSize: "1.5rem", fontWeight: 900, color: theme.textPrimary, display: "flex", alignItems: "center", gap: "8px" }}>
                  <span>📦⚡</span> Amazon-Style Lot Lifecycle Tracker
                </h1>
                <span style={{ fontSize: "0.8rem", color: theme.textSecondary, fontWeight: 600 }}>
                  End-to-End Production & Milestone Journey
                </span>
              </div>
            </div>

            <div style={{ display: "flex", gap: "10px", alignItems: "center" }}>
              {/* Theme Toggle Button */}
              <button
                onClick={() => setIsDarkMode(!isDarkMode)}
                title="Toggle Light / Dark Mode"
                style={{
                  background: isDarkMode ? "rgba(255, 255, 255, 0.1)" : "#ffffff",
                  border: `1.5px solid ${theme.cardBorder}`,
                  color: theme.textPrimary,
                  padding: "8px 14px",
                  borderRadius: "12px",
                  cursor: "pointer",
                  fontWeight: 800,
                  fontSize: "0.85rem",
                  display: "flex",
                  alignItems: "center",
                  gap: "6px",
                  boxShadow: isDarkMode ? "none" : "0 2px 5px rgba(0,0,0,0.04)"
                }}
              >
                {isDarkMode ? "☀️ Light Mode" : "🌙 Dark Mode"}
              </button>

              <button
                onClick={() => fetchTimeline(activeLot, true)}
                style={{
                  background: isDarkMode ? "rgba(255, 255, 255, 0.08)" : "#ffffff",
                  border: `1.5px solid ${theme.cardBorder}`,
                  color: theme.textPrimary,
                  padding: "8px 16px",
                  borderRadius: "12px",
                  cursor: "pointer",
                  fontWeight: 700,
                  fontSize: "0.85rem",
                  display: "flex",
                  alignItems: "center",
                  gap: "6px",
                  boxShadow: isDarkMode ? "none" : "0 2px 5px rgba(0,0,0,0.04)"
                }}
              >
                ↻ Refresh Live
              </button>
              <button
                onClick={() => window.print()}
                style={{
                  background: "linear-gradient(135deg, #f59e0b 0%, #d97706 100%)",
                  border: "none",
                  color: "#ffffff",
                  padding: "8px 18px",
                  borderRadius: "12px",
                  cursor: "pointer",
                  fontWeight: 800,
                  fontSize: "0.85rem",
                  boxShadow: "0 4px 14px rgba(245, 158, 11, 0.35)",
                  display: "flex",
                  alignItems: "center",
                  gap: "6px"
                }}
              >
                🖨️ Print Lot Passport
              </button>
            </div>
          </div>

          {/* Search Hero Card */}
          <div
            className="no-print"
            style={{
              background: theme.heroBg,
              borderRadius: "24px",
              padding: "28px 36px",
              border: `1.5px solid ${theme.heroBorder}`,
              boxShadow: theme.shadow,
              marginBottom: "32px"
            }}
          >
            <form onSubmit={handleSearch} style={{ display: "flex", gap: "12px", flexWrap: "wrap" }}>
              <div style={{ flex: 1, minWidth: "260px", position: "relative" }}>
                <input
                  type="text"
                  placeholder="Enter Lot Number (e.g. 76063, 65002)..."
                  value={lotInput}
                  onChange={(e) => setLotInput(e.target.value)}
                  style={{
                    width: "100%",
                    padding: "16px 20px 16px 48px",
                    borderRadius: "16px",
                    border: `2px solid ${theme.inputBorder}`,
                    background: theme.inputBg,
                    color: theme.inputText,
                    fontSize: "1.1rem",
                    fontWeight: 700,
                    outline: "none",
                    boxSizing: "border-box",
                    boxShadow: isDarkMode ? "none" : "inset 0 2px 4px rgba(0,0,0,0.02)"
                  }}
                />
                <span style={{ position: "absolute", left: "18px", top: "50%", transform: "translateY(-50%)", fontSize: "1.2rem", opacity: 0.6 }}>
                  🔍
                </span>
              </div>

              <button
                type="submit"
                disabled={loading}
                style={{
                  background: "linear-gradient(135deg, #f59e0b 0%, #ea580c 100%)",
                  border: "none",
                  color: "#ffffff",
                  padding: "0 32px",
                  borderRadius: "16px",
                  cursor: "pointer",
                  fontWeight: 900,
                  fontSize: "1.05rem",
                  letterSpacing: "0.5px",
                  boxShadow: "0 4px 15px rgba(245, 158, 11, 0.4)",
                  display: "flex",
                  alignItems: "center",
                  gap: "8px",
                  minHeight: "56px"
                }}
              >
                {loading ? "Searching..." : "Track Lot →"}
              </button>
            </form>

            {/* Quick Recent Search Tags */}
            <div style={{ display: "flex", alignItems: "center", gap: "10px", marginTop: "18px", flexWrap: "wrap" }}>
              <span style={{ fontSize: "0.8rem", color: theme.textSecondary, fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.5px" }}>
                Recent Searches:
              </span>
              {recentLots.map((lot) => (
                <button
                  key={lot}
                  type="button"
                  onClick={() => handleSelectRecent(lot)}
                  style={{
                    background: activeLot === lot ? (isDarkMode ? "#4338ca" : "#fef3c7") : theme.tagBg,
                    border: activeLot === lot ? "1.5px solid #f59e0b" : `1px solid ${theme.tagBorder}`,
                    color: activeLot === lot ? (isDarkMode ? "#ffffff" : "#b45309") : theme.tagText,
                    padding: "5px 14px",
                    borderRadius: "10px",
                    fontSize: "0.82rem",
                    fontWeight: 800,
                    cursor: "pointer",
                    transition: "all 0.18s"
                  }}
                >
                  Lot #{lot}
                </button>
              ))}
            </div>
          </div>

          {/* Loading State */}
          {loading && (
            <div style={{ textAlign: "center", padding: "80px 20px" }}>
              <div style={{ fontSize: "3rem", marginBottom: "16px", animation: "spin 1.5s linear infinite" }}>
                ⚡
              </div>
              <h3 style={{ margin: 0, fontSize: "1.3rem", fontWeight: 800, color: theme.textPrimary }}>
                Fetching Lot #{lotInput} Lifecycle Across All Stages...
              </h3>
              <p style={{ margin: "8px 0 0 0", color: theme.textSecondary, fontSize: "0.9rem" }}>
                Querying Cutting, Embroidery, Printing, Stitching, Kaj Button, and Packing records...
              </p>
            </div>
          )}

          {/* Error State */}
          {!loading && error && (
            <div
              style={{
                background: isDarkMode ? "rgba(239, 68, 68, 0.1)" : "#fef2f2",
                border: "1.5px solid #f87171",
                borderRadius: "20px",
                padding: "36px",
                textAlign: "center",
                color: isDarkMode ? "#fca5a5" : "#991b1b",
                boxShadow: theme.shadow
              }}
            >
              <div style={{ fontSize: "2.5rem", marginBottom: "12px" }}>⚠️</div>
              <h3 style={{ margin: 0, fontSize: "1.25rem", fontWeight: 800, color: isDarkMode ? "#ffffff" : "#991b1b" }}>
                {error}
              </h3>
              <p style={{ margin: "8px 0 20px 0", fontSize: "0.9rem", color: isDarkMode ? "#fca5a5" : "#b91c1c" }}>
                Please check the lot number or verify that it has been registered in Job Orders or Cutting.
              </p>
              <button
                onClick={() => fetchTimeline("76063")}
                style={{
                  background: "#ef4444",
                  border: "none",
                  color: "#ffffff",
                  padding: "8px 20px",
                  borderRadius: "10px",
                  fontWeight: 700,
                  cursor: "pointer"
                }}
              >
                Try Sample Lot #76063
              </button>
            </div>
          )}

          {/* Main Amazon-Style Order Status Display */}
          {!loading && timelineData && (
            <motion.div
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ duration: 0.4 }}
            >
              {/* Order Header Card */}
              <div
                style={{
                  background: theme.cardBg,
                  borderRadius: "24px",
                  border: `1.5px solid ${theme.cardBorder}`,
                  boxShadow: theme.shadow,
                  padding: "32px",
                  marginBottom: "32px",
                  position: "relative",
                  overflow: "hidden"
                }}
              >
                {/* Top Delivery Header Banner */}
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "flex-start",
                    flexWrap: "wrap",
                    gap: "20px",
                    borderBottom: `1px solid ${theme.cardBorder}`,
                    paddingBottom: "24px",
                    marginBottom: "24px"
                  }}
                >
                  <div>
                    <div
                      style={{
                        display: "inline-flex",
                        alignItems: "center",
                        gap: "8px",
                        background: timelineData.summary.isFullyComplete
                          ? (isDarkMode ? "rgba(16, 185, 129, 0.2)" : "#ecfdf5")
                          : (isDarkMode ? "rgba(245, 158, 11, 0.2)" : "#fffbeb"),
                        border: `1px solid ${timelineData.summary.isFullyComplete ? "#10b981" : "#f59e0b"}`,
                        padding: "5px 14px",
                        borderRadius: "9999px",
                        fontSize: "0.8rem",
                        fontWeight: 800,
                        color: timelineData.summary.isFullyComplete ? "#047857" : "#b45309",
                        marginBottom: "12px"
                      }}
                    >
                      <span
                        style={{
                          width: "8px",
                          height: "8px",
                          borderRadius: "50%",
                          background: timelineData.summary.isFullyComplete ? "#10b981" : "#f59e0b"
                        }}
                      />
                      <span>{timelineData.summary.isFullyComplete ? "MANUFACTURING COMPLETED" : "IN PRODUCTION PIPELINE"}</span>
                    </div>

                    <h2 style={{ margin: 0, fontSize: "1.85rem", fontWeight: 900, color: theme.textPrimary, letterSpacing: "-0.02em" }}>
                      {timelineData.summary.currentStage}
                    </h2>
                    <p style={{ margin: "8px 0 0 0", color: theme.textSecondary, fontSize: "0.95rem", lineHeight: "1.5" }}>
                      Lot #{timelineData.lotNumber} • Style: <strong style={{ color: theme.textPrimary }}>{timelineData.summary.style}</strong> • Fabric: <strong style={{ color: theme.textPrimary }}>{timelineData.summary.fabric}</strong>
                    </p>
                  </div>

                  {/* Summary Metric Badges */}
                  <div style={{ display: "flex", gap: "12px", flexWrap: "wrap" }}>
                    <div style={{ background: theme.metricBg, padding: "14px 20px", borderRadius: "16px", border: `1px solid ${theme.metricBorder}`, textAlign: "center", minWidth: "110px" }}>
                      <div style={{ fontSize: "1.35rem", fontWeight: 900, color: "#0284c7" }}>{timelineData.summary.totalPcs}</div>
                      <div style={{ fontSize: "0.72rem", fontWeight: 800, color: theme.textSecondary, textTransform: "uppercase", marginTop: "2px" }}>Total Pieces</div>
                    </div>
                    <div style={{ background: theme.metricBg, padding: "14px 20px", borderRadius: "16px", border: `1px solid ${theme.metricBorder}`, textAlign: "center", minWidth: "110px" }}>
                      <div style={{ fontSize: "1.35rem", fontWeight: 900, color: "#7c3aed" }}>{timelineData.summary.party}</div>
                      <div style={{ fontSize: "0.72rem", fontWeight: 800, color: theme.textSecondary, textTransform: "uppercase", marginTop: "2px" }}>Party / Client</div>
                    </div>
                    <div style={{ background: theme.metricBg, padding: "14px 20px", borderRadius: "16px", border: `1px solid ${theme.metricBorder}`, textAlign: "center", minWidth: "110px" }}>
                      <div style={{ fontSize: "1.35rem", fontWeight: 900, color: "#059669" }}>{timelineData.summary.totalLeadTimeDays}</div>
                      <div style={{ fontSize: "0.72rem", fontWeight: 800, color: theme.textSecondary, textTransform: "uppercase", marginTop: "2px" }}>Lead Time</div>
                    </div>
                  </div>
                </div>

                {/* Master Horizontal Progress Bar (Amazon Tracker Bar) */}
                <div>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "8px", fontSize: "0.85rem", fontWeight: 800, color: theme.textSecondary }}>
                    <span>Pipeline Progress</span>
                    <span style={{ color: theme.textPrimary, fontWeight: 900 }}>{timelineData.summary.progressPercent}% Completed</span>
                  </div>
                  <div style={{ width: "100%", height: "12px", background: isDarkMode ? "rgba(255, 255, 255, 0.1)" : "#e2e8f0", borderRadius: "9999px", overflow: "hidden" }}>
                    <motion.div
                      initial={{ width: 0 }}
                      animate={{ width: `${timelineData.summary.progressPercent}%` }}
                      transition={{ duration: 0.8, ease: "easeOut" }}
                      style={{
                        height: "100%",
                        background: timelineData.summary.isFullyComplete
                          ? "linear-gradient(90deg, #10b981 0%, #059669 100%)"
                          : "linear-gradient(90deg, #6366f1 0%, #ec4899 50%, #f59e0b 100%)",
                        borderRadius: "9999px"
                      }}
                    />
                  </div>
                </div>
              </div>

              {/* Vertical Amazon-Style Delivery Timeline Stepper */}
              <div
                style={{
                  background: theme.stepCardBg,
                  borderRadius: "24px",
                  border: `1.5px solid ${theme.cardBorder}`,
                  boxShadow: theme.shadow,
                  padding: "36px 32px"
                }}
              >
                <h3 style={{ margin: "0 0 32px 0", fontSize: "1.25rem", fontWeight: 900, color: theme.textPrimary, display: "flex", alignItems: "center", gap: "10px" }}>
                  <span>📍</span> Full Step-by-Step Manufacturing Journey (Issue & Completion Details)
                </h3>

                <div style={{ position: "relative", paddingLeft: "10px" }}>
                  {timelineData.milestones.map((step, index) => {
                    const isCompleted = step.status === "completed";
                    const isInProgress = step.status === "in_progress";
                    const isPending = step.status === "pending";
                    const isLast = index === timelineData.milestones.length - 1;

                    return (
                      <div
                        key={step.id}
                        style={{
                          position: "relative",
                          display: "flex",
                          gap: "24px",
                          paddingBottom: isLast ? "0" : "40px"
                        }}
                      >
                        {/* Connecting Vertical Track Line */}
                        {!isLast && (
                          <div
                            style={{
                              position: "absolute",
                              left: "22px",
                              top: "44px",
                              bottom: "0",
                              width: "3px",
                              background: isCompleted
                                ? "#10b981"
                                : isInProgress
                                ? "linear-gradient(180deg, #f59e0b 0%, #cbd5e1 100%)"
                                : theme.trackLinePending,
                              transition: "background 0.3s ease"
                            }}
                          />
                        )}

                        {/* Milestone Circle Marker */}
                        <div
                          style={{
                            width: "46px",
                            height: "46px",
                            borderRadius: "50%",
                            background: isCompleted
                              ? "#10b981"
                              : isInProgress
                              ? "#f59e0b"
                              : (isDarkMode ? "rgba(255, 255, 255, 0.08)" : "#f1f5f9"),
                            border: isInProgress
                              ? "3px solid #fef3c7"
                              : isCompleted
                              ? "2px solid #059669"
                              : `2px solid ${isDarkMode ? "rgba(255, 255, 255, 0.15)" : "#cbd5e1"}`,
                            boxShadow: isCompleted
                              ? "0 0 16px rgba(16, 185, 129, 0.35)"
                              : isInProgress
                              ? "0 0 20px rgba(245, 158, 11, 0.45)"
                              : "none",
                            color: isCompleted || isInProgress ? "#ffffff" : "#64748b",
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "center",
                            fontSize: "1.25rem",
                            flexShrink: 0,
                            zIndex: 2
                          }}
                        >
                          {isCompleted ? "✓" : isInProgress ? step.icon : step.icon}
                        </div>

                        {/* Milestone Content Card */}
                        <div
                          style={{
                            flex: 1,
                            background: isInProgress
                              ? (isDarkMode ? "rgba(245, 158, 11, 0.06)" : "#fffbeb")
                              : isCompleted
                              ? (isDarkMode ? "rgba(255, 255, 255, 0.03)" : "#ffffff")
                              : theme.stepItemBg,
                            border: isInProgress
                              ? "1.5px solid #f59e0b"
                              : isCompleted
                              ? `1.5px solid ${theme.cardBorder}`
                              : `1px solid ${theme.stepItemBorder}`,
                            borderRadius: "18px",
                            padding: "20px 24px",
                            boxShadow: isDarkMode ? "none" : "0 2px 8px rgba(0,0,0,0.03)"
                          }}
                        >
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: "12px", marginBottom: "12px" }}>
                            <div>
                              <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                                <h4 style={{ margin: 0, fontSize: "1.15rem", fontWeight: 800, color: isPending ? theme.textSecondary : theme.textPrimary }}>
                                  {step.title}
                                </h4>
                                <span
                                  style={{
                                    fontSize: "0.7rem",
                                    fontWeight: 900,
                                    padding: "3px 9px",
                                    borderRadius: "6px",
                                    background: isCompleted
                                      ? (isDarkMode ? "rgba(16, 185, 129, 0.2)" : "#d1fae5")
                                      : isInProgress
                                      ? (isDarkMode ? "rgba(245, 158, 11, 0.2)" : "#fef3c7")
                                      : (isDarkMode ? "rgba(255, 255, 255, 0.08)" : "#e2e8f0"),
                                    color: isCompleted
                                      ? "#047857"
                                      : isInProgress
                                      ? "#b45309"
                                      : "#64748b"
                                  }}
                                >
                                  {isCompleted ? "COMPLETED" : isInProgress ? "IN PROGRESS" : "UPCOMING"}
                                </span>
                              </div>

                              <p style={{ margin: "4px 0 0 0", fontSize: "0.85rem", color: theme.textSecondary, fontWeight: 500 }}>
                                {step.subtitle}
                              </p>
                            </div>

                            {/* Dual Date Tracker Badge (Issue Date -> Completion Date) */}
                            <div
                              style={{
                                display: "flex",
                                alignItems: "center",
                                gap: "10px",
                                flexWrap: "wrap",
                                background: theme.datePillBg,
                                border: `1.5px solid ${theme.cardBorder}`,
                                padding: "8px 14px",
                                borderRadius: "12px"
                              }}
                            >
                              <div style={{ display: "flex", alignItems: "center", gap: "5px" }}>
                                <span style={{ fontSize: "0.7rem", fontWeight: 800, color: theme.textSecondary, textTransform: "uppercase" }}>📤 ISSUE:</span>
                                <span style={{ fontSize: "0.82rem", fontWeight: 800, color: step.issueDate && step.issueDate !== "Pending" ? (isDarkMode ? "#38bdf8" : "#0369a1") : theme.textSecondary }}>
                                  {step.issueDate || "Pending"}
                                </span>
                              </div>

                              <span style={{ color: theme.textSecondary, opacity: 0.6, fontSize: "0.85rem" }}>➔</span>

                              <div style={{ display: "flex", alignItems: "center", gap: "5px" }}>
                                <span style={{ fontSize: "0.7rem", fontWeight: 800, color: theme.textSecondary, textTransform: "uppercase" }}>📥 COMPLETED:</span>
                                <span style={{ fontSize: "0.82rem", fontWeight: 800, color: isCompleted ? (isDarkMode ? "#34d399" : "#047857") : (isInProgress ? "#b45309" : theme.textSecondary) }}>
                                  {step.completeDate || (isInProgress ? "In Progress" : "Pending")}
                                </span>
                              </div>

                              {step.dwellDays && (
                                <span style={{ background: isDarkMode ? "rgba(16, 185, 129, 0.2)" : "#d1fae5", color: isDarkMode ? "#34d399" : "#065f46", padding: "2px 8px", borderRadius: "6px", fontSize: "0.72rem", fontWeight: 900 }}>
                                  ⏱️ {step.dwellDays}
                                </span>
                              )}
                            </div>
                          </div>

                          {/* Specific Detail Tags / Remarks */}
                          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", marginTop: "14px" }}>
                            {step.details.supervisor && step.details.supervisor !== "N/A" && (
                              <span style={{ background: isDarkMode ? "rgba(99, 102, 241, 0.15)" : "#e0e7ff", color: isDarkMode ? "#a5b4fc" : "#3730a3", border: "1px solid #c7d2fe", padding: "4px 10px", borderRadius: "8px", fontSize: "0.78rem", fontWeight: 700 }}>
                                👤 Supervisor: {step.details.supervisor}
                              </span>
                            )}
                            {step.details.cuttingQty && step.details.cuttingQty !== "N/A" && (
                              <span style={{ background: isDarkMode ? "rgba(37, 99, 235, 0.15)" : "#e0f2fe", color: isDarkMode ? "#93c5fd" : "#0369a1", border: "1px solid #bae6fd", padding: "4px 10px", borderRadius: "8px", fontSize: "0.78rem", fontWeight: 700 }}>
                                ✂️ Cutting Qty: {step.details.cuttingQty} Pcs
                              </span>
                            )}
                            {step.details.wipStatus && step.details.wipStatus !== "In Progress" && (
                              <span style={{ background: isDarkMode ? "rgba(245, 158, 11, 0.15)" : "#fef3c7", color: isDarkMode ? "#fde68a" : "#92400e", border: "1px solid #fde68a", padding: "4px 10px", borderRadius: "8px", fontSize: "0.78rem", fontWeight: 700 }}>
                                📝 Remarks: {step.details.wipStatus}
                              </span>
                            )}
                            {step.details.agingDays && (
                              <span style={{ background: isDarkMode ? "rgba(239, 68, 68, 0.15)" : "#fee2e2", color: isDarkMode ? "#fca5a5" : "#991b1b", border: "1px solid #fecaca", padding: "4px 10px", borderRadius: "8px", fontSize: "0.78rem", fontWeight: 700 }}>
                                ⏳ Stage Aging: {step.details.agingDays} Days
                              </span>
                            )}
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </motion.div>
          )}
        </div>
      </div>
    </>
  );
}
