import React, { useState, useEffect } from "react";
import { useHistory, useLocation } from "react-router-dom";
import { motion } from "framer-motion";
import { BACKEND_URL, SPREADSHEET_IDS, fetchSheetDataFromBackend } from "./config";

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

function parseDeptSheetMatch(rows, targetLot, deptKeywords) {
  if (!rows || rows.length < 2) return null;
  const headers = rows[0] || [];
  const lotIdx = findCol(headers, ["lot number", "lot no", "lot #", "lot"]);
  const dateIdx = findCol(headers, [...deptKeywords.map(k => `${k} date`), "date", "issue date", "saved at"]);
  const supIdx = findCol(headers, [...deptKeywords.map(k => `${k} supervisor`), "supervisor", "operator"]);
  const pcsIdx = findCol(headers, ["total pcs", "pcs", "quantity"]);
  const completeIdx = findCol(headers, [...deptKeywords.map(k => `${k} complete`), ...deptKeywords.map(k => `${k} completed`), "complete date", "completed date", "status"]);
  const wipIdx = findCol(headers, [...deptKeywords.map(k => `wip ${k}`), ...deptKeywords.map(k => `${k} wip`), "wip", "remarks", "recent remarks"]);
  const agingIdx = findCol(headers, ["aging"]);
  const stitchSupIdx = findCol(headers, ["stiching supervisor", "stitching supervisor"]);

  if (lotIdx === -1) return null;

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (normalizeLot(row[lotIdx]) === targetLot) {
      const completeRaw = row[completeIdx] || "";
      const wipRaw = row[wipIdx] || "";
      let isComplete = false;
      let completeDate = "";
      let wipRemarks = "";

      if (completeRaw && completeRaw !== "[]" && completeRaw !== "-") {
        if (typeof completeRaw === "string" && (completeRaw.toLowerCase().includes("complete") || completeRaw.toLowerCase().includes("done"))) {
          isComplete = true;
        }
        try {
          const parsed = JSON.parse(completeRaw);
          if (Array.isArray(parsed) && parsed.length > 0) {
            isComplete = true;
            completeDate = formatReadableDate(parsed[parsed.length - 1].timestamp || parsed[parsed.length - 1].date);
          }
        } catch (e) {
          const d = formatReadableDate(completeRaw);
          if (d) {
            isComplete = true;
            completeDate = d;
          }
        }
      }

      if (wipRaw && wipRaw !== "[]") {
        try {
          const parsed = JSON.parse(wipRaw);
          if (Array.isArray(parsed) && parsed.length > 0) {
            wipRemarks = parsed[parsed.length - 1].remarks || parsed[parsed.length - 1].status || "";
          }
        } catch (e) {
          wipRemarks = wipRaw;
        }
      }

      return {
        lotNo: targetLot,
        date: formatReadableDate(row[dateIdx]),
        supervisor: row[supIdx] || "",
        totalPcs: row[pcsIdx] || "",
        aging: row[agingIdx] || "0",
        isComplete,
        completeDate: completeDate || (isComplete ? "Completed" : ""),
        wipRemarks: wipRemarks || "In Progress",
        stitchingSupervisor: row[stitchSupIdx] || ""
      };
    }
  }
  return null;
}

export default function LotTimelineReport() {
  const history = useHistory();
  const location = useLocation();

  const queryParams = new URLSearchParams(location.search);
  const initialLot = queryParams.get("lot") || "11360";

  const [lotInput, setLotInput] = useState(initialLot);
  const [activeLot, setActiveLot] = useState(initialLot);
  const [loading, setLoading] = useState(false);
  const [timelineData, setTimelineData] = useState(null);
  const [error, setError] = useState(null);
  const [recentLots, setRecentLots] = useState(["11360", "76063", "65002", "11841", "12140", "11874"]);
  const [isDarkMode, setIsDarkMode] = useState(false);
  const [activeFilterTab, setActiveFilterTab] = useState("all");

  const buildTimelineFromSheets = async (cleanLot) => {
    const mainId = SPREADSHEET_IDS.MAIN || "1Hj3JeJEKB43aYYWv8gk2UhdU6BWuEQfCg5pBlTdBMNA";
    const jobOrderId = SPREADSHEET_IDS.JOBORDER || "1fKSwGBIpzWEFk566WRQ4bzQ0anJlmasoY8TwrTLQHXI";
    const issuesId = SPREADSHEET_IDS.ISSUES || "1uo14nKO_yHu4AJ2rOgaJajuprcinj6xw1AUMFJ6_zYM";
    const dailyStitchingId = SPREADSHEET_IDS.DAILY_STITCHING || "1IMhmYlJ3s2PPRgEQs1Ikd4O1OBXK4EYL1oV_-kWAkyg";
    const workingUpdatesId = SPREADSHEET_IDS.WORKING_UPDATES || "1Nh7XYE_MnAxtaTRUUntHvBpzctODwjnkbBYDiYLQgoc";
    const barcodeId = SPREADSHEET_IDS.BARCODE || "1dOCjNFwaAel5qun0_ZJVIGmREqjI76CJBBFIjM3NHv8";

    const [
      indexRes,
      jobOrderRes,
      issuesRes,
      kajRes,
      barcodeRes,
      feedUpRes,
      overlockRes,
      washingRes,
      foldingRes,
      elasticRes
    ] = await Promise.all([
      fetchSheetDataFromBackend(mainId, "Index!A:AA"),
      fetchSheetDataFromBackend(jobOrderId, "JobOrder!A:Z"),
      fetchSheetDataFromBackend(issuesId, "Issues!A:R"),
      fetchSheetDataFromBackend(workingUpdatesId, "KajButton!B:O"),
      fetchSheetDataFromBackend(barcodeId, "LotBarcodeData!A:Z"),
      fetchSheetDataFromBackend(workingUpdatesId, "FeedUp!B:O"),
      fetchSheetDataFromBackend(workingUpdatesId, "Overlock!B:O"),
      fetchSheetDataFromBackend(workingUpdatesId, "Washing!B:O"),
      fetchSheetDataFromBackend(workingUpdatesId, "Folding!B:O"),
      fetchSheetDataFromBackend(workingUpdatesId, "Elastic!B:O")
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

    // 4. Department Matches
    const kajMatch = parseDeptSheetMatch(kajRes.values, cleanLot, ["kaj", "kajbutton"]);
    const feedUpMatch = parseDeptSheetMatch(feedUpRes.values, cleanLot, ["feed up", "feedup"]);
    const overlockMatch = parseDeptSheetMatch(overlockRes.values, cleanLot, ["overlock"]);
    const washingMatch = parseDeptSheetMatch(washingRes.values, cleanLot, ["washing"]);
    const foldingMatch = parseDeptSheetMatch(foldingRes.values, cleanLot, ["folding"]);
    const elasticMatch = parseDeptSheetMatch(elasticRes.values, cleanLot, ["elastic"]);

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

    if (!jobOrderMatch && !indexMatch && !issuesMatch && !kajMatch && !feedUpMatch && !overlockMatch && !washingMatch && !foldingMatch && !elasticMatch && !barcodeMatch) {
      throw new Error(`No manufacturing records found for Lot #${cleanLot}`);
    }

    const bestParty = jobOrderMatch?.party || indexMatch?.party || "MH";
    const bestFabric = indexMatch?.fabric || jobOrderMatch?.fabric || "N/A";
    const bestStyle = indexMatch?.style || jobOrderMatch?.style || "Standard";
    const bestGarment = indexMatch?.garment || jobOrderMatch?.garment || "Garment";
    const bestTotalPcs = indexMatch?.cuttingQty || jobOrderMatch?.totalPcs || kajMatch?.totalPcs || "N/A";

    let embChallan = null;
    let printChallan = null;
    const rawChallan = indexMatch?.challanHistoryRaw || "";
    if (rawChallan) {
      try {
        if (rawChallan.trim().startsWith("[")) {
          const parsed = JSON.parse(rawChallan);
          if (Array.isArray(parsed) && parsed.length > 0) {
            parsed.forEach(c => {
              const num = String(c.number || c.challanNo || "").toLowerCase();
              if (num.includes("emb")) embChallan = c;
              if (num.includes("print")) printChallan = c;
            });
            if (!embChallan && !printChallan) {
              embChallan = parsed[0];
            }
          }
        }
      } catch (e) {}
    }

    const hasCut = !!indexMatch?.cutDate;
    const hasStitchIssue = !!indexMatch?.dateOfIssue;
    const hasStitchComplete =
      !!indexMatch?.completedStatus &&
      indexMatch.completedStatus !== "-" &&
      !indexMatch.completedStatus.toLowerCase().includes("pending");

    const milestones = [];
    let stageCounter = 1;

    // 1. Job Order Created
    const hasJob = !!jobOrderMatch || !!indexMatch;
    const jobOrderIssueDate = jobOrderMatch?.jobDate || indexMatch?.cutDate || "N/A";
    const jobOrderCompDate = jobOrderMatch?.jobDate || indexMatch?.cutDate || "Confirmed";
    milestones.push({
      id: "job_order",
      stageNumber: stageCounter++,
      title: "Job Order Created",
      department: "Planning & Job Order",
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

    // 2. Fabric Cutting
    const cutIssueDate = jobOrderMatch?.jobDate || indexMatch?.cutDate || "N/A";
    const cutCompleteDate = indexMatch?.cutDate || "Pending";
    milestones.push({
      id: "cutting",
      stageNumber: stageCounter++,
      title: "Fabric Cutting",
      department: "Cutting Department",
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

    // 3. Jaybir Printing
    const printRequired = jobOrderMatch?.printRequired || !!printChallan || rawChallan.toLowerCase().includes("print");
    if (printRequired) {
      const isPrintDone = printChallan?.completed || (printChallan?.date && hasStitchIssue);
      const printDate = formatReadableDate(printChallan?.date || printChallan?.dateOfIssue) || indexMatch?.cutDate || "Pending";
      const printCompDate = isPrintDone ? formatReadableDate(printChallan?.completedAt || printDate) : "In Progress";
      milestones.push({
        id: "printing",
        stageNumber: stageCounter++,
        title: "Jaybir Printing",
        department: "Printing Department",
        subtitle: isPrintDone ? "Screen printing & curing completed" : "Active in screen printing unit",
        icon: "🖼️",
        status: isPrintDone ? "completed" : hasCut ? "in_progress" : "pending",
        issueDate: printDate,
        completeDate: printCompDate,
        dwellDays: calculateDaysDiff(printDate, printCompDate),
        details: {
          challan: printChallan?.number || "Challan Issued",
          supervisor: "Jaybir Print Sup"
        }
      });
    }

    // 4. Jaybir Embroidery
    const embRequired = jobOrderMatch?.embRequired || !!embChallan || rawChallan.toLowerCase().includes("emb");
    if (embRequired) {
      const isEmbDone = embChallan?.completed || (embChallan?.date && hasStitchIssue);
      const embDate = formatReadableDate(embChallan?.date || embChallan?.dateOfIssue) || indexMatch?.cutDate || "Pending";
      const embCompDate = isEmbDone ? formatReadableDate(embChallan?.completedAt || embDate) : "In Progress";
      milestones.push({
        id: "embroidery",
        stageNumber: stageCounter++,
        title: "Jaybir Embroidery",
        department: "Embroidery Department",
        subtitle: isEmbDone ? "Multi-head embroidery stitching completed" : "Active in embroidery unit",
        icon: "🧵",
        status: isEmbDone ? "completed" : hasCut ? "in_progress" : "pending",
        issueDate: embDate,
        completeDate: embCompDate,
        dwellDays: calculateDaysDiff(embDate, embCompDate),
        details: {
          challan: embChallan?.number || "Challan Issued",
          supervisor: "Jaybir Emb Sup"
        }
      });
    }

    // 5. Elastic Attachment
    if (elasticMatch || (indexMatch?.wipStatus && indexMatch.wipStatus.toLowerCase().includes("elastic"))) {
      const isElasticDone = elasticMatch?.isComplete || hasStitchComplete;
      const elasticDate = elasticMatch?.date || indexMatch?.cutDate || "Pending";
      const elasticCompDate = elasticMatch?.completeDate || (isElasticDone ? "Completed" : "In Progress");
      milestones.push({
        id: "elastic",
        stageNumber: stageCounter++,
        title: "Elastic Attachment",
        department: "Elastic Department",
        subtitle: elasticMatch?.supervisor ? `Supervisor: ${elasticMatch.supervisor}` : "Waistband/Cuff elastic attachment",
        icon: "🪢",
        status: isElasticDone ? "completed" : hasCut ? "in_progress" : "pending",
        issueDate: elasticDate,
        completeDate: elasticCompDate,
        dwellDays: calculateDaysDiff(elasticDate, elasticCompDate),
        details: {
          supervisor: elasticMatch?.supervisor || "Elastic Sup",
          wipRemarks: elasticMatch?.wipRemarks || ""
        }
      });
    }

    // 6. Floor Stitching Assembly
    const stitchIssueDate = indexMatch?.dateOfIssue || "Pending";
    const stitchCompleteDate = indexMatch?.completedStatus || (hasStitchComplete ? "Completed" : "In Progress");
    milestones.push({
      id: "stitching",
      stageNumber: stageCounter++,
      title: "Floor Stitching Assembly",
      department: "Stitching Department",
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

    // 7. Feed Up Department
    if (feedUpMatch || (indexMatch?.wipStatus && indexMatch.wipStatus.toLowerCase().includes("feed"))) {
      const isFeedDone = feedUpMatch?.isComplete || hasStitchComplete;
      const feedDate = feedUpMatch?.date || indexMatch?.dateOfIssue || "Pending";
      const feedCompDate = feedUpMatch?.completeDate || (isFeedDone ? "Completed" : "In Progress");
      milestones.push({
        id: "feedup",
        stageNumber: stageCounter++,
        title: "Feed Up Department",
        department: "Feed Up Department",
        subtitle: feedUpMatch?.supervisor ? `Supervisor: ${feedUpMatch.supervisor}` : "Feed-up operations & seam joining",
        icon: "⚡",
        status: isFeedDone ? "completed" : hasStitchIssue ? "in_progress" : "pending",
        issueDate: feedDate,
        completeDate: feedCompDate,
        dwellDays: calculateDaysDiff(feedDate, feedCompDate),
        details: {
          supervisor: feedUpMatch?.supervisor || "Feed Up Sup",
          wipRemarks: feedUpMatch?.wipRemarks || ""
        }
      });
    }

    // 8. Daily Overlock
    if (overlockMatch || (indexMatch?.wipStatus && indexMatch.wipStatus.toLowerCase().includes("overlock"))) {
      const isOverlockDone = overlockMatch?.isComplete || hasStitchComplete;
      const overlockDate = overlockMatch?.date || indexMatch?.dateOfIssue || "Pending";
      const overlockCompDate = overlockMatch?.completeDate || (isOverlockDone ? "Completed" : "In Progress");
      milestones.push({
        id: "overlock",
        stageNumber: stageCounter++,
        title: "Daily Overlock",
        department: "Overlock Department",
        subtitle: overlockMatch?.supervisor ? `Supervisor: ${overlockMatch.supervisor}` : "Overlock stitching & edge trimming",
        icon: "➰",
        status: isOverlockDone ? "completed" : hasStitchIssue ? "in_progress" : "pending",
        issueDate: overlockDate,
        completeDate: overlockCompDate,
        dwellDays: calculateDaysDiff(overlockDate, overlockCompDate),
        details: {
          supervisor: overlockMatch?.supervisor || "Overlock Sup",
          aging: overlockMatch?.aging || "0",
          wipRemarks: overlockMatch?.wipRemarks || ""
        }
      });
    }

    // 9. Kaj Button & Secondary Work
    const isKajDone =
      kajMatch?.isComplete ||
      kajMatch?.status?.toLowerCase().includes("complete") ||
      kajMatch?.status?.toLowerCase().includes("done") ||
      !!issuesMatch?.pkgDate;
    const kajIssueDate = kajMatch?.date || indexMatch?.completedStatus || (hasStitchComplete ? "Issued" : "Pending");
    const kajCompDate = isKajDone ? (kajMatch?.completeDate || issuesMatch?.pkgDate || "Completed") : "In Progress";
    milestones.push({
      id: "kaj_button",
      stageNumber: stageCounter++,
      title: "Kaj Button & Secondary Work",
      department: "Kaj Button Department",
      subtitle: kajMatch?.supervisor ? `Supervisor: ${kajMatch.supervisor}` : "Button attachment & keyhole inspection",
      icon: "🔘",
      status: isKajDone ? "completed" : hasStitchComplete ? "in_progress" : "pending",
      issueDate: kajIssueDate,
      completeDate: kajCompDate,
      dwellDays: calculateDaysDiff(kajIssueDate, kajCompDate),
      details: {
        supervisor: kajMatch?.supervisor || "Kaj Sup",
        totalPcs: kajMatch?.totalPcs || bestTotalPcs,
        agingDays: kajMatch?.aging || "0",
        remarks: kajMatch?.remarks || kajMatch?.wipRemarks || "None"
      }
    });

    // 10. Washing Department
    if (washingMatch || (indexMatch?.wipStatus && indexMatch.wipStatus.toLowerCase().includes("wash"))) {
      const isWashDone = washingMatch?.isComplete || !!issuesMatch?.pkgDate;
      const washDate = washingMatch?.date || kajCompDate || "Pending";
      const washCompDate = washingMatch?.completeDate || (isWashDone ? "Completed" : "In Progress");
      milestones.push({
        id: "washing",
        stageNumber: stageCounter++,
        title: "Washing Department",
        department: "Washing Department",
        subtitle: washingMatch?.supervisor ? `Supervisor: ${washingMatch.supervisor}` : "Garment wash, softness & drying",
        icon: "🌊",
        status: isWashDone ? "completed" : isKajDone ? "in_progress" : "pending",
        issueDate: washDate,
        completeDate: washCompDate,
        dwellDays: calculateDaysDiff(washDate, washCompDate),
        details: {
          supervisor: washingMatch?.supervisor || "Washing Sup",
          wipRemarks: washingMatch?.wipRemarks || ""
        }
      });
    }

    // 11. Daily Folding
    if (foldingMatch || (indexMatch?.wipStatus && indexMatch.wipStatus.toLowerCase().includes("fold"))) {
      const isFoldDone = foldingMatch?.isComplete || !!issuesMatch?.pkgDate;
      const foldDate = foldingMatch?.date || kajCompDate || "Pending";
      const foldCompDate = foldingMatch?.completeDate || (isFoldDone ? "Completed" : "In Progress");
      milestones.push({
        id: "folding",
        stageNumber: stageCounter++,
        title: "Daily Folding",
        department: "Folding Department",
        subtitle: foldingMatch?.supervisor ? `Supervisor: ${foldingMatch.supervisor}` : "Folding, steam press & polybagging",
        icon: "📦",
        status: isFoldDone ? "completed" : isKajDone ? "in_progress" : "pending",
        issueDate: foldDate,
        completeDate: foldCompDate,
        dwellDays: calculateDaysDiff(foldDate, foldCompDate),
        details: {
          supervisor: foldingMatch?.supervisor || "Folding Sup",
          aging: foldingMatch?.aging || "0",
          wipRemarks: foldingMatch?.wipRemarks || ""
        }
      });
    }

    // 12. Packing & Finishing
    const hasPkgIssue = !!issuesMatch?.pkgDate;
    const hasPkgComplete =
      !!issuesMatch?.packingComplete &&
      issuesMatch.packingComplete !== "-" &&
      !issuesMatch.packingComplete.toLowerCase().includes("pending");
    const hasBarcode = !!barcodeMatch;
    const isFullyComplete = hasPkgComplete || hasBarcode;

    const pkgIssueDate = issuesMatch?.pkgDate || (isKajDone ? "Issued" : "Pending");
    const pkgCompDate = issuesMatch?.packingComplete || barcodeMatch?.barcodeDate || (isFullyComplete ? "Completed" : "In Progress");

    milestones.push({
      id: "packing",
      stageNumber: stageCounter++,
      title: "Packing & Finishing",
      department: "Packing Department",
      subtitle: issuesMatch?.pkgSupervisor ? `Supervisor: ${issuesMatch.pkgSupervisor}` : "Carton boxing & sticker allocation",
      icon: "🏷️",
      status: isFullyComplete ? "completed" : hasPkgIssue ? "in_progress" : "pending",
      issueDate: pkgIssueDate,
      completeDate: pkgCompDate,
      dwellDays: calculateDaysDiff(pkgIssueDate, pkgCompDate),
      details: {
        supervisor: issuesMatch?.pkgSupervisor || "Packing Sup",
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
        image: indexMatch?.image || "",
        totalDepartments: milestones.length,
        completedDepartments: completedCount
      },
      milestones
    };
  };

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
      try {
        const directData = await buildTimelineFromSheets(cleanLot);
        setTimelineData(directData);
        setActiveLot(cleanLot);
        setRecentLots((prev) => [cleanLot, ...prev.filter((l) => l !== cleanLot)].slice(0, 8));
      } catch (sheetErr) {
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

  const displayedMilestones = timelineData?.milestones.filter(m => {
    if (activeFilterTab === "completed") return m.status === "completed";
    if (activeFilterTab === "in_progress") return m.status === "in_progress";
    if (activeFilterTab === "pending") return m.status === "pending";
    return true;
  }) || [];

  return (
    <div className={`v-tracker-root ${isDarkMode ? "dark" : "light"}`}>
      <style>{`
        .v-tracker-root {
          min-height: 100vh;
          font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
          transition: background 0.3s ease, color 0.3s ease;
          padding: 24px 32px 80px;
          box-sizing: border-box;
          width: 100%;
        }

        .v-tracker-root.light {
          background: linear-gradient(180deg, #f8fafc 0%, #f1f5f9 100%);
          color: #0f172a;
        }

        .v-tracker-root.dark {
          background: radial-gradient(ellipse at top, #0d1527 0%, #060911 100%);
          color: #f8fafc;
        }

        .v-tracker-container {
          max-width: 1600px;
          margin: 0 auto;
          width: 100%;
        }

        /* TOP EXECUTIVE STRIP */
        .v-top-strip {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-bottom: 20px;
          flex-wrap: wrap;
          gap: 16px;
          padding-bottom: 16px;
          border-bottom: 1px solid rgba(226, 232, 240, 0.8);
        }
        .dark .v-top-strip {
          border-bottom-color: rgba(255, 255, 255, 0.08);
        }

        .v-brand-title {
          font-size: 24px;
          font-weight: 900;
          letter-spacing: -0.5px;
          color: #0f4c81;
          display: flex;
          align-items: center;
          gap: 10px;
          margin: 0;
        }
        .dark .v-brand-title { color: #38bdf8; }

        .v-brand-subtitle {
          font-size: 13px;
          font-weight: 600;
          color: #64748b;
          margin-left: 8px;
        }
        .dark .v-brand-subtitle { color: #94a3b8; }

        .v-actions-group {
          display: flex;
          align-items: center;
          gap: 10px;
        }

        .v-btn {
          padding: 9px 16px;
          border-radius: 10px;
          font-size: 13px;
          font-weight: 700;
          cursor: pointer;
          border: none;
          display: inline-flex;
          align-items: center;
          gap: 8px;
          transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
          box-shadow: 0 1px 2px rgba(0,0,0,0.05);
        }
        .v-btn:hover {
          transform: translateY(-1px);
        }
        .v-btn:active {
          transform: translateY(0);
        }

        .v-btn-secondary {
          background: #ffffff;
          color: #334155;
          border: 1px solid #cbd5e1;
        }
        .v-btn-secondary:hover {
          background: #f1f5f9;
          color: #0f172a;
          border-color: #94a3b8;
        }
        .dark .v-btn-secondary {
          background: rgba(255, 255, 255, 0.06);
          border-color: rgba(255, 255, 255, 0.12);
          color: #e2e8f0;
        }
        .dark .v-btn-secondary:hover {
          background: rgba(255, 255, 255, 0.12);
          border-color: rgba(255, 255, 255, 0.25);
        }

        .v-btn-primary {
          background: linear-gradient(135deg, #0f4c81 0%, #1e3a8a 100%);
          color: white;
          box-shadow: 0 4px 12px rgba(15, 76, 129, 0.25);
        }
        .v-btn-primary:hover {
          background: linear-gradient(135deg, #165b99 0%, #2546a8 100%);
          box-shadow: 0 6px 16px rgba(15, 76, 129, 0.35);
        }

        /* SEARCH HERO CARD */
        .v-search-box {
          border-radius: 16px;
          padding: 16px 24px;
          margin-bottom: 24px;
          display: flex;
          flex-direction: column;
          gap: 14px;
          transition: all 0.2s ease;
        }

        .light .v-search-box {
          background: #ffffff;
          border: 1px solid #e2e8f0;
          box-shadow: 0 4px 20px -4px rgba(0, 0, 0, 0.04);
        }

        .dark .v-search-box {
          background: #111827;
          border: 1px solid rgba(255, 255, 255, 0.08);
          box-shadow: 0 8px 24px rgba(0, 0, 0, 0.4);
        }

        .v-search-form {
          display: flex;
          gap: 12px;
          width: 100%;
        }

        .v-search-input-wrap {
          flex: 1;
          position: relative;
        }

        .v-search-input {
          width: 100%;
          padding: 13px 20px 13px 44px;
          border-radius: 12px;
          font-size: 15px;
          font-weight: 700;
          outline: none;
          box-sizing: border-box;
          border: 1.5px solid #cbd5e1;
          transition: all 0.2s ease;
        }
        .light .v-search-input { background: #f8fafc; color: #0f172a; }
        .light .v-search-input:focus { border-color: #0f4c81; background: #ffffff; box-shadow: 0 0 0 4px rgba(15, 76, 129, 0.1); }
        .dark .v-search-input { background: #0b0f19; border-color: rgba(255, 255, 255, 0.12); color: white; }
        .dark .v-search-input:focus { border-color: #38bdf8; box-shadow: 0 0 0 4px rgba(56, 189, 248, 0.15); }

        .v-search-icon {
          position: absolute;
          left: 16px;
          top: 50%;
          transform: translateY(-50%);
          font-size: 18px;
          opacity: 0.6;
        }

        .v-recent-row {
          display: flex;
          align-items: center;
          gap: 8px;
          flex-wrap: wrap;
        }

        .v-recent-chip {
          padding: 5px 12px;
          border-radius: 8px;
          font-size: 12px;
          font-weight: 700;
          border: 1px solid #e2e8f0;
          cursor: pointer;
          background: #f8fafc;
          color: #475569;
          transition: all 0.15s ease;
        }
        .v-recent-chip:hover {
          background: #e2e8f0;
          color: #0f172a;
          transform: translateY(-1px);
        }
        .v-recent-chip.active {
          background: #0f4c81;
          color: white;
          border-color: #0f4c81;
          box-shadow: 0 2px 6px rgba(15, 76, 129, 0.3);
        }
        .dark .v-recent-chip {
          background: rgba(255, 255, 255, 0.05);
          border-color: rgba(255, 255, 255, 0.08);
          color: #94a3b8;
        }
        .dark .v-recent-chip:hover {
          background: rgba(255, 255, 255, 0.12);
          color: #f8fafc;
        }
        .dark .v-recent-chip.active {
          background: #2563eb;
          color: white;
          border-color: #2563eb;
        }

        /* TOP EXECUTIVE KPI CARDS RIBBON */
        .v-kpi-ribbon {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
          gap: 16px;
          margin-bottom: 24px;
        }

        .v-kpi-card {
          border-radius: 14px;
          padding: 16px 20px;
          display: flex;
          align-items: center;
          gap: 14px;
          transition: all 0.2s ease;
        }
        .light .v-kpi-card {
          background: #ffffff;
          border: 1px solid #e2e8f0;
          box-shadow: 0 2px 10px rgba(0, 0, 0, 0.02);
        }
        .dark .v-kpi-card {
          background: #111827;
          border: 1px solid rgba(255, 255, 255, 0.08);
          box-shadow: 0 4px 14px rgba(0, 0, 0, 0.2);
        }
        .v-kpi-card:hover {
          transform: translateY(-2px);
          box-shadow: 0 6px 20px rgba(0, 0, 0, 0.06);
        }

        .v-kpi-icon-wrap {
          width: 46px;
          height: 46px;
          border-radius: 12px;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 22px;
          flex-shrink: 0;
        }

        .v-kpi-info {
          display: flex;
          flex-direction: column;
          min-width: 0;
        }

        .v-kpi-lbl {
          font-size: 11px;
          font-weight: 800;
          text-transform: uppercase;
          letter-spacing: 0.5px;
          color: #64748b;
          margin-bottom: 2px;
        }
        .dark .v-kpi-lbl { color: #94a3b8; }

        .v-kpi-val {
          font-size: 18px;
          font-weight: 900;
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
        }

        .v-kpi-sub {
          font-size: 11px;
          font-weight: 600;
          color: #94a3b8;
          margin-top: 1px;
        }

        /* TWO-COLUMN FULL-WIDTH MASTER GRID */
        .v-main-grid {
          display: grid;
          grid-template-columns: 360px 1fr;
          gap: 24px;
          align-items: start;
        }

        @media (max-width: 1024px) {
          .v-main-grid {
            grid-template-columns: 1fr;
          }
        }

        /* LEFT COLUMN: STICKY LOT OVERVIEW SIDEBAR */
        .v-sidebar-wrap {
          position: sticky;
          top: 24px;
          display: flex;
          flex-direction: column;
          gap: 16px;
        }

        .v-sidebar-card {
          border-radius: 16px;
          overflow: hidden;
          transition: all 0.2s ease;
        }
        .light .v-sidebar-card {
          background: #ffffff;
          border: 1px solid #e2e8f0;
          box-shadow: 0 4px 24px -4px rgba(0, 0, 0, 0.05);
        }
        .dark .v-sidebar-card {
          background: #111827;
          border: 1px solid rgba(255, 255, 255, 0.08);
          box-shadow: 0 8px 30px rgba(0, 0, 0, 0.4);
        }

        .v-sidebar-header {
          padding: 20px;
          background: linear-gradient(135deg, #0f4c81 0%, #1e3a8a 100%);
          color: white;
        }

        .v-sidebar-status-tag {
          display: inline-flex;
          align-items: center;
          gap: 6px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 10.5px;
          font-weight: 800;
          letter-spacing: 0.5px;
          margin-bottom: 8px;
        }
        .v-sidebar-status-tag.completed {
          background: #dcfce7;
          color: #15803d;
        }
        .v-sidebar-status-tag.in-progress {
          background: #fef3c7;
          color: #b45309;
        }

        .v-sidebar-lot-num {
          font-size: 32px;
          font-weight: 900;
          color: #ffffff;
          margin: 0 0 4px 0;
          letter-spacing: -0.5px;
          text-shadow: 0 2px 4px rgba(0,0,0,0.2);
        }

        .v-sidebar-stage-sub {
          font-size: 13.5px;
          font-weight: 700;
          color: rgba(255, 255, 255, 0.9);
          margin: 0;
        }

        .v-sidebar-body {
          padding: 20px;
          display: flex;
          flex-direction: column;
          gap: 16px;
        }

        /* SIDEBAR PROGRESS BAR */
        .v-progress-container {
          display: flex;
          flex-direction: column;
          gap: 6px;
        }

        .v-progress-header {
          display: flex;
          justify-content: space-between;
          font-size: 12.5px;
          font-weight: 800;
        }

        .v-progress-track {
          height: 10px;
          border-radius: 9999px;
          background: #e2e8f0;
          overflow: hidden;
        }
        .dark .v-progress-track { background: rgba(255, 255, 255, 0.1); }

        .v-progress-fill {
          height: 100%;
          border-radius: 9999px;
          background: linear-gradient(90deg, #10b981 0%, #059669 100%);
          transition: width 0.8s cubic-bezier(0.4, 0, 0.2, 1);
        }

        /* SPECS TABLE */
        .v-specs-table {
          display: flex;
          flex-direction: column;
          gap: 10px;
          border-top: 1px solid #f1f5f9;
          border-bottom: 1px solid #f1f5f9;
          padding: 14px 0;
        }
        .dark .v-specs-table { border-color: rgba(255, 255, 255, 0.08); }

        .v-spec-row {
          display: flex;
          justify-content: space-between;
          align-items: center;
          font-size: 13px;
        }

        .v-spec-lbl {
          color: #64748b;
          font-weight: 600;
          display: flex;
          align-items: center;
          gap: 6px;
        }
        .dark .v-spec-lbl { color: #94a3b8; }

        .v-spec-val {
          font-weight: 800;
          text-align: right;
        }

        /* STAGE QUICK-NAVIGATOR */
        .v-stage-nav-card {
          border-radius: 16px;
          padding: 16px 20px;
        }
        .light .v-stage-nav-card {
          background: #ffffff;
          border: 1px solid #e2e8f0;
        }
        .dark .v-stage-nav-card {
          background: #111827;
          border: 1px solid rgba(255, 255, 255, 0.08);
        }

        .v-stage-nav-title {
          font-size: 12px;
          font-weight: 800;
          text-transform: uppercase;
          letter-spacing: 0.5px;
          color: #64748b;
          margin: 0 0 12px 0;
        }
        .dark .v-stage-nav-title { color: #94a3b8; }

        .v-stage-nav-list {
          display: flex;
          flex-direction: column;
          gap: 6px;
          max-height: 280px;
          overflow-y: auto;
          padding-right: 4px;
        }

        .v-stage-nav-item {
          display: flex;
          align-items: center;
          justify-content: space-between;
          padding: 6px 10px;
          border-radius: 8px;
          font-size: 12px;
          font-weight: 700;
          cursor: pointer;
          transition: all 0.15s ease;
          border: 1px solid transparent;
        }
        .light .v-stage-nav-item { background: #f8fafc; color: #334155; }
        .light .v-stage-nav-item:hover { background: #e2e8f0; }
        .dark .v-stage-nav-item { background: rgba(255,255,255,0.04); color: #cbd5e1; }
        .dark .v-stage-nav-item:hover { background: rgba(255,255,255,0.08); }

        .v-stage-nav-item.completed { border-left: 3px solid #10b981; }
        .v-stage-nav-item.in_progress { border-left: 3px solid #f59e0b; background: rgba(245, 158, 11, 0.1); }
        .v-stage-nav-item.pending { border-left: 3px solid #cbd5e1; opacity: 0.7; }

        /* RIGHT COLUMN: PURE VERTICAL TIMELINE STREAM */
        .v-stream-container {
          display: flex;
          flex-direction: column;
          gap: 0;
          position: relative;
        }

        .v-filter-bar {
          display: flex;
          align-items: center;
          gap: 8px;
          margin-bottom: 20px;
          flex-wrap: wrap;
        }

        .v-tab-btn {
          padding: 7px 16px;
          border-radius: 10px;
          font-size: 12.5px;
          font-weight: 700;
          border: 1px solid #cbd5e1;
          background: #ffffff;
          color: #475569;
          cursor: pointer;
          transition: all 0.15s ease;
        }
        .v-tab-btn:hover {
          background: #f1f5f9;
          color: #0f172a;
        }
        .v-tab-btn.active {
          background: #0f4c81;
          color: white;
          border-color: #0f4c81;
          box-shadow: 0 2px 8px rgba(15, 76, 129, 0.25);
        }
        .dark .v-tab-btn {
          background: #111827;
          border-color: rgba(255, 255, 255, 0.1);
          color: #cbd5e1;
        }
        .dark .v-tab-btn:hover {
          background: rgba(255, 255, 255, 0.08);
          color: #ffffff;
        }
        .dark .v-tab-btn.active {
          background: #2563eb;
          color: white;
          border-color: #2563eb;
        }

        /* VERTICAL TIMELINE NODE ITEM */
        .v-node-item {
          display: flex;
          position: relative;
          gap: 20px;
          padding-bottom: 28px;
        }

        .v-node-item:last-child {
          padding-bottom: 0;
        }

        /* VERTICAL SPINE CONNECTOR LINE */
        .v-node-spine {
          position: absolute;
          left: 21px;
          top: 44px;
          bottom: 0;
          width: 4px;
          border-radius: 4px;
        }

        .v-node-spine.completed {
          background: linear-gradient(180deg, #10b981 0%, #059669 100%);
        }
        .v-node-spine.in-progress {
          background: linear-gradient(180deg, #f59e0b 0%, #cbd5e1 100%);
        }
        .v-node-spine.pending {
          background: #e2e8f0;
        }
        .dark .v-node-spine.pending {
          background: rgba(255, 255, 255, 0.1);
        }

        /* VERTICAL NODE MARKER */
        .v-node-marker {
          width: 44px;
          height: 44px;
          border-radius: 50%;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 18px;
          flex-shrink: 0;
          z-index: 2;
          font-weight: 800;
          transition: all 0.2s ease;
        }

        .v-node-marker.completed {
          background: linear-gradient(135deg, #10b981 0%, #059669 100%);
          color: white;
          border: 3px solid #ffffff;
          box-shadow: 0 0 16px rgba(16, 185, 129, 0.4);
        }
        .dark .v-node-marker.completed {
          border-color: #111827;
        }

        .v-node-marker.in-progress {
          background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%);
          color: white;
          border: 3px solid #ffffff;
          box-shadow: 0 0 20px rgba(245, 158, 11, 0.5);
          animation: markerPulse 2s infinite;
        }
        .dark .v-node-marker.in-progress {
          border-color: #111827;
        }

        @keyframes markerPulse {
          0% { transform: scale(1); }
          50% { transform: scale(1.08); box-shadow: 0 0 24px rgba(245, 158, 11, 0.7); }
          100% { transform: scale(1); }
        }

        .v-node-marker.pending {
          background: #f1f5f9;
          color: #94a3b8;
          border: 2px solid #cbd5e1;
        }
        .dark .v-node-marker.pending {
          background: #0b0f19;
          color: #64748b;
          border-color: rgba(255, 255, 255, 0.15);
        }

        /* VERTICAL NODE CARD */
        .v-node-card {
          flex: 1;
          border-radius: 16px;
          padding: 18px 22px;
          transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
        }

        .light .v-node-card {
          background: #ffffff;
          border: 1px solid #e2e8f0;
          box-shadow: 0 4px 16px rgba(0, 0, 0, 0.03);
        }
        .light .v-node-card:hover {
          transform: translateY(-2px);
          box-shadow: 0 8px 24px rgba(0, 0, 0, 0.06);
        }
        .light .v-node-card.in-progress {
          border-color: #f59e0b;
          background: #fffdf5;
          box-shadow: 0 4px 20px rgba(245, 158, 11, 0.12);
        }

        .dark .v-node-card {
          background: #111827;
          border: 1px solid rgba(255, 255, 255, 0.08);
          box-shadow: 0 4px 18px rgba(0, 0, 0, 0.3);
        }
        .dark .v-node-card:hover {
          transform: translateY(-2px);
          border-color: rgba(255, 255, 255, 0.18);
        }
        .dark .v-node-card.in-progress {
          border-color: #f59e0b;
          background: rgba(245, 158, 11, 0.07);
        }

        .v-card-top {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-bottom: 8px;
          flex-wrap: wrap;
          gap: 8px;
        }

        .v-stage-title-wrap {
          display: flex;
          align-items: center;
          gap: 10px;
        }

        .v-stage-num {
          font-size: 12px;
          font-weight: 800;
          color: #0f4c81;
          background: #eff6ff;
          padding: 3px 8px;
          border-radius: 6px;
        }
        .dark .v-stage-num {
          background: rgba(59, 130, 246, 0.2);
          color: #93c5fd;
        }

        .v-stage-name {
          font-size: 16px;
          font-weight: 800;
          margin: 0;
          letter-spacing: -0.3px;
        }
        .light .v-stage-name { color: #0f172a; }
        .dark .v-stage-name { color: #ffffff; }

        .v-stage-sub {
          font-size: 13px;
          color: #64748b;
          margin: 0 0 14px 0;
          font-weight: 500;
        }
        .dark .v-stage-sub { color: #94a3b8; }

        /* DUAL DATE TRACKER CAPSULE */
        .v-date-track {
          display: flex;
          align-items: center;
          gap: 12px;
          font-size: 13px;
          padding: 8px 14px;
          border-radius: 10px;
          margin-bottom: 12px;
          flex-wrap: wrap;
        }
        .light .v-date-track {
          background: #f8fafc;
          border: 1px solid #e2e8f0;
        }
        .dark .v-date-track {
          background: rgba(0, 0, 0, 0.25);
          border: 1px solid rgba(255, 255, 255, 0.08);
        }

        .v-date-seg {
          display: flex;
          align-items: center;
          gap: 6px;
        }

        .v-date-lbl {
          font-size: 11px;
          font-weight: 800;
          color: #64748b;
          text-transform: uppercase;
        }
        .dark .v-date-lbl { color: #94a3b8; }

        .v-date-val {
          font-weight: 800;
        }

        .v-dwell-chip {
          font-size: 11.5px;
          font-weight: 800;
          background: #dcfce7;
          color: #15803d;
          padding: 2px 8px;
          border-radius: 6px;
        }

        /* METADATA TAGS ROW */
        .v-meta-row {
          display: flex;
          gap: 8px;
          flex-wrap: wrap;
        }

        .v-meta-tag {
          font-size: 12px;
          font-weight: 700;
          padding: 4px 10px;
          border-radius: 6px;
        }
        .light .v-meta-tag {
          background: #f1f5f9;
          color: #334155;
          border: 1px solid #e2e8f0;
        }
        .dark .v-meta-tag {
          background: rgba(255, 255, 255, 0.06);
          color: #cbd5e1;
          border: 1px solid rgba(255, 255, 255, 0.1);
        }

        @media print {
          .no-print { display: none !important; }
          .v-tracker-root { padding: 0 !important; background: white !important; }
          .v-main-grid { display: block !important; }
          .v-sidebar-wrap { position: static !important; margin-bottom: 24px !important; }
        }
      `}</style>

      <div className="v-tracker-container">
        {/* TOP COMPACT NAVIGATION STRIP */}
        <div className="v-top-strip no-print">
          <div style={{ display: "flex", alignItems: "baseline", gap: "8px", flexWrap: "wrap" }}>
            <h1 className="v-brand-title">
              <span>🏭</span> FACTORY LOT LIFECYCLE
            </h1>
            <span className="v-brand-subtitle">
              Unified 12-Department Production Intelligence Stream
            </span>
          </div>

          <div className="v-actions-group">
            <button className="v-btn v-btn-secondary" onClick={() => history.push("/dashboard")}>
              ← Dashboard
            </button>
            <button className="v-btn v-btn-secondary" onClick={() => setIsDarkMode(!isDarkMode)}>
              {isDarkMode ? "☀️ Light" : "🌙 Dark"}
            </button>
            <button className="v-btn v-btn-secondary" onClick={() => fetchTimeline(activeLot, true)}>
              ↻ Refresh Live
            </button>
            <button className="v-btn v-btn-primary" onClick={() => window.print()}>
              🖨️ Export / Print
            </button>
          </div>
        </div>

        {/* SEARCH BAR */}
        <div className="v-search-box no-print">
          <form onSubmit={handleSearch} className="v-search-form">
            <div className="v-search-input-wrap">
              <span className="v-search-icon">🔍</span>
              <input
                type="text"
                className="v-search-input"
                placeholder="Search Lot Number (e.g. 11360, 76063, 65002)..."
                value={lotInput}
                onChange={(e) => setLotInput(e.target.value)}
              />
            </div>
            <button type="submit" className="v-btn v-btn-primary" disabled={loading}>
              {loading ? "Tracking..." : "Track Lot"}
            </button>
          </form>

          <div className="v-recent-row">
            <span style={{ fontSize: "11.5px", fontWeight: 800, color: "#64748b", textTransform: "uppercase" }}>
              Quick Lots:
            </span>
            {recentLots.map((lot) => (
              <button
                key={lot}
                type="button"
                className={`v-recent-chip ${activeLot === lot ? "active" : ""}`}
                onClick={() => handleSelectRecent(lot)}
              >
                Lot #{lot}
              </button>
            ))}
          </div>
        </div>

        {/* LOADING STATE */}
        {loading && (
          <div style={{ textAlign: "center", padding: "60px 0" }}>
            <div style={{ fontSize: "36px", animation: "markerPulse 1s infinite", display: "inline-block" }}>⚡</div>
            <p style={{ fontWeight: 800, fontSize: "16px", marginTop: "12px" }}>
              Syncing Lot #{lotInput} Real-time Data Across All 12 Departments...
            </p>
          </div>
        )}

        {/* ERROR STATE */}
        {!loading && error && (
          <div style={{ background: "#fef2f2", border: "1px solid #fecaca", borderRadius: "14px", padding: "24px", textAlign: "center", color: "#991b1b", marginBottom: "24px" }}>
            <p style={{ margin: "0 0 12px 0", fontWeight: 800, fontSize: "15px" }}>{error}</p>
            <button className="v-btn v-btn-primary" onClick={() => fetchTimeline("11360")}>
              Load Verified Sample Lot #11360
            </button>
          </div>
        )}

        {/* MAIN FULL-WIDTH TIMELINE DASHBOARD */}
        {!loading && timelineData && (
          <>
            {/* TOP EXECUTIVE KPI RIBBON */}
            <div className="v-kpi-ribbon no-print">
              <div className="v-kpi-card">
                <div className="v-kpi-icon-wrap" style={{ background: timelineData.summary.isFullyComplete ? "#dcfce7" : "#fef3c7", color: timelineData.summary.isFullyComplete ? "#15803d" : "#b45309" }}>
                  {timelineData.summary.isFullyComplete ? "✅" : "⚙️"}
                </div>
                <div className="v-kpi-info">
                  <span className="v-kpi-lbl">Current Stage</span>
                  <span className="v-kpi-val" style={{ color: timelineData.summary.isFullyComplete ? "#15803d" : "#b45309" }}>
                    {timelineData.summary.currentStage}
                  </span>
                  <span className="v-kpi-sub">
                    {timelineData.summary.isFullyComplete ? "All stages verified done" : "Active in production"}
                  </span>
                </div>
              </div>

              <div className="v-kpi-card">
                <div className="v-kpi-icon-wrap" style={{ background: "#e0f2fe", color: "#0284c7" }}>
                  📊
                </div>
                <div className="v-kpi-info">
                  <span className="v-kpi-lbl">Overall Progress</span>
                  <span className="v-kpi-val" style={{ color: "#0284c7" }}>
                    {timelineData.summary.progressPercent}%
                  </span>
                  <span className="v-kpi-sub">
                    {timelineData.summary.completedDepartments} of {timelineData.milestones.length} stages completed
                  </span>
                </div>
              </div>

              <div className="v-kpi-card">
                <div className="v-kpi-icon-wrap" style={{ background: "#ede9fe", color: "#7c3aed" }}>
                  👕
                </div>
                <div className="v-kpi-info">
                  <span className="v-kpi-lbl">Total Cut Quantity</span>
                  <span className="v-kpi-val" style={{ color: "#7c3aed" }}>
                    {timelineData.summary.totalPcs || "—"} Pcs
                  </span>
                  <span className="v-kpi-sub">
                    {timelineData.summary.garment || "Standard Garment"}
                  </span>
                </div>
              </div>

              <div className="v-kpi-card">
                <div className="v-kpi-icon-wrap" style={{ background: "#fce7f3", color: "#be185d" }}>
                  🏢
                </div>
                <div className="v-kpi-info">
                  <span className="v-kpi-lbl">Client / Party</span>
                  <span className="v-kpi-val" style={{ color: "#be185d" }}>
                    {timelineData.summary.party || "Direct Order"}
                  </span>
                  <span className="v-kpi-sub">
                    Style: {timelineData.summary.style || "—"}
                  </span>
                </div>
              </div>

              <div className="v-kpi-card">
                <div className="v-kpi-icon-wrap" style={{ background: "#fef9c3", color: "#a16207" }}>
                  ⏱️
                </div>
                <div className="v-kpi-info">
                  <span className="v-kpi-lbl">Lead Time</span>
                  <span className="v-kpi-val" style={{ color: "#a16207" }}>
                    {timelineData.summary.totalLeadTimeDays}
                  </span>
                  <span className="v-kpi-sub">Cross-factory sync</span>
                </div>
              </div>
            </div>

            {/* 2-COLUMN FULL-WIDTH LAYOUT */}
            <div className="v-main-grid">
              {/* LEFT SIDEBAR: STICKY LOT OVERVIEW & QUICK NAVIGATOR */}
              <div className="v-sidebar-wrap">
                <div className="v-sidebar-card">
                  <div className="v-sidebar-header">
                    <div className={`v-sidebar-status-tag ${timelineData.summary.isFullyComplete ? "completed" : "in-progress"}`}>
                      <span style={{ width: "6px", height: "6px", borderRadius: "50%", background: timelineData.summary.isFullyComplete ? "#15803d" : "#b45309" }} />
                      {timelineData.summary.isFullyComplete ? "COMPLETED" : "IN PRODUCTION"}
                    </div>

                    <h2 className="v-sidebar-lot-num">Lot #{timelineData.lotNumber}</h2>
                    <p className="v-sidebar-stage-sub">{timelineData.summary.currentStage}</p>
                  </div>

                  <div className="v-sidebar-body">
                    {/* PROGRESS BAR */}
                    <div className="v-progress-container">
                      <div className="v-progress-header">
                        <span>Production Progress</span>
                        <span>{timelineData.summary.progressPercent}%</span>
                      </div>
                      <div className="v-progress-track">
                        <div
                          className="v-progress-fill"
                          style={{
                            width: `${timelineData.summary.progressPercent}%`,
                            background: timelineData.summary.isFullyComplete
                              ? "linear-gradient(90deg, #10b981 0%, #059669 100%)"
                              : "linear-gradient(90deg, #0f4c81 0%, #38bdf8 100%)"
                          }}
                        />
                      </div>
                    </div>

                    {/* SPECS TABLE */}
                    <div className="v-specs-table">
                      <div className="v-spec-row">
                        <span className="v-spec-lbl">👕 Garment:</span>
                        <span className="v-spec-val">{timelineData.summary.garment || "—"}</span>
                      </div>
                      <div className="v-spec-row">
                        <span className="v-spec-lbl">🎨 Style:</span>
                        <span className="v-spec-val">{timelineData.summary.style || "—"}</span>
                      </div>
                      <div className="v-spec-row">
                        <span className="v-spec-lbl">🧵 Fabric:</span>
                        <span className="v-spec-val">{timelineData.summary.fabric || "—"}</span>
                      </div>
                      <div className="v-spec-row">
                        <span className="v-spec-lbl">🏢 Client / Party:</span>
                        <span className="v-spec-val" style={{ color: "#7c3aed" }}>{timelineData.summary.party || "—"}</span>
                      </div>
                      <div className="v-spec-row">
                        <span className="v-spec-lbl">🔢 Total Pieces:</span>
                        <span className="v-spec-val" style={{ color: "#0284c7" }}>{timelineData.summary.totalPcs}</span>
                      </div>
                      <div className="v-spec-row">
                        <span className="v-spec-lbl">✅ Stages Done:</span>
                        <span className="v-spec-val" style={{ color: "#15803d" }}>
                          {timelineData.summary.completedDepartments || 0} / {timelineData.milestones.length}
                        </span>
                      </div>
                      <div className="v-spec-row">
                        <span className="v-spec-lbl">⏳ Lead Time:</span>
                        <span className="v-spec-val">{timelineData.summary.totalLeadTimeDays}</span>
                      </div>
                    </div>

                    <div style={{ fontSize: "11px", color: "#64748b", textAlign: "center" }}>
                      Connected to 12 factory department spreadsheets
                    </div>
                  </div>
                </div>

                {/* STAGE QUICK NAVIGATOR */}
                <div className="v-stage-nav-card no-print">
                  <h4 className="v-stage-nav-title">Department Stage Index</h4>
                  <div className="v-stage-nav-list">
                    {timelineData.milestones.map((m) => (
                      <div
                        key={m.id}
                        className={`v-stage-nav-item ${m.status}`}
                        onClick={() => {
                          const el = document.getElementById(`node-stage-${m.id}`);
                          if (el) el.scrollIntoView({ behavior: "smooth", block: "center" });
                        }}
                      >
                        <span style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                          <span>{m.status === "completed" ? "✓" : m.icon}</span>
                          <span>{m.title}</span>
                        </span>
                        <span style={{ fontSize: "10.5px", opacity: 0.7 }}>
                          {m.status === "completed" ? "Done" : m.status === "in_progress" ? "Active" : "Pending"}
                        </span>
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              {/* RIGHT COLUMN: PURE VERTICAL TIMELINE STREAM */}
              <div className="v-stream-container">
                {/* FILTER TABS */}
                <div className="v-filter-bar no-print">
                  <button
                    className={`v-tab-btn ${activeFilterTab === "all" ? "active" : ""}`}
                    onClick={() => setActiveFilterTab("all")}
                  >
                    All Pipeline Stages ({timelineData.milestones.length})
                  </button>
                  <button
                    className={`v-tab-btn ${activeFilterTab === "completed" ? "active" : ""}`}
                    onClick={() => setActiveFilterTab("completed")}
                  >
                    Completed ({timelineData.milestones.filter(m => m.status === "completed").length})
                  </button>
                  <button
                    className={`v-tab-btn ${activeFilterTab === "in_progress" ? "active" : ""}`}
                    onClick={() => setActiveFilterTab("in_progress")}
                  >
                    In Progress ({timelineData.milestones.filter(m => m.status === "in_progress").length})
                  </button>
                  <button
                    className={`v-tab-btn ${activeFilterTab === "pending" ? "active" : ""}`}
                    onClick={() => setActiveFilterTab("pending")}
                  >
                    Upcoming ({timelineData.milestones.filter(m => m.status === "pending").length})
                  </button>
                </div>

                {/* VERTICAL STREAM NODES */}
                {displayedMilestones.map((step, index) => {
                  const isCompleted = step.status === "completed";
                  const isInProgress = step.status === "in_progress";
                  const isLast = index === displayedMilestones.length - 1;

                  return (
                    <div key={step.id} id={`node-stage-${step.id}`} className="v-node-item">
                      {/* SPINE CONNECTOR */}
                      {!isLast && (
                        <div
                          className={`v-node-spine ${
                            isCompleted ? "completed" : isInProgress ? "in-progress" : "pending"
                          }`}
                        />
                      )}

                      {/* NODE MARKER */}
                      <div
                        className={`v-node-marker ${
                          isCompleted ? "completed" : isInProgress ? "in-progress" : "pending"
                        }`}
                      >
                        {isCompleted ? "✓" : step.icon}
                      </div>

                      {/* VERTICAL CARD */}
                      <div className={`v-node-card ${isInProgress ? "in-progress" : ""}`}>
                        <div className="v-card-top">
                          <div className="v-stage-title-wrap">
                            <span className="v-stage-num">0{step.stageNumber}</span>
                            <h4 className="v-stage-name">{step.title}</h4>
                          </div>

                          <span
                            className={`v-sidebar-status-tag ${
                              isCompleted ? "completed" : isInProgress ? "in-progress" : "pending"
                            }`}
                            style={{ margin: 0, fontSize: "11px", padding: "3px 10px" }}
                          >
                            {isCompleted ? "COMPLETED" : isInProgress ? "IN PROGRESS" : "UPCOMING"}
                          </span>
                        </div>

                        <p className="v-stage-sub">{step.subtitle}</p>

                        {/* DUAL DATE TRACKER CAPSULE */}
                        <div className="v-date-track">
                          <div className="v-date-seg">
                            <span className="v-date-lbl">📤 ISSUE:</span>
                            <span className="v-date-val">{step.issueDate || "Pending"}</span>
                          </div>

                          <span style={{ opacity: 0.4, fontWeight: 900 }}>➔</span>

                          <div className="v-date-seg">
                            <span className="v-date-lbl">📥 DONE:</span>
                            <span
                              className="v-date-val"
                              style={{
                                color: isCompleted ? "#15803d" : isInProgress ? "#b45309" : undefined
                              }}
                            >
                              {step.completeDate || (isInProgress ? "In Progress" : "Pending")}
                            </span>
                          </div>

                          {step.dwellDays && (
                            <span className="v-dwell-chip">⏱️ {step.dwellDays}</span>
                          )}
                        </div>

                        {/* STAGE METADATA TAGS */}
                        <div className="v-meta-row">
                          {step.department && (
                            <span className="v-meta-tag">🏢 {step.department}</span>
                          )}

                          {step.details.supervisor && step.details.supervisor !== "N/A" && (
                            <span className="v-meta-tag">👤 {step.details.supervisor}</span>
                          )}

                          {step.details.cuttingQty && step.details.cuttingQty !== "N/A" && (
                            <span className="v-meta-tag">✂️ {step.details.cuttingQty} Pcs</span>
                          )}

                          {step.details.challan && (
                            <span className="v-meta-tag">📄 {step.details.challan}</span>
                          )}

                          {step.details.wipRemarks && step.details.wipRemarks !== "In Progress" && (
                            <span className="v-meta-tag">📝 {step.details.wipRemarks}</span>
                          )}

                          {step.details.aging && step.details.aging !== "0" && (
                            <span className="v-meta-tag">⏳ {step.details.aging} Days Aging</span>
                          )}

                          {step.details.agingDays && step.details.agingDays !== "0" && (
                            <span className="v-meta-tag">⏳ {step.details.agingDays} Days Aging</span>
                          )}
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
