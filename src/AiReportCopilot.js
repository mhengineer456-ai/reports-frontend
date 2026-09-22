// src/AiReportCopilot.js
import React, { useState, useEffect, useRef } from "react";
import { useHistory } from "react-router-dom";
import { BACKEND_API_BASE_URL, SPREADSHEET_IDS, SHEET_NAMES, fetchSheetDataFromBackend } from "./config";

/**
 * Factory Suite Pro — Standalone Global AI Assistant Pop-up Window
 * Handles Hinglish, Hindi, and English natural language production queries.
 * Examples:
 *  - "cutting pr ss k lot kitna pending h"
 *  - "embroidery me kitne lot pending hai"
 *  - "lot 2045 ka status kya h"
 *  - "5 din se jyada delay wale lots"
 *  - "stitching me highest pending quantity"
 */

export default function AiReportCopilot({
  isOpen: externalIsOpen,
  onClose: externalOnClose,
  activeStage = "all",
  allLotsData: passedLotsData = null,
  onApplyFilter = null,
  onSelectStage = null
}) {
  const history = useHistory();

  // Internal open state if used as floating widget
  const [internalIsOpen, setInternalIsOpen] = useState(false);
  const isOpen = externalIsOpen !== undefined ? externalIsOpen : internalIsOpen;
  const onClose = externalOnClose || (() => setInternalIsOpen(false));

  const [query, setQuery] = useState("");
  const [selectedDept, setSelectedDept] = useState("all");
  const [isListening, setIsListening] = useState(false);
  const [loading, setLoading] = useState(false);
  const [chatHistory, setChatHistory] = useState([]);
  const [copiedIndex, setCopiedIndex] = useState(null);
  const [fetchedLots, setFetchedLots] = useState([]);
  const [isFetchingData, setIsFetchingData] = useState(false);

  // Speech Synthesis (Text-to-Speech) State
  const [isSpeaking, setIsSpeaking] = useState(false);
  const [currentSpeakingId, setCurrentSpeakingId] = useState(null);
  const [autoSpeak, setAutoSpeak] = useState(true);

  // Gemini API Key (Optional custom key)
  const [geminiKey, setGeminiKey] = useState(localStorage.getItem("FACTORY_AI_GEMINI_KEY") || "");
  const [showKeyModal, setShowKeyModal] = useState(false);

  const inputRef = useRef(null);
  const chatEndRef = useRef(null);

  // Stop speech synthesis helper
  const stopSpeech = () => {
    if (window.speechSynthesis) {
      window.speechSynthesis.cancel();
    }
    setIsSpeaking(false);
    setCurrentSpeakingId(null);
  };

  // Speak response text using browser Web Speech Synthesis
  const speakText = (rawText, messageId = null) => {
    if (!window.speechSynthesis) {
      return;
    }

    // Toggle stop if already speaking this message
    if (isSpeaking && currentSpeakingId === messageId) {
      stopSpeech();
      return;
    }

    stopSpeech();

    // Clean text for natural speech (remove markdown symbols, emojis, bullets, LaTeX math)
    let speechString = String(rawText || "")
      .replace(/###\s*/g, "")
      .replace(/\*\*/g, "")
      .replace(/\*/g, "")
      .replace(/`[^`]+`/g, "")
      .replace(/\$\\text\{([^}]+)\}\$/g, "$1")
      .replace(/\\frac\{([^}]+)\}\{([^}]+)\}/g, "$1 divided by $2")
      .replace(/\\times/g, " multiply ")
      .replace(/\$+/g, "")
      .replace(/•\s*/g, ". ")
      .replace(/[📊💡🚨⏱️🏢📦✂️🧵🖨️📋🪡👉⚡📍🏷️⏳✓🤖👋😊]/g, "")
      .replace(/\n+/g, ". ")
      .replace(/\.+/g, ".")
      .trim();

    if (!speechString) return;

    try {
      const utterance = new SpeechSynthesisUtterance(speechString);
      utterance.rate = 0.95; // Natural clear pace
      utterance.pitch = 1.0;

      // Find Hindi or Indian English voice if available
      const voices = window.speechSynthesis.getVoices();
      const preferredVoice = voices.find(v =>
        v.lang.includes("hi") ||
        v.lang.includes("HI") ||
        v.lang.includes("en-IN") ||
        v.name.includes("India") ||
        v.name.includes("Hindi")
      );
      if (preferredVoice) {
        utterance.voice = preferredVoice;
      }

      utterance.onstart = () => {
        setIsSpeaking(true);
        setCurrentSpeakingId(messageId);
      };
      utterance.onend = () => {
        setIsSpeaking(false);
        setCurrentSpeakingId(null);
      };
      utterance.onerror = () => {
        setIsSpeaking(false);
        setCurrentSpeakingId(null);
      };

      window.speechSynthesis.speak(utterance);
    } catch (err) {
      console.warn("TTS Error:", err);
      setIsSpeaking(false);
      setCurrentSpeakingId(null);
    }
  };

  // Stop speech when window closes
  useEffect(() => {
    if (!isOpen) {
      stopSpeech();
    }
  }, [isOpen]);

  // Cleanup speech on unmount
  useEffect(() => {
    return () => {
      stopSpeech();
    };
  }, []);

  // Global event listener to open AI pop-up from anywhere in the app
  useEffect(() => {
    const handleGlobalOpen = (e) => {
      setInternalIsOpen(true);
      if (e?.detail?.query) {
        setQuery(e.detail.query);
        setTimeout(() => handleExecuteQuery(e.detail.query), 200);
      }
    };
    window.addEventListener("open-factory-ai", handleGlobalOpen);
    return () => window.removeEventListener("open-factory-ai", handleGlobalOpen);
  }, []);

  // Auto-focus input when opened
  useEffect(() => {
    if (isOpen) {
      setTimeout(() => inputRef.current?.focus(), 150);
      // If no lots passed from parent report, fetch production lots for global intelligence
      if (!passedLotsData && fetchedLots.length === 0 && !isFetchingData) {
        loadGlobalProductionLots();
      }
    }
  }, [isOpen]);

  // Scroll to bottom of chat
  useEffect(() => {
    chatEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [chatHistory, loading]);

  // Load global production intelligence across ALL linked sheets in the factory system
  const loadGlobalProductionLots = async () => {
    setIsFetchingData(true);
    try {
      // Parallel fetch across all factory department spreadsheets & sub-operation cards
      const [
        indexRes,
        jobRes,
        issuesRes,
        holdRes,
        stitchRes,
        rawpackRes,
        kajRes,
        overlockRes,
        feedUpRes,
        washRes,
        foldRes,
        elasticRes,
        barcodeRes
      ] = await Promise.allSettled([
        fetchSheetDataFromBackend(SPREADSHEET_IDS.MAIN, `${SHEET_NAMES.INDEX || 'Index'}!A1:ZZZ`),
        fetchSheetDataFromBackend(SPREADSHEET_IDS.JOBORDER, `${SHEET_NAMES.JOB_ORDER || 'JobOrder'}!A:AZ`),
        fetchSheetDataFromBackend(SPREADSHEET_IDS.ISSUES, `${SHEET_NAMES.ISSUES || 'Issues'}!A:AZ`),
        fetchSheetDataFromBackend(SPREADSHEET_IDS.HOLD_LOTS, `Hold_Lots!A:AZ`),
        fetchSheetDataFromBackend(SPREADSHEET_IDS.DAILY_STITCHING, `A:AZ`),
        fetchSheetDataFromBackend(SPREADSHEET_IDS.RAWPACK, `${SHEET_NAMES.RAWPACK || 'RAWPACK'}!A:AZ`),
        fetchSheetDataFromBackend(SPREADSHEET_IDS.WORKING_UPDATES, `KajButton!B:O`),
        fetchSheetDataFromBackend(SPREADSHEET_IDS.WORKING_UPDATES, `Overlock!B:O`),
        fetchSheetDataFromBackend(SPREADSHEET_IDS.WORKING_UPDATES, `FeedUp!B:O`),
        fetchSheetDataFromBackend(SPREADSHEET_IDS.WORKING_UPDATES, `Washing!B:O`),
        fetchSheetDataFromBackend(SPREADSHEET_IDS.WORKING_UPDATES, `Folding!B:O`),
        fetchSheetDataFromBackend(SPREADSHEET_IDS.WORKING_UPDATES, `Elastic!B:O`),
        fetchSheetDataFromBackend(SPREADSHEET_IDS.BARCODE, `${SHEET_NAMES.BARCODE || 'LotBarcodeData'}!A:Z`)
      ]);

      const normalize = (s) => String(s || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');

      const getVal = (row, headerMap, keys) => {
        for (const key of keys) {
          const idx = headerMap[normalize(key)];
          if (idx !== undefined && row[idx] !== undefined && String(row[idx]).trim() !== '') {
            return String(row[idx]).trim();
          }
        }
        return '';
      };

      const buildHeaderMap = (headersRow) => {
        const map = {};
        if (Array.isArray(headersRow)) {
          headersRow.forEach((h, idx) => {
            map[normalize(h)] = idx;
          });
        }
        return map;
      };

      const lotMap = new Map();

      // Helper to parse dates uniformly like OverallCuttingtoPacking.js
      const parseDateVal = (dStr) => {
        if (!dStr) return null;
        if (dStr instanceof Date) return isNaN(dStr.getTime()) ? null : dStr;
        const s = String(dStr).trim();
        if (!s || s === '-' || s.toLowerCase() === 'invalid date') return null;
        if (s.includes('/')) {
          const parts = s.split(' ')[0].split('/');
          if (parts.length === 3) {
            const m = parseInt(parts[0], 10) - 1;
            const d = parseInt(parts[1], 10);
            const y = parseInt(parts[2], 10);
            return new Date(y, m, d);
          }
        }
        if (s.includes('-')) {
          const parts = s.split(' ')[0].split('-');
          if (parts.length === 3) {
            const y = parseInt(parts[0], 10);
            const m = parseInt(parts[1], 10) - 1;
            const d = parseInt(parts[2], 10);
            return new Date(y, m, d);
          }
        }
        const parsed = new Date(s);
        return isNaN(parsed.getTime()) ? null : parsed;
      };

      const calcDaysDiff = (startStr, endStr) => {
        if (!startStr || !endStr) return '';
        const s = parseDateVal(startStr);
        const e = parseDateVal(endStr);
        if (!s || !e) return '';
        const diff = Math.round((e.getTime() - s.getTime()) / (1000 * 60 * 60 * 24));
        return diff >= 0 ? `${diff}d` : '';
      };

      // Step 1: Parse Master JobOrder Sheet (Master Registry)
      if (jobRes.status === 'fulfilled' && jobRes.value?.ok && Array.isArray(jobRes.value.values) && jobRes.value.values.length > 1) {
        const rows = jobRes.value.values;
        const headerMap = buildHeaderMap(rows[0]);

        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length === 0) continue;

          const lotNo = getVal(row, headerMap, ["Lot Number", "Lot No", "Lot", "Lot#"]) || (row[0] ? String(row[0]).trim() : '');
          if (!lotNo) continue;

          const cleanLotKey = String(lotNo).replace(/\s*\(cancel\)/gi, '').trim().toLowerCase();
          const rawQty = getVal(row, headerMap, ["Quantity", "Qty", "Pieces", "Pcs", "Total Qty", "Order Qty"]);
          const qty = parseFloat(String(rawQty).replace(/,/g, '')) || 0;
          const party = getVal(row, headerMap, ["Party Name", "Party", "Buyer", "Customer", "Client"]) || "Mohit Hosiery";
          const brand = getVal(row, headerMap, ["Brand", "BRAND"]) || "";
          const garmentType = getVal(row, headerMap, ["Garment Type", "Item", "Garment", "Product"]) || "";
          const style = getVal(row, headerMap, ["Style", "Style Name", "Description"]) || "";
          const fabric = getVal(row, headerMap, ["Fabric", "Fabric Type", "GSM", "Cloth", "Fabric Quality"]) || "";
          const season = getVal(row, headerMap, ["SEASON", "Season"]) || "";
          const section = getVal(row, headerMap, ["SECTION", "Section", "M/W/K"]) || "";
          const directStitching = getVal(row, headerMap, ["DIRECT STITCHING", "Direct Stitching", "D.Stitching"]) || "";
          const status = getVal(row, headerMap, ["Status", "Order Status", "Job Status"]) || "In Progress";
          const jobDate = getVal(row, headerMap, ["Date", "Job Date", "Order Date", "Creation Date"]);
          const embRequired = getVal(row, headerMap, ["Emb", "Embroidery"]);
          const printRequired = getVal(row, headerMap, ["Printing", "Print"]);

          lotMap.set(cleanLotKey, {
            id: `job-${i}`,
            lotNo: String(lotNo).trim(),
            party,
            brand,
            item: style ? `${garmentType} (${style})` : garmentType || "Garment",
            style,
            garmentType,
            fabric,
            season,
            section,
            directStitching,
            qty,
            cuttingQty: 0,
            stitchIssueQty: 0,
            stage: "Job Order Created",
            status,
            jobDate,
            cutDate: "",
            embIssueDate: "",
            embCompDate: "",
            stitchDate: "",
            stitchSup: "",
            wipStitch: "",
            compStitch: "",
            pkgSup: "",
            pkgDate: "",
            wipPkg: "",
            pkgComp: "",
            embRequired,
            printRequired,
            pendingDays: 2,
            isHold: false,
            holdReason: '',
            holdDept: '',
            issuedTo: '',
            issueDate: '',
            stitchingOutput: 0,
            packingStatus: '',
            kajDetails: null,
            overlockDetails: null,
            feedUpDetails: null,
            washingDetails: null,
            foldingDetails: null,
            elasticDetails: null,
            barcodeDetails: null,
            sourceSheets: ["JobOrder"]
          });
        }
      }

      // Step 2: Overlay Master Index Sheet (Cutting & Embellishment)
      if (indexRes.status === 'fulfilled' && indexRes.value?.ok && Array.isArray(indexRes.value.values) && indexRes.value.values.length > 1) {
        const rows = indexRes.value.values;
        const headerMap = buildHeaderMap(rows[0]);

        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length === 0) continue;

          const lotNo = getVal(row, headerMap, ["Lot Number", "Lot No", "Lot", "Lot#"]) || (row[0] ? String(row[0]).trim() : '');
          if (!lotNo) continue;

          const cleanLotKey = String(lotNo).replace(/\s*\(cancel\)/gi, '').trim().toLowerCase();
          const rawCuttingQty = getVal(row, headerMap, ["Cutting Qty", "Cutting Quantity", "Cut Qty", "Total Pieces", "Quantity", "Qty", "Pieces", "Pcs"]);
          const cuttingQty = parseFloat(String(rawCuttingQty).replace(/,/g, '')) || 0;
          const rawStitchIssueQty = getVal(row, headerMap, ["Stitching Issue Qty", "Stitch Issue Qty", "Issue Qty"]);
          const stitchIssueQty = parseFloat(String(rawStitchIssueQty).replace(/,/g, '')) || 0;

          const party = getVal(row, headerMap, ["PARTY NAME", "Party Name", "Party"]);
          const brand = getVal(row, headerMap, ["BRAND", "Brand"]);
          const garmentType = getVal(row, headerMap, ["Garment Type", "Item", "Garment"]);
          const style = getVal(row, headerMap, ["Style", "Style Name"]);
          const fabric = getVal(row, headerMap, ["Fabric", "Cloth"]);
          const sizes = getVal(row, headerMap, ["Sizes", "Size", "Size Ratio"]);
          const shades = getVal(row, headerMap, ["Shades", "Shade", "Colors"]);
          const supervisor = getVal(row, headerMap, ["Supervisor", "Cutting Master", "Stitching Supervisor"]);
          const season = getVal(row, headerMap, ["SEASON", "Season"]);
          const directStitching = getVal(row, headerMap, ["DIRECT STITCHING", "Direct Stitching"]);
          const challanHistoryRaw = getVal(row, headerMap, ["CHALLAN HISTORY", "Challan History"]);
          const wipStatusRaw = getVal(row, headerMap, ["WIP Status", "WIP"]);
          const completedStatusRaw = getVal(row, headerMap, ["Completed Status", "Complete Status", "Completed"]);
          const priority = getVal(row, headerMap, ["Prioirty", "Priority"]);
          const cutDate = getVal(row, headerMap, ["Saved At", "Cut Date", "Date"]);
          const stitchDate = getVal(row, headerMap, ["Date of Issue", "Stitching Issue", "Issue Date"]);

          // Parse JSON challan history
          let challanInfo = "";
          let embIssueDate = "";
          let embCompDate = "";
          let embCompleted = false;
          if (challanHistoryRaw && challanHistoryRaw.startsWith("[")) {
            try {
              const chList = JSON.parse(challanHistoryRaw);
              if (Array.isArray(chList) && chList.length > 0) {
                const latestCh = chList[chList.length - 1];
                challanInfo = `${latestCh.number || 'Challan'} (${latestCh.totalQty || cuttingQty} Pcs)`;
                embIssueDate = latestCh.date || '';
                embCompDate = latestCh.receivedDate || latestCh.embUpdatedAt || '';
                embCompleted = !!latestCh.embCompleted;
              }
            } catch (e) {}
          }

          let isStitchComplete = false;
          let compStitchDate = "";
          if (completedStatusRaw && completedStatusRaw !== '-' && completedStatusRaw.toLowerCase() !== 'invalid date') {
            if (completedStatusRaw.startsWith("[")) {
              try {
                const cList = JSON.parse(completedStatusRaw);
                if (Array.isArray(cList) && cList.length > 0) {
                  isStitchComplete = true;
                  compStitchDate = cList[cList.length - 1].timestamp || cList[cList.length - 1].date || '';
                }
              } catch (e) {}
            } else if (completedStatusRaw.toLowerCase().includes("complete") || completedStatusRaw.toLowerCase().includes("done")) {
              isStitchComplete = true;
              compStitchDate = completedStatusRaw;
            }
          }

          let existing = lotMap.get(cleanLotKey);
          if (!existing) {
            existing = {
              id: `index-${i}`,
              lotNo: String(lotNo).trim(),
              party: party || "Mohit Hosiery",
              brand: brand || "",
              item: style ? `${garmentType} (${style})` : garmentType || "Garment",
              style: style || "",
              garmentType: garmentType || "",
              fabric: fabric || "",
              season: season || "",
              section: "",
              directStitching: directStitching || "",
              qty: cuttingQty,
              cuttingQty,
              stitchIssueQty,
              stage: "Cutting Done",
              status: "In Progress",
              jobDate: "",
              cutDate,
              embIssueDate,
              embCompDate,
              stitchDate,
              stitchSup: supervisor,
              wipStitch: wipStatusRaw,
              compStitch: compStitchDate,
              pkgSup: "",
              pkgDate: "",
              wipPkg: "",
              pkgComp: "",
              pendingDays: 2,
              isHold: false,
              holdReason: '',
              holdDept: '',
              issuedTo: '',
              issueDate: '',
              stitchingOutput: 0,
              packingStatus: '',
              kajDetails: null,
              overlockDetails: null,
              feedUpDetails: null,
              washingDetails: null,
              foldingDetails: null,
              elasticDetails: null,
              barcodeDetails: null,
              sourceSheets: []
            };
            lotMap.set(cleanLotKey, existing);
          }

          // Merge Index fields
          if (cuttingQty > 0) {
            existing.cuttingQty = cuttingQty;
            existing.qty = cuttingQty;
          }
          if (stitchIssueQty > 0) existing.stitchIssueQty = stitchIssueQty;
          if (cutDate) existing.cutDate = cutDate;
          if (stitchDate) existing.stitchDate = stitchDate;
          if (supervisor) existing.stitchSup = supervisor;
          if (sizes) existing.sizes = sizes;
          if (shades) existing.shades = shades;
          if (priority) existing.priority = priority;
          if (party && !existing.party) existing.party = party;
          if (brand && !existing.brand) existing.brand = brand;
          if (fabric && !existing.fabric) existing.fabric = fabric;
          if (challanInfo) {
            existing.challanInfo = challanInfo;
            existing.embIssueDate = embIssueDate;
            existing.embCompDate = embCompDate;
            existing.embCompleted = embCompleted;
          }
          if (wipStatusRaw) existing.wipStitch = wipStatusRaw;
          if (isStitchComplete) existing.compStitch = compStitchDate || "Completed";

          // Compute active stage
          if (isStitchComplete) {
            existing.stage = "Stitching Done (Finishing / Packing Queue)";
          } else if (stitchIssueQty > 0 || stitchDate) {
            existing.stage = "Stitching Line Active";
          } else if (challanInfo) {
            existing.stage = embCompleted ? "EMB Done (Pending Stitching Issue)" : "At Embroidery / Printing";
          } else if (directStitching && directStitching.toLowerCase() === 'yes') {
            existing.stage = "Direct Stitching Queue";
          } else if (cutDate) {
            existing.stage = "Cutting Done";
          }

          existing.sourceSheets.push("Index");
        }
      }

      // Step 3: Overlay Issues & RAWPACK & Barcode Sheets (Packing & Finishing)
      if (issuesRes.status === 'fulfilled' && issuesRes.value?.ok && Array.isArray(issuesRes.value.values) && issuesRes.value.values.length > 1) {
        const rows = issuesRes.value.values;
        const headerMap = buildHeaderMap(rows[0]);
        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length === 0) continue;
          const lotNo = getVal(row, headerMap, ["Lot Number", "Lot No", "Lot"]);
          if (!lotNo) continue;
          const cleanKey = String(lotNo).replace(/\s*\(cancel\)/gi, '').trim().toLowerCase();
          const pkgDate = getVal(row, headerMap, ["Packing Date", "Date of Issue", "Date", "Timestamp"]);
          const pkgSup = getVal(row, headerMap, ["Packing Supervisor", "Supervisor"]);
          const wipPkg = getVal(row, headerMap, ["WIP Packing", "Remarks", "Status"]);
          const pkgComp = getVal(row, headerMap, ["Packing Complete", "Packing Completed", "Complete Date"]);

          if (lotMap.has(cleanKey)) {
            const existing = lotMap.get(cleanKey);
            if (pkgDate) existing.pkgDate = pkgDate;
            if (pkgSup) existing.pkgSup = pkgSup;
            if (wipPkg) existing.wipPkg = wipPkg;
            if (pkgComp && pkgComp !== '-') {
              existing.pkgComp = pkgComp;
              existing.stage = "Completed & Boxed";
              existing.status = "Completed";
            } else if (pkgDate) {
              existing.stage = "Packing Floor Active";
            }
            existing.sourceSheets.push("Issues");
          }
        }
      }

      // Step 3: Parse Sub-operations from Working_Updates (KajButton, Overlock, FeedUp, Washing, Folding, Elastic)
      const parseSubDept = (res, deptName, fieldKey) => {
        if (res.status === 'fulfilled' && res.value?.ok && Array.isArray(res.value.values) && res.value.values.length > 1) {
          const rows = res.value.values;
          const headerMap = buildHeaderMap(rows[0]);
          for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;
            const lotNo = getVal(row, headerMap, ["Lot Number", "Lot No", "Lot", "Lot#"]);
            if (!lotNo) continue;
            const cleanKey = String(lotNo).trim().toLowerCase();
            const date = getVal(row, headerMap, ["Date", "Issue Date", "Saved At"]);
            const supervisor = getVal(row, headerMap, ["Supervisor", "Operator", "Master"]);
            const pcs = getVal(row, headerMap, ["Pcs", "Pieces", "Total Pcs", "Quantity", "Output"]);
            const completeRaw = getVal(row, headerMap, ["Completed Status", "Complete Status", "Status", "Done"]);
            const isComplete = completeRaw.toLowerCase().includes("complete") || completeRaw.toLowerCase().includes("done");

            const info = { dept: deptName, date, supervisor, pcs, isComplete };
            if (lotMap.has(cleanKey)) {
              const existing = lotMap.get(cleanKey);
              existing[fieldKey] = info;
              existing.sourceSheets.push(deptName);
            }
          }
        }
      };

      parseSubDept(kajRes, "KajButton", "kajDetails");
      parseSubDept(overlockRes, "Overlock", "overlockDetails");
      parseSubDept(feedUpRes, "FeedUp", "feedUpDetails");
      parseSubDept(washRes, "Washing", "washingDetails");
      parseSubDept(foldRes, "Folding", "foldingDetails");
      parseSubDept(elasticRes, "Elastic", "elasticDetails");

      // Step 4: Parse Barcode Carton Scans
      if (barcodeRes.status === 'fulfilled' && barcodeRes.value?.ok && Array.isArray(barcodeRes.value.values) && barcodeRes.value.values.length > 1) {
        const rows = barcodeRes.value.values;
        const headerMap = buildHeaderMap(rows[0]);
        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length === 0) continue;
          const lotNo = getVal(row, headerMap, ["Lot Number", "Lot No", "Lot"]);
          if (!lotNo) continue;
          const cleanKey = String(lotNo).trim().toLowerCase();
          const cartonCount = getVal(row, headerMap, ["Carton", "Cartons", "Total Cartons", "Box Count"]);
          const scanDate = getVal(row, headerMap, ["Date", "Timestamp", "Scan Date"]);

          if (lotMap.has(cleanKey)) {
            const existing = lotMap.get(cleanKey);
            existing.barcodeDetails = { cartonCount, scanDate };
            existing.sourceSheets.push("LotBarcodeData");
          }
        }
      }

      // Step 5: Overlay Daily Stitching Sheet
      if (stitchRes.status === 'fulfilled' && stitchRes.value?.ok && Array.isArray(stitchRes.value.values) && stitchRes.value.values.length > 1) {
        const rows = stitchRes.value.values;
        const headerMap = buildHeaderMap(rows[0]);

        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length === 0) continue;
          const lotNo = getVal(row, headerMap, ["Lot Number", "Lot No", "Lot", "Lot#"]);
          if (!lotNo) continue;

          const cleanKey = String(lotNo).trim().toLowerCase();
          const stitchQty = parseFloat(getVal(row, headerMap, ["Stitching Output", "Stitched Pcs", "Output", "Total", "Done Qty"])) || 0;
          const line = getVal(row, headerMap, ["Line", "Line No", "Supervisor", "Operator"]);

          if (lotMap.has(cleanKey)) {
            const existing = lotMap.get(cleanKey);
            existing.stitchingOutput = (existing.stitchingOutput || 0) + stitchQty;
            if (line) existing.stitchingLine = line;
            if (!existing.stage || existing.stage.includes("Cutting") || existing.stage.includes("Done")) {
              existing.stage = "Stitching In-Progress";
            }
            existing.sourceSheets.push("DailyStitching");
          } else {
            lotMap.set(cleanKey, {
              id: `stitch-${i}`,
              lotNo: String(lotNo).trim(),
              party: getVal(row, headerMap, ["Party Name", "Party"]) || "Standard",
              item: getVal(row, headerMap, ["Item", "Style", "Garment"]) || "",
              fabric: "",
              qty: parseFloat(getVal(row, headerMap, ["Quantity", "Qty", "Total Qty"])) || stitchQty,
              stage: "Stitching WIP",
              status: "In Progress",
              dateStr: "",
              pendingDays: 2,
              isHold: false,
              holdReason: '',
              holdDept: '',
              issuedTo: '',
              issueDate: '',
              stitchingOutput: stitchQty,
              stitchingLine: line,
              packingStatus: '',
              sourceSheets: ["DailyStitching"]
            });
          }
        }
      }

      // Step 6: Overlay Hold Lots Sheet
      if (holdRes.status === 'fulfilled' && holdRes.value?.ok && Array.isArray(holdRes.value.values) && holdRes.value.values.length > 1) {
        const rows = holdRes.value.values;
        const headerMap = buildHeaderMap(rows[0]);

        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length === 0) continue;
          const lotNo = getVal(row, headerMap, ["Lot Number", "Lot No", "Lot", "Hold Lot"]);
          if (!lotNo) continue;

          const cleanKey = String(lotNo).trim().toLowerCase();
          const reason = getVal(row, headerMap, ["Reason", "Hold Reason", "Remarks", "Issue"]);
          const dept = getVal(row, headerMap, ["Department", "Dept", "Section", "Stage"]);

          if (lotMap.has(cleanKey)) {
            const existing = lotMap.get(cleanKey);
            existing.isHold = true;
            existing.holdReason = reason || "Hold by Department";
            existing.holdDept = dept || existing.stage;
            existing.sourceSheets.push("HoldLots");
          } else {
            lotMap.set(cleanKey, {
              id: `hold-${i}`,
              lotNo: String(lotNo).trim(),
              party: getVal(row, headerMap, ["Party Name", "Party"]) || "Standard",
              item: getVal(row, headerMap, ["Item", "Garment Type"]) || "",
              fabric: "",
              qty: parseFloat(getVal(row, headerMap, ["Quantity", "Qty"])) || 0,
              stage: dept ? `${dept} (ON HOLD)` : "ON HOLD",
              status: "ON HOLD",
              dateStr: "",
              pendingDays: 4,
              isHold: true,
              holdReason: reason || "On Hold",
              holdDept: dept || "Factory",
              issuedTo: "",
              issueDate: "",
              stitchingOutput: 0,
              packingStatus: "",
              sourceSheets: ["HoldLots"]
            });
          }
        }
      }

      // Step 5: Overlay Issues Sheet (Embroidery / Printing / Stitching Issues)
      if (issuesRes.status === 'fulfilled' && issuesRes.value?.ok && Array.isArray(issuesRes.value.values) && issuesRes.value.values.length > 1) {
        const rows = issuesRes.value.values;
        const headerMap = buildHeaderMap(rows[0]);

        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length === 0) continue;
          const lotNo = getVal(row, headerMap, ["Lot Number", "Lot No", "Lot", "Lot#"]);
          if (!lotNo) continue;

          const cleanKey = String(lotNo).trim().toLowerCase();
          const issueStage = getVal(row, headerMap, ["Issue To", "Stage", "Department", "Process", "Vendor", "Issue Type"]);
          const issueDate = getVal(row, headerMap, ["Date", "Issue Date", "Date of Issue"]);

          if (lotMap.has(cleanKey)) {
            const existing = lotMap.get(cleanKey);
            if (issueStage) {
              existing.issuedTo = issueStage;
              existing.stage = `Issued to ${issueStage}`;
            }
            if (issueDate) existing.issueDate = issueDate;
            existing.sourceSheets.push("Issues");
          }
        }
      }

      // Step 6: Overlay Rawpack Sheet (Packing Status)
      if (rawpackRes.status === 'fulfilled' && rawpackRes.value?.ok && Array.isArray(rawpackRes.value.values) && rawpackRes.value.values.length > 1) {
        const rows = rawpackRes.value.values;
        const headerMap = buildHeaderMap(rows[0]);

        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length === 0) continue;
          const lotNo = getVal(row, headerMap, ["Lot Number", "Lot No", "Lot"]);
          if (!lotNo) continue;

          const cleanKey = String(lotNo).trim().toLowerCase();
          const packStatus = getVal(row, headerMap, ["Status", "Packing Status", "Ready for Pack", "Remarks"]);

          if (lotMap.has(cleanKey)) {
            const existing = lotMap.get(cleanKey);
            existing.packingStatus = packStatus || "Packing Handover";
            if (packStatus.toLowerCase().includes("pack") || packStatus.toLowerCase().includes("ready")) {
              existing.stage = "Packing / Handover";
            }
            existing.sourceSheets.push("Rawpack");
          }
        }
      }

      // Format combined lots list
      const combinedLots = Array.from(lotMap.values()).map(l => ({
        ...l,
        severity: l.isHold ? 'CRITICAL' : l.pendingDays >= 5 ? 'CRITICAL' : l.pendingDays >= 3 ? 'HIGH' : 'MODERATE'
      }));

      setFetchedLots(combinedLots);
    } catch (e) {
      console.warn("Global multi-sheet lots fetch error:", e);
    } finally {
      setIsFetchingData(false);
    }
  };

  const activeLots = passedLotsData || fetchedLots;

  // Speech Recognition (Web Speech API)
  const handleToggleVoice = () => {
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SpeechRecognition) {
      alert("Voice input is not supported in this browser. Please use Google Chrome or Microsoft Edge.");
      return;
    }

    if (isListening) {
      setIsListening(false);
      return;
    }

    try {
      const recognition = new SpeechRecognition();
      recognition.lang = "hi-IN"; // Supports Hindi & Hinglish
      recognition.interimResults = false;
      recognition.maxAlternatives = 1;

      recognition.onstart = () => setIsListening(true);
      recognition.onend = () => setIsListening(false);
      recognition.onerror = (e) => {
        console.warn("Speech recognition error:", e.error);
        setIsListening(false);
      };
      recognition.onresult = (e) => {
        const transcript = e.results[0][0].transcript;
        if (transcript) {
          setQuery(transcript);
          handleExecuteQuery(transcript);
        }
      };

      recognition.start();
    } catch (err) {
      console.error("Speech recognition startup error:", err);
      setIsListening(false);
    }
  };

  /**
   * Smart Semantic Parser (Hinglish/English)
   */
  const processQueryLocally = (rawQuery, dataList) => {
    const q = String(rawQuery || "").toLowerCase().trim();
    const lots = Array.isArray(dataList) ? dataList : [];

    let filtered = [...lots];
    let detectedStage = null;
    let detectedItemType = null;
    let detectedMinDays = null;
    let targetReportPath = null;

    // Check for General Textile & Math Knowledge Queries
    const isGsm = q.includes("gsm") || q.includes("fabric weight") || q.includes("gram per square meter");
    const isConsumption = q.includes("consumption") || q.includes("kitna kapda") || q.includes("fabric required") || q.includes("kapda calculation");
    const isShrinkage = q.includes("shrinkage") || q.includes("sikudna") || q.includes("wash shrinkage");
    const isEfficiency = q.includes("sam") || q.includes("smv") || q.includes("efficiency") || q.includes("line output") || q.includes("pitch time");
    const isQuality = q.includes("dhu") || q.includes("aql") || q.includes("defect") || q.includes("quality");
    const isGreeting = q === "hi" || q === "hello" || q === "namaste" || q === "hey" || q.startsWith("hello") || q.startsWith("hi ") || q === "hlo" || q === "hy";
    const isWellBeing = q.includes("kaise ho") || q.includes("kya haal") || q.includes("how are you") || q.includes("sab badhiya") || q.includes("kya kr rhe ho") || q.includes("kya kar rahe ho") || q.includes("what are you doing") || q.includes("kya chal raha");
    const isThanks = q.includes("thank") || q.includes("shukriya") || q.includes("dhanyawad");
    const isHelp = q.includes("kaun ho") || q.includes("who are you") || q.includes("kya kar sakte ho") || q.includes("what can you do") || q.includes("help") || q.includes("madad");
    
    // Math pattern e.g. 2400 * 18 or 500 / 4
    const mathMatch = q.match(/(\d+(?:\.\d+)?)\s*([\+\-\*\/xX])\s*(\d+(?:\.\d+)?)/);
    const isMath = mathMatch && !q.includes("lot") && !q.includes("stage");

    if (isGsm || isConsumption || isShrinkage || isEfficiency || isQuality || isGreeting || isWellBeing || isThanks || isHelp || isMath) {
      let generalAnswer = "";
      if (isGreeting) {
        generalAnswer = `Namaste! 👋 Main Factory AI Assistant hoon.\n\nMain aapke **Factory Suite Pro** ke saare reports (Cutting, Embroidery, Printing, Stitching, Packing, Lots) aur garment calculations me help kar sakta hoon.\n\nAap aaj production bottlenecks check karna chahte hain, ya kisi specific lot ka status janna chahte hain?`;
      } else if (isWellBeing) {
        generalAnswer = `Main Factory Suite Pro ke saare production stages (Cutting, Embroidery, Printing, Stitching, Packing) ka live data monitor kar raha hoon aur aapke sawalon ke jawab dene ke liye taiyar hoon! 🚀\n\nAap kisi specific lot ka status check karna chahte hain, ya production bottleneck dekhna chahte hain?`;
      } else if (isThanks) {
        generalAnswer = `Aapka swagat hai! 😊 Agar production analysis, delayed lots, ya textile formulas me aur koi zaroorat ho, toh zaroor puchiye.`;
      } else if (isHelp) {
        generalAnswer = `### 🤖 Factory AI Assistant Capabilities\n\nMain Factory Suite Pro ka dedicated production intelligence copilot hoon:\n\n` +
          `1. **📊 Live Factory Stage Queries:**\n` +
          `   • *"cutting pr ss k lot kitna pending h"*\n` +
          `   • *"embroidery me 5 din se jyada delay wale lots"*\n` +
          `   • *"stitching me pending status"*\n` +
          `   • *"packing pending to issue kitna h"*\n\n` +
          `2. **🔍 Lot Tracking:**\n` +
          `   • *"lot 2045 ka current status kya h"*\n\n` +
          `3. **📐 Textile Formulas & Calculations:**\n` +
          `   • GSM, fabric consumption, wash shrinkage, SAM line efficiency & math calculations.\n\n` +
          `Aap voice mic 🎤 se bolkar ya text type karke sawal puch sakte hain!`;
      } else if (isGsm) {
        generalAnswer = `### 🧵 GSM (Grams per Square Meter) Calculation Guide\n\n` +
          `**Formula:**\n` +
          `• $\\text{GSM} = \\frac{\\text{Weight of Sample (grams)}}{\\text{Area (sq. meters)}} = \\frac{\\text{Weight in Grams} \\times 10,000}{\\text{Length (cm)} \\times \\text{Width (cm)}}$\n\n` +
          `**Quick GSM Round-Cutter Method:**\n` +
          `• 100 cm² Round GSM Cutter se sample cut karein.\n` +
          `• Electronic balance par weigh karein.\n` +
          `• $\\text{GSM} = \\text{Weight (grams)} \\times 100$\n\n` +
          `**Standard Garment GSM Ranges:**\n` +
          `• **Single Jersey T-Shirt:** 160 – 190 GSM\n` +
          `• **Pique Polo:** 200 – 240 GSM\n` +
          `• **Fleece / Hoodie:** 280 – 360 GSM\n` +
          `• **Rib / Collar:** 220 – 260 GSM`;
      } else if (isConsumption) {
        generalAnswer = `### 📐 Garment Fabric Consumption Calculation\n\n` +
          `**Basic T-Shirt Fabric Consumption (in Kg / Dozen):**\n` +
          `• $\\text{Grams/Pc} = \\frac{(\\text{Length} + \\text{Allowance}) \\times (\\text{Chest} + \\text{Allowance}) \\times 2 \\times \\text{GSM}}{10,000} + \\text{Sleeve Consumption}$\n\n` +
          `**Standard Estimations (per piece):**\n` +
          `• **Half Sleeve T-Shirt (180 GSM):** ~180 to 220 grams (4.5 - 5.5 pcs/Kg)\n` +
          `• **Full Sleeve T-Shirt (180 GSM):** ~240 to 280 grams (3.5 - 4.2 pcs/Kg)\n` +
          `• **Polo T-Shirt (220 GSM):** ~280 to 330 grams\n` +
          `• **Lower / Trackpant (240 GSM):** ~320 to 380 grams\n` +
          `• **Hoodie (320 GSM Fleece):** ~550 to 700 grams`;
      } else if (isShrinkage) {
        generalAnswer = `### 🧪 Fabric Wash Shrinkage Formula\n\n` +
          `**Formula:**\n` +
          `• $\\text{Shrinkage \\%} = \\frac{\\text{Original Length} - \\text{Length after Wash}}{\\text{Original Length}} \\times 100$\n\n` +
          `• **Acceptable Limit for Knits:** Length ±4% to 5%, Width ±3% to 4%.`;
      } else if (isEfficiency) {
        generalAnswer = `### ⏱️ Stitching Line Efficiency & SAM Formula\n\n` +
          `**Formula:**\n` +
          `• $\\text{Efficiency \\%} = \\frac{\\text{Total Output (Pcs)} \\times \\text{SAM (Minutes)}}{\\text{Total Operators} \\times \\text{Working Hours} \\times 60} \\times 100$\n\n` +
          `• **Example:** 1,000 Pcs output, 12 min SAM, 25 operators, 8 hours = 100% efficiency.`;
      } else if (isQuality) {
        generalAnswer = `### 🎯 Quality DHU & AQL Standards\n\n` +
          `**DHU (Defects per Hundred Units) Formula:**\n` +
          `• $\\text{DHU} = \\frac{\\text{Total Defects Found}}{\\text{Total Garments Inspected}} \\times 100$\n\n` +
          `**AQL Benchmark Standards:**\n` +
          `• **AQL 2.5:** Major Defects export benchmark.\n` +
          `• **AQL 4.0:** Minor Defects tolerance limit.`;
      } else if (isMath && mathMatch) {
        const n1 = parseFloat(mathMatch[1]);
        const op = mathMatch[2].toLowerCase();
        const n2 = parseFloat(mathMatch[3]);
        let resVal = 0;
        if (op === "+") resVal = n1 + n2;
        else if (op === "-") resVal = n1 - n2;
        else if (op === "*" || op === "x") resVal = n1 * n2;
        else if (op === "/" && n2 !== 0) resVal = n1 / n2;
        generalAnswer = `### 🔢 Calculation Result\n\n• **Question:** ${n1} ${op} ${n2}\n• **Answer:** **${Number.isInteger(resVal) ? resVal.toLocaleString() : resVal.toFixed(2)}**`;
      }

      return {
        isGeneralKnowledge: true,
        generalAnswer,
        totalLots: 0,
        totalQty: 0,
        avgDays: 0,
        matchedLots: []
      };
    }

    // Stage Detection
    if (q.includes("cut") || q.includes("cutting") || q.includes("katayi")) {
      detectedStage = "Cutting";
      targetReportPath = "/cutting-report";
      filtered = filtered.filter(l => {
        const s = `${l.stage || ''} ${l.stageId || ''}`.toLowerCase();
        return s.includes("cut");
      });
    } else if (q.includes("emb") || q.includes("embroidery") || q.includes("kadhai") || q.includes("embroidary")) {
      detectedStage = "Embroidery";
      targetReportPath = "/embroidery";
      filtered = filtered.filter(l => {
        const s = `${l.stage || ''} ${l.stageId || ''}`.toLowerCase();
        return s.includes("emb");
      });
    } else if (q.includes("print") || q.includes("printing") || q.includes("chapai")) {
      detectedStage = "Printing";
      targetReportPath = "/printing";
      filtered = filtered.filter(l => {
        const s = `${l.stage || ''} ${l.stageId || ''}`.toLowerCase();
        return s.includes("print");
      });
    } else if (q.includes("post") || q.includes("after emb") || q.includes("pending issue") || q.includes("issue to stitch")) {
      detectedStage = "Post-EMB/Print Issue";
      targetReportPath = "/pending-issue-to-stitching";
      filtered = filtered.filter(l => {
        const s = `${l.stage || ''} ${l.stageId || ''}`.toLowerCase();
        return s.includes("post") || s.includes("pending issue") || s.includes("after emb");
      });
    } else if (q.includes("stitch") || q.includes("silai") || q.includes("stitching")) {
      detectedStage = "Stitching WIP";
      targetReportPath = "/stitching-complete-lot";
      filtered = filtered.filter(l => {
        const s = `${l.stage || ''} ${l.stageId || ''}`.toLowerCase();
        return s.includes("stitch");
      });
    } else if (q.includes("pack") || q.includes("packing")) {
      detectedStage = "Packing Pending Handover";
      targetReportPath = "/pending-packing-issue";
      filtered = filtered.filter(l => {
        const s = `${l.stage || ''} ${l.stageId || ''}`.toLowerCase();
        return s.includes("pack");
      });
    } else if (q.includes("hold") || q.includes("roka") || q.includes("ruka hua") || q.includes("block")) {
      detectedStage = "Hold Lots";
      targetReportPath = "/hold-lot-manager";
      filtered = filtered.filter(l => l.isHold || String(l.stage || '').toLowerCase().includes("hold") || String(l.status || '').toLowerCase().includes("hold"));
    }

    // Garment Category Filter (e.g. Shirt, Shacket, Jacket, Lower, Hoodie, Sweatshirt, T-Shirt)
    if (/\b(shacke?t|shakt)\b/i.test(q)) {
      detectedItemType = "Shacket";
      filtered = filtered.filter(l => {
        const text = `${l.item || ''} ${l.style || ''} ${l.garmentType || ''}`.toLowerCase();
        return text.includes("shacket") || text.includes("shakt");
      });
    } else if (/\b(shirt)\b/i.test(q) && !/\b(t-?shirt|tshirt|sweatshirt|shacke?t|shakt)\b/i.test(q)) {
      detectedItemType = "Shirt";
      filtered = filtered.filter(l => {
        const text = `${l.item || ''} ${l.style || ''} ${l.garmentType || ''}`.toLowerCase();
        return text.includes("shirt") && !text.includes("shacket") && !text.includes("shakt") && !text.includes("t-shirt") && !text.includes("tshirt") && !text.includes("sweatshirt");
      });
    } else if (/\b(jacket|windcheater|bomber|puffer)\b/i.test(q)) {
      detectedItemType = "Jacket";
      filtered = filtered.filter(l => {
        const text = `${l.item || ''} ${l.style || ''} ${l.garmentType || ''}`.toLowerCase();
        return text.includes("jacket") || text.includes("windcheater") || text.includes("bomber") || text.includes("puffer");
      });
    } else if (/\b(lower|track|pant|jogger)\b/i.test(q)) {
      detectedItemType = "Lower / Track Pant";
      filtered = filtered.filter(l => {
        const text = `${l.item || ''} ${l.style || ''} ${l.garmentType || ''}`.toLowerCase();
        return text.includes("lower") || text.includes("track") || text.includes("pant") || text.includes("jogger");
      });
    } else if (/\b(hoodie|hood)\b/i.test(q)) {
      detectedItemType = "Hoodie";
      filtered = filtered.filter(l => {
        const text = `${l.item || ''} ${l.style || ''} ${l.garmentType || ''}`.toLowerCase();
        return text.includes("hood");
      });
    } else if (/\b(sweatshirt|sweat|fleece)\b/i.test(q)) {
      detectedItemType = "Sweatshirt";
      filtered = filtered.filter(l => {
        const text = `${l.item || ''} ${l.style || ''} ${l.garmentType || ''}`.toLowerCase();
        return text.includes("sweat") || text.includes("fleece");
      });
    } else if (/\b(t-?shirt|polo|tshirt)\b/i.test(q)) {
      detectedItemType = "T-Shirt / Polo";
      filtered = filtered.filter(l => {
        const text = `${l.item || ''} ${l.style || ''} ${l.garmentType || ''}`.toLowerCase();
        return text.includes("t-shirt") || text.includes("tshirt") || text.includes("polo") || text.includes("round neck");
      });
    } else if (/\b(ss|s\/s)\b/i.test(q)) {
      detectedItemType = "SS (Short Sleeve)";
      filtered = filtered.filter(l => {
        const text = `${l.item || ''} ${l.party || ''} ${l.description || ''} ${l.style || ''} ${l.lotNo || ''}`.toLowerCase();
        return text.includes("ss") || text.includes("short sleeve") || text.includes("half sleeve");
      });
    } else if (/\b(fs|f\/s)\b/i.test(q)) {
      detectedItemType = "FS (Full Sleeve)";
      filtered = filtered.filter(l => {
        const text = `${l.item || ''} ${l.party || ''} ${l.description || ''} ${l.style || ''} ${l.lotNo || ''}`.toLowerCase();
        return text.includes("fs") || text.includes("full sleeve");
      });
    }

    // Days / Delay Filter
    const daysMatch = q.match(/(\d+)\s*(?:din|days?|d)\b/i);
    if (daysMatch) {
      detectedMinDays = parseInt(daysMatch[1], 10);
      filtered = filtered.filter(l => {
        const days = parseFloat(l.pendingDays || l.delayDays || l.daysPending || 0);
        return days >= detectedMinDays;
      });
    } else if (q.includes("critical") || q.includes("red zone") || q.includes("sabse purana") || q.includes("oldest")) {
      filtered = filtered.filter(l => {
        const days = parseFloat(l.pendingDays || l.delayDays || l.daysPending || 0);
        const sev = String(l.severity || "").toUpperCase();
        return days >= 3 || sev.includes("CRITICAL") || sev.includes("HIGH");
      });
    }

    // Specific Lot Number lookup
    const ignoreWords = new Set([
      "kitna", "kitne", "kya", "kisko", "kiska", "kaunsa", "konsa", "pending", "pening", "hai", "h",
      "status", "detail", "details", "list", "ka", "ke", "ki", "me", "pr", "pe", "par", "kis",
      "process", "stage", "chal", "rha", "raha", "kaha", "kisme", "batao", "dikhao"
    ]);

    let searchedLotNo = null;
    // 1. Number before "lot" e.g. "61000 lot kis process pr hai"
    const numBeforeLot = q.match(/\b([a-zA-Z0-9_\-\/]{3,12})\s+lots?\b/i);
    if (numBeforeLot && /\d/.test(numBeforeLot[1])) {
      searchedLotNo = numBeforeLot[1].toLowerCase().trim();
    }

    // 2. "lot" followed by number e.g. "lot 61000", "lot #61000" (MUST contain at least one digit)
    if (!searchedLotNo) {
      const lotAfter = q.match(/\blots?\s*(?:no\.?|#|num|number)?\s*([a-zA-Z0-9_\-\/]+)/i);
      if (lotAfter && lotAfter[1]) {
        const cand = lotAfter[1].toLowerCase().trim();
        if (/\d/.test(cand) && !ignoreWords.has(cand)) {
          searchedLotNo = cand;
        }
      }
    }

    // 3. Standalone 4 to 7 digit lot number e.g. "61000 kis stage me hai"
    if (!searchedLotNo) {
      const standalone = q.match(/\b(\d{4,7}|[a-zA-Z]{1,3}[-_]\d{3,6})\b/i);
      if (standalone && !q.includes("gsm") && !q.includes("sam")) {
        searchedLotNo = standalone[1].toLowerCase().trim();
      }
    }

    if (searchedLotNo) {
      // Find exact or partial matching lot
      const exactMatches = lots.filter(l => String(l.lotNo || "").toLowerCase().trim() === searchedLotNo);
      const directMatches = exactMatches.length > 0 ? exactMatches : lots.filter(l => String(l.lotNo || "").toLowerCase().includes(searchedLotNo));

      if (directMatches.length > 0) {
        const found = directMatches[0];
        const days = parseFloat(found.pendingDays || found.delayDays || 0) || 0;
        const qtyFormatted = (parseFloat(found.qty) || 0).toLocaleString();

        // Calculate lag metrics (OverallCuttingtoPacking.js standard)
        const parseD = (dStr) => {
          if (!dStr) return null;
          if (dStr instanceof Date) return isNaN(dStr.getTime()) ? null : dStr;
          const s = String(dStr).trim();
          if (!s || s === '-' || s.toLowerCase() === 'invalid date') return null;
          const p = new Date(s);
          return isNaN(p.getTime()) ? null : p;
        };
        const calcDiff = (sStr, eStr) => {
          const s = parseD(sStr);
          const e = parseD(eStr);
          if (!s || !e) return '';
          const diff = Math.round((e.getTime() - s.getTime()) / (1000 * 60 * 60 * 24));
          return diff >= 0 ? `${diff} Days` : '';
        };

        const cutDays = calcDiff(found.jobDate, found.cutDate);
        const embDays = calcDiff(found.embIssueDate, found.embCompDate);
        const stitDays = calcDiff(found.stitchDate, found.compStitch);
        const pkgDays = calcDiff(found.pkgDate, found.pkgComp);
        const cutToEmb = calcDiff(found.cutDate, found.embIssueDate);
        const embToStit = calcDiff(found.embCompDate, found.stitchDate);
        const stitToPkg = calcDiff(found.compStitch, found.pkgDate);

        let lotResponse = "";

        // Detect user's specific question intent
        if (q.includes("cut") || q.includes("katayi") || q.includes("cutting kab") || q.includes("cut kab") || q.includes("cut date")) {
          lotResponse = `### ✂️ Lot #${found.lotNo} Cutting Details\n\n` +
            `• **Cut Date:** **${found.cutDate || 'Not Available'}** ${cutDays ? `*(Cut Lag from JobOrder: ${cutDays})*` : ''}\n` +
            `• **Job Order Date:** ${found.jobDate || 'Not Available'}\n` +
            `• **Cutting Quantity:** **${qtyFormatted} Pcs**\n` +
            `• **Garment Style:** ${found.item || 'Garment'} ${found.fabric ? `(${found.fabric})` : ''}\n` +
            `• **Buyer / Party:** **${found.party || 'Mohit Hosiery'}** ${found.brand ? `(${found.brand})` : ''}\n` +
            `• **Current Stage:** **${found.stage || 'Production WIP'}**`;
        } else if (q.includes("sup") || q.includes("supervisor") || q.includes("master") || q.includes("incharge")) {
          lotResponse = `### 👤 Lot #${found.lotNo} Supervisor Details\n\n` +
            `• **Stitching Supervisor:** **${found.stitchSup || 'Not Assigned'}**\n` +
            `• **Current Stage:** **${found.stage || 'Production WIP'}**\n` +
            `• **Cutting Quantity:** **${qtyFormatted} Pcs**\n` +
            `• **Buyer / Party:** **${found.party || 'Mohit Hosiery'}** ${found.brand ? `(${found.brand})` : ''}`;
        } else if (q.includes("emb") || q.includes("print") || q.includes("challan") || q.includes("kadhai") || q.includes("chapai")) {
          lotResponse = `### 🧵 Lot #${found.lotNo} Embroidery & Printing Status\n\n` +
            `• **Challan History:** 🧵 **${found.challanInfo || 'Direct Stitching / No External Challan'}**\n` +
            `• **Emb/Print Issue Date:** ${found.embIssueDate || 'Not Issued'} ${cutToEmb ? `*(Lag from Cut: ${cutToEmb})*` : ''}\n` +
            `• **Emb/Print Completion:** ${found.embCompDate || 'Pending / In Progress'} ${embDays ? `*(Process Time: ${embDays})*` : ''}\n` +
            `• **Direct Stitching:** ${found.directStitching ? found.directStitching.toUpperCase() : 'NO'}\n` +
            `• **Current Stage:** **${found.stage || 'Production WIP'}**`;
        } else if (q.includes("stitch") || q.includes("silai")) {
          lotResponse = `### 🪡 Lot #${found.lotNo} Stitching Details\n\n` +
            `• **Stitching Supervisor:** **${found.stitchSup || 'Not Assigned'}**\n` +
            `• **Stitch Start Date:** ${found.stitchDate || 'Not Started'} ${embToStit ? `*(Lag from Emb: ${embToStit})*` : ''}\n` +
            `• **Stitch Completion Date:** ${found.compStitch || 'In Progress'} ${stitDays ? `*(Process Time: ${stitDays})*` : ''}\n` +
            `• **Stitching Issue Qty:** **${found.stitchIssueQty ? found.stitchIssueQty.toLocaleString() : qtyFormatted} Pcs**\n` +
            (found.stitchingOutput ? `• **Stitching Output:** **${found.stitchingOutput.toLocaleString()} Pcs**\n` : '') +
            `• **Current Stage:** **${found.stage || 'Production WIP'}**`;
        } else if (q.includes("pack") || q.includes("packing") || q.includes("carton") || q.includes("barcode")) {
          lotResponse = `### 📦 Lot #${found.lotNo} Packing Details\n\n` +
            `• **Packing Issue Date:** ${found.pkgDate || 'Not Issued'} ${stitToPkg ? `*(Lag from Stitch: ${stitToPkg})*` : ''}\n` +
            `• **Packing Complete Date:** ${found.pkgComp || 'Pending'} ${pkgDays ? `*(Process: ${pkgDays})*` : ''}\n` +
            (found.barcodeDetails ? `• **Carton Barcodes:** **${found.barcodeDetails.cartonCount || 'Packed'} Cartons**\n` : '') +
            `• **Current Stage:** **${found.stage || 'Production WIP'}**`;
        } else if (q.includes("qty") || q.includes("quantity") || q.includes("piece") || q.includes("pcs") || q.includes("kitna piece") || q.includes("kitni qty")) {
          lotResponse = `### 📊 Lot #${found.lotNo} Quantity Breakdown\n\n` +
            `• **Cutting Quantity:** **${qtyFormatted} Pcs**\n` +
            (found.stitchIssueQty ? `• **Stitching Issue Qty:** **${found.stitchIssueQty.toLocaleString()} Pcs**\n` : '') +
            `• **Buyer / Party:** **${found.party || 'Mohit Hosiery'}** ${found.brand ? `(${found.brand})` : ''}\n` +
            `• **Garment Style:** ${found.item || 'Garment'}\n` +
            `• **Current Stage:** **${found.stage || 'Production WIP'}**`;
        } else if (q.includes("fabric") || q.includes("cloth") || q.includes("kapda") || q.includes("quality") || q.includes("shade") || q.includes("size") || q.includes("color")) {
          lotResponse = `### 🧵 Lot #${found.lotNo} Fabric & Style Info\n\n` +
            `• **Fabric / Quality:** **${found.fabric || 'Standard'}**\n` +
            `• **Garment Style:** **${found.item || 'Garment'}**\n` +
            (found.season ? `• **Season / Section:** ${found.season} ${found.section ? `(${found.section})` : ''}\n` : '') +
            (found.sizes ? `• **Sizes:** ${found.sizes}\n` : '') +
            (found.shades ? `• **Shades:** ${found.shades}\n` : '');
        } else if (q.includes("wash")) {
          const w = found.washingDetails;
          lotResponse = `### 🧼 Lot #${found.lotNo} Washing Status\n\n` +
            `• **Washing Status:** **${w ? (w.isComplete ? 'Completed ✅' : 'WIP ⏳') : 'Not in Washing / Standard'}**\n` +
            (w && w.supervisor ? `• **Supervisor:** ${w.supervisor}\n` : '') +
            `• **Current Stage:** **${found.stage || 'Production WIP'}**`;
        } else if (q.includes("overlock")) {
          const ov = found.overlockDetails;
          lotResponse = `### ⚡ Lot #${found.lotNo} Overlock Status\n\n` +
            `• **Overlock Status:** **${ov ? (ov.isComplete ? 'Completed ✅' : 'WIP ⏳') : 'Standard Line Flow'}**\n` +
            `• **Current Stage:** **${found.stage || 'Production WIP'}**`;
        } else if (q.includes("kaj") || q.includes("button")) {
          const kj = found.kajDetails;
          lotResponse = `### 🔘 Lot #${found.lotNo} Kaj / Button Status\n\n` +
            `• **Kaj Button Status:** **${kj ? (kj.isComplete ? 'Completed ✅' : 'WIP ⏳') : 'Standard Line Flow'}**\n` +
            `• **Current Stage:** **${found.stage || 'Production WIP'}**`;
        } else if (q.includes("hold") || q.includes("pause") || q.includes("ruka") || q.includes("quarantine")) {
          lotResponse = found.isHold
            ? `### 🚨 Lot #${found.lotNo} Hold Status\n\n• **Status:** **ON HOLD 🛑**\n• **Reason:** **${found.holdReason || 'Quarantined'}**\n• **Department:** **${found.holdDept || 'Factory Floor'}**`
            : `### ✅ Lot #${found.lotNo} Hold Status\n\n• **Status:** **Active (Not on hold)**\n• **Current Stage:** **${found.stage || 'Production WIP'}**`;
        } else if (q.includes("kaha") || q.includes("kis process") || q.includes("kis stage") || q.includes("kisme") || q.includes("status")) {
          lotResponse = `### 📍 Lot #${found.lotNo} Current Stage & Location\n\n` +
            `• **Current Stage:** **${found.stage || 'Production WIP'}**\n` +
            `• **Aging / Process Time:** ⏳ **${days} Days**\n` +
            `• **Buyer / Party:** **${found.party || 'Mohit Hosiery'}** ${found.brand ? `(${found.brand})` : ''}\n` +
            `• **Quantity:** **${qtyFormatted} Pcs**\n` +
            `• **Style:** ${found.item || 'Garment'}`;
        } else {
          // Full End-to-End Factory Department Traversal
          const subOps = found.sub_operations || {};
          const fmtSub = (op, icon) => {
            const o = subOps[op] || {};
            const st = o.status || 'Not Logged';
            const sup = o.supervisor ? ` | Sup: ${o.supervisor}` : '';
            const wip = (o.wip && o.wip !== '-' && o.wip !== '[]') ? ` — *${o.wip}*` : '';
            const em = st === 'Completed' ? '✅' : (st === 'WIP' ? '⏳' : '⚪');
            return `• ${icon} **${op}:** ${em} **${st}**${sup}${wip}`;
          };

          const zips = found.zip_details || [];
          const doris = found.dori_details || [];
          const zipStr = zips.length > 0 ? zips.map(z => `Supplier: ${z.supplier || '-'} (Qty: ${z.po_qty || '-'}, Entry: ${z.material_entry_date || 'Pending'})`).join(', ') : 'No specific zipper PO attached';
          const doriStr = doris.length > 0 ? doris.map(d => `Supplier: ${d.supplier || '-'} (Qty: ${d.po_qty || '-'}, Entry: ${d.entry_date || 'Pending'})`).join(', ') : 'No specific dori PO attached';

          lotResponse = `### 🏭 Lot #${found.lotNo} — Complete Factory Department Traversal\n\n` +
            `📍 **Live Current Location:** **${found.stage || 'Production WIP'}**\n` +
            `• **Party / Buyer:** **${found.party || 'Mohit Hosiery'}** ${found.brand ? `(${found.brand})` : ''} | **Style:** **${found.item || 'Garment'}**\n` +
            `• **Fabric / Quality:** ${found.fabric || 'Standard Fabric'} | **Total Cut Quantity:** **${qtyFormatted} Pcs**\n\n` +
            `---\n` +
            `#### 📋 1. Job Order Department (\`JobOrder\` Sheet)\n` +
            `• **Job Order Date:** ${found.jobDate || 'Not Available'} | **Season/Section:** ${found.season || '-'} (${found.section || '-'})\n` +
            `• **Sizes:** ${found.sizes || '-'} | **Shades:** ${found.shades || '-'}\n\n` +
            `#### ✂️ 2. Cutting Floor (\`Cutting\` & \`Index\` Sheet)\n` +
            `• **Cutting Date:** **${found.cutDate || 'Pending'}** ${cutDays ? `*(Cut Lag from JobOrder: ${cutDays})*` : ''}\n` +
            `• **Cutting Quantity:** **${qtyFormatted} Pcs**\n\n` +
            `#### 🧵 3. Embroidery & Printing Department (\`EmbroideryChallan\`, \`PrintingChallan\`)\n` +
            `• **Challan Issued:** ${found.embIssueDate || 'Direct Stitching / No External Challan'} ${cutToEmb ? `*(Lag from Cut: ${cutToEmb})*` : ''}\n` +
            `• **Challan Received / Comp:** ${found.embCompDate || 'Pending / In Progress'} ${embDays ? `*(Process: ${embDays})*` : ''}\n` +
            `• **Challan Details:** ${found.challanInfo || 'Standard Line Flow'}\n\n` +
            `#### 🪡 4. Stitching Floor (\`DailyStitching\` & \`Index\` Sheet)\n` +
            `• **Stitching Start Date:** ${found.stitchDate || 'Not Started'} ${embToStit ? `*(Lag from Emb: ${embToStit})*` : ''}\n` +
            `• **Stitching Supervisor:** **${found.stitchSup || 'Not Assigned'}**\n` +
            `• **Stitching Issue Qty:** **${found.stitchIssueQty ? found.stitchIssueQty.toLocaleString() : qtyFormatted} Pcs**\n` +
            `• **Stitching Completed:** ${found.compStitch || 'In Progress'} ${stitDays ? `*(Process: ${stitDays})*` : ''}\n\n` +
            `#### ⚙️ 5. Intermediate Finishing Sub-Operations (\`Working Updates\` - 6 Sheets)\n` +
            `${fmtSub("Kaj Button", "🔘")}\n` +
            `${fmtSub("Overlock", "🪡")}\n` +
            `${fmtSub("FeedUp", "🪡")}\n` +
            `${fmtSub("Washing", "🧼")}\n` +
            `${fmtSub("Folding", "📦")}\n` +
            `${fmtSub("Elastic", "🩳")}\n\n` +
            `#### 🤐 6. Trims & Accessories PO Status (\`ZipPurchaseOrders\`, \`DoriPurchaseOrders\`)\n` +
            `• 🤐 **Zipper PO:** ${zipStr}\n` +
            `• 🎗️ **Drawcord (Dori) PO:** ${doriStr}\n\n` +
            `#### 🛡️ 7. Quality & Hold Quarantine Status (\`Hold_Lots\` Sheet)\n` +
            `• **Quarantine Status:** ${found.isHold ? `🚨 **ON HOLD** — ${found.holdReason || 'Quarantined'} (${found.holdDept || 'Factory Floor'})` : '✅ **Active (Approved / Not on hold)**'}\n\n` +
            `#### 📦 8. Packing Floor & Final Yield (\`Issues\` & \`RAWPACK\` Sheet)\n` +
            `• **Packing Issue Date:** ${found.pkgDate || 'Not Issued'} ${stitToPkg ? `*(Lag from Stitch: ${stitToPkg})*` : ''}\n` +
            `• **Packing Supervisor:** **${found.pkgSup || 'Not Assigned'}**\n` +
            `• **Packing Complete Date:** ${found.pkgComp || 'Pending'}\n\n` +
            `---\n` +
            `📍 **Summary:** Lot #${found.lotNo} is currently at **${found.stage || 'Production WIP'}** in the factory pipeline.`;
        }

        return {
          isSpecificLot: true,
          isGeneralKnowledge: false,
          targetLotNo: found.lotNo,
          totalLots: directMatches.length,
          totalQty: directMatches.reduce((acc, l) => acc + (parseFloat(l.qty) || 0), 0),
          avgDays: days,
          detectedStage: found.stage,
          customResponse: lotResponse,
          matchedLots: directMatches
        };
      } else {
        const notFoundResponse = `🔍 **Lot #${searchedLotNo.toUpperCase()} Status Not Found**\n\n` +
          `• Lot **#${searchedLotNo.toUpperCase()}** factory production records mein nahi mila.\n` +
          `• **Possible Reasons:**\n` +
          `  1. Lot number me typing mistake ho sakti hai.\n` +
          `  2. Lot abhi JobOrder sheet me create nahi hua hai ya complete hokar dispatch ho chuka hai.\n` +
          `• **Tip:** Aap JobOrder ya Cutting report me lot number verify karein.`;

        return {
          isSpecificLot: true,
          isGeneralKnowledge: false,
          targetLotNo: searchedLotNo,
          totalLots: 0,
          totalQty: 0,
          avgDays: 0,
          detectedStage: null,
          customResponse: notFoundResponse,
          matchedLots: []
        };
      }
    }

    // If query has no stage, no lot, no delay filter, and no production keywords, answer conversationally
    const hasProdKeywords = q.includes("lot") || q.includes("pending") || q.includes("delay") || q.includes("stage") ||
      q.includes("report") || q.includes("piece") || q.includes("qty") || q.includes("quantity") || q.includes("kisme") ||
      q.includes("kaha") || q.includes("kaun") || q.includes("party") || q.includes("order") || q.includes("status") ||
      q.includes("cut") || q.includes("emb") || q.includes("print") || q.includes("stitch") || q.includes("pack");

    if (!detectedStage && !detectedItemType && !detectedMinDays && !searchedLotNo && !hasProdKeywords) {
      return {
        isGeneralKnowledge: true,
        generalAnswer: `Main Factory AI Copilot hoon! 😊\n\nMain aapke Factory Suite Pro ke saare production stages (Cutting, Embroidery, Printing, Stitching, Packing, Lots) aur garment calculations ko monitor karta hoon.\n\nAap kisi specific stage (Cutting, Embroidery, Stitching, Packing) ya specific lot number ke bare me puchna chahte hain?`,
        totalLots: 0,
        totalQty: 0,
        avgDays: 0,
        matchedLots: []
      };
    }

    // Sort by pending days / qty descending
    filtered.sort((a, b) => {
      const dA = parseFloat(a.pendingDays || a.delayDays || 0) || 0;
      const dB = parseFloat(b.pendingDays || b.delayDays || 0) || 0;
      return dB - dA;
    });

    const totalLots = filtered.length;
    const totalQty = filtered.reduce((acc, l) => acc + (parseFloat(l.qty || l.pieces || l.totalQty || 0) || 0), 0);
    const avgDays = totalLots > 0
      ? Math.round(filtered.reduce((acc, l) => acc + (parseFloat(l.pendingDays || l.delayDays || 0) || 0), 0) / totalLots)
      : 0;

    return {
      isGeneralKnowledge: false,
      totalLots,
      totalQty,
      avgDays,
      detectedStage,
      detectedItemType,
      detectedMinDays,
      targetReportPath,
      matchedLots: filtered
    };
  };

  const handleExecuteQuery = async (queryToRun) => {
    const qText = String(queryToRun || query || "").trim();
    if (!qText) return;

    const userMessage = {
      id: Date.now(),
      sender: "user",
      text: qText,
      timestamp: new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })
    };

    setChatHistory(prev => [...prev, userMessage]);
    setQuery("");
    setLoading(true);

    // 1. Process client-side analysis
    const localAnalysis = processQueryLocally(qText, activeLots);

    // 2. Prepare base structured response
    let responseText = "";
    if (localAnalysis.customResponse) {
      responseText = localAnalysis.customResponse;
    } else if (localAnalysis.isGeneralKnowledge && localAnalysis.generalAnswer) {
      responseText = localAnalysis.generalAnswer;
    } else if (localAnalysis.isGeneralKnowledge) {
      responseText = `Namaste! 👋 Main Factory AI Assistant hoon.\n\nMain aapke factory data (Cutting, Embroidery, Stitching, Packing, Lots) aur textile formulas ke sawal ka turant jawab de sakta hoon.`;
    } else if (localAnalysis.totalLots === 0) {
      responseText = `🔍 **Search Result:** Aapki query **"${qText}"** ke matching koi pending lot nahi mila.\n\n` +
        `• **Checked Stage:** ${localAnalysis.detectedStage || "All Factory Stages"}\n` +
        `• **Tip:** Aap stage name (Cutting, Embroidery, Stitching, Packing) ya specific lot number enter karke search kar sakte hain!`;
    } else {
      const stageStr = localAnalysis.detectedStage ? `📍 **Stage:** ${localAnalysis.detectedStage}` : "📍 **Stage:** All Stages";
      const itemStr = localAnalysis.detectedItemType ? ` | 🏷️ **Type:** ${localAnalysis.detectedItemType}` : "";
      
      responseText = `### 📊 AI Production Query Result\n` +
        `• **Total Pending Lots:** **${localAnalysis.totalLots} Lots**\n` +
        `• **Total Quantity:** **${localAnalysis.totalQty.toLocaleString()} Pcs**\n` +
        `• **Average Pending Time:** **${localAnalysis.avgDays} Days**\n` +
        `• ${stageStr}${itemStr}\n\n` +
        `Neeche matching lots ki detailed summary di gayi hai:`;
    }

    // Prepare conversation history payload for multi-turn conversational AI
    const historyPayload = chatHistory.slice(-10).map(m => ({
      role: m.sender === "user" ? "user" : "assistant",
      content: m.text
    }));

    // Try Python FastAPI microservice first, then Node.js backend
    try {
      // 1. Try Python FastAPI microservice (port 8000)
      const pyRes = await fetch("http://127.0.0.1:8000/api/ai/query", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ query: qText, history: historyPayload, department: selectedDept })
      });
      if (pyRes.ok) {
        const pyJson = await pyRes.json();
        if (pyJson && pyJson.answer) {
          responseText = pyJson.answer;
        }
      } else {
        throw new Error("Python microservice responded with error");
      }
    } catch (pyErr) {
      // 2. Fallback to Node.js backend (port 5000)
      try {
        const backendUrl = BACKEND_API_BASE_URL || "http://localhost:5000";
        const savedKey = localStorage.getItem("FACTORY_AI_GEMINI_KEY") || "";
        const res = await fetch(`${backendUrl}/api/ai/query`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            prompt: qText,
            apiKey: savedKey,
            context: {
              activeStage,
              department: selectedDept,
              totalLotsCount: activeLots.length,
              summaryKPIs: {
                matchedLotsCount: localAnalysis.totalLots,
                matchedTotalQty: localAnalysis.totalQty,
                avgDays: localAnalysis.avgDays
              },
              sampleLots: localAnalysis.matchedLots.slice(0, 40),
              history: historyPayload
            }
          })
        });

        if (res.ok) {
          const json = await res.json();
          if (json && json.answer) {
            responseText = json.answer;
          }
        }
      } catch (nodeErr) {
        // Fallback already in responseText
      }
    }

    const botMessage = {
      id: Date.now() + 1,
      sender: "bot",
      text: responseText,
      analysis: localAnalysis,
      timestamp: new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })
    };

    setChatHistory(prev => [...prev, botMessage]);
    setLoading(false);

    // Speak response automatically if voice auto-play is enabled
    if (autoSpeak) {
      setTimeout(() => speakText(responseText, botMessage.id), 150);
    }
  };

  const handleCopyText = (text, index) => {
    navigator.clipboard.writeText(text);
    setCopiedIndex(index);
    setTimeout(() => setCopiedIndex(null), 2000);
  };

  const handleNavigate = (path) => {
    if (path) {
      stopSpeech();
      onClose();
      history.push(path);
    }
  };

  return (
    <>
      {/* 1. Global Floating Button (Shown when window is closed) */}
      {!isOpen && (
        <button
          type="button"
          onClick={() => setInternalIsOpen(true)}
          title="Open Factory AI Assistant Window"
          style={styles.floatingTriggerBtn}
        >
          <span style={styles.floatingPulse}></span>
          <span style={{ fontSize: "20px" }}>✨</span>
          <span style={styles.floatingBtnText}>Ask AI Assistant</span>
        </button>
      )}

      {/* 2. Standalone AI Pop-Up Window */}
      {isOpen && (
        <div style={styles.overlay} onClick={() => { stopSpeech(); onClose(); }}>
          <div style={styles.popupWindow} onClick={(e) => e.stopPropagation()}>
            {/* Header */}
            <div style={styles.header}>
              <div style={styles.headerLeft}>
                <div style={styles.aiIconBadge}>✨</div>
                <div>
                  <div style={styles.titleRow}>
                    <span style={styles.title}>Factory AI Assistant</span>
                    <span style={styles.liveBadge}>● Active Assistant</span>
                  </div>
                  <p style={styles.subtitle}>
                    Ask anything in Hinglish / English across all factory production stages & lots
                  </p>
                </div>
              </div>
              <div style={{ display: "flex", gap: "10px", alignItems: "center" }}>
                {/* Gemini Key Settings Button */}
                <button
                  type="button"
                  onClick={() => setShowKeyModal(true)}
                  style={{
                    background: geminiKey ? "rgba(99, 102, 241, 0.25)" : "rgba(255, 255, 255, 0.1)",
                    color: geminiKey ? "#818cf8" : "#94a3b8",
                    border: `1px solid ${geminiKey ? "rgba(99, 102, 241, 0.5)" : "rgba(255, 255, 255, 0.2)"}`,
                    padding: "6px 10px",
                    borderRadius: "8px",
                    fontSize: "12px",
                    fontWeight: "700",
                    cursor: "pointer",
                    display: "flex",
                    alignItems: "center",
                    gap: "4px"
                  }}
                  title="Configure Gemini API Key (Optional for live LLM mode)"
                >
                  ⚙️ {geminiKey ? "Gemini Key: Active" : "Key: Default Brain"}
                </button>

                {chatHistory.length > 0 && (
                  <button
                    type="button"
                    onClick={() => {
                      stopSpeech();
                      setChatHistory([]);
                    }}
                    style={{
                      background: "rgba(255, 255, 255, 0.1)",
                      color: "#cbd5e1",
                      border: "1px solid rgba(255, 255, 255, 0.2)",
                      padding: "6px 10px",
                      borderRadius: "8px",
                      fontSize: "12px",
                      fontWeight: "700",
                      cursor: "pointer"
                    }}
                    title="Clear conversation history"
                  >
                    🧹 Clear
                  </button>
                )}

                {/* Voice Auto-Play Toggle */}
                <button
                  type="button"
                  onClick={() => {
                    if (isSpeaking) stopSpeech();
                    setAutoSpeak(!autoSpeak);
                  }}
                  style={{
                    background: autoSpeak ? "rgba(16, 185, 129, 0.2)" : "rgba(255, 255, 255, 0.1)",
                    color: autoSpeak ? "#34d399" : "#94a3b8",
                    border: `1px solid ${autoSpeak ? "rgba(16, 185, 129, 0.4)" : "rgba(255, 255, 255, 0.2)"}`,
                    padding: "6px 12px",
                    borderRadius: "8px",
                    fontSize: "12px",
                    fontWeight: "700",
                    cursor: "pointer",
                    display: "flex",
                    alignItems: "center",
                    gap: "5px"
                  }}
                  title={autoSpeak ? "Voice Auto-Speak is ON. Click to mute." : "Voice Auto-Speak is MUTED. Click to turn on."}
                >
                  {autoSpeak ? "🔊 Voice: ON" : "🔇 Voice: OFF"}
                </button>

                {isSpeaking && (
                  <button
                    type="button"
                    onClick={stopSpeech}
                    style={{
                      background: "#ef4444",
                      color: "#ffffff",
                      border: "none",
                      padding: "6px 12px",
                      borderRadius: "8px",
                      fontSize: "12px",
                      fontWeight: "800",
                      cursor: "pointer",
                      display: "flex",
                      alignItems: "center",
                      gap: "4px",
                      animation: "pulse 1.2s infinite"
                    }}
                    title="Stop current speech"
                  >
                    ⏹ Stop Voice
                  </button>
                )}

                <button
                  style={styles.closeBtn}
                  onClick={() => { stopSpeech(); onClose(); }}
                  title="Close Window"
                >
                  ✕
                </button>
              </div>
            </div>

            {/* Optional Gemini Key Settings Modal */}
            {showKeyModal && (
              <div style={styles.keyModalOverlay} onClick={() => setShowKeyModal(false)}>
                <div style={styles.keyModalContent} onClick={(e) => e.stopPropagation()}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "12px" }}>
                    <h4 style={{ margin: 0, fontSize: "15px", fontWeight: "800", color: "#0f172a" }}>⚙️ AI Engine Key Settings</h4>
                    <button onClick={() => setShowKeyModal(false)} style={{ border: "none", background: "none", cursor: "pointer", fontSize: "16px" }}>✕</button>
                  </div>
                  <p style={{ fontSize: "12.5px", color: "#64748b", margin: "0 0 12px 0", lineHeight: "1.4" }}>
                    Factory AI Assistant works out of the box with the built-in comprehensive brain. If you want full open-ended Google Gemini 1.5 live conversational reasoning, paste your free Gemini API key below:
                  </p>
                  <input
                    type="password"
                    placeholder="AIzaSy... (Paste Gemini API Key)"
                    value={geminiKey}
                    onChange={(e) => setGeminiKey(e.target.value)}
                    style={{ width: "100%", padding: "10px 12px", borderRadius: "8px", border: "1.5px solid #cbd5e1", fontSize: "13px", marginBottom: "12px", boxSizing: "border-box" }}
                  />
                  <div style={{ display: "flex", justifyContent: "flex-end", gap: "8px" }}>
                    <button
                      type="button"
                      onClick={() => {
                        localStorage.removeItem("FACTORY_AI_GEMINI_KEY");
                        setGeminiKey("");
                        setShowKeyModal(false);
                      }}
                      style={{ padding: "6px 12px", borderRadius: "6px", border: "1px solid #cbd5e1", background: "#f8fafc", fontSize: "12px", cursor: "pointer" }}
                    >
                      Reset to Default
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        localStorage.setItem("FACTORY_AI_GEMINI_KEY", geminiKey.trim());
                        setShowKeyModal(false);
                      }}
                      style={{ padding: "6px 14px", borderRadius: "6px", border: "none", background: "#4f46e5", color: "#ffffff", fontSize: "12px", fontWeight: "700", cursor: "pointer" }}
                    >
                      Save Key
                    </button>
                  </div>
                </div>
              </div>
            )}

            {/* Specialized Department Expert Selector Bar */}
            <div style={{
              background: "#1e1b4b",
              padding: "8px 16px",
              display: "flex",
              alignItems: "center",
              gap: "8px",
              overflowX: "auto",
              borderBottom: "1px solid rgba(255, 255, 255, 0.1)"
            }}>
              <span style={{ fontSize: "11.5px", fontWeight: "800", color: "#a5b4fc", textTransform: "uppercase", letterSpacing: "0.5px", whiteSpace: "nowrap" }}>
                🎯 Department Expert Mode:
              </span>
              {[
                { id: "all", label: "🏢 All Depts" },
                { id: "joborder", label: "📋 Job Order" },
                { id: "cutting", label: "✂️ Cutting" },
                { id: "embroidery", label: "🧵 Emb/Print" },
                { id: "stitching", label: "🪡 Stitching" },
                { id: "sub_ops", label: "⚙️ Sub-Ops" },
                { id: "trims", label: "🤐 Trims/PO" },
                { id: "quality", label: "🛡️ Quality/Hold" },
                { id: "packing", label: "📦 Packing" }
              ].map(dept => {
                const isActive = selectedDept === dept.id;
                return (
                  <button
                    key={dept.id}
                    type="button"
                    onClick={() => setSelectedDept(dept.id)}
                    style={{
                      background: isActive ? "linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%)" : "rgba(255, 255, 255, 0.08)",
                      color: isActive ? "#ffffff" : "#cbd5e1",
                      border: `1.5px solid ${isActive ? "#a78bfa" : "rgba(255, 255, 255, 0.15)"}`,
                      padding: "5px 12px",
                      borderRadius: "16px",
                      fontSize: "12px",
                      fontWeight: isActive ? "800" : "600",
                      cursor: "pointer",
                      whiteSpace: "nowrap",
                      boxShadow: isActive ? "0 2px 8px rgba(139, 92, 246, 0.3)" : "none",
                      transition: "all 0.2s ease"
                    }}
                  >
                    {dept.label}
                  </button>
                );
              })}
            </div>

            {/* Chat Body */}
            <div style={styles.chatBody}>
              {chatHistory.length === 0 ? (
                <div style={styles.welcomeBox}>
                  <div style={styles.welcomeIcon}>✨</div>
                  <h3 style={styles.welcomeHeading}>Factory AI Assistant</h3>
                  <p style={styles.welcomeDesc}>
                    Talk naturally with AI in Hinglish, Hindi, or English — just like ChatGPT & Gemini! Ask about production stages, lot status, delays, or textile math.
                  </p>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: "8px", justifyContent: "center", marginTop: "16px", maxWidth: "620px" }}>
                    {[
                      "👋 Hello AI Assistant",
                      "📊 Show factory production bottlenecks",
                      "🧵 Cutting pending lots summary",
                      "🪡 Stitching WIP status",
                      "🔍 Lot 1024 complete status",
                      "📐 Fabric GSM calculation formula",
                      "⏱️ Delayed lots > 5 days"
                    ].map((promptText, pIdx) => (
                      <button
                        key={pIdx}
                        type="button"
                        onClick={() => handleExecuteQuery(promptText)}
                        style={{
                          background: "#ffffff",
                          color: "#4f46e5",
                          border: "1.5px solid #e0e7ff",
                          padding: "8px 14px",
                          borderRadius: "20px",
                          fontSize: "12.5px",
                          fontWeight: "600",
                          cursor: "pointer",
                          boxShadow: "0 2px 6px rgba(99, 102, 241, 0.08)",
                          transition: "all 0.2s ease"
                        }}
                        onMouseEnter={(e) => {
                          e.currentTarget.style.borderColor = "#6366f1";
                          e.currentTarget.style.background = "#eef2ff";
                        }}
                        onMouseLeave={(e) => {
                          e.currentTarget.style.borderColor = "#e0e7ff";
                          e.currentTarget.style.background = "#ffffff";
                        }}
                      >
                        {promptText}
                      </button>
                    ))}
                  </div>
                </div>
              ) : (
                chatHistory.map((msg, idx) => (
                  <div
                    key={msg.id || idx}
                    style={{
                      ...styles.messageWrapper,
                      justifyContent: msg.sender === "user" ? "flex-end" : "flex-start"
                    }}
                  >
                    {msg.sender === "bot" && (
                      <div style={styles.botAvatar}>✨</div>
                    )}
                    <div
                      style={{
                        ...styles.messageBubble,
                        ...(msg.sender === "user" ? styles.userBubble : styles.botBubble)
                      }}
                    >
                      <div style={styles.messageHeader}>
                        <span style={styles.senderName}>
                          {msg.sender === "user" ? "You" : "Factory AI Assistant"}
                        </span>
                        <span style={styles.msgTime}>{msg.timestamp}</span>
                      </div>

                      <div style={styles.msgText}>
                        {msg.text.split("\n").map((rawLine, lIdx) => {
                          const line = rawLine.trim();
                          if (!line) {
                            return <div key={lIdx} style={{ height: "6px" }} />;
                          }
                          if (line.startsWith("### ")) {
                            return <h4 key={lIdx} style={styles.mdH4}>{line.replace("### ", "")}</h4>;
                          }
                          if (line.startsWith("## ")) {
                            return <h3 key={lIdx} style={{ ...styles.mdH4, fontSize: "16px", color: "#1e1b4b" }}>{line.replace("## ", "")}</h3>;
                          }
                          if (line.startsWith("• ") || line.startsWith("- ") || line.startsWith("* ") || rawLine.includes("• ")) {
                            const cleanBullet = line.replace(/^([•\-\*]\s*|\s*•\s*)/, "");
                            return (
                              <div key={lIdx} style={styles.mdBullet}>
                                <span style={styles.mdBulletDot}>•</span>
                                <span>{renderFormattedText(cleanBullet)}</span>
                              </div>
                            );
                          }
                          const numMatch = line.match(/^(\d+)\.\s+(.*)/);
                          if (numMatch) {
                            return (
                              <div key={lIdx} style={{ ...styles.mdBullet, alignItems: "flex-start", marginTop: "4px" }}>
                                <span style={{ background: "#ede9fe", color: "#6366f1", fontWeight: "800", fontSize: "11px", padding: "1px 6px", borderRadius: "10px", marginRight: "6px" }}>
                                  {numMatch[1]}
                                </span>
                                <span>{renderFormattedText(numMatch[2])}</span>
                              </div>
                            );
                          }
                          return <p key={lIdx} style={styles.mdP}>{renderFormattedText(line)}</p>;
                        })}
                      </div>

                      {/* Matching lots breakdown cards */}
                      {msg.analysis && msg.analysis.matchedLots && msg.analysis.matchedLots.length > 0 && (
                        <div style={styles.cardsSection}>
                          <div style={styles.cardHeaderBar}>
                            <span style={styles.cardSectionTitle}>
                              📋 Matching Lots Detail ({msg.analysis.matchedLots.length})
                            </span>
                            <div style={{ display: "flex", gap: "6px" }}>
                              {msg.analysis.targetReportPath && (
                                <button
                                  style={styles.openReportBtn}
                                  onClick={() => handleNavigate(msg.analysis.targetReportPath)}
                                >
                                  🚀 Open {msg.analysis.detectedStage || "Stage"} Report ↗
                                </button>
                              )}
                              {onApplyFilter && (
                                <button
                                  style={styles.applyFilterBtn}
                                  onClick={() => {
                                    onApplyFilter(msg.analysis.matchedLots.map(l => l.lotNo));
                                    onClose();
                                  }}
                                >
                                  🎯 Apply Filter to Table
                                </button>
                              )}
                            </div>
                          </div>

                          <div style={styles.lotCardsList}>
                            {msg.analysis.matchedLots.slice(0, 10).map((lot, lIdx) => {
                              const days = lot.pendingDays || lot.delayDays || 0;
                              const isCritical = days >= 5;
                              const isHigh = days >= 3 && days < 5;
                              return (
                                <div key={lIdx} style={styles.lotCard}>
                                  <div style={styles.lotCardTop}>
                                    <span style={styles.lotCardLotNo}>Lot #{lot.lotNo}</span>
                                    <span
                                      style={{
                                        ...styles.lotDaysBadge,
                                        backgroundColor: isCritical ? "#fee2e2" : isHigh ? "#ffedd5" : "#f1f5f9",
                                        color: isCritical ? "#dc2626" : isHigh ? "#c2410c" : "#475569",
                                        borderColor: isCritical ? "#fca5a5" : isHigh ? "#fdba74" : "#cbd5e1"
                                      }}
                                    >
                                      ⏳ {days} Days
                                    </span>
                                  </div>
                                  <div style={styles.lotCardDetails}>
                                    <div><strong>Party/Item:</strong> {lot.party || lot.item || "Standard"}</div>
                                    <div><strong>Quantity:</strong> <span style={{ color: "#059669", fontWeight: "700" }}>{lot.qty || 0} Pcs</span></div>
                                    <div><strong>Stage:</strong> {lot.stage || "Production WIP"}</div>
                                  </div>
                                </div>
                              );
                            })}
                          </div>

                          {msg.analysis.matchedLots.length > 10 && (
                            <div style={styles.moreLotsNotice}>
                              + {msg.analysis.matchedLots.length - 10} more lots matched.
                            </div>
                          )}
                        </div>
                      )}

                      {/* Bottom Actions: Speak / Stop / Copy */}
                      {msg.sender === "bot" && (
                        <div style={styles.msgBottomBar}>
                          <button
                            style={{
                              ...styles.speakBtn,
                              color: isSpeaking && currentSpeakingId === msg.id ? "#ef4444" : "#4f46e5"
                            }}
                            onClick={() => speakText(msg.text, msg.id)}
                            title={isSpeaking && currentSpeakingId === msg.id ? "Stop voice" : "Read this answer aloud"}
                          >
                            {isSpeaking && currentSpeakingId === msg.id ? "⏹ Stop" : "🔊 Listen"}
                          </button>

                          <button
                            style={styles.copyBtn}
                            onClick={() => handleCopyText(msg.text, idx)}
                            title="Copy response text"
                          >
                            {copiedIndex === idx ? "✓ Copied" : "📋 Copy"}
                          </button>
                        </div>
                      )}
                    </div>
                  </div>
                ))
              )}

              {loading && (
                <div style={{ ...styles.messageWrapper, justifyContent: "flex-start" }}>
                  <div style={styles.botAvatar}>✨</div>
                  <div style={{ ...styles.messageBubble, ...styles.botBubble, padding: "14px 20px" }}>
                    <div style={styles.loadingDots}>
                      <div style={styles.dot}></div>
                      <div style={{ ...styles.dot, animationDelay: "0.2s" }}></div>
                      <div style={{ ...styles.dot, animationDelay: "0.4s" }}></div>
                      <span style={{ marginLeft: "10px", fontSize: "13px", color: "#64748b" }}>
                        Analyzing production data...
                      </span>
                    </div>
                  </div>
                </div>
              )}
              <div ref={chatEndRef} />
            </div>

            {/* Active Voice Listening Banner */}
            {isListening && (
              <div style={{
                background: "linear-gradient(90deg, #ef4444 0%, #dc2626 100%)",
                color: "#ffffff",
                padding: "8px 16px",
                fontSize: "12.5px",
                fontWeight: "700",
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                borderBottom: "1px solid #b91c1c"
              }}>
                <span>🎙️ Voice Mic Active — Speak now in Hindi, Hinglish, or English...</span>
                <button
                  onClick={() => setIsListening(false)}
                  style={{ background: "rgba(255,255,255,0.2)", border: "none", color: "#ffffff", borderRadius: "4px", padding: "2px 8px", cursor: "pointer", fontWeight: "800" }}
                >
                  ✕ Stop Mic
                </button>
              </div>
            )}

            {/* Input Bar */}
            <div style={styles.inputContainer}>
              <button
                style={{
                  ...styles.voiceBtn,
                  backgroundColor: isListening ? "#ef4444" : "#f1f5f9",
                  color: isListening ? "#ffffff" : "#475569"
                }}
                onClick={handleToggleVoice}
                title={isListening ? "Listening... click to stop" : "Speak in Hindi or English"}
              >
                {isListening ? "🔴 Stop Mic" : "🎤 Voice"}
              </button>

              <input
                ref={inputRef}
                type="text"
                style={styles.textInput}
                placeholder="Ask anything in Hinglish / English..."
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === "Enter" && !e.shiftKey) {
                    e.preventDefault();
                    handleExecuteQuery(query);
                  }
                }}
              />

              <button
                style={{
                  ...styles.sendBtn,
                  opacity: query.trim() ? 1 : 0.6,
                  cursor: query.trim() ? "pointer" : "not-allowed"
                }}
                onClick={() => handleExecuteQuery(query)}
                disabled={!query.trim() || loading}
              >
                Ask AI ⚡
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}

// Markdown formatting helper
function renderFormattedText(txt) {
  if (!txt) return "";
  // Clean math delimiters e.g. $\text{...}$
  const cleanTxt = txt
    .replace(/\$\\text\{([^}]+)\}\$/g, "$1")
    .replace(/\$([^$]+)\$/g, "$1");

  const parts = cleanTxt.split(/(\*\*[^*]+\*\*|`[^`]+`|\*[^*]+\*)/g);
  return parts.map((part, i) => {
    if (part.startsWith("**") && part.endsWith("**")) {
      return <strong key={i} style={{ color: "#0f172a", fontWeight: "700" }}>{part.slice(2, -2)}</strong>;
    }
    if (part.startsWith("*") && part.endsWith("*") && !part.startsWith("**")) {
      return <em key={i} style={{ color: "#475569", fontStyle: "italic" }}>{part.slice(1, -1)}</em>;
    }
    if (part.startsWith("`") && part.endsWith("`")) {
      return (
        <code key={i} style={{ backgroundColor: "#e2e8f0", padding: "2px 6px", borderRadius: "4px", fontSize: "12px", color: "#0369a1" }}>
          {part.slice(1, -1)}
        </code>
      );
    }
    return part;
  });
}

// Complete styles system
const styles = {
  floatingTriggerBtn: {
    position: "fixed",
    bottom: "28px",
    right: "28px",
    zIndex: 99998,
    display: "flex",
    alignItems: "center",
    gap: "10px",
    background: "linear-gradient(135deg, #4f46e5 0%, #7c3aed 50%, #ec4899 100%)",
    color: "#ffffff",
    border: "2px solid rgba(255, 255, 255, 0.4)",
    padding: "12px 22px",
    borderRadius: "30px",
    boxShadow: "0 10px 30px rgba(124, 58, 237, 0.45)",
    cursor: "pointer",
    fontSize: "14px",
    fontWeight: "800",
    letterSpacing: "-0.01em",
    transition: "transform 0.2s, box-shadow 0.2s"
  },
  floatingPulse: {
    width: "10px",
    height: "10px",
    borderRadius: "50%",
    backgroundColor: "#10b981",
    boxShadow: "0 0 10px #10b981"
  },
  floatingBtnText: {
    textShadow: "0 1px 2px rgba(0,0,0,0.2)"
  },
  overlay: {
    position: "fixed",
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    backgroundColor: "rgba(15, 23, 42, 0.75)",
    backdropFilter: "blur(6px)",
    zIndex: 99999,
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    padding: "16px",
    animation: "fadeIn 0.2s ease-out"
  },
  popupWindow: {
    width: "100%",
    maxWidth: "880px",
    height: "85vh",
    maxHeight: "800px",
    backgroundColor: "#ffffff",
    borderRadius: "20px",
    boxShadow: "0 25px 60px -12px rgba(0, 0, 0, 0.45)",
    display: "flex",
    flexDirection: "column",
    overflow: "hidden",
    border: "1px solid #e2e8f0"
  },
  header: {
    padding: "18px 22px",
    backgroundColor: "#0f172a",
    color: "#ffffff",
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    borderBottom: "1px solid #334155"
  },
  headerLeft: {
    display: "flex",
    alignItems: "center",
    gap: "14px"
  },
  aiIconBadge: {
    width: "44px",
    height: "44px",
    borderRadius: "14px",
    background: "linear-gradient(135deg, #6366f1 0%, #a855f7 50%, #ec4899 100%)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    fontSize: "22px",
    boxShadow: "0 0 16px rgba(168, 85, 247, 0.5)"
  },
  titleRow: {
    display: "flex",
    alignItems: "center",
    gap: "10px"
  },
  title: {
    fontSize: "19px",
    fontWeight: "800",
    color: "#ffffff",
    letterSpacing: "-0.02em"
  },
  liveBadge: {
    fontSize: "11px",
    fontWeight: "700",
    backgroundColor: "rgba(16, 185, 129, 0.2)",
    color: "#10b981",
    padding: "2px 8px",
    borderRadius: "12px",
    border: "1px solid rgba(16, 185, 129, 0.3)"
  },
  subtitle: {
    margin: "4px 0 0 0",
    fontSize: "12.5px",
    color: "#94a3b8"
  },
  closeBtn: {
    background: "rgba(255,255,255,0.1)",
    border: "none",
    color: "#ffffff",
    fontSize: "18px",
    cursor: "pointer",
    padding: "8px 12px",
    borderRadius: "8px",
    transition: "all 0.2s"
  },
  suggestionsContainer: {
    padding: "10px 18px",
    backgroundColor: "#f8fafc",
    borderBottom: "1px solid #e2e8f0",
    display: "flex",
    alignItems: "center",
    gap: "10px"
  },
  suggestionTitle: {
    fontSize: "12px",
    fontWeight: "700",
    color: "#475569",
    whiteSpace: "nowrap"
  },
  suggestionScroll: {
    display: "flex",
    gap: "8px",
    overflowX: "auto",
    paddingBottom: "2px"
  },
  suggestionChip: {
    backgroundColor: "#ffffff",
    border: "1px solid #cbd5e1",
    padding: "5px 12px",
    borderRadius: "16px",
    fontSize: "12px",
    fontWeight: "600",
    color: "#334155",
    cursor: "pointer",
    whiteSpace: "nowrap",
    transition: "all 0.15s"
  },
  chatBody: {
    flex: 1,
    overflowY: "auto",
    padding: "20px",
    backgroundColor: "#f1f5f9",
    display: "flex",
    flexDirection: "column",
    gap: "16px"
  },
  welcomeBox: {
    margin: "auto",
    textAlign: "center",
    maxWidth: "560px",
    padding: "30px 20px"
  },
  welcomeIcon: {
    fontSize: "48px",
    marginBottom: "12px"
  },
  welcomeHeading: {
    fontSize: "18.5px",
    fontWeight: "800",
    color: "#0f172a",
    marginBottom: "8px"
  },
  welcomeDesc: {
    fontSize: "13px",
    color: "#64748b",
    marginBottom: "16px",
    lineHeight: "1.5"
  },
  exampleList: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    textAlign: "left"
  },
  exampleItem: {
    backgroundColor: "#ffffff",
    padding: "11px 15px",
    borderRadius: "10px",
    border: "1px solid #e2e8f0",
    fontSize: "13px",
    cursor: "pointer",
    color: "#334155",
    transition: "background 0.15s"
  },
  messageWrapper: {
    display: "flex",
    gap: "10px",
    width: "100%"
  },
  botAvatar: {
    width: "34px",
    height: "34px",
    borderRadius: "10px",
    background: "linear-gradient(135deg, #6366f1 0%, #a855f7 100%)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    fontSize: "16px",
    color: "#ffffff",
    flexShrink: 0
  },
  messageBubble: {
    maxWidth: "85%",
    borderRadius: "16px",
    padding: "15px 18px",
    position: "relative"
  },
  userBubble: {
    backgroundColor: "#2563eb",
    color: "#ffffff",
    borderBottomRightRadius: "4px"
  },
  botBubble: {
    backgroundColor: "#ffffff",
    color: "#1e293b",
    border: "1px solid #e2e8f0",
    boxShadow: "0 2px 4px rgba(0,0,0,0.04)",
    borderBottomLeftRadius: "4px"
  },
  messageHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: "8px",
    gap: "12px"
  },
  senderName: {
    fontSize: "12px",
    fontWeight: "700",
    opacity: 0.9
  },
  msgTime: {
    fontSize: "10px",
    opacity: 0.6
  },
  msgText: {
    fontSize: "13.5px",
    lineHeight: "1.55"
  },
  mdH4: {
    margin: "10px 0 6px 0",
    fontSize: "14px",
    fontWeight: "800",
    color: "#0f172a"
  },
  mdP: {
    margin: "0 0 6px 0"
  },
  mdBullet: {
    display: "flex",
    gap: "8px",
    margin: "3px 0"
  },
  mdBulletDot: {
    color: "#6366f1",
    fontWeight: "800"
  },
  cardsSection: {
    marginTop: "14px",
    paddingTop: "12px",
    borderTop: "1px solid #e2e8f0"
  },
  cardHeaderBar: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: "10px",
    flexWrap: "wrap",
    gap: "6px"
  },
  cardSectionTitle: {
    fontSize: "12px",
    fontWeight: "700",
    color: "#334155"
  },
  openReportBtn: {
    backgroundColor: "#4f46e5",
    color: "#ffffff",
    border: "none",
    padding: "6px 12px",
    borderRadius: "8px",
    fontSize: "12px",
    fontWeight: "700",
    cursor: "pointer"
  },
  applyFilterBtn: {
    backgroundColor: "#059669",
    color: "#ffffff",
    border: "none",
    padding: "6px 12px",
    borderRadius: "8px",
    fontSize: "12px",
    fontWeight: "700",
    cursor: "pointer"
  },
  lotCardsList: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))",
    gap: "8px"
  },
  lotCard: {
    backgroundColor: "#f8fafc",
    border: "1px solid #e2e8f0",
    borderRadius: "10px",
    padding: "10px 12px"
  },
  lotCardTop: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: "6px"
  },
  lotCardLotNo: {
    fontWeight: "800",
    fontSize: "13px",
    color: "#0f172a"
  },
  lotDaysBadge: {
    fontSize: "10px",
    fontWeight: "700",
    padding: "2px 6px",
    borderRadius: "6px",
    border: "1px solid"
  },
  lotCardDetails: {
    fontSize: "11.5px",
    color: "#475569",
    lineHeight: "1.4"
  },
  moreLotsNotice: {
    fontSize: "11px",
    color: "#64748b",
    marginTop: "8px",
    fontStyle: "italic",
    textAlign: "center"
  },
  msgBottomBar: {
    display: "flex",
    justifyContent: "flex-end",
    alignItems: "center",
    gap: "12px",
    marginTop: "8px"
  },
  speakBtn: {
    backgroundColor: "transparent",
    border: "none",
    fontSize: "11.5px",
    fontWeight: "700",
    cursor: "pointer",
    display: "inline-flex",
    alignItems: "center",
    gap: "3px"
  },
  copyBtn: {
    backgroundColor: "transparent",
    border: "none",
    fontSize: "11px",
    color: "#94a3b8",
    cursor: "pointer"
  },
  inputContainer: {
    padding: "14px 18px",
    backgroundColor: "#ffffff",
    borderTop: "1px solid #e2e8f0",
    display: "flex",
    gap: "10px",
    alignItems: "center"
  },
  voiceBtn: {
    padding: "10px 14px",
    borderRadius: "10px",
    border: "1px solid #cbd5e1",
    fontSize: "12px",
    fontWeight: "700",
    cursor: "pointer",
    whiteSpace: "nowrap"
  },
  textInput: {
    flex: 1,
    padding: "12px 16px",
    borderRadius: "10px",
    border: "1px solid #cbd5e1",
    fontSize: "14px",
    outline: "none",
    color: "#0f172a"
  },
  sendBtn: {
    backgroundColor: "#2563eb",
    color: "#ffffff",
    border: "none",
    padding: "12px 20px",
    borderRadius: "10px",
    fontSize: "13px",
    fontWeight: "700",
    transition: "all 0.2s"
  },
  loadingDots: {
    display: "flex",
    alignItems: "center"
  },
  dot: {
    width: "6px",
    height: "6px",
    backgroundColor: "#6366f1",
    borderRadius: "50%",
    margin: "0 2px"
  },
  keyModalOverlay: {
    position: "absolute",
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    backgroundColor: "rgba(15, 23, 42, 0.6)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    zIndex: 1000,
    padding: "20px"
  },
  keyModalContent: {
    backgroundColor: "#ffffff",
    borderRadius: "14px",
    padding: "20px",
    maxWidth: "460px",
    width: "100%",
    boxShadow: "0 20px 40px rgba(0,0,0,0.2)",
    border: "1px solid #e2e8f0"
  }
};
