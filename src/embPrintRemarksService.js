import { fetchSheetDataFromBackend, BACKEND_URL } from './config';

export const REMARKS_SPREADSHEET_ID =
  process.env.REACT_APP_EMB_PRINT_REMARKS_SPREADSHEET_ID ||
  '1ZAAVyKqAqQkBvwFv19pu1WT3g227XJ8ZpM_JSb_nMd8';

export const SHEET_TABS = {
  EMB: 'EMB REMARKS',
  PRINT: 'PRINT REMARKS',
  PACKING: 'PACKING REMARKS',
  PACKING_ALLOTED: 'PACKING ALLOTED REMARKS',
  PENDING_STITCHING: 'PENDING STITCHING REMARKS',
  CUTTING: 'CUTTING REMARKS'
};

export const WEBHOOK_URL =
  process.env.REACT_APP_EMB_PRINT_REMARKS_WEBHOOK_URL ||
  'https://script.google.com/macros/s/AKfycbyMDwX4P8mUmpkodGdoHQQvFMqW4z0LWvqeWFByh4pAF3GFDXrlLpGV9M7dHqHLB-bZ/exec';

/**
 * Recursively parses and cleans any remark value (JSON array, JSON object, stringified JSON, or plain text)
 * to extract ONLY the latest single human-readable remark/status string.
 */
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

    const hasTimestamp = validItems.some((item) => item && typeof item === "object" && (item.timestamp || item.date));
    let sorted = validItems;
    if (hasTimestamp) {
      sorted = [...validItems].sort((a, b) => {
        const timeA = a?.timestamp ? new Date(a.timestamp).getTime() : (a?.date ? new Date(a.date).getTime() : 0);
        const timeB = b?.timestamp ? new Date(b.timestamp).getTime() : (b?.date ? new Date(b.date).getTime() : 0);
        if (!isNaN(timeA) && !isNaN(timeB) && (timeA > 0 || timeB > 0)) {
          return timeB - timeA;
        }
        return 0;
      });
    }

    const latestItem = (hasTimestamp && sorted[0]?.timestamp && !isNaN(new Date(sorted[0].timestamp).getTime()))
      ? sorted[0]
      : (hasTimestamp && sorted[0]?.date && !isNaN(new Date(sorted[0].date).getTime()) ? sorted[0] : validItems[0]);

    return formatLatestRemark(latestItem, fallback);
  }

  // 3. If it's an Object (single update or remark object)
  if (val && typeof val === "object") {
    const rem = (val.remarks ?? val.remark ?? val.userRemarks ?? val.text ?? val.note ?? val.comment ?? "").toString().trim();
    const st = (val.status ?? val.wipStatus ?? val.updateType ?? val.compStatus ?? val.state ?? "").toString().trim();

    const isClean = (str) => str && str !== "—" && str !== "N/A" && str !== "-" && str !== "null" && str !== "undefined";

    if (isClean(st) && isClean(rem)) {
      if (st.toLowerCase() === rem.toLowerCase()) {
        return formatLatestRemark(st, fallback);
      }
      return `${formatLatestRemark(st, fallback)} - ${formatLatestRemark(rem, fallback)}`;
    }
    if (isClean(st)) {
      return formatLatestRemark(st, fallback);
    }
    if (isClean(rem)) {
      return formatLatestRemark(rem, fallback);
    }
    for (const v of Object.values(val)) {
      if (isClean(v)) {
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
  }
  return finalStr;
};

/** Helper function to format any date into DD-MM-YYYY */
/** Helper function to format any date into DD-MM-YYYY */
export const formatDateToDDMMYYYY = (dateInput) => {
  if (!dateInput) return null;
  const str = String(dateInput).trim();
  if (!str || str === 'N/A' || str === '-' || str === '—' || str === 'null' || str === 'undefined') return null;

  // 1. If string DD/MM/YYYY, DD-MM-YYYY, or YYYY-MM-DD
  const parts = str.split(/[\/\-\.]/);
  if (parts.length === 3) {
    const p0 = parseInt(parts[0], 10);
    const p1 = parseInt(parts[1], 10);
    const p2 = parseInt(parts[2], 10);

    // YYYY-MM-DD format (year first)
    if (p0 > 1000 && !isNaN(p0) && !isNaN(p1) && !isNaN(p2)) {
      return `${String(p2).padStart(2, '0')}-${String(p1).padStart(2, '0')}-${p0}`;
    }
    // DD-MM-YYYY or DD/MM/YYYY format (day first)
    if (p2 > 1000 || p2 >= 20) {
      const fullYear = p2 < 100 ? 2000 + p2 : p2;
      if (!isNaN(p0) && !isNaN(p1) && !isNaN(fullYear)) {
        return `${String(p0).padStart(2, '0')}-${String(p1).padStart(2, '0')}-${fullYear}`;
      }
    }
  }

  // 2. If ISO date, timestamp or standard Date string
  const d = new Date(str);
  if (!isNaN(d.getTime())) {
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    return `${day}-${month}-${year}`;
  }

  return null;
};

/**
 * Extracts latest remark AND appends the date of updation formatted as "Remark Text (DD-MM-YYYY)"
 */
export const formatLatestRemarkWithDate = (rawVal, fallback = "—", fallbackDate = null) => {
  const remarkText = formatLatestRemark(rawVal, fallback);
  if (!remarkText || remarkText === fallback || remarkText === "N/A" || remarkText === "—") {
    if (fallbackDate) {
      const formattedFallbackDate = formatDateToDDMMYYYY(fallbackDate);
      if (formattedFallbackDate) return `${fallback} (${formattedFallbackDate})`;
    }
    return fallback;
  }

  // Check if date is already in remarkText e.g. "Tailor Working (25-09-2026)" or "Tailor Working 25/09/2026"
  const existingDateMatch = remarkText.match(/\b(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})\b/);
  if (existingDateMatch && existingDateMatch[1]) {
    const cleanDate = formatDateToDDMMYYYY(existingDateMatch[1]);
    if (cleanDate) {
      if (remarkText.includes(`(${cleanDate})`)) {
        return remarkText;
      }
      const textWithoutDate = remarkText.replace(existingDateMatch[0], '').replace(/[\(\)\-\s]+$/, '').trim();
      return `${textWithoutDate || remarkText} (${cleanDate})`;
    }
  }

  // Try extracting date from JSON object / array inside rawVal
  let extractedDate = null;
  let parsed = rawVal;
  if (typeof rawVal === "string" && (rawVal.trim().startsWith("[") || rawVal.trim().startsWith("{") || rawVal.trim().startsWith('"['))) {
    try {
      parsed = JSON.parse(rawVal.trim());
    } catch (e) {
      try {
        parsed = JSON.parse(rawVal.trim().replace(/\\"/g, '"'));
      } catch (e2) {}
    }
  }

  // Handle double-stringified JSON
  if (typeof parsed === "string" && (parsed.trim().startsWith("[") || parsed.trim().startsWith("{"))) {
    try {
      parsed = JSON.parse(parsed.trim());
    } catch (e) {}
  }

  const getItemTimestamp = (item) => {
    if (!item || typeof item !== "object") return 0;
    const t = item.timestamp || item.date || item.updatedAt || item.time || item.created_at;
    if (!t) return 0;
    const time = new Date(t).getTime();
    return isNaN(time) ? 0 : time;
  };

  if (Array.isArray(parsed) && parsed.length > 0) {
    const validItems = parsed.filter((item) => item !== null && item !== undefined && item !== "");
    const sorted = [...validItems].sort((a, b) => getItemTimestamp(b) - getItemTimestamp(a));

    for (const item of sorted) {
      if (item && typeof item === "object") {
        const itemDate = item.timestamp || item.date || item.updatedAt || item.time || item.created_at;
        if (itemDate) {
          extractedDate = formatDateToDDMMYYYY(itemDate);
          if (extractedDate) break;
        }
      }
    }
  } else if (parsed && typeof parsed === "object") {
    const itemDate = parsed.timestamp || parsed.date || parsed.updatedAt || parsed.time || parsed.created_at;
    if (itemDate) {
      extractedDate = formatDateToDDMMYYYY(itemDate);
    }
  }

  const finalDate = extractedDate || formatDateToDDMMYYYY(fallbackDate);

  if (finalDate) {
    if (remarkText.includes(`(${finalDate})`)) {
      return remarkText;
    }
    return `${remarkText} (${finalDate})`;
  }

  return remarkText;
};

/**
 * Fetch remarks directly from Google Spreadsheet and fallback to localStorage cache.
 * Returns pure remark data keyed by Lot Number.
 */
export const fetchRemarksForTab = async (tabType) => {
  const tabName = SHEET_TABS[tabType] || tabType || 'PACKING REMARKS';
  const cacheKey = `fs_remarks_${tabType || 'PACKING'}`;
  let sheetMap = {};

  // Try loading from localStorage cache first for instant render
  try {
    const cached = localStorage.getItem(cacheKey);
    if (cached) {
      sheetMap = JSON.parse(cached) || {};
    }
  } catch (e) {}

  try {
    const res = await fetchSheetDataFromBackend(REMARKS_SPREADSHEET_ID, `${tabName}!A:H`);
    if (res && res.ok && Array.isArray(res.values) && res.values.length > 1) {
      const [headers, ...rows] = res.values;
      const freshMap = {};

      rows.forEach((row) => {
        const lotNumber = String(row[0] || '').trim();
        if (!lotNumber) return;

        const latestRemarkRaw = String(row[5] || '').trim();
        const historyJson = String(row[6] || '').trim();
        const updatedAt = String(row[7] || '').trim();

        let history = [];
        if (historyJson) {
          try {
            const parsed = JSON.parse(historyJson);
            if (Array.isArray(parsed)) {
              history = parsed.map((item) => {
                if (typeof item === 'object' && item !== null) {
                  return {
                    ...item,
                    text: formatLatestRemark(item.text || item.remarks || item.status || item, '')
                  };
                }
                return { text: formatLatestRemark(item, ''), timestamp: updatedAt || new Date().toLocaleString() };
              });
            }
          } catch {
            if (latestRemarkRaw) {
              history = [{ text: formatLatestRemark(latestRemarkRaw, ''), timestamp: updatedAt || new Date().toLocaleString() }];
            }
          }
        } else if (latestRemarkRaw) {
          history = [{ text: formatLatestRemark(latestRemarkRaw, ''), timestamp: updatedAt || new Date().toLocaleString() }];
        }

        freshMap[lotNumber] = history;
      });

      sheetMap = { ...sheetMap, ...freshMap };
      try {
        localStorage.setItem(cacheKey, JSON.stringify(sheetMap));
      } catch (e) {}
    }
  } catch (err) {
    console.warn(`Could not fetch remarks sheet data for ${tabName}:`, err.message);
  }

  return sheetMap;
};

/**
 * Save a new remark for a lot directly to Google Spreadsheet '1ZAAVyKqAqQkBvwFv19pu1WT3g227XJ8ZpM_JSb_nMd8'.
 * NO LOCALSTORAGE.
 * Enforces SINGLE ENTRY PER LOT in Google Sheet via Apps Script Webhook.
 */
export const saveRemarkForLot = async ({
  tabType, // 'EMB' | 'PRINT'
  lotNumber,
  challanNo = '',
  partyName = '',
  fabric = '',
  style = '',
  remarkText
}) => {
  if (!lotNumber || !remarkText || !remarkText.trim()) return null;

  const cleanLot = String(lotNumber).replace(/★\s*/, '').trim();
  const cleanText = remarkText.trim();
  const nowStr = new Date().toLocaleString('en-GB', {
    day: '2-digit',
    month: 'short',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });

  const tabName = SHEET_TABS[tabType] || tabType || 'PACKING REMARKS';
  const cacheKey = `fs_remarks_${tabType || 'PACKING'}`;

  // 1. Fetch current Google Sheet history for this lot
  const currentSheetMap = await fetchRemarksForTab(tabType);
  const existingHistory = currentSheetMap[cleanLot] || [];

  // Append new remark to single lot history array
  const updatedHistory = [...existingHistory, { text: cleanText, timestamp: nowStr }];

  // Immediately update local cache
  try {
    const updatedMap = { ...currentSheetMap, [cleanLot]: updatedHistory };
    localStorage.setItem(cacheKey, JSON.stringify(updatedMap));
  } catch (e) {}

  const payload = {
    spreadsheetId: REMARKS_SPREADSHEET_ID,
    sheetName: tabName,
    lotNumber: cleanLot,
    challanNo,
    partyName,
    fabric,
    style,
    latestRemark: cleanText,
    remarksHistoryJson: JSON.stringify(updatedHistory),
    updatedAt: nowStr
  };

  // 2. Post directly to Google Apps Script Webhook
  try {
    await fetch(WEBHOOK_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload),
      mode: 'no-cors'
    });
  } catch (err) {
    console.error('Error posting remark to Apps Script Webhook:', err);
  }

  // 3. Notify backend server to clear cache
  try {
    fetch(`${BACKEND_URL}/api/sheets/save-emb-print-remark`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    }).catch(() => {});
  } catch (err) {}

  // 4. Dispatch global custom event for instant UI render
  window.dispatchEvent(new CustomEvent('emb_print_remark_updated', {
    detail: { tabType, lotNumber: cleanLot, history: updatedHistory }
  }));

  return updatedHistory;
};
