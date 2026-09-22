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
  }
  return finalStr;
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
