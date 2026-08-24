import { fetchSheetDataFromBackend, BACKEND_URL } from './config';

export const REMARKS_SPREADSHEET_ID =
  process.env.REACT_APP_EMB_PRINT_REMARKS_SPREADSHEET_ID ||
  '1ZAAVyKqAqQkBvwFv19pu1WT3g227XJ8ZpM_JSb_nMd8';

export const SHEET_TABS = {
  EMB: 'EMB REMARKS',
  PRINT: 'PRINT REMARKS'
};

export const WEBHOOK_URL =
  process.env.REACT_APP_EMB_PRINT_REMARKS_WEBHOOK_URL ||
  'https://script.google.com/macros/s/AKfycbyMDwX4P8mUmpkodGdoHQQvFMqW4z0LWvqeWFByh4pAF3GFDXrlLpGV9M7dHqHLB-bZ/exec';

/**
 * Fetch remarks 100% directly from Google Spreadsheet '1ZAAVyKqAqQkBvwFv19pu1WT3g227XJ8ZpM_JSb_nMd8'.
 * NO LOCALSTORAGE. Returns pure Google Sheet data keyed by Lot Number.
 */
export const fetchRemarksForTab = async (tabType) => {
  const tabName = tabType === 'PRINT' ? SHEET_TABS.PRINT : SHEET_TABS.EMB;
  const sheetMap = {};

  try {
    const res = await fetchSheetDataFromBackend(REMARKS_SPREADSHEET_ID, `${tabName}!A:H`);
    if (res && res.ok && Array.isArray(res.values) && res.values.length > 1) {
      const [headers, ...rows] = res.values;

      rows.forEach((row) => {
        const lotNumber = String(row[0] || '').trim();
        if (!lotNumber) return;

        const latestRemark = String(row[5] || '').trim();
        const historyJson = String(row[6] || '').trim();
        const updatedAt = String(row[7] || '').trim();

        let history = [];
        if (historyJson) {
          try {
            history = JSON.parse(historyJson);
          } catch {
            if (latestRemark) {
              history = [{ text: latestRemark, timestamp: updatedAt || new Date().toLocaleString() }];
            }
          }
        } else if (latestRemark) {
          history = [{ text: latestRemark, timestamp: updatedAt || new Date().toLocaleString() }];
        }

        sheetMap[lotNumber] = history;
      });
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

  const tabName = tabType === 'PRINT' ? SHEET_TABS.PRINT : SHEET_TABS.EMB;

  // 1. Fetch current Google Sheet history for this lot
  const currentSheetMap = await fetchRemarksForTab(tabType);
  const existingHistory = currentSheetMap[cleanLot] || [];

  // Append new remark to single lot history array
  const updatedHistory = [...existingHistory, { text: cleanText, timestamp: nowStr }];

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
