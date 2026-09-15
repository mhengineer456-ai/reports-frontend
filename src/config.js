// src/config.js
// Centralized configuration module loading credentials safely from environment variables (.env)

export const GOOGLE_API_KEY =
  process.env.REACT_APP_GOOGLE_API_KEY || 'AIzaSyAomDFBkOySlIxKWSKGHe6ATv9gvaBr7uk';

export const SPREADSHEET_IDS = {
  MAIN: process.env.REACT_APP_MAIN_SPREADSHEET_ID || '1Hj3JeJEKB43aYYWv8gk2UhdU6BWuEQfCg5pBlTdBMNA',
  JOBORDER: process.env.REACT_APP_JOBORDER_SPREADSHEET_ID || '1fKSwGBIpzWEFk566WRQ4bzQ0anJlmasoY8TwrTLQHXI',
  ISSUES: process.env.REACT_APP_ISSUES_SPREADSHEET_ID || '1uo14nKO_yHu4AJ2rOgaJajuprcinj6xw1AUMFJ6_zYM',
  DAILY_STITCHING: process.env.REACT_APP_DAILY_STITCHING_SPREADSHEET_ID || '1IMhmYlJ3s2PPRgEQs1Ikd4O1OBXK4EYL1oV_-kWAkyg',
  WORKING_UPDATES: process.env.REACT_APP_WORKING_UPDATES_SPREADSHEET_ID || '1Nh7XYE_MnAxtaTRUUntHvBpzctODwjnkbBYDiYLQgoc',
  ZIP: process.env.REACT_APP_ZIP_SPREADSHEET_ID || '16mifNw0WMIlnZ1XRHsuH_8kVUm_6Y1O3uVsoM-Hjppo',
  SHADE_PO: process.env.REACT_APP_SHADE_PO_SPREADSHEET_ID || '1JgJF9Er7lYDW0rINQzUUafqonx0yxkVaAauPgX5QNfk',
  DORI: process.env.REACT_APP_DORI_SPREADSHEET_ID || '1LjwZqU26F0xwL1tEyps8txsM1qS8LLUuE-sy_4CQK6k',
  RAWPACK: process.env.REACT_APP_RAWPACK_SPREADSHEET_ID || '1xD8Uy1lUgvNTQ2RGRBI4ZjOrozbinUPRq2_UfIplP98',
  BARCODE: process.env.REACT_APP_BARCODE_SPREADSHEET_ID || '1dOCjNFwaAel5qun0_ZJVIGmREqjI76CJBBFIjM3NHv8',
  HOLD_LOTS: process.env.REACT_APP_HOLD_LOTS_SPREADSHEET_ID || '1oBetbe44z2lUXngctvk3J31WBiWTv07NIgx5jFlylOs',
  CUTTING_SCANS: process.env.REACT_APP_CUTTING_SCANS_SPREADSHEET_ID || '1UU9P1IOjuUYpm3Ojx06V5L-r6K2QRb4mpFQmlGblgDA',
};

export const HOLD_LOTS_WEBHOOK_URL =
  process.env.REACT_APP_HOLD_LOTS_WEBHOOK_URL ||
  'https://script.google.com/macros/s/AKfycbwhfQMI2uYzDc-VBoYbk6McZCArUShh-3xNE_qwV4MFEb4C53dRTxPpifjRp2sHoVd_Yg/exec';



export const SHEET_NAMES = {
  JOB_ORDER: process.env.REACT_APP_JOB_ORDER_SHEET_NAME || 'JobOrder',
  INDEX: process.env.REACT_APP_INDEX_SHEET_NAME || 'Index',
  ISSUES: process.env.REACT_APP_ISSUES_SHEET_NAME || 'Issues',
  CUTTING: process.env.REACT_APP_CUTTING_SHEET_NAME || 'Cutting',
  RAWPACK: process.env.REACT_APP_RAWPACK_SHEET_NAME || 'RAWPACK',
  BARCODE: process.env.REACT_APP_BARCODE_SHEET_NAME || 'LotBarcodeData',
};

export const BACKEND_URL = process.env.REACT_APP_BACKEND_URL || 'http://localhost:5000';
export const BACKEND_API_BASE_URL = BACKEND_URL;

/**
 * Fetch spreadsheet range via Node.js Express Backend API
 * Automatically falls back to direct Google Sheets API if backend is unavailable.
 */
export const fetchSheetDataFromBackend = async (spreadsheetId, range, forceRefresh = false) => {
  const refreshParam = forceRefresh ? '&refresh=true' : '';
  const backendUrl = `${BACKEND_URL}/api/sheets/fetch?spreadsheetId=${encodeURIComponent(spreadsheetId)}&range=${encodeURIComponent(range)}${refreshParam}`;
  try {
    const res = await fetch(backendUrl);
    if (res.ok) {
      const data = await res.json();
      if (data.success && Array.isArray(data.values)) {
        return { ok: true, values: data.values, source: data.source };
      }
    }
  } catch (err) {
    console.warn(`Primary backend fetch failed for [${spreadsheetId} - ${range}]:`, err.message);
  }

  // Try localhost:5000 if primary backend is remote and failed
  if (BACKEND_URL !== 'http://localhost:5000') {
    try {
      const localUrl = `http://localhost:5000/api/sheets/fetch?spreadsheetId=${encodeURIComponent(spreadsheetId)}&range=${encodeURIComponent(range)}${refreshParam}`;
      const localRes = await fetch(localUrl);
      if (localRes.ok) {
        const localData = await localRes.json();
        if (localData.success && Array.isArray(localData.values)) {
          return { ok: true, values: localData.values, source: 'localhost_fallback' };
        }
      }
    } catch (localErr) {}
  }

  // Fallback to direct Google Sheets API if backend servers are unreachable
  try {
    const directUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}?key=${GOOGLE_API_KEY}`;
    const directRes = await fetch(directUrl);
    if (directRes.ok) {
      const directData = await directRes.json();
      return { ok: true, values: directData.values || [], source: 'direct_fallback' };
    }
  } catch (fallbackErr) {
    console.error(`Direct fallback fetch failed for [${spreadsheetId} - ${range}]:`, fallbackErr);
  }

  return { ok: false, values: [] };
};

