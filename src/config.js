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
  DORI: process.env.REACT_APP_DORI_SPREADSHEET_ID || '1LjwZqU26F0xwL1tEyps8txsM1qS8LLUuE-sy_4CQK6k',
  RAWPACK: process.env.REACT_APP_RAWPACK_SPREADSHEET_ID || '1xD8Uy1lUgvNTQ2RGRBI4ZjOrozbinUPRq2_UfIplP98',
  BARCODE: process.env.REACT_APP_BARCODE_SPREADSHEET_ID || '1dOCjNFwaAel5qun0_ZJVIGmREqjI76CJBBFIjM3NHv8',
};

export const SHEET_NAMES = {
  JOB_ORDER: process.env.REACT_APP_JOB_ORDER_SHEET_NAME || 'JobOrder',
  INDEX: process.env.REACT_APP_INDEX_SHEET_NAME || 'Index',
  ISSUES: process.env.REACT_APP_ISSUES_SHEET_NAME || 'Issues',
  CUTTING: process.env.REACT_APP_CUTTING_SHEET_NAME || 'Cutting',
  RAWPACK: process.env.REACT_APP_RAWPACK_SHEET_NAME || 'RAWPACK',
  BARCODE: process.env.REACT_APP_BARCODE_SHEET_NAME || 'LotBarcodeData',
};

export const BACKEND_URL = process.env.REACT_APP_BACKEND_URL || 'http://localhost:5000';

/**
 * Fetch spreadsheet range via Node.js Express Backend API
 * Automatically falls back to direct Google Sheets API if backend is unavailable.
 */
export const fetchSheetDataFromBackend = async (spreadsheetId, range) => {
  const backendUrl = `${BACKEND_URL}/api/sheets/fetch?spreadsheetId=${encodeURIComponent(spreadsheetId)}&range=${encodeURIComponent(range)}`;
  try {
    const res = await fetch(backendUrl);
    if (res.ok) {
      const data = await res.json();
      if (data.success && Array.isArray(data.values)) {
        return { ok: true, values: data.values, source: data.source };
      }
    }
  } catch (err) {
    console.warn(`Backend fetch failed for [${spreadsheetId} - ${range}], using direct Google API fallback:`, err.message);
  }

  // Fallback to direct Google Sheets API if backend server is unreachable
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

