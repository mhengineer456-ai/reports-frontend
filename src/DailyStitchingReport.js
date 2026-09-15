import React, { useState, useEffect, useCallback, useRef } from 'react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

const DailyStitchingUpdation = () => {
  // Data states
  const [data, setData] = useState([]);
  const [filteredData, setFilteredData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [isInitialLoad, setIsInitialLoad] = useState(true);
  const [supervisorOptions, setSupervisorOptions] = useState(['Monu']);
const [isLoadingSupervisors, setIsLoadingSupervisors] = useState(false);
const [loadProgress, setLoadProgress] = useState({ current: 0, total: 0 });
const [visibleRows, setVisibleRows] = useState(50);
const [scrollPosition, setScrollPosition] = useState(0);
const tableRef = useRef(null);
  
  // Cache management
  const dataCache = useRef(new Map());
  const lastFetchTime = useRef(new Map());
  const CACHE_DURATION = 5 * 60 * 1000; // 5 minutes cache
  const BATCH_SIZE = 2000; // Process data in batches for large datasets
  
  // Color pending tracking
  const [colorPendingLots, setColorPendingLots] = useState({}); // {lotNumber: ['shade1', 'shade2']}
  
  // Filters
  // In the filters state, add:
const [filters, setFilters] = useState({
  lotNumber: '',
  fabric: '',
  garmentType: '',
  style: '',
  brand: '',
  partyName: '',
  supervisor: 'Monu', // Default supervisor
  season: '',
  mwk: '',
  directStitching: '',
  challanHistory: '',
  wipStatus: '',
  completedStatus: '',
  lotStatus: 'Pending', // Default to Pending for Monu
  priority: '' // ADD THIS LINE
});
  
 const [filterOptions, setFilterOptions] = useState({
  fabric: [],
  garmentType: [],
  style: [],
  brand: [],
  partyName: [],
  supervisor: [],
  season: [],
  mwk: [],
  directStitching: [],
  challanHistory: [],
  wipStatus: [],
  completedStatus: [],
  lotStatus: ['Pending', 'Completed'],
  priority: [] // ADD THIS LINE
});

  // Google Sheets configuration
  const SPREADSHEET_ID = '1Hj3JeJEKB43aYYWv8gk2UhdU6BWuEQfCg5pBlTdBMNA';
  const API_KEY = 'AIzaSyAomDFBkOySlIxKWSKGHe6ATv9gvaBr7uk';
  const SHEET_NAME = 'Index';
  const CUTTING_SHEET_NAME = 'Cutting';
  const CUTTING_RANGE = `${CUTTING_SHEET_NAME}!A1:ZZ40000`;
  const INDEX_RANGE = `${SHEET_NAME}!A:AG`; // For getting index data
  
  const normalizeKey = (s = "") => {
    return String(s || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
  };

  const HEADERS = [
    'Sr.No',
    'Lot Number',
    'Garment Type',
    'Style',
    'Fabric',
    'BRAND',
    'Total PCS',
    'M/W/K',
    'Season',
    'PARTY NAME',
    'Direct Stitching',
    'Stitching Days',
    'Emb/Print Date',
    'WIP Status',
    'Completion Date',
    'Status'
  ];

  // Helper functions - memoized
  const normalizeText = useCallback((text) => {
    if (!text || text.trim() === '') return '';
    return text.trim().toLowerCase();
  }, []);

  const normalizeAndCapitalize = useCallback((text) => {
    if (!text || text.trim() === '') return '';
    
    const trimmed = text.trim().toLowerCase();
    return trimmed.split(' ')
      .map(word => word.charAt(0).toUpperCase() + word.slice(1))
      .join(' ');
  }, []);

  // Function to parse index row (same as CuttingStatsReport)
// Function to parse index row (same as CuttingStatsReport)
// Function to parse index row (same as CuttingStatsReport)
const parseIndexRow = useCallback((header, row) => {
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
  const savedAt = get("savedat");
  
  // ADD SHADES FIELD FOR COLOR PENDING CHECK
  const sizes = String(get("sizes") || "")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);

  const shades = String(get("shades") || "")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);

  // ADD PRIORITY FIELD - UPDATED TO HANDLE TYPO "Prioirty"
  const priority = get("priority") || get("prioirty") || "";

  return { lot, startRow, numRows, headerCols, fabric, garmentType, style, sizes, shades, savedAt, priority };
}, []);

  // Function to find header row index (same as CuttingStatsReport)
  const findHeaderRowIndex = useCallback((windowValues, expectedSizesNorm) => {
    const hasSizeToken = (rowSet) => expectedSizesNorm.some((sz) => rowSet.has(sz));
    for (let i = 0; i < windowValues.length; i++) {
      const row = windowValues[i] || [];
      const set = new Set(row.map((c) => normalizeKey(c)));
      const hasShadeHeader = set.has("color") || set.has("shade") || set.has("shades");
      if (hasShadeHeader && hasSizeToken(set)) return i;
    }
    for (let i = 0; i < windowValues.length; i++) {
      const row = windowValues[i] || [];
      const set = new Set(row.map((c) => normalizeKey(c)));
      let matches = 0;
      expectedSizesNorm.forEach((sz) => {
        if (set.has(sz)) matches++;
      });
      if (matches >= 2) return i;
    }
    return 0;
  }, []);

  // Function to compute pending shades (EXACTLY SAME as CuttingStatsReport)
  const computePendingShades = useCallback((windowValues, sizes = [], shades = []) => {
    if (!windowValues || windowValues.length === 0) {
      return new Set(shades.map(normalizeKey));
    }

    const normalizedSizes = Array.from(new Set((sizes || []).map(s => normalizeKey(s)).filter(Boolean)));
    const headerRowIdx = findHeaderRowIndex(windowValues, normalizedSizes);
    const header = windowValues[headerRowIdx] || [];

    const hIdx = {};
    header.forEach((h, i) => {
      const k = normalizeKey(h);
      if (k && !(k in hIdx)) hIdx[k] = i;
    });

    const shadeColIndex = hIdx["color"] ?? hIdx["shade"] ?? hIdx["shades"] ?? 0;

    // NON-SIZE COLUMNS - Exactly as in CuttingStatsReport
    const nonSizeColumns = new Set([
      "color", "shade", "shades", "cuttingtable", "cutting", "table", 
      "total", "totalpcs", "totals", "grandtotal", "sum", "lot", "style",
      "fabric", "garment", "partyname", "brand", "section", "season"
    ]);

    // EXACT SAME LOGIC as CuttingStatsReport
    let sizeColIndices = [];
    header.forEach((h, i) => {
      const normalizedHeader = normalizeKey(h);
      if (normalizedHeader && !nonSizeColumns.has(normalizedHeader)) {
        sizeColIndices.push(i);
      }
    });

    // If no columns found with the above logic, try matching against expected sizes
    if (sizeColIndices.length === 0) {
      normalizedSizes.forEach((ns) => {
        if (ns in hIdx) sizeColIndices.push(hIdx[ns]);
      });
      
      if (sizeColIndices.length === 0) {
        const ct = hIdx["cuttingtable"];
        if (ct != null && ct >= 0) {
          const guessStart = ct + 1;
          const guessed = [];
          for (let k = 0; k < normalizedSizes.length; k++) guessed.push(guessStart + k);
          sizeColIndices = Array.from(new Set(guessed.filter((g) => g < header.length)));
        }
      }
    }

    if (sizeColIndices.length === 0) {
      return new Set(shades.map(normalizeKey));
    }

    const shadeStats = new Map(); // Track shade status: "found-with-data", "found-all-zero", or "not-found"
    
    for (let r = headerRowIdx + 1; r < windowValues.length; r++) {
      const row = windowValues[r] || [];
      const rawShade = String(row[shadeColIndex] || "").trim();
      const shadeKey = normalizeKey(rawShade);

      if (!shadeKey || shadeKey === "total" || shadeKey === "totals" || shadeKey === "grandtotal") {
        continue;
      }

      let hasPositiveData = false;
      let hasAnyData = false;
      let totalForThisRow = 0;
      
      // Check this shade's data
      sizeColIndices.forEach((c) => {
        const raw = row[c];
        if (raw != null && raw !== "") {
          hasAnyData = true;
          const n = parseFloat(String(raw).replace(/,/g, ""));
          if (!isNaN(n)) {
            totalForThisRow += n;
            if (n > 0) {
              hasPositiveData = true;
            }
          }
        }
      });

      // Update shade status
      if (hasAnyData) {
        if (hasPositiveData) {
          // Shade has some positive quantity (>0)
          shadeStats.set(shadeKey, "found-with-data");
        } else {
          // Shade exists but all quantities are 0 or empty
          // Only set if not already marked as "found-with-data"
          if (!shadeStats.has(shadeKey) || shadeStats.get(shadeKey) === "not-found") {
            shadeStats.set(shadeKey, "found-all-zero");
          }
        }
      } else {
        // Shade row exists but has no data in size columns
        if (!shadeStats.has(shadeKey)) {
          shadeStats.set(shadeKey, "found-no-data");
        }
      }
    }

    // Determine pending shades
    const pendingShadeKeys = new Set();
    const expectedShadeKeys = (shades || []).map(sh => normalizeKey(sh));

    expectedShadeKeys.forEach((shadeKey) => {
      const status = shadeStats.get(shadeKey);
      
      // Shade is PENDING only if:
      // 1. Not found at all (status is undefined) - shade doesn't exist in cutting sheet
      // 2. Found but has no data (status is "found-no-data") - shade exists but all cells are empty
      
      // Shade is NOT PENDING (cancelled/processed) if:
      // 1. Has positive data (status is "found-with-data") - shade has >0 quantity
      // 2. Has all zeros (status is "found-all-zero") - shade exists but all quantities are 0
      
      if (!status || status === "found-no-data") {
        // Shade not found OR found with no data → PENDING
        pendingShadeKeys.add(shadeKey);
      }
      // If status is "found-with-data" or "found-all-zero" → NOT PENDING (considered cancelled/processed)
    });

    return pendingShadeKeys;
  }, [findHeaderRowIndex]);

  // Function to slice cutting matrix (same as CuttingStatsReport)
  const sliceCuttingMatrix = useCallback((bigValues, startRow, numRows) => {
    if (!Array.isArray(bigValues) || bigValues.length === 0) return [];
    if (!(startRow > 0 && numRows > 0)) return [];
    const r0 = Math.max(0, startRow - 1);
    const r1 = Math.min(bigValues.length - 1, r0 + numRows - 1);
    return bigValues.slice(r0, r1 + 1);
  }, []);

  // Function to fetch index data (for shades information)
  const fetchIndexData = async (signal) => {
    try {
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(INDEX_RANGE)}?key=${API_KEY}`;
      const response = await fetch(url, { signal });
      
      if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
      }
      
      const result = await response.json();
      return result.values || [];
    } catch (err) {
      console.error('Error fetching index data:', err);
      return [];
    }
  };

  // Function to check color pending for all lots
// Function to check color pending for all lots
// Function to check color pending for all lots
const checkColorPendingForLots = useCallback(async (cuttingData, indexData) => {
  if (!cuttingData || cuttingData.length === 0 || !indexData || indexData.length === 0) {
    return {};
  }
  
  const indexHeader = indexData[0];
  const pendingMap = {};
  
  // Process each row in index sheet
  for (let i = 1; i < indexData.length; i++) {
    const entry = parseIndexRow(indexHeader, indexData[i]);
    if (!entry) continue;
    
    const { lot, startRow, numRows, sizes, shades, priority } = entry;
    
    // Store the priority for every lot (even if no shades)
    pendingMap[lot] = { 
      pendingColors: [],
      priority: priority || ''
    };
    
    // Only check for color pending if we have shades data
    if (shades && shades.length > 0) {
      const window = sliceCuttingMatrix(cuttingData, startRow, numRows);
      const pendingShadeKeys = computePendingShades(window, sizes, shades);
      
      if (pendingShadeKeys.size > 0) {
        const shadeKeyToOriginal = new Map((shades || []).map((sh) => [normalizeKey(sh), sh]));
        const pendingList = Array.from(pendingShadeKeys).map((k) => shadeKeyToOriginal.get(k) || k);
        pendingMap[lot].pendingColors = pendingList;
      }
    }
  }
  
  // Debug: Check priorities found
  console.log('Priorities found in checkColorPendingForLots:', 
    Object.entries(pendingMap)
      .filter(([lot, data]) => data.priority)
      .map(([lot, data]) => ({ lot, priority: data.priority }))
  );
  
  return pendingMap;
}, [parseIndexRow, sliceCuttingMatrix, computePendingShades]);
  // Function to get completion date from completedStatus
  const getCompletionDate = useCallback((completedStatus) => {
    if (!completedStatus || completedStatus.trim() === '') {
      return null;
    }
    
    try {
      if (completedStatus.startsWith('[')) {
        const statusArray = JSON.parse(completedStatus);
        
        if (Array.isArray(statusArray) && statusArray.length > 0) {
          // Find the completion status
          const completeEntry = statusArray.find(entry => 
            entry.status && normalizeText(entry.status).includes('complete')
          );
          
          if (completeEntry && completeEntry.timestamp) {
            return new Date(completeEntry.timestamp);
          }
        }
      }
      
      return null;
    } catch (error) {
      console.error('Error parsing completion date:', error);
      return null;
    }
  }, [normalizeText]);

  // Function to format completion date
  const getCompletionDateFormatted = useCallback((completedStatus) => {
    const date = getCompletionDate(completedStatus);
    if (!date) return '';
    
    return date.toLocaleDateString('en-US', {
      year: 'numeric',
      month: 'short',
      day: 'numeric'
    });
  }, [getCompletionDate]);

  const isLotCompleted = useCallback((completedStatus) => {
    if (!completedStatus || completedStatus.trim() === '') {
      return false;
    }
    
    try {
      if (completedStatus.startsWith('[')) {
        const statusArray = JSON.parse(completedStatus);
        
        if (Array.isArray(statusArray) && statusArray.length > 0) {
          const latestStatus = statusArray[statusArray.length - 1];
          if (latestStatus.status && normalizeText(latestStatus.status).includes('complete')) {
            return true;
          }
        }
      } else if (typeof completedStatus === 'string') {
        return normalizeText(completedStatus).includes('complete');
      }
      
      return false;
    } catch (error) {
      console.error('Error parsing Completed Status:', error);
      if (typeof completedStatus === 'string') {
        return normalizeText(completedStatus).includes('complete');
      }
      return false;
    }
  }, [normalizeText]);

  const getCompletedStatusText = useCallback((completedStatus) => {
    if (!completedStatus || completedStatus.trim() === '') {
      return 'N/A';
    }
    
    try {
      if (completedStatus.startsWith('[')) {
        const statusArray = JSON.parse(completedStatus);
        
        if (Array.isArray(statusArray) && statusArray.length > 0) {
          const latestStatus = statusArray[statusArray.length - 1];
          if (latestStatus.status) {
            return latestStatus.status;
          }
        }
      }
      
      return completedStatus;
    } catch (error) {
      console.error('Error getting completed status text:', error);
      return completedStatus;
    }
  }, []);

  const getEmbPrintDate = useCallback((challanHistory) => {
    if (!challanHistory || typeof challanHistory !== 'string' || challanHistory.trim() === '') {
      return '-';
    }
    
    const trimmedHistory = challanHistory.trim();
    
    try {
      if (trimmedHistory.startsWith('[') && trimmedHistory.endsWith(']')) {
        const historyArray = JSON.parse(trimmedHistory);
        
        if (!Array.isArray(historyArray) || historyArray.length === 0) {
          return '-';
        }
        
        let latestDate = null;
        
        historyArray.forEach(entry => {
          if (entry && entry.embUpdatedAt && typeof entry.embUpdatedAt === 'string') {
            const currentDate = new Date(entry.embUpdatedAt);
            if (!isNaN(currentDate.getTime())) {
              if (!latestDate || currentDate > latestDate) {
                latestDate = currentDate;
              }
            }
          }
        });
        
        if (latestDate) {
          return latestDate.toLocaleDateString('en-US', {
            year: 'numeric',
            month: 'short',
            day: 'numeric'
          });
        }
      }
      
      return '-';
    } catch (error) {
      console.error('Error parsing challan history in getEmbPrintDate:', error);
      return '-';
    }
  }, []);

  const fetchCuttingData = async (signal) => {
    try {
      const sheetNameEncoded = encodeURIComponent(CUTTING_RANGE);
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${sheetNameEncoded}?key=${API_KEY}`;
      
      const response = await fetch(url, { signal });
      
      if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
      }
      
      const result = await response.json();
      return result.values || [];
    } catch (err) {
      console.error('Error fetching cutting data:', err);
      return [];
    }
  };

  // Helper function to parse cutting data and calculate PCS for a lot
  const calculateTotalPCS = (cuttingData, startRow, numRows, sizes = []) => {
    if (!cuttingData || cuttingData.length === 0) return 0;
    if (!(startRow > 0 && numRows > 0)) return 0;
    
    // Slice the cutting matrix for the specific lot
    const r0 = Math.max(0, startRow - 1);
    const r1 = Math.min(cuttingData.length - 1, r0 + numRows - 1);
    const windowValues = cuttingData.slice(r0, r1 + 1);
    
    if (windowValues.length === 0) return 0;
    
    const normalizedSizes = Array.from(new Set((sizes || []).map(s => normalizeKey(s)).filter(Boolean)));
    
    // Find header row
    const findHeaderRowIndex = (windowValues, expectedSizesNorm) => {
      const hasSizeToken = (rowSet) => expectedSizesNorm.some((sz) => rowSet.has(sz));
      
      for (let i = 0; i < windowValues.length; i++) {
        const row = windowValues[i] || [];
        const set = new Set(row.map((c) => normalizeKey(c)));
        const hasShadeHeader = set.has("color") || set.has("shade") || set.has("shades");
        if (hasShadeHeader && hasSizeToken(set)) return i;
      }
      
      for (let i = 0; i < windowValues.length; i++) {
        const row = windowValues[i] || [];
        const set = new Set(row.map((c) => normalizeKey(c)));
        let matches = 0;
        expectedSizesNorm.forEach((sz) => {
          if (set.has(sz)) matches++;
        });
        if (matches >= 2) return i;
      }
      
      return 0;
    };
    
    const headerRowIdx = findHeaderRowIndex(windowValues, normalizedSizes);
    const header = windowValues[headerRowIdx] || [];
    
    const hIdx = {};
    header.forEach((h, i) => {
      const k = normalizeKey(h);
      if (k && !(k in hIdx)) hIdx[k] = i;
    });
    
    const nonSizeColumns = new Set([
      "color", "shade", "shades", "cuttingtable", "cutting", "table", 
      "total", "totalpcs", "totals", "grandtotal", "sum", "lot", "style",
      "fabric", "garment", "partyname", "brand", "section", "season"
    ]);
    
    let sizeColIndices = [];
    header.forEach((h, i) => {
      const normalizedHeader = normalizeKey(h);
      if (normalizedHeader && !nonSizeColumns.has(normalizedHeader)) {
        sizeColIndices.push(i);
      }
    });
    
    if (sizeColIndices.length === 0) {
      normalizedSizes.forEach((ns) => {
        if (ns in hIdx) sizeColIndices.push(hIdx[ns]);
      });
      
      if (sizeColIndices.length === 0) {
        const ct = hIdx["cuttingtable"];
        if (ct != null && ct >= 0) {
          const guessStart = ct + 1;
          const guessed = [];
          for (let k = 0; k < normalizedSizes.length; k++) guessed.push(guessStart + k);
          sizeColIndices = Array.from(new Set(guessed.filter((g) => g < header.length)));
        }
      }
    }
    
    if (sizeColIndices.length === 0) return 0;
    
    let totalQty = 0;
    
    for (let r = headerRowIdx + 1; r < windowValues.length; r++) {
      const row = windowValues[r] || [];
      const rawShade = String(row[hIdx["color"] || hIdx["shade"] || 0] || "").trim();
      const shadeKey = normalizeKey(rawShade);
      
      if (!shadeKey || shadeKey === "total" || shadeKey === "totals" || shadeKey === "grandtotal") continue;
      
      let rowTotal = 0;
      
      sizeColIndices.forEach((c) => {
        const raw = row[c];
        if (raw != null && raw !== "") {
          const n = parseFloat(String(raw).replace(/,/g, ""));
          if (!isNaN(n) && n > 0) {
            rowTotal += n;
          }
        }
      });
      
      totalQty += rowTotal;
    }
    
    return totalQty;
  };
  const isToday = useCallback((dateString) => {
  if (!dateString) return false;
  
  try {
    const date = new Date(dateString);
    if (isNaN(date.getTime())) return false;
    
    const today = new Date();
    return date.getDate() === today.getDate() &&
           date.getMonth() === today.getMonth() &&
           date.getFullYear() === today.getFullYear();
  } catch {
    return false;
  }
}, []);

const getLatestWipRemarks = useCallback((wipStatus) => {
  if (!wipStatus || wipStatus.trim() === '') {
    return 'Not Updated';
  }
  
  try {
    // If it's a simple string (not an array), check if it has today's date embedded
    if (typeof wipStatus === 'string' && !wipStatus.startsWith('[')) {
      // Try to extract date from the string if it contains one
      const dateMatch = wipStatus.match(/\d{1,4}[-/]\d{1,2}[-/]\d{1,4}/);
      if (dateMatch) {
        return isToday(dateMatch[0]) ? wipStatus : 'Not Updated';
      }
      return 'Not Updated'; // No date found in string
    }
    
    // Parse the JSON array
    const statusArray = JSON.parse(wipStatus);
    
    if (!Array.isArray(statusArray) || statusArray.length === 0) {
      return 'Not Updated';
    }
    
    // Sort by timestamp descending (newest first)
    const sortedStatuses = [...statusArray].sort((a, b) => {
      const dateA = new Date(a.timestamp).getTime();
      const dateB = new Date(b.timestamp).getTime();
      return dateB - dateA;
    });
    
    // Find the first entry with today's date
    const todayEntry = sortedStatuses.find(entry => 
      entry.timestamp && isToday(entry.timestamp)
    );
    
    if (todayEntry) {
      // Return remarks if available, otherwise return status
      if (todayEntry.remarks && todayEntry.remarks.trim() !== '') {
        return todayEntry.remarks;
      } else if (todayEntry.status && todayEntry.status.trim() !== '') {
        return todayEntry.status;
      }
    }
    
    return 'Not Updated'; // No entry with today's date found
  } catch (error) {
    console.error('Error parsing WIP Status:', error);
    return 'Not Updated';
  }
}, [isToday]);

  // Calculate stitching days - memoized
  const calculateStitchingDays = useCallback((dateOfIssue) => {
    if (!dateOfIssue || dateOfIssue.trim() === '') return 0;
    
    try {
      const issueDate = new Date(dateOfIssue);
      if (isNaN(issueDate.getTime())) {
        const parts = dateOfIssue.split(/[\/\-]/);
        if (parts.length === 3) {
          const newDate = new Date(parts[2], parts[1] - 1, parts[0]);
          if (!isNaN(newDate.getTime())) {
            const today = new Date();
            const timeDiff = today.getTime() - newDate.getTime();
            return Math.floor(timeDiff / (1000 * 3600 * 24));
          }
        }
        return 0;
      }
      
      const today = new Date();
      const timeDiff = today.getTime() - issueDate.getTime();
      return Math.floor(timeDiff / (1000 * 3600 * 24));
    } catch {
      return 0;
    }
  }, []);

  // Get stitching days color - memoized
  const getStitchingDaysColor = useCallback((days) => {
    if (days <= 6) return '#10b981'; // Green
    if (days <= 15) return '#f59e0b'; // Yellow
    return '#ef4444'; // Red
  }, []);

  const getStitchingDaysTextColor = useCallback((days) => {
    if (days <= 6) return '#ffffff'; // White text on green
    if (days <= 15) return '#000000'; // Black text on yellow
    return '#ffffff'; // White text on red
  }, []);

  // Format date - memoized
  const formatDate = useCallback((dateString) => {
    if (!dateString || dateString.trim() === '') return 'N/A';
    
    try {
      const date = new Date(dateString);
      
      if (isNaN(date.getTime())) {
        // Try different date formats
        const parts = dateString.split(/[\/\-]/);
        if (parts.length === 3) {
          const newDate = new Date(parts[2], parts[1] - 1, parts[0]);
          if (!isNaN(newDate.getTime())) {
            return newDate.toLocaleDateString('en-US', {
              year: 'numeric',
              month: 'short',
              day: 'numeric'
            });
          }
        }
        return dateString;
      }
      
      return date.toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'short',
        day: 'numeric'
      });
    } catch {
      return dateString;
    }
  }, []);

  // Process data in batches to prevent UI freeze
  const processDataInBatches = useCallback(async (rows, processFn) => {
    const batches = [];
    for (let i = 0; i < rows.length; i += BATCH_SIZE) {
      batches.push(rows.slice(i, i + BATCH_SIZE));
    }
    
    const results = [];
    for (let i = 0; i < batches.length; i++) {
      // Allow UI to update between batches
      await new Promise(resolve => setTimeout(resolve, 0));
      const batchResults = batches[i].map(processFn);
      results.push(...batchResults);
    }
    
    return results;
  }, []);

  // Extract filter options from processed data
 // In the extractFilterOptions function, add priority extraction:
// In the extractFilterOptions function, add priority extraction:
const extractFilterOptions = useCallback((data) => {
  const options = {
    fabric: [],
    garmentType: [],
    style: [],
    brand: [],
    partyName: [],
    supervisor: [],
    season: [],
    mwk: [],
    directStitching: [],
    challanHistory: [],
    wipStatus: [],
    completedStatus: [],
    lotStatus: ['Pending', 'Completed'],
    priority: []
  };

  // Use Set for faster lookups
  const optionSets = Object.keys(options).reduce((acc, key) => {
    acc[key] = new Set();
    return acc;
  }, {});

  // Process in batches
  const processRow = (item) => {
    Object.keys(options).forEach(key => {
      if (key === 'wipStatus') {
        const latestRemarks = getLatestWipRemarks(item[key]);
        // Only add to filter options if it has a value (meaning it's from today)
        if (latestRemarks && latestRemarks.trim() !== '') {
          const normalizedValue = normalizeAndCapitalize(latestRemarks);
          optionSets[key].add(normalizedValue);
        }
      } else if (key === 'completedStatus') {
        const statusText = getCompletedStatusText(item[key]);
        if (statusText && statusText.trim() !== '' && statusText !== 'N/A') {
          const normalizedValue = normalizeAndCapitalize(statusText);
          optionSets[key].add(normalizedValue);
        }
      } else if (key === 'lotStatus') {
        return;
      } else if (key === 'priority') {
        const priorityValue = item[key];
        if (priorityValue && typeof priorityValue === 'string' && priorityValue.trim() !== '') {
          const normalizedValue = normalizeAndCapitalize(priorityValue.trim());
          if (normalizedValue) {
            optionSets[key].add(normalizedValue);
          }
        }
      } else if (item[key] && item[key].trim() !== '') {
        const normalizedValue = normalizeAndCapitalize(item[key]);
        optionSets[key].add(normalizedValue);
      }
    });
  };

  // Process all rows
  data.forEach(processRow);

  // Convert Sets to sorted arrays
  Object.keys(options).forEach(key => {
    if (key !== 'lotStatus') {
      options[key] = Array.from(optionSets[key]).sort();
    }
  });

  setFilterOptions(options);
}, [getLatestWipRemarks, getCompletedStatusText, normalizeAndCapitalize]);


  // Generate cache key based on filters
  const getCacheKey = useCallback((supervisor = '') => {
    return supervisor || 'all';
  }, []);

  // Check if cache is valid
  const isCacheValid = useCallback((cacheKey) => {
    const lastFetch = lastFetchTime.current.get(cacheKey);
    if (!lastFetch) return false;
    
    const now = Date.now();
    return (now - lastFetch) < CACHE_DURATION;
  }, []);

  // Fetch data for specific supervisor or all data
// REPLACE the entire fetchDataForSupervisor function with this:

const fetchDataForSupervisor = useCallback(async (supervisor = '', signal) => {
  const cacheKey = getCacheKey(supervisor);
  
  // Check cache first
  if (isCacheValid(cacheKey) && dataCache.current.has(cacheKey)) {
    console.log('Using cached data for:', supervisor);
    return dataCache.current.get(cacheKey);
  }

  try {
    setLoading(true);
    setLoadProgress({ current: 5, total: 100 });
    
    // 1. Fetch cutting data (once for all lots)
    setLoadProgress({ current: 10, total: 100 });
    const cuttingData = await fetchCuttingData(signal);
    console.log('Cutting data fetched:', cuttingData.length, 'rows');
    
    // 2. Fetch index data for color pending check
    setLoadProgress({ current: 20, total: 100 });
    const indexData = await fetchIndexData(signal);
    console.log('Index data fetched:', indexData.length, 'rows');
    
    // 3. Check color pending for all lots AND get priority
    setLoadProgress({ current: 30, total: 100 });
    const colorPendingMap = await checkColorPendingForLots(cuttingData, indexData);
    setColorPendingLots(colorPendingMap);
    console.log('Color pending map created:', Object.keys(colorPendingMap).length, 'lots');
    
    // 4. Create priority lookup map from index data
    const indexHeader = indexData[0] || [];
    const priorityLookup = {};
    
    for (let i = 1; i < Math.min(indexData.length, 1000); i++) {
      const entry = parseIndexRow(indexHeader, indexData[i]);
      if (entry && entry.lot) {
        priorityLookup[entry.lot] = entry.priority || '';
      }
    }
    console.log('Priority lookup created:', Object.keys(priorityLookup).length, 'lots');
    
    // 5. Fetch main data in chunks
    setLoadProgress({ current: 40, total: 100 });
    const allData = [];
    let offset = 0;
    const CHUNK_SIZE = 1000;
    let totalRows = 5000; // Default estimate
    
    try {
      // Try to get approximate row count
      const countUrl = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(SHEET_NAME)}!A1:A?key=${API_KEY}`;
      const countResponse = await fetch(countUrl, { signal });
      const countResult = await countResponse.json();
      totalRows = countResult.values ? countResult.values.length : 5000;
      console.log('Estimated total rows:', totalRows);
    } catch (err) {
      console.log('Using default row estimate:', totalRows);
    }
    
    let headers = null;
    let columnMapping = null;
    
    // Fetch and process data in chunks
    while (offset < totalRows) {
      if (signal && signal.aborted) {
        console.log('Fetch aborted');
        break;
      }
      
      const range = `${SHEET_NAME}!A${offset + 1}:AG${offset + CHUNK_SIZE}`;
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(range)}?key=${API_KEY}`;
      
      console.log(`Fetching chunk ${offset}-${offset + CHUNK_SIZE}`);
      
      const response = await fetch(url, { signal });
      if (!response.ok) {
        console.log('Failed to fetch chunk:', response.status);
        break;
      }
      
      const result = await response.json();
      const chunkData = result.values || [];
      
      if (chunkData.length === 0) {
        console.log('No more data');
        break;
      }
      
      // Get headers from first chunk
      if (offset === 0) {
        headers = chunkData[0];
        columnMapping = createColumnMapping(headers);
        console.log('Headers found:', headers.length, 'columns');
      }
      
      // Process chunk (skip header row for first chunk)
      const startRow = offset === 0 ? 1 : 0;
      const chunkToProcess = startRow > 0 ? chunkData.slice(startRow) : chunkData;
      
      if (chunkToProcess.length > 0) {
        const processedChunk = await processChunkData(
          chunkToProcess, 
          columnMapping,
          cuttingData, 
          colorPendingMap, 
          priorityLookup
        );
        
        console.log(`Processed ${processedChunk.length} rows from chunk`);
        
        // Filter for supervisor if specified
        if (supervisor && supervisor.trim() !== '') {
          const filtered = processedChunk.filter(item => {
            if (!item.supervisor) return false;
            const itemSupervisor = normalizeText(item.supervisor);
            const targetSupervisor = normalizeText(supervisor);
            return itemSupervisor === targetSupervisor;
          });
          allData.push(...filtered);
          console.log(`Filtered to ${filtered.length} rows for supervisor ${supervisor}`);
        } else {
          allData.push(...processedChunk);
        }
      }
      
      offset += CHUNK_SIZE;
      const progress = 40 + Math.floor((offset / totalRows) * 50);
      setLoadProgress({ current: Math.min(progress, 95), total: 100 });
      
      // Yield to UI to prevent freezing (shorter delay)
      await new Promise(resolve => setTimeout(resolve, 5));
    }
    
    // Filter out invalid rows
    const validData = allData.filter(item => {
      const hasLotNumber = item.lotNumber && item.lotNumber.trim() !== '';
      const hasDateOfIssue = item.dateOfIssue && item.dateOfIssue.trim() !== '';
      const hasSupervisor = item.supervisor && item.supervisor.trim() !== '';
      
      return hasLotNumber && (hasDateOfIssue || hasSupervisor);
    });
    
    console.log(`Total valid data: ${validData.length} rows`);
    
    // Update cache
    setLoadProgress({ current: 98, total: 100 });
    dataCache.current.set(cacheKey, validData);
    lastFetchTime.current.set(cacheKey, Date.now());
    
    console.log(`Cached data for ${supervisor}: ${validData.length} rows`);
    
    setLoadProgress({ current: 100, total: 100 });
    
    return validData;
  } catch (err) {
    console.error('Error fetching data:', err);
    if (err.name !== 'AbortError') {
      setError(`Fetch error: ${err.message}`);
    }
    throw err;
  } finally {
    setLoading(false);
    // Clear progress after a delay
    setTimeout(() => setLoadProgress({ current: 0, total: 0 }), 500);
  }
}, [getCacheKey, isCacheValid, normalizeText, checkColorPendingForLots, parseIndexRow]);

// Add these functions AFTER fetchDataForSupervisor:

// Helper function to create column mapping
const createColumnMapping = useCallback((headers) => {
  const mapping = {};
  
  const columnDefinitions = [
    { keys: ['Lot Number', 'LotNumber', 'lotNumber'], target: 'lotNumber' },
    { keys: ['Fabric', 'FABRIC'], target: 'fabric' },
    { keys: ['Garment Type', 'GarmentType'], target: 'garmentType' },
    { keys: ['Style'], target: 'style' },
    { keys: ['BRAND', 'Brand'], target: 'brand' },
    { keys: ['PARTY NAME', 'Party Name', 'PartyName'], target: 'partyName' },
    { keys: ['Date of Issue', 'DateOfIssue'], target: 'dateOfIssue' },
    { keys: ['Supervisor'], target: 'supervisor' },
    { keys: ['SEASON', 'Season'], target: 'season' },
    { keys: ['DIRECT STITCHING', 'Direct Stitching'], target: 'directStitching' },
    { keys: ['CHALLAN HISTORY', 'Challan History'], target: 'challanHistory' },
    { keys: ['M/W/K', 'MWK'], target: 'mwk' },
    { keys: ['WIP Status', 'WIPStatus'], target: 'wipStatus' },
    { keys: ['Completed Status', 'CompletedStatus'], target: 'completedStatus' },
    { keys: ['StartRow', 'Start Row'], target: 'startRow' },
    { keys: ['NumRows', 'Num Rows'], target: 'numRows' },
    { keys: ['Manpower', 'MANPOWER', 'Man Power', 'manpower'], target: 'manpower' },
    { keys: ['Sizes'], target: 'sizes' },
    { keys: ['Priority', 'PRIORITY', 'Prioirty'], target: 'priority' }
  ];
  
  columnDefinitions.forEach(definition => {
    for (const key of definition.keys) {
      const index = headers.findIndex(
        header => header && normalizeKey(header) === normalizeKey(key)
      );
      if (index !== -1) {
        mapping[definition.target] = index;
        break;
      }
    }
  });
  
  return mapping;
}, [normalizeKey]);

// Process chunk data function
const processChunkData = useCallback(async (
  chunkData, 
  columnMapping,
  cuttingData, 
  colorPendingMap, 
  priorityLookup
) => {
  if (!chunkData || chunkData.length === 0) return [];
  
  const BATCH_SIZE = 200;
  const results = [];
  
  for (let i = 0; i < chunkData.length; i += BATCH_SIZE) {
    const batchEnd = Math.min(i + BATCH_SIZE, chunkData.length);
    const batch = chunkData.slice(i, batchEnd);
    
    const batchResults = batch.map(row => {
      const item = {};
      
      // Map all columns using the column mapping
      Object.entries(columnMapping).forEach(([targetKey, columnIndex]) => {
        if (columnIndex !== undefined && row[columnIndex] !== undefined) {
          const value = row[columnIndex];
          
          // Handle special parsing for certain fields
          switch (targetKey) {
            case 'startRow':
              const startRowVal = parseInt(value);
              item[targetKey] = isNaN(startRowVal) ? 0 : startRowVal;
              break;
            case 'numRows':
              const numRows = parseInt(value);
              item[targetKey] = isNaN(numRows) ? 0 : numRows;
              break;
            case 'sizes':
              item[targetKey] = value ? 
                value.split(',')
                  .map(s => s.trim())
                  .filter(Boolean) : [];
              break;
            default:
              item[targetKey] = value || '';
              break;
          }
        } else {
          // Set defaults for missing columns
          if (targetKey === 'startRow' || targetKey === 'numRows') {
            item[targetKey] = 0;
          } else if (targetKey === 'sizes') {
            item[targetKey] = [];
          } else {
            item[targetKey] = '';
          }
        }
      });
      
      // Calculate total PCS from cutting sheet
      if (cuttingData.length > 0 && item.startRow > 0 && item.numRows > 0) {
        try {
          const totalPCS = calculateTotalPCS(cuttingData, item.startRow, item.numRows, item.sizes);
          item['totalPCS'] = totalPCS;
        } catch (error) {
          console.error(`Error calculating PCS for lot ${item.lotNumber}:`, error);
          item['totalPCS'] = 0;
        }
      } else {
        item['totalPCS'] = 0;
      }
      
      // Get lot number
      const lotNumber = item.lotNumber || '';
      
      // Set color pending information
      if (lotNumber && colorPendingMap[lotNumber]) {
        item['hasColorPending'] = colorPendingMap[lotNumber].pendingColors.length > 0;
        item['pendingColors'] = colorPendingMap[lotNumber].pendingColors;
      } else {
        item['hasColorPending'] = false;
        item['pendingColors'] = [];
      }
      
      // Set priority
      if (lotNumber && priorityLookup[lotNumber]) {
        item['priority'] = priorityLookup[lotNumber];
      } else if (lotNumber && colorPendingMap[lotNumber] && colorPendingMap[lotNumber].priority) {
        item['priority'] = colorPendingMap[lotNumber].priority;
      } else if (item.priority) {
        item['priority'] = item.priority;
      } else {
        item['priority'] = '';
      }
      
      return item;
    });
    
    results.push(...batchResults);
    
    // Yield to UI between batches
    if (i % (BATCH_SIZE * 3) === 0) {
      await new Promise(resolve => setTimeout(resolve, 0));
    }
  }
  
  return results;
}, [calculateTotalPCS, normalizeKey]);

  // Apply filters to data
// In the applyFilters function, add priority filter handling:
const applyFilters = useCallback((currentFilters, currentData) => {
  if (currentData.length === 0) {
    setFilteredData([]);
    return;
  }

  // Use requestAnimationFrame for smooth filtering
  requestAnimationFrame(() => {
    let filtered = [...currentData];
    
    filtered = filtered.filter(item => {
      return Object.entries(currentFilters).every(([key, value]) => {
        if (!value || value.trim() === '') {
          return true;
        }
        
        const filterValue = normalizeText(value);
        
        // Handle lotStatus filter
        if (key === 'lotStatus') {
          if (filterValue === 'pending') {
            return !isLotCompleted(item.completedStatus);
          } else if (filterValue === 'completed') {
            return isLotCompleted(item.completedStatus);
          }
          return true;
        }
        
        // Handle lotNumber filter
        if (key === 'lotNumber') {
          const itemValue = item[key] ? normalizeText(item[key]) : '';
          return itemValue.includes(filterValue);
        }
        
        // Handle completedStatus filter
        if (key === 'completedStatus') {
          const statusText = getCompletedStatusText(item[key]);
          return normalizeText(statusText).includes(filterValue);
        }
        
        // Handle challanHistory filter
        if (key === 'challanHistory') {
          const embPrintDate = getEmbPrintDate(item[key]);
          return normalizeText(embPrintDate).includes(filterValue);
        }
        
        // Handle wipStatus filter - ONLY check today's status
        if (key === 'wipStatus') {
          const latestRemarks = getLatestWipRemarks(item[key]);
          // If there's no today's status, it won't match any filter
          return latestRemarks && normalizeText(latestRemarks).includes(filterValue);
        }
        
        // Handle priority filter
        if (key === 'priority') {
          const itemValue = item[key] ? normalizeText(item[key]) : '';
          return itemValue.includes(filterValue);
        }
        
        const itemValue = item[key] ? normalizeText(item[key]) : '';
        return itemValue.includes(filterValue);
      });
    });

    setFilteredData(filtered);
  });
}, [isLotCompleted, getCompletedStatusText, getEmbPrintDate, getLatestWipRemarks, normalizeText]);
  // Initial load - fetch ALL of Monu's data (including completed)
  useEffect(() => {
    const loadInitialData = async () => {
      try {
        setLoading(true);
        setIsInitialLoad(true);
        
        const monuData = await fetchDataForSupervisor('Monu');
        setData(monuData);
        
        // Show only PENDING lots by default for Monu
        const pendingData = monuData.filter(item => !isLotCompleted(item.completedStatus));
        setFilteredData(pendingData);
        
        // Extract filter options from Monu's data only
        extractFilterOptions(monuData);
        
      } catch (err) {
        setError(`Failed to load initial data: ${err.message}`);
      } finally {
        setLoading(false);
        setIsInitialLoad(false);
      }
    };

    loadInitialData();
  }, [fetchDataForSupervisor, extractFilterOptions, isLotCompleted]);

  // Handle filter changes
// REPLACE the entire handleFilterChange function with this:

const handleFilterChange = useCallback(async (e) => {
  const { name, value } = e.target;
  const newFilters = { ...filters, [name]: value };
  
  // Update filters immediately for UI responsiveness
  setFilters(newFilters);

  // If supervisor changed
  if (name === 'supervisor') {
    try {
      setLoading(true);
      setError(null);
      
      console.log(`Changing supervisor from ${filters.supervisor} to ${value}`);
      
      // Clear cache for the new supervisor only
      const newCacheKey = getCacheKey(value);
      dataCache.current.delete(newCacheKey);
      lastFetchTime.current.delete(newCacheKey);
      
      // Add timeout protection
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 60000);
      
      try {
        const supervisorData = await fetchDataForSupervisor(value, controller.signal);
        clearTimeout(timeoutId);
        
        if (!supervisorData) {
          throw new Error('No data returned from fetch');
        }
        
        console.log(`Successfully loaded ${supervisorData.length} rows for ${value}`);
        
        setData(supervisorData);
        
        // Apply filters based on supervisor
        if (value === 'Monu') {
          const pendingData = supervisorData.filter(item => !isLotCompleted(item.completedStatus));
          setFilteredData(pendingData);
          newFilters.lotStatus = 'Pending';
          setFilters(prev => ({ ...prev, lotStatus: 'Pending' }));
        } else {
          // For other supervisors, apply existing filters
          if (!newFilters.lotStatus || newFilters.lotStatus === '') {
            newFilters.lotStatus = '';
            setFilters(prev => ({ ...prev, lotStatus: '' }));
          }
          applyFilters(newFilters, supervisorData);
        }
        
        // CRITICAL FIX: Update filter options but PRESERVE existing supervisor options
        extractFilterOptionsPreservingSupervisors(supervisorData);
        
      } catch (fetchErr) {
        clearTimeout(timeoutId);
        throw fetchErr;
      }
      
    } catch (err) {
      console.error('Supervisor change error:', err);
      
      let errorMessage = `Failed to load data for ${value}: `;
      if (err.name === 'AbortError') {
        errorMessage += 'Request took too long.';
      } else if (err.message.includes('Failed to fetch')) {
        errorMessage += 'Network error.';
      } else {
        errorMessage += err.message;
      }
      
      setError(errorMessage);
      
      // Revert to previous supervisor on error
      setFilters(prev => ({ ...prev, supervisor: prev.supervisor }));
    } finally {
      setLoading(false);
    }
  } else {
    // For other filter changes, just apply filters to current data
    applyFilters(newFilters, data);
  }
}, [filters, data, fetchDataForSupervisor, applyFilters, getCacheKey, isLotCompleted]);

// Add this new function to preserve supervisor options
const extractFilterOptionsPreservingSupervisors = useCallback((data) => {
  const options = {
    fabric: [],
    garmentType: [],
    style: [],
    brand: [],
    partyName: [],
    supervisor: filterOptions.supervisor, // PRESERVE existing supervisor options
    season: [],
    mwk: [],
    directStitching: [],
    challanHistory: [],
    wipStatus: [],
    completedStatus: [],
    lotStatus: ['Pending', 'Completed'],
    priority: []
  };

  // Use Set for faster lookups
  const optionSets = {
    fabric: new Set(),
    garmentType: new Set(),
    style: new Set(),
    brand: new Set(),
    partyName: new Set(),
    season: new Set(),
    mwk: new Set(),
    directStitching: new Set(),
    challanHistory: new Set(),
    wipStatus: new Set(),
    completedStatus: new Set(),
    priority: new Set()
  };

  // Process data in batches
  const processRow = (item) => {
    // Process each field
    Object.keys(optionSets).forEach(key => {
      if (key === 'wipStatus') {
        const latestRemarks = getLatestWipRemarks(item[key]);
        if (latestRemarks && latestRemarks.trim() !== '' && latestRemarks !== 'Not Updated') {
          optionSets[key].add(normalizeAndCapitalize(latestRemarks));
        }
      } else if (key === 'completedStatus') {
        const statusText = getCompletedStatusText(item[key]);
        if (statusText && statusText.trim() !== '' && statusText !== 'N/A') {
          optionSets[key].add(normalizeAndCapitalize(statusText));
        }
      } else if (key === 'priority') {
        const priorityValue = item[key];
        if (priorityValue && typeof priorityValue === 'string' && priorityValue.trim() !== '') {
          optionSets[key].add(normalizeAndCapitalize(priorityValue.trim()));
        }
      } else if (item[key] && item[key].trim() !== '') {
        optionSets[key].add(normalizeAndCapitalize(item[key]));
      }
    });
  };

  // Process all rows
  data.forEach(processRow);

  // Convert Sets to sorted arrays
  Object.keys(optionSets).forEach(key => {
    options[key] = Array.from(optionSets[key]).sort();
  });

  // CRITICAL: Ensure supervisor options are preserved
  options.supervisor = filterOptions.supervisor; // Keep all supervisors
  
  setFilterOptions(options);
}, [filterOptions.supervisor, getLatestWipRemarks, getCompletedStatusText, normalizeAndCapitalize]);

// Also update the initial load useEffect
useEffect(() => {
  const loadInitialData = async () => {
    try {
      setLoading(true);
      setIsInitialLoad(true);
      
      const monuData = await fetchDataForSupervisor('Monu');
      setData(monuData);
      
      // Show only PENDING lots by default for Monu
      const pendingData = monuData.filter(item => !isLotCompleted(item.completedStatus));
      setFilteredData(pendingData);
      
      // Load ALL supervisors for dropdown
      await loadAllSupervisors();
      
      // Extract filter options from Monu's data
      extractFilterOptions(monuData);
      
    } catch (err) {
      setError(`Failed to load initial data: ${err.message}`);
    } finally {
      setLoading(false);
      setIsInitialLoad(false);
    }
  };

  loadInitialData();
}, [fetchDataForSupervisor, extractFilterOptions, isLotCompleted]);

// Add this function to load all supervisors separately
const loadAllSupervisors = useCallback(async () => {
  try {
    // Try to get supervisors from cache first
    const allCacheKey = getCacheKey('all');
    let allData;
    
    if (isCacheValid(allCacheKey) && dataCache.current.has(allCacheKey)) {
      allData = dataCache.current.get(allCacheKey);
    } else {
      // Fetch minimal data just to get supervisors
      allData = await fetchMinimalSupervisorData();
    }
    
    const allSupervisors = [...new Set(allData
      .filter(item => item.supervisor && item.supervisor.trim() !== '')
      .map(item => normalizeAndCapitalize(item.supervisor))
    )].sort();
    
    setFilterOptions(prev => ({
      ...prev,
      supervisor: allSupervisors
    }));
    
    return allSupervisors;
  } catch (err) {
    console.error('Error loading supervisor options:', err);
    return [];
  }
}, [getCacheKey, isCacheValid, normalizeAndCapitalize]);

// Add this function to fetch minimal supervisor data
const fetchMinimalSupervisorData = useCallback(async () => {
  try {
    // Fetch just the Supervisor column from the sheet
    const range = `${SHEET_NAME}!G:G`; // Assuming Supervisor is in column G
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(range)}?key=${API_KEY}`;
    
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error('Failed to fetch supervisor data');
    }
    
    const result = await response.json();
    const values = result.values || [];
    
    // Skip header row and create simple objects with supervisor field
    return values.slice(1).map(row => ({
      supervisor: row[0] || ''
    })).filter(item => item.supervisor.trim() !== '');
    
  } catch (err) {
    console.error('Error fetching minimal supervisor data:', err);
    return [];
  }
}, []);

// Update the clearFilters function to preserve all supervisors
const clearFilters = useCallback(async () => {
  const defaultFilters = {
    lotNumber: '',
    fabric: '',
    garmentType: '',
    style: '',
    brand: '',
    partyName: '',
    supervisor: 'Monu',
    season: '',
    mwk: '',
    directStitching: '',
    challanHistory: '',
    wipStatus: '',
    completedStatus: '',
    lotStatus: 'Pending',
    priority: ''
  };
  
  setFilters(defaultFilters);
  
  try {
    setLoading(true);
    
    // Clear Monu's cache
    const monuCacheKey = getCacheKey('Monu');
    dataCache.current.delete(monuCacheKey);
    lastFetchTime.current.delete(monuCacheKey);
    
    // Fetch fresh data for Monu
    const monuData = await fetchDataForSupervisor('Monu');
    setData(monuData);
    
    // Show only pending data for Monu
    const pendingData = monuData.filter(item => !isLotCompleted(item.completedStatus));
    setFilteredData(pendingData);
    
    // Update filter options but PRESERVE all supervisors
    extractFilterOptionsPreservingSupervisors(monuData);
    
  } catch (err) {
    setError(`Failed to load Monu's data: ${err.message}`);
  } finally {
    setLoading(false);
  }
}, [fetchDataForSupervisor, extractFilterOptionsPreservingSupervisors, isLotCompleted, getCacheKey]);

  // Load all data when needed (for export or supervisor dropdown)
  const loadAllData = useCallback(async () => {
    const cacheKey = getCacheKey('all');
    
    if (isCacheValid(cacheKey) && dataCache.current.has(cacheKey)) {
      return dataCache.current.get(cacheKey);
    }

    try {
      const allData = await fetchDataForSupervisor('');
      
      // Update supervisor filter options with all supervisors
      const allSupervisors = [...new Set(allData
        .filter(item => item.supervisor && item.supervisor.trim() !== '')
        .map(item => normalizeAndCapitalize(item.supervisor))
      )].sort();
      
      setFilterOptions(prev => ({
        ...prev,
        supervisor: allSupervisors
      }));

      return allData;
    } catch (err) {
      console.error('Error loading all data:', err);
      throw err;
    }
  }, [fetchDataForSupervisor, getCacheKey, isCacheValid, normalizeAndCapitalize]);

  // Enhanced download functions that load all data if needed
  const downloadExcel = useCallback(async () => {
    try {
      setLoading(true);
      
      // For Excel export, we might want all data or filtered data
      let exportData = filteredData;
      const exportFilters = { ...filters };
      
      // If no specific supervisor filter is applied, load all data for comprehensive export
      if (!filters.supervisor || filters.supervisor === '') {
        const allData = await loadAllData();
        exportData = allData;
        exportFilters.supervisor = 'All Supervisors';
      }

      if (exportData.length === 0) {
        alert('No data available to export.');
        return;
      }

      // Create worksheet data with ALL columns including Total PCS, Completion Date, and Status
      const worksheetData = [
        // Update HEADERS based on filter status
        filters.lotStatus === 'Pending' ? [
          'Sr.No',
          'Lot Number',
          'Fabric',
          'Garment Type',
          'Style',
          'BRAND',
          'PARTY NAME',
          'Season',
          'M/W/K',
          'Direct Stitching',
          'Stitching Days',
          'Emb/Print Date',
          'WIP Status',
          'Total PCS',
          'Color Status',      // NEW: Color Pending indicator
          'Pintu',
          'EA',
          'Status'
        ] : [
          'Sr.No',
          'Lot Number',
          'Fabric',
          'Garment Type',
          'Style',
          'BRAND',
          'PARTY NAME',
          'Season',
          'M/W/K',
          'Direct Stitching',
          'Stitching Days',
          'Emb/Print Date',
          'WIP Status',
          'Total PCS',
          'Completion Date',
          'Color Status',      // NEW: Color Pending indicator
          'Pintu',
          'EA',
          'Status'
        ],
        ...exportData.map((item, index) => {
          const hasColorPending = item.hasColorPending || false;
          const pendingColors = item.pendingColors || [];
          const colorStatus = hasColorPending ? `Color Pending (${pendingColors.length})` : 'OK';
          
          if (filters.lotStatus === 'Pending') {
            return [
              index + 1,
              item.lotNumber || '',
              item.fabric || '',
              item.garmentType || '',
              item.style || '',
              item.brand || '',
              item.partyName || '',
              item.season || '',
              item.mwk || '',
              item.directStitching || '',
              calculateStitchingDays(item.dateOfIssue),
              getEmbPrintDate(item.challanHistory),
              getLatestWipRemarks(item.wipStatus),
              item.totalPCS || 0,
              colorStatus, // Color Status column
              '', // Pintu - empty
              '', // EA - empty
              isLotCompleted(item.completedStatus) ? 'Completed' : 'Pending'
            ];
          } else {
            return [
              index + 1,
              item.lotNumber || '',
              item.fabric || '',
              item.garmentType || '',
              item.style || '',
              item.brand || '',
              item.partyName || '',
              item.season || '',
              item.mwk || '',
              item.directStitching || '',
              calculateStitchingDays(item.dateOfIssue),
              getEmbPrintDate(item.challanHistory),
              getLatestWipRemarks(item.wipStatus),
              item.totalPCS || 0,
              getCompletionDateFormatted(item.completedStatus) || '',
              colorStatus, // Color Status column
              '', // Pintu - empty
              '', // EA - empty
              isLotCompleted(item.completedStatus) ? 'Completed' : 'Pending'
            ];
          }
        })
      ];

      const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Stitching Issue Report');
      
      // Calculate column widths
      const maxWidths = HEADERS.map((header, colIndex) => {
        const maxLength = Math.max(
          header.length,
          ...exportData.map((item, rowIndex) => {
            const rowData = [
              rowIndex + 1,
              item.lotNumber || '',
              item.fabric || '',
              item.garmentType || '',
              item.style || '',
              item.brand || '',
              item.partyName || '',
              item.season || '',
              item.mwk || '',
              item.directStitching || '',
              calculateStitchingDays(item.dateOfIssue),
              getEmbPrintDate(item.challanHistory),
              getLatestWipRemarks(item.wipStatus),
              item.totalPCS || 0,
              getCompletionDateFormatted(item.completedStatus) || '',
              isLotCompleted(item.completedStatus) ? 'Completed' : 'Pending'
            ];
            return rowData[colIndex] ? rowData[colIndex].toString().length : 0;
          })
        );
        return { wch: Math.min(maxLength + 2, 50) };
      });
      
      worksheet['!cols'] = maxWidths;

      // Add formatting for important columns
      const range = XLSX.utils.decode_range(worksheet['!ref']);
      for (let row = 1; row <= exportData.length; row++) {
        // Format Total PCS column (column 13, 0-based)
        const pcsCellAddress = XLSX.utils.encode_cell({ r: row, c: 13 });
        if (worksheet[pcsCellAddress]) {
          worksheet[pcsCellAddress].s = {
            font: { bold: true },
            numFmt: '#,##0'
          };
        }
        
        // Format Status column (column 15, 0-based)
        const statusCellAddress = XLSX.utils.encode_cell({ r: row, c: 15 });
        if (worksheet[statusCellAddress]) {
          const isCompleted = worksheet[statusCellAddress].v === 'Completed';
          worksheet[statusCellAddress].s = {
            font: { bold: true },
            fill: {
              fgColor: { rgb: isCompleted ? 'C6EFCE' : 'FFC7CE' } // Green for completed, Red for pending
            }
          };
        }
      }

      const fileName = exportFilters.supervisor === 'All Supervisors' 
        ? `All_Supervisors_Stitching_Report_${new Date().toISOString().split('T')[0]}`
        : `${exportFilters.supervisor || 'Stitching'}_Report_${new Date().toISOString().split('T')[0]}`;
      
      XLSX.writeFile(workbook, `${fileName}.xlsx`);
    } catch (err) {
      setError(`Failed to export Excel: ${err.message}`);
    } finally {
      setLoading(false);
    }
  }, [filteredData, filters, loadAllData, calculateStitchingDays, getEmbPrintDate, getLatestWipRemarks, getCompletionDateFormatted, isLotCompleted]);

  // Add this helper function to parse Date of Issue properly
  const parseDateOfIssue = useCallback((dateString) => {
    if (!dateString || dateString.trim() === '') return null;
    
    try {
      // Try standard date parsing
      const date = new Date(dateString);
      if (!isNaN(date.getTime())) return date;
      
      // Try DD/MM/YYYY format (common in Indian dates)
      const parts = dateString.split('/');
      if (parts.length === 3) {
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10) - 1;
        const year = parseInt(parts[2], 10);
        const fullYear = year < 100 ? 2000 + year : year;
        const newDate = new Date(fullYear, month, day);
        if (!isNaN(newDate.getTime())) return newDate;
      }
      
      // Try MM/DD/YYYY format
      const parts2 = dateString.split('/');
      if (parts2.length === 3) {
        const month = parseInt(parts2[0], 10) - 1;
        const day = parseInt(parts2[1], 10);
        const year = parseInt(parts2[2], 10);
        const fullYear = year < 100 ? 2000 + year : year;
        const newDate = new Date(fullYear, month, day);
        if (!isNaN(newDate.getTime())) return newDate;
      }
      
      return null;
    } catch (error) {
      console.error('Error parsing date of issue:', error, dateString);
      return null;
    }
  }, []);

  // Function to get the latest manpower for each supervisor based on Date of Issue
  const getLatestManpowerBySupervisor = useCallback((data) => {
    const supervisorLatestManpower = {};
    const supervisorLatestDate = {};
    
    data.forEach(item => {
      const supervisor = normalizeText(item.supervisor);
      if (!supervisor || supervisor.trim() === '') return;
      
      const manpower = item.manpower ? String(item.manpower).trim() : '';
      if (!manpower) return; // Skip if no manpower data
      
      const dateOfIssue = parseDateOfIssue(item.dateOfIssue);
      if (!dateOfIssue) return; // Skip if no valid date
      
      const currentDate = supervisorLatestDate[supervisor];
      const currentManpower = supervisorLatestManpower[supervisor];
      
      // If this is the first entry for supervisor OR this date is more recent
      if (!currentDate || dateOfIssue.getTime() > currentDate.getTime()) {
        supervisorLatestDate[supervisor] = dateOfIssue;
        supervisorLatestManpower[supervisor] = manpower;
      }
      // If dates are equal, keep the one with valid manpower (not empty)
      else if (dateOfIssue.getTime() === currentDate.getTime() && 
               manpower && (!currentManpower || currentManpower.trim() === '')) {
        supervisorLatestManpower[supervisor] = manpower;
      }
    });
    
    return supervisorLatestManpower;
  }, [parseDateOfIssue, normalizeText]);

  // PDF download function - ENHANCED to match the first component's design
const downloadPDF = useCallback(async () => {
  try {
    setLoading(true);
    
    // Use filteredData directly
    const exportData = filteredData;
    
    if (exportData.length === 0) {
      alert('No data available to export.');
      return;
    }

    // Function to check if WIP status is updated today
    const isStatusUpdatedToday = (wipStatus) => {
      if (!wipStatus || wipStatus.trim() === '') return false;
      
      try {
        // Parse the JSON array
        const statusArray = JSON.parse(wipStatus);
        
        if (!Array.isArray(statusArray) || statusArray.length === 0) return false;
        
        // Get the most recent status
        const sortedStatuses = [...statusArray].sort((a, b) => {
          const dateA = new Date(a.timestamp).getTime();
          const dateB = new Date(b.timestamp).getTime();
          return dateB - dateA;
        });
        
        const latestEntry = sortedStatuses[0];
        if (!latestEntry || !latestEntry.timestamp) return false;
        
        // Check if latest entry is from today
        const today = new Date();
        const entryDate = new Date(latestEntry.timestamp);
        
        return entryDate.getDate() === today.getDate() &&
               entryDate.getMonth() === today.getMonth() &&
               entryDate.getFullYear() === today.getFullYear();
      } catch (error) {
        console.error('Error checking status update date:', error);
        return false;
      }
    };

    // Simple circle attention symbol
    const drawCircleAttentionSymbol = (doc, x, y, size = 2.5) => {
      const currentFillColor = doc.getFillColor();
      const currentDrawColor = doc.getDrawColor();
      const currentLineWidth = doc.internal.getLineWidth();
      
      // Draw orange circle with border
      doc.setFillColor(255, 165, 0);
      doc.setDrawColor(220, 100, 0);
      doc.setLineWidth(0.3);
      doc.circle(x, y, size, 'FD');
      
      // Draw exclamation mark inside
      doc.setFillColor(255, 255, 255);
      doc.setDrawColor(200, 80, 0);
      
      // Draw exclamation mark body (vertical bar)
      const barWidth = size * 0.25;
      const barHeight = size * 1.0;
      
      // Top part of exclamation
      doc.rect(
        x - barWidth/2, 
        y - barHeight/2 + size * 0.2, 
        barWidth, 
        barHeight * 0.7, 
        'F'
      );
      
      // Bottom dot of exclamation
      doc.circle(x, y + size * 0.4, barWidth * 0.8, 'F');
      
      // Add border to exclamation mark
      doc.setLineWidth(0.1);
      doc.rect(
        x - barWidth/2, 
        y - barHeight/2 + size * 0.2, 
        barWidth, 
        barHeight * 0.7, 
        'D'
      );
      doc.circle(x, y + size * 0.4, barWidth * 0.8, 'D');
      
      // Restore original state
      doc.setFillColor(currentFillColor);
      doc.setDrawColor(currentDrawColor);
      doc.setLineWidth(currentLineWidth);
    };

    const drawStarSymbol = (doc, x, y, size = 2.5) => {
      // Save current state
      const currentFillColor = doc.getFillColor();
      const currentDrawColor = doc.getDrawColor();
      
      // Draw a gold circle as the base
      doc.setFillColor(255, 215, 0); // Gold
      doc.setDrawColor(184, 134, 11); // Dark gold border
      doc.setLineWidth(0.3);
      doc.circle(x, y, size, 'FD');
      
      // Draw star points using simple lines
      doc.setDrawColor(160, 120, 10);
      doc.setLineWidth(0.4);
      
      // Draw 5 lines from center to create star shape
      const points = 5;
      for (let i = 0; i < points; i++) {
        const angle = (i * 2 * Math.PI / points) - Math.PI / 2;
        const endX = x + (size * 1.3) * Math.cos(angle);
        const endY = y + (size * 1.3) * Math.sin(angle);
        
        // Draw line from center to outer point
        doc.line(x, y, endX, endY);
        
        // Draw line from outer point to opposite inner point
        const nextAngle = angle + (Math.PI / points);
        const innerX = x + (size * 0.6) * Math.cos(nextAngle);
        const innerY = y + (size * 0.6) * Math.sin(nextAngle);
        doc.line(endX, endY, innerX, innerY);
      }
      
      // Draw a bright yellow center dot
      doc.setFillColor(255, 255, 180);
      doc.circle(x, y, size * 0.4, 'F');
      
      // Add an outer circle to make it look more like a star
      doc.setDrawColor(140, 100, 0);
      doc.setLineWidth(0.2);
      doc.circle(x, y, size * 0.95, 'D');
      
      // Restore original state
      doc.setFillColor(currentFillColor);
      doc.setDrawColor(currentDrawColor);
    };

    // Function to normalize supervisor name
    const normalizeSupervisorName = (name) => {
      if (!name || name.trim() === '') return 'Unassigned';
      
      const trimmedName = name.trim();
      return trimmedName.charAt(0).toUpperCase() + trimmedName.slice(1).toLowerCase();
    };

    // Function to abbreviate party names
    const abbreviatePartyName = (partyName) => {
      if (!partyName || partyName.trim() === '') return 'N/A';
      
      const name = partyName.trim();
      
      // Specific abbreviations
      if (name.toLowerCase().includes('mohit hosiery')) return 'MH';
      if (name.toLowerCase().includes('hosiery')) return name.split(' ')[0];
      
      // General abbreviation: take first letters of first two words
      const words = name.split(' ');
      if (words.length === 1) {
        return words[0].substring(0, 3).toUpperCase();
      } else if (words.length >= 2) {
        return (words[0].charAt(0) + words[1].charAt(0)).toUpperCase();
      }
      
      return name.substring(0, 4).toUpperCase();
    };

    // Function to abbreviate M/W/K for PDF
    const abbreviateMWKForPDF = (mwkValue) => {
      if (!mwkValue || typeof mwkValue !== 'string') return 'N/A';
      
      const value = mwkValue.trim().toLowerCase();
      
      if (value.includes('gents') || value === 'm' || value === 'mens') return 'M';
      if (value.includes('kids') || value === 'k') return 'K';
      if (value.includes('girls') || value === 'g' || value.includes('girlish') || value.includes('girls')) return 'G';
      if (value.includes('women') || value.includes('womens') || value === 'w' || value.includes('women')) return 'W';
      
      // Return first letter if no match
      return value.charAt(0).toUpperCase();
    };

    // Function to format date as "29/11/25" for PDF
    const formatDateToDDMMYYForPDF = (dateString) => {
      if (!dateString || dateString.trim() === '' || dateString === '-' || dateString === 'N/A') {
        return 'N/A';
      }
      
      try {
        const date = new Date(dateString);
        
        if (isNaN(date.getTime())) {
          // Try different date formats
          const parts = dateString.split(/[\/\-\.]/);
          if (parts.length === 3) {
            // Try day/month/year
            const day = parseInt(parts[0], 10);
            const month = parseInt(parts[1], 10);
            const year = parseInt(parts[2], 10);
            
            // Handle 2-digit and 4-digit years
            const fullYear = year < 100 ? 2000 + year : year;
            
            // Check if this is a valid date
            const testDate = new Date(fullYear, month - 1, day);
            if (!isNaN(testDate.getTime())) {
              return `${String(day).padStart(2, '0')}/${String(month).padStart(2, '0')}/${String(fullYear).slice(-2)}`;
            }
          }
          return dateString;
        }
        
        const day = String(date.getDate()).padStart(2, '0');
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const year = String(date.getFullYear()).slice(-2);
        return `${day}/${month}/${year}`;
      } catch {
        return dateString;
      }
    };

    // Function for season abbreviation
    const abbreviateSeason = (season) => {
      if (!season || season.trim() === '') return 'N/A';
      
      const seasonLower = season.trim().toLowerCase();
      
      if (seasonLower.includes('summer')) return 'S';
      if (seasonLower.includes('winter')) return 'W';
      if (seasonLower.includes('autumn')) return 'A';
      if (seasonLower.includes('spring')) return 'SP';
      
      // Return first letter if not a known season
      return season.charAt(0).toUpperCase();
    };

    // Function to format Date of Issue
    const formatDateOfIssue = (dateString) => {
      if (!dateString || dateString.trim() === '') return 'N/A';
      
      try {
        const date = new Date(dateString);
        
        if (isNaN(date.getTime())) {
          // Try different date formats
          const parts = dateString.split(/[\/\-\.]/);
          if (parts.length === 3) {
            const day = parseInt(parts[0], 10);
            const month = parseInt(parts[1], 10);
            const year = parseInt(parts[2], 10);
            
            const fullYear = year < 100 ? 2000 + year : year;
            
            const testDate = new Date(fullYear, month - 1, day);
            if (!isNaN(testDate.getTime())) {
              return `${String(day).padStart(2, '0')}/${String(month).padStart(2, '0')}/${String(fullYear).slice(-2)}`;
            }
          }
          return dateString;
        }
        
        const day = String(date.getDate()).padStart(2, '0');
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const year = String(date.getFullYear()).slice(-2);
        return `${day}/${month}/${year}`;
      } catch {
        return dateString;
      }
    };

    // Function to extract stage from WIP status
    const extractStageFromStatus = (wipRemarks) => {
      if (!wipRemarks || wipRemarks === 'N/A' || wipRemarks.trim() === '') return 'Other';
      
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
      } else if (remarks.includes('Tailor Working')) {
        return 'Tailor Working';
      } else if (remarks.includes('Emb Pending')) {
        return 'Emb Pending';
      } else {
        return 'On Stitching';
      }
    };

    // Create PDF in landscape mode
    const doc = new jsPDF({
      orientation: 'landscape',
      unit: 'mm',
      format: 'a3'
    });

    // Colors
    const headerColor = [15, 76, 129]; // Navy Blue
    const borderColor = [0, 0, 0]; // Black border
    const textColor = [17, 24, 39]; // Dark text
    const accentColor = [59, 130, 246]; // Blue accent
    const remarksColor = [239, 68, 68]; // Red for remarks
    const pcsColor = [239, 68, 68]; // RED for PCS
    const partyColor = [107, 33, 168]; // Purple for Party abbreviation
    const completedColor = [16, 185, 129]; // Green for completed status
    const pendingColor = [245, 158, 11]; // Yellow for pending status
    const issueDateColor = [139, 92, 246]; // Purple for Date of Issue
    const colorPendingColor = [255, 165, 0]; // Orange for color pending
    const repeatedLotColor = [255, 215, 0]; // Gold for repeated lot
    const stageAnalysisColor = [59, 130, 246]; // Blue for stage analysis
    const notUpdatedColor = [255, 235, 235]; // Light red for not updated status
    const notUpdatedTextColor = [220, 38, 38]; // Dark red text for not updated

    // Page dimensions
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const margin = 8;
    const contentWidth = pageWidth - (margin * 2);

    // Variable to track current Y position
    let currentY = 30;

    // Function to draw header
    const drawHeader = () => {
      // WHITE background for entire header area
      doc.setFillColor(255, 255, 255);
      doc.rect(0, 0, pageWidth, 30, 'F');
      
      // Calculate total PCS and other metrics
      const totalPCS = exportData.reduce((sum, item) => sum + (item.totalPCS || 0), 0);
      const totalLots = exportData.length;
      const totalCompleted = exportData.filter(item => isLotCompleted(item.completedStatus)).length;
      const totalPending = totalLots - totalCompleted;
      const totalColorPending = exportData.filter(item => item.hasColorPending).length;
      const totalNotUpdated = exportData.filter(item => !isStatusUpdatedToday(item.wipStatus)).length;
      
      // Calculate repeated lots
      const totalRepeatedLots = exportData.filter(item => {
        const hasRepeat = item.priority && normalizeText(item.priority).includes('repeated_lot');
        return hasRepeat;
      }).length;
      
      // Dynamic title based on supervisor
      let title = 'DAILY STITCHING ISSUE REPORT';
      
      // Add supervisor to title if filtered
      let supervisorName = '';
      if (filters.supervisor && filters.supervisor.trim() !== '') {
        supervisorName = filters.supervisor.toUpperCase();
        title = `${supervisorName} - DAILY STITCHING ISSUE REPORT`;
      }
      
      // Main Title
      doc.setFontSize(18);
      doc.setTextColor(15, 76, 129);
      doc.setFont('Times New Roman', 'bold');
      doc.text(title, pageWidth / 2, 12, { align: 'center' });
      
      // Add supervisor badge if supervisor filter is applied
      if (supervisorName) {
        doc.setFontSize(10);
        doc.setFont('times', 'bold');
        doc.setTextColor(15, 76, 129);
        doc.text(`SUPERVISOR: ${supervisorName}`, pageWidth / 2, 18, { align: 'center' });
      }
      
      // Key Metrics Row
      doc.setFontSize(10);
      doc.setFont('times', 'bold');
      doc.setTextColor(15, 76, 129);
      
      // Left side: Date
      const today = new Date();
      const reportDate = `${String(today.getDate()).padStart(2, '0')}/${String(today.getMonth() + 1).padStart(2, '0')}/${String(today.getFullYear()).slice(-2)}`;
      doc.text(`Report Date: ${reportDate}`, margin, 25);
      
      // Center: Totals summary
      const centerX = pageWidth / 2;
      doc.text(`Lots: ${totalLots} | PCS: ${totalPCS.toLocaleString()} | Completed: ${totalCompleted} | Pending: ${totalPending} | Color Pending: ${totalColorPending} | Not Updated: ${totalNotUpdated} | ★: ${totalRepeatedLots}`, 
                centerX, 25, { align: 'center' });
      
      // Right side: Filter status
      const activeFilterCount = Object.entries(filters)
        .filter(([key, value]) => {
          if (key === 'dateRange') {
            const { from, to } = value;
            return from || to;
          }
          return key !== 'supervisor' && typeof value === 'string' && value.trim() !== '';
        })
        .length;
      
      const filterText = activeFilterCount > 0 ? `Filtered: ${exportData.length} records` : `Total: ${exportData.length} records`;
      doc.text(filterText, pageWidth - margin, 25, { align: 'right' });
    };

    // Draw header on first page
    drawHeader();

    // Function to add color coding legend with not updated status
    const addColorLegend = (yPos) => {
      doc.setFontSize(8);
      doc.setFont('times', 'bold');
      
      // Legend title
      doc.setTextColor(59, 130, 246);
      doc.text('LEGEND:', margin, yPos);
      
      // Save current state
      const originalFont = doc.internal.getFont();
      const originalSize = doc.internal.getFontSize();
      const originalTextColor = doc.getTextColor();
      
      // Legend items with "Status Not Updated" item
      const legendItems = [
        { 
          type: 'color-box',
          color: [16, 185, 129], 
          text: '1-6 Days'
        },
        { 
          type: 'color-box',
          color: [245, 158, 11], 
          text: '7-15 Days'
        },
        { 
          type: 'color-box',
          color: [239, 68, 68], 
          text: '15+ Days'
        },
        { 
          type: 'attention',
          text: 'Color Pending'
        },
        { 
          type: 'star',
          text: 'Repeated Lot'
        },
        { 
          type: 'status-not-updated',
          text: 'Status Not Updated Today'
        }
      ];
      
      let legendX = margin + 35;
      const iconSize = 2;
      
      legendItems.forEach((item, index) => {
        if (item.type === 'color-box') {
          // Draw simple color box
          doc.setFillColor(...item.color);
          doc.rect(legendX, yPos - 3, 4, 4, 'F');
          doc.setTextColor(50, 50, 50);
          doc.setFont('times', 'normal');
          doc.text(item.text, legendX + 7, yPos);
          legendX += 50;
        } 
        else if (item.type === 'attention') {
          // Draw attention symbol
          const iconX = legendX + 2;
          const iconY = yPos - 1.5;
          drawCircleAttentionSymbol(doc, iconX, iconY, iconSize);
          doc.setTextColor(50, 50, 50);
          doc.setFont('times', 'normal');
          doc.text(item.text, iconX + 5, yPos);
          legendX += 60;
        }
        else if (item.type === 'star') {
          // Draw star symbol
          const iconX = legendX + 2;
          const iconY = yPos - 1.5;
          drawStarSymbol(doc, iconX, iconY, iconSize);
          doc.setTextColor(50, 50, 50);
          doc.setFont('times', 'normal');
          doc.text(item.text, iconX + 5, yPos);
          legendX += 60;
        }
        else if (item.type === 'status-not-updated') {
          // Draw a small red square for not updated status
          const iconX = legendX + 2;
          const iconY = yPos - 3;
          doc.setFillColor(255, 235, 235);
          doc.setDrawColor(220, 38, 38);
          doc.rect(iconX - 2, iconY, 4, 4, 'FD');
          doc.setTextColor(50, 50, 50);
          doc.setFont('times', 'normal');
          doc.text(item.text, iconX + 5, yPos);
          legendX += 70;
        }
      });
      
      // Restore original state
      doc.setFont(originalFont.fontName, originalFont.fontStyle);
      doc.setFontSize(originalSize);
      doc.setTextColor(originalTextColor[0], originalTextColor[1], originalTextColor[2]);
      
      return yPos + 10;
    };

    // Add color legend after header
    currentY = addColorLegend(35);

    // ===================== MAIN DATA TABLE =====================
    
    // Updated headers
    const PDF_HEADERS = filters.lotStatus === 'Pending' ? [
      'Sr',
      'Lot No',
      'Garment',
      'Style',
      'Fabric',
      'Brand',
      'Total PCS',
      'M/W/K',
      'Season',
      'Party',
      'Direct',
      'Supervisor',
      'Issue Date',
      'Days',
      'Emb/Print',
      'WIP Status',
      'Color Status',
      'Pintu',
      'EA',
      'Lot Status'
    ] : [
      'Sr',
      'Lot No',
      'Garment',
      'Style',
      'Fabric',
      'Brand',
      'Total PCS',
      'M/W/K',
      'Season',
      'Party',
      'Direct',
      'Supervisor',
      'Issue Date',
      'Days',
      'Emb/Print',
      'WIP Status',
      'Complete Date',
      'Color Status',
      'Pintu',
      'EA',
      'Lot Status'
    ];

    // Optimized column widths
    const columnWidths = filters.lotStatus === 'Pending' ? {
      0: 9,    // Sr
      1: 20,   // Lot No
      2: 28,   // Garment
      3: 31,   // Style
      4: 28,   // Fabric
      5: 25,   // Brand
      6: 18,   // Total PCS
      7: 12,   // M/W/K
      8: 9,    // Season
      9: 14,   // Party
      10: 12,  // Direct
      11: 21,  // Supervisor
      12: 20,  // Issue Date
      13: 15,  // Days
      14: 20,  // Emb/Print
      15: 40,  // WIP Status
      16: 25,  // Color Status
      17: 20,  // Pintu
      18: 22,  // EA
      19: 15   // Lot Status
    } : {
      0: 9,    // Sr
      1: 20,   // Lot No
      2: 29,   // Garment
      3: 31,   // Style
      4: 28,   // Fabric
      5: 31,   // Brand
      6: 25,   // Total PCS
      7: 12,   // M/W/K
      8: 9,    // Season
      9: 16,   // Party
      10: 12,  // Direct
      11: 29,  // Supervisor
      12: 20,  // Issue Date
      13: 15,  // Days
      14: 20,  // Emb/Print
      15: 48,  // WIP Status
      16: 25,  // Complete Date
      17: 25,  // Color Status
      18: 15,  // Pintu
      19: 15,  // EA
      20: 20   // Lot Status
    };
    
    // Prepare table headers
    const headers = [
      PDF_HEADERS.map((header, index) => ({
        content: header,
        styles: { 
          fontStyle: 'bold', 
          fillColor: headerColor, 
          textColor: [255, 255, 255], 
          cellWidth: columnWidths[index], 
          halign: 'center',
          fontSize: 9,
          cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
        }
      }))
    ];

    // Prepare table body with color coding and priority - UPDATED with status not updated check
    const tableRowsData = exportData.map((item, rowIndex) => {
      const stitchingDays = calculateStitchingDays(item.dateOfIssue);
      const stitchingDaysColor = getStitchingDaysColor(stitchingDays);
      const stitchingDaysTextColor = getStitchingDaysTextColor(stitchingDays);
      
      // Check if status is updated today
      const isStatusUpdated = isStatusUpdatedToday(item.wipStatus);
      const wipRemarks = getLatestWipRemarks(item.wipStatus);
      
      const embPrintDate = getEmbPrintDate(item.challanHistory);
      const abbreviatedParty = abbreviatePartyName(item.partyName);
      const abbreviatedSeason = abbreviateSeason(item.season);
      const totalPCS = item.totalPCS || 0;
      const completionDate = getCompletionDateFormatted(item.completedStatus);
      const isCompleted = isLotCompleted(item.completedStatus);
      const lotStatus = isCompleted ? 'Completed' : 'Pending';
      const issueDate = formatDateOfIssue(item.dateOfIssue);
      
      // Check for color pending
      const hasColorPending = item.hasColorPending || false;
      const pendingColors = item.pendingColors || [];
      
      // Check for repeated lot priority
      const isRepeatedLot = item.priority && (
        normalizeText(item.priority).includes('repeated_lot') || 
        normalizeText(item.priority).includes('repeatedlot') ||
        normalizeText(item.priority).includes('repeat_lot') ||
        normalizeText(item.priority).includes('repetitive_lot') ||
        (normalizeText(item.priority).includes('repeat') && normalizeText(item.priority).includes('lot'))
      );
      
      // Format color status text
      let colorStatusText = 'No Colour Pending';
      if (hasColorPending && pendingColors.length > 0) {
        const colorList = pendingColors.join(', ');
        colorStatusText = `Colour Pending (${colorList})`;
      }
      
      // Determine row background
      let rowBgColor = rowIndex % 2 === 0 ? [255, 255, 255] : [250, 250, 250];
      if (hasColorPending) {
        rowBgColor = [255, 245, 235]; // Light orange for color pending
      }
      if (isRepeatedLot) {
        rowBgColor = [255, 253, 231]; // Light yellow for repeated lots
      }
      
      // Convert RGB color to array
      const stitchingDaysRGB = stitchingDaysColor.split('#')[1] ? [
        parseInt(stitchingDaysColor.slice(1, 3), 16),
        parseInt(stitchingDaysColor.slice(3, 5), 16),
        parseInt(stitchingDaysColor.slice(5, 7), 16)
      ] : [16, 185, 129];
      
      // Stitching days text color
      const stitchingDaysTextRGB = stitchingDaysTextColor === '#ffffff' ? [255, 255, 255] : [0, 0, 0];
      
      // Lot status color
      const lotStatusColor = isCompleted ? completedColor : pendingColor;
      const lotStatusTextColor = isCompleted ? [255, 255, 255] : [0, 0, 0];
      
      // Color pending status color
      const colorStatusRGB = hasColorPending ? [255, 165, 0] : [16, 185, 129];
      const colorStatusTextRGB = hasColorPending ? [255, 255, 255] : [0, 0, 0];
      
      // Simple truncate function
      const truncateText = (text, maxLength) => {
        if (!text || text === 'N/A' || text === '-') return text;
        const cleanText = text.toString().trim();
        return cleanText.length > maxLength ? cleanText.substring(0, maxLength - 3) + '...' : cleanText;
      };
      
      // Build row cells
      const rowCells = [];
      
      // 0: Sr.No - Add star indicator for repeated lots
      const srNoContent = isRepeatedLot ? `★ ${rowIndex + 1}` : (rowIndex + 1).toString();
      rowCells.push({
        content: srNoContent,
        styles: {
          cellWidth: columnWidths[0],
          fontSize: isRepeatedLot ? 9 : 10,
          halign: 'center',
          fontStyle: isRepeatedLot ? 'bold' : 'bold',
          fillColor: rowBgColor,
          textColor: isRepeatedLot ? [218, 165, 32] : textColor,
          cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
        }
      });
      
      // 1: Lot No
      rowCells.push({
        content: item.lotNumber || 'N/A',
        styles: {
          cellWidth: columnWidths[1],
          fontSize: isRepeatedLot ? 10 : 10,
          halign: 'center',
          fontStyle: 'bold',
          textColor: hasColorPending ? [255, 0, 0] : (isRepeatedLot ? [218, 165, 32] : [255, 0, 0]),
          fillColor: rowBgColor,
          cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
        }
      });

      // 2: Garment
      rowCells.push({
        content: item.garmentType || 'N/A',
        styles: {
          cellWidth: columnWidths[2],
          fontSize: 10,
          halign: 'center',
          fillColor: rowBgColor,
          cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
        }
      });

      // 3: Style
      rowCells.push({
        content: item.style || 'N/A',
        styles: {
          cellWidth: columnWidths[3],
          fontSize: 10,
          halign: 'center',
          fillColor: rowBgColor,
          cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
        }
      });

      // 4: Fabric
      rowCells.push({
        content: item.fabric || 'N/A',
        styles: {
          cellWidth: columnWidths[4],
          fontSize: 10,
          halign: 'center',
          fillColor: rowBgColor,
          cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
        }
      });

      // 5: Brand
      rowCells.push({
        content: item.brand || 'N/A',
        styles: {
          cellWidth: columnWidths[5],
          fontSize: 10,
          halign: 'center',
          fillColor: rowBgColor,
          cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
        }
      });

      // 6: Total PCS
      rowCells.push({
        content: totalPCS > 0 ? totalPCS.toLocaleString() : 'N/A',
        styles: {
          cellWidth: columnWidths[6],
          fontSize: 9,
          halign: 'center',
          fontStyle: 'bold',
          fillColor: rowBgColor,
          textColor: pcsColor,
          cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
        }
      });
      
      // 7: M/W/K - abbreviated
      const abbreviatedMWK = abbreviateMWKForPDF(item.mwk);
      rowCells.push({
        content: abbreviatedMWK,
        styles: {
          cellWidth: columnWidths[7],
          fontSize: 10,
          halign: 'center',
          fontStyle: 'bold',
          fillColor: rowBgColor,
          textColor: abbreviatedMWK === 'M' ? [59, 130, 246] :
                     abbreviatedMWK === 'W' ? [239, 68, 68] :
                     abbreviatedMWK === 'K' ? [16, 185, 129] :
                     abbreviatedMWK === 'G' ? [168, 85, 247] :
                     [100, 100, 100],
          cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
        }
      });

      // 8: Season
      rowCells.push({
        content: abbreviatedSeason,
        styles: {
          cellWidth: columnWidths[8],
          fontSize: 10,
          halign: 'center',
          fontStyle: 'bold',
          fillColor: rowBgColor,
          textColor: abbreviatedSeason === 'S' ? [239, 68, 68] :
                     abbreviatedSeason === 'W' ? [59, 130, 246] :
                     abbreviatedSeason === 'A' ? [245, 158, 11] :
                     abbreviatedSeason === 'SP' ? [16, 185, 129] : [100, 100, 100],
          cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
        }
      });

      // 9: Party
      rowCells.push({
        content: abbreviatedParty,
        styles: {
          cellWidth: columnWidths[9],
          fontSize: 10,
          halign: 'center',
          fontStyle: 'bold',
          fillColor: rowBgColor,
          textColor: partyColor,
          cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
        }
      });
      
      // 10: Direct
      rowCells.push({
        content: truncateText(item.directStitching || 'N/A', 8),
        styles: {
          cellWidth: columnWidths[10],
          fontSize: 10,
          halign: 'center',
          fillColor: rowBgColor,
          cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
        }
      });

      // 11: Supervisor
      rowCells.push({
        content: truncateText(item.supervisor || 'N/A', 15),
        styles: {
          cellWidth: columnWidths[11],
          fontSize: 10,
          halign: 'center',
          fontStyle: 'bold',
          fillColor: rowBgColor,
          textColor: [59, 130, 246],
          cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
        }
      });
      
      // 12: Issue Date
      rowCells.push({
        content: truncateText(issueDate, 12),
        styles: {
          cellWidth: columnWidths[12],
          fontSize: 10,
          halign: 'center',
          fillColor: rowBgColor,
          fontStyle: 'bold',
          textColor: issueDateColor,
          cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
        }
      });
      
      // 13: Days
      rowCells.push({
        content: stitchingDays.toString(),
        styles: {
          cellWidth: columnWidths[13],
          fontSize: 10,
          halign: 'center',
          fontStyle: 'bold',
          fillColor: stitchingDaysRGB,
          textColor: stitchingDaysTextRGB,
          cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
        }
      });
      
      // 14: Emb/Print - formatted date
      const formattedEmbPrintDate = embPrintDate !== '-' ? formatDateToDDMMYYForPDF(embPrintDate) : '-';
      rowCells.push({
        content: truncateText(formattedEmbPrintDate, 14),
        styles: {
          cellWidth: columnWidths[14],
          fontSize: 10,
          halign: 'center',
          fillColor: rowBgColor,
          fontStyle: formattedEmbPrintDate !== '-' ? 'bold' : 'normal',
          textColor: formattedEmbPrintDate !== '-' ? accentColor : [100, 100, 100],
          cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
        }
      });
      
      // 15: WIP Status - UPDATED with not updated highlighting
      rowCells.push({
        content: wipRemarks || 'N/A',
        styles: {
          cellWidth: columnWidths[15],
          fontSize: 10,
          halign: 'center',
          fillColor: !isStatusUpdated ? [255, 235, 235] : rowBgColor, // Light red background if not updated
          fontStyle: !isStatusUpdated ? 'bolditalic' : (wipRemarks !== 'N/A' ? 'bold' : 'normal'),
          textColor: !isStatusUpdated ? [220, 38, 38] : (wipRemarks !== 'N/A' ? remarksColor : [100, 100, 100]),
          cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
        }
      });

      // Conditional: Complete Date column
      if (filters.lotStatus !== 'Pending') {
        const formattedCompletionDate = completionDate ? formatDateToDDMMYYForPDF(completionDate) : '-';
        rowCells.push({
          content: truncateText(formattedCompletionDate, 14),
          styles: {
            cellWidth: columnWidths[16],
            fontSize: 8,
            halign: 'center',
            fillColor: rowBgColor,
            fontStyle: formattedCompletionDate !== '-' ? 'bold' : 'normal',
            textColor: formattedCompletionDate !== '-' ? [59, 130, 246] : [100, 100, 100],
            cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
          }
        });
      }
      
      // Color Status column
      const colorStatusColIndex = filters.lotStatus === 'Pending' ? 16 : 17;
      rowCells.push({
        content: colorStatusText,
        styles: {
          cellWidth: columnWidths[colorStatusColIndex],
          fontSize: hasColorPending ? 7 : 8,
          halign: 'center',
          fontStyle: 'bold',
          fillColor: colorStatusRGB,
          textColor: colorStatusTextRGB,
          cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
        }
      });
      
      // Pintu column
      const pintuColIndex = filters.lotStatus === 'Pending' ? 17 : 18;
      rowCells.push({
        content: '',
        styles: {
          cellWidth: columnWidths[pintuColIndex],
          fontSize: 9,
          halign: 'center',
          fillColor: rowBgColor,
          fontStyle: 'normal',
          textColor: [59, 130, 246],
          cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
        }
      });
      
      // EA column
      const eaColIndex = filters.lotStatus === 'Pending' ? 18 : 19;
      rowCells.push({
        content: '',
        styles: {
          cellWidth: columnWidths[eaColIndex],
          fontSize: 9,
          halign: 'center',
          fillColor: rowBgColor,
          fontStyle: 'normal',
          textColor: [16, 185, 129],
          cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
        }
      });
      
      // Lot Status
      const lotStatusColIndex = filters.lotStatus === 'Pending' ? 19 : 20;
      rowCells.push({
        content: truncateText(lotStatus, 10),
        styles: {
          cellWidth: columnWidths[lotStatusColIndex],
          fontSize: 9,
          halign: 'center',
          fontStyle: 'bold',
          fillColor: lotStatusColor,
          textColor: lotStatusTextColor,
          cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
        }
      });
      
      return {
        cells: rowCells,
        hasColorPending,
        pendingColors,
        isRepeatedLot,
        lotNumber: item.lotNumber || 'N/A',
        priority: item.priority || '',
        colorStatusColIndex: colorStatusColIndex,
        colorStatusText: colorStatusText,
        wipRemarks: wipRemarks,
        totalPCS: totalPCS,
        supervisor: item.supervisor || '',
        isStatusUpdated: isStatusUpdated // Add this for reference
      };
    });

    // Extract just the cells for autoTable
    const body = tableRowsData.map(item => item.cells);

    // Create column styles
    const totalColumns = filters.lotStatus === 'Pending' ? 20 : 21;
    const columnStyles = {};
    
    for (let i = 0; i < totalColumns; i++) {
      columnStyles[i] = {
        cellWidth: columnWidths[i],
        halign: 'center',
        valign: 'middle'
      };
    }

    // Track the Y position manually
    let lastAutoTableY = currentY;

    // Create the table with precise styling and custom drawing
    autoTable(doc, {
      startY: currentY,
      head: headers,
      body: body,
      theme: 'grid',
      headStyles: {
        fillColor: headerColor,
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        fontSize: 8,
        cellPadding: { top: 3, right: 2, bottom: 3, left: 2 },
        lineWidth: 0.5,
        lineColor: headerColor,
        halign: 'center',
        valign: 'middle'
      },
      bodyStyles: {
        fontSize: 8,
        cellPadding: { top: 2, right: 1, bottom: 2, left: 1 },
        lineWidth: 0.3,
        lineColor: borderColor,
        textColor: textColor,
        fillColor: [255, 255, 255],
        font: 'times',
        valign: 'middle',
        overflow: 'linebreak',
        minCellHeight: 6,
        lineHeight: 1.1
      },
      styles: {
        fontSize: 8,
        cellPadding: { top: 2, right: 1, bottom: 2, left: 1 },
        lineWidth: 0.3,
        lineColor: borderColor,
        overflow: 'linebreak',
        font: 'times',
        valign: 'middle',
        lineHeight: 1.1
      },
      columnStyles: columnStyles,
      margin: { top: currentY, left: margin, right: margin },
      tableWidth: contentWidth,
      showHead: 'everyPage',
      showFoot: false,
      pageBreak: 'auto',
      rowPageBreak: 'avoid',
      tableLineWidth: 0.5,
      tableLineColor: borderColor,
      
      // Use willDrawCell to customize cell content before drawing
      willDrawCell: function(data) {
        const rowIndex = data.row.index;
        const rowData = tableRowsData[rowIndex];
        
        if (!rowData) return;
        
        const { hasColorPending, isRepeatedLot, lotNumber } = rowData;
        
        // Handle Lot Number column
        if (data.column.index === 1 && (hasColorPending || isRepeatedLot) && data.section === 'body') {
          // Store the original lot number
          data.cell.rawLotNumber = lotNumber;
          // Store priority info
          data.cell.isRepeatedLot = isRepeatedLot;
          data.cell.hasColorPending = hasColorPending;
          // Clear the auto-generated text so we can draw custom
          data.cell.text = [];
        }
      },
      
      // Use didDrawCell to draw custom content after cell is drawn
      didDrawCell: function(data) {
        const rowIndex = data.row.index;
        const rowData = tableRowsData[rowIndex];
        
        if (!rowData) return;
        
        const { hasColorPending, lotNumber, isRepeatedLot } = rowData;
        
        // DRAW SYMBOL FOR LOT NUMBER WITH PROPER CENTERING
        if (data.column.index === 1 && (hasColorPending || isRepeatedLot) && data.section === 'body') {
          const lotNum = lotNumber || '';
          
          if (lotNum) {
            // Save current state
            const originalFont = doc.internal.getFont();
            const originalSize = doc.internal.getFontSize();
            const originalColor = doc.getTextColor();
            
            // Get cell dimensions
            const cellX = data.cell.x;
            const cellY = data.cell.y;
            const cellWidth = data.cell.width;
            const cellHeight = data.cell.height;
            
            // Calculate center of cell
            const centerX = cellX + (cellWidth / 2);
            const centerY = cellY + (cellHeight / 2);
            
            // Determine which symbols to draw
            const iconSize = 2.5;
            const textWidth = doc.getTextWidth(lotNum) || 20;
            
            // Calculate total width based on which symbols to show
            let totalWidth = textWidth;
            let iconCount = 0;
            
            if (hasColorPending) {
              totalWidth += iconSize * 2 + 2;
              iconCount++;
            }
            if (isRepeatedLot) {
              totalWidth += iconSize * 2 + 2;
              iconCount++;
            }
            
            // Position icon and text as a group centered in cell
            const groupStartX = centerX - (totalWidth / 2);
            
            // Track current x position for drawing
            let currentX = groupStartX;
            
            // Draw attention circle for color pending
            if (hasColorPending) {
              const symbolX = currentX + iconSize;
              const symbolY = centerY;
              drawCircleAttentionSymbol(doc, symbolX, symbolY, iconSize);
              currentX += iconSize * 2 + 2; // Move right after icon
            }
            
            // Draw star for repeated lot
            if (isRepeatedLot) {
              const symbolX = currentX + iconSize;
              const symbolY = centerY;
              drawStarSymbol(doc, symbolX, symbolY, iconSize);
              currentX += iconSize * 2 + 2; // Move right after icon
            }
            
            // Draw lot number
            doc.setFont("times", "bold");
            doc.setFontSize(hasColorPending || isRepeatedLot ? 9.5 : 10);
            
            // Set text color
            if (hasColorPending) {
              doc.setTextColor(255, 0, 0);
            } else if (isRepeatedLot) {
              doc.setTextColor(218, 165, 32); // Gold color
            } else {
              doc.setTextColor(255, 0, 0);
            }
            
            const textX = currentX;
            const textY = centerY + 3;
            
            // Truncate if too long
            const maxTextWidth = cellWidth - (iconSize * 2 * iconCount) - (iconCount * 2);
            let displayText = lotNum;
            
            const actualTextWidth = doc.getTextWidth(lotNum);
            if (actualTextWidth > maxTextWidth) {
              for (let i = lotNum.length - 1; i > 5; i--) {
                displayText = lotNum.substring(0, i) + '..';
                if (doc.getTextWidth(displayText) <= maxTextWidth) {
                  break;
                }
              }
            }
            
            doc.text(displayText, textX, textY);
            
            // Restore original state
            doc.setFont(originalFont.fontName, originalFont.fontStyle);
            doc.setFontSize(originalSize);
            doc.setTextColor(originalColor[0], originalColor[1], originalColor[2]);
          }
        }
      },
      
      didDrawPage: function(data) {
        // Draw header on every page
        drawHeader();
        
        // Add page number at bottom
        const pageCount = doc.internal.getNumberOfPages();
        doc.setFontSize(8);
        doc.setTextColor(100, 100, 100);
        doc.text(
          `Page ${data.pageNumber} of ${pageCount}`,
          pageWidth / 2,
          pageHeight - 10,
          { align: 'center' }
        );
        
        // Update last Y position
        if (data.cursor && data.cursor.y) {
          lastAutoTableY = data.cursor.y;
        }
      }
    });

    // ===================== STAGE-WISE ANALYSIS =====================
    
    // Calculate stage-wise analysis
    const stageAnalysis = {};
    
    tableRowsData.forEach(row => {
      const stage = extractStageFromStatus(row.wipRemarks);
      if (!stageAnalysis[stage]) {
        stageAnalysis[stage] = {
          lots: 0,
          pcs: 0,
          colorPendingLots: 0,
          repeatedLots: 0,
          notUpdatedLots: 0
        };
      }
      
      stageAnalysis[stage].lots += 1;
      stageAnalysis[stage].pcs += row.totalPCS;
      
      // Count color pending lots
      const rowData = exportData.find(item => item.lotNumber === row.lotNumber);
      if (rowData && rowData.hasColorPending) {
        stageAnalysis[stage].colorPendingLots += 1;
      }
      
      if (row.isRepeatedLot) {
        stageAnalysis[stage].repeatedLots += 1;
      }
      
      if (!row.isStatusUpdated) {
        stageAnalysis[stage].notUpdatedLots += 1;
      }
    });

    // Convert to array and sort by lots (descending)
    const stageArray = Object.entries(stageAnalysis)
      .map(([stage, data]) => ({
        stage,
        lots: data.lots,
        pcs: data.pcs,
        colorPendingLots: data.colorPendingLots,
        repeatedLots: data.repeatedLots,
        notUpdatedLots: data.notUpdatedLots
      }))
      .sort((a, b) => b.lots - a.lots);

    // Add totals row
    const totalStageLots = stageArray.reduce((sum, item) => sum + item.lots, 0);
    const totalStagePCS = stageArray.reduce((sum, item) => sum + item.pcs, 0);
    const totalStageColorPending = stageArray.reduce((sum, item) => sum + item.colorPendingLots, 0);
    const totalStageRepeated = stageArray.reduce((sum, item) => sum + item.repeatedLots, 0);
    const totalStageNotUpdated = stageArray.reduce((sum, item) => sum + item.notUpdatedLots, 0);

    // ===================== SUPERVISOR WORKLOAD DISTRIBUTION =====================
    
    // Show supervisor summary for ALL reports
    let summaryStartY = lastAutoTableY + 7;

    // Check if we need a new page
    if (summaryStartY > pageHeight - 150) {
      doc.addPage();
      summaryStartY = 20;
      drawHeader();
      currentY = addColorLegend(32);
    }

    // Stage-wise Analysis Section
    doc.setFillColor(255, 255, 255);
    doc.rect(margin - 2, summaryStartY - 8, contentWidth + 4, 16, 'F');

    // Stage Analysis Title
    doc.setFontSize(14);
    doc.setFont('times', 'bold');
    doc.setTextColor(15, 76, 129);
    doc.text('STAGE-WISE ANALYSIS', pageWidth / 2, summaryStartY, { align: 'center' });

    let y = summaryStartY + 15;

    // Prepare stage analysis table data (6 columns now)
    const stageBody = stageArray.map(item => {
      const percentage = totalStageLots > 0 ? Math.round((item.lots / totalStageLots) * 100) : 0;
      
      return [
        { 
          content: item.stage,
          styles: { 
            halign: 'left', 
            fontSize: 9, 
            cellPadding: { top: 3, right: 3, bottom: 3, left: 6 },
            fontStyle: 'bold',
            fillColor: [240, 249, 255]
          } 
        },
        { 
          content: item.lots.toString(), 
          styles: { 
            halign: 'center', 
            fontSize: 10, 
            fontStyle: 'bold',
            cellPadding: { top: 3, right: 3, bottom: 3, left: 3 }
          } 
        },
        { 
          content: `${percentage}%`,
          styles: { 
            halign: 'center', 
            fontSize: 9, 
            cellPadding: { top: 3, right: 3, bottom: 3, left: 3 },
            textColor: percentage >= 50 ? [16, 185, 129] : 
                      percentage >= 30 ? [245, 158, 11] : 
                      [239, 68, 68]
          } 
        },
        { 
          content: item.pcs.toLocaleString(), 
          styles: { 
            halign: 'center', 
            fontSize: 10, 
            fontStyle: 'bold',
            cellPadding: { top: 3, right: 3, bottom: 3, left: 3 }
          } 
        },
        { 
          content: item.colorPendingLots > 0 ? `${item.colorPendingLots}` : '-',
          styles: { 
            halign: 'center', 
            fontSize: 9, 
            cellPadding: { top: 3, right: 3, bottom: 3, left: 3 },
            textColor: item.colorPendingLots > 0 ? [255, 165, 0] : [100, 100, 100]
          } 
        },
        { 
          content: item.notUpdatedLots > 0 ? `${item.notUpdatedLots}` : '-',
          styles: { 
            halign: 'center', 
            fontSize: 9, 
            cellPadding: { top: 3, right: 3, bottom: 3, left: 3 },
            textColor: item.notUpdatedLots > 0 ? [220, 38, 38] : [100, 100, 100],
            fontStyle: item.notUpdatedLots > 0 ? 'bold' : 'normal'
          } 
        }
      ];
    });

    // Add total row for stage analysis
    stageBody.push([
      { 
        content: 'TOTAL', 
        styles: { 
          halign: 'left', 
          fontSize: 10, 
          fontStyle: 'bold', 
          fillColor: [225, 239, 255],
          cellPadding: { top: 4, right: 3, bottom: 4, left: 6 }
        } 
      },
      { 
        content: totalStageLots.toString(), 
        styles: { 
          halign: 'center', 
          fontSize: 11, 
          fontStyle: 'bold', 
          fillColor: [225, 239, 255],
          cellPadding: { top: 4, right: 3, bottom: 4, left: 3 }
        } 
      },
      { 
        content: '100%',
        styles: { 
          halign: 'center', 
          fontSize: 10, 
          fontStyle: 'bold', 
          fillColor: [225, 239, 255],
          textColor: [59, 130, 246],
          cellPadding: { top: 4, right: 3, bottom: 4, left: 3 }
        } 
      },
      { 
        content: totalStagePCS.toLocaleString(), 
        styles: { 
          halign: 'center', 
          fontSize: 11, 
          fontStyle: 'bold', 
          fillColor: [225, 239, 255],
          cellPadding: { top: 4, right: 3, bottom: 4, left: 3 }
        } 
      },
      { 
        content: totalStageColorPending > 0 ? `${totalStageColorPending}` : '-',
        styles: { 
          halign: 'center', 
          fontSize: 10, 
          fontStyle: 'bold', 
          fillColor: [225, 239, 255],
          textColor: totalStageColorPending > 0 ? [255, 165, 0] : [100, 100, 100],
          cellPadding: { top: 4, right: 3, bottom: 4, left: 3 }
        } 
      },
      { 
        content: totalStageNotUpdated > 0 ? `${totalStageNotUpdated}` : '-',
        styles: { 
          halign: 'center', 
          fontSize: 10, 
          fontStyle: 'bold', 
          fillColor: [225, 239, 255],
          textColor: totalStageNotUpdated > 0 ? [220, 38, 38] : [100, 100, 100],
          cellPadding: { top: 4, right: 3, bottom: 4, left: 3 }
        } 
      }
    ]);

    // Calculate stage column widths (6 columns)
    const stageColumnWidths = [
      contentWidth * 0.40,  // Stage
      contentWidth * 0.10,  // Lots
      contentWidth * 0.07,  // %
      contentWidth * 0.15,  // PCS
      contentWidth * 0.14,  // Color Pending
      contentWidth * 0.14   // Not Updated
    ];

    // Create stage analysis table
    autoTable(doc, {
      startY: y,
      head: [[
        { content: 'WORK STAGE', styles: { halign: 'center', fontStyle: 'bold', cellWidth: stageColumnWidths[0], fillColor: stageAnalysisColor, textColor: [255, 255, 255] } },
        { content: 'LOTS', styles: { halign: 'center', fontStyle: 'bold', cellWidth: stageColumnWidths[1], fillColor: stageAnalysisColor, textColor: [255, 255, 255] } },
        { content: '%', styles: { halign: 'center', fontStyle: 'bold', cellWidth: stageColumnWidths[2], fillColor: stageAnalysisColor, textColor: [255, 255, 255] } },
        { content: 'TOTAL PCS', styles: { halign: 'center', fontStyle: 'bold', cellWidth: stageColumnWidths[3], fillColor: stageAnalysisColor, textColor: [255, 255, 255] } },
        { content: 'COLOR PENDING', styles: { halign: 'center', fontStyle: 'bold', cellWidth: stageColumnWidths[4], fillColor: stageAnalysisColor, textColor: [255, 255, 255] } },
        { content: 'NOT UPDATED', styles: { halign: 'center', fontStyle: 'bold', cellWidth: stageColumnWidths[5], fillColor: [220, 38, 38], textColor: [255, 255, 255] } }
      ]],
      body: stageBody.map(row => row.map((cell, colIndex) => ({
        content: cell.content,
        styles: {
          ...cell.styles,
          cellWidth: stageColumnWidths[colIndex]
        }
      }))),
      theme: 'grid',
      headStyles: {
        fillColor: stageAnalysisColor,
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        fontSize: 8,
        cellPadding: { top: 5, right: 3, bottom: 5, left: 3 },
        lineWidth: 0.5,
        lineColor: stageAnalysisColor,
        halign: 'center',
        valign: 'middle'
      },
      bodyStyles: {
        fontSize: 9,
        cellPadding: { top: 3, right: 3, bottom: 3, left: 3 },
        lineWidth: 0.3,
        lineColor: [220, 220, 220],
        textColor: textColor,
        font: 'times',
        valign: 'middle'
      },
      columnStyles: {
        0: { cellWidth: stageColumnWidths[0], halign: 'center' },
        1: { cellWidth: stageColumnWidths[1], halign: 'center' },
        2: { cellWidth: stageColumnWidths[2], halign: 'center' },
        3: { cellWidth: stageColumnWidths[3], halign: 'center' },
        4: { cellWidth: stageColumnWidths[4], halign: 'center' },
        5: { cellWidth: stageColumnWidths[5], halign: 'center' }
      },
      margin: { left: margin, right: margin },
      tableWidth: 'auto',
      showHead: 'everyPage',
      showFoot: false,
      pageBreak: 'auto',
      rowPageBreak: 'avoid'
    });

    // Update Y position after stage analysis
    y = doc.lastAutoTable.finalY + 10;

    // Check if we need a new page for supervisor distribution
    if (y > pageHeight - 120) {
      doc.addPage();
      y = 20;
      drawHeader();
    }

    // Supervisor Workload Distribution Title
    doc.setFillColor(255, 255, 255);
    doc.rect(margin - 2, y - 8, contentWidth + 4, 16, 'F');

    doc.setFontSize(14);
    doc.setFont('times', 'bold');
    doc.setTextColor(15, 76, 129);
    doc.text('SUPERVISOR WORKLOAD DISTRIBUTION', pageWidth / 2, y, { align: 'center' });

    y += 15;

    // Function to normalize supervisor name for matching
    const normalizeForMatching = (name) => {
      if (!name || name.trim() === '') return '';
      return name.trim().toLowerCase();
    };

    // Function to find manpower for a supervisor
    const getManpowerForSupervisor = (supervisor, supervisorManpowerData) => {
      const normalizedSupervisor = normalizeForMatching(supervisor);
      
      if (!supervisorManpowerData) return 0;
      
      // Try exact match first
      for (const [key, value] of Object.entries(supervisorManpowerData)) {
        if (normalizeForMatching(key) === normalizedSupervisor) {
          const manpowerNum = parseInt(value);
          return isNaN(manpowerNum) ? 0 : manpowerNum;
        }
      }
      
      // Try partial match
      for (const [key, value] of Object.entries(supervisorManpowerData)) {
        if (normalizeForMatching(key).includes(normalizedSupervisor) || 
            normalizedSupervisor.includes(normalizeForMatching(key))) {
          const manpowerNum = parseInt(value);
          return isNaN(manpowerNum) ? 0 : manpowerNum;
        }
      }
      
      return 0;
    };

    // Calculate supervisor summary by stitching days categories
    const supervisorSummary = exportData.reduce((acc, item) => {
      const supervisor = normalizeSupervisorName(item.supervisor);
      
      if (!acc[supervisor]) {
        acc[supervisor] = {
          totalLots: 0,
          totalPCS: 0,
          greenLots: 0,
          greenPCS: 0,
          yellowLots: 0,
          yellowPCS: 0,
          redLots: 0,
          redPCS: 0,
          repeatedLots: 0,
          colorPendingLots: 0,
          notUpdatedLots: 0,
          manpower: 0
        };
      }
      
      acc[supervisor].totalLots += 1;
      acc[supervisor].totalPCS += item.totalPCS || 0;
      
      // Check for repeated lot
      const isRepeat = item.priority && normalizeText(item.priority).includes('repeated_lot');
      if (isRepeat) {
        acc[supervisor].repeatedLots += 1;
      }
      
      // Check for color pending
      if (item.hasColorPending) {
        acc[supervisor].colorPendingLots += 1;
      }
      
      // Check if status not updated
      if (!isStatusUpdatedToday(item.wipStatus)) {
        acc[supervisor].notUpdatedLots += 1;
      }
      
      // Calculate stitching days for this lot
      const stitchingDays = calculateStitchingDays(item.dateOfIssue);
      
      // Categorize by stitching days
      if (stitchingDays <= 6) {
        acc[supervisor].greenLots += 1;
        acc[supervisor].greenPCS += item.totalPCS || 0;
      } else if (stitchingDays <= 15) {
        acc[supervisor].yellowLots += 1;
        acc[supervisor].yellowPCS += item.totalPCS || 0;
      } else {
        acc[supervisor].redLots += 1;
        acc[supervisor].redPCS += item.totalPCS || 0;
      }
      
      return acc;
    }, {});

    // Sort supervisors by total lots (descending)
    const sortedSupervisors = Object.entries(supervisorSummary)
      .sort(([, a], [, b]) => b.totalLots - a.totalLots);

    // Get latest manpower data for each supervisor
    const supervisorManpowerData = getLatestManpowerBySupervisor(exportData);

    // Update manpower for each supervisor
    sortedSupervisors.forEach(([supervisor, data]) => {
      data.manpower = getManpowerForSupervisor(supervisor, supervisorManpowerData);
    });

    // Prepare simplified summary body with not updated column
    const summaryBody = sortedSupervisors.map(([supervisor, data]) => {
      const manpower = data.manpower || 0;
      const avgPCS = manpower > 0 ? Math.round(data.totalPCS / manpower) : 0;
      
      // Calculate percentages for each category
      const greenPercent = data.greenLots > 0 ? Math.round((data.greenLots / data.totalLots) * 100) : 0;
      const yellowPercent = data.yellowLots > 0 ? Math.round((data.yellowLots / data.totalLots) * 100) : 0;
      const redPercent = data.redLots > 0 ? Math.round((data.redLots / data.totalLots) * 100) : 0;
      
      return [
        { 
          content: supervisor,
          styles: { 
            halign: 'center', 
            fontSize: 12, 
            cellPadding: { top: 3, right: 3, bottom: 3, left: 6 },
            fontStyle: 'bold'
          } 
        },
        { 
          content: manpower > 0 ? manpower.toString() : 'N/A',
          styles: { 
            halign: 'center', 
            fontSize: 12, 
            cellPadding: { top: 3, right: 3, bottom: 3, left: 3 },
            fontStyle: 'bold',
            fillColor: [240, 248, 255],
            textColor: [15, 76, 129]
          } 
        },
        { 
          content: data.totalLots.toString(), 
          styles: { 
            halign: 'center', 
            fontSize: 12, 
            fontStyle: 'bold',
            cellPadding: { top: 3, right: 3, bottom: 3, left: 3 }
          } 
        },
        { 
          content: data.totalPCS.toLocaleString(), 
          styles: { 
            halign: 'center', 
            fontSize: 12, 
            fontStyle: 'bold',
            cellPadding: { top: 3, right: 3, bottom: 3, left: 3 }
          } 
        },
        {
          content: avgPCS > 0 ? avgPCS.toLocaleString() : 'N/A',
          styles: { 
            halign: 'center', 
            fontSize: 12, 
            fontStyle: 'bold',
            cellPadding: { top: 3, right: 3, bottom: 3, left: 3 },
            fillColor: [255, 250, 240],
            textColor: [245, 158, 11]
          } 
        },
        { 
          content: data.notUpdatedLots > 0 ? data.notUpdatedLots.toString() : '-',
          styles: { 
            halign: 'center', 
            fontSize: 12, 
            cellPadding: { top: 3, right: 3, bottom: 3, left: 3 },
            fontStyle: data.notUpdatedLots > 0 ? 'bold' : 'normal',
            textColor: data.notUpdatedLots > 0 ? [220, 38, 38] : [100, 100, 100],
            fillColor: data.notUpdatedLots > 0 ? [255, 235, 235] : [255, 255, 255]
          } 
        },
        { 
          content: data.greenLots > 0 ? `${data.greenLots} (${greenPercent}%)${data.colorPendingLots > 0 ? ' ⚡' : ''}` : '-',
          styles: { 
            halign: 'center', 
            fontSize: 12, 
            cellPadding: { top: 3, right: 3, bottom: 3, left: 3 },
            fontStyle: 'bold',
            textColor: [16, 185, 129]
          } 
        },
        { 
          content: data.yellowLots > 0 ? `${data.yellowLots} (${yellowPercent}%)${data.colorPendingLots > 0 ? ' ⚡' : ''}` : '-',
          styles: { 
            halign: 'center', 
            fontSize: 12, 
            cellPadding: { top: 3, right: 3, bottom: 3, left: 3 },
            fontStyle: 'bold',
            textColor: [245, 158, 11]
          } 
        },
        { 
          content: data.redLots > 0 ? `${data.redLots} (${redPercent}%)${data.colorPendingLots > 0 ? ' ⚡' : ''}` : '-',
          styles: { 
            halign: 'center', 
            fontSize: 12, 
            cellPadding: { top: 3, right: 3, bottom: 3, left: 3 },
            fontStyle: 'bold',
            textColor: [239, 68, 68]
          } 
        }
      ];
    });

    // Calculate overall totals
    const overallTotals = exportData.reduce((acc, item) => {
      const stitchingDays = calculateStitchingDays(item.dateOfIssue);
      
      acc.totalLots += 1;
      acc.totalPCS += item.totalPCS || 0;
      
      // Check for color pending
      if (item.hasColorPending) {
        acc.colorPendingLots += 1;
      }
      
      // Check for not updated
      if (!isStatusUpdatedToday(item.wipStatus)) {
        acc.notUpdatedLots += 1;
      }
      
      if (stitchingDays <= 6) {
        acc.greenLots += 1;
        acc.greenPCS += item.totalPCS || 0;
      } else if (stitchingDays <= 15) {
        acc.yellowLots += 1;
        acc.yellowPCS += item.totalPCS || 0;
      } else {
        acc.redLots += 1;
        acc.redPCS += item.totalPCS || 0;
      }
      
      return acc;
    }, {
      totalLots: 0,
      totalPCS: 0,
      greenLots: 0,
      greenPCS: 0,
      yellowLots: 0,
      yellowPCS: 0,
      redLots: 0,
      redPCS: 0,
      colorPendingLots: 0,
      notUpdatedLots: 0
    });

    // Add total row
    const totalRowBackground = [225, 239, 255];
    const totalManpower = sortedSupervisors.reduce((sum, [, data]) => sum + (data.manpower || 0), 0);
    const overallAvgPCS = totalManpower > 0 ? Math.round(overallTotals.totalPCS / totalManpower) : 0;

    // Calculate overall percentages
    const overallGreenPercent = overallTotals.greenLots > 0 ? Math.round((overallTotals.greenLots / overallTotals.totalLots) * 100) : 0;
    const overallYellowPercent = overallTotals.yellowLots > 0 ? Math.round((overallTotals.yellowLots / overallTotals.totalLots) * 100) : 0;
    const overallRedPercent = overallTotals.redLots > 0 ? Math.round((overallTotals.redLots / overallTotals.totalLots) * 100) : 0;

    summaryBody.push([
      { 
        content: 'TOTALS', 
        styles: { 
          halign: 'center', 
          fontSize: 11, 
          fontStyle: 'bold', 
          fillColor: totalRowBackground,
          cellPadding: { top: 5, right: 3, bottom: 5, left: 6 }
        } 
      },
      { 
        content: totalManpower > 0 ? totalManpower.toString() : '-', 
        styles: { 
          halign: 'center', 
          fontSize: 11, 
          fontStyle: 'bold', 
          fillColor: totalRowBackground,
          cellPadding: { top: 5, right: 3, bottom: 5, left: 3 }
        } 
      },
      { 
        content: overallTotals.totalLots.toString(), 
        styles: { 
          halign: 'center', 
          fontSize: 11, 
          fontStyle: 'bold', 
          fillColor: totalRowBackground,
          cellPadding: { top: 5, right: 3, bottom: 5, left: 3 }
        } 
      },
      { 
        content: overallTotals.totalPCS.toLocaleString(), 
        styles: { 
          halign: 'center', 
          fontSize: 11, 
          fontStyle: 'bold', 
          fillColor: totalRowBackground,
          cellPadding: { top: 5, right: 3, bottom: 5, left: 3 }
        } 
      },
      {
        content: overallAvgPCS > 0 ? overallAvgPCS.toLocaleString() : '-',
        styles: { 
          halign: 'center', 
          fontSize: 11, 
          fontStyle: 'bold', 
          fillColor: totalRowBackground,
          textColor: [245, 158, 11],
          cellPadding: { top: 5, right: 3, bottom: 5, left: 3 }
        } 
      },
      { 
        content: overallTotals.notUpdatedLots > 0 ? overallTotals.notUpdatedLots.toString() : '-',
        styles: { 
          halign: 'center', 
          fontSize: 11, 
          fontStyle: 'bold', 
          fillColor: totalRowBackground,
          textColor: overallTotals.notUpdatedLots > 0 ? [220, 38, 38] : [100, 100, 100],
          cellPadding: { top: 5, right: 3, bottom: 5, left: 3 }
        } 
      },
      { 
        content: overallTotals.greenLots > 0 ? `${overallTotals.greenLots} (${overallGreenPercent}%)${overallTotals.colorPendingLots > 0 ? ' ⚡' : ''}` : '-',
        styles: { 
          halign: 'center', 
          fontSize: 11, 
          fontStyle: 'bold', 
          fillColor: totalRowBackground,
          textColor: [16, 185, 129],
          cellPadding: { top: 5, right: 3, bottom: 5, left: 3 }
        } 
      },
      { 
        content: overallTotals.yellowLots > 0 ? `${overallTotals.yellowLots} (${overallYellowPercent}%)${overallTotals.colorPendingLots > 0 ? ' ⚡' : ''}` : '-',
        styles: { 
          halign: 'center', 
          fontSize: 11, 
          fontStyle: 'bold', 
          fillColor: totalRowBackground,
          textColor: [245, 158, 11],
          cellPadding: { top: 5, right: 3, bottom: 5, left: 3 }
        } 
      },
      { 
        content: overallTotals.redLots > 0 ? `${overallTotals.redLots} (${overallRedPercent}%)${overallTotals.colorPendingLots > 0 ? ' ⚡' : ''}` : '-',
        styles: { 
          halign: 'center', 
          fontSize: 11, 
          fontStyle: 'bold', 
          fillColor: totalRowBackground,
          textColor: [239, 68, 68],
          cellPadding: { top: 5, right: 3, bottom: 5, left: 3 }
        } 
      }
    ]);

    // Calculate column widths for supervisor summary table - 9 COLUMNS NOW
    const totalAvailableWidth = contentWidth;
    const numberOfColumns = 9;
    const baseColumnWidth = totalAvailableWidth / numberOfColumns;

    const summaryColumnWidths = [
      baseColumnWidth * 1.0,  // Supervisor
      baseColumnWidth * 0.7,  // Manpower
      baseColumnWidth * 0.7,  // Total Lots
      baseColumnWidth * 0.8,  // Total PCS
      baseColumnWidth * 0.9,  // Avg PCS
      baseColumnWidth * 0.9,  // Not Updated
      baseColumnWidth * 1.1,  // Green
      baseColumnWidth * 1.1,  // Yellow
      baseColumnWidth * 1.1   // Red
    ];

    // Create supervisor summary table - 9 COLUMNS
    autoTable(doc, {
      startY: y,
      head: [[
        { content: 'SUPERVISOR', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[0] } },
        { content: 'MANPOWER', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[1], fillColor: [15, 76, 129], textColor: [255, 255, 255] } },
        { content: 'TOTAL LOTS', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[2] } },
        { content: 'TOTAL PCS', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[3] } },
        { content: 'AVG PCS/MAN', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[4], fillColor: [245, 158, 11], textColor: [255, 255, 255] } },
        { content: 'NOT UPDATED', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[5], fillColor: [220, 38, 38], textColor: [255, 255, 255] } },
        { content: 'GREEN (1-6 Days)', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[6], textColor: [16, 185, 129] } },
        { content: 'YELLOW (7-15 Days)', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[7], textColor: [245, 158, 11] } },
        { content: 'RED (15+ Days)', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[8], textColor: [239, 68, 68] } }
      ]],
      body: summaryBody,
      theme: 'grid',
      headStyles: {
        fillColor: [15, 76, 129],
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        fontSize: 8,
        cellPadding: { top: 5, right: 3, bottom: 5, left: 3 },
        lineWidth: 0.5,
        lineColor: [15, 76, 129],
        halign: 'center',
        valign: 'middle'
      },
      bodyStyles: {
        fontSize: 11,
        cellPadding: { top: 3, right: 3, bottom: 3, left: 3 },
        lineWidth: 0.3,
        lineColor: [220, 220, 220],
        textColor: textColor,
        font: 'times',
        valign: 'middle'
      },
      columnStyles: {
        0: { cellWidth: summaryColumnWidths[0], halign: 'center' },
        1: { cellWidth: summaryColumnWidths[1], halign: 'center' },
        2: { cellWidth: summaryColumnWidths[2], halign: 'center' },
        3: { cellWidth: summaryColumnWidths[3], halign: 'center' },
        4: { cellWidth: summaryColumnWidths[4], halign: 'center' },
        5: { cellWidth: summaryColumnWidths[5], halign: 'center' },
        6: { cellWidth: summaryColumnWidths[6], halign: 'center' },
        7: { cellWidth: summaryColumnWidths[7], halign: 'center' },
        8: { cellWidth: summaryColumnWidths[8], halign: 'center' }
      },
      margin: { left: margin, right: margin },
      tableWidth: 'auto',
      showHead: 'everyPage',
      showFoot: false,
      pageBreak: 'auto',
      rowPageBreak: 'avoid'
    });

    // Add COMPACT notes section
    const summaryEndY = doc.lastAutoTable.finalY;
    const notesY = summaryEndY + 4;

    doc.setFontSize(6.5);
    doc.setFont('times', 'normal');
    doc.setTextColor(100, 100, 100);

    // Compact notes in a single line
    const notes = [
      'Key: Green=1-6 Days • Yellow=7-15 Days • Red=15+ Days • ⚡=Color Pending lots in category • ★=Repeated Lot • Avg PCS/Man=Productivity • Numbers in ()=% of supervisor\'s total lots • NOT UPDATED=No WIP status update today (light red background)'
    ];

    notes.forEach((note, index) => {
      doc.text(note, margin, notesY + (index * 3));
    });

    // Generate filename
    const today = new Date();
    let fileName = 'Daily_Stitching_Report';
    
    // Add date
    fileName += `_${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
    
    // Add supervisor filter to filename
    if (filters.supervisor && filters.supervisor.trim() !== '') {
      const supervisorName = filters.supervisor.replace(/\s+/g, '_');
      fileName += `_${supervisorName}`;
    }
    
    // Add status filter to filename
    if (filters.lotStatus) {
      fileName += `_${filters.lotStatus}`;
    }
    
    // Add record count
    fileName += `_${exportData.length}_lots`;
    
    // Clean filename
    fileName = fileName.replace(/[^\w\-_]/g, '_').substring(0, 80);
    
    doc.save(`${fileName}.pdf`);
  } catch (err) {
    console.error('PDF Export Error:', err);
    setError(`Failed to export PDF: ${err.message}`);
  } finally {
    setLoading(false);
  }
}, [filteredData, filters, calculateStitchingDays, getStitchingDaysColor, getStitchingDaysTextColor, getEmbPrintDate, getLatestWipRemarks, getCompletionDateFormatted, isLotCompleted, getLatestManpowerBySupervisor, normalizeText]);

  const handleRefresh = useCallback(async () => {
    try {
      setLoading(true);
      setError(null);
      
      // Clear cache for current supervisor
      const cacheKey = getCacheKey(filters.supervisor);
      dataCache.current.delete(cacheKey);
      lastFetchTime.current.delete(cacheKey);
      
      const freshData = await fetchDataForSupervisor(filters.supervisor);
      setData(freshData);
      
      // If supervisor is Monu, show only pending by default
      if (filters.supervisor === 'Monu' && filters.lotStatus === 'Pending') {
        const pendingData = freshData.filter(item => !isLotCompleted(item.completedStatus));
        setFilteredData(pendingData);
      } else {
        applyFilters(filters, freshData);
      }
      
      // Update filter options
      extractFilterOptions(freshData);
      
    } catch (err) {
      setError(`Failed to refresh data: ${err.message}`);
    } finally {
      setLoading(false);
    }
  }, [filters, fetchDataForSupervisor, applyFilters, extractFilterOptions, getCacheKey, isLotCompleted]);

  // Load all supervisors for dropdown when component mounts
  useEffect(() => {
    const loadSupervisorOptions = async () => {
      try {
        const allData = await loadAllData();
        const allSupervisors = [...new Set(allData
          .filter(item => item.supervisor && item.supervisor.trim() !== '')
          .map(item => normalizeAndCapitalize(item.supervisor))
        )].sort();
        
        setFilterOptions(prev => ({
          ...prev,
          supervisor: allSupervisors
        }));
      } catch (err) {
        console.error('Error loading supervisor options:', err);
      }
    };

    loadSupervisorOptions();
  }, [loadAllData, normalizeAndCapitalize]);

  // Memoize render functions
  const renderDropdownFilter = useCallback((fieldName, label) => (
    <div key={fieldName} className="filter-item">
      <label className="filter-label">
        {label}
      </label>
      <select
        name={fieldName}
        value={filters[fieldName]}
        onChange={handleFilterChange}
        className="filter-select"
      >
        <option value="">All {label}</option>
        {filterOptions[fieldName].map((option, index) => (
          <option key={index} value={option}>
            {option}
          </option>
        ))}
      </select>
    </div>
  ), [filters, filterOptions, handleFilterChange]);

  // Loading and error states
  if (isInitialLoad && loading) {
    return (
      <div className="loading-container">
        <div className="spinner"></div>
        <p>Loading Monu's stitching data...</p>
        <p className="loading-subtext">Loading pending lots by default</p>
      </div>
    );
  }

  if (error) {
    return (
      <div className="error-container">
        <div className="error-content">
          <strong>Error:</strong> {error}
          <button 
            onClick={handleRefresh}
            className="retry-btn"
          >
            Retry
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="container">
      <div className="header">
        <div className="header-top">
          <div className="title-section">
            <h1>
              <span className="title-icon">📊</span>
              Daily Stitching Issue Report
            </h1>
            <p className="subtitle">
              {filters.supervisor ? `${filters.supervisor} - Stitching Report` : 'Real-time tracking of stitching operations'}
              {data.length > 0 && (
                <span className="data-source">
                  • {filters.supervisor === 'Monu' && filters.lotStatus === 'Pending' ? 
                     "Showing Monu's PENDING lots by default" : 
                     filters.lotStatus ? `Showing ${filters.lotStatus.toLowerCase()} lots` : 
                     `Showing all lots for ${filters.supervisor}`}
                  {Object.keys(colorPendingLots).length > 0 && ` • ${Object.keys(colorPendingLots).length} lots with color pending`}
                </span>
              )}
            </p>
          </div>
          
          <div className="stats-card">
            <div className="stat-item">
              <div className="stat-value">{filteredData.length}</div>
              <div className="stat-label">Filtered Lots</div>
            </div>
            <div className="stat-item">
              <div className="stat-value">
                {data.filter(item => isLotCompleted(item.completedStatus)).length}
              </div>
              <div className="stat-label">Completed</div>
            </div>
            <div className="stat-item">
              <div className="stat-value">{data.length}</div>
              <div className="stat-label">Total Lots</div>
            </div>
            <div className="stat-item">
              <div className="stat-value">
                {data.filter(item => item.hasColorPending).length}
              </div>
              <div className="stat-label">Color Pending</div>
            </div>
          </div>
        </div>

        <div className="header-actions">
          <div className="action-group">
            <button
              onClick={() => window.history.back()}
              className="btn btn-back"
            >
              <span className="btn-icon">←</span>
              Back
            </button>
            <button
              onClick={downloadExcel}
              className="btn btn-excel"
              disabled={filteredData.length === 0 || loading}
            >
              <span className="btn-icon">📊</span>
              Export Excel
              {(!filters.supervisor || filters.supervisor === '') && ' (All)'}
            </button>
            
            <button
              onClick={downloadPDF}
              className="btn btn-pdf"
              disabled={filteredData.length === 0 || loading}
            >
              <span className="btn-icon">📄</span>
              Export PDF
              {(!filters.supervisor || filters.supervisor === '') && ' (All)'}
            </button>
          </div>
          
          <div className="refresh-group">
            <button
              onClick={handleRefresh}
              className="btn btn-refresh"
              disabled={loading}
            >
              <span className="btn-icon">↻</span>
              {loading ? 'Refreshing...' : 'Refresh'}
            </button>
            <div className="cache-info">
              Cache: {dataCache.current.size} supervisor(s)
            </div>
          </div>
        </div>
      </div>

      {/* Filters Section */}
      <div className="filters-section">
        <div className="filters-header">
          <h3>Filters</h3>
          <div className="filters-info">
            <span className="active-filters">
              {Object.values(filters).filter(f => f && f.trim() !== '').length} active filters
            </span>
            <button
              onClick={clearFilters}
              className="btn-clear"
              disabled={loading}
            >
              {filters.supervisor === 'Monu' ? 'Clear Filters' : 'Reset to Monu'}
            </button>
          </div>
        </div>
        
        <div className="filters-grid">
          <div className="filter-group">
            <label className="group-label">Basic Filters</label>
            <div className="group-filters">
              <div className="filter-item">
                <label>Lot Number</label>
                <input
                  type="text"
                  name="lotNumber"
                  value={filters.lotNumber}
                  onChange={handleFilterChange}
                  placeholder="Search lot numbers..."
                  className="filter-input"
                  disabled={loading}
                />
              </div>
              
              <div className="filter-item">
                <label>Supervisor</label>
                <select
                  name="supervisor"
                  value={filters.supervisor}
                  onChange={handleFilterChange}
                  className="filter-select"
                  disabled={loading}
                >
                  <option value="Monu">Monu (Default)</option>
                  {filterOptions.supervisor
                    .filter(s => s !== 'Monu')
                    .map((option, index) => (
                      <option key={index} value={option}>
                        {option}
                      </option>
                    ))}
                </select>
                <div className="filter-hint">
                  Changing supervisor will load new data
                </div>
              </div>
              
              <div className="filter-item">
                <label>Party Name</label>
                <select
                  name="partyName"
                  value={filters.partyName}
                  onChange={handleFilterChange}
                  className="filter-select"
                  disabled={loading}
                >
                  <option value="">All Parties</option>
                  {filterOptions.partyName.map((option, index) => (
                    <option key={index} value={option}>
                      {option}
                    </option>
                  ))}
                </select>
              </div>
            </div>
          </div>
          
          <div className="filter-group">
            <label className="group-label">Product Details</label>
            <div className="group-filters">
              {renderDropdownFilter('fabric', 'Fabric')}
              {renderDropdownFilter('garmentType', 'Garment Type')}
              {renderDropdownFilter('style', 'Style')}
            </div>
          </div>
          
        {/* In the filters-grid, add a new filter-group or add to existing group */}
<div className="filter-group">
  <label className="group-label">Additional Filters</label>
  <div className="group-filters">
    {renderDropdownFilter('brand', 'Brand')}
    {renderDropdownFilter('season', 'Season')}
    {renderDropdownFilter('mwk', 'M/W/K')}
  </div>
</div>

<div className="filter-group">
  <label className="group-label">Priority & Status Filters</label>
  <div className="group-filters">
    {/* DYNAMIC PRIORITY FILTER */}
    <div className="filter-item">
      <label>Priority</label>
      <select
        name="priority"
        value={filters.priority}
        onChange={handleFilterChange}
        className="filter-select"
        disabled={loading}
      >
        <option value="">All Priorities</option>
        {filterOptions.priority && filterOptions.priority.length > 0 ? (
          filterOptions.priority.map((option, index) => (
            <option key={index} value={option}>
              {option}
            </option>
          ))
        ) : (
          <option value="" disabled>No priority data available</option>
        )}
      </select>
    </div>
    
    {renderDropdownFilter('wipStatus', 'WIP Status')}
    
    <div className="filter-item">
      <label>Lot Status</label>
      <select
        name="lotStatus"
        value={filters.lotStatus}
        onChange={handleFilterChange}
        className="filter-select"
        disabled={loading}
      >
        <option value="">All Lots</option>
        <option value="Pending">Pending Only</option>
        <option value="Completed">Completed Only</option>
      </select>
    </div>
    
    <div className="filter-item">
      <label>Direct Stitching</label>
      <select
        name="directStitching"
        value={filters.directStitching}
        onChange={handleFilterChange}
        className="filter-select"
        disabled={loading}
      >
        <option value="">All</option>
        <option value="Yes">Yes</option>
        <option value="No">No</option>
      </select>
    </div>
  </div>
</div>
        </div>
      </div>

      {/* Loading overlay for data fetching */}
      {loading && !isInitialLoad && (
        <div className="loading-overlay">
          <div className="loading-content">
            <div className="spinner small"></div>
            <p>Loading {filters.supervisor}'s data...</p>
          </div>
        </div>
      )}

      {filteredData.length === 0 ? (
        <div className="empty-state">
          <div className="empty-icon">📋</div>
          <h3>No Data Found</h3>
          <p>
            {data.length === 0 
              ? `No stitching issue data found for ${filters.supervisor}.` 
              : 'No records match the current filters.'}
          </p>
          <button onClick={clearFilters} className="btn btn-primary" disabled={loading}>
            {filters.supervisor === 'Monu' ? 'Clear Filters' : 'Reset to Monu'}
          </button>
        </div>
      ) : (
        <>
          <div className="table-section">
            <div className="table-header">
              <div className="table-info">
                Showing <strong>{filteredData.length}</strong> records for {filters.supervisor}
                {Object.values(filters).some(f => f && f.trim() !== '' && f !== 'supervisor') && ' (filtered)'}
                {data.filter(item => item.hasColorPending).length > 0 && ` • ${data.filter(item => item.hasColorPending).length} lots with color pending`}
              </div>
              <div className="table-legend">
                <div className="legend-item">
                  <span className="legend-color green"></span>
                  <span>1-6 days</span>
                </div>
                <div className="legend-item">
                  <span className="legend-color yellow"></span>
                  <span>7-15 days</span>
                </div>
                <div className="legend-item">
                  <span className="legend-color red"></span>
                  <span>15 days</span>
                </div>
                <div className="legend-item">
                  <span className="legend-color orange"></span>
                  <span>Color Pending</span>
                </div>
                <div className="legend-item">
                  <span className="legend-color completed"></span>
                  <span>Completed</span>
                </div>
                <div className="legend-item">
                  <span className="legend-color pending"></span>
                  <span>Pending</span>
                </div>
              </div>
            </div>
            
            <div className="table-container">
              <table className="data-table">
                <thead>
                  <tr>
                    {HEADERS.map((header, index) => (
                      <th key={index}>{header}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filteredData.map((item, index) => {
                    const stitchingDays = calculateStitchingDays(item.dateOfIssue);
                    const stitchingDaysColor = getStitchingDaysColor(stitchingDays);
                    const stitchingDaysTextColor = getStitchingDaysTextColor(stitchingDays);
                    const embPrintDate = getEmbPrintDate(item.challanHistory);
                    const isCompleted = isLotCompleted(item.completedStatus);
                    const completionDate = getCompletionDateFormatted(item.completedStatus);
                    const hasColorPending = item.hasColorPending || false;
                    
                    return (
                      <tr 
                        key={index} 
                        className={`table-row ${hasColorPending ? 'color-pending-row' : ''}`}
                        style={hasColorPending ? { backgroundColor: '#fff3e0' } : {}}
                      >
                        <td className="text-center font-semibold">{index + 1}</td>
                     <td className="text-center">
  <span 
    className={`lot-number ${hasColorPending ? 'color-pending' : ''} ${item.priority?.includes('REPEATED_LOT') ? 'repeated-lot' : ''}`}
    style={hasColorPending ? { color: '#ff0000', fontWeight: 'bold' } : 
           (item.priority?.includes('REPEATED_LOT') ? { color: '#d4af37', fontWeight: 'bold' } : {})}
  >
    {item.lotNumber || 'N/A'}
    {hasColorPending && ' ⚠️'}
    {item.priority?.includes('REPEATED_LOT') && ' ★'}
  </span>
</td>
                        <td className="text-center">{item.garmentType || 'N/A'}</td>
                        <td className="text-center">{item.style || 'N/A'}</td>
                        <td className="text-center">{item.fabric || 'N/A'}</td>
                        <td className="text-center">{item.brand || 'N/A'}</td>
                        <td className="text-center font-semibold">
                          <span className="total-pcs">
                            {item.totalPCS > 0 ? item.totalPCS.toLocaleString() : 'N/A'}
                          </span>
                        </td>
                        <td className="text-center">{item.mwk || 'N/A'}</td>
                        <td className="text-center">{item.season || 'N/A'}</td>
                        <td className="text-center">{item.partyName || 'N/A'}</td>
                        <td className="text-center">{item.directStitching || 'N/A'}</td>
                        <td className="text-center">
                          <span 
                            className="stitching-days"
                            style={{
                              backgroundColor: stitchingDaysColor,
                              color: stitchingDaysTextColor
                            }}
                          >
                            {stitchingDays}
                          </span>
                        </td>
                        <td className="text-center">
                          <span className={`emb-date ${embPrintDate === '-' ? 'no-date' : 'has-date'}`}>
                            {embPrintDate}
                          </span>
                        </td>
                        <td className="text-center">
                          <span className="wip-status">
                            {getLatestWipRemarks(item.wipStatus)}
                          </span>
                        </td>
                        <td className="text-center">
                          <span className={`completion-date ${isCompleted ? 'completed' : 'pending'}`}>
                            {completionDate || (isCompleted ? 'Completed' : 'Pending')}
                          </span>
                        </td>
                        <td className="text-center">
                          <span className={`status-badge ${isCompleted ? 'completed-badge' : 'pending-badge'} ${hasColorPending ? 'color-pending-badge' : ''}`}>
                            {isCompleted ? 'Completed' : 'Pending'}
                            {hasColorPending && ' (Color Pending)'}
                          </span>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            
            <div className="table-footer">
              <div className="footer-info">
                Showing <strong>{filteredData.length}</strong> of <strong>{data.length}</strong> records for {filters.supervisor}
                <span className="footer-stats">
                  • {data.filter(item => isLotCompleted(item.completedStatus)).length} completed • {data.filter(item => !isLotCompleted(item.completedStatus)).length} pending
                  • {data.filter(item => item.hasColorPending).length} color pending
                </span>
              </div>
              <div className="footer-actions">
                <button
                  onClick={downloadExcel}
                  className="btn btn-excel btn-small"
                  disabled={filteredData.length === 0 || loading}
                >
                  📊 Excel
                </button>
                <button
                  onClick={downloadPDF}
                  className="btn btn-pdf btn-small"
                  disabled={filteredData.length === 0 || loading}
                >
                  📄 PDF
                </button>
              </div>
            </div>
          </div>
        </>
      )}

      <style jsx>{`
        .container {
          min-height: 100vh;
          background: #f8fafc;
          padding: 20px;
          position: relative;
        }

        .loading-container {
          display: flex;
          justify-content: center;
          align-items: center;
          height: 100vh;
          flex-direction: column;
          gap: 20px;
        }

        .loading-subtext {
          font-size: 14px;
          color: #64748b;
          margin-top: 8px;
        }

        .spinner {
          width: 50px;
          height: 50px;
          border: 3px solid #e2e8f0;
          border-top-color: #3b82f6;
          border-radius: 50%;
          animation: spin 1s linear infinite;
        }

        .spinner.small {
          width: 30px;
          height: 30px;
          border-width: 2px;
        }

        @keyframes spin {
          to { transform: rotate(360deg); }
        }

        .error-container {
          background: white;
          border-radius: 12px;
          padding: 40px;
          margin: 20px;
          box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        }

        .error-content {
          text-align: center;
          color: #dc2626;
        }

        .retry-btn {
          margin-top: 20px;
          padding: 10px 24px;
          background: #3b82f6;
          color: white;
          border: none;
          border-radius: 8px;
          font-weight: 600;
          cursor: pointer;
          transition: all 0.2s;
        }

        .retry-btn:hover {
          background: #2563eb;
          transform: translateY(-1px);
        }

        /* Header Styles */
        .header {
          background: white;
          border-radius: 16px;
          padding: 24px;
          margin-bottom: 24px;
          box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
          border: 1px solid #e2e8f0;
        }

        .header-top {
          display: flex;
          justify-content: space-between;
          align-items: flex-start;
          margin-bottom: 24px;
        }

        .title-section h1 {
          font-size: 28px;
          font-weight: 700;
          color: #1e293b;
          margin: 0;
          display: flex;
          align-items: center;
          gap: 12px;
        }

        .title-icon {
          font-size: 32px;
        }

        .subtitle {
          font-size: 14px;
          color: #64748b;
          margin-top: 8px;
          display: flex;
          align-items: center;
          gap: 8px;
        }

        .data-source {
          font-size: 12px;
          background: #f1f5f9;
          padding: 2px 8px;
          border-radius: 12px;
        }

        .stats-card {
          display: flex;
          gap: 24px;
          background: #f1f5f9;
          padding: 16px 24px;
          border-radius: 12px;
          border: 1px solid #e2e8f0;
        }

        .stat-item {
          text-align: center;
          min-width: 80px;
        }

        .stat-value {
          font-size: 28px;
          font-weight: 700;
          color: #3b82f6;
          line-height: 1;
        }

        .stat-label {
          font-size: 12px;
          color: #64748b;
          margin-top: 4px;
          font-weight: 500;
          text-transform: uppercase;
          letter-spacing: 0.05em;
        }

        .header-actions {
          display: flex;
          justify-content: space-between;
          align-items: center;
          padding-top: 20px;
          border-top: 1px solid #e2e8f0;
        }

        .action-group {
          display: flex;
          gap: 12px;
        }

        .refresh-group {
          display: flex;
          align-items: center;
          gap: 12px;
        }

        .cache-info {
          font-size: 12px;
          color: #64748b;
          background: #f1f5f9;
          padding: 4px 12px;
          border-radius: 12px;
        }

        .btn {
          display: flex;
          align-items: center;
          gap: 8px;
          padding: 10px 20px;
          border: none;
          border-radius: 8px;
          font-weight: 600;
          font-size: 14px;
          cursor: pointer;
          transition: all 0.2s;
        }

        .btn:hover:not(:disabled) {
          transform: translateY(-1px);
        }

        .btn:disabled {
          opacity: 0.5;
          cursor: not-allowed;
        }

        .btn-back {
          background: #6b7280;
          color: white;
        }

        .btn-back:hover:not(:disabled) {
          background: #4b5563;
        }

        .btn-excel {
          background: #10b981;
          color: white;
        }

        .btn-excel:hover:not(:disabled) {
          background: #059669;
          box-shadow: 0 4px 12px rgba(16, 185, 129, 0.3);
        }

        .btn-pdf {
          background: #ef4444;
          color: white;
        }

        .btn-pdf:hover:not(:disabled) {
          background: #dc2626;
          box-shadow: 0 4px 12px rgba(239, 68, 68, 0.3);
        }

        .btn-refresh {
          background: #3b82f6;
          color: white;
        }

        .btn-refresh:hover:not(:disabled) {
          background: #2563eb;
          box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3);
        }

        .btn-small {
          padding: 6px 12px;
          font-size: 13px;
        }

        /* Loading Overlay */
        .loading-overlay {
          position: absolute;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          background: rgba(255, 255, 255, 0.9);
          display: flex;
          justify-content: center;
          align-items: center;
          z-index: 100;
          border-radius: 16px;
        }

        .loading-content {
          text-align: center;
          background: white;
          padding: 30px;
          border-radius: 12px;
          box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
        }

        /* Filters Section */
        .filters-section {
          background: white;
          border-radius: 16px;
          padding: 24px;
          margin-bottom: 24px;
          box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
          border: 1px solid #e2e8f0;
          position: relative;
        }

        .filters-header {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-bottom: 24px;
        }

        .filters-header h3 {
          font-size: 18px;
          color: #1e293b;
          margin: 0;
        }

        .filters-info {
          display: flex;
          align-items: center;
          gap: 16px;
        }

        .active-filters {
          font-size: 14px;
          color: #64748b;
          background: #f1f5f9;
          padding: 4px 12px;
          border-radius: 20px;
        }

        .btn-clear {
          padding: 6px 16px;
          background: #f1f5f9;
          color: #64748b;
          border: 1px solid #e2e8f0;
          border-radius: 6px;
          font-size: 14px;
          font-weight: 500;
          cursor: pointer;
          transition: all 0.2s;
        }

        .btn-clear:hover:not(:disabled) {
          background: #e2e8f0;
          color: #475569;
        }

        .btn-clear:disabled {
          opacity: 0.5;
          cursor: not-allowed;
        }

        .filters-grid {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
          gap: 24px;
        }

        .filter-group {
          display: flex;
          flex-direction: column;
          gap: 12px;
        }

        .group-label {
          font-size: 14px;
          font-weight: 600;
          color: #475569;
          text-transform: uppercase;
          letter-spacing: 0.05em;
        }

        .group-filters {
          display: flex;
          flex-direction: column;
          gap: 12px;
        }

        .filter-item {
          display: flex;
          flex-direction: column;
          gap: 4px;
        }

        .filter-item label {
          font-size: 12px;
          font-weight: 500;
          color: #64748b;
        }

        .filter-hint {
          font-size: 11px;
          color: #94a3b8;
          font-style: italic;
        }

        .filter-input,
        .filter-select {
          padding: 8px 12px;
          border: 1px solid #e2e8f0;
          border-radius: 6px;
          font-size: 14px;
          background: white;
          transition: all 0.2s;
        }

        .filter-input:focus,
        .filter-select:focus {
          outline: none;
          border-color: #3b82f6;
          box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.1);
        }

        /* Table Section */
        .table-section {
          background: white;
          border-radius: 16px;
          overflow: hidden;
          box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
          border: 1px solid #e2e8f0;
        }

        .table-header {
          padding: 20px 24px;
          background: #004386ff;
          border-bottom: 1px solid #e2e8f0;
          display: flex;
          justify-content: space-between;
          align-items: center;
        }

        .table-info {
          font-size: 14px;
          color: #ffffffff;
        }

        .table-info strong {
          color: #ffffffff;
        }

        .table-legend {
          display: flex;
          gap: 16px;
          flex-wrap: wrap;
        }

        .legend-item {
          display: flex;
          align-items: center;
          gap: 6px;
          font-size: 11px;
          color: #ffffffff;
        }

        .legend-color {
          width: 10px;
          height: 10px;
          border-radius: 2px;
        }

        .legend-color.green {
          background: #10b981;
        }

        .legend-color.yellow {
          background: #f59e0b;
        }

        .legend-color.red {
          background: #ef4444;
        }

        .legend-color.orange {
          background: #ffa500;
        }

        .legend-color.completed {
          background: #10b981;
        }

        .legend-color.pending {
          background: #f59e0b;
        }

        .table-container {
          overflow-x: auto;
        }

        .data-table {
          width: 100%;
          border-collapse: collapse;
          min-width: 1400px;
        }

        .data-table thead {
          background: #004c97ff;
        }

        .data-table th {
          padding: 12px;
          text-align: left;
          font-size: 11px;
          font-weight: 600;
          color: #ffffffff;
          text-transform: uppercase;
          letter-spacing: 0.05em;
          border: 2px solid #e2e8f0;
        }

        .data-table th:first-child {
          padding-left: 24px;
        }

        .data-table th:last-child {
          padding-right: 24px;
        }

        .table-row {
          border: 1px solid #e2e8f0;
          transition: background-color 0.2s;
        }
          .lot-number.repeated-lot {
  color: #d4af37 !important;
  background: #fffdf0;
  border: 1px solid #ffd700;
  position: relative;
}

.lot-number.repeated-lot::before {
  content: "★ ";
  font-size: 12px;
  vertical-align: middle;
  margin-right: 2px;
}

        .table-row:hover {
          background: #f8fafc;
        }

        .color-pending-row {
          background-color: #fff3e0 !important;
        }

        .color-pending-row:hover {
          background-color: #ffecb3 !important;
        }

        .data-table td {
          padding: 12px;
          font-size: 13px;
          color: #000000ff;
          border: 1px solid #e2e8f0;
        }

        .data-table td:first-child {
          padding-left: 24px;
        }

        .data-table td:last-child {
          padding-right: 24px;
        }

        .text-center {
          text-align: center;
        }

        .font-semibold {
          font-weight: 600;
        }

        .lot-number {
          font-weight: 600;
          color: #3b82f6;
          background: #eff6ff;
          padding: 4px 8px;
          border-radius: 4px;
          border: 1px solid #dbeafe;
        }

        .lot-number.color-pending {
          color: #ff0000;
          background: #ffebee;
          border: 1px solid #ffcdd2;
          font-weight: 800;
        }

        .stitching-days {
          display: inline-block;
          padding: 4px 8px;
          border-radius: 12px;
          font-weight: 600;
          font-size: 11px;
          min-width: 30px;
          text-align: center;
        }

        .emb-date {
          font-weight: 500;
          font-size: 12px;
        }

        .emb-date.has-date {
          color: #3b82f6;
        }

        .emb-date.no-date {
          color: #94a3b8;
          font-style: italic;
        }

        .wip-status {
          font-weight: 500;
          color: #475569;
          font-size: 12px;
        }

        .total-pcs {
          font-weight: 600;
          color: #059669;
          background: #d1fae5;
          padding: 4px 8px;
          border-radius: 12px;
          font-size: 12px;
        }

        .completion-date {
          font-size: 12px;
          font-weight: 500;
          padding: 4px 8px;
          border-radius: 12px;
        }

        .completion-date.completed {
          color: #059669;
          background: #d1fae5;
        }

        .completion-date.pending {
          color: #dc2626;
          background: #fee2e2;
        }

        .status-badge {
          font-size: 11px;
          font-weight: 600;
          padding: 4px 10px;
          border-radius: 20px;
          display: inline-block;
        }

        .status-badge.completed-badge {
          background: #10b981;
          color: white;
        }

        .status-badge.pending-badge {
          background: #f59e0b;
          color: white;
        }

        .status-badge.color-pending-badge {
          background: #ffa500;
          color: white;
          border: 1px solid #ff8c00;
        }

        .table-footer {
          padding: 16px 24px;
          background: #f8fafc;
          border-top: 1px solid #e2e8f0;
          display: flex;
          justify-content: space-between;
          align-items: center;
        }

        .footer-info {
          font-size: 14px;
          color: #64748b;
        }

        .footer-info strong {
          color: #1e293b;
        }

        .footer-stats {
          font-size: 13px;
          color: #6b7280;
          margin-left: 12px;
        }

        .footer-actions {
          display: flex;
          gap: 8px;
        }

        /* Empty State */
        .empty-state {
          background: white;
          border-radius: 16px;
          padding: 60px 20px;
          text-align: center;
          box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        }

        .empty-icon {
          font-size: 48px;
          margin-bottom: 20px;
          opacity: 0.5;
        }

        .empty-state h3 {
          color: #1e293b;
          margin-bottom: 12px;
          font-size: 20px;
        }

        .empty-state p {
          color: #64748b;
          max-width: 400px;
          margin: 0 auto 24px;
        }

        .btn-primary {
          background: #3b82f6;
          color: white;
          padding: 10px 24px;
          border: none;
          border-radius: 8px;
          font-weight: 600;
          cursor: pointer;
          transition: all 0.2s;
        }

        .btn-primary:hover:not(:disabled) {
          background: #2563eb;
        }

        /* Responsive Design */
        @media (max-width: 1200px) {
          .header-top {
            flex-direction: column;
            gap: 24px;
          }
          
          .stats-card {
            width: 100%;
            justify-content: space-around;
          }
        }

        @media (max-width: 768px) {
          .container {
            padding: 12px;
          }
          
          .header,
          .filters-section,
          .table-section {
            padding: 16px;
          }
          
          .title-section h1 {
            font-size: 24px;
          }
          
          .header-actions {
            flex-direction: column;
            gap: 12px;
          }
          
          .action-group {
            width: 100%;
            justify-content: center;
          }
          
          .table-header {
            flex-direction: column;
            gap: 16px;
            align-items: stretch;
          }
          
          .table-legend {
            justify-content: center;
          }
          
          .table-footer {
            flex-direction: column;
            gap: 16px;
            text-align: center;
          }
        }
      `}</style>
    </div>
  );
};

export default DailyStitchingUpdation;