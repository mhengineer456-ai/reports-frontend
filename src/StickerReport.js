import React, { useState, useEffect, useCallback, useRef } from 'react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

const StickerReport = () => {
  // Data states
  const [data, setData] = useState([]);
  const [filteredData, setFilteredData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [isInitialLoad, setIsInitialLoad] = useState(true);
  
  // Cache management
  const dataCache = useRef(new Map());
  const lastFetchTime = useRef(new Map());
  const CACHE_DURATION = 5 * 60 * 1000; // 5 minutes cache
  const BATCH_SIZE = 50; // Process data in batches for large datasets
  
  // Filters
 // Add this to your filters state (around line 33-34)
const [filters, setFilters] = useState({
  globalSearch: '',
  lotNumber: '',
  fabric: '',
  garmentType: '',
  style: '',
  brand: '',
  partyName: '',
  supervisor: '',
  season: '',
  mwk: '',
  directStitching: '',
  challanHistory: '',
  wipStatus: '',
  completedStatus: '',
  lotStatus: 'Completed',
  dateRange: { from: '', to: '' },
  stitchingDaysFilter: [],
  sticker: '' // ADDED: Sticker filter
});

  const [sortConfig, setSortConfig] = useState({
    key: 'completionDate',
    direction: 'desc' // Default: descending
  });
    const normalizeText = useCallback((text) => {
    if (!text || text.trim() === '') return '';
    return text.trim().toLowerCase();
  }, []);

  
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
  stitchingDaysFilter: ['green', 'yellow', 'red'],
  sticker: [] // ADDED: Sticker options
});

  // Google Sheets configuration
  const SPREADSHEET_ID = '1Hj3JeJEKB43aYYWv8gk2UhdU6BWuEQfCg5pBlTdBMNA';
  const API_KEY = 'AIzaSyAomDFBkOySlIxKWSKGHe6ATv9gvaBr7uk';
  const SHEET_NAME = 'Index';
  const CUTTING_SHEET_NAME = 'Cutting';
  const CUTTING_RANGE = `${CUTTING_SHEET_NAME}!A1:ZZ200000`;
  
  const normalizeKey = (s = "") => {
    return String(s || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
  };

// Remove the const HEADERS array and replace with this function
// Replace the HEADERS function with this version
// Replace the HEADERS function with this version
const HEADERS = () => {
  return [
    'Sr.No',
    'Lot Number',
    'Fabric',
    'Garment Type',
    'Style',
    'BRAND',
    'PARTY NAME',
    'Supervisor',
    'Season',
    'M/W/K',
    'Direct Stitching',
    'Date of Issue',
    'Stitching Days',
    'Emb/Print Date',
    'WIP Status',
    'Total PCS',
    'Completion Date',
    'Status',
    'Sticker' // 19th column
  ];
};
// Helper function to parse Date of Issue properly
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
  // Helper functions - memoized

const normalizeAndCapitalize = useCallback((text) => {
  if (!text || text.trim() === '') return '';
  
  const trimmed = text.trim().toLowerCase();
  
  // Special handling for yes/no values
  if (trimmed === 'yes' || trimmed === 'y') return 'Yes';
  if (trimmed === 'no' || trimmed === 'n') return 'No';
  if (trimmed === 'na' || trimmed === 'n/a') return 'N/A';
  
  // For other text, capitalize first letter of each word
  return trimmed.split(' ')
    .map(word => word.charAt(0).toUpperCase() + word.slice(1))
    .join(' ');
}, []);

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
  // Add this function near other helper functions (around line 200)
const abbreviateMWK = useCallback((mwkValue) => {
  if (!mwkValue || typeof mwkValue !== 'string') return 'N/A';
  
  const value = mwkValue.trim().toLowerCase();
  
  if (value.includes('gents') || value === 'm' || value === 'mens') return 'M';
  if (value.includes('kids') || value === 'k') return 'K';
  if (value.includes('girls') || value === 'g' || value.includes('girlish') || value.includes('girls')) return 'G';
  if (value.includes('women') || value.includes('womens') || value === 'w' || value.includes('women')) return 'W';
  
  // Return first letter if no match
  return value.charAt(0).toUpperCase();
}, []);

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
  // Add this function near other date functions (around line 200)
const formatDateToDDMMYY = useCallback((dateString) => {
  if (!dateString || dateString.trim() === '') return 'N/A';
  
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
        
        // Try month/day/year (US format)
        const usDate = new Date(year, day - 1, month);
        if (!isNaN(usDate.getTime())) {
          return `${String(month).padStart(2, '0')}/${String(day).padStart(2, '0')}/${String(fullYear).slice(-2)}`;
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
}, []);

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

const getLatestWipRemarks = useCallback((wipStatus, isCompleted = false) => {
  // If the lot is completed, always return "Done"
  if (isCompleted) {
    return 'Done';
  }
  
  if (!wipStatus || wipStatus.trim() === '') {
    return 'N/A';
  }
  
  try {
    if (typeof wipStatus === 'string' && !wipStatus.startsWith('[')) {
      return wipStatus;
    }
    
    const statusArray = JSON.parse(wipStatus);
    
    if (!Array.isArray(statusArray) || statusArray.length === 0) {
      return 'N/A';
    }
    
    const sortedStatuses = [...statusArray].sort((a, b) => {
      const dateA = new Date(a.timestamp).getTime();
      const dateB = new Date(b.timestamp).getTime();
      return dateB - dateA;
    });
    
    const latestStatus = sortedStatuses[0];
    
    if (latestStatus.remarks && latestStatus.remarks.trim() !== '') {
      return latestStatus.remarks;
    } else if (latestStatus.status && latestStatus.status.trim() !== '') {
      return latestStatus.status;
    } else {
      return 'N/A';
    }
  } catch (error) {
    console.error('Error parsing WIP Status for remarks:', error);
    return wipStatus;
  }
}, []);

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
// Update the extractFilterOptions function to properly normalize sticker values
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
    stitchingDaysFilter: ['green', 'yellow', 'red'],
    sticker: [] // Sticker options
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
        const isCompleted = isLotCompleted(item.completedStatus);
        const latestRemarks = getLatestWipRemarks(item[key], isCompleted);
        if (latestRemarks && latestRemarks.trim() !== '' && latestRemarks !== 'N/A') {
          const normalizedValue = normalizeAndCapitalize(latestRemarks);
          optionSets[key].add(normalizedValue);
        }
      } else if (key === 'completedStatus') {
        const statusText = getCompletedStatusText(item[key]);
        if (statusText && statusText.trim() !== '' && statusText !== 'N/A') {
          const normalizedValue = normalizeAndCapitalize(statusText);
          optionSets[key].add(normalizedValue);
        }
      } else if (key === 'sticker') {
        // Extract and normalize sticker values
        const stickerValue = item[key];
        if (stickerValue && stickerValue.trim() !== '' && stickerValue !== 'N/A') {
          // Normalize the sticker value - convert to consistent case
          let normalizedSticker = stickerValue.trim();
          
          // Handle yes/no variations
          const lowerSticker = normalizedSticker.toLowerCase();
          if (lowerSticker === 'yes' || lowerSticker === 'y') {
            normalizedSticker = 'Yes';
          } else if (lowerSticker === 'no' || lowerSticker === 'n') {
            normalizedSticker = 'No';
          } else {
            normalizedSticker = normalizeAndCapitalize(normalizedSticker);
          }
          
          optionSets[key].add(normalizedSticker);
        }
      } else if (key === 'lotStatus' || key === 'stitchingDaysFilter') {
        // Skip for predefined options
        return;
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
    if (key !== 'lotStatus' && key !== 'stitchingDaysFilter') {
      options[key] = Array.from(optionSets[key]).sort();
    }
  });

  setFilterOptions(options);
}, [getLatestWipRemarks, getCompletedStatusText, normalizeAndCapitalize, isLotCompleted]);

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

  const sortDataByCompletionDate = useCallback((data, direction = 'desc') => {
    return [...data].sort((a, b) => {
      const dateA = getCompletionDate(a.completedStatus);
      const dateB = getCompletionDate(b.completedStatus);
      
      // Handle cases where dates might be null
      if (!dateA && !dateB) return 0;
      if (!dateA) return direction === 'desc' ? 1 : -1;
      if (!dateB) return direction === 'desc' ? -1 : 1;
      
      const timeA = dateA.getTime();
      const timeB = dateB.getTime();
      
      if (direction === 'desc') {
        return timeB - timeA; // Newest first
      } else {
        return timeA - timeB; // Oldest first
      }
    });
  }, [getCompletionDate]);

  // Fetch data for specific supervisor or all data
  const fetchDataForSupervisor = useCallback(async (supervisor = '') => {
    const cacheKey = getCacheKey(supervisor);
    
    // Check cache first
    if (isCacheValid(cacheKey) && dataCache.current.has(cacheKey)) {
      return dataCache.current.get(cacheKey);
    }

    try {
      setLoading(true);
      
      // Fetch cutting data once for all lots
      const cuttingData = await fetchCuttingData();
      
      // Fetch index data
      const sheetNameEncoded = encodeURIComponent(SHEET_NAME);
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${sheetNameEncoded}?key=${API_KEY}`;

      const response = await fetch(url);
      
      if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
      }
      
      const result = await response.json();
      const rows = result.values;
      
      if (!rows || rows.length === 0) {
        return [];
      }

      const sheetHeaders = rows[0];
      
      // Create a more robust column mapping
 // Create a more robust column mapping
const createColumnMapping = () => {
  const mapping = {};
  
  // Define all possible column names and their variations
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
    { keys: ['Sticker', 'STICKER', 'sticker', 'Sticker'], target: 'sticker' } // FIXED: Added more variations
  ];
  
  // Find each column in the headers
  columnDefinitions.forEach(definition => {
    for (const key of definition.keys) {
      const index = sheetHeaders.findIndex(
        header => header && normalizeKey(header) === normalizeKey(key)
      );
      if (index !== -1) {
        mapping[definition.target] = index;
        console.log(`Mapped ${definition.target} to column ${index} (${key})`); // Debug log
        break;
      }
    }
  });
  
  return mapping;
};

      const columnMapping = createColumnMapping();
      
      // Process data in batches
      const processedData = await processDataInBatches(rows.slice(1), (row) => {
        const item = {};
        
        // Map all columns
        Object.entries(columnMapping).forEach(([targetKey, columnIndex]) => {
          if (columnIndex !== undefined && row[columnIndex] !== undefined) {
            const value = row[columnIndex];
            
            // Handle special parsing for certain fields
            switch (targetKey) {
              case 'startRow':
                const startRow = parseInt(value);
                item[targetKey] = isNaN(startRow) ? 0 : startRow;
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
        
        return item;
      });

      // Filter for supervisor if specified
      let filteredData = processedData;
      if (supervisor && supervisor.trim() !== '') {
        filteredData = processedData.filter(item => 
          item.supervisor && normalizeText(item.supervisor) === normalizeText(supervisor)
        );
      }

      // Filter out invalid rows
      const validData = filteredData.filter(item => {
        const hasLotNumber = item.lotNumber && item.lotNumber.trim() !== '';
        const hasDateOfIssue = item.dateOfIssue && item.dateOfIssue.trim() !== '';
        const hasSupervisor = item.supervisor && item.supervisor.trim() !== '';
        
        return hasLotNumber && (hasDateOfIssue || hasSupervisor);
      });

      // Update cache
      dataCache.current.set(cacheKey, validData);
      lastFetchTime.current.set(cacheKey, Date.now());

      return validData;
    } catch (err) {
      console.error('Error fetching data:', err);
      throw err;
    } finally {
      setLoading(false);
    }
  }, [getCacheKey, isCacheValid, processDataInBatches, normalizeText]);

  // Apply filters to data - UPDATED to handle dateRange object properly
 // Apply filters to data - UPDATED to handle stitchingDaysFilter array
// Update the applyFilters function (around line 450-460)
// Apply filters to data - UPDATED to handle stitchingDaysFilter array and global search
// Update the applyFilters function to handle sticker filtering case-insensitively
const applyFilters = useCallback((currentFilters, currentData) => {
  if (currentData.length === 0) {
    setFilteredData([]);
    return;
  }

  requestAnimationFrame(() => {
    let filtered = [...currentData];
    
    // Apply global search filter first if it exists
    if (currentFilters.globalSearch && currentFilters.globalSearch.trim() !== '') {
      const searchTerm = normalizeText(currentFilters.globalSearch);
      filtered = filtered.filter(item => {
        // Search in multiple fields including sticker
        const searchFields = [
          'lotNumber',
          'fabric',
          'garmentType',
          'style',
          'brand',
          'partyName',
          'supervisor',
          'season',
          'mwk',
          'directStitching',
          'sticker'
        ];
        
        // Check if any of the direct fields match
        const directMatch = searchFields.some(field => {
          let fieldValue = item[field] ? normalizeText(item[field]) : '';
          
          // Special handling for sticker to normalize yes/YES/Yes
          if (field === 'sticker' && fieldValue) {
            const lowerSticker = fieldValue.toLowerCase();
            if (lowerSticker === 'yes' || lowerSticker === 'y') {
              fieldValue = 'yes';
            } else if (lowerSticker === 'no' || lowerSticker === 'n') {
              fieldValue = 'no';
            }
          }
          
          return fieldValue.includes(searchTerm);
        });
        
        // Also search in WIP Status remarks
        const isCompleted = isLotCompleted(item.completedStatus);
        const wipRemarks = getLatestWipRemarks(item.wipStatus, isCompleted);
        const wipMatch = wipRemarks && wipRemarks !== 'N/A' ? 
          normalizeText(wipRemarks).includes(searchTerm) : false;
        
        // Also search in completed status
        const completedStatusText = getCompletedStatusText(item.completedStatus);
        const completedMatch = completedStatusText && completedStatusText !== 'N/A' ?
          normalizeText(completedStatusText).includes(searchTerm) : false;
        
        // Also search in Emb/Print date
        const embPrintDate = getEmbPrintDate(item.challanHistory);
        const embPrintMatch = embPrintDate && embPrintDate !== '-' ?
          normalizeText(embPrintDate).includes(searchTerm) : false;
        
        // Also search in stitching days (as text)
        const stitchingDays = calculateStitchingDays(item.dateOfIssue);
        const stitchingDaysMatch = stitchingDays.toString().includes(searchTerm);
        
        // Also search in total PCS
        const totalPCSMatch = item.totalPCS ? 
          item.totalPCS.toString().includes(searchTerm) : false;
        
        return directMatch || wipMatch || completedMatch || embPrintMatch || 
               stitchingDaysMatch || totalPCSMatch;
      });
    }
    
    // Then apply other filters
    filtered = filtered.filter(item => {
      return Object.entries(currentFilters).every(([key, value]) => {
        // Skip global search as we already applied it
        if (key === 'globalSearch') return true;
        
        if (!value || (typeof value === 'string' && value.trim() === '')) {
          return true;
        }
        
        // Handle date range filter separately
        if (key === 'dateRange') {
          const { from, to } = value;
          if (!from && !to) return true;
          
          const completionDate = getCompletionDate(item.completedStatus);
          if (!completionDate) return false;
          
          const itemDate = completionDate.getTime();
          const fromDate = from ? new Date(from).getTime() : 0;
          const toDate = to ? new Date(to).getTime() : Infinity;
          
          return itemDate >= fromDate && itemDate <= toDate;
        }
        
        // Handle stitchingDaysFilter array
        if (key === 'stitchingDaysFilter') {
          const selectedFilters = value;
          if (!selectedFilters || selectedFilters.length === 0) return true;
          
          const stitchingDays = calculateStitchingDays(item.dateOfIssue);
          
          // Check if stitching days fall into any selected category
          return selectedFilters.some(filter => {
            if (filter === 'green') {
              return stitchingDays <= 6;
            } else if (filter === 'yellow') {
              return stitchingDays >= 7 && stitchingDays <= 15;
            } else if (filter === 'red') {
              return stitchingDays > 15;
            }
            return false;
          });
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
        
        // Handle wipStatus filter
        if (key === 'wipStatus') {
          const isCompleted = isLotCompleted(item.completedStatus);
          const latestRemarks = getLatestWipRemarks(item[key], isCompleted);
          return normalizeText(latestRemarks).includes(filterValue);
        }
        
        // Handle sticker filter with case normalization
        if (key === 'sticker') {
          let itemValue = item[key] ? normalizeText(item[key]) : '';
          
          // Normalize both values for comparison
          const normalizedItemValue = itemValue.toLowerCase();
          const normalizedFilterValue = filterValue.toLowerCase();
          
          // Special handling for yes/no
          if ((normalizedItemValue === 'yes' || normalizedItemValue === 'y') && 
              (normalizedFilterValue === 'yes' || normalizedFilterValue === 'y')) {
            return true;
          }
          if ((normalizedItemValue === 'no' || normalizedItemValue === 'n') && 
              (normalizedFilterValue === 'no' || normalizedFilterValue === 'n')) {
            return true;
          }
          
          return normalizedItemValue.includes(normalizedFilterValue);
        }
        
        const itemValue = item[key] ? normalizeText(item[key]) : '';
        return itemValue.includes(filterValue);
      });
    });

    // Sort the filtered data by completion date in descending order
    if (currentFilters.lotStatus === 'Completed') {
      filtered = sortDataByCompletionDate(filtered, 'desc');
    }
    
    setFilteredData(filtered);
  });
}, [isLotCompleted, getCompletionDate, getCompletedStatusText, getEmbPrintDate, getLatestWipRemarks, normalizeText, sortDataByCompletionDate, calculateStitchingDays]);
  // Initial load - fetch ALL data (including all supervisors)
  useEffect(() => {
    const loadInitialData = async () => {
      try {
        setLoading(true);
        setIsInitialLoad(true);
        
        // Fetch ALL data, not just Monu's
        const allData = await fetchDataForSupervisor('');
        setData(allData);
        
        // Show only COMPLETED lots by default (not pending)
        const completedData = allData.filter(item => isLotCompleted(item.completedStatus));
        
        // Sort completed data in descending order by completion date
        const sortedCompletedData = sortDataByCompletionDate(completedData, 'desc');
        setFilteredData(sortedCompletedData);
        
        // Extract filter options from ALL data
        extractFilterOptions(allData);
        
      } catch (err) {
        setError(`Failed to load initial data: ${err.message}`);
      } finally {
        setLoading(false);
        setIsInitialLoad(false);
      }
    };

    loadInitialData();
  }, [fetchDataForSupervisor, extractFilterOptions, isLotCompleted, sortDataByCompletionDate]);

const handleFilterChange = useCallback(async (e) => {
  const { name, value, type, checked } = e.target;
   if (name === 'globalSearch') {
    const newFilters = { ...filters, globalSearch: value };
    setFilters(newFilters);
    
    // Apply filters immediately
    setTimeout(() => {
      applyFilters(newFilters, data);
    }, 0);
    return;
  }
  
  // Handle date range inputs separately
  if (name === 'dateFrom' || name === 'dateTo') {
    const newDateRange = {
      ...filters.dateRange,
      [name === 'dateFrom' ? 'from' : 'to']: value
    };
    
    const newFilters = { ...filters, dateRange: newDateRange };
    setFilters(newFilters);
    
    // Apply filters immediately
    setTimeout(() => {
      applyFilters(newFilters, data);
    }, 0);
  } 
  // Handle stitching days checkboxes
  else if (name === 'stitchingDaysFilter') {
    const currentFilters = [...filters.stitchingDaysFilter];
    let newFilters;
    
    if (checked) {
      // Add the value if checked
      if (!currentFilters.includes(value)) {
        newFilters = { ...filters, stitchingDaysFilter: [...currentFilters, value] };
      } else {
        newFilters = filters; // No change
      }
    } else {
      // Remove the value if unchecked
      newFilters = { 
        ...filters, 
        stitchingDaysFilter: currentFilters.filter(item => item !== value) 
      };
    }
    
    setFilters(newFilters);
    
    // Apply filters immediately
    setTimeout(() => {
      applyFilters(newFilters, data);
    }, 0);
  } else {
    const newFilters = { ...filters, [name]: value };
    setFilters(newFilters);
    
    // If supervisor is changing, fetch new data
    if (name === 'supervisor') {
      try {
        setLoading(true);
        
        // Clear cache for previous supervisor if needed
        if (filters.supervisor !== value) {
          const oldCacheKey = getCacheKey(filters.supervisor);
          if (dataCache.current.has(oldCacheKey)) {
            // Keep old cache for performance, just fetch new
          }
        }
        
        // Fetch data for the new supervisor
        const newData = await fetchDataForSupervisor(value);
        setData(newData);
        
        // Apply the new filters to the new data
        applyFilters(newFilters, newData);
        
        // Update filter options from new data
        extractFilterOptions(newData);
        
      } catch (err) {
        setError(`Failed to load data for ${value}: ${err.message}`);
      } finally {
        setLoading(false);
      }
    } else {
      // For other filters, just apply to existing data
      setTimeout(() => {
        applyFilters(newFilters, data);
      }, 0);
    }
  }
}, [filters, data, applyFilters, fetchDataForSupervisor, extractFilterOptions, getCacheKey]);

  // Add a separate function to apply filters when date range changes
  const applyDateRangeFilter = useCallback(() => {
    applyFilters(filters, data);
  }, [filters, data, applyFilters]);

  // Call applyDateRangeFilter when date range changes
  useEffect(() => {
    if (filters.dateRange.from || filters.dateRange.to) {
      applyDateRangeFilter();
    }
  }, [filters.dateRange, applyDateRangeFilter]);

  // Handle clear filters - UPDATED to reset to Completed (not Monu)
  // Handle clear filters - UPDATED to reset stitchingDaysFilter array
const clearFilters = useCallback(async () => {
  // Default filters with ALL supervisors and Completed status
  const defaultFilters = {
    lotNumber: '',
    fabric: '',
    garmentType: '',
    style: '',
    brand: '',
    partyName: '',
    supervisor: '', // Empty for all supervisors
    season: '',
    mwk: '',
    directStitching: '',
    challanHistory: '',
    wipStatus: '',
    completedStatus: '',
    lotStatus: 'Completed',  // Default to Completed for all
    dateRange: { from: '', to: '' }, // Reset date range
    stitchingDaysFilter: [] // Reset to empty array
  };
  
  setFilters(defaultFilters);
  
  // Clear cache and reload ALL data
  try {
    setLoading(true);
    
    // Clear all cache
    dataCache.current.clear();
    lastFetchTime.current.clear();
    
    // Fetch fresh ALL data
    const allData = await fetchDataForSupervisor('');
    setData(allData);
    
    // Show only completed data for all
    const completedData = allData.filter(item => isLotCompleted(item.completedStatus));
    // Sort in descending order
    const sortedCompletedData = sortDataByCompletionDate(completedData, 'desc');
    setFilteredData(sortedCompletedData);
    
    // Update filter options
    extractFilterOptions(allData);
    
  } catch (err) {
    setError(`Failed to load data: ${err.message}`);
  } finally {
    setLoading(false);
  }
}, [fetchDataForSupervisor, extractFilterOptions, isLotCompleted, sortDataByCompletionDate]);


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
    
    // Use filteredData directly
    const exportData = filteredData;
    
    if (exportData.length === 0) {
      alert('No data available to export.');
      return;
    }

    // Create worksheet data with dynamic headers based on lotStatus
    // In downloadExcel function, update worksheetData
const worksheetData = [
  HEADERS(), // Use the dynamic HEADERS function
  ...exportData.map((item, index) => {
    const rowData = [
      index + 1,
      item.lotNumber || '',
      item.fabric || '',
      item.garmentType || '',
      item.style || '',
      item.brand || '',
      item.partyName || '',
      item.supervisor || '',
      item.season || '',
      item.mwk || '',
      item.directStitching || '',
      formatDateToDDMMYY(item.dateOfIssue), // Date of Issue
      calculateStitchingDays(item.dateOfIssue),
      getEmbPrintDate(item.challanHistory),
    ];
    
    // WIP Status - Check if completed to show "Done"
    const isCompleted = isLotCompleted(item.completedStatus);
    const wipValue = isCompleted ? 'Done' : getLatestWipRemarks(item.wipStatus, false);
    rowData.push(wipValue);
    
    // Add remaining columns
    rowData.push(
      item.totalPCS || 0,
      getCompletionDateFormatted(item.completedStatus) || '',
      isLotCompleted(item.completedStatus) ? 'Completed' : 'Pending',
      item.sticker || '' // ADDED: Sticker value
    );
    
    return rowData;
  })
];

    const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Stitching Issue Report');
    
    // Add active filters to sheet name
    const activeFilters = Object.entries(filters)
      .filter(([key, value]) => {
        if (key === 'dateRange') {
          const { from, to } = value;
          return from || to;
        }
        return typeof value === 'string' && value.trim() !== '';
      })
      .map(([key, value]) => {
        if (key === 'dateRange') {
          const { from, to } = value;
          return `Date Range: ${from || ''} to ${to || ''}`;
        }
        return `${key}: ${value}`;
      });
    
    let sheetName = 'Stitching Report';
    if (activeFilters.length > 0) {
      sheetName = `Stitching (Filtered)`;
    }
    
    // Add a filter summary row
    if (activeFilters.length > 0) {
      const filterRow = ['Filter Summary:', ...activeFilters];
      const filterData = [['']]; // Empty row
      filterData.push(filterRow);
      
      const filterSheet = XLSX.utils.aoa_to_sheet(filterData);
      XLSX.utils.book_append_sheet(workbook, filterSheet, 'Filters');
    }

    // Calculate column widths based on dynamic headers
    const currentHeaders = HEADERS();
    const maxWidths = currentHeaders.map((header, colIndex) => {
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
            item.supervisor || '',
            item.season || '',
            item.mwk || '',
            item.directStitching || '',
            formatDateToDDMMYY(item.dateOfIssue),
            calculateStitchingDays(item.dateOfIssue),
            getEmbPrintDate(item.challanHistory),
          ];
          
          // Add WIP Status data - with "Done" for completed lots
          const isCompleted = isLotCompleted(item.completedStatus);
          const wipValue = isCompleted ? 'Done' : getLatestWipRemarks(item.wipStatus, false);
          rowData.push(wipValue);
          
          // Add remaining columns
          rowData.push(
            item.totalPCS || 0,
            getCompletionDateFormatted(item.completedStatus) || '',
            isLotCompleted(item.completedStatus) ? 'Completed' : 'Pending'
          );
          
          return rowData[colIndex] ? rowData[colIndex].toString().length : 0;
        })
      );
      return { wch: Math.min(maxLength + 2, 50) };
    });
    
    worksheet['!cols'] = maxWidths;

    // Generate filename with filter info
    const today = new Date();
    let fileName = `Stitching_Report_${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
    
    // Add filter info to filename if any filter is active
    if (activeFilters.length > 0) {
      const filterString = activeFilters
        .map(f => f.split(':')[1].trim())
        .join('_')
        .substring(0, 50); // Limit length
      fileName = `${fileName}_Filtered_${filterString}`;
    }
    
    XLSX.writeFile(workbook, `${fileName}.xlsx`);
  } catch (err) {
    setError(`Failed to export Excel: ${err.message}`);
  } finally {
    setLoading(false);
  }
}, [filteredData, filters, calculateStitchingDays, getEmbPrintDate, getLatestWipRemarks, getCompletionDateFormatted, isLotCompleted, formatDateToDDMMYY]);
const downloadPDF = useCallback(async () => {
  try {
    setLoading(true);
    
    // Use filteredData directly
    const exportData = filteredData;
    
    if (exportData.length === 0) {
      alert('No data available to export.');
      return;
    }

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

    // Function to format Date of Issue as 29/11/25
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
    const highlightColor = [16, 185, 129]; // Green for highlighting
    const remarksColor = [239, 68, 68]; // Red for remarks and lot number
    const pcsColor = [239, 68, 68]; // RED for PCS
    const lotNumberColor = [239, 68, 68]; // RED for lot number
    const partyColor = [107, 33, 168]; // Purple for Party abbreviation
    const completedColor = [16, 185, 129]; // Green for completed status
    const pendingColor = [245, 158, 11]; // Yellow for pending status
    const issueDateColor = [139, 92, 246]; // Purple for Date of Issue

    // Page dimensions
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const margin = 8; // Reduced margin for more width
    const contentWidth = pageWidth - (margin * 2);

    // Variable to track current Y position
    let currentY = 25;

    // Function to draw header with white background and blue text
    const drawHeader = () => {
      // WHITE background for entire header area
      doc.setFillColor(255, 255, 255);
      doc.rect(0, 0, pageWidth, 30, 'F');
      
      // Calculate total PCS and other metrics
      const totalPCS = exportData.reduce((sum, item) => sum + (item.totalPCS || 0), 0);
      const totalLots = exportData.length;
      const totalCompleted = exportData.filter(item => isLotCompleted(item.completedStatus)).length;
      const totalPending = totalLots - totalCompleted;
      
      // Dynamic title based on active filters
      let title = 'STITCHING PRODUCTION REPORT';
      
      // Add supervisor to title if filtered
      let supervisorName = '';
      if (filters.supervisor && filters.supervisor.trim() !== '') {
        supervisorName = filters.supervisor.toUpperCase();
        title = `${supervisorName} - STITCHING PRODUCTION REPORT`;
      }
      
      // Add status to title if filtered
      if (filters.lotStatus) {
        if (filters.lotStatus === 'Completed') {
          if (supervisorName) {
            title = `${supervisorName} - COMPLETED LOTS REPORT`;
          } else {
            title = 'COMPLETED LOTS REPORT';
          }
        } else if (filters.lotStatus === 'Pending') {
          if (supervisorName) {
            title = `${supervisorName} - PENDING LOTS REPORT`;
          } else {
            title = 'PENDING LOTS REPORT';
          }
        }
      }
      
      // Main Title - Blue text on white background
      doc.setFontSize(18);
      doc.setTextColor(15, 76, 129); // Dark blue text
      doc.setFont('Times New Roman', 'bold');
      doc.text(title, pageWidth / 2, 12, { align: 'center' });
      
      // Add supervisor badge if supervisor filter is applied
      if (supervisorName) {
        doc.setFontSize(10);
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(15, 76, 129); // Dark blue for supervisor badge
        doc.text(`SUPERVISOR: ${supervisorName}`, pageWidth / 2, 18, { align: 'center' });
      }
      
      // Key Metrics Row - Styled in blue
      doc.setFontSize(11);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(15, 76, 129); // Dark blue text
      
      // Left side: Date
      const today = new Date();
      const reportDate = `${String(today.getDate()).padStart(2, '0')}/${String(today.getMonth() + 1).padStart(2, '0')}/${String(today.getFullYear()).slice(-2)}`;
      doc.text(`Report Date: ${reportDate}`, margin, 25);
      
      // Center: Totals summary WITH TOTAL PCS
      const centerX = pageWidth / 2;
      doc.text(`Lots: ${totalLots} | Total PCS: ${totalPCS.toLocaleString()} | Completed: ${totalCompleted} | Pending: ${totalPending}`, 
                centerX, 25, { align: 'center' });
      
      // Right side: Filter status
      const activeFilterCount = Object.entries(filters)
        .filter(([key, value]) => {
          if (key === 'dateRange') {
            const { from, to } = value;
            return from || to;
          }
          return typeof value === 'string' && value.trim() !== '';
        })
        .length;
      
      const filterText = activeFilterCount > 0 ? `Filtered: ${exportData.length} records` : `Total: ${exportData.length} records`;
      doc.text(filterText, pageWidth - margin, 25, { align: 'right' });
    };

    // Draw header on first page
    drawHeader();

    // Add stitching days color coding legend
    const addColorLegend = (yPos) => {
      doc.setFontSize(8);
      doc.setFont('helvetica', 'bold');
      
      // Legend title
      doc.setTextColor(59, 130, 246);
      doc.text('STITCHING DAYS COLOR CODING:', margin, yPos);
      
      // Color boxes with descriptions
      const legendItems = [
        { color: [16, 185, 129], text: '1-6 Days: Good (Green)' },
        { color: [245, 158, 11], text: '7-15 Days: Average (Yellow)' },
        { color: [239, 68, 68], text: '15+ Days: Critical (Red)' }
      ];
      
      let legendX = margin + 70;
      legendItems.forEach((item, index) => {
        // Draw color box
        doc.setFillColor(...item.color);
        doc.rect(legendX, yPos - 3, 5, 5, 'F');
        
        // Draw text
        doc.setTextColor(50, 50, 50);
        doc.setFont('helvetica', 'normal');
        doc.text(item.text, legendX + 7, yPos);
        
        legendX += 70;
      });
      
      return yPos + 8;
    };

    // Add color legend after header
    currentY = addColorLegend(32);

    // ===================== MAIN DATA TABLE =====================
    
    // Dynamic headers - ALWAYS include WIP Status
// Dynamic headers - ALWAYS include WIP Status
const PDF_HEADERS = [
  'Sr',
  'Lot No',
  'Fabric',
  'Garment',
  'Style',
  'Brand',
  'Party',
  'Supervisor',
  'Season',
  'M/W/K',
  'Direct',
  'Issue Date',
  'Days',
  'Emb/Print',
  'WIP Status',
  'Total PCS',
  'Complete Date',
  'Lot Status',
  'Sticker'  // ADDED: Sticker column
];
    // Optimized column widths for better fit
   // Optimized column widths for better fit
// Optimized column widths for better fit
const columnWidths = {
  0: 9,    // Sr
  1: 18,   // Lot No
  2: 29,   // Fabric
  3: 30,   // Garment
  4: 32,   // Style
  5: 32,   // Brand
  6: 16,   // Party
  7: 30,   // Supervisor
  8: 9,    // Season
  9: 12,   // M/W/K
  10: 12,  // Direct
  11: 20,  // Issue Date
  12: 15,  // Days
  13: 20,  // Emb/Print
  14: 30,  // WIP Status - REDUCED from 50 to 30
  15: 25,  // Total PCS
  16: 25,  // Complete Date
  17: 20,  // Lot Status
  18: 20   // Sticker
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

    // Prepare table body with color coding
    // Prepare table body with color coding
// Prepare table body with color coding
// Prepare table body with color coding
const body = exportData.map((item, rowIndex) => {
  const stitchingDays = calculateStitchingDays(item.dateOfIssue);
  const stitchingDaysColor = getStitchingDaysColor(stitchingDays);
  const stitchingDaysTextColor = getStitchingDaysTextColor(stitchingDays);
  const wipRemarks = getLatestWipRemarks(item.wipStatus);
  const embPrintDate = getEmbPrintDate(item.challanHistory);
  const abbreviatedParty = abbreviatePartyName(item.partyName);
  const abbreviatedSeason = abbreviateSeason(item.season);
  const totalPCS = item.totalPCS || 0;
  const completionDate = getCompletionDateFormatted(item.completedStatus);
  const isCompleted = isLotCompleted(item.completedStatus);
  const wipDisplayValue = isCompleted ? 'Done' : wipRemarks;
  const lotStatus = isCompleted ? 'Completed' : 'Pending';
  const issueDate = formatDateOfIssue(item.dateOfIssue);
  
  const formattedEmbPrintDate = embPrintDate !== '-' ? formatDateToDDMMYYForPDF(embPrintDate) : '-';
  const formattedCompletionDate = completionDate ? formatDateToDDMMYYForPDF(completionDate) : '-';
  
  // Determine row background
  const rowBgColor = rowIndex % 2 === 0 ? [255, 255, 255] : [250, 250, 250];
  
  // Convert RGB color to array
  const stitchingDaysRGB = stitchingDaysColor.startsWith('#') ? [
    parseInt(stitchingDaysColor.slice(1, 3), 16),
    parseInt(stitchingDaysColor.slice(3, 5), 16),
    parseInt(stitchingDaysColor.slice(5, 7), 16)
  ] : [16, 185, 129];
  
  // Stitching days text color
  const stitchingDaysTextRGB = stitchingDaysTextColor === '#ffffff' ? [255, 255, 255] : [0, 0, 0];
  
  // Lot status color
  const lotStatusColor = isCompleted ? completedColor : pendingColor;
  const lotStatusTextColor = isCompleted ? [255, 255, 255] : [0, 0, 0];
  
  // Enhanced truncateText function
  const truncateText = (text, maxLength, columnType = 'general') => {
    if (!text || text === 'N/A' || text === '-') return text;
    
    const cleanText = text.toString().trim();
    
    // For very short text, return as is
    if (cleanText.length <= maxLength) return cleanText;
    
    // Special handling for different column types
    switch(columnType) {
      case 'lotNumber':
        return cleanText.length > maxLength ? cleanText.substring(0, maxLength - 3) + '...' : cleanText;
        
      case 'fabric':
      case 'style':
        // For fabric and style, try to keep important parts
        const words = cleanText.split(' ');
        if (words.length > 1) {
          // Try first word + initial of second
          const shortened = words[0] + ' ' + words[1].charAt(0);
          if (shortened.length <= maxLength) return shortened + '.';
        }
        return cleanText.substring(0, maxLength - 3) + '...';
        
      case 'party':
        return cleanText.substring(0, maxLength);
        
      case 'wip':
        // For WIP status, prioritize first part
        if (cleanText.includes('-')) {
          const parts = cleanText.split('-');
          if (parts[0].length <= maxLength) return parts[0];
        }
        return cleanText.substring(0, maxLength - 3) + '...';
        
      default:
        return cleanText.substring(0, maxLength - 3) + '...';
    }
  };
  
  // Build row cells
  const rowCells = [
    // Sr.No
    {
      content: (rowIndex + 1).toString(),
      styles: {
        cellWidth: columnWidths[0],
        fontSize: 11,
        halign: 'center',
        fontStyle: 'bold',
        fillColor: rowBgColor,
        cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
      }
    },
    
    // Lot No
    {
      content: item.lotNumber || 'N/A',
      styles: {
        cellWidth: columnWidths[1],
        fontSize: 12,
        halign: 'center',
        fontStyle: 'bold',
        textColor: [239, 68, 68],
        fillColor: rowBgColor,
        cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
      }
    },

    // Fabric
    {
      content: item.fabric || 'N/A',
      styles: {
        cellWidth: columnWidths[2],
        fontSize: 11,
        halign: 'center',
        fillColor: rowBgColor,
        cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
      }
    },

    // Garment
    {
      content: item.garmentType || 'N/A',
      styles: {
        cellWidth: columnWidths[3],
        fontSize: 11,
        halign: 'center',
        fillColor: rowBgColor,
        cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
      }
    },

    // Style
    {
      content: item.style || 'N/A',
      styles: {
        cellWidth: columnWidths[4],
        fontSize: 11,
        halign: 'center',
        fillColor: rowBgColor,
        cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
      }
    },

    // Brand
    {
      content: item.brand || 'N/A',
      styles: {
        cellWidth: columnWidths[5],
        fontSize: 11,
        halign: 'center',
        fillColor: rowBgColor,
        cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
      }
    },
    
    // Party
    {
      content: abbreviatedParty,
      styles: {
        cellWidth: columnWidths[6],
        fontSize: 10,
        halign: 'center',
        fontStyle: 'bold',
        fillColor: rowBgColor,
        textColor: partyColor,
        cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
      }
    },
    
    // Supervisor
    {
      content: truncateText(item.supervisor || 'N/A', 15),
      styles: {
        cellWidth: columnWidths[7],
        fontSize: 12,
        halign: 'center',
        fontStyle: 'bold',
        fillColor: rowBgColor,
        textColor: [59, 130, 246],
        cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
      }
    },
    
    // Season
    {
      content: abbreviatedSeason,
      styles: {
        cellWidth: columnWidths[8],
        fontSize: 12,
        halign: 'center',
        fontStyle: 'bold',
        fillColor: rowBgColor,
        textColor: abbreviatedSeason === 'S' ? [239, 68, 68] :
                   abbreviatedSeason === 'W' ? [59, 130, 246] :
                   abbreviatedSeason === 'A' ? [245, 158, 11] :
                   abbreviatedSeason === 'SP' ? [16, 185, 129] : [100, 100, 100],
        cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
      }
    },
    
    // M/W/K - abbreviated
    {
      content: abbreviateMWKForPDF(item.mwk),
      styles: {
        cellWidth: columnWidths[9],
        fontSize: 12,
        halign: 'center',
        fontStyle: 'bold',
        fillColor: rowBgColor,
        textColor: abbreviateMWKForPDF(item.mwk) === 'M' ? [59, 130, 246] :  // Blue for MENS
                   abbreviateMWKForPDF(item.mwk) === 'W' ? [239, 68, 68] :   // Red for WOMENS
                   abbreviateMWKForPDF(item.mwk) === 'K' ? [16, 185, 129] :  // Green for KIDS
                   abbreviateMWKForPDF(item.mwk) === 'G' ? [168, 85, 247] :  // Purple for GIRLS
                   [100, 100, 100],                                           // Gray for others
        cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
      }
    },
    
    // Direct
    {
      content: truncateText(item.directStitching || 'N/A', 8),
      styles: {
        cellWidth: columnWidths[10],
        fontSize: 11,
        halign: 'center',
        fillColor: rowBgColor,
        cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
      }
    },
    
    // Issue Date
    {
      content: truncateText(issueDate, 12),
      styles: {
        cellWidth: columnWidths[11],
        fontSize: 11,
        halign: 'center',
        fillColor: rowBgColor,
        fontStyle: 'bold',
        textColor: issueDateColor,
        cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
      }
    },
    
    // Days
    {
      content: stitchingDays.toString(),
      styles: {
        cellWidth: columnWidths[12],
        fontSize: 11,
        halign: 'center',
        fontStyle: 'bold',
        fillColor: stitchingDaysRGB,
        textColor: stitchingDaysTextRGB,
        cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
      }
    },
    
    // Emb/Print - formatted date
    {
      content: truncateText(formattedEmbPrintDate, 14),
      styles: {
        cellWidth: columnWidths[13],
        fontSize: 11,
        halign: 'center',
        fillColor: rowBgColor,
        fontStyle: formattedEmbPrintDate !== '-' ? 'bold' : 'normal',
        textColor: formattedEmbPrintDate !== '-' ? accentColor : [100, 100, 100],
        cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
      }
    },
    
    // WIP Status
    {
      content: wipDisplayValue || 'N/A',
      styles: {
        cellWidth: columnWidths[14],
        fontSize: 10,
        halign: 'center',
        fillColor: rowBgColor,
        fontStyle: isCompleted ? 'bold' : (wipRemarks !== 'N/A' ? 'bold' : 'normal'),
        textColor: isCompleted ? [16, 185, 129] : (wipRemarks !== 'N/A' ? remarksColor : [100, 100, 100]),
        cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
      }
    },

    // Total PCS
    {
      content: totalPCS > 0 ? totalPCS.toLocaleString() : 'N/A',
      styles: {
        cellWidth: columnWidths[15],
        fontSize: 11,
        halign: 'center',
        fontStyle: 'bold',
        fillColor: rowBgColor,
        textColor: pcsColor,
        cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
      }
    },
    
    // Complete Date - formatted date
    {
      content: truncateText(formattedCompletionDate, 14),
      styles: {
        cellWidth: columnWidths[16],
        fontSize: 11,
        halign: 'center',
        fillColor: rowBgColor,
        fontStyle: formattedCompletionDate !== '-' ? 'bold' : 'normal',
        textColor: formattedCompletionDate !== '-' ? [59, 130, 246] : [100, 100, 100],
        cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
      }
    },
    
    // Lot Status
    {
      content: truncateText(lotStatus, 10),
      styles: {
        cellWidth: columnWidths[17],
        fontSize: 11,
        halign: 'center',
        fontStyle: 'bold',
        fillColor: lotStatusColor,
        textColor: lotStatusTextColor,
        cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
      }
    },
    
    // ADDED: Sticker column
    {
      content: truncateText(item.sticker || 'N/A', 15, 'sticker'),
      styles: {
        cellWidth: columnWidths[18],
        fontSize: 11,
        halign: 'center',
        fillColor: rowBgColor,
        fontStyle: item.sticker ? 'bold' : 'normal',
        textColor: item.sticker ? [139, 92, 246] : [156, 163, 175], // Purple for sticker, gray for N/A
        cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
      }
    }
  ];
  
  return rowCells;
});
    // Create column styles based on dynamic structure
const columnStyles = {};
PDF_HEADERS.forEach((_, index) => {
  columnStyles[index] = {
    cellWidth: columnWidths[index],
    halign: 'center',
    valign: 'middle'
  };
});

    // Track the Y position manually
    let lastAutoTableY = currentY;

    // Create the table with precise styling
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
        font: 'helvetica',
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
        font: 'helvetica',
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

    // ===================== SUPERVISOR WORKLOAD DISTRIBUTION =====================
    
    // Show supervisor summary for ALL reports
    let summaryStartY = lastAutoTableY + 10;

    // Check if we need a new page
    if (summaryStartY > pageHeight - 100) {
      doc.addPage();
      summaryStartY = 20;
      drawHeader();
      currentY = addColorLegend(32);
    }

    // Simple Summary Section Header
    doc.setFillColor(255, 255, 255);
    doc.rect(margin - 2, summaryStartY - 8, contentWidth + 4, 16, 'F');

    // Section Title
    doc.setFontSize(14);
    doc.setFont('times', 'bold');
    doc.setTextColor(15, 76, 129);
    doc.text('SUPERVISOR WORKLOAD DISTRIBUTION', pageWidth / 2, summaryStartY, { align: 'center' });

    let y = summaryStartY + 15;

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
          redPCS: 0
        };
      }
      
      acc[supervisor].totalLots += 1;
      acc[supervisor].totalPCS += item.totalPCS || 0;
      
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

    // Function to normalize supervisor name for matching
    const normalizeForMatching = (name) => {
      if (!name || name.trim() === '') return '';
      return name.trim().toLowerCase();
    };

    // Function to find manpower for a supervisor
    const getManpowerForSupervisor = (supervisor) => {
      const normalizedSupervisor = normalizeForMatching(supervisor);
      
      // Try exact match first
      for (const [key, value] of Object.entries(supervisorManpowerData)) {
        if (normalizeForMatching(key) === normalizedSupervisor) {
          return value || 'N/A';
        }
      }
      
      // Try partial match
      for (const [key, value] of Object.entries(supervisorManpowerData)) {
        if (normalizeForMatching(key).includes(normalizedSupervisor) || 
            normalizedSupervisor.includes(normalizeForMatching(key))) {
          return value || 'N/A';
        }
      }
      
      return 'N/A';
    };

    // Prepare simplified summary body with MANPOWER column
    const summaryBody = sortedSupervisors.map(([supervisor, data]) => {
      const manpower = getManpowerForSupervisor(supervisor);
      
      return [
        { 
          content: supervisor,
          styles: { 
            halign: 'center', 
            fontSize: 13, 
            cellPadding: { top: 4, right: 3, bottom: 4, left: 6 },
            fontStyle: 'bold'
          } 
        },
        { 
          content: manpower,
          styles: { 
            halign: 'center', 
            fontSize: 13, 
            cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
            fontStyle: 'bold',
            fillColor: [240, 248, 255],
            textColor: [15, 76, 129]
          } 
        },
        { 
          content: data.totalLots.toString(), 
          styles: { 
            halign: 'center', 
            fontSize: 13, 
            fontStyle: 'bold',
            cellPadding: { top: 4, right: 3, bottom: 4, left: 3 }
          } 
        },
        { 
          content: data.totalPCS.toLocaleString(), 
          styles: { 
            halign: 'center', 
            fontSize: 13, 
            fontStyle: 'bold',
            cellPadding: { top: 4, right: 3, bottom: 4, left: 3 }
          } 
        },
        { 
          content: `${data.greenLots} (${data.greenPCS.toLocaleString()})`,
          styles: { 
            halign: 'center', 
            fontSize: 13, 
            cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
            fontStyle: 'bold',
            textColor: [16, 185, 129]
          } 
        },
        { 
          content: `${data.yellowLots} (${data.yellowPCS.toLocaleString()})`,
          styles: { 
            halign: 'center', 
            fontSize: 13, 
            cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
            fontStyle: 'bold',
            textColor: [245, 158, 11]
          } 
        },
        { 
          content: `${data.redLots} (${data.redPCS.toLocaleString()})`,
          styles: { 
            halign: 'center', 
            fontSize: 13, 
            cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
            fontStyle: 'bold',
            textColor: [239, 68, 68]
          } 
        }
      ];
    });

    // Calculate overall totals by color category
    const overallTotals = exportData.reduce((acc, item) => {
      const stitchingDays = calculateStitchingDays(item.dateOfIssue);
      
      acc.totalLots += 1;
      acc.totalPCS += item.totalPCS || 0;
      
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
      redPCS: 0
    });

    // Add total row
    const totalRowBackground = [225, 239, 255];

    summaryBody.push([
      { 
        content: 'OVERALL TOTALS', 
        styles: { 
          halign: 'left', 
          fontSize: 9, 
          fontStyle: 'bold', 
          fillColor: totalRowBackground,
          cellPadding: { top: 6, right: 3, bottom: 6, left: 6 }
        } 
      },
      { 
        content: '-', 
        styles: { 
          halign: 'center', 
          fontSize: 9, 
          fontStyle: 'bold', 
          fillColor: totalRowBackground,
          cellPadding: { top: 6, right: 3, bottom: 6, left: 3 }
        } 
      },
      { 
        content: overallTotals.totalLots.toString(), 
        styles: { 
          halign: 'center', 
          fontSize: 9, 
          fontStyle: 'bold', 
          fillColor: totalRowBackground,
          cellPadding: { top: 6, right: 3, bottom: 6, left: 3 }
        } 
      },
      { 
        content: overallTotals.totalPCS.toLocaleString(), 
        styles: { 
          halign: 'center', 
          fontSize: 9, 
          fontStyle: 'bold', 
          fillColor: totalRowBackground,
          cellPadding: { top: 6, right: 3, bottom: 6, left: 3 }
        } 
      },
      { 
        content: `${overallTotals.greenLots} (${overallTotals.greenPCS.toLocaleString()})`,
        styles: { 
          halign: 'center', 
          fontSize: 9, 
          fontStyle: 'bold', 
          fillColor: totalRowBackground,
          textColor: [16, 185, 129],
          cellPadding: { top: 6, right: 3, bottom: 6, left: 3 }
        } 
      },
      { 
        content: `${overallTotals.yellowLots} (${overallTotals.yellowPCS.toLocaleString()})`,
        styles: { 
          halign: 'center', 
          fontSize: 9, 
          fontStyle: 'bold', 
          fillColor: totalRowBackground,
          textColor: [245, 158, 11],
          cellPadding: { top: 6, right: 3, bottom: 6, left: 3 }
        } 
      },
      { 
        content: `${overallTotals.redLots} (${overallTotals.redPCS.toLocaleString()})`,
        styles: { 
          halign: 'center', 
          fontSize: 9, 
          fontStyle: 'bold', 
          fillColor: totalRowBackground,
          textColor: [239, 68, 68],
          cellPadding: { top: 6, right: 3, bottom: 6, left: 3 }
        } 
      }
    ]);

    // Calculate column widths for supervisor summary table
    const totalAvailableWidth = contentWidth;
    const numberOfColumns = 7;
    const baseColumnWidth = totalAvailableWidth / numberOfColumns;

    const summaryColumnWidths = [
      baseColumnWidth * 1.2,
      baseColumnWidth * 0.9,
      baseColumnWidth * 0.9,
      baseColumnWidth * 1.0,
      baseColumnWidth * 1.2,
      baseColumnWidth * 1.2,
      baseColumnWidth * 0.7,
    ];

    // Create supervisor summary table
    autoTable(doc, {
      startY: y,
      head: [[
        { content: 'SUPERVISOR', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[0] } },
        { content: 'MANPOWER', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[1], fillColor: [15, 76, 129], textColor: [255, 255, 255] } },
        { content: 'TOTAL LOTS', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[2] } },
        { content: 'TOTAL PCS', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[3] } },
        { content: 'GREEN (1-6 Days)', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[4], textColor: [16, 185, 129] } },
        { content: 'YELLOW (7-15 Days)', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[5], textColor: [245, 158, 11] } },
        { content: 'RED (15+ Days)', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[6], textColor: [239, 68, 68] } }
      ]],
      body: summaryBody.map(row => row.map((cell, colIndex) => ({
        content: cell.content,
        styles: {
          ...cell.styles,
          cellWidth: summaryColumnWidths[colIndex]
        }
      }))),
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
        fontSize: 12,
        cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
        lineWidth: 0.3,
        lineColor: [220, 220, 220],
        textColor: textColor,
        font: 'helvetica',
        valign: 'middle'
      },
      columnStyles: {
        0: { cellWidth: summaryColumnWidths[0], halign: 'left' },
        1: { cellWidth: summaryColumnWidths[1], halign: 'center' },
        2: { cellWidth: summaryColumnWidths[2], halign: 'center' },
        3: { cellWidth: summaryColumnWidths[3], halign: 'center' },
        4: { cellWidth: summaryColumnWidths[4], halign: 'center' },
        5: { cellWidth: summaryColumnWidths[5], halign: 'center' },
        6: { cellWidth: summaryColumnWidths[6], halign: 'center' }
      },
      margin: { left: margin, right: margin },
      tableWidth: 'auto',
      showHead: 'everyPage',
      showFoot: false,
      pageBreak: 'auto',
      rowPageBreak: 'avoid'
    });

    // Add summary notes
    const summaryEndY = doc.lastAutoTable.finalY;
    const notesY = summaryEndY + 8;

    doc.setFontSize(8);
    doc.setFont('helvetica', 'italic');
    doc.setTextColor(0, 0, 0);

    const notes = [
      '• Green (1-6 Days): Good performance, within target timeline',
      '• Yellow (7-15 Days): Average performance, needs monitoring',
      '• Red (15+ Days): Critical, requires immediate attention',
      '• Emb/Print Date column represent Emb/Print Done Date',
      '• Issue Date is a Date when the Lots is Issued For Stitching',
    ];

    notes.forEach((note, index) => {
      doc.text(note, margin, notesY + (index * 4));
    });

    // Generate filename with filter information
    const today = new Date();
    let fileName = 'Stitching_Report';
    
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
    fileName += `_${exportData.length}_records`;
    
    // Remove any special characters and limit length
    fileName = fileName.replace(/[^\w\-_]/g, '_').substring(0, 100);
    
    doc.save(`${fileName}.pdf`);
  } catch (err) {
    setError(`Failed to export PDF: ${err.message}`);
  } finally {
    setLoading(false);
  }
}, [filteredData, filters, calculateStitchingDays, getStitchingDaysColor, getStitchingDaysTextColor, getEmbPrintDate, getLatestWipRemarks, getCompletionDateFormatted, isLotCompleted, abbreviateMWK, formatDateToDDMMYY]);

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
  // Change the loading text in isInitialLoad check:
  if (isInitialLoad && loading) {
    return (
      <div className="loading-container">
        <div className="spinner"></div>
        <p>Loading all stitching data...</p>
        <p className="loading-subtext">Loading completed lots in descending order by default</p>
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
              StickerReport
            </h1>
            <p className="subtitle">
              {filters.supervisor ? `${filters.supervisor} - Stitching Report` : 'Real-time tracking of stitching operations'}
              {data.length > 0 && (
                <span className="data-source">
                  • {filters.supervisor === 'Monu' && filters.lotStatus === 'Pending' ? 
                     "Showing Monu's PENDING lots by default" : 
                     filters.lotStatus ? `Showing ${filters.lotStatus.toLowerCase()} lots` : 
                     `Showing all lots for ${filters.supervisor}`}
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
                {dataCache.current.size > 0 ? '🟢' : '⚪'}
              </div>
              <div className="stat-label">Cache</div>
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
           <div className="global-search-container">
    <div className="search-input-wrapper">
      <span className="search-icon">🔍</span>
      <input
        type="text"
        name="globalSearch"
        value={filters.globalSearch}
        onChange={handleFilterChange}
        placeholder="Search across all fields (Lot Number, Fabric, Style, Brand, Party, Supervisor...)"
        className="global-search-input"
        disabled={loading}
      />
      {filters.globalSearch && (
        <button
          onClick={() => {
            setFilters(prev => ({ ...prev, globalSearch: '' }));
            setTimeout(() => {
              applyFilters({ ...filters, globalSearch: '' }, data);
            }, 0);
          }}
          className="clear-search-btn"
          type="button"
          disabled={loading}
        >
          ✕
        </button>
      )}
    </div>
    <div className="search-hint">
      Searches in: Lot Number, Fabric, Garment Type, Style, Brand, Party Name, Supervisor, Season, M/W/K, Direct Stitching
    </div>
  </div>
  
          <div className="filters-info">
            {/* UPDATED active filters count to handle dateRange object */}
            <span className="active-filters">
              {Object.entries(filters).filter(([key, value]) => {
                if (key === 'dateRange') {
                  const { from, to } = value;
                  return from || to;
                }
                return typeof value === 'string' && value.trim() !== '';
              }).length} active filters
            </span>
            <button
              onClick={clearFilters}
              className="btn-clear"
              disabled={loading}
            >
              {filters.supervisor === 'Monu' ? 'Clear Filters' : ''}
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
  <option value="">All Supervisors</option> {/* CHANGED FROM "Monu (Default)" */}
  {filterOptions.supervisor.map((option, index) => (
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
          
          <div className="filter-group">
            <label className="group-label">Additional Filters</label>
            <div className="group-filters">
              {renderDropdownFilter('brand', 'Brand')}
              {renderDropdownFilter('season', 'Season')}
              {renderDropdownFilter('mwk', 'M/W/K')}
                {renderDropdownFilter('sticker', 'Sticker')}
            </div>
          </div>
          
          <div className="filter-group">
            <label className="group-label">Status Filters</label>
            <div className="group-filters">
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
                  <option value="YES">Yes</option>
                  <option value="NO">No</option>
                </select>
              </div>
            </div>
          </div>
          
          {/* Add Date Range Filter Group */}
          <div className="filter-group">
            <label className="group-label">Date Range (Completion)</label>
            <div className="group-filters">
              <div className="filter-item">
                <label>From Date</label>
                <input
                  type="date"
                  name="dateFrom"
                  value={filters.dateRange.from}
                  onChange={handleFilterChange}
                  className="filter-input"
                  disabled={loading}
                />
              </div>
              <div className="filter-item">
                <label>To Date</label>
                <input
                  type="date"
                  name="dateTo"
                  value={filters.dateRange.to}
                  onChange={handleFilterChange}
                  className="filter-input"
                  disabled={loading}
                />
              </div>
              <div className="filter-item">
                <label>Clear Dates</label>
                <button
                  onClick={() => {
                    setFilters(prev => ({
                      ...prev,
                      dateRange: { from: '', to: '' }
                    }));
                  }}
                  className="btn-clear-date"
                  disabled={loading}
                >
                  Clear
                </button>
              </div>
            </div>
          </div>
          {/* Add Stitching Days Filter Group with checkboxes */}
{/* Stitching Days Filter Group with horizontal checkboxes */}
<div className="filter-group horizontal-layout">
  <label className="group-label">Stitching Days</label>
  <div className="group-filters">
    <div className="filter-item">
      <div className="checkbox-group-horizontal">
        <label className="checkbox-item-horizontal">
          <input
            type="checkbox"
            name="stitchingDaysFilter"
            value="green"
            checked={filters.stitchingDaysFilter.includes('green')}
            onChange={handleFilterChange}
            disabled={loading}
            className="checkbox-input"
          />
          <span className="checkbox-custom"></span>
          <span className="checkbox-text">
            <span className="color-indicator green"></span>
            1-6 days
            {filters.stitchingDaysFilter.includes('green') && 
              <span className="selection-count-badge"></span>}
          </span>
        </label>
        
        <label className="checkbox-item-horizontal">
          <input
            type="checkbox"
            name="stitchingDaysFilter"
            value="yellow"
            checked={filters.stitchingDaysFilter.includes('yellow')}
            onChange={handleFilterChange}
            disabled={loading}
            className="checkbox-input"
          />
          <span className="checkbox-custom"></span>
          <span className="checkbox-text">
            <span className="color-indicator yellow"></span>
            7-15 days
            {filters.stitchingDaysFilter.includes('yellow') && 
              <span className="selection-count-badge"></span>}
          </span>
        </label>
        
        <label className="checkbox-item-horizontal">
          <input
            type="checkbox"
            name="stitchingDaysFilter"
            value="red"
            checked={filters.stitchingDaysFilter.includes('red')}
            onChange={handleFilterChange}
            disabled={loading}
            className="checkbox-input"
          />
          <span className="checkbox-custom"></span>
          <span className="checkbox-text">
            <span className="color-indicator red"></span>
            15+ days
            {filters.stitchingDaysFilter.includes('red') && 
              <span className="selection-count-badge"></span>}
          </span>
        </label>
        
        {/* Clear button in same row */}
        {filters.stitchingDaysFilter.length > 0 && (
          <button
            type="button"
            onClick={() => setFilters(prev => ({ ...prev, stitchingDaysFilter: [] }))}
            className="btn-clear-selection-horizontal"
            disabled={loading}
          >
            Clear ({filters.stitchingDaysFilter.length})
          </button>
        )}
      </div>
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
                {/* UPDATED to handle dateRange object */}
                {Object.entries(filters).some(([key, value]) => {
                  if (key === 'dateRange') {
                    const { from, to } = value;
                    return from || to;
                  }
                  return key !== 'supervisor' && typeof value === 'string' && value.trim() !== '';
                }) && ' (filtered)'}
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
    {HEADERS().map((header, index) => (
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
    
    return (
      <tr key={index} className="table-row">
        <td className="text-center font-semibold">{index + 1}</td>
        <td className="text-center">
          <span className="lot-number">{item.lotNumber || 'N/A'}</span>
        </td>
        <td className="text-center">{item.fabric || 'N/A'}</td>
        <td className="text-center">{item.garmentType || 'N/A'}</td>
        <td className="text-center">{item.style || 'N/A'}</td>
        <td className="text-center">{item.brand || 'N/A'}</td>
        <td className="text-center">{item.partyName || 'N/A'}</td>
        <td className="text-center">
          <span className="supervisor-name" style={{ color: '#3b82f6', fontWeight: 'bold' }}>
            {item.supervisor || 'N/A'}
          </span>
        </td>
        <td className="text-center">{item.season || 'N/A'}</td>
        <td className="text-center">
          <span className="mwk-abbreviation" style={{
            fontWeight: 'bold',
            color: abbreviateMWK(item.mwk) === 'M' ? '#3b82f6' :
                   abbreviateMWK(item.mwk) === 'W' ? '#ef4444' :
                   abbreviateMWK(item.mwk) === 'K' ? '#10b981' :
                   abbreviateMWK(item.mwk) === 'G' ? '#8b5cf6' : '#6b7280'
          }}>
            {abbreviateMWK(item.mwk)}
          </span>
        </td>
        <td className="text-center">{item.directStitching || 'N/A'}</td>
        
        {/* Date of Issue */}
        <td className="text-center">
          <span className="issue-date">
            {formatDateToDDMMYY(item.dateOfIssue)}
          </span>
        </td>
        
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
            {embPrintDate !== '-' ? formatDateToDDMMYY(embPrintDate) : embPrintDate}
          </span>
        </td>
        
        {/* WIP Status */}
        <td className="text-center">
          <span className="wip-status" style={{
            color: isCompleted ? '#10b981' : '#475569',
            fontWeight: isCompleted ? 'bold' : '500'
          }}>
            {isCompleted ? 'Done' : getLatestWipRemarks(item.wipStatus, false)}
          </span>
        </td>
        
        <td className="text-center font-semibold">
          <span className="total-pcs">
            {item.totalPCS > 0 ? item.totalPCS.toLocaleString() : 'N/A'}
          </span>
        </td>
        <td className="text-center">
          <span className={`completion-date ${isCompleted ? 'completed' : 'pending'}`}>
            {completionDate ? formatDateToDDMMYY(completionDate) : (isCompleted ? 'Completed' : 'Pending')}
          </span>
        </td>
        <td className="text-center">
          <span className={`status-badge ${isCompleted ? 'completed-badge' : 'pending-badge'}`}>
            {isCompleted ? 'Completed' : 'Pending'}
          </span>
        </td>
        {/* ADD THIS LINE FOR STICKER */}
        <td className="text-center">
          <span className="sticker-value" style={{
            fontWeight: '500',
            color: item.sticker ? '#8b5cf6' : '#94a3b8'
          }}>
            {item.sticker || 'N/A'}
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
          background: #ffffff;
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
// Add these styles to your existing CSS (around the filters section styles)

.global-search-container {
  margin-bottom: 24px;
  padding-bottom: 24px;
  border-bottom: 1px solid #e2e8f0;
}

.search-input-wrapper {
  position: relative;
  display: flex;
  align-items: center;
}

.search-icon {
  position: absolute;
  left: 12px;
  color: #64748b;
  font-size: 16px;
  z-index: 1;
}

.global-search-input {
  width: 100%;
  padding: 12px 20px 12px 40px;
  border: 2px solid #e2e8f0;
  border-radius: 10px;
  font-size: 14px;
  background: white;
  transition: all 0.2s;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
}

.global-search-input:focus {
  outline: none;
  border-color: #3b82f6;
  box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.1);
}

.global-search-input:disabled {
  background: #f8fafc;
  cursor: not-allowed;
}
  .sticker-value {
  font-size: 12px;
  padding: 4px 8px;
  border-radius: 4px;
  background: #f3e8ff;
  border: 1px solid #d8b4fe;
  display: inline-block;
}

.clear-search-btn {
  position: absolute;
  right: 12px;
  background: #e2e8f0;
  border: none;
  border-radius: 50%;
  width: 24px;
  height: 24px;
  display: flex;
  align-items: center;
  justify-content: center;
  cursor: pointer;
  font-size: 12px;
  color: #64748b;
  transition: all 0.2s;
}

.clear-search-btn:hover:not(:disabled) {
  background: #cbd5e1;
  color: #475569;
}

.clear-search-btn:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

.search-hint {
  font-size: 11px;
  color: #94a3b8;
  margin-top: 6px;
  padding-left: 12px;
  font-style: italic;
}

/* Responsive styles */
@media (max-width: 768px) {
  .global-search-input {
    font-size: 13px;
    padding: 10px 16px 10px 36px;
  }
  
  .search-hint {
    font-size: 10px;
  }
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
          .mwk-abbreviation {
  font-weight: 600;
  padding: 4px 8px;
  border-radius: 4px;
  background: #f8fafc;
  border: 1px solid #e2e8f0;
  display: inline-block;
  min-width: 20px;
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
          /* Horizontal Checkbox Styles */
.checkbox-group-horizontal {
  display: flex;
  align-items: center;
  gap: 16px;
  margin-top: 4px;
  flex-wrap: wrap;
}

.checkbox-item-horizontal {
  display: flex;
  align-items: center;
  gap: 6px;
  cursor: pointer;
  padding: 8px 12px;
  font-size: 13px;
  color: #475569;
  background: #f8fafc;
  border: 1px solid #e2e8f0;
  border-radius: 6px;
  transition: all 0.2s;
  min-height: 36px;
}

.checkbox-item-horizontal:hover {
  background: #f1f5f9;
  border-color: #cbd5e1;
}

.checkbox-input {
  display: none;
}

.checkbox-custom {
  width: 16px;
  height: 16px;
  border: 2px solid #cbd5e1;
  border-radius: 4px;
  background: white;
  position: relative;
  transition: all 0.2s;
  flex-shrink: 0;
}

.checkbox-input:checked + .checkbox-custom {
  background: #3b82f6;
  border-color: #3b82f6;
}

.checkbox-input:checked + .checkbox-custom::after {
  content: '✓';
  position: absolute;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%);
  color: white;
  font-size: 11px;
  font-weight: bold;
}

.checkbox-text {
  display: flex;
  align-items: center;
  gap: 6px;
  font-weight: 500;
  white-space: nowrap;
}

.color-indicator {
  width: 14px;
  height: 14px;
  border-radius: 3px;
  display: inline-block;
  border: 1px solid rgba(0, 0, 0, 0.1);
}

.color-indicator.green {
  background: #10b981;
}

.color-indicator.yellow {
  background: #f59e0b;
}

.color-indicator.red {
  background: #ef4444;
}

.btn-clear-selection-horizontal {
  margin-left: 8px;
  padding: 8px 12px;
  background: #f1f5f9;
  color: #64748b;
  border: 1px solid #e2e8f0;
  border-radius: 6px;
  font-size: 12px;
  font-weight: 500;
  cursor: pointer;
  transition: all 0.2s;
  height: 36px;
  white-space: nowrap;
  display: flex;
  align-items: center;
}

.btn-clear-selection-horizontal:hover:not(:disabled) {
  background: #e2e8f0;
  color: #475569;
}

.btn-clear-selection-horizontal:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

/* Selected count badge */
.selection-count-badge {
  background: #3b82f6;
  color: white;
  font-size: 11px;
  font-weight: 600;
  padding: 2px 6px;
  border-radius: 10px;
  margin-left: 4px;
}

/* Active selection highlight */
.checkbox-input:checked ~ .checkbox-text {
  font-weight: 600;
  color: #1e293b;
}

/* Make filter group more compact */
.filter-group.horizontal-layout {
  grid-column: span 2;
}

.filter-group.horizontal-layout .group-filters {
  padding-top: 4px;
}
  .issue-date {
  font-weight: 500;
  color: #8b5cf6; /* Purple color to match PDF */
  background: #f3e8ff;
  padding: 4px 8px;
  border-radius: 12px;
  font-size: 12px;
  display: inline-block;
  border: 1px solid #d8b4fe;
}

/* Responsive adjustments */
@media (max-width: 768px) {
  .checkbox-group-horizontal {
    gap: 12px;
  }
  
  .checkbox-item-horizontal {
    padding: 6px 10px;
    font-size: 12px;
  }
  
  .btn-clear-selection-horizontal {
    padding: 6px 10px;
    font-size: 11px;
  }
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

        .btn-clear-date {
          padding: 6px 16px;
          background: #f1f5f9;
          color: #64748b;
          border: 1px solid #e2e8f0;
          border-radius: 6px;
          font-size: 14px;
          font-weight: 500;
          cursor: pointer;
          transition: all 0.2s;
          margin-top: 20px;
        }

        .btn-clear-date:hover:not(:disabled) {
          background: #e2e8f0;
          color: #475569;
        }

        .btn-clear-date:disabled {
          opacity: 0.5;
          cursor: not-allowed;
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

        .table-row:hover {
          background: #f8fafc;
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

export default StickerReport;