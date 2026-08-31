import React, { useState, useEffect, useCallback, useRef } from 'react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { GOOGLE_API_KEY, SPREADSHEET_IDS, fetchSheetDataFromBackend } from './config';

const getDirectImageUrl = (url) => {
  if (!url) return '';
  if (url.includes('drive.google.com')) {
    const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (match && match[1]) {
      return `https://lh3.googleusercontent.com/u/0/d/${match[1]}`;
    }
    const idMatch = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (idMatch && idMatch[1]) {
      return `https://lh3.googleusercontent.com/u/0/d/${idMatch[1]}`;
    }
  }
  return url;
};

const MultiSelectDropdown = ({ options = [], selectedValues = [], onChange, placeholder, disabled }) => {
  const [isOpen, setIsOpen] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const containerRef = useRef(null);

  useEffect(() => {
    const handleClickOutside = (event) => {
      if (containerRef.current && !containerRef.current.contains(event.target)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, []);

  const handleToggleOption = (option) => {
    const isSelected = selectedValues.includes(option);
    let newSelected;
    if (isSelected) {
      newSelected = selectedValues.filter(v => v !== option);
    } else {
      newSelected = [...selectedValues, option];
    }
    onChange(newSelected);
  };

  const filteredOptions = options.filter(option =>
    String(option || '').toLowerCase().includes(searchTerm.toLowerCase())
  );

  const displayText = () => {
    if (!selectedValues || selectedValues.length === 0) {
      return placeholder || 'All';
    }
    if (selectedValues.length === 1) {
      return selectedValues[0];
    }
    if (selectedValues.length === options.length && options.length > 0) {
      return placeholder || 'All';
    }
    return `${selectedValues.length} Selected`;
  };

  return (
    <div className="multiselect-container" ref={containerRef} style={{ position: 'relative', width: '100%' }}>
      <div
        className="multiselect-select"
        onClick={() => !disabled && setIsOpen(!isOpen)}
        style={{
          padding: '8px 12px',
          background: disabled ? '#f1f5f9' : 'white',
          border: '1px solid #cbd5e1',
          borderRadius: '8px',
          cursor: disabled ? 'not-allowed' : 'pointer',
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          fontSize: '13px',
          color: selectedValues.length > 0 ? '#1e293b' : '#64748b',
          minHeight: '38px',
          boxSizing: 'border-box'
        }}
      >
        <span style={{
          whiteSpace: 'nowrap',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          maxWidth: '85%',
          display: 'block',
          fontWeight: selectedValues.length > 0 ? '600' : '400'
        }}>
          {displayText()}
        </span>
        <span style={{ fontSize: '10px', color: '#64748b' }}>{isOpen ? '▲' : '▼'}</span>
      </div>

      {isOpen && !disabled && (
        <div
          className="multiselect-dropdown"
          style={{
            position: 'absolute',
            top: '100%',
            left: 0,
            right: 0,
            background: 'white',
            border: '1px solid #cbd5e1',
            borderRadius: '8px',
            marginTop: '4px',
            boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.15)',
            zIndex: 1000,
            maxHeight: '250px',
            overflowY: 'auto',
            padding: '8px',
            display: 'flex',
            flexDirection: 'column',
            gap: '4px'
          }}
        >
          <input
            type="text"
            className="multiselect-search"
            placeholder="Search supervisors..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            style={{
              padding: '6px 10px',
              border: '1px solid #cbd5e1',
              borderRadius: '6px',
              width: '100%',
              fontSize: '12px',
              marginBottom: '6px',
              boxSizing: 'border-box',
              outline: 'none'
            }}
          />
          <div style={{ display: 'flex', gap: '6px', marginBottom: '6px' }}>
            <button
              type="button"
              onClick={() => onChange([])}
              style={{
                flex: 1,
                fontSize: '11px',
                padding: '4px',
                border: '1px solid #cbd5e1',
                borderRadius: '4px',
                background: '#f8fafc',
                cursor: 'pointer'
              }}
            >
              Clear All
            </button>
            <button
              type="button"
              onClick={() => onChange(options)}
              style={{
                flex: 1,
                fontSize: '11px',
                padding: '4px',
                border: '1px solid #cbd5e1',
                borderRadius: '4px',
                background: '#f8fafc',
                cursor: 'pointer'
              }}
            >
              Select All
            </button>
          </div>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '2px', overflowY: 'auto' }}>
            {filteredOptions.length === 0 ? (
              <span style={{ fontSize: '12px', color: '#94a3b8', textAlign: 'center', padding: '8px' }}>
                No options found
              </span>
            ) : (
              filteredOptions.map(option => {
                const isSelected = Array.isArray(selectedValues) && selectedValues.includes(option);
                return (
                  <label
                    key={option}
                    style={{
                      display: 'flex',
                      alignItems: 'center',
                      gap: '8px',
                      padding: '5px 8px',
                      borderRadius: '4px',
                      cursor: 'pointer',
                      background: isSelected ? '#f1f5f9' : 'transparent',
                      fontSize: '12px',
                      color: '#1e293b'
                    }}
                  >
                    <input
                      type="checkbox"
                      checked={isSelected}
                      onChange={() => handleToggleOption(option)}
                      style={{ cursor: 'pointer' }}
                    />
                    <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                      {option}
                    </span>
                  </label>
                );
              })
            )}
          </div>
        </div>
      )}
    </div>
  );
};

// Rate-limit proof fetch helper with Exponential Backoff
async function fetchSheetWithRetry(url, options = {}, retries = 4, baseDelayMs = 500) {
  let attempt = 0;
  while (true) {
    try {
      const response = await fetch(url, options);
      if (response.ok) return response;
      const text = await response.text();
      if ((response.status === 429 || response.status >= 500) && attempt < retries) {
        const delay = baseDelayMs * Math.pow(2, attempt) + Math.random() * 250;
        await new Promise((resolve) => setTimeout(resolve, delay));
        attempt++;
        continue;
      }
      throw new Error(`HTTP error! Status: ${response.status} - ${text}`);
    } catch (err) {
      if (attempt < retries && err.name !== 'AbortError') {
        const delay = baseDelayMs * Math.pow(2, attempt) + Math.random() * 250;
        await new Promise((resolve) => setTimeout(resolve, delay));
        attempt++;
        continue;
      }
      throw err;
    }
  }
}

const StitchingCompleteLot = () => {
  // Data states
  const [data, setData] = useState([]);
  const [filteredData, setFilteredData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [isInitialLoad, setIsInitialLoad] = useState(true);
  const [workingUpdatesMap, setWorkingUpdatesMap] = useState(new Map());
  const [viewImageSrc, setViewImageSrc] = useState(null);

  // Performance optimization state for large datasets
  const [displayLimit, setDisplayLimit] = useState(200);

  // Cache management
  const dataCache = useRef(new Map());
  const lastFetchTime = useRef(new Map());
  const cuttingDataCache = useRef({ data: null, timestamp: 0 });
  const designImageCache = useRef({ data: null, timestamp: 0 });
  const CACHE_DURATION = 5 * 60 * 1000; // 5 minutes cache
  const BATCH_SIZE = 500; // Process data in batches for large datasets

  // Filters
  // Add this to your filters state (around line 33-34)
  const [filters, setFilters] = useState({
    globalSearch: '',
    lotNumber: '',
    fabric: [],
    garmentType: [],
    style: [],
    brand: [],
    partyName: [],
    supervisor: [],
    season: [],
    mwk: [],
    directStitching: [],
    challanHistory: '',
    wipStatus: [],
    completedStatus: '',
    lotStatus: 'Completed',
    dateRange: { from: '', to: '' },
    stitchingDaysFilter: []
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
    stitchingDaysFilter: ['green', 'yellow', 'red'] // ADDED: Options for checkboxes
  });

  // Google Sheets configuration sourced from .env via config.js
  const SPREADSHEET_ID = SPREADSHEET_IDS.MAIN;
  const API_KEY = GOOGLE_API_KEY;
  const SHEET_NAME = 'Index';
  const CUTTING_SHEET_NAME = 'Cutting';
  const CUTTING_RANGE = `${CUTTING_SHEET_NAME}!A1:ZZ200000`;
  const WORKING_UPDATES_SPREADSHEET_ID = SPREADSHEET_IDS.WORKING_UPDATES;
  const WORKING_UPDATES_SHEET_NAME = 'Working Updates';
  const WORKING_UPDATES_RANGE = `${WORKING_UPDATES_SHEET_NAME}!A1:ZZ20000`;

  const normalizeKey = (s = "") => {
    return String(s || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
  };

  // Remove the const HEADERS array and replace with this function
  // Replace the HEADERS function with this version
  const HEADERS = () => {
    return [
      'Sr.No',
      'Image',
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
      'WIP Status',  // Always show WIP Status
      'Total PCS',
      'Completion Date',
      'Status'
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
  // Get Pintu status for a specific lot (for PDF only)
  const getPintuStatusForPDF = useCallback((lotNumber) => {
    if (!lotNumber) return 'N/A';
    const updates = workingUpdatesMap.get(lotNumber.trim());
    return updates?.pintuStatus || 'N/A';
  }, [workingUpdatesMap]);

  // Get EA status for a specific lot (for PDF only)
  const getEAStatusForPDF = useCallback((lotNumber) => {
    if (!lotNumber) return 'N/A';
    const updates = workingUpdatesMap.get(lotNumber.trim());
    return updates?.eaStatus || 'N/A';
  }, [workingUpdatesMap]);

  const normalizeAndCapitalize = useCallback((text) => {
    if (!text || text.trim() === '') return '';

    const trimmed = text.trim().toLowerCase();
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
    const now = Date.now();
    if (cuttingDataCache.current.data && (now - cuttingDataCache.current.timestamp < CACHE_DURATION)) {
      return cuttingDataCache.current.data;
    }
    try {
      const res = await fetchSheetDataFromBackend(SPREADSHEET_ID, CUTTING_RANGE);
      const val = res.ok ? res.values : [];
      cuttingDataCache.current = { data: val, timestamp: Date.now() };
      return val;
    } catch (err) {
      console.error('Error fetching cutting data:', err);
      return cuttingDataCache.current.data || [];
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

  // Calculate stitching days - For completed lots: Completion Date - Date of Issue
  // For pending lots: Today's Date - Date of Issue
  // Calculate stitching days - For completed lots: Completion Date - Date of Issue
  // For pending lots: Today's Date - Date of Issue
  const calculateStitchingDays = useCallback((dateOfIssue, completedStatus, isCompleted = false) => {
    if (!dateOfIssue || dateOfIssue.trim() === '') return 0;

    try {
      // Parse Date of Issue
      let issueDate = null;
      let testDate = new Date(dateOfIssue);
      if (!isNaN(testDate.getTime())) {
        issueDate = testDate;
      } else {
        const parts = dateOfIssue.split(/[\/\-]/);
        if (parts.length === 3) {
          const day = parseInt(parts[0], 10);
          const month = parseInt(parts[1], 10);
          const year = parseInt(parts[2], 10);
          const fullYear = year < 100 ? 2000 + year : year;
          testDate = new Date(fullYear, month - 1, day);
          if (!isNaN(testDate.getTime())) {
            issueDate = testDate;
          }
        }
      }

      if (!issueDate) return 0;

      let endDate;

      // FOR COMPLETED LOTS: Use Completion Date
      if (isCompleted && completedStatus) {
        try {
          if (typeof completedStatus === 'string' && completedStatus.startsWith('[')) {
            const statusArray = JSON.parse(completedStatus);
            if (Array.isArray(statusArray) && statusArray.length > 0) {
              const completeEntry = statusArray.find(entry =>
                entry.status && entry.status.toLowerCase().includes('complete')
              );
              if (completeEntry && completeEntry.timestamp) {
                endDate = new Date(completeEntry.timestamp);
              }
            }
          }

          if (!endDate || isNaN(endDate.getTime())) {
            endDate = new Date();
          }
        } catch (error) {
          console.error('Error parsing completion date:', error);
          endDate = new Date();
        }
      }
      // FOR PENDING LOTS: Use Today's Date
      else {
        endDate = new Date();
      }

      const timeDiff = endDate.getTime() - issueDate.getTime();
      const days = Math.floor(timeDiff / (1000 * 3600 * 24));
      return Math.max(0, days);
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
      lotStatus: ['Pending', 'Completed']
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
        } else if (key === 'lotStatus') {
          // Skip for lotStatus as it's predefined
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
      if (key !== 'lotStatus') {
        options[key] = Array.from(optionSets[key]).sort();
      }
    });

    setFilterOptions(options);
  }, [getLatestWipRemarks, getCompletedStatusText, normalizeAndCapitalize, isLotCompleted]);

  // Generate cache key based on filters
  const getCacheKey = useCallback((supervisor = '') => {
    if (Array.isArray(supervisor)) {
      return supervisor.length > 0 ? supervisor.slice().sort().join('_') : 'all';
    }
    if (typeof supervisor === 'string' && supervisor.trim() !== '') {
      return supervisor.trim();
    }
    return 'all';
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
  // Fetch working updates data for Pintu and EA status
  const fetchWorkingUpdatesData = useCallback(async () => {
    try {
      const res = await fetchSheetDataFromBackend(WORKING_UPDATES_SPREADSHEET_ID, WORKING_UPDATES_RANGE);
      const rows = res.ok ? res.values : [];

      if (rows.length === 0) return new Map();

      // Create a map of lot numbers to their Pintu and EA status
      const updatesMap = new Map();

      // Assuming the Working Updates sheet has columns:
      // Lot Number, Fabric, Brand, Cutting Qty, Update History, Last Updated, Last Updated By
      rows.slice(1).forEach(row => {
        if (row.length >= 5) {
          const lotNumber = row[0] || '';
          const updateHistory = row[4] || '[]'; // Update History column

          if (lotNumber) {
            // Parse update history to get Pintu and EA status
            let pintuStatus = 'N/A';
            let eaStatus = 'N/A';

            try {
              const history = JSON.parse(updateHistory);
              if (Array.isArray(history)) {
                // Get latest Pintu status
                const pintuUpdates = history.filter(h =>
                  h.updatedBy?.toLowerCase() === 'pintu'
                );
                if (pintuUpdates.length > 0) {
                  const latestPintu = pintuUpdates[pintuUpdates.length - 1];
                  pintuStatus = latestPintu.updateType || latestPintu.remarks || 'Updated';
                }

                // Get latest EA/WA status
                const eaUpdates = history.filter(h =>
                  h.updatedBy?.toLowerCase() === 'ea' ||
                  h.updatedBy?.toLowerCase() === 'wa'
                );
                if (eaUpdates.length > 0) {
                  const latestEA = eaUpdates[eaUpdates.length - 1];
                  eaStatus = latestEA.updateType || latestEA.remarks || 'Updated';
                }
              }
            } catch (e) {
              console.error('Error parsing update history for lot:', lotNumber, e);
            }

            updatesMap.set(lotNumber.trim(), {
              pintuStatus,
              eaStatus
            });
          }
        }
      });

      setWorkingUpdatesMap(updatesMap);
      return updatesMap;
    } catch (err) {
      console.error('Error fetching working updates data:', err);
      return new Map();
    }
  }, [WORKING_UPDATES_SPREADSHEET_ID, WORKING_UPDATES_RANGE, API_KEY]);

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

      // Fetch index data via backend
      const res = await fetchSheetDataFromBackend(SPREADSHEET_ID, SHEET_NAME);
      const rows = res.ok ? res.values : [];

      if (!rows || rows.length === 0) {
        return [];
      }

      const sheetHeaders = rows[0];

      // Create a more robust column mapping
      const createColumnMapping = () => {
        const mapping = {};

        // Define all possible column names and their variations
        const columnDefinitions = [
          { keys: ['Lot Number', 'LotNumber', 'lotNumber'], target: 'lotNumber' },
          { keys: ['Image', 'Image URL', 'IMAGE', 'ImageURL', 'img'], target: 'image' },
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
          { keys: ['Sizes'], target: 'sizes' }
        ];

        // Find each column in the headers
        columnDefinitions.forEach(definition => {
          for (const key of definition.keys) {
            const index = sheetHeaders.findIndex(
              header => header && normalizeKey(header) === normalizeKey(key)
            );
            if (index !== -1) {
              mapping[definition.target] = index;
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

      // Fetch Design Image fallback from JobOrder sheet via backend
      try {
        const now = Date.now();
        let designMap = designImageCache.current.data;
        if (!designMap || (now - designImageCache.current.timestamp > CACHE_DURATION)) {
          const designRes = await fetchSheetDataFromBackend(SPREADSHEET_IDS.JOBORDER, 'Design Image!A1:ZZZ');
          const designRows = designRes.ok ? designRes.values : [];
          designMap = {};
          if (designRows.length > 0) {
            const headers = designRows[0];
            const lotCol = headers.findIndex(h => h && h.trim().toLowerCase().includes('lot'));
            const imgCol = headers.findIndex(h => h && h.trim().toLowerCase().includes('image'));
            if (lotCol !== -1 && imgCol !== -1) {
              designRows.slice(1).forEach(r => {
                const lot = String(r?.[lotCol] || '').trim().toUpperCase();
                const img = String(r?.[imgCol] || '').trim();
                if (lot && img) {
                  designMap[lot] = getDirectImageUrl(img);
                }
              });
            }
          }
          designImageCache.current = { data: designMap, timestamp: Date.now() };
        }
        if (designMap) {
          processedData.forEach(item => {
            const lotKey = String(item.lotNumber || '').trim().toUpperCase();
            if ((!item.image || item.image.trim() === '') && designMap[lotKey]) {
              item.image = designMap[lotKey];
            } else if (item.image) {
              item.image = getDirectImageUrl(item.image);
            }
          });
        }
      } catch (err) {
        console.error("Error fetching Design Image fallback:", err);
      }

      // Filter for supervisor if specified
      let filteredData = processedData;
      if (supervisor) {
        if (Array.isArray(supervisor) && supervisor.length > 0) {
          filteredData = processedData.filter(item =>
            item.supervisor && supervisor.some(s => normalizeText(s) === normalizeText(item.supervisor))
          );
        } else if (typeof supervisor === 'string' && supervisor.trim() !== '') {
          filteredData = processedData.filter(item =>
            item.supervisor && normalizeText(item.supervisor) === normalizeText(supervisor)
          );
        }
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
          // Search in multiple fields
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
            'directStitching'
          ];

          // Check if any of the direct fields match
          const directMatch = searchFields.some(field => {
            const fieldValue = item[field] ? normalizeText(item[field]) : '';
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

          if (!value || (typeof value === 'string' && value.trim() === '') || (Array.isArray(value) && value.length === 0)) {
            return true;
          }

          // Handle stitchingDaysFilter array (MUST BE CHECKED BEFORE GENERIC ARRAY HANDLER!)
          if (key === 'stitchingDaysFilter') {
            const selectedFilters = value;
            if (!selectedFilters || selectedFilters.length === 0) return true;

            const isCompleted = isLotCompleted(item.completedStatus);
            const stitchingDays = calculateStitchingDays(item.dateOfIssue, item.completedStatus, isCompleted);

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

          // Handle array filters (multi-select for supervisor, fabric, garmentType, style, brand, partyName, season, mwk, directStitching, wipStatus)
          if (Array.isArray(value)) {
            if (value.length === 0) return true;

            if (key === 'supervisor') {
              const itemSupervisor = item.supervisor ? normalizeText(item.supervisor) : '';
              return value.some(v => normalizeText(v) === itemSupervisor);
            }

            if (key === 'wipStatus') {
              const isCompleted = isLotCompleted(item.completedStatus);
              const latestRemarks = getLatestWipRemarks(item[key], isCompleted);
              const remarkNorm = normalizeText(latestRemarks);
              return value.some(v => remarkNorm.includes(normalizeText(v)));
            }

            const itemVal = item[key] ? normalizeText(item[key]) : '';
            return value.some(v => {
              const normV = normalizeText(v);
              return itemVal === normV || itemVal.includes(normV);
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

          // Handle wipStatus filter (single string fallback)
          if (key === 'wipStatus') {
            const isCompleted = isLotCompleted(item.completedStatus);
            const latestRemarks = getLatestWipRemarks(item[key], isCompleted);
            return normalizeText(latestRemarks).includes(filterValue);
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

  // Initial load - fetch ALL data
  useEffect(() => {
    const loadInitialData = async () => {
      try {
        setLoading(true);
        setIsInitialLoad(true);

        // Fetch working updates data for PDF
        await fetchWorkingUpdatesData();

        // Fetch ALL stitching data
        const allData = await fetchDataForSupervisor('');
        setData(allData);

        // Show only COMPLETED lots by default
        const completedData = allData.filter(item => isLotCompleted(item.completedStatus));

        // Sort completed data in descending order
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
  }, [fetchDataForSupervisor, fetchWorkingUpdatesData, extractFilterOptions, isLotCompleted, sortDataByCompletionDate]);

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

  const handleSupervisorChange = useCallback((selectedSupervisors) => {
    const newFilters = { ...filters, supervisor: selectedSupervisors };
    setFilters(newFilters);
    setTimeout(() => {
      applyFilters(newFilters, data);
    }, 0);
  }, [filters, data, applyFilters]);

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
      globalSearch: '',
      lotNumber: '',
      fabric: [],
      garmentType: [],
      style: [],
      brand: [],
      partyName: [],
      supervisor: [],
      season: [],
      mwk: [],
      directStitching: [],
      challanHistory: '',
      wipStatus: [],
      completedStatus: '',
      lotStatus: 'Completed',
      dateRange: { from: '', to: '' },
      stitchingDaysFilter: []
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
            isLotCompleted(item.completedStatus) ? 'Completed' : 'Pending'
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

      // NEW: Function to extract stage from WIP status
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
        } else if (remarks.includes('tailor working')) {
          return 'Tailor Working';
        } else if (remarks.includes('emb pending')) {
          return 'Emb Pending';
        } else {
          return 'On Stitching';
        }
      };

      // Helper function to check if status was updated today
      const isStatusUpdatedToday = (wipStatus, completedStatus) => {
        // If lot is completed, we don't need to check
        if (isLotCompleted(completedStatus)) {
          return true; // Completed lots are considered updated
        }

        if (!wipStatus || wipStatus.trim() === '') {
          return false;
        }

        try {
          if (typeof wipStatus === 'string' && !wipStatus.startsWith('[')) {
            // If it's a simple string, assume it's not updated today
            return false;
          }

          const statusArray = JSON.parse(wipStatus);

          if (!Array.isArray(statusArray) || statusArray.length === 0) {
            return false;
          }

          // Get the latest status
          const sortedStatuses = [...statusArray].sort((a, b) => {
            const dateA = new Date(a.timestamp).getTime();
            const dateB = new Date(b.timestamp).getTime();
            return dateB - dateA;
          });

          const latestStatus = sortedStatuses[0];

          if (!latestStatus || !latestStatus.timestamp) {
            return false;
          }

          // Check if the latest status timestamp is from today
          const today = new Date();
          const statusDate = new Date(latestStatus.timestamp);

          // Compare dates (ignoring time)
          return statusDate.toDateString() === today.toDateString();

        } catch (error) {
          console.error('Error checking if status updated today:', error);
          return false;
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
      const notUpdatedColor = [185, 28, 28]; // Dark red for "Not updated"
      const pintuColor = [147, 51, 234]; // Purple for Pintu
      const eaColor = [219, 39, 119]; // Pink for EA
      const stageAnalysisColor = [59, 130, 246]; // Blue for stage analysis

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
        if (filters.supervisor && (Array.isArray(filters.supervisor) ? filters.supervisor.length > 0 : (typeof filters.supervisor === 'string' && filters.supervisor.trim() !== ''))) {
          supervisorName = Array.isArray(filters.supervisor)
            ? filters.supervisor.join(', ').toUpperCase()
            : filters.supervisor.toUpperCase();
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
            if (key === 'stitchingDaysFilter' || key === 'supervisor') {
              return Array.isArray(value) && value.length > 0;
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

      // Dynamic headers - ALWAYS include WIP Status, Pintu, and EA
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
        'WIP Status',  // Always included
        'Pintu',        // NEW COLUMN
        'EA',           // NEW COLUMN
        'Total PCS',
        'Complete Date',
        'Lot Status'
      ];

      // Optimized column widths for better fit - Adjusted to accommodate new columns
      const columnWidths = {
        0: 8,    // Sr
        1: 17,   // Lot No
        2: 26,   // Fabric
        3: 27,   // Garment
        4: 28,   // Style
        5: 28,   // Brand
        6: 15,   // Party
        7: 23,   // Supervisor
        8: 8,    // Season
        9: 11,   // M/W/K
        10: 11,  // Direct
        11: 18,  // Issue Date
        12: 13,  // Days
        13: 18,  // Emb/Print
        14: 45,  // WIP Status
        15: 25,  // Pintu (INCREASED from 15 to 20)
        16: 25,  // EA (INCREASED from 15 to 20)
        17: 22,  // Total PCS
        18: 22,  // Complete Date
        19: 15   // Lot Status
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
            fontSize: 8,
            cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
          }
        }))
      ];

      // Prepare table body with color coding
      const body = exportData.map((item, rowIndex) => {
        const isCompleted = isLotCompleted(item.completedStatus);
        const stitchingDays = calculateStitchingDays(item.dateOfIssue, item.completedStatus, isCompleted);
        const stitchingDaysColor = getStitchingDaysColor(stitchingDays);
        const stitchingDaysTextColor = getStitchingDaysTextColor(stitchingDays);
        const wipRemarks = getLatestWipRemarks(item.wipStatus);
        const embPrintDate = getEmbPrintDate(item.challanHistory);
        const abbreviatedParty = abbreviatePartyName(item.partyName);
        const abbreviatedSeason = abbreviateSeason(item.season);
        const totalPCS = item.totalPCS || 0;
        const completionDate = getCompletionDateFormatted(item.completedStatus);
        // const isCompleted = isLotCompleted(item.completedStatus);

        // Generate sample data for Pintu and EA columns
        const pintuValue = getPintuStatusForPDF ? getPintuStatusForPDF(item.lotNumber) : '';
        const eaValue = getEAStatusForPDF ? getEAStatusForPDF(item.lotNumber) : '';

        // NEW: Check if status was updated today
        const isUpdatedToday = isStatusUpdatedToday(item.wipStatus, item.completedStatus);
        const wipDisplayValue = isCompleted
          ? 'Done'
          : (isUpdatedToday ? wipRemarks : 'Not updated');

        const lotStatus = isCompleted ? 'Completed' : 'Pending';
        const issueDate = formatDateOfIssue(item.dateOfIssue);

        // Define the formatted date variables
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
          switch (columnType) {
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
              fontSize: 10,
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
              fontSize: 11,
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
              fontSize: 10,
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
              fontSize: 10,
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
              fontSize: 10,
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
              fontSize: 10,
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
              fontSize: 10,
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
          },

          // M/W/K - abbreviated
          {
            content: abbreviateMWKForPDF(item.mwk),
            styles: {
              cellWidth: columnWidths[9],
              fontSize: 10,
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
              fontSize: 10,
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
              fontSize: 10,
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
              fontSize: 10,
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
              fontSize: 10,
              halign: 'center',
              fillColor: rowBgColor,
              fontStyle: formattedEmbPrintDate !== '-' ? 'bold' : 'normal',
              textColor: formattedEmbPrintDate !== '-' ? accentColor : [100, 100, 100],
              cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
            }
          },

          // WIP Status - UPDATED with "Not updated" functionality
          {
            content: wipDisplayValue || 'N/A',
            styles: {
              cellWidth: columnWidths[14],
              fontSize: 10,
              halign: 'center',
              fillColor: !isCompleted && !isUpdatedToday ? [255, 235, 235] : rowBgColor, // Light red background for "Not updated"
              fontStyle: isCompleted ? 'bold' : (wipRemarks !== 'N/A' ? 'bold' : 'normal'),
              textColor: isCompleted
                ? [16, 185, 129]
                : (isUpdatedToday
                  ? (wipRemarks !== 'N/A' ? remarksColor : [100, 100, 100])
                  : notUpdatedColor), // Dark red text for "Not updated"
              cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
            }
          },

          // Pintu (NEW COLUMN)
          {
            content: pintuValue,
            styles: {
              cellWidth: columnWidths[15],
              fontSize: 9,
              halign: 'center',
              fontStyle: 'bold',
              fillColor: rowBgColor,
              textColor: pintuColor,
              cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
            }
          },

          // EA (NEW COLUMN)
          {
            content: eaValue,
            styles: {
              cellWidth: columnWidths[16],
              fontSize: 9,
              halign: 'center',
              fontStyle: 'bold',
              fillColor: rowBgColor,
              textColor: eaColor,
              cellPadding: { top: 2, right: 1, bottom: 2, left: 1 }
            }
          },

          // Total PCS
          {
            content: totalPCS > 0 ? totalPCS.toLocaleString() : 'N/A',
            styles: {
              cellWidth: columnWidths[17],
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
              cellWidth: columnWidths[18],
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
              cellWidth: columnWidths[19],
              fontSize: 7,
              halign: 'center',
              fontStyle: 'bold',
              fillColor: lotStatusColor,
              textColor: lotStatusTextColor,
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
        didDrawPage: function (data) {
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

      exportData.forEach(item => {
        const wipRemarks = getLatestWipRemarks(item.wipStatus);
        const stage = extractStageFromStatus(wipRemarks);

        if (!stageAnalysis[stage]) {
          stageAnalysis[stage] = {
            lots: 0,
            pcs: 0,
            notUpdatedLots: 0
          };
        }

        stageAnalysis[stage].lots += 1;
        stageAnalysis[stage].pcs += item.totalPCS || 0;

        // Check if status not updated
        if (!isStatusUpdatedToday(item.wipStatus, item.completedStatus)) {
          stageAnalysis[stage].notUpdatedLots += 1;
        }
      });

      // Convert to array and sort by lots (descending)
      const stageArray = Object.entries(stageAnalysis)
        .map(([stage, data]) => ({
          stage,
          lots: data.lots,
          pcs: data.pcs,
          notUpdatedLots: data.notUpdatedLots
        }))
        .sort((a, b) => b.lots - a.lots);

      // Add totals row
      const totalStageLots = stageArray.reduce((sum, item) => sum + item.lots, 0);
      const totalStagePCS = stageArray.reduce((sum, item) => sum + item.pcs, 0);
      const totalStageNotUpdated = stageArray.reduce((sum, item) => sum + item.notUpdatedLots, 0);

      // Show stage-wise analysis after main table
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

      // Prepare stage analysis table data
      const stageBody = stageArray.map(item => {
        const percentage = totalStageLots > 0 ? Math.round((item.lots / totalStageLots) * 100) : 0;

        return [
          {
            content: item.stage,
            styles: {
              halign: 'left',
              fontSize: 10,
              cellPadding: { top: 3, right: 3, bottom: 3, left: 6 },
              fontStyle: 'bold',
              fillColor: [240, 249, 255]
            }
          },
          {
            content: item.lots.toString(),
            styles: {
              halign: 'center',
              fontSize: 11,
              fontStyle: 'bold',
              cellPadding: { top: 3, right: 3, bottom: 3, left: 3 }
            }
          },
          {
            content: `${percentage}%`,
            styles: {
              halign: 'center',
              fontSize: 10,
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
              fontSize: 11,
              fontStyle: 'bold',
              cellPadding: { top: 3, right: 3, bottom: 3, left: 3 }
            }
          },
          {
            content: item.notUpdatedLots > 0 ? `${item.notUpdatedLots}` : '-',
            styles: {
              halign: 'center',
              fontSize: 10,
              cellPadding: { top: 3, right: 3, bottom: 3, left: 3 },
              textColor: item.notUpdatedLots > 0 ? [220, 38, 38] : [100, 100, 100],
              fontStyle: item.notUpdatedLots > 0 ? 'bold' : 'normal',
              fillColor: item.notUpdatedLots > 0 ? [255, 235, 235] : [255, 255, 255]
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
            fontSize: 11,
            fontStyle: 'bold',
            fillColor: [225, 239, 255],
            cellPadding: { top: 4, right: 3, bottom: 4, left: 6 }
          }
        },
        {
          content: totalStageLots.toString(),
          styles: {
            halign: 'center',
            fontSize: 12,
            fontStyle: 'bold',
            fillColor: [225, 239, 255],
            cellPadding: { top: 4, right: 3, bottom: 4, left: 3 }
          }
        },
        {
          content: '100%',
          styles: {
            halign: 'center',
            fontSize: 11,
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
            fontSize: 12,
            fontStyle: 'bold',
            fillColor: [225, 239, 255],
            cellPadding: { top: 4, right: 3, bottom: 4, left: 3 }
          }
        },
        {
          content: totalStageNotUpdated > 0 ? `${totalStageNotUpdated}` : '-',
          styles: {
            halign: 'center',
            fontSize: 11,
            fontStyle: 'bold',
            fillColor: [225, 239, 255],
            textColor: totalStageNotUpdated > 0 ? [220, 38, 38] : [100, 100, 100],
            cellPadding: { top: 4, right: 3, bottom: 4, left: 3 }
          }
        }
      ]);

      // Calculate stage column widths (5 columns)
      const stageColumnWidths = [
        contentWidth * 0.45,  // Stage
        contentWidth * 0.10,  // Lots
        contentWidth * 0.08,  // %
        contentWidth * 0.17,  // PCS
        contentWidth * 0.20   // Not Updated
      ];

      // Create stage analysis table
      autoTable(doc, {
        startY: y,
        head: [[
          { content: 'WORK STAGE', styles: { halign: 'center', fontStyle: 'bold', cellWidth: stageColumnWidths[0], fillColor: stageAnalysisColor, textColor: [255, 255, 255] } },
          { content: 'LOTS', styles: { halign: 'center', fontStyle: 'bold', cellWidth: stageColumnWidths[1], fillColor: stageAnalysisColor, textColor: [255, 255, 255] } },
          { content: '%', styles: { halign: 'center', fontStyle: 'bold', cellWidth: stageColumnWidths[2], fillColor: stageAnalysisColor, textColor: [255, 255, 255] } },
          { content: 'TOTAL PCS', styles: { halign: 'center', fontStyle: 'bold', cellWidth: stageColumnWidths[3], fillColor: stageAnalysisColor, textColor: [255, 255, 255] } },
          { content: 'NOT UPDATED', styles: { halign: 'center', fontStyle: 'bold', cellWidth: stageColumnWidths[4], fillColor: [220, 38, 38], textColor: [255, 255, 255] } }
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
          fontSize: 9,
          cellPadding: { top: 5, right: 3, bottom: 5, left: 3 },
          lineWidth: 0.5,
          lineColor: stageAnalysisColor,
          halign: 'center',
          valign: 'middle'
        },
        bodyStyles: {
          fontSize: 10,
          cellPadding: { top: 3, right: 3, bottom: 3, left: 3 },
          lineWidth: 0.3,
          lineColor: [220, 220, 220],
          textColor: textColor,
          font: 'helvetica',
          valign: 'middle'
        },
        columnStyles: {
          0: { cellWidth: stageColumnWidths[0], halign: 'left' },
          1: { cellWidth: stageColumnWidths[1], halign: 'center' },
          2: { cellWidth: stageColumnWidths[2], halign: 'center' },
          3: { cellWidth: stageColumnWidths[3], halign: 'center' },
          4: { cellWidth: stageColumnWidths[4], halign: 'center' }
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

      // ===================== SUPERVISOR WORKLOAD DISTRIBUTION =====================

      // Show supervisor summary for ALL reports
      summaryStartY = y;

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

      y = summaryStartY + 15;

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
            notUpdatedLots: 0
          };
        }

        acc[supervisor].totalLots += 1;
        acc[supervisor].totalPCS += item.totalPCS || 0;

        // Check if status not updated
        if (!isStatusUpdatedToday(item.wipStatus, item.completedStatus)) {
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

      // Prepare simplified summary body with MANPOWER and NOT UPDATED columns
      const summaryBody = sortedSupervisors.map(([supervisor, data]) => {
        const manpower = getManpowerForSupervisor(supervisor);

        // Calculate percentages for each category
        const greenPercent = data.greenLots > 0 ? Math.round((data.greenLots / data.totalLots) * 100) : 0;
        const yellowPercent = data.yellowLots > 0 ? Math.round((data.yellowLots / data.totalLots) * 100) : 0;
        const redPercent = data.redLots > 0 ? Math.round((data.redLots / data.totalLots) * 100) : 0;

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
            content: data.notUpdatedLots > 0 ? data.notUpdatedLots.toString() : '-',
            styles: {
              halign: 'center',
              fontSize: 13,
              cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
              fontStyle: data.notUpdatedLots > 0 ? 'bold' : 'normal',
              textColor: data.notUpdatedLots > 0 ? [220, 38, 38] : [100, 100, 100],
              fillColor: data.notUpdatedLots > 0 ? [255, 235, 235] : [255, 255, 255]
            }
          },
          {
            content: `${data.greenLots} (${greenPercent}%)`,
            styles: {
              halign: 'center',
              fontSize: 13,
              cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
              fontStyle: 'bold',
              textColor: [16, 185, 129]
            }
          },
          {
            content: `${data.yellowLots} (${yellowPercent}%)`,
            styles: {
              halign: 'center',
              fontSize: 13,
              cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
              fontStyle: 'bold',
              textColor: [245, 158, 11]
            }
          },
          {
            content: `${data.redLots} (${redPercent}%)`,
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

        // Check if status not updated
        if (!isStatusUpdatedToday(item.wipStatus, item.completedStatus)) {
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
        notUpdatedLots: 0
      });

      // Add total row
      const totalRowBackground = [225, 239, 255];
      const totalManpower = sortedSupervisors.reduce((sum, [supervisorName]) => {
        const manpower = parseInt(getManpowerForSupervisor(supervisorName)) || 0;
        return sum + manpower;
      }, 0);
      // Calculate overall percentages
      const overallGreenPercent = overallTotals.greenLots > 0 ? Math.round((overallTotals.greenLots / overallTotals.totalLots) * 100) : 0;
      const overallYellowPercent = overallTotals.yellowLots > 0 ? Math.round((overallTotals.yellowLots / overallTotals.totalLots) * 100) : 0;
      const overallRedPercent = overallTotals.redLots > 0 ? Math.round((overallTotals.redLots / overallTotals.totalLots) * 100) : 0;

      summaryBody.push([
        {
          content: 'OVERALL TOTALS',
          styles: {
            halign: 'left',
            fontSize: 11,
            fontStyle: 'bold',
            fillColor: totalRowBackground,
            cellPadding: { top: 6, right: 3, bottom: 6, left: 6 }
          }
        },
        {
          content: totalManpower > 0 ? totalManpower.toString() : '-',
          styles: {
            halign: 'center',
            fontSize: 11,
            fontStyle: 'bold',
            fillColor: totalRowBackground,
            cellPadding: { top: 6, right: 3, bottom: 6, left: 3 }
          }
        },
        {
          content: overallTotals.totalLots.toString(),
          styles: {
            halign: 'center',
            fontSize: 11,
            fontStyle: 'bold',
            fillColor: totalRowBackground,
            cellPadding: { top: 6, right: 3, bottom: 6, left: 3 }
          }
        },
        {
          content: overallTotals.totalPCS.toLocaleString(),
          styles: {
            halign: 'center',
            fontSize: 11,
            fontStyle: 'bold',
            fillColor: totalRowBackground,
            cellPadding: { top: 6, right: 3, bottom: 6, left: 3 }
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
            cellPadding: { top: 6, right: 3, bottom: 6, left: 3 }
          }
        },
        {
          content: `${overallTotals.greenLots} (${overallGreenPercent}%)`,
          styles: {
            halign: 'center',
            fontSize: 11,
            fontStyle: 'bold',
            fillColor: totalRowBackground,
            textColor: [16, 185, 129],
            cellPadding: { top: 6, right: 3, bottom: 6, left: 3 }
          }
        },
        {
          content: `${overallTotals.yellowLots} (${overallYellowPercent}%)`,
          styles: {
            halign: 'center',
            fontSize: 11,
            fontStyle: 'bold',
            fillColor: totalRowBackground,
            textColor: [245, 158, 11],
            cellPadding: { top: 6, right: 3, bottom: 6, left: 3 }
          }
        },
        {
          content: `${overallTotals.redLots} (${overallRedPercent}%)`,
          styles: {
            halign: 'center',
            fontSize: 11,
            fontStyle: 'bold',
            fillColor: totalRowBackground,
            textColor: [239, 68, 68],
            cellPadding: { top: 6, right: 3, bottom: 6, left: 3 }
          }
        }
      ]);

      // Calculate column widths for supervisor summary table (8 columns now)
      const totalAvailableWidth = contentWidth;
      const numberOfColumns = 8;
      const baseColumnWidth = totalAvailableWidth / numberOfColumns;

      const summaryColumnWidths = [
        baseColumnWidth * 1.2,  // Supervisor
        baseColumnWidth * 0.8,  // Manpower
        baseColumnWidth * 0.9,  // Total Lots
        baseColumnWidth * 1.0,  // Total PCS
        baseColumnWidth * 1.1,  // Not Updated
        baseColumnWidth * 1.0,  // Green
        baseColumnWidth * 1.0,  // Yellow
        baseColumnWidth * 1.0,  // Red
      ];

      // Create supervisor summary table
      autoTable(doc, {
        startY: y,
        head: [[
          { content: 'SUPERVISOR', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[0] } },
          { content: 'MANPOWER', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[1], fillColor: [15, 76, 129], textColor: [255, 255, 255] } },
          { content: 'TOTAL LOTS', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[2] } },
          { content: 'TOTAL PCS', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[3] } },
          { content: 'NOT UPDATED', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[4], fillColor: [220, 38, 38], textColor: [255, 255, 255] } },
          { content: 'GREEN (1-6 Days)', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[5], textColor: [16, 185, 129] } },
          { content: 'YELLOW (7-15 Days)', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[6], textColor: [245, 158, 11] } },
          { content: 'RED (15+ Days)', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[7], textColor: [239, 68, 68] } }
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
          fontSize: 9,
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
          6: { cellWidth: summaryColumnWidths[6], halign: 'center' },
          7: { cellWidth: summaryColumnWidths[7], halign: 'center' }
        },
        margin: { left: margin, right: margin },
        tableWidth: 'auto',
        showHead: 'everyPage',
        showFoot: false,
        pageBreak: 'auto',
        rowPageBreak: 'avoid'
      });

      // Add summary notes with the new "Not updated" note
      const summaryEndY = doc.lastAutoTable.finalY;
      const notesY = summaryEndY + 8;

      doc.setFontSize(8);
      doc.setFont('helvetica', 'italic');
      doc.setTextColor(0, 0, 0);

      const notes = [

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
      if (filters.supervisor && (Array.isArray(filters.supervisor) ? filters.supervisor.length > 0 : (typeof filters.supervisor === 'string' && filters.supervisor.trim() !== ''))) {
        const supervisorName = Array.isArray(filters.supervisor)
          ? filters.supervisor.join('_')
          : filters.supervisor.replace(/\s+/g, '_');
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
  }, [filteredData, filters, calculateStitchingDays, getStitchingDaysColor, getStitchingDaysTextColor, getEmbPrintDate, getLatestWipRemarks, getCompletionDateFormatted, isLotCompleted, getLatestManpowerBySupervisor]);
  const handleRefresh = useCallback(async () => {
    try {
      setLoading(true);
      setError(null);

      // Refresh working updates data for PDF
      await fetchWorkingUpdatesData();

      // Clear cache for current supervisor
      const cacheKey = getCacheKey(filters.supervisor);
      dataCache.current.delete(cacheKey);
      lastFetchTime.current.delete(cacheKey);

      const freshData = await fetchDataForSupervisor(filters.supervisor);
      setData(freshData);

      applyFilters(filters, freshData);
      extractFilterOptions(freshData);

    } catch (err) {
      setError(`Failed to refresh data: ${err.message}`);
    } finally {
      setLoading(false);
    }
  }, [filters, fetchDataForSupervisor, fetchWorkingUpdatesData, applyFilters, extractFilterOptions, getCacheKey]);



  // Memoize render functions with MultiSelectDropdown
  const renderDropdownFilter = useCallback((fieldName, label) => (
    <div key={fieldName} className="filter-item">
      <label className="filter-label">
        {label}
      </label>
      <MultiSelectDropdown
        options={filterOptions[fieldName] || []}
        selectedValues={Array.isArray(filters[fieldName]) ? filters[fieldName] : (filters[fieldName] ? [filters[fieldName]] : [])}
        onChange={(selected) => {
          const newFilters = { ...filters, [fieldName]: selected };
          setFilters(newFilters);
          setTimeout(() => {
            applyFilters(newFilters, data);
          }, 0);
        }}
        placeholder={`All ${label}`}
        disabled={loading}
      />
    </div>
  ), [filters, filterOptions, data, applyFilters, loading]);

  // Loading and error states
  if (isInitialLoad && loading) {
    return (
      <div style={{
        display: "flex",
        justifyContent: "center",
        alignItems: "center",
        minHeight: "100vh",
        backgroundColor: "#f8fafc",
        backgroundImage: "radial-gradient(at 0% 0%, rgba(99, 102, 241, 0.15) 0px, transparent 50%), radial-gradient(at 100% 100%, rgba(236, 72, 153, 0.12) 0px, transparent 50%)",
        padding: "24px",
        fontFamily: "'Plus Jakarta Sans', 'Inter', system-ui, sans-serif"
      }}>
        <style>{`
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          @keyframes pulse {
            0% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(99, 102, 241, 0.7); }
            70% { transform: scale(1); box-shadow: 0 0 0 8px rgba(99, 102, 241, 0); }
            100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(99, 102, 241, 0); }
          }
        `}</style>
        <div style={{
          background: "rgba(255, 255, 255, 0.95)",
          backdropFilter: "blur(16px)",
          WebkitBackdropFilter: "blur(16px)",
          border: "1px solid rgba(99, 102, 241, 0.2)",
          padding: "44px 40px",
          borderRadius: "24px",
          boxShadow: "0 25px 50px -12px rgba(30, 27, 75, 0.18)",
          display: "flex",
          flexDirection: "column",
          alignItems: "center",
          maxWidth: "440px",
          width: "100%",
          textAlign: "center"
        }}>
          <div style={{
            position: "relative",
            width: "72px",
            height: "72px",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            marginBottom: "24px"
          }}>
            <div style={{
              width: "100%",
              height: "100%",
              borderRadius: "50%",
              background: "conic-gradient(from 0deg, #6366f1, #ec4899, #10b981, #6366f1)",
              animation: "spin 1.2s linear infinite",
              padding: "4px",
              WebkitMask: "radial-gradient(farthest-side, transparent calc(100% - 4px), #fff 0)",
              mask: "radial-gradient(farthest-side, transparent calc(100% - 4px), #fff 0)"
            }}></div>
            <div style={{ position: "absolute", fontSize: "24px" }}>📊</div>
          </div>
          <h2 style={{ fontSize: "1.4rem", fontWeight: "800", color: "#1e1b4b", margin: "0 0 8px 0" }}>
            Loading Stitching Reports
          </h2>
          <p style={{ fontSize: "0.9rem", color: "#64748b", margin: "0 0 20px 0", lineHeight: "1.5" }}>
            Syncing Google Sheets data & calculating lot metrics...
          </p>
          <div style={{
            display: "inline-flex",
            alignItems: "center",
            gap: "8px",
            background: "#e0e7ff",
            color: "#3730a3",
            padding: "8px 18px",
            borderRadius: "20px",
            fontSize: "0.82rem",
            fontWeight: "700",
            border: "1px solid #c7d2fe"
          }}>
            <span style={{
              width: "8px",
              height: "8px",
              borderRadius: "50%",
              background: "#6366f1",
              display: "inline-block",
              animation: "pulse 1.5s infinite"
            }}></span>
            Sorting completed lots in descending order
          </div>
        </div>
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
              Overall Stitching Completed Report
            </h1>
            <p className="subtitle">
              {Array.isArray(filters.supervisor) && filters.supervisor.length > 0
                ? `${filters.supervisor.join(', ')} - Stitching Report`
                : typeof filters.supervisor === 'string' && filters.supervisor
                  ? `${filters.supervisor} - Stitching Report`
                  : 'Real-time tracking of stitching operations'}
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
            <span className="active-filters">
              {Object.entries(filters).filter(([key, value]) => {
                if (key === 'dateRange') {
                  const { from, to } = value;
                  return from || to;
                }
                if (Array.isArray(value)) {
                  return value.length > 0;
                }
                return typeof value === 'string' && value.trim() !== '';
              }).length} active filters
            </span>
            <button
              onClick={clearFilters}
              className="btn-clear"
              disabled={loading}
            >
              Clear Filters
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
                <MultiSelectDropdown
                  options={filterOptions.supervisor}
                  selectedValues={Array.isArray(filters.supervisor) ? filters.supervisor : (filters.supervisor ? [filters.supervisor] : [])}
                  onChange={handleSupervisorChange}
                  placeholder="All Supervisors"
                  disabled={loading}
                />
                <div className="filter-hint">
                  Select multiple supervisors
                </div>
              </div>

              <div className="filter-item">
                <label>Party Name</label>
                <MultiSelectDropdown
                  options={filterOptions.partyName || []}
                  selectedValues={Array.isArray(filters.partyName) ? filters.partyName : (filters.partyName ? [filters.partyName] : [])}
                  onChange={(selected) => {
                    const newFilters = { ...filters, partyName: selected };
                    setFilters(newFilters);
                    setTimeout(() => {
                      applyFilters(newFilters, data);
                    }, 0);
                  }}
                  placeholder="All Parties"
                  disabled={loading}
                />
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
                <MultiSelectDropdown
                  options={['Yes', 'No']}
                  selectedValues={Array.isArray(filters.directStitching) ? filters.directStitching : (filters.directStitching ? [filters.directStitching] : [])}
                  onChange={(selected) => {
                    const newFilters = { ...filters, directStitching: selected };
                    setFilters(newFilters);
                    setTimeout(() => {
                      applyFilters(newFilters, data);
                    }, 0);
                  }}
                  placeholder="All"
                  disabled={loading}
                />
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
            <div className="spinner-wrapper small">
              <div className="gradient-spinner"></div>
            </div>
            <p className="overlay-text">Updating {filters.supervisor ? filters.supervisor : 'All Supervisors'}'s data...</p>
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
                  {filteredData.slice(0, displayLimit).map((item, index) => {
                    const isCompleted = isLotCompleted(item.completedStatus);
                    const stitchingDays = calculateStitchingDays(item.dateOfIssue, item.completedStatus, isCompleted);
                    const stitchingDaysColor = getStitchingDaysColor(stitchingDays);
                    const stitchingDaysTextColor = getStitchingDaysTextColor(stitchingDays);
                    const embPrintDate = getEmbPrintDate(item.challanHistory);
                    // const isCompleted = isLotCompleted(item.completedStatus);
                    const completionDate = getCompletionDateFormatted(item.completedStatus);

                    return (
                      <tr key={index} className="table-row">
                        <td className="text-center font-semibold">{index + 1}</td>
                        <td className="text-center" style={{ verticalAlign: 'middle' }}>
                          {item.image ? (
                            <img
                              src={item.image}
                              alt="Style Preview"
                              style={{
                                width: "42px",
                                height: "42px",
                                objectFit: "cover",
                                borderRadius: "6px",
                                border: "1px solid #cbd5e1",
                                cursor: "pointer",
                                boxShadow: "0 2px 4px rgba(0,0,0,0.05)"
                              }}
                              referrerPolicy="no-referrer"
                              onClick={(e) => {
                                e.stopPropagation();
                                setViewImageSrc(item.image);
                              }}
                            />
                          ) : (
                            <span style={{ fontSize: "9px", color: "#94a3b8", textTransform: "uppercase" }}>No Image</span>
                          )}
                        </td>
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

                        {/* WIP Status - ALWAYS INCLUDED */}
                        {/* WIP Status - ALWAYS INCLUDED, with "Done" for completed lots */}
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
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            {/* Load More Pagination / Batching Control for Performance */}
            {filteredData.length > displayLimit && (
              <div style={{ padding: '16px 24px', background: '#f8fafc', borderTop: '1px solid #e2e8f0', display: 'flex', justifyContent: 'center', alignItems: 'center', gap: '12px' }}>
                <button
                  onClick={() => setDisplayLimit(prev => Math.min(prev + 300, filteredData.length))}
                  className="btn btn-primary btn-small"
                  style={{ background: 'linear-gradient(135deg, #4f46e5 0%, #4338ca 100%)', color: '#ffffff', border: 'none', padding: '8px 20px', borderRadius: '10px', fontWeight: '700' }}
                >
                  📥 Load More (Showing {displayLimit} of {filteredData.length} records)
                </button>
                <button
                  onClick={() => setDisplayLimit(filteredData.length)}
                  className="btn btn-small"
                  style={{ background: '#ffffff', color: '#475569', border: '1px solid #cbd5e1', padding: '8px 16px', borderRadius: '10px', fontWeight: '600' }}
                >
                  ⚡ Show All ({filteredData.length} records)
                </button>
              </div>
            )}

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

      {/* Image Preview Modal */}
      {viewImageSrc && (
        <div
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: "rgba(15, 23, 42, 0.8)",
            zIndex: 99999,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            backdropFilter: "blur(6px)"
          }}
          onClick={() => setViewImageSrc(null)}
        >
          <div
            style={{
              position: "relative",
              maxWidth: "90%",
              maxHeight: "90%",
              backgroundColor: "white",
              borderRadius: "16px",
              padding: "16px",
              boxShadow: "0 25px 50px -12px rgba(0, 0, 0, 0.5)"
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <button
              style={{
                position: "absolute",
                top: "-12px",
                right: "-12px",
                backgroundColor: "#0f172a",
                color: "white",
                border: "none",
                width: "32px",
                height: "32px",
                borderRadius: "50%",
                fontSize: "16px",
                fontWeight: "bold",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                cursor: "pointer",
                boxShadow: "0 4px 6px rgba(0, 0, 0, 0.2)"
              }}
              onClick={() => setViewImageSrc(null)}
            >
              ✕
            </button>
            <img
              src={viewImageSrc}
              alt="Enlarged style preview"
              style={{
                maxWidth: "100%",
                maxHeight: "80vh",
                borderRadius: "10px",
                objectFit: "contain"
              }}
            />
          </div>
        </div>
      )}

      <style jsx>{`
        .container {
          min-height: 100vh;
          background-color: #f8fafc;
          background-image: 
            radial-gradient(at 0% 0%, rgba(99, 102, 241, 0.08) 0px, transparent 50%), 
            radial-gradient(at 100% 0%, rgba(236, 72, 153, 0.06) 0px, transparent 50%), 
            radial-gradient(at 50% 100%, rgba(16, 185, 129, 0.06) 0px, transparent 50%);
          padding: 32px;
          font-family: 'Plus Jakarta Sans', 'Inter', system-ui, sans-serif;
          color: #0f172a;
          position: relative;
        }

        .loading-container {
          display: flex;
          justify-content: center;
          align-items: center;
          min-height: 100vh;
          background-color: #f8fafc;
          background-image: 
            radial-gradient(at 0% 0%, rgba(99, 102, 241, 0.12) 0px, transparent 50%), 
            radial-gradient(at 100% 100%, rgba(236, 72, 153, 0.1) 0px, transparent 50%);
          padding: 24px;
        }

        .loading-card {
          background: rgba(255, 255, 255, 0.9);
          backdrop-filter: blur(16px);
          border: 1px solid rgba(99, 102, 241, 0.2);
          padding: 48px 40px;
          border-radius: 24px;
          box-shadow: 0 25px 50px -12px rgba(30, 27, 75, 0.15);
          display: flex;
          flex-direction: column;
          align-items: center;
          max-width: 440px;
          width: 100%;
          text-align: center;
        }

        .spinner-wrapper {
          position: relative;
          width: 72px;
          height: 72px;
          display: flex;
          align-items: center;
          justify-content: center;
          margin-bottom: 24px;
        }

        .spinner-wrapper.small {
          width: 44px;
          height: 44px;
          margin-bottom: 12px;
        }

        .gradient-spinner {
          width: 100%;
          height: 100%;
          border-radius: 50%;
          background: conic-gradient(from 0deg, #6366f1, #ec4899, #10b981, #6366f1);
          animation: spin 1.2s linear infinite;
          padding: 4px;
          mask: radial-gradient(farthest-side, transparent calc(100% - 4px), #fff 0);
          -webkit-mask: radial-gradient(farthest-side, transparent calc(100% - 4px), #fff 0);
        }

        .spinner-center-icon {
          position: absolute;
          font-size: 24px;
        }

        .loading-title {
          font-size: 1.4rem;
          font-weight: 800;
          color: #1e1b4b;
          margin: 0 0 8px 0;
        }

        .loading-subtext {
          font-size: 0.9rem;
          color: #64748b;
          margin: 0 0 20px 0;
          line-height: 1.5;
        }

        .loading-badge {
          display: inline-flex;
          align-items: center;
          gap: 8px;
          background: #e0e7ff;
          color: #3730a3;
          padding: 6px 16px;
          border-radius: 20px;
          font-size: 0.8rem;
          font-weight: 700;
          border: 1px solid #c7d2fe;
        }

        .pulse-dot {
          width: 8px;
          height: 8px;
          border-radius: 50%;
          background: #6366f1;
          box-shadow: 0 0 0 0 rgba(99, 102, 241, 0.7);
          animation: pulse 1.5s infinite;
        }

        @keyframes pulse {
          0% {
            transform: scale(0.95);
            box-shadow: 0 0 0 0 rgba(99, 102, 241, 0.7);
          }
          70% {
            transform: scale(1);
            box-shadow: 0 0 0 8px rgba(99, 102, 241, 0);
          }
          100% {
            transform: scale(0.95);
            box-shadow: 0 0 0 0 rgba(99, 102, 241, 0);
          }
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
          background: linear-gradient(135deg, #1e1b4b 0%, #312e81 40%, #4338ca 100%);
          border-radius: 24px;
          padding: 32px 36px;
          margin-bottom: 28px;
          box-shadow: 0 20px 40px -10px rgba(49, 46, 129, 0.3);
          color: white;
        }

        .header-top {
          display: flex;
          justify-content: space-between;
          align-items: flex-start;
          margin-bottom: 24px;
        }

        .title-section h1 {
          font-size: 2.4rem;
          font-weight: 800;
          color: #ffffff;
          margin: 0;
          display: flex;
          align-items: center;
          gap: 12px;
        }

        .title-icon {
          font-size: 32px;
        }

        .subtitle {
          font-size: 1rem;
          color: #c7d2fe;
          margin-top: 8px;
          display: flex;
          align-items: center;
          gap: 8px;
        }

        .data-source {
          font-size: 12px;
          background: rgba(255, 255, 255, 0.15);
          color: #e0e7ff;
          padding: 2px 8px;
          border-radius: 12px;
          border: 1px solid rgba(255, 255, 255, 0.2);
        }

        .stats-card {
          display: flex;
          gap: 24px;
          background: rgba(255, 255, 255, 0.12);
          border: 1px solid rgba(255, 255, 255, 0.2);
          padding: 16px 24px;
          border-radius: 20px;
          backdrop-filter: blur(12px);
        }

        .stat-item {
          text-align: center;
          min-width: 80px;
        }

        .stat-value {
          font-size: 2rem;
          font-weight: 800;
          color: #ffffff;
          line-height: 1.2;
        }

        .stat-label {
          font-size: 0.75rem;
          color: #c7d2fe;
          margin-top: 4px;
          font-weight: 700;
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
          background: linear-gradient(135deg, #475569 0%, #334155 100%);
          color: white;
          box-shadow: 0 4px 12px rgba(51, 65, 85, 0.25);
        }

        .btn-back:hover:not(:disabled) {
          background: linear-gradient(135deg, #334155 0%, #1e293b 100%);
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
          background: rgba(15, 23, 42, 0.4);
          backdrop-filter: blur(6px);
          display: flex;
          justify-content: center;
          align-items: center;
          z-index: 1000;
          border-radius: 20px;
        }

        .loading-content {
          text-align: center;
          background: rgba(255, 255, 255, 0.95);
          padding: 24px 36px;
          border-radius: 20px;
          box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1);
          border: 1px solid #e2e8f0;
          display: flex;
          flex-direction: column;
          align-items: center;
        }

        .overlay-text {
          font-weight: 700;
          color: #1e1b4b;
          margin: 0;
          font-size: 0.95rem;
        }

        /* Filters Section */
        .filters-section {
          background: #ffffff;
          border-radius: 20px;
          padding: 24px;
          margin-bottom: 28px;
          box-shadow: 0 10px 30px rgba(0, 0, 0, 0.03);
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
          background: #ffffff;
          border-radius: 20px;
          overflow: hidden;
          box-shadow: 0 10px 30px rgba(0, 0, 0, 0.03);
          border: 1px solid #e2e8f0;
        }

        .table-header {
          padding: 20px 24px;
          background: linear-gradient(135deg, #1e1b4b 0%, #312e81 100%);
          border-bottom: 1px solid #e2e8f0;
          display: flex;
          justify-content: space-between;
          align-items: center;
        }

        .table-info {
          font-size: 14px;
          color: #ffffff;
        }

        .table-info strong {
          color: #ffffff;
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
          color: #ffffff;
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
          background: linear-gradient(135deg, #1e1b4b 0%, #312e81 100%);
        }

        .data-table th {
          padding: 14px 16px;
          text-align: left;
          font-size: 0.82rem;
          font-weight: 700;
          color: #ffffff;
          text-transform: uppercase;
          letter-spacing: 0.04em;
          border: 1px solid rgba(255, 255, 255, 0.1);
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

export default StitchingCompleteLot;