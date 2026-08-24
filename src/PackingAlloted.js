import React, { useState, useEffect, useMemo, useCallback } from 'react';
import './PendingPackingtoIssue.css';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { GOOGLE_API_KEY, SPREADSHEET_IDS, SHEET_NAMES, fetchSheetDataFromBackend } from './config';

const MultiSelectDropdown = ({ label, options, selectedValues, onChange }) => {
  const [isOpen, setIsOpen] = useState(false);
  const dropdownRef = React.useRef(null);

  useEffect(() => {
    const handleClickOutside = (event) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target)) {
        setIsOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  const toggleOption = (val) => {
    if (val === "all") {
      onChange([]);
      return;
    }
    if (selectedValues.includes(val)) {
      onChange(selectedValues.filter((v) => v !== val));
    } else {
      onChange([...selectedValues, val]);
    }
  };

  const isAllSelected = selectedValues.length === 0;

  const getDisplayText = () => {
    if (isAllSelected) return "All";
    if (selectedValues.length === 1) return selectedValues[0];
    return `${selectedValues.length} Selected`;
  };

  return (
    <div className="filter-item" ref={dropdownRef} style={{ position: "relative" }}>
      <label>{label}</label>
      <div
        onClick={() => setIsOpen(!isOpen)}
        className="filter-select"
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          userSelect: "none",
          background: selectedValues.length > 0 ? "#e0e7ff" : "#ffffff",
          borderColor: selectedValues.length > 0 ? "#6366f1" : "#cbd5e1",
          fontWeight: selectedValues.length > 0 ? "700" : "600",
          color: selectedValues.length > 0 ? "#3730a3" : "#0f172a"
        }}
      >
        <span>{getDisplayText()}</span>
        <span style={{ fontSize: "0.75rem", transform: isOpen ? "rotate(180deg)" : "rotate(0deg)", transition: "transform 0.2s" }}>
          ▼
        </span>
      </div>

      {isOpen && (
        <div
          style={{
            position: "absolute",
            top: "calc(100% + 4px)",
            left: 0,
            right: 0,
            zIndex: 100,
            background: "#ffffff",
            border: "1.5px solid #cbd5e1",
            borderRadius: "14px",
            boxShadow: "0 10px 25px rgba(0, 0, 0, 0.15)",
            maxHeight: "220px",
            overflowY: "auto",
            padding: "8px"
          }}
        >
          {options.map((opt) => {
            const isChecked = opt.value === "all" ? isAllSelected : selectedValues.includes(opt.value);
            return (
              <label
                key={opt.value}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: "8px",
                  padding: "8px 10px",
                  borderRadius: "8px",
                  cursor: "pointer",
                  fontSize: "0.85rem",
                  fontWeight: isChecked ? "700" : "500",
                  color: isChecked ? "#4338ca" : "#334155",
                  background: isChecked ? "#e0e7ff" : "transparent",
                  transition: "background 0.15s ease"
                }}
              >
                <input
                  type="checkbox"
                  checked={isChecked}
                  onChange={() => toggleOption(opt.value)}
                  style={{ cursor: "pointer", accentColor: "#4338ca" }}
                />
                <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                  {opt.label}
                </span>
              </label>
            );
          })}
        </div>
      )}
    </div>
  );
};

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

const getBase64ImageFromUrl = async (imageUrl) => {
  if (!imageUrl) return null;
  const trimmed = imageUrl.toString().trim();
  if (!trimmed) return null;

  let driveId = null;
  const match = trimmed.match(/\/d\/([a-zA-Z0-9_-]+)/) || trimmed.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (match && match[1]) {
    driveId = match[1];
  }

  const fetchAsBase64 = async (targetUrl) => {
    try {
      const res = await fetch(targetUrl);
      if (!res.ok) return null;
      const blob = await res.blob();
      return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result);
        reader.onerror = () => resolve(null);
        reader.readAsDataURL(blob);
      });
    } catch {
      return null;
    }
  };

  const loadViaCanvas = (targetUrl) => {
    return new Promise((resolve) => {
      const img = new Image();
      img.crossOrigin = 'anonymous';
      img.onload = () => {
        try {
          const canvas = document.createElement('canvas');
          canvas.width = img.naturalWidth || 80;
          canvas.height = img.naturalHeight || 80;
          const ctx = canvas.getContext('2d');
          ctx.drawImage(img, 0, 0);
          resolve(canvas.toDataURL('image/jpeg'));
        } catch {
          resolve(null);
        }
      };
      img.onerror = () => resolve(null);
      img.src = targetUrl;
    });
  };

  const candidates = [];
  if (driveId) {
    candidates.push(`https://images.weserv.nl/?url=${encodeURIComponent(`lh3.googleusercontent.com/d/${driveId}=s400`)}`);
    candidates.push(`https://images.weserv.nl/?url=${encodeURIComponent(`drive.google.com/thumbnail?id=${driveId}&sz=w400`)}`);
    candidates.push(`https://lh3.googleusercontent.com/d/${driveId}=s400`);
    candidates.push(`https://drive.google.com/thumbnail?id=${driveId}&sz=w400`);
  } else {
    candidates.push(`https://images.weserv.nl/?url=${encodeURIComponent(trimmed.replace(/^https?:\/\//, ''))}`);
    candidates.push(trimmed);
  }

  for (const url of candidates) {
    const res = await fetchAsBase64(url);
    if (res && res.startsWith('data:image')) return res;
  }

  const fallbackUrl = driveId ? `https://images.weserv.nl/?url=${encodeURIComponent(`lh3.googleusercontent.com/d/${driveId}=s400`)}` : trimmed;
  return loadViaCanvas(fallbackUrl);
};

/* ================== FINANCIAL YEAR HELPERS ================== */
const getFinancialYearFromDate = (dateStr) => {
  if (!dateStr) return null;
  let d = null;
  if (dateStr instanceof Date) {
    d = dateStr;
  } else if (typeof dateStr === "string") {
    const s = dateStr.trim();
    if (!s || s === "-") return null;
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
      const parts = s.split("-");
      d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    } else if (/^\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{4}/.test(s)) {
      const parts = s.split(/[\/\-\.]/);
      d = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
    } else {
      d = new Date(s);
    }
  }
  if (!d || isNaN(d.getTime())) return null;
  const month = d.getMonth();
  const year = d.getFullYear();
  return month >= 3 ? `${year}-${year + 1}` : `${year - 1}-${year}`;
};

const getCurrentFinancialYear = () => {
  return getFinancialYearFromDate(new Date());
};

const PackingAlloted = () => {
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [packingData, setPackingData] = useState([]);
  const [issuesData, setIssuesData] = useState([]);
  const [issuesLotMap, setIssuesLotMap] = useState(new Map());
  const [searchTerm, setSearchTerm] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const [itemsPerPage] = useState(50);
  const [exportLoading, setExportLoading] = useState(false);
  const [showFilters, setShowFilters] = useState(false);
  const [viewImageSrc, setViewImageSrc] = useState(null);
  const [financialYearFilter, setFinancialYearFilter] = useState(getCurrentFinancialYear());
  const [rawpackData, setRawpackData] = useState([]);
  const [rawpackLotMap, setRawpackLotMap] = useState(new Map());

  // Filter states - array of selected values for multi-selection
  const [filters, setFilters] = useState({
    priority: [],
    directStitching: [],
    stitchingSupervisor: [],
    packingSupervisor: [],
    brand: [],
    season: [],
    pendingDaysRange: [],
    partyName: []
  });

  // Google Sheets configuration from .env / config
  const API_KEY = GOOGLE_API_KEY;
  const SPREADSHEET_ID = SPREADSHEET_IDS.MAIN;
  const RANGE = `${SHEET_NAMES.INDEX}!A:AA`;

  // Issues spreadsheet configuration from .env / config
  const ISSUES_SPREADSHEET_ID = SPREADSHEET_IDS.ISSUES;
  const ISSUES_SPREADSHEET_RANGE = `${SHEET_NAMES.ISSUES}!A:R`;

  // RAWPACK spreadsheet configuration from .env / config
  const RAWPACK_SPREADSHEET_ID = SPREADSHEET_IDS.RAWPACK || '1xD8Uy1lUgvNTQ2RGRBI4ZjOrozbinUPRq2_UfIplP98';
  const RAWPACK_SPREADSHEET_RANGE = `${SHEET_NAMES.RAWPACK || 'RAWPACK'}!A:ZZ`;

  // Pending days ranges for filtering
  const pendingDaysRanges = [
    { value: 'all', label: 'All Days' },
    { value: '0-7', label: '0-7 Days' },
    { value: '8-15', label: '8-15 Days' },
    { value: '16-30', label: '16-30 Days' },
    { value: '30+', label: '30+ Days' }
  ];

  // Function to handle back navigation using window.history
  const handleGoBack = () => {
    window.history.back();
  };

  // Function to parse date from various formats
  const parseDate = (dateString) => {
    if (!dateString || dateString === '-') return null;

    try {
      if (typeof dateString === 'string' && dateString.includes('/')) {
        const parts = dateString.split('/');
        if (parts.length === 3) {
          const day = parseInt(parts[0], 10);
          const month = parseInt(parts[1], 10) - 1;
          const year = parseInt(parts[2], 10);

          if (!isNaN(day) && !isNaN(month) && !isNaN(year) &&
            day > 0 && day <= 31 && month >= 0 && month < 12 && year > 1900) {
            return new Date(year, month, day);
          }
        }
      }

      // Handle YYYY-MM-DD format
      if (typeof dateString === 'string' && dateString.includes('-')) {
        const parts = dateString.split('-');
        if (parts.length === 3) {
          const year = parseInt(parts[0], 10);
          const month = parseInt(parts[1], 10) - 1;
          const day = parseInt(parts[2], 10);

          if (!isNaN(day) && !isNaN(month) && !isNaN(year) &&
            day > 0 && day <= 31 && month >= 0 && month < 12 && year > 1900) {
            return new Date(year, month, day);
          }
        }
      }

      const date = new Date(dateString);
      if (!isNaN(date.getTime())) {
        return date;
      }

      return null;
    } catch {
      return null;
    }
  };

  // Function to check if packing is complete (has a date value)
  const isPackingComplete = (packingCompleteValue) => {
    if (!packingCompleteValue || packingCompleteValue === '' || packingCompleteValue === '-') {
      return false;
    }

    // If it's a valid date string (with numbers), consider it complete
    if (typeof packingCompleteValue === 'string') {
      // Check if it contains numbers (indicating a date)
      const hasNumbers = /\d/.test(packingCompleteValue);
      if (hasNumbers) {
        return true;
      }

      // Check for 'yes' case insensitive
      if (packingCompleteValue.toLowerCase() === 'yes') {
        return true;
      }
    }

    return false;
  };

  // Function to calculate pending days for packing
  const calculatePackingPendingDays = (packingAllocatedDate) => {
    if (!packingAllocatedDate || packingAllocatedDate === '-') return 0;

    try {
      const allocated = parseDate(packingAllocatedDate);
      if (!allocated) return 0;

      const today = new Date();
      today.setHours(0, 0, 0, 0);
      allocated.setHours(0, 0, 0, 0);

      const diffTime = today - allocated;
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

      return diffDays > 0 ? diffDays : 0;
    } catch (error) {
      console.error('Error calculating packing pending days:', error);
      return 0;
    }
  };

  // Function to parse completed status JSON and extract date
  const parseCompletedStatus = (statusData) => {
    if (!statusData) return { displayDate: '-', rawStatus: null };

    try {
      if (typeof statusData === 'object') {
        return {
          displayDate: extractDateFromTimestamp(statusData.timestamp),
          rawStatus: statusData
        };
      }

      const parsed = JSON.parse(statusData);
      if (Array.isArray(parsed) && parsed.length > 0) {
        const status = parsed[0];
        return {
          displayDate: extractDateFromTimestamp(status.timestamp),
          rawStatus: status
        };
      }
    } catch (e) {
      console.log('Failed to parse completed status:', statusData);
    }

    return { displayDate: statusData || '-', rawStatus: null };
  };

  const extractDateFromTimestamp = (timestamp) => {
    if (!timestamp) return '-';

    try {
      const date = new Date(timestamp);
      if (isNaN(date.getTime())) return timestamp;

      const day = date.getDate().toString().padStart(2, '0');
      const month = (date.getMonth() + 1).toString().padStart(2, '0');
      const year = date.getFullYear();

      return `${day}/${month}/${year}`;
    } catch {
      return timestamp;
    }
  };

  // Format date for display (always DD/MM/YYYY)
  const formatDate = (dateString) => {
    if (!dateString || dateString === '-') return '-';

    try {
      // Handle YYYY-MM-DD format (from Issues sheet)
      if (typeof dateString === 'string' && dateString.includes('-')) {
        const parts = dateString.split('-');
        if (parts.length === 3) {
          const year = parts[0];
          const month = parts[1];
          const day = parts[2];

          // Check if it's a valid date format with numbers
          if (/^\d{4}$/.test(year) && /^\d{2}$/.test(month) && /^\d{2}$/.test(day)) {
            return `${day}/${month}/${year}`;
          }
        }
      }

      // Handle DD/MM/YYYY format
      if (typeof dateString === 'string' && dateString.includes('/')) {
        const parts = dateString.split('/');
        if (parts.length === 3) {
          const day = parseInt(parts[0], 10);
          const month = parseInt(parts[1], 10);
          const year = parseInt(parts[2], 10);

          if (!isNaN(day) && !isNaN(month) && !isNaN(year) &&
            day > 0 && day <= 31 && month > 0 && month <= 12 && year > 1900) {
            return dateString;
          }
        }
      }

      const date = parseDate(dateString);
      if (date) {
        const day = date.getDate().toString().padStart(2, '0');
        const month = (date.getMonth() + 1).toString().padStart(2, '0');
        const year = date.getFullYear();
        return `${day}/${month}/${year}`;
      }

      return dateString;
    } catch {
      return dateString;
    }
  };

  // Define getFilteredLots BEFORE any useMemo that depends on it
  const getFilteredLots = useCallback(() => {
    if (!packingData.length) return [];

    return packingData.filter(item => {
      const lotNumber = item.lotNumber?.toString().trim();

      // Check if lot is in issues sheet (allocated to packing)
      const isInIssuesSheet = lotNumber && issuesLotMap.has(lotNumber);

      // Get issues info for this lot
      const issuesInfo = issuesLotMap.get(lotNumber);

      // Check if packing is NOT complete (packingComplete field is empty, '-', or not a date)
      const packingComplete = issuesInfo?.packingComplete;
      const isPackingNotComplete = !packingComplete ||
        packingComplete === '' ||
        packingComplete === '-' ||
        packingComplete === 'No' ||
        packingComplete.toString().toLowerCase() === 'no' ||
        !isPackingComplete(packingComplete);

      // RAWPACK Check: Exclude lot if Date of Packing Complete has a date AND Sticker issue is YES
      const rawpackInfo = lotNumber ? rawpackLotMap.get(lotNumber) : null;
      const isExcludedByRawpack = rawpackInfo && rawpackInfo.isExcluded;

      // Include lots that are in issues sheet AND packing is not complete AND not excluded by RAWPACK
      return isInIssuesSheet && isPackingNotComplete && !isExcludedByRawpack;
    });
  }, [packingData, issuesLotMap, rawpackLotMap]);

  // Get unique values for filter dropdowns
  const filterOptions = useMemo(() => {
    const mergedLots = getFilteredLots().map(item => {
      const lotNumber = item.lotNumber?.toString().trim();
      const issuesInfo = issuesLotMap.get(lotNumber) || {};
      return {
        ...item,
        packingSupervisor: issuesInfo.packingSupervisor || item.packingSupervisor || '-',
        brand: item.brand || issuesInfo.brand || '-',
        season: item.season || issuesInfo.season || '-',
        directStitching: item.directStitching || issuesInfo.directStitching || '-'
      };
    });

    const options = {
      priority: [
        { value: 'all', label: 'All Priorities' },
        ...Array.from(new Set(mergedLots.map(item => item.priority || 'Normal').filter(Boolean)))
          .map(value => ({ value, label: value }))
      ],
      directStitching: [
        { value: 'all', label: 'All' },
        ...Array.from(new Set(mergedLots.map(item => item.directStitching || 'No').filter(Boolean)))
          .map(value => ({ value, label: value }))
      ],
      stitchingSupervisor: [
        { value: 'all', label: 'All Stitching Supervisors' },
        ...Array.from(new Set(mergedLots.map(item => item.supervisor || item.stitchingSupervisor).filter(val => val && val !== '-')))
          .map(value => ({ value, label: value }))
      ],
      packingSupervisor: [
        { value: 'all', label: 'All Supervisors' },
        ...Array.from(new Set(mergedLots.map(item => item.packingSupervisor).filter(val => val && val !== '-')))
          .map(value => ({ value, label: value }))
      ],
      brand: [
        { value: 'all', label: 'All Brands' },
        ...Array.from(new Set(mergedLots.map(item => item.brand).filter(val => val && val !== '-')))
          .map(value => ({ value, label: value }))
      ],
      season: [
        { value: 'all', label: 'All Seasons' },
        ...Array.from(new Set(mergedLots.map(item => item.season).filter(val => val && val !== '-')))
          .map(value => ({ value, label: value }))
      ],
      partyName: [
        { value: 'all', label: 'All Parties' },
        ...Array.from(new Set(mergedLots.map(item => item.partyName).filter(val => val && val !== '-')))
          .map(value => ({ value, label: value }))
      ]
    };

    return options;
  }, [getFilteredLots, issuesLotMap]);

  // Compute all available Financial Years based on Packing Date
  const allFinancialYears = useMemo(() => {
    const set = new Set();
    set.add(getCurrentFinancialYear());

    (issuesData || []).forEach(item => {
      const fy = getFinancialYearFromDate(item.packingDate);
      if (fy) set.add(fy);
    });

    (packingData || []).forEach(item => {
      const lotNumber = item.lotNumber?.toString().trim();
      const issuesInfo = issuesLotMap.get(lotNumber);
      if (issuesInfo?.packingDate) {
        const fy = getFinancialYearFromDate(issuesInfo.packingDate);
        if (fy) set.add(fy);
      }
    });

    const sorted = Array.from(set).sort().reverse();
    return ["ALL", ...sorted];
  }, [issuesData, packingData, issuesLotMap]);

  // Optimized fetch with caching
  useEffect(() => {
    const fetchAllData = async () => {
      const cachedData = localStorage.getItem('packingData');
      const cachedIssuesData = localStorage.getItem('issuesData');
      const cachedIssuesMap = localStorage.getItem('issuesLotMap');
      const cachedTimestamp = localStorage.getItem('packingDataTimestamp');

      if (cachedData && cachedIssuesData && cachedIssuesMap && cachedTimestamp) {
        const now = new Date().getTime();
        if (now - parseInt(cachedTimestamp) < 5 * 60 * 1000) {
          const parsedData = JSON.parse(cachedData);
          const parsedMap = new Map(JSON.parse(cachedIssuesMap));

          const updatedData = parsedData.map(item => ({
            ...item,
            packingPendingDays: calculatePackingPendingDays(item.packingDate)
          }));

          setPackingData(updatedData);
          setIssuesData(JSON.parse(cachedIssuesData));
          setIssuesLotMap(parsedMap);
          setLoading(false);
          return;
        }
      }

      await fetchAllSheetData();
    };

    fetchAllData();
  }, []);

  const fetchAllSheetData = async () => {
    try {
      setLoading(true);
      setError(null);

      const [indexData, issuesSheetData, rawpackSheetData] = await Promise.all([
        fetchSheetData(SPREADSHEET_ID, RANGE),
        fetchSheetData(ISSUES_SPREADSHEET_ID, ISSUES_SPREADSHEET_RANGE),
        fetchSheetData(RAWPACK_SPREADSHEET_ID, RAWPACK_SPREADSHEET_RANGE).catch(err => {
          console.warn('Could not fetch RAWPACK sheet:', err);
          return { values: [] };
        })
      ]);

      const transformedIndexData = transformSheetData(indexData.values || []);
      const { issuesData: transformedIssuesData, lotMap } = transformIssuesData(issuesSheetData.values || []);
      const { rawpackData: transformedRawpackData, rawpackLotMap: rawpackMap } = transformRawpackData(rawpackSheetData.values || []);

      try {
        localStorage.setItem('packingData', JSON.stringify(transformedIndexData));
        localStorage.setItem('issuesData', JSON.stringify(transformedIssuesData));
        localStorage.setItem('issuesLotMap', JSON.stringify(Array.from(lotMap.entries())));
        localStorage.setItem('rawpackData', JSON.stringify(transformedRawpackData));
        localStorage.setItem('rawpackLotMap', JSON.stringify(Array.from(rawpackMap.entries())));
        localStorage.setItem('packingDataTimestamp', new Date().getTime().toString());
      } catch (storageErr) {
        console.warn('LocalStorage quota exceeded. Skipping cache persistence:', storageErr);
      }

      setPackingData(transformedIndexData);
      setIssuesData(transformedIssuesData);
      setIssuesLotMap(lotMap);
      setRawpackData(transformedRawpackData);
      setRawpackLotMap(rawpackMap);
      setError(null);
    } catch (err) {
      setError(err.message);
      console.error('Error fetching data:', err);

      const cachedData = localStorage.getItem('packingData');
      const cachedIssuesData = localStorage.getItem('issuesData');
      const cachedIssuesMap = localStorage.getItem('issuesLotMap');
      const cachedRawpackMap = localStorage.getItem('rawpackLotMap');

      if (cachedData && cachedIssuesData && cachedIssuesMap) {
        const parsedData = JSON.parse(cachedData);
        const parsedMap = new Map(JSON.parse(cachedIssuesMap));
        const parsedRawpackMap = cachedRawpackMap ? new Map(JSON.parse(cachedRawpackMap)) : new Map();

        const updatedData = parsedData.map(item => ({
          ...item,
          packingPendingDays: calculatePackingPendingDays(item.packingDate)
        }));

        setPackingData(updatedData);
        setIssuesData(JSON.parse(cachedIssuesData));
        setIssuesLotMap(parsedMap);
        setRawpackLotMap(parsedRawpackMap);
        setLoading(false);
      }
    } finally {
      setLoading(false);
    }
  };

  const fetchSheetData = async (spreadsheetId, range) => {
    try {
      const res = await fetchSheetDataFromBackend(spreadsheetId, range);
      if (res && res.ok && Array.isArray(res.values) && res.values.length > 0) {
        return { values: res.values };
      }
    } catch (error) {
      console.warn(`Backend proxy fetch failed for [${spreadsheetId} - ${range}], using direct fallback:`, error);
    }

    try {
      const response = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}?key=${API_KEY}`
      );

      if (!response.ok) {
        throw new Error(`Failed to fetch data from spreadsheet: ${spreadsheetId}`);
      }

      return await response.json();
    } catch (error) {
      console.error('Fetch error:', error);
      throw error;
    }
  };

  const transformIssuesData = (values) => {
    if (!values || values.length === 0) return { issuesData: [], lotMap: new Map() };

    const headers = values[0].map(header => header.toString().trim().toLowerCase());
    const rows = values.slice(1);

    // Log headers to debug
    console.log('Issues sheet headers:', headers);

    // Find indices with more flexible matching
    const lotNumberIndex = headers.findIndex(h =>
      h.includes('lot number') || h.includes('lot') || h === 'lot'
    );

    const packingSupervisorIndex = headers.findIndex(h =>
      h.includes('packing supervisor') || h.includes('supervisor')
    );

    const packingDateIndex = headers.findIndex(h =>
      h.includes('packing date')
    );

    const packingCompleteIndex = headers.findIndex(h =>
      h.includes('packing complete')
    );

    const totalPcsIndex = headers.findIndex(h =>
      h.includes('total pcs') || h.includes('total pieces') || h === 'total pcs'
    );

    const wipPackingIndex = headers.findIndex(h =>
      h.includes('wip packing')
    );

    const totalManpowerIndex = headers.findIndex(h =>
      h.includes('total manpower')
    );

    const stitchingIssueDateIndex = headers.findIndex(h =>
      h.includes('stitching issue date')
    );

    const stitchingSupervisorIndex = headers.findIndex(h =>
      h.includes('stitching supervisor')
    );

    const brandIndex = headers.findIndex(h =>
      h.includes('brand')
    );

    const seasonIndex = headers.findIndex(h =>
      h.includes('season')
    );

    const directStitchingIndex = headers.findIndex(h =>
      h.includes('direct stitching')
    );

    console.log('Found indices:', {
      lotNumber: lotNumberIndex,
      totalPcs: totalPcsIndex,
      packingSupervisor: packingSupervisorIndex,
      packingDate: packingDateIndex,
      packingComplete: packingCompleteIndex
    });

    const issuesData = [];
    const lotMap = new Map();

    rows.forEach((row, index) => {
      if (row[lotNumberIndex] && row[lotNumberIndex].toString().trim() !== '') {
        const lotNumber = row[lotNumberIndex].toString().trim();

        // Get values with proper parsing
        const totalPcsValue = totalPcsIndex !== -1 ? row[totalPcsIndex] : '0';
        const packingCompleteValue = packingCompleteIndex !== -1 ? row[packingCompleteIndex] : '';

        const issueItem = {
          id: `issues-${index}`,
          lotNumber: lotNumber,
          packingSupervisor: packingSupervisorIndex !== -1 ? row[packingSupervisorIndex] || '' : '',
          packingDate: packingDateIndex !== -1 ? row[packingDateIndex] || '' : '',
          packingComplete: packingCompleteValue,
          totalPcs: totalPcsValue,
          wipPacking: wipPackingIndex !== -1 ? row[wipPackingIndex] || '0' : '0',
          totalManpower: totalManpowerIndex !== -1 ? row[totalManpowerIndex] || '0' : '0',
          stitchingIssueDate: stitchingIssueDateIndex !== -1 ? row[stitchingIssueDateIndex] || '' : '',
          stitchingSupervisor: stitchingSupervisorIndex !== -1 ? row[stitchingSupervisorIndex] || '' : '',
          brand: brandIndex !== -1 ? row[brandIndex] || '' : '',
          season: seasonIndex !== -1 ? row[seasonIndex] || '' : '',
          directStitching: directStitchingIndex !== -1 ? row[directStitchingIndex] || '' : ''
        };

        issuesData.push(issueItem);
        lotMap.set(lotNumber, {
          packingSupervisor: issueItem.packingSupervisor,
          packingDate: issueItem.packingDate,
          packingComplete: issueItem.packingComplete,
          totalPcs: issueItem.totalPcs,
          wipPacking: issueItem.wipPacking,
          totalManpower: issueItem.totalManpower,
          stitchingIssueDate: issueItem.stitchingIssueDate,
          stitchingSupervisor: issueItem.stitchingSupervisor,
          brand: issueItem.brand,
          season: issueItem.season,
          directStitching: issueItem.directStitching
        });
      }
    });

    console.log('Sample issue item:', issuesData[0]);
    return { issuesData, lotMap };
  };

  const transformRawpackData = (values) => {
    if (!values || values.length === 0) return { rawpackData: [], rawpackLotMap: new Map() };

    let headerRowIdx = 0;
    for (let i = 0; i < Math.min(values.length, 5); i++) {
      const rowStr = (values[i] || []).join(' ').toLowerCase();
      if (rowStr.includes('lot no') || rowStr.includes('item') || rowStr.includes('packing person')) {
        headerRowIdx = i;
        break;
      }
    }

    const headers = (values[headerRowIdx] || []).map(h => h ? h.toString().trim().toLowerCase() : '');
    const rows = values.slice(headerRowIdx + 1);

    const getCol = (name) => headers.findIndex(h => h.includes(name));

    const lotCol = getCol('lot no') >= 0 ? getCol('lot no') : getCol('lot');
    const lot2Col = getCol('lot no.2') >= 0 ? getCol('lot no.2') : getCol('lot2');
    const packingCompleteDateCol = getCol('date of packing complete');
    const stickerIssueCol = getCol('sticker issue') >= 0 ? getCol('sticker issue') : getCol('sticker');

    const rawpackData = [];
    const rawpackLotMap = new Map();

    rows.forEach((row, idx) => {
      const lotVal1 = lotCol >= 0 ? (row[lotCol] || '').toString().trim() : '';
      const lotVal2 = lot2Col >= 0 ? (row[lot2Col] || '').toString().trim() : '';
      const lotNumber = lotVal1 || lotVal2;

      if (lotNumber && lotNumber !== '-' && lotNumber !== '0') {
        const packingCompleteDateVal = packingCompleteDateCol >= 0 ? (row[packingCompleteDateCol] || '').toString().trim() : '';
        const stickerIssueVal = stickerIssueCol >= 0 ? (row[stickerIssueCol] || '').toString().trim().toLowerCase() : '';

        const hasCompleteDate = packingCompleteDateVal && packingCompleteDateVal !== '-' && packingCompleteDateVal !== '#N/A' && packingCompleteDateVal !== '00/01/00';
        const isStickerIssueYes = stickerIssueVal === 'yes';

        // Exclude lot if Date of Packing Complete has a date AND Sticker issue is YES
        const isExcludedFromAlloted = hasCompleteDate && isStickerIssueYes;

        const item = {
          id: `rawpack-${idx}`,
          lotNumber: lotNumber,
          packingCompleteDate: packingCompleteDateVal,
          stickerIssue: stickerIssueVal,
          isExcluded: isExcludedFromAlloted
        };

        rawpackData.push(item);
        rawpackLotMap.set(lotNumber, item);
      }
    });

    return { rawpackData, rawpackLotMap };
  };

  const transformSheetData = (values) => {
    if (!values || values.length === 0) return [];

    const sheetHeaders = values[0].map(header => header.toString().trim().toLowerCase());
    const rows = values.slice(1);

    const headerIndices = {
      lotNumber: sheetHeaders.findIndex(h => h.includes('lot number')),
      fabric: sheetHeaders.findIndex(h => h.includes('fabric')),
      brand: sheetHeaders.findIndex(h => h.includes('brand')),
      garmentType: sheetHeaders.findIndex(h => h.includes('garment type')),
      style: sheetHeaders.findIndex(h => h.includes('style')),
      partyName: sheetHeaders.findIndex(h => h.includes('party name')),
      directStitching: sheetHeaders.findIndex(h => h.includes('direct stitching')),
      dateOfIssue: sheetHeaders.findIndex(h => h.includes('date of issue')),
      supervisor: sheetHeaders.findIndex(h => h.includes('supervisor')),
      stitchingIssueQty: sheetHeaders.findIndex(h => h.includes('stitching issue qty')),
      sizes: sheetHeaders.findIndex(h => h.includes('sizes')),
      shades: sheetHeaders.findIndex(h => h.includes('shades')),
      savedAt: sheetHeaders.findIndex(h => h.includes('saved at')),
      imageUrl: sheetHeaders.findIndex(h => h.includes('image url')),
      season: sheetHeaders.findIndex(h => h.includes('season')),
      challanHistory: sheetHeaders.findIndex(h => h.includes('challan history')),
      zipOrderDate: sheetHeaders.findIndex(h => h.includes('zip order date')),
      zipReceivedDate: sheetHeaders.findIndex(h => h.includes('zip received date')),
      wipStatus: sheetHeaders.findIndex(h => h.includes('wip status')),
      completedStatus: sheetHeaders.findIndex(h => h.includes('completed status')),
      mwk: sheetHeaders.findIndex(h => h.includes('m/w/k')),
      jobOrderDate: sheetHeaders.findIndex(h => h.includes('joborder date')),
      manpower: sheetHeaders.findIndex(h => h.includes('manpower')),
      cuttingQty: sheetHeaders.findIndex(h => h.includes('cutting qty')),
      priority: sheetHeaders.findIndex(h => h.includes('prioirty') || h.includes('priority')),
      sticker: sheetHeaders.findIndex(h => h.includes('sticker'))
    };

    return rows.reduce((acc, row, index) => {
      const lotNumber = headerIndices.lotNumber !== -1 ? row[headerIndices.lotNumber] : '';

      if (!lotNumber || lotNumber.toString().trim() === '') {
        return acc;
      }

      const item = {
        id: index,
        originalRow: index + 2
      };

      Object.entries(headerIndices).forEach(([fieldName, headerIndex]) => {
        if (headerIndex !== -1 && row[headerIndex] !== undefined) {
          item[fieldName] = row[headerIndex] || '';
        } else {
          item[fieldName] = '';
        }
      });

      if (item.completedStatus) {
        const parsedStatus = parseCompletedStatus(item.completedStatus);
        item.completedStatusDisplay = parsedStatus.displayDate;
        item.completedStatusRaw = parsedStatus.rawStatus;
      } else {
        item.completedStatusDisplay = '-';
      }

      if (item.stitchingIssueQty) {
        item.stitchingIssueQtyNum = parseInt(item.stitchingIssueQty) || 0;
      }

      if (item.dateOfIssue) {
        item.dateOfIssue = formatDate(item.dateOfIssue);
      }
      if (item.jobOrderDate) {
        item.jobOrderDate = formatDate(item.jobOrderDate);
      }

      acc.push(item);
      return acc;
    }, []);
  };

  // Check if a lot matches the pending days range filter
  const matchesPendingDaysRange = (pendingDays, rangeValue) => {
    if (rangeValue === 'all') return true;

    switch (rangeValue) {
      case '0-7': return pendingDays <= 7;
      case '8-15': return pendingDays > 7 && pendingDays <= 15;
      case '16-30': return pendingDays > 15 && pendingDays <= 30;
      case '30+': return pendingDays > 30;
      default: return true;
    }
  };

  // Memoized filtered data for display with issues data merged and filters applied
  const displayData = useMemo(() => {
    const filteredLots = getFilteredLots();

    const mergedData = filteredLots.map(item => {
      const lotNumber = item.lotNumber?.toString().trim();
      const issuesInfo = issuesLotMap.get(lotNumber) || {};

      // Calculate pending days based on packing date
      const packingPendingDays = calculatePackingPendingDays(issuesInfo.packingDate);

      // Check if packing is complete
      const packingComplete = isPackingComplete(issuesInfo.packingComplete);

      return {
        ...item,
        packingSupervisor: issuesInfo.packingSupervisor || '-',
        packingDate: issuesInfo.packingDate ? formatDate(issuesInfo.packingDate) : '-',
        packingComplete: packingComplete ? 'Yes' : 'No',
        packingCompleteRaw: issuesInfo.packingComplete || '-',
        packingPendingDays: packingPendingDays,
        totalPcs: issuesInfo.totalPcs || '0',
        totalManpower: issuesInfo.totalManpower || '0',
        wipPacking: issuesInfo.wipPacking || '0',
        stitchingIssueDate: issuesInfo.stitchingIssueDate || '-',
        stitchingSupervisor: issuesInfo.stitchingSupervisor || '-',
        issuesBrand: issuesInfo.brand || '-',
        issuesSeason: issuesInfo.season || '-',
        issuesDirectStitching: issuesInfo.directStitching || '-'
      };
    });

    // Apply multi-select filters
    let filteredData = mergedData;

    if (filters.priority.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.priority.includes(item.priority || 'Normal')
      );
    }

    if (filters.directStitching.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.directStitching.includes(item.directStitching || 'No') ||
        filters.directStitching.includes(item.issuesDirectStitching || 'No')
      );
    }

    if (filters.stitchingSupervisor.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.stitchingSupervisor.includes(item.supervisor) ||
        filters.stitchingSupervisor.includes(item.stitchingSupervisor)
      );
    }

    if (filters.packingSupervisor.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.packingSupervisor.includes(item.packingSupervisor)
      );
    }

    if (filters.brand.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.brand.includes(item.brand) ||
        filters.brand.includes(item.issuesBrand)
      );
    }

    if (filters.season.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.season.includes(item.season) ||
        filters.season.includes(item.issuesSeason)
      );
    }

    if (filters.partyName.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.partyName.includes(item.partyName)
      );
    }

    // Apply pending days range multi-filter using packing pending days
    if (filters.pendingDaysRange.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.pendingDaysRange.some(rangeVal => matchesPendingDaysRange(item.packingPendingDays || 0, rangeVal))
      );
    }

    // Apply search
    if (searchTerm) {
      const searchLower = searchTerm.toLowerCase();
      const fieldsToSearch = ['lotNumber', 'fabric', 'brand', 'garmentType', 'style', 'partyName', 'supervisor', 'season', 'mwk', 'priority', 'packingSupervisor', 'issuesBrand', 'stitchingSupervisor'];

      filteredData = filteredData.filter(item =>
        fieldsToSearch.some(field =>
          item[field] && item[field].toString().toLowerCase().includes(searchLower)
        )
      );
    }

    // Apply Financial Year filter based on Packing Date
    if (financialYearFilter !== "ALL") {
      filteredData = filteredData.filter((item) => {
        const packingDateRaw = item.packingDateRaw || issuesLotMap.get(item.lotNumber?.toString().trim())?.packingDate;
        const fy = getFinancialYearFromDate(packingDateRaw);
        return fy === financialYearFilter;
      });
    }

    return filteredData;
  }, [getFilteredLots, issuesLotMap, searchTerm, filters, financialYearFilter]);

  // Pagination
  const paginatedData = useMemo(() => {
    const startIndex = (currentPage - 1) * itemsPerPage;
    return displayData.slice(startIndex, startIndex + itemsPerPage);
  }, [displayData, currentPage, itemsPerPage]);

  const totalPages = Math.ceil(displayData.length / itemsPerPage);

  // Memoized summary stats - removed stitching qty references
  const summaryStats = useMemo(() => {
    const totalPackingLots = packingData.filter(item => {
      const lotNumber = item.lotNumber?.toString().trim();
      return lotNumber && issuesLotMap.has(lotNumber);
    }).length;

    const completedPackingLots = packingData.filter(item => {
      const lotNumber = item.lotNumber?.toString().trim();
      const issuesInfo = issuesLotMap.get(lotNumber);
      return lotNumber && issuesInfo && isPackingComplete(issuesInfo.packingComplete);
    }).length;

    const agingSummary = {
      '0-7 days': displayData.filter(item => item.packingPendingDays <= 7).length,
      '8-15 days': displayData.filter(item => item.packingPendingDays > 7 && item.packingPendingDays <= 15).length,
      '16-30 days': displayData.filter(item => item.packingPendingDays > 15 && item.packingPendingDays <= 30).length,
      '30+ days': displayData.filter(item => item.packingPendingDays > 30).length
    };

    const totalPcsSum = displayData.reduce((sum, item) => sum + (parseInt(item.totalPcs) || 0), 0);

    return {
      totalItems: displayData.length,
      totalPcs: totalPcsSum,
      highPriority: displayData.filter(item => item.priority?.toLowerCase() === 'high').length,
      directStitching: displayData.filter(item =>
        item.directStitching?.toLowerCase() === 'yes' ||
        item.issuesDirectStitching?.toLowerCase() === 'yes'
      ).length,
      totalInIssuesSheet: issuesLotMap.size,
      completedPackingLots: completedPackingLots,
      pendingPackingLots: totalPackingLots - completedPackingLots,
      avgPackingPendingDays: displayData.length > 0
        ? Math.round(displayData.reduce((sum, item) => sum + (item.packingPendingDays || 0), 0) / displayData.length)
        : 0,
      agingSummary
    };
  }, [displayData, packingData, issuesLotMap]);

  const handleSearch = useCallback((e) => {
    setSearchTerm(e.target.value);
    setCurrentPage(1);
  }, []);

  const handleFilterChange = useCallback((filterName, selectedValue, selectedLabel) => {
    setFilters(prev => ({
      ...prev,
      [filterName]: { value: selectedValue, label: selectedLabel }
    }));
    setCurrentPage(1);
  }, []);

  const clearAllFilters = useCallback(() => {
    setFilters({
      priority: [],
      directStitching: [],
      stitchingSupervisor: [],
      packingSupervisor: [],
      brand: [],
      season: [],
      pendingDaysRange: [],
      partyName: []
    });
    setFinancialYearFilter(getCurrentFinancialYear());
    setSearchTerm('');
    setCurrentPage(1);
  }, []);

  // Get color based on pending days for PDF
  const getPendingDaysColor = (days) => {
    if (days <= 7) return [40, 167, 69]; // Green
    if (days <= 15) return [255, 193, 7]; // Yellow/Orange
    if (days <= 30) return [253, 126, 20]; // Orange
    return [220, 53, 69]; // Red
  };

  // Export to PDF with professional formatting & embedded images
  const exportToPDF = useCallback(async () => {
    try {
      setExportLoading(true);

      const doc = new jsPDF({
        orientation: 'landscape',
        unit: 'mm',
        format: 'a4'
      });

      const headerColor = [15, 76, 129];
      const textColor = [17, 24, 39];
      const lotNumberColor = [239, 68, 68];
      const partyColor = [107, 33, 168];

      const pageWidth = doc.internal.pageSize.getWidth();
      const pageHeight = doc.internal.pageSize.getHeight();
      const margin = 8;
      const contentWidth = pageWidth - (margin * 2);

      // Pre-load images for PDF embedding
      const imageBase64Map = new Map();
      await Promise.all(
        displayData.map(async (item) => {
          if (item.imageUrl) {
            const directUrl = getDirectImageUrl(item.imageUrl);
            const b64 = await getBase64ImageFromUrl(directUrl);
            if (b64) {
              imageBase64Map.set(item.id, b64);
            }
          }
        })
      );

      let currentY = 25;

      const drawHeader = () => {
        doc.setFillColor(255, 255, 255);
        doc.rect(0, 0, pageWidth, 30, 'F');

        doc.setFontSize(18);
        doc.setTextColor(15, 76, 129);
        doc.setFont('times', 'bold');
        doc.text('PACKING ALLOTTED & IN PROGRESS REPORT', pageWidth / 2, 12, { align: 'center' });

        doc.setFontSize(10);
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(15, 76, 129);

        const today = new Date();
        const reportDate = `${String(today.getDate()).padStart(2, '0')}/${String(today.getMonth() + 1).padStart(2, '0')}/${String(today.getFullYear()).slice(-2)}`;
        doc.text(`Report Date: ${reportDate}`, margin, 25);

        const centerX = pageWidth / 2;
        doc.text(`Total Lots: ${displayData.length} | Total Pieces: ${(summaryStats.totalPcs || 0).toLocaleString()} | High Priority: ${summaryStats.highPriority}`, centerX, 25, { align: 'center' });

        const filterText = `Total: ${displayData.length} records`;
        doc.text(filterText, pageWidth - margin, 25, { align: 'right' });
      };

      drawHeader();

      currentY = 32;

      const PDF_HEADERS = [
        'Sr',
        'Image',
        'Lot No',
        'Party Name',
        'Fabric',
        'Brand',
        'Garment',
        'Style',
        'Total Pcs',
        'Priority',
        'Pkg Date',
        'Pkg Days',
        'Supervisor',
        'Pkg Supervisor'
      ];

      const baseWidths = {
        0: 6,   // Sr
        1: 12,  // Image
        2: 18,  // Lot No
        3: 26,  // Party Name
        4: 28,  // Fabric
        5: 22,  // Brand
        6: 22,  // Garment
        7: 22,  // Style
        8: 16,  // Total Pcs
        9: 16,  // Priority
        10: 20, // Pkg Date
        11: 16, // Pkg Days
        12: 22, // Supervisor
        13: 22  // Pkg Supervisor
      };

      const totalBaseWidth = Object.values(baseWidths).reduce((a, b) => a + b, 0);
      const scaleFactor = contentWidth / totalBaseWidth;

      const columnWidths = {};
      Object.keys(baseWidths).forEach((key) => {
        columnWidths[key] = baseWidths[key] * scaleFactor;
      });

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

      const body = displayData.map((item, rowIndex) => {
        const rowBgColor = rowIndex % 2 === 0 ? [255, 255, 255] : [250, 250, 250];

        return [
          { content: (rowIndex + 1).toString(), styles: { cellWidth: columnWidths[0], fontSize: 9, halign: 'center', fontStyle: 'bold', fillColor: rowBgColor } },
          { content: '', styles: { cellWidth: columnWidths[1], halign: 'center', fillColor: rowBgColor } },
          { content: item.lotNumber || '-', styles: { cellWidth: columnWidths[2], fontSize: 10, halign: 'center', fontStyle: 'bold', textColor: lotNumberColor, fillColor: rowBgColor } },
          { content: item.partyName || '-', styles: { cellWidth: columnWidths[3], fontSize: 9, halign: 'center', fontStyle: 'bold', textColor: partyColor, fillColor: rowBgColor } },
          { content: item.fabric || '-', styles: { cellWidth: columnWidths[4], fontSize: 9, halign: 'center', fillColor: rowBgColor } },
          { content: item.brand || item.issuesBrand || '-', styles: { cellWidth: columnWidths[5], fontSize: 9, halign: 'center', fillColor: rowBgColor } },
          { content: item.garmentType || '-', styles: { cellWidth: columnWidths[6], fontSize: 9, halign: 'center', fillColor: rowBgColor } },
          { content: item.style || '-', styles: { cellWidth: columnWidths[7], fontSize: 9, halign: 'center', fillColor: rowBgColor } },
          { content: (item.totalPcs || '0').toString(), styles: { cellWidth: columnWidths[8], fontSize: 10, halign: 'center', fontStyle: 'bold', textColor: [239, 68, 68], fillColor: rowBgColor } },
          { content: item.priority || 'Normal', styles: { cellWidth: columnWidths[9], fontSize: 9, halign: 'center', fillColor: rowBgColor } },
          { content: item.packingDate || '-', styles: { cellWidth: columnWidths[10], fontSize: 9, halign: 'center', fillColor: rowBgColor } },
          { content: (item.packingPendingDays || 0).toString(), styles: { cellWidth: columnWidths[11], fontSize: 9, halign: 'center', fontStyle: 'bold', fillColor: rowBgColor } },
          { content: item.supervisor || item.stitchingSupervisor || '-', styles: { cellWidth: columnWidths[12], fontSize: 9, halign: 'center', fillColor: rowBgColor } },
          { content: item.packingSupervisor || '-', styles: { cellWidth: columnWidths[13], fontSize: 9, halign: 'center', fillColor: rowBgColor } }
        ];
      });

      autoTable(doc, {
        head: headers,
        body: body,
        startY: currentY,
        margin: { left: margin, right: margin },
        tableWidth: contentWidth,
        styles: {
          fontSize: 8.5,
          minCellHeight: 12,
          cellPadding: 1.5,
          lineColor: [0, 0, 0],
          lineWidth: 0.25,
          textColor: textColor,
          halign: 'center',
          valign: 'middle',
          overflow: 'linebreak'
        },
        headStyles: {
          fillColor: headerColor,
          textColor: [255, 255, 255],
          fontStyle: 'bold',
          lineColor: [0, 0, 0],
          lineWidth: 0.35,
          halign: 'center'
        },
        bodyStyles: {
          lineColor: [0, 0, 0],
          lineWidth: 0.25
        },
        alternateRowStyles: {
          fillColor: [248, 250, 252],
          lineColor: [0, 0, 0],
          lineWidth: 0.25
        },
        didDrawCell: (data) => {
          // Embed Image Thumbnail into Column 1
          if (data.column.index === 1 && data.cell.section === 'body') {
            const rowItem = displayData[data.row.index];
            if (rowItem) {
              const b64 = imageBase64Map.get(rowItem.id);
              if (b64) {
                try {
                  const imgWidth = 10;
                  const imgHeight = 10;
                  const posX = data.cell.x + (data.cell.width - imgWidth) / 2;
                  const posY = data.cell.y + (data.cell.height - imgHeight) / 2;
                  doc.addImage(b64, posX, posY, imgWidth, imgHeight);
                } catch (e) {
                  console.warn('Could not draw image in PDF cell:', e);
                }
              }
            }
          }

          return true;
        }
      });

      const pageCount = doc.internal.getNumberOfPages();
      for (let i = 1; i <= pageCount; i++) {
        doc.setPage(i);
        doc.setFontSize(7);
        doc.setTextColor(100, 100, 100);
        doc.text(`Page ${i} of ${pageCount}`, pageWidth - margin, pageHeight - 5, { align: 'right' });
        doc.text('Confidential - Internal Use Only', margin, pageHeight - 5);
      }

      const date = new Date().toISOString().split('T')[0];
      doc.save(`Packing_Allotted_In_Progress_${date}.pdf`);
    } catch (error) {
      console.error('Error exporting to PDF:', error);
      alert('Error exporting to PDF. Please try again.');
    } finally {
      setExportLoading(false);
    }
  }, [displayData, summaryStats]);

  // Export to Excel - updated to remove wipPacking and stitching qty, use totalPcs
  const exportToExcel = useCallback(() => {
    try {
      setExportLoading(true);

      const exportData = displayData.map(item => ({
        'Lot Number': item.lotNumber,
        'Party Name': item.partyName,
        'Fabric': item.fabric,
        'Brand': item.brand || item.issuesBrand || '-',
        'Garment Type': item.garmentType,
        'Style': item.style,
        'Supervisor': item.supervisor || item.stitchingSupervisor || '-',
        'Direct Stitching': item.directStitching || item.issuesDirectStitching || '-',
        'Date of Issue': item.dateOfIssue || item.stitchingIssueDate || '-',
        'Total Pcs': item.totalPcs,
        'Priority': item.priority || 'Normal',
        'M/W/K': item.mwk || '-',
        'Season': item.season || item.issuesSeason || '-',
        'Packing Supervisor': item.packingSupervisor || '-',
        'Packing Date': item.packingDate || '-',
        'Packing Days': item.packingPendingDays || 0,
        'Packing Complete': item.packingComplete,
        'Total Manpower': item.totalManpower,
        'WIP Packing': item.wipPacking,
        'Cutting Qty': item.cuttingQty || '0',
        'Manpower': item.manpower || '0',
        'Job Order Date': item.jobOrderDate || '-',
        'Status': 'Packing In Progress'
      }));

      const ws = XLSX.utils.json_to_sheet(exportData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Packing In Progress');

      // Auto-size columns
      const colWidths = [];
      exportData.forEach(row => {
        Object.keys(row).forEach((key, i) => {
          const value = row[key] ? row[key].toString().length : 10;
          colWidths[i] = Math.max(colWidths[i] || 10, Math.min(value, 50));
        });
      });
      ws['!cols'] = colWidths.map(width => ({ wch: width }));

      const date = new Date().toISOString().split('T')[0];
      XLSX.writeFile(wb, `packing_in_progress_${date}.xlsx`);
    } catch (error) {
      console.error('Error exporting to Excel:', error);
      alert('Error exporting to Excel. Please try again.');
    } finally {
      setExportLoading(false);
    }
  }, [displayData]);

  const getPendingDaysClass = (days) => {
    if (days <= 7) return 'pending-days-low';
    if (days <= 15) return 'pending-days-medium';
    if (days <= 30) return 'pending-days-high';
    return 'pending-days-critical';
  };

  const handleRefresh = useCallback(() => {
    localStorage.removeItem('packingData');
    localStorage.removeItem('issuesData');
    localStorage.removeItem('issuesLotMap');
    localStorage.removeItem('packingDataTimestamp');
    fetchAllSheetData();
  }, []);

  const handlePageChange = (page) => {
    setCurrentPage(page);
    document.querySelector('.table-container')?.scrollTo(0, 0);
  };

  const toggleFilters = () => {
    setShowFilters(!showFilters);
  };

  if (loading) {
    return (
      <div className="pending-packing-container">
        <div className="loading-spinner">
          <div className="spinner"></div>
          <p>Loading packing in progress data from sheets...</p>
        </div>
      </div>
    );
  }

  return (
    <div className="pending-packing-container">
      <div className="header-section">
        <div className="title-section">
          <button onClick={handleGoBack} className="back-button" title="Go Back">
            ← Back
          </button>
          <h2>📦 Packing In Progress Report</h2>
          <div className="stats-badge">
            Lots in Packing: {displayData.length}
          </div>
        </div>

        <div className="actions">
          <input
            type="text"
            placeholder="🔍 Search by Lot, Fabric, Brand, Style, Party, Packing Supervisor..."
            value={searchTerm}
            onChange={handleSearch}
            className="search-input"
          />

          <button
            onClick={toggleFilters}
            className={`filter-toggle-btn ${showFilters ? 'active' : ''}`}
            title="Toggle Filters"
          >
            🎚️ Filters {showFilters ? '▲' : '▼'}
          </button>

          <div className="export-buttons">
            <button
              onClick={exportToExcel}
              className="export-btn excel-btn"
              disabled={exportLoading || displayData.length === 0}
              title="Export to Excel"
            >
              {exportLoading ? '⏳' : '📊'} Excel
            </button>
            <button
              onClick={exportToPDF}
              className="export-btn pdf-btn"
              disabled={exportLoading || displayData.length === 0}
              title="Export to PDF"
            >
              {exportLoading ? '⏳' : '📄'} PDF
            </button>
            <button
              onClick={handleRefresh}
              className="refresh-btn"
              disabled={loading}
            >
              🔄 Refresh
            </button>
          </div>
        </div>
      </div>

      {/* Filters Section */}
      {showFilters && (
        <div className="filters-section">
          <div className="filters-header">
            <h3>Filter Options</h3>
            <button onClick={clearAllFilters} className="clear-all-btn">
              Clear All Filters
            </button>
          </div>
          <div className="filters-grid">
            <div className="filter-item">
              <label style={{ fontWeight: '700', color: '#1e1b4b' }}>📅 Financial Year (Packing Date):</label>
              <select
                value={financialYearFilter}
                onChange={(e) => {
                  setFinancialYearFilter(e.target.value);
                  setCurrentPage(1);
                }}
                className="filter-select"
                style={{ fontWeight: '700', color: '#1e1b4b', border: '1.5px solid #6366f1' }}
              >
                {allFinancialYears.map((fy) => (
                  <option key={fy} value={fy}>
                    {fy === 'ALL' ? '🌐 All Data (Whole Data)' : `FY ${fy}`}
                  </option>
                ))}
              </select>
            </div>

            <MultiSelectDropdown
              label="Priority:"
              options={filterOptions.priority}
              selectedValues={filters.priority}
              onChange={(newVals) => {
                setFilters(prev => ({ ...prev, priority: newVals }));
                setCurrentPage(1);
              }}
            />

            <MultiSelectDropdown
              label="Direct Stitching:"
              options={filterOptions.directStitching}
              selectedValues={filters.directStitching}
              onChange={(newVals) => {
                setFilters(prev => ({ ...prev, directStitching: newVals }));
                setCurrentPage(1);
              }}
            />

            <MultiSelectDropdown
              label="Stitching Supervisor:"
              options={filterOptions.stitchingSupervisor}
              selectedValues={filters.stitchingSupervisor}
              onChange={(newVals) => {
                setFilters(prev => ({ ...prev, stitchingSupervisor: newVals }));
                setCurrentPage(1);
              }}
            />

            <MultiSelectDropdown
              label="Packing Supervisor:"
              options={filterOptions.packingSupervisor}
              selectedValues={filters.packingSupervisor}
              onChange={(newVals) => {
                setFilters(prev => ({ ...prev, packingSupervisor: newVals }));
                setCurrentPage(1);
              }}
            />

            <MultiSelectDropdown
              label="Brand:"
              options={filterOptions.brand}
              selectedValues={filters.brand}
              onChange={(newVals) => {
                setFilters(prev => ({ ...prev, brand: newVals }));
                setCurrentPage(1);
              }}
            />

            <MultiSelectDropdown
              label="Season:"
              options={filterOptions.season}
              selectedValues={filters.season}
              onChange={(newVals) => {
                setFilters(prev => ({ ...prev, season: newVals }));
                setCurrentPage(1);
              }}
            />

            <MultiSelectDropdown
              label="Party Name:"
              options={filterOptions.partyName}
              selectedValues={filters.partyName}
              onChange={(newVals) => {
                setFilters(prev => ({ ...prev, partyName: newVals }));
                setCurrentPage(1);
              }}
            />

            <MultiSelectDropdown
              label="Packing Days:"
              options={pendingDaysRanges}
              selectedValues={filters.pendingDaysRange}
              onChange={(newVals) => {
                setFilters(prev => ({ ...prev, pendingDaysRange: newVals }));
                setCurrentPage(1);
              }}
            />
          </div>

          {/* Active Filters Display */}
          <div className="active-filters">
            {Object.entries(filters).map(([key, selectedArr]) => {
              if (!selectedArr || selectedArr.length === 0) return null;
              return selectedArr.map((val) => (
                <span key={`${key}-${val}`} className="active-filter-tag">
                  {key}: <strong>{val}</strong>
                  <button
                    onClick={() => {
                      setFilters(prev => ({
                        ...prev,
                        [key]: prev[key].filter(v => v !== val)
                      }));
                    }}
                  >
                    ×
                  </button>
                </span>
              ));
            })}
          </div>
        </div>
      )}

      {error && (
        <div className="error-message">
          <p>⚠️ {error}</p>
        </div>
      )}

      {displayData.length === 0 ? (
        <div className="no-data">
          <p>📭 No lots currently in packing</p>
          <p>Lots allocated to packing will appear here</p>
          {(searchTerm || Object.values(filters).some(v => v.value !== 'all')) && (
            <button onClick={clearAllFilters} className="clear-search-btn">
              Clear All Filters
            </button>
          )}
        </div>
      ) : (
        <>
          <div className="table-container">
            <table className="packing-table">
              <thead>
                <tr>
                  <th>#</th>
                  <th>Image</th>
                  <th>Lot Number</th>
                  <th>Party Name</th>
                  <th>Fabric</th>
                  <th>Brand</th>
                  <th>Garment Type</th>
                  <th>Style</th>
                  <th>Supervisor</th>
                  <th>Direct Stitching</th>
                  <th>Date of Issue</th>
                  <th>Total Pcs</th>
                  <th>Priority</th>
                  <th>M/W/K</th>
                  <th>Season</th>
                  <th>Packing Supervisor</th>
                  <th>Packing Date</th>
                  <th>Packing Days</th>
                  <th>WIP Packing</th>
                  <th>Total Manpower</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody>
                {paginatedData.map((item, index) => (
                  <tr key={item.id} className="packing-in-progress-row">
                    <td>{(currentPage - 1) * itemsPerPage + index + 1}</td>
                    <td>
                      {item.imageUrl ? (
                        <img
                          src={getDirectImageUrl(item.imageUrl)}
                          alt={`Lot #${item.lotNumber}`}
                          onClick={() => setViewImageSrc(getDirectImageUrl(item.imageUrl))}
                          style={{
                            width: "44px",
                            height: "44px",
                            objectFit: "cover",
                            borderRadius: "8px",
                            border: "1px solid #cbd5e1",
                            cursor: "pointer",
                            boxShadow: "0 2px 6px rgba(0,0,0,0.1)"
                          }}
                          onError={(e) => { e.target.style.display = 'none'; }}
                        />
                      ) : (
                        <span style={{ color: "#94a3b8", fontSize: "0.78rem" }}>—</span>
                      )}
                    </td>
                    <td className="lot-number">{item.lotNumber || '-'}</td>
                    <td>{item.partyName || '-'}</td>
                    <td>{item.fabric || '-'}</td>
                    <td>{item.brand || item.issuesBrand || '-'}</td>
                    <td>{item.garmentType || '-'}</td>
                    <td>{item.style || '-'}</td>
                    <td>{item.supervisor || item.stitchingSupervisor || '-'}</td>
                    <td>
                      <span className={`status-badge ${(item.directStitching || item.issuesDirectStitching)?.toLowerCase() === 'yes' ? 'status-yes' : 'status-no'}`}>
                        {item.directStitching || item.issuesDirectStitching || '-'}
                      </span>
                    </td>
                    <td>{item.dateOfIssue || item.stitchingIssueDate || '-'}</td>
                    <td className="quantity-cell">
                      <span className={parseInt(item.totalPcs) > 0 ? 'positive-qty' : 'zero-qty'}>
                        {item.totalPcs || '0'}
                      </span>
                    </td>
                    <td>
                      <span className={`priority-badge priority-${(item.priority || 'normal').toLowerCase()}`}>
                        {item.priority || 'Normal'}
                      </span>
                    </td>
                    <td>{item.mwk || '-'}</td>
                    <td>{item.season || item.issuesSeason || '-'}</td>
                    <td>
                      <span className="packing-supervisor-badge">
                        {item.packingSupervisor || '-'}
                      </span>
                    </td>
                    <td>{item.packingDate || '-'}</td>
                    <td>
                      <span className={`pending-days-badge ${getPendingDaysClass(item.packingPendingDays)}`}>
                        {item.packingPendingDays || 0} {item.packingPendingDays === 1 ? 'day' : 'days'}
                      </span>
                    </td>
                    <td>{item.wipPacking || '0'}</td>
                    <td>{item.totalManpower || '0'}</td>
                    <td>
                      <span className="status-badge status-pending">
                        ⏳ In Progress
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* Pagination */}
          {totalPages > 1 && (
            <div className="pagination">
              <button
                onClick={() => handlePageChange(currentPage - 1)}
                disabled={currentPage === 1}
                className="page-btn"
              >
                ← Prev
              </button>

              <span className="page-info">
                Page {currentPage} of {totalPages}
              </span>

              <button
                onClick={() => handlePageChange(currentPage + 1)}
                disabled={currentPage === totalPages}
                className="page-btn"
              >
                Next →
              </button>
            </div>
          )}

          {/* Summary Section */}
          {/* <div className="summary-section">
            <div className="summary-stats">
              <span>📦 Lots in Packing: {summaryStats.totalItems}</span>
              <span>📦 Total Pcs: {summaryStats.totalPcs}</span>
              <span>⚡ High Priority: {summaryStats.highPriority}</span>
              <span>🎯 Direct Stitching: {summaryStats.directStitching}</span>
              <span>📋 Total Allocated: {summaryStats.totalInIssuesSheet}</span>
              <span>✅ Packing Complete: {summaryStats.completedPackingLots}</span>
              <span>⏳ Avg Packing Days: {summaryStats.avgPackingPendingDays}</span>
            </div>
            <div className="aging-summary">
              <span className="aging-title">Packing Aging:</span>
              <span className="aging-badge low">0-7 days: {summaryStats.agingSummary['0-7 days']}</span>
              <span className="aging-badge medium">8-15 days: {summaryStats.agingSummary['8-15 days']}</span>
              <span className="aging-badge high">16-30 days: {summaryStats.agingSummary['16-30 days']}</span>
              <span className="aging-badge critical">30+ days: {summaryStats.agingSummary['30+ days']}</span>
            </div>
          </div> */}
        </>
      )}

      {/* Image Preview Lightbox Modal */}
      {viewImageSrc && (
        <div
          onClick={() => setViewImageSrc(null)}
          style={{
            position: "fixed",
            top: 0, left: 0, right: 0, bottom: 0,
            background: "rgba(15, 23, 42, 0.85)",
            backdropFilter: "blur(8px)",
            zIndex: 9999,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            padding: "20px"
          }}
        >
          <div style={{ position: "relative", maxWidth: "90vw", maxHeight: "90vh" }}>
            <img
              src={viewImageSrc}
              alt="Garment Preview"
              style={{ maxWidth: "100%", maxHeight: "85vh", borderRadius: "16px", boxShadow: "0 25px 50px -12px rgba(0,0,0,0.5)" }}
            />
            <button
              onClick={() => setViewImageSrc(null)}
              style={{
                position: "absolute",
                top: "-15px",
                right: "-15px",
                background: "#ef4444",
                color: "#ffffff",
                border: "none",
                borderRadius: "50%",
                width: "36px",
                height: "36px",
                fontWeight: "800",
                fontSize: "1.1rem",
                cursor: "pointer",
                boxShadow: "0 4px 12px rgba(0,0,0,0.3)"
              }}
            >
              ✕
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

export default PackingAlloted;