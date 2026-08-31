import React, { useState, useEffect, useMemo, useCallback } from 'react';
import './PendingPackingtoIssue.css';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { GOOGLE_API_KEY, SPREADSHEET_IDS, SHEET_NAMES, fetchSheetDataFromBackend } from './config';

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

const PendingPackingtoIssue = () => {
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [packingData, setPackingData] = useState([]);
  const [issuesData, setIssuesData] = useState([]);
  const [issuesLotMap, setIssuesLotMap] = useState(new Map());
  const [rawpackData, setRawpackData] = useState([]);
  const [rawpackLotMap, setRawpackLotMap] = useState(new Map());
  const [viewImageSrc, setViewImageSrc] = useState(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const [itemsPerPage] = useState(50);
  const [exportLoading, setExportLoading] = useState(false);
  const [showFilters, setShowFilters] = useState(false);

  // Filter states - array of selected values for multi-selection
  const [filters, setFilters] = useState({
    supervisor: [],
    priority: [],
    directStitching: [],
    packingSupervisor: [],
    brand: [],
    fabric: [],
    garmentType: [],
    style: [],
    season: [],
    pendingDaysRange: [],
    partyName: []
  });

  // Google Sheets configuration for Index sheet from .env / config
  const API_KEY = GOOGLE_API_KEY;
  const SPREADSHEET_ID = SPREADSHEET_IDS.MAIN;
  const RANGE = `${SHEET_NAMES.INDEX}!A:AA`;

  // Issues spreadsheet configuration from .env / config
  const ISSUES_SPREADSHEET_ID = SPREADSHEET_IDS.ISSUES;
  const ISSUES_SPREADSHEET_RANGE = `${SHEET_NAMES.ISSUES}!A:R`;

  // RAWPACK spreadsheet configuration from .env / config for checking lot issue status
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

      const date = new Date(dateString);
      if (!isNaN(date.getTime())) {
        return date;
      }

      return null;
    } catch {
      return null;
    }
  };

  // Function to calculate pending days
  const calculatePendingDays = (completedDate) => {
    if (!completedDate || completedDate === '-') return 0;

    try {
      const completed = parseDate(completedDate);
      if (!completed) return 0;

      const today = new Date();
      today.setHours(0, 0, 0, 0);
      completed.setHours(0, 0, 0, 0);

      const diffTime = today - completed;
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

      return diffDays > 0 ? diffDays : 0;
    } catch (error) {
      console.error('Error calculating pending days:', error);
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
      if (!lotNumber) return false;
      const normalizedLot = lotNumber.toUpperCase();

      const hasCompletedStatus = item.completedStatus && item.completedStatus !== '';
      const isNotInIssuesSheet = !issuesLotMap.has(lotNumber) && !issuesLotMap.has(normalizedLot);

      const rawpackInfo = rawpackLotMap.get(lotNumber) || rawpackLotMap.get(normalizedLot) || null;
      const isRawpackIssuedOrCompleted = rawpackInfo && (Boolean(rawpackInfo.packingPerson) || Boolean(rawpackInfo.packingIssueDate) || rawpackInfo.isCompleted);

      // Exclude lot if present in Issues sheet OR issued/completed in RAWPACK sheet
      return hasCompletedStatus && isNotInIssuesSheet && !isRawpackIssuedOrCompleted;
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
      supervisor: [
        { value: 'all', label: 'All Supervisors' },
        ...Array.from(new Set(mergedLots.map(item => item.supervisor).filter(val => val && val !== '-')))
          .sort()
          .map(value => ({ value, label: value }))
      ],
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
      fabric: [
        { value: 'all', label: 'All Fabrics' },
        ...Array.from(new Set(mergedLots.map(item => item.fabric).filter(val => val && val !== '-')))
          .map(value => ({ value, label: value }))
      ],
      garmentType: [
        { value: 'all', label: 'All Garment Types' },
        ...Array.from(new Set(mergedLots.map(item => item.garmentType).filter(val => val && val !== '-')))
          .map(value => ({ value, label: value }))
      ],
      style: [
        { value: 'all', label: 'All Styles' },
        ...Array.from(new Set(mergedLots.map(item => item.style).filter(val => val && val !== '-')))
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

  // Always fetch fresh data on mount (bypassing stale localStorage)
  useEffect(() => {
    fetchAllSheetData();
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
        console.warn('LocalStorage quota exceeded. Skipping local storage persistence:', storageErr);
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
          pendingDays: calculatePendingDays(item.completedStatusDisplay)
        }));

        setPackingData(updatedData);
        setIssuesData(JSON.parse(cachedIssuesData));
        setIssuesLotMap(parsedMap);
        setRawpackLotMap(parsedRawpackMap);
        setError('Using cached data. Latest data could not be fetched.');
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

    const lotNumberIndex = headers.findIndex(h => h.includes('lot number') || h.includes('lot'));
    const packingSupervisorIndex = headers.findIndex(h => h.includes('packing supervisor'));
    const packingDateIndex = headers.findIndex(h => h.includes('packing date'));
    const packingCompleteIndex = headers.findIndex(h => h.includes('packing complete'));

    const issuesData = [];
    const lotMap = new Map();

    rows.forEach((row, index) => {
      if (row[lotNumberIndex]) {
        const lotNumber = row[lotNumberIndex].toString().trim();

        const issueItem = {
          id: `issues-${index}`,
          lotNumber: lotNumber,
          packingSupervisor: packingSupervisorIndex !== -1 ? row[packingSupervisorIndex] || '' : '',
          packingDate: packingDateIndex !== -1 ? row[packingDateIndex] || '' : '',
          packingComplete: packingCompleteIndex !== -1 ? row[packingCompleteIndex] || '' : ''
        };

        issuesData.push(issueItem);
        lotMap.set(lotNumber, {
          packingSupervisor: issueItem.packingSupervisor,
          packingDate: issueItem.packingDate,
          packingComplete: issueItem.packingComplete
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
    const packingPersonCol = getCol('packing person');
    const supervisorCol = getCol('supervisior') >= 0 ? getCol('supervisior') : getCol('supervisor');
    const packingIssueDateCol = getCol('date of packing issue');
    const packingCompleteDateCol = getCol('date of packing complete');
    const reportFindCol = getCol('report find');
    const completedCol = headers.findIndex(h => h && h.trim().toLowerCase() === 'completed') >= 0
      ? headers.findIndex(h => h && h.trim().toLowerCase() === 'completed')
      : getCol('completed');

    const rawpackData = [];
    const rawpackLotMap = new Map();

    rows.forEach((row, idx) => {
      const lotVal1 = lotCol >= 0 ? (row[lotCol] || '').toString().trim() : '';
      const lotVal2 = lot2Col >= 0 ? (row[lot2Col] || '').toString().trim() : '';
      const lotNumber = lotVal1 || lotVal2;

      if (lotNumber && lotNumber !== '-' && lotNumber !== '0') {
        const packingPersonVal = packingPersonCol >= 0 ? (row[packingPersonCol] || '').toString().trim() : '';
        const packingIssueDateVal = packingIssueDateCol >= 0 ? (row[packingIssueDateCol] || '').toString().trim() : '';
        const packingCompleteDateVal = packingCompleteDateCol >= 0 ? (row[packingCompleteDateCol] || '').toString().trim() : '';
        const reportFindVal = reportFindCol >= 0 ? (row[reportFindCol] || '').toString().trim().toLowerCase() : '';
        const completedVal = completedCol >= 0 ? (row[completedCol] || '').toString().trim().toLowerCase() : '';

        // Check if packing is completed (valid complete date or report find status) - ignoring remarks 3 / formula columns
        const isRawpackCompleted = (
          Boolean(packingCompleteDateVal && packingCompleteDateVal !== '-' && packingCompleteDateVal !== '#N/A' && packingCompleteDateVal !== '00/01/00') ||
          completedVal === 'yes' || completedVal === 'complete' || completedVal === 'completed' ||
          reportFindVal === 'complete' || reportFindVal === 'completed' || reportFindVal === 'yes'
        );

        // Check if lot has been issued to packing
        const isRawpackIssued = Boolean(packingPersonVal) || Boolean(packingIssueDateVal && packingIssueDateVal !== '-' && packingIssueDateVal !== '#N/A');

        const item = {
          id: `rawpack-${idx}`,
          lotNumber: lotNumber,
          packingPerson: packingPersonVal,
          supervisor: supervisorCol >= 0 ? row[supervisorCol] || '' : '',
          packingIssueDate: packingIssueDateVal,
          packingCompleteDate: packingCompleteDateVal,
          completedColValue: completedVal,
          isCompleted: isRawpackCompleted,
          isIssued: isRawpackIssued
        };

        rawpackData.push(item);
        rawpackLotMap.set(lotNumber, item);
        rawpackLotMap.set(lotNumber.toUpperCase(), item);
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

      item.pendingDays = calculatePendingDays(item.completedStatusDisplay);

      if (item.stitchingIssueQty) {
        item.stitchingIssueQtyNum = parseInt(item.stitchingIssueQty) || 0;
      }

      if (item.dateOfIssue) {
        item.dateOfIssue = formatDate(item.dateOfIssue);
      }
      if (item.jobOrderDate) {
        item.jobOrderDate = formatDate(item.jobOrderDate);
      }
      if (item.packingDate) {
        item.packingDate = formatDate(item.packingDate);
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

      return {
        ...item,
        packingSupervisor: issuesInfo.packingSupervisor || '-',
        packingDate: issuesInfo.packingDate ? formatDate(issuesInfo.packingDate) : '-',
        packingComplete: issuesInfo.packingComplete || '-'
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
        filters.directStitching.includes(item.directStitching || 'No')
      );
    }

    if (filters.supervisor.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.supervisor.includes(item.supervisor)
      );
    }

    if (filters.packingSupervisor.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.packingSupervisor.includes(item.packingSupervisor)
      );
    }

    if (filters.brand.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.brand.includes(item.brand)
      );
    }

    if (filters.fabric.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.fabric.includes(item.fabric)
      );
    }

    if (filters.garmentType.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.garmentType.includes(item.garmentType)
      );
    }

    if (filters.style.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.style.includes(item.style)
      );
    }

    if (filters.season.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.season.includes(item.season)
      );
    }

    if (filters.partyName.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.partyName.includes(item.partyName)
      );
    }

    // Apply pending days range multi-filter
    if (filters.pendingDaysRange.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.pendingDaysRange.some(rangeVal => matchesPendingDaysRange(item.pendingDays || 0, rangeVal))
      );
    }

    // Apply search
    if (searchTerm) {
      const searchLower = searchTerm.toLowerCase();
      const fieldsToSearch = ['lotNumber', 'fabric', 'brand', 'garmentType', 'style', 'partyName', 'supervisor', 'season', 'mwk', 'priority', 'packingSupervisor'];

      filteredData = filteredData.filter(item =>
        fieldsToSearch.some(field =>
          item[field] && item[field].toString().toLowerCase().includes(searchLower)
        )
      );
    }

    return filteredData;
  }, [getFilteredLots, issuesLotMap, searchTerm, filters]);

  // Pagination
  const paginatedData = useMemo(() => {
    const startIndex = (currentPage - 1) * itemsPerPage;
    return displayData.slice(startIndex, startIndex + itemsPerPage);
  }, [displayData, currentPage, itemsPerPage]);

  const totalPages = Math.ceil(displayData.length / itemsPerPage);

  // Memoized summary stats
  const summaryStats = useMemo(() => {
    const pendingLotsCount = packingData.filter(item => {
      const completedStatus = item.completedStatus?.toString().trim();
      return !completedStatus || completedStatus === '';
    }).length;

    const agingSummary = {
      '0-7 days': displayData.filter(item => item.pendingDays <= 7).length,
      '8-15 days': displayData.filter(item => item.pendingDays > 7 && item.pendingDays <= 15).length,
      '16-30 days': displayData.filter(item => item.pendingDays > 15 && item.pendingDays <= 30).length,
      '30+ days': displayData.filter(item => item.pendingDays > 30).length
    };

    const totalStitchingQty = displayData.reduce((sum, item) => sum + (item.stitchingIssueQtyNum || 0), 0);

    return {
      totalItems: displayData.length,
      totalStitchingQty: totalStitchingQty,
      highPriority: displayData.filter(item => item.priority?.toLowerCase() === 'high').length,
      directStitching: displayData.filter(item => item.directStitching?.toLowerCase() === 'yes').length,
      totalInIssuesSheet: issuesLotMap.size,
      pendingLotsExcluded: pendingLotsCount,
      avgPendingDays: displayData.length > 0
        ? Math.round(displayData.reduce((sum, item) => sum + (item.pendingDays || 0), 0) / displayData.length)
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
      supervisor: [],
      priority: [],
      directStitching: [],
      packingSupervisor: [],
      brand: [],
      fabric: [],
      garmentType: [],
      style: [],
      season: [],
      pendingDaysRange: [],
      partyName: []
    });
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

  // Export to PDF with professional formatting & embedded images - FULL PAGE WIDTH
  const exportToPDF = useCallback(async () => {
    try {
      setExportLoading(true);

      const doc = new jsPDF({
        orientation: 'landscape',
        unit: 'mm',
        format: 'a3'
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

        const totalQty = displayData.reduce((sum, item) => sum + (item.stitchingIssueQtyNum || 0), 0);
        const totalLots = displayData.length;

        doc.setFontSize(18);
        doc.setTextColor(15, 76, 129);
        doc.setFont('times', 'bold');
        doc.text('COMPLETED LOTS READY FOR PACKING REPORT', pageWidth / 2, 12, { align: 'center' });

        doc.setFontSize(10);
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(15, 76, 129);

        const today = new Date();
        const reportDate = `${String(today.getDate()).padStart(2, '0')}/${String(today.getMonth() + 1).padStart(2, '0')}/${String(today.getFullYear()).slice(-2)}`;
        doc.text(`Report Date: ${reportDate}`, margin, 25);

        const centerX = pageWidth / 2;
        doc.text(`Total Lots: ${totalLots} | Total Stitching Qty: ${totalQty.toLocaleString()} | High Priority: ${summaryStats.highPriority}`, centerX, 25, { align: 'center' });

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
        'Garment',
        'Style',
        'Brand',
        'Supervisor',
        'Direct',
        'Issue Date',
        'Stitching Qty',
        'Priority',
        'M/W/K',
        'Season',
        'Completed Date',
        'Pending Days',
        'Pkg Supervisor'
      ];

      const baseWidths = {
        0: 6,   // Sr
        1: 12,  // Image
        2: 16,  // Lot No
        3: 22,  // Party Name
        4: 24,  // Fabric
        5: 20,  // Garment
        6: 20,  // Style
        7: 20,  // Brand
        8: 18,  // Supervisor
        9: 12,  // Direct
        10: 18, // Issue Date
        11: 16, // Stitching Qty
        12: 14, // Priority
        13: 10, // M/W/K
        14: 12, // Season
        15: 18, // Completed Date
        16: 14, // Pending Days
        17: 20  // Pkg Supervisor
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
          { content: item.garmentType || '-', styles: { cellWidth: columnWidths[5], fontSize: 9, halign: 'center', fillColor: rowBgColor } },
          { content: item.style || '-', styles: { cellWidth: columnWidths[6], fontSize: 9, halign: 'center', fillColor: rowBgColor } },
          { content: item.brand || '-', styles: { cellWidth: columnWidths[7], fontSize: 9, halign: 'center', fillColor: rowBgColor } },
          { content: item.supervisor || '-', styles: { cellWidth: columnWidths[8], fontSize: 9, halign: 'center', fillColor: rowBgColor } },
          { content: item.directStitching || '-', styles: { cellWidth: columnWidths[9], fontSize: 9, halign: 'center', fillColor: rowBgColor } },
          { content: item.dateOfIssue || '-', styles: { cellWidth: columnWidths[10], fontSize: 9, halign: 'center', fillColor: rowBgColor } },
          { content: (item.stitchingIssueQty || '0').toString(), styles: { cellWidth: columnWidths[11], fontSize: 10, halign: 'center', fontStyle: 'bold', textColor: [239, 68, 68], fillColor: rowBgColor } },
          { content: item.priority || 'Normal', styles: { cellWidth: columnWidths[12], fontSize: 9, halign: 'center', fillColor: rowBgColor } },
          { content: item.mwk || '-', styles: { cellWidth: columnWidths[13], fontSize: 9, halign: 'center', fillColor: rowBgColor } },
          { content: item.season || '-', styles: { cellWidth: columnWidths[14], fontSize: 9, halign: 'center', fillColor: rowBgColor } },
          { content: item.completedStatusDisplay || '-', styles: { cellWidth: columnWidths[15], fontSize: 9, halign: 'center', fillColor: rowBgColor } },
          { content: (item.pendingDays || 0).toString(), styles: { cellWidth: columnWidths[16], fontSize: 9, halign: 'center', fontStyle: 'bold', fillColor: rowBgColor } },
          { content: item.packingSupervisor || '-', styles: { cellWidth: columnWidths[17], fontSize: 9, halign: 'center', fillColor: rowBgColor } }
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
      doc.save(`Completed_Lots_Ready_Packing_${date}.pdf`);
    } catch (error) {
      console.error('Error exporting to PDF:', error);
      alert('Error exporting to PDF. Please try again.');
    } finally {
      setExportLoading(false);
    }
  }, [displayData, summaryStats]);

  // Export to Excel
  const exportToExcel = useCallback(() => {
    try {
      setExportLoading(true);

      const exportData = displayData.map(item => ({
        'Lot Number': item.lotNumber,
        'Party Name': item.partyName,
        'Fabric': item.fabric,
        'Brand': item.brand,
        'Garment Type': item.garmentType,
        'Style': item.style,
        'Supervisor': item.supervisor,
        'Direct Stitching': item.directStitching,
        'Date of Issue': item.dateOfIssue || '-',
        'Stitching Qty': item.stitchingIssueQty,
        'Priority': item.priority,
        'M/W/K': item.mwk,
        'Season': item.season,
        'WIP Status': item.wipStatus,
        'Completed Date': item.completedStatusDisplay || '-',
        'Pending Days': item.pendingDays || 0,
        'Packing Supervisor': item.packingSupervisor,
        'Packing Date': item.packingDate || '-',
        'Packing Complete': item.packingComplete,
        'Cutting Qty': item.cuttingQty,
        'Manpower': item.manpower,
        'Job Order Date': item.jobOrderDate || '-',
        'Status': 'Completed - Ready for Packing'
      }));

      const ws = XLSX.utils.json_to_sheet(exportData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Completed Lots Ready');

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
      XLSX.writeFile(wb, `completed_lots_ready_for_packing_${date}.xlsx`);
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
    localStorage.removeItem('rawpackData');
    localStorage.removeItem('rawpackLotMap');
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
          <p>Loading completed lots data from sheets...</p>
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
          <h2>✅ Completed Lots Ready for Packing</h2>
          <div className="stats-badge">
            Total Records: {displayData.length}
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

      {/* Filters Section - Collapsible with all dropdowns */}
      {showFilters && (
        <div className="filters-section">
          <div className="filters-header">
            <h3>Filter Options</h3>
            <button onClick={clearAllFilters} className="clear-all-btn">
              Clear All Filters
            </button>
          </div>
          <div className="filters-grid">
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
              label="Supervisor:"
              options={filterOptions.supervisor}
              selectedValues={filters.supervisor}
              onChange={(newVals) => {
                setFilters(prev => ({ ...prev, supervisor: newVals }));
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
              label="Fabric:"
              options={filterOptions.fabric}
              selectedValues={filters.fabric}
              onChange={(newVals) => {
                setFilters(prev => ({ ...prev, fabric: newVals }));
                setCurrentPage(1);
              }}
            />

            <MultiSelectDropdown
              label="Garment Type:"
              options={filterOptions.garmentType}
              selectedValues={filters.garmentType}
              onChange={(newVals) => {
                setFilters(prev => ({ ...prev, garmentType: newVals }));
                setCurrentPage(1);
              }}
            />

            <MultiSelectDropdown
              label="Style:"
              options={filterOptions.style}
              selectedValues={filters.style}
              onChange={(newVals) => {
                setFilters(prev => ({ ...prev, style: newVals }));
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
              label="Pending Days:"
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
          <p>📭 No completed lots found</p>
          <p>Lots with completed status will appear here</p>
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
                  <th>Stitching Qty</th>
                  <th>Priority</th>
                  <th>M/W/K</th>
                  <th>Season</th>
                  <th>Stitching Completed Date</th>
                  <th>Pending Days</th>
                  <th>Packing Supervisor</th>
                </tr>
              </thead>
              <tbody>
                {paginatedData.map((item, index) => (
                  <tr key={item.id} className="completed-lot-row">
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
                    <td>{item.brand || '-'}</td>
                    <td>{item.garmentType || '-'}</td>
                    <td>{item.style || '-'}</td>
                    <td>{item.supervisor || '-'}</td>
                    <td>
                      <span className={`status-badge ${item.directStitching?.toLowerCase() === 'yes' ? 'status-yes' : 'status-no'}`}>
                        {item.directStitching || '-'}
                      </span>
                    </td>
                    <td>{item.dateOfIssue || '-'}</td>
                    <td className="quantity-cell">
                      <span className={item.stitchingIssueQtyNum > 0 ? 'positive-qty' : 'zero-qty'}>
                        {item.stitchingIssueQty || '0'}
                      </span>
                    </td>
                    <td>
                      <span className={`priority-badge priority-${(item.priority || 'normal').toLowerCase()}`}>
                        {item.priority || 'Normal'}
                      </span>
                    </td>
                    <td>{item.mwk || '-'}</td>
                    <td>{item.season || '-'}</td>
                    <td>
                      <span className="status-badge status-completed">
                        {item.completedStatusDisplay || '-'}
                      </span>
                    </td>
                    <td>
                      <span className={`pending-days-badge ${getPendingDaysClass(item.pendingDays)}`}>
                        {item.pendingDays || 0} {item.pendingDays === 1 ? 'day' : 'days'}
                      </span>
                    </td>
                    <td>{item.packingSupervisor || '-'}</td>
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
              <span>✅ Completed Lots: {summaryStats.totalItems}</span>
              <span>📦 Total Stitching Qty: {summaryStats.totalStitchingQty}</span>
              <span>⚡ High Priority: {summaryStats.highPriority}</span>
              <span>🎯 Direct Stitching: {summaryStats.directStitching}</span>
              <span>📋 In Issues Sheet: {summaryStats.totalInIssuesSheet}</span>
              <span>⏳ Pending Lots: {summaryStats.pendingLotsExcluded}</span>
              <span>📊 Avg Pending Days: {summaryStats.avgPendingDays}</span>
            </div>
            <div className="aging-summary">
              <span className="aging-title">Aging Summary:</span>
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

export default PendingPackingtoIssue;