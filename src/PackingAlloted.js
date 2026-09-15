import React, { useState, useEffect, useMemo, useCallback } from 'react';
import './PendingPackingtoIssue.css';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { GOOGLE_API_KEY, SPREADSHEET_IDS, SHEET_NAMES, fetchSheetDataFromBackend } from './config';
import { fetchRemarksForTab, saveRemarkForLot } from './embPrintRemarksService';

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
  const [barcodeData, setBarcodeData] = useState([]);
  const [barcodeLotMap, setBarcodeLotMap] = useState(new Map());

  // Filter states - array of selected values for multi-selection
  const [filters, setFilters] = useState({
    priority: [],
    directStitching: [],
    stitchingSupervisor: [],
    packingSupervisor: [],
    brand: [],
    fabric: [],
    garmentType: [],
    style: [],
    season: [],
    pendingDaysRange: [],
    partyName: [],
    remarks: []
  });

  // Remarks States & Realtime Synchronization
  const [remarksMap, setRemarksMap] = useState({});
  const [remarksModalOpen, setRemarksModalOpen] = useState(false);
  const [selectedRemarksLot, setSelectedRemarksLot] = useState(null);
  const [newRemarkInputText, setNewRemarkInputText] = useState('');
  const [savingRemark, setSavingRemark] = useState(false);

  // Fetch Remarks from Google Sheets & Subscribe to updates
  useEffect(() => {
    fetchRemarksForTab('PACKING_ALLOTED').then(map => {
      if (map && typeof map === 'object') {
        setRemarksMap(map);
      }
    });

    const handleRemarkUpdated = (e) => {
      if (e.detail && (e.detail.tabType === 'PACKING_ALLOTED' || e.detail.tabType === 'PACKING' || !e.detail.tabType)) {
        setRemarksMap(prev => ({
          ...prev,
          [e.detail.lotNumber]: e.detail.history
        }));
      }
    };

    window.addEventListener('emb_print_remark_updated', handleRemarkUpdated);
    return () => {
      window.removeEventListener('emb_print_remark_updated', handleRemarkUpdated);
    };
  }, []);

  const handleOpenRemarksModal = (item) => {
    setSelectedRemarksLot(item);
    setNewRemarkInputText('');
    setRemarksModalOpen(true);
  };

  const handleCloseRemarksModal = () => {
    setRemarksModalOpen(false);
    setSelectedRemarksLot(null);
    setNewRemarkInputText('');
  };

  const handleSaveRemark = async () => {
    if (!selectedRemarksLot || !newRemarkInputText.trim()) return;

    try {
      setSavingRemark(true);
      const lotNumber = selectedRemarksLot.lotNumber?.toString().trim();
      const updatedHistory = await saveRemarkForLot({
        tabType: 'PACKING_ALLOTED',
        lotNumber: lotNumber,
        partyName: selectedRemarksLot.partyName || '',
        fabric: selectedRemarksLot.fabric || '',
        style: selectedRemarksLot.style || selectedRemarksLot.garmentType || '',
        remarkText: newRemarkInputText.trim()
      });
      setRemarksMap(prev => ({
        ...prev,
        [lotNumber]: updatedHistory
      }));
      setNewRemarkInputText('');
      handleCloseRemarksModal();
    } catch (err) {
      console.error('Failed to save remark:', err);
      alert('Failed to save remark. Please try again.');
    } finally {
      setSavingRemark(false);
    }
  };

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

  // BARCODE spreadsheet configuration from .env / config
  const BARCODE_SPREADSHEET_ID = SPREADSHEET_IDS.BARCODE || '1dOCjNFwaAel5qun0_ZJVIGmREqjI76CJBBFIjM3NHv8';
  const BARCODE_SPREADSHEET_RANGE = `${SHEET_NAMES.BARCODE || 'LotBarcodeData'}!A:Z`;

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
      const normalizedLot = lotNumber ? lotNumber.toUpperCase() : '';

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

      // Barcode Check: Exclude lot if packing is marked complete in Barcode Data
      const barcodeInfo = normalizedLot ? barcodeLotMap.get(normalizedLot) : null;
      const isExcludedByBarcode = barcodeInfo && barcodeInfo.isCompleted;

      // Include lots that are in issues sheet AND packing is not complete AND not excluded by RAWPACK / Barcode
      return isInIssuesSheet && isPackingNotComplete && !isExcludedByRawpack && !isExcludedByBarcode;
    });
  }, [packingData, issuesLotMap, rawpackLotMap, barcodeLotMap]);

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
      remarks: [
        { value: 'all', label: 'All Remarks' },
        { value: 'WITH_REMARKS', label: '📝 With Remarks' },
        { value: 'WITHOUT_REMARKS', label: '⬜ Without Remarks' },
        ...Array.from(new Set(
          mergedLots
            .map(item => {
              const lot = item.lotNumber?.toString().trim();
              const hist = remarksMap[lot] || [];
              return hist.length > 0 ? hist[hist.length - 1].text : '';
            })
            .filter(Boolean)
        )).sort().map(text => ({ value: text, label: text }))
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
  }, [getFilteredLots, issuesLotMap, remarksMap]);

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

      const [indexData, issuesSheetData, rawpackSheetData, barcodeSheetData] = await Promise.all([
        fetchSheetData(SPREADSHEET_ID, RANGE),
        fetchSheetData(ISSUES_SPREADSHEET_ID, ISSUES_SPREADSHEET_RANGE),
        fetchSheetData(RAWPACK_SPREADSHEET_ID, RAWPACK_SPREADSHEET_RANGE).catch(err => {
          console.warn('Could not fetch RAWPACK sheet:', err);
          return { values: [] };
        }),
        fetchSheetData(BARCODE_SPREADSHEET_ID, BARCODE_SPREADSHEET_RANGE).catch(err => {
          console.warn('Could not fetch BARCODE sheet:', err);
          return { values: [] };
        })
      ]);

      const transformedIndexData = transformSheetData(indexData.values || []);
      const { issuesData: transformedIssuesData, lotMap } = transformIssuesData(issuesSheetData.values || []);
      const { rawpackData: transformedRawpackData, rawpackLotMap: rawpackMap } = transformRawpackData(rawpackSheetData.values || []);
      const { barcodeData: transformedBarcodeData, barcodeLotMap: barcodeMap } = transformBarcodeData(barcodeSheetData.values || []);

      try {
        localStorage.setItem('packingData', JSON.stringify(transformedIndexData));
        localStorage.setItem('issuesData', JSON.stringify(transformedIssuesData));
        localStorage.setItem('issuesLotMap', JSON.stringify(Array.from(lotMap.entries())));
        localStorage.setItem('rawpackData', JSON.stringify(transformedRawpackData));
        localStorage.setItem('rawpackLotMap', JSON.stringify(Array.from(rawpackMap.entries())));
        localStorage.setItem('barcodeData', JSON.stringify(transformedBarcodeData));
        localStorage.setItem('barcodeLotMap', JSON.stringify(Array.from(barcodeMap.entries())));
        localStorage.setItem('packingDataTimestamp', new Date().getTime().toString());
      } catch (storageErr) {
        console.warn('LocalStorage quota exceeded. Skipping cache persistence:', storageErr);
      }

      setPackingData(transformedIndexData);
      setIssuesData(transformedIssuesData);
      setIssuesLotMap(lotMap);
      setRawpackData(transformedRawpackData);
      setRawpackLotMap(rawpackMap);
      setBarcodeData(transformedBarcodeData);
      setBarcodeLotMap(barcodeMap);
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

  const transformBarcodeData = (values) => {
    if (!values || values.length === 0) return { barcodeData: [], barcodeLotMap: new Map() };

    const headers = (values[0] || []).map(h => (h || '').toString().trim().toLowerCase());
    const rows = values.slice(1);

    const lotCol = headers.findIndex(h => h.includes('lot number') || h.includes('lot no') || h === 'lot');
    const genDateCol = headers.findIndex(h => h.includes('generated date') || h.includes('timestamp') || h.includes('date'));
    const barcodeIdCol = headers.findIndex(h => h.includes('barcode id') || h.includes('barcode'));
    const statusCol = headers.findIndex(h => h.includes('status'));

    const barcodeData = [];
    const barcodeLotMap = new Map();

    rows.forEach((row, idx) => {
      const lotVal = lotCol >= 0 ? (row[lotCol] || '').toString().trim() : '';
      if (!lotVal || lotVal === '-' || lotVal === '0') return;

      const genDateRaw = genDateCol >= 0 ? (row[genDateCol] || '').toString().trim() : '';
      const barcodeId = barcodeIdCol >= 0 ? (row[barcodeIdCol] || '').toString().trim() : '';
      const status = statusCol >= 0 ? (row[statusCol] || '').toString().trim() : '';

      const normalizedLot = lotVal.toUpperCase();
      const hasDate = genDateRaw && genDateRaw !== '-' && genDateRaw !== '#N/A';

      const existing = barcodeLotMap.get(normalizedLot);
      if (!existing) {
        const item = {
          id: `barcode-${idx}`,
          lotNumber: lotVal,
          barcodeId: barcodeId,
          status: status,
          generatedDate: genDateRaw,
          packingCompleteDate: hasDate ? formatDate(genDateRaw) : '',
          isCompleted: Boolean(hasDate)
        };
        barcodeData.push(item);
        barcodeLotMap.set(normalizedLot, item);
      } else if (hasDate && !existing.isCompleted) {
        existing.generatedDate = genDateRaw;
        existing.packingCompleteDate = formatDate(genDateRaw);
        existing.isCompleted = true;
      }
    });

    return { barcodeData, barcodeLotMap };
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
        filters.season.includes(item.season) ||
        filters.season.includes(item.issuesSeason)
      );
    }

    if (filters.partyName.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.partyName.includes(item.partyName)
      );
    }

    // Apply remarks filter
    if (filters.remarks && filters.remarks.length > 0) {
      filteredData = filteredData.filter(item => {
        const lot = item.lotNumber?.toString().trim();
        const lotRemarks = remarksMap[lot] || [];
        const hasRemark = lotRemarks.length > 0 && Boolean(lotRemarks[lotRemarks.length - 1].text);
        const latestText = hasRemark ? lotRemarks[lotRemarks.length - 1].text : '';

        return filters.remarks.some(filterVal => {
          if (filterVal === 'WITH_REMARKS') return hasRemark;
          if (filterVal === 'WITHOUT_REMARKS') return !hasRemark;
          return latestText === filterVal || lotRemarks.some(r => r.text === filterVal);
        });
      });
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

      filteredData = filteredData.filter(item => {
        const directMatch = fieldsToSearch.some(field =>
          item[field] && item[field].toString().toLowerCase().includes(searchLower)
        );
        if (directMatch) return true;

        const lotNumber = item.lotNumber?.toString().trim();
        const lotRemarks = remarksMap[lotNumber] || [];
        return lotRemarks.some(r => r.text && r.text.toLowerCase().includes(searchLower));
      });
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
  }, [getFilteredLots, issuesLotMap, searchTerm, filters, financialYearFilter, remarksMap]);

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
      fabric: [],
      garmentType: [],
      style: [],
      season: [],
      pendingDaysRange: [],
      partyName: [],
      remarks: []
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

  // Professional PDF Export with centered cells, pure black text, embedded pictures, and executive summary
  const exportToPDF = useCallback(async () => {
    if (displayData.length === 0) {
      alert('No data to export');
      return;
    }

    try {
      setExportLoading(true);

      const doc = new jsPDF({
        orientation: 'landscape',
        unit: 'pt',
        format: 'a3'
      });

      const pageWidth = doc.internal.pageSize.getWidth();
      const pageHeight = doc.internal.pageSize.getHeight();
      const margin = 15;
      const contentWidth = pageWidth - (margin * 2);

      const totalLots = displayData.length;
      const totalPieces = displayData.reduce((sum, item) => sum + (Number(item.totalPcs) || 0), 0);
      const highPriorityCount = displayData.filter(item => (item.priority || '').toLowerCase() === 'high').length;
      const directLotsCount = displayData.filter(item => (item.directStitching || item.issuesDirectStitching || '').toLowerCase() === 'yes').length;
      const avgDays = totalLots > 0 ? Math.round(displayData.reduce((sum, item) => sum + (item.packingPendingDays || 0), 0) / totalLots) : 0;

      // Pre-load base64 images for all displayed items
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

      // Aggregations for Executive Summary
      const garmentMap = {};
      const seasonMap = {};
      const supervisorMap = {};
      const agingMap = {
        '0-7 Days': { totalLots: 0, totalPcs: 0 },
        '8-15 Days': { totalLots: 0, totalPcs: 0 },
        '16-30 Days': { totalLots: 0, totalPcs: 0 },
        '30+ Days': { totalLots: 0, totalPcs: 0 }
      };

      displayData.forEach(item => {
        const pcs = Number(item.totalPcs) || 0;
        const gType = (item.garmentType || 'Unknown').trim();
        const season = (item.season || item.issuesSeason || 'N/A').trim();
        const pSup = (item.packingSupervisor || 'Unassigned').trim();
        const days = item.packingPendingDays || 0;

        if (!garmentMap[gType]) garmentMap[gType] = { totalLots: 0, totalPcs: 0 };
        garmentMap[gType].totalLots += 1;
        garmentMap[gType].totalPcs += pcs;

        if (!seasonMap[season]) seasonMap[season] = { totalLots: 0, totalPcs: 0 };
        seasonMap[season].totalLots += 1;
        seasonMap[season].totalPcs += pcs;

        if (!supervisorMap[pSup]) supervisorMap[pSup] = { totalLots: 0, totalPcs: 0 };
        supervisorMap[pSup].totalLots += 1;
        supervisorMap[pSup].totalPcs += pcs;

        if (days <= 7) {
          agingMap['0-7 Days'].totalLots += 1;
          agingMap['0-7 Days'].totalPcs += pcs;
        } else if (days <= 15) {
          agingMap['8-15 Days'].totalLots += 1;
          agingMap['8-15 Days'].totalPcs += pcs;
        } else if (days <= 30) {
          agingMap['16-30 Days'].totalLots += 1;
          agingMap['16-30 Days'].totalPcs += pcs;
        } else {
          agingMap['30+ Days'].totalLots += 1;
          agingMap['30+ Days'].totalPcs += pcs;
        }
      });

      const sortedGarments = Object.keys(garmentMap).map(k => ({ name: k, totalLots: garmentMap[k].totalLots, totalPcs: garmentMap[k].totalPcs })).sort((a, b) => b.totalPcs - a.totalPcs);
      const sortedSeasons = Object.keys(seasonMap).map(k => ({ name: k, totalLots: seasonMap[k].totalLots, totalPcs: seasonMap[k].totalPcs })).sort((a, b) => b.totalPcs - a.totalPcs);
      const sortedSupervisors = Object.keys(supervisorMap).map(k => ({ name: k, totalLots: supervisorMap[k].totalLots, totalPcs: supervisorMap[k].totalPcs })).sort((a, b) => b.totalPcs - a.totalPcs);

      // Title Banner
      doc.setFillColor(15, 23, 42); // Dark Navy
      doc.rect(margin, 12, contentWidth, 48, 'F');

      doc.setFontSize(14);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(255, 255, 255);
      doc.text("MH FACTORY SUITE PRO - PACKING ALLOTTED & IN PROGRESS REPORT", pageWidth / 2, 30, { align: 'center' });

      doc.setFontSize(8.5);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(199, 210, 254);
      const subText = `Lots in Packing: ${totalLots}   |   Total Pieces: ${totalPieces.toLocaleString()}   |   Direct Lots: ${directLotsCount}   |   High Priority: ${highPriorityCount}   |   Avg Packing Days: ${avgDays}d   |   Generated: ${new Date().toLocaleDateString('en-IN')} ${new Date().toLocaleTimeString('en-IN')}`;
      doc.text(subText, pageWidth / 2, 48, { align: 'center' });

      // Filter Banner
      doc.setFillColor(241, 245, 249);
      doc.rect(margin, 63, contentWidth, 16, 'F');
      doc.setFontSize(8);
      doc.setFont('helvetica', 'italic');
      doc.setTextColor(0, 0, 0);
      const filterSummary = `Filters: Priority: ${filters.priority.length ? filters.priority.join(', ') : 'All'} | Direct: ${filters.directStitching.length ? filters.directStitching.join(', ') : 'All'} | Stitching Sup: ${filters.stitchingSupervisor.length ? filters.stitchingSupervisor.join(', ') : 'All'} | Pkg Sup: ${filters.packingSupervisor.length ? filters.packingSupervisor.join(', ') : 'All'} | Brand: ${filters.brand.length ? filters.brand.join(', ') : 'All'} | Fabric: ${filters.fabric.length ? filters.fabric.join(', ') : 'All'} | Garment: ${filters.garmentType.length ? filters.garmentType.join(', ') : 'All'} | Style: ${filters.style.length ? filters.style.join(', ') : 'All'} | Season: ${filters.season.length ? filters.season.join(', ') : 'All'}`;
      doc.text(filterSummary, pageWidth / 2, 74, { align: 'center' });

      // Table columns & rows
      const tableColumns = [
        '#',
        'Image',
        'Lot Number',
        'Garment Type',
        'Style',
        'Fabric',
        'Brand',
        'Total Pcs',
        'M/W/K',
        'Season',
        'Party Name',
        'Direct Stitching',
        'Supervisor',
        'Date of Issue',
        'Priority',
        'Packing Supervisor',
        'Packing Date',
        'Packing Days',
        'Status',
        'Remarks'
      ];

      const tableBody = displayData.map((item, idx) => {
        const lotNo = (item.lotNumber || '').toString().trim();
        const lotRemarks = remarksMap[lotNo] || [];
        const latestRemark = lotRemarks.length > 0 ? lotRemarks[lotRemarks.length - 1].text : '—';
        const pcs = Number(item.totalPcs) || 0;
        const isDirect = (item.directStitching || item.issuesDirectStitching || '').toLowerCase() === 'yes';

        return [
          (idx + 1).toString(),
          '', // Image cell rendered via didDrawCell
          lotNo || '—',
          item.garmentType || '—',
          item.style || '—',
          item.fabric || '—',
          item.brand || item.issuesBrand || '—',
          pcs.toLocaleString(),
          item.mwk || '—',
          item.season || item.issuesSeason || '—',
          item.partyName || '—',
          isDirect ? 'Yes' : 'No',
          item.supervisor || item.stitchingSupervisor || '—',
          item.dateOfIssue || item.stitchingIssueDate || '—',
          item.priority || 'Normal',
          item.packingSupervisor || '—',
          item.packingDate || '—',
          `${item.packingPendingDays || 0}d`,
          'In Progress',
          latestRemark
        ];
      });

      // Total Row
      tableBody.push([
        '',
        '',
        `TOTAL (${totalLots} Lots)`,
        '',
        '',
        '',
        '',
        totalPieces.toLocaleString(),
        '',
        '',
        '',
        `${directLotsCount} Direct`,
        '',
        '',
        `${highPriorityCount} High`,
        '',
        '',
        `${avgDays}d Avg`,
        'In Progress',
        ''
      ]);

      const baseWidths = {
        0: 24,   // #
        1: 36,   // Image
        2: 55,   // Lot Number
        3: 65,   // Garment Type
        4: 65,   // Style
        5: 72,   // Fabric
        6: 60,   // Brand
        7: 52,   // Total Pcs
        8: 42,   // M/W/K
        9: 48,   // Season
        10: 76,  // Party Name
        11: 46,  // Direct Stitching
        12: 65,  // Supervisor
        13: 56,  // Date of Issue
        14: 46,  // Priority
        15: 68,  // Packing Supervisor
        16: 56,  // Packing Date
        17: 46,  // Packing Days
        18: 56,  // Status
        19: 120  // Remarks
      };

      const sumBase = Object.values(baseWidths).reduce((a, b) => a + b, 0);
      const scale = contentWidth / sumBase;

      const columnStyles = {};
      Object.keys(baseWidths).forEach(k => {
        const w = baseWidths[k] * scale;
        columnStyles[k] = {
          cellWidth: w,
          halign: 'center',
          fontStyle: (k === '2' || k === '7' || k === '17') ? 'bold' : 'normal'
        };
      });

      autoTable(doc, {
        head: [tableColumns],
        body: tableBody,
        startY: 85,
        tableWidth: contentWidth,
        margin: { top: 85, right: margin, bottom: 25, left: margin },
        theme: "grid",
        styles: {
          fontSize: 8.5,
          cellPadding: { top: 3, right: 2, bottom: 3, left: 2 },
          overflow: "linebreak",
          valign: 'middle',
          halign: 'center',
          textColor: [0, 0, 0], // Pure Black
          lineColor: [0, 0, 0],
          lineWidth: 0.3,
          fontStyle: 'normal',
          minCellHeight: 25
        },
        headStyles: {
          fillColor: [15, 23, 42],
          textColor: [255, 255, 255],
          fontStyle: "bold",
          lineColor: [0, 0, 0],
          lineWidth: 0.5,
          halign: 'center',
          fontSize: 9,
          valign: 'middle',
          cellPadding: { top: 5, right: 2, bottom: 5, left: 2 },
          minCellHeight: 14
        },
        alternateRowStyles: {
          fillColor: [248, 250, 252]
        },
        columnStyles,
        didParseCell: function (data) {
          if (data.section === 'body') {
            const rowIndex = data.row.index;
            const isTotalRow = rowIndex === tableBody.length - 1;

            if (isTotalRow) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [226, 232, 240];
              data.cell.styles.textColor = [0, 0, 0];
              data.cell.styles.halign = 'center';
              return;
            }

            const row = displayData[rowIndex];
            if (!row) return;

            // Direct styling
            if (data.column.index === 11 && (row.directStitching || row.issuesDirectStitching || '').toLowerCase() === 'yes') {
              data.cell.styles.textColor = [21, 128, 61];
              data.cell.styles.fontStyle = 'bold';
            }

            // High Priority styling
            if (data.column.index === 14 && (row.priority || '').toLowerCase() === 'high') {
              data.cell.styles.textColor = [220, 38, 38];
              data.cell.styles.fontStyle = 'bold';
            }

            // Aging days styling
            if (data.column.index === 17) {
              const days = row.packingPendingDays || 0;
              if (days > 15) {
                data.cell.styles.textColor = [220, 38, 38];
                data.cell.styles.fontStyle = 'bold';
              } else if (days > 7) {
                data.cell.styles.textColor = [217, 119, 6];
              } else {
                data.cell.styles.textColor = [22, 163, 74];
              }
            }

            // Status styling
            if (data.column.index === 18) {
              data.cell.styles.textColor = [217, 119, 6];
              data.cell.styles.fontStyle = 'bold';
            }
          }
        },
        didDrawCell: function (data) {
          if (data.column.index === 1 && data.section === 'body') {
            const rowIndex = data.row.index;
            const isTotalRow = rowIndex === tableBody.length - 1;
            if (isTotalRow) return;

            const rowItem = displayData[rowIndex];
            if (!rowItem) return;

            const b64 = imageBase64Map.get(rowItem.id);
            if (b64) {
              try {
                const imgSize = 20;
                const posX = data.cell.x + (data.cell.width - imgSize) / 2;
                const posY = data.cell.y + (data.cell.height - imgSize) / 2;
                doc.addImage(b64, 'JPEG', posX, posY, imgSize, imgSize);
              } catch (e) {
                console.warn('Could not draw image in PDF cell:', e);
              }
            }
          }
        }
      });

      // --- 4-COLUMN SIDE-BY-SIDE EXECUTIVE SUMMARY ---
      const gBody = sortedGarments.map(item => [
        item.name,
        item.totalLots.toString(),
        item.totalPcs.toLocaleString(),
        totalPieces > 0 ? `${((item.totalPcs / totalPieces) * 100).toFixed(1)}%` : "0.0%"
      ]);
      gBody.push(["TOTAL", totalLots.toString(), totalPieces.toLocaleString(), "100.0%"]);

      const sBody = sortedSeasons.map(item => [
        item.name,
        item.totalLots.toString(),
        item.totalPcs.toLocaleString(),
        totalPieces > 0 ? `${((item.totalPcs / totalPieces) * 100).toFixed(1)}%` : "0.0%"
      ]);
      sBody.push(["TOTAL", totalLots.toString(), totalPieces.toLocaleString(), "100.0%"]);

      const supBody = sortedSupervisors.map(item => [
        item.name,
        item.totalLots.toString(),
        item.totalPcs.toLocaleString(),
        totalPieces > 0 ? `${((item.totalPcs / totalPieces) * 100).toFixed(1)}%` : "0.0%"
      ]);
      supBody.push(["TOTAL", totalLots.toString(), totalPieces.toLocaleString(), "100.0%"]);

      const ageBody = Object.entries(agingMap).map(([rangeName, data]) => [
        rangeName,
        data.totalLots.toString(),
        data.totalPcs.toLocaleString(),
        totalPieces > 0 ? `${((data.totalPcs / totalPieces) * 100).toFixed(1)}%` : "0.0%"
      ]);
      ageBody.push(["TOTAL", totalLots.toString(), totalPieces.toLocaleString(), "100.0%"]);

      const maxRows = Math.max(gBody.length, sBody.length, supBody.length, ageBody.length);
      const approxSummaryHeight = 55 + (maxRows * 18);

      let summaryStartY = doc.lastAutoTable.finalY + 22;
      const neededSpace = approxSummaryHeight + 35;
      if (summaryStartY + neededSpace > pageHeight - 30) {
        doc.addPage();
        summaryStartY = 40;
      } else {
        doc.setDrawColor(203, 213, 225);
        doc.setLineWidth(1);
        doc.line(margin, summaryStartY - 10, pageWidth - margin, summaryStartY - 10);
      }

      // Title Banner for Summary
      doc.setFillColor(15, 23, 42);
      doc.rect(margin, summaryStartY, contentWidth, 22, 'F');
      doc.setFontSize(10.5);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(255, 255, 255);
      doc.text("EXECUTIVE PRODUCTION SUMMARY & DISTRIBUTION", pageWidth / 2, summaryStartY + 14, { align: 'center' });

      const sumTableY = summaryStartY + 26;
      const totalAvailWidth = contentWidth;
      const colGap = 12;
      const singleTableWidth = (totalAvailWidth - (colGap * 3)) / 4;

      const summaryHeadStyles = {
        fontStyle: 'bold',
        textColor: [255, 255, 255],
        fontSize: 8,
        halign: 'center',
        valign: 'middle',
        cellPadding: 3
      };

      const summaryBodyStyles = {
        fontSize: 8,
        halign: 'center',
        valign: 'middle',
        cellPadding: 2.5,
        textColor: [0, 0, 0],
        lineColor: [0, 0, 0],
        lineWidth: 0.25
      };

      // 1. Garment Table
      autoTable(doc, {
        head: [['Garment Type', 'Lots', 'Qty', '%']],
        body: gBody,
        startY: sumTableY,
        margin: { left: margin },
        tableWidth: singleTableWidth,
        theme: 'grid',
        styles: summaryBodyStyles,
        headStyles: { ...summaryHeadStyles, fillColor: [15, 118, 110] },
        didParseCell: (d) => {
          if (d.section === 'body' && d.row.index === gBody.length - 1) {
            d.cell.styles.fontStyle = 'bold';
            d.cell.styles.fillColor = [226, 232, 240];
          }
        }
      });

      // 2. Season Table
      autoTable(doc, {
        head: [['Season', 'Lots', 'Qty', '%']],
        body: sBody,
        startY: sumTableY,
        margin: { left: margin + singleTableWidth + colGap },
        tableWidth: singleTableWidth,
        theme: 'grid',
        styles: summaryBodyStyles,
        headStyles: { ...summaryHeadStyles, fillColor: [67, 56, 202] },
        didParseCell: (d) => {
          if (d.section === 'body' && d.row.index === sBody.length - 1) {
            d.cell.styles.fontStyle = 'bold';
            d.cell.styles.fillColor = [226, 232, 240];
          }
        }
      });

      // 3. Packing Supervisor Table
      autoTable(doc, {
        head: [['Packing Sup', 'Lots', 'Qty', '%']],
        body: supBody,
        startY: sumTableY,
        margin: { left: margin + (singleTableWidth * 2) + (colGap * 2) },
        tableWidth: singleTableWidth,
        theme: 'grid',
        styles: summaryBodyStyles,
        headStyles: { ...summaryHeadStyles, fillColor: [30, 64, 175] },
        didParseCell: (d) => {
          if (d.section === 'body' && d.row.index === supBody.length - 1) {
            d.cell.styles.fontStyle = 'bold';
            d.cell.styles.fillColor = [226, 232, 240];
          }
        }
      });

      // 4. Aging Table
      autoTable(doc, {
        head: [['Aging Range', 'Lots', 'Qty', '%']],
        body: ageBody,
        startY: sumTableY,
        margin: { left: margin + (singleTableWidth * 3) + (colGap * 3) },
        tableWidth: singleTableWidth,
        theme: 'grid',
        styles: summaryBodyStyles,
        headStyles: { ...summaryHeadStyles, fillColor: [153, 27, 27] },
        didParseCell: (d) => {
          if (d.section === 'body' && d.row.index === ageBody.length - 1) {
            d.cell.styles.fontStyle = 'bold';
            d.cell.styles.fillColor = [226, 232, 240];
          }
        }
      });

      // Page numbers in footer
      const pageCount = doc.internal.getNumberOfPages();
      for (let i = 1; i <= pageCount; i++) {
        doc.setPage(i);
        doc.setFontSize(8);
        doc.setTextColor(100, 100, 100);
        doc.text(`Page ${i} of ${pageCount}`, pageWidth - margin, pageHeight - 10, { align: 'right' });
        doc.text('FACTORY SUITE PRO - CONFIDENTIAL', margin, pageHeight - 10);
      }

      const dateStr = new Date().toISOString().split('T')[0];
      doc.save(`Packing_Allotted_In_Progress_${dateStr}.pdf`);

    } catch (error) {
      console.error('Error exporting to PDF:', error);
      alert('Error exporting to PDF. Please try again.');
    } finally {
      setExportLoading(false);
    }
  }, [displayData, remarksMap, filters]);

  // Professional Multi-Sheet Excel Export with Executive Summary & Applied Filters
  const exportToExcel = useCallback(async () => {
    if (displayData.length === 0) {
      alert('No data to export');
      return;
    }

    try {
      setExportLoading(true);
      const workbook = new ExcelJS.Workbook();
      workbook.creator = 'MH Factory Suite Pro';
      workbook.created = new Date();

      const totalLots = displayData.length;
      const totalPieces = displayData.reduce((sum, item) => sum + (Number(item.totalPcs) || 0), 0);
      const highPriorityCount = displayData.filter(item => (item.priority || '').toLowerCase() === 'high').length;
      const directLotsCount = displayData.filter(item => (item.directStitching || item.issuesDirectStitching || '').toLowerCase() === 'yes').length;
      const avgDays = totalLots > 0 ? Math.round(displayData.reduce((sum, item) => sum + (item.packingPendingDays || 0), 0) / totalLots) : 0;

      // Aggregations for Executive Summary
      const garmentMap = {};
      const seasonMap = {};
      const supervisorMap = {};
      const partyMap = {};
      const agingMap = {
        '0-7 Days': { totalLots: 0, totalPcs: 0 },
        '8-15 Days': { totalLots: 0, totalPcs: 0 },
        '16-30 Days': { totalLots: 0, totalPcs: 0 },
        '30+ Days': { totalLots: 0, totalPcs: 0 }
      };

      displayData.forEach(item => {
        const pcs = Number(item.totalPcs) || 0;
        const gType = (item.garmentType || 'Unknown').trim();
        const season = (item.season || item.issuesSeason || 'N/A').trim();
        const pSup = (item.packingSupervisor || 'Unassigned').trim();
        const party = (item.partyName || (item.directStitching === 'yes' ? 'Direct Stitching' : '—')).trim();
        const days = item.packingPendingDays || 0;

        if (!garmentMap[gType]) garmentMap[gType] = { totalLots: 0, totalPcs: 0 };
        garmentMap[gType].totalLots += 1;
        garmentMap[gType].totalPcs += pcs;

        if (!seasonMap[season]) seasonMap[season] = { totalLots: 0, totalPcs: 0 };
        seasonMap[season].totalLots += 1;
        seasonMap[season].totalPcs += pcs;

        if (!supervisorMap[pSup]) supervisorMap[pSup] = { totalLots: 0, totalPcs: 0 };
        supervisorMap[pSup].totalLots += 1;
        supervisorMap[pSup].totalPcs += pcs;

        if (!partyMap[party]) partyMap[party] = { totalLots: 0, totalPcs: 0 };
        partyMap[party].totalLots += 1;
        partyMap[party].totalPcs += pcs;

        if (days <= 7) {
          agingMap['0-7 Days'].totalLots += 1;
          agingMap['0-7 Days'].totalPcs += pcs;
        } else if (days <= 15) {
          agingMap['8-15 Days'].totalLots += 1;
          agingMap['8-15 Days'].totalPcs += pcs;
        } else if (days <= 30) {
          agingMap['16-30 Days'].totalLots += 1;
          agingMap['16-30 Days'].totalPcs += pcs;
        } else {
          agingMap['30+ Days'].totalLots += 1;
          agingMap['30+ Days'].totalPcs += pcs;
        }
      });

      const sortedGarments = Object.keys(garmentMap).map(k => ({ name: k, totalLots: garmentMap[k].totalLots, totalPcs: garmentMap[k].totalPcs })).sort((a, b) => b.totalPcs - a.totalPcs);
      const sortedSeasons = Object.keys(seasonMap).map(k => ({ name: k, totalLots: seasonMap[k].totalLots, totalPcs: seasonMap[k].totalPcs })).sort((a, b) => b.totalPcs - a.totalPcs);
      const sortedSupervisors = Object.keys(supervisorMap).map(k => ({ name: k, totalLots: supervisorMap[k].totalLots, totalPcs: supervisorMap[k].totalPcs })).sort((a, b) => b.totalPcs - a.totalPcs);
      const sortedParties = Object.keys(partyMap).map(k => ({ name: k, totalLots: partyMap[k].totalLots, totalPcs: partyMap[k].totalPcs })).sort((a, b) => b.totalPcs - a.totalPcs);

      const thinBorder = {
        top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
      };

      // ================= SHEET 1: DATA TABLE =================
      const ws1 = workbook.addWorksheet('Packing In Progress', { views: [{ showGridLines: true }] });

      // Title Banner
      ws1.mergeCells('A1:U1');
      const titleCell = ws1.getCell('A1');
      titleCell.value = 'MH FACTORY SUITE PRO - PACKING ALLOTTED & IN PROGRESS REPORT';
      titleCell.font = { name: 'Segoe UI', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
      titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E1B4B' } };
      titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
      ws1.getRow(1).height = 30;

      // Subtitle KPI Banner
      ws1.mergeCells('A2:U2');
      const subCell = ws1.getCell('A2');
      subCell.value = `Lots in Packing: ${totalLots}   |   Total Pieces: ${totalPieces.toLocaleString()}   |   Direct Lots: ${directLotsCount}   |   High Priority: ${highPriorityCount}   |   Avg Packing Days: ${avgDays}d   |   Generated: ${new Date().toLocaleDateString('en-IN')} ${new Date().toLocaleTimeString('en-IN')}`;
      subCell.font = { name: 'Segoe UI', size: 9.5, color: { argb: 'FFC7D2FE' } };
      subCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF312E81' } };
      subCell.alignment = { horizontal: 'center', vertical: 'middle' };
      ws1.getRow(2).height = 22;

      // Table Header Row
      const tableHeaders = [
        '#', 'Lot Number', 'Garment Type', 'Style', 'Fabric', 'Brand',
        'Total Pcs', 'M/W/K', 'Season', 'Party Name', 'Direct Stitching',
        'Supervisor', 'Date of Issue', 'Priority', 'Packing Supervisor',
        'Packing Date', 'Packing Days', 'Status', 'Remarks'
      ];
      const headerRow = ws1.addRow(tableHeaders);
      headerRow.height = 24;
      headerRow.eachCell(cell => {
        cell.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = thinBorder;
      });

      // Data Rows
      displayData.forEach((item, index) => {
        const lotNo = (item.lotNumber || '').toString().trim();
        const lotRemarks = remarksMap[lotNo] || [];
        const latestRemark = lotRemarks.length > 0 ? lotRemarks[lotRemarks.length - 1].text : '';
        const pcs = Number(item.totalPcs) || 0;
        const days = item.packingPendingDays || 0;
        const isDirect = (item.directStitching || item.issuesDirectStitching || '').toLowerCase() === 'yes';

        const r = ws1.addRow([
          index + 1,
          lotNo,
          item.garmentType || '—',
          item.style || '—',
          item.fabric || '—',
          item.brand || item.issuesBrand || '—',
          pcs,
          item.mwk || '—',
          item.season || item.issuesSeason || '—',
          item.partyName || '—',
          isDirect ? 'Yes' : 'No',
          item.supervisor || item.stitchingSupervisor || '—',
          item.dateOfIssue || item.stitchingIssueDate || '—',
          item.priority || 'Normal',
          item.packingSupervisor || '—',
          item.packingDate || '—',
          days,
          'In Progress',
          latestRemark || '—'
        ]);

        r.height = 20;

        r.eachCell(cell => {
          cell.font = { name: 'Segoe UI', size: 9 };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.border = thinBorder;
          if (index % 2 === 1) {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
          }
        });

        // Total Pcs styling
        r.getCell(7).numFmt = '#,##0';
        r.getCell(7).font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FF1E40AF' } };

        // Days Aging formatting
        const daysCell = r.getCell(17);
        if (days <= 7) {
          daysCell.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FF16A34A' } };
        } else if (days <= 15) {
          daysCell.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FFD97706' } };
        } else {
          daysCell.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FFDC2626' } };
          daysCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
        }

        // Direct badge
        if (isDirect) {
          r.getCell(11).font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FF15803D' } };
        }
      });

      // Total Row
      const totalRow1 = ws1.addRow([
        '', `TOTAL (${totalLots} Lots)`, '', '', '', '',
        totalPieces, '', '', '', `${directLotsCount} Direct`,
        '', '', `${highPriorityCount} High`, '', '',
        `${avgDays}d Avg`, 'In Progress', ''
      ]);
      totalRow1.height = 24;
      totalRow1.eachCell(cell => {
        cell.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FF000000' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FF000000' } },
          bottom: { style: 'double', color: { argb: 'FF000000' } },
          left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
          right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
        };
      });
      totalRow1.getCell(7).numFmt = '#,##0';

      // Column widths
      const colWidths = [6, 15, 18, 18, 20, 16, 14, 10, 12, 22, 16, 18, 15, 12, 18, 15, 14, 15, 30];
      colWidths.forEach((w, i) => {
        ws1.getColumn(i + 1).width = w;
      });

      // ================= SHEET 2: EXECUTIVE SUMMARY =================
      const ws2 = workbook.addWorksheet('Executive Summary', { views: [{ showGridLines: true }] });

      // 1. Garment Type Breakdown
      ws2.mergeCells('A1:D1');
      const gTitle = ws2.getCell('A1');
      gTitle.value = '1. GARMENT TYPE BREAKDOWN (LOTS & PIECES DISTRIBUTION)';
      gTitle.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      gTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F766E' } };
      gTitle.alignment = { horizontal: 'left', vertical: 'middle' };
      ws2.getRow(1).height = 26;

      const gHeader = ws2.addRow(['Garment Type', 'Total Lots', 'Total Pieces (Qty)', 'Share %']);
      gHeader.height = 22;
      gHeader.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF134E4A' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = thinBorder;
      });

      sortedGarments.forEach((item, idx) => {
        const pct = totalPieces > 0 ? (item.totalPcs / totalPieces) : 0;
        const r = ws2.addRow([item.name, item.totalLots, item.totalPcs, pct]);
        r.height = 19;
        r.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(3).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(3).numFmt = '#,##0';
        r.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(4).numFmt = '0.0%';
        r.eachCell(c => {
          c.font = { name: 'Segoe UI', size: 9 };
          c.border = thinBorder;
          if (idx % 2 === 1) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        });
      });

      const gTotalRow = ws2.addRow(['TOTAL', totalLots, totalPieces, 1]);
      gTotalRow.height = 22;
      gTotalRow.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = { top: { style: 'thin' }, bottom: { style: 'double' }, left: { style: 'thin' }, right: { style: 'thin' } };
      });
      gTotalRow.getCell(3).numFmt = '#,##0';
      gTotalRow.getCell(4).numFmt = '0.0%';

      // Spacer
      ws2.addRow([]);

      // 2. Season Wise Breakdown
      const sStartRow = ws2.rowCount + 1;
      ws2.mergeCells(`A${sStartRow}:D${sStartRow}`);
      const sTitle = ws2.getCell(`A${sStartRow}`);
      sTitle.value = '2. SEASON WISE BREAKDOWN (LOTS & PIECES DISTRIBUTION)';
      sTitle.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      sTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4338CA' } };
      sTitle.alignment = { horizontal: 'left', vertical: 'middle' };
      ws2.getRow(sStartRow).height = 26;

      const sHeader = ws2.addRow(['Season', 'Total Lots', 'Total Pieces (Qty)', 'Share %']);
      sHeader.height = 22;
      sHeader.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF312E81' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = thinBorder;
      });

      sortedSeasons.forEach((item, idx) => {
        const pct = totalPieces > 0 ? (item.totalPcs / totalPieces) : 0;
        const r = ws2.addRow([item.name, item.totalLots, item.totalPcs, pct]);
        r.height = 19;
        r.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(3).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(3).numFmt = '#,##0';
        r.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(4).numFmt = '0.0%';
        r.eachCell(c => {
          c.font = { name: 'Segoe UI', size: 9 };
          c.border = thinBorder;
          if (idx % 2 === 1) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        });
      });

      const sTotalRow = ws2.addRow(['TOTAL', totalLots, totalPieces, 1]);
      sTotalRow.height = 22;
      sTotalRow.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = { top: { style: 'thin' }, bottom: { style: 'double' }, left: { style: 'thin' }, right: { style: 'thin' } };
      });
      sTotalRow.getCell(3).numFmt = '#,##0';
      sTotalRow.getCell(4).numFmt = '0.0%';

      // Spacer
      ws2.addRow([]);

      // 3. Packing Supervisor Breakdown
      const supStartRow = ws2.rowCount + 1;
      ws2.mergeCells(`A${supStartRow}:D${supStartRow}`);
      const supTitle = ws2.getCell(`A${supStartRow}`);
      supTitle.value = '3. PACKING SUPERVISOR BREAKDOWN';
      supTitle.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      supTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E40AF' } };
      supTitle.alignment = { horizontal: 'left', vertical: 'middle' };
      ws2.getRow(supStartRow).height = 26;

      const supHeader = ws2.addRow(['Packing Supervisor', 'Total Lots', 'Total Pieces (Qty)', 'Share %']);
      supHeader.height = 22;
      supHeader.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = thinBorder;
      });

      sortedSupervisors.forEach((item, idx) => {
        const pct = totalPieces > 0 ? (item.totalPcs / totalPieces) : 0;
        const r = ws2.addRow([item.name, item.totalLots, item.totalPcs, pct]);
        r.height = 19;
        r.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(3).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(3).numFmt = '#,##0';
        r.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(4).numFmt = '0.0%';
        r.eachCell(c => {
          c.font = { name: 'Segoe UI', size: 9 };
          c.border = thinBorder;
          if (idx % 2 === 1) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        });
      });

      const supTotalRow = ws2.addRow(['TOTAL', totalLots, totalPieces, 1]);
      supTotalRow.height = 22;
      supTotalRow.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = { top: { style: 'thin' }, bottom: { style: 'double' }, left: { style: 'thin' }, right: { style: 'thin' } };
      });
      supTotalRow.getCell(3).numFmt = '#,##0';
      supTotalRow.getCell(4).numFmt = '0.0%';

      // Spacer
      ws2.addRow([]);

      // 4. Aging Breakdown
      const ageStartRow = ws2.rowCount + 1;
      ws2.mergeCells(`A${ageStartRow}:D${ageStartRow}`);
      const ageTitle = ws2.getCell(`A${ageStartRow}`);
      ageTitle.value = '4. PACKING DAYS AGING BREAKDOWN';
      ageTitle.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      ageTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF991B1B' } };
      ageTitle.alignment = { horizontal: 'left', vertical: 'middle' };
      ws2.getRow(ageStartRow).height = 26;

      const ageHeader = ws2.addRow(['Aging Range', 'Total Lots', 'Total Pieces (Qty)', 'Share %']);
      ageHeader.height = 22;
      ageHeader.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7F1D1D' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = thinBorder;
      });

      Object.entries(agingMap).forEach(([rangeName, data], idx) => {
        const pct = totalPieces > 0 ? (data.totalPcs / totalPieces) : 0;
        const r = ws2.addRow([rangeName, data.totalLots, data.totalPcs, pct]);
        r.height = 19;
        r.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(3).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(3).numFmt = '#,##0';
        r.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(4).numFmt = '0.0%';
        r.eachCell(c => {
          c.font = { name: 'Segoe UI', size: 9 };
          c.border = thinBorder;
          if (idx % 2 === 1) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        });
      });

      const ageTotalRow = ws2.addRow(['TOTAL', totalLots, totalPieces, 1]);
      ageTotalRow.height = 22;
      ageTotalRow.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = { top: { style: 'thin' }, bottom: { style: 'double' }, left: { style: 'thin' }, right: { style: 'thin' } };
      });
      ageTotalRow.getCell(3).numFmt = '#,##0';
      ageTotalRow.getCell(4).numFmt = '0.0%';

      ws2.getColumn(1).width = 28;
      ws2.getColumn(2).width = 16;
      ws2.getColumn(3).width = 24;
      ws2.getColumn(4).width = 16;

      // ================= SHEET 3: APPLIED FILTERS =================
      const ws3 = workbook.addWorksheet('Applied Filters', { views: [{ showGridLines: true }] });
      ws3.mergeCells('A1:B1');
      const fTitle = ws3.getCell('A1');
      fTitle.value = 'APPLIED FILTERS & METADATA';
      fTitle.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      fTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E1B4B' } };
      fTitle.alignment = { horizontal: 'left', vertical: 'middle' };
      ws3.getRow(1).height = 26;

      const filterEntries = [
        ['Search Term', searchTerm || 'None'],
        ['Financial Year', financialYearFilter || 'All'],
        ['Priority', filters.priority.length ? filters.priority.join(', ') : 'All'],
        ['Direct Stitching', filters.directStitching.length ? filters.directStitching.join(', ') : 'All'],
        ['Stitching Supervisor', filters.stitchingSupervisor.length ? filters.stitchingSupervisor.join(', ') : 'All'],
        ['Packing Supervisor', filters.packingSupervisor.length ? filters.packingSupervisor.join(', ') : 'All'],
        ['Brand', filters.brand.length ? filters.brand.join(', ') : 'All'],
        ['Fabric', filters.fabric.length ? filters.fabric.join(', ') : 'All'],
        ['Garment Type', filters.garmentType.length ? filters.garmentType.join(', ') : 'All'],
        ['Style', filters.style.length ? filters.style.join(', ') : 'All'],
        ['Season', filters.season.length ? filters.season.join(', ') : 'All'],
        ['Party Name', filters.partyName.length ? filters.partyName.join(', ') : 'All'],
        ['Remarks', filters.remarks.length ? filters.remarks.join(', ') : 'All'],
        ['Packing Days Range', filters.pendingDaysRange.length ? filters.pendingDaysRange.join(', ') : 'All'],
        ['Matching Lots', totalLots.toString()],
        ['Total Pieces', totalPieces.toLocaleString()]
      ];

      filterEntries.forEach(([k, v], idx) => {
        const r = ws3.addRow([k, v]);
        r.height = 20;
        r.getCell(1).font = { name: 'Segoe UI', size: 9.5, bold: true };
        r.getCell(2).font = { name: 'Segoe UI', size: 9.5 };
        r.eachCell(c => {
          c.border = thinBorder;
          if (idx % 2 === 1) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        });
      });

      ws3.getColumn(1).width = 24;
      ws3.getColumn(2).width = 50;

      // Save Excel file
      const buffer = await workbook.xlsx.writeBuffer();
      const ts = new Date().toISOString().slice(0, 10);
      saveAs(new Blob([buffer]), `Packing_In_Progress_${ts}.xlsx`);

    } catch (err) {
      console.error('Error exporting to Excel:', err);
      alert(`Excel export failed: ${err.message}`);
    } finally {
      setExportLoading(false);
    }
  }, [displayData, remarksMap, searchTerm, financialYearFilter, filters]);

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

  const totalPcsCount = useMemo(() => {
    return displayData.reduce((sum, item) => {
      const val = parseInt(item.totalPcs, 10);
      return sum + (isNaN(val) ? 0 : val);
    }, 0);
  }, [displayData]);

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
          <div className="stats-badge">
            Total PCS: {totalPcsCount.toLocaleString()}
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
              label="Remarks Filter:"
              options={filterOptions.remarks || []}
              selectedValues={filters.remarks}
              onChange={(newVals) => {
                setFilters(prev => ({ ...prev, remarks: newVals }));
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
                  <th>Garment Type</th>
                  <th>Style</th>
                  <th>Fabric</th>
                  <th>Brand</th>
                  <th>Total Pcs</th>
                  <th>M/W/K</th>
                  <th>Season</th>
                  <th>Party Name</th>
                  <th>Direct Stitching</th>
                  <th>Supervisor</th>
                  <th>Date of Issue</th>
                  <th>Priority</th>
                  <th>Packing Supervisor</th>
                  <th>Packing Date</th>
                  <th>Packing Days</th>
                  <th>Status</th>
                  <th>Remarks</th>
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
                    <td>{item.garmentType || '-'}</td>
                    <td>{item.style || '-'}</td>
                    <td>{item.fabric || '-'}</td>
                    <td>{item.brand || item.issuesBrand || '-'}</td>
                    <td className="quantity-cell">
                      <span className={parseInt(item.totalPcs) > 0 ? 'positive-qty' : 'zero-qty'}>
                        {item.totalPcs || '0'}
                      </span>
                    </td>
                    <td>{item.mwk || '-'}</td>
                    <td>{item.season || item.issuesSeason || '-'}</td>
                    <td>{item.partyName || '-'}</td>
                    <td>
                      <span className={`status-badge ${(item.directStitching || item.issuesDirectStitching)?.toLowerCase() === 'yes' ? 'status-yes' : 'status-no'}`}>
                        {item.directStitching || item.issuesDirectStitching || '-'}
                      </span>
                    </td>
                    <td>{item.supervisor || item.stitchingSupervisor || '-'}</td>
                    <td>{item.dateOfIssue || item.stitchingIssueDate || '-'}</td>
                    <td>
                      <span className={`priority-badge priority-${(item.priority || 'normal').toLowerCase()}`}>
                        {item.priority || 'Normal'}
                      </span>
                    </td>
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
                    <td>
                      <span className="status-badge status-pending">
                        ⏳ In Progress
                      </span>
                    </td>

                    {/* Remarks Column */}
                    <td style={{ minWidth: '180px', maxWidth: '240px' }}>
                      {(() => {
                        const lot = (item.lotNumber || '').toString().trim();
                        const lotRemarks = remarksMap[lot] || [];
                        const latestRemark = lotRemarks.length > 0 ? lotRemarks[lotRemarks.length - 1] : null;

                        return (
                          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '5px' }}>
                            {latestRemark ? (
                              <div
                                onClick={() => handleOpenRemarksModal(item)}
                                style={{
                                  background: '#f8fafc',
                                  border: '1px solid #cbd5e1',
                                  borderLeft: '3.5px solid #6366f1',
                                  borderRadius: '8px',
                                  padding: '5px 8px',
                                  fontSize: '0.78rem',
                                  color: '#1e293b',
                                  width: '100%',
                                  boxSizing: 'border-box',
                                  textAlign: 'left',
                                  cursor: 'pointer'
                                }}
                                title="Click to view history or add remark"
                              >
                                <div style={{ fontWeight: 600, wordBreak: 'break-word', whiteSpace: 'normal', display: '-webkit-box', WebkitLineClamp: 2, WebkitBoxOrient: 'vertical', overflow: 'hidden' }}>
                                  {latestRemark.text}
                                </div>
                                <div style={{ fontSize: '0.68rem', color: '#64748b', marginTop: '3px' }}>
                                  🕒 {latestRemark.timestamp}
                                </div>
                              </div>
                            ) : (
                              <span style={{ color: '#94a3b8', fontSize: '0.76rem', fontStyle: 'italic' }}>No remarks</span>
                            )}

                            <button
                              type="button"
                              onClick={() => handleOpenRemarksModal(item)}
                              style={{
                                background: '#eef2ff',
                                color: '#4338ca',
                                border: '1px solid #c7d2fe',
                                padding: '3px 9px',
                                borderRadius: '6px',
                                fontSize: '0.72rem',
                                fontWeight: 700,
                                cursor: 'pointer',
                                display: 'inline-flex',
                                alignItems: 'center',
                                gap: '4px'
                              }}
                            >
                              <span>✏️</span> {latestRemark ? 'History / Edit' : '+ Add Remark'}
                            </button>
                          </div>
                        );
                      })()}
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
        </>
      )}

      {/* Interactive Remarks Modal */}
      {remarksModalOpen && selectedRemarksLot && (
        <div className="remarks-modal-overlay" onClick={handleCloseRemarksModal}>
          <div className="remarks-modal-box" onClick={(e) => e.stopPropagation()}>
            <div className="remarks-modal-header">
              <div className="remarks-modal-header-title">
                <span style={{ fontSize: '1.4rem' }}>💬</span>
                <h3>Lot Remarks</h3>
                <span className="remarks-modal-lot-tag">#{selectedRemarksLot.lotNumber}</span>
              </div>
              <button
                className="remarks-modal-close-btn"
                onClick={handleCloseRemarksModal}
                title="Close"
              >
                ✕
              </button>
            </div>

            <div className="remarks-modal-body">
              {/* Lot Information Summary */}
              <div className="remarks-lot-card">
                <div className="remarks-lot-field">
                  <span className="remarks-lot-field-label">Party</span>
                  <span className="remarks-lot-field-value">{selectedRemarksLot.partyName || '-'}</span>
                </div>
                <div className="remarks-lot-field">
                  <span className="remarks-lot-field-label">Fabric</span>
                  <span className="remarks-lot-field-value">{selectedRemarksLot.fabric || '-'}</span>
                </div>
                <div className="remarks-lot-field">
                  <span className="remarks-lot-field-label">Style / Garment</span>
                  <span className="remarks-lot-field-value">{selectedRemarksLot.style || selectedRemarksLot.garmentType || '-'}</span>
                </div>
                <div className="remarks-lot-field">
                  <span className="remarks-lot-field-label">Total Pcs</span>
                  <span className="remarks-lot-field-value" style={{ color: '#1e40af', fontWeight: '800' }}>
                    {selectedRemarksLot.totalPcs || '0'} pcs
                  </span>
                </div>
                <div className="remarks-lot-field">
                  <span className="remarks-lot-field-label">Packing Sup</span>
                  <span className="remarks-lot-field-value">{selectedRemarksLot.packingSupervisor || '-'}</span>
                </div>
                <div className="remarks-lot-field">
                  <span className="remarks-lot-field-label">Packing Days</span>
                  <span className="remarks-lot-field-value" style={{ color: '#4338ca', fontWeight: '800' }}>
                    {selectedRemarksLot.packingPendingDays || 0} days
                  </span>
                </div>
              </div>

              {/* Remarks History */}
              <div className="remarks-history-section">
                <h4 className="remarks-history-title">
                  <span>📜</span> Remarks History ({(remarksMap[selectedRemarksLot.lotNumber?.toString().trim()] || []).length})
                </h4>
                <div className="remarks-history-list">
                  {(() => {
                    const history = remarksMap[selectedRemarksLot.lotNumber?.toString().trim()] || [];
                    if (history.length === 0) {
                      return (
                        <div style={{ textAlign: 'center', padding: '16px', color: '#94a3b8', fontSize: '0.85rem' }}>
                          No previous remarks for this lot. Add the first remark below!
                        </div>
                      );
                    }
                    return history.map((item, i) => (
                      <div
                        key={i}
                        className={`remarks-history-item ${i === history.length - 1 ? 'latest' : ''}`}
                      >
                        <div className="remarks-history-text">{item.text}</div>
                        <div className="remarks-history-time">
                          <span>🕒</span> {item.timestamp}
                          {i === history.length - 1 && (
                            <span style={{ marginLeft: 'auto', color: '#4338ca', fontWeight: '700', fontSize: '0.7rem', background: '#e0e7ff', padding: '2px 6px', borderRadius: '4px' }}>
                              Latest
                            </span>
                          )}
                        </div>
                      </div>
                    ));
                  })()}
                </div>
              </div>

              {/* Enter New Remark Input */}
              <div className="remarks-input-section">
                <h4 className="remarks-input-label">
                  <span>✏️</span> Enter Your Remark
                </h4>

                {/* Quick Presets */}
                <div className="remarks-quick-presets">
                  {[
                    "Ready for Packing",
                    "In Ironing / Folding",
                    "Packing Material Pending",
                    "Tag / Barcode Pending",
                    "Urgent Packing Required",
                    "Hold for Quality Check",
                    "Partial Quantity Packed",
                    "Completed & Ready for Dispatch"
                  ].map((preset) => (
                    <button
                      key={preset}
                      type="button"
                      className="remark-preset-tag"
                      onClick={() => setNewRemarkInputText(preset)}
                    >
                      + {preset}
                    </button>
                  ))}
                </div>

                <textarea
                  className="remarks-textarea"
                  placeholder="Type your custom remark here..."
                  value={newRemarkInputText}
                  onChange={(e) => setNewRemarkInputText(e.target.value)}
                  rows={3}
                  autoFocus
                />
              </div>
            </div>

            <div className="remarks-modal-footer">
              <button
                type="button"
                className="remarks-cancel-btn"
                onClick={handleCloseRemarksModal}
                disabled={savingRemark}
              >
                Cancel
              </button>
              <button
                type="button"
                className="remarks-save-btn"
                onClick={handleSaveRemark}
                disabled={savingRemark || !newRemarkInputText.trim()}
              >
                {savingRemark ? '⏳ Saving...' : '💾 Save Remark'}
              </button>
            </div>
          </div>
        </div>
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