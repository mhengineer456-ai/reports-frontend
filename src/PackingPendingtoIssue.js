import React, { useState, useEffect, useMemo, useCallback } from 'react';
import { useLocation } from 'react-router-dom';
import './PendingPackingtoIssue.css';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { GOOGLE_API_KEY, SPREADSHEET_IDS, SHEET_NAMES, fetchSheetDataFromBackend } from './config';
import { fetchRemarksForTab, saveRemarkForLot } from './embPrintRemarksService';

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
      <label style={{ display: "block", marginBottom: "5px", fontWeight: "600", fontSize: "0.85rem", color: "#475569" }}>
        {label}
      </label>
      <div
        onClick={() => setIsOpen(!isOpen)}
        style={{
          border: "1px solid #cbd5e1",
          borderRadius: "8px",
          padding: "8px 12px",
          background: "#ffffff",
          cursor: "pointer",
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          minHeight: "38px",
          fontSize: "0.85rem",
          color: isAllSelected ? "#64748b" : "#1e293b",
          fontWeight: isAllSelected ? "normal" : "600",
          boxShadow: "0 1px 2px 0 rgba(0, 0, 0, 0.05)"
        }}
      >
        <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", marginRight: "8px" }}>
          {getDisplayText()}
        </span>
        <span style={{ fontSize: "0.7rem", color: "#94a3b8" }}>▼</span>
      </div>

      {isOpen && (
        <div
          style={{
            position: "absolute",
            top: "100%",
            left: 0,
            right: 0,
            zIndex: 50,
            background: "#ffffff",
            border: "1px solid #e2e8f0",
            borderRadius: "8px",
            boxShadow: "0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05)",
            marginTop: "4px",
            maxHeight: "220px",
            overflowY: "auto",
            padding: "6px"
          }}
        >
          <label
            style={{
              display: "flex",
              alignItems: "center",
              gap: "8px",
              padding: "6px 8px",
              borderRadius: "4px",
              cursor: "pointer",
              fontSize: "0.85rem",
              fontWeight: isAllSelected ? "700" : "500",
              color: isAllSelected ? "#4338ca" : "#334155",
              background: isAllSelected ? "#e0e7ff" : "transparent"
            }}
          >
            <input
              type="checkbox"
              checked={isAllSelected}
              onChange={() => toggleOption("all")}
              style={{ cursor: "pointer", accentColor: "#4338ca" }}
            />
            <span>All</span>
          </label>
          {options.map((opt) => {
            const isChecked = selectedValues.includes(opt.value);
            return (
              <label
                key={opt.value}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: "8px",
                  padding: "6px 8px",
                  borderRadius: "4px",
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
  const location = useLocation();
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [packingData, setPackingData] = useState([]);
  const [issuesData, setIssuesData] = useState([]);
  const [issuesLotMap, setIssuesLotMap] = useState(new Map());
  const [rawpackData, setRawpackData] = useState([]);
  const [rawpackLotMap, setRawpackLotMap] = useState(new Map());
  const [processRunningLotsSet, setProcessRunningLotsSet] = useState(new Set());
  const [processRunningLotsMap, setProcessRunningLotsMap] = useState(new Map());
  const [lastProcessCompletedMap, setLastProcessCompletedMap] = useState(new Map());
  const [activeView, setActiveView] = useState(() => {
    const params = new URLSearchParams(window.location.search);
    return params.get('view') === 'pending_process' ? 'pending_process' : 'ready';
  });
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

  // Sync activeView with URL query parameter
  useEffect(() => {
    if (location.search) {
      const params = new URLSearchParams(location.search);
      const viewParam = params.get('view');
      if (viewParam === 'pending_process') {
        setActiveView('pending_process');
      } else if (viewParam === 'ready') {
        setActiveView('ready');
      }
    }
  }, [location.search]);

  // Fetch Remarks from Google Sheets & Subscribe to updates
  useEffect(() => {
    fetchRemarksForTab('PACKING').then(map => {
      if (map && typeof map === 'object') {
        setRemarksMap(map);
      }
    });

    const handleRemarkUpdated = (e) => {
      if (e.detail && (e.detail.tabType === 'PACKING' || !e.detail.tabType)) {
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
        tabType: 'PACKING',
        lotNumber: lotNumber,
        partyName: selectedRemarksLot.partyName || '',
        fabric: selectedRemarksLot.fabric || '',
        style: selectedRemarksLot.style || selectedRemarksLot.garmentType || '',
        remarkText: newRemarkInputText.trim()
      });

      if (updatedHistory) {
        setRemarksMap(prev => ({
          ...prev,
          [lotNumber]: updatedHistory
        }));
      }

      setNewRemarkInputText('');
      setRemarksModalOpen(false);
    } catch (err) {
      console.error('Error saving remark:', err);
      alert('Failed to save remark. Please try again.');
    } finally {
      setSavingRemark(false);
    }
  };

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

  // Finishing / Intermediate process sheets from OVERLOCK..FOLDING..KAJBUTTON spreadsheet
  const FINISHING_PROCESS_SHEETS = [
    'Overlock',
    'Filling',
    'Press',
    'FeedUp',
    'Jaybir Printing',
    'Jaybir Embroidery',
    'Washing',
    'Bone',
    'Elastic',
    'Folding',
    'KajButton'
  ];

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

  // Comprehensive date parser for ISO, DD/MM/YYYY, MM/DD/YYYY, DD-MM-YYYY, and Date objects
  const parseAnyDate = (dateInput) => {
    if (!dateInput || dateInput === '-' || dateInput === 'null' || dateInput === 'undefined') return null;
    if (dateInput instanceof Date) {
      return isNaN(dateInput.getTime()) ? null : dateInput;
    }
    const str = String(dateInput).trim();
    if (!str) return null;

    try {
      // 1. Check DD/MM/YYYY or MM/DD/YYYY with optional time
      const slashMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+.*)?$/);
      if (slashMatch) {
        let p1 = parseInt(slashMatch[1], 10);
        let p2 = parseInt(slashMatch[2], 10);
        let y = parseInt(slashMatch[3], 10);
        if (y < 100) y += 2000;

        let d = p1;
        let m = p2;
        // In India DD/MM/YYYY is standard. If p2 > 12, p2 is day and p1 is month.
        if (p2 > 12 && p1 <= 12) {
          m = p1;
          d = p2;
        }

        if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
          return new Date(y, m - 1, d);
        }
      }

      // 2. Check YYYY-MM-DD or YYYY/MM/DD with optional time
      const isoMatch = str.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[T\s].*)?$/);
      if (isoMatch) {
        const y = parseInt(isoMatch[1], 10);
        const m = parseInt(isoMatch[2], 10);
        const d = parseInt(isoMatch[3], 10);
        if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
          return new Date(y, m - 1, d);
        }
      }

      // 3. Check DD-MM-YYYY
      const dashMatch = str.match(/^(\d{1,2})-(\d{1,2})-(\d{2,4})(?:\s+.*)?$/);
      if (dashMatch) {
        let d = parseInt(dashMatch[1], 10);
        let m = parseInt(dashMatch[2], 10);
        let y = parseInt(dashMatch[3], 10);
        if (y < 100) y += 2000;
        if (m > 12 && d <= 12) {
          const temp = d;
          d = m;
          m = temp;
        }
        if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
          return new Date(y, m - 1, d);
        }
      }

      // 4. Native JS Date parse
      const nativeDate = new Date(str);
      if (!isNaN(nativeDate.getTime()) && nativeDate.getFullYear() > 1990) {
        return nativeDate;
      }
    } catch (e) {}

    return null;
  };

  const extractDateFromValue = (val, fallbackDate) => {
    if (!val && !fallbackDate) return null;
    const str = String(val || '').trim();

    if (str.startsWith('[') || str.startsWith('{')) {
      try {
        const parsed = JSON.parse(str);
        if (Array.isArray(parsed) && parsed.length > 0) {
          const last = parsed[parsed.length - 1];
          if (last) {
            const possible = last.completionDate || last.completedDate || last.timestamp || last.date || last.updatedAt;
            if (possible) {
              const d = parseAnyDate(possible);
              if (d) return { dateObj: d, dateStr: String(possible) };
            }
          }
        } else if (parsed && typeof parsed === 'object') {
          const possible = parsed.completionDate || parsed.completedDate || parsed.timestamp || parsed.date;
          if (possible) {
            const d = parseAnyDate(possible);
            if (d) return { dateObj: d, dateStr: String(possible) };
          }
        }
      } catch (e) {}
    }

    if (str && str !== '-' && str.toLowerCase() !== 'yes' && str.toLowerCase() !== 'completed' && str.toLowerCase() !== 'done') {
      const d = parseAnyDate(str);
      if (d) return { dateObj: d, dateStr: str };
    }

    if (fallbackDate) {
      const d = parseAnyDate(fallbackDate);
      if (d) return { dateObj: d, dateStr: String(fallbackDate) };
    }

    return null;
  };

  // Function to parse date from various formats
  const parseDate = (dateString) => {
    return parseAnyDate(dateString);
  };

  // Function to calculate pending days
  const calculatePendingDays = (completedDate) => {
    if (!completedDate || completedDate === '-') return 0;

    try {
      const completed = parseAnyDate(completedDate);
      if (!completed || isNaN(completed.getTime())) return 0;

      const today = new Date();
      today.setHours(0, 0, 0, 0);
      completed.setHours(0, 0, 0, 0);

      const diffTime = today - completed;
      const diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24));

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

  // Format process date to "14 Sept 2026"
  const formatProcessDate = (dateStr) => {
    if (!dateStr || dateStr === '-' || dateStr === 'null' || dateStr === 'undefined') return '';
    const str = String(dateStr).trim();
    if (!str) return '';

    const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sept', 'Oct', 'Nov', 'Dec'];

    try {
      // 1. Check for standard M/D/YYYY or MM/DD/YYYY with optional time (e.g., 9/14/2026 15:05:04 or 9/14/2026)
      const slashMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+.*)?$/);
      if (slashMatch) {
        let p1 = parseInt(slashMatch[1], 10);
        let p2 = parseInt(slashMatch[2], 10);
        let y = parseInt(slashMatch[3], 10);
        if (y < 100) y += 2000;

        let d = p2;
        let m = p1;

        // If p1 > 12, it must be DD/MM/YYYY
        if (p1 > 12) {
          d = p1;
          m = p2;
        }

        if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
          return `${d} ${monthNames[m - 1]} ${y}`;
        }
      }

      // 2. Check for YYYY-MM-DD or YYYY/MM/DD with optional time
      const isoMatch = str.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[T\s].*)?$/);
      if (isoMatch) {
        const y = parseInt(isoMatch[1], 10);
        const m = parseInt(isoMatch[2], 10);
        const d = parseInt(isoMatch[3], 10);
        if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
          return `${d} ${monthNames[m - 1]} ${y}`;
        }
      }

      // 3. Check for DD-MM-YYYY
      const dashMatch = str.match(/^(\d{1,2})-(\d{1,2})-(\d{2,4})(?:\s+.*)?$/);
      if (dashMatch) {
        let d = parseInt(dashMatch[1], 10);
        let m = parseInt(dashMatch[2], 10);
        let y = parseInt(dashMatch[3], 10);
        if (y < 100) y += 2000;
        if (m > 12 && d <= 12) {
          const temp = d;
          d = m;
          m = temp;
        }
        if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
          return `${d} ${monthNames[m - 1]} ${y}`;
        }
      }

      // 4. Fallback to Date object parsing
      const parsed = new Date(str);
      if (!isNaN(parsed.getTime())) {
        const d = parsed.getDate();
        const m = parsed.getMonth();
        const y = parsed.getFullYear();
        if (y > 1900) {
          return `${d} ${monthNames[m]} ${y}`;
        }
      }
    } catch (e) {
      console.warn('Error formatting process date:', e);
    }

    return str;
  };

  // Define getFilteredLots BEFORE any useMemo that depends on it
  const getFilteredLots = useCallback(() => {
    if (!packingData.length) return [];

    return packingData.filter(item => {
      const lotNumber = item.lotNumber?.toString().trim();
      if (!lotNumber) return false;
      const normalizedLot = lotNumber.toUpperCase();
      const cleanKey = normalizedLot.replace(/[^A-Z0-9]/g, '');

      const hasCompletedStatus = item.completedStatus && item.completedStatus !== '';
      const isNotInIssuesSheet = !issuesLotMap.has(lotNumber) && !issuesLotMap.has(normalizedLot);

      const rawpackInfo = rawpackLotMap.get(lotNumber) || rawpackLotMap.get(normalizedLot) || null;
      const isRawpackIssuedOrCompleted = rawpackInfo && (Boolean(rawpackInfo.packingPerson) || Boolean(rawpackInfo.packingIssueDate) || rawpackInfo.isCompleted);

      // Check if lot has running intermediate processes
      const isRunningInProcess = processRunningLotsSet.has(lotNumber) || processRunningLotsSet.has(normalizedLot) || (cleanKey && processRunningLotsSet.has(cleanKey));

      const isEligibleBase = hasCompletedStatus && isNotInIssuesSheet && !isRawpackIssuedOrCompleted;
      if (!isEligibleBase) return false;

      if (activeView === 'pending_process') {
        // Show lots that ARE running on any process
        return isRunningInProcess;
      } else {
        // Show lots that are NOT running on any process (Ready for packing)
        return !isRunningInProcess;
      }
    });
  }, [packingData, issuesLotMap, rawpackLotMap, processRunningLotsSet, activeView]);

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
      ],
      remarks: [
        { value: 'all', label: 'All Remarks' },
        { value: 'WITH_REMARKS', label: '📌 With Remarks Only' },
        { value: 'WITHOUT_REMARKS', label: '⚪ Without Remarks Only' },
        ...Array.from(
          new Set(
            Object.values(remarksMap)
              .flat()
              .map(r => r.text)
              .filter(Boolean)
          )
        ).map(text => ({ value: text, label: text }))
      ]
    };

    return options;
  }, [getFilteredLots, issuesLotMap, remarksMap]);

  // Always fetch fresh data on mount (bypassing stale localStorage)
  useEffect(() => {
    fetchAllSheetData();
  }, []);

  const fetchAllSheetData = async () => {
    try {
      setLoading(true);
      setError(null);

      const [indexData, issuesSheetData, rawpackSheetData, ...processSheetResults] = await Promise.all([
        fetchSheetData(SPREADSHEET_ID, RANGE),
        fetchSheetData(ISSUES_SPREADSHEET_ID, ISSUES_SPREADSHEET_RANGE),
        fetchSheetData(RAWPACK_SPREADSHEET_ID, RAWPACK_SPREADSHEET_RANGE).catch(err => {
          console.warn('Could not fetch RAWPACK sheet:', err);
          return { values: [] };
        }),
        ...FINISHING_PROCESS_SHEETS.map(sheetName =>
          fetchSheetData(SPREADSHEET_IDS.DAILY_STITCHING, `'${sheetName}'!A:Z`).catch(err => {
            console.warn(`Could not fetch ${sheetName} sheet:`, err);
            return { values: [] };
          })
        )
      ]);

      const transformedIndexData = transformSheetData(indexData.values || []);
      const { issuesData: transformedIssuesData, lotMap } = transformIssuesData(issuesSheetData.values || []);
      const { rawpackData: transformedRawpackData, rawpackLotMap: rawpackMap } = transformRawpackData(rawpackSheetData.values || []);
      const { runningLotsMap, runningLotsSet, lastProcessCompletedMap: procCompletedMap } = transformProcessSheetsData(processSheetResults, FINISHING_PROCESS_SHEETS);

      try {
        localStorage.setItem('packingData', JSON.stringify(transformedIndexData));
        localStorage.setItem('issuesData', JSON.stringify(transformedIssuesData));
        localStorage.setItem('issuesLotMap', JSON.stringify(Array.from(lotMap.entries())));
        localStorage.setItem('rawpackData', JSON.stringify(transformedRawpackData));
        localStorage.setItem('rawpackLotMap', JSON.stringify(Array.from(rawpackMap.entries())));
        localStorage.setItem('processRunningLotsSet', JSON.stringify(Array.from(runningLotsSet)));
        localStorage.setItem('packingDataTimestamp', new Date().getTime().toString());
      } catch (storageErr) {
        console.warn('LocalStorage quota exceeded. Skipping local storage persistence:', storageErr);
      }

      setPackingData(transformedIndexData);
      setIssuesData(transformedIssuesData);
      setIssuesLotMap(lotMap);
      setRawpackData(transformedRawpackData);
      setRawpackLotMap(rawpackMap);
      setProcessRunningLotsSet(runningLotsSet);
      setProcessRunningLotsMap(runningLotsMap);
      setLastProcessCompletedMap(procCompletedMap);
      setError(null);
    } catch (err) {
      setError(err.message);
      console.error('Error fetching data:', err);

      const cachedData = localStorage.getItem('packingData');
      const cachedIssuesData = localStorage.getItem('issuesData');
      const cachedIssuesMap = localStorage.getItem('issuesLotMap');
      const cachedRawpackMap = localStorage.getItem('rawpackLotMap');
      const cachedRunningLots = localStorage.getItem('processRunningLotsSet');

      if (cachedData && cachedIssuesData && cachedIssuesMap) {
        const parsedData = JSON.parse(cachedData);
        const parsedMap = new Map(JSON.parse(cachedIssuesMap));
        const parsedRawpackMap = cachedRawpackMap ? new Map(JSON.parse(cachedRawpackMap)) : new Map();
        const parsedRunningLots = cachedRunningLots ? new Set(JSON.parse(cachedRunningLots)) : new Set();

        const updatedData = parsedData.map(item => ({
          ...item,
          pendingDays: calculatePendingDays(item.completedStatusDisplay)
        }));

        setPackingData(updatedData);
        setIssuesData(JSON.parse(cachedIssuesData));
        setIssuesLotMap(parsedMap);
        setRawpackLotMap(parsedRawpackMap);
        setProcessRunningLotsSet(parsedRunningLots);
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

  const transformProcessSheetsData = (sheetResults, sheetNames) => {
    const runningLotsMap = new Map();
    const runningLotsSet = new Set();
    const lastProcessCompletedMap = new Map();

    sheetResults.forEach((sheetData, sIdx) => {
      const sheetName = sheetNames[sIdx];
      const values = sheetData?.values || [];
      if (!values || values.length === 0) return;

      const headers = (values[0] || []).map(h => String(h || '').trim().toLowerCase());
      const normK = (s) => String(s || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
      const findCol = (kws) => headers.findIndex(h => kws.some(k => normK(h).includes(normK(k))));

      const lotIdx = findCol(['lot number', 'lot no', 'lot #', 'lot']);
      const compIdx = findCol(['complete', 'completed', 'completion date', 'completion', 'complete status', 'status']);
      const wipIdx = findCol(['wip', 'remarks']);
      const dateIdx = findCol(['date', 'issue date', 'timestamp']);

      const actualLotIdx = lotIdx !== -1 ? lotIdx : 1;

      for (let i = 1; i < values.length; i++) {
        const row = values[i];
        if (!row || row.length === 0) continue;

        const rawLot = String(row[actualLotIdx] || row[0] || '').trim();
        if (!rawLot || rawLot === '-' || rawLot === '0') continue;

        const cleanLot = rawLot.toUpperCase();
        const cleanKey = cleanLot.replace(/[^A-Z0-9]/g, '');
        const rawComp = compIdx !== -1 && row[compIdx] ? String(row[compIdx]).trim() : '';
        const rawWip = wipIdx !== -1 && row[wipIdx] ? String(row[wipIdx]).trim() : '';
        const rawDate = dateIdx !== -1 && row[dateIdx] ? String(row[dateIdx]).trim() : '';

        // Determine if this process entry is completed
        let isCompleted = false;
        let completionDateResult = null;

        if (rawComp && rawComp !== '[]' && rawComp !== '-' && rawComp.toLowerCase() !== 'null' && rawComp.toLowerCase() !== 'undefined') {
          const compLower = rawComp.toLowerCase();
          if (rawComp.startsWith('[') || rawComp.startsWith('{')) {
            try {
              const parsed = JSON.parse(rawComp);
              if (Array.isArray(parsed) && parsed.length > 0) {
                const last = parsed[parsed.length - 1];
                if (last && (last.timestamp || last.date || last.status || last.completionDate || last.completedDate)) {
                  isCompleted = true;
                  completionDateResult = extractDateFromValue(rawComp, rawDate);
                }
              } else if (parsed && typeof parsed === 'object' && Object.keys(parsed).length > 0) {
                isCompleted = true;
                completionDateResult = extractDateFromValue(rawComp, rawDate);
              }
            } catch (e) {
              isCompleted = true;
              completionDateResult = extractDateFromValue(rawComp, rawDate);
            }
          } else if (!compLower.includes('pending') && !compLower.includes('not') && compLower !== 'no') {
            isCompleted = true;
            completionDateResult = extractDateFromValue(rawComp, rawDate);
          }
        }

        if (isCompleted) {
          const completedDateObj = completionDateResult?.dateObj || parseAnyDate(rawDate);
          if (completedDateObj && !isNaN(completedDateObj.getTime())) {
            const keysToSet = [cleanLot, rawLot, cleanKey].filter(Boolean);
            keysToSet.forEach(key => {
              const existing = lastProcessCompletedMap.get(key);
              if (!existing || completedDateObj.getTime() > existing.latestCompletedDate.getTime()) {
                lastProcessCompletedMap.set(key, {
                  latestCompletedDate: completedDateObj,
                  latestCompletedDateStr: completionDateResult?.dateStr || rawDate || formatDate(completedDateObj),
                  sheetName: sheetName
                });
              }
            });
          }
        } else {
          // If not completed, this lot has an active/running process in this sheet
          runningLotsSet.add(cleanLot);
          runningLotsSet.add(rawLot);
          if (cleanKey) runningLotsSet.add(cleanKey);

          const existing = runningLotsMap.get(cleanLot) || runningLotsMap.get(rawLot) || [];
          const formattedIssueDate = formatProcessDate(rawDate) || rawDate;
          const alreadyAdded = existing.some(e => e.sheetName === sheetName && e.issueDate === formattedIssueDate);
          if (!alreadyAdded) {
            existing.push({
              sheetName,
              issueDate: formattedIssueDate,
              wipRemark: rawWip
            });
          }
          runningLotsMap.set(cleanLot, existing);
          runningLotsMap.set(rawLot, existing);
          if (cleanKey) runningLotsMap.set(cleanKey, existing);
        }
      }
    });

    return { runningLotsMap, runningLotsSet, lastProcessCompletedMap };
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
      const normalizedLot = lotNumber ? lotNumber.toUpperCase() : '';
      const cleanKey = normalizedLot ? normalizedLot.replace(/[^A-Z0-9]/g, '') : '';
      const issuesInfo = issuesLotMap.get(lotNumber) || issuesLotMap.get(normalizedLot) || {};
      const runningProcesses = processRunningLotsMap.get(normalizedLot) || processRunningLotsMap.get(lotNumber) || (cleanKey ? processRunningLotsMap.get(cleanKey) : null) || [];

      const lastProcessInfo = lastProcessCompletedMap.get(normalizedLot) ||
                              lastProcessCompletedMap.get(lotNumber) ||
                              (cleanKey ? lastProcessCompletedMap.get(cleanKey) : null) ||
                              null;

      let effectivePendingDays = item.pendingDays || 0;
      let effectiveCompletedDateDisplay = item.completedStatusDisplay || '-';
      let hasIntermediateProcess = false;
      let lastProcessName = null;

      if (activeView === 'ready') {
        if (lastProcessInfo && lastProcessInfo.latestCompletedDate) {
          effectivePendingDays = calculatePendingDays(lastProcessInfo.latestCompletedDate);
          effectiveCompletedDateDisplay = formatDate(lastProcessInfo.latestCompletedDateStr || lastProcessInfo.latestCompletedDate);
          hasIntermediateProcess = true;
          lastProcessName = lastProcessInfo.sheetName;
        } else {
          effectivePendingDays = calculatePendingDays(item.completedStatusDisplay);
          effectiveCompletedDateDisplay = item.completedStatusDisplay || '-';
        }
      } else {
        effectivePendingDays = calculatePendingDays(item.completedStatusDisplay);
        effectiveCompletedDateDisplay = item.completedStatusDisplay || '-';
      }

      return {
        ...item,
        pendingDays: effectivePendingDays,
        effectiveCompletedDateDisplay,
        hasIntermediateProcess,
        lastProcessName,
        packingSupervisor: issuesInfo.packingSupervisor || '-',
        packingDate: issuesInfo.packingDate ? formatDate(issuesInfo.packingDate) : '-',
        packingComplete: issuesInfo.packingComplete || '-',
        runningProcesses: runningProcesses,
        runningProcessesText: runningProcesses.map(p => p.sheetName).join(', ')
      };
    });

    // Apply multi-select filters
    let filteredData = mergedData;

    if (filters.priority?.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.priority.includes(item.priority || 'Normal')
      );
    }

    if (filters.directStitching?.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.directStitching.includes(item.directStitching || 'No')
      );
    }

    if (filters.supervisor?.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.supervisor.includes(item.supervisor)
      );
    }

    if (filters.brand?.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.brand.includes(item.brand)
      );
    }

    if (filters.fabric?.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.fabric.includes(item.fabric)
      );
    }

    if (filters.garmentType?.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.garmentType.includes(item.garmentType)
      );
    }

    if (filters.style?.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.style.includes(item.style)
      );
    }

    if (filters.season?.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.season.includes(item.season)
      );
    }

    if (filters.partyName?.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.partyName.includes(item.partyName)
      );
    }

    // Apply remarks filter
    if (filters.remarks?.length > 0) {
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

    // Apply pending days range multi-filter
    if (filters.pendingDaysRange?.length > 0) {
      filteredData = filteredData.filter(item =>
        filters.pendingDaysRange.some(rangeVal => matchesPendingDaysRange(item.pendingDays || 0, rangeVal))
      );
    }

    // Apply search
    if (searchTerm) {
      const searchLower = searchTerm.toLowerCase();
      const fieldsToSearch = ['lotNumber', 'fabric', 'brand', 'garmentType', 'style', 'partyName', 'supervisor', 'season', 'mwk', 'priority'];

      filteredData = filteredData.filter(item => {
        const directMatch = fieldsToSearch.some(field =>
          item[field] && item[field].toString().toLowerCase().includes(searchLower)
        );
        if (directMatch) return true;

        const lotRemarks = remarksMap[item.lotNumber?.toString().trim()] || [];
        return lotRemarks.some(r => r.text && r.text.toLowerCase().includes(searchLower));
      });
    }

    return filteredData;
  }, [getFilteredLots, issuesLotMap, processRunningLotsMap, lastProcessCompletedMap, activeView, searchTerm, filters, remarksMap]);

  // Pagination
  const paginatedData = useMemo(() => {
    const startIndex = (currentPage - 1) * itemsPerPage;
    return displayData.slice(startIndex, startIndex + itemsPerPage);
  }, [displayData, currentPage, itemsPerPage]);

  const totalPages = Math.ceil(displayData.length / itemsPerPage);

  // Counts for KPI Cards
  const readyLotsCount = useMemo(() => {
    return packingData.filter(item => {
      const lotNumber = item.lotNumber?.toString().trim();
      if (!lotNumber) return false;
      const normalizedLot = lotNumber.toUpperCase();
      const cleanKey = normalizedLot.replace(/[^A-Z0-9]/g, '');
      const hasCompletedStatus = item.completedStatus && item.completedStatus !== '';
      const isNotInIssuesSheet = !issuesLotMap.has(lotNumber) && !issuesLotMap.has(normalizedLot);
      const rawpackInfo = rawpackLotMap.get(lotNumber) || rawpackLotMap.get(normalizedLot) || null;
      const isRawpackIssuedOrCompleted = rawpackInfo && (Boolean(rawpackInfo.packingPerson) || Boolean(rawpackInfo.packingIssueDate) || rawpackInfo.isCompleted);
      const isRunningInProcess = processRunningLotsSet.has(lotNumber) || processRunningLotsSet.has(normalizedLot) || (cleanKey && processRunningLotsSet.has(cleanKey));
      return hasCompletedStatus && isNotInIssuesSheet && !isRawpackIssuedOrCompleted && !isRunningInProcess;
    }).length;
  }, [packingData, issuesLotMap, rawpackLotMap, processRunningLotsSet]);

  const readyPiecesCount = useMemo(() => {
    return packingData.filter(item => {
      const lotNumber = item.lotNumber?.toString().trim();
      if (!lotNumber) return false;
      const normalizedLot = lotNumber.toUpperCase();
      const cleanKey = normalizedLot.replace(/[^A-Z0-9]/g, '');
      const hasCompletedStatus = item.completedStatus && item.completedStatus !== '';
      const isNotInIssuesSheet = !issuesLotMap.has(lotNumber) && !issuesLotMap.has(normalizedLot);
      const rawpackInfo = rawpackLotMap.get(lotNumber) || rawpackLotMap.get(normalizedLot) || null;
      const isRawpackIssuedOrCompleted = rawpackInfo && (Boolean(rawpackInfo.packingPerson) || Boolean(rawpackInfo.packingIssueDate) || rawpackInfo.isCompleted);
      const isRunningInProcess = processRunningLotsSet.has(lotNumber) || processRunningLotsSet.has(normalizedLot) || (cleanKey && processRunningLotsSet.has(cleanKey));
      return hasCompletedStatus && isNotInIssuesSheet && !isRawpackIssuedOrCompleted && !isRunningInProcess;
    }).reduce((sum, item) => sum + (item.stitchingIssueQtyNum || 0), 0);
  }, [packingData, issuesLotMap, rawpackLotMap, processRunningLotsSet]);

  const pendingProcessLotsCount = useMemo(() => {
    return packingData.filter(item => {
      const lotNumber = item.lotNumber?.toString().trim();
      if (!lotNumber) return false;
      const normalizedLot = lotNumber.toUpperCase();
      const cleanKey = normalizedLot.replace(/[^A-Z0-9]/g, '');
      const hasCompletedStatus = item.completedStatus && item.completedStatus !== '';
      const isNotInIssuesSheet = !issuesLotMap.has(lotNumber) && !issuesLotMap.has(normalizedLot);
      const rawpackInfo = rawpackLotMap.get(lotNumber) || rawpackLotMap.get(normalizedLot) || null;
      const isRawpackIssuedOrCompleted = rawpackInfo && (Boolean(rawpackInfo.packingPerson) || Boolean(rawpackInfo.packingIssueDate) || rawpackInfo.isCompleted);
      const isRunningInProcess = processRunningLotsSet.has(lotNumber) || processRunningLotsSet.has(normalizedLot) || (cleanKey && processRunningLotsSet.has(cleanKey));
      return hasCompletedStatus && isNotInIssuesSheet && !isRawpackIssuedOrCompleted && isRunningInProcess;
    }).length;
  }, [packingData, issuesLotMap, rawpackLotMap, processRunningLotsSet]);

  const pendingProcessPiecesCount = useMemo(() => {
    return packingData.filter(item => {
      const lotNumber = item.lotNumber?.toString().trim();
      if (!lotNumber) return false;
      const normalizedLot = lotNumber.toUpperCase();
      const cleanKey = normalizedLot.replace(/[^A-Z0-9]/g, '');
      const hasCompletedStatus = item.completedStatus && item.completedStatus !== '';
      const isNotInIssuesSheet = !issuesLotMap.has(lotNumber) && !issuesLotMap.has(normalizedLot);
      const rawpackInfo = rawpackLotMap.get(lotNumber) || rawpackLotMap.get(normalizedLot) || null;
      const isRawpackIssuedOrCompleted = rawpackInfo && (Boolean(rawpackInfo.packingPerson) || Boolean(rawpackInfo.packingIssueDate) || rawpackInfo.isCompleted);
      const isRunningInProcess = processRunningLotsSet.has(lotNumber) || processRunningLotsSet.has(normalizedLot) || (cleanKey && processRunningLotsSet.has(cleanKey));
      return hasCompletedStatus && isNotInIssuesSheet && !isRawpackIssuedOrCompleted && isRunningInProcess;
    }).reduce((sum, item) => sum + (item.stitchingIssueQtyNum || 0), 0);
  }, [packingData, issuesLotMap, rawpackLotMap, processRunningLotsSet]);

  // Process-Wise Breakdown for Pending Process view
  const processWiseSummary = useMemo(() => {
    if (activeView !== 'pending_process') return [];
    const map = {};
    displayData.forEach(item => {
      const qty = Number(item.stitchingIssueQtyNum) || Number(item.stitchingIssueQty) || 0;
      if (item.runningProcesses && item.runningProcesses.length > 0) {
        item.runningProcesses.forEach(p => {
          const name = p.sheetName || 'In Progress';
          if (!map[name]) map[name] = { name, lots: 0, qty: 0 };
          map[name].lots += 1;
          map[name].qty += qty;
        });
      } else {
        const name = 'In Progress';
        if (!map[name]) map[name] = { name, lots: 0, qty: 0 };
        map[name].lots += 1;
        map[name].qty += qty;
      }
    });
    return Object.values(map).sort((a, b) => b.qty - a.qty);
  }, [displayData, activeView]);

  // Executive summary aggregates for UI Breakdown Section
  const executiveSummaryData = useMemo(() => {
    if (!displayData || displayData.length === 0) return null;

    const totalLots = displayData.length;
    const totalQty = displayData.reduce((sum, item) => sum + (Number(item.stitchingIssueQtyNum) || Number(item.stitchingIssueQty) || 0), 0);

    const garmentMap = {};
    const supervisorMap = {};
    const seasonMap = {};
    const agingMap = {
      '0-7 Days': { name: '0-7 Days', lots: 0, qty: 0 },
      '8-15 Days': { name: '8-15 Days', lots: 0, qty: 0 },
      '16-30 Days': { name: '16-30 Days', lots: 0, qty: 0 },
      '30+ Days': { name: '30+ Days', lots: 0, qty: 0 }
    };

    displayData.forEach(item => {
      const qty = Number(item.stitchingIssueQtyNum) || Number(item.stitchingIssueQty) || 0;
      const gType = (item.garmentType || 'Unknown').trim();
      const sup = (item.supervisor || 'Unassigned').trim();
      const season = (item.season || 'N/A').trim();
      const days = item.pendingDays || 0;

      if (!garmentMap[gType]) garmentMap[gType] = { name: gType, lots: 0, qty: 0 };
      garmentMap[gType].lots += 1;
      garmentMap[gType].qty += qty;

      if (!supervisorMap[sup]) supervisorMap[sup] = { name: sup, lots: 0, qty: 0 };
      supervisorMap[sup].lots += 1;
      supervisorMap[sup].qty += qty;

      if (!seasonMap[season]) seasonMap[season] = { name: season, lots: 0, qty: 0 };
      seasonMap[season].lots += 1;
      seasonMap[season].qty += qty;

      if (days <= 7) {
        agingMap['0-7 Days'].lots += 1;
        agingMap['0-7 Days'].qty += qty;
      } else if (days <= 15) {
        agingMap['8-15 Days'].lots += 1;
        agingMap['8-15 Days'].qty += qty;
      } else if (days <= 30) {
        agingMap['16-30 Days'].lots += 1;
        agingMap['16-30 Days'].qty += qty;
      } else {
        agingMap['30+ Days'].lots += 1;
        agingMap['30+ Days'].qty += qty;
      }
    });

    return {
      totalLots,
      totalQty,
      garments: Object.values(garmentMap).sort((a, b) => b.qty - a.qty),
      supervisors: Object.values(supervisorMap).sort((a, b) => b.qty - a.qty),
      seasons: Object.values(seasonMap).sort((a, b) => b.qty - a.qty),
      aging: Object.values(agingMap)
    };
  }, [displayData]);

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
      brand: [],
      fabric: [],
      garmentType: [],
      style: [],
      season: [],
      pendingDaysRange: [],
      partyName: [],
      remarks: []
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

  // Professional PDF Export with centered cells, pure black text, embedded pictures, total row, and executive summary
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
      const totalStitchingQty = displayData.reduce((sum, item) => sum + (Number(item.stitchingIssueQtyNum) || Number(item.stitchingIssueQty) || 0), 0);
      const highPriorityCount = displayData.filter(item => (item.priority || '').toLowerCase() === 'high').length;
      const directLotsCount = displayData.filter(item => (item.directStitching || '').toLowerCase() === 'yes').length;
      const avgPendingDays = totalLots > 0 ? Math.round(displayData.reduce((sum, item) => sum + (item.pendingDays || 0), 0) / totalLots) : 0;

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
      const processMap = {};
      const garmentMap = {};
      const seasonMap = {};
      const supervisorMap = {};
      const agingMap = {
        '0-7 Days': { totalLots: 0, totalQty: 0 },
        '8-15 Days': { totalLots: 0, totalQty: 0 },
        '16-30 Days': { totalLots: 0, totalQty: 0 },
        '30+ Days': { totalLots: 0, totalQty: 0 }
      };

      displayData.forEach(item => {
        const qty = Number(item.stitchingIssueQtyNum) || Number(item.stitchingIssueQty) || 0;
        const gType = (item.garmentType || 'Unknown').trim();
        const season = (item.season || 'N/A').trim();
        const sup = (item.supervisor || 'Unassigned').trim();
        const days = item.pendingDays || 0;

        if (item.runningProcesses && item.runningProcesses.length > 0) {
          item.runningProcesses.forEach(p => {
            const procName = (p.sheetName || 'In Progress').trim();
            if (!processMap[procName]) processMap[procName] = { totalLots: 0, totalQty: 0 };
            processMap[procName].totalLots += 1;
            processMap[procName].totalQty += qty;
          });
        } else {
          const procName = 'In Progress';
          if (!processMap[procName]) processMap[procName] = { totalLots: 0, totalQty: 0 };
          processMap[procName].totalLots += 1;
          processMap[procName].totalQty += qty;
        }

        if (!garmentMap[gType]) garmentMap[gType] = { totalLots: 0, totalQty: 0 };
        garmentMap[gType].totalLots += 1;
        garmentMap[gType].totalQty += qty;

        if (!seasonMap[season]) seasonMap[season] = { totalLots: 0, totalQty: 0 };
        seasonMap[season].totalLots += 1;
        seasonMap[season].totalQty += qty;

        if (!supervisorMap[sup]) supervisorMap[sup] = { totalLots: 0, totalQty: 0 };
        supervisorMap[sup].totalLots += 1;
        supervisorMap[sup].totalQty += qty;

        if (days <= 7) {
          agingMap['0-7 Days'].totalLots += 1;
          agingMap['0-7 Days'].totalQty += qty;
        } else if (days <= 15) {
          agingMap['8-15 Days'].totalLots += 1;
          agingMap['8-15 Days'].totalQty += qty;
        } else if (days <= 30) {
          agingMap['16-30 Days'].totalLots += 1;
          agingMap['16-30 Days'].totalQty += qty;
        } else {
          agingMap['30+ Days'].totalLots += 1;
          agingMap['30+ Days'].totalQty += qty;
        }
      });

      const sortedProcesses = Object.keys(processMap).map(k => ({ name: k, totalLots: processMap[k].totalLots, totalQty: processMap[k].totalQty })).sort((a, b) => b.totalQty - a.totalQty);
      const sortedGarments = Object.keys(garmentMap).map(k => ({ name: k, totalLots: garmentMap[k].totalLots, totalQty: garmentMap[k].totalQty })).sort((a, b) => b.totalQty - a.totalQty);
      const sortedSeasons = Object.keys(seasonMap).map(k => ({ name: k, totalLots: seasonMap[k].totalLots, totalQty: seasonMap[k].totalQty })).sort((a, b) => b.totalQty - a.totalQty);
      const sortedSupervisors = Object.keys(supervisorMap).map(k => ({ name: k, totalLots: supervisorMap[k].totalLots, totalQty: supervisorMap[k].totalQty })).sort((a, b) => b.totalQty - a.totalQty);

      const isPendingProcess = activeView === 'pending_process';

      // Title Banner
      doc.setFillColor(isPendingProcess ? 30 : 15, isPendingProcess ? 27 : 23, isPendingProcess ? 75 : 42); // Theme Navy / Dark Indigo
      doc.rect(margin, 12, contentWidth, 48, 'F');

      doc.setFontSize(14);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(255, 255, 255);
      const titleMain = isPendingProcess
        ? "MH FACTORY SUITE PRO - PENDING PROCESS AFTER STITCHING DONE REPORT"
        : "MH FACTORY SUITE PRO - COMPLETED LOTS READY FOR PACKING REPORT";
      doc.text(titleMain, pageWidth / 2, 30, { align: 'center' });

      doc.setFontSize(8.5);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(isPendingProcess ? 254 : 199, isPendingProcess ? 240 : 210, isPendingProcess ? 138 : 254);
      const subText = isPendingProcess
        ? `Pending Process Lots: ${totalLots}   |   Total Stitching Qty: ${totalStitchingQty.toLocaleString()}   |   Running on Process: ${totalLots} Lots   |   Direct Lots: ${directLotsCount}   |   High Priority: ${highPriorityCount}   |   Avg Pending Days: ${avgPendingDays}d   |   Generated: ${new Date().toLocaleDateString('en-IN')} ${new Date().toLocaleTimeString('en-IN')}`
        : `Completed Lots: ${totalLots}   |   Total Stitching Qty: ${totalStitchingQty.toLocaleString()}   |   Direct Lots: ${directLotsCount}   |   High Priority: ${highPriorityCount}   |   Avg Pending Days: ${avgPendingDays}d   |   Generated: ${new Date().toLocaleDateString('en-IN')} ${new Date().toLocaleTimeString('en-IN')}`;
      doc.text(subText, pageWidth / 2, 48, { align: 'center' });

      // Filter Banner
      doc.setFillColor(241, 245, 249);
      doc.rect(margin, 63, contentWidth, 16, 'F');
      doc.setFontSize(8);
      doc.setFont('helvetica', 'italic');
      doc.setTextColor(0, 0, 0);
      const filterSummary = `Filters: Priority: ${filters.priority.length ? filters.priority.join(', ') : 'All'} | Direct: ${filters.directStitching.length ? filters.directStitching.join(', ') : 'All'} | Supervisor: ${filters.supervisor.length ? filters.supervisor.join(', ') : 'All'} | Brand: ${filters.brand.length ? filters.brand.join(', ') : 'All'} | Fabric: ${filters.fabric.length ? filters.fabric.join(', ') : 'All'} | Garment: ${filters.garmentType.length ? filters.garmentType.join(', ') : 'All'} | Style: ${filters.style.length ? filters.style.join(', ') : 'All'} | Season: ${filters.season.length ? filters.season.join(', ') : 'All'} | Remarks: ${filters.remarks && filters.remarks.length ? filters.remarks.join(', ') : 'All'}`;
      doc.text(filterSummary, pageWidth / 2, 74, { align: 'center' });

      // Table columns & rows
      const tableColumns = isPendingProcess
        ? [
            '#',
            'Image',
            'Lot Number',
            'Garment Type',
            'Style',
            'Fabric',
            'Brand',
            'Stitching Qty',
            'M/W/K',
            'Season',
            'Party Name',
            'Direct Stitching',
            'Supervisor',
            'Date of Issue',
            'Priority',
            'Completed Date',
            'Pending Days',
            'Running Process',
            'Remarks'
          ]
        : [
            '#',
            'Image',
            'Lot Number',
            'Garment Type',
            'Style',
            'Fabric',
            'Brand',
            'Stitching Qty',
            'M/W/K',
            'Season',
            'Party Name',
            'Direct Stitching',
            'Supervisor',
            'Date of Issue',
            'Priority',
            'Completed Date',
            'Pending Days',
            'Remarks'
          ];

      const tableBody = displayData.map((item, idx) => {
        const lotNo = (item.lotNumber || '').toString().trim();
        const lotRemarks = remarksMap[lotNo] || [];
        const latestRemark = lotRemarks.length > 0 ? lotRemarks[lotRemarks.length - 1].text : '—';
        const qty = Number(item.stitchingIssueQtyNum) || Number(item.stitchingIssueQty) || 0;
        const isDirect = (item.directStitching || '').toLowerCase() === 'yes';

        if (isPendingProcess) {
          const runningProcText = (item.runningProcesses && item.runningProcesses.length > 0)
            ? item.runningProcesses.map(p => `${p.sheetName}${p.issueDate ? ` (${formatProcessDate(p.issueDate)})` : ''}`).join(', ')
            : 'In Progress';

          return [
            (idx + 1).toString(),
            '', // Image cell rendered via didDrawCell
            lotNo || '—',
            item.garmentType || '—',
            item.style || '—',
            item.fabric || '—',
            item.brand || '—',
            qty.toLocaleString(),
            item.mwk || '—',
            item.season || '—',
            item.partyName || '—',
            isDirect ? 'Yes' : 'No',
            item.supervisor || '—',
            item.dateOfIssue || '—',
            item.priority || 'Normal',
            item.completedStatusDisplay || '—',
            `${item.pendingDays || 0}d`,
            runningProcText,
            latestRemark
          ];
        }

        return [
          (idx + 1).toString(),
          '', // Image cell rendered via didDrawCell
          lotNo || '—',
          item.garmentType || '—',
          item.style || '—',
          item.fabric || '—',
          item.brand || '—',
          qty.toLocaleString(),
          item.mwk || '—',
          item.season || '—',
          item.partyName || '—',
          isDirect ? 'Yes' : 'No',
          item.supervisor || '—',
          item.dateOfIssue || '—',
          item.priority || 'Normal',
          item.effectiveCompletedDateDisplay || item.completedStatusDisplay || '—',
          `${item.pendingDays || 0}d`,
          latestRemark
        ];
      });

      // Total Row
      if (isPendingProcess) {
        tableBody.push([
          '',
          '',
          `TOTAL (${totalLots} Lots)`,
          '',
          '',
          '',
          '',
          totalStitchingQty.toLocaleString(),
          '',
          '',
          '',
          `${directLotsCount} Direct`,
          '',
          '',
          `${highPriorityCount} High`,
          '',
          `${avgPendingDays}d Avg`,
          'Running on Process',
          ''
        ]);
      } else {
        tableBody.push([
          '',
          '',
          `TOTAL (${totalLots} Lots)`,
          '',
          '',
          '',
          '',
          totalStitchingQty.toLocaleString(),
          '',
          '',
          '',
          `${directLotsCount} Direct`,
          '',
          '',
          `${highPriorityCount} High`,
          '',
          `${avgPendingDays}d Avg`,
          ''
        ]);
      }

      const baseWidths = isPendingProcess
        ? {
            0: 22,   // #
            1: 34,   // Image
            2: 52,   // Lot Number
            3: 65,   // Garment Type
            4: 65,   // Style
            5: 75,   // Fabric
            6: 62,   // Brand
            7: 54,   // Stitching Qty
            8: 44,   // M/W/K
            9: 50,   // Season
            10: 78,  // Party Name
            11: 48,  // Direct Stitching
            12: 68,  // Supervisor
            13: 60,  // Date of Issue
            14: 48,  // Priority
            15: 66,  // Completed Date
            16: 50,  // Pending Days
            17: 85,  // Running Process
            18: 130  // Remarks
          }
        : {
            0: 24,   // #
            1: 36,   // Image
            2: 55,   // Lot Number
            3: 72,   // Garment Type
            4: 72,   // Style
            5: 82,   // Fabric
            6: 68,   // Brand
            7: 58,   // Stitching Qty
            8: 48,   // M/W/K
            9: 54,   // Season
            10: 85,  // Party Name
            11: 52,  // Direct Stitching
            12: 74,  // Supervisor
            13: 65,  // Date of Issue
            14: 52,  // Priority
            15: 72,  // Completed Date
            16: 54,  // Pending Days
            17: 150  // Remarks
          };

      const sumBase = Object.values(baseWidths).reduce((a, b) => a + b, 0);
      const scale = contentWidth / sumBase;

      const columnStyles = {};
      Object.keys(baseWidths).forEach(k => {
        const w = baseWidths[k] * scale;
        columnStyles[k] = {
          cellWidth: w,
          halign: 'center',
          fontStyle: (k === '2' || k === '7' || k === '16') ? 'bold' : 'normal'
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
          fillColor: isPendingProcess ? [30, 27, 75] : [15, 23, 42],
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
            if (data.column.index === 11 && (row.directStitching || '').toLowerCase() === 'yes') {
              data.cell.styles.textColor = [21, 128, 61];
              data.cell.styles.fontStyle = 'bold';
            }

            // High Priority styling
            if (data.column.index === 14 && (row.priority || '').toLowerCase() === 'high') {
              data.cell.styles.textColor = [220, 38, 38];
              data.cell.styles.fontStyle = 'bold';
            }

            // Aging days styling
            if (data.column.index === 16) {
              const days = row.pendingDays || 0;
              if (days > 15) {
                data.cell.styles.textColor = [220, 38, 38];
                data.cell.styles.fontStyle = 'bold';
              } else if (days > 7) {
                data.cell.styles.textColor = [217, 119, 6];
              } else {
                data.cell.styles.textColor = [22, 163, 74];
              }
            }

            // Running Process styling in PDF
            if (isPendingProcess && data.column.index === 17) {
              data.cell.styles.textColor = [180, 83, 9];
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

      // --- SIDE-BY-SIDE EXECUTIVE SUMMARY TABLES ---
      const procTotalLots = sortedProcesses.reduce((s, p) => s + p.totalLots, 0);
      const procTotalQty = sortedProcesses.reduce((s, p) => s + p.totalQty, 0);

      const procBody = sortedProcesses.map(item => [
        item.name,
        item.totalLots.toString(),
        item.totalQty.toLocaleString(),
        procTotalQty > 0 ? `${((item.totalQty / procTotalQty) * 100).toFixed(1)}%` : '0%'
      ]);
      procBody.push(['TOTAL', procTotalLots.toString(), procTotalQty.toLocaleString(), '100%']);

      const gBody = sortedGarments.map(item => [
        item.name,
        item.totalLots.toString(),
        item.totalQty.toLocaleString(),
        totalStitchingQty > 0 ? `${((item.totalQty / totalStitchingQty) * 100).toFixed(1)}%` : '0%'
      ]);
      gBody.push(['TOTAL', totalLots.toString(), totalStitchingQty.toLocaleString(), '100%']);

      const sBody = sortedSeasons.map(item => [
        item.name,
        item.totalLots.toString(),
        item.totalQty.toLocaleString(),
        totalStitchingQty > 0 ? `${((item.totalQty / totalStitchingQty) * 100).toFixed(1)}%` : '0%'
      ]);
      sBody.push(['TOTAL', totalLots.toString(), totalStitchingQty.toLocaleString(), '100%']);

      const supBody = sortedSupervisors.map(item => [
        item.name,
        item.totalLots.toString(),
        item.totalQty.toLocaleString(),
        totalStitchingQty > 0 ? `${((item.totalQty / totalStitchingQty) * 100).toFixed(1)}%` : '0%'
      ]);
      supBody.push(['TOTAL', totalLots.toString(), totalStitchingQty.toLocaleString(), '100%']);

      const ageBody = Object.keys(agingMap).map(k => [
        k,
        agingMap[k].totalLots.toString(),
        agingMap[k].totalQty.toLocaleString(),
        totalStitchingQty > 0 ? `${((agingMap[k].totalQty / totalStitchingQty) * 100).toFixed(1)}%` : '0%'
      ]);
      ageBody.push(['TOTAL', totalLots.toString(), totalStitchingQty.toLocaleString(), '100%']);

      let finalY = doc.lastAutoTable ? doc.lastAutoTable.finalY + 20 : 300;
      if (finalY + 160 > pageHeight - 40) {
        doc.addPage();
        finalY = 30;
      }

      const summaryStartY = finalY;

      // Executive Summary Header Banner
      doc.setFillColor(15, 23, 42);
      doc.rect(margin, summaryStartY, pageWidth - (margin * 2), 20, 'F');
      doc.setFontSize(10);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(255, 255, 255);
      const sumHeading = isPendingProcess
        ? "EXECUTIVE SUMMARY & DISTRIBUTION (PENDING PROCESS AFTER STITCHING)"
        : "EXECUTIVE PRODUCTION SUMMARY & DISTRIBUTION";
      doc.text(sumHeading, pageWidth / 2, summaryStartY + 14, { align: 'center' });

      const sumTableY = summaryStartY + 26;
      const totalAvailWidth = pageWidth - (margin * 2);
      const numCols = isPendingProcess ? 5 : 4;
      const colGap = 10;
      const singleTableWidth = (totalAvailWidth - (colGap * (numCols - 1))) / numCols;

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

      let tableColIdx = 0;

      // 1. Process / Department Table (When Pending Process)
      if (isPendingProcess) {
        autoTable(doc, {
          head: [['Process / Dept', 'Lots', 'Qty', '%']],
          body: procBody,
          startY: sumTableY,
          margin: { left: margin + (singleTableWidth + colGap) * tableColIdx },
          tableWidth: singleTableWidth,
          theme: 'grid',
          styles: summaryBodyStyles,
          headStyles: { ...summaryHeadStyles, fillColor: [180, 83, 9] },
          didParseCell: (d) => {
            if (d.section === 'body' && d.row.index === procBody.length - 1) {
              d.cell.styles.fontStyle = 'bold';
              d.cell.styles.fillColor = [226, 232, 240];
            }
          }
        });
        tableColIdx++;
      }

      // 2. Garment Table
      autoTable(doc, {
        head: [['Garment Type', 'Lots', 'Qty', '%']],
        body: gBody,
        startY: sumTableY,
        margin: { left: margin + (singleTableWidth + colGap) * tableColIdx },
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
      tableColIdx++;

      // 3. Season Table
      autoTable(doc, {
        head: [['Season', 'Lots', 'Qty', '%']],
        body: sBody,
        startY: sumTableY,
        margin: { left: margin + (singleTableWidth + colGap) * tableColIdx },
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
      tableColIdx++;

      // 4. Supervisor Table
      autoTable(doc, {
        head: [['Supervisor', 'Lots', 'Qty', '%']],
        body: supBody,
        startY: sumTableY,
        margin: { left: margin + (singleTableWidth + colGap) * tableColIdx },
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
      tableColIdx++;

      // 5. Aging Table
      autoTable(doc, {
        head: [['Aging Range', 'Lots', 'Qty', '%']],
        body: ageBody,
        startY: sumTableY,
        margin: { left: margin + (singleTableWidth + colGap) * tableColIdx },
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
      const filename = isPendingProcess
        ? `Pending_Process_After_Stitching_${dateStr}.pdf`
        : `Completed_Lots_Ready_Packing_${dateStr}.pdf`;
      doc.save(filename);

    } catch (error) {
      console.error('Error exporting to PDF:', error);
      alert('Error exporting to PDF. Please try again.');
    } finally {
      setExportLoading(false);
    }
  }, [displayData, remarksMap, filters, activeView]);

  // Professional Multi-Sheet Excel Export with Executive Summary & Applied Filters
  const exportToExcel = useCallback(async () => {
    if (displayData.length === 0) {
      alert('No data to export');
      return;
    }

    try {
      setExportLoading(true);
      const isPendingProcess = activeView === 'pending_process';
      const workbook = new ExcelJS.Workbook();
      workbook.creator = 'MH Factory Suite Pro';
      workbook.created = new Date();

      const totalLots = displayData.length;
      const totalStitchingQty = displayData.reduce((sum, item) => sum + (Number(item.stitchingIssueQtyNum) || Number(item.stitchingIssueQty) || 0), 0);
      const highPriorityCount = displayData.filter(item => (item.priority || '').toLowerCase() === 'high').length;
      const directLotsCount = displayData.filter(item => (item.directStitching || '').toLowerCase() === 'yes').length;
      const avgPendingDays = totalLots > 0 ? Math.round(displayData.reduce((sum, item) => sum + (item.pendingDays || 0), 0) / totalLots) : 0;

      // Aggregations for Executive Summary
      const processMap = {};
      const garmentMap = {};
      const seasonMap = {};
      const supervisorMap = {};
      const partyMap = {};
      const agingMap = {
        '0-7 Days': { totalLots: 0, totalQty: 0 },
        '8-15 Days': { totalLots: 0, totalQty: 0 },
        '16-30 Days': { totalLots: 0, totalQty: 0 },
        '30+ Days': { totalLots: 0, totalQty: 0 }
      };

      displayData.forEach(item => {
        const qty = Number(item.stitchingIssueQtyNum) || Number(item.stitchingIssueQty) || 0;
        const gType = (item.garmentType || 'Unknown').trim();
        const season = (item.season || 'N/A').trim();
        const sup = (item.supervisor || 'Unassigned').trim();
        const party = (item.partyName || (item.directStitching === 'yes' ? 'Direct Stitching' : '—')).trim();
        const days = item.pendingDays || 0;

        if (item.runningProcesses && item.runningProcesses.length > 0) {
          item.runningProcesses.forEach(p => {
            const procName = (p.sheetName || 'In Progress').trim();
            if (!processMap[procName]) processMap[procName] = { totalLots: 0, totalQty: 0 };
            processMap[procName].totalLots += 1;
            processMap[procName].totalQty += qty;
          });
        } else {
          const procName = 'In Progress';
          if (!processMap[procName]) processMap[procName] = { totalLots: 0, totalQty: 0 };
          processMap[procName].totalLots += 1;
          processMap[procName].totalQty += qty;
        }

        if (!garmentMap[gType]) garmentMap[gType] = { totalLots: 0, totalQty: 0 };
        garmentMap[gType].totalLots += 1;
        garmentMap[gType].totalQty += qty;

        if (!seasonMap[season]) seasonMap[season] = { totalLots: 0, totalQty: 0 };
        seasonMap[season].totalLots += 1;
        seasonMap[season].totalQty += qty;

        if (!supervisorMap[sup]) supervisorMap[sup] = { totalLots: 0, totalQty: 0 };
        supervisorMap[sup].totalLots += 1;
        supervisorMap[sup].totalQty += qty;

        if (!partyMap[party]) partyMap[party] = { totalLots: 0, totalQty: 0 };
        partyMap[party].totalLots += 1;
        partyMap[party].totalQty += qty;

        if (days <= 7) {
          agingMap['0-7 Days'].totalLots += 1;
          agingMap['0-7 Days'].totalQty += qty;
        } else if (days <= 15) {
          agingMap['8-15 Days'].totalLots += 1;
          agingMap['8-15 Days'].totalQty += qty;
        } else if (days <= 30) {
          agingMap['16-30 Days'].totalLots += 1;
          agingMap['16-30 Days'].totalQty += qty;
        } else {
          agingMap['30+ Days'].totalLots += 1;
          agingMap['30+ Days'].totalQty += qty;
        }
      });

      const sortedProcesses = Object.keys(processMap).map(k => ({ name: k, totalLots: processMap[k].totalLots, totalQty: processMap[k].totalQty })).sort((a, b) => b.totalQty - a.totalQty);
      const sortedGarments = Object.keys(garmentMap).map(k => ({ name: k, totalLots: garmentMap[k].totalLots, totalQty: garmentMap[k].totalQty })).sort((a, b) => b.totalQty - a.totalQty);
      const sortedSeasons = Object.keys(seasonMap).map(k => ({ name: k, totalLots: seasonMap[k].totalLots, totalQty: seasonMap[k].totalQty })).sort((a, b) => b.totalQty - a.totalQty);
      const sortedSupervisors = Object.keys(supervisorMap).map(k => ({ name: k, totalLots: supervisorMap[k].totalLots, totalQty: supervisorMap[k].totalQty })).sort((a, b) => b.totalQty - a.totalQty);
      const sortedParties = Object.keys(partyMap).map(k => ({ name: k, totalLots: partyMap[k].totalLots, totalQty: partyMap[k].totalQty })).sort((a, b) => b.totalQty - a.totalQty);

      const thinBorder = {
        top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
      };

      // ================= SHEET 1: DATA TABLE =================
      const sheetName = isPendingProcess ? 'Pending Process After Stitch' : 'Completed Lots Ready';
      const ws1 = workbook.addWorksheet(sheetName, { views: [{ showGridLines: true }] });

      // Title Banner
      ws1.mergeCells('A1:R1');
      const titleCell = ws1.getCell('A1');
      titleCell.value = isPendingProcess
        ? 'MH FACTORY SUITE PRO - PENDING PROCESS AFTER STITCHING DONE REPORT'
        : 'MH FACTORY SUITE PRO - COMPLETED LOTS READY FOR PACKING REPORT';
      titleCell.font = { name: 'Segoe UI', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
      titleCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: isPendingProcess ? 'FF78350F' : 'FF1E1B4B' }
      };
      titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
      ws1.getRow(1).height = 30;

      // Subtitle KPI Banner
      ws1.mergeCells('A2:R2');
      const subCell = ws1.getCell('A2');
      subCell.value = isPendingProcess
        ? `Pending Process Lots: ${totalLots}   |   Total Stitching Qty: ${totalStitchingQty.toLocaleString()}   |   Running on Process: ${totalLots} Lots   |   Direct Lots: ${directLotsCount}   |   High Priority: ${highPriorityCount}   |   Avg Pending Days: ${avgPendingDays}d   |   Generated: ${new Date().toLocaleDateString('en-IN')} ${new Date().toLocaleTimeString('en-IN')}`
        : `Completed Lots: ${totalLots}   |   Total Stitching Qty: ${totalStitchingQty.toLocaleString()}   |   Direct Lots: ${directLotsCount}   |   High Priority: ${highPriorityCount}   |   Avg Pending Days: ${avgPendingDays}d   |   Generated: ${new Date().toLocaleDateString('en-IN')} ${new Date().toLocaleTimeString('en-IN')}`;
      subCell.font = { name: 'Segoe UI', size: 9.5, color: { argb: isPendingProcess ? 'FFFEF3C7' : 'FFC7D2FE' } };
      subCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: isPendingProcess ? 'FF92400E' : 'FF312E81' }
      };
      subCell.alignment = { horizontal: 'center', vertical: 'middle' };
      ws1.getRow(2).height = 22;

      // Table Header Row
      const tableHeaders = isPendingProcess
        ? [
            '#', 'Lot Number', 'Garment Type', 'Style', 'Fabric', 'Brand',
            'Stitching Qty', 'M/W/K', 'Season', 'Party Name', 'Direct Stitching',
            'Supervisor', 'Date of Issue', 'Priority', 'Completed Date',
            'Pending Days', 'Running Process', 'Remarks'
          ]
        : [
            '#', 'Lot Number', 'Garment Type', 'Style', 'Fabric', 'Brand',
            'Stitching Qty', 'M/W/K', 'Season', 'Party Name', 'Direct Stitching',
            'Supervisor', 'Date of Issue', 'Priority', 'Completed Date',
            'Pending Days', 'Status', 'Remarks'
          ];

      const headerRow = ws1.addRow(tableHeaders);
      headerRow.height = 24;
      headerRow.eachCell(cell => {
        cell.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: isPendingProcess ? 'FF1E293B' : 'FF0F172A' }
        };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = thinBorder;
      });

      // Data Rows
      displayData.forEach((item, index) => {
        const lotNo = (item.lotNumber || '').toString().trim();
        const lotRemarks = remarksMap[lotNo] || [];
        const latestRemark = lotRemarks.length > 0 ? lotRemarks[lotRemarks.length - 1].text : '';
        const qty = Number(item.stitchingIssueQtyNum) || Number(item.stitchingIssueQty) || 0;
        const days = item.pendingDays || 0;
        const isDirect = (item.directStitching || '').toLowerCase() === 'yes';

        const runningProcText = (item.runningProcesses && item.runningProcesses.length > 0)
          ? item.runningProcesses.map(p => `${p.sheetName}${p.issueDate ? ` (${formatProcessDate(p.issueDate)})` : ''}`).join(', ')
          : 'In Progress';

        const col17Value = isPendingProcess ? runningProcText : 'Ready for Packing';

        const r = ws1.addRow([
          index + 1,
          lotNo,
          item.garmentType || '—',
          item.style || '—',
          item.fabric || '—',
          item.brand || '—',
          qty,
          item.mwk || '—',
          item.season || '—',
          item.partyName || '—',
          isDirect ? 'Yes' : 'No',
          item.supervisor || '—',
          item.dateOfIssue || '—',
          item.priority || 'Normal',
          item.effectiveCompletedDateDisplay || item.completedStatusDisplay || '—',
          days,
          col17Value,
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

        // Stitching Qty styling
        r.getCell(7).numFmt = '#,##0';
        r.getCell(7).font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FFDC2626' } };

        // Days Aging formatting
        const daysCell = r.getCell(16);
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

        // Running Process styling in Excel
        if (isPendingProcess) {
          r.getCell(17).font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FFB45309' } };
        }
      });

      // Total Row
      const totalCol17 = isPendingProcess ? 'Running on Process' : 'Ready for Packing';
      const totalRow1 = ws1.addRow([
        '', `TOTAL (${totalLots} Lots)`, '', '', '', '',
        totalStitchingQty, '', '', '', `${directLotsCount} Direct`,
        '', '', `${highPriorityCount} High`, '',
        `${avgPendingDays}d Avg`, totalCol17, ''
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
      const colWidths = isPendingProcess
        ? [6, 15, 18, 18, 20, 16, 14, 10, 12, 22, 16, 18, 15, 12, 18, 14, 26, 30]
        : [6, 15, 18, 18, 20, 16, 14, 10, 12, 22, 16, 18, 15, 12, 18, 14, 18, 30];

      colWidths.forEach((w, i) => {
        ws1.getColumn(i + 1).width = w;
      });

      // ================= SHEET 2: EXECUTIVE SUMMARY =================
      const ws2 = workbook.addWorksheet('Executive Summary', { views: [{ showGridLines: true }] });

      // 0. Process / Department Breakdown (When in Pending Process mode)
      if (isPendingProcess) {
        const procTotalLots = sortedProcesses.reduce((s, p) => s + p.totalLots, 0);
        const procTotalQty = sortedProcesses.reduce((s, p) => s + p.totalQty, 0);

        ws2.mergeCells('A1:D1');
        const procTitle = ws2.getCell('A1');
        procTitle.value = '1. PROCESS / FINISHING DEPARTMENT WISE BREAKDOWN (LOTS & PCS)';
        procTitle.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        procTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB45309' } };
        procTitle.alignment = { horizontal: 'left', vertical: 'middle' };
        ws2.getRow(1).height = 26;

        const procHeader = ws2.addRow(['Process / Department', 'Total Lots', 'Total Stitching Qty', 'Share %']);
        procHeader.height = 22;
        procHeader.eachCell(c => {
          c.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF78350F' } };
          c.alignment = { horizontal: 'center', vertical: 'middle' };
          c.border = thinBorder;
        });

        sortedProcesses.forEach((item, idx) => {
          const pct = procTotalQty > 0 ? (item.totalQty / procTotalQty) : 0;
          const r = ws2.addRow([item.name, item.totalLots, item.totalQty, pct]);
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

        const procTotalRow = ws2.addRow(['TOTAL', procTotalLots, procTotalQty, 1]);
        procTotalRow.height = 22;
        procTotalRow.eachCell(c => {
          c.font = { name: 'Segoe UI', size: 9.5, bold: true };
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
          c.alignment = { horizontal: 'center', vertical: 'middle' };
          c.border = { top: { style: 'thin' }, bottom: { style: 'double' }, left: { style: 'thin' }, right: { style: 'thin' } };
        });
        procTotalRow.getCell(3).numFmt = '#,##0';
        procTotalRow.getCell(4).numFmt = '0.0%';

        ws2.addRow([]);
      }

      // Next Section: Garment Type Breakdown
      const gStartRow = ws2.rowCount + 1;
      ws2.mergeCells(`A${gStartRow}:D${gStartRow}`);
      const gTitle = ws2.getCell(`A${gStartRow}`);
      gTitle.value = isPendingProcess
        ? '2. GARMENT TYPE BREAKDOWN (PENDING PROCESS AFTER STITCHING)'
        : '1. GARMENT TYPE BREAKDOWN (LOTS & STITCHING QTY DISTRIBUTION)';
      gTitle.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      gTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F766E' } };
      gTitle.alignment = { horizontal: 'left', vertical: 'middle' };
      ws2.getRow(gStartRow).height = 26;

      const gHeader = ws2.addRow(['Garment Type', 'Total Lots', 'Total Stitching Qty', 'Share %']);
      gHeader.height = 22;
      gHeader.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF134E4A' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = thinBorder;
      });

      sortedGarments.forEach((item, idx) => {
        const pct = totalStitchingQty > 0 ? (item.totalQty / totalStitchingQty) : 0;
        const r = ws2.addRow([item.name, item.totalLots, item.totalQty, pct]);
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

      const gTotalRow = ws2.addRow(['TOTAL', totalLots, totalStitchingQty, 1]);
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
      sTitle.value = isPendingProcess
        ? '2. SEASON WISE BREAKDOWN (PENDING PROCESS AFTER STITCHING)'
        : '2. SEASON WISE BREAKDOWN (LOTS & STITCHING QTY DISTRIBUTION)';
      sTitle.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      sTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4338CA' } };
      sTitle.alignment = { horizontal: 'left', vertical: 'middle' };
      ws2.getRow(sStartRow).height = 26;

      const sHeader = ws2.addRow(['Season', 'Total Lots', 'Total Stitching Qty', 'Share %']);
      sHeader.height = 22;
      sHeader.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF312E81' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = thinBorder;
      });

      sortedSeasons.forEach((item, idx) => {
        const pct = totalStitchingQty > 0 ? (item.totalQty / totalStitchingQty) : 0;
        const r = ws2.addRow([item.name, item.totalLots, item.totalQty, pct]);
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

      const sTotalRow = ws2.addRow(['TOTAL', totalLots, totalStitchingQty, 1]);
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

      // 3. Supervisor Breakdown
      const supStartRow = ws2.rowCount + 1;
      ws2.mergeCells(`A${supStartRow}:D${supStartRow}`);
      const supTitle = ws2.getCell(`A${supStartRow}`);
      supTitle.value = isPendingProcess
        ? '3. SUPERVISOR WISE BREAKDOWN (PENDING PROCESS AFTER STITCHING)'
        : '3. SUPERVISOR WISE BREAKDOWN (LOTS & STITCHING QTY DISTRIBUTION)';
      supTitle.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      supTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E40AF' } };
      supTitle.alignment = { horizontal: 'left', vertical: 'middle' };
      ws2.getRow(supStartRow).height = 26;

      const supHeader = ws2.addRow(['Supervisor', 'Total Lots', 'Total Stitching Qty', 'Share %']);
      supHeader.height = 22;
      supHeader.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF172554' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = thinBorder;
      });

      sortedSupervisors.forEach((item, idx) => {
        const pct = totalStitchingQty > 0 ? (item.totalQty / totalStitchingQty) : 0;
        const r = ws2.addRow([item.name, item.totalLots, item.totalQty, pct]);
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

      const supTotalRow = ws2.addRow(['TOTAL', totalLots, totalStitchingQty, 1]);
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

      // 4. Party Wise Breakdown
      const pStartRow = ws2.rowCount + 1;
      ws2.mergeCells(`A${pStartRow}:D${pStartRow}`);
      const pTitle = ws2.getCell(`A${pStartRow}`);
      pTitle.value = isPendingProcess
        ? '4. PARTY WISE BREAKDOWN (PENDING PROCESS AFTER STITCHING)'
        : '4. PARTY WISE BREAKDOWN (LOTS & STITCHING QTY DISTRIBUTION)';
      pTitle.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      pTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9333EA' } };
      pTitle.alignment = { horizontal: 'left', vertical: 'middle' };
      ws2.getRow(pStartRow).height = 26;

      const pHeader = ws2.addRow(['Party Name', 'Total Lots', 'Total Stitching Qty', 'Share %']);
      pHeader.height = 22;
      pHeader.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF581C87' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = thinBorder;
      });

      sortedParties.forEach((item, idx) => {
        const pct = totalStitchingQty > 0 ? (item.totalQty / totalStitchingQty) : 0;
        const r = ws2.addRow([item.name, item.totalLots, item.totalQty, pct]);
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

      const pTotalRow = ws2.addRow(['TOTAL', totalLots, totalStitchingQty, 1]);
      pTotalRow.height = 22;
      pTotalRow.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = { top: { style: 'thin' }, bottom: { style: 'double' }, left: { style: 'thin' }, right: { style: 'thin' } };
      });
      pTotalRow.getCell(3).numFmt = '#,##0';
      pTotalRow.getCell(4).numFmt = '0.0%';

      // Spacer
      ws2.addRow([]);

      // 5. Aging Breakdown
      const ageStartRow = ws2.rowCount + 1;
      ws2.mergeCells(`A${ageStartRow}:D${ageStartRow}`);
      const ageTitle = ws2.getCell(`A${ageStartRow}`);
      ageTitle.value = isPendingProcess
        ? '5. AGING WISE BREAKDOWN (PENDING PROCESS AFTER STITCHING)'
        : '5. AGING WISE BREAKDOWN (PENDING DAYS)';
      ageTitle.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      ageTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF991B1B' } };
      ageTitle.alignment = { horizontal: 'left', vertical: 'middle' };
      ws2.getRow(ageStartRow).height = 26;

      const ageHeader = ws2.addRow(['Aging Range', 'Total Lots', 'Total Stitching Qty', 'Share %']);
      ageHeader.height = 22;
      ageHeader.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7F1D1D' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = thinBorder;
      });

      Object.keys(agingMap).forEach((k, idx) => {
        const item = agingMap[k];
        const pct = totalStitchingQty > 0 ? (item.totalQty / totalStitchingQty) : 0;
        const r = ws2.addRow([k, item.totalLots, item.totalQty, pct]);
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

      const ageTotalRow = ws2.addRow(['TOTAL', totalLots, totalStitchingQty, 1]);
      ageTotalRow.height = 22;
      ageTotalRow.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = { top: { style: 'thin' }, bottom: { style: 'double' }, left: { style: 'thin' }, right: { style: 'thin' } };
      });
      ageTotalRow.getCell(3).numFmt = '#,##0';
      ageTotalRow.getCell(4).numFmt = '0.0%';

      // Set column widths for summary sheet
      ws2.getColumn(1).width = 28;
      ws2.getColumn(2).width = 18;
      ws2.getColumn(3).width = 24;
      ws2.getColumn(4).width = 16;

      // Save file
      const dateStr = new Date().toISOString().split('T')[0];
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const excelFilename = isPendingProcess
        ? `Pending_Process_After_Stitching_${dateStr}.xlsx`
        : `Completed_Lots_Ready_Packing_${dateStr}.xlsx`;
      saveAs(blob, excelFilename);

    } catch (error) {
      console.error('Error exporting to Excel:', error);
      alert('Error exporting to Excel. Please try again.');
    } finally {
      setExportLoading(false);
    }
  }, [displayData, remarksMap, activeView]);

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
          <h2>{activeView === 'pending_process' ? '⚙️ Pending Process after Stitching Done' : '✅ Completed Lots Ready for Packing'}</h2>
          <div className="stats-badge">
            Total Records: {displayData.length}
          </div>
        </div>

        {/* Interactive KPI Cards Header Grid */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: '16px', marginBottom: '20px' }}>
          {/* Card 1: Ready for Packing */}
          <div
            onClick={() => { setActiveView('ready'); setCurrentPage(1); }}
            style={{
              background: activeView === 'ready' ? 'linear-gradient(135deg, #1e1b4b 0%, #312e81 50%, #4338ca 100%)' : 'rgba(255, 255, 255, 0.1)',
              border: activeView === 'ready' ? '2.5px solid #818cf8' : '1px solid rgba(255, 255, 255, 0.2)',
              borderRadius: '20px',
              padding: '18px 22px',
              cursor: 'pointer',
              transition: 'all 0.25s ease',
              boxShadow: activeView === 'ready' ? '0 12px 30px rgba(49, 46, 129, 0.45)' : 'none',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
              transform: activeView === 'ready' ? 'scale(1.02)' : 'scale(1)'
            }}
          >
            <div>
              <div style={{ fontSize: '0.8rem', color: activeView === 'ready' ? '#c7d2fe' : '#94a3b8', fontWeight: 800, textTransform: 'uppercase', letterSpacing: '0.05em' }}>
                Ready for Packing
              </div>
              <div style={{ fontSize: '2rem', fontWeight: 800, color: '#ffffff', marginTop: '4px', lineHeight: 1.1 }}>
                {readyLotsCount} <span style={{ fontSize: '0.95rem', fontWeight: 600, color: '#a5b4fc' }}>Lots</span>
              </div>
              <div style={{ fontSize: '0.82rem', color: activeView === 'ready' ? '#e0e7ff' : '#cbd5e1', marginTop: '4px', fontWeight: 500 }}>
                {readyPiecesCount.toLocaleString()} Pcs • All Processes Done
              </div>
            </div>
            <div style={{ fontSize: '2.4rem', opacity: activeView === 'ready' ? 1 : 0.6 }}>📦</div>
          </div>

          {/* Card 2: Pending Process after Stitching Done */}
          <div
            onClick={() => { setActiveView('pending_process'); setCurrentPage(1); }}
            style={{
              background: activeView === 'pending_process' ? 'linear-gradient(135deg, #78350f 0%, #92400e 50%, #b45309 100%)' : 'rgba(255, 255, 255, 0.1)',
              border: activeView === 'pending_process' ? '2.5px solid #fbbf24' : '1px solid rgba(255, 255, 255, 0.2)',
              borderRadius: '20px',
              padding: '18px 22px',
              cursor: 'pointer',
              transition: 'all 0.25s ease',
              boxShadow: activeView === 'pending_process' ? '0 12px 30px rgba(180, 83, 9, 0.45)' : 'none',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
              transform: activeView === 'pending_process' ? 'scale(1.02)' : 'scale(1)'
            }}
          >
            <div>
              <div style={{ fontSize: '0.8rem', color: activeView === 'pending_process' ? '#fde68a' : '#94a3b8', fontWeight: 800, textTransform: 'uppercase', letterSpacing: '0.05em' }}>
                Pending Process after stitching Done
              </div>
              <div style={{ fontSize: '2rem', fontWeight: 800, color: '#ffffff', marginTop: '4px', lineHeight: 1.1 }}>
                {pendingProcessLotsCount} <span style={{ fontSize: '0.95rem', fontWeight: 600, color: '#fef08a' }}>Lots</span>
              </div>
              <div style={{ fontSize: '0.82rem', color: activeView === 'pending_process' ? '#fef3c7' : '#cbd5e1', marginTop: '4px', fontWeight: 500 }}>
                {pendingProcessPiecesCount.toLocaleString()} Pcs • Running on Process
              </div>
            </div>
            <div style={{ fontSize: '2.4rem', opacity: activeView === 'pending_process' ? 1 : 0.6 }}>⚙️</div>
          </div>
        </div>

        {/* Process-Wise Department Breakdown Strip (When in Pending Process view) */}
        {activeView === 'pending_process' && processWiseSummary.length > 0 && (
          <div style={{
            display: 'flex',
            gap: '12px',
            flexWrap: 'wrap',
            marginBottom: '20px',
            background: 'rgba(255, 255, 255, 0.08)',
            padding: '14px 18px',
            borderRadius: '16px',
            border: '1px solid rgba(251, 191, 36, 0.25)',
            alignItems: 'center'
          }}>
            <span style={{ fontSize: '0.85rem', fontWeight: 800, color: '#fde68a', display: 'flex', alignItems: 'center', gap: '6px', textTransform: 'uppercase', letterSpacing: '0.05em' }}>
              <span>⚙️</span> Department / Process Breakdown:
            </span>
            <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap', flex: 1 }}>
              {processWiseSummary.map((proc, pIdx) => (
                <div
                  key={pIdx}
                  style={{
                    background: 'rgba(254, 243, 199, 0.15)',
                    border: '1px solid #fbbf24',
                    borderRadius: '10px',
                    padding: '6px 14px',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '8px',
                    color: '#ffffff'
                  }}
                >
                  <span style={{ fontWeight: 800, color: '#fef08a', fontSize: '0.88rem' }}>{proc.name}</span>
                  <span style={{ background: '#78350f', color: '#fef3c7', padding: '2px 8px', borderRadius: '9999px', fontSize: '0.78rem', fontWeight: 700 }}>
                    {proc.lots} {proc.lots === 1 ? 'Lot' : 'Lots'}
                  </span>
                  <span style={{ color: '#fed7aa', fontSize: '0.82rem', fontWeight: 600 }}>
                    {proc.qty.toLocaleString()} Pcs
                  </span>
                </div>
              ))}
            </div>
          </div>
        )}

        <div className="actions">
          <input
            type="text"
            placeholder="🔍 Search by Lot, Fabric, Brand, Style, Party, Supervisor..."
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

            <MultiSelectDropdown
              label="Remarks:"
              options={filterOptions.remarks}
              selectedValues={filters.remarks}
              onChange={(newVals) => {
                setFilters(prev => ({ ...prev, remarks: newVals }));
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
                  <th>Garment Type</th>
                  <th>Style</th>
                  <th>Fabric</th>
                  <th>Brand</th>
                  <th>Stitching Qty</th>
                  <th>M/W/K</th>
                  <th>Season</th>
                  <th>Party Name</th>
                  <th>Direct Stitching</th>
                  <th>Supervisor</th>
                  <th>Date of Issue</th>
                  <th>Priority</th>
                  <th>{activeView === 'ready' ? 'Completed Date' : 'Stitching Completed Date'}</th>
                  <th>Pending Days</th>
                  {activeView === 'pending_process' && <th>Running Process</th>}
                  <th>💬 Remarks</th>
                </tr>
              </thead>
              <tbody>
                {paginatedData.map((item, index) => {
                  const lotRemarks = remarksMap[item.lotNumber?.toString().trim()] || [];
                  const latestRemark = lotRemarks.length > 0 ? lotRemarks[lotRemarks.length - 1] : null;

                  return (
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
                      <td>{item.garmentType || '-'}</td>
                      <td>{item.style || '-'}</td>
                      <td>{item.fabric || '-'}</td>
                      <td>{item.brand || '-'}</td>
                      <td className="quantity-cell">
                        <span className={item.stitchingIssueQtyNum > 0 ? 'positive-qty' : 'zero-qty'}>
                          {item.stitchingIssueQty || '0'}
                        </span>
                      </td>
                      <td>{item.mwk || '-'}</td>
                      <td>{item.season || '-'}</td>
                      <td>{item.partyName || '-'}</td>
                      <td>
                        <span className={`status-badge ${item.directStitching?.toLowerCase() === 'yes' ? 'status-yes' : 'status-no'}`}>
                          {item.directStitching || '-'}
                        </span>
                      </td>
                      <td>{item.supervisor || '-'}</td>
                      <td>{item.dateOfIssue || '-'}</td>
                      <td>
                        <span className={`priority-badge priority-${(item.priority || 'normal').toLowerCase()}`}>
                          {item.priority || 'Normal'}
                        </span>
                      </td>
                      <td>
                        <span className="status-badge status-completed">
                          {item.effectiveCompletedDateDisplay || item.completedStatusDisplay || '-'}
                        </span>
                        {activeView === 'ready' && item.hasIntermediateProcess && item.lastProcessName && (
                          <span style={{ fontSize: '0.72rem', color: '#059669', display: 'block', fontWeight: 700, marginTop: '2px' }}>
                            ({item.lastProcessName} Done)
                          </span>
                        )}
                      </td>
                      <td>
                        <span className={`pending-days-badge ${getPendingDaysClass(item.pendingDays)}`}>
                          {item.pendingDays || 0} {item.pendingDays === 1 ? 'day' : 'days'}
                        </span>
                      </td>
                      {activeView === 'pending_process' && (
                        <td>
                          {item.runningProcesses && item.runningProcesses.length > 0 ? (
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', alignItems: 'center' }}>
                              {item.runningProcesses.map((p, pIdx) => (
                                <span
                                  key={pIdx}
                                  style={{
                                    background: '#fef3c7',
                                    color: '#92400e',
                                    border: '1px solid #fde68a',
                                    padding: '2px 8px',
                                    borderRadius: '6px',
                                    fontSize: '0.78rem',
                                    fontWeight: 700,
                                    whiteSpace: 'nowrap'
                                  }}
                                >
                                  ⚙️ {p.sheetName}
                                  {p.issueDate && <span style={{ fontSize: '0.7rem', color: '#78350f', marginLeft: '4px' }}>({formatProcessDate(p.issueDate)})</span>}
                                </span>
                              ))}
                            </div>
                          ) : (
                            <span style={{ color: '#d97706', fontSize: '0.8rem', fontWeight: 600 }}>In Progress</span>
                          )}
                        </td>
                      )}
                      <td className="remarks-table-cell">
                        <div className="remarks-cell-container">
                          {latestRemark ? (
                            <div
                              className="remark-bubble"
                              onClick={() => handleOpenRemarksModal(item)}
                              title="Click to view history or add new remark"
                            >
                              <div className="remark-text-preview">
                                {latestRemark.text}
                              </div>
                              <div className="remark-meta-preview">
                                <span>🕒 {latestRemark.timestamp}</span>
                              </div>
                            </div>
                          ) : (
                            <span className="no-remarks-placeholder">No remarks yet</span>
                          )}

                          <div className="remarks-actions-row">
                            <button
                              className="btn-add-remark-inline"
                              onClick={() => handleOpenRemarksModal(item)}
                              title="Add or update remark for this lot"
                            >
                              {latestRemark ? '✏️ Remark' : '+ Add Remark'}
                            </button>

                            {lotRemarks.length > 1 && (
                              <button
                                className="btn-remark-history-badge"
                                onClick={() => handleOpenRemarksModal(item)}
                                title="View all remarks history"
                              >
                                📜 ({lotRemarks.length})
                              </button>
                            )}
                          </div>
                        </div>
                      </td>
                    </tr>
                  );
                })}
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

          {/* Executive Summary & Distribution Section in UI */}
          {executiveSummaryData && (
            <div style={{ marginTop: '36px', marginBottom: '40px' }}>
              <div style={{
                background: 'linear-gradient(135deg, #0f172a 0%, #1e293b 100%)',
                padding: '14px 20px',
                borderRadius: '16px 16px 0 0',
                border: '1px solid #334155',
                borderBottom: 'none',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between'
              }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                  <span style={{ fontSize: '1.2rem' }}>📊</span>
                  <h3 style={{ margin: 0, fontSize: '1.05rem', fontWeight: 800, color: '#f8fafc', letterSpacing: '0.04em', textTransform: 'uppercase' }}>
                    {activeView === 'pending_process'
                      ? 'Executive Summary & Department Distribution (Pending Process After Stitching)'
                      : 'Executive Production Summary & Distribution (Ready for Packing)'}
                  </h3>
                </div>
                <div style={{ fontSize: '0.85rem', color: '#94a3b8', fontWeight: 600 }}>
                  Total: <strong style={{ color: '#ffffff' }}>{executiveSummaryData.totalLots} Lots</strong> • <strong style={{ color: '#38bdf8' }}>{executiveSummaryData.totalQty.toLocaleString()} Pcs</strong>
                </div>
              </div>

              <div style={{
                display: 'grid',
                gridTemplateColumns: activeView === 'pending_process' ? 'repeat(auto-fit, minmax(260px, 1fr))' : 'repeat(auto-fit, minmax(280px, 1fr))',
                gap: '16px',
                background: '#0f172a',
                padding: '18px',
                borderRadius: '0 0 16px 16px',
                border: '1px solid #334155',
                borderTop: 'none'
              }}>
                {/* 1. Process / Department Breakdown (When Pending Process) */}
                {activeView === 'pending_process' && processWiseSummary.length > 0 && (
                  <div style={{ background: '#1e293b', borderRadius: '12px', overflow: 'hidden', border: '1px solid #475569' }}>
                    <div style={{ background: '#b45309', padding: '8px 12px', color: '#ffffff', fontWeight: 800, fontSize: '0.85rem', textAlign: 'center', letterSpacing: '0.03em' }}>
                      ⚙️ PROCESS / DEPT
                    </div>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.82rem', color: '#f8fafc' }}>
                      <thead>
                        <tr style={{ background: '#78350f', color: '#fef3c7', fontSize: '0.78rem' }}>
                          <th style={{ padding: '6px 8px', textAlign: 'left', borderBottom: '1px solid #475569' }}>Department</th>
                          <th style={{ padding: '6px 8px', textAlign: 'center', borderBottom: '1px solid #475569' }}>Lots</th>
                          <th style={{ padding: '6px 8px', textAlign: 'center', borderBottom: '1px solid #475569' }}>Qty</th>
                          <th style={{ padding: '6px 8px', textAlign: 'center', borderBottom: '1px solid #475569' }}>%</th>
                        </tr>
                      </thead>
                      <tbody>
                        {processWiseSummary.map((p, idx) => {
                          const procTotalQty = processWiseSummary.reduce((s, x) => s + x.qty, 0);
                          const pct = procTotalQty > 0 ? ((p.qty / procTotalQty) * 100).toFixed(1) : '0';
                          return (
                            <tr key={idx} style={{ background: idx % 2 === 1 ? 'rgba(255,255,255,0.03)' : 'transparent', borderBottom: '1px solid #334155' }}>
                              <td style={{ padding: '6px 8px', fontWeight: 700, color: '#fde68a' }}>{p.name}</td>
                              <td style={{ padding: '6px 8px', textAlign: 'center' }}>{p.lots}</td>
                              <td style={{ padding: '6px 8px', textAlign: 'center', fontWeight: 700, color: '#ffffff' }}>{p.qty.toLocaleString()}</td>
                              <td style={{ padding: '6px 8px', textAlign: 'center', color: '#94a3b8' }}>{pct}%</td>
                            </tr>
                          );
                        })}
                        <tr style={{ background: '#334155', fontWeight: 800, color: '#ffffff' }}>
                          <td style={{ padding: '6px 8px' }}>TOTAL</td>
                          <td style={{ padding: '6px 8px', textAlign: 'center' }}>{processWiseSummary.reduce((s, x) => s + x.lots, 0)}</td>
                          <td style={{ padding: '6px 8px', textAlign: 'center', color: '#fde68a' }}>{processWiseSummary.reduce((s, x) => s + x.qty, 0).toLocaleString()}</td>
                          <td style={{ padding: '6px 8px', textAlign: 'center' }}>100%</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                )}

                {/* 2. Garment Breakdown */}
                <div style={{ background: '#1e293b', borderRadius: '12px', overflow: 'hidden', border: '1px solid #475569' }}>
                  <div style={{ background: '#0f766e', padding: '8px 12px', color: '#ffffff', fontWeight: 800, fontSize: '0.85rem', textAlign: 'center', letterSpacing: '0.03em' }}>
                    👕 GARMENT TYPE
                  </div>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.82rem', color: '#f8fafc' }}>
                    <thead>
                      <tr style={{ background: '#134e4a', color: '#ccfbf1', fontSize: '0.78rem' }}>
                        <th style={{ padding: '6px 8px', textAlign: 'left', borderBottom: '1px solid #475569' }}>Garment</th>
                        <th style={{ padding: '6px 8px', textAlign: 'center', borderBottom: '1px solid #475569' }}>Lots</th>
                        <th style={{ padding: '6px 8px', textAlign: 'center', borderBottom: '1px solid #475569' }}>Qty</th>
                        <th style={{ padding: '6px 8px', textAlign: 'center', borderBottom: '1px solid #475569' }}>%</th>
                      </tr>
                    </thead>
                    <tbody>
                      {executiveSummaryData.garments.map((g, idx) => {
                        const pct = executiveSummaryData.totalQty > 0 ? ((g.qty / executiveSummaryData.totalQty) * 100).toFixed(1) : '0';
                        return (
                          <tr key={idx} style={{ background: idx % 2 === 1 ? 'rgba(255,255,255,0.03)' : 'transparent', borderBottom: '1px solid #334155' }}>
                            <td style={{ padding: '6px 8px', fontWeight: 600 }}>{g.name}</td>
                            <td style={{ padding: '6px 8px', textAlign: 'center' }}>{g.lots}</td>
                            <td style={{ padding: '6px 8px', textAlign: 'center', fontWeight: 700, color: '#ffffff' }}>{g.qty.toLocaleString()}</td>
                            <td style={{ padding: '6px 8px', textAlign: 'center', color: '#94a3b8' }}>{pct}%</td>
                          </tr>
                        );
                      })}
                      <tr style={{ background: '#334155', fontWeight: 800, color: '#ffffff' }}>
                        <td style={{ padding: '6px 8px' }}>TOTAL</td>
                        <td style={{ padding: '6px 8px', textAlign: 'center' }}>{executiveSummaryData.totalLots}</td>
                        <td style={{ padding: '6px 8px', textAlign: 'center', color: '#5eead4' }}>{executiveSummaryData.totalQty.toLocaleString()}</td>
                        <td style={{ padding: '6px 8px', textAlign: 'center' }}>100%</td>
                      </tr>
                    </tbody>
                  </table>
                </div>

                {/* 3. Season Breakdown */}
                <div style={{ background: '#1e293b', borderRadius: '12px', overflow: 'hidden', border: '1px solid #475569' }}>
                  <div style={{ background: '#4338ca', padding: '8px 12px', color: '#ffffff', fontWeight: 800, fontSize: '0.85rem', textAlign: 'center', letterSpacing: '0.03em' }}>
                    ☀️ SEASON
                  </div>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.82rem', color: '#f8fafc' }}>
                    <thead>
                      <tr style={{ background: '#312e81', color: '#e0e7ff', fontSize: '0.78rem' }}>
                        <th style={{ padding: '6px 8px', textAlign: 'left', borderBottom: '1px solid #475569' }}>Season</th>
                        <th style={{ padding: '6px 8px', textAlign: 'center', borderBottom: '1px solid #475569' }}>Lots</th>
                        <th style={{ padding: '6px 8px', textAlign: 'center', borderBottom: '1px solid #475569' }}>Qty</th>
                        <th style={{ padding: '6px 8px', textAlign: 'center', borderBottom: '1px solid #475569' }}>%</th>
                      </tr>
                    </thead>
                    <tbody>
                      {executiveSummaryData.seasons.map((s, idx) => {
                        const pct = executiveSummaryData.totalQty > 0 ? ((s.qty / executiveSummaryData.totalQty) * 100).toFixed(1) : '0';
                        return (
                          <tr key={idx} style={{ background: idx % 2 === 1 ? 'rgba(255,255,255,0.03)' : 'transparent', borderBottom: '1px solid #334155' }}>
                            <td style={{ padding: '6px 8px', fontWeight: 600 }}>{s.name}</td>
                            <td style={{ padding: '6px 8px', textAlign: 'center' }}>{s.lots}</td>
                            <td style={{ padding: '6px 8px', textAlign: 'center', fontWeight: 700, color: '#ffffff' }}>{s.qty.toLocaleString()}</td>
                            <td style={{ padding: '6px 8px', textAlign: 'center', color: '#94a3b8' }}>{pct}%</td>
                          </tr>
                        );
                      })}
                      <tr style={{ background: '#334155', fontWeight: 800, color: '#ffffff' }}>
                        <td style={{ padding: '6px 8px' }}>TOTAL</td>
                        <td style={{ padding: '6px 8px', textAlign: 'center' }}>{executiveSummaryData.totalLots}</td>
                        <td style={{ padding: '6px 8px', textAlign: 'center', color: '#a5b4fc' }}>{executiveSummaryData.totalQty.toLocaleString()}</td>
                        <td style={{ padding: '6px 8px', textAlign: 'center' }}>100%</td>
                      </tr>
                    </tbody>
                  </table>
                </div>

                {/* 4. Supervisor Breakdown */}
                <div style={{ background: '#1e293b', borderRadius: '12px', overflow: 'hidden', border: '1px solid #475569' }}>
                  <div style={{ background: '#1e40af', padding: '8px 12px', color: '#ffffff', fontWeight: 800, fontSize: '0.85rem', textAlign: 'center', letterSpacing: '0.03em' }}>
                    👤 SUPERVISOR
                  </div>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.82rem', color: '#f8fafc' }}>
                    <thead>
                      <tr style={{ background: '#172554', color: '#dbeafe', fontSize: '0.78rem' }}>
                        <th style={{ padding: '6px 8px', textAlign: 'left', borderBottom: '1px solid #475569' }}>Supervisor</th>
                        <th style={{ padding: '6px 8px', textAlign: 'center', borderBottom: '1px solid #475569' }}>Lots</th>
                        <th style={{ padding: '6px 8px', textAlign: 'center', borderBottom: '1px solid #475569' }}>Qty</th>
                        <th style={{ padding: '6px 8px', textAlign: 'center', borderBottom: '1px solid #475569' }}>%</th>
                      </tr>
                    </thead>
                    <tbody>
                      {executiveSummaryData.supervisors.map((sup, idx) => {
                        const pct = executiveSummaryData.totalQty > 0 ? ((sup.qty / executiveSummaryData.totalQty) * 100).toFixed(1) : '0';
                        return (
                          <tr key={idx} style={{ background: idx % 2 === 1 ? 'rgba(255,255,255,0.03)' : 'transparent', borderBottom: '1px solid #334155' }}>
                            <td style={{ padding: '6px 8px', fontWeight: 600 }}>{sup.name}</td>
                            <td style={{ padding: '6px 8px', textAlign: 'center' }}>{sup.lots}</td>
                            <td style={{ padding: '6px 8px', textAlign: 'center', fontWeight: 700, color: '#ffffff' }}>{sup.qty.toLocaleString()}</td>
                            <td style={{ padding: '6px 8px', textAlign: 'center', color: '#94a3b8' }}>{pct}%</td>
                          </tr>
                        );
                      })}
                      <tr style={{ background: '#334155', fontWeight: 800, color: '#ffffff' }}>
                        <td style={{ padding: '6px 8px' }}>TOTAL</td>
                        <td style={{ padding: '6px 8px', textAlign: 'center' }}>{executiveSummaryData.totalLots}</td>
                        <td style={{ padding: '6px 8px', textAlign: 'center', color: '#93c5fd' }}>{executiveSummaryData.totalQty.toLocaleString()}</td>
                        <td style={{ padding: '6px 8px', textAlign: 'center' }}>100%</td>
                      </tr>
                    </tbody>
                  </table>
                </div>

                {/* 5. Aging Breakdown */}
                <div style={{ background: '#1e293b', borderRadius: '12px', overflow: 'hidden', border: '1px solid #475569' }}>
                  <div style={{ background: '#991b1b', padding: '8px 12px', color: '#ffffff', fontWeight: 800, fontSize: '0.85rem', textAlign: 'center', letterSpacing: '0.03em' }}>
                    ⏳ AGING RANGE
                  </div>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.82rem', color: '#f8fafc' }}>
                    <thead>
                      <tr style={{ background: '#7f1d1d', color: '#fee2e2', fontSize: '0.78rem' }}>
                        <th style={{ padding: '6px 8px', textAlign: 'left', borderBottom: '1px solid #475569' }}>Aging Range</th>
                        <th style={{ padding: '6px 8px', textAlign: 'center', borderBottom: '1px solid #475569' }}>Lots</th>
                        <th style={{ padding: '6px 8px', textAlign: 'center', borderBottom: '1px solid #475569' }}>Qty</th>
                        <th style={{ padding: '6px 8px', textAlign: 'center', borderBottom: '1px solid #475569' }}>%</th>
                      </tr>
                    </thead>
                    <tbody>
                      {executiveSummaryData.aging.map((ag, idx) => {
                        const pct = executiveSummaryData.totalQty > 0 ? ((ag.qty / executiveSummaryData.totalQty) * 100).toFixed(1) : '0';
                        return (
                          <tr key={idx} style={{ background: idx % 2 === 1 ? 'rgba(255,255,255,0.03)' : 'transparent', borderBottom: '1px solid #334155' }}>
                            <td style={{ padding: '6px 8px', fontWeight: 600 }}>{ag.name}</td>
                            <td style={{ padding: '6px 8px', textAlign: 'center' }}>{ag.lots}</td>
                            <td style={{ padding: '6px 8px', textAlign: 'center', fontWeight: 700, color: '#ffffff' }}>{ag.qty.toLocaleString()}</td>
                            <td style={{ padding: '6px 8px', textAlign: 'center', color: '#94a3b8' }}>{pct}%</td>
                          </tr>
                        );
                      })}
                      <tr style={{ background: '#334155', fontWeight: 800, color: '#ffffff' }}>
                        <td style={{ padding: '6px 8px' }}>TOTAL</td>
                        <td style={{ padding: '6px 8px', textAlign: 'center' }}>{executiveSummaryData.totalLots}</td>
                        <td style={{ padding: '6px 8px', textAlign: 'center', color: '#fca5a5' }}>{executiveSummaryData.totalQty.toLocaleString()}</td>
                        <td style={{ padding: '6px 8px', textAlign: 'center' }}>100%</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>
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
                  <span className="remarks-lot-field-label">Stitching Qty</span>
                  <span className="remarks-lot-field-value" style={{ color: '#ef4444', fontWeight: '800' }}>
                    {selectedRemarksLot.stitchingIssueQty || '0'} pcs
                  </span>
                </div>
                <div className="remarks-lot-field">
                  <span className="remarks-lot-field-label">Supervisor</span>
                  <span className="remarks-lot-field-value">{selectedRemarksLot.supervisor || '-'}</span>
                </div>
                <div className="remarks-lot-field">
                  <span className="remarks-lot-field-label">Pending Days</span>
                  <span className="remarks-lot-field-value" style={{ color: '#4338ca' }}>
                    {selectedRemarksLot.pendingDays || 0} days
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
                className="btn-cancel-remark"
                onClick={handleCloseRemarksModal}
                disabled={savingRemark}
              >
                Cancel
              </button>
              <button
                type="button"
                className="btn-save-remark"
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

export default PendingPackingtoIssue;