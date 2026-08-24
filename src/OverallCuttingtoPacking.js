import React, { useState, useEffect, useRef } from 'react';
import './OverallCuttingtoPacking.css';
import * as XLSX from 'xlsx'; // For Excel export
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { GOOGLE_API_KEY, SPREADSHEET_IDS, SHEET_NAMES, fetchSheetDataFromBackend } from './config';

const OverallCuttingToPacking = () => {
  // Credentials sourced from central .env configuration
  const SHEET_ID = SPREADSHEET_IDS.MAIN;
  const API_KEY = GOOGLE_API_KEY;
  const SHEET_IDD = SPREADSHEET_IDS.JOBORDER;
  const SHEET_IDDD = SPREADSHEET_IDS.ISSUES;
  
  // Sheet configurations
  const JOB_ORDER_SHEET = SHEET_NAMES.JOB_ORDER;
  const INDEX_SHEET = SHEET_NAMES.INDEX;
  const ISSUES_SHEET = SHEET_NAMES.ISSUES;
  const SHEET_ID_RAWPACK = SPREADSHEET_IDS.RAWPACK || '1xD8Uy1lUgvNTQ2RGRBI4ZjOrozbinUPRq2_UfIplP98';
  const RAWPACK_SHEET = SHEET_NAMES.RAWPACK || 'RAWPACK';

// Reusable Multi-Select Dropdown Filter Component
const MultiSelectFilter = ({ label, options = [], selected = [], onChange, placeholder = "Select..." }) => {
  const [isOpen, setIsOpen] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const dropdownRef = useRef(null);

  useEffect(() => {
    const handleClickOutside = (event) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const handleSelect = (item) => {
    if (selected.includes(item)) {
      onChange(selected.filter(i => i !== item));
    } else {
      onChange([...selected, item]);
    }
  };

  const handleSelectAll = () => {
    onChange([...options]);
  };

  const handleClearAll = () => {
    onChange([]);
    setSearchTerm('');
  };

  const filteredOptions = options.filter(opt => 
    String(opt).toLowerCase().includes(searchTerm.toLowerCase())
  );

  return (
    <div className="filter-group fabric-filter-group" ref={dropdownRef}>
      <label>
        {label} {selected.length > 0 && <span className="filter-count">({selected.length})</span>}
      </label>
      <div 
        className={`fabric-dropdown ${isOpen ? 'active' : ''}`}
        onClick={() => setIsOpen(!isOpen)}
      >
        <div className="fabric-dropdown-header">
          {selected.length === 0 ? (
            <span className="placeholder">{placeholder}</span>
          ) : (
            <span className="selected-count">{selected.length} selected</span>
          )}
          <span className="dropdown-arrow">▼</span>
        </div>
        
        {isOpen && (
          <div className="fabric-dropdown-content">
            <div className="fabric-search">
              <input
                type="text"
                placeholder={`Search ${label.toLowerCase()}...`}
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                onClick={(e) => e.stopPropagation()}
                className="fabric-search-input"
              />
            </div>
            
            <div className="fabric-actions">
              <button 
                type="button"
                onClick={(e) => { e.stopPropagation(); handleSelectAll(); }}
                className="fabric-action-btn"
              >
                Select All
              </button>
              <button 
                type="button"
                onClick={(e) => { e.stopPropagation(); handleClearAll(); }}
                className="fabric-action-btn"
              >
                Clear All
              </button>
            </div>
            
            <div className="fabric-options">
              {filteredOptions.map(opt => (
                <label key={opt} className="fabric-option" title={String(opt)}>
                  <input
                    type="checkbox"
                    checked={selected.includes(opt)}
                    onChange={() => handleSelect(opt)}
                    onClick={(e) => e.stopPropagation()}
                  />
                  <span className="fabric-name">{String(opt)}</span>
                </label>
              ))}
              
              {filteredOptions.length === 0 && (
                <div className="no-fabrics-found">No options found</div>
              )}
            </div>
          </div>
        )}
      </div>

      {selected.length > 0 && (
        <div className="selected-fabrics">
          {selected.map(item => (
            <span key={item} className="selected-fabric-tag">
              {item}
              <button 
                type="button"
                onClick={() => handleSelect(item)}
                className="remove-fabric"
              >
                ×
              </button>
            </span>
          ))}
        </div>
      )}
    </div>
  );
};
  
  // Headers from JobOrder sheet - Updated to include all required fields
// Headers from JobOrder sheet - Updated to include all required fields
// Headers from JobOrder sheet - Updated to include all required fields
// Headers from JobOrder sheet - Updated to include Fabric
const JOB_ORDER_HEADERS = [
  'Lot Number',
  'Garment Type',
  'Style',
  'Brand',
  'Party Name',
  'SEASON',
  'SECTION',
  'DIRECT STITCHING',
  'Date',
  'Status',
  'Emb',
  'Printing',
  'Fabric',
  'Size'
];
  // Headers from Index sheet (additional details)
  const INDEX_HEADERS = [
    'Lot Number',
    'Saved At',
    'Date of Issue',
    'Supervisor',
    'PARTY NAME',
    'BRAND',
    'SEASON',
    'DIRECT STITCHING',
    'CHALLAN HISTORY',
    'WIP Status',
    'Completed Status',
    'Cutting Qty',
    'Image URL',
    'Image'
  ];
  
  // Headers from Issues sheet (packing details)
// Headers from Issues sheet (packing details) - UPDATED
const ISSUES_HEADERS = [
  'Timestamp',
  'Lot Number',
  'Garment Type',
  'Fabric',        // Added this if it exists in your sheet
  'Style',
  'Packing Supervisor',
  'Packing Date',
  'Total Pcs',     // This might be the same as 'Total Pcs' in your sheet
  'WIP Packing',
  'Packing Complete'
];
  // Combined headers for display with short forms - INCLUDING DAYS COLUMNS
  // Combined headers for display with short forms - INCLUDING DAYS COLUMNS
const DISPLAY_HEADERS = [
  'Sr.',            // Serial number
  'Lot No',
  'Image',          // Image column from Index sheet
  'Fabric',        // MOVED HERE - after Lot No
  'Size',          // Size from JobOrder sheet
  'Item',
  'Brand',
  'Party',
  'Season',
  'Section',
  'Design Work',
  'Job Date',       // From JobOrder
  'PCS',           // Cutting Qty
  'Cut Date',      // Saved At
  'Emb/Print Issue', // From Challan History
  'Emb/print Comp', // From Challan History
  'Stit Date',      // Issue Date
  'Stit Sup',       // Supervisor
  'WIP Stit',       // WIP Status
  'Comp Stit',      // Comp. Status
  'Pkg Sup',        // From Issues
  'Pkg Date',       // From Issues
  'WIP Pkg',        // From Issues
  'Pkg Comp',       // From Issues
  // DAYS COLUMNS - ADDED AT THE END
  'Cut Days',
  'Emb/Print Days',
  'Stit Days',
  'Pkg Days',
  'Cut To Emb/Print',
  'Emb/Print To Stit',
  'Stit To Check Pack'
];
  const [data, setData] = useState([]);
  const [jobOrderData, setJobOrderData] = useState([]);
  const [indexData, setIndexData] = useState([]);
  const [issuesData, setIssuesData] = useState([]);
  const [rawpackData, setRawpackData] = useState([]);
  const [headers, setHeaders] = useState(DISPLAY_HEADERS);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [brandFilter, setBrandFilter] = useState([]);
  const [garmentTypeFilter, setGarmentTypeFilter] = useState([]);
  const [styleFilter, setStyleFilter] = useState([]);
  const [partyNameFilter, setPartyNameFilter] = useState([]);
  const [seasonFilter, setSeasonFilter] = useState([]);
  const [sectionFilter, setSectionFilter] = useState([]);
  const [directStitchingFilter, setDirectStitchingFilter] = useState([]);
  const [selectedChallan, setSelectedChallan] = useState(null);
  const [selectedWIP, setSelectedWIP] = useState(null);
  const [selectedCompleted, setSelectedCompleted] = useState(null);
  const [viewImageSrc, setViewImageSrc] = useState(null);
  const [exportLoading, setExportLoading] = useState(false);
  const tableRef = useRef(null);

  // Helper function to extract direct Google Drive image URL
  const getDirectImageUrl = (url) => {
    if (!url) return '';
    const trimmed = url.toString().trim();
    if (!trimmed) return '';
    
    if (!trimmed.includes('drive.google.com') && !trimmed.includes('docs.google.com')) {
      return trimmed;
    }
    
    let fileId = '';
    const dMatch = trimmed.match(/[?&]id=([^&]+)/);
    if (dMatch && dMatch[1]) {
      fileId = dMatch[1];
    } else {
      const fileMatch = trimmed.match(/\/file\/d\/([^/]+)/) || trimmed.match(/\/d\/([^/]+)/);
      if (fileMatch && fileMatch[1]) {
        fileId = fileMatch[1];
      }
    }
    
    if (fileId) {
      return `https://drive.google.com/thumbnail?id=${fileId}&sz=w800`;
    }
    
    return trimmed;
  };

  // Convert image URL to Base64 data URL for jsPDF embedding
  const getBase64ImageFromUrl = async (imageUrl) => {
    if (!imageUrl) return null;
    const trimmed = imageUrl.toString().trim();
    if (!trimmed) return null;

    // Helper: try fetch to base64
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

    // Helper: try Image element to base64 via Canvas
    const loadViaCanvas = (targetUrl) => {
      return new Promise((resolve) => {
        const img = new Image();
        img.crossOrigin = 'anonymous';
        img.onload = () => {
          try {
            const canvas = document.createElement('canvas');
            canvas.width = img.naturalWidth || 100;
            canvas.height = img.naturalHeight || 100;
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

    // 1. Try direct fetch
    let result = await fetchAsBase64(trimmed);
    if (result) return result;

    // 2. Try weserv.nl image proxy (bypasses Google Drive CORS & referrer restrictions)
    const cleanUrl = trimmed.replace(/^https?:\/\//, '');
    const proxiedUrl = `https://images.weserv.nl/?url=${encodeURIComponent(cleanUrl)}`;
    
    result = await fetchAsBase64(proxiedUrl);
    if (result) return result;

    // 3. Fallback to Canvas rendering with proxied URL
    result = await loadViaCanvas(proxiedUrl);
    if (result) return result;

    // 4. Final fallback to direct canvas load
    return await loadViaCanvas(trimmed);
  };
  const [showPartyInitials, setShowPartyInitials] = useState(true);
  const [supervisorFilter, setSupervisorFilter] = useState([]);
  const [embPrintCompFilter, setEmbPrintCompFilter] = useState([]);
  const [compStatusFilter, setCompStatusFilter] = useState([]);
  const [pkgSupervisorFilter, setPkgSupervisorFilter] = useState([]);
  const [pkgCompFilter, setPkgCompFilter] = useState([]);
  const [holdLotsFilter, setHoldLotsFilter] = useState(false);
  const [fabricFilter, setFabricFilter] = useState([]); // Array for multi-select
  const [fabricSearchTerm, setFabricSearchTerm] = useState('');
  const [showFabricDropdown, setShowFabricDropdown] = useState(false);
  const [dateFilterField, setDateFilterField] = useState('Job Date'); // 'Job Date' | 'Cut Date' | 'Both'
  const [startDateFilter, setStartDateFilter] = useState('');
  const [endDateFilter, setEndDateFilter] = useState('');
  const fabricDropdownRef = useRef(null);

  // Helper function to parse date values for date range filtering
  const parseDateForFilter = (dateValue) => {
    if (!dateValue || dateValue === '-' || dateValue === 'invalid date') return null;
    if (dateValue instanceof Date) return isNaN(dateValue.getTime()) ? null : dateValue;
    
    const str = String(dateValue).trim();
    if (!str) return null;

    if (str.includes('/')) {
      const parts = str.split(' ')[0].split('/');
      if (parts.length === 3) {
        let day = parseInt(parts[0], 10);
        let month = parseInt(parts[1], 10) - 1;
        let year = parseInt(parts[2], 10);
        if (year < 100) year = year <= 50 ? 2000 + year : 1900 + year;
        const d = new Date(year, month, day);
        if (!isNaN(d.getTime())) return d;
      }
    }

    if (str.includes('-')) {
      const parts = str.split(' ')[0].split('T')[0].split('-');
      if (parts.length === 3 && parts[0].length === 4) {
        let year = parseInt(parts[0], 10);
        let month = parseInt(parts[1], 10) - 1;
        let day = parseInt(parts[2], 10);
        const d = new Date(year, month, day);
        if (!isNaN(d.getTime())) return d;
      }
    }

    const d = new Date(str);
    return isNaN(d.getTime()) ? null : d;
  };

  // Function to format date from various formats to "DD/MM/YY"
// Function to format date from various formats to "DD/MM/YY"
const formatDate = (dateString) => {
  if (!dateString) {
    return '';
  }
  
  if (dateString instanceof Date) {
    if (isNaN(dateString.getTime())) {
      return '';
    }
    
    const day = dateString.getDate().toString().padStart(2, '0');
    const month = (dateString.getMonth() + 1).toString().padStart(2, '0');
    const year = dateString.getFullYear().toString().slice(-2);
    
    return `${day}/${month}/${year}`;
  }
  
  if (typeof dateString !== 'string') {
    try {
      dateString = String(dateString);
    } catch (e) {
      console.warn('Cannot convert date to string:', dateString);
      return '';
    }
  }
  
  if (dateString.trim() === '') {
    return '';
  }
  
  try {
    // NEW: Handle YYYY-MM-DD format (from Issues sheet)
    if (dateString.match(/^\d{4}-\d{1,2}-\d{1,2}$/)) {
      const parts = dateString.split('-');
      if (parts.length === 3) {
        const year = parseInt(parts[0]);
        const month = parseInt(parts[1]) - 1; // JS months are 0-indexed
        const day = parseInt(parts[2]);
        const date = new Date(year, month, day);
        
        if (!isNaN(date.getTime())) {
          const formattedDay = date.getDate().toString().padStart(2, '0');
          const formattedMonth = (date.getMonth() + 1).toString().padStart(2, '0');
          const formattedYear = date.getFullYear().toString().slice(-2);
          return `${formattedDay}/${formattedMonth}/${formattedYear}`;
        }
      }
    }
    
    let date = new Date(dateString);
    
    if (isNaN(date.getTime())) {
      // Try other formats
      if (dateString.match(/\d{1,2}\s+\w+\s+\d{4}/)) {
        const parts = dateString.split(' ');
        if (parts.length === 3) {
          const day = parseInt(parts[0]);
          const month = parts[1];
          const year = parseInt(parts[2]);
          
          const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
          const monthIndex = monthNames.findIndex(m => 
            m.toLowerCase() === month.slice(0, 3).toLowerCase()
          );
          
          if (monthIndex !== -1) {
            date = new Date(year, monthIndex, day);
          }
        }
      }
      
      if (isNaN(date.getTime())) {
        // Try DD/MM/YYYY or DD/MM/YY format
        if (dateString.includes('/')) {
          const parts = dateString.split('/');
          if (parts.length === 3) {
            const day = parseInt(parts[0], 10);
            const month = parseInt(parts[1], 10) - 1;
            const year = parseInt(parts[2], 10);
            const fullYear = year < 100 ? (year <= 50 ? 2000 + year : 1900 + year) : year;
            date = new Date(fullYear, month, day);
          }
        }
      }
      
      if (isNaN(date.getTime())) {
        // Return the original string if we can't parse it
        return dateString;
      }
    }
    
    const day = date.getDate().toString().padStart(2, '0');
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const year = date.getFullYear().toString().slice(-2);
    
    return `${day}/${month}/${year}`;
  } catch (error) {
    console.warn('Error formatting date:', dateString, error);
    return dateString;
  }
};

  // Function to calculate days difference between two dates
  const calculateDaysDifference = (date1Str, date2Str) => {
    if (!date1Str || !date2Str || date1Str === '' || date2Str === '' || 
        date1Str === '-' || date2Str === '-') {
      return '-';
    }
    
    const parseDate = (dateString) => {
      if (!dateString) return null;
      
      if (dateString instanceof Date) {
        return isNaN(dateString.getTime()) ? null : dateString;
      }
      
      if (typeof dateString !== 'string') {
        dateString = String(dateString);
      }
      
      if (dateString.trim() === '') {
        return null;
      }
      
      try {
        // Handle "DD/MM/YY" or "DD/MM/YYYY" format
        if (dateString.includes('/')) {
          const parts = dateString.split('/');
          if (parts.length === 3) {
            const day = parseInt(parts[0], 10);
            const month = parseInt(parts[1], 10) - 1;
            const year = parseInt(parts[2], 10);
            const fullYear = year < 100 ? (year <= 50 ? 2000 + year : 1900 + year) : year;
            
            const date = new Date(fullYear, month, day);
            if (date.getDate() === day && date.getMonth() === month && 
                date.getFullYear() === fullYear) {
              return date;
            }
          }
        }
        
        // Handle "DD MMM YYYY" format
        if (dateString.match(/\d{1,2}\s+[a-zA-Z]{3,}\s+\d{4}/)) {
          const monthMap = {
            'jan': 0, 'january': 0, 'feb': 1, 'february': 1, 'mar': 2, 'march': 2,
            'apr': 3, 'april': 3, 'may': 4, 'jun': 5, 'june': 5, 'jul': 6, 'july': 6,
            'aug': 7, 'august': 7, 'sep': 8, 'september': 8, 'oct': 9, 'october': 9,
            'nov': 10, 'november': 10, 'dec': 11, 'december': 11
          };
          
          const parts = dateString.split(' ');
          const day = parseInt(parts[0], 10);
          const monthName = parts[1].toLowerCase();
          const year = parseInt(parts[2], 10);
          
          if (monthMap[monthName] !== undefined) {
            const date = new Date(year, monthMap[monthName], day);
            if (date.getDate() === day && date.getMonth() === monthMap[monthName] && 
                date.getFullYear() === year) {
              return date;
            }
          }
        }
        
        // Try default Date constructor
        const date = new Date(dateString);
        if (!isNaN(date.getTime())) {
          return date;
        }
        
        return null;
      } catch (error) {
        console.warn('Error parsing date:', dateString, error);
        return null;
      }
    };
    
    const date1 = parseDate(date1Str);
    const date2 = parseDate(date2Str);
    
    if (!date1 || !date2) return '-';
    if (!(date1 instanceof Date) || isNaN(date1.getTime()) ||
        !(date2 instanceof Date) || isNaN(date2.getTime())) {
      return '-';
    }
    
    // Calculate difference in days (date2 - date1) to allow negative values
    const diffTime = date2.getTime() - date1.getTime();
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    
    // Return with minus sign for negative values
    if (diffDays < 0) {
      return diffDays.toString();
    } else if (diffDays === 0) {
      return '0';
    } else {
      return `+${diffDays}`;
    }
  };

  // Function to calculate all days for a row
  const calculateAllDays = (item) => {
    const days = {};
    
    // Cut Days: Cut Date - Job Date
    const cutDate = item['Cut Date'] || item['Saved At'] || '';
    const jobDate = item['Job Date'] || item['Date'] || '';
    days.cutDays = calculateDaysDifference(jobDate, cutDate);
    
    // Emb/Print Days: Emb/Print Comp - Emb/Print Issue
    const embPrintComp = item['Emb/print Comp'] || item['Emb/Print Comp.'] || '';
    const embPrintIssue = item['Emb/Print Issue'] || '';
    days.embPrintDays = calculateDaysDifference(embPrintIssue, embPrintComp);
    
    // Stit Days: Comp Stit - Stit Date
    const compStit = item['Comp Stit'] || item['Comp. Status'] || '';
    const stitDate = item['Stit Date'] || item['Issue Date'] || '';
    days.stitDays = calculateDaysDifference(stitDate, compStit);
    
    // Pkg Days: Pkg Comp - Pkg Date
    const pkgComp = item['Pkg Comp'] || item['Pkg. Comp.'] || '';
    const pkgDate = item['Pkg Date'] || item['Pkg. Date'] || '';
    days.pkgDays = calculateDaysDifference(pkgDate, pkgComp);
    
    // Cut To Emb/Print: Emb/Print Issue - Cut Date
    days.cutToEmbPrint = calculateDaysDifference(cutDate, embPrintIssue);
    
    // Emb/Print To Stit: Stit Date - Emb/Print Comp
    days.embPrintToStit = calculateDaysDifference(embPrintComp, stitDate);
    
    // Stit To Check Pack: Pkg Date - Comp Stit
    days.stitToCheckPack = calculateDaysDifference(compStit, pkgDate);
    
    return days;
  };

  // Function to convert party name to initials
  const getPartyInitials = (partyName) => {
    if (!partyName || typeof partyName !== 'string') {
      return '';
    }
    
    const trimmedParty = partyName.trim();
    
    if (trimmedParty.toLowerCase().includes('mohit hosiery')) {
      return 'MH';
    } else if (trimmedParty.toLowerCase().includes('imran')) {
      return 'Imr';
    }
    
    const words = trimmedParty.split(' ');
    if (words.length === 1) {
      return trimmedParty.substring(0, 3);
    } else {
      return words.map(word => word.charAt(0)).join('').substring(0, 3);
    }
  };

  // Function to parse date from various formats
  const parseDate = (dateString) => {
    if (!dateString) {
      return null;
    }
    
    if (dateString instanceof Date) {
      return isNaN(dateString.getTime()) ? null : dateString;
    }
    
    if (typeof dateString !== 'string') {
      dateString = String(dateString);
    }
    
    if (dateString.trim() === '') {
      return null;
    }
    
    try {
      if (dateString.includes('/')) {
        const parts = dateString.split(' ')[0].split('/');
        if (parts.length === 3) {
          const month = parseInt(parts[0]) - 1;
          const day = parseInt(parts[1]);
          const year = parseInt(parts[2]);
          return new Date(year, month, day);
        }
      }
      
      if (dateString.includes('-')) {
        const parts = dateString.split(' ')[0].split('-');
        if (parts.length === 3) {
          const year = parseInt(parts[0]);
          const month = parseInt(parts[1]) - 1;
          const day = parseInt(parts[2]);
          return new Date(year, month, day);
        }
      }
      
      const date = new Date(dateString);
      if (!isNaN(date.getTime())) {
        return date;
      }
      
      return null;
    } catch (error) {
      console.warn('Error parsing date:', dateString, error);
      return null;
    }
  };
  // Get unique fabric values for filter
// Get unique fabric values for filter
const getUniqueFabrics = () => {
  const fabrics = data
    .map(item => {
      // Get fabric from JobOrder sheet data
      let fabric = item['Fabric'] || item['Fabric_From_JobOrder'] || item._raw?.['Fabric'] || '';
      
      return fabric ? String(fabric).trim() : '';
    })
    .filter(fabric => fabric && fabric !== '')
    .filter((fabric, index, self) => self.indexOf(fabric) === index)
    .sort();
  
  return fabrics;
};
// Handle fabric selection/deselection
const handleFabricSelect = (fabric) => {
  setFabricFilter(prev => {
    if (prev.includes(fabric)) {
      return prev.filter(f => f !== fabric);
    } else {
      return [...prev, fabric];
    }
  });
};

// Select all fabrics
const selectAllFabrics = () => {
  const allFabrics = getUniqueFabrics();
  setFabricFilter(allFabrics);
};

// Clear all fabric selections
const clearFabricFilters = () => {
  setFabricFilter([]);
  setFabricSearchTerm('');
};

  // Helper function to extract date from object
  const getDateFromObject = (obj, dateFields) => {
    if (!obj) return null;
    
    for (const field of dateFields) {
      if (obj[field]) {
        try {
          if (field === 'date' && typeof obj[field] === 'string' && obj[field].match(/\d{1,2}\s+\w+\s+\d{4}/)) {
            const dateStr = obj[field];
            const parts = dateStr.split(' ');
            if (parts.length === 3) {
              const day = parseInt(parts[0]);
              const month = parts[1];
              const year = parseInt(parts[2]);
              
              const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
              const monthIndex = monthNames.findIndex(m => m.toLowerCase() === month.toLowerCase());
              
              if (monthIndex !== -1) {
                const date = new Date(year, monthIndex, day);
                if (!isNaN(date.getTime())) {
                  return date;
                }
              }
            }
          }
          
          const date = new Date(obj[field]);
          if (!isNaN(date.getTime())) {
            return date;
          }
        } catch (e) {
          // Try next field
        }
      }
    }
    return null;
  };

  const parseWIPPackingData = (content) => {
    if (!content || content.trim() === '') {
      return {
        fullData: [],
        displayValue: '',
        latestDate: '',
        latestRemarks: '',
        raw: content,
        isHold: false
      };
    }
    
    try {
      let statusData = [];
      
      if (content.trim().startsWith('[')) {
        statusData = JSON.parse(content);
      } else {
        const jsonMatch = content.match(/\[.*\]/s);
        if (jsonMatch) {
          statusData = JSON.parse(jsonMatch[0]);
        } else {
          return {
            fullData: [],
            displayValue: content,
            latestDate: '',
            latestRemarks: content,
            raw: content,
            isHold: content.toLowerCase().includes('hold')
          };
        }
      }
      
      if (!Array.isArray(statusData) || statusData.length === 0) {
        return {
          fullData: [],
          displayValue: content.length > 50 ? content.substring(0, 50) + '...' : content,
          latestDate: '',
          latestRemarks: content,
          raw: content,
          isHold: content.toLowerCase().includes('hold')
        };
      }
      
      const sortedData = [...statusData].sort((a, b) => {
        try {
          return new Date(b.timestamp) - new Date(a.timestamp);
        } catch (e) {
          return 0;
        }
      });
      
      const latest = sortedData[0] || {};
      
      let latestDate = '';
      if (latest.timestamp) {
        latestDate = formatDate(latest.timestamp);
      }
      
      const latestRemarks = latest.remarks || latest.status || '';
      const isHold = latestRemarks.toLowerCase().includes('hold') || 
                     (latest.status && latest.status.toLowerCase().includes('hold'));
      
      let displayValue = latestDate;
      if (latestRemarks && latestRemarks.trim() !== '') {
        displayValue = latestDate ? `${latestDate}\n${latestRemarks}` : latestRemarks;
      }
      
      return {
        fullData: sortedData,
        displayValue: displayValue || 'No Data',
        latestDate: latestDate,
        latestRemarks: latestRemarks,
        raw: content,
        isHold: isHold
      };
      
    } catch (error) {
      console.warn('Error parsing WIP Packing data:', error, content);
      return {
        fullData: [],
        displayValue: content.length > 50 ? content.substring(0, 50) + '...' : content,
        latestDate: '',
        latestRemarks: content,
        raw: content,
        isHold: content.toLowerCase().includes('hold')
      };
    }
  };

  // Function to convert season to single letter code
  const getSeasonCode = (season) => {
    if (!season || typeof season !== 'string') {
      return '';
    }
    
    const trimmedSeason = season.trim().toLowerCase();
    
    if (trimmedSeason.includes('summer')) {
      return 'S';
    } else if (trimmedSeason.includes('winter')) {
      return 'W';
    }
    
    return trimmedSeason.charAt(0).toUpperCase();
  };

  // Function to convert section to single letter code
  const getSectionCode = (section) => {
    if (!section || typeof section !== 'string') {
      return '';
    }
    
    const trimmedSection = section.trim().toLowerCase();
    
    if (trimmedSection.includes('gent') || trimmedSection.includes('men') || trimmedSection === 'm') {
      return 'M';
    } else if (trimmedSection.includes('girl') || trimmedSection.includes('women') || trimmedSection === 'g') {
      return 'G';
    } else if (trimmedSection.includes('kid') || trimmedSection.includes('child') || trimmedSection === 'k') {
      return 'K';
    }
    
    return trimmedSection.charAt(0).toUpperCase();
  };

  // Function to extract challan data from mixed content
  const extractChallanData = (content) => {
    if (!content || content.trim() === '') {
      return { challanData: null, wipData: null, raw: content };
    }
    
    const trimmed = content.trim();
    const jsonMatches = [];
    let start = 0;
    
    while (start < trimmed.length) {
      const arrayStart = trimmed.indexOf('[', start);
      const objectStart = trimmed.indexOf('{', start);
      
      let nextStart = -1;
      let isArray = true;
      
      if (arrayStart !== -1 && objectStart !== -1) {
        if (arrayStart < objectStart) {
          nextStart = arrayStart;
          isArray = true;
        } else {
          nextStart = objectStart;
          isArray = false;
        }
      } else if (arrayStart !== -1) {
        nextStart = arrayStart;
        isArray = true;
      } else if (objectStart !== -1) {
        nextStart = objectStart;
        isArray = false;
      }
      
      if (nextStart === -1) break;
      
      let jsonStr = '';
      let end = nextStart;
      let depth = 0;
      let inString = false;
      let escapeNext = false;
      
      for (let i = nextStart; i < trimmed.length; i++) {
        const char = trimmed[i];
        
        if (escapeNext) {
          escapeNext = false;
          continue;
        }
        
        if (char === '\\') {
          escapeNext = true;
          continue;
        }
        
        if (char === '"' && !inString) {
          inString = true;
        } else if (char === '"' && inString) {
          inString = false;
        }
        
        if (!inString) {
          if (char === '[' || char === '{') depth++;
          if (char === ']' || char === '}') {
            depth--;
            if (depth === 0) {
              end = i;
              break;
            }
          }
        }
      }
      
      if (depth === 0) {
        jsonStr = trimmed.substring(nextStart, end + 1);
        try {
          const parsed = JSON.parse(jsonStr);
          jsonMatches.push({
            data: parsed,
            start: nextStart,
            end: end + 1,
            isArray: isArray
          });
        } catch (e) {
          // Not valid JSON, continue
        }
      }
      
      start = end + 1;
    }
    
    let challanData = null;
    let wipData = null;
    
    jsonMatches.forEach(match => {
      if (Array.isArray(match.data)) {
        if (match.data.length > 0) {
          const firstItem = match.data[0];
          
          if (firstItem.number && (firstItem.number.toUpperCase().includes('CH-EMB-') || firstItem.number.toUpperCase().includes('CH-PRINT-'))) {
            challanData = match.data;
          } 
          else if (firstItem.status || firstItem.timestamp) {
            wipData = match.data;
          }
        }
      } else if (match.data.number && (match.data.number.toUpperCase().includes('CH-EMB-') || match.data.number.toUpperCase().includes('CH-PRINT-'))) {
        challanData = [match.data];
      }
    });
    
    if (!challanData || challanData.length === 0) {
      const plainChallanRegex = /CH-(EMB|PRINT)-\d+/gi;
      const plainMatches = trimmed.match(plainChallanRegex);
      
      if (plainMatches && plainMatches.length > 0) {
        challanData = plainMatches.map(number => ({
          number: number.toUpperCase(),
          totalQty: 0
        }));
      }
    }
    
    return { challanData, wipData, raw: content };
  };

  // Function to analyze challan data and get status with date
  const analyzeChallanData = (challanData) => {
    if (!challanData || !Array.isArray(challanData) || challanData.length === 0) {
      return { 
        status: 'No Challan', 
        type: 'none',
        date: '',
        display: 'No Challan',
        dateOfIssue: '',
        embUpdatedAt: '',
        hasMultiple: false,
        count: 0,
        totalQty: 0,
        sortedChallans: [],
        allCompleted: false,
        hasPending: false
      };
    }
    
    try {
      const sortedChallans = [...challanData].sort((a, b) => {
        const dateA = getDateFromObject(a, ['date']);
        const dateB = getDateFromObject(b, ['date']);
        return (dateA || new Date(0)) - (dateB || new Date(0));
      });
      
      const firstChallan = sortedChallans[0];
      let dateOfIssue = '';
      
      if (firstChallan?.date) {
        if (firstChallan.date.match(/\d{1,2}\s+\w+\s+\d{4}/)) {
          dateOfIssue = firstChallan.date;
        } else {
          const parsedDate = getDateFromObject(firstChallan, ['date']);
          if (parsedDate) {
            dateOfIssue = formatDate(parsedDate);
          }
        }
      }
      
      const lastChallan = sortedChallans[sortedChallans.length - 1];
      const lastChallanHasUpdate = lastChallan?.embUpdatedAt ? true : false;
      
      let embUpdatedAt = '';
      
      if (lastChallanHasUpdate) {
        const completedChallans = sortedChallans.filter(challan => 
          challan.embCompleted === true && challan.embUpdatedAt
        );
        
        if (completedChallans.length > 0) {
          completedChallans.sort((a, b) => {
            const dateA = getDateFromObject(a, ['embUpdatedAt']);
            const dateB = getDateFromObject(b, ['embUpdatedAt']);
            return (dateB || new Date(0)) - (dateA || new Date(0));
          });
          
          const lastCompletedChallan = completedChallans[0];
          if (lastCompletedChallan?.embUpdatedAt) {
            const parsedDate = getDateFromObject(lastCompletedChallan, ['embUpdatedAt']);
            if (parsedDate) {
              embUpdatedAt = formatDate(parsedDate);
            }
          }
        }
      }
      
      const allChallansCompleted = sortedChallans.every(challan => 
        challan.embCompleted === true
      );
      
      const hasPendingChallans = sortedChallans.some(challan => 
        !challan.embUpdatedAt || challan.embCompleted === false
      );
      
      const lastChallanPending = !lastChallan?.embUpdatedAt;
      
      const firstNumber = firstChallan?.number || '';
      const isPrinting = firstNumber.toUpperCase().includes('PRINT');
      
      const totalQty = sortedChallans.reduce((sum, challan) => {
        const qty = parseInt(challan.totalQty) || 0;
        return sum + qty;
      }, 0);
      
      let status = '';
      let type = '';
      let display = '';
      
      if (lastChallanHasUpdate && allChallansCompleted && embUpdatedAt) {
        status = isPrinting ? 'Printing Done' : 'Emb Done';
        type = isPrinting ? 'printing' : 'embroidery';
        display = `${firstNumber}\n${embUpdatedAt}`;
      } 
      else if (lastChallanPending || hasPendingChallans) {
        status = isPrinting ? 'Printing Pending' : 'Emb Pending';
        type = isPrinting ? 'printing' : 'embroidery';
        display = dateOfIssue ? `${firstNumber}\n${dateOfIssue}` : `${firstNumber}\nPending`;
      } 
      else if (dateOfIssue) {
        status = isPrinting ? 'Printing Pending' : 'Emb Pending';
        type = isPrinting ? 'printing' : 'embroidery';
        display = `${firstNumber}\n${dateOfIssue}`;
      } 
      else {
        status = isPrinting ? 'Printing Pending' : 'Emb Pending';
        type = isPrinting ? 'printing' : 'embroidery';
        display = `${firstNumber}\nPending`;
      }
      
      return {
        status,
        type,
        date: dateOfIssue || embUpdatedAt || '',
        dateOfIssue: dateOfIssue || '',
        embUpdatedAt: embUpdatedAt || '',
        display,
        hasMultiple: sortedChallans.length > 1,
        count: sortedChallans.length,
        totalQty,
        sortedChallans,
        allCompleted: allChallansCompleted && lastChallanHasUpdate,
        hasPending: lastChallanPending || hasPendingChallans,
        lastChallanHasUpdate: lastChallanHasUpdate
      };
      
    } catch (error) {
      console.warn('Error analyzing challan data:', error);
      
      const firstChallan = challanData[0] || {};
      const firstNumber = firstChallan.number || 'Challan';
      const isPrinting = firstNumber.toUpperCase().includes('PRINT');
      
      let dateOfIssue = '';
      
      if (firstChallan.date) {
        dateOfIssue = firstChallan.date.match(/\d{1,2}\s+\w+\s+\d{4}/) 
          ? firstChallan.date 
          : formatDate(firstChallan.date);
      }
      
      const lastChallan = challanData[challanData.length - 1];
      const lastChallanHasUpdate = lastChallan?.embUpdatedAt ? true : false;
      
      let embUpdatedAt = '';
      let status = isPrinting ? 'Printing Pending' : 'Emb Pending';
      let allCompleted = false;
      
      if (lastChallanHasUpdate) {
        allCompleted = challanData.every(c => c.embCompleted === true);
        
        if (allCompleted) {
          const completedChallans = challanData.filter(c => c.embUpdatedAt);
          if (completedChallans.length > 0) {
            completedChallans.sort((a, b) => {
              try {
                return new Date(b.embUpdatedAt) - new Date(a.embUpdatedAt);
              } catch (e) {
                return 0;
              }
            });
            embUpdatedAt = formatDate(completedChallans[0].embUpdatedAt);
            status = isPrinting ? 'Printing Done' : 'Emb Done';
          }
        }
      }
      
      return {
        status,
        type: isPrinting ? 'printing' : 'embroidery',
        date: dateOfIssue || embUpdatedAt || '',
        dateOfIssue: dateOfIssue || '',
        embUpdatedAt: embUpdatedAt || '',
        display: `${firstNumber}\n${dateOfIssue || 'Pending'}`,
        hasMultiple: challanData.length > 1,
        count: challanData.length,
        totalQty: 0,
        sortedChallans: challanData,
        allCompleted,
        hasPending: !lastChallanHasUpdate,
        lastChallanHasUpdate
      };
    }
  };

  // Function to parse and format Challan History for display
  const parseChallanHistory = (content) => {
    if (!content || content.trim() === '') {
      return {
        display: 'No Challan',
        status: 'No Challan',
        type: 'none',
        date: '',
        dateOfIssue: '',
        embUpdatedAt: '',
        embStatus: 'No Challan',
        raw: content,
        challanData: null,
        wipData: null,
        hasMultiple: false,
        count: 0,
        totalQty: 0,
        allCompleted: false,
        hasPending: false,
        lastChallanHasUpdate: false
      };
    }
    
    const { challanData, wipData, raw } = extractChallanData(content);
    
    if (!challanData || challanData.length === 0) {
      const challanRegex = /CH-(EMB|PRINT)-\d+/gi;
      const challanMatches = content.match(challanRegex);
      
      if (challanMatches && challanMatches.length > 0) {
        const plainTextChallans = challanMatches.map(number => ({
          number: number.toUpperCase(),
          totalQty: 0,
          embCompleted: false,
          embUpdatedAt: null
        }));
        
        const dateRegex = /\d{1,2}\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4}/gi;
        const dateMatches = content.match(dateRegex);
        let latestDate = '';
        
        if (dateMatches && dateMatches.length > 0) {
          try {
            const dates = dateMatches.map(date => new Date(date))
              .filter(date => !isNaN(date.getTime()));
            
            if (dates.length > 0) {
              const latest = new Date(Math.max(...dates.map(d => d.getTime())));
              latestDate = formatDate(latest);
            }
          } catch (e) {
            latestDate = dateMatches[0];
          }
        }
        
        const isPrinting = plainTextChallans[0]?.number.includes('PRINT');
        const status = isPrinting ? 'Printing Pending' : 'Emb Pending';
        const hasMultiple = plainTextChallans.length > 1;
        const lastChallanHasUpdate = false;
        
        return {
          display: `${plainTextChallans[0].number}\n${latestDate || 'Pending'}`,
          status: status,
          type: isPrinting ? 'printing' : 'embroidery',
          date: latestDate,
          dateOfIssue: latestDate || '',
          embUpdatedAt: '',
          embStatus: status,
          hasMultiple: hasMultiple,
          count: plainTextChallans.length,
          totalQty: 0,
          allCompleted: false,
          hasPending: true,
          lastChallanHasUpdate: lastChallanHasUpdate,
          challans: plainTextChallans,
          wipData: wipData,
          raw: content
        };
      }
      
      return {
        display: content.length > 30 ? content.substring(0, 30) + '...' : content,
        status: 'Text Data',
        type: 'text',
        date: '',
        dateOfIssue: '',
        embUpdatedAt: '',
        embStatus: 'Text Data',
        raw: content,
        challanData: null,
        wipData: wipData,
        hasMultiple: false,
        count: 0,
        totalQty: 0,
        allCompleted: false,
        hasPending: false,
        lastChallanHasUpdate: false
      };
    }
    
    try {
      const analysis = analyzeChallanData(challanData);
      
      let displayText = analysis.display;
      
      if (analysis.hasMultiple) {
        if (analysis.allCompleted && analysis.dateOfIssue && analysis.embUpdatedAt && analysis.lastChallanHasUpdate) {
          displayText = `${analysis.sortedChallans[0]?.number || 'Multiple'}\n${analysis.dateOfIssue} → ${analysis.embUpdatedAt}`;
        } else if (analysis.hasPending || !analysis.lastChallanHasUpdate) {
          if (analysis.dateOfIssue) {
            const completedCount = analysis.sortedChallans.filter(c => c.embCompleted === true && c.embUpdatedAt).length;
            const pendingCount = analysis.sortedChallans.filter(c => !c.embUpdatedAt || c.embCompleted === false).length;
            
            if (completedCount > 0 && pendingCount > 0) {
              displayText = `${analysis.sortedChallans[0]?.number || 'Multiple'}\n${analysis.dateOfIssue} (${completedCount}/${analysis.count})`;
            } else {
              displayText = `${analysis.sortedChallans[0]?.number || 'Multiple'}\n${analysis.dateOfIssue}`;
            }
          } else {
            displayText = `${analysis.sortedChallans[0]?.number || 'Multiple'}\nPending`;
          }
        } else if (analysis.dateOfIssue) {
          displayText = `${analysis.sortedChallans[0]?.number || 'Multiple'}\n${analysis.dateOfIssue}`;
        }
      } else if (analysis.dateOfIssue) {
        const hasUpdate = analysis.sortedChallans[0]?.embUpdatedAt ? true : false;
        if (hasUpdate && analysis.status.includes('Done')) {
          displayText = `${analysis.sortedChallans[0]?.number || 'Challan'}\n${analysis.embUpdatedAt || analysis.dateOfIssue}`;
        } else {
          displayText = `${analysis.sortedChallans[0]?.number || 'Challan'}\n${analysis.dateOfIssue}`;
        }
      }
      
      return {
        display: displayText,
        status: analysis.status,
        type: analysis.type,
        date: analysis.date,
        dateOfIssue: analysis.dateOfIssue,
        embUpdatedAt: analysis.embUpdatedAt,
        embStatus: analysis.status,
        hasMultiple: analysis.hasMultiple,
        count: analysis.count,
        totalQty: analysis.totalQty,
        allCompleted: analysis.allCompleted,
        hasPending: analysis.hasPending,
        lastChallanHasUpdate: analysis.lastChallanHasUpdate,
        challans: analysis.sortedChallans,
        wipData: wipData,
        raw: content
      };
    } catch (error) {
      console.warn('Error in parseChallanHistory:', error, content);
      
      const firstChallan = challanData[0];
      const isPrinting = firstChallan?.number?.toUpperCase().includes('PRINT');
      
      let dateOfIssue = '';
      let embUpdatedAt = '';
      
      if (firstChallan?.date) {
        if (firstChallan.date.match(/\d{1,2}\s+\w+\s+\d{4}/)) {
          dateOfIssue = firstChallan.date;
        } else {
          dateOfIssue = formatDate(firstChallan.date);
        }
      }
      
      const lastChallan = challanData[challanData.length - 1];
      const lastChallanHasUpdate = lastChallan?.embUpdatedAt ? true : false;
      
      if (lastChallanHasUpdate && challanData.length > 0) {
        const lastChallanWithUpdate = [...challanData]
          .filter(c => c.embUpdatedAt)
          .sort((a, b) => {
            try {
              return new Date(b.embUpdatedAt) - new Date(a.embUpdatedAt);
            } catch (e) {
              return 0;
            }
          })[0];
        
        if (lastChallanWithUpdate?.embUpdatedAt) {
          embUpdatedAt = formatDate(lastChallanWithUpdate.embUpdatedAt);
        }
      }
      
      const allCompleted = challanData.every(c => c.embCompleted === true && c.embUpdatedAt);
      const hasPending = challanData.some(c => !c.embUpdatedAt || c.embCompleted === false);
      
      let status = isPrinting ? 'Printing Pending' : 'Emb Pending';
      if (allCompleted && lastChallanHasUpdate && embUpdatedAt) {
        status = isPrinting ? 'Printing Done' : 'Emb Done';
      }
      
      let displayText = firstChallan?.number ? `${firstChallan.number}\n` : 'Challan\n';
      if (allCompleted && lastChallanHasUpdate && embUpdatedAt) {
        displayText += embUpdatedAt;
      } else if (dateOfIssue) {
        displayText += dateOfIssue;
      } else {
        displayText += 'Pending';
      }
      
      return {
        display: displayText,
        status: status,
        type: isPrinting ? 'printing' : 'embroidery',
        date: dateOfIssue,
        dateOfIssue: dateOfIssue,
        embUpdatedAt: embUpdatedAt,
        embStatus: status,
        raw: content,
        challanData: challanData,
        wipData: wipData,
        hasMultiple: challanData.length > 1,
        count: challanData.length,
        totalQty: 0,
        allCompleted: allCompleted,
        hasPending: hasPending,
        lastChallanHasUpdate: lastChallanHasUpdate
      };
    }
  };

  // Function to parse Status data
  const parseStatusData = (content, columnName) => {
    if (!content || content.trim() === '') {
      return {
        fullData: [],
        displayValue: '',
        count: 0,
        raw: content
      };
    }
    
    try {
      let statusData = [];
      
      if (content.trim().startsWith('[')) {
        statusData = JSON.parse(content);
      } else {
        const jsonMatch = content.match(/\[.*\]/s);
        if (jsonMatch) {
          statusData = JSON.parse(jsonMatch[0]);
        } else {
          return {
            fullData: [],
            displayValue: content,
            count: 0,
            raw: content
          };
        }
      }
      
      if (!Array.isArray(statusData) || statusData.length === 0) {
        return {
          fullData: [],
          displayValue: content.length > 50 ? content.substring(0, 50) + '...' : content,
          count: 0,
          raw: content
        };
      }
      
      const sortedData = [...statusData].sort((a, b) => {
        try {
          return new Date(b.timestamp) - new Date(a.timestamp);
        } catch (e) {
          return 0;
        }
      });
      
      const latest = sortedData[0] || {};
      
      let displayValue = '';
      
      if (columnName === 'WIP Status') {
        if (latest.remarks && latest.remarks.trim() !== '') {
          displayValue = latest.remarks;
        } else if (latest.status && latest.status.trim() !== '') {
          displayValue = latest.status;
        } else {
          displayValue = content.length > 50 ? content.substring(0, 50) + '...' : content;
        }
      } else if (columnName === 'Completed Status') {
        if (latest.timestamp) {
          displayValue = formatDate(latest.timestamp);
        } else if (latest.status && latest.status.trim() !== '') {
          displayValue = latest.status;
        } else {
          displayValue = content.length > 50 ? content.substring(0, 50) + '...' : content;
        }
      } else {
        displayValue = content.length > 50 ? content.substring(0, 50) + '...' : content;
      }
      
      return {
        fullData: sortedData,
        displayValue: displayValue,
        count: sortedData.length,
        raw: content
      };
      
    } catch (error) {
      console.warn(`Error parsing ${columnName}:`, error, content);
      return {
        fullData: [],
        displayValue: content.length > 50 ? content.substring(0, 50) + '...' : content,
        count: 0,
        raw: content
      };
    }
  };

  // Fetch data from all three Google Sheets
// Fetch data from all three Google Sheets
// Fetch data from all three Google Sheets
const fetchProductionData = async () => {
  setLoading(true);
  setError(null);
  
  try {
    console.log('Fetching data from all three sheets...');
    
    // ========== 1. FETCH JOBORDER DATA ==========
    const jobOrderRes = await fetchSheetDataFromBackend(SHEET_IDD, `${JOB_ORDER_SHEET}!A:Z`);
    if (!jobOrderRes.ok) {
      throw new Error(`HTTP error fetching JobOrder sheet data`);
    }

    const jobOrderRows = jobOrderRes.values || [];

console.log(`Received ${jobOrderRows.length} rows from JobOrder sheet`);

if (jobOrderRows.length === 0) {
  throw new Error('No data found in JobOrder sheet');
}

const jobOrderHeaders = jobOrderRows[0];
const jobOrderHeaderIndices = {};
jobOrderHeaders.forEach((header, index) => {
  const cleanHeader = header.trim();
  jobOrderHeaderIndices[cleanHeader] = index;
  if (!jobOrderHeaderIndices[cleanHeader.toUpperCase()]) {
    jobOrderHeaderIndices[cleanHeader.toUpperCase()] = index;
  }
  if (!jobOrderHeaderIndices[cleanHeader.toLowerCase()]) {
    jobOrderHeaderIndices[cleanHeader.toLowerCase()] = index;
  }
});

// DEBUG: Check what columns are available
console.log('All columns in JobOrder sheet:');
jobOrderHeaders.forEach((col, idx) => {
  console.log(`${idx}: "${col}"`);
});

// Look for Emb, Printing, and Fabric columns
const foundEmbColumn = jobOrderHeaders.find(h => 
  h.trim().toLowerCase() === 'emb' || 
  h.trim().toLowerCase() === 'embroidery' ||
  h.trim().toLowerCase().includes('emb')
);

const foundPrintingColumn = jobOrderHeaders.find(h => 
  h.trim().toLowerCase() === 'printing' || 
  h.trim().toLowerCase() === 'print' ||
  h.trim().toLowerCase().includes('print')
);

const foundFabricColumn = jobOrderHeaders.find(h => 
  h.trim().toLowerCase() === 'fabric' || 
  h.trim().toLowerCase() === 'fab' ||
  h.trim().toLowerCase().includes('fabric')
);

console.log('Found Emb column:', foundEmbColumn);
console.log('Found Printing column:', foundPrintingColumn);
console.log('Found Fabric column:', foundFabricColumn);

const processedJobOrderData = jobOrderRows.slice(1)
  .filter(row => row.length > 0 && row[0])
  .map((row, index) => {
    const rowObj = {};
    const rawObj = {};
    
    JOB_ORDER_HEADERS.forEach(header => {
      let colIndex = jobOrderHeaderIndices[header];
      if (colIndex === undefined) colIndex = jobOrderHeaderIndices[header.toLowerCase()];
      if (colIndex === undefined) colIndex = jobOrderHeaderIndices[header.toUpperCase()];
      
      let value = colIndex !== undefined && row[colIndex] !== undefined 
        ? row[colIndex] 
        : '';
      
      value = String(value).trim();
      rawObj[header] = value;
      
      switch(header) {
        case 'SEASON':
          const seasonCode = getSeasonCode(value);
          rowObj['Season'] = seasonCode;
          rowObj['Season_Full'] = value;
          break;
        case 'SECTION':
          const sectionCode = getSectionCode(value);
          rowObj['Section'] = sectionCode;
          rowObj['Section_Full'] = value;
          break;
        case 'DIRECT STITCHING':
          rowObj['Design Work'] = value;
          break;
        case 'Party Name':
          const partyInitials = getPartyInitials(value);
          rowObj['Party'] = partyInitials;
          rowObj['Party_Full'] = value;
          break;
        case 'Garment Type':
          rowObj['Item'] = value;
          break;
        case 'Lot Number':
          rowObj['Lot No'] = value;
          break;
        case 'Date':
          rowObj['Job Date'] = formatDate(value);
          rowObj['Date'] = value;
          rowObj['JobOrder_Date'] = value;
          break;
        case 'Status':
          rowObj['Status'] = value;
          rowObj['_isCancelled'] = value.toLowerCase().trim() === 'cancel';
          break;
        case 'Emb':
          rowObj['Emb'] = value;
          break;
        case 'Printing':
          rowObj['Printing'] = value;
          break;
        case 'Fabric':  // ADD FABRIC HANDLING
          rowObj['Fabric'] = value;
          rowObj['Fabric_From_JobOrder'] = value;
          break;
        case 'Size':
          rowObj['Size'] = value;
          break;
        default:
          if (header === 'Style' || header === 'Brand' || header === 'Size') {
            rowObj[header] = value;
          }
      }
    });
    
    return {
      ...rowObj,
      _id: `job-order-${index}`,
      _raw: rawObj,
      _isCancelled: rowObj['_isCancelled'] || false
    };
  });

setJobOrderData(processedJobOrderData);
console.log(`Processed ${processedJobOrderData.length} JobOrder records`);

// Debug: Check first few records for Emb/Printing/Fabric data
console.log('First 5 JobOrder records with Emb/Printing/Fabric data:');
processedJobOrderData.slice(0, 5).forEach((item, index) => {
  console.log(`Record ${index + 1}:`, {
    'Lot No': item['Lot No'],
    'Emb': item['Emb'],
    'Printing': item['Printing'],
    'Fabric': item['Fabric'],
    'D.Stitching': item['D.Stitching']
  });
});
    
    // ========== 2. FETCH INDEX DATA ==========
    const indexRes = await fetchSheetDataFromBackend(SHEET_ID, `${INDEX_SHEET}!A1:ZZZ`);
    if (!indexRes.ok) {
      throw new Error(`HTTP error fetching Index sheet data`);
    }

    const indexRows = indexRes.values || [];
    
    console.log(`Received ${indexRows.length} rows from Index sheet`);
    
    let processedIndexData = [];
    
    if (indexRows.length > 0) {
      const indexHeaders = indexRows[0];
      console.log('Actual headers in Index sheet:', indexHeaders);
      
      const indexHeaderIndices = {};
      indexHeaders.forEach((header, index) => {
        const cleanHeader = header.trim();
        indexHeaderIndices[cleanHeader] = index;
        
        const lowerHeader = cleanHeader.toLowerCase();
        const upperHeader = cleanHeader.toUpperCase();
        if (!indexHeaderIndices[lowerHeader]) indexHeaderIndices[lowerHeader] = index;
        if (!indexHeaderIndices[upperHeader]) indexHeaderIndices[upperHeader] = index;
      });
      
      processedIndexData = indexRows.slice(1)
        .filter(row => row.length > 0 && row[0])
        .map((row, rowIndex) => {
          const rowObj = {};
          const rawObj = {};
          
          // REQUIRED COLUMNS
          const lotNumberIndex = indexHeaderIndices['Lot Number'] || 
                                indexHeaderIndices['lot number'] || 
                                indexHeaderIndices['LOT NUMBER'];
          if (lotNumberIndex !== undefined && row[lotNumberIndex]) {
            const lotNumberValue = String(row[lotNumberIndex]).trim();
            rowObj['Lot Number'] = lotNumberValue;
            rawObj['Lot Number'] = lotNumberValue;
          }
          
          const savedAtIndex = indexHeaderIndices['Saved At'] || 
                             indexHeaderIndices['saved at'] || 
                             indexHeaderIndices['SAVED AT'];
          if (savedAtIndex !== undefined) {
            const savedAtValue = String(row[savedAtIndex] || '').trim();
            rawObj['Saved At'] = savedAtValue;
            rowObj['Cut Date'] = formatDate(savedAtValue);
            rowObj['Saved At'] = savedAtValue;
          }
          
          const dateOfIssueIndex = indexHeaderIndices['Date of Issue'] || 
                                 indexHeaderIndices['date of issue'] || 
                                 indexHeaderIndices['DATE OF ISSUE'];
          if (dateOfIssueIndex !== undefined) {
            const dateOfIssueValue = String(row[dateOfIssueIndex] || '').trim();
            rawObj['Date of Issue'] = dateOfIssueValue;
            rowObj['Stit Date'] = formatDate(dateOfIssueValue);
            rowObj['Issue Date'] = dateOfIssueValue;
          }
          
          const supervisorIndex = indexHeaderIndices['Supervisor'] || 
                                indexHeaderIndices['supervisor'] || 
                                indexHeaderIndices['SUPERVISOR'];
          if (supervisorIndex !== undefined) {
            const supervisorValue = String(row[supervisorIndex] || '').trim();
            rawObj['Supervisor'] = supervisorValue;
            rowObj['Stit Sup'] = supervisorValue;
            rowObj['Supervisor'] = supervisorValue;
          }
          
          const cuttingQtyIndex = indexHeaderIndices['Cutting Qty'] || 
                                indexHeaderIndices['cutting qty'] || 
                                indexHeaderIndices['CUTTING QTY'];
          if (cuttingQtyIndex !== undefined) {
            const cuttingQtyValue = String(row[cuttingQtyIndex] || '').trim();
            rawObj['Cutting Qty'] = cuttingQtyValue;
            const numValue = parseInt(cuttingQtyValue);
            rowObj['PCS'] = !isNaN(numValue) ? numValue : cuttingQtyValue;
            rowObj['Cutting Qty'] = cuttingQtyValue;
          }
          
          const challanHistoryIndex = indexHeaderIndices['CHALLAN HISTORY'] || 
                                    indexHeaderIndices['challan history'] || 
                                    indexHeaderIndices['Challan History'];
          if (challanHistoryIndex !== undefined) {
            const challanHistoryValue = String(row[challanHistoryIndex] || '').trim();
            rawObj['CHALLAN HISTORY'] = challanHistoryValue;
            
            const parsedChallan = parseChallanHistory(challanHistoryValue);
            rowObj['CHALLAN HISTORY_parsed'] = parsedChallan;
            rowObj['Emb/Print Issue'] = parsedChallan.dateOfIssue || '';
            rowObj['Emb/print Comp'] = parsedChallan.embUpdatedAt || '';
            rowObj['embStatus'] = parsedChallan.embStatus || 'No Challan';
          }
          
          const wipStatusIndex = indexHeaderIndices['WIP Status'] || 
                               indexHeaderIndices['wip status'] || 
                               indexHeaderIndices['Wip Status'];
          if (wipStatusIndex !== undefined) {
            const wipStatusValue = String(row[wipStatusIndex] || '').trim();
            rawObj['WIP Status'] = wipStatusValue;
            
            const parsedWIP = parseStatusData(wipStatusValue, 'WIP Status');
            rowObj['WIP Status_parsed'] = parsedWIP;
            rowObj['WIP Stit'] = parsedWIP.displayValue;
          }
          
          const completedStatusIndex = indexHeaderIndices['Completed Status'] || 
                                     indexHeaderIndices['completed status'] || 
                                     indexHeaderIndices['Completed status'];
          if (completedStatusIndex !== undefined) {
            const completedStatusValue = String(row[completedStatusIndex] || '').trim();
            rawObj['Completed Status'] = completedStatusValue;
            
            const parsedCompleted = parseStatusData(completedStatusValue, 'Completed Status');
            rowObj['Completed Status_parsed'] = parsedCompleted;
            rowObj['Comp Stit'] = parsedCompleted.displayValue;
          }
          
          // ADDITIONAL COLUMNS FROM INDEX
          const partyNameIndex = indexHeaderIndices['PARTY NAME'] || 
                               indexHeaderIndices['party name'] || 
                               indexHeaderIndices['Party Name'];
          if (partyNameIndex !== undefined) {
            const partyNameValue = String(row[partyNameIndex] || '').trim();
            rawObj['PARTY NAME'] = partyNameValue;
          }
          
          const brandIndex = indexHeaderIndices['BRAND'] || 
                           indexHeaderIndices['brand'] || 
                           indexHeaderIndices['Brand'];
          if (brandIndex !== undefined) {
            const brandValue = String(row[brandIndex] || '').trim();
            rawObj['BRAND'] = brandValue;
          }
          
          const seasonIndex = indexHeaderIndices['SEASON'] || 
                            indexHeaderIndices['season'] || 
                            indexHeaderIndices['Season'];
          if (seasonIndex !== undefined) {
            const seasonValue = String(row[seasonIndex] || '').trim();
            rawObj['SEASON'] = seasonValue;
          }
          
          const directStitchingIndex = indexHeaderIndices['DIRECT STITCHING'] || 
                                     indexHeaderIndices['direct stitching'] || 
                                     indexHeaderIndices['Direct Stitching'];
          if (directStitchingIndex !== undefined) {
            const directStitchingValue = String(row[directStitchingIndex] || '').trim();
            rawObj['DIRECT STITCHING'] = directStitchingValue;
          }
          
          // Image / Image URL column from Index sheet
          let imageColIndex = indexHeaderIndices['Image URL'] ?? 
                              indexHeaderIndices['image url'] ?? 
                              indexHeaderIndices['IMAGE URL'] ??
                              indexHeaderIndices['Image'] ?? 
                              indexHeaderIndices['image'] ??
                              indexHeaderIndices['IMAGE'];
          
          if (imageColIndex === undefined) {
            imageColIndex = indexHeaders.findIndex(h => h && h.trim().toLowerCase().includes('image'));
          }
          
          if (imageColIndex !== undefined && imageColIndex !== -1 && row[imageColIndex]) {
            const rawImageUrl = String(row[imageColIndex] || '').trim();
            rawObj['Image URL'] = rawImageUrl;
            const directUrl = getDirectImageUrl(rawImageUrl);
            rowObj['Image'] = directUrl;
            rowObj['Image URL'] = directUrl;
          }
          
          return {
            ...rowObj,
            _id: `index-${rowIndex}`,
            _raw: rawObj
          };
        });
    }
    
    setIndexData(processedIndexData);
    console.log(`Processed ${processedIndexData.length} Index records`);
    
    if (processedIndexData.length > 0) {
      const recordsWithImages = processedIndexData.filter(r => r['Image']);
      console.log(`Index sheet records with images: ${recordsWithImages.length} / ${processedIndexData.length}`);
      console.log('Sample Index data (first record):', {
        'Lot Number': processedIndexData[0]['Lot Number'],
        'Cut Date': processedIndexData[0]['Cut Date'],
        'Stit Date': processedIndexData[0]['Stit Date'],
        'Stit Sup': processedIndexData[0]['Stit Sup'],
        'PCS': processedIndexData[0]['PCS'],
        'Image': processedIndexData[0]['Image'] || 'No image'
      });
    }
    
    // ========== 2b. FETCH DESIGN IMAGE SHEET (FALLBACK) ==========
    try {
      const designRes = await fetchSheetDataFromBackend(SHEET_IDD, 'Design Image!A1:ZZZ');
      if (designRes.ok) {
        const designRows = designRes.values || [];
        if (designRows.length > 0) {
          const headers = designRows[0];
          const lotCol = headers.findIndex(h => h && h.trim().toLowerCase().includes('lot'));
          const imgCol = headers.findIndex(h => h && h.trim().toLowerCase().includes('image'));
          
          if (lotCol !== -1 && imgCol !== -1) {
            const designImageMap = {};
            designRows.slice(1).forEach(row => {
              const lot = String(row?.[lotCol] || '').trim();
              const img = String(row?.[imgCol] || '').trim();
              if (lot && img) {
                designImageMap[lot.toUpperCase()] = getDirectImageUrl(img);
              }
            });
            
            // Merge into processedIndexData
            processedIndexData.forEach(item => {
              const lotKey = String(item['Lot Number'] || '').trim().toUpperCase();
              if (!item['Image'] && designImageMap[lotKey]) {
                item['Image'] = designImageMap[lotKey];
                item['Image URL'] = designImageMap[lotKey];
              }
            });
          }
        }
      }
    } catch (dErr) {
      console.warn('Optional Design Image sheet fetch skipped:', dErr);
    }

    // ========== 3. FETCH ISSUES DATA ==========
    const issuesRes = await fetchSheetDataFromBackend(SHEET_IDDD, `${ISSUES_SHEET}!A:Z`);
    let processedIssuesData = [];

    if (!issuesRes.ok) {
      console.warn(`HTTP error fetching Issues sheet data`);
    } else {
      const issuesRows = issuesRes.values || [];
      
      console.log(`Received ${issuesRows.length} rows from Issues sheet`);
      
      if (issuesRows.length > 0) {
        const issuesHeaders = issuesRows[0];
        console.log('Actual headers in Issues sheet:', issuesHeaders);
        
        const issuesHeaderIndices = {};
        issuesHeaders.forEach((header, index) => {
          const cleanHeader = header.trim();
          issuesHeaderIndices[cleanHeader] = index;
          
          // Add case-insensitive variations for easier matching
          const lowerHeader = cleanHeader.toLowerCase();
          const upperHeader = cleanHeader.toUpperCase();
          if (!issuesHeaderIndices[lowerHeader]) {
            issuesHeaderIndices[lowerHeader] = index;
          }
          if (!issuesHeaderIndices[upperHeader]) {
            issuesHeaderIndices[upperHeader] = index;
          }
        });
        
        console.log('Issues header indices for reference:', issuesHeaderIndices);
        
        processedIssuesData = issuesRows.slice(1)
          .filter(row => row.length > 0 && row[0])
          .map((row, index) => {
            const rowObj = {};
            const rawObj = {};
            
            // Get Lot Number first
            const lotNumberIndex = issuesHeaderIndices['Lot Number'] || 
                                 issuesHeaderIndices['lot number'] || 
                                 issuesHeaderIndices['LOT NUMBER'];
            
            if (lotNumberIndex !== undefined && row[lotNumberIndex] !== undefined) {
              const lotNumberValue = String(row[lotNumberIndex]).trim();
              rawObj['Lot Number'] = lotNumberValue;
              rowObj['Lot Number'] = lotNumberValue;
            } else {
              // Skip if no lot number
              return null;
            }
            
            // Get Packing Date - check multiple possible column names
            let packingDateValue = '';
            
            // First try 'Packing Date'
            const packingDateIndex = issuesHeaderIndices['Packing Date'] || 
                                   issuesHeaderIndices['packing date'] || 
                                   issuesHeaderIndices['PACKING DATE'];
            
            if (packingDateIndex !== undefined && row[packingDateIndex] !== undefined) {
              packingDateValue = String(row[packingDateIndex]).trim();
            }
            
            // If Packing Date is empty, try 'Timestamp'
            if (!packingDateValue || packingDateValue === '') {
              const timestampIndex = issuesHeaderIndices['Timestamp'] || 
                                   issuesHeaderIndices['timestamp'] || 
                                   issuesHeaderIndices['TIMESTAMP'];
              
              if (timestampIndex !== undefined && row[timestampIndex] !== undefined) {
                const timestampValue = String(row[timestampIndex]).trim();
                // Extract date part from timestamp (e.g., "12/15/2025 14:19:18" -> "12/15/2025")
                if (timestampValue.includes(' ')) {
                  packingDateValue = timestampValue.split(' ')[0];
                } else {
                  packingDateValue = timestampValue;
                }
              }
            }
            
            // Format the packing date
            let formattedPackingDate = '';
            if (packingDateValue) {
              formattedPackingDate = formatDate(packingDateValue);
              rawObj['Packing Date'] = packingDateValue;
              rowObj['Packing Date'] = packingDateValue;
            }
            
            // Get Packing Supervisor
            const packingSupervisorIndex = issuesHeaderIndices['Packing Supervisor'] || 
                                         issuesHeaderIndices['packing supervisor'] || 
                                         issuesHeaderIndices['PACKING SUPERVISOR'];
            
            if (packingSupervisorIndex !== undefined && row[packingSupervisorIndex] !== undefined) {
              const packingSupervisorValue = String(row[packingSupervisorIndex]).trim();
              rawObj['Packing Supervisor'] = packingSupervisorValue;
              rowObj['Packing Supervisor'] = packingSupervisorValue;
            }
            
            // Get Total Pcs (check multiple possible column names)
            const totalPcsIndex = issuesHeaderIndices['Total Pcs'] || 
                                issuesHeaderIndices['total pcs'] || 
                                issuesHeaderIndices['TOTAL PCS'] || 
                                issuesHeaderIndices['Total PCS'] ||
                                issuesHeaderIndices['PCS'] || 
                                issuesHeaderIndices['pcs'] ||
                                issuesHeaderIndices['Total Pieces'] ||
                                issuesHeaderIndices['total pieces'];
            
            if (totalPcsIndex !== undefined && row[totalPcsIndex] !== undefined) {
              const totalPcsValue = String(row[totalPcsIndex]).trim();
              rawObj['Total Pcs'] = totalPcsValue;
              rowObj['Total Pcs'] = totalPcsValue;
            }
            
            // Get WIP Packing
            const wipPackingIndex = issuesHeaderIndices['WIP Packing'] || 
                                  issuesHeaderIndices['wip packing'] || 
                                  issuesHeaderIndices['WIP PACKING'];
            
            if (wipPackingIndex !== undefined && row[wipPackingIndex] !== undefined) {
              const wipPackingValue = String(row[wipPackingIndex]).trim();
              rawObj['WIP Packing'] = wipPackingValue;
              rowObj['WIP Packing'] = wipPackingValue;
            }
            
            // Get Packing Complete
            const packingCompleteIndex = issuesHeaderIndices['Packing Complete'] || 
                                       issuesHeaderIndices['packing complete'] || 
                                       issuesHeaderIndices['PACKING COMPLETE'];
            
            if (packingCompleteIndex !== undefined && row[packingCompleteIndex] !== undefined) {
              const packingCompleteValue = String(row[packingCompleteIndex]).trim();
              rawObj['Packing Complete'] = packingCompleteValue;
              rowObj['Packing Complete'] = packingCompleteValue;
            }
            
            // Get other columns if they exist
            const garmentTypeIndex = issuesHeaderIndices['Garment Type'] || 
                                   issuesHeaderIndices['garment type'] || 
                                   issuesHeaderIndices['GARMENT TYPE'];
            
            if (garmentTypeIndex !== undefined && row[garmentTypeIndex] !== undefined) {
              const garmentTypeValue = String(row[garmentTypeIndex]).trim();
              rawObj['Garment Type'] = garmentTypeValue;
              rowObj['Garment Type'] = garmentTypeValue;
            }
            
            // const fabricIndex = issuesHeaderIndices['Fabric'] || 
            //                   issuesHeaderIndices['fabric'] || 
            //                   issuesHeaderIndices['FABRIC'];
            
            // if (fabricIndex !== undefined && row[fabricIndex] !== undefined) {
            //   const fabricValue = String(row[fabricIndex]).trim();
            //   rawObj['Fabric'] = fabricValue;
            //   rowObj['Fabric'] = fabricValue;
            // }
            
            const styleIndex = issuesHeaderIndices['Style'] || 
                             issuesHeaderIndices['style'] || 
                             issuesHeaderIndices['STYLE'];
            
            if (styleIndex !== undefined && row[styleIndex] !== undefined) {
              const styleValue = String(row[styleIndex]).trim();
              rawObj['Style'] = styleValue;
              rowObj['Style'] = styleValue;
            }
            
            // Return the processed object
            return {
              ...rowObj,
              _id: `issues-${index}`,
              _raw: rawObj,
              // Pre-calculate display values for easier merging
              'Pkg Date': formattedPackingDate || '',
              'Pkg Sup': rowObj['Packing Supervisor'] || '',
              'WIP Pkg': rowObj['WIP Packing'] || '',
              'Pkg Comp': rowObj['Packing Complete'] || ''
            };
          })
          .filter(item => item !== null); // Remove null items (no lot number)
        
        console.log(`Processed ${processedIssuesData.length} Issues records`);
        
        if (processedIssuesData.length > 0) {
          console.log('Sample Issues data (first 3 records):');
          processedIssuesData.slice(0, 3).forEach((item, i) => {
            console.log(`Record ${i + 1}:`, {
              'Lot Number': item['Lot Number'],
              'Raw Packing Date': item['Packing Date'],
              'Formatted Pkg Date': item['Pkg Date'],
              'Packing Supervisor': item['Packing Supervisor'],
              'WIP Packing': item['WIP Packing'] ? item['WIP Packing'].substring(0, 50) + '...' : 'Empty',
              'Packing Complete': item['Packing Complete']
            });
          });
        }
      } else {
        console.log('Issues sheet is empty');
      }
    }
    
    setIssuesData(processedIssuesData);

    // ========== 4. FETCH RAWPACK DATA ==========
    let processedRawpackData = [];
    try {
      const rawpackRes = await fetchSheetDataFromBackend(SHEET_ID_RAWPACK, `${RAWPACK_SHEET}!A:Z`);
      if (rawpackRes.ok) {
        const rawpackRows = rawpackRes.values || [];
        console.log(`Received ${rawpackRows.length} rows from RAWPACK sheet`);
        
        if (rawpackRows.length > 0) {
          let headerIdx = rawpackRows.findIndex(r =>
            Array.isArray(r) && r.some(cell => {
              const str = String(cell || '').toLowerCase();
              return str.includes('lot no') || str.includes('lot number') || str.includes('item');
            })
          );
          if (headerIdx === -1) headerIdx = 0;

          const rawpackHeaders = rawpackRows[headerIdx] || [];
          console.log('Actual headers in RAWPACK sheet:', rawpackHeaders);

          const hMap = {};
          rawpackHeaders.forEach((h, idx) => {
            const clean = String(h || '').trim().toLowerCase();
            if (clean) hMap[clean] = idx;
          });

          const getColVal = (row, keyArr) => {
            for (const key of keyArr) {
              const idx = hMap[key];
              if (idx !== undefined && row[idx] !== undefined) {
                const val = String(row[idx]).trim();
                if (val && val !== '#N/A') return val;
              }
            }
            return '';
          };

          processedRawpackData = rawpackRows.slice(headerIdx + 1)
            .filter(r => r && r.length > 0)
            .map((row, index) => {
              const lotNo = getColVal(row, ['lot no.', 'lot no', 'lot no.2', 'lot number', 'lot']);
              if (!lotNo) return null;

              const supervisor = getColVal(row, ['packing person', 'packing supervisor', 'packing person name', 'supervisior', 'supervisor']);
              const pkgIssueDateRaw = getColVal(row, ['date of packing issue', 'sticker issue', 'date of packing', 'packing date', 'timestamp']);
              const pkgCompDateRaw = getColVal(row, ['date of packing complete', 'packing complete', 'complete date']);
              const remarks = getColVal(row, ['remarks', 'wip packing', 'wip']);
              const pcs = getColVal(row, ['pcs', 'total pcs']);

              return {
                _id: `rawpack-${index}`,
                'Lot Number': lotNo,
                'Packing Supervisor': supervisor,
                'Pkg Sup': supervisor,
                'Packing Date': pkgIssueDateRaw,
                'Pkg Date': pkgIssueDateRaw ? formatDate(pkgIssueDateRaw) : '',
                'Packing Complete': pkgCompDateRaw,
                'Pkg Comp': pkgCompDateRaw ? formatDate(pkgCompDateRaw) : '',
                'WIP Packing': remarks,
                'WIP Pkg': remarks,
                'Total Pcs': pcs
              };
            })
            .filter(Boolean);

          console.log(`Processed ${processedRawpackData.length} RAWPACK records`);
        }
      }
    } catch (rErr) {
      console.warn('RAWPACK sheet fetch error:', rErr);
    }

    setRawpackData(processedRawpackData);
    
    mergeData(processedJobOrderData, processedIndexData, processedIssuesData, processedRawpackData);
    
  } catch (err) {
    console.error('Error fetching data:', err);
    
    if (err.message.includes('404')) {
      setError('Google Sheet not found. Please check your SHEET_ID.');
    } else if (err.message.includes('403')) {
      setError('API Key rejected. Please check your API_KEY.');
    } else if (err.message.includes('Failed to fetch')) {
      setError('Network error. Please check your internet connection.');
    } else {
      setError(`Error loading data: ${err.message}`);
    }
    
    setData([]);
    setJobOrderData([]);
    setIndexData([]);
    setIssuesData([]);
    setRawpackData([]);
  } finally {
    setLoading(false);
  }
};

  // Merge data from all sheets
// Merge data from all sheets
// Merge data from all sheets
const mergeData = (jobOrderData, indexData, issuesData, rawpackData = []) => {
  console.log('Merging data from all sheets...');
  console.log('JobOrder records:', jobOrderData?.length || 0);
  console.log('Index records:', indexData?.length || 0);
  console.log('Issues records:', issuesData?.length || 0);
  console.log('RAWPACK records:', rawpackData?.length || 0);
  
  // Filter out cancelled lots from JobOrder data
  const validJobOrderData = jobOrderData.filter(item => !item._isCancelled);
  console.log(`Valid JobOrder records (excluding cancelled): ${validJobOrderData.length}`);
  console.log(`Cancelled lots filtered out: ${jobOrderData.length - validJobOrderData.length}`);
  
  if ((!issuesData || issuesData.length === 0) && (!rawpackData || rawpackData.length === 0)) {
    console.log('WARNING: Packing data (Issues & RAWPACK) is empty or undefined!');
    
    const mergedData = validJobOrderData.map(jobOrderItem => {
      const lotNumber = jobOrderItem['Lot No'];
      const indexItem = indexData.find(item => 
        item['Lot Number'] && 
        String(item['Lot Number']).trim().toUpperCase() === String(lotNumber).trim().toUpperCase()
      );
      
      const mergedItem = {};
      
      // Add JobOrder data INCLUDING Emb, Printing, and Fabric
      DISPLAY_HEADERS.forEach(header => {
        if (jobOrderItem[header] !== undefined) {
          mergedItem[header] = jobOrderItem[header];
        }
      });
      
      Object.keys(jobOrderItem).forEach(key => {
        if (!mergedItem[key] && key !== '_id' && key !== '_raw') {
          mergedItem[key] = jobOrderItem[key];
        }
      });
      
      mergedItem._id = jobOrderItem._id;
      mergedItem._raw = {
        ...jobOrderItem._raw
      };
      
      // Specifically add Job Date from JobOrder
      if (jobOrderItem['Job Date']) {
        mergedItem['Job Date'] = jobOrderItem['Job Date'];
        mergedItem._raw['Date'] = jobOrderItem._raw['Date'];
      }
      
      // Add Status from JobOrder
      if (jobOrderItem['Status']) {
        mergedItem['Status'] = jobOrderItem['Status'];
        mergedItem._isCancelled = jobOrderItem._isCancelled;
      }
      
      // Add Emb and Printing from JobOrder
      if (jobOrderItem['Emb']) {
        mergedItem['Emb'] = jobOrderItem['Emb'];
      }
      
      if (jobOrderItem['Printing']) {
        mergedItem['Printing'] = jobOrderItem['Printing'];
      }
      
      // ADD FABRIC FROM JOBORDER SHEET
      if (jobOrderItem['Fabric']) {
        mergedItem['Fabric'] = jobOrderItem['Fabric'];
        mergedItem._raw['Fabric'] = jobOrderItem['Fabric'];
        mergedItem['Fabric_From_JobOrder'] = jobOrderItem['Fabric'];
      } else {
        mergedItem['Fabric'] = '';
        mergedItem._raw['Fabric'] = '';
        mergedItem['Fabric_From_JobOrder'] = '';
      }
      
      // Add Index data if available
      if (indexItem) {
        // Add Cut Date
        if (indexItem['Cut Date']) {
          mergedItem['Cut Date'] = indexItem['Cut Date'];
          mergedItem._raw['Saved At'] = indexItem._raw['Saved At'];
        }
        
        // Add Stit Date
        if (indexItem['Stit Date']) {
          mergedItem['Stit Date'] = indexItem['Stit Date'];
          mergedItem._raw['Date of Issue'] = indexItem._raw['Date of Issue'];
        }
        
        // Add Stit Sup
        if (indexItem['Stit Sup']) {
          mergedItem['Stit Sup'] = indexItem['Stit Sup'];
          mergedItem._raw['Supervisor'] = indexItem._raw['Supervisor'];
        }
        
        // Add PCS
        if (indexItem['PCS']) {
          mergedItem['PCS'] = indexItem['PCS'];
          mergedItem._raw['Cutting Qty'] = indexItem._raw['Cutting Qty'];
        }
        
        // Add Emb/Print Issue and Emb/print Comp
        if (indexItem['CHALLAN HISTORY_parsed']) {
          mergedItem['CHALLAN HISTORY_parsed'] = indexItem['CHALLAN HISTORY_parsed'];
          mergedItem['Emb/Print Issue'] = indexItem['Emb/Print Issue'] || '';
          mergedItem['Emb/print Comp'] = indexItem['Emb/print Comp'] || '';
          mergedItem._raw['CHALLAN HISTORY'] = indexItem._raw['CHALLAN HISTORY'];
          mergedItem['embStatus'] = indexItem['embStatus'] || 'No Challan';
        }
        
        // Add WIP Stit
        if (indexItem['WIP Status_parsed']) {
          mergedItem['WIP Status_parsed'] = indexItem['WIP Status_parsed'];
          mergedItem['WIP Stit'] = indexItem['WIP Stit'];
          mergedItem._raw['WIP Status'] = indexItem._raw['WIP Status'];
        }
        
        // Add Comp Stit
        if (indexItem['Completed Status_parsed']) {
          mergedItem['Completed Status_parsed'] = indexItem['Completed Status_parsed'];
          mergedItem['Comp Stit'] = indexItem['Comp Stit'];
          mergedItem._raw['Completed Status'] = indexItem._raw['Completed Status'];
        }
        
        // If Comp Stit is completed, override WIP Stit with "Stitching Done"
        const fallbackCompStitVal = (mergedItem['Comp Stit'] || '').toString().trim();
        if (fallbackCompStitVal && 
            fallbackCompStitVal !== '-' && 
            fallbackCompStitVal.toLowerCase() !== 'invalid date' && 
            !fallbackCompStitVal.toLowerCase().includes('pending')) {
          mergedItem['WIP Stit'] = 'Stitching Done';
        }
        
        if (indexItem['Image'] || indexItem['Image URL']) {
          mergedItem['Image'] = indexItem['Image'] || indexItem['Image URL'];
          mergedItem['Image URL'] = indexItem['Image URL'] || indexItem['Image'];
        } else {
          mergedItem['Image'] = '';
          mergedItem['Image URL'] = '';
        }
        
        mergedItem._hasIndexData = true;
      } else {
        // Set defaults if no Index data
        mergedItem['Image'] = '';
        mergedItem['Image URL'] = '';
        mergedItem['Cut Date'] = '';
        mergedItem._raw['Saved At'] = '';
        mergedItem['Stit Date'] = '';
        mergedItem._raw['Date of Issue'] = '';
        mergedItem['Stit Sup'] = '';
        mergedItem._raw['Supervisor'] = '';
        mergedItem['PCS'] = '';
        mergedItem._raw['Cutting Qty'] = '';
        
        const parsedChallan = parseChallanHistory('');
        mergedItem['CHALLAN HISTORY_parsed'] = parsedChallan;
        mergedItem['Emb/Print Issue'] = '';
        mergedItem['Emb/print Comp'] = '';
        mergedItem._raw['CHALLAN HISTORY'] = '';
        mergedItem['embStatus'] = 'No Challan';
        
        mergedItem['WIP Stit'] = '';
        mergedItem._raw['WIP Status'] = '';
        mergedItem['Comp Stit'] = '';
        mergedItem._raw['Completed Status'] = '';
        
        mergedItem._hasIndexData = false;
      }
      
      // Set defaults for packing columns
      mergedItem['Pkg Sup'] = '';
      mergedItem._raw['Packing Supervisor'] = '';
      mergedItem['Pkg Date'] = '';
      mergedItem._raw['Packing Date'] = '';
      mergedItem['WIP Pkg'] = '';
      mergedItem._raw['WIP Packing'] = '';
      mergedItem['Pkg Comp'] = '';
      mergedItem._raw['Packing Complete'] = '';
      mergedItem._packingSource = 'No Data';
      mergedItem._issuesMatch = false;
      
      // Calculate all days for this row
      const days = calculateAllDays(mergedItem);
      mergedItem['Cut Days'] = days.cutDays;
      mergedItem['Emb/Print Days'] = days.embPrintDays;
      mergedItem['Stit Days'] = days.stitDays;
      mergedItem['Pkg Days'] = days.pkgDays;
      mergedItem['Cut To Emb/Print'] = days.cutToEmbPrint;
      mergedItem['Emb/Print To Stit'] = days.embPrintToStit;
      mergedItem['Stit To Check Pack'] = days.stitToCheckPack;
      
      // Ensure all display headers have values
      DISPLAY_HEADERS.forEach(header => {
        if (mergedItem[header] === undefined) {
          mergedItem[header] = '';
        }
      });
      
      return mergedItem;
    });
    
    console.log(`Created ${mergedData.length} merged records (cancelled lots excluded)`);
    
    // Debug: Check first few merged records
    console.log('First 3 merged records (without Issues data):');
    mergedData.slice(0, 3).forEach((item, index) => {
      console.log(`Record ${index + 1}:`, {
        'Lot No': item['Lot No'],
        'Fabric': item['Fabric'],
        'Emb': item['Emb'],
        'Printing': item['Printing'],
        'D.Stitching': item['D.Stitching']
      });
    });
    
    setData(mergedData);
    return;
  }
  
  // Create maps for quick lookup
  const indexMap = new Map();
  indexData.forEach(item => {
    const lotNumber = item['Lot Number'];
    if (lotNumber && String(lotNumber).trim() !== '') {
      const normalizedKey = String(lotNumber).trim().toUpperCase();
      indexMap.set(normalizedKey, item);
    }
  });
  
  const issuesMap = new Map();
  if (Array.isArray(issuesData)) {
    issuesData.forEach(item => {
      const lotNumber = item['Lot Number'];
      if (lotNumber && String(lotNumber).trim() !== '') {
        const normalizedKey = String(lotNumber).trim().toUpperCase();
        const existingItem = issuesMap.get(normalizedKey);
        if (existingItem) {
          const currentDate = parseDate(item['Packing Date']) || parseDate(item['Timestamp']);
          const existingDate = parseDate(existingItem['Packing Date']) || parseDate(existingItem['Timestamp']);
          if (currentDate && existingDate && currentDate > existingDate) {
            issuesMap.set(normalizedKey, item);
          }
        } else {
          issuesMap.set(normalizedKey, item);
        }
      }
    });
  }

  const rawpackMap = new Map();
  if (Array.isArray(rawpackData)) {
    rawpackData.forEach(item => {
      const lotNumber = item['Lot Number'];
      if (lotNumber && String(lotNumber).trim() !== '') {
        const normalizedKey = String(lotNumber).trim().toUpperCase();
        rawpackMap.set(normalizedKey, item);
      }
    });
  }
  
  // Merge validJobOrderData (excluding cancelled) with Index, Issues, and RAWPACK data
  const mergedData = validJobOrderData.map((jobOrderItem) => {
    const lotNumber = jobOrderItem['Lot No'];
    let indexItem = null;
    let issuesItem = null;
    let rawpackItem = null;
    
    if (lotNumber && String(lotNumber).trim() !== '') {
      const normalizedKey = String(lotNumber).trim().toUpperCase();
      indexItem = indexMap.get(normalizedKey);
      issuesItem = issuesMap.get(normalizedKey);
      rawpackItem = rawpackMap.get(normalizedKey);
    }
    
    const mergedItem = {};
    
    // Add JobOrder data INCLUDING Emb, Printing, and Fabric
    DISPLAY_HEADERS.forEach(header => {
      if (jobOrderItem[header] !== undefined) {
        mergedItem[header] = jobOrderItem[header];
      }
    });
    
    Object.keys(jobOrderItem).forEach(key => {
      if (!mergedItem[key] && key !== '_id' && key !== '_raw') {
        mergedItem[key] = jobOrderItem[key];
      }
    });
    
    mergedItem._id = jobOrderItem._id;
    mergedItem._raw = {
      ...jobOrderItem._raw
    };
    
    // Specifically add Job Date from JobOrder
    if (jobOrderItem['Job Date']) {
      mergedItem['Job Date'] = jobOrderItem['Job Date'];
      mergedItem._raw['Date'] = jobOrderItem._raw['Date'];
    }
    
    // Add Status from JobOrder
    if (jobOrderItem['Status']) {
      mergedItem['Status'] = jobOrderItem['Status'];
      mergedItem._isCancelled = jobOrderItem._isCancelled;
    }
    
    // Add Emb and Printing from JobOrder
    if (jobOrderItem['Emb']) {
      mergedItem['Emb'] = jobOrderItem['Emb'];
    }
    
    if (jobOrderItem['Printing']) {
      mergedItem['Printing'] = jobOrderItem['Printing'];
    }
    
    // ADD FABRIC FROM JOBORDER SHEET (BEFORE Index and Issues data)
    if (jobOrderItem['Fabric']) {
      mergedItem['Fabric'] = jobOrderItem['Fabric'];
      mergedItem._raw['Fabric'] = jobOrderItem['Fabric'];
      mergedItem['Fabric_From_JobOrder'] = jobOrderItem['Fabric'];
    } else {
      mergedItem['Fabric'] = '';
      mergedItem._raw['Fabric'] = '';
      mergedItem['Fabric_From_JobOrder'] = '';
    }
    
    // Add Index data if available
    if (indexItem) {
      // Add Cut Date
      if (indexItem['Cut Date']) {
        mergedItem['Cut Date'] = indexItem['Cut Date'];
        mergedItem._raw['Saved At'] = indexItem._raw['Saved At'];
      }
      
      // Add Stit Date
      if (indexItem['Stit Date']) {
        mergedItem['Stit Date'] = indexItem['Stit Date'];
        mergedItem._raw['Date of Issue'] = indexItem._raw['Date of Issue'];
      }
      
      // Add Stit Sup
      if (indexItem['Stit Sup']) {
        mergedItem['Stit Sup'] = indexItem['Stit Sup'];
        mergedItem._raw['Supervisor'] = indexItem._raw['Supervisor'];
      }
      
      // Add PCS
      if (indexItem['PCS']) {
        mergedItem['PCS'] = indexItem['PCS'];
        mergedItem._raw['Cutting Qty'] = indexItem._raw['Cutting Qty'];
      }
      
      // Add Emb/Print Issue and Emb/print Comp
      if (indexItem['CHALLAN HISTORY_parsed']) {
        mergedItem['CHALLAN HISTORY_parsed'] = indexItem['CHALLAN HISTORY_parsed'];
        mergedItem['Emb/Print Issue'] = indexItem['Emb/Print Issue'] || '';
        mergedItem['Emb/print Comp'] = indexItem['Emb/print Comp'] || '';
        mergedItem._raw['CHALLAN HISTORY'] = indexItem._raw['CHALLAN HISTORY'];
        mergedItem['embStatus'] = indexItem['embStatus'] || 'No Challan';
      }
      
      // Add WIP Stit
      if (indexItem['WIP Status_parsed']) {
        mergedItem['WIP Status_parsed'] = indexItem['WIP Status_parsed'];
        mergedItem['WIP Stit'] = indexItem['WIP Stit'];
        mergedItem._raw['WIP Status'] = indexItem._raw['WIP Status'];
      }
      
      // Add Comp Stit
      if (indexItem['Completed Status_parsed']) {
        mergedItem['Completed Status_parsed'] = indexItem['Completed Status_parsed'];
        mergedItem['Comp Stit'] = indexItem['Comp Stit'];
        mergedItem._raw['Completed Status'] = indexItem._raw['Completed Status'];
      }
      
      // If Comp Stit is completed, override WIP Stit with "Stitching Done"
      const compStitVal = (mergedItem['Comp Stit'] || '').toString().trim();
      if (compStitVal && 
          compStitVal !== '-' && 
          compStitVal.toLowerCase() !== 'invalid date' && 
          !compStitVal.toLowerCase().includes('pending')) {
        mergedItem['WIP Stit'] = 'Stitching Done';
      }
      
      // Add Image from Index sheet
      if (indexItem['Image'] || indexItem['Image URL']) {
        mergedItem['Image'] = indexItem['Image'] || indexItem['Image URL'];
        mergedItem['Image URL'] = indexItem['Image URL'] || indexItem['Image'];
        mergedItem._raw['Image URL'] = indexItem._raw?.['Image URL'] || indexItem['Image'];
      } else {
        mergedItem['Image'] = '';
        mergedItem['Image URL'] = '';
      }
      
      mergedItem._hasIndexData = true;
    } else {
      // Set defaults if no Index data
      mergedItem['Image'] = '';
      mergedItem['Image URL'] = '';
      mergedItem['Cut Date'] = '';
      mergedItem._raw['Saved At'] = '';
      mergedItem['Stit Date'] = '';
      mergedItem._raw['Date of Issue'] = '';
      mergedItem['Stit Sup'] = '';
      mergedItem._raw['Supervisor'] = '';
      mergedItem['PCS'] = '';
      mergedItem._raw['Cutting Qty'] = '';
      
      const parsedChallan = parseChallanHistory('');
      mergedItem['CHALLAN HISTORY_parsed'] = parsedChallan;
      mergedItem['Emb/Print Issue'] = '';
      mergedItem['Emb/print Comp'] = '';
      mergedItem._raw['CHALLAN HISTORY'] = '';
      mergedItem['embStatus'] = 'No Challan';
      
      mergedItem['WIP Stit'] = '';
      mergedItem._raw['WIP Status'] = '';
      mergedItem['Comp Stit'] = '';
      mergedItem._raw['Completed Status'] = '';
      
      mergedItem._hasIndexData = false;
    }
    
    // Add Packing data from Issues and/or RAWPACK sheet if available
    if (issuesItem || rawpackItem) {
      const pkgSupVal = issuesItem?.['Packing Supervisor'] || rawpackItem?.['Packing Supervisor'] || rawpackItem?.['Pkg Sup'] || '';
      mergedItem['Pkg Sup'] = pkgSupVal;
      mergedItem._raw['Packing Supervisor'] = pkgSupVal;
      
      let packingDateValue = issuesItem?.['Packing Date'] || issuesItem?.['Timestamp'] || rawpackItem?.['Packing Date'] || '';
      if (packingDateValue) {
        const parsedDate = parseDate(packingDateValue);
        if (parsedDate) {
          packingDateValue = formatDate(parsedDate);
        }
      }
      mergedItem['Pkg Date'] = packingDateValue;
      mergedItem._raw['Packing Date'] = packingDateValue;
      
      const wipPackingRaw = issuesItem?.['WIP Packing'] || rawpackItem?.['WIP Packing'] || rawpackItem?.['WIP Pkg'] || '';
      const parsedWIPPacking = parseWIPPackingData(wipPackingRaw);
      
      mergedItem['WIP Pkg'] = parsedWIPPacking.displayValue;
      mergedItem['WIP Pkg._parsed'] = parsedWIPPacking;
      mergedItem._raw['WIP Packing'] = wipPackingRaw;
      
      let packingCompleteValue = issuesItem?.['Packing Complete'] || rawpackItem?.['Packing Complete'] || rawpackItem?.['Pkg Comp'] || '';

      try {
        if (packingCompleteValue.trim().startsWith('[')) {
          const packingCompleteData = JSON.parse(packingCompleteValue);
          if (Array.isArray(packingCompleteData) && packingCompleteData.length > 0) {
            const sortedPackingData = [...packingCompleteData].sort((a, b) => {
              try {
                return new Date(b.timestamp) - new Date(a.timestamp);
              } catch (e) {
                return 0;
              }
            });
            
            const latestPacking = sortedPackingData[0];
            if (latestPacking?.timestamp) {
              packingCompleteValue = formatDate(latestPacking.timestamp);
            }
          }
        }
      } catch (error) {
        if (packingCompleteValue) {
          const parsedCompleteDate = parseDate(packingCompleteValue);
          if (parsedCompleteDate) {
            packingCompleteValue = formatDate(parsedCompleteDate);
          }
        }
      }

      if (!packingCompleteValue && parsedWIPPacking.fullData.length > 0) {
        const completedEntries = parsedWIPPacking.fullData.filter(item => 
          item.status && item.status.toLowerCase().includes('complete') ||
          item.remarks && item.remarks.toLowerCase().includes('complete')
        );
        
        if (completedEntries.length > 0) {
          completedEntries.sort((a, b) => {
            try {
              return new Date(b.timestamp) - new Date(a.timestamp);
            } catch (e) {
              return 0;
            }
          });
          
          const latestCompleted = completedEntries[0];
          if (latestCompleted.timestamp) {
            packingCompleteValue = formatDate(latestCompleted.timestamp);
          }
        }
      }
      
      mergedItem['Pkg Comp'] = packingCompleteValue;
      mergedItem._raw['Packing Complete'] = packingCompleteValue;
      
      mergedItem._packingSource = issuesItem && rawpackItem ? 'Issues & RAWPACK' : (issuesItem ? 'Issues Sheet' : 'RAWPACK Sheet');
      mergedItem._issuesMatch = true;
      mergedItem._isHoldLot = parsedWIPPacking.isHold || false;
    } else {
      mergedItem['Pkg Sup'] = '';
      mergedItem._raw['Packing Supervisor'] = '';
      mergedItem['Pkg Date'] = '';
      mergedItem._raw['Packing Date'] = '';
      mergedItem['WIP Pkg'] = '';
      mergedItem._raw['WIP Packing'] = '';
      mergedItem['Pkg Comp'] = '';
      mergedItem._raw['Packing Complete'] = '';
      mergedItem._packingSource = 'No Data';
      mergedItem._issuesMatch = false;
      mergedItem._isHoldLot = false;
    }
    
    // Calculate all days for this row
    const days = calculateAllDays(mergedItem);
    mergedItem['Cut Days'] = days.cutDays;
    mergedItem['Emb/Print Days'] = days.embPrintDays;
    mergedItem['Stit Days'] = days.stitDays;
    mergedItem['Pkg Days'] = days.pkgDays;
    mergedItem['Cut To Emb/Print'] = days.cutToEmbPrint;
    mergedItem['Emb/Print To Stit'] = days.embPrintToStit;
    mergedItem['Stit To Check Pack'] = days.stitToCheckPack;
    
    // Ensure all display headers have values
    DISPLAY_HEADERS.forEach(header => {
      if (mergedItem[header] === undefined) {
        mergedItem[header] = '';
      }
    });
    
    return mergedItem;
  });
  
  console.log(`Merged ${mergedData.length} records (cancelled lots excluded)`);
  
  // Debug: Check first few merged records with Fabric data
  console.log('First 5 merged records with Fabric data from JobOrder:');
  mergedData.slice(0, 5).forEach((item, index) => {
    console.log(`Record ${index + 1}:`, {
      'Lot No': item['Lot No'],
      'Fabric': item['Fabric'],
      'Fabric Source': item['Fabric_From_JobOrder'] ? 'JobOrder' : 'Not from JobOrder',
      'Emb': item['Emb'],
      'Printing': item['Printing'],
      'D.Stitching': item['D.Stitching'],
      'Has Fabric property': 'Fabric' in item,
      'Fabric value': item['Fabric']
    });
  });
  
  setData(mergedData);
};

  useEffect(() => {
    fetchProductionData();
  }, []);
  // Close fabric dropdown when clicking outside
useEffect(() => {
  const handleClickOutside = (event) => {
    if (fabricDropdownRef.current && !fabricDropdownRef.current.contains(event.target)) {
      setShowFabricDropdown(false);
    }
  };
  
  document.addEventListener('mousedown', handleClickOutside);
  return () => {
    document.removeEventListener('mousedown', handleClickOutside);
  };
}, []);

const filteredData = data.filter(item => {
  const rawItem = item._raw || item;
  
  // Search term - partial match across all fields
  if (searchTerm) {
    const matchesSearch = Object.values(item).some(value => 
      value && String(value).toLowerCase().includes(searchTerm.toLowerCase())
    );
    if (!matchesSearch) return false;
  }
  
  // Garment Type/Item filter (Multi-Select)
  if (garmentTypeFilter.length > 0) {
    const garmentTypeValue = (
      item['Garment Type'] || 
      item['Item'] || 
      rawItem['Garment Type'] || 
      rawItem['Item'] || 
      ''
    ).toString().trim();
    
    if (!garmentTypeFilter.some(sel => sel.toLowerCase() === garmentTypeValue.toLowerCase())) {
      return false;
    }
  }
  
  // Brand filter (Multi-Select)
  if (brandFilter.length > 0) {
    const brandValue = (
      item['Brand'] || 
      rawItem['Brand'] || 
      rawItem['BRAND'] || 
      ''
    ).toString().trim();
    
    if (!brandFilter.some(sel => sel.toLowerCase() === brandValue.toLowerCase())) {
      return false;
    }
  }
  
  // Style filter (Multi-Select)
  if (styleFilter.length > 0) {
    const styleValue = (
      item['Style'] || 
      rawItem['Style'] || 
      rawItem['STYLE'] || 
      ''
    ).toString().trim();
    
    if (!styleFilter.some(sel => sel.toLowerCase() === styleValue.toLowerCase())) {
      return false;
    }
  }
  
  // Party Name filter (Multi-Select)
  if (partyNameFilter.length > 0) {
    const partyValue = (
      rawItem['Party Name'] || 
      rawItem['PARTY NAME'] || 
      rawItem['Party'] || 
      rawItem['Party_Full'] || 
      item['Party'] || 
      item['Party_Full'] || 
      ''
    ).toString().trim();
    
    if (!partyNameFilter.some(sel => sel.toLowerCase() === partyValue.toLowerCase())) {
      return false;
    }
  }
  
  // Season filter (Multi-Select)
  if (seasonFilter.length > 0) {
    const seasonValue = (
      item['Season_Full'] || 
      item['Season'] || 
      rawItem['SEASON'] || 
      rawItem['Season'] || 
      rawItem['season'] || 
      item['SEASON'] ||
      ''
    ).toString().trim();
    
    if (!seasonFilter.some(sel => sel.toLowerCase() === seasonValue.toLowerCase())) {
      return false;
    }
  }
  
  // Section filter (Multi-Select)
  if (sectionFilter.length > 0) {
    const sectionValue = (
      item['Section_Full'] || 
      item['Section'] || 
      rawItem['SECTION'] || 
      rawItem['Section'] || 
      rawItem['section'] || 
      item['SECTION'] ||
      ''
    ).toString().trim();
    
    if (!sectionFilter.some(sel => sel.toLowerCase() === sectionValue.toLowerCase())) {
      return false;
    }
  }
  
  // Direct Stitching filter (Multi-Select)
  if (directStitchingFilter.length > 0) {
    const directStitchingValue = (
      rawItem['DIRECT STITCHING'] || 
      rawItem['Direct Stitching'] || 
      rawItem['D.Stitching'] || 
      rawItem['Design Work'] || 
      item['Design Work'] || 
      ''
    ).toString().trim().toUpperCase();

    const matchesDS = directStitchingFilter.some(filterValue => {
      const fUpper = filterValue.toUpperCase();
      if (fUpper === 'YES') {
        return directStitchingValue.includes('YES') || directStitchingValue === 'Y';
      } else if (fUpper === 'NO') {
        return directStitchingValue.includes('NO') || directStitchingValue === 'N';
      }
      return directStitchingValue === fUpper;
    });
    
    if (!matchesDS) return false;
  }
  
  // STIT SUPERVISOR filter (Multi-Select)
  if (supervisorFilter.length > 0) {
    const supervisorValue = (
      item['Stit Sup'] || 
      rawItem['Supervisor'] || 
      rawItem['supervisor'] || 
      item['Supervisor'] || 
      ''
    ).toString().trim().toLowerCase();
    
    if (!supervisorFilter.some(sel => supervisorValue.includes(sel.toLowerCase()))) {
      return false;
    }
  }
  
  // Emb/Print Completion filter (Multi-Select)
  if (embPrintCompFilter.length > 0) {
    const embPrintCompValue = (item['Emb/print Comp'] || '').toString().trim().toLowerCase();
    
    const matchesEmb = embPrintCompFilter.some(f => {
      if (f === 'Complete') {
        return embPrintCompValue && 
               embPrintCompValue !== '' && 
               embPrintCompValue !== 'invalid date' &&
               embPrintCompValue !== '-' &&
               !embPrintCompValue.includes('pending');
      } else if (f === 'Pending') {
        return !embPrintCompValue || 
               embPrintCompValue === '' || 
               embPrintCompValue === 'invalid date' ||
               embPrintCompValue === '-' ||
               embPrintCompValue.includes('pending');
      }
      return false;
    });
    
    if (!matchesEmb) return false;
  }
  
  // Comp Stit filter (Multi-Select)
  if (compStatusFilter.length > 0) {
    const compStatusValue = (item['Comp Stit'] || '').toString().trim().toLowerCase();
    
    const matchesComp = compStatusFilter.some(f => {
      if (f === 'Complete') {
        return compStatusValue && 
               compStatusValue !== '' && 
               compStatusValue !== 'invalid date' &&
               compStatusValue !== '-' &&
               !compStatusValue.includes('pending');
      } else if (f === 'Pending') {
        return !compStatusValue || 
               compStatusValue === '' || 
               compStatusValue === 'invalid date' ||
               compStatusValue === '-' ||
               compStatusValue.includes('pending');
      }
      return false;
    });

    if (!matchesComp) return false;
  }
  
  // Packing Supervisor filter (Multi-Select)
  if (pkgSupervisorFilter.length > 0) {
    const pkgSupervisorValue = (
      item['Pkg Sup'] || 
      rawItem['Packing Supervisor'] || 
      rawItem['packing supervisor'] || 
      ''
    ).toString().trim().toLowerCase();
    
    if (!pkgSupervisorFilter.some(sel => sel.toLowerCase() === pkgSupervisorValue)) {
      return false;
    }
  }
  
  // Packing Completion filter (Multi-Select)
  if (pkgCompFilter.length > 0) {
    const pkgCompValue = (item['Pkg Comp'] || '').toString().trim().toLowerCase();
    
    const matchesPkg = pkgCompFilter.some(f => {
      if (f === 'Complete') {
        return pkgCompValue && 
               pkgCompValue !== '' && 
               pkgCompValue !== 'invalid date' &&
               pkgCompValue !== '-' &&
               !pkgCompValue.includes('pending');
      } else if (f === 'Pending') {
        return !pkgCompValue || 
               pkgCompValue === '' || 
               pkgCompValue === 'invalid date' ||
               pkgCompValue === '-' ||
               pkgCompValue.includes('pending');
      }
      return false;
    });

    if (!matchesPkg) return false;
  }
  
  // Hold Lots filter
  if (holdLotsFilter) {
    const wipPkgParsed = item['WIP Pkg._parsed'];
    const isHoldLot = wipPkgParsed?.isHold || item._isHoldLot || false;
    
    if (!isHoldLot) {
      return false;
    }
  }
  
  // FABRIC FILTER
  if (fabricFilter.length > 0) {
    let fabricValue = '';
    
    if (item['Fabric']) {
      fabricValue = item['Fabric'];
    } else if (item._raw && item._raw['Fabric']) {
      fabricValue = item._raw['Fabric'];
    } else if (item._issuesMatch) {
      const issuesItem = issuesData.find(issues => 
        issues['Lot Number'] === item['Lot No']
      );
      if (issuesItem && issuesItem['Fabric']) {
        fabricValue = issuesItem['Fabric'];
      }
    }
    
    fabricValue = fabricValue.toString().trim().toLowerCase();
    
    const matchesFabric = fabricFilter.some(selectedFabric => 
      fabricValue === selectedFabric.toLowerCase()
    );
    
    if (!matchesFabric) {
      return false;
    }
  }

  // DATE RANGE FILTER (Job Date / Cut Date / Both)
  if (startDateFilter || endDateFilter) {
    const start = startDateFilter ? new Date(startDateFilter + 'T00:00:00') : null;
    const end = endDateFilter ? new Date(endDateFilter + 'T23:59:59') : null;

    const getRawDateValues = () => {
      if (dateFilterField === 'Job Date') {
        return [item['Job Date'], item._raw?.['Date'], item['Date']];
      } else if (dateFilterField === 'Cut Date') {
        return [item['Cut Date'], item._raw?.['Saved At'], item['Saved At']];
      } else {
        return [
          item['Job Date'], item._raw?.['Date'], item['Date'],
          item['Cut Date'], item._raw?.['Saved At'], item['Saved At']
        ];
      }
    };

    const candidateDates = getRawDateValues()
      .map(val => parseDateForFilter(val))
      .filter(Boolean);

    if (candidateDates.length === 0) {
      return false;
    }

    const matchesRange = candidateDates.some(d => {
      if (start && d < start) return false;
      if (end && d > end) return false;
      return true;
    });

    if (!matchesRange) {
      return false;
    }
  }
  
  return true;
});

const getUniqueValues = (fieldName) => {
  const displayToDataMap = {
    'Garment Type': ['Garment Type', 'Item'],
    'Item': ['Garment Type', 'Item'],
    'Brand': ['Brand', 'BRAND'],
    'Style': ['Style', 'STYLE'],
    'Party Name': ['Party Name', 'PARTY NAME', 'Party', 'Party_Full'],
    'Party': ['Party Name', 'PARTY NAME', 'Party', 'Party_Full'],
    'Season': ['Season_Full', 'Season', 'SEASON'],  // Prioritize Season_Full
    'Section': ['Section_Full', 'Section', 'SECTION'],  // Prioritize Section_Full
    'Stit Sup': ['Stit Sup', 'Supervisor', 'supervisor'],
    'Pkg Sup': ['Pkg Sup', 'Packing Supervisor', 'packing supervisor'],
    'Fabric': ['Fabric', 'FABRIC']
  };
  
  const possibleKeys = displayToDataMap[fieldName] || [fieldName];
  
  const values = data
    .map(item => {
      // Try all possible keys
      for (const key of possibleKeys) {
        let value = item[key];
        
        if (!value && item._raw) {
          value = item._raw[key];
        }
        
        if (!value) {
          // Try case variations
          const lowerKey = key.toLowerCase();
          const upperKey = key.toUpperCase();
          
          if (item[lowerKey]) value = item[lowerKey];
          else if (item[upperKey]) value = item[upperKey];
          else if (item._raw && item._raw[lowerKey]) value = item._raw[lowerKey];
          else if (item._raw && item._raw[upperKey]) value = item._raw[upperKey];
        }
        
        if (value && value.toString().trim() !== '') {
          return value;
        }
      }
      
      // Special handling for certain fields
      if (fieldName === 'Party' && item['Party_Full']) {
        return item['Party_Full'];
      }
      if (fieldName === 'Section' && item['Section_Full']) {
        return item['Section_Full'];
      }
      if (fieldName === 'Season' && item['Season_Full']) {
        return item['Season_Full'];
      }
      
      return '';
    })
    .filter(value => value && value.toString().trim() !== '')
    .map(value => value.toString().trim())
    .filter((value, index, self) => self.indexOf(value) === index)
    .sort();
  
  return values;
};

const clearAllFilters = () => {
  setSearchTerm('');
  setBrandFilter([]);
  setGarmentTypeFilter([]);
  setStyleFilter([]);
  setPartyNameFilter([]);
  setSeasonFilter([]);
  setSectionFilter([]);
  setDirectStitchingFilter([]);
  setSupervisorFilter([]);
  setEmbPrintCompFilter([]);
  setCompStatusFilter([]);
  setPkgSupervisorFilter([]);
  setPkgCompFilter([]);
  setHoldLotsFilter(false);
  setFabricFilter([]); // ADDED
  setFabricSearchTerm('');
  setDateFilterField('Job Date');
  setStartDateFilter('');
  setEndDateFilter('');
};

  // ========== EXPORT FUNCTIONS ==========

  // Function to prepare data for export
// Function to prepare data for export
const prepareExportData = (sourceData) => {
  // Filter out cancelled lots
  const filteredData = sourceData.filter(item => !item._isCancelled);
  
  console.log(`Export: Filtered out ${sourceData.length - filteredData.length} cancelled lots`);
  
  return filteredData.map(item => {
    const exportItem = {};
    
    DISPLAY_HEADERS.forEach(header => {
      let value = item[header] || '';
      
      if (header === 'Challan Hist.' && item['CHALLAN HISTORY_parsed']) {
        value = item['CHALLAN HISTORY_parsed'].display;
      } else if (header === 'Fabric') {
        // Ensure Fabric value is properly retrieved
        value = item['Fabric'] || item['Fabric_From_JobOrder'] || '';
      } else if (header === 'WIP Stit') {
        const compStitVal = (item['Comp Stit'] || '').toString().trim();
        if (compStitVal && 
            compStitVal !== '-' && 
            compStitVal.toLowerCase() !== 'invalid date' && 
            !compStitVal.toLowerCase().includes('pending')) {
          value = 'Stitching Done';
        }
      }
      
      exportItem[header] = value;
      
      if (header === 'PCS' && exportItem[header]) {
        exportItem[header] = parseInt(exportItem[header]) || exportItem[header];
      }
    });
    
    return exportItem;
  });
};
  // Export to Excel function with embedded images
  const exportToExcel = async (dataToExport = filteredData) => {
    setExportLoading(true);
    
    try {
      const recordsToExport = dataToExport.filter(item => !item._isCancelled);
      
      if (recordsToExport.length === 0) {
        alert('No data available to export.');
        setExportLoading(false);
        return;
      }

      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Production Data');

      // Define columns
      const columns = DISPLAY_HEADERS.map(header => {
        if (header === 'Image') return { header: 'Image', key: 'Image', width: 16 };
        if (header === 'Sr.') return { header: 'Sr.', key: 'Sr.', width: 8 };
        if (header === 'Lot No') return { header: 'Lot No', key: 'Lot No', width: 14 };
        return { header: header, key: header, width: 18 };
      });
      worksheet.columns = columns;

      // Style header row
      const headerRow = worksheet.getRow(1);
      headerRow.font = { bold: true, color: { argb: 'FFFFFF' }, size: 11 };
      headerRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '4F46E5' }
      };
      headerRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      headerRow.height = 28;

      const imageColIndex = DISPLAY_HEADERS.indexOf('Image');

      // Add rows & process images
      for (let i = 0; i < recordsToExport.length; i++) {
        const item = recordsToExport[i];
        const rowNum = i + 2; // 1-based index + 1 header row
        const rowData = {};

        DISPLAY_HEADERS.forEach((header) => {
          let value = item[header] || '';

          if (header === 'Sr.') {
            value = i + 1;
          } else if (header === 'Challan Hist.' && item['CHALLAN HISTORY_parsed']) {
            value = item['CHALLAN HISTORY_parsed'].display;
          } else if (header === 'Fabric') {
            value = item['Fabric'] || item['Fabric_From_JobOrder'] || '';
          } else if (header === 'PCS' && value) {
            value = parseInt(value) || value;
          }

          if (header === 'Image') {
            const rawUrl = getDirectImageUrl(item['Image'] || item['Image URL']);
            if (rawUrl) {
              rowData['Image'] = { formula: `IMAGE("${rawUrl}")` };
            } else {
              rowData['Image'] = '';
            }
          } else {
            rowData[header] = value;
          }
        });

        const row = worksheet.addRow(rowData);
        row.height = 55; // Row height for thumbnail preview
        row.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

        // Embed Base64 image into Excel cell if image URL exists
        const imgUrl = getDirectImageUrl(item['Image'] || item['Image URL']);
        if (imgUrl && imageColIndex !== -1) {
          try {
            const base64Data = await getBase64ImageFromUrl(imgUrl);
            if (base64Data && base64Data.startsWith('data:image')) {
              const mimeMatch = base64Data.match(/^data:image\/(png|jpeg|jpg|webp);base64,/);
              const extension = mimeMatch ? (mimeMatch[1] === 'jpg' ? 'jpeg' : mimeMatch[1]) : 'jpeg';
              const cleanBase64 = base64Data.replace(/^data:image\/\w+;base64,/, '');

              const imageId = workbook.addImage({
                base64: cleanBase64,
                extension: extension === 'webp' ? 'png' : extension
              });

              worksheet.addImage(imageId, {
                tl: { col: imageColIndex + 0.05, row: rowNum - 1 + 0.05 },
                ext: { width: 50, height: 50 }
              });
            }
          } catch (imgErr) {
            console.warn(`Image embed failed for row ${i + 1}:`, imgErr);
          }
        }
      }

      // Generate Excel file buffer and save
      const buffer = await workbook.xlsx.writeBuffer();
      const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
      const filename = `Production_Tracking_${timestamp}.xlsx`;

      saveAs(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), filename);
      console.log(`Successfully exported ${recordsToExport.length} records with embedded images to Excel`);
    } catch (error) {
      console.error('Error exporting to Excel:', error);
      alert('Error exporting to Excel: ' + error.message);
    } finally {
      setExportLoading(false);
    }
  };

  // Export to PDF function (same as your provided exportToPDF function)
// Export to PDF function (same as your provided exportToPDF function)
// Export to PDF function (same as your provided exportToPDF function)
// Export to PDF function
// Add this new function after your existing exportToPDF function

// Export to PDF WITHOUT DAYS columns
const exportToPDFWithoutDays = () => {
  setExportLoading(true);
  
  try {
    const exportData = prepareExportData(filteredData);
    
    if (exportData.length === 0) {
      alert('No data available to export.');
      setExportLoading(false);
      return;
    }

    // Calculate total PCS
    const totalPCS = exportData.reduce((sum, item) => {
      const pcsValue = item['PCS'];
      const numValue = parseInt(pcsValue);
      return sum + (isNaN(numValue) ? 0 : numValue);
    }, 0);
    
    const totalLots = exportData.length;

    // HEADER MAPPING WITHOUT DAYS COLUMNS
    const HEADER_MAPPING_NO_DAYS = [
      { pdf: 'Sr.', dataKey: 'Sr.', type: 'number', isSerial: true },
      { pdf: 'Lot', dataKey: 'Lot No', type: 'text' },
      { pdf: 'Fabric', dataKey: 'Fabric', type: 'text' },
      { pdf: 'Item', dataKey: 'Item', type: 'text' },
      { pdf: 'Brand', dataKey: 'Brand', type: 'text' },
      { pdf: 'Party', dataKey: 'Party', type: 'text' },
      { pdf: 'Sea', dataKey: 'Season', type: 'text' },
      { pdf: 'M/W/K', dataKey: 'Section', type: 'text' },
      { pdf: 'Design', dataKey: 'Design Work', type: 'stitching' }, 
      { pdf: 'Job Date', dataKey: 'Job Date', type: 'date' },
      { pdf: 'PCS', dataKey: 'PCS', type: 'number' },
      { pdf: 'Cut Date', dataKey: 'Cut Date', type: 'date' },
      { pdf: 'Emb/Print Issue', dataKey: 'Emb/Print Issue', type: 'date' },
      { pdf: 'Emb/print Comp', dataKey: 'Emb/print Comp', type: 'date' },
      { pdf: 'Stit Date', dataKey: 'Stit Date', type: 'date' },
      { pdf: 'Stit Sup', dataKey: 'Stit Sup', type: 'text' },
      { pdf: 'WIP Stit', dataKey: 'WIP Stit', type: 'text' },
      { pdf: 'Comp Stit', dataKey: 'Comp Stit', type: 'date' },
      { pdf: 'Pkg Sup', dataKey: 'Pkg Sup', type: 'text' },
      { pdf: 'Pkg Date', dataKey: 'Pkg Date', type: 'date' },
      { pdf: 'WIP Pkg', dataKey: 'WIP Pkg', type: 'text' },
      { pdf: 'Pkg Comp', dataKey: 'Pkg Comp', type: 'text' }
    ];

    // Column Definitions for Notes Section
    const COLUMN_DEFINITIONS_NO_DAYS = [
      { abbreviation: 'Sr.', fullForm: 'Serial Number' },
      { abbreviation: 'Lot', fullForm: 'Lot Number' },
      { abbreviation: 'Fabric', fullForm: 'Fabric Type' },
      { abbreviation: 'Item', fullForm: 'Item Description' },
      { abbreviation: 'Brand', fullForm: 'Brand Name' },
      { abbreviation: 'Party', fullForm: 'Party/Customer Name' },
      { abbreviation: 'Sea', fullForm: 'Season' },
      { abbreviation: 'M/W/K', fullForm: 'Men/Women/Kids Section' },
      { abbreviation: 'Design', fullForm: 'Design Work (Y=Direct, N=No, Emb=Embroidery, Print=Printing)' },
      { abbreviation: 'Job Date', fullForm: 'Job Order Date' },
      { abbreviation: 'PCS', fullForm: 'Number of Pieces' },
      { abbreviation: 'Cut Date', fullForm: 'Fabric Cutting Date' },
      { abbreviation: 'Emb/Print Issue', fullForm: 'Embroidery/Printing Issue Date' },
      { abbreviation: 'Emb/print Comp', fullForm: 'Embroidery/Printing Completion Date' },
      { abbreviation: 'Stit Date', fullForm: 'Stitching Start Date' },
      { abbreviation: 'Stit Sup', fullForm: 'Stitching Supervisor' },
      { abbreviation: 'WIP Stit', fullForm: 'Work In Progress - Stitching (Shows "Done" if Comp Stit available)' },
      { abbreviation: 'Comp Stit', fullForm: 'Stitching Completion Date' },
      { abbreviation: 'Pkg Sup', fullForm: 'Packing Supervisor' },
      { abbreviation: 'Pkg Date', fullForm: 'Packing Start Date' },
      { abbreviation: 'WIP Pkg', fullForm: 'Work In Progress - Packing' },
      { abbreviation: 'Pkg Comp', fullForm: 'Packing Completion Status' }
    ];

    const PDF_HEADERS = HEADER_MAPPING_NO_DAYS.map(h => h.pdf);

    const monthMap = {
      'jan': 0, 'january': 0, 'feb': 1, 'february': 1, 'mar': 2, 'march': 2,
      'apr': 3, 'april': 3, 'may': 4, 'jun': 5, 'june': 5, 'jul': 6, 'july': 6,
      'aug': 7, 'august': 7, 'sep': 8, 'september': 8, 'oct': 9, 'october': 9,
      'nov': 10, 'november': 10, 'dec': 11, 'december': 11
    };

    const parseDate = (dateString) => {
      if (!dateString || dateString === '' || dateString === '-' || 
          dateString.toLowerCase() === 'null' || dateString.toLowerCase() === 'invalid date') {
        return null;
      }
      
      if (dateString instanceof Date && !isNaN(dateString)) {
        return dateString;
      }
      
      let cleanDate = dateString.toString().trim().replace(/\s+/g, ' ').replace(/\n/g, ' ');
      
      try {
        if (/^\d{1,2}\s+[a-zA-Z]{3,}\s+\d{4}$/.test(cleanDate)) {
          const parts = cleanDate.split(' ');
          const day = parseInt(parts[0], 10);
          const monthName = parts[1].toLowerCase();
          const year = parseInt(parts[2], 10);
          
          if (monthMap[monthName] !== undefined) {
            const date = new Date(year, monthMap[monthName], day);
            if (date.getDate() === day && date.getMonth() === monthMap[monthName] && 
                date.getFullYear() === year) {
              return date;
            }
          }
        }
        
        if (cleanDate.includes('/')) {
          const parts = cleanDate.split('/');
          if (parts.length === 3) {
            const day = parseInt(parts[0], 10);
            const month = parseInt(parts[1], 10) - 1;
            const year = parseInt(parts[2], 10);
            const fullYear = year < 100 ? (year <= 50 ? 2000 + year : 1900 + year) : year;
            
            const date = new Date(fullYear, month, day);
            if (date.getDate() === day && date.getMonth() === month && 
                date.getFullYear() === fullYear) {
              return date;
            }
          }
        }
        
        const parsedDate = new Date(cleanDate);
        if (parsedDate && !isNaN(parsedDate.getTime())) {
          return parsedDate;
        }
        
        return null;
      } catch (error) {
        return null;
      }
    };

    const getDataValue = (item, dataKey, isSerial = false, rowIndex = 0) => {
      if (isSerial) return (rowIndex + 1).toString();
      
      if (dataKey === 'Design Work') {
        let rawValue = '';
        
        if (item[dataKey] !== undefined) {
          rawValue = item[dataKey];
        } else if (item._raw && item._raw[dataKey] !== undefined) {
          rawValue = item._raw[dataKey];
        } else if (item['Design Work'] !== undefined) {
          rawValue = item['Design Work'];
        } else if (item._raw && item._raw['Design Work'] !== undefined) {
          rawValue = item._raw['Design Work'];
        } else if (item._raw && item._raw['DIRECT STITCHING'] !== undefined) {
          rawValue = item._raw['DIRECT STITCHING'];
        } else if (item._raw && item._raw['D.Stitching'] !== undefined) {
          rawValue = item._raw['D.Stitching'];
        }
        
        const emb = item['Emb'] || '';
        const printing = item['Printing'] || '';
        
        const rawValueStr = rawValue ? String(rawValue).trim() : '';
        const embClean = emb ? String(emb).trim() : '';
        const printingClean = printing ? String(printing).trim() : '';
        
        const hasEmbFromJobOrder = embClean !== '' && 
                                  embClean.toUpperCase() !== 'NA' && 
                                  embClean.toUpperCase() !== 'N/A' &&
                                  embClean !== '-';
        
        const hasPrintingFromJobOrder = printingClean !== '' && 
                                       printingClean.toUpperCase() !== 'NA' && 
                                       printingClean.toUpperCase() !== 'N/A' &&
                                       printingClean !== '-';
        
        if (rawValueStr && (rawValueStr.toUpperCase().includes('NO') || rawValueStr.toUpperCase() === 'N')) {
          if (hasEmbFromJobOrder || hasPrintingFromJobOrder) {
            const embLower = hasEmbFromJobOrder ? embClean.toLowerCase() : '';
            const printingLower = hasPrintingFromJobOrder ? printingClean.toLowerCase() : '';
            
            const isJustName = (hasEmbFromJobOrder && !embLower.includes('complete') && 
                               !embLower.includes('done') && !embLower.includes('finished') &&
                               !embLower.includes('pending') && !embLower.includes('process') &&
                               !embLower.includes('ongoing') && !embLower.includes('wip')) ||
                              (hasPrintingFromJobOrder && !printingLower.includes('complete') && 
                               !printingLower.includes('done') && !printingLower.includes('finished') &&
                               !printingLower.includes('pending') && !printingLower.includes('process') &&
                               !printingLower.includes('ongoing') && !printingLower.includes('wip'));
            
            if (isJustName) {
              if (hasEmbFromJobOrder && hasPrintingFromJobOrder) {
                return 'Emb+Prt';
              } else if (hasEmbFromJobOrder) {
                return 'Emb';
              } else if (hasPrintingFromJobOrder) {
                return 'Print';
              }
            } else {
              const isComplete = (embLower.includes('complete') || embLower.includes('done') || 
                                 embLower.includes('finished') || embLower.includes('ready'));
              
              const isPending = (embLower.includes('pending') || embLower.includes('process') || 
                                embLower.includes('ongoing') || embLower.includes('wip'));
              
              if (isComplete) {
                return 'Emb/Print';
              } 
              else if (isPending) {
                return 'Emb/Prt Pending';
              }
              else {
                if (hasEmbFromJobOrder && hasPrintingFromJobOrder) {
                  return 'Emb+Prt';
                } else if (hasEmbFromJobOrder) {
                  return 'Emb';
                } else if (hasPrintingFromJobOrder) {
                  return 'Print';
                }
              }
            }
          }
          return 'N';
        }
        
        if (rawValueStr && (rawValueStr.toUpperCase().includes('YES') || rawValueStr.toUpperCase() === 'Y')) {
          return 'Y';
        }
        
        return rawValueStr || '-';
      }
      
      if ((dataKey === 'Emb/Print Issue' || dataKey === 'Emb/print Comp')) {
        const designWorkValue = getDataValue(item, 'Design Work', false, rowIndex);
        const isDirectStitching = designWorkValue === 'Y';
        const isOverriddenToEmbPrint = designWorkValue === 'Emb/Print' || 
                                       designWorkValue === 'Emb' || 
                                       designWorkValue === 'Print' || 
                                       designWorkValue === 'Emb+Prt' || 
                                       designWorkValue === 'Emb/Prt Pending';
        
        if (isDirectStitching || isOverriddenToEmbPrint) {
          return 'Direct';
        }
      }
      
      const possibleKeys = [
        dataKey, 
        dataKey.toLowerCase(), 
        dataKey.toUpperCase(),
        dataKey.replace(/[\.\s]/g, ''), 
        dataKey.replace(/[\.\s]/g, '').toLowerCase(),
        dataKey.replace(/[\.\s]/g, '').toUpperCase()
      ];
      
      for (const key of possibleKeys) {
        if (item[key] !== undefined && item[key] !== null && item[key] !== '') {
          return item[key];
        }
      }
      
      if (item._raw) {
        for (const key of possibleKeys) {
          if (item._raw[key] !== undefined && item._raw[key] !== null && item._raw[key] !== '') {
            return item._raw[key];
          }
        }
      }
      
      return '';
    };

    const isValidDate = (value) => {
      if (!value || value === '' || value === '-' || value.toLowerCase() === 'null') {
        return false;
      }
      return parseDate(value) !== null;
    };

    const formatDateForPDF = (dateString) => {
      if (!dateString || dateString === '' || dateString === '-' || 
          dateString.toLowerCase() === 'null' || dateString.toLowerCase() === 'invalid date') {
        return '-';
      }
      
      try {
        if (typeof dateString === 'string' && !isNaN(dateString) && dateString !== '-') {
          return dateString;
        }
        
        const date = parseDate(dateString);
        if (date) {
          const day = String(date.getDate()).padStart(2, '0');
          const month = String(date.getMonth() + 1).padStart(2, '0');
          const year = String(date.getFullYear()).slice(-2);
          return `${day}/${month}/${year}`;
        }
        
        const str = dateString.toString().trim();
        if (/^\d{1,2}\s+[a-zA-Z]{3,}\s+\d{4}$/.test(str)) {
          const parts = str.split(' ');
          const day = parts[0];
          const monthName = parts[1].toLowerCase();
          const year = parts[2].slice(-2);
          
          if (monthMap[monthName] !== undefined) {
            const monthNum = monthMap[monthName] + 1;
            return `${day.padStart(2, '0')}/${String(monthNum).padStart(2, '0')}/${year}`;
          }
        }
        
        return str;
      } catch {
        return dateString;
      }
    };

    const formatDirectStitching = (value, item = {}) => {
      if (!value || value === '') return '-';
      
      const str = String(value).toUpperCase();
      
      if (str.includes('YES') || str === 'Y') {
        return 'Direct';
      }
      
      if (str.includes('NO') || str === 'N') {
        const embFromJobOrder = item['Emb'] || item._raw?.['Emb'] || '';
        const printingFromJobOrder = item['Printing'] || item._raw?.['Printing'] || '';
        
        const embClean = embFromJobOrder ? String(embFromJobOrder).trim() : '';
        const printingClean = printingFromJobOrder ? String(printingFromJobOrder).trim() : '';
        
        const hasEmbFromJobOrder = embClean !== '' && 
                                  embClean.toUpperCase() !== 'NA' && 
                                  embClean.toUpperCase() !== 'N/A' &&
                                  embClean !== '-';
        
        const hasPrintingFromJobOrder = printingClean !== '' && 
                                       printingClean.toUpperCase() !== 'NA' && 
                                       printingClean.toUpperCase() !== 'N/A' &&
                                       printingClean !== '-';
        
        if (hasEmbFromJobOrder || hasPrintingFromJobOrder) {
          const embLower = hasEmbFromJobOrder ? embClean.toLowerCase() : '';
          const printingLower = hasPrintingFromJobOrder ? printingClean.toLowerCase() : '';
          
          const isJustName = (hasEmbFromJobOrder && !embLower.includes('complete') && 
                             !embLower.includes('done') && !embLower.includes('finished') &&
                             !embLower.includes('pending') && !embLower.includes('process') &&
                             !embLower.includes('ongoing') && !embLower.includes('wip')) ||
                            (hasPrintingFromJobOrder && !printingLower.includes('complete') && 
                             !printingLower.includes('done') && !printingLower.includes('finished') &&
                             !printingLower.includes('pending') && !printingLower.includes('process') &&
                             !printingLower.includes('ongoing') && !printingLower.includes('wip'));
          
          if (isJustName) {
            if (hasEmbFromJobOrder && hasPrintingFromJobOrder) {
              return 'Emb+Prt';
            } else if (hasEmbFromJobOrder) {
              return 'Emb';
            } else if (hasPrintingFromJobOrder) {
              return 'Print';
            }
          } else {
            const isComplete = (embLower.includes('complete') || embLower.includes('done') || 
                               embLower.includes('finished') || embLower.includes('ready'));
            
            const isPending = (embLower.includes('pending') || embLower.includes('process') || 
                              embLower.includes('ongoing') || embLower.includes('wip'));
            
            if (isComplete) {
              return 'Emb/Print';
            } 
            else if (isPending) {
              return 'Emb/Prt Pending';
            }
            else {
              if (hasEmbFromJobOrder && hasPrintingFromJobOrder) {
                return 'Emb+Prt';
              } else if (hasEmbFromJobOrder) {
                return 'Emb';
              } else if (hasPrintingFromJobOrder) {
                return 'Print';
              }
            }
          }
        }
        return 'N';
      }
      
      return str.charAt(0);
    };

    const doc = new jsPDF({
      orientation: 'landscape',
      unit: 'mm',
      format: 'a3',
      compress: true
    });

    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const margin = 8;
    const availableWidth = pageWidth - (margin * 2);
    
    // Define column widths in mm that will fit within available width
    const columnWidthsMM = {
      'Sr.': 10,
      'Lot': 22,
      'Fabric': 25,
      'Item': 28,
      'Brand': 22,
      'Party': 22,
      'Sea': 12,
      'M/W/K': 14,
      'Design': 20,
      'Job Date': 18,
      'PCS': 15,
      'Cut Date': 18,
      'Emb/Print Issue': 22,
      'Emb/print Comp': 22,
      'Stit Date': 18,
      'Stit Sup': 18,
      'WIP Stit': 20,
      'Comp Stit': 18,
      'Pkg Sup': 18,
      'Pkg Date': 18,
      'WIP Pkg': 20,
      'Pkg Comp': 18
    };
    
    // Verify total width doesn't exceed available width
    let totalWidth = 0;
    const columnOrder = ['Sr.', 'Lot', 'Fabric', 'Item', 'Brand', 'Party', 'Sea', 'M/W/K', 'Design', 
                         'Job Date', 'PCS', 'Cut Date', 'Emb/Print Issue', 'Emb/print Comp', 
                         'Stit Date', 'Stit Sup', 'WIP Stit', 'Comp Stit', 'Pkg Sup', 
                         'Pkg Date', 'WIP Pkg', 'Pkg Comp'];
    
    columnOrder.forEach(col => {
      totalWidth += columnWidthsMM[col];
    });
    
    // If total exceeds available, scale down proportionally
    let finalWidths = {};
    if (totalWidth > availableWidth) {
      const scaleFactor = availableWidth / totalWidth;
      columnOrder.forEach(col => {
        finalWidths[col] = columnWidthsMM[col] * scaleFactor;
      });
    } else {
      // If less than available, distribute extra space
      const extraSpace = availableWidth - totalWidth;
      const extraPerColumn = extraSpace / columnOrder.length;
      columnOrder.forEach(col => {
        finalWidths[col] = columnWidthsMM[col] + extraPerColumn;
      });
    }
    
    // Create array of widths in the correct order for autoTable
    const widthsArray = columnOrder.map(col => finalWidths[col]);

    const addPageHeader = () => {
      doc.setFontSize(14);
      doc.setFont('times', 'bold');
      doc.setTextColor(15, 76, 129);
      doc.text('PRODUCTION TRACKING REPORT', pageWidth / 2, 15, { align: 'center' });
      
      doc.setFontSize(10);
      doc.setFont('times', 'bold');
      doc.setTextColor(0, 100, 0);
      doc.text(`Total Lots: ${totalLots} | Total PCS: ${totalPCS.toLocaleString()}`, pageWidth / 2, 23, { align: 'center' });
      
      doc.setFontSize(8);
      doc.setFont('times', 'normal');
      doc.setTextColor(100, 100, 100);
      const today = new Date();
      const reportDate = `${String(today.getDate()).padStart(2, '0')}/${String(today.getMonth() + 1).padStart(2, '0')}/${String(today.getFullYear()).slice(-2)}`;
      doc.text(`Report Date: ${reportDate} | Records: ${exportData.length}`, pageWidth / 2, 30, { align: 'center' });
      
      doc.setDrawColor(200, 200, 200);
      doc.line(margin, 34, pageWidth - margin, 34);
    };

    const addColumnDefinitions = (startY) => {
      doc.setFontSize(10);
      doc.setFont('times', 'bold');
      doc.setTextColor(15, 76, 129);
      doc.text('Column Definitions:', margin, startY);
      
      doc.setFontSize(7);
      doc.setFont('times', 'normal');
      doc.setTextColor(0, 0, 0);
      
      let currentY = startY + 5;
      const lineHeight = 4;
      
      const itemsPerCol = Math.ceil(COLUMN_DEFINITIONS_NO_DAYS.length / 3);
      const col1 = COLUMN_DEFINITIONS_NO_DAYS.slice(0, itemsPerCol);
      const col2 = COLUMN_DEFINITIONS_NO_DAYS.slice(itemsPerCol, itemsPerCol * 2);
      const col3 = COLUMN_DEFINITIONS_NO_DAYS.slice(itemsPerCol * 2);
      
      const col1X = margin;
      const col2X = margin + 130;
      const col3X = margin + 260;
      
      col1.forEach((definition) => {
        doc.text(`${definition.abbreviation}: ${definition.fullForm}`, col1X, currentY);
        currentY += lineHeight;
      });
      
      currentY = startY + 5;
      col2.forEach((definition) => {
        doc.text(`${definition.abbreviation}: ${definition.fullForm}`, col2X, currentY);
        currentY += lineHeight;
      });
      
      currentY = startY + 5;
      col3.forEach((definition) => {
        doc.text(`${definition.abbreviation}: ${definition.fullForm}`, col3X, currentY);
        currentY += lineHeight;
      });
    };

    addPageHeader();

    const tableBody = exportData.map((item, rowIndex) => {
      // Get Comp Stit value for this row to check if stitching is completed
      const compStitValue = getDataValue(item, 'Comp Stit', false, rowIndex);
      const hasCompStit = compStitValue && compStitValue !== '' && compStitValue !== '-';
      
      return HEADER_MAPPING_NO_DAYS.map((headerInfo, colIndex) => {
        const rawValue = getDataValue(item, headerInfo.dataKey, headerInfo.isSerial, rowIndex);
        let displayValue = '';
        
        // Handle Emb/Print Issue and Emb/print Comp columns
        if (colIndex === 12 || colIndex === 13) {
          const designWorkValue = getDataValue(item, 'Design Work');
          const miniItem = {
            'Emb': item['Emb'] || (item._raw && item._raw['Emb']) || '',
            'Printing': item['Printing'] || (item._raw && item._raw['Printing']) || '',
            _raw: {
              'Emb': item['Emb'] || (item._raw && item._raw['Emb']) || '',
              'Printing': item['Printing'] || (item._raw && item._raw['Printing']) || ''
            }
          };
          const formattedDesignWork = formatDirectStitching(designWorkValue, miniItem);
          const isDirectStitching = formattedDesignWork === 'Y';
          
          if (isDirectStitching) {
            displayValue = 'Direct';
          } else {
            displayValue = formatDateForPDF(rawValue);
          }
        }
        // Handle WIP Stit column - Show "Done" if Comp Stit is available
        else if (colIndex === 16) {
          if (hasCompStit) {
            displayValue = 'Done';
          } else {
            displayValue = rawValue || '-';
          }
        }
        // Handle Comp Stit column - Show date
        else if (colIndex === 17) {
          displayValue = formatDateForPDF(rawValue);
        }
        // Handle all other columns
        else {
          switch(headerInfo.type) {
            case 'date': 
              displayValue = formatDateForPDF(rawValue); 
              break;
            case 'stitching': 
              const miniItemForDesignWork = {
                'Emb': item['Emb'] || (item._raw && item._raw['Emb']) || '',
                'Printing': item['Printing'] || (item._raw && item._raw['Printing']) || '',
                _raw: {
                  'Emb': item['Emb'] || (item._raw && item._raw['Emb']) || '',
                  'Printing': item['Printing'] || (item._raw && item._raw['Printing']) || ''
                }
              };
              displayValue = formatDirectStitching(rawValue, miniItemForDesignWork); 
              break;
            case 'number':
              if (headerInfo.isSerial) displayValue = rawValue;
              else if (rawValue && !isNaN(rawValue)) displayValue = parseInt(rawValue).toLocaleString();
              else displayValue = rawValue || '-';
              break;
            case 'text': displayValue = rawValue || '-'; break;
            default: displayValue = rawValue || '-';
          }
        }
        
        return displayValue;
      });
    });

   autoTable(doc, {
  head: [PDF_HEADERS],
  body: tableBody,
  startY: 40,
  theme: 'grid',
  headStyles: {
    fillColor: [15, 76, 129],
    textColor: [255, 255, 255],
    fontStyle: 'bold',
    fontSize: 8,
    cellPadding: { top: 2, right: 1, bottom: 2, left: 1 },
    halign: 'center',
    valign: 'middle',
    lineWidth: 0.5,
    lineColor: [0, 0, 0]
  },
  bodyStyles: {
    fontSize: 9,
    cellPadding: { top: 1, right: 1, bottom: 1, left: 1 },
    lineWidth: 0.5,
    lineColor: [0, 0, 0],
    textColor: [0, 0, 0],
    font: 'times',
    fontStyle: 'normal',
    valign: 'middle',
    overflow: 'linebreak',
    minCellHeight: 5,
    lineHeight: 1.1
  },
  styles: {
    fontSize: 9,
    cellPadding: 1.0,
    lineWidth: 0.5,
    lineColor: [0, 0, 0],
    font: 'times',
    fontStyle: 'normal',
    valign: 'middle',
    overflow: 'linebreak'
  },
  columnStyles: {
    0: { cellWidth: widthsArray[0], halign: 'center' },
    1: { cellWidth: widthsArray[1], halign: 'center' },
    2: { cellWidth: widthsArray[2], halign: 'center' },
    3: { cellWidth: widthsArray[3], halign: 'left' },
    4: { cellWidth: widthsArray[4], halign: 'center' },
    5: { cellWidth: widthsArray[5], halign: 'left' },
    6: { cellWidth: widthsArray[6], halign: 'center' },
    7: { cellWidth: widthsArray[7], halign: 'center' },
    8: { cellWidth: widthsArray[8], halign: 'center' },
    9: { cellWidth: widthsArray[9], halign: 'center' },
    10: { cellWidth: widthsArray[10], halign: 'center' },
    11: { cellWidth: widthsArray[11], halign: 'center' },
    12: { cellWidth: widthsArray[12], halign: 'center' },
    13: { cellWidth: widthsArray[13], halign: 'center' },
    14: { cellWidth: widthsArray[14], halign: 'center' },
    15: { cellWidth: widthsArray[15], halign: 'center' },
    16: { cellWidth: widthsArray[16], halign: 'center' },
    17: { cellWidth: widthsArray[17], halign: 'center' },
    18: { cellWidth: widthsArray[18], halign: 'center' },
    19: { cellWidth: widthsArray[19], halign: 'center' },
    20: { cellWidth: widthsArray[20], halign: 'left' },
    21: { cellWidth: widthsArray[21], halign: 'center' }
  },
  margin: { top: 40, left: margin, right: margin, bottom: 30 },
  tableWidth: availableWidth,
  showHead: 'firstPage',
  showFoot: false,
  pageBreak: 'auto',
  rowPageBreak: 'auto',
  didParseCell: function(data) {
    if (data.section === 'body') {
      if ((data.column.index === 12 || data.column.index === 13) && 
          data.cell.raw === 'Direct') {
        data.cell.styles.textColor = [0, 0, 255];
        data.cell.styles.fontStyle = 'bolditalic';
      }
      if (data.column.index === 16 && data.cell.raw === 'Done') {
        data.cell.styles.textColor = [0, 128, 0];
        data.cell.styles.fontStyle = 'bold';
      }
    }
  },
  didDrawPage: function(data) {
    const summaryY = pageHeight - 18;
    
    doc.setFontSize(8);
    doc.setFont('times', 'bold');
    doc.setTextColor(15, 76, 129);
    doc.text(`Total: ${totalLots} Lots | ${totalPCS.toLocaleString()} PCS`, pageWidth / 2, summaryY, { align: 'center' });
    
    doc.setFontSize(7);
    doc.setTextColor(100, 100, 100);
    doc.text(
      `Page ${data.pageNumber} of ${data.pageCount}`,
      pageWidth - margin - 5,
      pageHeight - 10
    );
    
    if (data.pageNumber === data.pageCount) {
      const finalY = data.cursor ? data.cursor.y : pageHeight - 50;
      const availableSpace = pageHeight - finalY - 15;
      
      if (availableSpace > 40) {
        addColumnDefinitions(finalY + 5);
      }
    }
  }
});

    const today = new Date();
    const fileName = `Production_Tracking_Report_No_Days_${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, '0')}${String(today.getDate()).padStart(2, '0')}.pdf`;
    
    doc.save(fileName);
    
    setTimeout(() => {
      alert(`✅ PDF Exported Successfully! (Without Days Calculation)\n📄 ${fileName}\n📊 ${totalLots} lots | ${totalPCS.toLocaleString()} PCS`);
    }, 500);
    
  } catch (error) {
    console.error('PDF Export Error Details:', error);
    alert(`❌ Failed to export PDF: ${error.message}`);
  } finally {
    setExportLoading(false);
  }
};
const exportToPDF = () => {
  setExportLoading(true);
  
  try {
    const exportData = prepareExportData(filteredData);
    
    if (exportData.length === 0) {
      alert('No data available to export.');
      setExportLoading(false);
      return;
    }

    // Calculate total PCS
    const totalPCS = exportData.reduce((sum, item) => {
      const pcsValue = item['PCS'];
      const numValue = parseInt(pcsValue);
      return sum + (isNaN(numValue) ? 0 : numValue);
    }, 0);
    
    const totalLots = exportData.length;

    const HEADER_MAPPING = [
      { pdf: 'Sr.', dataKey: 'Sr.', type: 'number', isSerial: true },
      { pdf: 'Lot', dataKey: 'Lot No', type: 'text' },
      { pdf: 'Item', dataKey: 'Item', type: 'text' },
      { pdf: 'Brand', dataKey: 'Brand', type: 'text' },
      { pdf: 'Party', dataKey: 'Party', type: 'text' },
      { pdf: 'Sea', dataKey: 'Season', type: 'text' },
      { pdf: 'M/W/K', dataKey: 'Section', type: 'text' },
      { pdf: 'Design', dataKey: 'Design Work', type: 'stitching' }, 
      { pdf: 'Job Date', dataKey: 'Job Date', type: 'date' },
      { pdf: 'PCS', dataKey: 'PCS', type: 'number' },
      { pdf: 'Cut Date', dataKey: 'Cut Date', type: 'date' },
      { pdf: 'Emb/Print Issue', dataKey: 'Emb/Print Issue', type: 'date' },
      { pdf: 'Emb/print Comp', dataKey: 'Emb/print Comp', type: 'date' },
      { pdf: 'Stit Date', dataKey: 'Stit Date', type: 'date' },
      { pdf: 'Stit Sup', dataKey: 'Stit Sup', type: 'text' },
      { pdf: 'WIP Stit', dataKey: 'WIP Stit', type: 'text' },
      { pdf: 'Comp Stit', dataKey: 'Comp Stit', type: 'date' },
      { pdf: 'Pkg Sup', dataKey: 'Pkg Sup', type: 'text' },
      { pdf: 'Pkg Date', dataKey: 'Pkg Date', type: 'date' },
      { pdf: 'WIP Pkg', dataKey: 'WIP Pkg', type: 'text' },
      { pdf: 'Pkg Comp', dataKey: 'Pkg Comp', type: 'text' },
      // DAYS COLUMNS - ALL MOVED TO THE END
      { pdf: 'Cut Days', dataKey: 'Cut Days', type: 'calculated', formula: 'Cut Date - Job Date' },
      { pdf: 'Emb/Print Days', dataKey: 'Emb/Print Days', type: 'calculated', formula: 'Emb/print Comp - Emb/Print Issue' },
      { pdf: 'Stit Days', dataKey: 'Stit Days', type: 'calculated', formula: 'Comp Stit - Stit Date' },
      { pdf: 'Pkg Days', dataKey: 'Pkg Days', type: 'calculated', formula: 'Pkg Comp - Pkg Date' },
      { pdf: 'Cut To Emb/Print', dataKey: 'Cut To Emb/Print', type: 'calculated', formula: 'Emb/Print Issue - Cut Date' },
      { pdf: 'Emb/Print To Stit', dataKey: 'Emb/Print To Stit', type: 'calculated', formula: 'Stit Date - Emb/print Comp' },
      { pdf: 'Stit To Check Pack', dataKey: 'Stit To Check Pack', type: 'calculated', formula: 'Pkg Date - Comp Stit' }
    ];

    // Column Definitions for Notes Section
    const COLUMN_DEFINITIONS = [
      { abbreviation: 'Sr.', fullForm: 'Serial Number' },
      { abbreviation: 'Lot', fullForm: 'Lot Number' },
      { abbreviation: 'Item', fullForm: 'Item Description' },
      { abbreviation: 'Brand', fullForm: 'Brand Name' },
      { abbreviation: 'Party', fullForm: 'Party/Customer Name' },
      { abbreviation: 'Sea', fullForm: 'Season' },
      { abbreviation: 'M/W/K', fullForm: 'Men/Women/Kids Section' },
      { abbreviation: 'Design', fullForm: 'Design Work (Y=Direct, N=No, Emb=Embroidery, Print=Printing)' },
      { abbreviation: 'Job Date', fullForm: 'Job Order Date' },
      { abbreviation: 'PCS', fullForm: 'Number of Pieces' },
      { abbreviation: 'Cut Date', fullForm: 'Fabric Cutting Date' },
      { abbreviation: 'Emb/Print Issue', fullForm: 'Embroidery/Printing Issue Date' },
      { abbreviation: 'Emb/print Comp', fullForm: 'Embroidery/Printing Completion Date' },
      { abbreviation: 'Stit Date', fullForm: 'Stitching Start Date' },
      { abbreviation: 'Stit Sup', fullForm: 'Stitching Supervisor' },
      { abbreviation: 'WIP Stit', fullForm: 'Work In Progress - Stitching' },
      { abbreviation: 'Comp Stit', fullForm: 'Stitching Completion Date' },
      { abbreviation: 'Pkg Sup', fullForm: 'Packing Supervisor' },
      { abbreviation: 'Pkg Date', fullForm: 'Packing Start Date' },
      { abbreviation: 'WIP Pkg', fullForm: 'Work In Progress - Packing' },
      { abbreviation: 'Pkg Comp', fullForm: 'Packing Completion Status' },
      { abbreviation: 'Cut Days', fullForm: 'Days from Job Date to Cut Date' },
      { abbreviation: 'Emb/Print Days', fullForm: 'Days for Embroidery/Printing Process' },
      { abbreviation: 'Stit Days', fullForm: 'Days for Stitching Process' },
      { abbreviation: 'Pkg Days', fullForm: 'Days for Packing Process' },
      { abbreviation: 'Cut To Emb/Print', fullForm: 'Days from Cutting to Embroidery/Printing Issue' },
      { abbreviation: 'Emb/Print To Stit', fullForm: 'Days from Embroidery/Printing to Stitching' },
      { abbreviation: 'Stit To Check Pack', fullForm: 'Days from Stitching to Packing' }
    ];

    const PDF_HEADERS = HEADER_MAPPING.map(h => h.pdf);

    const monthMap = {
      'jan': 0, 'january': 0, 'feb': 1, 'february': 1, 'mar': 2, 'march': 2,
      'apr': 3, 'april': 3, 'may': 4, 'jun': 5, 'june': 5, 'jul': 6, 'july': 6,
      'aug': 7, 'august': 7, 'sep': 8, 'september': 8, 'oct': 9, 'october': 9,
      'nov': 10, 'november': 10, 'dec': 11, 'december': 11
    };

    const parseDate = (dateString) => {
      if (!dateString || dateString === '' || dateString === '-' || 
          dateString.toLowerCase() === 'null' || dateString.toLowerCase() === 'invalid date') {
        return null;
      }
      
      if (dateString instanceof Date && !isNaN(dateString)) {
        return dateString;
      }
      
      let cleanDate = dateString.toString().trim().replace(/\s+/g, ' ').replace(/\n/g, ' ');
      
      try {
        if (/^\d{1,2}\s+[a-zA-Z]{3,}\s+\d{4}$/.test(cleanDate)) {
          const parts = cleanDate.split(' ');
          const day = parseInt(parts[0], 10);
          const monthName = parts[1].toLowerCase();
          const year = parseInt(parts[2], 10);
          
          if (monthMap[monthName] !== undefined) {
            const date = new Date(year, monthMap[monthName], day);
            if (date.getDate() === day && date.getMonth() === monthMap[monthName] && 
                date.getFullYear() === year) {
              return date;
            }
          }
        }
        
        if (cleanDate.includes('/')) {
          const parts = cleanDate.split('/');
          if (parts.length === 3) {
            const day = parseInt(parts[0], 10);
            const month = parseInt(parts[1], 10) - 1;
            const year = parseInt(parts[2], 10);
            const fullYear = year < 100 ? (year <= 50 ? 2000 + year : 1900 + year) : year;
            
            const date = new Date(fullYear, month, day);
            if (date.getDate() === day && date.getMonth() === month && 
                date.getFullYear() === fullYear) {
              return date;
            }
          }
        }
        
        const parsedDate = new Date(cleanDate);
        if (parsedDate && !isNaN(parsedDate.getTime())) {
          return parsedDate;
        }
        
        return null;
      } catch (error) {
        return null;
      }
    };

    const calculateDaysDifference = (date1Str, date2Str) => {
      const date1 = parseDate(date1Str);
      const date2 = parseDate(date2Str);
      
      if (!date1 || !date2) return '-';
      if (!(date1 instanceof Date) || isNaN(date1.getTime()) ||
          !(date2 instanceof Date) || isNaN(date2.getTime())) {
        return '-';
      }
      
      const diffTime = date2.getTime() - date1.getTime();
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      
      if (diffDays < 0) {
        return diffDays.toString();
      } else if (diffDays === 0) {
        return '0';
      } else {
        return `+${diffDays}`;
      }
    };

    const getDataValue = (item, dataKey, isSerial = false, rowIndex = 0) => {
      if (isSerial) return (rowIndex + 1).toString();
      
      // Handle Design Work specially to check JobOrder Emb/Printing columns
      if (dataKey === 'Design Work') {
        // Get the raw value from the item
        let rawValue = '';
        
        // Try different ways to get the Design Work value
        if (item[dataKey] !== undefined) {
          rawValue = item[dataKey];
        } else if (item._raw && item._raw[dataKey] !== undefined) {
          rawValue = item._raw[dataKey];
        } else if (item['Design Work'] !== undefined) {
          rawValue = item['Design Work'];
        } else if (item._raw && item._raw['Design Work'] !== undefined) {
          rawValue = item._raw['Design Work'];
        } else if (item._raw && item._raw['DIRECT STITCHING'] !== undefined) {
          rawValue = item._raw['DIRECT STITCHING'];
        } else if (item._raw && item._raw['D.Stitching'] !== undefined) {
          rawValue = item._raw['D.Stitching'];
        }
        
        // Get Emb and Printing values from the original item
        const emb = item['Emb'] || '';
        const printing = item['Printing'] || '';
        
        // Clean the values
        const rawValueStr = rawValue ? String(rawValue).trim() : '';
        const embClean = emb ? String(emb).trim() : '';
        const printingClean = printing ? String(printing).trim() : '';
        
        // Check if there's valid data in JobOrder columns
        const hasEmbFromJobOrder = embClean !== '' && 
                                  embClean.toUpperCase() !== 'NA' && 
                                  embClean.toUpperCase() !== 'N/A' &&
                                  embClean !== '-';
        
        const hasPrintingFromJobOrder = printingClean !== '' && 
                                       printingClean.toUpperCase() !== 'NA' && 
                                       printingClean.toUpperCase() !== 'N/A' &&
                                       printingClean !== '-';
        
        // If Design Work is "NO" and has Emb/Printing data in JobOrder, format accordingly
        if (rawValueStr && (rawValueStr.toUpperCase().includes('NO') || rawValueStr.toUpperCase() === 'N')) {
          
          if (hasEmbFromJobOrder || hasPrintingFromJobOrder) {
            
            const embLower = hasEmbFromJobOrder ? embClean.toLowerCase() : '';
            const printingLower = hasPrintingFromJobOrder ? printingClean.toLowerCase() : '';
            
            // Check if it's just a name/initials vs status text
            const isJustName = (hasEmbFromJobOrder && !embLower.includes('complete') && 
                               !embLower.includes('done') && !embLower.includes('finished') &&
                               !embLower.includes('pending') && !embLower.includes('process') &&
                               !embLower.includes('ongoing') && !embLower.includes('wip')) ||
                              (hasPrintingFromJobOrder && !printingLower.includes('complete') && 
                               !printingLower.includes('done') && !printingLower.includes('finished') &&
                               !printingLower.includes('pending') && !printingLower.includes('process') &&
                               !printingLower.includes('ongoing') && !printingLower.includes('wip'));
            
            if (isJustName) {
              // For names/initials like GOURAV, BHM
              if (hasEmbFromJobOrder && hasPrintingFromJobOrder) {
                return 'Emb+Prt';
              } else if (hasEmbFromJobOrder) {
                return 'Emb';
              } else if (hasPrintingFromJobOrder) {
                return 'Print';
              }
            } else {
              // Check status words
              const isComplete = (embLower.includes('complete') || embLower.includes('done') || 
                                 embLower.includes('finished') || embLower.includes('ready'));
              
              const isPending = (embLower.includes('pending') || embLower.includes('process') || 
                                embLower.includes('ongoing') || embLower.includes('wip'));
              
              if (isComplete) {
                return 'Emb/Print';
              } 
              else if (isPending) {
                return 'Emb/Prt Pending';
              }
              else {
                // Unknown status
                if (hasEmbFromJobOrder && hasPrintingFromJobOrder) {
                  return 'Emb+Prt';
                } else if (hasEmbFromJobOrder) {
                  return 'Emb';
                } else if (hasPrintingFromJobOrder) {
                  return 'Print';
                }
              }
            }
          }
          
          // Return 'N' if no Emb/Printing data in JobOrder
          return 'N';
        }
        
        // For YES or other values
        if (rawValueStr && (rawValueStr.toUpperCase().includes('YES') || rawValueStr.toUpperCase() === 'Y')) {
          return 'Y';
        }
        
        return rawValueStr || '-';
      }
      
      // For Emb/Print Issue and Emb/print Comp - handle direct stitching override
      if ((dataKey === 'Emb/Print Issue' || dataKey === 'Emb/print Comp')) {
        // Check if this is direct stitching
        const designWorkValue = getDataValue(item, 'Design Work', false, rowIndex);
        
        // Check if Design Work is "YES" or if it was "NO" but overridden to Emb/Print
        const isDirectStitching = designWorkValue === 'Y';
        const isOverriddenToEmbPrint = designWorkValue === 'Emb/Print' || 
                                       designWorkValue === 'Emb' || 
                                       designWorkValue === 'Print' || 
                                       designWorkValue === 'Emb+Prt' || 
                                       designWorkValue === 'Emb/Prt Pending';
        
        if (isDirectStitching || isOverriddenToEmbPrint) {
          return 'Direct';
        }
      }
      
      // For Emb and Printing columns themselves
      if (dataKey === 'Emb' || dataKey === 'Printing') {
        // Try to get the value directly
        if (item[dataKey] !== undefined && item[dataKey] !== null && item[dataKey] !== '') {
          return item[dataKey];
        }
        
        if (item._raw && item._raw[dataKey] !== undefined && item._raw[dataKey] !== null && item._raw[dataKey] !== '') {
          return item._raw[dataKey];
        }
        
        return '';
      }
      
      // Handle days calculation columns
      if (dataKey.includes('Days') || dataKey.includes('To') || dataKey.includes('Cut To')) {
        const headerInfo = HEADER_MAPPING.find(h => h.dataKey === dataKey);
        if (headerInfo && headerInfo.formula) {
          try {
            if (headerInfo.formula === 'Cut Date - Job Date') {
              const cutDate = getDataValue(item, 'Cut Date');
              const jobDate = getDataValue(item, 'Job Date');
              return calculateDaysDifference(jobDate, cutDate);
            }
            else if (headerInfo.formula === 'Emb/print Comp - Emb/Print Issue') {
              const embComp = getDataValue(item, 'Emb/print Comp');
              const embIssue = getDataValue(item, 'Emb/Print Issue');
              return calculateDaysDifference(embIssue, embComp);
            }
            else if (headerInfo.formula === 'Comp Stit - Stit Date') {
              const compStit = getDataValue(item, 'Comp Stit');
              const stitDate = getDataValue(item, 'Stit Date');
              return calculateDaysDifference(stitDate, compStit);
            }
            else if (headerInfo.formula === 'Pkg Comp - Pkg Date') {
              const pkgComp = getDataValue(item, 'Pkg Comp');
              const pkgDate = getDataValue(item, 'Pkg Date');
              return calculateDaysDifference(pkgDate, pkgComp);
            }
            else if (headerInfo.formula === 'Emb/Print Issue - Cut Date') {
              const embPrintIssue = getDataValue(item, 'Emb/Print Issue');
              const cutDate = getDataValue(item, 'Cut Date');
              return calculateDaysDifference(cutDate, embPrintIssue);
            }
            else if (headerInfo.formula === 'Stit Date - Emb/print Comp') {
              const stitDate = getDataValue(item, 'Stit Date');
              const embPrintComp = getDataValue(item, 'Emb/print Comp');
              return calculateDaysDifference(embPrintComp, stitDate);
            }
            else if (headerInfo.formula === 'Pkg Date - Comp Stit') {
              const pkgDate = getDataValue(item, 'Pkg Date');
              const compStit = getDataValue(item, 'Comp Stit');
              return calculateDaysDifference(compStit, pkgDate);
            }
          } catch (error) {
            return '-';
          }
        }
      }
      
      // For all other columns, use the original logic
      const possibleKeys = [
        dataKey, 
        dataKey.toLowerCase(), 
        dataKey.toUpperCase(),
        dataKey.replace(/[\.\s]/g, ''), 
        dataKey.replace(/[\.\s]/g, '').toLowerCase(),
        dataKey.replace(/[\.\s]/g, '').toUpperCase()
      ];
      
      // Try to get value from item
      for (const key of possibleKeys) {
        if (item[key] !== undefined && item[key] !== null && item[key] !== '') {
          return item[key];
        }
      }
      
      // Try to get value from _raw
      if (item._raw) {
        for (const key of possibleKeys) {
          if (item._raw[key] !== undefined && item._raw[key] !== null && item._raw[key] !== '') {
            return item._raw[key];
          }
        }
      }
      
      // Return empty string if nothing found
      return '';
    };

    const isValidDate = (value) => {
      if (!value || value === '' || value === '-' || value.toLowerCase() === 'null') {
        return false;
      }
      return parseDate(value) !== null;
    };

    const formatDateForPDF = (dateString) => {
      if (!dateString || dateString === '' || dateString === '-' || 
          dateString.toLowerCase() === 'null' || dateString.toLowerCase() === 'invalid date') {
        return '-';
      }
      
      try {
        if (typeof dateString === 'string' && !isNaN(dateString) && dateString !== '-') {
          return dateString;
        }
        
        const date = parseDate(dateString);
        if (date) {
          const day = String(date.getDate()).padStart(2, '0');
          const month = String(date.getMonth() + 1).padStart(2, '0');
          const year = String(date.getFullYear()).slice(-2);
          return `${day}/${month}/${year}`;
        }
        
        const str = dateString.toString().trim();
        if (/^\d{1,2}\s+[a-zA-Z]{3,}\s+\d{4}$/.test(str)) {
          const parts = str.split(' ');
          const day = parts[0];
          const monthName = parts[1].toLowerCase();
          const year = parts[2].slice(-2);
          
          if (monthMap[monthName] !== undefined) {
            const monthNum = monthMap[monthName] + 1;
            return `${day.padStart(2, '0')}/${String(monthNum).padStart(2, '0')}/${year}`;
          }
        }
        
        return str;
      } catch {
        return dateString;
      }
    };

    const formatDirectStitching = (value, item = {}) => {
      if (!value || value === '') return '-';
      
      const str = String(value).toUpperCase();
      
      // If it's "YES", return "Direct"
      if (str.includes('YES') || str === 'Y') {
        return 'Direct';
      }
      
      // If it's "NO", check ONLY JobOrder sheet columns
      if (str.includes('NO') || str === 'N') {
        // Check ONLY JobOrder sheet columns
        const embFromJobOrder = item['Emb'] || item._raw?.['Emb'] || '';
        const printingFromJobOrder = item['Printing'] || item._raw?.['Printing'] || '';
        
        // Clean JobOrder values
        const embClean = embFromJobOrder ? String(embFromJobOrder).trim() : '';
        const printingClean = printingFromJobOrder ? String(printingFromJobOrder).trim() : '';
        
        // Check if there's valid data in JobOrder columns (not "NA")
        const hasEmbFromJobOrder = embClean !== '' && 
                                  embClean.toUpperCase() !== 'NA' && 
                                  embClean.toUpperCase() !== 'N/A' &&
                                  embClean !== '-';
        
        const hasPrintingFromJobOrder = printingClean !== '' && 
                                       printingClean.toUpperCase() !== 'NA' && 
                                       printingClean.toUpperCase() !== 'N/A' &&
                                       printingClean !== '-';
        
        // If there's data in JobOrder Emb/Printing columns, show appropriate text
        if (hasEmbFromJobOrder || hasPrintingFromJobOrder) {
          
          const embLower = hasEmbFromJobOrder ? embClean.toLowerCase() : '';
          const printingLower = hasPrintingFromJobOrder ? printingClean.toLowerCase() : '';
          
          // Check if it's just a name/initials (like "GOURAV", "BHM") vs status text
          const isJustName = (hasEmbFromJobOrder && !embLower.includes('complete') && 
                             !embLower.includes('done') && !embLower.includes('finished') &&
                             !embLower.includes('pending') && !embLower.includes('process') &&
                             !embLower.includes('ongoing') && !embLower.includes('wip')) ||
                            (hasPrintingFromJobOrder && !printingLower.includes('complete') && 
                             !printingLower.includes('done') && !printingLower.includes('finished') &&
                             !printingLower.includes('pending') && !printingLower.includes('process') &&
                             !printingLower.includes('ongoing') && !printingLower.includes('wip'));
          
          if (isJustName) {
            // If it's just a name/initials (like GOURAV, BHM)
            if (hasEmbFromJobOrder && hasPrintingFromJobOrder) {
              return 'Emb+Prt';
            } else if (hasEmbFromJobOrder) {
              return 'Emb';
            } else if (hasPrintingFromJobOrder) {
              return 'Print';
            }
          } else {
            // If it contains status words, check status
            const isComplete = (embLower.includes('complete') || embLower.includes('done') || 
                               embLower.includes('finished') || embLower.includes('ready'));
            
            const isPending = (embLower.includes('pending') || embLower.includes('process') || 
                              embLower.includes('ongoing') || embLower.includes('wip'));
            
            if (isComplete) {
              return 'Emb/Print';
            } 
            else if (isPending) {
              return 'Emb/Prt Pending';
            }
            else {
              // Unknown status but has data
              if (hasEmbFromJobOrder && hasPrintingFromJobOrder) {
                return 'Emb+Prt';
              } else if (hasEmbFromJobOrder) {
                return 'Emb';
              } else if (hasPrintingFromJobOrder) {
                return 'Print';
              }
            }
          }
        }
        
        return 'N';  // Return N if no data in JobOrder Emb/Printing columns
      }
      
      return str.charAt(0);
    };

    const doc = new jsPDF({
      orientation: 'landscape',
      unit: 'mm',
      format: 'a3',
      compress: true
    });

    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const margin = 4;
    const contentWidth = pageWidth - (margin * 2);

    const addPageHeader = () => {
      doc.setFontSize(16);
      doc.setFont('times', 'bold');
      doc.setTextColor(15, 76, 129);
      doc.text('PRODUCTION TRACKING REPORT', pageWidth / 2, 12, { align: 'center' });
      
      // Add summary line with total lots and PCS
      doc.setFontSize(11);
      doc.setFont('times', 'bold');
      doc.setTextColor(0, 100, 0); // Green color for totals
      doc.text(`Total Lots: ${totalLots} | Total PCS: ${totalPCS.toLocaleString()}`, pageWidth / 2, 19, { align: 'center' });
      
      doc.setFontSize(9);
      doc.setFont('times', 'normal');
      doc.setTextColor(100, 100, 100);
      const today = new Date();
      const reportDate = `${String(today.getDate()).padStart(2, '0')}/${String(today.getMonth() + 1).padStart(2, '0')}/${String(today.getFullYear()).slice(-2)}`;
      doc.text(`Report Date: ${reportDate} | Records: ${exportData.length}`, pageWidth / 2, 25, { align: 'center' });
    };

    // Function to add column definitions notes section
    const addColumnDefinitions = (startY) => {
      // Section header
      doc.setFontSize(11);
      doc.setFont('times', 'bold');
      doc.setTextColor(15, 76, 129);
      doc.text('Column Definitions:', margin + 5, startY);
      
      doc.setFontSize(9);
      doc.setFont('times', 'normal');
      doc.setTextColor(0, 0, 0);
      
      let currentY = startY + 7;
      const lineHeight = 5;
      
      // Create definitions text in a simple format
      let definitionsText = '';
      COLUMN_DEFINITIONS.forEach((definition, index) => {
        const separator = index < COLUMN_DEFINITIONS.length - 1 ? ', ' : '';
        definitionsText += `${definition.abbreviation}: ${definition.fullForm}${separator}`;
      });
      
      // Split text into lines that fit the available width
      const maxLineWidth = pageWidth - (margin * 2) - 10;
      const textLines = doc.splitTextToSize(definitionsText, maxLineWidth);
      
      // Add all definition lines
      for (let i = 0; i < textLines.length; i++) {
        // Add the text line
        doc.text(textLines[i], margin + 5, currentY);
        currentY += lineHeight;
      }
    };

    addPageHeader();

    const columnCount = PDF_HEADERS.length;
    const baseColumnWidth = Math.min(13, contentWidth / columnCount);

    const columnWidths = PDF_HEADERS.map((header, index) => {
      if (index < 21) {
        switch(index) {
          case 0: return baseColumnWidth * 0.6;
          case 1: return baseColumnWidth * 1.0;
          case 2: return baseColumnWidth * 1.5;
          case 3: return baseColumnWidth * 1.5;
          case 4: return baseColumnWidth * 0.8;
          case 5: return baseColumnWidth * 0.6;
          case 6: return baseColumnWidth * 0.9;
          case 7: return baseColumnWidth * 0.9;
          case 8: return baseColumnWidth * 1.2;
          case 9: return baseColumnWidth * 0.9;
          case 10: return baseColumnWidth * 1.2;
          case 11: return baseColumnWidth * 1.2;
          case 12: return baseColumnWidth * 1.2;
          case 13: return baseColumnWidth * 1.2;
          case 14: return baseColumnWidth * 1.2;
          case 15: return baseColumnWidth * 1.5;
          case 16: return baseColumnWidth * 1.2;
          case 17: return baseColumnWidth * 1.5;
          case 18: return baseColumnWidth * 1.3;
          case 19: return baseColumnWidth * 1.2;
          case 20: return baseColumnWidth * 1.0;
          default: return baseColumnWidth * 1.0;
        }
      } else {
        switch(index) {
          case 21: return baseColumnWidth * 1.1;
          case 22: return baseColumnWidth * 1.0;
          case 23: return baseColumnWidth * 1.2;
          case 24: return baseColumnWidth * 1.0;
          case 25: return baseColumnWidth * 1.2;
          case 26: return baseColumnWidth * 1.2;
          case 27: return baseColumnWidth * 1.3;
          default: return baseColumnWidth * 1.0;
        }
      }
    });

    const tableBody = exportData.map((item, rowIndex) => {
      return HEADER_MAPPING.map((headerInfo, colIndex) => {
        const rawValue = getDataValue(item, headerInfo.dataKey, headerInfo.isSerial, rowIndex);
        let displayValue = '';
        
        // For Emb/Print columns, check if Design Work is direct stitching
        if (colIndex === 11 || colIndex === 12) {
          const designWorkValue = getDataValue(item, 'Design Work');
          const miniItem = {
            'Emb': item['Emb'] || (item._raw && item._raw['Emb']) || '',
            'Printing': item['Printing'] || (item._raw && item._raw['Printing']) || '',
            _raw: {
              'Emb': item['Emb'] || (item._raw && item._raw['Emb']) || '',
              'Printing': item['Printing'] || (item._raw && item._raw['Printing']) || ''
            }
          };
          const formattedDesignWork = formatDirectStitching(designWorkValue, miniItem);
          const isDirectStitching = formattedDesignWork === 'Y';
          
          if (isDirectStitching) {
            displayValue = 'Direct';
          }
        }
        
        if (colIndex === 15) {
          const compStitRawValue = getDataValue(item, HEADER_MAPPING[16].dataKey);
          const compStitValue = formatDateForPDF(compStitRawValue);
          
          if (isValidDate(compStitRawValue) || (compStitValue !== '-' && compStitValue !== '')) {
            displayValue = 'Done';
          } else {
            displayValue = rawValue || '-';
          }
        } 
        else if (headerInfo.type === 'calculated') {
          displayValue = rawValue || '-';
        }
        else {
          switch(headerInfo.type) {
            case 'date': 
              displayValue = formatDateForPDF(rawValue); 
              break;
            case 'stitching': 
              // Create mini-item with Emb/Printing data for Design Work column
              const miniItemForDesignWork = {
                'Emb': item['Emb'] || (item._raw && item._raw['Emb']) || '',
                'Printing': item['Printing'] || (item._raw && item._raw['Printing']) || '',
                _raw: {
                  'Emb': item['Emb'] || (item._raw && item._raw['Emb']) || '',
                  'Printing': item['Printing'] || (item._raw && item._raw['Printing']) || ''
                }
              };
              displayValue = formatDirectStitching(rawValue, miniItemForDesignWork); 
              break;
            case 'number':
              if (headerInfo.isSerial) displayValue = rawValue;
              else if (rawValue && !isNaN(rawValue)) displayValue = parseInt(rawValue).toLocaleString();
              else displayValue = rawValue || '-';
              break;
            case 'text': displayValue = rawValue || '-'; break;
            default: displayValue = rawValue || '-';
          }
        }
        
        return displayValue;
      });
    });

    autoTable(doc, {
      head: [PDF_HEADERS],
      body: tableBody,
      startY: 30, // Increased from 25 to accommodate the new summary line
      theme: 'grid',
      headStyles: {
        fillColor: [15, 76, 129],
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        fontSize: 8,
        cellPadding: { top: 3, right: 2, bottom: 3, left: 2 },
        halign: 'center',
        valign: 'middle'
      },
      bodyStyles: {
        fontSize: 10,
        cellPadding: { top: 2, right: 2, bottom: 2, left: 2 },
        lineWidth: 0.2,
        lineColor: [0, 0, 0],
        textColor: [0, 0, 0],
        font: 'times',
        fontStyle: 'normal',
        valign: 'middle',
        overflow: 'linebreak',
        minCellHeight: 6,
        lineHeight: 1.3
      },
      styles: {
        fontSize: 10,
        cellPadding: 1.5,
        lineWidth: 0.1,
        lineColor: [200, 200, 200],
        font: 'times',
        fontStyle: 'normal',
        valign: 'middle',
        overflow: 'linebreak'
      },
      columnStyles: PDF_HEADERS.reduce((styles, header, index) => {
        styles[index] = {
          cellWidth: columnWidths[index],
          halign: 'center',
          valign: 'middle',
          fontSize: 10,
          fontStyle: 'normal',
          overflow: 'linebreak',
          cellPadding: { top: 1.5, right: 1.5, bottom: 1.5, left: 1.5 }
        };
        
        if (index < 21) {
          if (index === 0) styles[index].textColor = [100, 100, 100];
          else if (index === 1 || index === 5 || index === 9) styles[index].textColor = [255, 0, 0];
          else if (index === 4) styles[index].textColor = [128, 0, 128];
          else if (index === 14) styles[index].textColor = [0, 0, 255];
          else if (index === 15) styles[index].textColor = [255, 0, 0];
        }
        
        return styles;
      }, {}),
      margin: { top: 30, left: margin, right: margin, bottom: 15 },
      tableWidth: 'auto',
      showHead: 'everyPage',
      showFoot: false,
      pageBreak: 'auto',
      rowPageBreak: 'auto',
      didParseCell: function(data) {
        if (data.section === 'body') {
          if (data.column.index === 15 && data.cell.raw === 'Done') {
            data.cell.styles.textColor = [0, 128, 0];
          }
          
          if ((data.column.index === 11 || data.column.index === 12) && 
              data.cell.raw === 'Direct') {
            data.cell.styles.textColor = [0, 0, 255];
            data.cell.styles.fontStyle = 'bolditalic';
          }
          
          const daysColumns = [21, 22, 23, 24, 25, 26, 27];
          if (daysColumns.includes(data.column.index)) {
            if (data.cell.raw === '-') {
              data.cell.styles.textColor = [100, 100, 100];
            } else {
              let daysStr = data.cell.raw;
              let days = 0;
              
              if (daysStr.startsWith('+')) {
                days = parseInt(daysStr.substring(1));
              } else if (daysStr.startsWith('-')) {
                days = parseInt(daysStr);
              } else if (!isNaN(parseInt(daysStr))) {
                days = parseInt(daysStr);
              }
              
              if (!isNaN(days)) {
                if (days < 0) {
                  data.cell.styles.textColor = [255, 0, 0];
                  data.cell.styles.fontStyle = 'bold';
                }
                else if (days >= 0) {
                  const positiveDays = Math.abs(days);
                  
                  if (data.column.index === 21) {
                    data.cell.styles.textColor = positiveDays > 2 ? [255, 0, 0] : [0, 128, 0];
                  } 
                  else if (data.column.index === 22) {
                    if (positiveDays > 10) data.cell.styles.textColor = [255, 0, 0];
                    else if (positiveDays > 5) data.cell.styles.textColor = [255, 165, 0];
                    else data.cell.styles.textColor = [0, 128, 0];
                  }
                  else if (data.column.index === 23) {
                    if (positiveDays <= 6) data.cell.styles.textColor = [0, 128, 0];
                    else if (positiveDays <= 10) data.cell.styles.textColor = [255, 165, 0];
                    else data.cell.styles.textColor = [255, 0, 0];
                  }
                  else if (data.column.index === 24) {
                    data.cell.styles.textColor = positiveDays > 2 ? [255, 0, 0] : [0, 128, 0];
                  }
                  else if (data.column.index === 25) {
                    if (positiveDays <= 1) data.cell.styles.textColor = [0, 128, 0];
                    else if (positiveDays <= 3) data.cell.styles.textColor = [255, 165, 0];
                    else data.cell.styles.textColor = [255, 0, 0];
                  }
                  else if (data.column.index === 26) {
                    if (positiveDays <= 1) data.cell.styles.textColor = [0, 128, 0];
                    else if (positiveDays <= 3) data.cell.styles.textColor = [255, 165, 0];
                    else data.cell.styles.textColor = [255, 0, 0];
                  }
                  else if (data.column.index === 27) {
                    if (positiveDays <= 1) data.cell.styles.textColor = [0, 128, 0];
                    else if (positiveDays <= 3) data.cell.styles.textColor = [255, 165, 0];
                    else data.cell.styles.textColor = [255, 0, 0];
                  }
                }
              }
            }
          }
        }
      },
      didDrawPage: function(data) {
        addPageHeader();
        
        // Add summary at the bottom of each page
        const summaryY = pageHeight - 10;
        
        doc.setFontSize(10);
        doc.setFont('times', 'bold');
        doc.setTextColor(15, 76, 129);
        doc.text(`Total: ${totalLots} Lots | ${totalPCS.toLocaleString()} PCS`, pageWidth / 2, summaryY, { align: 'center' });
        
        doc.setFontSize(9);
        doc.setTextColor(100, 100, 100);
        doc.text(
          `Page ${data.pageNumber} of ${data.pageCount}`,
          pageWidth - margin - 5,
          pageHeight - 5
        );
        
        // Add column definitions ONLY on the last page, after the table
        if (data.pageNumber === data.pageCount) {
          // Get the cursor position after the table ends
          const finalY = data.cursor ? data.cursor.y : pageHeight - 30;
          
          // Check if we have enough space for definitions (leave space for footer)
          const availableSpace = pageHeight - finalY - 20; // Increased from 15 to 20
          
          if (availableSpace > 30) { // If we have at least 30mm space
            // Add column definitions directly on the same page
            addColumnDefinitions(finalY + 10);
          }
          // If not enough space, don't add definitions - they'll show at the bottom of the last page
          // but won't create a new page
        }
      }
    });

    const today = new Date();
    const fileName = `Production_Tracking_Report_${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, '0')}${String(today.getDate()).padStart(2, '0')}.pdf`;
    
    doc.save(fileName);
    
    setTimeout(() => {
      alert(`✅ PDF Exported Successfully!\n📄 ${fileName}\n📊 ${totalLots} lots | ${totalPCS.toLocaleString()} PCS`);
    }, 500);
    
  } catch (error) {
    console.error('PDF Export Error Details:', error);
    alert(`❌ Failed to export PDF: ${error.message}`);
  } finally {
    setExportLoading(false);
  }
};

  const formatDateForPDF = (dateString) => {
    if (!dateString || dateString === '' || dateString === '-' || 
        dateString.toLowerCase() === 'null' || dateString.toLowerCase() === 'invalid date') {
      return '-';
    }
    
    try {
      if (typeof dateString === 'string' && !isNaN(dateString) && dateString !== '-') {
        return dateString;
      }
      
      const date = parseDate(dateString);
      if (date) {
        const day = String(date.getDate()).padStart(2, '0');
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const year = String(date.getFullYear()).slice(-2);
        return `${day}/${month}/${year}`;
      }
      
      return dateString.toString().trim();
    } catch {
      return dateString;
    }
  };

  const formatDirectStitching = (value, item = {}) => {
    if (!value || value === '') return '-';
    
    const str = String(value).toUpperCase();
    if (str.includes('YES') || str === 'Y') {
      return 'Direct';
    }
    
    if (str.includes('NO') || str === 'N') {
      const embFromJobOrder = item['Emb'] || item._raw?.['Emb'] || '';
      const printingFromJobOrder = item['Printing'] || item._raw?.['Printing'] || '';
      
      const embClean = embFromJobOrder ? String(embFromJobOrder).trim() : '';
      const printingClean = printingFromJobOrder ? String(printingFromJobOrder).trim() : '';
      
      const hasEmbFromJobOrder = embClean !== '' && 
                                embClean.toUpperCase() !== 'NA' && 
                                embClean.toUpperCase() !== 'N/A' &&
                                embClean !== '-';
      
      const hasPrintingFromJobOrder = printingClean !== '' && 
                                     printingClean.toUpperCase() !== 'NA' && 
                                     printingClean.toUpperCase() !== 'N/A' &&
                                     printingClean !== '-';
      
      if (hasEmbFromJobOrder && hasPrintingFromJobOrder) return 'Emb+Prt';
      if (hasEmbFromJobOrder) return 'Emb';
      if (hasPrintingFromJobOrder) return 'Print';
      return 'N';
    }
    
    return str.charAt(0);
  };

  // Export to PDF WITH IMAGES
  const exportToPDFWithImages = async () => {
    setExportLoading(true);
    
    try {
      const exportData = prepareExportData(filteredData);
      
      if (exportData.length === 0) {
        alert('No data available to export.');
        setExportLoading(false);
        return;
      }

      const totalPCS = exportData.reduce((sum, item) => {
        const pcsValue = item['PCS'];
        const numValue = parseInt(pcsValue);
        return sum + (isNaN(numValue) ? 0 : numValue);
      }, 0);
      
      const totalLots = exportData.length;

      const HEADER_MAPPING_IMAGES = [
        { pdf: 'Sr.', dataKey: 'Sr.', type: 'number', isSerial: true },
        { pdf: 'Lot', dataKey: 'Lot No', type: 'text' },
        { pdf: 'Image', dataKey: 'Image', type: 'image' },
        { pdf: 'Fabric', dataKey: 'Fabric', type: 'text' },
        { pdf: 'Size', dataKey: 'Size', type: 'text' },
        { pdf: 'Item', dataKey: 'Item', type: 'text' },
        { pdf: 'Brand', dataKey: 'Brand', type: 'text' },
        { pdf: 'Party', dataKey: 'Party', type: 'text' },
        { pdf: 'Sea', dataKey: 'Season', type: 'text' },
        { pdf: 'M/W/K', dataKey: 'Section', type: 'text' },
        { pdf: 'Design', dataKey: 'Design Work', type: 'stitching' }, 
        { pdf: 'Job Date', dataKey: 'Job Date', type: 'date' },
        { pdf: 'PCS', dataKey: 'PCS', type: 'number' },
        { pdf: 'Cut Date', dataKey: 'Cut Date', type: 'date' },
        { pdf: 'Emb/Print Issue', dataKey: 'Emb/Print Issue', type: 'date' },
        { pdf: 'Emb/print Comp', dataKey: 'Emb/print Comp', type: 'date' },
        { pdf: 'Stit Date', dataKey: 'Stit Date', type: 'date' },
        { pdf: 'Stit Sup', dataKey: 'Stit Sup', type: 'text' },
        { pdf: 'WIP Stit', dataKey: 'WIP Stit', type: 'text' },
        { pdf: 'Comp Stit', dataKey: 'Comp Stit', type: 'date' },
        { pdf: 'Pkg Sup', dataKey: 'Pkg Sup', type: 'text' },
        { pdf: 'Pkg Date', dataKey: 'Pkg Date', type: 'date' },
        { pdf: 'WIP Pkg', dataKey: 'WIP Pkg', type: 'text' },
        { pdf: 'Pkg Comp', dataKey: 'Pkg Comp', type: 'text' }
      ];

      const PDF_HEADERS = HEADER_MAPPING_IMAGES.map(h => h.pdf);

      // Pre-fetch Base64 images for all rows
      const imageMap = {};
      for (let i = 0; i < exportData.length; i++) {
        const item = exportData[i];
        const imgUrl = item['Image'] || item['Image URL'] || (item._raw && item._raw['Image URL']);
        if (imgUrl) {
          const base64 = await getBase64ImageFromUrl(imgUrl);
          if (base64) {
            imageMap[i] = base64;
          }
        }
      }

      const doc = new jsPDF({
        orientation: 'landscape',
        unit: 'mm',
        format: 'a3',
        compress: true
      });

      const pageWidth = doc.internal.pageSize.getWidth();
      const pageHeight = doc.internal.pageSize.getHeight();
      const margin = 6;
      const availableWidth = pageWidth - (margin * 2);

      const addPageHeader = () => {
        doc.setFontSize(16);
        doc.setFont('times', 'bold');
        doc.setTextColor(15, 76, 129);
        doc.text('PRODUCTION TRACKING REPORT (WITH IMAGES)', pageWidth / 2, 12, { align: 'center' });
        
        doc.setFontSize(11);
        doc.setFont('times', 'bold');
        doc.setTextColor(0, 100, 0);
        doc.text(`Total Lots: ${totalLots} | Total PCS: ${totalPCS.toLocaleString()}`, pageWidth / 2, 19, { align: 'center' });
        
        doc.setFontSize(9);
        doc.setFont('times', 'normal');
        doc.setTextColor(100, 100, 100);
        const today = new Date();
        const reportDate = `${String(today.getDate()).padStart(2, '0')}/${String(today.getMonth() + 1).padStart(2, '0')}/${String(today.getFullYear()).slice(-2)}`;
        doc.text(`Report Date: ${reportDate} | Records: ${exportData.length}`, pageWidth / 2, 25, { align: 'center' });
      };

      const tableBody = exportData.map((item, rowIndex) => {
        return HEADER_MAPPING_IMAGES.map((headerInfo) => {
          if (headerInfo.type === 'image') return '';
          const rawValue = item[headerInfo.dataKey] !== undefined ? item[headerInfo.dataKey] : (item._raw ? item._raw[headerInfo.dataKey] : '');
          if (headerInfo.isSerial) return (rowIndex + 1).toString();
          if (headerInfo.type === 'date') return formatDateForPDF(rawValue);
          if (headerInfo.type === 'stitching') {
            const miniItem = {
              'Emb': item['Emb'] || (item._raw && item._raw['Emb']) || '',
              'Printing': item['Printing'] || (item._raw && item._raw['Printing']) || '',
              _raw: {
                'Emb': item['Emb'] || (item._raw && item._raw['Emb']) || '',
                'Printing': item['Printing'] || (item._raw && item._raw['Printing']) || ''
              }
            };
            return formatDirectStitching(rawValue, miniItem);
          }
          if (headerInfo.type === 'number') {
            return rawValue && !isNaN(rawValue) ? parseInt(rawValue).toLocaleString() : (rawValue || '-');
          }
          return rawValue || '-';
        });
      });

    // Proportional column widths calculation to guarantee no overflow
    const baseWidths = [
      8,  // 0: Sr.
      18, // 1: Lot
      16, // 2: Image
      22, // 3: Fabric
      14, // 4: Size
      22, // 5: Item
      18, // 6: Brand
      18, // 7: Party
      10, // 8: Sea
      12, // 9: M/W/K
      16, // 10: Design
      16, // 11: Job Date
      14, // 12: PCS
      16, // 13: Cut Date
      18, // 14: Emb/Print Issue
      18, // 15: Emb/print Comp
      16, // 16: Stit Date
      16, // 17: Stit Sup
      18, // 18: WIP Stit
      16, // 19: Comp Stit
      16, // 20: Pkg Sup
      16, // 21: Pkg Date
      18, // 22: WIP Pkg
      16  // 23: Pkg Comp
    ];
    const totalBaseWidth = baseWidths.reduce((sum, w) => sum + w, 0);
    const scaleFactor = availableWidth / totalBaseWidth;
    const finalColumnStyles = baseWidths.reduce((acc, w, idx) => {
      acc[idx] = {
        cellWidth: w * scaleFactor,
        halign: idx === 4 || idx === 6 || idx === 21 ? 'left' : 'center',
        valign: 'middle'
      };
      return acc;
    }, {});

    autoTable(doc, {
      head: [PDF_HEADERS],
      body: tableBody,
      startY: 30,
      margin: { top: 30, left: margin, right: margin, bottom: 15 },
      tableWidth: availableWidth,
      theme: 'grid',
      styles: {
        lineWidth: 0.5,
        lineColor: [0, 0, 0],
        font: 'times',
        fontStyle: 'normal',
        valign: 'middle',
        overflow: 'linebreak'
      },
      headStyles: {
        fillColor: [15, 76, 129],
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        fontSize: 7.5,
        cellPadding: 1,
        halign: 'center',
        valign: 'middle',
        lineWidth: 0.5,
        lineColor: [0, 0, 0]
      },
      bodyStyles: {
        fontSize: 7.5,
        cellPadding: 1,
        minCellHeight: 14,
        valign: 'middle',
        overflow: 'linebreak',
        lineWidth: 0.5,
        lineColor: [0, 0, 0],
        textColor: [0, 0, 0]
      },
      columnStyles: finalColumnStyles,
      didDrawCell: (data) => {
        if (data.section === 'body' && data.column.index === 2) {
          const rowIndex = data.row.index;
          const base64Img = imageMap[rowIndex];
          if (base64Img) {
            try {
              const cell = data.cell;
              const imgWidth = Math.min(12, cell.width - 2);
              const imgHeight = Math.min(12, cell.height - 2);
              const x = cell.x + (cell.width - imgWidth) / 2;
              const y = cell.y + (cell.height - imgHeight) / 2;
              const format = base64Img.startsWith('data:image/png') ? 'PNG' : 'JPEG';
              doc.addImage(base64Img, format, x, y, imgWidth, imgHeight);
            } catch (imgErr) {
              console.warn('Error rendering cell image in PDF:', imgErr);
            }
          }
        }
      },
      didDrawPage: (data) => {
        addPageHeader();
        doc.setFontSize(9);
        doc.setTextColor(100, 100, 100);
        doc.text(
          `Page ${data.pageNumber} of ${data.pageCount}`,
          pageWidth - margin - 5,
          pageHeight - 5
        );
      }
    });

    const today = new Date();
    const fileName = `Production_Tracking_Report_With_Images_${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, '0')}${String(today.getDate()).padStart(2, '0')}.pdf`;
    
    doc.save(fileName);
    
    setTimeout(() => {
      alert(`✅ PDF With Images Exported Successfully!\n📄 ${fileName}\n📊 ${totalLots} lots | ${totalPCS.toLocaleString()} PCS`);
    }, 500);

  } catch (error) {
    console.error('PDF Export With Images Error:', error);
    alert(`❌ Failed to export PDF with images: ${error.message}`);
  } finally {
    setExportLoading(false);
  }
};

const exportCurrentView = (format) => {
  if (filteredData.length === 0) {
    alert('No data to export!');
    return;
  }
  
  if (format === 'excel') {
    exportToExcel();
  } else if (format === 'pdf') {
    exportToPDF();
  } else if (format === 'pdf-no-days') {
    exportToPDFWithoutDays();
  } else if (format === 'pdf-images') {
    exportToPDFWithImages();
  }
};

  // Export all data (unfiltered)
  const exportAllData = (format) => {
    if (data.length === 0) {
      alert('No data to export!');
      return;
    }
    
    if (format === 'excel') {
      const allValidData = data.filter(item => !item._isCancelled);
      exportToExcel(allValidData);
    } else if (format === 'pdf') {
       const allValidData = data.filter(item => !item._isCancelled);
       const allData = prepareExportData(allValidData);
      
      const doc = new jsPDF({
        orientation: 'landscape',
        unit: 'mm',
        format: 'a3'
      });
      
      doc.setFontSize(16);
      doc.text('All Production Data Report', 14, 15);
      doc.setFontSize(10);
      doc.text(`Generated on: ${new Date().toLocaleString()}`, 14, 22);
      doc.text(`Total Records: ${allData.length} (All Data)`, 14, 28);
      
      const tableData = allData.map(item => 
        DISPLAY_HEADERS.map(header => item[header] || '')
      );
      
      autoTable(doc, {
        head: [DISPLAY_HEADERS],
        body: tableData,
        startY: 35,
        theme: 'grid',
        styles: {
          fontSize: 7,
          cellPadding: 2,
          overflow: 'linebreak',
          halign: 'left'
        },
        headStyles: {
          fillColor: [41, 128, 185],
          textColor: 255,
          fontStyle: 'bold'
        },
        columnStyles: {
          0: { cellWidth: 15 },
          1: { cellWidth: 15 },
          2: { cellWidth: 15 },
          3: { cellWidth: 15 },
          4: { cellWidth: 15 },
          5: { cellWidth: 15 },
          6: { cellWidth: 15 },
          7: { cellWidth: 15 },
          8: { cellWidth: 15 },
          9: { cellWidth: 15 },
          10: { cellWidth: 20 },
          11: { cellWidth: 20 },
          12: { cellWidth: 20 },
          13: { cellWidth: 15 },
          14: { cellWidth: 15 },
          15: { cellWidth: 15 },
          16: { cellWidth: 15 },
          17: { cellWidth: 15 },
          18: { cellWidth: 15 },
          19: { cellWidth: 15 },
          20: { cellWidth: 15 },
          21: { cellWidth: 15 },
          22: { cellWidth: 15 },
          23: { cellWidth: 15 },
          24: { cellWidth: 15 },
          25: { cellWidth: 15 },
          26: { cellWidth: 15 },
          27: { cellWidth: 15 }
        }
      });
      
      const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
      const filename = `All_Production_Data_${timestamp}.pdf`;
      doc.save(filename);
      
      console.log(`Exported all ${allData.length} records to PDF`);
    }
  };

  return (
    <div className="overall-cutting-packing">
      {/* Header Section */}
      <div className="op-header">
        <div className="op-header-ambient-wrapper">
          <div className="op-header-ambient-1" />
          <div className="op-header-ambient-2" />
        </div>

        <div className="op-header-top-row">
          <div className="op-pill-badge">
            <span className="op-pulse-dot" />
            <span>Garment Production Suite</span>
            <span className="op-pill-divider">|</span>
            <span className="op-pill-sub">Overall Cutting to Packing Report</span>
          </div>

          <div className="op-header-right-meta">
            <span>📅 {new Date().toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' })}</span>
            <span className="op-status-active">
              <span className="op-pulse">🟢</span> Live Sync
            </span>
          </div>
        </div>

        <div className="op-header-main-grid">
          <div>
            <h1 className="op-title">
              Overall Cutting to Packing
            </h1>
            <p className="op-subtitle">
              Comprehensive real-time tracking across JobOrder, Index, and Issues sheets.
            </p>
          </div>

          {!loading && !error && data.length > 0 && (
            <div className="op-header-stats-grid">
              <div className="op-stat-card">
                <span className="op-stat-label">Total Lots</span>
                <span className="op-stat-value">{filteredData.length}</span>
                <span className="op-stat-sub">Active in view</span>
              </div>

              <div className="op-stat-card">
                <span className="op-stat-label">Stitching Done</span>
                <span className="op-stat-value" style={{ color: '#34d399' }}>
                  {filteredData.filter(i => {
                    const val = (i['Comp Stit'] || '').toString().toLowerCase();
                    return val && val !== '-' && !val.includes('pending');
                  }).length}
                </span>
                <span className="op-stat-sub">Completed</span>
              </div>

              <div className="op-stat-card">
                <span className="op-stat-label">Pkg Done</span>
                <span className="op-stat-value" style={{ color: '#38bdf8' }}>
                  {filteredData.filter(i => {
                    const val = (i['Pkg Comp'] || '').toString().toLowerCase();
                    return val && val !== '-' && !val.includes('pending');
                  }).length}
                </span>
                <span className="op-stat-sub">Completed</span>
              </div>

              <div className="op-stat-card">
                <span className="op-stat-label">Hold Lots</span>
                <span className="op-stat-value" style={{ color: '#fbbf24' }}>
                  {filteredData.filter(i => i['WIP Pkg._parsed']?.isHold || i._isHoldLot).length}
                </span>
                <span className="op-stat-sub">Requires action</span>
              </div>
            </div>
          )}
        </div>

        <div className="op-top-actions-bar">
          {!loading && !error && data.length > 0 && (
            <button 
              onClick={() => setShowPartyInitials(!showPartyInitials)}
              className="op-btn op-btn-toggle"
              title={showPartyInitials ? "Switch to full party names" : "Switch to party initials"}
            >
              {showPartyInitials ? 'Show Full Party Names' : 'Show Party Initials'}
            </button>
          )}

          {!loading && !error && data.length > 0 && (
            <div className="export-buttons">
              <div className="dropdown">
                <button 
                  className="op-btn op-btn-export"
                  disabled={exportLoading}
                >
                  {exportLoading ? '🔄 Exporting...' : '📥 Export Data'}
                </button>
                <div className="dropdown-content">
                  <div className="dropdown-section">
                    <strong>Export Current View</strong>
                    <p className="dropdown-info">({filteredData.length} filtered records)</p>
                    <button 
                      onClick={() => exportCurrentView('excel')}
                      className="dropdown-item"
                      disabled={exportLoading || filteredData.length === 0}
                    >
                      <span className="export-icon">📊</span> Excel (Current View)
                    </button>
                    <button 
                      onClick={() => exportCurrentView('pdf')}
                      className="dropdown-item"
                      disabled={exportLoading || filteredData.length === 0}
                    >
                      <span className="export-icon">📄</span> PDF (Current View)
                    </button>
                    <button 
                      onClick={() => exportCurrentView('pdf-no-days')}
                      className="dropdown-item"
                      disabled={exportLoading || filteredData.length === 0}
                    >
                      <span className="export-icon">📄</span> PDF (Without Days Calculation)
                    </button>
                    <button 
                      onClick={() => exportCurrentView('pdf-images')}
                      className="dropdown-item"
                      disabled={exportLoading || filteredData.length === 0}
                    >
                      <span className="export-icon">🖼️</span> PDF (With Images)
                    </button>
                  </div>
                  <div className="dropdown-section">
                    <strong>Export All Data</strong>
                    <p className="dropdown-info">({data.length} total records)</p>
                    <button 
                      onClick={() => exportAllData('excel')}
                      className="dropdown-item"
                      disabled={exportLoading}
                    >
                      <span className="export-icon">📊</span> Excel (All Data)
                    </button>
                    <button 
                      onClick={() => exportAllData('pdf')}
                      className="dropdown-item"
                      disabled={exportLoading}
                    >
                      <span className="export-icon">📄</span> PDF (All Data)
                    </button>
                  </div>
                </div>
              </div>
            </div>
          )}

          <button 
            onClick={fetchProductionData} 
            className="op-btn op-btn-refresh"
            disabled={loading}
          >
            {loading ? '🔄 Loading...' : '🔄 Refresh Data'}
          </button>

          <button 
            onClick={() => window.history.back()}
            className="op-btn op-btn-toggle op-btn-back"
            title="Go back to previous page"
          >
            ← Back
          </button>
        </div>
      </div>

      {/* Loading State */}
      {loading && (
        <div className="op-loading-overlay">
          <div className="op-spinner-large"></div>
          <p>Fetching data from JobOrder, Index, and Issues sheets...</p>
          <p>Looking for: Party, Season, Section, D.Stitching columns</p>
        </div>
      )}

      {/* Error State */}
      {error && !loading && (
        <div className="op-error-banner">
          <div className="op-error-content">
            <span className="error-icon">⚠️</span>
            <div>
              <h3>Connection Error</h3>
              <p>{error}</p>
              <button 
                onClick={fetchProductionData}
                className="op-btn op-btn-retry"
              >
                Try Again
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Filters Section */}
      {!loading && !error && data.length > 0 && (
        <div className="op-filters-section">
          {/* Main Search Bar */}
          <div className="op-search-box">
            <input
              type="text"
              placeholder="🔍 Search across all columns (Lot, Brand, Party, Style, Supervisor, Remarks...)"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="op-search-input"
            />
            {searchTerm && (
              <button 
                onClick={() => setSearchTerm('')}
                className="op-search-clear"
                title="Clear search"
              >
                ✕
              </button>
            )}
          </div>
          
          <div className="op-filter-grid">
            <MultiSelectFilter
              label="Item"
              placeholder="Select items..."
              options={getUniqueValues('Garment Type')}
              selected={garmentTypeFilter}
              onChange={setGarmentTypeFilter}
            />

            <MultiSelectFilter
              label="Brand"
              placeholder="Select brands..."
              options={getUniqueValues('Brand')}
              selected={brandFilter}
              onChange={setBrandFilter}
            />

            <MultiSelectFilter
              label="Style"
              placeholder="Select styles..."
              options={getUniqueValues('Style')}
              selected={styleFilter}
              onChange={setStyleFilter}
            />

            <MultiSelectFilter
              label="Party"
              placeholder="Select parties..."
              options={getUniqueValues('Party Name')}
              selected={partyNameFilter}
              onChange={setPartyNameFilter}
            />

            <MultiSelectFilter
              label="Fabric"
              placeholder="Select fabrics..."
              options={getUniqueFabrics()}
              selected={fabricFilter}
              onChange={setFabricFilter}
            />

            <MultiSelectFilter
              label="Stit Sup"
              placeholder="Select supervisors..."
              options={getUniqueValues('Stit Sup')}
              selected={supervisorFilter}
              onChange={setSupervisorFilter}
            />

            <MultiSelectFilter
              label="Pkg Sup"
              placeholder="Select pkg supervisors..."
              options={getUniqueValues('Pkg Sup')}
              selected={pkgSupervisorFilter}
              onChange={setPkgSupervisorFilter}
            />

            <MultiSelectFilter
              label="Emb/Print Comp"
              placeholder="Select status..."
              options={['Complete', 'Pending']}
              selected={embPrintCompFilter}
              onChange={setEmbPrintCompFilter}
            />

            <MultiSelectFilter
              label="Comp Stit"
              placeholder="Select status..."
              options={['Complete', 'Pending']}
              selected={compStatusFilter}
              onChange={setCompStatusFilter}
            />

            <MultiSelectFilter
              label="Pkg Comp"
              placeholder="Select status..."
              options={['Complete', 'Pending']}
              selected={pkgCompFilter}
              onChange={setPkgCompFilter}
            />

            <MultiSelectFilter
              label="D. Stitching"
              placeholder="Select D.Stitching..."
              options={['YES', 'NO']}
              selected={directStitchingFilter}
              onChange={setDirectStitchingFilter}
            />

            <MultiSelectFilter
              label="Season"
              placeholder="Select seasons..."
              options={getUniqueValues('Season')}
              selected={seasonFilter}
              onChange={setSeasonFilter}
            />

            <MultiSelectFilter
              label="Section"
              placeholder="Select sections..."
              options={getUniqueValues('Section')}
              selected={sectionFilter}
              onChange={setSectionFilter}
            />

            {/* Hold Lots Toggle Button */}
            <div className="filter-group">
              <label>Hold Lots</label>
              <button
                type="button"
                onClick={() => setHoldLotsFilter(!holdLotsFilter)}
                className={`op-filter-select op-hold-toggle ${holdLotsFilter ? 'active' : ''}`}
              >
                {holdLotsFilter ? '⚠️ Hold Lots Only' : 'All Lots'}
              </button>
            </div>

            {/* Date Filter Target */}
            <div className="filter-group">
              <label>Date Target</label>
              <select 
                value={dateFilterField} 
                onChange={(e) => setDateFilterField(e.target.value)}
                className="op-filter-select"
              >
                <option value="Job Date">Job Date</option>
                <option value="Cut Date">Cut Date</option>
                <option value="Both">Both (Job / Cut)</option>
              </select>
            </div>

            {/* Start Date */}
            <div className="filter-group">
              <label>Start Date</label>
              <input 
                type="date"
                value={startDateFilter}
                onChange={(e) => setStartDateFilter(e.target.value)}
                className="op-filter-select"
              />
            </div>

            {/* End Date */}
            <div className="filter-group">
              <label>End Date</label>
              <input 
                type="date"
                value={endDateFilter}
                onChange={(e) => setEndDateFilter(e.target.value)}
                className="op-filter-select"
              />
            </div>

            {/* Action Row */}
            <div className="filter-group filter-actions">
              <button 
                onClick={clearAllFilters}
                className="op-btn op-btn-clear"
              >
                Clear All Filters
              </button>
              <div className="filter-info">
                Showing {filteredData.length} of {data.length} lots
                {[
                  brandFilter.length > 0,
                  garmentTypeFilter.length > 0,
                  styleFilter.length > 0,
                  partyNameFilter.length > 0,
                  seasonFilter.length > 0,
                  sectionFilter.length > 0,
                  directStitchingFilter.length > 0,
                  supervisorFilter.length > 0,
                  embPrintCompFilter.length > 0,
                  compStatusFilter.length > 0,
                  pkgSupervisorFilter.length > 0,
                  pkgCompFilter.length > 0,
                  fabricFilter.length > 0,
                  holdLotsFilter,
                  Boolean(searchTerm),
                  Boolean(startDateFilter),
                  Boolean(endDateFilter)
                ].some(Boolean) && (
                  <span className="active-filters-count">
                    ({[
                      brandFilter.length > 0,
                      garmentTypeFilter.length > 0,
                      styleFilter.length > 0,
                      partyNameFilter.length > 0,
                      seasonFilter.length > 0,
                      sectionFilter.length > 0,
                      directStitchingFilter.length > 0,
                      supervisorFilter.length > 0,
                      embPrintCompFilter.length > 0,
                      compStatusFilter.length > 0,
                      pkgSupervisorFilter.length > 0,
                      pkgCompFilter.length > 0,
                      fabricFilter.length > 0,
                      holdLotsFilter,
                      Boolean(searchTerm),
                      Boolean(startDateFilter),
                      Boolean(endDateFilter)
                    ].filter(Boolean).length} active filters)
                  </span>
                )}
              </div>
            </div>
          </div>

          {fabricFilter.length > 0 && (
            <div className="selected-fabrics">
              {fabricFilter.map(fabric => (
                <span key={fabric} className="selected-fabric-tag">
                  {fabric}
                  <button 
                    onClick={() => handleFabricSelect(fabric)}
                    className="remove-fabric"
                  >
                    ×
                  </button>
                </span>
              ))}
            </div>
          )}
        </div>
      )}

      {/* Data Table */}
     {/* Data Table */}
<div className="op-data-table-container" ref={tableRef}>
  {loading ? (
    <div className="op-loading">
      <div className="op-spinner"></div>
      <p>Loading merged data from all sheets...</p>
      <p>Fetching: Party, Season, Section, D.Stitching from JobOrder</p>
      <p>Fetching: Packing details from Issues sheet</p>
    </div>
  ) : error ? (
    <div className="op-error">
      <div className="op-error-icon">⚠️</div>
      <div>
        <h3>Data Load Error</h3>
        <p>{error}</p>
      </div>
    </div>
  ) : data.length === 0 ? (
    <div className="op-no-data">
      <p>No production data found in the Google Sheets.</p>
      <button 
        onClick={fetchProductionData}
        className="op-btn op-btn-retry"
      >
        Try Again
      </button>
    </div>
  ) : (
    <>
      <div className="op-table-wrapper">
        <table className="op-data-table">
          <thead>
            <tr>
              {headers.map((header, index) => (
                <th key={index} className={`header-${header.replace(/\s+/g, '-').toLowerCase()}`}>
                  {header}
                  {header === 'WIP Stit' && (
                    <div className="column-help">
                      (Latest Status/Remarks)
                    </div>
                  )}
                  {header === 'Comp Stit' && (
                    <div className="column-help">
                      (Completion Date)
                    </div>
                  )}
                  {(header === 'Pkg Sup' || header === 'Pkg Date' || 
                    header === 'WIP Pkg' || header === 'Pkg Comp') && (
                    <div className="column-help">
                      (From Issues Sheet)
                    </div>
                  )}
                  {(header === 'Cut Days' || header === 'Emb/Print Days' || 
                    header === 'Stit Days' || header === 'Pkg Days' ||
                    header === 'Cut To Emb/Print' || header === 'Emb/Print To Stit' || 
                    header === 'Stit To Check Pack') && (
                    <div className="column-help">
                      (Days Calculation)
                    </div>
                  )}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filteredData.map((row, rowIndex) => {
              const serialNumber = rowIndex + 1; // Calculate serial number
              const challanInfo = row['CHALLAN HISTORY_parsed'];
              const wipInfo = row['WIP Status_parsed'];
              const completedInfo = row['Completed Status_parsed'];
              const hasIndexData = indexData.some(indexItem => 
                indexItem['Lot Number'] === row['Lot No']
              );
              const hasIssuesData = issuesData.some(issuesItem => 
                issuesItem['Lot Number'] === row['Lot No']
              );
              
              return (
                <tr key={row._id} className="op-data-row">
                  {headers.map((header, colIndex) => {
                    // Handle serial number column separately
                    if (header === 'Sr.') {
                      return (
                        <td 
                          key={`${rowIndex}-${colIndex}`} 
                          className="serial-number-cell"
                          title={`Serial Number: ${serialNumber}`}
                        >
                          {serialNumber}
                        </td>
                      );
                    }
                    
                    const value = row[header] || '';
                    const rawValue = row._raw?.[header] || '';
                    let cellClass = '';
                    let displayValue = value;
                    let titleText = `${header}: ${rawValue}`;
                    
                    // Add source indicator to title
                    if (INDEX_HEADERS.includes(header) && header !== 'Lot No') {
                      titleText += `\n(Source: ${hasIndexData ? 'Index Sheet' : 'Default/Empty'})`;
                    } else if (JOB_ORDER_HEADERS.includes(header) || 
                             header === 'Season' || 
                             header === 'Section' || 
                             header === 'Design Work' ||
                             header === 'Party') {
                      titleText += `\n(Source: JobOrder Sheet)`;
                    } else if (header === 'Pkg Sup' || header === 'Pkg Date' || 
                             header === 'WIP Pkg' || header === 'Pkg Comp') {
                      titleText += `\n(Source: ${hasIssuesData ? 'Issues Sheet' : 'No Data'})`;
                    }
                    
                    // Style specific columns
                    if (header === 'Party') {
                      cellClass = 'party-name-cell';
                      displayValue = showPartyInitials ? value : (row['Party_Full'] || rawValue || value);
                      displayValue = displayValue ? ` ${displayValue}` : displayValue;
                      titleText = `${header}: ${row['Party_Full'] || rawValue || value}`;
                    } else if (header === 'PCS') {
                      cellClass = 'numeric-cell';
                      displayValue = typeof value === 'number' ? value.toLocaleString() : value;
                    }
                    else if (header === 'Cut Date' || header === 'Stit Date' || header === 'Pkg Date' || 
                             header === 'Job Date' || header === 'Emb/Print Issue' || header === 'Emb/print Comp' || 
                             header === 'Comp Stit' || header === 'Pkg Comp') {
                      cellClass = 'date-cell';
                      if (value && value !== '-') {
                        displayValue = ` ${value}`;
                      }
                    } else if (header === 'Stit Sup') {
                      cellClass = 'supervisor-cell';
                    } else if (header === 'Section') {
                      cellClass = 'section-cell';
                      displayValue = value ? ` ${value}` : value;
                      const fullSection = row['Section_Full'] || rawValue;
                      titleText = `${header}: ${fullSection}`;
                    } else if (header === 'Season') {
                      cellClass = 'season-cell';
                      displayValue = value ? ` ${value}` : value;
                      const fullSeason = row['Season_Full'] || rawValue;
                      titleText = `${header}: ${fullSeason}`;
                    } else if (header === 'Design Work') {
                      cellClass = `direct-stitching ${(rawValue || value || '').toLowerCase()}`;
                      const displayValueRaw = rawValue || value || '';
                      
                      // Check ONLY JobOrder sheet columns
                      const embFromJobOrder = row['Emb'] || '';
                      const printingFromJobOrder = row['Printing'] || '';
                      
                      // Clean the values
                      const embClean = embFromJobOrder ? String(embFromJobOrder).trim() : '';
                      const printingClean = printingFromJobOrder ? String(printingFromJobOrder).trim() : '';
                      
                      // Check if there's valid data in JobOrder (not empty, not "NA")
                      const hasEmbFromJobOrder = embClean !== '' && 
                                                embClean.toUpperCase() !== 'NA' && 
                                                embClean.toUpperCase() !== 'N/A' &&
                                                embClean !== '-';
                      
                      const hasPrintingFromJobOrder = printingClean !== '' && 
                                                     printingClean.toUpperCase() !== 'NA' && 
                                                     printingClean.toUpperCase() !== 'N/A' &&
                                                     printingClean !== '-';
                      
                      // Check if we should override "NO"
                      const shouldOverride = (displayValueRaw === 'NO' || displayValueRaw === 'N' || displayValueRaw === 'no') && 
                                            (hasEmbFromJobOrder || hasPrintingFromJobOrder);
                      
                      if (shouldOverride) {
                        // Determine what to display
                        let displayText = '';
                        
                        const embLower = hasEmbFromJobOrder ? embClean.toLowerCase() : '';
                        const printingLower = hasPrintingFromJobOrder ? printingClean.toLowerCase() : '';
                        
                        // Check if it's just a name/initials
                        const isJustName = (hasEmbFromJobOrder && !embLower.includes('complete') && 
                                           !embLower.includes('done') && !embLower.includes('finished') &&
                                           !embLower.includes('pending') && !embLower.includes('process') &&
                                           !embLower.includes('ongoing') && !embLower.includes('wip')) ||
                                          (hasPrintingFromJobOrder && !printingLower.includes('complete') && 
                                           !printingLower.includes('done') && !printingLower.includes('finished') &&
                                           !printingLower.includes('pending') && !printingLower.includes('process') &&
                                           !printingLower.includes('ongoing') && !printingLower.includes('wip'));
                        
                        if (isJustName) {
                          // For names/initials like GOURAV, BHM
                          if (hasEmbFromJobOrder && hasPrintingFromJobOrder) {
                            displayText = 'Emb+Prt';
                            cellClass += ' emb-print-both';
                          } else if (hasEmbFromJobOrder) {
                            displayText = 'Emb';
                            cellClass += ' emb-only';
                          } else if (hasPrintingFromJobOrder) {
                            displayText = 'Print';
                            cellClass += ' printing-only';
                          }
                        } else {
                          // Check status words
                          const isComplete = (embLower.includes('complete') || embLower.includes('done') || 
                                             embLower.includes('finished') || embLower.includes('ready'));
                          
                          const isPending = (embLower.includes('pending') || embLower.includes('process') || 
                                            embLower.includes('ongoing') || embLower.includes('wip'));
                          
                          if (isComplete) {
                            displayText = 'Emb/Print';
                            cellClass += ' emb-print-complete';
                          }
                          else if (isPending) {
                            displayText = 'Emb/Prt Pending';
                            cellClass += ' emb-print-pending';
                          }
                          else {
                            // Unknown status
                            if (hasEmbFromJobOrder && hasPrintingFromJobOrder) {
                              displayText = 'Emb+Prt';
                              cellClass += ' emb-print-both';
                            } else if (hasEmbFromJobOrder) {
                              displayText = 'Emb';
                              cellClass += ' emb-only';
                            } else if (hasPrintingFromJobOrder) {
                              displayText = 'Print';
                              cellClass += ' printing-only';
                            }
                          }
                        }
                        
                        displayValue = ` ${displayText}`;
                        
                        // Update tooltip to show details
                        let tooltipDetails = `${header}: ${displayText}\n`;
                        tooltipDetails += `(Original D.Stitching: ${displayValueRaw})\n`;
                        tooltipDetails += `Data from JobOrder sheet:\n`;
                        
                        if (hasEmbFromJobOrder) tooltipDetails += `• Emb: "${embFromJobOrder}"\n`;
                        if (hasPrintingFromJobOrder) tooltipDetails += `• Printing: "${printingFromJobOrder}"`;
                        
                        titleText = tooltipDetails;
                      } else {
                        // Normal handling for YES or unknown
                        displayValue = displayValueRaw === 'YES' ? ' Yes' : 
                                      displayValueRaw === 'Y' ? ' Yes' : 
                                      displayValueRaw === 'yes' ? ' Yes' :
                                      displayValueRaw === 'NO' ? ' No' : 
                                      displayValueRaw === 'N' ? ' No' : 
                                      displayValueRaw === 'no' ? ' No' : 
                                      value || 'N/A';
                      }
                    } else if (header === 'Emb/Print Issue') {
                      cellClass = 'date-cell issue-date';
                      if (value) {
                        displayValue = ` ${value}`;
                        titleText = `First Issue Date: ${value}\n(From first challan in Challan History)`;
                      } else {
                        displayValue = '';
                      }
                    } else if (header === 'Emb/print Comp') {
                      cellClass = 'date-cell complete-date';
                      if (value) {
                        displayValue = ` ${value}`;
                        titleText = `Completion Date: ${value}\n(All challans completed)`;
                      } else {
                        displayValue = '';
                      }
                    } else if (header === 'WIP Stit') {
                      cellClass = 'wip-status-cell';
                      const compStitValue = (row['Comp Stit'] || '').toString().trim();
                      const isStitchingDone = compStitValue && 
                                             compStitValue !== '-' && 
                                             compStitValue.toLowerCase() !== 'invalid date' && 
                                             !compStitValue.toLowerCase().includes('pending');

                      if (isStitchingDone) {
                        displayValue = 'Stitching Done';
                        cellClass += ' wip-complete';
                      } else if (value) {
                        displayValue = value;
                        const lowerValue = value.toLowerCase();
                        if (lowerValue.includes('pending')) {
                          cellClass += ' wip-pending';
                        } else if (lowerValue.includes('working')) {
                          cellClass += ' wip-working';
                        } else if (lowerValue.includes('complete') || lowerValue.includes('done')) {
                          cellClass += ' wip-complete';
                        }
                      }
                    } else if (header === 'Comp Stit') {
                      cellClass = 'completed-status-cell date-cell';
                      if (value) {
                        displayValue = ` ${value}`;
                      }
                    } else if (header === 'WIP Pkg') {
                      cellClass = 'wip-packing-cell';
                      const wipPkgParsed = row['WIP Pkg._parsed'];
                      
                      if (wipPkgParsed && wipPkgParsed.displayValue) {
                        const lines = wipPkgParsed.displayValue.split('\n');
                        
                        return (
                          <td 
                            key={`${rowIndex}-${colIndex}`} 
                            className={cellClass}
                            title={`Latest Packing Update\n${wipPkgParsed.displayValue}`}
                          >
                            <div className="wip-packing-display">
                              {lines.map((line, lineIndex) => (
                                <div 
                                  key={lineIndex} 
                                  className={`wip-packing-line ${lineIndex === 0 ? 'date-line' : 'remarks-line'}`}
                                >
                                  {line}
                                </div>
                              ))}
                            </div>
                          </td>
                        );
                      }
                    } else if (header === 'Pkg Comp') {
                      cellClass = 'packing-complete-cell date-cell';
                      if (value) {
                        displayValue = ` ${value}`;
                        titleText = `Packing Complete Date: ${value}`;
                      }
                    } else if (header === 'Item') {
                      cellClass = 'garment-type-cell';
                    } else if (header === 'Style') {
                      cellClass = 'style-cell';
                    } else if (header === 'Brand') {
                      cellClass = 'brand-cell';
                    } else if (header === 'Pkg Sup') {
                      cellClass = 'packing-supervisor-cell';
                      displayValue = value ? ` ${value}` : value;
                    } else if (header === 'Lot No') {
                      cellClass = 'lot-no-cell';
                    } else if (header === 'Stit Sup') {
                      cellClass = 'supervisor-cell';
                    }
                    // Days columns styling
                    else if (header === 'Cut Days' || header === 'Emb/Print Days' || 
                             header === 'Stit Days' || header === 'Pkg Days' ||
                             header === 'Cut To Emb/Print' || header === 'Emb/Print To Stit' || 
                             header === 'Stit To Check Pack') {
                      cellClass = 'days-cell';
                      
                      // Color coding based on value
                      if (value && value !== '-') {
                        let daysValue = 0;
                        
                        if (value.startsWith('+')) {
                          daysValue = parseInt(value.substring(1));
                        } else if (value.startsWith('-')) {
                          daysValue = parseInt(value);
                        } else if (!isNaN(parseInt(value))) {
                          daysValue = parseInt(value);
                        }
                        
                        if (!isNaN(daysValue)) {
                          if (daysValue < 0) {
                            cellClass += ' days-negative';
                          } else if (daysValue === 0) {
                            cellClass += ' days-zero';
                          } else {
                            cellClass += ' days-positive';
                            
                            // Different color thresholds for different day types
                            if (header === 'Cut Days') {
                              if (daysValue > 2) cellClass += ' days-critical';
                              else if (daysValue > 0) cellClass += ' days-warning';
                              else cellClass += ' days-good';
                            } else if (header === 'Emb/Print Days') {
                              if (daysValue > 10) cellClass += ' days-critical';
                              else if (daysValue > 5) cellClass += ' days-warning';
                              else cellClass += ' days-good';
                            } else if (header === 'Stit Days') {
                              if (daysValue > 10) cellClass += ' days-critical';
                              else if (daysValue > 6) cellClass += ' days-warning';
                              else cellClass += ' days-good';
                            } else if (header === 'Pkg Days') {
                              if (daysValue > 2) cellClass += ' days-critical';
                              else if (daysValue > 0) cellClass += ' days-warning';
                              else cellClass += ' days-good';
                            } else if (header === 'Cut To Emb/Print' || 
                                       header === 'Emb/Print To Stit' || 
                                       header === 'Stit To Check Pack') {
                              if (daysValue > 3) cellClass += ' days-critical';
                              else if (daysValue > 1) cellClass += ' days-warning';
                              else cellClass += ' days-good';
                            }
                          }
                        }
                      }
                      
                      // Add calculation formula to tooltip
                      if (header === 'Cut Days') {
                        titleText = `Cut Days = Cut Date - Job Date\n${row['Cut Date']} - ${row['Job Date']} = ${value}`;
                      } else if (header === 'Emb/Print Days') {
                        titleText = `Emb/Print Days = Emb/print Comp - Emb/Print Issue\n${row['Emb/print Comp']} - ${row['Emb/Print Issue']} = ${value}`;
                      } else if (header === 'Stit Days') {
                        titleText = `Stit Days = Comp Stit - Stit Date\n${row['Comp Stit']} - ${row['Stit Date']} = ${value}`;
                      } else if (header === 'Pkg Days') {
                        titleText = `Pkg Days = Pkg Comp - Pkg Date\n${row['Pkg Comp']} - ${row['Pkg Date']} = ${value}`;
                      } else if (header === 'Cut To Emb/Print') {
                        titleText = `Cut To Emb/Print = Emb/Print Issue - Cut Date\n${row['Emb/Print Issue']} - ${row['Cut Date']} = ${value}`;
                      } else if (header === 'Emb/Print To Stit') {
                        titleText = `Emb/Print To Stit = Stit Date - Emb/print Comp\n${row['Stit Date']} - ${row['Emb/print Comp']} = ${value}`;
                      } else if (header === 'Stit To Check Pack') {
                        titleText = `Stit To Check Pack = Pkg Date - Comp Stit\n${row['Pkg Date']} - ${row['Comp Stit']} = ${value}`;
                      }
                    }
                    
                    // Add missing-data class for Index fields without data
                    if (INDEX_HEADERS.includes(header) && !hasIndexData && !value) {
                      cellClass += ' missing-index-data';
                    }
                    
                    // Add missing-data class for Issues fields without data
                    if ((header === 'Pkg Sup' || header === 'Pkg Date' || 
                         header === 'WIP Pkg' || header === 'Pkg Comp') && 
                        !hasIssuesData && !value) {
                      cellClass += ' missing-issues-data';
                    }
                    
                    return (
                      <td 
                        key={`${rowIndex}-${colIndex}`} 
                        className={cellClass}
                        title={titleText}
                        onClick={
                          header === 'Emb/Print Issue' && challanInfo 
                            ? () => setSelectedChallan(challanInfo) 
                            : header === 'WIP Stit' && wipInfo?.fullData?.length > 0
                            ? () => setSelectedWIP(wipInfo)
                            : header === 'Comp Stit' && completedInfo?.fullData?.length > 0
                            ? () => setSelectedCompleted(completedInfo)
                            : undefined
                        }
                        style={
                          header === 'Emb/Print Issue' || 
                          (header === 'WIP Stit' && wipInfo?.fullData?.length > 0) ||
                          (header === 'Comp Stit' && completedInfo?.fullData?.length > 0)
                            ? { cursor: 'pointer' } 
                            : {}
                        }
                      >
                        {header === 'Emb/Print Issue' && challanInfo ? (
                          <div className="challan-display">
                            <div className="challan-number-date">
                              <div className="challan-line challan-number">
                                {challanInfo.display.split('\n')[0]}
                              </div>
                              <div className="challan-line challan-date">
                                {challanInfo.display.split('\n')[1] || ''}
                              </div>
                            </div>
                            {challanInfo.hasMultiple && (
                              <div className="challan-count-badge" title={`${challanInfo.count} challans`}>
                                {challanInfo.count}
                              </div>
                            )}
                            {(challanInfo.status === 'Emb Done' || challanInfo.status === 'Printing Done') && (
                              <div className="status-icon-done" title="All Completed"></div>
                            )}
                            {(challanInfo.status === 'Emb Pending' || challanInfo.status === 'Printing Pending') && (
                              <div className="status-icon-pending" title="Some challans pending"></div>
                            )}
                            {challanInfo.hasMultiple && challanInfo.allCompleted && (
                              <div className="status-multiple-done" title="All challans completed"></div>
                            )}
                            {challanInfo.hasMultiple && challanInfo.hasPending && (
                              <div className="status-multiple-pending" title="Some challans pending"></div>
                            )}
                            {challanInfo.status === 'Text Data' && (
                              <div className="status-icon-text" title="Text Data"></div>
                            )}
                            {challanInfo.status === 'Error' && (
                              <div className="status-icon-error" title="Error parsing data"></div>
                            )}
                            {!hasIndexData && (
                              <div className="source-indicator-small" title="No matching data in Index sheet">
                                ⓘ
                              </div>
                            )}
                          </div>
                        ) : header === 'Emb/Print Issue' ? (
                          <div className="challan-display">
                            <div className="challan-number-date">
                              <div className="challan-line challan-number">No Challan</div>
                            </div>
                            {!hasIndexData && (
                              <div className="source-indicator-small" title="No matching data in Index sheet">
                                ⓘ
                              </div>
                            )}
                          </div>
                        ) : header === 'Image' ? (
                          <div className="op-image-cell">
                            {row['Image'] ? (
                              <img 
                                src={row['Image']} 
                                alt="Lot Preview" 
                                className="op-table-thumbnail" 
                                referrerPolicy="no-referrer"
                                onClick={() => setViewImageSrc(row['Image'])}
                                onError={(e) => {
                                  e.target.style.display = 'none';
                                  if (e.target.nextSibling) {
                                    e.target.nextSibling.style.display = 'inline';
                                  }
                                }}
                              />
                            ) : null}
                            <span 
                              className="op-no-image-text" 
                              style={{ display: row['Image'] ? 'none' : 'inline' }}
                            >
                              No Image
                            </span>
                          </div>
                        ) : header === 'WIP Stit' && wipInfo?.fullData?.length > 0 ? (
                          <div className="wip-display">
                            {displayValue}
                          </div>
                        ) : header === 'Comp Stit' && completedInfo?.fullData?.length > 0 ? (
                          <div className="completed-display">
                            {displayValue}
                          </div>
                        ) : (
                          <>
                            {displayValue}
                            {INDEX_HEADERS.includes(header) && !hasIndexData && !value && (
                              <div className="missing-data-indicator" title="No data in Index sheet">
                                —
                              </div>
                            )}
                            {(header === 'Pkg Sup' || header === 'Pkg Date' || 
                              header === 'WIP Pkg' || header === 'Pkg Comp') && 
                              !hasIssuesData && !value && (
                              <div className="missing-issues-indicator" title="No data in Issues sheet">
                                —
                              </div>
                            )}
                          </>
                        )}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      
      {filteredData.length === 0 && data.length > 0 && (
        <div className="op-no-filter-match">
          <div className="no-match-content">
            <span className="no-match-icon">🔍</span>
            <div>
              <h4>No lots match your filters</h4>
              <p>Try adjusting your filters or search term</p>
            </div>
          </div>
        </div>
      )}
    </>
  )}
</div>

      {/* Export Status */}
      {exportLoading && (
        <div className="export-loading-overlay">
          <div className="export-spinner"></div>
          <p>Preparing export file...</p>
        </div>
      )}

      {/* Challan Details Modal */}
      {selectedChallan && (
        <div className="challan-modal-overlay" onClick={() => setSelectedChallan(null)}>
          <div className="challan-modal-content" onClick={(e) => e.stopPropagation()}>
            <div className="challan-modal-header">
              <h3>Challan Details</h3>
              <button 
                className="challan-modal-close"
                onClick={() => setSelectedChallan(null)}
              >
                ✕
              </button>
            </div>
            <div className="challan-modal-body">
              <div className="challan-summary">
                <p><strong>Display:</strong> {selectedChallan.display}</p>
                <p><strong>Status:</strong> 
                  <span className={`emb-status-badge status-${selectedChallan.status?.replace(/\s+/g, '-').toLowerCase()}`}>
                    {selectedChallan.status === 'Emb Done' && '✅ '}
                    {selectedChallan.status === 'Emb Pending' && '⏳ '}
                    {selectedChallan.status === 'Printing Done' && '✅ '}
                    {selectedChallan.status === 'Printing Pending' && '⏳ '}
                    {selectedChallan.status}
                  </span>
                </p>
                <p><strong>Type:</strong> {selectedChallan.type}</p>
                {selectedChallan.dateOfIssue && (
                  <p><strong>First Issue Date:</strong> {selectedChallan.dateOfIssue}</p>
                )}
                {selectedChallan.embUpdatedAt && (
                  <p><strong>Last Completed Date:</strong> {selectedChallan.embUpdatedAt}</p>
                )}
                {selectedChallan.hasMultiple && (
                  <>
                    <p><strong>Multiple Challans:</strong> {selectedChallan.count} challans</p>
                    <p><strong>Completion Status:</strong> 
                      {selectedChallan.allCompleted ? '✅ All Completed' : '⚠️ Some Pending'}
                    </p>
                  </>
                )}
                {selectedChallan.totalQty > 0 && (
                  <p><strong>Total Quantity:</strong> {selectedChallan.totalQty.toLocaleString()}</p>
                )}
              </div>
              
              {selectedChallan.challans && (
                <div className="challan-list">
                  <h4>Challan Items ({selectedChallan.count})</h4>
                  {selectedChallan.challans.map((challan, index) => (
                    <div key={index} className="challan-item">
                      <div className="challan-item-header">
                        <span className="challan-number">{challan.number || 'No Number'}</span>
                        <span className="challan-date">{challan.date || formatDate(challan.receivedDate || challan.createdAt)}</span>
                      </div>
                      <div className="challan-item-details">
                        <p><strong>Total Qty:</strong> {challan.totalQty || 0}</p>
                        {challan.embUpdatedAt && (
                          <p><strong>Emb Updated:</strong> {formatDate(challan.embUpdatedAt)}</p>
                        )}
                        {challan.embCompleted !== undefined && (
                          <p><strong>Emb Completed:</strong> {challan.embCompleted ? ' Yes' : ' No'}</p>
                        )}
                        {challan.items && Array.isArray(challan.items) && challan.items.length > 0 && (
                          <div className="challan-shades">
                            <strong>Shades:</strong>
                            <ul>
                              {challan.items.map((item, idx) => (
                                <li key={idx}>
                                  {item.shade || 'No Shade'}: {item.qty || 0}
                                </li>
                              ))}
                            </ul>
                          </div>
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              )}
              
              {selectedChallan.wipData && (
                <div className="wip-data-section">
                  <h4>WIP/Completed Status</h4>
                  {selectedChallan.wipData.map((wipItem, index) => (
                    <div key={index} className="wip-item">
                      <p><strong>Status:</strong> {wipItem.status}</p>
                      <p><strong>Remarks:</strong> {wipItem.remarks || 'None'}</p>
                      <p><strong>Timestamp:</strong> {formatDate(wipItem.timestamp)}</p>
                    </div>
                  ))}
                </div>
              )}
              
              <div className="challan-raw-data">
                <details>
                  <summary>View Raw Data</summary>
                  <pre>{selectedChallan.raw}</pre>
                </details>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* WIP Details Modal */}
      {selectedWIP && (
        <div className="challan-modal-overlay" onClick={() => setSelectedWIP(null)}>
          <div className="challan-modal-content" onClick={(e) => e.stopPropagation()}>
            <div className="challan-modal-header">
              <h3>WIP Status History</h3>
              <button 
                className="challan-modal-close"
                onClick={() => setSelectedWIP(null)}
              >
                ✕
              </button>
            </div>
            <div className="challan-modal-body">
              <div className="wip-summary">
                <p><strong>Display Value:</strong> {selectedWIP.displayValue}</p>
                <p><strong>Total Entries:</strong> {selectedWIP.count}</p>
              </div>
              
              {selectedWIP.fullData && selectedWIP.fullData.length > 0 && (
                <div className="wip-full-history">
                  <h4>Complete WIP History ({selectedWIP.fullData.length} entries)</h4>
                  {selectedWIP.fullData.map((wipItem, index) => (
                    <div key={index} className="wip-history-item">
                      <div className="wip-history-header">
                        <span className="wip-history-status">{wipItem.status || 'No Status'}</span>
                        <span className="wip-history-time">
                          {formatDate(wipItem.timestamp)}
                          {index === 0 && <span className="wip-latest-badge">LATEST</span>}
                        </span>
                      </div>
                      {wipItem.remarks && (
                        <div className="wip-history-remarks">
                          <strong>Remarks:</strong> {wipItem.remarks}
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              )}
              
              <div className="challan-raw-data">
                <details>
                  <summary>View Raw Data</summary>
                  <pre>{selectedWIP.raw}</pre>
                </details>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Completed Status Details Modal */}
      {selectedCompleted && (
        <div className="challan-modal-overlay" onClick={() => setSelectedCompleted(null)}>
          <div className="challan-modal-content" onClick={(e) => e.stopPropagation()}>
            <div className="challan-modal-header">
              <h3>Completed Status History</h3>
              <button 
                className="challan-modal-close"
                onClick={() => setSelectedCompleted(null)}
              >
                ✕
              </button>
            </div>
            <div className="challan-modal-body">
              <div className="wip-summary">
                <p><strong>Completion Date:</strong> {selectedCompleted.displayValue}</p>
                <p><strong>Total Entries:</strong> {selectedCompleted.count}</p>
              </div>
              
              {selectedCompleted.fullData && selectedCompleted.fullData.length > 0 && (
                <div className="wip-full-history">
                  <h4>Complete History ({selectedCompleted.fullData.length} entries)</h4>
                  {selectedCompleted.fullData.map((item, index) => (
                    <div key={index} className="wip-history-item">
                      <div className="wip-history-header">
                        <span className="wip-history-status">{item.status || 'No Status'}</span>
                        <span className="wip-history-time">
                          {formatDate(item.timestamp)}
                          {index === 0 && <span className="wip-latest-badge">LATEST</span>}
                        </span>
                      </div>
                      {item.remarks && (
                        <div className="wip-history-remarks">
                          <strong>Remarks:</strong> {item.remarks}
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              )}
              
              <div className="challan-raw-data">
                <details>
                  <summary>View Raw Data</summary>
                  <pre>{selectedCompleted.raw}</pre>
                </details>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Image Lightbox Modal */}
      {viewImageSrc && (
        <div className="challan-modal-overlay" onClick={() => setViewImageSrc(null)}>
          <div className="op-image-modal-content" onClick={(e) => e.stopPropagation()}>
            <button 
              className="challan-modal-close" 
              onClick={() => setViewImageSrc(null)}
            >
              ✕
            </button>
            <img src={viewImageSrc} alt="Full Preview" className="op-image-modal-img" referrerPolicy="no-referrer" />
          </div>
        </div>
      )}
    </div>
  );
};

export default OverallCuttingToPacking;