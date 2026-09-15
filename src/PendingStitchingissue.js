import React, { useState, useEffect, useRef } from 'react';
import { GOOGLE_API_KEY, SPREADSHEET_IDS, fetchSheetDataFromBackend } from './config';
import { fetchRemarksForTab, saveRemarkForLot } from './embPrintRemarksService';
import axios from 'axios';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

const MultiSelectDropdown = ({ options, selectedValues, onChange, placeholder }) => {
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

  const handleSelectAll = () => {
    onChange([...options]);
  };

  const handleClearAll = () => {
    onChange([]);
  };

  const filteredOptions = options.filter(option =>
    option ? option.toString().toLowerCase().includes(searchTerm.toLowerCase()) : false
  );

  return (
    <div ref={containerRef} style={{ position: 'relative', width: '100%' }}>
      <div
        onClick={() => setIsOpen(!isOpen)}
        style={{
          padding: '0.625rem 0.75rem',
          borderRadius: '6px',
          background: 'white',
          border: '1px solid #cbd5e1',
          color: '#334155',
          fontSize: '0.875rem',
          cursor: 'pointer',
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          minHeight: '38px',
          boxSizing: 'border-box',
          userSelect: 'none',
        }}
      >
        <span style={{
          whiteSpace: 'nowrap',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          maxWidth: '85%'
        }}>
          {selectedValues.length === 0
            ? placeholder
            : selectedValues.length === options.length
              ? `All Selected (${selectedValues.length})`
              : `${selectedValues.length} Selected`}
        </span>
        <span style={{ fontSize: '0.6rem', color: '#64748b' }}>
          {isOpen ? '▲' : '▼'}
        </span>
      </div>
      {isOpen && (
        <div style={{
          position: 'absolute',
          top: '105%',
          left: 0,
          right: 0,
          background: 'white',
          border: '1px solid #cbd5e1',
          borderRadius: '6px',
          boxShadow: '0 4px 12px rgba(0, 0, 0, 0.1)',
          zIndex: 1000,
          maxHeight: '250px',
          display: 'flex',
          flexDirection: 'column',
        }}>
          {options.length > 5 && (
            <div style={{ padding: '0.5rem', borderBottom: '1px solid #f1f5f9' }}>
              <input
                type="text"
                placeholder="Search..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                style={{
                  width: '100%',
                  padding: '0.375rem 0.5rem',
                  borderRadius: '4px',
                  border: '1px solid #e2e8f0',
                  fontSize: '0.8125rem',
                  outline: 'none',
                  boxSizing: 'border-box',
                }}
              />
            </div>
          )}
          <div style={{
            display: 'flex',
            justifyContent: 'space-between',
            padding: '0.375rem 0.5rem',
            borderBottom: '1px solid #f1f5f9',
            fontSize: '0.75rem',
            background: '#f8fafc',
            userSelect: 'none',
          }}>
            <span
              onClick={handleSelectAll}
              style={{ color: '#2563eb', cursor: 'pointer', fontWeight: 500 }}
            >
              Select All
            </span>
            <span
              onClick={handleClearAll}
              style={{ color: '#dc2626', cursor: 'pointer', fontWeight: 500 }}
            >
              Clear
            </span>
          </div>
          <div style={{
            overflowY: 'auto',
            flex: 1,
            padding: '0.25rem 0'
          }}>
            {filteredOptions.length === 0 ? (
              <div style={{ padding: '0.5rem 0.75rem', fontSize: '0.8125rem', color: '#64748b', textAlign: 'center' }}>
                No options found
              </div>
            ) : (
              filteredOptions.map(option => {
                const isChecked = selectedValues.includes(option);
                return (
                  <label
                    key={option}
                    style={{
                      display: 'flex',
                      alignItems: 'center',
                      padding: '0.5rem 0.75rem',
                      fontSize: '0.8125rem',
                      cursor: 'pointer',
                      background: isChecked ? '#f1f5f9' : 'transparent',
                      userSelect: 'none',
                    }}
                    onMouseEnter={(e) => e.currentTarget.style.backgroundColor = '#f8fafc'}
                    onMouseLeave={(e) => e.currentTarget.style.backgroundColor = isChecked ? '#f1f5f9' : 'transparent'}
                  >
                    <input
                      type="checkbox"
                      checked={isChecked}
                      onChange={() => handleToggleOption(option)}
                      style={{ marginRight: '0.5rem', cursor: 'pointer' }}
                    />
                    <span style={{
                      whiteSpace: 'nowrap',
                      overflow: 'hidden',
                      textOverflow: 'ellipsis'
                    }}>
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

const getDirectImageUrl = (url) => {
  if (!url) return '';
  const trimmed = url.toString().trim();
  if (!trimmed) return '';

  if (!trimmed.includes('drive.google.com')) {
    return trimmed;
  }

  let fileId = '';
  const dMatch = trimmed.match(/id=([^&]+)/);
  if (dMatch && dMatch[1]) {
    fileId = dMatch[1];
  } else {
    const fileMatch = trimmed.match(/\/file\/d\/([^/]+)/);
    if (fileMatch && fileMatch[1]) {
      fileId = fileMatch[1];
    }
  }

  if (fileId) {
    return `https://lh3.googleusercontent.com/d/${fileId}`;
  }

  return trimmed;
};

const getBase64ImageFromUrl = async (imageUrl) => {
  if (!imageUrl || imageUrl === 'N/A') return null;
  try {
    const res = await fetch(imageUrl, { referrerPolicy: 'no-referrer' });
    const blob = await res.blob();
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  } catch (err) {
    console.error('Failed to convert image to base64:', err);
    return null;
  }
};

const PendingIssuetoStitching = () => {
  const [data, setData] = useState([]);
  const [filteredData, setFilteredData] = useState([]);
  const [matrixData, setMatrixData] = useState({});
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [viewImageSrc, setViewImageSrc] = useState(null);
  const [stats, setStats] = useState({
    totalLots: 0,
    totalPieces: 0,
    brands: new Set(),
    seasons: new Set(),
    colorPendingLots: 0,
    repeatedLots: 0,
    statusCounts: {
      direct: 0,
      ready: 0,
      printing: 0,
      embroidery: 0,
      pending: 0
    }
  });

  const [filters, setFilters] = useState({
    fabric: [],
    garmentType: [],
    style: [],
    brand: [],
    section: [],
    season: [],
    party: [],
    directStitch: '',
    colorPending: '',
    remarksStatus: '',
    status: [],
    search: ''
  });

  const [filterOptions, setFilterOptions] = useState({
    fabric: new Set(),
    garmentType: new Set(),
    style: new Set(),
    brand: new Set(),
    section: new Set(),
    season: new Set(),
    party: new Set(),
    directStitch: new Set(['Yes', 'No']),
    colorPending: new Set(['Yes', 'No']),
    status: new Set(['Direct', 'Ready for Stitching', 'Printing Working', 'Embroidery Working', 'Pending'])
  });

  // Color pending states
  const [colorPendingLots, setColorPendingLots] = useState({});
  const [lotPriorities, setLotPriorities] = useState({});

  // Remarks States & Realtime Synchronization
  const [remarksMap, setRemarksMap] = useState({});
  const [remarksModalOpen, setRemarksModalOpen] = useState(false);
  const [selectedRemarksLot, setSelectedRemarksLot] = useState(null);
  const [newRemarkInputText, setNewRemarkInputText] = useState('');
  const [savingRemark, setSavingRemark] = useState(false);

  // Fetch Remarks from Google Sheets & Subscribe to updates
  useEffect(() => {
    fetchRemarksForTab('PENDING_STITCHING').then(map => {
      if (map && typeof map === 'object') {
        setRemarksMap(map);
      }
    });

    const handleRemarkUpdated = (e) => {
      if (e.detail && (e.detail.tabType === 'PENDING_STITCHING' || !e.detail.tabType)) {
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
      const lotNumber = (selectedRemarksLot['Lot Number'] || selectedRemarksLot.lotNumber)?.toString().trim();
      const updatedHistory = await saveRemarkForLot({
        tabType: 'PENDING_STITCHING',
        lotNumber: lotNumber,
        partyName: selectedRemarksLot['Party Name'] || selectedRemarksLot.partyName || '',
        fabric: selectedRemarksLot['Fabric'] || selectedRemarksLot.fabric || '',
        style: selectedRemarksLot['Style'] || selectedRemarksLot.style || selectedRemarksLot['Garment Type'] || '',
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

  const tableRef = useRef();

  const SPREADSHEET_ID = SPREADSHEET_IDS.MAIN;
  const API_KEY = GOOGLE_API_KEY;
  const RANGE = 'Index!A:AG';
  const CUTTING_SHEET_RANGE = 'Cutting!A:Z';
  const INDEX_RANGE = 'Index!A:AG';

  const columnsConfig = [
    { key: 'srNo', label: '#', align: 'center', minWidth: '45px' },
    { key: 'Lot Number', label: 'LOT #', align: 'center', minWidth: '95px' },
    { key: 'Garment Type', label: 'GARMENT TYPE', align: 'center', minWidth: '140px' },
    { key: 'Style', label: 'STYLE', align: 'center', minWidth: '130px' },
    { key: 'Fabric', label: 'FABRIC', align: 'center', minWidth: '140px' },
    { key: 'Brand', label: 'BRAND', align: 'center', minWidth: '120px' },
    { key: 'Total Pcs', label: 'TOTAL PCS', align: 'center', minWidth: '100px' },
    { key: 'Section', label: 'SECTION', align: 'center', minWidth: '95px' },
    { key: 'Season', label: 'SEASON', align: 'center', minWidth: '95px' },
    { key: 'Party Name', label: 'PARTY NAME', align: 'center', minWidth: '140px' },
    { key: 'Direct Stitching', label: 'DIRECT', align: 'center', minWidth: '85px' },
    { key: 'Image', label: 'IMAGE', align: 'center', minWidth: '85px' },
    { key: 'Color Status', label: 'COLOR STATUS', align: 'center', minWidth: '140px' },
    { key: 'Status', label: 'STATUS', align: 'center', minWidth: '145px' },
    { key: 'Remarks', label: '💬 REMARKS', align: 'center', minWidth: '220px' }
  ];

  // Updated columns
  const targetColumns = [
    'Lot Number',
    'Fabric',
    'Garment Type',
    'Style',
    'Brand',
    'Total Pcs',
    'Section',
    'Season',
    'Party Name',
    'Direct Stitching',
    'Image',
    'Remarks'
  ];

  // Column headers display names
  const columnDisplayNames = {
    'Lot Number': 'Lot #',
    'Fabric': 'Fabric',
    'Garment Type': 'Garment Type',
    'Style': 'Style',
    'Brand': 'Brand',
    'Season': 'Season',
    'Direct Stitching': 'Direct',
    'Total Pcs': 'Total Pcs',
    'Status': 'Status',
    'Image': 'Image'
  };

  // Normalization functions
  const normalizeText = (text) => {
    if (!text || text === 'N/A') return text;

    return text
      .toString()
      .toLowerCase()
      .trim()
      .replace(/\s+/g, ' ')
      .replace(/[^\w\s]/g, '')
      .replace(/\b(?:rl|polo|polo t shirt|t shirt|tshirt)\b/gi, (match) => {
        const lower = match.toLowerCase();
        if (lower === 'rl') return 'rl';
        if (lower === 'polo' || lower === 'polo t shirt' || lower === 't shirt' || lower === 'tshirt') return 'polo';
        return match;
      });
  };

  const normalizeBrand = (brand) => {
    if (!brand || brand === 'N/A') return brand;

    const normalized = normalizeText(brand);

    const brandMappings = {
      'rl': 'RL',
      'ralph lauren': 'RL',
      'polo ralph lauren': 'RL',
      'tommy hilfiger': 'Tommy Hilfiger',
      'th': 'Tommy Hilfiger',
      'calvin klein': 'Calvin Klein',
      'ck': 'Calvin Klein',
      'nike': 'Nike',
      'adidas': 'Adidas',
      'puma': 'Puma',
      'levis': 'Levis',
      'lee': 'Lee',
      'wrangler': 'Wrangler'
    };

    return brandMappings[normalized] || brand.charAt(0).toUpperCase() + brand.slice(1);
  };

  const normalizeStyle = (style) => {
    if (!style || style === 'N/A') return style;

    const normalized = normalizeText(style);

    const styleMappings = {
      'polo': 'Polo',
      'tshirt': 'T-Shirt',
      't shirt': 'T-Shirt',
      'polo t shirt': 'Polo T-Shirt',
      'shirt': 'Shirt',
      'formal shirt': 'Formal Shirt',
      'casual shirt': 'Casual Shirt',
      'jeans': 'Jeans',
      'trousers': 'Trousers',
      'pants': 'Pants',
      'shorts': 'Shorts',
      'jacket': 'Jacket',
      'hoodie': 'Hoodie',
      'sweatshirt': 'Sweatshirt',
      'sweater': 'Sweater'
    };

    return styleMappings[normalized] || style.charAt(0).toUpperCase() + style.slice(1);
  };

  const normalizeSeason = (season) => {
    if (!season || season === 'N/A') return season;

    const normalized = normalizeText(season);

    const seasonMappings = {
      'ss24': 'SS24',
      'spring summer 24': 'SS24',
      'spring/summer 2024': 'SS24',
      'fw24': 'FW24',
      'fall winter 24': 'FW24',
      'fall/winter 2024': 'FW24',
      'aw24': 'AW24',
      'autumn winter 24': 'AW24',
      'autumn/winter 2024': 'AW24',
      'ss23': 'SS23',
      'fw23': 'FW23',
      'aw23': 'AW23',
      'ss22': 'SS22',
      'fw22': 'FW22',
      'aw22': 'AW22'
    };

    return seasonMappings[normalized] || season.toUpperCase();
  };

  // Color pending helper functions
  const normalizeKey = (s = "") => {
    return String(s || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
  };

  // Get Status - FIXED: If Direct Stitch is "yes", status should be "Direct" regardless of challan history
  const getStatus = (directStitching, challanHistory) => {
    // If Direct Stitch is "yes", it's always "Direct"
    const directStitchingLower = directStitching ? directStitching.toString().toLowerCase().trim() : '';
    if (directStitchingLower === 'yes' || directStitchingLower === 'y') {
      return 'Direct';
    }

    if (!challanHistory || challanHistory === 'N/A' || challanHistory.trim() === '') {
      return 'Pending';
    }

    try {
      const challans = JSON.parse(challanHistory);

      if (!Array.isArray(challans) || challans.length === 0) {
        return 'Pending';
      }

      const latestChallan = challans[0];

      if (latestChallan.embCompleted) {
        return 'Embroidery Working';
      }

      if (latestChallan.number && latestChallan.number.includes('PRINT')) {
        return 'Printing Working';
      }

      if (latestChallan.number && latestChallan.number.includes('EMB')) {
        return 'Embroidery Working';
      }

      return 'Pending';

    } catch (error) {
      console.error('Error parsing challan history:', error);
      return 'Pending';
    }
  };

  // Color pending check functions
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

  const computePendingShades = (windowValues, sizes = [], shades = []) => {
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

    if (sizeColIndices.length === 0) {
      return new Set(shades.map(normalizeKey));
    }

    const shadeStats = new Map();

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

      if (hasAnyData) {
        if (hasPositiveData) {
          shadeStats.set(shadeKey, "found-with-data");
        } else {
          if (!shadeStats.has(shadeKey) || shadeStats.get(shadeKey) === "not-found") {
            shadeStats.set(shadeKey, "found-all-zero");
          }
        }
      } else {
        if (!shadeStats.has(shadeKey)) {
          shadeStats.set(shadeKey, "found-no-data");
        }
      }
    }

    const pendingShadeKeys = new Set();
    const expectedShadeKeys = (shades || []).map(sh => normalizeKey(sh));

    expectedShadeKeys.forEach((shadeKey) => {
      const status = shadeStats.get(shadeKey);

      if (!status || status === "found-no-data") {
        pendingShadeKeys.add(shadeKey);
      }
    });

    return pendingShadeKeys;
  };

  const sliceCuttingMatrix = (bigValues, startRow, numRows) => {
    if (!Array.isArray(bigValues) || bigValues.length === 0) return [];
    if (!(startRow > 0 && numRows > 0)) return [];
    const r0 = Math.max(0, startRow - 1);
    const r1 = Math.min(bigValues.length - 1, r0 + numRows - 1);
    return bigValues.slice(r0, r1 + 1);
  };

  // Check color pending for all lots
  const checkColorPendingForLots = async () => {
    try {
      const [cuttingRes, indexRes] = await Promise.all([
        fetchSheetDataFromBackend(SPREADSHEET_ID, "Cutting!A:Z"),
        fetchSheetDataFromBackend(SPREADSHEET_ID, INDEX_RANGE)
      ]);

      const cuttingRows = cuttingRes.ok ? cuttingRes.values : [];
      const indexRows = indexRes.ok ? indexRes.values : [];

      if (indexRows.length === 0) {
        return {};
      }

      const pendingMap = {};
      const priorityMap = {};

      for (let i = 1; i < indexRows.length; i++) {
        const row = indexRows[i];

        const headers = indexRows[0];
        const headerIndices = {};
        headers.forEach((header, index) => {
          if (header) {
            headerIndices[normalizeKey(header)] = index;
          }
        });

        const lotNumber = (row[headerIndices['lotnumber']] ||
          row[headerIndices['lot no']] ||
          row[headerIndices['lot']] || '').toString().trim();

        if (!lotNumber || lotNumber === '') continue;

        const startRow = parseInt(row[headerIndices['startrow']] || '0', 10);
        const numRows = parseInt(row[headerIndices['numrows']] || '0', 10);

        const sizesStr = row[headerIndices['sizes']] || '';
        const shadesStr = row[headerIndices['shades']] || '';

        const sizes = sizesStr.split(',')
          .map(s => s.trim())
          .filter(Boolean);

        const shades = shadesStr.split(',')
          .map(s => s.trim())
          .filter(Boolean);

        const priority = (row[headerIndices['priority']] ||
          row[headerIndices['prioirty']] ||
          row[headerIndices['special']] || '').toString().trim();

        priorityMap[lotNumber] = priority;

        pendingMap[lotNumber] = {
          pendingColors: [],
          priority: priority,
          isRepeatedLot: priority && priority.toLowerCase().includes('repeated')
        };

        if (shades && shades.length > 0 && startRow > 0 && numRows > 0) {
          const window = sliceCuttingMatrix(cuttingRows, startRow, numRows);

          if (window.length > 0) {
            const pendingShadeKeys = computePendingShades(window, sizes, shades);

            if (pendingShadeKeys.size > 0) {
              const shadeKeyToOriginal = new Map();
              shades.forEach(shade => {
                shadeKeyToOriginal.set(normalizeKey(shade), shade);
              });

              const pendingList = Array.from(pendingShadeKeys)
                .map(key => shadeKeyToOriginal.get(key) || key)
                .filter(Boolean);

              pendingMap[lotNumber].pendingColors = pendingList;
            }
          }
        }
      }

      setColorPendingLots(pendingMap);
      setLotPriorities(priorityMap);

      return pendingMap;

    } catch (err) {
      console.error('Error checking color pending:', err);
      return {};
    }
  };

  useEffect(() => {
    fetchData();
  }, []);

  // Apply filters when data or filters change
  useEffect(() => {
    if (data.length > 0) {
      applyFilters();
    }
  }, [data, filters, remarksMap]);

  const fetchData = async () => {
    try {
      setLoading(true);
      setError(null);

      // Fetch all data in parallel via backend
      const [mainRes, cuttingRes, pendingData] = await Promise.all([
        fetchSheetDataFromBackend(SPREADSHEET_ID, RANGE),
        fetchSheetDataFromBackend(SPREADSHEET_ID, CUTTING_SHEET_RANGE),
        checkColorPendingForLots()
      ]);

      const rows = mainRes.ok ? mainRes.values : [];
      const cuttingRows = cuttingRes.ok ? cuttingRes.values : [];

      if (!rows || rows.length === 0) {
        setData([]);
        setFilteredData([]);
        setStats({
          totalLots: 0,
          totalPieces: 0,
          brands: new Set(),
          seasons: new Set(),
          colorPendingLots: 0,
          repeatedLots: 0,
          statusCounts: {
            direct: 0,
            ready: 0,
            printing: 0,
            embroidery: 0,
            pending: 0
          }
        });
        return;
      }

      const headers = rows[0].map(h => (h || '').toString().trim());

      // Find indices with exact match first, then alias support
      const findHeaderIndex = (patterns, excludePatterns = []) => {
        // First try exact match
        const exactIdx = headers.findIndex(h => {
          if (!h) return false;
          const cleanH = h.toString().toLowerCase().replace(/[^a-z0-9]/g, '');
          return patterns.some(p => {
            const cleanP = p.toLowerCase().replace(/[^a-z0-9]/g, '');
            return cleanH === cleanP;
          });
        });
        if (exactIdx !== -1) return exactIdx;

        // Fallback: substring match with exclusions
        return headers.findIndex(h => {
          if (!h) return false;
          const cleanH = h.toString().toLowerCase().replace(/[^a-z0-9]/g, '');
          const isExcluded = excludePatterns.some(ex => {
            const cleanEx = ex.toLowerCase().replace(/[^a-z0-9]/g, '');
            return cleanH.includes(cleanEx);
          });
          if (isExcluded) return false;

          return patterns.some(p => {
            const cleanP = p.toLowerCase().replace(/[^a-z0-9]/g, '');
            return cleanH.includes(cleanP);
          });
        });
      };

      const lotIndex = findHeaderIndex(['lotnumber', 'lot no', 'lot#', 'lot']);
      const fabricIndex = findHeaderIndex(['fabric', 'material']);
      const garmentTypeIndex = findHeaderIndex(['garmenttype', 'garment', 'item']);
      const styleIndex = findHeaderIndex(['style', 'styledesc', 'stylename']);
      const brandIndex = findHeaderIndex(['brand', 'brandname']);
      const sectionIndex = findHeaderIndex(['mwk', 'm/w/k', 'section', 'gender', 'm w k'], ['garment', 'style', 'fabric', 'party']);
      const seasonIndex = findHeaderIndex(['season', 'seasontype']);
      const partyIndex = findHeaderIndex(['partyname', 'party', 'embparty', 'vendor']);
      const directIndex = findHeaderIndex(['directstitching', 'directstitch', 'direct']);
      const totalPcsIndex = findHeaderIndex(['totalpcs', 'totalpieces', 'totalqty', 'total']);
      const imageIndex = findHeaderIndex(['imageurl', 'image', 'photo', 'picture']);
      const challanHistoryIndex = findHeaderIndex(['challanhistory', 'challan']);
      const supervisorIndex = findHeaderIndex(['supervisor']);

      // Filter and map data
      const filteredData = rows.slice(1)
        .filter(row => {
          // Check supervisor column is empty
          const supervisorEmpty = supervisorIndex === -1 || !row[supervisorIndex] || row[supervisorIndex].trim() === '';
          const lotNo = lotIndex !== -1 ? (row[lotIndex] || '').toString().trim() : '';
          return supervisorEmpty && lotNo !== '';
        })
        .map((row, idx) => {
          const lotNumber = lotIndex !== -1 ? (row[lotIndex] || '').toString().trim() : `Lot-${idx + 1}`;
          const fabric = fabricIndex !== -1 && row[fabricIndex] ? row[fabricIndex].toString().trim() : 'N/A';
          const garmentType = garmentTypeIndex !== -1 && row[garmentTypeIndex] ? row[garmentTypeIndex].toString().trim() : 'N/A';
          const style = styleIndex !== -1 && row[styleIndex] ? normalizeStyle(row[styleIndex].toString().trim()) : 'N/A';
          const brand = brandIndex !== -1 && row[brandIndex] ? normalizeBrand(row[brandIndex].toString().trim()) : 'N/A';
          const section = sectionIndex !== -1 && row[sectionIndex] && row[sectionIndex].toString().trim() !== '' ? row[sectionIndex].toString().trim() : '—';
          const rawSeason = seasonIndex !== -1 && row[seasonIndex] ? row[seasonIndex].toString().trim() : '';
          const season = rawSeason ? normalizeSeason(rawSeason) : 'N/A';
          const directStitching = directIndex !== -1 && row[directIndex] ? row[directIndex].toString().trim() : 'No';

          const directLower = directStitching.toLowerCase();
          const isDirect = directLower === 'yes' || directLower === 'y';

          let partyName = partyIndex !== -1 && row[partyIndex] ? row[partyIndex].toString().trim() : '';
          if (!partyName || partyName === 'N/A' || partyName === '-') {
            partyName = isDirect ? 'Direct Stitching' : '—';
          }

          const rawImage = imageIndex !== -1 && row[imageIndex] ? row[imageIndex].toString().trim() : '';
          const imageUrl = getDirectImageUrl(rawImage);

          const pendingInfo = pendingData[lotNumber] || {};
          const hasColorPending = pendingInfo.pendingColors?.length > 0 || false;
          const pendingColors = pendingInfo.pendingColors || [];
          const pendingColorsText = pendingColors.join(', ');
          const isRepeatedLot = pendingInfo.isRepeatedLot || false;
          const priority = pendingInfo.priority || '';

          const challanHistory = challanHistoryIndex !== -1 ? row[challanHistoryIndex] : '';
          const status = getStatus(directStitching, challanHistory);

          return {
            'Lot Number': lotNumber,
            lotNumber: lotNumber,
            'Fabric': fabric,
            fabric: fabric,
            'Garment Type': garmentType,
            garmentType: garmentType,
            'Style': style,
            style: style,
            'Brand': brand,
            brand: brand,
            'Section': section,
            'M/W/K': section,
            section: section,
            'Season': season,
            season: season,
            'Party Name': partyName,
            'Party': partyName,
            partyName: partyName,
            'Direct Stitching': isDirect ? 'Yes' : 'No',
            directStitching: isDirect ? 'yes' : 'no',
            'Image': imageUrl || 'N/A',
            imageUrl: imageUrl,
            'Total Pcs': totalPcsIndex !== -1 && row[totalPcsIndex] ? row[totalPcsIndex] : 'N/A',
            hasColorPending,
            pendingColors,
            pendingColorsText,
            isRepeatedLot,
            priority,
            status,
            rawRow: row
          };
        });

      setData(filteredData);
      setFilteredData(filteredData);

      // Calculate statistics and filter options
      const brands = new Set();
      const seasons = new Set();
      const fabricOptions = new Set();
      const garmentTypeOptions = new Set();
      const styleOptions = new Set();
      const sectionOptions = new Set();
      const partyOptions = new Set();
      const directStitchOptions = new Set();
      const statusOptions = new Set();

      let colorPendingCount = 0;
      let repeatedLotsCount = 0;
      const statusCounts = {
        direct: 0,
        ready: 0,
        printing: 0,
        embroidery: 0,
        pending: 0
      };

      filteredData.forEach(row => {
        if (row.Brand && row.Brand !== 'N/A' && row.Brand !== '-') brands.add(row.Brand);
        if (row.Season && row.Season !== 'N/A' && row.Season !== '-') seasons.add(row.Season);
        if (row.Fabric && row.Fabric !== 'N/A' && row.Fabric !== '-') fabricOptions.add(row.Fabric);
        if (row['Garment Type'] && row['Garment Type'] !== 'N/A' && row['Garment Type'] !== '-') garmentTypeOptions.add(row['Garment Type']);
        if (row.Style && row.Style !== 'N/A' && row.Style !== '-') styleOptions.add(row.Style);
        if (row.Section && row.Section !== '—' && row.Section !== 'N/A' && row.Section !== '-') sectionOptions.add(row.Section);
        if (row['Party Name'] && row['Party Name'] !== '—' && row['Party Name'] !== 'N/A' && row['Party Name'] !== '-') partyOptions.add(row['Party Name']);
        if (row.status) statusOptions.add(row.status);

        if (row.hasColorPending) colorPendingCount++;
        if (row.isRepeatedLot) repeatedLotsCount++;

        // Count statuses
        if (row.status) {
          switch (row.status) {
            case 'Direct':
              statusCounts.direct++;
              break;
            case 'Ready for Stitching':
              statusCounts.ready++;
              break;
            case 'Printing Working':
              statusCounts.printing++;
              break;
            case 'Embroidery Working':
              statusCounts.embroidery++;
              break;
            case 'Pending':
              statusCounts.pending++;
              break;
          }
        }
      });

      setFilterOptions({
        fabric: fabricOptions,
        garmentType: garmentTypeOptions,
        style: styleOptions,
        brand: brands,
        section: sectionOptions,
        season: seasons,
        party: partyOptions,
        directStitch: new Set(['Yes', 'No']),
        colorPending: new Set(['Yes', 'No']),
        status: statusOptions
      });

      // Process cutting matrix data
      if (cuttingRows && cuttingRows.length > 0) {
        processCuttingMatrixData(cuttingRows, filteredData, colorPendingCount, repeatedLotsCount, statusCounts);
      }

    } catch (err) {
      setError(err.message);
      console.error('Error fetching data:', err);
    } finally {
      setLoading(false);
    }
  };

  const processCuttingMatrixData = (cuttingRows, lotData, colorPendingCount, repeatedLotsCount, statusCounts) => {
    const matrixMap = {};
    let totalPieces = 0;

    let currentLot = null;
    let currentMatrix = [];

    for (let i = 0; i < cuttingRows.length; i++) {
      const row = cuttingRows[i];

      if (row && row[0] && row[0].toString().toLowerCase().includes('cutting matrix')) {
        if (currentLot && currentMatrix.length > 0) {
          matrixMap[currentLot] = [...currentMatrix];
        }

        const lotMatch = row[0].toString().match(/lot\s*(\d+)/i);
        if (lotMatch && lotMatch[1]) {
          currentLot = lotMatch[1];
          currentMatrix = [row];
        } else {
          currentLot = null;
          currentMatrix = [];
        }
      }
      else if (row && row[0] && row[0].toString().toLowerCase().includes('lot number:')) {
        const lotNumber = row[1];
        if (lotNumber) {
          if (currentLot && currentMatrix.length > 0) {
            matrixMap[currentLot] = [...currentMatrix];
          }

          currentLot = lotNumber.toString().trim();
          currentMatrix = [];
          for (let j = i; j < Math.min(i + 10, cuttingRows.length); j++) {
            if (cuttingRows[j] && cuttingRows[j].length > 0) {
              currentMatrix.push(cuttingRows[j]);
            }
          }
          i += 9;
        }
      }
      else if (currentLot && row && row.length > 0) {
        currentMatrix.push(row);
      }
    }

    if (currentLot && currentMatrix.length > 0) {
      matrixMap[currentLot] = [...currentMatrix];
    }

    // Calculate total pieces from matrices
    lotData.forEach(row => {
      const matrix = matrixMap[row['Lot Number']];
      if (matrix) {
        const total = getTotalPcsFromMatrix(matrix);
        if (typeof total === 'number' && total > 0) {
          totalPieces += total;
        }
      }
    });

    setMatrixData(matrixMap);
    setStats(prev => ({
      ...prev,
      totalPieces,
      colorPendingLots: colorPendingCount,
      repeatedLots: repeatedLotsCount,
      statusCounts: statusCounts
    }));
  };

  const applyFilters = () => {
    let result = [...data];

    // Apply each filter if set
    if (filters.fabric && filters.fabric.length > 0) {
      result = result.filter(row => filters.fabric.includes(row.Fabric));
    }
    if (filters.garmentType && filters.garmentType.length > 0) {
      result = result.filter(row => filters.garmentType.includes(row['Garment Type']));
    }
    if (filters.style && filters.style.length > 0) {
      result = result.filter(row => filters.style.includes(row.Style));
    }
    if (filters.brand && filters.brand.length > 0) {
      result = result.filter(row => filters.brand.includes(row.Brand));
    }
    if (filters.section && filters.section.length > 0) {
      result = result.filter(row => filters.section.includes(row.Section) || filters.section.includes(row['M/W/K']));
    }
    if (filters.season && filters.season.length > 0) {
      result = result.filter(row => filters.season.includes(row.Season));
    }
    if (filters.party && filters.party.length > 0) {
      result = result.filter(row => filters.party.includes(row['Party Name']) || filters.party.includes(row.Party));
    }
    if (filters.directStitch) {
      if (filters.directStitch === 'Yes') {
        result = result.filter(row => row['Direct Stitching'] && row['Direct Stitching'].toString().toLowerCase() === 'yes');
      } else if (filters.directStitch === 'No') {
        result = result.filter(row => !row['Direct Stitching'] || row['Direct Stitching'].toString().toLowerCase() !== 'yes');
      }
    }

    // Apply color pending filter
    if (filters.colorPending) {
      if (filters.colorPending === 'Yes') {
        result = result.filter(row => row.hasColorPending === true);
      } else if (filters.colorPending === 'No') {
        result = result.filter(row => row.hasColorPending === false || !row.hasColorPending);
      }
    }

    // Apply remarks filter
    if (filters.remarksStatus) {
      if (filters.remarksStatus === 'With Remarks') {
        result = result.filter(row => {
          const lotNo = (row['Lot Number'] || row.lotNumber)?.toString().trim();
          const lotRemarks = remarksMap[lotNo] || [];
          return lotRemarks.length > 0;
        });
      } else if (filters.remarksStatus === 'Without Remarks') {
        result = result.filter(row => {
          const lotNo = (row['Lot Number'] || row.lotNumber)?.toString().trim();
          const lotRemarks = remarksMap[lotNo] || [];
          return lotRemarks.length === 0;
        });
      }
    }

    // Apply status filter
    if (filters.status && filters.status.length > 0) {
      result = result.filter(row => filters.status.includes(row.status));
    }

    // Apply search filter if search term exists
    if (filters.search) {
      const searchLower = filters.search.toLowerCase();
      result = result.filter(row => {
        const directMatch = Object.values(row).some(value =>
          value && value.toString().toLowerCase().includes(searchLower)
        );
        if (directMatch) return true;

        const lotNumber = (row['Lot Number'] || row.lotNumber)?.toString().trim();
        const lotRemarks = remarksMap[lotNumber] || [];
        return lotRemarks.some(r => r.text && r.text.toLowerCase().includes(searchLower));
      });
    }

    setFilteredData(result);
  };

  const handleFilterChange = (filterName, value) => {
    setFilters(prev => ({
      ...prev,
      [filterName]: value
    }));
  };

  const clearAllFilters = () => {
    setFilters({
      fabric: [],
      garmentType: [],
      style: [],
      brand: [],
      section: [],
      season: [],
      party: [],
      directStitch: '',
      colorPending: '',
      remarksStatus: '',
      status: [],
      search: ''
    });
  };

  const getTotalPcsFromMatrix = (matrix) => {
    if (!matrix || !Array.isArray(matrix)) return 0;

    for (let i = 0; i < matrix.length; i++) {
      const row = matrix[i];
      if (row && row[0] && row[0].toString().toLowerCase().includes('total')) {
        const totalPcs = row[row.length - 1];
        const parsed = parseInt(String(totalPcs || '').replace(/,/g, '').trim(), 10);
        return isNaN(parsed) ? 0 : parsed;
      }
    }

    return 0;
  };

  const getTotalForLot = (lotNumber) => {
    const lotStr = String(lotNumber || '').trim();
    const matrix = matrixData[lotStr];
    if (matrix) {
      const val = getTotalPcsFromMatrix(matrix);
      if (typeof val === 'number' && val > 0) return val;
    }
    const item = data.find(d => String(d['Lot Number'] || d.lotNumber || '').trim() === lotStr);
    if (item && item['Total Pcs'] && item['Total Pcs'] !== 'N/A') {
      const parsed = parseInt(String(item['Total Pcs']).replace(/,/g, '').trim(), 10);
      if (!isNaN(parsed) && parsed > 0) return parsed;
    }
    return 0;
  };

  // Calculate total pieces for filtered data
  const calculateFilteredTotalPieces = () => {
    let total = 0;
    filteredData.forEach(row => {
      const lotNo = (row['Lot Number'] || row.lotNumber)?.toString().trim();
      const pcs = getTotalForLot(lotNo);
      total += (typeof pcs === 'number' ? pcs : (parseInt(pcs, 10) || 0));
    });
    return total;
  };

  // Calculate color pending stats for filtered data
  const calculateColorPendingStats = () => {
    const colorPendingCount = filteredData.filter(row => row.hasColorPending).length;
    const repeatedLotCount = filteredData.filter(row => row.isRepeatedLot).length;
    return { colorPendingCount, repeatedLotCount };
  };

  // Calculate status counts for filtered data
  const calculateStatusStats = () => {
    const statusCounts = {
      direct: 0,
      ready: 0,
      printing: 0,
      embroidery: 0,
      pending: 0
    };

    filteredData.forEach(row => {
      if (row.status) {
        switch (row.status) {
          case 'Direct':
            statusCounts.direct++;
            break;
          case 'Ready for Stitching':
            statusCounts.ready++;
            break;
          case 'Printing Working':
            statusCounts.printing++;
            break;
          case 'Embroidery Working':
            statusCounts.embroidery++;
            break;
          case 'Pending':
            statusCounts.pending++;
            break;
        }
      }
    });

    return statusCounts;
  };

  // Get status badge style
  const getStatusStyle = (status) => {
    switch (status) {
      case 'Direct':
        return {
          background: 'linear-gradient(135deg, #dbeafe 0%, #eff6ff 100%)',
          color: '#1d4ed8',
          border: '1px solid #93c5fd',
          boxShadow: '0 2px 6px rgba(37, 99, 235, 0.12)',
          icon: '⚡'
        };
      case 'Ready for Stitching':
        return {
          background: 'linear-gradient(135deg, #dcfce7 0%, #f0fdf4 100%)',
          color: '#15803d',
          border: '1px solid #86efac',
          boxShadow: '0 2px 6px rgba(22, 163, 74, 0.12)',
          icon: '✅'
        };
      case 'Printing Working':
        return {
          background: 'linear-gradient(135deg, #fef3c7 0%, #fffbeb 100%)',
          color: '#b45309',
          border: '1px solid #fde047',
          boxShadow: '0 2px 6px rgba(217, 119, 6, 0.12)',
          icon: '🖨️'
        };
      case 'Embroidery Working':
        return {
          background: 'linear-gradient(135deg, #f3e8ff 0%, #faf5ff 100%)',
          color: '#7e22ce',
          border: '1px solid #d8b4fe',
          boxShadow: '0 2px 6px rgba(126, 34, 206, 0.12)',
          icon: '🧵'
        };
      case 'Pending':
      default:
        return {
          background: 'linear-gradient(135deg, #f1f5f9 0%, #f8fafc 100%)',
          color: '#475569',
          border: '1px solid #cbd5e1',
          boxShadow: '0 2px 6px rgba(100, 116, 139, 0.08)',
          icon: '⏳'
        };
    }
  };

  // ================= EXCEL EXPORT (MULTI-SHEET) =================
  const exportToExcel = async () => {
    if (filteredData.length === 0) {
      alert('No data to download');
      return;
    }

    try {
      const workbook = new ExcelJS.Workbook();
      workbook.creator = 'Factory Suite Pro';
      workbook.created = new Date();

      const totalPieces = calculateFilteredTotalPieces();
      const { colorPendingCount, repeatedLotCount } = calculateColorPendingStats();
      const statusStats = calculateStatusStats();
      const totalLots = filteredData.length;

      // Grouping data for executive summary
      const garmentMap = {};
      const seasonMap = {};
      const partyMap = {};
      const statusMap = {};

      filteredData.forEach(row => {
        const lotNo = (row['Lot Number'] || row.lotNumber)?.toString().trim();
        const pcs = Number(getTotalForLot(lotNo)) || 0;
        const garment = (row['Garment Type'] || 'Unknown').trim();
        const season = (row['Season'] || 'N/A').trim();
        const party = (row['Party Name'] || (row['Direct Stitching'] === 'Yes' ? 'Direct Stitching' : '—')).trim();
        const status = (row.status || 'Pending').trim();

        if (!garmentMap[garment]) garmentMap[garment] = { totalLots: 0, totalPcs: 0 };
        garmentMap[garment].totalLots += 1;
        garmentMap[garment].totalPcs += pcs;

        if (!seasonMap[season]) seasonMap[season] = { totalLots: 0, totalPcs: 0 };
        seasonMap[season].totalLots += 1;
        seasonMap[season].totalPcs += pcs;

        if (!partyMap[party]) partyMap[party] = { totalLots: 0, totalPcs: 0 };
        partyMap[party].totalLots += 1;
        partyMap[party].totalPcs += pcs;

        if (!statusMap[status]) statusMap[status] = { totalLots: 0, totalPcs: 0 };
        statusMap[status].totalLots += 1;
        statusMap[status].totalPcs += pcs;
      });

      const sortedGarments = Object.keys(garmentMap).map(name => ({
        name,
        totalLots: garmentMap[name].totalLots,
        totalPcs: garmentMap[name].totalPcs
      })).sort((a, b) => b.totalPcs - a.totalPcs);

      const sortedSeasons = Object.keys(seasonMap).map(name => ({
        name,
        totalLots: seasonMap[name].totalLots,
        totalPcs: seasonMap[name].totalPcs
      })).sort((a, b) => b.totalPcs - a.totalPcs);

      const sortedParties = Object.keys(partyMap).map(name => ({
        name,
        totalLots: partyMap[name].totalLots,
        totalPcs: partyMap[name].totalPcs
      })).sort((a, b) => b.totalPcs - a.totalPcs);

      const sortedStatuses = Object.keys(statusMap).map(name => ({
        name,
        totalLots: statusMap[name].totalLots,
        totalPcs: statusMap[name].totalPcs
      })).sort((a, b) => b.totalPcs - a.totalPcs);

      const thinBorder = {
        top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
      };

      // ================= SHEET 1: DATA TABLE =================
      const ws1 = workbook.addWorksheet('Pending Stitching', {
        views: [{ showGridLines: true }]
      });

      // Title Banner
      ws1.mergeCells('A1:O1');
      const titleCell = ws1.getCell('A1');
      titleCell.value = 'FACTORY SUITE PRO - PENDING ISSUES TO STITCHING AFTER CUTTING';
      titleCell.font = { name: 'Segoe UI', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
      titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E1B4B' } };
      titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
      ws1.getRow(1).height = 30;

      // Subtitle KPI Banner
      ws1.mergeCells('A2:O2');
      const subCell = ws1.getCell('A2');
      subCell.value = `Total Lots: ${totalLots}   |   Total Pieces: ${totalPieces.toLocaleString()}   |   Color Pending: ${colorPendingCount}   |   Repeated Lots: ${repeatedLotCount}   |   Direct Lots: ${statusStats.direct}   |   Generated: ${new Date().toLocaleDateString('en-IN')} ${new Date().toLocaleTimeString('en-IN')}`;
      subCell.font = { name: 'Segoe UI', size: 9.5, color: { argb: 'FFC7D2FE' } };
      subCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF312E81' } };
      subCell.alignment = { horizontal: 'center', vertical: 'middle' };
      ws1.getRow(2).height = 22;

      // Header Row
      const tableHeaders = [
        '#', 'Lot #', 'Garment Type', 'Style', 'Fabric', 'Brand',
        'Total Pcs', 'Section', 'Season', 'Party Name', 'Direct',
        'Color Status', 'Status', 'Pending Colors', 'Remarks'
      ];
      const headerRow = ws1.addRow(tableHeaders);
      headerRow.height = 24;
      headerRow.eachCell((cell) => {
        cell.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = thinBorder;
      });

      // Data Rows
      filteredData.forEach((row, index) => {
        const lotNo = (row['Lot Number'] || row.lotNumber)?.toString().trim();
        const lotRemarks = remarksMap[lotNo] || [];
        const latestRemark = lotRemarks.length > 0 ? lotRemarks[lotRemarks.length - 1].text : '';
        const isRepeated = row.isRepeatedLot || false;
        const isColorPending = row.hasColorPending || false;
        const pcs = Number(getTotalForLot(lotNo)) || 0;

        const r = ws1.addRow([
          index + 1,
          isRepeated ? `★ ${row['Lot Number']}` : row['Lot Number'],
          row['Garment Type'] || 'N/A',
          row['Style'] || 'N/A',
          row['Fabric'] || 'N/A',
          row['Brand'] || 'N/A',
          pcs,
          row['Section'] || row['M/W/K'] || '—',
          row['Season'] || 'N/A',
          row['Party Name'] || row['Party'] || '—',
          row['Direct Stitching'] && row['Direct Stitching'].toString().toLowerCase() === 'yes' ? 'Yes' : 'No',
          isColorPending ? 'Color Pending' : 'OK',
          row.status || 'Pending',
          row.pendingColorsText || '',
          latestRemark || ''
        ]);

        r.height = 20;

        // Apply borders and alternating fill
        r.eachCell((cell) => {
          cell.font = { name: 'Segoe UI', size: 9 };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.border = thinBorder;
          if (index % 2 === 1) {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
          }
        });

        // Highlight Lot Number if repeated
        if (isRepeated) {
          r.getCell(2).font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFB45309' } };
          r.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
        }

        // Total Pcs formatting
        r.getCell(7).numFmt = '#,##0';
        r.getCell(7).font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FF1E40AF' } };

        // Color status badge
        if (isColorPending) {
          r.getCell(12).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
          r.getCell(12).font = { name: 'Segoe UI', size: 8.5, bold: true, color: { argb: 'FFDC2626' } };
        } else {
          r.getCell(12).font = { name: 'Segoe UI', size: 8.5, color: { argb: 'FF16A34A' } };
        }

        // Direct badge
        if (row['Direct Stitching'] && row['Direct Stitching'].toString().toLowerCase() === 'yes') {
          r.getCell(11).font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FF15803D' } };
        }
      });

      // Total Row
      const totalRow1 = ws1.addRow([
        '',
        `TOTAL (${totalLots} Lots)`,
        '',
        '',
        '',
        '',
        totalPieces,
        '',
        '',
        '',
        '',
        `${colorPendingCount} Color Pending`,
        `${statusStats.direct} Direct`,
        '',
        `${repeatedLotCount} Repeated`
      ]);
      totalRow1.height = 24;
      totalRow1.eachCell((cell) => {
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

      // Set column widths
      const colWidths1 = [6, 14, 18, 18, 18, 15, 14, 12, 14, 20, 12, 16, 18, 24, 30];
      colWidths1.forEach((w, i) => {
        ws1.getColumn(i + 1).width = w;
      });

      // ================= SHEET 2: EXECUTIVE SUMMARY =================
      const ws2 = workbook.addWorksheet('Executive Summary', {
        views: [{ showGridLines: true }]
      });

      // Section 1: Garment Type Breakdown
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

      // Section 2: Season Breakdown
      const seasonStartRow = ws2.rowCount + 1;
      ws2.mergeCells(`A${seasonStartRow}:D${seasonStartRow}`);
      const sTitle = ws2.getCell(`A${seasonStartRow}`);
      sTitle.value = '2. SEASON WISE BREAKDOWN (LOTS & PIECES DISTRIBUTION)';
      sTitle.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      sTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4338CA' } };
      sTitle.alignment = { horizontal: 'left', vertical: 'middle' };
      ws2.getRow(seasonStartRow).height = 26;

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

      // Section 3: Party Summary
      const partyStartRow = ws2.rowCount + 1;
      ws2.mergeCells(`A${partyStartRow}:D${partyStartRow}`);
      const pTitle = ws2.getCell(`A${partyStartRow}`);
      pTitle.value = '3. PARTY SUMMARY & WORKLOAD ALLOCATION';
      pTitle.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      pTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E40AF' } };
      pTitle.alignment = { horizontal: 'left', vertical: 'middle' };
      ws2.getRow(partyStartRow).height = 26;

      const pHeader = ws2.addRow(['Party Name', 'Total Lots', 'Total Pieces (Qty)', 'Share %']);
      pHeader.height = 22;
      pHeader.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = thinBorder;
      });

      sortedParties.forEach((item, idx) => {
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

      const pTotalRow = ws2.addRow(['TOTAL', totalLots, totalPieces, 1]);
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

      // Section 4: Status Breakdown
      const statusStartRow = ws2.rowCount + 1;
      ws2.mergeCells(`A${statusStartRow}:D${statusStartRow}`);
      const stTitle = ws2.getCell(`A${statusStartRow}`);
      stTitle.value = '4. STATUS & PROCESS WORKFLOW BREAKDOWN';
      stTitle.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      stTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9333EA' } };
      stTitle.alignment = { horizontal: 'left', vertical: 'middle' };
      ws2.getRow(statusStartRow).height = 26;

      const stHeader = ws2.addRow(['Status', 'Total Lots', 'Total Pieces (Qty)', 'Share %']);
      stHeader.height = 22;
      stHeader.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF6B21A8' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = thinBorder;
      });

      sortedStatuses.forEach((item, idx) => {
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

      const stTotalRow = ws2.addRow(['TOTAL', totalLots, totalPieces, 1]);
      stTotalRow.height = 22;
      stTotalRow.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = { top: { style: 'thin' }, bottom: { style: 'double' }, left: { style: 'thin' }, right: { style: 'thin' } };
      });
      stTotalRow.getCell(3).numFmt = '#,##0';
      stTotalRow.getCell(4).numFmt = '0.0%';

      ws2.getColumn(1).width = 34;
      ws2.getColumn(2).width = 16;
      ws2.getColumn(3).width = 22;
      ws2.getColumn(4).width = 16;

      // ================= SHEET 3: APPLIED FILTERS =================
      const ws3 = workbook.addWorksheet('Applied Filters', {
        views: [{ showGridLines: true }]
      });

      ws3.mergeCells('A1:B1');
      const fTitle = ws3.getCell('A1');
      fTitle.value = 'APPLIED FILTERS & REPORT METADATA';
      fTitle.font = { name: 'Segoe UI', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
      fTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
      fTitle.alignment = { horizontal: 'left', vertical: 'middle' };
      ws3.getRow(1).height = 28;

      const filterItems = [
        ['Report Generated', `${new Date().toLocaleDateString('en-IN')} ${new Date().toLocaleTimeString('en-IN')}`],
        ['Total Lots Exported', totalLots],
        ['Total Pieces Exported', totalPieces.toLocaleString()],
        ['Color Pending Lots', colorPendingCount],
        ['Repeated Lots', repeatedLotCount],
        ['Direct Lots', statusStats.direct],
        ['Fabric Filter', filters.fabric.length ? filters.fabric.join(', ') : 'All Fabrics'],
        ['Garment Type Filter', filters.garmentType.length ? filters.garmentType.join(', ') : 'All Types'],
        ['Style Filter', filters.style.length ? filters.style.join(', ') : 'All Styles'],
        ['Brand Filter', filters.brand.length ? filters.brand.join(', ') : 'All Brands'],
        ['Section Filter', filters.section.length ? filters.section.join(', ') : 'All Sections'],
        ['Season Filter', filters.season.length ? filters.season.join(', ') : 'All Seasons'],
        ['Party Name Filter', filters.party.length ? filters.party.join(', ') : 'All Parties'],
        ['Direct Stitching Filter', filters.directStitch || 'All'],
        ['Color Pending Filter', filters.colorPending || 'All'],
        ['Status Filter', filters.status.length ? filters.status.join(', ') : 'All Statuses'],
        ['Search Query', filters.search || 'None']
      ];

      filterItems.forEach(([k, v], idx) => {
        const r = ws3.addRow([k, v]);
        r.height = 20;
        r.getCell(1).font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FF1E293B' } };
        r.getCell(2).font = { name: 'Segoe UI', size: 9.5, color: { argb: 'FF334155' } };
        r.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
        r.getCell(2).alignment = { horizontal: 'left', vertical: 'middle' };
        r.eachCell(c => {
          c.border = thinBorder;
          if (idx % 2 === 1) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        });
      });

      ws3.getColumn(1).width = 24;
      ws3.getColumn(2).width = 50;

      // Generate and Save Excel File
      const buffer = await workbook.xlsx.writeBuffer();
      const ts = new Date().toISOString().slice(0, 10);
      saveAs(new Blob([buffer]), `Pending_Stitching_Issues_${ts}.xlsx`);

    } catch (e) {
      console.error('Excel export error:', e);
      alert(`Excel export failed: ${e.message}`);
    }
  };

  // ================= PROFESSIONAL PDF EXPORT (A3 LANDSCAPE WITH PICTURES) =================
  const exportToPDF = async () => {
    if (filteredData.length === 0) {
      alert('No data to download');
      return;
    }

    try {
      const doc = new jsPDF({
        orientation: "landscape",
        unit: "pt",
        format: "a3"
      });

      const pageW = doc.internal.pageSize.getWidth();
      const pageH = doc.internal.pageSize.getHeight();
      const margin = 15;
      const contentWidth = pageW - (margin * 2);

      const { colorPendingCount, repeatedLotCount } = calculateColorPendingStats();
      const totalPieces = calculateFilteredTotalPieces();
      const totalLots = filteredData.length;
      const statusStats = calculateStatusStats();

      // Pre-load base64 images for all filtered lots
      const imageBase64Map = new Map();
      await Promise.all(
        filteredData.map(async (row) => {
          const lotNo = (row['Lot Number'] || row.lotNumber)?.toString().trim();
          const imgUrl = row['Image'] && row['Image'] !== 'N/A' ? row['Image'] : '';
          if (imgUrl) {
            const directUrl = getDirectImageUrl(imgUrl);
            const b64 = await getBase64ImageFromUrl(directUrl);
            if (b64) {
              imageBase64Map.set(lotNo, b64);
            }
          }
        })
      );

      // Grouping data for executive summary
      const garmentMap = {};
      const seasonMap = {};
      const partyMap = {};
      const statusMap = {};

      filteredData.forEach(row => {
        const lotNo = (row['Lot Number'] || row.lotNumber)?.toString().trim();
        const pcs = Number(getTotalForLot(lotNo)) || 0;
        const garment = (row['Garment Type'] || 'Unknown').trim();
        const season = (row['Season'] || 'N/A').trim();
        const party = (row['Party Name'] || (row['Direct Stitching'] === 'Yes' ? 'Direct Stitching' : '—')).trim();
        const status = (row.status || 'Pending').trim();

        if (!garmentMap[garment]) garmentMap[garment] = { totalLots: 0, totalPcs: 0 };
        garmentMap[garment].totalLots += 1;
        garmentMap[garment].totalPcs += pcs;

        if (!seasonMap[season]) seasonMap[season] = { totalLots: 0, totalPcs: 0 };
        seasonMap[season].totalLots += 1;
        seasonMap[season].totalPcs += pcs;

        if (!partyMap[party]) partyMap[party] = { totalLots: 0, totalPcs: 0 };
        partyMap[party].totalLots += 1;
        partyMap[party].totalPcs += pcs;

        if (!statusMap[status]) statusMap[status] = { totalLots: 0, totalPcs: 0 };
        statusMap[status].totalLots += 1;
        statusMap[status].totalPcs += pcs;
      });

      const sortedGarments = Object.keys(garmentMap).map(name => ({
        name,
        totalLots: garmentMap[name].totalLots,
        totalPcs: garmentMap[name].totalPcs
      })).sort((a, b) => b.totalPcs - a.totalPcs);

      const sortedSeasons = Object.keys(seasonMap).map(name => ({
        name,
        totalLots: seasonMap[name].totalLots,
        totalPcs: seasonMap[name].totalPcs
      })).sort((a, b) => b.totalPcs - a.totalPcs);

      const sortedParties = Object.keys(partyMap).map(name => ({
        name,
        totalLots: partyMap[name].totalLots,
        totalPcs: partyMap[name].totalPcs
      })).sort((a, b) => b.totalPcs - a.totalPcs);

      const sortedStatuses = Object.keys(statusMap).map(name => ({
        name,
        totalLots: statusMap[name].totalLots,
        totalPcs: statusMap[name].totalPcs
      })).sort((a, b) => b.totalPcs - a.totalPcs);

      // Main Header Block
      doc.setFillColor(15, 23, 42); // Dark Navy
      doc.rect(margin, 12, contentWidth, 48, 'F');

      doc.setFontSize(14);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(255, 255, 255);
      doc.text("FACTORY SUITE PRO - PENDING ISSUES TO STITCHING", pageW / 2, 30, { align: 'center' });

      doc.setFontSize(8.5);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(199, 210, 254);
      const subText = `Total Lots: ${totalLots}   |   Total Pieces: ${totalPieces.toLocaleString()}   |   Color Pending: ${colorPendingCount}   |   Repeated Lots: ${repeatedLotCount}   |   Direct Lots: ${statusStats.direct}   |   Ready: ${statusStats.ready}   |   Printing: ${statusStats.printing}   |   Embroidery: ${statusStats.embroidery}   |   Pending: ${statusStats.pending}`;
      doc.text(subText, pageW / 2, 48, { align: 'center' });

      // Filter Banner
      doc.setFillColor(241, 245, 249);
      doc.rect(margin, 63, contentWidth, 16, 'F');
      doc.setFontSize(8);
      doc.setFont('helvetica', 'italic');
      doc.setTextColor(0, 0, 0); // Pure Black
      const filterSummary = `Filters: Fabric: ${filters.fabric.length ? filters.fabric.join(', ') : 'All'} | Garment: ${filters.garmentType.length ? filters.garmentType.join(', ') : 'All'} | Style: ${filters.style.length ? filters.style.join(', ') : 'All'} | Brand: ${filters.brand.length ? filters.brand.join(', ') : 'All'} | Section: ${filters.section.length ? filters.section.join(', ') : 'All'} | Season: ${filters.season.length ? filters.season.join(', ') : 'All'} | Party: ${filters.party.length ? filters.party.join(', ') : 'All'} | Direct: ${filters.directStitch || 'All'} | Remarks: ${filters.remarksStatus || 'All'} | Status: ${filters.status.length ? filters.status.join(', ') : 'All'}`;
      doc.text(filterSummary, pageW / 2, 74, { align: 'center' });

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
        'Section',
        'Season',
        'Party Name',
        'Direct',
        'Color Status',
        'Status',
        'Remarks'
      ];

      const tableBody = filteredData.map((row, idx) => {
        const lotNo = (row['Lot Number'] || row.lotNumber)?.toString().trim();
        const lotRemarks = remarksMap[lotNo] || [];
        const latestRemark = lotRemarks.length > 0 ? lotRemarks[lotRemarks.length - 1].text : '—';
        const isColorPending = row.hasColorPending || false;
        const pendingColors = row.pendingColors || [];
        const pcs = Number(getTotalForLot(lotNo)) || 0;

        let colorStatusText = 'OK';
        if (isColorPending) {
          colorStatusText = pendingColors.length > 0
            ? `Pending: ${pendingColors.slice(0, 2).join(', ')}${pendingColors.length > 2 ? '...' : ''}`
            : 'Color Pending';
        }

        return [
          (idx + 1).toString(),
          '', // Image cell rendered via didDrawCell
          row.isRepeatedLot ? `★ ${row['Lot Number']}` : row['Lot Number'],
          row['Garment Type'] || 'N/A',
          row['Style'] || 'N/A',
          row['Fabric'] || 'N/A',
          row['Brand'] || 'N/A',
          pcs.toLocaleString(),
          row['Section'] || row['M/W/K'] || '—',
          row['Season'] || 'N/A',
          row['Party Name'] || row['Party'] || '—',
          row['Direct Stitching'] && row['Direct Stitching'].toString().toLowerCase() === 'yes' ? 'Yes' : 'No',
          colorStatusText,
          row.status || 'Pending',
          latestRemark
        ];
      });

      // Add Total Row
      tableBody.push([
        '',
        '',
        `TOTAL (${totalLots})`,
        '',
        '',
        '',
        '',
        totalPieces.toLocaleString(),
        '',
        '',
        '',
        `${statusStats.direct} Direct`,
        `${colorPendingCount} Color Pending`,
        `${statusStats.pending} Pending`,
        `${repeatedLotCount} Repeated`
      ]);

      const baseWidths = [26, 42, 68, 88, 88, 90, 78, 62, 50, 62, 95, 48, 100, 90, 145];
      const sumBase = baseWidths.reduce((a, b) => a + b, 0);
      const scaleFactor = contentWidth / sumBase;
      const columnStyles = {};
      baseWidths.forEach((w, idx) => {
        columnStyles[idx] = {
          cellWidth: w * scaleFactor,
          halign: 'center',
          valign: 'middle'
        };
      });
      columnStyles[7].fontStyle = 'bold';

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
          lineColor: [0, 0, 0], // Black grid lines
          lineWidth: 0.3,
          fontStyle: 'normal',
          minCellHeight: 25,
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
          minCellHeight: 14,
        },
        alternateRowStyles: {
          fillColor: [248, 250, 252],
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
              data.cell.styles.valign = 'middle';
              return;
            }

            const row = filteredData[rowIndex];
            if (!row) return;

            // Repeated lot styling
            if (row.isRepeatedLot && data.column.index === 2) {
              data.cell.styles.fillColor = [254, 243, 199];
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.textColor = [146, 64, 14];
            }

            // Color status styling
            if (data.column.index === 12) {
              if (row.hasColorPending) {
                data.cell.styles.fillColor = [254, 226, 226];
                data.cell.styles.textColor = [220, 38, 38];
                data.cell.styles.fontStyle = 'bold';
              } else {
                data.cell.styles.textColor = [22, 163, 74];
              }
            }

            // Direct styling
            if (data.column.index === 11 && row['Direct Stitching'] && row['Direct Stitching'].toString().toLowerCase() === 'yes') {
              data.cell.styles.textColor = [21, 128, 61];
              data.cell.styles.fontStyle = 'bold';
            }
          }
        },
        didDrawCell: function (data) {
          if (data.column.index === 1 && data.section === 'body') {
            const rowIndex = data.row.index;
            const isTotalRow = rowIndex === tableBody.length - 1;
            if (isTotalRow) return;

            const row = filteredData[rowIndex];
            if (!row) return;

            const lotNo = (row['Lot Number'] || row.lotNumber)?.toString().trim();
            const b64 = imageBase64Map.get(lotNo);
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
      // 1. Garment Body
      const gBody = sortedGarments.map(item => {
        const pct = totalPieces > 0 ? ((item.totalPcs / totalPieces) * 100).toFixed(1) : "0.0";
        return [
          item.name,
          item.totalLots.toString(),
          item.totalPcs.toLocaleString(),
          `${pct}%`
        ];
      });
      gBody.push([
        "TOTAL",
        totalLots.toString(),
        totalPieces.toLocaleString(),
        "100.0%"
      ]);

      // 2. Season Body
      const sBody = sortedSeasons.map(item => {
        const pct = totalPieces > 0 ? ((item.totalPcs / totalPieces) * 100).toFixed(1) : "0.0";
        return [
          item.name,
          item.totalLots.toString(),
          item.totalPcs.toLocaleString(),
          `${pct}%`
        ];
      });
      sBody.push([
        "TOTAL",
        totalLots.toString(),
        totalPieces.toLocaleString(),
        "100.0%"
      ]);

      // 3. Party Body
      const pBody = sortedParties.map(party => {
        const pct = totalPieces > 0 ? ((party.totalPcs / totalPieces) * 100).toFixed(1) : "0.0";
        return [
          party.name,
          party.totalLots.toString(),
          party.totalPcs.toLocaleString(),
          `${pct}%`
        ];
      });
      pBody.push([
        "TOTAL",
        totalLots.toString(),
        totalPieces.toLocaleString(),
        "100.0%"
      ]);

      // 4. Status Body
      const stBody = sortedStatuses.map(st => {
        const pct = totalPieces > 0 ? ((st.totalPcs / totalPieces) * 100).toFixed(1) : "0.0";
        return [
          st.name,
          st.totalLots.toString(),
          st.totalPcs.toLocaleString(),
          `${pct}%`
        ];
      });
      stBody.push([
        "TOTAL",
        totalLots.toString(),
        totalPieces.toLocaleString(),
        "100.0%"
      ]);

      const maxRows = Math.max(gBody.length, sBody.length, pBody.length, stBody.length);
      const approxSummaryHeight = 55 + (maxRows * 18);

      let summaryStartY = doc.lastAutoTable.finalY + 22;
      const neededSpace = approxSummaryHeight + 35;
      if (summaryStartY + neededSpace > pageH - 30) {
        doc.addPage();
        summaryStartY = 40;
      } else {
        doc.setDrawColor(203, 213, 225);
        doc.setLineWidth(0.8);
        doc.line(margin, summaryStartY - 8, pageW - margin, summaryStartY - 8);
      }

      // Title & KPI Subtitle
      doc.setFontSize(14);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(0, 0, 0); // Pure Black
      const summaryTitle = "EXECUTIVE SUMMARY & PRODUCTION BREAKDOWN";
      doc.text(summaryTitle, pageW / 2, summaryStartY + 4, { align: 'center' });

      doc.setFontSize(9);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(0, 0, 0);
      const summarySub = `Total Lots: ${totalLots}   |   Total Pieces: ${totalPieces.toLocaleString()} Pcs   |   Seasons: ${sortedSeasons.length}   |   Parties: ${sortedParties.length}   |   Garment Types: ${sortedGarments.length}   |   Repeated Lots: ${repeatedLotCount}`;
      doc.text(summarySub, pageW / 2, summaryStartY + 16, { align: 'center' });

      const sectionTitleY = summaryStartY + 30;
      const tableStartY = sectionTitleY + 6;

      // 4 Columns Side-by-Side Configuration
      const gap = 12;
      const totalCols = 4;
      const totalGaps = gap * (totalCols - 1);
      const colWidth = (contentWidth - totalGaps) / totalCols;
      const col1X = margin;
      const col2X = col1X + colWidth + gap;
      const col3X = col2X + colWidth + gap;
      const col4X = col3X + colWidth + gap;

      // Section Titles above each Column
      doc.setFontSize(9.5);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(0, 0, 0);
      doc.text("1. GARMENT BREAKDOWN", col1X + colWidth / 2, sectionTitleY, { align: 'center' });
      doc.text("2. SEASON BREAKDOWN", col2X + colWidth / 2, sectionTitleY, { align: 'center' });
      doc.text("3. PARTY SUMMARY", col3X + colWidth / 2, sectionTitleY, { align: 'center' });
      doc.text("4. STATUS / PROCESS BREAKDOWN", col4X + colWidth / 2, sectionTitleY, { align: 'center' });

      const summaryColStyles = {
        0: { cellWidth: colWidth * 0.40, halign: 'center', valign: 'middle' },
        1: { cellWidth: colWidth * 0.18, halign: 'center', valign: 'middle' },
        2: { cellWidth: colWidth * 0.24, halign: 'center', valign: 'middle' },
        3: { cellWidth: colWidth * 0.18, halign: 'center', valign: 'middle' },
      };

      // Column 1 Table: Garment Type Breakdown
      autoTable(doc, {
        head: [['Garment Type', 'Lots', 'Total Pcs', 'Share %']],
        body: gBody,
        startY: tableStartY,
        tableWidth: colWidth,
        margin: { left: col1X, right: pageW - (col1X + colWidth) },
        theme: "grid",
        styles: {
          fontSize: 8.5,
          cellPadding: { top: 3.5, right: 2, bottom: 3.5, left: 2 },
          overflow: "linebreak",
          valign: 'middle',
          halign: 'center',
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          lineWidth: 0.3,
        },
        headStyles: {
          fillColor: [15, 118, 110], // Teal
          textColor: [255, 255, 255],
          fontStyle: "bold",
          fontSize: 8.5,
          halign: 'center',
          valign: 'middle',
          cellPadding: { top: 4, right: 2, bottom: 4, left: 2 },
        },
        columnStyles: summaryColStyles,
        didParseCell: function (data) {
          if (data.section === 'body') {
            data.cell.styles.textColor = [0, 0, 0];
            data.cell.styles.halign = 'center';
            data.cell.styles.valign = 'middle';
            if (data.row.index === gBody.length - 1) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });

      // Column 2 Table: Season Breakdown
      autoTable(doc, {
        head: [['Season', 'Lots', 'Total Pcs', 'Share %']],
        body: sBody,
        startY: tableStartY,
        tableWidth: colWidth,
        margin: { left: col2X, right: pageW - (col2X + colWidth) },
        theme: "grid",
        styles: {
          fontSize: 8.5,
          cellPadding: { top: 3.5, right: 2, bottom: 3.5, left: 2 },
          overflow: "linebreak",
          valign: 'middle',
          halign: 'center',
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          lineWidth: 0.3,
        },
        headStyles: {
          fillColor: [67, 56, 202], // Indigo
          textColor: [255, 255, 255],
          fontStyle: "bold",
          fontSize: 8.5,
          halign: 'center',
          valign: 'middle',
          cellPadding: { top: 4, right: 2, bottom: 4, left: 2 },
        },
        columnStyles: summaryColStyles,
        didParseCell: function (data) {
          if (data.section === 'body') {
            data.cell.styles.textColor = [0, 0, 0];
            data.cell.styles.halign = 'center';
            data.cell.styles.valign = 'middle';
            if (data.row.index === sBody.length - 1) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });

      // Column 3 Table: Party Summary
      autoTable(doc, {
        head: [['Party Name', 'Lots', 'Total Pcs', 'Share %']],
        body: pBody,
        startY: tableStartY,
        tableWidth: colWidth,
        margin: { left: col3X, right: pageW - (col3X + colWidth) },
        theme: "grid",
        styles: {
          fontSize: 8.5,
          cellPadding: { top: 3.5, right: 2, bottom: 3.5, left: 2 },
          overflow: "linebreak",
          valign: 'middle',
          halign: 'center',
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          lineWidth: 0.3,
        },
        headStyles: {
          fillColor: [30, 64, 175], // Blue
          textColor: [255, 255, 255],
          fontStyle: "bold",
          fontSize: 8.5,
          halign: 'center',
          valign: 'middle',
          cellPadding: { top: 4, right: 2, bottom: 4, left: 2 },
        },
        columnStyles: summaryColStyles,
        didParseCell: function (data) {
          if (data.section === 'body') {
            data.cell.styles.textColor = [0, 0, 0];
            data.cell.styles.halign = 'center';
            data.cell.styles.valign = 'middle';
            if (data.row.index === pBody.length - 1) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });

      // Column 4 Table: Status Breakdown
      autoTable(doc, {
        head: [['Status', 'Lots', 'Total Pcs', 'Share %']],
        body: stBody,
        startY: tableStartY,
        tableWidth: colWidth,
        margin: { left: col4X, right: pageW - (col4X + colWidth) },
        theme: "grid",
        styles: {
          fontSize: 8.5,
          cellPadding: { top: 3.5, right: 2, bottom: 3.5, left: 2 },
          overflow: "linebreak",
          valign: 'middle',
          halign: 'center',
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          lineWidth: 0.3,
        },
        headStyles: {
          fillColor: [147, 51, 234], // Purple
          textColor: [255, 255, 255],
          fontStyle: "bold",
          fontSize: 8.5,
          halign: 'center',
          valign: 'middle',
          cellPadding: { top: 4, right: 2, bottom: 4, left: 2 },
        },
        columnStyles: summaryColStyles,
        didParseCell: function (data) {
          if (data.section === 'body') {
            data.cell.styles.textColor = [0, 0, 0];
            data.cell.styles.halign = 'center';
            data.cell.styles.valign = 'middle';
            if (data.row.index === stBody.length - 1) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });

      // Footer with Page Numbers
      const totalPages = doc.internal.getNumberOfPages();
      for (let i = 1; i <= totalPages; i++) {
        doc.setPage(i);
        doc.setFontSize(8);
        doc.setTextColor(0, 0, 0); // Pure Black
        doc.text(
          `Factory Suite Pro  |  Confidential Production Report  |  Page ${i} of ${totalPages}`,
          pageW / 2,
          pageH - 12,
          { align: 'center' }
        );
      }

      // Save PDF File
      const ts = new Date().toISOString().slice(0, 10);
      doc.save(`Pending_Stitching_Report_${ts}.pdf`);

    } catch (e) {
      console.error('PDF export error:', e);
      alert(`PDF export failed: ${e.message}`);
    }
  };

  // Aliases for compatibility
  const downloadPDF = exportToPDF;
  const downloadCSV = exportToExcel;
  const exportToCSV = exportToExcel;

  const handleBack = () => {
    window.history.back();
  };

  // EmbroideryChallan Theme Styles
  const styles = {
    container: {
      minHeight: "100vh",
      backgroundColor: "#f8fafc",
      backgroundImage:
        "radial-gradient(at 0% 0%, rgba(99, 102, 241, 0.08) 0px, transparent 50%), " +
        "radial-gradient(at 100% 0%, rgba(236, 72, 153, 0.06) 0px, transparent 50%), " +
        "radial-gradient(at 50% 100%, rgba(16, 185, 129, 0.06) 0px, transparent 50%)",
      padding: "32px",
      fontFamily: "'Plus Jakarta Sans', 'Inter', system-ui, sans-serif",
      color: "#0f172a"
    },

    header: {
      background: "linear-gradient(135deg, #1e1b4b 0%, #312e81 40%, #4338ca 100%)",
      borderRadius: "24px",
      padding: "32px 36px",
      marginBottom: "28px",
      boxShadow: "0 20px 40px -10px rgba(49, 46, 129, 0.3)",
      color: "white"
    },

    headerContentBox: {
      display: "flex",
      justifyContent: "space-between",
      alignItems: "center",
      flexWrap: "wrap",
      gap: "20px"
    },

    titleSection: {
      flex: 1
    },

    title: {
      margin: 0,
      fontSize: "2.4rem",
      fontWeight: "800",
      color: "#ffffff"
    },

    subtitle: {
      margin: "8px 0 0 0",
      fontSize: "1rem",
      color: "#c7d2fe",
      fontWeight: "400"
    },

    statsCard: {
      display: "flex",
      gap: "24px",
      background: "rgba(255, 255, 255, 0.12)",
      border: "1px solid rgba(255, 255, 255, 0.2)",
      padding: "16px 24px",
      borderRadius: "20px",
      backdropFilter: "blur(12px)",
      flexWrap: "wrap"
    },

    statItem: {
      textAlign: "center",
      minWidth: "70px"
    },

    statNumber: {
      display: "block",
      fontSize: "1.8rem",
      fontWeight: "800",
      color: "#ffffff",
      lineHeight: "1.2"
    },

    statLabel: {
      fontSize: "0.75rem",
      color: "#c7d2fe",
      fontWeight: "700",
      textTransform: "uppercase",
      letterSpacing: "0.05em"
    },

    controlsContainer: {
      background: "#ffffff",
      borderRadius: "20px",
      padding: "24px",
      marginBottom: "28px",
      boxShadow: "0 10px 30px rgba(0, 0, 0, 0.03)",
      border: "1px solid #e2e8f0"
    },

    controlsTopRow: {
      display: "flex",
      justifyContent: "space-between",
      alignItems: "center",
      marginBottom: "20px",
      flexWrap: "wrap",
      gap: "20px"
    },

    searchContainer: {
      flex: 1,
      minWidth: "300px"
    },

    searchInput: {
      width: "100%",
      padding: "12px 20px",
      border: "1px solid #cbd5e1",
      borderRadius: "14px",
      fontSize: "0.95rem",
      background: "#f8fafc",
      color: "#0f172a",
      fontWeight: "500",
      outline: "none",
      transition: "all 0.2s ease",
      boxSizing: "border-box"
    },

    mainActions: {
      display: "flex",
      gap: "12px",
      flexWrap: "wrap"
    },

    backBtn: {
      background: "linear-gradient(135deg, #475569 0%, #334155 100%)",
      color: "#ffffff",
      border: "none",
      padding: "10px 20px",
      borderRadius: "12px",
      fontSize: "0.95rem",
      fontWeight: "600",
      cursor: "pointer",
      transition: "all 0.2s ease",
      boxShadow: "0 4px 12px rgba(51, 65, 85, 0.25)"
    },

    refreshBtn: {
      background: "linear-gradient(135deg, #10b981 0%, #059669 100%)",
      color: "white",
      border: "none",
      padding: "10px 20px",
      borderRadius: "12px",
      fontSize: "0.95rem",
      fontWeight: "600",
      cursor: "pointer",
      transition: "all 0.2s ease",
      display: "flex",
      alignItems: "center",
      gap: "8px",
      boxShadow: "0 4px 12px rgba(16, 185, 129, 0.2)"
    },

    exportBtnPdf: {
      background: "linear-gradient(135deg, #ef4444 0%, #dc2626 100%)",
      color: "white",
      border: "none",
      padding: "10px 20px",
      borderRadius: "12px",
      fontSize: "0.95rem",
      fontWeight: "600",
      cursor: "pointer",
      transition: "all 0.2s ease",
      boxShadow: "0 4px 12px rgba(239, 68, 68, 0.2)"
    },

    exportBtnExcel: {
      background: "linear-gradient(135deg, #4f46e5 0%, #4338ca 100%)",
      color: "white",
      border: "none",
      padding: "10px 20px",
      borderRadius: "12px",
      fontSize: "0.95rem",
      fontWeight: "600",
      cursor: "pointer",
      transition: "all 0.2s ease",
      boxShadow: "0 4px 12px rgba(79, 70, 229, 0.2)"
    },

    filterSection: {
      background: "#ffffff",
      borderRadius: "20px",
      padding: "24px",
      marginBottom: "28px",
      boxShadow: "0 10px 30px rgba(0, 0, 0, 0.03)",
      border: "1px solid #e2e8f0"
    },

    filterHeader: {
      display: "flex",
      justifyContent: "space-between",
      alignItems: "center",
      marginBottom: "20px"
    },

    filterTitle: {
      fontSize: "1.2rem",
      fontWeight: "700",
      color: "#0f172a",
      margin: 0,
      display: "flex",
      alignItems: "center",
      gap: "10px"
    },

    clearFiltersBtn: {
      background: "linear-gradient(135deg, #ef4444 0%, #dc2626 100%)",
      border: "none",
      color: "#ffffff",
      padding: "6px 16px",
      borderRadius: "10px",
      fontSize: "0.85rem",
      fontWeight: "600",
      cursor: "pointer",
      transition: "all 0.2s ease",
      boxShadow: "0 4px 12px rgba(239, 68, 68, 0.2)"
    },

    filterGrid: {
      display: "grid",
      gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))",
      gap: "16px"
    },

    filterGroup: {
      display: "flex",
      flexDirection: "column",
      gap: "6px"
    },

    filterLabel: {
      fontSize: "0.75rem",
      fontWeight: "700",
      color: "#475569",
      textTransform: "uppercase",
      letterSpacing: "0.05em"
    },

    filterSelect: {
      width: "100%",
      padding: "9px 12px",
      borderRadius: "12px",
      border: "1px solid #cbd5e1",
      fontSize: "0.9rem",
      background: "#f8fafc",
      color: "#0f172a",
      fontWeight: "500",
      cursor: "pointer",
      outline: "none",
      transition: "all 0.2s ease"
    },

    tableContainer: {
      background: "#ffffff",
      borderRadius: "16px",
      overflow: "hidden",
      boxShadow: "0 10px 30px rgba(0, 0, 0, 0.03)",
      marginBottom: "20px",
      overflowX: "auto",
      border: "1px solid #cbd5e1"
    },

    table: {
      width: "100%",
      borderCollapse: "collapse",
      minWidth: "1200px",
      border: "1px solid #cbd5e1"
    },

    tableHeader: {
      background: "linear-gradient(135deg, #1e1b4b 0%, #312e81 100%)",
      color: "#ffffff",
      padding: "12px 14px",
      textAlign: "center",
      verticalAlign: "middle",
      fontWeight: "700",
      fontSize: "0.82rem",
      border: "1px solid #334155",
      letterSpacing: "0.03em",
      textTransform: "uppercase"
    },

    tableCell: {
      padding: "10px 12px",
      border: "1px solid #cbd5e1",
      fontSize: "0.875rem",
      color: "#0f172a",
      textAlign: "center",
      verticalAlign: "middle"
    },

    tableRowEven: {
      background: "#ffffff"
    },

    tableRowOdd: {
      background: "#f8fafc"
    },

    lotCell: {
      fontWeight: "800",
      color: "#ef4444"
    },

    pcsBadge: {
      background: "#eff6ff",
      color: "#1e40af",
      padding: "4px 10px",
      borderRadius: "8px",
      fontSize: "0.8rem",
      fontWeight: "700",
      display: "inline-block",
      border: "1px solid #c7d2fe"
    },

    colorPendingStatus: {
      background: "#fef3c7",
      color: "#92400e",
      padding: "4px 12px",
      borderRadius: "20px",
      fontSize: "0.75rem",
      fontWeight: "700",
      display: "inline-block"
    },

    colorOkStatus: {
      background: "#dcfce7",
      color: "#166534",
      padding: "4px 12px",
      borderRadius: "20px",
      fontSize: "0.75rem",
      fontWeight: "700",
      display: "inline-block"
    },

    statusBadge: {
      padding: "6px 14px",
      borderRadius: "9999px",
      fontSize: "0.78rem",
      fontWeight: "700",
      display: "inline-flex",
      alignItems: "center",
      justifyContent: "center",
      gap: "6px",
      letterSpacing: "0.02em",
      whiteSpace: "nowrap",
      transition: "all 0.2s ease"
    },

    directYes: {
      background: "#dcfce7",
      color: "#15803d",
      padding: "4px 10px",
      borderRadius: "9999px",
      fontSize: "0.75rem",
      fontWeight: "700",
      display: "inline-block",
      border: "1px solid #86efac"
    },

    directNo: {
      background: "#f1f5f9",
      color: "#64748b",
      padding: "4px 10px",
      borderRadius: "9999px",
      fontSize: "0.75rem",
      fontWeight: "600",
      display: "inline-block",
      border: "1px solid #cbd5e1"
    },

    tableImage: {
      width: "40px",
      height: "40px",
      objectFit: "cover",
      borderRadius: "8px",
      border: "1px solid #e2e8f0",
      cursor: "pointer",
      boxShadow: "0 2px 6px rgba(0,0,0,0.05)"
    },

    noImagePlaceholder: {
      fontSize: "10px",
      color: "#94a3b8",
      textTransform: "uppercase"
    },

    imageModalBackdrop: {
      position: "fixed",
      top: 0,
      left: 0,
      right: 0,
      bottom: 0,
      backgroundColor: "rgba(15, 23, 42, 0.75)",
      backdropFilter: "blur(6px)",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      zIndex: 99999,
      padding: "20px"
    },

    imageModalContent: {
      position: "relative",
      background: "#ffffff",
      borderRadius: "20px",
      padding: "16px",
      maxWidth: "90vw",
      maxHeight: "90vh",
      boxShadow: "0 25px 50px -12px rgba(0, 0, 0, 0.35)",
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      justifyContent: "center",
      overflow: "hidden"
    },

    imageModalClose: {
      position: "absolute",
      top: "12px",
      right: "12px",
      background: "#ef4444",
      border: "none",
      color: "#ffffff",
      fontSize: "22px",
      fontWeight: "bold",
      borderRadius: "50%",
      width: "36px",
      height: "36px",
      cursor: "pointer",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      boxShadow: "0 4px 12px rgba(239, 68, 68, 0.4)",
      zIndex: 10
    },

    imageModalImg: {
      maxWidth: "85vw",
      maxHeight: "80vh",
      objectFit: "contain",
      borderRadius: "12px"
    },

    footer: {
      textAlign: "center",
      padding: "20px",
      fontSize: "0.85rem",
      color: "#64748b",
      background: "#ffffff",
      borderRadius: "16px",
      border: "1px solid #e2e8f0",
      marginTop: "20px"
    }
  };

  const [hoveredRow, setHoveredRow] = useState(null);
  const [hoveredButton, setHoveredButton] = useState(null);

  if (loading) {
    return (
      <div style={styles.container}>
        <style>{`
          @keyframes spinRing {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          .custom-spinner {
            width: 64px;
            height: 64px;
            border: 4px solid #e2e8f0;
            border-top: 4px solid #4f46e5;
            border-right: 4px solid #a855f7;
            border-radius: 50%;
            animation: spinRing 0.85s cubic-bezier(0.5, 0, 0.5, 1) infinite;
            box-shadow: 0 0 20px rgba(79, 70, 229, 0.25);
          }
        `}</style>
        <header style={styles.header}>
          <div style={styles.headerContentBox}>
            <div style={styles.titleSection}>
              <h1 style={styles.title}>Production Issue to Stitching after cutting</h1>
              <p style={styles.subtitle}>Pending Issues to Stitching • Real-time Monitoring</p>
            </div>
          </div>
        </header>

        <div style={{
          display: 'flex',
          flexDirection: 'column',
          alignItems: 'center',
          justifyContent: 'center',
          minHeight: '380px',
          background: '#ffffff',
          borderRadius: '24px',
          padding: '48px 32px',
          boxShadow: '0 10px 30px rgba(0, 0, 0, 0.03)',
          border: '1px solid #e2e8f0',
          textAlign: 'center',
          maxWidth: '500px',
          margin: '40px auto'
        }}>
          <div className="custom-spinner" style={{ marginBottom: '24px' }} />

          <div style={{
            fontSize: '11px',
            fontWeight: '800',
            letterSpacing: '1px',
            textTransform: 'uppercase',
            color: '#4f46e5',
            background: '#eef2ff',
            padding: '6px 16px',
            borderRadius: '20px',
            border: '1px solid #c7d2fe',
            marginBottom: '16px'
          }}>
            ⚡ MH FACTORY SUITE PRO API
          </div>

          <h3 style={{
            margin: '0 0 8px 0',
            fontSize: '1.4rem',
            fontWeight: '800',
            color: '#0f172a'
          }}>
            Fetching Live Production Data
          </h3>

          <p style={{
            margin: 0,
            fontSize: '0.9rem',
            color: '#64748b',
            lineHeight: '1.5'
          }}>
            Synchronizing Index & Cutting Matrix Sheets from Google Drive...
          </p>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div style={styles.container}>
        <div style={styles.header}>
          <div style={styles.headerContent}>
            <h1 style={styles.title}>PENDING ISSUE TO STTICHING AFTER CUTTING </h1>
            <p style={styles.subtitle}>Error loading data</p>
          </div>
        </div>
        <div style={styles.mainContent}>
          <div style={styles.emptyState}>
            <div style={{ ...styles.emptyIcon, color: '#dc2626' }}>⚠️</div>
            <h3 style={{ color: '#1e293b', marginBottom: '0.5rem' }}>Error Loading Data</h3>
            <p style={{ color: '#64748b', marginBottom: '1.5rem' }}>{error}</p>
            <button
              onClick={fetchData}
              style={{
                ...styles.button,
                ...styles.primaryButton
              }}
            >
              <span>🔄</span>
              Retry
            </button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={styles.container}>
      {/* Header */}
      <header style={styles.header}>
        <div style={styles.headerContentBox}>
          <div style={styles.titleSection}>
            <h1 style={styles.title}>Pendinng Issue to Stitching </h1>
            <p style={styles.subtitle}>Pending Issues to Stitching • Real-time Monitoring</p>
          </div>
          <div style={styles.statsCard}>
            <div style={styles.statItem}>
              <span style={styles.statNumber}>{filteredData.length}</span>
              <span style={styles.statLabel}>Total Lots</span>
            </div>
            <div style={styles.statItem}>
              <span style={styles.statNumber}>{calculateFilteredTotalPieces().toLocaleString()}</span>
              <span style={styles.statLabel}>Total Pieces</span>
            </div>
            <div style={styles.statItem}>
              <span style={styles.statNumber}>{calculateColorPendingStats().colorPendingCount}</span>
              <span style={styles.statLabel}>Color Pending</span>
            </div>
            <div style={styles.statItem}>
              <span style={styles.statNumber}>{calculateStatusStats().direct}</span>
              <span style={styles.statLabel}>Direct Lots</span>
            </div>
          </div>
        </div>
      </header>

      {/* Controls Container */}
      <div style={styles.controlsContainer}>
        <div style={styles.controlsTopRow}>
          <div style={styles.searchContainer}>
            <input
              type="text"
              placeholder="Search by lot number, fabric, brand, style, etc..."
              value={filters.search}
              onChange={(e) => handleFilterChange('search', e.target.value)}
              style={styles.searchInput}
            />
          </div>
          <div style={styles.mainActions}>
            <button onClick={handleBack} style={styles.backBtn}>
              ← Back
            </button>
            <button onClick={fetchData} style={styles.refreshBtn}>
              🔄 Refresh
            </button>
            <button
              onClick={exportToPDF}
              disabled={filteredData.length === 0}
              style={{
                ...styles.exportBtnPdf,
                ...(filteredData.length === 0 && { opacity: 0.5, cursor: 'not-allowed' })
              }}
            >
              📄 Export PDF
            </button>
            <button
              onClick={exportToExcel}
              disabled={filteredData.length === 0}
              style={{
                ...styles.exportBtnExcel,
                ...(filteredData.length === 0 && { opacity: 0.5, cursor: 'not-allowed' })
              }}
            >
              📊 Export Excel
            </button>
          </div>
        </div>

        {/* Filter Grid */}
        <div style={styles.filterSection}>
          <div style={styles.filterHeader}>
            <h3 style={styles.filterTitle}>
              🔍 Filter Options
              {Object.values(filters).some(f => Array.isArray(f) ? f.length > 0 : f) && (
                <span style={{
                  background: '#dbeafe',
                  color: '#1e40af',
                  padding: '2px 10px',
                  borderRadius: '20px',
                  fontSize: '0.8rem',
                  fontWeight: 600
                }}>
                  {filteredData.length} of {data.length} lots
                </span>
              )}
            </h3>
            {Object.values(filters).some(f => Array.isArray(f) ? f.length > 0 : f) && (
              <button onClick={clearAllFilters} style={styles.clearFiltersBtn}>
                🗑️ Clear Filters
              </button>
            )}
          </div>

          <div style={styles.filterGrid}>
            {/* Fabric Filter */}
            <div style={styles.filterGroup}>
              <label style={styles.filterLabel}>Fabric</label>
              <MultiSelectDropdown
                placeholder="All Fabrics"
                options={Array.from(filterOptions.fabric || []).sort()}
                selectedValues={filters.fabric || []}
                onChange={(val) => handleFilterChange('fabric', val)}
              />
            </div>

            {/* Garment Type Filter */}
            <div style={styles.filterGroup}>
              <label style={styles.filterLabel}>Garment Type</label>
              <MultiSelectDropdown
                placeholder="All Types"
                options={Array.from(filterOptions.garmentType || []).sort()}
                selectedValues={filters.garmentType || []}
                onChange={(val) => handleFilterChange('garmentType', val)}
              />
            </div>

            {/* Section Filter */}
            <div style={styles.filterGroup}>
              <label style={styles.filterLabel}>Section</label>
              <MultiSelectDropdown
                placeholder="All Sections"
                options={Array.from(filterOptions.section || []).sort()}
                selectedValues={filters.section || []}
                onChange={(val) => handleFilterChange('section', val)}
              />
            </div>

            {/* Season Filter */}
            <div style={styles.filterGroup}>
              <label style={styles.filterLabel}>Season</label>
              <MultiSelectDropdown
                placeholder="All Seasons"
                options={Array.from(filterOptions.season || []).sort()}
                selectedValues={filters.season || []}
                onChange={(val) => handleFilterChange('season', val)}
              />
            </div>

            {/* Style Filter */}
            <div style={styles.filterGroup}>
              <label style={styles.filterLabel}>Style</label>
              <MultiSelectDropdown
                placeholder="All Styles"
                options={Array.from(filterOptions.style || []).sort()}
                selectedValues={filters.style || []}
                onChange={(val) => handleFilterChange('style', val)}
              />
            </div>

            {/* Brand Filter */}
            <div style={styles.filterGroup}>
              <label style={styles.filterLabel}>Brand</label>
              <MultiSelectDropdown
                placeholder="All Brands"
                options={Array.from(filterOptions.brand || []).sort()}
                selectedValues={filters.brand || []}
                onChange={(val) => handleFilterChange('brand', val)}
              />
            </div>

            {/* Party Name Filter */}
            <div style={styles.filterGroup}>
              <label style={styles.filterLabel}>Party Name</label>
              <MultiSelectDropdown
                placeholder="All Parties"
                options={Array.from(filterOptions.party || []).sort()}
                selectedValues={filters.party || []}
                onChange={(val) => handleFilterChange('party', val)}
              />
            </div>

            {/* Direct Stitching Filter */}
            <div style={styles.filterGroup}>
              <label style={styles.filterLabel}>Direct Stitching</label>
              <select
                value={filters.directStitch || ''}
                onChange={(e) => handleFilterChange('directStitch', e.target.value)}
                style={styles.filterSelect}
              >
                <option value="">All</option>
                <option value="Yes">Direct (Yes)</option>
                <option value="No">Non-Direct (No)</option>
              </select>
            </div>

            {/* Status Filter */}
            <div style={styles.filterGroup}>
              <label style={styles.filterLabel}>Status</label>
              <MultiSelectDropdown
                placeholder="All Status"
                options={Array.from(filterOptions.status || []).sort()}
                selectedValues={filters.status || []}
                onChange={(val) => handleFilterChange('status', val)}
              />
            </div>

            {/* Color Status Filter */}
            <div style={styles.filterGroup}>
              <label style={styles.filterLabel}>Color Status</label>
              <select
                value={filters.colorPending || ''}
                onChange={(e) => handleFilterChange('colorPending', e.target.value)}
                style={styles.filterSelect}
              >
                <option value="">All</option>
                <option value="Yes">Color Pending</option>
                <option value="No">Color OK</option>
              </select>
            </div>

            {/* Remarks Filter */}
            <div style={styles.filterGroup}>
              <label style={styles.filterLabel}>Remarks</label>
              <select
                value={filters.remarksStatus || ''}
                onChange={(e) => handleFilterChange('remarksStatus', e.target.value)}
                style={styles.filterSelect}
              >
                <option value="">All Remarks</option>
                <option value="With Remarks">With Remarks</option>
                <option value="Without Remarks">Without Remarks</option>
              </select>
            </div>
          </div>
        </div>
      </div>

      {/* Main Content */}
      <main style={styles.mainContent}>
        {filteredData.length === 0 ? (
          <div style={styles.emptyState}>
            <div style={styles.emptyIcon}>📊</div>
            <h3 style={{ color: '#1e293b', marginBottom: '0.5rem' }}>
              {data.length === 0 ? 'No Pending Issues Found' : 'No Matching Results'}
            </h3>
            <p style={{ color: '#64748b' }}>
              {data.length === 0
                ? 'All lots have been assigned to supervisors.'
                : 'Try adjusting your filters to see more results.'
              }
            </p>
          </div>
        ) : (
          <>
            <div style={styles.tableContainer} ref={tableRef}>
              <table style={styles.table}>
                <thead>
                  <tr>
                    {columnsConfig.map((col) => (
                      <th
                        key={col.key}
                        style={{
                          ...styles.tableHeader,
                          textAlign: col.align,
                          minWidth: col.minWidth
                        }}
                      >
                        {col.label}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filteredData.map((row, index) => {
                    const isColorPending = row.hasColorPending || false;
                    const isRepeated = row.isRepeatedLot || false;
                    const pendingColors = row.pendingColors || [];
                    const pendingColorsText = row.pendingColorsText || '';
                    const statusStyle = getStatusStyle(row.status);

                    return (
                      <tr
                        key={index}
                        style={{
                          ...(index % 2 === 0 ? styles.tableRowEven : styles.tableRowOdd),
                          background: isColorPending ? '#fffbeb' :
                            isRepeated ? '#f5f3ff' :
                              row.status === 'Direct' ? '#eff6ff' : (index % 2 === 0 ? '#ffffff' : '#f8fafc'),
                          ...(hoveredRow === index && {
                            background: '#f1f5f9'
                          })
                        }}
                        onMouseEnter={() => setHoveredRow(index)}
                        onMouseLeave={() => setHoveredRow(null)}
                      >
                        {/* 1. Sr. No */}
                        <td style={{ ...styles.tableCell, textAlign: 'center', verticalAlign: 'middle', fontWeight: 600 }}>
                          {index + 1}
                          {isRepeated && <span style={{ color: '#f59e0b', marginLeft: '2px' }}>★</span>}
                        </td>

                        {/* 2. Lot Number */}
                        <td style={{ ...styles.tableCell, ...styles.lotCell, textAlign: 'center', verticalAlign: 'middle' }}>
                          {row['Lot Number']}
                          {isColorPending && <span style={{ marginLeft: '4px' }}>⚠️</span>}
                        </td>

                        {/* Garment Type */}
                        <td style={{ ...styles.tableCell, textAlign: 'center', verticalAlign: 'middle', fontWeight: 500 }}>
                          {row['Garment Type'] || 'N/A'}
                        </td>

                        {/* Style */}
                        <td style={{ ...styles.tableCell, textAlign: 'center', verticalAlign: 'middle', fontWeight: 500 }}>
                          {row['Style'] || 'N/A'}
                        </td>

                        {/* Fabric */}
                        <td style={{ ...styles.tableCell, textAlign: 'center', verticalAlign: 'middle', fontWeight: 600 }}>
                          {row['Fabric'] || 'N/A'}
                        </td>

                        {/* Brand */}
                        <td style={{ ...styles.tableCell, textAlign: 'center', verticalAlign: 'middle', fontWeight: 600, color: '#334155' }}>
                          {row['Brand'] || 'N/A'}
                        </td>

                        {/* Total Pcs */}
                        <td style={{ ...styles.tableCell, textAlign: 'center', verticalAlign: 'middle' }}>
                          <span style={styles.pcsBadge}>
                            {getTotalForLot(row['Lot Number']) || 0}
                          </span>
                        </td>

                        {/* Section */}
                        <td style={{ ...styles.tableCell, textAlign: 'center', verticalAlign: 'middle', fontWeight: 600, color: '#475569' }}>
                          {row['Section'] || row['M/W/K'] || '—'}
                        </td>

                        {/* Season */}
                        <td style={{ ...styles.tableCell, textAlign: 'center', verticalAlign: 'middle', fontWeight: 600, color: '#475569' }}>
                          {row['Season'] || 'N/A'}
                        </td>

                        {/* Party Name */}
                        <td style={{ ...styles.tableCell, textAlign: 'center', verticalAlign: 'middle', fontWeight: 500 }}>
                          {row['Party Name'] || row['Party'] || '—'}
                        </td>

                        {/* Direct Stitching */}
                        <td style={{ ...styles.tableCell, textAlign: 'center', verticalAlign: 'middle' }}>
                          {row['Direct Stitching'] && row['Direct Stitching'].toString().toLowerCase() === 'yes' ? (
                            <span style={styles.directYes}>✓ Yes</span>
                          ) : row['Direct Stitching'] && row['Direct Stitching'].toString().toLowerCase() === 'no' ? (
                            <span style={styles.directNo}>✗ No</span>
                          ) : (
                            row['Direct Stitching'] || 'N/A'
                          )}
                        </td>

                        {/* 10. Image */}
                        <td style={{ ...styles.tableCell, textAlign: 'center', verticalAlign: 'middle' }}>
                          {row['Image'] && row['Image'] !== 'N/A' ? (
                            <img
                              src={row['Image']}
                              alt="Style Preview"
                              style={styles.tableImage}
                              referrerPolicy="no-referrer"
                              onClick={(e) => {
                                e.stopPropagation();
                                setViewImageSrc(row['Image']);
                              }}
                            />
                          ) : (
                            <span style={styles.noImagePlaceholder}>No Image</span>
                          )}
                        </td>

                        {/* 11. Color Status */}
                        <td style={{ ...styles.tableCell, textAlign: 'center', verticalAlign: 'middle' }}>
                          {isColorPending ? (
                            <span
                              style={styles.colorPendingStatus}
                              title={pendingColorsText}
                            >
                              Pending: {pendingColors.slice(0, 2).join(', ')}
                              {pendingColors.length > 2 && ` +${pendingColors.length - 2}`}
                            </span>
                          ) : (
                            <span style={styles.colorOkStatus}>
                              OK
                            </span>
                          )}
                        </td>

                        {/* 12. Status */}
                        <td style={{ ...styles.tableCell, textAlign: 'center', verticalAlign: 'middle' }}>
                          <span style={{
                            ...styles.statusBadge,
                            ...statusStyle
                          }}>
                            <span style={{ fontSize: '0.85rem' }}>{statusStyle.icon}</span>
                            <span>{row.status || 'Pending'}</span>
                          </span>
                        </td>

                        {/* 13. Remarks */}
                        <td style={{ ...styles.tableCell, textAlign: 'center', verticalAlign: 'middle', minWidth: '200px', maxWidth: '260px' }}>
                          {(() => {
                            const lot = (row['Lot Number'] || row.lotNumber)?.toString().trim();
                            const lotRemarks = remarksMap[lot] || [];
                            const latestRemark = lotRemarks.length > 0 ? lotRemarks[lotRemarks.length - 1] : null;

                            return (
                              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '5px' }}>
                                {latestRemark ? (
                                  <div
                                    onClick={() => handleOpenRemarksModal(row)}
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

                                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '5px' }}>
                                  <button
                                    onClick={() => handleOpenRemarksModal(row)}
                                    style={{
                                      background: '#eef2ff',
                                      color: '#4338ca',
                                      border: '1px dashed #818cf8',
                                      borderRadius: '6px',
                                      padding: '3px 8px',
                                      fontSize: '0.72rem',
                                      fontWeight: '700',
                                      cursor: 'pointer',
                                      transition: 'all 0.15s ease'
                                    }}
                                  >
                                    {latestRemark ? '✏️ Remark' : '+ Add Remark'}
                                  </button>

                                  {lotRemarks.length > 1 && (
                                    <button
                                      onClick={() => handleOpenRemarksModal(row)}
                                      style={{
                                        background: '#ffffff',
                                        color: '#6366f1',
                                        border: '1px solid #c7d2fe',
                                        borderRadius: '6px',
                                        padding: '3px 7px',
                                        fontSize: '0.7rem',
                                        fontWeight: '700',
                                        cursor: 'pointer'
                                      }}
                                    >
                                      📜 ({lotRemarks.length})
                                    </button>
                                  )}
                                </div>
                              </div>
                            );
                          })()}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            <div style={styles.footer}>
              <p>
                Last updated: {new Date().toLocaleTimeString()} •
                Showing {filteredData.length} lots •
                Total pieces: {calculateFilteredTotalPieces().toLocaleString()} •
                Color Pending: {calculateColorPendingStats().colorPendingCount} •
                Direct Lots: {calculateStatusStats().direct}
              </p>
            </div>
          </>
        )}
      </main>

      {/* Interactive Remarks Modal */}
      {remarksModalOpen && selectedRemarksLot && (
        <div
          style={{
            position: 'fixed',
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: 'rgba(15, 23, 42, 0.65)',
            backdropFilter: 'blur(6px)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 99999,
            padding: '16px'
          }}
          onClick={handleCloseRemarksModal}
        >
          <div
            style={{
              background: '#ffffff',
              borderRadius: '24px',
              maxWidth: '560px',
              width: '100%',
              maxHeight: '88vh',
              display: 'flex',
              flexDirection: 'column',
              boxShadow: '0 25px 50px -12px rgba(15, 23, 42, 0.35)',
              border: '1px solid rgba(226, 232, 240, 0.8)',
              overflow: 'hidden'
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div
              style={{
                background: 'linear-gradient(135deg, #1e1b4b 0%, #312e81 40%, #4338ca 100%)',
                color: '#ffffff',
                padding: '18px 22px',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                borderBottom: '1px solid rgba(255, 255, 255, 0.1)'
              }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <span style={{ fontSize: '1.3rem' }}>💬</span>
                <h3 style={{ margin: 0, fontSize: '1.2rem', fontWeight: '800', color: '#ffffff' }}>
                  Stitching Issue Remarks
                </h3>
                <span style={{ background: 'rgba(255, 255, 255, 0.2)', color: '#ffd700', padding: '3px 10px', borderRadius: '12px', fontSize: '0.82rem', fontWeight: '800' }}>
                  #{selectedRemarksLot['Lot Number'] || selectedRemarksLot.lotNumber}
                </span>
              </div>
              <button
                style={{
                  background: 'rgba(255, 255, 255, 0.15)',
                  border: '1px solid rgba(255, 255, 255, 0.2)',
                  color: '#ffffff',
                  fontSize: '1.1rem',
                  borderRadius: '10px',
                  width: '32px',
                  height: '32px',
                  cursor: 'pointer',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center'
                }}
                onClick={handleCloseRemarksModal}
                title="Close"
              >
                ✕
              </button>
            </div>

            <div style={{ padding: '20px 22px', overflowY: 'auto', flex: 1, display: 'flex', flexDirection: 'column', gap: '16px' }}>
              {/* Lot Information Summary */}
              <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: '14px', padding: '12px 14px', display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '8px' }}>
                <div style={{ display: 'flex', flexDirection: 'column' }}>
                  <span style={{ fontSize: '0.68rem', fontWeight: '800', textTransform: 'uppercase', color: '#64748b' }}>Fabric</span>
                  <span style={{ fontSize: '0.84rem', fontWeight: '700', color: '#1e293b', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                    {selectedRemarksLot['Fabric'] || selectedRemarksLot.fabric || '-'}
                  </span>
                </div>
                <div style={{ display: 'flex', flexDirection: 'column' }}>
                  <span style={{ fontSize: '0.68rem', fontWeight: '800', textTransform: 'uppercase', color: '#64748b' }}>Garment / Style</span>
                  <span style={{ fontSize: '0.84rem', fontWeight: '700', color: '#1e293b', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                    {selectedRemarksLot['Garment Type'] || selectedRemarksLot['Style'] || '-'}
                  </span>
                </div>
                <div style={{ display: 'flex', flexDirection: 'column' }}>
                  <span style={{ fontSize: '0.68rem', fontWeight: '800', textTransform: 'uppercase', color: '#64748b' }}>Total Pcs</span>
                  <span style={{ fontSize: '0.84rem', fontWeight: '800', color: '#ef4444' }}>
                    {getTotalForLot(selectedRemarksLot['Lot Number']) || 0} pcs
                  </span>
                </div>
                <div style={{ display: 'flex', flexDirection: 'column' }}>
                  <span style={{ fontSize: '0.68rem', fontWeight: '800', textTransform: 'uppercase', color: '#64748b' }}>Brand</span>
                  <span style={{ fontSize: '0.84rem', fontWeight: '700', color: '#1e293b' }}>
                    {selectedRemarksLot['Brand'] || '-'}
                  </span>
                </div>
                <div style={{ display: 'flex', flexDirection: 'column' }}>
                  <span style={{ fontSize: '0.68rem', fontWeight: '800', textTransform: 'uppercase', color: '#64748b' }}>Direct</span>
                  <span style={{ fontSize: '0.84rem', fontWeight: '700', color: '#15803d' }}>
                    {selectedRemarksLot['Direct Stitching'] || '-'}
                  </span>
                </div>
                <div style={{ display: 'flex', flexDirection: 'column' }}>
                  <span style={{ fontSize: '0.68rem', fontWeight: '800', textTransform: 'uppercase', color: '#64748b' }}>Status</span>
                  <span style={{ fontSize: '0.84rem', fontWeight: '700', color: '#4338ca' }}>
                    {selectedRemarksLot.status || 'Pending'}
                  </span>
                </div>
              </div>

              {/* Remarks History */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                <h4 style={{ fontSize: '0.8rem', fontWeight: '800', color: '#475569', textTransform: 'uppercase', letterSpacing: '0.04em', margin: 0, display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <span>📜</span> Remarks History ({(remarksMap[(selectedRemarksLot['Lot Number'] || selectedRemarksLot.lotNumber)?.toString().trim()] || []).length})
                </h4>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', maxHeight: '160px', overflowY: 'auto', paddingRight: '4px' }}>
                  {(() => {
                    const lotNum = (selectedRemarksLot['Lot Number'] || selectedRemarksLot.lotNumber)?.toString().trim();
                    const history = remarksMap[lotNum] || [];
                    if (history.length === 0) {
                      return (
                        <div style={{ textAlign: 'center', padding: '14px', color: '#94a3b8', fontSize: '0.82rem' }}>
                          No previous remarks for this lot. Add the first remark below!
                        </div>
                      );
                    }
                    return history.map((item, i) => (
                      <div
                        key={i}
                        style={{
                          background: i === history.length - 1 ? '#f5f7ff' : '#ffffff',
                          border: `1px solid ${i === history.length - 1 ? '#c7d2fe' : '#e2e8f0'}`,
                          borderLeft: `4px solid ${i === history.length - 1 ? '#4338ca' : '#6366f1'}`,
                          borderRadius: '10px',
                          padding: '8px 12px',
                          display: 'flex',
                          flexDirection: 'column',
                          gap: '3px'
                        }}
                      >
                        <div style={{ fontSize: '0.85rem', color: '#0f172a', fontWeight: '600', wordBreak: 'break-word' }}>
                          {item.text}
                        </div>
                        <div style={{ fontSize: '0.7rem', color: '#64748b', fontWeight: '600', display: 'flex', alignItems: 'center', gap: '4px' }}>
                          <span>🕒</span> {item.timestamp}
                          {i === history.length - 1 && (
                            <span style={{ marginLeft: 'auto', color: '#4338ca', fontWeight: '700', fontSize: '0.68rem', background: '#e0e7ff', padding: '2px 6px', borderRadius: '4px' }}>
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
              <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                <h4 style={{ fontSize: '0.8rem', fontWeight: '800', color: '#1e1b4b', textTransform: 'uppercase', letterSpacing: '0.04em', margin: 0, display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <span>✏️</span> Enter Your Remark
                </h4>

                {/* Quick Presets */}
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: '5px' }}>
                  {[
                    "Fabric Issued to Stitching",
                    "Ready for Issue",
                    "Embroidery In Progress",
                    "Printing In Progress",
                    "Color Pending Issue",
                    "Urgent Stitching Required",
                    "Hold - Pattern Verification"
                  ].map((preset) => (
                    <button
                      key={preset}
                      type="button"
                      onClick={() => setNewRemarkInputText(preset)}
                      style={{
                        background: '#f1f5f9',
                        border: '1px solid #cbd5e1',
                        color: '#334155',
                        padding: '3px 8px',
                        borderRadius: '14px',
                        fontSize: '0.72rem',
                        fontWeight: '700',
                        cursor: 'pointer'
                      }}
                    >
                      + {preset}
                    </button>
                  ))}
                </div>

                <textarea
                  style={{
                    width: '100%',
                    minHeight: '80px',
                    padding: '10px 12px',
                    border: '1.5px solid #cbd5e1',
                    borderRadius: '12px',
                    fontFamily: 'inherit',
                    fontSize: '0.88rem',
                    color: '#0f172a',
                    boxSizing: 'border-box',
                    resize: 'vertical',
                    outline: 'none'
                  }}
                  placeholder="Type your custom remark here..."
                  value={newRemarkInputText}
                  onChange={(e) => setNewRemarkInputText(e.target.value)}
                  rows={3}
                  autoFocus
                />
              </div>
            </div>

            <div style={{ padding: '14px 22px', background: '#f8fafc', borderTop: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', justifyContent: 'flex-end', gap: '10px' }}>
              <button
                type="button"
                onClick={handleCloseRemarksModal}
                disabled={savingRemark}
                style={{
                  padding: '8px 18px',
                  background: '#ffffff',
                  color: '#475569',
                  border: '1.5px solid #cbd5e1',
                  borderRadius: '12px',
                  fontSize: '0.85rem',
                  fontWeight: '700',
                  cursor: 'pointer'
                }}
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleSaveRemark}
                disabled={savingRemark || !newRemarkInputText.trim()}
                style={{
                  padding: '8px 22px',
                  background: 'linear-gradient(135deg, #4338ca 0%, #3730a3 100%)',
                  color: '#ffffff',
                  border: 'none',
                  borderRadius: '12px',
                  fontSize: '0.85rem',
                  fontWeight: '800',
                  cursor: savingRemark || !newRemarkInputText.trim() ? 'not-allowed' : 'pointer',
                  opacity: savingRemark || !newRemarkInputText.trim() ? 0.5 : 1,
                  display: 'flex',
                  alignItems: 'center',
                  gap: '6px'
                }}
              >
                {savingRemark ? '⏳ Saving...' : '💾 Save Remark'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Image Lightbox Modal */}
      {viewImageSrc && (
        <div style={styles.imageModalBackdrop} onClick={() => setViewImageSrc(null)}>
          <div style={styles.imageModalContent} onClick={(e) => e.stopPropagation()}>
            <button style={styles.imageModalClose} onClick={() => setViewImageSrc(null)}>&times;</button>
            <img src={viewImageSrc} alt="Full Preview" style={styles.imageModalImg} />
          </div>
        </div>
      )}

      <style>
        {`
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          ::-webkit-scrollbar {
            width: 8px;
            height: 8px;
          }
          ::-webkit-scrollbar-track {
            background: #f1f5f9;
            border-radius: 4px;
          }
          ::-webkit-scrollbar-thumb {
            background: #cbd5e1;
            border-radius: 4px;
          }
          ::-webkit-scrollbar-thumb:hover {
            background: #94a3b8;
          }
        `}
      </style>
    </div>
  );
};

export default PendingIssuetoStitching;