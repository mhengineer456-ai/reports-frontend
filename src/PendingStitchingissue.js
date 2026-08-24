import React, { useState, useEffect, useRef } from 'react';
import { GOOGLE_API_KEY, SPREADSHEET_IDS, fetchSheetDataFromBackend } from './config';
import axios from 'axios';

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
    season: '',
    directStitch: '',
    colorPending: '',
    status: [],
    search: ''
  });

  const [filterOptions, setFilterOptions] = useState({
    fabric: new Set(),
    garmentType: new Set(),
    style: new Set(),
    brand: new Set(),
    season: new Set(),
    directStitch: new Set(),
    colorPending: new Set(['Yes', 'No']),
    status: new Set(['Direct', 'Ready for Stitching', 'Printing Working', 'Embroidery Working', 'Pending'])
  });

  // Color pending states
  const [colorPendingLots, setColorPendingLots] = useState({});
  const [lotPriorities, setLotPriorities] = useState({});

  const tableRef = useRef();

  const SPREADSHEET_ID = SPREADSHEET_IDS.MAIN;
  const API_KEY = GOOGLE_API_KEY;
  const RANGE = 'Index!A:AG';
  const CUTTING_SHEET_RANGE = 'Cutting!A:Z';
  const INDEX_RANGE = 'Index!A:AG';

  const columnsConfig = [
    { key: 'srNo', label: '#', align: 'center', minWidth: '45px' },
    { key: 'Lot Number', label: 'LOT #', align: 'left', minWidth: '95px' },
    { key: 'Fabric', label: 'FABRIC', align: 'left', minWidth: '140px' },
    { key: 'Garment Type', label: 'GARMENT TYPE', align: 'left', minWidth: '140px' },
    { key: 'Style', label: 'STYLE', align: 'left', minWidth: '130px' },
    { key: 'Brand', label: 'BRAND', align: 'left', minWidth: '120px' },
    { key: 'Season', label: 'SEASON', align: 'center', minWidth: '95px' },
    { key: 'Direct Stitching', label: 'DIRECT', align: 'center', minWidth: '85px' },
    { key: 'Total Pcs', label: 'TOTAL PCS', align: 'center', minWidth: '100px' },
    { key: 'Image', label: 'IMAGE', align: 'center', minWidth: '85px' },
    { key: 'Color Status', label: 'COLOR STATUS', align: 'center', minWidth: '140px' },
    { key: 'Status', label: 'STATUS', align: 'center', minWidth: '145px' }
  ];

  // Updated columns - REMOVED CHALLAN HISTORY
  const targetColumns = [
    'Lot Number',
    'Fabric',
    'Garment Type',
    'Style',
    'Brand',
    'Season',
    'Direct Stitching',
    'Total Pcs',
    'Image'
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
  }, [data, filters]);

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

      const headers = rows[0].map(h => h.trim());

      // Find indices of all target columns
      const columnIndices = {};
      targetColumns.forEach(column => {
        columnIndices[column] = headers.findIndex(header => {
          const hLower = header.toLowerCase();
          const cLower = column.toLowerCase();
          if (cLower === 'image') {
            return hLower === 'image' || hLower === 'image url' || hLower === 'imageurl' || hLower.includes('image');
          }
          return hLower === cLower;
        });
      });

      // Also find Challan History index for status calculation only
      const challanHistoryIndex = headers.findIndex(header => header.toLowerCase().includes('challan history'));

      // Find supervisor index for filtering
      const supervisorIndex = headers.findIndex(header =>
        header.toLowerCase().includes('supervisor')
      );

      if (supervisorIndex === -1) {
        throw new Error('Supervisor column not found');
      }

      // Filter and map data
      const filteredData = rows.slice(1)
        .filter(row => {
          // Check supervisor column is empty
          const supervisorEmpty = !row[supervisorIndex] || row[supervisorIndex].trim() === '';

          // Check if row has any data in target columns
          let hasData = false;
          targetColumns.forEach(column => {
            const columnIndex = columnIndices[column];
            if (columnIndex !== -1 && row[columnIndex] !== undefined && row[columnIndex].toString().trim() !== '') {
              hasData = true;
            }
          });

          return supervisorEmpty && hasData;
        })
        .map(row => {
          const rowData = {};
          targetColumns.forEach(column => {
            const columnIndex = columnIndices[column];
            if (columnIndex !== -1 && row[columnIndex] !== undefined && row[columnIndex].toString().trim() !== '') {
              let value = row[columnIndex].toString().trim();

              // Apply normalization for specific columns
              if (column === 'Brand') {
                value = normalizeBrand(value);
              } else if (column === 'Style') {
                value = normalizeStyle(value);
              } else if (column === 'Season') {
                value = normalizeSeason(value);
              } else if (column === 'Image') {
                value = getDirectImageUrl(value);
              }

              rowData[column] = value;
            } else {
              rowData[column] = 'N/A';
            }
          });

          // Add color pending information
          const lotNumber = rowData['Lot Number'];
          const pendingInfo = pendingData[lotNumber] || {};
          rowData.hasColorPending = pendingInfo.pendingColors?.length > 0 || false;
          rowData.pendingColors = pendingInfo.pendingColors || [];
          rowData.pendingColorsText = pendingInfo.pendingColors?.join(', ') || '';
          rowData.isRepeatedLot = pendingInfo.isRepeatedLot || false;
          rowData.priority = pendingInfo.priority || '';

          // Add status - FIXED: Direct Stitching "yes" = "Direct"
          const directStitching = rowData['Direct Stitching'];
          const challanHistory = challanHistoryIndex !== -1 ? row[challanHistoryIndex] : '';
          rowData.status = getStatus(directStitching, challanHistory);

          return rowData;
        });

      setData(filteredData);
      setFilteredData(filteredData);

      // Calculate statistics and filter options
      const brands = new Set();
      const seasons = new Set();
      const fabricOptions = new Set();
      const garmentTypeOptions = new Set();
      const styleOptions = new Set();
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
        if (row.Brand && row.Brand !== 'N/A') brands.add(row.Brand);
        if (row.Season && row.Season !== 'N/A') seasons.add(row.Season);
        if (row.Fabric && row.Fabric !== 'N/A') fabricOptions.add(row.Fabric);
        if (row['Garment Type'] && row['Garment Type'] !== 'N/A') garmentTypeOptions.add(row['Garment Type']);
        if (row.Style && row.Style !== 'N/A') styleOptions.add(row.Style);
        if (row['Direct Stitching'] && row['Direct Stitching'] !== 'N/A') directStitchOptions.add(row['Direct Stitching']);
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
        season: seasons,
        directStitch: directStitchOptions,
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
        if (total !== 'N/A' && !isNaN(total)) {
          totalPieces += parseInt(total);
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
    if (filters.season) {
      result = result.filter(row => row.Season === filters.season);
    }
    if (filters.directStitch) {
      result = result.filter(row => row['Direct Stitching'] === filters.directStitch);
    }

    // Apply color pending filter
    if (filters.colorPending) {
      if (filters.colorPending === 'Yes') {
        result = result.filter(row => row.hasColorPending === true);
      } else if (filters.colorPending === 'No') {
        result = result.filter(row => row.hasColorPending === false || !row.hasColorPending);
      }
    }

    // Apply status filter
    if (filters.status && filters.status.length > 0) {
      result = result.filter(row => filters.status.includes(row.status));
    }

    // Apply search filter if search term exists
    if (filters.search) {
      const searchLower = filters.search.toLowerCase();
      result = result.filter(row =>
        Object.values(row).some(value =>
          value && value.toString().toLowerCase().includes(searchLower)
        )
      );
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
      season: '',
      directStitch: '',
      colorPending: '',
      status: [],
      search: ''
    });
  };

  const getTotalPcsFromMatrix = (matrix) => {
    if (!matrix) return 'N/A';

    for (let i = 0; i < matrix.length; i++) {
      const row = matrix[i];
      if (row && row[0] && row[0].toString().toLowerCase().includes('total')) {
        const totalPcs = row[row.length - 1];
        return totalPcs || 'N/A';
      }
    }

    return 'N/A';
  };

  const getTotalForLot = (lotNumber) => {
    const matrix = matrixData[lotNumber];
    if (!matrix) return 'N/A';
    return getTotalPcsFromMatrix(matrix);
  };

  // Calculate total pieces for filtered data
  const calculateFilteredTotalPieces = () => {
    let total = 0;
    filteredData.forEach(row => {
      const matrix = matrixData[row['Lot Number']];
      if (matrix) {
        const totalPcs = getTotalPcsFromMatrix(matrix);
        if (totalPcs !== 'N/A' && !isNaN(totalPcs)) {
          total += parseInt(totalPcs);
        }
      }
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

  // Download PDF function - Fixed to show pending color names and "OK" instead of "Ready"
  // Download PDF function - Fixed to show pending color names and "OK" instead of "Ready"
  const downloadPDF = () => {
    if (filteredData.length === 0) {
      alert('No data to download');
      return;
    }

    try {
      const filteredTotalPieces = calculateFilteredTotalPieces();
      const { colorPendingCount, repeatedLotCount } = calculateColorPendingStats();
      const statusStats = calculateStatusStats();

      const htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <title>Production Status Report</title>
        <style>
          @page { 
            size: A4 landscape; 
            margin: 8mm; 
          }
          body {
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
            margin: 0;
            padding: 0;
            color: #000000;
            background-color: #ffffff;
            font-size: 10px;
          }
          
          .header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 2px solid #000000;
            padding-bottom: 8px;
            margin-bottom: 15px;
          }
          
          .title-area {
            text-align: left;
          }
          
          .title {
            font-size: 20px;
            font-weight: 800;
            color: #000000;
            margin: 0;
            letter-spacing: -0.025em;
          }
          
          .subtitle {
            font-size: 11px;
            color: #000000;
            margin: 4px 0 0 0;
            font-weight: 500;
          }
          
          .meta-area {
            text-align: right;
            font-size: 10px;
            color: #000000;
          }
          
          .meta-item {
            margin-bottom: 2px;
          }
          
          .meta-label {
            font-weight: 600;
          }
          
          /* Summary Cards */
          .summary-section {
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 12px;
            margin-bottom: 15px;
          }
          
          .summary-card {
            background: #ffffff;
            border: 1px solid #000000;
            border-radius: 6px;
            padding: 8px;
            text-align: center;
          }
          
          .summary-value {
            font-size: 14px;
            font-weight: 700;
            color: #000000;
            margin-bottom: 2px;
          }
          
          .summary-label {
            font-size: 8px;
            color: #000000;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.05em;
          }
          
          table {
            width: 100%;
            border-collapse: collapse;
            font-size: 9px;
            margin-bottom: 15px;
          }
          
          th {
            background-color: #f1f5f9;
            color: #000000;
            font-weight: 700;
            text-transform: uppercase;
            font-size: 8px;
            letter-spacing: 0.05em;
            padding: 6px;
            text-align: left;
            border: 1px solid #000000;
          }
          
          td {
            padding: 6px;
            border: 1px solid #000000;
            color: #000000;
            vertical-align: middle;
          }
          
          tr:nth-child(even) {
            background-color: #f8fafc;
          }
          
          .col-sr { width: 4%; text-align: center; }
          .col-lot { width: 9%; font-weight: 700; }
          .col-fabric { width: 11%; }
          .col-type { width: 9%; }
          .col-style { width: 9%; }
          .col-image { width: 8%; text-align: center; }
          .col-brand { width: 8%; }
          .col-season { width: 7%; }
          .col-direct { width: 6%; text-align: center; }
          .col-pcs { width: 7%; text-align: center; font-weight: 700; }
          .col-color { width: 12%; }
          .col-status { width: 10%; text-align: center; }
          
          .pcs-badge {
            font-weight: 700;
            display: inline-block;
          }
          
          .direct-yes {
            font-weight: 700;
          }
          
          .direct-no {
            font-weight: 500;
          }
          
          /* Status Badges */
          .status-badge {
            padding: 2px 6px;
            border: 1px solid #000000;
            border-radius: 4px;
            font-size: 8px;
            font-weight: 700;
            display: inline-block;
            text-align: center;
            color: #000000;
          }
          
          .color-pending {
            font-weight: 700;
            display: inline-block;
          }
          
          .color-ok {
            font-weight: 500;
          }
          
          .repeated-star {
            font-weight: 700;
            margin-right: 2px;
          }
          
          .footer {
            border-top: 1px solid #000000;
            padding-top: 8px;
            display: flex;
            justify-content: space-between;
            font-size: 8px;
            color: #000000;
          }
        </style>
      </head>
      <body>
        <div class="header">
          <div class="title-area">
            <h1 class="title">Production Status Report</h1>
            <p class="subtitle">Pending Issues to Stitching After Cutting</p>
          </div>
          <div class="meta-area">
            <div class="meta-item"><span class="meta-label">Date:</span> ${new Date().toLocaleDateString('en-GB')}</div>
            <div class="meta-item"><span class="meta-label">Time:</span> ${new Date().toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' })}</div>
          </div>
        </div>
        
        <div class="summary-section">
          <div class="summary-card">
            <div class="summary-value">${filteredData.length}</div>
            <div class="summary-label">Total Lots</div>
          </div>
          <div class="summary-card">
            <div class="summary-value">${filteredTotalPieces.toLocaleString()}</div>
            <div class="summary-label">Total Pieces</div>
          </div>
          <div class="summary-card">
            <div class="summary-value">${colorPendingCount}</div>
            <div class="summary-label">Color Pending</div>
          </div>
          <div class="summary-card">
            <div class="summary-value">${repeatedLotCount}</div>
            <div class="summary-label">Repeated Lots</div>
          </div>
          <div class="summary-card">
            <div class="summary-value">${statusStats.direct}</div>
            <div class="summary-label">Direct Lots</div>
          </div>
        </div>
        
        <table>
          <thead>
            <tr>
              <th class="col-sr">#</th>
              <th class="col-lot">LOT #</th>
              <th class="col-fabric">FABRIC</th>
              <th class="col-type">GARMENT TYPE</th>
              <th class="col-style">STYLE</th>
              <th class="col-brand">BRAND</th>
              <th class="col-season">SEASON</th>
              <th class="col-direct">DIRECT</th>
              <th class="col-pcs">TOTAL PCS</th>
              <th class="col-image">IMAGE</th>
              <th class="col-color">COLOR STATUS</th>
              <th class="col-status">STATUS</th>
            </tr>
          </thead>
          <tbody>
            ${filteredData.map((row, index) => {
        const isColorPending = row.hasColorPending || false;
        const isRepeated = row.isRepeatedLot || false;
        const pendingColors = row.pendingColors || [];
        const pendingColorsText = row.pendingColorsText || '';
        const statusClass = row.status ? row.status.toLowerCase().replace(/\s+/g, '-').replace('working', '') : 'pending';

        return `
                <tr>
                  <td class="col-sr" style="text-align: center;">${index + 1}</td>
                  <td class="col-lot">
                    ${isRepeated ? '<span class="repeated-star">★</span>' : ''}
                    ${row['Lot Number'] || 'N/A'}
                  </td>
                  <td>${row['Fabric'] || 'N/A'}</td>
                  <td>${row['Garment Type'] || 'N/A'}</td>
                  <td>${row['Style'] || 'N/A'}</td>
                  <td>${row['Brand'] || 'N/A'}</td>
                  <td style="text-align: center;">${row['Season'] || 'N/A'}</td>
                  <td class="col-direct" style="text-align: center;">
                    ${row['Direct Stitching'] && row['Direct Stitching'].toString().toLowerCase() === 'yes' ?
            '<span class="direct-yes">Yes</span>' :
            row['Direct Stitching'] && row['Direct Stitching'].toString().toLowerCase() === 'no' ?
              '<span class="direct-no">No</span>' :
              row['Direct Stitching'] || 'N/A'
          }
                  </td>
                  <td class="col-pcs" style="text-align: center;">
                    <span class="pcs-badge">${getTotalForLot(row['Lot Number']) || 0}</span>
                  </td>
                  <td class="col-image" style="text-align: center;">
                    ${row['Image'] && row['Image'] !== 'N/A'
            ? `<img src="${row['Image']}" referrerpolicy="no-referrer" style="width: 35px; height: 35px; object-fit: cover; border: 1px solid #000000; border-radius: 3px;" />`
            : '<span style="font-size: 8px; color: #666;">No Image</span>'}
                  </td>
                  <td class="col-color" style="text-align: center;">
                    ${isColorPending ?
            `<span class="color-pending" title="${pendingColorsText}">⚠️ Pending: ${pendingColors.slice(0, 2).join(', ')}${pendingColors.length > 2 ? '...' : ''}</span>` :
            '<span class="color-ok">✓ OK</span>'
          }
                  </td>
                  <td class="col-status" style="text-align: center;">
                    <span class="status-badge">
                      ${row.status || 'Pending'}
                    </span>
                  </td>
                </tr>
              `;
      }).join('')}
          </tbody>
        </table>
        
        <div class="footer">
          <div>System Generated Report | Confidential Production Data</div>
          <div>Direct: ${statusStats.direct} | Ready: ${statusStats.ready} | Printing: ${statusStats.printing} | Embroidery: ${statusStats.embroidery} | Pending: ${statusStats.pending}</div>
        </div>
      </body>
      </html>
    `;

      const blob = new Blob([htmlContent], { type: 'text/html' });
      const url = URL.createObjectURL(blob);
      const iframe = document.createElement('iframe');
      iframe.style.display = 'none';
      iframe.src = url;
      document.body.appendChild(iframe);

      iframe.onload = function () {
        setTimeout(() => {
          iframe.contentWindow.focus();
          iframe.contentWindow.print();
          setTimeout(() => {
            document.body.removeChild(iframe);
            URL.revokeObjectURL(url);
          }, 1000);
        }, 500);
      };

    } catch (error) {
      console.error('PDF Error:', error);
      alert('Error generating PDF.');
    }
  };
  const handleBack = () => {
    window.history.back();
  };

  const downloadCSV = () => {
    if (filteredData.length === 0) {
      alert('No data to download');
      return;
    }

    try {
      const { colorPendingCount, repeatedLotCount } = calculateColorPendingStats();
      const statusStats = calculateStatusStats();

      // Headers without Challan History
      const headers = ['#', 'Lot Number', 'Fabric', 'Garment Type', 'Style', 'Brand', 'Season', 'Direct Stitching', 'Total Pcs', 'Color Status', 'Pending Colors', 'Status', 'Priority'];
      const csvData = filteredData.map((row, index) => {
        const rowData = {
          '#': index + 1,
          'Lot Number': row['Lot Number'],
          'Fabric': row['Fabric'],
          'Garment Type': row['Garment Type'],
          'Style': row['Style'],
          'Brand': row['Brand'],
          'Season': row['Season'],
          'Direct Stitching': row['Direct Stitching'],
          'Total Pcs': getTotalForLot(row['Lot Number']),
          'Color Status': row.hasColorPending ? 'Color Pending' : 'OK',
          'Pending Colors': row.pendingColorsText || '',
          'Status': row.status || '',
          'Priority': row.priority || ''
        };

        return headers.map(column => {
          const value = rowData[column] || '';
          const escapedValue = String(value).replace(/"/g, '""');
          return escapedValue.includes(',') ? `"${escapedValue}"` : escapedValue;
        }).join(',');
      }).join('\n');

      // Add summary rows
      const summaryRows = [
        '',
        'SUMMARY',
        `Total Lots,${filteredData.length}`,
        `Total Pieces,${calculateFilteredTotalPieces()}`,
        `Color Pending Lots,${colorPendingCount}`,
        `Repeated Lots,${repeatedLotCount}`,
        'Status Counts',
        `Direct,${statusStats.direct}`,
        `Ready for Stitching,${statusStats.ready}`,
        `Printing Working,${statusStats.printing}`,
        `Embroidery Working,${statusStats.embroidery}`,
        `Pending,${statusStats.pending}`
      ].join('\n');

      const csvContent = `${headers.join(',')}\n${csvData}\n${summaryRows}`;
      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      const link = document.createElement('a');
      const url = URL.createObjectURL(blob);

      link.setAttribute('href', url);
      link.setAttribute('download', `Production_Status_${new Date().toISOString().split('T')[0]}.csv`);
      link.style.visibility = 'hidden';

      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (error) {
      console.error('Error generating CSV:', error);
      alert('Error generating CSV. Please try again.');
    }
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
      borderRadius: "20px",
      overflow: "hidden",
      boxShadow: "0 10px 30px rgba(0, 0, 0, 0.03)",
      marginBottom: "20px",
      overflowX: "auto",
      border: "1px solid #e2e8f0"
    },

    table: {
      width: "100%",
      borderCollapse: "collapse",
      minWidth: "1200px"
    },

    tableHeader: {
      background: "linear-gradient(135deg, #1e1b4b 0%, #312e81 100%)",
      color: "#ffffff",
      padding: "14px 16px",
      textAlign: "center",
      fontWeight: "700",
      fontSize: "0.82rem",
      borderRight: "1px solid rgba(255, 255, 255, 0.1)",
      letterSpacing: "0.03em",
      textTransform: "uppercase"
    },

    tableCell: {
      padding: "12px 16px",
      borderBottom: "1px solid #e2e8f0",
      fontSize: "0.875rem",
      color: "#0f172a"
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
              onClick={downloadPDF}
              disabled={filteredData.length === 0}
              style={{
                ...styles.exportBtnPdf,
                ...(filteredData.length === 0 && { opacity: 0.5, cursor: 'not-allowed' })
              }}
            >
              📄 Export PDF
            </button>
            <button
              onClick={downloadCSV}
              disabled={filteredData.length === 0}
              style={{
                ...styles.exportBtnExcel,
                ...(filteredData.length === 0 && { opacity: 0.5, cursor: 'not-allowed' })
              }}
            >
              📊 Export CSV
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
                options={Array.from(filterOptions.fabric).sort()}
                selectedValues={filters.fabric}
                onChange={(val) => handleFilterChange('fabric', val)}
              />
            </div>

            {/* Garment Type Filter */}
            <div style={styles.filterGroup}>
              <label style={styles.filterLabel}>Garment Type</label>
              <MultiSelectDropdown
                placeholder="All Types"
                options={Array.from(filterOptions.garmentType).sort()}
                selectedValues={filters.garmentType}
                onChange={(val) => handleFilterChange('garmentType', val)}
              />
            </div>

            {/* Style Filter */}
            <div style={styles.filterGroup}>
              <label style={styles.filterLabel}>Style</label>
              <MultiSelectDropdown
                placeholder="All Styles"
                options={Array.from(filterOptions.style).sort()}
                selectedValues={filters.style}
                onChange={(val) => handleFilterChange('style', val)}
              />
            </div>

            {/* Brand Filter */}
            <div style={styles.filterGroup}>
              <label style={styles.filterLabel}>Brand</label>
              <MultiSelectDropdown
                placeholder="All Brands"
                options={Array.from(filterOptions.brand).sort()}
                selectedValues={filters.brand}
                onChange={(val) => handleFilterChange('brand', val)}
              />
            </div>

            {/* Status Filter */}
            <div style={styles.filterGroup}>
              <label style={styles.filterLabel}>Status</label>
              <MultiSelectDropdown
                placeholder="All Status"
                options={Array.from(filterOptions.status).sort()}
                selectedValues={filters.status}
                onChange={(val) => handleFilterChange('status', val)}
              />
            </div>

            {/* Color Status Filter */}
            <div style={styles.filterGroup}>
              <label style={styles.filterLabel}>Color Status</label>
              <select
                value={filters.colorPending}
                onChange={(e) => handleFilterChange('colorPending', e.target.value)}
                style={styles.filterSelect}
              >
                <option value="">All</option>
                <option value="Yes">Color Pending</option>
                <option value="No">Color OK</option>
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
                        <td style={{ ...styles.tableCell, textAlign: 'center', fontWeight: 600 }}>
                          {index + 1}
                          {isRepeated && <span style={{ color: '#f59e0b', marginLeft: '2px' }}>★</span>}
                        </td>

                        {/* 2. Lot Number */}
                        <td style={{ ...styles.tableCell, ...styles.lotCell, textAlign: 'left' }}>
                          {row['Lot Number']}
                          {isColorPending && <span style={{ marginLeft: '4px' }}>⚠️</span>}
                        </td>

                        {/* 3. Fabric */}
                        <td style={{ ...styles.tableCell, textAlign: 'left', fontWeight: 600 }}>
                          {row['Fabric']}
                        </td>

                        {/* 4. Garment Type */}
                        <td style={{ ...styles.tableCell, textAlign: 'left', fontWeight: 500 }}>
                          {row['Garment Type']}
                        </td>

                        {/* 5. Style */}
                        <td style={{ ...styles.tableCell, textAlign: 'left', fontWeight: 500 }}>
                          {row['Style']}
                        </td>

                        {/* 6. Brand */}
                        <td style={{ ...styles.tableCell, textAlign: 'left', fontWeight: 600, color: '#334155' }}>
                          {row['Brand']}
                        </td>

                        {/* 7. Season */}
                        <td style={{ ...styles.tableCell, textAlign: 'center', fontWeight: 600, color: '#475569' }}>
                          {row['Season']}
                        </td>

                        {/* 8. Direct Stitching */}
                        <td style={{ ...styles.tableCell, textAlign: 'center' }}>
                          {row['Direct Stitching'] && row['Direct Stitching'].toString().toLowerCase() === 'yes' ? (
                            <span style={styles.directYes}>✓ Yes</span>
                          ) : row['Direct Stitching'] && row['Direct Stitching'].toString().toLowerCase() === 'no' ? (
                            <span style={styles.directNo}>✗ No</span>
                          ) : (
                            row['Direct Stitching']
                          )}
                        </td>

                        {/* 9. Total Pcs */}
                        <td style={{ ...styles.tableCell, textAlign: 'center' }}>
                          <span style={styles.pcsBadge}>
                            {getTotalForLot(row['Lot Number'])}
                          </span>
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
                        <td style={{ ...styles.tableCell, textAlign: 'center' }}>
                          {isColorPending ? (
                            <span
                              style={styles.colorPendingStatus}
                              title={pendingColorsText}
                            >
                              ⚠️ Pending: {pendingColors.slice(0, 2).join(', ')}
                              {pendingColors.length > 2 && ` +${pendingColors.length - 2}`}
                            </span>
                          ) : (
                            <span style={styles.colorOkStatus}>
                              ✓ OK
                            </span>
                          )}
                        </td>

                        {/* 12. Status */}
                        <td style={{ ...styles.tableCell, textAlign: 'center' }}>
                          <span style={{
                            ...styles.statusBadge,
                            ...statusStyle
                          }}>
                            <span style={{ fontSize: '0.85rem' }}>{statusStyle.icon}</span>
                            <span>{row.status || 'Pending'}</span>
                          </span>
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