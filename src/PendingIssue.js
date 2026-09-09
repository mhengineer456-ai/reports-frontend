import React, { useState, useEffect, useRef } from 'react';
import { GOOGLE_API_KEY, SPREADSHEET_IDS, fetchSheetDataFromBackend } from './config';

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
          padding: '10px 12px',
          borderRadius: '8px',
          background: 'white',
          border: '1px solid #e2e8f0',
          color: '#1e293b',
          fontSize: '14px',
          cursor: 'pointer',
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          minHeight: '41px',
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

const PendingIssue = ({ onBack }) => {
  const [issues, setIssues] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [filteredIssues, setFilteredIssues] = useState([]);
  const [searchTerm, setSearchTerm] = useState('');
  const [activeFilter, setActiveFilter] = useState('all');
  const [showExportMenu, setShowExportMenu] = useState(false);
  const [cuttingData, setCuttingData] = useState({});
  const [selectedIssue, setSelectedIssue] = useState(null);
  const [viewMode, setViewMode] = useState('table');
  const [viewImageSrc, setViewImageSrc] = useState(null);

  // New dropdown filter states
  const [brandFilter, setBrandFilter] = useState([]);
  const [fabricFilter, setFabricFilter] = useState([]);
  const [styleFilter, setStyleFilter] = useState([]);
  const [partyFilter, setPartyFilter] = useState([]);
  const [seasonFilter, setSeasonFilter] = useState([]); // Season filter
  const [daysFilter, setDaysFilter] = useState('');
  const [garmentTypeFilter, setGarmentTypeFilter] = useState([]);
  const [mwkFilter, setMwkFilter] = useState([]); // M/W/K filter

  // Available filter options
  const [availableBrands, setAvailableBrands] = useState([]);
  const [availableFabrics, setAvailableFabrics] = useState([]);
  const [availableStyles, setAvailableStyles] = useState([]);
  const [availableParties, setAvailableParties] = useState([]);
  const [availableSeasons, setAvailableSeasons] = useState([]); // Season options
  const [availableGarmentTypes, setAvailableGarmentTypes] = useState([]);
  const [availableMwkValues, setAvailableMwkValues] = useState([]); // M/W/K options

  // Color pending and priority states (NEW)
  const [colorPendingLots, setColorPendingLots] = useState({});
  const [lotPriorities, setLotPriorities] = useState({});

  const tableRef = useRef(null);

  // API Configuration
  const API_KEY = GOOGLE_API_KEY;
  const SPREADSHEET_ID = SPREADSHEET_IDS.MAIN;
  const SHEET_NAME = 'Index';
  const RANGE = 'A:AG';
  const CUTTING_SHEET_NAME = 'Cutting';
  const CUTTING_RANGE = 'A:K';
  const INDEX_RANGE = 'Index!A:AG'; // For color pending check

  useEffect(() => {
    fetchData();
  }, []);

  useEffect(() => {
    // Populate filter options when issues change
    if (issues.length > 0) {
      const brands = [...new Set(issues.map(issue => issue.brand).filter(b => b && b !== 'N/A'))].sort();
      const fabrics = [...new Set(issues.map(issue => issue.fabric).filter(f => f && f !== 'N/A'))].sort();
      const stylesList = [...new Set(issues.map(issue => issue.style).filter(s => s && s !== 'N/A'))].sort();
      const parties = [...new Set(issues.map(issue => issue.partyName).filter(p => p && p !== 'N/A'))].sort();
      const seasons = [...new Set(issues.map(issue => issue.season).filter(s => s && s !== 'N/A' && s !== ''))].sort();
      const garmentTypes = [...new Set(issues.map(issue => issue.garmentType).filter(g => g && g !== 'N/A'))].sort();
      const mwkValues = [...new Set(issues.map(issue => issue.mwk).filter(m => m && m !== 'N/A' && m !== ''))].sort();

      setAvailableBrands(brands);
      setAvailableFabrics(fabrics);
      setAvailableStyles(stylesList);
      setAvailableParties(parties);
      setAvailableSeasons(seasons);
      setAvailableGarmentTypes(garmentTypes);
      setAvailableMwkValues(mwkValues);
    }
  }, [issues]);

  useEffect(() => {
    let result = issues;

    if (searchTerm) {
      const term = searchTerm.toLowerCase();
      result = result.filter(issue =>
        issue.lotNumber.toLowerCase().includes(term) ||
        issue.fabric.toLowerCase().includes(term) ||
        issue.garmentType.toLowerCase().includes(term) ||
        issue.brand.toLowerCase().includes(term) ||
        issue.partyName.toLowerCase().includes(term) ||
        (issue.season && issue.season.toLowerCase().includes(term)) ||
        issue.mwk.toLowerCase().includes(term)
      );
    }

    switch (activeFilter) {
      case 'emb':
        result = result.filter(issue => issue.hasCompletedEmbChallans);
        break;
      case 'direct':
        result = result.filter(issue => issue.directStitching === 'yes');
        break;
      case 'both':
        result = result.filter(issue =>
          issue.hasCompletedEmbChallans && issue.directStitching === 'yes'
        );
        break;
      default:
        break;
    }

    // Apply dropdown filters
    if (brandFilter && brandFilter.length > 0) {
      result = result.filter(issue => brandFilter.includes(issue.brand));
    }

    if (fabricFilter && fabricFilter.length > 0) {
      result = result.filter(issue => fabricFilter.includes(issue.fabric));
    }

    if (styleFilter && styleFilter.length > 0) {
      result = result.filter(issue => styleFilter.includes(issue.style));
    }

    if (partyFilter && partyFilter.length > 0) {
      result = result.filter(issue => partyFilter.includes(issue.partyName));
    }

    if (seasonFilter && seasonFilter.length > 0) {
      result = result.filter(issue => seasonFilter.includes(issue.season));
    }

    if (garmentTypeFilter && garmentTypeFilter.length > 0) {
      result = result.filter(issue => garmentTypeFilter.includes(issue.garmentType));
    }

    if (mwkFilter && mwkFilter.length > 0) {
      result = result.filter(issue => mwkFilter.includes(issue.mwk));
    }

    // Apply days pending filter
    if (daysFilter) {
      switch (daysFilter) {
        case '7':
          result = result.filter(issue => issue.daysPending > 7);
          break;
        case '14':
          result = result.filter(issue => issue.daysPending > 14);
          break;
        case '30':
          result = result.filter(issue => issue.daysPending > 30);
          break;
        case 'urgent':
          result = result.filter(issue => issue.daysPending > 7 && issue.daysPending <= 14);
          break;
        case 'critical':
          result = result.filter(issue => issue.daysPending > 14);
          break;
        default:
          break;
      }
    }

    setFilteredIssues(result);
  }, [searchTerm, activeFilter, issues, brandFilter, fabricFilter, styleFilter, partyFilter, seasonFilter, garmentTypeFilter, mwkFilter, daysFilter]);

  // Clear all filters
  const clearAllFilters = () => {
    setBrandFilter([]);
    setFabricFilter([]);
    setStyleFilter([]);
    setPartyFilter([]);
    setSeasonFilter([]);
    setGarmentTypeFilter([]);
    setMwkFilter([]);
    setDaysFilter('');
    setActiveFilter('all');
    setSearchTerm('');
  };

  // Helper functions for color pending detection (SAME AS STITCHING COMPONENT)
  const normalizeKey = (s = "") => {
    return String(s || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
  };

  const parseIndexRow = (header, row) => {
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
    const fabric = get("fabric");
    const garmentType = get("garmenttype") || get("garment");
    const style = get("style");
    const savedAt = get("savedat");

    // SHADES FIELD FOR COLOR PENDING CHECK
    const sizes = String(get("sizes") || "")
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean);

    const shades = String(get("shades") || "")
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean);

    // PRIORITY FIELD - HANDLE TYPO "Prioirty"
    const priority = get("priority") || get("prioirty") || "";

    return { lot, startRow, numRows, fabric, garmentType, style, sizes, shades, savedAt, priority };
  };

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

    // NON-SIZE COLUMNS
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

    // Determine pending shades
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

  // Fetch index data for color pending and priority
  const fetchIndexData = async () => {
    try {
      const res = await fetchSheetDataFromBackend(SPREADSHEET_ID, INDEX_RANGE);
      return res.ok ? res.values : [];
    } catch (err) {
      console.error('Error fetching index data:', err);
      return [];
    }
  };

  // Check color pending for all lots by fetching cutting data
  const checkColorPendingForLots = async () => {
    try {
      // Fetch cutting data and index data in parallel via backend
      const [cuttingRes, indexRes] = await Promise.all([
        fetchSheetDataFromBackend(SPREADSHEET_ID, "Cutting!A:Z"),
        fetchSheetDataFromBackend(SPREADSHEET_ID, "Index!A:AG")
      ]);

      const cuttingRows = cuttingRes.ok ? cuttingRes.values : [];
      const indexRows = indexRes.ok ? indexRes.values : [];

      if (indexRows.length === 0) {
        return {};
      }

      const pendingMap = {};
      const priorityMap = {};

      // Process each lot from index sheet
      for (let i = 1; i < indexRows.length; i++) {
        const row = indexRows[i];

        // Find column indices dynamically
        const headers = indexRows[0];
        const headerIndices = {};
        headers.forEach((header, index) => {
          if (header) {
            headerIndices[normalizeKey(header)] = index;
          }
        });

        // Get lot number - try different possible column names
        const lotNumber = (row[headerIndices['lotnumber']] ||
          row[headerIndices['lot no']] ||
          row[headerIndices['lot']] || '').toString().trim();

        if (!lotNumber || lotNumber === '') continue;

        // Get start row and num rows
        const startRow = parseInt(row[headerIndices['startrow']] || '0', 10);
        const numRows = parseInt(row[headerIndices['numrows']] || '0', 10);

        // Get sizes and shades
        const sizesStr = row[headerIndices['sizes']] || '';
        const shadesStr = row[headerIndices['shades']] || '';

        const sizes = sizesStr.split(',')
          .map(s => s.trim())
          .filter(Boolean);

        const shades = shadesStr.split(',')
          .map(s => s.trim())
          .filter(Boolean);

        // Get priority - check multiple possible column names
        const priority = (row[headerIndices['priority']] ||
          row[headerIndices['prioirty']] ||
          row[headerIndices['special']] || '').toString().trim();

        // Store priority
        priorityMap[lotNumber] = priority;

        // Initialize entry
        pendingMap[lotNumber] = {
          pendingColors: [],
          priority: priority,
          isRepeatedLot: priority && priority.toLowerCase().includes('repeated')
        };

        // Only check for color pending if we have shades data and valid cutting matrix info
        if (shades && shades.length > 0 && startRow > 0 && numRows > 0) {
          // Extract cutting window for this lot
          const window = sliceCuttingMatrix(cuttingRows, startRow, numRows);

          if (window.length > 0) {
            const pendingShadeKeys = computePendingShades(window, sizes, shades);

            if (pendingShadeKeys.size > 0) {
              // Map shade keys back to original shade names
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

      // Set both states
      setColorPendingLots(pendingMap);
      setLotPriorities(priorityMap);

      console.log('Priority Map:', priorityMap);
      console.log('Repeated lots found:', Object.keys(priorityMap).filter(lot => priorityMap[lot] && priorityMap[lot].toLowerCase().includes('repeated')));

      return pendingMap;

    } catch (err) {
      console.error('Error checking color pending:', err);
      return {};
    }
  };

  // Parse date from Saved At column
  const parseSavedAtDate = (savedAtValue) => {
    console.log('Raw savedAtValue:', savedAtValue);

    if (!savedAtValue || savedAtValue === '' || savedAtValue === 'N/A' || savedAtValue === 'undefined') {
      return null;
    }

    try {
      const dateStr = savedAtValue.toString().trim();

      // Try multiple date parsing strategies
      const directDate = new Date(dateStr);
      if (!isNaN(directDate.getTime())) {
        console.log('Direct parse successful:', directDate);
        return directDate;
      }

      const regex1 = /(\w{3})\s+(\w{3})\s+(\d{1,2})\s+(\d{4})/;
      const match1 = dateStr.match(regex1);

      if (match1) {
        const [, , monthStr, dayStr, yearStr] = match1;
        const monthMap = {
          'Jan': 0, 'Feb': 1, 'Mar': 2, 'Apr': 3, 'May': 4, 'Jun': 5,
          'Jul': 6, 'Aug': 7, 'Sep': 8, 'Oct': 9, 'Nov': 10, 'Dec': 11
        };

        const month = monthMap[monthStr];
        if (month !== undefined) {
          const date = new Date(parseInt(yearStr), month, parseInt(dayStr));
          if (!isNaN(date.getTime())) {
            console.log('Regex1 parse successful:', date);
            return date;
          }
        }
      }

      const regex2 = /(\d{1,2})\/(\d{1,2})\/(\d{4})/;
      const match2 = dateStr.match(regex2);

      if (match2) {
        const [, month, day, year] = match2;
        const date = new Date(year, parseInt(month) - 1, day);
        if (!isNaN(date.getTime())) {
          console.log('Regex2 parse successful:', date);
          return date;
        }
      }

      if (dateStr.includes('T')) {
        const isoDate = new Date(dateStr);
        if (!isNaN(isoDate.getTime())) {
          console.log('ISO parse successful:', isoDate);
          return isoDate;
        }
      }

      console.log('All parsing strategies failed for:', savedAtValue);
      return null;

    } catch (e) {
      console.error('Error parsing date:', e, 'Value:', savedAtValue);
      return null;
    }
  };

  // Get last EMB date from challan history
  const getLastEmbDate = (challans) => {
    if (!challans || !Array.isArray(challans) || challans.length === 0) {
      return null;
    }

    try {
      const completedChallans = challans.filter(challan =>
        challan.embCompleted === true && challan.embUpdatedAt
      );

      if (completedChallans.length === 0) {
        return null;
      }

      const sortedChallans = [...completedChallans].sort((a, b) => {
        try {
          const dateA = new Date(a.embUpdatedAt);
          const dateB = new Date(b.embUpdatedAt);
          if (isNaN(dateA.getTime()) || isNaN(dateB.getTime())) {
            return 0;
          }
          return dateB - dateA;
        } catch (e) {
          return 0;
        }
      });

      const latestDate = sortedChallans[0]?.embUpdatedAt;
      if (latestDate) {
        const date = new Date(latestDate);
        return isNaN(date.getTime()) ? null : date;
      }

      return null;
    } catch (e) {
      console.error('Error getting last EMB date:', e);
      return null;
    }
  };

  // Format date for display
  const formatDateDisplay = (date, rawValue = '', issue = null) => {
    // If it's an Emb/Print date and there's no date but direct stitching is 'yes'
    if (issue && issue.directStitching === 'yes' && !date) {
      return 'Direct';
    }

    if (!date) {
      if (rawValue && rawValue !== '') {
        const dateMatch = rawValue.toString().match(/(\w{3}\s+\w{3}\s+\d{1,2}\s+\d{4})/);
        return dateMatch ? dateMatch[1] : rawValue.substring(0, 15);
      }
      return 'N/A';
    }

    try {
      const d = date instanceof Date ? date : new Date(date);
      if (isNaN(d.getTime())) {
        return 'N/A';
      }

      return d.toLocaleDateString('en-IN', {
        day: '2-digit',
        month: 'short',
        year: 'numeric'
      });
    } catch (e) {
      return 'N/A';
    }
  };

  // Calculate days difference
  const calculateDaysDifference = (directStitching, cuttingDate, lastEmbDate) => {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    if (directStitching === 'yes') {
      if (cuttingDate) {
        try {
          const cuttingDateObj = cuttingDate instanceof Date ?
            cuttingDate : new Date(cuttingDate);

          if (isNaN(cuttingDateObj.getTime())) {
            return null;
          }

          cuttingDateObj.setHours(0, 0, 0, 0);
          const diffTime = today - cuttingDateObj;
          const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
          return Math.max(0, diffDays);
        } catch (e) {
          console.error('Error calculating cutting days:', e);
          return null;
        }
      }
      return null;
    } else {
      if (lastEmbDate) {
        try {
          const embDate = lastEmbDate instanceof Date ?
            lastEmbDate : new Date(lastEmbDate);

          if (isNaN(embDate.getTime())) {
            return null;
          }

          embDate.setHours(0, 0, 0, 0);
          const diffTime = today - embDate;
          const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
          return Math.max(0, diffDays);
        } catch (e) {
          console.error('Error calculating EMB days:', e);
          return null;
        }
      }
      return null;
    }
  };

  const fetchCuttingData = async () => {
    try {
      const res = await fetchSheetDataFromBackend(SPREADSHEET_ID, `${CUTTING_SHEET_NAME}!${CUTTING_RANGE}`);
      const rows = res.ok ? res.values : [];

      if (!rows || rows.length === 0) {
        return {};
      }

      const matrixMap = {};
      let currentLot = null;
      let currentMatrix = [];

      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];

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
            for (let j = i; j < Math.min(i + 10, rows.length); j++) {
              if (rows[j] && rows[j].length > 0) {
                currentMatrix.push(rows[j]);
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

      return matrixMap;
    } catch (err) {
      console.error('Error fetching cutting data:', err);
      return {};
    }
  };

  const getTotalPcsFromMatrix = (matrix) => {
    if (!matrix) return 0;

    for (let i = 0; i < matrix.length; i++) {
      const row = matrix[i];
      if (row && row[0] && row[0].toString().toLowerCase().includes('total')) {
        const totalPcs = row[row.length - 1];
        return parseInt(totalPcs) || 0;
      }
    }

    return 0;
  };

  const fetchData = async () => {
    try {
      setLoading(true);
      setError(null);

      // Fetch all required data in parallel
      const [indexRes, pendingData, cuttingMatrixData] = await Promise.all([
        fetchSheetDataFromBackend(SPREADSHEET_ID, `${SHEET_NAME}!${RANGE}`),
        checkColorPendingForLots(),
        fetchCuttingData()
      ]);

      if (!indexRes.ok) {
        throw new Error(`Failed to fetch index data`);
      }

      const rows = indexRes.values || [];

      if (!rows || rows.length === 0) {
        setIssues([]);
        setFilteredIssues([]);
        setLoading(false);
        return;
      }

      const headers = rows[0];

      console.log('All headers:', headers.map((h, i) => `${i}: "${h}"`).join(', '));

      let savedAtIndex = -1;
      // First pass: exact match for date/cutting date/saved at
      headers.forEach((header, index) => {
        const headerLower = header.trim().toLowerCase();
        if (headerLower === 'saved at' || headerLower === 'cutting date' || headerLower === 'date' || headerLower === 'savedat') {
          console.log(`Found exact date column: ${index}: "${header}"`);
          savedAtIndex = index;
        }
      });

      // Second pass: fallback search if no exact match found
      if (savedAtIndex === -1) {
        headers.forEach((header, index) => {
          const headerLower = header.trim().toLowerCase();
          if ((headerLower.includes('saved') || headerLower.includes('cutting')) &&
            !headerLower.includes('pcs') &&
            !headerLower.includes('pieces') &&
            !headerLower.includes('qty') &&
            !headerLower.includes('total') &&
            !headerLower.includes('matrix')) {
            console.log(`Found fallback cutting date column: ${index}: "${header}"`);
            savedAtIndex = index;
          }
        });
      }

      if (savedAtIndex === -1) {
        console.warn('Could not find "Saved at" column. Trying default index...');
        savedAtIndex = headers.length > 16 ? 16 : (headers.length - 1);
      }

      let imageUrlIndex = -1;
      headers.forEach((header, index) => {
        const headerLower = header.trim().toLowerCase();
        if (headerLower === 'image url' || headerLower === 'imageurl' || headerLower === 'image' || headerLower.includes('image')) {
          console.log(`Found image URL column: ${index}: "${header}"`);
          imageUrlIndex = index;
        }
      });

      console.log(`Using column index ${savedAtIndex} for cutting date`);

      // Find M/W/K column index
      let mwkIndex = -1;
      headers.forEach((header, index) => {
        const headerLower = header.trim().toLowerCase();
        // Look for M/W/K variations
        if (headerLower.includes('m/w/k') || headerLower.includes('mwk') || headerLower.includes('m w k')) {
          console.log(`Found M/W/K column: ${index}: "${header}"`);
          mwkIndex = index;
        }
      });

      // If not found, try other common names
      if (mwkIndex === -1) {
        headers.forEach((header, index) => {
          const headerLower = header.trim().toLowerCase();
          if (headerLower === 'gender' || headerLower === 'type' || headerLower === 'category') {
            console.log(`Found alternative M/W/K column: ${index}: "${header}"`);
            mwkIndex = index;
          }
        });
      }

      console.log(`Using M/W/K column index: ${mwkIndex}`);

      // Find Season column index
      let seasonIndex = -1;
      headers.forEach((header, index) => {
        const headerLower = header.trim().toLowerCase();
        if (headerLower === 'season' || headerLower.includes('season')) {
          console.log(`Found Season column: ${index}: "${header}"`);
          seasonIndex = index;
        }
      });

      console.log(`Using Season column index: ${seasonIndex}`);

      const headerIndices = {};
      headers.forEach((header, index) => {
        headerIndices[header.trim()] = index;
      });

      const formattedData = rows.slice(1).map((row, rowIndex) => {
        if (rowIndex < 3) {
          console.log(`Row ${rowIndex} cutting date value:`, row[savedAtIndex]);
        }

        const supervisorValue = row[headerIndices['Supervisor']] || '';
        if (supervisorValue.trim() !== '') {
          return null;
        }

        const directStitchingValue = (row[headerIndices['DIRECT STITCHING']] || '').toLowerCase().trim();

        let challanHistory = [];
        try {
          const challanStr = row[headerIndices['CHALLAN HISTORY']] || '[]';
          challanHistory = JSON.parse(challanStr);
        } catch (e) {
          console.error('Error parsing challan history:', e);
          challanHistory = [];
        }

        // Check if ALL challans in the history are embCompleted true
        const hasCompletedEmbChallans = (() => {
          if (!challanHistory || challanHistory.length === 0) {
            return false;
          }

          // Get all challans that have items (valid challans)
          const validChallans = challanHistory.filter(challan =>
            challan.items && Array.isArray(challan.items) && challan.items.length > 0
          );

          // If no valid challans, return false
          if (validChallans.length === 0) {
            return false;
          }

          // Check if ALL valid challans are embCompleted true
          return validChallans.every(challan =>
            challan.embCompleted === true && challan.embUpdatedAt
          );
        })();

        if (directStitchingValue !== 'yes' && !hasCompletedEmbChallans) {
          return null;
        }

        const lotNumber = row[headerIndices['Lot Number']] || 'N/A';
        const cuttingMatrix = cuttingMatrixData[lotNumber];
        const totalPcs = getTotalPcsFromMatrix(cuttingMatrix);

        const savedAtValue = row[savedAtIndex] || '';
        const cuttingDate = parseSavedAtDate(savedAtValue);

        const imageUrlRaw = imageUrlIndex !== -1 ? row[imageUrlIndex] || '' : '';
        const imageUrl = getDirectImageUrl(imageUrlRaw);

        const lastEmbDate = getLastEmbDate(challanHistory);

        const daysPending = calculateDaysDifference(directStitchingValue, cuttingDate, lastEmbDate);

        // Get color pending and priority info from the pending data
        const pendingInfo = pendingData[lotNumber] || {};
        const hasColorPending = pendingInfo.pendingColors?.length > 0 || false;
        const pendingColors = pendingInfo.pendingColors || [];
        const priority = pendingInfo.priority || '';

        // Check if it's a repeated lot
        const isRepeatedLot = priority &&
          (priority.toLowerCase().includes('repeated') ||
            priority.toLowerCase().includes('repeat'));

        // Get M/W/K value
        const mwkValue = (mwkIndex !== -1 && row[mwkIndex]) ? row[mwkIndex].toString().trim() : 'N/A';

        // Get Season value
        const seasonValue = (seasonIndex !== -1 && row[seasonIndex])
          ? row[seasonIndex].toString().trim()
          : (row[headerIndices['SEASON']] || row[headerIndices['Season']] || row[headerIndices['season']] || 'N/A');

        console.log(`Lot ${lotNumber}: priority="${priority}", isRepeatedLot=${isRepeatedLot}, mwk="${mwkValue}", season="${seasonValue}"`);

        return {
          id: rowIndex + 1,
          lotNumber: lotNumber,
          fabric: row[headerIndices['Fabric']] || 'N/A',
          garmentType: row[headerIndices['Garment Type']] || 'N/A',
          style: row[headerIndices['Style']] || 'N/A',
          brand: row[headerIndices['BRAND']] || 'N/A',
          season: seasonValue,
          directStitching: directStitchingValue,
          challanHistory: challanHistory,
          partyName: row[headerIndices['PARTY NAME']] || 'N/A',
          sizes: row[headerIndices['Sizes']] || 'N/A',
          shades: row[headerIndices['Shades']] || 'N/A',
          hasCompletedEmbChallans: hasCompletedEmbChallans,
          totalPcs: totalPcs,
          hasCuttingMatrix: !!cuttingMatrix,
          challanCount: challanHistory.length,
          completedChallanCount: challanHistory.filter(ch => ch.embCompleted === true && ch.embUpdatedAt).length,
          cuttingDate: cuttingDate,
          lastEmbDate: lastEmbDate,
          daysPending: daysPending,
          imageUrl: imageUrl,
          savedAtValue: savedAtValue,
          hasColorPending: hasColorPending,
          pendingColors: pendingColors,
          priority: priority,
          isRepeatedLot: isRepeatedLot,
          mwk: mwkValue, // Add M/W/K field
          rawRow: row
        };
      }).filter(item => item !== null);

      console.log('Formatted data sample:', formattedData.slice(0, 5).map(item => ({
        lotNumber: item.lotNumber,
        priority: item.priority,
        isRepeatedLot: item.isRepeatedLot,
        hasColorPending: item.hasColorPending,
        pendingColors: item.pendingColors,
        mwk: item.mwk
      })));

      console.log('Total repeated lots:', formattedData.filter(item => item.isRepeatedLot).length);

      setIssues(formattedData);
      setFilteredIssues(formattedData);
    } catch (err) {
      setError(err.message);
      console.error('Error fetching data:', err);
    } finally {
      setLoading(false);
    }
  };

  // Helper function for text normalization (SAME AS STITCHING COMPONENT)
  const normalizeText = (text) => {
    if (!text || text.trim() === '') return '';
    return text.trim().toLowerCase();
  };

  const getCompletedEmbChallans = (challans) => {
    if (!challans || !Array.isArray(challans)) return [];
    return challans.filter(challan =>
      challan.embCompleted === true && challan.embUpdatedAt
    );
  };

  const calculateTotalPieces = (issue) => {
    return issue.totalPcs || 0;
  };

  // Calculate total pieces for filtered issues
  const calculateTotalPiecesForFiltered = () => {
    return filteredIssues.reduce((sum, issue) => sum + calculateTotalPieces(issue), 0);
  };

  // Calculate total color pending and repeated lots
  const calculateColorPendingStats = () => {
    const colorPendingCount = filteredIssues.filter(issue => issue.hasColorPending).length;
    const repeatedLotCount = filteredIssues.filter(issue => issue.isRepeatedLot).length;
    return { colorPendingCount, repeatedLotCount };
  };

  const exportToCSV = () => {
    const { colorPendingCount, repeatedLotCount } = calculateColorPendingStats();
    const totalPieces = calculateTotalPiecesForFiltered();
    const totalLots = filteredIssues.length;

    const csvData = filteredIssues.map(issue => ({
      'Lot Number': issue.lotNumber,
      'Fabric': issue.fabric,
      'Garment Type': issue.garmentType,
      'Brand': issue.brand,
      'Style': issue.style,
      'Season': issue.season || 'N/A',
      'M/W/K': issue.mwk,
      'Total Pcs': calculateTotalPieces(issue),
      'Cutting Date': formatDateDisplay(issue.cuttingDate),
      'Emb/Printing Date': formatDateDisplay(issue.lastEmbDate),
      'Days Pending': issue.daysPending || 'N/A',
      'Direct Stitching': issue.directStitching === 'yes' ? 'Yes' : 'No',
      'Party Name': issue.partyName,
      'Color Status': issue.hasColorPending ? `Color Pending (${issue.pendingColors.length})` : 'OK',
      'Priority': issue.priority || 'N/A'
    }));

    // Add summary rows
    csvData.push({});
    csvData.push({
      'Lot Number': 'TOTAL LOTS',
      'Fabric': '',
      'Garment Type': '',
      'Brand': '',
      'Style': '',
      'Season': '',
      'M/W/K': '',
      'Total Pcs': '',
      'Cutting Date': '',
      'Emb/Printing Date': '',
      'Days Pending': '',
      'Direct Stitching': '',
      'Party Name': `${totalLots} Lots`,
      'Color Status': '',
      'Priority': ''
    });
    csvData.push({
      'Lot Number': 'TOTAL PIECES',
      'Fabric': '',
      'Garment Type': '',
      'Brand': '',
      'Style': '',
      'Season': '',
      'M/W/K': '',
      'Total Pcs': totalPieces,
      'Cutting Date': '',
      'Emb/Printing Date': '',
      'Days Pending': '',
      'Direct Stitching': '',
      'Party Name': '',
      'Color Status': '',
      'Priority': ''
    });
    csvData.push({
      'Lot Number': 'COLOR PENDING',
      'Fabric': '',
      'Garment Type': '',
      'Brand': '',
      'Style': '',
      'Season': '',
      'M/W/K': '',
      'Total Pcs': '',
      'Cutting Date': '',
      'Emb/Printing Date': '',
      'Days Pending': '',
      'Direct Stitching': '',
      'Party Name': '',
      'Color Status': `${colorPendingCount} Lots`,
      'Priority': ''
    });
    csvData.push({
      'Lot Number': 'REPEATED LOTS',
      'Fabric': '',
      'Garment Type': '',
      'Brand': '',
      'Style': '',
      'Season': '',
      'M/W/K': '',
      'Total Pcs': '',
      'Cutting Date': '',
      'Emb/Printing Date': '',
      'Days Pending': '',
      'Direct Stitching': '',
      'Party Name': '',
      'Color Status': '',
      'Priority': `${repeatedLotCount} Lots`
    });

    const headers = Object.keys(csvData[0]);
    const csv = [
      headers.join(','),
      ...csvData.map(row => headers.map(header =>
        `"${row[header]}"`
      ).join(','))
    ].join('\n');

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', `pending_issues_${new Date().toISOString().split('T')[0]}.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  // Enhanced PDF export with color pending and priority
  const exportToPDF = () => {
    const printWindow = window.open('', '_blank');
    const date = new Date().toLocaleDateString('en-IN', {
      day: '2-digit',
      month: 'short',
      year: 'numeric'
    });
    const time = new Date().toLocaleTimeString('en-IN', {
      hour: '2-digit',
      minute: '2-digit'
    });

    const totalPieces = calculateTotalPiecesForFiltered();
    const totalLots = filteredIssues.length;
    const { colorPendingCount, repeatedLotCount } = calculateColorPendingStats();

    printWindow.document.write(`
      <!DOCTYPE html>
      <html lang="en">
      <head>
        <meta charset="UTF-8">
        <title>Pending Issue Report - ${date}</title>
        <style>
          /* A4 Setup */
          @page {
            size: A4 landscape;
            margin: 8mm;
          }
          
          body { 
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif; 
            margin: 0; 
            padding: 0; 
            color: #000000;
            line-height: 1.4;
            background-color: #ffffff;
            font-size: 10px;
          }

          .report-wrapper {
            width: 100%;
            margin: auto;
          }

          /* Professional Header */
          .report-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 2px solid #000000;
            padding-bottom: 8px;
            margin-bottom: 15px;
          }

          .company-name h1 {
            margin: 0;
            font-size: 20px;
            font-weight: 800;
            color: #000000;
            letter-spacing: -0.025em;
          }

          .report-info {
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

          /* Stats Grid */
          .stats-container {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 12px;
            margin-bottom: 15px;
          }
          
          .stat-box {
            background: #ffffff;
            border: 1px solid #000000;
            border-radius: 6px;
            padding: 8px;
            text-align: center;
          }
          
          .stat-box .val {
            display: block;
            font-size: 15px;
            font-weight: 700;
            color: #000000;
            margin-bottom: 2px;
          }
          
          .stat-box .lbl {
            font-size: 8px;
            color: #000000;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.05em;
          }

          /* Table Design */
          table { 
            width: 100%; 
            border-collapse: collapse; 
            font-size: 9px;
            margin-bottom: 15px;
          }
          
          th, td { 
            padding: 6px;
            border: 1px solid #000000;
            color: #000000;
            word-wrap: break-word;
          }

          th { 
            background-color: #f1f5f9; 
            color: #000000;
            font-weight: 700;
            text-transform: uppercase;
            font-size: 8px;
            letter-spacing: 0.05em;
            text-align: left;
          }
          
          tr:nth-child(even) {
            background-color: #f8fafc;
          }

          /* Status Badges */
          .badge {
            display: inline-block;
            padding: 2px 6px;
            border: 1px solid #000000;
            border-radius: 4px;
            font-size: 8px;
            font-weight: 700;
            text-align: center;
            color: #000000;
          }
          
          .color-pending-tag {
            font-weight: 700;
            color: #000000;
            display: inline-block;
          }
          
          .color-ready-tag {
            color: #000000;
            font-weight: 500;
          }

          .footer {
            border-top: 1px solid #000000;
            padding-top: 8px;
            font-size: 8px;
            color: #000000;
            text-align: center;
          }

          @media print {
            .no-print { display: none; }
            body { padding: 0; }
          }
        </style>
      </head>
      <body>
        <div class="report-wrapper">
          <div class="report-header">
            <div class="company-name">
              <h1>Pending Issues Report After Emb/Print Done</h1>
              <div style="font-size: 10px; color: #000000; font-weight: 500; margin-top: 4px;">Supervision & Stitching Status</div>
            </div>
            <div class="report-info">
              <div class="meta-item"><span class="meta-label">Date:</span> ${date}</div>
              <div class="meta-item"><span class="meta-label">Time:</span> ${time}</div>
            </div>
          </div>

          <div class="stats-container">
            <div class="stat-box">
              <span class="val">${totalLots}</span>
              <span class="lbl">Total Lots</span>
            </div>
            <div class="stat-box">
              <span class="val">${totalPieces.toLocaleString()}</span>
              <span class="lbl">Total Pieces</span>
            </div>
            <div class="stat-box">
              <span class="val">${colorPendingCount}</span>
              <span class="lbl">Color Pending</span>
            </div>
            <div class="stat-box">
              <span class="val">${repeatedLotCount}</span>
              <span class="lbl">Repeated Lots</span>
            </div>
          </div>

          <table>
            <thead>
              <tr>
                <th style="width: 3%; text-align: center;">#</th>
                <th style="width: 9%;">Lot Number</th>
                <th style="width: 10%;">Fabric</th>
                <th style="width: 10%;">Garment Type</th>
                <th style="width: 10%;">Style</th>
                <th style="width: 8%;">Season</th>
                <th style="width: 7%;">Image</th>
                <th style="width: 8%;">Brand</th>
                <th style="width: 6%; text-align: center;">M/W/K</th>
                <th style="width: 7%; text-align: center;">Pieces</th>
                <th style="width: 10%;">Emb/Print Date</th>
                <th style="width: 6%; text-align: center;">Days</th>
                <th style="width: 12%;">Color/Status</th>
              </tr>
            </thead>
            <tbody>
              ${filteredIssues.map((issue, index) => {
      const statusClass = issue.daysPending > 14 ? 'critical' : issue.daysPending > 7 ? 'urgent' : 'normal';
      const rowStyle = issue.isRepeatedLot ? 'row-repeated' : (issue.hasColorPending ? 'row-pending' : '');

      return `
                  <tr>
                    <td style="text-align: center;">${index + 1}</td>
                    <td style="font-weight: bold;">
                      ${issue.isRepeatedLot ? '★ ' : ''}
                      ${issue.lotNumber}
                    </td>
                    <td>${issue.fabric}</td>
                    <td>${issue.garmentType}</td>
                    <td>${issue.style}</td>
                    <td>${issue.season || 'N/A'}</td>
                    <td style="text-align: center;">
                      ${issue.imageUrl
          ? `<img src="${issue.imageUrl}" referrerpolicy="no-referrer" style="width: 35px; height: 35px; object-fit: cover; border: 1px solid #000000; border-radius: 3px;" />`
          : '<span style="font-size: 8px; color: #666;">No Image</span>'}
                    </td>
                    <td>${issue.brand}</td>
                    <td style="text-align: center;">${issue.mwk}</td>
                    <td style="text-align: center; font-weight: bold;">${calculateTotalPieces(issue).toLocaleString()}</td>
                    <td>${formatDateDisplay(issue.lastEmbDate, '', issue)}</td>
                    <td style="text-align: center;">
                      <span class="badge">${issue.daysPending || 0}d</span>
                    </td>
                    <td>
                      ${issue.hasColorPending
          ? `<span class="color-pending-tag">⚠️ Pending: ${issue.pendingColors.slice(0, 2).join(', ')}${issue.pendingColors.length > 2 ? '...' : ''}</span>`
          : '<span class="color-ready-tag">✓ READY</span>'}
                    </td>
                  </tr>
                `;
    }).join('')}
            </tbody>
          </table>

          <div class="footer">
            <p>System Generated Report | Confidential Production Data</p>
            <p style="font-size: 8px; margin-top: 4px; color: #94a3b8;">★ Repeated Lot | ⚠️ Color Pending</p>
          </div>

          <div class="no-print" style="margin-top: 30px; text-align: center;">
            <button onclick="window.print()" style="background: #4f46e5; color: white; padding: 10px 25px; border: none; border-radius: 8px; cursor: pointer; font-weight: bold; font-size: 14px; box-shadow: 0 4px 12px rgba(79, 70, 229, 0.3);">
              PRINT REPORT
            </button>
          </div>
        </div>
      </body>
      </html>
    `);

    printWindow.document.close();
  };

  const IssueDetailModal = ({ issue, onClose }) => {
    if (!issue) return null;

    const completedChallans = getCompletedEmbChallans(issue.challanHistory);

    return (
      <div className="modal-overlay" onClick={onClose}>
        <div className="modal-content" onClick={e => e.stopPropagation()}>
          <div className="modal-header">
            <h2>Lot Details: {issue.lotNumber} {issue.isRepeatedLot && '★'}</h2>
            <button className="modal-close" onClick={onClose}>×</button>
          </div>

          <div className="modal-body">
            <div className="detail-grid">
              <div className="detail-section">
                <h3>Basic Information</h3>
                <div className="detail-row">
                  <span className="detail-label">Fabric:</span>
                  <span className="detail-value">{issue.fabric}</span>
                </div>
                <div className="detail-row">
                  <span className="detail-label">Garment Type:</span>
                  <span className="detail-value">{issue.garmentType}</span>
                </div>
                <div className="detail-row">
                  <span className="detail-label">Style:</span>
                  <span className="detail-value">{issue.style}</span>
                </div>
                <div className="detail-row">
                  <span className="detail-label">Brand:</span>
                  <span className="detail-value">{issue.brand}</span>
                </div>
                <div className="detail-row">
                  <span className="detail-label">Season:</span>
                  <span className="detail-value">{issue.season || 'N/A'}</span>
                </div>
                <div className="detail-row">
                  <span className="detail-label">M/W/K:</span>
                  <span className="detail-value">{issue.mwk}</span>
                </div>
              </div>

              <div className="detail-section">
                <h3>Production Timeline</h3>
                <div className="detail-row">
                  <span className="detail-label">Cutting Date:</span>
                  <span className="detail-value">{formatDateDisplay(issue.cuttingDate)}</span>
                </div>
                <div className="detail-row">
                  <span className="detail-label">Emb/Print Date:</span>
                  <span className="detail-value">{formatDateDisplay(issue.lastEmbDate)}</span>
                </div>
                <div className="detail-row">
                  <span className="detail-label">Days Pending:</span>
                  <span className="detail-value highlight">
                    {issue.daysPending || 'N/A'} days
                  </span>
                </div>
                <div className="detail-row">
                  <span className="detail-label">Total Pieces:</span>
                  <span className="detail-value highlight">{calculateTotalPieces(issue)}</span>
                </div>
              </div>

              <div className="detail-section">
                <h3>Status & Priority</h3>
                <div className="status-tags">
                  <span className={`status-tag ${issue.directStitching === 'yes' ? 'tag-direct' : 'tag-regular'}`}>
                    {issue.directStitching === 'yes' ? 'Direct Stitching' : 'Regular Production'}
                  </span>
                  <span className={`status-tag ${issue.hasCompletedEmbChallans ? 'tag-emb-completed' : 'tag-emb-pending'}`}>
                    {issue.hasCompletedEmbChallans ? 'EMB Completed' : 'EMB Pending'}
                  </span>
                  {issue.hasColorPending && (
                    <span className="status-tag tag-color-pending">
                      Color Pending ({issue.pendingColors.length})
                    </span>
                  )}
                  {issue.isRepeatedLot && (
                    <span className="status-tag tag-repeated">
                      ★ Repeated Lot
                    </span>
                  )}
                </div>
                <div className="detail-row">
                  <span className="detail-label">Party Name:</span>
                  <span className="detail-value">{issue.partyName}</span>
                </div>
                <div className="detail-row">
                  <span className="detail-label">Sizes:</span>
                  <span className="detail-value">{issue.sizes}</span>
                </div>
                <div className="detail-row">
                  <span className="detail-label">Shades:</span>
                  <span className="detail-value">{issue.shades}</span>
                </div>
                {issue.priority && (
                  <div className="detail-row">
                    <span className="detail-label">Priority:</span>
                    <span className="detail-value">{issue.priority}</span>
                  </div>
                )}
              </div>

              {issue.hasColorPending && issue.pendingColors.length > 0 && (
                <div className="detail-section full-width">
                  <h3>Pending Colors ({issue.pendingColors.length})</h3>
                  <div className="challan-grid">
                    {issue.pendingColors.map((color, index) => (
                      <div key={index} className="challan-card color-pending-card">
                        <div className="challan-header">
                          <span className="challan-number">Color {index + 1}</span>
                          <span className="pending-badge">Pending</span>
                        </div>
                        <div className="challan-body">
                          <div><strong>{color}</strong></div>
                          <div>Status: Awaiting cutting</div>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {completedChallans.length > 0 && (
                <div className="detail-section full-width">
                  <h3>Completed Challans ({completedChallans.length})</h3>
                  <div className="challan-grid">
                    {completedChallans.map(challan => (
                      <div key={challan.number} className="challan-card">
                        <div className="challan-header">
                          <span className="challan-number">Challan #{challan.number}</span>
                          {challan.completeLot === false && (
                            <span className="partial-badge">Partial</span>
                          )}
                        </div>
                        <div className="challan-body">
                          <div>Quantity: <strong>{challan.totalQty}</strong></div>
                          <div>Completed: {formatDateDisplay(challan.embUpdatedAt)}</div>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
    );
  };

  const TableRow = ({ issue, index }) => {
    const totalPcs = calculateTotalPieces(issue);
    const daysPending = issue.daysPending;

    const getDaysPendingClass = () => {
      if (!daysPending && daysPending !== 0) return 'days-normal';
      if (daysPending > 14) return 'days-critical';
      if (daysPending > 7) return 'days-urgent';
      return 'days-normal';
    };

    // Format color status text - shows actual pending colors
    const getColorStatusText = () => {
      if (!issue.hasColorPending || issue.pendingColors.length === 0) return 'No Colour Pending';

      // Show up to 3 colors, add "+X more" if there are more
      const displayColors = issue.pendingColors.slice(0, 3);
      let text = `Colour Pending: ${displayColors.join(', ')}`;

      if (issue.pendingColors.length > 3) {
        text += ` +${issue.pendingColors.length - 3} more`;
      }

      return text;
    };

    return (
      <tr
        className={`table-row ${issue.hasColorPending ? 'color-pending-row' : ''} ${issue.isRepeatedLot ? 'repeated-lot-row' : ''}`}
        onClick={() => setSelectedIssue(issue)}
      >
        <td className="text-center">
          <div className="row-number">
            {issue.isRepeatedLot ? '★ ' : ''}{index + 1}
          </div>
        </td>
        <td className="text-center" style={{ verticalAlign: 'middle' }}>
          {issue.imageUrl ? (
            <img
              src={issue.imageUrl}
              alt="Style Preview"
              className="table-image"
              referrerPolicy="no-referrer"
              onClick={(e) => {
                e.stopPropagation();
                setViewImageSrc(issue.imageUrl);
              }}
            />
          ) : (
            <span className="no-image-placeholder">No Image</span>
          )}
        </td>
        <td>
          <div className="lot-cell">
            <div className={`lot-number ${issue.hasColorPending ? 'color-pending' : ''} ${issue.isRepeatedLot ? 'repeated-lot' : ''}`}>
              {issue.lotNumber}
              {issue.hasColorPending && ' ⚠️'}
              {issue.isRepeatedLot && ' ★'}
            </div>
          </div>
        </td>
        <td>
          <div className="fabric-cell">{issue.fabric}</div>
        </td>
        <td>
          <div className="garment-type-cell">{issue.garmentType}</div>
        </td>
        <td>
          <div className="style-cell">{issue.style}</div>
        </td>
        <td>
          <div className="brand-cell">{issue.brand}</div>
        </td>
        <td>
          <div className="season-cell">{issue.season || 'N/A'}</div>
        </td>
        <td>
          <div className="mwk-cell">{issue.mwk}</div>
        </td>
        <td className="text-center">
          <div className={`piece-count ${totalPcs > 0 ? 'has-pieces' : 'no-pieces'}`}>
            {totalPcs > 0 ? totalPcs.toLocaleString() : '-'}
          </div>
        </td>
        <td>
          <div className="date-cell">{formatDateDisplay(issue.cuttingDate)}</div>
        </td>
        <td>
          <div className="date-cell">{formatDateDisplay(issue.lastEmbDate, '', issue)}</div>
        </td>
        <td className="text-center">
          <div className={`days-pending ${getDaysPendingClass()}`}>
            {daysPending || daysPending === 0 ? `${daysPending} days` : 'N/A'}
          </div>
        </td>
        <td className="text-center">
          <div className={`color-status ${issue.hasColorPending ? 'has-color-pending' : 'no-color-pending'}`}
            title={issue.hasColorPending ? issue.pendingColors.join(', ') : 'No pending colors'}>
            {getColorStatusText()}
          </div>
        </td>
      </tr>
    );
  };

  const CardView = () => (
    <div className="card-grid">
      {filteredIssues.map((issue, index) => {
        const totalPcs = calculateTotalPieces(issue);
        const daysPending = issue.daysPending;

        const getDaysPendingClass = () => {
          if (!daysPending && daysPending !== 0) return 'days-normal';
          if (daysPending > 14) return 'days-critical';
          if (daysPending > 7) return 'days-urgent';
          return 'days-normal';
        };

        return (
          <div
            key={issue.id}
            className={`issue-card ${issue.hasColorPending ? 'color-pending-card' : ''} ${issue.isRepeatedLot ? 'repeated-lot-card' : ''}`}
            onClick={() => setSelectedIssue(issue)}
          >
            <div className="card-header" style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
              {issue.imageUrl ? (
                <img
                  src={issue.imageUrl}
                  alt="Style Preview"
                  className="card-image"
                  referrerPolicy="no-referrer"
                  onClick={(e) => {
                    e.stopPropagation();
                    setViewImageSrc(issue.imageUrl);
                  }}
                />
              ) : (
                <div className="card-no-image-placeholder">No Image</div>
              )}
              <div className="card-header-text" style={{ flex: 1 }}>
                <div className="card-lot">
                  {issue.lotNumber}
                  {issue.hasColorPending && ' ⚠️'}
                  {issue.isRepeatedLot && ' ★'}
                </div>
                <div className={`card-days ${getDaysPendingClass()}`}>
                  {daysPending || daysPending === 0 ? `${daysPending} days pending` : 'No date'}
                </div>
              </div>
            </div>
            <div className="card-body">
              <div className="card-info">
                <div className="info-row">
                  <span className="info-label">Fabric:</span>
                  <span className="info-value">{issue.fabric}</span>
                </div>
                <div className="info-row">
                  <span className="info-label">Type:</span>
                  <span className="info-value">{issue.garmentType}</span>
                </div>
                <div className="info-row">
                  <span className="info-label">Style:</span>
                  <span className="info-value">{issue.style}</span>
                </div>
                <div className="info-row">
                  <span className="info-label">Brand:</span>
                  <span className="info-value">{issue.brand}</span>
                </div>
                <div className="info-row">
                  <span className="info-label">Season:</span>
                  <span className="info-value">{issue.season || 'N/A'}</span>
                </div>
                <div className="info-row">
                  <span className="info-label">M/W/K:</span>
                  <span className="info-value">{issue.mwk}</span>
                </div>
              </div>
              <div className="card-dates">
                <div className="date-row">
                  <span className="date-label">Cutting:</span>
                  <span className="date-value">{formatDateDisplay(issue.cuttingDate)}</span>
                </div>
                <div className="date-row">
                  <span className="date-label">Emb/Print:</span>
                  <span className="date-value">{formatDateDisplay(issue.lastEmbDate)}</span>
                </div>
              </div>
              <div className="card-stats">
                <div className="stat-item">
                  <div className="stat-value">{totalPcs > 0 ? totalPcs.toLocaleString() : '-'}</div>
                  <div className="stat-label">Pieces</div>
                </div>
                <div className="stat-item">
                  <div className="stat-value">{issue.hasColorPending ? issue.pendingColors.length : 0}</div>
                  <div className="stat-label">Pending Colors</div>
                </div>
              </div>
            </div>
            <div className="card-footer">
              <div className="status-tags">
                <span className={`status-tag ${issue.directStitching === 'yes' ? 'tag-direct' : 'tag-regular'}`}>
                  {issue.directStitching === 'yes' ? 'Direct' : 'Regular'}
                </span>
                <span className={`status-tag ${issue.hasCompletedEmbChallans ? 'tag-emb-done' : 'tag-emb-pending'}`}>
                  {issue.hasCompletedEmbChallans ? 'EMB Done' : 'No EMB'}
                </span>
                {issue.hasColorPending && (
                  <span className="status-tag tag-color-pending">
                    Color Pending
                  </span>
                )}
                {issue.isRepeatedLot && (
                  <span className="status-tag tag-repeated">
                    Repeated
                  </span>
                )}
              </div>
              <button className="card-action-btn" onClick={(e) => {
                e.stopPropagation();
                setSelectedIssue(issue);
              }}>
                Details →
              </button>
            </div>
          </div>
        );
      })}
    </div>
  );

  const TableView = () => {
    const { colorPendingCount, repeatedLotCount } = calculateColorPendingStats();

    return (
      <div className="table-wrapper">
        <div className="table-legend">
          <div className="legend-item">
            <div className="legend-color color-pending"></div>
            <span>Color Pending ({colorPendingCount})</span>
          </div>
          <div className="legend-item">
            <div className="legend-color repeated-lot"></div>
            <span>Repeated Lot ({repeatedLotCount})</span>
          </div>
          <div className="legend-item">
            <div className="legend-color urgent"></div>
            <span>Urgent (7-14 days)</span>
          </div>
          <div className="legend-item">
            <div className="legend-color critical"></div>
            <span>Critical (Less than 14 days)</span>
          </div>
        </div>

        <table className="main-table">
          <thead>
            <tr>
              <th className="text-center">#</th>
              <th>Image</th>
              <th>Lot Number</th>
              <th>Fabric</th>
              <th>Garment Type</th>
              <th>Style</th>
              <th>Brand</th>
              <th>Season</th>
              <th>M/W/K</th>
              <th className="text-center">Total Pcs</th>
              <th>Cutting Date</th>
              <th>Emb/Printing Date</th>
              <th className="text-center">Days Pending</th>
              <th className="text-center">Color Status</th>
            </tr>
          </thead>
          <tbody>
            {filteredIssues.length === 0 ? (
              <tr>
                <td colSpan="14" className="empty-state">
                  <div className="empty-message">
                    <span className="empty-icon">📭</span>
                    <h3>No pending issues found</h3>
                    <p>Try adjusting your search or filters</p>
                  </div>
                </td>
              </tr>
            ) : (
              filteredIssues.map((issue, index) => (
                <TableRow key={issue.id} issue={issue} index={index} />
              ))
            )}
          </tbody>
        </table>
      </div>
    );
  };

  // Calculate stats for filtered issues
  const stats = {
    total: filteredIssues.length,
    embCompleted: filteredIssues.filter(i => i.hasCompletedEmbChallans).length,
    directStitching: filteredIssues.filter(i => i.directStitching === 'yes').length,
    totalPieces: calculateTotalPiecesForFiltered(),
    urgent: filteredIssues.filter(i => i.daysPending > 7 && i.daysPending <= 14).length,
    critical: filteredIssues.filter(i => i.daysPending > 14).length,
    colorPending: filteredIssues.filter(i => i.hasColorPending).length,
    repeatedLots: filteredIssues.filter(i => i.isRepeatedLot).length
  };

  const filters = [
    { id: 'all', label: 'All Lots', count: issues.length },
    { id: 'emb', label: 'EMB Completed', count: issues.filter(i => i.hasCompletedEmbChallans).length },
    { id: 'direct', label: 'Direct Stitching', count: issues.filter(i => i.directStitching === 'yes').length },
    { id: 'both', label: 'Both', count: issues.filter(i => i.hasCompletedEmbChallans && i.directStitching === 'yes').length },
    { id: 'urgent', label: 'Urgent (>7 days)', count: issues.filter(i => i.daysPending > 7).length },
    { id: 'critical', label: 'Critical (>14 days)', count: issues.filter(i => i.daysPending > 14).length },
    { id: 'colorPending', label: 'Color Pending', count: issues.filter(i => i.hasColorPending).length },
    { id: 'repeated', label: 'Repeated Lots', count: issues.filter(i => i.isRepeatedLot).length }
  ];

  // Count active filters
  const activeFilterCount = [
    brandFilter && brandFilter.length > 0 ? 1 : 0,
    fabricFilter && fabricFilter.length > 0 ? 1 : 0,
    styleFilter && styleFilter.length > 0 ? 1 : 0,
    partyFilter && partyFilter.length > 0 ? 1 : 0,
    seasonFilter && seasonFilter.length > 0 ? 1 : 0,
    garmentTypeFilter && garmentTypeFilter.length > 0 ? 1 : 0,
    mwkFilter && mwkFilter.length > 0 ? 1 : 0,
    daysFilter ? 1 : 0,
    activeFilter !== 'all' ? 1 : 0,
    searchTerm ? 1 : 0
  ].filter(Boolean).length;

  // Dashboard-themed CSS
  const styles = `
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    .pending-issues-container {
      font-family: 'Plus Jakarta Sans', 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
      background-color: #f8fafc;
      background-image: radial-gradient(at 0% 0%, rgba(99, 102, 241, 0.08) 0px, transparent 50%),
                        radial-gradient(at 100% 0%, rgba(236, 72, 153, 0.06) 0px, transparent 50%),
                        radial-gradient(at 50% 100%, rgba(16, 185, 129, 0.06) 0px, transparent 50%);
      min-height: 100vh;
      padding: 0;
      color: #0f172a;
    }

    /* ===== DASHBOARD HERO HEADER ===== */
    .page-header-wrapper {
      width: 100%;
      padding: 24px 36px 0;
      box-sizing: border-box;
    }

    .dashboard-header {
      width: 100%;
      background: linear-gradient(135deg, #1e1b4b 0%, #312e81 40%, #4338ca 100%);
      border-radius: 24px;
      padding: 36px 40px;
      box-shadow: 0 20px 40px -10px rgba(49, 46, 129, 0.3);
      color: #ffffff;
      position: relative;
      overflow: visible;
      z-index: 10;
      box-sizing: border-box;
      margin-bottom: 0;
    }

    .header-bg-glow {
      position: absolute;
      inset: 0;
      border-radius: 24px;
      overflow: hidden;
      pointer-events: none;
    }

    .header-bg-glow::before {
      content: '';
      position: absolute;
      top: -50px;
      right: -50px;
      width: 250px;
      height: 250px;
      border-radius: 50%;
      background: rgba(255,255,255,0.08);
    }

    .header-bg-glow::after {
      content: '';
      position: absolute;
      bottom: -80px;
      right: 150px;
      width: 300px;
      height: 300px;
      border-radius: 50%;
      background: rgba(99,102,241,0.2);
    }

    .header-system-badge {
      display: inline-flex;
      align-items: center;
      gap: 10px;
      padding: 6px 16px;
      border-radius: 9999px;
      background: rgba(255,255,255,0.15);
      backdrop-filter: blur(12px);
      font-size: 0.85rem;
      font-weight: 600;
      color: #e0e7ff;
      border: 1px solid rgba(255,255,255,0.2);
      margin-bottom: 20px;
    }

    .header-top {
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 24px;
      flex-wrap: wrap;
      position: relative;
      z-index: 1;
    }

    .header-title h1 {
      font-size: clamp(1.6rem, 2.5vw, 2.2rem);
      font-weight: 800;
      color: #ffffff;
      margin: 0 0 8px 0;
      letter-spacing: -0.025em;
      line-height: 1.2;
    }

    .header-title p {
      color: #c7d2fe;
      font-size: 1rem;
      margin: 0;
      font-weight: 500;
    }

    .header-title .stats-line {
      font-size: 12px;
      color: rgba(255,255,255,0.7);
      margin-top: 10px;
      display: flex;
      align-items: center;
      gap: 12px;
      flex-wrap: wrap;
    }

    .header-title .stats-line span {
      padding: 4px 10px;
      border-radius: 8px;
      font-size: 11px;
      font-weight: 600;
      background: rgba(255,255,255,0.12);
      border: 1px solid rgba(255,255,255,0.2);
      color: #e0e7ff;
    }

    .header-title .color-pending-count {
      background: rgba(251,191,36,0.2);
      color: #fde68a;
      border: 1px solid rgba(251,191,36,0.3);
    }

    .header-title .repeated-count {
      background: rgba(167,139,250,0.2);
      color: #ddd6fe;
      border: 1px solid rgba(167,139,250,0.3);
    }

    /* Header KPI Glass Cards */
    .header-kpi-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(110px, 1fr));
      gap: 12px;
      min-width: 300px;
    }

    .header-kpi-card {
      background: rgba(255,255,255,0.12);
      border: 1px solid rgba(255,255,255,0.2);
      backdrop-filter: blur(12px);
      border-radius: 16px;
      padding: 16px 14px;
      text-align: center;
    }

    .header-kpi-value {
      font-size: 1.6rem;
      font-weight: 800;
      color: #ffffff;
      display: block;
      line-height: 1;
      margin-bottom: 4px;
    }

    .header-kpi-label {
      font-size: 0.72rem;
      font-weight: 600;
      color: #c7d2fe;
      text-transform: uppercase;
      letter-spacing: 0.05em;
    }

    .header-actions {
      display: flex;
      gap: 10px;
      align-items: center;
      flex-wrap: wrap;
      margin-top: 20px;
      padding-top: 20px;
      border-top: 1px solid rgba(255,255,255,0.15);
      position: relative;
      z-index: 1;
    }

    /* ===== CONTROLS / FILTERS PANEL (Dashboard Style) ===== */
    .page-body {
      padding: 24px 36px;
      box-sizing: border-box;
    }

    .dropdown-filters-section {
      background: #ffffff;
      border-radius: 20px;
      padding: 24px;
      margin-bottom: 24px;
      box-shadow: 0 10px 30px rgba(0,0,0,0.03);
      border: 1px solid #e2e8f0;
    }

    .filters-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 20px;
      padding-bottom: 16px;
      border-bottom: 1px solid #f1f5f9;
    }

    .filters-title {
      font-size: 1rem;
      font-weight: 700;
      color: #1e293b;
    }

    .active-filters-count {
      background: linear-gradient(135deg, #4f46e5 0%, #3730a3 100%);
      color: white;
      padding: 4px 12px;
      border-radius: 20px;
      font-size: 12px;
      font-weight: 600;
      box-shadow: 0 4px 10px rgba(79,70,229,0.2);
    }

    .dropdown-filters-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
      gap: 16px;
      margin-bottom: 20px;
    }

    .filter-group {
      display: flex;
      flex-direction: column;
      gap: 8px;
    }

    .filter-label {
      font-size: 11px;
      font-weight: 700;
      color: #64748b;
      text-transform: uppercase;
      letter-spacing: 0.06em;
    }

    .filter-select {
      padding: 10px 12px;
      border: 1.5px solid #e2e8f0;
      border-radius: 12px;
      font-size: 14px;
      font-weight: 500;
      color: #334155;
      background: #f8fafc;
      cursor: pointer;
      transition: all 0.2s;
      outline: none;
    }

    .filter-select:focus {
      border-color: #4f46e5;
      background: #ffffff;
      box-shadow: 0 0 0 3px rgba(79,70,229,0.1);
    }

    .filter-select:hover {
      border-color: #a5b4fc;
      background: #ffffff;
    }

    .clear-filters-btn {
      padding: 10px 20px;
      background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%);
      color: white;
      border: none;
      border-radius: 12px;
      font-size: 13px;
      font-weight: 700;
      cursor: pointer;
      transition: all 0.2s;
      display: flex;
      align-items: center;
      gap: 8px;
      width: 100%;
      justify-content: center;
      box-shadow: 0 4px 12px rgba(239,68,68,0.25);
    }

    .clear-filters-btn:hover {
      transform: translateY(-1px);
      box-shadow: 0 6px 16px rgba(239,68,68,0.3);
    }

    .clear-filters-btn:disabled {
      background: #e2e8f0;
      color: #94a3b8;
      border: 1px solid #cbd5e1;
      cursor: not-allowed;
      box-shadow: none;
      transform: none;
    }

    /* Search */
    .search-input-wrapper {
      position: relative;
      margin-bottom: 16px;
    }

    .search-input {
      width: 100%;
      box-sizing: border-box;
      padding: 14px 20px 14px 48px;
      border: 1.5px solid #cbd5e1;
      border-radius: 14px;
      font-size: 0.95rem;
      font-weight: 500;
      transition: all 0.25s;
      background: #f8fafc;
      color: #0f172a;
      outline: none;
    }

    .search-input:focus {
      border-color: #4f46e5;
      box-shadow: 0 0 0 3px rgba(79,70,229,0.1);
      background: white;
    }

    .search-icon {
      position: absolute;
      left: 16px;
      top: 50%;
      transform: translateY(-50%);
      color: #94a3b8;
      font-size: 1.1rem;
    }

    /* ===== FILTER TABS (Dashboard category style) ===== */
    .filter-tabs {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
    }

    .filter-tab {
      padding: 10px 18px;
      border: 1.5px solid #e2e8f0;
      border-radius: 12px;
      background: #f8fafc;
      color: #475569;
      font-size: 0.875rem;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.25s ease;
      display: flex;
      align-items: center;
      gap: 6px;
    }

    .filter-tab:hover {
      border-color: #a5b4fc;
      background: #eef2ff;
      color: #4f46e5;
    }

    .filter-tab.active {
      background: linear-gradient(135deg, #4f46e5 0%, #3730a3 100%);
      color: white;
      border-color: #4338ca;
      box-shadow: 0 6px 16px rgba(79, 70, 229, 0.25);
    }

    .filter-count {
      background: rgba(255,255,255,0.2);
      padding: 2px 8px;
      border-radius: 10px;
      font-size: 11px;
      font-weight: 700;
    }

    /* View Toggle */
    .view-toggle {
      display: flex;
      background: #f1f5f9;
      border-radius: 12px;
      padding: 4px;
      margin-left: auto;
      border: 1px solid #e2e8f0;
    }

    .view-btn {
      padding: 8px 18px;
      border: none;
      background: none;
      border-radius: 999;
      font-size: 0.875rem;
      font-weight: 700;
      color: #64748b;
      cursor: pointer;
      transition: all 0.2s;
      display: flex;
      align-items: center;
      gap: 6px;
    }

    .view-btn.active {
      background: white;
      color: #4f46e5;
      box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    }

    /* ===== STATS GRID (Dashboard style) ===== */
    .stats-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
      gap: 16px;
      margin-bottom: 24px;
    }

    .stat-card {
      background: white;
      border-radius: 18px;
      padding: 22px 20px;
      box-shadow: 0 10px 25px -5px rgba(0,0,0,0.04);
      border: 1.5px solid #e2e8f0;
      transition: all 0.25s ease;
      position: relative;
      overflow: hidden;
    }

    .stat-card::before {
      content: '';
      position: absolute;
      top: 0;
      left: 0;
      right: 0;
      height: 3px;
      background: linear-gradient(135deg, #4f46e5 0%, #7c3aed 100%);
      opacity: 0;
      transition: opacity 0.25s;
    }

    .stat-card:hover {
      transform: translateY(-3px);
      box-shadow: 0 16px 32px -8px rgba(79,70,229,0.15);
      border-color: #c7d2fe;
    }

    .stat-card:hover::before {
      opacity: 1;
    }

    .stat-card.critical {
      border-left: 4px solid #ef4444;
    }

    .stat-card.urgent {
      border-left: 4px solid #f59e0b;
    }

    .stat-card.normal {
      border-left: 4px solid #10b981;
    }

    .stat-card.pieces {
      border-left: 4px solid #3b82f6;
    }

    .stat-card.color-pending {
      border-left: 4px solid #ff9800;
    }

    .stat-card.repeated {
      border-left: 4px solid #8b5cf6;
    }

    .stat-value {
      font-size: 2rem;
      font-weight: 800;
      color: #1e293b;
      margin-bottom: 4px;
      line-height: 1;
    }

    .stat-label {
      font-size: 0.75rem;
      color: #64748b;
      font-weight: 700;
      display: flex;
      align-items: center;
      gap: 6px;
      text-transform: uppercase;
      letter-spacing: 0.05em;
      margin-top: 6px;
    }

    /* Table Legend */
    .table-legend {
      display: flex;
      justify-content: center;
      gap: 20px;
      margin-bottom: 16px;
      padding: 12px;
      background: #f8fafc;
      border-radius: 8px;
      border: 1px solid #e2e8f0;
    }

    .legend-item {
      display: flex;
      align-items: center;
      gap: 8px;
      font-size: 12px;
      color: #475569;
      font-weight: 500;
    }

    .legend-color {
      width: 12px;
      height: 12px;
      border-radius: 3px;
    }

    .legend-color.color-pending {
      background-color: #ffa500;
    }

    .legend-color.repeated-lot {
      background-color: #8b5cf6;
    }

    .legend-color.urgent {
      background-color: #f59e0b;
    }

    .legend-color.critical {
      background-color: #ef4444;
    }

    /* Table Styles */
    .table-wrapper {
      background: white;
      border-radius: 20px;
      overflow: hidden;
      box-shadow: 0 10px 30px rgba(0,0,0,0.04);
      border: 1.5px solid #e2e8f0;
    }

    .main-table {
      width: 100%;
      border-collapse: collapse;
    }

    .main-table th {
      background: #0f172a;
      padding: 16px 20px;
      text-align: left;
      font-size: 12px;
      font-weight: 600;
      color: #ffffff;
      text-transform: uppercase;
      letter-spacing: 0.05em;
      border: 1px solid #334155;
    }

    .main-table th.text-center {
      text-align: center;
    }

    .table-row {
      border-bottom: 1px solid #cbd5e1;
      cursor: pointer;
      transition: background-color 0.2s;
    }

    .table-row:hover {
      background-color: #f8fafc;
    }

    .table-row:last-child {
      border-bottom: none;
    }

    .color-pending-row {
      background-color: #fff3e0 !important;
    }

    .color-pending-row:hover {
      background-color: #ffecb3 !important;
    }

    .repeated-lot-row {
      background-color: #f5f3ff !important;
    }

    .repeated-lot-row:hover {
      background-color: #ede9fe !important;
    }

    .table-row td {
      padding: 16px 20px;
      vertical-align: middle;
      border: 1px solid #cbd5e1;
    }

    .table-image {
      width: 40px;
      height: 40px;
      object-fit: cover;
      border-radius: 6px;
      border: 1px solid #e2e8f0;
      cursor: pointer;
      transition: transform 0.2s, box-shadow 0.2s;
    }
    
    .table-image:hover {
      transform: scale(1.15);
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    
    .no-image-placeholder {
      font-size: 9px;
      color: #94a3b8;
      text-transform: uppercase;
      letter-spacing: 0.05em;
    }

    .card-image {
      width: 50px;
      height: 50px;
      object-fit: cover;
      border-radius: 8px;
      border: 1px solid #e2e8f0;
      cursor: pointer;
    }

    .card-no-image-placeholder {
      width: 50px;
      height: 50px;
      display: flex;
      align-items: center;
      justify-content: center;
      background: #f1f5f9;
      border-radius: 8px;
      font-size: 8px;
      color: #94a3b8;
      border: 1px solid #e2e8f0;
    }
    
    /* Image Modal Lightbox */
    .image-modal-backdrop {
      position: fixed;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background: rgba(0, 0, 0, 0.75);
      z-index: 9999;
      display: flex;
      align-items: center;
      justify-content: center;
      backdrop-filter: blur(4px);
    }
    
    .image-modal-content {
      position: relative;
      max-width: 90%;
      max-height: 90%;
      background: white;
      border-radius: 12px;
      padding: 16px;
      box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);
    }
    
    .image-modal-img {
      max-width: 100%;
      max-height: 80vh;
      border-radius: 6px;
      object-fit: contain;
    }
    
    .image-modal-close {
      position: absolute;
      top: -12px;
      right: -12px;
      background: #0f172a;
      color: white;
      border: none;
      width: 28px;
      height: 28px;
      border-radius: 50%;
      font-size: 18px;
      display: flex;
      align-items: center;
      justify-content: center;
      cursor: pointer;
      box-shadow: 0 4px 6px rgba(0, 0, 0, 0.15);
      transition: background-color 0.2s;
    }
    
    .image-modal-close:hover {
      background: #1e293b;
    }

    /* Table Cell Styles */
    .lot-cell, .fabric-cell, .garment-type-cell, .brand-cell, .style-cell, .season-cell, .date-cell, .mwk-cell {
      font-size: 14px;
      color: #334155;
      font-weight: 500;
    }

    .lot-number {
      font-weight: 700;
      color: #1e40af;
    }

    .lot-number.color-pending {
      color: #b45309;
      background: #fffbeb;
      padding: 4px 8px;
      border-radius: 6px;
      border: 1px solid #fde68a;
      font-weight: 800;
    }

    .lot-number.repeated-lot {
      color: #6d28d9 !important;
      background: #f5f3ff;
      border: 1px solid #ddd6fe;
      padding: 4px 8px;
      border-radius: 6px;
    }

    .row-number {
      font-size: 14px;
      color: #64748b;
      font-weight: 600;
    }

    .piece-count {
      font-size: 16px;
      font-weight: 700;
    }

    .has-pieces {
      color: #dc2626;
    }

    .no-pieces {
      color: #94a3b8;
    }

    /* Days Pending Styles */
    .days-pending {
      padding: 6px 12px;
      border-radius: 12px;
      font-size: 13px;
      font-weight: 700;
      display: inline-block;
    }

    .days-critical {
      background: #fee2e2;
      color: #dc2626;
      border: 1px solid #fecaca;
    }

    .days-urgent {
      background: #fffbeb;
      color: #b45309;
      border: 1px solid #fde68a;
    }

    .days-normal {
      background: #ecfdf5;
      color: #047857;
      border: 1px solid #a7f3d0;
    }

    /* Color Status */
    .color-status {
      padding: 6px 12px;
      border-radius: 6px;
      font-size: 12px;
      font-weight: 700;
      display: inline-block;
    }

    .color-status.has-color-pending {
      background: #fffbeb;
      color: #b45309;
      border: 1px solid #fde68a;
    }

    .color-status.no-color-pending {
      background: #ecfdf5;
      color: #047857;
      border: 1px solid #a7f3d0;
    }

    /* Card View */
    .card-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
      gap: 16px;
      margin-bottom: 24px;
    }

    .issue-card {
      background: white;
      border-radius: 14px;
      padding: 20px;
      box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.02);
      border: 1px solid #e2e8f0;
      cursor: pointer;
      transition: all 0.25s ease;
    }

    .issue-card:hover {
      transform: translateY(-2px);
      box-shadow: 0 12px 20px -3px rgba(0, 0, 0, 0.05);
    }

    .color-pending-card {
      background: #fffbeb;
      border-color: #fde68a;
    }

    .repeated-lot-card {
      background: #f5f3ff;
      border-color: #ddd6fe;
    }

    .card-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 16px;
      padding-bottom: 12px;
      border-bottom: 1px solid #f1f5f9;
    }

    .card-lot {
      font-weight: 700;
      color: #1e293b;
      font-size: 16px;
    }

    .card-days {
      font-size: 12px;
      font-weight: 700;
      padding: 4px 10px;
      border-radius: 12px;
    }

    .card-body {
      margin-bottom: 16px;
    }

    .card-info {
      display: flex;
      flex-direction: column;
      gap: 8px;
    }

    .info-row {
      display: flex;
      justify-content: space-between;
      font-size: 13px;
    }

    .info-label {
      color: #64748b;
      font-weight: 500;
    }

    .info-value {
      color: #1e293b;
      font-weight: 600;
    }

    .card-pieces-row {
      border-top: 1px solid #f1f5f9;
      margin-top: 12px;
      padding-top: 12px;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }

    .card-pieces-label {
      font-size: 13px;
      color: #64748b;
      font-weight: 600;
    }

    .card-pieces-value {
      font-size: 18px;
      font-weight: 700;
      color: #dc2626;
    }

    .card-dates {
      background: #f8fafc;
      border-radius: 8px;
      padding: 12px;
      margin-bottom: 16px;
      border: 1px solid #e2e8f0;
    }

    .date-row {
      display: flex;
      justify-content: space-between;
      font-size: 13px;
      margin-bottom: 6px;
    }

    .date-row:last-child {
      margin-bottom: 0;
    }

    .date-label {
      color: #64748b;
      font-weight: 500;
    }

    .date-value {
      color: #1e293b;
      font-weight: 600;
    }

    .card-stats {
      display: flex;
      gap: 16px;
    }

    .stat-item {
      flex: 1;
      text-align: center;
      padding: 12px;
      background: #f8fafc;
      border-radius: 8px;
      border: 1px solid #e2e8f0;
    }

    .stat-item .stat-value {
      font-size: 20px;
      font-weight: 700;
      margin-bottom: 4px;
    }

    .stat-item .stat-label {
      font-size: 11px;
      color: #64748b;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.025em;
    }

    .card-footer {
      display: flex;
      justify-content: space-between;
      align-items: center;
      border-top: 1px solid #f1f5f9;
      padding-top: 12px;
      margin-top: 12px;
    }

    .status-tags {
      display: flex;
      gap: 6px;
      flex-wrap: wrap;
    }

    .status-tag {
      padding: 4px 10px;
      border-radius: 12px;
      font-size: 11px;
      font-weight: 700;
    }

    .tag-direct {
      background: #eff6ff;
      color: #1e40af;
      border: 1px solid #bfdbfe;
    }

    .tag-regular {
      background: #f3f4f6;
      color: #4b5563;
      border: 1px solid #e5e7eb;
    }

    .tag-emb-done {
      background: #ecfdf5;
      color: #047857;
      border: 1px solid #a7f3d0;
    }

    .tag-emb-pending {
      background: #fee2e2;
      color: #dc2626;
      border: 1px solid #fecaca;
    }

    .tag-color-pending {
      background: #fffbeb;
      color: #b45309;
      border: 1px solid #fde68a;
    }

    .tag-repeated {
      background: #f5f3ff;
      color: #6d28d9;
      border: 1px solid #ddd6fe;
    }

    .card-action-btn {
      padding: 6px 12px;
      border: 1px solid #cbd5e1;
      border-radius: 8px;
      background: white;
      color: #475569;
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.2s;
    }

    .card-action-btn:hover {
      background: #f8fafc;
      color: #4f46e5;
      border-color: #cbd5e1;
    }

    /* Empty State */
    .empty-state {
      padding: 60px 20px;
      text-align: center;
      background: white;
      border-radius: 14px;
      border: 1px solid #e2e8f0;
    }

    .empty-message {
      max-width: 400px;
      margin: 0 auto;
    }

    .empty-icon {
      font-size: 48px;
      color: #94a3b8;
      margin-bottom: 16px;
      display: block;
    }

    .empty-message h3 {
      font-size: 18px;
      font-weight: 700;
      color: #1e293b;
      margin: 0 0 8px 0;
    }

    .empty-message p {
      color: #64748b;
      font-size: 14px;
      margin: 0;
    }

    /* Modal - Updated for color pending */
    .modal-overlay {
      position: fixed;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background: rgba(15, 23, 42, 0.6);
      display: flex;
      align-items: center;
      justify-content: center;
      z-index: 1000;
      padding: 20px;
      backdrop-filter: blur(4px);
    }

    .modal-content {
      background: white;
      border-radius: 16px;
      max-width: 800px;
      width: 100%;
      max-height: 90vh;
      overflow-y: auto;
      box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.25);
      border: 1px solid #e2e8f0;
    }

    .modal-header {
      padding: 24px 32px;
      border-bottom: 1px solid #e2e8f0;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }

    .modal-header h2 {
      margin: 0;
      font-size: 20px;
      font-weight: 700;
      color: #1e293b;
    }

    .modal-close {
      background: none;
      border: none;
      font-size: 24px;
      color: #64748b;
      cursor: pointer;
      padding: 0;
      width: 32px;
      height: 32px;
      display: flex;
      align-items: center;
      justify-content: center;
      border-radius: 8px;
      transition: all 0.2s;
    }

    .modal-close:hover {
      background: #f1f5f9;
      color: #1e293b;
    }

    .modal-body {
      padding: 32px;
    }

    .detail-grid {
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: 24px;
    }

    .detail-section.full-width {
      grid-column: 1 / -1;
    }

    .detail-section h3 {
      font-size: 16px;
      font-weight: 700;
      color: #1e293b;
      margin: 0 0 16px 0;
      border-bottom: 2px solid #f1f5f9;
      padding-bottom: 8px;
    }

    .detail-row {
      display: flex;
      justify-content: space-between;
      padding: 8px 0;
      border-bottom: 1px solid #f8fafc;
      font-size: 14px;
    }

    .detail-row:last-child {
      border-bottom: none;
    }

    .detail-label {
      color: #64748b;
      font-weight: 500;
    }

    .detail-value {
      color: #1e293b;
      font-weight: 600;
    }

    .detail-value.highlight {
      color: #dc2626;
      font-weight: 700;
    }

    /* Challan details */
    .challan-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(220px, 1fr));
      gap: 16px;
    }

    .challan-card {
      background: #f8fafc;
      border: 1px solid #e2e8f0;
      border-radius: 8px;
      padding: 16px;
    }

    .challan-card.color-pending-card {
      background: #fffbeb;
      border-color: #fde68a;
    }

    .challan-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 12px;
      border-bottom: 1px solid #e2e8f0;
      padding-bottom: 8px;
    }

    .challan-number {
      font-weight: 700;
      font-size: 13px;
      color: #475569;
    }

    .partial-badge {
      background: #fffbeb;
      color: #d97706;
      padding: 2px 6px;
      border-radius: 4px;
      font-size: 10px;
      font-weight: 700;
      border: 1px solid #fcd34d;
    }

    .pending-badge {
      background: #fee2e2;
      color: #dc2626;
      padding: 2px 6px;
      border-radius: 4px;
      font-size: 10px;
      font-weight: 700;
      border: 1px solid #fecaca;
    }

    .challan-body {
      font-size: 13px;
      color: #475569;
      display: flex;
      flex-direction: column;
      gap: 4px;
    }

    /* Buttons */
    .btn {
      padding: 10px 20px;
      border-radius: 8px;
      font-size: 14px;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.2s;
      border: 1px solid transparent;
      display: inline-flex;
      align-items: center;
      gap: 8px;
    }

    .btn-primary {
      background: linear-gradient(135deg, #4f46e5 0%, #3730a3 100%);
      color: white;
      box-shadow: 0 4px 12px rgba(79, 70, 229, 0.3);
    }

    .btn-primary:hover {
      transform: translateY(-2px);
      box-shadow: 0 6px 20px rgba(79, 70, 229, 0.4);
    }

    .btn-secondary {
      background: white;
      color: #475569;
      border-color: #cbd5e1;
      box-shadow: 0 2px 4px rgba(0, 0, 0, 0.02);
    }

    .btn-secondary:hover {
      background: #f8fafc;
      transform: translateY(-2px);
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
    }

    .btn-success {
      background: linear-gradient(135deg, #10b981 0%, #059669 100%);
      color: white;
      box-shadow: 0 4px 12px rgba(16, 185, 129, 0.3);
    }

    .btn-success:hover {
      transform: translateY(-2px);
      box-shadow: 0 6px 20px rgba(16, 185, 129, 0.4);
    }

    /* Export Menu */
    .export-menu {
      position: absolute;
      top: calc(100% + 8px);
      right: 0;
      background: #ffffff;
      border-radius: 14px;
      box-shadow: 0 20px 40px rgba(0, 0, 0, 0.25), 0 0 0 1px rgba(0, 0, 0, 0.08);
      padding: 8px;
      min-width: 220px;
      z-index: 99999;
    }

    .export-option {
      padding: 10px 16px;
      border: none;
      background: none;
      width: 100%;
      text-align: left;
      color: #64748b;
      font-size: 14px;
      cursor: pointer;
      border-radius: 6px;
      display: flex;
      align-items: center;
      gap: 8px;
      transition: all 0.2s;
    }

    .export-option:hover {
      background: #f8fafc;
      color: #1e293b;
    }

    /* ===== LOADING / ERROR STATES (Dashboard style) ===== */
    @keyframes dashSpin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }

    .loading-container {
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      min-height: 380px;
      background: #ffffff;
      border-radius: 24px;
      padding: 48px 32px;
      box-shadow: 0 10px 30px rgba(0,0,0,0.03);
      border: 1px solid #e2e8f0;
      text-align: center;
      max-width: 500px;
      margin: 40px auto;
    }

    .spinner {
      width: 64px;
      height: 64px;
      border: 4px solid #e2e8f0;
      border-top: 4px solid #4f46e5;
      border-right: 4px solid #a855f7;
      border-radius: 50%;
      animation: dashSpin 0.85s cubic-bezier(0.5, 0, 0.5, 1) infinite;
      box-shadow: 0 0 20px rgba(79,70,229,0.2);
      margin-bottom: 24px;
    }

    .loading-container h3 {
      font-size: 1.4rem;
      font-weight: 800;
      color: #0f172a;
      margin: 0 0 8px 0;
    }

    .loading-container p {
      color: #64748b;
      font-size: 0.9rem;
      margin: 0;
    }

    .error-container {
      text-align: center;
      padding: 60px 20px;
      max-width: 500px;
      margin: 40px auto;
      background: white;
      border-radius: 24px;
      box-shadow: 0 10px 30px rgba(0,0,0,0.03);
      border: 1px solid #e2e8f0;
    }

    .error-icon {
      font-size: 48px;
      color: #ef4444;
      margin-bottom: 20px;
    }

    @keyframes spin {
      to { transform: rotate(360deg); }
    }

    /* Footer */
    .footer-info {
      background: white;
      border-radius: 18px;
      padding: 20px 28px;
      margin-top: 24px;
      box-shadow: 0 4px 16px rgba(0,0,0,0.03);
      border: 1.5px solid #e2e8f0;
      text-align: center;
      font-size: 13px;
      color: #64748b;
    }

    /* Responsive */
    @media (max-width: 1024px) {
      .main-table {
        font-size: 13px;
      }
      
      .main-table th,
      .main-table td {
        padding: 12px 16px;
      }
      
      .dropdown-filters-grid {
        grid-template-columns: repeat(2, 1fr);
      }
    }

    @media (max-width: 768px) {
      .pending-issues-container {
        padding: 12px;
      }

      .dashboard-header {
        padding: 16px;
      }

      .header-top {
        flex-direction: column;
        gap: 16px;
      }

      .header-actions {
        width: 100%;
        flex-direction: column;
      }

      .stats-grid {
        grid-template-columns: repeat(2, 1fr);
      }

      .filter-tabs {
        overflow-x: auto;
        padding-bottom: 8px;
      }

      .table-wrapper {
        overflow-x: auto;
      }

      .main-table {
        min-width: 1100px;
      }

      .detail-grid {
        grid-template-columns: 1fr;
      }

      .card-grid {
        grid-template-columns: 1fr;
      }

      .modal-content {
        margin: 20px;
        max-height: 80vh;
      }
      
      .dropdown-filters-grid {
        grid-template-columns: 1fr;
      }
      
      .filters-header {
        flex-direction: column;
        align-items: flex-start;
        gap: 12px;
      }
    }

    @media (max-width: 640px) {
      .stats-grid {
        grid-template-columns: 1fr;
      }

      .filter-tabs {
        justify-content: flex-start;
      }

      .search-input {
        width: 100%;
      }
    }
  `;

  if (loading) {
    return (
      <>
        <style>{styles}</style>
        <div className="pending-issues-container">
          <div className="loading-container">
            <div className="spinner"></div>
            <div style={{ fontSize: '11px', fontWeight: '800', letterSpacing: '1px', textTransform: 'uppercase', color: '#4f46e5', background: '#eef2ff', padding: '6px 16px', borderRadius: '20px', border: '1px solid #c7d2fe', marginBottom: '16px' }}>
              ⚡ MH FACTORY SUITE PRO
            </div>
            <h3>Fetching Live Production Data</h3>
            <p>Synchronizing EMB/Print Issue data from Google Sheets...</p>
          </div>
        </div>
      </>
    );
  }

  if (error) {
    return (
      <>
        <style>{styles}</style>
        <div className="pending-issues-container">
          <div className="error-container">
            <div className="error-icon">⚠️</div>
            <h3 style={{ color: '#1e293b', margin: '0 0 8px 0' }}>Error Loading Data</h3>
            <p style={{ color: '#64748b', marginBottom: '20px' }}>{error}</p>
            <div style={{ display: 'flex', gap: 12, justifyContent: 'center' }}>
              <button
                onClick={fetchData}
                style={{ padding: '10px 24px', background: 'linear-gradient(135deg, #4f46e5 0%, #3730a3 100%)', color: 'white', border: 'none', borderRadius: '12px', fontWeight: '700', cursor: 'pointer', fontSize: '14px', boxShadow: '0 4px 12px rgba(79,70,229,0.25)' }}
              >
                🔄 Retry
              </button>
              {onBack && (
                <button
                  onClick={onBack}
                  style={{ padding: '10px 24px', background: '#f1f5f9', color: '#475569', border: '1.5px solid #e2e8f0', borderRadius: '12px', fontWeight: '700', cursor: 'pointer', fontSize: '14px' }}
                >
                  ← Back
                </button>
              )}
            </div>
          </div>
        </div>
      </>
    );
  }

  return (
    <>
      <style>{styles}</style>
      <div className="pending-issues-container">

        {/* ====== DASHBOARD-STYLE HERO HEADER ====== */}
        <div className="page-header-wrapper">
          <div className="dashboard-header">
            <div className="header-bg-glow" />
            {/* Top badge row */}
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '12px', marginBottom: '20px', position: 'relative', zIndex: 1 }}>
              <div className="header-system-badge">
                <span style={{ display: 'inline-block', width: '8px', height: '8px', borderRadius: '50%', background: '#34d399', boxShadow: '0 0 10px #34d399' }} />
                <span>Garment Production Suite</span>
                <span style={{ color: 'rgba(255,255,255,0.4)' }}>|</span>
                <span style={{ color: '#c7d2fe' }}>After EMB/PRINT DONE</span>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '16px', fontSize: '0.85rem', color: '#c7d2fe' }}>
                <span>📅 {new Date().toLocaleDateString('en-IN', { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' })}</span>
                <span style={{ color: '#34d399', fontWeight: 600 }}>🟢 System Active</span>
              </div>
            </div>

            {/* Main Title & KPI Row */}
            <div className="header-top">
              <div className="header-title">
                <h1>Pending Issues to Stitching</h1>
                <p>After Emb/Printing Done • Supervisor Assignment Pending</p>
                <div className="stats-line">
                  <span>Total Pieces: {stats.totalPieces.toLocaleString()}</span>
                  <span className="color-pending-count">⚠️ Color Pending: {stats.colorPending}</span>
                  <span className="repeated-count">★ Repeated Lots: {stats.repeatedLots}</span>
                  <span>Filters Active: {activeFilterCount}</span>
                </div>
              </div>

              {/* Glass KPI Cards */}
              <div className="header-kpi-grid">
                <div className="header-kpi-card">
                  <span className="header-kpi-value">{stats.total}</span>
                  <span className="header-kpi-label">Total Lots</span>
                </div>
                <div className="header-kpi-card">
                  <span className="header-kpi-value" style={{ color: '#86efac' }}>{stats.embCompleted}</span>
                  <span className="header-kpi-label">EMB Done</span>
                </div>
                <div className="header-kpi-card">
                  <span className="header-kpi-value" style={{ color: '#fde68a' }}>{stats.colorPending}</span>
                  <span className="header-kpi-label">Color ⚠️</span>
                </div>
                <div className="header-kpi-card">
                  <span className="header-kpi-value" style={{ color: '#fca5a5' }}>{stats.critical}</span>
                  <span className="header-kpi-label">Critical</span>
                </div>
              </div>
            </div>

            {/* Action Buttons Bar */}
            <div className="header-actions">
              <button
                onClick={() => window.history.back()}
                style={{ padding: '10px 20px', background: 'rgba(255,255,255,0.15)', color: '#ffffff', border: '1.5px solid rgba(255,255,255,0.25)', borderRadius: '12px', fontWeight: '700', cursor: 'pointer', fontSize: '14px', backdropFilter: 'blur(8px)', transition: 'all 0.2s' }}
              >
                ← Back
              </button>

              <button
                onClick={fetchData}
                style={{ padding: '10px 20px', background: 'rgba(52,211,153,0.2)', color: '#86efac', border: '1.5px solid rgba(52,211,153,0.3)', borderRadius: '12px', fontWeight: '700', cursor: 'pointer', fontSize: '14px', backdropFilter: 'blur(8px)' }}
              >
                🔄 Refresh
              </button>

              <div style={{ position: 'relative', display: 'inline-block' }}>
                <button
                  onClick={() => setShowExportMenu(!showExportMenu)}
                  style={{ padding: '10px 20px', background: 'rgba(79,70,229,0.3)', color: '#c7d2fe', border: '1.5px solid rgba(99,102,241,0.4)', borderRadius: '12px', fontWeight: '700', cursor: 'pointer', fontSize: '14px', backdropFilter: 'blur(8px)' }}
                >
                  📥 Export
                </button>
                {showExportMenu && (
                  <div className="export-menu">
                    <button className="export-option" onClick={() => { exportToPDF(); setShowExportMenu(false); }}>
                      📄 PDF Report
                    </button>
                    <button className="export-option" onClick={() => { exportToCSV(); setShowExportMenu(false); }}>
                      📊 Excel (CSV)
                    </button>
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>

        {/* Body wrapper */}
        <div className="page-body">

        {/* Dropdown Filters Section */}
        <div className="dropdown-filters-section">
          <div className="filters-header">
            <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
              <span className="filters-title">Advanced Filters</span>
              {activeFilterCount > 0 && (
                <span className="active-filters-count">{activeFilterCount} active</span>
              )}
            </div>
          </div>

          <div className="dropdown-filters-grid">
            <div className="filter-group">
              <label className="filter-label">Brand</label>
              <MultiSelectDropdown
                placeholder="All Brands"
                options={availableBrands}
                selectedValues={brandFilter}
                onChange={setBrandFilter}
              />
            </div>

            <div className="filter-group">
              <label className="filter-label">Fabric</label>
              <MultiSelectDropdown
                placeholder="All Fabrics"
                options={availableFabrics}
                selectedValues={fabricFilter}
                onChange={setFabricFilter}
              />
            </div>

            <div className="filter-group">
              <label className="filter-label">Party</label>
              <MultiSelectDropdown
                placeholder="All Parties"
                options={availableParties}
                selectedValues={partyFilter}
                onChange={setPartyFilter}
              />
            </div>

            <div className="filter-group">
              <label className="filter-label">Garment Type</label>
              <MultiSelectDropdown
                placeholder="All Types"
                options={availableGarmentTypes}
                selectedValues={garmentTypeFilter}
                onChange={setGarmentTypeFilter}
              />
            </div>

            <div className="filter-group">
              <label className="filter-label">Style</label>
              <MultiSelectDropdown
                placeholder="All Styles"
                options={availableStyles}
                selectedValues={styleFilter}
                onChange={setStyleFilter}
              />
            </div>

            <div className="filter-group">
              <label className="filter-label">Season</label>
              <MultiSelectDropdown
                placeholder="All Seasons"
                options={availableSeasons}
                selectedValues={seasonFilter}
                onChange={setSeasonFilter}
              />
            </div>

            <div className="filter-group">
              <label className="filter-label">M/W/K</label>
              <MultiSelectDropdown
                placeholder="All"
                options={availableMwkValues}
                selectedValues={mwkFilter}
                onChange={setMwkFilter}
              />
            </div>

            <div className="filter-group">
              <label className="filter-label">Days Pending</label>
              <select
                className="filter-select"
                value={daysFilter}
                onChange={(e) => setDaysFilter(e.target.value)}
              >
                <option value="">All Days</option>
                <option value="urgent">Urgent (7-14 days)</option>
                <option value="critical">Critical (Less than 14 days)</option>
                <option value="7">More than 7 days</option>
                <option value="14">More than 14 days</option>
                <option value="30">More than 30 days</option>
              </select>
            </div>

            <div className="filter-group">
              <label className="filter-label" style={{ visibility: 'hidden' }}>Clear</label>
              <button
                className="clear-filters-btn"
                onClick={clearAllFilters}
                disabled={activeFilterCount === 0}
              >
                🗑️ Clear All Filters
              </button>
            </div>
          </div>

          {/* Search and quick filters */}
          <div style={{ marginTop: '20px', paddingTop: '20px', borderTop: '1px solid #e2e8f0' }}>
            <div className="search-input-wrapper">
              <span className="search-icon">🔍</span>
              <input
                type="text"
                className="search-input"
                placeholder="Search by Lot Number, Fabric, Brand, Party, Season, M/W/K..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: 16, marginTop: '16px', flexWrap: 'wrap' }}>
              <div className="filter-tabs">
                {filters.map(filter => (
                  <button
                    key={filter.id}
                    className={`filter-tab ${activeFilter === filter.id ? 'active' : ''}`}
                    onClick={() => setActiveFilter(filter.id)}
                  >
                    {filter.label}
                    <span className="filter-count">{filter.count}</span>
                  </button>
                ))}
              </div>
              <div className="view-toggle">
                <button
                  className={`view-btn ${viewMode === 'table' ? 'active' : ''}`}
                  onClick={() => setViewMode('table')}
                >
                  📋 Table
                </button>
                <button
                  className={`view-btn ${viewMode === 'card' ? 'active' : ''}`}
                  onClick={() => setViewMode('card')}
                >
                  🗂️ Cards
                </button>
              </div>
            </div>
          </div>
        </div>

        {/* Statistics */}
        <div className="stats-grid">
          <div className="stat-card">
            <div className="stat-value">{stats.total}</div>
            <div className="stat-label">
              Total Pending Lots
            </div>
          </div>
          <div className="stat-card pieces">
            <div className="stat-value">{stats.totalPieces.toLocaleString()}</div>
            <div className="stat-label">
              Total Pieces
            </div>
          </div>
          <div className="stat-card color-pending">
            <div className="stat-value" style={{ color: '#ff9800' }}>{stats.colorPending}</div>
            <div className="stat-label">
              Color Pending Lots
            </div>
          </div>
          <div className="stat-card repeated">
            <div className="stat-value" style={{ color: '#f57c00' }}>{stats.repeatedLots}</div>
            <div className="stat-label">
              Repeated Lots
            </div>
          </div>
          <div className="stat-card">
            <div className="stat-value" style={{ color: '#10b981' }}>{stats.embCompleted}</div>
            <div className="stat-label">
              EMB Completed
            </div>
          </div>
          <div className="stat-card">
            <div className="stat-value" style={{ color: '#f59e0b' }}>{stats.directStitching}</div>
            <div className="stat-label">
              Direct Stitching
            </div>
          </div>
          <div className={`stat-card ${stats.urgent > 0 ? 'urgent' : 'normal'}`}>
            <div className="stat-value" style={{ color: stats.urgent > 0 ? '#f59e0b' : '#10b981' }}>
              {stats.urgent}
            </div>
            <div className="stat-label">
              Urgent (7-14 days)
            </div>
          </div>
          <div className={`stat-card ${stats.critical > 0 ? 'critical' : 'normal'}`}>
            <div className="stat-value" style={{ color: stats.critical > 0 ? '#dc2626' : '#10b981' }}>
              {stats.critical}
            </div>
            <div className="stat-label">
              Critical (less than 14 days)
            </div>
          </div>
        </div>

        {/* Main Content */}
        {viewMode === 'table' ? <TableView /> : <CardView />}

        {/* Footer Info */}
        <div className="footer-info">
          <p>
            Showing <strong>{filteredIssues.length}</strong> of <strong>{issues.length}</strong> pending lots •
            Total Pieces: <strong>{stats.totalPieces.toLocaleString()}</strong> •
            Color Pending: <strong>{stats.colorPending}</strong> •
            Repeated Lots: <strong>{stats.repeatedLots}</strong> •
            Active Filters: <strong>{activeFilterCount}</strong>
          </p>
          <p style={{ fontSize: 12, marginTop: 8, color: '#94a3b8' }}>
            Last updated: {new Date().toLocaleTimeString('en-IN', {
              hour: '2-digit',
              minute: '2-digit',
              second: '2-digit'
            })} • Click any row to view details
          </p>
        </div>

        {/* Issue Detail Modal */}
        {selectedIssue && (
          <IssueDetailModal
            issue={selectedIssue}
            onClose={() => setSelectedIssue(null)}
          />
        )}

        {/* Image Lightbox Modal */}
        {viewImageSrc && (
          <div className="image-modal-backdrop" onClick={() => setViewImageSrc(null)}>
            <div className="image-modal-content" onClick={(e) => e.stopPropagation()}>
              <button className="image-modal-close" onClick={() => setViewImageSrc(null)}>&times;</button>
              <img src={viewImageSrc} alt="Full Preview" className="image-modal-img" />
            </div>
          </div>
        )}
        </div>
      </div>
    </>
  );
};

export default PendingIssue;