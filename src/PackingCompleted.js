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

  const handleToggle = (val) => {
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
          background: selectedValues.length > 0 ? "#eff6ff" : "#ffffff",
          borderColor: selectedValues.length > 0 ? "#3b82f6" : "#cbd5e1",
          fontWeight: selectedValues.length > 0 ? "700" : "600",
          color: selectedValues.length > 0 ? "#1d4ed8" : "#0f172a"
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
              <div
                key={opt.value}
                onClick={(e) => {
                  e.preventDefault();
                  e.stopPropagation();
                  handleToggle(opt.value);
                }}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: "8px",
                  padding: "8px 10px",
                  borderRadius: "8px",
                  cursor: "pointer",
                  fontSize: "0.85rem",
                  fontWeight: isChecked ? "700" : "500",
                  color: isChecked ? "#1d4ed8" : "#334155",
                  background: isChecked ? "#eff6ff" : "transparent",
                  transition: "background 0.15s ease"
                }}
              >
                <input
                  type="checkbox"
                  checked={isChecked}
                  readOnly
                  style={{ cursor: "pointer", accentColor: "#2563eb", pointerEvents: "none" }}
                />
                <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                  {opt.label}
                </span>
              </div>
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

const PackingCompleted = () => {
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [packingData, setPackingData] = useState([]);
  const [issuesData, setIssuesData] = useState([]);
  const [issuesLotMap, setIssuesLotMap] = useState(new Map());
  const [rawpackData, setRawpackData] = useState([]);
  const [rawpackLotMap, setRawpackLotMap] = useState(new Map());
  const [barcodeData, setBarcodeData] = useState([]);
  const [barcodeLotMap, setBarcodeLotMap] = useState(new Map());

  const [searchTerm, setSearchTerm] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const [itemsPerPage] = useState(50);
  const [exportLoading, setExportLoading] = useState(false);
  const [showFilters, setShowFilters] = useState(false);
  const [viewImageSrc, setViewImageSrc] = useState(null);
  const [financialYearFilter, setFinancialYearFilter] = useState("ALL");

  const [filters, setFilters] = useState({
    packingSupervisor: [],
    stitchingSupervisor: [],
    brand: [],
    fabric: [],
    garmentType: [],
    style: [],
    season: [],
    partyName: [],
    source: [],
    agingRange: []
  });

  const API_KEY = GOOGLE_API_KEY;
  const SPREADSHEET_ID = SPREADSHEET_IDS.MAIN;
  const RANGE = `${SHEET_NAMES.INDEX}!A:AA`;

  const ISSUES_SPREADSHEET_ID = SPREADSHEET_IDS.ISSUES;
  const ISSUES_SPREADSHEET_RANGE = `${SHEET_NAMES.ISSUES}!A:R`;

  const RAWPACK_SPREADSHEET_ID = SPREADSHEET_IDS.RAWPACK || '1xD8Uy1lUgvNTQ2RGRBI4ZjOrozbinUPRq2_UfIplP98';
  const RAWPACK_SPREADSHEET_RANGE = `${SHEET_NAMES.RAWPACK || 'RAWPACK'}!A:ZZ`;

  const BARCODE_SPREADSHEET_ID = SPREADSHEET_IDS.BARCODE || '1dOCjNFwaAel5qun0_ZJVIGmREqjI76CJBBFIjM3NHv8';
  const BARCODE_SPREADSHEET_RANGE = `${SHEET_NAMES.BARCODE || 'LotBarcodeData'}!A:Z`;

  const handleGoBack = () => {
    window.history.back();
  };

  const parseDate = (dateString) => {
    if (!dateString || dateString === '-') return null;
    try {
      if (typeof dateString === 'string' && dateString.includes('/')) {
        const parts = dateString.split('/');
        if (parts.length === 3) {
          const day = parseInt(parts[0], 10);
          const month = parseInt(parts[1], 10) - 1;
          let year = parseInt(parts[2], 10);
          if (year < 100) year = year <= 50 ? 2000 + year : 1900 + year;
          if (!isNaN(day) && !isNaN(month) && !isNaN(year) &&
            day > 0 && day <= 31 && month >= 0 && month < 12 && year > 1900) {
            return new Date(year, month, day);
          }
        }
      }
      if (typeof dateString === 'string' && dateString.includes('-')) {
        const parts = dateString.split('T')[0].split('-');
        if (parts.length === 3) {
          let year = parseInt(parts[0], 10);
          let month = parseInt(parts[1], 10) - 1;
          let day = parseInt(parts[2], 10);
          if (parts[0].length <= 2 && parts[2].length === 4) {
            day = parseInt(parts[0], 10);
            month = parseInt(parts[1], 10) - 1;
            year = parseInt(parts[2], 10);
          }
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

  const formatDate = (dateString) => {
    if (!dateString || dateString === '-') return '-';
    try {
      if (typeof dateString === 'string' && dateString.includes('-')) {
        const parts = dateString.split('T')[0].split('-');
        if (parts.length === 3 && /^\d{4}$/.test(parts[0])) {
          return `${parts[2].padStart(2, '0')}/${parts[1].padStart(2, '0')}/${parts[0]}`;
        }
      }
      if (typeof dateString === 'string' && dateString.includes('/')) {
        const parts = dateString.split('/');
        if (parts.length === 3) {
          const year = parts[2].length === 2 ? `20${parts[2]}` : parts[2];
          return `${parts[0].padStart(2, '0')}/${parts[1].padStart(2, '0')}/${year}`;
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

  // Aging calculation formula: Packing Complete Date - Date of Issue (in days)
  const calculateAgingDays = (issueDateStr, completeDateStr) => {
    if (!issueDateStr || !completeDateStr || issueDateStr === '-' || completeDateStr === '-') return null;
    const d1 = parseDate(issueDateStr);
    const d2 = parseDate(completeDateStr);
    if (!d1 || !d2) return null;
    d1.setHours(0, 0, 0, 0);
    d2.setHours(0, 0, 0, 0);
    const diffTime = d2.getTime() - d1.getTime();
    const diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24));
    return isNaN(diffDays) ? null : diffDays;
  };

  const isPackingComplete = (val) => {
    if (!val || val === '' || val === '-') return false;
    if (typeof val === 'string') {
      if (/\d/.test(val)) return true;
      if (val.toLowerCase() === 'yes') return true;
    }
    return false;
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
    const headers = values[0].map(h => (h || '').toString().trim().toLowerCase());
    const rows = values.slice(1);

    const lotNumberIndex = headers.findIndex(h => h.includes('lot number') || h.includes('lot') || h === 'lot');
    const packingSupervisorIndex = headers.findIndex(h => h.includes('packing supervisor') || h.includes('supervisor'));
    const packingDateIndex = headers.findIndex(h => h.includes('packing date'));
    const packingCompleteIndex = headers.findIndex(h => h.includes('packing complete'));
    const totalPcsIndex = headers.findIndex(h => h.includes('total pcs') || h.includes('total pieces') || h === 'total pcs' || h.includes('pcs'));
    const stitchingSupervisorIndex = headers.findIndex(h => h.includes('stitching supervisor'));
    const brandIndex = headers.findIndex(h => h.includes('brand'));
    const seasonIndex = headers.findIndex(h => h.includes('season'));

    const issuesData = [];
    const lotMap = new Map();

    rows.forEach((row, index) => {
      if (row[lotNumberIndex] && row[lotNumberIndex].toString().trim() !== '') {
        const lotNumber = row[lotNumberIndex].toString().trim();
        const packingCompleteValue = packingCompleteIndex !== -1 ? row[packingCompleteIndex] : '';

        const issueItem = {
          id: `issues-${index}`,
          lotNumber: lotNumber,
          packingSupervisor: packingSupervisorIndex !== -1 ? row[packingSupervisorIndex] || '' : '',
          packingDate: packingDateIndex !== -1 ? row[packingDateIndex] || '' : '',
          packingComplete: packingCompleteValue,
          totalPcs: totalPcsIndex !== -1 ? row[totalPcsIndex] || '0' : '0',
          stitchingSupervisor: stitchingSupervisorIndex !== -1 ? row[stitchingSupervisorIndex] || '' : '',
          brand: brandIndex !== -1 ? row[brandIndex] || '' : '',
          season: seasonIndex !== -1 ? row[seasonIndex] || '' : ''
        };

        issuesData.push(issueItem);
        lotMap.set(lotNumber, issueItem);
        lotMap.set(lotNumber.toUpperCase(), issueItem);
      }
    });

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
    const supervisorCol = getCol('packing person') >= 0 ? getCol('packing person') : getCol('supervisor');
    const pcsCol = getCol('pcs') >= 0 ? getCol('pcs') : getCol('total pcs');

    const rawpackData = [];
    const rawpackLotMap = new Map();

    rows.forEach((row, idx) => {
      const lotVal1 = lotCol >= 0 ? (row[lotCol] || '').toString().trim() : '';
      const lotVal2 = lot2Col >= 0 ? (row[lot2Col] || '').toString().trim() : '';
      const lotNumber = lotVal1 || lotVal2;

      if (lotNumber && lotNumber !== '-' && lotNumber !== '0') {
        const packingCompleteDateVal = packingCompleteDateCol >= 0 ? (row[packingCompleteDateCol] || '').toString().trim() : '';
        const supervisor = supervisorCol >= 0 ? (row[supervisorCol] || '').toString().trim() : '';
        const pcs = pcsCol >= 0 ? (row[pcsCol] || '').toString().trim() : '';

        const item = {
          id: `rawpack-${idx}`,
          lotNumber: lotNumber,
          packingCompleteDate: packingCompleteDateVal,
          packingSupervisor: supervisor,
          totalPcs: pcs,
          hasCompleteDate: Boolean(packingCompleteDateVal && packingCompleteDateVal !== '-' && packingCompleteDateVal !== '#N/A' && packingCompleteDateVal !== '00/01/00')
        };

        rawpackData.push(item);
        rawpackLotMap.set(lotNumber, item);
        rawpackLotMap.set(lotNumber.toUpperCase(), item);
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
    const brandCol = headers.findIndex(h => h.includes('brand'));
    const styleCol = headers.findIndex(h => h.includes('style'));

    const barcodeData = [];
    const barcodeLotMap = new Map();

    rows.forEach((row, idx) => {
      const lotVal = lotCol >= 0 ? (row[lotCol] || '').toString().trim() : '';
      if (!lotVal || lotVal === '-' || lotVal === '0') return;

      const genDateRaw = genDateCol >= 0 ? (row[genDateCol] || '').toString().trim() : '';
      const barcodeId = barcodeIdCol >= 0 ? (row[barcodeIdCol] || '').toString().trim() : '';
      const status = statusCol >= 0 ? (row[statusCol] || '').toString().trim() : '';
      const brand = brandCol >= 0 ? (row[brandCol] || '').toString().trim() : '';
      const style = styleCol >= 0 ? (row[styleCol] || '').toString().trim() : '';

      const normalizedLot = lotVal.toUpperCase();
      const hasDate = genDateRaw && genDateRaw !== '-' && genDateRaw !== '#N/A';

      const existing = barcodeLotMap.get(normalizedLot);
      if (!existing) {
        const item = {
          id: `barcode-${idx}`,
          lotNumber: lotVal,
          barcodeId: barcodeId,
          status: status,
          brand: brand,
          style: style,
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
    const sheetHeaders = values[0].map(h => (h || '').toString().trim().toLowerCase());
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
      imageUrl: sheetHeaders.findIndex(h => h.includes('image url') || h === 'image'),
      season: sheetHeaders.findIndex(h => h.includes('season')),
      mwk: sheetHeaders.findIndex(h => h.includes('m/w/k')),
      cuttingQty: sheetHeaders.findIndex(h => h.includes('cutting qty') || h.includes('pcs'))
    };

    return rows.reduce((acc, row, index) => {
      const lotNumber = headerIndices.lotNumber !== -1 ? row[headerIndices.lotNumber] : '';
      if (!lotNumber || lotNumber.toString().trim() === '') return acc;

      const item = { id: index, originalRow: index + 2 };
      Object.entries(headerIndices).forEach(([fieldName, headerIndex]) => {
        item[fieldName] = (headerIndex !== -1 && row[headerIndex] !== undefined) ? (row[headerIndex] || '').toString().trim() : '';
      });

      if (item.dateOfIssue) item.dateOfIssue = formatDate(item.dateOfIssue);

      acc.push(item);
      return acc;
    }, []);
  };

  const fetchAllSheetData = async () => {
    try {
      setLoading(true);
      setError(null);

      const [indexData, issuesSheetData, rawpackSheetData, barcodeSheetData] = await Promise.all([
        fetchSheetData(SPREADSHEET_ID, RANGE),
        fetchSheetData(ISSUES_SPREADSHEET_ID, ISSUES_SPREADSHEET_RANGE),
        fetchSheetData(RAWPACK_SPREADSHEET_ID, RAWPACK_SPREADSHEET_RANGE).catch(() => ({ values: [] })),
        fetchSheetData(BARCODE_SPREADSHEET_ID, BARCODE_SPREADSHEET_RANGE).catch(() => ({ values: [] }))
      ]);

      const transformedIndexData = transformSheetData(indexData.values || []);
      const { issuesData: transformedIssuesData, lotMap } = transformIssuesData(issuesSheetData.values || []);
      const { rawpackData: transformedRawpackData, rawpackLotMap: rawpackMap } = transformRawpackData(rawpackSheetData.values || []);
      const { barcodeData: transformedBarcodeData, barcodeLotMap: barcodeMap } = transformBarcodeData(barcodeSheetData.values || []);

      setPackingData(transformedIndexData);
      setIssuesData(transformedIssuesData);
      setIssuesLotMap(lotMap);
      setRawpackData(transformedRawpackData);
      setRawpackLotMap(rawpackMap);
      setBarcodeData(transformedBarcodeData);
      setBarcodeLotMap(barcodeMap);
    } catch (err) {
      console.error('Error loading packing completed data:', err);
      setError(err.message || 'Failed to fetch data');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchAllSheetData();
  }, []);

  // Filter completed packing lots
  const completedLots = useMemo(() => {
    if (!packingData.length && !issuesData.length && !barcodeData.length) return [];

    const indexMap = new Map();
    packingData.forEach(item => {
      if (item.lotNumber) indexMap.set(item.lotNumber.toUpperCase(), item);
    });

    const allLotNumbers = new Set();
    issuesData.forEach(i => i.lotNumber && allLotNumbers.add(i.lotNumber.toUpperCase()));
    rawpackData.forEach(r => r.lotNumber && allLotNumbers.add(r.lotNumber.toUpperCase()));
    barcodeData.forEach(b => b.lotNumber && allLotNumbers.add(b.lotNumber.toUpperCase()));
    packingData.forEach(p => p.lotNumber && allLotNumbers.add(p.lotNumber.toUpperCase()));

    const results = [];

    allLotNumbers.forEach(lotKey => {
      const indexItem = indexMap.get(lotKey) || {};
      const issuesInfo = issuesLotMap.get(lotKey);
      const rawpackInfo = rawpackLotMap.get(lotKey);
      const barcodeInfo = barcodeLotMap.get(lotKey);

      // Determine Packing Complete Date from available sources (Barcode > Issues > RAWPACK)
      let completeDate = '';
      let sourceName = '';

      if (barcodeInfo && barcodeInfo.isCompleted && barcodeInfo.packingCompleteDate) {
        completeDate = barcodeInfo.packingCompleteDate;
        sourceName = 'Barcode Data';
      } else if (issuesInfo && isPackingComplete(issuesInfo.packingComplete)) {
        completeDate = formatDate(issuesInfo.packingComplete);
        sourceName = 'Issues Sheet';
      } else if (rawpackInfo && rawpackInfo.hasCompleteDate) {
        completeDate = formatDate(rawpackInfo.packingCompleteDate);
        sourceName = 'RAWPACK Sheet';
      }

      // If lot is packing completed, construct row
      if (completeDate && completeDate !== '-') {
        const lotNumber = indexItem.lotNumber || issuesInfo?.lotNumber || barcodeInfo?.lotNumber || rawpackInfo?.lotNumber || lotKey;
        const totalPcs = issuesInfo?.totalPcs || rawpackInfo?.totalPcs || indexItem.cuttingQty || '-';
        const packingSupervisor = issuesInfo?.packingSupervisor || rawpackInfo?.packingSupervisor || '-';
        const dateOfIssue = issuesInfo?.packingDate ? formatDate(issuesInfo.packingDate) : (indexItem.dateOfIssue || '-');
        const agingDays = calculateAgingDays(dateOfIssue, completeDate);

        results.push({
          lotNumber: lotNumber,
          partyName: indexItem.partyName || '-',
          fabric: indexItem.fabric || '-',
          brand: indexItem.brand || barcodeInfo?.brand || issuesInfo?.brand || '-',
          garmentType: indexItem.garmentType || '-',
          style: indexItem.style || barcodeInfo?.style || '-',
          supervisor: packingSupervisor,
          stitchingSupervisor: indexItem.supervisor || issuesInfo?.stitchingSupervisor || '-',
          dateOfIssue: dateOfIssue,
          totalPcs: totalPcs,
          packingCompleteDate: completeDate,
          agingDays: agingDays,
          completionSource: sourceName,
          barcodeId: barcodeInfo?.barcodeId || '-',
          mwk: indexItem.mwk || '-',
          season: indexItem.season || issuesInfo?.season || '-',
          imageUrl: indexItem.imageUrl || '',
          directStitching: indexItem.directStitching || '-'
        });
      }
    });

    return results;
  }, [packingData, issuesData, rawpackData, barcodeData, issuesLotMap, rawpackLotMap, barcodeLotMap]);

  // Unique filter options
  const filterOptions = useMemo(() => {
    const getOptions = (getter, defaultLabel) => {
      const set = new Set();
      completedLots.forEach(item => {
        const val = (getter(item) || '').toString().trim();
        if (val && val !== '-' && val !== '0' && val !== '--') {
          set.add(val);
        }
      });
      const sorted = Array.from(set).sort((a, b) => a.localeCompare(b));
      return [
        { value: 'all', label: defaultLabel },
        ...sorted.map(val => ({ value: val, label: val }))
      ];
    };

    return {
      packingSupervisor: getOptions(item => item.supervisor, 'All Supervisors'),
      stitchingSupervisor: getOptions(item => item.stitchingSupervisor, 'All Stitching Supervisors'),
      brand: getOptions(item => item.brand, 'All Brands'),
      fabric: getOptions(item => item.fabric, 'All Fabrics'),
      garmentType: getOptions(item => item.garmentType, 'All Garment Types'),
      style: getOptions(item => item.style, 'All Styles'),
      season: getOptions(item => item.season, 'All Seasons'),
      partyName: getOptions(item => item.partyName, 'All Parties'),
      source: getOptions(item => item.completionSource, 'All Sources'),
      agingRange: [
        { value: 'all', label: 'All Aging' },
        { value: '0-3', label: '0-3 Days (Fast)' },
        { value: '4-7', label: '4-7 Days (Normal)' },
        { value: '8-14', label: '8-14 Days (Moderate)' },
        { value: '15+', label: '15+ Days (Delayed)' }
      ]
    };
  }, [completedLots]);

  const allFinancialYears = useMemo(() => {
    const set = new Set();
    set.add(getCurrentFinancialYear());
    completedLots.forEach(item => {
      const fy = getFinancialYearFromDate(item.packingCompleteDate);
      if (fy) set.add(fy);
    });
    const sorted = Array.from(set).sort().reverse();
    return ["ALL", ...sorted];
  }, [completedLots]);

  const matchesAgingRange = (days, rangeVal) => {
    if (rangeVal === 'all') return true;
    if (days === null || days === undefined) return false;
    switch (rangeVal) {
      case '0-3': return days >= 0 && days <= 3;
      case '4-7': return days >= 4 && days <= 7;
      case '8-14': return days >= 8 && days <= 14;
      case '15+': return days >= 15;
      default: return true;
    }
  };

  // Reset to page 1 whenever any filter or search changes
  useEffect(() => {
    setCurrentPage(1);
  }, [filters, searchTerm, financialYearFilter]);

  // Filtered dataset
  const displayData = useMemo(() => {
    let filtered = completedLots;

    if (filters.packingSupervisor.length > 0) {
      const selected = filters.packingSupervisor.map(s => s.toLowerCase().trim());
      filtered = filtered.filter(item => selected.includes((item.supervisor || '').toLowerCase().trim()));
    }
    if (filters.stitchingSupervisor.length > 0) {
      const selected = filters.stitchingSupervisor.map(s => s.toLowerCase().trim());
      filtered = filtered.filter(item => selected.includes((item.stitchingSupervisor || '').toLowerCase().trim()));
    }
    if (filters.brand.length > 0) {
      const selected = filters.brand.map(s => s.toLowerCase().trim());
      filtered = filtered.filter(item => selected.includes((item.brand || '').toLowerCase().trim()));
    }
    if (filters.fabric.length > 0) {
      const selected = filters.fabric.map(s => s.toLowerCase().trim());
      filtered = filtered.filter(item => selected.includes((item.fabric || '').toLowerCase().trim()));
    }
    if (filters.garmentType.length > 0) {
      const selected = filters.garmentType.map(s => s.toLowerCase().trim());
      filtered = filtered.filter(item => selected.includes((item.garmentType || '').toLowerCase().trim()));
    }
    if (filters.style.length > 0) {
      const selected = filters.style.map(s => s.toLowerCase().trim());
      filtered = filtered.filter(item => selected.includes((item.style || '').toLowerCase().trim()));
    }
    if (filters.season.length > 0) {
      const selected = filters.season.map(s => s.toLowerCase().trim());
      filtered = filtered.filter(item => selected.includes((item.season || '').toLowerCase().trim()));
    }
    if (filters.partyName.length > 0) {
      const selected = filters.partyName.map(s => s.toLowerCase().trim());
      filtered = filtered.filter(item => selected.includes((item.partyName || '').toLowerCase().trim()));
    }
    if (filters.source.length > 0) {
      const selected = filters.source.map(s => s.toLowerCase().trim());
      filtered = filtered.filter(item => selected.includes((item.completionSource || '').toLowerCase().trim()));
    }
    if (filters.agingRange.length > 0) {
      filtered = filtered.filter(item =>
        filters.agingRange.some(rangeVal => matchesAgingRange(item.agingDays, rangeVal))
      );
    }

    if (searchTerm) {
      const s = searchTerm.toLowerCase().trim();
      filtered = filtered.filter(item =>
        ['lotNumber', 'fabric', 'brand', 'garmentType', 'style', 'partyName', 'supervisor', 'season', 'barcodeId', 'completionSource']
          .some(field => item[field] && item[field].toString().toLowerCase().includes(s))
      );
    }

    if (financialYearFilter !== "ALL") {
      filtered = filtered.filter(item => {
        const fy = getFinancialYearFromDate(item.packingCompleteDate);
        return fy === financialYearFilter;
      });
    }

    return filtered;
  }, [completedLots, filters, searchTerm, financialYearFilter]);

  // Pagination
  const paginatedData = useMemo(() => {
    const start = (currentPage - 1) * itemsPerPage;
    return displayData.slice(start, start + itemsPerPage);
  }, [displayData, currentPage, itemsPerPage]);

  const totalPages = Math.ceil(displayData.length / itemsPerPage);

  const totalCompletedPcs = useMemo(() => {
    return displayData.reduce((acc, item) => {
      const num = parseInt(item.totalPcs, 10);
      return acc + (isNaN(num) ? 0 : num);
    }, 0);
  }, [displayData]);

  // Export to Excel
  const exportToExcel = () => {
    setExportLoading(true);
    try {
      const rows = displayData.map((item, idx) => ({
        '#': idx + 1,
        'Lot Number': item.lotNumber,
        'Party Name': item.partyName,
        'Fabric': item.fabric,
        'Brand': item.brand,
        'Garment Type': item.garmentType,
        'Style': item.style,
        'Packing Supervisor': item.supervisor,
        'Date of Issue': item.dateOfIssue,
        'Total PCS': item.totalPcs,
        'Packing Complete Date': item.packingCompleteDate,
        'Aging (Days)': item.agingDays !== null ? `${item.agingDays} Days` : '-',
        'Source': item.completionSource,
        'Barcode ID': item.barcodeId,
        'M/W/K': item.mwk,
        'Season': item.season
      }));

      const ws = XLSX.utils.json_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Packing Completed Lots');
      XLSX.writeFile(wb, `Packing_Completed_Lots_${new Date().toISOString().slice(0, 10)}.xlsx`);
    } catch (e) {
      console.error('Excel export error:', e);
      alert('Failed to export Excel file');
    } finally {
      setExportLoading(false);
    }
  };

  // Export to PDF
  const exportToPDF = () => {
    setExportLoading(true);
    try {
      const doc = new jsPDF('landscape');
      doc.setFontSize(16);
      doc.setTextColor(30, 58, 138); // Deep Blue
      doc.text('PACKING COMPLETED LOTS REPORT', 14, 15);
      doc.setFontSize(9);
      doc.setTextColor(71, 85, 105);
      doc.text(`Generated on: ${new Date().toLocaleDateString()} | Total Completed Lots: ${displayData.length} | Total PCS: ${totalCompletedPcs.toLocaleString()}`, 14, 22);

      const tableData = displayData.map((item, idx) => [
        idx + 1,
        item.lotNumber,
        item.partyName,
        item.fabric,
        item.brand,
        item.style,
        item.supervisor,
        item.dateOfIssue,
        item.totalPcs,
        item.packingCompleteDate,
        item.agingDays !== null ? `${item.agingDays} d` : '-',
        item.completionSource
      ]);

      autoTable(doc, {
        head: [['#', 'Lot No', 'Party', 'Fabric', 'Brand', 'Style', 'Supervisor', 'Issue Date', 'PCS', 'Comp Date', 'Aging', 'Source']],
        body: tableData,
        startY: 26,
        styles: { fontSize: 7.5, cellPadding: 2, halign: 'center' },
        headStyles: { fillColor: [30, 58, 138], textColor: [255, 255, 255], fontStyle: 'bold' },
        columnStyles: {
          1: { halign: 'left', fontStyle: 'bold' },
          2: { halign: 'left' },
          4: { halign: 'left' }
        }
      });

      doc.save(`Packing_Completed_Lots_${new Date().toISOString().slice(0, 10)}.pdf`);
    } catch (e) {
      console.error('PDF export error:', e);
      alert('Failed to export PDF file');
    } finally {
      setExportLoading(false);
    }
  };

  // Helper function for aging badge styling
  const getAgingBadge = (days) => {
    if (days === null || days === undefined) {
      return <span style={{ color: '#94a3b8' }}>—</span>;
    }
    if (days <= 3) {
      return (
        <span style={{ background: '#dcfce7', color: '#15803d', padding: '3px 8px', borderRadius: '10px', fontWeight: '800', fontSize: '0.75rem', border: '1px solid #86efac' }}>
          {days} {days === 1 ? 'Day' : 'Days'}
        </span>
      );
    }
    if (days <= 7) {
      return (
        <span style={{ background: '#e0f2fe', color: '#0369a1', padding: '3px 8px', borderRadius: '10px', fontWeight: '800', fontSize: '0.75rem', border: '1px solid #7dd3fc' }}>
          {days} Days
        </span>
      );
    }
    if (days <= 14) {
      return (
        <span style={{ background: '#fef3c7', color: '#b45309', padding: '3px 8px', borderRadius: '10px', fontWeight: '800', fontSize: '0.75rem', border: '1px solid #fde047' }}>
          {days} Days
        </span>
      );
    }
    return (
      <span style={{ background: '#fee2e2', color: '#b91c1c', padding: '3px 8px', borderRadius: '10px', fontWeight: '800', fontSize: '0.75rem', border: '1px solid #fca5a5' }}>
        {days} Days
      </span>
    );
  };

  return (
    <div className="pending-packing-container">
      {/* Header Section */}
      <div className="header-section" style={{ background: 'linear-gradient(135deg, #1e1b4b 0%, #1e3a8a 45%, #2563eb 100%)', boxShadow: '0 20px 40px -15px rgba(30, 58, 138, 0.35)' }}>
        <div className="title-section">
          <button className="back-button" onClick={handleGoBack}>
            ← Back
          </button>
          <h2>
            📦✅ Packing Completed Lots Report
          </h2>
        </div>

        {/* Actions Row */}
        <div className="actions">
          <input
            type="text"
            placeholder="🔍 Search by Lot, Fabric, Brand, Style, Party, Supervisor, Barcode..."
            value={searchTerm}
            onChange={(e) => {
              setSearchTerm(e.target.value);
              setCurrentPage(1);
            }}
            className="search-input"
          />

          <button
            onClick={() => setShowFilters(!showFilters)}
            className={`filter-toggle-btn ${showFilters ? 'active' : ''}`}
            title="Toggle Filters"
          >
            🎚️ Filters {showFilters ? '▲' : '▼'} {Object.values(filters).some(arr => arr.length > 0) && '●'}
          </button>

          <div className="export-buttons">
            <button
              onClick={exportToExcel}
              disabled={exportLoading || displayData.length === 0}
              className="export-btn excel-btn"
              title="Export to Excel"
            >
              {exportLoading ? '⏳' : '📊'} Excel
            </button>

            <button
              onClick={exportToPDF}
              disabled={exportLoading || displayData.length === 0}
              className="export-btn pdf-btn"
              title="Export to PDF"
            >
              {exportLoading ? '⏳' : '📄'} PDF
            </button>

            <button
              onClick={fetchAllSheetData}
              disabled={loading}
              className="refresh-btn"
              title="Refresh Data"
            >
              🔄 Refresh
            </button>
          </div>
        </div>
      </div>

      {/* Filter Section */}
      {showFilters && (
        <div className="filters-section">
          <div className="filters-header">
            <h3>Filter Completed Packing Lots</h3>
            {Object.values(filters).some(arr => arr.length > 0) && (
              <button
                className="clear-all-btn"
                onClick={() => setFilters({
                  packingSupervisor: [],
                  stitchingSupervisor: [],
                  brand: [],
                  fabric: [],
                  garmentType: [],
                  style: [],
                  season: [],
                  partyName: [],
                  source: [],
                  agingRange: []
                })}
              >
                ✕ Clear All Filters
              </button>
            )}
          </div>

          <div className="filters-grid">
            <MultiSelectDropdown
              label="Packing Supervisor"
              options={filterOptions.packingSupervisor}
              selectedValues={filters.packingSupervisor}
              onChange={(val) => setFilters(prev => ({ ...prev, packingSupervisor: val }))}
            />
            <MultiSelectDropdown
              label="Stitching Supervisor"
              options={filterOptions.stitchingSupervisor}
              selectedValues={filters.stitchingSupervisor}
              onChange={(val) => setFilters(prev => ({ ...prev, stitchingSupervisor: val }))}
            />
            <MultiSelectDropdown
              label="Brand"
              options={filterOptions.brand}
              selectedValues={filters.brand}
              onChange={(val) => setFilters(prev => ({ ...prev, brand: val }))}
            />
            <MultiSelectDropdown
              label="Fabric"
              options={filterOptions.fabric}
              selectedValues={filters.fabric}
              onChange={(val) => setFilters(prev => ({ ...prev, fabric: val }))}
            />
            <MultiSelectDropdown
              label="Garment Type"
              options={filterOptions.garmentType}
              selectedValues={filters.garmentType}
              onChange={(val) => setFilters(prev => ({ ...prev, garmentType: val }))}
            />
            <MultiSelectDropdown
              label="Style"
              options={filterOptions.style}
              selectedValues={filters.style}
              onChange={(val) => setFilters(prev => ({ ...prev, style: val }))}
            />
            <MultiSelectDropdown
              label="Season"
              options={filterOptions.season}
              selectedValues={filters.season}
              onChange={(val) => setFilters(prev => ({ ...prev, season: val }))}
            />
            <MultiSelectDropdown
              label="Party Name"
              options={filterOptions.partyName}
              selectedValues={filters.partyName}
              onChange={(val) => setFilters(prev => ({ ...prev, partyName: val }))}
            />
            <MultiSelectDropdown
              label="Aging (Days)"
              options={filterOptions.agingRange}
              selectedValues={filters.agingRange}
              onChange={(val) => setFilters(prev => ({ ...prev, agingRange: val }))}
            />
            <MultiSelectDropdown
              label="Data Source"
              options={filterOptions.source}
              selectedValues={filters.source}
              onChange={(val) => setFilters(prev => ({ ...prev, source: val }))}
            />

            <div className="filter-item">
              <label>Financial Year</label>
              <select
                value={financialYearFilter}
                onChange={(e) => setFinancialYearFilter(e.target.value)}
                className="filter-select"
              >
                {allFinancialYears.map(fy => (
                  <option key={fy} value={fy}>{fy === 'ALL' ? 'All Financial Years' : `FY ${fy}`}</option>
                ))}
              </select>
            </div>
          </div>
        </div>
      )}

      {/* Main Table Content */}
      {loading ? (
        <div className="loading-spinner">
          <div className="spinner" style={{ borderTopColor: '#2563eb' }}></div>
          <p style={{ color: '#1e3a8a', fontWeight: '700' }}>Loading Packing Completed Lots Data...</p>
        </div>
      ) : error ? (
        <div className="error-message">
          <span>⚠️</span>
          <p>{error}</p>
          <button onClick={fetchAllSheetData} style={{ marginLeft: 'auto', padding: '8px 16px', background: '#dc2626', color: '#fff', border: 'none', borderRadius: '8px', cursor: 'pointer', fontWeight: '700' }}>Retry</button>
        </div>
      ) : displayData.length === 0 ? (
        <div style={{ padding: '48px', textAlign: 'center', background: '#ffffff', borderRadius: '24px', border: '1px solid #e2e8f0', boxShadow: '0 10px 25px rgba(0,0,0,0.03)' }}>
          <span style={{ fontSize: '3rem' }}>🔍</span>
          <h3 style={{ margin: '16px 0 8px', color: '#1e293b' }}>No Completed Packing Lots Found</h3>
          <p style={{ color: '#64748b' }}>Try adjusting your search criteria or filter options</p>
        </div>
      ) : (
        <div className="table-container">
          <table className="packing-table">
            <thead>
              <tr>
                <th style={{ textAlign: 'center', width: '50px' }}>#</th>
                <th style={{ textAlign: 'center', width: '70px' }}>IMAGE</th>
                <th>LOT NUMBER</th>
                <th>PARTY NAME</th>
                <th>FABRIC</th>
                <th>BRAND</th>
                <th>GARMENT TYPE</th>
                <th>STYLE</th>
                <th>SUPERVISOR</th>
                <th style={{ textAlign: 'center' }}>DATE OF ISSUE</th>
                <th style={{ textAlign: 'center' }}>TOTAL PCS</th>
                <th style={{ textAlign: 'center', background: '#1d4ed8' }}>PACKING COMPLETE DATE</th>
                <th style={{ textAlign: 'center', background: '#2563eb' }}>AGING</th>
                <th style={{ textAlign: 'center' }}>BARCODE ID</th>
                <th style={{ textAlign: 'center' }}>SOURCE</th>
                <th style={{ textAlign: 'center' }}>M/W/K</th>
                <th style={{ textAlign: 'center' }}>SEASON</th>
              </tr>
            </thead>
            <tbody>
              {paginatedData.map((item, index) => {
                const rowIndex = (currentPage - 1) * itemsPerPage + index + 1;
                const directImg = getDirectImageUrl(item.imageUrl);
                return (
                  <tr key={item.lotNumber + '-' + index}>
                    <td style={{ textAlign: 'center', fontWeight: '700', color: '#64748b' }}>
                      {rowIndex}
                    </td>
                    <td style={{ textAlign: 'center', padding: '6px' }}>
                      {directImg ? (
                        <img
                          src={directImg}
                          alt="Lot preview"
                          onClick={() => setViewImageSrc(directImg)}
                          style={{ width: '38px', height: '38px', objectFit: 'cover', borderRadius: '8px', cursor: 'pointer', border: '1.5px solid #cbd5e1' }}
                          onError={(e) => { e.target.style.display = 'none'; }}
                        />
                      ) : (
                        <span style={{ color: '#94a3b8', fontSize: '0.75rem' }}>—</span>
                      )}
                    </td>
                    <td>
                      <span className="lot-number" style={{ background: '#dbeafe', color: '#1e40af' }}>
                        {item.lotNumber}
                      </span>
                    </td>
                    <td style={{ fontWeight: '600', color: '#1e293b' }}>{item.partyName}</td>
                    <td style={{ color: '#334155' }}>{item.fabric}</td>
                    <td style={{ fontWeight: '700', color: '#0f172a' }}>{item.brand}</td>
                    <td style={{ color: '#475569' }}>{item.garmentType}</td>
                    <td style={{ color: '#475569' }}>{item.style}</td>
                    <td style={{ fontWeight: '600', color: '#1e40af' }}>{item.supervisor}</td>
                    <td style={{ textAlign: 'center', color: '#475569' }}>{item.dateOfIssue}</td>
                    <td style={{ textAlign: 'center', fontWeight: '700', color: '#0f172a' }}>
                      {item.totalPcs}
                    </td>
                    <td style={{ textAlign: 'center' }}>
                      <span style={{ background: '#dbeafe', color: '#1d4ed8', padding: '5px 12px', borderRadius: '12px', fontWeight: '800', fontSize: '0.85rem', display: 'inline-flex', alignItems: 'center', gap: '4px', border: '1px solid #bfdbfe' }}>
                        ✓ {item.packingCompleteDate}
                      </span>
                    </td>
                    <td style={{ textAlign: 'center' }}>
                      {getAgingBadge(item.agingDays)}
                    </td>
                    <td style={{ textAlign: 'center', fontFamily: 'monospace', fontSize: '0.8rem', color: '#2563eb', fontWeight: '700' }}>
                      {item.barcodeId}
                    </td>
                    <td style={{ textAlign: 'center' }}>
                      <span style={{ fontSize: '0.75rem', padding: '3px 8px', borderRadius: '6px', background: item.completionSource === 'Barcode Data' ? '#dbeafe' : '#fef3c7', color: item.completionSource === 'Barcode Data' ? '#1e40af' : '#92400e', fontWeight: '700' }}>
                        {item.completionSource}
                      </span>
                    </td>
                    <td style={{ textAlign: 'center', color: '#64748b' }}>{item.mwk}</td>
                    <td style={{ textAlign: 'center', color: '#64748b' }}>{item.season}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>

          {/* Pagination Controls */}
          {totalPages > 1 && (
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '16px 24px', background: '#f8fafc', borderTop: '1px solid #e2e8f0', position: 'sticky', bottom: 0, zIndex: 5 }}>
              <span style={{ fontSize: '0.85rem', color: '#64748b', fontWeight: '600' }}>
                Showing {(currentPage - 1) * itemsPerPage + 1} to {Math.min(currentPage * itemsPerPage, displayData.length)} of {displayData.length} records
              </span>
              <div style={{ display: 'flex', gap: '8px' }}>
                <button
                  onClick={() => setCurrentPage(p => Math.max(p - 1, 1))}
                  disabled={currentPage === 1}
                  style={{ padding: '6px 14px', borderRadius: '8px', border: '1px solid #cbd5e1', background: currentPage === 1 ? '#f1f5f9' : '#ffffff', cursor: currentPage === 1 ? 'not-allowed' : 'pointer', fontWeight: '700' }}
                >
                  Previous
                </button>
                <span style={{ padding: '6px 14px', fontWeight: '700', color: '#1e40af' }}>
                  Page {currentPage} of {totalPages}
                </span>
                <button
                  onClick={() => setCurrentPage(p => Math.min(p + 1, totalPages))}
                  disabled={currentPage === totalPages}
                  style={{ padding: '6px 14px', borderRadius: '8px', border: '1px solid #cbd5e1', background: currentPage === totalPages ? '#f1f5f9' : '#ffffff', cursor: currentPage === totalPages ? 'not-allowed' : 'pointer', fontWeight: '700' }}
                >
                  Next
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {/* Image Modal Lightbox */}
      {viewImageSrc && (
        <div
          onClick={() => setViewImageSrc(null)}
          style={{
            position: 'fixed',
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            background: 'rgba(0, 0, 0, 0.75)',
            backdropFilter: 'blur(8px)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 10000,
            padding: '24px'
          }}
        >
          <div style={{ position: 'relative', maxWidth: '800px', maxHeight: '80vh' }} onClick={e => e.stopPropagation()}>
            <button
              onClick={() => setViewImageSrc(null)}
              style={{
                position: 'absolute',
                top: '-16px',
                right: '-16px',
                width: '36px',
                height: '36px',
                borderRadius: '50%',
                background: '#ef4444',
                color: '#ffffff',
                border: 'none',
                fontWeight: 'bold',
                cursor: 'pointer',
                fontSize: '1.1rem',
                boxShadow: '0 4px 12px rgba(0,0,0,0.3)'
              }}
            >
              ✕
            </button>
            <img
              src={viewImageSrc}
              alt="Preview"
              style={{ width: '100%', maxHeight: '75vh', objectFit: 'contain', borderRadius: '16px', background: '#fff', boxShadow: '0 25px 50px -12px rgba(0,0,0,0.5)' }}
            />
          </div>
        </div>
      )}
    </div>
  );
};

export default PackingCompleted;
