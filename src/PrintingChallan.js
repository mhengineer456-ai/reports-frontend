import React, { useEffect, useMemo, useState } from "react";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { GOOGLE_API_KEY, SPREADSHEET_IDS, fetchSheetDataFromBackend } from './config';
import { fetchRemarksForTab, saveRemarkForLot } from './embPrintRemarksService';

const MultiSelectDropdown = ({ options, selectedValues, onChange, placeholder }) => {
  const [isOpen, setIsOpen] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const containerRef = React.useRef(null);

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

  return (
    <div className="multiselect-container" ref={containerRef} style={{ position: 'relative', width: '100%' }}>
      <div
        className="multiselect-select"
        onClick={() => setIsOpen(!isOpen)}
        style={{
          padding: '10px 14px',
          background: 'white',
          border: '1px solid #cbd5e1',
          borderRadius: '8px',
          cursor: 'pointer',
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          fontSize: '14px',
          color: selectedValues.length > 0 ? '#1e293b' : '#94a3b8',
          minHeight: '38px',
          boxSizing: 'border-box'
        }}
      >
        <span style={{
          whiteSpace: 'nowrap',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          maxWidth: '85%',
          display: 'block'
        }}>
          {selectedValues.length > 0
            ? `${selectedValues.length} selected`
            : placeholder}
        </span>
        <span style={{ fontSize: '10px', color: '#64748b' }}>{isOpen ? '▲' : '▼'}</span>
      </div>

      {isOpen && (
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
            boxShadow: '0 10px 15px -3px rgba(0, 0, 0, 0.1)',
            zIndex: 1000,
            maxHeight: '220px',
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
            placeholder="Search..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            style={{
              padding: '6px 10px',
              border: '1px solid #cbd5e1',
              borderRadius: '6px',
              width: '100%',
              fontSize: '13px',
              marginBottom: '6px',
              boxSizing: 'border-box',
              outline: 'none'
            }}
          />
          <div style={{ display: 'flex', gap: '8px', marginBottom: '8px' }}>
            <button
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
                const isSelected = selectedValues.includes(option);
                return (
                  <label
                    key={option}
                    style={{
                      display: 'flex',
                      alignItems: 'center',
                      gap: '8px',
                      padding: '6px 8px',
                      borderRadius: '4px',
                      cursor: 'pointer',
                      background: isSelected ? '#f1f5f9' : 'transparent',
                      fontSize: '13px',
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

function buildImageMap(indexValues) {
  const map = {};
  if (!Array.isArray(indexValues)) return map;
  const [headers, ...rows] = indexValues;
  if (!headers) return map;

  let imageCol = -1;
  headers.forEach((header, index) => {
    const headerLower = header.trim().toLowerCase();
    if (headerLower === 'image url' || headerLower === 'imageurl' || headerLower === 'image' || headerLower.includes('image')) {
      imageCol = index;
    }
  });

  const lotCol = headers.findIndex(h => h.trim().toLowerCase() === 'lot number');

  if (imageCol === -1 || lotCol === -1) return map;

  rows.forEach((row) => {
    const lot = String(row?.[lotCol] || "").trim();
    const url = String(row?.[imageCol] || "").trim();
    if (lot && url) {
      map[lot] = getDirectImageUrl(url);
    }
  });

  return map;
}

/* ================== CONFIG ================== */
const SHEET_ID = SPREADSHEET_IDS.JOBORDER;
const API_KEY = GOOGLE_API_KEY;
const TAB_NAME = "JobOrder";
const RANGE = "A1:ZZZ";

const BUDGET_SHEET_ID = SPREADSHEET_IDS.MAIN;
const INDEX_SHEET_NAME = "Index";
const CUTTING_SHEET_NAME = "Cutting";

// --- Design Image sheet config ---
const DESIGN_SHEET_NAME = "Design Image";
const API_URL_DESIGN =
  `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${DESIGN_SHEET_NAME}!A1:ZZZ?key=${API_KEY}`;

const API_URL_JOB =
  `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${TAB_NAME}!${RANGE}?key=${API_KEY}`;
const API_URL_INDEX =
  `https://sheets.googleapis.com/v4/spreadsheets/${BUDGET_SHEET_ID}/values/${INDEX_SHEET_NAME}!A1:ZZZ?key=${API_KEY}`;
const API_URL_CUTTING =
  `https://sheets.googleapis.com/v4/spreadsheets/${BUDGET_SHEET_ID}/values/${CUTTING_SHEET_NAME}!A1:ZZZ?key=${API_KEY}`;

/* ================== DISPLAY COLUMNS ================== */
const DISPLAY_HEADERS = [
  "Lot Number",
  "Garment Type",
  "Section",
  "Style",
  "Fabric",
  "Brand",
  "Printing",
  "Party Name",
  "Priority", // Added Priority for repeated lots
  "Challan No",
  "Challan Date",
  "Challan Total Qty",
];

const RemarksHistoryModal = ({ isOpen, onClose, lotNumber, remarksHistory }) => {
  if (!isOpen) return null;

  return (
    <div
      style={{
        position: 'fixed',
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        backgroundColor: 'rgba(15, 23, 42, 0.65)',
        backdropFilter: 'blur(4px)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        zIndex: 99999,
        padding: '20px'
      }}
      onClick={onClose}
    >
      <div
        style={{
          background: '#ffffff',
          borderRadius: '16px',
          maxWidth: '540px',
          width: '100%',
          maxHeight: '80vh',
          display: 'flex',
          flexDirection: 'column',
          boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.25)',
          overflow: 'hidden'
        }}
        onClick={(e) => e.stopPropagation()}
      >
        <div
          style={{
            background: 'linear-gradient(135deg, #1e1b4b 0%, #312e81 100%)',
            color: '#ffffff',
            padding: '16px 20px',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between'
          }}
        >
          <div>
            <h3 style={{ margin: 0, fontSize: '1.1rem', fontWeight: '800' }}>
              📝 Remarks History
            </h3>
            <p style={{ margin: '2px 0 0 0', fontSize: '0.82rem', color: '#c7d2fe' }}>
              Lot Number: <strong>#{lotNumber}</strong>
            </p>
          </div>
          <button
            onClick={onClose}
            style={{
              background: 'rgba(255, 255, 255, 0.15)',
              border: 'none',
              color: '#ffffff',
              fontSize: '1.2rem',
              borderRadius: '8px',
              width: '32px',
              height: '32px',
              cursor: 'pointer',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center'
            }}
          >
            ✕
          </button>
        </div>

        <div style={{ padding: '20px', overflowY: 'auto', flex: 1, display: 'flex', flexDirection: 'column', gap: '10px' }}>
          {(!remarksHistory || remarksHistory.length === 0) ? (
            <div style={{ textAlign: 'center', padding: '30px', color: '#94a3b8' }}>
              <div style={{ fontSize: '2rem', marginBottom: '8px' }}>📭</div>
              <p style={{ margin: 0, fontWeight: '600' }}>No remarks history found for this lot.</p>
            </div>
          ) : (
            remarksHistory.map((item, idx) => (
              <div
                key={idx}
                style={{
                  background: '#f8fafc',
                  border: '1px solid #e2e8f0',
                  borderLeft: '4px solid #6366f1',
                  borderRadius: '10px',
                  padding: '12px 14px',
                  display: 'flex',
                  flexDirection: 'column',
                  gap: '4px'
                }}
              >
                <div style={{ fontSize: '0.92rem', fontWeight: '700', color: '#0f172a', lineHeight: '1.4' }}>
                  {item.text}
                </div>
                <div style={{ fontSize: '0.78rem', color: '#6366f1', fontWeight: '600', display: 'flex', alignItems: 'center', gap: '4px' }}>
                  <span>🕒</span> {item.timestamp}
                </div>
              </div>
            ))
          )}
        </div>

        <div
          style={{
            padding: '14px 20px',
            background: '#f8fafc',
            borderTop: '1px solid #e2e8f0',
            display: 'flex',
            justifyContent: 'flex-end'
          }}
        >
          <button
            onClick={onClose}
            style={{
              background: '#334155',
              color: '#ffffff',
              border: 'none',
              padding: '8px 20px',
              borderRadius: '8px',
              fontWeight: '700',
              fontSize: '0.88rem',
              cursor: 'pointer'
            }}
          >
            Close
          </button>
        </div>
      </div>
    </div>
  );
};

// Desired front-order
const PRIORITY_HEADERS = [
  "Lot Number",
  "Fabric",
  "Brand",
  "Style",
  "Section",
  "Garment Type",
  "Party Name",
  "Priority",
];

// Final order: priority first, then the rest in their original order
const ORDERED_DISPLAY_HEADERS = [
  ...PRIORITY_HEADERS,
  ...DISPLAY_HEADERS.filter((h) => !PRIORITY_HEADERS.includes(h)),
];

// Label mapping (display only)
const headerLabel = (h) => (h === "Lot Number" ? "Lot No." : h);

// Filter headers (the ones you want dropdowns for)
const FILTER_HEADERS = [
  "Fabric",
  "Brand",
  "Style",
  "Garment Type",
  "Section",
  "Party Name",
  "Printing",
  "Priority"
];

/* ================== REPEATED LOT HELPER ================== */
// Function to check if a lot is repeated based on Priority column
const isRepeatedLot = (priorityValue) => {
  if (!priorityValue) return false;
  const priority = String(priorityValue).trim().toUpperCase();
  return priority.includes("REPEATED_LOT");
};

/* ================== SHADES HELPERS ================== */
const parseShadeList = (s) => {
  if (!s) return [];

  const shadeString = String(s).trim();
  if (!shadeString) return [];

  // Split by comma, slash, or newline
  return shadeString
    .split(/[,/\n]/)
    .map((x) => {
      let shade = x.trim();

      // Remove quantity suffix (like "-50", "- 50", " -50", etc.)
      shade = shade.replace(/\s*-\s*\d+\s*$/, '').trim();

      // Normalize: lowercase
      shade = shade.toLowerCase();

      // Handle common variations
      shade = shade
        .replace(/\./g, ' ')      // Replace dots with spaces
        .replace(/gray/g, 'grey') // Standardize spelling
        .replace(/\s+/g, ' ')     // Normalize spaces
        .trim();

      // Remove content in parentheses for comparison
      shade = shade.replace(/\([^)]*\)/g, '').trim();

      return shade;
    })
    .filter(x => x &&
      !['null', 'undefined', 'na', 'n/a', 'none', '-', ''].includes(x));
};

// Helper function to parse shades while preserving original names
function parseOriginalShadeList(s) {
  if (!s) return [];

  const shadeString = String(s).trim();
  if (!shadeString) return [];

  // Split by comma, slash, or newline
  return shadeString
    .split(/[,/\n]/)
    .map(x => x.trim())
    .filter(x => x &&
      !['null', 'undefined', 'na', 'n/a', 'none', '-', ''].includes(x.toLowerCase()));
}

function buildShadeStateMap(indexValues, cuttingValues) {
  const result = { cancelledByLot: {}, emptySizeByLot: {} };
  if (!Array.isArray(indexValues) || !Array.isArray(cuttingValues)) return result;

  const [idxHeaders, ...idxRows] = indexValues;
  const h = (name) => idxHeaders.indexOf(name);

  const lotCol = h("Lot Number");
  const startRowCol = h("StartRow");
  const numRowsCol = h("NumRows");

  idxRows.forEach((row) => {
    const lot = String(row?.[lotCol] || "").trim();
    const startRow = parseInt(row?.[startRowCol], 10);
    const numRows = parseInt(row?.[numRowsCol], 10);
    if (!lot || !Number.isFinite(startRow) || !Number.isFinite(numRows)) return;

    const start = Math.max(0, startRow - 1);
    const end = Math.min(cuttingValues.length, start + numRows);
    const region = cuttingValues.slice(start, end);

    const headerIdx = region.findIndex(
      (r) => String(r?.[0] || "").trim().toLowerCase() === "color"
    );
    if (headerIdx === -1) return;

    const headerRow = region[headerIdx];
    const totalPcsCol = headerRow.findIndex(
      (x) => String(x || "").trim().toLowerCase() === "total pcs"
    );

    const sizeStart = 2;
    const sizeEnd = totalPcsCol > sizeStart ? totalPcsCol : headerRow.length;
    const sizeColIdxs = [];
    for (let i = sizeStart; i < sizeEnd; i++) {
      const name = String(headerRow[i] || "").trim();
      if (name) sizeColIdxs.push(i);
    }

    const cancelled = new Set();
    const emptySizes = new Set();

    for (let i = headerIdx + 1; i < region.length; i++) {
      const shadeName = String(region[i]?.[0] || "").trim();
      if (!shadeName) continue;
      const low = shadeName.toLowerCase();
      if (low === "total") break;

      // Normalize shade name - MUST MATCH parseShadeList normalization
      const normalizedShade = low
        .replace(/\([^)]*\)/g, '')  // Remove parentheses content
        .replace(/\./g, ' ')        // Replace dots with spaces
        .replace(/gray/g, 'grey')   // Standardize spelling
        .replace(/\s+/g, ' ')       // Normalize spaces
        .trim();

      const values = sizeColIdxs.map((ci) => {
        const val = region[i]?.[ci];
        if (val === undefined || val === null) return "";
        return String(val).trim();
      });

      // Check if all values are zeros (or zero-like)
      const allZeros = values.length > 0 && values.every((v) => {
        const num = parseFloat(v.replace(/\s/g, ''));
        return !isNaN(num) && num === 0;
      });

      // Check if any value is empty (for Cutting Pending)
      const hasEmpty = values.some((v) => v === "" || v === null || v === undefined);

      if (allZeros) cancelled.add(normalizedShade);
      if (hasEmpty) emptySizes.add(normalizedShade);
    }

    if (cancelled.size) result.cancelledByLot[lot] = cancelled;
    if (emptySizes.size) result.emptySizeByLot[lot] = emptySizes;
  });

  return result;
}

/* ================== DESIGN IMAGE HELPERS ================== */
function buildDesignDateMap(designValues) {
  const map = {};
  if (!Array.isArray(designValues)) return map;
  const [headers, ...rows] = designValues;
  if (!headers) return map;

  const h = (name) => headers.indexOf(name);
  const tsCol = h("Timestamp");
  const lotCol = h("LOT NUMBER");

  rows.forEach((r) => {
    const lot = String(r?.[lotCol] || "").trim();
    const tsRaw = r?.[tsCol];
    const d = parseDateSafe(tsRaw);
    if (!lot || !d) return;
    if (!map[lot]) map[lot] = [];
    map[lot].push(d);
  });

  Object.keys(map).forEach((lot) => map[lot].sort((a, b) => a - b));
  return map;
}

/** Find the earliest design timestamp that is >= receiveDate */
function getDesignDateOnOrAfter(designMap, lot, receiveDate) {
  if (!lot || !receiveDate) return null;
  const arr = designMap[lot];
  if (!Array.isArray(arr) || arr.length === 0) return null;
  const recvMidnight = new Date(receiveDate);
  recvMidnight.setHours(0, 0, 0, 0);
  for (const d of arr) {
    const dd = new Date(d); dd.setHours(0, 0, 0, 0);
    if (dd >= recvMidnight) return d;
  }
  return null;
}

function formatDDMMMYYYY(d) {
  if (!d || isNaN(d)) return "";
  return d.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
}

function shadesFromHistory(challanHistoryJson) {
  try {
    const history = JSON.parse(challanHistoryJson || "[]");
    if (!Array.isArray(history)) return [];

    const shades = [];
    history.forEach((h) => {
      if (h && Array.isArray(h.items)) {
        h.items.forEach((it) => {
          if (it && it.shade !== undefined && it.shade !== null) {
            let shade = String(it.shade).trim();

            // Apply SAME normalization as parseShadeList
            shade = shade.replace(/\s*-\s*\d+\s*$/, '').trim();
            shade = shade.toLowerCase();
            shade = shade
              .replace(/\./g, ' ')
              .replace(/gray/g, 'grey')
              .replace(/\s+/g, ' ')
              .trim();
            shade = shade.replace(/\([^)]*\)/g, '').trim();

            shades.push(shade);
          }
        });
      }
    });
    return shades;
  } catch {
    return [];
  }
}

function computeStatus(shadeString, challanHistoryJson, cancelledSet) {
  const all = parseShadeList(shadeString);

  // Filter out cancelled shades
  const effective = all.filter((s) => !(cancelledSet && cancelledSet.has(s)));

  // If all shades are cancelled
  if (all.length > 0 && effective.length === 0) {
    return { text: "All Shades Cancelled", type: "neutral" };
  }

  // If no effective shades after filtering cancelled ones
  if (!effective.length) return { text: "Comp.Challan Created", type: "success" };

  const pending = getPendingShadesArray(shadeString, challanHistoryJson, cancelledSet);

  return pending.length === 0
    ? { text: "Comp.Challan Created", type: "success" }
    : { text: "Partial pending", type: "error" };
}

function getPendingShadesArray(shadeString, challanHistoryJson, cancelledSet) {
  // First, parse the original shade list
  const originalShades = parseOriginalShadeList(shadeString);

  // Create normalized versions for comparison
  const normalizedShades = originalShades.map(original => {
    let normalized = original.toLowerCase().trim();

    // Remove quantity suffix for comparison only
    normalized = normalized.replace(/\s*-\s*\d+\s*$/, '').trim();
    normalized = normalized
      .replace(/\./g, ' ')
      .replace(/gray/g, 'grey')
      .replace(/\s+/g, ' ')
      .trim();

    // Remove content in parentheses for comparison
    normalized = normalized.replace(/\([^)]*\)/g, '').trim();

    return { original, normalized };
  });

  // Filter out cancelled shades
  const effectiveShades = normalizedShades.filter(({ normalized }) =>
    !(cancelledSet && cancelledSet.has(normalized))
  );

  if (effectiveShades.length === 0) return [];

  // Get done shades from history
  const doneShades = shadesFromHistory(challanHistoryJson);
  const done = new Set(doneShades);

  // Return original shade names that are pending AND not cancelled
  return effectiveShades
    .filter(({ normalized }) => !done.has(normalized))
    .map(({ original }) => original.trim());
}

function getRemarks({ lot, shadeString, challanHistoryJson, cancelledByLot, emptySizeByLot }) {
  const cancelledSet = cancelledByLot[lot];
  const emptySet = emptySizeByLot[lot];

  // Get all shades and filter out cancelled ones
  const allShades = parseShadeList(shadeString);
  const effectiveShades = allShades.filter(
    (s) => !(cancelledSet && cancelledSet.has(s))
  );

  // If all shades are cancelled
  if (allShades.length > 0 && effectiveShades.length === 0) {
    return "All Shades Cancelled";
  }

  // If no effective shades
  if (effectiveShades.length === 0) return "";

  // Check for empty sizes in effective shades only
  const hasEmpty = effectiveShades.some((s) => emptySet && emptySet.has(s));
  if (hasEmpty) return "Cutting Pending";

  // Get pending shades (already excludes cancelled ones)
  const pending = getPendingShadesArray(shadeString, challanHistoryJson, cancelledSet);
  if (pending.length > 0) return "Challan Pending";

  return "";
}

/* ================== OTHER HELPERS ================== */
const getPrintingStatus = (challanHistoryJson) => {
  if (!challanHistoryJson) return { text: "Unknown", type: "neutral" };
  try {
    const history = JSON.parse(challanHistoryJson);
    if (!Array.isArray(history)) return { text: "Invalid History", type: "error" };
    const allCompleted = history.every((e) => e.embCompleted === true);
    return allCompleted
      ? { text: "Printing Done", type: "success" }
      : { text: "Printing Pending", type: "warning" };
  } catch {
    return { text: "Error Parsing JSON", type: "error" };
  }
};

const getLatestPrintingUpdatedAt = (challanHistoryJson) => {
  try {
    const history = JSON.parse(challanHistoryJson || "[]");
    if (!Array.isArray(history)) return "";
    const dates = history
      .map((e) => new Date(e.embUpdatedAt))
      .filter((d) => !isNaN(d));
    if (!dates.length) return "";
    const latest = new Date(Math.max(...dates));
    return latest.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
  } catch {
    return "";
  }
};

const getLatestPrintingUpdatedAtObj = (challanHistoryJson) => {
  try {
    const history = JSON.parse(challanHistoryJson || "[]");
    if (!Array.isArray(history)) return null;
    const dates = history
      .map((e) => new Date(e.embUpdatedAt))
      .filter((d) => !isNaN(d));
    if (!dates.length) return null;
    return new Date(Math.max(...dates));
  } catch {
    return null;
  }
};

/* ======= dates + day-diff ======= */
function parseDateSafe(input) {
  if (!input) return null;
  if (input instanceof Date && !isNaN(input)) return input;
  const s = String(input).trim();
  if (!s) return null;

  const d1 = new Date(s);
  if (!isNaN(d1)) return d1;

  const m1 = s.match(/^(\d{1,2})[\/](\d{1,2})[\/](\d{2,4})$/);
  if (m1) {
    const [_, dd, mm, yyyy] = m1;
    const y = Number(yyyy.length === 2 ? (Number(yyyy) + 2000) : yyyy);
    const d = new Date(y, Number(mm) - 1, Number(dd));
    return isNaN(d) ? null : d;
  }

  const m2 = s.match(/^(\d{1,2})[- ]([A-Za-z]{3})[- ](\d{2,4})$/);
  if (m2) {
    const monMap = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
    const dd = Number(m2[1]);
    const mon = monMap[m2[2].toLowerCase()];
    let y = Number(m2[3]);
    if (y < 100) y += 2000;
    const d = new Date(y, mon, dd);
    return isNaN(d) ? null : d;
  }

  return null;
}

function daysBetween(startDate, endDate) {
  const start = parseDateSafe(startDate);
  const end = parseDateSafe(endDate);
  if (!start || !end) return "";
  const ms = end.setHours(0, 0, 0, 0) - start.setHours(0, 0, 0, 0);
  const days = Math.ceil(ms / 86400000);
  return Math.max(0, days);
}

const getDaysValForRow = (row, cancelledByLot) => {
  if (!row) return "";
  const lot = String(row["Lot Number"] || "").trim();
  const cancelledSet = cancelledByLot?.[lot];
  const status = computeStatus(row["Shade"], row["Challan History JSON"], cancelledSet);
  const printingStatus = getPrintingStatus(row["Challan History JSON"]);
  const receiveObj =
    status.type === "success" && printingStatus.type === "success"
      ? getLatestPrintingUpdatedAtObj(row["Challan History JSON"])
      : null;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  return daysBetween(row["Challan Date"], (receiveObj || today));
};


/* ================== LOADERS ================== */
const loadScriptOnce = (src, globalVar) =>
  new Promise((resolve, reject) => {
    if (globalVar && window[globalVar]) return resolve(window[globalVar]);
    const existing = document.querySelector(`script[src="${src}"]`);
    if (existing) {
      existing.addEventListener("load", () => resolve(globalVar ? window[globalVar] : true));
      existing.addEventListener("error", reject);
      return;
    }
    const s = document.createElement("script");
    s.src = src;
    s.async = true;
    s.onload = () => resolve(globalVar ? window[globalVar] : true);
    s.onerror = () => reject(new Error(`Failed to load ${src}`));
    document.head.appendChild(s);
  });

async function ensureXLSX() {
  try {
    const mod = await import("xlsx");
    return mod.default || mod;
  } catch {
    await loadScriptOnce("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js", "XLSX");
    return window.XLSX;
  }
}

const downloadBlob = (blob, filename) => {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  URL.revokeObjectURL(url);
  a.remove();
};

/* ================== NEW HELPER FUNCTIONS ================== */
// Function to map Section values to M/W/K
const mapSectionValue = (section) => {
  if (!section) return "";
  const sectionStr = String(section).trim().toLowerCase();
  if (sectionStr.includes("gent") || sectionStr === "m") return "M";
  if (sectionStr.includes("girl") || sectionStr.includes("wom") || sectionStr === "w") return "W";
  if (sectionStr.includes("kid") || sectionStr === "k") return "K";
  return sectionStr.toUpperCase().substring(0, 1);
};

// Function to abbreviate Party Name to initials
const abbreviatePartyName = (partyName) => {
  if (!partyName) return "";
  const name = String(partyName).trim();

  const cleanName = name
    .replace(/\b(pvt|ltd|limited|llp|inc|corporation|corp)\b/gi, '')
    .replace(/[^\w\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  const words = cleanName.split(/\s+/);
  if (words.length === 0) return name.substring(0, 3).toUpperCase();

  const initials = words.map(word => word.charAt(0).toUpperCase()).join('');

  if (initials.length <= 1) return name.substring(0, 3).toUpperCase();
  if (initials.length <= 4) return initials;
  return initials.substring(0, 3);
};

/* ================== FINANCIAL YEAR HELPERS ================== */
const getFinancialYearFromDate = (dateStr) => {
  if (!dateStr) return null;
  let d = null;
  if (dateStr instanceof Date) {
    d = dateStr;
  } else if (typeof dateStr === "string") {
    const s = dateStr.trim();
    if (!s) return null;
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

/* ================== COMPONENT ================== */
export default function PrintingChallan({ initialPrintingStatusFilter = "pending", isEmbedded = false }) {
  const [challanData, setChallanData] = useState([]);
  const [cancelledByLot, setCancelledByLot] = useState({});
  const [emptySizeByLot, setEmptySizeByLot] = useState({});
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [searchTerm, setSearchTerm] = useState("");
  const [designDatesByLot, setDesignDatesByLot] = useState({});
  const [lotImages, setLotImages] = useState({});
  const [viewImageSrc, setViewImageSrc] = useState(null);

  // Filter states
  const [fabricFilter, setFabricFilter] = useState([]);
  const [brandFilter, setBrandFilter] = useState([]);
  const [styleFilter, setStyleFilter] = useState([]);
  const [garmentTypeFilter, setGarmentTypeFilter] = useState([]);
  const [sectionFilter, setSectionFilter] = useState([]);
  const [partyNameFilter, setPartyNameFilter] = useState([]);
  const [printingFilterDropdown, setPrintingFilterDropdown] = useState([]);
  const [priorityFilter, setPriorityFilter] = useState([]);

  // Printing Status filter
  const [printingStatusFilter, setPrintingStatusFilter] = useState(initialPrintingStatusFilter);
  const [daysFilter, setDaysFilter] = useState("all");
  const [financialYearFilter, setFinancialYearFilter] = useState(getCurrentFinancialYear());

  // Custom Head Remarks History state (Stored 100% in Google Sheets)
  const [headRemarksMap, setHeadRemarksMap] = useState({});
  const [newRemarkInput, setNewRemarkInput] = useState({});
  const [remarksModalOpen, setRemarksModalOpen] = useState(false);
  const [selectedRemarksLot, setSelectedRemarksLot] = useState(null);
  const [selectedRemarksHistory, setSelectedRemarksHistory] = useState([]);

  const openRemarksModal = (lot, history) => {
    setSelectedRemarksLot(lot);
    setSelectedRemarksHistory(history || []);
    setRemarksModalOpen(true);
  };

  useEffect(() => {
    fetchRemarksForTab('PRINT').then(map => {
      if (map && typeof map === 'object') {
        setHeadRemarksMap(map);
      }
    });

    const handleRemarkUpdated = (e) => {
      if (e.detail && e.detail.tabType === 'PRINT') {
        setHeadRemarksMap(prev => ({
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

  const handleSaveHeadRemark = (lot, rowData = {}) => {
    const text = String(newRemarkInput[lot] || "").trim();
    if (!text || !lot) return;
    saveRemarkForLot({
      tabType: 'PRINT',
      lotNumber: lot,
      challanNo: rowData["Challan No"] || "",
      partyName: rowData["Party Name"] || "",
      fabric: rowData["Fabric"] || "",
      style: rowData["Style"] || "",
      remarkText: text
    });
    setNewRemarkInput(prev => ({ ...prev, [lot]: "" }));
  };

  const redZoneCount = useMemo(() => {
    return challanData.filter((row) => {
      const daysVal = getDaysValForRow(row, cancelledByLot);
      return typeof daysVal === "number" && daysVal > 5;
    }).length;
  }, [challanData, cancelledByLot]);


  const loadAllData = async () => {
    setLoading(true);
    setError("");
    try {
      const [jobRes, idxRes, cutRes, designRes] = await Promise.all([
        fetchSheetDataFromBackend(SPREADSHEET_IDS.JOBORDER, `${TAB_NAME}!${RANGE}`),
        fetchSheetDataFromBackend(SPREADSHEET_IDS.MAIN, `${INDEX_SHEET_NAME}!A1:ZZZ`),
        fetchSheetDataFromBackend(SPREADSHEET_IDS.MAIN, `${CUTTING_SHEET_NAME}!A1:ZZZ`),
        fetchSheetDataFromBackend(SPREADSHEET_IDS.JOBORDER, `${DESIGN_SHEET_NAME}!A1:ZZZ`),
      ]);

      if (!jobRes || !Array.isArray(jobRes.values) || jobRes.values.length === 0) {
        if (!jobRes || !jobRes.ok) throw new Error("Failed to fetch JobOrder");
      }

      const [headers, ...rows] = jobRes.values || [];
      const formatted = (rows || [])
        .map((row) => {
          const entry = {};
          (headers || []).forEach((header, i) => (entry[header] = row[i] ?? ""));
          return entry;
        })
        .filter((row) => row["Challan No"]?.trim().startsWith("CH-PRINT-"));

      let cancelledMap = {};
      let emptyMap = {};
      let lotImagesMap = {};

      if (idxRes && idxRes.ok && cutRes && cutRes.ok) {
        const { cancelledByLot: cMap, emptySizeByLot: eMap } = buildShadeStateMap(
          idxRes.values || [],
          cutRes.values || []
        );
        cancelledMap = cMap;
        emptyMap = eMap;
        lotImagesMap = buildImageMap(idxRes.values || []);
      }

      if (designRes && designRes.ok) {
        const designMap = buildDesignDateMap(designRes.values || []);
        setDesignDatesByLot(designMap);
      } else {
        setDesignDatesByLot({});
      }

      setChallanData(formatted);
      setCancelledByLot(cancelledMap);
      setEmptySizeByLot(emptyMap);
      setLotImages(lotImagesMap);
    } catch (e) {
      console.error(e);
      setError("Fetch error. Check API key / sheet IDs / network.");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadAllData();
  }, []);

  const matchesPrintingStatusFilter = (statusText) => {
    switch (printingStatusFilter) {
      case "done":
        return statusText === "Printing Done";
      case "pending":
        return statusText === "Printing Pending";
      case "unknown":
        return statusText !== "Printing Done" && statusText !== "Printing Pending";
      default:
        return true;
    }
  };

  const availableFYs = useMemo(() => {
    const set = new Set();
    const currentFY = getCurrentFinancialYear();
    if (currentFY) set.add(currentFY);
    challanData.forEach((row) => {
      const fy = getFinancialYearFromDate(row["Challan Date"]);
      if (fy) set.add(fy);
    });
    return Array.from(set).sort().reverse();
  }, [challanData]);

  const filteredData = useMemo(() => {
    return challanData.filter((row) => {
      if (financialYearFilter && financialYearFilter !== "ALL") {
        const rowFY = getFinancialYearFromDate(row["Challan Date"]);
        if (rowFY !== financialYearFilter) return false;
      }
      // Text search
      const textMatch = Object.values(row).some((value) =>
        String(value).toLowerCase().includes(searchTerm.toLowerCase())
      );
      if (!textMatch) return false;

      // Printing status filter
      const printingStatus = getPrintingStatus(row["Challan History JSON"]).text;
      if (!matchesPrintingStatusFilter(printingStatus)) return false;

      // Header dropdown filters with normalization
      const normalizeValue = (val) =>
        String(val || "").trim()
          .toLowerCase()
          .replace(/\s+/g, '')
          .replace(/[^a-z0-9]/g, '');

      if (fabricFilter && fabricFilter.length > 0) {
        const rowFabric = normalizeValue(row["Fabric"]);
        const matched = fabricFilter.some(f => normalizeValue(f) === rowFabric);
        if (!matched) return false;
      }

      if (brandFilter && brandFilter.length > 0) {
        const rowBrand = normalizeValue(row["Brand"]);
        const matched = brandFilter.some(b => normalizeValue(b) === rowBrand);
        if (!matched) return false;
      }

      if (styleFilter && styleFilter.length > 0) {
        const rowStyle = normalizeValue(row["Style"]);
        const matched = styleFilter.some(s => normalizeValue(s) === rowStyle);
        if (!matched) return false;
      }

      if (garmentTypeFilter && garmentTypeFilter.length > 0) {
        const rowGarmentType = normalizeValue(row["Garment Type"]);
        const matched = garmentTypeFilter.some(g => normalizeValue(g) === rowGarmentType);
        if (!matched) return false;
      }

      if (sectionFilter && sectionFilter.length > 0) {
        const rowSection = normalizeValue(row["Section"]);
        const matched = sectionFilter.some(s => normalizeValue(s) === rowSection);
        if (!matched) return false;
      }

      if (partyNameFilter && partyNameFilter.length > 0) {
        const rowPartyName = normalizeValue(row["Party Name"]);
        const matched = partyNameFilter.some(p => normalizeValue(p) === rowPartyName);
        if (!matched) return false;
      }

      if (printingFilterDropdown && printingFilterDropdown.length > 0) {
        const rowPrinting = normalizeValue(row["Printing"]);
        const matched = printingFilterDropdown.some(p => normalizeValue(p) === rowPrinting);
        if (!matched) return false;
      }

      if (priorityFilter && priorityFilter.length > 0) {
        const rowPriority = normalizeValue(row["Priority"]);
        const matched = priorityFilter.some(p => normalizeValue(p) === rowPriority);
        if (!matched) return false;
      }

      if (daysFilter && daysFilter !== "all") {
        const daysVal = getDaysValForRow(row, cancelledByLot);
        if (typeof daysVal !== "number") return false;
        if (daysFilter === "red" && daysVal <= 5) return false;
        if (daysFilter === "green" && daysVal > 5) return false;
        if (daysFilter === "gt10" && daysVal <= 10) return false;
        if (daysFilter === "gt15" && daysVal <= 15) return false;
      }

      return true;
    });
  }, [
    challanData, searchTerm, financialYearFilter, printingStatusFilter, daysFilter, cancelledByLot,
    fabricFilter, brandFilter, styleFilter, garmentTypeFilter,
    sectionFilter, partyNameFilter, printingFilterDropdown, priorityFilter
  ]);

  // Calculate total PCs for filtered data
  const totalFilteredPCs = useMemo(() => {
    return filteredData.reduce((sum, row) => {
      const qty = parseFloat(row["Challan Total Qty"] || 0);
      return sum + (isNaN(qty) ? 0 : qty);
    }, 0);
  }, [filteredData]);

  // Calculate total PCs for all data
  const totalAllPCs = useMemo(() => {
    return challanData.reduce((sum, row) => {
      const qty = parseFloat(row["Challan Total Qty"] || 0);
      return sum + (isNaN(qty) ? 0 : qty);
    }, 0);
  }, [challanData]);

  // Build export rows + Days
  const exportedRows = useMemo(() => {
    const today = new Date(); today.setHours(0, 0, 0, 0);

    return filteredData.map((row, index) => {
      const lot = String(row["Lot Number"] || "").trim();
      const cancelledSet = cancelledByLot[lot];
      const status = computeStatus(row["Shade"], row["Challan History JSON"], cancelledSet);
      const printingStatus = getPrintingStatus(row["Challan History JSON"]);
      const pendingArr = getPendingShadesArray(row["Shade"], row["Challan History JSON"], cancelledSet);
      const pending = pendingArr.join(", ").toUpperCase();
      const remarks = getRemarks({
        lot,
        shadeString: row["Shade"],
        challanHistoryJson: row["Challan History JSON"],
        cancelledByLot,
        emptySizeByLot,
      });

      // Check if this is a repeated lot
      const isRepeated = isRepeatedLot(row["Priority"]);

      const receiveDateStr =
        status.type === "success" && printingStatus.type === "success"
          ? getLatestPrintingUpdatedAt(row["Challan History JSON"])
          : "";

      const receiveObj =
        status.type === "success" && printingStatus.type === "success"
          ? getLatestPrintingUpdatedAtObj(row["Challan History JSON"])
          : null;

      const designDateObj = receiveObj
        ? getDesignDateOnOrAfter(designDatesByLot, lot, receiveObj)
        : null;
      const designDateStr = formatDDMMMYYYY(designDateObj);

      const refDate = receiveObj || today;
      const days = daysBetween(row["Challan Date"], refDate);

      const base = {};
      // Add S. No column
      base["S. No"] = index + 1;

      // Process display headers with transformations
      ORDERED_DISPLAY_HEADERS.forEach((h) => {
        let value = row[h] ?? "";
        // Apply transformations for specific columns
        if (h === "Section") {
          value = mapSectionValue(value);
        } else if (h === "Party Name") {
          value = abbreviatePartyName(value);
        } else if (h === "Lot Number" && isRepeated) {
          // Add star indicator for repeated lots
          value = `★ ${value}`;
        }
        base[headerLabel(h)] = String(value || "");
      });

      // Store repeated flag for PDF styling
      base["_isRepeated"] = isRepeated;
      base["_priority"] = row["Priority"];

      // Add extra columns with proper defaults
      base["Status"] = String(status.text || "");
      base["Printing Status"] = String(printingStatus.text || "");
      base["Pending Challan Shade"] = String(pending || "");
      base["Remarks"] = String(remarks || "");
      base["Printing Done"] = String(receiveDateStr || "");
      base["Days"] = String(days || "");

      return base;
    });
  }, [filteredData, cancelledByLot, emptySizeByLot, designDatesByLot]);

  // Helper function to get filter value by header
  const getFilterValue = (header) => {
    switch (header) {
      case "Fabric": return fabricFilter;
      case "Brand": return brandFilter;
      case "Style": return styleFilter;
      case "Garment Type": return garmentTypeFilter;
      case "Section": return sectionFilter;
      case "Party Name": return partyNameFilter;
      case "Printing": return printingFilterDropdown;
      case "Priority": return priorityFilter;
      default: return [];
    }
  };

  // Helper function to set filter value
  const handleFilterChange = (header, value) => {
    switch (header) {
      case "Fabric": setFabricFilter(value); break;
      case "Brand": setBrandFilter(value); break;
      case "Style": setStyleFilter(value); break;
      case "Garment Type": setGarmentTypeFilter(value); break;
      case "Section": setSectionFilter(value); break;
      case "Party Name": setPartyNameFilter(value); break;
      case "Printing": setPrintingFilterDropdown(value); break;
      case "Priority": setPriorityFilter(value); break;
      default: break;
    }
  };

  // Function to extract unique values for dropdown filters with normalization
  const getUniqueValues = (data, header) => {
    const normalizedMap = new Map();
    data.forEach(row => {
      const value = String(row[header] || "").trim();
      if (!value) return;

      const normalized = value
        .toLowerCase()
        .replace(/\s+/g, '')
        .replace(/[^a-z0-9]/g, '');

      if (!normalizedMap.has(normalized)) {
        normalizedMap.set(normalized, value);
      }
    });

    return Array.from(normalizedMap.values()).sort();
  };

  const exportToExcel = async () => {
    if (!exportedRows.length) return;
    try {
      const XLSX = await ensureXLSX();
      const ws = XLSX.utils.json_to_sheet(exportedRows);
      const cols = Object.keys(exportedRows[0] || {});
      ws["!cols"] = cols.map((c) => {
        const max = Math.max(c.length, ...exportedRows.map((r) => String(r[c] ?? "").length));
        return { wch: Math.min(Math.max(max + 2, 10), 60) };
      });
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Challans");
      const wbout = XLSX.write(wb, { type: "array", bookType: "xlsx" });
      downloadBlob(
        new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
        `Printing-Challans-${new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-")}.xlsx`
      );
    } catch (e) {
      console.error(e);
      const cols = Object.keys(exportedRows[0] || {});
      const csv =
        [cols.join(",")]
          .concat(
            exportedRows.map((r) => cols.map((k) => `"${String(r[k] ?? "").replace(/"/g, '""')}"`).join(","))
          )
          .join("\n") + "\n";
      downloadBlob(new Blob([csv], { type: "text/csv;charset=utf-8" }), "Printing-Challans.csv");
      alert("Excel export fell back to CSV because the XLSX library couldn't be loaded.");
    }
  };

  const exportToPDF = async () => {
    if (!exportedRows.length) return;

    try {
      const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "A3" });

      // ---- Palette ----
      // ---- Palette (matching StitchingCompleted.js) ----
      const COLOR = {
        headerFill: [15, 76, 129],      // Deep Navy Blue
        headerText: [255, 255, 255],    // White bold text
        grid: [0, 0, 0],                // Solid Black grid lines
        stripe: [248, 250, 252],        // Light slate alternating stripe
        text: [15, 23, 42],             // Deep Charcoal text
        accent: [59, 130, 246],         // Vibrant Blue Accent
        daysGood: [220, 252, 231],      // Soft Green badge (#dcfce7)
        daysGoodText: [21, 128, 61],    // Green text (#15803d)
        daysBad: [239, 68, 68],         // Vibrant Solid Red (#ef4444) for Days > 5
        daysBadText: [255, 255, 255],   // Bold White Text (#ffffff)
        pageBorder: [0, 0, 0],          // Black page border
        summaryBg: [239, 246, 255],     // Light blue summary box (#eff6ff)
        summaryText: [30, 58, 138],     // Navy summary text
        totalBg: [239, 246, 255],       // Soft Navy/Blue total summary banner
        repeatedLotBg: [254, 243, 199], // Light amber highlight (#fef3c7)
        starColor: [217, 119, 6],       // Gold star
      };

      // ---- ENHANCED STAR DRAWING FUNCTION ----
      const drawStarSymbol = (doc, x, y, size = 4.5) => {
        const currentFillColor = doc.getFillColor();
        const currentDrawColor = doc.getDrawColor();
        const currentLineWidth = doc.internal.getLineWidth();

        doc.setFillColor(255, 215, 0);
        doc.setDrawColor(184, 134, 11);
        doc.setLineWidth(0.4);
        doc.circle(x, y, size, 'FD');

        doc.setDrawColor(160, 120, 10);
        doc.setLineWidth(0.5);

        const points = 5;
        for (let i = 0; i < points; i++) {
          const angle = (i * 2 * Math.PI / points) - Math.PI / 2;
          const endX = x + (size * 1.4) * Math.cos(angle);
          const endY = y + (size * 1.4) * Math.sin(angle);
          doc.line(x, y, endX, endY);

          const nextAngle = angle + (Math.PI / points);
          const innerX = x + (size * 0.7) * Math.cos(nextAngle);
          const innerY = y + (size * 0.7) * Math.sin(nextAngle);
          doc.line(endX, endY, innerX, innerY);
        }

        doc.setFillColor(255, 255, 180);
        doc.circle(x, y, size * 0.5, 'F');

        doc.setDrawColor(140, 100, 0);
        doc.setLineWidth(0.3);
        doc.circle(x, y, size * 1.1, 'D');

        doc.setFillColor(currentFillColor);
        doc.setDrawColor(currentDrawColor);
        doc.setLineWidth(currentLineWidth);
      };

      // ---- Calculate totals ----
      const totalLots = exportedRows.length;
      const totalPcs = exportedRows.reduce((sum, row) => {
        const qty = parseFloat(row["Challan Total Qty"]) || 0;
        return sum + qty;
      }, 0);
      const repeatedLotsCount = exportedRows.filter(row => row["_isRepeated"]).length;

      // ---- Calculate Printing party statistics ----
      const printingParties = {};
      exportedRows.forEach(row => {
        const printingParty = row["Printing"] || "Unassigned";
        const qty = parseFloat(row["Challan Total Qty"]) || 0;

        if (!printingParties[printingParty]) {
          printingParties[printingParty] = {
            name: printingParty,
            totalLots: 0,
            totalPcs: 0
          };
        }

        printingParties[printingParty].totalLots += 1;
        printingParties[printingParty].totalPcs += qty;
      });

      const sortedPrintingParties = Object.values(printingParties).sort((a, b) => b.totalPcs - a.totalPcs);

      // ---- Title block ----
      const title = "PRINTING CHALLAN PRODUCTION REPORT";
      const now = new Date();
      const subtitle = `Report Date: ${now.toLocaleDateString()} ${now.toLocaleTimeString()} • Factory Suite Pro`;

      const pageW = doc.internal.pageSize.getWidth();
      const pageH = doc.internal.pageSize.getHeight();

      doc.setFontSize(20);
      doc.setTextColor(15, 76, 129);
      doc.setFont('helvetica', 'bold');
      doc.text(title, pageW / 2, 38, { align: 'center' });

      doc.setFontSize(10);
      doc.setTextColor(71, 85, 105);
      doc.setFont('helvetica', 'normal');
      doc.text(subtitle, pageW / 2, 54, { align: 'center' });

      // ---- Main Summary section ----
      let currentY = 66;

      doc.setFillColor(...COLOR.totalBg);
      doc.roundedRect(30, currentY - 5, pageW - 60, 24, 6, 6, 'F');
      doc.setDrawColor(191, 219, 254);
      doc.roundedRect(30, currentY - 5, pageW - 60, 24, 6, 6, 'D');

      doc.setFontSize(11);
      doc.setTextColor(30, 58, 138);
      doc.setFont('helvetica', 'bold');

      const summaryText = `Total Records: ${totalLots}   |   Total Pcs / Qty: ${totalPcs.toLocaleString()}   |   Repeated Lots: ${repeatedLotsCount}`;
      const summaryTextWidth = doc.getTextWidth(summaryText);
      const summaryX = (pageW - summaryTextWidth) / 2;

      doc.text(summaryText, summaryX, currentY + 11);
      doc.setFont('helvetica', 'normal');

      currentY += 26;

      // ---- Define columns ----
      const columns = [
        "S. No",
        "Lot No.",
        "Fabric",
        "Brand",
        "Style",
        "Section",
        "Garment Type",
        "Party Name",
        "Printing",
        "Challan Date",
        "Challan Total Qty",
        "Printing Status",
        "Pending Challan Shade",
        "Remarks",
        "Printing Done",
        "Days",
        "HOD Remarks"
      ];

      // Build body data
      const body = exportedRows.map((row, index) => {
        const lotNum = row["Lot No."] ? String(row["Lot No."]).replace(/★\s*/, '').trim() : "";
        const lotHist = headRemarksMap[lotNum] || [];
        const latestHod = lotHist.length > 0 ? lotHist[lotHist.length - 1].text : "";

        return [
          String(index + 1),
          lotNum,
          row["Fabric"] || "",
          row["Brand"] || "",
          row["Style"] || "",
          row["Section"] || "",
          row["Garment Type"] || "",
          row["Party Name"] || "",
          row["Printing"] || "",
          row["Challan Date"] || "",
          row["Challan Total Qty"] || "",
          row["Printing Status"] || "",
          row["Pending Challan Shade"] || "",
          row["Remarks"] || "",
          row["Printing Done"] || "",
          row["Days"] || "",
          latestHod || ""
        ];
      });

      const daysColIdx = columns.indexOf("Days");
      const lotColIdx = columns.indexOf("Lot No.");

      const availableWidth = pageW - 30;

      // Proportional base widths for 17 columns (sum exactly to availableWidth)
      const baseWidths = [
        28,  // 0: S. No
        55,  // 1: Lot No.
        60,  // 2: Fabric
        65,  // 3: Brand
        70,  // 4: Style
        30,  // 5: Section
        65,  // 6: Garment Type
        40,  // 7: Party Name
        42,  // 8: Printing
        55,  // 9: Challan Date
        45,  // 10: Challan Total Qty
        58,  // 11: Printing Status
        70,  // 12: Pending Challan Shade
        60,  // 13: Remarks
        55,  // 14: Printing Done
        35,  // 15: Days
        82   // 16: HOD Remarks
      ];

      const baseSum = baseWidths.reduce((a, b) => a + b, 0);
      const columnStyles = {};
      baseWidths.forEach((w, i) => {
        columnStyles[i] = {
          cellWidth: (w / baseSum) * availableWidth,
          halign: 'center'
        };
      });

      const starPositions = [];

      // ---- Create main table - USING autoTable(doc, {...}) like Embroidery component ----
      autoTable(doc, {
        head: [columns],
        body,
        startY: currentY,
        margin: { left: 15, right: 15 },
        tableWidth: availableWidth,
        theme: "grid",
        styles: {
          fontSize: 9,
          cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
          overflow: "linebreak",
          valign: 'middle',
          textColor: COLOR.text,
          lineColor: [0, 0, 0],
          lineWidth: 0.5,
          fontStyle: 'normal',
          minCellHeight: 12,
        },
        headStyles: {
          fillColor: COLOR.headerFill,
          textColor: COLOR.headerText,
          fontStyle: "bold",
          lineColor: [0, 0, 0],
          lineWidth: 0.5,
          halign: 'center',
          fontSize: 9.5,
          valign: 'middle',
          cellPadding: { top: 5, right: 3, bottom: 5, left: 3 },
        },
        bodyStyles: {
          halign: 'center',
          valign: 'middle',
        },
        alternateRowStyles: {
          fillColor: COLOR.stripe,
        },
        columnStyles,

        didParseCell: function (data) {
          if (data.section === "body") {
            const rowIndex = data.row.index;
            const row = exportedRows[rowIndex];
            const isRepeated = row && row["_isRepeated"];

            if (isRepeated && data.column.index === lotColIdx) {
              data.cell.styles.fillColor = COLOR.repeatedLotBg;
              data.cell.styles.fontStyle = "bold";
              data.cell.styles.fontSize = 10;
            }

            if (data.column.index === daysColIdx) {
              const rawVal = data.cell.raw !== undefined && data.cell.raw !== null ? String(data.cell.raw).trim() : "";
              const n = parseFloat(rawVal);
              if (!isNaN(n) && rawVal !== "") {
                if (n > 5) {
                  data.cell.styles.fillColor = [239, 68, 68];   // Solid Vibrant Red (#ef4444)
                  data.cell.styles.textColor = [255, 255, 255]; // Pure White Bold Text (#ffffff)
                  data.cell.styles.fontStyle = "bold";
                } else {
                  data.cell.styles.fillColor = [220, 252, 231]; // Soft Green badge (#dcfce7)
                  data.cell.styles.textColor = [21, 128, 61];   // Dark Green text (#15803d)
                  data.cell.styles.fontStyle = "bold";
                }
              }
            }
          }
        },

        willDrawCell: function (data) {
          if (data.section === "body" && data.column.index === lotColIdx) {
            const rowIndex = data.row.index;
            const row = exportedRows[rowIndex];
            const isRepeated = row && row["_isRepeated"];

            if (isRepeated) {
              const lotNumber = row["Lot No."] ? String(row["Lot No."]).replace(/★\s*/, '').trim() : "";

              starPositions.push({
                lotNumber: lotNumber,
                rowIndex: rowIndex,
                cell: data.cell
              });

              data.cell.text = [];
            }
          }
        },

        didDrawCell: function (data) {
          if (data.section === "body" && data.column.index === lotColIdx) {
            const rowIndex = data.row.index;
            const starData = starPositions.find(s => s.rowIndex === rowIndex);

            if (starData) {
              try {
                const { lotNumber } = starData;

                const cellX = data.cell.x;
                const cellY = data.cell.y;
                const cellWidth = data.cell.width;
                const cellHeight = data.cell.height;

                const centerX = cellX + (cellWidth / 2);
                const centerY = cellY + (cellHeight / 2);

                const originalFont = doc.internal.getFont();
                const originalSize = doc.internal.getFontSize();
                const originalColor = doc.getTextColor();

                doc.setFontSize(10);
                doc.setFont(undefined, 'bold');

                const textWidth = doc.getTextWidth(lotNumber);
                const iconSize = 4.5;

                const starX = centerX - (textWidth / 2) - iconSize - 5;
                const starY = centerY;

                drawStarSymbol(doc, starX, starY, iconSize);

                const textX = centerX - (textWidth / 2);
                const textY = centerY + 5;

                doc.setTextColor(0, 0, 0);
                doc.text(lotNumber, textX, textY);

                doc.setFont(originalFont.fontName, originalFont.fontStyle);
                doc.setFontSize(originalSize);
                doc.setTextColor(originalColor[0], originalColor[1], originalColor[2]);

                const index = starPositions.findIndex(s => s.rowIndex === rowIndex);
                if (index !== -1) {
                  starPositions.splice(index, 1);
                }

              } catch (e) {
                console.warn("Error drawing star for row", rowIndex, e);
              }
            }
          }
        },

        didDrawPage: (hookData) => {
          const pw = doc.internal.pageSize.getWidth();
          const ph = doc.internal.pageSize.getHeight();

          doc.setDrawColor(...COLOR.pageBorder);
          doc.setLineWidth(0.5);
          doc.roundedRect(10, 10, pw - 20, ph - 20, 2, 2, "S");

          doc.setFontSize(8.5);
          doc.setTextColor(100, 116, 139);
          doc.setFont('helvetica', 'normal');

          doc.text(`Factory Suite Pro • Printing Challan Report`, 18, ph - 14);

          const totalPages = doc.internal.getNumberOfPages();
          const label = `Page ${hookData.pageNumber} of ${totalPages}`;
          doc.text(label, pw / 2, ph - 14, { align: 'center' });

          const timeStr = `${new Date().toLocaleDateString()} ${new Date().toLocaleTimeString()}`;
          doc.text(`Generated: ${timeStr}`, pw - 18, ph - 14, { align: 'right' });
        },

        margin: { top: currentY, right: 15, bottom: 25, left: 15 },
      });

      // ---- FALLBACK: Draw any remaining stars ----
      if (starPositions.length > 0) {
        starPositions.forEach(({ lotNumber, rowIndex, cell }) => {
          try {
            const cellX = cell.x;
            const cellY = cell.y;
            const cellWidth = cell.width;
            const cellHeight = cell.height;

            const centerX = cellX + (cellWidth / 2);
            const centerY = cellY + (cellHeight / 2);

            doc.setFontSize(10);
            doc.setFont(undefined, 'bold');

            const textWidth = doc.getTextWidth(lotNumber);
            const iconSize = 4.5;

            const starX = centerX - (textWidth / 2) - iconSize - 5;
            const starY = centerY;

            drawStarSymbol(doc, starX, starY, iconSize);

            const textX = centerX - (textWidth / 2);
            const textY = centerY + 5;

            doc.setTextColor(0, 0, 0);
            doc.text(lotNumber, textX, textY);

            doc.setFont(undefined, 'normal');
          } catch (e) {
            console.warn("Error drawing star for row", rowIndex, e);
          }
        });
      }

      // ---- PRINTING PARTY SUMMARY ----

      let lastAutoTable = doc.lastAutoTable;
      let yPosition = lastAutoTable.finalY + 40;

      if (yPosition > pageH - 100) {
        doc.addPage();
        yPosition = 60;
      } else {
        yPosition += 10;
        doc.setDrawColor(...COLOR.headerFill);
        doc.setLineWidth(1);
        doc.line(50, yPosition - 5, pageW - 50, yPosition - 5);
      }

      if (sortedPrintingParties.length > 0) {

        doc.setFontSize(22);
        doc.setFont(undefined, 'bold');
        doc.setTextColor(...COLOR.headerFill);

        const summaryTitle = "PRINTING PARTY SUMMARY REPORT";
        const titleWidth = doc.getTextWidth(summaryTitle);
        const titleX = (pageW - titleWidth) / 2;
        doc.text(summaryTitle, titleX, yPosition);

        yPosition += 25;

        doc.setFontSize(14);
        doc.setFont(undefined, 'normal');
        doc.setTextColor(...COLOR.text);

        const totalText = `Total Printing Parties: ${sortedPrintingParties.length}  |  Total Pieces: ${totalPcs.toLocaleString()}  |  Total Lots: ${totalLots}`;
        const totalTextWidth = doc.getTextWidth(totalText);
        const totalTextX = (pageW - totalTextWidth) / 2;
        doc.text(totalText, totalTextX, yPosition);

        yPosition += 25;

        const tableColumnWidths = [200, 120, 140, 120];
        const totalTableWidth = tableColumnWidths.reduce((a, b) => a + b, 0);
        const tableMarginLeft = (pageW - totalTableWidth) / 2;

        autoTable(doc, {
          head: [['Printing Party', 'Total Lots', 'Total Pcs', 'Share %']],
          body: sortedPrintingParties.map((party) => {
            const percentage = (party.totalPcs / totalPcs * 100).toFixed(1);
            return [
              party.name,
              party.totalLots.toString(),
              party.totalPcs.toLocaleString(),
              `${percentage}%`
            ];
          }),
          startY: yPosition,
          theme: "grid",
          styles: {
            fontSize: 11,
            cellPadding: 8,
            overflow: "linebreak",
            valign: 'middle',
            halign: 'center',
            textColor: COLOR.text,
            lineColor: COLOR.grid,
            lineWidth: 0.2,
          },
          headStyles: {
            fillColor: [79, 70, 229],
            textColor: [255, 255, 255],
            fontStyle: "bold",
            fontSize: 12,
            halign: 'center',
            cellPadding: 10,
          },
          columnStyles: {
            0: { cellWidth: 200, halign: 'left' },
            1: { cellWidth: 120, halign: 'center' },
            2: { cellWidth: 140, halign: 'right' },
            3: { cellWidth: 120, halign: 'center' },
          },
          margin: { left: tableMarginLeft, right: tableMarginLeft },
          didDrawPage: (hookData) => {
            if (hookData.pageNumber > doc.internal.getNumberOfPages() - 1 || hookData.pageNumber === doc.internal.getNumberOfPages()) {
              const pw = doc.internal.pageSize.getWidth();
              const ph = doc.internal.pageSize.getHeight();

              doc.setDrawColor(...COLOR.pageBorder);
              doc.setLineWidth(0.5);
              doc.roundedRect(10, 10, pw - 20, ph - 20, 2, 2, "S");

              doc.setFontSize(9);
              doc.setTextColor(...COLOR.accent);
              const totalPages = doc.internal.getNumberOfPages();
              const label = `Page ${hookData.pageNumber} of ${totalPages}`;
              const textW = doc.getTextWidth(label);
              doc.text(label, pw - 15 - textW, ph - 15);
              doc.text("Printing Party Summary", 15, ph - 15);
            }
          }
        });

        const tableEndY = doc.lastAutoTable.finalY + 30;

        const finalY = tableEndY + 20;
        doc.setDrawColor(...COLOR.headerFill);
        doc.setLineWidth(0.5);
        doc.line(50, finalY, pageW - 50, finalY);

        doc.setFontSize(11);
        doc.setFont(undefined, 'italic');
        doc.setTextColor(100, 100, 100);
        const thankYouText = "Report Generated Successfully";
        const thankYouWidth = doc.getTextWidth(thankYouText);
        const thankYouX = (pageW - thankYouWidth) / 2;
        doc.text(thankYouText, thankYouX, finalY + 20);
      }

      const ts = new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
      doc.save(`Printing-Challans-${ts}.pdf`);

    } catch (e) {
      console.error("PDF export error:", e);
      alert(`PDF export failed: ${e.message}`);
    }
  };
  const goBack = () => {
    try {
      if (window.history.length > 1) return window.history.back();
    } catch { }
    window.location.href = "/";
  };

  if (loading) {
    return (
      <div style={styles.loadingContainer}>
        <div style={styles.spinner}></div>
        <p style={styles.loadingText}>🧵 Loading Printing Challans...</p>
      </div>
    );
  }

  if (error) {
    return (
      <div style={styles.errorContainer}>
        <div style={styles.errorIcon}>⚠️</div>
        <h3 style={styles.errorTitle}>Connection Error</h3>
        <p style={styles.errorMessage}>{error}</p>
        <button style={styles.retryButton} onClick={() => window.location.reload()}>
          🔄 Retry
        </button>
      </div>
    );
  }

  // Helpers for day coloring in UI
  const dayBgStyle = (daysVal) => {
    if (typeof daysVal !== "number") return styles.dayNeutral;
    return daysVal <= 5 ? styles.dayOk : styles.dayLate;
  };

  return (
    <div style={styles.container}>
      {/* Image Lightbox Modal */}
      {viewImageSrc && (
        <div style={styles.imageModalBackdrop} onClick={() => setViewImageSrc(null)}>
          <div style={styles.imageModalContent} onClick={(e) => e.stopPropagation()}>
            <button style={styles.imageModalClose} onClick={() => setViewImageSrc(null)}>&times;</button>
            <img src={viewImageSrc} alt="Full Preview" style={styles.imageModalImg} />
          </div>
        </div>
      )}

      {/* Header */}
      {!isEmbedded && (
        <div style={styles.header}>
          <div style={styles.headerContentBox}>
            <div style={styles.titleSection}>
              <h1 style={styles.title}>🧵 Printing Challan Records</h1>
              <p style={styles.subtitle}>Manage and track printing challan progress</p>
            </div>
            <div style={styles.statsCard}>
              <div style={styles.statItem}>
                <span style={styles.statNumber}>{challanData.length}</span>
                <span style={styles.statLabel}>Total Challans</span>
              </div>
              <div style={styles.statItem}>
                <span style={styles.statNumber}>
                  {
                    challanData.filter((row) => {
                      const lot = String(row["Lot Number"] || "").trim();
                      const cancelledSet = cancelledByLot[lot];
                      return (
                        computeStatus(row["Shade"], row["Challan History JSON"], cancelledSet).type === "success"
                      );
                    }).length
                  }
                </span>
                <span style={styles.statLabel}>Completed Lots</span>
              </div>
              <div style={styles.statItem}>
                <span style={styles.statNumber}>
                  {
                    challanData.filter((row) => isRepeatedLot(row["Priority"])).length
                  }
                </span>
                <span style={styles.statLabel}>Repeated Lots</span>
              </div>
              <div
                style={{ ...styles.statItem, cursor: 'pointer' }}
                onClick={() => setDaysFilter(daysFilter === "red" ? "all" : "red")}
                title="Click to toggle Red Zone filter (>5 Days)"
              >
                <span style={{ ...styles.statNumber, color: '#f87171' }}>{redZoneCount}</span>
                <span style={styles.statLabel}>🔴 Red Zone (&gt;5 Days)</span>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Controls */}
      <div style={styles.controlsContainer}>
        {/* Top Row: Search and Main Actions */}
        <div style={styles.controlsTopRow}>
          <div style={styles.searchContainer}>
            <div style={styles.searchBox}>
              <span style={styles.searchIcon}>🔍</span>
              <input
                type="text"
                placeholder="Search challans..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                style={styles.searchInput}
                onKeyDown={(e) => e.key === 'Escape' && setSearchTerm('')}
              />
              {searchTerm && (
                <button
                  onClick={() => setSearchTerm("")}
                  style={styles.clearSearchBtn}
                  title="Clear search"
                >
                  ✕
                </button>
              )}
            </div>
            <div style={styles.searchHint}>
              <small>Press Esc to clear search</small>
            </div>
          </div>

          <div style={styles.mainActions}>
            <button
              style={loading ? styles.refreshBtnLoading : styles.refreshBtn}
              onClick={loadAllData}
              disabled={loading}
              title="Refresh data from Google Sheets"
            >
              {loading ? "⏳ Refreshing..." : "⟳ Refresh"}
            </button>
            <button style={styles.exportBtnExcel} onClick={exportToExcel} title="Export to Excel">
              📊 Excel
            </button>
            <button style={styles.exportBtnPdf} onClick={exportToPDF} title="Export to PDF">
              📄 PDF
            </button>
            <button style={styles.backBtn} onClick={goBack} title="Go back">
              ← Back
            </button>
          </div>
        </div>

        {/* Filter Section */}
        <div style={styles.filtersSection}>
          <div style={styles.filtersHeader}>
            <h3 style={styles.filtersTitle}>Filters</h3>
            <div style={styles.activeFilterInfo}>
              {[
                fabricFilter.length > 0,
                brandFilter.length > 0,
                styleFilter.length > 0,
                garmentTypeFilter.length > 0,
                sectionFilter.length > 0,
                partyNameFilter.length > 0,
                printingFilterDropdown.length > 0,
                priorityFilter.length > 0,
                printingStatusFilter !== "all",
                daysFilter !== "all"
              ].filter(Boolean).length > 0 && (
                  <>
                    <span style={styles.activeFilterCount}>
                      {[
                        fabricFilter.length > 0,
                        brandFilter.length > 0,
                        styleFilter.length > 0,
                        garmentTypeFilter.length > 0,
                        sectionFilter.length > 0,
                        partyNameFilter.length > 0,
                        printingFilterDropdown.length > 0,
                        priorityFilter.length > 0,
                        printingStatusFilter !== "all" ? "Printing Status" : null,
                        daysFilter !== "all" ? "Days" : null
                      ].filter(Boolean).length} filter(s) active
                    </span>
                    <button
                      onClick={() => {
                        setFabricFilter([]);
                        setBrandFilter([]);
                        setStyleFilter([]);
                        setGarmentTypeFilter([]);
                        setSectionFilter([]);
                        setPartyNameFilter([]);
                        setPrintingFilterDropdown([]);
                        setPriorityFilter([]);
                        setPrintingStatusFilter("all");
                        setDaysFilter("all");
                      }}
                      style={styles.clearFiltersBtn}
                    >
                      Clear All
                    </button>
                  </>
                )}
            </div>
          </div>

          <div style={styles.filtersGrid}>
            {FILTER_HEADERS.map(header => (
              <div key={header} style={styles.filterCard}>
                <div style={styles.filterHeader}>
                  <label htmlFor={`${header}Filter`} style={styles.filterLabel}>
                    {headerLabel(header)}
                  </label>
                  {getFilterValue(header).length > 0 && (
                    <button
                      onClick={() => handleFilterChange(header, [])}
                      style={styles.clearSingleFilter}
                      title={`Clear ${headerLabel(header)} filter`}
                    >
                      ✕
                    </button>
                  )}
                </div>
                <MultiSelectDropdown
                  placeholder={`All ${headerLabel(header)}`}
                  options={getUniqueValues(challanData, header)}
                  selectedValues={getFilterValue(header)}
                  onChange={(val) => handleFilterChange(header, val)}
                />
              </div>
            ))}

            {/* Financial Year Filter */}
            <div style={styles.filterCard}>
              <div style={styles.filterHeader}>
                <label htmlFor="financialYearFilter" style={styles.filterLabel}>📅 Financial Year</label>
                {financialYearFilter !== "ALL" && (
                  <button
                    onClick={() => setFinancialYearFilter("ALL")}
                    style={styles.clearSingleFilter}
                    title="Show All Data (Whole Data)"
                  >
                    ✕
                  </button>
                )}
              </div>
              <select
                id="financialYearFilter"
                value={financialYearFilter}
                onChange={(e) => setFinancialYearFilter(e.target.value)}
                style={{
                  ...styles.filterSelect,
                  ...(financialYearFilter !== "ALL" && styles.filterSelectActive)
                }}
              >
                <option value="ALL">🌐 All Data (Whole Data)</option>
                {availableFYs.map((fy) => (
                  <option key={fy} value={fy}>
                    FY {fy} {fy === getCurrentFinancialYear() ? "(Current)" : ""}
                  </option>
                ))}
              </select>
            </div>

            {/* Printing Status Filter */}
            <div style={styles.filterCard}>
              <div style={styles.filterHeader}>
                <label htmlFor="printingStatusFilter" style={styles.filterLabel}>Printing Status</label>
                {printingStatusFilter !== "all" && (
                  <button
                    onClick={() => setPrintingStatusFilter("all")}
                    style={styles.clearSingleFilter}
                    title="Clear Printing Status filter"
                  >
                    ✕
                  </button>
                )}
              </div>
              <select
                id="printingStatusFilter"
                value={printingStatusFilter}
                onChange={(e) => setPrintingStatusFilter(e.target.value)}
                style={{
                  ...styles.filterSelect,
                  ...(printingStatusFilter !== "all" && styles.filterSelectActive)
                }}
              >
                <option value="all">All Status</option>
                <option value="done">✅ Printing Done</option>
                <option value="pending">⏳ Printing Pending</option>
                <option value="unknown">❓ Unknown</option>
              </select>
            </div>

            {/* Repeated Lots Filter */}
            <div style={styles.filterCard}>
              <div style={styles.filterHeader}>
                <label htmlFor="repeatedFilter" style={styles.filterLabel}>Repeated Lots</label>
              </div>
              <div style={styles.filterButtons}>
                <button
                  onClick={() => {
                    setPriorityFilter("REPEATED_LOT");
                  }}
                  style={{
                    ...styles.filterButton,
                    ...(priorityFilter === "REPEATED_LOT" && styles.filterButtonActive)
                  }}
                >
                  ★ Show Repeated
                </button>
                <button
                  onClick={() => {
                    setPriorityFilter("");
                  }}
                  style={styles.filterButton}
                >
                  Show All
                </button>
              </div>
            </div>

            {/* Days / Red Zone Filter */}
            <div style={styles.filterCard}>
              <div style={styles.filterHeader}>
                <label htmlFor="daysFilter" style={styles.filterLabel}>⏱️ Days / Zone Filter</label>
                {daysFilter !== "all" && (
                  <button
                    onClick={() => setDaysFilter("all")}
                    style={styles.clearSingleFilter}
                    title="Clear Days filter"
                  >
                    ✕
                  </button>
                )}
              </div>
              <select
                id="daysFilter"
                value={daysFilter}
                onChange={(e) => setDaysFilter(e.target.value)}
                style={{
                  ...styles.filterSelect,
                  ...(daysFilter !== "all" && styles.filterSelectActive)
                }}
              >
                <option value="all">🌐 All Days</option>
                <option value="red">🔴 Red Zone (&gt; 5 Days)</option>
                <option value="green">🟢 Green Zone (≤ 5 Days)</option>
                <option value="gt10">🔴 &gt; 10 Days</option>
                <option value="gt15">🔴 &gt; 15 Days</option>
              </select>
              <div style={{ display: 'flex', gap: '6px', marginTop: '8px' }}>
                <button
                  onClick={() => setDaysFilter(daysFilter === "red" ? "all" : "red")}
                  style={{
                    ...styles.filterButton,
                    flex: 1,
                    fontSize: '0.78rem',
                    padding: '6px 8px',
                    backgroundColor: daysFilter === "red" ? '#fee2e2' : '#f8fafc',
                    color: daysFilter === "red" ? '#991b1b' : '#475569',
                    borderColor: daysFilter === "red" ? '#ef4444' : '#cbd5e1',
                    fontWeight: daysFilter === "red" ? '700' : '500',
                  }}
                >
                  🔴 Red Zone ({redZoneCount})
                </button>
                <button
                  onClick={() => setDaysFilter(daysFilter === "green" ? "all" : "green")}
                  style={{
                    ...styles.filterButton,
                    flex: 1,
                    fontSize: '0.78rem',
                    padding: '6px 8px',
                    backgroundColor: daysFilter === "green" ? '#dcfce7' : '#f8fafc',
                    color: daysFilter === "green" ? '#166534' : '#475569',
                    borderColor: daysFilter === "green" ? '#22c55e' : '#cbd5e1',
                    fontWeight: daysFilter === "green" ? '700' : '500',
                  }}
                >
                  🟢 Green Zone
                </button>
              </div>
            </div>
          </div>
        </div>

        {/* Results Info */}
        <div style={styles.resultsSummary}>
          <div style={styles.resultsInfoCard}>
            <div style={styles.resultStat}>
              <span style={styles.resultNumber}>{filteredData.length}</span>
              <span style={styles.resultLabel}>Showing</span>
            </div>
            <div style={styles.resultStat}>
              <span style={styles.resultNumber}>{challanData.length}</span>
              <span style={styles.resultLabel}>Total</span>
            </div>
            <div style={styles.resultStat}>
              <span style={styles.resultNumber}>
                {filteredData.filter(row => isRepeatedLot(row["Priority"])).length}
              </span>
              <span style={styles.resultLabel}>Repeated</span>
            </div>
            <div style={styles.resultStat}>
              <span style={{ ...styles.resultNumber, color: '#ef4444' }}>
                {filteredData.filter(row => {
                  const daysVal = getDaysValForRow(row, cancelledByLot);
                  return typeof daysVal === "number" && daysVal > 5;
                }).length}
              </span>
              <span style={styles.resultLabel}>Red Zone (&gt;5d)</span>
            </div>
            {/* <div style={styles.resultStat}>
              <span style={styles.resultNumber}>{totalFilteredPCs.toLocaleString()}</span>
              <span style={styles.resultLabel}>Filtered PCs</span>
            </div> */}
          </div>
          <div style={styles.resultsText}>
            {filteredData.length === 0 ? (
              <span style={styles.noResults}>No records found. Try adjusting your filters.</span>
            ) : (
              <span>
                Showing <strong>{filteredData.length}</strong> of <strong>{challanData.length}</strong> records
                {searchTerm && ` matching "${searchTerm}"`}
                {filteredData.filter(row => isRepeatedLot(row["Priority"])).length > 0 && (
                  <span style={{ color: '#d97706', marginLeft: '10px' }}>
                    ★ {filteredData.filter(row => isRepeatedLot(row["Priority"])).length} repeated lots
                  </span>
                )}
              </span>
            )}
          </div>
        </div>
      </div>

      {/* Table - Hide Priority column */}
      <div style={styles.tableContainer}>
        <table style={styles.table}>
          <thead>
            <tr>
              <th style={styles.tableHeader}>S. No</th>
              <th style={styles.tableHeader}>Image</th>
              {ORDERED_DISPLAY_HEADERS
                .filter(header => header !== "Priority" && !(isEmbedded && (header === "Section" || header === "Challan No" || header === "Party Name")))
                .map((header) => (
                  <th key={header} style={styles.tableHeader}>{headerLabel(header)}</th>
                ))}
              {!isEmbedded && <th style={styles.tableHeader}>Status</th>}
              {!isEmbedded && <th style={styles.tableHeader}>Printing Status</th>}
              <th style={styles.tableHeader}>Pending Challan Shade</th>
              <th style={styles.tableHeader}>Remarks</th>
              <th style={styles.tableHeader}>Printing Done</th>
              <th style={styles.tableHeader}>Days</th>
              <th style={{ ...styles.tableHeader, background: '#312e81' }}>HOD Remarks</th>
              {isEmbedded && (
                <>
                  <th style={{ ...styles.tableHeader, background: '#312e81' }}>Remarks History</th>
                  <th style={{ ...styles.tableHeader, background: '#312e81' }}>Add New Remarks</th>
                </>
              )}
            </tr>
          </thead>
          <tbody>
            {filteredData.map((row, idx) => {
              const lot = String(row["Lot Number"] || "").trim();
              const cancelledSet = cancelledByLot[lot];
              const status = computeStatus(row["Shade"], row["Challan History JSON"], cancelledSet);
              const printingStatus = getPrintingStatus(row["Challan History JSON"]);
              const pendingArr = getPendingShadesArray(row["Shade"], row["Challan History JSON"], cancelledSet);
              const pending = pendingArr.join(", ").toUpperCase();
              const remarks = getRemarks({
                lot,
                shadeString: row["Shade"],
                challanHistoryJson: row["Challan History JSON"],
                cancelledByLot,
                emptySizeByLot,
              });

              // Check if this is a repeated lot
              const isRepeated = isRepeatedLot(row["Priority"]);

              const receiveDateStr =
                status.type === "success" && printingStatus.type === "success"
                  ? getLatestPrintingUpdatedAt(row["Challan History JSON"])
                  : "";

              const receiveObj =
                status.type === "success" && printingStatus.type === "success"
                  ? getLatestPrintingUpdatedAtObj(row["Challan History JSON"])
                  : null;

              const designDateObj = receiveObj
                ? getDesignDateOnOrAfter(designDatesByLot, lot, receiveObj)
                : null;
              const designDateStr = formatDDMMMYYYY(designDateObj);

              const daysVal = getDaysValForRow(row, cancelledByLot);

              const lotRemarksHistory = headRemarksMap[lot] || [];

              return (
                <tr key={idx} style={{
                  ...(idx % 2 === 0 ? styles.tableRowEven : styles.tableRowOdd),
                  ...(isRepeated ? styles.repeatedRow : {})
                }}>
                  <td style={styles.tableCell}>{idx + 1}</td>
                  <td style={{ ...styles.tableCell, textAlign: 'center', verticalAlign: 'middle' }}>
                    {lotImages[lot] ? (
                      <img
                        src={lotImages[lot]}
                        alt="Style Preview"
                        style={styles.tableImage}
                        referrerPolicy="no-referrer"
                        onClick={(e) => {
                          e.stopPropagation();
                          setViewImageSrc(lotImages[lot]);
                        }}
                      />
                    ) : (
                      <span style={styles.noImagePlaceholder}>No Image</span>
                    )}
                  </td>
                  {ORDERED_DISPLAY_HEADERS
                    .filter(header => header !== "Priority" && !(isEmbedded && (header === "Section" || header === "Challan No" || header === "Party Name")))
                    .map((header) => {
                      let value = row[header];
                      // Apply transformations for specific columns
                      if (header === "Section") {
                        value = mapSectionValue(value);
                      } else if (header === "Party Name") {
                        value = abbreviatePartyName(value);
                      } else if (header === "Lot Number" && isRepeated) {
                        // Add golden star indicator for repeated lots
                        value = (
                          <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                            <span style={{ color: '#FFD700', fontSize: '16px', fontWeight: 'bold' }}>★</span>
                            <span>{value}</span>
                          </div>
                        );
                      }
                      return (
                        <td key={header} style={{
                          ...styles.tableCell,
                          ...(isRepeated && header === "Lot Number" ? styles.repeatedLotCell : {}),
                        }}>{value}</td>
                      );
                    })}
                  {!isEmbedded && <td style={styles.tableCell}>{status.text}</td>}
                  {!isEmbedded && <td style={styles.tableCell}>{printingStatus.text}</td>}
                  <td style={styles.tableCell}>{pending}</td>
                  <td style={styles.tableCell}>{remarks}</td>
                  <td style={styles.tableCell}>{receiveDateStr}</td>
                  <td style={{
                    ...styles.tableCell,
                    ...dayBgStyle(daysVal),
                    color: "#000",
                    fontWeight: 700,
                    ...(isRepeated ? styles.repeatedDaysCell : {})
                  }}>
                    {daysVal}
                  </td>
                  {/* HOD Remarks Column */}
                  <td style={{ ...styles.tableCell, minWidth: '170px', textAlign: 'center' }}>
                    {lotRemarksHistory.length > 0 ? (
                      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '4px' }}>
                        <div
                          style={{
                            background: '#eef2ff',
                            border: '1px solid #c7d2fe',
                            color: '#1e1b4b',
                            padding: '4px 8px',
                            borderRadius: '6px',
                            fontSize: '11px',
                            fontWeight: '700',
                            maxWidth: '220px',
                            wordBreak: 'break-word',
                            lineHeight: '1.3'
                          }}
                        >
                          {lotRemarksHistory[lotRemarksHistory.length - 1].text}
                        </div>
                        {lotRemarksHistory.length > 1 && (
                          <button
                            onClick={() => openRemarksModal(lot, lotRemarksHistory)}
                            style={{
                              border: 'none',
                              background: 'transparent',
                              color: '#4f46e5',
                              fontSize: '10px',
                              fontWeight: '700',
                              cursor: 'pointer',
                              textDecoration: 'underline'
                            }}
                          >
                            + {lotRemarksHistory.length - 1} earlier remark(s)
                          </button>
                        )}
                      </div>
                    ) : (
                      <span style={{ color: '#cbd5e1', fontSize: '11px' }}>-</span>
                    )}
                  </td>
                  {isEmbedded && (
                    <>
                      {/* Remarks History Column */}
                      <td style={{ ...styles.tableCell, textAlign: 'center', minWidth: '150px' }}>
                        {lotRemarksHistory.length === 0 ? (
                          <span style={{ color: '#94a3b8', fontSize: '11px', fontStyle: 'italic' }}>No remarks history</span>
                        ) : (
                          <button
                            onClick={() => openRemarksModal(lot, lotRemarksHistory)}
                            style={{
                              padding: '6px 14px',
                              borderRadius: '8px',
                              border: '1px solid #c7d2fe',
                              background: '#eef2ff',
                              color: '#3730a3',
                              fontSize: '11px',
                              fontWeight: '700',
                              cursor: 'pointer',
                              display: 'inline-flex',
                              alignItems: 'center',
                              gap: '6px',
                              boxShadow: '0 2px 4px rgba(99, 102, 241, 0.12)'
                            }}
                          >
                            👁️ View Remarks ({lotRemarksHistory.length})
                          </button>
                        )}
                      </td>

                      {/* Add New Remarks Column */}
                      <td style={{ ...styles.tableCell, minWidth: '220px' }}>
                        <div style={{ display: 'flex', gap: '6px', alignItems: 'center' }}>
                          <input
                            type="text"
                            placeholder="Enter new remark..."
                            value={newRemarkInput[lot] || ""}
                            onChange={(e) => {
                              const val = e.target.value;
                              setNewRemarkInput(prev => ({ ...prev, [lot]: val }));
                            }}
                            onKeyDown={(e) => {
                              if (e.key === "Enter") handleSaveHeadRemark(lot, row);
                            }}
                            style={{
                              padding: '6px 10px',
                              borderRadius: '8px',
                              border: '1px solid #cbd5e1',
                              fontSize: '12px',
                              width: '130px',
                              outline: 'none'
                            }}
                          />
                          <button
                            onClick={() => handleSaveHeadRemark(lot, row)}
                            style={{
                              padding: '6px 12px',
                              borderRadius: '8px',
                              border: 'none',
                              background: 'linear-gradient(135deg, #4f46e5 0%, #3730a3 100%)',
                              color: '#ffffff',
                              fontSize: '11px',
                              fontWeight: '700',
                              cursor: 'pointer',
                              boxShadow: '0 2px 6px rgba(79, 70, 229, 0.3)'
                            }}
                          >
                            + Save
                          </button>
                        </div>
                      </td>
                    </>
                  )}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <div style={styles.footer}>
        <p style={styles.footerText}>
          💫 Last updated: {new Date().toLocaleDateString()} • 🧵 Printing Management System • ★ = Repeated Lot •
          Total: {challanData.length} Lots, {totalAllPCs.toLocaleString()} PCs •
          Filtered: {filteredData.length} Lots, {totalFilteredPCs.toLocaleString()} PCs
        </p>
      </div>

      <RemarksHistoryModal
        isOpen={remarksModalOpen}
        onClose={() => setRemarksModalOpen(false)}
        lotNumber={selectedRemarksLot}
        remarksHistory={selectedRemarksHistory}
      />
    </div>
  );
}

/* ================== STYLES ================== */
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
  headerContentBox: { display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "20px" },
  titleSection: { flex: 1 },
  title: { margin: 0, fontSize: "2.4rem", fontWeight: "800", color: "#ffffff" },
  subtitle: { margin: "8px 0 0 0", fontSize: "1rem", color: "#c7d2fe", fontWeight: "400" },
  statsCard: { display: "flex", gap: "30px", background: "rgba(255,255,255,0.12)", border: "1px solid rgba(255,255,255,0.2)", padding: "16px 24px", borderRadius: "20px", backdropFilter: "blur(12px)", flexWrap: "wrap" },
  statItem: { textAlign: "center" },
  statNumber: { display: "block", fontSize: "2rem", fontWeight: "800", color: "#ffffff" },
  statLabel: { fontSize: "0.75rem", color: "#c7d2fe", fontWeight: "700", textTransform: "uppercase", letterSpacing: "0.05em" },

  // Controls Container
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
    marginBottom: "24px",
    flexWrap: "wrap",
    gap: "20px"
  },

  searchContainer: {
    flex: 1,
    minWidth: "300px"
  },

  searchBox: {
    position: "relative",
    width: "100%"
  },

  searchIcon: {
    position: "absolute",
    left: "16px",
    top: "50%",
    transform: "translateY(-50%)",
    fontSize: "1.2rem",
    color: "#94a3b8"
  },

  searchInput: {
    width: "100%",
    padding: "12px 20px 12px 48px",
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

  clearSearchBtn: {
    position: "absolute",
    right: "16px",
    top: "50%",
    transform: "translateY(-50%)",
    background: "#e2e8f0",
    border: "none",
    color: "#64748b",
    fontSize: "0.9rem",
    cursor: "pointer",
    padding: "2px",
    borderRadius: "50%",
    width: "22px",
    height: "22px",
    display: "flex",
    alignItems: "center",
    justifyContent: "center"
  },

  searchHint: {
    marginTop: "8px",
    color: "#64748b",
    fontSize: "0.85rem",
    paddingLeft: "4px"
  },

  mainActions: {
    display: "flex",
    gap: "12px",
    flexWrap: "wrap"
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

  refreshBtnLoading: {
    background: "linear-gradient(135deg, #6b7280 0%, #4b5563 100%)",
    color: "white",
    border: "none",
    padding: "10px 20px",
    borderRadius: "12px",
    fontSize: "0.95rem",
    fontWeight: "600",
    cursor: "not-allowed",
    opacity: 0.7
  },

  exportBtnExcel: {
    background: "linear-gradient(135deg, #10b981 0%, #059669 100%)",
    color: "white",
    border: "none",
    padding: "10px 20px",
    borderRadius: "12px",
    fontSize: "0.95rem",
    fontWeight: "600",
    cursor: "pointer",
    transition: "all 0.2s ease",
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

  filtersSection: {
    background: "#ffffff",
    borderRadius: "20px",
    padding: "24px",
    marginBottom: "28px",
    boxShadow: "0 10px 30px rgba(0, 0, 0, 0.03)",
    border: "1px solid #e2e8f0"
  },

  filtersHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: "20px"
  },

  filtersTitle: {
    fontSize: "1.2rem",
    fontWeight: "700",
    color: "#0f172a",
    margin: 0
  },

  activeFilterInfo: {
    display: "flex",
    alignItems: "center",
    gap: "12px"
  },

  activeFilterCount: {
    background: "linear-gradient(135deg, #4f46e5 0%, #4338ca 100%)",
    color: "white",
    padding: "6px 16px",
    borderRadius: "30px",
    fontSize: "0.85rem",
    fontWeight: "700"
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

  filtersGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))",
    gap: "16px"
  },

  filterCard: {
    background: "#ffffff",
    padding: "16px",
    borderRadius: "14px",
    border: "1px solid #cbd5e1",
    transition: "all 0.2s ease"
  },

  filterHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: "8px"
  },

  filterLabel: {
    fontSize: "0.75rem",
    fontWeight: "700",
    color: "#475569",
    textTransform: "uppercase",
    letterSpacing: "0.05em"
  },

  clearSingleFilter: {
    background: "#e2e8f0",
    border: "none",
    color: "#64748b",
    width: "22px",
    height: "22px",
    borderRadius: "50%",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    fontSize: "0.8rem",
    cursor: "pointer",
    padding: 0
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

  filterSelectActive: {
    borderColor: "#4f46e5",
    background: "#ffffff",
    boxShadow: "0 0 0 3px rgba(79, 70, 229, 0.1)"
  },

  filterButtons: {
    display: "flex",
    gap: "8px",
    marginTop: "8px"
  },

  filterButton: {
    flex: 1,
    padding: "8px 12px",
    borderRadius: "10px",
    border: "1px solid #cbd5e1",
    background: "#f8fafc",
    color: "#475569",
    fontSize: "0.85rem",
    fontWeight: "500",
    cursor: "pointer",
    transition: "all 0.2s ease",
    textAlign: "center"
  },

  filterButtonActive: {
    background: "#fef3c7",
    borderColor: "#f59e0b",
    color: "#b45309",
    fontWeight: "700"
  },

  resultsSummary: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    paddingTop: "16px",
    borderTop: "1px solid #f1f5f9"
  },

  resultsInfoCard: {
    display: "flex",
    gap: "20px"
  },

  resultStat: {
    textAlign: "center",
    minWidth: "80px"
  },

  resultNumber: {
    display: "block",
    fontSize: "1.8rem",
    fontWeight: "800",
    color: "#4f46e5",
    marginBottom: "4px"
  },

  resultLabel: {
    fontSize: "0.75rem",
    color: "#64748b",
    textTransform: "uppercase",
    letterSpacing: "0.05em",
    fontWeight: "700"
  },

  resultsText: {
    fontSize: "0.95rem",
    color: "#0f172a",
    fontWeight: "600"
  },

  noResults: {
    color: "#ef4444",
    fontWeight: "600"
  },

  // Table
  tableContainer: { background: "#ffffff", borderRadius: "20px", overflow: "hidden", boxShadow: "0 10px 30px rgba(0, 0, 0, 0.03)", marginBottom: "20px", overflowX: "auto", border: "1px solid #e2e8f0" },
  table: { width: "100%", borderCollapse: "collapse", minWidth: "1450px" },
  tableHeader: { background: "linear-gradient(135deg, #1e1b4b 0%, #312e81 100%)", color: "white", padding: "14px 16px", textAlign: "left", fontWeight: "700", fontSize: "0.82rem", textTransform: "uppercase", letterSpacing: "0.04em" },
  tableRowEven: { background: "#ffffff" },
  tableRowOdd: { background: "#f8fafc" },
  repeatedRow: {
    borderLeft: "4px solid #d97706",
    background: "linear-gradient(90deg, rgba(254, 243, 199, 0.3) 0%, rgba(254, 243, 199, 0.1) 100%)"
  },
  tableCell: {
    padding: "14px 16px",
    borderBottom: "1px solid #e2e8f0",
    fontSize: "0.88rem",
    color: "#0f172a",
    border: "1px solid #e2e8f0",
    textAlign: "center",
    verticalAlign: "top",
    wordWrap: "break-word",
    whiteSpace: "normal",
    maxWidth: "150px"
  },
  repeatedLotCell: {
    background: "#fef3c7",
    fontWeight: "600",
    color: "#92400e"
  },
  repeatedDaysCell: {
    background: "#fecaca",
    fontWeight: "800"
  },

  // day backgrounds
  dayOk: { background: "#E4F7E4" },
  dayLate: { background: "#FDE2E2" },
  dayNeutral: {},

  loadingContainer: { display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", minHeight: "400px", background: "#ffffff", borderRadius: "20px", border: "1px solid #e2e8f0" },
  spinner: { width: "50px", height: "50px", border: "4px solid #e2e8f0", borderLeft: "4px solid #4f46e5", borderRadius: "50%", animation: "spin 1s linear infinite", marginBottom: "20px" },
  loadingText: { fontSize: "1.2rem", color: "#4f46e5", fontWeight: "600" },
  errorContainer: { textAlign: "center", padding: "60px 20px", background: "white", borderRadius: "20px", boxShadow: "0 10px 30px rgba(0,0,0,0.03)", border: "1px solid #e2e8f0", maxWidth: "500px", margin: "40px auto" },
  errorIcon: { fontSize: "4rem", marginBottom: "20px" },
  errorTitle: { color: "#dc2626", marginBottom: "10px", fontSize: "1.5rem" },
  errorMessage: { color: "#6b7280", marginBottom: "30px" },
  retryButton: { background: "linear-gradient(135deg,#ef4444,#dc2626)", color: "white", border: "none", padding: "12px 30px", borderRadius: "50px", fontSize: "1rem", fontWeight: "600", cursor: "pointer" },
  footerText: { margin: 0 },
  tableImage: {
    width: "40px",
    height: "40px",
    objectFit: "cover",
    borderRadius: "6px",
    border: "1px solid #e2e8f0",
    cursor: "pointer"
  },
  noImagePlaceholder: {
    fontSize: "9px",
    color: "#94a3b8",
    textTransform: "uppercase",
    letterSpacing: "0.05em"
  },
  imageModalBackdrop: {
    position: "fixed",
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    backgroundColor: "rgba(0, 0, 0, 0.75)",
    zIndex: 9999,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    backdropFilter: "blur(4px)"
  },
  imageModalContent: {
    position: "relative",
    maxWidth: "90%",
    maxHeight: "90%",
    backgroundColor: "white",
    borderRadius: "12px",
    padding: "16px",
    boxShadow: "0 25px 50px -12px rgba(0, 0, 0, 0.5)"
  },
  imageModalImg: {
    maxWidth: "100%",
    maxHeight: "80vh",
    borderRadius: "6px",
    objectFit: "contain"
  },
  imageModalClose: {
    position: "absolute",
    top: "-12px",
    right: "-12px",
    backgroundColor: "#0f172a",
    color: "white",
    border: "none",
    width: "28px",
    height: "28px",
    borderRadius: "50%",
    fontSize: "18px",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    cursor: "pointer",
    boxShadow: "0 4px 6px rgba(0, 0, 0, 0.15)"
  }
};

// Add CSS animations
if (typeof document !== 'undefined') {
  const style = document.createElement('style');
  style.textContent = `
    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }
    .search-input:focus { 
      border-color: #4f46e5; 
      box-shadow: 0 0 0 3px rgba(79, 70, 229, 0.1); 
    }
    .filter-select:focus { 
      border-color: #4f46e5; 
      box-shadow: 0 0 0 3px rgba(79, 70, 229, 0.1); 
    }
  `;
  document.head.appendChild(style);
}