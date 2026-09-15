import React, { useEffect, useMemo, useState } from "react";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
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
  "Style",
  "Fabric",
  "Brand",
  "Challan Total Qty",
  "Section",
  "Season",
  "Party Name",
  "Direct Stitching",
  "Priority", // Added Priority for repeated lots
  "Printing",
  "Challan No",
  "Challan Date",
];

/* ================== EXPORT CONFIG & MODAL ================== */
const ALL_EXPORT_COLUMNS = [
  { id: 'S. No', label: 'S. No', key: 'S. No', defaultSelected: true, baseWidth: 28 },
  { id: 'Lot No.', label: 'Lot No.', key: 'Lot No.', defaultSelected: true, baseWidth: 55 },
  { id: 'Garment Type', label: 'Garment Type', key: 'Garment Type', defaultSelected: true, baseWidth: 65 },
  { id: 'Style', label: 'Style', key: 'Style', defaultSelected: true, baseWidth: 70 },
  { id: 'Fabric', label: 'Fabric', key: 'Fabric', defaultSelected: true, baseWidth: 60 },
  { id: 'Brand', label: 'Brand', key: 'Brand', defaultSelected: true, baseWidth: 65 },
  { id: 'Challan Total Qty', label: 'Challan Qty', key: 'Challan Total Qty', defaultSelected: true, baseWidth: 45 },
  { id: 'Section', label: 'Section (M/W/K)', key: 'Section', defaultSelected: true, baseWidth: 32 },
  { id: 'Season', label: 'Season', key: 'Season', defaultSelected: true, baseWidth: 45 },
  { id: 'Party Name', label: 'Party Name', key: 'Party Name', defaultSelected: true, baseWidth: 45 },
  { id: 'Direct Stitching', label: 'Direct Stitching', key: 'Direct Stitching', defaultSelected: true, baseWidth: 45 },
  { id: 'Printing', label: 'Printing Party', key: 'Printing', defaultSelected: true, baseWidth: 45 },
  { id: 'Challan Date', label: 'Challan Date', key: 'Challan Date', defaultSelected: true, baseWidth: 55 },
  { id: 'Printing Status', label: 'Printing Status', key: 'Printing Status', defaultSelected: true, baseWidth: 58 },
  { id: 'Pending Challan Shade', label: 'Pending Shade', key: 'Pending Challan Shade', defaultSelected: true, baseWidth: 70 },
  { id: 'Remarks', label: 'Remarks', key: 'Remarks', defaultSelected: true, baseWidth: 60 },
  { id: 'Printing Done', label: 'Printing Done Date', key: 'Printing Done', defaultSelected: true, baseWidth: 55 },
  { id: 'Days', label: 'Days', key: 'Days', defaultSelected: true, baseWidth: 35 },
  { id: 'HOD Remarks', label: 'HOD Remarks', key: 'HOD Remarks', defaultSelected: true, baseWidth: 80 },
];

const ExportColumnModal = ({
  isOpen,
  onClose,
  exportType,
  allColumns,
  selectedColumns,
  setSelectedColumns,
  onConfirmExport
}) => {
  const [searchTerm, setSearchTerm] = useState('');

  if (!isOpen) return null;

  const isExcel = exportType === 'excel';
  const headerBg = isExcel
    ? 'linear-gradient(135deg, #065f46 0%, #047857 50%, #059669 100%)'
    : 'linear-gradient(135deg, #1e1b4b 0%, #312e81 50%, #4338ca 100%)';
  const confirmBtnBg = isExcel
    ? 'linear-gradient(135deg, #10b981 0%, #059669 100%)'
    : 'linear-gradient(135deg, #6366f1 0%, #4f46e5 100%)';

  const handleToggle = (id) => {
    if (selectedColumns.includes(id)) {
      setSelectedColumns(selectedColumns.filter(colId => colId !== id));
    } else {
      setSelectedColumns([...selectedColumns, id]);
    }
  };

  const handleSelectAll = () => {
    setSelectedColumns(allColumns.map(c => c.id));
  };

  const handleClearAll = () => {
    setSelectedColumns([]);
  };

  const handleResetDefault = () => {
    setSelectedColumns(allColumns.filter(c => c.defaultSelected !== false).map(c => c.id));
  };

  const filteredColumns = allColumns.filter(c =>
    c.label.toLowerCase().includes(searchTerm.toLowerCase()) ||
    c.id.toLowerCase().includes(searchTerm.toLowerCase())
  );

  return (
    <div
      style={{
        position: 'fixed',
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        backgroundColor: 'rgba(15, 23, 42, 0.7)',
        backdropFilter: 'blur(6px)',
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
          borderRadius: '20px',
          maxWidth: '660px',
          width: '100%',
          maxHeight: '85vh',
          display: 'flex',
          flexDirection: 'column',
          boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.35)',
          overflow: 'hidden',
          border: '1px solid rgba(226, 232, 240, 0.8)'
        }}
        onClick={(e) => e.stopPropagation()}
      >
        <div
          style={{
            background: headerBg,
            color: '#ffffff',
            padding: '20px 24px',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between'
          }}
        >
          <div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
              <span style={{ fontSize: '1.4rem' }}>{isExcel ? '📊' : '📄'}</span>
              <h3 style={{ margin: 0, fontSize: '1.2rem', fontWeight: '800', letterSpacing: '-0.02em' }}>
                Select Headers for {isExcel ? 'Excel (.xlsx)' : 'PDF'} Export
              </h3>
            </div>
            <p style={{ margin: '4px 0 0 0', fontSize: '0.85rem', color: isExcel ? '#a7f3d0' : '#c7d2fe' }}>
              Choose which columns you want to include in the downloaded file
            </p>
          </div>
          <button
            onClick={onClose}
            style={{
              background: 'rgba(255, 255, 255, 0.2)',
              border: 'none',
              color: '#ffffff',
              fontSize: '1.2rem',
              borderRadius: '10px',
              width: '36px',
              height: '36px',
              cursor: 'pointer',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center'
            }}
          >
            ✕
          </button>
        </div>

        <div
          style={{
            padding: '14px 24px',
            background: '#f8fafc',
            borderBottom: '1px solid #e2e8f0',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            flexWrap: 'wrap',
            gap: '10px'
          }}
        >
          <div style={{ display: 'flex', gap: '8px' }}>
            <button
              onClick={handleSelectAll}
              style={{
                padding: '6px 12px',
                fontSize: '12px',
                fontWeight: '600',
                background: '#ffffff',
                border: '1px solid #cbd5e1',
                borderRadius: '6px',
                color: '#2563eb',
                cursor: 'pointer'
              }}
            >
              ✓ Select All
            </button>
            <button
              onClick={handleClearAll}
              style={{
                padding: '6px 12px',
                fontSize: '12px',
                fontWeight: '600',
                background: '#ffffff',
                border: '1px solid #cbd5e1',
                borderRadius: '6px',
                color: '#dc2626',
                cursor: 'pointer'
              }}
            >
              ✕ Clear All
            </button>
            <button
              onClick={handleResetDefault}
              style={{
                padding: '6px 12px',
                fontSize: '12px',
                fontWeight: '600',
                background: '#ffffff',
                border: '1px solid #cbd5e1',
                borderRadius: '6px',
                color: '#475569',
                cursor: 'pointer'
              }}
            >
              ↺ Reset Default
            </button>
          </div>

          <div style={{
            fontSize: '13px',
            fontWeight: '700',
            color: selectedColumns.length > 0 ? (isExcel ? '#059669' : '#4f46e5') : '#dc2626',
            background: selectedColumns.length > 0 ? (isExcel ? '#ecfdf5' : '#eef2ff') : '#fee2e2',
            padding: '4px 12px',
            borderRadius: '20px',
            border: `1px solid ${selectedColumns.length > 0 ? (isExcel ? '#a7f3d0' : '#c7d2fe') : '#fecaca'}`
          }}>
            {selectedColumns.length} of {allColumns.length} headers selected
          </div>
        </div>

        <div style={{ padding: '12px 24px 0 24px' }}>
          <input
            type="text"
            placeholder="🔍 Search column headers..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            style={{
              width: '100%',
              padding: '8px 12px',
              borderRadius: '8px',
              border: '1.5px solid #cbd5e1',
              fontSize: '13px',
              outline: 'none',
              boxSizing: 'border-box'
            }}
          />
        </div>

        <div style={{
          padding: '16px 24px',
          overflowY: 'auto',
          flex: 1,
          display: 'grid',
          gridTemplateColumns: 'repeat(auto-fill, minmax(180px, 1fr))',
          gap: '10px'
        }}>
          {filteredColumns.map(col => {
            const isSelected = selectedColumns.includes(col.id);
            return (
              <label
                key={col.id}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '10px',
                  padding: '10px 12px',
                  borderRadius: '10px',
                  border: `1.5px solid ${isSelected ? (isExcel ? '#059669' : '#6366f1') : '#e2e8f0'}`,
                  background: isSelected ? (isExcel ? '#f0fdf4' : '#f5f3ff') : '#ffffff',
                  cursor: 'pointer',
                  transition: 'all 0.15s ease',
                  userSelect: 'none'
                }}
              >
                <input
                  type="checkbox"
                  checked={isSelected}
                  onChange={() => handleToggle(col.id)}
                  style={{
                    width: '17px',
                    height: '17px',
                    accentColor: isExcel ? '#059669' : '#4f46e5',
                    cursor: 'pointer'
                  }}
                />
                <span style={{
                  fontSize: '13px',
                  fontWeight: isSelected ? '700' : '500',
                  color: isSelected ? (isExcel ? '#065f46' : '#1e1b4b') : '#475569'
                }}>
                  {col.label}
                </span>
              </label>
            );
          })}
        </div>

        <div
          style={{
            padding: '16px 24px',
            background: '#f8fafc',
            borderTop: '1px solid #e2e8f0',
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center'
          }}
        >
          <button
            onClick={onClose}
            style={{
              background: '#e2e8f0',
              color: '#334155',
              border: 'none',
              padding: '10px 20px',
              borderRadius: '10px',
              fontWeight: '700',
              fontSize: '0.9rem',
              cursor: 'pointer'
            }}
          >
            Cancel
          </button>
          <button
            onClick={() => onConfirmExport(selectedColumns)}
            disabled={selectedColumns.length === 0}
            style={{
              background: confirmBtnBg,
              color: '#ffffff',
              border: 'none',
              padding: '10px 26px',
              borderRadius: '10px',
              fontWeight: '800',
              fontSize: '0.92rem',
              cursor: selectedColumns.length === 0 ? 'not-allowed' : 'pointer',
              opacity: selectedColumns.length === 0 ? 0.5 : 1,
              boxShadow: isExcel ? '0 4px 12px rgba(16, 185, 129, 0.3)' : '0 4px 12px rgba(99, 102, 241, 0.35)',
              display: 'flex',
              alignItems: 'center',
              gap: '8px'
            }}
          >
            <span>{isExcel ? '📊 Download Excel' : '📄 Download PDF'}</span>
            <span style={{ fontSize: '0.85rem', opacity: 0.9 }}>({selectedColumns.length} cols)</span>
          </button>
        </div>
      </div>
    </div>
  );
};

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
  "Garment Type",
  "Style",
  "Fabric",
  "Brand",
  "Challan Total Qty",
  "Section",
  "Season",
  "Party Name",
  "Direct Stitching",
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
  "Season",
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
  const [seasonFilter, setSeasonFilter] = useState([]); // Season filter
  const [garmentTypeFilter, setGarmentTypeFilter] = useState([]);
  const [sectionFilter, setSectionFilter] = useState([]);
  const [partyNameFilter, setPartyNameFilter] = useState([]);
  const [printingFilterDropdown, setPrintingFilterDropdown] = useState([]);
  const [priorityFilter, setPriorityFilter] = useState([]);

  // Printing Status filter
  const [printingStatusFilter, setPrintingStatusFilter] = useState(initialPrintingStatusFilter);
  const [daysFilter, setDaysFilter] = useState("all");
  const [financialYearFilter, setFinancialYearFilter] = useState(getCurrentFinancialYear());

  // Export Header Selection Modal State
  const [exportModalConfig, setExportModalConfig] = useState({ isOpen: false, type: 'excel' });
  const [selectedExportColumns, setSelectedExportColumns] = useState(() => ALL_EXPORT_COLUMNS.filter(c => c.defaultSelected !== false).map(c => c.id));

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
      const seasonCol = headers ? headers.findIndex(h => h && h.trim().toLowerCase() === 'season') : -1;

      const formatted = (rows || [])
        .map((row) => {
          const entry = {};
          (headers || []).forEach((header, i) => (entry[header] = row[i] ?? ""));
          const seasonVal = seasonCol !== -1 && row[seasonCol] !== undefined
            ? String(row[seasonCol] || '').trim()
            : String(row[headers.indexOf('SEASON')] || row[headers.indexOf('Season')] || '').trim();
          entry["Season"] = seasonVal;
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

      if (seasonFilter && seasonFilter.length > 0) {
        const rowSeason = normalizeValue(row["Season"] || row["SEASON"]);
        const matched = seasonFilter.some(s => normalizeValue(s) === rowSeason);
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
    fabricFilter, brandFilter, styleFilter, seasonFilter, garmentTypeFilter,
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
      base["Season"] = String(row["Season"] || "");

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
      case "Season": return seasonFilter;
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
      case "Season": setSeasonFilter(value); break;
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

  const exportToExcelWithColumns = async (selectedCols) => {
    if (!exportedRows.length || !selectedCols.length) return;
    try {
      const now = new Date();
      const reportDateStr = `${String(now.getDate()).padStart(2, '0')}/${String(now.getMonth() + 1).padStart(2, '0')}/${now.getFullYear()} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;

      const totalLots = exportedRows.length;
      let totalQty = 0;
      let repeatedLotsCount = 0;
      let redZoneLots = 0;
      let redZoneQty = 0;
      let normalLots = 0;
      let normalQty = 0;

      const garmentAnalysis = {};
      const printingPartyAnalysis = {};

      exportedRows.forEach((row) => {
        const qty = parseFloat(row["Challan Total Qty"]) || 0;
        totalQty += qty;

        const isRepeated = !!row["_isRepeated"];
        if (isRepeated) repeatedLotsCount++;

        const days = parseFloat(row["Days"]);
        if (!isNaN(days) && days > 5) {
          redZoneLots++;
          redZoneQty += qty;
        } else {
          normalLots++;
          normalQty += qty;
        }

        const garment = String(row["Garment Type"] || "Unassigned").trim() || "Unassigned";
        if (!garmentAnalysis[garment]) {
          garmentAnalysis[garment] = { name: garment, lots: 0, qty: 0 };
        }
        garmentAnalysis[garment].lots += 1;
        garmentAnalysis[garment].qty += qty;

        const printingParty = String(row["Printing"] || "Unassigned").trim() || "Unassigned";
        if (!printingPartyAnalysis[printingParty]) {
          printingPartyAnalysis[printingParty] = { name: printingParty, lots: 0, qty: 0 };
        }
        printingPartyAnalysis[printingParty].lots += 1;
        printingPartyAnalysis[printingParty].qty += qty;
      });

      const sortedGarments = Object.values(garmentAnalysis).sort((a, b) => b.qty - a.qty);
      const sortedPrintingParties = Object.values(printingPartyAnalysis).sort((a, b) => b.qty - a.qty);

      const activeColumns = selectedCols.map(id => {
        const def = ALL_EXPORT_COLUMNS.find(c => c.id === id);
        return def || { id, label: id, key: id, baseWidth: 50 };
      });

      const workbook = new ExcelJS.Workbook();
      workbook.creator = "Factory Suite Pro";
      workbook.created = now;

      // Styling Helpers
      const thinBorder = {
        top: { style: 'thin', color: { argb: 'CBD5E1' } },
        left: { style: 'thin', color: { argb: 'CBD5E1' } },
        bottom: { style: 'thin', color: { argb: 'CBD5E1' } },
        right: { style: 'thin', color: { argb: 'CBD5E1' } }
      };

      const headerBorder = {
        top: { style: 'thin', color: { argb: '0F172A' } },
        left: { style: 'thin', color: { argb: '0F172A' } },
        bottom: { style: 'medium', color: { argb: '0F172A' } },
        right: { style: 'thin', color: { argb: '0F172A' } }
      };

      const totalBorder = {
        top: { style: 'thin', color: { argb: '0F172A' } },
        left: { style: 'thin', color: { argb: 'CBD5E1' } },
        bottom: { style: 'double', color: { argb: '0F172A' } },
        right: { style: 'thin', color: { argb: 'CBD5E1' } }
      };

      // ==========================================
      // SHEET 1: CHALLAN RECORDS
      // ==========================================
      const ws1 = workbook.addWorksheet("Challan Records", {
        views: [{ showGridLines: true, state: 'frozen', xSplit: 0, ySplit: 6 }]
      });

      const numCols1 = activeColumns.length;
      ws1.columns = activeColumns.map(col => ({
        header: col.label,
        key: col.id,
        width: Math.min(Math.max(Math.round((col.baseWidth || 50) / 3.8), 12), 40)
      }));

      // Row 1: Title Banner
      const titleRow1 = ws1.getRow(1);
      titleRow1.values = ["FACTORY SUITE PRO - PRINTING CHALLAN PRODUCTION REPORT"];
      ws1.mergeCells(1, 1, 1, numCols1);
      titleRow1.font = { name: "Segoe UI", size: 14, bold: true, color: { argb: "FFFFFFFF" } };
      titleRow1.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F4C81" } };
      titleRow1.alignment = { vertical: "middle", horizontal: "center" };
      titleRow1.height = 34;

      // Row 2: Metadata Banner
      const metaRow1 = ws1.getRow(2);
      metaRow1.values = [`Exported on: ${reportDateStr}  |  Total Lots: ${totalLots}  |  Total Quantity: ${totalQty.toLocaleString()}  |  Printing Parties: ${sortedPrintingParties.length}  |  Garment Types: ${sortedGarments.length}  |  Repeated Lots: ${repeatedLotsCount}`];
      ws1.mergeCells(2, 1, 2, numCols1);
      metaRow1.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF1E293B" } };
      metaRow1.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } };
      metaRow1.alignment = { vertical: "middle", horizontal: "center" };
      metaRow1.height = 22;

      // Row 3: Blank Row
      const blankRow3 = ws1.getRow(3);
      blankRow3.values = [];
      blankRow3.height = 6;

      // Row 4: KPI Summary Banner
      const normalPct = totalLots > 0 ? Math.round((normalLots / totalLots) * 100) : 0;
      const redPct = totalLots > 0 ? Math.round((redZoneLots / totalLots) * 100) : 0;
      const kpiRow1 = ws1.getRow(4);
      kpiRow1.values = [`KPI SUMMARY  |  TOTAL LOTS: ${totalLots}  |  TOTAL QTY: ${totalQty.toLocaleString()}  |  NORMAL (<=5d): ${normalLots} (${normalPct}%)  |  RED ZONE (>5d): ${redZoneLots} (${redPct}%)  |  REPEATED LOTS: ${repeatedLotsCount}`];
      ws1.mergeCells(4, 1, 4, numCols1);
      kpiRow1.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FF0369A1" } };
      kpiRow1.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE0F2FE" } };
      kpiRow1.alignment = { vertical: "middle", horizontal: "center" };
      kpiRow1.height = 24;

      // Row 5: Blank Row
      const blankRow5 = ws1.getRow(5);
      blankRow5.values = [];
      blankRow5.height = 6;

      // Row 6: Table Headers Row
      const headerRow1 = ws1.getRow(6);
      headerRow1.values = activeColumns.map(c => c.label);
      headerRow1.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
      headerRow1.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F4C81" } };
      headerRow1.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      headerRow1.height = 30;
      for (let c = 1; c <= numCols1; c++) {
        headerRow1.getCell(c).border = headerBorder;
      }

      // Add Data Rows
      let rIdx1 = 7;
      exportedRows.forEach((row, index) => {
        const lotNum = row["Lot No."] ? String(row["Lot No."]).replace(/★\s*/, '').trim() : "";
        const isRepeated = !!row["_isRepeated"];
        const lotHist = headRemarksMap[lotNum] || [];
        const latestHod = lotHist.length > 0 ? lotHist[lotHist.length - 1].text : "";
        const days = parseFloat(row["Days"]);

        const rowValues = activeColumns.map(col => {
          if (col.id === 'S. No') return index + 1;
          if (col.id === 'Lot No.') return lotNum || "—";
          if (col.id === 'HOD Remarks') return latestHod || "—";
          if (col.id === 'Challan Total Qty') {
            const q = parseFloat(row[col.key]);
            return isNaN(q) ? 0 : q;
          }
          if (col.id === 'Days') {
            return !isNaN(days) ? days : (row[col.key] || "—");
          }
          return row[col.key] || "—";
        });

        const dataRow = ws1.getRow(rIdx1);
        dataRow.values = rowValues;
        dataRow.height = 22;

        activeColumns.forEach((col, colIdx) => {
          const cell = dataRow.getCell(colIdx + 1);
          cell.border = thinBorder;
          cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
          cell.font = { name: "Segoe UI", size: 9.5, color: { argb: "FF0F172A" } };

          if (col.id === 'Challan Total Qty') {
            cell.numFmt = '#,##0';
          }

          if (col.id === 'Days' && !isNaN(days)) {
            if (days > 5) {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
              cell.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FF991B1B' } };
            } else {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } };
              cell.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FF15803D' } };
            }
          }

          if (isRepeated && (col.id === 'Lot No.' || col.id === 'S. No')) {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
            cell.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFB45309' } };
          }
        });

        rIdx1++;
      });

      // Total Row on Sheet 1
      const totalRow1 = ws1.getRow(rIdx1);
      const totalRowValues1 = activeColumns.map(col => {
        if (col.id === 'S. No' || col.id === 'Lot No.') return `TOTAL: ${totalLots} LOTS`;
        if (col.id === 'Challan Total Qty') return totalQty;
        return "";
      });
      totalRow1.values = totalRowValues1;
      totalRow1.height = 26;
      totalRow1.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
      totalRow1.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F172A" } };

      for (let c = 1; c <= numCols1; c++) {
        const cell = totalRow1.getCell(c);
        cell.border = totalBorder;
        cell.alignment = { vertical: "middle", horizontal: "center" };
        if (activeColumns[c - 1].id === 'Challan Total Qty') {
          cell.numFmt = '#,##0';
        }
      }

      // ==========================================
      // SHEET 2: EXECUTIVE SUMMARY & BREAKDOWN
      // ==========================================
      const ws2 = workbook.addWorksheet("Executive Summary", {
        views: [{ showGridLines: true }]
      });

      ws2.columns = [
        { header: "Category / Name", key: "col1", width: 28 },
        { header: "Total Lots", key: "col2", width: 16 },
        { header: "Share %", key: "col3", width: 16 },
        { header: "Total Pieces (Qty)", key: "col4", width: 22 }
      ];

      // Row 1: Executive Title
      const exTitle = ws2.getRow(1);
      exTitle.values = ["PRINTING PRODUCTION - EXECUTIVE SUMMARY & BREAKDOWN"];
      ws2.mergeCells(1, 1, 1, 4);
      exTitle.font = { name: "Segoe UI", size: 14, bold: true, color: { argb: "FFFFFFFF" } };
      exTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F4C81" } };
      exTitle.alignment = { vertical: "middle", horizontal: "center" };
      exTitle.height = 32;

      // Row 2: Metadata
      const exMeta = ws2.getRow(2);
      exMeta.values = [`Generated on: ${reportDateStr}  |  Total Records: ${totalLots} Lots  |  Total Quantity: ${totalQty.toLocaleString()} Pcs`];
      ws2.mergeCells(2, 1, 2, 4);
      exMeta.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF1E293B" } };
      exMeta.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F5F9" } };
      exMeta.alignment = { vertical: "middle", horizontal: "center" };
      exMeta.height = 22;

      let rIdx2 = 4;

      // --- SECTION 1: GARMENT TYPE BREAKDOWN ---
      const gHeaderRow = ws2.getRow(rIdx2);
      gHeaderRow.values = ["🧵 1. GARMENT TYPE BREAKDOWN (LOTS & PIECES)"];
      ws2.mergeCells(rIdx2, 1, rIdx2, 4);
      gHeaderRow.font = { name: "Segoe UI", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      gHeaderRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F766E" } }; // Deep Teal
      gHeaderRow.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
      gHeaderRow.height = 26;
      rIdx2++;

      const gColHeaderRow = ws2.getRow(rIdx2);
      gColHeaderRow.values = ["Garment Type", "Total Lots", "Share %", "Total Pieces"];
      gColHeaderRow.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
      gColHeaderRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF14B8A6" } };
      gColHeaderRow.alignment = { vertical: "middle", horizontal: "center" };
      gColHeaderRow.height = 24;
      for (let c = 1; c <= 4; c++) gColHeaderRow.getCell(c).border = headerBorder;
      rIdx2++;

      sortedGarments.forEach((item) => {
        const row = ws2.getRow(rIdx2);
        const share = totalQty > 0 ? ((item.qty / totalQty) * 100).toFixed(1) : "0.0";
        row.values = [item.name, item.lots, `${share}%`, item.qty];
        row.height = 20;

        for (let c = 1; c <= 4; c++) {
          const cell = row.getCell(c);
          cell.border = thinBorder;
          cell.alignment = { vertical: "middle", horizontal: "center" };
          cell.font = { name: "Segoe UI", size: 9.5, color: { argb: "FF0F172A" } };
        }
        row.getCell(4).numFmt = '#,##0';
        row.getCell(2).numFmt = '#,##0';
        rIdx2++;
      });

      const gTotalRow = ws2.getRow(rIdx2);
      gTotalRow.values = ["TOTAL GARMENTS", totalLots, "100.0%", totalQty];
      gTotalRow.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
      gTotalRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F766E" } };
      gTotalRow.height = 24;
      for (let c = 1; c <= 4; c++) {
        const cell = gTotalRow.getCell(c);
        cell.border = totalBorder;
        cell.alignment = { vertical: "middle", horizontal: "center" };
      }
      gTotalRow.getCell(4).numFmt = '#,##0';
      gTotalRow.getCell(2).numFmt = '#,##0';
      rIdx2 += 3;

      // --- SECTION 2: PRINTING PARTY SUMMARY ---
      const pHeaderRow = ws2.getRow(rIdx2);
      pHeaderRow.values = ["🏢 2. PRINTING PARTY SUMMARY & WORKLOAD ALLOCATION"];
      ws2.mergeCells(rIdx2, 1, rIdx2, 4);
      pHeaderRow.font = { name: "Segoe UI", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      pHeaderRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E40AF" } }; // Royal Blue
      pHeaderRow.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
      pHeaderRow.height = 26;
      rIdx2++;

      const pColHeaderRow = ws2.getRow(rIdx2);
      pColHeaderRow.values = ["Printing Party", "Total Lots", "Share %", "Total Pieces"];
      pColHeaderRow.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
      pColHeaderRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF3B82F6" } };
      pColHeaderRow.alignment = { vertical: "middle", horizontal: "center" };
      pColHeaderRow.height = 24;
      for (let c = 1; c <= 4; c++) pColHeaderRow.getCell(c).border = headerBorder;
      rIdx2++;

      sortedPrintingParties.forEach((item) => {
        const row = ws2.getRow(rIdx2);
        const share = totalQty > 0 ? ((item.qty / totalQty) * 100).toFixed(1) : "0.0";
        row.values = [item.name, item.lots, `${share}%`, item.qty];
        row.height = 20;

        for (let c = 1; c <= 4; c++) {
          const cell = row.getCell(c);
          cell.border = thinBorder;
          cell.alignment = { vertical: "middle", horizontal: "center" };
          cell.font = { name: "Segoe UI", size: 9.5, color: { argb: "FF0F172A" } };
        }
        row.getCell(4).numFmt = '#,##0';
        row.getCell(2).numFmt = '#,##0';
        rIdx2++;
      });

      const pTotalRow = ws2.getRow(rIdx2);
      pTotalRow.values = ["TOTAL PRINTING PARTIES", totalLots, "100.0%", totalQty];
      pTotalRow.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
      pTotalRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E40AF" } };
      pTotalRow.height = 24;
      for (let c = 1; c <= 4; c++) {
        const cell = pTotalRow.getCell(c);
        cell.border = totalBorder;
        cell.alignment = { vertical: "middle", horizontal: "center" };
      }
      pTotalRow.getCell(4).numFmt = '#,##0';
      pTotalRow.getCell(2).numFmt = '#,##0';
      rIdx2 += 3;

      // --- SECTION 3: DAYS AGING BREAKDOWN ---
      const aHeaderRow = ws2.getRow(rIdx2);
      aHeaderRow.values = ["⏱️ 3. SLA & DAYS AGING BREAKDOWN"];
      ws2.mergeCells(rIdx2, 1, rIdx2, 4);
      aHeaderRow.font = { name: "Segoe UI", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      aHeaderRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFB45309" } }; // Amber
      aHeaderRow.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
      aHeaderRow.height = 26;
      rIdx2++;

      const aColHeaderRow = ws2.getRow(rIdx2);
      aColHeaderRow.values = ["Aging Status", "Lot Count", "Share %", "Total Pieces"];
      aColHeaderRow.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
      aColHeaderRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF59E0B" } };
      aColHeaderRow.alignment = { vertical: "middle", horizontal: "center" };
      aColHeaderRow.height = 24;
      for (let c = 1; c <= 4; c++) aColHeaderRow.getCell(c).border = headerBorder;
      rIdx2++;

      // Row Normal (<=5d)
      const normRow = ws2.getRow(rIdx2);
      normRow.values = ["<= 5 Days (On-Time / Normal)", normalLots, `${totalLots > 0 ? ((normalLots / totalLots) * 100).toFixed(1) : 0}%`, normalQty];
      normRow.height = 20;
      for (let c = 1; c <= 4; c++) {
        const cell = normRow.getCell(c);
        cell.border = thinBorder;
        cell.alignment = { vertical: "middle", horizontal: "center" };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } };
        cell.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF15803D" } };
      }
      normRow.getCell(4).numFmt = '#,##0';
      normRow.getCell(2).numFmt = '#,##0';
      rIdx2++;

      // Row Red Zone (>5d)
      const redRow = ws2.getRow(rIdx2);
      redRow.values = ["> 5 Days (Red Zone / Delayed)", redZoneLots, `${totalLots > 0 ? ((redZoneLots / totalLots) * 100).toFixed(1) : 0}%`, redZoneQty];
      redRow.height = 20;
      for (let c = 1; c <= 4; c++) {
        const cell = redRow.getCell(c);
        cell.border = thinBorder;
        cell.alignment = { vertical: "middle", horizontal: "center" };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
        cell.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF991B1B" } };
      }
      redRow.getCell(4).numFmt = '#,##0';
      redRow.getCell(2).numFmt = '#,##0';
      rIdx2++;

      const aTotalRow = ws2.getRow(rIdx2);
      aTotalRow.values = ["TOTAL SLA PERFORMANCE", totalLots, "100.0%", totalQty];
      aTotalRow.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
      aTotalRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFB45309" } };
      aTotalRow.height = 24;
      for (let c = 1; c <= 4; c++) {
        const cell = aTotalRow.getCell(c);
        cell.border = totalBorder;
        cell.alignment = { vertical: "middle", horizontal: "center" };
      }
      aTotalRow.getCell(4).numFmt = '#,##0';
      aTotalRow.getCell(2).numFmt = '#,##0';

      // ==========================================
      // SHEET 3: APPLIED FILTERS AUDIT
      // ==========================================
      const ws3 = workbook.addWorksheet("Applied Filters", {
        views: [{ showGridLines: true }]
      });

      ws3.columns = [
        { header: "Filter Parameter", key: "param", width: 28 },
        { header: "Active Selection", key: "value", width: 50 }
      ];

      const fTitle = ws3.getRow(1);
      fTitle.values = ["APPLIED REPORT FILTERS & AUDIT METADATA"];
      ws3.mergeCells(1, 1, 1, 2);
      fTitle.font = { name: "Segoe UI", size: 13, bold: true, color: { argb: "FFFFFFFF" } };
      fTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F4C81" } };
      fTitle.alignment = { vertical: "middle", horizontal: "center" };
      fTitle.height = 30;

      const filterList = [
        { param: "Financial Year", val: financialYearFilter || "ALL" },
        { param: "Printing Status Filter", val: (printingStatusFilter || "ALL").toUpperCase() },
        { param: "Days SLA Filter", val: daysFilter.toUpperCase() },
        { param: "Fabric Filter", val: fabricFilter.length > 0 ? fabricFilter.join(", ") : "All Fabrics" },
        { param: "Brand Filter", val: brandFilter.length > 0 ? brandFilter.join(", ") : "All Brands" },
        { param: "Style Filter", val: styleFilter.length > 0 ? styleFilter.join(", ") : "All Styles" },
        { param: "Season Filter", val: seasonFilter.length > 0 ? seasonFilter.join(", ") : "All Seasons" },
        { param: "Garment Type Filter", val: garmentTypeFilter.length > 0 ? garmentTypeFilter.join(", ") : "All Garments" },
        { param: "Section Filter", val: sectionFilter.length > 0 ? sectionFilter.join(", ") : "All Sections" },
        { param: "Party Name Filter", val: partyNameFilter.length > 0 ? partyNameFilter.join(", ") : "All Parties" },
        { param: "Printing Dropdown Filter", val: printingFilterDropdown.length > 0 ? printingFilterDropdown.join(", ") : "All Printing Parties" },
        { param: "Priority Filter", val: priorityFilter.length > 0 ? priorityFilter.join(", ") : "All Priorities" },
        { param: "Search Query", val: searchTerm || "None" },
        { param: "Export Timestamp", val: reportDateStr },
      ];

      let rIdx3 = 3;
      filterList.forEach(f => {
        const r = ws3.getRow(rIdx3);
        r.values = [f.param, f.val];
        r.height = 20;
        const cell1 = r.getCell(1);
        const cell2 = r.getCell(2);
        cell1.border = thinBorder;
        cell2.border = thinBorder;
        cell1.font = { name: "Segoe UI", size: 9.5, bold: true, color: { argb: "FF1E293B" } };
        cell2.font = { name: "Segoe UI", size: 9.5, color: { argb: "FF334155" } };
        cell1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        rIdx3++;
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const filename = `Printing-Challans-${new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-")}.xlsx`;
      saveAs(new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), filename);

    } catch (e) {
      console.error("ExcelJS export error:", e);
      alert(`Excel export error: ${e.message}`);
    }
  };

  const exportToPDFWithColumns = async (selectedCols) => {
    if (!exportedRows.length || !selectedCols.length) return;

    try {
      const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "A3" });

      const COLOR = {
        headerFill: [15, 76, 129],      // Deep Navy Blue
        headerText: [255, 255, 255],    // White bold text
        grid: [0, 0, 0],                // Solid Black grid lines
        stripe: [248, 250, 252],        // Light slate alternating stripe
        text: [0, 0, 0],                // Crisp Pure Black Text
        accent: [0, 0, 0],              // Pure Black Accent
        pageBorder: [0, 0, 0],          // Pure Black page border
        summaryBg: [239, 246, 255],     // Light blue summary box (#eff6ff)
        summaryText: [0, 0, 0],         // Pure Black summary text
        totalBg: [239, 246, 255],       // Soft Navy/Blue total summary banner
        repeatedLotBg: [254, 243, 199], // Light amber highlight (#fef3c7)
      };

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

      const totalLots = exportedRows.length;
      const totalPcs = exportedRows.reduce((sum, row) => {
        const qty = parseFloat(row["Challan Total Qty"]) || 0;
        return sum + qty;
      }, 0);
      const repeatedLotsCount = exportedRows.filter(row => row["_isRepeated"]).length;

      // Garment Type Breakdown Map
      const garmentMap = {};
      const printingParties = {};
      let normalLotsCount = 0;
      let normalPcsCount = 0;
      let redLotsCount = 0;
      let redPcsCount = 0;

      exportedRows.forEach(row => {
        const garment = row["Garment Type"] || "Unassigned";
        const printingParty = row["Printing"] || "Unassigned";
        const qty = parseFloat(row["Challan Total Qty"]) || 0;
        const days = parseFloat(row["Days"]);

        if (!garmentMap[garment]) {
          garmentMap[garment] = { name: garment, totalLots: 0, totalPcs: 0 };
        }
        garmentMap[garment].totalLots += 1;
        garmentMap[garment].totalPcs += qty;

        if (!printingParties[printingParty]) {
          printingParties[printingParty] = { name: printingParty, totalLots: 0, totalPcs: 0 };
        }
        printingParties[printingParty].totalLots += 1;
        printingParties[printingParty].totalPcs += qty;

        if (!isNaN(days) && days > 5) {
          redLotsCount++;
          redPcsCount += qty;
        } else {
          normalLotsCount++;
          normalPcsCount += qty;
        }
      });

      const sortedGarments = Object.values(garmentMap).sort((a, b) => b.totalPcs - a.totalPcs);
      const sortedPrintingParties = Object.values(printingParties).sort((a, b) => b.totalPcs - a.totalPcs);

      const title = "PRINTING CHALLAN PRODUCTION REPORT";
      const now = new Date();
      const subtitle = `Report Date: ${now.toLocaleDateString()} ${now.toLocaleTimeString()} • Factory Suite Pro`;

      const pageW = doc.internal.pageSize.getWidth();
      const pageH = doc.internal.pageSize.getHeight();

      doc.setFontSize(20);
      doc.setTextColor(0, 0, 0); // Pure Black Heading
      doc.setFont('helvetica', 'bold');
      doc.text(title, pageW / 2, 38, { align: 'center' });

      doc.setFontSize(10);
      doc.setTextColor(50, 50, 50);
      doc.setFont('helvetica', 'normal');
      doc.text(subtitle, pageW / 2, 54, { align: 'center' });

      let currentY = 66;

      // Summary Bar
      doc.setFillColor(...COLOR.totalBg);
      doc.roundedRect(30, currentY - 5, pageW - 60, 24, 6, 6, 'F');
      doc.setDrawColor(191, 219, 254);
      doc.roundedRect(30, currentY - 5, pageW - 60, 24, 6, 6, 'D');

      doc.setFontSize(11);
      doc.setTextColor(0, 0, 0); // Pure Black
      doc.setFont('helvetica', 'bold');

      const summaryText = `Total Records: ${totalLots}   |   Total Pcs / Qty: ${totalPcs.toLocaleString()}   |   Garment Types: ${sortedGarments.length}   |   Printing Parties: ${sortedPrintingParties.length}   |   Repeated Lots: ${repeatedLotsCount}`;
      const summaryTextWidth = doc.getTextWidth(summaryText);
      const summaryX = (pageW - summaryTextWidth) / 2;

      doc.text(summaryText, summaryX, currentY + 11);
      doc.setFont('helvetica', 'normal');

      currentY += 26;

      const activeColumns = selectedCols.map(id => {
        const def = ALL_EXPORT_COLUMNS.find(c => c.id === id);
        return def || { id, label: id, key: id, baseWidth: 50 };
      });

      const columns = activeColumns.map(c => c.label);

      const body = exportedRows.map((row, index) => {
        const lotNum = row["Lot No."] ? String(row["Lot No."]).replace(/★\s*/, '').trim() : "";
        const lotHist = headRemarksMap[lotNum] || [];
        const latestHod = lotHist.length > 0 ? lotHist[lotHist.length - 1].text : "";

        return activeColumns.map(col => {
          if (col.id === "S. No") return String(index + 1);
          if (col.id === "Lot No.") return lotNum;
          if (col.id === "HOD Remarks") return latestHod || "";
          return String(row[col.key] || "");
        });
      });

      const daysColIdx = columns.indexOf("Days");
      const lotColIdx = columns.indexOf("Lot No.");

      const availableWidth = pageW - 30;
      const baseWidths = activeColumns.map(c => c.baseWidth || 50);
      const baseSum = baseWidths.reduce((a, b) => a + b, 0);
      const columnStyles = {};
      baseWidths.forEach((w, i) => {
        columnStyles[i] = {
          cellWidth: (w / baseSum) * availableWidth,
          halign: 'center'
        };
      });

      const starPositions = [];

      autoTable(doc, {
        head: [columns],
        body,
        startY: currentY,
        tableWidth: availableWidth,
        theme: "grid",
        styles: {
          fontSize: 9,
          cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
          overflow: "linebreak",
          valign: 'middle',
          textColor: [0, 0, 0], // Pure Black
          lineColor: [0, 0, 0], // Black grid lines
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

            if (isRepeated && data.column.index === lotColIdx && lotColIdx !== -1) {
              data.cell.styles.fillColor = COLOR.repeatedLotBg;
              data.cell.styles.fontStyle = "bold";
              data.cell.styles.fontSize = 10;
              data.cell.styles.textColor = [0, 0, 0];
            }

            if (data.column.index === daysColIdx && daysColIdx !== -1) {
              const rawVal = data.cell.raw !== undefined && data.cell.raw !== null ? String(data.cell.raw).trim() : "";
              const n = parseFloat(rawVal);
              if (!isNaN(n) && rawVal !== "") {
                if (n > 5) {
                  data.cell.styles.fillColor = [239, 68, 68];   // Vibrant Red
                  data.cell.styles.textColor = [255, 255, 255]; // White bold
                  data.cell.styles.fontStyle = "bold";
                } else {
                  data.cell.styles.fillColor = [220, 252, 231]; // Soft Green
                  data.cell.styles.textColor = [21, 128, 61];   // Green text
                  data.cell.styles.fontStyle = "bold";
                }
              }
            }
          }
        },

        willDrawCell: function (data) {
          if (data.section === "body" && data.column.index === lotColIdx && lotColIdx !== -1) {
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
          if (data.section === "body" && data.column.index === lotColIdx && lotColIdx !== -1) {
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
          doc.setTextColor(0, 0, 0);
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

      // --- 3-COLUMN SIDE-BY-SIDE EXECUTIVE SUMMARY ---
      const gBody = sortedGarments.map(item => {
        const pct = totalPcs > 0 ? ((item.totalPcs / totalPcs) * 100).toFixed(1) : "0.0";
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
        totalPcs.toLocaleString(),
        "100.0%"
      ]);

      const pBody = sortedPrintingParties.map(party => {
        const pct = totalPcs > 0 ? ((party.totalPcs / totalPcs) * 100).toFixed(1) : "0.0";
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
        totalPcs.toLocaleString(),
        "100.0%"
      ]);

      const aBody = [
        [
          "<= 5 Days (On-Time)",
          normalLotsCount.toString(),
          normalPcsCount.toLocaleString(),
          `${totalLots > 0 ? ((normalLotsCount / totalLots) * 100).toFixed(1) : 0}%`
        ],
        [
          "> 5 Days (Delayed)",
          redLotsCount.toString(),
          redPcsCount.toLocaleString(),
          `${totalLots > 0 ? ((redLotsCount / totalLots) * 100).toFixed(1) : 0}%`
        ],
        [
          "TOTAL",
          totalLots.toString(),
          totalPcs.toLocaleString(),
          "100.0%"
        ]
      ];

      const maxRows = Math.max(gBody.length, pBody.length, aBody.length);
      const approxSummaryHeight = 55 + (maxRows * 18);
      let lastAutoTable = doc.lastAutoTable;
      const mainFinalY = lastAutoTable ? lastAutoTable.finalY : 100;

      let summaryStartY;
      if (mainFinalY + approxSummaryHeight + 35 <= pageH - 30) {
        // Fits on the same page!
        summaryStartY = mainFinalY + 22;
        doc.setDrawColor(203, 213, 225);
        doc.setLineWidth(0.8);
        doc.line(20, summaryStartY - 8, pageW - 20, summaryStartY - 8);
      } else {
        // Create a new clean page
        doc.addPage();
        summaryStartY = 40;
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
      const summarySub = `Total Records: ${totalLots} Lots   |   Total Pieces: ${totalPcs.toLocaleString()} Pcs   |   Printing Parties: ${sortedPrintingParties.length}   |   Garment Types: ${sortedGarments.length}   |   Repeated Lots: ${repeatedLotsCount}`;
      doc.text(summarySub, pageW / 2, summaryStartY + 16, { align: 'center' });

      const sectionTitleY = summaryStartY + 30;
      const tableStartY = sectionTitleY + 6;

      // 3 Columns Side-by-Side Configuration
      const colWidth = 370;
      const col1X = 20;
      const col2X = 410;
      const col3X = 800;

      // Section Titles above each Column
      doc.setFontSize(10);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(0, 0, 0);
      doc.text("1. GARMENT TYPE BREAKDOWN", col1X, sectionTitleY);
      doc.text("2. PRINTING PARTY SUMMARY", col2X, sectionTitleY);
      doc.text("3. SLA & DAYS AGING", col3X, sectionTitleY);

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
          cellPadding: { top: 3.5, right: 3, bottom: 3.5, left: 3 },
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
          fontSize: 9,
          halign: 'center',
          cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
        },
        columnStyles: {
          0: { cellWidth: 145, halign: 'left' },
          1: { cellWidth: 60, halign: 'center' },
          2: { cellWidth: 100, halign: 'right' },
          3: { cellWidth: 65, halign: 'center' },
        },
        didParseCell: function (data) {
          if (data.section === 'body') {
            data.cell.styles.textColor = [0, 0, 0];
            if (data.row.index === gBody.length - 1) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });
      const endY1 = doc.lastAutoTable.finalY;

      // Column 2 Table: Printing Party Summary
      autoTable(doc, {
        head: [['Printing Party', 'Lots', 'Total Pcs', 'Share %']],
        body: pBody,
        startY: tableStartY,
        tableWidth: colWidth,
        margin: { left: col2X, right: pageW - (col2X + colWidth) },
        theme: "grid",
        styles: {
          fontSize: 8.5,
          cellPadding: { top: 3.5, right: 3, bottom: 3.5, left: 3 },
          overflow: "linebreak",
          valign: 'middle',
          halign: 'center',
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          lineWidth: 0.3,
        },
        headStyles: {
          fillColor: [30, 64, 175], // Royal Blue
          textColor: [255, 255, 255],
          fontStyle: "bold",
          fontSize: 9,
          halign: 'center',
          cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
        },
        columnStyles: {
          0: { cellWidth: 145, halign: 'left' },
          1: { cellWidth: 60, halign: 'center' },
          2: { cellWidth: 100, halign: 'right' },
          3: { cellWidth: 65, halign: 'center' },
        },
        didParseCell: function (data) {
          if (data.section === 'body') {
            data.cell.styles.textColor = [0, 0, 0];
            if (data.row.index === pBody.length - 1) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });
      const endY2 = doc.lastAutoTable.finalY;

      // Column 3 Table: SLA & Days Aging Breakdown
      autoTable(doc, {
        head: [['Aging Status', 'Lots', 'Total Pcs', 'Share %']],
        body: aBody,
        startY: tableStartY,
        tableWidth: colWidth,
        margin: { left: col3X, right: pageW - (col3X + colWidth) },
        theme: "grid",
        styles: {
          fontSize: 8.5,
          cellPadding: { top: 3.5, right: 3, bottom: 3.5, left: 3 },
          overflow: "linebreak",
          valign: 'middle',
          halign: 'center',
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          lineWidth: 0.3,
        },
        headStyles: {
          fillColor: [180, 83, 9], // Amber
          textColor: [255, 255, 255],
          fontStyle: "bold",
          fontSize: 9,
          halign: 'center',
          cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
        },
        columnStyles: {
          0: { cellWidth: 145, halign: 'left' },
          1: { cellWidth: 60, halign: 'center' },
          2: { cellWidth: 100, halign: 'right' },
          3: { cellWidth: 65, halign: 'center' },
        },
        didParseCell: function (data) {
          if (data.section === 'body') {
            data.cell.styles.textColor = [0, 0, 0];
            if (data.row.index === 0) {
              data.cell.styles.fillColor = [220, 252, 231]; // Soft Green
              data.cell.styles.fontStyle = 'bold';
            } else if (data.row.index === 1) {
              data.cell.styles.fillColor = [254, 226, 226]; // Soft Red
              data.cell.styles.fontStyle = 'bold';
            } else if (data.row.index === 2) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });
      const endY3 = doc.lastAutoTable.finalY;

      const maxEndY = Math.max(endY1, endY2, endY3);

      const finalY = maxEndY + 16;
      if (finalY <= pageH - 22) {
        doc.setDrawColor(0, 0, 0);
        doc.setLineWidth(0.5);
        doc.line(20, finalY, pageW - 20, finalY);

        doc.setFontSize(8.5);
        doc.setFont('helvetica', 'italic');
        doc.setTextColor(0, 0, 0);
        const thankYouText = "Factory Suite Pro • Report Generated Successfully";
        const thankYouWidth = doc.getTextWidth(thankYouText);
        const thankYouX = (pageW - thankYouWidth) / 2;
        doc.text(thankYouText, thankYouX, finalY + 11);
      }

      // Page Border & Footer for ALL Pages
      const totalPages = doc.internal.getNumberOfPages();
      for (let i = 1; i <= totalPages; i++) {
        doc.setPage(i);
        doc.setDrawColor(0, 0, 0);
        doc.setLineWidth(0.5);
        doc.roundedRect(10, 10, pageW - 20, pageH - 20, 2, 2, "S");

        doc.setFontSize(8.5);
        doc.setTextColor(0, 0, 0);
        doc.setFont('helvetica', 'normal');

        doc.text(`Factory Suite Pro • Printing Challan Report`, 18, pageH - 14);
        const label = `Page ${i} of ${totalPages}`;
        doc.text(label, pageW / 2, pageH - 14, { align: 'center' });
        const timeStr = `${new Date().toLocaleDateString()} ${new Date().toLocaleTimeString()}`;
        doc.text(`Generated: ${timeStr}`, pageW - 18, pageH - 14, { align: 'right' });
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
            <button
              style={styles.exportBtnExcel}
              onClick={() => setExportModalConfig({ isOpen: true, type: 'excel' })}
              title="Export to Excel"
            >
              📊 Excel
            </button>
            <button
              style={styles.exportBtnPdf}
              onClick={() => setExportModalConfig({ isOpen: true, type: 'pdf' })}
              title="Export to PDF"
            >
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
                seasonFilter.length > 0,
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
                        seasonFilter.length > 0,
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
                        setSeasonFilter([]);
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

      {/* Column Selection Export Modal */}
      <ExportColumnModal
        isOpen={exportModalConfig.isOpen}
        onClose={() => setExportModalConfig({ isOpen: false, type: 'excel' })}
        exportType={exportModalConfig.type}
        allColumns={ALL_EXPORT_COLUMNS}
        selectedColumns={selectedExportColumns}
        setSelectedColumns={setSelectedExportColumns}
        onConfirmExport={(cols) => {
          if (exportModalConfig.type === 'excel') {
            exportToExcelWithColumns(cols);
          } else {
            exportToPDFWithColumns(cols);
          }
          setExportModalConfig({ isOpen: false, type: 'excel' });
        }}
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