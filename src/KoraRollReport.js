import React, { useState, useEffect, useMemo, useRef } from 'react';
import { useHistory } from "react-router-dom";
import axios from "axios";
import { saveAs } from "file-saver";
import dayjs from "dayjs";
import isBetweenPlugin from "dayjs/plugin/isBetween";
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';

// Extend dayjs with plugins
dayjs.extend(isBetweenPlugin);

// Credentials & Configuration
const API_KEY = process.env.REACT_APP_GOOGLE_API_KEY || "AIzaSyAomDFBkOySlIxKWSKGHe6ATv9gvaBr7uk";
const SPREADSHEET_ID = process.env.REACT_APP_KNITTING_SPREADSHEET_ID || "1yHVieyNb7A5rds3oBEaUlfxxgG04QetLpb8T9g_xQPw";
const SHEET_RANGE = process.env.REACT_APP_KORA_ROLLS_SHEET_RANGE || "kora rolls!A1:H1000";

// Columns in each data row:
const DATA_HEADERS = [
  "Sr. No.",
  "Item Description",
  "OP Rolls",
  "Prod Rolls",
  "Issue Rolls",
  "Bal Rolls",
  "Issue to Whom"
];

const MOCK_DATES = ["Wednesday, August 5, 2026"];

const KoraRollReport = () => {
  const history = useHistory();
  const reportRef = useRef(null);

  // States
  const [dataByDate, setDataByDate] = useState({});
  const [allDates, setAllDates] = useState(MOCK_DATES);
  const [selectedDate, setSelectedDate] = useState(MOCK_DATES[0]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState(null);
  const [searchTerm, setSearchTerm] = useState("");
  const [showExportMenu, setShowExportMenu] = useState(false);

  // Filters
  const [filters, setFilters] = useState({
    issueToWhom: "",
    item: ""
  });
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");

  // ---------------------- FETCH DATA ----------------------
  const fetchSheetData = async () => {
    setIsLoading(true);
    setError(null);

    try {
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(SHEET_RANGE)}?key=${API_KEY}`;
      const response = await axios.get(url);
      const rows = response.data.values || [];

      const parsed = parseSheetRows(rows);
      setDataByDate(parsed);

      const datesFound = Object.keys(parsed).sort((a, b) => {
        return dayjs(b, "dddd, MMMM D, YYYY").diff(
          dayjs(a, "dddd, MMMM D, YYYY")
        );
      });

      setAllDates(datesFound.length > 0 ? datesFound : MOCK_DATES);
      if (datesFound.length > 0) {
        setSelectedDate(datesFound[0]);
      } else {
        setSelectedDate(null);
      }
    } catch (err) {
      console.error("Error fetching Kora Rolls data:", err);
      const apiMsg = err.response?.data?.error?.message;
      setError(apiMsg ? `Google Sheets API Error: ${apiMsg}` : "Failed to load Kora Rolls data. Please check connection.");
      setDataByDate({});
      setAllDates(MOCK_DATES);
      setSelectedDate(MOCK_DATES[0]);
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    fetchSheetData();
  }, []);

  // ---------------------- PARSE ROWS ----------------------
  const parseSheetRows = (rows) => {
    const result = {};
    let currentDate = "";
    let inDataBlock = false;
    let headerOffset = 0;

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.every((cell) => !cell.trim())) continue;

      const joinedRow = row.join(" ").trim();
      const rowString = joinedRow.toLowerCase();

      // Check for Date Header (e.g. Wednesday, August 5, 2026)
      const dateRegex = /(monday|tuesday|wednesday|thursday|friday|saturday|sunday),\s+\w+\s+\d{1,2},\s+\d{4}/i;
      const containsDate = dateRegex.test(joinedRow);

      if (containsDate) {
        let finalDate = joinedRow;
        if (rowString.includes("knitting department summary")) {
          finalDate = finalDate.replace(/knitting department summary\s*/i, "").trim();
        }
        currentDate = finalDate;
        result[currentDate] = [];
        inDataBlock = false;
        continue;
      }

      // Check for Table Header row
      const srIdx = row.findIndex(c => c.toLowerCase().includes("sr. no.") || c.toLowerCase().includes("sr no"));
      if (srIdx !== -1) {
        headerOffset = srIdx;
        inDataBlock = true;
        continue;
      }

      if (inDataBlock && currentDate) {
        const srNo = row[headerOffset] || "";
        const item = row[headerOffset + 1] || "";

        // Skip section subheaders like "KORA ROLLS" header row
        if (item.toLowerCase() === "kora rolls" && !row[headerOffset + 2]) continue;

        const record = {
          "Sr. No.": srNo,
          "Item Description": item,
          "OP Rolls": row[headerOffset + 2] || "",
          "Prod Rolls": row[headerOffset + 3] || "",
          "Issue Rolls": row[headerOffset + 4] || "",
          "Bal Rolls": row[headerOffset + 5] || "",
          "Issue to Whom": row[headerOffset + 6] || ""
        };
        result[currentDate].push(record);
      }
    }

    return result;
  };

  // Unique filter options
  const filterOptions = useMemo(() => {
    const options = {
      issueToWhom: new Set(),
      items: new Set()
    };

    Object.entries(dataByDate).forEach(([_, records]) => {
      records.forEach((item) => {
        if (item["Issue to Whom"]) options.issueToWhom.add(item["Issue to Whom"]);
        if (item["Item Description"]) options.items.add(item["Item Description"]);
      });
    });

    return {
      issueToWhom: Array.from(options.issueToWhom).filter(Boolean).sort(),
      items: Array.from(options.items).filter(v => v && v !== "Total").sort()
    };
  }, [dataByDate]);

  // Filtered Data (No Pagination)
  const filteredData = useMemo(() => {
    let combined = [];
    const allDateKeys = Object.keys(dataByDate);
    let datesToInclude = [];

    if (selectedDate === null) {
      datesToInclude = allDateKeys;
    } else if (startDate && endDate) {
      datesToInclude = allDateKeys.filter((dateStr) => {
        const dateObj = dayjs(dateStr, "dddd, MMMM D, YYYY");
        return dateObj.isBetween(dayjs(startDate), dayjs(endDate), "day", "[]");
      });
    } else if (selectedDate) {
      datesToInclude = [selectedDate];
    }

    if (datesToInclude.length === 0 && allDateKeys.length > 0) {
      datesToInclude = allDateKeys;
    }

    datesToInclude.forEach((date) => {
      if (dataByDate[date]) {
        combined = combined.concat(dataByDate[date]);
      }
    });

    return combined.filter((row) => {
      if (searchTerm) {
        const term = searchTerm.toLowerCase();
        const matches = Object.values(row).some((val) =>
          val.toLowerCase().includes(term)
        );
        if (!matches) return false;
      }

      if (filters.issueToWhom && row["Issue to Whom"] !== filters.issueToWhom) return false;
      if (filters.item && row["Item Description"] !== filters.item) return false;

      return true;
    });
  }, [dataByDate, selectedDate, startDate, endDate, searchTerm, filters]);

  // Aggregated Totals for non-summary rows
  const aggregatedTotals = useMemo(() => {
    let opSum = 0;
    let prodSum = 0;
    let issueSum = 0;
    let balSum = 0;

    filteredData.forEach((row) => {
      const itemUpper = (row["Item Description"] || "").trim().toUpperCase();
      const srUpper = (row["Sr. No."] || "").trim().toUpperCase();
      const isTotalRow = itemUpper === "TOTAL" || srUpper === "TOTAL";

      if (!isTotalRow) {
        opSum += parseFloat(row["OP Rolls"] || 0);
        prodSum += parseFloat(row["Prod Rolls"] || 0);
        issueSum += parseFloat(row["Issue Rolls"] || 0);
        balSum += parseFloat(row["Bal Rolls"] || 0);
      }
    });

    return {
      opSum,
      prodSum,
      issueSum,
      balSum
    };
  }, [filteredData]);

  // Navigation handlers
  const handleDateNavigation = (direction) => {
    if (selectedDate === null) return;
    const idx = allDates.indexOf(selectedDate);
    if (direction === "prev" && idx < allDates.length - 1) {
      setSelectedDate(allDates[idx + 1]);
    } else if (direction === "next" && idx > 0) {
      setSelectedDate(allDates[idx - 1]);
    }
  };

  const clearFilters = () => {
    setFilters({ issueToWhom: "", item: "" });
    setStartDate("");
    setEndDate("");
    setSearchTerm("");
  };

  const activeFilterCount = [
    filters.issueToWhom,
    filters.item,
    startDate,
    endDate,
    searchTerm
  ].filter(Boolean).length;

  // Export CSV
  const exportCSV = () => {
    const csvHeaders = DATA_HEADERS.join(",") + "\n";
    const csvRows = filteredData.map((item) =>
      DATA_HEADERS.map((hdr) => `"${item[hdr] || ''}"`).join(",")
    );
    const csvContent = csvHeaders + csvRows.join("\n");
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    saveAs(blob, `Kora_Rolls_Report_${selectedDate || 'All'}.csv`);
  };

  // Export PDF (Single Page Fit with Bold Text & Yellow Totals)
  const exportPDF = () => {
    const doc = new jsPDF('landscape');

    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(6, 78, 59);
    doc.text("KORA ROLLS PRODUCTION & INVENTORY REPORT", 10, 14);

    doc.setFontSize(9);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(71, 85, 105);
    doc.text(`Period: ${selectedDate || 'All Data'} | Total Records: ${filteredData.length}`, 10, 20);
    doc.text(`Generated: ${new Date().toLocaleString('en-IN')}`, 287, 20, { align: 'right' });

    const headers = [DATA_HEADERS];
    const body = filteredData.map(row => DATA_HEADERS.map(h => row[h] || '-'));

    // Dynamic single-page layout scaling
    const totalRows = filteredData.length + 1;
    const availHeight = 180;
    const maxRowHeight = availHeight / totalRows;

    let fontSize = 8.5;
    let headFontSize = 9.5;
    let cellPadding = 2;
    let minCellHeight = 5;

    if (maxRowHeight >= 10) {
      fontSize = 9.5;
      headFontSize = 10.5;
      cellPadding = 3;
      minCellHeight = 7;
    } else if (maxRowHeight >= 8) {
      fontSize = 8.5;
      headFontSize = 9.5;
      cellPadding = 2.2;
      minCellHeight = 5.5;
    } else if (maxRowHeight >= 6) {
      fontSize = 7.5;
      headFontSize = 8.5;
      cellPadding = 1.5;
      minCellHeight = 4.5;
    } else if (maxRowHeight >= 4.5) {
      fontSize = 6.5;
      headFontSize = 7.5;
      cellPadding = 1.0;
      minCellHeight = 3.5;
    } else {
      fontSize = 5.5;
      headFontSize = 6.5;
      cellPadding = 0.6;
      minCellHeight = 2.8;
    }

    autoTable(doc, {
      head: headers,
      body: body,
      startY: 22,
      margin: { top: 6, bottom: 6, left: 10, right: 10 },
      pageBreak: 'avoid',
      theme: 'grid',
      styles: {
        fontSize: fontSize,
        halign: 'center',
        valign: 'middle',
        fontStyle: 'bold',
        textColor: [0, 0, 0],
        lineColor: [0, 0, 0],
        lineWidth: 0.2,
        cellPadding: cellPadding,
        minCellHeight: minCellHeight
      },
      headStyles: {
        fillColor: [30, 27, 75],
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        fontSize: headFontSize,
        lineWidth: 0.2,
        lineColor: [30, 27, 75],
        minCellHeight: minCellHeight + 1
      },
      alternateRowStyles: {
        fillColor: [248, 250, 252]
      },
      didParseCell: function (data) {
        if (data.row.section === 'body') {
          const rowData = data.row.raw;
          const srVal = (rowData[0] || '').toString().trim().toUpperCase();
          const itemVal = (rowData[1] || '').toString().trim().toUpperCase();

          const isTotal = itemVal === "TOTAL" || srVal === "TOTAL";

          if (isTotal) {
            data.cell.styles.fillColor = [254, 240, 138];
            data.cell.styles.textColor = [0, 0, 0];
            data.cell.styles.fontStyle = 'bold';
            data.cell.styles.lineWidth = 0.3;
            data.cell.styles.lineColor = [0, 0, 0];
          }

          // Prod Rolls (Col Index 3) Green Text
          if (data.column.index === 3 && rowData[3] && parseFloat(rowData[3]) > 0) {
            data.cell.styles.textColor = [22, 163, 74];
            data.cell.styles.fontStyle = 'bold';
          }

          // Issue Rolls (Col Index 4) Red Text
          if (data.column.index === 4 && rowData[4] && parseFloat(rowData[4]) > 0) {
            data.cell.styles.textColor = [220, 38, 38];
            data.cell.styles.fontStyle = 'bold';
          }
        }
      }
    });

    doc.save(`Kora_Rolls_Report_${selectedDate || 'All'}.pdf`);
  };

  const styles = `
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    .kora-container {
      font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
      background-color: #f8fafc;
      background-image: radial-gradient(at 0% 0%, rgba(99, 102, 241, 0.08) 0px, transparent 50%),
                        radial-gradient(at 100% 0%, rgba(236, 72, 153, 0.06) 0px, transparent 50%),
                        radial-gradient(at 50% 100%, rgba(16, 185, 129, 0.06) 0px, transparent 50%);
      min-height: 100vh;
      padding: 0;
      color: #0f172a;
    }

    .page-header-wrapper {
      width: 100%;
      padding: 20px 28px 0;
      box-sizing: border-box;
    }

    .dashboard-header {
      width: 100%;
      background: linear-gradient(135deg, #1e1b4b 0%, #312e81 40%, #4338ca 100%);
      border-radius: 20px;
      padding: 22px 28px;
      box-shadow: 0 15px 35px -8px rgba(49, 46, 129, 0.25);
      color: #ffffff;
      position: relative;
      overflow: visible;
      z-index: 10;
      box-sizing: border-box;
    }

    .header-bg-glow {
      position: absolute;
      inset: 0;
      border-radius: 20px;
      overflow: hidden;
      pointer-events: none;
    }

    .header-system-badge {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 4px 14px;
      border-radius: 9999px;
      background: rgba(255,255,255,0.15);
      backdrop-filter: blur(12px);
      font-size: 0.8rem;
      font-weight: 600;
      color: #e0e7ff;
      border: 1px solid rgba(255,255,255,0.2);
      margin-bottom: 10px;
    }

    .header-top {
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 20px;
      flex-wrap: wrap;
      position: relative;
      z-index: 1;
    }

    .header-title h1 {
      font-size: 1.6rem;
      font-weight: 800;
      color: #ffffff;
      margin: 0 0 4px 0;
      letter-spacing: -0.02em;
    }

    .header-title p {
      color: #c7d2fe;
      font-size: 0.88rem;
      margin: 0;
      font-weight: 500;
    }

    .header-kpi-grid {
      display: flex;
      gap: 10px;
      align-items: center;
    }

    .header-kpi-card {
      background: rgba(255,255,255,0.12);
      border: 1px solid rgba(255,255,255,0.2);
      backdrop-filter: blur(12px);
      border-radius: 14px;
      padding: 8px 16px;
      text-align: center;
      min-width: 90px;
    }

    .header-kpi-value {
      font-size: 1.3rem;
      font-weight: 800;
      color: #ffffff;
      display: block;
      line-height: 1.1;
      margin-bottom: 2px;
    }

    .header-kpi-label {
      font-size: 0.65rem;
      font-weight: 700;
      color: #c7d2fe;
      text-transform: uppercase;
      letter-spacing: 0.04em;
    }

    .header-actions {
      display: flex;
      gap: 10px;
      align-items: center;
      margin-top: 14px;
      padding-top: 14px;
      border-top: 1px solid rgba(255,255,255,0.15);
      position: relative;
      z-index: 1;
    }

    .btn-header {
      padding: 8px 16px;
      background: rgba(255,255,255,0.15);
      color: #ffffff;
      border: 1.2px solid rgba(255,255,255,0.25);
      border-radius: 12px;
      font-size: 13px;
      font-weight: 700;
      cursor: pointer;
      backdrop-filter: blur(8px);
      transition: all 0.2s;
      display: inline-flex;
      align-items: center;
      gap: 6px;
    }

    .btn-header:hover {
      background: rgba(255,255,255,0.25);
      transform: translateY(-1px);
    }

    .export-menu {
      position: absolute;
      top: calc(100% + 6px);
      right: 0;
      background: #ffffff;
      border-radius: 12px;
      box-shadow: 0 15px 30px rgba(0, 0, 0, 0.2);
      padding: 6px;
      min-width: 160px;
      z-index: 99999;
    }

    .export-option {
      padding: 8px 14px;
      border: none;
      background: none;
      width: 100%;
      text-align: left;
      color: #334155;
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
      border-radius: 6px;
      display: flex;
      align-items: center;
      gap: 6px;
    }

    .export-option:hover {
      background: #f1f5f9;
      color: #4338ca;
    }

    .page-body {
      padding: 20px 28px;
      box-sizing: border-box;
    }

    /* Date Switcher Bar */
    .date-switcher-card {
      background: #ffffff;
      border-radius: 16px;
      padding: 12px 20px;
      margin-bottom: 16px;
      box-shadow: 0 6px 20px rgba(0,0,0,0.03);
      border: 1px solid #e2e8f0;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 14px;
      flex-wrap: wrap;
    }

    .date-nav-group {
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .nav-arrow-btn {
      width: 34px;
      height: 34px;
      border-radius: 10px;
      border: 1.2px solid #cbd5e1;
      background: #ffffff;
      color: #334155;
      font-size: 14px;
      font-weight: 700;
      cursor: pointer;
      display: flex;
      align-items: center;
      justify-content: center;
      transition: all 0.2s;
    }

    .nav-arrow-btn:hover:not(:disabled) {
      border-color: #4338ca;
      color: #4338ca;
      background: #e0e7ff;
    }

    .nav-arrow-btn:disabled {
      opacity: 0.35;
      cursor: not-allowed;
    }

    .current-date-pill {
      padding: 8px 16px;
      background: linear-gradient(135deg, #1e1b4b 0%, #4338ca 100%);
      color: #ffffff;
      border-radius: 12px;
      font-weight: 700;
      font-size: 0.9rem;
    }

    .all-dates-btn {
      padding: 8px 14px;
      border: 1.2px solid #cbd5e1;
      border-radius: 10px;
      background: #f8fafc;
      color: #475569;
      font-size: 12.5px;
      font-weight: 700;
      cursor: pointer;
    }

    .all-dates-btn.active {
      background: #4338ca;
      color: white;
      border-color: #312e81;
    }

    /* Filter Controls Bar */
    .filter-section {
      background: #ffffff;
      border-radius: 16px;
      padding: 14px 20px;
      margin-bottom: 18px;
      box-shadow: 0 6px 20px rgba(0,0,0,0.03);
      border: 1px solid #e2e8f0;
    }

    .filter-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
      gap: 12px;
    }

    .filter-control {
      display: flex;
      flex-direction: column;
      gap: 4px;
    }

    .filter-label {
      font-size: 11px;
      font-weight: 700;
      color: #64748b;
      text-transform: uppercase;
      letter-spacing: 0.04em;
    }

    .filter-select, .filter-input {
      padding: 8px 12px;
      border: 1.2px solid #cbd5e1;
      border-radius: 10px;
      font-size: 13px;
      font-weight: 600;
      color: #1e293b;
      background: #f8fafc;
      outline: none;
    }

    .filter-select:focus, .filter-input:focus {
      border-color: #4338ca;
      background: #ffffff;
    }

    /* Full-Screen Table Styling */
    .table-wrapper {
      background: #ffffff;
      border-radius: 16px;
      overflow: hidden;
      box-shadow: 0 6px 24px rgba(0,0,0,0.03);
      border: 1.2px solid #e2e8f0;
    }

    .full-report-table {
      width: 100%;
      border-collapse: collapse;
      text-align: center;
      font-size: 13.5px;
    }

    .full-report-table th {
      background: linear-gradient(135deg, #1e1b4b 0%, #4338ca 100%);
      color: #ffffff;
      font-weight: 700;
      padding: 11px 16px;
      font-size: 12.5px;
      text-transform: uppercase;
      letter-spacing: 0.04em;
      border-bottom: 2px solid #312e81;
    }

    .full-report-table td {
      padding: 10px 16px;
      border-bottom: 1px solid #f1f5f9;
      color: #334155;
      font-weight: 600;
    }

    .full-report-table tr:nth-child(even) td {
      background: #f8fafc;
    }

    .full-report-table tr:hover td {
      background: #e0e7ff;
    }

    .full-report-table tr.total-row-yellowish td {
      background: #fef08a !important;
      color: #0f172a !important;
      font-weight: 800 !important;
      font-size: 14px !important;
      border-top: 1.5px solid #eab308 !important;
      border-bottom: 1.5px solid #eab308 !important;
    }

    .text-green-bold {
      color: #16a34a !important;
      font-weight: 800;
    }

    .text-red-bold {
      color: #dc2626 !important;
      font-weight: 800;
    }

    .issue-badge {
      display: inline-block;
      padding: 3px 12px;
      border-radius: 9999px;
      background: #f3e8ff;
      color: #7e22ce;
      font-weight: 700;
      font-size: 12px;
      border: 1px solid #d8b4fe;
    }

    .loading-glass-card {
      max-width: 440px;
      margin: 60px auto;
      background: white;
      border-radius: 20px;
      padding: 40px;
      text-align: center;
      box-shadow: 0 10px 30px rgba(0,0,0,0.04);
      border: 1px solid #e2e8f0;
    }

    .spinner-ring {
      width: 48px;
      height: 48px;
      border: 3.5px solid #e2e8f0;
      border-top: 3.5px solid #4338ca;
      border-radius: 50%;
      animation: spin 0.8s linear infinite;
      margin: 0 auto 18px;
    }

    @keyframes spin {
      to { transform: rotate(360deg); }
    }
  `;

  if (isLoading) {
    return (
      <>
        <style>{styles}</style>
        <div className="kora-container">
          <div className="loading-glass-card">
            <div className="spinner-ring" />
            <h3 style={{ fontSize: '1.25rem', color: '#0f172a', margin: '0 0 6px 0' }}>Fetching Kora Rolls Summary</h3>
            <p style={{ color: '#64748b', margin: 0, fontSize: '0.9rem' }}>Connecting to Google Sheets live data...</p>
          </div>
        </div>
      </>
    );
  }

  if (error) {
    return (
      <>
        <style>{styles}</style>
        <div className="kora-container">
          <div className="loading-glass-card">
            <div style={{ fontSize: '44px', marginBottom: '14px' }}>⚠️</div>
            <h3 style={{ color: '#1e293b', margin: '0 0 6px 0' }}>Data Sync Error</h3>
            <p style={{ color: '#64748b', marginBottom: '18px', fontSize: '0.88rem' }}>{error}</p>
            <button className="btn-header" style={{ background: '#4338ca', margin: '0 auto' }} onClick={fetchSheetData}>
              🔄 Retry Sync
            </button>
          </div>
        </div>
      </>
    );
  }

  return (
    <>
      <style>{styles}</style>
      <div className="kora-container" ref={reportRef}>
        {/* ====== HERO HEADER ====== */}
        <div className="page-header-wrapper">
          <div className="dashboard-header">
            <div className="header-bg-glow" />
            
            <div className="header-system-badge">
              <span style={{ display: 'inline-block', width: '7px', height: '7px', borderRadius: '50%', background: '#34d399', boxShadow: '0 0 8px #34d399' }} />
              <span>Garment Production Suite</span>
              <span style={{ color: 'rgba(255,255,255,0.4)' }}>|</span>
              <span style={{ color: '#a7f3d0' }}>KORA ROLL INVENTORY ANALYTICS</span>
            </div>

            <div className="header-top">
              <div className="header-title">
                <h1>Kora Roll Production & Inventory</h1>
                <p>Real-Time Fabric Roll Balances, Issuance & Production Tracker</p>
              </div>

              {/* Glass KPI Cards */}
              <div className="header-kpi-grid">
                <div className="header-kpi-card">
                  <span className="header-kpi-value">{aggregatedTotals.opSum.toFixed(0)}</span>
                  <span className="header-kpi-label">OP Rolls</span>
                </div>
                <div className="header-kpi-card">
                  <span className="header-kpi-value" style={{ color: '#86efac' }}>{aggregatedTotals.prodSum.toFixed(0)}</span>
                  <span className="header-kpi-label">Prod Rolls</span>
                </div>
                <div className="header-kpi-card">
                  <span className="header-kpi-value" style={{ color: '#fca5a5' }}>{aggregatedTotals.issueSum.toFixed(0)}</span>
                  <span className="header-kpi-label">Issue Rolls</span>
                </div>
                <div className="header-kpi-card">
                  <span className="header-kpi-value" style={{ color: '#fde68a' }}>{aggregatedTotals.balSum.toFixed(0)}</span>
                  <span className="header-kpi-label">Bal Rolls</span>
                </div>
              </div>
            </div>

            {/* Action Bar */}
            <div className="header-actions">
              <button className="btn-header" onClick={() => history.goBack()}>
                ← Back
              </button>
              <button className="btn-header" onClick={fetchSheetData}>
                🔄 Refresh
              </button>
              <div style={{ position: 'relative' }}>
                <button className="btn-header" onClick={() => setShowExportMenu(!showExportMenu)}>
                  📥 Export Options
                </button>
                {showExportMenu && (
                  <div className="export-menu">
                    <button className="export-option" onClick={() => { exportPDF(); setShowExportMenu(false); }}>
                      📄 PDF Report
                    </button>
                    <button className="export-option" onClick={() => { exportCSV(); setShowExportMenu(false); }}>
                      📊 Excel CSV
                    </button>
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>

        {/* ====== BODY SECTION ====== */}
        <div className="page-body">
          {/* Date Switcher Bar */}
          <div className="date-switcher-card">
            <div className="date-nav-group">
              <button
                className="nav-arrow-btn"
                onClick={() => handleDateNavigation("prev")}
                disabled={selectedDate === null || allDates.indexOf(selectedDate) >= allDates.length - 1}
                title="Previous Day"
              >
                ◀
              </button>
              <div className="current-date-pill">
                📅 {selectedDate || "All Kora Rolls Dates"}
              </div>
              <button
                className="nav-arrow-btn"
                onClick={() => handleDateNavigation("next")}
                disabled={selectedDate === null || allDates.indexOf(selectedDate) <= 0}
                title="Next Day"
              >
                ▶
              </button>
            </div>

            <div style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
              <button
                className={`all-dates-btn ${selectedDate === null ? 'active' : ''}`}
                onClick={() => setSelectedDate(null)}
              >
                Show All Dates ({allDates.length} Days)
              </button>
              <span style={{ fontSize: '13px', color: '#64748b', fontWeight: 600 }}>
                Showing {filteredData.length} records
              </span>
            </div>
          </div>

          {/* Filter Controls Bar */}
          <div className="filter-section">
            <div className="filter-grid">
              <div className="filter-control">
                <label className="filter-label">Search</label>
                <input
                  type="text"
                  className="filter-input"
                  placeholder="Search item, processor..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                />
              </div>

              <div className="filter-control">
                <label className="filter-label">Item Description</label>
                <select
                  className="filter-select"
                  value={filters.item}
                  onChange={(e) => setFilters(prev => ({ ...prev, item: e.target.value }))}
                >
                  <option value="">All Fabrics</option>
                  {filterOptions.items.map(i => (
                    <option key={i} value={i}>{i}</option>
                  ))}
                </select>
              </div>

              <div className="filter-control">
                <label className="filter-label">Issue to Whom</label>
                <select
                  className="filter-select"
                  value={filters.issueToWhom}
                  onChange={(e) => setFilters(prev => ({ ...prev, issueToWhom: e.target.value }))}
                >
                  <option value="">All Processors / Parties</option>
                  {filterOptions.issueToWhom.map(p => (
                    <option key={p} value={p}>{p}</option>
                  ))}
                </select>
              </div>

              {activeFilterCount > 0 && (
                <div className="filter-control" style={{ justifyContent: 'flex-end' }}>
                  <button className="all-dates-btn" style={{ background: '#ef4444', color: 'white', borderColor: '#dc2626' }} onClick={clearFilters}>
                    Clear Filters ({activeFilterCount})
                  </button>
                </div>
              )}
            </div>
          </div>

          {/* Full Screen Unpaginated Table */}
          <div className="table-wrapper">
            <table className="full-report-table">
              <thead>
                <tr>
                  <th>SR. NO.</th>
                  <th>ITEM DESCRIPTION</th>
                  <th>OP ROLLS</th>
                  <th>PROD ROLLS</th>
                  <th>ISSUE ROLLS</th>
                  <th>BAL ROLLS</th>
                  <th>ISSUE TO WHOM</th>
                </tr>
              </thead>
              <tbody>
                {filteredData.length === 0 ? (
                  <tr>
                    <td colSpan={7} style={{ padding: '36px', color: '#64748b' }}>
                      No Kora Roll records found for the selected date/filters.
                    </td>
                  </tr>
                ) : (
                  filteredData.map((row, idx) => {
                    const itemUpper = (row["Item Description"] || "").trim().toUpperCase();
                    const srUpper = (row["Sr. No."] || "").trim().toUpperCase();
                    const isTotalRow = itemUpper === "TOTAL" || srUpper === "TOTAL";

                    const prodRolls = parseFloat(row["Prod Rolls"] || 0);
                    const issueRolls = parseFloat(row["Issue Rolls"] || 0);
                    const issueToWhom = row["Issue to Whom"];

                    return (
                      <tr key={idx} className={isTotalRow ? "total-row-yellowish" : ""}>
                        <td style={{ fontWeight: 700, color: isTotalRow ? '#0f172a' : '#1e293b' }}>{row["Sr. No."] || "-"}</td>
                        <td style={{ fontWeight: 700, textAlign: 'left', paddingLeft: '24px' }}>{row["Item Description"] || "-"}</td>
                        <td>{row["OP Rolls"] || "0"}</td>
                        <td className={prodRolls > 0 ? "text-green-bold" : ""}>
                          {row["Prod Rolls"] || "-"}
                        </td>
                        <td className={issueRolls > 0 ? "text-red-bold" : ""}>
                          {row["Issue Rolls"] || "-"}
                        </td>
                        <td style={{ fontWeight: 800, color: '#0f172a' }}>{row["Bal Rolls"] || "0"}</td>
                        <td>
                          {issueToWhom ? (
                            <span className="issue-badge">{issueToWhom}</span>
                          ) : (
                            "-"
                          )}
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </>
  );
};

export default KoraRollReport;
