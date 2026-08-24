import React, { useState, useEffect, useMemo, useRef } from "react";
import { useHistory } from "react-router-dom";
import dayjs from "dayjs";
import isBetween from "dayjs/plugin/isBetween";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";

dayjs.extend(isBetween);

// Google Sheets Credentials from Environment Variables with fallbacks
const SPREADSHEET_ID =
  process.env.REACT_APP_KNITTING_SPREADSHEET_ID ||
  "1yHVieyNb7A5rds3oBEaUlfxxgG04QetLpb8T9g_xQPw";
const API_KEY = process.env.REACT_APP_GOOGLE_API_KEY || "";
const SHEET_RANGE =
  process.env.REACT_APP_YARN_SHEET_RANGE || "yarn!A1:C1000";

const DATA_HEADERS = ["Item Description", "Pkgs", "Bal WT (Kgs)"];

const MOCK_DATES = ["Wednesday, August 5, 2026"];

const YarnReport = () => {
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
    item: ""
  });

  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");

  // Fetch Data from Google Sheets API
  const fetchSheetData = async () => {
    setIsLoading(true);
    setError(null);

    try {
      const encodedRange = encodeURIComponent(SHEET_RANGE);
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodedRange}?key=${API_KEY}`;
      const response = await fetch(url);

      if (!response.ok) {
        throw new Error(`Google Sheets API returned status ${response.status}`);
      }

      const result = await response.json();
      const rows = result.values || [];

      if (rows.length === 0) {
        setError("No data found in the Yarn Stock spreadsheet range.");
        setIsLoading(false);
        return;
      }

      const parsedData = parseYarnRows(rows);
      const datesFound = Object.keys(parsedData);

      if (datesFound.length > 0) {
        setDataByDate(parsedData);
        setAllDates(datesFound);
        setSelectedDate(datesFound[0]);
      } else {
        setError("No valid date sections found in the yarn sheet.");
      }
    } catch (err) {
      console.error("Yarn Stock API Error:", err);
      setError(`Failed to fetch live Yarn Stock data: ${err.message}`);
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    fetchSheetData();
  }, []);

  // Parser for Yarn Stock rows with Date Blocks
  const parseYarnRows = (rows) => {
    const result = {};
    let currentDate = null;
    let inDataBlock = false;
    const defaultDate = "Wednesday, August 5, 2026";

    const dateRegex = /(monday|tuesday|wednesday|thursday|friday|saturday|sunday),\s+\w+\s+\d{1,2},\s+\d{4}/i;

    for (let i = 0; i < rows.length; i++) {
      const rawRow = rows[i] || [];
      if (!rawRow || rawRow.every((cell) => !cell || !String(cell).trim())) continue;

      const joinedRow = rawRow.map((c) => (c ? String(c).trim() : "")).join(" ").trim();

      // Check for Date Heading anywhere in the row
      const dateMatch = joinedRow.match(dateRegex);
      if (dateMatch) {
        currentDate = dateMatch[0];
        if (!result[currentDate]) {
          result[currentDate] = [];
        }
        inDataBlock = false;
        continue;
      }

      const col0 = String(rawRow[0] || "").trim();
      const col1 = String(rawRow[1] || "").trim();
      const col2 = String(rawRow[2] || "").trim();
      const itemUpper = col0.toUpperCase();

      // Detect Header Row (Item Description)
      if (itemUpper.includes("ITEM DESCRIPTION") || itemUpper.includes("ITEM")) {
        inDataBlock = true;
        if (!currentDate) {
          currentDate = defaultDate;
          if (!result[currentDate]) result[currentDate] = [];
        }
        continue;
      }

      if (inDataBlock) {
        // Skip sub-headers or title lines except Total
        if ((itemUpper === "YARN STOCK" || col0 === "") && itemUpper !== "TOTAL") {
          continue;
        }

        if (!currentDate) {
          currentDate = defaultDate;
          if (!result[currentDate]) result[currentDate] = [];
        }

        const record = {
          "Item Description": col0,
          "Pkgs": col1 !== "" ? col1 : "0",
          "Bal WT (Kgs)": col2 !== "" ? col2 : "0"
        };

        if (record["Item Description"] || record["Pkgs"] !== "0" || record["Bal WT (Kgs)"] !== "0") {
          result[currentDate].push(record);
        }
      }
    }

    // Fallback: If no date sections populated but rows exist
    if (Object.keys(result).length === 0 && rows.length > 0) {
      const fallbackRecords = [];
      rows.forEach((r) => {
        const item = String(r[0] || "").trim();
        const itemUpper = item.toUpperCase();
        if (item && !itemUpper.includes("KNITTING") && !itemUpper.includes("YARN STOCK") && !itemUpper.includes("ITEM DESCRIPTION")) {
          fallbackRecords.push({
            "Item Description": item,
            "Pkgs": String(r[1] || "0").trim(),
            "Bal WT (Kgs)": String(r[2] || "0").trim()
          });
        }
      });
      if (fallbackRecords.length > 0) {
        result[defaultDate] = fallbackRecords;
      }
    }

    return result;
  };

  // Unique filter options
  const filterOptions = useMemo(() => {
    const options = {
      items: new Set()
    };

    Object.entries(dataByDate).forEach(([_, records]) => {
      records.forEach((item) => {
        if (item["Item Description"]) options.items.add(item["Item Description"]);
      });
    });

    return {
      items: Array.from(options.items).filter(Boolean).sort()
    };
  }, [dataByDate]);

  // Filtered Data Display
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

      if (filters.item && row["Item Description"] !== filters.item) return false;

      return true;
    });
  }, [dataByDate, selectedDate, startDate, endDate, searchTerm, filters]);

  // Aggregated Totals
  const aggregatedTotals = useMemo(() => {
    let pkgsSum = 0;
    let balWtSum = 0;

    filteredData.forEach((row) => {
      const itemUpper = (row["Item Description"] || "").trim().toUpperCase();
      if (itemUpper === "TOTAL") return; // Avoid double counting

      const pkgs = parseFloat((row["Pkgs"] || "0").replace(/,/g, "")) || 0;
      const balWt = parseFloat((row["Bal WT (Kgs)"] || "0").replace(/,/g, "")) || 0;

      pkgsSum += pkgs;
      balWtSum += balWt;
    });

    return { pkgsSum, balWtSum };
  }, [filteredData]);

  // Navigation handlers
  const handleDateNavigation = (direction) => {
    if (!selectedDate || allDates.length === 0) return;
    const currentIndex = allDates.indexOf(selectedDate);
    if (direction === "prev" && currentIndex < allDates.length - 1) {
      setSelectedDate(allDates[currentIndex + 1]);
    } else if (direction === "next" && currentIndex > 0) {
      setSelectedDate(allDates[currentIndex - 1]);
    }
  };

  const clearFilters = () => {
    setSearchTerm("");
    setFilters({ item: "" });
    setStartDate("");
    setEndDate("");
  };

  const activeFilterCount = useMemo(() => {
    let count = 0;
    if (searchTerm) count++;
    if (filters.item) count++;
    if (startDate || endDate) count++;
    return count;
  }, [searchTerm, filters, startDate, endDate]);

  // Export to Excel CSV
  const exportCSV = () => {
    if (filteredData.length === 0) return;

    const headers = ["Item Description", "Pkgs", "Bal WT (Kgs)"];
    const csvRows = [headers.join(",")];

    filteredData.forEach((row) => {
      const values = headers.map((header) => {
        const val = row[header] || "";
        const escaped = ('' + val).replace(/"/g, '""');
        return `"${escaped}"`;
      });
      csvRows.push(values.join(","));
    });

    const blob = new Blob([csvRows.join("\n")], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `Yarn_Stock_Report_${selectedDate || "All"}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  // Strict Single-Page PDF Export with Yellow Total Row & Red Negative Values
  const exportPDF = () => {
    const doc = new jsPDF({
      orientation: 'portrait',
      unit: 'mm',
      format: 'a4'
    });

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(14);
    doc.setTextColor(30, 27, 75);
    doc.text('KNITTING DEPARTMENT SUMMARY - YARN STOCK', 105, 12, { align: 'center' });

    doc.setFontSize(10);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(100, 116, 139);
    doc.text(`Date: ${selectedDate || 'All Dates'} | Total Items: ${filteredData.length}`, 105, 17, { align: 'center' });

    const headers = [["ITEM DESCRIPTION", "PKGS", "BAL WT (KGS)"]];

    const body = filteredData.map((row) => [
      row["Item Description"] || "-",
      row["Pkgs"] || "0",
      row["Bal WT (Kgs)"] || "0"
    ]);

    const totalRows = body.length + 1;
    const availHeight = 255;
    let minCellHeight = Math.floor(availHeight / totalRows);

    if (minCellHeight > 10) minCellHeight = 10;
    if (minCellHeight < 5) minCellHeight = 5;

    let fontSize = 9;
    let headFontSize = 10;
    let cellPadding = 2;

    if (totalRows > 35) {
      fontSize = 7;
      headFontSize = 8;
      cellPadding = 1;
      minCellHeight = 4.5;
    } else if (totalRows > 25) {
      fontSize = 8;
      headFontSize = 9;
      cellPadding = 1.5;
      minCellHeight = 6;
    }

    autoTable(doc, {
      head: headers,
      body: body,
      startY: 22,
      margin: { top: 6, bottom: 6, left: 15, right: 15 },
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
      columnStyles: {
        0: { halign: 'left', cellWidth: 100 }
      },
      alternateRowStyles: {
        fillColor: [248, 250, 252]
      },
      didParseCell: function (data) {
        if (data.row.section === 'body') {
          const rowData = data.row.raw;
          const itemVal = (rowData[0] || '').toString().trim().toUpperCase();

          const isTotal = itemVal === "TOTAL";

          if (isTotal) {
            data.cell.styles.fillColor = [254, 240, 138];
            data.cell.styles.textColor = [0, 0, 0];
            data.cell.styles.fontStyle = 'bold';
            data.cell.styles.lineWidth = 0.3;
            data.cell.styles.lineColor = [0, 0, 0];
          }

          // Negative Value Handling in Red
          const pkgsVal = parseFloat((rowData[1] || '0').replace(/,/g, ''));
          const wtVal = parseFloat((rowData[2] || '0').replace(/,/g, ''));

          if (data.column.index === 1 && pkgsVal < 0) {
            data.cell.styles.textColor = [220, 38, 38];
            data.cell.styles.fontStyle = 'bold';
          }
          if (data.column.index === 2 && wtVal < 0) {
            data.cell.styles.textColor = [220, 38, 38];
            data.cell.styles.fontStyle = 'bold';
          }
        }
      }
    });

    doc.save(`Yarn_Stock_Report_${selectedDate || 'All'}.pdf`);
  };

  const styles = `
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    .yarn-container {
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
      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
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

    .text-negative-red {
      color: #dc2626 !important;
      font-weight: 800;
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
        <div className="yarn-container">
          <div className="loading-glass-card">
            <div className="spinner-ring" />
            <h3 style={{ fontSize: '1.25rem', color: '#0f172a', margin: '0 0 6px 0' }}>Fetching Yarn Stock Summary</h3>
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
        <div className="yarn-container">
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
      <div className="yarn-container" ref={reportRef}>
        {/* ====== HERO HEADER ====== */}
        <div className="page-header-wrapper">
          <div className="dashboard-header">
            <div className="header-bg-glow" />
            
            <div className="header-system-badge">
              <span style={{ display: 'inline-block', width: '7px', height: '7px', borderRadius: '50%', background: '#34d399', boxShadow: '0 0 8px #34d399' }} />
              <span>Garment Production Suite</span>
              <span style={{ color: 'rgba(255,255,255,0.4)' }}>|</span>
              <span style={{ color: '#e0e7ff' }}>YARN STOCK ANALYTICS</span>
            </div>

            <div className="header-top">
              <div className="header-title">
                <h1>Knitting Department Summary - Yarn Stock</h1>
                <p>Real-Time Yarn Inventory, Packages & Balance Weight (Kgs) Tracker</p>
              </div>

              {/* Glass KPI Cards */}
              <div className="header-kpi-grid">
                <div className="header-kpi-card">
                  <span className="header-kpi-value">{aggregatedTotals.pkgsSum.toFixed(0)}</span>
                  <span className="header-kpi-label">Total Pkgs</span>
                </div>
                <div className="header-kpi-card">
                  <span className="header-kpi-value" style={{ color: '#86efac' }}>{aggregatedTotals.balWtSum.toLocaleString()} kg</span>
                  <span className="header-kpi-label">Total Bal Wt</span>
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
                📅 {selectedDate || "All Yarn Stock Dates"}
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
                  placeholder="Search item description..."
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
                  <option value="">All Yarn Items</option>
                  {filterOptions.items.map(i => (
                    <option key={i} value={i}>{i}</option>
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
                  <th style={{ textAlign: 'left', paddingLeft: '24px' }}>ITEM DESCRIPTION</th>
                  <th>PKGS</th>
                  <th>BAL WT (KGS)</th>
                </tr>
              </thead>
              <tbody>
                {filteredData.length === 0 ? (
                  <tr>
                    <td colSpan={3} style={{ padding: '36px', color: '#64748b' }}>
                      No Yarn Stock records found for the selected date/filters.
                    </td>
                  </tr>
                ) : (
                  filteredData.map((row, idx) => {
                    const itemUpper = (row["Item Description"] || "").trim().toUpperCase();
                    const isTotalRow = itemUpper === "TOTAL";

                    const pkgsVal = parseFloat((row["Pkgs"] || "0").replace(/,/g, ""));
                    const wtVal = parseFloat((row["Bal WT (Kgs)"] || "0").replace(/,/g, ""));

                    return (
                      <tr key={idx} className={isTotalRow ? "total-row-yellowish" : ""}>
                        <td style={{ fontWeight: 700, textAlign: 'left', paddingLeft: '24px', color: isTotalRow ? '#0f172a' : '#1e293b' }}>
                          {row["Item Description"] || "-"}
                        </td>
                        <td className={pkgsVal < 0 ? "text-negative-red" : ""} style={{ fontWeight: 700 }}>
                          {row["Pkgs"] || "0"}
                        </td>
                        <td className={wtVal < 0 ? "text-negative-red" : ""} style={{ fontWeight: 700 }}>
                          {row["Bal WT (Kgs)"] || "0"}
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

export default YarnReport;
