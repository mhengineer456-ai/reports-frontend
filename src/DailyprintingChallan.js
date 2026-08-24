// src/components/DailyEmbroideryChallan.jsx
import React, { useEffect, useMemo, useState } from "react";

const SHEET_ID = "1fKSwGBIpzWEFk566WRQ4bzQ0anJlmasoY8TwrTLQHXI";
const API_KEY = "AIzaSyAomDFBkOySlIxKWSKGHe6ATv9gvaBr7uk";
const TAB_NAME = "JobOrder";
const RANGE = "A1:ZZZ";

const API_URL = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${TAB_NAME}!${RANGE}?key=${API_KEY}`;

// add two headers
const DAILY_HEADERS = [
  "Lot Number",
  "Challan No",
  "Issue Date",
  "Receive Date",            // from embUpdatedAt
  "Material Receive Date",   // ← NEW (from receivedDate)
  "Printing",
  "Shades (in challan)",     // ← NEW (CSV of all shades in this challan)
  "Challan Qty (date-wise)", // total per challan
];



function formatDisplayDate(d) {
  if (!d) return "";
  const parsed = new Date(d);
  if (isNaN(parsed)) return String(d); // e.g., "20 Sept 2025"
  return parsed.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
}

function safeSumItemsQty(items) {
  if (!Array.isArray(items)) return 0;
  return items.reduce((acc, it) => acc + (Number(it?.qty) || 0), 0);
}

// Download utilities
function downloadCSV(data, filename) {
  const headers = DAILY_HEADERS;
  const csvContent = [
    headers.join(","),
    ...data.map((row) =>
      headers
        .map((header) => {
          const value = row[header] || "";
          return `"${String(value).replace(/"/g, '""')}"`;
        })
        .join(",")
    ),
  ].join("\n");

  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const link = document.createElement("a");
  const url = URL.createObjectURL(blob);
  link.setAttribute("href", url);
  link.setAttribute("download", filename);
  link.style.visibility = "hidden";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

function downloadPDF(data) {
  const printWindow = window.open("", "_blank");
  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <title>Daily Printing Challan Report</title>
        <style>
          body { font-family: 'Inter', Arial, sans-serif; margin: 40px; color: #1a202c; }
          .header { text-align: center; margin-bottom: 40px; border-bottom: 2px solid #e2e8f0; padding-bottom: 20px; }
          .header h1 { color: #2d3748; margin-bottom: 8px; }
          table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 14px; }
          th, td { border: 1px solid #e2e8f0; padding: 12px; text-align: left; }
          th { background-color: #f7fafc; font-weight: 600; color: #4a5568; }
          .total { font-weight: 600; margin-top: 30px; padding: 16px; background: #f7fafc; border-radius: 8px; }
          .print-date { color: #718096; font-size: 14px; }
        </style>
      </head>
      <body>
        <div class="header">
          <h1>Daily Printing Challan Report</h1>
          <p class="print-date">Generated on: ${new Date().toLocaleDateString('en-GB', { 
            day: '2-digit', month: 'short', year: 'numeric'
          })} ${new Date().toLocaleTimeString([], {hour: '2-digit', minute: '2-digit'})}</p>
        </div>
        <table>
          <thead>
            <tr>
              ${DAILY_HEADERS.map((h) => `<th>${h}</th>`).join("")}
            </tr>
          </thead>
          <tbody>
            ${data
              .map(
                (row) => `
              <tr>
                ${DAILY_HEADERS.map((header) => `<td>${row[header] || ""}</td>`).join("")}
              </tr>
            `
              )
              .join("")}
          </tbody>
        </table>
        <div class="total">
          Total Records: ${data.length}
        </div>
      </body>
    </html>
  `;
  printWindow.document.write(htmlContent);
  printWindow.document.close();
  printWindow.print();
}

export default function DailyPrintingChallan() {
  const [rawRows, setRawRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [filterLot, setFilterLot] = useState("");
  const [filterIssueDate, setFilterIssueDate] = useState(""); // display date
  const [filterReceiveDate, setFilterReceiveDate] = useState(""); // display date
  const [filterPrinting, setFilterPrinting] = useState(""); // ✅ renamed for clarity
  const [searchTerm, setSearchTerm] = useState("");
  const [filterShade, setFilterShade] = useState("");   // ← NEW


  // ✅ Back navigation
  const goBack = () => {
    try {
      if (window.history.length > 1) {
        window.history.back();
        return;
      }
    } catch {}
    window.location.href = "/"; // fallback route — change to your dashboard if needed
  };

  useEffect(() => {
    (async () => {
      try {
        const res = await fetch(API_URL);
        if (!res.ok) throw new Error("Failed to fetch data");
        const data = await res.json();
        const [headers, ...rows] = data.values;

        const formatted = rows.map((row) => {
          const obj = {};
          headers.forEach((h, i) => {
            obj[h] = row[i] ?? "";
          });
          return obj;
        });

        // Only printing challans
        const printingOnly = formatted.filter((r) =>
          String(r["Challan No"] || "").trim().startsWith("CH-PRINT-")
        );

        setRawRows(printingOnly);
      } catch (e) {
        console.error(e);
        setError("Failed to fetch data.");
      } finally {
        setLoading(false);
      }
    })();
  }, []);

  // Unique values for filters
  const uniqueLots = useMemo(() => {
    const set = new Set(
      rawRows
        .map((r) => String(r["Lot Number"] || "").trim())
        .filter((v) => v.length > 0)
    );
    return Array.from(set).sort((a, b) =>
      a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" })
    );
  }, [rawRows]);

  // ✅ Use "Printing" column (was "Emb")
  const uniquePrintings = useMemo(() => {
    const set = new Set(
      rawRows
        .map((r) => String(r["Printing"] || "").trim())
        .filter((v) => v.length > 0)
    );
    return Array.from(set).sort();
  }, [rawRows]);

  // Expand each row's history to per-challan entries
  // helper to read Shade if your sheet also contains a master Shade column (fallback)
const getRowShade = (r) =>
  String(r["Shade"] || r["Shade No"] || r["Shade Number"] || r["Color"] || "").trim();
const toShadesCSV = (items) =>
  Array.from(
    new Set(
      (Array.isArray(items) ? items : [])
        .map(it => String(it?.shade || "").trim())
        .filter(Boolean)
    )
  ).join(", ");

function sumItemsQty(items) {
  return (Array.isArray(items) ? items : []).reduce((s, it) => s + (Number(it?.qty) || 0), 0);
}


const expandedRows = useMemo(() => {
  const out = [];
  for (const r of rawRows) {
    const lotNo = r["Lot Number"] || "";
    const printing = r["Printing"] || "";
    const challanHistoryStr = r["Challan History JSON"] || "";

    let history = [];
    try {
      const parsed = JSON.parse(challanHistoryStr);
      if (Array.isArray(parsed)) history = parsed;
    } catch {}

    if (history.length === 0) continue;

    for (const entry of history) {
      const challanNo = entry?.number || r["Challan No"] || "";
      const issueDate = entry?.date || r["Challan Date"] || "";
      const receiveDate = entry?.embUpdatedAt || "";
      const materialReceiveDate = entry?.receivedDate || "";

      const items = Array.isArray(entry?.items) ? entry.items : [];
      const totalQty = Number(entry?.totalQty) || sumItemsQty(items);
      const shadesCSV = toShadesCSV(items); // e.g. "BLACK, 15%, D.GREY, ..."

      out.push({
        "Lot Number": lotNo,
        "Challan No": challanNo,
        "Issue Date": formatDisplayDate(issueDate),
        "Receive Date": formatDisplayDate(receiveDate),
        "Material Receive Date": formatDisplayDate(materialReceiveDate), // ← NEW
        Printing: printing,
        "Shades (in challan)": shadesCSV,                                // ← NEW
        "Challan Qty (date-wise)": totalQty,

        // raw fields for sorting only
        rawIssueDate: issueDate,
        rawReceiveDate: receiveDate,
      });
    }
  }

  // newest first by raw Issue Date
  out.sort((a, b) => {
    const da = new Date(a.rawIssueDate);
    const db = new Date(b.rawReceiveDate); // keep your original? likely issue date sort:
    // If you want by issue date:
    const d1 = new Date(a.rawIssueDate), d2 = new Date(b.rawIssueDate);
    if (isNaN(d1) && isNaN(d2)) return 0;
    if (isNaN(d1)) return 1;
    if (isNaN(d2)) return -1;
    return d2 - d1;  // descending
  });

  return out;
}, [rawRows]);

  // Build dropdown options for dates from expandedRows (use DISPLAY values)
  const uniqueIssueDates = useMemo(() => {
    const set = new Set(
      expandedRows.map((r) => r["Issue Date"]).filter((v) => v && v.length > 0)
    );
    return Array.from(set).sort((a, b) => {
      const da = new Date(a);
      const db = new Date(b);
      if (isNaN(da) || isNaN(db)) return a.localeCompare(b);
      return da - db;
    });
  }, [expandedRows]);

  const uniqueReceiveDates = useMemo(() => {
    const set = new Set(
      expandedRows.map((r) => r["Receive Date"]).filter((v) => v && v.length > 0)
    );
    return Array.from(set).sort((a, b) => {
      const da = new Date(a);
      const db = new Date(b);
      if (isNaN(da) || isNaN(db)) return a.localeCompare(b);
      return da - db;
    });
  }, [expandedRows]);

  // Apply filters
  const filteredRows = useMemo(() => {
    let filtered = expandedRows;

    if (filterLot) filtered = filtered.filter((r) => r["Lot Number"] === filterLot);
    if (filterIssueDate) filtered = filtered.filter((r) => r["Issue Date"] === filterIssueDate);
    if (filterReceiveDate)
      filtered = filtered.filter((r) => r["Receive Date"] === filterReceiveDate);

    // ✅ filter by "Printing" column (was r["Emb"])
    if (filterPrinting) filtered = filtered.filter((r) => r["Printing"] === filterPrinting);

    if (searchTerm) {
      const term = searchTerm.toLowerCase();
      filtered = filtered.filter((r) =>
        Object.values(r).some((value) => String(value).toLowerCase().includes(term))
      );
    }

    return filtered;
  }, [expandedRows, filterLot, filterIssueDate, filterReceiveDate, filterPrinting, searchTerm]);

  const clearFilters = () => {
    setFilterLot("");
    setFilterIssueDate("");
    setFilterReceiveDate("");
    setFilterPrinting("");
    setSearchTerm("");
  };

  const hasActiveFilters =
    filterLot || filterIssueDate || filterReceiveDate || filterPrinting || searchTerm;

  return (
    <div className="modern-container">
      {/* Header Section */}
      <div className="header-section">
        <div className="header-content">
          <div className="title-group">
            <h1 className="main-title">Daily Printing Challan</h1>
            <p className="subtitle">Per-Challan Overview & Management</p>
          </div>
          <div className="header-actions">
            <button className="back-btn" onClick={goBack} title="Go back">
              ← Back
            </button>
            <button
              className="export-btn excel-btn"
              onClick={() =>
                downloadCSV(
                  filteredRows,
                  `printing-challan-${new Date().toISOString().split("T")[0]}.csv`
                )
              }
            >
              <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                <polyline points="14,2 14,8 20,8"/>
                <path d="M16 13H8"/>
                <path d="M16 17H8"/>
                <path d="M10 9H8"/>
              </svg>
              Export Excel
            </button>
            <button className="export-btn pdf-btn" onClick={() => downloadPDF(filteredRows)}>
              <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                <polyline points="14,2 14,8 20,8"/>
                <path d="M16 13H8"/>
                <path d="M16 17H8"/>
                <path d="M10 9H8"/>
              </svg>
              Export PDF
            </button>
          </div>
        </div>
      </div>

      {/* Stats Cards */}
      <div className="stats-container">
        <div className="stat-card">
          <div className="stat-icon total">
            <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
              <path d="M9 17H7v-7h2v7zm4 0h-2V7h2v10zm4 0h-2v-4h2v4zm2 2H5V5h14v14zm0-16H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2z"/>
            </svg>
          </div>
          <div className="stat-info">
            <span className="stat-value">{expandedRows.length}</span>
            <span className="stat-label">Total Records</span>
          </div>
        </div>
        <div className="stat-card">
          <div className="stat-icon filtered">
            <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
              <path d="M10 18h4v-2h-4v2zM3 6v2h18V6H3zm3 7h12v-2H6v2z"/>
            </svg>
          </div>
          <div className="stat-info">
            <span className="stat-value">{filteredRows.length}</span>
            <span className="stat-label">Filtered Results</span>
          </div>
        </div>
        <div className="stat-card">
          <div className="stat-icon lots">
            <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
              <path d="M17 3H7c-1.1 0-1.99.9-1.99 2L5 21l7-3 7 3V5c0-1.1-.9-2-2-2z"/>
            </svg>
          </div>
          <div className="stat-info">
            <span className="stat-value">{uniqueLots.length}</span>
            <span className="stat-label">Unique Lots</span>
          </div>
        </div>
      </div>

      {/* Search and Filters */}
      <div className="filters-card">
        <div className="search-section">
          <div className="search-wrapper">
            <svg className="search-icon" width="18" height="18" viewBox="0 0 24 24" fill="currentColor">
              <path d="M15.5 14h-.79l-.28-.27C15.41 12.59 16 11.11 16 9.5 16 5.91 13.09 3 9.5 3S3 5.91 3 9.5 5.91 16 9.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14z"/>
            </svg>
            <input
              type="text"
              placeholder="Search across all columns..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="search-input"
            />
          </div>
        </div>

        <div className="filters-section">
          <div className="filter-row">
            <div className="filter-group">
              <label className="filter-label">Lot Number</label>
              <select
                value={filterLot}
                onChange={(e) => setFilterLot(e.target.value)}
                className="filter-select"
              >
                <option value="">All Lots</option>
                {uniqueLots.map((lot) => (
                  <option key={lot} value={lot}>
                    {lot}
                  </option>
                ))}
              </select>
            </div>

            <div className="filter-group">
              <label className="filter-label">Issue Date</label>
              <select
                value={filterIssueDate}
                onChange={(e) => setFilterIssueDate(e.target.value)}
                className="filter-select"
              >
                <option value="">All Issue Dates</option>
                {uniqueIssueDates.map((date) => (
                  <option key={date} value={date}>
                    {date}
                  </option>
                ))}
              </select>
            </div>

            <div className="filter-group">
              <label className="filter-label">Receive Date</label>
              <select
                value={filterReceiveDate}
                onChange={(e) => setFilterReceiveDate(e.target.value)}
                className="filter-select"
              >
                <option value="">All Receive Dates</option>
                {uniqueReceiveDates.map((date) => (
                  <option key={date} value={date}>
                    {date}
                  </option>
                ))}
              </select>
            </div>

            <div className="filter-group">
              <label className="filter-label">Printing</label>
              <select
                value={filterPrinting}
                onChange={(e) => setFilterPrinting(e.target.value)}
                className="filter-select"
              >
                <option value="">All PRINTING</option>
                {uniquePrintings.map((p) => (
                  <option key={p} value={p}>
                    {p}
                  </option>
                ))}
              </select>
            </div>
          </div>

          {hasActiveFilters && (
            <div className="filter-actions">
              <button onClick={clearFilters} className="clear-filters-btn">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/>
                </svg>
                Clear Filters
              </button>
              <span className="results-info">
                Showing {filteredRows.length} of {expandedRows.length} records
              </span>
            </div>
          )}
        </div>
      </div>

      {/* Data Table */}
      <div className="table-card">
        {loading && (
          <div className="loading-state">
            <div className="loading-spinner"></div>
            <p>Loading Printing challan data...</p>
          </div>
        )}

        {error && (
          <div className="error-state">
            <svg width="24" height="24" viewBox="0 0 24 24" fill="currentColor">
              <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/>
            </svg>
            <p>{error}</p>
          </div>
        )}

        {!loading && !error && (
          <>
            <div className="table-container">
              <table className="modern-table">
                <thead>
                  <tr>
                    {DAILY_HEADERS.map((h) => (
                      <th key={h}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filteredRows.map((row, idx) => (
                    <tr key={idx}>
                      {DAILY_HEADERS.map((h) => (
                        <td key={h}>
                          <span className="cell-content">{row[h]}</span>
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {filteredRows.length === 0 && (
              <div className="empty-state">
                <div className="empty-icon">
                  <svg width="48" height="48" viewBox="0 0 24 24" fill="currentColor">
                    <path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-5 14H7v-2h7v2zm3-4H7v-2h10v2zm0-4H7V7h10v2z"/>
                  </svg>
                </div>
                <h3>No challan entries found</h3>
                <p>Try adjusting your filters or search criteria</p>
                {hasActiveFilters && (
                  <button onClick={clearFilters} className="primary-btn">
                    Clear All Filters
                  </button>
                )}
              </div>
            )}
          </>
        )}
      </div>

      <style jsx>{`
        .modern-container {
          min-height: 100vh;
          background: #ffffffff;
          padding: 24px;
          font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        }
        .header-section {  background: #4331a8ff;color: white border-radius: 12px; padding: 24px; margin-bottom: 24px; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1); border: 1px solid #e2e8f0; }
        .header-content { display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; }
        .title-group { flex: 1; }
        .main-title { font-size: 24px; font-weight: 700; color: #ffffffff; margin: 0 0 4px 0; }
        .subtitle { font-size: 14px; color: #ffffffff; margin: 0; }
        .header-actions { display: flex; gap: 10px; flex-wrap: wrap; }
        .back-btn {
          display: inline-flex;
          align-items: center;
          gap: 6px;
          padding: 10px 14px;
          border: 1px solid #e5e7eb;
          border-radius: 8px;
          background: #111827;
          color: #ffffff;
          font-weight: 600;
          font-size: 14px;
          cursor: pointer;
          transition: transform .15s ease, box-shadow .15s ease, opacity .15s ease;
        }
        .back-btn:hover { transform: translateY(-1px); box-shadow: 0 2px 8px rgba(0,0,0,0.15); }
        .export-btn { display: flex; align-items: center; gap: 8px; padding: 10px 16px; border: 1px solid #e2e8f0; border-radius: 8px; background: white; color: #4a5568; font-weight: 500; font-size: 14px; cursor: pointer; transition: all 0.2s ease; }
        .export-btn:hover { transform: translateY(-1px); box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1); }
        .excel-btn:hover { border-color: #10b981; color: #10b981; }
        .pdf-btn:hover { border-color: #ef4444; color: #ef4444; }
        .stats-container { display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 16px; margin-bottom: 24px; }
        .stat-card { background: white; padding: 20px; border-radius: 12px; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1); border: 1px solid #e2e8f0; display: flex; align-items: center; gap: 16px; }
        .stat-icon { width: 48px; height: 48px; border-radius: 10px; display: flex; align-items: center; justify-content: center; color: white; }
        .stat-icon.total { background: linear-gradient(135deg, #667eea, #764ba2); }
        .stat-icon.filtered { background: linear-gradient(135deg, #f093fb, #f5576c); }
        .stat-icon.lots { background: linear-gradient(135deg, #4facfe, #00f2fe); }
        .stat-value { display: block; font-size: 24px; font-weight: 700; color: #1a202c; }
        .stat-label { font-size: 14px; color: #718096; }
        .filters-card { background: white; border-radius: 12px; padding: 24px; margin-bottom: 24px; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1); border: 1px solid #e2e8f0; }
        .search-section { margin-bottom: 20px; }
        .search-wrapper { position: relative; max-width: 400px; }
        .search-icon { position: absolute; left: 12px; top: 50%; transform: translateY(-50%); color: #a0aec0; }
        .search-input { width: 100%; padding: 12px 12px 12px 40px; border: 1px solid #e2e8f0; border-radius: 8px; font-size: 14px; transition: all 0.2s ease; }
        .search-input:focus { outline: none; border-color: #4299e1; box-shadow: 0 0 0 3px rgba(66, 153, 225, 0.1); }
        .filter-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; }
        .filter-group { display: flex; flex-direction: column; gap: 6px; }
        .filter-label { font-size: 12px; font-weight: 600; color: #4a5568; text-transform: uppercase; letter-spacing: 0.5px; }
        .filter-select, .filter-input { padding: 10px 12px; border: 1px solid #e2e8f0; border-radius: 6px; font-size: 14px; background: white; transition: all 0.2s ease; }
        .filter-select:focus, .filter-input:focus { outline: none; border-color: #4299e1; box-shadow: 0 0 0 3px rgba(66, 153, 225, 0.1); }
        .filter-actions { display: flex; justify-content: space-between; align-items: center; margin-top: 16px; padding-top: 16px; border-top: 1px solid #e2e8f0; }
        .clear-filters-btn { display: flex; align-items: center; gap: 6px; padding: 8px 12px; border: 1px solid #fed7d7; border-radius: 6px; background: #fff5f5; color: #c53030; font-size: 12px; font-weight: 500; cursor: pointer; transition: all 0.2s ease; }
        .clear-filters-btn:hover { background: #fed7d7; }
        .results-info { font-size: 12px; color: #718096; font-weight: 500; }
        .table-card { background: white; border-radius: 12px; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1); border: 1px solid #e2e8f0; overflow: hidden; }
        .table-container { overflow-x: auto; }
        .modern-table { border-collapse: collapse; width: 100%; }
        .modern-table th {
          background: #4331a8ff;
          padding: 16px;
          text-align: center;
          font-weight: 600;
          color: #ffffffff;
          border: 1px solid #e2e8f0;
          font-size: 12px;
          text-transform: uppercase;
          letter-spacing: 0.5px;
        }
        .modern-table td {
          padding: 16px;
          border: 1px solid #e2e8f0;
          color: #000000ff;
          text-align: center;
        }
        .modern-table tr:last-child td { border-bottom: none; }
        .modern-table tr:hover { background: #f7fafc; }
        .cell-content { display: inline-flex; align-items: center; min-height: 20px; }
        .loading-state { padding: 60px 20px; text-align: center; color: #718096; }
        .loading-spinner { border: 3px solid #f3f3f3; border-top: 3px solid #4299e1; border-radius: 50%; width: 40px; height: 40px; animation: spin 1s linear infinite; margin: 0 auto 16px; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        .error-state { padding: 40px 20px; text-align: center; color: #e53e3e; background: #fed7d7; margin: 20px; border-radius: 8px; }
        .error-state svg { margin-bottom: 12px; }
        .empty-state { padding: 60px 20px; text-align: center; color: #718096; }
        .empty-icon { margin-bottom: 16px; color: #cbd5e0; }
        .empty-state h3 { margin: 0 0 8px 0; color: #4a5568; font-size: 18px; }
        .empty-state p { margin: 0 0 20px 0; font-size: 14px; }
        .primary-btn { padding: 10px 20px; background: #4299e1; color: white; border: none; border-radius: 6px; font-weight: 500; cursor: pointer; transition: background 0.2s ease; }
        .primary-btn:hover { background: #3182ce; }
        @media (max-width: 768px) {
          .modern-container { padding: 16px; }
          .header-content { flex-direction: column; gap: 16px; }
          .header-actions { width: 100%; justify-content: stretch; }
          .back-btn, .export-btn { flex: 1; justify-content: center; }
          .filter-row { grid-template-columns: 1fr; }
          .stats-container { grid-template-columns: 1fr; }
          .filter-actions { flex-direction: column; gap: 12px; align-items: stretch; }
        }
      `}</style>
    </div>
  );
}
