// src/CancelledLotsReport.js
import React, { useState, useEffect, useMemo, useRef } from "react";
import { useHistory } from "react-router-dom";
import * as XLSX from "xlsx-js-style";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import { SPREADSHEET_IDS, fetchSheetDataFromBackend, BACKEND_URL } from "./config";

/** Google Drive direct image URL helper */
const getDirectImageUrl = (url) => {
  if (!url) return "";
  const reg = /id=([a-zA-Z0-9_-]+)/;
  const match = String(url).match(reg);
  if (match && match[1]) {
    return `https://lh3.googleusercontent.com/d/${match[1]}`;
  }
  const dreg = /\/d\/([a-zA-Z0-9_-]+)/;
  const dmatch = String(url).match(dreg);
  if (dmatch && dmatch[1]) {
    return `https://lh3.googleusercontent.com/d/${dmatch[1]}`;
  }
  return url;
};

/** Date format helper */
function formatDate(dateStr) {
  if (!dateStr) return "-";
  const str = String(dateStr).trim();
  if (str.includes("T")) {
    const d = new Date(str);
    if (!isNaN(d.getTime())) {
      return d.toLocaleDateString("en-IN", {
        day: "2-digit",
        month: "short",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit"
      });
    }
  }
  const parts = str.split("-");
  if (parts.length === 3) {
    const [y, m, d] = parts.map((p) => parseInt(p, 10));
    if (y && m && d) {
      const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      return `${d} ${monthNames[m - 1] || m} ${y}`;
    }
  }
  return str;
}

/** Financial Year helper */
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

// Multi-select dropdown component with Vanilla CSS
const MultiSelectDropdown = ({ label, options, selectedValues, onChange, placeholder = "Select..." }) => {
  const [isOpen, setIsOpen] = useState(false);
  const dropdownRef = useRef(null);

  useEffect(() => {
    const handleClickOutside = (e) => {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target)) {
        setIsOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  const toggleOption = (val) => {
    if (selectedValues.includes(val)) {
      onChange(selectedValues.filter((v) => v !== val));
    } else {
      onChange([...selectedValues, val]);
    }
  };

  return (
    <div className="clr-dropdown-wrapper" ref={dropdownRef}>
      <label className="clr-field-label">{label}</label>
      <button
        type="button"
        onClick={() => setIsOpen(!isOpen)}
        className="clr-select-trigger"
      >
        <span className="clr-select-trigger-text">
          {selectedValues.length === 0
            ? placeholder
            : `${selectedValues.length} selected`}
        </span>
        <span style={{ fontSize: "10px", color: "#64748b" }}>{isOpen ? "▲" : "▼"}</span>
      </button>

      {isOpen && (
        <div className="clr-dropdown-popover">
          <div className="clr-popover-header">
            <span>{options.length} Options</span>
            {selectedValues.length > 0 && (
              <button
                type="button"
                onClick={() => onChange([])}
                className="clr-clear-btn"
              >
                Clear
              </button>
            )}
          </div>
          <div className="clr-popover-options">
            {options.length === 0 ? (
              <div style={{ padding: "8px", color: "#94a3b8", textAlign: "center" }}>No options available</div>
            ) : (
              options.map((opt) => (
                <label key={opt} className="clr-popover-item">
                  <input
                    type="checkbox"
                    checked={selectedValues.includes(opt)}
                    onChange={() => toggleOption(opt)}
                    style={{ marginRight: "6px", cursor: "pointer" }}
                  />
                  <span>{opt}</span>
                </label>
              ))
            )}
          </div>
        </div>
      )}
    </div>
  );
};

export default function CancelledLotsReport() {
  const history = useHistory();
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [refreshing, setRefreshing] = useState(false);
  const [error, setError] = useState("");

  // Filters State
  const [search, setSearch] = useState("");
  const [financialYear, setFinancialYear] = useState(getCurrentFinancialYear());
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [selectedParties, setSelectedParties] = useState([]);
  const [selectedBrands, setSelectedBrands] = useState([]);
  const [selectedFabrics, setSelectedFabrics] = useState([]);
  const [selectedStyles, setSelectedStyles] = useState([]);
  const [selectedReasons, setSelectedReasons] = useState([]);
  const [selectedApprovedBy, setSelectedApprovedBy] = useState([]);

  // Pagination & Modals
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(25);
  const [activeLotDetail, setActiveLotDetail] = useState(null);
  const [lightboxImage, setLightboxImage] = useState(null);

  // Fetch only Cancelled Lots
  const fetchCancelledLots = async (isRefresh = false) => {
    if (isRefresh) setRefreshing(true);
    else setLoading(true);
    setError("");

    try {
      // Step 1: Try dedicated backend endpoint
      const backendRes = await fetch(
        `${BACKEND_URL}/api/sheets/cancelled-lots?refresh=${isRefresh}`
      );

      if (backendRes.ok) {
        const json = await backendRes.json();
        if (json.success && Array.isArray(json.data)) {
          setData(json.data);
          return;
        }
      }

      // Step 2: Fallback to direct JobOrder sheet read
      console.warn("Using fallback sheet fetch for JobOrder cancelled lots...");
      const sheetRes = await fetchSheetDataFromBackend(
        SPREADSHEET_IDS.JOBORDER,
        "JobOrder!A1:AZ50000"
      );

      if (sheetRes.ok && Array.isArray(sheetRes.values) && sheetRes.values.length > 0) {
        const rows = sheetRes.values;
        const headers = rows[0].map((h) => String(h || "").trim());
        const normalize = (s) => String(s || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");

        const headerMap = {};
        headers.forEach((h, idx) => {
          headerMap[normalize(h)] = idx;
        });

        const getVal = (row, key) => {
          const idx = headerMap[normalize(key)];
          return idx !== undefined && row[idx] !== undefined ? String(row[idx]).trim() : "";
        };

        const statusIdx = headerMap["status"];
        const cancelled = [];

        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length === 0) continue;

          const statusVal = statusIdx !== undefined ? String(row[statusIdx] || "").trim() : "";
          if (statusVal.toLowerCase().includes("cancel")) {
            cancelled.push({
              rowIndex: i + 1,
              jobOrderNo: getVal(row, "Job Order No") || getVal(row, "JobOrderNo") || getVal(row, "Order No"),
              date: getVal(row, "Date"),
              fabric: getVal(row, "Fabric"),
              brand: getVal(row, "Brand"),
              shade: getVal(row, "Shade"),
              size: getVal(row, "Size"),
              quantity: getVal(row, "Quantity"),
              unit: getVal(row, "Unit"),
              partyName: getVal(row, "Party Name") || getVal(row, "Party"),
              garmentType: getVal(row, "Garment Type") || getVal(row, "Garment"),
              section: getVal(row, "Section"),
              season: getVal(row, "Season"),
              emb: getVal(row, "Emb"),
              embDetails: getVal(row, "Emb Details"),
              printing: getVal(row, "Printing"),
              printingDetails: getVal(row, "Printing Details"),
              pattern: getVal(row, "Pattern"),
              style: getVal(row, "Style"),
              remarks: getVal(row, "Remarks"),
              directStitching: getVal(row, "Direct Stitching"),
              submittedBy: getVal(row, "Submitted By"),
              imageUrl: getVal(row, "Image URL") || getVal(row, "Image"),
              lotNumber: getVal(row, "Lot Number") || getVal(row, "Lot No") || getVal(row, "Lot"),
              component: getVal(row, "Component"),
              challanNo: getVal(row, "Challan No"),
              challanDate: getVal(row, "Challan Date"),
              challanItemsJson: getVal(row, "Challan Items JSON"),
              challanHistoryJson: getVal(row, "Challan History JSON"),
              challanTotalQty: getVal(row, "Challan Total Qty"),
              challanCompleteLot: getVal(row, "Challan Complete Lot"),
              challanBy: getVal(row, "Challan By"),
              challanPdfUrl: getVal(row, "Challan PDF URL"),
              priority: getVal(row, "Priority"),
              cancellationTimestamp: getVal(row, "Cancellation Timestamp") || getVal(row, "Cancelled At"),
              cancelledBy: getVal(row, "Cancelled By"),
              cancellationApprovedFrom: getVal(row, "Cancellation Approved From") || getVal(row, "Approved By"),
              cancellationReason: getVal(row, "Cancellation Reason") || getVal(row, "Reason"),
              status: statusVal || "Cancel"
            });
          }
        }

        setData(cancelled);
      } else {
        throw new Error("Unable to retrieve JobOrder sheet data");
      }
    } catch (err) {
      console.error("Error fetching cancelled lots:", err);
      setError(err.message || "Failed to load cancelled lots report");
    } finally {
      setLoading(false);
      setRefreshing(false);
    }
  };

  useEffect(() => {
    fetchCancelledLots();
  }, []);

  // Distinct Options for Dropdown Filters
  const distinctParties = useMemo(() => {
    return Array.from(new Set(data.map((d) => d.partyName).filter(Boolean))).sort();
  }, [data]);

  const distinctBrands = useMemo(() => {
    return Array.from(new Set(data.map((d) => d.brand).filter(Boolean))).sort();
  }, [data]);

  const distinctFabrics = useMemo(() => {
    return Array.from(new Set(data.map((d) => d.fabric).filter(Boolean))).sort();
  }, [data]);

  const distinctStyles = useMemo(() => {
    return Array.from(new Set(data.map((d) => d.style).filter(Boolean))).sort();
  }, [data]);

  const distinctReasons = useMemo(() => {
    return Array.from(new Set(data.map((d) => d.cancellationReason).filter(Boolean))).sort();
  }, [data]);

  const distinctApprovedBy = useMemo(() => {
    return Array.from(new Set(data.map((d) => d.cancellationApprovedFrom).filter(Boolean))).sort();
  }, [data]);

  const availableFYs = useMemo(() => {
    const set = new Set();
    const currentFY = getCurrentFinancialYear();
    if (currentFY) set.add(currentFY);
    data.forEach((r) => {
      const fy = getFinancialYearFromDate(r.date || r.cancellationTimestamp);
      if (fy) set.add(fy);
    });
    return Array.from(set).sort().reverse();
  }, [data]);

  // Filtering Logic
  const filteredData = useMemo(() => {
    return data.filter((item) => {
      // Global Search
      if (search.trim()) {
        const q = search.toLowerCase();
        const searchTarget = [
          item.lotNumber,
          item.jobOrderNo,
          item.partyName,
          item.brand,
          item.style,
          item.fabric,
          item.garmentType,
          item.cancellationReason,
          item.cancelledBy,
          item.cancellationApprovedFrom,
          item.remarks
        ]
          .join(" ")
          .toLowerCase();

        if (!searchTarget.includes(q)) return false;
      }

      // Financial Year Filter
      if (financialYear) {
        const itemFY = getFinancialYearFromDate(item.date || item.cancellationTimestamp);
        if (itemFY && itemFY !== financialYear) return false;
      }

      // Date Range Filter
      const itemDate = item.date || (item.cancellationTimestamp ? item.cancellationTimestamp.slice(0, 10) : "");
      if (startDate && itemDate && itemDate < startDate) return false;
      if (endDate && itemDate && itemDate > endDate) return false;

      // Multi-select filters
      if (selectedParties.length > 0 && !selectedParties.includes(item.partyName)) return false;
      if (selectedBrands.length > 0 && !selectedBrands.includes(item.brand)) return false;
      if (selectedFabrics.length > 0 && !selectedFabrics.includes(item.fabric)) return false;
      if (selectedStyles.length > 0 && !selectedStyles.includes(item.style)) return false;
      if (selectedReasons.length > 0 && !selectedReasons.includes(item.cancellationReason)) return false;
      if (selectedApprovedBy.length > 0 && !selectedApprovedBy.includes(item.cancellationApprovedFrom)) return false;

      return true;
    });
  }, [
    data,
    search,
    financialYear,
    startDate,
    endDate,
    selectedParties,
    selectedBrands,
    selectedFabrics,
    selectedStyles,
    selectedReasons,
    selectedApprovedBy
  ]);

  // Aggregate Metrics
  const stats = useMemo(() => {
    const totalLots = filteredData.length;
    let totalQty = 0;
    const partiesSet = new Set();
    const brandsSet = new Set();
    const reasonCounts = {};

    filteredData.forEach((item) => {
      const q = parseFloat(String(item.quantity || "0").replace(/,/g, ""));
      if (!isNaN(q)) totalQty += q;
      if (item.partyName) partiesSet.add(item.partyName);
      if (item.brand) brandsSet.add(item.brand);

      const r = item.cancellationReason || "Not Specified";
      reasonCounts[r] = (reasonCounts[r] || 0) + 1;
    });

    let topReason = "None";
    let maxCount = 0;
    Object.entries(reasonCounts).forEach(([r, count]) => {
      if (count > maxCount) {
        maxCount = count;
        topReason = `${r} (${count})`;
      }
    });

    return {
      totalLots,
      totalQty,
      uniqueParties: partiesSet.size,
      uniqueBrands: brandsSet.size,
      topReason
    };
  }, [filteredData]);

  // Clear All Filters
  const clearFilters = () => {
    setSearch("");
    setStartDate("");
    setEndDate("");
    setSelectedParties([]);
    setSelectedBrands([]);
    setSelectedFabrics([]);
    setSelectedStyles([]);
    setSelectedReasons([]);
    setSelectedApprovedBy([]);
  };

  // Pagination logic
  const totalPages = Math.ceil(filteredData.length / pageSize) || 1;
  const paginatedData = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return filteredData.slice(start, start + pageSize);
  }, [filteredData, currentPage, pageSize]);

  useEffect(() => {
    setCurrentPage(1);
  }, [search, financialYear, startDate, endDate, selectedParties, selectedBrands, selectedFabrics, selectedStyles, selectedReasons, selectedApprovedBy, pageSize]);

  // Export to Excel
  const exportToExcel = () => {
    if (filteredData.length === 0) {
      alert("No cancelled lots data to export.");
      return;
    }

    const exportRows = filteredData.map((item, idx) => ({
      "S.No": idx + 1,
      "Lot No": item.lotNumber || "-",
      "Job Order No": item.jobOrderNo || "-",
      "Date (PO Date)": item.date || "-",
      "Party Name": item.partyName || "-",
      "Brand": item.brand || "-",
      "Style": item.style || "-",
      "Fabric": item.fabric || "-",
      "Garment Type": item.garmentType || "-",
      "Section": item.section || "-",
      "Season": item.season || "-",
      "Quantity": item.quantity || "0",
      "Unit": item.unit || "-",
      "Shade": item.shade || "-",
      "Size": item.size || "-",
      "Cancellation Reason": item.cancellationReason || "-",
      "Cancellation Approved From": item.cancellationApprovedFrom || "-",
      "Cancelled By": item.cancelledBy || "-",
      "Cancellation Timestamp": item.cancellationTimestamp || "-",
      "Priority": item.priority || "-",
      "Direct Stitching": item.directStitching || "-",
      "Emb": item.emb || "-",
      "Emb Details": item.embDetails || "-",
      "Printing": item.printing || "-",
      "Printing Details": item.printingDetails || "-",
      "Pattern": item.pattern || "-",
      "Submitted By": item.submittedBy || "-",
      "Remarks": item.remarks || "-",
      "Status": item.status || "Cancel"
    }));

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const colKeys = Object.keys(exportRows[0]);
    ws["!cols"] = colKeys.map((k) => ({
      wch: Math.max(k.length + 4, 14)
    }));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Cancelled Lots");
    XLSX.writeFile(wb, `Cancelled_Lots_Report_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  // Export to PDF
  const exportToPDF = () => {
    if (filteredData.length === 0) {
      alert("No data available to export to PDF.");
      return;
    }

    const doc = new jsPDF("landscape", "mm", "a4");

    doc.setFont("helvetica", "bold");
    doc.setFontSize(16);
    doc.setTextColor(225, 29, 72);
    doc.text("MH FACTORY SUITE PRO — CANCELLED LOTS REPORT", 14, 15);

    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.setTextColor(100, 116, 139);
    doc.text(`Generated on: ${new Date().toLocaleString()} | Total Records: ${filteredData.length} | Financial Year: ${financialYear || "All"}`, 14, 21);

    const tableRows = filteredData.map((item, idx) => [
      idx + 1,
      item.lotNumber || "-",
      item.jobOrderNo || "-",
      item.date || "-",
      item.partyName || "-",
      item.brand || "-",
      item.style || "-",
      item.fabric || "-",
      `${item.quantity || "0"} ${item.unit || ""}`,
      item.cancellationReason || "-",
      item.cancellationApprovedFrom || item.cancelledBy || "-",
      item.status || "Cancel"
    ]);

    autoTable(doc, {
      startY: 26,
      head: [[
        "#", "Lot No", "Job Order No", "Date", "Party Name", "Brand", "Style", "Fabric", "Qty", "Cancellation Reason", "Approved / By", "Status"
      ]],
      body: tableRows,
      theme: "grid",
      headStyles: {
        fillColor: [225, 29, 72],
        textColor: 255,
        fontStyle: "bold",
        fontSize: 8,
        halign: "center"
      },
      styles: {
        fontSize: 7.5,
        cellPadding: 2,
        overflow: "linebreak"
      },
      alternateRowStyles: {
        fillColor: [255, 241, 242]
      }
    });

    doc.save(`Cancelled_Lots_Report_${new Date().toISOString().slice(0, 10)}.pdf`);
  };

  return (
    <>
      <style>{`
        .clr-container {
          min-height: 100vh;
          background: #f8fafc;
          padding: 24px 32px 60px;
          font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
          color: #1e293b;
        }
        .clr-header-card {
          background: #ffffff;
          border-radius: 16px;
          border: 1px solid #e2e8f0;
          padding: 20px 24px;
          display: flex;
          justify-content: space-between;
          align-items: center;
          flex-wrap: wrap;
          gap: 16px;
          box-shadow: 0 2px 6px rgba(0,0,0,0.03);
          margin-bottom: 20px;
        }
        .clr-btn {
          display: inline-flex;
          align-items: center;
          gap: 6px;
          padding: 8px 16px;
          font-size: 12px;
          font-weight: 700;
          border-radius: 10px;
          cursor: pointer;
          transition: all 0.2s;
          border: none;
        }
        .clr-btn-back {
          background: #f1f5f9;
          color: #334155;
          border: 1px solid #cbd5e1;
        }
        .clr-btn-back:hover {
          background: #e2e8f0;
        }
        .clr-btn-refresh {
          background: #ffffff;
          color: #334155;
          border: 1px solid #cbd5e1;
        }
        .clr-btn-refresh:hover {
          background: #f8fafc;
        }
        .clr-btn-excel {
          background: #059669;
          color: #ffffff;
        }
        .clr-btn-excel:hover {
          background: #047857;
        }
        .clr-btn-pdf {
          background: #e11d48;
          color: #ffffff;
        }
        .clr-btn-pdf:hover {
          background: #be123c;
        }
        .clr-stats-grid {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
          gap: 16px;
          margin-bottom: 20px;
        }
        .clr-stat-card {
          background: #ffffff;
          border-radius: 14px;
          padding: 16px 20px;
          border: 1px solid #e2e8f0;
          box-shadow: 0 2px 4px rgba(0,0,0,0.02);
          display: flex;
          flex-direction: column;
          gap: 4px;
        }
        .clr-stat-title {
          font-size: 11px;
          font-weight: 800;
          text-transform: uppercase;
          letter-spacing: 0.5px;
          color: #64748b;
        }
        .clr-stat-value {
          font-size: 24px;
          font-weight: 900;
          color: #0f172a;
        }
        .clr-stat-subtitle {
          font-size: 11px;
          color: #94a3b8;
          font-weight: 600;
        }
        .clr-filter-card {
          background: #ffffff;
          border-radius: 16px;
          border: 1px solid #e2e8f0;
          padding: 20px 24px;
          box-shadow: 0 2px 6px rgba(0,0,0,0.03);
          margin-bottom: 20px;
          display: flex;
          flex-direction: column;
          gap: 16px;
        }
        .clr-filter-row-1 {
          display: grid;
          grid-template-columns: 2fr 1fr 1fr;
          gap: 16px;
        }
        @media (max-width: 900px) {
          .clr-filter-row-1 {
            grid-template-columns: 1fr;
          }
        }
        .clr-filter-row-2 {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
          gap: 12px;
          padding-top: 14px;
          border-top: 1px solid #f1f5f9;
        }
        .clr-field-label {
          display: block;
          font-size: 11px;
          font-weight: 700;
          text-transform: uppercase;
          color: #475569;
          margin-bottom: 4px;
        }
        .clr-input {
          width: 100%;
          box-sizing: border-box;
          padding: 8px 12px;
          font-size: 12px;
          border-radius: 8px;
          border: 1px solid #cbd5e1;
          outline: none;
          background: #ffffff;
          transition: border-color 0.2s;
        }
        .clr-input:focus {
          border-color: #e11d48;
        }
        .clr-dropdown-wrapper {
          position: relative;
        }
        .clr-select-trigger {
          width: 100%;
          box-sizing: border-box;
          padding: 8px 12px;
          font-size: 12px;
          border-radius: 8px;
          border: 1px solid #cbd5e1;
          background: #ffffff;
          display: flex;
          justify-content: space-between;
          align-items: center;
          cursor: pointer;
        }
        .clr-select-trigger-text {
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
          color: #1e293b;
          font-weight: 600;
        }
        .clr-dropdown-popover {
          position: absolute;
          top: calc(100% + 4px);
          left: 0;
          width: 220px;
          max-height: 240px;
          background: #ffffff;
          border: 1px solid #cbd5e1;
          border-radius: 10px;
          box-shadow: 0 10px 25px rgba(0,0,0,0.1);
          z-index: 100;
          padding: 8px;
          box-sizing: border-box;
        }
        .clr-popover-header {
          display: flex;
          justify-content: space-between;
          align-items: center;
          padding-bottom: 6px;
          margin-bottom: 6px;
          border-bottom: 1px solid #f1f5f9;
          font-size: 11px;
          font-weight: 700;
          color: #64748b;
        }
        .clr-clear-btn {
          background: none;
          border: none;
          color: #e11d48;
          cursor: pointer;
          font-weight: 800;
          font-size: 11px;
        }
        .clr-popover-options {
          max-height: 180px;
          overflow-y: auto;
        }
        .clr-popover-item {
          display: flex;
          align-items: center;
          padding: 5px 6px;
          border-radius: 6px;
          cursor: pointer;
          font-size: 12px;
          color: #334155;
          font-weight: 500;
        }
        .clr-popover-item:hover {
          background: #fff1f2;
        }
        .clr-table-card {
          background: #ffffff;
          border-radius: 16px;
          border: 1px solid #e2e8f0;
          box-shadow: 0 2px 6px rgba(0,0,0,0.03);
          overflow: hidden;
        }
        .clr-table-wrapper {
          overflow-x: auto;
          max-height: 700px;
        }
        .clr-table {
          width: 100%;
          border-collapse: collapse;
          text-align: left;
          font-size: 12px;
        }
        .clr-table th {
          background: #0f172a;
          color: #f8fafc;
          padding: 12px 14px;
          font-size: 11px;
          font-weight: 800;
          text-transform: uppercase;
          letter-spacing: 0.5px;
          white-space: nowrap;
          position: sticky;
          top: 0;
          z-index: 10;
        }
        .clr-table td {
          padding: 12px 14px;
          border-bottom: 1px solid #f1f5f9;
          color: #334155;
        }
        .clr-table tr:hover td {
          background: #fff1f2;
        }
        .clr-badge-lot {
          background: #f1f5f9;
          border: 1px solid #cbd5e1;
          color: #0f172a;
          padding: 3px 8px;
          border-radius: 6px;
          font-weight: 800;
        }
        .clr-badge-cancel {
          background: #ffe4e6;
          color: #be123c;
          border: 1px solid #fecdd3;
          padding: 2px 8px;
          border-radius: 20px;
          font-weight: 800;
          font-size: 11px;
          text-transform: uppercase;
          display: inline-flex;
          align-items: center;
          gap: 4px;
        }
        .clr-reason-box {
          background: #fff1f2;
          color: #9f1239;
          border: 1px solid #fecdd3;
          padding: 4px 8px;
          border-radius: 6px;
          font-weight: 700;
          font-size: 11px;
          display: inline-block;
          max-width: 220px;
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
        }
        .clr-thumb {
          width: 36px;
          height: 36px;
          border-radius: 8px;
          object-fit: cover;
          border: 1px solid #cbd5e1;
          cursor: pointer;
        }
        .clr-pagination {
          display: flex;
          justify-content: space-between;
          align-items: center;
          padding: 14px 20px;
          background: #f8fafc;
          border-top: 1px solid #e2e8f0;
          font-size: 12px;
          font-weight: 600;
          flex-wrap: wrap;
          gap: 12px;
        }
        .clr-modal-overlay {
          position: fixed;
          inset: 0;
          background: rgba(0,0,0,0.6);
          backdrop-filter: blur(4px);
          display: flex;
          align-items: center;
          justify-content: center;
          z-index: 9999;
          padding: 20px;
        }
        .clr-modal-card {
          background: #ffffff;
          border-radius: 16px;
          width: 100%;
          max-width: 650px;
          max-height: 85vh;
          overflow-y: auto;
          box-shadow: 0 20px 40px rgba(0,0,0,0.25);
        }
        .clr-modal-header {
          background: #0f172a;
          color: #ffffff;
          padding: 16px 20px;
          display: flex;
          justify-content: space-between;
          align-items: center;
        }
        .clr-modal-body {
          padding: 20px;
          display: flex;
          flex-direction: column;
          gap: 16px;
          font-size: 12px;
        }
        .clr-detail-grid {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
          gap: 12px;
        }
        .clr-detail-item {
          background: #f8fafc;
          border: 1px solid #e2e8f0;
          border-radius: 10px;
          padding: 10px 14px;
        }
        .clr-detail-item-title {
          font-size: 10px;
          font-weight: 800;
          text-transform: uppercase;
          color: #64748b;
          margin-bottom: 2px;
        }
        .clr-detail-item-val {
          font-size: 12px;
          font-weight: 700;
          color: #0f172a;
        }
      `}</style>

      <div className="clr-container">
        {/* Header Card */}
        <div className="clr-header-card">
          <div style={{ display: "flex", alignItems: "center", gap: "14px" }}>
            <button
              onClick={() => history.push("/dashboard")}
              className="clr-btn clr-btn-back"
            >
              ← Dashboard
            </button>
            <div>
              <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                <span style={{ fontSize: "22px" }}>🚫</span>
                <h1 style={{ margin: 0, fontSize: "20px", fontWeight: 900, color: "#0f172a" }}>
                  Cancelled Lots Report
                </h1>
                <span style={{ fontSize: "11px", fontWeight: 800, background: "#ffe4e6", color: "#be123c", padding: "2px 8px", borderRadius: "12px", border: "1px solid #fecdd3" }}>
                  JobOrder Sheet
                </span>
              </div>
              <span style={{ fontSize: "12px", color: "#64748b", fontWeight: 500 }}>
                Dedicated audit & real-time monitoring of all cancelled job orders and cancelled lot records
              </span>
            </div>
          </div>

          <div style={{ display: "flex", gap: "10px", alignItems: "center", flexWrap: "wrap" }}>
            <button
              onClick={() => fetchCancelledLots(true)}
              disabled={refreshing || loading}
              className="clr-btn clr-btn-refresh"
            >
              ↻ {refreshing ? "Refreshing..." : "Refresh Data"}
            </button>
            <button
              onClick={exportToExcel}
              className="clr-btn clr-btn-excel"
            >
              📊 Export Excel
            </button>
            <button
              onClick={exportToPDF}
              className="clr-btn clr-btn-pdf"
            >
              📄 Export PDF
            </button>
          </div>
        </div>

        {/* Error Notification */}
        {error && (
          <div style={{ background: "#fef2f2", border: "1px solid #fecaca", color: "#b91c1c", padding: "12px 16px", borderRadius: "10px", marginBottom: "16px", fontSize: "12px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span>⚠️ {error}</span>
            <button onClick={() => fetchCancelledLots(true)} style={{ fontWeight: 800, background: "none", border: "none", color: "#991b1b", cursor: "pointer" }}>Retry</button>
          </div>
        )}

        {/* Top Metric Cards */}
        <div className="clr-stats-grid">
          <div className="clr-stat-card" style={{ borderLeft: "4px solid #e11d48" }}>
            <span className="clr-stat-title">Total Cancelled Lots</span>
            <span className="clr-stat-value">{stats.totalLots.toLocaleString()}</span>
            <span className="clr-stat-subtitle">cancelled orders</span>
          </div>

          <div className="clr-stat-card" style={{ borderLeft: "4px solid #f59e0b" }}>
            <span className="clr-stat-title">Cancelled Volume</span>
            <span className="clr-stat-value">{stats.totalQty.toLocaleString()}</span>
            <span className="clr-stat-subtitle">PCS / Sets</span>
          </div>

          <div className="clr-stat-card" style={{ borderLeft: "4px solid #3b82f6" }}>
            <span className="clr-stat-title">Affected Parties</span>
            <span className="clr-stat-value">{stats.uniqueParties}</span>
            <span className="clr-stat-subtitle">clients</span>
          </div>

          <div className="clr-stat-card" style={{ borderLeft: "4px solid #8b5cf6" }}>
            <span className="clr-stat-title">Affected Brands</span>
            <span className="clr-stat-value">{stats.uniqueBrands}</span>
            <span className="clr-stat-subtitle">brands</span>
          </div>

          <div className="clr-stat-card" style={{ borderLeft: "4px solid #64748b" }}>
            <span className="clr-stat-title">Top Cancel Reason</span>
            <span style={{ fontSize: "13px", fontWeight: 800, color: "#be123c", marginTop: "4px", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }} title={stats.topReason}>
              {stats.topReason}
            </span>
          </div>
        </div>

        {/* Filter Controls Card */}
        <div className="clr-filter-card">
          <div className="clr-filter-row-1">
            <div>
              <label className="clr-field-label">Quick Search</label>
              <input
                type="text"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="Search Lot #, Job Order #, Party, Style, Reason, Fabric..."
                className="clr-input"
              />
            </div>

            <div>
              <label className="clr-field-label">Financial Year</label>
              <select
                value={financialYear}
                onChange={(e) => setFinancialYear(e.target.value)}
                className="clr-input"
                style={{ fontWeight: 600 }}
              >
                <option value="">All Financial Years</option>
                {availableFYs.map((fy) => (
                  <option key={fy} value={fy}>
                    FY {fy}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label className="clr-field-label">Date Range</label>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "6px" }}>
                <input
                  type="date"
                  value={startDate}
                  onChange={(e) => setStartDate(e.target.value)}
                  className="clr-input"
                  title="From Date"
                />
                <input
                  type="date"
                  value={endDate}
                  onChange={(e) => setEndDate(e.target.value)}
                  className="clr-input"
                  title="To Date"
                />
              </div>
            </div>
          </div>

          <div className="clr-filter-row-2">
            <MultiSelectDropdown
              label="Party Name"
              options={distinctParties}
              selectedValues={selectedParties}
              onChange={setSelectedParties}
              placeholder="All Parties"
            />
            <MultiSelectDropdown
              label="Brand"
              options={distinctBrands}
              selectedValues={selectedBrands}
              onChange={setSelectedBrands}
              placeholder="All Brands"
            />
            <MultiSelectDropdown
              label="Fabric"
              options={distinctFabrics}
              selectedValues={selectedFabrics}
              onChange={setSelectedFabrics}
              placeholder="All Fabrics"
            />
            <MultiSelectDropdown
              label="Style"
              options={distinctStyles}
              selectedValues={selectedStyles}
              onChange={setSelectedStyles}
              placeholder="All Styles"
            />
            <MultiSelectDropdown
              label="Cancel Reason"
              options={distinctReasons}
              selectedValues={selectedReasons}
              onChange={setSelectedReasons}
              placeholder="All Reasons"
            />
            <MultiSelectDropdown
              label="Approved By"
              options={distinctApprovedBy}
              selectedValues={selectedApprovedBy}
              onChange={setSelectedApprovedBy}
              placeholder="All Approvers"
            />
          </div>

          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", paddingTop: "8px", fontSize: "11px", color: "#64748b", fontWeight: 700 }}>
            <span>Showing {filteredData.length} of {data.length} cancelled lots</span>
            {(search || startDate || endDate || selectedParties.length > 0 || selectedBrands.length > 0 || selectedFabrics.length > 0 || selectedStyles.length > 0 || selectedReasons.length > 0 || selectedApprovedBy.length > 0) && (
              <button
                onClick={clearFilters}
                style={{ color: "#e11d48", fontWeight: 800, background: "none", border: "none", cursor: "pointer", textDecoration: "underline" }}
              >
                Reset All Filters
              </button>
            )}
          </div>
        </div>

        {/* Data Table Card */}
        <div className="clr-table-card">
          {loading ? (
            <div style={{ padding: "80px", textAlign: "center", color: "#64748b", fontWeight: 700 }}>
              <div style={{ fontSize: "24px", marginBottom: "8px" }}>⌛</div>
              Loading cancelled lots directly from JobOrder sheet...
            </div>
          ) : filteredData.length === 0 ? (
            <div style={{ padding: "80px", textAlign: "center", color: "#64748b" }}>
              <div style={{ fontSize: "32px", marginBottom: "8px" }}>🚫</div>
              <h3 style={{ margin: "0 0 4px", fontSize: "16px", color: "#0f172a", fontWeight: 800 }}>No Cancelled Lots Found</h3>
              <span style={{ fontSize: "12px", color: "#94a3b8" }}>Try adjusting your search terms or filters</span>
            </div>
          ) : (
            <div className="clr-table-wrapper">
              <table className="clr-table">
                <thead>
                  <tr>
                    <th style={{ width: "40px", textAlign: "center" }}>#</th>
                    <th>Lot No</th>
                    <th>Job Order No</th>
                    <th>Image</th>
                    <th>PO Date</th>
                    <th>Party Name</th>
                    <th>Brand & Style</th>
                    <th>Fabric</th>
                    <th style={{ textAlign: "right" }}>Quantity</th>
                    <th>Shade & Size</th>
                    <th>Cancellation Reason</th>
                    <th>Approved / By</th>
                    <th>Cancel Date</th>
                    <th style={{ textAlign: "center" }}>Status</th>
                    <th style={{ textAlign: "center" }}>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {paginatedData.map((row, idx) => {
                    const directImg = getDirectImageUrl(row.imageUrl);
                    const globalIdx = (currentPage - 1) * pageSize + idx + 1;

                    return (
                      <tr key={`${row.jobOrderNo}-${row.lotNumber}-${idx}`}>
                        <td style={{ textAlign: "center", fontWeight: 800, color: "#94a3b8" }}>{globalIdx}</td>
                        <td>
                          <span className="clr-badge-lot">{row.lotNumber || "-"}</span>
                        </td>
                        <td style={{ fontWeight: 800, color: "#be123c" }}>{row.jobOrderNo || "-"}</td>
                        <td>
                          {directImg ? (
                            <img
                              src={directImg}
                              alt="Garment"
                              className="clr-thumb"
                              onClick={() => setLightboxImage(directImg)}
                              onError={(e) => { e.target.style.display = "none"; }}
                            />
                          ) : (
                            <span style={{ fontSize: "10px", color: "#94a3b8" }}>N/A</span>
                          )}
                        </td>
                        <td style={{ whiteSpace: "nowrap" }}>{formatDate(row.date)}</td>
                        <td style={{ fontWeight: 700, color: "#0f172a" }}>{row.partyName || "-"}</td>
                        <td>
                          <div style={{ fontWeight: 800, color: "#0f172a" }}>{row.brand || "-"}</div>
                          <div style={{ fontSize: "11px", color: "#64748b", maxWidth: "130px", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={row.style}>
                            {row.style || "-"}
                          </div>
                        </td>
                        <td>{row.fabric || "-"}</td>
                        <td style={{ textAlign: "right" }}>
                          <div style={{ fontWeight: 900, color: "#be123c", fontSize: "13px" }}>{row.quantity || "0"}</div>
                          <div style={{ fontSize: "9px", textTransform: "uppercase", color: "#94a3b8", fontWeight: 700 }}>{row.unit || "PCS"}</div>
                        </td>
                        <td>
                          <div style={{ fontSize: "11px", fontWeight: 700, maxWidth: "120px", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={row.shade}>
                            {row.shade || "-"}
                          </div>
                          <div style={{ fontSize: "10px", color: "#64748b", maxWidth: "120px", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={row.size}>
                            {row.size || "-"}
                          </div>
                        </td>
                        <td>
                          <span className="clr-reason-box" title={row.cancellationReason}>
                            ⚠️ {row.cancellationReason || "No Reason"}
                          </span>
                        </td>
                        <td>
                          <div style={{ fontWeight: 700, fontSize: "11px", color: "#0f172a" }}>
                            {row.cancellationApprovedFrom ? `Appr: ${row.cancellationApprovedFrom}` : "-"}
                          </div>
                          {row.cancelledBy && (
                            <div style={{ fontSize: "10px", color: "#64748b" }}>By: {row.cancelledBy}</div>
                          )}
                        </td>
                        <td style={{ fontSize: "11px", whiteSpace: "nowrap" }}>{formatDate(row.cancellationTimestamp)}</td>
                        <td style={{ textAlign: "center" }}>
                          <span className="clr-badge-cancel">
                            ❌ {row.status || "Cancel"}
                          </span>
                        </td>
                        <td style={{ textAlign: "center" }}>
                          <button
                            type="button"
                            onClick={() => setActiveLotDetail(row)}
                            style={{ background: "#ffe4e6", border: "1px solid #fecdd3", color: "#be123c", padding: "4px 8px", borderRadius: "6px", cursor: "pointer", fontWeight: 800, fontSize: "11px" }}
                          >
                            Details
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}

          {/* Pagination */}
          {!loading && filteredData.length > 0 && (
            <div className="clr-pagination">
              <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                <span>Rows per page:</span>
                <select
                  value={pageSize}
                  onChange={(e) => setPageSize(Number(e.target.value))}
                  style={{ padding: "4px 8px", borderRadius: "6px", border: "1px solid #cbd5e1", fontWeight: 700 }}
                >
                  <option value={10}>10</option>
                  <option value={25}>25</option>
                  <option value={50}>50</option>
                  <option value={100}>100</option>
                  <option value={500}>500</option>
                </select>
                <span style={{ color: "#64748b" }}>
                  Page {currentPage} of {totalPages} ({filteredData.length} records)
                </span>
              </div>

              <div style={{ display: "flex", gap: "6px", alignItems: "center" }}>
                <button
                  onClick={() => setCurrentPage(1)}
                  disabled={currentPage === 1}
                  style={{ padding: "4px 10px", borderRadius: "6px", border: "1px solid #cbd5e1", background: "#ffffff", cursor: "pointer", fontWeight: 700 }}
                >
                  First
                </button>
                <button
                  onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                  disabled={currentPage === 1}
                  style={{ padding: "4px 10px", borderRadius: "6px", border: "1px solid #cbd5e1", background: "#ffffff", cursor: "pointer", fontWeight: 700 }}
                >
                  Prev
                </button>
                <span style={{ padding: "4px 10px", background: "#ffe4e6", color: "#be123c", borderRadius: "6px", fontWeight: 900 }}>
                  {currentPage}
                </span>
                <button
                  onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                  disabled={currentPage === totalPages}
                  style={{ padding: "4px 10px", borderRadius: "6px", border: "1px solid #cbd5e1", background: "#ffffff", cursor: "pointer", fontWeight: 700 }}
                >
                  Next
                </button>
                <button
                  onClick={() => setCurrentPage(totalPages)}
                  disabled={currentPage === totalPages}
                  style={{ padding: "4px 10px", borderRadius: "6px", border: "1px solid #cbd5e1", background: "#ffffff", cursor: "pointer", fontWeight: 700 }}
                >
                  Last
                </button>
              </div>
            </div>
          )}
        </div>

        {/* Lot Detail Modal */}
        {activeLotDetail && (
          <div className="clr-modal-overlay" onClick={() => setActiveLotDetail(null)}>
            <div className="clr-modal-card" onClick={(e) => e.stopPropagation()}>
              <div className="clr-modal-header">
                <div>
                  <h3 style={{ margin: 0, fontSize: "15px", fontWeight: 800 }}>🚫 Cancelled Lot Specification</h3>
                  <span style={{ fontSize: "11px", color: "#94a3b8" }}>
                    Lot #{activeLotDetail.lotNumber} | Job Order #{activeLotDetail.jobOrderNo}
                  </span>
                </div>
                <button
                  onClick={() => setActiveLotDetail(null)}
                  style={{ background: "none", border: "none", color: "#ffffff", fontSize: "16px", cursor: "pointer" }}
                >
                  ✕
                </button>
              </div>

              <div className="clr-modal-body">
                {/* Cancellation Alert Callout */}
                <div style={{ background: "#fff1f2", border: "1.5px solid #fecdd3", borderRadius: "10px", padding: "12px 16px", color: "#9f1239" }}>
                  <span style={{ fontSize: "10px", fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.5px" }}>Cancellation Reason:</span>
                  <div style={{ fontSize: "14px", fontWeight: 900, marginTop: "2px" }}>
                    {activeLotDetail.cancellationReason || "No Specific Reason Recorded"}
                  </div>
                  <div style={{ display: "flex", gap: "16px", marginTop: "8px", fontSize: "11px", fontWeight: 700, flexWrap: "wrap" }}>
                    {activeLotDetail.cancellationApprovedFrom && <span>🛡️ Approved From: {activeLotDetail.cancellationApprovedFrom}</span>}
                    {activeLotDetail.cancelledBy && <span>👤 Cancelled By: {activeLotDetail.cancelledBy}</span>}
                    {activeLotDetail.cancellationTimestamp && <span>⏰ Date: {formatDate(activeLotDetail.cancellationTimestamp)}</span>}
                  </div>
                </div>

                <div className="clr-detail-grid">
                  <div className="clr-detail-item">
                    <span className="clr-detail-item-title">Party Name</span>
                    <span className="clr-detail-item-val">{activeLotDetail.partyName || "-"}</span>
                  </div>
                  <div className="clr-detail-item">
                    <span className="clr-detail-item-title">Brand</span>
                    <span className="clr-detail-item-val">{activeLotDetail.brand || "-"}</span>
                  </div>
                  <div className="clr-detail-item">
                    <span className="clr-detail-item-title">Fabric</span>
                    <span className="clr-detail-item-val">{activeLotDetail.fabric || "-"}</span>
                  </div>
                  <div className="clr-detail-item">
                    <span className="clr-detail-item-title">Style</span>
                    <span className="clr-detail-item-val">{activeLotDetail.style || "-"}</span>
                  </div>
                  <div className="clr-detail-item">
                    <span className="clr-detail-item-title">Garment Type</span>
                    <span className="clr-detail-item-val">{activeLotDetail.garmentType || "-"}</span>
                  </div>
                  <div className="clr-detail-item">
                    <span className="clr-detail-item-title">Quantity</span>
                    <span className="clr-detail-item-val" style={{ color: "#be123c", fontSize: "14px" }}>
                      {activeLotDetail.quantity || "0"} {activeLotDetail.unit || "PCS"}
                    </span>
                  </div>
                  <div className="clr-detail-item">
                    <span className="clr-detail-item-title">Section</span>
                    <span className="clr-detail-item-val">{activeLotDetail.section || "-"}</span>
                  </div>
                  <div className="clr-detail-item">
                    <span className="clr-detail-item-title">Season</span>
                    <span className="clr-detail-item-val">{activeLotDetail.season || "-"}</span>
                  </div>
                  <div className="clr-detail-item">
                    <span className="clr-detail-item-title">Priority</span>
                    <span className="clr-detail-item-val">{activeLotDetail.priority || "-"}</span>
                  </div>
                </div>

                <div className="clr-detail-item">
                  <span className="clr-detail-item-title">Shades</span>
                  <span className="clr-detail-item-val">{activeLotDetail.shade || "-"}</span>
                </div>

                <div className="clr-detail-item">
                  <span className="clr-detail-item-title">Sizes</span>
                  <span className="clr-detail-item-val">{activeLotDetail.size || "-"}</span>
                </div>

                {activeLotDetail.remarks && (
                  <div style={{ background: "#fffbeb", border: "1px solid #fde68a", borderRadius: "10px", padding: "10px 14px", color: "#92400e" }}>
                    <span style={{ fontSize: "10px", fontWeight: 800, textTransform: "uppercase" }}>Remarks</span>
                    <div style={{ fontWeight: 600, marginTop: "2px" }}>{activeLotDetail.remarks}</div>
                  </div>
                )}
              </div>

              <div style={{ padding: "12px 20px", background: "#f8fafc", borderTop: "1px solid #e2e8f0", textAlign: "right" }}>
                <button
                  onClick={() => setActiveLotDetail(null)}
                  style={{ background: "#0f172a", color: "#ffffff", border: "none", padding: "8px 16px", borderRadius: "8px", fontWeight: 700, fontSize: "12px", cursor: "pointer" }}
                >
                  Close
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Lightbox Modal */}
        {lightboxImage && (
          <div
            className="clr-modal-overlay"
            onClick={() => setLightboxImage(null)}
          >
            <div style={{ position: "relative", maxWidth: "80vw", maxHeight: "85vh", background: "#ffffff", padding: "8px", borderRadius: "14px" }} onClick={(e) => e.stopPropagation()}>
              <button
                onClick={() => setLightboxImage(null)}
                style={{ position: "absolute", top: "-10px", right: "-10px", background: "#0f172a", color: "#ffffff", border: "none", borderRadius: "50%", width: "26px", height: "26px", cursor: "pointer", fontWeight: 900 }}
              >
                ✕
              </button>
              <img
                src={lightboxImage}
                alt="Preview"
                style={{ maxWidth: "100%", maxHeight: "80vh", borderRadius: "10px", display: "block" }}
              />
            </div>
          </div>
        )}
      </div>
    </>
  );
}
