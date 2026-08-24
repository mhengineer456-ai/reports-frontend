import { useState, useEffect, useMemo } from 'react';
import { useHistory } from 'react-router-dom';
import { store } from './store.js';
import { fetchSheetDataFromBackend, SPREADSHEET_IDS } from './config.js';
import {
  Calendar, Search, Download, RefreshCw, FileText,
  Layers, Scale, Tag, Scissors, ArrowLeft, Clock
} from 'lucide-react';
import * as XLSX from "xlsx-js-style";
import { jsPDF } from 'jspdf';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell,
  AreaChart, Area, PieChart, Pie, Legend, ComposedChart, Line, LineChart
} from 'recharts';

const CHART_COLORS = ['#2563EB', '#10B981', '#F59E0B', '#EF4444', '#8B5CF6', '#EC4899', '#06B6D4', '#F97316'];

// Custom Glassmorphic Tooltip for Professional Graphs
const CustomTooltip = ({ active, payload, label }) => {
  if (active && payload && payload.length) {
    return (
      <div style={{
        background: 'rgba(15, 23, 42, 0.95)',
        backdropFilter: 'blur(8px)',
        border: '1px solid rgba(255, 255, 255, 0.15)',
        padding: '10px 14px',
        borderRadius: '8px',
        boxShadow: '0 6px 20px rgba(0,0,0,0.15)',
        color: '#fff',
        fontSize: '11px',
        fontFamily: 'inherit'
      }}>
        <p style={{ margin: '0 0 6px 0', fontWeight: 800, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.5px' }}>{label}</p>
        {payload.map((entry, idx) => (
          <div key={idx} style={{ display: 'flex', alignItems: 'center', gap: 8, margin: '4px 0' }}>
            <span style={{ width: 8, height: 8, borderRadius: '50%', background: entry.color || entry.fill, display: 'inline-block' }} />
            <span style={{ color: '#cbd5e1' }}>{entry.name}:</span>
            <span style={{ fontWeight: 800, color: '#f8fafc' }}>
              {typeof entry.value === 'number' ? entry.value.toFixed(1) : entry.value}
              {entry.name.toLowerCase().includes('weight') ? ' KG' : ' Roll(s)'}
            </span>
          </div>
        ))}
      </div>
    );
  }
  return null;
};

export default function DailyFabricIssueReport() {
  const history = useHistory();
  const [startDate, setStartDate] = useState(() => {
    const d = new Date();
    d.setDate(d.getDate() - 7); // Default to last 7 days
    return d.toISOString().slice(0, 10);
  });
  const [endDate, setEndDate] = useState(() => {
    return new Date().toISOString().slice(0, 10);
  });
  const [loading, setLoading] = useState(false);
  const [reportData, setReportData] = useState([]);
  const [searchTerm, setSearchTerm] = useState('');

  // Entry Pending Cutting Lot State
  const [isPendingCuttingFilter, setIsPendingCuttingFilter] = useState(false);
  const [pendingCuttingCount, setPendingCuttingCount] = useState(0);

  const fetchReport = async () => {
    setLoading(true);
    setIsPendingCuttingFilter(false);
    try {
      const response = await store.getDailyFabricIssuanceReport(startDate, endDate);
      if (Array.isArray(response)) {
        setReportData(response);
      } else if (response && response.success && Array.isArray(response.data)) {
        setReportData(response.data);
      } else if (response && Array.isArray(response.data)) {
        setReportData(response.data);
      } else {
        setReportData([]);
      }
    } catch (err) {
      console.error("Error loading daily fabric issue report:", err);
      alert("Failed to load report: " + err.message);
    } finally {
      setLoading(false);
    }
  };

  const fetchPendingCuttingLots = async () => {
    setLoading(true);
    setIsPendingCuttingFilter(true);
    const jul1Date = '2026-07-01';
    setStartDate(jul1Date);

    try {
      // 1. Fetch fabric issuance data from 1 July 2026 onwards
      const response = await store.getDailyFabricIssuanceReport(jul1Date, endDate);
      let allIssuances = [];
      if (Array.isArray(response)) {
        allIssuances = response;
      } else if (response && response.success && Array.isArray(response.data)) {
        allIssuances = response.data;
      } else if (response && Array.isArray(response.data)) {
        allIssuances = response.data;
      }

      // 2. Fetch completed lot numbers from Index Sheet
      const completedLotsSet = new Set();

      // Fallback A: Try Store completed report endpoint
      try {
        if (store.getDailyCuttingCompletedReport) {
          const completedRes = await store.getDailyCuttingCompletedReport(true).catch(() => null);
          if (completedRes && Array.isArray(completedRes.data)) {
            completedRes.data.forEach(item => {
              const l = String(item.lotNo || item.lotNumber || item.lot || '').trim().toLowerCase();
              if (l) completedLotsSet.add(l);
            });
          }
        }
      } catch (e) {
        console.warn("Backend completed report fetch error:", e);
      }

      // Fallback B: Fetch directly from Index sheet (Index!A:Z)
      try {
        const indexRes = await fetchSheetDataFromBackend(SPREADSHEET_IDS.MAIN, 'Index!A:Z');
        if (indexRes && indexRes.ok && Array.isArray(indexRes.values) && indexRes.values.length > 0) {
          const headers = indexRes.values[0].map(h => String(h || '').trim().toLowerCase());
          let lotColIdx = headers.findIndex(h => h.includes('lot') || h.includes('lot no') || h.includes('lot number'));
          if (lotColIdx === -1) lotColIdx = 0;

          for (let i = 1; i < indexRes.values.length; i++) {
            const row = indexRes.values[i];
            if (row && row[lotColIdx]) {
              const val = String(row[lotColIdx]).trim().toLowerCase();
              if (val && val !== '-' && val !== 'null') {
                completedLotsSet.add(val);
              }
            }
          }
        }
      } catch (e) {
        console.warn("Index sheet direct fetch error:", e);
      }

      // 3. Filter Fabric Issuance Lots where Lot Number is NOT in Index Sheet & DEDUPLICATE (Single Lot = Single Entry)
      const lotMap = new Map();

      allIssuances.forEach(item => {
        const rawLot = String(item.lotNumber || item.lotNo || item.lot || '').trim();
        if (!rawLot) return;
        const key = rawLot.toLowerCase();

        // Include ONLY if Lot Number is NOT in Index sheet
        if (!completedLotsSet.has(key)) {
          const tbl = item.tableNumber || item.table || item.tableName || item.tableNo || item.section || '—';
          if (!lotMap.has(key)) {
            lotMap.set(key, {
              ...item,
              lotNumber: rawLot,
              date: item.date || item.issueDate || '-',
              fabric: item.fabric || item.fabricDescription || item.fabricName || '—',
              tableNumber: tbl,
              rolls: item.rolls || 1,
              weight: item.weight || 0
            });
          } else {
            // Aggregate rolls & weight for identical lot number
            const existing = lotMap.get(key);
            existing.rolls += (item.rolls || 1);
            existing.weight += (item.weight || 0);
            if (!existing.tableNumber || existing.tableNumber === '—') {
              existing.tableNumber = tbl;
            }
          }
        }
      });

      const uniquePendingData = Array.from(lotMap.values());
      setReportData(uniquePendingData);
      setPendingCuttingCount(uniquePendingData.length);
    } catch (err) {
      console.error("Error loading pending cutting lots:", err);
      alert("Failed to load pending cutting lots: " + err.message);
    } finally {
      setLoading(false);
    }
  };

  const handlePendingToggle = () => {
    if (isPendingCuttingFilter) {
      setIsPendingCuttingFilter(false);
      fetchReport();
    } else {
      fetchPendingCuttingLots();
    }
  };

  useEffect(() => {
    if (!isPendingCuttingFilter) {
      fetchReport();
    }
  }, [startDate, endDate]);

  // General Text Filtering
  const filteredData = useMemo(() => {
    return reportData.filter(item => {
      const q = searchTerm.toLowerCase().trim();
      if (!q) return true;
      return (
        String(item.tableNumber).toLowerCase().includes(q) ||
        String(item.fabric).toLowerCase().includes(q) ||
        String(item.lotNumber).toLowerCase().includes(q) ||
        String(item.jobOrderNo).toLowerCase().includes(q) ||
        String(item.shade).toLowerCase().includes(q) ||
        String(item.issuedBy).toLowerCase().includes(q)
      );
    });
  }, [reportData, searchTerm]);

  // Statistics
  const stats = useMemo(() => {
    return filteredData.reduce((acc, curr) => {
      acc.totalRolls += curr.rolls || 0;
      acc.totalWeight += curr.weight || 0;
      acc.uniqueLots.add(curr.lotNumber);
      acc.activeTables.add(curr.tableNumber);
      acc.uniqueFabrics.add(curr.fabric);
      return acc;
    }, {
      totalRolls: 0,
      totalWeight: 0,
      uniqueLots: new Set(),
      activeTables: new Set(),
      uniqueFabrics: new Set()
    });
  }, [filteredData]);

  // Aggregation 1: Table-wise Issuance Summary
  const tableSummary = useMemo(() => {
    const summaryMap = {};
    filteredData.forEach(item => {
      const tbl = item.tableNumber || 'N/A';
      if (!summaryMap[tbl]) {
        summaryMap[tbl] = { name: tbl, rolls: 0, weight: 0, lots: new Set() };
      }
      summaryMap[tbl].rolls += item.rolls || 0;
      summaryMap[tbl].weight += item.weight || 0;
      summaryMap[tbl].lots.add(item.lotNumber);
    });

    return Object.values(summaryMap)
      .map(item => ({
        ...item,
        uniqueLotsCount: item.lots.size,
        percentage: stats.totalRolls > 0 ? Math.round((item.rolls / stats.totalRolls) * 100) : 0
      }))
      .sort((a, b) => b.rolls - a.rolls);
  }, [filteredData, stats.totalRolls]);

  // Aggregation 2: Fabric-wise Summary
  const fabricSummary = useMemo(() => {
    const summaryMap = {};
    filteredData.forEach(item => {
      const fab = item.fabric || '—';
      if (!summaryMap[fab]) {
        summaryMap[fab] = { name: fab, rolls: 0, weight: 0, lots: new Set(), shades: new Set() };
      }
      summaryMap[fab].rolls += item.rolls || 0;
      summaryMap[fab].weight += item.weight || 0;
      summaryMap[fab].lots.add(item.lotNumber);
      summaryMap[fab].shades.add(item.shade);
    });

    return Object.values(summaryMap)
      .map(item => ({
        ...item,
        uniqueLotsCount: item.lots.size,
        uniqueShadesCount: item.shades.size,
        percentage: stats.totalRolls > 0 ? Math.round((item.rolls / stats.totalRolls) * 100) : 0
      }))
      .sort((a, b) => b.rolls - a.rolls);
  }, [filteredData, stats.totalRolls]);

  // Aggregation 3: Daily Issuance Trend
  const trendChartData = useMemo(() => {
    const dailyMap = {};
    filteredData.forEach(item => {
      const date = item.date || 'No Date';
      if (!dailyMap[date]) {
        dailyMap[date] = { date, rolls: 0, weight: 0 };
      }
      dailyMap[date].rolls += item.rolls || 0;
      dailyMap[date].weight += item.weight || 0;
    });
    return Object.values(dailyMap)
      .map(item => ({
        ...item,
        formattedDate: new Date(item.date).toLocaleDateString(undefined, { month: 'short', day: 'numeric' })
      }))
      .sort((a, b) => a.date.localeCompare(b.date));
  }, [filteredData]);

  // Chart 1: Table-wise Composed Data (rolls and weight combined)
  const tableChartData = useMemo(() => {
    return tableSummary.map(item => ({
      name: item.name,
      rolls: item.rolls,
      weight: item.weight
    }));
  }, [tableSummary]);

  // Chart 2: Fabric-wise (Top 6)
  const fabricChartData = useMemo(() => {
    const data = fabricSummary.map(item => ({
      name: item.name.length > 20 ? item.name.slice(0, 18) + '...' : item.name,
      value: item.rolls
    }));
    if (data.length <= 6) return data;
    const top = data.slice(0, 5);
    const otherRolls = data.slice(5).reduce((acc, curr) => acc + curr.value, 0);
    top.push({ name: 'Other Fabrics', value: otherRolls });
    return top;
  }, [fabricSummary]);

  // Excel exporter (Multi-sheet summary with lot list)
  const exportToExcel = () => {
    if (filteredData.length === 0) {
      alert("No data available to export.");
      return;
    }

    const tableData = tableSummary.map((item, idx) => ({
      "SR": idx + 1,
      "Table Name": item.name,
      "Lot Number(s)": Array.from(item.lots).sort().join(', '),
      "Total Rolls Issued": item.rolls,
      "Total Weight Issued (KG)": parseFloat(item.weight.toFixed(2)),
      "Unique Lots Count": item.uniqueLotsCount,
      "Roll Share (%)": `${item.percentage}%`
    }));

    const fabricData = fabricSummary.map((item, idx) => ({
      "SR": idx + 1,
      "Fabric Description": item.name,
      "Lot Number(s)": Array.from(item.lots).sort().join(', '),
      "Total Rolls Issued": item.rolls,
      "Total Weight Issued (KG)": parseFloat(item.weight.toFixed(2)),
      "Unique Lots Count": item.uniqueLotsCount,
      "Unique Shades Count": item.uniqueShadesCount,
      "Roll Share (%)": `${item.percentage}%`
    }));

    const wb = XLSX.utils.book_new();

    // Sheet 1: Table Summary
    const wsTable = XLSX.utils.json_to_sheet(tableData);
    let range = XLSX.utils.decode_range(wsTable['!ref']);
    for (let c = range.s.c; c <= range.e.c; c++) {
      const headerCell = XLSX.utils.encode_cell({ r: 0, c });
      if (wsTable[headerCell]) {
        wsTable[headerCell].s = {
          fill: { fgColor: { rgb: "334155" } }, // Slate-700
          font: { bold: true, color: { rgb: "FFFFFF" }, name: "Calibri", sz: 11 },
          alignment: { horizontal: "center", vertical: "center" }
        };
      }
    }
    wsTable['!cols'] = [{ wch: 6 }, { wch: 18 }, { wch: 25 }, { wch: 18 }, { wch: 22 }, { wch: 18 }, { wch: 15 }];
    XLSX.utils.book_append_sheet(wb, wsTable, "Table-wise Summary");

    // Sheet 2: Fabric Summary
    const wsFabric = XLSX.utils.json_to_sheet(fabricData);
    range = XLSX.utils.decode_range(wsFabric['!ref']);
    for (let c = range.s.c; c <= range.e.c; c++) {
      const headerCell = XLSX.utils.encode_cell({ r: 0, c });
      if (wsFabric[headerCell]) {
        wsFabric[headerCell].s = {
          fill: { fgColor: { rgb: "475569" } }, // Slate-600
          font: { bold: true, color: { rgb: "FFFFFF" }, name: "Calibri", sz: 11 },
          alignment: { horizontal: "center", vertical: "center" }
        };
      }
    }
    wsFabric['!cols'] = [{ wch: 6 }, { wch: 30 }, { wch: 25 }, { wch: 18 }, { wch: 22 }, { wch: 18 }, { wch: 18 }, { wch: 15 }];
    XLSX.utils.book_append_sheet(wb, wsFabric, "Fabric-wise Summary");

    XLSX.writeFile(wb, `Daily_Fabric_Issue_Summary_${startDate}_to_${endDate}.xlsx`);
  };

  const getTodayAttendanceText = async () => {
    const todayStr = new Date().toISOString().slice(0, 10);
    let hodsPresent = 0;
    let supervisorsPresent = 0;
    let helpersPresent = 0;
    const absentees = [];

    const safeParseJSON = (val) => {
      if (!val) return [];
      if (typeof val === 'object') return val;
      try { return JSON.parse(val); } catch (e) { return []; }
    };

    try {
      const attRes = await store.getAttendance(todayStr);
      if (attRes && attRes.success && attRes.data) {
        attRes.data.forEach(record => {
          const recordHods = safeParseJSON(record.hods);
          const recordSups = safeParseJSON(record.supervisors);
          const recordHelpers = safeParseJSON(record.helpers);

          recordHods.forEach(h => {
            if (h.status === 'Present' || h.status === 'Half Day') {
              hodsPresent++;
            } else if (h.status === 'Absent') {
              absentees.push(`${h.name} (HOD)`);
            }
          });

          recordSups.forEach(s => {
            if (s.status === 'Present' || s.status === 'Half Day') {
              supervisorsPresent++;
            } else if (s.status === 'Absent') {
              absentees.push(`${s.name} (Supervisor)`);
            }
          });

          recordHelpers.forEach(hp => {
            if (hp.status === 'Present' || hp.status === 'Half Day') {
              helpersPresent++;
            } else if (hp.status === 'Absent') {
              absentees.push(`${hp.name} (Helper)`);
            }
          });
        });
      }
    } catch (e) {
      console.error("Failed to load today's attendance for PDF:", e);
    }

    const uniqueAbsentees = [...new Set(absentees)];

    return {
      summary: `HODs Present: ${hodsPresent} | Supervisors Present: ${supervisorsPresent} | Helpers Present: ${helpersPresent}`,
      absenteesText: uniqueAbsentees.length > 0 ? `Absentees: ${uniqueAbsentees.join(', ')}` : "Absentees: None"
    };
  };

  // Separate PDF Exporter for "ENTRY PENDING LOT"
  const exportPendingCuttingPdf = () => {
    if (filteredData.length === 0) {
      alert("No pending cutting data available to export.");
      return;
    }

    const doc = new jsPDF({
      orientation: "portrait",
      unit: "pt",
      format: "a4"
    });

    const PAGE_W = doc.internal.pageSize.getWidth();
    const PAGE_H = doc.internal.pageSize.getHeight();
    const M = 36; // Margin
    const contentW = PAGE_W - (M * 2);
    let y = 30;

    const setFont = (style, size) => {
      doc.setFont("helvetica", style);
      doc.setFontSize(size);
    };

    // --- HERO HEADER BANNER ---
    doc.setFillColor(30, 27, 75); // Deep Indigo (#1e1b4b)
    doc.rect(0, 0, PAGE_W, 60, "F");

    setFont("bold", 15);
    doc.setTextColor(255, 255, 255);
    doc.text("FABRIC ISSUED BUT NOT ENTERED IN SYSTEM", M, 34);

    setFont("normal", 9);
    doc.setTextColor(199, 210, 254);
    doc.text(`Cutting Entry Pending Lots  |  Issued Period: ${startDate} to ${endDate}  |  Generated: ${new Date().toLocaleDateString()}`, M, 48);

    y = 75;

    // --- WARNING ALERT STRIP ---
    doc.setFillColor(254, 243, 199); // Soft Amber (#fef3c7)
    doc.setDrawColor(245, 158, 11); // Amber border (#f59e0b)
    doc.setLineWidth(1);
    doc.roundedRect(M, y, contentW, 26, 4, 4, "FD");

    setFont("bold", 9.5);
    doc.setTextColor(146, 64, 14); // Dark Amber (#92400e)
    doc.text(`PENDING SUMMARY: ${filteredData.length} Lots Fabric Issued but Not Entered in System `, M + 10, y + 16);

    y += 40;

    // --- TABLE OF PENDING LOTS (NO WEIGHT, NO ROLLS!) ---
    // Columns: [Sr No, Issue Date, Lot Number, Fabric Description, Cutting Table, Status]
    const tableHeaders = ["#", "Issue Date", "Lot Number", "Fabric Description", "Cutting Table", "Status"];
    const colWidths = [30, 75, 95, 175, 75, 73];

    // Table Header Row
    doc.setFillColor(30, 41, 59); // Slate Navy (#1e293b)
    doc.rect(M, y, contentW, 22, "F");

    setFont("bold", 8.5);
    doc.setTextColor(255, 255, 255);
    let tx = M;
    tableHeaders.forEach((h, idx) => {
      const align = (idx === 0 || idx === 1 || idx === 4 || idx === 5) ? "center" : "left";
      let offset = 8;
      if (align === "center") offset = colWidths[idx] / 2;
      doc.text(h, tx + offset, y + 14, { align });
      tx += colWidths[idx];
    });

    y += 22;

    // Table Rows
    setFont("normal", 8.5);
    filteredData.forEach((item, index) => {
      // Page break check
      if (y + 24 > PAGE_H - 40) {
        // Page Footer
        setFont("normal", 8);
        doc.setTextColor(148, 163, 184);
        doc.text(`Page ${doc.internal.getNumberOfPages()}  |  FABRIC ISSUED BUT CUTTING PENDING `, M, PAGE_H - 15);

        doc.addPage();
        y = 35;

        // Repeat Table Header Row on new page
        doc.setFillColor(30, 41, 59);
        doc.rect(M, y, contentW, 22, "F");

        setFont("bold", 8.5);
        doc.setTextColor(255, 255, 255);
        let rx = M;
        tableHeaders.forEach((h, idx) => {
          const align = (idx === 0 || idx === 1 || idx === 4 || idx === 5) ? "center" : "left";
          let offset = 8;
          if (align === "center") offset = colWidths[idx] / 2;
          doc.text(h, rx + offset, y + 14, { align });
          rx += colWidths[idx];
        });
        y += 22;
      }

      // Zebra striping
      if (index % 2 === 1) {
        doc.setFillColor(248, 250, 252);
        doc.rect(M, y, contentW, 20, "F");
      }

      const srNo = String(index + 1);
      const issueDate = item.date || item.issueDate || '-';
      const lotNo = String(item.lotNumber || item.lotNo || item.lot || '-');
      const fabricDesc = String(item.fabric || item.fabricDescription || item.fabricName || '-');
      const tableNo = String(item.tableNumber || item.table || item.tableName || item.tableNo || item.section || '-');
      const statusText = "PENDING ENTRY";

      let cx = M;

      // 0: #
      setFont("normal", 8);
      doc.setTextColor(100, 116, 139);
      doc.text(srNo, cx + (colWidths[0] / 2), y + 13, { align: "center" });
      cx += colWidths[0];

      // 1: Issue Date
      doc.setTextColor(51, 65, 85);
      doc.text(issueDate, cx + (colWidths[1] / 2), y + 13, { align: "center" });
      cx += colWidths[1];

      // 2: Lot Number (Bold Royal Blue)
      setFont("bold", 8.5);
      doc.setTextColor(29, 78, 216);
      doc.text(lotNo, cx + 8, y + 13);
      cx += colWidths[2];

      // 3: Fabric Description
      setFont("normal", 8.5);
      doc.setTextColor(15, 23, 42);
      const truncatedFabric = fabricDesc.length > 32 ? fabricDesc.slice(0, 30) + '...' : fabricDesc;
      doc.text(truncatedFabric, cx + 8, y + 13);
      cx += colWidths[3];

      // 4: Cutting Table
      setFont("normal", 8.5);
      doc.setTextColor(30, 41, 59);
      doc.text(tableNo, cx + (colWidths[4] / 2), y + 13, { align: "center" });
      cx += colWidths[4];

      // 5: Status (Amber Bold)
      setFont("bold", 8);
      doc.setTextColor(180, 83, 9);
      doc.text(statusText, cx + (colWidths[5] / 2), y + 13, { align: "center" });

      y += 20;
      doc.setDrawColor(226, 232, 240);
      doc.setLineWidth(0.5);
      doc.line(M, y, M + contentW, y);
    });

    // Final Page Footer
    setFont("normal", 8);
    doc.setTextColor(148, 163, 184);
    doc.text(`Page ${doc.internal.getNumberOfPages()}  |  FABRIC ISSUED BUT CUTTING PENDING   |  Total Pending: ${filteredData.length}`, M, PAGE_H - 15);

    doc.save(`Fabric_Issued_Not_Entered_${startDate}_to_${endDate}.pdf`);
  };

  // PDF exporter (Grayscale / Professional Layout)
  const exportToPdf = async () => {
    if (isPendingCuttingFilter) {
      exportPendingCuttingPdf();
      return;
    }

    if (filteredData.length === 0) {
      alert("No data available to export.");
      return;
    }

    // Fetch today's attendance
    const attData = await getTodayAttendanceText();

    const doc = new jsPDF({
      orientation: "portrait",
      unit: "pt",
      format: "a4"
    });

    const PAGE_W = doc.internal.pageSize.getWidth();
    const PAGE_H = doc.internal.pageSize.getHeight();
    const M = 40; // Margin
    let y = 35;

    const setFont = (style, size) => {
      doc.setFont("helvetica", style);
      doc.setFontSize(size);
    };

    // Left Side - Header Title
    setFont("bold", 14);
    doc.setTextColor(15, 23, 42); // slate-900
    doc.text("DAILY FABRIC ISSUANCE ANALYSIS", M, y + 15);

    setFont("normal", 8.5);
    doc.setTextColor(100, 116, 139); // slate-500
    doc.text(`Period: ${startDate} to ${endDate}  |  Generated: ${new Date().toLocaleDateString()}`, M, y + 28);

    // Right Side - Today's Attendance Block
    doc.setTextColor(15, 23, 42);
    setFont("bold", 8);
    doc.text("TODAY'S ATTENDANCE SUMMARY", PAGE_W - M - 230, y + 10);
    setFont("normal", 7.5);
    doc.setTextColor(71, 85, 105);
    doc.text(attData.summary, PAGE_W - M - 230, y + 21);
    doc.text(attData.absenteesText, PAGE_W - M - 230, y + 31);

    // Divider Line
    doc.setDrawColor(15, 23, 42);
    doc.setLineWidth(1);
    doc.line(M, y + 39, PAGE_W - M, y + 39);

    y += 54;

    // Grayscale Summary Box
    doc.setDrawColor(226, 232, 240); // slate-200
    doc.setLineWidth(0.5);
    doc.rect(M, y, PAGE_W - 2 * M, 45);

    setFont("bold", 8.5);
    doc.setTextColor(71, 85, 105); // slate-600
    doc.text("TOTAL ROLLS ISSUED", M + 20, y + 18);
    doc.text("TOTAL WEIGHT ISSUED", M + 150, y + 18);
    doc.text("ENGAGED TABLES", M + 280, y + 18);
    doc.text("UNIQUE FABRIC TYPES", M + 400, y + 18);

    setFont("bold", 12);
    doc.setTextColor(15, 23, 42);
    doc.text(`${stats.totalRolls}`, M + 20, y + 34);
    doc.text(`${stats.totalWeight.toFixed(1)} kg`, M + 150, y + 34);
    doc.text(`${stats.activeTables.size}`, M + 280, y + 34);
    doc.text(`${stats.uniqueFabrics.size}`, M + 400, y + 34);
    y += 70;

    // ── SECTION 1: CUTTING TABLE SUMMARY ──────────────────────────────────
    setFont("bold", 11);
    doc.setTextColor(30, 41, 59); // slate-800
    doc.text("1. Cutting Table Summary", M, y);
    y += 12;

    const tHeaders = ["Table", "Lot Numbers", "Rolls", "Weight (KG)", "Share (%)"];
    const tColWidths = [100, 180, 60, 100, 75];
    let tTotalW = tColWidths.reduce((a, b) => a + b, 0);

    // Thick border above table header
    doc.setDrawColor(15, 23, 42);
    doc.setLineWidth(1.5);
    doc.line(M, y, M + tTotalW, y);

    setFont("bold", 9);
    doc.setTextColor(15, 23, 42);
    let tx = M;
    tHeaders.forEach((h, idx) => {
      const align = (idx === 2 || idx === 3 || idx === 4) ? "right" : "left";
      const offset = align === "right" ? tColWidths[idx] - 10 : 10;
      doc.text(h, tx + offset, y + 14, { align });
      tx += tColWidths[idx];
    });

    // Divider line underneath headers
    doc.setLineWidth(0.75);
    doc.line(M, y + 20, M + tTotalW, y + 20);
    y += 20;

    setFont("normal", 8.5);
    doc.setTextColor(51, 65, 85);
    tableSummary.forEach((item) => {
      const lotsStr = Array.from(item.lots).sort().join(', ');
      const truncatedLots = lotsStr.length > 38 ? lotsStr.slice(0, 35) + '...' : lotsStr;

      let rx = M;
      doc.text(item.name, rx + 10, y + 11); rx += tColWidths[0];
      doc.text(truncatedLots, rx + 10, y + 11); rx += tColWidths[1];
      doc.text(String(item.rolls), rx + tColWidths[2] - 10, y + 11, { align: "right" }); rx += tColWidths[2];
      doc.text(item.weight.toFixed(1), rx + tColWidths[3] - 10, y + 11, { align: "right" }); rx += tColWidths[3];
      doc.text(`${item.percentage}%`, rx + tColWidths[4] - 10, y + 11, { align: "right" });

      y += 16;
      doc.setDrawColor(226, 232, 240);
      doc.setLineWidth(0.5);
      doc.line(M, y, M + tTotalW, y);
    });

    // Total Row Table 1
    const tTotalRolls = tableSummary.reduce((sum, item) => sum + item.rolls, 0);
    const tTotalWeight = tableSummary.reduce((sum, item) => sum + item.weight, 0);

    setFont("bold", 8.5);
    doc.setTextColor(15, 23, 42);
    let rx = M;
    doc.text("Total", rx + 10, y + 11); rx += tColWidths[0];
    doc.text("", rx + 10, y + 11); rx += tColWidths[1];
    doc.text(String(tTotalRolls), rx + tColWidths[2] - 10, y + 11, { align: "right" }); rx += tColWidths[2];
    doc.text(tTotalWeight.toFixed(1), rx + tColWidths[3] - 10, y + 11, { align: "right" }); rx += tColWidths[3];
    doc.text("100%", rx + tColWidths[4] - 10, y + 11, { align: "right" });

    y += 16;
    doc.setDrawColor(15, 23, 42);
    doc.setLineWidth(1);
    doc.line(M, y, M + tTotalW, y);
    y += 35;

    // ── SECTION 2: FABRIC TYPE SUMMARY ────────────────────────────────────
    if (y + 160 > PAGE_H) {
      doc.addPage();
      y = 50;
    }

    setFont("bold", 11);
    doc.setTextColor(30, 41, 59);
    doc.text("2. Fabric Description Summary", M, y);
    y += 12;

    const fHeaders = ["Fabric Description", "Lot Numbers", "Rolls", "Weight (KG)", "Share (%)"];
    const fColWidths = [120, 150, 60, 100, 75];
    let fTotalW = fColWidths.reduce((a, b) => a + b, 0);

    // Thick border above table header
    doc.setDrawColor(15, 23, 42);
    doc.setLineWidth(1.5);
    doc.line(M, y, M + fTotalW, y);

    setFont("bold", 9);
    doc.setTextColor(15, 23, 42);
    tx = M;
    fHeaders.forEach((h, idx) => {
      const align = (idx === 2 || idx === 3 || idx === 4) ? "right" : "left";
      const offset = align === "right" ? fColWidths[idx] - 10 : 10;
      doc.text(h, tx + offset, y + 14, { align });
      tx += fColWidths[idx];
    });

    // Divider line underneath headers
    doc.setLineWidth(0.75);
    doc.line(M, y + 20, M + fTotalW, y + 20);
    y += 20;

    setFont("normal", 8.5);
    doc.setTextColor(51, 65, 85);
    fabricSummary.forEach((item) => {
      if (y + 20 > PAGE_H - 40) {
        doc.addPage();
        y = 50;
        doc.setDrawColor(15, 23, 42);
        doc.setLineWidth(1.5);
        doc.line(M, y, M + fTotalW, y);

        setFont("bold", 9);
        doc.setTextColor(15, 23, 42);
        let tfx = M;
        fHeaders.forEach((h, fIdx) => {
          const align = (fIdx === 2 || fIdx === 3 || fIdx === 4) ? "right" : "left";
          const offset = align === "right" ? fColWidths[fIdx] - 10 : 10;
          doc.text(h, tfx + offset, y + 14, { align });
          tfx += fColWidths[fIdx];
        });
        y += 20;
        setFont("normal", 8.5);
        doc.setTextColor(51, 65, 85);
      }

      const lotsStr = Array.from(item.lots).sort().join(', ');
      const truncatedLots = lotsStr.length > 32 ? lotsStr.slice(0, 29) + '...' : lotsStr;

      let rx = M;
      doc.text(item.name.length > 25 ? item.name.slice(0, 22) + '...' : item.name, rx + 10, y + 11); rx += fColWidths[0];
      doc.text(truncatedLots, rx + 10, y + 11); rx += fColWidths[1];
      doc.text(String(item.rolls), rx + fColWidths[2] - 10, y + 11, { align: "right" }); rx += fColWidths[2];
      doc.text(item.weight.toFixed(1), rx + fColWidths[3] - 10, y + 11, { align: "right" }); rx += fColWidths[3];
      doc.text(`${item.percentage}%`, rx + fColWidths[4] - 10, y + 11, { align: "right" });

      y += 16;
      doc.setDrawColor(226, 232, 240);
      doc.setLineWidth(0.5);
      doc.line(M, y, M + fTotalW, y);
    });

    // Total Row Table 2
    const fTotalRolls = fabricSummary.reduce((sum, item) => sum + item.rolls, 0);
    const fTotalWeight = fabricSummary.reduce((sum, item) => sum + item.weight, 0);

    setFont("bold", 8.5);
    doc.setTextColor(15, 23, 42);
    let frx = M;
    doc.text("Total", frx + 10, y + 11); frx += fColWidths[0];
    doc.text("", frx + 10, y + 11); frx += fColWidths[1];
    doc.text(String(fTotalRolls), frx + fColWidths[2] - 10, y + 11, { align: "right" }); frx += fColWidths[2];
    doc.text(fTotalWeight.toFixed(1), frx + fColWidths[3] - 10, y + 11, { align: "right" }); frx += fColWidths[3];
    doc.text("100%", frx + fColWidths[4] - 10, y + 11, { align: "right" });

    y += 16;
    doc.setDrawColor(15, 23, 42);
    doc.setLineWidth(1);
    doc.line(M, y, M + fTotalW, y);

    const pages = doc.internal.getNumberOfPages();
    for (let p = 1; p <= pages; p++) {
      doc.setPage(p);
      setFont("italic", 8);
      doc.setTextColor(148, 163, 184);
      doc.text(`Daily Fabric Issuance Analysis  |  Page ${p} of ${pages}`, M, PAGE_H - 20);
    }

    doc.save(`Daily_Fabric_Issue_Summary_${startDate}_to_${endDate}.pdf`);
  };

  return (
    <div className="wip-report-container">
      {/* Styles Injection for Factory Suite Pro Theme */}
      <style>{`
        .gradient-title {
          background: linear-gradient(135deg, #1d4ed8 0%, #3b82f6 100%);
          -webkit-background-clip: text;
          -webkit-text-fill-color: transparent;
          font-weight: 900;
          letter-spacing: -0.8px;
        }
        
        .premium-card {
          background: #ffffff !important;
          border: 1px solid rgba(37, 99, 235, 0.12) !important;
          box-shadow: 0 4px 20px -2px rgba(37, 99, 235, 0.04) !important;
          border-radius: 12px !important;
          transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }
        
        .premium-card:hover {
          transform: translateY(-2px);
          box-shadow: 0 12px 28px -5px rgba(37, 99, 235, 0.1) !important;
          border-color: rgba(37, 99, 235, 0.25) !important;
        }

        .kpi-card-glow {
          position: relative;
          overflow: hidden;
        }

        .kpi-card-glow::before {
          content: '';
          position: absolute;
          top: 0; left: 0; width: 4px; height: 100%;
        }
        
        .kpi-purple::before { background: #2563eb; }
        .kpi-emerald::before { background: #3b82f6; }
        .kpi-amber::before { background: #1d4ed8; }
        .kpi-sky::before { background: #60a5fa; }

        .glow-icon-box {
          display: flex;
          align-items: center;
          justify-content: center;
          width: 44px;
          height: 44px;
          border-radius: 12px;
          transition: transform 0.4s ease;
        }

        .premium-card:hover .glow-icon-box {
          transform: scale(1.1) rotate(4deg);
        }

        .custom-gradient-progress {
          background: linear-gradient(90deg, #2563eb 0%, #60a5fa 100%) !important;
        }

        .wip-report-container {
          padding: 32px;
          background-color: #f8fafc;
          background-image: 
            radial-gradient(at 0% 0%, rgba(99, 102, 241, 0.08) 0px, transparent 50%),
            radial-gradient(at 100% 0%, rgba(236, 72, 153, 0.06) 0px, transparent 50%),
            radial-gradient(at 50% 100%, rgba(16, 185, 129, 0.06) 0px, transparent 50%);
          min-height: 100vh;
          font-family: 'Plus Jakarta Sans', 'Inter', system-ui, -apple-system, sans-serif;
          color: #0f172a;
        }
        
        .header {
          background: linear-gradient(135deg, #1e1b4b 0%, #312e81 40%, #4338ca 100%);
          border-radius: 24px;
          padding: 32px 36px;
          margin-bottom: 28px;
          box-shadow: 0 20px 40px -10px rgba(49, 46, 129, 0.3);
          color: white;
          position: relative;
          overflow: hidden;
        }

        .header::before {
          content: '';
          position: absolute;
          top: -50%;
          right: -20%;
          width: 400px;
          height: 400px;
          background: radial-gradient(circle, rgba(99, 102, 241, 0.25) 0%, transparent 70%);
          pointer-events: none;
        }

        .header-content {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-bottom: 24px;
          flex-wrap: wrap;
          gap: 20px;
        }

        .header-title-box {
          display: flex;
          flex-direction: column;
          gap: 6px;
        }
        
        .header h1 {
          margin: 0;
          font-size: 2.2rem;
          font-weight: 800;
          color: #ffffff;
          display: flex;
          align-items: center;
          gap: 14px;
          letter-spacing: -0.5px;
        }

        .live-status-badge {
          display: inline-flex;
          align-items: center;
          gap: 6px;
          background: rgba(16, 185, 129, 0.2);
          border: 1px solid rgba(16, 185, 129, 0.4);
          color: #34d399;
          padding: 4px 12px;
          border-radius: 9999px;
          font-size: 12px;
          font-weight: 700;
          letter-spacing: 0.5px;
          text-transform: uppercase;
        }

        .live-dot {
          width: 8px;
          height: 8px;
          background-color: #34d399;
          border-radius: 50%;
          box-shadow: 0 0 8px #34d399;
          animation: pulse-dot 1.5s infinite;
        }

        @keyframes pulse-dot {
          0% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(52, 211, 153, 0.7); }
          70% { transform: scale(1); box-shadow: 0 0 0 8px rgba(52, 211, 153, 0); }
          100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(52, 211, 153, 0); }
        }
        
        .header-subtitle {
          margin: 0;
          font-size: 0.95rem;
          color: #c7d2fe;
          font-weight: 400;
        }
        
        .header-controls {
          display: flex;
          gap: 10px;
          flex-wrap: wrap;
        }
        
        .icon-btn {
          background: rgba(255, 255, 255, 0.12);
          border: 1px solid rgba(255, 255, 255, 0.2);
          border-radius: 14px;
          padding: 10px 18px;
          display: inline-flex;
          align-items: center;
          justify-content: center;
          gap: 8px;
          cursor: pointer;
          transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
          color: #ffffff;
          font-size: 13px;
          font-weight: 600;
          backdrop-filter: blur(8px);
        }
        
        .icon-btn:hover {
          background: rgba(255, 255, 255, 0.22);
          transform: translateY(-2px);
          box-shadow: 0 8px 16px rgba(0, 0, 0, 0.15);
        }

        .btn-excel {
          background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
          border: none !important;
          box-shadow: 0 4px 12px rgba(16, 185, 129, 0.3) !important;
        }

        .btn-pdf {
          background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%) !important;
          border: none !important;
          box-shadow: 0 4px 12px rgba(239, 68, 68, 0.3) !important;
        }

        .btn-back {
          background: rgba(255, 255, 255, 0.15) !important;
        }

        .btn-refresh {
          background: linear-gradient(135deg, #6366f1 0%, #4f46e5 100%) !important;
          border: none !important;
          box-shadow: 0 4px 12px rgba(99, 102, 241, 0.3) !important;
        }

        .summary-cards {
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          gap: 18px;
          background: rgba(255, 255, 255, 0.1);
          border: 1px solid rgba(255, 255, 255, 0.2);
          padding: 16px 20px;
          border-radius: 20px;
          backdrop-filter: blur(12px);
        }
        
        .summary-card {
          background: rgba(255, 255, 255, 0.95);
          border-radius: 16px;
          padding: 18px 22px;
          box-shadow: 0 4px 15px rgba(0, 0, 0, 0.05);
          transition: all 0.25s ease;
          border-left: 5px solid #6366f1;
          display: flex;
          flex-direction: column;
          justify-content: center;
        }
        
        .summary-card:hover {
          transform: translateY(-3px);
          box-shadow: 0 10px 20px rgba(0, 0, 0, 0.08);
        }
        
        .summary-card.total-lots { border-left-color: #6366f1; }
        .summary-card.total-pieces { border-left-color: #10b981; }
        .summary-card.total-supervisors { border-left-color: #a855f7; }
        .summary-card.avg-aging { border-left-color: #f59e0b; }

        .card-label {
          font-size: 11px;
          font-weight: 800;
          color: #64748b;
          text-transform: uppercase;
          letter-spacing: 0.5px;
          margin-bottom: 4px;
        }
        
        .card-value {
          font-size: 26px;
          font-weight: 800;
          color: #0f172a;
        }

        .filter-section {
          background: #ffffff;
          border-radius: 20px;
          padding: 24px;
          box-shadow: 0 10px 30px -5px rgba(0, 0, 0, 0.05);
          margin-bottom: 28px;
          border: 1px solid #e2e8f0;
        }

        .custom-table-bordered {
          border-collapse: collapse !important;
          width: 100% !important;
          border: 2px solid #cbd5e1 !important;
        }

        .custom-table-bordered th {
          background: #1e293b !important;
          color: #ffffff !important;
          font-weight: 800 !important;
          border: 1px solid #475569 !important;
          padding: 12px 14px !important;
          font-size: 11px !important;
          text-transform: uppercase;
        }

        .custom-table-bordered td {
          color: #0f172a !important;
          font-weight: 600 !important;
          border: 1px solid #cbd5e1 !important;
          padding: 12px 14px !important;
          background: #ffffff !important;
        }

        .custom-table-bordered tr:nth-child(even) td {
          background: #f8fafc !important;
        }

        .spin {
          animation: spin 1s linear infinite;
        }
        @keyframes spin {
          100% { transform: rotate(360deg); }
        }
      `}</style>

      {/* System Theme Hero Header */}
      <div className="header">
        <div className="header-content">
          <div className="header-title-box">
            <h1>
              <span>🧵 Daily Fabric Issue Analytics</span>
              <span className="live-status-badge">
                <span className="live-dot"></span> LIVE
              </span>
            </h1>
            <p className="header-subtitle">
              Visual distribution matrix mapping overall volume and weights across tables & fabric styles.
            </p>
          </div>
          <div className="header-controls">
            <button className="icon-btn btn-back" onClick={() => history.push('/dashboard')}>
              <ArrowLeft size={16} /> Dashboard
            </button>
            <button
              className={`icon-btn ${isPendingCuttingFilter ? 'btn-amber' : ''}`}
              onClick={handlePendingToggle}
              style={{
                background: isPendingCuttingFilter
                  ? 'linear-gradient(135deg, #f59e0b 0%, #d97706 100%)'
                  : 'rgba(255, 255, 255, 0.15)',
                border: isPendingCuttingFilter ? 'none' : '1px solid rgba(255, 255, 255, 0.2)',
                boxShadow: isPendingCuttingFilter ? '0 4px 12px rgba(245, 158, 11, 0.35)' : 'none'
              }}
              title="Filter Fabric Issued lots with Cutting Entry Pending in Index Sheet from 1 July 2026"
            >
              <Clock size={16} /> {isPendingCuttingFilter ? 'SHOW ALL LOTS' : 'ENTRY PENDING LOT'}
            </button>
            <button className="icon-btn btn-refresh" onClick={isPendingCuttingFilter ? fetchPendingCuttingLots : fetchReport}>
              <RefreshCw size={16} className={loading ? "spin" : ""} /> Reload
            </button>
            <button className="icon-btn btn-excel" onClick={exportToExcel} disabled={loading || filteredData.length === 0}>
              <Download size={16} /> Export Excel
            </button>
            <button className="icon-btn btn-pdf" onClick={exportToPdf} disabled={loading || filteredData.length === 0}>
              <FileText size={16} /> Export PDF
            </button>
          </div>
        </div>

        {/* Integrated Glassmorphic KPI Summary Cards */}
        <div className="summary-cards">
          <div className="summary-card total-lots">
            <span className="card-label">ISSUED ROLLS</span>
            <span className="card-value">{stats.totalRolls}</span>
          </div>
          <div className="summary-card total-pieces">
            <span className="card-label">TOTAL WEIGHT</span>
            <span className="card-value">{stats.totalWeight.toFixed(1)} <small style={{ fontSize: 13, color: '#64748b' }}>KG</small></span>
          </div>
          <div className="summary-card total-supervisors">
            <span className="card-label">UNIQUE LOTS</span>
            <span className="card-value">{stats.uniqueLots.size}</span>
          </div>
          <div className="summary-card avg-aging">
            <span className="card-label">CUTTING TABLES</span>
            <span className="card-value">{stats.activeTables.size}</span>
          </div>
        </div>
      </div>

      {/* Warning Banner when Entry Pending Filter is Active */}
      {isPendingCuttingFilter && (
        <div style={{
          background: 'linear-gradient(135deg, rgba(254, 243, 199, 0.9) 0%, rgba(253, 230, 138, 0.95) 100%)',
          border: '1.5px solid #f59e0b',
          borderRadius: '16px',
          padding: '16px 22px',
          marginBottom: '24px',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          flexWrap: 'wrap',
          gap: 12,
          boxShadow: '0 8px 20px rgba(245, 158, 11, 0.15)'
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 14 }}>
            <div style={{ width: 42, height: 42, borderRadius: 10, background: '#f59e0b', display: 'flex', alignItems: 'center', justify: 'center', color: '#ffffff', flexShrink: 0, boxShadow: '0 4px 10px rgba(245, 158, 11, 0.3)' }}>
              <Clock size={22} />
            </div>
            <div>
              <div style={{ fontWeight: 850, fontSize: '15px', color: '#92400e' }}>
                ⚠️ ENTRY PENDING CUTTING LOTS (From 1 July 2026)
              </div>
              <div style={{ fontSize: '13px', fontWeight: 600, color: '#b45309', marginTop: 2 }}>
                Found <strong>{pendingCuttingCount}</strong> issued lots where Fabric was issued but Cutting Entry is <strong>PENDING</strong> in Index Sheet.
              </div>
            </div>
          </div>
          <button
            onClick={handlePendingToggle}
            style={{
              background: '#ffffff',
              border: '1.5px solid #d97706',
              color: '#b45309',
              borderRadius: '10px',
              padding: '8px 16px',
              fontSize: '12px',
              fontWeight: 800,
              cursor: 'pointer',
              boxShadow: '0 2px 6px rgba(217, 119, 6, 0.15)'
            }}
          >
            Show All Issued Lots
          </button>
        </div>
      )}

      {/* Filter Section */}
      <div className="filter-section" style={{ width: '100%', boxSizing: 'border-box' }}>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: 16, alignItems: 'flex-end', width: '100%' }}>

          {/* Start Date */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
            <label style={{ fontSize: 11, fontWeight: 800, color: '#475569', textTransform: 'uppercase', letterSpacing: '0.6px' }}>Start Date</label>
            <div style={{ position: 'relative' }}>
              <input
                type="date"
                style={{ width: '100%', paddingLeft: 34, height: 42, borderRadius: 10, border: '1.5px solid #cbd5e1', fontSize: '13px', fontWeight: 700, outline: 'none', boxSizing: 'border-box' }}
                value={startDate}
                onChange={e => setStartDate(e.target.value)}
              />
              <Calendar size={15} style={{ position: 'absolute', left: 12, top: '50%', transform: 'translateY(-50%)', color: '#6366f1' }} />
            </div>
          </div>

          {/* End Date */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
            <label style={{ fontSize: 11, fontWeight: 800, color: '#475569', textTransform: 'uppercase', letterSpacing: '0.6px' }}>End Date</label>
            <div style={{ position: 'relative' }}>
              <input
                type="date"
                style={{ width: '100%', paddingLeft: 34, height: 42, borderRadius: 10, border: '1.5px solid #cbd5e1', fontSize: '13px', fontWeight: 700, outline: 'none', boxSizing: 'border-box' }}
                value={endDate}
                onChange={e => setEndDate(e.target.value)}
              />
              <Calendar size={15} style={{ position: 'absolute', left: 12, top: '50%', transform: 'translateY(-50%)', color: '#6366f1' }} />
            </div>
          </div>

          {/* Text Search */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 6, gridColumn: 'span 2' }}>
            <label style={{ fontSize: 11, fontWeight: 800, color: '#475569', textTransform: 'uppercase', letterSpacing: '0.6px' }}>Search Filter</label>
            <div style={{ position: 'relative' }}>
              <input
                type="text"
                placeholder="Search summaries by lot, shade, fabric name..."
                style={{ width: '100%', paddingLeft: 38, height: 42, borderRadius: 10, border: '1.5px solid #cbd5e1', fontSize: '13px', fontWeight: 600, outline: 'none', boxSizing: 'border-box' }}
                value={searchTerm}
                onChange={e => setSearchTerm(e.target.value)}
              />
              <Search size={15} style={{ position: 'absolute', left: 12, top: '50%', transform: 'translateY(-50%)', color: '#64748b' }} />
            </div>
          </div>

        </div>
      </div>

      {/* Modern High-Quality Chart 1: Daily Issuance Trend (Area Chart) */}
      {filteredData.length > 0 && (
        <div className="card premium-card" style={{ overflow: 'hidden' }}>
          <div className="card-header" style={{ padding: '16px 20px', borderBottom: '2px solid #e2e8f0', background: '#f8fafc', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <div style={{ width: 4, height: 14, background: '#2563eb', borderRadius: 2 }} />
              <div style={{ fontSize: '13px', fontWeight: 850, color: '#0f172a', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Issuance Volume & Weight Trend</div>
            </div>
            <span style={{ fontSize: '11px', color: '#2563eb', fontWeight: 800 }}>Daily Performance Curve</span>
          </div>
          <div className="card-body" style={{ padding: '24px 20px 10px 10px' }}>
            <ResponsiveContainer width="100%" height={280}>
              <ComposedChart data={trendChartData}>
                <defs>
                  <linearGradient id="weightTrendGrad" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#10b981" stopOpacity={0.24} />
                    <stop offset="95%" stopColor="#10b981" stopOpacity={0.0} />
                  </linearGradient>
                  <linearGradient id="rollsTrendGrad" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#2563eb" stopOpacity={0.08} />
                    <stop offset="95%" stopColor="#2563eb" stopOpacity={0.0} />
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" vertical={false} />
                <XAxis
                  dataKey="formattedDate"
                  tick={{ fontSize: 10, fill: '#64748b', fontWeight: 600 }}
                  axisLine={{ stroke: '#cbd5e1' }}
                  tickLine={{ stroke: '#cbd5e1' }}
                />
                <YAxis
                  yAxisId="left"
                  tick={{ fontSize: 10, fill: '#64748b', fontWeight: 600 }}
                  axisLine={false}
                  tickLine={false}
                  label={{ value: 'Rolls Issued', angle: -90, position: 'insideLeft', offset: 0, fill: '#2563eb', fontSize: 11, fontWeight: 700 }}
                />
                <YAxis
                  yAxisId="right"
                  orientation="right"
                  tick={{ fontSize: 10, fill: '#64748b', fontWeight: 600 }}
                  axisLine={false}
                  tickLine={false}
                  label={{ value: 'Weight (KG)', angle: 90, position: 'insideRight', offset: 0, fill: '#10b981', fontSize: 11, fontWeight: 700 }}
                />
                <Tooltip content={<CustomTooltip />} cursor={{ stroke: 'rgba(37, 99, 235, 0.08)', strokeWidth: 2 }} />
                <Legend
                  verticalAlign="top"
                  height={40}
                  iconType="circle"
                  iconSize={8}
                  wrapperStyle={{ fontSize: '11px', fontWeight: 700, color: '#1e293b', paddingBottom: '10px' }}
                />
                <Area
                  yAxisId="right"
                  type="monotone"
                  dataKey="weight"
                  name="Weight (KG)"
                  fill="url(#weightTrendGrad)"
                  stroke="#10b981"
                  strokeWidth={2.5}
                  dot={{ r: 3, fill: '#10b981', strokeWidth: 0 }}
                  activeDot={{ r: 5, stroke: '#10b981', strokeWidth: 1.5, fill: '#ffffff' }}
                />
                <Line
                  yAxisId="left"
                  type="monotone"
                  dataKey="rolls"
                  name="Rolls Issued"
                  stroke="#2563eb"
                  strokeWidth={3}
                  dot={{ r: 4, stroke: '#2563eb', strokeWidth: 2, fill: '#ffffff' }}
                  activeDot={{ r: 6, stroke: '#2563eb', strokeWidth: 2.5, fill: '#ffffff' }}
                />
              </ComposedChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}

      {/* Main Aggregated Summaries (Dual Cards side-by-side with proper borders) */}
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1.2fr', gap: 24 }}>

        {/* Table-wise Summary Card */}
        <div className="card premium-card" style={{ overflow: 'hidden', padding: 0 }}>
          <div className="card-header" style={{ padding: '16px 20px', borderBottom: '2px solid #cbd5e1', background: '#f1f5f9', display: 'flex', alignItems: 'center', gap: 8 }}>
            <div style={{ width: 4, height: 14, background: '#2563eb', borderRadius: 2 }} />
            <div style={{ fontSize: '13px', fontWeight: 850, color: '#0f172a', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Table-wise Summary</div>
          </div>
          <div className="card-body" style={{ padding: 16 }}>
            {loading ? (
              <div style={{ padding: 40, textAlign: 'center' }}><div className="spinner" style={{ margin: '0 auto' }} /></div>
            ) : tableSummary.length === 0 ? (
              <div style={{ padding: 30, textAlign: 'center', color: '#0f172a', fontSize: 13 }}>No records found.</div>
            ) : (
              <div className="table-wrap" style={{ border: 'none' }}>
                <table className="custom-table-bordered">
                  <thead>
                    <tr>
                      <th style={{ textAlign: 'left', minWidth: '100px' }}>Table</th>
                      <th style={{ textAlign: 'left', minWidth: '220px' }}>Lot Numbers</th>
                      <th style={{ textAlign: 'right' }}>Rolls</th>
                      <th style={{ textAlign: 'right' }}>Weight (KG)</th>
                      <th style={{ textAlign: 'center' }}>Share</th>
                    </tr>
                  </thead>
                  <tbody>
                    {tableSummary.map(item => (
                      <tr key={item.name}>
                        <td style={{ fontWeight: 800 }}>
                          <span style={{ fontSize: '12px', color: '#1e3a8a' }}>
                            {item.name}
                          </span>
                        </td>
                        <td>
                          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, minWidth: '220px' }}>
                            {Array.from(item.lots).sort().map(lot => (
                              <span key={lot} style={{ background: '#eff6ff', color: '#1d4ed8', border: '1px solid #bfdbfe', borderRadius: 6, padding: '4px 9px', fontSize: '11px', fontWeight: 750, letterSpacing: '0.3px', boxShadow: '0 1px 2px rgba(0,0,0,0.03)' }}>
                                {lot}
                              </span>
                            ))}
                          </div>
                        </td>
                        <td style={{ fontWeight: 900, textAlign: 'right', fontSize: '13px' }}>{item.rolls}</td>
                        <td style={{ fontWeight: 900, textAlign: 'right', fontSize: '13px', color: '#10b981' }}>{item.weight.toFixed(1)}</td>
                        <td style={{ width: '100px' }}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: 6, justifyContent: 'center' }}>
                            <div style={{ flex: 1, height: 6, background: '#e2e8f0', borderRadius: 3, overflow: 'hidden' }}>
                              <div className="custom-gradient-progress" style={{ width: `${item.percentage}%`, height: '100%', borderRadius: 3 }} />
                            </div>
                            <span style={{ fontSize: '11px', fontWeight: 800, color: '#2563eb' }}>{item.percentage}%</span>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>

        {/* Fabric-wise Summary Card */}
        <div className="card premium-card" style={{ overflow: 'hidden', padding: 0 }}>
          <div className="card-header" style={{ padding: '16px 20px', borderBottom: '2px solid #cbd5e1', background: '#f1f5f9', display: 'flex', alignItems: 'center', gap: 8 }}>
            <div style={{ width: 4, height: 14, background: '#2563eb', borderRadius: 2 }} />
            <div style={{ fontSize: '13px', fontWeight: 850, color: '#0f172a', textTransform: 'uppercase', letterSpacing: '0.5px' }}>Fabric-wise Summary</div>
          </div>
          <div className="card-body" style={{ padding: 16 }}>
            {loading ? (
              <div style={{ padding: 40, textAlign: 'center' }}><div className="spinner" style={{ margin: '0 auto' }} /></div>
            ) : fabricSummary.length === 0 ? (
              <div style={{ padding: 30, textAlign: 'center', color: '#0f172a', fontSize: 13 }}>No fabric records found.</div>
            ) : (
              <div className="table-wrap" style={{ border: 'none' }}>
                <table className="custom-table-bordered">
                  <thead>
                    <tr>
                      <th style={{ textAlign: 'left', minWidth: '130px' }}>Fabric Description</th>
                      <th style={{ textAlign: 'left', minWidth: '220px' }}>Lot Numbers</th>
                      <th style={{ textAlign: 'right' }}>Rolls</th>
                      <th style={{ textAlign: 'right' }}>Weight (KG)</th>
                      <th style={{ textAlign: 'center' }}>Share</th>
                    </tr>
                  </thead>
                  <tbody>
                    {fabricSummary.map(item => (
                      <tr key={item.name}>
                        <td style={{ fontWeight: 850, maxWidth: '150px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }} title={item.name}>
                          {item.name}
                        </td>
                        <td>
                          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, minWidth: '220px' }}>
                            {Array.from(item.lots).sort().map(lot => (
                              <span key={lot} style={{ background: '#eff6ff', color: '#1d4ed8', border: '1.5px solid #3b82f6', borderRadius: 6, padding: '4px 9px', fontSize: '11px', fontWeight: 750, letterSpacing: '0.3px', boxShadow: '0 1px 2px rgba(0,0,0,0.03)' }}>
                                {lot}
                              </span>
                            ))}
                          </div>
                        </td>
                        <td style={{ fontWeight: 900, textAlign: 'right', fontSize: '13px' }}>{item.rolls}</td>
                        <td style={{ fontWeight: 900, textAlign: 'right', fontSize: '13px', color: '#10b981' }}>{item.weight.toFixed(1)}</td>
                        <td style={{ width: '100px' }}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: 6, justifyContent: 'center' }}>
                            <div style={{ flex: 1, height: 6, background: '#e2e8f0', borderRadius: 3, overflow: 'hidden' }}>
                              <div className="custom-gradient-progress" style={{ width: `${item.percentage}%`, height: '100%', borderRadius: 3 }} />
                            </div>
                            <span style={{ fontSize: '11px', fontWeight: 800, color: '#2563eb' }}>{item.percentage}%</span>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>

      </div>

      {/* Sub-Charts Section: Table ComposedChart + Fabric Doughnut Chart */}
      {filteredData.length > 0 && (
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1.2fr', gap: 24 }}>

          {/* Chart 2: Table-wise Composed Chart (Vibrant Bars + curved Line overlay) */}
          <div className="card premium-card" style={{ overflow: 'hidden' }}>
            <div className="card-header" style={{ padding: '16px 20px', borderBottom: '2px solid #e2e8f0', background: '#f8fafc' }}>
              <div className="card-title" style={{ fontSize: '13px', fontWeight: 850, textTransform: 'uppercase', color: '#0f172a', letterSpacing: '0.5px' }}>Table Performance (Rolls & Weights)</div>
            </div>
            <div className="card-body" style={{ padding: '20px 10px 10px 10px' }}>
              <ResponsiveContainer width="100%" height={230}>
                <ComposedChart data={tableChartData}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" vertical={false} />
                  <XAxis
                    dataKey="name"
                    tick={{ fontSize: 10, fill: '#64748b', fontWeight: 600 }}
                    axisLine={{ stroke: '#cbd5e1' }}
                    tickLine={{ stroke: '#cbd5e1' }}
                  />
                  <YAxis
                    yAxisId="left"
                    tick={{ fontSize: 10, fill: '#64748b', fontWeight: 600 }}
                    axisLine={false}
                    tickLine={false}
                  />
                  <YAxis
                    yAxisId="right"
                    orientation="right"
                    tick={{ fontSize: 10, fill: '#64748b', fontWeight: 600 }}
                    axisLine={false}
                    tickLine={false}
                  />
                  <Tooltip content={<CustomTooltip />} cursor={{ fill: 'rgba(37, 99, 235, 0.02)' }} />
                  <Legend
                    verticalAlign="top"
                    height={36}
                    iconType="circle"
                    iconSize={8}
                    wrapperStyle={{ fontSize: '11px', fontWeight: 700, color: '#1e293b' }}
                  />
                  <Bar yAxisId="left" dataKey="rolls" name="Rolls Issued" radius={[6, 6, 0, 0]} maxBarSize={28}>
                    {tableChartData.map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={CHART_COLORS[index % CHART_COLORS.length]} />
                    ))}
                  </Bar>
                  <Line
                    yAxisId="right"
                    type="monotone"
                    dataKey="weight"
                    name="Weight (KG)"
                    stroke="#e11d48"
                    strokeWidth={3}
                    dot={{ r: 4, stroke: '#e11d48', strokeWidth: 2, fill: '#ffffff' }}
                    activeDot={{ r: 6, stroke: '#e11d48', strokeWidth: 2.5, fill: '#ffffff' }}
                  />
                </ComposedChart>
              </ResponsiveContainer>
            </div>
          </div>

          {/* Chart 3: Fabric-wise Doughnut (Pie) Chart (Colorful Slices) */}
          <div className="card premium-card" style={{ overflow: 'hidden' }}>
            <div className="card-header" style={{ padding: '16px 20px', borderBottom: '2px solid #e2e8f0', background: '#f8fafc' }}>
              <div className="card-title" style={{ fontSize: '13px', fontWeight: 850, textTransform: 'uppercase', color: '#0f172a', letterSpacing: '0.5px' }}>Fabric Share Ratio</div>
            </div>
            <div className="card-body" style={{ padding: '10px', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <div style={{ flex: 1.2, height: 230 }}>
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie
                      data={fabricChartData}
                      cx="50%"
                      cy="50%"
                      innerRadius={50}
                      outerRadius={80}
                      paddingAngle={4}
                      dataKey="value"
                      nameKey="name"
                    >
                      {fabricChartData.map((entry, index) => (
                        <Cell key={`cell-${index}`} fill={CHART_COLORS[index % CHART_COLORS.length]} stroke="#ffffff" strokeWidth={1.5} />
                      ))}
                    </Pie>
                    <Tooltip content={<CustomTooltip />} />
                  </PieChart>
                </ResponsiveContainer>
              </div>
              {/* Custom aligned side legend */}
              <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 8, paddingLeft: 12 }}>
                {fabricChartData.map((entry, idx) => (
                  <div key={entry.name} style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '11px', fontWeight: 700 }}>
                    <span style={{ width: 8, height: 8, borderRadius: '50%', background: CHART_COLORS[idx % CHART_COLORS.length], flexShrink: 0 }} />
                    <span style={{ color: '#0f172a', textOverflow: 'ellipsis', overflow: 'hidden', whiteSpace: 'nowrap', maxWidth: '120px' }} title={entry.name}>
                      {entry.name}
                    </span>
                    <span style={{ color: '#475569', marginLeft: 'auto' }}>({entry.value})</span>
                  </div>
                ))}
              </div>
            </div>
          </div>

        </div>
      )}

    </div>
  );
}
