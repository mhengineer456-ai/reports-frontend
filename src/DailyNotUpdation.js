import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

const DailyNotUpdation = () => {
  // Google Sheets configuration
  const SPREADSHEET_ID = '1Hj3JeJEKB43aYYWv8gk2UhdU6BWuEQfCg5pBlTdBMNA';
  const API_KEY = 'AIzaSyAomDFBkOySlIxKWSKGHe6ATv9gvaBr7uk';
  const RANGE = 'Index!A1:Z';

  // State management
  const [data, setData] = useState([]);
  const [headers, setHeaders] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [filters, setFilters] = useState({
    lotNumber: '',
    fabric: '',
    garmentType: '',
    partyName: '',
    brand: '',
    season: '',
    wipStatus: '',
    supervisor: ''
  });
  const [searchTerm, setSearchTerm] = useState('');
  const [stats, setStats] = useState({
    totalEntries: 0,
    pendingWIP: 0,
    completed: 0,
    totalStyles: 0,
    notUpdatedToday: 0
  });
  const [todayDate, setTodayDate] = useState('');
  const [exporting, setExporting] = useState(false);
  const [selectedRows, setSelectedRows] = useState([]);
  const [sortConfig, setSortConfig] = useState({ key: null, direction: 'asc' });
  const [viewMode, setViewMode] = useState('table'); // 'table' or 'card'
  const tableRef = useRef(null);

  // Normalize supervisor name
  const normalizeSupervisor = (name) => {
    if (!name || typeof name !== 'string') return '';
    return name
      .toLowerCase()
      .split(' ')
      .map(word => word.charAt(0).toUpperCase() + word.slice(1))
      .join(' ')
      .trim();
  };

  // Helper function to check if item belongs to Dummy, Dushyant, or Jain Hosiery
  const isExcludedItem = (item) => {
    if (!item) return false;
    const sup = String(item['NormalizedSupervisor'] || item['Supervisor'] || '').trim().toLowerCase();
    const party = String(item['PARTY NAME'] || '').trim().toLowerCase();
    const brand = String(item['BRAND'] || '').trim().toLowerCase();

    const excludeKeys = ['dummy', 'dushyant', 'jain hosiery', 'JAINHOSIERY'];
    return excludeKeys.some((ex) => sup.includes(ex) || party.includes(ex) || brand.includes(ex));
  };

  // Set today's date on component mount
  useEffect(() => {
    const today = new Date();
    const formattedDate = today.toLocaleDateString('en-IN', {
      day: '2-digit',
      month: 'short',
      year: 'numeric'
    });
    setTodayDate(formattedDate);
  }, []);

  // Helper function to parse JSON status data
  const parseStatusData = (statusString) => {
    try {
      if (!statusString || statusString.trim() === '') return [];
      const cleanedString = statusString.trim();
      return JSON.parse(cleanedString);
    } catch (error) {
      console.error('Error parsing status data:', error);
      return [];
    }
  };

  // Helper function to get latest status with safe defaults
  const getLatestStatus = (statusArray) => {
    if (!Array.isArray(statusArray) || statusArray.length === 0) {
      return { status: 'Not Started', timestamp: null, remarks: '' };
    }
    try {
      const sorted = [...statusArray].sort((a, b) =>
        new Date(b.timestamp || 0) - new Date(a.timestamp || 0)
      );
      return {
        status: sorted[0]?.status || 'Not Started',
        timestamp: sorted[0]?.timestamp || null,
        remarks: sorted[0]?.remarks || ''
      };
    } catch (error) {
      return { status: 'Not Started', timestamp: null, remarks: '' };
    }
  };

  // Helper function to check if timestamp is from today
  const isFromToday = (timestamp) => {
    if (!timestamp) return false;
    try {
      const date = new Date(timestamp);
      const today = new Date();
      return (
        date.getDate() === today.getDate() &&
        date.getMonth() === today.getMonth() &&
        date.getFullYear() === today.getFullYear()
      );
    } catch (error) {
      return false;
    }
  };

  // Helper function to format date from timestamp
  const formatDateFromTimestamp = (timestamp) => {
    if (!timestamp) return 'Never';
    try {
      const date = new Date(timestamp);
      return date.toLocaleDateString('en-IN', {
        day: '2-digit',
        month: 'short',
        year: 'numeric'
      });
    } catch (error) {
      return 'Invalid Date';
    }
  };

  // Safe string comparison with null checks
  const safeStringCompare = (str1, str2) => {
    const s1 = String(str1 || '').toLowerCase();
    const s2 = String(str2 || '').toLowerCase();
    return s1 === s2;
  };

  // Safe string contains with null checks
  const safeStringContains = (str, search) => {
    const s1 = String(str || '').toLowerCase();
    const s2 = String(search || '').toLowerCase();
    return s1.includes(s2);
  };

  // Check if both Date of Issue and Supervisor are present
  const hasRequiredData = (item) => {
    const dateOfIssue = item['Date of Issue'] || '';
    const supervisor = item['Supervisor'] || '';
    return dateOfIssue.trim() !== '' && supervisor.trim() !== '';
  };
  const handleBack = () => {
    window.history.back();
  };
  // Fetch data from Google Sheets
  const fetchData = async () => {
    setLoading(true);
    setError(null);
    setSelectedRows([]);

    try {
      const response = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${RANGE}?key=${API_KEY}`
      );
      const result = await response.json();

      if (result.values && result.values.length > 0) {
        const [headerRow, ...rows] = result.values;
        setHeaders(headerRow);

        const formattedData = rows.map((row, index) => {
          const item = {};
          headerRow.forEach((header, colIndex) => {
            item[header] = row[colIndex] || '';
          });

          // Normalize supervisor name
          if (item['Supervisor']) {
            item['NormalizedSupervisor'] = normalizeSupervisor(item['Supervisor']);
          } else {
            item['NormalizedSupervisor'] = '';
          }

          item['parsedWIPStatus'] = parseStatusData(item['WIP Status']);
          item['parsedCompletedStatus'] = parseStatusData(item['Completed Status']);

          item['latestWIPStatus'] = getLatestStatus(item['parsedWIPStatus']);
          item['latestCompletedStatus'] = getLatestStatus(item['parsedCompletedStatus']);

          return item;
        });

        setData(formattedData);
      } else {
        setError('No data found in sheet');
      }
    } catch (err) {
      console.error('Failed to fetch data:', err);
      setError('Failed to fetch data from Google Sheets');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchData();
  }, []);

  // Calculate statistics with safe checks
  useEffect(() => {
    if (data.length > 0 && todayDate) {
      const validData = data.filter(item => {
        const completedStatus = item['latestCompletedStatus']?.status || '';
        const isCompleted = safeStringCompare(completedStatus, 'complete lot');
        const hasRequired = hasRequiredData(item);
        return !isCompleted && hasRequired && !isExcludedItem(item);
      });

      const notUpdatedToday = validData.filter(item => {
        const latestWIP = item['latestWIPStatus'] || {};
        return !isFromToday(latestWIP.timestamp);
      });

      const totalEntries = notUpdatedToday.length;
      const pendingWIP = notUpdatedToday.filter(item => {
        const wipStatus = item['latestWIPStatus']?.status || '';
        return safeStringContains(wipStatus, 'pending');
      }).length;

      const completed = data.filter(item => {
        const completedStatus = item['latestCompletedStatus']?.status || '';
        return safeStringCompare(completedStatus, 'complete lot');
      }).length;

      const uniqueStyles = [...new Set(
        notUpdatedToday.map(item => item['Style'] || '').filter(Boolean)
      )].length;

      setStats({
        totalEntries,
        pendingWIP,
        completed,
        totalStyles: uniqueStyles,
        notUpdatedToday: notUpdatedToday.length
      });
    }
  }, [data, todayDate]);

  // Calculate Supervisor-wise Updation Summary Report
  const supervisorSummaryData = useMemo(() => {
    if (!data || data.length === 0) return [];

    const map = {};

    data.forEach((item) => {
      // Exclude complete lots
      const completedStatus = item['latestCompletedStatus']?.status || '';
      const isCompleted = safeStringCompare(completedStatus, 'complete lot');
      if (isCompleted) return;

      // Ensure required data & exclude Dummy, Dushyant, Jain Hosiery
      if (!hasRequiredData(item) || isExcludedItem(item)) return;

      const supName = (item['NormalizedSupervisor'] || item['Supervisor'] || '').trim();
      if (!supName) return;

      if (!map[supName]) {
        map[supName] = {
          supervisorName: supName,
          reportUpdated: 0,
          partialUpdated: 0,
          notUpdated: 0,
          totalLots: 0,
          totalNotUpdatedReport: 0,
          updatedPercentage: 0
        };
      }

      const rec = map[supName];
      rec.totalLots += 1;

      const latestWIP = item['latestWIPStatus'] || {};
      const wipLogs = item['parsedWIPStatus'] || [];

      if (isFromToday(latestWIP.timestamp)) {
        rec.reportUpdated += 1;
      } else if (wipLogs.length > 0) {
        rec.partialUpdated += 1;
        rec.totalNotUpdatedReport += 1;
      } else {
        rec.notUpdated += 1;
        rec.totalNotUpdatedReport += 1;
      }
    });

    return Object.values(map)
      .map((sup) => {
        const pct = sup.totalLots > 0 ? (sup.reportUpdated / sup.totalLots) * 100 : 0;
        return {
          ...sup,
          updatedPercentage: Number(pct.toFixed(1))
        };
      })
      .sort((a, b) => a.supervisorName.localeCompare(b.supervisorName));
  }, [data]);

  // Overall Totals for Supervisor Summary Report
  const supervisorSummaryTotals = useMemo(() => {
    let reportUpdated = 0;
    let partialUpdated = 0;
    let notUpdated = 0;
    let totalLots = 0;
    let totalNotUpdatedReport = 0;

    supervisorSummaryData.forEach((s) => {
      reportUpdated += s.reportUpdated;
      partialUpdated += s.partialUpdated;
      notUpdated += s.notUpdated;
      totalLots += s.totalLots;
      totalNotUpdatedReport += s.totalNotUpdatedReport;
    });

    const updatedPercentage = totalLots > 0 ? Number(((reportUpdated / totalLots) * 100).toFixed(1)) : 0;

    return {
      supervisorName: 'TOTAL SUMMARY',
      reportUpdated,
      partialUpdated,
      notUpdated,
      totalLots,
      totalNotUpdatedReport,
      updatedPercentage
    };
  }, [supervisorSummaryData]);

  // Calculate Garment Type-wise Updation Summary Report
  const garmentTypeSummaryData = useMemo(() => {
    if (!data || data.length === 0) return [];

    const map = {};

    data.forEach((item) => {
      // Exclude complete lots
      const completedStatus = item['latestCompletedStatus']?.status || '';
      const isCompleted = safeStringCompare(completedStatus, 'complete lot');
      if (isCompleted) return;

      // Ensure required data & exclude Dummy, Dushyant, Jain Hosiery
      if (!hasRequiredData(item) || isExcludedItem(item)) return;

      const garmentType = (item['Garment Type'] || 'Unspecified').trim().toUpperCase();

      if (!map[garmentType]) {
        map[garmentType] = {
          garmentType,
          reportUpdated: 0,
          partialUpdated: 0,
          notUpdated: 0,
          totalLots: 0,
          totalNotUpdatedReport: 0,
          updatedPercentage: 0
        };
      }

      const rec = map[garmentType];
      rec.totalLots += 1;

      const latestWIP = item['latestWIPStatus'] || {};
      const wipLogs = item['parsedWIPStatus'] || [];

      if (isFromToday(latestWIP.timestamp)) {
        rec.reportUpdated += 1;
      } else if (wipLogs.length > 0) {
        rec.partialUpdated += 1;
        rec.totalNotUpdatedReport += 1;
      } else {
        rec.notUpdated += 1;
        rec.totalNotUpdatedReport += 1;
      }
    });

    return Object.values(map)
      .map((item) => {
        const pct = item.totalLots > 0 ? (item.reportUpdated / item.totalLots) * 100 : 0;
        return {
          ...item,
          updatedPercentage: Number(pct.toFixed(1))
        };
      })
      .sort((a, b) => a.garmentType.localeCompare(b.garmentType));
  }, [data]);

  // Overall Totals for Garment Type Summary Report
  const garmentTypeSummaryTotals = useMemo(() => {
    let reportUpdated = 0;
    let partialUpdated = 0;
    let notUpdated = 0;
    let totalLots = 0;
    let totalNotUpdatedReport = 0;

    garmentTypeSummaryData.forEach((g) => {
      reportUpdated += g.reportUpdated;
      partialUpdated += g.partialUpdated;
      notUpdated += g.notUpdated;
      totalLots += g.totalLots;
      totalNotUpdatedReport += g.totalNotUpdatedReport;
    });

    const updatedPercentage = totalLots > 0 ? Number(((reportUpdated / totalLots) * 100).toFixed(1)) : 0;

    return {
      garmentType: 'TOTAL SUMMARY',
      reportUpdated,
      partialUpdated,
      notUpdated,
      totalLots,
      totalNotUpdatedReport,
      updatedPercentage
    };
  }, [garmentTypeSummaryData]);

  // Export Supervisor Summary to Excel
  const exportSupervisorSummaryToExcel = () => {
    try {
      const exportRows = supervisorSummaryData.map((s) => ({
        'Supervisor Name': s.supervisorName,
        'Report Updated': s.reportUpdated,
        'Partial Updated': s.partialUpdated,
        'Not Updated': s.notUpdated,
        'Total Lots': s.totalLots,
        'Total Not Updated Report': s.totalNotUpdatedReport,
        'Percentage of Updated Lots (%)': `${s.updatedPercentage}%`
      }));

      exportRows.push({
        'Supervisor Name': 'TOTAL SUMMARY',
        'Report Updated': supervisorSummaryTotals.reportUpdated,
        'Partial Updated': supervisorSummaryTotals.partialUpdated,
        'Not Updated': supervisorSummaryTotals.notUpdated,
        'Total Lots': supervisorSummaryTotals.totalLots,
        'Total Not Updated Report': supervisorSummaryTotals.totalNotUpdatedReport,
        'Percentage of Updated Lots (%)': `${supervisorSummaryTotals.updatedPercentage}%`
      });

      const ws = XLSX.utils.json_to_sheet(exportRows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Supervisor Summary');
      XLSX.writeFile(wb, `Supervisor_WIP_Updation_Summary_${new Date().toISOString().split('T')[0]}.xlsx`);
    } catch (err) {
      console.error('Error exporting Supervisor Summary:', err);
      alert('Failed to export Supervisor Summary Excel');
    }
  };

  // Filter data based on filters and search
  const filteredData = useMemo(() => {
    return data.filter(item => {
      const completedStatus = item['latestCompletedStatus']?.status || '';
      const isCompleted = safeStringCompare(completedStatus, 'complete lot');
      if (isCompleted) return false;

      if (!hasRequiredData(item)) return false;

      const latestWIP = item['latestWIPStatus'] || {};
      const updatedToday = isFromToday(latestWIP.timestamp);
      if (updatedToday) return false;

      const wipStatus = item['latestWIPStatus']?.status || '';
      const normalizedSupervisor = item['NormalizedSupervisor'] || '';

      const matchesFilters =
        (filters.lotNumber === '' || safeStringContains(item['Lot Number'] || '', filters.lotNumber)) &&
        (filters.fabric === '' || safeStringCompare(item['Fabric'] || '', filters.fabric)) &&
        (filters.garmentType === '' || safeStringCompare(item['Garment Type'] || '', filters.garmentType)) &&
        (filters.partyName === '' || safeStringCompare(item['PARTY NAME'] || '', filters.partyName)) &&
        (filters.brand === '' || safeStringCompare(item['BRAND'] || '', filters.brand)) &&
        (filters.season === '' || safeStringCompare(item['SEASON'] || '', filters.season)) &&
        (filters.supervisor === '' || safeStringCompare(normalizedSupervisor, normalizeSupervisor(filters.supervisor))) &&
        (filters.wipStatus === '' || safeStringCompare(wipStatus, filters.wipStatus));

      const matchesSearch = searchTerm === '' ||
        Object.values(item).some(value =>
          safeStringContains(String(value || ''), searchTerm)
        );

      return matchesFilters && matchesSearch;
    });
  }, [data, filters, searchTerm, todayDate]);

  // Handle sort
  const handleSort = (key) => {
    let direction = 'asc';
    if (sortConfig.key === key && sortConfig.direction === 'asc') {
      direction = 'desc';
    }
    setSortConfig({ key, direction });
  };

  // Sort data
  const sortedData = useMemo(() => {
    if (!sortConfig.key) return filteredData;

    return [...filteredData].sort((a, b) => {
      let aValue, bValue;

      // Special handling for nested properties
      if (sortConfig.key === 'latestWIPStatus.timestamp') {
        aValue = a['latestWIPStatus']?.timestamp || '';
        bValue = b['latestWIPStatus']?.timestamp || '';
      } else if (sortConfig.key === 'latestWIPStatus.status') {
        aValue = a['latestWIPStatus']?.status || '';
        bValue = b['latestWIPStatus']?.status || '';
      } else if (sortConfig.key === 'latestWIPStatus.remarks') {
        aValue = a['latestWIPStatus']?.remarks || '';
        bValue = b['latestWIPStatus']?.remarks || '';
      } else if (sortConfig.key === 'NormalizedSupervisor') {
        aValue = normalizeSupervisor(a[sortConfig.key] || '');
        bValue = normalizeSupervisor(b[sortConfig.key] || '');
      } else {
        aValue = a[sortConfig.key] || '';
        bValue = b[sortConfig.key] || '';
      }

      if (aValue < bValue) {
        return sortConfig.direction === 'asc' ? -1 : 1;
      }
      if (aValue > bValue) {
        return sortConfig.direction === 'asc' ? 1 : -1;
      }
      return 0;
    });
  }, [filteredData, sortConfig]);

  // Get unique values for filter dropdowns (with normalized supervisors)
  const getUniqueValues = (field) => {
    if (field === 'WIP Status') {
      const allStatuses = data.flatMap(item =>
        (item['parsedWIPStatus'] || []).map(s => s?.status || '').filter(Boolean)
      );
      return ['All', ...new Set(allStatuses)];
    } else if (field === 'Supervisor') {
      const supervisors = data.map(item => item['NormalizedSupervisor'] || '').filter(Boolean);
      return ['All', ...new Set(supervisors)];
    } else {
      const values = data.map(item => item[field] || '');
      return ['All', ...new Set(values.filter(value => value))];
    }
  };

  // Handle filter changes
  const handleFilterChange = (e) => {
    const { name, value } = e.target;
    setFilters(prev => ({
      ...prev,
      [name]: value === 'All' ? '' : value
    }));
  };

  // Clear all filters
  const clearFilters = () => {
    setFilters({
      lotNumber: '',
      fabric: '',
      garmentType: '',
      partyName: '',
      brand: '',
      season: '',
      wipStatus: '',
      supervisor: ''
    });
    setSearchTerm('');
    setSelectedRows([]);
  };

  // Row selection handlers
  const handleSelectAll = (e) => {
    if (e.target.checked) {
      setSelectedRows(sortedData.map((_, index) => index));
    } else {
      setSelectedRows([]);
    }
  };

  const handleSelectRow = (index) => {
    if (selectedRows.includes(index)) {
      setSelectedRows(selectedRows.filter(i => i !== index));
    } else {
      setSelectedRows([...selectedRows, index]);
    }
  };

  // Export to Excel
  const exportToExcel = () => {
    setExporting(true);

    try {
      const exportData = sortedData.map(item => ({
        'Lot Number': item['Lot Number'],
        'Fabric': item['Fabric'],
        'Garment Type': item['Garment Type'],
        'Style': item['Style'],
        'Party Name': item['PARTY NAME'],
        'Brand': item['BRAND'],
        'Season': item['SEASON'],
        'Current WIP Status': item['latestWIPStatus']?.status || 'Not Started',
        'Last Updated': formatDateFromTimestamp(item['latestWIPStatus']?.timestamp),
        'Remarks': item['latestWIPStatus']?.remarks || '',
        'Date of Issue': item['Date of Issue'],
        'Supervisor': item['NormalizedSupervisor'] || item['Supervisor'],
        'Total WIP Updates': item['parsedWIPStatus']?.length || 0
      }));

      const ws = XLSX.utils.json_to_sheet(exportData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'WIP Status Report');

      const fileName = `WIP_Status_Report_${new Date().toISOString().split('T')[0]}.xlsx`;
      XLSX.writeFile(wb, fileName);
    } catch (err) {
      console.error('Error exporting to Excel:', err);
      alert('Failed to export to Excel. Please try again.');
    } finally {
      setExporting(false);
    }
  };

  // Native Vector PDF Export for Main WIP Report (No html2canvas screenshot)
  const exportToPDF = () => {
    setExporting(true);
    try {
      const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });

      // Clean Title & Header
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(14);
      doc.setTextColor(0, 0, 0);
      doc.text('MH FACTORY SUITE PRO — DAILY STITCHING NOT UPDATION REPORT', 12, 13);

      doc.setFont('helvetica', 'normal');
      doc.setFontSize(8.5);
      doc.setTextColor(60, 60, 60);
      doc.text('Real-time WIP Status Monitoring & Un-updated Production Lots Report', 12, 19);

      // Meta Info Badge (Right Top)
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(8.5);
      doc.setTextColor(0, 0, 0);
      doc.text(`DATE: ${todayDate || new Date().toLocaleDateString('en-IN')}`, 285, 11, { align: 'right' });
      doc.setFont('helvetica', 'normal');
      doc.text(`TOTAL REQUIRING ATTENTION: ${sortedData.length}`, 285, 16, { align: 'right' });

      // Solid Black Separator Line
      doc.setDrawColor(0, 0, 0);
      doc.setLineWidth(0.5);
      doc.line(12, 23, 285, 23);

      const tableRows = sortedData
        .filter(item => !isExcludedItem(item))
        .map((item, idx) => {
          const latestWIP = item['latestWIPStatus'] || {};
          const lastUpdatedDate = formatDateFromTimestamp(latestWIP.timestamp);
          return [
            idx + 1,
            item['Lot Number'] || '-',
            item['Fabric'] || '-',
            item['Garment Type'] || '-',
            item['Style'] || '-',
            item['PARTY NAME'] || '-',
            latestWIP.status || 'Not Started',
            lastUpdatedDate,
            latestWIP.remarks || '-',
            item['Date of Issue'] || '-',
            item['NormalizedSupervisor'] || item['Supervisor'] || '-'
          ];
        });

      autoTable(doc, {
        head: [['#', 'Lot No', 'Fabric', 'Garment', 'Style', 'Party Name', 'WIP Status', 'Last Updated', 'Remarks', 'Issue Date', 'Supervisor']],
        body: tableRows,
        startY: 26,
        margin: { top: 26, bottom: 15, left: 10, right: 10 },
        styles: {
          fontSize: 9,
          fontStyle: 'bold',
          cellPadding: 3,
          textColor: [0, 0, 0],
          valign: 'middle',
          lineColor: [0, 0, 0],
          lineWidth: 0.2
        },
        headStyles: {
          fillColor: [31, 73, 125], // Navy Blue (#1F497D)
          textColor: [255, 255, 255],
          fontStyle: 'bold',
          fontSize: 9.5,
          halign: 'center',
          lineColor: [0, 0, 0],
          lineWidth: 0.2
        },
        columnStyles: {
          0: { halign: 'center', cellWidth: 8 },
          1: { fontStyle: 'bold', cellWidth: 18 },
          2: { cellWidth: 23 },
          3: { cellWidth: 22 },
          4: { fontStyle: 'bold', cellWidth: 22 },
          5: { cellWidth: 30 },
          6: { fontStyle: 'bold', cellWidth: 26 },
          7: { cellWidth: 23 },
          8: { cellWidth: 40 },
          9: { cellWidth: 22 },
          10: { fontStyle: 'bold', cellWidth: 28 }
        },
        didDrawPage: (data) => {
          const str = `Page ${doc.internal.getNumberOfPages()}`;
          doc.setFontSize(8.5);
          doc.setFont('helvetica', 'bold');
          doc.setTextColor(80, 80, 80);
          doc.text(str, data.settings.margin.left, doc.internal.pageSize.height - 8);
          doc.text(
            'CONFIDENTIAL — FOR INTERNAL FACTORY USE ONLY',
            doc.internal.pageSize.width - data.settings.margin.right,
            doc.internal.pageSize.height - 8,
            { align: 'right' }
          );
        }
      });

      doc.save(`Daily_Not_Updation_Report_${new Date().toISOString().split('T')[0]}.pdf`);
    } catch (err) {
      console.error('Error generating PDF:', err);
      alert('Failed to generate PDF. Please try again.');
    } finally {
      setExporting(false);
    }
  };

  // Native Vector PDF Export for Supervisor & Garment Type Summary Report (Fits Single Page, Increased SR NO width)
  const exportSupervisorSummaryToPDF = () => {
    try {
      const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });

      doc.setFont('helvetica', 'bold');
      doc.setFontSize(13);
      doc.setTextColor(0, 0, 0);
      doc.text('MH FACTORY SUITE PRO — DAILY UPDATION SUMMARY REPORT', 10, 10);

      doc.setFont('helvetica', 'bold');
      doc.setFontSize(8.5);
      doc.setTextColor(50, 50, 50);
      doc.text('Supervisor & Garment Type Performance Breakdown for Daily Stitching WIP Updates', 10, 15);

      doc.setFont('helvetica', 'bold');
      doc.setFontSize(8.5);
      doc.setTextColor(0, 0, 0);
      doc.text(`DATE: ${todayDate || new Date().toLocaleDateString('en-IN')}`, 287, 10, { align: 'right' });

      doc.setDrawColor(0, 0, 0);
      doc.setLineWidth(0.5);
      doc.line(10, 18, 287, 18);

      // Table 1: Supervisor / Thekedar Summary
      const supRows = supervisorSummaryData.map((s, i) => {
        const isUpdated = s.reportUpdated > 0;
        return [
          i + 1,
          String(s.supervisorName || '').toUpperCase(),
          isUpdated ? 'Yes' : 'No',
          s.reportUpdated,
          s.notUpdated,
          s.totalLots,
          s.totalNotUpdatedReport,
          `${s.updatedPercentage}%`
        ];
      });

      supRows.push([
        '',
        'TOTAL SUPERVISOR SUMMARY',
        supervisorSummaryTotals.reportUpdated > 0 ? 'Yes' : 'No',
        supervisorSummaryTotals.reportUpdated,
        supervisorSummaryTotals.notUpdated,
        supervisorSummaryTotals.totalLots,
        supervisorSummaryTotals.totalNotUpdatedReport,
        `${supervisorSummaryTotals.updatedPercentage}%`
      ]);

      autoTable(doc, {
        head: [['SR NO', 'Supervisor/Thekedar', 'Status (Yes/No)', 'Report Updated', 'Not Updated', 'Total Lots', 'Total Pending', 'Updated (%)']],
        body: supRows,
        startY: 20,
        margin: { top: 20, bottom: 8, left: 10, right: 10 },
        styles: {
          fontSize: 8.5,
          fontStyle: 'bold',
          cellPadding: 2,
          textColor: [0, 0, 0],
          valign: 'middle',
          lineColor: [0, 0, 0],
          lineWidth: 0.2
        },
        headStyles: {
          fillColor: [31, 73, 125], // Navy Blue (#1F497D)
          textColor: [255, 255, 255],
          fontStyle: 'bold',
          fontSize: 9,
          halign: 'center',
          lineColor: [0, 0, 0],
          lineWidth: 0.2
        },
        columnStyles: {
          0: { halign: 'center', cellWidth: 15 },
          1: { fontStyle: 'bold', cellWidth: 55 },
          2: { halign: 'center', fontStyle: 'bold', cellWidth: 32 },
          3: { halign: 'center', cellWidth: 32 },
          4: { halign: 'center', cellWidth: 32 },
          5: { halign: 'center', fontStyle: 'bold', cellWidth: 32 },
          6: { halign: 'center', fontStyle: 'bold', cellWidth: 35 },
          7: { halign: 'center', fontStyle: 'bold', cellWidth: 34 }
        },
        didParseCell: (data) => {
          if (data.section === 'body') {
            if (data.column.index === 2) {
              const val = String(data.cell.raw || '').trim();
              if (val === 'Yes') {
                data.cell.styles.fillColor = [216, 228, 188];
                data.cell.styles.textColor = [0, 0, 0];
              } else if (val === 'No') {
                data.cell.styles.fillColor = [242, 220, 219];
                data.cell.styles.textColor = [0, 0, 0];
              }
            }
            if (data.row.index === supRows.length - 1) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [230, 235, 245];
            }
          }
        }
      });

      // Table 2: Garment Type Summary (Rendered below Table 1)
      const nextY = (doc.lastAutoTable && doc.lastAutoTable.finalY) ? doc.lastAutoTable.finalY + 6 : 110;

      const garmentRows = garmentTypeSummaryData.map((g, i) => {
        const isUpdated = g.reportUpdated > 0;
        return [
          i + 1,
          String(g.garmentType || '').toUpperCase(),
          isUpdated ? 'Yes' : 'No',
          g.reportUpdated,
          g.notUpdated,
          g.totalLots,
          g.totalNotUpdatedReport,
          `${g.updatedPercentage}%`
        ];
      });

      garmentRows.push([
        '',
        'TOTAL GARMENT SUMMARY',
        garmentTypeSummaryTotals.reportUpdated > 0 ? 'Yes' : 'No',
        garmentTypeSummaryTotals.reportUpdated,
        garmentTypeSummaryTotals.notUpdated,
        garmentTypeSummaryTotals.totalLots,
        garmentTypeSummaryTotals.totalNotUpdatedReport,
        `${garmentTypeSummaryTotals.updatedPercentage}%`
      ]);

      autoTable(doc, {
        head: [['SR NO', 'Garment Type', 'Status (Yes/No)', 'Report Updated', 'Not Updated', 'Total Lots', 'Total Pending', 'Updated (%)']],
        body: garmentRows,
        startY: nextY,
        margin: { top: 10, bottom: 8, left: 10, right: 10 },
        styles: {
          fontSize: 8.5,
          fontStyle: 'bold',
          cellPadding: 2,
          textColor: [0, 0, 0],
          valign: 'middle',
          lineColor: [0, 0, 0],
          lineWidth: 0.2
        },
        headStyles: {
          fillColor: [31, 73, 125], // Navy Blue (#1F497D)
          textColor: [255, 255, 255],
          fontStyle: 'bold',
          fontSize: 9,
          halign: 'center',
          lineColor: [0, 0, 0],
          lineWidth: 0.2
        },
        columnStyles: {
          0: { halign: 'center', cellWidth: 15 },
          1: { fontStyle: 'bold', cellWidth: 55 },
          2: { halign: 'center', fontStyle: 'bold', cellWidth: 32 },
          3: { halign: 'center', cellWidth: 32 },
          4: { halign: 'center', cellWidth: 32 },
          5: { halign: 'center', fontStyle: 'bold', cellWidth: 32 },
          6: { halign: 'center', fontStyle: 'bold', cellWidth: 35 },
          7: { halign: 'center', fontStyle: 'bold', cellWidth: 34 }
        },
        didParseCell: (data) => {
          if (data.section === 'body') {
            if (data.column.index === 2) {
              const val = String(data.cell.raw || '').trim();
              if (val === 'Yes') {
                data.cell.styles.fillColor = [216, 228, 188];
                data.cell.styles.textColor = [0, 0, 0];
              } else if (val === 'No') {
                data.cell.styles.fillColor = [242, 220, 219];
                data.cell.styles.textColor = [0, 0, 0];
              }
            }
            if (data.row.index === garmentRows.length - 1) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [230, 235, 245];
            }
          }
        },
        didDrawPage: (data) => {
          const str = `Page ${doc.internal.getNumberOfPages()}`;
          doc.setFontSize(8);
          doc.setFont('helvetica', 'bold');
          doc.setTextColor(80, 80, 80);
          doc.text(str, data.settings.margin.left, doc.internal.pageSize.height - 5);
          doc.text(
            'CONFIDENTIAL — FOR INTERNAL FACTORY USE ONLY',
            doc.internal.pageSize.width - data.settings.margin.right,
            doc.internal.pageSize.height - 5,
            { align: 'right' }
          );
        }
      });

      doc.save(`Daily_Updation_Summary_${new Date().toISOString().split('T')[0]}.pdf`);
    } catch (err) {
      console.error('Error generating Supervisor Summary PDF:', err);
      alert('Failed to generate PDF. Please try again.');
    }
  };

  // PDF Export matching user's EXACT 3-column screenshot style (Date, Supervisor/Thekedar, Update Yes/No)
  const exportSupervisorStatusPDF = () => {
    try {
      const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });

      const d = new Date();
      const day = String(d.getDate()).padStart(2, '0');
      const month = String(d.getMonth() + 1).padStart(2, '0');
      const year = d.getFullYear();
      const dateStr = `${day}-${month}-${year}`;

      const tableRows = supervisorSummaryData.map((s) => {
        const isUpdated = s.reportUpdated > 0;
        return [
          dateStr,
          String(s.supervisorName || '').toUpperCase(),
          isUpdated ? 'Yes' : 'No'
        ];
      });

      autoTable(doc, {
        head: [['Date', 'Supervisor/Thekedar', 'Update (Yes/No)']],
        body: tableRows,
        startY: 15,
        margin: { top: 15, bottom: 15, left: 35, right: 35 },
        styles: {
          fontSize: 9.5,
          cellPadding: 3,
          textColor: [0, 0, 0],
          valign: 'middle',
          halign: 'center',
          lineColor: [0, 0, 0],
          lineWidth: 0.2
        },
        headStyles: {
          fillColor: [31, 73, 125], // Exact Navy Blue (#1F497D)
          textColor: [255, 255, 255],
          fontStyle: 'bold',
          halign: 'center',
          lineColor: [0, 0, 0],
          lineWidth: 0.2
        },
        columnStyles: {
          0: { cellWidth: 35, halign: 'center' },
          1: { cellWidth: 65, halign: 'center', fontStyle: 'bold' },
          2: { cellWidth: 40, halign: 'center' }
        },
        didParseCell: (data) => {
          if (data.section === 'body' && data.column.index === 2) {
            const val = String(data.cell.raw || '').trim();
            if (val === 'Yes') {
              data.cell.styles.fillColor = [216, 228, 188]; // Light Sage Green (#D8E4BC)
              data.cell.styles.textColor = [0, 0, 0];
              data.cell.styles.fontStyle = 'bold';
            } else if (val === 'No') {
              data.cell.styles.fillColor = [242, 220, 219]; // Light Soft Pink (#F2DCDB)
              data.cell.styles.textColor = [0, 0, 0];
              data.cell.styles.fontStyle = 'bold';
            }
          }
        }
      });

      doc.save(`Supervisor_Daily_Update_Status_${dateStr}.pdf`);
    } catch (err) {
      console.error('Error generating Supervisor Status PDF:', err);
      alert('Failed to generate PDF. Please try again.');
    }
  };

  // Get status badge color
  const getStatusColor = (status) => {
    const statusStr = String(status || '').toLowerCase();
    if (statusStr.includes('pending')) return '#f59e0b';
    if (statusStr.includes('complete') || statusStr.includes('completed')) return '#10b981';
    if (statusStr.includes('working') || statusStr.includes('in progress')) return '#3b82f6';
    if (statusStr === 'not started') return '#ef4444';
    if (statusStr.includes('hold') || statusStr.includes('stopped')) return '#8b5cf6';
    if (statusStr.includes('review') || statusStr.includes('check')) return '#f97316';
    return '#6b7280';
  };

  // Get status icon
  const getStatusIcon = (status) => {
    const statusStr = String(status || '').toLowerCase();
    if (statusStr.includes('pending')) return '⏳';
    if (statusStr.includes('complete') || statusStr.includes('completed')) return '✅';
    if (statusStr.includes('working') || statusStr.includes('in progress')) return '⚡';
    if (statusStr === 'not started') return '🔴';
    if (statusStr.includes('hold') || statusStr.includes('stopped')) return '⏸️';
    if (statusStr.includes('review') || statusStr.includes('check')) return '👁️';
    return '📝';
  };

  // Render sort indicator
  const renderSortIndicator = (key) => {
    if (sortConfig.key !== key) return '↕️';
    return sortConfig.direction === 'asc' ? '⬆️' : '⬇️';
  };

  return (
    <div style={styles.container}>
      {/* Header */}
      <header style={styles.header}>
        <div style={styles.headerContent}>
          <div style={styles.headerTitle}>
            <div style={styles.logoContainer}>
              <div style={styles.logo}>🏭</div>
              <div>
                <h1 style={styles.title}>Daily Stitching Not Updation Report</h1>
                <p style={styles.subtitle}>Real-time WIP Status Monitoring & Management</p>
              </div>
            </div>
            <div style={styles.headerActions}>

            </div>
          </div>

          <div style={styles.dateContainer}>
            <span style={styles.dateIcon}>📅</span>
            <span style={styles.dateText}>{todayDate}</span>
            <div style={styles.updateInfo}>
              <span style={styles.updateDot}></span>
              Auto-refresh in 5 min
            </div>
            <button
              style={styles.refreshButton}
              onClick={fetchData}
              disabled={loading}
            >
              <span style={styles.buttonIcon}>
                {loading ? '🔄' : '🔄'}
              </span>
              {loading ? 'Refreshing...' : 'Refresh'}
            </button>
            <button
              style={styles.refreshButton}
              onClick={handleBack}
              title="Go back"
            >
              <span style={styles.backIcon}>←</span>
              Back
            </button>
          </div>
        </div>
      </header>

      {/* Statistics Cards */}
      {/* <div style={styles.statsGrid}>
        <div style={styles.statCard}>
          <div style={styles.statIconContainer} className="stat-icon-attention">
            <span style={styles.statIcon}>⚠️</span>
          </div>
          <div style={styles.statContent}>
            <h3 style={styles.statValue}>{stats.notUpdatedToday}</h3>
            <p style={styles.statLabel}>Require Attention</p>
            <div style={styles.statProgress}>
              <div style={styles.progressBar}>
                <div style={{
                  ...styles.progressFill,
                  width: `${Math.min((stats.notUpdatedToday / Math.max(stats.totalEntries, 1)) * 100, 100)}%`,
                  backgroundColor: '#ef4444'
                }}></div>
              </div>
            </div>
          </div>
        </div>
        
        <div style={styles.statCard}>
          <div style={styles.statIconContainer} className="stat-icon-pending">
            <span style={styles.statIcon}>⏳</span>
          </div>
          <div style={styles.statContent}>
            <h3 style={styles.statValue}>{stats.pendingWIP}</h3>
            <p style={styles.statLabel}>Pending WIP</p>
            <div style={styles.statProgress}>
              <div style={styles.progressBar}>
                <div style={{
                  ...styles.progressFill,
                  width: `${Math.min((stats.pendingWIP / Math.max(stats.notUpdatedToday, 1)) * 100, 100)}%`,
                  backgroundColor: '#f59e0b'
                }}></div>
              </div>
            </div>
          </div>
        </div>
        
        <div style={styles.statCard}>
          <div style={styles.statIconContainer} className="stat-icon-completed">
            <span style={styles.statIcon}>✅</span>
          </div>
          <div style={styles.statContent}>
            <h3 style={styles.statValue}>{stats.completed}</h3>
            <p style={styles.statLabel}>Completed Lots</p>
            <div style={styles.statProgress}>
              <div style={styles.progressBar}>
                <div style={{
                  ...styles.progressFill,
                  width: `${Math.min((stats.completed / Math.max(data.length, 1)) * 100, 100)}%`,
                  backgroundColor: '#10b981'
                }}></div>
              </div>
            </div>
          </div>
        </div>
        
        <div style={styles.statCard}>
          <div style={styles.statIconContainer} className="stat-icon-styles">
            <span style={styles.statIcon}>👕</span>
          </div>
          <div style={styles.statContent}>
            <h3 style={styles.statValue}>{stats.totalStyles}</h3>
            <p style={styles.statLabel}>Active Styles</p>
            <div style={styles.statProgress}>
              <div style={styles.progressBar}>
                <div style={{
                  ...styles.progressFill,
                  width: `${Math.min((stats.totalStyles / Math.max(stats.notUpdatedToday, 1)) * 100, 100)}%`,
                  backgroundColor: '#3b82f6'
                }}></div>
              </div>
            </div>
          </div>
        </div>
      </div> */}

      {/* Filters Section */}
      <div style={styles.filtersSection}>
        <div style={styles.filtersHeader}>
          <div style={styles.filtersTitle}>
            <span style={styles.filterIcon}>🔍</span>
            <h2>Filters & Search</h2>
          </div>
          <div style={styles.viewToggle}>
            <button
              style={{
                ...styles.viewToggleButton,
                ...(viewMode === 'table' ? styles.viewToggleActive : {})
              }}
              onClick={() => setViewMode('table')}
            >
              <span>📊</span> Table
            </button>
            <button
              style={{
                ...styles.viewToggleButton,
                ...(viewMode === 'card' ? styles.viewToggleActive : {})
              }}
              onClick={() => setViewMode('card')}
            >
              <span>🃏</span> Cards
            </button>
            <button
              style={{
                ...styles.viewToggleButton,
                ...(viewMode === 'summary' ? styles.viewToggleActive : {})
              }}
              onClick={() => setViewMode('summary')}
            >
              <span>📋</span> Supervisor Summary
            </button>
          </div>
        </div>

        <div style={styles.searchBox}>
          <span style={styles.searchIcon}>🔍</span>
          <input
            type="text"
            placeholder="Search across all fields..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            style={styles.searchInput}
          />
          {searchTerm && (
            <button
              onClick={() => setSearchTerm('')}
              style={styles.clearSearchButton}
            >
              ✕
            </button>
          )}
        </div>

        <div style={styles.filtersGrid}>
          {[
            { label: 'Lot Number', key: 'lotNumber', type: 'text', placeholder: 'Enter lot number' },
            { label: 'Fabric', key: 'fabric', type: 'select', options: getUniqueValues('Fabric') },
            { label: 'Garment Type', key: 'garmentType', type: 'select', options: getUniqueValues('Garment Type') },
            { label: 'WIP Status', key: 'wipStatus', type: 'select', options: getUniqueValues('WIP Status') },
            { label: 'Supervisor', key: 'supervisor', type: 'select', options: getUniqueValues('Supervisor') },
            { label: 'Party Name', key: 'partyName', type: 'select', options: getUniqueValues('PARTY NAME') },
            { label: 'Brand', key: 'brand', type: 'select', options: getUniqueValues('BRAND') },
            { label: 'Season', key: 'season', type: 'select', options: getUniqueValues('SEASON') }
          ].map((filter) => (
            <div key={filter.key} style={styles.filterGroup}>
              <label style={styles.filterLabel}>{filter.label}</label>
              {filter.type === 'select' ? (
                <select
                  name={filter.key}
                  value={filters[filter.key]}
                  onChange={handleFilterChange}
                  style={styles.filterSelect}
                >
                  {filter.options.map((option, i) => (
                    <option key={i} value={option}>{option}</option>
                  ))}
                </select>
              ) : (
                <input
                  type="text"
                  name={filter.key}
                  value={filters[filter.key]}
                  onChange={handleFilterChange}
                  placeholder={filter.placeholder}
                  style={styles.filterInput}
                />
              )}
            </div>
          ))}
        </div>

        <div style={styles.filtersActions}>
          <div style={styles.selectedInfo}>
            {selectedRows.length > 0 && (
              <>
                <span style={styles.selectedCount}>{selectedRows.length}</span>
                <span style={styles.selectedText}>rows selected</span>
              </>
            )}
          </div>
          <div style={styles.actionButtons}>
            <button
              style={styles.exportButton}
              onClick={exportToExcel}
              disabled={exporting || sortedData.length === 0}
            >
              <span style={styles.buttonIcon}>📊</span>
              Excel
            </button>
            <button
              style={styles.exportButton}
              onClick={exportToPDF}
              disabled={exporting || sortedData.length === 0}
            >
              <span style={styles.buttonIcon}>📄</span>
              PDF
            </button>
            <button
              style={styles.clearButton}
              onClick={clearFilters}
            >
              <span style={styles.buttonIcon}>🗑️</span>
              Clear All
            </button>
          </div>
        </div>
      </div>

      {/* Data Display */}
      <div style={styles.dataSection} ref={tableRef}>
        <div style={styles.dataHeader}>
          <div style={styles.dataInfo}>
            <input
              type="checkbox"
              checked={selectedRows.length === sortedData.length && sortedData.length > 0}
              onChange={handleSelectAll}
              style={styles.checkbox}
            />
            <span style={styles.dataCount}>
              Showing <strong>{sortedData.length}</strong> of{' '}
              <strong>{stats.notUpdatedToday}</strong> lots requiring attention
            </span>
          </div>
          <div style={styles.lastUpdated}>
            <span style={styles.lastUpdatedText}>
              Updated: {new Date().toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' })}
            </span>
          </div>
        </div>

        {loading && (
          <div style={styles.loadingContainer}>
            <div style={styles.spinner}></div>
            <p style={styles.loadingText}>Loading production data...</p>
          </div>
        )}

        {error && (
          <div style={styles.errorContainer}>
            <div style={styles.errorIcon}>⚠️</div>
            <div style={styles.errorContent}>
              <h3>Error Loading Data</h3>
              <p>{error}</p>
            </div>
            <button
              onClick={fetchData}
              style={styles.retryButton}
            >
              Retry
            </button>
          </div>
        )}

        {!loading && !error && sortedData.length === 0 && (
          <div style={styles.emptyState}>
            <div style={styles.emptyIcon}>🎉</div>
            <h3 style={styles.emptyTitle}>All caught up!</h3>
            <p style={styles.emptyText}>
              All WIP statuses have been updated for today. Great work!
            </p>
            <button
              onClick={fetchData}
              style={styles.checkUpdatesButton}
            >
              Check for Updates
            </button>
          </div>
        )}

        {!loading && !error && sortedData.length > 0 && viewMode === 'table' && (
          <div style={styles.tableContainer}>
            <table style={styles.table}>
              <thead style={styles.tableHead}>
                <tr>
                  <th style={styles.tableHeader}>
                    <input
                      type="checkbox"
                      checked={selectedRows.length === sortedData.length}
                      onChange={handleSelectAll}
                      style={styles.checkbox}
                    />
                  </th>
                  {[
                    { key: 'Lot Number', label: 'Lot No.', sortable: true },
                    { key: 'Fabric', label: 'Fabric', sortable: true },
                    { key: 'Garment Type', label: 'Type', sortable: true },
                    { key: 'Style', label: 'Style', sortable: true },
                    { key: 'PARTY NAME', label: 'Party', sortable: true },
                    { key: 'latestWIPStatus.status', label: 'WIP Status', sortable: false },
                    { key: 'latestWIPStatus.timestamp', label: 'Last Updated', sortable: true },
                    { key: 'latestWIPStatus.remarks', label: 'Remarks', sortable: false },
                    { key: 'Date of Issue', label: 'Issue Date', sortable: true },
                    { key: 'NormalizedSupervisor', label: 'Supervisor', sortable: true }
                  ].map((column) => (
                    <th
                      key={column.key}
                      style={styles.tableHeader}
                      onClick={column.sortable ? () => handleSort(column.key) : undefined}
                      className={column.sortable ? 'sortable-header' : ''}
                    >
                      <div style={styles.headerContent}>
                        {column.label}
                        {column.sortable && (
                          <span style={styles.sortIndicator}>
                            {renderSortIndicator(column.key)}
                          </span>
                        )}
                      </div>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {sortedData.map((item, index) => {
                  const latestWIP = item['latestWIPStatus'] || {};
                  const lastUpdatedDate = formatDateFromTimestamp(latestWIP.timestamp);
                  const statusColor = getStatusColor(latestWIP.status);
                  const statusIcon = getStatusIcon(latestWIP.status);

                  return (
                    <tr
                      key={index}
                      style={{
                        ...styles.tableRow,
                        backgroundColor: selectedRows.includes(index) ? '#f0f9ff' : 'white'
                      }}
                    >
                      <td style={styles.tableCell}>
                        <input
                          type="checkbox"
                          checked={selectedRows.includes(index)}
                          onChange={() => handleSelectRow(index)}
                          style={styles.checkbox}
                        />
                      </td>
                      <td style={{ ...styles.tableCell, ...styles.lotNumberCell }}>
                        <span style={styles.lotNumber}>{item['Lot Number'] || 'N/A'}</span>
                      </td>
                      <td style={styles.tableCell}>{item['Fabric'] || '-'}</td>
                      <td style={styles.tableCell}>{item['Garment Type'] || '-'}</td>
                      <td style={styles.tableCell}>
                        <span style={styles.styleBadge}>{item['Style'] || '-'}</span>
                      </td>
                      <td style={styles.tableCell}>{item['PARTY NAME'] || '-'}</td>
                      <td style={styles.tableCell}>
                        <div style={styles.statusContainer}>
                          <span style={{
                            ...styles.statusBadge,
                            backgroundColor: `${statusColor}15`,
                            color: statusColor,
                            border: `1px solid ${statusColor}30`
                          }}>
                            {latestWIP.status || 'Not Started'}
                          </span>
                        </div>
                        {(item['parsedWIPStatus'] || []).length > 1 && (
                          <div style={styles.updateCount}>
                            {(item['parsedWIPStatus'] || []).length} updates
                          </div>
                        )}
                      </td>
                      <td style={styles.tableCell}>
                        <div style={styles.lastUpdatedCell}>
                          <span style={{
                            color: lastUpdatedDate === 'Never' ? '#ef4444' : '#4b5563',
                            fontWeight: lastUpdatedDate === 'Never' ? '600' : '400'
                          }}>
                            {lastUpdatedDate}
                          </span>
                          {lastUpdatedDate !== 'Never' && (
                            <span style={styles.timeAgo}>
                              {(() => {
                                const now = new Date();
                                const updated = new Date(latestWIP.timestamp);
                                const diffHours = Math.floor((now - updated) / (1000 * 60 * 60));
                                if (diffHours < 1) return 'Just now';
                                if (diffHours < 24) return `${diffHours}h ago`;
                                const diffDays = Math.floor(diffHours / 24);
                                return `${diffDays}d ago`;
                              })()}
                            </span>
                          )}
                        </div>
                      </td>
                      <td style={styles.tableCell}>
                        <div
                          style={styles.remarksCell}
                          title={latestWIP.remarks || ''}
                        >
                          {latestWIP.remarks ? (
                            <>
                              {latestWIP.remarks.length > 30
                                ? `${latestWIP.remarks.substring(0, 30)}...`
                                : latestWIP.remarks}
                              {latestWIP.remarks.length > 30 && (
                                <span style={styles.viewMore}>View</span>
                              )}
                            </>
                          ) : '-'}
                        </div>
                      </td>
                      <td style={styles.tableCell}>
                        <span style={styles.dateCell}>{item['Date of Issue'] || '-'}</span>
                      </td>
                      <td style={styles.tableCell}>
                        <div style={styles.supervisorCell}>
                          <span style={styles.supervisorAvatar}>
                            {item['NormalizedSupervisor']?.charAt(0) || '?'}
                          </span>
                          <span style={styles.supervisorName}>
                            {item['NormalizedSupervisor'] || '-'}
                          </span>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}

        {/* Supervisor Summary Report View */}
        {!loading && !error && viewMode === 'summary' && (
          <div style={{ padding: '28px 32px', backgroundColor: '#ffffff', borderRadius: '24px', border: '1px solid #e2e8f0', boxShadow: '0 4px 20px rgba(0,0,0,0.03)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px', flexWrap: 'wrap', gap: '12px' }}>
              <div>
                <h2 style={{ margin: 0, fontSize: '1.4rem', fontWeight: '800', color: '#0f172a', letterSpacing: '-0.02em' }}>
                  📋 Supervisor Updation Summary Report
                </h2>
                <p style={{ margin: '4px 0 0', fontSize: '0.88rem', color: '#64748b' }}>
                  Performance breakdown of WIP updates across all active supervisors
                </p>
              </div>
              <div style={{ display: 'flex', gap: '10px' }}>
                <button
                  onClick={exportSupervisorStatusPDF}
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: '8px',
                    background: 'linear-gradient(135deg, #1f497d 0%, #1e3a8a 100%)',
                    color: '#ffffff',
                    border: '1px solid rgba(255, 255, 255, 0.2)',
                    padding: '10px 18px',
                    borderRadius: '10px',
                    fontWeight: '700',
                    fontSize: '0.88rem',
                    cursor: 'pointer',
                    boxShadow: '0 4px 12px rgba(31, 73, 125, 0.4)'
                  }}
                  title="Export 3-Column Status PDF (Date, Supervisor, Yes/No)"
                >
                  <span>📋</span> Status PDF (Yes/No)
                </button>
                <button
                  onClick={exportSupervisorSummaryToExcel}
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: '8px',
                    background: 'linear-gradient(135deg, #059669 0%, #10b981 100%)',
                    color: '#ffffff',
                    border: 'none',
                    padding: '10px 18px',
                    borderRadius: '10px',
                    fontWeight: '700',
                    fontSize: '0.88rem',
                    cursor: 'pointer',
                    boxShadow: '0 4px 12px rgba(5, 150, 105, 0.3)'
                  }}
                >
                  <span>📊</span> Excel
                </button>
                <button
                  onClick={exportSupervisorSummaryToPDF}
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: '8px',
                    background: 'linear-gradient(135deg, #4f46e5 0%, #3b82f6 100%)',
                    color: '#ffffff',
                    border: 'none',
                    padding: '10px 18px',
                    borderRadius: '10px',
                    fontWeight: '700',
                    fontSize: '0.88rem',
                    cursor: 'pointer',
                    boxShadow: '0 4px 12px rgba(79, 70, 229, 0.3)'
                  }}
                >
                  <span>📄</span> Full Summary PDF
                </button>
              </div>
            </div>

            {/* Supervisor Summary Table */}
            <div style={{ overflowX: 'auto', borderRadius: '14px', border: '1px solid #e2e8f0' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.9rem', textAlign: 'left' }}>
                <thead>
                  <tr style={{ background: 'linear-gradient(135deg, #1e1b4b 0%, #312e81 100%)', color: '#ffffff' }}>
                    <th style={{ padding: '14px 16px', fontWeight: '700' }}>Supervisor Name</th>
                    <th style={{ padding: '14px 16px', fontWeight: '700', textAlign: 'center' }}>Status (Yes/No)</th>
                    <th style={{ padding: '14px 16px', fontWeight: '700', textAlign: 'center' }}>Report Updated</th>
                    <th style={{ padding: '14px 16px', fontWeight: '700', textAlign: 'center' }}>Partial Updated</th>
                    <th style={{ padding: '14px 16px', fontWeight: '700', textAlign: 'center' }}>Not Updated</th>
                    <th style={{ padding: '14px 16px', fontWeight: '700', textAlign: 'center' }}>Total Lots</th>
                    <th style={{ padding: '14px 16px', fontWeight: '700', textAlign: 'center' }}>Total Not Updated Report</th>
                    <th style={{ padding: '14px 16px', fontWeight: '700', textAlign: 'center' }}>Percentage of Updated Lots (%)</th>
                  </tr>
                </thead>
                <tbody>
                  {supervisorSummaryData.map((row, idx) => {
                    const pctColor = row.updatedPercentage >= 80 ? '#059669' : row.updatedPercentage >= 50 ? '#d97706' : '#dc2626';
                    const isUpdated = row.reportUpdated > 0;
                    return (
                      <tr key={idx} style={{ borderBottom: '1px solid #f1f5f9', backgroundColor: idx % 2 === 0 ? '#ffffff' : '#f8fafc' }}>
                        <td style={{ padding: '14px 16px', fontWeight: '700', color: '#0f172a', display: 'flex', alignItems: 'center', gap: '10px' }}>
                          <span style={{ width: '32px', height: '32px', borderRadius: '50%', background: 'linear-gradient(135deg, #4f46e5, #3b82f6)', color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '0.85rem', fontWeight: '800' }}>
                            {row.supervisorName.charAt(0)}
                          </span>
                          {row.supervisorName}
                        </td>
                        <td style={{ padding: '14px 16px', textAlign: 'center' }}>
                          <span style={{ background: isUpdated ? '#dcfce7' : '#fee2e2', color: isUpdated ? '#15803d' : '#b91c1c', fontWeight: '800', padding: '4px 14px', borderRadius: '8px', border: `1px solid ${isUpdated ? '#bbf7d0' : '#fca5a5'}` }}>
                            {isUpdated ? 'Yes' : 'No'}
                          </span>
                        </td>
                        <td style={{ padding: '14px 16px', textAlign: 'center' }}>
                          <span style={{ background: '#dcfce7', color: '#15803d', fontWeight: '700', padding: '4px 12px', borderRadius: '8px', border: '1px solid #bbf7d0' }}>
                            {row.reportUpdated}
                          </span>
                        </td>
                        <td style={{ padding: '14px 16px', textAlign: 'center' }}>
                          <span style={{ background: '#fef3c7', color: '#b45309', fontWeight: '700', padding: '4px 12px', borderRadius: '8px', border: '1px solid #fde68a' }}>
                            {row.partialUpdated}
                          </span>
                        </td>
                        <td style={{ padding: '14px 16px', textAlign: 'center' }}>
                          <span style={{ background: '#fee2e2', color: '#b91c1c', fontWeight: '700', padding: '4px 12px', borderRadius: '8px', border: '1px solid #fca5a5' }}>
                            {row.notUpdated}
                          </span>
                        </td>
                        <td style={{ padding: '14px 16px', textAlign: 'center' }}>
                          <span style={{ background: '#e0e7ff', color: '#3730a3', fontWeight: '800', padding: '4px 12px', borderRadius: '8px', border: '1px solid #c7d2fe' }}>
                            {row.totalLots}
                          </span>
                        </td>
                        <td style={{ padding: '14px 16px', textAlign: 'center' }}>
                          <span style={{ background: '#f3e8ff', color: '#6b21a8', fontWeight: '800', padding: '4px 12px', borderRadius: '8px', border: '1px solid #e9d5ff' }}>
                            {row.totalNotUpdatedReport}
                          </span>
                        </td>
                        <td style={{ padding: '14px 16px', textAlign: 'center' }}>
                          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '10px' }}>
                            <div style={{ width: '70px', height: '8px', backgroundColor: '#e2e8f0', borderRadius: '4px', overflow: 'hidden' }}>
                              <div style={{ width: `${Math.min(row.updatedPercentage, 100)}%`, height: '100%', backgroundColor: pctColor }} />
                            </div>
                            <span style={{ fontWeight: '800', color: pctColor }}>
                              {row.updatedPercentage}%
                            </span>
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
                <tfoot>
                  <tr style={{ background: '#1e1b4b', color: '#ffffff', fontWeight: '800', borderTop: '2px solid #cbd5e1' }}>
                    <td style={{ padding: '16px', fontSize: '0.95rem' }}>{supervisorSummaryTotals.supervisorName}</td>
                    <td style={{ padding: '16px', textAlign: 'center', color: supervisorSummaryTotals.reportUpdated > 0 ? '#4ade80' : '#f87171' }}>{supervisorSummaryTotals.reportUpdated > 0 ? 'Yes' : 'No'}</td>
                    <td style={{ padding: '16px', textAlign: 'center', color: '#4ade80' }}>{supervisorSummaryTotals.reportUpdated}</td>
                    <td style={{ padding: '16px', textAlign: 'center', color: '#fbbf24' }}>{supervisorSummaryTotals.partialUpdated}</td>
                    <td style={{ padding: '16px', textAlign: 'center', color: '#f87171' }}>{supervisorSummaryTotals.notUpdated}</td>
                    <td style={{ padding: '16px', textAlign: 'center', color: '#818cf8' }}>{supervisorSummaryTotals.totalLots}</td>
                    <td style={{ padding: '16px', textAlign: 'center', color: '#c084fc' }}>{supervisorSummaryTotals.totalNotUpdatedReport}</td>
                    <td style={{ padding: '16px', textAlign: 'center', color: '#38bdf8' }}>{supervisorSummaryTotals.updatedPercentage}%</td>
                  </tr>
                </tfoot>
              </table>
            </div>

            {/* Garment Type Updation Summary Table */}
            <div style={{ marginTop: '32px' }}>
              <div style={{ marginBottom: '16px' }}>
                <h3 style={{ margin: 0, fontSize: '1.2rem', fontWeight: '800', color: '#0f172a', letterSpacing: '-0.02em', display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <span>👕</span> Status Updation as per Garment Type
                </h3>
                <p style={{ margin: '4px 0 0', fontSize: '0.85rem', color: '#64748b' }}>
                  Performance breakdown of WIP updates grouped by Garment Type
                </p>
              </div>

              <div style={{ overflowX: 'auto', borderRadius: '14px', border: '1px solid #e2e8f0' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.9rem', textAlign: 'left' }}>
                  <thead>
                    <tr style={{ background: 'linear-gradient(135deg, #0f172a 0%, #1e293b 100%)', color: '#ffffff' }}>
                      <th style={{ padding: '14px 16px', fontWeight: '700' }}>Garment Type</th>
                      <th style={{ padding: '14px 16px', fontWeight: '700', textAlign: 'center' }}>Status (Yes/No)</th>
                      <th style={{ padding: '14px 16px', fontWeight: '700', textAlign: 'center' }}>Report Updated</th>
                      <th style={{ padding: '14px 16px', fontWeight: '700', textAlign: 'center' }}>Partial Updated</th>
                      <th style={{ padding: '14px 16px', fontWeight: '700', textAlign: 'center' }}>Not Updated</th>
                      <th style={{ padding: '14px 16px', fontWeight: '700', textAlign: 'center' }}>Total Lots</th>
                      <th style={{ padding: '14px 16px', fontWeight: '700', textAlign: 'center' }}>Total Not Updated Report</th>
                      <th style={{ padding: '14px 16px', fontWeight: '700', textAlign: 'center' }}>Percentage of Updated Lots (%)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {garmentTypeSummaryData.map((row, idx) => {
                      const pctColor = row.updatedPercentage >= 80 ? '#059669' : row.updatedPercentage >= 50 ? '#d97706' : '#dc2626';
                      const isUpdated = row.reportUpdated > 0;
                      return (
                        <tr key={idx} style={{ borderBottom: '1px solid #f1f5f9', backgroundColor: idx % 2 === 0 ? '#ffffff' : '#f8fafc' }}>
                          <td style={{ padding: '14px 16px', fontWeight: '700', color: '#0f172a', display: 'flex', alignItems: 'center', gap: '10px' }}>
                            <span style={{ width: '32px', height: '32px', borderRadius: '8px', background: 'linear-gradient(135deg, #06b6d4, #3b82f6)', color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '0.85rem', fontWeight: '800' }}>
                              👕
                            </span>
                            {row.garmentType}
                          </td>
                          <td style={{ padding: '14px 16px', textAlign: 'center' }}>
                            <span style={{ background: isUpdated ? '#dcfce7' : '#fee2e2', color: isUpdated ? '#15803d' : '#b91c1c', fontWeight: '800', padding: '4px 14px', borderRadius: '8px', border: `1px solid ${isUpdated ? '#bbf7d0' : '#fca5a5'}` }}>
                              {isUpdated ? 'Yes' : 'No'}
                            </span>
                          </td>
                          <td style={{ padding: '14px 16px', textAlign: 'center' }}>
                            <span style={{ background: '#dcfce7', color: '#15803d', fontWeight: '700', padding: '4px 12px', borderRadius: '8px', border: '1px solid #bbf7d0' }}>
                              {row.reportUpdated}
                            </span>
                          </td>
                          <td style={{ padding: '14px 16px', textAlign: 'center' }}>
                            <span style={{ background: '#fef3c7', color: '#b45309', fontWeight: '700', padding: '4px 12px', borderRadius: '8px', border: '1px solid #fde68a' }}>
                              {row.partialUpdated}
                            </span>
                          </td>
                          <td style={{ padding: '14px 16px', textAlign: 'center' }}>
                            <span style={{ background: '#fee2e2', color: '#b91c1c', fontWeight: '700', padding: '4px 12px', borderRadius: '8px', border: '1px solid #fca5a5' }}>
                              {row.notUpdated}
                            </span>
                          </td>
                          <td style={{ padding: '14px 16px', textAlign: 'center' }}>
                            <span style={{ background: '#e0e7ff', color: '#3730a3', fontWeight: '800', padding: '4px 12px', borderRadius: '8px', border: '1px solid #c7d2fe' }}>
                              {row.totalLots}
                            </span>
                          </td>
                          <td style={{ padding: '14px 16px', textAlign: 'center' }}>
                            <span style={{ background: '#f3e8ff', color: '#6b21a8', fontWeight: '800', padding: '4px 12px', borderRadius: '8px', border: '1px solid #e9d5ff' }}>
                              {row.totalNotUpdatedReport}
                            </span>
                          </td>
                          <td style={{ padding: '14px 16px', textAlign: 'center' }}>
                            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '10px' }}>
                              <div style={{ width: '70px', height: '8px', backgroundColor: '#e2e8f0', borderRadius: '4px', overflow: 'hidden' }}>
                                <div style={{ width: `${Math.min(row.updatedPercentage, 100)}%`, height: '100%', backgroundColor: pctColor }} />
                              </div>
                              <span style={{ fontWeight: '800', color: pctColor }}>
                                {row.updatedPercentage}%
                              </span>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                  <tfoot>
                    <tr style={{ background: '#0f172a', color: '#ffffff', fontWeight: '800', borderTop: '2px solid #cbd5e1' }}>
                      <td style={{ padding: '16px', fontSize: '0.95rem' }}>{garmentTypeSummaryTotals.garmentType}</td>
                      <td style={{ padding: '16px', textAlign: 'center', color: garmentTypeSummaryTotals.reportUpdated > 0 ? '#34d399' : '#f87171' }}>{garmentTypeSummaryTotals.reportUpdated > 0 ? 'Yes' : 'No'}</td>
                      <td style={{ padding: '16px', textAlign: 'center', color: '#34d399' }}>{garmentTypeSummaryTotals.reportUpdated}</td>
                      <td style={{ padding: '16px', textAlign: 'center', color: '#fbbf24' }}>{garmentTypeSummaryTotals.partialUpdated}</td>
                      <td style={{ padding: '16px', textAlign: 'center', color: '#f87171' }}>{garmentTypeSummaryTotals.notUpdated}</td>
                      <td style={{ padding: '16px', textAlign: 'center', color: '#818cf8' }}>{garmentTypeSummaryTotals.totalLots}</td>
                      <td style={{ padding: '16px', textAlign: 'center', color: '#c084fc' }}>{garmentTypeSummaryTotals.totalNotUpdatedReport}</td>
                      <td style={{ padding: '16px', textAlign: 'center', color: '#38bdf8' }}>{garmentTypeSummaryTotals.updatedPercentage}%</td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            </div>
          </div>
        )}

        {!loading && !error && sortedData.length > 0 && viewMode === 'card' && (
          <div style={styles.cardsContainer}>
            {sortedData.map((item, index) => {
              const latestWIP = item['latestWIPStatus'] || {};
              const lastUpdatedDate = formatDateFromTimestamp(latestWIP.timestamp);
              const statusColor = getStatusColor(latestWIP.status);
              const statusIcon = getStatusIcon(latestWIP.status);

              return (
                <div
                  key={index}
                  style={{
                    ...styles.card,
                    borderLeft: `4px solid ${statusColor}`,
                    backgroundColor: selectedRows.includes(index) ? '#f0f9ff' : 'white'
                  }}
                >
                  <div style={styles.cardHeader}>
                    <input
                      type="checkbox"
                      checked={selectedRows.includes(index)}
                      onChange={() => handleSelectRow(index)}
                      style={styles.checkbox}
                    />
                    <div style={styles.cardTitle}>
                      <h3 style={styles.cardLotNumber}>{item['Lot Number'] || 'N/A'}</h3>
                      <span style={styles.cardStyle}>{item['Style'] || 'No Style'}</span>
                    </div>
                    <span style={{
                      ...styles.cardStatus,
                      backgroundColor: `${statusColor}15`,
                      color: statusColor
                    }}>
                      {statusIcon} {latestWIP.status || 'Not Started'}
                    </span>
                  </div>

                  <div style={styles.cardBody}>
                    <div style={styles.cardGrid}>
                      <div style={styles.cardItem}>
                        <label>Fabric</label>
                        <span>{item['Fabric'] || '-'}</span>
                      </div>
                      <div style={styles.cardItem}>
                        <label>Type</label>
                        <span>{item['Garment Type'] || '-'}</span>
                      </div>
                      <div style={styles.cardItem}>
                        <label>Party</label>
                        <span>{item['PARTY NAME'] || '-'}</span>
                      </div>
                      <div style={styles.cardItem}>
                        <label>Issue Date</label>
                        <span>{item['Date of Issue'] || '-'}</span>
                      </div>
                    </div>

                    <div style={styles.cardSupervisor}>
                      <div style={styles.supervisorInfo}>
                        <span style={styles.supervisorAvatarCard}>
                          {item['NormalizedSupervisor']?.charAt(0) || '?'}
                        </span>
                        <div>
                          <div style={styles.supervisorNameCard}>
                            {item['NormalizedSupervisor'] || '-'}
                          </div>
                          <div style={styles.supervisorRole}>Supervisor</div>
                        </div>
                      </div>
                      <div style={styles.lastUpdateCard}>
                        <div style={styles.lastUpdateLabel}>Last Update</div>
                        <div style={{
                          color: lastUpdatedDate === 'Never' ? '#ef4444' : '#4b5563',
                          fontWeight: '500'
                        }}>
                          {lastUpdatedDate}
                        </div>
                      </div>
                    </div>

                    {latestWIP.remarks && (
                      <div style={styles.cardRemarks}>
                        <label>Remarks</label>
                        <p>{latestWIP.remarks}</p>
                      </div>
                    )}
                  </div>

                  <div style={styles.cardFooter}>
                    <span style={styles.updateCountCard}>
                      {(item['parsedWIPStatus'] || []).length} updates
                    </span>
                    <span style={styles.cardBrand}>
                      {item['BRAND'] ? `Brand: ${item['BRAND']}` : 'No Brand'}
                    </span>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>

      {/* Add CSS animations and styles */}
      <style>
        {`
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          
          @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
          }
          
          @keyframes pulse {
            0%, 100% { opacity: 1; }
            50% { opacity: 0.7; }
          }
          
          @keyframes slideIn {
            from { transform: translateX(-10px); opacity: 0; }
            to { transform: translateX(0); opacity: 1; }
          }
          
          .sortable-header:hover {
            background-color: #f3f4f6;
            cursor: pointer;
          }
          
          .stat-icon-attention {
            background: linear-gradient(135deg, #fef2f2 0%, #fee2e2 100%);
          }
          
          .stat-icon-pending {
            background: linear-gradient(135deg, #fffbeb 0%, #fef3c7 100%);
          }
          
          .stat-icon-completed {
            background: linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%);
          }
          
          .stat-icon-styles {
            background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%);
          }
          
          .fade-in {
            animation: fadeIn 0.5s ease-out;
          }
          
          .slide-in {
            animation: slideIn 0.3s ease-out;
          }
          
          .pulse {
            animation: pulse 2s infinite;
          }
          
          ::-webkit-scrollbar {
            width: 8px;
            height: 8px;
          }
          
          ::-webkit-scrollbar-track {
            background: #f1f5f9;
            border-radius: 4px;
          }
          
          @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
          @keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
          .fade-in { animation: fadeIn 0.5s ease-out; }
          ::-webkit-scrollbar { width: 8px; height: 8px; }
          ::-webkit-scrollbar-track { background: #0f172a; }
          ::-webkit-scrollbar-thumb { background: #334155; border-radius: 4px; }
          ::-webkit-scrollbar-thumb:hover { background: #475569; }
        `}
      </style>
    </div>
  );
};

// Enhanced Dashboard Light Theme Styles
const styles = {
  container: {
    minHeight: '100vh',
    backgroundColor: '#f8fafc',
    backgroundImage: `
      radial-gradient(at 0% 0%, rgba(99, 102, 241, 0.08) 0px, transparent 50%),
      radial-gradient(at 100% 0%, rgba(236, 72, 153, 0.06) 0px, transparent 50%),
      radial-gradient(at 50% 100%, rgba(16, 185, 129, 0.06) 0px, transparent 50%)
    `,
    color: '#0f172a',
    fontFamily: "'Plus Jakarta Sans', 'Inter', -apple-system, sans-serif",
    padding: '24px 36px 80px',
    boxSizing: 'border-box',
  },

  header: {
    background: 'linear-gradient(135deg, #1e1b4b 0%, #312e81 40%, #4338ca 100%)',
    color: 'white',
    padding: '28px 36px',
    borderRadius: '24px',
    boxShadow: '0 20px 40px -10px rgba(49, 46, 129, 0.3)',
    border: '1px solid rgba(255, 255, 255, 0.15)',
    marginBottom: '28px',
  },
  headerContent: {
    maxWidth: '2000px',
    margin: '0 auto',
  },
  headerTitle: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: '16px',
  },
  logoContainer: {
    display: 'flex',
    alignItems: 'center',
    gap: '16px',
  },
  logo: {
    fontSize: '32px',
    background: 'linear-gradient(135deg, #4f46e5 0%, #3b82f6 50%, #06b6d4 100%)',
    width: '56px',
    height: '56px',
    borderRadius: '16px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    boxShadow: '0 10px 25px -5px rgba(79, 70, 229, 0.5)',
    border: '1px solid rgba(255, 255, 255, 0.2)',
  },
  title: {
    fontSize: '28px',
    fontWeight: '800',
    margin: '0',
    color: '#ffffff',
    letterSpacing: '-0.02em',
  },
  subtitle: {
    fontSize: '14px',
    color: '#c7d2fe',
    margin: '4px 0 0 0',
    fontWeight: '500',
  },
  headerActions: {
    display: 'flex',
    gap: '12px',
  },
  refreshButton: {
    padding: '10px 20px',
    backgroundColor: 'rgba(255, 255, 255, 0.15)',
    border: '1px solid rgba(255, 255, 255, 0.25)',
    borderRadius: '12px',
    color: 'white',
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    fontSize: '14px',
    fontWeight: '700',
    transition: 'all 0.2s ease',
    backdropFilter: 'blur(10px)',
  },
  buttonIcon: {
    fontSize: '16px',
  },
  dateContainer: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    padding: '10px 18px',
    backgroundColor: 'rgba(255, 255, 255, 0.12)',
    borderRadius: '12px',
    border: '1px solid rgba(255, 255, 255, 0.2)',
    width: 'fit-content',
  },
  dateIcon: {
    fontSize: '18px',
  },
  dateText: {
    fontSize: '14px',
    fontWeight: '600',
    color: '#ffffff',
  },
  updateInfo: {
    fontSize: '12px',
    color: '#c7d2fe',
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    marginLeft: '16px',
    fontWeight: '600',
  },
  updateDot: {
    width: '8px',
    height: '8px',
    backgroundColor: '#34d399',
    borderRadius: '50%',
    boxShadow: '0 0 8px #34d399',
  },

  statsGrid: {
    maxWidth: '2000px',
    margin: '0 auto 28px',
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))',
    gap: '20px',
  },
  statCard: {
    background: '#ffffff',
    borderRadius: '16px',
    padding: '24px',
    boxShadow: '0 4px 20px rgba(0, 0, 0, 0.03)',
    display: 'flex',
    alignItems: 'center',
    gap: '20px',
    transition: 'all 0.3s ease',
    border: '1px solid #e2e8f0',
  },
  statIconContainer: {
    width: '60px',
    height: '60px',
    borderRadius: '16px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    flexShrink: '0',
    background: '#f8fafc',
    border: '1px solid #e2e8f0',
  },
  statIcon: {
    fontSize: '26px',
  },
  statContent: {
    flex: '1',
  },
  statValue: {
    fontSize: '32px',
    fontWeight: '800',
    margin: '0 0 4px 0',
    color: '#0f172a',
  },
  statLabel: {
    fontSize: '13px',
    color: '#64748b',
    fontWeight: '600',
    margin: '0 0 10px 0',
    textTransform: 'uppercase',
    letterSpacing: '0.04em',
  },
  statProgress: {
    marginTop: '10px',
  },
  progressBar: {
    height: '6px',
    backgroundColor: '#e2e8f0',
    borderRadius: '3px',
    overflow: 'hidden',
  },
  progressFill: {
    height: '100%',
    borderRadius: '3px',
    transition: 'width 0.5s ease',
  },

  filtersSection: {
    maxWidth: '2000px',
    margin: '0 auto 28px',
    padding: '28px 32px',
    backgroundColor: '#ffffff',
    border: '1px solid #e2e8f0',
    borderRadius: '24px',
    boxShadow: '0 4px 20px rgba(0, 0, 0, 0.03)',
  },
  filtersHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: '20px',
  },
  filtersTitle: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    '& h2': {
      margin: '0',
      fontSize: '20px',
      color: '#0f172a',
      fontWeight: '800',
    },
  },
  filterIcon: {
    fontSize: '20px',
    background: 'linear-gradient(135deg, #4f46e5 0%, #3b82f6 100%)',
    width: '40px',
    height: '40px',
    borderRadius: '12px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    color: '#ffffff',
  },
  viewToggle: {
    display: 'flex',
    gap: '6px',
    backgroundColor: '#f1f5f9',
    padding: '6px',
    borderRadius: '12px',
    border: '1px solid #e2e8f0',
  },
  viewToggleButton: {
    padding: '8px 16px',
    border: 'none',
    background: 'transparent',
    borderRadius: '8px',
    cursor: 'pointer',
    fontSize: '13px',
    fontWeight: '700',
    color: '#64748b',
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    transition: 'all 0.2s ease',
  },
  viewToggleActive: {
    backgroundColor: '#4338ca',
    color: '#ffffff',
    boxShadow: '0 4px 12px rgba(67, 56, 202, 0.3)',
  },
  searchBox: {
    position: 'relative',
    marginBottom: '20px',
  },
  searchIcon: {
    position: 'absolute',
    left: '16px',
    top: '50%',
    transform: 'translateY(-50%)',
    color: '#6366f1',
    fontSize: '18px',
  },
  searchInput: {
    width: '100%',
    padding: '14px 14px 14px 48px',
    border: '2px solid #e2e8f0',
    borderRadius: '14px',
    fontSize: '15px',
    backgroundColor: '#ffffff',
    color: '#0f172a',
    transition: 'all 0.2s ease',
    boxSizing: 'border-box',
    '&:focus': {
      outline: 'none',
      borderColor: '#4338ca',
      boxShadow: '0 0 0 4px rgba(67, 56, 202, 0.15)',
    },
    '&::placeholder': {
      color: '#94a3b8',
    },
  },
  clearSearchButton: {
    position: 'absolute',
    right: '16px',
    top: '50%',
    transform: 'translateY(-50%)',
    background: 'none',
    border: 'none',
    color: '#94a3b8',
    cursor: 'pointer',
    fontSize: '18px',
    padding: '4px',
    borderRadius: '4px',
  },
  filtersGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill, minmax(220px, 1fr))',
    gap: '16px',
    marginBottom: '20px',
  },
  filterGroup: {
    display: 'flex',
    flexDirection: 'column',
    gap: '6px',
  },
  filterLabel: {
    fontSize: '11px',
    fontWeight: '700',
    color: '#475569',
    textTransform: 'uppercase',
    letterSpacing: '0.05em',
  },
  filterInput: {
    padding: '11px 14px',
    border: '2px solid #e2e8f0',
    borderRadius: '10px',
    fontSize: '13.5px',
    backgroundColor: '#ffffff',
    color: '#0f172a',
    transition: 'all 0.2s ease',
  },
  filterSelect: {
    padding: '11px 14px',
    border: '2px solid #e2e8f0',
    borderRadius: '10px',
    fontSize: '13.5px',
    backgroundColor: '#ffffff',
    color: '#0f172a',
    cursor: 'pointer',
    transition: 'all 0.2s ease',
  },
  filtersActions: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    paddingTop: '16px',
    borderTop: '1px solid #f1f5f9',
  },
  selectedInfo: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  selectedCount: {
    backgroundColor: '#4338ca',
    color: 'white',
    width: '28px',
    height: '28px',
    borderRadius: '50%',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: '12px',
    fontWeight: '700',
  },
  selectedText: {
    fontSize: '13px',
    color: '#64748b',
    fontWeight: '500',
  },
  actionButtons: {
    display: 'flex',
    gap: '10px',
  },
  exportButton: {
    padding: '10px 18px',
    backgroundColor: '#ffffff',
    border: '2px solid #e2e8f0',
    borderRadius: '10px',
    color: '#0f172a',
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    fontSize: '13px',
    fontWeight: '700',
    transition: 'all 0.2s ease',
  },
  clearButton: {
    padding: '10px 18px',
    backgroundColor: '#fef2f2',
    border: '2px solid #fee2e2',
    borderRadius: '10px',
    color: '#dc2626',
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    fontSize: '13px',
    fontWeight: '700',
    transition: 'all 0.2s ease',
  },

  dataSection: {
    maxWidth: '2000px',
    margin: '0 auto',
    padding: '28px 32px',
    backgroundColor: '#ffffff',
    border: '1px solid #e2e8f0',
    borderRadius: '24px',
    boxShadow: '0 4px 20px rgba(0, 0, 0, 0.03)',
  },
  dataHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: '20px',
  },
  dataInfo: {
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
  },
  dataCount: {
    fontSize: '14px',
    color: '#64748b',
  },
  lastUpdated: {
    fontSize: '13px',
    color: '#94a3b8',
  },
  tableContainer: {
    overflowX: 'auto',
    borderRadius: '16px',
    border: '1px solid #e2e8f0',
  },
  table: {
    width: '100%',
    borderCollapse: 'collapse',
    fontSize: '13.5px',
    textAlign: 'left',
  },
  tableHeader: {
    background: 'linear-gradient(135deg, #1e1b4b 0%, #312e81 100%)',
    color: '#ffffff',
    padding: '14px 16px',
    fontWeight: '700',
    fontSize: '13px',
  },
  tableRow: {
    borderBottom: '1px solid #f1f5f9',
    transition: 'background-color 0.2s ease',
  },
  tableCell: {
    padding: '14px 16px',
    color: '#334155',
  },
  lotNumberCell: {
    fontWeight: '700',
    color: '#312e81',
  },
  lotNumber: {
    fontFamily: "'Outfit', sans-serif",
    fontSize: '14px',
  },
  styleBadge: {
    background: '#e0e7ff',
    color: '#3730a3',
    padding: '3px 8px',
    borderRadius: '6px',
    fontWeight: '600',
    fontSize: '12px',
    border: '1px solid #c7d2fe',
  },
  statusContainer: {
    display: 'flex',
    alignItems: 'center',
  },
  statusBadge: {
    padding: '4px 10px',
    borderRadius: '8px',
    fontSize: '12px',
    fontWeight: '700',
    display: 'inline-flex',
    alignItems: 'center',
  },
  updateCount: {
    fontSize: '11px',
    color: '#64748b',
    marginTop: '2px',
  },
  lastUpdatedCell: {
    display: 'flex',
    flexDirection: 'column',
  },
  timeAgo: {
    fontSize: '11px',
    color: '#64748b',
  },
  remarksCell: {
    maxWidth: '220px',
    fontSize: '13px',
    color: '#64748b',
  },
  viewMore: {
    color: '#4338ca',
    marginLeft: '4px',
    cursor: 'pointer',
    fontWeight: '600',
  },
  dateCell: {
    color: '#475569',
    fontWeight: '500',
  },
  supervisorCell: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  supervisorAvatar: {
    width: '28px',
    height: '28px',
    borderRadius: '50%',
    background: 'linear-gradient(135deg, #4f46e5, #3b82f6)',
    color: 'white',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: '12px',
    fontWeight: '800',
  },
  supervisorName: {
    fontWeight: '600',
    color: '#0f172a',
  },
  checkbox: {
    width: '16px',
    height: '16px',
    cursor: 'pointer',
    accentColor: '#4338ca',
  },

  cardsContainer: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill, minmax(320px, 1fr))',
    gap: '20px',
  },
  card: {
    background: '#ffffff',
    borderRadius: '16px',
    padding: '20px',
    border: '1px solid #e2e8f0',
    boxShadow: '0 4px 20px rgba(0, 0, 0, 0.03)',
  },
  cardHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: '16px',
  },
  cardTitle: {
    flex: '1',
    marginLeft: '12px',
  },
  cardLotNumber: {
    fontSize: '18px',
    fontWeight: '800',
    color: '#312e81',
    margin: '0',
  },
  cardStyle: {
    fontSize: '12px',
    color: '#64748b',
  },
  cardStatus: {
    padding: '4px 10px',
    borderRadius: '8px',
    fontSize: '12px',
    fontWeight: '700',
  },
  cardBody: {
    marginBottom: '16px',
  },
  cardGrid: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    gap: '12px',
    marginBottom: '16px',
  },
  cardItem: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
    '& label': {
      fontSize: '11px',
      color: '#94a3b8',
      textTransform: 'uppercase',
    },
    '& span': {
      fontSize: '13px',
      color: '#0f172a',
      fontWeight: '500',
    },
  },
  cardSupervisor: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    padding: '12px',
    backgroundColor: '#f8fafc',
    borderRadius: '12px',
    border: '1px solid #f1f5f9',
  },
  supervisorInfo: {
    display: 'flex',
    alignItems: 'center',
    gap: '10px',
  },
  supervisorAvatarCard: {
    width: '36px',
    height: '36px',
    background: 'linear-gradient(135deg, #4f46e5, #3b82f6)',
    color: 'white',
    borderRadius: '50%',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: '14px',
    fontWeight: '800',
  },
  supervisorNameCard: {
    fontSize: '14px',
    fontWeight: '700',
    color: '#0f172a',
  },
  supervisorRole: {
    fontSize: '11px',
    color: '#64748b',
  },
  lastUpdateCard: {
    textAlign: 'right',
  },
  lastUpdateLabel: {
    fontSize: '11px',
    color: '#64748b',
  },
  cardRemarks: {
    marginTop: '12px',
    padding: '12px',
    backgroundColor: '#fefce8',
    border: '1px solid #fef08a',
    borderRadius: '10px',
    '& label': {
      fontSize: '11px',
      color: '#ca8a04',
      textTransform: 'uppercase',
      letterSpacing: '0.5px',
      marginBottom: '4px',
    },
    '& p': {
      margin: '0',
      fontSize: '13px',
      color: '#854d0e',
    },
  },
  cardFooter: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    paddingTop: '14px',
    borderTop: '1px solid #f1f5f9',
  },
  updateCountCard: {
    fontSize: '11px',
    color: '#64748b',
  },
  cardBrand: {
    fontSize: '12px',
    color: '#475569',
    fontWeight: '500',
  },
};

export default DailyNotUpdation;