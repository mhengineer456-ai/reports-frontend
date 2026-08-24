import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx'; 
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
const DailyKajButtonReport = () => {
  const [data, setData] = useState([]);
  const [filteredData, setFilteredData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [filters, setFilters] = useState({
    lotNumber: '',
    supervisor: '',
    garmentType: '',
    fabric: '',
    brand: '',
    stitchingSupervisor: '',
    minAging: '',
    maxAging: '',
    status: 'active',
    holdLots: false // NEW: Filter for Hold Lots
  });
  const [selectedRow, setSelectedRow] = useState(null);
  const [modalOpen, setModalOpen] = useState(false);

  // Google Sheets API configuration
  const SPREADSHEET_ID = '1IMhmYlJ3s2PPRgEQs1Ikd4O1OBXK4EYL1oV_-kWAkyg';
  const API_KEY = 'AIzaSyAomDFBkOySlIxKWSKGHe6ATv9gvaBr7uk';
  const RANGE = 'KajButton!B:O';

  useEffect(() => {
    fetchData();
  }, []);

  useEffect(() => {
    filterData();
  }, [data, filters]);

 const fetchData = async () => {
  setLoading(true);
  setError(null);
  
  try {
    const response = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${RANGE}?key=${API_KEY}`
    );
    
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    
    const result = await response.json();
    
    if (result.values && result.values.length > 0) {
      const headers = result.values[0];
      const rows = result.values.slice(1);
      
      const formattedData = rows
        .filter(row => {
          // Skip completely empty rows
          if (!row || row.length === 0) return false;
          
          // Skip rows where the first column (Lot Number) is empty
          if (!row[0] || row[0].trim() === '') return false;
          
          // You can add more conditions here if needed
          // For example, skip rows where all cells are empty
          const hasAnyData = row.some(cell => cell && cell.toString().trim() !== '');
          return hasAnyData;
        })
        .map(row => {
          const record = {};
          headers.forEach((header, index) => {
            record[header] = row[index] || '';
          });
          
          // Calculate aging
          record['Aging'] = calculateAging(record);
          record['Status'] = getLotStatus(record);
          
          return record;
        });
      
      setData(formattedData);
    }
  } catch (err) {
    setError(err.message || 'Failed to fetch data from Google Sheets');
    console.error('Error fetching data:', err);
  } finally {
    setLoading(false);
  }
};

  const getLotStatus = (record) => {
    const packingComplete = record['KajButton Complete'] || '';
    const wipPacking = record['WIP KajButton'] || '';
    
    if (packingComplete.trim() && packingComplete !== '[]') {
      return 'Completed';
    } else if (wipPacking.trim() && wipPacking !== '[]') {
      return 'WIP';
    } else {
      return 'Not Started';
    }
  };

  // NEW: Check if lot is on hold
  const isLotOnHold = (record) => {
    const remarks = getRecentRemarks(record).toLowerCase();
    return remarks.includes('hold');
  };

  const calculateAging = (record) => {
    try {
      const packingDate = record['KajButton Date'];
      if (!packingDate) return 0;

      const today = new Date();
      today.setHours(0, 0, 0, 0);

      const packingDateObj = new Date(packingDate);
      packingDateObj.setHours(0, 0, 0, 0);

      let endDate = today;

      const packingComplete = record['KajButton Complete'] || '';
      if (packingComplete.trim() && packingComplete !== '[]') {
        try {
          const statusData = JSON.parse(packingComplete);
          if (statusData.length > 0) {
            const latestStatus = statusData[statusData.length - 1];
            endDate = new Date(latestStatus.timestamp);
            endDate.setHours(0, 0, 0, 0);
          }
        } catch (e) {
          console.error('Error parsing KajButton Complete date:', e);
          endDate = today;
        }
      }
      
      const diffTime = endDate.getTime() - packingDateObj.getTime();
      const diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24));
      
      if (diffDays < 0) return 0;
      
      return diffDays;

    } catch (error) {
      console.error('Error calculating aging:', error);
      return 0;
    }
  };

  // Get recent remarks from WIP KajButton data
  // Get recent remarks from WIP KajButton data
const getRecentRemarks = (record) => {
  try {
    const wipPacking = record['WIP KajButton'] || '';
    if (wipPacking.trim() && wipPacking !== '[]') {
      const statusData = JSON.parse(wipPacking);
      if (statusData.length > 0) {
        // Sort by timestamp in descending order (newest first)
        const sortedData = statusData.sort((a, b) => 
          new Date(b.timestamp) - new Date(a.timestamp)
        );
        
        // Get the most recent entry (first after sorting)
        const latest = sortedData[0];
        // Return remarks or "No remarks" if empty
        return latest.remarks || (latest.status ? latest.status : 'No remarks');
      }
    }
    
    // For completed lots, get remarks from KajButton Complete
    const packingComplete = record['KajButton Complete'] || '';
    if (packingComplete.trim() && packingComplete !== '[]') {
      try {
        const statusData = JSON.parse(packingComplete);
        if (statusData.length > 0) {
          // Sort by timestamp in descending order (newest first)
          const sortedData = statusData.sort((a, b) => 
            new Date(b.timestamp) - new Date(a.timestamp)
          );
          
          const latest = sortedData[0];
          return latest.remarks || 'Completed';
        }
      } catch (e) {
        console.error('Error parsing KajButton Complete remarks:', e);
      }
    }
    
    return 'No remarks';
  } catch (error) {
    console.error('Error getting recent remarks:', error);
    return 'Error loading remarks';
  }
};

  // Get latest status details (for modal)
  const getLatestStatusDetails = (record) => {
    try {
      let statusData = [];
      
      // Check WIP KajButton first
      const wipPacking = record['WIP KajButton'] || '';
      if (wipPacking.trim() && wipPacking !== '[]') {
        statusData = JSON.parse(wipPacking);
      }
      
      // If no WIP data, check KajButton Complete
      if (statusData.length === 0) {
        const packingComplete = record['KajButton Complete'] || '';
        if (packingComplete.trim() && packingComplete !== '[]') {
          statusData = JSON.parse(packingComplete);
        }
      }
      
      if (statusData.length > 0) {
        return statusData[statusData.length - 1];
      }
      return null;
    } catch (error) {
      console.error('Error getting latest status:', error);
      return null;
    }
  };

  // Get all status history (for modal)
  const getAllStatusHistory = (record) => {
    try {
      let allStatuses = [];
      
      // Parse WIP KajButton data
      const wipPacking = record['WIP KajButton'] || '';
      if (wipPacking.trim() && wipPacking !== '[]') {
        const wipData = JSON.parse(wipPacking);
        allStatuses = [...allStatuses, ...wipData];
      }
      
      // Parse KajButton Complete data
      const packingComplete = record['KajButton Complete'] || '';
      if (packingComplete.trim() && packingComplete !== '[]') {
        const completeData = JSON.parse(packingComplete);
        allStatuses = [...allStatuses, ...completeData];
      }
      
      // Sort by timestamp (most recent first)
      return allStatuses.sort((a, b) => 
        new Date(b.timestamp) - new Date(a.timestamp)
      );
    } catch (error) {
      console.error('Error getting status history:', error);
      return [];
    }
  };

  const filterData = () => {
    let filtered = [...data];
    
    // Apply status filter - show only active (WIP + Not Started) by default
    if (filters.status === 'active') {
      filtered = filtered.filter(item => getLotStatus(item) !== 'Completed');
    } else if (filters.status !== 'all') {
      filtered = filtered.filter(item => getLotStatus(item) === filters.status);
    }
    
    // NEW: Apply Hold Lots filter
    if (filters.holdLots) {
      filtered = filtered.filter(item => isLotOnHold(item));
    }
    
    // Apply other filters
    if (filters.lotNumber) {
      filtered = filtered.filter(item => 
        item['Lot Number'].toLowerCase().includes(filters.lotNumber.toLowerCase())
      );
    }
    
    if (filters.supervisor) {
      filtered = filtered.filter(item => 
        item['kajButton Supervisor'].toLowerCase().includes(filters.supervisor.toLowerCase())
      );
    }
    
    if (filters.garmentType) {
      filtered = filtered.filter(item => 
        item['Garment Type'].toLowerCase().includes(filters.garmentType.toLowerCase())
      );
    }
    
    if (filters.fabric) {
      filtered = filtered.filter(item => 
        item['Fabric'].toLowerCase().includes(filters.fabric.toLowerCase())
      );
    }
    
    if (filters.brand) {
      filtered = filtered.filter(item => 
        item['BRAND']?.toLowerCase().includes(filters.brand.toLowerCase())
      );
    }
    
    if (filters.stitchingSupervisor) {
      filtered = filtered.filter(item => 
        item['Stiching Supervisor']?.toLowerCase().includes(filters.stitchingSupervisor.toLowerCase())
      );
    }
    
    if (filters.minAging) {
      filtered = filtered.filter(item => 
        parseInt(item['Aging']) >= parseInt(filters.minAging)
      );
    }
    
    if (filters.maxAging) {
      filtered = filtered.filter(item => 
        parseInt(item['Aging']) <= parseInt(filters.maxAging)
      );
    }
    
    setFilteredData(filtered);
  };

  // Export to Excel Function (with status)
  const exportToExcel = () => {
    const exportColumns = [
      'Lot Number', 'Fabric', 'Garment Type', 'Style', 'BRAND', 'Total Pcs', 
      'KajButton Date', 'kajButton Supervisor', 'Aging', 'Status', 'Recent Remarks', 
      'Stiching Supervisor'
    ];
    
    const dataToExport = filteredData.map(item => {
      const row = {};
      exportColumns.forEach(col => {
        if (col === 'Recent Remarks') {
          row[col] = getRecentRemarks(item);
        } else if (col === 'Status') {
          row[col] = getLotStatus(item);
        } else if (col === 'Aging') {
          row[col] = calculateAging(item);
        } else {
          row[col] = item[col] || '';
        }
      });
      return row;
    });

    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "KajButtonReport");
    XLSX.writeFile(workbook, "KajButton.xlsx");
  };

  // Export to PDF Function
// Export to PDF Function - FIXED VERSION
// Export to PDF Function - COMPLETELY REWRITTEN
// Export to PDF Function - PRINT VERSION (Opens Print Dialog)
const exportToPDF = () => {
  if (filteredData.length === 0) {
    alert('No data available to export.');
    return;
  }

  try {
    // Create PDF in landscape mode
    const doc = new jsPDF({
      orientation: 'landscape',
      unit: 'mm',
      format: 'a4'
    });

    // Color Palette matching Stitching Production Report
    const headerColor = [15, 76, 129];      // Deep Navy Blue
    const daysGood = [220, 252, 231];       // Green badge background
    const daysGoodText = [21, 128, 61];     // Green text
    const daysWarning = [254, 243, 199];    // Amber badge background
    const daysWarningText = [180, 83, 9];   // Amber text
    const daysBad = [239, 68, 68];          // Solid Vibrant Red background (#ef4444)
    const daysBadText = [255, 255, 255];    // Bold White text (#ffffff)

    // Page dimensions
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const margin = 10;
    const contentWidth = pageWidth - (margin * 2);

    let currentY = 25;

    // ---- 1. HEADER (White background with Navy Blue Title & Metrics) ----
    const drawHeader = () => {
      doc.setFillColor(255, 255, 255);
      doc.rect(0, 0, pageWidth, 28, 'F');

      const totalPCS = filteredData.reduce((sum, item) => sum + (parseInt(item['Total Pcs']) || 0), 0);
      const totalLots = filteredData.length;
      const holdLotsCount = filteredData.filter(item => isLotOnHold(item)).length;

      let title = filters.holdLots ? 'KAJBUTTON HOLD LOTS REPORT' : 'KAJBUTTON PRODUCTION REPORT';

      // Main Title - Deep Navy Blue on White background
      doc.setFontSize(18);
      doc.setTextColor(...headerColor);
      doc.setFont('Times New Roman', 'bold');
      doc.text(title, pageWidth / 2, 12, { align: 'center' });

      // Key Metrics Row - Styled in Navy Blue
      doc.setFontSize(10);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(...headerColor);

      // Left side: Date
      const today = new Date();
      const reportDate = `${String(today.getDate()).padStart(2, '0')}/${String(today.getMonth() + 1).padStart(2, '0')}/${today.getFullYear()}`;
      doc.text(`Report Date: ${reportDate}`, margin, 24);

      // Center: Summary Metrics
      const centerSummary = `Total Lots: ${totalLots}  |  Total PCS: ${totalPCS.toLocaleString()}  |  Hold Lots: ${holdLotsCount}`;
      doc.text(centerSummary, pageWidth / 2, 24, { align: 'center' });

      // Right side: Record count
      doc.text(`Showing: ${filteredData.length} records`, pageWidth - margin, 24, { align: 'right' });
    };

    drawHeader();

    // ---- 2. COLOR CODING LEGEND (Matching Stitching Report) ----
    const addColorLegend = (yPos) => {
      doc.setFontSize(8);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(...headerColor);
      doc.text('AGING DAYS COLOR CODING:', margin, yPos);

      const legendItems = [
        { color: [16, 185, 129], text: '0-2 Days: Good (Green)' },
        { color: [245, 158, 11], text: '3-5 Days: Warning (Yellow)' },
        { color: [239, 68, 68], text: '5+ Days: Critical (Red)' }
      ];

      let legendX = margin + 55;
      legendItems.forEach((item) => {
        doc.setFillColor(...item.color);
        doc.rect(legendX, yPos - 3, 5, 5, 'F');
        doc.setDrawColor(0, 0, 0);
        doc.rect(legendX, yPos - 3, 5, 5, 'D');

        doc.setTextColor(50, 50, 50);
        doc.setFont('helvetica', 'normal');
        doc.text(item.text, legendX + 7, yPos);

        legendX += 65;
      });

      return yPos + 6;
    };

    currentY = addColorLegend(30);

    // ---- 3. MAIN DATA TABLE ----
    const columns = [
      'Sr',
      'Lot No',
      'Fabric',
      'Garment',
      'Style',
      'Brand',
      'Total PCS',
      'KajButton Date',
      'Supervisor',
      'Aging (Days)',
      'Status',
      'Recent Remarks',
      'Stitching Sup'
    ];

    const body = filteredData.map((item, index) => [
      String(index + 1),
      item['Lot Number'] || '',
      item['Fabric'] || '',
      item['Garment Type'] || '',
      item['Style'] || '',
      item['BRAND'] || '',
      (parseInt(item['Total Pcs']) || 0).toLocaleString(),
      item['KajButton Date'] || '',
      item['kajButton Supervisor'] || '',
      String(item['Aging'] || 0),
      item['Status'] || 'Not Started',
      getRecentRemarks(item) || 'No remarks',
      item['Stiching Supervisor'] || ''
    ]);

    const agingColIdx = columns.indexOf('Aging (Days)');
    const statusColIdx = columns.indexOf('Status');
    const remarksColIdx = columns.indexOf('Recent Remarks');

    // Calculate proportional column widths so total sum EXACTLY equals contentWidth
    const baseWidths = [
      10, // 0: Sr
      18, // 1: Lot No
      24, // 2: Fabric
      22, // 3: Garment
      24, // 4: Style
      20, // 5: Brand
      18, // 6: Total PCS
      22, // 7: KajButton Date
      24, // 8: Supervisor
      16, // 9: Aging (Days)
      18, // 10: Status
      48, // 11: Recent Remarks (allocated extra width for full remarks)
      22  // 12: Stitching Sup
    ];

    const baseSum = baseWidths.reduce((a, b) => a + b, 0);
    const columnStyles = {};
    baseWidths.forEach((w, i) => {
      columnStyles[i] = {
        cellWidth: (w / baseSum) * contentWidth,
        halign: i === 11 ? 'left' : 'center',
        ...(i === 1 || i === 6 ? { fontStyle: 'bold' } : {})
      };
    });

    autoTable(doc, {
      head: [columns],
      body,
      startY: currentY,
      margin: { left: margin, right: margin },
      tableWidth: contentWidth,
      theme: 'grid',
      styles: {
        fontSize: 8,
        cellPadding: { top: 3, right: 2, bottom: 3, left: 2 },
        overflow: 'linebreak',
        valign: 'middle',
        textColor: [17, 24, 39],
        lineColor: [0, 0, 0],
        lineWidth: 0.4,
        fontStyle: 'normal',
      },
      headStyles: {
        fillColor: headerColor,
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        lineColor: [0, 0, 0],
        lineWidth: 0.5,
        halign: 'center',
        fontSize: 8.5,
        valign: 'middle',
        cellPadding: { top: 4, right: 2, bottom: 4, left: 2 },
      },
      bodyStyles: {
        halign: 'center',
        valign: 'middle',
      },
      alternateRowStyles: {
        fillColor: [248, 250, 252],
      },
      columnStyles,

      didParseCell: function (data) {
        if (data.section === 'body') {
          // Color code Aging column
          if (data.column.index === agingColIdx) {
            const days = parseInt(data.cell.raw) || 0;
            if (days <= 2) {
              data.cell.styles.fillColor = daysGood;
              data.cell.styles.textColor = daysGoodText;
              data.cell.styles.fontStyle = 'bold';
            } else if (days <= 5) {
              data.cell.styles.fillColor = daysWarning;
              data.cell.styles.textColor = daysWarningText;
              data.cell.styles.fontStyle = 'bold';
            } else {
              data.cell.styles.fillColor = daysBad;       // Solid Vibrant Red (#ef4444)
              data.cell.styles.textColor = daysBadText;   // Bold White text (#ffffff)
              data.cell.styles.fontStyle = 'bold';
            }
          }

          // Color code Status column
          if (data.column.index === statusColIdx) {
            const status = String(data.cell.raw).trim();
            if (status === 'Completed') {
              data.cell.styles.fillColor = [220, 252, 231];
              data.cell.styles.textColor = [21, 128, 61];
              data.cell.styles.fontStyle = 'bold';
            } else if (status === 'WIP') {
              data.cell.styles.fillColor = [224, 231, 255];
              data.cell.styles.textColor = [55, 48, 163];
              data.cell.styles.fontStyle = 'bold';
            }
          }

          // Highlight Hold Lots in Remarks
          if (data.column.index === remarksColIdx) {
            const text = String(data.cell.raw).toLowerCase();
            if (text.includes('hold')) {
              data.cell.styles.fillColor = [254, 226, 226];
              data.cell.styles.textColor = [185, 28, 28];
              data.cell.styles.fontStyle = 'bold';
            }
          }
        }
      },

      didDrawPage: function (data) {
        const pageCount = doc.internal.getNumberOfPages();
        doc.setFontSize(8);
        doc.setTextColor(100, 116, 139);
        doc.text(
          `Page ${data.pageNumber} of ${pageCount}`,
          pageWidth / 2,
          pageHeight - 6,
          { align: 'center' }
        );
      }
    });

    // ---- 4. SUPERVISOR WORKLOAD SUMMARY TABLE ----
    const supervisorTotals = {};
    filteredData.forEach(item => {
      const supervisor = item['kajButton Supervisor'] || 'Unassigned';
      const pieces = parseInt(item['Total Pcs']) || 0;
      
      if (!supervisorTotals[supervisor]) {
        supervisorTotals[supervisor] = { pieces: 0, lots: 0 };
      }
      supervisorTotals[supervisor].pieces += pieces;
      supervisorTotals[supervisor].lots += 1;
    });

    const supervisorArray = Object.entries(supervisorTotals).map(([name, data]) => ({
      name,
      lots: data.lots,
      pieces: data.pieces,
    })).sort((a, b) => b.pieces - a.pieces);

    const totalPieces = supervisorArray.reduce((sum, s) => sum + s.pieces, 0);
    const totalLots = supervisorArray.reduce((sum, s) => sum + s.lots, 0);

    const summaryHeaders = [['KajButton Supervisor', 'Lots Count', 'Total Pieces', '% of Total Production']];
    const summaryData = supervisorArray.map(s => [
      s.name,
      s.lots.toString(),
      s.pieces.toLocaleString(),
      totalPieces > 0 ? ((s.pieces / totalPieces) * 100).toFixed(1) + '%' : '0%'
    ]);

    summaryData.push([
      'TOTAL SUMMARY',
      totalLots.toString(),
      totalPieces.toLocaleString(),
      '100%'
    ]);

    let finalY = doc.lastAutoTable.finalY + 10;
    if (finalY > pageHeight - 45) {
      doc.addPage();
      finalY = 20;
    }

    doc.setFontSize(11);
    doc.setTextColor(...headerColor);
    doc.setFont('helvetica', 'bold');
    doc.text('SUPERVISOR WORKLOAD SUMMARY', margin, finalY);

    const summaryBaseWidths = [70, 45, 55, 55];
    const summaryBaseSum = summaryBaseWidths.reduce((a, b) => a + b, 0);
    const summaryColumnStyles = {};
    summaryBaseWidths.forEach((w, i) => {
      summaryColumnStyles[i] = {
        cellWidth: (w / summaryBaseSum) * contentWidth,
        halign: i === 0 ? 'left' : 'center',
        ...(i === 0 || i === 2 ? { fontStyle: 'bold' } : {})
      };
    });

    autoTable(doc, {
      head: summaryHeaders,
      body: summaryData,
      startY: finalY + 4,
      margin: { left: margin, right: margin },
      tableWidth: contentWidth,
      theme: 'grid',
      styles: {
        fontSize: 8.5,
        cellPadding: 3,
        lineColor: [0, 0, 0],
        lineWidth: 0.4,
      },
      headStyles: {
        fillColor: headerColor,
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        lineColor: [0, 0, 0],
        lineWidth: 0.5,
        halign: 'center'
      },
      bodyStyles: {
        textColor: [17, 24, 39],
        halign: 'center'
      },
      columnStyles: summaryColumnStyles,
      margin: { left: margin, right: margin }
    });

    // Print & Open Dialog
    doc.autoPrint();
    window.open(doc.output('bloburl'), '_blank');

  } catch (error) {
    console.error('PDF Generation Error Details:', error);
    alert(`Error generating PDF: ${error.message || 'Unknown error'}`);
  }
};
  // Helper function for aging text color in PDF
  const getAgingPDFColor = (days) => {
    const aging = parseInt(days) || 0;
    if (aging <= 2) return [16, 185, 129]; // Green
    if (aging <= 5) return [245, 158, 11]; // Yellow/Orange
    return [239, 68, 68]; // Red
  };

  const calculateTotals = () => {
    // Calculate unique supervisors
    const uniqueSupervisors = new Set(
      filteredData.map(item => item['kajButton Supervisor']).filter(Boolean)
    );
    
    return {
      totalLots: filteredData.length,
      totalPieces: filteredData.reduce((sum, item) => sum + (parseInt(item['Total Pcs']) || 0), 0),
      totalSupervisors: uniqueSupervisors.size, // NEW: Count of unique supervisors
      avgAging: filteredData.length > 0 ? 
        (filteredData.reduce((sum, item) => sum + parseInt(item['Aging']), 0) / filteredData.length).toFixed(1) : 0
    };
  };

  const getUniqueValues = (field) => {
    const values = new Set(data.map(item => item[field]).filter(Boolean));
    return Array.from(values).sort();
  };

  const handleFilterChange = (field, value) => {
    setFilters(prev => ({
      ...prev,
      [field]: value
    }));
  };

  const clearFilters = () => {
    setFilters({
      lotNumber: '',
      supervisor: '',
      garmentType: '',
      fabric: '',
      brand: '',
      stitchingSupervisor: '',
      minAging: '',
      maxAging: '',
      status: 'active',
      holdLots: false // NEW: Reset hold filter
    });
  };

  const openRowDetails = (record) => {
    setSelectedRow(record);
    setModalOpen(true);
  };

  const closeModal = () => {
    setModalOpen(false);
    setSelectedRow(null);
  };

  const getAgingColor = (days) => {
    if (days <= 3) return '#10B981';
    if (days <= 7) return '#F59E0B';
    return '#EF4444';
  };

  const getStatusColor = (status) => {
    switch(status) {
      case 'Completed': return '#10B981';
      case 'WIP': return '#3B82F6';
      case 'Not Started': return '#6B7280';
      default: return '#6B7280';
    }
  };

  if (loading) {
    return (
      <div className="loading-container">
        <div className="spinner"></div>
        <p>Loading KajButton Report...</p>
      </div>
    );
  }

  if (error) {
    return (
      <div className="error-message">
        <h3>Error Loading Data</h3>
        <p>{error}</p>
        <button onClick={fetchData} className="retry-btn">Retry</button>
      </div>
    );
  }

  const totals = calculateTotals();

  return (
    <>
      <style>
        {`
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
          .btn-excel:hover {
            box-shadow: 0 6px 16px rgba(16, 185, 129, 0.4) !important;
          }

          .btn-pdf {
            background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%) !important;
            border: none !important;
            box-shadow: 0 4px 12px rgba(239, 68, 68, 0.3) !important;
          }
          .btn-pdf:hover {
            box-shadow: 0 6px 16px rgba(239, 68, 68, 0.4) !important;
          }

          .btn-back {
            background: rgba(255, 255, 255, 0.15) !important;
          }

          .btn-refresh {
            background: linear-gradient(135deg, #6366f1 0%, #4f46e5 100%) !important;
            border: none !important;
            box-shadow: 0 4px 12px rgba(99, 102, 241, 0.3) !important;
          }

          .btn-clear {
            background: rgba(255, 255, 255, 0.1) !important;
          }
          
          .summary-cards {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
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
          
          .summary-card.total-lots {
            border-left-color: #6366f1;
          }
          
          .summary-card.total-pieces {
            border-left-color: #10b981;
          }
          
          .summary-card.total-supervisors {
            border-left-color: #a855f7;
          }
          
          .summary-card.avg-aging {
            border-left-color: #f59e0b;
          }
          
          .card-label {
            font-size: 12px;
            font-weight: 700;
            color: #64748b;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 6px;
            display: flex;
            align-items: center;
            gap: 6px;
          }
          
          .card-value {
            font-size: 28px;
            font-weight: 800;
            color: #0f172a;
            margin: 0;
            line-height: 1.1;
          }
          
          .filters-container {
            background: #ffffff;
            border-radius: 20px;
            padding: 26px 28px;
            margin-bottom: 28px;
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.05);
            border: 1px solid #f1f5f9;
          }
          
          .filters-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            padding-bottom: 14px;
            border-bottom: 1px solid #f1f5f9;
          }
          
          .filters-title {
            font-size: 1.1rem;
            font-weight: 800;
            color: #1e1b4b;
            display: flex;
            align-items: center;
            gap: 8px;
            margin: 0;
          }
          
          .filters-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
            gap: 16px;
          }
          
          .filter-group {
            display: flex;
            flex-direction: column;
          }
          
          .filter-label {
            font-size: 12px;
            font-weight: 700;
            color: #475569;
            margin-bottom: 6px;
            text-transform: uppercase;
            letter-spacing: 0.4px;
          }
          
          .filter-input,
          .filter-select {
            padding: 10px 14px;
            border: 2px solid #e2e8f0;
            border-radius: 12px;
            font-size: 13px;
            font-weight: 500;
            transition: all 0.2s;
            background: #ffffff;
            color: #0f172a;
          }
          
          .filter-input:focus,
          .filter-select:focus {
            outline: none;
            border-color: #6366f1;
            box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.15);
          }
          
          /* Checkbox styling for Hold Lots filter */
          .checkbox-filter {
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 10px 14px;
            border: 2px solid #e2e8f0;
            border-radius: 12px;
            cursor: pointer;
            transition: all 0.2s;
            background: #ffffff;
          }
          
          .checkbox-filter:hover {
            border-color: #6366f1;
            background: #f8fafc;
          }
          
          .checkbox-filter input[type="checkbox"] {
            width: 18px;
            height: 18px;
            cursor: pointer;
            accent-color: #ef4444;
          }
          
          .checkbox-filter label {
            font-size: 13px;
            font-weight: 600;
            color: #334155;
            cursor: pointer;
            flex: 1;
          }
          
          .checkbox-filter.active {
            border-color: #ef4444;
            background: #fef2f2;
          }
          
          .hold-indicator {
            background: #ef4444;
            color: white;
            padding: 4px 10px;
            border-radius: 9999px;
            font-size: 11px;
            font-weight: 700;
            margin-left: 8px;
            text-transform: uppercase;
          }
          
          .table-container {
            background: #ffffff;
            border-radius: 20px;
            overflow: hidden;
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.05);
            border: 1px solid #f1f5f9;
            overflow-x: auto;
          }
          
          .table-header {
            padding: 20px 28px;
            border-bottom: 1px solid #f1f5f9;
            display: flex;
            justify-content: space-between;
            align-items: center;
            background: #ffffff;
          }
          
          .table-title {
            font-size: 1.15rem;
            font-weight: 800;
            color: #1e1b4b;
            margin: 0;
          }
          
          .records-count {
            font-size: 13px;
            color: #64748b;
            font-weight: 600;
          }
          
          .data-table {
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            min-width: 1500px;
          }
          
          /* Modern Indigo/Navy Table Headers matching Dashboard */
          .data-table thead {
            background: linear-gradient(135deg, #1e1b4b 0%, #312e81 100%) !important;
          }
          
          .data-table th {
            background: #1e1b4b !important;
            color: #ffffff !important;
            padding: 16px 20px;
            text-align: left;
            font-weight: 800;
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.6px;
            border-bottom: 2px solid #312e81;
            position: sticky;
            top: 0;
            z-index: 10;
          }
          
          .data-table td {
            padding: 14px 20px;
            border-bottom: 1px solid #f1f5f9;
            color: #1e293b;
            font-size: 13.5px;
            font-weight: 500;
            vertical-align: middle;
          }
          
          .data-table tbody tr {
            transition: all 0.2s ease;
            cursor: pointer;
          }

          .data-table tbody tr:nth-child(even) {
            background-color: #f8fafc;
          }
          
          .data-table tbody tr:hover {
            background-color: #eef2ff !important;
          }
          
          .status-badge {
            padding: 6px 14px;
            border-radius: 9999px;
            font-size: 12px;
            font-weight: 700;
            display: inline-flex;
            align-items: center;
            gap: 4px;
          }
          
          .aging-badge {
            padding: 6px 14px;
            border-radius: 9999px;
            font-size: 12px;
            font-weight: 800;
            color: white;
            display: inline-block;
            min-width: 50px;
            text-align: center;
            box-shadow: 0 2px 6px rgba(0, 0, 0, 0.1);
          }
          
          .supervisor-tag {
            background: #e0e7ff;
            color: #3730a3;
            padding: 4px 12px;
            border-radius: 9999px;
            font-size: 12px;
            font-weight: 700;
            display: inline-block;
          }
          
          .stitching-supervisor-tag {
            background: #f0f9ff;
            color: #0369a1;
            padding: 4px 12px;
            border-radius: 9999px;
            font-size: 12px;
            font-weight: 700;
            display: inline-block;
          }
          
          .brand-tag {
            background: #fef3c7;
            color: #92400e;
            padding: 4px 12px;
            border-radius: 9999px;
            font-size: 12px;
            font-weight: 700;
            display: inline-block;
          }
          
          .remarks-cell {
            max-width: 220px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
            font-size: 12px;
            color: #475569;
            background: #f8fafc;
            padding: 8px 12px;
            border-radius: 10px;
            border-left: 4px solid #6366f1;
            font-weight: 500;
          }
          
          .remarks-cell.hold {
            border-left-color: #ef4444;
            background: #fef2f2;
            color: #991b1b;
            font-weight: 700;
          }
          
          .loading-container {
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            min-height: 400px;
            color: #6366f1;
          }
          
          .spinner {
            width: 50px;
            height: 50px;
            border: 4px solid #e2e8f0;
            border-top: 4px solid #6366f1;
            border-radius: 50%;
            animation: spin 1s linear infinite;
            margin-bottom: 20px;
          }
          
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          
          .error-message {
            background: #fef2f2;
            border: 1px solid #fecaca;
            border-radius: 16px;
            padding: 28px;
            margin: 32px;
            text-align: center;
            box-shadow: 0 10px 25px -5px rgba(239, 68, 68, 0.1);
          }
          
          .error-message h3 {
            color: #dc2626;
            margin: 0 0 12px 0;
            font-size: 1.3rem;
            font-weight: 800;
          }
          
          .error-message p {
            color: #991b1b;
            margin: 0 0 20px 0;
          }
          
          .retry-btn {
            background: #dc2626;
            color: white;
            border: none;
            padding: 10px 26px;
            border-radius: 12px;
            cursor: pointer;
            font-weight: 700;
            transition: all 0.2s;
            box-shadow: 0 4px 12px rgba(220, 38, 38, 0.3);
          }
          
          .retry-btn:hover {
            background: #b91c1c;
            transform: translateY(-2px);
          }
          
          .modal-overlay {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(15, 23, 42, 0.75);
            display: flex;
            justify-content: center;
            align-items: center;
            z-index: 1000;
            padding: 20px;
            backdrop-filter: blur(6px);
          }
          
          .modal-content {
            background: #ffffff;
            border-radius: 24px;
            max-width: 800px;
            width: 100%;
            max-height: 90vh;
            overflow-y: auto;
            box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.35);
            animation: slideIn 0.3s ease-out;
            border: 1px solid #f1f5f9;
          }
          
          @keyframes slideIn {
            from {
              opacity: 0;
              transform: translateY(20px);
            }
            to {
              opacity: 1;
              transform: translateY(0);
            }
          }
          
          .modal-header {
            background: linear-gradient(135deg, #1e1b4b 0%, #312e81 100%);
            color: white;
            padding: 24px 32px;
            border-radius: 24px 24px 0 0;
            display: flex;
            justify-content: space-between;
            align-items: center;
          }
          
          .modal-title {
            font-size: 1.4rem;
            font-weight: 800;
            margin: 0;
            color: #ffffff;
          }
          
          .close-button {
            background: rgba(255, 255, 255, 0.15);
            border: none;
            color: white;
            font-size: 22px;
            cursor: pointer;
            width: 36px;
            height: 36px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: all 0.2s;
          }
          
          .close-button:hover {
            background: #ef4444;
            transform: rotate(90deg);
          }
          
          .modal-body {
            padding: 28px;
          }
          
          .row-details-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 16px;
            margin-bottom: 32px;
          }
          
          .detail-card {
            background: #f8fafc;
            border-radius: 14px;
            padding: 18px;
            border: 1px solid #f1f5f9;
          }
          
          .detail-label {
            font-size: 11px;
            font-weight: 700;
            color: #64748b;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 6px;
          }
          
          .detail-value {
            font-size: 15px;
            font-weight: 700;
            color: #0f172a;
          }
          
          .status-section {
            margin-bottom: 32px;
          }
          
          .section-title {
            font-size: 1.1rem;
            font-weight: 800;
            color: #1e1b4b;
            margin-bottom: 18px;
            padding-bottom: 10px;
            border-bottom: 2px solid #f1f5f9;
          }
          
          .status-log-item {
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 14px;
            padding: 18px;
            margin-bottom: 14px;
            box-shadow: 0 2px 6px rgba(0, 0, 0, 0.02);
          }
          
          .log-item-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 12px;
          }
          
          .log-status {
            font-weight: 800;
            font-size: 15px;
            color: #0f172a;
          }
          
          .log-timestamp {
            font-size: 12px;
            color: #64748b;
            font-weight: 600;
          }
          
          .log-details {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 16px;
          }
          
          .log-detail-item {
            display: flex;
            flex-direction: column;
          }
          
          .log-detail-label {
            font-size: 11px;
            color: #64748b;
            margin-bottom: 4px;
            font-weight: 600;
            text-transform: uppercase;
          }
          
          .log-detail-value {
            font-size: 13.5px;
            color: #0f172a;
            font-weight: 600;
          }
          
          .no-data {
            text-align: center;
            padding: 60px 20px;
            color: #64748b;
            font-weight: 600;
          }
          
          .no-data-icon {
            font-size: 48px;
            margin-bottom: 16px;
            opacity: 0.6;
          }
          
          @media (max-width: 768px) {
            .wip-report-container {
              padding: 16px;
            }
            
            .header-content {
              flex-direction: column;
              gap: 16px;
              align-items: flex-start;
            }
            
            .header-controls {
              width: 100%;
              justify-content: flex-start;
            }
            
            .summary-cards {
              grid-template-columns: 1fr;
            }
            
            .filters-grid {
              grid-template-columns: 1fr;
            }
            
            .table-header {
              flex-direction: column;
              gap: 12px;
              align-items: flex-start;
            }
            
            .modal-content {
              margin: 10px;
            }
          }
        `}
      </style>
      
      {/* Row Details Modal */}
      {modalOpen && selectedRow && (
        <div className="modal-overlay" onClick={closeModal}>
          <div className="modal-content" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <h2 className="modal-title">
                Lot #{selectedRow['Lot Number']} - {getLotStatus(selectedRow)}
              </h2>
              <button className="close-button" onClick={closeModal}>
                ×
              </button>
            </div>
            <div className="modal-body">
              <div className="row-details-grid">
                <div className="detail-card">
                  <div className="detail-label">Garment Type</div>
                  <div className="detail-value">{selectedRow['Garment Type']}</div>
                </div>
                <div className="detail-card">
                  <div className="detail-label">Fabric</div>
                  <div className="detail-value">{selectedRow['Fabric']}</div>
                </div>
                <div className="detail-card">
                  <div className="detail-label">Style</div>
                  <div className="detail-value">{selectedRow['Style']}</div>
                </div>
                <div className="detail-card">
                  <div className="detail-label">Brand</div>
                  <div className="detail-value">{selectedRow['BRAND'] || 'N/A'}</div>
                </div>
                <div className="detail-card">
                  <div className="detail-label">KajButton Date</div>
                  <div className="detail-value">{selectedRow['KajButton Date']}</div>
                </div>
                <div className="detail-card">
                  <div className="detail-label">Total Pieces</div>
                  <div className="detail-value">{selectedRow['Total Pcs']}</div>
                </div>
                <div className="detail-card">
                  <div className="detail-label">Manpower</div>
                  <div className="detail-value">{selectedRow['Total Manpower'] || '0'}</div>
                </div>
                <div className="detail-card">
                  <div className="detail-label">Aging (Days)</div>
                  <div className="detail-value">{selectedRow['Aging']}</div>
                </div>
                <div className="detail-card">
                  <div className="detail-label">kajButton Supervisor</div>
                  <div className="detail-value">{selectedRow['kajButton Supervisor']}</div>
                </div>
                <div className="detail-card">
                  <div className="detail-label">Stiching Supervisor</div>
                  <div className="detail-value">{selectedRow['Stiching Supervisor'] || 'N/A'}</div>
                </div>
                <div className="detail-card">
                  <div className="detail-label">Status</div>
                  <div className="detail-value" style={{ 
                    color: getStatusColor(selectedRow['Status']),
                    fontWeight: 'bold'
                  }}>
                    {selectedRow['Status']}
                  </div>
                </div>
                <div className="detail-card">
                  <div className="detail-label">Recent Remarks</div>
                  <div className="detail-value">{getRecentRemarks(selectedRow)}</div>
                </div>
                <div className="detail-card">
                  <div className="detail-label">Hold Status</div>
                  <div className="detail-value" style={{ 
                    color: isLotOnHold(selectedRow) ? '#EF4444' : '#10B981',
                    fontWeight: 'bold'
                  }}>
                    {isLotOnHold(selectedRow) ? 'On Hold' : 'Not on Hold'}
                  </div>
                </div>
              </div>
              
              {/* Status History Section */}
              <div className="status-section">
                <h3 className="section-title">Status History</h3>
                {getAllStatusHistory(selectedRow).length > 0 ? (
                  getAllStatusHistory(selectedRow).map((status, index) => (
                    <div className="status-log-item" key={index}>
                      <div className="log-item-header">
                        <div className="log-status">
                          {status.status}
                        </div>
                        <div className="log-timestamp">
                          {new Date(status.timestamp).toLocaleString()}
                        </div>
                      </div>
                      <div className="log-details">
                        <div className="log-detail-item">
                          <div className="log-detail-label">Supervisor</div>
                          <div className="log-detail-value">
                            {status.supervisor}
                          </div>
                        </div>
                        <div className="log-detail-item">
                          <div className="log-detail-label">Remarks</div>
                          <div className="log-detail-value">
                            {status.remarks || 'No remarks'}
                          </div>
                        </div>
                        <div className="log-detail-item">
                          <div className="log-detail-label">Date</div>
                          <div className="log-detail-value">
                            {status.date}
                          </div>
                        </div>
                        <div className="log-detail-item">
                          <div className="log-detail-label">Time</div>
                          <div className="log-detail-value">
                            {status.time}
                          </div>
                        </div>
                      </div>
                    </div>
                  ))
                ) : (
                  <p>No status history available</p>
                )}
              </div>
            </div>
          </div>
        </div>
      )}
      
      <div className="wip-report-container">
        {/* Header */}
        <div className="header">
          <div className="header-content">
            <div className="header-title-box">
              <h1>
                <span>🔘</span> KajButton Production Report
                <span className="live-status-badge">
                  <span className="live-dot"></span> LIVE
                </span>
              </h1>
              <p className="header-subtitle">Real-time Monitoring • WIP & Completed Lots Tracking</p>
            </div>
            <div className="header-controls">
              <button 
                className="icon-btn btn-back" 
                onClick={() => window.history.back()}
                title="Go Back"
              >
                <span>←</span> Back
              </button>
              <button 
                className="icon-btn btn-excel" 
                onClick={exportToExcel} 
                title="Download as Excel (with Status column)"
              >
                <span>📊</span> Export Excel
              </button>
              
              <button 
                className="icon-btn btn-pdf" 
                onClick={exportToPDF} 
                title="Download as PDF (without Status column)"
              >
                <span>📄</span> Export PDF
              </button>
              
              <button className="icon-btn btn-refresh" onClick={fetchData} title="Refresh data">
                <span>↻</span> Refresh
              </button>
              <button className="icon-btn btn-clear" onClick={clearFilters} title="Clear all filters">
                <span>✕</span> Clear
              </button>
            </div>
          </div>
          
          {/* Summary Cards - CHANGED: Total Manpower card replaced with Total kajButton Supervisor card */}
          <div className="summary-cards">
            <div className="summary-card total-lots">
              <div className="card-label">
                <span>📋</span> Total Lots
              </div>
              <p className="card-value">{totals.totalLots}</p>
            </div>
            
            <div className="summary-card total-pieces">
              <div className="card-label">
                <span>👕</span> Total Pieces
              </div>
              <p className="card-value">{totals.totalPieces.toLocaleString()}</p>
            </div>
            
            <div className="summary-card total-supervisors">
              <div className="card-label">
                <span>👤</span> kajButton Supervisors
              </div>
              <p className="card-value">{totals.totalSupervisors}</p>
            </div>
            
            <div className="summary-card avg-aging">
              <div className="card-label">
                <span>📅</span> Average Aging (Days)
              </div>
              <p className="card-value">{totals.avgAging}</p>
            </div>
          </div>
        </div>

        {/* Filters Section - ADDED: Hold Lots filter */}
        <div className="filters-container">
          <div className="filters-header">
            <h3 className="filters-title">
              <span>🔍</span> Filter Options
            </h3>
            <span className="records-count">
              {filteredData.length} lots found
              {filters.holdLots && (
                <span className="hold-indicator">Hold Lots Only</span>
              )}
            </span>
          </div>
          
          <div className="filters-grid">
            <div className="filter-group">
              <label className="filter-label">Lot Number</label>
              <input
                type="text"
                className="filter-input"
                placeholder="Search by lot..."
                value={filters.lotNumber}
                onChange={(e) => handleFilterChange('lotNumber', e.target.value)}
              />
            </div>
            
            <div className="filter-group">
              <label className="filter-label">kajButton Supervisor</label>
              <select
                className="filter-select"
                value={filters.supervisor}
                onChange={(e) => handleFilterChange('supervisor', e.target.value)}
              >
                <option value="">All Supervisors</option>
                {getUniqueValues('kajButton Supervisor').map((supervisor, index) => (
                  <option key={index} value={supervisor}>
                    {supervisor}
                  </option>
                ))}
              </select>
            </div>
            
            <div className="filter-group">
              <label className="filter-label">Garment Type</label>
              <select
                className="filter-select"
                value={filters.garmentType}
                onChange={(e) => handleFilterChange('garmentType', e.target.value)}
              >
                <option value="">All Types</option>
                {getUniqueValues('Garment Type').map((type, index) => (
                  <option key={index} value={type}>
                    {type}
                  </option>
                ))}
              </select>
            </div>
            
            <div className="filter-group">
              <label className="filter-label">Fabric</label>
              <select
                className="filter-select"
                value={filters.fabric}
                onChange={(e) => handleFilterChange('fabric', e.target.value)}
              >
                <option value="">All Fabrics</option>
                {getUniqueValues('Fabric').map((fabric, index) => (
                  <option key={index} value={fabric}>
                    {fabric}
                  </option>
                ))}
              </select>
            </div>
            
            <div className="filter-group">
              <label className="filter-label">Brand</label>
              <select
                className="filter-select"
                value={filters.brand}
                onChange={(e) => handleFilterChange('brand', e.target.value)}
              >
                <option value="">All Brands</option>
                {getUniqueValues('BRAND').map((brand, index) => (
                  <option key={index} value={brand}>
                    {brand}
                  </option>
                ))}
              </select>
            </div>
            
            <div className="filter-group">
              <label className="filter-label">Stiching Supervisor</label>
              <select
                className="filter-select"
                value={filters.stitchingSupervisor}
                onChange={(e) => handleFilterChange('stitchingSupervisor', e.target.value)}
              >
                <option value="">All Stiching Supervisors</option>
                {getUniqueValues('Stiching Supervisor').map((supervisor, index) => (
                  <option key={index} value={supervisor}>
                    {supervisor}
                  </option>
                ))}
              </select>
            </div>
            
            <div className="filter-group">
              <label className="filter-label">Status</label>
              <select
                className="filter-select"
                value={filters.status}
                onChange={(e) => handleFilterChange('status', e.target.value)}
              >
                <option value="active">Active (WIP + Not Started)</option>
                <option value="all">All Status</option>
                <option value="WIP">WIP Only</option>
                <option value="Completed">Completed Only</option>
                <option value="Not Started">Not Started Only</option>
              </select>
            </div>
            
            <div className="filter-group">
              <label className="filter-label">Min Aging (Days)</label>
              <input
                type="number"
                className="filter-input"
                placeholder="e.g. 5"
                value={filters.minAging}
                onChange={(e) => handleFilterChange('minAging', e.target.value)}
              />
            </div>
            
            <div className="filter-group">
              <label className="filter-label">Max Aging (Days)</label>
              <input
                type="number"
                className="filter-input"
                placeholder="e.g. 10"
                value={filters.maxAging}
                onChange={(e) => handleFilterChange('maxAging', e.target.value)}
              />
            </div>
            
            {/* NEW: Hold Lots filter */}
            <div className={`checkbox-filter ${filters.holdLots ? 'active' : ''}`}>
              <input
                type="checkbox"
                id="holdLots"
                checked={filters.holdLots}
                onChange={(e) => handleFilterChange('holdLots', e.target.checked)}
              />
              <label htmlFor="holdLots">
                Show Hold Lots Only
                <span style={{ fontSize: '12px', color: '#6B7280', fontWeight: 'normal', marginLeft: '8px' }}>
                  (Filter lots with "hold" in remarks)
                </span>
              </label>
            </div>
            
          </div>
        </div>
        
        {/* Main Table with Navy Blue Headers */}
        <div className="table-container">
          <div className="table-header">
            <h3 className="table-title">
              KajButton Lots Details
              {filters.holdLots && (
                <span style={{ color: '#EF4444', marginLeft: '10px' }}>
                  (Hold Lots Only)
                </span>
              )}
            </h3>
            <span className="records-count">
              Showing {filteredData.length} of {data.length} records
              <span style={{ color: '#0f4c81', fontWeight: 'bold', marginLeft: '8px' }}>
                (Showing Active Lots Only by Default)
              </span>
            </span>
          </div>
          
          <table className="data-table">
            <thead>
              <tr>
                <th>Lot #</th>
                <th>Fabric</th>
                <th>Garment Type</th>
                <th>Style</th>
                <th>Brand</th>
                <th>Total Pcs</th>
                <th>KajButton Date</th>
                <th>kajButton Supervisor</th>
                <th>Aging (Days)</th>
                <th>Status</th>
                <th>Recent Remarks</th>
                <th>Stiching Supervisor</th>
              </tr>
            </thead>
            <tbody>
              {filteredData.length > 0 ? (
                filteredData.map((item, index) => {
                  const isHold = isLotOnHold(item);
                  return (
                    <tr key={index} onClick={() => openRowDetails(item)}>
                      <td>{item['Lot Number']}</td>
                      <td>{item['Fabric']}</td>
                      <td>{item['Garment Type']}</td>
                      <td>{item['Style']}</td>
                      <td>
                        {item['BRAND'] && (
                          <span className="brand-tag">{item['BRAND']}</span>
                        )}
                      </td>
                      <td>{item['Total Pcs']}</td>
                      <td>{item['KajButton Date']}</td>
                      <td><span className="supervisor-tag">{item['kajButton Supervisor']}</span></td>
                      <td>
                        <span 
                          className="aging-badge" 
                          style={{ backgroundColor: getAgingColor(item['Aging']) }}
                        >
                          {item['Aging']}
                        </span>
                      </td>
                      <td>
                        <span 
                          className="status-badge" 
                          style={{ 
                            backgroundColor: `${getStatusColor(item['Status'])}1A`, 
                            color: getStatusColor(item['Status']) 
                          }}
                        >
                          {item['Status']}
                        </span>
                      </td>
                      <td>
                        <div 
                          className={`remarks-cell ${isHold ? 'hold' : ''}`} 
                          title={getRecentRemarks(item)}
                        >
                          {isHold && <span style={{ color: '#EF4444', fontWeight: 'bold' }}>⏸️ </span>}
                          {getRecentRemarks(item)}
                        </div>
                      </td>
                      <td>
                        {item['Stiching Supervisor'] && (
                          <span className="stitching-supervisor-tag">
                            {item['Stiching Supervisor']}
                          </span>
                        )}
                      </td>
                    </tr>
                  );
                })
              ) : (
                <tr>
                  <td colSpan="12">
                    <div className="no-data">
                      <div className="no-data-icon">📭</div>
                      No kajbutton lots match your current filters.
                    </div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </>
  );
};

export default DailyKajButtonReport;