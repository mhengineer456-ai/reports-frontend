import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx'; 
import { jsPDF } from 'jspdf';
import { autoTable } from 'jspdf-autotable';

const DailyPackingReport = () => {
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
  const SPREADSHEET_ID = '1uo14nKO_yHu4AJ2rOgaJajuprcinj6xw1AUMFJ6_zYM';
  const API_KEY = 'AIzaSyAomDFBkOySlIxKWSKGHe6ATv9gvaBr7uk';
  const RANGE = 'Issues!B:O';

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
    const packingComplete = record['Packing Complete'] || '';
    const wipPacking = record['WIP Packing'] || '';
    
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
      const packingDate = record['Packing Date'];
      if (!packingDate) return 0;

      const today = new Date();
      today.setHours(0, 0, 0, 0);

      const packingDateObj = new Date(packingDate);
      packingDateObj.setHours(0, 0, 0, 0);

      let endDate = today;

      const packingComplete = record['Packing Complete'] || '';
      if (packingComplete.trim() && packingComplete !== '[]') {
        try {
          const statusData = JSON.parse(packingComplete);
          if (statusData.length > 0) {
            const latestStatus = statusData[statusData.length - 1];
            endDate = new Date(latestStatus.timestamp);
            endDate.setHours(0, 0, 0, 0);
          }
        } catch (e) {
          console.error('Error parsing packing complete date:', e);
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

  // Get recent remarks from WIP Packing data
  // Get recent remarks from WIP Packing data
const getRecentRemarks = (record) => {
  try {
    const wipPacking = record['WIP Packing'] || '';
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
    
    // For completed lots, get remarks from Packing Complete
    const packingComplete = record['Packing Complete'] || '';
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
        console.error('Error parsing packing complete remarks:', e);
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
      
      // Check WIP Packing first
      const wipPacking = record['WIP Packing'] || '';
      if (wipPacking.trim() && wipPacking !== '[]') {
        statusData = JSON.parse(wipPacking);
      }
      
      // If no WIP data, check Packing Complete
      if (statusData.length === 0) {
        const packingComplete = record['Packing Complete'] || '';
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
      
      // Parse WIP Packing data
      const wipPacking = record['WIP Packing'] || '';
      if (wipPacking.trim() && wipPacking !== '[]') {
        const wipData = JSON.parse(wipPacking);
        allStatuses = [...allStatuses, ...wipData];
      }
      
      // Parse Packing Complete data
      const packingComplete = record['Packing Complete'] || '';
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
        item['Packing Supervisor'].toLowerCase().includes(filters.supervisor.toLowerCase())
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
        item['STITCHING SUPERVISOR']?.toLowerCase().includes(filters.stitchingSupervisor.toLowerCase())
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
// Export to Excel Function (with status) - UPDATED
const exportToExcel = () => {
  const exportColumns = [
    'Lot Number', 'Fabric', 'Garment Type', 'Style', 'BRAND', 'Total Pcs', 
    'Packing Date', 'Packing Supervisor', 'Aging', 'Status', 'Recent Remarks', 
    'STITCHING SUPERVISOR'
  ];
  
  // Sort data in ascending order by Packing Date (or another field if you prefer)
  const sortedData = [...filteredData].sort((a, b) => {
    // Sort by Packing Date in ascending order
    const dateA = a['Packing Date'] ? new Date(a['Packing Date']) : new Date(0);
    const dateB = b['Packing Date'] ? new Date(b['Packing Date']) : new Date(0);
    return dateA - dateB;
  });
  
  const dataToExport = sortedData.map(item => {
    const row = {};
    exportColumns.forEach(col => {
      if (col === 'Recent Remarks') {
        row[col] = getRecentRemarks(item);
      } else if (col === 'Status') {
        row[col] = getLotStatus(item);
      } else if (col === 'Aging') {
        row[col] = calculateAging(item);
      } else if (col === 'Packing Date') {
        // Format date to show only date without time
        const dateStr = item['Packing Date'] || '';
        if (dateStr) {
          try {
            // Parse the date and format it as YYYY-MM-DD or any other date-only format
            const date = new Date(dateStr);
            if (!isNaN(date.getTime())) {
              // Format as YYYY-MM-DD
              const year = date.getFullYear();
              const month = String(date.getMonth() + 1).padStart(2, '0');
              const day = String(date.getDate()).padStart(2, '0');
              row[col] = `${year}-${month}-${day}`;
            } else {
              // If date parsing fails, keep original value
              row[col] = dateStr;
            }
          } catch (e) {
            row[col] = dateStr;
          }
        } else {
          row[col] = '';
        }
      } else {
        row[col] = item[col] || '';
      }
    });
    return row;
  });

  const worksheet = XLSX.utils.json_to_sheet(dataToExport);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "PackingReport");
  
  // Generate filename with current date
  const today = new Date();
  const dateStr = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
  XLSX.writeFile(workbook, `PackingReport_${dateStr}.xlsx`);
};

  // Export to PDF Function
  const exportToPDF = () => {
    if (filteredData.length === 0) {
      alert('No data available to export.');
      return;
    }

    // Create PDF in landscape mode
    const doc = new jsPDF({
      orientation: 'landscape',
      unit: 'mm',
      format: 'a4'
    });

    // Colors
    const headerColor = [15, 76, 129]; // Navy Blue
    const borderColor = [0, 0, 0]; // Light gray border
    const textColor = [17, 24, 39]; // Dark text
    const accentColor = [59, 130, 246]; // Blue accent
    const highlightColor = [16, 185, 129]; // Green for highlighting
    const remarksColor = [239, 68, 68]; // Red for remarks

    // Page dimensions
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const margin = 10;
    const contentWidth = pageWidth - (margin * 2);

    // Variable to track current Y position
    let currentY = 30; // Start Y position

    // Function to draw header (will be called on each page)
    const drawHeader = () => {
      // White background for header area
      doc.setFillColor(255, 255, 255);
      doc.rect(0, 0, pageWidth, 25, 'F');
      
      // Navy blue header bar
      doc.setFillColor(...headerColor);
      doc.rect(0, 0, pageWidth, 20, 'F');
      
      // Dynamic title based on Hold Lots filter
      const title = filters.holdLots ? 'PACKING HOLD LOTS REPORT' : 'PACKING ALLOTED REPORT';
      doc.setFontSize(18);
      doc.setTextColor(255, 255, 255);
      doc.setFont('Times New Roman', 'bold');
      doc.text(title, pageWidth / 2, 12, { align: 'center' });
      
      // Date
      doc.setFontSize(9);
      doc.setFont('helvetica', 'normal');
      const reportDate = new Date().toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'short',
        day: 'numeric'
      });
      doc.text(`Date: ${reportDate}`, margin, 25);
      
      // Page info
      doc.setFontSize(8);
      doc.setTextColor(100, 100, 100);
    };

    // Draw header on first page
    drawHeader();

    // Get today's date in multiple formats for comparison
    const today = new Date();
    const todayStr = today.toISOString().split('T')[0];
    const todayDisplay = today.toLocaleDateString('en-US', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit'
    });
    
    // Helper function to extract date from timestamp or packing date
    const extractDate = (record) => {
      // First try to get the timestamp
      const timestamp = record['Timestamp'];
      if (timestamp) {
        const datePart = timestamp.split(' ')[0];
        // Convert to mm/dd/yyyy format
        const parts = datePart.split('/');
        if (parts.length === 3) {
          return `${parts[1]}/${parts[0]}/${parts[2]}`;
        }
      }
      
      // Fallback to packing date
      return record['Packing Date'] || '';
    };

    // Helper function to find the most common manpower value (mode)
    const findMostCommonManpower = (manpowerValues) => {
      if (manpowerValues.length === 0) return 0;
      
      // Count frequency of each manpower value
      const frequency = {};
      manpowerValues.forEach(value => {
        frequency[value] = (frequency[value] || 0) + 1;
      });
      
      // Find the most common value
      let mostCommon = manpowerValues[0];
      let maxCount = 0;
      
      Object.entries(frequency).forEach(([value, count]) => {
        if (count > maxCount) {
          maxCount = count;
          mostCommon = parseInt(value);
        }
      });
      
      return mostCommon;
    };

    // Calculate supervisor summary with manpower logic
    const supervisorSummary = filteredData.reduce((acc, item) => {
      const supervisor = item['Packing Supervisor'] || 'Unassigned';
      const pieces = parseInt(item['Total Pcs']) || 0;
      const manpower = parseInt(item['Total Manpower']) || 0;
      const recordDate = extractDate(item);
      const isToday = recordDate === todayDisplay;
      
      if (!acc[supervisor]) {
        acc[supervisor] = {
          totalPieces: 0,
          lotCount: 0,
          todayManpowerValues: [],  // Store individual manpower values for today
          recentManpowerValues: [], // Store individual manpower values for recent dates
          todayDates: new Set(),
          otherDates: new Set()
        };
      }
      
      acc[supervisor].totalPieces += pieces;
      acc[supervisor].lotCount += 1;
      
      // Store manpower value with its date
      if (isToday) {
        acc[supervisor].todayManpowerValues.push(manpower);
        acc[supervisor].todayDates.add(recordDate);
      } else {
        acc[supervisor].recentManpowerValues.push({
          manpower: manpower,
          date: recordDate
        });
        acc[supervisor].otherDates.add(recordDate);
      }
      
      return acc;
    }, {});

    // Determine the manpower to show for each supervisor
    Object.keys(supervisorSummary).forEach(supervisor => {
      const summary = supervisorSummary[supervisor];
      
      // Try to get today's manpower
      if (summary.todayManpowerValues.length > 0) {
        // Find the most common manpower value for today
        summary.displayManpower = findMostCommonManpower(summary.todayManpowerValues);
        summary.manpowerSource = 'today';
        summary.manpowerDate = Array.from(summary.todayDates)[0] || todayDisplay;
      } 
      // If no today's data, look for most recent data
      else if (summary.recentManpowerValues.length > 0) {
        // Group by date to find the most recent date with data
        const groupedByDate = {};
        summary.recentManpowerValues.forEach(item => {
          if (!groupedByDate[item.date]) {
            groupedByDate[item.date] = [];
          }
          groupedByDate[item.date].push(item.manpower);
        });
        
        // Sort dates (most recent first)
        const sortedDates = Object.keys(groupedByDate).sort((a, b) => {
          const dateA = parseDate(a);
          const dateB = parseDate(b);
          return dateB - dateA;
        });
        
        // Get most recent date
        const mostRecentDate = sortedDates[0];
        if (mostRecentDate && groupedByDate[mostRecentDate]) {
          // Find most common manpower for the most recent date
          summary.displayManpower = findMostCommonManpower(groupedByDate[mostRecentDate]);
          summary.manpowerSource = 'recent';
          summary.manpowerDate = mostRecentDate;
        } else {
          summary.displayManpower = 0;
          summary.manpowerSource = 'none';
          summary.manpowerDate = '';
        }
      } else {
        summary.displayManpower = 0;
        summary.manpowerSource = 'none';
        summary.manpowerDate = '';
      }
    });

    // Helper function to parse date string to Date object
    function parseDate(dateStr) {
      if (!dateStr) return new Date(0);
      
      try {
        // Format: mm/dd/yyyy
        if (dateStr.includes('/')) {
          const parts = dateStr.split('/');
          if (parts.length === 3) {
            return new Date(parts[2], parts[1] - 1, parts[0]);
          }
        }
        
        // Format: yyyy-mm-dd
        if (dateStr.includes('-')) {
          const parts = dateStr.split('-');
          if (parts.length === 3) {
            return new Date(parts[0], parts[1] - 1, parts[2]);
          }
        }
        
        return new Date(dateStr);
      } catch (error) {
        console.error('Error parsing date:', dateStr, error);
        return new Date(0);
      }
    }

    // Sort supervisors by total pieces (descending)
    const sortedSupervisors = Object.entries(supervisorSummary)
      .sort(([, a], [, b]) => b.totalPieces - a.totalPieces);

    // Calculate totals for all supervisors
    const totalPiecesAll = sortedSupervisors.reduce((total, [supervisor, data]) => {
      return total + data.totalPieces;
    }, 0);
    
    // Count total manpower (using display manpower per supervisor)
    const totalManpowerAll = sortedSupervisors.reduce((total, [supervisor, data]) => {
      return total + data.displayManpower;
    }, 0);

    // ===================== MAIN DATA TABLE =====================
    
    // Calculate column widths to use full page width
    const columnWidths = {
      0: 17,   // Lot # - 17mm
      1: 35,   // Fabric - 35mm
      2: 25,   // Garment - 25mm
      3: 35,   // Style - 35mm
      4: 23,   // Brand - 23mm
      5: 15,   // Pcs - 15mm
      6: 25,   // Issue Date - 25mm (Changed header from Packing Date)
      7: 30,   // Supervisor - 30mm
      8: 15,   // Aging - 15mm
      9: 37,   // Remarks - 37mm
      10: 20   // Stitching Sup - 20mm
    };

    // Verify total width fits page
    const totalWidth = Object.values(columnWidths).reduce((a, b) => a + b, 0);
    // Adjust if total width is less than content width
    if (totalWidth < contentWidth) {
      const extraSpace = contentWidth - totalWidth;
      // Distribute extra space to wider columns
      columnWidths[9] += Math.floor(extraSpace * 0.4); // 40% to Remarks
      columnWidths[3] += Math.floor(extraSpace * 0.3); // 30% to Style
      columnWidths[1] += Math.floor(extraSpace * 0.2); // 20% to Fabric
      columnWidths[10] += Math.floor(extraSpace * 0.1); // 10% to Stitching Sup
    }

    // Prepare table headers with fixed width styling
    // CHANGED: "PACKING DATE" to "ISSUE DATE" in header
    const headers = [
      [
        { content: 'LOT #', styles: { fontStyle: 'bold', fillColor: headerColor, textColor: [255, 255, 255], cellWidth: columnWidths[0], halign: 'center' } },
        { content: 'FABRIC', styles: { fontStyle: 'bold', fillColor: headerColor, textColor: [255, 255, 255], cellWidth: columnWidths[1], halign: 'center' } },
        { content: 'GARMENT', styles: { fontStyle: 'bold', fillColor: headerColor, textColor: [255, 255, 255], cellWidth: columnWidths[2], halign: 'center' } },
        { content: 'STYLE', styles: { fontStyle: 'bold', fillColor: headerColor, textColor: [255, 255, 255], cellWidth: columnWidths[3], halign: 'center' } },
        { content: 'BRAND', styles: { fontStyle: 'bold', fillColor: headerColor, textColor: [255, 255, 255], cellWidth: columnWidths[4], halign: 'center' } },
        { content: 'PCS', styles: { fontStyle: 'bold', fillColor: headerColor, textColor: [255, 255, 255], cellWidth: columnWidths[5], halign: 'center' } },
        // CHANGED: Header text from "PACKING DATE" to "ISSUE DATE"
        { content: 'ISSUE DATE', styles: { fontStyle: 'bold', fillColor: headerColor, textColor: [255, 255, 255], cellWidth: columnWidths[6], halign: 'center' } },
        { content: 'SUPERVISOR', styles: { fontStyle: 'bold', fillColor: headerColor, textColor: [255, 255, 255], cellWidth: columnWidths[7], halign: 'center' } },
        { content: 'AGING', styles: { fontStyle: 'bold', fillColor: headerColor, textColor: [255, 255, 255], cellWidth: columnWidths[8], halign: 'center' } },
        { content: 'REMARKS', styles: { fontStyle: 'bold', fillColor: headerColor, textColor: [255, 255, 255], cellWidth: columnWidths[9], halign: 'center' } },
        { content: 'STITCHING SUP', styles: { fontStyle: 'bold', fillColor: headerColor, textColor: [255, 255, 255], cellWidth: columnWidths[10], halign: 'center' } }
      ]
    ];

    // Prepare table body with proper formatting
    // Note: Still using 'Packing Date' data but header shows as 'Issue Date'
    const body = filteredData.map(item => [
      { content: item['Lot Number'] || 'N/A', styles: { cellWidth: columnWidths[0], fontStyle: 'bold', fontSize: 9, halign: 'center',textColor: 'red' } },
      { content: item['Fabric'] || 'N/A', styles: { cellWidth: columnWidths[1], fontSize: 9, halign: 'center' } },
      { content: item['Garment Type'] || 'N/A', styles: { cellWidth: columnWidths[2], fontSize: 9, halign: 'center' } },
      { content: item['Style'] || 'N/A', styles: { cellWidth: columnWidths[3], fontSize: 9, halign: 'center' } },
      { content: item['BRAND'] || 'N/A', styles: { cellWidth: columnWidths[4], fontSize: 9, halign: 'center' } },
      { content: item['Total Pcs'] || '0', styles: { cellWidth: columnWidths[5], halign: 'center', fontSize: 10 ,fontStyle: ' bold', textColor:'blue'} },
      // Data still comes from 'Packing Date' field
      { content: item['Packing Date'] || 'N/A', styles: { cellWidth: columnWidths[6], fontSize: 9, halign: 'center' } },
      { content: item['Packing Supervisor'] || 'N/A', styles: { cellWidth: columnWidths[7], fontSize: 9, halign: 'center' } },
      { 
        content: item['Aging'] || '0', 
        styles: { 
          cellWidth: columnWidths[8],
          halign: 'center',
          fontStyle: 'bold',
          fontSize: 9,
          textColor: getAgingPDFColor(item['Aging'])
        }
      },
      { 
        content: getRecentRemarks(item).substring(0, 30) || 'No remarks', 
        styles: { 
          cellWidth: columnWidths[9],
          fontSize: 8,
          halign: 'center',
          fontStyle: 'bold',
          textColor: remarksColor // RED TEXT FOR REMARKS
        }
      },
      { content: item['STITCHING SUPERVISOR'] || 'N/A', styles: { cellWidth: columnWidths[10], fontSize: 9, halign: 'center' } }
    ]);

    // Track the Y position manually
    let lastAutoTableY = currentY;

    // Create the table with precise styling
    autoTable(doc, {
      startY: currentY,
      head: headers,
      body: body,
      theme: 'grid',
      headStyles: {
        fillColor: headerColor,
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        fontSize: 8,
        cellPadding: { top: 3, right: 2, bottom: 3, left: 2 },
        lineWidth: 0.5,
        lineColor: headerColor,
        halign: 'center',
        valign: 'middle',
        borderColor: 'black'
      },
      bodyStyles: {
        fontSize: 8,
        cellPadding: { top: 2, right: 2, bottom: 2, left: 2 },
        lineWidth: 0.3,
        lineColor: borderColor,
        textColor: textColor,
        fillColor: [255, 255, 255],
        font: 'helvetica',
        valign: 'middle'
      },
      styles: {
        fontSize: 8,
        cellPadding: { top: 2, right: 2, bottom: 2, left: 2 },
        lineWidth: 0.3,
        lineColor: borderColor,
        overflow: 'linebreak',
        font: 'helvetica',
        valign: 'middle'
      },
      columnStyles: {
        0: { cellWidth: columnWidths[0], fontStyle: 'bold', halign: 'center' },
        1: { cellWidth: columnWidths[1], halign: 'center' },
        2: { cellWidth: columnWidths[2], halign: 'center' },
        3: { cellWidth: columnWidths[3], halign: 'center' },
        4: { cellWidth: columnWidths[4], halign: 'center' },
        5: { cellWidth: columnWidths[5], halign: 'center' },
        6: { cellWidth: columnWidths[6], halign: 'center' },
        7: { cellWidth: columnWidths[7], halign: 'center' },
        8: { cellWidth: columnWidths[8], halign: 'center' },
        9: { cellWidth: columnWidths[9], fontSize: 8, halign: 'center', textColor: remarksColor, fontStyle: 'bold' }, // Red remarks
        10: { cellWidth: columnWidths[10], halign: 'center' }
      },
      margin: { top: currentY, left: margin, right: margin },
      tableWidth: contentWidth,
      showHead: 'everyPage',
      showFoot: false,
      pageBreak: 'auto',
      rowPageBreak: 'avoid',
      tableLineWidth: 0.5,
      tableLineColor: borderColor,
      didDrawPage: function(data) {
        // Draw header on every page
        if (data.pageNumber > 1) {
          drawHeader();
        }
        
        // Add page number at bottom
        const pageCount = doc.internal.getNumberOfPages();
        doc.setFontSize(8);
        doc.setTextColor(100, 100, 100);
        doc.text(
          `Page ${data.pageNumber} of ${pageCount}`,
          pageWidth / 2,
          pageHeight - 5,
          { align: 'center' }
        );
        
        // Add record count on first page
        if (data.pageNumber === 1) {
          doc.setFontSize(8);
          doc.setTextColor(100, 100, 100);
          doc.text(
            `Total Records: ${filteredData.length}`,
            pageWidth - margin,
            25,
            { align: 'right' }
          );
        }
        
        // Update last Y position
        if (data.cursor && data.cursor.y) {
          lastAutoTableY = data.cursor.y;
        }
      }
    });

    // ===================== PROFESSIONAL SUPERVISOR SUMMARY =====================
    
    // Add supervisor summary section after the table
    let summaryStartY = lastAutoTableY + 15;
    
    // Check if we need a new page
    if (summaryStartY > pageHeight - 50) {
      doc.addPage();
      summaryStartY = 30;
      drawHeader();
    }
    
    // Summary Section Header
    doc.setFontSize(13);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(...headerColor);
    doc.text('SUPERVISOR WORKLOAD DISTRIBUTION', pageWidth / 2, summaryStartY, { align: 'center' });
    
    // Today's date note
    doc.setFontSize(8);
    doc.setFont('helvetica', 'italic');
    doc.setTextColor(100, 100, 100);
    
    // Add note about manpower calculation
    const manpowerNote = "*Manpower shown per supervisor (most common value) | Pcs/Man = Total Pieces ÷ Manpower";
    doc.text(manpowerNote, pageWidth / 2, summaryStartY + 4, { align: 'center' });
    
    // Subtle underline
    doc.setLineWidth(0.5);
    doc.setDrawColor(220, 220, 220);
    doc.line(pageWidth / 2 - 80, summaryStartY + 6, pageWidth / 2 + 80, summaryStartY + 6);
    
    let y = summaryStartY + 12;
    
    // Prepare summary body data WITH PCS/MAN column
    const summaryBody = sortedSupervisors.map(([supervisor, data], index) => {
      const percentage = totalPiecesAll > 0 ? ((data.totalPieces / totalPiecesAll) * 100).toFixed(1) : '0.0';
      
      // Get manpower display information
      let manpowerDisplay = data.displayManpower > 0 ? data.displayManpower.toString() : '-';
      let manpowerColor = textColor;
      let manpowerNote = '';
      
      // Calculate Pcs/Man (only if we have valid manpower)
      let pcsPerMan = '-';
      let pcsPerManColor = textColor;
      if (data.displayManpower > 0) {
        const pcsPerManValue = Math.round(data.totalPieces / data.displayManpower);
        pcsPerMan = pcsPerManValue.toLocaleString();
        
        // Color coding based on productivity
        if (pcsPerManValue > 500) {
          pcsPerManColor = [16, 185, 129]; // Green for high productivity
        } else if (pcsPerManValue > 300) {
          pcsPerManColor = [245, 158, 11]; // Orange for medium productivity
        } else {
          pcsPerManColor = [239, 68, 68]; // Red for low productivity
        }
      }
      
      if (data.displayManpower > 0) {
        if (data.manpowerSource === 'today') {
          manpowerColor = accentColor;
          manpowerNote = `(Today: ${data.manpowerDate})`;
        } else if (data.manpowerSource === 'recent') {
          manpowerColor = [245, 158, 11]; // Orange for recent data
          manpowerNote = `(Recent: ${data.manpowerDate})`;
        }
      } else {
        manpowerColor = [150, 150, 150]; // Gray for no data
        manpowerNote = '(No data)';
      }
      
      // Determine row background color
      let rowBackground = [255, 255, 255];
      if (data.totalPieces === Math.max(...sortedSupervisors.map(([, d]) => d.totalPieces))) {
        rowBackground = [245, 247, 255]; // Light blue for top performer
      } else if (index % 2 === 0) {
        rowBackground = [250, 250, 250]; // Light gray for even rows
      }
      
      // Determine text color for percentage
      let percentageColor = textColor;
      const percentageNum = parseFloat(percentage);
      if (percentageNum > 30) {
        percentageColor = [16, 185, 129]; // Green for high percentage
      } else if (percentageNum > 15) {
        percentageColor = [245, 158, 11]; // Orange for medium percentage
      }
      
      return [
        { 
          content: supervisor, 
          styles: { 
            halign: 'left', 
            fontSize: 9, 
            fillColor: rowBackground,
            cellPadding: { top: 4, right: 4, bottom: 4, left: 4 },
            fontStyle: 'bold'
          } 
        },
        { 
          content: data.lotCount.toString(), 
          styles: { 
            halign: 'center', 
            fontSize: 9, 
            fillColor: rowBackground,
            cellPadding: { top: 4, right: 4, bottom: 4, left: 4 },
            fontStyle: 'bold'
          } 
        },
        { 
          content: data.totalPieces.toLocaleString(), 
          styles: { 
            halign: 'center', 
            fontSize: 10, 
            fontStyle: 'bold', 
            fillColor: rowBackground,
            cellPadding: { top: 4, right: 4, bottom: 4, left: 4 }
        } 
        },
        { 
          content: manpowerDisplay, 
          styles: { 
            halign: 'center', 
            fontSize: 10, 
            fillColor: rowBackground, 
            fontStyle: 'bold',
            textColor: manpowerColor,
            cellPadding: { top: 4, right: 4, bottom: 4, left: 4 }
          } 
        },
        // NEW: PCS/MAN column
        { 
          content: pcsPerMan, 
          styles: { 
            halign: 'center', 
            fontSize: 10, 
            fillColor: rowBackground, 
            fontStyle: 'bold',
            textColor: pcsPerManColor,
            cellPadding: { top: 4, right: 4, bottom: 4, left: 4 }
          } 
        },
        { 
          content: `${percentage}%`, 
          styles: { 
            halign: 'center', 
            fontSize: 10, 
            fontStyle: 'bold', 
            fillColor: rowBackground, 
            textColor: percentageColor,
            cellPadding: { top: 4, right: 4, bottom: 4, left: 4 }
          } 
        }
      ];
    });

    // Add total row
    const totalRowBackground = [240, 240, 240];
    
    // Calculate overall Pcs/Man for total row
    const overallPcsPerMan = totalManpowerAll > 0 ? Math.round(totalPiecesAll / totalManpowerAll) : '-';
    let overallPcsPerManColor = textColor;
    if (overallPcsPerMan !== '-') {
      if (overallPcsPerMan > 500) {
        overallPcsPerManColor = [16, 185, 129];
      } else if (overallPcsPerMan > 300) {
        overallPcsPerManColor = [245, 158, 11];
      } else {
        overallPcsPerManColor = [239, 68, 68];
      }
    }
    
    summaryBody.push([
      { 
        content: 'TOTAL', 
        styles: { 
          halign: 'left', 
          fontSize: 11, 
          fontStyle: 'bold', 
          fillColor: totalRowBackground,
          cellPadding: { top: 4, right: 4, bottom: 4, left: 4 }
        } 
      },
      { 
        content: filteredData.length.toString(), 
        styles: { 
          halign: 'center', 
          fontSize: 11, 
          fontStyle: 'bold', 
          fillColor: totalRowBackground,
          cellPadding: { top: 4, right: 4, bottom: 4, left: 4 }
        } 
      },
      { 
        content: totalPiecesAll.toLocaleString(), 
        styles: { 
          halign: 'center', 
          fontSize: 11, 
          fontStyle: 'bold', 
          fillColor: totalRowBackground,
          cellPadding: { top: 4, right: 4, bottom: 4, left: 4 }
        } 
      },
      { 
        content: totalManpowerAll > 0 ? totalManpowerAll.toString() : '-', 
        styles: { 
          halign: 'center', 
          fontSize: 11, 
          fontStyle: 'bold', 
          fillColor: totalRowBackground, 
          textColor: totalManpowerAll > 0 ? accentColor : [150, 150, 150],
          cellPadding: { top: 4, right: 4, bottom: 4, left: 4 }
        } 
      },
      // NEW: Total Pcs/Man column
      { 
        content: overallPcsPerMan !== '-' ? overallPcsPerMan.toLocaleString() : '-', 
        styles: { 
          halign: 'center', 
          fontSize: 11, 
          fontStyle: 'bold', 
          fillColor: totalRowBackground,
          textColor: overallPcsPerManColor,
          cellPadding: { top: 4, right: 4, bottom: 4, left: 4 }
        } 
      },
      { 
        content: '100%', 
        styles: { 
          halign: 'center', 
          fontSize: 11, 
          fontStyle: 'bold', 
          fillColor: totalRowBackground,
          cellPadding: { top: 4, right: 4, bottom: 4, left: 4 }
        } 
      }
    ]);

    // Calculate column widths for supervisor summary table
    // Adjusted widths to accommodate the new column
    const summaryColumnWidths = [
      65,  // Supervisor
      25,  // Lots
      40,  // Total Pieces
      35,  // Manpower
      35,  // NEW: Pcs/Man column
      40   // % of Total
    ];

    // Adjust if total is less than content width
    const totalSummaryWidth = summaryColumnWidths.reduce((a, b) => a + b, 0);
    if (totalSummaryWidth < contentWidth) {
      const extraSpace = contentWidth - totalSummaryWidth;
      // Distribute extra space proportionally
      summaryColumnWidths[0] += Math.floor(extraSpace * 0.3); // 30% to Supervisor
      summaryColumnWidths[2] += Math.floor(extraSpace * 0.25); // 25% to Total Pieces
      summaryColumnWidths[5] += Math.floor(extraSpace * 0.2); // 20% to % of Total
      summaryColumnWidths[1] += Math.floor(extraSpace * 0.15); // 15% to Lots
      summaryColumnWidths[4] += Math.floor(extraSpace * 0.1); // 10% to Pcs/Man
    }

    // Create the supervisor summary table with NEW column header
    autoTable(doc, {
      startY: y,
      head: [[
        { content: 'PACKING SUPERVISOR', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[0] } },
        { content: 'LOTS', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[1] } },
        { content: 'TOTAL PIECES', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[2] } },
        { content: 'MANPOWER*', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[3] } },
        // NEW: PCS/MAN column header
        { content: 'PCS/MAN', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[4] } },
        { content: '% OF PCS', styles: { halign: 'center', fontStyle: 'bold', cellWidth: summaryColumnWidths[5] } }
      ]],
      body: summaryBody.map(row => row.map((cell, colIndex) => ({
        content: cell.content,
        styles: {
          ...cell.styles,
          cellWidth: summaryColumnWidths[colIndex]
        }
      }))),
      theme: 'grid',
      headStyles: {
        fillColor: headerColor,
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        fontSize: 10,
        cellPadding: { top: 5, right: 4, bottom: 5, left: 4 },
        lineWidth: 0.5,
        lineColor: headerColor,
        halign: 'center',
        valign: 'middle'
      },
      bodyStyles: {
        fontSize: 9,
        cellPadding: { top: 4, right: 4, bottom: 4, left: 4 },
        lineWidth: 0.3,
        lineColor: borderColor,
        textColor: textColor,
        fillColor: [255, 255, 255],
        font: 'helvetica',
        valign: 'middle'
      },
      styles: {
        fontSize: 9,
        cellPadding: { top: 4, right: 4, bottom: 4, left: 4 },
        lineWidth: 0.3,
        lineColor: borderColor,
        overflow: 'linebreak',
        font: 'helvetica',
        valign: 'middle'
      },
      columnStyles: {
        0: { cellWidth: summaryColumnWidths[0], halign: 'center', fontStyle: 'bold' },
        1: { cellWidth: summaryColumnWidths[1], halign: 'center' },
        2: { cellWidth: summaryColumnWidths[2], halign: 'center' },
        3: { cellWidth: summaryColumnWidths[3], halign: 'center' },
        4: { cellWidth: summaryColumnWidths[4], halign: 'center' }, // NEW: Pcs/Man column
        5: { cellWidth: summaryColumnWidths[5], halign: 'center' }
      },
      margin: { left: margin, right: margin },
      tableWidth: contentWidth,
      showHead: 'everyPage',
      showFoot: false,
      pageBreak: 'auto',
      rowPageBreak: 'avoid',
      tableLineWidth: 0.5,
      tableLineColor: borderColor,
      didDrawPage: function(data) {
        // Update last Y position
        if (data.cursor && data.cursor.y) {
          lastAutoTableY = data.cursor.y;
        }
      }
    });

    // Add final insights
    const finalY = lastAutoTableY + 10;
    
    if (finalY < pageHeight - 20) {
      doc.setFontSize(9);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(100, 100, 100);
      
      // Overall statistics
      const avgPiecesPerLot = totalPiecesAll > 0 ? (totalPiecesAll / filteredData.length).toFixed(0) : 0;
      const supervisorsWithManpower = sortedSupervisors.filter(([_, data]) => data.displayManpower > 0).length;
      
      // Calculate average Pcs/Man (excluding supervisors without manpower)
      let avgPcsPerMan = '-';
      if (supervisorsWithManpower > 0) {
        const totalPcsForSupervisorsWithManpower = sortedSupervisors
          .filter(([_, data]) => data.displayManpower > 0)
          .reduce((total, [_, data]) => total + data.totalPieces, 0);
        
        const totalManpowerForSupervisorsWithManpower = sortedSupervisors
          .filter(([_, data]) => data.displayManpower > 0)
          .reduce((total, [_, data]) => total + data.displayManpower, 0);
        
        if (totalManpowerForSupervisorsWithManpower > 0) {
          avgPcsPerMan = Math.round(totalPcsForSupervisorsWithManpower / totalManpowerForSupervisorsWithManpower);
        }
      }
      
      const summaryText = `${sortedSupervisors.length} supervisors • ${supervisorsWithManpower} with manpower data • ${filteredData.length} lots • ${totalPiecesAll.toLocaleString()} total pieces • ${totalManpowerAll} total manpower`;
      doc.text(summaryText, pageWidth / 2, finalY, { align: 'center' });
      
      // Highlight top performers
      if (sortedSupervisors.length > 0) {
        const topSupervisor = sortedSupervisors[0];
        const topPercentage = ((topSupervisor[1].totalPieces / totalPiecesAll) * 100).toFixed(1);
        
        // Find supervisor with highest Pcs/Man (productivity)
        let mostProductiveSupervisor = null;
        let highestPcsPerMan = 0;
        
        sortedSupervisors.forEach(([name, data]) => {
          if (data.displayManpower > 0) {
            const pcsPerMan = Math.round(data.totalPieces / data.displayManpower);
            if (pcsPerMan > highestPcsPerMan) {
              highestPcsPerMan = pcsPerMan;
              mostProductiveSupervisor = { name, pcsPerMan };
            }
          }
        });
        
        doc.setFont('helvetica', 'italic');
        doc.setTextColor(...headerColor);
        doc.text(`Top Material Holder: ${topSupervisor[0]} - ${topPercentage}% of total pieces`, 
                 pageWidth / 2, finalY + 7, { align: 'center' });
        
        if (mostProductiveSupervisor) {
          doc.text(`Most Productive: ${mostProductiveSupervisor.name} - ${mostProductiveSupervisor.pcsPerMan.toLocaleString()} pcs/man`, 
                   pageWidth / 2, finalY + 14, { align: 'center' });
        }
      }
    }
    
    // Save with proper filename
    const fileName = filters.holdLots 
      ? `Packing_Hold_Lots_${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}.pdf`
      : `Packing_Report_${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}.pdf`;
    doc.save(fileName);
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
      filteredData.map(item => item['Packing Supervisor']).filter(Boolean)
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
        <p>Loading Packing Report...</p>
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
            padding: 24px;
            background: linear-gradient(135deg, #ffffff15 0%, #ffffff15 100%);
            min-height: 100vh;
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
          }
          
          .header {
            background: white;
            border-radius: 16px;
            padding: 24px;
            margin-bottom: 24px;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
            border-left: 6px solid #3B82F6;
          }
          
          .header-content {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
          }
          
          .header h1 {
            margin: 0;
            font-size: 28px;
            font-weight: 700;
            color: #1F2937;
            display: flex;
            align-items: center;
            gap: 12px;
          }
          
          .header h1 svg {
            color: #3B82F6;
          }
          
          .header-controls {
            display: flex;
            gap: 12px;
          }
          
          .icon-btn {
            background: #F3F4F6;
            border: none;
            border-radius: 12px;
            width: 48px;
            height: 48px;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            transition: all 0.2s;
            color: #4B5563;
          }
          
          .icon-btn:hover {
            background: #E5E7EB;
            transform: translateY(-2px);
          }
          
          .refresh-btn {
            color: #3B82F6;
          }
          
          .clear-btn {
            color: #EF4444;
          }
          
          .summary-cards {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
            gap: 20px;
            margin-bottom: 24px;
          }
          
          .summary-card {
            background: white;
            border-radius: 16px;
            padding: 24px;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
            transition: transform 0.2s;
            border-top: 4px solid;
          }
          
          .summary-card:hover {
            transform: translateY(-4px);
          }
          
          .summary-card.total-lots {
            border-top-color: #3B82F6;
          }
          
          .summary-card.total-pieces {
            border-top-color: #10B981;
          }
          
          .summary-card.total-supervisors {
            border-top-color: #8B5CF6;
          }
          
          .summary-card.avg-aging {
            border-top-color: #F59E0B;
          }
          
          .card-label {
            font-size: 14px;
            font-weight: 600;
            color: #6B7280;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 8px;
            display: flex;
            align-items: center;
            gap: 8px;
          }
          
          .card-value {
            font-size: 36px;
            font-weight: 800;
            color: #1F2937;
            margin: 0;
          }
          
          .filters-container {
            background: white;
            border-radius: 16px;
            padding: 24px;
            margin-bottom: 24px;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
          }
          
          .filters-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
          }
          
          .filters-title {
            font-size: 18px;
            font-weight: 700;
            color: #1F2937;
            display: flex;
            align-items: center;
            gap: 10px;
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
            font-size: 14px;
            font-weight: 600;
            color: #4B5563;
            margin-bottom: 8px;
          }
          
          .filter-input,
          .filter-select {
            padding: 12px 16px;
            border: 2px solid #E5E7EB;
            border-radius: 10px;
            font-size: 14px;
            transition: all 0.2s;
            background: white;
          }
          
          .filter-input:focus,
          .filter-select:focus {
            outline: none;
            border-color: #3B82F6;
            box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.1);
          }
          
          /* Checkbox styling for Hold Lots filter */
          .checkbox-filter {
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 12px 16px;
            border: 2px solid #E5E7EB;
            border-radius: 10px;
            cursor: pointer;
            transition: all 0.2s;
          }
          
          .checkbox-filter:hover {
            border-color: #3B82F6;
            background: #F9FAFB;
          }
          
          .checkbox-filter input[type="checkbox"] {
            width: 18px;
            height: 18px;
            cursor: pointer;
          }
          
          .checkbox-filter label {
            font-size: 14px;
            font-weight: 600;
            color: #4B5563;
            cursor: pointer;
            flex: 1;
          }
          
          .checkbox-filter.active {
            border-color: #EF4444;
            background: #FEF2F2;
          }
          
          .hold-indicator {
            background: #EF4444;
            color: white;
            padding: 4px 8px;
            border-radius: 6px;
            font-size: 12px;
            font-weight: 600;
            margin-left: 8px;
          }
          
          .table-container {
            background: white;
            border-radius: 16px;
            overflow: hidden;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
            overflow-x: auto;
          }
          
          .table-header {
            padding: 20px 24px;
            border-bottom: 1px solid #E5E7EB;
            display: flex;
            justify-content: space-between;
            align-items: center;
          }
          
          .table-title {
            font-size: 18px;
            font-weight: 700;
            color: #1F2937;
          }
          
          .records-count {
            font-size: 14px;
            color: #6B7280;
            font-weight: 500;
          }
          
          .data-table {
            width: 100%;
            border-collapse: collapse;
            min-width: 1500px;
          }
          
          /* Navy Blue Table Headers */
          .data-table thead {
            background: #0f4c81 !important; /* Navy Blue */
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
          }
          
          .data-table th {
            background-color: #0f4c81; /* Navy Blue */
            color: white !important;
            padding: 16px 20px;
            text-align: left;
            font-weight: 700;
            font-size: 14px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            border-bottom: none;
            position: sticky;
            top: 0;
            z-index: 10;
              border: 1px solid #F3F4F6;
          }
          
          .data-table td {
            padding: 16px 20px;
            border: 1px solid #F3F4F6;
            color: #1F2937;
            font-size: 14px;
          }
          
          .data-table tbody tr {
            transition: background-color 0.2s;
            cursor: pointer;
          }
          
          .data-table tbody tr:hover {
            background-color: #F9FAFB;
          }
          
          .status-badge {
            padding: 6px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            display: inline-block;
          }
          
          .aging-badge {
            padding: 6px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 700;
            color: white;
            display: inline-block;
            min-width: 60px;
            text-align: center;
          }
          
          .supervisor-tag {
            background: #DBEAFE;
            color: #1E40AF;
            padding: 4px 12px;
            border-radius: 12px;
            font-size: 12px;
            font-weight: 600;
            display: inline-block;
          }
          
          .stitching-supervisor-tag {
            background: #F0F9FF;
            color: #0369A1;
            padding: 4px 12px;
            border-radius: 12px;
            font-size: 12px;
            font-weight: 600;
            display: inline-block;
          }
          
          .brand-tag {
            background: #FEF3C7;
            color: #92400E;
            padding: 4px 12px;
            border-radius: 12px;
            font-size: 12px;
            font-weight: 600;
            display: inline-block;
          }
          
          .remarks-cell {
            max-width: 200px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
            font-size: 12px;
            color: #6B7280;
            background: #F9FAFB;
            padding: 8px 12px;
            border-radius: 8px;
            border-left: 3px solid #3B82F6;
          }
          
          .remarks-cell.hold {
            border-left-color: #EF4444;
            background: #FEF2F2;
            color: #7F1D1D;
          }
          
          .loading-container {
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            min-height: 400px;
            color: #3B82F6;
          }
          
          .spinner {
            width: 50px;
            height: 50px;
            border: 4px solid #E5E7EB;
            border-top: 4px solid #3B82F6;
            border-radius: 50%;
            animation: spin 1s linear infinite;
            margin-bottom: 20px;
          }
          
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          
          .error-message {
            background: #FEF2F2;
            border: 1px solid #FECACA;
            border-radius: 12px;
            padding: 24px;
            margin: 24px;
            text-align: center;
          }
          
          .error-message h3 {
            color: #DC2626;
            margin: 0 0 12px 0;
          }
          
          .error-message p {
            color: #7F1D1D;
            margin: 0 0 20px 0;
          }
          
          .retry-btn {
            background: #DC2626;
            color: white;
            border: none;
            padding: 10px 24px;
            border-radius: 10px;
            cursor: pointer;
            font-weight: 600;
            transition: background 0.2s;
          }
          
          .retry-btn:hover {
            background: #B91C1C;
          }
          
          .modal-overlay {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(0, 0, 0, 0.7);
            display: flex;
            justify-content: center;
            align-items: center;
            z-index: 1000;
            padding: 20px;
            backdrop-filter: blur(4px);
          }
          
          .modal-content {
            background: white;
            border-radius: 20px;
            max-width: 800px;
            width: 100%;
            max-height: 90vh;
            overflow-y: auto;
            box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.25);
            animation: slideIn 0.3s ease-out;
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
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 24px;
            border-radius: 20px 20px 0 0;
            display: flex;
            justify-content: space-between;
            align-items: center;
          }
          
          .modal-title {
            font-size: 24px;
            font-weight: 700;
            margin: 0;
          }
          
          .close-button {
            background: rgba(255, 255, 255, 0.2);
            border: none;
            color: white;
            font-size: 24px;
            cursor: pointer;
            width: 40px;
            height: 40px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: background 0.2s;
          }
          
          .close-button:hover {
            background: rgba(255, 255, 255, 0.3);
          }
          
          .modal-body {
            padding: 24px;
          }
          
          .row-details-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 32px;
          }
          
          .detail-card {
            background: #F9FAFB;
            border-radius: 12px;
            padding: 20px;
          }
          
          .detail-label {
            font-size: 12px;
            font-weight: 600;
            color: #6B7280;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 8px;
          }
          
          .detail-value {
            font-size: 16px;
            font-weight: 600;
            color: #1F2937;
          }
          
          .status-section {
            margin-bottom: 32px;
          }
          
          .section-title {
            font-size: 18px;
            font-weight: 700;
            color: #374151;
            margin-bottom: 20px;
            padding-bottom: 12px;
            border-bottom: 2px solid #E5E7EB;
          }
          
          .status-log-item {
            background: white;
            border: 1px solid #E5E7EB;
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 16px;
          }
          
          .log-item-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 12px;
          }
          
          .log-status {
            font-weight: 700;
            font-size: 16px;
            color: #1F2937;
          }
          
          .log-timestamp {
            font-size: 12px;
            color: #6B7280;
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
            font-size: 12px;
            color: #6B7280;
            margin-bottom: 4px;
          }
          
          .log-detail-value {
            font-size: 14px;
            color: #1F2937;
            font-weight: 500;
          }
          
          .no-data {
            text-align: center;
            padding: 60px 20px;
            color: #6B7280;
          }
          
          .no-data-icon {
            font-size: 48px;
            margin-bottom: 16px;
            opacity: 0.5;
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
              align-self: flex-end;
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
                  <div className="detail-label">Packing Date</div>
                  <div className="detail-value">{selectedRow['Packing Date']}</div>
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
                  <div className="detail-label">Packing Supervisor</div>
                  <div className="detail-value">{selectedRow['Packing Supervisor']}</div>
                </div>
                <div className="detail-card">
                  <div className="detail-label">Stitching Supervisor</div>
                  <div className="detail-value">{selectedRow['STITCHING SUPERVISOR'] || 'N/A'}</div>
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
            <h1>
              📦 Packing Dashboard
              <span style={{ fontSize: '14px', color: '#6B7280', fontWeight: 'normal' }}>
                (WIP & Completed Lots)
              </span>
            </h1>
            <div className="header-controls">
                <button 
    className="icon-btn" 
    style={{ backgroundColor: '#6B7280', color: 'white' }}
    onClick={() => window.history.back()}
    title="Go Back"
  >
    ←
  </button>
              <button 
                className="icon-btn" 
                style={{ backgroundColor: '#10B981', color: 'white' }}
                onClick={exportToExcel} 
                title="Download as Excel (with Status column)"
              >
                <span style={{ fontSize: '20px' }}>📊</span>
              </button>
              
              <button 
                className="icon-btn" 
                style={{ backgroundColor: '#EF4444', color: 'white' }}
                onClick={exportToPDF} 
                title="Download as PDF (without Status column)"
              >
                <span style={{ fontSize: '20px' }}>📄</span>
              </button>
              
              <button className="icon-btn refresh-btn" onClick={fetchData} title="Refresh data">
                ↻
              </button>
              <button className="icon-btn clear-btn" onClick={clearFilters} title="Clear all filters">
                ✕
              </button>
            </div>
          </div>
          
          {/* Summary Cards - CHANGED: Total Manpower card replaced with Total Packing Supervisor card */}
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
                <span>👤</span> Packing Supervisors
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
              <label className="filter-label">Packing Supervisor</label>
              <select
                className="filter-select"
                value={filters.supervisor}
                onChange={(e) => handleFilterChange('supervisor', e.target.value)}
              >
                <option value="">All Supervisors</option>
                {getUniqueValues('Packing Supervisor').map((supervisor, index) => (
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
              <label className="filter-label">Stitching Supervisor</label>
              <select
                className="filter-select"
                value={filters.stitchingSupervisor}
                onChange={(e) => handleFilterChange('stitchingSupervisor', e.target.value)}
              >
                <option value="">All Stitching Supervisors</option>
                {getUniqueValues('STITCHING SUPERVISOR').map((supervisor, index) => (
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
              Packing Lots Details
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
                <th>Packing Date</th>
                <th>Packing Supervisor</th>
                <th>Aging (Days)</th>
                <th>Status</th>
                <th>Recent Remarks</th>
                <th>Stitching Supervisor</th>
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
                      <td>{item['Packing Date']}</td>
                      <td><span className="supervisor-tag">{item['Packing Supervisor']}</span></td>
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
                        {item['STITCHING SUPERVISOR'] && (
                          <span className="stitching-supervisor-tag">
                            {item['STITCHING SUPERVISOR']}
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
                      No packing lots match your current filters.
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

export default DailyPackingReport;