import React, { useState, useEffect } from 'react';
import axios from 'axios';
import * as XLSX from 'xlsx';

const IssueToPacking = () => {
  const [data, setData] = useState([]);
  const [filteredData, setFilteredData] = useState([]);
  const [recentLots, setRecentLots] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [stats, setStats] = useState({
    totalRecords: 0,
    totalPcs: 0,
    packingSupervisors: [],
    garmentTypes: []
  });

  // Filters
  const [filters, setFilters] = useState({
    lotNumber: '',
    garmentType: '',
    fabric: '',
    style: '',
    packingSupervisor: '',
    startDate: '',
    endDate: ''
  });

  // Dropdown options
  const [dropdownOptions, setDropdownOptions] = useState({
    lotNumbers: [],
    garmentTypes: [],
    fabrics: [],
    styles: [],
    packingSupervisors: []
  });

  // Replace these with your actual values
  const API_KEY = 'AIzaSyAomDFBkOySlIxKWSKGHe6ATv9gvaBr7uk';
  const SPREADSHEET_ID = '1uo14nKO_yHu4AJ2rOgaJajuprcinj6xw1AUMFJ6_zYM';
  const SHEET_NAME = 'Issues';
  const RANGE = 'A:H';

  useEffect(() => {
    fetchData();
  }, []);

  useEffect(() => {
    applyFilters();
  }, [data, filters]);

const fetchData = async () => {
  try {
    setLoading(true);
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${SHEET_NAME}!${RANGE}?key=${API_KEY}`;
    
    const response = await axios.get(url);
    
    if (response.data.values) {
      const rows = response.data.values.slice(1); // Skip header row
      
      // Filter out completely empty rows
      const validRows = rows.filter(row => {
        // Check if row exists and is an array
        if (!row || !Array.isArray(row)) return false;
        
        // Check if row has ANY non-empty cell
        const hasData = row.some(cell => {
          return cell !== undefined && 
                 cell !== null && 
                 cell !== '' && 
                 (typeof cell !== 'string' || cell.trim() !== '');
        });
        
        return hasData;
      });
      
      console.log(`Loaded ${validRows.length} valid records (skipped ${rows.length - validRows.length} empty rows)`);
      
      const formattedData = validRows.map(row => ({
        timestamp: row[0] || '',
        lotNumber: row[1] || '',
        garmentType: row[2] || '',
        fabric: row[3] || '',
        style: row[4] || '',
        packingSupervisor: row[5] || '',
        packingDate: row[6] || '',
        totalPcs: parseInt(row[7]) || 0,
        // Calculate if lot is recently issued (within last 24 hours)
        isRecent: isRecentLot(row[0])
      }));
      
      setData(formattedData);
      
      // Separate recent lots
      const recent = formattedData.filter(item => item.isRecent);
      setRecentLots(recent);
      
      extractDropdownOptions(formattedData);
      calculateStats(formattedData);
    }
    setLoading(false);
  } catch (err) {
    setError('Error fetching data. Please check your API key and spreadsheet ID.');
    setLoading(false);
    console.error('Error:', err);
  }
};

  const isRecentLot = (timestamp) => {
    if (!timestamp) return false;
    const lotDate = new Date(timestamp);
    const now = new Date();
    const hoursDifference = (now - lotDate) / (1000 * 60 * 60);
    return hoursDifference <= 24;
  };

  const extractDropdownOptions = (dataArray) => {
    const lotNumbers = [...new Set(dataArray.map(item => item.lotNumber))].filter(Boolean).sort();
    const garmentTypes = [...new Set(dataArray.map(item => item.garmentType))].filter(Boolean).sort();
    const fabrics = [...new Set(dataArray.map(item => item.fabric))].filter(Boolean).sort();
    const styles = [...new Set(dataArray.map(item => item.style))].filter(Boolean).sort();
    const packingSupervisors = [...new Set(dataArray.map(item => item.packingSupervisor))].filter(Boolean).sort();

    setDropdownOptions({
      lotNumbers,
      garmentTypes,
      fabrics,
      styles,
      packingSupervisors
    });
  };

  const calculateStats = (dataArray) => {
    const totalPcs = dataArray.reduce((sum, item) => sum + item.totalPcs, 0);
    
    // Count supervisors with their data
    const supervisorCounts = dataArray.reduce((acc, item) => {
      if (item.packingSupervisor) {
        acc[item.packingSupervisor] = (acc[item.packingSupervisor] || 0) + 1;
      }
      return acc;
    }, {});
    
    const supervisors = Object.entries(supervisorCounts)
      .map(([name, count]) => ({ name, count }))
      .sort((a, b) => b.count - a.count);
    
    // Count garment types with their data
    const garmentTypeCounts = dataArray.reduce((acc, item) => {
      if (item.garmentType) {
        acc[item.garmentType] = (acc[item.garmentType] || 0) + 1;
      }
      return acc;
    }, {});
    
    const garmentTypes = Object.entries(garmentTypeCounts)
      .map(([type, count]) => ({ type, count }))
      .sort((a, b) => b.count - a.count);
    
    setStats({
      totalRecords: dataArray.length,
      totalPcs,
      packingSupervisors: supervisors,
      garmentTypes
    });
  };

  const applyFilters = () => {
    let filtered = [...data];

    if (filters.lotNumber) {
      filtered = filtered.filter(item => item.lotNumber === filters.lotNumber);
    }

    if (filters.garmentType) {
      filtered = filtered.filter(item => item.garmentType === filters.garmentType);
    }

    if (filters.fabric) {
      filtered = filtered.filter(item => item.fabric === filters.fabric);
    }

    if (filters.style) {
      filtered = filtered.filter(item => item.style === filters.style);
    }

    if (filters.packingSupervisor) {
      filtered = filtered.filter(item => item.packingSupervisor === filters.packingSupervisor);
    }

    if (filters.startDate) {
      filtered = filtered.filter(item => 
        new Date(item.packingDate) >= new Date(filters.startDate)
      );
    }

    if (filters.endDate) {
      filtered = filtered.filter(item => 
        new Date(item.packingDate) <= new Date(filters.endDate)
      );
    }

    setFilteredData(filtered);
    // Update recent lots based on filtered data
    const recent = filtered.filter(item => item.isRecent);
    setRecentLots(recent);
    calculateStats(filtered);
  };

  const handleFilterChange = (e) => {
    const { name, value } = e.target;
    setFilters(prev => ({
      ...prev,
      [name]: value
    }));
  };

  const resetFilters = () => {
    setFilters({
      lotNumber: '',
      garmentType: '',
      fabric: '',
      style: '',
      packingSupervisor: '',
      startDate: '',
      endDate: ''
    });
  };

const exportToExcel = () => {
  // Filter out empty rows before exporting
  const exportData = filteredData
    .filter(item => {
      // Check if the row has any meaningful data
      return (
        (item.lotNumber && item.lotNumber.trim() !== '') ||
        (item.garmentType && item.garmentType.trim() !== '') ||
        (item.fabric && item.fabric.trim() !== '') ||
        (item.style && item.style.trim() !== '') ||
        (item.packingSupervisor && item.packingSupervisor.trim() !== '') ||
        (item.packingDate && item.packingDate.trim() !== '') ||
        item.totalPcs > 0
      );
    })
    .map(item => ({
      'Timestamp': item.timestamp || '',
      'Lot Number': item.lotNumber || '',
      'Garment Type': item.garmentType || '',
      'Fabric': item.fabric || '',
      'Style': item.style || '',
      'Packing Supervisor': item.packingSupervisor || '',
      'Packing Date': item.packingDate || '',
      'Total Pcs': item.totalPcs || 0
    }));

  // Check if there's data to export
  if (exportData.length === 0) {
    alert('No valid data to export. All rows appear to be empty.');
    return;
  }

  const worksheet = XLSX.utils.json_to_sheet(exportData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Issue To Packing');
  XLSX.writeFile(workbook, `issue_to_packing_${new Date().toISOString().split('T')[0]}.xlsx`);
};
const exportToPDF = () => {
  // Filter out empty rows for PDF export too
  const pdfData = filteredData.filter(item => {
    return (
      (item.lotNumber && item.lotNumber.trim() !== '') ||
      (item.garmentType && item.garmentType.trim() !== '') ||
      (item.fabric && item.fabric.trim() !== '') ||
      (item.style && item.style.trim() !== '') ||
      (item.packingSupervisor && item.packingSupervisor.trim() !== '') ||
      (item.packingDate && item.packingDate.trim() !== '') ||
      item.totalPcs > 0
    );
  });

  // Check if there's data to export
  if (pdfData.length === 0) {
    alert('No valid data to export. All rows appear to be empty.');
    return;
  }

  const totalPcs = pdfData.reduce((sum, item) => sum + item.totalPcs, 0);
  const uniqueSupervisors = new Set(pdfData.map(item => item.packingSupervisor)).size;
  const uniqueGarmentTypes = new Set(pdfData.map(item => item.garmentType)).size;

  const printWindow = window.open('', '_blank');
  
  const content = `
    <!DOCTYPE html>
    <html>
    <head>
      <title>Issue to Packing Report</title>
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
        
        * {
          margin: 0;
          padding: 0;
          box-sizing: border-box;
        }
        
        body {
          font-family: 'Inter', sans-serif;
          line-height: 1.4;
          color: #333;
          background: #ffffff;
          min-height: 100vh;
          padding: 40px 20px;
        }
        
        .report-container {
          max-width: 1100px;
          margin: 0 auto;
          background: white;
          border: 1px solid #e5e7eb;
          box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
          position: relative;
        }
        
        .report-header {
          background: #ffffff;
          padding: 40px 40px 20px 40px;
          border-bottom: 3px solid #1e40af;
          position: relative;
        }
        
        .company-header {
          display: flex;
          justify-content: space-between;
          align-items: flex-end;
          margin-bottom: 30px;
        }
        
        .company-info {
          flex: 1;
        }
        
        .company-name {
          font-size: 28px;
          font-weight: 700;
          color: #1e293b;
          margin-bottom: 5px;
          letter-spacing: -0.5px;
        }
        
        .company-tagline {
          font-size: 14px;
          color: #64748b;
          font-weight: 400;
        }
        
        .report-info {
          text-align: right;
        }
        
        .report-title {
          font-size: 24px;
          font-weight: 600;
          color: #1e293b;
          margin-bottom: 5px;
        }
        
        .report-id {
          font-size: 12px;
          color: #64748b;
          font-weight: 500;
        }
        
        .report-meta {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-top: 25px;
          padding: 20px 0;
          border-top: 1px solid #e5e7eb;
          border-bottom: 1px solid #e5e7eb;
        }
        
        .meta-item {
          text-align: center;
          flex: 1;
        }
        
        .meta-label {
          font-size: 11px;
          color: #64748b;
          text-transform: uppercase;
          letter-spacing: 0.5px;
          font-weight: 600;
          margin-bottom: 4px;
        }
        
        .meta-value {
          font-size: 14px;
          color: #1e293b;
          font-weight: 500;
        }
        
        .section-header {
          padding: 15px 40px;
          background: #f8fafc;
          border-bottom: 1px solid #e5e7eb;
          font-size: 12px;
          font-weight: 600;
          color: #475569;
          text-transform: uppercase;
          letter-spacing: 0.5px;
        }
        
        .summary-section {
          padding: 30px 40px;
          background: #f8fafc;
          border-bottom: 1px solid #e5e7eb;
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          gap: 20px;
        }
        
        .summary-item {
          text-align: center;
        }
        
        .summary-value {
          font-size: 24px;
          font-weight: 700;
          color: #1e40af;
          margin-bottom: 4px;
          line-height: 1;
        }
        
        .summary-label {
          font-size: 10px;
          color: #64748b;
          text-transform: uppercase;
          letter-spacing: 0.5px;
          font-weight: 600;
        }
        
        .table-container {
          padding: 0;
          overflow: hidden;
        }
        
        .data-table {
          width: 100%;
          border-collapse: collapse;
          font-size: 11px;
          table-layout: fixed;
          border: 1px solid #e5e7eb;
        }
        
        .data-table thead {
          background: linear-gradient(135deg, #1e40af 0%, #3b82f6 100%);
        }
        
        .data-table th {
          color: #ffffff;
          font-weight: 600;
          text-transform: uppercase;
          font-size: 10px;
          letter-spacing: 0.5px;
          padding: 12px 8px;
          text-align: left;
          border-right: 1px solid rgba(255, 255, 255, 0.1);
          border-bottom: 2px solid #1e3a8a;
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
        }
        
        .data-table th:nth-child(1) { width: 10%; }
        .data-table th:nth-child(2) { width: 10%; }
        .data-table th:nth-child(3) { width: 12%; }
        .data-table th:nth-child(4) { width: 12%; }
        .data-table th:nth-child(5) { width: 15%; }
        .data-table th:nth-child(6) { width: 12%; }
        .data-table th:nth-child(7) { width: 10%; }
        .data-table th:nth-child(8) { width: 8%; }
        
        .data-table th:last-child {
          border-right: none;
          text-align: right;
        }
        
        .data-table td {
          padding: 10px 8px;
          border-bottom: 1px solid #e5e7eb;
          border-right: 1px solid #e5e7eb;
          vertical-align: middle;
          color: #334155;
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: nowrap;
        }
        
        .data-table td:last-child {
          border-right: none;
          text-align: right;
          font-family: 'Courier New', monospace;
          font-weight: 600;
        }
        
        .data-table tbody tr:last-child td {
          border-bottom: none;
        }
        
        .data-table tbody tr:hover {
          background-color: #f8fafc;
        }
        
        .timestamp-cell {
          font-size: 10px;
          color: #64748b;
        }
        
        .lot-number-cell {
          font-weight: 600;
          color: #1e40af;
        }
        
        .recent-lot {
          background: linear-gradient(90deg, rgba(16, 185, 129, 0.05) 0%, rgba(16, 185, 129, 0.02) 100%);
          border-left: 3px solid #10b981;
        }
        
        .recent-indicator {
          display: inline-block;
          width: 6px;
          height: 6px;
          background: #10b981;
          border-radius: 50%;
          margin-right: 4px;
          vertical-align: middle;
        }
        
        .supervisor-cell {
          display: inline-block;
          padding: 4px 10px;
          background: #e0f2fe;
          color: #0369a1;
          border-radius: 12px;
          font-size: 10px;
          font-weight: 500;
          white-space: nowrap;
        }
        
        .pcs-cell {
          text-align: right;
          font-weight: 600;
          font-family: 'Courier New', monospace;
        }
        
        .total-row {
          background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
          font-weight: 600;
        }
        
        .total-row td {
          border-top: 2px solid #cbd5e1;
          padding: 12px 8px;
          color: #1e293b;
        }
        
        .total-row td:first-child {
          text-align: right;
          font-size: 11px;
          padding-right: 20px;
        }
        
        .total-row td:last-child {
          font-size: 12px;
          color: #1e40af;
        }
        
        .signature-section {
          margin-top: 40px;
          display: grid;
          grid-template-columns: repeat(3, 1fr);
          gap: 30px;
          padding-top: 30px;
          border-top: 1px solid #e5e7eb;
        }
        
        .signature-line {
          border-top: 1px solid #cbd5e1;
          margin-top: 30px;
          padding-top: 8px;
          font-size: 10px;
          color: #64748b;
          text-align: center;
        }
        
        .page-info {
          text-align: center;
          margin-top: 30px;
          font-size: 10px;
          color: #94a3b8;
        }
        
        .watermark {
          position: fixed;
          top: 50%;
          left: 50%;
          transform: translate(-50%, -50%) rotate(-45deg);
          font-size: 80px;
          opacity: 0.03;
          color: #000;
          pointer-events: none;
          z-index: -1;
          font-weight: 900;
          white-space: nowrap;
        }
        
        @media print {
          body {
            padding: 0;
            background: white;
          }
          
          .report-container {
            box-shadow: none;
            border: none;
            max-width: 100%;
            margin: 0;
            padding: 0;
          }
          
          .data-table th {
            -webkit-print-color-adjust: exact;
            color-adjust: exact;
          }
          
          .data-table thead {
            display: table-header-group;
          }
          
          .data-table {
            font-size: 9px;
          }
          
          .data-table th,
          .data-table td {
            padding: 8px 6px;
          }
        }
      </style>
    </head>
    <body>
      <div class="watermark">CONFIDENTIAL</div>
      
      <div class="report-container">
        <div class="report-header">
          <div class="company-header">
            <div class="company-info">
              <div class="company-name">PACKING ISSUE</div>
              <div class="company-tagline">Issue to Packing Management System</div>
            </div>
            <div class="report-info">
              <div class="report-title">Issue to Packing Report</div>
              <div class="report-id">DOC-${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, '0')}-${Math.floor(Math.random() * 1000).toString().padStart(3, '0')}</div>
            </div>
          </div>
          
          <div class="report-meta">
            <div class="meta-item">
              <div class="meta-label">Report Period</div>
              <div class="meta-value">${filters.startDate || 'Start'} - ${filters.endDate || 'End'}</div>
            </div>
            <div class="meta-item">
              <div class="meta-label">Generated On</div>
              <div class="meta-value">${new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' })}</div>
            </div>
            <div class="meta-item">
              <div class="meta-label">Generated At</div>
              <div class="meta-value">${new Date().toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' })}</div>
            </div>
            <div class="meta-item">
              <div class="meta-label">Total Records</div>
              <div class="meta-value">${pdfData.length}</div>
            </div>
          </div>
        </div>
        
        <div class="section-header">
          Executive Summary
        </div>
        
        <div class="summary-section">
          <div class="summary-item">
            <div class="summary-value">${pdfData.length}</div>
            <div class="summary-label">Total Lots</div>
          </div>
          <div class="summary-item">
            <div class="summary-value">${totalPcs.toLocaleString()}</div>
            <div class="summary-label">Total Pieces</div>
          </div>
          <div class="summary-item">
            <div class="summary-value">${uniqueSupervisors}</div>
            <div class="summary-label">Supervisors</div>
          </div>
          <div class="summary-item">
            <div class="summary-value">${uniqueGarmentTypes}</div>
            <div class="summary-label">Product Types</div>
          </div>
        </div>
        
        <div class="section-header">
          Detailed Issue to Packing Records
        </div>
        
        <div class="table-container">
          <table class="data-table">
            <thead>
              <tr>
                <th>TIMESTAMP</th>
                <th>LOT NUMBER</th>
                <th>GARMENT TYPE</th>
                <th>FABRIC</th>
                <th>STYLE</th>
                <th>SUPERVISOR</th>
                <th>PACKING DATE</th>
                <th>TOTAL PCS</th>
              </tr>
            </thead>
            <tbody>
              ${pdfData.map((item, index) => `
                <tr class="${item.isRecent ? 'recent-lot' : ''}">
                  <td class="timestamp-cell">
                    ${item.isRecent ? '<span class="recent-indicator"></span>' : ''}
                    ${item.timestamp ? new Date(item.timestamp).toLocaleDateString('en-US', { 
                      month: 'short', 
                      day: 'numeric', 
                      year: 'numeric' 
                    }) : 'N/A'}
                  </td>
                  <td class="lot-number-cell">${item.lotNumber || 'N/A'}</td>
                  <td>${item.garmentType || 'N/A'}</td>
                  <td>${item.fabric || 'N/A'}</td>
                  <td>${item.style || 'N/A'}</td>
                  <td>
                    <span class="supervisor-cell">${item.packingSupervisor || 'N/A'}</span>
                  </td>
                  <td>${item.packingDate || 'N/A'}</td>
                  <td class="pcs-cell">${item.totalPcs ? item.totalPcs.toLocaleString() : '0'}</td>
                </tr>
              `).join('')}
              
              <tr class="total-row">
                <td colspan="7" style="text-align: right; padding-right: 20px;">GRAND TOTAL</td>
                <td class="pcs-cell">${totalPcs.toLocaleString()}</td>
              </tr>
            </tbody>
          </table>
        </div>
        
        <div class="signature-section">
          <div>
            <div class="signature-line"></div>
            <div style="text-align: center; font-size: 10px; color: #64748b;">Prepared By</div>
            <div style="text-align: center; font-size: 9px; color: #94a3b8;">System Administrator</div>
          </div>
          
          <div>
            <div class="signature-line"></div>
            <div style="text-align: center; font-size: 10px; color: #64748b;">Reviewed By</div>
            <div style="text-align: center; font-size: 9px; color: #94a3b8;">Production Manager</div>
          </div>
          
          <div>
            <div class="signature-line"></div>
            <div style="text-align: center; font-size: 10px; color: #64748b;">Approved By</div>
            <div style="text-align: center; font-size: 9px; color: #94a3b8;">Plant Director</div>
          </div>
        </div>
        
        <div class="page-info">
          Page 1 of 1 • Generated by Issue to Packing System • ${new Date().toLocaleString('en-US', { 
            year: 'numeric', 
            month: 'long', 
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit'
          })}
        </div>
      </div>
      
      <script>
        setTimeout(() => {
          window.print();
          setTimeout(() => {
            window.close();
          }, 500);
        }, 1000);
      </script>
    </body>
    </html>
  `;

  printWindow.document.write(content);
  printWindow.document.close();
};

  // Enhanced styles with modern design
  const styles = {
    container: {
      fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
      padding: '20px',
      maxWidth: '1900px',
      margin: '0 auto',
      minHeight: '100vh',
      background: 'linear-gradient(135deg, #ffffffff 0%, #ffffffff 100%)'
    },
    header: {
      background: 'white',
      padding: '30px',
      borderRadius: '20px',
      boxShadow: '0 10px 40px rgba(0,0,0,0.08)',
      marginBottom: '30px',
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      borderLeft: '5px solid #3498db',
      position: 'relative',
      overflow: 'hidden'
    },
    headerContent: {
      flex: 1
    },
    title: {
      color: '#2c3e50',
      fontSize: '32px',
      margin: '0 0 10px 0',
      fontWeight: '700',
      letterSpacing: '-0.5px'
    },
    subtitle: {
      color: '#7f8c8d',
      fontSize: '14px',
      margin: '0',
      fontWeight: '400'
    },
    headerDecoration: {
      position: 'absolute',
      right: '-50px',
      top: '-50px',
      width: '200px',
      height: '200px',
      background: 'linear-gradient(135deg, rgba(52, 152, 219, 0.1) 0%, rgba(41, 128, 185, 0.05) 100%)',
      borderRadius: '50%',
      zIndex: '0'
    },
    actionButtons: {
      display: 'flex',
      gap: '12px',
      flexWrap: 'wrap',
      zIndex: '1'
    },
    button: {
      padding: '12px 24px',
      border: 'none',
      borderRadius: '10px',
      cursor: 'pointer',
      fontWeight: '600',
      fontSize: '14px',
      transition: 'all 0.3s cubic-bezier(0.4, 0, 0.2, 1)',
      display: 'flex',
      alignItems: 'center',
      gap: '8px',
      letterSpacing: '0.3px',
      position: 'relative',
      overflow: 'hidden'
    },
    buttonHover: {
      transform: 'translateY(-2px)',
      boxShadow: '0 5px 15px rgba(0,0,0,0.1)'
    },
    primaryButton: {
      background: 'linear-gradient(135deg, #3498db 0%, #2980b9 100%)',
      color: 'white'
    },
    successButton: {
      background: 'linear-gradient(135deg, #27ae60 0%, #229954 100%)',
      color: 'white'
    },
    dangerButton: {
      background: 'linear-gradient(135deg, #e74c3c 0%, #c0392b 100%)',
      color: 'white'
    },
    warningButton: {
      background: 'linear-gradient(135deg, #f39c12 0%, #d68910 100%)',
      color: 'white'
    },
    infoButton: {
      background: 'linear-gradient(135deg, #9b59b6 0%, #8e44ad 100%)',
      color: 'white'
    },
    statsGrid: {
      display: 'grid',
      gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))',
      gap: '15px',
      marginBottom: '30px'
    },
    statCard: {
      background: 'white',
      padding: '30px',
      borderRadius: '15px',
      boxShadow: '0 8px 25px rgba(0,0,0,0.06)',
      textAlign: 'center',
      borderTop: '4px solid',
      transition: 'transform 0.3s cubic-bezier(0.4, 0, 0.2, 1)',
      position: 'relative',
      overflow: 'hidden'
    },
    recentLotsContainer: {
      background: 'white',
      padding: '30px',
      borderRadius: '15px',
      boxShadow: '0 8px 25px rgba(0,0,0,0.06)',
      marginBottom: '30px',
      borderTop: '4px solid #2ecc71'
    },
    recentLotsHeader: {
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      marginBottom: '25px',
      paddingBottom: '15px',
      borderBottom: '2px solid #ecf0f1'
    },
    recentLotsTitle: {
      fontSize: '24px',
      fontWeight: '700',
      color: '#2c3e50',
      display: 'flex',
      alignItems: 'center',
      gap: '12px'
    },
    recentLotsBadge: {
      background: 'linear-gradient(135deg, #2ecc71 0%, #27ae60 100%)',
      color: 'white',
      padding: '6px 16px',
      borderRadius: '20px',
      fontSize: '14px',
      fontWeight: '700'
    },
    recentLotsTable: {
      width: '100%',
      borderCollapse: 'collapse',
      border: '1px solid #e0e0e0',
      borderRadius: '10px',
      overflow: 'hidden'
    },
    recentTh: {
      padding: '16px 20px',
      textAlign: 'left',
      fontWeight: '600',
      fontSize: '13px',
      color: '#ffffffff',
      textTransform: 'uppercase',
      letterSpacing: '0.5px',
      borderRight: '1px solid #3498db',
      borderBottom: '2px solid #2980b9',
      background: 'linear-gradient(135deg, #3498db 0%, #2980b9 100%)'
    },
    recentTd: {
      padding: '16px 20px',
      borderBottom: '1px solid #e0e0e0',
      borderRight: '1px solid #e0e0e0',
      fontSize: '14px',
      verticalAlign: 'middle'
    },
    emptyRecentLots: {
      padding: '40px',
      textAlign: 'center',
      color: '#95a5a6',
      border: '2px dashed #ecf0f1',
      borderRadius: '10px'
    },
    filtersContainer: {
      background: 'white',
      padding: '30px',
      borderRadius: '15px',
      boxShadow: '0 8px 25px rgba(0,0,0,0.06)',
      marginBottom: '40px'
    },
    filterTitle: {
      fontSize: '20px',
      fontWeight: '600',
      color: '#2c3e50',
      marginBottom: '25px',
      display: 'flex',
      alignItems: 'center',
      gap: '10px'
    },
    filterGrid: {
      display: 'grid',
      gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))',
      gap: '25px',
      marginBottom: '30px'
    },
    filterGroup: {
      display: 'flex',
      flexDirection: 'column'
    },
    filterLabel: {
      marginBottom: '10px',
      fontWeight: '600',
      color: '#2c3e50',
      fontSize: '14px',
      display: 'flex',
      alignItems: 'center',
      gap: '6px'
    },
    filterSelect: {
      padding: '14px',
      border: '2px solid #e9ecef',
      borderRadius: '10px',
      fontSize: '14px',
      background: '#f8f9fa',
      transition: 'all 0.3s ease',
      appearance: 'none',
      backgroundImage: 'url("data:image/svg+xml;charset=UTF-8,%3csvg xmlns=\'http://www.w3.org/2000/svg\' viewBox=\'0 0 24 24\' fill=\'none\' stroke=\'%237f8c8d\' stroke-width=\'2\' stroke-linecap=\'round\' stroke-linejoin=\'round\'%3e%3cpolyline points=\'6 9 12 15 18 9\'%3e%3c/polyline%3e%3c/svg%3e")',
      backgroundRepeat: 'no-repeat',
      backgroundPosition: 'right 14px center',
      backgroundSize: '16px',
      paddingRight: '40px'
    },
    filterSelectFocus: {
      outline: 'none',
      borderColor: '#3498db',
      background: 'white',
      boxShadow: '0 0 0 3px rgba(52,152,219,0.1)'
    },
    filterActions: {
      display: 'flex',
      justifyContent: 'flex-end',
      gap: '15px',
      paddingTop: '25px',
      borderTop: '1px solid #e9ecef'
    },
    tableContainer: {
      background: 'white',
      borderRadius: '15px',
      boxShadow: '0 8px 25px rgba(0,0,0,0.06)',
      overflow: 'hidden',
      marginBottom: '40px',
      border: '1px solid #e0e0e0'
    },
    tableHeader: {
      padding: '25px',
      background: '#f8f9fa',
      borderBottom: '1px solid #e0e0e0'
    },
    tableTitle: {
      fontSize: '20px',
      fontWeight: '600',
      color: '#2c3e50',
      margin: '0'
    },
    table: {
      width: '100%',
      borderCollapse: 'collapse',
      border: '1px solid #e0e0e0'
    },
    th: {
      padding: '18px 20px',
      textAlign: 'left',
      fontWeight: '600',
      fontSize: '13px',
      color: '#ffffffff',
      textTransform: 'uppercase',
      letterSpacing: '0.5px',
      borderRight: '1px solid #004080',
      borderBottom: '2px solid #003366',
      background: '#004080ff',
      position: 'sticky',
      top: 0,
      zIndex: 10
    },
    td: {
      padding: '18px 20px',
      borderBottom: '1px solid #e0e0e0',
      borderRight: '1px solid #e0e0e0',
      fontSize: '14px',
      verticalAlign: 'middle'
    },
    recentLotRow: {
      background: 'linear-gradient(90deg, rgba(46, 204, 113, 0.1) 0%, rgba(39, 174, 96, 0.05) 100%)',
      borderLeft: '3px solid #2ecc71',
      position: 'relative'
    },
    recentBadge: {
      display: 'inline-block',
      padding: '3px 10px',
      borderRadius: '20px',
      fontSize: '11px',
      fontWeight: '700',
      background: 'linear-gradient(135deg, #2ecc71 0%, #27ae60 100%)',
      color: 'white',
      marginLeft: '10px',
      textTransform: 'uppercase'
    },
    tableFooter: {
      padding: '20px',
      background: '#f8f9fa',
      borderTop: '1px solid #e0e0e0',
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center'
    },
    loadingOverlay: {
      position: 'fixed',
      top: 0,
      left: 0,
      right: 0,
      bottom: 0,
      background: 'rgba(255,255,255,0.9)',
      display: 'flex',
      flexDirection: 'column',
      justifyContent: 'center',
      alignItems: 'center',
      zIndex: 1000
    },
    loadingSpinner: {
      border: '4px solid #f3f3f3',
      borderTop: '4px solid #3498db',
      borderRadius: '50%',
      width: '50px',
      height: '50px',
      animation: 'spin 1s linear infinite',
      marginBottom: '20px'
    },
    errorAlert: {
      background: 'linear-gradient(135deg, #ffcccc 0%, #ff9999 100%)',
      padding: '30px',
      borderRadius: '15px',
      textAlign: 'center',
      boxShadow: '0 5px 15px rgba(0,0,0,0.1)',
      margin: '40px 0'
    },
    emptyState: {
      padding: '60px 30px',
      textAlign: 'center',
      color: '#95a5a6'
    }
  };

  return (
    <div style={styles.container}>
      {/* Loading Overlay */}
      {loading && (
        <div style={styles.loadingOverlay}>
          <div style={styles.loadingSpinner}></div>
          <div style={{ fontSize: '18px', color: '#2c3e50', fontWeight: '600' }}>
            Loading Issue to Packing Data...
          </div>
          <div style={{ fontSize: '14px', color: '#7f8c8d', marginTop: '10px' }}>
            Fetching data from Google Sheets
          </div>
        </div>
      )}

      {/* Header */}
      <div style={styles.header}>
        <div style={styles.headerContent}>
          <h1 style={styles.title}>📦 Issue to Packing Dashboard</h1>
          <p style={styles.subtitle}>
            Monitor and manage all packing activities • Real-time data from Google Sheets
          </p>
        </div>
        <div style={styles.headerDecoration}></div>
        <div style={styles.actionButtons}>
          <button
            style={{ ...styles.button, ...styles.infoButton }}
            onClick={() => window.history.back()}
            onMouseEnter={(e) => Object.assign(e.target.style, styles.buttonHover)}
            onMouseLeave={(e) => {
              e.target.style.transform = 'translateY(0)';
              e.target.style.boxShadow = 'none';
            }}
          >
            ← Back
          </button>
          <button
            style={{ ...styles.button, ...styles.primaryButton }}
            onClick={fetchData}
            disabled={loading}
            onMouseEnter={(e) => Object.assign(e.target.style, styles.buttonHover)}
            onMouseLeave={(e) => {
              e.target.style.transform = 'translateY(0)';
              e.target.style.boxShadow = 'none';
            }}
          >
            🔄 Refresh
          </button>
          <button
            style={{ ...styles.button, ...styles.successButton }}
            onClick={exportToExcel}
            disabled={loading || filteredData.length === 0}
            onMouseEnter={(e) => Object.assign(e.target.style, styles.buttonHover)}
            onMouseLeave={(e) => {
              e.target.style.transform = 'translateY(0)';
              e.target.style.boxShadow = 'none';
            }}
          >
            📊 Export Excel
          </button>
          <button
            style={{ ...styles.button, ...styles.dangerButton }}
            onClick={exportToPDF}
            disabled={loading || filteredData.length === 0}
            onMouseEnter={(e) => Object.assign(e.target.style, styles.buttonHover)}
            onMouseLeave={(e) => {
              e.target.style.transform = 'translateY(0)';
              e.target.style.boxShadow = 'none';
            }}
          >
            📄 Export PDF
          </button>
        </div>
      </div>

      {/* Recently Issued Lots Section */}
      <div style={styles.recentLotsContainer}>
        <div style={styles.recentLotsHeader}>
          <div style={styles.recentLotsTitle}>
            <span>🚀 Recently Issued Lots (Last 24 Hours)</span>
            <span style={styles.recentLotsBadge}>
              {recentLots.length} Lots
            </span>
          </div>
          <div style={{ fontSize: '14px', color: '#7f8c8d' }}>
            Total Pieces: {recentLots.reduce((sum, item) => sum + item.totalPcs, 0).toLocaleString()}
          </div>
        </div>

        {recentLots.length === 0 ? (
          <div style={styles.emptyRecentLots}>
            <div style={{ fontSize: '64px', marginBottom: '20px', opacity: '0.3' }}>⏰</div>
            <div style={{ fontSize: '18px', fontWeight: '600', marginBottom: '10px' }}>
              No recently issued lots in the last 24 hours
            </div>
            <div style={{ fontSize: '14px', color: '#95a5a6' }}>
              New lots issued within 24 hours will appear here
            </div>
          </div>
        ) : (
          <div style={{ overflowX: 'auto' }}>
            <table style={styles.recentLotsTable}>
              <thead>
                <tr>
                  <th style={styles.recentTh}>Timestamp</th>
                  <th style={styles.recentTh}>Lot Number</th>
                  <th style={styles.recentTh}>Garment Type</th>
                  <th style={styles.recentTh}>Fabric</th>
                  <th style={styles.recentTh}>Style</th>
                  <th style={styles.recentTh}>Packing Supervisor</th>
                  <th style={styles.recentTh}>Packing Date</th>
                  <th style={{...styles.recentTh, borderRight: 'none'}}>Total Pcs</th>
                </tr>
              </thead>
              <tbody>
                {recentLots.map((item, index) => (
                  <tr
                    key={`recent-${index}`}
                    style={{
                      backgroundColor: index % 2 === 0 ? '#f8fff9' : '#f0f9f2',
                      transition: 'background-color 0.2s ease'
                    }}
                    onMouseEnter={(e) => e.currentTarget.style.backgroundColor = '#e8f8ee'}
                    onMouseLeave={(e) => {
                      e.currentTarget.style.backgroundColor = index % 2 === 0 ? '#f8fff9' : '#f0f9f2';
                    }}
                  >
                    <td style={styles.recentTd}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <div style={{
                          width: '8px',
                          height: '8px',
                          background: '#2ecc71',
                          borderRadius: '50%',
                          animation: 'pulse 1.5s infinite'
                        }}></div>
                        {item.timestamp}
                      </div>
                    </td>
                    <td style={styles.recentTd}>
                      <strong style={{ color: '#2c3e50' }}>{item.lotNumber}</strong>
                    </td>
                    <td style={styles.recentTd}>{item.garmentType}</td>
                    <td style={styles.recentTd}>{item.fabric}</td>
                    <td style={styles.recentTd}>{item.style}</td>
                    <td style={styles.recentTd}>
                      <span style={{
                        ...styles.button,
                        padding: '6px 15px',
                        background: 'linear-gradient(135deg, #e8f6ef 0%, #d1f2eb 100%)',
                        color: '#27ae60',
                        border: '1px solid #a3e4d7',
                        fontSize: '12px',
                        fontWeight: '600'
                      }}>
                        {item.packingSupervisor}
                      </span>
                    </td>
                    <td style={styles.recentTd}>{item.packingDate}</td>
                    <td style={{...styles.recentTd, borderRight: 'none'}}>
                      <span style={{
                        ...styles.button,
                        padding: '6px 15px',
                        background: 'linear-gradient(135deg, #d5f4e6 0%, #c8f7dc 100%)',
                        color: '#229954',
                        border: '1px solid #82e5aa',
                        fontSize: '12px',
                        fontWeight: '700'
                      }}>
                        {item.totalPcs.toLocaleString()}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* Filters Section */}
      <div style={styles.filtersContainer}>
        <div style={styles.filterTitle}>
          <span>🔍 Filter Data</span>
          <div style={{ fontSize: '14px', color: '#7f8c8d', fontWeight: 'normal', marginLeft: '10px' }}>
            Use dropdowns to filter records
          </div>
        </div>
        
        <div style={styles.filterGrid}>
          <div style={styles.filterGroup}>
            <label style={styles.filterLabel}>
              <span>📦 Lot Number</span>
            </label>
            <select
              name="lotNumber"
              value={filters.lotNumber}
              onChange={handleFilterChange}
              style={styles.filterSelect}
              onFocus={(e) => Object.assign(e.target.style, styles.filterSelectFocus)}
              onBlur={(e) => {
                e.target.style.borderColor = '#e9ecef';
                e.target.style.background = '#f8f9fa';
                e.target.style.boxShadow = 'none';
              }}
            >
              <option value="">All Lot Numbers</option>
              {dropdownOptions.lotNumbers.map((lot, index) => (
                <option key={index} value={lot}>{lot}</option>
              ))}
            </select>
          </div>
          
          <div style={styles.filterGroup}>
            <label style={styles.filterLabel}>
              <span>👕 Garment Type</span>
            </label>
            <select
              name="garmentType"
              value={filters.garmentType}
              onChange={handleFilterChange}
              style={styles.filterSelect}
              onFocus={(e) => Object.assign(e.target.style, styles.filterSelectFocus)}
              onBlur={(e) => {
                e.target.style.borderColor = '#e9ecef';
                e.target.style.background = '#f8f9fa';
                e.target.style.boxShadow = 'none';
              }}
            >
              <option value="">All Garment Types</option>
              {dropdownOptions.garmentTypes.map((type, index) => (
                <option key={index} value={type}>{type}</option>
              ))}
            </select>
          </div>
          
          <div style={styles.filterGroup}>
            <label style={styles.filterLabel}>
              <span>🧵 Fabric</span>
            </label>
            <select
              name="fabric"
              value={filters.fabric}
              onChange={handleFilterChange}
              style={styles.filterSelect}
              onFocus={(e) => Object.assign(e.target.style, styles.filterSelectFocus)}
              onBlur={(e) => {
                e.target.style.borderColor = '#e9ecef';
                e.target.style.background = '#f8f9fa';
                e.target.style.boxShadow = 'none';
              }}
            >
              <option value="">All Fabrics</option>
              {dropdownOptions.fabrics.map((fabric, index) => (
                <option key={index} value={fabric}>{fabric}</option>
              ))}
            </select>
          </div>
          
          <div style={styles.filterGroup}>
            <label style={styles.filterLabel}>
              <span>👤 Packing Supervisor</span>
            </label>
            <select
              name="packingSupervisor"
              value={filters.packingSupervisor}
              onChange={handleFilterChange}
              style={styles.filterSelect}
              onFocus={(e) => Object.assign(e.target.style, styles.filterSelectFocus)}
              onBlur={(e) => {
                e.target.style.borderColor = '#e9ecef';
                e.target.style.background = '#f8f9fa';
                e.target.style.boxShadow = 'none';
              }}
            >
              <option value="">All Supervisors</option>
              {dropdownOptions.packingSupervisors.map((supervisor, index) => (
                <option key={index} value={supervisor}>{supervisor}</option>
              ))}
            </select>
          </div>
          
          <div style={styles.filterGroup}>
            <label style={styles.filterLabel}>
              <span>📅 Start Date</span>
            </label>
            <input
              type="date"
              name="startDate"
              value={filters.startDate}
              onChange={handleFilterChange}
              style={{ ...styles.filterSelect, backgroundImage: 'none' }}
              onFocus={(e) => Object.assign(e.target.style, styles.filterSelectFocus)}
              onBlur={(e) => {
                e.target.style.borderColor = '#e9ecef';
                e.target.style.background = '#f8f9fa';
                e.target.style.boxShadow = 'none';
              }}
            />
          </div>
          
          <div style={styles.filterGroup}>
            <label style={styles.filterLabel}>
              <span>📅 End Date</span>
            </label>
            <input
              type="date"
              name="endDate"
              value={filters.endDate}
              onChange={handleFilterChange}
              style={{ ...styles.filterSelect, backgroundImage: 'none' }}
              onFocus={(e) => Object.assign(e.target.style, styles.filterSelectFocus)}
              onBlur={(e) => {
                e.target.style.borderColor = '#e9ecef';
                e.target.style.background = '#f8f9fa';
                e.target.style.boxShadow = 'none';
              }}
            />
          </div>
        </div>
        
        <div style={styles.filterActions}>
          <button
            style={{ ...styles.button, ...styles.warningButton }}
            onClick={resetFilters}
            onMouseEnter={(e) => Object.assign(e.target.style, styles.buttonHover)}
            onMouseLeave={(e) => {
              e.target.style.transform = 'translateY(0)';
              e.target.style.boxShadow = 'none';
            }}
          >
            🗑️ Clear All Filters
          </button>
        </div>
      </div>

      {/* Error Display */}
      {error && (
        <div style={styles.errorAlert}>
          <div style={{ fontSize: '24px', marginBottom: '15px' }}>⚠️</div>
          <div style={{ fontSize: '18px', fontWeight: '600', marginBottom: '10px', color: '#c0392b' }}>
            {error}
          </div>
          <button
            style={{ ...styles.button, ...styles.primaryButton, marginTop: '15px' }}
            onClick={fetchData}
          >
            Try Again
          </button>
        </div>
      )}

      {/* Main Data Table */}
      <div style={styles.tableContainer}>
        <div style={styles.tableHeader}>
          <h3 style={styles.tableTitle}>📋 All Issue to Packing Records</h3>
          <div style={{ fontSize: '14px', color: '#7f8c8d', marginTop: '5px' }}>
            Showing {filteredData.length} of {data.length} records • {stats.totalPcs.toLocaleString()} total pieces
          </div>
        </div>
        
        {filteredData.length === 0 && !loading && !error ? (
          <div style={styles.emptyState}>
            <div style={{ fontSize: '64px', marginBottom: '20px', opacity: '0.3' }}>📭</div>
            <div style={{ fontSize: '18px', fontWeight: '600', marginBottom: '10px' }}>
              No records found matching your filters
            </div>
            <div style={{ fontSize: '14px', color: '#95a5a6', marginBottom: '20px' }}>
              Try adjusting your filters or check the data source
            </div>
            <button
              style={{ ...styles.button, ...styles.infoButton }}
              onClick={resetFilters}
            >
              Reset All Filters
            </button>
          </div>
        ) : (
          <>
            <div style={{ overflowX: 'auto' }}>
              <table style={styles.table}>
                <thead>
                  <tr>
                    <th style={styles.th}>Timestamp</th>
                    <th style={styles.th}>Lot Number</th>
                    <th style={styles.th}>Garment Type</th>
                    <th style={styles.th}>Fabric</th>
                    <th style={styles.th}>Style</th>
                    <th style={styles.th}>Packing Supervisor</th>
                    <th style={styles.th}>Packing Date</th>
                    <th style={{...styles.th, borderRight: 'none'}}>Total Pcs</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredData.map((item, index) => (
                    <tr
                      key={index}
                      style={item.isRecent ? {
                        ...styles.recentLotRow,
                        backgroundColor: index % 2 === 0 ? '#f8fff9' : '#f0f9f2'
                      } : {
                        backgroundColor: index % 2 === 0 ? '#f8f9fa' : 'white'
                      }}
                      onMouseEnter={(e) => e.currentTarget.style.backgroundColor = '#e8f4fc'}
                      onMouseLeave={(e) => {
                        if (item.isRecent) {
                          e.currentTarget.style.backgroundColor = index % 2 === 0 ? '#f8fff9' : '#f0f9f2';
                        } else {
                          e.currentTarget.style.backgroundColor = index % 2 === 0 ? '#f8f9fa' : 'white';
                        }
                      }}
                    >
                      <td style={styles.td}>{item.timestamp}</td>
                      <td style={styles.td}>
                        <strong style={{ color: '#2c3e50' }}>{item.lotNumber}</strong>
                        {item.isRecent && <span style={styles.recentBadge}>RECENT</span>}
                      </td>
                      <td style={styles.td}>{item.garmentType}</td>
                      <td style={styles.td}>{item.fabric}</td>
                      <td style={styles.td}>{item.style}</td>
                      <td style={styles.td}>
                        <span style={{
                          ...styles.button,
                          padding: '6px 15px',
                          background: 'linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%)',
                          color: '#1565c0',
                          border: '1px solid #90caf9',
                          fontSize: '12px',
                          fontWeight: '600'
                        }}>
                          {item.packingSupervisor}
                        </span>
                      </td>
                      <td style={styles.td}>{item.packingDate}</td>
                      <td style={{...styles.td, borderRight: 'none'}}>
                        <span style={{
                          ...styles.button,
                          padding: '6px 15px',
                          background: item.totalPcs > 500 
                            ? 'linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%)'
                            : item.totalPcs > 300
                            ? 'linear-gradient(135deg, #fff3e0 0%, #ffe0b2 100%)'
                            : 'linear-gradient(135deg, #ffebee 0%, #ffcdd2 100%)',
                          color: item.totalPcs > 500 ? '#2e7d32' : 
                                item.totalPcs > 300 ? '#f57c00' : '#c62828',
                          border: item.totalPcs > 500 ? '1px solid #a5d6a7' : 
                                 item.totalPcs > 300 ? '1px solid #ffcc80' : '1px solid #ef9a9a',
                          fontSize: '12px',
                          fontWeight: '700'
                        }}>
                          {item.totalPcs.toLocaleString()}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            
            <div style={styles.tableFooter}>
              <div style={{ fontSize: '14px', color: '#6c757d' }}>
                Showing {filteredData.length} records • {stats.totalPcs.toLocaleString()} total pieces
              </div>
              <div style={{ fontSize: '14px', color: '#2c3e50', fontWeight: '600' }}>
                Recent Lots: {recentLots.length}
              </div>
            </div>
          </>
        )}
      </div>

      <style>
        {`
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          
          @keyframes pulse {
            0% { transform: scale(0.95); opacity: 0.7; }
            50% { transform: scale(1.05); opacity: 1; }
            100% { transform: scale(0.95); opacity: 0.7; }
          }
        `}
      </style>
    </div>
  );
};

export default IssueToPacking;