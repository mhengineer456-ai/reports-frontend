import React, { useState, useEffect, useRef } from 'react';
import axios from 'axios';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import './ShortSummaryReport.css'; // Import the external CSS

const ShortSummaryReport = () => {
  const [sheetData, setSheetData] = useState([]);
  const [jobOrderData, setJobOrderData] = useState([]);
  const [cuttingData, setCuttingData] = useState([]);
  const [indexData, setIndexData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [headers, setHeaders] = useState([]);
  const [jobOrderHeaders, setJobOrderHeaders] = useState([]);
  const [activeDepartment, setActiveDepartment] = useState('all');
  const [departmentSummary, setDepartmentSummary] = useState({});
  const [filteredDepartmentSummary, setFilteredDepartmentSummary] = useState({});
  const [colourPendingLots, setColourPendingLots] = useState([]);
  const [colourPendingDetails, setColourPendingDetails] = useState({});

  // New state for Pending Issue after Emb/Print department
  const [pendingIssueLots, setPendingIssueLots] = useState([]);
  const [pendingIssueDetails, setPendingIssueDetails] = useState({});

  // Password protection state
  const [showPasswordModal, setShowPasswordModal] = useState(false);
  const [passwordInput, setPasswordInput] = useState('');
  const [passwordError, setPasswordError] = useState('');
  const [downloadMode, setDownloadMode] = useState('client'); // 'client' or 'md'
  const [pendingPdfData, setPendingPdfData] = useState(null);

  // Supervisor summary state
  const [supervisorSummary, setSupervisorSummary] = useState({});
  const [filteredSupervisorSummary, setFilteredSupervisorSummary] = useState({});
  const [showSupervisorDetails, setShowSupervisorDetails] = useState(false);

  // Filter states
  const [filters, setFilters] = useState({
    party: 'all',
    brand: 'all',
    item: 'all',
    season: 'all'
  });

  // Filter options
  const [filterOptions, setFilterOptions] = useState({
    parties: [],
    brands: [],
    items: [],
    seasons: []
  });

  // Ref for PDF export
  const summaryRef = useRef(null);

  // Password configuration - change these as needed
  const MD_PASSWORD = 'md123'; // Password for MD access (full data)
  const CLIENT_PASSWORD = 'client123'; // Password for client access (limited data)

  // Configuration - replace with your actual values
  const API_KEY = 'AIzaSyAomDFBkOySlIxKWSKGHe6ATv9gvaBr7uk';
  const SPREADSHEET_ID = '1Hj3JeJEKB43aYYWv8gk2UhdU6BWuEQfCg5pBlTdBMNA';
  const JOBORDER_SHEET_ID = '1fKSwGBIpzWEFk566WRQ4bzQ0anJlmasoY8TwrTLQHXI';
  const INDEX_SHEET_NAME = 'Index';
  const JOBORDER_SHEET_NAME = 'JobOrder';
  const CUTTING_SHEET_NAME = 'Cutting';
  const RANGE = 'A:AL';
  const CUTTING_BIG_RANGE = `${CUTTING_SHEET_NAME}!A1:ZZ200000`;

  // Define departments with their search criteria
  const departments = [
    {
      id: 'all',
      name: 'All Data',
      color: '#667eea',
      icon: '📊',
      searchTerms: [],
      description: 'Complete dataset overview from both sheets'
    },
    {
      id: 'cutting_done',
      name: 'Cutting Done',
      subtitle: 'All Lots from Index Sheet',
      color: '#FFA500',
      icon: '✂️',
      searchTerms: ['cutting done'],
      description: 'All lots present in Index sheet (All cutting records)'
    },
    {
      id: 'fabric_pending',
      name: 'Fabric Issue Pending',
      subtitle: 'Missing in Index Sheet',
      color: '#FF5252',
      icon: '🧵',
      searchTerms: ['fabric', 'issue', 'pending'],
      conditions: {
        missingInIndex: true
      },
      description: 'Lots from JobOrder sheet that are missing in Index sheet (Fabric issue pending)'
    },
    {
      id: 'colour_pending',
      name: 'Colour Pending on Cutting',
      subtitle: 'Pending Colour/Shade Lots',
      color: '#E91E63',
      icon: '🎨',
      searchTerms: ['colour pending', 'color pending'],
      conditions: {
        type: 'colour_pending'
      },
      description: 'Lots with pending colour/shade completion from cutting data'
    },
    {
      id: 'direct_stitching',
      name: 'Total Direct Stitching Lots',
      subtitle: 'Yes in Direct Stitching Column',
      color: '#10b981',
      icon: '🧷',
      searchTerms: ['direct stitching'],
      conditions: {
        directStitching: 'Yes'
      },
      description: 'Lots marked as Direct Stitching in Index sheet'
    },
    {
      id: 'direct_pending_issue',
      name: 'Direct Lots Pending Issue to Stitching',
      subtitle: 'Direct Stitching = Yes, No Issue Date/Supervisor',
      color: '#FF8C00',
      icon: '⏱️',
      searchTerms: ['direct pending', 'direct stitching pending issue'],
      conditions: {
        directStitching: 'Yes',
        dateOfIssueEmpty: true,
        supervisorEmpty: true
      },
      description: 'Lots marked as Direct Stitching in Index sheet but Pending Issue to Stitching After Cutting (empty Date of Issue and Supervisor)'
    },
    {
      id: 'pending_stitching',
      name: 'Pending Issue to Stitching After Cutting',
      subtitle: 'To Stitching',
      color: '#FF5722',
      icon: '⏳',
      searchTerms: ['pending stitching', 'pending to stitch'],
      conditions: {
        supervisorEmpty: true
      },
      description: 'Lots Pending Issue to Stitching After Cutting'
    },
    {
      id: 'pending_after_emb_print',
      name: 'Pending Issue After Emb/Print Done',
      subtitle: 'Emb/Print Completed Only',
      color: '#4CAF50',
      icon: '✅',
      searchTerms: ['pending after emb', 'pending after print', 'emb completed'],
      conditions: {
        type: 'pending_after_emb_print',
        supervisorEmpty: true,
        embPrintCompletedOnly: true // Only Emb/Print completed, no Direct Stitching
      },
      description: 'Lots that have completed Embroidery/Printing (excluding Direct Stitching) and are ready for supervisor assignment'
    },
    {
      id: 'embroidery',
      name: 'Embroidery WIP Record',
      subtitle: 'Pending Lots',
      color: '#9C27B0',
      icon: '🧵',
      searchTerms: ['embroidery', 'emb'],
      conditions: {
        type: 'embroidery',
        pendingOnly: true
      },
      description: 'Pending embroidery challans'
    },
    {
      id: 'printing',
      name: 'Printing WIP Record',
      subtitle: 'Pending Lots',
      color: '#FF9800',
      icon: '🖨️',
      searchTerms: ['printing', 'print'],
      conditions: {
        type: 'printing',
        pendingOnly: true
      },
      description: 'Pending printing challans'
    },
    {
      id: 'stitching',
      name: 'Stitching WIP Lots',
      subtitle: 'Issued Lots',
      color: '#2196F3',
      icon: '🪡',
      searchTerms: ['stitching', 'stitch'],
      conditions: {
        issued: true,
        notCompleted: true
      },
      description: 'Issued but not completed lots'
    },



    {
      id: 'packing',
      name: 'Packing Report',
      subtitle: 'Packing Status',
      color: '#795548',
      icon: '📦',
      searchTerms: ['packing', 'pack'],
      description: 'Packing department report'
    }
  ];

  // Define the specific columns to display in table
  const displayColumns = [
    { key: 'srNo', name: 'Sr. No', searchTerms: ['sr', 'no', 'sr.', 'serial'], width: '80px' },
    { key: 'lotNumber', name: 'Lot Number', searchTerms: ['lot number', 'lot no'], width: '120px' },
    { key: 'fabric', name: 'Fabric', searchTerms: ['fabric'], width: '120px' },
    { key: 'garmentType', name: 'Garment Type', searchTerms: ['garment type'], width: '130px' },
    { key: 'style', name: 'Style', searchTerms: ['style'], width: '100px' },
    { key: 'brand', name: 'Brand', searchTerms: ['brand'], width: '100px' },
    { key: 'season', name: 'Season', searchTerms: ['season'], width: '100px' },
    { key: 'directStitching', name: 'Direct Stitching', searchTerms: ['direct stitching'], width: '120px' },
    { key: 'cuttingQty', name: 'Cutting Qty', searchTerms: ['cutting qty'], width: '110px' },
    { key: 'completeStatus', name: 'Complete Status', searchTerms: ['complete status', 'completed status', 'status'], width: '150px' }
  ];

  // Columns for JobOrder sheet display
  const jobOrderDisplayColumns = [
    { key: 'jobOrderNo', name: 'Job Order No', searchTerms: ['job order no'], width: '100px' },
    { key: 'date', name: 'Date', searchTerms: ['date'], width: '100px' },
    { key: 'fabric', name: 'Fabric', searchTerms: ['fabric'], width: '100px' },
    { key: 'brand', name: 'Brand', searchTerms: ['brand'], width: '100px' },
    { key: 'shade', name: 'Shade', searchTerms: ['shade'], width: '80px' },
    { key: 'size', name: 'Size', searchTerms: ['size'], width: '80px' },
    { key: 'quantity', name: 'Quantity', searchTerms: ['quantity'], width: '100px' },
    { key: 'unit', name: 'Unit', searchTerms: ['unit'], width: '80px' },
    { key: 'partyName', name: 'Party Name', searchTerms: ['party name'], width: '120px' },
    { key: 'garmentType', name: 'Garment Type', searchTerms: ['garment type'], width: '120px' },
    { key: 'section', name: 'Section', searchTerms: ['section'], width: '100px' },
    { key: 'season', name: 'Season', searchTerms: ['season'], width: '100px' },
    { key: 'emb', name: 'Emb', searchTerms: ['emb'], width: '80px' },
    { key: 'embDetails', name: 'Emb Details', searchTerms: ['emb details'], width: '120px' },
    { key: 'printing', name: 'Printing', searchTerms: ['printing'], width: '100px' },
    { key: 'printingDetails', name: 'Printing Details', searchTerms: ['printing details'], width: '120px' },
    { key: 'pattern', name: 'Pattern', searchTerms: ['pattern'], width: '100px' },
    { key: 'style', name: 'Style', searchTerms: ['style'], width: '100px' },
    { key: 'remarks', name: 'Remarks', searchTerms: ['remarks'], width: '120px' },
    { key: 'directStitching', name: 'Direct Stitching', searchTerms: ['direct stitching'], width: '120px' },
    { key: 'submittedBy', name: 'Submitted By', searchTerms: ['submitted by'], width: '120px' },
    { key: 'imageUrl', name: 'Image URL', searchTerms: ['image url'], width: '150px' },
    { key: 'lotNumber', name: 'Lot Number', searchTerms: ['lot number'], width: '120px' },
    { key: 'status', name: 'Status', searchTerms: ['status'], width: '100px' }
  ];

  // State to store column indices for display columns
  const [displayColumnIndices, setDisplayColumnIndices] = useState({});
  const [jobOrderColumnIndices, setJobOrderColumnIndices] = useState({});

  useEffect(() => {
    fetchData();
  }, []);

  // Update filtered summary when filters change
  useEffect(() => {
    if (sheetData.length > 0 || jobOrderData.length > 0 || cuttingData.length > 0) {
      calculateFilteredDepartmentSummaries();
    }
  }, [filters, sheetData, jobOrderData, cuttingData, indexData, displayColumnIndices, jobOrderColumnIndices]);

  // Function to find column index
  const findColumnIndex = (headersArray, columnNames) => {
    if (!headersArray || !Array.isArray(headersArray)) return -1;

    for (let name of columnNames) {
      const index = headersArray.findIndex(h =>
        h && h.toString().toLowerCase().trim().includes(name.toLowerCase())
      );
      if (index !== -1) return index;
    }
    return -1;
  };

  // Function to check if a cell is empty
  const isEmptyCell = (cellValue) => {
    if (cellValue === undefined || cellValue === null) return true;

    const strValue = cellValue.toString().trim();
    return strValue === '' || strValue === '-' || strValue === 'null' || strValue === 'undefined' || strValue === 'N/A';
  };

  // Function to normalize supervisor name
  const normalizeSupervisorName = (name) => {
    if (!name) return '';

    // Convert to string, trim, and convert to proper case (first letter capital, rest lowercase)
    const normalized = name.toString()
      .trim()
      .toLowerCase()
      .split(' ')
      .map(word => word.charAt(0).toUpperCase() + word.slice(1))
      .join(' ');

    return normalized;
  };

  // Function to check if lot is completed based on Completed Status column
  const isLotCompleted = (completeStatus) => {
    if (!completeStatus || isEmptyCell(completeStatus)) return false;

    const statusStr = completeStatus.toString().toLowerCase().trim();
    // If it contains any value that's not empty, consider it completed
    return statusStr.length > 0 && statusStr !== 'pending' && statusStr !== 'in progress';
  };

  // Function to check if lot is cancelled in JobOrder
  const isLotCancelled = (row) => {
    const statusIndex = findColumnIndex(jobOrderHeaders, ['status']);
    if (statusIndex === -1) return false;

    const statusValue = row[statusIndex];
    if (!statusValue) return false;

    const statusStr = statusValue.toString().toLowerCase().trim();
    return statusStr.includes('cancel') || statusStr.includes('cancelled');
  };

  // Function to parse complete status JSON
  const parseCompleteStatus = (completeStatus) => {
    if (!completeStatus || isEmptyCell(completeStatus)) return null;

    try {
      let cleaned = completeStatus.trim();

      if (cleaned.startsWith('"') && cleaned.endsWith('"')) {
        cleaned = cleaned.slice(1, -1);
      }

      cleaned = cleaned.replace(/\\"/g, '"');
      const parsed = JSON.parse(cleaned);

      if (Array.isArray(parsed) && parsed.length > 0) {
        return parsed[0];
      } else if (typeof parsed === 'object' && parsed !== null) {
        return parsed;
      }

      return null;
    } catch (error) {
      if (completeStatus.toLowerCase().includes('complete')) {
        return { status: 'Complete', timestamp: '' };
      }

      return null;
    }
  };

  // Function to parse challan history JSON
  const parseChallanHistory = (challanHistory) => {
    if (!challanHistory || isEmptyCell(challanHistory)) return [];

    try {
      let cleaned = challanHistory.trim();

      if (cleaned.startsWith('"') && cleaned.endsWith('"')) {
        cleaned = cleaned.slice(1, -1);
      }

      cleaned = cleaned.replace(/\\"/g, '"');
      const parsed = JSON.parse(cleaned);

      if (Array.isArray(parsed)) {
        return parsed;
      } else if (typeof parsed === 'object' && parsed !== null) {
        return [parsed];
      }

      return [];
    } catch (error) {
      try {
        const arrayMatch = challanHistory.match(/\[.*\]/);
        if (arrayMatch) {
          return JSON.parse(arrayMatch[0]);
        }

        const objectMatch = challanHistory.match(/\{[^}]+\}/);
        if (objectMatch) {
          return [JSON.parse(objectMatch[0])];
        }
      } catch (e) {
        console.error('Alternative parsing failed:', e);
      }

      return [];
    }
  };

  // Function to check if lot is pending in embroidery
  const isEmbroideryLotPending = (challanHistory) => {
    const challans = parseChallanHistory(challanHistory);

    for (const challan of challans) {
      const challanNumber = String(challan.number || '').toUpperCase();

      if (challanNumber.startsWith('CH-EMB')) {
        const isCompleted = challan.embCompleted === true ||
          challan.completed === true ||
          challan.emb_completed === true ||
          challan.isCompleted === true;

        const isPending = challan.embCompleted === false ||
          challan.completed === false ||
          challan.emb_completed === false ||
          challan.isCompleted === false;

        if (isPending) {
          return true;
        }

        if (challan.embCompleted === undefined && challanNumber.startsWith('CH-EMB')) {
          const hasUpdatedAt = challan.embUpdatedAt || challan.updatedAt;
          const hasItems = challan.items && Array.isArray(challan.items) && challan.items.length > 0;

          if (hasItems) {
            const allItemsCompleted = challan.items.every(item =>
              item.completed === true || item.done === true
            );
            if (!allItemsCompleted) {
              return true;
            }
          } else if (!hasUpdatedAt) {
            return true;
          }
        }
      }
    }

    return false;
  };

  // Function to check if lot has completed embroidery (all challans done)
  const hasCompletedEmbroidery = (challanHistory) => {
    const challans = parseChallanHistory(challanHistory);

    // If no challans, return false
    if (challans.length === 0) return false;

    // Get only embroidery challans
    const embChallans = challans.filter(challan =>
      String(challan.number || '').toUpperCase().startsWith('CH-EMB')
    );

    // If no embroidery challans, return false
    if (embChallans.length === 0) return false;

    // Check if ALL embroidery challans are completed
    return embChallans.every(challan =>
      challan.embCompleted === true && challan.embUpdatedAt
    );
  };

  // Function to check if lot has completed printing (all challans done)
  const hasCompletedPrinting = (challanHistory) => {
    const challans = parseChallanHistory(challanHistory);

    // If no challans, return false
    if (challans.length === 0) return false;

    // Get only printing challans
    const printChallans = challans.filter(challan =>
      String(challan.number || '').toUpperCase().startsWith('CH-PRINT')
    );

    // If no printing challans, return false
    if (printChallans.length === 0) return false;

    // Check if ALL printing challans are completed
    return printChallans.every(challan =>
      challan.printCompleted === true || challan.embCompleted === true
    );
  };

  // Function to check if lot has completed embroidery OR printing (for pending after emb/print department)
  const hasCompletedEmbroideryOrPrinting = (challanHistory) => {
    const challans = parseChallanHistory(challanHistory);

    // If no challans, return false
    if (challans.length === 0) return false;

    // Check if there are any embroidery or printing challans
    const hasEmbOrPrintChallans = challans.some(challan => {
      const challanNumber = String(challan.number || '').toUpperCase();
      return challanNumber.startsWith('CH-EMB') || challanNumber.startsWith('CH-PRINT');
    });

    if (!hasEmbOrPrintChallans) return false;

    // Check if ALL embroidery challans are completed
    const embChallans = challans.filter(challan =>
      String(challan.number || '').toUpperCase().startsWith('CH-EMB')
    );
    const allEmbCompleted = embChallans.length === 0 || embChallans.every(challan =>
      challan.embCompleted === true
    );

    // Check if ALL printing challans are completed
    const printChallans = challans.filter(challan =>
      String(challan.number || '').toUpperCase().startsWith('CH-PRINT')
    );
    const allPrintCompleted = printChallans.length === 0 || printChallans.every(challan =>
      challan.printCompleted === true || challan.embCompleted === true
    );

    return allEmbCompleted && allPrintCompleted;
  };

  // Function to get last EMB/Print date
  const getLastEmbPrintDate = (challanHistory) => {
    const challans = parseChallanHistory(challanHistory);

    if (!challans || challans.length === 0) return null;

    try {
      const completedChallans = challans.filter(challan => {
        const challanNumber = String(challan.number || '').toUpperCase();
        if (challanNumber.startsWith('CH-EMB')) {
          return challan.embCompleted === true && challan.embUpdatedAt;
        } else if (challanNumber.startsWith('CH-PRINT')) {
          return challan.printCompleted === true && challan.printUpdatedAt;
        }
        return false;
      });

      if (completedChallans.length === 0) return null;

      const sortedChallans = [...completedChallans].sort((a, b) => {
        try {
          const dateA = new Date(a.embUpdatedAt || a.printUpdatedAt || a.updatedAt);
          const dateB = new Date(b.embUpdatedAt || b.printUpdatedAt || b.updatedAt);
          if (isNaN(dateA.getTime()) || isNaN(dateB.getTime())) return 0;
          return dateB - dateA;
        } catch (e) {
          return 0;
        }
      });

      const latestDate = sortedChallans[0]?.embUpdatedAt ||
        sortedChallans[0]?.printUpdatedAt ||
        sortedChallans[0]?.updatedAt;

      if (latestDate) {
        const date = new Date(latestDate);
        return isNaN(date.getTime()) ? null : date;
      }

      return null;
    } catch (e) {
      console.error('Error getting last EMB/Print date:', e);
      return null;
    }
  };

  // Function to check if lot is pending in printing
  const isPrintingLotPending = (challanHistory) => {
    const challans = parseChallanHistory(challanHistory);

    for (const challan of challans) {
      const challanNumber = String(challan.number || '').toUpperCase();

      if (challanNumber.startsWith('CH-PRINT')) {
        const isCompleted = challan.embCompleted === true ||
          challan.printCompleted === true ||
          challan.print_completed === true ||
          challan.completed === true ||
          challan.isCompleted === true;

        const isPending = challan.embCompleted === false ||
          challan.printCompleted === false ||
          challan.completed === false ||
          challan.isCompleted === false;

        if (isPending) {
          return true;
        }

        if ((challan.embCompleted === undefined && challan.printCompleted === undefined)
          && challanNumber.startsWith('CH-PRINT')) {
          const hasUpdatedAt = challan.printUpdatedAt || challan.embUpdatedAt || challan.updatedAt;
          const hasItems = challan.items && Array.isArray(challan.items) && challan.items.length > 0;

          if (hasItems) {
            const allItemsCompleted = challan.items.every(item =>
              item.completed === true || item.done === true
            );
            if (!allItemsCompleted) {
              return true;
            }
          } else if (!hasUpdatedAt) {
            return true;
          }
        }
      }
    }

    return false;
  };

  // Function to get pending embroidery details
  const getPendingEmbroideryDetails = (challanHistory) => {
    const challans = parseChallanHistory(challanHistory);
    const pendingLots = [];

    for (const challan of challans) {
      const challanNumber = String(challan.number || '').toUpperCase();

      if (challanNumber.startsWith('CH-EMB')) {
        const isCompleted = challan.embCompleted === true;
        const isPending = challan.embCompleted === false;

        if (isPending || challan.embCompleted === undefined) {
          const pendingItems = [];

          if (challan.items && Array.isArray(challan.items)) {
            challan.items.forEach(item => {
              if (item.completed !== true && item.done !== true) {
                pendingItems.push({
                  shade: item.shade || 'Unknown',
                  qty: item.qty || 0,
                  completed: item.completed || false
                });
              }
            });
          }

          pendingLots.push({
            lotNumber: challan.number || 'Unknown',
            date: challan.date || challan.challanDate || 'N/A',
            totalQty: challan.totalQty || challan.qty || 0,
            items: pendingItems.length > 0 ? pendingItems : challan.items || [],
            receivedDate: challan.receivedDate || challan.updatedAt || 'N/A',
            completeLot: challan.completeLot || false,
            embCompleted: challan.embCompleted,
            embUpdatedAt: challan.embUpdatedAt
          });
        }
      }
    }

    return pendingLots;
  };

  // Function to get completed embroidery challans
  const getCompletedEmbroideryChallans = (challanHistory) => {
    const challans = parseChallanHistory(challanHistory);

    return challans.filter(challan => {
      const challanNumber = String(challan.number || '').toUpperCase();
      return challanNumber.startsWith('CH-EMB') && challan.embCompleted === true;
    });
  };

  // Function to get completed printing challans
  const getCompletedPrintingChallans = (challanHistory) => {
    const challans = parseChallanHistory(challanHistory);

    return challans.filter(challan => {
      const challanNumber = String(challan.number || '').toUpperCase();
      return challanNumber.startsWith('CH-PRINT') && challan.printCompleted === true;
    });
  };

  // Function to get pending printing details
  const getPendingPrintingDetails = (challanHistory) => {
    const challans = parseChallanHistory(challanHistory);
    const pendingLots = [];

    for (const challan of challans) {
      const challanNumber = String(challan.number || '').toUpperCase();

      if (challanNumber.startsWith('CH-PRINT')) {
        const isCompleted = challan.embCompleted === true || challan.printCompleted === true;
        const isPending = challan.embCompleted === false || challan.printCompleted === false;

        if (isPending || (challan.embCompleted === undefined && challan.printCompleted === undefined)) {
          const pendingItems = [];

          if (challan.items && Array.isArray(challan.items)) {
            challan.items.forEach(item => {
              if (item.completed !== true && item.done !== true) {
                pendingItems.push({
                  shade: item.shade || 'Unknown',
                  qty: item.qty || 0,
                  completed: item.completed || false
                });
              }
            });
          }

          pendingLots.push({
            lotNumber: challan.number || 'Unknown',
            date: challan.date || challan.challanDate || 'N/A',
            totalQty: challan.totalQty || challan.qty || 0,
            items: pendingItems.length > 0 ? pendingItems : challan.items || [],
            receivedDate: challan.receivedDate || challan.updatedAt || 'N/A',
            completeLot: challan.completeLot || false,
            embCompleted: challan.embCompleted,
            printCompleted: challan.printCompleted,
            embUpdatedAt: challan.embUpdatedAt,
            printUpdatedAt: challan.printUpdatedAt
          });
        }
      }
    }

    return pendingLots;
  };

  // ===== CUTTING SHEET HELPER FUNCTIONS =====

  // Function to normalize keys for comparison
  const normalizeKey = (s) => {
    return String(s || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
  };

  // Function to parse Index row to get lot information
  const parseIndexRow = (header, row) => {
    const hmap = {};
    header.forEach((h, i) => (hmap[normalizeKey(h)] = i));

    const get = (key) => {
      const i = hmap[key];
      return i == null || i < 0 ? "" : row[i] ?? "";
    };

    const lot = String(get("lotnumber") || get("lot number") || get("lotno")).trim();
    if (!lot) return null;

    const startRow = parseInt(get("startrow") || "0", 10);
    const numRows = parseInt(get("numrows") || "0", 10);
    const headerCols = parseInt(get("headercols") || "0", 10);
    const fabric = get("fabric");
    const garmentType = get("garmenttype") || get("garment");
    const style = get("style");
    const savedAt = get("savedat");

    const sizes = String(get("sizes") || "")
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean);

    const shades = String(get("shades") || "")
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean);

    return { lot, startRow, numRows, headerCols, fabric, garmentType, style, sizes, shades, savedAt };
  };

  // Function to slice cutting matrix based on start row and number of rows
  const sliceCuttingMatrix = (bigValues, startRow, numRows) => {
    if (!Array.isArray(bigValues) || bigValues.length === 0) return [];
    if (!(startRow > 0 && numRows > 0)) return [];
    const r0 = Math.max(0, startRow - 1);
    const r1 = Math.min(bigValues.length - 1, r0 + numRows - 1);
    return bigValues.slice(r0, r1 + 1);
  };

  // Function to find header row index in cutting window
  const findHeaderRowIndex = (windowValues, expectedSizesNorm) => {
    const hasSizeToken = (rowSet) => expectedSizesNorm.some((sz) => rowSet.has(sz));

    for (let i = 0; i < windowValues.length; i++) {
      const row = windowValues[i] || [];
      const set = new Set(row.map((c) => normalizeKey(c)));
      const hasShadeHeader = set.has("color") || set.has("shade") || set.has("shades");
      if (hasShadeHeader && hasSizeToken(set)) return i;
    }

    for (let i = 0; i < windowValues.length; i++) {
      const row = windowValues[i] || [];
      const set = new Set(row.map((c) => normalizeKey(c)));
      let matches = 0;
      expectedSizesNorm.forEach((sz) => {
        if (set.has(sz)) matches++;
      });
      if (matches >= 2) return i;
    }
    return 0;
  };

  // Function to compute pending shades for a lot
  const computePendingShades = (windowValues, sizes = [], shades = []) => {
    if (!windowValues || windowValues.length === 0) {
      return new Set(shades.map(normalizeKey));
    }

    const normalizedSizes = Array.from(
      new Set((sizes || []).map((s) => normalizeKey(s)).filter(Boolean))
    );
    const headerRowIdx = findHeaderRowIndex(windowValues, normalizedSizes);
    const header = windowValues[headerRowIdx] || [];

    const hIdx = {};
    header.forEach((h, i) => {
      const k = normalizeKey(h);
      if (k && !(k in hIdx)) hIdx[k] = i;
    });

    const shadeColIndex = hIdx["color"] ?? hIdx["shade"] ?? hIdx["shades"] ?? 0;

    const nonSizeColumns = new Set([
      "color",
      "shade",
      "shades",
      "cuttingtable",
      "cutting",
      "table",
      "total",
      "totalpcs",
      "totals",
      "grandtotal",
      "sum",
      "lot",
      "style",
      "fabric",
      "garment",
      "partyname",
      "brand",
      "section",
      "season",
    ]);

    let sizeColIndices = [];
    header.forEach((h, i) => {
      const normalizedHeader = normalizeKey(h);
      if (normalizedHeader && !nonSizeColumns.has(normalizedHeader)) {
        sizeColIndices.push(i);
      }
    });

    if (sizeColIndices.length === 0) {
      normalizedSizes.forEach((ns) => {
        if (ns in hIdx) sizeColIndices.push(hIdx[ns]);
      });

      if (sizeColIndices.length === 0) {
        const ct = hIdx["cuttingtable"];
        if (ct != null && ct >= 0) {
          const guessStart = ct + 1;
          const guessed = [];
          for (let k = 0; k < normalizedSizes.length; k++) guessed.push(guessStart + k);
          sizeColIndices = Array.from(new Set(guessed.filter((g) => g < header.length)));
        }
      }
    }

    if (sizeColIndices.length === 0) {
      return new Set(shades.map(normalizeKey));
    }

    const shadeStats = new Map();

    for (let r = headerRowIdx + 1; r < windowValues.length; r++) {
      const row = windowValues[r] || [];
      const rawShade = String(row[shadeColIndex] || "").trim();
      const shadeKey = normalizeKey(rawShade);

      if (!shadeKey || shadeKey === "total" || shadeKey === "totals" || shadeKey === "grandtotal") {
        continue;
      }

      let hasPositiveData = false;
      let hasAnyData = false;
      let totalForThisRow = 0;

      sizeColIndices.forEach((c) => {
        const raw = row[c];
        if (raw != null && raw !== "") {
          hasAnyData = true;
          const n = parseFloat(String(raw).replace(/,/g, ""));
          if (!isNaN(n)) {
            totalForThisRow += n;
            if (n > 0) {
              hasPositiveData = true;
            }
          }
        }
      });

      if (hasAnyData) {
        if (hasPositiveData) {
          shadeStats.set(shadeKey, "found-with-data");
        } else {
          if (!shadeStats.has(shadeKey) || shadeStats.get(shadeKey) === "not-found") {
            shadeStats.set(shadeKey, "found-all-zero");
          }
        }
      } else {
        if (!shadeStats.has(shadeKey)) {
          shadeStats.set(shadeKey, "found-no-data");
        }
      }
    }

    const pendingShadeKeys = new Set();
    const expectedShadeKeys = (shades || []).map((sh) => normalizeKey(sh));

    expectedShadeKeys.forEach((shadeKey) => {
      const status = shadeStats.get(shadeKey);

      if (!status || status === "found-no-data") {
        pendingShadeKeys.add(shadeKey);
      }
    });

    return pendingShadeKeys;
  };

  // Function to format savedAt date
  const formatSavedAtToYMD = (savedAt) => {
    if (!savedAt) return "";
    const d = new Date(savedAt);
    if (isNaN(d.getTime())) return "";

    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
      "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const day = d.getDate();
    const monthName = monthNames[d.getMonth()];
    const year = d.getFullYear();

    return `${day} ${monthName} ${year}`;
  };

  // ===== SUPERVISOR SUMMARY FUNCTION WITH PCS QUANTITY =====
  const calculateSupervisorSummary = (data, headersArray, columnIndices) => {
    const supervisorIndex = findColumnIndex(headersArray, ['supervisor']);
    const completeStatusIndex = columnIndices.completeStatus;
    const cuttingQtyIndex = columnIndices.cuttingQty; // Get Cutting Qty index

    if (supervisorIndex === -1) {
      console.log('Supervisor column not found');
      return {};
    }

    const supervisorMap = new Map();

    // Function to parse PCS value
    const parsePcsValue = (value) => {
      if (!value) return 0;

      // Try to parse as integer first
      const num = parseInt(value);
      if (!isNaN(num)) return num;

      // If it's a string with commas or other formatting, extract numbers
      const match = value.toString().match(/\d+/);
      return match ? parseInt(match[0]) : 0;
    };

    data.forEach((row, idx) => {
      const supervisor = row[supervisorIndex];
      if (!supervisor || isEmptyCell(supervisor)) return;

      // Normalize the supervisor name
      const normalizedSupervisorName = normalizeSupervisorName(supervisor);
      const completeStatus = completeStatusIndex !== -1 ? row[completeStatusIndex] : null;
      const cuttingQty = cuttingQtyIndex !== -1 ? row[cuttingQtyIndex] : 0;
      const pcsValue = parsePcsValue(cuttingQty);

      // Check if lot is completed based on Completed Status column
      const isCompleted = isLotCompleted(completeStatus);

      if (!supervisorMap.has(normalizedSupervisorName)) {
        supervisorMap.set(normalizedSupervisorName, {
          name: normalizedSupervisorName,
          originalNames: new Set([supervisor.toString().trim()]),
          totalLots: 0,
          completedLots: 0,
          pendingLots: 0,
          totalPcs: 0,           // Total PCS across all lots
          completedPcs: 0,        // PCS from completed lots
          pendingPcs: 0,          // PCS from pending lots
          lotNumbers: [],
          pendingLotNumbers: [],
          completedLotNumbers: []
        });
      }

      const supervisorData = supervisorMap.get(normalizedSupervisorName);
      supervisorData.totalLots++;
      supervisorData.totalPcs += pcsValue;
      supervisorData.originalNames.add(supervisor.toString().trim());

      // Get lot number for reference
      const lotNumberIndex = columnIndices.lotNumber;
      const lotNumber = lotNumberIndex !== -1 ? row[lotNumberIndex] : `Lot-${idx}`;
      const lotNumberStr = lotNumber ? lotNumber.toString().trim() : `Lot-${idx}`;

      if (isCompleted) {
        supervisorData.completedLots++;
        supervisorData.completedPcs += pcsValue;
        supervisorData.completedLotNumbers.push(lotNumberStr);
      } else {
        supervisorData.pendingLots++;
        supervisorData.pendingPcs += pcsValue;
        supervisorData.pendingLotNumbers.push(lotNumberStr);
      }

      supervisorData.lotNumbers.push(lotNumberStr);
    });

    // Convert map to object and sort by supervisor name
    const result = {};
    Array.from(supervisorMap.keys())
      .sort()
      .forEach(key => {
        const data = supervisorMap.get(key);
        // Convert Set to Array for original names
        result[key] = {
          ...data,
          originalNames: Array.from(data.originalNames)
        };
      });

    return result;
  };

  // ===== FUNCTION FOR PENDING ISSUE AFTER EMB/PRINT DEPARTMENT =====
  // ===== UPDATED FUNCTION FOR PENDING ISSUE AFTER EMB/PRINT DEPARTMENT (EXCLUDING DIRECT) =====
  // ===== UPDATED FUNCTION FOR PENDING ISSUE AFTER EMB/PRINT DEPARTMENT (EXCLUDING DIRECT) =====
  const getPendingIssueAfterEmbPrintLots = () => {
    if (!sheetData.length) return [];

    // Apply filters to Index data
    let filteredIndexData = sheetData;
    if (filters.party !== 'all' || filters.brand !== 'all' || filters.item !== 'all') {
      filteredIndexData = applyFilters(sheetData, 'index');
    }

    // Get column indices
    const supervisorIndex = findColumnIndex(headers, ['supervisor']);
    const directStitchingIndex = displayColumnIndices.directStitching;
    const challanHistoryIndex = findColumnIndex(headers, ['challan history']);
    const completeStatusIndex = displayColumnIndices.completeStatus;
    const dateOfIssueIndex = findColumnIndex(headers, ['date of issue']);

    // Filter lots that meet the criteria:
    // 1. Supervisor is empty (not assigned)
    // 2. Has completed ALL embroidery/printing challans (NO Direct Stitching)
    // 3. Not completed yet
    // 4. Date of Issue is empty
    // 5. EXCLUDE Direct Stitching lots

    const pendingLots = filteredIndexData.filter(row => {
      const supervisorValue = supervisorIndex !== -1 ? row[supervisorIndex] : null;
      const directStitchingValue = directStitchingIndex !== -1 ? row[directStitchingIndex] : null;
      const challanHistory = challanHistoryIndex !== -1 ? row[challanHistoryIndex] : null;
      const completeStatus = completeStatusIndex !== -1 ? row[completeStatusIndex] : null;
      const dateOfIssueValue = dateOfIssueIndex !== -1 ? row[dateOfIssueIndex] : null;

      // EXCLUDE Direct Stitching lots
      const isDirectStitching = directStitchingValue &&
        directStitchingValue.toString().toLowerCase().includes('yes');
      if (isDirectStitching) return false;

      // Check if supervisor is empty
      const isSupervisorEmpty = isEmptyCell(supervisorValue);
      if (!isSupervisorEmpty) return false;

      // Check if lot is completed
      if (isLotCompleted(completeStatus)) return false;

      // Check if date of issue is empty (pending issue)
      const isDateEmpty = isEmptyCell(dateOfIssueValue);
      if (!isDateEmpty) return false;

      // Check if all embroidery/printing challans are completed
      const hasCompletedEmbPrint = challanHistory && hasCompletedEmbroideryOrPrinting(challanHistory);

      // Include ONLY if emb/print completed (no direct stitching)
      return hasCompletedEmbPrint;
    });

    return pendingLots;
  };

  // ===== FUNCTION TO ENRICH PENDING ISSUE LOTS WITH DETAILS =====
  // ===== UPDATED FUNCTION TO ENRICH PENDING ISSUE LOTS WITH DETAILS =====
  const enrichPendingIssueLotsWithDetails = (lots) => {
    if (!lots.length) return { lots, details: {} };

    const details = {};

    const lotNumberIndex = displayColumnIndices.lotNumber;
    const challanHistoryIndex = findColumnIndex(headers, ['challan history']);
    const cuttingQtyIndex = displayColumnIndices.cuttingQty;
    const fabricIndex = findColumnIndex(headers, ['fabric']);
    const garmentTypeIndex = findColumnIndex(headers, ['garment type']);
    const styleIndex = displayColumnIndices.style;
    const brandIndex = findColumnIndex(headers, ['brand']);

    lots.forEach(row => {
      const lotNumber = lotNumberIndex !== -1 ? row[lotNumberIndex] : 'Unknown';
      if (!lotNumber) return;

      const lotStr = lotNumber.toString().trim();
      const challanHistory = challanHistoryIndex !== -1 ? row[challanHistoryIndex] : null;
      const cuttingQty = cuttingQtyIndex !== -1 ? row[cuttingQtyIndex] : 0;
      const fabric = fabricIndex !== -1 ? row[fabricIndex] : 'Unknown';
      const garmentType = garmentTypeIndex !== -1 ? row[garmentTypeIndex] : 'Unknown';
      const style = styleIndex !== -1 ? row[styleIndex] : 'Unknown';
      const brand = brandIndex !== -1 ? row[brandIndex] : 'Unknown';

      // Get completed embroidery challans
      const completedEmbChallans = challanHistory ? getCompletedEmbroideryChallans(challanHistory) : [];

      // Get completed printing challans
      const completedPrintChallans = challanHistory ? getCompletedPrintingChallans(challanHistory) : [];

      // Get last emb/print date
      const lastEmbPrintDate = challanHistory ? getLastEmbPrintDate(challanHistory) : null;

      // Calculate days since last completed
      let daysSinceCompletion = null;
      if (lastEmbPrintDate) {
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const completionDate = new Date(lastEmbPrintDate);
        completionDate.setHours(0, 0, 0, 0);
        const diffTime = today - completionDate;
        daysSinceCompletion = Math.floor(diffTime / (1000 * 60 * 60 * 24));
      }

      // Parse PCS value
      const parsePcsValue = (value) => {
        if (!value) return 0;
        const num = parseInt(value);
        if (!isNaN(num)) return num;
        const match = value.toString().match(/\d+/);
        return match ? parseInt(match[0]) : 0;
      };

      details[lotStr] = {
        lotNumber: lotStr,
        fabric,
        garmentType,
        style,
        brand,
        completedEmbChallans: completedEmbChallans.length,
        completedPrintChallans: completedPrintChallans.length,
        totalCompletedChallans: completedEmbChallans.length + completedPrintChallans.length,
        lastCompletionDate: lastEmbPrintDate ? formatSavedAtToYMD(lastEmbPrintDate) : null,
        daysSinceCompletion: daysSinceCompletion,
        totalPcs: parsePcsValue(cuttingQty),
        // Track which type was completed
        hasEmbroidery: completedEmbChallans.length > 0,
        hasPrinting: completedPrintChallans.length > 0
      };
    });

    return { lots, details };
  };

  // ===== UPDATED FUNCTION: Get ALL lots from Index sheet =====
  const getCuttingDoneLots = () => {
    if (!sheetData.length) return [];

    // Apply filters to Index data
    let filteredIndexData = sheetData;
    if (filters.party !== 'all' || filters.brand !== 'all' || filters.item !== 'all') {
      filteredIndexData = applyFilters(sheetData, 'index');
    }

    return filteredIndexData;
  };

  // ===== UPDATED FUNCTION: Get Index data for Cutting Done =====
  const getIndexDataForCuttingDone = () => {
    if (!sheetData.length) return [];

    // Apply filters to Index data
    let filteredIndexData = sheetData;
    if (filters.party !== 'all' || filters.brand !== 'all' || filters.item !== 'all') {
      filteredIndexData = applyFilters(sheetData, 'index');
    }

    return filteredIndexData;
  };

  // Function to get JobOrder lots that are missing in Index sheet (excluding cancelled lots)
  const getMissingLotsFromJobOrder = () => {
    if (!jobOrderData.length || !sheetData.length) return [];

    const lotNumberIndexInJobOrder = findColumnIndex(jobOrderHeaders, ['lot number', 'lot no']);
    const lotNumberIndexInIndex = displayColumnIndices.lotNumber;

    if (lotNumberIndexInJobOrder === -1 || lotNumberIndexInIndex === -1) {
      console.log('Could not find lot number columns');
      return [];
    }

    // Get all UNIQUE lot numbers from Index sheet
    const indexLotNumbers = new Set();
    sheetData.forEach((row) => {
      const lotNumber = row[lotNumberIndexInIndex];
      if (lotNumber && !isEmptyCell(lotNumber)) {
        // Normalize the lot number (remove extra spaces, convert to uppercase for consistency)
        const cleanLot = lotNumber.toString().trim().toUpperCase().replace(/\s+/g, ' ');
        indexLotNumbers.add(cleanLot);
      }
    });

    // Track unique lot numbers from JobOrder to avoid counting duplicates
    const processedJobLots = new Set();
    const missingRows = [];
    const cancelledLots = [];

    jobOrderData.forEach(row => {
      const jobOrderLotNumber = row[lotNumberIndexInJobOrder];
      if (!jobOrderLotNumber || isEmptyCell(jobOrderLotNumber)) {
        return;
      }

      // Normalize the lot number the same way
      const cleanJobLot = jobOrderLotNumber.toString().trim().toUpperCase().replace(/\s+/g, ' ');

      // Skip if we've already processed this lot number (avoid duplicates)
      if (processedJobLots.has(cleanJobLot)) {
        return;
      }
      processedJobLots.add(cleanJobLot);

      if (isLotCancelled(row)) {
        cancelledLots.push(cleanJobLot);
      } else {
        // Check if this lot exists in Index sheet
        const isInIndex = indexLotNumbers.has(cleanJobLot);

        if (!isInIndex) {
          missingRows.push(row);
        }
      }
    });

    console.log('Fabric Pending Calculation:', {
      uniqueJobLots: processedJobLots.size,
      uniqueIndexLots: indexLotNumbers.size,
      cancelledLots: cancelledLots.length,
      missingLots: missingRows.length,
      formula: `${processedJobLots.size} (Unique Job Lots) - ${cancelledLots.length} (Cancelled) - ${processedJobLots.size - cancelledLots.length - missingRows.length} (Found in Index) = ${missingRows.length} (Missing)`
    });

    return missingRows;
  };

  // ===== UPDATED FUNCTION: Get lots with pending colours/shades =====
  const getColourPendingLots = (filteredJobData = jobOrderData) => {
    if (!cuttingData.length || !indexData.length || !filteredJobData.length) {
      return { lots: [], details: {} };
    }

    const colourPendingLotsList = [];
    const colourPendingDetailsMap = {};

    // Create a map of index data by lot number
    const indexMap = new Map();
    indexData.forEach((entry) => {
      if (entry && entry.lot) {
        indexMap.set(entry.lot, entry);
      }
    });

    // Get lot numbers from JobOrder data
    const lotNumberIndexInJobOrder = findColumnIndex(jobOrderHeaders, ['lot number', 'lot no']);
    if (lotNumberIndexInJobOrder === -1) return { lots: [], details: {} };

    // For each lot in filtered JobOrder data, check if it has pending shades
    filteredJobData.forEach((row) => {
      const lotNumber = row[lotNumberIndexInJobOrder];
      if (!lotNumber || isEmptyCell(lotNumber)) return;

      const lotStr = lotNumber.toString().trim();
      const indexEntry = indexMap.get(lotStr);

      // Skip if no index entry (fabric pending lots)
      if (!indexEntry) return;

      // Get cutting window for this lot
      const window = sliceCuttingMatrix(cuttingData, indexEntry.startRow, indexEntry.numRows);

      // Compute pending shades
      const pendingShadeKeys = computePendingShades(window, indexEntry.sizes, indexEntry.shades);

      if (pendingShadeKeys.size > 0) {
        // Map shade keys back to original shade names
        const shadeKeyToOriginal = new Map((indexEntry.shades || []).map((sh) => [normalizeKey(sh), sh]));
        const pendingList = Array.from(pendingShadeKeys).map(
          (k) => shadeKeyToOriginal.get(k) || k
        );

        // Find the matching row in sheetData (Index sheet) for display
        const lotNumberIndexInIndex = displayColumnIndices.lotNumber;
        let indexRow = null;

        if (lotNumberIndexInIndex !== -1) {
          indexRow = sheetData.find(r => {
            const idxLot = r[lotNumberIndexInIndex];
            return idxLot && idxLot.toString().trim() === lotStr;
          });
        }

        // If we found a matching index row, use that for display
        if (indexRow) {
          colourPendingLotsList.push(indexRow);
        } else {
          // Create a minimal row with just the lot number
          const minimalRow = [...(new Array(headers.length).fill(''))];
          if (lotNumberIndexInIndex !== -1) {
            minimalRow[lotNumberIndexInIndex] = lotStr;
          }
          colourPendingLotsList.push(minimalRow);
        }

        colourPendingDetailsMap[lotStr] = {
          lotNumber: lotStr,
          pendingShades: pendingList,
          totalPending: pendingList.length,
          fabric: indexEntry.fabric,
          garmentType: indexEntry.garmentType,
          style: indexEntry.style,
          cuttingDate: formatSavedAtToYMD(indexEntry.savedAt)
        };
      }
    });

    return { lots: colourPendingLotsList, details: colourPendingDetailsMap };
  };

  // Extract filter options from data
  const extractFilterOptions = (data, headersArray) => {
    const parties = new Set();
    const brands = new Set();
    const items = new Set();
    const seasons = new Set();

    // Find column indices for party, brand, item, and season columns
    const partyIndex = findColumnIndex(headersArray, ['party', 'client', 'customer', 'buyer']);
    const brandIndex = findColumnIndex(headersArray, ['brand']);
    const itemIndex = findColumnIndex(headersArray, ['item', 'product', 'garment', 'style', 'garment type']);
    const seasonIndex = findColumnIndex(headersArray, ['season']);

    data.forEach(row => {
      if (partyIndex !== -1 && row[partyIndex] && !isEmptyCell(row[partyIndex])) {
        parties.add(row[partyIndex].toString().trim());
      }

      if (brandIndex !== -1 && row[brandIndex] && !isEmptyCell(row[brandIndex])) {
        brands.add(row[brandIndex].toString().trim());
      }

      if (itemIndex !== -1 && row[itemIndex] && !isEmptyCell(row[itemIndex])) {
        items.add(row[itemIndex].toString().trim());
      }

      // Extract seasons
      if (seasonIndex !== -1 && row[seasonIndex] && !isEmptyCell(row[seasonIndex])) {
        seasons.add(row[seasonIndex].toString().trim());
      }
    });

    return {
      parties: ['all', ...Array.from(parties).sort()],
      brands: ['all', ...Array.from(brands).sort()],
      items: ['all', ...Array.from(items).sort()],
      seasons: ['all', ...Array.from(seasons).sort()]
    };
  };

  // Fetch all sheets from different spreadsheets
  const fetchData = async () => {
    try {
      setLoading(true);
      setError(null);

      setDepartmentSummary({});
      setFilteredDepartmentSummary({});

      // Fetch Index sheet from first spreadsheet
      const indexUrl = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${INDEX_SHEET_NAME}!${RANGE}?key=${API_KEY}`;

      // Fetch JobOrder sheet from second spreadsheet
      const jobOrderUrl = `https://sheets.googleapis.com/v4/spreadsheets/${JOBORDER_SHEET_ID}/values/${JOBORDER_SHEET_NAME}!${RANGE}?key=${API_KEY}`;

      // Fetch Cutting sheet from first spreadsheet
      const cuttingUrl = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${CUTTING_BIG_RANGE}?key=${API_KEY}`;

      // Fetch all sheets in parallel
      const [indexResponse, jobOrderResponse, cuttingResponse] = await Promise.all([
        axios.get(indexUrl),
        axios.get(jobOrderUrl),
        axios.get(cuttingUrl)
      ]);

      // Process Index sheet - WITH BLANK ROW FILTERING
      if (indexResponse.data && indexResponse.data.values) {
        const indexData = indexResponse.data.values;

        if (indexData.length > 0) {
          setHeaders(indexData[0]);

          // Filter out completely blank rows from Index sheet
          const allIndexData = indexData.slice(1).filter(row => {
            // Check if row exists and has at least one non-empty cell
            if (!row || !Array.isArray(row)) return false;

            // Check if the row has any meaningful data
            return row.some(cell => {
              if (cell === undefined || cell === null) return false;

              const cellValue = cell.toString().trim();
              // Return true if cell has actual content (not empty, not placeholder)
              return cellValue !== '' &&
                cellValue !== '-' &&
                cellValue !== 'null' &&
                cellValue !== 'undefined' &&
                cellValue !== 'N/A';
            });
          });

          console.log('Index sheet loaded:', allIndexData.length, 'rows (filtered from', indexData.length - 1, 'total rows)');
          setSheetData(allIndexData);

          // Find indices for display columns
          const columnIndices = {};
          displayColumns.forEach(col => {
            const index = findColumnIndex(indexData[0], col.searchTerms);
            columnIndices[col.key] = index;
          });
          setDisplayColumnIndices(columnIndices);

          // Calculate supervisor summary with PCS quantities
          const supSummary = calculateSupervisorSummary(allIndexData, indexData[0], columnIndices);
          setSupervisorSummary(supSummary);
          setFilteredSupervisorSummary(supSummary);

          // Extract filter options from Index sheet
          const options = extractFilterOptions(allIndexData, indexData[0]);
          setFilterOptions(options);

          // Parse Index sheet for cutting data reference
          const parsedIndexEntries = [];
          for (let i = 1; i < indexData.length; i++) {
            const entry = parseIndexRow(indexData[0], indexData[i]);
            if (entry) parsedIndexEntries.push(entry);
          }
          setIndexData(parsedIndexEntries);
          console.log('Parsed Index entries:', parsedIndexEntries.length);
        } else {
          console.log('Index sheet is empty');
          setSheetData([]);
          setIndexData([]);
        }
      } else {
        throw new Error('No data found in the Index sheet');
      }

      // Process JobOrder sheet
      if (jobOrderResponse.data && jobOrderResponse.data.values) {
        const jobOrderDataValues = jobOrderResponse.data.values;

        if (jobOrderDataValues.length > 0) {
          setJobOrderHeaders(jobOrderDataValues[0]);
          const allJobOrderData = jobOrderDataValues.slice(1);
          setJobOrderData(allJobOrderData);

          // Find indices for JobOrder display columns
          const jobOrderColumnIndices = {};
          jobOrderDisplayColumns.forEach(col => {
            const index = findColumnIndex(jobOrderDataValues[0], col.searchTerms);
            jobOrderColumnIndices[col.key] = index;
          });
          setJobOrderColumnIndices(jobOrderColumnIndices);

        } else {
          console.log('JobOrder sheet is empty');
          setJobOrderData([]);
        }
      } else {
        console.warn('No data found in the JobOrder sheet, but continuing...');
        setJobOrderData([]);
      }

      // Process Cutting sheet
      if (cuttingResponse.data && cuttingResponse.data.values) {
        const cuttingDataValues = cuttingResponse.data.values;
        setCuttingData(cuttingDataValues);
        console.log('Cutting sheet loaded:', cuttingDataValues.length, 'rows');
      } else {
        console.warn('No data found in the Cutting sheet');
        setCuttingData([]);
      }

      // Calculate department summaries
      if (sheetData.length > 0 || jobOrderData.length > 0 || cuttingData.length > 0) {
        // Get colour pending lots with current jobOrderData
        const { lots, details } = getColourPendingLots(jobOrderData);
        setColourPendingLots(lots);
        setColourPendingDetails(details);

        // Get pending issue after emb/print lots
        const pendingLots = getPendingIssueAfterEmbPrintLots();
        const { lots: enrichedLots, details: enrichedDetails } = enrichPendingIssueLotsWithDetails(pendingLots);
        setPendingIssueLots(enrichedLots);
        setPendingIssueDetails(enrichedDetails);

        const summaries = calculateDepartmentSummaries(
          sheetData,
          headers,
          displayColumnIndices,
          jobOrderData,
          jobOrderHeaders,
          cuttingData,
          indexData,
          lots,
          details,
          enrichedLots,
          enrichedDetails
        );
        setDepartmentSummary(summaries);
        setFilteredDepartmentSummary(summaries);
      }

    } catch (err) {
      console.error('Error fetching Google Sheet data:', err);

      // Check if it's just the JobOrder sheet that failed
      if (err.message.includes('JobOrder') || err.response?.status === 404) {
        setError(`Error loading JobOrder sheet: ${err.message}. The Index and Cutting sheets data is loaded.`);
        // Continue with just Index and Cutting data
        if (sheetData.length > 0) {
          const summaries = calculateDepartmentSummaries(sheetData, headers, displayColumnIndices, [], [], cuttingData, indexData, [], {}, [], {});
          setDepartmentSummary(summaries);
          setFilteredDepartmentSummary(summaries);
        }
      } else {
        setError(`Error: ${err.message}. Please check your API key, spreadsheet IDs, and sheet names.`);
      }
    } finally {
      setLoading(false);
    }
  };

  // Function to get pending details for a row
  const getPendingDetails = (row) => {
    if (activeDepartment !== 'embroidery' && activeDepartment !== 'printing' &&
      activeDepartment !== 'colour_pending' && activeDepartment !== 'pending_after_emb_print') return null;

    if (activeDepartment === 'colour_pending') {
      const lotNumberIndex = displayColumnIndices.lotNumber;
      if (lotNumberIndex === -1) return null;

      const lotNumber = row[lotNumberIndex];
      if (!lotNumber) return null;

      return colourPendingDetails[lotNumber.toString().trim()] || null;
    }

    if (activeDepartment === 'pending_after_emb_print') {
      const lotNumberIndex = displayColumnIndices.lotNumber;
      if (lotNumberIndex === -1) return null;

      const lotNumber = row[lotNumberIndex];
      if (!lotNumber) return null;

      return pendingIssueDetails[lotNumber.toString().trim()] || null;
    }

    const challanHistoryIndex = findColumnIndex(headers, ['challan history']);
    const challanHistory = row[challanHistoryIndex];

    if (activeDepartment === 'embroidery') {
      return getPendingEmbroideryDetails(challanHistory);
    } else if (activeDepartment === 'printing') {
      return getPendingPrintingDetails(challanHistory);
    }

    return null;
  };

  // Function to calculate unique styles from JobOrder data
  const getUniqueStylesFromJobOrder = (data) => {
    const styleIndex = findColumnIndex(jobOrderHeaders, ['style']);
    if (styleIndex === -1) return 0;

    const styles = new Set();
    data.forEach(row => {
      const style = row[styleIndex];
      if (style && !isEmptyCell(style)) {
        styles.add(style.toString().trim());
      }
    });

    return styles.size;
  };

  // Function to calculate total quantity from Index data
  const getTotalQuantityFromIndex = (data) => {
    const pcsIndex = displayColumnIndices.cuttingQty;
    if (pcsIndex === -1) return 0;

    const parsePcsValue = (value) => {
      if (!value) return 0;

      const num = parseInt(value);
      if (!isNaN(num)) return num;

      const match = value.toString().match(/\d+/);
      return match ? parseInt(match[0]) : 0;
    };

    return data.reduce((sum, row) => sum + parsePcsValue(row[pcsIndex]), 0);
  };

  // Function to calculate total quantity from JobOrder data
  const getTotalQuantityFromJobOrder = (data) => {
    const quantityIndex = findColumnIndex(jobOrderHeaders, ['quantity']);
    if (quantityIndex === -1) return 0;

    const parseQtyValue = (value) => {
      if (!value) return 0;

      const num = parseInt(value);
      if (!isNaN(num)) return num;

      const match = value.toString().match(/\d+/);
      return match ? parseInt(match[0]) : 0;
    };

    return data.reduce((sum, row) => sum + parseQtyValue(row[quantityIndex]), 0);
  };

  const calculateDepartmentSummaries = (
    indexData,
    headersArray,
    columnIndices,
    jobData,
    jobHeaders,
    cuttingSheetData,
    indexEntries,
    filteredColourLots = colourPendingLots,
    filteredColourDetails = colourPendingDetails,
    filteredPendingIssueLots = pendingIssueLots,
    filteredPendingIssueDetails = pendingIssueDetails
  ) => {
    const summaries = {};

    // Get indices from columnIndices
    const styleIndex = columnIndices.style;
    const pcsIndex = columnIndices.cuttingQty;
    const directStitchingIndex = columnIndices.directStitching;
    const completeStatusIndex = columnIndices.completeStatus;
    const challanHistoryIndex = findColumnIndex(headersArray, ['challan history']);
    const dateOfIssueIndex = findColumnIndex(headersArray, ['date of issue']);
    const supervisorIndex = findColumnIndex(headersArray, ['supervisor']);
    const dateIndex = findColumnIndex(headersArray, ['date', 'saved at', 'joborder date']);

    // Function to parse PCS value
    const parsePcsValue = (value) => {
      if (!value) return 0;

      const num = parseInt(value);
      if (!isNaN(num)) return num;

      const match = value.toString().match(/\d+/);
      return match ? parseInt(match[0]) : 0;
    };

    // Helper function to check if a cell is empty
    const isEmptyCell = (cellValue) => {
      if (cellValue === undefined || cellValue === null) return true;

      const strValue = cellValue.toString().trim();
      return strValue === '' || strValue === '-' || strValue === 'null' || strValue === 'undefined' || strValue === 'N/A';
    };

    // Helper function to check if lot is cancelled in JobOrder
    const isLotCancelled = (row) => {
      const statusIndex = findColumnIndex(jobHeaders, ['status']);
      if (statusIndex === -1) return false;

      const statusValue = row[statusIndex];
      if (!statusValue) return false;

      const statusStr = statusValue.toString().toLowerCase().trim();
      return statusStr.includes('cancel') || statusStr.includes('cancelled');
    };

    // Helper function to get missing lots from JobOrder (not in Index)
    const getMissingLotsFromJobOrder = () => {
      if (!jobData.length || !indexData.length) return [];

      const lotNumberIndexInJobOrder = findColumnIndex(jobHeaders, ['lot number', 'lot no']);
      const lotNumberIndexInIndex = columnIndices.lotNumber;

      if (lotNumberIndexInJobOrder === -1 || lotNumberIndexInIndex === -1) return [];

      // Get all lot numbers from Index sheet
      const indexLotNumbers = new Set();
      indexData.forEach((row) => {
        const lotNumber = row[lotNumberIndexInIndex];
        if (lotNumber && !isEmptyCell(lotNumber)) {
          indexLotNumbers.add(lotNumber.toString().trim());
        }
      });

      // Get all non-cancelled JobOrder lots
      const nonCancelledJobOrders = [];

      jobData.forEach(row => {
        const jobOrderLotNumber = row[lotNumberIndexInJobOrder];
        if (!jobOrderLotNumber || isEmptyCell(jobOrderLotNumber)) {
          return;
        }

        const cleanJobLot = jobOrderLotNumber.toString().trim();

        if (!isLotCancelled(row)) {
          nonCancelledJobOrders.push({
            lotNumber: cleanJobLot,
            row: row
          });
        }
      });

      // Find JobOrder lots that are NOT in Index sheet
      const missingLots = nonCancelledJobOrders
        .filter(item => !indexLotNumbers.has(item.lotNumber))
        .map(item => item.row);

      return missingLots;
    };

    // Helper function to get unique styles from JobOrder data
    const getUniqueStylesFromJobOrder = (data) => {
      const styleIndex = findColumnIndex(jobHeaders, ['style']);
      if (styleIndex === -1) return 0;

      const styles = new Set();
      data.forEach(row => {
        const style = row[styleIndex];
        if (style && !isEmptyCell(style)) {
          styles.add(style.toString().trim());
        }
      });

      return styles.size;
    };

    // Helper function to get total quantity from Index data
    const getTotalQuantityFromIndex = (data) => {
      const pcsIndex = columnIndices.cuttingQty;
      if (pcsIndex === -1) return 0;

      return data.reduce((sum, row) => sum + parsePcsValue(row[pcsIndex]), 0);
    };

    departments.forEach(dept => {
      if (dept.id === 'all') {
        // For All Data department:
        // Total Records = JobOrder total records (excluding cancelled)
        const nonCancelledJobOrders = jobData.filter(row => !isLotCancelled(row)).length;
        const jobOrderStyles = getUniqueStylesFromJobOrder(jobData);
        const indexTotalQty = getTotalQuantityFromIndex(indexData);

        // Count Direct Stitching lots from Index
        let directStitchingLots = 0;
        if (directStitchingIndex !== -1) {
          directStitchingLots = indexData.filter(row => {
            const dsValue = row[directStitchingIndex];
            return dsValue && dsValue.toString().toLowerCase().includes('yes');
          }).length;
        }

        summaries['all'] = {
          totalRows: nonCancelledJobOrders,
          totalPcs: indexTotalQty,
          styles: jobOrderStyles,
          lastUpdated: indexData[0]?.[dateIndex] || 'N/A',
          indexRecords: indexData.length,
          jobOrderRecords: nonCancelledJobOrders,
          directStitchingLots: directStitchingLots
        };
        return;
      }

      // Cutting Done = ALL Index sheet data
      if (dept.id === 'cutting_done') {
        const totalQty = getTotalQuantityFromIndex(indexData);

        const styles = new Set();
        indexData.forEach(row => {
          if (styleIndex !== -1) {
            const style = row[styleIndex];
            if (style && !isEmptyCell(style)) {
              styles.add(style.toString().trim());
            }
          }
        });

        summaries['cutting_done'] = {
          totalRows: indexData.length,
          totalPcs: totalQty,
          styles: styles.size,
          lastUpdated: indexData[0]?.[dateIndex] || 'N/A',
          indexTotalRecords: indexData.length
        };
        return;
      }

      if (dept.id === 'fabric_pending') {
        // Get missing lots from JobOrder (not in Index) - using UNIQUE lots
        const missingLots = getMissingLotsFromJobOrder();

        // Count cancelled lots in JobOrder (unique lots)
        const lotNumberIndexInJobOrder = findColumnIndex(jobHeaders, ['lot number', 'lot no']);
        const processedCancelled = new Set();
        let cancelledLotsCount = 0;

        jobData.forEach(row => {
          if (isLotCancelled(row)) {
            const lotNumber = row[lotNumberIndexInJobOrder];
            if (lotNumber && !isEmptyCell(lotNumber)) {
              const cleanLot = lotNumber.toString().trim().toUpperCase();
              if (!processedCancelled.has(cleanLot)) {
                processedCancelled.add(cleanLot);
                cancelledLotsCount++;
              }
            }
          }
        });

        // Get total unique non-cancelled JobOrder lots
        const processedNonCancelled = new Set();
        jobData.forEach(row => {
          if (!isLotCancelled(row)) {
            const lotNumber = row[lotNumberIndexInJobOrder];
            if (lotNumber && !isEmptyCell(lotNumber)) {
              const cleanLot = lotNumber.toString().trim().toUpperCase();
              processedNonCancelled.add(cleanLot);
            }
          }
        });

        // Get unique Index lots count
        const lotNumberIndexInIndex = columnIndices.lotNumber;
        const uniqueIndexLots = new Set();
        indexData.forEach(row => {
          const lotNumber = row[lotNumberIndexInIndex];
          if (lotNumber && !isEmptyCell(lotNumber)) {
            uniqueIndexLots.add(lotNumber.toString().trim().toUpperCase());
          }
        });

        // Get unique styles from missing lots
        const styleIndexJobOrder = findColumnIndex(jobHeaders, ['style']);
        const styles = new Set();
        missingLots.forEach(row => {
          if (styleIndexJobOrder !== -1) {
            const style = row[styleIndexJobOrder];
            if (style && !isEmptyCell(style)) {
              styles.add(style.toString().trim());
            }
          }
        });

        summaries['fabric_pending'] = {
          totalRows: missingLots.length, // Number of UNIQUE missing lots
          totalPcs: 0,
          styles: styles.size,
          lastUpdated: 'N/A',
          missingLotsCount: missingLots.length,
          jobOrderTotal: processedNonCancelled.size + cancelledLotsCount, // Total unique lots
          cancelledLots: cancelledLotsCount,
          nonCancelledJobOrders: processedNonCancelled.size,
          indexRecords: uniqueIndexLots.size,
          calculation: `${processedNonCancelled.size} (Non-Cancelled Jobs) - ${uniqueIndexLots.size} (Index) = ${missingLots.length} (Pending)`,
          // Add detailed breakdown
          breakdown: {
            totalUniqueJobLots: processedNonCancelled.size + cancelledLotsCount,
            uniqueIndexLots: uniqueIndexLots.size,
            uniqueCancelledLots: cancelledLotsCount,
            uniqueNonCancelledLots: processedNonCancelled.size,
            uniqueMissingLots: missingLots.length,
            matchCount: processedNonCancelled.size - missingLots.length
          }
        };
        return;
      }

      if (dept.id === 'direct_stitching') {
        // Direct Stitching lots from Index sheet
        let directStitchingData = [];
        if (directStitchingIndex !== -1) {
          directStitchingData = indexData.filter(row => {
            const dsValue = row[directStitchingIndex];
            return dsValue && dsValue.toString().toLowerCase().includes('yes');
          });
        }

        const departmentPcs = directStitchingData.reduce((sum, row) => sum + parsePcsValue(row[pcsIndex]), 0);

        summaries['direct_stitching'] = {
          totalRows: directStitchingData.length,
          totalPcs: departmentPcs || 0,
          styles: [...new Set(directStitchingData.map(row => row[styleIndex]).filter(Boolean))].length,
          lastUpdated: directStitchingData[0]?.[dateIndex] || 'N/A',
          yesCount: directStitchingData.length,
          noCount: indexData.length - directStitchingData.length
        };
        return;
      }

      // For Direct Pending Issue department
      if (dept.id === 'direct_pending_issue' && dept.conditions) {
        const headersArray = headers;
        const dateOfIssueIndex = findColumnIndex(headersArray, ['date of issue']);
        const supervisorIndex = findColumnIndex(headersArray, ['supervisor']);
        const directStitchingIndex = columnIndices.directStitching;

        // Filter index data for:
        // 1. Direct Stitching = Yes
        // 2. Date of Issue is empty
        // 3. Supervisor is empty
        let deptData = [];
        if (directStitchingIndex !== -1 && dateOfIssueIndex !== -1 && supervisorIndex !== -1) {
          deptData = indexData.filter(row => {
            const dsValue = row[directStitchingIndex];
            const isDirectStitching = dsValue && dsValue.toString().toLowerCase().includes('yes');

            const dateOfIssueValue = row[dateOfIssueIndex];
            const isDateEmpty = isEmptyCell(dateOfIssueValue);

            const supervisorValue = row[supervisorIndex];
            const isSupervisorEmpty = isEmptyCell(supervisorValue);

            return isDirectStitching && isDateEmpty && isSupervisorEmpty;
          });
        }

        const departmentPcs = deptData.reduce((sum, row) => sum + parsePcsValue(row[pcsIndex]), 0);

        summaries['direct_pending_issue'] = {
          totalRows: deptData.length,
          totalPcs: departmentPcs || 0,
          styles: [...new Set(deptData.map(row => row[styleIndex]).filter(Boolean))].length,
          lastUpdated: deptData[0]?.[dateIndex] || 'N/A',
          pendingIssueCount: deptData.length
        };
        return;
      }

      // For stitching department
      if (dept.id === 'stitching' && dept.conditions) {
        const deptData = indexData.filter(row => {
          const dateOfIssueValue = row[dateOfIssueIndex];
          const supervisorValue = row[supervisorIndex];
          const completeStatusValue = row[completeStatusIndex];

          const isIssued = !isEmptyCell(dateOfIssueValue) && !isEmptyCell(supervisorValue);
          const isNotCompleted = !isLotCompleted(completeStatusValue);

          return isIssued && isNotCompleted;
        });

        const departmentPcs = deptData.reduce((sum, row) => sum + parsePcsValue(row[pcsIndex]), 0);

        summaries['stitching'] = {
          totalRows: deptData.length,
          totalPcs: departmentPcs || 0,
          styles: [...new Set(deptData.map(row => row[styleIndex]).filter(Boolean))].length,
          lastUpdated: deptData[0]?.[dateIndex] || 'N/A',
          stitchingIssued: deptData.length
        };
        return;
      }

      // For pending stitching department
      if (dept.id === 'pending_stitching' && dept.conditions) {
        const deptData = indexData.filter(row => {
          const supervisorValue = row[supervisorIndex];

          const isSupervisorEmpty = isEmptyCell(supervisorValue);

          const hasDisplayColumnData = Object.values(columnIndices).some(index => {
            if (index === -1) return false;
            const cellValue = row[index];
            return !isEmptyCell(cellValue);
          });

          const hasAnyData = row.some((cell, idx) => {
            if (idx === supervisorIndex) return false;
            return !isEmptyCell(cell);
          });

          return isSupervisorEmpty && (hasDisplayColumnData || hasAnyData);
        });

        const departmentPcs = deptData.reduce((sum, row) => sum + parsePcsValue(row[pcsIndex]), 0);

        summaries['pending_stitching'] = {
          totalRows: deptData.length,
          totalPcs: departmentPcs || 0,
          styles: [...new Set(deptData.map(row => row[styleIndex]).filter(Boolean))].length,
          lastUpdated: deptData[0]?.[dateIndex] || 'N/A',
          pendingStitching: deptData.length
        };
        return;
      }

      // For embroidery department
      if (dept.id === 'embroidery' && dept.conditions) {
        const deptData = indexData.filter(row => {
          const challanHistory = row[challanHistoryIndex];
          if (!challanHistory) return false;

          return isEmbroideryLotPending(challanHistory);
        });

        const departmentPcs = deptData.reduce((sum, row) => sum + parsePcsValue(row[pcsIndex]), 0);

        let pendingLotsCount = 0;
        let totalPendingQty = 0;

        deptData.forEach(row => {
          const challanHistory = row[challanHistoryIndex];
          const pendingLots = getPendingEmbroideryDetails(challanHistory);

          pendingLotsCount += pendingLots.length;
          totalPendingQty += pendingLots.reduce((sum, lot) => sum + (lot.totalQty || 0), 0);
        });

        summaries['embroidery'] = {
          totalRows: deptData.length,
          totalPcs: departmentPcs || 0,
          styles: [...new Set(deptData.map(row => row[styleIndex]).filter(Boolean))].length,
          lastUpdated: deptData[0]?.[dateIndex] || 'N/A',
          pendingLots: pendingLotsCount,
          pendingQty: totalPendingQty
        };
        return;
      }

      // For printing department
      if (dept.id === 'printing' && dept.conditions) {
        const deptData = indexData.filter(row => {
          const challanHistory = row[challanHistoryIndex];
          if (!challanHistory) return false;

          return isPrintingLotPending(challanHistory);
        });

        const departmentPcs = deptData.reduce((sum, row) => sum + parsePcsValue(row[pcsIndex]), 0);

        let pendingLotsCount = 0;
        let totalPendingQty = 0;

        deptData.forEach(row => {
          const challanHistory = row[challanHistoryIndex];
          const pendingLots = getPendingPrintingDetails(challanHistory);

          pendingLotsCount += pendingLots.length;
          totalPendingQty += pendingLots.reduce((sum, lot) => sum + (lot.totalQty || 0), 0);
        });

        summaries['printing'] = {
          totalRows: deptData.length,
          totalPcs: departmentPcs || 0,
          styles: [...new Set(deptData.map(row => row[styleIndex]).filter(Boolean))].length,
          lastUpdated: deptData[0]?.[dateIndex] || 'N/A',
          pendingLots: pendingLotsCount,
          pendingQty: totalPendingQty
        };
        return;
      }

      // For pending after emb/print department
      // For pending after emb/print department (EXCLUDING DIRECT STITCHING)
      if (dept.id === 'pending_after_emb_print' && dept.conditions) {
        // Get pending issue after emb/print lots
        const pendingData = filteredPendingIssueLots;

        // Get unique styles from pending lots
        const styles = new Set();
        pendingData.forEach(row => {
          if (styleIndex !== -1) {
            const style = row[styleIndex];
            if (style && !isEmptyCell(style)) {
              styles.add(style.toString().trim());
            }
          }
        });

        // Calculate total PCS
        const totalPcs = pendingData.reduce((sum, row) => sum + parsePcsValue(row[pcsIndex]), 0);

        // Calculate counts for embroidery vs printing completed
        let embroideryOnlyCount = 0;
        let printingOnlyCount = 0;
        let bothEmbPrintCount = 0;

        pendingData.forEach(row => {
          const lotNumberIndex = columnIndices.lotNumber;
          if (lotNumberIndex !== -1) {
            const lotNumber = row[lotNumberIndex];
            if (lotNumber) {
              const details = filteredPendingIssueDetails[lotNumber.toString().trim()];
              if (details) {
                if (details.hasEmbroidery && details.hasPrinting) {
                  bothEmbPrintCount++;
                } else if (details.hasEmbroidery) {
                  embroideryOnlyCount++;
                } else if (details.hasPrinting) {
                  printingOnlyCount++;
                }
              }
            }
          }
        });

        // Calculate average days since completion
        let totalDays = 0;
        let lotsWithDays = 0;

        pendingData.forEach(row => {
          const lotNumberIndex = columnIndices.lotNumber;
          if (lotNumberIndex !== -1) {
            const lotNumber = row[lotNumberIndex];
            if (lotNumber) {
              const details = filteredPendingIssueDetails[lotNumber.toString().trim()];
              if (details && details.daysSinceCompletion !== null) {
                totalDays += details.daysSinceCompletion;
                lotsWithDays++;
              }
            }
          }
        });

        const avgDays = lotsWithDays > 0 ? Math.round(totalDays / lotsWithDays) : 0;

        summaries['pending_after_emb_print'] = {
          totalRows: pendingData.length,
          totalPcs: totalPcs,
          styles: styles.size,
          lastUpdated: pendingData[0]?.[dateIndex] || 'N/A',
          embroideryOnlyCount: embroideryOnlyCount,
          printingOnlyCount: printingOnlyCount,
          bothEmbPrintCount: bothEmbPrintCount,
          avgDaysSinceCompletion: avgDays
        };
        return;
      }
      // For colour pending department
      if (dept.id === 'colour_pending' && dept.conditions) {
        // Get colour pending lots from filtered data
        const colourPendingData = filteredColourLots;

        // Get unique styles from colour pending lots
        const styles = new Set();
        colourPendingData.forEach(row => {
          if (styleIndex !== -1) {
            const style = row[styleIndex];
            if (style && !isEmptyCell(style)) {
              styles.add(style.toString().trim());
            }
          }
        });

        // Calculate total pending shades count
        let totalPendingShades = 0;
        colourPendingData.forEach(row => {
          const lotNumberIndex = columnIndices.lotNumber;
          if (lotNumberIndex !== -1) {
            const lotNumber = row[lotNumberIndex];
            if (lotNumber) {
              const details = filteredColourDetails[lotNumber.toString().trim()];
              if (details) {
                totalPendingShades += details.totalPending || 0;
              }
            }
          }
        });

        summaries['colour_pending'] = {
          totalRows: colourPendingData.length,
          totalPcs: 0,
          styles: styles.size,
          lastUpdated: colourPendingData[0]?.[dateIndex] || 'N/A',
          pendingShades: totalPendingShades
        };
        return;
      }

      // For packing department
      if (dept.id === 'packing') {
        const deptData = indexData.filter(row => {
          return dept.searchTerms.some(term =>
            row.some(cell =>
              cell && cell.toString().toLowerCase().includes(term.toLowerCase())
            )
          );
        });

        const departmentPcs = deptData.reduce((sum, row) => sum + parsePcsValue(row[pcsIndex]), 0);

        summaries['packing'] = {
          totalRows: deptData.length,
          totalPcs: departmentPcs || 0,
          styles: [...new Set(deptData.map(row => row[styleIndex]).filter(Boolean))].length,
          lastUpdated: deptData[0]?.[dateIndex] || 'N/A'
        };
        return;
      }

      // For any other departments, use search terms
      const deptData = indexData.filter(row => {
        return dept.searchTerms.some(term =>
          row.some(cell =>
            cell && cell.toString().toLowerCase().includes(term.toLowerCase())
          )
        );
      });

      const departmentPcs = deptData.reduce((sum, row) => sum + parsePcsValue(row[pcsIndex]), 0);

      summaries[dept.id] = {
        totalRows: deptData.length,
        totalPcs: departmentPcs || 0,
        styles: [...new Set(deptData.map(row => row[styleIndex]).filter(Boolean))].length,
        lastUpdated: deptData[0]?.[dateIndex] || 'N/A'
      };
    });

    return summaries;
  };

  // Calculate filtered department summaries
  const calculateFilteredDepartmentSummaries = () => {
    // Apply filters to Index data
    const filteredIndexData = applyFilters(sheetData, 'index');

    // Apply filters to JobOrder data
    const filteredJobData = applyFilters(jobOrderData, 'joborder');

    // Update filtered supervisor summary with PCS quantities
    const filteredSupSummary = calculateSupervisorSummary(filteredIndexData, headers, displayColumnIndices);
    setFilteredSupervisorSummary(filteredSupSummary);

    // Recalculate colour pending lots with filters
    const filteredColourPendingLots = [];
    const filteredColourPendingDetails = {};

    if (cuttingData.length && indexData.length && filteredJobData.length) {
      // Create a map of index data by lot number
      const indexMap = new Map();
      indexData.forEach((entry) => {
        if (entry && entry.lot) {
          indexMap.set(entry.lot, entry);
        }
      });

      // Get lot numbers from filtered JobOrder data
      const lotNumberIndexInJobOrder = findColumnIndex(jobOrderHeaders, ['lot number', 'lot no']);

      if (lotNumberIndexInJobOrder !== -1) {
        // For each filtered JobOrder row, check if it has pending shades
        filteredJobData.forEach((row) => {
          const lotNumber = row[lotNumberIndexInJobOrder];
          if (!lotNumber || isEmptyCell(lotNumber)) return;

          const lotStr = lotNumber.toString().trim();
          const indexEntry = indexMap.get(lotStr);

          // Skip if no index entry (fabric pending lots)
          if (!indexEntry) return;

          // Get cutting window for this lot
          const window = sliceCuttingMatrix(cuttingData, indexEntry.startRow, indexEntry.numRows);

          // Compute pending shades
          const pendingShadeKeys = computePendingShades(window, indexEntry.sizes, indexEntry.shades);

          if (pendingShadeKeys.size > 0) {
            // Map shade keys back to original shade names
            const shadeKeyToOriginal = new Map((indexEntry.shades || []).map((sh) => [normalizeKey(sh), sh]));
            const pendingList = Array.from(pendingShadeKeys).map(
              (k) => shadeKeyToOriginal.get(k) || k
            );

            // Find the matching row in sheetData (Index sheet) for display
            const lotNumberIndexInIndex = displayColumnIndices.lotNumber;
            let indexRow = null;

            if (lotNumberIndexInIndex !== -1) {
              indexRow = sheetData.find(r => {
                const idxLot = r[lotNumberIndexInIndex];
                return idxLot && idxLot.toString().trim() === lotStr;
              });
            }

            // If we found a matching index row, use that for display
            if (indexRow) {
              filteredColourPendingLots.push(indexRow);
            } else {
              // Create a minimal row with just the lot number
              const minimalRow = [...(new Array(headers.length).fill(''))];
              if (lotNumberIndexInIndex !== -1) {
                minimalRow[lotNumberIndexInIndex] = lotStr;
              }
              filteredColourPendingLots.push(minimalRow);
            }

            filteredColourPendingDetails[lotStr] = {
              lotNumber: lotStr,
              pendingShades: pendingList,
              totalPending: pendingList.length,
              fabric: indexEntry.fabric,
              garmentType: indexEntry.garmentType,
              style: indexEntry.style,
              cuttingDate: formatSavedAtToYMD(indexEntry.savedAt)
            };
          }
        });
      }
    }

    // Update state with filtered colour pending data
    setColourPendingLots(filteredColourPendingLots);
    setColourPendingDetails(filteredColourPendingDetails);

    // Recalculate pending issue after emb/print lots with filters
    const filteredPendingIssueLots = getPendingIssueAfterEmbPrintLots();
    const { lots: enrichedLots, details: enrichedDetails } = enrichPendingIssueLotsWithDetails(filteredPendingIssueLots);
    setPendingIssueLots(enrichedLots);
    setPendingIssueDetails(enrichedDetails);

    const summaries = calculateDepartmentSummaries(
      filteredIndexData,
      headers,
      displayColumnIndices,
      filteredJobData,
      jobOrderHeaders,
      cuttingData,
      indexData,
      filteredColourPendingLots,
      filteredColourPendingDetails,
      enrichedLots,
      enrichedDetails
    );

    setFilteredDepartmentSummary(summaries);
  };

  // Apply filters to data
  const applyFilters = (data, dataType = 'index') => {
    if (!data || data.length === 0) return [];

    let filteredData = [...data];

    // Use different headers based on data type
    const headersArray = dataType === 'joborder' ? jobOrderHeaders : headers;

    // Find column indices for filters based on data type
    let partyIndex = -1;
    let brandIndex = -1;
    let itemIndex = -1;
    let seasonIndex = -1;

    // For JobOrder, try multiple possible column names for party
    if (dataType === 'joborder') {
      // Try multiple possible column names for party in JobOrder
      const partySearchTerms = ['party name', 'party', 'client', 'customer', 'buyer', 'name'];
      partyIndex = findColumnIndex(headersArray, partySearchTerms);
    } else {
      partyIndex = findColumnIndex(headersArray, ['party', 'client', 'customer', 'buyer']);
    }

    brandIndex = findColumnIndex(headersArray, ['brand']);
    seasonIndex = findColumnIndex(headersArray, ['season']);

    if (dataType === 'joborder') {
      // For JobOrder, garment type might be in different columns
      itemIndex = findColumnIndex(headersArray, ['garment type', 'item', 'product', 'garment', 'style', 'description']);
    } else {
      itemIndex = findColumnIndex(headersArray, ['item', 'product', 'garment', 'style', 'garment type']);
    }

    // Apply party filter
    if (filters.party !== 'all' && partyIndex !== -1) {
      filteredData = filteredData.filter(row => {
        const rowPartyValue = row[partyIndex];
        if (!rowPartyValue || isEmptyCell(rowPartyValue)) {
          return false;
        }

        const rowParty = rowPartyValue.toString().trim();
        const filterParty = filters.party.toString().trim();
        return rowParty === filterParty;
      });
    }

    // Apply brand filter
    if (filters.brand !== 'all' && brandIndex !== -1) {
      filteredData = filteredData.filter(row => {
        const rowBrandValue = row[brandIndex];
        if (!rowBrandValue || isEmptyCell(rowBrandValue)) {
          return false;
        }

        const rowBrand = rowBrandValue.toString().trim();
        const filterBrand = filters.brand.toString().trim();
        return rowBrand === filterBrand;
      });
    }

    // Apply item filter
    if (filters.item !== 'all' && itemIndex !== -1) {
      filteredData = filteredData.filter(row => {
        const rowItemValue = row[itemIndex];
        if (!rowItemValue || isEmptyCell(rowItemValue)) {
          return false;
        }

        const rowItem = rowItemValue.toString().trim();
        const filterItem = filters.item.toString().trim();
        return rowItem === filterItem;
      });
    }

    // Apply season filter
    if (filters.season !== 'all' && seasonIndex !== -1) {
      filteredData = filteredData.filter(row => {
        const rowSeasonValue = row[seasonIndex];
        if (!rowSeasonValue || isEmptyCell(rowSeasonValue)) {
          return false;
        }

        const rowSeason = rowSeasonValue.toString().trim();
        const filterSeason = filters.season.toString().trim();
        return rowSeason === filterSeason;
      });
    }

    return filteredData;
  };

  const getDepartmentData = (departmentId) => {
    // For cutting_done, return ALL Index data
    if (departmentId === 'cutting_done') {
      return applyFilters(sheetData, 'index');
    }

    if (departmentId === 'fabric_pending') {
      // Return JobOrder lots that are missing in Index sheet
      let missingLots = getMissingLotsFromJobOrder();

      // Apply filters to the missing lots
      if (filters.party !== 'all' || filters.brand !== 'all' || filters.item !== 'all') {
        missingLots = applyFilters(missingLots, 'joborder');
      }

      return missingLots;
    }

    if (departmentId === 'colour_pending') {
      // Return colour pending lots
      let colourPendingData = colourPendingLots;

      // Apply filters
      if (filters.party !== 'all' || filters.brand !== 'all' || filters.item !== 'all') {
        colourPendingData = applyFilters(colourPendingData, 'index');
      }

      return colourPendingData;
    }

    if (departmentId === 'pending_after_emb_print') {
      // Return pending issue after emb/print lots
      let pendingData = pendingIssueLots;

      // Apply filters
      if (filters.party !== 'all' || filters.brand !== 'all' || filters.item !== 'all') {
        pendingData = applyFilters(pendingData, 'index');
      }

      return pendingData;
    }

    let data = applyFilters(sheetData);

    if (departmentId === 'all') {
      return data;
    }

    const department = departments.find(d => d.id === departmentId);
    if (!department) return data;

    if (departmentId === 'direct_stitching') {
      const directStitchingIndex = displayColumnIndices.directStitching;
      if (directStitchingIndex !== -1) {
        data = data.filter(row => {
          const dsValue = row[directStitchingIndex];
          return dsValue && dsValue.toString().toLowerCase().includes('yes');
        });
      }
    }
    else if (departmentId === 'direct_pending_issue') {
      const headersArray = headers;
      const dateOfIssueIndex = findColumnIndex(headersArray, ['date of issue']);
      const supervisorIndex = findColumnIndex(headersArray, ['supervisor']);
      const directStitchingIndex = displayColumnIndices.directStitching;

      if (directStitchingIndex !== -1 && dateOfIssueIndex !== -1 && supervisorIndex !== -1) {
        data = data.filter(row => {
          const dsValue = row[directStitchingIndex];
          const isDirectStitching = dsValue && dsValue.toString().toLowerCase().includes('yes');

          const dateOfIssueValue = row[dateOfIssueIndex];
          const isDateEmpty = isEmptyCell(dateOfIssueValue);

          const supervisorValue = row[supervisorIndex];
          const isSupervisorEmpty = isEmptyCell(supervisorValue);

          return isDirectStitching && isDateEmpty && isSupervisorEmpty;
        });
      }
    }
    else if (departmentId === 'stitching' && department.conditions) {
      const headersArray = headers;
      const completeStatusIndex = displayColumnIndices.completeStatus;
      const dateOfIssueIndex = findColumnIndex(headersArray, ['date of issue']);
      const supervisorIndex = findColumnIndex(headersArray, ['supervisor']);

      data = data.filter(row => {
        const dateOfIssueValue = row[dateOfIssueIndex];
        const supervisorValue = row[supervisorIndex];
        const completeStatusValue = row[completeStatusIndex];

        const isIssued = !isEmptyCell(dateOfIssueValue) && !isEmptyCell(supervisorValue);
        const isNotCompleted = !isLotCompleted(completeStatusValue);

        return isIssued && isNotCompleted;
      });
    }
    else if (departmentId === 'pending_stitching' && department.conditions) {
      const headersArray = headers;
      const supervisorIndex = findColumnIndex(headersArray, ['supervisor']);

      data = data.filter(row => {
        const supervisorValue = row[supervisorIndex];

        const isSupervisorEmpty = isEmptyCell(supervisorValue);

        const hasDisplayColumnData = Object.values(displayColumnIndices).some(index => {
          if (index === -1) return false;
          const cellValue = row[index];
          return !isEmptyCell(cellValue);
        });

        const hasAnyData = row.some((cell, idx) => {
          if (idx === supervisorIndex) return false;
          return !isEmptyCell(cell);
        });

        return isSupervisorEmpty && (hasDisplayColumnData || hasAnyData);
      });
    }
    else if (departmentId === 'embroidery' && department.conditions) {
      const challanHistoryIndex = findColumnIndex(headers, ['challan history']);

      data = data.filter(row => {
        const challanHistory = row[challanHistoryIndex];
        return isEmbroideryLotPending(challanHistory);
      });
    }
    else if (departmentId === 'printing' && department.conditions) {
      const challanHistoryIndex = findColumnIndex(headers, ['challan history']);

      data = data.filter(row => {
        const challanHistory = row[challanHistoryIndex];
        return isPrintingLotPending(challanHistory);
      });
    }
    else {
      data = data.filter(row => {
        return department.searchTerms.some(term =>
          row.some(cell =>
            cell && cell.toString().toLowerCase().includes(term.toLowerCase())
          )
        );
      });
    }

    return data;
  };

  // Handle filter changes
  const handleFilterChange = (filterType, value) => {
    setFilters(prev => ({
      ...prev,
      [filterType]: value
    }));
  };

  // Reset all filters
  const resetFilters = () => {
    setFilters({
      party: 'all',
      brand: 'all',
      item: 'all',
      season: 'all'
    });
  };

  const refreshData = () => {
    fetchData();
  };

  // Password modal handlers
  const handlePasswordSubmit = () => {
    if (passwordInput === MD_PASSWORD) {
      setDownloadMode('md');
      setPasswordError('');
      setShowPasswordModal(false);
      setPasswordInput('');

      // If there's pending PDF data, generate it with MD permissions
      if (pendingPdfData) {
        setTimeout(() => {
          generatePDFWithMode('md');
          setPendingPdfData(null);
        }, 100);
      }
    } else if (passwordInput === CLIENT_PASSWORD) {
      setDownloadMode('client');
      setPasswordError('');
      setShowPasswordModal(false);
      setPasswordInput('');

      // If there's pending PDF data, generate it with Client permissions
      if (pendingPdfData) {
        setTimeout(() => {
          generatePDFWithMode('client');
          setPendingPdfData(null);
        }, 100);
      }
    } else {
      setPasswordError('Incorrect password. Please try again.');
    }
  };

  const handlePasswordCancel = () => {
    setShowPasswordModal(false);
    setPasswordInput('');
    setPasswordError('');
    setPendingPdfData(null);
  };

  const openPasswordModal = () => {
    setShowPasswordModal(true);
    setPasswordInput('');
    setPasswordError('');
    setPendingPdfData(new Date()); // Just a trigger to indicate PDF is pending
  };

  const generatePDFWithMode = (mode) => {
    // Create PDF with professional settings
    const doc = new jsPDF({
      orientation: 'landscape',
      unit: 'mm',
      format: 'a4'
    });

    const pageWidth = doc.internal.pageSize.width;
    const pageHeight = doc.internal.pageSize.height;
    const margin = 15;
    const contentWidth = pageWidth - (margin * 2);

    // ===== PROFESSIONAL COLOR PALETTE =====
    const colors = {
      primary: [52, 73, 94],        // Professional Blue
      secondary: [52, 73, 94],         // Dark Blue-Gray
      accent: [192, 57, 43],           // Professional Red
      light: [236, 240, 241],          // Light Gray
      border: [0, 0, 0],               // Black border
      text: [44, 62, 80],              // Dark Gray
      textLight: [127, 140, 141],      // Medium Gray
      success: [39, 174, 96],           // Green
      warning: [230, 126, 34],          // Orange
      white: [255, 255, 255],
      black: [0, 0, 0]
    };

    // ===== HELPER FUNCTIONS =====
    const addHeader = (pageNum) => {
      // Clean header with solid color bar
      doc.setFillColor(colors.primary[0], colors.primary[1], colors.primary[2]);
      doc.rect(0, 0, pageWidth, 12, 'F');

      doc.setTextColor(colors.white[0], colors.white[1], colors.white[2]);
      doc.setFontSize(14);
      doc.setFont('times', 'bold');
      doc.text('PRODUCTION SUMMARY REPORT', margin, 9);

      doc.setFontSize(8);
      doc.setFont('times', 'normal');
      const dateStr = `Generated: ${new Date().toLocaleString()}`;
      doc.text(dateStr, pageWidth - margin - 80, 9);
      doc.text(`Page ${pageNum}`, pageWidth - margin - 15, 9);

      // Mode indicator
      doc.setFontSize(8);
      doc.setFont('times', 'bold');
      if (mode === 'md') {
        doc.text('MANAGEMENT VIEW', pageWidth / 2 - 30, 9);
      } else {
        doc.text('CLIENT VIEW', pageWidth / 2 - 25, 9);
      }
    };

    const addFooter = (pageNum, totalPages) => {
      const footerY = pageHeight - 8;

      doc.setDrawColor(colors.border[0], colors.border[1], colors.border[2]);
      doc.setLineWidth(0.3);
      doc.line(margin, footerY - 4, pageWidth - margin, footerY - 4);

      doc.setFontSize(7);
      doc.setTextColor(colors.textLight[0], colors.textLight[1], colors.textLight[2]);
      doc.text('Production Management System', margin, footerY);
      doc.text(`Page ${pageNum} of ${totalPages}`, pageWidth / 2, footerY, { align: 'center' });
      doc.text(new Date().toLocaleDateString(), pageWidth - margin, footerY, { align: 'right' });
    };

    // Function to format quantity based on mode
    const formatQuantity = (value, hideForClient = true, forceHide = false) => {
      // If forceHide is true, always show asterisks
      if (forceHide) {
        return '***';
      }

      // For client mode, hide quantities
      if (mode === 'client') {
        return '***';
      }

      // For MD mode, show actual values
      if (value === undefined || value === null) return '0';
      if (typeof value === 'number') return value.toLocaleString();
      return value.toString();
    };

    // Start with header
    addHeader(1);

    // ===== TITLE AND REPORT INFO =====
    let yPos = 22;

    // Report Title
    doc.setTextColor(colors.secondary[0], colors.secondary[1], colors.secondary[2]);
    doc.setFontSize(16);
    doc.setFont('times', 'bold');
    doc.text('Production Dashboard Summary', margin, yPos);

    yPos += 5;
    doc.setFontSize(9);
    doc.setFont('times', 'normal');
    doc.setTextColor(colors.textLight[0], colors.textLight[1], colors.textLight[2]);
    doc.text('Comprehensive overview of production departments and supervisor performance', margin, yPos);

    yPos += 10;

    // ===== FILTER SECTION =====
    doc.setDrawColor(colors.border[0], colors.border[1], colors.border[2]);
    doc.setFillColor(colors.light[0], colors.light[1], colors.light[2]);
    doc.roundedRect(margin, yPos, contentWidth, 12, 1.5, 1.5, 'FD');

    doc.setFontSize(8);
    doc.setFont('times', 'bold');
    doc.setTextColor(colors.secondary[0], colors.secondary[1], colors.secondary[2]);
    doc.text('ACTIVE FILTERS:', margin + 4, yPos + 7);

    doc.setFont('times', 'normal');
    doc.setTextColor(colors.text[0], colors.text[1], colors.text[2]);

    const activeFilters = [];
    if (filters.party !== 'all') activeFilters.push(`Party: ${filters.party}`);
    if (filters.brand !== 'all') activeFilters.push(`Brand: ${filters.brand}`);
    if (filters.item !== 'all') activeFilters.push(`Item: ${filters.item}`);
    if (filters.season !== 'all') activeFilters.push(`Season: ${filters.season}`);

    const filterText = activeFilters.length > 0 ? activeFilters.join('  |  ') : 'No filters applied';
    doc.text(filterText, margin + 45, yPos + 7);

    yPos += 20;

    // ===== QUICK SUMMARY STATISTICS =====
    // Calculate summary statistics
    const supSummary = filteredSupervisorSummary;
    const supArray = Object.keys(supSummary).sort();

    const totalLots = supArray.length > 0 ? Object.values(supSummary).reduce((sum, s) => sum + s.totalLots, 0) : 0;
    const totalPendingLots = supArray.length > 0 ? Object.values(supSummary).reduce((sum, s) => sum + s.pendingLots, 0) : 0;
    const totalCompleted = supArray.length > 0 ? Object.values(supSummary).reduce((sum, s) => sum + s.completedLots, 0) : 0;
    const totalPcs = supArray.length > 0 ? Object.values(supSummary).reduce((sum, s) => sum + s.totalPcs, 0) : 0;
    const totalPendingPcs = supArray.length > 0 ? Object.values(supSummary).reduce((sum, s) => sum + s.pendingPcs, 0) : 0;
    const overallCompletion = totalLots > 0 ? (totalCompleted / totalLots * 100).toFixed(1) : '0';

    // Summary statistics box
    doc.setDrawColor(colors.primary[0], colors.primary[1], colors.primary[2]);
    doc.setFillColor(colors.white[0], colors.white[1], colors.white[2]);
    doc.roundedRect(margin, yPos, contentWidth, 24, 1.5, 1.5, 'FD');

    doc.setFontSize(10);
    doc.setFont('times', 'bold');
    doc.setTextColor(colors.primary[0], colors.primary[1], colors.primary[2]);
    doc.text('STITCHING SUMMARY STATISTICS', margin + 4, yPos + 7);

    doc.setFont('times', 'normal');
    doc.setTextColor(colors.text[0], colors.text[1], colors.text[2]);
    doc.setFontSize(9.5);

    // Show/hide quantities based on mode
    if (mode === 'client') {
      // Client view - hide all quantities
      doc.text(
        `Supervisors: ${supArray.length} | Total Lots: ${totalLots} | Pending Lots: ${totalPendingLots} | Completion Rate: ${overallCompletion}%`,
        margin + 4,
        yPos + 13
      );
      doc.text(
        `Total Pcs: *** | Pending Pcs: ***`,
        margin + 4,
        yPos + 19
      );
    } else {
      // MD view - show all quantities
      doc.text(
        `Supervisors: ${supArray.length} | Total Lots: ${totalLots} | Pending Lots: ${totalPendingLots} | Completion Rate: ${overallCompletion}%`,
        margin + 4,
        yPos + 13
      );
      doc.text(
        `Total Pcs: ${totalPcs.toLocaleString()} | Pending Pcs: ${totalPendingPcs.toLocaleString()}`,
        margin + 4,
        yPos + 19
      );
    }

    yPos += 32;

    // ===== DEPARTMENT SUMMARY TABLE =====
    doc.setFontSize(14);
    doc.setFont('times', 'bold');
    doc.setTextColor(colors.secondary[0], colors.secondary[1], colors.secondary[2]);
    doc.text('Department Summary', margin, yPos);
    yPos += 8;

    // Prepare department table data
    const departmentTableData = departments.map(dept => {
      const summary = filteredDepartmentSummary[dept.id];
      if (!summary) return null;

      let status = '';
      let statusColor = colors.text;

      if (dept.id === 'fabric_pending') {
        status = summary.totalRows > 0 ? 'Pending' : 'Complete';
        statusColor = summary.totalRows > 0 ? colors.warning : colors.success;
      } else if (dept.id === 'embroidery' || dept.id === 'printing') {
        status = summary.pendingLots > 0 ? 'In Progress' : 'Complete';
        statusColor = summary.pendingLots > 0 ? colors.warning : colors.success;
      } else if (dept.id === 'colour_pending') {
        status = summary.totalRows > 0 ? 'Pending' : 'Complete';
        statusColor = summary.totalRows > 0 ? colors.warning : colors.success;
      } else if (dept.id === 'pending_after_emb_print') {
        status = summary.totalRows > 0 ? 'Ready' : 'No Data';
        statusColor = summary.totalRows > 0 ? colors.success : colors.textLight;
      } else if (summary.totalRows > 0) {
        status = 'Active';
        statusColor = colors.success;
      } else {
        status = 'No Data';
        statusColor = colors.textLight;
      }

      // Determine if quantity should be hidden for this department
      // Hide for 'all' and 'cutting_done' departments in MD view
      // Hide for all departments in Client view
      const hideQuantity = (mode === 'client') || (mode === 'md' && (dept.id === 'all' || dept.id === 'cutting_done'));

      // Additional info based on department
      let additionalInfo = '';
      if (dept.id === 'pending_after_emb_print' && summary) {
        additionalInfo = `DS:${summary.directStitchingCount || 0} | EP:${summary.embPrintCompletedCount || 0}`;
      }

      return {
        data: [
          dept.name,
          summary.totalRows.toString(),
          summary.styles.toString(),
          hideQuantity ? '***' : formatQuantity(summary.totalPcs, false), // Don't hide in formatQuantity, we handle it here
          status,
          additionalInfo
        ],
        statusColor
      };
    }).filter(row => row !== null);

    // Department table with increased font size and centered content - UPDATED
    autoTable(doc, {  // Added doc as first parameter
      head: [['Department', 'Records', 'Styles', 'Quantity (PCS)', 'Status', 'Any Remarks']],
      body: departmentTableData.map(item => item.data),
      startY: yPos,
      theme: 'grid',
      styles: {
        fontSize: 10,
        cellPadding: 2,
        lineColor: [0, 0, 0],
        lineWidth: 0.5,
        textColor: [colors.text[0], colors.text[1], colors.text[2]],
        halign: 'center',
        valign: 'middle',
        fontStyle: 'normal'
      },
      headStyles: {
        fillColor: [colors.secondary[0], colors.secondary[1], colors.secondary[2]],
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        fontSize: 11,
        halign: 'center',
        valign: 'middle',
        cellPadding: 2,
        lineColor: [0, 0, 0],
        lineWidth: 0.5
      },
      columnStyles: {
        0: { cellWidth: contentWidth * 0.25, halign: 'center' },
        1: { cellWidth: contentWidth * 0.12, halign: 'center' },
        2: { cellWidth: contentWidth * 0.12, halign: 'center' },
        3: { cellWidth: contentWidth * 0.15, halign: 'center' },
        4: { cellWidth: contentWidth * 0.15, halign: 'center' },
        5: { cellWidth: contentWidth * 0.21, halign: 'center' }
      },
      margin: { left: margin, right: margin },
      didParseCell: function (data) {
        // Color code status cells
        if (data.section === 'body' && data.column.index === 4) {
          const statusData = departmentTableData[data.row.index];
          if (statusData && statusData.statusColor) {
            data.cell.styles.textColor = statusData.statusColor;
            data.cell.styles.fontStyle = 'bold';
          }
        }
      }
    });

    yPos = doc.lastAutoTable.finalY + 15;

    // ===== SUPERVISOR PERFORMANCE SECTION =====
    // Check if we need a new page
    if (yPos > pageHeight - 70) {
      doc.addPage();
      addHeader(doc.internal.getCurrentPageInfo().pageNumber);
      yPos = 22;
    }

    doc.setFontSize(14);
    doc.setFont('times', 'bold');
    doc.setTextColor(colors.secondary[0], colors.secondary[1], colors.secondary[2]);
    doc.text('Supervisor Performance', margin, yPos);
    yPos += 8;

    if (supArray.length > 0) {
      // Prepare supervisor table data
      const supervisorTableData = supArray.map(supName => {
        const data = supSummary[supName];
        const completionRate = data.totalLots > 0 ? ((data.completedLots / data.totalLots) * 100).toFixed(1) : '0';

        // In MD view, show actual quantities. In Client view, hide them
        const totalPcsDisplay = mode === 'client' ? '***' : data.totalPcs.toLocaleString();
        const pendingPcsDisplay = mode === 'client' ? '***' : data.pendingPcs.toLocaleString();

        return [
          supName,
          data.totalLots.toString(),
          data.pendingLots.toString(),
          totalPcsDisplay,
          pendingPcsDisplay,
          `${completionRate}%`
        ];
      });

      // Supervisor table with increased font size, centered content, and full width - UPDATED
      autoTable(doc, {  // Added doc as first parameter
        head: [['Supervisor', 'Total Lots', 'Pending', 'Total Pcs', 'Pending Pcs', 'Complete %']],
        body: supervisorTableData,
        startY: yPos,
        theme: 'grid',
        styles: {
          fontSize: 10,
          cellPadding: 2,
          lineColor: [0, 0, 0],
          lineWidth: 0.5,
          textColor: [colors.text[0], colors.text[1], colors.text[2]],
          halign: 'center',
          valign: 'middle',
          fontStyle: 'normal'
        },
        headStyles: {
          fillColor: [colors.secondary[0], colors.secondary[1], colors.secondary[2]],
          textColor: [255, 255, 255],
          fontStyle: 'bold',
          fontSize: 11,
          halign: 'center',
          valign: 'middle',
          cellPadding: 2,
          lineColor: [0, 0, 0],
          lineWidth: 0.5
        },
        columnStyles: {
          0: { cellWidth: contentWidth * 0.25, halign: 'center' },
          1: { cellWidth: contentWidth * 0.12, halign: 'center' },
          2: { cellWidth: contentWidth * 0.12, halign: 'center' },
          3: { cellWidth: contentWidth * 0.18, halign: 'center' },
          4: { cellWidth: contentWidth * 0.18, halign: 'center' },
          5: { cellWidth: contentWidth * 0.15, halign: 'center' }
        },
        margin: { left: margin, right: margin },
        didParseCell: function (data) {
          // Color code completion percentages
          if (data.section === 'body' && data.column.index === 5) {
            const rate = parseFloat(data.cell.raw);
            if (rate >= 75) data.cell.styles.textColor = colors.success;
            else if (rate >= 50) data.cell.styles.textColor = colors.warning;
            else if (rate > 0) data.cell.styles.textColor = colors.accent;
            data.cell.styles.fontStyle = 'bold';
          }
        }
      });

      yPos = doc.lastAutoTable.finalY + 15;
    }

    // ===== COLOUR PENDING SECTION =====
    if (colourPendingLots && colourPendingLots.length > 0) {
      // Check if we need a new page
      if (yPos > pageHeight - 60) {
        doc.addPage();
        addHeader(doc.internal.getCurrentPageInfo().pageNumber);
        yPos = 22;
      }

      doc.setFontSize(14);
      doc.setFont('times', 'bold');
      doc.setTextColor(colors.secondary[0], colors.secondary[1], colors.secondary[2]);
      doc.text('Colour Pending Analysis', margin, yPos);
      yPos += 8;

      // Prepare pending shades data - show ALL shades without truncation
      const pendingShadesData = [];

      colourPendingLots.slice(0, 25).forEach((row) => {
        const lotNumberIndex = displayColumnIndices.lotNumber;
        const lotNumber = lotNumberIndex !== -1 ? row[lotNumberIndex] : 'Unknown';
        const lotStr = lotNumber?.toString().trim() || 'Unknown';
        const details = colourPendingDetails[lotStr];

        if (details && details.pendingShades && details.pendingShades.length > 0) {
          // Join ALL shades with commas - no truncation
          const allShades = details.pendingShades.join(', ');

          pendingShadesData.push([
            lotStr,
            details.pendingShades.length.toString(),
            allShades // Display all shades without truncation
          ]);
        }
      });

      if (pendingShadesData.length > 0) {
        // Calculate available height for the table
        const remainingHeight = pageHeight - yPos - 40; // Leave space for footer
        const estimatedRowHeight = 8; // Approximate height per row including padding
        const maxRows = Math.floor(remainingHeight / estimatedRowHeight);

        // Limit rows if needed to fit on page
        const displayData = pendingShadesData.slice(0, maxRows);
        const remainingCount = pendingShadesData.length - maxRows;

        // Colour Pending Analysis table - UPDATED
        autoTable(doc, {  // Added doc as first parameter
          head: [['Lot Number', 'Pending Shades', 'Shade Details']],
          body: displayData,
          startY: yPos,
          theme: 'grid',
          styles: {
            fontSize: 9,
            cellPadding: 2,
            lineColor: [0, 0, 0],
            lineWidth: 0.5,
            textColor: [colors.text[0], colors.text[1], colors.text[2]],
            halign: 'left',
            valign: 'middle',
            overflow: 'linebreak'
          },
          headStyles: {
            fillColor: [colors.secondary[0], colors.secondary[1], colors.secondary[2]],
            textColor: [255, 255, 255],
            fontStyle: 'bold',
            fontSize: 10,
            halign: 'center',
            valign: 'middle',
            cellPadding: 2,
            lineColor: [0, 0, 0],
            lineWidth: 0.5
          },
          columnStyles: {
            0: { cellWidth: contentWidth * 0.2, halign: 'center' },
            1: { cellWidth: contentWidth * 0.15, halign: 'center' },
            2: { cellWidth: contentWidth * 0.65, halign: 'left' }
          },
          margin: { left: margin, right: margin },
          didParseCell: function (data) {
            if (data.section === 'body' && data.column.index === 1) {
              const count = parseInt(data.cell.raw);
              if (count > 5) {
                data.cell.styles.textColor = colors.warning;
                data.cell.styles.fontStyle = 'bold';
              }
            }
            if (data.section === 'body' && data.column.index === 2) {
              data.cell.styles.overflow = 'linebreak';
              data.cell.styles.cellWidth = 'wrap';
            }
          }
        });

        // Calculate total pending shades only (no quantity)
        const totalPendingShades = colourPendingLots.reduce((total, row) => {
          const lotNumberIndex = displayColumnIndices.lotNumber;
          const lotNumber = lotNumberIndex !== -1 ? row[lotNumberIndex] : 'Unknown';
          const lotStr = lotNumber?.toString().trim() || 'Unknown';
          const details = colourPendingDetails[lotStr];
          return total + (details?.totalPending || 0);
        }, 0);

        // Summary with information about remaining lots if truncated
        const summaryY = doc.lastAutoTable.finalY + 5;

        doc.setFillColor(colors.light[0], colors.light[1], colors.light[2]);

        const summaryHeight = remainingCount > 0 ? 18 : 10;
        doc.roundedRect(margin, summaryY, contentWidth, summaryHeight, 1, 1, 'F');

        doc.setFontSize(9);
        doc.setFont('times', 'normal');
        doc.setTextColor(colors.text[0], colors.text[1], colors.text[2]);

        doc.text(
          `Lots with pending colours: ${colourPendingLots.length} | Total pending shades: ${totalPendingShades} | Average per lot: ${(totalPendingShades / colourPendingLots.length).toFixed(1)}`,
          margin + 4,
          summaryY + 6
        );

        if (remainingCount > 0) {
          doc.setTextColor(colors.warning[0], colors.warning[1], colors.warning[2]);
          doc.setFont('times', 'bold');
          doc.text(
            `Note: ${remainingCount} more lot(s) with pending colours not shown due to space. Please see detailed report for complete list.`,
            margin + 4,
            summaryY + 13
          );
        }

        if (pendingShadesData.length > maxRows) {
          yPos = summaryY + summaryHeight + 5;

          const remainingPages = Math.ceil((pendingShadesData.length - maxRows) / maxRows);

          for (let page = 1; page < remainingPages; page++) {
            doc.addPage();
            addHeader(doc.internal.getCurrentPageInfo().pageNumber);

            const newYPos = 22;

            doc.setFontSize(14);
            doc.setFont('times', 'bold');
            doc.setTextColor(colors.secondary[0], colors.secondary[1], colors.secondary[2]);
            doc.text('Colour Pending Analysis (Continued)', margin, newYPos);

            const nextBatchStart = maxRows + (page * maxRows);
            const nextBatchEnd = Math.min(nextBatchStart + maxRows, pendingShadesData.length);
            const nextBatchData = pendingShadesData.slice(maxRows + ((page - 1) * maxRows), nextBatchEnd);

            // Colour Pending Analysis (Continued) table - UPDATED
            autoTable(doc, {  // Added doc as first parameter
              head: [['Lot Number', 'Pending Shades', 'Shade Details']],
              body: nextBatchData,
              startY: newYPos + 8,
              theme: 'grid',
              styles: {
                fontSize: 9,
                cellPadding: 2,
                lineColor: [0, 0, 0],
                lineWidth: 0.5,
                textColor: [colors.text[0], colors.text[1], colors.text[2]],
                halign: 'left',
                valign: 'middle',
                overflow: 'linebreak'
              },
              headStyles: {
                fillColor: [colors.secondary[0], colors.secondary[1], colors.secondary[2]],
                textColor: [255, 255, 255],
                fontStyle: 'bold',
                fontSize: 10,
                halign: 'center',
                valign: 'middle',
                cellPadding: 2,
                lineColor: [0, 0, 0],
                lineWidth: 0.5
              },
              columnStyles: {
                0: { cellWidth: contentWidth * 0.2, halign: 'center' },
                1: { cellWidth: contentWidth * 0.15, halign: 'center' },
                2: { cellWidth: contentWidth * 0.65, halign: 'left' }
              },
              margin: { left: margin, right: margin }
            });
          }
        }
      } else {
        doc.setFontSize(10);
        doc.setFont('times', 'italic');
        doc.setTextColor(colors.textLight[0], colors.textLight[1], colors.textLight[2]);
        doc.text('No pending shades found', margin, yPos + 10);
      }
    }

    // ===== FOOTER ON ALL PAGES =====
    const pageCount = doc.internal.getNumberOfPages();
    for (let i = 1; i <= pageCount; i++) {
      doc.setPage(i);
      addFooter(i, pageCount);
    }

    // Save PDF with mode indicator in filename
    const today = new Date();
    const modeSuffix = mode === 'md' ? 'MD' : 'Client';
    const fileName = `Production_Report_${modeSuffix}_${today.getFullYear()}-${(today.getMonth() + 1).toString().padStart(2, '0')}-${today.getDate().toString().padStart(2, '0')}.pdf`;
    doc.save(fileName);
  };
  // Export Summary to PDF with password protection
  const exportSummaryToPDF = () => {
    openPasswordModal();
  };

  if (loading) {
    return (
      <div className="short-summary-report-container">
        <div className="short-summary-report-loading">
          <div className="short-summary-report-spinner"></div>
          <p style={{ color: '#64748b', fontSize: '16px' }}>Loading data from Google Sheets...</p>
          <p style={{ color: '#94a3b8', fontSize: '14px', marginTop: '8px' }}>Fetching data from multiple spreadsheets</p>
        </div>
      </div>
    );
  }

  if (error && sheetData.length === 0 && jobOrderData.length === 0) {
    return (
      <div className="short-summary-report-container">
        <div className="short-summary-report-error-container">
          <h3 className="short-summary-report-error-title">⚠️ Error Loading Data</h3>
          <p className="short-summary-report-error-message">{error}</p>
          <div style={{ display: 'flex', gap: '12px', justifyContent: 'center' }}>
            <button
              onClick={refreshData}
              className="short-summary-report-action-button short-summary-report-error-button"
            >
              ↻ Retry Loading
            </button>
          </div>
        </div>
      </div>
    );
  }

  const currentDept = departments.find(d => d.id === activeDepartment);

  return (
    <div className="short-summary-report-container">
      {/* Password Modal */}
      {showPasswordModal && (
        <div className="short-summary-report-modal-overlay">
          <div className="short-summary-report-modal">
            <div className="short-summary-report-modal-header">
              <h3 className="short-summary-report-modal-title">🔒 Password Protected Download</h3>
              <button className="short-summary-report-modal-close" onClick={handlePasswordCancel}>×</button>
            </div>
            <div className="short-summary-report-modal-body">
              <p className="short-summary-report-modal-text">
                Please enter the password to download the PDF report:
              </p>
              <p className="short-summary-report-modal-hint">
                <small>MD Password: For full data access (includes all quantities)</small><br />
                <small>Client Password: For limited data (quantities hidden)</small>
              </p>
              <input
                type="password"
                className="short-summary-report-modal-input"
                placeholder="Enter password"
                value={passwordInput}
                onChange={(e) => setPasswordInput(e.target.value)}
                onKeyPress={(e) => e.key === 'Enter' && handlePasswordSubmit()}
                autoFocus
              />
              {passwordError && (
                <p className="short-summary-report-modal-error">{passwordError}</p>
              )}
            </div>
            <div className="short-summary-report-modal-footer">
              <button className="short-summary-report-modal-button short-summary-report-modal-button-secondary" onClick={handlePasswordCancel}>
                Cancel
              </button>
              <button className="short-summary-report-modal-button short-summary-report-modal-button-primary" onClick={handlePasswordSubmit}>
                Download PDF
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Header */}
      <div className="short-summary-report-header">
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
          <div style={{ flex: 1 }}>
            <h2 className="short-summary-report-header-title">Summary Report of Production</h2>
            {/* Error warning if JobOrder failed but Index loaded */}
            {error && sheetData.length > 0 && (
              <div style={{ marginTop: '15px', padding: '12px', backgroundColor: '#fef3c7', borderRadius: '8px', border: '1px solid #fbbf24' }}>
                <span style={{ fontWeight: '600', color: '#92400e', fontSize: '14px' }}>⚠️ Partial Load: </span>
                <span style={{ color: '#92400e', fontSize: '13px' }}>
                  Index and Cutting sheets loaded successfully, but JobOrder sheet failed: {error}. Some data may not be accurate.
                </span>
              </div>
            )}

            {/* Filter Summary */}
            {(filters.party !== 'all' || filters.brand !== 'all' || filters.item !== 'all' || filters.season !== 'all') && (
              <div style={{ marginTop: '15px', padding: '12px', backgroundColor: '#f0f9ff', borderRadius: '8px', border: '1px solid #bae6fd' }}>
                <span style={{ fontWeight: '600', color: '#0369a1', fontSize: '14px' }}>🎯 Active Filters: </span>
                {filters.party !== 'all' && (
                  <span className="short-summary-report-filter-badge">👥 Party: {filters.party}</span>
                )}
                {filters.brand !== 'all' && (
                  <span className="short-summary-report-filter-badge">🏷️ Brand: {filters.brand}</span>
                )}
                {filters.item !== 'all' && (
                  <span className="short-summary-report-filter-badge">👕 Item: {filters.item}</span>
                )}
                {filters.season !== 'all' && (
                  <span className="short-summary-report-filter-badge">📅 Season: {filters.season}</span>
                )}
              </div>
            )}
          </div>
          {/* Back Button - Simple Version */}
          <div style={{ marginBottom: '20px' }}>
            <button
              onClick={() => window.history.back()}
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                gap: '8px',
                padding: '8px 16px',
                background: '#f8fafc',
                border: '1px solid #cbd5e1',
                borderRadius: '6px',
                color: '#334155',
                fontSize: '14px',
                fontWeight: 500,
                cursor: 'pointer'
              }}
              onMouseOver={(e) => {
                e.currentTarget.style.background = '#e2e8f0';
              }}
              onMouseOut={(e) => {
                e.currentTarget.style.background = '#f8fafc';
              }}
            >
              <span style={{ fontSize: '18px' }}>←</span>
              <span>Back</span>
            </button>
          </div>

          <div style={{ display: 'flex', gap: '10px' }}>
            <button
              onClick={refreshData}
              className="short-summary-report-action-button short-summary-report-refresh-button"
            >
              ↻ Refresh
            </button>
            <button
              onClick={exportSummaryToPDF}
              className="short-summary-report-action-button short-summary-report-pdf-button"
            >
              📄 Summary PDF
            </button>
          </div>
        </div>
      </div>

      {/* Filters Section */}
      <div className="short-summary-report-filter-section">
        <h3 className="short-summary-report-filter-title">
          <span>🔍 Filter Data</span>
        </h3>
        <div className="short-summary-report-filter-grid">
          <div className="short-summary-report-filter-group">
            <label className="short-summary-report-filter-label">👥 Party Wise</label>
            <select
              className="short-summary-report-filter-select"
              value={filters.party}
              onChange={(e) => handleFilterChange('party', e.target.value)}
            >
              {filterOptions.parties.map(party => (
                <option key={party} value={party}>
                  {party === 'all' ? '📊 All Parties' : party}
                </option>
              ))}
            </select>
          </div>

          <div className="short-summary-report-filter-group">
            <label className="short-summary-report-filter-label">🏷️ Brand Wise</label>
            <select
              className="short-summary-report-filter-select"
              value={filters.brand}
              onChange={(e) => handleFilterChange('brand', e.target.value)}
            >
              {filterOptions.brands.map(brand => (
                <option key={brand} value={brand}>
                  {brand === 'all' ? '📦 All Brands' : brand}
                </option>
              ))}
            </select>
          </div>

          <div className="short-summary-report-filter-group">
            <label className="short-summary-report-filter-label">👕 Item Wise</label>
            <select
              className="short-summary-report-filter-select"
              value={filters.item}
              onChange={(e) => handleFilterChange('item', e.target.value)}
            >
              {filterOptions.items.map(item => (
                <option key={item} value={item}>
                  {item === 'all' ? '👚 All Items' : item}
                </option>
              ))}
            </select>
          </div>

          {/* Season Filter */}
          <div className="short-summary-report-filter-group">
            <label className="short-summary-report-filter-label">📅 Season Wise</label>
            <select
              className="short-summary-report-filter-select"
              value={filters.season}
              onChange={(e) => handleFilterChange('season', e.target.value)}
            >
              {filterOptions.seasons.map(season => (
                <option key={season} value={season}>
                  {season === 'all' ? '🗓️ All Seasons' : season}
                </option>
              ))}
            </select>
          </div>
        </div>

        <div className="short-summary-report-filter-actions">
          <div className="short-summary-report-filter-stats">
            <div className="short-summary-report-filter-stat">
              <span className="short-summary-report-filter-stat-label">Filter Applied</span>
              <span className="short-summary-report-filter-stat-value">
                {[filters.party, filters.brand, filters.item, filters.season].filter(f => f !== 'all').length}
              </span>
            </div>
          </div>
          <button
            onClick={resetFilters}
            className="short-summary-report-reset-button"
          >
            🗑️ Clear Filters
          </button>
        </div>
      </div>

      {/* Department Tabs */}
      <div className="short-summary-report-department-tabs">
        {departments.map(dept => {
          const summary = filteredDepartmentSummary[dept.id];
          return (
            <button
              key={dept.id}
              className={`short-summary-report-dept-tab ${activeDepartment === dept.id ? 'short-summary-report-dept-tab-active' : ''}`}
              style={activeDepartment === dept.id ? {
                background: `linear-gradient(135deg, ${dept.color} 0%, ${dept.color}99 100%)`
              } : {}}
              onClick={() => setActiveDepartment(dept.id)}
              onMouseOver={e => {
                if (activeDepartment !== dept.id) {
                  e.currentTarget.style.backgroundColor = `${dept.color}15`;
                }
              }}
              onMouseOut={e => {
                if (activeDepartment !== dept.id) {
                  e.currentTarget.style.backgroundColor = '#f8fafc';
                }
              }}
            >
              <span className="short-summary-report-dept-icon">{dept.icon}</span>
              <span className="short-summary-report-dept-name">{dept.name}</span>
              {dept.subtitle && (
                <span className="short-summary-report-dept-subtitle">{dept.subtitle}</span>
              )}
              {summary && (
                <span className={`short-summary-report-dept-count ${activeDepartment === dept.id ? 'short-summary-report-dept-count-active' : ''}`}>
                  {summary.totalRows}
                </span>
              )}
            </button>
          );
        })}
      </div>

      {/* Department Summary Table */}
      <div ref={summaryRef} className="short-summary-report-summary-table">
        <div className="short-summary-report-summary-table-header">
          <span>📈 Department Summary</span>
          <div className="short-summary-report-export-buttons">
            <button
              onClick={exportSummaryToPDF}
              className="short-summary-report-action-button"
              style={{
                padding: '8px 16px',
                fontSize: '12px',
                backgroundColor: 'rgba(255,255,255,0.2)',
                backdropFilter: 'blur(10px)'
              }}
            >
              📄 Export PDF
            </button>
          </div>
        </div>
        <div className="short-summary-report-summary-table-body">
          <div className="short-summary-report-summary-table-row short-summary-report-summary-table-cell-header">
            <div className="short-summary-report-summary-table-cell short-summary-report-summary-table-cell-header">
              Department
            </div>
            <div className="short-summary-report-summary-table-cell short-summary-report-summary-table-cell-header">
              Total Records
            </div>
            <div className="short-summary-report-summary-table-cell short-summary-report-summary-table-cell-header">
              Styles
            </div>
            <div className="short-summary-report-summary-table-cell short-summary-report-summary-table-cell-header">
              Total Quantity
            </div>
            <div className="short-summary-report-summary-table-cell short-summary-report-summary-table-cell-header">
              Status
            </div>
            <div className="short-summary-report-summary-table-cell short-summary-report-summary-table-cell-header">
              Additional Info
            </div>
          </div>

          {departments.map(dept => {
            const summary = filteredDepartmentSummary[dept.id];
            if (!summary) return null;

            return (
              <div
                key={dept.id}
                className="short-summary-report-summary-table-row"
                style={activeDepartment === dept.id ? {
                  backgroundColor: `${dept.color}10`
                } : {}}
                onClick={() => setActiveDepartment(dept.id)}
              >
                <div className="short-summary-report-summary-table-cell" style={{ fontWeight: '500' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                    <span style={{ color: dept.color, fontSize: '16px' }}>
                      {dept.icon}
                    </span>
                    <span style={{ color: '#334155' }}>
                      {dept.name}
                    </span>
                  </div>
                </div>
                <div className="short-summary-report-summary-table-cell">
                  <span style={{ fontWeight: '600', color: '#1a237e' }}>{summary.totalRows}</span>
                  {departmentSummary[dept.id] && departmentSummary[dept.id].totalRows !== summary.totalRows && (
                    <span style={{
                      fontSize: '11px',
                      color: '#64748b',
                      marginLeft: '6px',
                      backgroundColor: '#f1f5f9',
                      padding: '2px 6px',
                      borderRadius: '4px'
                    }}>
                      (filtered)
                    </span>
                  )}
                </div>
                <div className="short-summary-report-summary-table-cell">
                  {summary.styles}
                </div>
                {/* Hide Total Quantity for 'all' and 'cutting_done' departments */}
                <div className="short-summary-report-summary-table-cell">
                  {dept.id === 'all' || dept.id === 'cutting_done' ? (
                    <span style={{
                      fontWeight: '600',
                      color: '#94a3b8',
                      fontStyle: 'italic'
                    }}>
                      ***
                    </span>
                  ) : (
                    <span style={{ fontWeight: '600', color: '#059669' }}>
                      {summary.totalPcs?.toLocaleString() || '0'}
                    </span>
                  )}
                </div>
                <div className="short-summary-report-summary-table-cell">
                  {dept.id === 'all' ? (
                    <span style={{
                      backgroundColor: '#e0f2fe',
                      color: '#0369a1',
                      padding: '4px 10px',
                      borderRadius: '6px',
                      fontSize: '12px',
                      fontWeight: '500',
                      display: 'inline-flex',
                      alignItems: 'center',
                      gap: '4px'
                    }}>
                      📊 Combined View
                    </span>
                  ) : dept.id === 'cutting_done' ? (
                    <span style={{
                      backgroundColor: summary.totalRows > 0 ? '#d1fae5' : '#fee2e2',
                      color: summary.totalRows > 0 ? '#065f46' : '#991b1b',
                      padding: '4px 10px',
                      borderRadius: '6px',
                      fontSize: '12px',
                      fontWeight: '500',
                      display: 'inline-flex',
                      alignItems: 'center',
                      gap: '4px'
                    }}>
                      {summary.totalRows > 0 ? '✅ Cutting Records' : 'No Data'}
                    </span>
                  ) : dept.id === 'direct_stitching' ? (
                    <span style={{
                      backgroundColor: '#d1fae5',
                      color: '#065f46',
                      padding: '4px 10px',
                      borderRadius: '6px',
                      fontSize: '12px',
                      fontWeight: '500',
                      display: 'inline-flex',
                      alignItems: 'center',
                      gap: '4px'
                    }}>
                      ✨ Direct Lots
                    </span>
                  ) : dept.id === 'direct_pending_issue' ? (
                    <span style={{
                      backgroundColor: summary.totalRows > 0 ? '#fef3c7' : '#d1fae5',
                      color: summary.totalRows > 0 ? '#92400e' : '#065f46',
                      padding: '4px 10px',
                      borderRadius: '6px',
                      fontSize: '12px',
                      fontWeight: '500',
                      display: 'inline-flex',
                      alignItems: 'center',
                      gap: '4px'
                    }}>
                      {summary.totalRows > 0 ? `⏱️ ${summary.totalRows} Pending` : '✅ All Issued'}
                    </span>
                  ) : dept.id === 'fabric_pending' ? (
                    <span style={{
                      backgroundColor: summary.totalRows > 0 ? '#fef3c7' : '#d1fae5',
                      color: summary.totalRows > 0 ? '#92400e' : '#065f46',
                      padding: '4px 10px',
                      borderRadius: '6px',
                      fontSize: '12px',
                      fontWeight: '500',
                      display: 'inline-flex',
                      alignItems: 'center',
                      gap: '4px'
                    }}>
                      {summary.totalRows > 0 ? `⏳ ${summary.totalRows} Pending` : '✅ All Issued'}
                    </span>
                  ) : dept.id === 'embroidery' || dept.id === 'printing' ? (
                    <span style={{
                      backgroundColor: summary.pendingLots > 0 ? '#fef3c7' : '#d1fae5',
                      color: summary.pendingLots > 0 ? '#92400e' : '#065f46',
                      padding: '4px 10px',
                      borderRadius: '6px',
                      fontSize: '12px',
                      fontWeight: '500',
                      display: 'inline-flex',
                      alignItems: 'center',
                      gap: '4px'
                    }}>
                      {summary.pendingLots > 0 ? '⏳ Pending' : '✅ Complete'}
                    </span>
                  ) : dept.id === 'colour_pending' ? (
                    <span style={{
                      backgroundColor: summary.totalRows > 0 ? '#fce7f3' : '#d1fae5',
                      color: summary.totalRows > 0 ? '#9d174d' : '#065f46',
                      padding: '4px 10px',
                      borderRadius: '6px',
                      fontSize: '12px',
                      fontWeight: '500',
                      display: 'inline-flex',
                      alignItems: 'center',
                      gap: '4px'
                    }}>
                      {summary.totalRows > 0 ? `🎨 ${summary.totalRows} Pending` : '✅ Complete'}
                    </span>
                  ) : dept.id === 'pending_after_emb_print' ? (
                    <span style={{
                      backgroundColor: summary.totalRows > 0 ? '#d1fae5' : '#fee2e2',
                      color: summary.totalRows > 0 ? '#065f46' : '#991b1b',
                      padding: '4px 10px',
                      borderRadius: '6px',
                      fontSize: '12px',
                      fontWeight: '500',
                      display: 'inline-flex',
                      alignItems: 'center',
                      gap: '4px'
                    }}>
                      {summary.totalRows > 0 ? '✅ Ready (No Direct)' : 'No Data'}
                    </span>
                  ) : (
                    <span style={{
                      backgroundColor: summary.totalRows > 0 ? '#d1fae5' : '#fee2e2',
                      color: summary.totalRows > 0 ? '#065f46' : '#991b1b',
                      padding: '4px 10px',
                      borderRadius: '6px',
                      fontSize: '12px',
                      fontWeight: '500'
                    }}>
                      {summary.totalRows > 0 ? 'Active' : 'No Data'}
                    </span>
                  )}
                </div>
                <div className="short-summary-report-summary-table-cell">
                  {dept.id === 'all' ? (
                    <span style={{ fontSize: '12px', color: '#475569' }}>
                      JobOrder: {summary.jobOrderRecords || 0} | Index: {summary.indexRecords || 0} | Direct: {summary.directStitchingLots || 0}
                    </span>
                  ) : dept.id === 'cutting_done' ? (
                    <span style={{ fontSize: '12px', color: '#475569' }}>
                      Total Index Records: {summary.indexTotalRecords || 0}
                    </span>
                  ) : dept.id === 'direct_stitching' ? (
                    <span style={{ fontSize: '12px', color: '#475569' }}>
                      Yes: {summary.yesCount || 0} | No: {summary.noCount || 0}
                    </span>
                  ) : dept.id === 'fabric_pending' ? (
                    <span style={{ fontSize: '12px', color: '#475569' }}>
                      Missing: {summary.missingLotsCount || 0} | Cancelled: {summary.cancelledLots || 0}
                    </span>
                  ) : dept.id === 'embroidery' || dept.id === 'printing' ? (
                    <span style={{ fontSize: '12px', color: '#475569' }}>
                      Pending: {summary.pendingLots || 0} lots ({summary.pendingQty || 0} pcs)
                    </span>
                  ) : dept.id === 'colour_pending' ? (
                    <span style={{ fontSize: '12px', color: '#475569' }}>
                      Pending Shades: {summary.pendingShades || 0}
                    </span>
                  ) : dept.id === 'stitching' ? (
                    <span style={{ fontSize: '12px', color: '#475569' }}>
                      Issued: {summary.stitchingIssued || 0}
                    </span>
                  ) : dept.id === 'pending_stitching' ? (
                    <span style={{ fontSize: '12px', color: '#475569' }}>
                      Pending Issue to Stitching After Cutting: {summary.pendingStitching || 0}
                    </span>
                  ) : dept.id === 'pending_after_emb_print' ? (
                    <span style={{ fontSize: '12px', color: '#475569' }}>
                      Direct: {summary.directStitchingCount || 0} | Emb/Print: {summary.embPrintCompletedCount || 0} | Avg Days: {summary.avgDaysSinceCompletion || 0}
                    </span>
                  ) : (
                    <span style={{ fontSize: '12px', color: '#94a3b8' }}>-</span>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      </div>

      {/* Supervisor Summary Section */}
      <div className="short-summary-report-supervisor-section">
        <div className="short-summary-report-supervisor-header" onClick={() => setShowSupervisorDetails(!showSupervisorDetails)}>
          <h3 className="short-summary-report-supervisor-title">
            <span>👥 Supervisor Wise Performance</span>
            <span className="short-summary-report-supervisor-toggle">
              {showSupervisorDetails ? '▼' : '▶'}
            </span>
          </h3>
          <div className="short-summary-report-supervisor-stats">
            <span className="short-summary-report-supervisor-stat">
              Total Supervisors: <strong>{Object.keys(filteredSupervisorSummary).length}</strong>
            </span>
            <span className="short-summary-report-supervisor-stat">
              Total Lots: <strong>{Object.values(filteredSupervisorSummary).reduce((sum, s) => sum + s.totalLots, 0)}</strong>
            </span>
            <span className="short-summary-report-supervisor-stat">
              Total Pcs: <strong>{Object.values(filteredSupervisorSummary).reduce((sum, s) => sum + s.totalPcs, 0).toLocaleString()}</strong>
            </span>
            <span className="short-summary-report-supervisor-stat">
              Pending Pcs: <strong>{Object.values(filteredSupervisorSummary).reduce((sum, s) => sum + s.pendingPcs, 0).toLocaleString()}</strong>
            </span>
          </div>
        </div>

        {showSupervisorDetails && (
          <div className="short-summary-report-supervisor-content">
            <div className="short-summary-report-supervisor-grid">
              {Object.keys(filteredSupervisorSummary).sort().map(supName => {
                const data = filteredSupervisorSummary[supName];
                const completionRate = ((data.completedLots / data.totalLots) * 100).toFixed(1);

                return (
                  <div key={supName} className="short-summary-report-supervisor-card">
                    <div className="short-summary-report-supervisor-card-header">
                      <span className="short-summary-report-supervisor-card-name">{supName}</span>
                      <span className={`short-summary-report-supervisor-card-badge ${data.pendingLots > 0 ? 'short-summary-report-supervisor-card-badge-pending' : 'short-summary-report-supervisor-card-badge-complete'
                        }`}>
                        {data.pendingLots > 0 ? `${data.pendingLots} Pending` : 'Complete'}
                      </span>
                    </div>

                    <div className="short-summary-report-supervisor-card-stats">
                      <div className="short-summary-report-supervisor-card-stat">
                        <span className="short-summary-report-supervisor-card-stat-label">Total Lots</span>
                        <span className="short-summary-report-supervisor-card-stat-value">{data.totalLots}</span>
                      </div>
                      <div className="short-summary-report-supervisor-card-stat">
                        <span className="short-summary-report-supervisor-card-stat-label">Completed</span>
                        <span className="short-summary-report-supervisor-card-stat-value" style={{ color: '#059669' }}>{data.completedLots}</span>
                      </div>
                      <div className="short-summary-report-supervisor-card-stat">
                        <span className="short-summary-report-supervisor-card-stat-label">Pending</span>
                        <span className="short-summary-report-supervisor-card-stat-value" style={{ color: '#dc2626' }}>{data.pendingLots}</span>
                      </div>
                      <div className="short-summary-report-supervisor-card-stat">
                        <span className="short-summary-report-supervisor-card-stat-label">Total Pcs</span>
                        <span className="short-summary-report-supervisor-card-stat-value">{data.totalPcs.toLocaleString()}</span>
                      </div>
                      <div className="short-summary-report-supervisor-card-stat">
                        <span className="short-summary-report-supervisor-card-stat-label">Pending Pcs</span>
                        <span className="short-summary-report-supervisor-card-stat-value" style={{ color: '#dc2626' }}>{data.pendingPcs.toLocaleString()}</span>
                      </div>
                    </div>

                    <div className="short-summary-report-supervisor-card-progress">
                      <div className="short-summary-report-supervisor-card-progress-bar">
                        <div
                          className="short-summary-report-supervisor-card-progress-fill"
                          style={{ width: `${completionRate}%` }}
                        ></div>
                      </div>
                      <span className="short-summary-report-supervisor-card-progress-text">{completionRate}% Complete</span>
                    </div>

                    {data.pendingLots > 0 && (
                      <div className="short-summary-report-supervisor-card-pending">
                        <span className="short-summary-report-supervisor-card-pending-title">Pending Lot Numbers:</span>
                        <div className="short-summary-report-supervisor-card-pending-list">
                          {data.pendingLotNumbers.slice(0, 3).map((lot, idx) => (
                            <span key={idx} className="short-summary-report-supervisor-card-pending-item">{lot}</span>
                          ))}
                          {data.pendingLotNumbers.length > 3 && (
                            <span className="short-summary-report-supervisor-card-pending-more">
                              +{data.pendingLotNumbers.length - 3} more
                            </span>
                          )}
                        </div>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </div>

      {/* Department Description */}
      <div className="short-summary-report-dept-description">
        <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
          <span style={{
            backgroundColor: currentDept.color,
            color: 'white',
            width: '32px',
            height: '32px',
            borderRadius: '8px',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            fontSize: '16px'
          }}>
            {currentDept.icon}
          </span>
          <div>
            <h3 style={{ margin: '0 0 4px 0', color: '#1a237e', fontSize: '16px' }}>
              {currentDept.name}
            </h3>
            <p style={{ margin: 0, color: '#64748b', fontSize: '14px' }}>
              {currentDept.description}
            </p>
          </div>
        </div>
      </div>

      {/* Special Message for Cutting Done */}
      {activeDepartment === 'cutting_done' && (
        <div className="short-summary-report-special-criteria">
          <h3 className="short-summary-report-criteria-title">
            <span>📋 Cutting Done Information:</span>
          </h3>
          <div className="short-summary-report-criteria-list">
            <div className="short-summary-report-criterion">
              <span className="short-summary-report-criterion-icon">✅</span>
              <span className="short-summary-report-criterion-text">
                <strong>All Index Records</strong> - Showing all lots from the Index sheet (cutting records)
              </span>
            </div>
            <div className="short-summary-report-criterion">
              <span className="short-summary-report-criterion-icon">📊</span>
              <span className="short-summary-report-criterion-text">
                Total Records: <strong>{filteredDepartmentSummary.cutting_done?.totalRows || 0}</strong>
              </span>
            </div>
            <div className="short-summary-report-criterion">
              <span className="short-summary-report-criterion-icon">🔍</span>
              <span className="short-summary-report-criterion-text">
                <strong>Data Source:</strong> Directly from Index sheet (No intersection with JobOrder)
              </span>
            </div>
          </div>
        </div>
      )}

      {/* Special Message for Direct Stitching */}
      {activeDepartment === 'direct_stitching' && (
        <div className="short-summary-report-special-criteria">
          <h3 className="short-summary-report-criteria-title">
            <span>📋 Direct Stitching Information:</span>
          </h3>
          <div className="short-summary-report-criteria-list">
            <div className="short-summary-report-criterion">
              <span className="short-summary-report-criterion-icon">✅</span>
              <span className="short-summary-report-criterion-text">
                <strong>Yes</strong> - Lots marked as Direct Stitching in Index sheet
              </span>
            </div>
            <div className="short-summary-report-criterion">
              <span className="short-summary-report-criterion-icon">❌</span>
              <span className="short-summary-report-criterion-text">
                <strong>No</strong> - Lots not marked as Direct Stitching
              </span>
            </div>
            <div className="short-summary-report-criterion">
              <span className="short-summary-report-criterion-icon">📊</span>
              <span className="short-summary-report-criterion-text">
                Total Direct Lots: <strong>{filteredDepartmentSummary.direct_stitching?.totalRows || 0}</strong>
              </span>
            </div>
          </div>
        </div>
      )}

      {/* Special Message for Fabric Pending */}
      {activeDepartment === 'fabric_pending' && (
        <div className="short-summary-report-special-criteria">
          <h3 className="short-summary-report-criteria-title">
            <span>📋 Fabric Issue Pending Criteria:</span>
          </h3>
          <div className="short-summary-report-criteria-list">
            <div className="short-summary-report-criterion">
              <span className="short-summary-report-criterion-icon">🔍</span>
              <span className="short-summary-report-criterion-text">
                These lots are from <strong>JobOrder sheet</strong> but are <strong>missing in Index sheet</strong>
              </span>
            </div>
            <div className="short-summary-report-criterion">
              <span className="short-summary-report-criterion-icon">❌</span>
              <span className="short-summary-report-criterion-text">
                <strong>Excluding cancelled lots</strong> (Status contains "Cancel" or "Cancelled")
              </span>
            </div>
            <div className="short-summary-report-criterion">
              <span className="short-summary-report-criterion-icon">📊</span>
              <span className="short-summary-report-criterion-text">
                Total JobOrder Lots: <strong>{jobOrderData.length}</strong> |
                Cancelled Lots: <strong>{filteredDepartmentSummary.fabric_pending?.cancelledLots || 0}</strong> |
                Missing Lots: <strong>{filteredDepartmentSummary.fabric_pending?.totalRows || 0}</strong>
              </span>
            </div>
          </div>
        </div>
      )}

      {/* Special Message for Pending After Emb/Print */}
      {/* Special Message for Pending After Emb/Print */}
      {activeDepartment === 'pending_after_emb_print' && (
        <div className="short-summary-report-special-criteria">
          <h3 className="short-summary-report-criteria-title">
            <span>✅ Pending Issue After Emb/Print Done (Excluding Direct Stitching)</span>
          </h3>
          <div className="short-summary-report-criteria-list">
            <div className="short-summary-report-criterion">
              <span className="short-summary-report-criterion-icon">🔍</span>
              <span className="short-summary-report-criterion-text">
                <strong>Lots ready for supervisor assignment</strong> - Completed Embroidery/Printing only (Direct Stitching EXCLUDED)
              </span>
            </div>
            <div className="short-summary-report-criterion">
              <span className="short-summary-report-criterion-icon">✅</span>
              <span className="short-summary-report-criterion-text">
                <strong>Criteria:</strong> Supervisor empty, Date of Issue empty, Not completed, All Emb/Print challans completed, and NOT Direct Stitching
              </span>
            </div>
            <div className="short-summary-report-criterion">
              <span className="short-summary-report-criterion-icon">📊</span>
              <span className="short-summary-report-criterion-text">
                Total Ready Lots: <strong>{filteredDepartmentSummary.pending_after_emb_print?.totalRows || 0}</strong> |
                Total Pcs: <strong>{filteredDepartmentSummary.pending_after_emb_print?.totalPcs?.toLocaleString() || 0}</strong>
              </span>
            </div>
            <div className="short-summary-report-criterion">
              <span className="short-summary-report-criterion-icon">📋</span>
              <span className="short-summary-report-criterion-text">
                Embroidery Only: <strong>{filteredDepartmentSummary.pending_after_emb_print?.embroideryOnlyCount || 0}</strong> |
                Printing Only: <strong>{filteredDepartmentSummary.pending_after_emb_print?.printingOnlyCount || 0}</strong> |
                Both: <strong>{filteredDepartmentSummary.pending_after_emb_print?.bothEmbPrintCount || 0}</strong> |
                Avg Days: <strong>{filteredDepartmentSummary.pending_after_emb_print?.avgDaysSinceCompletion || 0}</strong>
              </span>
            </div>
            {pendingIssueLots.length > 0 && (
              <div className="short-summary-report-criterion" style={{ flexDirection: 'column', alignItems: 'flex-start' }}>
                <span className="short-summary-report-criterion-icon">📋</span>
                <span className="short-summary-report-criterion-text">
                  <strong>Ready Lots Summary (Emb/Print Completed Only):</strong>
                </span>
                <div style={{ marginTop: '8px', maxHeight: '200px', overflowY: 'auto', width: '100%' }}>
                  {pendingIssueLots.slice(0, 5).map((row, idx) => {
                    const lotNumberIndex = displayColumnIndices.lotNumber;
                    const lotNumber = lotNumberIndex !== -1 ? row[lotNumberIndex] : 'Unknown';
                    const details = pendingIssueDetails[lotNumber?.toString().trim()];

                    let type = '';
                    if (details?.hasEmbroidery && details?.hasPrinting) type = 'Emb+Print';
                    else if (details?.hasEmbroidery) type = 'Emb Only';
                    else if (details?.hasPrinting) type = 'Print Only';

                    return (
                      <div key={idx} style={{
                        padding: '8px',
                        backgroundColor: '#f8fafc',
                        borderRadius: '6px',
                        marginBottom: '4px',
                        fontSize: '12px'
                      }}>
                        <strong>Lot {lotNumber}:</strong> {type} ({details?.totalCompletedChallans || 0} challans) |
                        Days: {details?.daysSinceCompletion || 'N/A'} | Pcs: {details?.totalPcs || 0}
                      </div>
                    );
                  })}
                  {pendingIssueLots.length > 5 && (
                    <div style={{ padding: '8px', color: '#64748b', fontSize: '12px' }}>
                      ... and {pendingIssueLots.length - 5} more lots
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      {/* Special Message for Colour Pending */}
      {activeDepartment === 'colour_pending' && (
        <div className="short-summary-report-special-criteria">
          <h3 className="short-summary-report-criteria-title">
            <span>🎨 Colour Pending Information:</span>
          </h3>
          <div className="short-summary-report-criteria-list">
            <div className="short-summary-report-criterion">
              <span className="short-summary-report-criterion-icon">🔍</span>
              <span className="short-summary-report-criterion-text">
                <strong>Analysis from Cutting sheet</strong> - Lots where not all shades/colours have been cut
              </span>
            </div>
            <div className="short-summary-report-criterion">
              <span className="short-summary-report-criterion-icon">📊</span>
              <span className="short-summary-report-criterion-text">
                <strong>Logic:</strong> For each lot, we check the Cutting sheet data against expected shades from Index sheet
              </span>
            </div>
            <div className="short-summary-report-criterion">
              <span className="short-summary-report-criterion-icon">📈</span>
              <span className="short-summary-report-criterion-text">
                Total Pending Lots: <strong>{filteredDepartmentSummary.colour_pending?.totalRows || 0}</strong> |
                Pending Shades: <strong>{filteredDepartmentSummary.colour_pending?.pendingShades || 0}</strong>
              </span>
            </div>
            {colourPendingLots.length > 0 && (
              <div className="short-summary-report-criterion" style={{ flexDirection: 'column', alignItems: 'flex-start' }}>
                <span className="short-summary-report-criterion-icon">📋</span>
                <span className="short-summary-report-criterion-text">
                  <strong>Pending Shades Summary:</strong>
                </span>
                <div style={{ marginTop: '8px', maxHeight: '200px', overflowY: 'auto', width: '100%' }}>
                  {colourPendingLots.slice(0, 5).map((row, idx) => {
                    const lotNumberIndex = displayColumnIndices.lotNumber;
                    const lotNumber = lotNumberIndex !== -1 ? row[lotNumberIndex] : 'Unknown';
                    const details = colourPendingDetails[lotNumber?.toString().trim()];
                    return (
                      <div key={idx} style={{
                        padding: '8px',
                        backgroundColor: '#f8fafc',
                        borderRadius: '6px',
                        marginBottom: '4px',
                        fontSize: '12px'
                      }}>
                        <strong>Lot {lotNumber}:</strong> {details?.pendingShades?.join(', ') || 'Unknown shades'}
                      </div>
                    );
                  })}
                  {colourPendingLots.length > 5 && (
                    <div style={{ padding: '8px', color: '#64748b', fontSize: '12px' }}>
                      ... and {colourPendingLots.length - 5} more lots
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
};

export default ShortSummaryReport;