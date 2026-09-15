import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx'; 
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { SPREADSHEET_IDS, fetchSheetDataFromBackend } from './config';

// Reusable Multi-Select Dropdown Component matching the KajButton Theme
const MultiSelectDropdown = ({ 
  label, 
  options = [], 
  selectedValues = [], 
  onChange, 
  placeholder = "All", 
  themeColor = "#4f46e5" 
}) => {
  const [isOpen, setIsOpen] = useState(false);
  const [searchTerm, setSearchTerm] = useState("");
  const dropdownRef = useRef(null);

  useEffect(() => {
    const handleClickOutside = (event) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target)) {
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

  const isAllSelected = selectedValues.length === 0;

  const getDisplayText = () => {
    if (isAllSelected) return placeholder;
    if (selectedValues.length === 1) return selectedValues[0];
    return `${selectedValues.length} Selected`;
  };

  const filteredOptions = options.filter((opt) => {
    const lbl = typeof opt === "object" ? opt.label : opt;
    return String(lbl || "").toLowerCase().includes(searchTerm.toLowerCase());
  });

  return (
    <div className="filter-group" ref={dropdownRef} style={{ position: "relative" }}>
      <label className="filter-label">{label}</label>
      <div
        onClick={() => setIsOpen(!isOpen)}
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          userSelect: "none",
          padding: "10px 14px",
          borderRadius: "12px",
          border: `1.5px solid ${selectedValues.length > 0 ? themeColor : "#e2e8f0"}`,
          background: selectedValues.length > 0 ? "#eef2ff" : "#ffffff",
          fontWeight: selectedValues.length > 0 ? "700" : "500",
          color: selectedValues.length > 0 ? themeColor : "#0f172a",
          cursor: "pointer",
          fontSize: "13px",
          minHeight: "42px",
          boxSizing: "border-box",
          transition: "all 0.2s"
        }}
      >
        <span style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: "160px" }}>
          {getDisplayText()}
        </span>
        <span style={{ fontSize: "10px", transform: isOpen ? "rotate(180deg)" : "rotate(0deg)", transition: "transform 0.2s", color: "#64748b" }}>
          ▼
        </span>
      </div>

      {isOpen && (
        <div
          style={{
            position: "absolute",
            top: "calc(100% + 4px)",
            left: 0,
            zIndex: 1000,
            background: "#ffffff",
            border: "1.5px solid #e2e8f0",
            borderRadius: "12px",
            boxShadow: "0 10px 25px rgba(0, 0, 0, 0.15)",
            minWidth: "220px",
            maxHeight: "260px",
            overflowY: "auto",
            padding: "8px"
          }}
        >
          {options.length > 5 && (
            <input
              type="text"
              placeholder={`Search ${label}...`}
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              onClick={(e) => e.stopPropagation()}
              style={{
                width: "100%",
                padding: "6px 10px",
                borderRadius: "8px",
                border: "1px solid #cbd5e1",
                fontSize: "12px",
                marginBottom: "6px",
                boxSizing: "border-box",
                outline: "none"
              }}
            />
          )}

          <div
            onClick={() => onChange([])}
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              padding: "6px 10px",
              borderRadius: "8px",
              cursor: "pointer",
              fontSize: "12px",
              fontWeight: 700,
              color: isAllSelected ? themeColor : "#64748b",
              background: isAllSelected ? "#eef2ff" : "transparent",
              borderBottom: "1px solid #f1f5f9",
              marginBottom: "4px"
            }}
          >
            <span>All (Clear Selection)</span>
            {isAllSelected && <span>✓</span>}
          </div>

          {filteredOptions.length > 0 ? (
            filteredOptions.map((opt) => {
              const val = typeof opt === "object" ? opt.value : opt;
              const lbl = typeof opt === "object" ? opt.label : opt;
              const isChecked = selectedValues.includes(val);
              return (
                <div
                  key={val}
                  onClick={() => toggleOption(val)}
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: "8px",
                    padding: "6px 10px",
                    borderRadius: "8px",
                    cursor: "pointer",
                    fontSize: "12px",
                    fontWeight: isChecked ? "700" : "500",
                    color: isChecked ? themeColor : "#1e293b",
                    background: isChecked ? "#eef2ff" : "transparent",
                    transition: "background 0.15s"
                  }}
                >
                  <input
                    type="checkbox"
                    checked={isChecked}
                    onChange={() => {}}
                    style={{ accentColor: themeColor, cursor: "pointer" }}
                  />
                  <span style={{ flex: 1 }}>{lbl}</span>
                </div>
              );
            })
          ) : (
            <div style={{ padding: "8px 10px", fontSize: "12px", color: "#94a3b8", textAlign: "center" }}>
              No options found
            </div>
          )}
        </div>
      )}
    </div>
  );
};

const DailyKajButtonReport = () => {
  const [data, setData] = useState([]);
  const [filteredData, setFilteredData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [filters, setFilters] = useState({
    lotNumber: '',
    supervisors: [],
    garmentTypes: [],
    fabrics: [],
    brands: [],
    parties: [],
    seasons: [],
    sections: [],
    stitchingSupervisors: [],
    statuses: ['WIP', 'Not Started'], // Active Lots by default
    minAging: '',
    maxAging: '',
    holdLots: false
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
      // Fetch KajButton data and JobOrder data in parallel
      const [kajButtonRes, jobOrderRes] = await Promise.allSettled([
        fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${RANGE}?key=${API_KEY}`),
        fetchSheetDataFromBackend(SPREADSHEET_IDS.JOBORDER, 'JobOrder!A:AZ')
      ]);

      // Parse JobOrder Sheet for Brand / Season / Section lookup against lots
      const lotToJobInfo = {};
      if (jobOrderRes.status === 'fulfilled' && jobOrderRes.value?.ok && Array.isArray(jobOrderRes.value.values)) {
        const jRows = jobOrderRes.value.values;
        const jHeaders = jRows[0] || [];
        const normalize = (s) => String(s || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
        const headerMap = {};
        jHeaders.forEach((h, idx) => {
          headerMap[normalize(h)] = idx;
        });

        const getVal = (row, key) => {
          const idx = headerMap[normalize(key)];
          return idx !== undefined && row[idx] !== undefined ? String(row[idx]).trim() : "";
        };

        for (let i = 1; i < jRows.length; i++) {
          const r = jRows[i];
          const lotNo = normalize(getVal(r, "Lot Number") || getVal(r, "Lot No") || getVal(r, "Lot"));
          if (!lotNo) continue;
          lotToJobInfo[lotNo] = {
            brand: getVal(r, "Brand") || getVal(r, "Brand Name") || getVal(r, "Party Name") || getVal(r, "Party"),
            party: getVal(r, "Party Name") || getVal(r, "Party"),
            garment: getVal(r, "Garment Type") || getVal(r, "Garment"),
            style: getVal(r, "Style"),
            fabric: getVal(r, "Fabric"),
            season: getVal(r, "Season"),
            section: getVal(r, "Section"),
            directStitching: getVal(r, "Direct Stitching") || getVal(r, "Direct")
          };
        }
      }

      let kajButtonValues = [];
      if (kajButtonRes.status === 'fulfilled' && kajButtonRes.value.ok) {
        const result = await kajButtonRes.value.json();
        if (result.values) kajButtonValues = result.values;
      } else {
        const backendKajButton = await fetchSheetDataFromBackend(SPREADSHEET_ID, RANGE);
        if (backendKajButton.ok && Array.isArray(backendKajButton.values)) {
          kajButtonValues = backendKajButton.values;
        }
      }
      
      if (kajButtonValues && kajButtonValues.length > 0) {
        const headers = kajButtonValues[0];
        const rows = kajButtonValues.slice(1);
        
        const formattedData = rows
          .filter(row => {
            // Skip completely empty rows
            if (!row || row.length === 0) return false;
            
            // Skip rows where the first column (Lot Number) is empty
            if (!row[0] || row[0].trim() === '') return false;
            
            const hasAnyData = row.some(cell => cell && cell.toString().trim() !== '');
            return hasAnyData;
          })
          .map(row => {
            const record = {};
            headers.forEach((header, index) => {
              record[header] = row[index] || '';
            });

            const rawLot = String(record['Lot Number'] || record['Lot No'] || row[0] || '').trim();
            const cleanLot = rawLot.toLowerCase().replace(/[^a-z0-9]/g, '');
            const jobInfo = lotToJobInfo[cleanLot] || {};
            
            const sheetBrand = String(record['BRAND'] || record['Brand'] || '').trim();
            const finalBrand = sheetBrand && sheetBrand !== '-' ? sheetBrand : (jobInfo.brand || jobInfo.party || '');
            record['BRAND'] = finalBrand;
            record['Brand'] = finalBrand;
            
            const sheetParty = String(record['Party Name'] || record['Party'] || '').trim();
            const finalParty = sheetParty && sheetParty !== '-' ? sheetParty : (jobInfo.party || jobInfo.brand || '');
            record['Party Name'] = finalParty;
            record['Party'] = finalParty;

            if ((!record['Garment Type'] || record['Garment Type'] === '-') && jobInfo.garment) record['Garment Type'] = jobInfo.garment;
            if ((!record['Style'] || record['Style'] === '-') && jobInfo.style) record['Style'] = jobInfo.style;
            if ((!record['Fabric'] || record['Fabric'] === '-') && jobInfo.fabric) record['Fabric'] = jobInfo.fabric;
            
            const sheetSeason = String(record['Season'] || '').trim();
            record['Season'] = sheetSeason && sheetSeason !== '-' ? sheetSeason : (jobInfo.season || '');
            
            const sheetSection = String(record['Section'] || '').trim();
            record['Section'] = sheetSection && sheetSection !== '-' ? sheetSection : (jobInfo.section || '');
            record['Direct Stitching'] = jobInfo.directStitching || '';
            
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
    
    // Apply multi-select status filter
    if (filters.statuses && filters.statuses.length > 0) {
      filtered = filtered.filter(item => {
        const itemStatus = getLotStatus(item);
        return filters.statuses.includes(itemStatus);
      });
    }
    
    // Apply Hold Lots filter
    if (filters.holdLots) {
      filtered = filtered.filter(item => isLotOnHold(item));
    }
    
    // Apply Lot Number search
    if (filters.lotNumber) {
      filtered = filtered.filter(item => 
        item['Lot Number']?.toLowerCase().includes(filters.lotNumber.toLowerCase())
      );
    }
    
    // Multi-select KajButton Supervisors
    if (filters.supervisors && filters.supervisors.length > 0) {
      filtered = filtered.filter(item => 
        filters.supervisors.includes(item['kajButton Supervisor'])
      );
    }
    
    // Multi-select Garment Types
    if (filters.garmentTypes && filters.garmentTypes.length > 0) {
      filtered = filtered.filter(item => 
        filters.garmentTypes.includes(item['Garment Type'])
      );
    }
    
    // Multi-select Fabrics
    if (filters.fabrics && filters.fabrics.length > 0) {
      filtered = filtered.filter(item => 
        filters.fabrics.includes(item['Fabric'])
      );
    }
    
    // Multi-select Brands
    if (filters.brands && filters.brands.length > 0) {
      filtered = filtered.filter(item => 
        filters.brands.includes(item['BRAND'])
      );
    }

    // Multi-select Party Names
    if (filters.parties && filters.parties.length > 0) {
      filtered = filtered.filter(item => 
        filters.parties.includes(item['Party Name'])
      );
    }
    
    // Multi-select Seasons
    if (filters.seasons && filters.seasons.length > 0) {
      filtered = filtered.filter(item => 
        filters.seasons.includes(item['Season'])
      );
    }
    
    // Multi-select Sections
    if (filters.sections && filters.sections.length > 0) {
      filtered = filtered.filter(item => 
        filters.sections.includes(item['Section'])
      );
    }
    
    // Multi-select Stitching Supervisors
    if (filters.stitchingSupervisors && filters.stitchingSupervisors.length > 0) {
      filtered = filtered.filter(item => 
        filters.stitchingSupervisors.includes(item['Stiching Supervisor'])
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

  // Professional Multi-Sheet Excel Export (matching PendingIssue format)
  const exportToExcel = async () => {
    if (filteredData.length === 0) {
      alert('No data available to export.');
      return;
    }

    try {
      const workbook = new ExcelJS.Workbook();
      workbook.creator = 'Factory Suite Pro';
      workbook.created = new Date();

      const totalLots = filteredData.length;
      const totalPieces = filteredData.reduce((sum, item) => sum + (parseInt(item['Total Pcs']) || 0), 0);

      // Grouping for Executive Summary
      const garmentMap = {};
      const supervisorMap = {};
      const seasonMap = {};
      let goodAgingLots = 0, goodAgingPcs = 0;
      let warnAgingLots = 0, warnAgingPcs = 0;
      let critAgingLots = 0, critAgingPcs = 0;
      let completedLots = 0, completedPcs = 0;
      let wipLots = 0, wipPcs = 0;
      let notStartedLots = 0, notStartedPcs = 0;

      filteredData.forEach(item => {
        const pcs = parseInt(item['Total Pcs']) || 0;
        const gType = (item['Garment Type'] || 'Unknown').trim();
        const sup = (item['kajButton Supervisor'] || 'Unassigned').trim();
        const season = (item['Season'] || 'Other / NA').trim();
        const aging = parseInt(item['Aging']) || 0;
        const status = getLotStatus(item);

        if (!garmentMap[gType]) garmentMap[gType] = { totalLots: 0, totalPcs: 0 };
        garmentMap[gType].totalLots += 1;
        garmentMap[gType].totalPcs += pcs;

        if (!supervisorMap[sup]) supervisorMap[sup] = { totalLots: 0, totalPcs: 0 };
        supervisorMap[sup].totalLots += 1;
        supervisorMap[sup].totalPcs += pcs;

        if (!seasonMap[season]) seasonMap[season] = { totalLots: 0, totalPcs: 0 };
        seasonMap[season].totalLots += 1;
        seasonMap[season].totalPcs += pcs;

        if (aging <= 2) {
          goodAgingLots += 1;
          goodAgingPcs += pcs;
        } else if (aging <= 5) {
          warnAgingLots += 1;
          warnAgingPcs += pcs;
        } else {
          critAgingLots += 1;
          critAgingPcs += pcs;
        }

        if (status === 'Completed') {
          completedLots += 1;
          completedPcs += pcs;
        } else if (status === 'WIP') {
          wipLots += 1;
          wipPcs += pcs;
        } else {
          notStartedLots += 1;
          notStartedPcs += pcs;
        }
      });

      const sortedGarments = Object.keys(garmentMap).map(name => ({
        name,
        totalLots: garmentMap[name].totalLots,
        totalPcs: garmentMap[name].totalPcs
      })).sort((a, b) => b.totalPcs - a.totalPcs);

      const sortedSupervisors = Object.keys(supervisorMap).map(name => ({
        name,
        totalLots: supervisorMap[name].totalLots,
        totalPcs: supervisorMap[name].totalPcs
      })).sort((a, b) => b.totalPcs - a.totalPcs);

      const thinBorder = {
        top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
      };

      // ================= SHEET 1: MAIN DATA =================
      const ws1 = workbook.addWorksheet('KajButton Report', {
        views: [{ showGridLines: true }]
      });

      // Title Banner
      ws1.mergeCells('A1:Q1');
      const titleCell = ws1.getCell('A1');
      titleCell.value = filters.holdLots ? 'FACTORY SUITE PRO - KAJBUTTON HOLD LOTS REPORT' : 'FACTORY SUITE PRO - DAILY KAJBUTTON PRODUCTION REPORT';
      titleCell.font = { name: 'Segoe UI', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
      titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
      titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
      ws1.getRow(1).height = 32;

      // Subtitle
      ws1.mergeCells('A2:Q2');
      const subCell = ws1.getCell('A2');
      subCell.value = `Report Date: ${new Date().toLocaleDateString('en-IN')}  |  Total Lots: ${totalLots}  |  Total Pieces: ${totalPieces.toLocaleString()}  |  Completed: ${completedLots}  |  WIP: ${wipLots}  |  Not Started: ${notStartedLots}`;
      subCell.font = { name: 'Segoe UI', size: 9.5, color: { argb: 'FFCBD5E1' } };
      subCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
      subCell.alignment = { horizontal: 'center', vertical: 'middle' };
      ws1.getRow(2).height = 22;

      // Spacer
      ws1.addRow([]);

      // Table Headers
      const headers1 = [
        'Sr No.', 'Lot Number', 'Garment Type', 'Style', 'Fabric', 'Brand',
        'Total Pcs', 'Section', 'Season', 'Party Name', 'Direct Stitching',
        'KajButton Date', 'KajButton Supervisor', 'Aging (Days)', 'Status',
        'Recent Remarks', 'Stitching Supervisor'
      ];
      const headerRow = ws1.addRow(headers1);
      headerRow.height = 25;
      headerRow.eachCell((cell) => {
        cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = thinBorder;
      });

      // Data Rows
      filteredData.forEach((item, idx) => {
        const pcs = parseInt(item['Total Pcs']) || 0;
        const aging = parseInt(item['Aging']) || 0;
        const status = getLotStatus(item);
        const remarks = getRecentRemarks(item);
        const isHold = isLotOnHold(item);

        const row = ws1.addRow([
          idx + 1,
          item['Lot Number'] || '—',
          item['Garment Type'] || '—',
          item['Style'] || '—',
          item['Fabric'] || '—',
          item['BRAND'] || '—',
          pcs,
          item['Section'] || '—',
          item['Season'] || '—',
          item['Party Name'] || '—',
          item['Direct Stitching'] || '—',
          item['KajButton Date'] || '—',
          item['kajButton Supervisor'] || '—',
          aging,
          status,
          remarks || '—',
          item['Stiching Supervisor'] || '—'
        ]);

        row.height = 20;
        row.eachCell((cell) => {
          cell.font = { name: 'Segoe UI', size: 9.5 };
          cell.border = thinBorder;
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          if (idx % 2 === 1) {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
          }
        });

        row.getCell(2).font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFDC2626' } };
        row.getCell(7).numFmt = '#,##0';
        row.getCell(7).font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFDC2626' } };

        if (aging <= 2) {
          row.getCell(14).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } };
          row.getCell(14).font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FF15803D' } };
        } else if (aging <= 5) {
          row.getCell(14).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
          row.getCell(14).font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FFB45309' } };
        } else {
          row.getCell(14).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
          row.getCell(14).font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FFDC2626' } };
        }

        if (status === 'Completed') {
          row.getCell(15).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } };
          row.getCell(15).font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FF15803D' } };
        } else if (status === 'WIP') {
          row.getCell(15).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E7FF' } };
          row.getCell(15).font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FF3730A3' } };
        }

        if (isHold) {
          row.getCell(16).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
          row.getCell(16).font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FFDC2626' } };
        }
      });

      // Total Row
      const totalRow1 = ws1.addRow([
        '',
        `TOTAL (${totalLots} Lots)`,
        '',
        '',
        '',
        '',
        totalPieces,
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        `${completedLots} Comp | ${wipLots} WIP`,
        '',
        ''
      ]);
      totalRow1.height = 24;
      totalRow1.eachCell((cell) => {
        cell.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FF000000' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FF000000' } },
          bottom: { style: 'double', color: { argb: 'FF000000' } },
          left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
          right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
        };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      });
      totalRow1.getCell(7).numFmt = '#,##0';

      const colWidths = [8, 16, 20, 20, 20, 16, 14, 12, 14, 20, 16, 16, 22, 14, 16, 30, 20];
      colWidths.forEach((w, i) => {
        ws1.getColumn(i + 1).width = w;
      });

      // ================= SHEET 2: EXECUTIVE SUMMARY =================
      const ws2 = workbook.addWorksheet('Executive Summary', {
        views: [{ showGridLines: true }]
      });

      // Section 1: Garment Type Breakdown
      ws2.mergeCells('A1:D1');
      const gTitle = ws2.getCell('A1');
      gTitle.value = '1. GARMENT TYPE BREAKDOWN (LOTS & PIECES DISTRIBUTION)';
      gTitle.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      gTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F766E' } };
      gTitle.alignment = { horizontal: 'left', vertical: 'middle' };
      ws2.getRow(1).height = 26;

      const gHeader = ws2.addRow(['Garment Type', 'Total Lots', 'Total Pieces (Qty)', 'Share %']);
      gHeader.height = 22;
      gHeader.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF134E4A' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = thinBorder;
      });

      sortedGarments.forEach((item, idx) => {
        const pct = totalPieces > 0 ? (item.totalPcs / totalPieces) : 0;
        const r = ws2.addRow([item.name, item.totalLots, item.totalPcs, pct]);
        r.height = 19;
        r.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(3).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(3).numFmt = '#,##0';
        r.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(4).numFmt = '0.0%';
        r.eachCell(c => {
          c.font = { name: 'Segoe UI', size: 9 };
          c.border = thinBorder;
          if (idx % 2 === 1) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        });
      });

      const gTotalRow = ws2.addRow(['TOTAL', totalLots, totalPieces, 1]);
      gTotalRow.height = 22;
      gTotalRow.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
        c.border = { top: { style: 'thin' }, bottom: { style: 'double' }, left: { style: 'thin' }, right: { style: 'thin' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
      });
      gTotalRow.getCell(3).numFmt = '#,##0';
      gTotalRow.getCell(4).numFmt = '0.0%';

      // Spacer
      ws2.addRow([]);

      // Section 2: Supervisor Breakdown
      const supStartRow = ws2.rowCount + 1;
      ws2.mergeCells(`A${supStartRow}:D${supStartRow}`);
      const supTitle = ws2.getCell(`A${supStartRow}`);
      supTitle.value = '2. SUPERVISOR WORKLOAD DISTRIBUTION';
      supTitle.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      supTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E40AF' } };
      supTitle.alignment = { horizontal: 'left', vertical: 'middle' };
      ws2.getRow(supStartRow).height = 26;

      const supHeader = ws2.addRow(['Supervisor Name', 'Total Lots', 'Total Pieces (Qty)', 'Share %']);
      supHeader.height = 22;
      supHeader.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = thinBorder;
      });

      sortedSupervisors.forEach((item, idx) => {
        const pct = totalPieces > 0 ? (item.totalPcs / totalPieces) : 0;
        const r = ws2.addRow([item.name, item.totalLots, item.totalPcs, pct]);
        r.height = 19;
        r.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(3).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(3).numFmt = '#,##0';
        r.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(4).numFmt = '0.0%';
        r.eachCell(c => {
          c.font = { name: 'Segoe UI', size: 9 };
          c.border = thinBorder;
          if (idx % 2 === 1) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        });
      });

      const supTotalRow = ws2.addRow(['TOTAL', totalLots, totalPieces, 1]);
      supTotalRow.height = 22;
      supTotalRow.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
        c.border = { top: { style: 'thin' }, bottom: { style: 'double' }, left: { style: 'thin' }, right: { style: 'thin' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
      });
      supTotalRow.getCell(3).numFmt = '#,##0';
      supTotalRow.getCell(4).numFmt = '0.0%';

      // Spacer
      ws2.addRow([]);

      // Section 3: Aging Breakdown
      const slaStartRow = ws2.rowCount + 1;
      ws2.mergeCells(`A${slaStartRow}:D${slaStartRow}`);
      const slaTitle = ws2.getCell(`A${slaStartRow}`);
      slaTitle.value = '3. AGING & SLA PERFORMANCE';
      slaTitle.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      slaTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB45309' } };
      slaTitle.alignment = { horizontal: 'left', vertical: 'middle' };
      ws2.getRow(slaStartRow).height = 26;

      const slaHeader = ws2.addRow(['Aging Bracket', 'Total Lots', 'Total Pieces (Qty)', 'Share %']);
      slaHeader.height = 22;
      slaHeader.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF78350F' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = thinBorder;
      });

      const slaData = [
        { name: '<= 2 Days (Good / On Track)', lots: goodAgingLots, pcs: goodAgingPcs, bg: 'FFDCFCE7', fg: 'FF15803D' },
        { name: '3 - 5 Days (Warning Zone)', lots: warnAgingLots, pcs: warnAgingPcs, bg: 'FFFEF3C7', fg: 'FFB45309' },
        { name: '> 5 Days (Critical Delay)', lots: critAgingLots, pcs: critAgingPcs, bg: 'FFFEE2E2', fg: 'FFDC2626' }
      ];

      slaData.forEach(item => {
        const pct = totalPieces > 0 ? (item.pcs / totalPieces) : 0;
        const r = ws2.addRow([item.name, item.lots, item.pcs, pct]);
        r.height = 20;
        r.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(3).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(3).numFmt = '#,##0';
        r.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
        r.getCell(4).numFmt = '0.0%';
        r.eachCell(c => {
          c.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: item.fg } };
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: item.bg } };
          c.border = thinBorder;
        });
      });

      const slaTotalRow = ws2.addRow(['TOTAL', totalLots, totalPieces, 1]);
      slaTotalRow.height = 22;
      slaTotalRow.eachCell(c => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
        c.border = { top: { style: 'thin' }, bottom: { style: 'double' }, left: { style: 'thin' }, right: { style: 'thin' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
      });
      slaTotalRow.getCell(3).numFmt = '#,##0';
      slaTotalRow.getCell(4).numFmt = '0.0%';

      ws2.getColumn(1).width = 34;
      ws2.getColumn(2).width = 16;
      ws2.getColumn(3).width = 22;
      ws2.getColumn(4).width = 16;

      // ================= SHEET 3: APPLIED FILTERS =================
      const ws3 = workbook.addWorksheet('Applied Filters', {
        views: [{ showGridLines: true }]
      });

      ws3.mergeCells('A1:B1');
      const fTitle = ws3.getCell('A1');
      fTitle.value = 'APPLIED FILTERS & METADATA';
      fTitle.font = { name: 'Segoe UI', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
      fTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
      fTitle.alignment = { horizontal: 'left', vertical: 'middle' };
      ws3.getRow(1).height = 28;

      const filterItems = [
        ['Report Generated', `${new Date().toLocaleDateString('en-IN')} ${new Date().toLocaleTimeString('en-IN')}`],
        ['Total Lots Exported', totalLots],
        ['Total Pieces Exported', totalPieces.toLocaleString()],
        ['Hold Lots Filter', filters.holdLots ? 'Yes (Hold Lots Only)' : 'All Lots'],
        ['Status Filter', filters.statuses?.length ? filters.statuses.join(', ') : 'All Statuses'],
        ['Supervisor Filter', filters.supervisors?.length ? filters.supervisors.join(', ') : 'All Supervisors'],
        ['Garment Type Filter', filters.garmentTypes?.length ? filters.garmentTypes.join(', ') : 'All Types'],
        ['Fabric Filter', filters.fabrics?.length ? filters.fabrics.join(', ') : 'All Fabrics'],
        ['Brand Filter', filters.brands?.length ? filters.brands.join(', ') : 'All Brands'],
        ['Party Filter', filters.parties?.length ? filters.parties.join(', ') : 'All Parties'],
        ['Season Filter', filters.seasons?.length ? filters.seasons.join(', ') : 'All Seasons'],
        ['Aging Range', `${filters.minAging || 0} - ${filters.maxAging || 'Max'} Days`],
        ['Lot Search Query', filters.lotNumber || 'None']
      ];

      filterItems.forEach(([k, v], idx) => {
        const r = ws3.addRow([k, v]);
        r.height = 20;
        r.getCell(1).font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FF1E293B' } };
        r.getCell(2).font = { name: 'Segoe UI', size: 9.5, color: { argb: 'FF334155' } };
        r.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
        r.getCell(2).alignment = { horizontal: 'left', vertical: 'middle' };
        r.eachCell(c => {
          c.border = thinBorder;
          if (idx % 2 === 1) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        });
      });

      ws3.getColumn(1).width = 24;
      ws3.getColumn(2).width = 50;

      // Generate and Save Excel File
      const buffer = await workbook.xlsx.writeBuffer();
      const ts = new Date().toISOString().slice(0, 10);
      saveAs(new Blob([buffer]), `Daily_KajButton_Report_${ts}.xlsx`);

    } catch (e) {
      console.error('Excel export error:', e);
      alert(`Excel export failed: ${e.message}`);
    }
  };

  // Export to PDF Function
// Export to PDF Function - FIXED VERSION
// Export to PDF Function - COMPLETELY REWRITTEN
// Export to PDF Function - PRINT VERSION (Opens Print Dialog)
  // --- Professional PDF Export (A3 Landscape - Matching Factory Suite Pro Standard) ---
  const exportToPDF = () => {
    if (filteredData.length === 0) {
      alert('No data available to export.');
      return;
    }

    try {
      const doc = new jsPDF({
        orientation: "landscape",
        unit: "pt",
        format: "a3"
      });

      const pageW = doc.internal.pageSize.getWidth();
      const pageH = doc.internal.pageSize.getHeight();
      const totalPieces = filteredData.reduce((sum, item) => sum + (parseInt(item['Total Pcs']) || 0), 0);
      const totalLots = filteredData.length;

      // Grouping for 4-Column Side-by-Side Executive Summary
      const garmentMap = {};
      const supervisorMap = {};
      const seasonMap = {};
      let goodAgingLots = 0, goodAgingPcs = 0;
      let warnAgingLots = 0, warnAgingPcs = 0;
      let critAgingLots = 0, critAgingPcs = 0;
      let completedLots = 0, completedPcs = 0;
      let wipLots = 0, wipPcs = 0;
      let notStartedLots = 0, notStartedPcs = 0;

      filteredData.forEach(item => {
        const pcs = parseInt(item['Total Pcs']) || 0;
        const gType = (item['Garment Type'] || 'Unknown').trim();
        const sup = (item['kajButton Supervisor'] || 'Unassigned').trim();
        const season = (item['Season'] || 'Other / NA').trim();
        const aging = parseInt(item['Aging']) || 0;
        const status = getLotStatus(item);

        if (!garmentMap[gType]) garmentMap[gType] = { totalLots: 0, totalPcs: 0 };
        garmentMap[gType].totalLots += 1;
        garmentMap[gType].totalPcs += pcs;

        if (!supervisorMap[sup]) supervisorMap[sup] = { totalLots: 0, totalPcs: 0 };
        supervisorMap[sup].totalLots += 1;
        supervisorMap[sup].totalPcs += pcs;

        if (!seasonMap[season]) seasonMap[season] = { totalLots: 0, totalPcs: 0 };
        seasonMap[season].totalLots += 1;
        seasonMap[season].totalPcs += pcs;

        if (aging <= 2) {
          goodAgingLots += 1;
          goodAgingPcs += pcs;
        } else if (aging <= 5) {
          warnAgingLots += 1;
          warnAgingPcs += pcs;
        } else {
          critAgingLots += 1;
          critAgingPcs += pcs;
        }

        if (status === 'Completed') {
          completedLots += 1;
          completedPcs += pcs;
        } else if (status === 'WIP') {
          wipLots += 1;
          wipPcs += pcs;
        } else {
          notStartedLots += 1;
          notStartedPcs += pcs;
        }
      });

      const sortedGarments = Object.keys(garmentMap).map(name => ({
        name,
        totalLots: garmentMap[name].totalLots,
        totalPcs: garmentMap[name].totalPcs
      })).sort((a, b) => b.totalPcs - a.totalPcs);

      const sortedSupervisors = Object.keys(supervisorMap).map(name => ({
        name,
        totalLots: supervisorMap[name].totalLots,
        totalPcs: supervisorMap[name].totalPcs
      })).sort((a, b) => b.totalPcs - a.totalPcs);

      const sortedSeasons = Object.keys(seasonMap).map(name => ({
        name,
        totalLots: seasonMap[name].totalLots,
        totalPcs: seasonMap[name].totalPcs
      })).sort((a, b) => b.totalPcs - a.totalPcs);

      // 1. Main Header Block
      doc.setFillColor(15, 23, 42); // Dark Navy #0F172A
      doc.rect(15, 12, pageW - 30, 48, 'F');

      doc.setFontSize(14);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(255, 255, 255);
      const reportTitle = filters.holdLots 
        ? "FACTORY SUITE PRO - KAJBUTTON HOLD LOTS REPORT" 
        : "FACTORY SUITE PRO - DAILY KAJBUTTON PRODUCTION REPORT";
      doc.text(reportTitle, pageW / 2, 30, { align: 'center' });

      doc.setFontSize(8.5);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(199, 210, 254);
      const subText = `Total Lots: ${totalLots}   |   Total Pieces: ${totalPieces.toLocaleString()}   |   Completed: ${completedLots}   |   WIP: ${wipLots}   |   Not Started: ${notStartedLots}   |   Aging <=2d: ${goodAgingLots}   |   3-5d: ${warnAgingLots}   |   >5d: ${critAgingLots}`;
      doc.text(subText, pageW / 2, 48, { align: 'center' });

      // 2. Filter Banner
      doc.setFillColor(241, 245, 249);
      doc.rect(15, 63, pageW - 30, 16, 'F');
      doc.setFontSize(8);
      doc.setFont('helvetica', 'italic');
      doc.setTextColor(0, 0, 0);
      const filterSummary = `Filters: Status: ${filters.statuses?.length ? filters.statuses.join(', ') : 'All'} | Supervisors: ${filters.supervisors?.length ? filters.supervisors.join(', ') : 'All'} | Garments: ${filters.garmentTypes?.length ? filters.garmentTypes.join(', ') : 'All'} | Fabrics: ${filters.fabrics?.length ? filters.fabrics.join(', ') : 'All'} | Brands: ${filters.brands?.length ? filters.brands.join(', ') : 'All'} | Seasons: ${filters.seasons?.length ? filters.seasons.join(', ') : 'All'} | Aging: ${filters.minAging || 0} - ${filters.maxAging || 'Max'}d | Search: ${filters.lotNumber || 'None'}`;
      doc.text(filterSummary, pageW / 2, 74, { align: 'center' });

      // 3. Main Data Table
      const tableColumns = [
        '#',
        'Lot Number',
        'Garment Type',
        'Style',
        'Fabric',
        'Brand',
        'Total Pcs',
        'Section',
        'Season',
        'Party Name',
        'Direct Stitching',
        'KajButton Date',
        'Supervisor',
        'Aging (Days)',
        'Status',
        'Recent Remarks',
        'Stitching Sup'
      ];

      const tableBody = filteredData.map((item, idx) => {
        const pcs = parseInt(item['Total Pcs']) || 0;
        const aging = parseInt(item['Aging']) || 0;
        const status = getLotStatus(item);
        const remarks = getRecentRemarks(item);

        return [
          (idx + 1).toString(),
          item['Lot Number'] || '—',
          item['Garment Type'] || '—',
          item['Style'] || '—',
          item['Fabric'] || '—',
          item['BRAND'] || '—',
          pcs.toLocaleString(),
          item['Section'] || '—',
          item['Season'] || '—',
          item['Party Name'] || '—',
          item['Direct Stitching'] || '—',
          item['KajButton Date'] || '—',
          item['kajButton Supervisor'] || '—',
          `${aging}d`,
          status,
          remarks || 'No remarks',
          item['Stiching Supervisor'] || '—'
        ];
      });

      // Total Row
      tableBody.push([
        '',
        `TOTAL (${totalLots})`,
        '',
        '',
        '',
        '',
        totalPieces.toLocaleString(),
        '',
        '',
        '',
        '',
        '',
        `${sortedSupervisors.length} Sups`,
        '',
        `${completedLots} Comp | ${wipLots} WIP`,
        '',
        ''
      ]);

      const columnStyles = {
        0: { cellWidth: 25, halign: 'center' },
        1: { cellWidth: 65, halign: 'center', fontStyle: 'bold' },
        2: { cellWidth: 80, halign: 'center' },
        3: { cellWidth: 80, halign: 'center' },
        4: { cellWidth: 80, halign: 'center' },
        5: { cellWidth: 65, halign: 'center' },
        6: { cellWidth: 55, halign: 'center', fontStyle: 'bold' },
        7: { cellWidth: 45, halign: 'center' },
        8: { cellWidth: 55, halign: 'center' },
        9: { cellWidth: 80, halign: 'center' },
        10: { cellWidth: 50, halign: 'center' },
        11: { cellWidth: 65, halign: 'center' },
        12: { cellWidth: 75, halign: 'center' },
        13: { cellWidth: 50, halign: 'center' },
        14: { cellWidth: 65, halign: 'center' },
        15: { cellWidth: 160, halign: 'center' },
        16: { cellWidth: 65, halign: 'center' }
      };

      autoTable(doc, {
        head: [tableColumns],
        body: tableBody,
        startY: 85,
        tableWidth: pageW - 30,
        margin: { top: 85, right: 15, bottom: 25, left: 15 },
        theme: "grid",
        styles: {
          fontSize: 8.5,
          cellPadding: { top: 4, right: 3, bottom: 4, left: 3 },
          overflow: "linebreak",
          valign: 'middle',
          halign: 'center',
          textColor: [0, 0, 0], // Pure Black
          lineColor: [0, 0, 0], // Black grid lines
          lineWidth: 0.3,
          fontStyle: 'normal',
          minCellHeight: 12,
        },
        headStyles: {
          fillColor: [15, 23, 42],
          textColor: [255, 255, 255],
          fontStyle: "bold",
          lineColor: [0, 0, 0],
          lineWidth: 0.5,
          halign: 'center',
          fontSize: 9,
          valign: 'middle',
          cellPadding: { top: 5, right: 3, bottom: 5, left: 3 },
        },
        alternateRowStyles: {
          fillColor: [248, 250, 252],
        },
        columnStyles,
        didParseCell: function (data) {
          if (data.section === 'body') {
            const rowIndex = data.row.index;
            const isTotalRow = rowIndex === tableBody.length - 1;

            if (isTotalRow) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [226, 232, 240];
              data.cell.styles.textColor = [0, 0, 0];
              data.cell.styles.halign = 'center';
              return;
            }

            const item = filteredData[rowIndex];
            if (!item) return;

            // Lot number styling
            if (data.column.index === 1) {
              data.cell.styles.textColor = [220, 38, 38];
              data.cell.styles.fontStyle = 'bold';
            }

            // Total Pcs styling
            if (data.column.index === 6) {
              data.cell.styles.textColor = [220, 38, 38];
              data.cell.styles.fontStyle = 'bold';
            }

            // Aging styling
            if (data.column.index === 13) {
              const days = parseInt(item['Aging']) || 0;
              if (days > 5) {
                data.cell.styles.fillColor = [239, 68, 68];
                data.cell.styles.textColor = [255, 255, 255];
                data.cell.styles.fontStyle = "bold";
              } else if (days >= 3) {
                data.cell.styles.fillColor = [254, 243, 199];
                data.cell.styles.textColor = [180, 83, 9];
                data.cell.styles.fontStyle = "bold";
              } else {
                data.cell.styles.fillColor = [220, 252, 231];
                data.cell.styles.textColor = [21, 128, 61];
                data.cell.styles.fontStyle = "bold";
              }
            }

            // Status styling
            if (data.column.index === 14) {
              const status = getLotStatus(item);
              if (status === 'Completed') {
                data.cell.styles.fillColor = [220, 252, 231];
                data.cell.styles.textColor = [21, 128, 61];
                data.cell.styles.fontStyle = 'bold';
              } else if (status === 'WIP') {
                data.cell.styles.fillColor = [224, 231, 255];
                data.cell.styles.textColor = [55, 48, 163];
                data.cell.styles.fontStyle = 'bold';
              } else {
                data.cell.styles.fillColor = [241, 245, 249];
                data.cell.styles.textColor = [100, 116, 139];
              }
            }

            // Remarks hold styling
            if (data.column.index === 15 && isLotOnHold(item)) {
              data.cell.styles.fillColor = [254, 226, 226];
              data.cell.styles.textColor = [185, 28, 28];
              data.cell.styles.fontStyle = 'bold';
            }
          }
        }
      });

      // --- 4-COLUMN SIDE-BY-SIDE EXECUTIVE SUMMARY ---
      const gBody = sortedGarments.map(item => {
        const pct = totalPieces > 0 ? ((item.totalPcs / totalPieces) * 100).toFixed(1) : "0.0";
        return [item.name, item.totalLots.toString(), item.totalPcs.toLocaleString(), `${pct}%`];
      });
      gBody.push(["TOTAL", totalLots.toString(), totalPieces.toLocaleString(), "100.0%"]);

      const supBody = sortedSupervisors.map(item => {
        const pct = totalPieces > 0 ? ((item.totalPcs / totalPieces) * 100).toFixed(1) : "0.0";
        return [item.name, item.totalLots.toString(), item.totalPcs.toLocaleString(), `${pct}%`];
      });
      supBody.push(["TOTAL", totalLots.toString(), totalPieces.toLocaleString(), "100.0%"]);

      const seasonBody = sortedSeasons.map(item => {
        const pct = totalPieces > 0 ? ((item.totalPcs / totalPieces) * 100).toFixed(1) : "0.0";
        return [item.name, item.totalLots.toString(), item.totalPcs.toLocaleString(), `${pct}%`];
      });
      seasonBody.push(["TOTAL", totalLots.toString(), totalPieces.toLocaleString(), "100.0%"]);

      const slaBody = [
        ["<= 2 Days (Good)", goodAgingLots.toString(), goodAgingPcs.toLocaleString(), `${totalLots > 0 ? ((goodAgingLots / totalLots) * 100).toFixed(1) : 0}%`],
        ["3 - 5 Days (Warning)", warnAgingLots.toString(), warnAgingPcs.toLocaleString(), `${totalLots > 0 ? ((warnAgingLots / totalLots) * 100).toFixed(1) : 0}%`],
        ["> 5 Days (Critical)", critAgingLots.toString(), critAgingPcs.toLocaleString(), `${totalLots > 0 ? ((critAgingLots / totalLots) * 100).toFixed(1) : 0}%`],
        ["TOTAL", totalLots.toString(), totalPieces.toLocaleString(), "100.0%"]
      ];

      const maxRows = Math.max(gBody.length, supBody.length, seasonBody.length, slaBody.length);
      const approxSummaryHeight = 55 + (maxRows * 18);

      let summaryStartY = doc.lastAutoTable.finalY + 22;
      const neededSpace = approxSummaryHeight + 35;
      if (summaryStartY + neededSpace > pageH - 30) {
        doc.addPage();
        summaryStartY = 40;
      } else {
        doc.setDrawColor(203, 213, 225);
        doc.setLineWidth(0.8);
        doc.line(15, summaryStartY - 8, pageW - 15, summaryStartY - 8);
      }

      // Title & KPI Subtitle
      doc.setFontSize(14);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(0, 0, 0);
      doc.text("EXECUTIVE SUMMARY & PRODUCTION BREAKDOWN", pageW / 2, summaryStartY + 4, { align: 'center' });

      doc.setFontSize(9);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(0, 0, 0);
      const summarySub = `Total Lots: ${totalLots}   |   Total Pieces: ${totalPieces.toLocaleString()} Pcs   |   Supervisors: ${sortedSupervisors.length}   |   Garments: ${sortedGarments.length}   |   Seasons: ${sortedSeasons.length}`;
      doc.text(summarySub, pageW / 2, summaryStartY + 16, { align: 'center' });

      const sectionTitleY = summaryStartY + 30;
      const tableStartY = sectionTitleY + 6;

      // 4 Columns Side-by-Side Configuration (Exactly matching full page width)
      const colWidth = 278;
      const gap = 16;
      const col1X = 15;
      const col2X = col1X + colWidth + gap; // 309
      const col3X = col2X + colWidth + gap; // 603
      const col4X = col3X + colWidth + gap; // 897

      doc.setFontSize(9.5);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(0, 0, 0);
      doc.text("1. GARMENT BREAKDOWN", col1X, sectionTitleY);
      doc.text("2. SUPERVISOR BREAKDOWN", col2X, sectionTitleY);
      doc.text("3. SEASON BREAKDOWN", col3X, sectionTitleY);
      doc.text("4. AGING & SLA BREAKDOWN", col4X, sectionTitleY);

      const summaryColStyles = {
        0: { cellWidth: 110, halign: 'center' },
        1: { cellWidth: 45, halign: 'center' },
        2: { cellWidth: 68, halign: 'center' },
        3: { cellWidth: 55, halign: 'center' },
      };

      // Column 1 Table: Garment Breakdown
      autoTable(doc, {
        head: [['Garment Type', 'Lots', 'Total Pcs', 'Share %']],
        body: gBody,
        startY: tableStartY,
        tableWidth: colWidth,
        margin: { left: col1X, right: pageW - (col1X + colWidth) },
        theme: "grid",
        styles: {
          fontSize: 8.5,
          cellPadding: { top: 3.5, right: 2, bottom: 3.5, left: 2 },
          overflow: "linebreak",
          valign: 'middle',
          halign: 'center',
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          lineWidth: 0.3,
        },
        headStyles: {
          fillColor: [15, 118, 110], // Teal
          textColor: [255, 255, 255],
          fontStyle: "bold",
          fontSize: 8.5,
          halign: 'center',
          cellPadding: { top: 4, right: 2, bottom: 4, left: 2 },
        },
        columnStyles: summaryColStyles,
        didParseCell: function (data) {
          if (data.section === 'body') {
            data.cell.styles.textColor = [0, 0, 0];
            data.cell.styles.halign = 'center';
            if (data.row.index === gBody.length - 1) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });
      const endY1 = doc.lastAutoTable.finalY;

      // Column 2 Table: Supervisor Breakdown
      autoTable(doc, {
        head: [['Supervisor', 'Lots', 'Total Pcs', 'Share %']],
        body: supBody,
        startY: tableStartY,
        tableWidth: colWidth,
        margin: { left: col2X, right: pageW - (col2X + colWidth) },
        theme: "grid",
        styles: {
          fontSize: 8.5,
          cellPadding: { top: 3.5, right: 2, bottom: 3.5, left: 2 },
          overflow: "linebreak",
          valign: 'middle',
          halign: 'center',
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          lineWidth: 0.3,
        },
        headStyles: {
          fillColor: [67, 56, 202], // Indigo
          textColor: [255, 255, 255],
          fontStyle: "bold",
          fontSize: 8.5,
          halign: 'center',
          cellPadding: { top: 4, right: 2, bottom: 4, left: 2 },
        },
        columnStyles: summaryColStyles,
        didParseCell: function (data) {
          if (data.section === 'body') {
            data.cell.styles.textColor = [0, 0, 0];
            data.cell.styles.halign = 'center';
            if (data.row.index === supBody.length - 1) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });
      const endY2 = doc.lastAutoTable.finalY;

      // Column 3 Table: Season Breakdown
      autoTable(doc, {
        head: [['Season', 'Lots', 'Total Pcs', 'Share %']],
        body: seasonBody,
        startY: tableStartY,
        tableWidth: colWidth,
        margin: { left: col3X, right: pageW - (col3X + colWidth) },
        theme: "grid",
        styles: {
          fontSize: 8.5,
          cellPadding: { top: 3.5, right: 2, bottom: 3.5, left: 2 },
          overflow: "linebreak",
          valign: 'middle',
          halign: 'center',
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          lineWidth: 0.3,
        },
        headStyles: {
          fillColor: [30, 64, 175], // Royal Blue
          textColor: [255, 255, 255],
          fontStyle: "bold",
          fontSize: 8.5,
          halign: 'center',
          cellPadding: { top: 4, right: 2, bottom: 4, left: 2 },
        },
        columnStyles: summaryColStyles,
        didParseCell: function (data) {
          if (data.section === 'body') {
            data.cell.styles.textColor = [0, 0, 0];
            data.cell.styles.halign = 'center';
            if (data.row.index === seasonBody.length - 1) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });
      const endY3 = doc.lastAutoTable.finalY;

      // Column 4 Table: SLA & Aging Breakdown
      autoTable(doc, {
        head: [['Aging Bracket', 'Lots', 'Total Pcs', 'Share %']],
        body: slaBody,
        startY: tableStartY,
        tableWidth: colWidth,
        margin: { left: col4X, right: pageW - (col4X + colWidth) },
        theme: "grid",
        styles: {
          fontSize: 8.5,
          cellPadding: { top: 3.5, right: 2, bottom: 3.5, left: 2 },
          overflow: "linebreak",
          valign: 'middle',
          halign: 'center',
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          lineWidth: 0.3,
        },
        headStyles: {
          fillColor: [180, 83, 9], // Amber
          textColor: [255, 255, 255],
          fontStyle: "bold",
          fontSize: 8.5,
          halign: 'center',
          cellPadding: { top: 4, right: 2, bottom: 4, left: 2 },
        },
        columnStyles: summaryColStyles,
        didParseCell: function (data) {
          if (data.section === 'body') {
            data.cell.styles.textColor = [0, 0, 0];
            data.cell.styles.halign = 'center';
            if (data.row.index === 0) {
              data.cell.styles.fillColor = [220, 252, 231]; // Soft Green
              data.cell.styles.fontStyle = 'bold';
            } else if (data.row.index === 1) {
              data.cell.styles.fillColor = [254, 243, 199]; // Soft Amber
              data.cell.styles.fontStyle = 'bold';
            } else if (data.row.index === 2) {
              data.cell.styles.fillColor = [254, 226, 226]; // Soft Red
              data.cell.styles.fontStyle = 'bold';
            } else if (data.row.index === 3) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });
      const endY4 = doc.lastAutoTable.finalY;

      const maxEndY = Math.max(endY1, endY2, endY3, endY4);
      const finalY = maxEndY + 16;
      if (finalY <= pageH - 22) {
        doc.setDrawColor(0, 0, 0);
        doc.setLineWidth(0.5);
        doc.line(15, finalY, pageW - 15, finalY);

        doc.setFontSize(8.5);
        doc.setFont('helvetica', 'italic');
        doc.setTextColor(0, 0, 0);
        doc.text("Daily Kaj Button Department Report — Factory Suite Pro", 15, finalY + 12);
      }

      // Page Numbering Loop
      const totalPages = doc.internal.getNumberOfPages();
      const timeStr = `${new Date().toLocaleDateString('en-IN')} ${new Date().toLocaleTimeString('en-IN')}`;
      for (let i = 1; i <= totalPages; i++) {
        doc.setPage(i);
        doc.setFontSize(8.5);
        doc.setFont('helvetica', 'normal');
        doc.setTextColor(0, 0, 0);

        doc.setDrawColor(203, 213, 225);
        doc.setLineWidth(0.5);
        doc.line(15, pageH - 22, pageW - 15, pageH - 22);

        doc.text(`Page ${i} of ${totalPages}`, pageW / 2, pageH - 12, { align: 'center' });
        doc.text(`Generated: ${timeStr}`, pageW - 18, pageH - 12, { align: 'right' });
      }

      const ts = new Date().toISOString().slice(0, 10);
      doc.save(filters.holdLots ? `Daily_KajButton_Hold_Lots_${ts}.pdf` : `Daily_KajButton_Report_${ts}.pdf`);

    } catch (error) {
      console.error('Error generating PDF:', error);
      alert('Failed to generate PDF. Please try again.');
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
      supervisors: [],
      garmentTypes: [],
      fabrics: [],
      brands: [],
      parties: [],
      seasons: [],
      sections: [],
      stitchingSupervisors: [],
      statuses: ['WIP', 'Not Started'],
      minAging: '',
      maxAging: '',
      holdLots: false
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
          
          .table-card {
            background: #ffffff;
            border-radius: 20px;
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.05);
            border: 1px solid #f1f5f9;
            overflow: hidden;
            margin-bottom: 40px;
          }

          .table-container {
            overflow: auto;
            max-height: calc(100vh - 280px);
            min-height: 350px;
            position: relative;
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
                  <div className="detail-label">Party Name</div>
                  <div className="detail-value">{selectedRow['Party Name'] || 'N/A'}</div>
                </div>
                <div className="detail-card">
                  <div className="detail-label">Season</div>
                  <div className="detail-value">{selectedRow['Season'] || 'N/A'}</div>
                </div>
                <div className="detail-card">
                  <div className="detail-label">Section</div>
                  <div className="detail-value">{selectedRow['Section'] || 'N/A'}</div>
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
            
            <MultiSelectDropdown
              label="KajButton Supervisor"
              options={getUniqueValues('kajButton Supervisor')}
              selectedValues={filters.supervisors}
              onChange={(vals) => handleFilterChange('supervisors', vals)}
              placeholder="All Supervisors"
            />
            
            <MultiSelectDropdown
              label="Garment Type"
              options={getUniqueValues('Garment Type')}
              selectedValues={filters.garmentTypes}
              onChange={(vals) => handleFilterChange('garmentTypes', vals)}
              placeholder="All Types"
            />
            
            <MultiSelectDropdown
              label="Fabric"
              options={getUniqueValues('Fabric')}
              selectedValues={filters.fabrics}
              onChange={(vals) => handleFilterChange('fabrics', vals)}
              placeholder="All Fabrics"
            />
            
            <MultiSelectDropdown
              label="Brand"
              options={getUniqueValues('BRAND')}
              selectedValues={filters.brands}
              onChange={(vals) => handleFilterChange('brands', vals)}
              placeholder="All Brands"
            />

            <MultiSelectDropdown
              label="Party Name"
              options={getUniqueValues('Party Name')}
              selectedValues={filters.parties}
              onChange={(vals) => handleFilterChange('parties', vals)}
              placeholder="All Parties"
            />
            
            <MultiSelectDropdown
              label="Season"
              options={getUniqueValues('Season')}
              selectedValues={filters.seasons}
              onChange={(vals) => handleFilterChange('seasons', vals)}
              placeholder="All Seasons"
            />

            <MultiSelectDropdown
              label="Section"
              options={getUniqueValues('Section')}
              selectedValues={filters.sections}
              onChange={(vals) => handleFilterChange('sections', vals)}
              placeholder="All Sections"
            />
            
            <MultiSelectDropdown
              label="Stitching Supervisor"
              options={getUniqueValues('Stiching Supervisor')}
              selectedValues={filters.stitchingSupervisors}
              onChange={(vals) => handleFilterChange('stitchingSupervisors', vals)}
              placeholder="All Stiching Supervisors"
            />
            
            <MultiSelectDropdown
              label="Status"
              options={[
                { value: 'WIP', label: 'WIP' },
                { value: 'Not Started', label: 'Not Started' },
                { value: 'Completed', label: 'Completed' }
              ]}
              selectedValues={filters.statuses}
              onChange={(vals) => handleFilterChange('statuses', vals)}
              placeholder="All Status"
            />
            
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
        <div className="table-card">
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
          
          <div className="table-container">
            <table className="data-table">
              <thead>
                <tr>
                  <th>Lot #</th>
                  <th>Garment Type</th>
                  <th>Style</th>
                  <th>Fabric</th>
                  <th>Brand</th>
                  <th>Total Pcs</th>
                  <th>Section</th>
                  <th>Season</th>
                  <th>Party Name</th>
                  <th>Direct Stitching</th>
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
                        <td>{item['Garment Type']}</td>
                        <td>{item['Style']}</td>
                        <td>{item['Fabric']}</td>
                        <td>
                          {item['BRAND'] && (
                            <span className="brand-tag">{item['BRAND']}</span>
                          )}
                        </td>
                        <td>{item['Total Pcs']}</td>
                        <td>
                          {item['Section'] ? (
                            <span className="brand-tag" style={{ background: '#fef3c7', color: '#92400e', border: '1px solid #fde68a' }}>
                              {item['Section']}
                            </span>
                          ) : '-'}
                        </td>
                        <td>
                          {item['Season'] ? (
                            <span className="brand-tag" style={{ background: '#f0fdf4', color: '#166534', border: '1px solid #bbf7d0' }}>
                              {item['Season']}
                            </span>
                          ) : '-'}
                        </td>
                        <td>
                          {item['Party Name'] ? (
                            <span className="party-tag" style={{ background: '#f5f3ff', color: '#6d28d9', border: '1px solid #ddd6fe', padding: '4px 10px', borderRadius: '9999px', fontSize: '12px', fontWeight: '700', display: 'inline-block' }}>
                              {item['Party Name']}
                            </span>
                          ) : '-'}
                        </td>
                        <td>{item['Direct Stitching'] || '-'}</td>
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
                    <td colSpan="15">
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
      </div>
    </>
  );
};

export default DailyKajButtonReport;