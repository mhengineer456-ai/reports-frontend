import React, { useState, useEffect } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { jsPDF } from 'jspdf';
import { autoTable } from 'jspdf-autotable';
import { SPREADSHEET_IDS, fetchSheetDataFromBackend } from './config';

const DailyFoldingReport = () => {
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
    holdLots: false
  });
  const [selectedRow, setSelectedRow] = useState(null);
  const [modalOpen, setModalOpen] = useState(false);

  // Google Sheets API configuration
  const SPREADSHEET_ID = '1IMhmYlJ3s2PPRgEQs1Ikd4O1OBXK4EYL1oV_-kWAkyg';
  const API_KEY = 'AIzaSyAomDFBkOySlIxKWSKGHe6ATv9gvaBr7uk';
  const RANGE = 'Folding!B:O';

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
      // Fetch Folding data and JobOrder data in parallel
      const [foldingRes, jobOrderRes] = await Promise.allSettled([
        fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${RANGE}?key=${API_KEY}`),
        fetchSheetDataFromBackend(SPREADSHEET_IDS.JOBORDER, 'JobOrder!A:AZ')
      ]);

      // Parse JobOrder Sheet for Brand / Party lookup against lots
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

      let foldingValues = [];
      if (foldingRes.status === 'fulfilled' && foldingRes.value.ok) {
        const result = await foldingRes.value.json();
        if (result.values) foldingValues = result.values;
      } else {
        const backendFolding = await fetchSheetDataFromBackend(SPREADSHEET_ID, RANGE);
        if (backendFolding.ok && Array.isArray(backendFolding.values)) {
          foldingValues = backendFolding.values;
        }
      }
      
      if (foldingValues && foldingValues.length > 0) {
        const headers = foldingValues[0];
        const rows = foldingValues.slice(1);
        
        const formattedData = rows
          .filter(row => {
            if (!row || row.length === 0) return false;
            if (!row[0] || row[0].trim() === '') return false;
            return row.some(cell => cell && cell.toString().trim() !== '');
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
            if ((!record['Garment Type'] || record['Garment Type'] === '-') && jobInfo.garment) record['Garment Type'] = jobInfo.garment;
            if ((!record['Style'] || record['Style'] === '-') && jobInfo.style) record['Style'] = jobInfo.style;
            if ((!record['Fabric'] || record['Fabric'] === '-') && jobInfo.fabric) record['Fabric'] = jobInfo.fabric;
            if ((!record['Season'] || record['Season'] === '-') && jobInfo.season) record['Season'] = jobInfo.season;
            if ((!record['Section'] || record['Section'] === '-') && jobInfo.section) record['Section'] = jobInfo.section;
            if ((!record['Party Name'] || record['Party Name'] === '-') && jobInfo.party) record['Party Name'] = jobInfo.party;
            if ((!record['Direct Stitching'] || record['Direct Stitching'] === '-') && jobInfo.directStitching) record['Direct Stitching'] = jobInfo.directStitching;
            
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
    const complete = record['Folding Complete'] || '';
    const wip = record['WIP Folding'] || '';
    
    if (complete.trim() && complete !== '[]') {
      return 'Completed';
    } else if (wip.trim() && wip !== '[]') {
      return 'WIP';
    } else {
      return 'Not Started';
    }
  };

  const isLotOnHold = (record) => {
    const remarks = getRecentRemarks(record).toLowerCase();
    return remarks.includes('hold');
  };

  const calculateAging = (record) => {
    try {
      const foldingDate = record['Folding Date'];
      if (!foldingDate) return 0;

      const today = new Date();
      today.setHours(0, 0, 0, 0);

      const foldingDateObj = new Date(foldingDate);
      foldingDateObj.setHours(0, 0, 0, 0);

      let endDate = today;

      const complete = record['Folding Complete'] || '';
      if (complete.trim() && complete !== '[]') {
        try {
          const statusData = JSON.parse(complete);
          if (statusData.length > 0) {
            const latestStatus = statusData[statusData.length - 1];
            endDate = new Date(latestStatus.timestamp);
            endDate.setHours(0, 0, 0, 0);
          }
        } catch (e) {
          endDate = today;
        }
      }
      
      const diffTime = endDate.getTime() - foldingDateObj.getTime();
      const diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24));
      
      return diffDays < 0 ? 0 : diffDays;
    } catch (error) {
      return 0;
    }
  };

  const getRecentRemarks = (record) => {
    try {
      const wip = record['WIP Folding'] || '';
      if (wip.trim() && wip !== '[]') {
        const statusData = JSON.parse(wip);
        if (statusData.length > 0) {
          const sortedData = statusData.sort((a, b) => 
            new Date(b.timestamp) - new Date(a.timestamp)
          );
          const latest = sortedData[0];
          return latest.remarks || (latest.status ? latest.status : 'No remarks');
        }
      }
      
      const complete = record['Folding Complete'] || '';
      if (complete.trim() && complete !== '[]') {
        try {
          const statusData = JSON.parse(complete);
          if (statusData.length > 0) {
            const sortedData = statusData.sort((a, b) => 
              new Date(b.timestamp) - new Date(a.timestamp)
            );
            const latest = sortedData[0];
            return latest.remarks || 'Completed';
          }
        } catch (e) {}
      }
      
      return 'No remarks';
    } catch (error) {
      return 'No remarks';
    }
  };

  const getAllStatusHistory = (record) => {
    try {
      let allStatuses = [];
      const wip = record['WIP Folding'] || '';
      if (wip.trim() && wip !== '[]') {
        const wipData = JSON.parse(wip);
        allStatuses = [...allStatuses, ...wipData];
      }
      
      const complete = record['Folding Complete'] || '';
      if (complete.trim() && complete !== '[]') {
        const completeData = JSON.parse(complete);
        allStatuses = [...allStatuses, ...completeData];
      }
      
      return allStatuses.sort((a, b) => 
        new Date(b.timestamp) - new Date(a.timestamp)
      );
    } catch (error) {
      return [];
    }
  };

  const filterData = () => {
    let filtered = [...data];
    
    if (filters.status === 'active') {
      filtered = filtered.filter(item => getLotStatus(item) !== 'Completed');
    } else if (filters.status !== 'all') {
      filtered = filtered.filter(item => getLotStatus(item) === filters.status);
    }
    
    if (filters.holdLots) {
      filtered = filtered.filter(item => isLotOnHold(item));
    }
    
    if (filters.lotNumber) {
      filtered = filtered.filter(item => 
        (item['Lot Number'] || '').toLowerCase().includes(filters.lotNumber.toLowerCase())
      );
    }
    
    if (filters.supervisor) {
      filtered = filtered.filter(item => 
        (item['Folding Supervisor'] || '').toLowerCase().includes(filters.supervisor.toLowerCase())
      );
    }
    
    if (filters.garmentType) {
      filtered = filtered.filter(item => 
        (item['Garment Type'] || '').trim().toLowerCase() === filters.garmentType.trim().toLowerCase()
      );
    }
    
    if (filters.fabric) {
      filtered = filtered.filter(item => 
        (item['Fabric'] || '').toLowerCase().includes(filters.fabric.toLowerCase())
      );
    }
    
    if (filters.brand) {
      filtered = filtered.filter(item => 
        (item['BRAND'] || '').toLowerCase().includes(filters.brand.toLowerCase())
      );
    }
    
    if (filters.stitchingSupervisor) {
      filtered = filtered.filter(item => 
        (item['Stiching Supervisor'] || '').toLowerCase().includes(filters.stitchingSupervisor.toLowerCase())
      );
    }
    
    if (filters.minAging) {
      filtered = filtered.filter(item => 
        parseInt(item['Aging'] || 0) >= parseInt(filters.minAging)
      );
    }
    
    if (filters.maxAging) {
      filtered = filtered.filter(item => 
        parseInt(item['Aging'] || 0) <= parseInt(filters.maxAging)
      );
    }
    
    setFilteredData(filtered);
  };

  // Professional Multi-Sheet Excel Export (Matching Factory Suite Pro Standard)
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

      // Aggregations for Executive Summary
      const garmentMap = {};
      const supervisorMap = {};
      const stitchingSupMap = {};
      let goodAgingLots = 0, goodAgingPcs = 0;
      let warnAgingLots = 0, warnAgingPcs = 0;
      let critAgingLots = 0, critAgingPcs = 0;
      let completedLots = 0, completedPcs = 0;
      let wipLots = 0, wipPcs = 0;
      let notStartedLots = 0, notStartedPcs = 0;

      filteredData.forEach(item => {
        const pcs = parseInt(item['Total Pcs']) || 0;
        const gType = (item['Garment Type'] || 'Unknown').trim();
        const sup = (item['Folding Supervisor'] || 'Unassigned').trim();
        const stSup = (item['Stiching Supervisor'] || 'Unassigned').trim();
        const aging = calculateAging(item);
        const status = getLotStatus(item);

        if (!garmentMap[gType]) garmentMap[gType] = { lots: 0, pcs: 0, compLots: 0, wipLots: 0 };
        garmentMap[gType].lots += 1;
        garmentMap[gType].pcs += pcs;
        if (status === 'Completed') garmentMap[gType].compLots += 1;
        else garmentMap[gType].wipLots += 1;

        if (!supervisorMap[sup]) supervisorMap[sup] = { lots: 0, pcs: 0, compLots: 0, wipLots: 0 };
        supervisorMap[sup].lots += 1;
        supervisorMap[sup].pcs += pcs;
        if (status === 'Completed') supervisorMap[sup].compLots += 1;
        else supervisorMap[sup].wipLots += 1;

        if (!stitchingSupMap[stSup]) stitchingSupMap[stSup] = { lots: 0, pcs: 0, compLots: 0, wipLots: 0 };
        stitchingSupMap[stSup].lots += 1;
        stitchingSupMap[stSup].pcs += pcs;
        if (status === 'Completed') stitchingSupMap[stSup].compLots += 1;
        else stitchingSupMap[stSup].wipLots += 1;

        if (aging <= 3) {
          goodAgingLots += 1;
          goodAgingPcs += pcs;
        } else if (aging <= 7) {
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

      const sortedGarments = Object.keys(garmentMap).map(k => ({ name: k, ...garmentMap[k] })).sort((a, b) => b.pcs - a.pcs);
      const sortedSupervisors = Object.keys(supervisorMap).map(k => ({ name: k, ...supervisorMap[k] })).sort((a, b) => b.pcs - a.pcs);
      const sortedStitchingSups = Object.keys(stitchingSupMap).map(k => ({ name: k, ...stitchingSupMap[k] })).sort((a, b) => b.pcs - a.pcs);

      const thinBorder = {
        top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
      };

      // ================= SHEET 1: MAIN DATA =================
      const ws1 = workbook.addWorksheet('Folding Report', {
        views: [{ showGridLines: true }]
      });

      // Title Banner
      ws1.mergeCells('A1:Q1');
      const titleCell = ws1.getCell('A1');
      titleCell.value = filters.holdLots ? 'FACTORY SUITE PRO - FOLDING HOLD LOTS REPORT' : 'FACTORY SUITE PRO - DAILY FOLDING PRODUCTION REPORT';
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
        'Folding Date', 'Folding Supervisor', 'Aging (Days)', 'Status',
        'Recent Remarks', 'Stitching Supervisor'
      ];
      const headerRow = ws1.addRow(headers1);
      headerRow.height = 25;
      headerRow.eachCell((cell) => {
        cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = thinBorder;
      });

      // Data Rows
      filteredData.forEach((item, idx) => {
        const pcs = parseInt(item['Total Pcs']) || 0;
        const aging = calculateAging(item);
        const status = getLotStatus(item);
        const remarks = getRecentRemarks(item);

        const rowData = [
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
          item['Folding Date'] || '—',
          item['Folding Supervisor'] || '—',
          aging,
          status,
          remarks,
          item['Stiching Supervisor'] || '—'
        ];

        const row = ws1.addRow(rowData);
        row.height = 20;

        const isEven = idx % 2 === 0;
        const rowBgColor = isEven ? 'FFFFFFFF' : 'FFF8FAFC';

        row.eachCell((cell, colNumber) => {
          cell.border = thinBorder;
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBgColor } };
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.font = { name: 'Segoe UI', size: 9, color: { argb: 'FF1E293B' } };

          // Lot Number styling
          if (colNumber === 2) {
            cell.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFDC2626' } };
          }
          // Brand styling
          if (colNumber === 6) {
            cell.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FF1E293B' } };
          }
          // Total Pcs styling
          if (colNumber === 7) {
            cell.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFDC2626' } };
            cell.numFmt = '#,##0';
          }
          // Supervisor styling
          if (colNumber === 13) {
            cell.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FF1E293B' } };
          }
          // Aging column badge
          if (colNumber === 14) {
            if (aging <= 3) {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } };
              cell.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FF15803D' } };
            } else if (aging <= 7) {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
              cell.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FFB45309' } };
            } else {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
              cell.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FFB91C1C' } };
            }
          }
          // Status column badge
          if (colNumber === 15) {
            if (status === 'Completed') {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } };
              cell.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FF15803D' } };
            } else if (status === 'WIP') {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E7FF' } };
              cell.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FF3730A3' } };
            } else {
              cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
              cell.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FF64748B' } };
            }
          }
        });
      });

      // Total Summary Row
      const totalRow = ws1.addRow([
        '',
        'TOTAL',
        `${totalLots} Lots`,
        '',
        '',
        'TOTAL PCS:',
        totalPieces,
        '',
        '',
        '',
        '',
        '',
        `${sortedSupervisors.length} Sups`,
        '',
        `${completedLots} Comp / ${wipLots} WIP`,
        '',
        ''
      ]);
      totalRow.height = 24;
      totalRow.eachCell((cell, colNumber) => {
        cell.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FF0F172A' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = {
          top: { style: 'medium', color: { argb: 'FF0F172A' } },
          bottom: { style: 'double', color: { argb: 'FF0F172A' } },
          left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
          right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
        };
        if (colNumber === 7) {
          cell.numFmt = '#,##0';
          cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFDC2626' } };
        }
      });

      // Auto-fit column widths
      ws1.columns.forEach((col) => {
        let maxLen = 12;
        col.eachCell({ includeEmpty: false }, (cell) => {
          const valStr = cell.value ? cell.value.toString() : '';
          if (valStr.length > maxLen) {
            maxLen = Math.min(valStr.length, 36);
          }
        });
        col.width = Math.max(maxLen + 3, 11);
      });

      // ================= SHEET 2: EXECUTIVE SUMMARY =================
      const ws2 = workbook.addWorksheet('Executive Summary', {
        views: [{ showGridLines: true }]
      });

      // Summary Header
      ws2.mergeCells('A1:G1');
      const sumTitle = ws2.getCell('A1');
      sumTitle.value = 'FOLDING DEPARTMENT - EXECUTIVE KPI & WORKLOAD BREAKDOWN';
      sumTitle.font = { name: 'Segoe UI', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
      sumTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
      sumTitle.alignment = { horizontal: 'center', vertical: 'middle' };
      ws2.getRow(1).height = 30;

      // Section 1: Garment Type Summary
      let curRow = 3;
      ws2.mergeCells(`A${curRow}:F${curRow}`);
      const gSec = ws2.getCell(`A${curRow}`);
      gSec.value = '👕 GARMENT TYPE DISTRIBUTION';
      gSec.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      gSec.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F766E' } };
      gSec.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
      ws2.getRow(curRow).height = 24;

      curRow++;
      const gHeaders = ['Garment Type', 'Total Lots', 'Total Pieces', 'Completed Lots', 'WIP Lots', '% Share'];
      const gHeaderRow = ws2.addRow(gHeaders);
      gHeaderRow.height = 22;
      gHeaderRow.eachCell((c) => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = thinBorder;
      });

      sortedGarments.forEach((g) => {
        curRow++;
        const share = totalPieces > 0 ? (g.pcs / totalPieces) * 100 : 0;
        const r = ws2.addRow([
          g.name,
          g.lots,
          g.pcs,
          g.compLots,
          g.wipLots,
          `${share.toFixed(1)}%`
        ]);
        r.height = 19;
        r.eachCell((c, colIdx) => {
          c.border = thinBorder;
          c.alignment = { horizontal: 'center', vertical: 'middle' };
          c.font = { name: 'Segoe UI', size: 9 };
          if (colIdx === 1) {
            c.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
            c.font = { name: 'Segoe UI', size: 9, bold: true };
          }
          if (colIdx === 3) {
            c.numFmt = '#,##0';
            c.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FFDC2626' } };
          }
        });
      });

      // Section 2: Folding Supervisor Summary
      curRow += 3;
      ws2.mergeCells(`A${curRow}:F${curRow}`);
      const sSec = ws2.getCell(`A${curRow}`);
      sSec.value = '👤 FOLDING SUPERVISOR ALLOCATION';
      sSec.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      sSec.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E40AF' } };
      sSec.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
      ws2.getRow(curRow).height = 24;

      curRow++;
      const sHeaders = ['Folding Supervisor', 'Total Lots', 'Total Pieces', 'Completed Lots', 'WIP Lots', '% Share'];
      const sHeaderRow = ws2.addRow(sHeaders);
      sHeaderRow.height = 22;
      sHeaderRow.eachCell((c) => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = thinBorder;
      });

      sortedSupervisors.forEach((s) => {
        curRow++;
        const share = totalPieces > 0 ? (s.pcs / totalPieces) * 100 : 0;
        const r = ws2.addRow([
          s.name,
          s.lots,
          s.pcs,
          s.compLots,
          s.wipLots,
          `${share.toFixed(1)}%`
        ]);
        r.height = 19;
        r.eachCell((c, colIdx) => {
          c.border = thinBorder;
          c.alignment = { horizontal: 'center', vertical: 'middle' };
          c.font = { name: 'Segoe UI', size: 9 };
          if (colIdx === 1) {
            c.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
            c.font = { name: 'Segoe UI', size: 9, bold: true };
          }
          if (colIdx === 3) {
            c.numFmt = '#,##0';
            c.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FFDC2626' } };
          }
        });
      });

      // Section 3: Stitching Supervisor Summary
      curRow += 3;
      ws2.mergeCells(`A${curRow}:E${curRow}`);
      const stSec = ws2.getCell(`A${curRow}`);
      stSec.value = '🧵 STITCHING SUPERVISOR ORIGIN';
      stSec.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      stSec.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF6B21A8' } };
      stSec.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
      ws2.getRow(curRow).height = 24;

      curRow++;
      const stHeaders = ['Stitching Supervisor', 'Total Lots', 'Total Pieces', 'Completed Lots', 'WIP Lots'];
      const stHeaderRow = ws2.addRow(stHeaders);
      stHeaderRow.height = 22;
      stHeaderRow.eachCell((c) => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = thinBorder;
      });

      sortedStitchingSups.forEach((st) => {
        curRow++;
        const r = ws2.addRow([
          st.name,
          st.lots,
          st.pcs,
          st.compLots,
          st.wipLots
        ]);
        r.height = 19;
        r.eachCell((c, colIdx) => {
          c.border = thinBorder;
          c.alignment = { horizontal: 'center', vertical: 'middle' };
          c.font = { name: 'Segoe UI', size: 9 };
          if (colIdx === 1) {
            c.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
            c.font = { name: 'Segoe UI', size: 9, bold: true };
          }
          if (colIdx === 3) {
            c.numFmt = '#,##0';
            c.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FFDC2626' } };
          }
        });
      });

      // Section 4: Aging Breakdown
      curRow += 3;
      ws2.mergeCells(`A${curRow}:E${curRow}`);
      const aSec = ws2.getCell(`A${curRow}`);
      aSec.value = '⏱️ AGING & SLA PERFORMANCE';
      aSec.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
      aSec.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB45309' } };
      aSec.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
      ws2.getRow(curRow).height = 24;

      curRow++;
      const aHeaders = ['Aging Category', 'SLA Days', 'Total Lots', 'Total Pieces', '% Share'];
      const aHeaderRow = ws2.addRow(aHeaders);
      aHeaderRow.height = 22;
      aHeaderRow.eachCell((c) => {
        c.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        c.border = thinBorder;
      });

      const agingTiers = [
        { name: '🟢 Fast Track (Good)', days: '0 - 3 Days', lots: goodAgingLots, pcs: goodAgingPcs },
        { name: '🟡 Moderate (Warning)', days: '4 - 7 Days', lots: warnAgingLots, pcs: warnAgingPcs },
        { name: '🔴 High Aging (Critical)', days: '> 7 Days', lots: critAgingLots, pcs: critAgingPcs }
      ];

      agingTiers.forEach((tier) => {
        curRow++;
        const share = totalPieces > 0 ? (tier.pcs / totalPieces) * 100 : 0;
        const r = ws2.addRow([
          tier.name,
          tier.days,
          tier.lots,
          tier.pcs,
          `${share.toFixed(1)}%`
        ]);
        r.height = 19;
        r.eachCell((c, colIdx) => {
          c.border = thinBorder;
          c.alignment = { horizontal: 'center', vertical: 'middle' };
          c.font = { name: 'Segoe UI', size: 9 };
          if (colIdx === 1) {
            c.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
            c.font = { name: 'Segoe UI', size: 9, bold: true };
          }
          if (colIdx === 4) {
            c.numFmt = '#,##0';
            c.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FFDC2626' } };
          }
        });
      });

      // Column widths for Sheet 2
      ws2.columns = [
        { width: 28 },
        { width: 16 },
        { width: 16 },
        { width: 16 },
        { width: 16 },
        { width: 15 },
        { width: 15 }
      ];

      // ================= SHEET 3: APPLIED FILTERS & METADATA =================
      const ws3 = workbook.addWorksheet('Applied Filters', {
        views: [{ showGridLines: true }]
      });

      ws3.mergeCells('A1:D1');
      const fTitle = ws3.getCell('A1');
      fTitle.value = 'FACTORY SUITE PRO - REPORT PARAMETERS & METADATA';
      fTitle.font = { name: 'Segoe UI', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
      fTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
      fTitle.alignment = { horizontal: 'center', vertical: 'middle' };
      ws3.getRow(1).height = 28;

      const metadata = [
        ['Report Type', filters.holdLots ? 'Daily Folding - Hold Lots' : 'Daily Folding Production Report'],
        ['Generated At', new Date().toLocaleString('en-IN')],
        ['Total Lots Count', totalLots],
        ['Total Pieces Count', totalPieces],
        ['Completed Lots', completedLots],
        ['WIP Lots', wipLots],
        ['Not Started Lots', notStartedLots],
        ['Lot Number Search', filters.lotNumber || '(None)'],
        ['Folding Supervisor Filter', filters.supervisor || 'All Supervisors'],
        ['Garment Type Filter', filters.garmentType || 'All Garments'],
        ['Fabric Filter', filters.fabric || 'All Fabrics'],
        ['Brand Filter', filters.brand || 'All Brands'],
        ['Stitching Supervisor Filter', filters.stitchingSupervisor || 'All Stitching Supervisors'],
        ['Min Aging Filter', filters.minAging ? `${filters.minAging} Days` : 'None'],
        ['Max Aging Filter', filters.maxAging ? `${filters.maxAging} Days` : 'None'],
        ['Status Filter', filters.status || 'active'],
        ['Hold Lots Mode', filters.holdLots ? 'Yes (Hold Lots Only)' : 'No (All Lots)']
      ];

      metadata.forEach(([k, v]) => {
        const r = ws3.addRow([k, v]);
        r.height = 20;
        r.getCell(1).font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FF0F172A' } };
        r.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
        r.getCell(1).border = thinBorder;
        r.getCell(2).font = { name: 'Segoe UI', size: 9.5, color: { argb: 'FF334155' } };
        r.getCell(2).border = thinBorder;
      });

      ws3.columns = [{ width: 26 }, { width: 45 }, { width: 15 }, { width: 15 }];

      // Save file
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const fileDate = new Date().toISOString().slice(0, 10);
      const filename = filters.holdLots 
        ? `Folding_Hold_Lots_${fileDate}.xlsx` 
        : `Daily_Folding_Report_${fileDate}.xlsx`;
      saveAs(blob, filename);
    } catch (err) {
      console.error('Error generating Folding Report Excel:', err);
      alert('Failed to export Excel file. Please try again.');
    }
  };

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
      const stitchingSupMap = {};
      let goodAgingLots = 0, goodAgingPcs = 0;
      let warnAgingLots = 0, warnAgingPcs = 0;
      let critAgingLots = 0, critAgingPcs = 0;
      let completedLots = 0, completedPcs = 0;
      let wipLots = 0, wipPcs = 0;
      let notStartedLots = 0, notStartedPcs = 0;

      filteredData.forEach(item => {
        const pcs = parseInt(item['Total Pcs']) || 0;
        const gType = (item['Garment Type'] || 'Unknown').trim();
        const sup = (item['Folding Supervisor'] || 'Unassigned').trim();
        const stSup = (item['Stiching Supervisor'] || 'Unassigned').trim();
        const aging = calculateAging(item);
        const status = getLotStatus(item);

        if (!garmentMap[gType]) garmentMap[gType] = { totalLots: 0, totalPcs: 0 };
        garmentMap[gType].totalLots += 1;
        garmentMap[gType].totalPcs += pcs;

        if (!supervisorMap[sup]) supervisorMap[sup] = { totalLots: 0, totalPcs: 0 };
        supervisorMap[sup].totalLots += 1;
        supervisorMap[sup].totalPcs += pcs;

        if (!stitchingSupMap[stSup]) stitchingSupMap[stSup] = { totalLots: 0, totalPcs: 0 };
        stitchingSupMap[stSup].totalLots += 1;
        stitchingSupMap[stSup].totalPcs += pcs;

        if (aging <= 3) {
          goodAgingLots += 1;
          goodAgingPcs += pcs;
        } else if (aging <= 7) {
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

      const sortedStitchingSups = Object.keys(stitchingSupMap).map(name => ({
        name,
        totalLots: stitchingSupMap[name].totalLots,
        totalPcs: stitchingSupMap[name].totalPcs
      })).sort((a, b) => b.totalPcs - a.totalPcs);

      // 1. Main Header Block
      doc.setFillColor(15, 23, 42); // Dark Navy #0F172A
      doc.rect(15, 12, pageW - 30, 48, 'F');

      doc.setFontSize(14);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(255, 255, 255);
      const reportTitle = filters.holdLots 
        ? "FACTORY SUITE PRO - FOLDING HOLD LOTS REPORT" 
        : "FACTORY SUITE PRO - DAILY FOLDING PRODUCTION REPORT";
      doc.text(reportTitle, pageW / 2, 30, { align: 'center' });

      doc.setFontSize(8.5);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(199, 210, 254);
      const subText = `Total Lots: ${totalLots}   |   Total Pieces: ${totalPieces.toLocaleString()}   |   Completed: ${completedLots}   |   WIP: ${wipLots}   |   Not Started: ${notStartedLots}   |   Aging <=3d: ${goodAgingLots}   |   4-7d: ${warnAgingLots}   |   >7d: ${critAgingLots}`;
      doc.text(subText, pageW / 2, 48, { align: 'center' });

      // 2. Filter Banner
      doc.setFillColor(241, 245, 249);
      doc.rect(15, 63, pageW - 30, 16, 'F');
      doc.setFontSize(8);
      doc.setFont('helvetica', 'italic');
      doc.setTextColor(0, 0, 0);
      const filterSummary = `Filters: Status: ${filters.status || 'active'} | Supervisors: ${filters.supervisor || 'All'} | Garments: ${filters.garmentType || 'All'} | Fabrics: ${filters.fabric || 'All'} | Brands: ${filters.brand || 'All'} | Stitching Sups: ${filters.stitchingSupervisor || 'All'} | Aging: ${filters.minAging || 0} - ${filters.maxAging || 'Max'}d | Search: ${filters.lotNumber || 'None'}`;
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
        'Folding Date',
        'Supervisor',
        'Aging (Days)',
        'Status',
        'Recent Remarks',
        'Stitching Sup'
      ];

      const tableBody = filteredData.map((item, idx) => {
        const pcs = parseInt(item['Total Pcs']) || 0;
        const aging = calculateAging(item);
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
          item['Folding Date'] || '—',
          item['Folding Supervisor'] || '—',
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
              const days = calculateAging(item);
              if (days > 7) {
                data.cell.styles.fillColor = [239, 68, 68];
                data.cell.styles.textColor = [255, 255, 255];
                data.cell.styles.fontStyle = "bold";
              } else if (days >= 4) {
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

      const stBody = sortedStitchingSups.map(item => {
        const pct = totalPieces > 0 ? ((item.totalPcs / totalPieces) * 100).toFixed(1) : "0.0";
        return [item.name, item.totalLots.toString(), item.totalPcs.toLocaleString(), `${pct}%`];
      });
      stBody.push(["TOTAL", totalLots.toString(), totalPieces.toLocaleString(), "100.0%"]);

      const slaBody = [
        ["<= 3 Days (Good)", goodAgingLots.toString(), goodAgingPcs.toLocaleString(), `${totalLots > 0 ? ((goodAgingLots / totalLots) * 100).toFixed(1) : 0}%`],
        ["4 - 7 Days (Warning)", warnAgingLots.toString(), warnAgingPcs.toLocaleString(), `${totalLots > 0 ? ((warnAgingLots / totalLots) * 100).toFixed(1) : 0}%`],
        ["> 7 Days (Critical)", critAgingLots.toString(), critAgingPcs.toLocaleString(), `${totalLots > 0 ? ((critAgingLots / totalLots) * 100).toFixed(1) : 0}%`],
        ["TOTAL", totalLots.toString(), totalPieces.toLocaleString(), "100.0%"]
      ];

      const maxRows = Math.max(gBody.length, supBody.length, stBody.length, slaBody.length);
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
      const summarySub = `Total Lots: ${totalLots}   |   Total Pieces: ${totalPieces.toLocaleString()} Pcs   |   Folding Sups: ${sortedSupervisors.length}   |   Garments: ${sortedGarments.length}   |   Stitching Sups: ${sortedStitchingSups.length}`;
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
      doc.text("2. FOLDING SUPERVISOR BREAKDOWN", col2X, sectionTitleY);
      doc.text("3. STITCHING SUPERVISOR BREAKDOWN", col3X, sectionTitleY);
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

      // Column 2 Table: Folding Supervisor Breakdown
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

      // Column 3 Table: Stitching Supervisor Breakdown
      autoTable(doc, {
        head: [['Stitching Sup', 'Lots', 'Total Pcs', 'Share %']],
        body: stBody,
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
            if (data.row.index === stBody.length - 1) {
              data.cell.styles.fontStyle = 'bold';
              data.cell.styles.fillColor = [241, 245, 249];
            }
          }
        }
      });
      const endY3 = doc.lastAutoTable.finalY;

      // Column 4 Table: Aging & SLA Breakdown
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
        doc.text("Daily Folding Department Report — Factory Suite Pro", 15, finalY + 12);
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

      const fileName = filters.holdLots 
        ? `Folding_Hold_Lots_${new Date().toISOString().split('T')[0]}.pdf`
        : `Daily_Folding_Report_${new Date().toISOString().split('T')[0]}.pdf`;
      doc.save(fileName);

    } catch (error) {
      console.error('Error generating Folding Report PDF:', error);
      alert('Failed to generate PDF. Please try again.');
    }
  };

  const calculateTotals = () => {
    const uniqueSupervisors = new Set(
      filteredData.map(item => item['Folding Supervisor']).filter(Boolean)
    );
    
    return {
      totalLots: filteredData.length,
      totalPieces: filteredData.reduce((sum, item) => sum + (parseInt(item['Total Pcs']) || 0), 0),
      totalSupervisors: uniqueSupervisors.size,
      avgAging: filteredData.length > 0 ? 
        (filteredData.reduce((sum, item) => sum + parseInt(item['Aging'] || 0), 0) / filteredData.length).toFixed(1) : 0
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

  const getAgingColorClass = (days) => {
    const d = parseInt(days) || 0;
    if (d <= 3) return 'aging-pill-green';
    if (d <= 7) return 'aging-pill-amber';
    return 'aging-pill-red';
  };

  const getStatusColorClass = (status) => {
    switch(status) {
      case 'Completed': return 'status-pill-green';
      case 'WIP': return 'status-pill-blue';
      case 'Not Started': return 'status-pill-slate';
      default: return 'status-pill-slate';
    }
  };

  const totals = calculateTotals();
  const activeFiltersCount = Object.entries(filters).filter(([k, v]) => {
    if (k === 'status') return v !== 'active';
    if (k === 'holdLots') return v === true;
    return v !== '';
  }).length;

  if (loading) {
    return (
      <div className="dept-loading-container">
        <div className="dept-spinner"></div>
        <p className="dept-loading-text">Loading Daily Folding Report...</p>
        <span className="dept-loading-sub">Connecting to Factory Suite Real-Time Database</span>
      </div>
    );
  }

  if (error) {
    return (
      <div className="dept-error-container">
        <div className="dept-error-card">
          <div className="dept-error-icon">⚠️</div>
          <h3 className="dept-error-title">Failed to Load Folding Data</h3>
          <p className="dept-error-msg">{error}</p>
          <button onClick={fetchData} className="dept-btn dept-btn-primary">
            ↻ Retry Connection
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="dept-dashboard-container">
      <style>{`
        .dept-dashboard-container {
          padding: 24px 32px;
          background: #f8fafc;
          min-height: 100vh;
          font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
          color: #1e293b;
        }

        /* HEADER BANNER */
        .dept-header-card {
          background: linear-gradient(135deg, #0f4c81 0%, #1e3a8a 100%);
          border-radius: 16px;
          padding: 22px 28px;
          margin-bottom: 24px;
          box-shadow: 0 10px 25px -5px rgba(15, 76, 129, 0.25), 0 8px 10px -6px rgba(15, 76, 129, 0.2);
          display: flex;
          justify-content: space-between;
          align-items: center;
          flex-wrap: wrap;
          gap: 16px;
          color: white;
          position: relative;
          overflow: hidden;
        }

        .dept-header-card::after {
          content: "";
          position: absolute;
          top: -50%;
          right: -10%;
          width: 300px;
          height: 300px;
          background: radial-gradient(circle, rgba(255,255,255,0.1) 0%, rgba(255,255,255,0) 70%);
          pointer-events: none;
        }

        .dept-header-left {
          display: flex;
          align-items: center;
          gap: 16px;
        }

        .dept-header-icon {
          width: 52px;
          height: 52px;
          border-radius: 14px;
          background: rgba(255, 255, 255, 0.15);
          backdrop-filter: blur(8px);
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 26px;
          border: 1px solid rgba(255, 255, 255, 0.2);
        }

        .dept-header-title {
          font-size: 24px;
          font-weight: 800;
          letter-spacing: -0.02em;
          margin: 0 0 4px 0;
          display: flex;
          align-items: center;
          gap: 10px;
        }

        .dept-badge-pill {
          font-size: 11px;
          font-weight: 700;
          padding: 3px 10px;
          border-radius: 20px;
          background: rgba(255, 255, 255, 0.2);
          border: 1px solid rgba(255, 255, 255, 0.3);
          letter-spacing: 0.05em;
          text-transform: uppercase;
        }

        .dept-header-subtitle {
          margin: 0;
          font-size: 13px;
          color: #bfdbfe;
          font-weight: 500;
        }

        .dept-header-actions {
          display: flex;
          align-items: center;
          gap: 10px;
          flex-wrap: wrap;
        }

        .dept-btn {
          display: inline-flex;
          align-items: center;
          gap: 6px;
          padding: 9px 16px;
          border-radius: 10px;
          font-size: 13px;
          font-weight: 700;
          cursor: pointer;
          transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
          border: none;
          outline: none;
          box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }

        .dept-btn:hover {
          transform: translateY(-2px);
          box-shadow: 0 4px 10px rgba(0,0,0,0.15);
        }

        .dept-btn-back {
          background: rgba(255, 255, 255, 0.15);
          color: white;
          border: 1px solid rgba(255, 255, 255, 0.25);
        }
        .dept-btn-back:hover {
          background: rgba(255, 255, 255, 0.25);
        }

        .dept-btn-excel {
          background: #10b981;
          color: white;
        }
        .dept-btn-excel:hover {
          background: #059669;
        }

        .dept-btn-pdf {
          background: #ef4444;
          color: white;
        }
        .dept-btn-pdf:hover {
          background: #dc2626;
        }

        .dept-btn-refresh {
          background: rgba(255, 255, 255, 0.15);
          color: white;
          border: 1px solid rgba(255, 255, 255, 0.25);
        }
        .dept-btn-refresh:hover {
          background: rgba(255, 255, 255, 0.25);
        }

        .dept-btn-clear {
          background: rgba(239, 68, 68, 0.2);
          color: #fca5a5;
          border: 1px solid rgba(239, 68, 68, 0.4);
        }
        .dept-btn-clear:hover {
          background: rgba(239, 68, 68, 0.35);
          color: white;
        }

        /* METRICS / STATS CARDS */
        .dept-stats-grid {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(230px, 1fr));
          gap: 16px;
          margin-bottom: 24px;
        }

        .dept-stat-card {
          background: white;
          border-radius: 14px;
          padding: 18px 20px;
          border: 1px solid #e2e8f0;
          box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
          display: flex;
          align-items: center;
          gap: 16px;
          transition: all 0.25s ease;
          position: relative;
          overflow: hidden;
        }

        .dept-stat-card:hover {
          transform: translateY(-3px);
          box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.08);
          border-color: #cbd5e1;
        }

        .dept-stat-icon-wrapper {
          width: 48px;
          height: 48px;
          border-radius: 12px;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 22px;
          flex-shrink: 0;
        }

        .dept-stat-content {
          flex: 1;
        }

        .dept-stat-label {
          font-size: 12px;
          font-weight: 700;
          color: #64748b;
          text-transform: uppercase;
          letter-spacing: 0.04em;
          margin-bottom: 2px;
        }

        .dept-stat-val {
          font-size: 24px;
          font-weight: 800;
          color: #0f172a;
          line-height: 1.2;
          margin: 0;
        }

        .dept-stat-sub {
          font-size: 11px;
          font-weight: 600;
          margin-top: 3px;
        }

        .card-blue .dept-stat-icon-wrapper { background: #eff6ff; color: #2563eb; }
        .card-blue .dept-stat-sub { color: #3b82f6; }
        
        .card-emerald .dept-stat-icon-wrapper { background: #ecfdf5; color: #059669; }
        .card-emerald .dept-stat-sub { color: #10b981; }

        .card-violet .dept-stat-icon-wrapper { background: #f5f3ff; color: #7c3aed; }
        .card-violet .dept-stat-sub { color: #8b5cf6; }

        .card-amber .dept-stat-icon-wrapper { background: #fffbeb; color: #d97706; }
        .card-amber .dept-stat-sub { color: #f59e0b; }

        /* FILTER CONSOLE */
        .dept-filter-card {
          background: white;
          border-radius: 16px;
          padding: 20px 24px;
          border: 1px solid #e2e8f0;
          box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
          margin-bottom: 24px;
        }

        .dept-filter-top {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-bottom: 16px;
          padding-bottom: 12px;
          border-bottom: 1px solid #f1f5f9;
          flex-wrap: wrap;
          gap: 12px;
        }

        .dept-filter-title-group {
          display: flex;
          align-items: center;
          gap: 10px;
        }

        .dept-filter-title {
          font-size: 15px;
          font-weight: 800;
          color: #1e293b;
          display: flex;
          align-items: center;
          gap: 6px;
          margin: 0;
        }

        .dept-filter-counter {
          font-size: 12px;
          font-weight: 700;
          padding: 2px 8px;
          border-radius: 12px;
          background: #f1f5f9;
          color: #475569;
        }

        /* STATUS SEGMENTED CONTROL */
        .dept-status-segments {
          display: flex;
          align-items: center;
          gap: 6px;
          background: #f8fafc;
          padding: 4px;
          border-radius: 12px;
          border: 1px solid #e2e8f0;
          flex-wrap: wrap;
        }

        .dept-segment-btn {
          border: none;
          background: transparent;
          padding: 6px 12px;
          border-radius: 8px;
          font-size: 12px;
          font-weight: 700;
          color: #64748b;
          cursor: pointer;
          transition: all 0.2s ease;
        }

        .dept-segment-btn:hover {
          color: #0f172a;
          background: rgba(0,0,0,0.04);
        }

        .dept-segment-btn.active {
          background: white;
          color: #0f4c81;
          box-shadow: 0 2px 4px rgba(0,0,0,0.06);
        }

        .dept-segment-hold {
          border: 1px solid #fecaca !important;
          background: #fff1f2 !important;
          color: #e11d48 !important;
        }

        .dept-segment-hold.active {
          background: #e11d48 !important;
          color: white !important;
          box-shadow: 0 2px 6px rgba(225, 29, 72, 0.3) !important;
        }

        /* FILTER FORM GRID */
        .dept-filter-grid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(190px, 1fr));
          gap: 12px;
        }

        .dept-input-group {
          display: flex;
          flex-direction: column;
          gap: 5px;
        }

        .dept-input-label {
          font-size: 11px;
          font-weight: 700;
          color: #64748b;
          text-transform: uppercase;
          letter-spacing: 0.03em;
        }

        .dept-form-control {
          width: 100%;
          padding: 8px 12px;
          border: 1.5px solid #e2e8f0;
          border-radius: 8px;
          font-size: 13px;
          font-weight: 500;
          color: #1e293b;
          background: #ffffff;
          transition: all 0.2s ease;
          outline: none;
          box-sizing: border-box;
        }

        .dept-form-control:focus {
          border-color: #0f4c81;
          box-shadow: 0 0 0 3px rgba(15, 76, 129, 0.12);
        }

        /* DATA TABLE CARD */
        .dept-table-card {
          background: white;
          border-radius: 16px;
          border: 1px solid #e2e8f0;
          box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
          overflow: hidden;
          margin-bottom: 30px;
        }

        .dept-table-top-bar {
          padding: 16px 24px;
          display: flex;
          justify-content: space-between;
          align-items: center;
          border-bottom: 1px solid #e2e8f0;
          background: #ffffff;
          flex-wrap: wrap;
          gap: 10px;
        }

        .dept-table-heading {
          font-size: 16px;
          font-weight: 800;
          color: #0f172a;
          margin: 0;
          display: flex;
          align-items: center;
          gap: 8px;
        }

        .dept-table-records-info {
          font-size: 13px;
          color: #64748b;
          font-weight: 600;
        }

        .dept-table-wrapper {
          overflow-x: auto;
          max-height: calc(100vh - 300px);
          min-height: 380px;
          position: relative;
        }

        .dept-table {
          width: 100%;
          border-collapse: collapse;
          min-width: 1400px;
          text-align: center;
        }

        .dept-table thead {
          position: sticky;
          top: 0;
          z-index: 10;
        }

        .dept-table th {
          background: #0f4c81 !important;
          color: #ffffff !important;
          font-size: 12px;
          font-weight: 800;
          text-transform: uppercase;
          letter-spacing: 0.05em;
          padding: 12px 10px;
          border: 1px solid #1e3a8a;
          white-space: nowrap;
        }

        .dept-table tbody tr {
          border-bottom: 1px solid #f1f5f9;
          transition: background-color 0.15s ease;
          cursor: pointer;
        }

        .dept-table tbody tr:nth-child(even) {
          background-color: #f8fafc;
        }

        .dept-table tbody tr:hover {
          background-color: #eff6ff !important;
        }

        .dept-table td {
          padding: 11px 10px;
          font-size: 13px;
          color: #1e293b;
          border: 1px solid #e2e8f0;
          white-space: nowrap;
          vertical-align: middle;
        }

        .cell-lot-no {
          font-weight: 800;
          color: #dc2626 !important;
          font-size: 13.5px;
        }

        .cell-total-pcs {
          font-weight: 800;
          color: #1e40af;
          font-size: 13px;
        }

        /* PILLS & BADGES */
        .dept-tag {
          display: inline-block;
          padding: 3px 8px;
          border-radius: 6px;
          font-size: 11.5px;
          font-weight: 700;
        }

        .dept-tag-brand {
          background: #fef3c7;
          color: #92400e;
        }

        .dept-tag-sup {
          background: #eff6ff;
          color: #1e40af;
          border: 1px solid #dbeafe;
        }

        .dept-tag-stitching-sup {
          background: #f0fdf4;
          color: #166534;
          border: 1px solid #dcfce7;
        }

        .aging-pill-green {
          background: #dcfce7;
          color: #15803d;
          padding: 3px 10px;
          border-radius: 20px;
          font-weight: 800;
          font-size: 12px;
        }

        .aging-pill-amber {
          background: #fef3c7;
          color: #b45309;
          padding: 3px 10px;
          border-radius: 20px;
          font-weight: 800;
          font-size: 12px;
        }

        .aging-pill-red {
          background: #fee2e2;
          color: #b91c1c;
          padding: 3px 10px;
          border-radius: 20px;
          font-weight: 800;
          font-size: 12px;
        }

        .status-pill-green {
          background: #dcfce7;
          color: #15803d;
          padding: 3px 10px;
          border-radius: 20px;
          font-weight: 800;
          font-size: 11.5px;
        }

        .status-pill-blue {
          background: #dbeafe;
          color: #1d4ed8;
          padding: 3px 10px;
          border-radius: 20px;
          font-weight: 800;
          font-size: 11.5px;
        }

        .status-pill-slate {
          background: #f1f5f9;
          color: #475569;
          padding: 3px 10px;
          border-radius: 20px;
          font-weight: 700;
          font-size: 11.5px;
        }

        .dept-remarks-cell {
          max-width: 220px;
          overflow: hidden;
          text-overflow: ellipsis;
          white-space: nowrap;
          font-size: 12px;
          color: #475569;
          display: inline-block;
        }

        .dept-remarks-hold {
          color: #dc2626 !important;
          font-weight: 700;
        }

        .dept-hold-badge {
          background: #fee2e2;
          color: #dc2626;
          padding: 2px 6px;
          border-radius: 4px;
          font-size: 10.5px;
          font-weight: 800;
          margin-right: 4px;
        }

        /* TABLE FOOTER */
        .dept-table tfoot td {
          background: #f1f5f9 !important;
          font-weight: 800;
          color: #0f4c81;
          padding: 12px 10px;
          border-top: 2px solid #cbd5e1;
        }

        /* EMPTY STATE */
        .dept-empty-state {
          padding: 60px 20px;
          text-align: center;
          color: #64748b;
        }
        .dept-empty-icon {
          font-size: 44px;
          margin-bottom: 12px;
        }
        .dept-empty-text {
          font-size: 16px;
          font-weight: 700;
          color: #1e293b;
          margin-bottom: 6px;
        }
        .dept-empty-sub {
          font-size: 13px;
          margin-bottom: 16px;
        }

        /* MODAL */
        .dept-modal-backdrop {
          position: fixed;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          background: rgba(15, 23, 42, 0.65);
          backdrop-filter: blur(6px);
          display: flex;
          align-items: center;
          justify-content: center;
          z-index: 9999;
          padding: 20px;
        }

        .dept-modal-card {
          background: white;
          border-radius: 20px;
          width: 100%;
          max-width: 780px;
          max-height: 90vh;
          overflow-y: auto;
          box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.35);
          animation: deptModalSlide 0.25s ease-out;
        }

        @keyframes deptModalSlide {
          from { opacity: 0; transform: translateY(16px) scale(0.98); }
          to { opacity: 1; transform: translateY(0) scale(1); }
        }

        .dept-modal-header {
          background: linear-gradient(135deg, #0f4c81 0%, #1e3a8a 100%);
          color: white;
          padding: 20px 26px;
          display: flex;
          justify-content: space-between;
          align-items: center;
        }

        .dept-modal-title {
          font-size: 20px;
          font-weight: 800;
          margin: 0;
          display: flex;
          align-items: center;
          gap: 10px;
        }

        .dept-modal-close {
          background: rgba(255, 255, 255, 0.2);
          border: none;
          color: white;
          font-size: 22px;
          width: 36px;
          height: 36px;
          border-radius: 50%;
          cursor: pointer;
          display: flex;
          align-items: center;
          justify-content: center;
          transition: background 0.2s;
        }
        .dept-modal-close:hover {
          background: rgba(255, 255, 255, 0.35);
        }

        .dept-modal-body {
          padding: 24px;
        }

        .dept-modal-grid {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
          gap: 14px;
          margin-bottom: 24px;
        }

        .dept-modal-cell {
          background: #f8fafc;
          border: 1px solid #e2e8f0;
          border-radius: 12px;
          padding: 12px 16px;
        }

        .dept-modal-cell-label {
          font-size: 11px;
          font-weight: 700;
          color: #64748b;
          text-transform: uppercase;
          margin-bottom: 4px;
        }

        .dept-modal-cell-val {
          font-size: 15px;
          font-weight: 800;
          color: #0f172a;
        }

        .dept-history-title {
          font-size: 16px;
          font-weight: 800;
          color: #1e293b;
          margin: 0 0 16px 0;
          padding-bottom: 8px;
          border-bottom: 2px solid #e2e8f0;
        }

        .dept-timeline {
          position: relative;
          padding-left: 24px;
          border-left: 2px solid #e2e8f0;
          margin-left: 12px;
        }

        .dept-timeline-item {
          position: relative;
          margin-bottom: 18px;
        }

        .dept-timeline-dot {
          position: absolute;
          left: -31px;
          top: 4px;
          width: 12px;
          height: 12px;
          border-radius: 50%;
          background: #0f4c81;
          border: 2px solid white;
          box-shadow: 0 0 0 2px #0f4c81;
        }

        .dept-timeline-card {
          background: #ffffff;
          border: 1px solid #e2e8f0;
          border-radius: 10px;
          padding: 12px 16px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.03);
        }

        .dept-timeline-header {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-bottom: 6px;
        }

        .dept-timeline-status {
          font-weight: 800;
          font-size: 13.5px;
          color: #0f4c81;
        }

        .dept-timeline-time {
          font-size: 11.5px;
          color: #64748b;
          font-weight: 500;
        }

        .dept-timeline-remarks {
          font-size: 13px;
          color: #334155;
          margin: 0;
        }

        /* LOADING & ERROR */
        .dept-loading-container, .dept-error-container {
          min-height: 80vh;
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          text-align: center;
          padding: 20px;
        }

        .dept-spinner {
          width: 50px;
          height: 50px;
          border: 4px solid #e2e8f0;
          border-top-color: #0f4c81;
          border-radius: 50%;
          animation: deptSpin 0.8s linear infinite;
          margin-bottom: 16px;
        }

        @keyframes deptSpin {
          to { transform: rotate(360deg); }
        }

        .dept-loading-text {
          font-size: 18px;
          font-weight: 800;
          color: #0f172a;
          margin: 0 0 6px 0;
        }

        .dept-loading-sub {
          font-size: 13px;
          color: #64748b;
        }

        .dept-error-card {
          background: white;
          padding: 36px;
          border-radius: 16px;
          border: 1px solid #fee2e2;
          box-shadow: 0 10px 25px -5px rgba(220, 38, 38, 0.1);
          max-width: 440px;
        }

        .dept-error-icon {
          font-size: 48px;
          margin-bottom: 12px;
        }

        .dept-error-title {
          font-size: 18px;
          font-weight: 800;
          color: #991b1b;
          margin: 0 0 8px 0;
        }

        .dept-error-msg {
          font-size: 13px;
          color: #64748b;
          margin: 0 0 20px 0;
        }
      `}</style>

      {/* HEADER HERO */}
      <div className="dept-header-card">
        <div className="dept-header-left">
          <div className="dept-header-icon">📦</div>
          <div>
            <h1 className="dept-header-title">
              DAILY FOLDING DASHBOARD
              <span className="dept-badge-pill">
                {filters.holdLots ? 'Hold Lots Only' : filters.status === 'active' ? 'Active Lots (WIP)' : filters.status}
              </span>
            </h1>
            <p className="dept-header-subtitle">
              Live tracking of lots, piece counts, aging analysis, and department supervisor allocation
            </p>
          </div>
        </div>

        <div className="dept-header-actions">
          <button 
            className="dept-btn dept-btn-back"
            onClick={() => window.history.back()}
            title="Go Back"
          >
            ← Back
          </button>
          
          <button 
            className="dept-btn dept-btn-excel"
            onClick={exportToExcel}
            title="Export full data to Excel"
          >
            📊 Export Excel
          </button>

          <button 
            className="dept-btn dept-btn-pdf"
            onClick={exportToPDF}
            title="Download PDF report (Stitching style layout)"
          >
            📄 Download PDF
          </button>

          <button 
            className="dept-btn dept-btn-refresh"
            onClick={fetchData}
            title="Refresh latest data"
          >
            ↻ Refresh
          </button>

          {activeFiltersCount > 0 && (
            <button 
              className="dept-btn dept-btn-clear"
              onClick={clearFilters}
              title="Reset all active filters"
            >
              ✕ Reset ({activeFiltersCount})
            </button>
          )}
        </div>
      </div>

      {/* EXECUTIVE KPI STATS */}
      <div className="dept-stats-grid">
        <div className="dept-stat-card card-blue">
          <div className="dept-stat-icon-wrapper">📋</div>
          <div className="dept-stat-content">
            <div className="dept-stat-label">Total Lots</div>
            <p className="dept-stat-val">{totals.totalLots}</p>
            <div className="dept-stat-sub">
              {filteredData.filter(i => getLotStatus(i) === 'WIP').length} WIP • {filteredData.filter(i => getLotStatus(i) === 'Completed').length} Completed
            </div>
          </div>
        </div>

        <div className="dept-stat-card card-emerald">
          <div className="dept-stat-icon-wrapper">👕</div>
          <div className="dept-stat-content">
            <div className="dept-stat-label">Total Pieces</div>
            <p className="dept-stat-val">{totals.totalPieces.toLocaleString()}</p>
            <div className="dept-stat-sub">
              Avg {(totals.totalPieces / (totals.totalLots || 1)).toFixed(0)} pcs/lot
            </div>
          </div>
        </div>

        <div className="dept-stat-card card-violet">
          <div className="dept-stat-icon-wrapper">👤</div>
          <div className="dept-stat-content">
            <div className="dept-stat-label">Folding Supervisors</div>
            <p className="dept-stat-val">{totals.totalSupervisors}</p>
            <div className="dept-stat-sub">Active in Department</div>
          </div>
        </div>

        <div className="dept-stat-card card-amber">
          <div className="dept-stat-icon-wrapper">⏱️</div>
          <div className="dept-stat-content">
            <div className="dept-stat-label">Average Aging</div>
            <p className="dept-stat-val">{totals.avgAging} <span style={{ fontSize: '14px', fontWeight: 600 }}>Days</span></p>
            <div className="dept-stat-sub">Processing timeline</div>
          </div>
        </div>
      </div>

      {/* FILTER CONSOLE */}
      <div className="dept-filter-card">
        <div className="dept-filter-top">
          <div className="dept-filter-title-group">
            <h3 className="dept-filter-title">🔍 Filter & Search Console</h3>
            <span className="dept-filter-counter">{filteredData.length} lots matching</span>
          </div>

          <div className="dept-status-segments">
            <button 
              className={`dept-segment-btn ${filters.status === 'active' && !filters.holdLots ? 'active' : ''}`}
              onClick={() => { handleFilterChange('status', 'active'); handleFilterChange('holdLots', false); }}
            >
              Active Lots
            </button>
            <button 
              className={`dept-segment-btn ${filters.status === 'all' && !filters.holdLots ? 'active' : ''}`}
              onClick={() => { handleFilterChange('status', 'all'); handleFilterChange('holdLots', false); }}
            >
              All Status
            </button>
            <button 
              className={`dept-segment-btn ${filters.status === 'WIP' && !filters.holdLots ? 'active' : ''}`}
              onClick={() => { handleFilterChange('status', 'WIP'); handleFilterChange('holdLots', false); }}
            >
              WIP Only
            </button>
            <button 
              className={`dept-segment-btn ${filters.status === 'Completed' && !filters.holdLots ? 'active' : ''}`}
              onClick={() => { handleFilterChange('status', 'Completed'); handleFilterChange('holdLots', false); }}
            >
              Completed
            </button>
            <button 
              className={`dept-segment-btn ${filters.status === 'Not Started' && !filters.holdLots ? 'active' : ''}`}
              onClick={() => { handleFilterChange('status', 'Not Started'); handleFilterChange('holdLots', false); }}
            >
              Not Started
            </button>
            <button 
              className={`dept-segment-btn dept-segment-hold ${filters.holdLots ? 'active' : ''}`}
              onClick={() => handleFilterChange('holdLots', !filters.holdLots)}
            >
              ⏸️ Hold Lots Only
            </button>
          </div>
        </div>

        <div className="dept-filter-grid">
          <div className="dept-input-group">
            <label className="dept-input-label">Lot Number</label>
            <input 
              type="text"
              className="dept-form-control"
              placeholder="Search lot number..."
              value={filters.lotNumber}
              onChange={(e) => handleFilterChange('lotNumber', e.target.value)}
            />
          </div>

          <div className="dept-input-group">
            <label className="dept-input-label">Folding Supervisor</label>
            <select
              className="dept-form-control"
              value={filters.supervisor}
              onChange={(e) => handleFilterChange('supervisor', e.target.value)}
            >
              <option value="">All Supervisors</option>
              {getUniqueValues('Folding Supervisor').map((s, idx) => (
                <option key={idx} value={s}>{s}</option>
              ))}
            </select>
          </div>

          <div className="dept-input-group">
            <label className="dept-input-label">Stitching Supervisor</label>
            <select
              className="dept-form-control"
              value={filters.stitchingSupervisor}
              onChange={(e) => handleFilterChange('stitchingSupervisor', e.target.value)}
            >
              <option value="">All Stitching Sups</option>
              {getUniqueValues('Stiching Supervisor').map((s, idx) => (
                <option key={idx} value={s}>{s}</option>
              ))}
            </select>
          </div>

          <div className="dept-input-group">
            <label className="dept-input-label">Garment Type</label>
            <select
              className="dept-form-control"
              value={filters.garmentType}
              onChange={(e) => handleFilterChange('garmentType', e.target.value)}
            >
              <option value="">All Garment Types</option>
              {getUniqueValues('Garment Type').map((g, idx) => (
                <option key={idx} value={g}>{g}</option>
              ))}
            </select>
          </div>

          <div className="dept-input-group">
            <label className="dept-input-label">Fabric</label>
            <select
              className="dept-form-control"
              value={filters.fabric}
              onChange={(e) => handleFilterChange('fabric', e.target.value)}
            >
              <option value="">All Fabrics</option>
              {getUniqueValues('Fabric').map((f, idx) => (
                <option key={idx} value={f}>{f}</option>
              ))}
            </select>
          </div>

          <div className="dept-input-group">
            <label className="dept-input-label">Brand</label>
            <select
              className="dept-form-control"
              value={filters.brand}
              onChange={(e) => handleFilterChange('brand', e.target.value)}
            >
              <option value="">All Brands</option>
              {getUniqueValues('BRAND').map((b, idx) => (
                <option key={idx} value={b}>{b}</option>
              ))}
            </select>
          </div>

          <div className="dept-input-group">
            <label className="dept-input-label">Min Aging (Days)</label>
            <input 
              type="number"
              className="dept-form-control"
              placeholder="e.g. 3"
              value={filters.minAging}
              onChange={(e) => handleFilterChange('minAging', e.target.value)}
            />
          </div>

          <div className="dept-input-group">
            <label className="dept-input-label">Max Aging (Days)</label>
            <input 
              type="number"
              className="dept-form-control"
              placeholder="e.g. 10"
              value={filters.maxAging}
              onChange={(e) => handleFilterChange('maxAging', e.target.value)}
            />
          </div>
        </div>
      </div>

      {/* DATA TABLE */}
      <div className="dept-table-card">
        <div className="dept-table-top-bar">
          <h3 className="dept-table-heading">
            📋 Folding Lots Register
            {filters.holdLots && <span className="dept-hold-badge">HOLD LOTS ONLY</span>}
          </h3>
          <span className="dept-table-records-info">
            Displaying {filteredData.length} of {data.length} total lots
          </span>
        </div>

        <div className="dept-table-wrapper">
          <table className="dept-table">
            <thead>
              <tr>
                <th>Sr No.</th>
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
                <th>Folding Date</th>
                <th>Supervisor</th>
                <th>Aging</th>
                <th>Status</th>
                <th>Recent Remarks</th>
                <th>Stitching Sup</th>
              </tr>
            </thead>
            <tbody>
              {filteredData.length > 0 ? (
                filteredData.map((item, index) => {
                  const isHold = isLotOnHold(item);
                  const status = getLotStatus(item);
                  const remarks = getRecentRemarks(item);

                  return (
                    <tr key={index} onClick={() => openRowDetails(item)}>
                      <td style={{ color: '#94a3b8', fontWeight: 600 }}>{index + 1}</td>
                      <td className="cell-lot-no">{item['Lot Number'] || '—'}</td>
                      <td>{item['Garment Type'] || '—'}</td>
                      <td style={{ fontWeight: 600 }}>{item['Style'] || '—'}</td>
                      <td style={{ fontWeight: 600 }}>{item['Fabric'] || '—'}</td>
                      <td>
                        {item['BRAND'] ? (
                          <span className="dept-tag dept-tag-brand">{item['BRAND']}</span>
                        ) : '—'}
                      </td>
                      <td className="cell-total-pcs">
                        {(parseInt(item['Total Pcs']) || 0).toLocaleString()}
                      </td>
                      <td>{item['Section'] || '—'}</td>
                      <td>{item['Season'] || '—'}</td>
                      <td>{item['Party Name'] || '—'}</td>
                      <td>{item['Direct Stitching'] || '—'}</td>
                      <td style={{ fontWeight: 600 }}>{item['Folding Date'] || '—'}</td>
                      <td>
                        {item['Folding Supervisor'] ? (
                          <span className="dept-tag dept-tag-sup">{item['Folding Supervisor']}</span>
                        ) : '—'}
                      </td>
                      <td>
                        <span className={getAgingColorClass(item['Aging'])}>
                          {item['Aging']}d
                        </span>
                      </td>
                      <td>
                        <span className={getStatusColorClass(status)}>
                          {status}
                        </span>
                      </td>
                      <td>
                        <div className={`dept-remarks-cell ${isHold ? 'dept-remarks-hold' : ''}`} title={remarks}>
                          {isHold && <span className="dept-hold-badge">HOLD</span>}
                          {remarks}
                        </div>
                      </td>
                      <td>
                        {item['Stiching Supervisor'] ? (
                          <span className="dept-tag dept-tag-stitching-sup">{item['Stiching Supervisor']}</span>
                        ) : '—'}
                      </td>
                    </tr>
                  );
                })
              ) : (
                <tr>
                  <td colSpan="13">
                    <div className="dept-empty-state">
                      <div className="dept-empty-icon">📭</div>
                      <div className="dept-empty-text">No Folding Lots Found</div>
                      <div className="dept-empty-sub">Try changing or clearing your active filters</div>
                      <button className="dept-btn dept-btn-clear" onClick={clearFilters}>
                        Clear All Filters
                      </button>
                    </div>
                  </td>
                </tr>
              )}
            </tbody>
            {filteredData.length > 0 && (
              <tfoot>
                <tr>
                  <td></td>
                  <td style={{ color: '#0f4c81' }}>TOTAL</td>
                  <td></td>
                  <td></td>
                  <td></td>
                  <td style={{ color: '#0f4c81' }}>TOTAL PCS:</td>
                  <td style={{ color: '#dc2626', fontSize: '14px' }}>
                    {totals.totalPieces.toLocaleString()}
                  </td>
                  <td></td>
                  <td>{totals.totalSupervisors} Sups</td>
                  <td>Avg {totals.avgAging}d</td>
                  <td>{filteredData.length} Lots</td>
                  <td></td>
                  <td></td>
                </tr>
              </tfoot>
            )}
          </table>
        </div>
      </div>

      {/* DETAIL MODAL */}
      {modalOpen && selectedRow && (
        <div className="dept-modal-backdrop" onClick={closeModal}>
          <div className="dept-modal-card" onClick={(e) => e.stopPropagation()}>
            <div className="dept-modal-header">
              <h2 className="dept-modal-title">
                <span>📦</span> Lot #{selectedRow['Lot Number']} Details
              </h2>
              <button className="dept-modal-close" onClick={closeModal}>×</button>
            </div>

            <div className="dept-modal-body">
              <div className="dept-modal-grid">
                <div className="dept-modal-cell">
                  <div className="dept-modal-cell-label">Garment Type</div>
                  <div className="dept-modal-cell-val">{selectedRow['Garment Type'] || '—'}</div>
                </div>
                <div className="dept-modal-cell">
                  <div className="dept-modal-cell-label">Fabric</div>
                  <div className="dept-modal-cell-val">{selectedRow['Fabric'] || '—'}</div>
                </div>
                <div className="dept-modal-cell">
                  <div className="dept-modal-cell-label">Style</div>
                  <div className="dept-modal-cell-val">{selectedRow['Style'] || '—'}</div>
                </div>
                <div className="dept-modal-cell">
                  <div className="dept-modal-cell-label">Brand</div>
                  <div className="dept-modal-cell-val">{selectedRow['BRAND'] || '—'}</div>
                </div>
                <div className="dept-modal-cell">
                  <div className="dept-modal-cell-label">Total Pieces</div>
                  <div className="dept-modal-cell-val" style={{ color: '#1e40af' }}>
                    {(parseInt(selectedRow['Total Pcs']) || 0).toLocaleString()}
                  </div>
                </div>
                <div className="dept-modal-cell">
                  <div className="dept-modal-cell-label">Folding Date</div>
                  <div className="dept-modal-cell-val">{selectedRow['Folding Date'] || '—'}</div>
                </div>
                <div className="dept-modal-cell">
                  <div className="dept-modal-cell-label">Folding Supervisor</div>
                  <div className="dept-modal-cell-val">{selectedRow['Folding Supervisor'] || '—'}</div>
                </div>
                <div className="dept-modal-cell">
                  <div className="dept-modal-cell-label">Stitching Supervisor</div>
                  <div className="dept-modal-cell-val">{selectedRow['Stiching Supervisor'] || '—'}</div>
                </div>
                <div className="dept-modal-cell">
                  <div className="dept-modal-cell-label">Aging (Days)</div>
                  <div className="dept-modal-cell-val">{selectedRow['Aging']} Days</div>
                </div>
                <div className="dept-modal-cell">
                  <div className="dept-modal-cell-label">Hold Status</div>
                  <div className="dept-modal-cell-val" style={{ color: isLotOnHold(selectedRow) ? '#dc2626' : '#16a34a' }}>
                    {isLotOnHold(selectedRow) ? '⏸️ On Hold' : '✅ Normal'}
                  </div>
                </div>
              </div>

              <h3 className="dept-history-title">⏱️ Status History & Timeline</h3>
              {getAllStatusHistory(selectedRow).length > 0 ? (
                <div className="dept-timeline">
                  {getAllStatusHistory(selectedRow).map((status, idx) => (
                    <div className="dept-timeline-item" key={idx}>
                      <div className="dept-timeline-dot"></div>
                      <div className="dept-timeline-card">
                        <div className="dept-timeline-header">
                          <span className="dept-timeline-status">{status.status || 'Status Update'}</span>
                          <span className="dept-timeline-time">{new Date(status.timestamp).toLocaleString()}</span>
                        </div>
                        <p className="dept-timeline-remarks">
                          {status.remarks ? `Remarks: ${status.remarks}` : 'No remarks'}
                          {status.supervisor && ` • By: ${status.supervisor}`}
                        </p>
                      </div>
                    </div>
                  ))}
                </div>
              ) : (
                <p style={{ color: '#64748b', fontSize: '13px', textAlign: 'center', padding: '20px 0' }}>
                  No historical log entries available for this lot.
                </p>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default DailyFoldingReport;