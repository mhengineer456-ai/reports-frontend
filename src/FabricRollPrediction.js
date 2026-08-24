// src/pages/FabricRollPrediction.jsx
import React, { useState, useEffect, useCallback } from 'react';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import './FabricRollPrediction.css';

export default function FabricRollPrediction() {
  const [isLoading, setIsLoading] = useState(true);
  const [selectedGarmentType, setSelectedGarmentType] = useState('');
  const [garmentTypes, setGarmentTypes] = useState([]);
  const [selectedSection, setSelectedSection] = useState('');
  const [sections, setSections] = useState([]);
  const [selectedSeason, setSelectedSeason] = useState('');
  const [seasons, setSeasons] = useState([]);
  const [selectedFabric, setSelectedFabric] = useState('');
  const [fabrics, setFabrics] = useState([]);
  const [allStyles, setAllStyles] = useState([]);
  const [filteredStyles, setFilteredStyles] = useState([]);
  const [predictionData, setPredictionData] = useState(null);
  const [error, setError] = useState(null);
  const [cutLotNumbers, setCutLotNumbers] = useState(new Set());

  // Google Sheets configuration
  const SHEET_ID = '1fKSwGBIpzWEFk566WRQ4bzQ0anJlmasoY8TwrTLQHXI';
  const API_KEY = 'AIzaSyAomDFBkOySlIxKWSKGHe6ATv9gvaBr7uk';
  const SHEET_NAME = 'JobOrder';
  
  // Cut sheet configuration
  const CUT_SHEET_ID = '1Hj3JeJEKB43aYYWv8gk2UhdU6BWuEQfCg5pBlTdBMNA';
  const CUT_SHEET_NAME = 'Index';

  // Parse shade column: "Black[1], off-white[2], olive[2]"
  const parseShades = (shadeString) => {
    if (!shadeString) return [];
    
    const shades = [];
    const pattern = /([A-Za-z\s\.\-]+?)\s*\[(\d+)\]/g;
    let match;
    
    while ((match = pattern.exec(shadeString)) !== null) {
      shades.push({
        color: match[1].trim(),
        rollsRequired: parseInt(match[2])
      });
    }
    
    if (shades.length === 0 && shadeString.includes(',')) {
      const parts = shadeString.split(',');
      parts.forEach(part => {
        const trimmed = part.trim();
        shades.push({
          color: trimmed,
          rollsRequired: 1
        });
      });
    } else if (shades.length === 0 && shadeString.trim()) {
      shades.push({
        color: shadeString.trim(),
        rollsRequired: 1
      });
    }
    
    return shades;
  };

  // Fetch cut lot numbers from the Index sheet
  const fetchCutLotNumbers = useCallback(async () => {
    try {
      console.log('Fetching cut sheet from ID:', CUT_SHEET_ID);
      
      const response = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${CUT_SHEET_ID}/values/${CUT_SHEET_NAME}?key=${API_KEY}`
      );
      
      if (!response.ok) {
        console.warn('Could not fetch cut sheet, proceeding without filtering', response.status);
        return new Set();
      }
      
      const data = await response.json();
      const rows = data.values;
      
      if (!rows || rows.length === 0) {
        console.warn('Cut sheet is empty');
        return new Set();
      }
      
      console.log('Cut sheet headers:', rows[0]);
      
      const headers = rows[0];
      
      // Find the Lot Number column
      const lotNumberIndex = headers.findIndex(h => 
        h && h.toString().toLowerCase().trim() === 'lot number'
      );
      
      if (lotNumberIndex === -1) {
        console.error('No "Lot Number" column found in cut sheet. Headers:', headers);
        return new Set();
      }
      
      console.log(`Found Lot Number column at index ${lotNumberIndex}`);
      
      const cutLots = new Set();
      const rawLotNumbers = [];
      
      // Collect all lot numbers from cut sheet (starting from row 1, skip header)
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        let lotNumber = row[lotNumberIndex];
        
        if (!lotNumber) continue;
        
        lotNumber = lotNumber.toString().trim();
        rawLotNumbers.push(lotNumber);
        cutLots.add(lotNumber);
      }
      
      console.log(`Found ${cutLots.size} unique cut lot numbers`);
      console.log('Sample cut lot numbers (first 10):', rawLotNumbers.slice(0, 10));
      
      return cutLots;
      
    } catch (err) {
      console.error('Error fetching cut sheet:', err);
      return new Set();
    }
  }, [CUT_SHEET_ID, CUT_SHEET_NAME, API_KEY]);

  const fetchDataFromSheet = useCallback(async () => {
    try {
      setIsLoading(true);
      setError(null);
      
      // First fetch cut lot numbers
      const cutLots = await fetchCutLotNumbers();
      setCutLotNumbers(cutLots);
      
      console.log('Cut lot numbers set size:', cutLots.size);
      
      const response = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${SHEET_NAME}?key=${API_KEY}`
      );
      
      if (!response.ok) {
        throw new Error(`Failed to fetch data: ${response.status} ${response.statusText}`);
      }
      
      const data = await response.json();
      const rows = data.values;
      
      if (!rows || rows.length === 0) {
        throw new Error('No data found in sheet');
      }
      
      const headers = rows[0];
      console.log('Main sheet headers:', headers);
      
      // Find the Lot Number column in main sheet
      const lotNumberIndex = headers.findIndex(h => 
        h && h.toString().toLowerCase().trim() === 'lot number'
      );
      
      if (lotNumberIndex === -1) {
        console.error('No "Lot Number" column found in main sheet. Headers:', headers);
        throw new Error('Lot Number column not found in main sheet');
      }
      
      const dateIndex = headers.findIndex(h => h === 'Date');
      const fabricIndex = headers.findIndex(h => h === 'Fabric');
      const brandIndex = headers.findIndex(h => h === 'Brand');
      const shadeIndex = headers.findIndex(h => h === 'Shade');
      const sizeIndex = headers.findIndex(h => h === 'Size');
      const quantityIndex = headers.findIndex(h => h === 'Quantity');
      const unitIndex = headers.findIndex(h => h === 'Unit');
      const partyNameIndex = headers.findIndex(h => h === 'Party Name');
      const garmentTypeIndex = headers.findIndex(h => h === 'Garment Type');
      const sectionIndex = headers.findIndex(h => h === 'Section');
      const seasonIndex = headers.findIndex(h => h === 'Season');
      const remarksIndex = headers.findIndex(h => h === 'Remarks');
      const jobOrderNoIndex = headers.findIndex(h => h === 'Job Order No');
      
      const processedStyles = [];
      const uniqueGarmentTypes = new Set();
      const uniqueSections = new Set();
      const uniqueSeasons = new Set();
      const uniqueFabrics = new Set();
      
      let startProcessing = false;
      let skippedCount = 0;
      let cutCount = 0;
      let cutLotNumbersFound = [];
      
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const lotNumber = row[lotNumberIndex];
        const jobOrderNo = row[jobOrderNoIndex];
        
        if (!lotNumber) continue;
        
        const lotNumberStr = lotNumber.toString().trim();
        const jobOrderNoStr = jobOrderNo ? jobOrderNo.toString().trim() : '';
        
        // Start processing from JO-1109 based on Job Order No
        if (jobOrderNoStr === 'JO-1109') {
          startProcessing = true;
        }
        
        if (!startProcessing) continue;
        
        // Check if this lot number has been cut
        if (cutLots.has(lotNumberStr)) {
          cutCount++;
          cutLotNumbersFound.push(lotNumberStr);
          console.log(`Skipping cut lot: ${lotNumberStr} (JO: ${jobOrderNoStr})`);
          continue;
        }
        
        const garmentType = row[garmentTypeIndex] || 'OTHER';
        const section = row[sectionIndex] || 'UNISEX';
        const season = row[seasonIndex] || 'SUMMER';
        const fabric = row[fabricIndex] || 'N/A';
        const quantity = parseInt(row[quantityIndex]) || 0;
        
        if (quantity === 0) {
          skippedCount++;
          continue;
        }
        
        const shades = parseShades(row[shadeIndex]);
        const totalRollsRequired = shades.reduce((sum, s) => sum + s.rollsRequired, 0);
        
        if (totalRollsRequired === 0) {
          skippedCount++;
          continue;
        }
        
        processedStyles.push({
          id: i,
          lotNumber: lotNumberStr,
          jobOrderNo: jobOrderNoStr,
          date: row[dateIndex] || '',
          fabric: fabric,
          brand: row[brandIndex] || 'N/A',
          partyName: row[partyNameIndex] || 'N/A',
          garmentType: garmentType,
          section: section,
          season: season,
          size: row[sizeIndex] || '',
          quantity: quantity,
          unit: row[unitIndex] || 'SETS',
          shades: shades,
          totalRollsRequired: totalRollsRequired,
          remarks: row[remarksIndex] || '',
          isCut: false
        });
        
        if (garmentType && garmentType !== 'OTHER') uniqueGarmentTypes.add(garmentType);
        if (section) uniqueSections.add(section);
        if (season) uniqueSeasons.add(season);
        if (fabric && fabric !== 'N/A') uniqueFabrics.add(fabric);
      }
      
      console.log(`Cut lots found in main sheet: ${cutCount}`);
      if (cutLotNumbersFound.length > 0) {
        console.log('Cut lot numbers found:', cutLotNumbersFound.slice(0, 20));
      }
      
      if (processedStyles.length === 0) {
        throw new Error(`No data found from JO-1109 onwards. Cut lots filtered: ${cutCount}`);
      }
      
      setAllStyles(processedStyles);
      setFilteredStyles(processedStyles);
      setGarmentTypes(Array.from(uniqueGarmentTypes).sort());
      setSections(Array.from(uniqueSections).sort());
      setSeasons(Array.from(uniqueSeasons).sort());
      setFabrics(Array.from(uniqueFabrics).sort());
      
      console.log(`Loaded ${processedStyles.length} uncut lots. Skipped ${skippedCount} orders with 0 quantity. Filtered out ${cutCount} cut lots.`);
      
    } catch (err) {
      console.error('Error fetching data:', err);
      setError(err.message);
    } finally {
      setIsLoading(false);
    }
  }, [SHEET_ID, API_KEY, SHEET_NAME, fetchCutLotNumbers]);

  useEffect(() => {
    fetchDataFromSheet();
  }, [fetchDataFromSheet]);

  // Apply filters and auto-generate prediction when filters change
  useEffect(() => {
    let filtered = [...allStyles];
    
    if (selectedGarmentType) {
      filtered = filtered.filter(style => style.garmentType === selectedGarmentType);
    }
    
    if (selectedSection) {
      filtered = filtered.filter(style => style.section === selectedSection);
    }
    
    if (selectedSeason) {
      filtered = filtered.filter(style => style.season === selectedSeason);
    }
    
    if (selectedFabric) {
      filtered = filtered.filter(style => style.fabric === selectedFabric);
    }
    
    setFilteredStyles(filtered);
    
    // Auto-generate prediction when filters are applied
    if (filtered.length > 0 && selectedGarmentType) {
      generateAggregatedPrediction(filtered, selectedGarmentType, selectedSection, selectedSeason, selectedFabric);
    } else if (filtered.length > 0 && !selectedGarmentType && garmentTypes.length === 1) {
      setSelectedGarmentType(garmentTypes[0]);
    } else {
      setPredictionData(null);
    }
  }, [selectedGarmentType, selectedSection, selectedSeason, selectedFabric, allStyles, garmentTypes]);

  const generateAggregatedPrediction = (styles, garmentType, section, season, fabric) => {
    // Aggregate all shades across all lots
    const shadeAggregation = new Map();
    let totalRolls = 0;
    let lotNumbersList = [];
    let jobOrdersList = [];
    let uniqueBrands = new Set();
    let uniqueParties = new Set();
    
    styles.forEach(style => {
      lotNumbersList.push(style.lotNumber);
      if (style.jobOrderNo) jobOrdersList.push(style.jobOrderNo);
      uniqueBrands.add(style.brand);
      uniqueParties.add(style.partyName);
      
      style.shades.forEach(shade => {
        const existing = shadeAggregation.get(shade.color);
        if (existing) {
          existing.rollsRequired += shade.rollsRequired;
        } else {
          shadeAggregation.set(shade.color, {
            color: shade.color,
            rollsRequired: shade.rollsRequired,
            garmentType: style.garmentType,
            fabric: style.fabric
          });
        }
        totalRolls += shade.rollsRequired;
      });
    });
    
    const aggregatedShades = Array.from(shadeAggregation.values());
    // Sort shades by rolls required (descending)
    aggregatedShades.sort((a, b) => b.rollsRequired - a.rollsRequired);
    
    setPredictionData({
      garmentType: garmentType,
      fabric: fabric || 'All Fabrics',
      section: section || 'All Sections',
      season: season || 'All Seasons',
      totalLots: styles.length,
      lotNumbers: lotNumbersList,
      jobOrders: jobOrdersList,
      brands: Array.from(uniqueBrands).join(', '),
      parties: Array.from(uniqueParties).join(', '),
      totalRollsRequired: totalRolls,
      shadeWiseRequirements: aggregatedShades
    });
  };

  // Professional Black & White PDF Export Function
// Professional Black & White PDF Export Function - Compact Version
const exportToPDF = () => {
  if (!predictionData) return;
  
  // Create PDF document
  const doc = new jsPDF('p', 'mm', 'a4');
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  
  // Simple Header - Just text, no boxes
  doc.setTextColor(0, 0, 0);
  doc.setFontSize(16);
  doc.setFont('helvetica', 'bold');
  doc.text('FABRIC ROLL REQUIREMENT REPORT', pageWidth / 2, 20, { align: 'center' });
  
  doc.setFontSize(9);
  doc.setFont('helvetica', 'normal');
  doc.text(`Generated: ${new Date().toLocaleString()}`, pageWidth / 2, 28, { align: 'center' });
  
  // Key info in one line - compact
  doc.setFontSize(9);
  doc.text(`Garment: ${predictionData.garmentType} | Fabric: ${predictionData.fabric} | Section: ${predictionData.section} | Season: ${predictionData.season} | Total Rolls: ${predictionData.totalRollsRequired}`, pageWidth / 2, 36, { align: 'center' });
  
  let yPosition = 45;
  
  // Prepare table data
  const tableData = predictionData.shadeWiseRequirements.map((item, idx) => [
    idx + 1,
    item.garmentType || predictionData.garmentType,
    item.fabric || predictionData.fabric,
    item.color,
    `${item.rollsRequired}`,
    item.rollsRequired > 10 ? 'High' : item.rollsRequired > 5 ? 'Medium' : 'Normal'
  ]);
  
  // Add main table
  autoTable(doc, {
    startY: yPosition,
    head: [['S.No', 'Garment Type', 'Fabric', 'Shade / Color', 'Rolls', 'Priority']],
    body: tableData,
    foot: [['', '', '', 'TOTAL', `${predictionData.totalRollsRequired}`, '']],
    theme: 'plain',
    headStyles: {
      fillColor: [0, 0, 0],
      textColor: [255, 255, 255],
      fontSize: 9,
      fontStyle: 'bold',
      halign: 'center',
      lineWidth: 0.2,
      lineColor: [0, 0, 0]
    },
    footStyles: {
      fillColor: [240, 240, 240],
      textColor: [0, 0, 0],
      fontSize: 9,
      fontStyle: 'bold',
      halign: 'center',
      lineWidth: 0.2,
      lineColor: [0, 0, 0]
    },
    bodyStyles: {
      fontSize: 8,
      cellPadding: 2,
      textColor: [0, 0, 0],
      lineColor: [200, 200, 200],
      lineWidth: 0.1
    },
    alternateRowStyles: {
      fillColor: [245, 245, 245]
    },
    columnStyles: {
      0: { cellWidth: 12, halign: 'center' },
      1: { cellWidth: 35, halign: 'center' },
      2: { cellWidth: 35, halign: 'center' },
      3: { cellWidth: 50, halign: 'center' },
      4: { cellWidth: 18, halign: 'center' },
      5: { cellWidth: 22, halign: 'center' }
    },
    margin: { left: 15, right: 15 },
    didDrawPage: (data) => {
      // Get the final Y position after table
      const finalY = data.cursor ? data.cursor.y : yPosition + 50;
      
      // Add Lot Numbers section below table if space permits
      if (finalY < pageHeight - 30) {
        doc.setFontSize(8);
        doc.setFont('helvetica', 'bold');
        doc.text('Included Lot Numbers:', 15, finalY + 5);
        
        const lotNumbersStr = predictionData.lotNumbers.join(', ');
        doc.setFont('helvetica', 'normal');
        const splitLotNumbers = doc.splitTextToSize(lotNumbersStr, pageWidth - 30);
        doc.text(splitLotNumbers, 15, finalY + 12);
      } else {
        // If no space, add to next page
        doc.addPage();
        doc.setFontSize(8);
        doc.setFont('helvetica', 'bold');
        doc.text('Included Lot Numbers:', 15, 20);
        
        const lotNumbersStr = predictionData.lotNumbers.join(', ');
        doc.setFont('helvetica', 'normal');
        const splitLotNumbers = doc.splitTextToSize(lotNumbersStr, pageWidth - 30);
        doc.text(splitLotNumbers, 15, 27);
      }
      
      // Footer
      const pageCount = doc.internal.getNumberOfPages();
      doc.setDrawColor(0, 0, 0);
      doc.setLineWidth(0.2);
      doc.line(15, pageHeight - 8, pageWidth - 15, pageHeight - 8);
      
      doc.setFontSize(7);
      doc.setTextColor(0, 0, 0);
      doc.text(
        `Page ${data.pageNumber} of ${pageCount}`,
        pageWidth / 2,
        pageHeight - 3,
        { align: 'center' }
      );
    }
  });
  
  // Save the PDF
  const fileName = `Roll_Req_${predictionData.garmentType}_${new Date().toISOString().split('T')[0]}.pdf`;
  doc.save(fileName);
};

  const handleGarmentTypeSelect = (garmentType) => {
    setSelectedGarmentType(garmentType);
    setSelectedFabric('');
  };

  const clearFilters = () => {
    setSelectedGarmentType('');
    setSelectedSection('');
    setSelectedSeason('');
    setSelectedFabric('');
  };

  if (isLoading) {
    return (
      <div className="frp-loading-container">
        <div className="frp-loading-card">
          <div className="frp-spinner" />
          <p className="frp-loading-text">Fetching data from JO-1109 onwards...</p>
          <p className="frp-loading-subtext">Checking for cut lots...</p>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="frp-error-container">
        <div className="frp-error-card">
          <div className="frp-error-icon">⚠️</div>
          <h2 className="frp-error-title">Error Loading Data</h2>
          <p className="frp-error-message">{error}</p>
          <button
            onClick={() => window.location.reload()}
            className="frp-retry-button"
          >
            Retry
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="frp-container">
      <div className="frp-main-container">
        {/* Header */}
        <div className="frp-header">
          <div className="frp-header-content">
            <div className="frp-header-icon">
              📦🧮
            </div>
            <div className="frp-header-text">
              <h1 className="frp-header-title">
                Fabric Roll Requirement Report
              </h1>
              <p className="frp-header-subtitle">
                Select Garment Type and Fabric to view shade-wise roll requirements
              </p>
              {allStyles.length > 0 && cutLotNumbers.size > 0 && (
                <p className="frp-header-stats">
                  Loaded {allStyles.length} uncut lots | {garmentTypes.length} Garment Types | {fabrics.length} Fabrics
                  <span className="frp-cut-badge">✓ Excluding {cutLotNumbers.size} cut lots</span>
                </p>
              )}
              {allStyles.length > 0 && cutLotNumbers.size === 0 && (
                <p className="frp-header-stats">
                  Loaded {allStyles.length} lots | {garmentTypes.length} Garment Types | {fabrics.length} Fabrics
                  <span className="frp-cut-badge-warning">⚠ No cut sheet data loaded</span>
                </p>
              )}
            </div>
            <button
              onClick={() => window.history.back()}
              className="frp-back-button"
            >
              ← Back
            </button>
          </div>
        </div>

        {/* Garment Type Selection - Main Category Cards */}
        <div className="frp-card">
          <h2 className="frp-card-title">
            🏷️ Select Garment Type (Main Category)
          </h2>
          <div className="frp-garment-grid">
            {garmentTypes.map(type => {
              const orderCount = allStyles.filter(s => s.garmentType === type).length;
              const totalRolls = allStyles
                .filter(s => s.garmentType === type)
                .reduce((sum, s) => sum + s.totalRollsRequired, 0);
              
              return (
                <button
                  key={type}
                  onClick={() => handleGarmentTypeSelect(type)}
                  className={`frp-garment-card ${
                    selectedGarmentType === type 
                      ? 'frp-garment-card-selected'
                      : 'frp-garment-card-unselected'
                  }`}
                >
                  <div className="frp-garment-icon">
                    {type === 'LOWER' && '👖'}
                    {type === 'UPPER' && '👕'}
                    {type === 'SWEATSHIRT' && '👕'}
                    {type === 'T-SHIRT' && '👚'}
                    {type === 'HOODIE' && '🧥'}
                    {type === 'JACKET' && '🧥'}
                    {type === 'DRESS' && '👗'}
                    {!['LOWER','UPPER','SWEATSHIRT','T-SHIRT','HOODIE','JACKET','DRESS'].includes(type) && '👔'}
                  </div>
                  <div className="frp-garment-name">
                    {type}
                  </div>
                  <div className="frp-garment-stats">
                    {orderCount} Lots • {totalRolls} Rolls
                  </div>
                </button>
              );
            })}
          </div>
        </div>

        {/* Additional Filters (Section, Season, Fabric) */}
        {selectedGarmentType && (
          <div className="frp-card">
            <h2 className="frp-card-title">
              Refine Filters (Optional)
            </h2>
            <div className="frp-filters-grid">
              <div>
                <label className="frp-filter-label">
                  Section
                </label>
                <select
                  value={selectedSection}
                  onChange={(e) => setSelectedSection(e.target.value)}
                  className="frp-filter-select"
                >
                  <option value="">All Sections</option>
                  {sections.map(section => (
                    <option key={section} value={section}>{section}</option>
                  ))}
                </select>
              </div>
              
              <div>
                <label className="frp-filter-label">
                  Season
                </label>
                <select
                  value={selectedSeason}
                  onChange={(e) => setSelectedSeason(e.target.value)}
                  className="frp-filter-select"
                >
                  <option value="">All Seasons</option>
                  {seasons.map(season => (
                    <option key={season} value={season}>{season}</option>
                  ))}
                </select>
              </div>
              
              <div>
                <label className="frp-filter-label">
                  Fabric
                </label>
                <select
                  value={selectedFabric}
                  onChange={(e) => setSelectedFabric(e.target.value)}
                  className="frp-filter-select"
                >
                  <option value="">All Fabrics</option>
                  {fabrics.map(fabric => {
                    const fabricOrders = allStyles.filter(s => 
                      s.garmentType === selectedGarmentType && s.fabric === fabric
                    ).length;
                    return (
                      <option key={fabric} value={fabric}>
                        {fabric} ({fabricOrders} lots)
                      </option>
                    );
                  })}
                </select>
              </div>
              
              <div style={{ display: "flex", alignItems: "flex-end" }}>
                <button
                  onClick={clearFilters}
                  className="frp-clear-button"
                >
                  Clear All Filters
                </button>
              </div>
            </div>
            
            {/* Active Filters Display */}
            {(selectedSection || selectedSeason || selectedFabric) && (
              <div className="frp-active-filters">
                <strong>Active Filters:</strong>
                {selectedSection && <span className="frp-filter-tag">Section: {selectedSection}</span>}
                {selectedSeason && <span className="frp-filter-tag">Season: {selectedSeason}</span>}
                {selectedFabric && <span className="frp-filter-tag">Fabric: {selectedFabric}</span>}
              </div>
            )}
          </div>
        )}

        {/* Results Table - Garment Type, Fabric, Shade, Rolls Qty */}
        {predictionData && (
          <>
            {/* Summary Banner - Black and White version */}
            <div className="frp-summary-banner" style={{ 
              background: '#f5f5f5', 
              border: '2px solid #000',
              borderRadius: '8px'
            }}>
              <div className="frp-summary-content">
                <div>
                  <h2 className="frp-summary-title">
                    📊 {predictionData.garmentType} - Roll Requirements Summary
                    {predictionData.fabric !== 'All Fabrics' && ` (Fabric: ${predictionData.fabric})`}
                  </h2>
                  <p className="frp-summary-subtitle">
                    {predictionData.totalLots} Uncut Lots • {predictionData.brands}
                  </p>
                  {predictionData.section !== 'All Sections' && (
                    <p className="frp-summary-detail">
                      Section: {predictionData.section} | Season: {predictionData.season}
                    </p>
                  )}
                </div>
                <div className="frp-summary-total-box" style={{ 
                  border: '2px solid #000',
                  background: '#fff'
                }}>
                  <div className="frp-summary-total-label">Total Rolls Required</div>
                  <div className="frp-summary-total-value">
                    {predictionData.totalRollsRequired}
                  </div>
                </div>
              </div>
            </div>

            {/* Main Table - Garment Type, Fabric, Shade, Rolls Qty */}
            <div className="frp-card">
              <h2 className="frp-card-title">
                📋 Roll Requirements by Shade
              </h2>
              
              <div className="frp-table-container">
                <table className="frp-table">
                  <thead className="frp-table-header">
                    <tr style={{ background: '#000', color: '#fff' }}>
                      <th style={{ padding: "15px", textAlign: "left", color: '#fff' }}>S.No</th>
                      <th style={{ padding: "15px", textAlign: "left", color: '#fff' }}>Garment Type</th>
                      <th style={{ padding: "15px", textAlign: "left", color: '#fff' }}>Fabric</th>
                      <th style={{ padding: "15px", textAlign: "left", color: '#fff' }}>Shade / Color</th>
                      <th style={{ padding: "15px", textAlign: "center", color: '#fff' }}>Rolls Required</th>
                      <th style={{ padding: "15px", textAlign: "center", color: '#fff' }}>Priority</th>
                    </tr>
                  </thead>
                  <tbody>
                    {predictionData.shadeWiseRequirements.map((item, idx) => (
                      <tr key={idx} className={`frp-table-row ${idx % 2 === 0 ? 'frp-table-row-even' : 'frp-table-row-odd'}`}>
                        <td className="frp-table-cell" style={{ fontWeight: "500" }}>{idx + 1}</td>
                        <td className="frp-table-cell" style={{ fontWeight: "500" }}>
                          <span className="frp-badge-garment">
                            {item.garmentType || predictionData.garmentType}
                          </span>
                        </td>
                        <td className="frp-table-cell">
                          <span className="frp-badge-fabric">
                            {item.fabric || predictionData.fabric}
                          </span>
                        </td>
                        <td className="frp-table-cell" style={{ fontWeight: "500" }}>
                          {item.color}
                        </td>
                        <td className="frp-table-cell frp-table-cell-center">
                          <span className="frp-rolls-badge">
                            {item.rollsRequired} Rolls
                          </span>
                        </td>
                        <td className="frp-table-cell frp-table-cell-center">
                          <span>
                            {item.rollsRequired > 10 ? "High" : item.rollsRequired > 5 ? "Medium" : "Normal"}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot>
                    <tr className="frp-table-footer" style={{ background: '#f0f0f0', fontWeight: 'bold' }}>
                      <td className="frp-table-footer-cell" colSpan="4">GRAND TOTAL</td>
                      <td className="frp-table-footer-cell frp-table-cell-center">
                        <span className="frp-grand-total">
                          {predictionData.totalRollsRequired} Rolls
                        </span>
                      </td>
                      <td className="frp-table-footer-cell frp-table-cell-center"></td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            </div>

            {/* List of Lot Numbers included */}
            <div className="frp-job-orders">
              <h3 className="frp-job-orders-title">
                📋 Included Uncut Lot Numbers
              </h3>
              <div className="frp-job-orders-container">
                {predictionData.lotNumbers.map((lot, idx) => (
                  <span key={idx} className="frp-job-order-tag">
                    {lot}
                  </span>
                ))}
              </div>
            </div>

            {/* Action Buttons */}
            <div className="frp-action-buttons">
              <button
                onClick={() => window.print()}
                className="frp-print-button"
              >
                🖨️ Print Report
              </button>
              <button
                onClick={exportToPDF}
                className="frp-pdf-button"
                style={{
                  background: '#000',
                  color: 'white',
                  border: 'none',
                  padding: '12px 24px',
                  fontSize: '14px',
                  fontWeight: '600',
                  borderRadius: '8px',
                  cursor: 'pointer'
                }}
              >
                📄 Download PDF
              </button>
              <button
                onClick={() => {
                  const reportData = {
                    garmentType: predictionData.garmentType,
                    fabric: predictionData.fabric,
                    section: predictionData.section,
                    season: predictionData.season,
                    totalLots: predictionData.totalLots,
                    lotNumbers: predictionData.lotNumbers,
                    jobOrders: predictionData.jobOrders,
                    totalRollsRequired: predictionData.totalRollsRequired,
                    shadeWiseRequirements: predictionData.shadeWiseRequirements,
                    generatedAt: new Date().toISOString()
                  };
                  const blob = new Blob([JSON.stringify(reportData, null, 2)], { type: 'application/json' });
                  const url = URL.createObjectURL(blob);
                  const a = document.createElement('a');
                  a.href = url;
                  a.download = `roll_requirements_${predictionData.garmentType}${predictionData.fabric !== 'All Fabrics' ? '_' + predictionData.fabric : ''}.json`;
                  a.click();
                  URL.revokeObjectURL(url);
                }}
                className="frp-export-button"
              >
                📥 Export JSON
              </button>
            </div>
          </>
        )}

        {/* No selection message */}
        {!predictionData && !isLoading && garmentTypes.length > 0 && (
          <div className="frp-empty-state">
            <div className="frp-empty-icon">👆</div>
            <h3 className="frp-empty-title">Select a Garment Type Above</h3>
            <p>Choose a main category like LOWER, UPPER, or SWEATSHIRT to view roll requirements</p>
            <p className="frp-empty-subtitle">
              You can then filter by Fabric, Section, or Season for more specific analysis
            </p>
          </div>
        )}
      </div>
    </div>
  );
}