// src/EmbroideryMaterialReceiving.jsx
import React, { useState, useEffect } from "react";
import { useHistory } from "react-router-dom";

export default function EmbroideryMaterialReceiving() {
  const history = useHistory();
  const [jobOrders, setJobOrders] = useState([]);
  const [filteredData, setFilteredData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [searchTerm, setSearchTerm] = useState("");
  const [showDebug, setShowDebug] = useState(false);
  const [rawSheetData, setRawSheetData] = useState(null);
  const [visibleColumns, setVisibleColumns] = useState({});
  const [challanFilter, setChallanFilter] = useState("CH-EMB"); // Filter for Challan No starting with

  // ========== GOOGLE SHEETS CONFIGURATION ==========
  const SHEET_ID = "1fKSwGBIpzWEFk566WRQ4bzQ0anJlmasoY8TwrTLQHXI";
  const API_KEY = "AIzaSyAomDFBkOySlIxKWSKGHe6ATv9gvaBr7uk";
  const SHEET_NAME = "JobOrder";
  
  const SHEET_URL = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${SHEET_NAME}?key=${API_KEY}`;
  // ================================================

  // All column headers in your desired order
  const ALL_COLUMNS = [
    "Sr. No",
    "Lot Number",
    "Date",
    "Fabric",
    "Brand",
    "Garment Type",
    "Style",
    "Party Name",
    "Emb",
    "Emb Details",
    "Challan No",
    "Challan Date",
    "Challan History JSON",
    "Challan Total Qty"
  ];

  useEffect(() => {
    fetchJobOrdersFromSheet();
  }, []);

  // Apply filters whenever jobOrders or filter criteria change
  useEffect(() => {
    applyFilters();
  }, [jobOrders, searchTerm, challanFilter]);

  const fetchJobOrdersFromSheet = async () => {
    setLoading(true);
    setError(null);
    
    try {
      const response = await fetch(SHEET_URL);
      
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      
      const data = await response.json();
      console.log("Raw API Response:", data);
      setRawSheetData(data);
      
      if (!data.values || data.values.length === 0) {
        throw new Error("No data found in sheet");
      }
      
      const headers = data.values[0];
      const rows = data.values.slice(1);
      
      console.log("Headers found:", headers);
      console.log("Total rows:", rows.length);
      
      // Find column indices for all columns
      const columnIndices = {};
      ALL_COLUMNS.forEach(columnName => {
        columnIndices[columnName] = headers.findIndex(h => 
          h && h.toString().trim().toLowerCase() === columnName.toLowerCase()
        );
      });
      
      console.log("Column indices:", columnIndices);
      
      // Transform data - include ALL columns
      const transformedData = rows.map((row, index) => {
        const jobOrder = {
          id: index + 1,
        };
        
        // Add all columns to the job order object
        ALL_COLUMNS.forEach(columnName => {
          const index = columnIndices[columnName];
          jobOrder[columnName] = index !== -1 ? (row[index] || "") : "";
        });
        
        return jobOrder;
      });
      
      console.log("Transformed data count:", transformedData.length);
      console.log("First row sample:", transformedData[0]);
      
      setJobOrders(transformedData);
      
      // Initialize visible columns (all true by default)
      const initialVisible = {};
      ALL_COLUMNS.forEach(col => {
        initialVisible[col] = true;
      });
      setVisibleColumns(initialVisible);
      
    } catch (err) {
      console.error("Error fetching from Google Sheets:", err);
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const applyFilters = () => {
    let filtered = [...jobOrders];
    
    // Filter 1: Challan No starts with specific prefix (e.g., "CH-EMB")
    if (challanFilter) {
      filtered = filtered.filter(order => {
        const challanNo = order["Challan No"] || "";
        return challanNo.startsWith(challanFilter);
      });
    }
    
    // Filter 2: Search term
    if (searchTerm) {
      const searchLower = searchTerm.toLowerCase();
      filtered = filtered.filter(order => {
        return (
          (order["Lot Number"] && order["Lot Number"].toLowerCase().includes(searchLower)) ||
          (order["Party Name"] && order["Party Name"].toLowerCase().includes(searchLower)) ||
          (order["Fabric"] && order["Fabric"].toLowerCase().includes(searchLower)) ||
          (order["Brand"] && order["Brand"].toLowerCase().includes(searchLower)) ||
          (order["Challan No"] && order["Challan No"].toLowerCase().includes(searchLower))
        );
      });
    }
    
    setFilteredData(filtered);
  };

  const refreshData = () => {
    fetchJobOrdersFromSheet();
  };

  const toggleColumn = (columnName) => {
    setVisibleColumns(prev => ({
      ...prev,
      [columnName]: !prev[columnName]
    }));
  };

  const selectAllColumns = () => {
    const allVisible = {};
    ALL_COLUMNS.forEach(col => {
      allVisible[col] = true;
    });
    setVisibleColumns(allVisible);
  };

  const deselectAllColumns = () => {
    const allVisible = {};
    ALL_COLUMNS.forEach(col => {
      allVisible[col] = false;
    });
    setVisibleColumns(allVisible);
  };

  // Statistics
  const stats = {
    totalOrders: jobOrders.length,
    filteredOrders: filteredData.length,
    totalQuantity: filteredData.reduce((sum, order) => sum + (parseInt(order["Challan Total Qty"]) || 0), 0),
    uniqueParties: new Set(filteredData.map(order => order["Party Name"]).filter(Boolean)).size,
    uniqueFabrics: new Set(filteredData.map(order => order["Fabric"]).filter(Boolean)).size,
    uniqueChallans: new Set(filteredData.map(order => order["Challan No"]).filter(Boolean)).size
  };

  if (loading) {
    return (
      <div style={{
        minHeight: "100vh",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        background: "linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%)"
      }}>
        <div style={{ textAlign: "center" }}>
          <div style={{ fontSize: "3rem", marginBottom: "20px" }}>🔄</div>
          <h2>Loading Job Orders...</h2>
          <p>Fetching data from Google Sheets</p>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div style={{
        minHeight: "100vh",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        background: "linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%)"
      }}>
        <div style={{ textAlign: "center", maxWidth: "600px", padding: "20px" }}>
          <div style={{ fontSize: "3rem", marginBottom: "20px" }}>⚠️</div>
          <h2 style={{ color: "#EF4444" }}>Error Loading Data</h2>
          <p>{error}</p>
          <button
            onClick={refreshData}
            style={{
              background: "linear-gradient(135deg, #667eea 0%, #764ba2 100%)",
              color: "white",
              border: "none",
              padding: "10px 20px",
              borderRadius: "10px",
              cursor: "pointer",
              marginTop: "20px"
            }}
          >
            Try Again
          </button>
        </div>
      </div>
    );
  }

  const visibleColumnsList = ALL_COLUMNS.filter(col => visibleColumns[col]);

  return (
    <div style={{
      minHeight: "100vh",
      background: "linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%)",
      padding: "40px 20px"
    }}>
      <div style={{ maxWidth: "100%", margin: "0 auto" }}>
        {/* Header */}
        <div style={{
          borderRadius: "20px",
          padding: "30px",
          marginBottom: "30px",
          background: "linear-gradient(135deg, #667eea 0%, #764ba2 100%)",
          color: "white"
        }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "20px" }}>
            <div>
              <div style={{ fontSize: "3rem", marginBottom: "10px" }}>🧵📥</div>
              <h1 style={{ margin: "0 0 10px 0", fontSize: "2rem" }}>Embroidery Material Receiving Report</h1>
              <p style={{ margin: 0, opacity: 0.9 }}>
                Showing records with Challan No starting with "{challanFilter}"
              </p>
            </div>
            <div style={{ display: "flex", gap: "10px" }}>
              <button
                onClick={() => setShowDebug(!showDebug)}
                style={{
                  background: "rgba(255,255,255,0.2)",
                  border: "1px solid rgba(255,255,255,0.3)",
                  padding: "12px 24px",
                  borderRadius: "12px",
                  color: "white",
                  cursor: "pointer",
                  fontSize: "1rem",
                  fontWeight: "500"
                }}
              >
                {showDebug ? "Hide Debug" : "Show Debug"}
              </button>
              <button
                onClick={refreshData}
                style={{
                  background: "rgba(255,255,255,0.2)",
                  border: "1px solid rgba(255,255,255,0.3)",
                  padding: "12px 24px",
                  borderRadius: "12px",
                  color: "white",
                  cursor: "pointer",
                  fontSize: "1rem",
                  fontWeight: "500"
                }}
              >
                🔄 Refresh
              </button>
              <button
                onClick={() => history.goBack()}
                style={{
                  background: "rgba(255,255,255,0.2)",
                  border: "1px solid rgba(255,255,255,0.3)",
                  padding: "12px 24px",
                  borderRadius: "12px",
                  color: "white",
                  cursor: "pointer",
                  fontSize: "1rem",
                  fontWeight: "500"
                }}
              >
                ← Back
              </button>
            </div>
          </div>
        </div>

        {/* Filter Section */}
        <div style={{
          background: "white",
          borderRadius: "16px",
          padding: "20px",
          marginBottom: "20px",
          boxShadow: "0 2px 10px rgba(0,0,0,0.05)"
        }}>
          <div style={{ display: "flex", gap: "20px", flexWrap: "wrap", alignItems: "flex-end" }}>
            <div style={{ flex: 1, minWidth: "200px" }}>
              <label style={{ display: "block", marginBottom: "8px", fontWeight: "500", color: "#475569" }}>
                Challan No Filter
              </label>
              <select
                value={challanFilter}
                onChange={(e) => setChallanFilter(e.target.value)}
                style={{
                  width: "100%",
                  padding: "10px 15px",
                  borderRadius: "10px",
                  border: "1px solid #e2e8f0",
                  fontSize: "0.95rem",
                  background: "white"
                }}
              >
                <option value="CH-EMB">CH-EMB (Embroidery Challans)</option>
                <option value="CH-PRN">CH-PRN (Printing Challans)</option>
                <option value="CH-ST">CH-ST (Stitching Challans)</option>
                <option value="CH-PK">CH-PK (Packing Challans)</option>
                <option value="">All Challans (No Filter)</option>
              </select>
            </div>
            <div style={{ flex: 2, minWidth: "300px" }}>
              <label style={{ display: "block", marginBottom: "8px", fontWeight: "500", color: "#475569" }}>
                Search
              </label>
              <input
                type="text"
                placeholder="🔍 Search by Lot Number, Party, Fabric, Brand..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                style={{
                  width: "100%",
                  padding: "10px 15px",
                  borderRadius: "10px",
                  border: "1px solid #e2e8f0",
                  fontSize: "0.95rem",
                  outline: "none"
                }}
              />
            </div>
            <div>
              <button
                onClick={() => {
                  setSearchTerm("");
                  setChallanFilter("CH-EMB");
                }}
                style={{
                  padding: "10px 20px",
                  background: "#EF4444",
                  color: "white",
                  border: "none",
                  borderRadius: "10px",
                  cursor: "pointer",
                  fontSize: "0.95rem"
                }}
              >
                Clear Filters
              </button>
            </div>
          </div>
        </div>

        {/* Column Visibility Controls */}
        <div style={{
          background: "white",
          borderRadius: "16px",
          padding: "20px",
          marginBottom: "20px",
          boxShadow: "0 2px 10px rgba(0,0,0,0.05)"
        }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "15px", flexWrap: "wrap", gap: "10px" }}>
            <h3 style={{ margin: 0, color: "#1e293b" }}>📊 Column Visibility ({visibleColumnsList.length}/{ALL_COLUMNS.length})</h3>
            <div style={{ display: "flex", gap: "10px" }}>
              <button
                onClick={selectAllColumns}
                style={{
                  padding: "6px 12px",
                  background: "#10B981",
                  color: "white",
                  border: "none",
                  borderRadius: "8px",
                  cursor: "pointer",
                  fontSize: "0.85rem"
                }}
              >
                Select All
              </button>
              <button
                onClick={deselectAllColumns}
                style={{
                  padding: "6px 12px",
                  background: "#EF4444",
                  color: "white",
                  border: "none",
                  borderRadius: "8px",
                  cursor: "pointer",
                  fontSize: "0.85rem"
                }}
              >
                Deselect All
              </button>
            </div>
          </div>
          <div style={{
            display: "flex",
            flexWrap: "wrap",
            gap: "10px",
            maxHeight: "150px",
            overflowY: "auto",
            padding: "10px",
            background: "#f8fafc",
            borderRadius: "12px"
          }}>
            {ALL_COLUMNS.map(column => (
              <label key={column} style={{
                display: "flex",
                alignItems: "center",
                gap: "8px",
                padding: "6px 12px",
                background: visibleColumns[column] ? "#667eea20" : "white",
                borderRadius: "8px",
                cursor: "pointer",
                fontSize: "0.85rem",
                border: "1px solid #e2e8f0"
              }}>
                <input
                  type="checkbox"
                  checked={visibleColumns[column]}
                  onChange={() => toggleColumn(column)}
                  style={{ cursor: "pointer" }}
                />
                <span style={{ color: visibleColumns[column] ? "#667eea" : "#64748b" }}>
                  {column}
                </span>
              </label>
            ))}
          </div>
        </div>

        {/* Debug Panel */}
        {showDebug && rawSheetData && (
          <div style={{
            background: "#1e1e1e",
            color: "#d4d4d4",
            borderRadius: "16px",
            padding: "20px",
            marginBottom: "20px",
            fontFamily: "monospace",
            fontSize: "12px",
            overflowX: "auto"
          }}>
            <h3 style={{ color: "#fff", marginTop: 0 }}>🔍 Debug Information</h3>
            <div><strong>Total Rows in Sheet:</strong> {rawSheetData.values?.length || 0}</div>
            <div><strong>Rows After Filter:</strong> {filteredData.length}</div>
            <div><strong>Challan Filter:</strong> {challanFilter || "None"}</div>
            <details>
              <summary style={{ cursor: "pointer", marginTop: "10px" }}>View Sample Filtered Data</summary>
              <pre style={{ fontSize: "11px", marginTop: "10px" }}>{JSON.stringify(filteredData.slice(0, 2), null, 2)}</pre>
            </details>
          </div>
        )}

        {/* Statistics Cards */}
        <div style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))",
          gap: "15px",
          marginBottom: "20px"
        }}>
          <div style={{
            background: "white",
            padding: "15px",
            borderRadius: "12px",
            borderLeft: "4px solid #667eea"
          }}>
            <div style={{ fontSize: "0.85rem", color: "#64748b" }}>Total Records</div>
            <div style={{ fontSize: "1.8rem", fontWeight: "bold", color: "#1e293b" }}>{stats.filteredOrders}</div>
            <div style={{ fontSize: "0.75rem", color: "#94a3b8" }}>of {stats.totalOrders} total</div>
          </div>
          <div style={{
            background: "white",
            padding: "15px",
            borderRadius: "12px",
            borderLeft: "4px solid #10B981"
          }}>
            <div style={{ fontSize: "0.85rem", color: "#64748b" }}>Total Quantity</div>
            <div style={{ fontSize: "1.8rem", fontWeight: "bold", color: "#1e293b" }}>{stats.totalQuantity.toLocaleString()}</div>
          </div>
          <div style={{
            background: "white",
            padding: "15px",
            borderRadius: "12px",
            borderLeft: "4px solid #8B5CF6"
          }}>
            <div style={{ fontSize: "0.85rem", color: "#64748b" }}>Unique Parties</div>
            <div style={{ fontSize: "1.8rem", fontWeight: "bold", color: "#1e293b" }}>{stats.uniqueParties}</div>
          </div>
          <div style={{
            background: "white",
            padding: "15px",
            borderRadius: "12px",
            borderLeft: "4px solid #F59E0B"
          }}>
            <div style={{ fontSize: "0.85rem", color: "#64748b" }}>Unique Challans</div>
            <div style={{ fontSize: "1.8rem", fontWeight: "bold", color: "#1e293b" }}>{stats.uniqueChallans}</div>
          </div>
        </div>

        {/* Job Orders Table */}
        <div style={{
          background: "white",
          borderRadius: "16px",
          overflow: "hidden",
          boxShadow: "0 2px 10px rgba(0,0,0,0.05)"
        }}>
          <div style={{ overflowX: "auto", maxHeight: "60vh", overflowY: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", minWidth: "800px" }}>
              <thead style={{ 
                background: "#f8fafc", 
                borderBottom: "2px solid #e2e8f0",
                position: "sticky",
                top: 0,
                zIndex: 10
              }}>
                <tr>
                  <th style={{ padding: "12px 15px", textAlign: "left", fontWeight: "600", background: "#f8fafc" }}>#</th>
                  {visibleColumnsList.map(column => (
                    <th key={column} style={{ 
                      padding: "12px 15px", 
                      textAlign: "left", 
                      fontWeight: "600",
                      background: "#f8fafc",
                      minWidth: "150px"
                    }}>
                      {column}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filteredData.length === 0 ? (
                  <tr>
                    <td colSpan={visibleColumnsList.length + 1} style={{ padding: "60px", textAlign: "center", color: "#64748b" }}>
                      No records found with Challan No starting with "{challanFilter}"
                    </td>
                  </tr>
                ) : (
                  filteredData.map((order, index) => (
                    <tr 
                      key={order.id} 
                      style={{ 
                        borderBottom: "1px solid #e2e8f0",
                        backgroundColor: index % 2 === 0 ? "white" : "#fafafa",
                      }}
                    >
                      <td style={{ padding: "10px 15px", fontWeight: "500", color: "#667eea" }}>
                        {index + 1}
                      </td>
                      {visibleColumnsList.map(column => {
                        let value = order[column];
                        const isChallanNo = column === "Challan No";
                        const isSrNo = column === "Sr. No";
                        
                        // Format JSON fields for better display
                        if (column === "Challan History JSON" && value) {
                          try {
                            const parsed = JSON.parse(value);
                            value = JSON.stringify(parsed, null, 2);
                          } catch(e) {
                            // Keep as is if not valid JSON
                          }
                        }
                        // Truncate long text
                        const displayValue = value && value.length > 50 ? value.substring(0, 50) + "..." : value;
                        
                        return (
                          <td 
                            key={column} 
                            style={{ 
                              padding: "10px 15px",
                              fontSize: "0.85rem",
                              verticalAlign: "top",
                              ...(isChallanNo && { 
                                background: "#fef3c7",
                                fontWeight: "500",
                                color: "#92400e"
                              }),
                              ...(isSrNo && {
                                fontWeight: "500",
                                color: "#667eea"
                              })
                            }}
                            title={value && value.length > 50 ? value : ""}
                          >
                            {column === "Challan Total Qty" && value ? (
                              <span style={{ fontWeight: "500", color: "#8B5CF6" }}>
                                {parseInt(value).toLocaleString()}
                              </span>
                            ) : (
                              displayValue || "-"
                            )}
                          </td>
                        );
                      })}
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>

        {/* Filter Summary Footer */}
        {filteredData.length > 0 && (
          <div style={{
            marginTop: "20px",
            padding: "15px",
            background: "white",
            borderRadius: "12px",
            textAlign: "center",
            color: "#64748b",
            fontSize: "0.85rem",
            border: "1px solid #e2e8f0"
          }}>
            Showing {filteredData.length} record(s) where Challan No starts with "{challanFilter}"
            {searchTerm && ` and matches "${searchTerm}"`}
          </div>
        )}
      </div>
    </div>
  );
}