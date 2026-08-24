// src/PrintingMaterialReceiving.jsx
import React, { useState } from "react";
import { useHistory } from "react-router-dom";

export default function PrintingMaterialReceiving() {
  const history = useHistory();
  const [receipts, setReceipts] = useState([
    {
      id: "PMR-001",
      date: "2024-01-15",
      supplier: "Ink Masters Ltd.",
      materialType: "Ink",
      materialName: "Plastisol Ink - White",
      color: "White",
      quantity: 50,
      unit: "kg",
      batchNo: "INK-2401-001",
      receivedBy: "Rajesh Kumar",
      qualityStatus: "Approved",
      notes: "High quality ink received"
    },
    {
      id: "PMR-002",
      date: "2024-01-15",
      supplier: "Screen Supply Co.",
      materialType: "Screen",
      materialName: "Mesh Screen 120T",
      meshCount: "120",
      quantity: 30,
      unit: "pieces",
      batchNo: "SCR-2401-045",
      receivedBy: "Rajesh Kumar",
      qualityStatus: "Approved",
      notes: "Good quality screens"
    },
    {
      id: "PMR-003",
      date: "2024-01-14",
      supplier: "Chemical Solutions",
      materialType: "Chemical",
      materialName: "Emulsion Remover",
      quantity: 20,
      unit: "liters",
      batchNo: "CHM-2401-089",
      receivedBy: "Suresh Patil",
      qualityStatus: "Pending Check",
      notes: "Sample testing required"
    },
    {
      id: "PMR-004",
      date: "2024-01-14",
      supplier: "Color World",
      materialType: "Ink",
      materialName: "Water Based Ink - Red",
      color: "Red",
      quantity: 25,
      unit: "kg",
      batchNo: "INK-2401-089",
      receivedBy: "Suresh Patil",
      qualityStatus: "Approved",
      notes: "Color matching perfect"
    },
    {
      id: "PMR-005",
      date: "2024-01-13",
      supplier: "Screen Supply Co.",
      materialType: "Screen",
      materialName: "Mesh Screen 156T",
      meshCount: "156",
      quantity: 25,
      unit: "pieces",
      batchNo: "SCR-2401-067",
      receivedBy: "Amit Sharma",
      qualityStatus: "Approved",
      notes: "Premium mesh quality"
    }
  ]);

  const [searchTerm, setSearchTerm] = useState("");
  const [filterType, setFilterType] = useState("All");
  const [filterColor, setFilterColor] = useState("All");
  const [showAddModal, setShowAddModal] = useState(false);
  const [newReceipt, setNewReceipt] = useState({
    date: new Date().toISOString().split('T')[0],
    supplier: "",
    materialType: "Ink",
    materialName: "",
    color: "",
    meshCount: "",
    quantity: "",
    unit: "kg",
    batchNo: "",
    receivedBy: "",
    qualityStatus: "Pending Check",
    notes: ""
  });

  const materialTypes = ["All", "Ink", "Screen", "Chemical", "Consumables"];
  const colors = ["All", "White", "Red", "Blue", "Yellow", "Black", "Green", "Other"];
  const units = {
    Ink: ["kg", "liters", "gallons"],
    Screen: ["pieces", "boxes"],
    Chemical: ["liters", "kg", "bottles"],
    Consumables: ["rolls", "pieces", "boxes"]
  };

  const handleInputChange = (e) => {
    setNewReceipt({
      ...newReceipt,
      [e.target.name]: e.target.value
    });
  };

  const handleAddReceipt = () => {
    const newId = `PMR-${String(receipts.length + 1).padStart(3, '0')}`;
    const receiptToAdd = {
      id: newId,
      ...newReceipt,
      quantity: parseInt(newReceipt.quantity)
    };
    if (newReceipt.materialType !== "Screen") {
      delete receiptToAdd.meshCount;
    }
    setReceipts([receiptToAdd, ...receipts]);
    setShowAddModal(false);
    setNewReceipt({
      date: new Date().toISOString().split('T')[0],
      supplier: "",
      materialType: "Ink",
      materialName: "",
      color: "",
      meshCount: "",
      quantity: "",
      unit: "kg",
      batchNo: "",
      receivedBy: "",
      qualityStatus: "Pending Check",
      notes: ""
    });
  };

  const filteredReceipts = receipts.filter(receipt => {
    const matchesSearch = receipt.materialName.toLowerCase().includes(searchTerm.toLowerCase()) ||
                         receipt.supplier.toLowerCase().includes(searchTerm.toLowerCase()) ||
                         receipt.batchNo.toLowerCase().includes(searchTerm.toLowerCase());
    const matchesType = filterType === "All" || receipt.materialType === filterType;
    const matchesColor = filterColor === "All" || receipt.color === filterColor;
    return matchesSearch && matchesType && matchesColor;
  });

  const getStatusColor = (status) => {
    switch(status) {
      case 'Approved': return '#10B981';
      case 'Pending Check': return '#F59E0B';
      case 'Rejected': return '#EF4444';
      default: return '#6B7280';
    }
  };

  const getMaterialTypeIcon = (type) => {
    switch(type) {
      case 'Ink': return '🎨';
      case 'Screen': return '🖥️';
      case 'Chemical': return '⚗️';
      default: return '📦';
    }
  };

  // Statistics
  const stats = {
    totalReceipts: receipts.length,
    totalInk: receipts.filter(r => r.materialType === "Ink").reduce((sum, r) => sum + r.quantity, 0),
    totalScreens: receipts.filter(r => r.materialType === "Screen").reduce((sum, r) => sum + r.quantity, 0),
    totalChemicals: receipts.filter(r => r.materialType === "Chemical").reduce((sum, r) => sum + r.quantity, 0),
    pendingChecks: receipts.filter(r => r.qualityStatus === "Pending Check").length
  };

  return (
    <div style={{
      minHeight: "100vh",
      background: "linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%)",
      padding: "40px 20px"
    }}>
      <div style={{ maxWidth: "1400px", margin: "0 auto" }}>
        {/* Header */}
        <div style={{
          background: "white",
          borderRadius: "20px",
          padding: "30px",
          marginBottom: "30px",
          boxShadow: "0 10px 40px rgba(0,0,0,0.1)",
          background: "linear-gradient(135deg, #f093fb 0%, #f5576c 100%)",
          color: "white"
        }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "20px" }}>
            <div>
              <div style={{ fontSize: "3rem", marginBottom: "10px" }}>🖨️📥</div>
              <h1 style={{ margin: "0 0 10px 0", fontSize: "2rem" }}>Printing Material Receiving Report</h1>
              <p style={{ margin: 0, opacity: 0.9 }}>Track and manage all printing raw material receipts including inks, screens, and chemicals</p>
            </div>
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
                fontWeight: "500",
                backdropFilter: "blur(10px)"
              }}
            >
              ← Back to Dashboard
            </button>
          </div>
        </div>

        {/* Statistics Cards */}
        <div style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))",
          gap: "20px",
          marginBottom: "30px"
        }}>
          <div style={{
            background: "white",
            padding: "20px",
            borderRadius: "16px",
            boxShadow: "0 2px 10px rgba(0,0,0,0.05)",
            borderLeft: "4px solid #f093fb"
          }}>
            <div style={{ fontSize: "2rem", marginBottom: "5px" }}>📦</div>
            <div style={{ fontSize: "1.8rem", fontWeight: "bold", color: "#1e293b" }}>{stats.totalReceipts}</div>
            <div style={{ color: "#64748b" }}>Total Receipts</div>
          </div>
          <div style={{
            background: "white",
            padding: "20px",
            borderRadius: "16px",
            boxShadow: "0 2px 10px rgba(0,0,0,0.05)",
            borderLeft: "4px solid #10B981"
          }}>
            <div style={{ fontSize: "2rem", marginBottom: "5px" }}>🎨</div>
            <div style={{ fontSize: "1.8rem", fontWeight: "bold", color: "#1e293b" }}>{stats.totalInk}</div>
            <div style={{ color: "#64748b" }}>Ink (kg/liters)</div>
          </div>
          <div style={{
            background: "white",
            padding: "20px",
            borderRadius: "16px",
            boxShadow: "0 2px 10px rgba(0,0,0,0.05)",
            borderLeft: "4px solid #8B5CF6"
          }}>
            <div style={{ fontSize: "2rem", marginBottom: "5px" }}>🖥️</div>
            <div style={{ fontSize: "1.8rem", fontWeight: "bold", color: "#1e293b" }}>{stats.totalScreens}</div>
            <div style={{ color: "#64748b" }}>Screens Received</div>
          </div>
          <div style={{
            background: "white",
            padding: "20px",
            borderRadius: "16px",
            boxShadow: "0 2px 10px rgba(0,0,0,0.05)",
            borderLeft: "4px solid #F59E0B"
          }}>
            <div style={{ fontSize: "2rem", marginBottom: "5px" }}>⚠️</div>
            <div style={{ fontSize: "1.8rem", fontWeight: "bold", color: "#1e293b" }}>{stats.pendingChecks}</div>
            <div style={{ color: "#64748b" }}>Pending Quality Checks</div>
          </div>
        </div>

        {/* Filters and Search */}
        <div style={{
          background: "white",
          borderRadius: "16px",
          padding: "20px",
          marginBottom: "20px",
          display: "flex",
          gap: "15px",
          flexWrap: "wrap",
          alignItems: "center",
          justifyContent: "space-between"
        }}>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
            {materialTypes.map(type => (
              <button
                key={type}
                onClick={() => setFilterType(type)}
                style={{
                  padding: "8px 20px",
                  borderRadius: "10px",
                  border: "1px solid #e2e8f0",
                  background: filterType === type ? "#f093fb" : "white",
                  color: filterType === type ? "white" : "#475569",
                  cursor: "pointer",
                  fontWeight: "500",
                  transition: "all 0.3s ease"
                }}
              >
                {type}
              </button>
            ))}
          </div>
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
            <select
              value={filterColor}
              onChange={(e) => setFilterColor(e.target.value)}
              style={{
                padding: "10px 15px",
                borderRadius: "10px",
                border: "1px solid #e2e8f0",
                fontSize: "0.9rem"
              }}
            >
              {colors.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
            <input
              type="text"
              placeholder="Search by material, supplier or batch..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              style={{
                padding: "10px 15px",
                borderRadius: "10px",
                border: "1px solid #e2e8f0",
                width: "250px",
                fontSize: "0.9rem"
              }}
            />
            <button
              onClick={() => setShowAddModal(true)}
              style={{
                background: "linear-gradient(135deg, #f093fb 0%, #f5576c 100%)",
                color: "white",
                border: "none",
                padding: "10px 20px",
                borderRadius: "10px",
                cursor: "pointer",
                fontWeight: "500",
                display: "flex",
                alignItems: "center",
                gap: "8px"
              }}
            >
              + New Receipt
            </button>
          </div>
        </div>

        {/* Receipts Table */}
        <div style={{
          background: "white",
          borderRadius: "16px",
          overflow: "hidden",
          boxShadow: "0 2px 10px rgba(0,0,0,0.05)"
        }}>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead style={{ background: "#f8fafc", borderBottom: "2px solid #e2e8f0" }}>
                <tr>
                  <th style={{ padding: "15px", textAlign: "left" }}>ID</th>
                  <th style={{ padding: "15px", textAlign: "left" }}>Date</th>
                  <th style={{ padding: "15px", textAlign: "left" }}>Material</th>
                  <th style={{ padding: "15px", textAlign: "left" }}>Supplier</th>
                  <th style={{ padding: "15px", textAlign: "right" }}>Quantity</th>
                  <th style={{ padding: "15px", textAlign: "left" }}>Batch No</th>
                  <th style={{ padding: "15px", textAlign: "left" }}>Received By</th>
                  <th style={{ padding: "15px", textAlign: "left" }}>Status</th>
                </tr>
              </thead>
              <tbody>
                {filteredReceipts.map(receipt => (
                  <tr key={receipt.id} style={{ borderBottom: "1px solid #e2e8f0" }}>
                    <td style={{ padding: "15px", fontWeight: "500", color: "#f093fb" }}>{receipt.id}</td>
                    <td style={{ padding: "15px" }}>{receipt.date}</td>
                    <td style={{ padding: "15px" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                        <span style={{ fontSize: "1.2rem" }}>{getMaterialTypeIcon(receipt.materialType)}</span>
                        <div>
                          <div style={{ fontWeight: "500" }}>{receipt.materialName}</div>
                          <div style={{ fontSize: "0.8rem", color: "#64748b" }}>
                            {receipt.materialType}
                            {receipt.color && ` • ${receipt.color}`}
                            {receipt.meshCount && ` • ${receipt.meshCount}T`}
                          </div>
                        </div>
                      </div>
                    </td>
                    <td style={{ padding: "15px" }}>{receipt.supplier}</td>
                    <td style={{ padding: "15px", textAlign: "right" }}>{receipt.quantity} {receipt.unit}</td>
                    <td style={{ padding: "15px" }}>
                      <code style={{ background: "#f1f5f9", padding: "4px 8px", borderRadius: "6px", fontSize: "0.85rem" }}>
                        {receipt.batchNo}
                      </code>
                    </td>
                    <td style={{ padding: "15px" }}>{receipt.receivedBy}</td>
                    <td style={{ padding: "15px" }}>
                      <span style={{
                        display: "inline-block",
                        padding: "4px 12px",
                        borderRadius: "20px",
                        background: `${getStatusColor(receipt.qualityStatus)}20`,
                        color: getStatusColor(receipt.qualityStatus),
                        fontWeight: "500",
                        fontSize: "0.85rem"
                      }}>
                        {receipt.qualityStatus}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {/* Add Modal */}
      {showAddModal && (
        <div style={{
          position: "fixed",
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          background: "rgba(0,0,0,0.5)",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          zIndex: 1000
        }}>
          <div style={{
            background: "white",
            borderRadius: "20px",
            padding: "30px",
            width: "90%",
            maxWidth: "500px",
            maxHeight: "80vh",
            overflow: "auto"
          }}>
            <h2 style={{ margin: "0 0 20px 0" }}>New Material Receipt</h2>
            
            <div style={{ display: "flex", flexDirection: "column", gap: "15px" }}>
              <input type="date" name="date" value={newReceipt.date} onChange={handleInputChange} style={inputStyle} />
              <input type="text" name="supplier" placeholder="Supplier Name" value={newReceipt.supplier} onChange={handleInputChange} style={inputStyle} />
              <select name="materialType" value={newReceipt.materialType} onChange={handleInputChange} style={inputStyle}>
                <option value="Ink">Ink</option>
                <option value="Screen">Screen</option>
                <option value="Chemical">Chemical</option>
                <option value="Consumables">Consumables</option>
              </select>
              <input type="text" name="materialName" placeholder="Material Name" value={newReceipt.materialName} onChange={handleInputChange} style={inputStyle} />
              
              {newReceipt.materialType === "Ink" && (
                <input type="text" name="color" placeholder="Color" value={newReceipt.color} onChange={handleInputChange} style={inputStyle} />
              )}
              
              {newReceipt.materialType === "Screen" && (
                <input type="text" name="meshCount" placeholder="Mesh Count (e.g., 120, 156)" value={newReceipt.meshCount} onChange={handleInputChange} style={inputStyle} />
              )}
              
              <div style={{ display: "flex", gap: "10px" }}>
                <input type="number" name="quantity" placeholder="Quantity" value={newReceipt.quantity} onChange={handleInputChange} style={{...inputStyle, flex: 1}} />
                <select name="unit" value={newReceipt.unit} onChange={handleInputChange} style={{...inputStyle, width: "120px"}}>
                  {units[newReceipt.materialType]?.map(u => <option key={u} value={u}>{u}</option>)}
                </select>
              </div>
              
              <input type="text" name="batchNo" placeholder="Batch Number" value={newReceipt.batchNo} onChange={handleInputChange} style={inputStyle} />
              <input type="text" name="receivedBy" placeholder="Received By" value={newReceipt.receivedBy} onChange={handleInputChange} style={inputStyle} />
              <textarea name="notes" placeholder="Notes" value={newReceipt.notes} onChange={handleInputChange} style={{...inputStyle, minHeight: "80px"}} />
            </div>
            
            <div style={{ display: "flex", gap: "10px", marginTop: "20px" }}>
              <button onClick={handleAddReceipt} style={{...buttonStyle, background: "#10B981" }}>Save Receipt</button>
              <button onClick={() => setShowAddModal(false)} style={{...buttonStyle, background: "#EF4444" }}>Cancel</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

const inputStyle = {
  padding: "10px 15px",
  borderRadius: "10px",
  border: "1px solid #e2e8f0",
  fontSize: "0.9rem"
};

const buttonStyle = {
  flex: 1,
  padding: "12px",
  borderRadius: "10px",
  border: "none",
  color: "white",
  cursor: "pointer",
  fontWeight: "500"
};