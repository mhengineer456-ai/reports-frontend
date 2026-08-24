// src/LotLogs.js
import React, { useState, useEffect, useMemo } from "react";
import { useHistory } from "react-router-dom";
import { getCurrentUser } from "./auth";
import { getLotLogs, addLotLog, subscribeToLotLogs, fetchLotLogsFromSheet } from "./lotLogService";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";

export default function LotLogs() {
  const history = useHistory();
  const currentUser = getCurrentUser();

  const [isAuthorized, setIsAuthorized] = useState(false);
  const [accessInputPass, setAccessInputPass] = useState("");
  const [accessError, setAccessError] = useState("");

  const [logs, setLogs] = useState([]);
  const [searchQuery, setSearchQuery] = useState("");
  const [categoryFilter, setCategoryFilter] = useState("ALL");
  const [submitting, setSubmitting] = useState(false);
  const [successBanner, setSuccessBanner] = useState("");
  const [errorBanner, setErrorBanner] = useState("");

  // Form State
  const [formState, setFormState] = useState({
    lotNumber: "11028",
    changeDetails: "Updated Job Order against lot number 11028 - revised order qty and stitch process",
    changedBy: currentUser?.name ? `${currentUser.name} (${currentUser.role || 'User'})` : "Gourav (Admin)",
    permissionBy: "Monu (Production Head)",
    category: "Job Order",
    priority: "High Alert",
    authPassword: ""
  });

  const handleUnlockAccess = (e) => {
    e.preventDefault();
    if (accessInputPass.trim() === "987456123") {
      setIsAuthorized(true);
      setAccessError("");
    } else {
      setAccessError("⛔ Incorrect Security Password! Access Denied.");
    }
  };

  // Load logs on mount and subscribe to live changes from Google Sheets
  useEffect(() => {
    // Initial fast load from local cache
    setLogs(getLotLogs());

    // Fetch fresh live logs from Google Sheets
    fetchLotLogsFromSheet().then((liveLogs) => {
      if (liveLogs && liveLogs.length > 0) {
        setLogs(liveLogs);
      }
    });

    const unsubscribe = subscribeToLotLogs((newLog) => {
      setLogs((prev) => [newLog, ...prev.filter((item) => item.id !== newLog.id)]);
    });

    return () => unsubscribe();
  }, []);

  const handleChange = (e) => {
    const { name, value } = e.target;
    setFormState((prev) => ({ ...prev, [name]: value }));
  };

  const handleSubmit = (e) => {
    e.preventDefault();
    setErrorBanner("");

    if (!formState.lotNumber || !formState.changeDetails) {
      alert("Please enter Lot Number and Change Details!");
      return;
    }

    // Security Authorization Password Check
    const VALID_PASSWORDS = ["987456123", "MH@2026", "1234", "admin", "MHENGINEER", "monu123"];
    if (!formState.authPassword || !VALID_PASSWORDS.includes(formState.authPassword.trim())) {
      setErrorBanner("⛔ Authorization Failed! Invalid Security Password. Only authorized personnel can submit lot change logs.");
      setTimeout(() => setErrorBanner(""), 6000);
      return;
    }

    setSubmitting(true);

    setTimeout(() => {
      const addedLog = addLotLog({
        lotNumber: formState.lotNumber,
        changeDetails: formState.changeDetails,
        changedBy: formState.changedBy,
        permissionBy: formState.permissionBy,
        category: formState.category,
        priority: formState.priority
      });

      setSubmitting(false);
      setSuccessBanner(`✅ Change logged for Lot #${addedLog.lotNumber}! Live real-time notification sent to all users.`);

      // Reset form slightly but keep author fields and clear password
      setFormState((prev) => ({
        ...prev,
        lotNumber: "",
        changeDetails: "",
        authPassword: ""
      }));

      setTimeout(() => setSuccessBanner(""), 6000);
    }, 400);
  };

  const filteredLogs = useMemo(() => {
    return logs.filter((log) => {
      if (categoryFilter !== "ALL" && log.category !== categoryFilter) {
        return false;
      }
      if (searchQuery.trim() !== "") {
        const q = searchQuery.toLowerCase();
        const lotMatch = String(log.lotNumber || "").toLowerCase().includes(q);
        const descMatch = String(log.changeDetails || "").toLowerCase().includes(q);
        const byMatch = String(log.changedBy || "").toLowerCase().includes(q);
        const permMatch = String(log.permissionBy || "").toLowerCase().includes(q);
        return lotMatch || descMatch || byMatch || permMatch;
      }
      return true;
    });
  }, [logs, searchQuery, categoryFilter]);

  const exportExcel = () => {
    const dataToExport = filteredLogs.map((l) => ({
      "Log ID": l.id,
      "Lot Number": l.lotNumber,
      "Change Details": l.changeDetails,
      "Changed By": l.changedBy,
      "Permission By": l.permissionBy,
      Category: l.category,
      Priority: l.priority,
      Timestamp: new Date(l.timestamp).toLocaleString()
    }));
    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Lot Changes");
    XLSX.writeFile(workbook, `Lot_Change_Logs_${Date.now()}.xlsx`);
  };

  const exportPDF = () => {
    const doc = new jsPDF({ orientation: "landscape" });
    doc.setFontSize(16);
    doc.text("Factory Suite Pro - Lot Change Audit Logs", 14, 15);
    doc.setFontSize(10);
    doc.text(`Exported on: ${new Date().toLocaleString()}`, 14, 22);

    const tableData = filteredLogs.map((l) => [
      l.id,
      l.lotNumber,
      l.changeDetails,
      l.changedBy,
      l.permissionBy,
      l.category,
      new Date(l.timestamp).toLocaleString()
    ]);

    autoTable(doc, {
      startY: 28,
      head: [["Log ID", "Lot #", "Change Description", "Changed By", "Permission By", "Category", "Date & Time"]],
      body: tableData,
      theme: "grid",
      headStyles: { fillColor: [30, 27, 75] }
    });

    doc.save(`Lot_Change_Logs_${Date.now()}.pdf`);
  };

  if (!isAuthorized) {
    return (
      <div
        style={{
          minHeight: "100vh",
          backgroundColor: "#0f172a",
          backgroundImage: "radial-gradient(at 50% 0%, rgba(99, 102, 241, 0.25) 0px, transparent 50%), radial-gradient(at 100% 100%, rgba(236, 72, 153, 0.2) 0px, transparent 50%)",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          padding: "20px",
          fontFamily: "'Plus Jakarta Sans', 'Inter', sans-serif"
        }}
      >
        <div
          style={{
            background: "rgba(30, 27, 75, 0.85)",
            backdropFilter: "blur(20px)",
            WebkitBackdropFilter: "blur(20px)",
            border: "1px solid rgba(99, 102, 241, 0.4)",
            borderRadius: "28px",
            padding: "40px 36px",
            width: "100%",
            maxWidth: "440px",
            color: "#ffffff",
            boxShadow: "0 25px 50px -12px rgba(0, 0, 0, 0.5), 0 0 30px rgba(99, 102, 241, 0.3)",
            textAlign: "center"
          }}
        >
          <div
            style={{
              width: "64px",
              height: "64px",
              borderRadius: "50%",
              background: "linear-gradient(135deg, #6366f1 0%, #4338ca 100%)",
              margin: "0 auto 20px auto",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontSize: "1.8rem",
              boxShadow: "0 10px 25px rgba(99, 102, 241, 0.4)"
            }}
          >
            🔒
          </div>

          <h2 style={{ margin: "0 0 8px 0", fontSize: "1.6rem", fontWeight: 800 }}>
            Restricted Audit Access
          </h2>
          <p style={{ margin: "0 0 24px 0", color: "#c7d2fe", fontSize: "0.9rem", lineHeight: 1.5 }}>
            Please enter the security access password to view & manage Lot Change Audit Logs.
          </p>

          <form onSubmit={handleUnlockAccess}>
            <div style={{ marginBottom: "20px", textAlign: "left" }}>
              <label style={{ display: "block", fontSize: "0.78rem", fontWeight: 800, color: "#a5b4fc", textTransform: "uppercase", marginBottom: "8px", letterSpacing: "0.05em" }}>
                Security Access Password *
              </label>
              <input
                type="password"
                placeholder="Enter Access Password"
                value={accessInputPass}
                onChange={(e) => setAccessInputPass(e.target.value)}
                autoFocus
                required
                style={{
                  width: "100%",
                  padding: "14px 16px",
                  borderRadius: "14px",
                  border: "1.5px solid rgba(165, 180, 252, 0.4)",
                  background: "rgba(15, 23, 42, 0.6)",
                  color: "#ffffff",
                  fontSize: "1rem",
                  fontWeight: 700,
                  boxSizing: "border-box",
                  outline: "none"
                }}
              />
            </div>

            {accessError && (
              <div style={{ background: "rgba(239, 68, 68, 0.2)", border: "1px solid #f87171", color: "#fca5a5", padding: "10px 14px", borderRadius: "10px", fontSize: "0.82rem", fontWeight: 700, marginBottom: "20px" }}>
                {accessError}
              </div>
            )}

            <div style={{ display: "flex", gap: "12px" }}>
              <button
                type="button"
                onClick={() => history.push("/dashboard")}
                style={{
                  flex: 1,
                  padding: "12px",
                  borderRadius: "12px",
                  border: "1px solid rgba(255, 255, 255, 0.2)",
                  background: "rgba(255, 255, 255, 0.1)",
                  color: "#ffffff",
                  fontSize: "0.9rem",
                  fontWeight: 700,
                  cursor: "pointer"
                }}
              >
                Cancel
              </button>
              <button
                type="submit"
                style={{
                  flex: 1.5,
                  padding: "12px",
                  borderRadius: "12px",
                  border: "none",
                  background: "linear-gradient(135deg, #6366f1 0%, #4338ca 100%)",
                  color: "#ffffff",
                  fontSize: "0.9rem",
                  fontWeight: 800,
                  cursor: "pointer",
                  boxShadow: "0 4px 15px rgba(99, 102, 241, 0.4)"
                }}
              >
                🔓 Unlock Access
              </button>
            </div>
          </form>
        </div>
      </div>
    );
  }

  return (
    <div className="lot-logs-container">
      <style>{`
        .lot-logs-container {
          min-height: 100vh;
          background-color: #f8fafc;
          background-image: 
            radial-gradient(at 0% 0%, rgba(99, 102, 241, 0.08) 0px, transparent 50%), 
            radial-gradient(at 100% 0%, rgba(236, 72, 153, 0.06) 0px, transparent 50%), 
            radial-gradient(at 50% 100%, rgba(16, 185, 129, 0.06) 0px, transparent 50%);
          font-family: 'Plus Jakarta Sans', 'Inter', system-ui, sans-serif;
          color: #0f172a;
          padding: 24px 32px;
        }

        .main-content {
          max-width: 100%;
          margin: 0 auto;
        }

        .header-card {
          background: linear-gradient(135deg, #1e1b4b 0%, #312e81 40%, #4338ca 100%);
          border-radius: 24px;
          padding: 32px 36px;
          margin-bottom: 28px;
          color: #ffffff;
          box-shadow: 0 20px 40px -15px rgba(30, 27, 75, 0.25);
          display: flex;
          justify-content: space-between;
          align-items: center;
          position: relative;
          overflow: hidden;
        }

        .header-title {
          margin: 0 0 6px 0;
          font-size: 2.1rem;
          font-weight: 800;
          letter-spacing: -0.02em;
        }

        .header-subtitle {
          margin: 0;
          font-size: 0.95rem;
          color: #c7d2fe;
          font-weight: 500;
        }

        .stats-badge-grid {
          display: flex;
          gap: 16px;
        }

        .stat-pill {
          background: rgba(255, 255, 255, 0.12);
          backdrop-filter: blur(12px);
          border: 1px solid rgba(255, 255, 255, 0.2);
          border-radius: 18px;
          padding: 12px 24px;
          text-align: center;
          min-width: 120px;
        }

        .stat-num {
          display: block;
          font-size: 1.7rem;
          font-weight: 800;
          color: #ffffff;
        }

        .stat-lbl {
          font-size: 0.72rem;
          color: #e0e7ff;
          text-transform: uppercase;
          font-weight: 700;
          letter-spacing: 0.06em;
        }

        .banner-success {
          background: #dcfce7;
          border: 1px solid #86efac;
          color: #166534;
          padding: 14px 20px;
          border-radius: 16px;
          margin-bottom: 24px;
          font-weight: 700;
          font-size: 0.95rem;
          display: flex;
          align-items: center;
          gap: 10px;
          box-shadow: 0 4px 12px rgba(22, 101, 52, 0.1);
        }

        /* Form Card */
        .card-panel {
          background: #ffffff;
          border-radius: 24px;
          border: 1px solid #e2e8f0;
          box-shadow: 0 10px 30px rgba(0, 0, 0, 0.03);
          padding: 28px 32px;
          margin-bottom: 28px;
        }

        .card-title {
          font-size: 1.15rem;
          font-weight: 800;
          color: #1e1b4b;
          margin: 0 0 20px 0;
          display: flex;
          align-items: center;
          gap: 10px;
        }

        .card-title::before {
          content: '';
          width: 5px;
          height: 20px;
          background: linear-gradient(135deg, #6366f1, #4338ca);
          border-radius: 4px;
        }

        .form-grid {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
          gap: 20px;
          margin-bottom: 20px;
        }

        .form-group {
          display: flex;
          flex-direction: column;
          gap: 6px;
        }

        .form-label {
          font-size: 0.82rem;
          color: #475569;
          font-weight: 700;
          text-transform: uppercase;
          letter-spacing: 0.04em;
        }

        .form-input, .form-select, .form-textarea {
          padding: 11px 16px;
          border: 1.5px solid #cbd5e1;
          border-radius: 12px;
          background: #ffffff;
          color: #0f172a;
          font-weight: 600;
          font-size: 0.9rem;
          transition: all 0.2s ease;
          width: 100%;
          box-sizing: border-box;
        }

        .form-input:focus, .form-select:focus, .form-textarea:focus {
          outline: none;
          border-color: #6366f1;
          box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.15);
        }

        .form-textarea {
          resize: vertical;
          min-height: 80px;
        }

        .submit-btn {
          background: linear-gradient(135deg, #4f46e5 0%, #312e81 100%);
          color: #ffffff;
          border: none;
          padding: 14px 28px;
          border-radius: 14px;
          font-size: 0.95rem;
          font-weight: 800;
          cursor: pointer;
          transition: all 0.2s ease;
          box-shadow: 0 8px 20px rgba(79, 70, 229, 0.3);
          display: inline-flex;
          align-items: center;
          gap: 10px;
        }

        .submit-btn:hover:not(:disabled) {
          transform: translateY(-2px);
          box-shadow: 0 12px 25px rgba(79, 70, 229, 0.4);
        }

        /* Filter Controls */
        .toolbar-grid {
          display: flex;
          justify-content: space-between;
          align-items: center;
          gap: 16px;
          flex-wrap: wrap;
          margin-bottom: 20px;
        }

        .search-box {
          flex: 1;
          min-width: 260px;
        }

        .action-btns {
          display: flex;
          gap: 12px;
        }

        .btn-exp {
          padding: 10px 18px;
          border: none;
          border-radius: 12px;
          font-size: 0.88rem;
          font-weight: 700;
          cursor: pointer;
          display: inline-flex;
          align-items: center;
          gap: 8px;
          transition: all 0.2s ease;
        }

        .btn-excel {
          background: linear-gradient(135deg, #10b981 0%, #059669 100%);
          color: #ffffff;
          box-shadow: 0 4px 12px rgba(16, 185, 129, 0.25);
        }

        .btn-pdf {
          background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%);
          color: #ffffff;
          box-shadow: 0 4px 12px rgba(239, 68, 68, 0.25);
        }

        /* Table */
        .table-container {
          background: #ffffff;
          border-radius: 20px;
          border: 1px solid #e2e8f0;
          overflow: hidden;
          box-shadow: 0 10px 30px rgba(0, 0, 0, 0.03);
        }

        .data-table {
          width: 100%;
          border-collapse: separate;
          border-spacing: 0;
          font-size: 0.88rem;
        }

        .table-header {
          background: linear-gradient(135deg, #1e1b4b 0%, #312e81 100%);
          color: #ffffff;
        }

        .th-cell {
          padding: 16px 14px;
          text-align: left;
          font-weight: 800;
          font-size: 0.75rem;
          text-transform: uppercase;
          letter-spacing: 0.05em;
          border-bottom: 2px solid #312e81;
        }

        .td-cell {
          padding: 14px;
          border-bottom: 1px solid #f1f5f9;
          color: #1e293b;
          vertical-align: middle;
        }

        .table-row:nth-child(even) .td-cell {
          background: #f8fafc;
        }

        .table-row:hover .td-cell {
          background: #e0e7ff;
        }

        .lot-badge {
          background: linear-gradient(135deg, #4338ca 0%, #312e81 100%);
          color: #ffffff;
          padding: 4px 12px;
          border-radius: 20px;
          font-weight: 800;
          font-size: 0.82rem;
          display: inline-block;
        }

        .priority-chip {
          padding: 3px 10px;
          border-radius: 12px;
          font-size: 0.75rem;
          font-weight: 700;
          display: inline-block;
        }

        .priority-high {
          background: #ffe4e6;
          color: #be123c;
          border: 1px solid #fecdd3;
        }

        .priority-standard {
          background: #e0e7ff;
          color: #3730a3;
          border: 1px solid #c7d2fe;
        }
      `}</style>

      <div className="main-content">
        {/* Header */}
        <div className="header-card">
          <div>
            <h1 className="header-title">📋 Lot Change Logs & Audit Center</h1>
            <p className="header-subtitle">
              Record lot modifications, specify permission authority, and broadcast real-time notifications to all users
            </p>
          </div>
          <div className="stats-badge-grid" style={{ alignItems: "center" }}>
            <button
              onClick={() => {
                if (history.length > 1) {
                  history.goBack();
                } else {
                  history.push("/dashboard");
                }
              }}
              style={{
                background: "rgba(255, 255, 255, 0.18)",
                backdropFilter: "blur(12px)",
                border: "1px solid rgba(255, 255, 255, 0.3)",
                color: "#ffffff",
                padding: "12px 20px",
                borderRadius: "16px",
                fontSize: "0.9rem",
                fontWeight: 800,
                cursor: "pointer",
                display: "inline-flex",
                alignItems: "center",
                gap: "8px",
                transition: "all 0.2s ease"
              }}
            >
              ⬅️ Back
            </button>
            <div className="stat-pill">
              <span className="stat-num">{logs.length}</span>
              <span className="stat-lbl">Total Logs</span>
            </div>
            <div className="stat-pill">
              <span className="stat-num">{filteredLogs.length}</span>
              <span className="stat-lbl">Filtered</span>
            </div>
          </div>
        </div>

        {errorBanner && (
          <div style={{ background: "#fef2f2", border: "1.5px solid #f87171", color: "#991b1b", padding: "14px 20px", borderRadius: "14px", fontWeight: "700", marginBottom: "20px", fontSize: "0.95rem" }}>
            {errorBanner}
          </div>
        )}

        {successBanner && <div className="banner-success">{successBanner}</div>}

        {/* New Log Submission Form */}
        <div className="card-panel">
          <h3 className="card-title">📝 Log New Change Against Lot Number</h3>
          <form onSubmit={handleSubmit}>
            <div className="form-grid">
              <div className="form-group">
                <label className="form-label">Lot Number *</label>
                <input
                  type="text"
                  name="lotNumber"
                  className="form-input"
                  placeholder="e.g. 11028"
                  value={formState.lotNumber}
                  onChange={handleChange}
                  required
                />
              </div>

              <div className="form-group">
                <label className="form-label">Changed By (User) *</label>
                <input
                  type="text"
                  name="changedBy"
                  className="form-input"
                  placeholder="e.g. Gourav (Admin)"
                  value={formState.changedBy}
                  onChange={handleChange}
                  required
                />
              </div>

              <div className="form-group">
                <label className="form-label">By Whom Permission (Authorized By) *</label>
                <input
                  type="text"
                  name="permissionBy"
                  className="form-input"
                  placeholder="e.g. Monu (Production Head)"
                  value={formState.permissionBy}
                  onChange={handleChange}
                  required
                />
              </div>

              <div className="form-group">
                <label className="form-label">Category / Department</label>
                <select
                  name="category"
                  className="form-select"
                  value={formState.category}
                  onChange={handleChange}
                >
                  <option value="Job Order">Job Order</option>
                  <option value="Stitching">Stitching</option>
                  <option value="Embroidery">Embroidery</option>
                  <option value="Printing">Printing</option>
                  <option value="Cutting">Cutting</option>
                  <option value="Packing">Packing</option>
                  <option value="Fabric Stock">Fabric Stock</option>
                  <option value="General">General</option>
                </select>
              </div>

              <div className="form-group">
                <label className="form-label">Priority Alert Level</label>
                <select
                  name="priority"
                  className="form-select"
                  value={formState.priority}
                  onChange={handleChange}
                >
                  <option value="High Alert">High Alert</option>
                  <option value="Standard Update">Standard Update</option>
                  <option value="Urgent Notice">Urgent Notice</option>
                </select>
              </div>
            </div>

            <div className="form-group" style={{ marginBottom: "20px" }}>
              <label className="form-label">What Changes Were Made? (Detailed Description) *</label>
              <textarea
                name="changeDetails"
                className="form-textarea"
                placeholder="Explain what was changed in code/job order against this lot number..."
                value={formState.changeDetails}
                onChange={handleChange}
                required
              />
            </div>

            <div className="form-group" style={{ marginBottom: "20px" }}>
              <label className="form-label" style={{ color: "#dc2626", fontWeight: "800" }}>
                🔒 Authorization Security Password *
              </label>
              <input
                type="password"
                name="authPassword"
                className="form-input"
                placeholder="Enter Security Password to confirm submission (e.g. MH@2026)"
                value={formState.authPassword}
                onChange={handleChange}
                required
                style={{ border: "1.5px solid #f87171", background: "#fff5f5" }}
              />
            </div>

            <button type="submit" className="submit-btn" disabled={submitting}>
              {submitting ? "Broadcasting..." : "🚀 Submit Change & Broadcast Notification"}
            </button>
          </form>
        </div>

        {/* Audit Log Table */}
        <div className="card-panel">
          <h3 className="card-title">🔍 Audit Trail & Change Logs History</h3>

          <div className="toolbar-grid">
            <div className="search-box">
              <input
                type="text"
                className="form-input"
                placeholder="Search by Lot #, User, Approver, or Details..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
              />
            </div>

            <div style={{ display: "flex", gap: "10px", alignItems: "center" }}>
              <select
                className="form-select"
                value={categoryFilter}
                onChange={(e) => setCategoryFilter(e.target.value)}
                style={{ width: "180px" }}
              >
                <option value="ALL">All Categories</option>
                <option value="Job Order">Job Order</option>
                <option value="Stitching">Stitching</option>
                <option value="Embroidery">Embroidery</option>
                <option value="Printing">Printing</option>
                <option value="Cutting">Cutting</option>
                <option value="Packing">Packing</option>
                <option value="Fabric Stock">Fabric Stock</option>
              </select>
            </div>

            <div className="action-btns">
              <button className="btn-exp btn-excel" onClick={exportExcel}>
                📊 Export Excel
              </button>
              <button className="btn-exp btn-pdf" onClick={exportPDF}>
                📄 Export PDF
              </button>
            </div>
          </div>

          <div className="table-container">
            <table className="data-table">
              <thead>
                <tr className="table-header">
                  <th className="th-cell">Log ID</th>
                  <th className="th-cell">Lot Number</th>
                  <th className="th-cell" style={{ width: "35%" }}>Change Description</th>
                  <th className="th-cell">Changed By</th>
                  <th className="th-cell">Permission Given By</th>
                  <th className="th-cell">Category</th>
                  <th className="th-cell">Date & Time</th>
                </tr>
              </thead>
              <tbody>
                {filteredLogs.length === 0 ? (
                  <tr>
                    <td colSpan="7" style={{ textAlign: "center", padding: "30px", color: "#64748b" }}>
                      No change logs found. Use the form above to log a new change!
                    </td>
                  </tr>
                ) : (
                  filteredLogs.map((log) => (
                    <tr key={log.id} className="table-row">
                      <td className="td-cell" style={{ fontWeight: "700", color: "#64748b" }}>
                        {log.id}
                      </td>
                      <td className="td-cell">
                        <span className="lot-badge">#{log.lotNumber}</span>
                      </td>
                      <td className="td-cell" style={{ fontWeight: "600" }}>
                        {log.changeDetails}
                      </td>
                      <td className="td-cell" style={{ fontWeight: "700", color: "#0f172a" }}>
                        {log.changedBy}
                      </td>
                      <td className="td-cell" style={{ fontWeight: "700", color: "#4338ca" }}>
                        👤 {log.permissionBy}
                      </td>
                      <td className="td-cell">
                        <span
                          className={`priority-chip ${
                            log.priority === "High Alert" ? "priority-high" : "priority-standard"
                          }`}
                        >
                          {log.category}
                        </span>
                      </td>
                      <td className="td-cell" style={{ fontSize: "0.8rem", color: "#64748b" }}>
                        {new Date(log.timestamp).toLocaleString()}
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}
