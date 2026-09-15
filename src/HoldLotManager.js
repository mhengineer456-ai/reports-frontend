// src/HoldLotManager.js
import React, { useState, useEffect, useMemo, useRef } from "react";
import { useHistory } from "react-router-dom";
import { SPREADSHEET_IDS, fetchSheetDataFromBackend, BACKEND_URL, HOLD_LOTS_WEBHOOK_URL } from "./config";
import { getCurrentUser } from "./auth";

/** Department list with rich metadata */
const DEPARTMENTS = [
  { id: "cutting", name: "Cutting Department", icon: "✂️", color: "#10b981", desc: "Fabric spreading, marker planning & cut pieces" },
  { id: "embroidery", name: "Embroidery Department", icon: "🧵", color: "#8b5cf6", desc: "Challan issuance, multi-head embroidery & thread work" },
  { id: "printing", name: "Printing Department", icon: "🖨️", color: "#ec4899", desc: "Screen printing, heat transfer & curing operations" },
  { id: "stitching", name: "Stitching Department", icon: "🪡", color: "#3b82f6", desc: "Line assembly, floor operations & machine lines" },
  { id: "overlock", name: "Overlock Department", icon: "🧶", color: "#6366f1", desc: "Overlock seaming, edge finishing & thread trims" },
  { id: "kaj_button", name: "Kaj Button Department", icon: "🔘", color: "#f59e0b", desc: "Button attachment, kaj holes & eyelet punching" },
  { id: "washing", name: "Washing Department", icon: "🧺", color: "#06b6d4", desc: "Fabric enzyme washing, softening & dye treatment" },
  { id: "packing", name: "Packing & Finishing", icon: "📦", color: "#f97316", desc: "Folding, iron press, tag placement & master boxing" },
  { id: "quality", name: "Quality & Dispatch", icon: "🔍", color: "#e11d48", desc: "Final audit, defect quarantine & warehouse dispatch" }
];

/** Google Drive direct image URL helper */
const getDirectImageUrl = (url) => {
  if (!url) return "";
  const reg = /id=([a-zA-Z0-9_-]+)/;
  const match = String(url).match(reg);
  if (match && match[1]) {
    return `https://lh3.googleusercontent.com/d/${match[1]}`;
  }
  const dreg = /\/d\/([a-zA-Z0-9_-]+)/;
  const dmatch = String(url).match(dreg);
  if (dmatch && dmatch[1]) {
    return `https://lh3.googleusercontent.com/d/${dmatch[1]}`;
  }
  return url;
};

export default function HoldLotManager() {
  const history = useHistory();
  const currentUser = getCurrentUser();
  const defaultUserName = currentUser?.name || currentUser?.username || "Production Operator";

  // Flow State
  const [selectedDept, setSelectedDept] = useState(null);
  const [lotSearchInput, setLotSearchInput] = useState("");
  const [searchingLot, setSearchingLot] = useState(false);
  const [lotDetails, setLotDetails] = useState(null);
  const [lotNotFound, setLotNotFound] = useState(false);

  // All Job Orders Cache
  const [jobOrderCache, setJobOrderCache] = useState([]);
  const [loadingJobOrders, setLoadingJobOrders] = useState(false);

  // Hold Modal State
  const [isHoldModalOpen, setIsHoldModalOpen] = useState(false);
  const [holdBy, setHoldBy] = useState(defaultUserName);
  const [approvedBy, setApprovedBy] = useState("");
  const [location, setLocation] = useState("");
  const [holdReason, setHoldReason] = useState("Fabric Defect / Mismatch");
  const [customReason, setCustomReason] = useState("");
  const [priority, setPriority] = useState("High Alert");
  const [submittingHold, setSubmittingHold] = useState(false);
  const [successToast, setSuccessToast] = useState(null);

  // Live Hold Lots from Google Sheet (Spreadsheet ID: 1oBetbe44z2lUXngctvk3J31WBiWTv07NIgx5jFlylOs)
  const [holdRecords, setHoldRecords] = useState([]);
  const [loadingHolds, setLoadingHolds] = useState(false);

  // Fetch Live Hold Records from Google Sheet
  const fetchHoldRecordsFromSheet = async (deptName = "") => {
    setLoadingHolds(true);
    try {
      // Step 1: Try backend /api/sheets/hold-lots
      const queryParam = deptName ? `?department=${encodeURIComponent(deptName)}` : "";
      const res = await fetch(`${BACKEND_URL}/api/sheets/hold-lots${queryParam}`);
      if (res.ok) {
        const json = await res.json();
        if (json.success && Array.isArray(json.data)) {
          setHoldRecords(json.data);
          return;
        }
      }

      // Step 2: Fallback to direct sheet fetch
      const holdSpreadsheetId = SPREADSHEET_IDS.HOLD_LOTS || "1oBetbe44z2lUXngctvk3J31WBiWTv07NIgx5jFlylOs";
      const tab = deptName ? deptName.replace(/Department/gi, "").trim() : "All Holds";
      const sheetRes = await fetchSheetDataFromBackend(holdSpreadsheetId, `${tab}!A1:V1000`);

      if (sheetRes.ok && Array.isArray(sheetRes.values) && sheetRes.values.length > 1) {
        const rows = sheetRes.values;
        const headers = rows[0].map((h) => String(h || "").trim());
        const normalize = (s) => String(s || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
        const headerMap = {};
        headers.forEach((h, idx) => {
          headerMap[normalize(h)] = idx;
        });

        const getVal = (row, key) => {
          const idx = headerMap[normalize(key)];
          return idx !== undefined && row[idx] !== undefined ? String(row[idx]).trim() : "";
        };

        const list = [];
        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || (!row[0] && !row[1])) continue;
          list.push({
            timestamp: getVal(row, "Timestamp") || row[0],
            id: getVal(row, "Hold ID") || row[1],
            department: getVal(row, "Department") || row[2],
            lotNumber: getVal(row, "Lot Number") || row[3],
            jobOrderNo: getVal(row, "Job Order No") || row[4],
            date: getVal(row, "PO Date") || row[5],
            partyName: getVal(row, "Party Name") || row[6],
            brand: getVal(row, "Brand") || row[7],
            style: getVal(row, "Style") || row[8],
            fabric: getVal(row, "Fabric") || row[9],
            quantity: getVal(row, "Quantity") || row[10],
            unit: getVal(row, "Unit") || row[11],
            shade: getVal(row, "Shade") || row[12],
            size: getVal(row, "Size") || row[13],
            reason: getVal(row, "Hold Reason") || row[14],
            location: getVal(row, "Location") || getVal(row, "Hold Location") || getVal(row, "Placement") || row[15] || "",
            holdBy: getVal(row, "Hold By") || row[16] || row[15],
            approvedBy: getVal(row, "Approved By") || row[17] || row[16],
            priority: getVal(row, "Priority") || row[18] || row[17],
            status: getVal(row, "Status") || row[19] || row[18] || "ON HOLD",
            releasedAt: getVal(row, "Released At") || row[20] || row[19] || "",
            releasedBy: getVal(row, "Released By") || row[21] || row[20] || ""
          });
        }
        setHoldRecords(list.reverse());
      }
    } catch (err) {
      console.warn("Could not fetch hold records from Google Sheet:", err);
    } finally {
      setLoadingHolds(false);
    }
  };

  useEffect(() => {
    fetchHoldRecordsFromSheet(selectedDept?.name || "");
  }, [selectedDept]);

  // Load JobOrder rows once for lightning-fast search
  const loadJobOrders = async () => {
    setLoadingJobOrders(true);
    try {
      const res = await fetchSheetDataFromBackend(
        SPREADSHEET_IDS.JOBORDER,
        "JobOrder!A1:AZ50000"
      );

      if (res.ok && Array.isArray(res.values) && res.values.length > 0) {
        const rows = res.values;
        const headers = rows[0].map((h) => String(h || "").trim());
        const normalize = (s) => String(s || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");

        const headerMap = {};
        headers.forEach((h, idx) => {
          headerMap[normalize(h)] = idx;
        });

        const getVal = (row, key) => {
          const idx = headerMap[normalize(key)];
          return idx !== undefined && row[idx] !== undefined ? String(row[idx]).trim() : "";
        };

        const parsed = [];
        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length === 0) continue;

          const lotNo = getVal(row, "Lot Number") || getVal(row, "Lot No") || getVal(row, "Lot");
          if (!lotNo) continue;

          parsed.push({
            lotNumber: lotNo,
            jobOrderNo: getVal(row, "Job Order No") || getVal(row, "JobOrderNo") || getVal(row, "Order No"),
            date: getVal(row, "Date"),
            fabric: getVal(row, "Fabric"),
            brand: getVal(row, "Brand"),
            shade: getVal(row, "Shade"),
            size: getVal(row, "Size"),
            quantity: getVal(row, "Quantity"),
            unit: getVal(row, "Unit") || "PCS",
            partyName: getVal(row, "Party Name") || getVal(row, "Party"),
            garmentType: getVal(row, "Garment Type") || getVal(row, "Garment"),
            section: getVal(row, "Section"),
            season: getVal(row, "Season"),
            style: getVal(row, "Style"),
            pattern: getVal(row, "Pattern"),
            remarks: getVal(row, "Remarks"),
            imageUrl: getVal(row, "Image URL") || getVal(row, "Image"),
            status: getVal(row, "Status") || "Active"
          });
        }
        setJobOrderCache(parsed);
      }
    } catch (err) {
      console.error("Error preloading JobOrders for hold manager:", err);
    } finally {
      setLoadingJobOrders(false);
    }
  };

  useEffect(() => {
    loadJobOrders();
  }, []);

  // Search Lot Details from Cache
  const handleSearchLot = (targetLotNum) => {
    const cleanNum = String(targetLotNum || lotSearchInput).trim().toUpperCase().replace(/^LOT[-#\s]*/i, "");
    if (!cleanNum) return;

    setSearchingLot(true);
    setLotNotFound(false);

    const match = jobOrderCache.find(
      (j) => String(j.lotNumber).trim().toUpperCase().replace(/^LOT[-#\s]*/i, "") === cleanNum
    );

    if (match) {
      setLotDetails(match);
      setLotNotFound(false);
    } else {
      setLotDetails(null);
      setLotNotFound(true);
    }
    setSearchingLot(false);
  };

  // Open Hold Form Modal
  const openHoldModal = () => {
    if (!lotDetails) return;
    setIsHoldModalOpen(true);
    setApprovedBy("");
    setLocation("");
    setHoldReason("Fabric Defect / Mismatch");
    setCustomReason("");
    setPriority("High Alert");
  };

  // Submit Hold Data to Google Sheet (1oBetbe44z2lUXngctvk3J31WBiWTv07NIgx5jFlylOs)
  const handleSubmitHold = async (e) => {
    e.preventDefault();
    if (!lotDetails || !selectedDept) return;

    if (!location.trim()) {
      alert("Please enter the location / rack / bin where this hold lot is placed.");
      return;
    }
    if (!holdBy.trim()) {
      alert("Please enter 'Hold By' name.");
      return;
    }
    if (!approvedBy.trim()) {
      alert("Please enter 'Approved By' authorized person name.");
      return;
    }

    const finalReason = holdReason === "Other" ? (customReason.trim() || "Unspecified Hold Reason") : holdReason;
    const holdId = `HOLD-${Date.now().toString().slice(-6)}`;
    const nowIso = new Date().toISOString();
    const cleanLocation = location.trim();

    setSubmittingHold(true);

    const holdPayload = {
      action: "add_hold",
      id: holdId,
      department: selectedDept.name,
      lotNumber: lotDetails.lotNumber,
      jobOrderNo: lotDetails.jobOrderNo,
      date: lotDetails.date,
      partyName: lotDetails.partyName,
      brand: lotDetails.brand,
      style: lotDetails.style,
      fabric: lotDetails.fabric,
      quantity: lotDetails.quantity,
      unit: lotDetails.unit,
      shade: lotDetails.shade,
      size: lotDetails.size,
      reason: finalReason,
      location: cleanLocation,
      holdBy: holdBy.trim(),
      approvedBy: approvedBy.trim(),
      priority: priority,
      timestamp: nowIso
    };

    try {
      // 1. Post to Backend /api/sheets/hold-lot
      try {
        await fetch(`${BACKEND_URL}/api/sheets/hold-lot`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(holdPayload)
        });
      } catch (backendErr) {
        console.warn("Backend hold sync fallback to direct webhook:", backendErr);
      }

      // 2. Direct Apps Script Webhook to Hold Lots Sheet
      if (HOLD_LOTS_WEBHOOK_URL) {
        try {
          await fetch(HOLD_LOTS_WEBHOOK_URL, {
            method: "POST",
            headers: { "Content-Type": "text/plain;charset=utf-8" },
            body: JSON.stringify(holdPayload),
            mode: "no-cors"
          });
        } catch (webhookErr) {
          console.warn("Hold Lots Webhook fetch note:", webhookErr);
        }
      }

      // 3. Update UI state
      setHoldRecords((prev) => [
        {
          ...holdPayload,
          status: "ON HOLD"
        },
        ...prev
      ]);

      setIsHoldModalOpen(false);
      setSuccessToast(`Lot #${lotDetails.lotNumber} stored successfully in '${selectedDept.name}' tab at location "${cleanLocation}"!`);

      setTimeout(() => {
        setSuccessToast(null);
      }, 5000);
    } catch (err) {
      console.error("Error saving hold lot to Google Sheet:", err);
      alert("Failed to submit hold data: " + err.message);
    } finally {
      setSubmittingHold(false);
    }
  };

  // Release a lot from hold
  const handleReleaseHold = async (holdId, lotNo, deptName) => {
    if (!window.confirm(`Are you sure you want to release Lot #${lotNo} from hold?`)) {
      return;
    }

    try {
      // 1. Post release to backend
      try {
        await fetch(`${BACKEND_URL}/api/sheets/release-hold`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            holdId,
            lotNumber: lotNo,
            department: deptName,
            releasedBy: defaultUserName
          })
        });
      } catch (e) {}

      // 2. Direct Webhook backup
      if (HOLD_LOTS_WEBHOOK_URL) {
        try {
          await fetch(HOLD_LOTS_WEBHOOK_URL, {
            method: "POST",
            headers: { "Content-Type": "text/plain;charset=utf-8" },
            body: JSON.stringify({
              action: "release_hold",
              id: holdId,
              lotNumber: lotNo,
              department: deptName,
              releasedBy: defaultUserName
            }),
            mode: "no-cors"
          });
        } catch (e) {}
      }

      // 3. Update UI
      setHoldRecords((prev) =>
        prev.map((r) => (r.id === holdId ? { ...r, status: "RELEASED", releasedBy: defaultUserName, releasedAt: new Date().toISOString() } : r))
      );

      setSuccessToast(`Lot #${lotNo} released from hold successfully.`);
      setTimeout(() => setSuccessToast(null), 4000);
    } catch (err) {
      console.error("Error releasing lot from hold:", err);
    }
  };

  const activeHolds = useMemo(() => {
    return holdRecords.filter((r) => String(r.status || "").toUpperCase() === "ON HOLD");
  }, [holdRecords]);

  return (
    <>
      <style>{`
        .hlm-container {
          min-height: 100vh;
          background: #f8fafc;
          padding: 24px 32px 60px;
          font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
          color: #1e293b;
        }
        .hlm-header-card {
          background: #ffffff;
          border-radius: 16px;
          border: 1px solid #e2e8f0;
          padding: 20px 24px;
          display: flex;
          justify-content: space-between;
          align-items: center;
          flex-wrap: wrap;
          gap: 16px;
          box-shadow: 0 2px 6px rgba(0,0,0,0.03);
          margin-bottom: 24px;
        }
        .hlm-btn {
          display: inline-flex;
          align-items: center;
          gap: 6px;
          padding: 8px 16px;
          font-size: 12px;
          font-weight: 700;
          border-radius: 10px;
          cursor: pointer;
          transition: all 0.2s;
          border: none;
        }
        .hlm-btn-back {
          background: #f1f5f9;
          color: #334155;
          border: 1px solid #cbd5e1;
        }
        .hlm-btn-back:hover {
          background: #e2e8f0;
        }
        .hlm-dept-grid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
          gap: 16px;
          margin-bottom: 28px;
        }
        .hlm-dept-card {
          background: #ffffff;
          border: 2px solid #e2e8f0;
          border-radius: 14px;
          padding: 16px;
          cursor: pointer;
          transition: all 0.2s;
          display: flex;
          flex-direction: column;
          gap: 6px;
          position: relative;
        }
        .hlm-dept-card:hover {
          border-color: #f59e0b;
          transform: translateY(-2px);
          box-shadow: 0 6px 14px rgba(245, 158, 11, 0.12);
        }
        .hlm-dept-card.selected {
          border-color: #d97706;
          background: #fffbeb;
          box-shadow: 0 4px 12px rgba(217, 119, 6, 0.15);
        }
        .hlm-section-card {
          background: #ffffff;
          border-radius: 16px;
          border: 1px solid #e2e8f0;
          padding: 24px;
          box-shadow: 0 2px 6px rgba(0,0,0,0.03);
          margin-bottom: 28px;
        }
        .hlm-input {
          width: 100%;
          box-sizing: border-box;
          padding: 10px 14px;
          font-size: 13px;
          border-radius: 10px;
          border: 1.5px solid #cbd5e1;
          outline: none;
          background: #ffffff;
          transition: border-color 0.2s;
        }
        .hlm-input:focus {
          border-color: #d97706;
        }
        .hlm-lot-card {
          background: #ffffff;
          border: 1.5px solid #fed7aa;
          border-radius: 14px;
          padding: 20px;
          margin-top: 20px;
          background: linear-gradient(180deg, #fffaf5 0%, #ffffff 100%);
        }
        .hlm-lot-grid {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
          gap: 12px;
          margin-top: 14px;
        }
        .hlm-lot-item {
          background: #ffffff;
          border: 1px solid #fed7aa;
          border-radius: 8px;
          padding: 10px 12px;
        }
        .hlm-lot-item-title {
          font-size: 10px;
          font-weight: 800;
          text-transform: uppercase;
          color: #9a3412;
          margin-bottom: 2px;
          display: block;
        }
        .hlm-lot-item-val {
          font-size: 13px;
          font-weight: 700;
          color: #1e293b;
        }
        .hlm-modal-overlay {
          position: fixed;
          inset: 0;
          background: rgba(15, 23, 42, 0.7);
          backdrop-filter: blur(4px);
          display: flex;
          align-items: center;
          justify-content: center;
          z-index: 9999;
          padding: 20px;
        }
        .hlm-modal-card {
          background: #ffffff;
          border-radius: 18px;
          width: 100%;
          max-width: 580px;
          max-height: 90vh;
          overflow-y: auto;
          box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.25);
          border: 1px solid #cbd5e1;
        }
        .hlm-modal-header {
          background: linear-gradient(135deg, #d97706 0%, #b45309 100%);
          color: #ffffff;
          padding: 18px 24px;
          display: flex;
          justify-content: space-between;
          align-items: center;
        }
        .hlm-form-group {
          display: flex;
          flex-direction: column;
          gap: 6px;
          margin-bottom: 14px;
        }
        .hlm-form-label {
          font-size: 11px;
          font-weight: 800;
          text-transform: uppercase;
          color: #475569;
          letter-spacing: 0.5px;
        }
        .hlm-table {
          width: 100%;
          border-collapse: collapse;
          text-align: left;
          font-size: 12px;
        }
        .hlm-table th {
          background: #0f172a;
          color: #f8fafc;
          padding: 12px 14px;
          font-size: 11px;
          font-weight: 800;
          text-transform: uppercase;
          letter-spacing: 0.5px;
        }
        .hlm-table td {
          padding: 12px 14px;
          border-bottom: 1px solid #f1f5f9;
        }
        .hlm-table tr:hover td {
          background: #fffbeb;
        }
      `}</style>

      <div className="hlm-container">
        {/* Top Header Card */}
        <div className="hlm-header-card">
          <div style={{ display: "flex", alignItems: "center", gap: "14px" }}>
            <button
              onClick={() => history.push("/dashboard")}
              className="hlm-btn hlm-btn-back"
            >
              ← Dashboard
            </button>
            <div>
              <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                <span style={{ fontSize: "24px" }}>⏸️</span>
                <h1 style={{ margin: 0, fontSize: "22px", fontWeight: 900, color: "#0f172a" }}>
                  Hold Lot Action Center
                </h1>
                <span style={{ fontSize: "11px", fontWeight: 800, background: "#fef3c7", color: "#92400e", padding: "2px 10px", borderRadius: "12px", border: "1px solid #fde68a" }}>
                  Google Sheets Synchronized
                </span>
              </div>
              <span style={{ fontSize: "12px", color: "#64748b", fontWeight: 500 }}>
                Select Department → Search Lot # from JobOrder → Submit Hold to Department Sheet
              </span>
            </div>
          </div>

          <div style={{ display: "flex", gap: "10px", alignItems: "center", flexWrap: "wrap" }}>
            <button
              onClick={() => history.push("/hold-lots-report")}
              className="hlm-btn"
              style={{ background: "#7c2d12", color: "#ffffff", border: "none", fontWeight: 800 }}
            >
              🚨 View Hold Lots Report
            </button>
            <a
              href={`https://docs.google.com/spreadsheets/d/${SPREADSHEET_IDS.HOLD_LOTS || "1oBetbe44z2lUXngctvk3J31WBiWTv07NIgx5jFlylOs"}/edit`}
              target="_blank"
              rel="noopener noreferrer"
              className="hlm-btn"
              style={{ background: "#059669", color: "#ffffff", textDecoration: "none" }}
            >
              📊 Open Google Sheet
            </a>
            <span style={{ fontSize: "12px", fontWeight: 700, color: "#92400e", background: "#fef3c7", padding: "6px 14px", borderRadius: "10px", border: "1px solid #fde68a" }}>
              Active Holds: {activeHolds.length} Lots
            </span>
          </div>
        </div>

        {/* Success Toast Alert */}
        {successToast && (
          <div style={{ background: "#ecfdf5", border: "1.5px solid #a7f3d0", color: "#065f46", padding: "14px 20px", borderRadius: "12px", marginBottom: "20px", fontSize: "13px", fontWeight: 700, display: "flex", alignItems: "center", gap: "10px" }}>
            <span>✅</span>
            <span>{successToast}</span>
          </div>
        )}

        {/* STEP 1: Select Department */}
        <div className="hlm-section-card">
          <div style={{ display: "flex", alignItems: "center", gap: "10px", marginBottom: "16px" }}>
            <span style={{ background: "#d97706", color: "#fff", width: "24px", height: "24px", borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 900, fontSize: "12px" }}>1</span>
            <h2 style={{ margin: 0, fontSize: "16px", fontWeight: 800, color: "#0f172a" }}>
              Select Department to Store in Sheet
            </h2>
            {selectedDept && (
              <span style={{ fontSize: "11px", fontWeight: 800, color: "#059669", background: "#d1fae5", padding: "2px 8px", borderRadius: "6px" }}>
                Target Sheet Tab: "{selectedDept.name.replace(/Department/gi, "").trim()}"
              </span>
            )}
          </div>

          <div className="hlm-dept-grid">
            {DEPARTMENTS.map((dept) => {
              const isSelected = selectedDept?.id === dept.id;
              return (
                <div
                  key={dept.id}
                  onClick={() => {
                    setSelectedDept(dept);
                    setLotDetails(null);
                    setLotSearchInput("");
                    setLotNotFound(false);
                  }}
                  className={`hlm-dept-card ${isSelected ? "selected" : ""}`}
                >
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                    <span style={{ fontSize: "24px" }}>{dept.icon}</span>
                    {isSelected && (
                      <span style={{ fontSize: "11px", fontWeight: 900, color: "#d97706", background: "#fef3c7", padding: "2px 6px", borderRadius: "4px" }}>
                        ACTIVE
                      </span>
                    )}
                  </div>
                  <div style={{ fontSize: "14px", fontWeight: 800, color: isSelected ? "#92400e" : "#0f172a", marginTop: "4px" }}>
                    {dept.name}
                  </div>
                  <div style={{ fontSize: "11px", color: "#64748b", fontWeight: 500 }}>
                    {dept.desc}
                  </div>
                </div>
              );
            })}
          </div>
        </div>

        {/* STEP 2: Search Lot Number & View JobOrder Details */}
        {selectedDept && (
          <div className="hlm-section-card" style={{ borderLeft: `6px solid #d97706` }}>
            <div style={{ display: "flex", alignItems: "center", gap: "10px", marginBottom: "16px" }}>
              <span style={{ background: "#d97706", color: "#fff", width: "24px", height: "24px", borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 900, fontSize: "12px" }}>2</span>
              <h2 style={{ margin: 0, fontSize: "16px", fontWeight: 800, color: "#0f172a" }}>
                Enter Lot Number for {selectedDept.name}
              </h2>
            </div>

            <div style={{ display: "flex", gap: "12px", maxWidth: "600px" }}>
              <input
                type="text"
                value={lotSearchInput}
                onChange={(e) => setLotSearchInput(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && handleSearchLot()}
                placeholder="Enter Lot Number (e.g., 11046, 11102, 76063)..."
                className="hlm-input"
                autoFocus
              />
              <button
                type="button"
                onClick={() => handleSearchLot()}
                disabled={searchingLot || !lotSearchInput.trim()}
                style={{ background: "#0f172a", color: "#ffffff", border: "none", padding: "0 20px", borderRadius: "10px", fontWeight: 800, fontSize: "13px", cursor: "pointer", whiteSpace: "nowrap" }}
              >
                {searchingLot ? "Fetching..." : "Fetch Lot Details"}
              </button>
            </div>

            {loadingJobOrders && (
              <div style={{ fontSize: "11px", color: "#64748b", marginTop: "8px" }}>
                ⏳ Syncing JobOrder spreadsheet data...
              </div>
            )}

            {/* Lot Not Found Alert */}
            {lotNotFound && (
              <div style={{ marginTop: "16px", background: "#fef2f2", border: "1px solid #fecaca", color: "#991b1b", padding: "12px 16px", borderRadius: "10px", fontSize: "12px", fontWeight: 700 }}>
                ❌ Lot Number "{lotSearchInput}" not found in JobOrder sheet. Please check the lot number and try again.
              </div>
            )}

            {/* Lot Found Details Card */}
            {lotDetails && (
              <div className="hlm-lot-card">
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: "12px" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
                    {lotDetails.imageUrl ? (
                      <img
                        src={getDirectImageUrl(lotDetails.imageUrl)}
                        alt="Garment"
                        style={{ width: "60px", height: "60px", borderRadius: "10px", objectFit: "cover", border: "1.5px solid #fed7aa" }}
                        onError={(e) => { e.target.style.display = "none"; }}
                      />
                    ) : (
                      <div style={{ width: "60px", height: "60px", borderRadius: "10px", background: "#fff7ed", border: "1.5px solid #fed7aa", display: "flex", alignItems: "center", justifyContent: "center", fontSize: "24px" }}>
                        📦
                      </div>
                    )}
                    <div>
                      <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                        <span style={{ fontSize: "18px", fontWeight: 900, color: "#0f172a" }}>
                          Lot #{lotDetails.lotNumber}
                        </span>
                        <span style={{ fontSize: "11px", fontWeight: 800, background: "#d1fae5", color: "#065f46", padding: "2px 8px", borderRadius: "6px" }}>
                          Job Order #{lotDetails.jobOrderNo}
                        </span>
                      </div>
                      <div style={{ fontSize: "12px", color: "#64748b", fontWeight: 600, marginTop: "2px" }}>
                        {lotDetails.partyName} • {lotDetails.brand} • {lotDetails.style}
                      </div>
                    </div>
                  </div>

                  {/* Hold Action Button */}
                  <button
                    type="button"
                    onClick={openHoldModal}
                    style={{ background: "linear-gradient(135deg, #e11d48 0%, #be123c 100%)", color: "#ffffff", border: "none", padding: "10px 22px", borderRadius: "12px", fontWeight: 900, fontSize: "13px", cursor: "pointer", boxShadow: "0 4px 12px rgba(225, 29, 72, 0.3)", display: "flex", alignItems: "center", gap: "6px" }}
                  >
                    <span>⏸️</span>
                    <span>Put Lot on Hold</span>
                  </button>
                </div>

                {/* Grid Details */}
                <div className="hlm-lot-grid">
                  <div className="hlm-lot-item">
                    <span className="hlm-lot-item-title">Fabric</span>
                    <span className="hlm-lot-item-val">{lotDetails.fabric || "-"}</span>
                  </div>
                  <div className="hlm-lot-item">
                    <span className="hlm-lot-item-title">Total Quantity</span>
                    <span className="hlm-lot-item-val" style={{ color: "#d97706" }}>
                      {lotDetails.quantity} {lotDetails.unit}
                    </span>
                  </div>
                  <div className="hlm-lot-item">
                    <span className="hlm-lot-item-title">Garment / Section</span>
                    <span className="hlm-lot-item-val">{lotDetails.garmentType || "-"} ({lotDetails.section || "-"})</span>
                  </div>
                  <div className="hlm-lot-item">
                    <span className="hlm-lot-item-title">Season</span>
                    <span className="hlm-lot-item-val">{lotDetails.season || "-"}</span>
                  </div>
                  <div className="hlm-lot-item">
                    <span className="hlm-lot-item-title">PO Date</span>
                    <span className="hlm-lot-item-val">{lotDetails.date || "-"}</span>
                  </div>
                  <div className="hlm-lot-item">
                    <span className="hlm-lot-item-title">Current Status</span>
                    <span className="hlm-lot-item-val">{lotDetails.status || "Active"}</span>
                  </div>
                </div>

                {/* Shades & Sizes */}
                <div style={{ marginTop: "12px", padding: "10px 14px", background: "#ffffff", borderRadius: "8px", border: "1px solid #fed7aa", fontSize: "12px" }}>
                  <span style={{ fontWeight: 800, color: "#9a3412", marginRight: "6px" }}>Shades:</span>
                  <span style={{ fontWeight: 600, color: "#1e293b" }}>{lotDetails.shade || "-"}</span>
                  <span style={{ margin: "0 10px", color: "#cbd5e1" }}>|</span>
                  <span style={{ fontWeight: 800, color: "#9a3412", marginRight: "6px" }}>Sizes:</span>
                  <span style={{ fontWeight: 600, color: "#1e293b" }}>{lotDetails.size || "-"}</span>
                </div>
              </div>
            )}
          </div>
        )}

        {/* ACTIVE HOLDS TABLE (Direct from Google Sheet) */}
        <div className="hlm-section-card">
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "16px", flexWrap: "wrap", gap: "10px" }}>
            <div>
              <h2 style={{ margin: 0, fontSize: "16px", fontWeight: 800, color: "#0f172a" }}>
                📋 Hold Lots Live from Spreadsheet ({holdRecords.length} Total Records)
              </h2>
              <span style={{ fontSize: "11px", color: "#64748b" }}>
                Spreadsheet: HOLD LOT ACTION (ID: {SPREADSHEET_IDS.HOLD_LOTS || "1oBetbe44z2lUXngctvk3J31WBiWTv07NIgx5jFlylOs"})
              </span>
            </div>
            <button
              type="button"
              onClick={() => fetchHoldRecordsFromSheet(selectedDept?.name || "")}
              className="hlm-btn hlm-btn-back"
            >
              ↻ Refresh Sheet Records
            </button>
          </div>

          {loadingHolds ? (
            <div style={{ padding: "40px", textAlign: "center", color: "#64748b", fontWeight: 700 }}>
              ⏳ Fetching live records from Google Sheet...
            </div>
          ) : holdRecords.length === 0 ? (
            <div style={{ padding: "40px", textAlign: "center", color: "#94a3b8" }}>
              <div style={{ fontSize: "28px", marginBottom: "6px" }}>✨</div>
              <span style={{ fontSize: "13px", fontWeight: 600 }}>No hold lot entries in the spreadsheet yet. Submit your first hold above!</span>
            </div>
          ) : (
            <div style={{ overflowX: "auto" }}>
              <table className="hlm-table">
                <thead>
                  <tr>
                    <th>Hold ID</th>
                    <th>Status</th>
                    <th>Lot Number</th>
                    <th>Job Order No</th>
                    <th>Department</th>
                    <th>Party & Style</th>
                    <th>Quantity</th>
                    <th>Hold Reason</th>
                    <th>📍 Location</th>
                    <th>Hold By</th>
                    <th>Approved By</th>
                    <th>Hold Time</th>
                    <th style={{ textAlign: "center" }}>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {holdRecords.map((item) => {
                    const isReleased = String(item.status || "").toUpperCase() === "RELEASED";
                    return (
                      <tr key={item.id}>
                        <td style={{ fontWeight: 800, color: "#64748b" }}>{item.id}</td>
                        <td>
                          <span style={{ background: isReleased ? "#ecfdf5" : "#fee2e2", color: isReleased ? "#065f46" : "#991b1b", padding: "2px 8px", borderRadius: "12px", fontWeight: 900, fontSize: "11px", border: `1px solid ${isReleased ? "#a7f3d0" : "#fecaca"}` }}>
                            {isReleased ? "✓ RELEASED" : "⏸️ ON HOLD"}
                          </span>
                        </td>
                        <td>
                          <span style={{ background: "#fef3c7", color: "#92400e", padding: "2px 8px", borderRadius: "6px", fontWeight: 900, border: "1px solid #fde68a" }}>
                            {item.lotNumber}
                          </span>
                        </td>
                        <td style={{ fontWeight: 700 }}>{item.jobOrderNo}</td>
                        <td>
                          <span style={{ fontWeight: 700, color: "#0f172a" }}>
                            {item.department}
                          </span>
                        </td>
                        <td>
                          <div style={{ fontWeight: 700 }}>{item.partyName}</div>
                          <div style={{ fontSize: "11px", color: "#64748b" }}>{item.style}</div>
                        </td>
                        <td style={{ fontWeight: 800, color: "#d97706" }}>
                          {item.quantity} {item.unit}
                        </td>
                        <td>
                          <span style={{ background: "#fee2e2", color: "#991b1b", padding: "3px 8px", borderRadius: "6px", fontWeight: 700, fontSize: "11px", border: "1px solid #fecaca" }}>
                            ⚠️ {item.reason}
                          </span>
                        </td>
                        <td>
                          <span style={{ background: "#f1f5f9", color: "#0f172a", padding: "3px 8px", borderRadius: "6px", fontWeight: 700, fontSize: "11px", border: "1px solid #cbd5e1", display: "inline-flex", alignItems: "center", gap: "4px", whiteSpace: "nowrap" }}>
                            📍 {item.location || "Floor / Unassigned"}
                          </span>
                        </td>
                        <td style={{ fontWeight: 600 }}>{item.holdBy}</td>
                        <td style={{ fontWeight: 800, color: "#065f46" }}>🛡️ {item.approvedBy}</td>
                        <td style={{ fontSize: "11px", color: "#64748b", whiteSpace: "nowrap" }}>
                          {item.timestamp}
                        </td>
                        <td style={{ textAlign: "center" }}>
                          {!isReleased ? (
                            <button
                              type="button"
                              onClick={() => handleReleaseHold(item.id, item.lotNumber, item.department)}
                              style={{ background: "#ecfdf5", border: "1px solid #a7f3d0", color: "#065f46", padding: "4px 10px", borderRadius: "6px", cursor: "pointer", fontWeight: 800, fontSize: "11px" }}
                              title="Release this lot from hold"
                            >
                              ✓ Release
                            </button>
                          ) : (
                            <span style={{ fontSize: "11px", color: "#059669", fontWeight: 700 }}>
                              By {item.releasedBy || "Admin"}
                            </span>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>

        {/* HOLD AUTHORIZATION MODAL */}
        {isHoldModalOpen && lotDetails && selectedDept && (
          <div className="hlm-modal-overlay" onClick={() => setIsHoldModalOpen(false)}>
            <div className="hlm-modal-card" onClick={(e) => e.stopPropagation()}>
              <div className="hlm-modal-header">
                <div>
                  <h3 style={{ margin: 0, fontSize: "16px", fontWeight: 800 }}>
                    ⏸️ Place Lot #{lotDetails.lotNumber} on Hold
                  </h3>
                  <span style={{ fontSize: "11px", opacity: 0.9 }}>
                    Writing to Sheet Tab: "{selectedDept.name.replace(/Department/gi, "").trim()}"
                  </span>
                </div>
                <button
                  type="button"
                  onClick={() => setIsHoldModalOpen(false)}
                  style={{ background: "none", border: "none", color: "#ffffff", fontSize: "18px", cursor: "pointer" }}
                >
                  ✕
                </button>
              </div>

              <form onSubmit={handleSubmitHold} style={{ padding: "24px" }}>
                {/* Hold Reason Dropdown */}
                <div className="hlm-form-group">
                  <label className="hlm-form-label">Hold Reason / Issue Category *</label>
                  <select
                    value={holdReason}
                    onChange={(e) => setHoldReason(e.target.value)}
                    className="hlm-input"
                    required
                  >
                    <option value="Fabric Defect / Mismatch">Fabric Defect / Mismatch</option>
                    <option value="Printing Quality / Screen Issue">Printing Quality / Screen Issue</option>
                    <option value="Embroidery Error / Thread Problem">Embroidery Error / Thread Problem</option>
                    <option value="Stitching Size / Measurement Issue">Stitching Size / Measurement Issue</option>
                    <option value="Shortage of Accessories / Zip / Dori">Shortage of Accessories / Zip / Dori</option>
                    <option value="Pattern / Measurement Discrepancy">Pattern / Measurement Discrepancy</option>
                    <option value="Customer / Buyer Hold Request">Customer / Buyer Hold Request</option>
                    <option value="Machine Breakdown / Line Delay">Machine Breakdown / Line Delay</option>
                    <option value="Other">Other Custom Reason</option>
                  </select>
                </div>

                {holdReason === "Other" && (
                  <div className="hlm-form-group">
                    <label className="hlm-form-label">Specify Custom Reason *</label>
                    <input
                      type="text"
                      value={customReason}
                      onChange={(e) => setCustomReason(e.target.value)}
                      placeholder="Type detailed reason for hold..."
                      className="hlm-input"
                      required
                    />
                  </div>
                )}

                {/* Location / Placement Area */}
                <div className="hlm-form-group">
                  <label className="hlm-form-label">📍 Hold Lot Placement Location (Rack / Bin / Table / Area) *</label>
                  <input
                    type="text"
                    value={location}
                    onChange={(e) => setLocation(e.target.value)}
                    placeholder="e.g., Rack A1, Bin 4, Table 2, Floor Quarantine, QC Area..."
                    className="hlm-input"
                    required
                  />
                </div>

                {/* Hold By Input Label */}
                <div className="hlm-form-group">
                  <label className="hlm-form-label">Hold By (Operator / Supervisor Name) *</label>
                  <input
                    type="text"
                    value={holdBy}
                    onChange={(e) => setHoldBy(e.target.value)}
                    placeholder="Enter your name / operator title..."
                    className="hlm-input"
                    required
                  />
                </div>

                {/* Approved By Input Label */}
                <div className="hlm-form-group">
                  <label className="hlm-form-label">Approved By (Authorized Production Head / Manager) *</label>
                  <input
                    type="text"
                    value={approvedBy}
                    onChange={(e) => setApprovedBy(e.target.value)}
                    placeholder="Enter authorizing authority name (e.g., Monu Sir, Mohit Ji, Production Head)..."
                    className="hlm-input"
                    required
                  />
                </div>

                {/* Priority Selection */}
                <div className="hlm-form-group">
                  <label className="hlm-form-label">Hold Priority Alert</label>
                  <select
                    value={priority}
                    onChange={(e) => setPriority(e.target.value)}
                    className="hlm-input"
                  >
                    <option value="High Alert">🚨 High Alert</option>
                    <option value="Critical Hold">🛑 Critical Hold</option>
                    <option value="Standard Update">⚠️ Standard Hold</option>
                  </select>
                </div>

                {/* Action Buttons */}
                <div style={{ display: "flex", justifyContent: "flex-end", gap: "10px", marginTop: "24px" }}>
                  <button
                    type="button"
                    onClick={() => setIsHoldModalOpen(false)}
                    style={{ padding: "10px 18px", borderRadius: "10px", border: "1px solid #cbd5e1", background: "#ffffff", fontWeight: 700, fontSize: "12px", cursor: "pointer" }}
                  >
                    Cancel
                  </button>
                  <button
                    type="submit"
                    disabled={submittingHold}
                    style={{ padding: "10px 22px", borderRadius: "10px", border: "none", background: "linear-gradient(135deg, #d97706 0%, #b45309 100%)", color: "#ffffff", fontWeight: 900, fontSize: "12px", cursor: "pointer" }}
                  >
                    {submittingHold ? "Storing to Sheet..." : "Confirm & Store in Sheet"}
                  </button>
                </div>
              </form>
            </div>
          </div>
        )}
      </div>
    </>
  );
}
