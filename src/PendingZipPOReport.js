import React, { useState, useEffect, useMemo, useRef } from "react";
import { useHistory } from "react-router-dom";
import Navbar from "./Navbar";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import { fetchSheetDataFromBackend } from "./config";

// Multi-Select Dropdown Component with Checkboxes & Search
const MultiSelectDropdown = ({ title, options, selected, onChange, icon }) => {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState("");
  const dropdownRef = useRef(null);

  useEffect(() => {
    const handleClickOutside = (e) => {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target)) {
        setOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  const toggleOption = (option) => {
    if (selected.includes(option)) {
      onChange(selected.filter((item) => item !== option));
    } else {
      onChange([...selected, option]);
    }
  };

  const selectAll = () => {
    if (selected.length === options.length) {
      onChange([]);
    } else {
      onChange([...options]);
    }
  };

  const filteredOptions = useMemo(() => {
    if (!search.trim()) return options;
    return options.filter((o) => String(o).toLowerCase().includes(search.toLowerCase().trim()));
  }, [options, search]);

  const buttonText = useMemo(() => {
    if (selected.length === 0) return `${icon} All ${title}s`;
    if (selected.length === 1) return `${icon} ${selected[0]}`;
    return `${icon} ${selected.length} ${title}s Selected`;
  }, [selected, title, icon]);

  return (
    <div ref={dropdownRef} style={{ position: "relative", width: "100%" }}>
      <button
        type="button"
        className="zpr-select"
        onClick={() => setOpen((prev) => !prev)}
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          cursor: "pointer",
          textAlign: "left",
          background: selected.length > 0 ? "#e0e7ff" : "#ffffff",
          borderColor: selected.length > 0 ? "#4338ca" : "#cbd5e1",
          color: selected.length > 0 ? "#312e81" : "#1e293b",
          fontWeight: selected.length > 0 ? 700 : 500
        }}
      >
        <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{buttonText}</span>
        <span style={{ fontSize: "0.75rem", marginLeft: "0.3rem" }}>{open ? "▲" : "▼"}</span>
      </button>

      {open && (
        <div
          style={{
            position: "absolute",
            top: "calc(100% + 4px)",
            left: 0,
            right: 0,
            background: "#ffffff",
            border: "1px solid #cbd5e1",
            borderRadius: "12px",
            boxShadow: "0 10px 25px -5px rgba(0, 0, 0, 0.15)",
            zIndex: 100,
            padding: "0.6rem",
            maxHeight: "260px",
            display: "flex",
            flexDirection: "column",
            gap: "0.5rem"
          }}
        >
          {options.length > 5 && (
            <input
              type="text"
              placeholder={`Search ${title}...`}
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              style={{
                width: "100%",
                padding: "0.4rem 0.6rem",
                borderRadius: "6px",
                border: "1px solid #cbd5e1",
                fontSize: "0.82rem",
                boxSizing: "border-box"
              }}
            />
          )}

          <div style={{ display: "flex", justifyContent: "space-between", fontSize: "0.78rem", padding: "0.2rem 0.4rem" }}>
            <span
              onClick={selectAll}
              style={{ color: "#4338ca", fontWeight: 700, cursor: "pointer", textDecoration: "underline" }}
            >
              {selected.length === options.length ? "Deselect All" : "Select All"}
            </span>
            {selected.length > 0 && (
              <span
                onClick={() => onChange([])}
                style={{ color: "#dc2626", fontWeight: 700, cursor: "pointer", textDecoration: "underline" }}
              >
                Clear ({selected.length})
              </span>
            )}
          </div>

          <div style={{ overflowY: "auto", flex: 1, display: "flex", flexDirection: "column", gap: "0.2rem" }}>
            {filteredOptions.length === 0 ? (
              <div style={{ fontSize: "0.8rem", color: "#94a3b8", padding: "0.5rem", textAlign: "center" }}>No matches found</div>
            ) : (
              filteredOptions.map((option) => {
                const isChecked = selected.includes(option);
                return (
                  <label
                    key={option}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: "0.5rem",
                      padding: "0.35rem 0.5rem",
                      borderRadius: "6px",
                      background: isChecked ? "#f1f5f9" : "transparent",
                      cursor: "pointer",
                      fontSize: "0.85rem",
                      color: "#1e293b",
                      userSelect: "none"
                    }}
                  >
                    <input
                      type="checkbox"
                      checked={isChecked}
                      onChange={() => toggleOption(option)}
                      style={{ cursor: "pointer", accentColor: "#4338ca" }}
                    />
                    <span>{option}</span>
                  </label>
                );
              })
            )}
          </div>
        </div>
      )}
    </div>
  );
};

// Spreadsheet IDs & API Configuration
const API_KEY = process.env.REACT_APP_GOOGLE_API_KEY || "AIzaSyAomDFBkOySlIxKWSKGHe6ATv9gvaBr7uk";
const JOB_ORDERS_SPREADSHEET_ID = "1fKSwGBIpzWEFk566WRQ4bzQ0anJlmasoY8TwrTLQHXI";
const JOB_ORDERS_RANGE = "JobOrder!A1:BZ2000";

const ZIP_PO_SPREADSHEET_ID = "16mifNw0WMIlnZ1XRHsuH_8kVUm_6Y1O3uVsoM-Hjppo";
const ZIP_PO_RANGE = "ZipPurchaseOrders!A:V";

const SHADE_PO_SPREADSHEET_ID = "1JgJF9Er7lYDW0rINQzUUafqonx0yxkVaAauPgX5QNfk";
const SHADE_PO_RANGE = "POs!A:M";

const INDEX_SPREADSHEET_ID = "1Hj3JeJEKB43aYYWv8gk2UhdU6BWuEQfCg5pBlTdBMNA";
const INDEX_RANGE = "Index!A1:AF3000";

const PendingZipPOReport = () => {
  const history = useHistory();

  // State Management
  const [jobOrders, setJobOrders] = useState([]);
  const [existingPOLots, setExistingPOLots] = useState(new Set());
  const [completedIndexLots, setCompletedIndexLots] = useState(new Set());
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [refreshing, setRefreshing] = useState(false);

  // Multi-Select Filters
  const [searchQuery, setSearchQuery] = useState("");
  const [selectedParties, setSelectedParties] = useState([]);
  const [selectedFabrics, setSelectedFabrics] = useState([]);
  const [selectedGarments, setSelectedGarments] = useState([]);
  const [selectedPriorities, setSelectedPriorities] = useState([]);
  const [selectedFYs, setSelectedFYs] = useState([]);

  // Pagination
  const [currentPage, setCurrentPage] = useState(1);
  const pageSize = 20;

  // Expander
  const [expandedRows, setExpandedRows] = useState(new Set());

  // Helper: Normalize string for comparison
  const normalize = (str) => String(str || "").trim().toLowerCase();

  // Helper: Check if Zip is required for a row
  const isZipRequired = (zipValue) => {
    const val = normalize(zipValue);
    if (!val) return false;
    const ignoreList = [
      "no", "n/a", "na", "none", "—", "-", "nil",
      "not decided", "notdecided", "to be decided", "tbd", "undecided"
    ];
    return !ignoreList.includes(val);
  };

  // Helper: Check if Completed Status in Index Sheet contains a date/timestamp
  const isCompletedWithDate = (statusValue) => {
    if (!statusValue) return false;
    const str = String(statusValue).trim();
    if (!str || str === "[]" || str === "—" || str === "-") return false;

    // Try parsing JSON array/object
    if (str.startsWith("[") || str.startsWith("{")) {
      try {
        const parsed = JSON.parse(str);
        const items = Array.isArray(parsed) ? parsed : [parsed];
        return items.some((item) => {
          if (!item) return false;
          const st = String(item.status || "").toLowerCase();
          const ts = String(item.timestamp || item.date || "").trim();
          return st.includes("complete") || ts.length > 0;
        });
      } catch {
        // Fallback to string check
      }
    }

    // String fallback regex check for date indicators
    const normStr = str.toLowerCase();
    const dateRegex = /\d{4}|\d{1,2}[\/\-]\d{1,2}|gmt|ist|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|complete/i;
    return dateRegex.test(normStr);
  };

  // Helper: Exponential Backoff Fetch to prevent HTTP 429 Too Many Requests
  const fetchWithRetry = async (url, retries = 4, delay = 1000) => {
    for (let i = 0; i < retries; i++) {
      try {
        const res = await fetch(url);
        if (res.status === 429) {
          console.warn(`Google Sheets API 429 Rate Limit. Retrying in ${delay}ms... (Attempt ${i + 1}/${retries})`);
          await new Promise((resolve) => setTimeout(resolve, delay));
          delay *= 2; // Exponential backoff: 1s, 2s, 4s, 8s
          continue;
        }
        return res;
      } catch (err) {
        if (i === retries - 1) throw err;
        await new Promise((resolve) => setTimeout(resolve, delay));
        delay *= 2;
      }
    }
    return fetch(url);
  };

  // Helper: Check if Supervisor or Party Name is excluded (e.g. Dushyant, Jain Hosiery)
  const isExcludedSupervisorOrParty = (row) => {
    const sup = normalize(row.supervisor || row.indexSupervisor || "");
    const sub = normalize(row.submittedBy || "");
    const party = normalize(row.partyName || row.indexPartyName || "");

    const excludedTerms = ["dushyant", "jainhosiery", "jain hosiery", "jain_hosiery"];

    const matchSup = excludedTerms.some((term) => sup.includes(term) || sub.includes(term));
    const matchParty = excludedTerms.some((term) => party.includes(term));

    return matchSup || matchParty;
  };

  // Helper: Calculate Indian Financial Year
  const getFinancialYear = (dateInput) => {
    if (!dateInput) return "";
    const d = new Date(dateInput);
    if (isNaN(d.getTime())) return "";
    const yr = d.getFullYear();
    const m = d.getMonth();
    return m >= 3 ? `${yr}-${yr + 1}` : `${yr - 1}-${yr}`;
  };

  // Fetch Data from Google Sheets
  const fetchData = async ({ isRefresh = false } = {}) => {
    try {
      if (isRefresh) setRefreshing(true);
      else setLoading(true);
      setError("");

      // 1. Fetch Zip Purchase Orders & Accumulate PO Quantities per Lot
      const zipUrl = `https://sheets.googleapis.com/v4/spreadsheets/${ZIP_PO_SPREADSHEET_ID}/values/${encodeURIComponent(ZIP_PO_RANGE)}?key=${API_KEY}`;
      const zipRes = await fetchWithRetry(zipUrl);
      let poLots = new Set();
      let poQtyMap = new Map(); // lotNo -> cumulative total PO pieces created

      if (zipRes.ok) {
        const zipData = await zipRes.json();
        const zipValues = zipData.values || [];
        if (zipValues.length > 1) {
          const zipHeaders = zipValues[0].map((h) => normalize(h));
          const lotColIdx = zipHeaders.findIndex((h) => h.includes("lot"));
          const pcsColIdx = zipHeaders.findIndex((h) => h.includes("pieces") || h.includes("pcs") || h.includes("total pcs") || h.includes("qty"));

          if (lotColIdx !== -1) {
            zipValues.slice(1).forEach((r) => {
              const lotNo = normalize(r[lotColIdx]);
              if (lotNo) {
                poLots.add(lotNo);
                const pcs = pcsColIdx !== -1 ? Number((r[pcsColIdx] || "").toString().replace(/,/g, "").trim()) || 0 : 0;
                poQtyMap.set(lotNo, (poQtyMap.get(lotNo) || 0) + pcs);
              }
            });
          }
        }
      }
      setExistingPOLots(poLots);

      // 1b. Fetch PO As Per Shade sheet and identify lots where Supplier Name contains 'ZIP'
      let shadeZipPOLots = new Set();
      try {
        let shadePoValues = [];
        const backendRes = await fetchSheetDataFromBackend(SHADE_PO_SPREADSHEET_ID, SHADE_PO_RANGE);
        if (backendRes && backendRes.ok && Array.isArray(backendRes.values) && backendRes.values.length > 0) {
          shadePoValues = backendRes.values;
        } else {
          const shadePoUrl = `https://sheets.googleapis.com/v4/spreadsheets/${SHADE_PO_SPREADSHEET_ID}/values/${encodeURIComponent(SHADE_PO_RANGE)}?key=${API_KEY}`;
          const shadePoRes = await fetchWithRetry(shadePoUrl);
          if (shadePoRes.ok) {
            const shadePoData = await shadePoRes.json();
            shadePoValues = shadePoData.values || [];
          }
        }

        if (shadePoValues.length > 1) {
          const shadeHeaders = shadePoValues[0].map((h) => normalize(h));
          let lotColIdx = shadeHeaders.findIndex((h) => h === "lot number" || h === "lot no" || h.includes("lot"));
          let supplierColIdx = shadeHeaders.findIndex((h) => h === "supplier name" || h === "supplier" || h.includes("supplier"));

          if (lotColIdx === -1) lotColIdx = 1;
          if (supplierColIdx === -1) supplierColIdx = 4;

          shadePoValues.slice(1).forEach((r) => {
            const supplier = normalize(r[supplierColIdx]);
            const rawLot = (r[lotColIdx] || "").toString().trim();
            if (supplier.includes("zip") && rawLot) {
              shadeZipPOLots.add(normalize(rawLot));
              rawLot.split(/[,;\s]+/).forEach((part) => {
                const normPart = normalize(part);
                if (normPart) shadeZipPOLots.add(normPart);
              });
            }
          });
        }
      } catch (shadeErr) {
        console.warn("Could not fetch Shade PO sheet:", shadeErr);
      }

      // 2. Fetch Index Sheet Cut Lots, Cutting Qty, Completed Lots, Supervisors & Parties
      const indexUrl = `https://sheets.googleapis.com/v4/spreadsheets/${INDEX_SPREADSHEET_ID}/values/${encodeURIComponent(INDEX_RANGE)}?key=${API_KEY}`;
      const indexRes = await fetchWithRetry(indexUrl);
      let compLots = new Set();
      let cutLotsMap = new Map();
      if (indexRes.ok) {
        const indexData = await indexRes.json();
        const indexValues = indexData.values || [];
        if (indexValues.length > 1) {
          const indexHeaders = indexValues[0].map((h) => normalize(h));
          const lotColIdx = indexHeaders.findIndex((h) => h === "lot number" || h.includes("lot"));
          const compColIdx = indexHeaders.findIndex((h) => h.includes("completed status") || h.includes("completion"));
          const supColIdx = indexHeaders.findIndex((h) => h === "supervisor" || h.includes("superv"));
          const partyColIdx = indexHeaders.findIndex((h) => h === "party name" || h.includes("party"));
          const cutQtyColIdx = indexHeaders.findIndex((h) => h === "cutting qty" || h.includes("cutting"));

          if (lotColIdx !== -1) {
            indexValues.slice(1).forEach((r) => {
              const lotNo = normalize(r[lotColIdx]);
              if (lotNo) {
                const supVal = supColIdx !== -1 ? (r[supColIdx] || "").toString().trim() : "";
                const partyVal = partyColIdx !== -1 ? (r[partyColIdx] || "").toString().trim() : "";
                const cutQtyVal = cutQtyColIdx !== -1 ? Number((r[cutQtyColIdx] || "").toString().replace(/,/g, "").trim()) || 0 : 0;

                cutLotsMap.set(lotNo, { supervisor: supVal, partyName: partyVal, cuttingQty: cutQtyVal });

                if (compColIdx !== -1 && isCompletedWithDate(r[compColIdx])) {
                  compLots.add(lotNo);
                }
              }
            });
          }
        }
      }
      setCompletedIndexLots(compLots);

      // 3. Fetch Job Orders
      const joUrl = `https://sheets.googleapis.com/v4/spreadsheets/${JOB_ORDERS_SPREADSHEET_ID}/values/${encodeURIComponent(JOB_ORDERS_RANGE)}?key=${API_KEY}`;
      const joRes = await fetchWithRetry(joUrl);
      if (!joRes.ok) throw new Error(`HTTP Error ${joRes.status}`);

      const joData = await joRes.json();
      const values = joData.values || [];
      if (values.length < 2) throw new Error("No Job Orders found.");

      const headers = values[0].map((h) => (h || "").trim());
      const hIndex = {};
      headers.forEach((h, i) => {
        const exact = h.trim();
        const norm = normalize(h);
        if (!(exact in hIndex)) hIndex[exact] = i;
        if (!(norm in hIndex)) hIndex[norm] = i;
      });

      const parsedRows = values
        .slice(1)
        .filter((r) => r.some((cell) => (cell ?? "").toString().trim() !== ""))
        .map((r, i) => {
          const getVal = (colName) => (r[hIndex[colName]] ?? "").toString().trim();
          const zipVal = getVal("Zip") || getVal("Zipper");
          const lotNo = getVal("Lot Number");
          const normLot = normalize(lotNo);
          const req = isZipRequired(zipVal);

          const indexInfo = normLot ? cutLotsMap.get(normLot) : null;
          const isCutInIndex = !!indexInfo;
          const indexSupervisor = indexInfo?.supervisor || "";
          const indexPartyName = indexInfo?.partyName || "";

          // Fetch Cutting Qty from Index Sheet (fallback to JobOrder Quantity if 0)
          const indexCuttingQty = indexInfo?.cuttingQty || 0;
          const jobOrderQty = Number((getVal("Quantity") || "").replace(/,/g, "")) || 0;
          const cuttingQty = indexCuttingQty > 0 ? indexCuttingQty : jobOrderQty;

          // Accumulated PO Pieces created so far
          const createdPOQty = normLot ? (poQtyMap.get(normLot) || 0) : 0;
          const pendingPOQty = Math.max(0, cuttingQty - createdPOQty);
          const hasFullPO = createdPOQty >= cuttingQty && cuttingQty > 0;

          return {
            id: i + 1,
            jobOrderNo: getVal("Job Order No"),
            date: getVal("Date"),
            fabric: getVal("Fabric"),
            brand: getVal("Brand"),
            shade: getVal("Shade"),
            size: getVal("Size"),
            quantity: cuttingQty,
            cuttingQty: cuttingQty,
            createdPOQty: createdPOQty,
            pendingPOQty: pendingPOQty,
            unit: "PCS",
            partyName: getVal("Party Name") || indexPartyName,
            garmentType: getVal("Garment Type"),
            section: getVal("Section"),
            season: getVal("Season"),
            pattern: getVal("Pattern"),
            style: getVal("Style"),
            remarks: getVal("Remarks"),
            submittedBy: getVal("Submitted By"),
            lotNumber: lotNo,
            directStitching: getVal("Direct Stitching") || getVal("Direct_Stitching") || "—",
            priority: getVal("Priority") || "NORMAL",
            zipDetails: zipVal,
            orderNo: getVal("Order No."),
            supervisor: getVal("FABRIC_SUPERVISOR") || indexSupervisor,
            indexSupervisor,
            indexPartyName,
            isZipReq: req,
            isCutInIndex: isCutInIndex,
            hasPO: createdPOQty > 0,
            hasFullPO: hasFullPO,
            isCompletedInIndex: normLot ? compLots.has(normLot) : false,
            financialYear: getFinancialYear(getVal("Date"))
          };
        });

      // Include ONLY lots where LOT IS CUT (present in Index), ZIP IS REQUIRED, PENDING PO QTY > 0, NOT COMPLETED, NOT DUSHYANT / JAINHOSIERY, AND NOT PRESENT IN SHADE PO SHEET WITH ZIP SUPPLIER
      const pendingZipRows = parsedRows.filter(
        (r) =>
          r.isCutInIndex &&
          r.isZipReq &&
          !r.hasFullPO &&
          r.pendingPOQty > 0 &&
          !r.isCompletedInIndex &&
          !isExcludedSupervisorOrParty(r) &&
          !shadeZipPOLots.has(normalize(r.lotNumber))
      );
      setJobOrders(pendingZipRows);
    } catch (err) {
      console.error("Failed to load Pending Zip PO data:", err);
      setError(err.message || "Failed to load data from Google Sheets.");
    } finally {
      setLoading(false);
      setRefreshing(false);
    }
  };

  useEffect(() => {
    fetchData();
  }, []);

  // Filter Unique Options
  const partyOptions = useMemo(() => Array.from(new Set(jobOrders.map((r) => r.partyName).filter(Boolean))).sort(), [jobOrders]);
  const fabricOptions = useMemo(() => Array.from(new Set(jobOrders.map((r) => r.fabric).filter(Boolean))).sort(), [jobOrders]);
  const garmentOptions = useMemo(() => Array.from(new Set(jobOrders.map((r) => r.garmentType).filter(Boolean))).sort(), [jobOrders]);
  const priorityOptions = useMemo(() => Array.from(new Set(jobOrders.map((r) => r.priority).filter(Boolean))).sort(), [jobOrders]);
  const fyOptions = useMemo(() => Array.from(new Set(jobOrders.map((r) => r.financialYear).filter(Boolean))).sort((a, b) => b.localeCompare(a)), [jobOrders]);

  // Filtered Job Orders
  const filteredOrders = useMemo(() => {
    const q = searchQuery.toLowerCase().trim();
    return jobOrders.filter((r) => {
      const matchSearch =
        !q ||
        r.jobOrderNo.toLowerCase().includes(q) ||
        r.lotNumber.toLowerCase().includes(q) ||
        r.partyName.toLowerCase().includes(q) ||
        r.fabric.toLowerCase().includes(q) ||
        r.zipDetails.toLowerCase().includes(q) ||
        r.garmentType.toLowerCase().includes(q);

      const matchParty = selectedParties.length === 0 || selectedParties.includes(r.partyName);
      const matchFabric = selectedFabrics.length === 0 || selectedFabrics.includes(r.fabric);
      const matchGarment = selectedGarments.length === 0 || selectedGarments.includes(r.garmentType);
      const matchPriority = selectedPriorities.length === 0 || selectedPriorities.includes(r.priority);
      const matchFY = selectedFYs.length === 0 || selectedFYs.includes(r.financialYear);

      return matchSearch && matchParty && matchFabric && matchGarment && matchPriority && matchFY;
    });
  }, [jobOrders, searchQuery, selectedParties, selectedFabrics, selectedGarments, selectedPriorities, selectedFYs]);

  // Stats Calculations
  const stats = useMemo(() => {
    const totalLots = filteredOrders.length;
    const totalCuttingQty = filteredOrders.reduce((sum, r) => sum + r.cuttingQty, 0);
    const totalPOCreatedQty = filteredOrders.reduce((sum, r) => sum + r.createdPOQty, 0);
    const totalPendingPOQty = filteredOrders.reduce((sum, r) => sum + r.pendingPOQty, 0);
    const highPriority = filteredOrders.filter((r) => normalize(r.priority).startsWith("h")).length;
    const partiesCount = new Set(filteredOrders.map((r) => r.partyName).filter(Boolean)).size;

    return { totalLots, totalCuttingQty, totalPOCreatedQty, totalPendingPOQty, highPriority, partiesCount };
  }, [filteredOrders]);

  // Pagination Slice
  const totalPages = Math.max(1, Math.ceil(filteredOrders.length / pageSize));
  const currentPageSafe = Math.min(currentPage, totalPages);
  const paginatedOrders = useMemo(() => {
    const start = (currentPageSafe - 1) * pageSize;
    return filteredOrders.slice(start, start + pageSize);
  }, [filteredOrders, currentPageSafe, pageSize]);

  // Toggle Row Expansion
  const toggleRow = (id) => {
    setExpandedRows((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  // Clear Filters
  const clearFilters = () => {
    setSearchQuery("");
    setSelectedParties([]);
    setSelectedFabrics([]);
    setSelectedGarments([]);
    setSelectedPriorities([]);
    setSelectedFYs([]);
    setCurrentPage(1);
  };

  // Export CSV
  const exportCSV = () => {
    const headers = [
      "Lot No",
      "Garment Type",
      "Style",
      "Fabric",
      "Brand",
      "Pcs",
      "Section",
      "Season",
      "Party Name",
      "Direct Stitching",
      "Job Order No",
      "Date",
      "Shade",
      "PO Created Qty (PCS)",
      "Pending PO Qty (PCS Left)",
      "Zip Details",
      "Priority",
      "Submitted By",
      "Financial Year",
      "PO Status"
    ];

    const rows = filteredOrders.map((r) => [
      r.lotNumber,
      r.garmentType,
      r.style,
      r.fabric,
      r.brand,
      r.cuttingQty,
      r.section,
      r.season,
      r.partyName,
      r.directStitching,
      r.jobOrderNo,
      r.date,
      r.shade,
      r.createdPOQty,
      r.pendingPOQty,
      r.zipDetails,
      r.priority,
      r.submittedBy,
      r.financialYear,
      r.createdPOQty > 0 ? `Partial PO (${r.createdPOQty}/${r.cuttingQty})` : "PO NOT CREATED"
    ]);

    const csvContent = [headers.join(","), ...rows.map((e) => e.map((val) => `"${String(val || "").replace(/"/g, '""')}"`).join(","))].join("\n");
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `Pending_Zip_PO_Report_${new Date().toISOString().slice(0, 10)}.csv`;
    link.click();
    URL.revokeObjectURL(url);
  };

  // Export Monochrome Black & White PDF
  const exportPDF = () => {
    const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a3" });
    
    // Clean Header (White Background with Black Text)
    doc.setFont("helvetica", "bold");
    doc.setFontSize(14);
    doc.setTextColor(0, 0, 0);
    doc.text("MH FACTORY SUITE PRO — PENDING ZIP PO REPORT", 12, 13);

    doc.setFont("helvetica", "normal");
    doc.setFontSize(8.5);
    doc.setTextColor(60, 60, 60);
    doc.text("Cut production lots requiring Zipper Purchase Orders where PO has not been fully created", 12, 19);

    // Meta Information Badge (Right Top - Black Text)
    doc.setFont("helvetica", "bold");
    doc.setFontSize(8.5);
    doc.setTextColor(0, 0, 0);
    doc.text(`DATE: ${new Date().toLocaleDateString("en-IN")}`, 400, 11, { align: "right" });
    doc.setFont("helvetica", "normal");
    doc.text(`TOTAL PENDING LOTS: ${filteredOrders.length}`, 400, 16, { align: "right" });
    doc.text(`TOTAL PENDING PCS: ${stats.totalPendingPOQty.toLocaleString("en-IN")} PCS`, 400, 21, { align: "right" });

    // Solid Black Header Line
    doc.setDrawColor(0, 0, 0);
    doc.setLineWidth(0.5);
    doc.line(12, 24, 400, 24);

    const tableData = filteredOrders.map((r, i) => [
      i + 1,
      r.lotNumber || "-",
      r.garmentType || "-",
      r.style || "-",
      r.fabric || "-",
      r.brand || "-",
      `${r.cuttingQty?.toLocaleString("en-IN")} PCS`,
      r.section || "-",
      r.season || "-",
      r.partyName || "-",
      r.directStitching || "-",
      r.jobOrderNo || "-",
      r.date || "-",
      `${r.createdPOQty > 0 ? r.createdPOQty.toLocaleString("en-IN") : "0"} PCS`,
      `${r.pendingPOQty?.toLocaleString("en-IN")} PCS Left`,
      r.zipDetails || "-",
      r.priority || "NORMAL",
      r.createdPOQty > 0 ? `Partial (${r.createdPOQty}/${r.cuttingQty})` : "PO NOT CREATED"
    ]);

    autoTable(doc, {
      head: [["#", "Lot No", "Garment Type", "Style", "Fabric", "Brand", "Pcs", "Section", "Season", "Party Name", "Direct Stitching", "JO No", "Date", "PO Created", "Pending Qty", "Zip Requirement", "Priority", "PO Status"]],
      body: tableData,
      startY: 27,
      margin: { top: 27, bottom: 15, left: 12, right: 12 },
      styles: {
        fontSize: 7.5,
        cellPadding: 2,
        textColor: [0, 0, 0],
        valign: "middle",
        lineColor: [0, 0, 0],
        lineWidth: 0.15
      },
      headStyles: {
        fillColor: [0, 0, 0],
        textColor: [255, 255, 255],
        fontStyle: "bold",
        lineColor: [0, 0, 0],
        lineWidth: 0.2,
        halign: "center"
      },
      alternateRowStyles: { fillColor: [248, 248, 248] },
      didDrawPage: (data) => {
        // Sub-page header
        if (data.pageNumber > 1) {
          doc.setFont("helvetica", "bold");
          doc.setFontSize(9.5);
          doc.setTextColor(0, 0, 0);
          doc.text("MH FACTORY SUITE PRO — PENDING ZIP PO REPORT (CONTINUED)", 12, 10);
          doc.setDrawColor(0, 0, 0);
          doc.setLineWidth(0.4);
          doc.line(12, 12, 400, 12);
        }

        // Footer rule and text
        const totalPages = doc.internal.getNumberOfPages();
        doc.setDrawColor(0, 0, 0);
        doc.setLineWidth(0.3);
        doc.line(12, 280, 400, 280);

        doc.setFont("helvetica", "normal");
        doc.setFontSize(7.5);
        doc.setTextColor(0, 0, 0);
        doc.text("MH Factory Suite Pro — Internal Production Report", 12, 285);
        doc.text(`Page ${data.pageNumber} of ${totalPages}`, 400, 285, { align: "right" });
      }
    });

    doc.save(`Pending_Zip_PO_Report_${new Date().toISOString().slice(0, 10)}.pdf`);
  };

  return (
    <>
      <style>{`
        .zpr-container {
          max-width: 2400px;
          margin: 0 auto;
          padding: 1.5rem 1rem;
          font-family: 'Plus Jakarta Sans', system-ui, -apple-system, sans-serif;
          background: #f8fafc;
          min-height: 100vh;
          color: #0f172a;
        }
        @media (min-width: 768px) { .zpr-container { padding: 2rem; } }

        .zpr-header {
          background: linear-gradient(135deg, #1e1b4b 0%, #312e81 50%, #4338ca 100%);
          border-radius: 16px;
          padding: 1.25rem 1.5rem;
          color: #ffffff;
          box-shadow: 0 10px 25px -5px rgba(49, 46, 129, 0.25);
          display: flex;
          flex-wrap: wrap;
          align-items: center;
          justify-content: space-between;
          gap: 1rem;
        }
        .zpr-title {
          font-size: 1.4rem;
          font-weight: 800;
          margin: 0;
          display: flex;
          align-items: center;
          gap: 0.6rem;
        }
        .zpr-subtitle {
          font-size: 0.9rem;
          opacity: 0.9;
          margin-top: 0.2rem;
          font-weight: 500;
        }

        .zpr-actions {
          display: flex;
          align-items: center;
          gap: 0.75rem;
          flex-wrap: wrap;
        }
        .zpr-btn {
          display: inline-flex;
          align-items: center;
          gap: 0.5rem;
          background: #ffffff;
          color: #312e81;
          border: 0;
          border-radius: 10px;
          padding: 0.65rem 1.1rem;
          font-weight: 700;
          font-size: 0.88rem;
          cursor: pointer;
          box-shadow: 0 4px 10px rgba(0, 0, 0, 0.1);
          transition: all 0.15s ease;
        }
        .zpr-btn:hover { background: #e0e7ff; transform: translateY(-1px); }
        .zpr-btn--ghost {
          background: rgba(255, 255, 255, 0.15);
          color: #ffffff;
          border: 1px solid rgba(255, 255, 255, 0.3);
        }
        .zpr-btn--ghost:hover { background: rgba(255, 255, 255, 0.25); color: #ffffff; }

        /* Stats Grid */
        .zpr-stats-grid {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
          gap: 1rem;
          margin-top: 1.25rem;
        }
        .zpr-stat-card {
          background: #ffffff;
          border: 1px solid #e2e8f0;
          border-radius: 14px;
          padding: 1.1rem;
          box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.04);
          display: flex;
          align-items: center;
          gap: 1rem;
        }
        .zpr-stat-icon {
          width: 48px;
          height: 48px;
          border-radius: 12px;
          display: grid;
          place-items: center;
          font-size: 1.4rem;
        }
        .zpr-stat-val { font-size: 1.5rem; font-weight: 800; color: #0f172a; line-height: 1.1; }
        .zpr-stat-lbl { font-size: 0.8rem; font-weight: 600; color: #64748b; text-transform: uppercase; margin-top: 0.2rem; }

        /* Filters Panel */
        .zpr-panel {
          background: #ffffff;
          border: 1px solid #e2e8f0;
          border-radius: 16px;
          padding: 1.25rem;
          margin-top: 1.25rem;
          box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.04);
        }
        .zpr-filters-grid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(210px, 1fr));
          gap: 0.85rem;
          align-items: end;
        }
        .zpr-input, .zpr-select {
          width: 100%;
          padding: 0.65rem 0.9rem;
          border: 1px solid #cbd5e1;
          border-radius: 10px;
          background: #ffffff;
          font-size: 0.9rem;
          font-weight: 500;
          color: #1e293b;
          box-sizing: border-box;
        }
        .zpr-input:focus, .zpr-select:focus { outline: none; border-color: #4338ca; box-shadow: 0 0 0 3px rgba(67, 56, 202, 0.18); }

        /* Table Styling */
        .zpr-tablewrap {
          overflow-x: auto;
          border: 1px solid #e2e8f0;
          border-radius: 16px;
          background: #ffffff;
          margin-top: 1.25rem;
          box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.04);
        }
        .zpr-table { width: 100%; min-width: 1200px; border-collapse: collapse; font-size: 0.9rem; }
        .zpr-th {
          background: #1e1b4b;
          color: #ffffff;
          padding: 0.9rem 1rem;
          text-align: left;
          font-weight: 700;
          font-size: 0.8rem;
          text-transform: uppercase;
          letter-spacing: 0.04em;
          white-space: nowrap;
        }
        .zpr-td { padding: 0.85rem 1rem; border-bottom: 1px solid #f1f5f9; color: #1e293b; font-weight: 500; }
        .zpr-tr:nth-child(even) { background: #f8fafc; }
        .zpr-tr:hover { background: #f1f5f9; }

        .zpr-badge-pending {
          display: inline-flex;
          align-items: center;
          gap: 0.35rem;
          background: #ffe4e6;
          color: #9f1239;
          border: 1px solid #fecdd3;
          border-radius: 999px;
          padding: 0.25rem 0.65rem;
          font-size: 0.78rem;
          font-weight: 800;
        }
        .zpr-badge-zip {
          display: inline-flex;
          align-items: center;
          gap: 0.35rem;
          background: #f0f9ff;
          color: #0369a1;
          border: 1px solid #bae6fd;
          border-radius: 8px;
          padding: 0.2rem 0.55rem;
          font-size: 0.82rem;
          font-weight: 700;
        }

        .zpr-pager {
          display: flex;
          align-items: center;
          justify-content: space-between;
          padding: 1rem 0.5rem;
          flex-wrap: wrap;
          gap: 1rem;
        }
      `}</style>

      <div className="zpr-container">
        {/* Header */}
        <header className="zpr-header">
          <div>
            <h1 className="zpr-title">🤐 Pending Zip PO Report</h1>
            <div className="zpr-subtitle">Lots with Zip required where Purchase Orders (PO) have NOT been created</div>
          </div>
          <div className="zpr-actions">
            <button className="zpr-btn zpr-btn--ghost" onClick={() => history.push("/dashboard")}>
              <span>⬅️</span>
              <span>Back</span>
            </button>
            <button className="zpr-btn zpr-btn--ghost" onClick={() => fetchData({ isRefresh: true })} disabled={loading || refreshing}>
              <span>{refreshing ? "⏳" : "🔄"}</span>
              <span>{refreshing ? "Refreshing…" : "Refresh"}</span>
            </button>
            <button className="zpr-btn" onClick={() => history.push("/zip-report")}>
              <span>📦</span>
              <span>Create Zip PO</span>
            </button>
            <button className="zpr-btn zpr-btn--ghost" onClick={exportCSV} disabled={filteredOrders.length === 0}>
              <span>📊</span>
              <span>Export CSV</span>
            </button>
            <button className="zpr-btn zpr-btn--ghost" onClick={exportPDF} disabled={filteredOrders.length === 0}>
              <span>📄</span>
              <span>Export PDF</span>
            </button>
          </div>
        </header>

        {/* Stats Grid */}
        <div className="zpr-stats-grid">
          <div className="zpr-stat-card">
            <div className="zpr-stat-icon" style={{ background: "#ffe4e6", color: "#be123c" }}>🤐</div>
            <div>
              <div className="zpr-stat-val">{stats.totalLots.toLocaleString("en-IN")}</div>
              <div className="zpr-stat-lbl">Pending Zip Lots</div>
            </div>
          </div>
          <div className="zpr-stat-card">
            <div className="zpr-stat-icon" style={{ background: "#fef2f2", color: "#dc2626" }}>⚡</div>
            <div>
              <div className="zpr-stat-val">{stats.highPriority.toLocaleString("en-IN")}</div>
              <div className="zpr-stat-lbl">High Priority Lots</div>
            </div>
          </div>
          <div className="zpr-stat-card">
            <div className="zpr-stat-icon" style={{ background: "#f0fdf4", color: "#166534" }}>🏢</div>
            <div>
              <div className="zpr-stat-val">{stats.partiesCount.toLocaleString("en-IN")}</div>
              <div className="zpr-stat-lbl">Parties Affected</div>
            </div>
          </div>
        </div>

        {/* Filters Section */}
        <section className="zpr-panel">
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1rem" }}>
            <h3 style={{ margin: 0, fontWeight: 700, fontSize: "1.1rem" }}>🔍 Filters & Search</h3>
            <button className="zpr-btn zpr-btn--ghost" style={{ color: "#475569", borderColor: "#cbd5e1" }} onClick={clearFilters}>
              🔄 Reset Filters
            </button>
          </div>

          <div className="zpr-filters-grid">
            <input
              className="zpr-input"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="Search Lot, JO No, Party, Fabric, Zip..."
            />
            <MultiSelectDropdown
              title="Party"
              icon="👥"
              options={partyOptions}
              selected={selectedParties}
              onChange={setSelectedParties}
            />
            <MultiSelectDropdown
              title="Fabric"
              icon="👗"
              options={fabricOptions}
              selected={selectedFabrics}
              onChange={setSelectedFabrics}
            />
            <MultiSelectDropdown
              title="Garment Type"
              icon="👕"
              options={garmentOptions}
              selected={selectedGarments}
              onChange={setSelectedGarments}
            />
            <MultiSelectDropdown
              title="Priority"
              icon="⚡"
              options={priorityOptions}
              selected={selectedPriorities}
              onChange={setSelectedPriorities}
            />
            <MultiSelectDropdown
              title="Financial Year"
              icon="📅"
              options={fyOptions}
              selected={selectedFYs}
              onChange={setSelectedFYs}
            />
          </div>
        </section>

        {/* Status / Errors */}
        {error && (
          <div style={{ padding: "1rem", background: "#fee2e2", color: "#b91c1c", borderRadius: "12px", marginTop: "1rem", fontWeight: 600 }}>
            ⚠️ {error}
          </div>
        )}

        {/* Data Table */}
        <div className="zpr-tablewrap">
          {loading ? (
            <div style={{ padding: "3rem", textAlign: "center", color: "#64748b", fontWeight: 600 }}>
              ⏳ Loading Pending Zip PO Data...
            </div>
          ) : paginatedOrders.length === 0 ? (
            <div style={{ padding: "3rem", textAlign: "center", color: "#64748b", fontWeight: 600 }}>
              🎉 Excellent! No pending Zip PO lots found for the selected criteria.
            </div>
          ) : (
            <table className="zpr-table">
              <thead>
                <tr>
                  <th className="zpr-th" style={{ width: "40px" }}></th>
                  <th className="zpr-th">Lot No</th>
                  <th className="zpr-th">Garment Type</th>
                  <th className="zpr-th">Style</th>
                  <th className="zpr-th">Fabric</th>
                  <th className="zpr-th">Brand</th>
                  <th className="zpr-th">Pcs</th>
                  <th className="zpr-th">Section</th>
                  <th className="zpr-th">Season</th>
                  <th className="zpr-th">Party Name</th>
                  <th className="zpr-th">Direct Stitching</th>
                  <th className="zpr-th">JO No</th>
                  <th className="zpr-th">Date</th>
                  <th className="zpr-th">PO Created Qty</th>
                  <th className="zpr-th">Pending PO (Left)</th>
                  <th className="zpr-th">Zip Requirement</th>
                  <th className="zpr-th">Priority</th>
                  <th className="zpr-th">PO Status</th>
                </tr>
              </thead>
              <tbody>
                {paginatedOrders.map((r) => {
                  const isExp = expandedRows.has(r.id);
                  return (
                    <React.Fragment key={r.id}>
                      <tr className="zpr-tr">
                        <td className="zpr-td" style={{ textAlign: "center", cursor: "pointer" }} onClick={() => toggleRow(r.id)}>
                          {isExp ? "▼" : "▶"}
                        </td>
                        <td className="zpr-td" style={{ fontWeight: 700 }}>{r.lotNumber || "-"}</td>
                        <td className="zpr-td">{r.garmentType || "-"}</td>
                        <td className="zpr-td">{r.style || "-"}</td>
                        <td className="zpr-td">{r.fabric || "-"}</td>
                        <td className="zpr-td">{r.brand || "-"}</td>
                        <td className="zpr-td" style={{ fontWeight: 800, color: "#1e1b4b" }}>
                          {r.cuttingQty?.toLocaleString("en-IN")} PCS
                        </td>
                        <td className="zpr-td">{r.section || "-"}</td>
                        <td className="zpr-td">{r.season || "-"}</td>
                        <td className="zpr-td">{r.partyName || "-"}</td>
                        <td className="zpr-td">{r.directStitching || "-"}</td>
                        <td className="zpr-td" style={{ fontWeight: 800, color: "#4338ca" }}>{r.jobOrderNo || "-"}</td>
                        <td className="zpr-td">{r.date || "-"}</td>
                        <td className="zpr-td" style={{ fontWeight: 700, color: "#15803d" }}>
                          {r.createdPOQty > 0 ? `${r.createdPOQty.toLocaleString("en-IN")} PCS` : "0 PCS"}
                        </td>
                        <td className="zpr-td">
                          <span className="zpr-badge-pending" style={{ background: "#fff7ed", color: "#c2410c", borderColor: "#ffedd5" }}>
                            ⏳ {r.pendingPOQty?.toLocaleString("en-IN")} PCS Left for Order
                          </span>
                        </td>
                        <td className="zpr-td">
                          <span className="zpr-badge-zip">🤐 {r.zipDetails}</span>
                        </td>
                        <td className="zpr-td" style={{ fontWeight: 700 }}>{r.priority}</td>
                        <td className="zpr-td">
                          {r.createdPOQty > 0 ? (
                            <span className="zpr-badge-zip" style={{ background: "#fef3c7", color: "#92400e", borderColor: "#fde68a" }}>
                              ⚡ Partial PO ({r.createdPOQty} / {r.cuttingQty} PCS)
                            </span>
                          ) : (
                            <span className="zpr-badge-pending">⚠️ PO NOT CREATED</span>
                          )}
                        </td>
                      </tr>

                      {/* Expander Row */}
                      {isExp && (
                        <tr style={{ background: "#f8fafc" }}>
                          <td></td>
                          <td colSpan={17} style={{ padding: "1rem", borderBottom: "1px solid #e2e8f0" }}>
                            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: "0.75rem" }}>
                              <div><strong>Pattern:</strong> {r.pattern || "—"}</div>
                              <div><strong>Shade:</strong> {r.shade || "—"}</div>
                              <div><strong>Submitted By:</strong> {r.submittedBy || "—"}</div>
                              <div><strong>Order No:</strong> {r.orderNo || "—"}</div>
                              <div><strong>Supervisor:</strong> {r.supervisor || "—"}</div>
                              <div><strong>Remarks:</strong> {r.remarks || "—"}</div>
                            </div>
                          </td>
                        </tr>
                      )}
                    </React.Fragment>
                  );
                })}
              </tbody>
            </table>
          )}
        </div>

        {/* Pagination */}
        {!loading && filteredOrders.length > 0 && (
          <div className="zpr-pager">
            <div style={{ color: "#64748b", fontWeight: 600, fontSize: "0.9rem" }}>
              📊 Showing {(currentPageSafe - 1) * pageSize + 1} - {Math.min(currentPageSafe * pageSize, filteredOrders.length)} of {filteredOrders.length} lots
            </div>
            <div style={{ display: "flex", gap: "0.5rem", alignItems: "center" }}>
              <button
                className="zpr-btn zpr-btn--ghost"
                style={{ color: "#1e293b", borderColor: "#cbd5e1" }}
                disabled={currentPageSafe === 1}
                onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
              >
                ⬅️ Previous
              </button>
              <span style={{ fontWeight: 700, padding: "0 0.5rem" }}>Page {currentPageSafe} of {totalPages}</span>
              <button
                className="zpr-btn zpr-btn--ghost"
                style={{ color: "#1e293b", borderColor: "#cbd5e1" }}
                disabled={currentPageSafe === totalPages}
                onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
              >
                Next ➡️
              </button>
            </div>
          </div>
        )}
      </div>
    </>
  );
};

export default PendingZipPOReport;
