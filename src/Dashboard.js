import React, { useState, useEffect, useMemo, useRef } from "react";
import { Link, useHistory } from "react-router-dom";
import { motion, AnimatePresence } from "framer-motion";
import { getCurrentUser, logoutUser } from "./auth";

export default function Dashboard() {
  const history = useHistory();
  const currentUser = getCurrentUser();
  const isAdmin = (currentUser?.role || "").trim().toLowerCase() === "admin";

  const handleLogout = () => {
    logoutUser();
    history.push("/");
  };

  const [activeSectionTab, setActiveSectionTab] = useState("all"); // 'all' | 'critical' | 'departments' | 'admin' | 'other' | 'favorites'
  const [searchQuery, setSearchQuery] = useState("");
  const [viewMode, setViewMode] = useState("grid"); // 'grid' | 'list'
  const [showBackToTop, setShowBackToTop] = useState(false);
  const searchInputRef = useRef(null);

  const [pinnedModules, setPinnedModules] = useState(() => {
    try {
      const saved = localStorage.getItem("pinned_dashboard_modules");
      return saved ? JSON.parse(saved) : ["Cutting Report", "Overall Stitching Report", "Cut to Pack Report"];
    } catch {
      return ["Cutting Report", "Overall Stitching Report", "Cut to Pack Report"];
    }
  });

  const [hoveredCard, setHoveredCard] = useState(null);

  // Keyboard shortcut for fast search focus (Ctrl+K or /)
  useEffect(() => {
    const handleKeyDown = (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "k") {
        e.preventDefault();
        searchInputRef.current?.focus();
      } else if (e.key === "/" && document.activeElement !== searchInputRef.current && document.activeElement?.tagName !== "INPUT") {
        e.preventDefault();
        searchInputRef.current?.focus();
      }
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, []);

  // Scroll listener for floating back-to-top button
  useEffect(() => {
    const handleScroll = () => {
      setShowBackToTop(window.scrollY > 350);
    };
    window.addEventListener("scroll", handleScroll);
    return () => window.removeEventListener("scroll", handleScroll);
  }, []);

  // Save pinned modules to localStorage
  useEffect(() => {
    try {
      localStorage.setItem("pinned_dashboard_modules", JSON.stringify(pinnedModules));
    } catch (e) {
      console.error("Failed to save pinned modules:", e);
    }
  }, [pinnedModules]);

  const togglePin = (title, e) => {
    e.preventDefault();
    e.stopPropagation();
    setPinnedModules((prev) =>
      prev.includes(title) ? prev.filter((t) => t !== title) : [...prev, title]
    );
  };

  // ============================================================================
  // VIBRANT COLORFUL MASTER MODULE DATA ACROSS 4 OPERATIONAL TIERS
  // ============================================================================

  // 1. SECTION: Most Critical Reports – Necessity Daily Checking Reports (10 Reports)
  const criticalReports = useMemo(() => [
    {
      id: "crit-1",
      step: 1,
      title: "Cutting Report",
      emoji: "✂️",
      path: "/cutting-report",
      section: "critical",
      gradient: "linear-gradient(135deg, #059669 0%, #10b981 100%)",
      glowColor: "rgba(16, 185, 129, 0.45)",
      badgeColor: "#059669",
      accentBg: "#ecfdf5",
      accentBorder: "#a7f3d0",
      description: "Cutting operations master report with fabric consumption, table-wise layer plies, and marker efficiency.",
      whyEssential: "Validates daily cutting yardage against BOM standards and catches layer wastage before pieces leave the cutting hall.",
      department: "Cutting",
      features: ["Fabric Consumption", "Layer Breakdown", "Marker Efficiency", "Cut Pcs Count"],
      lastUpdated: "Real-time"
    },
    {
      id: "crit-2",
      step: 2,
      title: "Embroidery Report",
      emoji: "🧵",
      path: "/embroidery",
      section: "critical",
      gradient: "linear-gradient(135deg, #7c3aed 0%, #a855f7 100%)",
      glowColor: "rgba(124, 58, 237, 0.45)",
      badgeColor: "#7c3aed",
      accentBg: "#f5f3ff",
      accentBorder: "#ddd6fe",
      description: "Create and manage embroidery production challans with machine stitch tracking, thread usage, and vendor dispatch.",
      whyEssential: "Prevents outbound delivery delays by tracking embroidery challan issuance and vendor completion timelines.",
      department: "Embroidery & Printing",
      features: ["Batch Tracking", "Quality Check", "Thread Usage", "Challan Issue"],
      lastUpdated: "Today"
    },
    {
      id: "crit-3",
      step: 3,
      title: "Printing Report",
      emoji: "🖨️",
      path: "/printing",
      section: "critical",
      gradient: "linear-gradient(135deg, #ec4899 0%, #f43f5e 100%)",
      glowColor: "rgba(236, 72, 153, 0.45)",
      badgeColor: "#db2777",
      accentBg: "#fdf2f8",
      accentBorder: "#fbcfe8",
      description: "Screen printing job management, color registration tracking, table allotment, and daily output logs.",
      whyEssential: "Ensures printed panels meet shade and curing quality standards while keeping printing turnaround under SLA.",
      department: "Embroidery & Printing",
      features: ["Screen Allocation", "Color Registration", "Curing Log", "Daily Output"],
      lastUpdated: "Today"
    },
    {
      id: "crit-4",
      step: 4,
      title: "After EMB/PRINT Done Report",
      emoji: "📋",
      path: "/pending-issue-to-stitching",
      section: "critical",
      gradient: "linear-gradient(135deg, #ef4444 0%, #f97316 100%)",
      glowColor: "rgba(239, 68, 68, 0.45)",
      badgeColor: "#dc2626",
      accentBg: "#fef2f2",
      accentBorder: "#fecaca",
      description: "Tracks embellished cut panels returned from embroidery/printing that await issuance to sewing lines.",
      whyEssential: "Eliminates assembly bottlenecks by making sure completed embellishment lots are immediately loaded onto sewing lines.",
      department: "Embroidery & Printing",
      features: ["WIP Handover", "Ready for Stitching", "Aging in Transit", "Batch Transfer"],
      lastUpdated: "Live Feed"
    },
    {
      id: "crit-5",
      step: 5,
      title: "After Cutting Done Report",
      emoji: "⏳",
      path: "/pending-stitching",
      section: "critical",
      gradient: "linear-gradient(135deg, #f59e0b 0%, #d97706 100%)",
      glowColor: "rgba(245, 158, 11, 0.45)",
      badgeColor: "#d97706",
      accentBg: "#fffbeb",
      accentBorder: "#fde68a",
      description: "Monitors cut garment bundles lying in storage awaiting line loading, secondary processes, or stitch allotment.",
      whyEssential: "Prevents high bundle aging and fabric crease damage by flagging cut lots sitting idle for more than 48 hours.",
      department: "Stitching",
      features: ["Bundle Aging", "Line Allocation Queue", "Cut-to-Stitch Lag", "Delayed Lots"],
      lastUpdated: "Hourly"
    },
    {
      id: "crit-6",
      step: 6,
      title: "Overall Stitching Report",
      emoji: "🪡",
      path: "/stitching-complete-lot",
      section: "critical",
      gradient: "linear-gradient(135deg, #0284c7 0%, #2563eb 100%)",
      glowColor: "rgba(2, 132, 199, 0.45)",
      badgeColor: "#0284c7",
      accentBg: "#f0f9ff",
      accentBorder: "#bae6fd",
      description: "Comprehensive live report of finished sewn garments, line productivity, operator piece rates, and defect passes.",
      whyEssential: "Provides real-time line balance visibility, ensuring hourly targets match daily shipment commitments.",
      department: "Stitching",
      features: ["Complete Lots", "Line Productivity", "Efficiency Rate", "Batch Inspection"],
      lastUpdated: "Real-time"
    },
    {
      id: "crit-7",
      step: 7,
      title: "Pending Packing to Issue",
      emoji: "📦⏳",
      path: "/pending-packing-issue",
      section: "critical",
      gradient: "linear-gradient(135deg, #ea580c 0%, #c2410c 100%)",
      glowColor: "rgba(234, 88, 12, 0.45)",
      badgeColor: "#ea580c",
      accentBg: "#fff7ed",
      accentBorder: "#fed7aa",
      description: "Tracks stitched garments waiting for final finishing clearance, thread trimming, and packing department handover.",
      whyEssential: "Stops finished goods from accumulating on the sewing floor and accelerates final iron and carton packing.",
      department: "Packing & Finishing",
      features: ["Stitching Done Queue", "Ready for Iron/Pack", "Floor Clearance", "Batch Transfer"],
      lastUpdated: "Live Feed"
    },
    {
      id: "crit-8",
      step: 8,
      title: "Pending Process after Stitching Done",
      emoji: "⚙️⏳",
      path: "/pending-packing-issue?view=pending_process",
      section: "critical",
      gradient: "linear-gradient(135deg, #b45309 0%, #78350f 100%)",
      glowColor: "rgba(180, 83, 9, 0.45)",
      badgeColor: "#b45309",
      accentBg: "#fffbeb",
      accentBorder: "#fde68a",
      description: "Tracks stitched lots currently running in intermediate finishing processes (Overlock, Kaj Button, Washing, Elastic).",
      whyEssential: "Pinpoints exact intermediate operation bottlenecks before garments reach final polybag packing.",
      department: "Packing & Finishing",
      features: ["Inter-Process WIP", "Overlock / Kaj / Wash", "Process Aging", "Bottleneck Audit"],
      lastUpdated: "Live Feed"
    },
    {
      id: "crit-9",
      step: 9,
      title: "Packing Alloted Report",
      emoji: "📦📋",
      path: "/packing-alloted-lot",
      section: "critical",
      gradient: "linear-gradient(135deg, #4338ca 0%, #312e81 100%)",
      glowColor: "rgba(67, 56, 202, 0.45)",
      badgeColor: "#4338ca",
      accentBg: "#eef2ff",
      accentBorder: "#c7d2fe",
      description: "Manages lots allocated for carton packing, operator bench assignment, polybag barcodes, and dispatch planning.",
      whyEssential: "Ensures packing tables have assigned carton capacities to prevent dispatch day shipping panics.",
      department: "Packing & Finishing",
      features: ["Table Assignment", "Carton Planning", "Barcode Sync", "Operator Allocation"],
      lastUpdated: "Today"
    },
    {
      id: "crit-10",
      step: 10,
      title: "Cut to Pack Report",
      emoji: "📊🏆",
      path: "/overall-cutting-to-packing-report",
      section: "critical",
      gradient: "linear-gradient(135deg, #1e1b4b 0%, #4f46e5 100%)",
      glowColor: "rgba(79, 70, 229, 0.5)",
      badgeColor: "#1e1b4b",
      accentBg: "#e0e7ff",
      accentBorder: "#a5b4fc",
      description: "The ultimate factory audit matrix comparing cutting plies vs packed carton pieces to calculate factory net yield.",
      whyEssential: "Directly detects production losses, rejects, and short shipments across the entire factory manufacturing pipeline.",
      department: "Executive & Analytics",
      features: ["Cutting vs Packed Pcs", "Factory Yield %", "Loss & Rejection Audit", "Style-Wise P&L"],
      lastUpdated: "Real-time"
    }
  ], []);

  // 2. SECTION: Individual Department Reports (8 Reports)
  const departmentReports = useMemo(() => [
    {
      id: "dept-1",
      title: "Daily Kaj Button Report",
      emoji: "🔘",
      path: "/daily-kaj-button-report",
      section: "departments",
      gradient: "linear-gradient(135deg, #0284c7 0%, #06b6d4 100%)",
      glowColor: "rgba(2, 132, 199, 0.45)",
      badgeColor: "#0284c7",
      accentBg: "#f0f9ff",
      accentBorder: "#bae6fd",
      description: "Daily button attaching, buttonhole (kaj) punching output, machine rate logs, and thread color matches.",
      whyEssential: "Maintains uninterrupted buttoning flow on shirts and polo garments before garments move to folding.",
      department: "Stitching Operations",
      features: ["Button Attachment Log", "Kaj Hole Quality", "Operator Piece Rate", "Shift Output"],
      lastUpdated: "Daily"
    },
    {
      id: "dept-2",
      title: "Daily Overlock Report",
      emoji: "🪡⚡",
      path: "/daily-overlock-report",
      section: "departments",
      gradient: "linear-gradient(135deg, #0d9488 0%, #14b8a6 100%)",
      glowColor: "rgba(13, 148, 136, 0.45)",
      badgeColor: "#0d9488",
      accentBg: "#ccfbf1",
      accentBorder: "#99f6e4",
      description: "Operator-wise overlock stitch logs, multi-thread seam tension checks, and edge finish quality metrics.",
      whyEssential: "Audits seam strength and high-speed overlock output across knit t-shirts and fleece joggers.",
      department: "Stitching Operations",
      features: ["Overlock Seam Log", "Operator Target", "Tension Inspection", "Daily Output"],
      lastUpdated: "Daily"
    },
    {
      id: "dept-3",
      title: "Daily Folding Report",
      emoji: "👔📦",
      path: "/daily-folding-report",
      section: "departments",
      gradient: "linear-gradient(135deg, #0891b2 0%, #06b6d4 100%)",
      glowColor: "rgba(6, 182, 212, 0.45)",
      badgeColor: "#0891b2",
      accentBg: "#ecfeff",
      accentBorder: "#a5f3fc",
      description: "Garment steam pressing, folding, tag insertion, and polybag packaging production records linked live to the Folding department sheet.",
      whyEssential: "Ensures finishing tables maintain daily folding pace to support packing department carton loading and prevent dispatch delays.",
      department: "Finishing Operations",
      features: ["Live Sheet Sync", "Folding Output", "Supervisor Log", "Daily Tally"],
      lastUpdated: "Daily"
    },
    {
      id: "dept-4",
      title: "Elastic Report",
      emoji: "🎗️",
      path: "/elastic-report",
      section: "departments",
      gradient: "linear-gradient(135deg, #f59e0b 0%, #fbbf24 100%)",
      glowColor: "rgba(245, 158, 11, 0.45)",
      badgeColor: "#d97706",
      accentBg: "#fffbeb",
      accentBorder: "#fde68a",
      description: "Waistband elastic attachment tracking, elastic tension testing, roll consumption, and stitch elongation logs.",
      whyEssential: "Ensures waistband stretch recovery meets brand compliance and prevents elastic puckering rejections.",
      department: "Stitching Operations",
      features: ["Waistband Assembly", "Elastic Roll Usage", "Tension Test", "Defect Checks"],
      lastUpdated: "Daily"
    },
    {
      id: "dept-5",
      title: "Feed Up Report",
      emoji: "⚙️🧵",
      path: "/feed-up-report",
      section: "departments",
      gradient: "linear-gradient(135deg, #6366f1 0%, #818cf8 100%)",
      glowColor: "rgba(99, 102, 241, 0.45)",
      badgeColor: "#6366f1",
      accentBg: "#eef2ff",
      accentBorder: "#c7d2fe",
      description: "Feed-off-the-arm (Feed-Up) machine lap-seam logs, inseam joint audits, and specialized heavy-stitch production.",
      whyEssential: "Monitors specialized structural seams on track pants, hoodies, and jackets to prevent seam unraveling.",
      department: "Stitching Operations",
      features: ["Lap-Seam Production", "Feed-Up Machine Log", "Operator Rate", "Inseam Quality"],
      lastUpdated: "Daily"
    },
    {
      id: "dept-6",
      title: "Washing Report",
      emoji: "🧼",
      path: "/washing-report",
      section: "departments",
      gradient: "linear-gradient(135deg, #06b6d4 0%, #38bdf8 100%)",
      glowColor: "rgba(6, 182, 212, 0.45)",
      badgeColor: "#0284c7",
      accentBg: "#ecfeff",
      accentBorder: "#a5f3fc",
      description: "Garment washing cycles, bio-polishing, enzyme softening, shrinkage measurement, and shade consistency tests.",
      whyEssential: "Guarantees exact post-wash dimensional stability and premium soft handfeel before final pressing.",
      department: "Finishing Operations",
      features: ["Softening Cycles", "Shrinkage Testing", "Shade Lot Matching", "Moisture Control"],
      lastUpdated: "Daily"
    },
    {
      id: "dept-7",
      title: "Jaybir Embroidery Report",
      emoji: "🧵✨",
      path: "/jaybir-embroidery-report",
      section: "departments",
      gradient: "linear-gradient(135deg, #8b5cf6 0%, #c084fc 100%)",
      glowColor: "rgba(139, 92, 246, 0.45)",
      badgeColor: "#7c3aed",
      accentBg: "#f5f3ff",
      accentBorder: "#ddd6fe",
      description: "Dedicated production challan tracker and SLA aging monitor linked live to the Jaybir Embroidery unit.",
      whyEssential: "Enables instant vendor SLA accountability by tracking pending pieces, aging days, and verified completion dates.",
      department: "Vendor Partner",
      features: ["Live Sheet Sync", "Pending Pieces", "SLA Aging Days", "Done Verification"],
      lastUpdated: "Live Google Sheet"
    },
    {
      id: "dept-8",
      title: "Jaybir Printing Report",
      emoji: "🎨✨",
      path: "/jaybir-printing-report",
      section: "departments",
      gradient: "linear-gradient(135deg, #ec4899 0%, #f472b6 100%)",
      glowColor: "rgba(236, 72, 153, 0.45)",
      badgeColor: "#db2777",
      accentBg: "#fdf2f8",
      accentBorder: "#fbcfe8",
      description: "Dedicated production challan tracker and SLA aging monitor linked live to the Jaybir Screen Printing unit.",
      whyEssential: "Tracks outside printing turnaround times, pending pieces per lot, and verifies completed vendor challans.",
      department: "Vendor Partner",
      features: ["Live Sheet Sync", "Pending Pieces", "SLA Aging Days", "Done Verification"],
      lastUpdated: "Live Google Sheet"
    }
  ], []);

  // 3. SECTION: Special Reports for Admin (5 Reports)
  const adminReports = useMemo(() => [
    {
      id: "admin-1",
      title: "Knitting Report",
      emoji: "🧶",
      path: "/knitting-report",
      section: "admin",
      gradient: "linear-gradient(135deg, #b45309 0%, #d97706 100%)",
      glowColor: "rgba(180, 83, 9, 0.45)",
      badgeColor: "#b45309",
      accentBg: "#fffbeb",
      accentBorder: "#fde68a",
      description: "Greige fabric circular knitting machine output, yarn lot consumption, loop density, and daily knitted kg rolls.",
      whyEssential: "Tracks raw greige fabric production to ensure sufficient fabric supply for scheduled cutting orders.",
      department: "Material & Supply Chain",
      adminOnly: true,
      features: ["Knitting Machine Output", "GSM Control", "Yarn Consumption", "Roll Weight Audit"],
      lastUpdated: "Daily"
    },
    {
      id: "admin-2",
      title: "Collar Report",
      emoji: "👕",
      path: "/collar-report",
      section: "admin",
      gradient: "linear-gradient(135deg, #7c3aed 0%, #6366f1 100%)",
      glowColor: "rgba(124, 58, 237, 0.45)",
      badgeColor: "#7c3aed",
      accentBg: "#f5f3ff",
      accentBorder: "#ddd6fe",
      description: "Flat knit collar and rib cuff manufacturing output, tipping design matching, yarn count, and lot allocation.",
      whyEssential: "Prevents polo shirt assembly halts by synchronizing collar/cuff quantities with cutting job orders.",
      department: "Material & Supply Chain",
      adminOnly: true,
      features: ["Flat Knit Output", "Collar Matching", "Tipping Design", "Lot Balance"],
      lastUpdated: "Daily"
    },
    {
      id: "admin-3",
      title: "Yarn Report & Stock",
      emoji: "🧵⚖️",
      path: "/yarn-report",
      section: "admin",
      gradient: "linear-gradient(135deg, #059669 0%, #047857 100%)",
      glowColor: "rgba(5, 150, 105, 0.45)",
      badgeColor: "#059669",
      accentBg: "#ecfdf5",
      accentBorder: "#a7f3d0",
      description: "Master spun yarn inventory, mill-wise inward receipts, lot bag stock balances, and yarn weight allocations.",
      whyEssential: "Directly protects factory capital by maintaining tight controls over high-value yarn warehouse inventory.",
      department: "Material & Supply Chain",
      adminOnly: true,
      features: ["Yarn Bag Stock", "Mill Inward Ledger", "Lot Balance Weight", "Count Analysis"],
      lastUpdated: "Daily"
    },
    {
      id: "admin-4",
      title: "Kora Roll Report",
      emoji: "📜",
      path: "/kora-roll-report",
      section: "admin",
      gradient: "linear-gradient(135deg, #ea580c 0%, #c2410c 100%)",
      glowColor: "rgba(234, 88, 12, 0.45)",
      badgeColor: "#ea580c",
      accentBg: "#fff7ed",
      accentBorder: "#fed7aa",
      description: "Raw un-dyed greige (Kora) roll storage, lot dispatch to process houses, and shrinkage/loss records.",
      whyEssential: "Controls greige roll allotment to dye mills and ensures zero unauthorized roll leakage.",
      department: "Material & Supply Chain",
      adminOnly: true,
      features: ["Greige Roll Stock", "Dyeing Allotment", "Shrinkage Audit", "Weight Verification"],
      lastUpdated: "Daily"
    },
    {
      id: "admin-5",
      title: "Pending Zip PO Report",
      emoji: "🤐📦",
      path: "/pending-zip-po-report",
      section: "admin",
      gradient: "linear-gradient(135deg, #dc2626 0%, #991b1b 100%)",
      glowColor: "rgba(220, 38, 38, 0.45)",
      badgeColor: "#dc2626",
      accentBg: "#fef2f2",
      accentBorder: "#fecaca",
      description: "Zipper procurement purchase orders, supplier dispatch tracking, style-wise zipper lengths, and PO balances.",
      whyEssential: "Eliminates high-risk zipper shortages on jacket and sweatshirt lines before sewing operations stall.",
      department: "Procurement & Trims",
      features: ["Zip PO Ledger", "Supplier Dispatch SLA", "Length & Teeth Spec", "PO Balance"],
      lastUpdated: "Live"
    }
  ], []);

  // 4. SECTION: Other Factory Tools & Auxiliary Reports
  const otherReports = useMemo(() => [
    {
      id: "other-1",
      title: "Stitching Issue",
      emoji: "🧶",
      path: "/stitching",
      section: "other",
      gradient: "linear-gradient(135deg, #06b6d4 0%, #3b82f6 100%)",
      glowColor: "rgba(6, 182, 212, 0.45)",
      badgeColor: "#0284c7",
      accentBg: "#ecfeff",
      accentBorder: "#a5f3fc",
      description: "Manage daily stitching line assignments, operator tracking, hourly output targets, and sewing floor allotments.",
      whyEssential: "Core operational module for issuing cut bundles directly to stitching line leaders.",
      department: "Stitching Operations",
      features: ["Line Wise", "Operator Tracking", "Hourly Target", "Bundle Issue"],
      lastUpdated: "Today"
    },
    {
      id: "other-2",
      title: "Issue To Packing",
      emoji: "📦",
      path: "/issue-to-packing",
      section: "other",
      gradient: "linear-gradient(135deg, #10b981 0%, #0d9488 100%)",
      glowColor: "rgba(16, 185, 129, 0.45)",
      badgeColor: "#059669",
      accentBg: "#ecfdf5",
      accentBorder: "#a7f3d0",
      description: "Manage packing operations, carton preparation, size tagging, and shipment readiness.",
      whyEssential: "Official floor handover tool moving sewn garments to finishing tables.",
      department: "Packing & Finishing",
      features: ["Carton Tracking", "Quality Check", "Shipment Prep", "Batch Transfer"],
      lastUpdated: "Today"
    },
    {
      id: "other-3",
      title: "Daily Embroidery Challan",
      emoji: "📈",
      path: "/daily-embroidery-challan",
      section: "other",
      gradient: "linear-gradient(135deg, #00c6ff 0%, #0072ff 100%)",
      glowColor: "rgba(0, 198, 255, 0.45)",
      badgeColor: "#0284c7",
      accentBg: "#f0f9ff",
      accentBorder: "#bae6fd",
      description: "Daily embroidery machine outputs, stitch counts, thread consumption, and operator logs.",
      whyEssential: "Monitors shift-by-shift embroidery machine run time and productivity.",
      department: "Embroidery & Printing",
      features: ["Machine Wise", "Stitch Count", "Thread Usage", "Operator Log"],
      lastUpdated: "Daily"
    },
    {
      id: "other-4",
      title: "Daily Printing Challan",
      emoji: "📊",
      path: "/daily-printing-challan",
      section: "other",
      gradient: "linear-gradient(135deg, #ff7e5f 0%, #feb47b 100%)",
      glowColor: "rgba(255, 126, 95, 0.45)",
      badgeColor: "#ea580c",
      accentBg: "#fff7ed",
      accentBorder: "#fed7aa",
      description: "Daily printing screen usage, color runs, operator logs, and panel outputs.",
      whyEssential: "Tracks daily screen utilization and printing color formulations.",
      department: "Embroidery & Printing",
      features: ["Screen Utilization", "Color Output", "Operator Wise", "Daily Tally"],
      lastUpdated: "Daily"
    },
    {
      id: "other-5",
      title: "Short Summary Report",
      emoji: "📑⚡",
      path: "/short-summary-report",
      section: "other",
      gradient: "linear-gradient(135deg, #4f46e5 0%, #3730a3 100%)",
      glowColor: "rgba(79, 70, 229, 0.45)",
      badgeColor: "#4f46e5",
      accentBg: "#eef2ff",
      accentBorder: "#c7d2fe",
      description: "Executive high-level snapshot summarizing active factory lots across all stages into a concise single sheet.",
      whyEssential: "Gives general managers an instantaneous overview of entire factory WIP in under 60 seconds.",
      department: "Executive & Analytics",
      adminOnly: true,
      features: ["Executive Snapshot", "Stage Counts", "Quick Export", "Lot Status"],
      lastUpdated: "Real-time"
    },
    {
      id: "other-6",
      title: "Daily Fabric Issue Analytics",
      emoji: "🧵📊",
      path: "/daily-fabric-issue-report",
      section: "other",
      gradient: "linear-gradient(135deg, #0284c7 0%, #0369a1 100%)",
      glowColor: "rgba(2, 132, 199, 0.45)",
      badgeColor: "#0284c7",
      accentBg: "#f0f9ff",
      accentBorder: "#bae6fd",
      description: "Fabric roll issuance matrix mapping total volume, roll weights, and table assignments across styles.",
      whyEssential: "Ensures cutting masters receive the exact planned fabric quantities without over-issuing rolls.",
      department: "Cutting Department",
      features: ["Roll Weight Matrix", "Table Allocation", "Fabric Audit", "Shift Issue Log"],
      lastUpdated: "Daily"
    },
    {
      id: "other-7",
      title: "All Cutting Job Orders",
      emoji: "📝",
      path: "/all-cutting-joborders",
      section: "other",
      gradient: "linear-gradient(135deg, #059669 0%, #047857 100%)",
      glowColor: "rgba(5, 150, 105, 0.45)",
      badgeColor: "#059669",
      accentBg: "#ecfdf5",
      accentBorder: "#a7f3d0",
      description: "Master digital registry of all active and historical cutting job order cards with complete ratio tables.",
      whyEssential: "Centralizes job order specifications so cutting masters always work from approved size ratios.",
      department: "Cutting Department",
      features: ["Job Order Registry", "Size Ratio Table", "Marker Notes", "Order Specs"],
      lastUpdated: "Real-time"
    },
    {
      id: "other-8",
      title: "Fabric Roll Prediction (AI)",
      emoji: "🤖📏",
      path: "/fabric-roll-prediction",
      section: "other",
      gradient: "linear-gradient(135deg, #8b5cf6 0%, #6d28d9 100%)",
      glowColor: "rgba(139, 92, 246, 0.45)",
      badgeColor: "#7c3aed",
      accentBg: "#f5f3ff",
      accentBorder: "#ddd6fe",
      description: "AI-driven fabric roll combination forecaster that computes optimal cutting lay lengths to minimize end-bits.",
      whyEssential: "Reduces fabric end-bit scrap by 3% to 5% through intelligent mathematical roll grouping.",
      department: "Cutting Department",
      features: ["AI Roll Optimization", "End-Bit Minimizer", "Lay Plan Generator", "Yarn Savings"],
      lastUpdated: "AI Powered"
    },
    {
      id: "other-9",
      title: "Sticker & Barcode Report",
      emoji: "🏷️",
      path: "/sticker-report",
      section: "other",
      gradient: "linear-gradient(135deg, #f59e0b 0%, #d97706 100%)",
      glowColor: "rgba(245, 158, 11, 0.45)",
      badgeColor: "#d97706",
      accentBg: "#fffbeb",
      accentBorder: "#fde68a",
      description: "Bundle sticker numbering, barcode batch printing, ply sequence tags, and bundle count auditing.",
      whyEssential: "Guarantees bundle part traceability across all downstream sewing and finishing stages.",
      department: "Cutting Department",
      features: ["Bundle Barcodes", "Ply Sequence", "Batch Printing", "Part Verification"],
      lastUpdated: "Real-time"
    },
    {
      id: "other-10",
      title: "Daily Packing Report",
      emoji: "📦✨",
      path: "/daily-packing-report",
      section: "other",
      gradient: "linear-gradient(135deg, #10b981 0%, #059669 100%)",
      glowColor: "rgba(16, 185, 129, 0.45)",
      badgeColor: "#059669",
      accentBg: "#ecfdf5",
      accentBorder: "#a7f3d0",
      description: "Daily carton boxing records, size assortment verification, barcode scanning, and dispatch readiness.",
      whyEssential: "Validates that boxed assortments exactly match brand purchase order packing lists.",
      department: "Packing & Finishing",
      features: ["Carton Assortment", "Packing Tally", "Barcode Scan Log", "Dispatch Ready"],
      lastUpdated: "Daily"
    },
    {
      id: "other-11",
      title: "Packing Complete Lots",
      emoji: "📦✅",
      path: "/packing-completed-lots",
      section: "other",
      gradient: "linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%)",
      glowColor: "rgba(37, 99, 235, 0.45)",
      badgeColor: "#1e3a8a",
      accentBg: "#eff6ff",
      accentBorder: "#bfdbfe",
      description: "Archive of 100% completed and boxed lots with turnaround time metrics, completion dates, and export ledgers.",
      whyEssential: "Maintains official historical records of all completed lots for buyer shipping reconciliation.",
      department: "Packing & Finishing",
      features: ["Completion Date Audit", "Aging Analysis", "Barcode Linkage", "Excel/PDF Export"],
      lastUpdated: "Real-time"
    },
    {
      id: "other-13",
      title: "Hold Lot Manager",
      emoji: "⏸️",
      path: "/hold-lot",
      section: "other",
      gradient: "linear-gradient(135deg, #e11d48 0%, #be123c 100%)",
      glowColor: "rgba(225, 29, 72, 0.45)",
      badgeColor: "#e11d48",
      accentBg: "#fff1f2",
      accentBorder: "#fecdd3",
      description: "Action console to pause lots on the factory floor, select hold reasons, and notify supervisors.",
      whyEssential: "Instantly locks suspect lots from being processed further.",
      department: "Quality & Compliance",
      features: ["Hold Action", "Reason Select", "Floor Lock", "Audit Trail"],
      lastUpdated: "Real-time"
    },
    {
      id: "other-14",
      title: "Hold Lots Report",
      emoji: "⏸️📋",
      path: "/hold-lots-report",
      section: "other",
      gradient: "linear-gradient(135deg, #9f1239 0%, #881337 100%)",
      glowColor: "rgba(159, 18, 57, 0.45)",
      badgeColor: "#9f1239",
      accentBg: "#fff1f2",
      accentBorder: "#fecdd3",
      description: "Quarantine manager report for lots placed on hold with aging breakdown and release tracking.",
      whyEssential: "Ensures quarantined lots are resolved quickly without being forgotten.",
      department: "Quality & Compliance",
      features: ["Quarantine Log", "Hold Reason", "Release Approval", "Aging Days"],
      lastUpdated: "Live"
    },
    {
      id: "other-15",
      title: "Cancelled Lots Report",
      emoji: "🚫❌",
      path: "/cancelled-lots-report",
      section: "other",
      gradient: "linear-gradient(135deg, #64748b 0%, #334155 100%)",
      glowColor: "rgba(100, 116, 139, 0.45)",
      badgeColor: "#475569",
      accentBg: "#f8fafc",
      accentBorder: "#cbd5e1",
      description: "Official cancellation log for cancelled job orders, recording management authorization stamps and fabric salvage notes.",
      whyEssential: "Maintains financial accountability for cancelled orders and prevents accidental continuation of cancelled runs.",
      department: "Executive & Analytics",
      features: ["Cancellation Reason", "Authorization Stamp", "Fabric Salvage", "Brand Audit"],
      lastUpdated: "Real-time"
    },
    {
      id: "other-16",
      title: "Lot Change Logs",
      emoji: "📋🔍",
      path: "/lot-logs",
      section: "other",
      gradient: "linear-gradient(135deg, #4f46e5 0%, #312e81 100%)",
      glowColor: "rgba(79, 70, 229, 0.45)",
      badgeColor: "#4f46e5",
      accentBg: "#eef2ff",
      accentBorder: "#c7d2fe",
      description: "Tamper-evident audit trail tracking every edit made to lot piece quantities, style codes, or ratio breakdowns.",
      whyEssential: "Ensures complete accountability across production staff by logging who modified what lot and when.",
      department: "System Administration",
      features: ["Permission Log", "Lot Search", "Modification History", "Live Broadcast"],
      lastUpdated: "Real-time"
    },
    {
      id: "other-17",
      title: "Lot Lifecycle Timeline (Tracker)",
      emoji: "🚚⏱️",
      path: "/lot-timeline",
      section: "other",
      gradient: "linear-gradient(135deg, #f59e0b 0%, #ea580c 50%, #dc2626 100%)",
      glowColor: "rgba(245, 158, 11, 0.45)",
      badgeColor: "#d97706",
      accentBg: "#fffbeb",
      accentBorder: "#fde68a",
      description: "Amazon Order Tracking-style visual pipeline showing stage-by-stage milestone completion and aging for any lot number.",
      whyEssential: "Empowers customer service and merchandising teams to answer buyer progress inquiries in seconds.",
      department: "Executive & Analytics",
      features: ["Amazon Stepper UI", "Live Milestone Aging", "Lot Number Search", "Stage History"],
      lastUpdated: "Real-time"
    },
    {
      id: "other-18",
      title: "Production Flowchart Poster",
      emoji: "🗺️📊",
      path: "/production-flowchart",
      section: "other",
      gradient: "linear-gradient(135deg, #0284c7 0%, #4f46e5 100%)",
      glowColor: "rgba(2, 132, 199, 0.45)",
      badgeColor: "#0284c7",
      accentBg: "#f0f9ff",
      accentBorder: "#bae6fd",
      description: "Visual high-resolution flowchart mapping complete factory sequence from raw yarn inward to finished carton dispatch.",
      whyEssential: "Acts as a visual standard operating procedure (SOP) for floor supervisors and buyer factory audits.",
      department: "Quality & Compliance",
      features: ["Visual Flowchart", "SOP Mapping", "Stage Connections", "Printable Poster"],
      lastUpdated: "SOP Reference"
    },
    {
      id: "other-19",
      title: "Notification Center",
      emoji: "🔔",
      path: "/notifications",
      section: "other",
      gradient: "linear-gradient(135deg, #ec4899 0%, #be185d 100%)",
      glowColor: "rgba(236, 72, 153, 0.45)",
      badgeColor: "#db2777",
      accentBg: "#fdf2f8",
      accentBorder: "#fbcfe8",
      description: "Centralized broadcast center for floor-wide delay alerts, management notices, and critical lot priority alarms.",
      whyEssential: "Instantly communicates urgent production notices across all connected supervisor terminals.",
      department: "Executive & Analytics",
      features: ["Live Broadcasts", "Priority Alarms", "Department Routing", "Read Receipts"],
      lastUpdated: "Real-time"
    },
    {
      id: "other-20",
      title: "EMB / Print Remarks Manager",
      emoji: "📝💬",
      path: "/emb-print-remarks",
      section: "other",
      gradient: "linear-gradient(135deg, #8b5cf6 0%, #d946ef 100%)",
      glowColor: "rgba(139, 92, 246, 0.45)",
      badgeColor: "#7c3aed",
      accentBg: "#f5f3ff",
      accentBorder: "#ddd6fe",
      description: "Quick remarks entry console allowing embroidery and printing floor supervisors to log specific delay reasons.",
      whyEssential: "Captures on-the-ground operational notes (thread shortage, screen repair) directly into centralized sheets.",
      department: "Embroidery & Printing",
      features: ["Remarks Entry", "Delay Tagging", "Google Sheets Sync", "Lot Search"],
      lastUpdated: "Real-time"
    },
    {
      id: "other-21",
      title: "Embroidery Material Receiving",
      emoji: "📥🧵",
      path: "/embroidery-material-receiving",
      section: "other",
      gradient: "linear-gradient(135deg, #10b981 0%, #047857 100%)",
      glowColor: "rgba(16, 185, 129, 0.45)",
      badgeColor: "#059669",
      accentBg: "#ecfdf5",
      accentBorder: "#a7f3d0",
      description: "Receiving gate inspection for thread cones, embroidery backing paper, and embellished cut panel returns.",
      whyEssential: "Stops defective or short-shipped embroidery accessories at the factory gate before production begins.",
      department: "Materials & Inventory",
      features: ["Gate Inward Check", "Thread Receiving", "Defect Logging", "Receipt Approval"],
      lastUpdated: "Daily"
    },
    {
      id: "other-22",
      title: "Printing Material Receiving",
      emoji: "📥🖨️",
      path: "/printing-material-receiving",
      section: "other",
      gradient: "linear-gradient(135deg, #06b6d4 0%, #0284c7 100%)",
      glowColor: "rgba(6, 182, 212, 0.45)",
      badgeColor: "#0891b2",
      accentBg: "#ecfeff",
      accentBorder: "#a5f3fc",
      description: "Receiving gate inspection for printing screens, plastisol/waterbase inks, and printed cut panel returns.",
      whyEssential: "Stops incorrect chemical or ink batches before screens are mounted.",
      department: "Materials & Inventory",
      features: ["Gate Inward Check", "Ink Receiving", "Screen Log", "Receipt Approval"],
      lastUpdated: "Daily"
    },
    {
      id: "other-23",
      title: "Zip Purchase Dashboard",
      emoji: "🤐",
      path: "/zip-report",
      section: "other",
      gradient: "linear-gradient(135deg, #f97316 0%, #ea580c 100%)",
      glowColor: "rgba(249, 115, 22, 0.45)",
      badgeColor: "#ea580c",
      accentBg: "#fff7ed",
      accentBorder: "#fed7aa",
      description: "Zipper stock inventory, style-wise zipper consumption, and supplier procurement metrics.",
      whyEssential: "Tracks zipper inventory balances for sweatshirts and winter jackets.",
      department: "Materials & Inventory",
      features: ["Zip Stock", "Style Match", "Supplier Log", "Balance Ledger"],
      lastUpdated: "Live"
    },
    {
      id: "other-24",
      title: "Dori Purchase Dashboard",
      emoji: "🪢",
      path: "/dori-report",
      section: "other",
      gradient: "linear-gradient(135deg, #84cc16 0%, #65a30d 100%)",
      glowColor: "rgba(132, 204, 22, 0.45)",
      badgeColor: "#65a30d",
      accentBg: "#f7fee7",
      accentBorder: "#d9f99d",
      description: "Drawstring (dori) and cord inventory, aglet tip specifications, and hoodie trims tracking.",
      whyEssential: "Manages drawstring inventory to prevent delays on hoodie and jogger assembly lines.",
      department: "Materials & Inventory",
      features: ["Dori Inventory", "Aglet Specs", "Color Matching", "Consumption Log"],
      lastUpdated: "Live"
    }
  ], []);

  // Filter accessible modules based on admin status
  const filterByRole = (list) => {
    return list.filter((m) => {
      if (m.adminOnly && !isAdmin) return false;
      return true;
    });
  };

  const filteredCritical = useMemo(() => filterByRole(criticalReports), [criticalReports, isAdmin]);
  const filteredDepartments = useMemo(() => filterByRole(departmentReports), [departmentReports, isAdmin]);
  const filteredAdmin = useMemo(() => filterByRole(adminReports), [adminReports, isAdmin]);
  const filteredOther = useMemo(() => filterByRole(otherReports), [otherReports, isAdmin]);

  const allAccessibleModules = useMemo(() => [
    ...filteredCritical,
    ...filteredDepartments,
    ...filteredAdmin,
    ...filteredOther
  ], [filteredCritical, filteredDepartments, filteredAdmin, filteredOther]);

  // Global search filtering
  const searchResults = useMemo(() => {
    if (!searchQuery.trim()) return null;
    const term = searchQuery.toLowerCase().trim();
    return allAccessibleModules.filter((m) =>
      m.title.toLowerCase().includes(term) ||
      m.description.toLowerCase().includes(term) ||
      m.whyEssential.toLowerCase().includes(term) ||
      m.department.toLowerCase().includes(term) ||
      m.features.some((f) => f.toLowerCase().includes(term))
    );
  }, [allAccessibleModules, searchQuery]);

  return (
    <div style={{ minHeight: "100vh", width: "100%", background: "#f8fafc", color: "#0f172a", fontFamily: "'Inter', system-ui, -apple-system, sans-serif", overflowX: "hidden" }}>
      {/* ============================================================================ */}
      {/* 3D CSS STYLE INJECTIONS FOR 3D PERSPECTIVE & VIBRANT PALETTES                */}
      {/* ============================================================================ */}
      <style>{`
        .full-dashboard-wrapper {
          width: 100%;
          box-sizing: border-box;
        }
        .perspective-card-container {
          perspective: 1200px;
          perspective-origin: center;
        }
        .card-3d-wrapper {
          transform-style: preserve-3d;
          transition: transform 0.35s cubic-bezier(0.2, 0.8, 0.2, 1), box-shadow 0.35s ease, border-color 0.25s ease;
        }
        .card-3d-wrapper:hover {
          transform: translateY(-8px) rotateX(4deg) rotateY(-2deg) scale(1.015);
        }
        .card-3d-floating-icon {
          transform: translateZ(28px);
          transition: transform 0.35s cubic-bezier(0.2, 0.8, 0.2, 1);
        }
        .card-3d-wrapper:hover .card-3d-floating-icon {
          transform: translateZ(42px) scale(1.12) rotate(4deg);
        }
        .card-3d-floating-badge {
          transform: translateZ(20px);
        }
        .card-3d-floating-btn {
          transform: translateZ(22px);
          transition: all 0.25s ease;
        }
        .card-3d-wrapper:hover .card-3d-floating-btn {
          transform: translateZ(32px) scale(1.02);
        }
      `}</style>

      {/* ============================================================================ */}
      {/* TOP HEADER: VIBRANT COLORFUL NAVBAR (FULL WIDTH)                              */}
      {/* ============================================================================ */}
      <header
        style={{
          width: "100%",
          background: "#ffffff",
          borderBottom: "1.5px solid #e2e8f0",
          boxShadow: "0 4px 25px -4px rgba(0, 0, 0, 0.05)",
          position: "sticky",
          top: 0,
          zIndex: 40,
          boxSizing: "border-box"
        }}
      >
        <div style={{ width: "100%", padding: "16px 36px", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: "16px", boxSizing: "border-box" }}>
          {/* Brand & Factory Status */}
          <div style={{ display: "flex", alignItems: "center", gap: "16px" }}>
            <div
              style={{
                width: "48px",
                height: "48px",
                borderRadius: "14px",
                background: "linear-gradient(135deg, #4f46e5 0%, #ec4899 50%, #f59e0b 100%)",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                color: "#ffffff",
                fontSize: "1.7rem",
                boxShadow: "0 8px 20px rgba(236, 72, 153, 0.35)"
              }}
            >
              🏭
            </div>
            <div>
              <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                <h1 style={{ margin: 0, fontSize: "1.45rem", fontWeight: 900, color: "#0f172a", letterSpacing: "-0.02em" }}>
                  Factory Suite <span style={{ background: "linear-gradient(135deg, #4f46e5 0%, #ec4899 100%)", WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent" }}>Pro</span>
                </h1>
                <span
                  style={{
                    background: "#dcfce7",
                    color: "#15803d",
                    border: "1.5px solid #86efac",
                    padding: "3px 10px",
                    borderRadius: "9999px",
                    fontSize: "0.72rem",
                    fontWeight: 800,
                    display: "inline-flex",
                    alignItems: "center",
                    gap: "5px"
                  }}
                >
                  <span style={{ width: "7px", height: "7px", borderRadius: "50%", background: "#22c55e", display: "inline-block" }}></span>
                  LIVE COLORFUL COMMAND CENTER
                </span>
              </div>
              <p style={{ margin: "3px 0 0 0", fontSize: "0.84rem", color: "#64748b", fontWeight: 500 }}>
                Enterprise Production Pipeline & Departmental Audit System (Vibrant Multi-Color View)
              </p>
            </div>
          </div>

          {/* User Profile & Quick Actions */}
          <div style={{ display: "flex", alignItems: "center", gap: "14px", flexWrap: "wrap" }}>
            <div
              style={{
                background: "#f8fafc",
                border: "1.5px solid #e2e8f0",
                padding: "6px 14px",
                borderRadius: "12px",
                display: "flex",
                alignItems: "center",
                gap: "10px"
              }}
            >
              <div
                style={{
                  width: "34px",
                  height: "34px",
                  borderRadius: "50%",
                  background: isAdmin ? "linear-gradient(135deg, #d97706 0%, #f59e0b 100%)" : "linear-gradient(135deg, #4f46e5 0%, #06b6d4 100%)",
                  color: "#ffffff",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  fontSize: "0.9rem",
                  fontWeight: 900,
                  boxShadow: "0 4px 10px rgba(0,0,0,0.12)"
                }}
              >
                {(currentUser?.name || "U")[0].toUpperCase()}
              </div>
              <div>
                <div style={{ fontSize: "0.85rem", fontWeight: 800, color: "#0f172a" }}>
                  {currentUser?.name || "Authorized User"}
                </div>
                <div style={{ fontSize: "0.72rem", fontWeight: 800, color: isAdmin ? "#d97706" : "#4f46e5", textTransform: "uppercase" }}>
                  {isAdmin ? "👑 Administrator" : "👷 Production Operator"}
                </div>
              </div>
            </div>

            <button
              onClick={handleLogout}
              style={{
                background: "#fee2e2",
                color: "#dc2626",
                border: "1.5px solid #fca5a5",
                padding: "9px 18px",
                borderRadius: "12px",
                fontSize: "0.84rem",
                fontWeight: 800,
                cursor: "pointer",
                display: "inline-flex",
                alignItems: "center",
                gap: "6px",
                boxShadow: "0 2px 8px rgba(220, 38, 38, 0.12)",
                transition: "all 0.2s"
              }}
              onMouseEnter={(e) => {
                e.currentTarget.style.background = "#fecaca";
                e.currentTarget.style.borderColor = "#f87171";
              }}
              onMouseLeave={(e) => {
                e.currentTarget.style.background = "#fee2e2";
                e.currentTarget.style.borderColor = "#fca5a5";
              }}
            >
              <span>🚪</span> Logout
            </button>
          </div>
        </div>

        {/* 4-Section Colorful KPI Summary Strip (Full Width) */}
        <div style={{ width: "100%", background: "#f8fafc", borderTop: "1px solid #e2e8f0", padding: "12px 36px", boxSizing: "border-box" }}>
          <div style={{ width: "100%", display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))", gap: "14px" }}>
            <HeroSectionStat
              title="1. Critical Daily Reports"
              count={filteredCritical.length}
              icon="🚨"
              color="#dc2626"
              bg="#fef2f2"
              border="#fca5a5"
              active={activeSectionTab === "critical"}
              onClick={() => setActiveSectionTab(activeSectionTab === "critical" ? "all" : "critical")}
            />
            <HeroSectionStat
              title="2. Departmental Operations"
              count={filteredDepartments.length}
              icon="🏢"
              color="#0284c7"
              bg="#f0f9ff"
              border="#bae6fd"
              active={activeSectionTab === "departments"}
              onClick={() => setActiveSectionTab(activeSectionTab === "departments" ? "all" : "departments")}
            />
            <HeroSectionStat
              title="3. Admin Special Reports"
              count={filteredAdmin.length}
              icon="👑"
              color="#d97706"
              bg="#fffbeb"
              border="#fde68a"
              active={activeSectionTab === "admin"}
              onClick={() => setActiveSectionTab(activeSectionTab === "admin" ? "all" : "admin")}
            />
            <HeroSectionStat
              title="4. Other Factory Tools"
              count={filteredOther.length}
              icon="🛠️"
              color="#7c3aed"
              bg="#f5f3ff"
              border="#ddd6fe"
              active={activeSectionTab === "other"}
              onClick={() => setActiveSectionTab(activeSectionTab === "other" ? "all" : "other")}
            />
          </div>
        </div>
      </header>

      {/* ============================================================================ */}
      {/* INTERACTIVE CONTROLS: COLORFUL SEARCH BAR & FILTER PILLS                     */}
      {/* ============================================================================ */}
      <div style={{ width: "100%", padding: "24px 36px 0 36px", boxSizing: "border-box" }}>
        <div
          style={{
            width: "100%",
            background: "#ffffff",
            borderRadius: "20px",
            border: "1.5px solid #e2e8f0",
            padding: "20px 28px",
            boxShadow: "0 6px 20px rgba(0, 0, 0, 0.03)",
            display: "flex",
            flexDirection: "column",
            gap: "16px",
            boxSizing: "border-box"
          }}
        >
          {/* Top Search & View Switcher */}
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: "16px" }}>
            {/* Search Input Box */}
            <div style={{ position: "relative", flex: "1 1 400px", maxWidth: "800px" }}>
              <span style={{ position: "absolute", left: "16px", top: "50%", transform: "translateY(-50%)", fontSize: "1.2rem", color: "#6366f1" }}>
                🔍
              </span>
              <input
                ref={searchInputRef}
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Search across all reports, why it's essential, departments, or tags... (Press / or Ctrl+K)"
                style={{
                  width: "100%",
                  padding: "14px 105px 14px 48px",
                  borderRadius: "14px",
                  background: "#f8fafc",
                  border: searchQuery ? "2px solid #6366f1" : "1.5px solid #cbd5e1",
                  color: "#0f172a",
                  fontSize: "0.94rem",
                  fontWeight: 600,
                  outline: "none",
                  boxSizing: "border-box",
                  transition: "all 0.2s"
                }}
              />
              <div style={{ position: "absolute", right: "14px", top: "50%", transform: "translateY(-50%)", display: "flex", alignItems: "center", gap: "6px" }}>
                {searchQuery ? (
                  <button
                    onClick={() => setSearchQuery("")}
                    style={{
                      background: "#e2e8f0",
                      border: "none",
                      color: "#475569",
                      borderRadius: "50%",
                      width: "24px",
                      height: "24px",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      cursor: "pointer",
                      fontSize: "0.75rem",
                      fontWeight: 800
                    }}
                  >
                    ✕
                  </button>
                ) : (
                  <span style={{ background: "#e0e7ff", color: "#4338ca", padding: "2px 8px", borderRadius: "6px", fontSize: "0.72rem", fontWeight: 800 }}>
                    Ctrl+K
                  </span>
                )}
              </div>
            </div>

            {/* View Mode Switcher */}
            <div style={{ display: "flex", alignItems: "center", gap: "6px", background: "#f1f5f9", padding: "4px", borderRadius: "12px", border: "1.5px solid #e2e8f0" }}>
              <button
                onClick={() => setViewMode("grid")}
                style={{
                  padding: "9px 18px",
                  borderRadius: "9px",
                  border: "none",
                  background: viewMode === "grid" ? "linear-gradient(135deg, #4f46e5 0%, #6366f1 100%)" : "transparent",
                  color: viewMode === "grid" ? "#ffffff" : "#64748b",
                  boxShadow: viewMode === "grid" ? "0 4px 12px rgba(79, 70, 229, 0.3)" : "none",
                  fontWeight: 800,
                  fontSize: "0.84rem",
                  cursor: "pointer",
                  display: "flex",
                  alignItems: "center",
                  gap: "6px"
                }}
              >
                <span>🏓</span> 3D Grid
              </button>
              <button
                onClick={() => setViewMode("list")}
                style={{
                  padding: "9px 18px",
                  borderRadius: "9px",
                  border: "none",
                  background: viewMode === "list" ? "linear-gradient(135deg, #4f46e5 0%, #6366f1 100%)" : "transparent",
                  color: viewMode === "list" ? "#ffffff" : "#64748b",
                  boxShadow: viewMode === "list" ? "0 4px 12px rgba(79, 70, 229, 0.3)" : "none",
                  fontWeight: 800,
                  fontSize: "0.84rem",
                  cursor: "pointer",
                  display: "flex",
                  alignItems: "center",
                  gap: "6px"
                }}
              >
                <span>📄</span> List View
              </button>
            </div>
          </div>

          {/* Section Navigation Tabs */}
          <div style={{ display: "flex", alignItems: "center", gap: "10px", flexWrap: "wrap", borderTop: "1px solid #f1f5f9", paddingTop: "14px" }}>
            <span style={{ fontSize: "0.78rem", fontWeight: 800, color: "#64748b", textTransform: "uppercase", letterSpacing: "0.5px" }}>
              Section Filter:
            </span>
            <SectionPillTab
              id="all"
              label="✨ All 4 Sections"
              count={allAccessibleModules.length}
              active={activeSectionTab === "all"}
              onClick={() => setActiveSectionTab("all")}
              highlightColor="linear-gradient(135deg, #4f46e5 0%, #3730a3 100%)"
            />
            <SectionPillTab
              id="critical"
              label="🚨 1. Most Critical Reports"
              count={filteredCritical.length}
              active={activeSectionTab === "critical"}
              onClick={() => setActiveSectionTab("critical")}
              highlightColor="linear-gradient(135deg, #dc2626 0%, #b91c1c 100%)"
            />
            <SectionPillTab
              id="departments"
              label="🏢 2. Department Reports"
              count={filteredDepartments.length}
              active={activeSectionTab === "departments"}
              onClick={() => setActiveSectionTab("departments")}
              highlightColor="linear-gradient(135deg, #0284c7 0%, #0369a1 100%)"
            />
            <SectionPillTab
              id="admin"
              label="👑 3. Admin Special"
              count={filteredAdmin.length}
              active={activeSectionTab === "admin"}
              onClick={() => setActiveSectionTab("admin")}
              highlightColor="linear-gradient(135deg, #d97706 0%, #b45309 100%)"
            />
            <SectionPillTab
              id="other"
              label="🛠️ 4. Other Reports"
              count={filteredOther.length}
              active={activeSectionTab === "other"}
              onClick={() => setActiveSectionTab("other")}
              highlightColor="linear-gradient(135deg, #7c3aed 0%, #6d28d9 100%)"
            />
            {pinnedModules.length > 0 && (
              <SectionPillTab
                id="favorites"
                label="⭐ Bookmarked Favorites"
                count={pinnedModules.length}
                active={activeSectionTab === "favorites"}
                onClick={() => setActiveSectionTab("favorites")}
                highlightColor="linear-gradient(135deg, #f59e0b 0%, #d97706 100%)"
              />
            )}
          </div>
        </div>
      </div>

      {/* ============================================================================ */}
      {/* MAIN CONTENT AREA: COLORFUL 3D CARDS (FULL WIDTH)                            */}
      {/* ============================================================================ */}
      <main style={{ width: "100%", padding: "32px 36px 60px 36px", boxSizing: "border-box" }}>
        {/* If Active Search Query is present */}
        {searchResults !== null ? (
          <div>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: "20px" }}>
              <div>
                <h2 style={{ fontSize: "1.45rem", fontWeight: 900, color: "#0f172a", margin: 0 }}>
                  Search Results for <span style={{ color: "#4f46e5" }}>"{searchQuery}"</span>
                </h2>
                <p style={{ margin: "4px 0 0 0", fontSize: "0.88rem", color: "#64748b" }}>
                  Found {searchResults.length} matching modules across all production sections.
                </p>
              </div>
              <button
                onClick={() => setSearchQuery("")}
                style={{
                  background: "#ffffff",
                  border: "1.5px solid #cbd5e1",
                  color: "#0f172a",
                  padding: "8px 18px",
                  borderRadius: "10px",
                  fontWeight: 800,
                  fontSize: "0.84rem",
                  cursor: "pointer"
                }}
              >
                Clear Search
              </button>
            </div>

            {searchResults.length === 0 ? (
              <div style={{ textAlign: "center", padding: "80px 20px", background: "#ffffff", borderRadius: "24px", border: "1.5px dashed #cbd5e1" }}>
                <div style={{ fontSize: "3rem", marginBottom: "12px" }}>🔍</div>
                <h3 style={{ fontSize: "1.25rem", fontWeight: 800, color: "#0f172a" }}>No Matching Reports Found</h3>
                <p style={{ color: "#64748b", fontSize: "0.92rem" }}>Try searching with a different term like "Stitching", "Jaybir", "Fabric", or "Packing".</p>
              </div>
            ) : (
              <div style={{ width: "100%", display: "grid", gridTemplateColumns: viewMode === "grid" ? "repeat(auto-fill, minmax(340px, 1fr))" : "1fr", gap: "24px" }}>
                {searchResults.map((mod) => (
                  <ColorfulReportCard
                    key={mod.id + mod.path}
                    module={mod}
                    isPinned={pinnedModules.includes(mod.title)}
                    togglePin={togglePin}
                    viewMode={viewMode}
                    hoveredCard={hoveredCard}
                    setHoveredCard={setHoveredCard}
                  />
                ))}
              </div>
            )}
          </div>
        ) : activeSectionTab === "favorites" ? (
          /* Bookmarked Favorites View */
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: "10px", marginBottom: "22px" }}>
              <span style={{ fontSize: "1.6rem" }}>⭐</span>
              <h2 style={{ fontSize: "1.45rem", fontWeight: 900, color: "#0f172a", margin: 0 }}>
                Bookmarked Quick Access Reports
              </h2>
            </div>
            <div style={{ width: "100%", display: "grid", gridTemplateColumns: viewMode === "grid" ? "repeat(auto-fill, minmax(340px, 1fr))" : "1fr", gap: "24px" }}>
              {allAccessibleModules
                .filter((m) => pinnedModules.includes(m.title))
                .map((mod) => (
                  <ColorfulReportCard
                    key={'fav-' + mod.id}
                    module={mod}
                    isPinned={true}
                    togglePin={togglePin}
                    viewMode={viewMode}
                    hoveredCard={hoveredCard}
                    setHoveredCard={setHoveredCard}
                  />
                ))}
            </div>
          </div>
        ) : (
          <div style={{ width: "100%", display: "flex", flexDirection: "column", gap: "52px" }}>
            {/* ========================================================================= */}
            {/* 🚨 SECTION 1: MOST CRITICAL REPORTS - MANDATORY DAILY CHECKING REPORTS     */}
            {/* ========================================================================= */}
            {(activeSectionTab === "all" || activeSectionTab === "critical") && (
              <section id="critical" style={{ width: "100%" }}>
                <SectionHeaderBlock
                  tierNumber="SECTION 1"
                  emoji="🚨"
                  title="Most Critical Reports – Necessity Daily Checking Reports"
                  subtitle="Mandatory step-by-step production pipeline tracking bottleneck stages from cutting lay to packed carton handoff."
                  badgeText="MANDATORY DAILY CHECKPOINTS"
                  badgeBg="#fee2e2"
                  badgeColor="#dc2626"
                  badgeBorder="#fca5a5"
                  count={filteredCritical.length}
                  bannerGradient="linear-gradient(90deg, #dc2626 0%, #ea580c 35%, #ec4899 70%, #4f46e5 100%)"
                />

                {/* Workflow Sequence Indicator (Colorful) */}
                <div
                  style={{
                    width: "100%",
                    background: "#ffffff",
                    borderRadius: "16px",
                    border: "1.5px solid #fecdd3",
                    padding: "14px 20px",
                    marginBottom: "22px",
                    display: "flex",
                    alignItems: "center",
                    gap: "10px",
                    overflowX: "auto",
                    boxShadow: "0 4px 14px rgba(225, 29, 72, 0.05)",
                    boxSizing: "border-box"
                  }}
                >
                  <span style={{ fontSize: "0.82rem", fontWeight: 800, color: "#991b1b", textTransform: "uppercase", whiteSpace: "nowrap" }}>
                    🔄 Daily Audit Flow:
                  </span>
                  {filteredCritical.map((m, idx) => (
                    <React.Fragment key={'flow-' + m.id}>
                      <div
                        onClick={() => history.push(m.path)}
                        style={{
                          display: "inline-flex",
                          alignItems: "center",
                          gap: "6px",
                          background: m.accentBg || "#fff1f2",
                          border: `1px solid ${m.accentBorder || "#fecdd3"}`,
                          padding: "6px 12px",
                          borderRadius: "8px",
                          fontSize: "0.8rem",
                          fontWeight: 700,
                          color: m.badgeColor || "#9f1239",
                          whiteSpace: "nowrap",
                          cursor: "pointer",
                          transition: "all 0.2s"
                        }}
                        onMouseEnter={(e) => {
                          e.currentTarget.style.transform = "scale(1.04)";
                          e.currentTarget.style.boxShadow = `0 4px 12px ${m.glowColor}`;
                        }}
                        onMouseLeave={(e) => {
                          e.currentTarget.style.transform = "scale(1)";
                          e.currentTarget.style.boxShadow = "none";
                        }}
                      >
                        <span style={{ background: m.badgeColor || "#e11d48", color: "#ffffff", padding: "2px 6px", borderRadius: "5px", fontSize: "0.7rem", fontWeight: 900 }}>
                          #{idx + 1}
                        </span>
                        <span>{m.title}</span>
                      </div>
                      {idx < filteredCritical.length - 1 && (
                        <span style={{ color: "#f43f5e", fontWeight: 800, fontSize: "0.88rem" }}>➔</span>
                      )}
                    </React.Fragment>
                  ))}
                </div>

                {/* Critical Grid */}
                <div style={{ width: "100%", display: "grid", gridTemplateColumns: viewMode === "grid" ? "repeat(auto-fill, minmax(340px, 1fr))" : "1fr", gap: "24px" }}>
                  {filteredCritical.map((mod) => (
                    <ColorfulReportCard
                      key={mod.id}
                      module={mod}
                      isPinned={pinnedModules.includes(mod.title)}
                      togglePin={togglePin}
                      viewMode={viewMode}
                      hoveredCard={hoveredCard}
                      setHoveredCard={setHoveredCard}
                      isCriticalStep={true}
                    />
                  ))}
                </div>
              </section>
            )}

            {/* ========================================================================= */}
            {/* 🏢 SECTION 2: INDIVIDUAL DEPARTMENT REPORTS                               */}
            {/* ========================================================================= */}
            {(activeSectionTab === "all" || activeSectionTab === "departments") && (
              <section id="departments" style={{ width: "100%" }}>
                <SectionHeaderBlock
                  tierNumber="SECTION 2"
                  emoji="🏢"
                  title="Individual Department Reports"
                  subtitle="Dedicated floor-level machine logs, specialized seam tracking, and external Jaybir vendor production reconciliations."
                  badgeText="FLOOR-LEVEL OPERATIONS & VENDORS"
                  badgeBg="#e0f2fe"
                  badgeColor="#0284c7"
                  badgeBorder="#bae6fd"
                  count={filteredDepartments.length}
                  bannerGradient="linear-gradient(90deg, #0284c7 0%, #0d9488 35%, #f59e0b 70%, #ec4899 100%)"
                />

                <div style={{ width: "100%", display: "grid", gridTemplateColumns: viewMode === "grid" ? "repeat(auto-fill, minmax(340px, 1fr))" : "1fr", gap: "24px" }}>
                  {filteredDepartments.map((mod) => (
                    <ColorfulReportCard
                      key={mod.id}
                      module={mod}
                      isPinned={pinnedModules.includes(mod.title)}
                      togglePin={togglePin}
                      viewMode={viewMode}
                      hoveredCard={hoveredCard}
                      setHoveredCard={setHoveredCard}
                    />
                  ))}
                </div>
              </section>
            )}

            {/* ========================================================================= */}
            {/* 👑 SECTION 3: SPECIAL REPORTS FOR ADMIN                                    */}
            {/* ========================================================================= */}
            {(activeSectionTab === "all" || activeSectionTab === "admin") && (
              <section id="admin" style={{ width: "100%" }}>
                <SectionHeaderBlock
                  tierNumber="SECTION 3"
                  emoji="👑"
                  title="Special Reports for Admin"
                  subtitle="Executive control over high-value greige knitting, collar/cuff output, yarn warehouse balances, and procurement POs."
                  badgeText="EXECUTIVE & RAW MATERIALS"
                  badgeBg="#fef3c7"
                  badgeColor="#b45309"
                  badgeBorder="#fde68a"
                  count={filteredAdmin.length}
                  bannerGradient="linear-gradient(90deg, #b45309 0%, #7c3aed 35%, #059669 70%, #dc2626 100%)"
                />

                <div style={{ width: "100%", display: "grid", gridTemplateColumns: viewMode === "grid" ? "repeat(auto-fill, minmax(340px, 1fr))" : "1fr", gap: "24px" }}>
                  {filteredAdmin.map((mod) => (
                    <ColorfulReportCard
                      key={mod.id}
                      module={mod}
                      isPinned={pinnedModules.includes(mod.title)}
                      togglePin={togglePin}
                      viewMode={viewMode}
                      hoveredCard={hoveredCard}
                      setHoveredCard={setHoveredCard}
                      isAdminSpecial={true}
                    />
                  ))}
                </div>
              </section>
            )}

            {/* ========================================================================= */}
            {/* 🛠️ SECTION 4: OTHER FACTORY TOOLS & AUXILIARY REPORTS                      */}
            {/* ========================================================================= */}
            {(activeSectionTab === "all" || activeSectionTab === "other") && (
              <section id="other" style={{ width: "100%" }}>
                <SectionHeaderBlock
                  tierNumber="SECTION 4"
                  emoji="🛠️"
                  title="Other Factory Tools & Auxiliary Reports"
                  subtitle="AI forecasting, job orders, lot modifications, quarantine hold/cancellations, and visual factory tracking systems."
                  badgeText="ANALYTICS, LOGS & UTILITIES"
                  badgeBg="#f5f3ff"
                  badgeColor="#7c3aed"
                  badgeBorder="#ddd6fe"
                  count={filteredOther.length}
                  bannerGradient="linear-gradient(90deg, #06b6d4 0%, #10b981 25%, #f59e0b 50%, #ec4899 75%, #4f46e5 100%)"
                />

                <div style={{ width: "100%", display: "grid", gridTemplateColumns: viewMode === "grid" ? "repeat(auto-fill, minmax(340px, 1fr))" : "1fr", gap: "24px" }}>
                  {filteredOther.map((mod) => (
                    <ColorfulReportCard
                      key={mod.id}
                      module={mod}
                      isPinned={pinnedModules.includes(mod.title)}
                      togglePin={togglePin}
                      viewMode={viewMode}
                      hoveredCard={hoveredCard}
                      setHoveredCard={setHoveredCard}
                    />
                  ))}
                </div>
              </section>
            )}
          </div>
        )}

        {/* Floating Back to Top Button */}
        <AnimatePresence>
          {showBackToTop && (
            <motion.button
              initial={{ opacity: 0, scale: 0.8, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.8, y: 20 }}
              onClick={() => window.scrollTo({ top: 0, behavior: "smooth" })}
              style={{
                position: "fixed",
                bottom: "28px",
                right: "36px",
                padding: "12px 22px",
                borderRadius: "9999px",
                background: "linear-gradient(135deg, #1e1b4b 0%, #4f46e5 100%)",
                color: "#ffffff",
                border: "1.5px solid rgba(255, 255, 255, 0.2)",
                boxShadow: "0 12px 28px rgba(79, 70, 229, 0.35)",
                fontSize: "0.88rem",
                fontWeight: 800,
                cursor: "pointer",
                display: "flex",
                alignItems: "center",
                gap: "8px",
                zIndex: 99
              }}
            >
              <span>⬆️</span> Back to Top
            </motion.button>
          )}
        </AnimatePresence>
      </main>
    </div>
  );
}

// ============================================================================
// SUBCOMPONENT: SECTION HEADER BLOCK (COLORFUL GRADIENTS)
// ============================================================================
function SectionHeaderBlock({ tierNumber, emoji, title, subtitle, badgeText, badgeBg, badgeColor, badgeBorder, count, bannerGradient }) {
  return (
    <div style={{ marginBottom: "22px", position: "relative" }}>
      {/* Top Accent Line */}
      <div style={{ height: "4.5px", background: bannerGradient, borderRadius: "9999px", marginBottom: "14px", boxShadow: "0 2px 8px rgba(0,0,0,0.08)" }} />

      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", flexWrap: "wrap", gap: "14px" }}>
        <div>
          <div style={{ display: "flex", alignItems: "center", gap: "10px", flexWrap: "wrap", marginBottom: "6px" }}>
            <span
              style={{
                background: badgeBg,
                color: badgeColor,
                border: `1.5px solid ${badgeBorder}`,
                padding: "3px 10px",
                borderRadius: "8px",
                fontSize: "0.74rem",
                fontWeight: 900,
                letterSpacing: "0.5px"
              }}
            >
              {tierNumber} • {badgeText}
            </span>
            <span style={{ fontSize: "0.8rem", fontWeight: 800, color: "#64748b" }}>
              ({count} Active Modules)
            </span>
          </div>
          <h2 style={{ margin: 0, fontSize: "1.5rem", fontWeight: 900, color: "#0f172a", display: "flex", alignItems: "center", gap: "10px" }}>
            <span>{emoji}</span>
            <span>{title}</span>
          </h2>
          <p style={{ margin: "5px 0 0 0", fontSize: "0.92rem", color: "#475569", maxWidth: "950px", lineHeight: 1.45 }}>
            {subtitle}
          </p>
        </div>
      </div>
    </div>
  );
}

// ============================================================================
// SUBCOMPONENT: 3D COLORFUL REPORT CARD
// ============================================================================
function ColorfulReportCard({ module, isPinned, togglePin, viewMode, hoveredCard, setHoveredCard, isCriticalStep, isAdminSpecial }) {
  const isHovered = hoveredCard === module.id;

  if (viewMode === "list") {
    return (
      <motion.div
        layout
        whileHover={{ x: 6 }}
        transition={{ duration: 0.2 }}
        onMouseEnter={() => setHoveredCard?.(module.id)}
        onMouseLeave={() => setHoveredCard?.(null)}
        style={{
          background: "#ffffff",
          border: isHovered ? `1.5px solid ${module.badgeColor || "#4f46e5"}` : "1.5px solid #e2e8f0",
          borderRadius: "18px",
          padding: "16px 24px",
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          gap: "18px",
          boxShadow: isHovered ? `0 12px 28px ${module.glowColor}` : "0 2px 6px rgba(0, 0, 0, 0.02)",
          transition: "all 0.2s"
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: "16px", flex: 1, minWidth: 0 }}>
          {module.step && (
            <div style={{ background: module.badgeColor || "#1e1b4b", color: "#ffffff", fontSize: "0.75rem", fontWeight: 900, padding: "5px 10px", borderRadius: "8px", boxShadow: "0 2px 6px rgba(0,0,0,0.15)" }}>
              STEP #{module.step}
            </div>
          )}
          <div
            style={{
              width: "48px",
              height: "48px",
              borderRadius: "14px",
              background: module.gradient,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontSize: "1.7rem",
              boxShadow: `0 6px 14px ${module.glowColor}`,
              flexShrink: 0
            }}
          >
            {module.emoji}
          </div>
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap" }}>
              <h3 style={{ margin: 0, fontSize: "1.1rem", fontWeight: 800, color: "#0f172a" }}>
                {module.title}
              </h3>
              <span style={{ background: module.accentBg || "#f1f5f9", color: module.badgeColor || "#475569", border: `1px solid ${module.accentBorder || "#e2e8f0"}`, padding: "2px 8px", borderRadius: "6px", fontSize: "0.74rem", fontWeight: 700 }}>
                {module.department}
              </span>
            </div>
            <p style={{ margin: "4px 0 0 0", fontSize: "0.85rem", color: "#64748b", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
              <strong style={{ color: module.badgeColor || "#059669" }}>💡 Essential:</strong> {module.whyEssential}
            </p>
          </div>
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: "10px", flexShrink: 0 }}>
          <button
            onClick={(e) => togglePin(module.title, e)}
            style={{
              background: isPinned ? "#fef3c7" : "#f8fafc",
              border: isPinned ? "1.5px solid #f59e0b" : "1.5px solid #e2e8f0",
              color: isPinned ? "#d97706" : "#94a3b8",
              borderRadius: "10px",
              width: "36px",
              height: "36px",
              fontSize: "1.1rem",
              cursor: "pointer",
              display: "flex",
              alignItems: "center",
              justifyContent: "center"
            }}
          >
            ★
          </button>
          <Link
            to={module.path}
            style={{
              textDecoration: "none",
              padding: "10px 20px",
              borderRadius: "11px",
              background: module.gradient,
              color: "#ffffff",
              fontSize: "0.86rem",
              fontWeight: 800,
              display: "inline-flex",
              alignItems: "center",
              gap: "6px",
              boxShadow: `0 4px 12px ${module.glowColor}`
            }}
          >
            <span>Open</span>
            <span>➔</span>
          </Link>
        </div>
      </motion.div>
    );
  }

  // 3D Grid Card View
  return (
    <div className="perspective-card-container" style={{ width: "100%" }}>
      <div
        className="card-3d-wrapper"
        onMouseEnter={() => setHoveredCard?.(module.id)}
        onMouseLeave={() => setHoveredCard?.(null)}
        style={{
          background: "#ffffff",
          border: isHovered ? `1.5px solid ${module.badgeColor || "#4f46e5"}` : "1.5px solid #e2e8f0",
          borderRadius: "24px",
          padding: "24px",
          display: "flex",
          flexDirection: "column",
          justifyContent: "space-between",
          minHeight: "340px",
          position: "relative",
          overflow: "hidden",
          boxShadow: isHovered
            ? `0 24px 48px -12px ${module.glowColor}, 0 4px 16px rgba(0,0,0,0.04)`
            : "0 4px 14px rgba(0, 0, 0, 0.03)",
          boxSizing: "border-box"
        }}
      >
        {/* Top Accent Gradient Bar */}
        <div style={{ position: "absolute", top: 0, left: 0, right: 0, height: "4.5px", background: module.gradient }} />

        <div>
          {/* Card Header Badges (3D floating) */}
          <div className="card-3d-floating-badge" style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: "16px" }}>
            <div style={{ display: "flex", alignItems: "center", gap: "6px", flexWrap: "wrap" }}>
              {module.step && (
                <span style={{ background: module.badgeColor || "#1e1b4b", color: "#ffffff", fontSize: "0.74rem", fontWeight: 900, padding: "3px 9px", borderRadius: "6px", letterSpacing: "0.5px", boxShadow: "0 2px 6px rgba(0,0,0,0.15)" }}>
                  STEP #{String(module.step).padStart(2, '0')}
                </span>
              )}
              <span style={{ background: module.accentBg || "#f1f5f9", color: module.badgeColor || "#334155", border: `1px solid ${module.accentBorder || "#e2e8f0"}`, padding: "3px 10px", borderRadius: "9999px", fontSize: "0.74rem", fontWeight: 800 }}>
                🏢 {module.department}
              </span>
              {isAdminSpecial && (
                <span style={{ background: "#fef3c7", color: "#b45309", border: "1px solid #fde68a", padding: "3px 9px", borderRadius: "9999px", fontSize: "0.72rem", fontWeight: 800 }}>
                  👑 Admin
                </span>
              )}
            </div>

            <button
              onClick={(e) => togglePin(module.title, e)}
              title={isPinned ? "Unpin module" : "Pin to quick access"}
              style={{
                background: isPinned ? "#fef3c7" : "#f8fafc",
                border: isPinned ? "1.5px solid #f59e0b" : "1.5px solid #e2e8f0",
                color: isPinned ? "#d97706" : "#94a3b8",
                borderRadius: "10px",
                width: "34px",
                height: "34px",
                fontSize: "1.05rem",
                cursor: "pointer",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                transition: "all 0.2s"
              }}
            >
              ★
            </button>
          </div>

          {/* Title & 3D Floating Icon Header */}
          <div style={{ display: "flex", alignItems: "flex-start", gap: "16px", marginBottom: "14px" }}>
            <div
              className="card-3d-floating-icon"
              style={{
                width: "56px",
                height: "56px",
                borderRadius: "18px",
                background: module.gradient,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontSize: "2.1rem",
                boxShadow: `0 10px 22px ${module.glowColor}`,
                flexShrink: 0
              }}
            >
              {module.emoji}
            </div>
            <div style={{ flex: 1 }}>
              <h3 style={{ margin: "0 0 4px 0", fontSize: "1.18rem", fontWeight: 900, color: "#0f172a", lineHeight: 1.25, letterSpacing: "-0.01em" }}>
                {module.title}
              </h3>
              <span style={{ fontSize: "0.76rem", color: "#64748b", fontWeight: 600 }}>🕒 {module.lastUpdated}</span>
            </div>
          </div>

          {/* General Description */}
          <p style={{ margin: "0 0 14px 0", fontSize: "0.88rem", color: "#475569", lineHeight: 1.45 }}>
            {module.description}
          </p>

          {/* 💡 WHY THIS REPORT IS ESSENTIAL BOX (Colorful & Tinted) */}
          <div
            style={{
              background: module.accentBg || "#f8fafc",
              border: `1px solid ${module.accentBorder || "#e2e8f0"}`,
              borderLeft: `4px solid ${module.badgeColor || "#4f46e5"}`,
              borderRadius: "12px",
              padding: "11px 14px",
              marginBottom: "16px",
              boxShadow: "0 2px 6px rgba(0,0,0,0.02)"
            }}
          >
            <div style={{ fontSize: "0.74rem", fontWeight: 900, color: module.badgeColor || "#4f46e5", textTransform: "uppercase", letterSpacing: "0.5px", marginBottom: "4px", display: "flex", alignItems: "center", gap: "5px" }}>
              <span>💡</span> Why This Report Is Essential
            </div>
            <div style={{ fontSize: "0.82rem", color: "#1e293b", fontWeight: 600, lineHeight: 1.45 }}>
              {module.whyEssential}
            </div>
          </div>

          {/* Feature Pills */}
          <div style={{ display: "flex", gap: "6px", flexWrap: "wrap", marginBottom: "18px" }}>
            {module.features.map((feat, i) => (
              <span
                key={i}
                style={{
                  background: "#f1f5f9",
                  color: "#334155",
                  padding: "3px 9px",
                  borderRadius: "7px",
                  fontSize: "0.74rem",
                  fontWeight: 600,
                  border: "1px solid #e2e8f0"
                }}
              >
                {feat}
              </span>
            ))}
          </div>
        </div>

        {/* 3D Floating Colorful Action Button */}
        <Link
          to={module.path}
          className="card-3d-floating-btn"
          style={{
            textDecoration: "none",
            padding: "12px 18px",
            borderRadius: "14px",
            background: isHovered ? module.gradient : (module.accentBg || "#f8fafc"),
            color: isHovered ? "#ffffff" : (module.badgeColor || "#0f172a"),
            fontSize: "0.88rem",
            fontWeight: 900,
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            border: isHovered ? `1px solid ${module.badgeColor}` : `1.5px solid ${module.accentBorder || "#e2e8f0"}`,
            boxShadow: isHovered ? `0 8px 20px ${module.glowColor}` : "none",
            transition: "all 0.25s ease"
          }}
        >
          <span>Open {isCriticalStep ? "Daily Check" : "Report"}</span>
          <span style={{ transform: isHovered ? "translateX(5px)" : "translateX(0)", transition: "transform 0.2s" }}>
            ➔
          </span>
        </Link>
      </div>
    </div>
  );
}

// ============================================================================
// SUBCOMPONENT: SECTION PILL TAB (COLORFUL GRADIENTS)
// ============================================================================
function SectionPillTab({ id, label, count, active, onClick, highlightColor }) {
  const activeBg = highlightColor || "linear-gradient(135deg, #0f172a 0%, #1e293b 100%)";

  return (
    <button
      onClick={onClick}
      style={{
        background: active ? activeBg : "#ffffff",
        color: active ? "#ffffff" : "#334155",
        border: active ? "1.5px solid transparent" : "1.5px solid #cbd5e1",
        padding: "8px 16px",
        borderRadius: "9999px",
        fontSize: "0.82rem",
        fontWeight: 800,
        cursor: "pointer",
        display: "inline-flex",
        alignItems: "center",
        gap: "7px",
        transition: "all 0.2s",
        boxShadow: active ? "0 4px 14px rgba(0, 0, 0, 0.15)" : "none"
      }}
    >
      <span>{label}</span>
      <span
        style={{
          background: active ? "rgba(255, 255, 255, 0.25)" : "#f1f5f9",
          color: active ? "#ffffff" : "#475569",
          padding: "2px 7px",
          borderRadius: "9999px",
          fontSize: "0.72rem",
          fontWeight: 900
        }}
      >
        {count}
      </span>
    </button>
  );
}

// ============================================================================
// SUBCOMPONENT: HERO SECTION KPI STAT (COLORFUL ACCENTS)
// ============================================================================
function HeroSectionStat({ title, count, icon, color, bg, border, active, onClick }) {
  return (
    <div
      onClick={onClick}
      style={{
        background: active ? bg : "#ffffff",
        border: active ? `2px solid ${color}` : `1.5px solid ${border}`,
        borderRadius: "14px",
        padding: "12px 16px",
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        cursor: "pointer",
        transition: "all 0.22s ease",
        boxShadow: active ? `0 6px 16px ${color}35` : "0 2px 6px rgba(0, 0, 0, 0.03)",
        transform: active ? "scale(1.02)" : "scale(1)"
      }}
    >
      <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
        <span style={{ fontSize: "1.35rem" }}>{icon}</span>
        <div>
          <div style={{ fontSize: "0.82rem", fontWeight: 800, color: "#0f172a" }}>{title}</div>
          <div style={{ fontSize: "0.72rem", fontWeight: 600, color: "#64748b" }}>Click to filter</div>
        </div>
      </div>
      <span
        style={{
          background: color,
          color: "#ffffff",
          padding: "3px 10px",
          borderRadius: "8px",
          fontSize: "0.88rem",
          fontWeight: 900,
          boxShadow: `0 2px 6px ${color}50`
        }}
      >
        {count}
      </span>
    </div>
  );
}
