import React, { useState, useEffect, useMemo } from "react";
import { Link, useHistory } from "react-router-dom";
import { motion, AnimatePresence } from "framer-motion";
import { getCurrentUser, logoutUser } from "./auth";

export default function Dashboard() {
  const history = useHistory();
  const currentUser = getCurrentUser();

  const handleLogout = () => {
    logoutUser();
    history.push("/");
  };
  const [activeCategory, setActiveCategory] = useState("All");
  const [searchQuery, setSearchQuery] = useState("");
  const [quickTag, setQuickTag] = useState("ALL"); // 'ALL' | 'CRITICAL' | 'LIVE' | 'AI' | 'PENDING'
  const [viewMode, setViewMode] = useState("grid"); // 'grid' | 'list' | 'index'
  const [showBackToTop, setShowBackToTop] = useState(false);
  const searchInputRef = React.useRef(null);
  const [pinnedModules, setPinnedModules] = useState(() => {
    try {
      const saved = localStorage.getItem("pinned_dashboard_modules");
      return saved ? JSON.parse(saved) : ["Short Summary Report", "Stitching Issue", "Cut to Pack Report"];
    } catch {
      return ["Short Summary Report", "Stitching Issue", "Cut to Pack Report"];
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
      if (window.scrollY > 400) {
        setShowBackToTop(true);
      } else {
        setShowBackToTop(false);
      }
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

  const modules = [
    {
      title: "Cutting Report",
      emoji: "✂️",
      path: "/cutting-report",
      gradient: "linear-gradient(135deg, #11998e 0%, #38ef7d 100%)",
      glowColor: "rgba(17, 153, 142, 0.4)",
      description: "Cutting operations report with fabric consumption and layer details",
      category: "Reports",
      department: "Cutting",
      status: "reports",
      features: ["Fabric usage", "Layer summary", "Marker efficiency"],
      lastUpdated: "Yesterday"
    },
    {
      title: "Embroidery Challan",
      emoji: "🧵",
      path: "/embroidery",
      gradient: "linear-gradient(135deg, #6366f1 0%, #a855f7 100%)",
      glowColor: "rgba(99, 102, 241, 0.4)",
      description: "Create and manage embroidery production challans with real-time tracking",
      category: "Production",
      department: "Embroidery & Printing",
      status: "active",
      features: ["Batch tracking", "Quality check", "Thread consumption"],
      lastUpdated: "Today"
    },
    {
      title: "Printing Challan",
      emoji: "🖨️",
      path: "/printing",
      gradient: "linear-gradient(135deg, #ec4899 0%, #f43f5e 100%)",
      glowColor: "rgba(236, 72, 153, 0.4)",
      description: "Handle printing orders, screen management, and production workflows",
      category: "Production",
      department: "Embroidery & Printing",
      status: "active",
      features: ["Screen tracking", "Color wise", "Print quality"],
      lastUpdated: "Today"
    },
    {
      title: "Overall Stitching Report",
      emoji: "✅",
      path: "/stitching-complete-lot",
      gradient: "linear-gradient(135deg, #059669 0%, #10b981 100%)",
      glowColor: "rgba(5, 150, 105, 0.4)",
      description: "View and manage completed stitching lots with batch-wise performance metrics",
      category: "Production",
      department: "Stitching",
      status: "active",
      features: ["Lot summary", "Efficiency calc", "Defect analysis"],
      lastUpdated: "Yesterday"
    },
    {
      title: "After EMB/PRINT DONE",
      emoji: "📋",
      path: "/pending-issue-to-stitching",
      gradient: "linear-gradient(135deg, #ef4444 0%, #f97316 100%)",
      glowColor: "rgba(239, 68, 68, 0.4)",
      description: "Track garments pending for stitching after embroidery/printing completion",
      category: "Production",
      department: "Embroidery & Printing",
      status: "pending",
      features: ["Ready for stitch", "Batch transfer", "Quality pending"],
      lastUpdated: "2 hrs ago"
    },
    {
      title: "After Cutting Done",
      emoji: "⏳",
      path: "/pending-stitching",
      gradient: "linear-gradient(135deg, #f59e0b 0%, #d97706 100%)",
      glowColor: "rgba(245, 158, 11, 0.4)",
      description: "Monitor pending stitching operations and delayed batches after cutting",
      category: "Production",
      department: "Stitching",
      status: "pending",
      features: ["Delay analysis", "Priority queue", "Resource allocation"],
      lastUpdated: "5 hrs ago"
    },
    {
      title: "Pending Packing to Issue",
      emoji: "📦⏳",
      path: "/pending-packing-issue",
      gradient: "linear-gradient(135deg, #f97316 0%, #ea580c 100%)",
      glowColor: "rgba(249, 115, 22, 0.4)",
      description: "Track garments that completed stitching and await packing operations",
      category: "Production",
      department: "Packing & Finishing",
      status: "pending",
      features: ["Stitching done", "Ready for pack", "Quality pending", "Batch transfer"],
      lastUpdated: "Live"
    },
    {
      title: "Packing Alloted Lot",
      emoji: "📦⏳",
      path: "/packing-alloted-lot",
      gradient: "linear-gradient(135deg, #0284c7 0%, #2563eb 100%)",
      glowColor: "rgba(2, 132, 199, 0.4)",
      description: "Manage lots allocated for packing with operator assignment and carton planning",
      category: "Production",
      department: "Packing & Finishing",
      status: "active",
      features: ["Lot allocation", "Carton planning", "Operator assign", "Target tracking"],
      lastUpdated: "Today"
    },
    {
      title: "Stitching Issue",
      emoji: "🧶",
      path: "/stitching",
      gradient: "linear-gradient(135deg, #06b6d4 0%, #3b82f6 100%)",
      glowColor: "rgba(6, 182, 212, 0.4)",
      description: "Manage daily stitching operations, line assignments, and production tracking",
      category: "Production",
      department: "Stitching",
      status: "active",
      features: ["Line wise", "Operator tracking", "Hourly target"],
      lastUpdated: "Today"
    },
    {
      title: "Issue To Packing",
      emoji: "📦",
      path: "/issue-to-packing",
      gradient: "linear-gradient(135deg, #10b981 0%, #0d9488 100%)",
      glowColor: "rgba(16, 185, 129, 0.4)",
      description: "Manage packing operations, carton preparation, and shipment readiness",
      category: "Production",
      department: "Packing & Finishing",
      status: "active",
      features: ["Carton tracking", "Quality check", "Shipment prep"],
      lastUpdated: "Today"
    },
    {
      title: "Packing Complete Lots",
      emoji: "📦✅",
      path: "/packing-completed-lots",
      gradient: "linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%)",
      glowColor: "rgba(37, 99, 235, 0.4)",
      description: "View all packing issue lots with verified packing complete dates, aging analysis, and barcode linkages",
      category: "Production",
      department: "Packing & Finishing",
      status: "active",
      features: ["Packing complete date", "Aging days analysis", "Barcode linkage", "Excel & PDF export"],
      lastUpdated: "Real-time"
    },
    {
      title: "EMB/PRINT REMARKS",
      emoji: "📝",
      path: "/emb-print-remarks",
      gradient: "linear-gradient(135deg, #8b5cf6 0%, #d946ef 100%)",
      glowColor: "rgba(139, 92, 246, 0.4)",
      description: "View and update remarks for pending Embroidery & Printing lots",
      category: "Reports",
      department: "Embroidery & Printing",
      status: "reports",
      features: ["Embroidery Pending", "Printing Pending", "Remarks Tracking"],
      lastUpdated: "Real-time"
    },
    {
      title: "Daily Printing Challan",
      emoji: "📊",
      path: "/daily-printing-challan",
      gradient: "linear-gradient(135deg, #ff7e5f 0%, #feb47b 100%)",
      glowColor: "rgba(255, 126, 95, 0.4)",
      description: "Daily printing reports, screen usage, and production analytics",
      category: "Reports",
      department: "Embroidery & Printing",
      status: "reports",
      features: ["Screen utilization", "Color output", "Operator wise"],
      lastUpdated: "Daily"
    },
    {
      title: "Lot Change Logs",
      emoji: "📋",
      path: "/lot-logs",
      gradient: "linear-gradient(135deg, #4f46e5 0%, #312e81 100%)",
      glowColor: "rgba(79, 70, 229, 0.4)",
      description: "Log lot modifications, authorization authority, and broadcast notifications",
      category: "Management",
      department: "Executive & Analytics",
      status: "active",
      features: ["Permission log", "Lot # search", "Live Notifications"],
      lastUpdated: "Realtime"
    },
    {
      title: "Daily Embroidery Challan",
      emoji: "📈",
      path: "/daily-embroidery-challan",
      gradient: "linear-gradient(135deg, #00c6ff 0%, #0072ff 100%)",
      glowColor: "rgba(0, 198, 255, 0.4)",
      description: "Daily embroidery progress, machine efficiency, and thread consumption",
      category: "Reports",
      department: "Embroidery & Printing",
      status: "reports",
      features: ["Machine wise", "Stitch count", "Thread usage"],
      lastUpdated: "Daily"
    },
    {
      title: "Cancelled Lots Report",
      emoji: "🚫",
      path: "/cancelled-lots-report",
      gradient: "linear-gradient(135deg, #e11d48 0%, #be123c 100%)",
      glowColor: "rgba(225, 29, 72, 0.4)",
      description: "Detailed report and tracking of all cancelled job orders and cancelled lots directly from JobOrder sheet",
      category: "Reports",
      department: "Cutting",
      status: "reports",
      features: ["Cancelled lots audit", "Cancellation reasons & approvals", "Party & brand breakdown", "Excel & PDF export"],
      lastUpdated: "Real-time"
    },
    {
      title: "Hold Lot Action",
      emoji: "⏸️",
      path: "/hold-lot",
      gradient: "linear-gradient(135deg, #f59e0b 0%, #d97706 100%)",
      glowColor: "rgba(245, 158, 11, 0.4)",
      description: "Put production lots on hold by department, fetch JobOrder details, and submit with authorized approvals",
      category: "Management",
      department: "Cutting",
      status: "active",
      features: ["Department selection", "JobOrder details fetch", "Hold By & Approved By audit", "Live factory alerts"],
      lastUpdated: "Live"
    },
    {
      title: "Daily Fabric Issue Analytics",
      emoji: "🧵",
      path: "/daily-fabric-issue-report",
      gradient: "linear-gradient(135deg, #2563eb 0%, #3b82f6 100%)",
      glowColor: "rgba(37, 99, 235, 0.4)",
      description: "Visual distribution matrix mapping overall volume and weights across tables & fabric styles",
      category: "Reports",
      department: "Cutting",
      status: "reports",
      features: ["Table summary", "Fabric summary", "Weight & rolls analytics"],
      lastUpdated: "Daily"
    },
    {
      title: "Daily Packing Report",
      emoji: "📦",
      path: "/daily-packing-report",
      gradient: "linear-gradient(135deg, #f59e0b 0%, #b45309 100%)",
      glowColor: "rgba(245, 158, 11, 0.4)",
      description: "Daily packing output, carton utilization, and shipment tracking",
      category: "Reports",
      department: "Packing & Finishing",
      status: "reports",
      features: ["Carton count", "Shipment ready", "Packing efficiency"],
      lastUpdated: "Daily"
    },
    {
      title: "Daily Folding Report",
      emoji: "👕",
      path: "/daily-folding-report",
      gradient: "linear-gradient(135deg, #10b981 0%, #047857 100%)",
      glowColor: "rgba(16, 185, 129, 0.4)",
      description: "Daily folding operations, style-wise output, and productivity tracking",
      category: "Reports",
      department: "Packing & Finishing",
      status: "reports",
      features: ["Style wise", "Operator output", "Quality check"],
      lastUpdated: "Daily"
    },
    {
      title: "Daily Kaj Button Report",
      emoji: "🔘",
      path: "/daily-kaj-button-report",
      gradient: "linear-gradient(135deg, #ec4899 0%, #9d174d 100%)",
      glowColor: "rgba(236, 72, 153, 0.4)",
      description: "Daily button attachment, kaj work, and accessory tracking",
      category: "Reports",
      department: "Stitching",
      status: "reports",
      features: ["Button count", "Kaj quantity", "Accessory used"],
      lastUpdated: "Daily"
    },
    {
      title: "Daily Overlock Report",
      emoji: "🧵",
      path: "/daily-overlock-report",
      gradient: "linear-gradient(135deg, #6366f1 0%, #3730a3 100%)",
      glowColor: "rgba(99, 102, 241, 0.4)",
      description: "Daily overlock stitching production and quality check reports",
      category: "Reports",
      department: "Stitching",
      status: "reports",
      features: ["Machine output", "Thread consumption", "Quality check"],
      lastUpdated: "Daily"
    },
    {
      title: "Not Updation Report",
      emoji: "⚠️",
      path: "/daily-stitching-report-not-updation",
      gradient: "linear-gradient(135deg, #f97316 0%, #c2410c 100%)",
      glowColor: "rgba(249, 115, 22, 0.4)",
      description: "Track stitching reports that require updates or pending entries",
      category: "Reports",
      department: "Stitching",
      status: "reports",
      features: ["Missing entries", "Due reports", "Line pending"],
      lastUpdated: "2 hrs ago"
    },
    {
      title: "Cut to Pack Report",
      emoji: "📊",
      path: "/overall-cutting-to-packing-report",
      gradient: "linear-gradient(135deg, #8b5cf6 0%, #5b21b6 100%)",
      glowColor: "rgba(139, 92, 246, 0.4)",
      description: "Complete production tracking from cutting to packing with stage analysis",
      category: "Reports",
      department: "Executive & Analytics",
      status: "reports",
      features: ["WIP tracking", "Stage analysis", "Bottleneck detect"],
      lastUpdated: "Weekly"
    },
    {
      title: "Short Summary Report",
      emoji: "⚡",
      path: "/short-summary-report",
      gradient: "linear-gradient(135deg, #ff416c 0%, #ff4b2b 100%)",
      glowColor: "rgba(255, 65, 108, 0.5)",
      description: "Quick overview of key production metrics across all departments",
      category: "Reports",
      department: "Executive & Analytics",
      status: "reports",
      isLive: true,
      features: ["Live dashboard", "KPI tracking", "Alerts & notifications"],
      lastUpdated: "Real-time"
    },
    {
      title: "Sticker Report",
      emoji: "🏷️",
      path: "/sticker-report",
      gradient: "linear-gradient(135deg, #ff512f 0%, #f09819 100%)",
      glowColor: "rgba(255, 81, 47, 0.4)",
      description: "Track sticker printing, usage, and inventory management for labels and tags",
      category: "Materials & Inventory",
      department: "Materials & Inventory",
      status: "reports",
      features: ["Sticker inventory", "Size wise", "Print quality", "Usage tracking"],
      lastUpdated: "Today"
    },
    {
      title: "Fabric Roll Prediction",
      emoji: "📏",
      path: "/fabric-roll-prediction",
      gradient: "linear-gradient(135deg, #1a2980 0%, #26d0ce 100%)",
      glowColor: "rgba(38, 208, 206, 0.4)",
      description: "AI-powered fabric consumption prediction, roll optimization, and forecasting",
      category: "Reports",
      department: "Cutting",
      status: "reports",
      isAI: true,
      features: ["Roll optimization", "Consumption forecast", "Smart cutting plan"],
      lastUpdated: "Real-time"
    },
    {
      title: "All Cutting Job Orders",
      emoji: "📝",
      path: "/all-cutting-joborders",
      gradient: "linear-gradient(135deg, #475569 0%, #1e293b 100%)",
      glowColor: "rgba(71, 85, 105, 0.4)",
      description: "Centralized view of all job orders, cutting assignments, and order status",
      category: "Production",
      department: "Cutting",
      status: "active",
      features: ["Job tracking", "Order status", "Cutting queues"],
      lastUpdated: "Today"
    },
    {
      title: "Embroidery Material Receiving",
      emoji: "🧵📥",
      path: "/embroidery-material-receiving",
      gradient: "linear-gradient(135deg, #8b5cf6 0%, #6d28d9 100%)",
      glowColor: "rgba(139, 92, 246, 0.4)",
      description: "Track embroidery raw material receipts including threads, backing, needles",
      category: "Materials & Inventory",
      department: "Materials & Inventory",
      status: "active",
      features: ["Thread inventory", "Backing stock", "Supplier wise"],
      lastUpdated: "Today"
    },
    {
      title: "Printing Material Receiving",
      emoji: "🖨️📥",
      path: "/printing-material-receiving",
      gradient: "linear-gradient(135deg, #ec4899 0%, #be185d 100%)",
      glowColor: "rgba(236, 72, 153, 0.4)",
      description: "Monitor printing raw material receipts including inks, screens, and chemicals",
      category: "Materials & Inventory",
      department: "Materials & Inventory",
      status: "active",
      features: ["Ink inventory", "Screen stock", "Batch tracking"],
      lastUpdated: "Today"
    },
    {
      title: "Notification Center",
      emoji: "🔔",
      path: "/notifications",
      gradient: "linear-gradient(135deg, #ec4899 0%, #8b5cf6 100%)",
      glowColor: "rgba(236, 72, 153, 0.4)",
      description: "Full-screen activity feed for all factory alerts, stitching, cutting, and lot changes",
      category: "Reports",
      department: "Executive & Analytics",
      status: "active",
      isLive: true,
      features: ["Instagram feed", "Category filters", "Real-time alerts", "1-click navigation"],
      lastUpdated: "Real-time"
    },
    {
      title: "Knitting Report",
      emoji: "🧶",
      path: "/knitting-report",
      gradient: "linear-gradient(135deg, #a855f7 0%, #6366f1 100%)",
      glowColor: "rgba(168, 85, 247, 0.4)",
      description: "Track knitting machine production, yarn conversion efficiency, and fabric roll outputs",
      category: "Reports",
      department: "Materials & Inventory",
      status: "reports",
      features: ["Machine gauge", "Yarn conversion", "GSM & Loop length", "Shift output"],
      lastUpdated: "Real-time"
    },
    {
      title: "Collar Report",
      emoji: "👔",
      path: "/collar-report",
      gradient: "linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%)",
      glowColor: "rgba(59, 130, 246, 0.4)",
      description: "Monitor collar and cuff knitting production, size-wise counts, and machine utilization",
      category: "Reports",
      department: "Materials & Inventory",
      status: "reports",
      features: ["Collar & cuff count", "Size-wise breakdown", "Flat knitting gauge", "Defect tracking"],
      lastUpdated: "Daily"
    },
    {
      title: "Kora Roll Report",
      emoji: "📜",
      path: "/kora-roll-report",
      gradient: "linear-gradient(135deg, #10b981 0%, #059669 100%)",
      glowColor: "rgba(168, 85, 247, 0.4)",
      description: "Comprehensive tracking for un-dyed kora fabric rolls, weight in kgs, meters, and stock status",
      category: "Materials & Inventory",
      department: "Materials & Inventory",
      status: "reports",
      features: ["Kora roll inventory", "Weight (Kgs) & Meters", "Dyeing dispatch", "Roll barcode tracking"],
      lastUpdated: "Real-time"
    },
    {
      title: "Yarn Stock Report",
      emoji: "🧵",
      path: "/yarn-report",
      gradient: "linear-gradient(135deg, #4338ca 0%, #1e1b4b 100%)",
      glowColor: "rgba(67, 56, 202, 0.4)",
      description: "Knitting department yarn stock summary, packages count, and balance weight in kgs",
      category: "Materials & Inventory",
      department: "Materials & Inventory",
      status: "reports",
      features: ["Yarn stock inventory", "Pkgs & Bal WT (Kgs)", "Negative stock alerts", "Date block navigation"],
      lastUpdated: "Real-time"
    },
    {
      title: "Yarn Report",
      emoji: "🧵",
      path: "/yarn-report",
      gradient: "linear-gradient(135deg, #f59e0b 0%, #d97706 100%)",
      glowColor: "rgba(245, 158, 11, 0.4)",
      description: "Manage raw yarn stock, count-wise inventory, supplier receipts, and consumption analytics",
      category: "Materials & Inventory",
      department: "Materials & Inventory",
      status: "reports",
      features: ["Yarn count & lot", "Bale & Bag stock", "Supplier receipts", "Consumption analytics"],
      lastUpdated: "Today"
    },
    {
      title: "Pending Zip PO Report",
      emoji: "🤐⏳",
      path: "/pending-zip-po-report",
      gradient: "linear-gradient(135deg, #1e1b4b 0%, #4338ca 100%)",
      glowColor: "rgba(67, 56, 202, 0.4)",
      description: "Track production lots requiring zips where Purchase Orders (PO) have NOT been created",
      category: "Reports",
      department: "Materials & Inventory",
      status: "reports",
      features: ["Zip Required", "PO Pending", "Lot-wise tracking", "Export PDF/CSV"],
      lastUpdated: "Real-time"
    },
    {
      title: "Production Flowchart Poster",
      emoji: "🗺️",
      path: "/production-flowchart",
      gradient: "linear-gradient(135deg, #6366f1 0%, #a855f7 100%)",
      glowColor: "rgba(99, 102, 241, 0.5)",
      description: "Comprehensive visual blueprint and infographic poster mapping all 37 reports across 6 stages",
      category: "Reports",
      department: "Executive & Analytics",
      status: "active",
      isLive: true,
      features: ["Infographic Poster", "6 Production Stages", "Print / PDF Mode", "1-Click Launch"],
      lastUpdated: "Real-time"
    },
    {
      title: "Jaybir Printing Report",
      emoji: "🖨️",
      path: "/jaybir-printing-report",
      gradient: "linear-gradient(135deg, #ec4899 0%, #be185d 100%)",
      glowColor: "rgba(236, 72, 153, 0.4)",
      description: "Jaybir printing unit production report, screen tracking, and lot-wise dispatch analytics",
      category: "Reports",
      department: "Embroidery & Printing",
      status: "reports",
      features: ["Jaybir Unit", "Screen tracking", "Lot-wise Output", "Dispatch Analytics"],
      lastUpdated: "Daily"
    },
    {
      title: "Jaybir Embroidery Report",
      emoji: "🧵",
      path: "/jaybir-embroidery-report",
      gradient: "linear-gradient(135deg, #8b5cf6 0%, #6d28d9 100%)",
      glowColor: "rgba(139, 92, 246, 0.4)",
      description: "Jaybir embroidery unit daily stitch output, head utilization, and challan tracking",
      category: "Reports",
      department: "Embroidery & Printing",
      status: "reports",
      features: ["Jaybir Unit", "Stitch count", "Challan Tracking", "Quality Inspection"],
      lastUpdated: "Daily"
    },
    {
      title: "Washing Report",
      emoji: "🧼",
      path: "/washing-report",
      gradient: "linear-gradient(135deg, #0ea5e9 0%, #0284c7 100%)",
      glowColor: "rgba(14, 165, 233, 0.4)",
      description: "Garment washing cycle status, lot receipts, shrinkage testing, and finishing handover",
      category: "Reports",
      department: "Packing & Finishing",
      status: "reports",
      features: ["Wash cycle tracking", "Shrinkage & Shade", "Lot receipts", "Dispatch to Packing"],
      lastUpdated: "Real-time"
    },
    {
      title: "Feed Up Report",
      emoji: "🧵",
      path: "/feed-up-report",
      gradient: "linear-gradient(135deg, #0284c7 0%, #0369a1 100%)",
      glowColor: "rgba(2, 132, 199, 0.4)",
      description: "Feed up department output, supervisor tracking, WIP aging, and completed lot dispatch",
      category: "Reports",
      department: "Stitching & Assembly",
      status: "reports",
      features: ["Feed up tracking", "Supervisor wise", "Lot completion", "Live Sheet Sync"],
      lastUpdated: "Real-time"
    },
    {
      title: "Lot Lifecycle Timeline (Amazon Tracker)",
      emoji: "📦⚡",
      path: "/lot-timeline",
      gradient: "linear-gradient(135deg, #f59e0b 0%, #ea580c 50%, #dc2626 100%)",
      glowColor: "rgba(245, 158, 11, 0.5)",
      description: "Amazon Order Tracking-style visual timeline: Cutting Date, EMB/Print, Stitching, Kaj Button, and Packing milestones",
      category: "Management",
      department: "Executive & Analytics",
      status: "active",
      isLive: true,
      features: ["Amazon Stepper UI", "Cutting to Pack", "Live Stage Aging", "Lot Search"],
      lastUpdated: "Real-time"
    }
  ];

  const isAdmin = (currentUser?.role || "").trim().toLowerCase() === "admin";

  const adminOnlyPaths = [
    "/collar-report",
    "/yarn-report",
    "/knitting-report",
    "/short-summary-report",
    "/kora-roll-report"
  ];

  // Filter modules accessible by current user role
  const accessibleModules = useMemo(() => {
    return modules.filter((m) => {
      if (adminOnlyPaths.includes(m.path) && !isAdmin) {
        return false;
      }
      return true;
    });
  }, [modules, isAdmin]);

  // Urgent reports displayed in exact priority sequence for mandatory daily checking
  const urgentPaths = useMemo(() => [
    "/cutting-report",
    "/embroidery",
    "/printing",
    "/stitching-complete-lot",
    "/pending-issue-to-stitching",
    "/pending-stitching",
    "/pending-packing-issue",
    "/packing-alloted-lot"
  ], []);

  const urgentModules = useMemo(() => {
    return urgentPaths
      .map((path) => accessibleModules.find((m) => m.path === path))
      .filter(Boolean);
  }, [accessibleModules, urgentPaths]);

  const regularModules = useMemo(() => {
    return accessibleModules.filter((m) => !urgentPaths.includes(m.path));
  }, [accessibleModules, urgentPaths]);

  // Department definitions & colors
  const departments = useMemo(() => [
    { id: "Cutting", label: "Cutting", emoji: "✂️", color: "#2563eb", gradient: "linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%)" },
    { id: "Stitching", label: "Stitching", emoji: "🪡", color: "#06b6d4", gradient: "linear-gradient(135deg, #0e7490 0%, #06b6d4 100%)" },
    { id: "Embroidery & Printing", label: "Emb & Print", emoji: "🎨", color: "#a855f7", gradient: "linear-gradient(135deg, #6b21a8 0%, #a855f7 100%)" },
    { id: "Packing & Finishing", label: "Packing & Finishing", emoji: "📦", color: "#10b981", gradient: "linear-gradient(135deg, #047857 0%, #10b981 100%)" },
    { id: "Materials & Inventory", label: "Materials & Inventory", emoji: "🧵", color: "#f59e0b", gradient: "linear-gradient(135deg, #b45309 0%, #f59e0b 100%)" },
    { id: "Executive & Analytics", label: "Executive & Analytics", emoji: "📊", color: "#8b5cf6", gradient: "linear-gradient(135deg, #4c1d95 0%, #8b5cf6 100%)" }
  ], []);

  // Department-wise counts mapping
  const departmentCounts = useMemo(() => {
    const counts = {};
    departments.forEach(dept => {
      counts[dept.id] = accessibleModules.filter(m => m.department === dept.id).length;
    });
    return counts;
  }, [accessibleModules, departments]);

  // Filter modules based on search query, quickTag, and active department/category
  const filteredModules = useMemo(() => {
    return accessibleModules.filter((m) => {
      const matchesSearch =
        searchQuery.trim() === "" ||
        m.title.toLowerCase().includes(searchQuery.toLowerCase()) ||
        m.description.toLowerCase().includes(searchQuery.toLowerCase()) ||
        (m.department && m.department.toLowerCase().includes(searchQuery.toLowerCase())) ||
        m.features.some((f) => f.toLowerCase().includes(searchQuery.toLowerCase()));

      if (!matchesSearch) return false;

      // Quick Tag filters
      if (quickTag === "CRITICAL" && !urgentPaths.includes(m.path)) return false;
      if (quickTag === "LIVE" && !m.isLive) return false;
      if (quickTag === "AI" && !m.isAI) return false;
      if (quickTag === "PENDING" && m.status !== "pending") return false;

      if (activeCategory === "All") return true;
      if (activeCategory === "Favorites") return pinnedModules.includes(m.title);
      return m.department === activeCategory || m.category === activeCategory;
    });
  }, [accessibleModules, searchQuery, activeCategory, pinnedModules, quickTag, urgentPaths]);

  // Quick Stats
  const totalCount = accessibleModules.length;

  const categories = [
    { id: "All", label: "All Reports", emoji: "📊", count: totalCount },
    { id: "Favorites", label: "Favorites", emoji: "⭐", count: pinnedModules.length },
    ...departments.map(d => ({
      id: d.id,
      label: d.label,
      emoji: d.emoji,
      count: departmentCounts[d.id] || 0
    }))
  ];

  // Departmental Hierarchy Directory Index (Parent Module -> Sub-Reports)
  const departmentIndex = [
    {
      id: "cutting",
      title: "Cutting Department",
      emoji: "✂️",
      color: "#2563eb",
      gradient: "linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%)",
      description: "Fabric roll issuance, cutting table assignments, layer planning & AI roll optimization",
      reports: [
        {
          title: "Daily Fabric Issue Analytics",
          path: "/daily-fabric-issue-report",
          emoji: "🧵",
          desc: "Visual distribution matrix mapping overall volume and weights across tables & fabric styles",
          badge: "NEW & LIVE"
        },
        {
          title: "Cutting Stats & Analysis",
          path: "/cutting-stats",
          emoji: "📊",
          desc: "Complete cutting statistics, lot efficiency, layer summaries and cut vs pack metrics"
        },
        {
          title: "Cutting Report",
          path: "/cutting-report",
          emoji: "✂️",
          desc: "Cutting operations report with fabric consumption details and layer summaries"
        },
        {
          title: "All Cutting Job Orders",
          path: "/all-cutting-joborders",
          emoji: "📝",
          desc: "Centralized view of all active job orders, cutting assignments, and order status"
        },
        {
          title: "Fabric Roll Prediction (AI)",
          path: "/fabric-roll-prediction",
          emoji: "📏",
          desc: "AI-powered fabric roll optimization, consumption forecast, and smart cutting plans",
          badge: "AI POWERED"
        }
      ]
    },
    {
      id: "stitching",
      title: "Stitching & Production Operations",
      emoji: "🪡",
      color: "#06b6d4",
      gradient: "linear-gradient(135deg, #0e7490 0%, #06b6d4 100%)",
      description: "Daily line assignments, hourly target tracking, overlock, button/kaj work & stitching output",
      reports: [
        {
          title: "Stitching Issue Operations",
          path: "/stitching",
          emoji: "🧶",
          desc: "Manage daily stitching operations, line assignments, operator tracking and hourly targets"
        },
        {
          title: "After Cutting Done (Pending Stitching)",
          path: "/pending-stitching",
          emoji: "⏳",
          desc: "Monitor pending stitching operations and delayed batches after cutting completion"
        },
        {
          title: "Overall Stitching Report (Completed Lots)",
          path: "/stitching-complete-lot",
          emoji: "✅",
          desc: "View and manage completed stitching lots with batch-wise performance metrics"
        },
        {
          title: "Daily Kaj Button Report",
          path: "/daily-kaj-button-report",
          emoji: "🔘",
          desc: "Daily button attachment, kaj work output, and accessory tracking"
        },
        {
          title: "Daily Overlock Report",
          path: "/daily-overlock-report",
          emoji: "🧵",
          desc: "Daily overlock stitching production, machine output, and quality check reports"
        },
        {
          title: "Not Updation Report",
          path: "/daily-stitching-report-not-updation",
          emoji: "⚠️",
          desc: "Track stitching reports requiring updates, missing entries, or pending line updates"
        }
      ]
    },
    {
      id: "embroidery_printing",
      title: "Embroidery & Printing Sub-Department",
      emoji: "🎨",
      color: "#a855f7",
      gradient: "linear-gradient(135deg, #6b21a8 0%, #a855f7 100%)",
      description: "Challan issuance, screen & thread tracking, remarks log & post-embroidery stitching dispatch",
      reports: [
        {
          title: "Embroidery Challan Tracker",
          path: "/embroidery",
          emoji: "🧵",
          desc: "Create and manage embroidery production challans with real-time tracking"
        },
        {
          title: "Printing Challan Tracker",
          path: "/printing",
          emoji: "🖨️",
          desc: "Handle printing orders, screen management, color output, and production workflows"
        },
        {
          title: "EMB / PRINT REMARKS Hub",
          path: "/emb-print-remarks",
          emoji: "📝",
          desc: "View and update live remarks directly for pending Embroidery & Printing lots",
          badge: "REAL-TIME"
        },
        {
          title: "After EMB / PRINT DONE",
          path: "/pending-issue-to-stitching",
          emoji: "📋",
          desc: "Track garments pending for stitching after embroidery/printing completion"
        },
        {
          title: "Daily Printing Challan Report",
          path: "/daily-printing-challan",
          emoji: "📊",
          desc: "Daily printing output, screen utilization, and color-wise operator analytics"
        },
        {
          title: "Daily Embroidery Challan Report",
          path: "/daily-embroidery-challan",
          emoji: "📈",
          desc: "Daily embroidery progress, machine efficiency, stitch counts, and thread usage"
        }
      ]
    },
    {
      id: "packing_finishing",
      title: "Packing & Finishing Department",
      emoji: "📦",
      color: "#10b981",
      gradient: "linear-gradient(135deg, #047857 0%, #10b981 100%)",
      description: "Carton preparation, operator packing allocation, folding, and shipment dispatch",
      reports: [
        {
          title: "Issue To Packing",
          path: "/issue-to-packing",
          emoji: "📦",
          desc: "Manage packing operations, carton preparation, and shipment readiness"
        },
        {
          title: "Pending Packing to Issue",
          path: "/pending-packing-issue",
          emoji: "📦⏳",
          desc: "Track garments that completed stitching and await packing operations",
          badge: "LIVE"
        },
        {
          title: "Packing Alloted Lot",
          path: "/packing-alloted-lot",
          emoji: "📦✅",
          desc: "Manage lots allocated for packing with operator assignment and carton planning"
        },
        {
          title: "Daily Packing Report",
          path: "/daily-packing-report",
          emoji: "📦",
          desc: "Daily packing output, carton count, shipment readiness, and packing efficiency"
        },
        {
          title: "Daily Folding Report",
          path: "/daily-folding-report",
          emoji: "👕",
          desc: "Daily folding operations, style-wise output, operator efficiency, and quality checks"
        }
      ]
    },
    {
      id: "materials_inventory",
      title: "Materials, Inventory & Accessories",
      emoji: "📦",
      color: "#f59e0b",
      gradient: "linear-gradient(135deg, #b45309 0%, #f59e0b 100%)",
      description: "Raw material receipts, yarn, kora fabric, knitting, zips, stickers, and trims",
      reports: [
        {
          title: "Embroidery Material Receiving",
          path: "/embroidery-material-receiving",
          emoji: "🧵📥",
          desc: "Track embroidery raw material receipts including threads, backing, needles, and stock"
        },
        {
          title: "Printing Material Receiving",
          path: "/printing-material-receiving",
          emoji: "🖨️📥",
          desc: "Monitor printing raw material receipts including inks, screens, and chemicals"
        },
        {
          title: "Knitting Report",
          path: "/knitting-report",
          emoji: "🧶",
          desc: "Track knitting machine production, yarn conversion efficiency, and fabric roll outputs"
        },
        {
          title: "Collar Report",
          path: "/collar-report",
          emoji: "👔",
          desc: "Monitor collar and cuff knitting production, size-wise counts, and machine utilization"
        },
        {
          title: "Kora Roll Report",
          path: "/kora-roll-report",
          emoji: "📜",
          desc: "Un-dyed kora fabric rolls, weight in kgs, meters, dyeing dispatch, and stock status"
        },
        {
          title: "Yarn Stock Report",
          path: "/yarn-report",
          emoji: "🧵",
          desc: "Knitting department yarn stock summary, packages count, and balance weight in kgs"
        },
        {
          title: "Sticker Report",
          path: "/sticker-report",
          emoji: "🏷️",
          desc: "Track sticker printing, usage, size-wise breakdown, and label inventory management"
        },
        {
          title: "Pending Zip PO Report",
          path: "/pending-zip-po-report",
          emoji: "🤐⏳",
          desc: "Track production lots requiring zips where Purchase Orders (PO) have NOT been created"
        }
      ]
    },
    {
      id: "executive_milestones",
      title: "Executive Milestones & Matrix",
      emoji: "📊",
      color: "#8b5cf6",
      gradient: "linear-gradient(135deg, #4c1d95 0%, #8b5cf6 100%)",
      description: "Cross-departmental WIP stage analysis, live executive dashboard, audit logs & milestone tracking",
      reports: [
        {
          title: "Cut to Pack Report (Overall Matrix)",
          path: "/overall-cutting-to-packing-report",
          emoji: "📊",
          desc: "Complete production tracking from cutting to packing with stage analysis and bottleneck detection"
        },
        {
          title: "Short Summary Report (Live Dashboard)",
          path: "/short-summary-report",
          emoji: "⚡",
          desc: "Quick overview of key production metrics across all factory departments",
          badge: "EXECUTIVE LIVE"
        },
        {
          title: "Lot Change Logs",
          path: "/lot-logs",
          emoji: "📋",
          desc: "Log lot modifications, authorization details, broadcast notifications, and audit history"
        },
        {
          title: "Notification Center",
          path: "/notifications",
          emoji: "🔔",
          desc: "Full-screen activity feed for all factory alerts, stitching, cutting, and lot changes"
        }
      ]
    }
  ];

  return (
    <div
      style={{
        minHeight: "100vh",
        width: "100%",
        backgroundColor: "#f8fafc",
        backgroundImage: "radial-gradient(at 0% 0%, rgba(99, 102, 241, 0.06) 0px, transparent 50%), radial-gradient(at 100% 0%, rgba(236, 72, 153, 0.04) 0px, transparent 50%), radial-gradient(at 50% 100%, rgba(16, 185, 129, 0.04) 0px, transparent 50%)",
        color: "#0f172a",
        fontFamily: "'Plus Jakarta Sans', 'Inter', -apple-system, sans-serif",
        position: "relative",
        paddingBottom: "80px",
        boxSizing: "border-box"
      }}
    >
      {/* Top Banner & Header Section (Full Screen Width) */}
      <header
        style={{
          width: "100%",
          padding: "24px 36px 20px",
          boxSizing: "border-box"
        }}
      >
        {/* Signature Modern Hero Box */}
        <motion.div
          initial={{ opacity: 0, y: -15 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.5 }}
          style={{
            width: "100%",
            background: "linear-gradient(135deg, #1e1b4b 0%, #312e81 40%, #4338ca 100%)",
            borderRadius: "26px",
            padding: "36px 40px",
            boxShadow: "0 20px 45px -10px rgba(30, 27, 75, 0.35)",
            color: "#ffffff",
            position: "relative",
            overflow: "hidden",
            boxSizing: "border-box",
            border: "1px solid rgba(255, 255, 255, 0.15)"
          }}
        >
          {/* Decorative Ambient Shapes */}
          <div style={{ position: "absolute", top: "-60px", right: "-60px", width: "300px", height: "300px", borderRadius: "50%", background: "radial-gradient(circle, rgba(99, 102, 241, 0.3) 0%, transparent 70%)", pointerEvents: "none" }} />
          <div style={{ position: "absolute", bottom: "-100px", right: "200px", width: "350px", height: "350px", borderRadius: "50%", background: "radial-gradient(circle, rgba(236, 72, 153, 0.18) 0%, transparent 70%)", pointerEvents: "none" }} />
          <div style={{ position: "absolute", top: "40%", left: "-60px", width: "200px", height: "200px", borderRadius: "50%", background: "radial-gradient(circle, rgba(16, 185, 129, 0.15) 0%, transparent 70%)", pointerEvents: "none" }} />

          {/* Top Operational Info Line */}
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: "12px", marginBottom: "22px" }}>
            <div style={{ display: "flex", alignItems: "center", gap: "10px", flexWrap: "wrap" }}>
              <div
                style={{
                  display: "inline-flex",
                  alignItems: "center",
                  gap: "8px",
                  padding: "6px 16px",
                  borderRadius: "9999px",
                  background: "rgba(255, 255, 255, 0.15)",
                  backdropFilter: "blur(12px)",
                  fontSize: "0.84rem",
                  fontWeight: 700,
                  color: "#e0e7ff",
                  border: "1px solid rgba(255, 255, 255, 0.2)"
                }}
              >
                <span style={{ display: "inline-block", width: "8px", height: "8px", borderRadius: "50%", background: "#34d399", boxShadow: "0 0 10px #34d399" }} />
                <span>Garment Production Suite</span>
                <span style={{ color: "rgba(255,255,255,0.3)" }}>|</span>
                <span style={{ color: "#a5b4fc" }}>Enterprise Hub v2.5</span>
              </div>

              <div
                style={{
                  display: "inline-flex",
                  alignItems: "center",
                  gap: "6px",
                  padding: "6px 14px",
                  borderRadius: "9999px",
                  background: "rgba(239, 68, 68, 0.25)",
                  backdropFilter: "blur(12px)",
                  fontSize: "0.8rem",
                  fontWeight: 800,
                  color: "#fecdd3",
                  border: "1px solid rgba(239, 68, 68, 0.4)"
                }}
              >
                <span>🚨</span> 8 Daily Critical Checks
              </div>
            </div>

            <div style={{ display: "flex", alignItems: "center", gap: "16px", fontSize: "0.85rem", color: "#c7d2fe" }}>
              <span>📅 {new Date().toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' })}</span>
              <span style={{ display: "inline-flex", alignItems: "center", gap: "6px", color: "#34d399", fontWeight: 700 }}>
                <span style={{ animation: "pulse 2s infinite", display: "inline-block" }}>🟢</span> Shift Active
              </span>
            </div>
          </div>

          {/* Main Title & Hero KPI Grid */}
          <div style={{ display: "flex", flexDirection: "column", gap: "28px" }}>
            <div style={{ maxWidth: "800px" }}>
              <h1
                style={{
                  fontSize: "clamp(2.3rem, 3.8vw, 3.2rem)",
                  fontWeight: 900,
                  margin: 0,
                  letterSpacing: "-0.035em",
                  color: "#ffffff",
                  lineHeight: 1.15
                }}
              >
                Factory Command Center
              </h1>
              <p
                style={{
                  color: "#c7d2fe",
                  fontSize: "1.05rem",
                  marginTop: "12px",
                  marginBottom: 0,
                  lineHeight: 1.5,
                  fontWeight: 500
                }}
              >
                Real-time operational suite for cutting, embroidery, printing, stitching, packing, and high-level factory analytics.
              </p>
            </div>

            {/* Interactive KPI Quick-Filter Deck */}
            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(auto-fit, minmax(130px, 1fr))",
                gap: "12px",
                width: "100%"
              }}
            >
              <HeroStatCard
                label="All Modules"
                value={totalCount}
                icon="📊"
                color="#818cf8"
                active={activeCategory === "All" && quickTag === "ALL"}
                onClick={() => {
                  setActiveCategory("All");
                  setQuickTag("ALL");
                }}
              />
              <HeroStatCard
                label="Daily Critical"
                value={urgentModules.length}
                icon="🚨"
                color="#f43f5e"
                alert={true}
                active={quickTag === "CRITICAL"}
                onClick={() => {
                  setActiveCategory("All");
                  setQuickTag("CRITICAL");
                }}
              />
              <HeroStatCard
                label="Cutting"
                value={departmentCounts["Cutting"] || 0}
                icon="✂️"
                color="#38bdf8"
                active={activeCategory === "Cutting"}
                onClick={() => {
                  setActiveCategory("Cutting");
                  setQuickTag("ALL");
                }}
              />
              <HeroStatCard
                label="Stitching"
                value={departmentCounts["Stitching"] || 0}
                icon="🪡"
                color="#22d3ee"
                active={activeCategory === "Stitching"}
                onClick={() => {
                  setActiveCategory("Stitching");
                  setQuickTag("ALL");
                }}
              />
              <HeroStatCard
                label="Emb & Print"
                value={departmentCounts["Embroidery & Printing"] || 0}
                icon="🎨"
                color="#c084fc"
                active={activeCategory === "Embroidery & Printing"}
                onClick={() => {
                  setActiveCategory("Embroidery & Printing");
                  setQuickTag("ALL");
                }}
              />
              <HeroStatCard
                label="Packing"
                value={departmentCounts["Packing & Finishing"] || 0}
                icon="📦"
                color="#34d399"
                active={activeCategory === "Packing & Finishing"}
                onClick={() => {
                  setActiveCategory("Packing & Finishing");
                  setQuickTag("ALL");
                }}
              />
              <HeroStatCard
                label="Materials"
                value={departmentCounts["Materials & Inventory"] || 0}
                icon="🧵"
                color="#fbbf24"
                active={activeCategory === "Materials & Inventory"}
                onClick={() => {
                  setActiveCategory("Materials & Inventory");
                  setQuickTag("ALL");
                }}
              />
              <HeroStatCard
                label="Executive"
                value={departmentCounts["Executive & Analytics"] || 0}
                icon="📈"
                color="#a78bfa"
                active={activeCategory === "Executive & Analytics"}
                onClick={() => {
                  setActiveCategory("Executive & Analytics");
                  setQuickTag("ALL");
                }}
              />
            </div>
          </div>
        </motion.div>

        {/* Search, View Mode & Filter Controls Container */}
        <motion.div
          initial={{ opacity: 0, y: 15 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.5, delay: 0.2 }}
          style={{
            marginTop: "28px",
            background: "#ffffff",
            border: "1px solid #e2e8f0",
            borderRadius: "22px",
            padding: "20px 24px",
            boxShadow: "0 10px 30px rgba(0, 0, 0, 0.04)",
            display: "flex",
            flexDirection: "column",
            gap: "18px",
            width: "100%",
            boxSizing: "border-box"
          }}
        >
          {/* Top Controls Row: Search Input + View Mode Switcher */}
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: "16px" }}>
            {/* Search Input Box */}
            <div style={{ position: "relative", flex: "1 1 360px", maxWidth: "720px" }}>
              <span style={{ position: "absolute", left: "16px", top: "50%", transform: "translateY(-50%)", fontSize: "1.2rem", color: "#64748b" }}>
                🔍
              </span>
              <input
                ref={searchInputRef}
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Search by report title, department, feature, or keyword..."
                style={{
                  width: "100%",
                  padding: "14px 110px 14px 48px",
                  borderRadius: "14px",
                  background: "#f8fafc",
                  border: searchQuery ? "2px solid #4f46e5" : "1.5px solid #cbd5e1",
                  color: "#0f172a",
                  fontSize: "0.95rem",
                  fontWeight: 500,
                  outline: "none",
                  boxSizing: "border-box",
                  transition: "all 0.25s ease"
                }}
              />

              {/* Clear / Shortcut Badge */}
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
                      fontSize: "0.8rem",
                      fontWeight: 800
                    }}
                  >
                    ✕
                  </button>
                ) : (
                  <span
                    style={{
                      background: "#e2e8f0",
                      color: "#64748b",
                      padding: "3px 8px",
                      borderRadius: "6px",
                      fontSize: "0.75rem",
                      fontWeight: 700,
                      pointerEvents: "none"
                    }}
                  >
                    Ctrl+K
                  </span>
                )}
              </div>
            </div>

            {/* Quick View Mode Switcher */}
            <div style={{ display: "flex", alignItems: "center", gap: "6px", background: "#f1f5f9", padding: "4px", borderRadius: "12px", border: "1px solid #e2e8f0" }}>
              <button
                onClick={() => setViewMode("grid")}
                style={{
                  padding: "8px 18px",
                  borderRadius: "9px",
                  border: "none",
                  background: viewMode === "grid" ? "#ffffff" : "transparent",
                  color: viewMode === "grid" ? "#4f46e5" : "#64748b",
                  boxShadow: viewMode === "grid" ? "0 2px 8px rgba(0, 0, 0, 0.08)" : "none",
                  fontWeight: 700,
                  fontSize: "0.85rem",
                  cursor: "pointer",
                  display: "flex",
                  alignItems: "center",
                  gap: "6px",
                  transition: "all 0.2s ease"
                }}
              >
                <span>🏓</span> Grid
              </button>
              <button
                onClick={() => setViewMode("list")}
                style={{
                  padding: "8px 18px",
                  borderRadius: "9px",
                  border: "none",
                  background: viewMode === "list" ? "#ffffff" : "transparent",
                  color: viewMode === "list" ? "#4f46e5" : "#64748b",
                  boxShadow: viewMode === "list" ? "0 2px 8px rgba(0, 0, 0, 0.08)" : "none",
                  fontWeight: 700,
                  fontSize: "0.85rem",
                  cursor: "pointer",
                  display: "flex",
                  alignItems: "center",
                  gap: "6px",
                  transition: "all 0.2s ease"
                }}
              >
                <span>📄</span> List
              </button>
              <button
                onClick={() => setViewMode("index")}
                style={{
                  padding: "8px 18px",
                  borderRadius: "9px",
                  border: "none",
                  background: viewMode === "index" ? "linear-gradient(135deg, #4f46e5 0%, #3730a3 100%)" : "transparent",
                  color: viewMode === "index" ? "#ffffff" : "#64748b",
                  boxShadow: viewMode === "index" ? "0 4px 12px rgba(79, 70, 229, 0.25)" : "none",
                  fontWeight: 700,
                  fontSize: "0.85rem",
                  cursor: "pointer",
                  display: "flex",
                  alignItems: "center",
                  gap: "6px",
                  transition: "all 0.2s ease"
                }}
              >
                <span>📑</span> Index Panel
              </button>
            </div>
          </div>

          {/* Quick Tag Pills (All, Critical, Live, AI, Pending) */}
          <div style={{ display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap" }}>
            <span style={{ fontSize: "0.8rem", fontWeight: 700, color: "#64748b", marginRight: "4px" }}>
              Filter By:
            </span>
            {[
              { id: "ALL", label: "✨ All Reports", count: totalCount },
              { id: "CRITICAL", label: "🚨 Critical Checkpoints", count: 8 },
              { id: "LIVE", label: "⚡ Live Updates", count: accessibleModules.filter(m => m.isLive).length },
              { id: "AI", label: "🤖 AI Forecasts", count: accessibleModules.filter(m => m.isAI).length },
              { id: "PENDING", label: "⏳ Action Required", count: accessibleModules.filter(m => m.status === "pending").length }
            ].map(tag => {
              const isSelected = quickTag === tag.id;
              return (
                <button
                  key={tag.id}
                  onClick={() => setQuickTag(tag.id)}
                  style={{
                    background: isSelected ? "#0f172a" : "#f1f5f9",
                    color: isSelected ? "#ffffff" : "#334155",
                    border: isSelected ? "1px solid #0f172a" : "1px solid #e2e8f0",
                    padding: "6px 14px",
                    borderRadius: "9999px",
                    fontSize: "0.8rem",
                    fontWeight: 700,
                    cursor: "pointer",
                    display: "inline-flex",
                    alignItems: "center",
                    gap: "6px",
                    transition: "all 0.2s ease",
                    boxShadow: isSelected ? "0 2px 8px rgba(15, 23, 42, 0.2)" : "none"
                  }}
                >
                  <span>{tag.label}</span>
                  <span
                    style={{
                      background: isSelected ? "rgba(255,255,255,0.2)" : "#cbd5e1",
                      color: isSelected ? "#ffffff" : "#475569",
                      padding: "1px 6px",
                      borderRadius: "9999px",
                      fontSize: "0.7rem",
                      fontWeight: 800
                    }}
                  >
                    {tag.count}
                  </span>
                </button>
              );
            })}
          </div>

          {/* Department-Wise Filter Tabs */}
          <div style={{ display: "flex", gap: "8px", flexWrap: "wrap", borderTop: "1px solid #f1f5f9", paddingTop: "14px" }}>
            {categories.map((cat) => {
              const isSelected = activeCategory === cat.id;

              return (
                <button
                  key={cat.id}
                  onClick={() => {
                    setActiveCategory(cat.id);
                    if (quickTag !== "ALL") setQuickTag("ALL");
                  }}
                  style={{
                    background: isSelected
                      ? "linear-gradient(135deg, #4f46e5 0%, #3730a3 100%)"
                      : "#f8fafc",
                    color: isSelected ? "#ffffff" : "#475569",
                    border: isSelected ? "1px solid #4338ca" : "1px solid #e2e8f0",
                    padding: "9px 16px",
                    borderRadius: "12px",
                    fontSize: "0.86rem",
                    fontWeight: 700,
                    cursor: "pointer",
                    display: "flex",
                    alignItems: "center",
                    gap: "8px",
                    transition: "all 0.22s ease",
                    boxShadow: isSelected ? "0 6px 16px rgba(79, 70, 229, 0.25)" : "none"
                  }}
                >
                  <span>{cat.emoji}</span>
                  <span>{cat.label}</span>
                  <span
                    style={{
                      background: isSelected ? "rgba(255, 255, 255, 0.25)" : "#e2e8f0",
                      color: isSelected ? "#ffffff" : "#475569",
                      padding: "2px 7px",
                      borderRadius: "9999px",
                      fontSize: "0.74rem",
                      fontWeight: 800
                    }}
                  >
                    {cat.count}
                  </span>
                </button>
              );
            })}
          </div>
        </motion.div>
      </header>

      {/* Main Content Area (Full Screen Width) */}
      <main style={{ width: "100%", padding: "0 36px", boxSizing: "border-box" }}>
        {/* Pinned Quick Access Row (If on 'All' tab and user has pinned items) */}
        {activeCategory === "All" && searchQuery === "" && quickTag === "ALL" && pinnedModules.length > 0 && (
          <div style={{ marginBottom: "36px" }}>
            <div style={{ display: "flex", alignItems: "center", gap: "10px", marginBottom: "16px" }}>
              <span style={{ fontSize: "1.3rem" }}>⭐</span>
              <h2 style={{ fontSize: "1.3rem", fontWeight: 800, margin: 0, color: "#0f172a" }}>
                Pinned Quick Access
              </h2>
              <span style={{ fontSize: "0.85rem", color: "#64748b", fontWeight: 500 }}>(Top bookmarked modules)</span>
            </div>

            <div
              style={{
                display: "grid",
                gridTemplateColumns: viewMode === "grid" ? "repeat(auto-fill, minmax(340px, 1fr))" : "1fr",
                gap: "20px",
                width: "100%"
              }}
            >
              {modules
                .filter((m) => pinnedModules.includes(m.title))
                .map((module) => (
                  <LightModuleCard
                    key={'pinned-' + module.title}
                    module={module}
                    isPinned={true}
                    togglePin={togglePin}
                    viewMode={viewMode}
                    hoveredCard={hoveredCard}
                    setHoveredCard={setHoveredCard}
                  />
                ))}
            </div>
          </div>
        )}

        {/* INDEX PANEL VIEW (Hierarchical Department Directory & Sub-Reports Index) */}
        {viewMode === "index" ? (
          <div style={{ display: "flex", flexDirection: "column", gap: "28px", marginTop: "12px", marginBottom: "40px" }}>
            {/* Index Directory Header Box */}
            <div style={{
              background: "linear-gradient(135deg, #0f172a 0%, #1e293b 100%)",
              borderRadius: "20px",
              padding: "24px 32px",
              color: "#ffffff",
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              flexWrap: "wrap",
              gap: "16px",
              boxShadow: "0 10px 30px rgba(15, 23, 42, 0.2)",
              border: "1px solid rgba(255, 255, 255, 0.1)"
            }}>
              <div>
                <h2 style={{ margin: 0, fontSize: "1.5rem", fontWeight: 800, display: "flex", alignItems: "center", gap: "12px" }}>
                  <span>📑</span> Production Department Directory & Sub-Reports Index
                </h2>
                <p style={{ margin: "6px 0 0 0", color: "#94a3b8", fontSize: "0.92rem", fontWeight: 500, maxWidth: "700px" }}>
                  Hierarchical structure listing core production departments (Cutting, Stitching, Embroidery/Printing, Packing, Materials) and their detailed sub-reports.
                </p>
              </div>
              <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
                <span style={{ background: "rgba(255, 255, 255, 0.1)", backdropFilter: "blur(8px)", padding: "8px 16px", borderRadius: "9999px", fontSize: "0.85rem", fontWeight: 700, color: "#e2e8f0", border: "1px solid rgba(255,255,255,0.15)" }}>
                  🏢 6 Core Departments
                </span>
                <span style={{ background: "rgba(99, 102, 241, 0.25)", backdropFilter: "blur(8px)", padding: "8px 16px", borderRadius: "9999px", fontSize: "0.85rem", fontWeight: 700, color: "#c7d2fe", border: "1px solid rgba(99,102,241,0.3)" }}>
                  📊 34 Sub-Reports Listed
                </span>
              </div>
            </div>

            {/* Department Accordion Directory Cards */}
            {departmentIndex.map((dept) => {
              const deptReports = dept.reports.filter(r =>
                searchQuery === "" ||
                r.title.toLowerCase().includes(searchQuery.toLowerCase()) ||
                r.desc.toLowerCase().includes(searchQuery.toLowerCase())
              );

              if (searchQuery !== "" && deptReports.length === 0) return null;

              return (
                <div key={dept.id} style={{
                  background: "#ffffff",
                  borderRadius: "22px",
                  border: "1px solid #e2e8f0",
                  boxShadow: "0 8px 24px rgba(0,0,0,0.03)",
                  overflow: "hidden"
                }}>
                  {/* Department Banner Header */}
                  <div style={{
                    background: dept.gradient,
                    padding: "22px 30px",
                    color: "#ffffff",
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    flexWrap: "wrap",
                    gap: "14px"
                  }}>
                    <div style={{ display: "flex", alignItems: "center", gap: "16px" }}>
                      <div style={{ width: "48px", height: "48px", borderRadius: "14px", background: "rgba(255,255,255,0.18)", backdropFilter: "blur(10px)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: "1.8rem" }}>
                        {dept.emoji}
                      </div>
                      <div>
                        <h3 style={{ margin: 0, fontSize: "1.3rem", fontWeight: 800, color: "#ffffff", letterSpacing: "-0.01em" }}>
                          {dept.title}
                        </h3>
                        <p style={{ margin: "3px 0 0 0", fontSize: "0.88rem", color: "rgba(255, 255, 255, 0.88)", fontWeight: 500 }}>
                          {dept.description}
                        </p>
                      </div>
                    </div>

                    <span style={{
                      background: "rgba(255, 255, 255, 0.22)",
                      backdropFilter: "blur(10px)",
                      padding: "6px 16px",
                      borderRadius: "9999px",
                      fontSize: "0.82rem",
                      fontWeight: 800,
                      color: "#ffffff",
                      border: "1px solid rgba(255, 255, 255, 0.3)"
                    }}>
                      {deptReports.length} Sub-Reports
                    </span>
                  </div>

                  {/* Sub-Reports Grid */}
                  <div style={{
                    padding: "24px",
                    display: "grid",
                    gridTemplateColumns: "repeat(auto-fill, minmax(310px, 1fr))",
                    gap: "18px",
                    background: "#f8fafc"
                  }}>
                    {deptReports.map((report) => (
                      <div
                        key={report.title}
                        onClick={() => history.push(report.path)}
                        style={{
                          background: "#ffffff",
                          borderRadius: "16px",
                          padding: "20px 22px",
                          border: "1.5px solid #e2e8f0",
                          boxShadow: "0 2px 8px rgba(0,0,0,0.02)",
                          cursor: "pointer",
                          transition: "all 0.22s ease",
                          display: "flex",
                          flexDirection: "column",
                          justifyContent: "space-between",
                          position: "relative"
                        }}
                        onMouseEnter={(e) => {
                          e.currentTarget.style.transform = "translateY(-4px)";
                          e.currentTarget.style.borderColor = dept.color || "#4f46e5";
                          e.currentTarget.style.boxShadow = "0 12px 24px rgba(79, 70, 229, 0.12)";
                        }}
                        onMouseLeave={(e) => {
                          e.currentTarget.style.transform = "translateY(0)";
                          e.currentTarget.style.borderColor = "#e2e8f0";
                          e.currentTarget.style.boxShadow = "0 2px 8px rgba(0,0,0,0.02)";
                        }}
                      >
                        <div>
                          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: "12px" }}>
                            <span style={{ fontSize: "1.5rem" }}>{report.emoji}</span>
                            {report.badge && (
                              <span style={{
                                background: "#e0e7ff",
                                color: "#3730a3",
                                fontSize: "0.7rem",
                                fontWeight: 800,
                                padding: "4px 9px",
                                borderRadius: "8px",
                                letterSpacing: "0.5px"
                              }}>
                                {report.badge}
                              </span>
                            )}
                          </div>
                          <h4 style={{ margin: "0 0 8px 0", fontSize: "1.02rem", fontWeight: 800, color: "#0f172a", lineHeight: 1.3 }}>
                            {report.title}
                          </h4>
                          <p style={{ margin: 0, fontSize: "0.84rem", color: "#64748b", lineHeight: 1.45, fontWeight: 500 }}>
                            {report.desc}
                          </p>
                        </div>

                        <div style={{
                          marginTop: "18px",
                          paddingTop: "12px",
                          borderTop: "1px solid #f1f5f9",
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "space-between",
                          fontSize: "0.82rem",
                          fontWeight: 700,
                          color: dept.color || "#4f46e5"
                        }}>
                          <span>Open Sub-Report</span>
                          <span>→</span>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })}
          </div>
        ) : activeCategory === "All" && searchQuery === "" && quickTag === "ALL" ? (
          <>
            {/* 🚨 DEDICATED HIGHLIGHT SECTION: DAILY BASIS CRITICAL REPORTS */}
            <motion.section
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ duration: 0.4 }}
              style={{
                marginBottom: "44px",
                background: "#ffffff",
                border: "2px solid #e0e7ff",
                borderRadius: "26px",
                padding: "32px 36px",
                boxShadow: "0 10px 30px rgba(79, 70, 229, 0.06)",
                position: "relative",
                overflow: "hidden"
              }}
            >
              {/* Top Accent Gradient Line */}
              <div
                style={{
                  position: "absolute",
                  top: 0,
                  left: 0,
                  right: 0,
                  height: "5px",
                  background: "linear-gradient(90deg, #4f46e5 0%, #06b6d4 35%, #ec4899 70%, #f59e0b 100%)"
                }}
              />

              {/* Section Header */}
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "space-between",
                  flexWrap: "wrap",
                  gap: "18px",
                  marginBottom: "22px"
                }}
              >
                <div>
                  <div style={{ display: "flex", alignItems: "center", gap: "12px", flexWrap: "wrap" }}>
                    <span
                      style={{
                        display: "inline-flex",
                        alignItems: "center",
                        justifyContent: "center",
                        width: "44px",
                        height: "44px",
                        borderRadius: "14px",
                        background: "linear-gradient(135deg, #1e1b4b 0%, #4338ca 100%)",
                        color: "#ffffff",
                        fontSize: "1.4rem",
                        boxShadow: "0 4px 14px rgba(79, 70, 229, 0.3)"
                      }}
                    >
                      🚨
                    </span>
                    <h2
                      style={{
                        fontSize: "1.55rem",
                        fontWeight: 900,
                        margin: 0,
                        color: "#0f172a",
                        letterSpacing: "-0.025em"
                      }}
                    >
                      Daily Critical Factory Inspection Pipeline
                    </h2>
                    <span
                      style={{
                        background: "#fef2f2",
                        color: "#dc2626",
                        padding: "5px 14px",
                        borderRadius: "9999px",
                        fontSize: "0.78rem",
                        fontWeight: 800,
                        border: "1px solid #fca5a5",
                        display: "inline-flex",
                        alignItems: "center",
                        gap: "6px"
                      }}
                    >
                      <span
                        style={{
                          width: "7px",
                          height: "7px",
                          borderRadius: "50%",
                          background: "#ef4444",
                          display: "inline-block"
                        }}
                      />
                      MANDATORY 8-STAGE AUDIT
                    </span>
                  </div>
                  <p
                    style={{
                      margin: "8px 0 0 0",
                      color: "#64748b",
                      fontSize: "0.94rem",
                      fontWeight: 500,
                      maxWidth: "900px",
                      lineHeight: 1.5
                    }}
                  >
                    High-priority sequential monitoring across Cutting, Embroidery, Printing, Stitching, and Packing stages to prevent production bottlenecks.
                  </p>
                </div>

                <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                  <span
                    style={{
                      background: "#e0e7ff",
                      color: "#3730a3",
                      border: "1px solid #c7d2fe",
                      padding: "7px 16px",
                      borderRadius: "12px",
                      fontSize: "0.82rem",
                      fontWeight: 800
                    }}
                  >
                    ⚡ Top Priority Sequence
                  </span>
                </div>
              </div>

              {/* Interactive Factory Workflow Stepper Bar */}
              <div
                style={{
                  background: "#f8fafc",
                  border: "1px solid #e2e8f0",
                  borderRadius: "18px",
                  padding: "14px 18px",
                  marginBottom: "26px",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "space-between",
                  gap: "10px",
                  overflowX: "auto"
                }}
              >
                {urgentModules.map((m, idx) => (
                  <React.Fragment key={m.path}>
                    <div
                      onClick={() => history.push(m.path)}
                      title={m.title}
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: "8px",
                        padding: "7px 12px",
                        borderRadius: "10px",
                        background: "#ffffff",
                        border: "1px solid #cbd5e1",
                        cursor: "pointer",
                        whiteSpace: "nowrap",
                        transition: "all 0.2s ease",
                        flexShrink: 0,
                        boxShadow: "0 2px 4px rgba(0,0,0,0.02)"
                      }}
                      onMouseEnter={(e) => {
                        e.currentTarget.style.background = "#e0e7ff";
                        e.currentTarget.style.borderColor = "#6366f1";
                        e.currentTarget.style.transform = "scale(1.04)";
                      }}
                      onMouseLeave={(e) => {
                        e.currentTarget.style.background = "#ffffff";
                        e.currentTarget.style.borderColor = "#cbd5e1";
                        e.currentTarget.style.transform = "scale(1)";
                      }}
                    >
                      <span
                        style={{
                          background: "linear-gradient(135deg, #1e1b4b 0%, #4338ca 100%)",
                          color: "#ffffff",
                          fontSize: "0.7rem",
                          fontWeight: 800,
                          padding: "2px 6px",
                          borderRadius: "5px"
                        }}
                      >
                        #{idx + 1}
                      </span>
                      <span style={{ fontSize: "1.05rem" }}>{m.emoji}</span>
                      <span style={{ fontSize: "0.82rem", fontWeight: 700, color: "#1e1b4b" }}>{m.title}</span>
                    </div>
                    {idx < urgentModules.length - 1 && (
                      <span style={{ color: "#818cf8", fontSize: "0.9rem", fontWeight: 800, flexShrink: 0 }}>➔</span>
                    )}
                  </React.Fragment>
                ))}
              </div>

              {/* Priority Cards Grid */}
              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: viewMode === "grid" ? "repeat(auto-fill, minmax(330px, 1fr))" : "1fr",
                  gap: "22px",
                  width: "100%"
                }}
              >
                {urgentModules.map((module, idx) => (
                  <HighlightModuleCard
                    key={'urgent-' + module.title}
                    module={module}
                    stepNumber={idx + 1}
                    isPinned={pinnedModules.includes(module.title)}
                    togglePin={togglePin}
                    viewMode={viewMode}
                    hoveredCard={hoveredCard}
                    setHoveredCard={setHoveredCard}
                  />
                ))}
              </div>
            </motion.section>

            {/* 📁 ALL OTHER PRODUCTION & AUXILIARY REPORTS */}
            <div>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: "22px" }}>
                <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
                  <span
                    style={{
                      display: "inline-flex",
                      alignItems: "center",
                      justifyContent: "center",
                      width: "36px",
                      height: "36px",
                      borderRadius: "10px",
                      background: "linear-gradient(135deg, #4f46e5 0%, #3730a3 100%)",
                      color: "#ffffff",
                      fontSize: "1.1rem"
                    }}
                  >
                    📁
                  </span>
                  <div>
                    <h2 style={{ fontSize: "1.4rem", fontWeight: 800, margin: 0, color: "#0f172a" }}>
                      All Other Production & Departmental Reports
                    </h2>
                    <span style={{ fontSize: "0.85rem", fontWeight: 600, color: "#64748b" }}>
                      {regularModules.length} Auxiliary & Analytical Modules Available
                    </span>
                  </div>
                </div>
              </div>

              <motion.div
                layout
                style={{
                  display: "grid",
                  gridTemplateColumns: viewMode === "grid" ? "repeat(auto-fill, minmax(340px, 1fr))" : "1fr",
                  gap: "24px",
                  width: "100%"
                }}
              >
                <AnimatePresence>
                  {regularModules.map((module) => (
                    <LightModuleCard
                      key={module.title}
                      module={module}
                      isPinned={pinnedModules.includes(module.title)}
                      togglePin={togglePin}
                      viewMode={viewMode}
                      hoveredCard={hoveredCard}
                      setHoveredCard={setHoveredCard}
                    />
                  ))}
                </AnimatePresence>
              </motion.div>
            </div>
          </>
        ) : (
          <>
            {/* Dynamic Header showing active selection / search results */}
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: "20px" }}>
              <h2 style={{ fontSize: "1.35rem", fontWeight: 800, margin: 0, color: "#0f172a", display: "flex", alignItems: "center", gap: "10px" }}>
                <span>
                  {searchQuery
                    ? 'Search Results for "' + searchQuery + '"'
                    : quickTag !== "ALL"
                      ? quickTag + ' Reports'
                      : activeCategory === "All"
                        ? "All Production Modules & Analytics"
                        : activeCategory}
                </span>
                <span style={{ fontSize: "0.95rem", fontWeight: 600, color: "#64748b" }}>({filteredModules.length})</span>
              </h2>
            </div>

            {/* Zero state if no modules match search */}
            {filteredModules.length === 0 ? (
              <motion.div
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                style={{
                  textAlign: "center",
                  padding: "60px 20px",
                  background: "#ffffff",
                  borderRadius: "24px",
                  border: "1px dashed #cbd5e1"
                }}
              >
                <div style={{ fontSize: "3rem", marginBottom: "16px" }}>🔍</div>
                <h3 style={{ fontSize: "1.3rem", color: "#0f172a", margin: "0 0 8px 0" }}>No matching modules found</h3>
                <p style={{ color: "#64748b", fontSize: "0.95rem", margin: "0 0 20px 0" }}>
                  We couldn't find any module matching your filters. Try clearing your search or filters.
                </p>
                <button
                  onClick={() => {
                    setSearchQuery("");
                    setQuickTag("ALL");
                    setActiveCategory("All");
                  }}
                  style={{
                    padding: "10px 24px",
                    borderRadius: "12px",
                    background: "#4f46e5",
                    color: "#ffffff",
                    border: "none",
                    fontWeight: 600,
                    cursor: "pointer"
                  }}
                >
                  Clear All Filters
                </button>
              </motion.div>
            ) : (
              /* Modules Layout Grid or List */
              <motion.div
                layout
                style={{
                  display: "grid",
                  gridTemplateColumns: viewMode === "grid" ? "repeat(auto-fill, minmax(340px, 1fr))" : "1fr",
                  gap: "24px",
                  width: "100%"
                }}
              >
                <AnimatePresence>
                  {filteredModules.map((module) => (
                    <LightModuleCard
                      key={module.title}
                      module={module}
                      isPinned={pinnedModules.includes(module.title)}
                      togglePin={togglePin}
                      viewMode={viewMode}
                      hoveredCard={hoveredCard}
                      setHoveredCard={setHoveredCard}
                    />
                  ))}
                </AnimatePresence>
              </motion.div>
            )}
          </>
        )}

        {/* Floating Back to Top / Quick Jump Button */}
        <AnimatePresence>
          {showBackToTop && (
            <motion.button
              initial={{ opacity: 0, scale: 0.8, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.8, y: 20 }}
              onClick={() => {
                window.scrollTo({ top: 0, behavior: "smooth" });
                searchInputRef.current?.focus();
              }}
              style={{
                position: "fixed",
                bottom: "28px",
                right: "36px",
                padding: "12px 20px",
                borderRadius: "9999px",
                background: "linear-gradient(135deg, #1e1b4b 0%, #4338ca 100%)",
                color: "#ffffff",
                border: "1px solid rgba(255,255,255,0.2)",
                boxShadow: "0 12px 28px rgba(30, 27, 75, 0.35)",
                fontSize: "0.86rem",
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

        {/* Dashboard Operational Tips Section */}
        <motion.div
          initial={{ opacity: 0, y: 20 }}
          whileInView={{ opacity: 1, y: 0 }}
          viewport={{ once: true }}
          transition={{ duration: 0.5 }}
          style={{
            marginTop: "56px",
            background: "#ffffff",
            borderRadius: "24px",
            border: "1px solid #e2e8f0",
            padding: "32px 36px",
            boxShadow: "0 10px 30px rgba(0, 0, 0, 0.03)",
            width: "100%",
            boxSizing: "border-box"
          }}
        >
          <div style={{ display: "flex", alignItems: "flex-start", gap: "20px", flexWrap: "wrap" }}>
            <div
              style={{
                width: "54px",
                height: "54px",
                borderRadius: "16px",
                background: "linear-gradient(135deg, #4f46e5 0%, #3730a3 100%)",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontSize: "1.8rem",
                boxShadow: "0 10px 20px rgba(79, 70, 229, 0.25)",
                flexShrink: 0
              }}
            >
              💡
            </div>
            <div style={{ flex: 1, minWidth: "280px" }}>
              <h3 style={{ margin: "0 0 8px 0", fontSize: "1.3rem", fontWeight: 800, color: "#0f172a" }}>
                Production Dashboard Quick Tips
              </h3>
              <p style={{ color: "#64748b", fontSize: "0.95rem", margin: "0 0 20px 0" }}>
                Streamline operations across cutting, embroidery, printing, stitching, and packing departments.
              </p>

              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))", gap: "16px" }}>
                <LightTipItem
                  icon="⭐"
                  title="Bookmark Top Modules"
                  desc="Click the star icon on any card to pin your daily modules to the top of your dashboard."
                />
                <LightTipItem
                  icon="🔍"
                  title="Instant Tag Search"
                  desc="Type any keyword like 'WIP', 'Fabric', or 'Carton' in the search bar for instant filtering."
                />
                <LightTipItem
                  icon="⚡"
                  title="Real-time Reports"
                  desc="Modules tagged with 'LIVE' update metrics automatically throughout production shifts."
                />
              </div>
            </div>
          </div>
        </motion.div>
      </main>
    </div>
  );
}

// Subcomponent: Hero Stat Card inside Top Header
function HeroStatCard({ label, value, icon, color, alert, active, onClick }) {
  const [isHovered, setIsHovered] = React.useState(false);

  return (
    <div
      onClick={onClick}
      onMouseEnter={() => setIsHovered(true)}
      onMouseLeave={() => setIsHovered(false)}
      style={{
        background: active
          ? "rgba(255, 255, 255, 0.28)"
          : isHovered
            ? "rgba(255, 255, 255, 0.2)"
            : "rgba(255, 255, 255, 0.12)",
        border: active
          ? "2px solid #ffffff"
          : isHovered
            ? "1.5px solid rgba(255, 255, 255, 0.5)"
            : alert
              ? "1px solid rgba(244, 63, 94, 0.6)"
              : "1px solid rgba(255, 255, 255, 0.18)",
        borderRadius: "16px",
        padding: "14px 16px",
        backdropFilter: "blur(12px)",
        display: "flex",
        alignItems: "center",
        gap: "12px",
        cursor: onClick ? "pointer" : "default",
        transform: active || isHovered ? "scale(1.04)" : "scale(1)",
        transition: "all 0.22s ease",
        boxShadow: active
          ? "0 10px 24px rgba(0,0,0,0.25)"
          : isHovered
            ? "0 8px 20px rgba(0,0,0,0.18)"
            : "none"
      }}
    >
      <div
        style={{
          width: "42px",
          height: "42px",
          borderRadius: "12px",
          background: alert ? "linear-gradient(135deg, #f43f5e 0%, #be123c 100%)" : "rgba(255, 255, 255, 0.2)",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          fontSize: "1.3rem",
          boxShadow: alert ? "0 4px 12px rgba(244, 63, 94, 0.4)" : "none",
          flexShrink: 0
        }}
      >
        {icon}
      </div>
      <div>
        <div style={{ fontSize: "1.45rem", fontWeight: 900, color: "#ffffff", lineHeight: 1.1 }}>{value}</div>
        <div style={{ fontSize: "0.74rem", fontWeight: 700, color: alert ? "#fecdd3" : "#c7d2fe", marginTop: "4px", whiteSpace: "nowrap" }}>{label}</div>
      </div>
    </div>
  );
}

// Subcomponent: Light Module Card (Grid & List View)
function LightModuleCard({ module, isPinned, togglePin, viewMode, hoveredCard, setHoveredCard }) {
  const isHovered = hoveredCard === module.title;

  if (viewMode === "list") {
    return (
      <motion.div
        layout
        initial={{ opacity: 0, y: 10 }}
        animate={{ opacity: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.95 }}
        whileHover={{ x: 6 }}
        transition={{ duration: 0.2 }}
        onMouseEnter={() => setHoveredCard?.(module.title)}
        onMouseLeave={() => setHoveredCard?.(null)}
        style={{
          background: isHovered ? "linear-gradient(135deg, #ffffff 0%, #f8faff 100%)" : "#ffffff",
          border: isHovered ? "1.5px solid #6366f1" : "1.5px solid #e2e8f0",
          borderRadius: "18px",
          padding: "16px 24px",
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          gap: "18px",
          boxShadow: isHovered ? "0 12px 28px rgba(99, 102, 241, 0.12)" : "0 2px 8px rgba(0, 0, 0, 0.02)",
          transition: "all 0.25s ease"
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: "16px", flex: 1, minWidth: 0 }}>
          <div
            style={{
              width: "50px",
              height: "50px",
              borderRadius: "14px",
              background: module.gradient,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontSize: "1.7rem",
              boxShadow: '0 8px 18px ' + module.glowColor,
              transform: isHovered ? "scale(1.08)" : "scale(1)",
              transition: "transform 0.25s ease",
              flexShrink: 0
            }}
          >
            {module.emoji}
          </div>
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap" }}>
              <h3 style={{ margin: 0, fontSize: "1.1rem", fontWeight: 800, color: "#0f172a", letterSpacing: "-0.01em" }}>
                {module.title}
              </h3>
              {module.department && (
                <span style={{ background: "#f8fafc", color: "#475569", border: "1px solid #cbd5e1", padding: "2px 9px", borderRadius: "9999px", fontSize: "0.72rem", fontWeight: 700 }}>
                  🏢 {module.department}
                </span>
              )}
              {module.isLive && (
                <span style={{ background: "#fee2e2", color: "#dc2626", border: "1px solid #fca5a5", padding: "2px 8px", borderRadius: "9999px", fontSize: "0.7rem", fontWeight: 800 }}>
                  ⚡ LIVE
                </span>
              )}
              {module.isAI && (
                <span style={{ background: "#e0f2fe", color: "#0284c7", border: "1px solid #7dd3fc", padding: "2px 8px", borderRadius: "9999px", fontSize: "0.7rem", fontWeight: 800 }}>
                  🤖 AI
                </span>
              )}
              {module.status === "pending" && (
                <span style={{ background: "#fef3c7", color: "#d97706", border: "1px solid #fde68a", padding: "2px 8px", borderRadius: "9999px", fontSize: "0.7rem", fontWeight: 800 }}>
                  ⏳ Action
                </span>
              )}
            </div>
            <p style={{ margin: "4px 0 0 0", fontSize: "0.85rem", color: "#64748b", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
              {module.description}
            </p>
          </div>
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: "12px", flexShrink: 0 }}>
          <button
            onClick={(e) => togglePin(module.title, e)}
            title={isPinned ? "Unpin module" : "Pin module to top"}
            style={{
              background: isPinned ? "#fef3c7" : "#f1f5f9",
              border: isPinned ? "1px solid #f59e0b" : "1px solid #e2e8f0",
              color: isPinned ? "#d97706" : "#94a3b8",
              borderRadius: "10px",
              width: "36px",
              height: "36px",
              fontSize: "1.1rem",
              cursor: "pointer",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              transition: "all 0.2s ease"
            }}
          >
            ★
          </button>

          <Link
            to={module.path}
            style={{
              textDecoration: "none",
              padding: "10px 22px",
              borderRadius: "12px",
              background: "linear-gradient(135deg, #1e1b4b 0%, #4338ca 100%)",
              color: "#ffffff",
              fontSize: "0.86rem",
              fontWeight: 800,
              display: "inline-flex",
              alignItems: "center",
              gap: "6px",
              boxShadow: "0 4px 12px rgba(30, 27, 75, 0.2)"
            }}
          >
            <span>Open</span>
            <span>➔</span>
          </Link>
        </div>
      </motion.div>
    );
  }

  // Grid View Card
  return (
    <motion.div
      layout
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0, scale: 0.95 }}
      whileHover={{ y: -7, transition: { duration: 0.25 } }}
      onMouseEnter={() => setHoveredCard?.(module.title)}
      onMouseLeave={() => setHoveredCard?.(null)}
      style={{
        position: "relative",
        background: isHovered ? "linear-gradient(180deg, #ffffff 0%, #fbfcfe 100%)" : "#ffffff",
        border: isHovered ? "1.5px solid #6366f1" : "1.5px solid #e2e8f0",
        borderRadius: "24px",
        padding: "24px",
        boxShadow: isHovered
          ? '0 20px 40px -12px ' + module.glowColor + ', 0 6px 16px rgba(0,0,0,0.03)'
          : "0 4px 18px rgba(0, 0, 0, 0.03)",
        display: "flex",
        flexDirection: "column",
        justifyContent: "space-between",
        minHeight: "290px",
        transition: "all 0.3s cubic-bezier(0.16, 1, 0.3, 1)",
        overflow: "hidden"
      }}
    >
      {/* Top Accent Gradient Bar */}
      <div
        style={{
          position: "absolute",
          top: 0,
          left: 0,
          right: 0,
          height: "4px",
          background: module.gradient,
          opacity: isHovered ? 1 : 0.85
        }}
      />

      {/* Header Row: Department Badge, Status Badges & Pin Star */}
      <div>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: "16px" }}>
          <div style={{ display: "flex", alignItems: "center", gap: "6px", flexWrap: "wrap" }}>
            {module.department && (
              <span
                style={{
                  background: "#f8fafc",
                  color: "#334155",
                  border: "1px solid #cbd5e1",
                  padding: "4px 10px",
                  borderRadius: "9999px",
                  fontSize: "0.74rem",
                  fontWeight: 700
                }}
              >
                {module.department === "Cutting" && "✂️ "}
                {module.department === "Stitching" && "🪡 "}
                {module.department === "Embroidery & Printing" && "🎨 "}
                {module.department === "Packing & Finishing" && "📦 "}
                {module.department === "Materials & Inventory" && "🧵 "}
                {module.department === "Executive & Analytics" && "📊 "}
                {module.department}
              </span>
            )}

            {module.status === "pending" && (
              <span
                style={{
                  background: "#fef3c7",
                  color: "#d97706",
                  border: "1px solid #fde68a",
                  padding: "4px 10px",
                  borderRadius: "9999px",
                  fontSize: "0.74rem",
                  fontWeight: 800
                }}
              >
                ⏳ Attention
              </span>
            )}
            {module.isLive && (
              <span
                style={{
                  background: "#fee2e2",
                  color: "#dc2626",
                  border: "1px solid #fca5a5",
                  padding: "4px 10px",
                  borderRadius: "9999px",
                  fontSize: "0.74rem",
                  fontWeight: 800
                }}
              >
                ⚡ LIVE
              </span>
            )}
            {module.isAI && (
              <span
                style={{
                  background: "#e0f2fe",
                  color: "#0284c7",
                  border: "1px solid #7dd3fc",
                  padding: "4px 10px",
                  borderRadius: "9999px",
                  fontSize: "0.74rem",
                  fontWeight: 800
                }}
              >
                🤖 AI
              </span>
            )}
          </div>

          <button
            onClick={(e) => togglePin(module.title, e)}
            title={isPinned ? "Unpin from quick access" : "Pin to quick access"}
            style={{
              background: isPinned ? "#fef3c7" : "#f1f5f9",
              border: isPinned ? "1px solid #f59e0b" : "1px solid #e2e8f0",
              color: isPinned ? "#d97706" : "#94a3b8",
              borderRadius: "12px",
              width: "36px",
              height: "36px",
              fontSize: "1.1rem",
              cursor: "pointer",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              transition: "all 0.2s ease"
            }}
          >
            ★
          </button>
        </div>

        {/* Title & Emoji Header */}
        <div style={{ display: "flex", alignItems: "flex-start", gap: "16px", marginBottom: "14px" }}>
          <div
            style={{
              width: "56px",
              height: "56px",
              borderRadius: "18px",
              background: module.gradient,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontSize: "2.1rem",
              boxShadow: '0 10px 22px ' + module.glowColor,
              transform: isHovered ? "scale(1.1) rotate(4deg)" : "scale(1)",
              transition: "transform 0.3s cubic-bezier(0.34, 1.56, 0.64, 1)",
              flexShrink: 0
            }}
          >
            {module.emoji}
          </div>
          <div style={{ flex: 1 }}>
            <h3
              style={{
                margin: "0 0 4px 0",
                fontSize: "1.18rem",
                fontWeight: 800,
                color: "#0f172a",
                lineHeight: 1.3,
                letterSpacing: "-0.015em"
              }}
            >
              {module.title}
            </h3>
            <span style={{ fontSize: "0.75rem", color: "#64748b", fontWeight: 600 }}>🕒 {module.lastUpdated}</span>
          </div>
        </div>

        {/* Module Description */}
        <p
          style={{
            margin: "0 0 16px 0",
            fontSize: "0.89rem",
            color: "#475569",
            lineHeight: 1.5,
            display: "-webkit-box",
            WebkitLineClamp: 2,
            WebkitBoxOrient: "vertical",
            overflow: "hidden"
          }}
        >
          {module.description}
        </p>

        {/* Feature Tags */}
        <div style={{ display: "flex", gap: "6px", flexWrap: "wrap", marginBottom: "20px" }}>
          {module.features.map((feat, i) => (
            <span
              key={i}
              style={{
                background: "#f1f5f9",
                color: "#334155",
                padding: "3px 10px",
                borderRadius: "8px",
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

      {/* Card Action Link */}
      <Link
        to={module.path}
        style={{
          textDecoration: "none",
          padding: "12px 18px",
          borderRadius: "14px",
          background: isHovered
            ? "linear-gradient(135deg, #1e1b4b 0%, #4338ca 100%)"
            : "#f8fafc",
          color: isHovered ? "#ffffff" : "#1e1b4b",
          fontSize: "0.88rem",
          fontWeight: 800,
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          border: isHovered ? "1px solid #1e1b4b" : "1px solid #e2e8f0",
          transition: "all 0.3s ease",
          boxShadow: isHovered ? "0 8px 20px rgba(30, 27, 75, 0.25)" : "none"
        }}
      >
        <span>Open Module</span>
        <span
          style={{
            transform: isHovered ? "translateX(4px)" : "translateX(0)",
            transition: "transform 0.3s ease"
          }}
        >
          ➔
        </span>
      </Link>
    </motion.div>
  );
}

// Subcomponent: Highlight Module Card for Daily Critical Section
function HighlightModuleCard({ module, stepNumber, isPinned, togglePin, viewMode, hoveredCard, setHoveredCard }) {
  const cardKey = 'urgent-' + module.title;
  const isHovered = hoveredCard === cardKey;

  if (viewMode === "list") {
    return (
      <motion.div
        layout
        whileHover={{ x: 6 }}
        transition={{ duration: 0.2 }}
        onMouseEnter={() => setHoveredCard?.(cardKey)}
        onMouseLeave={() => setHoveredCard?.(null)}
        style={{
          background: isHovered ? "linear-gradient(135deg, #ffffff 0%, #fef2f2 100%)" : "#ffffff",
          border: isHovered ? "2px solid #e11d48" : "1.5px solid #e2e8f0",
          borderRadius: "18px",
          padding: "18px 24px",
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          gap: "18px",
          boxShadow: isHovered
            ? "0 12px 28px rgba(225, 29, 72, 0.14)"
            : "0 2px 8px rgba(0, 0, 0, 0.02)",
          transition: "all 0.25s ease"
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: "16px", flex: 1, minWidth: 0 }}>
          <div
            style={{
              background: "linear-gradient(135deg, #1e1b4b 0%, #4338ca 100%)",
              color: "#ffffff",
              fontSize: "0.78rem",
              fontWeight: 900,
              padding: "5px 12px",
              borderRadius: "8px",
              boxShadow: "0 2px 8px rgba(30, 27, 75, 0.25)",
              whiteSpace: "nowrap",
              letterSpacing: "0.5px"
            }}
          >
            STEP #{String(stepNumber).padStart(2, '0')}
          </div>
          <div
            style={{
              width: "50px",
              height: "50px",
              borderRadius: "14px",
              background: module.gradient,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontSize: "1.7rem",
              boxShadow: '0 8px 18px ' + module.glowColor,
              transform: isHovered ? "scale(1.08)" : "scale(1)",
              transition: "transform 0.25s ease",
              flexShrink: 0
            }}
          >
            {module.emoji}
          </div>
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap" }}>
              <h3 style={{ margin: 0, fontSize: "1.12rem", fontWeight: 800, color: "#0f172a", letterSpacing: "-0.01em" }}>
                {module.title}
              </h3>
              <span style={{ background: "#fef2f2", color: "#dc2626", border: "1px solid #fca5a5", padding: "2px 10px", borderRadius: "9999px", fontSize: "0.72rem", fontWeight: 800 }}>
                🚨 DAILY CRITICAL
              </span>
            </div>
            <p style={{ margin: "4px 0 0 0", fontSize: "0.86rem", color: "#64748b", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
              {module.description}
            </p>
          </div>
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: "12px", flexShrink: 0 }}>
          <button
            onClick={(e) => togglePin(module.title, e)}
            title={isPinned ? "Unpin module" : "Pin module"}
            style={{
              background: isPinned ? "#fef3c7" : "#f1f5f9",
              border: isPinned ? "1px solid #f59e0b" : "1px solid #e2e8f0",
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
              padding: "10px 22px",
              borderRadius: "12px",
              background: "linear-gradient(135deg, #e11d48 0%, #be123c 100%)",
              color: "#ffffff",
              fontSize: "0.88rem",
              fontWeight: 800,
              display: "inline-flex",
              alignItems: "center",
              gap: "8px",
              boxShadow: "0 4px 14px rgba(225, 29, 72, 0.25)"
            }}
          >
            <span>Open Check</span>
            <span>➔</span>
          </Link>
        </div>
      </motion.div>
    );
  }

  return (
    <motion.div
      layout
      whileHover={{ y: -7, transition: { duration: 0.25 } }}
      onMouseEnter={() => setHoveredCard?.(cardKey)}
      onMouseLeave={() => setHoveredCard?.(null)}
      style={{
        position: "relative",
        background: isHovered ? "linear-gradient(180deg, #ffffff 0%, #fffbfa 100%)" : "#ffffff",
        border: isHovered ? "2px solid #e11d48" : "1.5px solid #e2e8f0",
        borderRadius: "24px",
        padding: "24px",
        boxShadow: isHovered
          ? '0 20px 42px -10px ' + module.glowColor + ', 0 6px 16px rgba(0,0,0,0.03)'
          : "0 4px 18px rgba(0, 0, 0, 0.03)",
        display: "flex",
        flexDirection: "column",
        justifyContent: "space-between",
        minHeight: "290px",
        transition: "all 0.3s cubic-bezier(0.16, 1, 0.3, 1)",
        overflow: "hidden"
      }}
    >
      {/* Top Accent Gradient Bar */}
      <div
        style={{
          position: "absolute",
          top: 0,
          left: 0,
          right: 0,
          height: "4px",
          background: module.gradient
        }}
      />

      <div>
        {/* Header Badges Row */}
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: "16px" }}>
          <div style={{ display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap" }}>
            <span
              style={{
                background: "linear-gradient(135deg, #1e1b4b 0%, #4338ca 100%)",
                color: "#ffffff",
                fontSize: "0.76rem",
                fontWeight: 900,
                padding: "4px 10px",
                borderRadius: "8px",
                boxShadow: "0 2px 6px rgba(30, 27, 75, 0.25)",
                letterSpacing: "0.5px"
              }}
            >
              STEP #{String(stepNumber).padStart(2, '0')}
            </span>
            <span
              style={{
                background: "#fef2f2",
                color: "#dc2626",
                border: "1px solid #fca5a5",
                padding: "3px 10px",
                borderRadius: "9999px",
                fontSize: "0.72rem",
                fontWeight: 800
              }}
            >
              🔥 CRITICAL AUDIT
            </span>
          </div>

          <button
            onClick={(e) => togglePin(module.title, e)}
            title={isPinned ? "Unpin module" : "Pin module"}
            style={{
              background: isPinned ? "#fef3c7" : "#f8fafc",
              border: isPinned ? "1px solid #f59e0b" : "1px solid #e2e8f0",
              color: isPinned ? "#d97706" : "#94a3b8",
              borderRadius: "10px",
              width: "34px",
              height: "34px",
              fontSize: "1rem",
              cursor: "pointer",
              display: "flex",
              alignItems: "center",
              justifyContent: "center"
            }}
          >
            ★
          </button>
        </div>

        {/* Title & Emoji */}
        <div style={{ display: "flex", alignItems: "flex-start", gap: "16px", marginBottom: "14px" }}>
          <div
            style={{
              width: "56px",
              height: "56px",
              borderRadius: "18px",
              background: module.gradient,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontSize: "2.1rem",
              boxShadow: '0 10px 22px ' + module.glowColor,
              transform: isHovered ? "scale(1.1) rotate(4deg)" : "scale(1)",
              transition: "transform 0.3s cubic-bezier(0.34, 1.56, 0.64, 1)",
              flexShrink: 0
            }}
          >
            {module.emoji}
          </div>
          <div style={{ flex: 1 }}>
            <h3
              style={{
                margin: "0 0 4px 0",
                fontSize: "1.18rem",
                fontWeight: 900,
                color: "#0f172a",
                lineHeight: 1.3,
                letterSpacing: "-0.015em"
              }}
            >
              {module.title}
            </h3>
            <span style={{ fontSize: "0.76rem", color: "#4f46e5", fontWeight: 700 }}>
              🏢 {module.department} • 🕒 {module.lastUpdated}
            </span>
          </div>
        </div>

        {/* Description */}
        <p
          style={{
            margin: "0 0 16px 0",
            fontSize: "0.89rem",
            color: "#475569",
            lineHeight: 1.5,
            display: "-webkit-box",
            WebkitLineClamp: 2,
            WebkitBoxOrient: "vertical",
            overflow: "hidden"
          }}
        >
          {module.description}
        </p>

        {/* Feature Tags */}
        <div style={{ display: "flex", gap: "6px", flexWrap: "wrap", marginBottom: "20px" }}>
          {module.features.map((feat, i) => (
            <span
              key={i}
              style={{
                background: "#f1f5f9",
                color: "#334155",
                padding: "3px 10px",
                borderRadius: "8px",
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

      {/* Action Button */}
      <Link
        to={module.path}
        style={{
          textDecoration: "none",
          padding: "12px 18px",
          borderRadius: "14px",
          background: isHovered
            ? "linear-gradient(135deg, #e11d48 0%, #be123c 100%)"
            : "#fff1f2",
          color: isHovered ? "#ffffff" : "#be123c",
          fontSize: "0.88rem",
          fontWeight: 800,
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          border: isHovered ? "1px solid #be123c" : "1px solid #fecdd3",
          transition: "all 0.3s ease",
          boxShadow: isHovered ? "0 8px 20px rgba(225, 29, 72, 0.25)" : "none"
        }}
      >
        <span>Open Daily Check</span>
        <span
          style={{
            transform: isHovered ? "translateX(4px)" : "translateX(0)",
            transition: "transform 0.3s ease"
          }}
        >
          ➔
        </span>
      </Link>
    </motion.div>
  );
}

// Subcomponent: Tip Item for Light Theme
function LightTipItem({ icon, title, desc }) {
  return (
    <div
      style={{
        background: "#f8fafc",
        border: "1px solid #e2e8f0",
        borderRadius: "14px",
        padding: "16px"
      }}
    >
      <div style={{ fontSize: "1.2rem", marginBottom: "6px" }}>{icon}</div>
      <div style={{ fontWeight: 700, color: "#0f172a", fontSize: "0.95rem", marginBottom: "4px" }}>{title}</div>
      <div style={{ color: "#64748b", fontSize: "0.85rem", lineHeight: 1.4 }}>{desc}</div>
    </div>
  );
}

