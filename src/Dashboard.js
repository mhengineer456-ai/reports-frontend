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
  const [viewMode, setViewMode] = useState("grid"); // 'grid' | 'list'
  const [pinnedModules, setPinnedModules] = useState(() => {
    try {
      const saved = localStorage.getItem("pinned_dashboard_modules");
      return saved ? JSON.parse(saved) : ["Short Summary Report", "Stitching Issue", "Cut to Pack Report"];
    } catch {
      return ["Short Summary Report", "Stitching Issue", "Cut to Pack Report"];
    }
  });
  const [hoveredCard, setHoveredCard] = useState(null);

  // Notification System State
  const [notificationsOpen, setNotificationsOpen] = useState(false);
  const [notifications, setNotifications] = useState([
    { id: 1, title: "Hold Lot Alert", desc: "3 lots marked on hold in Packaging Section", time: "10m ago", unread: true, type: "warning" },
    { id: 2, title: "Stitching Output", desc: "Daily Stitching Report updated with 1,200 pcs", time: "1h ago", unread: true, type: "info" },
    { id: 3, title: "Fabric Roll Prediction", desc: "AI cutting forecast completed for Lot #4050", time: "2h ago", unread: false, type: "success" }
  ]);

  const notifRef = React.useRef(null);

  useEffect(() => {
    const handleClickOutside = (e) => {
      if (notifRef.current && !notifRef.current.contains(e.target)) {
        setNotificationsOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  const unreadCount = notifications.filter(n => n.unread).length;

  const markAllRead = () => {
    setNotifications(prev => prev.map(n => ({ ...n, unread: false })));
  };

  const removeNotification = (id, e) => {
    e.stopPropagation();
    setNotifications(prev => prev.filter(n => n.id !== id));
  };

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
      title: "Embroidery Challan",
      emoji: "🧵",
      path: "/embroidery",
      gradient: "linear-gradient(135deg, #6366f1 0%, #a855f7 100%)",
      glowColor: "rgba(99, 102, 241, 0.4)",
      description: "Create and manage embroidery production challans with real-time tracking",
      category: "Production",
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
      status: "active",
      features: ["Screen tracking", "Color wise", "Print quality"],
      lastUpdated: "Today"
    },
    {
      title: "EMB/PRINT REMARKS",
      emoji: "📝",
      path: "/emb-print-remarks",
      gradient: "linear-gradient(135deg, #8b5cf6 0%, #d946ef 100%)",
      glowColor: "rgba(139, 92, 246, 0.4)",
      description: "View and update remarks for pending Embroidery & Printing lots",
      category: "Reports",
      status: "reports",
      features: ["Embroidery Pending", "Printing Pending", "Remarks Tracking"],
      lastUpdated: "Real-time"
    },
    {
      title: "After EMB/PRINT DONE",
      emoji: "📋",
      path: "/pending-issue-to-stitching",
      gradient: "linear-gradient(135deg, #ef4444 0%, #f97316 100%)",
      glowColor: "rgba(239, 68, 68, 0.4)",
      description: "Track garments pending for stitching after embroidery/printing completion",
      category: "Production",
      status: "pending",
      features: ["Ready for stitch", "Batch transfer", "Quality pending"],
      lastUpdated: "2 hrs ago"
    },
    {
      title: "Stitching Issue",
      emoji: "🧶",
      path: "/stitching",
      gradient: "linear-gradient(135deg, #06b6d4 0%, #3b82f6 100%)",
      glowColor: "rgba(6, 182, 212, 0.4)",
      description: "Manage daily stitching operations, line assignments, and production tracking",
      category: "Production",
      status: "active",
      features: ["Line wise", "Operator tracking", "Hourly target"],
      lastUpdated: "Today"
    },
    {
      title: "After Cutting Done",
      emoji: "⏳",
      path: "/pending-stitching",
      gradient: "linear-gradient(135deg, #f59e0b 0%, #d97706 100%)",
      glowColor: "rgba(245, 158, 11, 0.4)",
      description: "Monitor pending stitching operations and delayed batches after cutting",
      category: "Production",
      status: "pending",
      features: ["Delay analysis", "Priority queue", "Resource allocation"],
      lastUpdated: "5 hrs ago"
    },
    {
      title: "Issue To Packing",
      emoji: "📦",
      path: "/issue-to-packing",
      gradient: "linear-gradient(135deg, #10b981 0%, #0d9488 100%)",
      glowColor: "rgba(16, 185, 129, 0.4)",
      description: "Manage packing operations, carton preparation, and shipment readiness",
      category: "Production",
      status: "active",
      features: ["Carton tracking", "Quality check", "Shipment prep"],
      lastUpdated: "Today"
    },
    {
      title: "Pending Packing to Issue",
      emoji: "📦⏳",
      path: "/pending-packing-issue",
      gradient: "linear-gradient(135deg, #f97316 0%, #ea580c 100%)",
      glowColor: "rgba(249, 115, 22, 0.4)",
      description: "Track garments that completed stitching and await packing operations",
      category: "Production",
      status: "pending",
      features: ["Stitching done", "Ready for pack", "Quality pending", "Batch transfer"],
      lastUpdated: "Live"
    },
    {
      title: "Packing Alloted Lot",
      emoji: "📦✅",
      path: "/packing-alloted-lot",
      gradient: "linear-gradient(135deg, #0284c7 0%, #2563eb 100%)",
      glowColor: "rgba(2, 132, 199, 0.4)",
      description: "Manage lots allocated for packing with operator assignment and carton planning",
      category: "Production",
      status: "active",
      features: ["Lot allocation", "Carton planning", "Operator assign", "Target tracking"],
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
      status: "active",
      features: ["Lot summary", "Efficiency calc", "Defect analysis"],
      lastUpdated: "Yesterday"
    },
    {
      title: "Daily Printing Challan",
      emoji: "📊",
      path: "/daily-printing-challan",
      gradient: "linear-gradient(135deg, #ff7e5f 0%, #feb47b 100%)",
      glowColor: "rgba(255, 126, 95, 0.4)",
      description: "Daily printing reports, screen usage, and production analytics",
      category: "Reports",
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
      status: "reports",
      features: ["Machine wise", "Stitch count", "Thread usage"],
      lastUpdated: "Daily"
    },
    {
      title: "Cutting Report",
      emoji: "✂️",
      path: "/cutting-report",
      gradient: "linear-gradient(135deg, #11998e 0%, #38ef7d 100%)",
      glowColor: "rgba(17, 153, 142, 0.4)",
      description: "Cutting operations report with fabric consumption and layer details",
      category: "Reports",
      status: "reports",
      features: ["Fabric usage", "Layer summary", "Marker efficiency"],
      lastUpdated: "Yesterday"
    },
    {
      title: "Daily Fabric Issue Analytics",
      emoji: "🧵",
      path: "/daily-fabric-issue-report",
      gradient: "linear-gradient(135deg, #2563eb 0%, #3b82f6 100%)",
      glowColor: "rgba(37, 99, 235, 0.4)",
      description: "Visual distribution matrix mapping overall volume and weights across tables & fabric styles",
      category: "Reports",
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
      status: "reports",
      isLive: true,
      features: ["Live dashboard", "KPI tracking", "Alerts & notifications"],
      lastUpdated: "Real-time"
    },
    // {
    //   title: "Daily Stitching Report",
    //   emoji: "🪡",
    //   path: "/daily-stitching-report",
    //   gradient: "linear-gradient(135deg, #0ea5e9 0%, #0369a1 100%)",
    //   glowColor: "rgba(14, 165, 233, 0.4)",
    //   description: "Daily stitching production, line performance, and quality metrics",
    //   category: "Reports",
    //   status: "reports",
    //   features: ["Line output", "SMV achieved", "Defect rate"],
    //   lastUpdated: "Daily"
    // },
    {
      title: "Sticker Report",
      emoji: "🏷️",
      path: "/sticker-report",
      gradient: "linear-gradient(135deg, #ff512f 0%, #f09819 100%)",
      glowColor: "rgba(255, 81, 47, 0.4)",
      description: "Track sticker printing, usage, and inventory management for labels and tags",
      category: "Materials & Inventory",
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
      status: "active",
      features: ["Ink inventory", "Screen stock", "Batch tracking"],
      lastUpdated: "Today"
    },
    {
      title: "Lot Change Logs",
      emoji: "📋",
      path: "/lot-logs",
      gradient: "linear-gradient(135deg, #f59e0b 0%, #d97706 100%)",
      glowColor: "rgba(245, 158, 11, 0.4)",
      description: "Log and track lot number changes, modifications, authorizer details, and audit history",
      category: "Reports",
      status: "active",
      features: ["Lot audit logs", "Real-time broadcast", "Permission tracking", "Excel & PDF export"],
      lastUpdated: "Real-time"
    },
    {
      title: "Notification Center",
      emoji: "🔔",
      path: "/notifications",
      gradient: "linear-gradient(135deg, #ec4899 0%, #8b5cf6 100%)",
      glowColor: "rgba(236, 72, 153, 0.4)",
      description: "Full-screen Instagram-style activity feed for all factory alerts, stitching, cutting, and lot changes",
      category: "Reports",
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
      status: "reports",
      features: ["Collar & cuff count", "Size-wise breakdown", "Flat knitting gauge", "Defect tracking"],
      lastUpdated: "Daily"
    },
    {
      title: "Kora Roll Report",
      emoji: "📜",
      path: "/kora-roll-report",
      gradient: "linear-gradient(135deg, #10b981 0%, #059669 100%)",
      glowColor: "rgba(16, 185, 129, 0.4)",
      description: "Comprehensive tracking for un-dyed kora fabric rolls, weight in kgs, meters, and stock status",
      category: "Materials & Inventory",
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
      status: "reports",
      features: ["Zip Required", "PO Pending", "Lot-wise tracking", "Export PDF/CSV"],
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

  // Filter modules based on search query and category
  const filteredModules = useMemo(() => {
    return accessibleModules.filter((m) => {
      const matchesSearch =
        searchQuery.trim() === "" ||
        m.title.toLowerCase().includes(searchQuery.toLowerCase()) ||
        m.description.toLowerCase().includes(searchQuery.toLowerCase()) ||
        m.features.some((f) => f.toLowerCase().includes(searchQuery.toLowerCase()));

      if (!matchesSearch) return false;

      if (activeCategory === "All") return true;
      if (activeCategory === "Favorites") return pinnedModules.includes(m.title);
      return m.category === activeCategory;
    });
  }, [accessibleModules, searchQuery, activeCategory, pinnedModules]);

  // Quick Stats
  const totalCount = accessibleModules.length;
  const prodCount = accessibleModules.filter((m) => m.category === "Production").length;
  const reportCount = accessibleModules.filter((m) => m.category === "Reports").length;
  const pendingCount = accessibleModules.filter((m) => m.status === "pending").length;

  const categories = ["All", "Favorites", "Production", "Reports", "Materials & Inventory"];

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
        backgroundImage: "radial-gradient(at 0% 0%, rgba(99, 102, 241, 0.08) 0px, transparent 50%), radial-gradient(at 100% 0%, rgba(236, 72, 153, 0.06) 0px, transparent 50%), radial-gradient(at 50% 100%, rgba(16, 185, 129, 0.06) 0px, transparent 50%)",
        color: "#0f172a",
        fontFamily: "'Plus Jakarta Sans', 'Inter', -apple-system, sans-serif",
        position: "relative",
        paddingBottom: "80px",
        boxSizing: "border-box"
      }}
    >
      {/* Top Navigation Bar with User Profile, Notifications & Quick Nav */}


      {/* Top Banner & Header Section (Full Screen Width) */}
      <header
        style={{
          width: "100%",
          padding: "24px 36px 20px",
          boxSizing: "border-box"
        }}
      >
        {/* Modern Top Hero Box */}
        <motion.div
          initial={{ opacity: 0, y: -15 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.5 }}
          style={{
            width: "100%",
            background: "linear-gradient(135deg, #1e1b4b 0%, #312e81 40%, #4338ca 100%)",
            borderRadius: "24px",
            padding: "36px 40px",
            boxShadow: "0 20px 40px -10px rgba(49, 46, 129, 0.3)",
            color: "#ffffff",
            position: "relative",
            overflow: "hidden",
            boxSizing: "border-box"
          }}
        >
          {/* Decorative Ambient Shapes */}
          <div style={{ position: "absolute", top: "-50px", right: "-50px", width: "250px", height: "250px", borderRadius: "50%", background: "rgba(255, 255, 255, 0.08)", pointerEvents: "none" }} />
          <div style={{ position: "absolute", bottom: "-80px", right: "150px", width: "300px", height: "300px", borderRadius: "50%", background: "rgba(99, 102, 241, 0.2)", pointerEvents: "none" }} />

          {/* Top Operational Info Line */}
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: "12px", marginBottom: "20px" }}>
            <div
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: "10px",
                padding: "6px 16px",
                borderRadius: "9999px",
                background: "rgba(255, 255, 255, 0.15)",
                backdropFilter: "blur(12px)",
                fontSize: "0.85rem",
                fontWeight: 600,
                color: "#e0e7ff",
                border: "1px solid rgba(255, 255, 255, 0.2)"
              }}
            >
              <span style={{ display: "inline-block", width: "8px", height: "8px", borderRadius: "50%", background: "#34d399", boxShadow: "0 0 10px #34d399" }} />
              <span>Garment Production Suite</span>
              <span style={{ color: "rgba(255,255,255,0.4)" }}>|</span>
              <span style={{ color: "#c7d2fe" }}>v2.5 Enterprise</span>
            </div>

            <div style={{ display: "flex", alignItems: "center", gap: "16px", fontSize: "0.85rem", color: "#c7d2fe" }}>
              <span>📅 {new Date().toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' })}</span>
              <span style={{ display: "inline-flex", alignItems: "center", gap: "6px", color: "#34d399", fontWeight: 600 }}>
                <span style={{ animation: "pulse 2s infinite", display: "inline-block" }}>🟢</span> System Active
              </span>
            </div>
          </div>

          {/* Main Title & KPI Cards Row */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(320px, 1fr))", gap: "32px", alignItems: "center" }}>
            <div>
              <h1
                style={{
                  fontSize: "clamp(2.2rem, 3.5vw, 3rem)",
                  fontWeight: 800,
                  margin: 0,
                  letterSpacing: "-0.03em",
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
                  marginTop: "10px",
                  marginBottom: 0,
                  maxWidth: "650px",
                  lineHeight: 1.5
                }}
              >
                Comprehensive tracking for cutting, embroidery, printing, stitching, packing, and analytical reporting.
              </p>
            </div>

            {/* Header KPI Stats Cards */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(130px, 1fr))", gap: "14px" }}>
              <HeroStatCard label="Total Modules" value={totalCount} icon="📦" color="#818cf8" />
              <HeroStatCard label="Production" value={prodCount} icon="🏭" color="#38bdf8" />
              <HeroStatCard label="Analytics" value={reportCount} icon="📊" color="#34d399" />
              <HeroStatCard label="Attention" value={pendingCount} icon="⏳" color="#fbbf24" alert={pendingCount > 0} />
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
            borderRadius: "20px",
            padding: "20px 24px",
            boxShadow: "0 10px 30px rgba(0, 0, 0, 0.04)",
            display: "flex",
            flexDirection: "column",
            gap: "20px",
            width: "100%",
            boxSizing: "border-box"
          }}
        >
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: "16px" }}>
            {/* Search Input Box */}
            <div style={{ position: "relative", flex: "1 1 340px", maxWidth: "700px" }}>
              <span style={{ position: "absolute", left: "16px", top: "50%", transform: "translateY(-50%)", fontSize: "1.2rem", color: "#64748b" }}>
                🔍
              </span>
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Search modules by title, feature, or description..."
                style={{
                  width: "100%",
                  padding: "14px 44px 14px 48px",
                  borderRadius: "14px",
                  background: "#f8fafc",
                  border: searchQuery ? "2px solid #4f46e5" : "1px solid #cbd5e1",
                  color: "#0f172a",
                  fontSize: "0.95rem",
                  fontWeight: 500,
                  outline: "none",
                  boxSizing: "border-box",
                  transition: "all 0.25s ease"
                }}
              />
              {searchQuery && (
                <button
                  onClick={() => setSearchQuery("")}
                  style={{
                    position: "absolute",
                    right: "14px",
                    top: "50%",
                    transform: "translateY(-50%)",
                    background: "#e2e8f0",
                    border: "none",
                    color: "#475569",
                    borderRadius: "50%",
                    width: "24px",
                    height: "24px",
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    cursor: "pointer"
                  }}
                >
                  ✕
                </button>
              )}
            </div>

            {/* View Mode Toggle Switcher */}
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

          {/* Category Tabs */}
          <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", borderTop: "1px solid #f1f5f9", paddingTop: "16px" }}>
            {categories.map((cat) => {
              const isSelected = activeCategory === cat;
              let count = modules.length;
              if (cat === "Favorites") count = pinnedModules.length;
              else if (cat !== "All") count = modules.filter((m) => m.category === cat).length;

              return (
                <button
                  key={cat}
                  onClick={() => setActiveCategory(cat)}
                  style={{
                    background: isSelected
                      ? "linear-gradient(135deg, #4f46e5 0%, #3730a3 100%)"
                      : "#f8fafc",
                    color: isSelected ? "#ffffff" : "#475569",
                    border: isSelected ? "1px solid #4338ca" : "1px solid #e2e8f0",
                    padding: "10px 20px",
                    borderRadius: "12px",
                    fontSize: "0.9rem",
                    fontWeight: 600,
                    cursor: "pointer",
                    display: "flex",
                    alignItems: "center",
                    gap: "8px",
                    transition: "all 0.25s ease",
                    boxShadow: isSelected ? "0 6px 16px rgba(79, 70, 229, 0.25)" : "none"
                  }}
                >
                  {cat === "All" && "📊"}
                  {cat === "Favorites" && "⭐"}
                  {cat === "Production" && "🏭"}
                  {cat === "Reports" && "📈"}
                  {cat === "Materials & Inventory" && "📦"}
                  <span>{cat}</span>
                  <span
                    style={{
                      background: isSelected ? "rgba(255, 255, 255, 0.25)" : "#e2e8f0",
                      color: isSelected ? "#ffffff" : "#475569",
                      padding: "2px 8px",
                      borderRadius: "9999px",
                      fontSize: "0.75rem",
                      fontWeight: 700
                    }}
                  >
                    {count}
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
        {activeCategory === "All" && searchQuery === "" && pinnedModules.length > 0 && (
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
        ) : (
          <>
            {/* Dynamic Header showing active selection / search results */}
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: "20px" }}>
              <h2 style={{ fontSize: "1.35rem", fontWeight: 800, margin: 0, color: "#0f172a", display: "flex", alignItems: "center", gap: "10px" }}>
                <span>
                  {searchQuery
                    ? 'Search Results for "' + searchQuery + '"'
                    : activeCategory === "All"
                      ? "All Production Modules & Analytics"
                      : activeCategory}
                </span>
                <span style={{ fontSize: "0.95rem", fontWeight: 600, color: "#64748b" }}>({filteredModules.length})</span>
              </h2>
            </div>
          </>
        )}

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
              We couldn't find any module matching "{searchQuery}". Try clearing your search query.
            </p>
            <button
              onClick={() => setSearchQuery("")}
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
              Clear Search Filter
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
function HeroStatCard({ label, value, icon, color, alert }) {
  return (
    <div
      style={{
        background: "rgba(255, 255, 255, 0.12)",
        border: alert ? "1px solid rgba(251, 191, 36, 0.6)" : "1px solid rgba(255, 255, 255, 0.18)",
        borderRadius: "16px",
        padding: "16px",
        backdropFilter: "blur(12px)",
        display: "flex",
        alignItems: "center",
        gap: "14px"
      }}
    >
      <div
        style={{
          width: "44px",
          height: "44px",
          borderRadius: "12px",
          background: "rgba(255, 255, 255, 0.2)",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          fontSize: "1.3rem",
          flexShrink: 0
        }}
      >
        {icon}
      </div>
      <div>
        <div style={{ fontSize: "1.5rem", fontWeight: 800, color: "#ffffff", lineHeight: 1.1 }}>{value}</div>
        <div style={{ fontSize: "0.75rem", fontWeight: 600, color: "#c7d2fe", marginTop: "4px" }}>{label}</div>
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
        whileHover={{ x: 4 }}
        transition={{ duration: 0.2 }}
        onMouseEnter={() => setHoveredCard?.(module.title)}
        onMouseLeave={() => setHoveredCard?.(null)}
        style={{
          background: "#ffffff",
          border: isHovered ? "1px solid #818cf8" : "1px solid #e2e8f0",
          borderRadius: "16px",
          padding: "16px 24px",
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          gap: "16px",
          boxShadow: isHovered ? "0 10px 25px rgba(79, 70, 229, 0.12)" : "0 2px 8px rgba(0, 0, 0, 0.02)",
          transition: "all 0.25s ease"
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: "16px", flex: 1, minWidth: 0 }}>
          <div
            style={{
              width: "48px",
              height: "48px",
              borderRadius: "14px",
              background: module.gradient,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontSize: "1.6rem",
              boxShadow: '0 8px 16px ' + module.glowColor,
              flexShrink: 0
            }}
          >
            {module.emoji}
          </div>
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
              <h3 style={{ margin: 0, fontSize: "1.1rem", fontWeight: 700, color: "#0f172a" }}>{module.title}</h3>
              {module.isLive && (
                <span style={{ background: "#fee2e2", color: "#ef4444", padding: "2px 8px", borderRadius: "9999px", fontSize: "0.7rem", fontWeight: 700 }}>
                  ⚡ LIVE
                </span>
              )}
              {module.isAI && (
                <span style={{ background: "#e0f2fe", color: "#0284c7", padding: "2px 8px", borderRadius: "9999px", fontSize: "0.7rem", fontWeight: 700 }}>
                  🤖 AI
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
              padding: "9px 20px",
              borderRadius: "10px",
              background: "linear-gradient(135deg, #4f46e5 0%, #3730a3 100%)",
              color: "#ffffff",
              fontSize: "0.85rem",
              fontWeight: 700,
              display: "inline-flex",
              alignItems: "center",
              gap: "6px",
              boxShadow: "0 4px 12px rgba(79, 70, 229, 0.2)"
            }}
          >
            Open →
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
      whileHover={{ y: -6, transition: { duration: 0.25 } }}
      onMouseEnter={() => setHoveredCard?.(module.title)}
      onMouseLeave={() => setHoveredCard?.(null)}
      style={{
        position: "relative",
        background: "#ffffff",
        border: isHovered ? "1px solid #818cf8" : "1px solid #e2e8f0",
        borderRadius: "24px",
        padding: "24px",
        boxShadow: isHovered
          ? '0 20px 40px -10px ' + module.glowColor + ', 0 10px 25px rgba(0,0,0,0.05)'
          : "0 4px 20px rgba(0, 0, 0, 0.03)",
        display: "flex",
        flexDirection: "column",
        justifyContent: "space-between",
        minHeight: "285px",
        transition: "all 0.3s ease",
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
          height: "5px",
          background: module.gradient,
          opacity: isHovered ? 1 : 0.8
        }}
      />

      {/* Header Row: Category Badge & Pin Star */}
      <div>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: "16px" }}>
          <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
            <span
              style={{
                background: module.category === "Production"
                  ? "#eff6ff"
                  : module.category === "Reports"
                    ? "#ecfdf5"
                    : "#f3e8ff",
                color: module.category === "Production"
                  ? "#2563eb"
                  : module.category === "Reports"
                    ? "#059669"
                    : "#7c3aed",
                border: "1px solid " + (module.category === "Production" ? "#bfdbfe" : module.category === "Reports" ? "#a7f3d0" : "#ddd6fe"),
                padding: "4px 12px",
                borderRadius: "9999px",
                fontSize: "0.75rem",
                fontWeight: 700
              }}
            >
              {module.category}
            </span>

            {module.status === "pending" && (
              <span
                style={{
                  background: "#fef3c7",
                  color: "#d97706",
                  border: "1px solid #fde68a",
                  padding: "4px 10px",
                  borderRadius: "9999px",
                  fontSize: "0.75rem",
                  fontWeight: 700
                }}
              >
                ⏳ Attention
              </span>
            )}
            {module.isLive && (
              <span
                style={{
                  background: "#fee2e2",
                  color: "#ef4444",
                  border: "1px solid #fca5a5",
                  padding: "4px 10px",
                  borderRadius: "9999px",
                  fontSize: "0.75rem",
                  fontWeight: 700
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
                  fontSize: "0.75rem",
                  fontWeight: 700
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
              width: "58px",
              height: "58px",
              borderRadius: "18px",
              background: module.gradient,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontSize: "2.1rem",
              boxShadow: '0 10px 20px ' + module.glowColor,
              transform: isHovered ? "scale(1.08) rotate(3deg)" : "scale(1)",
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
                fontSize: "1.2rem",
                fontWeight: 800,
                color: "#0f172a",
                lineHeight: 1.3
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
            fontSize: "0.9rem",
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
                fontSize: "0.75rem",
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
            ? "linear-gradient(135deg, #4f46e5 0%, #3730a3 100%)"
            : "#f8fafc",
          color: isHovered ? "#ffffff" : "#4f46e5",
          fontSize: "0.9rem",
          fontWeight: 700,
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          border: isHovered ? "1px solid #4338ca" : "1px solid #e2e8f0",
          transition: "all 0.3s ease",
          boxShadow: isHovered ? "0 8px 20px rgba(79, 70, 229, 0.25)" : "none"
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

