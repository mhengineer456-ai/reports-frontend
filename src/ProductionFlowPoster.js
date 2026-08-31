import React, { useState, useRef } from "react";
import { useHistory } from "react-router-dom";
import { motion } from "framer-motion";

export default function ProductionFlowPoster() {
  const history = useHistory();
  const posterRef = useRef(null);

  const [activeDept, setActiveDept] = useState("all");
  const [zoomLevel, setZoomLevel] = useState(1);
  const [showMetrics, setShowMetrics] = useState(true);

  const handlePrint = () => {
    window.print();
  };

  const handleZoom = (delta) => {
    setZoomLevel((prev) => Math.min(Math.max(0.6, prev + delta), 1.6));
  };

  const resetZoom = () => {
    setZoomLevel(1);
  };

  const departments = [
    { id: "materials", name: "1. Materials & Yarn", icon: "🧵", color: "#f59e0b", badge: "Stage 01" },
    { id: "cutting", name: "2. Cutting Department", icon: "✂️", color: "#2563eb", badge: "Stage 02" },
    { id: "emb_print", name: "3. Embroidery & Printing", icon: "🎨", color: "#a855f7", badge: "Stage 03" },
    { id: "stitching", name: "4. Stitching Operations", icon: "🪡", color: "#06b6d4", badge: "Stage 04" },
    { id: "packing", name: "5. Packing & Finishing", icon: "📦", color: "#10b981", badge: "Stage 05" },
    { id: "executive", name: "6. Executive Matrix", icon: "📊", color: "#8b5cf6", badge: "Stage 06" }
  ];

  const workflowStages = [
    {
      stageId: "materials",
      stageNumber: "STAGE 01",
      title: "Raw Materials, Yarn & Trims Inventory",
      tagline: "Receipts, stock ledgers, kora rolls, knitting & accessories procurement",
      color: "#f59e0b",
      gradient: "linear-gradient(135deg, #78350f 0%, #b45309 50%, #f59e0b 100%)",
      glowColor: "rgba(245, 158, 11, 0.4)",
      reports: [
        {
          code: "R-01",
          title: "Yarn Stock Report",
          path: "/yarn-report",
          emoji: "🧵",
          desc: "Knitting yarn inventory, package counts & balance weight in KGs",
          tags: ["Yarn Stock", "Bal WT (KGs)", "Pkgs"],
          adminOnly: true
        },
        {
          code: "R-02",
          title: "Kora Roll Report",
          path: "/kora-roll-report",
          emoji: "📜",
          desc: "Un-dyed kora rolls, weight (KGs), meters, dyeing dispatch",
          tags: ["Kora Rolls", "Weight / Mtr", "Dyeing"],
          adminOnly: true
        },
        {
          code: "R-03",
          title: "Knitting Machine Report",
          path: "/knitting-report",
          emoji: "🧶",
          desc: "Knitting output, yarn conversion efficiency & roll outputs",
          tags: ["Machine Gauge", "GSM", "Shift Output"],
          adminOnly: true
        },
        {
          code: "R-04",
          title: "Collar & Cuff Report",
          path: "/collar-report",
          emoji: "👔",
          desc: "Flat knitting production, size-wise collar counts & machine output",
          tags: ["Collar Count", "Flat Knitting", "Cuffs"],
          adminOnly: true
        },
        {
          code: "R-05",
          title: "Pending Zip PO Report",
          path: "/pending-zip-po-report",
          emoji: "🤐⏳",
          desc: "Production lots requiring zips with pending Purchase Orders",
          tags: ["Zip Required", "PO Pending", "Trims"]
        },
        {
          code: "R-06",
          title: "Sticker & Tag Report",
          path: "/sticker-report",
          emoji: "🏷️",
          desc: "Sticker printing, label usage & inventory management",
          tags: ["Label Stock", "Size Breakdown", "Barcode"]
        },
        {
          code: "R-07",
          title: "Embroidery Material Receiving",
          path: "/embroidery-material-receiving",
          emoji: "🧵📥",
          desc: "Thread cones, backing paper, needles & raw receipts",
          tags: ["Thread Stock", "Backing", "Receipts"]
        },
        {
          code: "R-08",
          title: "Printing Material Receiving",
          path: "/printing-material-receiving",
          emoji: "🖨️📥",
          desc: "Printing inks, chemicals, screens & supplier deliveries",
          tags: ["Ink Inventory", "Screen Mesh", "Chemicals"]
        }
      ]
    },
    {
      stageId: "cutting",
      stageNumber: "STAGE 02",
      title: "Cutting Department & Fabric Allocation",
      tagline: "Fabric roll distribution, cutting table assignment, AI optimization & layer plans",
      color: "#2563eb",
      gradient: "linear-gradient(135deg, #1e3a8a 0%, #1d4ed8 50%, #2563eb 100%)",
      glowColor: "rgba(37, 99, 235, 0.4)",
      reports: [
        {
          code: "R-09",
          title: "Daily Fabric Issue Analytics",
          path: "/daily-fabric-issue-report",
          emoji: "🧵",
          desc: "Visual distribution matrix mapping weights across tables & fabric styles",
          tags: ["Table Wise", "Fabric Summary", "Rolls"],
          badge: "LIVE MATRIX"
        },
        {
          code: "R-10",
          title: "All Cutting Job Orders",
          path: "/all-cutting-joborders",
          emoji: "📝",
          desc: "Centralized view of all active job orders, assignments & queue status",
          tags: ["Job Order #", "Cutting Queue", "Order Status"]
        },
        {
          code: "R-11",
          title: "Fabric Roll Prediction (AI)",
          path: "/fabric-roll-prediction",
          emoji: "📏",
          desc: "AI fabric consumption forecast, roll optimization & smart cutting plan",
          tags: ["AI Optimization", "Roll Forecast", "Consumption"],
          badge: "AI POWERED"
        },
        {
          code: "R-12",
          title: "Cutting Operations Report",
          path: "/cutting-report",
          emoji: "✂️",
          desc: "Cutting statistics, fabric consumption details, layers & marker output",
          tags: ["Layers Summary", "Fabric Usage", "Efficiency"]
        }
      ]
    },
    {
      stageId: "emb_print",
      stageNumber: "STAGE 03",
      title: "Embroidery & Printing Processing",
      tagline: "Challan issuance, screen management, remarks tracking & stitching dispatch",
      color: "#a855f7",
      gradient: "linear-gradient(135deg, #581c87 0%, #7e22ce 50%, #a855f7 100%)",
      glowColor: "rgba(168, 85, 247, 0.4)",
      reports: [
        {
          code: "R-13",
          title: "Embroidery Challan Tracker",
          path: "/embroidery",
          emoji: "🧵",
          desc: "Create & manage embroidery production challans with batch tracking",
          tags: ["Challan History", "Stitch Count", "Batch Wise"]
        },
        {
          code: "R-14",
          title: "Printing Challan Tracker",
          path: "/printing",
          emoji: "🖨️",
          desc: "Manage printing orders, screen tracking, color output & workflows",
          tags: ["Screen Tracking", "Color Wise", "Workflows"]
        },
        {
          code: "R-15",
          title: "EMB / PRINT REMARKS Hub",
          path: "/emb-print-remarks",
          emoji: "📝",
          desc: "Live remarks tracker for pending Embroidery & Printing batches",
          tags: ["Live Remarks", "Pending Lots", "Issue Updates"],
          badge: "REAL-TIME"
        },
        {
          code: "R-16",
          title: "Daily Embroidery Challan",
          path: "/daily-embroidery-challan",
          emoji: "📈",
          desc: "Daily embroidery progress, machine efficiency & thread consumption",
          tags: ["Daily Output", "Machine Wise", "Thread Usage"]
        },
        {
          code: "R-17",
          title: "Daily Printing Challan",
          path: "/daily-printing-challan",
          emoji: "📊",
          desc: "Daily screen utilization, color output & operator analytics",
          tags: ["Daily Printing", "Screen Usage", "Operators"]
        },
        {
          code: "R-18",
          title: "After EMB/PRINT DONE",
          path: "/pending-issue-to-stitching",
          emoji: "📋",
          desc: "Track completed embroidery/print garments pending for stitching dispatch",
          tags: ["Ready for Stitch", "Quality Check", "Transfer Queue"]
        },
        {
          code: "R-19",
          title: "Jaybir Printing Report",
          path: "/jaybir-printing-report",
          emoji: "🖨️",
          desc: "Jaybir printing unit production report, screen tracking & lot dispatch",
          tags: ["Jaybir Unit", "Screen Tracking", "Lot Output"]
        },
        {
          code: "R-20",
          title: "Jaybir Embroidery Report",
          path: "/jaybir-embroidery-report",
          emoji: "🧵",
          desc: "Jaybir embroidery stitch output, machine utilization & challans",
          tags: ["Jaybir Unit", "Stitch Count", "Challan Tracking"]
        }
      ]
    },
    {
      stageId: "stitching",
      stageNumber: "STAGE 04",
      title: "Stitching & Floor Assembly Operations",
      tagline: "Line allocation, hourly targets, overlock, button/kaj work & lot completion",
      color: "#06b6d4",
      gradient: "linear-gradient(135deg, #164e63 0%, #0e7490 50%, #06b6d4 100%)",
      glowColor: "rgba(6, 182, 212, 0.4)",
      reports: [
        {
          code: "R-19",
          title: "After Cutting Done (Pending Stitching)",
          path: "/pending-stitching",
          emoji: "⏳",
          desc: "Monitor cut lots awaiting line issue & delayed production batches",
          tags: ["Pending Queue", "Priority Aging", "Delay Analysis"]
        },
        {
          code: "R-20",
          title: "Daily Stitching Issue Operations",
          path: "/stitching",
          emoji: "🧶",
          desc: "Daily line assignments, operator allocation & hourly target tracking",
          tags: ["Line Allocation", "Operator Wise", "Hourly Output"]
        },
        {
          code: "R-21",
          title: "Daily Kaj Button Report",
          path: "/daily-kaj-button-report",
          emoji: "🔘",
          desc: "Daily button attachment, kaj work output, supervisor & aging tracking",
          tags: ["Button Count", "Kaj Quantity", "Supervisor Wise"]
        },
        {
          code: "R-22",
          title: "Daily Overlock Report",
          path: "/daily-overlock-report",
          emoji: "🧵",
          desc: "Overlock stitching production, machine output & quality inspection",
          tags: ["Overlock Output", "Thread Used", "Quality Check"]
        },
        {
          code: "R-23",
          title: "Not Updation Report",
          path: "/daily-stitching-report-not-updation",
          emoji: "⚠️",
          desc: "Identify line reports requiring updates or missing production entries",
          tags: ["Missing Entries", "Due Reports", "Line Pending"]
        },
        {
          code: "R-24",
          title: "Overall Stitching Report (Completed Lots)",
          path: "/stitching-complete-lot",
          emoji: "✅",
          desc: "Verified completed stitching lots with batch-wise performance metrics",
          tags: ["Completed Lots", "Efficiency", "Stitching Done"]
        }
      ]
    },
    {
      stageId: "packing",
      stageNumber: "STAGE 05",
      title: "Finishing, Packaging & Carton Dispatch",
      tagline: "Quality inspection, sticker issuance, folding, carton preparation & dispatch",
      color: "#10b981",
      gradient: "linear-gradient(135deg, #064e3b 0%, #047857 50%, #10b981 100%)",
      glowColor: "rgba(16, 185, 129, 0.4)",
      reports: [
        {
          code: "R-25",
          title: "Pending Packing to Issue",
          path: "/pending-packing-issue",
          emoji: "📦⏳",
          desc: "Stitching-completed lots ready and awaiting packing line assignment",
          tags: ["Ready to Pack", "Stitch Complete", "Pending Issue"],
          badge: "CRITICAL QUEUE"
        },
        {
          code: "R-26",
          title: "Issue To Packing Operations",
          path: "/issue-to-packing",
          emoji: "📦",
          desc: "Manage packing allocation, carton planning & shipment readiness",
          tags: ["Carton Tracking", "Shipment Prep", "Quality Pass"]
        },
        {
          code: "R-27",
          title: "Packing Alloted Lot",
          path: "/packing-alloted-lot",
          emoji: "📦✅",
          desc: "Lots allocated for packaging with operator assignment and cartons",
          tags: ["Lot Allocation", "Carton Plan", "Operator Assign"]
        },
        {
          code: "R-28",
          title: "Packing Complete Lots",
          path: "/packing-completed-lots",
          emoji: "📦🏆",
          desc: "All completed packing lots with verified complete dates & aging",
          tags: ["Complete Date", "Aging Days", "Barcode Linkage"]
        },
        {
          code: "R-29",
          title: "Daily Packing Report",
          path: "/daily-packing-report",
          emoji: "📦",
          desc: "Daily packed carton output, shift summary & shipment efficiency",
          tags: ["Carton Count", "Daily Packed", "Efficiency"]
        },
        {
          code: "R-32",
          title: "Daily Folding Report",
          path: "/daily-folding-report",
          emoji: "👕",
          desc: "Daily garment folding, style-wise breakdown & supervisor analytics",
          tags: ["Style Breakdown", "Folding Output", "Quality Pass"]
        },
        {
          code: "R-33",
          title: "Washing & Treatment Report",
          path: "/washing-report",
          emoji: "🧼",
          desc: "Garment washing cycles, shrinkage testing & dispatch handover",
          tags: ["Wash Cycle", "Shrinkage %", "Dispatch to Pack"]
        }
      ]
    },
    {
      stageId: "executive",
      stageNumber: "STAGE 06",
      title: "Executive Master Control & Milestone Analytics",
      tagline: "Cross-departmental WIP matrix, bottleneck detection, live KPI dashboard & audit logs",
      color: "#8b5cf6",
      gradient: "linear-gradient(135deg, #2e1065 0%, #5b21b6 50%, #8b5cf6 100%)",
      glowColor: "rgba(139, 92, 246, 0.4)",
      reports: [
        {
          code: "R-31",
          title: "Cut to Pack Report (WIP Matrix)",
          path: "/overall-cutting-to-packing-report",
          emoji: "📊",
          desc: "End-to-end production matrix from cutting to packing with stage aging",
          tags: ["WIP Matrix", "Bottlenecks", "Cut To Pack"],
          badge: "MASTER MATRIX"
        },
        {
          code: "R-32",
          title: "Short Summary Report (Live KPIs)",
          path: "/short-summary-report",
          emoji: "⚡",
          desc: "Executive command center overview of KPIs across all departments",
          tags: ["Live KPIs", "Executive Summary", "Alerts"],
          adminOnly: true,
          badge: "EXECUTIVE LIVE"
        },
        {
          code: "R-33",
          title: "Lot Change Logs (Audit Trail)",
          path: "/lot-logs",
          emoji: "📋",
          desc: "Log lot modifications, authority authorizations & real-time broadcast",
          tags: ["Audit Trail", "Modifications", "Permissions"]
        },
        {
          code: "R-34",
          title: "Lot Lifecycle Timeline (Amazon Tracker)",
          path: "/lot-timeline",
          emoji: "📦⚡",
          desc: "Amazon Order Tracking-style visual timeline: Cutting, EMB/Print, Stitching, Kaj Button & Packing",
          tags: ["Amazon Stepper", "Lot Milestones", "Stage Aging"],
          badge: "AMAZON TRACKER"
        },
        {
          code: "R-35",
          title: "Notification Center",
          path: "/notifications",
          emoji: "🔔",
          desc: "Full-screen activity feed for factory alerts, hold lots & milestone events",
          tags: ["Activity Feed", "Live Alerts", "Lot Updates"]
        }
      ]
    }
  ];

  const totalReportsCount = workflowStages.reduce((acc, stage) => acc + stage.reports.length, 0);

  return (
    <>
      <style>{`
        @media print {
          body {
            background: #ffffff !important;
            color: #000000 !important;
          }
          .no-print {
            display: none !important;
          }
          .poster-main-container {
            padding: 0 !important;
            background: #ffffff !important;
            max-width: 100% !important;
          }
          .flowchart-stage-box {
            break-inside: avoid !important;
            box-shadow: none !important;
            border: 1px solid #cbd5e1 !important;
            page-break-inside: avoid !important;
            margin-bottom: 20px !important;
          }
          .flow-report-card {
            box-shadow: none !important;
            border: 1px solid #94a3b8 !important;
            background: #f8fafc !important;
          }
        }
      `}</style>

      <div
        style={{
          minHeight: "100vh",
          backgroundColor: "#090d16",
          backgroundImage: `
            radial-gradient(at 0% 0%, rgba(99, 102, 241, 0.15) 0px, transparent 50%),
            radial-gradient(at 100% 0%, rgba(236, 72, 153, 0.12) 0px, transparent 50%),
            radial-gradient(at 50% 100%, rgba(16, 185, 129, 0.12) 0px, transparent 50%),
            linear-gradient(180deg, #090d16 0%, #0f172a 100%)
          `,
          color: "#f8fafc",
          fontFamily: "'Plus Jakarta Sans', 'Inter', system-ui, -apple-system, sans-serif",
          paddingBottom: "80px"
        }}
      >
        {/* Floating Top Poster Control Bar (Fixed) */}
        <div
          className="no-print"
          style={{
            position: "sticky",
            top: 0,
            zIndex: 100,
            background: "rgba(15, 23, 42, 0.85)",
            backdropFilter: "blur(16px)",
            borderBottom: "1px solid rgba(255, 255, 255, 0.12)",
            padding: "14px 32px",
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            flexWrap: "wrap",
            gap: "16px",
            boxShadow: "0 10px 30px rgba(0, 0, 0, 0.5)"
          }}
        >
          {/* Left branding */}
          <div style={{ display: "flex", alignItems: "center", gap: "14px" }}>
            <button
              onClick={() => history.push("/dashboard")}
              style={{
                background: "rgba(255, 255, 255, 0.1)",
                border: "1px solid rgba(255, 255, 255, 0.2)",
                color: "#ffffff",
                padding: "8px 14px",
                borderRadius: "10px",
                cursor: "pointer",
                fontWeight: 700,
                fontSize: "0.85rem",
                display: "flex",
                alignItems: "center",
                gap: "6px"
              }}
            >
              ← Dashboard
            </button>

            <div>
              <h2 style={{ margin: 0, fontSize: "1.15rem", fontWeight: 800, color: "#ffffff", display: "flex", alignItems: "center", gap: "8px" }}>
                <span>🗺️</span> Factory Production Flowchart & Reports Poster
              </h2>
              <span style={{ fontSize: "0.78rem", color: "#94a3b8", fontWeight: 500 }}>
                High-definition operational pipeline mapping {totalReportsCount} reports across 6 stages
              </span>
            </div>
          </div>

          {/* Department Quick Filter Buttons */}
          <div style={{ display: "flex", alignItems: "center", gap: "6px", flexWrap: "wrap" }}>
            <button
              onClick={() => setActiveDept("all")}
              style={{
                padding: "6px 14px",
                borderRadius: "8px",
                border: activeDept === "all" ? "1px solid #6366f1" : "1px solid rgba(255,255,255,0.15)",
                background: activeDept === "all" ? "linear-gradient(135deg, #4f46e5 0%, #3730a3 100%)" : "rgba(255,255,255,0.06)",
                color: "#ffffff",
                fontWeight: 700,
                fontSize: "0.8rem",
                cursor: "pointer"
              }}
            >
              All 6 Stages
            </button>
            {departments.map((d) => (
              <button
                key={d.id}
                onClick={() => setActiveDept(d.id)}
                style={{
                  padding: "6px 12px",
                  borderRadius: "8px",
                  border: activeDept === d.id ? `1px solid ${d.color}` : "1px solid rgba(255,255,255,0.12)",
                  background: activeDept === d.id ? d.color : "rgba(255,255,255,0.06)",
                  color: "#ffffff",
                  fontWeight: 700,
                  fontSize: "0.78rem",
                  cursor: "pointer",
                  display: "flex",
                  alignItems: "center",
                  gap: "5px"
                }}
              >
                <span>{d.icon}</span>
                <span>{d.name.split(". ")[1]}</span>
              </button>
            ))}
          </div>

          {/* Right Action Tools (Zoom, Metrics, Print) */}
          <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
            <div style={{ display: "flex", alignItems: "center", background: "rgba(255, 255, 255, 0.08)", padding: "4px", borderRadius: "10px", border: "1px solid rgba(255,255,255,0.15)" }}>
              <button
                onClick={() => handleZoom(-0.1)}
                title="Zoom Out"
                style={{ background: "transparent", border: "none", color: "#e2e8f0", padding: "6px 10px", cursor: "pointer", fontWeight: 800, fontSize: "0.9rem" }}
              >
                −
              </button>
              <button
                onClick={resetZoom}
                title="Reset Zoom"
                style={{ background: "transparent", border: "none", color: "#c7d2fe", padding: "6px 10px", cursor: "pointer", fontWeight: 700, fontSize: "0.8rem" }}
              >
                {Math.round(zoomLevel * 100)}%
              </button>
              <button
                onClick={() => handleZoom(0.1)}
                title="Zoom In"
                style={{ background: "transparent", border: "none", color: "#e2e8f0", padding: "6px 10px", cursor: "pointer", fontWeight: 800, fontSize: "0.9rem" }}
              >
                +
              </button>
            </div>

            <button
              onClick={() => setShowMetrics(!showMetrics)}
              style={{
                background: showMetrics ? "rgba(16, 185, 129, 0.2)" : "rgba(255, 255, 255, 0.08)",
                border: showMetrics ? "1px solid #10b981" : "1px solid rgba(255, 255, 255, 0.2)",
                color: showMetrics ? "#34d399" : "#cbd5e1",
                padding: "8px 14px",
                borderRadius: "10px",
                cursor: "pointer",
                fontWeight: 700,
                fontSize: "0.82rem",
                display: "flex",
                alignItems: "center",
                gap: "6px"
              }}
            >
              🏷️ {showMetrics ? "Tags On" : "Tags Off"}
            </button>

            <button
              onClick={handlePrint}
              style={{
                background: "linear-gradient(135deg, #6366f1 0%, #4338ca 100%)",
                border: "none",
                color: "#ffffff",
                padding: "8px 18px",
                borderRadius: "10px",
                cursor: "pointer",
                fontWeight: 800,
                fontSize: "0.85rem",
                display: "flex",
                alignItems: "center",
                gap: "6px",
                boxShadow: "0 4px 15px rgba(99, 102, 241, 0.4)"
              }}
            >
              🖨️ Print Poster / PDF
            </button>
          </div>
        </div>

        {/* Poster Wrapper Container with Zoom Transform */}
        <div
          ref={posterRef}
          className="poster-main-container"
          style={{
            maxWidth: "1850px",
            margin: "0 auto",
            padding: "40px 32px 60px",
            transform: `scale(${zoomLevel})`,
            transformOrigin: "top center",
            transition: "transform 0.2s ease"
          }}
        >
          {/* Poster Top Hero Banner */}
          <motion.div
            initial={{ opacity: 0, y: -20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.6 }}
            style={{
              background: "linear-gradient(135deg, #1e1b4b 0%, #312e81 40%, #0f172a 100%)",
              borderRadius: "28px",
              padding: "36px 44px",
              border: "1px solid rgba(255, 255, 255, 0.15)",
              boxShadow: "0 25px 50px -12px rgba(0, 0, 0, 0.6)",
              marginBottom: "40px",
              position: "relative",
              overflow: "hidden"
            }}
          >
            {/* Ambient Lighting Gradients */}
            <div style={{ position: "absolute", top: "-80px", right: "-80px", width: "350px", height: "350px", borderRadius: "50%", background: "rgba(99, 102, 241, 0.25)", filter: "blur(60px)", pointerEvents: "none" }} />
            <div style={{ position: "absolute", bottom: "-80px", left: "15%", width: "400px", height: "400px", borderRadius: "50%", background: "rgba(236, 72, 153, 0.15)", filter: "blur(70px)", pointerEvents: "none" }} />

            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: "20px", position: "relative", zIndex: 2 }}>
              <div>
                <div style={{ display: "inline-flex", alignItems: "center", gap: "10px", padding: "6px 16px", borderRadius: "9999px", background: "rgba(255, 255, 255, 0.1)", backdropFilter: "blur(10px)", border: "1px solid rgba(255, 255, 255, 0.2)", fontSize: "0.85rem", fontWeight: 700, color: "#c7d2fe", marginBottom: "14px" }}>
                  <span style={{ display: "inline-block", width: "8px", height: "8px", borderRadius: "50%", background: "#34d399", boxShadow: "0 0 10px #34d399" }} />
                  <span>MH Enterprise Garment Suite</span>
                  <span style={{ opacity: 0.4 }}>•</span>
                  <span>End-to-End Manufacturing Architecture</span>
                </div>

                <h1 style={{ margin: 0, fontSize: "clamp(2.2rem, 3.8vw, 3.2rem)", fontWeight: 900, letterSpacing: "-0.03em", color: "#ffffff", lineHeight: 1.15 }}>
                  Factory Production Flowchart & Master Reports Poster
                </h1>
                <p style={{ margin: "12px 0 0 0", fontSize: "1.05rem", color: "#cbd5e1", maxWidth: "880px", lineHeight: 1.5, fontWeight: 400 }}>
                  Interactive visual blueprint detailing each department's operational role, from initial yarn/fabric procurement, through cutting, embroidery, printing, stitching lines, and final packaging to executive milestone analytics.
                </p>
              </div>

              {/* Poster Key KPI Badges */}
              <div style={{ display: "flex", gap: "14px", flexWrap: "wrap" }}>
                <div style={{ background: "rgba(255, 255, 255, 0.08)", backdropFilter: "blur(12px)", padding: "18px 24px", borderRadius: "18px", border: "1px solid rgba(255, 255, 255, 0.15)", textAlign: "center" }}>
                  <div style={{ fontSize: "2rem", fontWeight: 900, color: "#818cf8", lineHeight: 1 }}>6</div>
                  <div style={{ fontSize: "0.75rem", fontWeight: 700, color: "#cbd5e1", marginTop: "4px", textTransform: "uppercase", letterSpacing: "0.5px" }}>Core Stages</div>
                </div>
                <div style={{ background: "rgba(255, 255, 255, 0.08)", backdropFilter: "blur(12px)", padding: "18px 24px", borderRadius: "18px", border: "1px solid rgba(255, 255, 255, 0.15)", textAlign: "center" }}>
                  <div style={{ fontSize: "2rem", fontWeight: 900, color: "#34d399", lineHeight: 1 }}>{totalReportsCount}</div>
                  <div style={{ fontSize: "0.75rem", fontWeight: 700, color: "#cbd5e1", marginTop: "4px", textTransform: "uppercase", letterSpacing: "0.5px" }}>Total Reports</div>
                </div>
                <div style={{ background: "rgba(255, 255, 255, 0.08)", backdropFilter: "blur(12px)", padding: "18px 24px", borderRadius: "18px", border: "1px solid rgba(255, 255, 255, 0.15)", textAlign: "center" }}>
                  <div style={{ fontSize: "2rem", fontWeight: 900, color: "#fbbf24", lineHeight: 1 }}>100%</div>
                  <div style={{ fontSize: "0.75rem", fontWeight: 700, color: "#cbd5e1", marginTop: "4px", textTransform: "uppercase", letterSpacing: "0.5px" }}>Live Sync</div>
                </div>
              </div>
            </div>
          </motion.div>

          {/* Master Sequential Flowchart Lane Cards */}
          <div style={{ display: "flex", flexDirection: "column", gap: "40px" }}>
            {workflowStages.map((stage, sIdx) => {
              if (activeDept !== "all" && activeDept !== stage.stageId) {
                return null;
              }

              return (
                <div key={stage.stageId} style={{ position: "relative" }}>
                  {/* Flow Connector Arrow to Next Stage */}
                  {sIdx < workflowStages.length - 1 && activeDept === "all" && (
                    <div
                      className="no-print"
                      style={{
                        position: "absolute",
                        bottom: "-32px",
                        left: "50%",
                        transform: "translateX(-50%)",
                        zIndex: 10,
                        display: "flex",
                        flexDirection: "column",
                        alignItems: "center",
                        gap: "2px"
                      }}
                    >
                      <span style={{ fontSize: "1.2rem", color: stage.color, filter: `drop-shadow(0 0 8px ${stage.color})` }}>
                        ↓
                      </span>
                    </div>
                  )}

                  {/* Stage Container Box */}
                  <motion.div
                    initial={{ opacity: 0, y: 30 }}
                    whileInView={{ opacity: 1, y: 0 }}
                    viewport={{ once: true }}
                    transition={{ duration: 0.5, delay: sIdx * 0.1 }}
                    className="flowchart-stage-box"
                    style={{
                      background: "#111827",
                      borderRadius: "26px",
                      border: `1.5px solid rgba(255, 255, 255, 0.12)`,
                      boxShadow: `0 20px 40px -10px rgba(0, 0, 0, 0.5), 0 0 30px -10px ${stage.glowColor}`,
                      overflow: "hidden"
                    }}
                  >
                    {/* Stage Header Banner */}
                    <div
                      style={{
                        background: stage.gradient,
                        padding: "24px 36px",
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "space-between",
                        flexWrap: "wrap",
                        gap: "16px",
                        color: "#ffffff"
                      }}
                    >
                      <div style={{ display: "flex", alignItems: "center", gap: "18px" }}>
                        <div
                          style={{
                            background: "rgba(255, 255, 255, 0.2)",
                            backdropFilter: "blur(12px)",
                            border: "1px solid rgba(255, 255, 255, 0.3)",
                            borderRadius: "16px",
                            padding: "8px 18px",
                            fontSize: "0.88rem",
                            fontWeight: 900,
                            letterSpacing: "1px",
                            textTransform: "uppercase"
                          }}
                        >
                          {stage.stageNumber}
                        </div>

                        <div>
                          <h3 style={{ margin: 0, fontSize: "1.45rem", fontWeight: 900, letterSpacing: "-0.01em", color: "#ffffff" }}>
                            {stage.title}
                          </h3>
                          <p style={{ margin: "4px 0 0 0", fontSize: "0.9rem", color: "rgba(255, 255, 255, 0.9)", fontWeight: 500 }}>
                            {stage.tagline}
                          </p>
                        </div>
                      </div>

                      <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
                        <span
                          style={{
                            background: "rgba(0, 0, 0, 0.3)",
                            padding: "6px 16px",
                            borderRadius: "9999px",
                            fontSize: "0.82rem",
                            fontWeight: 800,
                            border: "1px solid rgba(255, 255, 255, 0.2)"
                          }}
                        >
                          {stage.reports.length} Interactive Reports
                        </span>
                      </div>
                    </div>

                    {/* Stage Reports Grid */}
                    <div
                      style={{
                        padding: "28px 32px",
                        display: "grid",
                        gridTemplateColumns: "repeat(auto-fill, minmax(340px, 1fr))",
                        gap: "20px",
                        background: "rgba(15, 23, 42, 0.6)"
                      }}
                    >
                      {stage.reports.map((report) => (
                        <div
                          key={report.code + report.title}
                          onClick={() => history.push(report.path)}
                          className="flow-report-card"
                          style={{
                            background: "#1e293b",
                            borderRadius: "18px",
                            padding: "22px 24px",
                            border: "1.5px solid rgba(255, 255, 255, 0.08)",
                            boxShadow: "0 4px 14px rgba(0, 0, 0, 0.2)",
                            cursor: "pointer",
                            transition: "all 0.22s cubic-bezier(0.4, 0, 0.2, 1)",
                            display: "flex",
                            flexDirection: "column",
                            justifyContent: "space-between",
                            position: "relative",
                            overflow: "hidden"
                          }}
                          onMouseEnter={(e) => {
                            e.currentTarget.style.transform = "translateY(-5px)";
                            e.currentTarget.style.borderColor = stage.color;
                            e.currentTarget.style.boxShadow = `0 14px 28px -5px ${stage.glowColor}, 0 4px 10px rgba(0,0,0,0.3)`;
                          }}
                          onMouseLeave={(e) => {
                            e.currentTarget.style.transform = "translateY(0)";
                            e.currentTarget.style.borderColor = "rgba(255, 255, 255, 0.08)";
                            e.currentTarget.style.boxShadow = "0 4px 14px rgba(0, 0, 0, 0.2)";
                          }}
                        >
                          {/* Card Top Accent Line */}
                          <div style={{ position: "absolute", top: 0, left: 0, right: 0, height: "4px", background: stage.gradient }} />

                          <div>
                            {/* Card Header Row */}
                            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: "12px" }}>
                              <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                                <span style={{ fontSize: "1.8rem" }}>{report.emoji}</span>
                                <span style={{ background: "rgba(255, 255, 255, 0.08)", color: "#cbd5e1", padding: "2px 8px", borderRadius: "6px", fontSize: "0.72rem", fontWeight: 800, letterSpacing: "0.5px" }}>
                                  {report.code}
                                </span>
                              </div>

                              <div style={{ display: "flex", gap: "6px" }}>
                                {report.badge && (
                                  <span style={{ background: "rgba(99, 102, 241, 0.25)", color: "#a5b4fc", border: "1px solid rgba(99, 102, 241, 0.4)", fontSize: "0.68rem", fontWeight: 800, padding: "3px 8px", borderRadius: "6px" }}>
                                    {report.badge}
                                  </span>
                                )}
                                {report.adminOnly && (
                                  <span style={{ background: "rgba(239, 68, 68, 0.2)", color: "#fca5a5", border: "1px solid rgba(239, 68, 68, 0.4)", fontSize: "0.68rem", fontWeight: 800, padding: "3px 8px", borderRadius: "6px" }}>
                                    ADMIN
                                  </span>
                                )}
                              </div>
                            </div>

                            {/* Title & Description */}
                            <h4 style={{ margin: "0 0 6px 0", fontSize: "1.1rem", fontWeight: 800, color: "#ffffff", lineHeight: 1.3 }}>
                              {report.title}
                            </h4>
                            <p style={{ margin: 0, fontSize: "0.85rem", color: "#94a3b8", lineHeight: 1.45, fontWeight: 400 }}>
                              {report.desc}
                            </p>

                            {/* Metrics / Key Tags */}
                            {showMetrics && (
                              <div style={{ display: "flex", gap: "6px", flexWrap: "wrap", marginTop: "14px" }}>
                                {report.tags.map((tag) => (
                                  <span
                                    key={tag}
                                    style={{
                                      background: "rgba(255, 255, 255, 0.05)",
                                      color: "#cbd5e1",
                                      border: "1px solid rgba(255, 255, 255, 0.1)",
                                      padding: "3px 8px",
                                      borderRadius: "6px",
                                      fontSize: "0.72rem",
                                      fontWeight: 600
                                    }}
                                  >
                                    • {tag}
                                  </span>
                                ))}
                              </div>
                            )}
                          </div>

                          {/* Action Footer */}
                          <div
                            style={{
                              marginTop: "18px",
                              paddingTop: "12px",
                              borderTop: "1px solid rgba(255, 255, 255, 0.08)",
                              display: "flex",
                              alignItems: "center",
                              justifyContent: "space-between",
                              fontSize: "0.82rem",
                              fontWeight: 700,
                              color: stage.color
                            }}
                          >
                            <span>Open Report</span>
                            <span style={{ fontSize: "1rem" }}>↗</span>
                          </div>
                        </div>
                      ))}
                    </div>
                  </motion.div>
                </div>
              );
            })}
          </div>

          {/* Bottom Poster Footer Info */}
          <div
            style={{
              marginTop: "50px",
              padding: "24px 32px",
              background: "rgba(255, 255, 255, 0.03)",
              border: "1px solid rgba(255, 255, 255, 0.1)",
              borderRadius: "20px",
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              flexWrap: "wrap",
              gap: "16px",
              color: "#94a3b8",
              fontSize: "0.85rem"
            }}
          >
            <div>
              <strong style={{ color: "#ffffff" }}>MH Enterprise Garment Suite</strong> • Official Production Workflow Architecture
            </div>
            <div>
              📅 Generated: {new Date().toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })} • Verified Live System
            </div>
          </div>
        </div>
      </div>
    </>
  );
}
